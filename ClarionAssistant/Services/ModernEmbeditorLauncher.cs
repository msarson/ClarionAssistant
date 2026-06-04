using System;
using System.Collections.Generic;
using System.Threading;
using System.Windows.Forms;
using ICSharpCode.SharpDevelop.Gui;
using ClarionAssistant.Terminal;

namespace ClarionAssistant.Services
{
    /// <summary>
    /// Path B multi-editor: open a procedure's embed source in a Monaco view via the mirror+snapshot
    /// model. Clarion's native generator allows only ONE embeditor at a time, so for each procedure we:
    ///   1. OpenProcedureEmbed(name)  — native generation + Clarion's embeditor (transient)
    ///   2. mirror the live buffer (source + editable-region map)
    ///   3. CancelEmbeditor()         — discard/close to release the native single-embeditor lock
    ///   4. ShowView(new ModernEmbeditorViewContent)
    /// The snapshot lives in our own tab, so any number of procedures can be open at once.
    ///
    /// MUST run on the UI thread (OpenProcedureEmbed drives native focus + Application.DoEvents).
    /// Snapshots are read-only-of-truth for now; the save round-trip (re-open → write → save → close)
    /// is M2. If the .app is regenerated underneath, an open snapshot can go stale (reload to refresh).
    /// </summary>
    public static class ModernEmbeditorLauncher
    {
        /// <summary>Opens one procedure as a Monaco snapshot tab. Returns null on success, else an error message.</summary>
        public static string OpenProcedure(string procName, bool isDark)
        {
            if (string.IsNullOrWhiteSpace(procName)) return "No procedure specified.";
            EnterBusy();
            try
            {
                TimingMark("=== OpenProcedure '" + procName + "' START ===");
                var swTotal = System.Diagnostics.Stopwatch.StartNew();
                var appTree = new AppTreeService();

                string source, error;
                List<int[]> ranges;
                if (!OpenAndMirror(appTree, procName, out source, out ranges, out error))
                    return error;

                // ABC is loaded now (OpenAndMirror just opened the native embed, which triggers the lazy ABC
                // load). Warm the LSP HERE — deliberately not at picker-start — so its background solution parse
                // runs during the WebView2/Monaco load below rather than competing with the ABC load and
                // ~halving it. Idempotent, fire-and-forget.
                try { EmbeditorCompletionService.LspStarter?.Invoke(); } catch { }

                TimingMark("LSP kicked; about to CancelEmbeditor");
                // OpenAndMirror leaves the embeditor open; we made no edits, so discard/close to free the lock.
                var swClose = System.Diagnostics.Stopwatch.StartNew();
                try { appTree.CancelEmbeditor(); } catch { }
                TimingMark("CancelEmbeditor returned; waiting for closed");
                WaitForEmbedClosed(appTree, 3000);
                TimingLog("CancelEmbeditor+WaitClosed", swClose.ElapsedMilliseconds);
                TimingLog("OpenProcedure total (pre-deferred-ShowView)", swTotal.ElapsedMilliseconds);

                // CRITICAL — do NOT create the WebView2 view on THIS call stack. We are still unwinding the
                // nested Application.DoEvents() pumps that drove the native embeditor (SetFocus / AttachThreadInput
                // / WM_CHAR / BM_CLICK in OpenProcedureEmbed). WebView2's EnsureCoreWebView2Async — kicked off by
                // ShowView -> Panel.HandleCreated -> async OnHandleCreated — needs a SETTLED, non-reentrant
                // message-loop turn to complete; created on this reentrant/unsettled stack its await continuation
                // can't progress and the whole IDE hard-hangs. (This is the freeze: a manual idle GAP before
                // opening avoided it; ABC warmth was a red herring.) Post ShowView so this entire stack unwinds
                // and the message/input state drains first — deterministically reproducing that gap.
                var ctx = WindowsFormsSynchronizationContext.Current ?? new WindowsFormsSynchronizationContext();
                string capProc = procName; string capSrc = source; List<int[]> capRanges = ranges; bool capDark = isDark;
                ctx.Post(_ =>
                {
                    try
                    {
                        var view = new ModernEmbeditorViewContent(capProc, capSrc, capRanges, "clarion", capDark, capProc);
                        WorkbenchSingleton.Workbench.ShowView(view);
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine("[ModernEmbeditorLauncher] deferred ShowView failed: " + ex.Message);
                    }
                }, null);
                return null;
            }
            finally { LeaveBusy(); }
        }

        /// <summary>
        /// Force the IDE's lazy ABC class load to happen NOW (in isolation) so the user's first Modern
        /// Embeditor open doesn't pay it concurrently with the WebView2 open — the conflict that freezes
        /// Clarion. Reuses the proven open/cancel pattern: open the first procedure's native embeditor
        /// (whose source generation loads ABC), then CancelEmbeditor + WaitForEmbedClosed (which pumps
        /// Application.DoEvents so the native view actually tears down — the step my earlier raw cancel
        /// missed). NO Monaco view, NO WebView2 — that's what keeps it freeze-free. UI thread only.
        /// Returns a short diagnostic. Safe to call once per app load (guarded by IsBusy).
        /// </summary>
        public static string WarmupAbc()
        {
            EnterBusy();
            try
            {
                var appTree = new AppTreeService();
                var procs = appTree.GetProcedureNames();
                if (procs == null || procs.Count == 0)
                    return "ABC warmup skipped: no procedures in the open app.";
                string proc = procs[0];

                var sw = System.Diagnostics.Stopwatch.StartNew();
                string source, error;
                List<int[]> ranges;
                bool opened = OpenAndMirror(appTree, proc, out source, out ranges, out error);

                // Always close + WAIT (pumps DoEvents until GetEmbedInfo()==null) so the native embeditor
                // actually tears down — whether or not the mirror read succeeded, the open itself loads ABC.
                try { appTree.CancelEmbeditor(); } catch { }
                bool closed = WaitForEmbedClosed(appTree, 5000);
                appTree.ActivateAppView();
                sw.Stop();

                if (!opened)
                    return "ABC warmup: open of '" + proc + "' had trouble (" + error + ") after "
                           + sw.ElapsedMilliseconds + "ms; ABC may still have loaded. closed=" + closed;
                return "ABC warmup OK via '" + proc + "' — embed opened+closed in "
                       + sw.ElapsedMilliseconds + "ms (closed=" + closed + ").";
            }
            finally { LeaveBusy(); }
        }

        // Locator typing speed (ms/char): a quick first pass, then a slower, very reliable retry.
        // ClaList drops keystrokes typed too fast, so if the quick pass selects the wrong procedure
        // the verify step below catches it and we retry slower.
        private static readonly int[] CharDelaysMs = { 70, 130 };

        /// <summary>
        /// Reliably open the procedure's embeditor and mirror its source + editable-range map, leaving the
        /// embeditor OPEN on success (caller mirrors/edits then closes). Types the name into the locator at a
        /// quick speed first; if the WRONG procedure opened (keystrokes dropped), closes and retries slower.
        /// Verifies the opened source actually belongs to the procedure, so we never proceed on a mis-selected
        /// one. UI thread only.
        /// </summary>
        /// <summary>&gt;0 while an embeditor open/save is driving the IDE — pads should not auto-refresh then.</summary>
        private static int _busyCount;
        public static bool IsBusy { get { return _busyCount > 0; } }

        // TEMP diagnostic: append a phase timing to C:\Temp\modern-open-timing.log so we can break down where
        // the Modern-open wall-clock goes (open vs retry vs mirror vs pad-refresh vs WebView2). Remove once tuned.
        internal static void TimingLog(string label, long ms) { TimingMark(label + " = " + ms + " ms"); }
        internal static void TimingMark(string text)
        {
            try { System.IO.File.AppendAllText(@"C:\Temp\modern-open-timing.log",
                System.DateTime.Now.ToString("HH:mm:ss.fff") + "  " + text + "\r\n"); }
            catch { }
        }
        internal static void EnterBusy() { System.Threading.Interlocked.Increment(ref _busyCount); }
        internal static void LeaveBusy() { System.Threading.Interlocked.Decrement(ref _busyCount); }

        internal static bool OpenAndMirror(AppTreeService appTree, string procName,
            out string source, out List<int[]> ranges, out string error)
        {
            source = null; ranges = null; error = null;
            EnterBusy();
            try
            {
            for (int attempt = 0; attempt < CharDelaysMs.Length; attempt++)
            {
                if (!WaitForEmbedClosed(appTree, 3000))
                { error = "An embeditor is still open; close it and try again."; return false; }

                // Bring the app tree to the front so the native automation works even when a Modern
                // Embeditor tab is the active document.
                TimingMark("  attempt " + attempt + " (" + CharDelaysMs[attempt] + "ms/char) START");
                appTree.ActivateAppView();
                var swOpen = System.Diagnostics.Stopwatch.StartNew();
                appTree.OpenProcedureEmbed(procName, CharDelaysMs[attempt]);
                TimingLog("  attempt " + attempt + " OpenProcedureEmbed", swOpen.ElapsedMilliseconds);

                // First open loads the ABC libraries and can take many seconds; wait generously.
                var swWait = System.Diagnostics.Stopwatch.StartNew();
                if (!WaitForEmbedOpen(appTree, 45000))
                {
                    try { appTree.CancelEmbeditor(); } catch { }
                    error = "Embeditor did not open for '" + procName + "' within 45s.";
                    continue;
                }
                TimingLog("  attempt " + attempt + " WaitForEmbedOpen (ABC/native open)", swWait.ElapsedMilliseconds);

                var swMirror = System.Diagnostics.Stopwatch.StartNew();
                string title, ferr;
                if (!EmbeditorCompletionService.TryGetActiveEmbeditorSource(out title, out source, out ranges, out ferr))
                {
                    try { appTree.CancelEmbeditor(); } catch { }
                    error = "Could not read embed source for '" + procName + "': " + ferr;
                    continue;
                }
                TimingLog("  attempt " + attempt + " TryGetActiveEmbeditorSource (mirror)", swMirror.ElapsedMilliseconds);

                if (SourceMentionsProcedure(source, procName))
                { TimingMark("  attempt " + attempt + " VERIFIED OK"); return true; } // correct procedure — leave the embeditor open

                // Wrong procedure: keystrokes were dropped at this speed. Close and retry slower.
                try { appTree.CancelEmbeditor(); } catch { }
                TimingMark("  attempt " + attempt + " MIS-SELECTED — will retry slower");
                error = "Opened a different procedure than '" + procName + "' — the locator search missed.";
                source = null; ranges = null;
            }
            return false;
            }
            finally { LeaveBusy(); }
        }

        /// <summary>
        /// Sanity check that the assembled embed source belongs to the procedure: its own name appears in
        /// its generated source (e.g. "Name PROCEDURE"), so if it's absent we almost certainly opened the
        /// wrong procedure.
        /// </summary>
        private static bool SourceMentionsProcedure(string source, string procName)
        {
            if (string.IsNullOrEmpty(source) || string.IsNullOrWhiteSpace(procName)) return false;
            try
            {
                return System.Text.RegularExpressions.Regex.IsMatch(
                    source, @"\b" + System.Text.RegularExpressions.Regex.Escape(procName) + @"\b",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            }
            catch { return source.IndexOf(procName, StringComparison.OrdinalIgnoreCase) >= 0; }
        }

        // The native embed/ABC open + close are driven by the UI-thread MESSAGE LOOP. The old coarse
        // Sleep(50)-per-iteration starved that loop (we slept ~50ms of every tick), inflating a ~2s native
        // open to ~19s and leaving the embed half-settled for the close. These now pump Application.DoEvents()
        // at full speed (like the native click path) and only run the reflection-heavy GetEmbedInfo() poll
        // ~every 120ms, with a 1ms yield to avoid a 100% busy-spin.
        private const int EmbedPollIntervalMs = 120;

        internal static bool WaitForEmbedOpen(AppTreeService appTree, int timeoutMs)
        {
            var sw = System.Diagnostics.Stopwatch.StartNew();
            long lastPoll = -EmbedPollIntervalMs;
            while (sw.ElapsedMilliseconds < timeoutMs)
            {
                Application.DoEvents();
                if (sw.ElapsedMilliseconds - lastPoll >= EmbedPollIntervalMs)
                {
                    lastPoll = sw.ElapsedMilliseconds;
                    if (appTree.GetEmbedInfo() != null) return true;
                }
                Thread.Sleep(1);
            }
            return appTree.GetEmbedInfo() != null;
        }

        internal static bool WaitForEmbedClosed(AppTreeService appTree, int timeoutMs)
        {
            var sw = System.Diagnostics.Stopwatch.StartNew();
            long lastPoll = -EmbedPollIntervalMs;
            while (sw.ElapsedMilliseconds < timeoutMs)
            {
                Application.DoEvents();
                if (sw.ElapsedMilliseconds - lastPoll >= EmbedPollIntervalMs)
                {
                    lastPoll = sw.ElapsedMilliseconds;
                    if (appTree.GetEmbedInfo() == null) return true;
                }
                Thread.Sleep(1);
            }
            return appTree.GetEmbedInfo() == null;
        }
    }
}
