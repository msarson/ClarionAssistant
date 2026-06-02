using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using ICSharpCode.SharpDevelop.Gui;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using ClarionAssistant.Services;

namespace ClarionAssistant.Terminal
{
    /// <summary>
    /// Path B — Modern Embeditor (M1 spike, read-only render).
    /// Hosts a Monaco editor in WebView2 as a SharpDevelop view, showing the assembled
    /// embeditor source. Generation + parse-back + persistence remain Clarion-owned; this
    /// view is a parallel surface (mirror model — see docs/ModernEmbeditor-PathA.md, Path B).
    ///
    /// M1 scope: scaffold + render only. The editable-region map (read-only guard) and the
    /// save round-trip back through WriteEmbedContentByLine / SaveAndCloseEmbeditor are M2.
    ///
    /// Mirrors the proven WebView2-as-view pattern from DiffViewContent.cs: shared environment
    /// cache, virtual-host folder mapping for large-buffer transfer, and a JS to C# message bridge.
    /// </summary>
    public class ModernEmbeditorViewContent : AbstractViewContent
    {
        private Panel _panel;
        private WebView2 _webView;
        private bool _isInitialized;
        private bool _isInitializing;

        private string _title;
        private string _sourceText;
        private string _language;
        private bool _isDark = true;
        private List<int[]> _editableRanges; // 1-based inclusive [start,end] embed-slot ranges
        private readonly string _procedureName;     // set when opened from the picker (enables save)
        private List<string> _originalSlotTexts;     // baseline slot contents for change detection
        private readonly bool _saveEnabled;
        private readonly string _lspFileName;        // synthetic .clw URI for LSP completion/hover requests

        private string _tempDir;
        private const string VIRTUAL_HOST = "clarion-embeditor-data";

        // Find/Replace history scope: per-version (storage layer) + per-solution (folder) + per-procedure
        // (the "This procedure" group). Resolved once from the IDE when the page first asks for source.
        private string _histSolutionPath;
        private string _histProcKey;
        private bool _histScopeResolved;

        private static readonly List<ModernEmbeditorViewContent> _instances = new List<ModernEmbeditorViewContent>();

        public override Control Control { get { return _panel; } }

        /// <summary>The procedure this tab represents (null/empty in mirror mode).</summary>
        public string ProcedureName { get { return _procedureName; } }

        /// <summary>The Modern Embeditor tab that's currently the active document, or null.</summary>
        public static ModernEmbeditorViewContent ActiveModernView()
        {
            try
            {
                var wb = WorkbenchSingleton.Workbench;
                if (wb != null)
                {
                    // Reflect ActiveWorkbenchWindow -> ViewContent (the property is explicit-interface on the
                    // workbench itself, so GetProperty by name there returns null — go via the window).
                    var aw = GetProp(wb, "ActiveWorkbenchWindow");
                    if (aw != null)
                    {
                        var vc = GetProp(aw, "ActiveViewContent") ?? GetProp(aw, "ViewContent");
                        var m = vc as ModernEmbeditorViewContent;
                        if (m != null) return m;
                    }
                }
                // Fallback: if exactly one Modern Embeditor is open, it's unambiguous.
                lock (_instances) { if (_instances.Count == 1) return _instances[0]; }
                return null;
            }
            catch { return null; }
        }

        private static object GetProp(object obj, string name)
        {
            if (obj == null) return null;
            try
            {
                var p = obj.GetType().GetProperty(name,
                    BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
                return (p != null && p.GetIndexParameters().Length == 0) ? p.GetValue(obj, null) : null;
            }
            catch { return null; }
        }

        /// <summary>
        /// Data symbols (locals/globals/structures) for this procedure, from the LSP document-symbol tree
        /// over the opened source. Each entry: { name, kind, detail }. Empty if the LSP isn't running.
        /// </summary>
        public List<Dictionary<string, object>> GetDataSymbols()
        {
            var result = new List<Dictionary<string, object>>();
            try
            {
                var lsp = LspClient.Active;
                if (lsp == null) return result;
                var resp = lsp.GetDocumentSymbols(_lspFileName, _sourceText);
                object res = (resp != null && resp.ContainsKey("result")) ? resp["result"] : null;
                CollectSymbols(res, result);
            }
            catch (Exception ex) { System.Diagnostics.Debug.WriteLine("[ModernEmbeditor] GetDataSymbols: " + ex.Message); }
            return result;
        }

        // LSP documentSymbol returns either DocumentSymbol[] (hierarchical, has children) or
        // SymbolInformation[] (flat). Collect leaf names + kinds from either shape.
        private static void CollectSymbols(object node, List<Dictionary<string, object>> into)
        {
            var list = node as System.Collections.IEnumerable;
            if (list == null) return;
            foreach (var item in list)
            {
                var d = item as Dictionary<string, object>;
                if (d == null) continue;
                string name = d.ContainsKey("name") ? d["name"] as string : null;
                int kind = 0;
                if (d.ContainsKey("kind")) { try { kind = Convert.ToInt32(d["kind"]); } catch { } }
                string detail = d.ContainsKey("detail") ? d["detail"] as string : null;
                if (!string.IsNullOrEmpty(name))
                    into.Add(new Dictionary<string, object> { { "name", name }, { "kind", kind }, { "detail", detail } });
                if (d.ContainsKey("children")) CollectSymbols(d["children"], into);
            }
        }

        /// <summary>
        /// Combined data payload for the Modern Data pad: the procedure's local symbols (LSP) plus the
        /// dictionary tables it references (parsed from the generated &lt;app&gt;.clw, filtered to used ones).
        /// </summary>
        public Dictionary<string, object> GetPadData()
        {
            var locals = new List<Dictionary<string, object>>();
            var routines = new List<Dictionary<string, object>>();
            var globals = new List<Dictionary<string, object>>();
            try
            {
                foreach (var d in ClarionAppDataReader.ParseLocalData(_sourceText, _procedureName))
                    locals.Add(new Dictionary<string, object> { { "name", d.Name }, { "type", d.Type } });

                foreach (var r in ClarionAppDataReader.ParseRoutines(_sourceText, _procedureName))
                    routines.Add(new Dictionary<string, object> { { "name", r }, { "type", "ROUTINE" } });

                string appClw = ClarionAppDataReader.FindAppClwPath();
                if (appClw != null)
                    foreach (var g in ClarionAppDataReader.ParseGlobalData(appClw))
                        globals.Add(new Dictionary<string, object> { { "name", g.Name }, { "type", g.Type } });
            }
            catch (Exception ex) { System.Diagnostics.Debug.WriteLine("[ModernEmbeditor] GetPadData parse: " + ex.Message); }

            var moduleData = new List<Dictionary<string, object>>();
            try
            {
                string modClw = ClarionAppDataReader.FindModuleClwForProcedure(_procedureName);
                foreach (var d in ClarionAppDataReader.ParseModuleData(modClw))
                    moduleData.Add(new Dictionary<string, object> { { "name", d.Name }, { "type", d.Type } });
            }
            catch (Exception ex) { System.Diagnostics.Debug.WriteLine("[ModernEmbeditor] ParseModuleData: " + ex.Message); }

            var procedures = new List<Dictionary<string, object>>();
            try
            {
                foreach (var p in new AppTreeService().GetProcedureDetails())
                {
                    string n = (p != null && p.ContainsKey("name")) ? p["name"]?.ToString() : null;
                    if (!string.IsNullOrWhiteSpace(n))
                        procedures.Add(new Dictionary<string, object> { { "name", n } });
                }
            }
            catch (Exception ex) { System.Diagnostics.Debug.WriteLine("[ModernEmbeditor] procedures: " + ex.Message); }

            var data = new Dictionary<string, object>
            {
                { "procedure", _procedureName ?? "" },
                { "locals", locals },
                { "routines", routines },
                { "moduleData", moduleData },
                { "globals", globals },
                { "tables", GetUsedTables() },
                { "procedures", procedures }
            };
            return data;
        }

        /// <summary>Navigate this editor to a ROUTINE's declaration (Modern Data pad "go to routine" button).</summary>
        public void GotoRoutine(string name)
        {
            if (string.IsNullOrEmpty(name)) return;
            Action post = () =>
            {
                if (_webView == null || _webView.CoreWebView2 == null) return;
                try { _webView.CoreWebView2.PostWebMessageAsJson("{\"type\":\"gotoRoutine\",\"name\":" + JsonString(name) + "}"); }
                catch { }
            };
            try { if (_panel != null && _panel.InvokeRequired) _panel.BeginInvoke(post); else post(); }
            catch { }
        }

        /// <summary>If a Modern Embeditor tab for this procedure is already open, focus it. Returns true if found.</summary>
        public static bool TryFocusExisting(string procName)
        {
            if (string.IsNullOrWhiteSpace(procName)) return false;
            lock (_instances)
            {
                foreach (var inst in _instances)
                {
                    if (string.Equals(inst._procedureName, procName, StringComparison.OrdinalIgnoreCase))
                    {
                        inst.BringToFront();
                        return true;
                    }
                }
            }
            return false;
        }

        private List<Dictionary<string, object>> GetUsedTables()
        {
            var outp = new List<Dictionary<string, object>>();
            try
            {
                string appClw = ClarionAppDataReader.FindAppClwPath();
                if (appClw == null) return outp;
                var tables = ClarionAppDataReader.ParseTables(appClw);
                string src = _sourceText ?? "";
                foreach (var t in tables)
                {
                    if (!IsTableUsed(t, src)) continue;
                    var cols = new List<Dictionary<string, object>>();
                    foreach (var f in t.Fields)
                        cols.Add(new Dictionary<string, object> { { "name", f.Name }, { "type", f.Type } });
                    outp.Add(new Dictionary<string, object>
                    {
                        { "name", t.Name }, { "prefix", t.Prefix }, { "columns", cols }, { "keys", t.Keys }
                    });
                }
            }
            catch (Exception ex) { System.Diagnostics.Debug.WriteLine("[ModernEmbeditor] GetUsedTables: " + ex.Message); }
            return outp;
        }

        // A table is "used" by the procedure if its PRE: prefix or its name appears in the mirrored source.
        private static bool IsTableUsed(ClarionAppDataReader.TableDef t, string src)
        {
            if (string.IsNullOrEmpty(src)) return false;
            if (!string.IsNullOrEmpty(t.Prefix) && src.IndexOf(t.Prefix + ":", StringComparison.OrdinalIgnoreCase) >= 0)
                return true;
            if (!string.IsNullOrEmpty(t.Name) &&
                System.Text.RegularExpressions.Regex.IsMatch(src,
                    @"\b" + System.Text.RegularExpressions.Regex.Escape(t.Name) + @"\b",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase))
                return true;
            return false;
        }

        /// <summary>Insert text at the editor's cursor (used by the Modern Data pad's double-click-insert).</summary>
        public void InsertAtCursor(string text)
        {
            if (string.IsNullOrEmpty(text)) return;
            Action post = () =>
            {
                if (_webView == null || _webView.CoreWebView2 == null) return;
                try
                {
                    string json = "{\"type\":\"insertText\",\"text\":" + JsonString(text) + "}";
                    _webView.CoreWebView2.PostWebMessageAsJson(json);
                }
                catch { }
            };
            try { if (_panel != null && _panel.InvokeRequired) _panel.BeginInvoke(post); else post(); }
            catch { }
        }

        public ModernEmbeditorViewContent(string title, string sourceText, List<int[]> editableRanges,
            string language = "clarion", bool isDark = true, string procedureName = null)
        {
            _title = title ?? "Embeditor";
            _sourceText = sourceText ?? "";
            _editableRanges = editableRanges ?? new List<int[]>();
            _language = language ?? "clarion";
            _isDark = isDark;
            _procedureName = procedureName;
            _saveEnabled = !string.IsNullOrWhiteSpace(procedureName);
            _originalSlotTexts = ModernEmbeditorSaver.ExtractSlotTexts(_sourceText, _editableRanges);
            _lspFileName = MakeLspFileName(procedureName);
            TitleName = "Modern: " + _title;

            _panel = new Panel { Dock = DockStyle.Fill, BackColor = isDark ? Color.FromArgb(30, 30, 46) : Color.FromArgb(239, 241, 245) };
            // Plain WebView2 — Monaco's native mouseWheelZoom handles Ctrl+wheel inside the
            // renderer (a WinForms WndProc override never sees WebView2's inner Chrome wheel msg).
            _webView = new WebView2 { Dock = DockStyle.Fill };
            _panel.Controls.Add(_webView);

            lock (_instances) { _instances.Add(this); }
            _panel.HandleCreated += OnHandleCreated;
        }

        private async void OnHandleCreated(object sender, EventArgs e)
        {
            if (_isInitializing || _isInitialized) return;
            _isInitializing = true;

            try
            {
                var environment = await WebView2EnvironmentCache.GetEnvironmentAsync();
                await _webView.EnsureCoreWebView2Async(environment);

                _tempDir = Path.Combine(Path.GetTempPath(), "ClarionEmbeditor_" + Guid.NewGuid().ToString("N").Substring(0, 8));
                Directory.CreateDirectory(_tempDir);
                _webView.CoreWebView2.SetVirtualHostNameToFolderMapping(
                    VIRTUAL_HOST, _tempDir,
                    CoreWebView2HostResourceAccessKind.Allow);

                var settings = _webView.CoreWebView2.Settings;
                settings.IsScriptEnabled = true;
                settings.AreDefaultContextMenusEnabled = false;
                settings.AreDevToolsEnabled = true;
                settings.IsStatusBarEnabled = false;
                settings.IsZoomControlEnabled = false;
                settings.AreBrowserAcceleratorKeysEnabled = false; // let Monaco own Ctrl+S, not the browser

                _webView.CoreWebView2.WebMessageReceived += OnWebMessageReceived;
                _webView.CoreWebView2.NavigationCompleted += OnNavigationCompleted;

                string htmlPath = GetHtmlPath();
                if (File.Exists(htmlPath))
                    _webView.CoreWebView2.Navigate(new Uri(htmlPath).AbsoluteUri + "?v=" + File.GetLastWriteTimeUtc(htmlPath).Ticks);
            }
            catch (Exception ex)
            {
                _isInitializing = false; // allow retry
                System.Diagnostics.Debug.WriteLine("[ModernEmbeditorViewContent] Init error: " + ex.Message);
            }
        }

        private void OnNavigationCompleted(object sender, CoreWebView2NavigationCompletedEventArgs e)
        {
            _isInitialized = e.IsSuccess;
            _isInitializing = false;
            // SendSource is triggered by the JS "ready" message, not here — avoids double-send.
        }

        private void OnWebMessageReceived(object sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            try
            {
                string json = e.TryGetWebMessageAsString();
                string action = ExtractJsonValue(json, "action");
                if (action == "ready")
                    SendSource();
                else if (action == "save")
                    HandleSave(json);
                else if (action == "clipboard")
                    HandleClipboard(json);
                else if (action == "completion")
                    HandleCompletion(json);
                else if (action == "hover")
                    HandleHover(json);
                else if (action == "diagnostics")
                    HandleDiagnostics(json);
                else if (action == "saveSettings")
                    HandleSaveSettings(json);
                else if (action == "saveHistory")
                    HandleSaveHistory(json);
                else if (action == "saveCursor")
                    HandleSaveCursor(json);
                else if (action == "saveBookmarks")
                    HandleSaveBookmarks(json);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[ModernEmbeditorViewContent] Message error: " + ex.Message);
            }
        }

        /// <summary>Persist the user's edits: parse the per-slot payload and run the save round-trip.</summary>
        private void HandleSave(string json)
        {
            if (!_saveEnabled || string.IsNullOrWhiteSpace(_procedureName))
            {
                PostSaveResult(false, "Save isn't available — this tab was opened in mirror mode, not from the procedure picker.");
                return;
            }

            List<string> current;
            try
            {
                var ser = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };
                var data = ser.DeserializeObject(json) as Dictionary<string, object>;
                var arr = (data != null && data.ContainsKey("slots")) ? data["slots"] as object[] : null;
                if (arr == null) { PostSaveResult(false, "Save failed: malformed payload (no slots)."); return; }
                current = arr.Select(o => o == null ? "" : o.ToString()).ToList();
            }
            catch (Exception ex)
            {
                PostSaveResult(false, "Save failed parsing the editor payload: " + ex.Message);
                return;
            }

            bool ok;
            string msg = ModernEmbeditorSaver.Save(_procedureName, _editableRanges, _originalSlotTexts, current, out ok);
            // On success, the saved content is the new baseline so a follow-up save sees no changes.
            if (ok && current.Count == _originalSlotTexts.Count) _originalSlotTexts = current;
            // The save activated the app tree to drive the embeditor — bring this tab back to the front.
            BringToFront();
            PostSaveResult(ok, msg);
        }

        /// <summary>Re-select this view's tab (the save round-trip activates the app tree to drive the embeditor).</summary>
        private void BringToFront()
        {
            try
            {
                var w = WorkbenchWindow;
                if (w != null) w.GetType().GetMethod("SelectWindow", Type.EmptyTypes)?.Invoke(w, null);
            }
            catch { }
        }

        /// <summary>
        /// LSP completion request from Monaco. Uses the context-free language set (keywords, builtins,
        /// datatypes, attributes, controls) — no per-keystroke buffer sync needed. Runs off the UI thread
        /// and posts the result back keyed by reqId.
        /// </summary>
        /// <summary>
        /// Kick the shared LSP self-heal (idempotent, fire-and-forget) when no client is running yet.
        /// Mirrors the native embeditor's completion-time self-heal (EmbeditorCompletionService.LspStarter,
        /// wired to EnsureLspRunningInBackground) so the Modern editor can also recover the language server —
        /// completion, hover, AND the LSP diagnostics pass all depend on it. The first request after a cold
        /// start still returns empty (server warming); the next one succeeds.
        /// </summary>
        private static void EnsureLspStarted()
        {
            try
            {
                var lsp = LspClient.Active;
                if (lsp == null || !lsp.IsRunning)
                    EmbeditorCompletionService.LspStarter?.Invoke();
            }
            catch { }
        }

        private void HandleCompletion(string json)
        {
            int reqId, line, column;
            if (!ParseRequest(json, out reqId, out line, out column, out _)) return;
            Task.Run(() =>
            {
                var items = new List<Dictionary<string, object>>();
                string lspStatus;
                try
                {
                    EnsureLspStarted();
                    var lsp = LspClient.Active;
                    if (lsp == null) lspStatus = "not started";
                    else if (!lsp.IsRunning) lspStatus = "starting";
                    else
                    {
                        // Context-free: pass no buffer; the server returns the language item set.
                        var comps = lsp.GetCompletion(_lspFileName, Math.Max(0, line - 1), Math.Max(0, column - 1), 2500, null);
                        if (comps != null)
                            foreach (var c in comps)
                                items.Add(new Dictionary<string, object>
                                {
                                    { "label", c.Label },
                                    { "kind", c.Kind },
                                    { "detail", c.Detail },
                                    { "documentation", c.Documentation },
                                    { "insertText", c.InsertText }
                                });
                        lspStatus = string.IsNullOrEmpty(lsp.LastCompletionDiagnostic) ? "ok" : lsp.LastCompletionDiagnostic;
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine("[ModernEmbeditor] completion: " + ex.Message);
                    lspStatus = "error: " + ex.Message;
                }
                PostResponse(reqId, new Dictionary<string, object> { { "items", items }, { "lsp", lspStatus } });
            });
        }

        /// <summary>LSP hover request from Monaco. Syncs the current buffer (needed to resolve the symbol).</summary>
        private void HandleHover(string json)
        {
            int reqId, line, column; string buffer;
            if (!ParseRequest(json, out reqId, out line, out column, out buffer)) return;
            Task.Run(() =>
            {
                string contents = null;
                try
                {
                    EnsureLspStarted();
                    var lsp = LspClient.Active;
                    if (lsp != null && lsp.IsRunning)
                    {
                        var resp = lsp.GetHover(_lspFileName, Math.Max(0, line - 1), Math.Max(0, column - 1), buffer);
                        contents = ExtractHoverString(resp);
                    }
                }
                catch (Exception ex) { System.Diagnostics.Debug.WriteLine("[ModernEmbeditor] hover: " + ex.Message); }
                PostResponse(reqId, new Dictionary<string, object> { { "contents", contents } });
            });
        }

        /// <summary>
        /// Diagnostics request from Monaco (debounced after edits + once after load). Runs the hybrid
        /// ModernEmbeditorDiagnostics over the LIVE buffer + LIVE editable ranges — Monaco passes its
        /// decoration-tracked ranges because slots grow as the user types, so the load-time
        /// _editableRanges snapshot would be stale. Runs off the UI thread (the LSP sub-pass blocks),
        /// then posts back a unified marker list for setModelMarkers.
        /// </summary>
        private void HandleDiagnostics(string json)
        {
            int reqId; string buffer; List<int[]> ranges;
            if (!ParseDiagnosticsRequest(json, out reqId, out buffer, out ranges)) return;
            Task.Run(() =>
            {
                var markers = new List<Dictionary<string, object>>();
                try
                {
                    markers = ModernEmbeditorDiagnostics.Compute(
                        _lspFileName,
                        buffer ?? _sourceText,
                        (ranges != null && ranges.Count > 0) ? ranges : _editableRanges,
                        _procedureName);
                }
                catch (Exception ex) { System.Diagnostics.Debug.WriteLine("[ModernEmbeditor] diagnostics: " + ex.Message); }
                PostResponse(reqId, new Dictionary<string, object> { { "markers", markers } });
            });
        }

        // Parses a diagnostics request: reqId, the live buffer text, and the live editable ranges
        // (an array of [start,end] line pairs from Monaco's tracked decorations).
        private bool ParseDiagnosticsRequest(string json, out int reqId, out string buffer, out List<int[]> ranges)
        {
            reqId = 0; buffer = null; ranges = null;
            try
            {
                var data = new JavaScriptSerializer { MaxJsonLength = int.MaxValue }.DeserializeObject(json) as Dictionary<string, object>;
                if (data == null) return false;
                if (data.ContainsKey("reqId")) reqId = Convert.ToInt32(data["reqId"]);
                if (data.ContainsKey("buffer")) buffer = data["buffer"] as string;
                if (data.ContainsKey("ranges"))
                {
                    var arr = data["ranges"] as object[];
                    if (arr != null)
                    {
                        ranges = new List<int[]>();
                        foreach (var item in arr)
                        {
                            var pair = item as object[];
                            if (pair != null && pair.Length >= 2)
                                ranges.Add(new[] { Convert.ToInt32(pair[0]), Convert.ToInt32(pair[1]) });
                        }
                    }
                }
                return true;
            }
            catch { return false; }
        }

        /// <summary>
        /// Persist the dev's editor settings (from the gear panel) and broadcast them to every open
        /// Modern Embeditor tab so the change is consistent across tabs. Persist failures are logged but
        /// don't block the broadcast — the live editors still reflect the new options for this session.
        /// </summary>
        // Small fixed-cap parse for the tiny save* bridge payloads (cursor / bookmarks / settings). These
        // are page-supplied (untrusted) and bounded by design — refuse to materialize an oversized payload
        // BEFORE deserializing rather than trimming after. (Security gate finding.)
        private const int MaxBridgeJsonBytes = 65536;   // 64 KB — far above any legit save* message
        private static Dictionary<string, object> ParseBoundedBridgeJson(string json)
        {
            if (string.IsNullOrEmpty(json) || json.Length > MaxBridgeJsonBytes) return null;
            try { return new JavaScriptSerializer { MaxJsonLength = MaxBridgeJsonBytes }.DeserializeObject(json) as Dictionary<string, object>; }
            catch { return null; }
        }

        private void HandleSaveSettings(string json)
        {
            try
            {
                var data = ParseBoundedBridgeJson(json);
                var sd = (data != null && data.ContainsKey("settings")) ? data["settings"] as Dictionary<string, object> : null;
                if (sd == null) return;
                var settings = ModernEmbeditorSettings.FromDict(sd);
                try { settings.Save(); }
                catch (Exception ex) { System.Diagnostics.Debug.WriteLine("[ModernEmbeditor] saveSettings persist: " + ex.Message); }
                ApplySettingsToAll(settings); // broadcast to every open tab (incl. this one — idempotent)
            }
            catch (Exception ex) { System.Diagnostics.Debug.WriteLine("[ModernEmbeditor] saveSettings: " + ex.Message); }
        }

        /// <summary>Push the given settings to this tab's Monaco (gear panel + live updateOptions).</summary>
        public void ApplySettings(ModernEmbeditorSettings settings)
        {
            if (settings == null) return;
            string sjson;
            try { sjson = new JavaScriptSerializer().Serialize(settings.ToDict()); }
            catch { return; }
            Action post = () =>
            {
                if (_webView == null || _webView.CoreWebView2 == null) return;
                try { _webView.CoreWebView2.PostWebMessageAsJson("{\"type\":\"applySettings\",\"settings\":" + sjson + "}"); }
                catch { }
            };
            try { if (_panel != null && _panel.InvokeRequired) _panel.BeginInvoke(post); else post(); }
            catch { }
        }

        /// <summary>Broadcast editor settings to all open Modern Embeditor tabs (mirrors ApplyThemeToAll).</summary>
        public static void ApplySettingsToAll(ModernEmbeditorSettings settings)
        {
            lock (_instances) { foreach (var inst in _instances) inst.ApplySettings(settings); }
        }

        /// <summary>
        /// Persist the Find/Replace dropdown history (sent by JS as full arrays) and broadcast the saved
        /// lists to every open tab so all tabs converge. The incoming list is authoritative, so per-entry
        /// delete and "clear history" stick. Persist failures are logged but never block the broadcast.
        /// </summary>
        private void HandleSaveHistory(string json)
        {
            try
            {
                var data = new JavaScriptSerializer { MaxJsonLength = int.MaxValue }.DeserializeObject(json) as Dictionary<string, object>;
                if (data == null) return;
                var find = ToStringList(data, "find");
                var replace = ToStringList(data, "replace");
                var proc = ToStringList(data, "proc");
                EnsureHistoryScope();
                List<string> savedFind, savedReplace;
                ModernEmbeditorHistory.Save(_histSolutionPath, _histProcKey, find, replace, proc, out savedFind, out savedReplace);
                // Broadcast solution-wide lists only — each tab keeps its own procedure's recent terms.
                ApplyHistoryToAll(savedFind, savedReplace);
            }
            catch (Exception ex) { System.Diagnostics.Debug.WriteLine("[ModernEmbeditor] saveHistory: " + ex.Message); }
        }

        /// <summary>Persist the cursor position (sent on Ctrl+S) per solution+procedure for restore-on-open.</summary>
        private void HandleSaveCursor(string json)
        {
            try
            {
                var data = ParseBoundedBridgeJson(json);
                if (data == null) return;
                int line = data.ContainsKey("line") ? Convert.ToInt32(data["line"]) : 0;
                int column = data.ContainsKey("column") ? Convert.ToInt32(data["column"]) : 0;
                if (line < 1) return;
                EnsureHistoryScope();
                ModernEmbeditorState.SaveCursor(_histSolutionPath, _histProcKey, line, column);
            }
            catch (Exception ex) { System.Diagnostics.Debug.WriteLine("[ModernEmbeditor] saveCursor: " + ex.Message); }
        }

        /// <summary>Persist the bookmark line set (sent whenever it changes) per solution+procedure.</summary>
        private void HandleSaveBookmarks(string json)
        {
            try
            {
                var data = ParseBoundedBridgeJson(json);
                if (data == null) return;
                var lines = new List<int>();
                object o;
                if (data.TryGetValue("bookmarks", out o) && o is object[])
                {
                    // Bound ingestion: stop collecting once we have comfortably more than the persist cap
                    // (200) so a hostile/oversized array from the page can't force a huge allocation before
                    // CleanLines trims it. (Security gate finding.)
                    var arr = (object[])o;
                    for (int i = 0; i < arr.Length && lines.Count < 1000; i++)
                        if (arr[i] != null) { try { lines.Add(Convert.ToInt32(arr[i])); } catch { } }
                }
                EnsureHistoryScope();
                ModernEmbeditorState.SaveBookmarks(_histSolutionPath, _histProcKey, lines);
            }
            catch (Exception ex) { System.Diagnostics.Debug.WriteLine("[ModernEmbeditor] saveBookmarks: " + ex.Message); }
        }

        /// <summary>
        /// Resolve (once) the history scope from the IDE: the open solution (folder) and an app::procedure
        /// key (the "This procedure" group). Cached for this tab's lifetime.
        /// </summary>
        private void EnsureHistoryScope()
        {
            if (_histScopeResolved) return;
            _histScopeResolved = true;
            try { _histSolutionPath = EditorService.GetOpenSolutionPath(); } catch { _histSolutionPath = null; }
            string appName = null;
            try
            {
                var info = new AppTreeService().GetAppInfo();
                if (info != null && info.ContainsKey("name") && info["name"] != null) appName = info["name"].ToString();
            }
            catch { }
            string key = ((appName ?? "") + "::" + (_procedureName ?? "")).Trim(':');
            _histProcKey = string.IsNullOrEmpty(key) ? "" : key;
        }

        /// <summary>Coerce a JSON array field (object[] from DeserializeObject) into a string list.</summary>
        private static List<string> ToStringList(Dictionary<string, object> d, string key)
        {
            var res = new List<string>();
            object o;
            if (d != null && d.TryGetValue(key, out o) && o is object[])
            {
                foreach (var item in (object[])o)
                    if (item != null) res.Add(item.ToString());
            }
            return res;
        }

        /// <summary>Push Find/Replace history to this tab's dropdowns.</summary>
        public void ApplyHistory(IList<string> find, IList<string> replace)
        {
            string fj = ModernEmbeditorHistory.ToJson(find);
            string rj = ModernEmbeditorHistory.ToJson(replace);
            Action post = () =>
            {
                if (_webView == null || _webView.CoreWebView2 == null) return;
                try { _webView.CoreWebView2.PostWebMessageAsJson("{\"type\":\"applyHistory\",\"find\":" + fj + ",\"replace\":" + rj + "}"); }
                catch { }
            };
            try { if (_panel != null && _panel.InvokeRequired) _panel.BeginInvoke(post); else post(); }
            catch { }
        }

        /// <summary>Broadcast Find/Replace history to all open Modern Embeditor tabs.</summary>
        public static void ApplyHistoryToAll(IList<string> find, IList<string> replace)
        {
            lock (_instances) { foreach (var inst in _instances) inst.ApplyHistory(find, replace); }
        }

        private bool ParseRequest(string json, out int reqId, out int line, out int column, out string buffer)
        {
            reqId = 0; line = 0; column = 0; buffer = null;
            try
            {
                var data = new JavaScriptSerializer { MaxJsonLength = int.MaxValue }.DeserializeObject(json) as Dictionary<string, object>;
                if (data == null) return false;
                if (data.ContainsKey("reqId")) reqId = Convert.ToInt32(data["reqId"]);
                if (data.ContainsKey("line")) line = Convert.ToInt32(data["line"]);
                if (data.ContainsKey("column")) column = Convert.ToInt32(data["column"]);
                if (data.ContainsKey("buffer")) buffer = data["buffer"] as string;
                return true;
            }
            catch { return false; }
        }

        /// <summary>Posts a {type:"response", reqId, data} message back to Monaco (marshaled to the UI thread).</summary>
        private void PostResponse(int reqId, Dictionary<string, object> data)
        {
            Action post = () =>
            {
                if (_webView == null || _webView.CoreWebView2 == null) return;
                try
                {
                    var ser = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };
                    string json = ser.Serialize(new Dictionary<string, object>
                    {
                        { "type", "response" }, { "reqId", reqId }, { "data", data }
                    });
                    _webView.CoreWebView2.PostWebMessageAsJson(json);
                }
                catch { }
            };
            try { if (_panel != null && _panel.InvokeRequired) _panel.BeginInvoke(post); else post(); }
            catch { }
        }

        /// <summary>Pulls a plain string out of an LSP textDocument/hover response (MarkupContent/string/array).</summary>
        private static string ExtractHoverString(Dictionary<string, object> resp)
        {
            if (resp == null) return null;
            object result = resp.ContainsKey("result") ? resp["result"] : null;
            var rd = result as Dictionary<string, object>;
            object contents = rd != null && rd.ContainsKey("contents") ? rd["contents"] : result;
            return HoverPartToString(contents);
        }

        private static string HoverPartToString(object contents)
        {
            if (contents == null) return null;
            var s = contents as string;
            if (s != null) return s;
            var d = contents as Dictionary<string, object>;
            if (d != null && d.ContainsKey("value")) return d["value"] as string;
            var list = contents as System.Collections.IEnumerable;
            if (list != null)
            {
                var sb = new StringBuilder();
                foreach (var part in list)
                {
                    string p = HoverPartToString(part);
                    if (!string.IsNullOrEmpty(p)) { if (sb.Length > 0) sb.Append("\n\n"); sb.Append(p); }
                }
                return sb.Length > 0 ? sb.ToString() : null;
            }
            return null;
        }

        private static string MakeLspFileName(string procName)
        {
            string baseName = string.IsNullOrWhiteSpace(procName) ? "modern_embeditor" : procName;
            var sb = new StringBuilder();
            foreach (char c in baseName) sb.Append(char.IsLetterOrDigit(c) || c == '_' ? c : '_');
            return sb.ToString() + ".clw";
        }

        /// <summary>Put text on the Windows clipboard (Clarion-style Ctrl+X cut from the editor).</summary>
        private void HandleClipboard(string json)
        {
            try
            {
                var ser = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };
                var data = ser.DeserializeObject(json) as Dictionary<string, object>;
                string text = (data != null && data.ContainsKey("text")) ? (data["text"]?.ToString() ?? "") : null;
                if (text != null) Clipboard.SetText(text.Length == 0 ? " " : text);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[ModernEmbeditorViewContent] Clipboard error: " + ex.Message);
            }
        }

        private void PostSaveResult(bool ok, string message)
        {
            PostSaveResultOnce(ok, message);
            // Backup re-post on the next message-loop turn — delivery of the first post can race with
            // the embeditor open/close churn that just happened during the save.
            try { _panel?.BeginInvoke((Action)(() => PostSaveResultOnce(ok, message))); }
            catch { }
        }

        private void PostSaveResultOnce(bool ok, string message)
        {
            if (_webView == null || _webView.CoreWebView2 == null) return;
            string json = "{\"type\":\"saveResult\",\"ok\":" + (ok ? "true" : "false") +
                          ",\"message\":" + JsonString(message) + "}";
            try { _webView.CoreWebView2.PostWebMessageAsJson(json); } catch { }
        }

        /// <summary>Update the displayed source. Sends immediately if ready, else waits for the JS "ready".</summary>
        public void SetSource(string title, string sourceText, string language = null)
        {
            _title = title ?? _title;
            _sourceText = sourceText ?? "";
            if (language != null) _language = language;
            TitleName = "Modern: " + _title;

            if (_isInitialized)
                SendSource();
        }

        private void SendSource()
        {
            if (_webView.CoreWebView2 == null) return;

            // Warm the language server as soon as the editor opens, so completion/hover/LSP-diagnostics
            // are ready by the time the dev uses them (self-heal if eager-start never fired).
            EnsureLspStarted();

            try
            {
                // Transfer source via the virtual host (temp file) to avoid huge postMessage payloads.
                string sourceFile = Path.Combine(_tempDir, "source.txt");
                File.WriteAllText(sourceFile, _sourceText ?? "", Encoding.UTF8);

                string settingsJson;
                try { settingsJson = new JavaScriptSerializer().Serialize(ModernEmbeditorSettings.Load().ToDict()); }
                catch { settingsJson = "null"; }

                string findHistJson = "[]", replHistJson = "[]", procHistJson = "[]";
                int cursorLine = 0, cursorColumn = 0;
                string bookmarksJson = "[]";
                try
                {
                    EnsureHistoryScope();
                    List<string> hf, hr, hp;
                    ModernEmbeditorHistory.Load(_histSolutionPath, _histProcKey, out hf, out hr, out hp);
                    findHistJson = ModernEmbeditorHistory.ToJson(hf);
                    replHistJson = ModernEmbeditorHistory.ToJson(hr);
                    procHistJson = ModernEmbeditorHistory.ToJson(hp);
                    List<int> bms;
                    ModernEmbeditorState.Load(_histSolutionPath, _histProcKey, out cursorLine, out cursorColumn, out bms);
                    bookmarksJson = ModernEmbeditorState.BookmarksJson(bms);
                }
                catch { }

                string json = "{\"type\":\"setSource\"," +
                    "\"title\":" + JsonString(_title) + "," +
                    "\"language\":" + JsonString(_language) + "," +
                    "\"isDark\":" + (_isDark ? "true" : "false") + "," +
                    "\"saveEnabled\":" + (_saveEnabled ? "true" : "false") + "," +
                    "\"editableRanges\":" + RangesJson() + "," +
                    "\"settings\":" + settingsJson + "," +
                    "\"findHistory\":" + findHistJson + "," +
                    "\"replaceHistory\":" + replHistJson + "," +
                    "\"procHistory\":" + procHistJson + "," +
                    "\"cursorLine\":" + cursorLine + "," +
                    "\"cursorColumn\":" + cursorColumn + "," +
                    "\"bookmarks\":" + bookmarksJson + "," +
                    "\"sourceUrl\":\"https://" + VIRTUAL_HOST + "/source.txt\"}";
                _webView.CoreWebView2.PostWebMessageAsJson(json);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[ModernEmbeditorViewContent] SendSource error: " + ex.Message);
            }
        }

        public void ApplyTheme(bool isDark)
        {
            _isDark = isDark;
            if (_panel != null)
                _panel.BackColor = isDark ? Color.FromArgb(30, 30, 46) : Color.FromArgb(239, 241, 245);
            if (_isInitialized && _webView?.CoreWebView2 != null)
                _webView.CoreWebView2.PostWebMessageAsJson("{\"type\":\"applyTheme\",\"isDark\":" + (isDark ? "true" : "false") + "}");
        }

        public static void ApplyThemeToAll(bool isDark)
        {
            lock (_instances)
            {
                foreach (var inst in _instances)
                    inst.ApplyTheme(isDark);
            }
        }

        private string GetHtmlPath()
        {
            string assemblyDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            string path = Path.Combine(assemblyDir, "Terminal", "monaco-embeditor.html");
            if (File.Exists(path)) return path;
            path = Path.Combine(assemblyDir, "monaco-embeditor.html");
            if (File.Exists(path)) return path;
            return Path.Combine(assemblyDir, "Terminal", "monaco-embeditor.html");
        }

        /// <summary>Serializes the editable ranges as a JSON array of [start,end] pairs (1-based, inclusive).</summary>
        private string RangesJson()
        {
            if (_editableRanges == null || _editableRanges.Count == 0) return "[]";
            var sb = new StringBuilder("[");
            for (int i = 0; i < _editableRanges.Count; i++)
            {
                var r = _editableRanges[i];
                if (r == null || r.Length < 2) continue;
                if (sb.Length > 1) sb.Append(',');
                sb.Append('[').Append(r[0]).Append(',').Append(r[1]).Append(']');
            }
            sb.Append(']');
            return sb.ToString();
        }

        private static string JsonString(string s)
        {
            if (s == null) return "null";
            var sb = new StringBuilder(s.Length + 20);
            sb.Append('"');
            foreach (char c in s)
            {
                switch (c)
                {
                    case '\\': sb.Append("\\\\"); break;
                    case '"': sb.Append("\\\""); break;
                    case '\n': sb.Append("\\n"); break;
                    case '\r': sb.Append("\\r"); break;
                    case '\t': sb.Append("\\t"); break;
                    case '\b': sb.Append("\\b"); break;
                    case '\f': sb.Append("\\f"); break;
                    default:
                        if (c < ' ')
                            sb.AppendFormat("\\u{0:X4}", (int)c);
                        else
                            sb.Append(c);
                        break;
                }
            }
            sb.Append('"');
            return sb.ToString();
        }

        private static string ExtractJsonValue(string json, string key)
        {
            if (json == null) return null;
            string search = "\"" + key + "\":";
            int idx = json.IndexOf(search, StringComparison.Ordinal);
            if (idx < 0) return null;
            idx += search.Length;
            while (idx < json.Length && json[idx] == ' ') idx++;
            if (idx >= json.Length) return null;
            if (json[idx] == 'n') return null;
            if (json[idx] == '"')
            {
                idx++;
                var sb = new StringBuilder();
                while (idx < json.Length)
                {
                    char c = json[idx];
                    if (c == '\\' && idx + 1 < json.Length)
                    {
                        char next = json[idx + 1];
                        if (next == '"') { sb.Append('"'); idx += 2; continue; }
                        if (next == '\\') { sb.Append('\\'); idx += 2; continue; }
                        if (next == 'n') { sb.Append('\n'); idx += 2; continue; }
                        if (next == 'r') { sb.Append('\r'); idx += 2; continue; }
                        if (next == 't') { sb.Append('\t'); idx += 2; continue; }
                        sb.Append(c); idx++; continue;
                    }
                    if (c == '"') break;
                    sb.Append(c);
                    idx++;
                }
                return sb.ToString();
            }
            int start = idx;
            while (idx < json.Length && json[idx] != ',' && json[idx] != '}') idx++;
            return json.Substring(start, idx - start).Trim();
        }

        public override void Dispose()
        {
            lock (_instances) { _instances.Remove(this); }
            if (_webView != null)
            {
                _webView.Dispose();
                _webView = null;
            }
            if (_panel != null)
            {
                _panel.Dispose();
                _panel = null;
            }
            CleanupTempDir();
            base.Dispose();
        }

        private void CleanupTempDir()
        {
            if (_tempDir != null && Directory.Exists(_tempDir))
            {
                try { Directory.Delete(_tempDir, true); } catch { }
                _tempDir = null;
            }
        }
    }
}
