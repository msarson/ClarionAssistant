using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using ICSharpCode.SharpDevelop.Gui;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

namespace ClarionAssistant
{
    /// <summary>
    /// Modern Data pad (Path B). Lists the active Modern Embeditor tab's data symbols (locals/globals/
    /// structures, from the LSP document-symbol tree) and lets the developer double-click one to insert it
    /// at the cursor in that tab. Our managed replacement for Clarion's native Data/Tables field selector,
    /// which can't drive a non-ICSharpCode editor. Works with the snapshot model, so multi-open is preserved.
    /// </summary>
    public class ModernDataPad : AbstractPadContent
    {
        private Panel _panel;
        private WebView2 _webView;
        private bool _isInitialized;
        private bool _isInitializing;

        public override Control Control
        {
            get
            {
                if (_panel == null)
                {
                    _panel = new Panel { Dock = DockStyle.Fill, BackColor = Color.White };
                    _webView = new WebView2 { Dock = DockStyle.Fill };
                    _panel.Controls.Add(_webView);
                    _panel.HandleCreated += OnHandleCreated;
                }
                return _panel;
            }
        }

        private async void OnHandleCreated(object sender, EventArgs e)
        {
            if (_isInitializing || _isInitialized) return;
            _isInitializing = true;
            try
            {
                var environment = await Terminal.WebView2EnvironmentCache.GetEnvironmentAsync();
                await _webView.EnsureCoreWebView2Async(environment);

                var settings = _webView.CoreWebView2.Settings;
                settings.IsScriptEnabled = true;
                settings.AreDefaultContextMenusEnabled = false;
                settings.IsStatusBarEnabled = false;
                settings.AreBrowserAcceleratorKeysEnabled = false;

                _webView.CoreWebView2.WebMessageReceived += OnWebMessageReceived;
                StartAutoRefreshTimer();

                string htmlPath = GetHtmlPath();
                if (File.Exists(htmlPath))
                    _webView.CoreWebView2.Navigate(new Uri(htmlPath).AbsoluteUri + "?v=" + File.GetLastWriteTimeUtc(htmlPath).Ticks);
            }
            catch (Exception ex)
            {
                _isInitializing = false;
                System.Diagnostics.Debug.WriteLine("[ModernDataPad] Init error: " + ex.Message);
            }
        }

        private void OnWebMessageReceived(object sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            try
            {
                string json = e.TryGetWebMessageAsString();
                string action = ExtractJsonValue(json, "action");
                if (action == "ready") { _isInitialized = true; Refresh(); }
                else if (action == "refresh")
                {
                    // Explicit Refresh re-exports the whole-app .txa FIRST (UI thread, silent) so it picks up
                    // changes made outside the Modern Embeditor — e.g. dictionary/table edits — then re-parses
                    // the fresh cache. Deferred out of the WebView2 callback (like 'open') to avoid re-entrancy.
                    if (_panel != null)
                        _panel.BeginInvoke((Action)(() =>
                        {
                            Terminal.ModernEmbeditorViewContent.RefreshPadSources();
                            Refresh();
                        }));
                    else
                        Refresh();
                }
                else if (action == "insert")
                {
                    string name = ExtractJsonValue(json, "name");
                    var view = Terminal.ModernEmbeditorViewContent.ActiveModernView();
                    if (view != null && !string.IsNullOrEmpty(name)) view.InsertAtCursor(name);
                }
                else if (action == "open")
                {
                    string name = ExtractJsonValue(json, "name");
                    if (!string.IsNullOrEmpty(name) && _panel != null)
                    {
                        // Defer OUT of the WebView2 message callback — running the heavy open synchronously
                        // here re-enters the message pump. On the next UI turn: focus if already open, else open.
                        _panel.BeginInvoke((Action)(() =>
                        {
                            // Move focus off the pad's WebView2 to the main IDE window first — the native
                            // embeditor close (Discard) deadlocks if a WebView2 holds focus during it.
                            try
                            {
                                var mainForm = ICSharpCode.SharpDevelop.Gui.WorkbenchSingleton.Workbench as Form;
                                if (mainForm != null) { mainForm.Activate(); Application.DoEvents(); }
                            }
                            catch { }
                            try
                            {
                                if (!Terminal.ModernEmbeditorViewContent.TryFocusExisting(name))
                                    Services.ModernEmbeditorLauncher.OpenProcedure(name, false);
                            }
                            catch (Exception ex) { System.Diagnostics.Debug.WriteLine("[ModernDataPad] open: " + ex.Message); }
                            _lastShownProc = name; // about to show this proc; don't double-refresh on the next tick
                            Refresh();
                        }));
                    }
                }
                else if (action == "goto")
                {
                    string name = ExtractJsonValue(json, "name");
                    var view = Terminal.ModernEmbeditorViewContent.ActiveModernView();
                    if (view != null && !string.IsNullOrEmpty(name)) view.GotoRoutine(name);
                }
            }
            catch (Exception ex) { System.Diagnostics.Debug.WriteLine("[ModernDataPad] Message error: " + ex.Message); }
        }

        /// <summary>Pull the active Modern Embeditor tab's data (locals + used tables) off the UI thread.</summary>
        public void Refresh()
        {
            var view = Terminal.ModernEmbeditorViewContent.ActiveModernView();
            if (view == null)
            {
                Post(new Dictionary<string, object>
                {
                    { "title", "(no Modern Embeditor active)" },
                    { "locals", new List<object>() }, { "tables", new List<object>() }
                });
                return;
            }
            if (_refreshing) return; // coalesce — don't pile up overlapping parses
            _refreshing = true;
            Task.Run(() =>
            {
                try
                {
                    Dictionary<string, object> data;
                    try { data = view.GetPadData(); }
                    catch { data = new Dictionary<string, object> { { "locals", new List<object>() }, { "tables", new List<object>() } }; }
                    object proc; data.TryGetValue("procedure", out proc);
                    data["title"] = string.IsNullOrEmpty(proc as string) ? "Data" : (string)proc;
                    Post(data);
                }
                finally { _refreshing = false; }
            });
        }

        private void Post(Dictionary<string, object> data)
        {
            Action post = () =>
            {
                if (_webView == null || _webView.CoreWebView2 == null) return;
                try
                {
                    var ser = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };
                    data["type"] = "setSymbols";
                    _webView.CoreWebView2.PostWebMessageAsJson(ser.Serialize(data));
                }
                catch { }
            };
            try { if (_panel != null && _panel.InvokeRequired) _panel.BeginInvoke(post); else post(); }
            catch { }
        }

        private volatile bool _refreshing;
        private Timer _autoTimer;
        private string _lastShownProc = " "; // sentinel so the first tick always refreshes

        /// <summary>
        /// Auto-refresh via a timer (NOT a workbench event subscription). The event approach re-entered
        /// during the native embeditor close and deadlocked; a timer tick runs in a clean message-loop
        /// context and skips while an open/save is busy.
        /// </summary>
        private void StartAutoRefreshTimer()
        {
            if (_autoTimer != null) return;
            _autoTimer = new Timer { Interval = 750 };
            _autoTimer.Tick += OnAutoRefreshTick;
            _autoTimer.Start();
        }

        private void OnAutoRefreshTick(object sender, EventArgs e)
        {
            if (Services.ModernEmbeditorLauncher.IsBusy) return;  // never touch the IDE mid open/save
            string proc;
            try
            {
                var view = Terminal.ModernEmbeditorViewContent.ActiveModernView();
                proc = view != null ? (view.ProcedureName ?? "") : null;
            }
            catch { return; }
            // Refresh only when the active Modern procedure actually changed (or went away).
            if (!string.Equals(proc, _lastShownProc, StringComparison.OrdinalIgnoreCase))
            {
                _lastShownProc = proc;
                Refresh();
            }
        }

        public override void RedrawContent() { _lastShownProc = " "; Refresh(); }

        public override void Dispose()
        {
            if (_autoTimer != null) { _autoTimer.Stop(); _autoTimer.Dispose(); _autoTimer = null; }
            if (_webView != null) { _webView.Dispose(); _webView = null; }
            if (_panel != null) { _panel.Dispose(); _panel = null; }
            base.Dispose();
        }

        private static string GetHtmlPath()
        {
            string dir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            string path = Path.Combine(dir, "Terminal", "modern-data-pad.html");
            return File.Exists(path) ? path : Path.Combine(dir, "modern-data-pad.html");
        }

        private static string ExtractJsonValue(string json, string key)
        {
            if (json == null) return null;
            string search = "\"" + key + "\":";
            int idx = json.IndexOf(search, StringComparison.Ordinal);
            if (idx < 0) return null;
            idx += search.Length;
            while (idx < json.Length && json[idx] == ' ') idx++;
            if (idx >= json.Length || json[idx] != '"') return null;
            idx++;
            var sb = new System.Text.StringBuilder();
            while (idx < json.Length)
            {
                char c = json[idx];
                if (c == '\\' && idx + 1 < json.Length)
                {
                    char n = json[idx + 1];
                    if (n == '"') { sb.Append('"'); idx += 2; continue; }
                    if (n == '\\') { sb.Append('\\'); idx += 2; continue; }
                    if (n == 'n') { sb.Append('\n'); idx += 2; continue; }
                    if (n == 't') { sb.Append('\t'); idx += 2; continue; }
                    sb.Append(c); idx++; continue;
                }
                if (c == '"') break;
                sb.Append(c); idx++;
            }
            return sb.ToString();
        }
    }
}
