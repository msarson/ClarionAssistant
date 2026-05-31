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
                    _panel = new Panel { Dock = DockStyle.Fill, BackColor = Color.FromArgb(239, 241, 245) };
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
                else if (action == "refresh") Refresh();
                else if (action == "insert")
                {
                    string name = ExtractJsonValue(json, "name");
                    var view = Terminal.ModernEmbeditorViewContent.ActiveModernView();
                    if (view != null && !string.IsNullOrEmpty(name)) view.InsertAtCursor(name);
                }
            }
            catch (Exception ex) { System.Diagnostics.Debug.WriteLine("[ModernDataPad] Message error: " + ex.Message); }
        }

        /// <summary>Pull the active Modern Embeditor tab's data symbols (off the UI thread) and send to the pad.</summary>
        public void Refresh()
        {
            var view = Terminal.ModernEmbeditorViewContent.ActiveModernView();
            if (view == null) { Post(new Dictionary<string, object> { { "title", "(no Modern Embeditor active)" }, { "items", new List<object>() } }); return; }
            string title = "Data";
            Task.Run(() =>
            {
                List<Dictionary<string, object>> items;
                try { items = view.GetDataSymbols(); }
                catch { items = new List<Dictionary<string, object>>(); }
                Post(new Dictionary<string, object> { { "title", title }, { "items", items } });
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

        public override void RedrawContent() { Refresh(); }

        public override void Dispose()
        {
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
