using System;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using ICSharpCode.SharpDevelop.Gui;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

namespace ClarionAssistant.Terminal
{
    /// <summary>
    /// ViewContent that hosts a Monaco diff editor in the IDE's main editor panel.
    /// Opens as a tab alongside .clw files.
    /// </summary>
    public class DiffViewContent : AbstractViewContent
    {
        private Panel _panel;
        private WebView2 _webView;
        private bool _isInitialized;
        private bool _isInitializing;

        private string _title;
        private string _originalText;
        private string _modifiedText;
        private string _language;
        private bool _ignoreWhitespace;
        private bool _pendingLoad;

        private string _tempDir;
        private const string VIRTUAL_HOST = "clarion-diff-data";

        /// <summary>Fires when the user clicks Apply with the final (possibly edited) text.</summary>
        public event Action<string> Applied;

        /// <summary>Fires when the user clicks Cancel.</summary>
        public event Action Cancelled;

        public override Control Control { get { return _panel; } }

        public DiffViewContent(string title, string originalText, string modifiedText, string language = "clarion",
            bool ignoreWhitespace = false)
        {
            _title = title ?? "Diff";
            _originalText = originalText ?? "";
            _modifiedText = modifiedText ?? "";
            _language = language ?? "clarion";
            _ignoreWhitespace = ignoreWhitespace;
            TitleName = "Diff: " + _title;

            _panel = new Panel { Dock = DockStyle.Fill, BackColor = Color.FromArgb(30, 30, 46) };
            _webView = new ZoomableWebView2 { Dock = DockStyle.Fill };
            _panel.Controls.Add(_webView);

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

                // Set up temp directory and virtual host mapping for large file transfer
                _tempDir = Path.Combine(Path.GetTempPath(), "ClarionDiff_" + Guid.NewGuid().ToString("N").Substring(0, 8));
                Directory.CreateDirectory(_tempDir);
                _webView.CoreWebView2.SetVirtualHostNameToFolderMapping(
                    VIRTUAL_HOST, _tempDir,
                    CoreWebView2HostResourceAccessKind.Allow);

                var settings = _webView.CoreWebView2.Settings;
                settings.IsScriptEnabled = true;
                settings.AreDefaultContextMenusEnabled = false;
                settings.AreDevToolsEnabled = true;
                settings.IsStatusBarEnabled = false;
                settings.AreBrowserAcceleratorKeysEnabled = false;
                settings.IsZoomControlEnabled = false;

                _webView.CoreWebView2.WebMessageReceived += OnWebMessageReceived;
                _webView.CoreWebView2.NavigationCompleted += OnNavigationCompleted;

                string htmlPath = GetHtmlPath();
                if (File.Exists(htmlPath))
                    _webView.CoreWebView2.Navigate(new Uri(htmlPath).AbsoluteUri + "?v=" + File.GetLastWriteTimeUtc(htmlPath).Ticks);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[DiffViewContent] Init error: " + ex.Message);
            }
        }

        private void OnNavigationCompleted(object sender, CoreWebView2NavigationCompletedEventArgs e)
        {
            _isInitialized = true;
            _isInitializing = false;

            // If we already have diff data queued, send it now
            if (_pendingLoad)
            {
                _pendingLoad = false;
                SendDiffData();
            }
        }

        private void OnWebMessageReceived(object sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            try
            {
                string json = e.TryGetWebMessageAsString();
                string action = ExtractJsonValue(json, "action");

                if (action == "ready")
                {
                    SendDiffData();
                }
                else if (action == "apply")
                {
                    string text = ExtractJsonValue(json, "text");
                    Applied?.Invoke(text);
                }
                else if (action == "cancel")
                {
                    Cancelled?.Invoke();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[DiffViewContent] Message error: " + ex.Message);
            }
        }

        /// <summary>Update the diff content (can be called before or after initialization).</summary>
        public void SetDiff(string title, string originalText, string modifiedText, string language = null)
        {
            _title = title ?? _title;
            _originalText = originalText ?? "";
            _modifiedText = modifiedText ?? "";
            if (language != null) _language = language;
            TitleName = "Diff: " + _title;

            if (_isInitialized)
                SendDiffData();
            else
                _pendingLoad = true;
        }

        private void SendDiffData()
        {
            if (_webView.CoreWebView2 == null) return;

            try
            {
                // Write texts to temp files so JavaScript can fetch them.
                // This avoids PostWebMessageAsJson corruption with large payloads (200KB+).
                string origFile = Path.Combine(_tempDir, "original.txt");
                string modFile = Path.Combine(_tempDir, "modified.txt");
                File.WriteAllText(origFile, _originalText ?? "", System.Text.Encoding.UTF8);
                File.WriteAllText(modFile, _modifiedText ?? "", System.Text.Encoding.UTF8);

                // Send only metadata — JavaScript fetches the actual content via virtual host URLs
                string json = "{\"type\":\"setDiff\"," +
                    "\"title\":" + JsonString(_title) + "," +
                    "\"language\":" + JsonString(_language) + "," +
                    "\"ignoreWhitespace\":" + (_ignoreWhitespace ? "true" : "false") + "," +
                    "\"originalUrl\":\"https://" + VIRTUAL_HOST + "/original.txt\"," +
                    "\"modifiedUrl\":\"https://" + VIRTUAL_HOST + "/modified.txt\"}";
                _webView.CoreWebView2.PostWebMessageAsJson(json);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[DiffViewContent] SendDiffData error: " + ex.Message);
            }
        }

        private string GetHtmlPath()
        {
            string assemblyDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            string path = Path.Combine(assemblyDir, "Terminal", "diff.html");
            if (File.Exists(path)) return path;
            path = Path.Combine(assemblyDir, "diff.html");
            if (File.Exists(path)) return path;
            return Path.Combine(assemblyDir, "Terminal", "diff.html");
        }

        private static string JsonString(string s)
        {
            if (s == null) return "null";
            var sb = new System.Text.StringBuilder(s.Length + 20);
            sb.Append('"');
            foreach (char c in s)
            {
                switch (c)
                {
                    case '\\': sb.Append("\\\\"); break;
                    case '"':  sb.Append("\\\""); break;
                    case '\n': sb.Append("\\n");  break;
                    case '\r': sb.Append("\\r");  break;
                    case '\t': sb.Append("\\t");  break;
                    case '\b': sb.Append("\\b");  break;
                    case '\f': sb.Append("\\f");  break;
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
                var sb = new System.Text.StringBuilder();
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

        /// <summary>
        /// WebView2 subclass that intercepts Ctrl+MouseWheel at the Win32 message level
        /// and forwards it to JavaScript for font size changes.
        /// WebView2 swallows WM_MOUSEWHEEL+Ctrl internally even when IsZoomControlEnabled=false,
        /// so the wheel event never reaches JavaScript. This override catches it first.
        /// </summary>
        private class ZoomableWebView2 : WebView2
        {
            private const int WM_MOUSEWHEEL = 0x020A;
            private const int MK_CONTROL = 0x0008;

            protected override void WndProc(ref Message m)
            {
                if (m.Msg == WM_MOUSEWHEEL)
                {
                    int wParam = m.WParam.ToInt32();
                    int keys = wParam & 0xFFFF;
                    if ((keys & MK_CONTROL) != 0 && CoreWebView2 != null)
                    {
                        short delta = (short)(wParam >> 16);
                        int change = delta > 0 ? 1 : -1;
                        CoreWebView2.ExecuteScriptAsync(
                            "if(typeof applyFontSize==='function')applyFontSize(savedFontSize+(" + change + "))");
                        return; // Consume — don't let WebView2 swallow it
                    }
                }
                base.WndProc(ref m);
            }
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
