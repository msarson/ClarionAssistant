using System;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

namespace ClarionAssistant.Terminal
{
    public class HeaderActionEventArgs : EventArgs
    {
        public string Action { get; private set; }
        public string Data { get; private set; }
        public HeaderActionEventArgs(string action, string data) { Action = action; Data = data; }
    }

    public class HeaderWebView : UserControl
    {
        private WebView2 _webView;
        private bool _isInitialized;
        private bool _isInitializing;

        public event EventHandler<HeaderActionEventArgs> ActionReceived;
        public event EventHandler HeaderReady;

        public bool IsReady { get { return _isInitialized; } }

        public HeaderWebView()
        {
            SuspendLayout();
            BackColor = Color.FromArgb(30, 30, 46);
            Height = 130;
            Dock = DockStyle.Top;

            _webView = new WebView2 { Dock = DockStyle.Fill, Name = "headerWebView" };
            Controls.Add(_webView);
            ResumeLayout(false);

            HandleCreated += OnHandleCreated;
        }

        private async void OnHandleCreated(object sender, EventArgs e)
        {
            if (_isInitializing || _isInitialized) return;
            _isInitializing = true;

            try
            {
                var environment = await WebView2EnvironmentCache.GetEnvironmentAsync();
                await _webView.EnsureCoreWebView2Async(environment);

                var settings = _webView.CoreWebView2.Settings;
                settings.IsScriptEnabled = true;
                settings.AreDefaultContextMenusEnabled = false;
                settings.AreDevToolsEnabled = true;
                settings.IsStatusBarEnabled = false;
                settings.AreBrowserAcceleratorKeysEnabled = false;

                _webView.CoreWebView2.WebMessageReceived += OnWebMessageReceived;
                _webView.CoreWebView2.NavigationCompleted += OnNavigationCompleted;

                string htmlPath = GetHtmlPath();
                if (File.Exists(htmlPath))
                    _webView.CoreWebView2.Navigate(new Uri(htmlPath).AbsoluteUri);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[HeaderWebView] Init error: " + ex.Message);
            }
        }

        private void OnNavigationCompleted(object sender, CoreWebView2NavigationCompletedEventArgs e)
        {
            _isInitialized = true;
            _isInitializing = false;
            HeaderReady?.Invoke(this, EventArgs.Empty);
        }

        private void OnWebMessageReceived(object sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            try
            {
                string json = e.TryGetWebMessageAsString();
                // Simple JSON parse — avoid dependency on JSON library
                string action = ExtractJsonValue(json, "action");
                string data = ExtractJsonValue(json, "data");
                if (!string.IsNullOrEmpty(action))
                    ActionReceived?.Invoke(this, new HeaderActionEventArgs(action, data));
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[HeaderWebView] Message error: " + ex.Message);
            }
        }

        /// <summary>Send a JSON message to the header JavaScript.</summary>
        public void SendMessage(string json)
        {
            if (!_isInitialized || _webView.CoreWebView2 == null) return;
            _webView.CoreWebView2.PostWebMessageAsString(json);
        }

        /// <summary>Set the version dropdown items.</summary>
        public void SetVersions(string[] labels, string[] values, int selectedIndex)
        {
            var items = new System.Text.StringBuilder("[");
            for (int i = 0; i < labels.Length; i++)
            {
                if (i > 0) items.Append(",");
                items.AppendFormat("{{\"label\":\"{0}\",\"value\":\"{1}\",\"selected\":{2}}}",
                    EscapeJson(labels[i]), EscapeJson(values[i]), i == selectedIndex ? "true" : "false");
            }
            items.Append("]");
            SendMessage("{\"type\":\"setVersions\",\"items\":" + items + "}");
        }

        /// <summary>Set the solution dropdown items.</summary>
        public void SetSolutions(string[] paths, int selectedIndex)
        {
            var items = new System.Text.StringBuilder("[");
            for (int i = 0; i < paths.Length; i++)
            {
                if (i > 0) items.Append(",");
                string label = paths[i].Length > 60
                    ? "..." + paths[i].Substring(paths[i].Length - 57)
                    : paths[i];
                items.AppendFormat("{{\"label\":\"{0}\",\"value\":\"{1}\",\"selected\":{2}}}",
                    EscapeJson(label), EscapeJson(paths[i]), i == selectedIndex ? "true" : "false");
            }
            items.Append("]");
            SendMessage("{\"type\":\"setSolutions\",\"items\":" + items + "}");
        }

        /// <summary>Update the MCP/status text in the header.</summary>
        public void SetStatus(string text, string cssClass = "")
        {
            SendMessage("{\"type\":\"setStatus\",\"text\":\"" + EscapeJson(text) + "\",\"css\":\"" + EscapeJson(cssClass) + "\"}");
        }

        /// <summary>Update the index status text.</summary>
        public void SetIndexStatus(string text, string cssClass = "")
        {
            SendMessage("{\"type\":\"setIndexStatus\",\"text\":\"" + EscapeJson(text) + "\",\"css\":\"" + EscapeJson(cssClass) + "\"}");
        }

        /// <summary>Enable or disable the index buttons.</summary>
        public void SetIndexButtonsEnabled(bool enabled)
        {
            SendMessage("{\"type\":\"setIndexButtons\",\"enabled\":" + (enabled ? "true" : "false") + "}");
        }

        private string GetHtmlPath()
        {
            string assemblyDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            string path = Path.Combine(assemblyDir, "Terminal", "header.html");
            if (File.Exists(path)) return path;
            path = Path.Combine(assemblyDir, "header.html");
            if (File.Exists(path)) return path;
            return Path.Combine(assemblyDir, "Terminal", "header.html");
        }

        public static string EscapeJsonStatic(string s) { return EscapeJson(s); }

        private static string EscapeJson(string s)
        {
            if (s == null) return "";
            return s.Replace("\\", "\\\\").Replace("\"", "\\\"").Replace("\n", "\\n").Replace("\r", "\\r");
        }

        private static string ExtractJsonValue(string json, string key)
        {
            string search = "\"" + key + "\":";
            int idx = json.IndexOf(search, StringComparison.Ordinal);
            if (idx < 0) return null;
            idx += search.Length;
            // Skip whitespace
            while (idx < json.Length && json[idx] == ' ') idx++;
            if (idx >= json.Length) return null;
            if (json[idx] == 'n') return null; // null
            if (json[idx] == '"')
            {
                idx++; // skip opening quote
                int end = json.IndexOf('"', idx);
                return end > idx ? json.Substring(idx, end - idx) : "";
            }
            // Number or boolean
            int start = idx;
            while (idx < json.Length && json[idx] != ',' && json[idx] != '}') idx++;
            return json.Substring(start, idx - start).Trim();
        }
    }
}
