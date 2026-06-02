using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Web.Script.Serialization;

namespace ClarionAssistant.Services
{
    /// <summary>
    /// Path B — Modern Embeditor: per-procedure editor STATE that should survive closing/reopening a tab —
    /// the last cursor position (saved on Ctrl+S, restored on first open) and the set of bookmark lines
    /// (saved whenever they change, restored on open). Stored next to the Find history, one file per
    /// version+solution:
    ///   %APPDATA%\ClarionAssistant\&lt;ClarionVersion&gt;\&lt;Solution&gt;\embed-state.json
    /// shape: { "&lt;app::procedure&gt;": { "cursorLine": N, "cursorColumn": N, "bookmarks": [N, ...] } }.
    /// The version/solution scope reuses <see cref="ModernEmbeditorHistory"/>'s tags. Writes are
    /// read-modify-write preserving every other procedure's record and the untouched field (cursor vs
    /// bookmarks), matching the history class's last-writer-wins simplicity — this is non-critical state.
    /// </summary>
    public static class ModernEmbeditorState
    {
        private const int BookmarkCap = 200;       // guard against an unbounded bookmark list bloating the file
        private const int MaxStateFileBytes = 2 * 1024 * 1024;  // refuse to parse an oversized state file (DoS guard)

        // Read + deserialize the state file with size/length caps so a hostile or corrupt file can't force
        // a huge parse. Returns null when missing, oversized, or unparseable. (Security gate finding.)
        private static Dictionary<string, object> ReadStateFile(string path)
        {
            try
            {
                if (!File.Exists(path)) return null;
                if (new FileInfo(path).Length > MaxStateFileBytes)
                {
                    System.Diagnostics.Debug.WriteLine("[ModernEmbeditorState] state file exceeds cap; ignoring: " + path);
                    return null;
                }
                return new JavaScriptSerializer { MaxJsonLength = MaxStateFileBytes }
                    .DeserializeObject(File.ReadAllText(path)) as Dictionary<string, object>;
            }
            catch (Exception ex) { System.Diagnostics.Debug.WriteLine("[ModernEmbeditorState] ReadStateFile: " + ex.Message); return null; }
        }

        /// <summary>Read the saved cursor + bookmarks for one procedure (zeros / empty when none).</summary>
        public static void Load(string solutionPath, string procKey,
            out int cursorLine, out int cursorColumn, out List<int> bookmarks)
        {
            cursorLine = 0; cursorColumn = 0; bookmarks = new List<int>();
            if (string.IsNullOrEmpty(procKey)) return;
            try
            {
                var root = ReadStateFile(FilePath(solutionPath));
                object pv;
                if (root == null || !root.TryGetValue(procKey, out pv)) return;
                var rec = pv as Dictionary<string, object>;
                if (rec == null) return;
                cursorLine = ToInt(rec, "cursorLine");
                cursorColumn = ToInt(rec, "cursorColumn");
                bookmarks = CleanLines(ToIntList(rec, "bookmarks"));
            }
            catch (Exception ex) { System.Diagnostics.Debug.WriteLine("[ModernEmbeditorState] Load: " + ex.Message); }
        }

        /// <summary>Persist only the cursor for one procedure, preserving its bookmarks + other procs.</summary>
        public static void SaveCursor(string solutionPath, string procKey, int line, int column)
        {
            Update(solutionPath, procKey, rec =>
            {
                rec["cursorLine"] = line < 1 ? 1 : line;
                rec["cursorColumn"] = column < 1 ? 1 : column;
            });
        }

        /// <summary>Persist only the bookmark lines for one procedure, preserving its cursor + other procs.</summary>
        public static void SaveBookmarks(string solutionPath, string procKey, IList<int> bookmarks)
        {
            Update(solutionPath, procKey, rec => { rec["bookmarks"] = CleanLines(bookmarks); });
        }

        // Read-modify-write the whole file: load (or start) the proc-keyed map, mutate this proc's record,
        // write it all back. Preserves every other procedure and any field the mutator doesn't touch.
        private static void Update(string solutionPath, string procKey, Action<Dictionary<string, object>> mutate)
        {
            if (string.IsNullOrEmpty(procKey)) return;
            try
            {
                string path = FilePath(solutionPath);
                var root = ReadStateFile(path) ?? new Dictionary<string, object>();
                var rec = (root.ContainsKey(procKey) ? root[procKey] as Dictionary<string, object> : null)
                          ?? new Dictionary<string, object>();
                mutate(rec);
                root[procKey] = rec;
                File.WriteAllText(path, new JavaScriptSerializer().Serialize(root), Encoding.UTF8);
            }
            catch (Exception ex) { System.Diagnostics.Debug.WriteLine("[ModernEmbeditorState] Update: " + ex.Message); }
        }

        public static string BookmarksJson(IList<int> bookmarks)
        {
            try { return new JavaScriptSerializer().Serialize(CleanLines(bookmarks)); }
            catch { return "[]"; }
        }

        /// <summary>Version+solution-scoped state file path (folders created on demand).</summary>
        private static string FilePath(string solutionPath)
        {
            string dir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "ClarionAssistant", ModernEmbeditorHistory.VersionTag(), ModernEmbeditorHistory.SolutionTag(solutionPath));
            if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
            return Path.Combine(dir, "embed-state.json");
        }

        // Distinct, positive, ascending, capped — keeps the bookmark list tidy and bounded.
        private static List<int> CleanLines(IList<int> lines)
        {
            var outp = new List<int>();
            if (lines == null) return outp;
            var seen = new HashSet<int>();
            foreach (var n in lines)
            {
                if (n < 1 || seen.Contains(n)) continue;
                seen.Add(n); outp.Add(n);
                if (outp.Count >= BookmarkCap) break;
            }
            outp.Sort();
            return outp;
        }

        private static int ToInt(IDictionary<string, object> d, string key)
        {
            object o;
            if (d != null && d.TryGetValue(key, out o) && o != null)
            {
                try { return Convert.ToInt32(o); } catch { }
            }
            return 0;
        }

        private static List<int> ToIntList(IDictionary<string, object> d, string key)
        {
            var res = new List<int>();
            object o;
            if (d != null && d.TryGetValue(key, out o) && o is object[])
            {
                foreach (var item in (object[])o)
                {
                    if (item == null) continue;
                    try { res.Add(Convert.ToInt32(item)); } catch { }
                }
            }
            return res;
        }
    }
}
