using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace ClarionAssistant.Services
{
    /// <summary>
    /// Registers ClarionAssistant's local MCP server in <c>~/.codex/config.toml</c>
    /// so a Codex CLI terminal launched from CA can reach the IDE-driving tools.
    ///
    /// Codex CLI does not accept an <c>--mcp-config</c> flag (unlike Claude / Copilot);
    /// MCP servers must be declared in the user-global config.toml. CA's MCP server
    /// runs HTTP on localhost with a per-session bearer token, so the TOML block
    /// uses a stdio bridge via a globally-installed <c>mcp-remote</c> shim
    /// (Codex CLI's native HTTP transport is unreliable against Streamable-HTTP).
    ///
    /// SECURITY: We require a pre-installed, locally-resolved <c>mcp-remote.cmd</c>
    /// rather than letting <c>npx</c> auto-fetch from the public registry at launch
    /// time. The bridge process receives the live MCP bearer token, so a typosquat
    /// or registry compromise of <c>mcp-remote</c> would get authenticated access
    /// to the IDE bridge. Fail-closed if the shim isn't installed; surface an
    /// install instruction to the user instead.
    ///
    /// CONCURRENCY: A static lock serializes all in-process writers to the same
    /// config file (multiple Codex tabs launching simultaneously would otherwise
    /// race the read-modify-write cycle and clobber each other). Cross-process
    /// safety relies on <see cref="File.Replace(string,string,string)"/>'s atomic
    /// ReplaceFileW semantics on NTFS.
    ///
    /// MARKER COEXISTENCE: Codex CLI rewrites <c>config.toml</c> through its own
    /// TOML serializer, which silently drops the trailing end-marker comment
    /// when it appends state tables (e.g. <c>[tui.model_availability_nux]</c>).
    /// That leaves the file with a begin marker but no matching end marker; a
    /// naive replace-or-append would then APPEND a second managed block,
    /// producing duplicate <c>[mcp_servers.clarion-assistant]</c> tables and
    /// breaking TOML parsing. <see cref="ReplaceOrAppendManagedBlock"/> is
    /// therefore self-healing: it strips ALL CA markers and ALL
    /// <c>[mcp_servers.clarion-assistant.*]</c> sections anywhere in the file,
    /// lifts foreign tables out of any (possibly broken) managed region, and
    /// appends one fresh canonical block.
    /// </summary>
    public static class CodexConfigService
    {
        private const string ManagedMarkerBegin = "# >>> CLARIONASSISTANT MANAGED — do not edit (begin) <<<";
        private const string ManagedMarkerEnd   = "# <<< CLARIONASSISTANT MANAGED — do not edit (end) >>>";
        private const string ServerName         = "clarion-assistant";
        private const string ManagedTablePrefix = "mcp_servers." + ServerName;

        // Serializes in-process writers. Cross-process is handled by File.Replace's
        // atomic ReplaceFileW semantics; this lock prevents two CA tabs in the same
        // process from racing the read-modify-write cycle.
        private static readonly object _writeLock = new object();

        public static string GetCodexConfigPath()
        {
            string userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            return Path.Combine(userProfile, ".codex", "config.toml");
        }

        /// <summary>
        /// Refresh the managed MCP block in <c>~/.codex/config.toml</c>. Idempotent
        /// and content-diffing — no write if the existing block already matches.
        /// Returns the config path on success, or null on any failure (caller
        /// surfaces the error to the user).
        ///
        /// On failure, sets <paramref name="failureReason"/> with a single-line
        /// human-readable cause so the launcher can put it in the terminal banner.
        /// </summary>
        public static string EnsureMcpRegistration(string mcpUrl, string bearerToken, out string failureReason)
        {
            failureReason = null;

            if (string.IsNullOrWhiteSpace(mcpUrl))
            {
                failureReason = "MCP server URL is empty (server may have stopped before launch).";
                return null;
            }

            string mcpRemotePath = CodexProcessManager.FindMcpRemotePath();
            if (string.IsNullOrEmpty(mcpRemotePath))
            {
                failureReason = "mcp-remote not installed. Run: npm install -g mcp-remote@" + CodexProcessManager.McpRemoteVersion;
                return null;
            }

            // Version pin enforcement. Without this, an older mcp-remote satisfies
            // the presence check and CA writes a path that may behave differently
            // than the version we tested. Hard-fail with a distinct reason so the
            // user can tell version-skew apart from missing-install.
            string installedVersion = CodexProcessManager.GetInstalledMcpRemoteVersion();
            if (!string.Equals(installedVersion, CodexProcessManager.McpRemoteVersion, StringComparison.Ordinal))
            {
                failureReason = "mcp-remote version mismatch (installed: "
                    + (installedVersion ?? "unknown")
                    + ", expected: " + CodexProcessManager.McpRemoteVersion
                    + "). Run: npm install -g mcp-remote@" + CodexProcessManager.McpRemoteVersion;
                return null;
            }

            lock (_writeLock)
            {
                try
                {
                    string managedBlock = BuildManagedTomlBlock(mcpUrl, bearerToken, mcpRemotePath);
                    string configPath = GetCodexConfigPath();
                    string configDir = Path.GetDirectoryName(configPath);
                    if (!string.IsNullOrEmpty(configDir))
                        Directory.CreateDirectory(configDir);

                    bool fileExisted = File.Exists(configPath);
                    string existing = fileExisted ? File.ReadAllText(configPath) : string.Empty;
                    string updated = ReplaceOrAppendManagedBlock(existing, managedBlock);

                    if (string.Equals(existing, updated, StringComparison.Ordinal))
                        return configPath;

                    if (fileExisted)
                    {
                        string bak = configPath + ".clarionassistant.bak";
                        if (!File.Exists(bak))
                        {
                            try { File.Copy(configPath, bak, overwrite: false); }
                            catch { }
                        }
                    }

                    // Atomic write. File.Replace uses ReplaceFileW on NTFS — single
                    // syscall, concurrent readers see either old or new, never a
                    // half-written file. Falls back to first-write Move when the
                    // destination doesn't exist yet (Replace requires it to exist).
                    //
                    // Per-invocation unique tmp name: cross-process safety. Two CA
                    // processes (e.g. two IDE instances) share ~/.codex/config.toml
                    // but NOT _writeLock. A fixed ".ca-new" name would let process A
                    // overwrite process B's staged content. PID + GUID in the
                    // filename eliminates the collision entirely.
                    string tmpPath = configPath + ".ca-new-"
                        + System.Diagnostics.Process.GetCurrentProcess().Id + "-"
                        + Guid.NewGuid().ToString("N");
                    File.WriteAllText(tmpPath, updated, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
                    try
                    {
                        if (fileExisted)
                        {
                            File.Replace(tmpPath, configPath, destinationBackupFileName: null, ignoreMetadataErrors: true);
                        }
                        else
                        {
                            // First-write cross-process race: another process may have
                            // created the file after our File.Exists check but before
                            // we Move. If Move fails because dest now exists, retry as
                            // Replace so the latest writer wins atomically rather than
                            // dropping the registration with an exception.
                            try
                            {
                                File.Move(tmpPath, configPath);
                            }
                            catch (IOException) when (File.Exists(configPath))
                            {
                                File.Replace(tmpPath, configPath, destinationBackupFileName: null, ignoreMetadataErrors: true);
                            }
                        }
                    }
                    catch
                    {
                        try { if (File.Exists(tmpPath)) File.Delete(tmpPath); } catch { }
                        throw;
                    }

                    return configPath;
                }
                catch (Exception ex)
                {
                    failureReason = "Writing ~/.codex/config.toml failed: " + ex.Message;
                    System.Diagnostics.Debug.WriteLine("[CodexConfigService] EnsureMcpRegistration failed: " + ex);
                    return null;
                }
            }
        }

        private static string BuildManagedTomlBlock(string mcpUrl, string bearerToken, string mcpRemotePath)
        {
            // Resolved local path for mcp-remote.cmd — no npx, no registry fetch.
            // The user installs it once with `npm install -g mcp-remote@<version>`
            // and CA then invokes it directly. This eliminates the supply-chain
            // surface that runtime-`npx` would expose to the bearer token.
            var sb = new StringBuilder();
            sb.AppendLine(ManagedMarkerBegin);
            sb.AppendLine("# Generated by ClarionAssistant — points Codex CLI at the IDE's local MCP server.");
            sb.AppendLine("# Content between the markers is replaced on every Codex terminal launch.");
            sb.AppendLine("# CA rewrites the bearer token each launch, so do not hand-edit it here.");
            sb.AppendLine("# To change the bridge: npm install -g mcp-remote@" + CodexProcessManager.McpRemoteVersion);
            sb.AppendLine();
            sb.AppendLine("[mcp_servers." + ServerName + "]");
            sb.AppendLine("command = " + TomlQuote(mcpRemotePath));

            var args = new List<string> { mcpUrl };
            if (!string.IsNullOrEmpty(bearerToken))
            {
                args.Add("--header");
                args.Add("Authorization:Bearer " + bearerToken);
            }
            sb.Append("args = [");
            for (int i = 0; i < args.Count; i++)
            {
                if (i > 0) sb.Append(", ");
                sb.Append(TomlQuote(args[i]));
            }
            sb.AppendLine("]");
            sb.AppendLine("startup_timeout_sec = 30");
            sb.AppendLine();
            sb.Append(ManagedMarkerEnd);
            return sb.ToString();
        }

        /// <summary>
        /// Rewrite the managed block in a self-healing way:
        /// <list type="bullet">
        ///   <item>Drop every CA begin/end marker line (paired or orphan).</item>
        ///   <item>Drop every <c>[mcp_servers.clarion-assistant.*]</c> section
        ///         anywhere in the file. The fresh block owns this content;
        ///         leftover copies would crash TOML parsing with a duplicate-key
        ///         error.</item>
        ///   <item>Lift foreign top-level sections out of any region that was
        ///         between CA markers (e.g. <c>[tui.model_availability_nux]</c>
        ///         that Codex CLI wrote inside our markers after stripping our
        ///         trailing end-marker comment) so they survive the rewrite.</item>
        ///   <item>Append one fresh canonical managed block at the end.</item>
        /// </list>
        ///
        /// NOTE: This is a line scanner, not a full TOML parser. It tracks
        /// multi-line basic / literal string state (<c>"""..."""</c> /
        /// <c>'''...'''</c>) so a bracketed token at column 0 inside a multi-line
        /// string isn't misclassified as a section header.
        /// </summary>
        private static string ReplaceOrAppendManagedBlock(string existing, string managedBlock)
        {
            if (existing == null) existing = string.Empty;
            var lines = existing.Split(new[] { "\r\n", "\n", "\r" }, StringSplitOptions.None);

            var keep = new StringBuilder();        // content outside any CA managed region (minus managed sections)
            var lifted = new StringBuilder();      // foreign sections that were inside CA markers
            bool inManaged = false;
            string currentSection = null;
            bool inBasicMultiline = false;
            bool inLiteralMultiline = false;

            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i];

                // CA markers are pure single-line comments and can't appear
                // inside a multi-line string, so we check them before updating
                // multi-line state. Whole-line match (with surrounding whitespace
                // tolerated) avoids false positives.
                string trimAll = line.Trim();
                if (trimAll.Equals(ManagedMarkerBegin, StringComparison.Ordinal))
                {
                    inManaged = true;
                    currentSection = null;
                    continue; // drop marker line
                }
                if (trimAll.Equals(ManagedMarkerEnd, StringComparison.Ordinal))
                {
                    inManaged = false;
                    currentSection = null;
                    continue; // drop marker line
                }

                int basicCount = CountOccurrences(line, "\"\"\"");
                int literalCount = CountOccurrences(line, "'''");
                bool wasInMultiline = inBasicMultiline || inLiteralMultiline;
                if ((basicCount % 2) != 0) inBasicMultiline = !inBasicMultiline;
                if ((literalCount % 2) != 0) inLiteralMultiline = !inLiteralMultiline;
                bool nowInMultiline = inBasicMultiline || inLiteralMultiline;

                // Update currentSection only when not inside a multi-line string.
                if (!wasInMultiline && !nowInMultiline)
                {
                    string ltrim = line.TrimStart();
                    if (ltrim.StartsWith("[", StringComparison.Ordinal))
                    {
                        int closeBracket = ltrim.IndexOf(']');
                        if (closeBracket > 0)
                        {
                            string name = ltrim.Substring(1, closeBracket - 1);
                            if (name.StartsWith("[", StringComparison.Ordinal)) name = name.Substring(1);
                            currentSection = name.Trim();
                        }
                    }
                }

                bool isManagedSection = currentSection != null
                    && (currentSection.Equals(ManagedTablePrefix, StringComparison.Ordinal)
                        || currentSection.StartsWith(ManagedTablePrefix + ".", StringComparison.Ordinal));

                if (isManagedSection)
                {
                    // Drop — fresh block owns this. Duplicate would crash TOML.
                    continue;
                }

                if (inManaged)
                {
                    lifted.Append(line).Append(Environment.NewLine);
                }
                else
                {
                    keep.Append(line).Append(Environment.NewLine);
                }
            }

            string keepStr = TrimTrailingBlankLines(keep.ToString());
            string liftedStr = TrimTrailingBlankLines(lifted.ToString());

            var sb = new StringBuilder();
            if (keepStr.Length > 0)
            {
                sb.Append(keepStr);
                sb.Append(Environment.NewLine);
                sb.Append(Environment.NewLine);
            }
            sb.Append(managedBlock);
            sb.Append(Environment.NewLine);
            if (liftedStr.Length > 0)
            {
                sb.Append(Environment.NewLine);
                sb.Append(liftedStr);
                sb.Append(Environment.NewLine);
            }
            return sb.ToString();
        }

        private static string TrimTrailingBlankLines(string s)
        {
            if (string.IsNullOrEmpty(s)) return string.Empty;
            int end = s.Length;
            while (end > 0)
            {
                char c = s[end - 1];
                if (c == ' ' || c == '\t' || c == '\r' || c == '\n') { end--; continue; }
                break;
            }
            return s.Substring(0, end);
        }

        private static int CountOccurrences(string s, string needle)
        {
            if (string.IsNullOrEmpty(s) || string.IsNullOrEmpty(needle)) return 0;
            int n = 0;
            int idx = 0;
            while ((idx = s.IndexOf(needle, idx, StringComparison.Ordinal)) >= 0)
            {
                n++;
                idx += needle.Length;
            }
            return n;
        }

        private static string TomlQuote(string value)
        {
            if (value == null) value = string.Empty;
            var sb = new StringBuilder(value.Length + 2);
            sb.Append('"');
            foreach (char c in value)
            {
                switch (c)
                {
                    case '\\': sb.Append("\\\\"); break;
                    case '"':  sb.Append("\\\""); break;
                    case '\b': sb.Append("\\b"); break;
                    case '\f': sb.Append("\\f"); break;
                    case '\n': sb.Append("\\n"); break;
                    case '\r': sb.Append("\\r"); break;
                    case '\t': sb.Append("\\t"); break;
                    default:
                        if (c < 0x20) sb.Append("\\u").Append(((int)c).ToString("X4"));
                        else sb.Append(c);
                        break;
                }
            }
            sb.Append('"');
            return sb.ToString();
        }
    }
}
