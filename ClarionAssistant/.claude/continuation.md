# Continuation: ClarionAssistant

## NEXT SESSION — TEST THIS

"all deployed. load the classify app. load the txa file classifyit.txa into the ide editor. there was another file CreateTheClass_optimized.clw where you made some changes in the code. we were trying to see if you could go to the embeditor for the CreateTheClass procedure and make the changes you made in the .clw in the correct embeds in the embeditor"

### What to do:
1. Open ClassifyIt.app with `open_app` (`H:\Dev\C11Apps\ClassifyIt\ClassifyIt.app`)
2. Open `H:\Dev\C11Apps\ClassifyIt\ClassifyIt.txa` in the editor with `open_file`
3. Read `H:\Dev\C11Apps\ClassifyIt\CreateTheClass_optimized.clw` into context with `read_file`
4. Open embeditor for CreateTheClass with `open_procedure_embed`
5. Test `get_embed_info` — should now work (fixed this session)
6. Use `get_lines_range` (NEW tool) to read embed code in bulk instead of line-by-line
7. Use `find_in_file` to locate key landmarks (stQueuesLine, PathToSaveClass, etc.)
8. Apply optimizations from the .clw file using `replace_range` / `delete_range`, working bottom-to-top

### The 12 edits to apply (bottom-to-top order):
1. **Lines ~12022-12024**: Use `BasePath` in final SaveFile calls
2. **Lines ~12006-12015**: Consolidate 8 method stub Append calls into 1
3. **Lines ~11996-11999**: Replace no-op ELSE copy loop with `stFinalTemplate.SetValue(stTemplate.GetValue())`
4. **Line ~11977**: Delete dead `stQueuesLine.SetValue` in Destruct section
5. **Line ~11948**: Delete dead `stQueuesLine.SetValue` in Construct section
6. **Lines ~11853-11855**: Use `SavedCLIGUID` instead of `qDerived.CLIGUID`
7. **Lines ~11784-11846**: BIG CHANGE — Replace line-by-line .inc reparse with direct `stTemplate.Replace()` calls. Save from `stTemplate` (not `stFinalTemplate`). Use `BasePath`.
8. **Line ~11676**: Add `SavedCLIGUID = LocalCLIGUID` before `ADD(qDerived)`
9. **Line ~11628**: Use `BasePath` in EXISTS check
10. **After line ~11624**: Add `BasePath = CLIP(PathToSaveClass) & '\' & CLIP(NewClassName)` after CODE
11. **After line ~11622**: Add `BasePath STRING(260)` and `SavedCLIGUID STRING(16)` variable declarations
12. **Line ~11594**: Delete dead `stQueuesLine` variable declaration

**IMPORTANT**: Line numbers are approximate — use `find_in_file` to find exact positions before editing. The embed code starts around line 11573 (CreateTheClass PROCEDURE) in the embeditor.

### Speed optimization strategy:
- Use `get_lines_range(start, end)` to read blocks of 100-200 lines at once
- Use `find_in_file` to pinpoint exact line numbers for edits
- Do NOT use `get_line_text` in a loop — that's what made the previous attempt slow

## Current Status

### New MCP tools — BUILT 2026-03-25, AWAITING DEPLOY + TEST

#### `get_lines_range` tool
- Added to EditorService.cs: `GetLinesRange(int startLine, int endLine)` method
- Gets full text content ONCE, then extracts all line segments in a loop
- Returns `lineNumber\tcontent` per line
- Registered in McpToolRegistry.cs with `start_line` and `end_line` params
- Should turn 800 get_line_text calls into 4-5 get_lines_range calls

#### Improved embeditor detection
- `GetClaGenEditor()` now searches ALL workbench windows (not just ActiveWorkbenchWindow)
- Checks both primary ViewContent AND SecondaryViewContents for ClaGenEditor
- Same pattern as the `FindAppViewContent()` fix from earlier
- `GetEmbedInfo()` now uses `IGeneratorDialog` interface check instead of unreliable `AppName` check
- Also returns `editorType` field in the response

### WebView2 Header — DEPLOYED, AWAITING TEST
Replaced the WinForms solution bar (Panel) and toolbar (ToolStrip) with a single WebView2 header.
- `Terminal/HeaderWebView.cs` — C# wrapper with postMessage bridge
- `Terminal/header.html` — Dark themed (Catppuccin Mocha) HTML/CSS header
- Title: "Clarion Assistant" with MCP status indicator
- Version dropdown + refresh, Solution dropdown + browse + Full Index / Update
- Action buttons: New Chat, Settings, Create COM
- All button clicks route through `OnHeaderAction()` in `ClaudeChatControl.cs`
- `HeaderReady` event triggers `LoadVersions()` and `LoadSolutionHistory()` to populate dropdowns

**What to test after deploy:**
1. Header renders with dark theme, title shows "Clarion Assistant"
2. Version dropdown populates from ClarionVersionService
3. Solution dropdown populates from history, auto-detects from IDE
4. Refresh button re-detects version + solution from IDE
5. Browse (...) opens file dialog, updates solution
6. Full Index / Update buttons trigger indexing, status updates
7. New Chat, Settings, Create COM buttons all work
8. MCP status shows in top-right corner ("MCP: port XXXXX | N tools")
9. If header height needs adjusting, change `Height = 105` in `HeaderWebView.cs`

### Library CodeGraph — TESTED, WORKING
Deployed and verified. 1,082 symbols indexed from Clarion LibSrc equate files. Confirmed EVENT:, PROP:, COLOR: lookups all work via `query_codegraph` with `db_path` pointing to `C:\Clarion12\Accessory\AddIns\ClarionAssistant\ClarionLib.codegraph.db`. Shows up in `list_codegraph_databases`.

### select_procedure — TESTED, WORKING
- `select_procedure` MCP tool selects a procedure in the ClaList by name
- Uses `PostMessage` + `WM_CHAR` only
- VK_DOWN + VK_UP after typing clears the locator buffer
- Tested 2026-03-22: successfully selected "ScanClass" in the app tree
- Now includes embeditor-open guard (returns error if embeditor is active)

### open_procedure_embed — TESTED, WORKING
- Timing fix deployed and verified 2026-03-22
- 100ms per-char delay + `DoEvents()` after each character
- Tested: ScanClass opened 3 times, FillCheckBoxes opened 3 times — all correct
- Now includes embeditor-open guard (returns error if embeditor is active, tells user to close it first)
- **2026-03-25 FIX**: Fixed `FindAppViewContent()` bug — now searches `WorkbenchWindowCollection` for any open `.app` ViewContent

### save_and_close_embeditor — TESTED, WORKING
### cancel_embeditor — TESTED, WORKING
### Embeditor lifecycle — COMPLETE
### Embed navigation tools — TESTED, WORKING
### "Create COM" toolbar button — TESTED, WORKING

**Design direction (discussed 2026-03-22):**
- ClarionAssistant is evolving into a Clarion-focused skill dispatcher with IDE-native UI buttons
- COM controls and addins live in centralized folders (shared across solutions)
- CodeGraph analysis stays per-solution
- Future buttons: "Create Addin", "Analyze Solution", etc.

### ClassifyIt TXA Code Review — IN PROGRESS
- Exported ClassifyIt.app to `H:\Dev\C11Apps\ClassifyIt\ClassifyIt.txa` (128K lines, 4.75 MB)
- Analyzed all embed code in the class scanning procedures (ScanClass, ScanAllClasses, ScanOneClass, CreateClass, CalculateMethodRow, ProcessDerivedClass, ClassPeek, PropertiesAndMethods, ExportPropertiesMethods)
- Wrote optimized `CreateTheClass()` to `H:\Dev\C11Apps\ClassifyIt\CreateTheClass_optimized.clw`
- **Optimizations applied**: removed dead variable, computed BasePath once, saved CLIGUID before FREE, replaced line-by-line reparse with Replace() calls, eliminated no-op copy loop, unified Construct/Destruct injection into shared ROUTINE, consolidated method stub Append calls
- **First attempt (2026-03-25)**: Successfully applied all 12 edits via embeditor, but it took too long (~10 min) because reading 800 lines required 800 individual `get_line_text` calls. User cancelled the edits so we can re-test with `get_lines_range`.
- **Also noted but not yet addressed**: Duplicated `AddClassNames` parser in both ScanClass and ScanOneClass — maintenance risk, should consolidate

## Build notes
- ONLY build with MSBuild: `MSYS_NO_PATHCONV=1 "C:/Program Files/Microsoft Visual Studio/2022/Community/MSBuild/Current/Bin/MSBuild.exe" ClarionAssistant.csproj /p:Configuration=Debug /v:minimal`
- Need `MSYS_NO_PATHCONV=1` prefix in bash
- Do NOT use `dotnet build` (WebView2 resolution fails) or `deploy.ps1` (DLLs locked)
- User deploys manually (copies DLL to `C:\Clarion12\accessory\addins\ClarionAssistant`)
- Pre-existing warning in LspClient.cs (CS0414 `_initialized`) — ignore

## Architecture
- **Header**: `Terminal/HeaderWebView.cs` + `Terminal/header.html` — WebView2 with postMessage bridge
- **Terminal**: `Terminal/WebViewTerminalRenderer.cs` + `Terminal/terminal.html` — xterm.js in WebView2
- **Both share** `Terminal/WebView2EnvironmentCache.cs` for the CoreWebView2Environment
- **Layout**: HeaderWebView (Dock=Top, 105px) → WebViewTerminalRenderer (Dock=Fill)

## Known issues / TODO
- **IntPtr overflow pattern**: Any hex constant with bit 31 set (>=0x80000000) will overflow `IntPtr` cast on 32-bit. Always use `new IntPtr(unchecked((int)0x...))`.
- **Library CodeGraph**: Consider adding builtins.clw function/procedure declarations (not just EQUATE lines) in a future iteration.
- **ClaList is confirmed NOT a standard listbox** — LB_GETCOUNT, LB_FINDSTRINGEXACT, LB_GETTEXTLEN all return 0. May be a treeview or fully custom control. Keystroke approach is the only viable method for now.

## Locator clearing approach (important pattern)
- **VK_ESCAPE**: DO NOT USE — triggers IDE exit dialog
- **VK_DOWN + VK_UP**: Clears the locator buffer. Send after typing + settling. The selection moves down one item then back up, ending on the same item but with a clean locator.
- This must happen AFTER the WM_CHAR typing and 500ms settle, BEFORE the Embeditor button click (in OpenProcedureEmbed) or thread detach (in SelectProcedure).

## Keystroke timing (current values as of 2026-03-22)
- **100ms** between WM_CHAR posts + `DoEvents()` after each char
- **500ms** settle after all chars typed
- **100ms + DoEvents** between VK_DOWN and VK_UP for locator clear
- **300ms** final settle after locator clear
- If still too fast, try 150-200ms per char next
