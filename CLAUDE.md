# CLAUDE.md — Dossier

## Project Overview

Dossier is a cross-platform desktop app that automates repetitive admissions workflows in UCL's Portico student record system. Staff load a spreadsheet of student records and the app batch-processes them via browser automation — either processing Accept/Reject decisions or generating merged overview PDFs. It also includes a PDF File Tools module (ported from PDFusion) for renaming and ranking the downloaded PDFs.

## Tech Stack

- **Language:** C# / .NET 10
- **UI Framework:** Avalonia UI 11 (cross-platform — Windows, macOS, Linux)
- **Browser Automation:** Microsoft Playwright 1.49 (targets Microsoft Edge)
- **Excel Parsing:** EPPlus 7.5.2
- **Architecture:** MVVM (Model-View-ViewModel)

## Build & Run

```bash
dotnet restore
dotnet build
dotnet run
```

Playwright browsers must be installed once: `playwright install msedge`

## Project Structure

```
├── MainWindow.axaml / .cs            # Main UI — file loading, mode selection, start processing
├── SettingsWindow.axaml / .cs        # Settings dialog (headless mode, delays, URL)
├── Views/
│   ├── ProcessingWindow.axaml / .cs  # Live processing UI — progress bar, student list, log, timer
│   └── PdfToolsWindow.axaml / .cs    # PDF File Tools — standalone rename/rank window
├── Models/
│   ├── StudentRecord.cs              # Student data model (StudentNo, Decision, Programme,
│   │                                 #   Batch, FeeStatus, UKGrade, ApplicationQualityRank, etc.)
│   ├── AppConfig.cs                  # Configuration model
│   ├── ProcessingStudentViewModel.cs # ViewModel for processing window student rows
│   └── PdfRenamePreviewItem.cs       # ViewModel for PDF rename preview DataGrid rows
├── Services/
│   ├── ExcelService.cs               # Excel (.xlsx/.xls) and CSV parsing with fuzzy column detection
│   ├── IExcelService.cs
│   ├── PorticoAutomationService.cs   # Core Playwright automation — all browser interactions
│   ├── IPorticoAutomationService.cs
│   ├── PdfRenameService.cs           # PDF rename/rank logic (ported from PDFusion)
│   ├── IPdfRenameService.cs
│   └── SleepInhibitor.cs             # Cross-platform OS sleep prevention
├── Dossier.Tests/                    # xUnit tests (ExcelService CSV + Excel, models)
├── Dossier.csproj
└── Dossier.sln
```

## Key Concepts

### Processing Modes
1. **Process Accepts / Process Rejects** — Reads Decision column from spreadsheet, navigates to each student in Portico, selects the correct programme, goes to Actions tab, clicks "Recommend Offer or Reject", selects the decision, and clicks Process.
2. **Merge Overview** — Before Playwright starts, prompts for a batch number. Then for each student, navigates to their record, opens Documents & Uploads tab, clicks "Create Overview.pdf" or "Amend Overview.pdf", waits for server processing, clicks "Merge Documents", and downloads the merged PDF. Downloads land in `~/Desktop/{PROGRAMME}_LATEST/Batch N - MMM D - MMM D/` and are auto-renamed immediately on download.

### Folder & File Naming (Merge Overview)

**Programme folder** (`GetBatchFolderName()` in `MainWindow.axaml.cs`):
- Reads the `Programme` column from loaded students and produces `CSML_LATEST`, `DSML_LATEST`, etc.
- If students have >3 distinct programmes, falls back to `LATEST_BATCH`.

**Batch subfolder** (`ComputeBatchFolderName(int n)` in `MainWindow.axaml.cs`):
- Format: `Batch N - MMM D - MMM D` (e.g. `Batch 7 - Apr 10 - Apr 27`)
- Date range is min/max `ReceivedDate` from loaded students. Omitted if no dates present.
- Batch number is entered by the user in a dialog that appears before Playwright starts.

**Auto-rename** (`AutoRenameStudentPdf()` in `MainWindow.axaml.cs`):
- Immediately after each successful download, finds `{studentNo}-*.PDF` in the batch folder and `File.Move`s it to the formatted name: `b7 John Smith 26049530 H 2_1.pdf`.
- Requires `Batch`, `FeeStatus`, `UKGrade` columns for a meaningful name; gracefully degrades without them.

**Resulting structure on disk:**
```
~/Desktop/DSML_LATEST/
  Batch 7 - Apr 10 - Apr 27/
    b7 John Smith 26049530 H 2_1.pdf
    b7 Jane Doe 22082064 OS 2_1.pdf
    ...
```

### PDF File Tools (`PdfToolsWindow`)

A standalone window (also reachable via "📄 Rename & Rank PDFs" button in MainWindow's action panel) for post-processing downloaded PDFs without running Playwright.

**Features:**
- **Load Spreadsheet** — independent spreadsheet loader (CSV or Excel) within the window; pre-populated when opened from MainWindow with students already loaded.
- **Preview** — scans selected folder, matches PDFs to students by student number, shows `CurrentFilename → NewFilename` in a DataGrid with colour-coded status.
- **Rename All** — prompts for batch number, computes `Batch N - MMM D - MMM D` folder name, then `File.Move`s all matched files into that subfolder. Undo is supported.
- **Append Ranking** — explicit manual button; only enabled when the spreadsheet has a ranking column (`AT Note (Ranking)`, `ApplicationQualityRank`, `Ranking`, etc.). Copies files to a `RankRenamed/` subfolder prefixed with the ranking letter: `A - b7 John Smith...pdf`. Originals unchanged.
- **Open Folder** — opens the current folder in Finder/Explorer.

**Filename format** (`PdfRenameService.GenerateNewFilename()`):
```
b{batch} {Forename} {Surname} {StudentNo} {FeeCode} {Grade}.pdf
```
- `b{batch}` — digits extracted from the `Batch` column (e.g. "Batch 7" → "b7")
- `FeeCode` — `H` (Home), `OS` (Overseas/European), `?` (unknown)
- `Grade` — `1`, `2_1`, `2_2`, `3`, `Masters`, `XX` (unknown)

**Student number extraction** (`ExtractStudentNumber()` in `PdfRenameService.cs`):
- Handles two formats: original `26049530-01-01-OVERVIEW.PDF` (digits before first dash) and already-renamed `b7 John Smith 26049530 H 2_1.pdf` (longest 7–10 digit space-delimited token).

### Programme Code Mapping
Short codes in the spreadsheet (ML, CS, FT, DSML, CSML, etc.) are mapped to Portico programme codes (TMSCOMSMCL01, etc.) inside `PorticoAutomationService.cs`. This mapping is used to click the correct programme row when a student has multiple applications.

### Browser Automation Patterns
- **Persistent browser context** is stored at `~/.dossier-auth` so SSO/MFA sessions survive between runs.
- **JavaScript `el.click()`** is used instead of Playwright's `.ClickAsync()` for buttons that trigger server-side page reloads — this avoids Playwright's navigation tracking hanging on slow server responses.
- **DOM polling** (every 2 seconds, up to 120s timeout) checks `document.readyState === "complete"` AND whether the trigger button has disappeared, to detect when a slow server-side operation has finished.
- **Failure recovery:** Each student is processed in a try/catch. On failure, the catch block calls `NavigateToUclSelectAsync()` to return the browser to a known-good state before continuing to the next student.

### Sleep Prevention
`SleepInhibitor.cs` prevents the OS from sleeping during batch runs:
- macOS: spawns `caffeinate -di`
- Windows: `SetThreadExecutionState` P/Invoke
- Linux: `systemd-inhibit`
Activated at the start of processing, released automatically via `IDisposable` when processing ends.

### Input Format
- Accepts `.xlsx`, `.xls`, and `.csv` files
- Column detection is fuzzy — two-pass matching (exact then contains) for headers like StudentNo, Decision, Programme, Forename, Surname
- CSV auto-detects delimiter (comma or tab) and handles quoted fields (RFC 4180)
- Decision column is optional (only needed for Accept/Reject mode)
- PDF Tools columns (`Batch`, `FeeStatus`, `UKGrade`, `ApplicationQualityRank`) are optional; detected automatically if present

## Common Pitfalls

- **Amend vs Create Overview:** The Portico system shows either "Create Overview.pdf" or "Amend Overview.pdf" depending on whether an overview already exists. The code handles both via JavaScript text matching across all element types (input, button, a).
- **Slow server responses:** The Portico server can take 5–40 seconds to process overview generation. Never use `WaitForLoadStateAsync` or `RunAndWaitForNavigationAsync` for these — they either resolve too early or hang. Use the JS click + DOM polling pattern instead.
- **Navigation state after failures:** If a student record fails, the browser is left on an unknown page. Always call `NavigateToUclSelectAsync()` in catch blocks to recover.
- **Auto-rename requires PDF Tools columns:** `AutoRenameStudentPdf()` calls `GenerateNewFilename()` which needs `Batch`, `FeeStatus`, `UKGrade`. If these columns are absent, the rename still runs but produces `b0 ... ? XX.pdf`. This is intentional — the file is still findable by student number for the skip check.
- **Batch folder is pre-created before Playwright:** `batchDownloadPath` is set and the directory created before `InitialiseAsync` is called. If the user cancels the batch number dialog, the entire run is aborted cleanly before any browser opens.

## Style Notes

- UI uses glassmorphic design: acrylic blur, soft shadows, rounded corners, colour-coded status indicators
- ProcessingWindow footer shows live progress counter with elapsed timer
- Keep the README professional — the project is shown to hiring managers
