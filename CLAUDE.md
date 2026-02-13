# CLAUDE.md — Dossier (formerly Dossier)

## Project Overview

Dossier is a cross-platform desktop app that automates repetitive admissions workflows in UCL's Portico student record system. Staff load a spreadsheet of student records and the app batch-processes them via browser automation — either processing Accept/Reject decisions or generating merged overview PDFs.

**Rename in progress:** The project is being renamed from "Dossier" to "Dossier". The GitHub repo rename should happen first, then namespaces/files/references in the codebase will be updated.

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
├── MainWindow.axaml / .cs       # Main UI — file loading, mode selection, start processing
├── SettingsWindow.axaml / .cs   # Settings dialog (headless mode, delays, URL)
├── Views/
│   └── ProcessingWindow.axaml / .cs  # Live processing UI — progress bar, student list, log, timer
├── Models/
│   ├── StudentRecord.cs         # Student data model (StudentNo, Decision, Programme, etc.)
│   ├── AppConfig.cs             # Configuration model
│   └── ProcessingStudentViewModel.cs  # ViewModel for processing window student rows
├── Services/
│   ├── ExcelService.cs          # Excel (.xlsx/.xls) and CSV parsing with fuzzy column detection
│   ├── IExcelService.cs         # Interface
│   ├── PorticoAutomationService.cs   # Core Playwright automation — all browser interactions
│   ├── IPorticoAutomationService.cs  # Interface
│   └── SleepInhibitor.cs        # Cross-platform OS sleep prevention (caffeinate/SetThreadExecutionState/systemd-inhibit)
├── Dossier.csproj
└── Dossier.sln
```

## Key Concepts

### Processing Modes
1. **Process Accepts / Process Rejects** — Reads Decision column from spreadsheet, navigates to each student in Portico, selects the correct programme, goes to Actions tab, clicks "Recommend Offer or Reject", selects the decision, and clicks Process.
2. **Merge Overview** — For each student, navigates to their record, opens Documents & Uploads tab, clicks "Create Overview.pdf" or "Amend Overview.pdf", waits for server processing, clicks "Merge Documents", downloads the merged PDF to `~/Desktop/LATEST_BATCH/`.

### Programme Code Mapping
Short codes in the spreadsheet (ML, CS, FT, DSML, etc.) are mapped to Portico programme codes (TMSCOMSMCL01, etc.) inside `PorticoAutomationService.cs`. This mapping is used to click the correct programme row when a student has multiple applications.

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

## Common Pitfalls

- **Amend vs Create Overview:** The Portico system shows either "Create Overview.pdf" or "Amend Overview.pdf" depending on whether an overview already exists. The code handles both via JavaScript text matching across all element types (input, button, a).
- **Slow server responses:** The Portico server can take 5–40 seconds to process overview generation. Never use `WaitForLoadStateAsync` or `RunAndWaitForNavigationAsync` for these — they either resolve too early or hang. Use the JS click + DOM polling pattern instead.
- **Navigation state after failures:** If a student record fails, the browser is left on an unknown page. Always call `NavigateToUclSelectAsync()` in catch blocks to recover.

## Style Notes

- UI uses glassmorphic design: acrylic blur, soft shadows, rounded corners, color-coded status indicators
- ProcessingWindow footer shows live progress counter with elapsed timer
- Keep the README professional — the project is shown to hiring managers
