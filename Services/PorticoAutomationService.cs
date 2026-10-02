// Services/PorticoAutomationService.cs

using System.Diagnostics;
using Microsoft.Playwright;
using Dossier.Models;

namespace Dossier.Services;

public class PorticoAutomationService : IPorticoAutomationService
{
    // Single browser shared by every tab — the Edge profile only supports one
    // running browser at a time, so all automation must go through one instance.
    public static PorticoAutomationService Shared { get; } = new();

    private const string HandoffPidFileName = "dossier-handoff.pid";

    private IPlaywright? _playwright;
    private IBrowser? _browser;
    private IBrowserContext? _context;
    private IPage? _page;
    private AppConfig? _config;
    private string? _userDataDir;

    public bool DebugMode { get; set; } = false;

    // When set (e.g. "2026/27"), SearchForStudentAsync selects that academic year in the
    // Portico "Year" dropdown before searching. Null = default "Current Applications".
    public string? TargetAcademicYear { get; set; } = null;

    private readonly Dictionary<string, string> _shortToLongProgCodes = new(StringComparer.OrdinalIgnoreCase)
    {
        { "AIBH", "TMSARTSINT03" },
        { "AISD", "TMSARTSINT02" },
        { "AIDE", "TMSCOMSSAD18" },
        { "ISEC", "TMSCOMSINF01" },
        { "CF",   "TMSCOMSCFI01" },
        { "FRM",  "TMSCOMSFRM01" },
        { "FT",   "TMSFINSTEC01" },
        { "EDT",  "TMSCOMSEDT01" },
        { "ML",   "TMSCOMSMCL01" },
        { "DSML", "TMSDATSMLE01" },
        { "CSML", "TMSCOMSSML01" },
        { "RAI",  "TMSROBAARI01" },
        { "SEIOT","TMSCOMSEIT01" },
        { "DDI",  "TMSCOMSDDI19" },
        { "CS",   "TMSCOMSING01" },
        { "SSE",  "TMSCOMSSSE01" },
        { "CGVI", "TMSCOMSCGV01" }
    };

    public event EventHandler<string>? StatusUpdated;
    public event EventHandler<StudentRecord>? StudentProcessed;

    public bool IsInitialised => _page != null;

    public async Task InitialiseAsync(AppConfig config)
    {
        _config = config;

        // Reuse the browser left open from a previous run, if it's still alive.
        if (_context != null)
        {
            try
            {
                _page = _context.Pages.FirstOrDefault(p => !p.IsClosed) ?? await _context.NewPageAsync();
                LogStatus("Reusing existing browser session (launch settings unchanged).");
                return;
            }
            catch
            {
                LogStatus("Previous browser was closed — launching a fresh one.");
                try { await _context.CloseAsync(); } catch { }
                _context = null;
                _playwright?.Dispose();
                _playwright = null;
                _page = null;
            }
        }

        LogStatus("Initialising Playwright...");
        try
        {
            _playwright = await Playwright.CreateAsync();
        }
        catch (Exception ex)
        {
            LogStatus($"ERROR — Playwright driver failed to start: {ex.GetType().Name}: {ex.Message}");
            LogStatus("Check that Microsoft Edge is installed, or run: playwright install msedge");
            throw;
        }

        var userDataDir = config.EdgeUserDataDir;
        if (string.IsNullOrEmpty(userDataDir))
        {
            var appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            userDataDir = Path.Combine(appDataPath, "Dossier", "EdgeProfile");
            Directory.CreateDirectory(userDataDir);
        }
        _userDataDir = userDataDir;

        // If a hand-off browser from a previous session still owns the profile, close it first.
        await CloseStaleHandoffBrowserAsync(userDataDir);

        var contextOptions = new BrowserTypeLaunchPersistentContextOptions
        {
            Headless = config.HeadlessMode,
            Channel = "msedge",
            SlowMo = config.ActionDelayMs,
            AcceptDownloads = true,
            ViewportSize = null,
            Args = new[] { "--start-maximized" }
        };

        LogStatus("Launching Microsoft Edge...");
        try
        {
            _context = await _playwright.Chromium.LaunchPersistentContextAsync(userDataDir, contextOptions);
        }
        catch (Exception ex)
        {
            LogStatus($"ERROR — Edge failed to launch: {ex.GetType().Name}: {ex.Message}");
            LogStatus("Ensure Microsoft Edge is installed at /Applications/Microsoft Edge.app");
            throw;
        }

        _page = _context.Pages.FirstOrDefault() ?? await _context.NewPageAsync();
        LogStatus("Browser initialised.");
    }

    public async Task<bool> LoginAsync()
    {
        if (_page == null || _config == null) throw new InvalidOperationException("Service not initialised.");
        
        LogStatus($"Navigating to Portico: {_config.PorticoUrl}");
        await _page.GotoAsync(_config.PorticoUrl);
        
        try 
        {
            await _page.WaitForSelectorAsync("text=My Portico", new PageWaitForSelectorOptions { Timeout = 5000 });
            LogStatus("Session valid. Already logged in.");
            return true;
        } 
        catch 
        { 
            LogStatus("Session check: Login required."); 
        }

        var staffLoginButton = _page.GetByRole(AriaRole.Button, new() { Name = "Staff and Students Login" });
        if (await staffLoginButton.IsVisibleAsync()) 
            await staffLoginButton.ClickAsync();

        LogStatus("Waiting for manual SSO/MFA authentication...");
        
        try 
        {
            await _page.WaitForSelectorAsync("text=My Portico", new PageWaitForSelectorOptions { Timeout = 240000 });
            LogStatus("Successfully logged in to Portico.");
            return true;
        } 
        catch 
        { 
            return false; 
        }
    }

    public async Task<bool> NavigateToUclSelectAsync()
    {
        if (_page == null) throw new InvalidOperationException("Service not initialised.");
        
        LogStatus("Navigating to UCLSelect...");
        var uclSelectLink = _page.Locator("text=UCLSelect").First;
        if (await uclSelectLink.IsVisibleAsync()) 
        {
            await uclSelectLink.ClickAsync();
            await _page.WaitForLoadStateAsync(LoadState.NetworkIdle);
        }

        LogStatus("Clicking Search tab...");
        var searchTab = _page.Locator("a").Filter(new() { HasText = "Search" }).First;
        if (await searchTab.IsVisibleAsync()) 
        {
            await searchTab.ClickAsync();
            await _page.WaitForLoadStateAsync(LoadState.NetworkIdle);
        }
        
        LogStatus("Ready to search.");
        return true;
    }

    public async Task ProcessStudentAcceptAsync(StudentRecord student)
    {
        if (_page == null) throw new InvalidOperationException("Service not initialised.");
        
        student.Status = ProcessingStatus.Processing;
        StudentProcessed?.Invoke(this, student);
        
        try
        {
            LogStatus($"Processing OFFER for: {student.StudentNo} (Prog: '{student.Programme}')");
            await SearchForStudentAsync(student.StudentNo);
            await ClickStudentLinkAsync(student); 
            await NavigateToActionsTabAsync();
            await RecommendOfferAsync();
            
            student.Status = ProcessingStatus.Success;
            LogStatus($"SUCCESS: Offer processed for {student.StudentNo}");
        }
        catch (Exception ex)
        {
            student.Status = ProcessingStatus.Failed;
            student.ErrorMessage = ex.Message;
            LogStatus($"FAILED {student.StudentNo}: {ex.Message}");
        }
        
        StudentProcessed?.Invoke(this, student);
    }
    
    public async Task ProcessStudentRejectAsync(StudentRecord student)
    {
        if (_page == null) throw new InvalidOperationException("Service not initialised.");
        
        student.Status = ProcessingStatus.Processing;
        StudentProcessed?.Invoke(this, student);
        
        try
        {
            LogStatus($"Processing REJECT for: {student.StudentNo} (Prog: '{student.Programme}')");
            await SearchForStudentAsync(student.StudentNo);
            await ClickStudentLinkAsync(student); 
            await NavigateToActionsTabAsync();
            await RecommendRejectAsync();
            
            student.Status = ProcessingStatus.Success;
            LogStatus($"SUCCESS: Rejection processed for {student.StudentNo}");
        }
        catch (Exception ex)
        {
            student.Status = ProcessingStatus.Failed;
            student.ErrorMessage = ex.Message;
            LogStatus($"FAILED {student.StudentNo}: {ex.Message}");
        }
        
        StudentProcessed?.Invoke(this, student);
    }

    private async Task SearchForStudentAsync(string studentNo)
    {
        if (_page == null) return;
        
        LogStatus($"Searching: {studentNo}");

        // Past-applicant mode: switch the "Year" dropdown to the target academic year
        // before searching. The search tab defaults back to "Current Applications" on each
        // visit, so this must run for every student.
        if (!string.IsNullOrWhiteSpace(TargetAcademicYear))
            await SelectAcademicYearAsync(TargetAcademicYear);

        var radioLabel = _page.Locator("text=Student Number").First;
        if (await radioLabel.IsVisibleAsync()) 
            await radioLabel.ClickAsync();

        ILocator? searchInput = null;
        var textboxes = _page.GetByRole(AriaRole.Textbox);
        if (await textboxes.CountAsync() > 0) 
            searchInput = textboxes.First;
        
        if (searchInput == null || !await searchInput.IsVisibleAsync())
            searchInput = _page.Locator("input[type='text']").First;

        if (searchInput != null && await searchInput.IsVisibleAsync()) 
        {
            await searchInput.ClickAsync();
            await searchInput.ClearAsync();
            await searchInput.FillAsync(studentNo);
        } 
        else 
        {
            throw new Exception("Could not find search input field.");
        }
        
        var searchBtn = _page.Locator("input[value='Search']").First;
        if (!await searchBtn.IsVisibleAsync()) 
        {
            searchBtn = _page.Locator("button").Filter(new() { HasText = "Search" }).First;
        }
        
        await searchBtn.ClickAsync();
        await _page.WaitForLoadStateAsync(LoadState.NetworkIdle);
        await Task.Delay(500);
    }

    // Selects the academic year in the Portico search "Year" dropdown.
    // Matches the <option> whose text contains the target year (e.g. "2026/27").
    private async Task SelectAcademicYearAsync(string yearText)
    {
        if (_page == null) return;

        LogStatus($"Selecting academic year '{yearText}' in Year dropdown...");

        // Give the search form a moment to render its dropdowns.
        try
        {
            await _page.WaitForSelectorAsync("select", new PageWaitForSelectorOptions { Timeout = 10000 });
        }
        catch
        {
            throw new Exception("Year dropdown did not appear on the search page.");
        }

        var jsResult = await _page.EvaluateAsync<string>(@"
            (target) => {
                const selects = Array.from(document.querySelectorAll('select'));
                // Find the dropdown that actually contains an academic-year option.
                for (const select of selects) {
                    for (let i = 0; i < select.options.length; i++) {
                        if (select.options[i].text.includes(target)) {
                            select.selectedIndex = i;
                            select.value = select.options[i].value;
                            select.dispatchEvent(new Event('change', { bubbles: true }));
                            return 'SUCCESS: ' + select.options[i].text;
                        }
                    }
                }
                return 'FAILED: No option containing ""' + target + '"" found in any dropdown.';
            }
        ", yearText);

        LogStatus(jsResult);

        if (!jsResult.StartsWith("SUCCESS"))
            throw new Exception($"Could not select academic year '{yearText}'. {jsResult}");

        // Let any year-driven form refresh settle before continuing.
        await Task.Delay(500);
    }

    private async Task ClickStudentLinkAsync(StudentRecord student)
    {
        if (_page == null) return;

        LogStatus($"Looking for student {student.StudentNo} with Prog '{student.Programme}'");

        try 
        {
            await _page.WaitForSelectorAsync("table", new PageWaitForSelectorOptions { Timeout = 10000 });
            await Task.Delay(500);
        } 
        catch 
        {
            throw new Exception("Search results table did not appear.");
        }

        string inputProg = student.Programme?.Trim() ?? "";
        
        if (string.IsNullOrEmpty(inputProg))
        {
            throw new Exception("Programme column is empty.");
        }

        string searchCode = _shortToLongProgCodes.TryGetValue(inputProg, out var longCode) ? longCode : inputProg;
        
        if (searchCode != inputProg)
        {
            LogStatus($"Mapped '{inputProg}' -> '{searchCode}'");
        }

        var allLinks = await _page.Locator("table tbody tr td a").AllAsync();
        LogStatus($"Found {allLinks.Count} links in result rows");
        
        ILocator? targetLink = null;
        
        for (int i = 0; i < allLinks.Count; i++)
        {
            var link = allLinks[i];
            var parentRow = link.Locator("xpath=ancestor::tr[1]");
            var rowText = await parentRow.InnerTextAsync();
            
            bool hasStudentNo = rowText.Contains(student.StudentNo);
            bool hasProgCode = rowText.Contains(searchCode);
            
            var linkText = await link.InnerTextAsync();
            LogStatus($"Link {i}: '{linkText.Trim()}' | StudentNo={hasStudentNo} | ProgCode={hasProgCode}");
            
            if (hasStudentNo && hasProgCode)
            {
                var href = await link.GetAttributeAsync("href") ?? "";
                LogStatus($"Found matching link, href ends: ...{href.Substring(Math.Max(0, href.Length - 50))}");
                targetLink = link;
                break;
            }
        }
        
        if (targetLink == null)
        {
            throw new Exception($"Could not find link in row with StudentNo='{student.StudentNo}' AND ProgCode='{searchCode}'");
        }
        
        LogStatus("Clicking the matched link...");
        await targetLink.ClickAsync();
        await _page.WaitForLoadStateAsync(LoadState.NetworkIdle);
    }

    private async Task NavigateToActionsTabAsync()
    {
        if (_page == null) return;
        
        LogStatus("Clicking Actions tab...");
        var actionsTab = _page.Locator("text=Actions").First;
        await actionsTab.ClickAsync();
        await _page.WaitForLoadStateAsync(LoadState.NetworkIdle);
        await Task.Delay(300);
    }

    private async Task RecommendOfferAsync()
    {
        if (_page == null) return;
        
        LogStatus("Clicking 'Recommend Offer or Reject'...");
        var recommendLink = _page.Locator("text=Recommend Offer or Reject").First;
        await recommendLink.ClickAsync();
        await _page.WaitForLoadStateAsync(LoadState.NetworkIdle);
        await Task.Delay(300);

        LogStatus("Selecting 'Offer recommendation'...");
        var offerRadio = _page.Locator("text=Offer recommendation").First;
        await offerRadio.ClickAsync();
        await Task.Delay(200);

        if (DebugMode)
        {
            LogStatus("DEBUG MODE: Paused before clicking Process");
            LogStatus("Verify 'Offer recommendation' is selected, then manually click Process if correct");
            return;
        }

        LogStatus("Clicking Process...");
        var processBtn = _page.Locator("input[value='Process']").First;
        await processBtn.ClickAsync();
        await _page.WaitForLoadStateAsync(LoadState.NetworkIdle);
        await Task.Delay(500);
        
        LogStatus("Offer recommendation processed.");
    }

    private async Task RecommendRejectAsync()
    {
        if (_page == null) return;
        
        LogStatus("Clicking 'Recommend Offer or Reject'...");
        var recommendLink = _page.Locator("a").Filter(new() { HasText = "Recommend Offer or Reject" }).First;
        if (!await recommendLink.IsVisibleAsync())
        {
            recommendLink = _page.Locator("text=Recommend Offer or Reject").First;
        }
        await recommendLink.ClickAsync();
        await _page.WaitForLoadStateAsync(LoadState.NetworkIdle);
        await Task.Delay(1000);

        LogStatus("Selecting 'Reject' radio button...");
        var radioButtons = await _page.Locator("input[type='radio']").AllAsync();
        
        if (radioButtons.Count >= 2)
        {
            await radioButtons[1].ClickAsync();
        }
        else
        {
            var rejectRadio = _page.GetByLabel("Reject");
            await rejectRadio.ClickAsync();
        }
        
        await Task.Delay(2000);
        
        LogStatus("Selecting Reason 1 dropdown option 8...");
        var allSelects = await _page.Locator("select").AllAsync();
        
        if (allSelects.Count < 2)
        {
            throw new Exception($"Expected at least 2 dropdowns, found {allSelects.Count}");
        }
        
        var jsCode = @"
            () => {
                const selects = document.querySelectorAll('select');
                if (selects.length < 2) return 'ERROR: Need at least 2 dropdowns';
                
                const select = selects[1];
                let debugInfo = 'Reason 1 options:\n';
                
                for (let i = 0; i < select.options.length; i++) {
                    debugInfo += '  ' + i + ': ' + select.options[i].text + '\n';
                }
                
                let foundOption = null;
                
                for (let i = 0; i < select.options.length; i++) {
                    const text = select.options[i].text;
                    if ((text.startsWith('8.') || text.startsWith('8 ')) && 
                        text.toLowerCase().includes('not competitive')) {
                        foundOption = i;
                        break;
                    }
                }
                
                if (foundOption === null) {
                    for (let i = 0; i < select.options.length; i++) {
                        const text = select.options[i].text.toLowerCase();
                        if (text.includes('not competitive') && text.includes('oversubscribed')) {
                            foundOption = i;
                            break;
                        }
                    }
                }
                
                if (foundOption !== null) {
                    select.selectedIndex = foundOption;
                    select.value = select.options[foundOption].value;
                    select.dispatchEvent(new Event('change', { bubbles: true }));
                    
                    return 'SUCCESS: Selected option ' + foundOption + ': ' + select.options[foundOption].text;
                }
                
                return 'FAILED: Could not find option 8\n' + debugInfo;
            }
        ";
        
        var jsResult = await _page.EvaluateAsync<string>(jsCode);
        LogStatus(jsResult);
        
        if (!jsResult.Contains("SUCCESS"))
        {
            throw new Exception("Failed to select option 8 in Reason 1 dropdown\n" + jsResult);
        }
        
        await Task.Delay(1000);

        if (DebugMode)
        {
            LogStatus("DEBUG MODE: Paused before clicking Process");
            LogStatus("Verify 'Reject' is selected and 'Reason 1' shows option 8");
            LogStatus("Manually click Process if correct");
            return;
        }

        LogStatus("Clicking Process button...");
        var processBtn = _page.Locator("input[value='Process']").First;
        if (!await processBtn.IsVisibleAsync())
        {
            processBtn = _page.Locator("button").Filter(new() { HasText = "Process" }).First;
        }
        await processBtn.ClickAsync();
        await _page.WaitForLoadStateAsync(LoadState.NetworkIdle);
        await Task.Delay(500);
        
        LogStatus("Rejection processed.");
    }

    public async Task ProcessStudentMergeOverviewAsync(StudentRecord student, string downloadPath)
    {
        if (_page == null) throw new InvalidOperationException("Service not initialised.");

        student.Status = ProcessingStatus.Processing;
        StudentProcessed?.Invoke(this, student);

        try
        {
            LogStatus($"=== MERGE OVERVIEW: {student.StudentNo} ===");

            // Step 1: Search and enter student record
            LogStatus("[Step 1] Searching for student...");
            await SearchForStudentAsync(student.StudentNo);
            await ClickStudentLinkAsync(student);
            LogStatus("[Step 1] Entered student record.");

            // Step 2: Click "Documents & Uploads" tab
            LogStatus("[Step 2] Clicking 'Documents & Uploads' tab...");
            var docsTab = _page.Locator("a").Filter(new() { HasText = "Documents" }).First;
            await docsTab.ClickAsync();
            await _page.WaitForLoadStateAsync(LoadState.NetworkIdle);
            await Task.Delay(1000);
            LogStatus("[Step 2] Documents tab loaded.");

            // Step 3: Click "Create Overview.pdf" or "Amend Overview.pdf"
            LogStatus("[Step 3] Looking for Create/Amend Overview button...");

            // Debug: log every input/button on the page so we can see what's there
            var pageButtons = await _page.EvaluateAsync<string>(@"() => {
                const elements = document.querySelectorAll('input[type=submit], input[type=button], button, input[type=reset]');
                return Array.from(elements).map(el =>
                    el.tagName + ' | type=' + el.type + ' | value=""' + (el.value || '') + '"" | text=""' + (el.textContent || '').trim() + '""'
                ).join('\n');
            }");
            LogStatus($"[Step 3] Buttons found on page:\n{pageButtons}");

            // Use JavaScript to find and click the overview button — most reliable approach
            var clickResult = await _page.EvaluateAsync<string>(@"() => {
                const elements = document.querySelectorAll('input, button, a');
                for (const el of elements) {
                    const text = ((el.value || '') + ' ' + (el.textContent || '')).toLowerCase();
                    if (text.includes('create overview') || text.includes('amend overview')) {
                        el.click();
                        return 'CLICKED: ' + el.tagName + ' | ' + (el.value || el.textContent || '').trim();
                    }
                }
                return 'NOT_FOUND';
            }");
            LogStatus($"[Step 3] {clickResult}");

            if (clickResult == "NOT_FOUND")
                throw new Exception("Could not find any Create/Amend Overview button on the page.");

            // The JS click triggers a form POST. The page will navigate.
            // Poll until we can confirm the page has reloaded (readyState goes through loading→complete)
            LogStatus("[Step 3] Waiting for page to load (up to 2 minutes)...");
            var waited = 0;
            while (waited < 120000)
            {
                await Task.Delay(2000);
                waited += 2000;
                try
                {
                    var state = await _page.EvaluateAsync<string>("() => document.readyState");
                    if (state == "complete")
                    {
                        // Page has loaded — but is it the NEW page? Check if the overview button is gone
                        var stillHasOverview = await _page.EvaluateAsync<bool>(@"() => {
                            const elements = document.querySelectorAll('input, button, a');
                            for (const el of elements) {
                                const text = ((el.value || '') + ' ' + (el.textContent || '')).toLowerCase();
                                if (text.includes('create overview') || text.includes('amend overview')) return true;
                            }
                            return false;
                        }");
                        // If the overview button is gone, the page has changed
                        if (!stillHasOverview)
                        {
                            LogStatus($"[Step 3] Page reloaded (took ~{waited / 1000}s).");
                            break;
                        }
                    }
                }
                catch
                {
                    // Page is mid-navigation, keep waiting
                    LogStatus($"[Step 3] Still loading... ({waited / 1000}s)");
                }
            }
            await _page.WaitForLoadStateAsync(LoadState.NetworkIdle, new PageWaitForLoadStateOptions { Timeout = 60000 });
            await Task.Delay(2000);
            LogStatus("[Step 3] Page loaded after overview creation.");

            // Step 4: Click "Merge Documents"
            LogStatus("[Step 4] Looking for 'Merge Documents' button...");
            ILocator? mergeBtn = null;

            var mergeBtnSelectors = new[]
            {
                "input[value='Merge Documents']",
                "button:has-text('Merge Documents')",
                "a:has-text('Merge Documents')",
            };

            foreach (var selector in mergeBtnSelectors)
            {
                var candidate = _page.Locator(selector).First;
                if (await candidate.IsVisibleAsync())
                {
                    mergeBtn = candidate;
                    LogStatus($"[Step 4] Found with selector: {selector}");
                    break;
                }
            }

            if (mergeBtn == null)
                throw new Exception("Could not find 'Merge Documents' button on the page.");

            await mergeBtn.ScrollIntoViewIfNeededAsync();
            await Task.Delay(500);
            LogStatus("[Step 4] Clicking 'Merge Documents'...");
            await mergeBtn.ClickAsync();
            await Task.Delay(2000);

            // Step 5: Confirmation modal — click "Yes"
            LogStatus("[Step 5] Waiting for confirmation modal...");
            ILocator? yesBtn = null;

            var yesBtnSelectors = new[]
            {
                "input[value='Yes']",
                "button:has-text('Yes')",
                "a:has-text('Yes')",
            };

            // Wait for the modal to appear
            for (int attempt = 0; attempt < 10; attempt++)
            {
                foreach (var selector in yesBtnSelectors)
                {
                    var candidate = _page.Locator(selector).First;
                    if (await candidate.IsVisibleAsync())
                    {
                        yesBtn = candidate;
                        break;
                    }
                }
                if (yesBtn != null) break;
                await Task.Delay(1000);
            }

            if (yesBtn == null)
                throw new Exception("Confirmation modal 'Yes' button did not appear.");

            LogStatus("[Step 5] Clicking 'Yes' via JS and waiting for merge processing...");
            await _page.EvaluateAsync(@"() => {
                const elements = document.querySelectorAll('input, button, a');
                for (const el of elements) {
                    const text = ((el.value || '') + ' ' + (el.textContent || '')).toLowerCase().trim();
                    if (text === 'yes') { el.click(); return; }
                }
            }");

            // Poll until the Yes button disappears (modal closed, page navigating)
            var yesWaited = 0;
            while (yesWaited < 120000)
            {
                await Task.Delay(2000);
                yesWaited += 2000;
                try
                {
                    var state = await _page.EvaluateAsync<string>("() => document.readyState");
                    // Check if the Yes button / modal is gone
                    var stillHasYes = await _page.EvaluateAsync<bool>(@"() => {
                        const elements = document.querySelectorAll('input, button');
                        for (const el of elements) {
                            const text = ((el.value || '') + ' ' + (el.textContent || '')).toLowerCase().trim();
                            if (text === 'yes') return true;
                        }
                        return false;
                    }");
                    if (state == "complete" && !stillHasYes)
                    {
                        LogStatus($"[Step 5] Merge complete (took ~{yesWaited / 1000}s).");
                        break;
                    }
                }
                catch
                {
                    LogStatus($"[Step 5] Still processing... ({yesWaited / 1000}s)");
                }
            }
            await _page.WaitForLoadStateAsync(LoadState.NetworkIdle, new PageWaitForLoadStateOptions { Timeout = 60000 });
            await Task.Delay(2000);
            LogStatus("[Step 5] Merge processing complete.");

            // Step 6: Download the overview PDF
            LogStatus("[Step 6] Looking for overview PDF link...");
            var pdfLink = _page.Locator("a").Filter(
                new() { HasTextRegex = new System.Text.RegularExpressions.Regex(
                    $@"{System.Text.RegularExpressions.Regex.Escape(student.StudentNo)}.*OVERVIEW",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase) }).First;

            if (await pdfLink.IsVisibleAsync())
            {
                Directory.CreateDirectory(downloadPath);
                LogStatus("[Step 6] Downloading PDF...");
                var download = await _page.RunAndWaitForDownloadAsync(async () =>
                {
                    await pdfLink.ClickAsync();
                });

                var fileName = download.SuggestedFilename;
                if (string.IsNullOrEmpty(fileName))
                    fileName = $"{student.StudentNo}-01-01-OVERVIEW.PDF";

                var savePath = Path.Combine(downloadPath, fileName);
                await download.SaveAsAsync(savePath);
                LogStatus($"[Step 6] PDF saved: {savePath}");
            }
            else
            {
                LogStatus("[Step 6] WARNING: Overview PDF link not found. Continuing...");
            }

            // Step 7: Click "Exit"
            LogStatus("[Step 7] Clicking 'Exit'...");
            var exitBtn = _page.Locator("input[value='Exit']").First;
            if (!await exitBtn.IsVisibleAsync())
                exitBtn = _page.Locator("button:has-text('Exit')").First;
            if (!await exitBtn.IsVisibleAsync())
                exitBtn = _page.Locator("a:has-text('Exit')").First;
            await exitBtn.ScrollIntoViewIfNeededAsync();
            await Task.Delay(300);
            await exitBtn.ClickAsync();
            await _page.WaitForLoadStateAsync(LoadState.NetworkIdle, new PageWaitForLoadStateOptions { Timeout = 30000 });
            await Task.Delay(500);
            LogStatus("[Step 7] Exited record.");

            // Step 8: Click "Search" to go back
            LogStatus("[Step 8] Returning to search screen...");
            var searchNav = _page.Locator("a").Filter(new() { HasText = "Search" }).First;
            await searchNav.ClickAsync();
            await _page.WaitForLoadStateAsync(LoadState.NetworkIdle, new PageWaitForLoadStateOptions { Timeout = 30000 });
            await Task.Delay(500);
            LogStatus("[Step 8] Back on search screen.");

            student.Status = ProcessingStatus.Success;
            LogStatus($"=== SUCCESS: {student.StudentNo} ===");
        }
        catch (Exception ex)
        {
            student.Status = ProcessingStatus.Failed;
            student.ErrorMessage = ex.Message;
            LogStatus($"=== FAILED {student.StudentNo}: {ex.Message} ===");

            // RECOVERY: Navigate back to search screen so the next student can be processed
            try
            {
                LogStatus("Recovering — navigating back to search...");
                await NavigateToUclSelectAsync();
            }
            catch
            {
                LogStatus("Recovery failed — browser may be in an unexpected state.");
            }
        }

        StudentProcessed?.Invoke(this, student);
    }

    // ============================ PAT — Personal Tutor Allocations ============================
    //
    // Flow (from the My Portico home, after login):
    //   1. NavigateToPersonalTutorAllocationsAsync — click "Personal Tutor Allocations", land on
    //      the "All Students" sub-tab.
    //   2. ApplyPatDepartmentAsync("Computer Science") — set the Department, Apply New Criteria. ONCE.
    //   3. ApplyPatProgrammeAsync(code) — type the route code (e.g. TMSDATSMLE01) into the Programme
    //      filter box, Apply New Criteria. ONCE per programme.
    //   4. ProcessStudentPatAsync — PER STUDENT: type the student number into Student ID, Apply New
    //      Criteria; the student's row appears with an editable "Personal Tutor" box; type the PIVOT
    //      code into it, let it resolve to the staff name, click Save, then OK on the confirmation.
    //
    // The page is a SITS/eVision grid. Selectors are defensive (JS text matching + placeholder
    // lookups + logging) because the live page can only be tuned against an authenticated session.

    public async Task<bool> NavigateToPersonalTutorAllocationsAsync()
    {
        if (_page == null) throw new InvalidOperationException("Service not initialised.");

        LogStatus("Navigating to Personal Tutor Allocations...");

        // Make sure we're on the My Portico home where the "Personal Tutors" container lives.
        bool homeVisible = false;
        try { homeVisible = await _page.Locator("text=My Portico").First.IsVisibleAsync(); } catch { }
        if (!homeVisible)
        {
            await _page.GotoAsync(_config?.PorticoUrl ?? "https://evision.ucl.ac.uk/urd/sits.urd/run/siw_lgn");
            await _page.WaitForSelectorAsync("text=My Portico", new PageWaitForSelectorOptions { Timeout = 30000 });
        }

        // Click the "Personal Tutor Allocations" link — not "... User Guide"/"... Export ..."/dashboard.
        var clicked = await _page.EvaluateAsync<string>(@"() => {
            const norm = s => (s || '').replace(/\s+/g, ' ').trim().toLowerCase();
            const target = 'personal tutor allocations';
            const links = Array.from(document.querySelectorAll('a'));
            for (const a of links) { if (norm(a.textContent) === target) { a.scrollIntoView(); a.click(); return 'CLICKED_EXACT'; } }
            for (const a of links) {
                const t = norm(a.textContent);
                if (t.startsWith(target) && !t.includes('guide') && !t.includes('export')) { a.scrollIntoView(); a.click(); return 'CLICKED_STARTS'; }
            }
            return 'NOT_FOUND';
        }");
        LogStatus($"[PAT] Allocations link: {clicked}");
        if (clicked == "NOT_FOUND")
            throw new Exception("Could not find the 'Personal Tutor Allocations' link on the My Portico home page.");

        await _page.WaitForLoadStateAsync(LoadState.NetworkIdle, new PageWaitForLoadStateOptions { Timeout = 30000 });
        await Task.Delay(1500);

        try { await _page.WaitForSelectorAsync("text=Personal Tutor Allocations", new PageWaitForSelectorOptions { Timeout = 15000 }); }
        catch { LogStatus("[PAT] WARNING: 'Personal Tutor Allocations' heading not detected — layout may differ."); }

        // Ensure the "All Students" sub-tab (recover if a prior run left "Unassigned Students" selected).
        var sub = await _page.EvaluateAsync<string>(@"() => {
            const norm = s => (s || '').replace(/\s+/g, ' ').trim().toLowerCase();
            const links = Array.from(document.querySelectorAll('a'));
            for (const a of links) { if (norm(a.textContent) === 'all students') { a.scrollIntoView(); a.click(); return 'CLICKED'; } }
            return 'NOT_FOUND';
        }");
        if (sub == "CLICKED")
        {
            await _page.WaitForLoadStateAsync(LoadState.NetworkIdle, new PageWaitForLoadStateOptions { Timeout = 30000 });
            await Task.Delay(1000);
        }
        LogStatus("[PAT] On Personal Tutor Allocations page (All Students).");
        return true;
    }

    // ONE-OFF setup: set the Department combobox, then Apply New Criteria.
    public async Task ApplyPatDepartmentAsync(string department)
    {
        if (_page == null) throw new InvalidOperationException("Service not initialised.");

        LogStatus($"[PAT] Setting Department = '{department}'...");

        // The Department control sits next to the "Department:" label and behaves like a searchable
        // dropdown. Open it, type the name, then click the matching option (handles select2 / native).
        var dept = _page.GetByPlaceholder("Department", new() { Exact = false }).First;
        try
        {
            if (await dept.CountAsync() > 0 && await dept.IsVisibleAsync())
                await dept.ClickAsync();
            else
                await _page.EvaluateAsync(@"() => {
                    const norm = s => (s || '').replace(/\s+/g,' ').trim().toLowerCase();
                    const lab = Array.from(document.querySelectorAll('*')).find(e => e.children.length === 0 && /^department:?$/.test(norm(e.textContent)));
                    if (lab) { const box = (lab.parentElement && lab.parentElement.querySelector('select, input, .select2, [role=combobox]')) || lab.nextElementSibling; if (box) box.click(); }
                }");
        }
        catch { }

        await Task.Delay(500);
        await _page.Keyboard.TypeAsync(department, new KeyboardTypeOptions { Delay = 60 });
        await Task.Delay(1500);

        var picked = await _page.EvaluateAsync<string>(@"(dep) => {
            const norm = s => (s || '').replace(/\s+/g,' ').trim().toLowerCase();
            const target = norm(dep);
            const opts = Array.from(document.querySelectorAll('li, option, [role=option], .select2-results__option, .dropdown-item, a, span, div'));
            // Native <select>: set the value directly.
            for (const o of opts) {
                if (o.tagName === 'OPTION' && norm(o.textContent) === target) {
                    const sel = o.closest('select'); if (sel) { sel.value = o.value; sel.dispatchEvent(new Event('change', { bubbles: true })); return 'SELECTED_OPTION'; }
                }
            }
            for (const o of opts) { if (o.offsetParent === null) continue; if (norm(o.textContent) === target) { o.scrollIntoView(); o.click(); return 'CLICKED'; } }
            for (const o of opts) { if (o.offsetParent === null) continue; const t = norm(o.textContent); if (t.includes(target) && t.length < 60) { o.scrollIntoView(); o.click(); return 'CLICKED_CONTAINS: ' + t; } }
            return 'NOT_FOUND';
        }", department);
        LogStatus($"[PAT] Department option: {picked}");
        await Task.Delay(500);

        await ClickApplyNewCriteriaAsync();
    }

    // ONE-OFF-per-programme setup: type the route code into the Programme filter box, Apply New Criteria.
    public async Task ApplyPatProgrammeAsync(string programmeCode)
    {
        if (_page == null) throw new InvalidOperationException("Service not initialised.");

        LogStatus($"[PAT] Setting Programme (route code) = '{programmeCode}'...");

        var input = await ResolvePatInputAsync(new[] { "Programme" });
        if (input == null)
            throw new Exception("Could not find the 'Programme' filter field.");

        await input.ClickAsync();
        await input.FillAsync("");
        await _page.Keyboard.TypeAsync(programmeCode, new KeyboardTypeOptions { Delay = 60 });
        await Task.Delay(1500);
        LogStatus($"[PAT] Programme field now holds: '{await ReadInputAsync(input)}'");

        await ClickApplyNewCriteriaAsync();
    }

    public async Task ProcessStudentPatAsync(StudentRecord student)
    {
        if (_page == null) throw new InvalidOperationException("Service not initialised.");

        student.Status = ProcessingStatus.Processing;
        StudentProcessed?.Invoke(this, student);

        try
        {
            LogStatus($"=== PAT: {student.StudentNo} → {student.PersonalTutor} ===");

            if (string.IsNullOrWhiteSpace(student.PersonalTutor))
                throw new Exception("No Personal Tutor code (PIVOT column) for this student.");

            // 1. Enter the student number in the Student ID filter and Apply New Criteria.
            await EnterStudentFilterAsync(student.StudentNo);
            await ClickApplyNewCriteriaAsync();

            // 2. Wait for the student's result row (with its editable Personal Tutor box) to appear.
            await WaitForResultRowAsync(student.StudentNo);

            // 3. Type the tutor code into the row's Personal Tutor box and let it resolve to the name.
            //    If the row already has a tutor, skip without saving (no overwrite).
            var assigned = await AssignTutorInResultRowAsync(student);
            if (!assigned)
            {
                student.Status = ProcessingStatus.Skipped;
                student.ErrorMessage = "Already allocated in Portico — skipped (not overwritten).";
                LogStatus($"=== SKIPPED {student.StudentNo}: already has a tutor ===");
                StudentProcessed?.Invoke(this, student);
                return;
            }

            if (DebugMode)
            {
                LogStatus("[PAT] DEBUG MODE: paused before 'Save'. Verify the tutor resolved on the row, then click Save/OK manually.");
                student.Status = ProcessingStatus.Success;
                StudentProcessed?.Invoke(this, student);
                return;
            }

            // 4. Save, then OK on the "Personal Tutors Updated" confirmation.
            await ClickSaveAsync();
            await ClickOkConfirmationAsync();

            student.Status = ProcessingStatus.Success;
            LogStatus($"=== SUCCESS: {student.StudentNo} allocated to {student.PersonalTutor} ===");
        }
        catch (Exception ex)
        {
            student.Status = ProcessingStatus.Failed;
            student.ErrorMessage = ex.Message;
            LogStatus($"=== FAILED {student.StudentNo}: {ex.Message} ===");
        }

        StudentProcessed?.Invoke(this, student);
    }

    // Finds a visible filter input by trying each placeholder candidate (case-insensitive substring).
    private async Task<ILocator?> ResolvePatInputAsync(string[] placeholderCandidates)
    {
        if (_page == null) return null;

        foreach (var ph in placeholderCandidates)
        {
            var loc = _page.GetByPlaceholder(ph, new() { Exact = false }).First;
            try
            {
                if (await loc.CountAsync() > 0 && await loc.IsVisibleAsync())
                {
                    LogStatus($"[PAT] Matched filter field by placeholder '{ph}'.");
                    return loc;
                }
            }
            catch { /* try the next candidate */ }
        }
        return null;
    }

    // Reads back what an input currently holds — proves whether typing actually landed.
    private static async Task<string> ReadInputAsync(ILocator input)
    {
        try { return await input.InputValueAsync(); } catch { return "<unreadable>"; }
    }

    // Type the student number into the Student ID filter (overwriting any previous student).
    private async Task EnterStudentFilterAsync(string studentNo)
    {
        if (_page == null) return;

        LogStatus($"[PAT] Entering student number {studentNo}...");

        var input = await ResolvePatInputAsync(new[] { "Student ID or Name", "Student ID", "Student" });
        if (input == null)
            throw new Exception("Could not find the 'Student ID or Name' filter field.");

        await input.ClickAsync();
        await input.FillAsync("");
        await _page.Keyboard.TypeAsync(studentNo, new KeyboardTypeOptions { Delay = 60 });
        LogStatus($"[PAT] Student ID field now holds: '{await ReadInputAsync(input)}'");
        await Task.Delay(200);
    }

    // Polls for the student's result row (a row whose text contains the student number) to render.
    private async Task WaitForResultRowAsync(string studentNo)
    {
        if (_page == null) return;

        LogStatus("[PAT] Waiting for the student's result row...");
        for (int attempt = 0; attempt < 20; attempt++)   // up to ~20s
        {
            var has = await _page.EvaluateAsync<bool>(@"(no) => {
                const rows = document.querySelectorAll('tr');
                for (const r of rows) {
                    if (r.offsetParent === null) continue;
                    if ((r.textContent || '').includes(no) && r.querySelector('input')) return true;
                }
                return false;
            }", studentNo);
            if (has) { LogStatus("[PAT] Result row is present."); return; }
            await Task.Delay(1000);
        }
        throw new Exception($"Student {studentNo}'s result row did not appear after Apply New Criteria.");
    }

    // Types the tutor code into the Personal Tutor box ON THE STUDENT'S ROW and accepts the match.
    // Returns true if a tutor was entered; false if the row already had one (left untouched).
    private async Task<bool> AssignTutorInResultRowAsync(StudentRecord student)
    {
        if (_page == null) return false;

        // The result row is the <tr> that contains the student number (the filter row does not).
        var row = _page.Locator("tr").Filter(new() { HasText = student.StudentNo });
        var tutorInput = row.GetByPlaceholder("Personal Tutor", new() { Exact = false }).Last;
        if (await tutorInput.CountAsync() == 0)
            throw new Exception("Could not find the Personal Tutor box on the student's row.");

        // Skip if the row already has a tutor, so an existing allocation is never overwritten.
        var existing = (await ReadInputAsync(tutorInput)).Trim();
        if (!string.IsNullOrWhiteSpace(existing) && existing != "<unreadable>")
        {
            student.PersonalTutorName = existing;
            LogStatus($"[PAT] {student.StudentNo} already has a tutor ('{existing}') — skipping (not overwritten).");
            return false;
        }

        LogStatus($"[PAT] Entering tutor code {student.PersonalTutor} on {student.StudentNo}'s row...");

        await tutorInput.ScrollIntoViewIfNeededAsync();
        await tutorInput.ClickAsync();
        try { await tutorInput.FillAsync(""); } catch { }
        await _page.Keyboard.TypeAsync(student.PersonalTutor, new KeyboardTypeOptions { Delay = 40 });

        var code = student.PersonalTutor;

        // Wait responsively for the autocomplete suggestion to render — break as soon as it appears.
        bool suggestionVisible = false;
        for (int i = 0; i < 40; i++)   // up to ~6s, 150ms granularity
        {
            if (await IsTutorSuggestionOpenAsync(code)) { suggestionVisible = true; break; }
            await Task.Delay(150);
        }
        if (!suggestionVisible)
            LogStatus("[PAT] WARNING: tutor suggestion did not appear — attempting selection anyway.");

        // Capture the staff name the code resolved to (for the results CSV) before the dropdown closes.
        var suggestionText = await GetTutorSuggestionTextAsync(code);
        if (!string.IsNullOrWhiteSpace(suggestionText))
        {
            student.PersonalTutorName = ExtractStaffName(suggestionText, code);
            LogStatus($"[PAT] Resolved tutor: {student.PersonalTutorName}");
        }

        // Commit the selection by keyboard — a synthetic el.click() on this widget does NOT register
        // the choice (the hidden staff-id stays empty and Save stalls): highlight first match, Enter.
        await _page.Keyboard.PressAsync("ArrowDown");
        await Task.Delay(100);
        await _page.Keyboard.PressAsync("Enter");

        // Wait responsively for the dropdown to close (= committed). Fall back to real mouse events.
        if (!await WaitForTutorCommittedAsync(code, 3000))
        {
            LogStatus("[PAT] Keyboard select did not commit — trying mouse-event fallback...");
            var picked = await _page.EvaluateAsync<string>(@"(code) => {
                const norm = s => (s || '').replace(/\s+/g, ' ').trim();
                const cl = code.toLowerCase();
                const els = Array.from(document.querySelectorAll('li, [role=option], .select2-results__option, .dropdown-item, .ui-menu-item, a, span, div'));
                for (const e of els) {
                    if (e.offsetParent === null) continue;
                    const t = norm(e.textContent);
                    if (!t || t.length > 90) continue;
                    if (t.toLowerCase().includes(cl)) {
                        e.scrollIntoView();
                        const opts = { bubbles: true, cancelable: true, view: window };
                        e.dispatchEvent(new MouseEvent('mousedown', opts));
                        e.dispatchEvent(new MouseEvent('mouseup', opts));
                        e.dispatchEvent(new MouseEvent('click', opts));
                        return 'CLICKED: ' + t;
                    }
                }
                return 'NO_SUGGESTION';
            }", code);
            LogStatus($"[PAT] Tutor suggestion (fallback): {picked}");
            await WaitForTutorCommittedAsync(code, 2000);
        }

        if (await IsTutorSuggestionOpenAsync(code))
            LogStatus($"[PAT] WARNING: tutor '{code}' may not have resolved on {student.StudentNo}'s row — verify before Save.");
        else
            LogStatus($"[PAT] Tutor resolved on {student.StudentNo}'s row.");

        return true;
    }

    // Polls (every 150ms, up to timeoutMs) until the tutor autocomplete dropdown has closed.
    private async Task<bool> WaitForTutorCommittedAsync(string code, int timeoutMs)
    {
        var waited = 0;
        while (waited < timeoutMs)
        {
            if (!await IsTutorSuggestionOpenAsync(code)) return true;
            await Task.Delay(150);
            waited += 150;
        }
        return !await IsTutorSuggestionOpenAsync(code);
    }

    // Returns the full text of the autocomplete option for this code (e.g. "DADAM27 Dmitry Adamskiy"),
    // or "" if no option is currently visible.
    private async Task<string> GetTutorSuggestionTextAsync(string code)
    {
        if (_page == null) return "";
        return await _page.EvaluateAsync<string>(@"(code) => {
            const norm = s => (s || '').replace(/\s+/g, ' ').trim();
            const cl = code.toLowerCase();
            const els = Array.from(document.querySelectorAll('li, [role=option], .select2-results__option, .dropdown-item, .ui-menu-item'));
            for (const e of els) {
                if (e.offsetParent === null) continue;
                const t = norm(e.textContent);
                if (t && t.length <= 90 && t.toLowerCase().includes(cl)) return t;
            }
            return '';
        }", code);
    }

    // Strips the tutor code (and any leftover separators/brackets) from the suggestion text,
    // leaving just the staff name. "DADAM27 Dmitry Adamskiy" / "Dmitry Adamskiy (DADAM27)" → "Dmitry Adamskiy".
    private static string ExtractStaffName(string suggestionText, string code)
    {
        var name = System.Text.RegularExpressions.Regex.Replace(
            suggestionText, System.Text.RegularExpressions.Regex.Escape(code), "",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        name = name.Replace("()", " ");
        name = System.Text.RegularExpressions.Regex.Replace(name, @"\s+", " ").Trim();
        return name.Trim('(', ')', '/', '-', '–', ',', ' ');
    }

    // True while the tutor autocomplete dropdown is still showing an option for this code
    // (i.e. the selection has NOT yet been committed).
    private async Task<bool> IsTutorSuggestionOpenAsync(string code)
    {
        if (_page == null) return false;
        return await _page.EvaluateAsync<bool>(@"(code) => {
            const cl = code.toLowerCase();
            const els = Array.from(document.querySelectorAll('li, [role=option], .select2-results__option, .dropdown-item, .ui-menu-item'));
            return els.some(e => {
                if (e.offsetParent === null) return false;
                const t = (e.textContent || '').replace(/\s+/g, ' ').trim();
                return t.length > 0 && t.length <= 90 && t.toLowerCase().includes(cl);
            });
        }", code);
    }

    private async Task ClickSaveAsync()
    {
        if (_page == null) return;

        LogStatus("[PAT] Clicking 'Save'...");
        var clicked = await _page.EvaluateAsync<string>(@"() => {
            const norm = s => (s || '').replace(/\s+/g, ' ').trim().toLowerCase();
            const els = document.querySelectorAll('input[type=submit], input[type=button], button, a');
            for (const el of els) { if (el.offsetParent === null) continue; const t = norm(el.value) || norm(el.textContent); if (t === 'save') { el.scrollIntoView(); el.click(); return 'CLICKED'; } }
            return 'NOT_FOUND';
        }");
        LogStatus($"[PAT] Save: {clicked}");
        if (clicked == "NOT_FOUND")
            throw new Exception("Could not find the 'Save' button.");

        await _page.WaitForLoadStateAsync(LoadState.NetworkIdle, new PageWaitForLoadStateOptions { Timeout = 30000 });
        await Task.Delay(400);
    }

    // Waits for the "Personal Tutors Updated" confirmation and clicks OK.
    private async Task ClickOkConfirmationAsync()
    {
        if (_page == null) return;

        LogStatus("[PAT] Waiting for 'Personal Tutors Updated' confirmation...");
        for (int attempt = 0; attempt < 30; attempt++)   // poll every 200ms, up to ~6s
        {
            var clicked = await _page.EvaluateAsync<string>(@"() => {
                const norm = s => (s || '').replace(/\s+/g, ' ').trim().toLowerCase();
                const els = document.querySelectorAll('input[type=submit], input[type=button], button, a');
                for (const el of els) { if (el.offsetParent === null) continue; const t = norm(el.value) || norm(el.textContent); if (t === 'ok') { el.scrollIntoView(); el.click(); return 'CLICKED'; } }
                return 'NOT_FOUND';
            }");
            if (clicked == "CLICKED")
            {
                LogStatus("[PAT] Confirmation OK clicked.");
                await _page.WaitForLoadStateAsync(LoadState.NetworkIdle, new PageWaitForLoadStateOptions { Timeout = 30000 });
                await Task.Delay(400);
                return;
            }
            await Task.Delay(200);
        }
        LogStatus("[PAT] WARNING: no 'OK' confirmation button appeared — continuing.");
    }

    private async Task ClickApplyNewCriteriaAsync()
    {
        if (_page == null) return;

        LogStatus("[PAT] Clicking 'Apply New Criteria'...");
        var clicked = await _page.EvaluateAsync<string>(@"() => {
            const norm = s => (s || '').replace(/\s+/g, ' ').trim().toLowerCase();
            const els = document.querySelectorAll('input[type=submit], input[type=button], button, a');
            for (const el of els) {
                const t = norm(el.value) || norm(el.textContent);
                if (t === 'apply new criteria' || t === 'apply criteria' || (t.includes('apply') && t.includes('criteria'))) {
                    el.scrollIntoView(); el.click(); return 'CLICKED: ' + t;
                }
            }
            return 'NOT_FOUND';
        }");
        LogStatus($"[PAT] {clicked}");
        if (clicked == "NOT_FOUND")
            throw new Exception("Could not find the 'Apply New Criteria' button.");

        await _page.WaitForLoadStateAsync(LoadState.NetworkIdle, new PageWaitForLoadStateOptions { Timeout = 30000 });
        await Task.Delay(1200);
    }

    public async Task<string> DownloadDepartmentReportAsync(string fullProgrammeName, string downloadDir)
    {
        if (_page == null) throw new InvalidOperationException("Service not initialised.");

        LogStatus("Looking for Department Application Reports link...");
        var deptLink = _page.Locator("a").Filter(new() { HasText = "Department Application Reports" }).First;
        if (!await deptLink.IsVisibleAsync())
        {
            LogStatus("Link not visible — navigating to portal home...");
            await _page.GotoAsync("https://evision.ucl.ac.uk/urd/sits.urd/run/siw_lgn");
            await _page.WaitForSelectorAsync("text=My Portico", new PageWaitForSelectorOptions { Timeout = 30000 });
            deptLink = _page.Locator("a").Filter(new() { HasText = "Department Application Reports" }).First;
        }

        await deptLink.ClickAsync();
        await _page.WaitForLoadStateAsync(LoadState.NetworkIdle);
        await Task.Delay(1000);
        LogStatus("On Report Parameters page.");

        // Clear any existing programme code tag selections
        var removed = await _page.EvaluateAsync<int>(@"() => {
            const btns = document.querySelectorAll('button[title=""Remove""], .sit-remove, [class*=""remove""][class*=""btn""]');
            btns.forEach(b => b.click());
            return btns.length;
        }");
        if (removed > 0) await Task.Delay(400);

        // Find the Programme code autocomplete input — it is typically the last text input on the form
        LogStatus($"Entering programme name: {fullProgrammeName}");
        var textInputs = _page.Locator("input[type='text']");
        int inputCount = await textInputs.CountAsync();
        var progCodeInput = inputCount >= 2 ? textInputs.Nth(inputCount - 1) : textInputs.First;

        await progCodeInput.ClickAsync();
        await progCodeInput.ClearAsync();

        // Type first 8 chars to trigger the autocomplete dropdown
        var searchTerm = fullProgrammeName.Length > 8 ? fullProgrammeName[..8] : fullProgrammeName;
        await _page.Keyboard.TypeAsync(searchTerm, new KeyboardTypeOptions { Delay = 90 });
        await Task.Delay(2000);

        // Click the matching suggestion
        var suggestion = _page.Locator("li").Filter(new() { HasText = fullProgrammeName }).First;
        if (!await suggestion.IsVisibleAsync())
            suggestion = _page.Locator($"[class*='autocomplete'] li, [class*='dropdown'] li, ul li").Filter(new() { HasText = fullProgrammeName[..Math.Min(15, fullProgrammeName.Length)] }).First;

        if (await suggestion.IsVisibleAsync())
        {
            LogStatus("Found autocomplete suggestion — clicking.");
            await suggestion.ClickAsync();
        }
        else
        {
            // JS fallback: click a list item containing the programme name
            var jsClicked = await _page.EvaluateAsync<bool>($@"() => {{
                const items = document.querySelectorAll('li, .autocomplete-option, .dropdown-item');
                for (const item of items) {{
                    if (item.textContent && item.textContent.includes('{fullProgrammeName[..Math.Min(10, fullProgrammeName.Length)]}')) {{
                        item.click(); return true;
                    }}
                }}
                return false;
            }}");
            LogStatus(jsClicked ? "Suggestion clicked via JS." : "No suggestion found — pressing Enter.");
            if (!jsClicked) await progCodeInput.PressAsync("Enter");
        }

        await Task.Delay(500);

        LogStatus("Clicking Run Report...");
        var runBtn = _page.Locator("input[value='Run Report'], button:has-text('Run Report')").First;
        await runBtn.ClickAsync();

        LogStatus("Waiting for report generation (may take up to 3 min)...");
        await _page.WaitForSelectorAsync("text=Completed", new PageWaitForSelectorOptions { Timeout = 180000 });
        LogStatus("Report generated. Starting download...");

        Directory.CreateDirectory(downloadDir);
        var savePath = Path.Combine(downloadDir, $"DeptAppReport_{DateTime.Now:yyyyMMddHHmmss}.csv");

        var download = await _page.RunAndWaitForDownloadAsync(async () =>
        {
            var okBtn = _page.Locator("button:has-text('Ok'), button:has-text('OK')").First;
            await okBtn.ClickAsync();
        }, new PageRunAndWaitForDownloadOptions { Timeout = 30000 });

        await download.SaveAsAsync(savePath);
        LogStatus($"Report downloaded: {Path.GetFileName(savePath)}");
        return savePath;
    }

    /// <summary>
    /// Individual Student Overview (ISO) flow:
    /// Awards, Assessments and Achievements → Data Quality Reports → Individual Student Overview.
    /// Enters the student code, picks the autocomplete match and programme route, retrieves the
    /// record, then downloads the "Download as CSV" export. Returns the saved CSV path.
    /// </summary>
    public async Task<string> DownloadIndividualStudentOverviewCsvAsync(string studentNumber, string downloadDir)
    {
        if (_page == null) throw new InvalidOperationException("Service not initialised.");

        studentNumber = studentNumber.Trim();
        LogStatus($"=== ISO: {studentNumber} ===");

        // Step 1: Open the "Awards, Assessments and Achievements" menu (left-nav anchor)
        LogStatus("[Step 1] Opening 'Awards, Assessments and Achievements'...");
        var awardsLink = _page.Locator("a").Filter(
            new() { HasText = "Awards, Assessments and Achievements" }).First;
        if (!await awardsLink.IsVisibleAsync())
        {
            LogStatus("[Step 1] Menu not visible — navigating to portal home first...");
            await _page.GotoAsync(_config?.PorticoUrl ?? "https://evision.ucl.ac.uk/urd/sits.urd/run/siw_lgn");
            await _page.WaitForSelectorAsync("text=My Portico", new PageWaitForSelectorOptions { Timeout = 30000 });
            awardsLink = _page.Locator("a").Filter(
                new() { HasText = "Awards, Assessments and Achievements" }).First;
        }
        if (!await awardsLink.IsVisibleAsync())
            awardsLink = _page.Locator("a, span, div, li").Filter(
                new() { HasText = "Awards, Assessments and Achievements" }).First;
        await awardsLink.ClickAsync();
        await _page.WaitForLoadStateAsync(LoadState.NetworkIdle, new PageWaitForLoadStateOptions { Timeout = 30000 });
        await Task.Delay(2000);

        // Step 2: Click the "Individual Student Overview" report link.
        // The link lives in the "Data Quality Reports" container on the Awards portal page,
        // which can take a moment to render — poll, normalising whitespace and preferring anchors.
        LogStatus("[Step 2] Locating 'Individual Student Overview' link...");
        string clickInfo = "NOT_FOUND";
        for (int attempt = 0; attempt < 10; attempt++)
        {
            clickInfo = await _page.EvaluateAsync<string>(@"() => {
                const norm = s => (s || '').replace(/\s+/g, ' ').trim().toLowerCase();
                const target = 'individual student overview';
                // Pass 1: exact match, anchors/buttons first (real clickable controls)
                const clickable = Array.from(document.querySelectorAll('a, button, input[type=button], input[type=submit]'));
                for (const el of clickable) {
                    const t = norm(el.textContent) || norm(el.value);
                    if (t === target) { el.scrollIntoView(); el.click(); return 'CLICKED_ANCHOR'; }
                }
                // Pass 2: any element whose *own* text is exactly the target
                const all = Array.from(document.querySelectorAll('span, div, li, td'));
                for (const el of all) {
                    if (norm(el.textContent) === target && el.children.length === 0) {
                        el.scrollIntoView(); el.click(); return 'CLICKED_TEXT';
                    }
                }
                // Pass 3: anchor whose text merely contains the target
                for (const el of clickable) {
                    if (norm(el.textContent).includes(target)) { el.scrollIntoView(); el.click(); return 'CLICKED_CONTAINS'; }
                }
                return 'NOT_FOUND';
            }");
            if (clickInfo.StartsWith("CLICKED")) break;
            await Task.Delay(1500);
        }

        if (!clickInfo.StartsWith("CLICKED"))
        {
            // Diagnostics: dump the link/report texts actually present so we can see what to match.
            var linkDump = await _page.EvaluateAsync<string>(@"() => {
                const out = [];
                document.querySelectorAll('a, button, input[type=button], input[type=submit]').forEach(el => {
                    const t = ((el.textContent || '') + (el.value || '')).replace(/\s+/g, ' ').trim();
                    if (t) out.push(t);
                });
                return out.slice(0, 150).join('  |  ');
            }");
            LogStatus("[Step 2] Clickable items on page: " + linkDump);
            throw new Exception("Could not find the 'Individual Student Overview' report link. See log for the links found on the page.");
        }

        LogStatus($"[Step 2] {clickInfo}.");
        await _page.WaitForLoadStateAsync(LoadState.NetworkIdle, new PageWaitForLoadStateOptions { Timeout = 30000 });
        await Task.Delay(1500);

        // Step 3: Enter the Student Code and select the AJAX autocomplete match.
        // The form is a SITS "ttq" form — the code box triggers an AJAX lookup whose selection
        // populates the Programme Route <select data-ttq-field="SPRCode">. Only a *real* mouse
        // click on the suggestion row fires that lookup (a synthetic click / keyboard does not),
        // so click the row and confirm SPRCode fills as proof the student was selected.
        LogStatus("[Step 3] Entering student code...");
        await _page.WaitForSelectorAsync("input[type='text']", new PageWaitForSelectorOptions { Timeout = 20000 });
        var codeInput = _page.Locator("input[type='text']:visible").First;
        if (!await codeInput.IsVisibleAsync())
            codeInput = _page.Locator("input[type='text']").First;

        await codeInput.ClickAsync();
        await codeInput.ClearAsync();
        await _page.Keyboard.TypeAsync(studentNumber, new KeyboardTypeOptions { Delay = 80 });
        await Task.Delay(1500);

        async Task<bool> RouteReadyAsync() => await _page!.EvaluateAsync<bool>(@"() => {
            const sel = document.querySelector('[data-ttq-field=""SPRCode""]')
                     || Array.from(document.querySelectorAll('select')).find(s => s.offsetParent !== null);
            return !!(sel && sel.options && sel.options.length > 0 && (sel.value || '').trim() !== '');
        }");

        // Real mouse click on the dropdown row that starts with the student number.
        var suggestion = _page.Locator("li, tr, td, div, a")
            .Filter(new() { HasTextRegex = new System.Text.RegularExpressions.Regex($@"^\s*{System.Text.RegularExpressions.Regex.Escape(studentNumber)}\b") })
            .Last;
        try
        {
            if (await suggestion.IsVisibleAsync())
                await suggestion.ClickAsync();
        }
        catch { /* fall back to keyboard below */ }

        bool routeReady = false;
        for (int i = 0; i < 12; i++)
        {
            if (await RouteReadyAsync()) { routeReady = true; break; }
            await Task.Delay(500);
        }

        if (!routeReady)
        {
            LogStatus("[Step 3] Suggestion click didn't populate route — trying ArrowDown → Enter...");
            await codeInput.PressAsync("ArrowDown");
            await Task.Delay(400);
            await codeInput.PressAsync("Enter");
            for (int i = 0; i < 10; i++)
            {
                if (await RouteReadyAsync()) { routeReady = true; break; }
                await Task.Delay(500);
            }
        }

        LogStatus(routeReady
            ? "[Step 3] Student selected — programme route populated."
            : "[Step 3] WARNING: programme route still empty; continuing anyway.");

        // Step 4: Ensure the Programme Route has a value (auto-fills on selection; force first real option if not)
        var routeResult = await _page.EvaluateAsync<string>(@"() => {
            const sel = document.querySelector('[data-ttq-field=""SPRCode""]')
                     || Array.from(document.querySelectorAll('select')).find(s => s.offsetParent !== null);
            if (!sel || sel.options.length === 0) return 'NO_SELECT';
            if ((sel.value || '').trim() !== '') return 'ALREADY: ' + sel.options[sel.selectedIndex].text;
            for (let i = 0; i < sel.options.length; i++) {
                const v = (sel.options[i].value || '').trim();
                const txt = (sel.options[i].text || '').trim().toLowerCase();
                if (v !== '' && !txt.includes('select') && !txt.includes('please')) {
                    sel.selectedIndex = i; sel.value = sel.options[i].value;
                    sel.dispatchEvent(new Event('change', { bubbles: true }));
                    return 'SELECTED: ' + sel.options[i].text;
                }
            }
            return 'NO_SELECT';
        }");
        LogStatus($"[Step 4] Programme route: {routeResult}");
        await Task.Delay(500);

        // Step 5: Click the "Retrieve student information" button via its ttq field hook
        LogStatus("[Step 5] Clicking 'Retrieve student information'...");
        var retrieveBtn = _page.Locator("[data-ttq-field='retrieve']").First;
        if (await retrieveBtn.CountAsync() == 0 || !await retrieveBtn.IsVisibleAsync())
            retrieveBtn = _page.Locator("input[value*='Retrieve'], button:has-text('Retrieve'), a:has-text('Retrieve')").First;
        if (await retrieveBtn.CountAsync() == 0)
            throw new Exception("Could not find the 'Retrieve student information' button.");
        await retrieveBtn.ClickAsync();

        // Wait for the results page: poll for the green "Download as CSV" control (up to 40s)
        LogStatus("[Step 5] Waiting for results to load...");
        bool resultsReady = false;
        for (int i = 0; i < 40; i++)
        {
            var has = await _page.EvaluateAsync<bool>(@"() => {
                const norm = s => (s || '').replace(/\s+/g, ' ').trim().toLowerCase();
                const els = document.querySelectorAll('a, button, input');
                for (const el of els) {
                    if ((norm(el.value) + ' ' + norm(el.textContent)).includes('download as csv')) return true;
                }
                return false;
            }");
            if (has) { resultsReady = true; break; }
            await Task.Delay(1000);
        }
        if (!resultsReady)
            throw new Exception("Results page did not load — no 'Download as CSV' button appeared after Retrieve.");
        LogStatus("[Step 5] Results page loaded.");
        await Task.Delay(500);

        // Step 6: Scrape the student name from the "Student Details [STU]" table
        LogStatus("[Step 6] Reading student name...");
        var scrapedName = await _page.EvaluateAsync<string>(@"() => {
            // SITS responsive tables embed the column header inside each data cell (a visually
            // hidden span), so cell.textContent reads like 'SurnameYang'. Strip that header prefix.
            const norm = s => (s || '').replace(/\s+/g, ' ').trim();
            const strip = (val, header) => {
                let v = norm(val), h = norm(header);
                if (h && v.toLowerCase().startsWith(h.toLowerCase())) v = v.slice(h.length).trim();
                return v;
            };
            const tables = document.querySelectorAll('table');
            for (const table of tables) {
                const rows = table.querySelectorAll('tr');
                let surnameCol = -1, firstCol = -1, headerRow = -1, surnameHdr = '', firstHdr = '';
                for (let r = 0; r < rows.length; r++) {
                    const cells = rows[r].querySelectorAll('th, td');
                    for (let c = 0; c < cells.length; c++) {
                        const raw = norm(cells[c].textContent);
                        const h = raw.toLowerCase();
                        if (h === 'surname') { surnameCol = c; surnameHdr = raw; }
                        if (h === 'first names' || h === 'first name' || h === 'forename' || h === 'forenames') { firstCol = c; firstHdr = raw; }
                    }
                    if (surnameCol >= 0 && firstCol >= 0) { headerRow = r; break; }
                }
                if (headerRow >= 0) {
                    for (let r = headerRow + 1; r < rows.length; r++) {
                        const cells = rows[r].querySelectorAll('td, th');
                        if (cells.length > Math.max(surnameCol, firstCol)) {
                            const surname = strip(cells[surnameCol].textContent, surnameHdr);
                            const first = strip(cells[firstCol].textContent, firstHdr);
                            if (surname || first) return first + '|' + surname;
                        }
                    }
                }
            }
            return '';
        }");

        string firstNames = "", surname = "";
        if (!string.IsNullOrWhiteSpace(scrapedName) && scrapedName.Contains('|'))
        {
            var parts = scrapedName.Split('|', 2);
            firstNames = parts[0].Trim();
            surname = parts[1].Trim();
            LogStatus($"[Step 6] Student: {firstNames} {surname}");
        }
        else
        {
            LogStatus("[Step 6] WARNING: Could not read student name — using student number only.");
        }

        // Build "{studentNumber} {FirstNames} {Surname} - ISO.csv" (or fall back to number only)
        var namePart = string.Join(" ", new[] { firstNames, surname }
            .Where(s => !string.IsNullOrWhiteSpace(s))).Trim();
        var baseName = string.IsNullOrWhiteSpace(namePart)
            ? $"{studentNumber} - ISO"
            : $"{studentNumber} {namePart} - ISO";
        var invalid = Path.GetInvalidFileNameChars();
        var safeName = new string(baseName.Select(c => invalid.Contains(c) ? '_' : c).ToArray());
        var fileName = safeName + ".csv";

        // Step 7: Click the green "Download as CSV" button with a REAL mouse click (a synthetic
        // el.click() is ignored by this SITS UI). The export may either stream a download on the
        // current page OR open the CSV in a new tab/popup — arm waiters for both before clicking.
        LogStatus("[Step 7] Locating 'Download as CSV'...");
        var csvButton = _page.GetByText("Download as CSV", new() { Exact = false }).First;
        if (await csvButton.CountAsync() == 0)
            csvButton = _page.Locator("input[value*='Download as CSV'], a:has-text('Download'), button:has-text('Download')").First;
        if (await csvButton.CountAsync() == 0)
            throw new Exception("'Download as CSV' button not found on the results page.");

        Directory.CreateDirectory(downloadDir);
        LogStatus("[Step 7] Clicking 'Download as CSV'...");

        // Arm both waiters BEFORE the click so the event can't be missed.
        var pageDownloadTask = _page.WaitForDownloadAsync(new PageWaitForDownloadOptions { Timeout = 60000 });
        var popupTask = _page.WaitForPopupAsync(new PageWaitForPopupOptions { Timeout = 8000 });

        await csvButton.ScrollIntoViewIfNeededAsync();
        await csvButton.ClickAsync();

        IPage? popup = null;
        try { popup = await popupTask; } catch { /* no popup opened */ }

        IDownload download;
        if (popup != null)
        {
            LogStatus("[Step 7] CSV opened in a new tab — capturing its download...");
            try
            {
                download = await popup.WaitForDownloadAsync(new PageWaitForDownloadOptions { Timeout = 60000 });
            }
            catch
            {
                // The popup may already have streamed the file to the main page instead.
                download = await pageDownloadTask;
            }
        }
        else
        {
            download = await pageDownloadTask;
        }

        var savePath = Path.Combine(downloadDir, fileName);
        await download.SaveAsAsync(savePath);
        LogStatus($"=== ISO SAVED: {savePath} ===");
        return savePath;
    }

    public async Task CloseAsync(bool handOffToUser = false)
    {
        LogStatus("Closing browser...");
        if (_context != null)
        {
            try { await _context.CloseAsync(); } catch { }
            _context = null;
        }
        _playwright?.Dispose();
        _playwright = null;
        _page = null;

        if (handOffToUser && !string.IsNullOrEmpty(_userDataDir))
        {
            try
            {
                await Task.Delay(800); // let Edge fully exit and release the profile lock
                var pid = LaunchHandoffBrowser(_userDataDir, _config?.PorticoUrl);
                if (pid > 0)
                    File.WriteAllText(Path.Combine(_userDataDir, HandoffPidFileName), pid.ToString());
                LogStatus("Edge reopened for normal use — your Portico session is still signed in.");
            }
            catch (Exception ex)
            {
                LogStatus($"Could not reopen Edge for normal use: {ex.Message}");
            }
        }
    }

    private static int LaunchHandoffBrowser(string userDataDir, string? url)
    {
        ProcessStartInfo psi;
        bool pidIsBrowser = true;

        if (OperatingSystem.IsMacOS())
        {
            var edgeBinary = "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge";
            if (File.Exists(edgeBinary))
            {
                psi = new ProcessStartInfo(edgeBinary);
            }
            else
            {
                // 'open' exits immediately, so the browser PID can't be tracked
                psi = new ProcessStartInfo("open");
                psi.ArgumentList.Add("-na");
                psi.ArgumentList.Add("Microsoft Edge");
                psi.ArgumentList.Add("--args");
                pidIsBrowser = false;
            }
        }
        else if (OperatingSystem.IsWindows())
        {
            var candidates = new[]
            {
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "Microsoft", "Edge", "Application", "msedge.exe"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Microsoft", "Edge", "Application", "msedge.exe"),
            };
            psi = new ProcessStartInfo(candidates.FirstOrDefault(File.Exists) ?? "msedge.exe");
        }
        else
        {
            psi = new ProcessStartInfo("microsoft-edge");
        }

        psi.ArgumentList.Add($"--user-data-dir={userDataDir}");
        if (!string.IsNullOrEmpty(url))
            psi.ArgumentList.Add(url);
        psi.UseShellExecute = false;

        var proc = Process.Start(psi);
        return pidIsBrowser && proc != null ? proc.Id : 0;
    }

    private async Task CloseStaleHandoffBrowserAsync(string userDataDir)
    {
        var pidFile = Path.Combine(userDataDir, HandoffPidFileName);
        if (!File.Exists(pidFile)) return;

        try
        {
            if (int.TryParse(File.ReadAllText(pidFile).Trim(), out var pid))
            {
                var proc = Process.GetProcessById(pid);
                // PID may have been recycled by the OS — only touch it if it's actually Edge
                if (!proc.ProcessName.Contains("edge", StringComparison.OrdinalIgnoreCase))
                    return;

                LogStatus("Closing the hand-off Edge window so automation can use the profile...");
                if (OperatingSystem.IsWindows())
                {
                    proc.CloseMainWindow();
                }
                else
                {
                    var kill = Process.Start("kill", $"-TERM {pid}");
                    kill?.WaitForExit(2000);
                }

                if (!proc.WaitForExit(5000))
                    proc.Kill(entireProcessTree: true);

                await Task.Delay(1000); // let the profile lock release
            }
        }
        catch (ArgumentException)
        {
            // Process already gone — nothing to close.
        }
        catch (Exception ex)
        {
            LogStatus($"Note: could not close previous Edge window: {ex.Message}");
        }
        finally
        {
            try { File.Delete(pidFile); } catch { }
        }
    }

    /// <summary>
    /// Waits for a page reload by polling a JS marker that gets cleared on navigation.
    /// Call EvaluateAsync("() => window.__pw_marker = true") BEFORE the action that triggers reload.
    /// </summary>
    private async Task WaitForPageReloadAsync(int timeoutMs = 120000)
    {
        var pollMs = 1000;
        var elapsed = 0;

        while (elapsed < timeoutMs)
        {
            try
            {
                // If we can evaluate JS and the marker is gone, the page has fully reloaded
                var markerExists = await _page!.EvaluateAsync<bool>("() => window.__pw_marker === true");
                if (!markerExists)
                {
                    return;
                }
            }
            catch
            {
                // JS evaluation failed — page is mid-navigation, keep polling
            }

            await Task.Delay(pollMs);
            elapsed += pollMs;
        }

        throw new Exception($"Page did not reload within {timeoutMs / 1000} seconds.");
    }

    private void LogStatus(string message) => StatusUpdated?.Invoke(this, message);
}
