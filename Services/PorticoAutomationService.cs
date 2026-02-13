// Services/PorticoAutomationService.cs

using Microsoft.Playwright;
using Dossier.Models;

namespace Dossier.Services;

public class PorticoAutomationService : IPorticoAutomationService
{
    private IPlaywright? _playwright;
    private IBrowser? _browser;
    private IBrowserContext? _context;
    private IPage? _page;
    private AppConfig? _config;

    public bool DebugMode { get; set; } = false;

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
        LogStatus("Initialising Playwright...");
        _playwright = await Playwright.CreateAsync();

        var userDataDir = config.EdgeUserDataDir;
        if (string.IsNullOrEmpty(userDataDir))
        {
            var appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            userDataDir = Path.Combine(appDataPath, "Dossier", "EdgeProfile");
            Directory.CreateDirectory(userDataDir);
        }

        var contextOptions = new BrowserTypeLaunchPersistentContextOptions
        {
            Headless = config.HeadlessMode,
            Channel = "msedge",
            SlowMo = config.ActionDelayMs,
            AcceptDownloads = true,
            ViewportSize = null,
            Args = new[] { "--start-maximized" }
        };

        _context = await _playwright.Chromium.LaunchPersistentContextAsync(userDataDir, contextOptions);
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

    public async Task CloseAsync()
    {
        LogStatus("Closing browser...");
        if (_context != null) 
        { 
            await _context.CloseAsync(); 
            _context = null; 
        }
        _playwright?.Dispose(); 
        _playwright = null;
        _page = null;
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
