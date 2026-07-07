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
