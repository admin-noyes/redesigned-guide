````markdown
# URL Validator Excel Add-In Documentation

## Overview
The URL Validator Add-In is an Excel tool that scans the active worksheet for URL-like values (in any column, starting from row 2), checks each URL via HTTP HEAD requests, logs failures, and generates a dashboard with statistics and a pie chart. The updated version uses PowerShell for multi-threaded parallel processing, enabling efficient handling of large datasets (up to 100,000+ URLs).

## Features
- Validates URLs by checking HTTP status codes
- Retries failed requests up to 3 times with delays
- Logs failed and invalid URLs to a "URL_Log" sheet (records source cell address, original URL, and status)
- Creates a "URL_Dashboard" sheet with summary statistics and pie chart
- Multi-threaded processing using PowerShell (up to 50 concurrent threads)
- Optimized for large files with reduced timeouts and runspace parallelism
- Writes status next to each source cell (one column to the right)

## Installation Instructions

### Step 1: Package the VBA Code as an Add-In
1. Open Microsoft Excel.
2. Create a new workbook (Ctrl+N).
3. Press Alt+F11 to open the Visual Basic for Applications (VBA) editor.
4. In the VBA editor:
   - Insert a new module: **Insert > Module**
   - Copy and paste the code from `modURLChecker` section of `url-val.vba` into the module.
   - Repeat for `modDashboard` and `modProgress` modules.
   - For the UserForm: **Insert > UserForm**
     - Name it `frmProgress`.
     - Add a Label named `lblProgress` (for progress text).
     - Add a Frame named `Frame1`.
     - Inside the Frame, add a Label named `bar` (this will be the progress bar - set BackColor to blue or preferred color).
     - Copy the UserForm code from `url-val.vba` into the form's code window.
   - For ThisWorkbook: In the Project Explorer, double-click "ThisWorkbook" and paste the code.
5. Save the workbook as an Excel Add-In:
   - **File > Save As**
   - Choose "Excel Add-In (*.xlam)" from the file type dropdown.
   - Name it `URL_Validator.xlam` and save to a convenient location (e.g., Desktop or Documents folder).
6. Place `URLChecker.ps1` in the same directory as `URL_Validator.xlam` (the VBA module launches PowerShell using the add-in path).

### Step 2: Install the Add-In
1. Open Excel (if not already open).
2. Go to **File > Options > Add-Ins**.
3. At the bottom, select "Excel Add-ins" from the Manage dropdown and click "Go".
4. In the Add-Ins dialog, click "Browse".
5. Navigate to and select the `URL_Validator.xlam` file you created.
6. Click "OK" to install.
7. Ensure "URL Validator" is checked in the list and click "OK".

The add-in is now installed and will load automatically with Excel.

## Usage Instructions

### Preparing Your Data
1. Open a new or existing Excel workbook.
2. In the active worksheet, enter URLs anywhere in the sheet starting from row 2 (row 1 is treated as header row). The add-in will detect URL-like values (starting with `http`, containing `www.`, or containing `://`) in any column and check them.
   - The status for each URL will be written to the cell one column to the right of the source URL cell.
   - Example:
     - Put `https://www.example.com` in `B2` -> status written to `C2`.

### Running the URL Check
1. Press Alt+F8 to open the Macro dialog.
2. Select "RunURLCheck" from the list (macro description: "Check URLs in the sheet and generate dashboard report").
3. Click "Run".

### What Happens During Execution
- The add-in collects all URL-like cells and invokes PowerShell for parallel checking. Each URL is identified by its cell address so duplicate URLs are handled per-cell.
- Status codes (e.g., "200" for success) will be written to the cell one column to the right of the URL.
- Failed URLs will be logged in the "URL_Log" worksheet; the log stores the source cell address, original URL, and status.
- A "URL_Dashboard" worksheet will be created/updated with:
  - Total URLs checked
  - Successful checks (200 status)
  - Failures
  - Invalid URLs
  - Failure percentage
  - Pie chart visualization

### After Completion
- A message box will confirm "URL Check Complete!".
- Review results in the worksheet (statuses next to source cells), `URL_Log`, and `URL_Dashboard`.

## Performance Tuning
- For large files (50,000+), adjust `MaxThreads` in `URLChecker.ps1` (default 50; reduce if network throttles).
- Ensure PowerShell 5.1+; upgrade to PowerShell 7+ for better performance.
- Monitor system resources; use high-speed internet for optimal speed.

## Troubleshooting
- **Macro not appearing**: Ensure add-in is installed and `URLChecker.ps1` is in the same folder.
- **PowerShell missing or blocked**: Ensure PowerShell is available and script execution is permitted. To allow scripts for the current user run in an elevated (or appropriate) PowerShell prompt:

```
Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
```

- **PowerShell errors**: Check execution policy and the path to `URLChecker.ps1`.
- **Unsaved Add-In path**: If you edit VBA inside a workbook and haven't saved the add-in, `ThisWorkbook.Path` may be empty; save the workbook/add-in to a folder where `URLChecker.ps1` is located.
- **Slow performance**: Reduce threads or check network speed.
- **VBProject access error**: Enable "Trust access to the VBA project object model" in Excel Trust Center.
- **Large datasets**: Split into batches if memory issues occur.

## Code Review Summary
The updated VBA integrates PowerShell runspaces for parallelism, reduces per-URL overhead by batching via CSV input/output, and maps results by source cell address. This improves robustness when duplicate URLs exist in multiple cells and enables multi-column scanning. Suitable for production use on large datasets when deployed with the accompanying `URLChecker.ps1`.
````
