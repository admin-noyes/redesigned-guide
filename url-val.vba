' modURLChecker

Option Explicit

Public Sub RunURLCheck()

    Dim ws As Worksheet
    Dim logWs As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim url As String
    Dim result As String
    Dim logRow As Long
    Dim urls() As String
    Dim results As Dictionary
    Dim tempInput As String
    Dim tempOutput As String
    Dim psCommand As String
    
    Set ws = ActiveSheet
    
    ' Create / Get Log Sheet
    On Error Resume Next
    Set logWs = Worksheets("URL_Log")
    If logWs Is Nothing Then
        Set logWs = Worksheets.Add
        logWs.Name = "URL_Log"
        logWs.Range("A1:C1").Value = Array("Row", "URL", "Status")
    End If
    On Error GoTo 0
    
    logRow = logWs.Cells(logWs.Rows.Count, 1).End(xlUp).Row + 1
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Collect URLs
    ReDim urls(1 To lastRow - 1)
    For i = 2 To lastRow
        urls(i - 1) = ws.Cells(i, 1).Value
    Next i
    
    ' Performance optimization
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' Write URLs to temp file
    tempInput = Environ("TEMP") & "\urls.txt"
    tempOutput = Environ("TEMP") & "\results.csv"
    
    Open tempInput For Output As #1
    For i = LBound(urls) To UBound(urls)
        Print #1, urls(i)
    Next i
    Close #1
    
    ' Run PowerShell script for parallel checking
    psCommand = "powershell.exe -ExecutionPolicy Bypass -File """ & ThisWorkbook.Path & "\URLChecker.ps1"" -InputFile """ & tempInput & """ -OutputFile """ & tempOutput & """"
    Shell psCommand, vbNormalFocus
    
    ' Wait for completion (simple wait, in real scenario might poll file existence)
    Do While Dir(tempOutput) = ""
        DoEvents
        Application.Wait Now + TimeValue("00:00:01")
    Loop
    
    ' Read results
    Set results = New Dictionary
    Dim csvLine As String
    Open tempOutput For Input As #1
    Line Input #1, csvLine ' Skip header
    Do While Not EOF(1)
        Line Input #1, csvLine
        Dim parts() As String
        parts = Split(csvLine, ",")
        If UBound(parts) >= 1 Then
            results.Add Trim(Replace(parts(0), """", "")), Trim(Replace(parts(1), """", ""))
        End If
    Loop
    Close #1
    
    ' Process results
    For i = 2 To lastRow
        url = ws.Cells(i, 1).Value
        If results.Exists(url) Then
            result = results(url)
            ws.Cells(i, 2).Value = result
            If result <> "200" Then
                logWs.Cells(logRow, 1).Value = i
                logWs.Cells(logRow, 2).Value = url
                logWs.Cells(logRow, 3).Value = result
                logRow = logRow + 1
            End If
        End If
    Next i
    
    ' Clean up temp files
    Kill tempInput
    Kill tempOutput
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    BuildDashboard
    
    MsgBox "URL Check Complete!", vbInformation

End Sub


Public Function CheckURLStatus(url As String, Optional retries As Integer = 3) As String

    Dim request As Object
    Dim attempt As Integer
    Dim trimmedUrl As String
    
    trimmedUrl = Trim(url)
    
    If LCase(Left(trimmedUrl, 4)) <> "http" Then
        CheckURLStatus = "Invalid URL"
        Exit Function
    End If
    
    For attempt = 1 To retries
    
        On Error GoTo retryLogic
        
        Set request = CreateObject("WinHttp.WinHttpRequest.5.1")
        request.SetTimeouts 3000, 3000, 3000, 5000
        request.Open "HEAD", trimmedUrl, False
        request.Send
        
        CheckURLStatus = request.Status
        Exit Function
        
retryLogic:
        If attempt < retries Then
            Application.Wait Now + TimeValue("00:00:02")
        Else
            CheckURLStatus = "Failed After " & retries & " Attempts"
        End If
        
        Err.Clear
        
    Next attempt

End Function


















' modDashboard

Option Explicit

Public Sub BuildDashboard()

    Dim ws As Worksheet
    Dim dash As Worksheet
    Dim lastRow As Long
    Dim successCount As Long
    Dim failCount As Long
    Dim invalidCount As Long
    Dim i As Long
    Dim total As Long
    
    Set ws = ActiveSheet
    
    On Error Resume Next
    Set dash = Worksheets("URL_Dashboard")
    
    If dash Is Nothing Then
        Set dash = Worksheets.Add
        dash.Name = "URL_Dashboard"
    End If
    On Error GoTo 0
    
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    total = lastRow - 1
    
    For i = 2 To lastRow
    
        Select Case ws.Cells(i, 2).Value
    Dim lastCol As Long
    Dim r As Long, c As Long
    Dim startRow As Long

    startRow = 2 ' assume row 1 is header
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    ' Collect URLs by scanning all used columns starting at row 2
    Dim entryCount As Long
    entryCount = 0
    ' We'll write directly to tempInput as CSV with Address,URL
        
    Next i
    
    dash.Cells.Clear
    
    dash.Range("A1:B1").Value = Array("Metric", "Value")
    ' Write URLs to temp file (CSV: Address,URL)
    tempInput = Environ("TEMP") & "\urls_input.csv"
    ' Aggregate status values across the sheet (any column where results were written)
    Dim ur As Range
    Dim cell As Range
    Set ur = ws.UsedRange
    total = 0
    For Each cell In ur.Cells
        If Not IsEmpty(cell.Value) Then
            Dim v As String
            v = Trim(CStr(cell.Value))
            If v = "200" Or v = "Invalid URL" Or InStr(1, v, "Failed", vbTextCompare) > 0 Or InStr(1, v, "Error", vbTextCompare) > 0 Then
                total = total + 1
                Select Case True
                    Case v = "200"
                        successCount = successCount + 1
                    Case v = "Invalid URL"
                        invalidCount = invalidCount + 1
                    Case InStr(1, v, "Failed", vbTextCompare) > 0 Or InStr(1, v, "Error", vbTextCompare) > 0 Or (IsNumeric(v) And Val(v) >= 400)
                        failCount = failCount + 1
                End Select
            End If
        End If
    Next cell
    dash.Range("A2:B6").Value = Array( _
        Array("Total URLs", total), _
        Array("Successful (200)", successCount), _
        Array("Failures", failCount), _
        Array("Invalid URLs", invalidCount), _
        Array("Failure %", IIf(total > 0, Format((failCount / total), "0.00%"), "0.00%")) _

    Dim dash As Worksheet
    Set dash = Worksheets("URL_Dashboard")
    
    Dim chartObj As ChartObject
    Set chartObj = dash.ChartObjects.Add(Left:=250, Width:=400, Top:=50, Height:=300)
    ' Read results (CSV: Address,Status)
    Set results = CreateObject("Scripting.Dictionary")
    Dim csvLine As String
    Dim addr As String, stat As String
    Open tempOutput For Input As #1
    If Not EOF(1) Then Line Input #1, csvLine ' Skip header
    Do While Not EOF(1)
        Line Input #1, csvLine
        Dim parts() As String
        parts = Split(csvLine, ",")
        If UBound(parts) >= 1 Then
            addr = Trim(Replace(parts(0), """", """"))
            stat = Trim(Replace(parts(1), """", """"))
            If Not results.Exists(addr) Then results.Add addr, stat
        End If
    Loop
    Close #1


    ' Process results and write status to cell to the right of the source cell
    Dim srcCell As Range
    For Each addr In results.Keys
        On Error Resume Next
        Set srcCell = ws.Range(addr)
        On Error GoTo 0
        If Not srcCell Is Nothing Then
            result = results(addr)
            srcCell.Offset(0, 1).Value = result
            If result <> "200" Then
                logWs.Cells(logRow, 1).Value = srcCell.Address(False, False)
                logWs.Cells(logRow, 2).Value = srcCell.Value
                logWs.Cells(logRow, 3).Value = result
                logRow = logRow + 1
            End If
        End If
        Set srcCell = Nothing
    Next addr
    With frmProgress
        .lblProgress.Caption = "Processing " & current & " of " & total
        .bar.Width = (.Frame1.Width) * (current / total)
        DoEvents
    End With

End Sub

















' ThisWorkbook

Private Sub Workbook_Open()

    Application.MacroOptions _
        Macro:="RunURLCheck", _
        Description:="Check URLs in Column A and generate dashboard report"

End Sub

















' UserForm: frmProgress

Private Sub UserForm_Initialize()
    bar.Width = 0
End Sub