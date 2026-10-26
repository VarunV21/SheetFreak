Attribute VB_Name = "SheetFreakTests"
Option Explicit
Option Private Module
Private mLog As String

Private Sub Check(ByVal condition As Boolean, ByVal label As String)
    If Not condition Then Err.Raise 513, "SheetFreakTests", "FAIL: " & label
    mLog = mLog & "PASS: " & label & vbCrLf
End Sub

Public Sub RunAll()
    Dim book As Workbook, sheet As Worksheet, i As Long, savedCalc As XlCalculation
    Dim calc As XlCalculation, events As Boolean, screen As Boolean, rejected As Boolean
    Dim oldBooks As Long, report As Integer, path As String
    On Error GoTo Failed
    mLog = "SheetFreak native Excel tests - " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf
    Check SF_ZeroFormat("#,##0.00") = "#,##0", "comma format"
    Check SF_ZeroFormat("$#,##0.00;[Red]($#,##0.00)") = "$#,##0;[Red]($#,##0)", "currency, red negatives, sections"
    Check SF_ZeroFormat("0.00%") = "0%", "percentage"
    Check SF_ZeroFormat("0.00E+00") = "0E+00", "scientific"
    Check SF_ZeroFormat("General") = "0", "General to zero decimals"
    Check SF_ZeroFormat("mm/dd/yyyy") = "mm/dd/yyyy", "date untouched"
    Check SF_ZeroFormat("hh:mm:ss.00") = "hh:mm:ss.00", "time untouched"
    Check SF_ZeroFormat("[h]:mm:ss.00") = "[h]:mm:ss.00", "elapsed time untouched"
    Check SF_ZeroFormat("# ?/?") = "# ?/?", "fraction untouched"
    Check SF_ZeroFormat("0.00 ""version 1.00""") = "0 ""version 1.00""", "quoted decimal literal preserved"
    Check SF_ZeroFormat(".00") = "0", "missing integer placeholder"
    Check SF_ZeroFormat("@") = "@", "text format untouched"
    Check SF_ZeroFormat("[>=1000]#,##0.00;0.000") = "[>=1000]#,##0;0", "conditions preserved"
    Set book = Workbooks.Add(xlWBATWorksheet): Set sheet = book.Worksheets(1)
    sheet.Name = "SheetFreak test"
    sheet.Range("A1").Value2 = 1234.567
    sheet.Range("A1").NumberFormat = "#,##0.00"
    sheet.Range("A2").Value2 = 0.12345
    sheet.Range("A2").NumberFormat = "0.00%"
    sheet.Range("A3").Value2 = 45000
    sheet.Range("A3").NumberFormat = "mm/dd/yyyy"
    sheet.Range("A1:A3").Select: SF_ZeroDecimals
    Check sheet.Range("A1").NumberFormat = "#,##0", "mixed-selection comma format"
    Check sheet.Range("A2").NumberFormat = "0%", "mixed-selection percentage"
    Check sheet.Range("A3").NumberFormat = "mm/dd/yyyy", "mixed-selection date preserved"
    Check sheet.Range("A1").Value2 = 1234.567, "underlying precision preserved"
    sheet.Range("B1:C2").Select: SF_Header
    Check sheet.Range("B1").Interior.Color = RGB(64, 64, 64), "header gray"
    Check sheet.Range("B1").Font.Color = vbWhite And sheet.Range("B1").Font.Bold, "header font"
    Check sheet.Range("B1:C2").Borders(xlInsideVertical).Color = vbWhite, "white inside vertical borders"
    Check sheet.Range("B1:C2").Borders(xlInsideHorizontal).Color = vbWhite, "white inside horizontal borders"
    sheet.Range("D1").Select: SF_Assumption
    Check sheet.Range("D1").Font.Color = RGB(0, 0, 255), "assumption blue"
    Check sheet.Range("D1").Interior.Color = RGB(255, 242, 204), "assumption yellow"
    sheet.Range("E1").Formula2 = "=A1+$A$2"
    sheet.Range("F1").Value2 = "'=literal"
    sheet.Range("E1:F1").Select: SF_CopyExact
    sheet.Range("H5").Select: SF_PasteExact
    Check sheet.Range("H5").Formula2 = "=A1+$A$2", "exact relative and absolute formula text"
    Check sheet.Range("I5").Value2 = "=literal" And Not sheet.Range("I5").HasFormula, "literal equals text remains text"
    sheet.Range("E1").Select: SF_CopyExact
    sheet.Range("H6").Select: SF_PasteExact
    Check sheet.Range("H6").Formula2 = "=A1+$A$2", "single-cell exact copy"
    sheet.Range("Z100").Select: SF_Home
    Check Selection.Address = "$A$1", "home selects A1"
    Check ActiveWindow.Zoom >= 70 And ActiveWindow.Zoom <= 150, "home readable zoom bounds"
    Check Not ActiveWindow.DisplayGridlines, "home hides gridlines"
    sheet.Range("A1:Z100").Clear
    sheet.Range("A1").Value2 = "Key"
    sheet.Range("A2:A100002").FormulaR1C1 = "=ROW()"
    sheet.Range("A2:A100002").Value2 = sheet.Range("A2:A100002").Value2
    sheet.Range("B2").Formula2 = "=A2*2"
    sheet.Range("C2").Formula2 = "=B2+1"
    sheet.Range("B2").NumberFormat = "#,##0.00"
    sheet.Range("C2").NumberFormat = "0%"
    sheet.Range("B3:C4").Value2 = "Original sentinel"
    savedCalc = Application.Calculation
    Application.Calculation = xlCalculationManual
    calc = Application.Calculation: events = Application.EnableEvents: screen = Application.ScreenUpdating
    oldBooks = Workbooks.Count
    SF_ExecuteFill sheet.Range("B2:C2"), sheet.Range("A1")
    Check Workbooks.Count = oldBooks, "fill creates no additional workbook"
    Check sheet.Range("B2").HasFormula And sheet.Range("C2").HasFormula, "starting formulas retained"
    Check Not sheet.Range("B3:C100002").HasFormula, "100000 lower rows frozen"
    Check sheet.Range("B100002").Value2 = 200004 And sheet.Range("C100002").Value2 = 200005, "cross-column dependencies calculate correctly"
    Check sheet.Range("B50001").Value2 = 100002, "middle-of-range result"
    Check sheet.Range("C100002").NumberFormat = "0%", "column-specific number formats"
    Check sheet.Range("B100002").Font.Color = vbBlack, "lower rows black"
    Check sheet.Range("B2:C2").Font.Color = RGB(112, 48, 160), "seed formulas purple"
    Check Application.Calculation = calc And Application.EnableEvents = events And Application.ScreenUpdating = screen, "application state restored after success"
    sheet.Range("B2").Value2 = 7
    On Error Resume Next
    SF_ValidateFill sheet.Range("B2:C2"), sheet.Range("A1")
    rejected = Err.Number <> 0: Err.Clear
    On Error GoTo Failed
    Check rejected, "constant seed rejected"
    sheet.Range("B2").Formula2 = "=A2*2"
    sheet.Protect
    rejected = False
    On Error Resume Next
    SF_ValidateFill sheet.Range("B2:C2"), sheet.Range("A1")
    rejected = Err.Number <> 0: Err.Clear
    On Error GoTo Failed
    Check rejected, "protected sheet rejected"
    sheet.Unprotect
    sheet.Range("A1:C100002").AutoFilter Field:=1, Criteria1:=">10"
    rejected = False
    On Error Resume Next
    SF_ValidateFill sheet.Range("B2:C2"), sheet.Range("A1")
    rejected = Err.Number <> 0: Err.Clear
    On Error GoTo Failed
    Check rejected, "active filter rejected"
    sheet.AutoFilterMode = False
    sheet.Range("B2").Formula2 = "=1/0"
    rejected = False
    On Error Resume Next
    SF_ExecuteFill sheet.Range("B2:C2"), sheet.Range("A1")
    rejected = Err.Number <> 0: Err.Clear
    On Error GoTo Failed
    Check rejected, "calculation errors stop freezing"
    Check sheet.Range("B100002").HasFormula, "error range remains formulas"
    Check Workbooks.Count = oldBooks, "error creates no additional workbook"
    Check Application.Calculation = calc And Application.EnableEvents = events And Application.ScreenUpdating = screen, "application state restored after error"
    Application.Calculation = savedCalc
    mLog = mLog & "ALL TESTS PASSED" & vbCrLf
    GoTo WriteReport
Failed:
    mLog = mLog & "FAIL: " & Err.Description & vbCrLf
WriteReport:
    path = ThisWorkbook.Path & Application.PathSeparator & "native-test-results.txt"
    report = FreeFile
    Open path For Output As #report
    Print #report, mLog
    Close #report
    MsgBox mLog, vbInformation, "SheetFreak native test report"
End Sub
