Attribute VB_Name = "WorkbookViewTests"
Option Explicit
Option Private Module

' Load the candidate SheetFreak.xlam, import this module into a disposable host,
' then call RunWorkbookViewChecks with an explicit report filename.
Public Sub RunWorkbookViewChecks(ByVal reportPath As String)
    Dim book As Workbook, other As Workbook, originalBook As Workbook
    Dim first As Worksheet, second As Worksheet, hidden As Worksheet
    Dim veryHidden As Worksheet, chart As Chart, sheet As Worksheet
    Dim originalSheet As Object, file As Integer, log As String
    Dim expected As Long, width As Double, oldEvents As Boolean, oldScreen As Boolean
    On Error GoTo Failed
    Set originalBook = ActiveWorkbook: Set originalSheet = ActiveSheet
    oldEvents = Application.EnableEvents: oldScreen = Application.ScreenUpdating
    Set book = Workbooks.Add(xlWBATWorksheet)
    Set first = book.Worksheets(1): first.Name = "Compact"
    first.Range("J17").Value2 = 5233
    first.Columns("A:J").ColumnWidth = 5
    Set second = book.Worksheets.Add(After:=first): second.Name = "Wide"
    second.Range("AZ17").Value2 = 42
    second.Columns("A:AZ").ColumnWidth = 12
    Set hidden = book.Worksheets.Add(After:=second): hidden.Name = "Hidden"
    ActiveWindow.Zoom = 88: ActiveWindow.DisplayGridlines = True
    hidden.Visible = xlSheetHidden
    Set veryHidden = book.Worksheets.Add(After:=second): veryHidden.Name = "Very hidden"
    ActiveWindow.Zoom = 91: ActiveWindow.DisplayGridlines = True
    veryHidden.Visible = xlSheetVeryHidden
    Set chart = book.Charts.Add: chart.Name = "Chart tab"
    Set other = Workbooks.Add(xlWBATWorksheet)
    ActiveWindow.Zoom = 93: ActiveWindow.DisplayGridlines = True
    book.Activate: second.Activate
    second.Range("A1:AZ1").Select: ActiveWindow.Zoom = True
    expected = Int(CDbl(ActiveWindow.Zoom) * 0.9)
    If expected < 70 Then expected = 70
    If expected > 150 Then expected = 150
    width = second.Columns("A:AZ").Width
    first.Activate
    Application.Run "'SheetFreak.xlam'!SF_ResizeAll"
    Check ActiveSheet Is first, "returns to original worksheet", log
    For Each sheet In ArraySheets(first, second)
        sheet.Activate
        Check Selection.Address = "$A$1", sheet.Name & " selects A1", log
        Check Not ActiveWindow.DisplayGridlines, sheet.Name & " gridlines off", log
        Check ActiveWindow.Zoom >= 70 And ActiveWindow.Zoom <= 150, sheet.Name & " zoom bounds", log
    Next sheet
    Check ActiveWindow.Zoom = expected, "wide sheet applies margin or readability floor", log
    Check second.Columns("A:AZ").Width = width, "wide column widths preserved", log
    Check first.Range("J17").Value2 = 5233 And second.Range("AZ17").Value2 = 42, "contents preserved", log
    Check hidden.Visible = xlSheetHidden And veryHidden.Visible = xlSheetVeryHidden, "hidden states preserved", log
    hidden.Visible = xlSheetVisible: hidden.Activate
    Check ActiveWindow.Zoom = 88 And ActiveWindow.DisplayGridlines, "hidden view untouched", log
    hidden.Visible = xlSheetHidden
    veryHidden.Visible = xlSheetVisible: veryHidden.Activate
    Check ActiveWindow.Zoom = 91 And ActiveWindow.DisplayGridlines, "very hidden view untouched", log
    veryHidden.Visible = xlSheetVeryHidden
    chart.Activate
    Application.Run "'SheetFreak.xlam'!SF_ResizeAll"
    Check ActiveSheet Is chart, "chart tab restored without resizing chart", log
    Check Application.EnableEvents = oldEvents And Application.ScreenUpdating = oldScreen, "application state restored", log
    other.Activate
    Check ActiveWindow.Zoom = 93 And ActiveWindow.DisplayGridlines, "other workbook view untouched", log
    log = log & "ALL WORKBOOK VIEW CHECKS PASSED" & vbCrLf
    GoTo Finish
Failed:
    log = log & "FAIL: " & Err.Description & vbCrLf
Finish:
    On Error Resume Next
    If Not other Is Nothing Then other.Close False
    If Not book Is Nothing Then book.Close False
    If Not originalBook Is Nothing Then originalBook.Activate
    If Not originalSheet Is Nothing Then originalSheet.Activate
    Application.EnableEvents = oldEvents: Application.ScreenUpdating = oldScreen
    On Error GoTo 0
    file = FreeFile
    Open reportPath For Output As #file
    Print #file, log
    Close #file
End Sub

Private Function ArraySheets(ByVal first As Worksheet, ByVal second As Worksheet) As Collection
    Dim result As New Collection
    result.Add first: result.Add second
    Set ArraySheets = result
End Function

Private Sub Check(ByVal condition As Boolean, ByVal description As String, ByRef log As String)
    If Not condition Then Err.Raise 513, "SheetFreak workbook view tests", description
    log = log & "PASS: " & description & vbCrLf
End Sub
