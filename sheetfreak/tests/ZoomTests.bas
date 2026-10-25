Attribute VB_Name = "ZoomTests"
Option Explicit
Option Private Module

' Run against the newly built add-in from a disposable test workbook.
' The report path must be supplied explicitly. Existing user workbooks are untouched.
Public Sub RunZoomChecks(ByVal reportPath As String)
    Dim book As Workbook, sheet As Worksheet, file As Integer, log As String
    Dim width As Double, beforeCount As Long, lastColumn As Long, expected As Long
    Dim originalBook As Workbook, originalSheet As Object, originalSelection As String
    On Error GoTo Failed
    Set originalBook = ActiveWorkbook
    Set originalSheet = ActiveSheet
    If TypeName(Selection) = "Range" Then originalSelection = Selection.Address
    beforeCount = Workbooks.Count
    Set book = Workbooks.Add(xlWBATWorksheet)
    Set sheet = book.Worksheets(1)
    sheet.Range("J17").Value2 = 5233
    sheet.Columns("A:J").ColumnWidth = 5
    width = sheet.Columns("A:J").Width
    sheet.Range("A1:J1").Select
    ActiveWindow.Zoom = True
    expected = ActiveWindow.Zoom
    If expected < 70 Then expected = 70
    If expected > 200 Then expected = 200
    sheet.Range("J17").Select
    Application.Run "'SheetFreak.xlam'!SF_Home"
    Check ActiveWindow.Zoom = expected, "compact sheet fits width within 70-200%", log
    Check ActiveWindow.Zoom > 100, "compact sheet zooms above 100%", log
    Check Selection.Address = "$A$1", "returns to A1", log
    Check sheet.Columns("A:J").Width = width, "column widths unchanged", log
    Check sheet.Range("J17").Value2 = 5233, "cell contents unchanged", log
    sheet.Range("J17").ClearContents
    sheet.Range("AZ17").Value2 = 42
    sheet.Columns("A:AZ").ColumnWidth = 12
    sheet.Range("A1:AZ1").Select
    ActiveWindow.Zoom = True
    expected = ActiveWindow.Zoom
    If expected < 70 Then expected = 70
    If expected > 200 Then expected = 200
    sheet.Range("AZ17").Select
    Application.Run "'SheetFreak.xlam'!SF_Home"
    Check ActiveWindow.Zoom = expected, "wide sheet still fits within bounds", log
    sheet.Range("AZ17").ClearContents
    Application.Run "'SheetFreak.xlam'!SF_Home"
    Check ActiveWindow.Zoom = 100, "empty sheet uses 100%", log
    Check Workbooks.Count = beforeCount + 1, "Auto resize creates no extra workbook", log
    log = log & "ALL ZOOM CHECKS PASSED" & vbCrLf
    GoTo Finish
Failed:
    log = log & "FAIL: " & Err.Description & vbCrLf
Finish:
    On Error Resume Next
    If Not book Is Nothing Then book.Close SaveChanges:=False
    If Not originalBook Is Nothing Then originalBook.Activate
    If Not originalSheet Is Nothing Then originalSheet.Activate
    If Len(originalSelection) > 0 Then originalSheet.Range(originalSelection).Select
    On Error GoTo 0
    file = FreeFile
    Open reportPath For Output As #file
    Print #file, log
    Close #file
End Sub

Private Sub Check(ByVal condition As Boolean, ByVal description As String, ByRef log As String)
    If Not condition Then Err.Raise 513, "SheetFreak zoom tests", description
    log = log & "PASS: " & description & vbCrLf
End Sub
