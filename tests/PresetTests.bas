Attribute VB_Name = "PresetTests"
Option Explicit
Option Private Module

Public Sub RunPresetChecks(ByVal reportPath As String)
    Dim book As Workbook, sheet As Worksheet, target As Range, log As String
    Dim original As Workbook, file As Integer, message As String
    On Error GoTo Failed
    Set original = ActiveWorkbook
    Set book = Workbooks.Add(xlWBATWorksheet): Set sheet = book.Worksheets(1)
    Set target = sheet.Range("A1:A3")
    target.Value2 = 1234.567: target.Font.Bold = True
    target.Interior.Color = RGB(200, 220, 240)
    target.NumberFormat = "0.00": target.Select
    Application.Run "'SheetFreak.xlam'!SF_Notes"
    Check target.Font.Color = RGB(128, 128, 128) And target.Font.Italic, "Notes applies gray italics", log
    Check target.Font.Bold And target.Interior.Color = RGB(200, 220, 240), "Notes preserves bold and fill", log
    Application.Run "'SheetFreak.xlam'!SF_BlueFont"
    Check target.Font.Color = RGB(0, 0, 255), "blue font on range", log
    Application.Run "'SheetFreak.xlam'!SF_GreenFont"
    Check target.Font.Color = RGB(0, 128, 0), "green font on range", log
    Application.Run "'SheetFreak.xlam'!SF_RedFont"
    Check target.Font.Color = RGB(255, 0, 0), "red font on range", log
    Check target.Font.Bold And target.Font.Italic And target.NumberFormat = "0.00", "font colors preserve font style and number format", log
    Check target.Interior.Color = RGB(200, 220, 240), "font colors preserve fill", log
    Application.Run "'SheetFreak.xlam'!SF_YellowFill"
    Check target.Interior.Color = RGB(255, 242, 204) And target.Interior.Pattern = xlSolid, "highlight applies pale yellow", log
    Check target.Font.Color = RGB(255, 0, 0) And target.Font.Italic, "highlight preserves red italic font", log
    sheet.Range("A2").Value2 = -1234.567: sheet.Range("A3").Value2 = 0
    sheet.Columns("A").ColumnWidth = 24
    Application.Run "'SheetFreak.xlam'!SF_Currency"
    Check sheet.Range("A1").Text = "$1,235", "currency displays dollar with zero decimals", log
    Check sheet.Range("A2").Text = "($1,235)" And sheet.Range("A3").Text = "$0", "currency negative and zero display", log
    Check InStr(target.NumberFormat, "*") = 0, "currency has no accounting alignment", log
    Check sheet.Range("A1").Value2 = 1234.567, "currency preserves stored precision", log
    Application.Run "'SheetFreak.xlam'!SF_Comma"
    Check sheet.Range("A1").Text = "1,235" And sheet.Range("A2").Text = "(1,235)", "comma preset displays integers and negatives", log
    Check sheet.Range("A3").Text = "0", "comma zero display", log
    sheet.Range("A1").Value2 = 0.12567
    Application.Run "'SheetFreak.xlam'!SF_Percent"
    Check sheet.Range("A1").Text = "12.6%" And sheet.Range("A1").Value2 = 0.12567, "percent shows one decimal without changing value", log
    Check target.Font.Color = RGB(255, 0, 0) And target.Font.Italic And target.Interior.Color = RGB(255, 242, 204), "number presets preserve font and fill", log
    sheet.Range("B1").Formula2 = "=1/8": sheet.Range("B1").Select
    Application.Run "'SheetFreak.xlam'!SF_Percent"
    Check sheet.Range("B1").Formula2 = "=1/8", "formatting preserves formula", log
    sheet.Range("C1,E1").Select
    Application.Run "'SheetFreak.xlam'!SF_GreenFont"
    Check sheet.Range("C1").Font.Color = RGB(0, 128, 0) And sheet.Range("E1").Font.Color = RGB(0, 128, 0), "font preset supports nonadjacent cells", log
    Check sheet.Range("D1").Font.Color <> RGB(0, 128, 0), "nonselected cell unchanged", log
    Set target = sheet.Range("G1:H2"): target.Value2 = "Header"
    target.Borders.LineStyle = xlContinuous: target.Borders.Color = vbWhite
    target.Select: Application.Run "'SheetFreak.xlam'!SF_Header"
    Check target.Interior.Color = RGB(64, 64, 64) And target.Font.Color = vbWhite And target.Font.Bold, "dark header colors and bold", log
    Check target.Borders(xlInsideVertical).LineStyle = xlNone And target.Borders(xlInsideHorizontal).LineStyle = xlNone, "dark header removes inside borders", log
    Check target.Borders(xlEdgeTop).LineStyle = xlNone And target.Borders(xlEdgeBottom).LineStyle = xlNone, "dark header removes top and bottom borders", log
    Check target.Borders(xlEdgeLeft).LineStyle = xlNone And target.Borders(xlEdgeRight).LineStyle = xlNone, "dark header removes left and right borders", log
    target.Borders.LineStyle = xlContinuous: target.Borders.Color = vbWhite
    Application.Run "'SheetFreak.xlam'!SF_LightHeader"
    Check target.Interior.Color = RGB(217, 217, 217) And target.Font.Color = vbBlack And target.Font.Bold, "light header colors and bold", log
    Check target.Borders(xlInsideVertical).LineStyle = xlNone And target.Borders(xlInsideHorizontal).LineStyle = xlNone, "light header removes inside borders", log
    Check target.Borders(xlEdgeTop).LineStyle = xlNone And target.Borders(xlEdgeBottom).LineStyle = xlNone, "light header removes top and bottom borders", log
    Check target.Borders(xlEdgeLeft).LineStyle = xlNone And target.Borders(xlEdgeRight).LineStyle = xlNone And target.Cells(1, 1).Value2 = "Header", "light header removes side borders and preserves text", log
    log = log & "ALL PRESET CHECKS PASSED" & vbCrLf
    GoTo Cleanup
Failed:
    message = Err.Description: log = log & "FAIL: " & message & vbCrLf
Cleanup:
    On Error Resume Next
    If Not book Is Nothing Then book.Close False
    If Not original Is Nothing Then original.Activate
    file = FreeFile: Open reportPath For Output As #file: Print #file, log: Close #file
    On Error GoTo 0
    If Len(message) > 0 Then Err.Raise 5, , message
End Sub

Private Sub Check(ByVal condition As Boolean, ByVal description As String, ByRef log As String)
    If Not condition Then Err.Raise 5, , description
    log = log & "PASS: " & description & vbCrLf
End Sub
