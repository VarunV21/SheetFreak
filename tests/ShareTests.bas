Attribute VB_Name = "ShareTests"
Option Explicit
Option Private Module

' Import into a disposable test host and run against the candidate add-in.
Public Sub RunShareChecks(ByVal reportPath As String)
    Dim book As Workbook, copy As Workbook, fso As Object, root As String
    Dim first As String, second As String, originalName As String, originalSaved As Boolean
    Dim log As String, file As Integer, oldEvents As Boolean, oldAlerts As Boolean
    Dim beforeCount As Long, failure As Long, note As String
    oldEvents = Application.EnableEvents: oldAlerts = Application.DisplayAlerts
    On Error GoTo Failed
    Set fso = CreateObject("Scripting.FileSystemObject")
    root = fso.BuildPath(fso.GetParentFolderName(reportPath), "share-test-" & fso.GetTempName)
    fso.CreateFolder root
    Set book = Workbooks.Add(xlWBATWorksheet)
    book.Worksheets(1).Range("A1").Value2 = "Saved version"
    book.SaveAs fso.BuildPath(root, "Source.xlsx"), xlOpenXMLWorkbook
    book.Worksheets(1).Range("A1").Value2 = "Unsaved change"
    book.Worksheets(1).Range("B1").Formula2 = "=1+2"
    book.Worksheets(1).Range("B1").Font.Color = RGB(112, 48, 160)
    book.Worksheets.Add(After:=book.Worksheets(book.Worksheets.Count)).Name = "Hidden data"
    book.Worksheets("Hidden data").Range("A1").Value2 = "Preserved"
    book.Worksheets("Hidden data").Visible = xlSheetVeryHidden
    book.Worksheets(1).Activate
    originalName = book.FullName: originalSaved = book.Saved: beforeCount = Workbooks.Count
    first = Application.Run("'SheetFreak.xlam'!SF_CreateShareCopy", book, root)
    Check fso.FileExists(first), "snapshot created", log
    Check fso.GetFileName(first) = "Source_vS.xlsx", "snapshot uses _vS suffix", log
    Check book.FullName = originalName And book.Saved = originalSaved, "source name and unsaved state unchanged", log
    Check Workbooks.Count = beforeCount And ActiveWorkbook Is book, "source stays active and no snapshot workbook is opened", log
    Check Application.EnableEvents = oldEvents, "event state restored after success", log
    second = Application.Run("'SheetFreak.xlam'!SF_CreateShareCopy", book, root)
    Check first <> second And fso.FileExists(first), "repeated snapshots do not overwrite", log
    Set copy = Workbooks.Open(first, UpdateLinks:=0, ReadOnly:=True)
    Check copy.Worksheets(1).Range("A1").Value2 = "Unsaved change", "includes unsaved edits", log
    Check copy.Worksheets(1).Range("B1").Formula2 = "=1+2", "preserves formulas", log
    Check copy.Worksheets(1).Range("B1").Font.Color = RGB(112, 48, 160), "preserves formatting", log
    Check copy.Worksheets("Hidden data").Visible = xlSheetVeryHidden, "preserves hidden sheets", log
    copy.Close False: Set copy = Nothing
    book.Close False: Set book = Nothing
    Set copy = Workbooks.Open(originalName, UpdateLinks:=0, ReadOnly:=True)
    Check copy.Worksheets(1).Range("A1").Value2 = "Saved version", "original disk file was not saved", log
    copy.Close False: Set copy = Nothing
    Set book = Workbooks.Add(xlWBATWorksheet)
    book.Worksheets(1).Range("A1").Value2 = "Never saved"
    first = Application.Run("'SheetFreak.xlam'!SF_CreateShareCopy", book, root)
    Check book.Path = "" And Not book.Saved And fso.FileExists(first), "never-saved workbook supported without Save As", log
    Application.EnableEvents = False
    second = Application.Run("'SheetFreak.xlam'!SF_CreateShareCopy", book, root)
    Check Not Application.EnableEvents, "disabled events remain disabled", log
    Application.EnableEvents = oldEvents
    On Error Resume Next
    Err.Clear
    second = Application.Run("'SheetFreak.xlam'!SF_CreateShareCopy", book, fso.BuildPath(root, "missing\nested"))
    failure = Err.Number: Err.Clear
    On Error GoTo Failed
    Check failure <> 0 And Application.EnableEvents = oldEvents, "failure reported without changing event state", log
    log = log & "ALL SHARE CHECKS PASSED" & vbCrLf
    GoTo Cleanup
Failed:
    note = Err.Description
    log = log & "FAIL: " & note & vbCrLf
Cleanup:
    On Error Resume Next
    If Not copy Is Nothing Then copy.Close False
    If Not book Is Nothing Then book.Close False
    Application.EnableEvents = oldEvents: Application.DisplayAlerts = oldAlerts
    file = FreeFile: Open reportPath For Output As #file: Print #file, log: Close #file
    ' Leave the synthetic snapshots available for inspection, not deletion.
    On Error GoTo 0
    If Len(note) > 0 Then Err.Raise 5, , note
End Sub

Private Sub Check(ByVal condition As Boolean, ByVal description As String, ByRef log As String)
    If Not condition Then Err.Raise 5, , description
    log = log & "PASS: " & description & vbCrLf
End Sub
