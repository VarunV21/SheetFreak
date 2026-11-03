Attribute VB_Name = "SheetFreakShare"
Option Explicit

' Create a local, closed snapshot; the user chooses whether/where to upload it.
Public Sub SF_ShareExcel(Optional control As Variant)
    Dim book As Workbook, snapshot As String, shellApp As Object
    On Error GoTo Failed
    Set book = Application.ActiveWorkbook
    snapshot = SF_CreateShareCopy(book)
    On Error GoTo RevealFailed
    Set shellApp = CreateObject("Shell.Application")
    shellApp.ShellExecute "explorer.exe", "/select," & Chr$(34) & snapshot & Chr$(34), "", "open", 1
    Exit Sub
RevealFailed:
    MsgBox "Your snapshot is ready at:" & vbCrLf & snapshot & vbCrLf & vbCrLf & _
        "File Explorer could not be opened. Attach this file using your chat's upload button.", vbInformation, "SheetFreak - Share Excel"
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "SheetFreak - Share Excel"
End Sub

' Public for native tests. Each invocation owns a new directory, never overwrites.
Public Function SF_CreateShareCopy(ByVal book As Workbook, Optional ByVal shareRoot As String = "") As String
    Dim fso As Object, folder As String, snapshot As String, extension As String
    Dim stem As String, bad As Variant, oldEvents As Boolean, changedEvents As Boolean
    Dim failureNumber As Long, failureMessage As String
    On Error GoTo Failed
    If book Is Nothing Then Err.Raise 5, , "Open the workbook you want to share first."
    If book Is ThisWorkbook Then Err.Raise 5, , "Choose your working workbook first."
    If book.IsAddin Then Err.Raise 5, , "Choose a workbook, rather than an Excel add-in."
    Select Case book.FileFormat
        Case xlOpenXMLWorkbook: extension = ".xlsx"
        Case xlOpenXMLWorkbookMacroEnabled: extension = ".xlsm"
        Case xlExcel12: extension = ".xlsb"
        Case xlExcel8: extension = ".xls"
        Case Else: Err.Raise 5, , "Save this workbook as .xlsx or .xlsm first, then try Share Excel again."
    End Select
    If extension = ".xlsx" And book.HasVBProject Then _
        Err.Raise 5, , "Save this workbook as .xlsm first so the snapshot can preserve its macros."
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Len(shareRoot) = 0 Then
        shareRoot = fso.BuildPath(Environ$("TEMP"), "SheetFreak")
        If Not fso.FolderExists(shareRoot) Then fso.CreateFolder shareRoot
        shareRoot = fso.BuildPath(shareRoot, "Share")
    End If
    If Not fso.FolderExists(shareRoot) Then fso.CreateFolder shareRoot
    folder = fso.BuildPath(shareRoot, Format$(Now, "yyyymmdd-hhnnss") & "-" & fso.GetTempName)
    fso.CreateFolder folder
    stem = fso.GetBaseName(book.Name)
    For Each bad In Array("\", "/", ":", "*", "?", Chr$(34), "<", ">", "|", "[", "]")
        stem = Replace(stem, CStr(bad), "_")
    Next bad
    stem = Left$(stem, 60)
    snapshot = fso.BuildPath(folder, stem & " - Share" & extension)
    oldEvents = Application.EnableEvents: changedEvents = True
    Application.EnableEvents = False
    book.SaveCopyAs snapshot
    Application.EnableEvents = oldEvents: changedEvents = False
    If Not fso.FileExists(snapshot) Then Err.Raise 5, , "Excel did not create the snapshot. Try again."
    SF_CreateShareCopy = snapshot
    Exit Function
Failed:
    failureNumber = Err.Number: failureMessage = Err.Description
    On Error Resume Next
    If changedEvents Then Application.EnableEvents = oldEvents
    On Error GoTo 0
    Err.Raise failureNumber, "SheetFreak", "Share Excel could not create a snapshot. " & failureMessage
End Function
