Attribute VB_Name = "SheetFreak"
Option Explicit

' SheetFreak 0.2.1 | Windows desktop Excel | MIT license.
Private mCopied As Variant
Private mCopyRows As Long, mCopyCols As Long
Private mHasCopy As Boolean
Private mCopyIsFormula() As Boolean
Private Const MAX_FORMAT_CELLS As Long = 250000
Private Const MAX_COPY_CELLS As Long = 100000
Private Const MAX_FILL_CELLS As Long = 5000000

Public Sub Auto_Open()
    SF_RegisterKeys
End Sub

Public Sub Auto_Close()
    SF_UnregisterKeys
End Sub

Public Sub SF_RegisterKeys()
    Dim names As Variant, keys As Variant, i As Long
    names = Array("SF_ZeroDecimals", "SF_CopyExact", "SF_PasteExact", "SF_Home", "SF_Assumption", "SF_Header", "SF_FillFreeze")
    keys = Array("^%0", "^%c", "^%v", "^%h", "^%a", "^%t", "^%f")
    For i = 0 To UBound(keys)
        Application.OnKey CStr(keys(i)), "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & names(i)
    Next i
End Sub

Public Sub SF_UnregisterKeys()
    Dim key As Variant
    For Each key In Array("^%0", "^%c", "^%v", "^%h", "^%a", "^%t", "^%f")
        Application.OnKey CStr(key)
    Next key
End Sub

Private Function SelectedCells(Optional ByVal allowManySheets As Boolean = False) As Range
    If TypeName(Application.Selection) <> "Range" Then Err.Raise 5, , "Select worksheet cells first."
    If ActiveWorkbook Is ThisWorkbook Then Err.Raise 5, , "Select cells in your working workbook."
    If Not allowManySheets Then
        If ActiveWindow.SelectedSheets.Count <> 1 Then Err.Raise 5, , "Ungroup the worksheet tabs first."
    End If
    Set SelectedCells = Application.Selection
End Function

Private Sub Writable(ByVal target As Range)
    If target.Worksheet.ProtectContents Then Err.Raise 5, , "Unprotect this sheet first."
    If target.Worksheet.Parent.ReadOnly Then Err.Raise 5, , "This workbook is read-only. Save an editable copy first."
End Sub

Private Sub SmallSelection(ByVal target As Range, ByVal limit As Long)
    If target.CountLarge > limit Then Err.Raise 5, , "Select a smaller range (limit: " & Format$(limit, "#,##0") & " cells)."
End Sub

Private Sub NoMerged(ByVal target As Range)
    Dim merged As Variant
    merged = target.MergeCells
    If IsNull(merged) Then Err.Raise 5, , "Merged cells are not supported for this operation."
    If merged Then Err.Raise 5, , "Merged cells are not supported for this operation."
End Sub

Private Sub Fail(ByVal description As String)
    MsgBox description, vbExclamation, "SheetFreak"
End Sub

Public Sub SF_Assumption(Optional control As Variant)
    Dim target As Range
    On Error GoTo Failed
    Set target = SelectedCells(): Writable target: SmallSelection target, MAX_FORMAT_CELLS
    With target
        .Font.Color = RGB(0, 0, 255)
        .Interior.Pattern = xlSolid
        .Interior.Color = RGB(255, 242, 204)
    End With
    Exit Sub
Failed:
    Fail Err.Description
End Sub

Public Sub SF_Header(Optional control As Variant)
    Dim target As Range, area As Range, edge As Variant
    On Error GoTo Failed
    Set target = SelectedCells(): Writable target: SmallSelection target, MAX_FORMAT_CELLS
    NoMerged target
    With target
        .Font.Color = vbWhite
        .Font.Bold = True
        .Interior.Pattern = xlSolid
        .Interior.Color = RGB(64, 64, 64)
    End With
    For Each area In target.Areas
        For Each edge In Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight, xlInsideVertical, xlInsideHorizontal)
            With area.Borders(CLng(edge))
                .LineStyle = xlContinuous
                .Weight = xlThin
                .Color = vbWhite
            End With
        Next edge
    Next area
    Exit Sub
Failed:
    Fail Err.Description
End Sub

' Strips quoted/escaped literals and bracket tokens for format classification.
Private Function FormatTokens(ByVal code As String) As String
    Dim i As Long, ch As String, quoted As Boolean, bracket As Boolean, result As String
    For i = 1 To Len(code)
        ch = Mid$(code, i, 1)
        If ch = """" And Not bracket Then
            quoted = Not quoted
        ElseIf Not quoted Then
            If ch = "\" Or ch = "_" Or ch = "*" Then
                i = i + 1
            ElseIf ch = "[" Then
                bracket = True
            ElseIf ch = "]" Then
                bracket = False
            ElseIf Not bracket Then
                result = result & LCase$(ch)
            End If
        End If
    Next i
    FormatTokens = result
End Function

' Pure function: unsupported formats return unchanged. NumberFormat is invariant.
Public Function SF_ZeroFormat(ByVal code As String) As String
    Dim tokens As String, i As Long, j As Long, ch As String, result As String
    Dim quoted As Boolean, bracket As Boolean, nextCh As String
    SF_ZeroFormat = code
    If LCase$(code) = "general" Then SF_ZeroFormat = "0": Exit Function
    tokens = FormatTokens(code)
    ' Date/time, elapsed time, fractions and text-only formats are intentionally preserved.
    If InStr(tokens, "y") Or InStr(tokens, "d") Or InStr(tokens, "h") Or InStr(tokens, "s") Or InStr(tokens, "m") Or InStr(tokens, "/") Then Exit Function
    If InStr(LCase$(code), "[h") Or InStr(LCase$(code), "[m") Or InStr(LCase$(code), "[s") Then Exit Function
    If InStr(tokens, "0") = 0 And InStr(tokens, "#") = 0 And InStr(tokens, "?") = 0 Then Exit Function
    i = 1
    Do While i <= Len(code)
        ch = Mid$(code, i, 1)
        If ch = """" And Not bracket Then
            quoted = Not quoted
        ElseIf Not quoted Then
            If ch = "\" Or ch = "_" Or ch = "*" Then
                result = result & Mid$(code, i, 2): i = i + 2: GoTo NextCharacter
            ElseIf ch = "[" Then
                bracket = True
            ElseIf ch = "]" Then
                bracket = False
            ElseIf ch = "." And Not bracket Then
                j = i + 1
                Do While j <= Len(code)
                    nextCh = Mid$(code, j, 1)
                    If nextCh <> "0" And nextCh <> "#" And nextCh <> "?" Then Exit Do
                    j = j + 1
                Loop
                If j > i + 1 Then
                    ' .00 or .## without an integer placeholder needs a leading zero.
                    If i = 1 Then
                        result = result & "0"
                    ElseIf InStr("0#?", Mid$(code, i - 1, 1)) = 0 Then
                        result = result & "0"
                    End If
                    i = j: GoTo NextCharacter
                End If
            End If
        End If
        result = result & ch
        i = i + 1
NextCharacter:
    Loop
    SF_ZeroFormat = result
End Function

Public Sub SF_ZeroDecimals(Optional control As Variant)
    Dim target As Range, area As Range, cell As Range, code As Variant
    Dim oldScreen As Boolean, message As String, changedState As Boolean
    On Error GoTo Failed
    Set target = SelectedCells(): Writable target: SmallSelection target, MAX_FORMAT_CELLS
    oldScreen = Application.ScreenUpdating: changedState = True
    Application.ScreenUpdating = False
    For Each area In target.Areas
        code = area.NumberFormat
        If Not IsNull(code) Then
            area.NumberFormat = SF_ZeroFormat(CStr(code))
        Else
            For Each cell In area.Cells
                cell.NumberFormat = SF_ZeroFormat(CStr(cell.NumberFormat))
            Next cell
        End If
    Next area
    Application.ScreenUpdating = oldScreen
    Exit Sub
Failed:
    message = Err.Description
    If changedState Then Application.ScreenUpdating = oldScreen
    Fail message
End Sub

Public Sub SF_CopyExact(Optional control As Variant)
    Dim target As Range, row As Long, col As Long
    On Error GoTo Failed
    Set target = SelectedCells()
    If target.Areas.Count <> 1 Then Err.Raise 5, , "Copy one rectangular range."
    SmallSelection target, MAX_COPY_CELLS: NoMerged target
    mCopied = target.Formula2
    mCopyRows = target.Rows.Count: mCopyCols = target.Columns.Count: mHasCopy = True
    ReDim mCopyIsFormula(1 To mCopyRows, 1 To mCopyCols)
    For row = 1 To mCopyRows
        For col = 1 To mCopyCols
            mCopyIsFormula(row, col) = target.Cells(row, col).HasFormula
        Next col
    Next row
    Application.StatusBar = "SheetFreak: exact formulas copied. Choose a destination and press Ctrl+Alt+V."
    Exit Sub
Failed:
    Fail Err.Description
End Sub

Public Sub SF_PasteExact(Optional control As Variant)
    Dim selected As Range, target As Range, area As Range
    Dim cell As Range, data As Variant, row As Long, col As Long, value As Variant
    Dim oldEvents As Boolean, oldScreen As Boolean, changedState As Boolean, message As String
    On Error GoTo Failed
    If Not mHasCopy Then Err.Raise 5, , "Use SheetFreak Copy exact formulas first (Ctrl+Alt+C)."
    Set selected = SelectedCells(): Writable selected
    If selected.Areas.Count <> 1 Then Err.Raise 5, , "Paste into one rectangular range."
    If selected.CountLarge = 1 Then
        If selected.Row + mCopyRows - 1 > selected.Worksheet.Rows.Count Or selected.Column + mCopyCols - 1 > selected.Worksheet.Columns.Count Then Err.Raise 5, , "The copied range would extend beyond the sheet."
        Set target = selected.Resize(mCopyRows, mCopyCols)
    Else
        If selected.Rows.Count <> mCopyRows Or selected.Columns.Count <> mCopyCols Then Err.Raise 5, , "Select one destination cell or a range the same size as the copy."
        Set target = selected
    End If
    NoMerged target: ValidateOrdinaryRange target
    If target.Worksheet.FilterMode Then Err.Raise 5, , "Clear active filters before pasting exact formulas."
    If Application.CountA(target) > 0 Then
        If MsgBox("Replace existing content in " & target.Address(External:=True) & "?" & vbCrLf & "Normal Excel Undo will not restore this paste.", vbYesNo + vbQuestion + vbDefaultButton2, "SheetFreak - Paste exact") <> vbYes Then Exit Sub
    End If
    oldEvents = Application.EnableEvents: oldScreen = Application.ScreenUpdating: changedState = True
    Application.EnableEvents = False: Application.ScreenUpdating = False
    ' Per-cell assignment distinguishes literal text beginning with =, + or - from formulas.
    For row = 1 To mCopyRows
        For col = 1 To mCopyCols
            If mCopyRows = 1 And mCopyCols = 1 Then value = mCopied Else value = mCopied(row, col)
            Set cell = target.Cells(row, col)
            If VarType(value) = vbString Then
                If mCopyIsFormula(row, col) Then
                    cell.Formula2 = value
                Else
                    cell.Value2 = "'" & CStr(value)
                End If
            Else
                cell.Value2 = value
            End If
        Next col
    Next row
    Application.EnableEvents = oldEvents: Application.ScreenUpdating = oldScreen
    Application.StatusBar = False
    Exit Sub
Failed:
    message = Err.Description
    If changedState Then Application.EnableEvents = oldEvents: Application.ScreenUpdating = oldScreen
    Fail message
End Sub

Public Sub SF_Home(Optional control As Variant)
    Dim target As Range, lastCell As Range, fitRange As Range, zoom As Long
    On Error GoTo Failed
    Set target = SelectedCells()
    Set lastCell = target.Worksheet.Cells.Find(What:="*", After:=target.Worksheet.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False, SearchFormat:=False)
    If lastCell Is Nothing Then
        ActiveWindow.Zoom = 100
    Else
        Set fitRange = target.Worksheet.Range(target.Worksheet.Cells(1, 1), target.Worksheet.Cells(1, lastCell.Column))
        fitRange.Select
        ActiveWindow.Zoom = True
        zoom = ActiveWindow.Zoom
        If zoom < 70 Then zoom = 70
        If zoom > 200 Then zoom = 200
        ActiveWindow.Zoom = zoom
    End If
    Application.Goto target.Worksheet.Range("A1"), True
    Exit Sub
Failed:
    Fail Err.Description
End Sub

Private Sub ValidateOrdinaryRange(ByVal target As Range)
    Dim table As ListObject, arrayState As Variant, spillState As Variant
    Writable target: NoMerged target
    arrayState = target.HasArray
    If IsNull(arrayState) Then Err.Raise 5, , "Legacy array formulas are not supported in the target range."
    If arrayState Then Err.Raise 5, , "Legacy array formulas are not supported in the target range."
    spillState = target.HasSpill
    If IsNull(spillState) Then Err.Raise 5, , "Dynamic spill formulas are not supported in the target range."
    If spillState Then Err.Raise 5, , "Dynamic spill formulas are not supported in the target range."
    For Each table In target.Worksheet.ListObjects
        If Not Application.Intersect(target, table.Range) Is Nothing Then Err.Raise 5, , "Convert the target Excel Table to a normal range before using this tool."
    Next table
End Sub

' Non-UI validation also used by the native Excel regression tests.
Public Function SF_ValidateFill(ByVal top As Range, ByVal key As Range) As Long
    Dim cell As Range, lastRow As Long, target As Range
    If top.Areas.Count <> 1 Or top.Rows.Count <> 1 Then Err.Raise 5, , "Select one row containing only the starting formulas."
    If Not top.Worksheet Is key.Worksheet Then Err.Raise 5, , "Choose a last-row reference on the same sheet."
    If key.Columns.Count <> 1 Then Err.Raise 5, , "Choose cells in just one reference column."
    If top.Worksheet.Parent Is ThisWorkbook Then Err.Raise 5, , "Choose your working workbook."
    If top.Worksheet.FilterMode Then Err.Raise 5, , "Clear active filters first. Fill-and-freeze affects every row, including hidden rows."
    For Each cell In top.Cells
        If Not cell.HasFormula Then Err.Raise 5, , "Every selected starting cell must contain a formula."
    Next cell
    lastRow = key.Worksheet.Cells(key.Worksheet.Rows.Count, key.Column).End(xlUp).Row
    If lastRow <= top.Row Then Err.Raise 5, , "No data rows were found below the starting formulas in that column."
    Set target = top.Resize(lastRow - top.Row + 1)
    SmallSelection target, MAX_FILL_CELLS
    ValidateOrdinaryRange target
    SF_ValidateFill = lastRow
End Function

Public Sub SF_FillFreeze(Optional control As Variant)
    Dim top As Range, key As Range, lastRow As Long, target As Range, message As String
    On Error GoTo Failed
    Set top = SelectedCells()
    If top.Areas.Count <> 1 Or top.Rows.Count <> 1 Then Err.Raise 5, , "Select one row containing the starting formulas first."
    On Error Resume Next
    Set key = Application.InputBox("Select any cell in the column that determines the last data row." & vbCrLf & "Use the same worksheet as your starting formulas.", "SheetFreak - Last-row column", Type:=8)
    On Error GoTo Failed
    If key Is Nothing Then Exit Sub
    lastRow = SF_ValidateFill(top, key)
    Set target = top.Resize(lastRow - top.Row + 1)
    message = "Workbook: " & top.Worksheet.Parent.Name & vbCrLf & "Sheet: " & top.Worksheet.Name & vbCrLf & _
        "Keep starting formulas: " & top.Address(False, False) & vbCrLf & _
        "Overwrite and freeze: " & target.Offset(1).Resize(target.Rows.Count - 1).Address(False, False) & vbCrLf & _
        "Last-row column: " & Split(key.EntireColumn.Address(False, False), ":")(0) & vbCrLf & _
        "Last row: " & Format$(lastRow, "#,##0") & vbCrLf & _
        "Cells replaced: " & Format$(target.CountLarge - top.CountLarge, "#,##0") & vbCrLf & vbCrLf & _
        "ALL rows in this range will be overwritten, including hidden rows." & vbCrLf & _
        "Lower rows become values with black text; affected columns are AutoFit." & vbCrLf & _
        "Results are snapshots. Normal Excel Undo cannot restore them." & vbCrLf & vbCrLf & _
        "No backup or recovery workbook will be created." & vbCrLf & _
        "Save your working workbook before continuing if you need a recovery point." & vbCrLf & vbCrLf & "Continue?"
    If MsgBox(message, vbYesNo + vbExclamation + vbDefaultButton2, "SheetFreak - Confirm fill and freeze") <> vbYes Then Exit Sub
    SF_ExecuteFill top, key
    Exit Sub
Failed:
    Fail Err.Description
End Sub

' Internal engine. Call SF_FillFreeze for the interactive confirmation flow.
' All target formulas are populated and calculated before ANY formulas are frozen.
' This preserves cross-column and previous/next-row dependencies.
Public Sub SF_ExecuteFill(ByVal top As Range, ByVal key As Range)
    Dim lastRow As Long, target As Range, lower As Range, part As Range
    Dim oldCalc As XlCalculation, oldEvents As Boolean, oldScreen As Boolean
    Dim oldCancel As XlEnableCancelKey, oldStatus As Variant, changedState As Boolean
    Dim message As String, row As Long, rowsPerBlock As Long
    Dim count As Long, col As Long, errors As Range, originalNumber As Long
    On Error GoTo Failed
    lastRow = SF_ValidateFill(top, key)
    Set target = top.Resize(lastRow - top.Row + 1)
    Set lower = target.Offset(1).Resize(target.Rows.Count - 1)
    oldCalc = Application.Calculation: oldEvents = Application.EnableEvents
    oldScreen = Application.ScreenUpdating: oldStatus = Application.StatusBar
    oldCancel = Application.EnableCancelKey: changedState = True
    Application.EnableEvents = False: Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual: Application.EnableCancelKey = xlErrorHandler
    Application.StatusBar = "SheetFreak: filling formulas..."
    target.FillDown
    Application.StatusBar = "SheetFreak: calculating (this may take time)..."
    Application.Calculate
    If Application.CalculationState <> xlDone Then Err.Raise 5, , "Calculation is still pending. Nothing has been frozen."
    On Error Resume Next
    Set errors = target.SpecialCells(xlCellTypeFormulas, xlErrors)
    On Error GoTo Failed
    If Not errors Is Nothing Then Err.Raise 5, , "The calculated range contains errors. Nothing has been frozen. The target has been filled with formulas; check them before retrying."
    rowsPerBlock = 100000 \ top.Columns.Count
    If rowsPerBlock < 1 Then rowsPerBlock = 1
    For row = lower.Row To lastRow Step rowsPerBlock
        count = rowsPerBlock
        If row + count - 1 > lastRow Then count = lastRow - row + 1
        Set part = lower.Worksheet.Cells(row, lower.Column).Resize(count, lower.Columns.Count)
        Application.StatusBar = "SheetFreak: freezing through row " & Format$(row + count - 1, "#,##0") & "..."
        part.Value2 = part.Value2
    Next row
    For col = 1 To top.Columns.Count
        lower.Columns(col).NumberFormat = top.Cells(1, col).NumberFormat
    Next col
    lower.Font.Color = vbBlack
    target.EntireColumn.AutoFit
    top.Select
    GoTo RestoreState
Failed:
    originalNumber = Err.Number: message = Err.Description
RestoreState:
    On Error Resume Next
    If changedState Then
        Application.Calculation = oldCalc: Application.EnableEvents = oldEvents
        Application.ScreenUpdating = oldScreen: Application.StatusBar = oldStatus
        Application.EnableCancelKey = oldCancel
    End If
    On Error GoTo 0
    If originalNumber <> 0 Then
        message = message & vbCrLf & "The working range may be partially changed. No recovery workbook was created."
        Err.Raise originalNumber, "SheetFreak", message
    End If
End Sub

Public Sub SF_Help(Optional control As Variant)
    MsgBox "SheetFreak 0.2.1" & vbCrLf & vbCrLf & _
        "Ctrl+Alt+0  Zero decimals (preserves number style)" & vbCrLf & _
        "Ctrl+Alt+C  Copy exact formulas" & vbCrLf & _
        "Ctrl+Alt+V  Paste exact formulas" & vbCrLf & _
        "Ctrl+Alt+H  Auto resize (A1 and 70-200% zoom)" & vbCrLf & _
        "Ctrl+Alt+A  Assumption format" & vbCrLf & _
        "Ctrl+Alt+T  Header format" & vbCrLf & _
        "Ctrl+Alt+F  Fill and freeze (confirmation required)" & vbCrLf & vbCrLf & _
        "Fill and freeze: select one row of formulas, then choose a last-row column." & vbCrLf & _
        "No recovery workbook is created. Save your work before fill and freeze." & vbCrLf & _
        "Use the ribbon if your keyboard layout or another add-in conflicts." & vbCrLf & _
        "VBA edits clear normal Undo. No files are automatically saved.", vbInformation, "SheetFreak"
End Sub
