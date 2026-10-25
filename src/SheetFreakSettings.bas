Attribute VB_Name = "SheetFreakSettings"
Option Explicit

Public Sub Auto_Add()
    SF_RegisterKeys
End Sub

Public Sub Auto_Remove()
    SF_UnregisterKeys
End Sub

Public Sub SF_EnableKeys(Optional control As Variant)
    SF_RegisterKeys
    MsgBox "SheetFreak shortcuts are enabled for this Excel session.", vbInformation, "SheetFreak"
End Sub

Public Sub SF_DisableKeys(Optional control As Variant)
    SF_UnregisterKeys
    MsgBox "SheetFreak shortcuts are disabled for this session. The ribbon buttons still work." & vbCrLf & "Shortcuts are enabled again when SheetFreak next loads.", vbInformation, "SheetFreak"
End Sub
