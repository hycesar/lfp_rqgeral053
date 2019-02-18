Attribute VB_Name = "protectup"
Option Explicit
Public Sub ToggleCutCopyAndPaste(Allow As Boolean)
     'Activate/deactivate cut, copy, paste and pastespecial menu items
    Call EnableMenuItem(21, Allow) ' cut
    Call EnableMenuItem(19, Allow) ' copy
    Call EnableMenuItem(22, Allow) ' paste
    Call EnableMenuItem(755, Allow) ' pastespecial
     
     'Activate/deactivate drag and drop ability
    Application.CellDragAndDrop = Allow
     
     'Activate/deactivate cut, copy, paste and pastespecial shortcut keys
    With Application
        Select Case Allow
        Case Is = False
            .OnKey "^c", "CutCopyPasteDisabled"
            .OnKey "^v", "CutCopyPasteDisabled"
            .OnKey "^x", "CutCopyPasteDisabled"
            .OnKey "+{DEL}", "CutCopyPasteDisabled"
            .OnKey "^{INSERT}", "CutCopyPasteDisabled"
        Case Is = True
            .OnKey "^c"
            .OnKey "^v"
            .OnKey "^x"
            .OnKey "+{DEL}"
            .OnKey "^{INSERT}"
        End Select
    End With
End Sub
Public Sub EnableMenuItem(ctlId As Integer, Enabled As Boolean)
     'Activate/Deactivate specific menu item
    Dim cBar As CommandBar
    Dim cBarCtrl As CommandBarControl
    For Each cBar In Application.CommandBars
        If cBar.Name <> "Clipboard" Then
            Set cBarCtrl = cBar.FindControl(ID:=ctlId, recursive:=True)
            If Not cBarCtrl Is Nothing Then cBarCtrl.Enabled = Enabled
        End If
    Next
End Sub
Public Sub CutCopyPasteDisabled()
     'Inform user that the functions have been disabled
    MsgBox "Desculpe. Mas, por segurança, operação de copiar foi desabilitada!"
End Sub
Public Function f_protect(ws As Worksheet)
    'Protect
    ws.Unprotect Password:="0"
    ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:="0"
    ws.EnableSelection = xlUnlockedCells
End Function
