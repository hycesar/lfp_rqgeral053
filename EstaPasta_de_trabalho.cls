VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EstaPasta_de_trabalho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Call performance.speedup(True)
    With ThisWorkbook
        If Not .Saved Then
            Select Case MsgBox("Salvar '" & .Name & "' ?", vbYesNo + vbExclamation, "LAFEPE S/A")
                Case Is = vbYes
                    Cancel = False
                    Call CustomSave
                Case Is = vbNo
                    Cancel = True
            End Select
        End If
        Call performance.speedup(False)
        If Cancel = False Then
            .Saved = True
            Application.Quit
        End If
    End With
End Sub
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    Call performance.speedup(True)
    Call CustomSave
    Cancel = True
    ThisWorkbook.Saved = True
    Call performance.speedup(False)
End Sub
Private Sub CustomSave()
    On Error Resume Next
    Dim aWs As Worksheet, newFname As String
    Set aWs = ActiveSheet
    Call HideAllSheets
        Dim a, b As String
        If Dir(ActiveWorkbook.Path & "\backup de liberações", vbDirectory) = "" Then
            MkDir ActiveWorkbook.Path + "\backup de liberações"
        Else
            ChDir ActiveWorkbook.Path + "\backup de liberações"
        End If
        newFname = "RQ GERAL 053 REV 003 " + Format(Now(), "yymmddhhmmssms") + ".xlsm"
        If Not newFname = "Falso" Then ThisWorkbook.SaveAs newFname
    Call ShowAllSheets
    aWs.Activate
End Sub
Private Sub Workbook_Open()
    Call performance.speedup(True)
    Call ShowAllSheets
    Call performance.speedup(False)
End Sub
Private Sub HideAllSheets()
    Dim ws As Worksheet
    With Worksheets("HABILITE MACROS")
        .Visible = xlSheetVisible
        .Activate
    End With
    For Each ws In ThisWorkbook.Worksheets
        If Not ws.Name = "HABILITE MACROS" Then ws.Visible = xlSheetVeryHidden
    Next ws
End Sub
Private Sub ShowAllSheets()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = "Uniformidade DE Conteúdo" Or ws.Name = "Dissolução" Or ws.Name = "Teor" Then ws.Visible = xlSheetVisible
    Next ws
    With Worksheets("HABILITE MACROS")
        .Visible = xlSheetVeryHidden
    End With
    Call protection.lockup(True)
End Sub
