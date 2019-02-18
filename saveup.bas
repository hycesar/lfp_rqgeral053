Attribute VB_Name = "saveup"
Public Sub CustomSave(Optional SaveAs As Boolean)
    Dim ws As Worksheet, aWs As Worksheet, newFname As String
    
    'Record active worksheet
    Set aWs = ActiveSheet
    
    'Hide all sheets
    Call s_hideallsheets
    
    'Save workbook directly or prompt for saveas filename
    If SaveAs = True Then
        newFname = Application.GetSaveAsFilename(fileFilter:="Excel Files (*.xlsm), *.xlsm")
        If Not newFname = "Falso" Then ThisWorkbook.SaveAs newFname
    Else
        ThisWorkbook.Save
    End If
    
    'Restore file to where user was
    Call s_showallsheets
    aWs.Activate

End Sub
