Attribute VB_Name = "IntegrityCheck"
Private Sub Workbook_BeforePrint(Cancel As Boolean)
    If IsEmpty(Worksheets("Teor").Range("AF23")) Or IsEmpty(Worksheets("Teor").Range("AF25")) Or IsEmpty(Worksheets("Teor").Range("AM7")) Then
        MsgBox "Por favor, verifique a planilha antes de imprimir pois há campos não preenchidos!"
    Else
        If Worksheets("Teor").Range("AF23") / Worksheets("Teor").Range("AF23") = 1 Then
            MsgBox "Esta igual!"
        End If
    End If
End Sub

