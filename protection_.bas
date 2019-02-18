Attribute VB_Name = "protection"
Public Sub lockup(status As Boolean)
    On Error Resume Next
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        With ws
        .Unprotect Password:=.Name
        End With
    Next ws
    If status = True Then
        For Each ws In ThisWorkbook.Worksheets
            With ws
            .EnableSelection = xlUnlockedCells
            .Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=.Name
            End With
        Next ws
    End If
End Sub
