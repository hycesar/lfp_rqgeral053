Attribute VB_Name = "performance"
Public Sub speedup(status As Boolean)
    On Error Resume Next
    With Application
        If status = True Then
                .ScreenUpdating = False
                .FeatureInstall = msoFeatureInstallOnDemandWithUI
                .EnableLivePreview = False
                .EnableEvents = False
                .DisplayStatusBar = False
                .DisplayScrollBars = False
                .DisplayFormulaBar = False
                .DisplayAlerts = False
                .Calculation = xlCalculationManual
                .AlertBeforeOverwriting = False
                .ActiveSheet.DisplayPageBreaks = False
        Else
                .ActiveSheet.DisplayPageBreaks = True
                .AlertBeforeOverwriting = True
                .Calculation = xlCalculationAutomatic
                .DisplayAlerts = True
                .DisplayFormulaBar = True
                .DisplayScrollBars = True
                .DisplayStatusBar = True
                .EnableEvents = True
                .EnableLivePreview = True
                .FeatureInstall = msoFeatureInstallNone
                .ScreenUpdating = True
        End If
    End With
End Sub
