Attribute VB_Name = "showup"
Public Sub s_showallsheets()
    'Code
    Worksheets("Teor").Visible = xlSheetVisible
    Worksheets("Uniformidade DE Conte�do").Visible = xlSheetVisible
    Worksheets("Dissolu��o").Visible = xlSheetVisible
    Worksheets("Teor").Activate
    Worksheets("Macros").Visible = xlSheetVeryHidden
End Sub
Public Sub s_hideallsheets()
    'Code
    Worksheets("Macros").Visible = xlSheetVisible
    Worksheets("Macros").Activate
    Worksheets("Teor").Visible = xlSheetVeryHidden
    Worksheets("Uniformidade DE Conte�do").Visible = xlSheetVeryHidden
    Worksheets("Dissolu��o").Visible = xlSheetVeryHidden
End Sub
