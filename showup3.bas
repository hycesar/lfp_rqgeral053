Attribute VB_Name = "showup"
Public Sub s_showallsheets()
    'Code
    Worksheets("Teor").Visible = xlSheetVisible
    Worksheets("Uniformidade DE Conteúdo").Visible = xlSheetVisible
    Worksheets("Dissolução").Visible = xlSheetVisible
    Worksheets("Teor").Activate
    Worksheets("Macros").Visible = xlSheetVeryHidden
End Sub
Public Sub s_hideallsheets()
    'Code
    Worksheets("Macros").Visible = xlSheetVisible
    Worksheets("Macros").Activate
    Worksheets("Teor").Visible = xlSheetVeryHidden
    Worksheets("Uniformidade DE Conteúdo").Visible = xlSheetVeryHidden
    Worksheets("Dissolução").Visible = xlSheetVeryHidden
End Sub
