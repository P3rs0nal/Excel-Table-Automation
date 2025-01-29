VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub WorkBook_Open()
    ufProg.StartUpPosition = 0
    ufProg.Left = Application.Left + (0.5 * Application.Width) - (0.5 * ufProg.Width)
    ufProg.Top = Application.Top + (0.5 * Application.Height) - (0.5 * ufProg.Height)
    ufProg.Show
    ufProg.LabelProg.Width = 0
    ufProg.LabelCaption.Caption = "Refreshing Data, Please Wait..."
    ThisWorkbook.RefreshAll
    Unload ufProg
End Sub
