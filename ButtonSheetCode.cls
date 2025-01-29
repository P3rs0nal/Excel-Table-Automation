VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
    Application.ScreenUpdating = False
    'ThisWorkbook.RefreshAll
    If buttonValue = 1 Then
        'change to blue for word
        ActiveSheet.CommandButton1.BackColor = RGB(173, 216, 230)
        ActiveSheet.CommandButton1.Caption = "Word"
    ElseIf buttonValue = 2 Then
        'change to yellow for email draft
        ActiveSheet.CommandButton1.BackColor = RGB(255, 255, 0)
        ActiveSheet.CommandButton1.Caption = "Email"
    Else
        'change to green for direct email send
        ActiveSheet.CommandButton1.BackColor = RGB(153, 255, 102)
        ActiveSheet.CommandButton1.Caption = "Direct Email Send"
    End If
    buttonValue = (buttonValue + 1) Mod 3
    ThisWorkbook.Sheets("Email Generator").Range("A31").Value = buttonValue
    Application.ScreenUpdating = True
End Sub
