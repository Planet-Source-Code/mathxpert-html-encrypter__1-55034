Attribute VB_Name = "basMain"
Sub Main()
If Screen.Width < 800 And Screen.Width < 600 Then
    MsgBox "Please set your screen resolution to 800x600 or higher.", 48, "Screen Resolution Too Small"
    End
Else
    frmMain.Show
End If
End Sub
