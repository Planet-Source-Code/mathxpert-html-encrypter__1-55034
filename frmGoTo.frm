VERSION 5.00
Begin VB.Form frmGoTo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Go To"
   ClientHeight    =   960
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   960
   ScaleWidth      =   3600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1853
      TabIndex        =   3
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   773
      TabIndex        =   2
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox txtLine 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label lblLine 
      Caption         =   "&Line Number:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmGoTo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
On Error GoTo hndl
If compF Then If CLng(txtLine) > SendMessage(frmMain.txtCompiled.hWnd, EM_GETLINECOUNT, 0, 0) Or CLng(txtLine) < 1 Then GoTo hndl
If origF Then If CLng(txtLine) > SendMessage(frmMain.txtOriginal.hWnd, EM_GETLINECOUNT, 0, 0) Or CLng(txtLine) < 1 Then GoTo hndl
tmpPos = CLng(txtLine)
frmMain.GoToBM
Unload Me
Exit Sub
hndl:
    MsgBox "Line number out of range", , "Go To"
    txtLine.SetFocus
    txtLine.SelStart = 0
    txtLine.SelLength = Len(txtLine)
End Sub

Private Sub Form_Load()
Left = frmMain.Left + 180
Top = frmMain.Top + 1110
txtLine = tmpText
txtLine.SelStart = 0
txtLine.SelLength = Len(txtLine)
End Sub
