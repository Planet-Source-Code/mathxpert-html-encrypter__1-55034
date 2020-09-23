VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Untitled - HTML Encrypter"
   ClientHeight    =   7080
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   6720
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7080
   ScaleWidth      =   6720
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDecompile 
      Caption         =   "Decompile"
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmdCompile 
      Caption         =   "Compile"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmdClearAll 
      Caption         =   "Clear All"
      Height          =   375
      Left            =   5400
      TabIndex        =   6
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmdClearCompiled 
      Caption         =   "Clear Compiled"
      Height          =   375
      Left            =   4080
      TabIndex        =   5
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmdClearOriginal 
      Caption         =   "Clear Original"
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox txtCompiled 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   8
      Top             =   4080
      Width           =   6495
   End
   Begin VB.TextBox txtOriginal 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   360
      Width           =   6495
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   5760
      Top             =   0
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   6240
      Top             =   0
   End
   Begin VB.Label lblStat 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Line: 1        Col: 1"
      Height          =   255
      Left            =   4440
      TabIndex        =   9
      Top             =   0
      Width           =   2175
   End
   Begin VB.Line Line5 
      X1              =   6000
      X2              =   6000
      Y1              =   3240
      Y2              =   4080
   End
   Begin VB.Line Line4 
      X1              =   4680
      X2              =   4680
      Y1              =   3720
      Y2              =   4080
   End
   Begin VB.Line Line3 
      X1              =   3360
      X2              =   3360
      Y1              =   3240
      Y2              =   3360
   End
   Begin VB.Line Line2 
      X1              =   2040
      X2              =   2040
      Y1              =   3720
      Y2              =   4080
   End
   Begin VB.Line Line1 
      X1              =   720
      X2              =   720
      Y1              =   3240
      Y2              =   3360
   End
   Begin VB.Label lblCompiled 
      Caption         =   "Compiled:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   3840
      Width           =   735
   End
   Begin VB.Label lblOriginal 
      Caption         =   "Original:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnuLine 
         Caption         =   "-"
      End
      Begin VB.Menu mnuImport 
         Caption         =   "&Import HTML Document..."
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuExport 
         Caption         =   "&Export Compiled Code..."
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuUndo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuLine3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "De&lete"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuLine4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGoTo 
         Caption         =   "&Go To..."
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuLine5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelAll 
         Caption         =   "Select &All"
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const APPNM = "HTMLEncrypter"
Private Const FTR = "HTML Encrypter Files (*.enc)|*.enc"
Private Const FTR2 = "HTML Documents (*.htm, *.html, *.shtm, *.shtml, *.jhtml, *.htx, *.alx, *.stm, *.xhtm, *.xhtml, *.ssi, *.lbi)|*.htm;*.html;*.shtm;*.shtml;*.jhtml;*.htx;*.alx;*.stm;*.xhtm;*.xhtml;*.ssi;*.lbi"

Dim tmpFN As String, Saved As Boolean, tmpX As Single, tmpY As Single, tmpL As Single, tmpW As Single, tmpSt1 As Long, tmpSt2 As Long, tmpOrig1 As Boolean, tmpOrig2 As Boolean, tmpComp1 As Boolean, tmpComp2 As Boolean

Private Function AddNull(HexVal As String) As String
AddNull = IIf(Len(HexVal) = 1, "0" & HexVal, HexVal)
End Function

Private Function DelNull(HexVal As String) As String
DelNull = IIf(Left(HexVal, 1) = "0", Right(HexVal, 1), HexVal)
End Function

Private Function RetBool(StringCheck As String, StringMatch As String) As Boolean
RetBool = (InStrRev(StringCheck, StringMatch) > 0)
End Function

Private Sub SaveChanges()
DoSaveAs (tmpFN = "")
End Sub

Public Sub GoToBM()
Dim tmpPosB As Long
tmpPos = tmpPos - 1
If origF Then
    tmpPosB = SendMessage(txtOriginal.hWnd, EM_LINEINDEX, tmpPos, 0)
    SendMessage txtOriginal.hWnd, EM_SETSEL, tmpPosB, tmpPosB
    SendMessage txtOriginal.hWnd, EM_SCROLLCARET, 0, 0
End If
If compF Then
    tmpPosB = SendMessage(txtCompiled.hWnd, EM_LINEINDEX, tmpPos, 0)
    SendMessage txtCompiled.hWnd, EM_SETSEL, tmpPosB, tmpPosB
    SendMessage txtCompiled.hWnd, EM_SCROLLCARET, 0, 0
End If
End Sub

Private Function CanUndo(TBox As TextBox) As Boolean
CanUndo = (SendMessage(TBox.hWnd, EM_CANUNDO, 0, 0) <> 0)
End Function

Private Sub SaveStg()
SaveSetting APPNM, "WindowState", "Height", tmpL
SaveSetting APPNM, "WindowState", "Width", tmpW
SaveSetting APPNM, "WindowState", "Left", tmpX
SaveSetting APPNM, "WindowState", "Top", tmpY
If WindowState <> 1 Then SaveSetting APPNM, "WindowState", "CurrState", WindowState
Unload frmGoTo
Unhook
If Not (cmdCompile.Enabled Or cmdDecompile.Enabled) Then End
If Not (cmdCompile.Enabled And cmdDecompile.Enabled) Then End
End Sub

Private Sub GetTBPos(TBox As TextBox)
On Error Resume Next
Dim tmpPos1 As Long, tmpPos2 As Long, cLine As Long, cChar As Long

tmpPos1 = SendMessage(TBox.hWnd, EM_GETSEL, 0, 0) / &H10000
cLine = SendMessage(TBox.hWnd, EM_LINEFROMCHAR, tmpPos1, 0)

If tmpText <> CStr(cLine + 1) Then tmpText = CStr(cLine + 1)

tmpPos2 = SendMessageLong(TBox.hWnd, EM_LINEINDEX, cLine, 0)
cChar = tmpPos1 - tmpPos2

If lblStat.Caption <> "Line: " & CStr(cLine + 1) & TB & "Col: " & CStr(cChar + 1) Then lblStat.Caption = "Line: " & CStr(cLine + 1) & TB & "Col: " & CStr(cChar + 1)
End Sub

Private Function GetSize(Optional UseWidth As Boolean = True) As Single
Dim tX As String, tY As String
If UseWidth Then
    tX = GetSetting(APPNM, "WindowState", "Width", 6840)
    If Not IsNumeric(tX) Then GetSize = 6840 Else GetSize = CSng(tX)
Else
    tY = GetSetting(APPNM, "WindowState", "Height", 7590)
    If Not IsNumeric(tY) Then GetSize = 7590 Else GetSize = CSng(tY)
End If
End Function

Private Function GetPos(Optional UseLeft As Boolean = True) As Single
Dim tL As String, tT As String
If UseLeft Then
    tL = GetSetting(APPNM, "WindowState", "Left", Left)
    If Not IsNumeric(tL) Then GetPos = Left Else GetPos = CSng(tL)
Else
    tT = GetSetting(APPNM, "WindowState", "Top", Top)
    If Not IsNumeric(tT) Then GetPos = Top Else GetPos = CSng(tT)
End If
End Function

Private Function GetSt() As Single
Dim tmp As String
tmp = GetSetting(APPNM, "WindowState", "CurrState", 0)
If Not IsNumeric(tmp) Then GetSt = 0 Else GetSt = CSng(tmp)
End Function

Private Sub DoOpen()
Dim fn As String, ff As Integer, ln As String, orig As Boolean, ch As Boolean, str1 As String, str2 As String
ff = FreeFile
fn = "": str1 = "": str2 = ""
fn = ShowOpenDlg(Me, FTR, "Open")
If fn <> "" Then
    If Dir(fn) = "" Then
        tmpFN = ""
    Else
        tmpFN = fn
        Open tmpFN For Input As #ff
            Do Until EOF(ff)
                Line Input #ff, ln
                If InStrRev(ln, "Orig" & Chr(21) & Chr(12) & Chr(15) & Chr(20) & Chr(8)) > 0 Then orig = True: ch = True
                If InStrRev(ln, "Comp" & Chr(24) & Chr(15) & Chr(16) & Chr(6) & Chr(2)) > 0 Then orig = False: ch = True
                If orig Then
                    If ch Then
                        ch = False
                    Else
                        str1 = str1 & ln & vbCrLf
                    End If
                Else
                    If ch Then
                        ch = False
                    Else
                        str2 = str2 & ln & vbCrLf
                    End If
                End If
            Loop
        Close #ff
        txtOriginal = Left(str1, Len(str1) - 2)
        txtCompiled = Left(str2, Len(str2) - 2)
        Caption = Replace(Dir(fn), ".enc", "") & " - HTML Encrypter"
    End If
End If
End Sub

Private Sub DoSave(strFN As String)
Dim ff As Integer
ff = FreeFile
Open strFN For Output As #ff
    Print #ff, "Orig" & Chr(21) & Chr(12) & Chr(15) & Chr(20) & Chr(8)
    Print #ff, txtOriginal
    Print #ff, "Comp" & Chr(24) & Chr(15) & Chr(16) & Chr(6) & Chr(2)
    Print #ff, txtCompiled
Close #ff
Caption = Replace(Dir(strFN), ".enc", "") & " - HTML Encrypter"
tmpFN = strFN
End Sub

Private Sub DoSaveAs(Optional ShowDlg As Boolean = True)
On Error Resume Next
Dim fn As String, ff As Integer, ln As String, orig As Boolean, ch As Boolean, str1 As String, str2 As String
If ShowDlg Then
    ff = FreeFile
    fn = "": str1 = "": str2 = ""
    fn = ShowSaveDlg(Me, FTR, "Save As", , "*.enc")
    If fn = "" Then
        Saved = False
    Else
        DoSave fn
    End If
Else
    DoSave tmpFN
    Saved = True
End If
End Sub

Private Sub cmdClearAll_Click()
If txtOriginal <> "" Then txtOriginal = ""
If txtCompiled <> "" Then txtCompiled = ""
End Sub

Private Sub cmdClearAll_GotFocus()
origF = False
compF = False
End Sub

Private Sub cmdClearCompiled_Click()
If txtCompiled <> "" Then txtCompiled = ""
End Sub

Private Sub cmdClearCompiled_GotFocus()
origF = False
compF = False
End Sub

Private Sub cmdClearOriginal_Click()
If txtOriginal <> "" Then txtOriginal = ""
End Sub

Private Sub cmdClearOriginal_GotFocus()
origF = False
compF = False
End Sub

Private Sub cmdCompile_Click()
On Error Resume Next
Dim i As Long, tmp As String, tmp2 As String
txtCompiled.SetFocus
cmdCompile.Enabled = False
tmp = txtOriginal
If tmp <> "" And InStrRev(tmp, "<script language=""JavaScript"">" & vbCrLf & "document.write(unescape(""\x") = 0 Then
    tmp2 = "<script language=""JavaScript"">" & vbCrLf & "document.write(unescape("""
    For i = 1 To Len(tmp)
        tmp2 = tmp2 & "\x" & LCase(AddNull(Hex(Asc(Mid(tmp, i, 1)))))
        If txtCompiled <> "Please wait..." Then txtCompiled = "Please wait..."
        DoEvents
    Next
    tmp2 = tmp2 & """))" & vbCrLf & "</script>" & vbCrLf & "<noscript>JavaScript is required to view this page.</noscript>"
    txtCompiled = tmp2
End If
cmdCompile.Enabled = True
End Sub

Private Sub cmdCompile_GotFocus()
origF = False
compF = False
End Sub

Private Sub cmdDecompile_Click()
On Error Resume Next
Dim i As Long, j As Long, tmp As String, tmp2 As String
txtOriginal.SetFocus
cmdDecompile.Enabled = False
tmp = txtCompiled
If tmp <> "" And InStrRev(tmp, "<script language=""JavaScript"">" & vbCrLf & "document.write(unescape(""\x") > 0 Then
    For i = 1 To Len(tmp)
        If Mid(tmp, i, 14) = """))" & vbCrLf & "</script>" Then Exit For
        i = InStr(i, tmp, "\x")
        If i > 0 Then
            For j = 0 To 255
                If Hex(j) = UCase(DelNull(Mid(tmp, i + 2, 2))) Then
                    tmp2 = tmp2 & Chr(j)
                    Exit For
                End If
            Next
            i = i + 1
        Else
            Exit For
        End If
        If txtOriginal <> "Please wait..." Then txtOriginal = "Please wait..."
        DoEvents
    Next
    txtOriginal = tmp2
End If
cmdDecompile.Enabled = True
End Sub

Private Sub cmdDecompile_GotFocus()
origF = False
compF = False
End Sub

Private Sub Form_Load()
Width = GetSize
Height = GetSize(False)
Left = GetPos
Top = GetPos(False)
WindowState = GetSt
origF = True
compF = False
gHW = Me.hWnd
Hook
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim ff As Integer, str1 As String, str2 As String, ln As String, orig As Boolean, ch As Boolean
ff = FreeFile
Saved = True
If tmpFN = "" Then
    If txtCompiled <> "" Then
        Select Case MsgBox("Do you want to save changes?", 35, "Changes Detected")
            Case vbYes
                SaveChanges
                GoTo hndl
            Case vbNo
                SaveStg
                Exit Sub
            Case Else
                Cancel = 1
                Exit Sub
        End Select
    End If
    If txtOriginal <> "" Then
        Select Case MsgBox("Do you want to save changes?", 35, "Changes Detected")
            Case vbYes
                SaveChanges
                GoTo hndl
            Case vbNo
                SaveStg
                Exit Sub
            Case Else
                Cancel = 1
                Exit Sub
        End Select
    End If
Else
    Open tmpFN For Input As #ff
        Do Until EOF(ff)
            Line Input #ff, ln
            If InStrRev(ln, "Orig" & Chr(21) & Chr(12) & Chr(15) & Chr(20) & Chr(8)) > 0 Then orig = True: ch = True
            If InStrRev(ln, "Comp" & Chr(24) & Chr(15) & Chr(16) & Chr(6) & Chr(2)) > 0 Then orig = False: ch = True
            If orig Then
                If ch Then
                    ch = False
                Else
                    str1 = str1 & ln & vbCrLf
                End If
            Else
                If ch Then
                    ch = False
                Else
                    str2 = str2 & ln & vbCrLf
                End If
            End If
        Loop
    Close #ff
    If txtOriginal <> Left(str1, Len(str1) - 2) Then
        Select Case MsgBox("Do you want to save changes?", 35, "Changes Detected")
            Case vbYes
                SaveChanges
                GoTo hndl
            Case vbNo
                SaveStg
                Exit Sub
            Case Else
                Cancel = 1
                Exit Sub
        End Select
    End If
    If txtCompiled <> Left(str2, Len(str2) - 2) Then
        Select Case MsgBox("Do you want to save changes?", 35, "Changes Detected")
            Case vbYes
                SaveChanges
                GoTo hndl
            Case vbNo
                SaveStg
                Exit Sub
            Case Else
                Cancel = 1
                Exit Sub
        End Select
    End If
End If
hndl:
If Saved Then
    SaveStg
Else
    Cancel = 1
End If
End Sub

Private Sub Form_Resize()
On Error Resume Next
Dim bWidth As Single
If WindowState <> 1 Then
    lblStat.Left = ScaleWidth - 2295
    txtOriginal.Height = ScaleHeight - (ScaleHeight / 2) - 653
    txtCompiled.Height = ScaleHeight - (ScaleHeight / 2) - 653
    txtOriginal.Width = ScaleWidth - 240
    txtCompiled.Width = ScaleWidth - 240
    With txtOriginal
        bWidth = (.Width / 5) - 95
        Line1.Y1 = .Height + 360
        Line1.Y2 = .Height + 480
        Line3.Y1 = .Height + 360
        Line3.Y2 = .Height + 480
        cmdCompile.Top = .Height + 480
        cmdCompile.Width = bWidth
        cmdDecompile.Top = .Height + 480
        cmdDecompile.Left = bWidth + 240
        cmdDecompile.Width = bWidth
        cmdClearOriginal.Top = .Height + 480
        cmdClearOriginal.Left = (bWidth * 2) + 360
        cmdClearOriginal.Width = bWidth
        cmdClearCompiled.Top = .Height + 480
        cmdClearCompiled.Left = (bWidth * 3) + 480
        cmdClearCompiled.Width = bWidth
        cmdClearAll.Top = .Height + 480
        cmdClearAll.Left = (bWidth * 4) + 600
        cmdClearAll.Width = bWidth
        lblCompiled.Top = cmdCompile.Top + cmdCompile.Height + 120
        txtCompiled.Top = .Height + 1200
        Line5.Y1 = .Height + 360
        Line5.Y2 = txtCompiled.Top
        Line2.Y1 = cmdDecompile.Height + .Height + 480
        Line2.Y2 = cmdDecompile.Height + .Height + 840
        Line4.Y1 = cmdClearCompiled.Height + .Height + 480
        Line4.Y2 = cmdClearCompiled.Height + .Height + 840
        Line1.X1 = (bWidth / 2) + 120
        Line1.X2 = (bWidth / 2) + 120
        Line2.X1 = (bWidth + (bWidth / 2)) + 240
        Line2.X2 = (bWidth + (bWidth / 2)) + 240
        Line3.X1 = ((bWidth * 2) + (bWidth / 2)) + 360
        Line3.X2 = ((bWidth * 2) + (bWidth / 2)) + 360
        Line4.X1 = ((bWidth * 3) + (bWidth / 2)) + 480
        Line4.X2 = ((bWidth * 3) + (bWidth / 2)) + 480
        Line5.X1 = ((bWidth * 4) + (bWidth / 2)) + 600
        Line5.X2 = ((bWidth * 4) + (bWidth / 2)) + 600
    End With
End If
End Sub

Private Sub mnuCopy_Click()
If compF Then SendMessage txtCompiled.hWnd, WM_COPY, 0, 0
If origF Then SendMessage txtOriginal.hWnd, WM_COPY, 0, 0
End Sub

Private Sub mnuCut_Click()
If compF Then SendMessage txtCompiled.hWnd, WM_CUT, 0, 0
If origF Then SendMessage txtOriginal.hWnd, WM_CUT, 0, 0
End Sub

Private Sub mnuDelete_Click()
If compF Then SendMessage txtCompiled.hWnd, WM_CLEAR, 0, 0
If origF Then SendMessage txtOriginal.hWnd, WM_CLEAR, 0, 0
End Sub

Private Sub mnuEdit_Click()
Dim TextSelected As Boolean
Dim SelGTZero As Boolean

TextSelected = (txtCompiled.SelLength > 0 Or txtOriginal.SelLength > 0)
SelGTZero = (Len(txtCompiled) > 0 Or Len(txtOriginal) > 0)

If origF Then If CanUndo(txtOriginal) Then mnuUndo.Enabled = True: mnuUndo.Caption = "&Undo" Else mnuUndo.Enabled = False: mnuUndo.Caption = "Can't Undo"
If compF Then If CanUndo(txtCompiled) Then mnuUndo.Enabled = True: mnuUndo.Caption = "&Undo" Else mnuUndo.Enabled = False: mnuUndo.Caption = "Can't Undo"

If Not origF And Not compF Then
    mnuUndo.Enabled = False
    mnuUndo.Caption = "Can't Undo"
    mnuDelete.Enabled = False
    mnuCut.Enabled = False
    mnuCopy.Enabled = False
    mnuPaste.Enabled = False
    mnuGoTo.Enabled = False
    mnuSelAll.Enabled = False
Else
    If origF Then If CanUndo(txtOriginal) Then mnuUndo.Enabled = True: mnuUndo.Caption = "&Undo" Else mnuUndo.Enabled = False: mnuUndo.Caption = "Can't Undo"
    If compF Then If CanUndo(txtCompiled) Then mnuUndo.Enabled = True: mnuUndo.Caption = "&Undo" Else mnuUndo.Enabled = False: mnuUndo.Caption = "Can't Undo"
    mnuDelete.Enabled = TextSelected
    mnuCut.Enabled = TextSelected
    mnuCopy.Enabled = TextSelected
    mnuPaste.Enabled = Clipboard.GetFormat(vbCFText)
    mnuGoTo.Enabled = True
    If origF Then mnuSelAll.Enabled = (SelGTZero And txtOriginal.SelLength <> Len(txtOriginal))
    If compF Then mnuSelAll.Enabled = (SelGTZero And txtCompiled.SelLength <> Len(txtCompiled))
End If
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub mnuExport_Click()
On Error Resume Next
Dim ff As Integer, fn As String
ff = FreeFile
fn = ShowSaveDlg(Me, FTR2, "Export Compiled Code", , "*.htm")
If fn <> "" Then
    Open fn For Output As #ff
        Print #ff, txtCompiled
    Close #ff
End If
End Sub

Private Sub mnuFile_Click()
mnuImport.Enabled = origF Or compF
mnuExport.Enabled = (txtCompiled <> "")
End Sub

Private Sub mnuGoTo_Click()
frmGoTo.Show 1
End Sub

Private Sub mnuImport_Click()
On Error Resume Next
Dim ff As Integer, fn As String, tmpHTML As String, IsValid As Boolean

IsValid = False
ff = FreeFile
fn = ShowOpenDlg(Me, FTR2, "Import HTML Document")

If fn <> "" Then

    If RetBool(fn, ".htm") Then IsValid = True
    If RetBool(fn, ".html") Then IsValid = True
    If RetBool(fn, ".shtm") Then IsValid = True
    If RetBool(fn, ".shtml") Then IsValid = True
    If RetBool(fn, ".jhtml") Then IsValid = True
    If RetBool(fn, ".htx") Then IsValid = True
    If RetBool(fn, ".alx") Then IsValid = True
    If RetBool(fn, ".stm") Then IsValid = True
    If RetBool(fn, ".xhtm") Then IsValid = True
    If RetBool(fn, ".xhtml") Then IsValid = True
    If RetBool(fn, ".ssi") Then IsValid = True
    If RetBool(fn, ".lbi") Then IsValid = True

    If IsValid Then
        Open fn For Input As #ff
            tmpHTML = Input(LOF(ff), #ff)
        Close #ff
        If origF Then txtOriginal.SelText = tmpHTML
        If compF Then txtCompiled.SelText = tmpHTML
    Else
        MsgBox "Invalid file type.", 48, "Error"
    End If

End If
End Sub

Private Sub mnuNew_Click()
On Error Resume Next
Dim ff As Integer, str1 As String, str2 As String, ln As String, orig As Boolean, ch As Boolean
ff = FreeFile
Saved = True
If tmpFN = "" Then
    If txtCompiled <> "" Then
        Select Case MsgBox("Do you want to save changes?", 35, "Changes Detected")
            Case vbYes
                SaveChanges
                GoTo hndl
            Case vbNo
                txtCompiled = ""
                txtOriginal = ""
                tmpFN = ""
                Caption = "Untitled - HTML Encrypter"
                Exit Sub
            Case Else
                Exit Sub
        End Select
    End If
    If txtOriginal <> "" Then
        Select Case MsgBox("Do you want to save changes?", 35, "Changes Detected")
            Case vbYes
                SaveChanges
                GoTo hndl
            Case vbNo
                txtCompiled = ""
                txtOriginal = ""
                tmpFN = ""
                Caption = "Untitled - HTML Encrypter"
                Exit Sub
            Case Else
                Exit Sub
        End Select
    End If
Else
    Open tmpFN For Input As #ff
        Do Until EOF(ff)
            Line Input #ff, ln
            If InStrRev(ln, "Orig" & Chr(21) & Chr(12) & Chr(15) & Chr(20) & Chr(8)) > 0 Then orig = True: ch = True
            If InStrRev(ln, "Comp" & Chr(24) & Chr(15) & Chr(16) & Chr(6) & Chr(2)) > 0 Then orig = False: ch = True
            If orig Then
                If ch Then
                    ch = False
                Else
                    str1 = str1 & ln & vbCrLf
                End If
            Else
                If ch Then
                    ch = False
                Else
                    str2 = str2 & ln & vbCrLf
                End If
            End If
        Loop
    Close #ff
    If txtOriginal <> Left(str1, Len(str1) - 2) Then
        Select Case MsgBox("Do you want to save changes?", 35, "Changes Detected")
            Case vbYes
                SaveChanges
                GoTo hndl
            Case vbNo
                txtCompiled = ""
                txtOriginal = ""
                tmpFN = ""
                Caption = "Untitled - HTML Encrypter"
                Exit Sub
            Case Else
                Exit Sub
        End Select
    End If
    If txtCompiled <> Left(str2, Len(str2) - 2) Then
        Select Case MsgBox("Do you want to save changes?", 35, "Changes Detected")
            Case vbYes
                SaveChanges
                GoTo hndl
            Case vbNo
                txtCompiled = ""
                txtOriginal = ""
                tmpFN = ""
                Caption = "Untitled - HTML Encrypter"
                Exit Sub
            Case Else
                Exit Sub
        End Select
    End If
End If
hndl:
If Saved Then
    txtCompiled = ""
    txtOriginal = ""
    tmpFN = ""
    Caption = "Untitled - HTML Encrypter"
End If
End Sub

Private Sub mnuOpen_Click()
On Error Resume Next
Dim ff As Integer
ff = FreeFile
Saved = True
If tmpFN = "" Then
    If txtCompiled <> "" Then
        Select Case MsgBox("Do you want to save changes?", 35, "Changes Detected")
            Case vbYes
                SaveChanges
                GoTo hndl
            Case vbNo
                DoOpen
                Exit Sub
            Case Else
                Exit Sub
        End Select
    End If
    If txtOriginal <> "" Then
        Select Case MsgBox("Do you want to save changes?", 35, "Changes Detected")
            Case vbYes
                SaveChanges
                GoTo hndl
            Case vbNo
                DoOpen
                Exit Sub
            Case Else
                Exit Sub
        End Select
    End If
Else
    Open tmpFN For Input As #ff
        Do Until EOF(ff)
            Line Input #ff, ln
            If InStrRev(ln, "Orig" & Chr(21) & Chr(12) & Chr(15) & Chr(20) & Chr(8)) > 0 Then orig = True: ch = True
            If InStrRev(ln, "Comp" & Chr(24) & Chr(15) & Chr(16) & Chr(6) & Chr(2)) > 0 Then orig = False: ch = True
            If orig Then
                If ch Then
                    ch = False
                Else
                    str1 = str1 & ln & vbCrLf
                End If
            Else
                If ch Then
                    ch = False
                Else
                    str2 = str2 & ln & vbCrLf
                End If
            End If
        Loop
    Close #ff
    If txtOriginal <> Left(str1, Len(str1) - 2) Then
        Select Case MsgBox("Do you want to save changes?", 35, "Changes Detected")
            Case vbYes
                SaveChanges
                GoTo hndl
            Case vbNo
                DoOpen
                Exit Sub
            Case Else
                Exit Sub
        End Select
    End If
    If txtCompiled <> Left(str2, Len(str2) - 2) Then
        Select Case MsgBox("Do you want to save changes?", 35, "Changes Detected")
            Case vbYes
                SaveChanges
                GoTo hndl
            Case vbNo
                DoOpen
                Exit Sub
            Case Else
                Exit Sub
        End Select
    End If
End If
hndl:
If Saved Then
    DoOpen
End If
End Sub

Private Sub mnuPaste_Click()
If compF Then SendMessage txtCompiled.hWnd, WM_PASTE, 0, 0
If origF Then SendMessage txtOriginal.hWnd, WM_PASTE, 0, 0
End Sub

Private Sub mnuSave_Click()
SaveChanges
End Sub

Private Sub mnuSaveAs_Click()
DoSaveAs
End Sub

Private Sub mnuSelAll_Click()
On Error Resume Next
If compF Then SendMessage txtCompiled.hWnd, EM_SETSEL, 0, -1
If origF Then SendMessage txtOriginal.hWnd, EM_SETSEL, 0, -1
End Sub

Private Sub mnuUndo_Click()
If compF Then SendMessage txtCompiled.hWnd, WM_UNDO, 0, 0
If origF Then SendMessage txtOriginal.hWnd, WM_UNDO, 0, 0
End Sub

Private Sub Timer1_Timer()
If WindowState <> 1 Then
    If WindowState <> 2 Then
        If tmpL <> Height Then tmpL = Height
        If tmpW <> Width Then tmpW = Width
        If tmpX <> Left Then tmpX = Left
        If tmpY <> Top Then tmpY = Top
    End If
End If
End Sub

Private Sub Timer2_Timer()
On Error Resume Next
If tmpComp2 <> compF Then tmpComp2 = compF
If tmpOrig2 <> origF Then tmpOrig2 = origF
If origF Then
    With txtOriginal
        If tmpSt2 <> .SelStart Then tmpSt2 = .SelStart
        If tmpOrig1 <> tmpOrig2 Then
            GetTBPos txtOriginal
        Else
            If tmpSt1 <> tmpSt2 Then GetTBPos txtOriginal
        End If
        If tmpSt1 <> .SelStart Then tmpSt1 = .SelStart
    End With
ElseIf compF Then
    With txtCompiled
        If tmpSt2 <> .SelStart Then tmpSt2 = .SelStart
        If tmpComp1 <> tmpComp2 Then
            GetTBPos txtCompiled
        Else
            If tmpSt1 <> tmpSt2 Then GetTBPos txtCompiled
        End If
        If tmpSt1 <> .SelStart Then tmpSt1 = .SelStart
    End With
Else
    If lblStat.Caption <> "" Then lblStat.Caption = ""
End If
If tmpComp1 <> compF Then tmpComp1 = compF
If tmpOrig1 <> origF Then tmpOrig1 = origF
End Sub

Private Sub txtCompiled_GotFocus()
origF = False
compF = True
End Sub

Private Sub txtOriginal_GotFocus()
origF = True
compF = False
End Sub
