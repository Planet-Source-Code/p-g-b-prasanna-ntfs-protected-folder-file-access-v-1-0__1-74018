VERSION 5.00
Begin VB.Form frmFeatures 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Help"
   ClientHeight    =   4530
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4290
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   4290
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdEnterEsc 
      Cancel          =   -1  'True
      Caption         =   "EnterEsc"
      Default         =   -1  'True
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   10000
      Width           =   375
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2760
      Top             =   4080
   End
   Begin VB.TextBox txtInfo 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   150
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   850
      Width           =   3975
   End
   Begin VB.Image imgClose 
      Height          =   330
      Left            =   3240
      MouseIcon       =   "frmFeatures.frx":0000
      MousePointer    =   99  'Custom
      Top             =   4075
      Width           =   300
   End
   Begin VB.Label lblClose 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   3375
      LinkTimeout     =   100
      MouseIcon       =   "frmFeatures.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   4135
      Width           =   675
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   4690
   End
   Begin VB.Line Line3 
      X1              =   4260
      X2              =   4260
      Y1              =   0
      Y2              =   4570
   End
   Begin VB.Image Image5 
      Height          =   80
      Left            =   0
      Stretch         =   -1  'True
      Top             =   4450
      Width           =   4335
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00000000&
      X1              =   0
      X2              =   4270
      Y1              =   400
      Y2              =   400
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Features..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      OLEDropMode     =   1  'Manual
      TabIndex        =   3
      Top             =   105
      Width           =   885
   End
   Begin VB.Image imgCloseBottom 
      Height          =   315
      Left            =   3810
      MouseIcon       =   "frmFeatures.frx":02A4
      MousePointer    =   99  'Custom
      ToolTipText     =   "Close"
      Top             =   60
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image imgCloseTop 
      Height          =   210
      Left            =   3840
      Top             =   120
      Width           =   240
   End
   Begin VB.Image imgTop 
      Height          =   400
      Left            =   10
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4255
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      X1              =   150
      X2              =   4130
      Y1              =   3970
      Y2              =   3970
   End
   Begin VB.Image Image1 
      Height          =   405
      Left            =   105
      Top             =   4030
      Width           =   435
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "What is NTFS Protected Folder/File Accessr?"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Left            =   160
      TabIndex        =   1
      Top             =   505
      Width           =   2775
   End
   Begin VB.Image Image2 
      Height          =   415
      Left            =   0
      Stretch         =   -1  'True
      Top             =   360
      Width           =   4255
   End
   Begin VB.Image Image3 
      Height          =   5635
      Left            =   0
      Top             =   720
      Width           =   75
   End
   Begin VB.Image Image4 
      Height          =   5635
      Left            =   4200
      Top             =   0
      Width           =   75
   End
End
Attribute VB_Name = "frmFeatures"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================
'=  Developed by P. G. B. Prasanna       =
'=  A Software Developer from Sri Lanka  =
'=  E-Mail: pgbsoft@gmail.com            =
'=========================================

Private Sub cmdOK_Click()
Unload Me
End Sub
Private Sub cmdOk_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Move_Control Me
End Sub

Private Sub cmdEnterEsc_Click()
Unload Me
End Sub

Private Sub Form_Activate()
If intIsFirstFlag = 1 Then frmAbout.Show 1
End Sub
Private Sub Form_Load()
On Error Resume Next
Apply_Theme 2, Me, intSkin
Set_Round Me
Dim strFeatures As String
strFeatures = "--------------------------------------------------------------------" & vbCrLf
strFeatures = strFeatures & "Features of NTFS Protected Folder/File Access" & vbCrLf
strFeatures = strFeatures & "--------------------------------------------------------------------" & vbCrLf & vbCrLf
strFeatures = strFeatures & "• NTFS Protected Folder/File Access is a useful and  " & vbCrLf
strFeatures = strFeatures & "  powerful tool, which provides you the access to any" & vbCrLf
strFeatures = strFeatures & "  NTFS Protected Folder/File on your drives" & vbCrLf
strFeatures = strFeatures & "  without resetting the permission of them." & vbCrLf & vbCrLf
strFeatures = strFeatures & "• The tool can be used with NTFS-Formatted" & vbCrLf
strFeatures = strFeatures & "  Removable drives as well." & vbCrLf & vbCrLf
strFeatures = strFeatures & "• It provides a great interface to interact with its operations" & vbCrLf
strFeatures = strFeatures & "  together with easy, efficient and effective usage." & vbCrLf & vbCrLf
strFeatures = strFeatures & "• No matter how many files/folders are protected with" & vbCrLf
strFeatures = strFeatures & "  NTFS on a drive(s), you will have instant access to" & vbCrLf
strFeatures = strFeatures & "  them with a single click." & vbCrLf & vbCrLf
strFeatures = strFeatures & "You can use this tool," & vbCrLf
strFeatures = strFeatures & "-------------------------------" & vbCrLf
strFeatures = strFeatures & "• to instantly access any drive(s), which you get the" & vbCrLf
strFeatures = strFeatures & "  message 'Access is denied' because of NTFS" & vbCrLf
strFeatures = strFeatures & "  permission." & vbCrLf & vbCrLf
strFeatures = strFeatures & "• to access folders/files when you mess up with NTFS" & vbCrLf
strFeatures = strFeatures & "  permission of a drive." & vbCrLf & vbCrLf
strFeatures = strFeatures & "• when you cannot take the ownership of folders/files" & vbCrLf
strFeatures = strFeatures & "  on a drive to use them, because the drive is read only" & vbCrLf
strFeatures = strFeatures & "  (Eg: read-only hard drive.) or mounted as read-only." & vbCrLf & vbCrLf
strFeatures = strFeatures & "• when considering the time which takes to gain the" & vbCrLf
strFeatures = strFeatures & "  ownership of millions of files of a drive in order to" & vbCrLf
strFeatures = strFeatures & "  use them." & vbCrLf & vbCrLf
strFeatures = strFeatures & "• to temporarily access folders/files which you cannot" & vbCrLf
strFeatures = strFeatures & "  access or modify because of the guard provided by " & vbCrLf
strFeatures = strFeatures & "  some of the security products on the market." & vbCrLf & vbCrLf
strFeatures = strFeatures & "• to access the files which seem to have been" & vbCrLf
strFeatures = strFeatures & "  corrupted because of NTFS permission." & vbCrLf & vbCrLf
strFeatures = strFeatures & "• to access the NTFS protected Drives/Folders/Files," & vbCrLf
strFeatures = strFeatures & "  when you can not reset to access them because" & vbCrLf
strFeatures = strFeatures & "  Security Tab is missing on Properties and you can't get " & vbCrLf
strFeatures = strFeatures & "  it reset due to Folder Options/Use simple file sharing" & vbCrLf
strFeatures = strFeatures & "  is restricted or missing." & vbCrLf & vbCrLf
strFeatures = strFeatures & "How to use the program:" & vbCrLf
strFeatures = strFeatures & "-----------------------------------" & vbCrLf
strFeatures = strFeatures & "• Click on 'Test Service Task' first. (Recommended)" & vbCrLf
strFeatures = strFeatures & "  It will open 'System Volume Information' folder" & vbCrLf
strFeatures = strFeatures & "  of your system drive, which you don't generally have " & vbCrLf
strFeatures = strFeatures & "  access to. If not succeeded, click on 'Recover/Restore" & vbCrLf
strFeatures = strFeatures & "  Filter Service'." & vbCrLf & vbCrLf
strFeatures = strFeatures & "• Select drive(s)on which you have NTFS-Protected data," & vbCrLf
strFeatures = strFeatures & "  which you cannot access to." & vbCrLf & vbCrLf
strFeatures = strFeatures & "• Then, click on 'Connect to Service'." & vbCrLf & vbCrLf
strFeatures = strFeatures & "• Now you will wonderfully have access to NTFS-Protected" & vbCrLf
strFeatures = strFeatures & "  data without resetting the permission of them." & vbCrLf & vbCrLf
strFeatures = strFeatures & "Supported operating systems:" & vbCrLf
strFeatures = strFeatures & "------------------------------------------" & vbCrLf
strFeatures = strFeatures & "• Windows 2000 SP4 Rollup 1" & vbCrLf
strFeatures = strFeatures & "• Windows XP SP2, SP3" & vbCrLf
strFeatures = strFeatures & "• Windows Server 2003 32-bit" & vbCrLf
strFeatures = strFeatures & "• Windows Vista RTM, SP1 32-bit" & vbCrLf
strFeatures = strFeatures & "• Windows Server 2008 32-bit" & vbCrLf
strFeatures = strFeatures & "• Windows 7 32-bit" & vbCrLf & vbCrLf
strFeatures = strFeatures & "Note:" & vbCrLf
strFeatures = strFeatures & "---------" & vbCrLf
strFeatures = strFeatures & "• NTFS Protected Folder/File Access can only be run " & vbCrLf
strFeatures = strFeatures & "  with required administrative privileges." & vbCrLf & vbCrLf
strFeatures = strFeatures & "• Supported only for 32-bit versions of Windows." & vbCrLf & vbCrLf
strFeatures = strFeatures & "   ----------------------------------------------------------------------------" & vbCrLf & vbCrLf
strFeatures = strFeatures & "   Developed by P. G. B. Prasanna" & vbCrLf
strFeatures = strFeatures & "   E-mail: pgbsoft@gmail.com"
txtInfo = strFeatures

Me.Left = frmMain.Left + (frmMain.Width - Me.Width) / 2 + 600
Me.Top = frmMain.Top + (frmMain.Height - Me.Height) / 2

lStartStyle = GetWindowLong(hwnd, GWL_EXSTYLE)
SetWindowLong hwnd, GWL_EXSTYLE, lStartStyle Or WS_EX_LAYERED
bTrans_P_Level = 0
Timer1.Enabled = True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Move_Control Me
If Button = 1 Then
    Getmove Me
End If
lblClose.ForeColor = vbBlack
End Sub
Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Move_Control Me
If Button = 1 Then
    Getmove Me
End If
lblClose.ForeColor = vbBlack
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Move_Control Me
If Button = 1 Then
    Getmove Me
End If
lblClose.ForeColor = vbBlack
End Sub

Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Move_Control Me
If Button = 1 Then
    Getmove Me
End If
lblClose.ForeColor = vbBlack
End Sub

Private Sub Image4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Move_Control Me
If Button = 1 Then
    Getmove Me
End If
lblClose.ForeColor = vbBlack
End Sub

Private Sub Image5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Move_Control Me
If Button = 1 Then
    Getmove Me
End If
lblClose.ForeColor = vbBlack
End Sub
Private Sub imgClose_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblClose_MouseDown 0, 0, 0, 0
End Sub

Private Sub imgClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblClose_MouseMove 0, 0, 0, 0
End Sub

Private Sub imgClose_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblClose_MouseUp 0, 0, 0, 0
End Sub

Private Sub imgCloseBottom_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Close_Button_Action Me, 1
End Sub

Private Sub imgCloseBottom_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Close_Button_Action Me, 2
End Sub

Private Sub imgCloseTop_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgCloseTop.Visible = False
imgCloseBottom.Visible = True
End Sub

Private Sub imgTop_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Move_Control Me
If Button = 1 Then
    Getmove Me
End If
lblClose.ForeColor = vbBlack
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Move_Control Me
If Button = 1 Then
    Getmove Me
End If
lblClose.ForeColor = vbBlack
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Move_Control Me
If Button = 1 Then
    Getmove Me
End If
lblClose.ForeColor = vbBlack
End Sub

Private Sub lblClose_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmMain.Mouse_Up_Down_Handling lblClose, imgClose
End Sub

Private Sub lblClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblClose.ForeColor = &H80FF&
End Sub

Private Sub lblClose_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmMain.Mouse_Up_Down_Handling lblClose, imgClose, 1
Unload Me
End Sub

Private Sub Timer1_Timer()
Control_Transparent_Effect Me
End Sub

Private Sub txtInfo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Move_Control Me
lblClose.ForeColor = vbBlack
End Sub
