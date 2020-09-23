VERSION 5.00
Begin VB.Form frmOptions 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "FC6"
   ClientHeight    =   7530
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4815
   LinkTopic       =   "Form1"
   ScaleHeight     =   7530
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   2520
      Top             =   6840
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   5535
      Left            =   180
      TabIndex        =   1
      Top             =   960
      Width           =   4440
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   5295
         Left            =   20
         ScaleHeight     =   5295
         ScaleWidth      =   4380
         TabIndex        =   2
         Top             =   120
         Width           =   4380
         Begin VB.CheckBox chkUninstallService 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Completely uninstall Filter Service when exiting."
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   795
            TabIndex        =   15
            Top             =   4260
            Width           =   3165
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Execute at Windows Startup"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   795
            TabIndex        =   14
            Top             =   4560
            Width           =   3165
         End
         Begin VB.ComboBox cmbSkin 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   4920
            Width           =   1365
         End
         Begin VB.ListBox lstDrivesList 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1185
            Left            =   120
            Style           =   1  'Checkbox
            TabIndex        =   7
            Top             =   195
            Width           =   2535
         End
         Begin VB.CheckBox chkSelectAll 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Select All Drives"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            TabIndex        =   6
            Top             =   1435
            Width           =   1485
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Please select the NTFS drive(s) and  click Connect to Service (Drive can be Fixed or Removable)"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1095
            Left            =   2760
            TabIndex        =   11
            Top             =   180
            Width           =   1695
         End
         Begin VB.Label Label4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "(Use Recover/Restore Filter Service only if you can not get the work done with the filter service.)"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1080
            TabIndex        =   20
            Top             =   3600
            Width           =   2970
            WordWrap        =   -1  'True
         End
         Begin VB.Image Image6 
            Height          =   255
            Left            =   765
            Picture         =   "FRMOPT~1.frx":0000
            Top             =   3600
            Width           =   240
         End
         Begin VB.Image Image9 
            Height          =   600
            Left            =   240
            Picture         =   "FRMOPT~1.frx":0372
            Top             =   2760
            Width           =   525
         End
         Begin VB.Image imgReloadFilter 
            Height          =   285
            Left            =   1065
            MouseIcon       =   "FRMOPT~1.frx":1494
            MousePointer    =   99  'Custom
            Picture         =   "FRMOPT~1.frx":15E6
            Top             =   3180
            Width           =   300
         End
         Begin VB.Image imgTestService 
            Height          =   390
            Left            =   1065
            MouseIcon       =   "FRMOPT~1.frx":1A9C
            MousePointer    =   99  'Custom
            Picture         =   "FRMOPT~1.frx":1BEE
            Top             =   2715
            Width           =   330
         End
         Begin VB.Label lblTestService 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Test the Service Task"
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
            Left            =   1170
            LinkTimeout     =   100
            MouseIcon       =   "FRMOPT~1.frx":2318
            MousePointer    =   99  'Custom
            TabIndex        =   19
            Top             =   2790
            Width           =   2070
         End
         Begin VB.Label lblRecoverFilterDrv 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Recover/Restore Filter Service"
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
            Left            =   1110
            LinkTimeout     =   100
            MouseIcon       =   "FRMOPT~1.frx":246A
            MousePointer    =   99  'Custom
            TabIndex        =   18
            Top             =   3225
            Width           =   2925
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ttransparency"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   795
            TabIndex        =   17
            Top             =   4920
            Width           =   885
         End
         Begin VB.Line Line7 
            X1              =   960
            X2              =   4080
            Y1              =   4080
            Y2              =   4080
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "General"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   120
            TabIndex        =   16
            Top             =   3960
            Width           =   600
         End
         Begin VB.Image Image2 
            Height          =   495
            Left            =   120
            Picture         =   "FRMOPT~1.frx":25BC
            Top             =   4320
            Width           =   525
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Comman"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   120
            TabIndex        =   12
            Top             =   2355
            Width           =   690
         End
         Begin VB.Image Image7 
            Height          =   675
            Left            =   840
            Picture         =   "FRMOPT~1.frx":33EA
            Stretch         =   -1  'True
            Top             =   0
            Width           =   3240
         End
         Begin VB.Line Line2 
            X1              =   120
            X2              =   4080
            Y1              =   1755
            Y2              =   1755
         End
         Begin VB.Image imgReset 
            Height          =   330
            Left            =   120
            MouseIcon       =   "FRMOPT~1.frx":A614
            MousePointer    =   99  'Custom
            Picture         =   "FRMOPT~1.frx":A766
            Top             =   1920
            Width           =   375
         End
         Begin VB.Label lblReset 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Reset Service"
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
            Left            =   285
            LinkTimeout     =   100
            MouseIcon       =   "FRMOPT~1.frx":AE30
            MousePointer    =   99  'Custom
            TabIndex        =   10
            Top             =   1950
            Width           =   1410
         End
         Begin VB.Line Line1 
            X1              =   960
            X2              =   4080
            Y1              =   2475
            Y2              =   2475
         End
         Begin VB.Label lblADFilter 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Connect to Service..."
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
            Left            =   2100
            LinkTimeout     =   100
            MouseIcon       =   "FRMOPT~1.frx":AF82
            MousePointer    =   99  'Custom
            TabIndex        =   9
            Top             =   1950
            Width           =   1965
         End
         Begin VB.Image imgADFilter 
            Height          =   405
            Left            =   1935
            MouseIcon       =   "FRMOPT~1.frx":B0D4
            MousePointer    =   99  'Custom
            Picture         =   "FRMOPT~1.frx":B226
            Top             =   1875
            Width           =   375
         End
         Begin VB.Image imgRefresh 
            Height          =   330
            Left            =   2760
            MouseIcon       =   "FRMOPT~1.frx":BA6C
            MousePointer    =   99  'Custom
            Picture         =   "FRMOPT~1.frx":BBBE
            Top             =   1080
            Width           =   300
         End
         Begin VB.Label lblRefresh 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Refresh..."
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
            Left            =   2865
            LinkTimeout     =   100
            MouseIcon       =   "FRMOPT~1.frx":C128
            MousePointer    =   99  'Custom
            TabIndex        =   8
            Top             =   1125
            Width           =   1020
         End
      End
   End
   Begin VB.Line Line8 
      X1              =   120
      X2              =   4545
      Y1              =   6720
      Y2              =   6720
   End
   Begin VB.Label lblEmail 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "pgbsoft@gmail.com"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   165
      Left            =   600
      MouseIcon       =   "FRMOPT~1.frx":C27A
      MousePointer    =   99  'Custom
      TabIndex        =   23
      Top             =   6960
      Width           =   1260
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail:"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   165
      Left            =   120
      MouseIcon       =   "FRMOPT~1.frx":C3CC
      MousePointer    =   99  'Custom
      TabIndex        =   22
      Top             =   6960
      Width           =   420
   End
   Begin VB.Image imgClose 
      Height          =   330
      Left            =   3720
      MouseIcon       =   "FRMOPT~1.frx":C51E
      MousePointer    =   99  'Custom
      Top             =   6885
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
      Left            =   3855
      LinkTimeout     =   100
      MouseIcon       =   "FRMOPT~1.frx":C670
      MousePointer    =   99  'Custom
      TabIndex        =   21
      Top             =   6960
      Width           =   675
   End
   Begin VB.Label lblAbout 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "About..."
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Left            =   4080
      MouseIcon       =   "FRMOPT~1.frx":C7C2
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   600
      Width           =   480
   End
   Begin VB.Label lblHelp 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Features..."
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Left            =   3360
      MouseIcon       =   "FRMOPT~1.frx":C914
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   600
      Width           =   645
   End
   Begin VB.Image imgCloseBottom 
      Height          =   315
      Left            =   4320
      MouseIcon       =   "FRMOPT~1.frx":CA66
      MousePointer    =   99  'Custom
      ToolTipText     =   "Close"
      Top             =   42
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image imgCloseTop 
      Height          =   240
      Left            =   4350
      Top             =   120
      Width           =   225
   End
   Begin VB.Line Line6 
      BorderColor     =   &H006167E7&
      X1              =   10
      X2              =   4770
      Y1              =   400
      Y2              =   400
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Options..."
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
      Left            =   105
      TabIndex        =   3
      Top             =   105
      Width           =   780
   End
   Begin VB.Line Line4 
      X1              =   4785
      X2              =   4800
      Y1              =   360
      Y2              =   7560
   End
   Begin VB.Line Line5 
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   7560
   End
   Begin VB.Image Image5 
      Height          =   90
      Left            =   -840
      Stretch         =   -1  'True
      Top             =   7440
      Width           =   5865
   End
   Begin VB.Image imgTop 
      Height          =   405
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4800
   End
   Begin VB.Line Line3 
      X1              =   180
      X2              =   4605
      Y1              =   6000
      Y2              =   6000
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "General and Security Settings..."
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
      Left            =   240
      TabIndex        =   0
      Top             =   580
      Width           =   2655
   End
   Begin VB.Image Image3 
      Height          =   525
      Left            =   0
      Stretch         =   -1  'True
      Top             =   420
      Width           =   4780
   End
   Begin VB.Image Image1 
      Height          =   6735
      Left            =   0
      Stretch         =   -1  'True
      Top             =   840
      Width           =   75
   End
   Begin VB.Image Image4 
      Height          =   6735
      Left            =   4725
      Stretch         =   -1  'True
      Top             =   840
      Width           =   75
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================
'=  Developed by P. G. B. Prasanna       =
'=  A Software Developer from Sri Lanka  =
'=  E-Mail: pgbsoft@gmail.com            =
'=========================================

Private Sub chkSelectAll_Click()
If chkSelectAll.Value = 1 Then
    For i = 0 To lstDrivesList.ListCount - 1
        lstDrivesList.Selected(i) = True
    Next
Else
    For i = 0 To lstDrivesList.ListCount - 1
        lstDrivesList.Selected(i) = False
    Next
End If
lstDrivesList.ListIndex = -1
End Sub

Private Sub chkSelectAll_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Mouse_Move_Effect
End Sub

Private Sub chkUninstallService_Click()
Call Write_Reg(R_LocApp, chkUninstallService.Value)
Select Case chkUninstallService.Value
    Case 0: intService_Uninstall = 0
    Case 1: intService_Uninstall = 1
End Select
End Sub

Private Sub chkUninstallService_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Mouse_Move_Effect
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
lblADFilter_MouseDown 0, 0, 0, 0
lblADFilter_MouseUp 0, 0, 0, 0
End Sub
Private Sub Form_Activate()
lstDrivesList.ListIndex = -1
If intIsFirstFlag = 1 Then frmFeatures.Show 1
End Sub

Private Sub Form_Load()
On Error Resume Next
Me.Label3.Caption = AppTitles(0)
Get_Requried_Drives
Set_Round Me
Apply_Theme 1, Me

Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2

chkUninstallService.Value = intService_Uninstall
If intAdmin = 0 Then
    Enable_Disable_Control False
    imgCloseTop.Enabled = True
    imgClose.Enabled = True
    lblClose.Enabled = True
End If

On Error Resume Next
lStartStyle = GetWindowLong(hWnd, GWL_EXSTYLE)
SetWindowLong hWnd, GWL_EXSTYLE, lStartStyle Or WS_EX_LAYERED
bTrans_P_Level = 0
Timer1.Enabled = True
intActionMode = 1
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Mouse_Move_Control Me
If Button = 1 Then
    Getmove Me
End If
Mouse_Move_Effect
End Sub

Private Sub Form_Unload(Cancel As Integer)
chkSelectAll_Click
chkSelectAll.Value = 0
Attach_Detach_Filter
Control_MiniFilter_Driver 4
If intService_Uninstall = 1 Then: Control_MiniFilter_Driver 6: Config_Drv 2
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Mouse_Move_Control Me
If Button = 1 Then
    Getmove Me
End If
Mouse_Move_Effect
End Sub

Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Mouse_Move_Control Me
If Button = 1 Then
    Getmove Me
End If
Mouse_Move_Effect
End Sub

Private Sub Image4_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Mouse_Move_Control Me
If Button = 1 Then
    Getmove Me
End If
Mouse_Move_Effect
End Sub

Private Sub Image5_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Mouse_Move_Control Me
If Button = 1 Then
    Getmove Me
End If
Mouse_Move_Effect
End Sub
Private Sub Image6_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Mouse_Move_Control Me
If Button = 1 Then
    Getmove Me
End If
End Sub
Private Sub Image7_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Mouse_Move_Control Me
If Button = 1 Then
    Getmove Me
End If
Mouse_Move_Effect
End Sub
Private Sub Image9_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Mouse_Move_Control Me
If Button = 1 Then
    Getmove Me
End If
Mouse_Move_Effect
End Sub

Private Sub imgADFilter_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
lblADFilter_MouseDown 0, 0, 0, 0
End Sub

Private Sub imgADFilter_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
lblADFilter_MouseMove 0, 0, 0, 0
End Sub

Private Sub imgADFilter_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
lblADFilter_MouseUp 0, 0, 0, 0
End Sub

Private Sub imgClose_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
lblClose_MouseDown 0, 0, 0, 0
End Sub

Private Sub imgClose_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
lblClose_MouseMove 0, 0, 0, 0
End Sub

Private Sub imgClose_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
lblClose_MouseUp 0, 0, 0, 0
End Sub

Private Sub imgCloseBottom_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Close_Button_Action Me, 1
End Sub

Private Sub imgCloseBottom_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Close_Button_Action Me, 2
End Sub

Private Sub imgCloseTop_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
imgCloseTop.Visible = False
imgCloseBottom.Visible = True
End Sub
Private Sub imgRefresh_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
lblRefresh_MouseDown 0, 0, 0, 0
End Sub

Private Sub imgRefresh_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
lblRefresh_MouseMove 0, 0, 0, 0
End Sub

Private Sub imgRefresh_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
lblRefresh_MouseUp 0, 0, 0, 0
End Sub

Private Sub imgReloadFilter_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
lblRecoverFilterDrv_MouseDown 0, 0, 0, 0
End Sub

Private Sub imgReloadFilter_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
lblRecoverFilterDrv_MouseMove 0, 0, 0, 0
End Sub

Private Sub imgReloadFilter_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
lblRecoverFilterDrv_MouseUp 0, 0, 0, 0
End Sub

Private Sub imgReset_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
lblReset_MouseDown 0, 0, 0, 0
End Sub

Private Sub imgReset_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
lblReset_MouseMove 0, 0, 0, 0
End Sub

Private Sub imgReset_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
lblReset_MouseUp 0, 0, 0, 0
End Sub

Private Sub imgTestService_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
lblTestService_MouseDown 0, 0, 0, 0
End Sub

Private Sub imgTestService_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
lblTestService_MouseMove 0, 0, 0, 0
End Sub

Private Sub imgTestService_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
lblTestService_MouseUp 0, 0, 0, 0
End Sub

Private Sub imgTop_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Mouse_Move_Control Me
If Button = 1 Then
    Getmove Me
End If
Mouse_Move_Effect
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Mouse_Move_Control Me
If Button = 1 Then
    Getmove Me
End If
Mouse_Move_Effect
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Mouse_Move_Control Me
If Button = 1 Then
    Getmove Me
End If
Mouse_Move_Effect
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Mouse_Move_Control Me
If Button = 1 Then
    Getmove Me
End If
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Mouse_Move_Control Me
If Button = 1 Then
    Getmove Me
End If
Mouse_Move_Effect
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Mouse_Move_Effect
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Mouse_Move_Effect
End Sub

Private Sub lblAbout_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Mouse_Up_Down_Handling lblAbout
End Sub

Private Sub lblAbout_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Mouse_Move_Effect
End Sub

Private Sub lblAbout_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Mouse_Up_Down_Handling lblAbout, , 1
frmAbout.Show 1
End Sub

Private Sub lblADFilter_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Mouse_Up_Down_Handling lblADFilter, imgADFilter
End Sub

Private Sub lblADFilter_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Mouse_Move_Effect
lblADFilter.ForeColor = &H80FF&
End Sub

Private Sub lblADFilter_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Mouse_Up_Down_Handling lblADFilter, imgADFilter, 1

If Check_For_Selected_Drives = False Then Exit Sub

Enable_Disable_Control False
Attach_Detach_Filter
CenterMessage Me, "Filter service attached to the selected drive(s)...", vbInformation, App.Title
Enable_Disable_Control True
SHChangeNotify SHCNE_ASSOCCHANGED, SHCNF_FLUSH, 0, 0
End Sub

Private Sub lblClose_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Mouse_Up_Down_Handling lblClose, imgClose
End Sub

Private Sub lblClose_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Mouse_Move_Effect
lblClose.ForeColor = &H80FF&
End Sub

Private Sub lblClose_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Mouse_Up_Down_Handling lblClose, imgClose, 1
Unload Me
End Sub

Private Sub lblEmail_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Mouse_Up_Down_Handling lblEmail
End Sub

Private Sub lblEmail_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Mouse_Move_Effect
lblEmail.ForeColor = &HC0&
End Sub

Private Sub lblEmail_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Mouse_Up_Down_Handling lblEmail, , 1
Mail_To frmMain
End Sub

Private Sub lblHelp_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Mouse_Up_Down_Handling lblHelp
End Sub

Private Sub lblHelp_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Mouse_Move_Effect
End Sub

Private Sub lblHelp_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Mouse_Up_Down_Handling lblHelp, , 1
frmFeatures.Show 1
End Sub
Private Sub lstDriveList_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Mouse_Move_Control Me
If Button = 1 Then
    Getmove Me
End If
End Sub

Private Sub lblRecoverFilterDrv_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Mouse_Up_Down_Handling lblRecoverFilterDrv, imgReloadFilter
End Sub

Private Sub lblRecoverFilterDrv_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Mouse_Move_Effect
lblRecoverFilterDrv.ForeColor = &H80FF&
End Sub

Private Sub lblRecoverFilterDrv_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Mouse_Up_Down_Handling lblRecoverFilterDrv, imgReloadFilter, 1

Enable_Disable_Control False

chkSelectAll_Click
chkSelectAll.Value = 0
Attach_Detach_Filter
Control_MiniFilter_Driver 4
Control_MiniFilter_Driver 5
Control_MiniFilter_Driver 6
Config_Drv 2
Sleep 100
Config_Drv 1
Control_MiniFilter_Driver 3
CenterMessage Me, "Filter Service was recovered..." & vbCrLf & vbCrLf & "Tip" & vbCrLf & _
       "----" & vbCrLf & "If not succeeded, try again.", vbInformation, App.Title
Enable_Disable_Control True
SHChangeNotify SHCNE_ASSOCCHANGED, SHCNF_FLUSH, 0, 0
End Sub

Private Sub lblRefresh_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Mouse_Up_Down_Handling lblRefresh, imgRefresh
End Sub

Private Sub lblRefresh_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Mouse_Move_Effect
lblRefresh.ForeColor = &H80FF&
End Sub

Private Sub lblRefresh_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
lblRefresh.Left = 2865: lblRefresh.Top = 1920
imgRefresh.Left = 2760: imgRefresh.Top = 1880
Enable_Disable_Control False
chkSelectAll_Click
chkSelectAll.Value = 0
Attach_Detach_Filter
Get_Requried_Drives
Enable_Disable_Control True
End Sub

Private Sub lblReset_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Mouse_Up_Down_Handling lblReset, imgReset
End Sub

Private Sub lblReset_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Mouse_Move_Effect
lblReset.ForeColor = &H80FF&
End Sub

Private Sub lblReset_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Mouse_Up_Down_Handling lblReset, imgReset, 1
Enable_Disable_Control False
chkSelectAll_Click
chkSelectAll.Value = 0
Attach_Detach_Filter
'Control_MiniFilter_Driver 4
'Control_MiniFilter_Driver 3
CenterMessage Me, "Filter service was reset...", vbInformation, App.Title
Enable_Disable_Control True
SHChangeNotify SHCNE_ASSOCCHANGED, SHCNF_FLUSH, 0, 0
End Sub

Private Sub lblTestService_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Mouse_Up_Down_Handling lblTestService, imgTestService
End Sub

Private Sub lblTestService_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Mouse_Move_Effect
lblTestService.ForeColor = &H80FF&
End Sub

Private Sub lblTestService_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Mouse_Up_Down_Handling lblTestService, imgTestService, 1

Enable_Disable_Control False
Dim sysVI_On_Sys As String
sysVI_On_Sys = Environ$("SystemDrive") & "\System Volume Information"
If PathIsDirectory(sysVI_On_Sys) = 0 Then: CenterMessage Me, "Unexpected Error", vbCritical, AppTitles(3): Enable_Disable_Control True: Exit Sub
    
On Error Resume Next
Dim sd As Integer
Dim Found_Sys_Drive As Integer
Found_Sys_Drive = 0
For sd = 0 To lstDrivesList.ListCount - 1
    If UCase(Mid(lstDrivesList.List(sd), Len(lstDrivesList.List(sd)) - 2, 2)) = UCase(Environ$("SystemDrive")) Then
        lstDrivesList.Selected(sd) = True: Found_Sys_Drive = 1
        Exit For
    End If
Next
CenterMessage Me, "Program will now open " & sysVI_On_Sys & " directory" & vbCrLf & _
       "generally protected with NTFS permission, for testing purpose." & vbCrLf & _
       "Tip:" & vbCrLf & "-----" & vbCrLf & _
       "If succeeded, filter service is working properly.", vbInformation, App.Title
If Found_Sys_Drive = 1 Then Attach_Detach_Filter: Shell "Explorer " & sysVI_On_Sys, vbNormalFocus: Enable_Disable_Control True
End Sub

Private Sub lstDrivesList_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Mouse_Move_Effect
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Mouse_Move_Control Me
If Button = 1 Then
    Getmove Me
End If
Mouse_Move_Effect
End Sub

Private Sub Timer1_Timer()
Call Control_Transparent_Effect(Me)
End Sub
'Getting the NTFS Drives (Fixed and Removable)
Private Sub Get_Requried_Drives()
On Error GoTo ErrDrive
Dim Drv, DrvCol, Vol_N
Dim intErrNotReady As Integer
Set DrvCol = fs_obj.Drives
lstDrivesList.Clear
    For Each Drv In DrvCol
        If UCase(Drv.DriveLetter) <> "A" And UCase(Drv.DriveLetter) <> "B" Then
            Vol_N = ""
            intErrNotReady = 0
            On Error GoTo ErrDrive
                Select Case Drv.DriveType
                    Case 1: Vol_N = Drv.VolumeName
                            If Vol_N = "" Then Vol_N = "Removable Disk"
                                If intErrNotReady = 0 Then
                                    If Determine_NTFS(Drv.DriveLetter & ":\") = True Then
                                        lstDrivesList.AddItem Vol_N & " (" & Drv.DriveLetter & ":)"
                                    End If
                                End If
                    Case 2: Vol_N = Drv.VolumeName
                            If Vol_N = "" Then Vol_N = "Local Disk"
                                If intErrNotReady = 0 Then
                                    If Determine_NTFS(Drv.DriveLetter & ":\") = True Then
                                        lstDrivesList.AddItem Vol_N & " (" & Drv.DriveLetter & ":)"
                                    End If
                                End If
                End Select
        End If
    Next
    
    Exit Sub
ErrDrive:
intErrNotReady = 1
Resume Next
End Sub
'Determine the File System of the drive is NTFS
Private Function Determine_NTFS(ByVal DrvR As String) As Boolean
Dim sVNBuffer As String
Dim sFSNBuffer As String
Dim lSerialNumber As Long
Dim lMaxCompLen As Long
Dim lFileSysFlags As Long

sVNBuffer = String(255, 0)
sFSNBuffer = String(255, 0)
Determine_NTFS = False

Call GetVolumeInformation(DrvR, sVNBuffer, 255&, lSerialNumber, lMaxCompLen, lFileSysFlags, sFSNBuffer, 255&)
If UCase(Left$(sFSNBuffer, 4)) = "NTFS" Then Determine_NTFS = True
End Function

Private Function Check_For_Selected_Drives() As Boolean
Dim i As Integer
Check_For_Selected_Drives = False
For i = 0 To lstDrivesList.ListCount - 1
     If lstDrivesList.Selected(i) = True Then
        Check_For_Selected_Drives = True
        Exit Function
     End If
Next
If Check_For_Selected_Drives = False Then CenterMessage Me, "Please select the drive(s)... !", vbExclamation, App.Title: lstDrivesList.SetFocus
End Function
Private Sub Attach_Detach_Filter()
Dim i As Integer
    For i = 0 To lstDrivesList.ListCount - 1
        strDrvL = Mid(lstDrivesList.List(i), Len(lstDrivesList.List(i)) - 2, 2)
        If lstDrivesList.Selected(i) = True Then: Control_MiniFilter_Driver 1, strDrvL
        If lstDrivesList.Selected(i) = False Then: Control_MiniFilter_Driver 2, strDrvL
        Sleep 10
    Next
End Sub
Private Sub Enable_Disable_Control(ByVal val As Boolean)
lstDrivesList.Enabled = val: lstDrivesList.Refresh
chkSelectAll.Enabled = val: chkSelectAll.Refresh
lblReset.Enabled = val: lblReset.Refresh
imgReset.Enabled = val: imgReset.Refresh
lblADFilter.Enabled = val: lblADFilter.Refresh
imgADFilter.Enabled = val: imgADFilter.Refresh
lblRecoverFilterDrv.Enabled = val: lblRecoverFilterDrv.Refresh
imgReloadFilter.Enabled = val: imgReloadFilter.Refresh
cmdClose.Enabled = val: cmdClose.Refresh
imgCloseTop.Enabled = val: imgCloseTop.Refresh
imgClose.Enabled = val: imgClose.Refresh
lblClose.Enabled = val: lblClose.Refresh
lblTestService.Enabled = val: lblTestService.Refresh
imgTestService.Enabled = val: imgTestService.Refresh
lblRefresh.Enabled = val: lblRefresh.Refresh
imgRefresh.Enabled = val: imgRefresh.Refresh
chkUninstallService.Enabled = val: chkUninstallService.Refresh
End Sub
Public Sub Mouse_Move_Effect()
lblRefresh.ForeColor = vbBlack
lblReset.ForeColor = vbBlack
lblADFilter.ForeColor = vbBlack
lblTestService.ForeColor = vbBlack
lblRecoverFilterDrv.ForeColor = vbBlack
lblClose.ForeColor = vbBlack
lblEmail.ForeColor = vbBlack
End Sub
Public Sub Mouse_Up_Down_Handling(lblName As Label, Optional imgName As Image, Optional Eventopt As Integer = 0)
On Error Resume Next
Select Case Eventopt
    Case 0: lblName.Left = lblName.Left + 20: lblName.Top = lblName.Top + 20
            imgName.Left = imgName.Left + 20: imgName.Top = imgName.Top + 20
    Case 1: lblName.Left = lblName.Left - 20: lblName.Top = lblName.Top - 20
            imgName.Left = imgName.Left - 20: imgName.Top = imgName.Top - 20
End Select
End Sub


