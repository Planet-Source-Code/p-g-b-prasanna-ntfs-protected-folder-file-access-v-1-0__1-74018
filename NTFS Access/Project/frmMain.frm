VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "FC6"
   ClientHeight    =   7800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4815
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7800
   ScaleWidth      =   4815
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      MouseIcon       =   "frmMain.frx":0CCA
      MousePointer    =   99  'Custom
      TabIndex        =   27
      Top             =   10000
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   26
      Top             =   10000
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   3000
      Top             =   7320
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   6130
      Left            =   180
      TabIndex        =   14
      Top             =   960
      Width           =   4440
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   5950
         Left            =   20
         ScaleHeight     =   5955
         ScaleWidth      =   4380
         TabIndex        =   15
         Top             =   120
         Width           =   4380
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
            Left            =   2760
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   5110
            Width           =   1300
         End
         Begin VB.CheckBox chkUninstallService 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Completely uninstall Filter Service when exiting"
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
            Left            =   1065
            TabIndex        =   7
            Top             =   4210
            Width           =   3165
         End
         Begin VB.CheckBox chkStartup 
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
            Left            =   1065
            TabIndex        =   8
            Top             =   4510
            Width           =   2070
         End
         Begin VB.ComboBox cmbTransparency 
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
            Left            =   1065
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   5110
            Width           =   1300
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
            TabIndex        =   0
            Top             =   115
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
            TabIndex        =   2
            Top             =   1355
            Width           =   1485
         End
         Begin VB.Image Image8 
            Height          =   555
            Left            =   120
            Picture         =   "frmMain.frx":0E1C
            Top             =   4200
            Width           =   525
         End
         Begin VB.Label lblUninstall 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Uninstall the product..."
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
            MouseIcon       =   "frmMain.frx":1DFA
            MousePointer    =   99  'Custom
            TabIndex        =   11
            Top             =   5640
            Width           =   2160
         End
         Begin VB.Image imgUninstall 
            Height          =   360
            Left            =   1060
            MouseIcon       =   "frmMain.frx":1F4C
            MousePointer    =   99  'Custom
            Picture         =   "frmMain.frx":209E
            Top             =   5550
            Width           =   270
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Set Skin:"
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
            Left            =   2760
            TabIndex        =   28
            Top             =   4875
            Width           =   540
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Please select the NTFS drive(s) and  click Connect to Service. (Drive can be Fixed or Removable)"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   2760
            TabIndex        =   19
            Top             =   100
            Width           =   1575
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
            TabIndex        =   23
            Top             =   3470
            Width           =   2970
            WordWrap        =   -1  'True
         End
         Begin VB.Image Image6 
            Height          =   255
            Left            =   765
            Picture         =   "frmMain.frx":2620
            Top             =   3470
            Width           =   240
         End
         Begin VB.Image Image9 
            Height          =   570
            Left            =   120
            Picture         =   "frmMain.frx":2992
            Top             =   2560
            Width           =   510
         End
         Begin VB.Image imgReloadFilter 
            Height          =   285
            Left            =   1065
            MouseIcon       =   "frmMain.frx":3944
            MousePointer    =   99  'Custom
            Picture         =   "frmMain.frx":3A96
            Top             =   3050
            Width           =   300
         End
         Begin VB.Image imgTestService 
            Height          =   390
            Left            =   1065
            MouseIcon       =   "frmMain.frx":3F4C
            MousePointer    =   99  'Custom
            Picture         =   "frmMain.frx":409E
            Top             =   2585
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
            Left            =   1090
            LinkTimeout     =   100
            MouseIcon       =   "frmMain.frx":47C8
            MousePointer    =   99  'Custom
            TabIndex        =   5
            Top             =   2655
            Width           =   2140
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
            MouseIcon       =   "frmMain.frx":491A
            MousePointer    =   99  'Custom
            TabIndex        =   6
            Top             =   3095
            Width           =   2925
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Set Transparency:"
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
            Left            =   1065
            TabIndex        =   22
            Top             =   4870
            Width           =   1110
         End
         Begin VB.Line Line7 
            BorderColor     =   &H00C0C0C0&
            X1              =   840
            X2              =   4320
            Y1              =   4035
            Y2              =   4035
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "General"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   21
            Top             =   3910
            Width           =   660
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Check"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   20
            Top             =   2275
            Width           =   510
         End
         Begin VB.Image Image7 
            Height          =   675
            Left            =   900
            Picture         =   "frmMain.frx":4A6C
            Stretch         =   -1  'True
            Top             =   0
            Width           =   3480
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00C0C0C0&
            X1              =   120
            X2              =   4320
            Y1              =   1680
            Y2              =   1680
         End
         Begin VB.Image imgReset 
            Height          =   330
            Left            =   120
            MouseIcon       =   "frmMain.frx":BC96
            MousePointer    =   99  'Custom
            Picture         =   "frmMain.frx":BDE8
            Top             =   1840
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
            MouseIcon       =   "frmMain.frx":C4B2
            MousePointer    =   99  'Custom
            TabIndex        =   3
            Top             =   1870
            Width           =   1410
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00C0C0C0&
            X1              =   720
            X2              =   4320
            Y1              =   2400
            Y2              =   2400
         End
         Begin VB.Label lblADFilter 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Connect to Service"
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
            Left            =   2190
            LinkTimeout     =   100
            MouseIcon       =   "frmMain.frx":C604
            MousePointer    =   99  'Custom
            TabIndex        =   4
            Top             =   1875
            Width           =   1875
         End
         Begin VB.Image imgADFilter 
            Height          =   405
            Left            =   2080
            MouseIcon       =   "frmMain.frx":C756
            MousePointer    =   99  'Custom
            Picture         =   "frmMain.frx":C8A8
            Top             =   1800
            Width           =   375
         End
         Begin VB.Image imgRefresh 
            Height          =   330
            Left            =   2760
            MouseIcon       =   "frmMain.frx":D0EE
            MousePointer    =   99  'Custom
            Picture         =   "frmMain.frx":D240
            Top             =   1000
            Width           =   300
         End
         Begin VB.Label lblRefresh 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Refresh List"
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
            Left            =   2880
            LinkTimeout     =   100
            MouseIcon       =   "frmMain.frx":D7AA
            MousePointer    =   99  'Custom
            TabIndex        =   1
            Top             =   1050
            Width           =   1210
         End
      End
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00C0C0C0&
      X1              =   200
      X2              =   4605
      Y1              =   7245
      Y2              =   7245
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
      Left            =   645
      MouseIcon       =   "frmMain.frx":D8FC
      MousePointer    =   99  'Custom
      TabIndex        =   25
      Top             =   7440
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
      Left            =   165
      TabIndex        =   24
      Top             =   7440
      Width           =   420
   End
   Begin VB.Image imgClose 
      Height          =   330
      Left            =   3765
      MouseIcon       =   "frmMain.frx":DA4E
      MousePointer    =   99  'Custom
      Top             =   7350
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
      Left            =   3900
      LinkTimeout     =   100
      MouseIcon       =   "frmMain.frx":DBA0
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   7410
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
      MouseIcon       =   "frmMain.frx":DCF2
      MousePointer    =   99  'Custom
      TabIndex        =   18
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
      MouseIcon       =   "frmMain.frx":DE44
      MousePointer    =   99  'Custom
      TabIndex        =   17
      Top             =   600
      Width           =   645
   End
   Begin VB.Image imgCloseBottom 
      Height          =   315
      Left            =   4320
      MouseIcon       =   "frmMain.frx":DF96
      MousePointer    =   99  'Custom
      ToolTipText     =   "Close"
      Top             =   62
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
      BorderColor     =   &H00000000&
      X1              =   10
      X2              =   4770
      Y1              =   400
      Y2              =   400
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      TabIndex        =   16
      Top             =   105
      Width           =   45
   End
   Begin VB.Line Line4 
      X1              =   4785
      X2              =   4785
      Y1              =   360
      Y2              =   7800
   End
   Begin VB.Line Line5 
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   7800
   End
   Begin VB.Image Image5 
      Height          =   75
      Left            =   -840
      Stretch         =   -1  'True
      Top             =   7725
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
      Caption         =   "Operation and Configuration:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   195
      TabIndex        =   13
      Top             =   585
      Width           =   2115
   End
   Begin VB.Image Image3 
      Height          =   525
      Left            =   0
      Stretch         =   -1  'True
      Top             =   420
      Width           =   4780
   End
   Begin VB.Image Image1 
      Height          =   6975
      Left            =   0
      Stretch         =   -1  'True
      Top             =   840
      Width           =   75
   End
   Begin VB.Image Image4 
      Height          =   6975
      Left            =   4725
      Stretch         =   -1  'True
      Top             =   840
      Width           =   75
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================
'=  Developed by P. G. B. Prasanna       =
'=  A Software Developer from Sri Lanka  =
'=  E-Mail: pgbsoft@gmail.com            =
'=========================================

Public Sub chkSelectAll_Click()
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

Private Sub chkSelectAll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Move_Effect
End Sub

Private Sub chkStartup_Click()
If chkStartup.Value = 1 Then
    Call Set_Reg_Data(R_Startup, Get_Full_App_Path & " /chkadmin", 1)
Else
   Call Remove_Reg_Data(R_Startup)
End If
End Sub

Private Sub chkUninstallService_Click()
Call Set_Reg_Data(R_LocApp, chkUninstallService.Value)
Select Case chkUninstallService.Value
    Case 0: intService_Uninstall = 0
    Case 1: intService_Uninstall = 1
End Select
End Sub

Private Sub chkUninstallService_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Move_Effect
End Sub
Private Sub cmbSkin_Click()
If intAdmin = 1 Then Set_Reg_Data R_Skin, cmbSkin.ListIndex
intSkin = cmbSkin.ListIndex
Apply_Theme 1, Me, intSkin
End Sub

Private Sub cmbTransparency_Click()
If intAdmin = 1 Then Set_Reg_Data R_Tpncy, cmbTransparency.ListIndex
Select Case cmbTransparency.ListIndex
    Case 0: Set_Manual_Transparency 255: bTrans_P_Level_Limit = 255
    Case 1: Set_Manual_Transparency 240: bTrans_P_Level_Limit = 240
    Case 2: Set_Manual_Transparency 200: bTrans_P_Level_Limit = 200
    Case 3: Set_Manual_Transparency 150: bTrans_P_Level_Limit = 150
End Select
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
If intAdmin = 0 Then: Exit Sub
Mouse_Move_Effect
lblADFilter.ForeColor = &H80FF&
lblADFilter_MouseDown 0, 0, 0, 0
lblADFilter_MouseUp 0, 0, 0, 0
'Mouse_Move_Effect
End Sub
Private Sub Form_Activate()
On Error Resume Next
lstDrivesList.ListIndex = -1
If intAdmin = 0 Then: cmdClose.SetFocus
If intIsFirstFlag = 1 Then frmFeatures.Show 1
End Sub

Private Sub Form_Load()
'On Error Resume Next
Me.Label3.Caption = AppTitles(0)
Me.Caption = AppTitles(0)
Get_Requried_Drives
Set_Round Me
Load_Settings
Apply_Theme 1, Me, intSkin

Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2

'chkUninstallService.Value = intService_Uninstall
If intAdmin = 0 Then
    Enable_Disable_Control False
    'imgCloseTop.Enabled = True
    'imgClose.Enabled = True
    'lblClose.Enabled = True
End If

On Error Resume Next
lStartStyle = GetWindowLong(hwnd, GWL_EXSTYLE)
SetWindowLong hwnd, GWL_EXSTYLE, lStartStyle Or WS_EX_LAYERED
bTrans_P_Level = 0
Timer1.Enabled = True
intActionMode = 1
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Move_Control Me
If Button = 1 Then
    Getmove Me
End If
Mouse_Move_Effect
End Sub

Private Sub Form_Unload(Cancel As Integer)
If intAdmin = 1 Then
    chkSelectAll_Click
    chkSelectAll.Value = 0
    Attach_Detach_Filter
    Control_MiniFilter_Driver 4
    If intService_Uninstall = 1 Then: Control_MiniFilter_Driver 6: Config_Drv 2
End If
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Move_Control Me
If Button = 1 Then
    Getmove Me
End If
Mouse_Move_Effect
End Sub

Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Move_Control Me
If Button = 1 Then
    Getmove Me
End If
Mouse_Move_Effect
End Sub

Private Sub Image4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Move_Control Me
If Button = 1 Then
    Getmove Me
End If
Mouse_Move_Effect
End Sub

Private Sub Image5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Move_Control Me
If Button = 1 Then
    Getmove Me
End If
Mouse_Move_Effect
End Sub
Private Sub Image6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Move_Control Me
If Button = 1 Then
    Getmove Me
End If
End Sub
Private Sub Image7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Move_Control Me
If Button = 1 Then
    Getmove Me
End If
Mouse_Move_Effect
End Sub

Private Sub Image8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Move_Control Me
If Button = 1 Then
    Getmove Me
End If
Mouse_Move_Effect

End Sub

Private Sub Image9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Move_Control Me
If Button = 1 Then
    Getmove Me
End If
Mouse_Move_Effect
End Sub

Private Sub imgADFilter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblADFilter_MouseDown 0, 0, 0, 0
End Sub

Private Sub imgADFilter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblADFilter_MouseMove 0, 0, 0, 0
End Sub

Private Sub imgADFilter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblADFilter_MouseUp 0, 0, 0, 0
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
Private Sub imgRefresh_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblRefresh_MouseDown 0, 0, 0, 0
End Sub

Private Sub imgRefresh_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblRefresh_MouseMove 0, 0, 0, 0
End Sub

Private Sub imgRefresh_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblRefresh_MouseUp 0, 0, 0, 0
End Sub

Private Sub imgReloadFilter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblRecoverFilterDrv_MouseDown 0, 0, 0, 0
End Sub

Private Sub imgReloadFilter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblRecoverFilterDrv_MouseMove 0, 0, 0, 0
End Sub

Private Sub imgReloadFilter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblRecoverFilterDrv_MouseUp 0, 0, 0, 0
End Sub

Private Sub imgReset_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblReset_MouseDown 0, 0, 0, 0
End Sub

Private Sub imgReset_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblReset_MouseMove 0, 0, 0, 0
End Sub

Private Sub imgReset_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblReset_MouseUp 0, 0, 0, 0
End Sub

Private Sub imgTestService_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblTestService_MouseDown 0, 0, 0, 0
End Sub

Private Sub imgTestService_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblTestService_MouseMove 0, 0, 0, 0
End Sub

Private Sub imgTestService_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblTestService_MouseUp 0, 0, 0, 0
End Sub

Private Sub imgTop_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Move_Control Me
If Button = 1 Then
    Getmove Me
End If
Mouse_Move_Effect
End Sub
Private Sub imgUninstall_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblUninstall_MouseDown 0, 0, 0, 0
End Sub

Private Sub imgUninstall_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblUninstall_MouseMove 0, 0, 0, 0
End Sub
Private Sub imgUninstall_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblUninstall_MouseUp 0, 0, 0, 0
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Move_Control Me
If Button = 1 Then
    Getmove Me
End If
Mouse_Move_Effect
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Move_Control Me
If Button = 1 Then
    Getmove Me
End If
Mouse_Move_Effect
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Move_Control Me
If Button = 1 Then
    Getmove Me
End If
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Move_Control Me
If Button = 1 Then
    Getmove Me
End If
Mouse_Move_Effect
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Move_Effect
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Move_Control Me
If Button = 1 Then
    Getmove Me
End If
Mouse_Move_Effect
End Sub

Private Sub lblAbout_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Up_Down_Handling lblAbout
End Sub

Private Sub lblAbout_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Move_Effect
End Sub

Private Sub lblAbout_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Up_Down_Handling lblAbout, , 1
frmAbout.Show 1
End Sub

Private Sub lblADFilter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Up_Down_Handling lblADFilter, imgADFilter
End Sub

Private Sub lblADFilter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Move_Effect
lblADFilter.ForeColor = &H80FF&
End Sub

Private Sub lblADFilter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Up_Down_Handling lblADFilter, imgADFilter, 1

If Check_For_Selected_Drives = False Then Exit Sub

Enable_Disable_Control False
Attach_Detach_Filter
CenterMessage Me, "Filter service attached to the selected drive(s)..." & vbCrLf & _
                  "Now you can access any NTFS-Protected folder/file on the" & vbCrLf & _
                  "selected drive(s).", vbInformation, App.Title
Enable_Disable_Control True
SHChangeNotify SHCNE_ASSOCCHANGED, SHCNF_FLUSH, 0, 0
cmdOK.SetFocus
End Sub

Private Sub lblClose_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Up_Down_Handling lblClose, imgClose
End Sub

Private Sub lblClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Move_Effect
lblClose.ForeColor = &H80FF&
End Sub

Private Sub lblClose_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Up_Down_Handling lblClose, imgClose, 1
Unload Me
End Sub

Private Sub lblEmail_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Up_Down_Handling lblEmail
End Sub

Private Sub lblEmail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Move_Effect
lblEmail.ForeColor = &HC0&
End Sub

Private Sub lblEmail_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Up_Down_Handling lblEmail, , 1
Call Mail_To
End Sub

Private Sub lblHelp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Up_Down_Handling lblHelp
End Sub

Private Sub lblHelp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Move_Effect
End Sub

Private Sub lblHelp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Up_Down_Handling lblHelp, , 1
frmFeatures.Show 1
End Sub
Private Sub lstDriveList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Move_Control Me
If Button = 1 Then
    Getmove Me
End If
End Sub

Private Sub lblRecoverFilterDrv_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Up_Down_Handling lblRecoverFilterDrv, imgReloadFilter
End Sub

Private Sub lblRecoverFilterDrv_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Move_Effect
lblRecoverFilterDrv.ForeColor = &H80FF&
End Sub

Private Sub lblRecoverFilterDrv_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
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
cmdOK.SetFocus
End Sub

Private Sub lblRefresh_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Up_Down_Handling lblRefresh, imgRefresh
End Sub

Private Sub lblRefresh_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Move_Effect
lblRefresh.ForeColor = &H80FF&
End Sub

Private Sub lblRefresh_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblRefresh.Left = 2880: lblRefresh.Top = 1050
imgRefresh.Left = 2760: imgRefresh.Top = 1000
Enable_Disable_Control False
chkSelectAll_Click
chkSelectAll.Value = 0
Attach_Detach_Filter
Get_Requried_Drives
Enable_Disable_Control True
cmdOK.SetFocus
End Sub

Private Sub lblReset_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Up_Down_Handling lblReset, imgReset
End Sub

Private Sub lblReset_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Move_Effect
lblReset.ForeColor = &H80FF&
End Sub

Private Sub lblReset_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
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
cmdOK.SetFocus
End Sub

Private Sub lblTestService_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Up_Down_Handling lblTestService, imgTestService
End Sub

Private Sub lblTestService_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Move_Effect
lblTestService.ForeColor = &H80FF&
End Sub

Private Sub lblTestService_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Up_Down_Handling lblTestService, imgTestService, 1

Enable_Disable_Control False
Dim sysVI_On_Sys As String
sysVI_On_Sys = Environ$("SystemDrive") & "\System Volume Information"
If PathIsDirectory(sysVI_On_Sys) = 0 Then CenterMessage Me, "Unexpected Error", vbCritical, AppTitles(3): _
                                     Enable_Disable_Control True: cmdOK.SetFocus: Exit Sub
    
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
       "If succeeded, filter service is working properly and" & vbCrLf & _
       "if not, click on Recover/Restore Filter Service and try again.", vbInformation, App.Title
If Found_Sys_Drive = 1 Then Attach_Detach_Filter: Shell "Explorer " & sysVI_On_Sys, vbNormalFocus: Enable_Disable_Control True: cmdOK.SetFocus
End Sub

Private Sub lblUninstall_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Up_Down_Handling lblUninstall, imgUninstall
End Sub

Private Sub lblUninstall_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Move_Effect
lblUninstall.ForeColor = &H80FF&
End Sub

Private Sub lblUninstall_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Up_Down_Handling lblUninstall, imgUninstall, 1
If CenterMessage(Me, "Are you sure you want to completely remove " & _
                AppTitles(0) & " and all of its components?", vbQuestion + vbYesNo + _
                vbDefaultButton2, AppTitles(0) & " Uninstall") = vbYes Then
                cmbTransparency = cmbTransparency.List(0)
                lblTestService.Visible = False: imgTestService.Visible = False
                frmUninstall.Show 1
End If
Mouse_Move_Effect
cmdOK.SetFocus
End Sub
Private Sub lstDrivesList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Move_Effect
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
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
                                'If removable drive is ready we check for NTFS
                                If intErrNotReady = 0 Then
                                    If Determine_NTFS(Drv.DriveLetter & ":\") = True Then
                                        'If the drive is not a virtual drive we add it to the list.(No need to add virtual drives to the list
                                        'as they are associated with physical drives.)
                                        If Is_Virtual_Drive(Drv.DriveLetter & ":\") = False Then
                                            lstDrivesList.AddItem Vol_N & " (" & Drv.DriveLetter & ":)"
                                        End If
                                    End If
                                Else
                                   'A Removable Drive will be detected as not ready by the system, even when you do not have
                                   'the permission to access it. So, we should add the drive to the list in order to use it
                                   'with the program.
                                   lstDrivesList.AddItem Vol_N & " (" & Drv.DriveLetter & ":)"
                                End If
                                
                    Case 2: Vol_N = Drv.VolumeName
                            If Vol_N = "" Then Vol_N = "Local Disk"
                                'If fixed drive is ready we check for NTFS
                                If intErrNotReady = 0 Then
                                    If Determine_NTFS(Drv.DriveLetter & ":\") = True Then
                                        'If the drive is not a virtual drive we add it to the list.(No need to add virtual drives to the list
                                        'as they are associated with physical drives.)
                                        If Is_Virtual_Drive(Drv.DriveLetter & ":\") = False Then
                                            lstDrivesList.AddItem Vol_N & " (" & Drv.DriveLetter & ":)"
                                        End If
                                    End If
                                Else
                                   'A Fixed Drive will be detected as not ready by the system, when you do not have
                                   'the permission to access it. So, we should add the drive to the list in order to use it
                                   'with the program.
                                    lstDrivesList.AddItem Vol_N & " (" & Drv.DriveLetter & ":)"
                                End If
                End Select
        End If
    Next
    
Exit Sub
ErrDrive:
intErrNotReady = 1
Resume Next
End Sub
'Determine that the File System of the drive is NTFS
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
'Determine whether the checking drive is virtual one or not
Private Function Is_Virtual_Drive(ByVal sDrive As String) As Boolean
Const MAX_PATH = 260
Dim bDrive() As Byte
Dim bResult() As Byte
Dim lR As Long
Dim sDeviceName As String

Is_Virtual_Drive = False

If Right(sDrive, 1) = "\" Then
    If Len(sDrive) > 1 Then
         sDrive = Left(sDrive, Len(sDrive) - 1)
    End If
End If
bDrive = sDrive
ReDim Preserve bDrive(0 To UBound(bDrive) + 2) As Byte
ReDim bResult(0 To MAX_PATH * 2 + 1) As Byte
lR = QueryDosDeviceW(VarPtr(bDrive(0)), VarPtr(bResult(0)), MAX_PATH)
    If (lR > 2) Then
        sDeviceName = bResult
        sDeviceName = Left(sDeviceName, lR - 2)
            'Following condition determines whether the checking drive is a virtual drive
            '(Eg: A Physical Drive returns something like Device\HarddiskVolume1 while
            'a Virtual Drive returns something like \??\C:\Windows)
            If Mid(sDeviceName, 2, 1) = "?" Then Is_Virtual_Drive = True
    End If
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
Public Sub Attach_Detach_Filter()
Dim i As Integer
    For i = 0 To lstDrivesList.ListCount - 1
        strDrvL = Mid(lstDrivesList.List(i), Len(lstDrivesList.List(i)) - 2, 2)
        If lstDrivesList.Selected(i) Then
            Control_MiniFilter_Driver 1, strDrvL
        Else
            Control_MiniFilter_Driver 2, strDrvL
        End If
        Sleep 10
    Next
End Sub
Public Sub Enable_Disable_Control(ByVal val As Boolean)
lstDrivesList.Enabled = val: lstDrivesList.Refresh
chkSelectAll.Enabled = val: chkSelectAll.Refresh
lblReset.Enabled = val: lblReset.Refresh
imgReset.Enabled = val: imgReset.Refresh
lblADFilter.Enabled = val: lblADFilter.Refresh
imgADFilter.Enabled = val: imgADFilter.Refresh
lblRecoverFilterDrv.Enabled = val: lblRecoverFilterDrv.Refresh
imgReloadFilter.Enabled = val: imgReloadFilter.Refresh
If intAdmin = 1 Then
    cmdClose.Enabled = val: cmdClose.Refresh
    imgCloseTop.Enabled = val: imgCloseTop.Refresh
    imgClose.Enabled = val: imgClose.Refresh
    lblClose.Enabled = val: lblClose.Refresh
End If
lblTestService.Enabled = val: lblTestService.Refresh
imgTestService.Enabled = val: imgTestService.Refresh
lblRefresh.Enabled = val: lblRefresh.Refresh
imgRefresh.Enabled = val: imgRefresh.Refresh
chkUninstallService.Enabled = val: chkUninstallService.Refresh
chkStartup.Enabled = val: chkStartup.Refresh
'cmbTransparency.Enabled = val: cmbTransparency.Refresh
'cmbSkin.Enabled = val: cmbSkin.Refresh
lblUninstall.Enabled = val: lblUninstall.Refresh
imgUninstall.Enabled = val: imgUninstall.Refresh
End Sub
Public Sub Mouse_Move_Effect()
lblRefresh.ForeColor = vbBlack
lblReset.ForeColor = vbBlack
lblADFilter.ForeColor = vbBlack
lblTestService.ForeColor = vbBlack
lblRecoverFilterDrv.ForeColor = vbBlack
lblClose.ForeColor = vbBlack
lblEmail.ForeColor = vbBlack
lblUninstall.ForeColor = vbBlack
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
Public Sub Load_Settings()
Dim intStartup, intTP As Integer
Dim i, c1, c2 As Integer
Dim CmbVal(6) As String
Dim validate As Boolean

CmbVal(0) = "-None-": CmbVal(1) = "Low"
CmbVal(2) = "Medium": CmbVal(3) = "High"
CmbVal(4) = "Layer-Green": CmbVal(5) = "Layer-Red"
CmbVal(6) = "Layer-Yellow"

cmbTransparency.Clear
cmbSkin.Clear

For c1 = 0 To 3: cmbTransparency.AddItem CmbVal(c1): Next
For c2 = 4 To 6: cmbSkin.AddItem CmbVal(c2): Next


If intIsFirstFlag = 1 Then
    intStartup = 1
    intTP = 1
    intSkin = 0
    
Else
   'Reading for the Registry Values
    If Read_Reg_Data(R_Startup) = 0 And UCase(Read_Reg_Data(R_Startup)) <> UCase(Get_Full_App_Path) & " /chkadmin" Then
        intStartup = 0
    Else
        intStartup = 1
    End If
    intTP = Read_Reg_Data(R_Tpncy)
       'Validate value for Transparency - Prevent from loading the program with improper values.
        validate = False
        For v = 0 To 3
            If intTP = v Then validate = True: Exit For
        Next
        If validate = False Then intTP = 1
        
    intSkin = Read_Reg_Data(R_Skin)
        'Validate value for Skin - Prevent from loading the program with improper values.
        validate = False
        For v = 0 To 2
            If intSkin = v Then validate = True: Exit For
        Next
        If validate = False Then intSkin = 0
End If

'Applying Values
chkUninstallService.Value = intService_Uninstall
chkStartup = intStartup
'For i = 0 To 3: If intTP = i Then cmbTransparency = cmbTransparency.List(i): Exit For
'Next
cmbTransparency = Choose((intTP + 1), _
                        cmbTransparency.List(0), _
                        cmbTransparency.List(1), _
                        cmbTransparency.List(2), _
                        cmbTransparency.List(3))
'cmbSkin = IIf(intSkin = 0, cmbSkin.List(0), cmbSkin.List(1))
For i = 0 To 2: If intSkin = i Then cmbSkin = cmbSkin.List(i): Exit For
Next
End Sub
Private Sub Set_Manual_Transparency(TpLevel As Byte)
SetLayeredWindowAttributes Me.hwnd, 0, TpLevel, LWA_ALPHA
End Sub
