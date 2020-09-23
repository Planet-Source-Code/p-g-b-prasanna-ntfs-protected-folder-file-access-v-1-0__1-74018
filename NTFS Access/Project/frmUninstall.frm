VERSION 5.00
Begin VB.Form frmUninstall 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   450
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2535
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   450
   ScaleWidth      =   2535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label lblProgress1 
      Height          =   30
      Left            =   135
      TabIndex        =   2
      Top             =   255
      Width           =   15
   End
   Begin VB.Label lblProgress2 
      Height          =   60
      Left            =   135
      TabIndex        =   1
      Top             =   285
      Width           =   15
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   30
   End
   Begin VB.Image imgPro1 
      Height          =   120
      Left            =   120
      Picture         =   "frmUninstall.frx":0000
      Stretch         =   -1  'True
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "frmUninstall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
Uninstall_With_Progress
End Sub
Private Sub Form_Load()
lblProgress1.Width = 0
lblProgress2.Width = 0
Set_Round Me
Apply_Theme 4, Me, intSkin
'lStartStyle = GetWindowLong(hwnd, GWL_EXSTYLE)
'SetWindowLong hwnd, GWL_EXSTYLE, lStartStyle Or WS_EX_LAYERED
'SetLayeredWindowAttributes Me.hwnd, 0, bTrans_P_Level_Limit, LWA_ALPHA
End Sub
Private Sub Uninstall_With_Progress()
Dim i As Integer
Dim strUnStatus(3) As String

strUnStatus(0) = "Preparing to uninstall, please wait..."
strUnStatus(1) = "Unloding services, please wait..."
strUnStatus(2) = "Deleting files, please wait..."
strUnStatus(3) = "Removing registry entries, please wait..."

For i = 0 To 14
    Select Case i
        Case 0: lblStatus.Caption = strUnStatus(0)
                frmMain.Enable_Disable_Control False
                lblProgress1.Width = i / 14 * 2265: lblProgress2.Width = i / 14 * 2265
                lblProgress1.Refresh: lblProgress2.Refresh: Sleep 100
        Case 1: frmMain.chkSelectAll_Click
                lblProgress1.Width = i / 14 * 2265: lblProgress2.Width = i / 14 * 2265
                lblProgress1.Refresh: lblProgress2.Refresh: Sleep 100
        Case 2: frmMain.chkSelectAll.Value = 0
                lblProgress1.Width = i / 14 * 2265: lblProgress2.Width = i / 14 * 2265
                lblProgress1.Refresh: lblProgress2.Refresh: Sleep 100
        Case 3: frmMain.Attach_Detach_Filter
                lblProgress1.Width = i / 14 * 2265: lblProgress2.Width = i / 14 * 2265
                lblProgress1.Refresh: lblProgress2.Refresh: Sleep 100
        Case 4: lblStatus.Caption = strUnStatus(1)
                Control_MiniFilter_Driver 4
                lblProgress1.Width = i / 14 * 2265: lblProgress2.Width = i / 14 * 2265
                lblProgress1.Refresh: lblProgress2.Refresh: Sleep 100
        Case 5: Control_MiniFilter_Driver 5
                lblProgress1.Width = i / 14 * 2265: lblProgress2.Width = i / 14 * 2265
                lblProgress1.Refresh: lblProgress2.Refresh: Sleep 100
        Case 6: Control_MiniFilter_Driver 6
                lblProgress1.Width = i / 14 * 2265: lblProgress2.Width = i / 14 * 2265
                lblProgress1.Refresh: lblProgress2.Refresh: Sleep 100
        Case 7: lblStatus.Caption = strUnStatus(2)
                Config_Drv 2
                SHChangeNotify SHCNE_ASSOCCHANGED, SHCNF_FLUSH, 0, 0
                lblProgress1.Width = i / 14 * 2265: lblProgress2.Width = i / 14 * 2265
                lblProgress1.Refresh: lblProgress2.Refresh: Sleep 100
        Case 8: lblStatus.Caption = strUnStatus(3)
                Call Remove_Reg_Data(R_LocApp)
                lblProgress1.Width = i / 14 * 2265: lblProgress2.Width = i / 14 * 2265
                lblProgress1.Refresh: lblProgress2.Refresh: Sleep 100
        Case 9: Call Remove_Reg_Data(R_IsFirst)
                lblProgress1.Width = i / 14 * 2265: lblProgress2.Width = i / 14 * 2265
                lblProgress1.Refresh: lblProgress2.Refresh: Sleep 100
       Case 10: Call Remove_Reg_Data(R_Tpncy)
                lblProgress1.Width = i / 14 * 2265: lblProgress2.Width = i / 14 * 2265
                lblProgress1.Refresh: lblProgress2.Refresh: Sleep 100
       Case 11: Call Remove_Reg_Data(R_Skin)
                lblProgress1.Width = i / 14 * 2265: lblProgress2.Width = i / 14 * 2265
                lblProgress1.Refresh: lblProgress2.Refresh: Sleep 100
       Case 12: Call Remove_Reg_Data(R_Startup)
                lblProgress1.Width = i / 14 * 2265: lblProgress2.Width = i / 14 * 2265
                lblProgress1.Refresh: lblProgress2.Refresh: Sleep 100
       Case 13: Call Remove_Reg_Data(Left(R_LocApp, Len(R_LocApp) - 7))
                lblProgress1.Width = i / 14 * 2265: lblProgress2.Width = i / 14 * 2265
                lblProgress1.Refresh: lblProgress2.Refresh: Sleep 100
       Case 14: Call Remove_Reg_Data(Left(R_LocApp, Len(R_LocApp) - 16))
                lblProgress1.Width = i / 14 * 2265: lblProgress2.Width = i / 14 * 2265
                lblProgress1.Refresh: lblProgress2.Refresh: Sleep 100
    End Select
Next
Call Mail_To(1)
End
End Sub
