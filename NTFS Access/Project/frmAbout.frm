VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1365
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4545
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   1365
   ScaleWidth      =   4545
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   120
      Top             =   360
   End
   Begin VB.PictureBox picShape 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   0
      ScaleHeight     =   1575
      ScaleWidth      =   4575
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NTFS Protected Folder/File Access"
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
         Left            =   360
         TabIndex        =   8
         Top             =   120
         Width           =   2865
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Freeware"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00321DC7&
         Height          =   165
         Left            =   3480
         TabIndex        =   7
         Top             =   650
         Width           =   750
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "E-Mail:"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   165
         Left            =   720
         TabIndex        =   6
         Top             =   900
         Width           =   435
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright © 2011 Bandula"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   165
         Left            =   960
         TabIndex        =   5
         Top             =   405
         Width           =   1545
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Programmer:"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   165
         Left            =   720
         TabIndex        =   4
         Top             =   620
         Width           =   825
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "P. G. B. Prasanna "
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
         Left            =   1605
         TabIndex        =   3
         Top             =   620
         Width           =   1320
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "V 1.0"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00321DC7&
         Height          =   165
         Left            =   3525
         TabIndex        =   2
         Top             =   840
         Width           =   270
      End
      Begin VB.Label lblContact 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   165
         Left            =   1195
         MouseIcon       =   "frmAbout.frx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   1
         Top             =   900
         Width           =   1260
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================
'=  Developed by P. G. B. Prasanna       =
'=  A Software Developer from Sri Lanka  =
'=  E-Mail: pgbsoft@gmail.com            =
'=========================================

Dim l, k As String

Private Const DIB_RGB_COLORS = 0&
Private Const BI_RGB = 0&

Private Const pixR As Integer = 3
Private Const pixG As Integer = 2
Private Const pixB As Integer = 1

Private Sub UnRGB(ByRef color As Long, ByRef R As Byte, ByRef g As Byte, ByRef b As Byte)
    R = color And &HFF&
    g = (color And &HFF00&) \ &H100&
    b = (color And &HFF0000) \ &H10000
End Sub

Private Sub ShapeForm(ByVal pic As PictureBox, ByVal transparent_color As Long)
On Error Resume Next
Dim bytes_per_scanLine As Integer
Dim wid As Long
Dim hgt As Long
Dim bitmap_info As BITMAPINFO
Dim pixels() As Byte
Dim buffer() As Byte
Dim transparent_r As Byte
Dim transparent_g As Byte
Dim transparent_b As Byte
Dim border_width As Single
Dim title_height As Single
Dim x0 As Long
Dim y0 As Long
Dim start_c As Integer
Dim stop_c As Integer
Dim R As Integer
Dim C As Integer
Dim combined_rgn As Long
Dim new_rgn As Long

    ScaleMode = vbPixels
    pic.ScaleMode = vbPixels
    pic.AutoRedraw = True
    pic.Picture = pic.Image

    ' Prepare the bitmap description.
    wid = pic.ScaleWidth
    hgt = pic.ScaleHeight
    With bitmap_info.bmiHeader
        .biSize = 40
        .biWidth = wid
        ' Use negative height to scan top-down.
        .biHeight = -hgt
        .biPlanes = 1
        .biBitCount = 32
        .biCompression = BI_RGB
        bytes_per_scanLine = ((((.biWidth * .biBitCount) + 31) \ 32) * 4)
        .biSizeImage = bytes_per_scanLine * hgt
    End With

    ' Load the bitmap's data.
    ReDim pixels(1 To 4, 1 To wid, 1 To hgt)
    GetDIBits pic.hDC, pic.Image, _
        0, hgt, pixels(1, 1, 1), _
        bitmap_info, DIB_RGB_COLORS

    ' Process the pixels.
    ' Break the tansparent color apart.
    UnRGB transparent_color, transparent_r, transparent_g, transparent_b

    ' Find the form's corner.
    border_width = (ScaleX(Width, vbTwips, vbPixels) - ScaleWidth) / 2
    title_height = ScaleX(Height, vbTwips, vbPixels) - border_width - ScaleHeight

    ' Find the picture's corner.
    x0 = pic.Left + border_width '- 1
   y0 = pic.Top + title_height '- 1
    ' Create the form's regions.
    For R = 1 To hgt
        ' Create a region for this row.
        C = 1
        Do While C <= wid
            start_c = 1
            stop_c = 1

            ' Find the next non-white column.
            Do While C <= wid
                If pixels(pixR, C, R) <> transparent_r Or _
                   pixels(pixG, C, R) <> transparent_g Or _
                   pixels(pixB, C, R) <> transparent_b _
                Then
                    Exit Do
                End If
                C = C + 1
            Loop
            start_c = C

            ' Find the next white column.
            Do While C <= wid
                If pixels(pixR, C, R) = transparent_r And _
                   pixels(pixG, C, R) = transparent_g And _
                   pixels(pixB, C, R) = transparent_b _
                Then
                    Exit Do
                End If
                C = C + 1
            Loop
            stop_c = C

            ' Make a region from start_c to stop_c.
            If start_c <= wid Then
                If stop_c > wid Then stop_c = wid

                ' Create the region.
                new_rgn = CreateRectRgn( _
                    start_c + x0, R + y0, _
                    stop_c + x0, R + y0 + 1)

                ' Add it to what we have so far.
                If combined_rgn = 0 Then
                    combined_rgn = new_rgn
                Else
                    CombineRgn combined_rgn, _
                        combined_rgn, new_rgn, RGN_OR
                    DeleteObject new_rgn
                End If
            End If
        Loop
    Next R

    ' Restrict the form to the region.
    SetWindowRgn hwnd, combined_rgn, True
    DeleteObject combined_rgn
End Sub

Private Sub Command1_Click()
End
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then Unload Me
    If KeyAscii = vbKeyReturn Then Unload Me
End Sub

Private Sub Form_Load()
Apply_Theme 3, Me, intSkin
KeyPreview = True
ShapeForm picShape, &HFFFFFF
If intIsFirstFlag = 1 Then
    Me.Left = frmFeatures.Left + (frmFeatures.Width - Me.Width) / 2 + 500
    Me.Top = frmFeatures.Top + (frmFeatures.Height - Me.Height) / 2 + 150
    intIsFirstFlag = 0
Else
    Me.Left = frmMain.Left + (frmMain.Width - Me.Width) / 2 + 500
    Me.Top = frmMain.Top + (frmMain.Height - Me.Height) / 2 + 150
End If
On Error Resume Next
lStartStyle = GetWindowLong(hwnd, GWL_EXSTYLE)
SetWindowLong hwnd, GWL_EXSTYLE, lStartStyle Or WS_EX_LAYERED
bTrans_P_Level = 0
Timer1.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If intselectDriveLoaded = 1 Then
         frmMain.lblAbout.MousePointer = vbCustom
         SetWindowPos frmMain.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
  End If
  intMenuLoad = 0
 End Sub
Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    Getmove Me
    Unload Me
End If
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Unload Me
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    Getmove Me
    Unload Me
End If
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    Getmove Me
    Unload Me
End If
End Sub
Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    Getmove Me
    Unload Me
End If
End Sub

Private Sub Label4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Unload Me
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    Getmove Me
    Unload Me
End If
End Sub

Private Sub Label5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Unload Me
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    Getmove Me
    Unload Me
End If

End Sub

Private Sub Label6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Unload Me
End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    Getmove Me
    Unload Me
End If
End Sub

Private Sub Label7_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Unload Me
End Sub

Private Sub lblContact_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblContact.Top = lblContact.Top + 1
lblContact.Left = lblContact.Left + 1
End Sub

Private Sub lblContact_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblContact.Top = 60
lblContact.Left = 79.66666
Call Mail_To
End Sub

Private Sub picShape_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    Getmove Me
    Unload Me
End If
End Sub

Private Sub picShape_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Unload Me
End Sub

Private Sub Timer1_Timer()
Control_Transparent_Effect Me
End Sub
