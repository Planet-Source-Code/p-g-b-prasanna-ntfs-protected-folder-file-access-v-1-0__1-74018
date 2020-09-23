Attribute VB_Name = "ModUserInterface"
'=========================================
'=  Developed by P. G. B. Prasanna       =
'=  A Software Developer from Sri Lanka  =
'=  E-Mail: pgbsoft@gmail.com            =
'=========================================

'Make a form round
Public Sub Set_Round(s_control As Form)
mShape = 1
xp = Screen.TwipsPerPixelX
yp = Screen.TwipsPerPixelY
      
If mShape = 1 Then
    mChildFormRegion = CreateRoundRectRgn(0, 0, s_control.Width / xp, s_control.Height / yp, 9, 9)
Else
    mChildFormRegion = CreateEllipticRgn(0, 0, s_control.Width / xp, s_control.Height / yp)
End If
    
SetWindowRgn s_control.hwnd, mChildFormRegion, False
End Sub

'Procedure to help moving forms
Public Sub Getmove(p_control As Form)
Dim lngreturnvalue As Long
    Call ReleaseCapture
    lngreturnvalue = SendMessage(p_control.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End Sub

'Control transparency of forms
Public Sub Control_Transparent_Effect(appHwnd As Form)
On Error Resume Next
bTrans_P_Level = bTrans_P_Level + 5
If bTrans_P_Level >= bTrans_P_Level_Limit Then
    appHwnd.Timer1.Enabled = False
    Exit Sub
End If
SetLayeredWindowAttributes appHwnd.hwnd, 0, bTrans_P_Level, LWA_ALPHA
End Sub
'Load requried theme settings
Public Sub Apply_Theme(ByVal Form_index As Integer, Form_name As Form, ByVal LayerIndex As Integer)
With Form_name
    If LayerIndex = 0 Then
        Select Case Form_index
            Case 1: .Image3.Picture = LoadResPicture("BANNER_G", 0)
                    .imgTop = LoadResPicture("T_BANNER_G", 0)
                    .Image1.Picture = LoadResPicture("L_BORDER_G", 0)
                    .Image4.Picture = LoadResPicture("R_BORDER_G", 0)
                    .Image5.Picture = LoadResPicture("B_BORDER_G", 0)
                    .imgCloseBottom.Picture = LoadResPicture("C_BOTTOM_G", 0)
                    .imgCloseTop.Picture = LoadResPicture("C_TOP_G", 0)
                    .imgClose.Picture = LoadResPicture("IMG_CLOSE", 0)
                    .Line6.BorderColor = &H23BA16
            Case 2: .Image1.Picture = LoadResPicture("LOGO", 0)
                    .Image2.Picture = LoadResPicture("BANNER_G", 0)
                    .imgTop = LoadResPicture("T_BANNER_G", 0)
                    .Image3.Picture = LoadResPicture("L_BORDER_G", 0)
                    .Image4.Picture = LoadResPicture("R_BORDER_G", 0)
                    .Image5.Picture = LoadResPicture("B_BORDER_G", 0)
                    .imgCloseBottom.Picture = LoadResPicture("C_BOTTOM_G", 0)
                    .imgCloseTop.Picture = LoadResPicture("C_TOP_G", 0)
                    .imgClose.Picture = LoadResPicture("IMG_CLOSE", 0)
                    .Line6.BorderColor = &H23BA16
            Case 3: .picShape.Picture = LoadResPicture("L_ABOUT_G", 0)
            Case 4: .lblProgress1.BackColor = &HC0FFC0
                    .lblProgress2.BackColor = &H24CD16
        End Select
    ElseIf LayerIndex = 1 Then
        Select Case Form_index
            Case 1: .Image3.Picture = LoadResPicture("BANNER_R", 0)
                    .imgTop = LoadResPicture("T_BANNER_R", 0)
                    .Image1.Picture = LoadResPicture("L_BORDER_R", 0)
                    .Image4.Picture = LoadResPicture("R_BORDER_R", 0)
                    .Image5.Picture = LoadResPicture("B_BORDER_R", 0)
                    .imgCloseBottom.Picture = LoadResPicture("C_BOTTOM_R", 0)
                    .imgCloseTop.Picture = LoadResPicture("C_TOP_R", 0)
                    .imgClose.Picture = LoadResPicture("IMG_CLOSE", 0)
                    .Line6.BorderColor = &H6167E7
            Case 2: .Image1.Picture = LoadResPicture("LOGO", 0)
                    .Image2.Picture = LoadResPicture("BANNER_R", 0)
                    .imgTop = LoadResPicture("T_BANNER_R", 0)
                    .Image3.Picture = LoadResPicture("L_BORDER_R", 0)
                    .Image4.Picture = LoadResPicture("R_BORDER_R", 0)
                    .Image5.Picture = LoadResPicture("B_BORDER_R", 0)
                    .imgCloseBottom.Picture = LoadResPicture("C_BOTTOM_R", 0)
                    .imgCloseTop.Picture = LoadResPicture("C_TOP_R", 0)
                    .imgClose.Picture = LoadResPicture("IMG_CLOSE", 0)
                    .Line6.BorderColor = &H6167E7
            Case 3: .picShape.Picture = LoadResPicture("L_ABOUT_R", 0)
            Case 4: .lblProgress1.BackColor = &HC0C0FF
                    .lblProgress2.BackColor = &H4759FE
        End Select
    ElseIf LayerIndex = 2 Then
        Select Case Form_index
            Case 1: .Image3.Picture = LoadResPicture("BANNER_Y", 0)
                    .imgTop = LoadResPicture("T_BANNER_Y", 0)
                    .Image1.Picture = LoadResPicture("L_BORDER_Y", 0)
                    .Image4.Picture = LoadResPicture("R_BORDER_Y", 0)
                    .Image5.Picture = LoadResPicture("B_BORDER_Y", 0)
                    .imgCloseBottom.Picture = LoadResPicture("C_BOTTOM_Y", 0)
                    .imgCloseTop.Picture = LoadResPicture("C_TOP_Y", 0)
                    .imgClose.Picture = LoadResPicture("IMG_CLOSE", 0)
                    '.Line6.BorderColor = &HC0C0&
                    .Line6.BorderColor = &H51C2CC
            Case 2: .Image1.Picture = LoadResPicture("LOGO", 0)
                    .Image2.Picture = LoadResPicture("BANNER_Y", 0)
                    .imgTop = LoadResPicture("T_BANNER_Y", 0)
                    .Image3.Picture = LoadResPicture("L_BORDER_Y", 0)
                    .Image4.Picture = LoadResPicture("R_BORDER_Y", 0)
                    .Image5.Picture = LoadResPicture("B_BORDER_Y", 0)
                    .imgCloseBottom.Picture = LoadResPicture("C_BOTTOM_Y", 0)
                    .imgCloseTop.Picture = LoadResPicture("C_TOP_Y", 0)
                    .imgClose.Picture = LoadResPicture("IMG_CLOSE", 0)
                    .Line6.BorderColor = &H51C2CC
            Case 3: .picShape.Picture = LoadResPicture("L_ABOUT_Y", 0)
            Case 4: .lblProgress1.BackColor = &HC0FFFF
                    .lblProgress2.BackColor = &HC0C0&
        End Select
    End If
End With
End Sub

'Control close button(Image) behavior
Public Sub Close_Button_Action(F_name As Form, Opt As Integer)
With F_name
    Select Case Opt
        Case 1
            .imgCloseBottom.Top = .imgCloseBottom.Top + 20
            .imgCloseBottom.Left = .imgCloseBottom.Left + 20
        Case 2
            .imgCloseBottom.Top = .imgCloseBottom.Top - 20
            .imgCloseBottom.Left = .imgCloseBottom.Left - 20
            Unload F_name
    End Select
End With
End Sub

Public Sub Mouse_Move_Control(frm As Form)
On Error Resume Next
frm.imgCloseBottom.Visible = False
frm.imgCloseTop.Visible = True
End Sub

