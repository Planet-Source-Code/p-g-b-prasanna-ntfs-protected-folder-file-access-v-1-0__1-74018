Attribute VB_Name = "modCenterMessage"
'=========================================
'=  Developed by P. G. B. Prasanna       =
'=  A Software Developer from Sri Lanka  =
'=  E-Mail: pgbsoft@gmail.com            =
'=========================================
Option Explicit

Public Function CenterMessage(ParentForm As Form, Msg As String, Optional Buttons As VbMsgBoxStyle = vbOKOnly, Optional Title As String = vbNullString) As VbMsgBoxResult
  Dim hInst As Long
  Dim Thread As Long

  FrmhWnd = ParentForm.hwnd
  hInst = GetWindowLong(ParentForm.hwnd, GWL_HINSTANCE)
  Thread = GetCurrentThreadId()
  hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProc1, hInst, Thread)
  CenterMessage = MsgBox(Msg, Buttons, Title)
End Function

Private Function WinProc1(ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Dim rectForm As RECT, rectMsg As RECT
  Dim x As Long, Y As Long
  If lMsg = HCBT_ACTIVATE Then
    GetWindowRect FrmhWnd, rectForm
    GetWindowRect wParam, rectMsg
    x = (rectForm.Left + (rectForm.Right - rectForm.Left) / 2) - ((rectMsg.Right - rectMsg.Left) / 2)
    Y = (rectForm.Top + (rectForm.Bottom - rectForm.Top) / 2) - ((rectMsg.Bottom - rectMsg.Top) / 2)
    SetWindowPos wParam, 0, x, Y, 0, 0, SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE
    UnhookWindowsHookEx hHook
  End If
  WinProc1 = False
End Function

