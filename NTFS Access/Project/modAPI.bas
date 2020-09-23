Attribute VB_Name = "modAPI"
'=========================================
'=  Developed by P. G. B. Prasanna       =
'=  A Software Developer from Sri Lanka  =
'=  E-Mail: pgbsoft@gmail.com            =
'=========================================

'API Functions used in the program -----------------------------------------------------------------------------

Public Declare Function SetForegroundWindow Lib "user32" _
    (ByVal hwnd As Long) As Long
    
Public Declare Function ShellExecute Lib "Shell32.dll" _
    Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    
Public Declare Function SetWindowPos Lib "user32" _
    (ByVal hwnd As Long, ByVal hWndInsertAfter As WndInsertAfterEnum, _
    ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As SetWindowPosFlagsEnum) As Long
    
Public Declare Function InitCommonControlsEx Lib "comctl32.dll" _
    (iccex As tagInitCommonControlsEx) As Boolean
    
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
(ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" _
(ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Public Declare Function IsUserAnAdmin Lib "shell32" () As Long

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, _
    ByVal X2 As Long, ByVal Y2 As Long) As Long
    
Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, _
    ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
    
Public Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, _
    ByVal X2 As Long, ByVal Y2 As Long) As Long
    
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, _
    ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
    
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, _
ByVal bRedraw As Boolean) As Long

Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Public Declare Function ReleaseCapture Lib "user32" () As Long

Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crey As Byte, _
    ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
    
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, _
    ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, _
    ByVal nIndex As Long) As Long
    
Public Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, _
    lpObject As Any) As Long
    
Public Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, _
    ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
    
Public Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal _
    lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, _
    lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal _
    lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
    
Public Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" _
    (ByVal pszPath As String) As Long
    
Public Declare Function PathIsDirectory Lib "shlwapi.dll" Alias "PathIsDirectoryA" _
    (ByVal pszPath As String) As Long
    
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Declare Function SHChangeNotify Lib "Shell32.dll" (ByVal wEventID As Long, _
    ByVal uFlags As Long, ByVal dwItem1 As Long, ByVal dwItem2 As Long) As Long

Public Declare Function GetProcAddress Lib "kernel32" _
    (ByVal hModule As Long, ByVal lpProcName As String) As Long

Public Declare Function GetModuleHandle Lib "kernel32" _
    Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long

Public Declare Function GetCurrentProcess Lib "kernel32" () As Long

Public Declare Function IsWow64Process Lib "kernel32" _
    (ByVal hProc As Long, ByRef bWow64Process As Boolean) As Long

Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long

Public Declare Function GetCurrentThreadId Lib "kernel32" () As Long

Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" _
    (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
    
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Public Declare Function QueryDosDeviceW Lib "kernel32.dll" (ByVal lpDeviceName As Long, _
    ByVal lpTargetPath As Long, ByVal ucchMax As Long) As Long

Public Declare Function GetVersionExA Lib "kernel32" (lpVersionInformation As OSVERSIONINFO) As Integer

'Types, Enums, Constants -----------------------------------------------------------------------------------------

Public Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type

Public Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Public Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Public Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type

Public Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type

Public Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Public Type OSVERSIONINFO
dwOSVersionInfoSize As Long
dwMajorVersion As Long
dwMinorVersion As Long
dwBuildNumber As Long
dwPlatformId As Long
szCSDVersion As String * 128
End Type

Public Enum WndInsertAfterEnum
    HWND_BOTTOM = 1
    HWND_BROADCAST = &HFFFF&
    HWND_DESKTOP = 0
    HWND_NOTOPMOST = -2
    HWND_TOP = 0
    HWND_TOPMOST = -1
End Enum

Public Enum SetWindowPosFlagsEnum
   SWP_FRAMECHANGED = &H20
   SWP_DRAWFRAME = SWP_FRAMECHANGED
   SWP_HIDEWINDOW = &H80
   SWP_NOACTIVATE = &H10
   SWP_NOCOPYBITS = &H100
   SWP_NOMOVE = &H2
   SWP_NOOWNERZORDER = &H200
   SWP_NOREDRAW = &H8
   SWP_NOREPOSITION = SWP_NOOWNERZORDER
   SWP_NOSIZE = &H1
   SWP_NOZORDER = &H4
   SWP_SHOWWINDOW = &H40
End Enum

Public Const SW_SHOW = 5

Public Const ICC_USEREX_CLASSES = &H200

Public Const SHCNE_ASSOCCHANGED = &H8000000
Public Const SHCNF_FLUSH = &H1000

Public Const RGN_OR = 2
Public Const RGN_DIFF = 4
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2

Public Const GWL_EXSTYLE = (-20)
Public Const WS_EX_LAYERED = &H80000
Public Const LWA_ALPHA = &H2&

Public Const GWL_HINSTANCE = (-6)
'Public Const SWP_NOSIZE = &H1
'Public Const SWP_NOZORDER = &H4
'Public Const SWP_NOACTIVATE = &H10
Public Const HCBT_ACTIVATE = 5
Public Const WH_CBT = 5

'Custom defined Variables, Constants -------------------------------------------------------------------------------

Public xp As Long, yp As Long
Public mShape As Integer
Public mChildFormRegion As Long

Public hHook As Long
Public FrmhWnd As Long

'-----------------------------------------------------------------------

Public InfStr(27), DesInf, DesSys As String
Public strDrvL As String
Public fs_obj, reg_obj As Object
Public bTrans_P_Level, bTrans_P_Level_Limit As Byte
Public intService_Uninstall As Integer
Public intAdmin As Integer
Public intIsFirstFlag As Integer
Public intActionMode As Integer
Public AppTitles(3) As String
Public intSkin As Integer

Public Const R_LocApp = "HKEY_LOCAL_MACHINE\SOFTWARE\NTFSFF-Access\Settings\Opt_Val"
Public Const R_IsFirst = "HKEY_LOCAL_MACHINE\SOFTWARE\NTFSFF-Access\Settings\IsFirst"
Public Const R_Tpncy = "HKEY_LOCAL_MACHINE\SOFTWARE\NTFSFF-Access\Settings\Transparency"
Public Const R_Skin = "HKEY_LOCAL_MACHINE\SOFTWARE\NTFSFF-Access\Settings\Skin"
Public Const R_LocDrv = "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\NtfsAD\DisplayName"
Public Const R_Startup = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\NTFSAccess"

