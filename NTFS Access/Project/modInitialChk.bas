Attribute VB_Name = "modInitialchk"
'=========================================
'=  Developed by P. G. B. Prasanna       =
'=  A Software Developer from Sri Lanka  =
'=  E-Mail: pgbsoft@gmail.com            =
'=========================================

'Function to determine the app is running
Public Function IS_APP_RUNNING() As Boolean
Dim cT As Integer
Dim hwndAP As Long
IS_APP_RUNNING = False
    For cT = LBound(AppTitles) To UBound(AppTitles)
        hwndAP = FindWindow(vbNullString, AppTitles(cT)): If hwndAP <> 0 Then IS_APP_RUNNING = True: Exit For
    Next
End Function
'Function to determine the User Type
Public Function IS_USER_AN_ADMIN() As Boolean
IS_USER_AN_ADMIN = True

intAdmin = IsUserAnAdmin()
'intAdmin = 0 'for checking purpose
If intAdmin = 0 Then IS_USER_AN_ADMIN = False
End Function

'This function is used to determine the mode of Windows (Windows x64 or Windows x86)
Public Function Is_OS_64() As Boolean
On Error GoTo Err

   Dim lngHandle As Long
   Dim is64Bit As Boolean

    ' Assume initially that this is not a WOW64 process
    is64Bit = False

    ' Then try to prove that wrong by attempting to load the
    ' IsWow64Process function dynamically
    lngHandle = GetProcAddress(GetModuleHandle("kernel32"), "IsWow64Process")

    ' The function exists, so call it
    If lngHandle <> 0 Then
        IsWow64Process GetCurrentProcess(), is64Bit
    End If

    ' Return the value
    Is_OS_64 = is64Bit
  
   'If is64Bit returns 'False' then the OS is 32-bit, otherwise 64-bit.
    If Not Is_OS_64 = False Then Is_OS_64 = True

Exit Function
Err:
If intActionMode = 1 Then
    CenterMessage frmMain, "Function: Is_OS_64" & vbCrLf & Err.Description & vbCrLf & Err.Number, vbCritical, AppTitles(3)
Else
    MsgBox "Function: Is_OS_64" & vbCrLf & Err.Description & vbCrLf & Err.Number, vbCritical, AppTitles(3): End
End If
End Function

'--------------------------------------------------------------------------------------------------------
'Public Function Is_OS_64() As Boolean
'Dim ProcessorArchitec As String
'Is_OS_64 = False
'ProcessorArchitec = Environ$("PROCESSOR_ARCHITECTURE")
'If UCase(ProcessorArchitec) = "AMD64" Then Is_OS_64 = True
'End Function
'--------------------------------------------------------------------------------------------------------

'Checking for Windows version
Public Function Check_Windows_Version() As Boolean
On Error GoTo Err
Dim OSinfo As OSVERSIONINFO
Dim RetValue As Integer
Dim ID As String

OSinfo.dwOSVersionInfoSize = 148
OSinfo.szCSDVersion = Space$(128)
RetValue = GetVersionExA(OSinfo)

Check_Windows_Version = True

With OSinfo

    If .dwPlatformId < 2 And .dwMajorVersion < 5 Then Check_Windows_Version = False
End With
'Check_Windows_Version = False
Exit Function

Err:
Check_Windows_Version = True
End Function
