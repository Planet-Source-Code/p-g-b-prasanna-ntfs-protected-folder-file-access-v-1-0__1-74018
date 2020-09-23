Attribute VB_Name = "modMain"
'=========================================
'=  Developed by P. G. B. Prasanna       =
'=  A Software Developer from Sri Lanka  =
'=  E-Mail: pgbsoft@gmail.com            =
'=========================================

Option Explicit
Public Sub Main()
On Error GoTo Err

AppTitles(0) = "NTFS Protected Folder/File Access v 1.0"
AppTitles(1) = "Incompatible Version"
AppTitles(2) = "Administrative privileges required"
AppTitles(3) = "Error Detected"

'Checking whether the application is already running.
If IS_APP_RUNNING = True Then End

'If the OS is 64-bit, we stop here because the filter driver used with this program is 32-bit.
If Is_OS_64 = True Then MsgBox "This program is designed to run only on 32-bit versions of Windows." & vbCrLf & _
                               App.Title & " cannot continue...", vbCritical, AppTitles(1): End
                               
'Checking for required Windows version.
If Check_Windows_Version = False Then MsgBox "This version of Windows is not supported." & vbCrLf & _
                               App.Title & " cannot continue...", vbCritical, AppTitles(1): End

'Checking whether the program is running with administrative privileges.
If IS_USER_AN_ADMIN = False Then
    
    If UCase(Command$) <> UCase("/chkadmin") Then
        MsgBox "You do not possess required administrative privileges." & vbCrLf & _
        "(Right click the Exe and use ""Run as administrator"")" & vbCrLf & vbCrLf & _
        "Tip:" & vbCrLf & _
        "-----" & vbCrLf & _
        "If you get this message even in administrative user accounts" & vbCrLf & _
        "on Windows Vista/Windows7, you may do the followings." & vbCrLf & vbCrLf & _
        "• Right click the exe and try with ""Run as administrator"" or" & vbCrLf & _
        "• Disable UAC from Control Panel and restart the machine or" & vbCrLf & _
        "• Login to the system as built-in Administrator", vbCritical, AppTitles(2)
    'Else
    '   End
    End If
   
End If

Dim iccex As tagInitCommonControlsEx
   With iccex
       .lngSize = LenB(iccex)
       .lngICC = ICC_USEREX_CLASSES
   End With
InitCommonControlsEx iccex

'Creating required objects
Set reg_obj = CreateObject("Wscript.Shell"): Set fs_obj = CreateObject("Scripting.FileSystemObject")

intService_Uninstall = Read_Reg_Data(R_LocApp)
intActionMode = 0
 If intAdmin = 1 Then
    DesInf = Environ$("SystemRoot") & "\System32\ntfsaccess.inf"
    DesSys = Environ$("SystemRoot") & "\System32\drivers\ntfsaccess.sys"
 
    If intService_Uninstall = 1 Then Call Config_Drv(1): Call Control_MiniFilter_Driver(3)
    If Read_Reg_Data(R_IsFirst) = 0 Then intIsFirstFlag = 1: Set_Reg_Data (R_IsFirst), 1: Set_Reg_Data (R_LocApp), 0: intService_Uninstall = 0: Call Config_Drv(1)
    
    Call Control_MiniFilter_Driver(3)
 End If
frmMain.Show

Exit Sub
Err:
MsgBox "Procedure: Main" & vbCrLf & Err.Description & vbCrLf & Err.Number, vbCritical, AppTitles(3): End
End Sub

Public Function Get_Full_App_Path() As String
If Right(App.Path, 1) = "\" Then
    Get_Full_App_Path = App.Path & App.EXEName & ".exe"
Else
    Get_Full_App_Path = App.Path & "\" & App.EXEName & ".exe"
End If
End Function

Public Sub Mail_To(Optional M_opt As Integer = 0)
Select Case M_opt
    Case 0
            ShellExecute 0, "Open", "mailto:pgbsoft@gmail.com?Subject=" & _
            AppTitles(0) & " - Feedback", vbNullString, vbNullString, SW_SHOW
    Case 1
            ShellExecute 0, "Open", "mailto:pgbsoft@gmail.com?Subject=" & _
            AppTitles(0) & " - Uninstall&Body=Please let me know your idea about " & _
            AppTitles(0) & ".", vbNullString, vbNullString, SW_SHOW
End Select
End Sub
