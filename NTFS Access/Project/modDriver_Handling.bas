Attribute VB_Name = "modDriver_Handling"
'=========================================
'=  Developed by P. G. B. Prasanna       =
'=  A Software Developer from Sri Lanka  =
'=  E-Mail: pgbsoft@gmail.com            =
'=========================================
Option Explicit

'Installing the Driver configuration with inf
Public Sub Config_Setup_Inf()
On Error GoTo Err
If Read_Reg_Data(R_LocDrv) = 0 Then
    If PathFileExists(DesInf) = 0 Then
        InfStr(0) = "[Version]": InfStr(1) = "Signature=""$Windows NT$"""
        InfStr(2) = "Class=""ActivityMonitor""": InfStr(3) = "ClassGuid={b86dff51-a31e-4bac-b3cf-e8cfe75c9fc2}"
        InfStr(4) = "Provider=%Twindows%": InfStr(5) = "[DefaultInstall.Services]"
        InfStr(6) = "AddService=%ServiceName%,,NtfsAD.Service": InfStr(7) = "[NtfsAD.Service]"
        InfStr(8) = "DisplayName=""Ntfs Access Driver""": InfStr(9) = "ServiceBinary=%12%\%DriverName%.sys"
        InfStr(10) = "Dependencies=""FltMgr""": InfStr(11) = "ServiceType=2": InfStr(12) = "StartType=3"
        InfStr(13) = "ErrorControl=1": InfStr(14) = "LoadOrderGroup=""FSFilter Activity Monitor"""
        InfStr(15) = "AddReg=NtfsAD.AddRegistry": InfStr(16) = "[NtfsAD.AddRegistry]"
        InfStr(17) = "HKR,""Instances"",""DefaultInstance"",0x00000000,%DefaultInstance%"
        InfStr(18) = "HKR,""Instances\""%Instance1.Name%,""Altitude"",0x00000000,%Instance1.Altitude%"
        InfStr(19) = "HKR,""Instances\""%Instance1.Name%,""Flags"",0x00010001,%Instance1.Flags%"
        InfStr(20) = "[Strings]": InfStr(21) = "ServiceName=""NtfsAD""": InfStr(22) = "DriverName=""ntfsaccess"""
        InfStr(23) = "DefaultInstance=""NtfsAccess Instance""": InfStr(24) = "Instance1.Name=""NtfsAccess Instance"""
        InfStr(25) = "Instance1.Altitude=""370020""": InfStr(26) = "Instance1.Flags=0x1"
        
        'Creating the inf
        Dim i As Integer
        Open DesInf For Output As #1: For i = LBound(InfStr) To UBound(InfStr): Print #1, InfStr(i): Next: Close #1
        
    End If
    'Installing the inf
    If PathFileExists(DesInf) = 1 Then Shell "RUNDLL32 " & _
                                   "SETUPAPI.DLL,InstallHinfSection DefaultInstall 132 " & _
                                   DesInf, vbHide
End If

Exit Sub
Err:
If intActionMode = 1 Then
    CenterMessage frmMain, "Procedure: Config_Setup_Inf" & vbCrLf & Err.Description & vbCrLf & Err.Number, vbCritical, AppTitles(3)
Else
    MsgBox "Procedure: Config_Setup_Inf" & vbCrLf & Err.Description & vbCrLf & Err.Number, vbCritical, AppTitles(3): End
End If
End Sub

'Configuring the driver
Public Sub Config_Drv(ByVal Opt As Integer)
On Error GoTo Err

Select Case Opt
    Case 1: 'If Is_OS_64 = True Then
            '   ExtractDRV "x64": Config_Setup_Inf
            'Else
               ExtractDRV "x86": Config_Setup_Inf
               'SetAttr DesSys, vbHidden + vbSystem: SetAttr DesInf, vbHidden + vbSystem
            'End If
    Case 2: If PathFileExists(DesInf) = 1 Then SetAttr DesInf, vbNormal: Kill DesInf: If PathFileExists(DesSys) = 1 Then SetAttr DesSys, vbNormal: Kill DesSys
End Select

Exit Sub
Err:
If intActionMode = 1 Then
    CenterMessage frmMain, "Procedure: Config_Drv" & vbCrLf & Err.Description & vbCrLf & Err.Number, vbCritical, AppTitles(3)
Else
    MsgBox "Procedure: Config_Drv" & vbCrLf & Err.Description & vbCrLf & Err.Number, vbCritical, AppTitles(3): End
End If
End Sub

'Extract the required system file
Public Sub ExtractDRV(DataIndex As String)
On Error GoTo Err
     Dim b_data() As Byte
     Dim file_index As Long
     If PathFileExists(DesSys) = 0 Then
        b_data = LoadResData(DataIndex, "CUSTOM")
        file_index = FreeFile
        Open DesSys For Binary Access Write As #file_index
        Put #file_index, , b_data
        Close #file_index
     End If
     
Exit Sub
Err:
If intActionMode = 1 Then
    CenterMessage frmMain, "Procedure: ExtractDRV" & vbCrLf & Err.Description & vbCrLf & Err.Number, vbCritical, AppTitles(3)
Else
    MsgBox "Procedure: ExtractDRV" & vbCrLf & Err.Description & vbCrLf & Err.Number, vbCritical, AppTitles(3): End
End If
End Sub

'Control the filter driver
Public Sub Control_MiniFilter_Driver(Opt As Long, Optional ADDrv As String)
On Error GoTo Err
Select Case Opt
    Case 1: Shell "fltmc " & " Attach NtfsAd " & ADDrv, vbHide
    Case 2: Shell "fltmc " & " Detach NtfsAd " & ADDrv, vbHide
    Case 3: Shell "fltmc " & " Load NtfsAd ", vbHide
    Case 4: Shell "fltmc " & " Unload NtfsAd ", vbHide
    Case 5: Shell "sc " & " Stop NtfsAd", vbHide
    Case 6: Shell "sc " & " Delete NtfsAd", vbHide
End Select

Exit Sub
Err:
If intActionMode = 1 Then
    CenterMessage frmMain, "Procedure: Control_MiniFilter_Driver" & vbCrLf & Err.Description & vbCrLf & Err.Number, vbCritical, AppTitles(3)
Else
    MsgBox "Procedure: Control_MiniFilter_Driver" & vbCrLf & Err.Description & vbCrLf & Err.Number, vbCritical, AppTitles(3): End
End If
End Sub

