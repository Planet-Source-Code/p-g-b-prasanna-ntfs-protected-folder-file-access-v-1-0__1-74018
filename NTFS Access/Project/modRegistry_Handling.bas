Attribute VB_Name = "modRegistry_Handling"
'=========================================
'=  Developed by P. G. B. Prasanna       =
'=  A Software Developer from Sri Lanka  =
'=  E-Mail: pgbsoft@gmail.com            =
'=========================================

'Read Registry Value
Public Function Read_Reg_Data(ByVal R_Loc As String) As Variant
On Error Resume Next
Read_Reg_Data = 0: Read_Reg_Data = reg_obj.RegRead(R_Loc)
End Function

'Set Registry Value
Public Sub Set_Reg_Data(ByVal R_Loc As String, ByVal VData As Variant, Optional RType As Integer = 0)
On Error Resume Next
Select Case RType
    Case 0: reg_obj.RegWrite (R_Loc), VData, "REG_DWORD"
    Case Else: reg_obj.RegWrite (R_Loc), VData
End Select
End Sub

'Delete Registry Value
Public Sub Remove_Reg_Data(ByVal R_Loc As String)
On Error Resume Next
reg_obj.RegDelete (R_Loc)
End Sub

