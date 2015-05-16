Attribute VB_Name = "提权"
'提升到DEBUG权限的声明

Private Const ANYSIZE_ARRAY = 1
Private Const TOKEN_ADJUST_PRIVILEGES = &H20
Private Const TOKEN_QUERY = &H8
Private Const SE_PRIVILEGE_ENABLED = &H2
Private Type LUID
lowpart As Long
highpart As Long
End Type
Private Type LUID_AND_ATTRIBUTES
pLuid As LUID
Attributes As Long
End Type
Private Type TOKEN_PRIVILEGES
PrivilegeCount As Long
Privileges(ANYSIZE_ARRAY) As LUID_AND_ATTRIBUTES
End Type
Private Declare Function SetSystemPowerState Lib "kernel32" (ByVal fSuspend As Long, ByVal fForce As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function OpenProcessToken Lib "advapi32" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long
Private Declare Function AdjustTokenPrivileges Lib "advapi32" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
Const STANDARD_RIGHTS_REQUIRED = &HF0000
Const TOKEN_ASSIGN_PRIMARY = &H1
Const TOKEN_DUPLICATE = (&H2)
Const TOKEN_IMPERSONATE = (&H4)
Const TOKEN_QUERY_SOURCE = (&H10)
Const TOKEN_ADJUST_GROUPS = (&H40)
Const TOKEN_ADJUST_DEFAULT = (&H80)
Const TOKEN_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or TOKEN_ASSIGN_PRIMARY Or _
TOKEN_DUPLICATE Or TOKEN_IMPERSONATE Or TOKEN_QUERY Or TOKEN_QUERY_SOURCE Or _
TOKEN_ADJUST_PRIVILEGES Or TOKEN_ADJUST_GROUPS Or TOKEN_ADJUST_DEFAULT)

 '提升到DEBUG权限的代码
Sub TokenPrivileges()
Dim hdlProcessHandle As Long
Dim hdlTokenHandle As Long
Dim tmpLuid As LUID
Dim tkp As TOKEN_PRIVILEGES
Dim tkpNewButIgnored As TOKEN_PRIVILEGES
Dim lBufferNeeded As Long
Dim lp As Long
hdlProcessHandle = GetCurrentProcess()
lp = OpenProcessToken(hdlProcessHandle, TOKEN_ALL_ACCESS, hdlTokenHandle)
lp = LookupPrivilegeValue("", "SeDebugPrivilege", tmpLuid)
tkp.PrivilegeCount = 1
tkp.Privileges(0).pLuid = tmpLuid
tkp.Privileges(0).Attributes = SE_PRIVILEGE_ENABLED
lp = AdjustTokenPrivileges(hdlTokenHandle, False, tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded)
End Sub




