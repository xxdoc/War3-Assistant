Attribute VB_Name = "获取程序路径"
Private Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Private Declare Function GetModuleFileNameEx Lib "psapi.dll" Alias "GetModuleFileNameExA" (ByVal hProcess As Long, ByVal hModule As Long, ByVal lpFilename As String, ByVal nSize As Long) As Long

Public Function GetProcessPath(ByVal PID As Long) As String
'返回程路径。
    On Error GoTo Z
    Dim cbNeeded As Long
    Dim szBuf(1 To 250) As Long
    Dim Ret As Long
    Dim szPathName As String
    Dim nSize As Long
    Dim hProcess As Long
    hProcess = OpenProcess(&H400 Or &H10, 0, PID)
    If hProcess <> 0 Then
        Ret = EnumProcessModules(hProcess, szBuf(1), 250, cbNeeded)
        If Ret <> 0 Then
            szPathName = Space(260)
            nSize = 500
            Ret = GetModuleFileNameEx(hProcess, szBuf(1), szPathName, nSize)
            GetProcessPath = left(szPathName, Ret)
        End If
    End If
    Ret = CloseHandle(hProcess)
    Exit Function
Z:
    GetProcessPath = ""
End Function
