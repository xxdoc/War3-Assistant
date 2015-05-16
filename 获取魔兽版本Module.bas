Attribute VB_Name = "获取魔兽版本Module"
Public Function 获取魔兽版本() As String

Dim y As Long
Dim hwnd As Long, Handle As Long, PID As Long, Address As Long
hwnd = FindWindow(vbNullString, "Warcraft III")
GetWindowThreadProcessId hwnd, PID
Handle = OpenProcess(PROCESS_ALL_ACCESS, False, PID)
Call Getgamedll
ReadProcessMemory Handle, ByVal Wargamedll + &H636F5D, ByVal VarPtr(y), 4, 0&
        If y = 27836032 Then
            获取魔兽版本 = "1.24B"
        ElseIf y = 993808523 Then
            获取魔兽版本 = "1.24E"
        ElseIf y = 1408011093 Then
            获取魔兽版本 = "1.20E"
'        ElseIf y = 74777673 Then
'            获取魔兽版本 = "1.21"
        Else
            获取魔兽版本 = "未知版本，如有需要请联系作者"
Debug.Print y
        End If
CloseHandle Handle
End Function
