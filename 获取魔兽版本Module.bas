Attribute VB_Name = "��ȡħ�ް汾Module"
Public Function ��ȡħ�ް汾() As String

Dim y As Long
Dim hwnd As Long, Handle As Long, PID As Long, Address As Long
hwnd = FindWindow(vbNullString, "Warcraft III")
GetWindowThreadProcessId hwnd, PID
Handle = OpenProcess(PROCESS_ALL_ACCESS, False, PID)
Call Getgamedll
ReadProcessMemory Handle, ByVal Wargamedll + &H636F5D, ByVal VarPtr(y), 4, 0&
        If y = 27836032 Then
            ��ȡħ�ް汾 = "1.24B"
        ElseIf y = 993808523 Then
            ��ȡħ�ް汾 = "1.24E"
        ElseIf y = 1408011093 Then
            ��ȡħ�ް汾 = "1.20E"
'        ElseIf y = 74777673 Then
'            ��ȡħ�ް汾 = "1.21"
        Else
            ��ȡħ�ް汾 = "δ֪�汾��������Ҫ����ϵ����"
Debug.Print y
        End If
CloseHandle Handle
End Function
