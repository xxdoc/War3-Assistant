Attribute VB_Name = "��ȡ����״̬"

Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long

Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long

Public Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long

Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long

Public Const PROCESS_ALL_ACCESS = &H1F0FFF

Public Function ChatState() As Long
    Select Case ��ȡħ�ް汾
    Case "1.24E": ChatState = ChatState124E
    Case "1.24B": ChatState = ChatState124B
    Case "1.20E": ChatState = ChatState120E
'    Case "1.21": ChatState = ChatState121
    End Select
End Function
Public Function ChatState124E() As Long                                                   '��ȡ����״̬��0Ϊ�����죬1Ϊ����
Dim hwnd As Long, Handle As Long, PID As Long

hwnd = FindWindow(vbNullString, "Warcraft III")

GetWindowThreadProcessId hwnd, PID

Handle = OpenProcess(PROCESS_ALL_ACCESS, False, PID)

Call Getgamedll

ReadProcessMemory ByVal Handle, ByVal Wargamedll, ByVal VarPtr(ChatState124E), 4, 0&

Wargamedll = Wargamedll + &HAE8450

ReadProcessMemory ByVal Handle, ByVal Wargamedll, ByVal VarPtr(ChatState124E), 4, 0&

CloseHandle Handle

End Function
Public Function ChatState124B() As Long                                                   '��ȡ����״̬��0Ϊ�����죬1Ϊ����
Dim hwnd As Long, Handle As Long, PID As Long

hwnd = FindWindow(vbNullString, "Warcraft III")

GetWindowThreadProcessId hwnd, PID

Handle = OpenProcess(PROCESS_ALL_ACCESS, False, PID)

Call Getgamedll

ReadProcessMemory ByVal Handle, ByVal Wargamedll, ByVal VarPtr(ChatState124B), 4, 0&

Wargamedll = Wargamedll + &HAE8450

ReadProcessMemory ByVal Handle, ByVal Wargamedll, ByVal VarPtr(ChatState124B), 4, 0&

CloseHandle Handle

End Function
Public Function ChatState120E() As Long                                                   '��ȡ����״̬��0Ϊ�����죬1Ϊ����
Dim hwnd As Long, Handle As Long, PID As Long

hwnd = FindWindow(vbNullString, "Warcraft III")

GetWindowThreadProcessId hwnd, PID

Handle = OpenProcess(PROCESS_ALL_ACCESS, False, PID)

Call Getgamedll

ReadProcessMemory ByVal Handle, ByVal &H45CB8C, ByVal VarPtr(ChatState120E), 4, 0&

CloseHandle Handle

End Function
'Public Function ChatState121() As Long                                                   '��ȡ����״̬��0Ϊ�����죬1Ϊ����
'Dim hwnd As Long, Handle As Long, PID As Long

'hwnd = FindWindow(vbNullString, "Warcraft III")

'GetWindowThreadProcessId hwnd, PID

'Handle = OpenProcess(PROCESS_ALL_ACCESS, False, PID)

'Call Getgamedll

'ReadProcessMemory ByVal Handle, ByVal &H45CB8C, ByVal VarPtr(ChatState121), 4, 0&

'CloseHandle Handle

'End Function
