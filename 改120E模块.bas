Attribute VB_Name = "��120Eģ��"
Sub ��120E()
Dim hwnd As Long, Handle As Long, PID As Long
hwnd = FindWindow(vbNullString, "Warcraft III")
GetWindowThreadProcessId hwnd, PID
Handle = OpenProcess(PROCESS_ALL_ACCESS, False, PID)
Call Getgamedll
    Call С��ͼ��ʾ����120E(Handle)
    Call С��ͼ��ʾ��λ120E(Handle)
    Call ���ͼ��ʾ����120E(Handle)
    Call ���ͼ��ʾ��λ120E(Handle)
CloseHandle Handle
End Sub
Sub ����120E()
Dim hwnd As Long, Handle As Long, PID As Long
hwnd = FindWindow(vbNullString, "Warcraft III")
GetWindowThreadProcessId hwnd, PID
Handle = OpenProcess(PROCESS_ALL_ACCESS, False, PID)
Call Getgamedll
    Call С��ͼ����ʾ����120E(Handle)
    Call С��ͼ����ʾ��λ120E(Handle)
    Call ���ͼ����ʾ����120E(Handle)
    Call ���ͼ����ʾ��λ120E(Handle)
CloseHandle Handle
End Sub
Sub ���ͼ��ʾ��λ120E(Handle As Long)
    WriteProcessMemory Handle, ByVal Wargamedll + &H2A0923, &HC03340, 3, 0&
    WriteProcessMemory Handle, ByVal Wargamedll + &H2A0926, &HD23342, 3, 0&
End Sub
Sub ���ͼ����ʾ��λ120E(Handle As Long)
    WriteProcessMemory Handle, ByVal Wargamedll + &H2A0923, &H14428B, 3, 0&
    WriteProcessMemory Handle, ByVal Wargamedll + &H2A0926, &H10528B, 3, 0&
End Sub
Sub ���ͼ��ʾ����120E(Handle As Long)
    WriteProcessMemory Handle, ByVal Wargamedll + &H17D4E8, &H90909090, 4, 0&
    WriteProcessMemory Handle, ByVal Wargamedll + &H17D4EC, &H1B8, 4, 0&
End Sub
Sub ���ͼ����ʾ����120E(Handle As Long)
    WriteProcessMemory Handle, ByVal Wargamedll + &H17D4E8, &H5650006A, 4, 0&
    WriteProcessMemory Handle, ByVal Wargamedll + &H17D4EC, &H12355FE8, 4, 0&
End Sub
Sub С��ͼ��ʾ��λ120E(Handle As Long)
    WriteProcessMemory Handle, ByVal Wargamedll + &H1491A8, &H0, 1, 0&
End Sub
Sub С��ͼ����ʾ��λ120E(Handle As Long)
    WriteProcessMemory Handle, ByVal Wargamedll + &H1491A8, &H1, 1, 0&
End Sub
Sub С��ͼ��ʾ����120E(Handle As Long)
    WriteProcessMemory Handle, ByVal Wargamedll + &H1494E0, &H39, 1, 0&
    WriteProcessMemory Handle, ByVal Wargamedll + &H1494E3, &H85, 1, 0&
End Sub
Sub С��ͼ����ʾ����120E(Handle As Long)
    WriteProcessMemory Handle, ByVal Wargamedll + &H1494E0, &H85, 1, 0&
    WriteProcessMemory Handle, ByVal Wargamedll + &H1494E3, &H84, 1, 0&
End Sub
Sub ��ʾ����Ѫ��120E()
Dim hwnd As Long, Handle As Long, PID As Long
hwnd = FindWindow(vbNullString, "Warcraft III")
GetWindowThreadProcessId hwnd, PID
Handle = OpenProcess(PROCESS_ALL_ACCESS, False, PID)
Call Getgamedll
    WriteProcessMemory Handle, ByVal Wargamedll + &H17F133, &HEB, 1, 0&
CloseHandle Handle
End Sub

Sub ���ؼ���Ѫ��120E()
Dim hwnd As Long, Handle As Long, PID As Long
hwnd = FindWindow(vbNullString, "Warcraft III")
GetWindowThreadProcessId hwnd, PID
Handle = OpenProcess(PROCESS_ALL_ACCESS, False, PID)
Call Getgamedll
    WriteProcessMemory Handle, ByVal Wargamedll + &H17F133, &H75, 1, 0&
 CloseHandle Handle
End Sub

Sub ��ʾ�з�Ѫ��120E()
Dim hwnd As Long, Handle As Long, PID As Long
hwnd = FindWindow(vbNullString, "Warcraft III")
GetWindowThreadProcessId hwnd, PID
Handle = OpenProcess(PROCESS_ALL_ACCESS, False, PID)
Call Getgamedll
    WriteProcessMemory Handle, ByVal Wargamedll + &H17F150, &HEB, 1, 0&
CloseHandle Handle
End Sub

Sub ���صз�Ѫ��120E()
Dim hwnd As Long, Handle As Long, PID As Long
hwnd = FindWindow(vbNullString, "Warcraft III")
GetWindowThreadProcessId hwnd, PID
Handle = OpenProcess(PROCESS_ALL_ACCESS, False, PID)
Call Getgamedll
    WriteProcessMemory Handle, ByVal Wargamedll + &H17F150, &H75, 1, 0&
CloseHandle Handle
End Sub
Sub ��Ѫ120E()
    Call ��ʾ�з�Ѫ��120E
    Call ��ʾ����Ѫ��120E
End Sub
Sub ����Ѫ120E()
    Call ���صз�Ѫ��120E
    Call ���ؼ���Ѫ��120E
End Sub
Function ��ȡ��Ѫ״̬120E() As Long
Dim hwnd As Long, Handle As Long, PID As Long
hwnd = FindWindow(vbNullString, "Warcraft III")
GetWindowThreadProcessId hwnd, PID
Handle = OpenProcess(PROCESS_ALL_ACCESS, False, PID)
Call Getgamedll
ReadProcessMemory Handle, ByVal Wargamedll + &H17F150, ��ȡ��Ѫ״̬120E, 1, 0&
If ��ȡ��Ѫ״̬120E = &H75 Then ��ȡ��Ѫ״̬120E = 0
If ��ȡ��Ѫ״̬120E = &HEB Then ��ȡ��Ѫ״̬120E = 1
End Function
Function ��ȡ��Ϸ״̬120E() As Long  '0Ϊδ��ʼ��1Ϊ�ѿ�ʼ
Dim hwnd As Long, Handle As Long, PID As Long, Address As Long
hwnd = FindWindow(vbNullString, "Warcraft III")
GetWindowThreadProcessId hwnd, PID
Handle = OpenProcess(PROCESS_ALL_ACCESS, False, PID)
Call Getgamedll
ReadProcessMemory Handle, ByVal Wargamedll + &H870A48, ��ȡ��Ϸ״̬120E, 1, 0&
CloseHandle Handle
End Function
