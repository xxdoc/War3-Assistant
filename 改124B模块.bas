Attribute VB_Name = "��124Bģ��"
Sub ��124B()
Dim hWnd As Long, Handle As Long, PID As Long
hWnd = FindWindow(vbNullString, "Warcraft III")
GetWindowThreadProcessId hWnd, PID
Handle = OpenProcess(PROCESS_ALL_ACCESS, False, PID)
Call Getgamedll
    Call С��ͼ��ʾ����124B(Handle)
    Call С��ͼ��ʾ��λ124B(Handle)
    Call ���ͼ��ʾ����124B(Handle)
    Call ���ͼ��ʾ��λ124B(Handle)
    Call ���ͼ����з���Ұ124B(Handle)
    Call С��ͼ����з���Ұ124B(Handle)
CloseHandle Handle
End Sub
Sub ����124B()
Dim hWnd As Long, Handle As Long, PID As Long
hWnd = FindWindow(vbNullString, "Warcraft III")
GetWindowThreadProcessId hWnd, PID
Handle = OpenProcess(PROCESS_ALL_ACCESS, False, PID)
Call Getgamedll
    Call С��ͼ����ʾ����124B(Handle)
    Call С��ͼ����ʾ��λ124B(Handle)
    Call ���ͼ����ʾ����124B(Handle)
    Call ���ͼ����ʾ��λ124B(Handle)
    Call ���ͼ������з���Ұ124B(Handle)
    Call С��ͼ������з���Ұ124B(Handle)
CloseHandle Handle
End Sub
Sub ���ͼ����з���Ұ124B(Handle As Long)
    WriteProcessMemory Handle, ByVal Wargamedll + &H74D0D1, 104, 1, 0&
    WriteProcessMemory Handle, ByVal Wargamedll + &H74D0D2, 255, 1, 0&
    WriteProcessMemory Handle, ByVal Wargamedll + &H74D0D3, 15, 1, 0&
End Sub
Sub С��ͼ����з���Ұ124B(Handle As Long)
    WriteProcessMemory Handle, ByVal Wargamedll + &H356F39, 33, 1, 0&
    WriteProcessMemory Handle, ByVal Wargamedll + &H356F3A, 192, 1, 0&
End Sub
Sub ���ͼ������з���Ұ124B(Handle As Long)
    WriteProcessMemory Handle, ByVal Wargamedll + &H74D0D1, 139, 1, 0&
    WriteProcessMemory Handle, ByVal Wargamedll + &H74D0D2, 84, 1, 0&
    WriteProcessMemory Handle, ByVal Wargamedll + &H74D0D3, 36, 1, 0&
End Sub
Sub С��ͼ������з���Ұ124B(Handle As Long)
    WriteProcessMemory Handle, ByVal Wargamedll + &H356F39, 35, 1, 0&
    WriteProcessMemory Handle, ByVal Wargamedll + &H356F3A, 194, 1, 0&
End Sub
Sub С��ͼ��ʾ����124B(Handle As Long) 'С��ͼ��ʾ���ε�λ1.24b
WriteProcessMemory Handle, ByVal Wargamedll + &H361EBC, &H0, 1, 0&
End Sub
Sub С��ͼ����ʾ����124B(Handle As Long)
WriteProcessMemory Handle, ByVal Wargamedll + &H361EBC, &H1, 1, 0&
End Sub
Sub С��ͼ��ʾ��λ124B(Handle As Long) 'С��ͼ��λ��ʾ1.24b
WriteProcessMemory Handle, ByVal Wargamedll + &H361EAB, &H9090, 2, 0&
End Sub
Sub С��ͼ����ʾ��λ124B(Handle As Long)
WriteProcessMemory Handle, ByVal Wargamedll + &H361EAB, &H750C, 2, 0&
End Sub
Sub ���ͼ��ʾ����124B(Handle As Long) '���ͼ��ʾ���ε�λ1.24b
    Dim ASMA(6) As Byte
    ASMA(0) = &H90
    ASMA(1) = &H90
    ASMA(2) = &H90
    ASMA(3) = &H90
    ASMA(4) = &H90
    ASMA(5) = &H90
    Dim ASMB(11) As Byte
    ASMB(0) = &H90
    ASMB(1) = &H90
    ASMB(2) = &H90
    ASMB(3) = &H90
    ASMB(4) = &H90
    ASMB(5) = &H90
    ASMB(6) = &H90
    ASMB(7) = &H90
    ASMB(8) = &H33
    ASMB(9) = &HC0
    ASMB(10) = &H40
    WriteProcessMemory Handle, ByVal Wargamedll + &H3622D1, &H3B, 1, 0&
    WriteProcessMemory Handle, ByVal Wargamedll + &H3622D4, &H85, 1, 0&
    WriteProcessMemory Handle, ByVal Wargamedll + &H39A45B, ASMA(0), 6, 0&
    WriteProcessMemory Handle, ByVal Wargamedll + &H39A46E, ASMB(0), 11, 0&
End Sub
Sub ���ͼ����ʾ����124B(Handle As Long)
    Dim ASMA(6) As Byte
    ASMA(0) = &H8B
    ASMA(1) = &H97
    ASMA(2) = &H98
    ASMA(3) = &H1
    ASMA(4) = &H0
    ASMA(5) = &H0
    Dim ASMB(11) As Byte
    ASMB(0) = &HF
    ASMB(1) = &HB7
    ASMB(2) = &H0
    ASMB(3) = &H55
    ASMB(4) = &H50
    ASMB(5) = &H56
    ASMB(6) = &HE8
    ASMB(7) = &HF7
    ASMB(8) = &H7B
    ASMB(9) = &H0
    ASMB(10) = &H0
    WriteProcessMemory Handle, ByVal Wargamedll + &H3622D1, &H85, 1, 0&
    WriteProcessMemory Handle, ByVal Wargamedll + &H3622D4, &H84, 1, 0&
    WriteProcessMemory Handle, ByVal Wargamedll + &H39A45B, ASMA(0), 6, 0&
    WriteProcessMemory Handle, ByVal Wargamedll + &H39A46E, ASMB(0), 11, 0&
End Sub
Sub ���ͼ��ʾ��λ124B(Handle As Long) '���ͼ��λ��ʾ1.24b
WriteProcessMemory Handle, ByVal Wargamedll + &H3A201D, &HEB, 1, 0&
End Sub
Sub ���ͼ����ʾ��λ124B(Handle As Long)
WriteProcessMemory Handle, ByVal Wargamedll + &H3A201D, &H75, 1, 0&
End Sub
Sub ��Ѫ124B()
Dim hWnd As Long, Handle As Long, PID As Long, Address As Long
hWnd = FindWindow(vbNullString, "Warcraft III")
GetWindowThreadProcessId hWnd, PID
Handle = OpenProcess(PROCESS_ALL_ACCESS, False, PID)
Call Getgamedll
    ReadProcessMemory Handle, ByVal Wargamedll + &HACBDD8, Address, 4, 0&
    WriteProcessMemory Handle, ByVal Address + &H1D8, 1, 2, 0&
CloseHandle Handle
End Sub
Sub ����Ѫ124B()
Dim hWnd As Long, Handle As Long, PID As Long, Address As Long
hWnd = FindWindow(vbNullString, "Warcraft III")
GetWindowThreadProcessId hWnd, PID
Handle = OpenProcess(PROCESS_ALL_ACCESS, False, PID)
Call Getgamedll
    ReadProcessMemory Handle, ByVal Wargamedll + &HACBDD8, Address, 4, 0&
    WriteProcessMemory Handle, ByVal Address + &H1D8, 0, 2, 0&
CloseHandle Handle
End Sub
Function ��ȡ��Ѫ״̬124B() As Long
Dim hWnd As Long, Handle As Long, PID As Long, Address As Long
hWnd = FindWindow(vbNullString, "Warcraft III")
GetWindowThreadProcessId hWnd, PID
Handle = OpenProcess(PROCESS_ALL_ACCESS, False, PID)
Call Getgamedll
    ReadProcessMemory Handle, ByVal Wargamedll + &HACBDD8, Address, 4, 0&
    ReadProcessMemory Handle, ByVal Address + &H1D8, ��ȡ��Ѫ״̬124B, 4, 0&
CloseHandle Handle
End Function

Function ��ȡ��Ϸ״̬124B() As Long  '0Ϊδ��ʼ������0Ϊ�ѿ�ʼ
Dim hWnd As Long, Handle As Long, PID As Long, Address As Long
hWnd = FindWindow(vbNullString, "Warcraft III")
GetWindowThreadProcessId hWnd, PID
Handle = OpenProcess(PROCESS_ALL_ACCESS, False, PID)
Call Getgamedll
ReadProcessMemory Handle, ByVal Wargamedll + &HACAA5C, ��ȡ��Ϸ״̬124B, 1, 0&
CloseHandle Handle
End Function
