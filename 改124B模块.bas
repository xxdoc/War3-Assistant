Attribute VB_Name = "改124B模块"
Sub 改124B()
Dim hWnd As Long, Handle As Long, PID As Long
hWnd = FindWindow(vbNullString, "Warcraft III")
GetWindowThreadProcessId hWnd, PID
Handle = OpenProcess(PROCESS_ALL_ACCESS, False, PID)
Call Getgamedll
    Call 小地图显示隐形124B(Handle)
    Call 小地图显示单位124B(Handle)
    Call 大地图显示隐形124B(Handle)
    Call 大地图显示单位124B(Handle)
    Call 大地图共享敌方视野124B(Handle)
    Call 小地图共享敌方视野124B(Handle)
CloseHandle Handle
End Sub
Sub 不改124B()
Dim hWnd As Long, Handle As Long, PID As Long
hWnd = FindWindow(vbNullString, "Warcraft III")
GetWindowThreadProcessId hWnd, PID
Handle = OpenProcess(PROCESS_ALL_ACCESS, False, PID)
Call Getgamedll
    Call 小地图不显示隐形124B(Handle)
    Call 小地图不显示单位124B(Handle)
    Call 大地图不显示隐形124B(Handle)
    Call 大地图不显示单位124B(Handle)
    Call 大地图不共享敌方视野124B(Handle)
    Call 小地图不共享敌方视野124B(Handle)
CloseHandle Handle
End Sub
Sub 大地图共享敌方视野124B(Handle As Long)
    WriteProcessMemory Handle, ByVal Wargamedll + &H74D0D1, 104, 1, 0&
    WriteProcessMemory Handle, ByVal Wargamedll + &H74D0D2, 255, 1, 0&
    WriteProcessMemory Handle, ByVal Wargamedll + &H74D0D3, 15, 1, 0&
End Sub
Sub 小地图共享敌方视野124B(Handle As Long)
    WriteProcessMemory Handle, ByVal Wargamedll + &H356F39, 33, 1, 0&
    WriteProcessMemory Handle, ByVal Wargamedll + &H356F3A, 192, 1, 0&
End Sub
Sub 大地图不共享敌方视野124B(Handle As Long)
    WriteProcessMemory Handle, ByVal Wargamedll + &H74D0D1, 139, 1, 0&
    WriteProcessMemory Handle, ByVal Wargamedll + &H74D0D2, 84, 1, 0&
    WriteProcessMemory Handle, ByVal Wargamedll + &H74D0D3, 36, 1, 0&
End Sub
Sub 小地图不共享敌方视野124B(Handle As Long)
    WriteProcessMemory Handle, ByVal Wargamedll + &H356F39, 35, 1, 0&
    WriteProcessMemory Handle, ByVal Wargamedll + &H356F3A, 194, 1, 0&
End Sub
Sub 小地图显示隐形124B(Handle As Long) '小地图显示隐形单位1.24b
WriteProcessMemory Handle, ByVal Wargamedll + &H361EBC, &H0, 1, 0&
End Sub
Sub 小地图不显示隐形124B(Handle As Long)
WriteProcessMemory Handle, ByVal Wargamedll + &H361EBC, &H1, 1, 0&
End Sub
Sub 小地图显示单位124B(Handle As Long) '小地图单位显示1.24b
WriteProcessMemory Handle, ByVal Wargamedll + &H361EAB, &H9090, 2, 0&
End Sub
Sub 小地图不显示单位124B(Handle As Long)
WriteProcessMemory Handle, ByVal Wargamedll + &H361EAB, &H750C, 2, 0&
End Sub
Sub 大地图显示隐形124B(Handle As Long) '大地图显示隐形单位1.24b
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
Sub 大地图不显示隐形124B(Handle As Long)
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
Sub 大地图显示单位124B(Handle As Long) '大地图单位显示1.24b
WriteProcessMemory Handle, ByVal Wargamedll + &H3A201D, &HEB, 1, 0&
End Sub
Sub 大地图不显示单位124B(Handle As Long)
WriteProcessMemory Handle, ByVal Wargamedll + &H3A201D, &H75, 1, 0&
End Sub
Sub 显血124B()
Dim hWnd As Long, Handle As Long, PID As Long, Address As Long
hWnd = FindWindow(vbNullString, "Warcraft III")
GetWindowThreadProcessId hWnd, PID
Handle = OpenProcess(PROCESS_ALL_ACCESS, False, PID)
Call Getgamedll
    ReadProcessMemory Handle, ByVal Wargamedll + &HACBDD8, Address, 4, 0&
    WriteProcessMemory Handle, ByVal Address + &H1D8, 1, 2, 0&
CloseHandle Handle
End Sub
Sub 不显血124B()
Dim hWnd As Long, Handle As Long, PID As Long, Address As Long
hWnd = FindWindow(vbNullString, "Warcraft III")
GetWindowThreadProcessId hWnd, PID
Handle = OpenProcess(PROCESS_ALL_ACCESS, False, PID)
Call Getgamedll
    ReadProcessMemory Handle, ByVal Wargamedll + &HACBDD8, Address, 4, 0&
    WriteProcessMemory Handle, ByVal Address + &H1D8, 0, 2, 0&
CloseHandle Handle
End Sub
Function 获取显血状态124B() As Long
Dim hWnd As Long, Handle As Long, PID As Long, Address As Long
hWnd = FindWindow(vbNullString, "Warcraft III")
GetWindowThreadProcessId hWnd, PID
Handle = OpenProcess(PROCESS_ALL_ACCESS, False, PID)
Call Getgamedll
    ReadProcessMemory Handle, ByVal Wargamedll + &HACBDD8, Address, 4, 0&
    ReadProcessMemory Handle, ByVal Address + &H1D8, 获取显血状态124B, 4, 0&
CloseHandle Handle
End Function

Function 获取游戏状态124B() As Long  '0为未开始，大于0为已开始
Dim hWnd As Long, Handle As Long, PID As Long, Address As Long
hWnd = FindWindow(vbNullString, "Warcraft III")
GetWindowThreadProcessId hWnd, PID
Handle = OpenProcess(PROCESS_ALL_ACCESS, False, PID)
Call Getgamedll
ReadProcessMemory Handle, ByVal Wargamedll + &HACAA5C, 获取游戏状态124B, 1, 0&
CloseHandle Handle
End Function
