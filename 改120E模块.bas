Attribute VB_Name = "改120E模块"
Sub 改120E()
Dim hwnd As Long, Handle As Long, PID As Long
hwnd = FindWindow(vbNullString, "Warcraft III")
GetWindowThreadProcessId hwnd, PID
Handle = OpenProcess(PROCESS_ALL_ACCESS, False, PID)
Call Getgamedll
    Call 小地图显示隐形120E(Handle)
    Call 小地图显示单位120E(Handle)
    Call 大地图显示隐形120E(Handle)
    Call 大地图显示单位120E(Handle)
CloseHandle Handle
End Sub
Sub 不改120E()
Dim hwnd As Long, Handle As Long, PID As Long
hwnd = FindWindow(vbNullString, "Warcraft III")
GetWindowThreadProcessId hwnd, PID
Handle = OpenProcess(PROCESS_ALL_ACCESS, False, PID)
Call Getgamedll
    Call 小地图不显示隐形120E(Handle)
    Call 小地图不显示单位120E(Handle)
    Call 大地图不显示隐形120E(Handle)
    Call 大地图不显示单位120E(Handle)
CloseHandle Handle
End Sub
Sub 大地图显示单位120E(Handle As Long)
    WriteProcessMemory Handle, ByVal Wargamedll + &H2A0923, &HC03340, 3, 0&
    WriteProcessMemory Handle, ByVal Wargamedll + &H2A0926, &HD23342, 3, 0&
End Sub
Sub 大地图不显示单位120E(Handle As Long)
    WriteProcessMemory Handle, ByVal Wargamedll + &H2A0923, &H14428B, 3, 0&
    WriteProcessMemory Handle, ByVal Wargamedll + &H2A0926, &H10528B, 3, 0&
End Sub
Sub 大地图显示隐形120E(Handle As Long)
    WriteProcessMemory Handle, ByVal Wargamedll + &H17D4E8, &H90909090, 4, 0&
    WriteProcessMemory Handle, ByVal Wargamedll + &H17D4EC, &H1B8, 4, 0&
End Sub
Sub 大地图不显示隐形120E(Handle As Long)
    WriteProcessMemory Handle, ByVal Wargamedll + &H17D4E8, &H5650006A, 4, 0&
    WriteProcessMemory Handle, ByVal Wargamedll + &H17D4EC, &H12355FE8, 4, 0&
End Sub
Sub 小地图显示单位120E(Handle As Long)
    WriteProcessMemory Handle, ByVal Wargamedll + &H1491A8, &H0, 1, 0&
End Sub
Sub 小地图不显示单位120E(Handle As Long)
    WriteProcessMemory Handle, ByVal Wargamedll + &H1491A8, &H1, 1, 0&
End Sub
Sub 小地图显示隐形120E(Handle As Long)
    WriteProcessMemory Handle, ByVal Wargamedll + &H1494E0, &H39, 1, 0&
    WriteProcessMemory Handle, ByVal Wargamedll + &H1494E3, &H85, 1, 0&
End Sub
Sub 小地图不显示隐形120E(Handle As Long)
    WriteProcessMemory Handle, ByVal Wargamedll + &H1494E0, &H85, 1, 0&
    WriteProcessMemory Handle, ByVal Wargamedll + &H1494E3, &H84, 1, 0&
End Sub
Sub 显示己方血条120E()
Dim hwnd As Long, Handle As Long, PID As Long
hwnd = FindWindow(vbNullString, "Warcraft III")
GetWindowThreadProcessId hwnd, PID
Handle = OpenProcess(PROCESS_ALL_ACCESS, False, PID)
Call Getgamedll
    WriteProcessMemory Handle, ByVal Wargamedll + &H17F133, &HEB, 1, 0&
CloseHandle Handle
End Sub

Sub 隐藏己方血条120E()
Dim hwnd As Long, Handle As Long, PID As Long
hwnd = FindWindow(vbNullString, "Warcraft III")
GetWindowThreadProcessId hwnd, PID
Handle = OpenProcess(PROCESS_ALL_ACCESS, False, PID)
Call Getgamedll
    WriteProcessMemory Handle, ByVal Wargamedll + &H17F133, &H75, 1, 0&
 CloseHandle Handle
End Sub

Sub 显示敌方血条120E()
Dim hwnd As Long, Handle As Long, PID As Long
hwnd = FindWindow(vbNullString, "Warcraft III")
GetWindowThreadProcessId hwnd, PID
Handle = OpenProcess(PROCESS_ALL_ACCESS, False, PID)
Call Getgamedll
    WriteProcessMemory Handle, ByVal Wargamedll + &H17F150, &HEB, 1, 0&
CloseHandle Handle
End Sub

Sub 隐藏敌方血条120E()
Dim hwnd As Long, Handle As Long, PID As Long
hwnd = FindWindow(vbNullString, "Warcraft III")
GetWindowThreadProcessId hwnd, PID
Handle = OpenProcess(PROCESS_ALL_ACCESS, False, PID)
Call Getgamedll
    WriteProcessMemory Handle, ByVal Wargamedll + &H17F150, &H75, 1, 0&
CloseHandle Handle
End Sub
Sub 显血120E()
    Call 显示敌方血条120E
    Call 显示己方血条120E
End Sub
Sub 不显血120E()
    Call 隐藏敌方血条120E
    Call 隐藏己方血条120E
End Sub
Function 获取显血状态120E() As Long
Dim hwnd As Long, Handle As Long, PID As Long
hwnd = FindWindow(vbNullString, "Warcraft III")
GetWindowThreadProcessId hwnd, PID
Handle = OpenProcess(PROCESS_ALL_ACCESS, False, PID)
Call Getgamedll
ReadProcessMemory Handle, ByVal Wargamedll + &H17F150, 获取显血状态120E, 1, 0&
If 获取显血状态120E = &H75 Then 获取显血状态120E = 0
If 获取显血状态120E = &HEB Then 获取显血状态120E = 1
End Function
Function 获取游戏状态120E() As Long  '0为未开始，1为已开始
Dim hwnd As Long, Handle As Long, PID As Long, Address As Long
hwnd = FindWindow(vbNullString, "Warcraft III")
GetWindowThreadProcessId hwnd, PID
Handle = OpenProcess(PROCESS_ALL_ACCESS, False, PID)
Call Getgamedll
ReadProcessMemory Handle, ByVal Wargamedll + &H870A48, 获取游戏状态120E, 1, 0&
CloseHandle Handle
End Function
