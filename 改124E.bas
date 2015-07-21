Attribute VB_Name = "改124E模块"
Sub 改124E()
Dim hwnd As Long, Handle As Long, PID As Long
hwnd = FindWindow(vbNullString, "Warcraft III")
GetWindowThreadProcessId hwnd, PID
Handle = OpenProcess(PROCESS_ALL_ACCESS, False, PID)
Call Getgamedll
        Call 大地图显示隐形124E(Handle)
        Call 小地图显示单位124E(Handle)
        Call 大地图显示单位124E(Handle)
        Call 小地图显示隐形124E(Handle)
        Call 大地图共享敌方视野124E(Handle)
        Call 小地图共享敌方视野124E(Handle)
        Call 显示技能124E(Handle)
        Call 分辨幻影124E(Handle)
        If Form3.Check2.Value = 1 Then Call 可选视野外单位124E(Handle)
        If Form3.Check3.Value = 1 Then Call 资源124E(Handle)
        If Form3.Check4.Value = 1 Then Call 清除迷雾124E(Handle)
        If Form3.Check5.Value = 1 Then Call 显示敌方信号124E(Handle)
        If Form3.Check6.Value = 1 Then Call 显示盟军头像124E(Handle)
        If Form3.Check7.Value = 1 Then Call 显示敌方头像124E(Handle)
CloseHandle Handle
End Sub
Sub 不改124E()
Dim hwnd As Long, Handle As Long, PID As Long
hwnd = FindWindow(vbNullString, "Warcraft III")
GetWindowThreadProcessId hwnd, PID
Handle = OpenProcess(PROCESS_ALL_ACCESS, False, PID)
Call Getgamedll
        Call 大地图不显示隐形124E(Handle)
        Call 小地图不显示单位124E(Handle)
        Call 大地图不显示单位124E(Handle)
        Call 小地图不显示隐形124E(Handle)
        Call 大地图不共享敌方视野124E(Handle)
        Call 小地图不共享敌方视野124E(Handle)
        Call 不显示技能124E(Handle)
        Call 不分辨幻影124E(Handle)
        If Form3.Check2.Value = 1 Then Call 不可选视野外单位124E(Handle)
        'If Form3.Check3.Value = 1 Then Call 不资源124E(Handle)
        If Form3.Check4.Value = 1 Then Call 不清除迷雾124E(Handle)
        If Form3.Check5.Value = 1 Then Call 不显示敌方信号124E(Handle)
        If Form3.Check6.Value = 1 Then Call 不显示头像124E(Handle)
        If Form3.Check7.Value = 1 Then Call 不显示头像124E(Handle)
CloseHandle Handle
End Sub

Sub 分辨幻影124E(Handle As Long)  '分辨幻影1.24e
    Dim b(2) As Byte
        b(0) = &H40
        b(1) = &HC3
    WriteProcessMemory Handle, ByVal Wargamedll + &H28357C, b(0), 2, 0&
End Sub

Sub 不分辨幻影124E(Handle As Long)  '不分辨幻影1.24e
    Dim b(2) As Byte
        b(0) = &HC3
        b(1) = &HCC
    WriteProcessMemory Handle, ByVal Wargamedll + &H28357C, b(0), 2, 0&
End Sub

Sub 显示盟军头像124E(Handle As Long) '显示盟友英雄头像1.24e
    Dim TEXSL(18) As Byte
    TEXSL(0) = &HE8
    TEXSL(1) = &H3B
    TEXSL(2) = &H28
    TEXSL(3) = &H3
    TEXSL(4) = &H0
    TEXSL(5) = &H85
    TEXSL(6) = &HC0
    TEXSL(7) = &HF
    TEXSL(8) = &H84
    TEXSL(9) = &H8F
    TEXSL(10) = &H2
    TEXSL(11) = &H0
    TEXSL(12) = &H0
    TEXSL(13) = &HEB
    TEXSL(14) = &HC9
    TEXSL(15) = &H90
    TEXSL(16) = &H90
    TEXSL(17) = &H90
    TEXSL(18) = &H90
    WriteProcessMemory Handle, ByVal Wargamedll + &H371700, TEXSL(0), 19, 0&
End Sub

Sub 不显示头像124E(Handle As Long) '不显示英雄头像1.24e
    Dim TEXSL(18) As Byte
    TEXSL(0) = &HE8
    TEXSL(1) = &HFB
    TEXSL(2) = &H29
    TEXSL(3) = &H3
    TEXSL(4) = &H0
    TEXSL(5) = &H85
    TEXSL(6) = &HC0
    TEXSL(7) = &HF
    TEXSL(8) = &H84
    TEXSL(9) = &H8F
    TEXSL(10) = &H2
    TEXSL(11) = &H0
    TEXSL(12) = &H0
    TEXSL(13) = &H8B
    TEXSL(14) = &H85
    TEXSL(15) = &H80
    TEXSL(16) = &H1
    TEXSL(17) = &H0
    TEXSL(18) = &H0
    WriteProcessMemory Handle, ByVal Wargamedll + &H371700, TEXSL(0), 19, 0&
End Sub

Sub 显示敌方头像124E(Handle As Long) '显示敌方英雄头像1.24e
    Dim TXSl(18) As Byte
    TXSl(0) = &HE8
    TXSl(1) = &H3B
    TXSl(2) = &H28
    TXSl(3) = &H3
    TXSl(4) = &H0
    TXSl(5) = &H85
    TXSl(6) = &HC0
    TXSl(7) = &HF
    TXSl(8) = &H85
    TXSl(9) = &H8F
    TXSl(10) = &H2
    TXSl(11) = &H0
    TXSl(12) = &H0
    TXSl(13) = &HEB
    TXSl(14) = &HC9
    TXSl(15) = &H90
    TXSl(16) = &H90
    TXSl(17) = &H90
    TXSl(18) = &H90
    WriteProcessMemory Handle, ByVal Wargamedll + &H371700, TXSl(0), 19, 0&
End Sub
Sub 资源124E(Handle As Long)
    WriteProcessMemory Handle, ByVal Wargamedll + &H359A61, &HEB02, 2, 0&
    WriteProcessMemory Handle, ByVal Wargamedll + &H359AED, &HEB02, 2, 0&
    WriteProcessMemory Handle, ByVal Wargamedll + &H35A1DF, &HEB02, 2, 0&
    WriteProcessMemory Handle, ByVal Wargamedll + &H35A29F, &HEB02, 2, 0&
    WriteProcessMemory Handle, ByVal Wargamedll + &H35A3D0, &HEB02, 2, 0&
    WriteProcessMemory Handle, ByVal Wargamedll + &H28EAFA, &HEB02, 2, 0&
    WriteProcessMemory Handle, ByVal Wargamedll + &H4172EB, &HEB02, 2, 0&
    WriteProcessMemory Handle, ByVal Wargamedll + &H5B2D77, &HEB, 1, 0&
    WriteProcessMemory Handle, ByVal Wargamedll + &H5B2D8B, &H3, 1, 0&
End Sub

Sub 不资源124E(Handle As Long)
    WriteProcessMemory Handle, ByVal Wargamedll + &H359A61, &HC085, 2, 0&
    WriteProcessMemory Handle, ByVal Wargamedll + &H359AED, &HC085, 2, 0&
    WriteProcessMemory Handle, ByVal Wargamedll + &H35A1DF, &HC085, 2, 0&
    WriteProcessMemory Handle, ByVal Wargamedll + &H35A29F, &HC085, 2, 0&
    WriteProcessMemory Handle, ByVal Wargamedll + &H35A3D0, &HC085, 2, 0&
    WriteProcessMemory Handle, ByVal Wargamedll + &H28EAFA, &HC085, 2, 0&
    WriteProcessMemory Handle, ByVal Wargamedll + &H4172EB, &HC185, 2, 0&
    WriteProcessMemory Handle, ByVal Wargamedll + &H5B2D77, &H0, 1, 0&
    WriteProcessMemory Handle, ByVal Wargamedll + &H5B2D8B, &H1, 1, 0&
End Sub

Sub 可选视野外单位124E(Handle As Long) '可选视野外单位1.24e
    WriteProcessMemory Handle, ByVal Wargamedll + &H285CBC, &H9090, 2, 0&
    WriteProcessMemory Handle, ByVal Wargamedll + &H285CD2, &HEB, 1, 0&
End Sub

Sub 不可选视野外单位124E(Handle As Long) '可选视野外单位1.24e
    WriteProcessMemory Handle, ByVal Wargamedll + &H285CBC, &H2A74, 2, 0&
    WriteProcessMemory Handle, ByVal Wargamedll + &H285CD2, &H75, 1, 0&
End Sub

Sub 清除迷雾124E(Handle As Long)
    Dim Asc(5) As Byte
    Asc(0) = &HB2
    Asc(1) = &H0
    Asc(2) = &H90
    Asc(3) = &H90
    Asc(4) = &H90
    Asc(5) = &H90
    WriteProcessMemory Handle, ByVal Wargamedll + &H74D1B9, Asc(0), 6, 0&
    WriteProcessMemory Handle, ByVal Wargamedll + &H357065, &H9090, 2, 0&
End Sub
Sub 不清除迷雾124E(Handle As Long)
    Dim Asc(5) As Byte
    Asc(0) = &H8A
    Asc(1) = &H90
    Asc(2) = &H6C
    Asc(3) = &H7E
    Asc(4) = &HAB
    Asc(5) = &H6F
    WriteProcessMemory Handle, ByVal Wargamedll + &H74D1B9, Asc(0), 6, 0&
    WriteProcessMemory Handle, ByVal Wargamedll + &H357065, &H188, 2, 0&
End Sub
Sub 显示敌方信号124E(Handle As Long)
    WriteProcessMemory Handle, ByVal Wargamedll + &H43F9A6, &H3B, 1, 0&
    WriteProcessMemory Handle, ByVal Wargamedll + &H43F9A9, &H85, 1, 0&
    WriteProcessMemory Handle, ByVal Wargamedll + &H43F9B9, &H3B, 1, 0&
    WriteProcessMemory Handle, ByVal Wargamedll + &H43F9BC, &H85, 1, 0&
End Sub
Sub 不显示敌方信号124E(Handle As Long)
    WriteProcessMemory Handle, ByVal Wargamedll + &H43F9A6, &H85, 1, 0&
    WriteProcessMemory Handle, ByVal Wargamedll + &H43F9A9, &H84, 1, 0&
    WriteProcessMemory Handle, ByVal Wargamedll + &H43F9B9, &H85, 1, 0&
    WriteProcessMemory Handle, ByVal Wargamedll + &H43F9BC, &H84, 1, 0&
End Sub
Sub 显示技能124E(Handle As Long) '显示技能1.24e
    Dim AscSl(5) As Byte
    AscSl(0) = &H90
    AscSl(1) = &H90
    AscSl(2) = &H90
    AscSl(3) = &H90
    AscSl(4) = &H90
    AscSl(5) = &H90
    WriteProcessMemory Handle, ByVal Wargamedll + &H2031EC, AscSl(0), 6, 0&
    WriteProcessMemory Handle, ByVal Wargamedll + &H34FDE8, &H9090, 2, 0&
'顺便带上技能CD的显示
'地址：(28ECFE , "EB");
'地址：(34FE26 , "90 , 90 , 90 , 90");
    WriteProcessMemory Handle, ByVal Wargamedll + &H28ECFE, &HEB, 1, 0&
    WriteProcessMemory Handle, ByVal Wargamedll + &H34FE26, &H90909090, 4, 0&
End Sub
Sub 不显示技能124E(Handle As Long) '显示技能1.24e
    Dim AscSl(5) As Byte
    AscSl(0) = &HF
    AscSl(1) = &H84
    AscSl(2) = &H5F
    AscSl(3) = &H1
    AscSl(4) = &H0
    AscSl(5) = &H0
    WriteProcessMemory Handle, ByVal Wargamedll + &H2031EC, AscSl(0), 6, 0&
    WriteProcessMemory Handle, ByVal Wargamedll + &H34FDE8, &H874, 2, 0&
'顺便带上技能CD的显示
'地址：(28ECFE , "EB");
'地址：(34FE26 , "90 , 90 , 90 , 90");
    WriteProcessMemory Handle, ByVal Wargamedll + &H28ECFE, &H75, 1, 0&
    WriteProcessMemory Handle, ByVal Wargamedll + &H34FE26, &H874C085, 4, 0&
End Sub
Sub 大地图共享敌方视野124E(Handle As Long)
    WriteProcessMemory Handle, ByVal Wargamedll + &H74D191, 104, 1, 0&
    WriteProcessMemory Handle, ByVal Wargamedll + &H74D192, 255, 1, 0&
    WriteProcessMemory Handle, ByVal Wargamedll + &H74D193, 15, 1, 0&
End Sub
Sub 小地图共享敌方视野124E(Handle As Long)
    WriteProcessMemory Handle, ByVal Wargamedll + &H356FF9, 33, 1, 0&
    WriteProcessMemory Handle, ByVal Wargamedll + &H356FFA, 192, 1, 0&
End Sub
Sub 大地图不共享敌方视野124E(Handle As Long)
    WriteProcessMemory Handle, ByVal Wargamedll + &H74D191, 139, 1, 0&
    WriteProcessMemory Handle, ByVal Wargamedll + &H74D192, 84, 1, 0&
    WriteProcessMemory Handle, ByVal Wargamedll + &H74D193, 36, 1, 0&
End Sub
Sub 小地图不共享敌方视野124E(Handle As Long)
    WriteProcessMemory Handle, ByVal Wargamedll + &H356FF9, 35, 1, 0&
    WriteProcessMemory Handle, ByVal Wargamedll + &H356FFA, 194, 1, 0&
End Sub
Sub 大地图显示单位124E(Handle As Long)  '大地图单位显示1.24e
'地址：(3A201B , "EB");
'地址：(40A864 , "90 , 90");
'这里是显示物品的
WriteProcessMemory Handle, ByVal Wargamedll + &H3A201B, &HEB, 1, 0&
WriteProcessMemory Handle, ByVal Wargamedll + &H40A864, &H9090, 2, 0&
'这里是显示物品的------------------------------------————————————————————————————————
WriteProcessMemory Handle, ByVal Wargamedll + &H39EBBC, &H75, 1, 0&
WriteProcessMemory Handle, ByVal Wargamedll + &H3A2030, &H9090, 2, 0&
WriteProcessMemory Handle, ByVal Wargamedll + &H3A20DB, &H9090, 2, 0&
End Sub
Sub 大地图不显示单位124E(Handle As Long)
WriteProcessMemory Handle, ByVal Wargamedll + &H3A201B, &H75, 1, 0&
WriteProcessMemory Handle, ByVal Wargamedll + &H40A864, &HA75, 2, 0&
WriteProcessMemory Handle, ByVal Wargamedll + &H39EBBC, &H74, 1, 0&
WriteProcessMemory Handle, ByVal Wargamedll + &H3A2030, &H9EB, 2, 0&
WriteProcessMemory Handle, ByVal Wargamedll + &H3A20DB, &HCA23, 2, 0&
End Sub
Sub 小地图显示单位124E(Handle As Long)  '小地图单位显示1.24e
WriteProcessMemory Handle, ByVal Wargamedll + &H361F7C, &H0, 1, 0&
End Sub
Sub 小地图不显示单位124E(Handle As Long)
WriteProcessMemory Handle, ByVal Wargamedll + &H361F7C, &H1, 1, 0&
End Sub

Sub 大地图显示隐形124E(Handle As Long)  '大地图显示隐形单位1.24e
    Dim ASMC(5) As Byte
    ASMC(0) = &H90
    ASMC(1) = &H90
    ASMC(2) = &H90
    ASMC(3) = &H90
    ASMC(4) = &H90
    ASMC(5) = &H90
    Dim ASMCSL(10) As Byte
    ASMCSL(0) = &H90
    ASMCSL(1) = &H90
    ASMCSL(2) = &H90
    ASMCSL(3) = &H90
    ASMCSL(4) = &H90
    ASMCSL(5) = &H90
    ASMCSL(6) = &H90
    ASMCSL(7) = &H90
    ASMCSL(8) = &H33
    ASMCSL(9) = &HC0
    ASMCSL(10) = &H40
    WriteProcessMemory Handle, ByVal Wargamedll + &H362391, &H3B, 1, 0&
    WriteProcessMemory Handle, ByVal Wargamedll + &H362394, &H85, 1, 0&
    WriteProcessMemory Handle, ByVal Wargamedll + &H39A51B, ASMC(0), 6, 0&
    WriteProcessMemory Handle, ByVal Wargamedll + &H39A52E, ASMCSL(0), 11, 0&
End Sub
Sub 大地图不显示隐形124E(Handle As Long)
    Dim ASMC(5) As Byte
    ASMC(0) = &H8B
    ASMC(1) = &H97
    ASMC(2) = &H98
    ASMC(3) = &H1
    ASMC(4) = &H0
    ASMC(5) = &H0
    Dim ASMCSL(10) As Byte
    ASMCSL(0) = &HF
    ASMCSL(1) = &HB7
    ASMCSL(2) = &H0
    ASMCSL(3) = &H55
    ASMCSL(4) = &H50
    ASMCSL(5) = &H56
    ASMCSL(6) = &HE8
    ASMCSL(7) = &HF7
    ASMCSL(8) = &H7B
    ASMCSL(9) = &H0
    ASMCSL(10) = &H0
    WriteProcessMemory Handle, ByVal Wargamedll + &H362391, &H85, 1, 0&
    WriteProcessMemory Handle, ByVal Wargamedll + &H362394, &H84, 1, 0&
    WriteProcessMemory Handle, ByVal Wargamedll + &H39A51B, ASMC(0), 6, 0&
    WriteProcessMemory Handle, ByVal Wargamedll + &H39A52E, ASMCSL(0), 11, 0&
End Sub
Sub 小地图显示隐形124E(Handle As Long)  '小地图显示隐形单位1.24e
'''    地址：(362391 , "3B");
'''    地址：(362394 , "85");
'''    地址：(39A51B , "90 , 90 , 90 , 90 , 90 , 90");
'''    地址：(39A52E , "90 , 90 , 90 , 90 , 90 , 90 , 90 , 90 , 33 , C0 , 40");
    WriteProcessMemory Handle, ByVal Wargamedll + &H362391, &H3B, 1, 0&
    WriteProcessMemory Handle, ByVal Wargamedll + &H362394, &H85, 1, 0&
    Dim b(5) As Byte
        b(0) = &H90
        b(1) = &H90
        b(2) = &H90
        b(3) = &H90
        b(4) = &H90
        b(5) = &H90
    WriteProcessMemory Handle, ByVal Wargamedll + &H39A51B, b(0), 6, 0&
    Dim c(10) As Byte
        c(0) = &H90
        c(1) = &H90
        c(2) = &H90
        c(3) = &H90
        c(4) = &H90
        c(5) = &H90
        c(6) = &H90
        c(7) = &H90
        c(8) = &H33
        c(9) = &HC0
        c(10) = &H40
    WriteProcessMemory Handle, ByVal Wargamedll + &H39A52E, c(0), 11, 0&
End Sub

Sub 小地图不显示隐形124E(Handle As Long)
    WriteProcessMemory Handle, ByVal Wargamedll + &H362391, &H85, 1, 0&
    WriteProcessMemory Handle, ByVal Wargamedll + &H362394, &H84, 1, 0&
    Dim b(5) As Byte
        b(0) = &H8B
        b(1) = &H97
        b(2) = &H98
        b(3) = &H1
        b(4) = &H0
        b(5) = &H0
    WriteProcessMemory Handle, ByVal Wargamedll + &H39A51B, b(0), 6, 0&
    Dim c(10) As Byte
        c(0) = &HF
        c(1) = &HB7
        c(2) = &H0
        c(3) = &H55
        c(4) = &H50
        c(5) = &H56
        c(6) = &HE8
        c(7) = &HF7
        c(8) = &H7B
        c(9) = &H0
        c(10) = &H0
    WriteProcessMemory Handle, ByVal Wargamedll + &H39A52E, c(0), 11, 0&
End Sub
Sub 显血124E()
Dim hwnd As Long, Handle As Long, PID As Long, Address As Long
hwnd = FindWindow(vbNullString, "Warcraft III")
GetWindowThreadProcessId hwnd, PID
Handle = OpenProcess(PROCESS_ALL_ACCESS, False, PID)
Call Getgamedll
    ReadProcessMemory Handle, ByVal Wargamedll + &HACBDD8, Address, 4, 0&
    WriteProcessMemory Handle, ByVal Address + &H1D8, 1, 2, 0&
CloseHandle Handle
End Sub
Sub 不显血124E()
Dim hwnd As Long, Handle As Long, PID As Long, Address As Long
hwnd = FindWindow(vbNullString, "Warcraft III")
GetWindowThreadProcessId hwnd, PID
Handle = OpenProcess(PROCESS_ALL_ACCESS, False, PID)
Call Getgamedll
    ReadProcessMemory Handle, ByVal Wargamedll + &HACBDD8, Address, 4, 0&
    WriteProcessMemory Handle, ByVal Address + &H1D8, 0, 2, 0&
CloseHandle Handle
End Sub
Function 获取显血状态124E() As Long
Dim hwnd As Long, Handle As Long, PID As Long, Address As Long
hwnd = FindWindow(vbNullString, "Warcraft III")
GetWindowThreadProcessId hwnd, PID
Handle = OpenProcess(PROCESS_ALL_ACCESS, False, PID)
Call Getgamedll
    ReadProcessMemory Handle, ByVal Wargamedll + &HACBDD8, Address, 4, 0&
    ReadProcessMemory Handle, ByVal Address + &H1D8, 获取显血状态124E, 4, 0&
CloseHandle Handle
End Function
Function 获取游戏状态124E() As Long  '0为未开始，大于0为已开始
Dim hwnd As Long, Handle As Long, PID As Long, Address As Long
hwnd = FindWindow(vbNullString, "Warcraft III")
GetWindowThreadProcessId hwnd, PID
Handle = OpenProcess(PROCESS_ALL_ACCESS, False, PID)
Call Getgamedll
ReadProcessMemory Handle, ByVal Wargamedll + &HACAA5C, 获取游戏状态124E, 1, 0&
CloseHandle Handle
End Function
Function 进入房间状态124E() As Long  '0为未进入，1为已进入
Dim hwnd As Long, Handle As Long, PID As Long
hwnd = FindWindow(vbNullString, "Warcraft III")
GetWindowThreadProcessId hwnd, PID
Handle = OpenProcess(PROCESS_ALL_ACCESS, False, PID)
Call Getgamedll
ReadProcessMemory Handle, ByVal Wargamedll + &HAE8450, 进入房间状态124E, 1, 0&
CloseHandle Handle
End Function

