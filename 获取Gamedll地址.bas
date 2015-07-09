Attribute VB_Name = "获取Gamedll地址"
Public Wargamedll As Long
Public WarStormdll As Long
Public WAR3dll As Long
Public HFGameShelldll As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function Module32First Lib "kernel32" (ByVal hSnapShot As Long, uProcess As MODULEENTRY32) As Long
Private Declare Function Module32Next Lib "kernel32" (ByVal hSnapShot As Long, uProcess As MODULEENTRY32) As Long
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Private Declare Function IsIconic Lib "user32" (ByVal hwnd As Long) As Long
Private Type MODULEENTRY32
   dwSize As Long
   th32ModuleID As Long
   th32ProcessID As Long
   GlblcntUsage As Long
   ProccntUsage As Long
   modBaseAddr As Long
   modBaseSize As Long
   hModule As Long
   szModule As String * 256
   szExePath As String * 260
End Type
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Public Function dlljz(p As Long, na As String) As Long
Dim WindowHandle As Long, ChildWindowHandle As Long
         Dim hSnapShot, N
         Dim uProcess As MODULEENTRY32
         uProcess.th32ProcessID = p
         hSnapShot = CreateToolhelp32Snapshot(8, uProcess.th32ProcessID)
         uProcess.dwSize = Len(uProcess)
         N = Module32First(hSnapShot, uProcess)
         Do While N
                If LCase(left(uProcess.szModule, InStr(uProcess.szModule, Chr(0)) - 1)) = na Then
                 dlljz = CLng(uProcess.modBaseAddr)
                 GoTo guanbi
                End If
         N = Module32Next(hSnapShot, uProcess)
         Loop
guanbi: CloseHandle hSnapShot
End Function
Public Function Getgamedll()
    Dim a As Long, hwnd As Long, PID As Long
    hwnd = FindWindow(vbNullString, "Warcraft III")
    GetWindowThreadProcessId hwnd, PID
    Wargamedll = dlljz(PID, "game.dll")
End Function

Public Function GetHFGameShelldll()
    Dim a As Long, hwnd As Long, PID As Long
    hwnd = FindWindow(vbNullString, "浩方电竞平台 - 6.0.0.0521(RC7)")
    GetWindowThreadProcessId hwnd, PID
    HFGameShelldll = dlljz(PID, "gameshell.dll")
End Function

Public Function GetGGWAR3dll()
    Dim a As Long, hwnd As Long, PID As Long
    hwnd = FindWindow(vbNullString, "Warcraft III")
    GetWindowThreadProcessId hwnd, PID
    WAR3dll = dlljz(PID, "ggwar3.dll")
End Function
