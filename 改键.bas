Attribute VB_Name = "改键"

Public Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public WindowText As String * 255   '前台窗体文本
Public Const WH_KEYBOARD_LL = 13
Private Const WM_KEYDOWN = &H100
Private Const WM_KEYUP = &H101
     Public Type MSG
           vKey As Long
           sKey As Long
           flag As Long
           time As Long
     End Type
Public mymsg As MSG
Public aKeycode(1 To 10) As Long '改键aKeycode→bKeycode
Public bKeycode(1 To 10) As Long
Dim KS As Integer
Public Function MyKBHook(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim hwnd As Long, q As Integer
    GetWindowText GetForegroundWindow, WindowText, 255                  '获取前台窗体标题
    If nCode = 0 And Left(WindowText, 12) = "Warcraft III" And ChatState = 0 Then
        CopyMemory mymsg, ByVal lParam, Len(mymsg)
        If Form1.CHKkey.Value = 1 Then
            For i = 1 To 10
                If aKeycode(i) > 0 Then
                    If mymsg.sKey > 0 Then '如果扫描码大于0，'为了区别是按下键盘发生的还是程序模拟的，因为模拟按键时没有发送扫描码
                        If mymsg.vKey = aKeycode(i) And bKeycode(i) > 0 Then
                            hwnd = FindWindow(vbNullString, "Warcraft III")
                            If wParam = WM_KEYDOWN Then PostMessage hwnd, WM_KEYDOWN, bKeycode(i), 0
                            If wParam = WM_KEYUP Then PostMessage hwnd, WM_KEYUP, bKeycode(i), 0
                            MyKBHook = 1 '吃掉消息
                        End If
                    End If
                End If
            Next
        End If
        '按下~后喊话，为了避免按下~后选中农民，用钩子吃掉消息
        For i = 1 To 10
            If aKeycode(i) = 192 And Form1.CHKkey.Value = 1 Then Exit Function
        Next
        If Form1.CHKHh.Value = 1 And Form1.HhText5.Text <> "" Then '如果已开启喊话并且对应文本框不为空
            If mymsg.vKey = 192 And wParam = WM_KEYDOWN Then  '按下~
                    If KS = 0 Then SendString Form1.HhText5.Text
                    MyKBHook = 1
                    If KS = 0 Then KS = 1
            End If
        End If
        If mymsg.vKey = 192 And wParam = WM_KEYUP And KS = 1 Then KS = 0
        '
    Else
        MyKBHook = CallNextHookEx(hHook, nCode, wParam, lParam)
    End If
End Function
