Attribute VB_Name = "报时"
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Const WM_KEYDOWN = &H100
Private Const WM_KEYUP = &H101
Public Sub GetTime()
Dim gTime As String, a As String, i As Integer, hWnd As Long, bStr As String, yStr As String, num As Integer, tt As Integer, N As Integer
'----------------------------------------------'修改时间格式
    If left(time, 1) = "上" Or left(time, 1) = "下" Then  '如果是十二小时制
        If left(time, 1) = "上" Then
            gTime = time
            gTime = right(gTime, 8)
            gTime = Date & "  " & gTime
        ElseIf left(time, 1) = "下" And Mid(time, 4, 2) = 12 Then
            gTime = time
            gTime = right(gTime, 8)
            gTime = Date & "  " & gTime
        Else
            gTime = time
            gTime = right(gTime, 8)
            a = left(gTime, 2)
            i = Val(a) + 12
            gTime = i & right(gTime, 6)
            gTime = Date & "  " & gTime
        End If
    For N = 5 To 11
        If Mid(gTime, N, 1) = "-" And tt = 0 Then
            Mid(gTime, N, 1) = "年"
            tt = 1
        ElseIf Mid(gTime, N, 1) = "-" And tt = 1 Then
                Mid(gTime, N, 1) = "月"
                tt = 3
        ElseIf Mid(gTime, N, 1) = " " And tt = 3 Then
            Mid(gTime, N, 1) = "日"
            tt = 0
        End If
    Next N
    
    Else                                               '如果是二十四小时制
        gTime = Date
        num = 1
        For i = 5 To 11
            If Mid(gTime, i, 1) = "/" Or Mid(gTime, i, 1) = "-" Then
                If num = 1 Then
                    Mid(gTime, i, 1) = "年"
                    num = 2
                ElseIf num = 2 Then
                    Mid(gTime, i, 1) = "月"
                    num = 1
                End If
            End If
        Next
        gTime = gTime & "日 "
        If Len(time) = 7 Then gTime = gTime & "0" & time
        If Len(time) = 8 Then gTime = gTime & time
    End If
'-----------------------------------------------------------------------
    hWnd = FindWindow(vbNullString, "Warcraft III")
If ChatState = 0 And 获取游戏状态 > 0 Then
    PostMessage hWnd, WM_KEYDOWN, 16, 0         '按下SHIFT
    SendString gTime
    PostMessage hWnd, WM_KEYUP, 16, 0           '松开HIFT
    Delay 900
End If
If ChatState = 1 And 进入房间状态124E = 1 Then '如果是在局域网房间内游戏还没开始，要按下两次退格删掉多余的V

        yStr = Clipboard.GetText                    '保存原剪贴板文本
        Clipboard.Clear                             '清空原剪贴板文本
        ClipboardSetText hWnd, UTF8_Encode(gTime)   '设置剪贴板文本
        PostMessage hWnd, WM_KEYDOWN, 17, 0         '按下Ctrl
        PostMessage hWnd, WM_KEYDOWN, 86, 0         '按下v
        PostMessage hWnd, WM_KEYUP, 86, 0           '松开v
        PostMessage hWnd, WM_KEYUP, 17, 0           '松开Ctrl
        PostMessage hWnd, WM_KEYDOWN, 13, 0         '按下回车
        PostMessage hWnd, WM_KEYUP, 13, 0           '松开回车
        Delay 100
        PostMessage hWnd, WM_KEYDOWN, 8, 0          '按下退格
        PostMessage hWnd, WM_KEYUP, 8, 0            '松开退格
        PostMessage hWnd, WM_KEYDOWN, 8, 0          '按下退格
        PostMessage hWnd, WM_KEYUP, 8, 0            '松开退格
        Delay 100
        Clipboard.Clear                             '清空原剪贴板文本
        Clipboard.SetText yStr                      '恢复原剪贴板文本
        Delay 800
End If
End Sub

Public Function 获取游戏状态() As Long
Select Case 获取魔兽版本
Case "1.24E": 获取游戏状态 = 获取游戏状态124E
Case "1.24B": 获取游戏状态 = 获取游戏状态124B
Case "1.20E": 获取游戏状态 = 获取游戏状态120E
End Select
End Function
