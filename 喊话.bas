Attribute VB_Name = "喊话"
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const WM_KEYDOWN = &H100
Private Const WM_KEYUP = &H101


Public Sub SendString(Str As String, Optional i As Integer = 0) 'I=1则对所有人喊话
Dim hWnd As Long, aStr As String
hWnd = FindWindow(vbNullString, "Warcraft III")
If Str <> "" Then
    If i = 1 Then
        aStr = Clipboard.GetText                    '保存原剪贴板文本
        Clipboard.Clear                             '清空原剪贴板文本
        ClipboardSetText hWnd, UTF8_Encode(Str)     '设置剪贴板文本
        PostMessage hWnd, WM_KEYDOWN, 16, 0         '按下SHIFT
        PostMessage hWnd, WM_KEYDOWN, 13, 0         '按下回车
        PostMessage hWnd, WM_KEYUP, 16, 0           '松开SHIFT
        PostMessage hWnd, WM_KEYDOWN, 17, 0         '按下Ctrl
        PostMessage hWnd, WM_KEYDOWN, 86, 0         '按下v
        PostMessage hWnd, WM_KEYUP, 86, 0           '松开v
        PostMessage hWnd, WM_KEYUP, 17, 0           '松开Ctrl
        PostMessage hWnd, WM_KEYDOWN, 13, 0         '按下回车
        PostMessage hWnd, WM_KEYUP, 13, 0           '松开回车
        Delay 100
        Clipboard.Clear                             '清空原剪贴板文本
        Clipboard.SetText aStr                      '恢复原剪贴板文本
    ElseIf i = 2 Then
        aStr = Clipboard.GetText                    '保存原剪贴板文本
        Clipboard.Clear                             '清空原剪贴板文本
        ClipboardSetText hWnd, UTF8_Encode(Str)     '设置剪贴板文本
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
        Clipboard.SetText aStr                      '恢复原剪贴板文本
    ElseIf i = 0 Then
        aStr = Clipboard.GetText                    '保存原剪贴板文本
        Clipboard.Clear                             '清空原剪贴板文本
        ClipboardSetText hWnd, UTF8_Encode(Str)     '设置剪贴板文本
        PostMessage hWnd, WM_KEYDOWN, 17, 0         '按下Ctrl
        PostMessage hWnd, WM_KEYDOWN, 13, 0         '按下回车
        PostMessage hWnd, WM_KEYUP, 17, 0           '松开Ctrl
        PostMessage hWnd, WM_KEYDOWN, 17, 0         '按下Ctrl
        PostMessage hWnd, WM_KEYDOWN, 86, 0         '按下v
        PostMessage hWnd, WM_KEYUP, 86, 0           '松开v
        PostMessage hWnd, WM_KEYUP, 17, 0           '松开Ctrl
        PostMessage hWnd, WM_KEYDOWN, 13, 0         '按下回车
        PostMessage hWnd, WM_KEYUP, 13, 0           '松开回车
        Delay 100
        Clipboard.Clear                             '清空原剪贴板文本
        Clipboard.SetText aStr                      '恢复原剪贴板文本
    End If
End If
End Sub





                                    
'转码模块摘自网络，有待修改:HhAPi和字符编码转换
