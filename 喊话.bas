Attribute VB_Name = "����"
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const WM_KEYDOWN = &H100
Private Const WM_KEYUP = &H101


Public Sub SendString(Str As String, Optional i As Integer = 0) 'I=1��������˺���
Dim hWnd As Long, aStr As String
hWnd = FindWindow(vbNullString, "Warcraft III")
If Str <> "" Then
    If i = 1 Then
        aStr = Clipboard.GetText                    '����ԭ�������ı�
        Clipboard.Clear                             '���ԭ�������ı�
        ClipboardSetText hWnd, UTF8_Encode(Str)     '���ü������ı�
        PostMessage hWnd, WM_KEYDOWN, 16, 0         '����SHIFT
        PostMessage hWnd, WM_KEYDOWN, 13, 0         '���»س�
        PostMessage hWnd, WM_KEYUP, 16, 0           '�ɿ�SHIFT
        PostMessage hWnd, WM_KEYDOWN, 17, 0         '����Ctrl
        PostMessage hWnd, WM_KEYDOWN, 86, 0         '����v
        PostMessage hWnd, WM_KEYUP, 86, 0           '�ɿ�v
        PostMessage hWnd, WM_KEYUP, 17, 0           '�ɿ�Ctrl
        PostMessage hWnd, WM_KEYDOWN, 13, 0         '���»س�
        PostMessage hWnd, WM_KEYUP, 13, 0           '�ɿ��س�
        Delay 100
        Clipboard.Clear                             '���ԭ�������ı�
        Clipboard.SetText aStr                      '�ָ�ԭ�������ı�
    ElseIf i = 2 Then
        aStr = Clipboard.GetText                    '����ԭ�������ı�
        Clipboard.Clear                             '���ԭ�������ı�
        ClipboardSetText hWnd, UTF8_Encode(Str)     '���ü������ı�
        PostMessage hWnd, WM_KEYDOWN, 17, 0         '����Ctrl
        PostMessage hWnd, WM_KEYDOWN, 86, 0         '����v
        PostMessage hWnd, WM_KEYUP, 86, 0           '�ɿ�v
        PostMessage hWnd, WM_KEYUP, 17, 0           '�ɿ�Ctrl
        PostMessage hWnd, WM_KEYDOWN, 13, 0         '���»س�
        PostMessage hWnd, WM_KEYUP, 13, 0           '�ɿ��س�
       Delay 100
        PostMessage hWnd, WM_KEYDOWN, 8, 0          '�����˸�
        PostMessage hWnd, WM_KEYUP, 8, 0            '�ɿ��˸�
        PostMessage hWnd, WM_KEYDOWN, 8, 0          '�����˸�
        PostMessage hWnd, WM_KEYUP, 8, 0            '�ɿ��˸�
        Delay 100
        Clipboard.Clear                             '���ԭ�������ı�
        Clipboard.SetText aStr                      '�ָ�ԭ�������ı�
    ElseIf i = 0 Then
        aStr = Clipboard.GetText                    '����ԭ�������ı�
        Clipboard.Clear                             '���ԭ�������ı�
        ClipboardSetText hWnd, UTF8_Encode(Str)     '���ü������ı�
        PostMessage hWnd, WM_KEYDOWN, 17, 0         '����Ctrl
        PostMessage hWnd, WM_KEYDOWN, 13, 0         '���»س�
        PostMessage hWnd, WM_KEYUP, 17, 0           '�ɿ�Ctrl
        PostMessage hWnd, WM_KEYDOWN, 17, 0         '����Ctrl
        PostMessage hWnd, WM_KEYDOWN, 86, 0         '����v
        PostMessage hWnd, WM_KEYUP, 86, 0           '�ɿ�v
        PostMessage hWnd, WM_KEYUP, 17, 0           '�ɿ�Ctrl
        PostMessage hWnd, WM_KEYDOWN, 13, 0         '���»س�
        PostMessage hWnd, WM_KEYUP, 13, 0           '�ɿ��س�
        Delay 100
        Clipboard.Clear                             '���ԭ�������ı�
        Clipboard.SetText aStr                      '�ָ�ԭ�������ı�
    End If
End If
End Sub





                                    
'ת��ģ��ժ�����磬�д��޸�:HhAPi���ַ�����ת��
