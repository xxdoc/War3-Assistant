Attribute VB_Name = "��ʱ"
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Const WM_KEYDOWN = &H100
Private Const WM_KEYUP = &H101
Public Sub GetTime()
Dim gTime As String, a As String, i As Integer, hWnd As Long, bStr As String, yStr As String, num As Integer, tt As Integer, N As Integer
'----------------------------------------------'�޸�ʱ���ʽ
    If left(time, 1) = "��" Or left(time, 1) = "��" Then  '�����ʮ��Сʱ��
        If left(time, 1) = "��" Then
            gTime = time
            gTime = right(gTime, 8)
            gTime = Date & "  " & gTime
        ElseIf left(time, 1) = "��" And Mid(time, 4, 2) = 12 Then
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
            Mid(gTime, N, 1) = "��"
            tt = 1
        ElseIf Mid(gTime, N, 1) = "-" And tt = 1 Then
                Mid(gTime, N, 1) = "��"
                tt = 3
        ElseIf Mid(gTime, N, 1) = " " And tt = 3 Then
            Mid(gTime, N, 1) = "��"
            tt = 0
        End If
    Next N
    
    Else                                               '����Ƕ�ʮ��Сʱ��
        gTime = Date
        num = 1
        For i = 5 To 11
            If Mid(gTime, i, 1) = "/" Or Mid(gTime, i, 1) = "-" Then
                If num = 1 Then
                    Mid(gTime, i, 1) = "��"
                    num = 2
                ElseIf num = 2 Then
                    Mid(gTime, i, 1) = "��"
                    num = 1
                End If
            End If
        Next
        gTime = gTime & "�� "
        If Len(time) = 7 Then gTime = gTime & "0" & time
        If Len(time) = 8 Then gTime = gTime & time
    End If
'-----------------------------------------------------------------------
    hWnd = FindWindow(vbNullString, "Warcraft III")
If ChatState = 0 And ��ȡ��Ϸ״̬ > 0 Then
    PostMessage hWnd, WM_KEYDOWN, 16, 0         '����SHIFT
    SendString gTime
    PostMessage hWnd, WM_KEYUP, 16, 0           '�ɿ�HIFT
    Delay 900
End If
If ChatState = 1 And ���뷿��״̬124E = 1 Then '������ھ�������������Ϸ��û��ʼ��Ҫ���������˸�ɾ�������V

        yStr = Clipboard.GetText                    '����ԭ�������ı�
        Clipboard.Clear                             '���ԭ�������ı�
        ClipboardSetText hWnd, UTF8_Encode(gTime)   '���ü������ı�
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
        Clipboard.SetText yStr                      '�ָ�ԭ�������ı�
        Delay 800
End If
End Sub

Public Function ��ȡ��Ϸ״̬() As Long
Select Case ��ȡħ�ް汾
Case "1.24E": ��ȡ��Ϸ״̬ = ��ȡ��Ϸ״̬124E
Case "1.24B": ��ȡ��Ϸ״̬ = ��ȡ��Ϸ״̬124B
Case "1.20E": ��ȡ��Ϸ״̬ = ��ȡ��Ϸ״̬120E
End Select
End Function
