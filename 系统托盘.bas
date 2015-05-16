Attribute VB_Name = "ϵͳ����"
Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Const NIM_ADD = 0
Const NIM_DELETE = 2
Const NIM_MODIFY = 1
Private Const WM_USER = &H400
Private Const NIN_BALLOONUSERCLICK = (WM_USER + &H5)
Private Const NIN_BALLOONTIMEOUT = (WM_USER + &H4)
Private Const GWL_WNDPROC = (-4)
Private Const WM_NOTIFYICON = WM_USER + 1
Private lngPreWndProc As Long
Public Const NIIF_INFO = &H1
Public Const WM_RBUTTONUP = &H205
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONDBLCLK = &H206
Public Type NOTIFYICONDATA
          cbSize As Long
          hwnd As Long
          Uid As Long
          UFlags As Long
          UCallbackMessage As Long
          Hicon As Long
          SzTip As String * 128
          SzInfo As String * 256
          SzInfoTitle As String * 64
          DwInfoFlags As Long
          uTimeoutOrVersion As Long   '����VB��û��Union���ͣ�ֻ����Long�ʹ���
          dwState As Long
          dwStateMask As Long
End Type

Private TheData  As NOTIFYICONDATA         '��������ͼ������
Private Bremind As NOTIFYICONDATA
Public Sub ChangeIcon(Handle As Long)
With TheData
        .Uid = 0
        .hwnd = Form1.hwnd                   '���ΪForm1�ľ��
        .cbSize = Len(TheData)
        .Hicon = Handle
        .UFlags = 2
        .UCallbackMessage = WM_MOUSEMOVE     '�ص�����MOUSEMOVE
        .UFlags = &H1 Or &H2 Or &H4 Or &H10
        .cbSize = Len(TheData)
      End With
      Shell_NotifyIcon NIM_MODIFY, TheData
End Sub
Public Sub DelIcon()
With TheData
        .UFlags = 0
End With
Shell_NotifyIcon NIM_DELETE, TheData
SetWindowLong Form1.hwnd, GWL_WNDPROC, lngPreWndProc
lngPreWndProc = 0
End Sub
Public Sub AddIcon()
With TheData
        .Uid = 0
        .hwnd = Form1.hwnd
        .cbSize = Len(TheData)
        .Hicon = Form1.ImageA.Picture.Handle
        .UFlags = 2 Or 1 Or 4
        .UCallbackMessage = WM_MOUSEMOVE
      End With
      Shell_NotifyIcon NIM_ADD, TheData
End Sub
Public Sub Remind() '����������Ϸ�ѿ�ʼ
Form1.Timer1.Enabled = False
With Bremind
        .cbSize = Len(Bremind)
        .hwnd = Form1.hwnd
        .Uid = 0
        .UFlags = &H1 Or &H2 Or &H4 Or &H10
        .Hicon = Form1.Image3.Picture.Handle
        .SzTip = "                ��ܰ��ʾ" & vbNullChar
        .dwState = 0
        .dwStateMask = 0
        .SzInfo = "        ��Ϸ�Ѿ���ʼ�ˡ�����" & vbNullChar
        .uTimeoutOrVersion = 120000
        .SzInfoTitle = "            �����İ���" & vbNullChar
        .DwInfoFlags = 1
        .UCallbackMessage = WM_NOTIFYICON
    End With
    Shell_NotifyIcon NIM_MODIFY, Bremind
    If lngPreWndProc = 0 Then lngPreWndProc = SetWindowLong(Form1.hwnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub
Function WindowProc(ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
   On Error Resume Next
    If msg = WM_NOTIFYICON Then
        Select Case lParam
            Case WM_LBUTTONUP '����������ͼ����Ӧ�¼�
                ShowWindow FindWindow(vbNullString, "Warcraft III"), 1
                Form1.Timer1.Enabled = True
            Case WM_RBUTTONUP '�Ҽ��������ͼ����Ӧ�¼�
                ShowWindow FindWindow(vbNullString, "Warcraft III"), 1
                Form1.Timer1.Enabled = True
            Case WM_MOUSEUP '����Ƶ�����ͼ����Ӧ�¼�
                 
            
            Case NIN_BALLOONTIMEOUT '������ʾ��ʧ
                Call Remind
            Case NIN_BALLOONUSERCLICK '����������ʾ��Ӧ�¼�
                ShowWindow FindWindow(vbNullString, "Warcraft III"), 1
                Form1.Timer1.Enabled = True
        End Select
    End If
    WindowProc = CallWindowProc(lngPreWndProc, hwnd, msg, wParam, lParam)
End Function
