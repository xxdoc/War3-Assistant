VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "War3 Assistant"
   ClientHeight    =   3735
   ClientLeft      =   8235
   ClientTop       =   5520
   ClientWidth     =   4710
   BeginProperty Font 
      Name            =   "����"
      Size            =   21.75
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command13 
      Caption         =   "����Ƥ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   19
      ToolTipText     =   "���������������������򼴿�"
      Top             =   3240
      Width           =   975
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3960
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command12 
      Caption         =   "��ȡ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3840
      TabIndex        =   18
      ToolTipText     =   "�����Ϸ����ʱ������Զ���ȡ"
      Top             =   2640
      Width           =   735
   End
   Begin VB.TextBox Text5 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1080
      TabIndex        =   17
      Top             =   2640
      Width           =   2655
   End
   Begin VB.CommandButton Command11 
      Caption         =   "��ϵ����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   10
      ToolTipText     =   "��ϵ����"
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton Command10 
      Caption         =   "ʹ�ð���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   12
      ToolTipText     =   "�鿴����"
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton Command9 
      Caption         =   "����Ĭ������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   11
      ToolTipText     =   "������������Ϣ������������Ȼ����"
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   840
      Width           =   735
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   120
      Top             =   4200
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   3360
      Top             =   4320
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   720
      TabIndex        =   9
      ToolTipText     =   "���ѵ�����"
      Top             =   2040
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   7
      ToolTipText     =   "���ʱ��"
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "��������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1440
      Width           =   1575
   End
   Begin VB.TextBox GMTX 
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   180
      TabIndex        =   4
      ToolTipText     =   "����"
      Top             =   840
      Width           =   2745
   End
   Begin VB.CommandButton Command2 
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "����"
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "����ħ�޷ֱ���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "���ħ��ȫ��ʱ�޷�ȫ�����ô˹����޸�����ֵ����Ļ�ֱ�����ͬ������ħ�޺���Ч"
      Top             =   120
      Width           =   1395
   End
   Begin VB.TextBox HeightTx 
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1920
      TabIndex        =   2
      Text            =   "768"
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox WidthTx 
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   240
      TabIndex        =   0
      Text            =   "1024"
      Top             =   120
      Width           =   855
   End
   Begin VB.Line Line6 
      X1              =   0
      X2              =   5040
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "��Ϸ·��:"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   16
      Top             =   2760
      Width           =   945
   End
   Begin VB.Line Line5 
      Visible         =   0   'False
      X1              =   120
      X2              =   0
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line Line4 
      X1              =   0
      X2              =   4560
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   4560
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   4680
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line1 
      Visible         =   0   'False
      X1              =   120
      X2              =   4560
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1560
      TabIndex        =   15
      Top             =   1560
      Width           =   480
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "���ݣ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   14
      Top             =   2040
      Width           =   720
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "�����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   13
      Top             =   1560
      Width           =   720
   End
   Begin VB.Label Label3 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "����"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Top             =   0
      Width           =   375
   End
   Begin VB.Menu PF 
      Caption         =   "����Ƥ��"
      Visible         =   0   'False
      Begin VB.Menu PF0 
         Caption         =   "Ĭ��Ƥ��"
      End
      Begin VB.Menu PF1 
         Caption         =   "china"
      End
      Begin VB.Menu PF2 
         Caption         =   "MSN"
      End
      Begin VB.Menu PF4 
         Caption         =   "QQ2008"
      End
      Begin VB.Menu PF3 
         Caption         =   "QQ2009"
      End
      Begin VB.Menu PF5 
         Caption         =   "QQ2010"
      End
      Begin VB.Menu PF6 
         Caption         =   "QQӰ��"
      End
      Begin VB.Menu PF7 
         Caption         =   "QQGame"
      End
      Begin VB.Menu PF8 
         Caption         =   "REAL"
      End
      Begin VB.Menu PF9 
         Caption         =   "adamant"
      End
      Begin VB.Menu PF10 
         Caption         =   "asus"
      End
      Begin VB.Menu PF11 
         Caption         =   "black"
      End
      Begin VB.Menu PF12 
         Caption         =   "darkroyale"
      End
      Begin VB.Menu PF13 
         Caption         =   "dogmax"
      End
      Begin VB.Menu PF14 
         Caption         =   "elegance"
      End
      Begin VB.Menu PF15 
         Caption         =   "enjoy"
      End
      Begin VB.Menu PF16 
         Caption         =   "gem"
      End
      Begin VB.Menu PF17 
         Caption         =   "hlong"
      End
      Begin VB.Menu PF18 
         Caption         =   "homestead"
      End
      Begin VB.Menu PF19 
         Caption         =   "itunes"
      End
      Begin VB.Menu PF20 
         Caption         =   "longhorn"
      End
      Begin VB.Menu PF21 
         Caption         =   "office2007"
      End
      Begin VB.Menu PF22 
         Caption         =   "pixos"
      End
      Begin VB.Menu PF23 
         Caption         =   "royale"
      End
      Begin VB.Menu PF24 
         Caption         =   "storm"
      End
      Begin VB.Menu PF25 
         Caption         =   "vista"
      End
      Begin VB.Menu PF26 
         Caption         =   "whitefire"
      End
      Begin VB.Menu PF27 
         Caption         =   "wish"
      End
      Begin VB.Menu PF28 
         Caption         =   "��ľ"
      End
      Begin VB.Menu PF29 
         Caption         =   "����"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef dwFlags As Long, ByVal dwReserved As Long) As Long
Dim LoopNum As Long '�����������ӵ�Timerѭ������
Dim NowNum As Long '�������浱ǰ���ӵ�Timerѭ������
Dim St As Integer '����������Ƿ��Ѿ�������ֵ�UTF8�����ı�
Dim MyName As String
Private Sub Command1_Click() '����ħ�޷ֱ���
If WidthTx.Text <> "" And HeightTx.Text <> "" Then
    For i = 1 To Len(WidthTx.Text) '����ı������ݣ�����з������ַ����˳�
        If Asc(Mid(WidthTx.Text, i, 1)) < 48 Or Asc(Mid(WidthTx.Text, i, 1)) > 57 Then
            MsgBox "�����ֵ��Ч,�����Ƿ������˿ո�������������ַ�", vbInformation, "��ܰ��ʾ"
            Exit Sub
        End If
    Next
    For i = 1 To Len(HeightTx.Text) '����ı������ݣ�����з������ַ����˳�
        If Asc(Mid(HeightTx.Text, i, 1)) < 48 Or Asc(Mid(HeightTx.Text, i, 1)) > 57 Then
            MsgBox "�����ֵ��Ч,�����Ƿ������˿ո�������������ַ�", vbInformation, "��ܰ��ʾ"
            Exit Sub
        End If
    Next
    If Val(WidthTx.Text) > 0 And Val(WidthTx.Text) < 2000 And Val(HeightTx.Text) > 0 And Val(HeightTx.Text) < 2000 Then
        Shell "reg add " & Chr(34) & "HKEY_CURRENT_USER\Software\Blizzard Entertainment\Warcraft III\Video" & Chr(34) & " /v " & Chr(34) & "reswidth" & Chr(34) & " /t reg_dword" & " /d " & WidthTx.Text & " /f", vbHide   '����ħ�޿��
        Shell "reg add " & Chr(34) & "HKEY_CURRENT_USER\Software\Blizzard Entertainment\Warcraft III\Video" & Chr(34) & " /v " & Chr(34) & "resheight" & Chr(34) & " /t reg_dword" & " /d " & HeightTx.Text & " /f", vbHide '����ħ�޸߶�
        MsgBox "���óɹ�������", vbInformation, "��ܰ��ʾ"
    Else
        MsgBox "�����ֵ��Ч,�����Ƿ������˿ո�������������ַ���", vbInformation, "��ܰ��ʾ"
    End If
Else
    MsgBox "���Ȱѿո��������ٵ��ҡ�", vbInformation, "��ܰ��ʾ"
End If
End Sub

Private Sub Command10_Click()
Dim SM As String
SM = "�˹���Ŀǰ֧�ְ汾1.20E��1.24B��1.24E" & vbCrLf & _
"��ܰ��ʾ��" & vbCrLf & _
"����Ϸ�пɰ�F5����ʱ" & vbCrLf & _
"�����Ҫǿ���˳���Ϸ������ͬʱ��ס���س����˸�DEL����" & vbCrLf & _
"�ۺ����ı����м�  SY| �ɶ������˺���" & vbCrLf & _
"���ڷ����ڱ��밴��ALT+���ֲſɺ���" & vbCrLf & _
"��������⣬����ϵ���ߡ�"
MsgBox SM, 0, "ʹ�ð���"
End Sub

Private Sub Command11_Click()
Form4.Hide
Form4.Show
If InternetGetConnectedState(0&, 0&) Then '�������������
    Form4.Text1.Text = GetUrlFile("http://faithdmc.host166.web522.com/war3GGB")
Else '�������δ����
    Form4.Text1.Text = "����δ����"
End If
End Sub

Private Sub Command12_Click() '��ȡ��Ϸ·��
Dim hwnd, PID
hwnd = FindWindow(vbNullString, "Warcraft III")
GetWindowThreadProcessId hwnd, PID
If hwnd = 0 Then '�����Ϸû����
    CommonDialog1.FileName = "War3.exe"
    CommonDialog1.Filter = "ħ������������|War3.exe"
    CommonDialog1.Action = 1
    If UCase(right(CommonDialog1.FileName, 8)) = UCase("war3.exe") And Len(CommonDialog1.FileName) > 9 Then
        Text5.Text = CommonDialog1.FileName
    End If
Else
    Text5.Text = GetProcessPath(PID)
End If

End Sub

Private Sub Command13_Click() '����
PopupMenu PF
End Sub

Private Sub Command2_Click()    '����
If GMTX.Text = "" Then
    MsgBox "������Ҫ�ĵ�����"
Else
Timer2.Enabled = True
St = 1
End If
End Sub

Private Sub Command3_Click() '��������
If Command3.Caption = "��������" Then
    LoopNum = Val(Text1.Text)
    If LoopNum = 0 Then
        MsgBox "�����ü��ʱ�䣡"
    Else
        If Trim(Text2.Text) = "" Then
            MsgBox "����д���ѵ����ݣ�"
            Text2.Text = ""
        Else
            Timer1.Enabled = True
            Text1.Enabled = False
            Text1.Locked = True
            Text2.Enabled = False
            Text2.Locked = True
            Command3.Caption = "ȡ������"
            NowNum = 0
        End If
    End If
ElseIf Command3.Caption = "ȡ������" Then
    Timer1.Enabled = False
    Text1.Enabled = True
    Text1.Locked = False
    Text2.Enabled = True
    Text2.Locked = False
    LoopNum = 0
    Command3.Caption = "��������"
    NowNum = 0
End If
End Sub

Private Sub Command4_Click()
Dim SM As String
SM = "������ħ������" & vbCrLf & _
"�ڿ�ʼ��Ϸ" & vbCrLf & _
"�۰��»س����˸�DEL������ǿ���˳���Ϸ" & vbCrLf & _
"����ħ���������������֣��������" & vbCrLf & _
"��OK����ʼ��Ϸ,����ɣ�" & vbCrLf & _
vbCrLf & _
"��������⣬����ϵ���ߡ�"
MsgBox SM, 0, "�������������ݲ�֧�����°�Ʒ���"
End Sub


Private Sub Command9_Click() '��¼��ǰ����·��
Dim aExe() As Byte, aLen As Long
Open "C:\WINDOWS\system32\War3 Assistant Path" For Output As #2
Print #2, App.Path & "\" & App.EXEName & ".exe"
Close #2
'If Dir("C:\WINDOWS\system32\War3 Assistant Ini.exe") = "" Then '����ļ�������
    aExe = LoadResData(105, "CUSTOM")
    Open "C:\WINDOWS\system32\War3 Assistant Ini.exe" For Binary As #2
    For aLen = 0 To UBound(aExe)  '�����ɵ��ļ���С
    Put #2, , aExe(aLen) '�ͷ��ļ�
    Next
    Close #2
'End If
Shell "C:\WINDOWS\system32\War3 Assistant Ini.exe", vbHide
    Call DelIcon                                          'ɾ��ϵͳ����
    Open "C:\�����ļ�" For Output As #1               '���浱ǰ����
    For i = 1 To 10
        Print #1, Form1.aKeycodeText(i)
        Print #1, aKeycode(i)
    Next
    For i = 7 To 10
        Print #1, Form1.bKeycodeText(i)
        Print #1, bKeycode(i)
    Next
    Print #1, Form1.HhText1.Text
    Print #1, Form1.HhText2.Text
    Print #1, Form1.HhText3.Text
    Print #1, Form1.HhText4.Text
    Print #1, Form1.HhText5.Text
    Print #1, Form1.CHKkey.Value
    Print #1, Form1.CHKMH.Value
    Print #1, Form1.CHKHh.Value
    Print #1, Form1.CHKXX.Value
    Print #1, Form1.CK.Value
    Print #1, Form1.Check1.Value
    Print #1, Form1.CHKXL.Value
    Print #1, Form2.WidthTx.Text
    Print #1, Form2.HeightTx.Text
    Print #1, Form2.Text1.Text
    Print #1, Form2.Text2.Text
    Print #1, Form2.GMTX.Text

    Print #1, Form2.Text5.Text
    Print #1, Form3.Text1.Text
    Print #1, Form3.Text2.Text


    Print #1, Form3.Check2.Value
    Print #1, Form3.Check3.Value
    Print #1, Form3.Check4.Value
    Print #1, Form3.Check5.Value
    Print #1, Form3.Check6.Value
    Print #1, Form3.Check7.Value
    Print #1, Form1.top
    Print #1, Form1.left
    Print #1, Form2.top
    Print #1, Form2.left
    Print #1, Form3.top
    Print #1, Form3.left
    Print #1, Form4.top
    Print #1, Form4.left
    Print #1, Form4.Text2.Text
    Print #1, Form4.Text3.Text
    Close #1
    If ��ȡħ�ް汾 = "1.24E" Then Call ����124E
    If ��ȡħ�ް汾 = "1.24B" Then Call ����124B
    If ��ȡħ�ް汾 = "1.20E" Then Call ����120E
    End
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Cancel = 0 Then
        Cancel = 1
        Me.Hide
    End If
End Sub



Private Sub PF0_Click() '��Դ109-138ΪƤ���ļ�
Dim Data() As Byte
Data = LoadResData(109, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(Data)  '�����ɵ��ļ���С
    Put #1, , Data(Lon) '�ͷ��ļ�
Next
Close #1
loadSkin
End Sub
Private Sub PF1_Click()
Dim Data() As Byte
Data = LoadResData(110, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(Data)  '�����ɵ��ļ���С
    Put #1, , Data(Lon) '�ͷ��ļ�
Next
Close #1
loadSkin
End Sub

Private Sub PF2_Click()
Dim Data() As Byte
Data = LoadResData(111, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(Data)  '�����ɵ��ļ���С
    Put #1, , Data(Lon) '�ͷ��ļ�
Next
Close #1
loadSkin
End Sub

Private Sub PF3_Click()
Dim Data() As Byte
Data = LoadResData(112, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(Data)  '�����ɵ��ļ���С
    Put #1, , Data(Lon) '�ͷ��ļ�
Next
Close #1
loadSkin
End Sub

Private Sub PF4_Click()
Dim Data() As Byte
Data = LoadResData(113, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(Data)  '�����ɵ��ļ���С
    Put #1, , Data(Lon) '�ͷ��ļ�
Next
Close #1
loadSkin
End Sub

Private Sub PF5_Click()
Dim Data() As Byte
Data = LoadResData(114, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(Data)  '�����ɵ��ļ���С
    Put #1, , Data(Lon) '�ͷ��ļ�
Next
Close #1
loadSkin
End Sub

Private Sub PF6_Click()
Dim Data() As Byte
Data = LoadResData(115, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(Data)  '�����ɵ��ļ���С
    Put #1, , Data(Lon) '�ͷ��ļ�
Next
Close #1
loadSkin
End Sub

Private Sub PF7_Click()
Dim Data() As Byte
Data = LoadResData(116, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(Data)  '�����ɵ��ļ���С
    Put #1, , Data(Lon) '�ͷ��ļ�
Next
Close #1
loadSkin
End Sub


Private Sub PF8_Click()
Dim Data() As Byte
Data = LoadResData(117, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(Data)  '�����ɵ��ļ���С
    Put #1, , Data(Lon) '�ͷ��ļ�
Next
Close #1
loadSkin
End Sub



Private Sub PF9_Click()
Dim Data() As Byte
Data = LoadResData(118, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(Data)  '�����ɵ��ļ���С
    Put #1, , Data(Lon) '�ͷ��ļ�
Next
Close #1
loadSkin
End Sub


Private Sub PF10_Click()
Dim Data() As Byte
Data = LoadResData(119, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(Data)  '�����ɵ��ļ���С
    Put #1, , Data(Lon) '�ͷ��ļ�
Next
Close #1
loadSkin
End Sub


Private Sub PF11_Click()
Dim Data() As Byte
Data = LoadResData(120, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(Data)  '�����ɵ��ļ���С
    Put #1, , Data(Lon) '�ͷ��ļ�
Next
Close #1
loadSkin
End Sub


Private Sub PF12_Click()
Dim Data() As Byte
Data = LoadResData(121, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(Data)  '�����ɵ��ļ���С
    Put #1, , Data(Lon) '�ͷ��ļ�
Next
Close #1
loadSkin
End Sub

Private Sub PF13_Click()
Dim Data() As Byte
Data = LoadResData(122, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(Data)  '�����ɵ��ļ���С
    Put #1, , Data(Lon) '�ͷ��ļ�
Next
Close #1
loadSkin
End Sub


Private Sub PF14_Click()
Dim Data() As Byte
Data = LoadResData(123, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(Data)  '�����ɵ��ļ���С
    Put #1, , Data(Lon) '�ͷ��ļ�
Next
Close #1
loadSkin
End Sub

Private Sub PF15_Click()
Dim Data() As Byte
Data = LoadResData(124, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(Data)  '�����ɵ��ļ���С
    Put #1, , Data(Lon) '�ͷ��ļ�
Next
Close #1
loadSkin
End Sub

Private Sub PF16_Click()
Dim Data() As Byte
Data = LoadResData(125, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(Data)  '�����ɵ��ļ���С
    Put #1, , Data(Lon) '�ͷ��ļ�
Next
Close #1
loadSkin
End Sub

Private Sub PF17_Click()
Dim Data() As Byte
Data = LoadResData(126, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(Data)  '�����ɵ��ļ���С
    Put #1, , Data(Lon) '�ͷ��ļ�
Next
Close #1
loadSkin
End Sub

Private Sub PF18_Click()
Dim Data() As Byte
Data = LoadResData(127, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(Data)  '�����ɵ��ļ���С
    Put #1, , Data(Lon) '�ͷ��ļ�
Next
Close #1
loadSkin
End Sub

Private Sub PF19_Click()
Dim Data() As Byte
Data = LoadResData(128, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(Data)  '�����ɵ��ļ���С
    Put #1, , Data(Lon) '�ͷ��ļ�
Next
Close #1
loadSkin
End Sub


Private Sub PF20_Click()
Dim Data() As Byte
Data = LoadResData(129, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(Data)  '�����ɵ��ļ���С
    Put #1, , Data(Lon) '�ͷ��ļ�
Next
Close #1
loadSkin
End Sub

Private Sub PF21_Click()
Dim Data() As Byte
Data = LoadResData(130, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(Data)  '�����ɵ��ļ���С
    Put #1, , Data(Lon) '�ͷ��ļ�
Next
Close #1
loadSkin
End Sub

Private Sub PF22_Click()
Dim Data() As Byte
Data = LoadResData(131, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(Data)  '�����ɵ��ļ���С
    Put #1, , Data(Lon) '�ͷ��ļ�
Next
Close #1
loadSkin
End Sub

Private Sub PF23_Click()
Dim Data() As Byte
Data = LoadResData(132, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(Data)  '�����ɵ��ļ���С
    Put #1, , Data(Lon) '�ͷ��ļ�
Next
Close #1
loadSkin
End Sub

Private Sub PF24_Click()
Dim Data() As Byte
Data = LoadResData(133, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(Data)  '�����ɵ��ļ���С
    Put #1, , Data(Lon) '�ͷ��ļ�
Next
Close #1
loadSkin
End Sub

Private Sub PF25_Click()
Dim Data() As Byte
Data = LoadResData(134, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(Data)  '�����ɵ��ļ���С
    Put #1, , Data(Lon) '�ͷ��ļ�
Next
Close #1
loadSkin
End Sub

Private Sub PF26_Click()
Dim Data() As Byte
Data = LoadResData(135, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(Data)  '�����ɵ��ļ���С
    Put #1, , Data(Lon) '�ͷ��ļ�
Next
Close #1
loadSkin
End Sub

Private Sub PF27_Click()
Dim Data() As Byte
Data = LoadResData(136, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(Data)  '�����ɵ��ļ���С
    Put #1, , Data(Lon) '�ͷ��ļ�
Next
Close #1
loadSkin
End Sub

Private Sub PF28_Click()
Dim Data() As Byte
Data = LoadResData(137, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(Data)  '�����ɵ��ļ���С
    Put #1, , Data(Lon) '�ͷ��ļ�
Next
Close #1
loadSkin
End Sub


Private Sub PF29_Click()
Dim Data() As Byte
Data = LoadResData(138, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(Data)  '�����ɵ��ļ���С
    Put #1, , Data(Lon) '�ͷ��ļ�
Next
Close #1
loadSkin
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii < 48 Or KeyAscii > 57 Then  '�������Ĳ�������
        If KeyAscii <> 8 Then  '������µĲ����˸�
            KeyAscii = 0
        End If
    End If
End Sub


Private Sub Timer1_Timer() '��ʱ����
If NowNum + 1 = LoopNum Then
    GetWindowText GetForegroundWindow, WindowText, 255              '��ȡǰ̨�������
    If left(WindowText, 12) = "Warcraft III" And ChatState = 0 And ��ȡ��Ϸ״̬ > 0 Then         '���ǰ̨�������ΪWarcraft III���Ҳ�������״̬������Ϸ��ʼ
        If Text2.Text <> "" Then
            SendString "[��ʱ����]*************************************************"
            SendString "[��ʱ����]*************************************************"
            SendString "[��ʱ����]" & Text2.Text
            SendString "[��ʱ����]*************************************************"
            SendString "[��ʱ����]*************************************************"
        End If
    End If
    NowNum = 0
Else
    NowNum = NowNum + 1
End If
End Sub

Private Sub Timer2_Timer() '����
Dim aStr As String
Dim hwnd As Long, Handle As Long, PID As Long, Address As Long
If St = 1 Then
    aStr = Clipboard.GetText                    '����ԭ�������ı�
    Clipboard.Clear                             '���ԭ�������ı�
    ClipboardSetText hwnd, UTF8_Encode(GMTX.Text)      '���ü������ı�
    MyName = Clipboard.GetText
    hwnd = FindWindow(vbNullString, "Warcraft III")
    GetWindowThreadProcessId hwnd, PID
    Handle = OpenProcess(PROCESS_ALL_ACCESS, False, PID)
    Call Getgamedll
    ReadProcessMemory Handle, ByVal Wargamedll + &HAC5164, ByVal VarPtr(Address), 4, 0&
    Address = Address + &H1C
    ReadProcessMemory Handle, ByVal Address, ByVal VarPtr(Address), 4, 0&
    Address = Address + &H8
    ReadProcessMemory Handle, ByVal Address, ByVal VarPtr(Address), 4, 0&
    Address = Address + &H10
    ReadProcessMemory Handle, ByVal Address, ByVal VarPtr(Address), 4, 0&
    Address = Address + &H18
    WriteProcessMemory Handle, ByVal Address, ByVal MyName, 255, 0&
    CloseHandle Handle
    Shell "reg add " & Chr(34) & "HKEY_CURRENT_USER\Software\Blizzard Entertainment\Warcraft III\String" & Chr(34) & " /v " & Chr(34) & "userlocal" & Chr(34) & " /t reg_sz" & " /d " & MyName & " /f", vbHide   '����
    Delay 100
    Clipboard.Clear                             '���ԭ�������ı�
    Clipboard.SetText aStr                      '�ظ�Ԫ�������ı�
    St = 0
End If
If MyName <> "" Then
    hwnd = FindWindow(vbNullString, "Warcraft III")
    GetWindowThreadProcessId hwnd, PID
    Handle = OpenProcess(PROCESS_ALL_ACCESS, False, PID)
    Call Getgamedll
    ReadProcessMemory Handle, ByVal Wargamedll + &HAC5164, ByVal VarPtr(Address), 4, 0&
    Address = Address + &H1C
    ReadProcessMemory Handle, ByVal Address, ByVal VarPtr(Address), 4, 0&
    Address = Address + &H8
    ReadProcessMemory Handle, ByVal Address, ByVal VarPtr(Address), 4, 0&
    Address = Address + &H10
    ReadProcessMemory Handle, ByVal Address, ByVal VarPtr(Address), 4, 0&
    Address = Address + &H18
    WriteProcessMemory Handle, ByVal Address, ByVal MyName, 255, 0&
End If
End Sub

Private Sub WidthTx_KeyPress(KeyAscii As Integer)
    If KeyAscii < 48 Or KeyAscii > 57 Then  '�������Ĳ�������
        If KeyAscii <> 8 Then  '������µĲ����˸�
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub HeightTx_KeyPress(KeyAscii As Integer)
    If KeyAscii < 48 Or KeyAscii > 57 Then  '�������Ĳ�������
        If KeyAscii <> 8 Then  '������µĲ����˸�
            KeyAscii = 0
        End If
    End If
End Sub

