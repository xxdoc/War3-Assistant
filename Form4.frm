VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��ϵ����"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6405
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   6405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.TextBox Text3 
      Height          =   1455
      Left            =   2760
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "Form4.frx":0000
      ToolTipText     =   "���뽨���BUG"
      Top             =   120
      Width           =   3495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ˢ��"
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      ToolTipText     =   "ˢ�¹�����"
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ȡ��"
      Height          =   375
      Left            =   5160
      TabIndex        =   5
      ToolTipText     =   "ȡ��"
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�ύ"
      Height          =   375
      Left            =   3960
      TabIndex        =   4
      ToolTipText     =   "�ύ�����BUG"
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   3960
      TabIndex        =   2
      Text            =   "Text2"
      ToolTipText     =   "���������������Ա���ϵ��"
      Top             =   1800
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   2655
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "Form4.frx":0006
      ToolTipText     =   "������˵..."
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "http://faithdmc.web-108.com/"
      ForeColor       =   &H80000011&
      Height          =   180
      Left            =   3360
      MouseIcon       =   "Form4.frx":000C
      MousePointer    =   99  'Custom
      TabIndex        =   7
      ToolTipText     =   "��������վ"
      Top             =   2160
      Width           =   2520
   End
   Begin VB.Label Label1 
      Caption         =   "�������䣺"
      Height          =   255
      Left            =   3000
      TabIndex        =   6
      ToolTipText     =   "���������������Ա���ϵ��"
      Top             =   1800
      Width           =   975
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Sub Command1_Click()
If InternetGetConnectedState(0&, 0&) = 0 Then MsgBox "����δ����", vbInformation: Exit Sub             '�������δ����
If Len(Text3.Text) < 1 Then MsgBox "����д�����BUG": Exit Sub
If (right(Text2.Text, 4) = ".com" Or right(Text2.Text, 4) = ".Com" Or right(Text2.Text, 4) = ".COm" Or right(Text2.Text, 4) = ".COM" Or right(Text2.Text, 4) = ".cOm" Or right(Text2.Text, 4) = ".cOM" Or right(Text2.Text, 4) = ".coM") And Len(Text2.Text) > 7 Then
    a = MsgBox("�����Ҫ�ύ?ȷ��?", vbYesNo, "��ʾ:")
    If a = vbNo Then Exit Sub
Else
    a = MsgBox("����δ��ȷ��д,ȷ���ύ?", vbYesNo, "��ʾ:")
    If a = vbNo Then Exit Sub
End If
Command1.Enabled = False
Command1.Caption = "�����ύ"

    Dim NameS As String
    Dim Email As Object
    
    NameS = "http://schemas.microsoft.com/cdo/configuration/"
    
    Set Email = CreateObject("CDO.Message")
    Email.From = "********"  '�����˵������ַ
    Email.To = "********"   ' �ռ��˵������ַ
    Email.Subject = "War3 Assistant BUG ����" & Text2.Text   '�ʼ�����
    Email.Textbody = Text3.Text  '�ʼ�����
    Email.Configuration.Fields.Item(NameS & "sendusing") = 2
    Email.Configuration.Fields.Item(NameS & "smtpserver") = "smtp.163.com" '�ʼ����ͷ�����"
    Email.Configuration.Fields.Item(NameS & "smtpserverport") = 25 '�ʼ����ͷ��������ŵĶ˿ں�
    Email.Configuration.Fields.Item(NameS & "smtpauthenticate") = 1
    Email.Configuration.Fields.Item(NameS & "sendusername") = "********" '�����˵��ʺ�
    Email.Configuration.Fields.Item(NameS & "sendpassword") = "********" '�����˵�����
    Email.Configuration.Fields.Update
    Email.send
Command1.Enabled = True
Command1.Caption = "�ύ"
End Sub

Private Sub Command2_Click()
Form4.Hide
End Sub

Private Sub Command3_Click() '��ʾ����
If InternetGetConnectedState(0&, 0&) Then '�������������
    Text1.Text = GetUrlFile("http://faithdmc.host166.web522.com/war3GGB")
Else '�������δ����
    Form4.Text1.Text = "����δ����"
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
If Cancel = 0 Then
    Cancel = 1                            '�������Ĺرգ��򽫴�������
    Form4.Hide
End If
End Sub

Private Sub Label2_Click()
ShellExecute 0, vbNullString, "http://faithdmc.web-108.com/", vbNullString, vbNullString, 1
End Sub
