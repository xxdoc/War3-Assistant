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
      Name            =   "宋体"
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
      Caption         =   "更换皮肤"
      BeginProperty Font 
         Name            =   "宋体"
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
      ToolTipText     =   "如果换肤后窗体变形重启程序即可"
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
      Caption         =   "获取"
      BeginProperty Font 
         Name            =   "宋体"
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
      ToolTipText     =   "如果游戏运行时点击可自动获取"
      Top             =   2640
      Width           =   735
   End
   Begin VB.TextBox Text5 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
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
      Caption         =   "联系作者"
      BeginProperty Font 
         Name            =   "宋体"
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
      ToolTipText     =   "联系作者"
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton Command10 
      Caption         =   "使用帮助"
      BeginProperty Font 
         Name            =   "宋体"
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
      ToolTipText     =   "查看帮助"
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton Command9 
      Caption         =   "保存默认配置"
      BeginProperty Font 
         Name            =   "宋体"
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
      ToolTipText     =   "保存后该配置信息在其他电脑依然存在"
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "帮助"
      BeginProperty Font 
         Name            =   "宋体"
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
         Name            =   "宋体"
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
      ToolTipText     =   "提醒的内容"
      Top             =   2040
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
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
      ToolTipText     =   "间隔时间"
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "设置提醒"
      BeginProperty Font 
         Name            =   "宋体"
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
         Name            =   "宋体"
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
      ToolTipText     =   "名字"
      Top             =   840
      Width           =   2745
   End
   Begin VB.CommandButton Command2 
      Caption         =   "改名"
      BeginProperty Font 
         Name            =   "宋体"
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
      ToolTipText     =   "改名"
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "设置魔兽分辨率"
      BeginProperty Font 
         Name            =   "宋体"
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
      ToolTipText     =   "如果魔兽全屏时无法全屏可用此功能修复，其值与屏幕分辨率相同，重启魔兽后生效"
      Top             =   120
      Width           =   1395
   End
   Begin VB.TextBox HeightTx 
      BeginProperty Font 
         Name            =   "宋体"
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
         Name            =   "宋体"
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
      Caption         =   "游戏路径:"
      BeginProperty Font 
         Name            =   "宋体"
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
      Caption         =   "分钟"
      BeginProperty Font 
         Name            =   "宋体"
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
      Caption         =   "内容："
      BeginProperty Font 
         Name            =   "宋体"
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
      Caption         =   "间隔："
      BeginProperty Font 
         Name            =   "宋体"
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
         Name            =   "宋体"
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
      Caption         =   "更换皮肤"
      Visible         =   0   'False
      Begin VB.Menu PF0 
         Caption         =   "默认皮肤"
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
         Caption         =   "QQ影音"
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
         Caption         =   "积木"
      End
      Begin VB.Menu PF29 
         Caption         =   "炫绿"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef dwFlags As Long, ByVal dwReserved As Long) As Long
Dim LoopNum As Long '用来保存闹钟的Timer循环次数
Dim NowNum As Long '用来保存当前闹钟的Timer循环次数
Dim St As Integer '保存改名是是否已经获得名字的UTF8编码文本
Dim MyName As String
Private Sub Command1_Click() '设置魔兽分辨率
If WidthTx.Text <> "" And HeightTx.Text <> "" Then
    For i = 1 To Len(WidthTx.Text) '检测文本框内容，如果有非数字字符则退出
        If Asc(Mid(WidthTx.Text, i, 1)) < 48 Or Asc(Mid(WidthTx.Text, i, 1)) > 57 Then
            MsgBox "输入的值无效,请检查是否输入了空格或其他非数字字符", vbInformation, "温馨提示"
            Exit Sub
        End If
    Next
    For i = 1 To Len(HeightTx.Text) '检测文本框内容，如果有非数字字符则退出
        If Asc(Mid(HeightTx.Text, i, 1)) < 48 Or Asc(Mid(HeightTx.Text, i, 1)) > 57 Then
            MsgBox "输入的值无效,请检查是否输入了空格或其他非数字字符", vbInformation, "温馨提示"
            Exit Sub
        End If
    Next
    If Val(WidthTx.Text) > 0 And Val(WidthTx.Text) < 2000 And Val(HeightTx.Text) > 0 And Val(HeightTx.Text) < 2000 Then
        Shell "reg add " & Chr(34) & "HKEY_CURRENT_USER\Software\Blizzard Entertainment\Warcraft III\Video" & Chr(34) & " /v " & Chr(34) & "reswidth" & Chr(34) & " /t reg_dword" & " /d " & WidthTx.Text & " /f", vbHide   '设置魔兽宽度
        Shell "reg add " & Chr(34) & "HKEY_CURRENT_USER\Software\Blizzard Entertainment\Warcraft III\Video" & Chr(34) & " /v " & Chr(34) & "resheight" & Chr(34) & " /t reg_dword" & " /d " & HeightTx.Text & " /f", vbHide '设置魔兽高度
        MsgBox "设置成功！！！", vbInformation, "温馨提示"
    Else
        MsgBox "输入的值无效,请检查是否输入了空格或其他非数字字符。", vbInformation, "温馨提示"
    End If
Else
    MsgBox "请先把空格填完整再点我。", vbInformation, "温馨提示"
End If
End Sub

Private Sub Command10_Click()
Dim SM As String
SM = "此工具目前支持版本1.20E、1.24B、1.24E" & vbCrLf & _
"温馨提示：" & vbCrLf & _
"①游戏中可按F5键报时" & vbCrLf & _
"②如果要强制退出游戏，可以同时按住：回车、退格、DEL键。" & vbCrLf & _
"③喊话文本框中加  SY| 可对所有人喊话" & vbCrLf & _
"④在房间内必须按下ALT+数字才可喊话" & vbCrLf & _
"如果有问题，请联系作者。"
MsgBox SM, 0, "使用帮助"
End Sub

Private Sub Command11_Click()
Form4.Hide
Form4.Show
If InternetGetConnectedState(0&, 0&) Then '如果网络已连接
    Form4.Text1.Text = GetUrlFile("http://faithdmc.host166.web522.com/war3GGB")
Else '如果网络未连接
    Form4.Text1.Text = "网络未连接"
End If
End Sub

Private Sub Command12_Click() '获取游戏路径
Dim hwnd, PID
hwnd = FindWindow(vbNullString, "Warcraft III")
GetWindowThreadProcessId hwnd, PID
If hwnd = 0 Then '如果游戏没运行
    CommonDialog1.FileName = "War3.exe"
    CommonDialog1.Filter = "魔兽争霸主程序|War3.exe"
    CommonDialog1.Action = 1
    If UCase(right(CommonDialog1.FileName, 8)) = UCase("war3.exe") And Len(CommonDialog1.FileName) > 9 Then
        Text5.Text = CommonDialog1.FileName
    End If
Else
    Text5.Text = GetProcessPath(PID)
End If

End Sub

Private Sub Command13_Click() '换肤
PopupMenu PF
End Sub

Private Sub Command2_Click()    '改名
If GMTX.Text = "" Then
    MsgBox "请输入要改的名字"
Else
Timer2.Enabled = True
St = 1
End If
End Sub

Private Sub Command3_Click() '设置提醒
If Command3.Caption = "设置提醒" Then
    LoopNum = Val(Text1.Text)
    If LoopNum = 0 Then
        MsgBox "请设置间隔时间！"
    Else
        If Trim(Text2.Text) = "" Then
            MsgBox "请填写提醒的内容！"
            Text2.Text = ""
        Else
            Timer1.Enabled = True
            Text1.Enabled = False
            Text1.Locked = True
            Text2.Enabled = False
            Text2.Locked = True
            Command3.Caption = "取消提醒"
            NowNum = 0
        End If
    End If
ElseIf Command3.Caption = "取消提醒" Then
    Timer1.Enabled = False
    Text1.Enabled = True
    Text1.Locked = False
    Text2.Enabled = True
    Text2.Locked = False
    LoopNum = 0
    Command3.Caption = "设置提醒"
    NowNum = 0
End If
End Sub

Private Sub Command4_Click()
Dim SM As String
SM = "①运行魔兽助手" & vbCrLf & _
"②开始游戏" & vbCrLf & _
"③按下回车，退格，DEL三个键强制退出游戏" & vbCrLf & _
"④在魔兽助手上输入名字，点击改名" & vbCrLf & _
"⑤OK，开始游戏,快玩吧！" & vbCrLf & _
vbCrLf & _
"如果有问题，请联系作者。"
MsgBox SM, 0, "改名方法：（暂不支持最新版浩方）"
End Sub


Private Sub Command9_Click() '记录当前完整路径
Dim aExe() As Byte, aLen As Long
Open "C:\WINDOWS\system32\War3 Assistant Path" For Output As #2
Print #2, App.Path & "\" & App.EXEName & ".exe"
Close #2
'If Dir("C:\WINDOWS\system32\War3 Assistant Ini.exe") = "" Then '如果文件不存在
    aExe = LoadResData(105, "CUSTOM")
    Open "C:\WINDOWS\system32\War3 Assistant Ini.exe" For Binary As #2
    For aLen = 0 To UBound(aExe)  '欲生成的文件大小
    Put #2, , aExe(aLen) '释放文件
    Next
    Close #2
'End If
Shell "C:\WINDOWS\system32\War3 Assistant Ini.exe", vbHide
    Call DelIcon                                          '删除系统托盘
    Open "C:\配置文件" For Output As #1               '保存当前设置
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
    If 获取魔兽版本 = "1.24E" Then Call 不改124E
    If 获取魔兽版本 = "1.24B" Then Call 不改124B
    If 获取魔兽版本 = "1.20E" Then Call 不改120E
    End
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Cancel = 0 Then
        Cancel = 1
        Me.Hide
    End If
End Sub



Private Sub PF0_Click() '资源109-138为皮肤文件
Dim Data() As Byte
Data = LoadResData(109, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(Data)  '欲生成的文件大小
    Put #1, , Data(Lon) '释放文件
Next
Close #1
loadSkin
End Sub
Private Sub PF1_Click()
Dim Data() As Byte
Data = LoadResData(110, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(Data)  '欲生成的文件大小
    Put #1, , Data(Lon) '释放文件
Next
Close #1
loadSkin
End Sub

Private Sub PF2_Click()
Dim Data() As Byte
Data = LoadResData(111, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(Data)  '欲生成的文件大小
    Put #1, , Data(Lon) '释放文件
Next
Close #1
loadSkin
End Sub

Private Sub PF3_Click()
Dim Data() As Byte
Data = LoadResData(112, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(Data)  '欲生成的文件大小
    Put #1, , Data(Lon) '释放文件
Next
Close #1
loadSkin
End Sub

Private Sub PF4_Click()
Dim Data() As Byte
Data = LoadResData(113, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(Data)  '欲生成的文件大小
    Put #1, , Data(Lon) '释放文件
Next
Close #1
loadSkin
End Sub

Private Sub PF5_Click()
Dim Data() As Byte
Data = LoadResData(114, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(Data)  '欲生成的文件大小
    Put #1, , Data(Lon) '释放文件
Next
Close #1
loadSkin
End Sub

Private Sub PF6_Click()
Dim Data() As Byte
Data = LoadResData(115, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(Data)  '欲生成的文件大小
    Put #1, , Data(Lon) '释放文件
Next
Close #1
loadSkin
End Sub

Private Sub PF7_Click()
Dim Data() As Byte
Data = LoadResData(116, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(Data)  '欲生成的文件大小
    Put #1, , Data(Lon) '释放文件
Next
Close #1
loadSkin
End Sub


Private Sub PF8_Click()
Dim Data() As Byte
Data = LoadResData(117, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(Data)  '欲生成的文件大小
    Put #1, , Data(Lon) '释放文件
Next
Close #1
loadSkin
End Sub



Private Sub PF9_Click()
Dim Data() As Byte
Data = LoadResData(118, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(Data)  '欲生成的文件大小
    Put #1, , Data(Lon) '释放文件
Next
Close #1
loadSkin
End Sub


Private Sub PF10_Click()
Dim Data() As Byte
Data = LoadResData(119, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(Data)  '欲生成的文件大小
    Put #1, , Data(Lon) '释放文件
Next
Close #1
loadSkin
End Sub


Private Sub PF11_Click()
Dim Data() As Byte
Data = LoadResData(120, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(Data)  '欲生成的文件大小
    Put #1, , Data(Lon) '释放文件
Next
Close #1
loadSkin
End Sub


Private Sub PF12_Click()
Dim Data() As Byte
Data = LoadResData(121, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(Data)  '欲生成的文件大小
    Put #1, , Data(Lon) '释放文件
Next
Close #1
loadSkin
End Sub

Private Sub PF13_Click()
Dim Data() As Byte
Data = LoadResData(122, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(Data)  '欲生成的文件大小
    Put #1, , Data(Lon) '释放文件
Next
Close #1
loadSkin
End Sub


Private Sub PF14_Click()
Dim Data() As Byte
Data = LoadResData(123, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(Data)  '欲生成的文件大小
    Put #1, , Data(Lon) '释放文件
Next
Close #1
loadSkin
End Sub

Private Sub PF15_Click()
Dim Data() As Byte
Data = LoadResData(124, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(Data)  '欲生成的文件大小
    Put #1, , Data(Lon) '释放文件
Next
Close #1
loadSkin
End Sub

Private Sub PF16_Click()
Dim Data() As Byte
Data = LoadResData(125, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(Data)  '欲生成的文件大小
    Put #1, , Data(Lon) '释放文件
Next
Close #1
loadSkin
End Sub

Private Sub PF17_Click()
Dim Data() As Byte
Data = LoadResData(126, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(Data)  '欲生成的文件大小
    Put #1, , Data(Lon) '释放文件
Next
Close #1
loadSkin
End Sub

Private Sub PF18_Click()
Dim Data() As Byte
Data = LoadResData(127, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(Data)  '欲生成的文件大小
    Put #1, , Data(Lon) '释放文件
Next
Close #1
loadSkin
End Sub

Private Sub PF19_Click()
Dim Data() As Byte
Data = LoadResData(128, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(Data)  '欲生成的文件大小
    Put #1, , Data(Lon) '释放文件
Next
Close #1
loadSkin
End Sub


Private Sub PF20_Click()
Dim Data() As Byte
Data = LoadResData(129, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(Data)  '欲生成的文件大小
    Put #1, , Data(Lon) '释放文件
Next
Close #1
loadSkin
End Sub

Private Sub PF21_Click()
Dim Data() As Byte
Data = LoadResData(130, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(Data)  '欲生成的文件大小
    Put #1, , Data(Lon) '释放文件
Next
Close #1
loadSkin
End Sub

Private Sub PF22_Click()
Dim Data() As Byte
Data = LoadResData(131, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(Data)  '欲生成的文件大小
    Put #1, , Data(Lon) '释放文件
Next
Close #1
loadSkin
End Sub

Private Sub PF23_Click()
Dim Data() As Byte
Data = LoadResData(132, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(Data)  '欲生成的文件大小
    Put #1, , Data(Lon) '释放文件
Next
Close #1
loadSkin
End Sub

Private Sub PF24_Click()
Dim Data() As Byte
Data = LoadResData(133, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(Data)  '欲生成的文件大小
    Put #1, , Data(Lon) '释放文件
Next
Close #1
loadSkin
End Sub

Private Sub PF25_Click()
Dim Data() As Byte
Data = LoadResData(134, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(Data)  '欲生成的文件大小
    Put #1, , Data(Lon) '释放文件
Next
Close #1
loadSkin
End Sub

Private Sub PF26_Click()
Dim Data() As Byte
Data = LoadResData(135, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(Data)  '欲生成的文件大小
    Put #1, , Data(Lon) '释放文件
Next
Close #1
loadSkin
End Sub

Private Sub PF27_Click()
Dim Data() As Byte
Data = LoadResData(136, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(Data)  '欲生成的文件大小
    Put #1, , Data(Lon) '释放文件
Next
Close #1
loadSkin
End Sub

Private Sub PF28_Click()
Dim Data() As Byte
Data = LoadResData(137, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(Data)  '欲生成的文件大小
    Put #1, , Data(Lon) '释放文件
Next
Close #1
loadSkin
End Sub


Private Sub PF29_Click()
Dim Data() As Byte
Data = LoadResData(138, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(Data)  '欲生成的文件大小
    Put #1, , Data(Lon) '释放文件
Next
Close #1
loadSkin
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii < 48 Or KeyAscii > 57 Then  '如果输入的不是数字
        If KeyAscii <> 8 Then  '如果按下的不是退格
            KeyAscii = 0
        End If
    End If
End Sub


Private Sub Timer1_Timer() '定时提醒
If NowNum + 1 = LoopNum Then
    GetWindowText GetForegroundWindow, WindowText, 255              '获取前台窗体标题
    If left(WindowText, 12) = "Warcraft III" And ChatState = 0 And 获取游戏状态 > 0 Then         '如果前台窗体标题为Warcraft III并且不在聊天状态并且游戏开始
        If Text2.Text <> "" Then
            SendString "[定时提醒]*************************************************"
            SendString "[定时提醒]*************************************************"
            SendString "[定时提醒]" & Text2.Text
            SendString "[定时提醒]*************************************************"
            SendString "[定时提醒]*************************************************"
        End If
    End If
    NowNum = 0
Else
    NowNum = NowNum + 1
End If
End Sub

Private Sub Timer2_Timer() '改名
Dim aStr As String
Dim hwnd As Long, Handle As Long, PID As Long, Address As Long
If St = 1 Then
    aStr = Clipboard.GetText                    '保存原剪贴板文本
    Clipboard.Clear                             '清空原剪贴板文本
    ClipboardSetText hwnd, UTF8_Encode(GMTX.Text)      '设置剪贴板文本
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
    Shell "reg add " & Chr(34) & "HKEY_CURRENT_USER\Software\Blizzard Entertainment\Warcraft III\String" & Chr(34) & " /v " & Chr(34) & "userlocal" & Chr(34) & " /t reg_sz" & " /d " & MyName & " /f", vbHide   '改名
    Delay 100
    Clipboard.Clear                             '清空原剪贴板文本
    Clipboard.SetText aStr                      '回复元剪贴板文本
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
    If KeyAscii < 48 Or KeyAscii > 57 Then  '如果输入的不是数字
        If KeyAscii <> 8 Then  '如果按下的不是退格
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub HeightTx_KeyPress(KeyAscii As Integer)
    If KeyAscii < 48 Or KeyAscii > 57 Then  '如果输入的不是数字
        If KeyAscii <> 8 Then  '如果按下的不是退格
            KeyAscii = 0
        End If
    End If
End Sub

