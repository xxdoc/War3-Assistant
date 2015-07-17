VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "War3 Assistant"
   ClientHeight    =   4212
   ClientLeft      =   8232
   ClientTop       =   5520
   ClientWidth     =   4620
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   21.6
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
   ScaleHeight     =   4212
   ScaleWidth      =   4620
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2040
      Top             =   720
   End
   Begin VB.CommandButton Command5 
      Caption         =   "游戏加速"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   432
      Left            =   3120
      TabIndex        =   20
      ToolTipText     =   "开启后游戏中按小键盘+可加速，按小键盘-可恢复"
      Top             =   720
      Width           =   1332
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   13.8
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   444
      Left            =   240
      TabIndex        =   19
      Text            =   "64"
      Top             =   720
      Width           =   2772
   End
   Begin VB.CommandButton Command13 
      Caption         =   "更换皮肤"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   18
      ToolTipText     =   "如果换肤后窗体变形重启程序即可"
      Top             =   3720
      Width           =   1332
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3960
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command12 
      Caption         =   "获取"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3840
      TabIndex        =   17
      ToolTipText     =   "如果游戏运行时点击可自动获取"
      Top             =   3120
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
      TabIndex        =   16
      Top             =   3120
      Width           =   2655
   End
   Begin VB.CommandButton Command11 
      Caption         =   "获取最新版本"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.8
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
      Top             =   3720
      Width           =   1452
   End
   Begin VB.CommandButton Command10 
      Caption         =   "使用帮助"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   11
      ToolTipText     =   "查看帮助"
      Top             =   3720
      Width           =   1332
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
      Top             =   1320
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   3360
      Top             =   4800
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.4
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
      Top             =   2520
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.4
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
      Top             =   1920
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
      Top             =   1920
      Width           =   1575
   End
   Begin VB.TextBox GMTX 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.4
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
      Top             =   1320
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
      Top             =   1320
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
         Size            =   14.4
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
         Size            =   14.4
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
   Begin VB.Line Line7 
      X1              =   0
      X2              =   5760
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line6 
      X1              =   0
      X2              =   5040
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "游戏路径:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   216
      Left            =   120
      TabIndex        =   15
      Top             =   3240
      Width           =   948
   End
   Begin VB.Line Line5 
      Visible         =   0   'False
      X1              =   120
      X2              =   0
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line Line4 
      X1              =   0
      X2              =   4560
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   4560
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   4680
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line1 
      Visible         =   0   'False
      X1              =   120
      X2              =   4560
      Y1              =   3000
      Y2              =   3000
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
      TabIndex        =   14
      Top             =   2040
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
      TabIndex        =   13
      Top             =   2520
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
      TabIndex        =   12
      Top             =   2040
      Width           =   720
   End
   Begin VB.Label Label3 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.4
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
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Dim LoopNum As Long '用来保存闹钟的Timer循环次数
Dim NowNum As Long '用来保存当前闹钟的Timer循环次数
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
    ShellExecute 0, vbNullString, "https://raw.githubusercontent.com/ranulldd/War3-Assistant/master/War3%20Assistant.exe", vbNullString, vbNullString, vbNormalFocus
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
Dim hwnd As Long, Handle As Long, PID As Long, Address As Long
    If GMTX.Text = "" Then
        MsgBox "请输入要改的名字"
    Else
        hwnd = FindWindow(vbNullString, "浩方电竞平台 - 6.0.0.0521(RC7)")
        If hwnd = 0 Then MsgBox "请先运行浩方电竞平台": Exit Sub
        GetWindowThreadProcessId hwnd, PID
        Handle = OpenProcess(PROCESS_ALL_ACCESS, False, PID)
        Call GetHFGameShelldll
        
        WriteProcessMemory Handle, ByVal HFGameShelldll + &H5A7A0, ByVal GMTX.Text, 56, 0& '先写入昵称
        
        Dim data(24) As Byte
        data(0) = &HF3
        data(1) = &HA5
        data(2) = &H8B
        data(3) = &H83
        data(4) = &H58
        data(5) = &H5
        data(6) = &H0
        data(7) = &H0
        data(8) = &HB9
        data(9) = &HE
        data(10) = &H0
        data(11) = &H0
        data(12) = &H0
        data(13) = &HBE
        data(14) = &HA0
        data(15) = &HA7
        data(16) = &H5
        data(17) = &H10
        data(18) = &HF3
        data(19) = &HA5
        data(20) = &HE9
        data(21) = &HD4
        data(22) = &HF0
        data(23) = &HFC
        data(24) = &HFF
        WriteProcessMemory Handle, ByVal HFGameShelldll + &H5A6D0, data(0), 25, 0&
        
        data(0) = &HB9
        data(1) = &H8
        data(2) = &H0
        data(3) = &H0
        data(4) = &H0
        data(5) = &H8D
        data(6) = &H74
        data(7) = &H24
        data(8) = &H20
        data(9) = &HE9
        data(10) = &H16
        data(11) = &HF
        data(12) = &H3
        data(13) = &H0
        data(14) = &H90
        data(15) = &H90
        data(16) = &H90
        WriteProcessMemory Handle, ByVal HFGameShelldll + &H297AC, data(0), 17, 0&
        
        
        '-------don't need reenter the room------
        'ReadProcessMemory ByVal Handle, ByVal &H5D62CC, Address, 4, 0&
        'WriteProcessMemory Handle, ByVal Address, ByVal GMTX.Text, 56, 0& '写入昵称
                
        CloseHandle Handle
        
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
"②打开浩方并登陆" & vbCrLf & _
"③输入新昵称，点击更改" & vbCrLf & _
"④OK，开始游戏,快玩吧！" & vbCrLf & _
vbCrLf & _
"如要重新改名，请更改后重新进入房间。"
MsgBox SM, 0, "改名方法：（仅支持浩方电竞平台 - 6.0.0.0521(RC7)）"
End Sub

Private Sub Command5_Click() '游戏加速
Dim hwnd As Long, Handle As Long, PID As Long
If Val(Text3.Text) < 1 Or Val(Text3.Text) > 120 Then Text3.Text = 20
hwnd = FindWindow(vbNullString, "Warcraft III")
If hwnd = 0 Then MsgBox "请先运行游戏": Exit Sub
GetWindowThreadProcessId hwnd, PID
Handle = OpenProcess(PROCESS_ALL_ACCESS, False, PID)
Call GetKernelBasedll

Dim Tmp As Long
Dim Base(1) As Byte

Tmp = KernelBasedll
Tmp = Tmp / &H10000
CopyMemory Base(0), Tmp, 2

Dim data(52) As Byte
data(0) = &HB9
data(1) = &H0
data(2) = &H0
data(3) = &H0
data(4) = &H0
data(5) = &H3B
data(6) = &HD
data(7) = &HF0
data(8) = &HF
data(9) = Base(0)
data(10) = Base(1)
data(11) = &H75
data(12) = &HB
data(13) = &HA3
data(14) = &HF0
data(15) = &HF
data(16) = Base(0)
data(17) = Base(1)
data(18) = &HA3
data(19) = &HF8
data(20) = &HF
data(21) = Base(0)
data(22) = Base(1)
data(23) = &HC3
data(24) = &H8B
data(25) = &HC8
data(26) = &H2B
data(27) = &HD
data(28) = &HF8
data(29) = &HF
data(30) = Base(0)
data(31) = Base(1)
data(32) = &H6B
data(33) = &HC9
data(34) = &H5
data(35) = &HA3
data(36) = &HF8
data(37) = &HF
data(38) = Base(0)
data(39) = Base(1)
data(40) = &HA1
data(41) = &HF0
data(42) = &HF
data(43) = Base(0)
data(44) = Base(1)
data(45) = &H3
data(46) = &HC1
data(47) = &HA3
data(48) = &HF0
data(49) = &HF
data(50) = Base(0)
data(51) = Base(1)
data(52) = &HC3
WriteProcessMemory Handle, ByVal KernelBasedll + &H770, data(0), 53, 0& '更改GetTickCount的返回值

data(0) = 1
WriteProcessMemory Handle, ByVal KernelBasedll + &H792, data(0), 1, 0& '更改速率

data(0) = &HE9
data(1) = &H7C
data(2) = &H77
data(3) = &HFF
data(4) = &HFF
WriteProcessMemory Handle, ByVal KernelBasedll + &H8FEF&, data(0), 5, 0& '跳转过去

CloseHandle Handle

Timer2.Enabled = True

End Sub



Private Sub Form_Unload(Cancel As Integer)
    If Cancel = 0 Then
        Cancel = 1
        Me.Hide
    End If
End Sub



Private Sub PF0_Click() '资源109-138为皮肤文件
Dim data() As Byte
data = LoadResData(109, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(data)  '欲生成的文件大小
    Put #1, , data(Lon) '释放文件
Next
Close #1
loadSkin
End Sub
Private Sub PF1_Click()
Dim data() As Byte
data = LoadResData(110, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(data)  '欲生成的文件大小
    Put #1, , data(Lon) '释放文件
Next
Close #1
loadSkin
End Sub

Private Sub PF2_Click()
Dim data() As Byte
data = LoadResData(111, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(data)  '欲生成的文件大小
    Put #1, , data(Lon) '释放文件
Next
Close #1
loadSkin
End Sub

Private Sub PF3_Click()
Dim data() As Byte
data = LoadResData(112, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(data)  '欲生成的文件大小
    Put #1, , data(Lon) '释放文件
Next
Close #1
loadSkin
End Sub

Private Sub PF4_Click()
Dim data() As Byte
data = LoadResData(113, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(data)  '欲生成的文件大小
    Put #1, , data(Lon) '释放文件
Next
Close #1
loadSkin
End Sub

Private Sub PF5_Click()
Dim data() As Byte
data = LoadResData(114, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(data)  '欲生成的文件大小
    Put #1, , data(Lon) '释放文件
Next
Close #1
loadSkin
End Sub

Private Sub PF6_Click()
Dim data() As Byte
data = LoadResData(115, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(data)  '欲生成的文件大小
    Put #1, , data(Lon) '释放文件
Next
Close #1
loadSkin
End Sub

Private Sub PF7_Click()
Dim data() As Byte
data = LoadResData(116, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(data)  '欲生成的文件大小
    Put #1, , data(Lon) '释放文件
Next
Close #1
loadSkin
End Sub


Private Sub PF8_Click()
Dim data() As Byte
data = LoadResData(117, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(data)  '欲生成的文件大小
    Put #1, , data(Lon) '释放文件
Next
Close #1
loadSkin
End Sub



Private Sub PF9_Click()
Dim data() As Byte
data = LoadResData(118, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(data)  '欲生成的文件大小
    Put #1, , data(Lon) '释放文件
Next
Close #1
loadSkin
End Sub


Private Sub PF10_Click()
Dim data() As Byte
data = LoadResData(119, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(data)  '欲生成的文件大小
    Put #1, , data(Lon) '释放文件
Next
Close #1
loadSkin
End Sub


Private Sub PF11_Click()
Dim data() As Byte
data = LoadResData(120, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(data)  '欲生成的文件大小
    Put #1, , data(Lon) '释放文件
Next
Close #1
loadSkin
End Sub


Private Sub PF12_Click()
Dim data() As Byte
data = LoadResData(121, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(data)  '欲生成的文件大小
    Put #1, , data(Lon) '释放文件
Next
Close #1
loadSkin
End Sub

Private Sub PF13_Click()
Dim data() As Byte
data = LoadResData(122, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(data)  '欲生成的文件大小
    Put #1, , data(Lon) '释放文件
Next
Close #1
loadSkin
End Sub


Private Sub PF14_Click()
Dim data() As Byte
data = LoadResData(123, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(data)  '欲生成的文件大小
    Put #1, , data(Lon) '释放文件
Next
Close #1
loadSkin
End Sub

Private Sub PF15_Click()
Dim data() As Byte
data = LoadResData(124, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(data)  '欲生成的文件大小
    Put #1, , data(Lon) '释放文件
Next
Close #1
loadSkin
End Sub

Private Sub PF16_Click()
Dim data() As Byte
data = LoadResData(125, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(data)  '欲生成的文件大小
    Put #1, , data(Lon) '释放文件
Next
Close #1
loadSkin
End Sub

Private Sub PF17_Click()
Dim data() As Byte
data = LoadResData(126, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(data)  '欲生成的文件大小
    Put #1, , data(Lon) '释放文件
Next
Close #1
loadSkin
End Sub

Private Sub PF18_Click()
Dim data() As Byte
data = LoadResData(127, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(data)  '欲生成的文件大小
    Put #1, , data(Lon) '释放文件
Next
Close #1
loadSkin
End Sub

Private Sub PF19_Click()
Dim data() As Byte
data = LoadResData(128, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(data)  '欲生成的文件大小
    Put #1, , data(Lon) '释放文件
Next
Close #1
loadSkin
End Sub


Private Sub PF20_Click()
Dim data() As Byte
data = LoadResData(129, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(data)  '欲生成的文件大小
    Put #1, , data(Lon) '释放文件
Next
Close #1
loadSkin
End Sub

Private Sub PF21_Click()
Dim data() As Byte
data = LoadResData(130, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(data)  '欲生成的文件大小
    Put #1, , data(Lon) '释放文件
Next
Close #1
loadSkin
End Sub

Private Sub PF22_Click()
Dim data() As Byte
data = LoadResData(131, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(data)  '欲生成的文件大小
    Put #1, , data(Lon) '释放文件
Next
Close #1
loadSkin
End Sub

Private Sub PF23_Click()
Dim data() As Byte
data = LoadResData(132, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(data)  '欲生成的文件大小
    Put #1, , data(Lon) '释放文件
Next
Close #1
loadSkin
End Sub

Private Sub PF24_Click()
Dim data() As Byte
data = LoadResData(133, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(data)  '欲生成的文件大小
    Put #1, , data(Lon) '释放文件
Next
Close #1
loadSkin
End Sub

Private Sub PF25_Click()
Dim data() As Byte
data = LoadResData(134, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(data)  '欲生成的文件大小
    Put #1, , data(Lon) '释放文件
Next
Close #1
loadSkin
End Sub

Private Sub PF26_Click()
Dim data() As Byte
data = LoadResData(135, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(data)  '欲生成的文件大小
    Put #1, , data(Lon) '释放文件
Next
Close #1
loadSkin
End Sub

Private Sub PF27_Click()
Dim data() As Byte
data = LoadResData(136, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(data)  '欲生成的文件大小
    Put #1, , data(Lon) '释放文件
Next
Close #1
loadSkin
End Sub

Private Sub PF28_Click()
Dim data() As Byte
data = LoadResData(137, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(data)  '欲生成的文件大小
    Put #1, , data(Lon) '释放文件
Next
Close #1
loadSkin
End Sub


Private Sub PF29_Click()
Dim data() As Byte
data = LoadResData(138, "CUSTOM")
Open "C:\WINDOWS\system32\Pifu.she" For Binary As #1
For Lon = 0 To UBound(data)  '欲生成的文件大小
    Put #1, , data(Lon) '释放文件
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


Private Sub Text3_KeyPress(KeyAscii As Integer)
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


Private Sub Timer2_Timer() '加速
Dim hwnd As Long, Handle As Long, PID As Long
Dim data(1) As Byte
If GetAsyncKeyState(107) < 0 Then  '按下+

    If Val(Text3.Text) < 1 Or Val(Text3.Text) > 120 Then Text3.Text = 20
    hwnd = FindWindow(vbNullString, "Warcraft III")
    If hwnd = 0 Then MsgBox "请先运行游戏": Exit Sub
    GetWindowThreadProcessId hwnd, PID
    Handle = OpenProcess(PROCESS_ALL_ACCESS, False, PID)
    Call GetKernelBasedll
    
    data(0) = Val(Text3.Text)
    WriteProcessMemory Handle, ByVal KernelBasedll + &H792, data(0), 1, 0& '更改速率
    
    CloseHandle Handle
    
ElseIf GetAsyncKeyState(109) < 0 Then  '按下-

    If Val(Text3.Text) < 1 Or Val(Text3.Text) > 120 Then Text3.Text = 20
    hwnd = FindWindow(vbNullString, "Warcraft III")
    If hwnd = 0 Then MsgBox "请先运行游戏": Exit Sub
    GetWindowThreadProcessId hwnd, PID
    Handle = OpenProcess(PROCESS_ALL_ACCESS, False, PID)
    Call GetKernelBasedll
    
    data(0) = 1
    WriteProcessMemory Handle, ByVal KernelBasedll + &H792, data(0), 1, 0& '更改速率
    
    CloseHandle Handle
    
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

