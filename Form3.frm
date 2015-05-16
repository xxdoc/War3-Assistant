VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "作弊"
   ClientHeight    =   4644
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   5724
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4644
   ScaleWidth      =   5724
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame1 
      Caption         =   "开图相关"
      Height          =   1335
      Left            =   3000
      TabIndex        =   44
      ToolTipText     =   "此面板功能可能会导致掉线，请谨慎使用"
      Top             =   2520
      Width           =   2655
      Begin VB.CheckBox Check7 
         Caption         =   "敌军头像"
         Height          =   255
         Left            =   1440
         TabIndex        =   50
         ToolTipText     =   "此面板功能可能会导致掉线，请谨慎使用"
         Top             =   960
         Width           =   1095
      End
      Begin VB.CheckBox Check6 
         Caption         =   "盟军头像"
         Height          =   255
         Left            =   120
         TabIndex        =   49
         ToolTipText     =   "此面板功能可能会导致掉线，请谨慎使用"
         Top             =   960
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         Caption         =   "视野外点击"
         Height          =   255
         Left            =   120
         TabIndex        =   45
         ToolTipText     =   "此面板功能可能会导致掉线，请谨慎使用"
         Top             =   240
         Width           =   1215
      End
      Begin VB.CheckBox Check5 
         Caption         =   "敌方信号"
         Height          =   300
         Left            =   1440
         TabIndex        =   48
         ToolTipText     =   "此面板功能可能会导致掉线，请谨慎使用"
         Top             =   600
         Width           =   1095
      End
      Begin VB.CheckBox Check3 
         Caption         =   "显示资源"
         Height          =   255
         Left            =   120
         TabIndex        =   46
         ToolTipText     =   "此面板功能可能会导致掉线，请谨慎使用"
         Top             =   600
         Width           =   1095
      End
      Begin VB.CheckBox Check4 
         Caption         =   "清除迷雾"
         Height          =   300
         Left            =   1440
         TabIndex        =   47
         ToolTipText     =   "此面板功能可能会导致掉线，请谨慎使用"
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "锁定"
      Height          =   375
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   43
      ToolTipText     =   "锁定资源（多人玩时必须确保更改或锁定资源的数据一致）"
      Top             =   1200
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   3960
      TabIndex        =   4
      ToolTipText     =   "更改的资源数，不填表示100万"
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "更改"
      Height          =   375
      Left            =   4800
      TabIndex        =   3
      ToolTipText     =   "更改资源，更改时最好暂停游戏，待所有玩家更改完毕后再继续游戏"
      Top             =   1200
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3840
      TabIndex        =   1
      ToolTipText     =   "输入要更改资源的玩家代号，用空格隔开。"
      Top             =   600
      Width           =   1815
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   2760
      Top             =   3960
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "玩家代号:"
      Height          =   180
      Left            =   3000
      TabIndex        =   51
      ToolTipText     =   "输入要更改资源的玩家代号，用空格隔开。"
      Top             =   720
      Width           =   810
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "此页功能仅支持124E"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000003&
      Height          =   240
      Left            =   3240
      TabIndex        =   42
      Top             =   4200
      Width           =   2160
   End
   Begin VB.Label Label9 
      Caption         =   "更改数目:"
      Height          =   255
      Left            =   3000
      TabIndex        =   41
      ToolTipText     =   "更改的资源数，不填表示100万"
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "木头"
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
      Left            =   2280
      TabIndex        =   40
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "金钱"
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
      Left            =   1320
      TabIndex        =   39
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Label3"
      Height          =   180
      Index           =   2
      Left            =   2160
      TabIndex        =   38
      Top             =   840
      Width           =   540
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Label3"
      Height          =   180
      Index           =   3
      Left            =   2160
      TabIndex        =   37
      Top             =   1200
      Width           =   540
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Label3"
      Height          =   180
      Index           =   4
      Left            =   2160
      TabIndex        =   36
      Top             =   1560
      Width           =   540
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Label3"
      Height          =   180
      Index           =   5
      Left            =   2160
      TabIndex        =   35
      Top             =   1920
      Width           =   540
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Label3"
      Height          =   180
      Index           =   6
      Left            =   2160
      TabIndex        =   34
      Top             =   2280
      Width           =   540
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Label3"
      Height          =   180
      Index           =   7
      Left            =   2160
      TabIndex        =   33
      Top             =   2640
      Width           =   540
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Label3"
      Height          =   180
      Index           =   8
      Left            =   2160
      TabIndex        =   32
      Top             =   3000
      Width           =   540
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Label3"
      Height          =   180
      Index           =   9
      Left            =   2160
      TabIndex        =   31
      Top             =   3360
      Width           =   540
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Label3"
      Height          =   180
      Index           =   10
      Left            =   2160
      TabIndex        =   30
      Top             =   3720
      Width           =   540
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Label3"
      Height          =   180
      Index           =   11
      Left            =   2160
      TabIndex        =   29
      Top             =   4080
      Width           =   540
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Label3"
      Height          =   180
      Index           =   12
      Left            =   2160
      TabIndex        =   28
      Top             =   4440
      Width           =   540
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Label3"
      Height          =   180
      Index           =   1
      Left            =   2160
      TabIndex        =   27
      Top             =   480
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   180
      Index           =   3
      Left            =   1200
      TabIndex        =   26
      Top             =   1200
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   180
      Index           =   4
      Left            =   1200
      TabIndex        =   25
      Top             =   1560
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   180
      Index           =   5
      Left            =   1200
      TabIndex        =   24
      Top             =   1920
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   180
      Index           =   6
      Left            =   1200
      TabIndex        =   23
      Top             =   2280
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   180
      Index           =   7
      Left            =   1200
      TabIndex        =   22
      Top             =   2640
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   180
      Index           =   8
      Left            =   1200
      TabIndex        =   21
      Top             =   3000
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   180
      Index           =   9
      Left            =   1200
      TabIndex        =   20
      Top             =   3360
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   180
      Index           =   10
      Left            =   1200
      TabIndex        =   19
      Top             =   3720
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   180
      Index           =   11
      Left            =   1200
      TabIndex        =   18
      Top             =   4080
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   180
      Index           =   12
      Left            =   1200
      TabIndex        =   17
      Top             =   4440
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   180
      Index           =   2
      Left            =   1200
      TabIndex        =   16
      Top             =   840
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   180
      Index           =   1
      Left            =   1200
      TabIndex        =   15
      Top             =   480
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "玩家7"
      Height          =   180
      Index           =   7
      Left            =   360
      TabIndex        =   14
      Top             =   2640
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "玩家8"
      Height          =   180
      Index           =   8
      Left            =   360
      TabIndex        =   13
      Top             =   3000
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "玩家9"
      Height          =   180
      Index           =   9
      Left            =   360
      TabIndex        =   12
      Top             =   3360
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "玩家10"
      Height          =   180
      Index           =   10
      Left            =   360
      TabIndex        =   11
      Top             =   3720
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "玩家11"
      Height          =   180
      Index           =   11
      Left            =   360
      TabIndex        =   10
      Top             =   4080
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "玩家12"
      Height          =   180
      Index           =   12
      Left            =   360
      TabIndex        =   9
      Top             =   4440
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "玩家6"
      Height          =   180
      Index           =   6
      Left            =   360
      TabIndex        =   8
      Top             =   2280
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "玩家5"
      Height          =   180
      Index           =   5
      Left            =   360
      TabIndex        =   7
      Top             =   1920
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "玩家4"
      Height          =   180
      Index           =   4
      Left            =   360
      TabIndex        =   6
      Top             =   1560
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "玩家3"
      Height          =   180
      Index           =   3
      Left            =   360
      TabIndex        =   5
      Top             =   1200
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "玩家2"
      Height          =   180
      Index           =   2
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "玩家1"
      Height          =   180
      Index           =   1
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   450
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const WM_KEYDOWN = &H100
Private Const WM_KEYUP = &H101
Dim HHwnd As Long, Handle As Long, PID As Long, MMY As Long

Private Sub Check2_Click()
HHwnd = FindWindow(vbNullString, "Warcraft III")
GetWindowThreadProcessId HHwnd, PID
Handle = OpenProcess(PROCESS_ALL_ACCESS, False, PID)
If Check2.Value = 0 Then
    Call 不可选视野外单位124E(Handle)
Else
    Call 可选视野外单位124E(Handle)
End If
End Sub

Private Sub Check3_Click()
HHwnd = FindWindow(vbNullString, "Warcraft III")
GetWindowThreadProcessId HHwnd, PID
Handle = OpenProcess(PROCESS_ALL_ACCESS, False, PID)
If Check3.Value = 0 Then
    Call 不资源124E(Handle)
Else
    Call 资源124E(Handle)
End If
End Sub

Private Sub Check4_Click()
HHwnd = FindWindow(vbNullString, "Warcraft III")
GetWindowThreadProcessId HHwnd, PID
Handle = OpenProcess(PROCESS_ALL_ACCESS, False, PID)
If Check4.Value = 0 Then
    Call 不清除迷雾124E(Handle)
Else
    Call 清除迷雾124E(Handle)
End If
End Sub

Private Sub Check5_Click()
HHwnd = FindWindow(vbNullString, "Warcraft III")
GetWindowThreadProcessId HHwnd, PID
Handle = OpenProcess(PROCESS_ALL_ACCESS, False, PID)
If Check5.Value = 0 Then
    Call 不显示敌方信号124E(Handle)
Else
    Call 显示敌方信号124E(Handle)
End If
End Sub

Private Sub Check6_Click()
HHwnd = FindWindow(vbNullString, "Warcraft III")
GetWindowThreadProcessId HHwnd, PID
Handle = OpenProcess(PROCESS_ALL_ACCESS, False, PID)
If Check6.Value = 0 Then
    Call 不显示头像124E(Handle)
Else
    Call 显示盟军头像124E(Handle)
End If
End Sub

Private Sub Check7_Click()
HHwnd = FindWindow(vbNullString, "Warcraft III")
GetWindowThreadProcessId HHwnd, PID
Handle = OpenProcess(PROCESS_ALL_ACCESS, False, PID)
If Check7.Value = 0 Then
    Call 不显示头像124E(Handle)
Else
    Call 显示敌方头像124E(Handle)
End If
End Sub

Private Sub Command1_Click() '更改
Dim PLR() As String
Dim Addr As Long, yStr As String
HHwnd = FindWindow(vbNullString, "Warcraft III")
GetWindowThreadProcessId HHwnd, PID
Handle = OpenProcess(PROCESS_ALL_ACCESS, False, PID)
MMY = 10000000

If Val(Text1.Text) > 0 Then

    If Text2.Text <> "" Then
        If Val(Text2.Text) > 1000000 Then
        MMY = 10000000
        Text2.Text = "1000000"
        Else
        MMY = Val(Text2.Text) * 10
        End If
    End If
    If right(Text1.Text, 1) <> " " Then
        Text1.Text = Text1.Text & " "
    End If
    PLR() = Split(Text1.Text, " ")
    For i = 0 To UBound(PLR) - 1
        If Val(PLR(i)) = 1 Then
ReadProcessMemory Handle, ByVal Wargamedll + &HACE5E0, Addr, 4, 0&
Addr = Addr + &HC
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H1E4
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H408
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H4
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H7C
WriteProcessMemory Handle, ByVal Addr, MMY, 4, 0&

ReadProcessMemory Handle, ByVal Wargamedll + &HACE5E0, Addr, 4, 0&
Addr = Addr + &H4
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H1C
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H24
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H4
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H54
WriteProcessMemory Handle, ByVal Addr, MMY, 4, 0&

        ElseIf Val(PLR(i)) = 2 Then
ReadProcessMemory Handle, ByVal Wargamedll + &HACE5E4, Addr, 4, 0&
Addr = Addr + &H4
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H13C
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H128
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H0
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H54
WriteProcessMemory Handle, ByVal Addr, MMY, 4, 0&

ReadProcessMemory Handle, ByVal Wargamedll + &HACE5E0, Addr, 4, 0&
Addr = Addr + &H4
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H12C
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H24
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H284
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H54
WriteProcessMemory Handle, ByVal Addr, MMY, 4, 0&

         ElseIf Val(PLR(i)) = 3 Then
ReadProcessMemory Handle, ByVal Wargamedll + &HACE5E0, Addr, 4, 0&
Addr = Addr + &H4
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H254
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H1FC
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H130
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H78
WriteProcessMemory Handle, ByVal Addr, MMY, 4, 0&

ReadProcessMemory Handle, ByVal Wargamedll + &HACE5E0, Addr, 4, 0&
Addr = Addr + &H4
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H244
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H2A8
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H318
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H78
WriteProcessMemory Handle, ByVal Addr, MMY, 4, 0&

         ElseIf Val(PLR(i)) = 4 Then
ReadProcessMemory Handle, ByVal Wargamedll + &HACE5E0, Addr, 4, 0&
Addr = Addr + &H4
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H3D4
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H24
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H4
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H54
WriteProcessMemory Handle, ByVal Addr, MMY, 4, 0&

ReadProcessMemory Handle, ByVal Wargamedll + &HACE5E0, Addr, 4, 0&
Addr = Addr + &H4
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H334
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H6BC
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H2AC
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H78
WriteProcessMemory Handle, ByVal Addr, MMY, 4, 0&

         ElseIf Val(PLR(i)) = 5 Then
ReadProcessMemory Handle, ByVal Wargamedll + &HACE5E0, Addr, 4, 0&
Addr = Addr + &H4
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H514
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H24
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H4
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H54
WriteProcessMemory Handle, ByVal Addr, MMY, 4, 0&

ReadProcessMemory Handle, ByVal Wargamedll + &HACE5E0, Addr, 4, 0&
Addr = Addr + &H4
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H51C
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H24
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H4
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H54
WriteProcessMemory Handle, ByVal Addr, MMY, 4, 0&
         
         ElseIf Val(PLR(i)) = 6 Then
ReadProcessMemory Handle, ByVal Wargamedll + &HACE5E0, Addr, 4, 0&
Addr = Addr + &H4
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H654
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H24
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H4
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H54
WriteProcessMemory Handle, ByVal Addr, MMY, 4, 0&

ReadProcessMemory Handle, ByVal Wargamedll + &HACE5E0, Addr, 4, 0&
Addr = Addr + &H4
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H65C
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H24
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H4
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H54
WriteProcessMemory Handle, ByVal Addr, MMY, 4, 0&
                  
         ElseIf Val(PLR(i)) = 7 Then

ReadProcessMemory Handle, ByVal Wargamedll + &HACE5E0, Addr, 4, 0&
Addr = Addr + &H4
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H794
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H24
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H4
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H54
WriteProcessMemory Handle, ByVal Addr, MMY, 4, 0&

ReadProcessMemory Handle, ByVal Wargamedll + &HACE5E0, Addr, 4, 0&
Addr = Addr + &H4
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H79C
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H24
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H4
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H54
WriteProcessMemory Handle, ByVal Addr, MMY, 4, 0&

         ElseIf Val(PLR(i)) = 8 Then

ReadProcessMemory Handle, ByVal Wargamedll + &HACE5E0, Addr, 4, 0&
Addr = Addr + &H4
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H7CC
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H7A4
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H594
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H254
WriteProcessMemory Handle, ByVal Addr, MMY, 4, 0&

ReadProcessMemory Handle, ByVal Wargamedll + &HACE5E0, Addr, 4, 0&
Addr = Addr + &H4
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H7DC
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H6A4
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H710
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H54
WriteProcessMemory Handle, ByVal Addr, MMY, 4, 0&

         ElseIf Val(PLR(i)) = 9 Then

ReadProcessMemory Handle, ByVal Wargamedll + &HACE5E0, Addr, 4, 0&
Addr = Addr + &H4
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H424
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H340
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H728
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H154
WriteProcessMemory Handle, ByVal Addr, MMY, 4, 0&

ReadProcessMemory Handle, ByVal Wargamedll + &HACE5E0, Addr, 4, 0&
Addr = Addr + &H4
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H284
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H7E4
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H688
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &HF8
WriteProcessMemory Handle, ByVal Addr, MMY, 4, 0&

         ElseIf Val(PLR(i)) = 10 Then

ReadProcessMemory Handle, ByVal Wargamedll + &HACE5E0, Addr, 4, 0&
Addr = Addr + &H4
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H784
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H4C4
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H0
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &HF8
WriteProcessMemory Handle, ByVal Addr, MMY, 4, 0&

ReadProcessMemory Handle, ByVal Wargamedll + &HACE5E0, Addr, 4, 0&
Addr = Addr + &H4
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H784
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H438
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H188
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H78
WriteProcessMemory Handle, ByVal Addr, MMY, 4, 0&
         ElseIf Val(PLR(i)) = 11 Then

ReadProcessMemory Handle, ByVal Wargamedll + &HACE5E0, Addr, 4, 0&
Addr = Addr + &H4
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H784
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H5B0
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H80
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H54
WriteProcessMemory Handle, ByVal Addr, MMY, 4, 0&

ReadProcessMemory Handle, ByVal Wargamedll + &HACE5E0, Addr, 4, 0&
Addr = Addr + &H4
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H784
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H5B0
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H100
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H54
WriteProcessMemory Handle, ByVal Addr, MMY, 4, 0&


         ElseIf Val(PLR(i)) = 12 Then
ReadProcessMemory Handle, ByVal Wargamedll + &HACE5E0, Addr, 4, 0&
Addr = Addr + &H4
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H784
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H728
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H80
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H54
WriteProcessMemory Handle, ByVal Addr, MMY, 4, 0&
             
ReadProcessMemory Handle, ByVal Wargamedll + &HACE5E0, Addr, 4, 0&
Addr = Addr + &H4
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H784
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H728
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H100
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H54
WriteProcessMemory Handle, ByVal Addr, MMY, 4, 0&
         End If
    Next
End If
End Sub

Private Sub Command2_Click() '锁定
If Command2.Caption = "锁定" Then
    Command2.Caption = "已锁定"
    Command2.BackColor = vbRed
    
    Text1.Enabled = False
    Text2.Enabled = False
    
    ZYSUB '改资源过程
    
    HHwnd = FindWindow(vbNullString, "Warcraft III")
    GetWindowThreadProcessId HHwnd, PID
    Handle = OpenProcess(PROCESS_ALL_ACCESS, False, PID)
    
    WriteProcessMemory Handle, ByVal Wargamedll + &H473C5E, &HE5EB, 2, 0& 'jmp 到该资源指令
Else
    Command2.Caption = "锁定"
    Command2.BackColor = &H8000000F
    
    HHwnd = FindWindow(vbNullString, "Warcraft III")
    GetWindowThreadProcessId HHwnd, PID
    Handle = OpenProcess(PROCESS_ALL_ACCESS, False, PID)
    
    WriteProcessMemory Handle, ByVal Wargamedll + &H473C5E, &H548B, 2, 0& '恢复
    
    Text1.Enabled = True
    Text2.Enabled = True
End If
End Sub

Private Sub Form_Activate()
Form1.CHKHh.Refresh
Form1.CHKkey.Refresh
Form1.CHKMH.Refresh
Form1.CHKXX.Refresh
Form3.Command1.Refresh
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If Cancel = 0 Then
        Cancel = 1
        Me.Hide
        Form3.Timer1 = False
    End If
End Sub

Private Sub Timer1_Timer()
Dim Money As Long, Mutou As Long, Addr As Long, Player As Long, ZT As Long, WJS As Integer
HHwnd = FindWindow(vbNullString, "Warcraft III")
GetWindowThreadProcessId HHwnd, PID
Handle = OpenProcess(PROCESS_ALL_ACCESS, False, PID)
Call Getgamedll
Call GetGGWAR3dll
ReadProcessMemory Handle, ByVal Wargamedll + &HACE5E0, Addr, 4, 0&
Addr = Addr + &HC
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H1E4
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H408
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H4
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H7C
ReadProcessMemory Handle, ByVal Addr, Money, 4, 0&
Label2(1) = Money / 10

ReadProcessMemory Handle, ByVal Wargamedll + &HACE5E4, Addr, 4, 0&
Addr = Addr + &H4
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H13C
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H128
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H0
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H54
ReadProcessMemory Handle, ByVal Addr, Money, 4, 0&
Label2(2) = Money / 10

ReadProcessMemory Handle, ByVal Wargamedll + &HACE5E0, Addr, 4, 0&
Addr = Addr + &H4
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H254
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H1FC
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H130
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H78
ReadProcessMemory Handle, ByVal Addr, Money, 4, 0&
Label2(3) = Money / 10

ReadProcessMemory Handle, ByVal Wargamedll + &HACE5E0, Addr, 4, 0&
Addr = Addr + &H4
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H3D4
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H24
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H4
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H54
ReadProcessMemory Handle, ByVal Addr, Money, 4, 0&
Label2(4) = Money / 10

ReadProcessMemory Handle, ByVal Wargamedll + &HACE5E0, Addr, 4, 0&
Addr = Addr + &H4
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H514
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H24
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H4
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H54
ReadProcessMemory Handle, ByVal Addr, Money, 4, 0&
Label2(5) = Money / 10

ReadProcessMemory Handle, ByVal Wargamedll + &HACE5E0, Addr, 4, 0&
Addr = Addr + &H4
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H654
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H24
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H4
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H54
ReadProcessMemory Handle, ByVal Addr, Money, 4, 0&
Label2(6) = Money / 10

ReadProcessMemory Handle, ByVal Wargamedll + &HACE5E0, Addr, 4, 0&
Addr = Addr + &H4
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H794
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H24
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H4
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H54
ReadProcessMemory Handle, ByVal Addr, Money, 4, 0&
Label2(7) = Money / 10

ReadProcessMemory Handle, ByVal Wargamedll + &HACE5E0, Addr, 4, 0&
Addr = Addr + &H4
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H7CC
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H7A4
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H594
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H254
ReadProcessMemory Handle, ByVal Addr, Money, 4, 0&
Label2(8) = Money / 10

ReadProcessMemory Handle, ByVal Wargamedll + &HACE5E0, Addr, 4, 0&
Addr = Addr + &H4
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H424
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H340
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H728
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H154
ReadProcessMemory Handle, ByVal Addr, Money, 4, 0&
Label2(9) = Money / 10

ReadProcessMemory Handle, ByVal Wargamedll + &HACE5E0, Addr, 4, 0&
Addr = Addr + &H4
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H784
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H4C4
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H0
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &HF8
ReadProcessMemory Handle, ByVal Addr, Money, 4, 0&
Label2(10) = Money / 10

ReadProcessMemory Handle, ByVal Wargamedll + &HACE5E0, Addr, 4, 0&
Addr = Addr + &H4
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H784
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H5B0
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H80
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H54
ReadProcessMemory Handle, ByVal Addr, Money, 4, 0&
Label2(11) = Money / 10

ReadProcessMemory Handle, ByVal Wargamedll + &HACE5E0, Addr, 4, 0&
Addr = Addr + &H4
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H784
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H728
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H80
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H54
ReadProcessMemory Handle, ByVal Addr, Money, 4, 0&
Label2(12) = Money / 10





ReadProcessMemory Handle, ByVal Wargamedll + &HACE5E0, Addr, 4, 0&
Addr = Addr + &H4
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H1C
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H24
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H4
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H54
ReadProcessMemory Handle, ByVal Addr, Mutou, 4, 0&
Label3(1) = Mutou / 10

ReadProcessMemory Handle, ByVal Wargamedll + &HACE5E0, Addr, 4, 0&
Addr = Addr + &H4
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H12C
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H24
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H284
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H54
ReadProcessMemory Handle, ByVal Addr, Mutou, 4, 0&
Label3(2) = Mutou / 10

ReadProcessMemory Handle, ByVal Wargamedll + &HACE5E0, Addr, 4, 0&
Addr = Addr + &H4
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H244
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H2A8
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H318
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H78
ReadProcessMemory Handle, ByVal Addr, Mutou, 4, 0&
Label3(3) = Mutou / 10

ReadProcessMemory Handle, ByVal Wargamedll + &HACE5E0, Addr, 4, 0&
Addr = Addr + &H4
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H334
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H6BC
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H2AC
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H78
ReadProcessMemory Handle, ByVal Addr, Mutou, 4, 0&
Label3(4) = Mutou / 10

ReadProcessMemory Handle, ByVal Wargamedll + &HACE5E0, Addr, 4, 0&
Addr = Addr + &H4
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H51C
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H24
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H4
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H54
ReadProcessMemory Handle, ByVal Addr, Mutou, 4, 0&
Label3(5) = Mutou / 10

ReadProcessMemory Handle, ByVal Wargamedll + &HACE5E0, Addr, 4, 0&
Addr = Addr + &H4
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H65C
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H24
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H4
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H54
ReadProcessMemory Handle, ByVal Addr, Mutou, 4, 0&
Label3(6) = Mutou / 10

ReadProcessMemory Handle, ByVal Wargamedll + &HACE5E0, Addr, 4, 0&
Addr = Addr + &H4
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H79C
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H24
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H4
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H54
ReadProcessMemory Handle, ByVal Addr, Mutou, 4, 0&
Label3(7) = Mutou / 10

ReadProcessMemory Handle, ByVal Wargamedll + &HACE5E0, Addr, 4, 0&
Addr = Addr + &H4
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H7DC
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H6A4
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H710
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H54
ReadProcessMemory Handle, ByVal Addr, Mutou, 4, 0&
Label3(8) = Mutou / 10

ReadProcessMemory Handle, ByVal Wargamedll + &HACE5E0, Addr, 4, 0&
Addr = Addr + &H4
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H284
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H7E4
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H688
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &HF8
ReadProcessMemory Handle, ByVal Addr, Mutou, 4, 0&
Label3(9) = Mutou / 10

ReadProcessMemory Handle, ByVal Wargamedll + &HACE5E0, Addr, 4, 0&
Addr = Addr + &H4
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H784
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H438
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H188
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H78
ReadProcessMemory Handle, ByVal Addr, Mutou, 4, 0&
Label3(10) = Mutou / 10

ReadProcessMemory Handle, ByVal Wargamedll + &HACE5E0, Addr, 4, 0&
Addr = Addr + &H4
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H784
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H5B0
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H100
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H54
ReadProcessMemory Handle, ByVal Addr, Mutou, 4, 0&
Label3(11) = Mutou / 10

ReadProcessMemory Handle, ByVal Wargamedll + &HACE5E0, Addr, 4, 0&
Addr = Addr + &H4
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H784
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H728
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H100
ReadProcessMemory Handle, ByVal Addr, Addr, 4, 0&
Addr = Addr + &H54
ReadProcessMemory Handle, ByVal Addr, Mutou, 4, 0&
Label3(12) = Mutou / 10
CloseHandle Handle
End Sub

Private Sub ZYSUB()
Dim PLR() As String
Dim Addr As Long, yStr As String, dByte(50) As Byte
HHwnd = FindWindow(vbNullString, "Warcraft III")
GetWindowThreadProcessId HHwnd, PID
Handle = OpenProcess(PROCESS_ALL_ACCESS, False, PID)
Call Getgamedll
MMY = 10000000

If Val(Text1.Text) > 0 Then

    If Text2.Text <> "" Then
        If Val(Text2.Text) > 1000000 Then
        MMY = 10000000
        Text2.Text = "1000000"
        Else
        MMY = Val(Text2.Text) * 10
        End If
    End If
    If right(Text1.Text, 1) <> " " Then
        Text1.Text = Text1.Text & " "
    End If
    
    WriteProcessMemory Handle, ByVal Wargamedll + &H87EFFC, MMY, 4, 0& '先写入要改的资源
    
    dByte(0) = &H8B 'mov edx,dword ptr ds:[0x6F87EFFC]
    dByte(1) = &H15
    dByte(2) = &HFC
    dByte(3) = &HEF
    dByte(4) = &H87
    dByte(5) = &H6F
    dByte(6) = &HE9 'jmp game.6F473C62
    dByte(7) = &HB5
    dByte(8) = &H4C
    dByte(9) = &HBF
    dByte(10) = &HFF
    WriteProcessMemory Handle, ByVal Wargamedll + &H87EFA2, dByte(0), 11, 0& '写入改资源代码
    WriteProcessMemory Handle, ByVal Wargamedll + &H87EFA4, Wargamedll + &H87EFFC, 4, 0& '修正位置
    
    dByte(0) = &H83 'cmp edx,0x3
    dByte(1) = &HFA
    dByte(2) = &H2
    dByte(3) = &H90
    dByte(4) = &H90
    dByte(5) = &H90
    dByte(6) = &H90
    dByte(7) = &H90
    dByte(8) = &H90
    dByte(9) = &HEB 'jmp Xgame.6F473C72
    dByte(10) = &H22
    WriteProcessMemory Handle, ByVal Wargamedll + &H473C45, dByte(0), 11, 0& '1J
    
    dByte(0) = &H83 'cmp edx,0x3
    dByte(1) = &HFA
    dByte(2) = &H3
    dByte(3) = &H90
    dByte(4) = &H90
    dByte(5) = &H90
    dByte(6) = &H90
    dByte(7) = &H90
    dByte(8) = &H90
    dByte(9) = &HE9 'jmp game.6F473D84
    dByte(10) = &H4
    dByte(11) = &H1
    dByte(12) = &H0
    dByte(13) = &H0
    WriteProcessMemory Handle, ByVal Wargamedll + &H473C72, dByte(0), 14, 0& '1M
    
    dByte(0) = &H83 'cmp edx,0x2A
    dByte(1) = &HFA
    dByte(2) = &H2A
    dByte(3) = &H90
    dByte(4) = &H90
    dByte(5) = &H90
    dByte(6) = &H90
    dByte(7) = &H90
    dByte(8) = &H90
    dByte(9) = &HEB 'jmp Xgame.6F473D95
    dByte(10) = &H6
    WriteProcessMemory Handle, ByVal Wargamedll + &H473D84, dByte(0), 11, 0& '2J
    
    dByte(0) = &H83 'cmp edx,0x2B
    dByte(1) = &HFA
    dByte(2) = &H2B
    dByte(3) = &H90
    dByte(4) = &H90
    dByte(5) = &H90
    dByte(6) = &H90
    dByte(7) = &H90
    dByte(8) = &H90
    dByte(9) = &HEB 'jmp Xgame.6F473E01
    dByte(10) = &H61
    WriteProcessMemory Handle, ByVal Wargamedll + &H473D95, dByte(0), 11, 0& '2J
    
    dByte(0) = &H83 'cmp edx,0x52
    dByte(1) = &HFA
    dByte(2) = &H52
    dByte(3) = &H90
    dByte(4) = &H90
    dByte(5) = &H90
    dByte(6) = &H90
    dByte(7) = &H90
    dByte(8) = &H90
    dByte(9) = &HE9 'jmp game.6F473F04
    dByte(10) = &HF5
    dByte(11) = &H0
    dByte(12) = &H0
    dByte(13) = &H0
    WriteProcessMemory Handle, ByVal Wargamedll + &H473E01, dByte(0), 14, 0& '3J
    
    dByte(0) = &H83 'cmp edx,0x53
    dByte(1) = &HFA
    dByte(2) = &H53
    dByte(3) = &H90
    dByte(4) = &H90
    dByte(5) = &H90
    dByte(6) = &H90
    dByte(7) = &H90
    dByte(8) = &H90
    dByte(9) = &HEB 'jmp Xgame.6F473F62
    dByte(10) = &H53
    WriteProcessMemory Handle, ByVal Wargamedll + &H473F04, dByte(0), 11, 0& '3M
    
    dByte(0) = &H83 'cmp edx,0x7A
    dByte(1) = &HFA
    dByte(2) = &H7A
    dByte(3) = &H90
    dByte(4) = &H90
    dByte(5) = &H90
    dByte(6) = &H90
    dByte(7) = &H90
    dByte(8) = &H90
    dByte(9) = &HEB 'jmp Xgame.6F473FC2
    dByte(10) = &H55
    WriteProcessMemory Handle, ByVal Wargamedll + &H473F62, dByte(0), 11, 0& '4J
    
    dByte(0) = &H83 'cmp edx,0x7B
    dByte(1) = &HFA
    dByte(2) = &H7B
    dByte(3) = &H90
    dByte(4) = &H90
    dByte(5) = &H90
    dByte(6) = &H90
    dByte(7) = &H90
    dByte(8) = &H90
    dByte(9) = &HE9 'jmp game.6F87EA61
    dByte(10) = &H91
    dByte(11) = &HAA
    dByte(12) = &H40
    dByte(13) = &H0
    WriteProcessMemory Handle, ByVal Wargamedll + &H473FC2, dByte(0), 14, 0& '4M
    
    dByte(0) = &H81 'cmp edx,0xA2
    dByte(1) = &HFA
    dByte(2) = &HA2
    dByte(3) = &H0
    dByte(4) = &H0
    dByte(5) = &H0
    dByte(6) = &H90
    dByte(7) = &H90
    dByte(8) = &H90
    dByte(9) = &H90
    dByte(10) = &H90
    dByte(11) = &H90
    dByte(12) = &HEB 'jmp Xgame.6F87EA71
    dByte(13) = &H2
    WriteProcessMemory Handle, ByVal Wargamedll + &H87EA61, dByte(0), 14, 0& '5J
    
    dByte(0) = &H81 'cmp edx,0xA3
    dByte(1) = &HFA
    dByte(2) = &HA3
    dByte(3) = &H0
    dByte(4) = &H0
    dByte(5) = &H0
    dByte(6) = &H90
    dByte(7) = &H90
    dByte(8) = &H90
    dByte(9) = &H90
    dByte(10) = &H90
    dByte(11) = &H90
    dByte(12) = &HEB 'jmp Xgame.6F87EA81
    dByte(13) = &H2
    WriteProcessMemory Handle, ByVal Wargamedll + &H87EA71, dByte(0), 14, 0& '5M
    
    dByte(0) = &H81 'cmp edx,0xCA
    dByte(1) = &HFA
    dByte(2) = &HCA
    dByte(3) = &H0
    dByte(4) = &H0
    dByte(5) = &H0
    dByte(6) = &H90
    dByte(7) = &H90
    dByte(8) = &H90
    dByte(9) = &H90
    dByte(10) = &H90
    dByte(11) = &H90
    dByte(12) = &HEB 'jmp Xgame.6F87EA91
    dByte(13) = &H2
    WriteProcessMemory Handle, ByVal Wargamedll + &H87EA81, dByte(0), 14, 0& '6J
    
    dByte(0) = &H81 'cmp edx,0xCB
    dByte(1) = &HFA
    dByte(2) = &HCB
    dByte(3) = &H0
    dByte(4) = &H0
    dByte(5) = &H0
    dByte(6) = &H90
    dByte(7) = &H90
    dByte(8) = &H90
    dByte(9) = &H90
    dByte(10) = &H90
    dByte(11) = &H90
    dByte(12) = &HEB 'jmp Xgame.6F87EAA1
    dByte(13) = &H2
    WriteProcessMemory Handle, ByVal Wargamedll + &H87EA91, dByte(0), 14, 0& '6M
     
    dByte(0) = &H81 'cmp edx,0xF2
    dByte(1) = &HFA
    dByte(2) = &HF2
    dByte(3) = &H0
    dByte(4) = &H0
    dByte(5) = &H0
    dByte(6) = &H90
    dByte(7) = &H90
    dByte(8) = &H90
    dByte(9) = &H90
    dByte(10) = &H90
    dByte(11) = &H90
    dByte(12) = &HEB 'jmp Xgame.6F87EAB1
    dByte(13) = &H2
    WriteProcessMemory Handle, ByVal Wargamedll + &H87EAA1, dByte(0), 14, 0& '7J
    
    dByte(0) = &H81 'cmp edx,0xF3
    dByte(1) = &HFA
    dByte(2) = &HF3
    dByte(3) = &H0
    dByte(4) = &H0
    dByte(5) = &H0
    dByte(6) = &H90
    dByte(7) = &H90
    dByte(8) = &H90
    dByte(9) = &H90
    dByte(10) = &H90
    dByte(11) = &H90
    dByte(12) = &HEB 'jmp Xgame.6F87EB11
    dByte(13) = &H52
    WriteProcessMemory Handle, ByVal Wargamedll + &H87EAB1, dByte(0), 14, 0& '7M
    
    dByte(0) = &H81 'cmp edx,0x11A
    dByte(1) = &HFA
    dByte(2) = &H1A
    dByte(3) = &H1
    dByte(4) = &H0
    dByte(5) = &H0
    dByte(6) = &H90
    dByte(7) = &H90
    dByte(8) = &H90
    dByte(9) = &H90
    dByte(10) = &H90
    dByte(11) = &H90
    dByte(12) = &HEB 'jmp Xgame.6F87EB71
    dByte(13) = &H52
    WriteProcessMemory Handle, ByVal Wargamedll + &H87EB11, dByte(0), 14, 0& '8J
    
    dByte(0) = &H81 'cmp edx,0x11B
    dByte(1) = &HFA
    dByte(2) = &H1B
    dByte(3) = &H1
    dByte(4) = &H0
    dByte(5) = &H0
    dByte(6) = &H90
    dByte(7) = &H90
    dByte(8) = &H90
    dByte(9) = &H90
    dByte(10) = &H90
    dByte(11) = &H90
    dByte(12) = &HEB 'jmp Xgame.6F87EB71
    dByte(13) = &H55
    WriteProcessMemory Handle, ByVal Wargamedll + &H87EB71, dByte(0), 14, 0& '8M
    
    dByte(0) = &HE9 'jmp game.6F87ECA1
    dByte(1) = &HC8
    dByte(2) = &H0
    dByte(3) = &H0
    dByte(4) = &H0
    WriteProcessMemory Handle, ByVal Wargamedll + &H87EBD4, dByte(0), 5, 0& '跳板
    
    dByte(0) = &H81 'cmp edx,0x142
    dByte(1) = &HFA
    dByte(2) = &H42
    dByte(3) = &H1
    dByte(4) = &H0
    dByte(5) = &H0
    dByte(6) = &H90
    dByte(7) = &H90
    dByte(8) = &H90
    dByte(9) = &H90
    dByte(10) = &H90
    dByte(11) = &H90
    dByte(12) = &HEB 'jmp Xgame.6F87EB71
    dByte(13) = &H52
    WriteProcessMemory Handle, ByVal Wargamedll + &H87ECA1, dByte(0), 14, 0& '9J
    
    dByte(0) = &H81 'cmp edx,0x143
    dByte(1) = &HFA
    dByte(2) = &H43
    dByte(3) = &H1
    dByte(4) = &H0
    dByte(5) = &H0
    dByte(6) = &H90
    dByte(7) = &H90
    dByte(8) = &H90
    dByte(9) = &H90
    dByte(10) = &H90
    dByte(11) = &H90
    dByte(12) = &HEB 'jmp Xgame.6F87EB71
    dByte(13) = &HB
    WriteProcessMemory Handle, ByVal Wargamedll + &H87ED01, dByte(0), 14, 0& '9M
    
    WriteProcessMemory Handle, ByVal Wargamedll + &H87ED1A, &H75EB, 2, 0& '跳板
    
    dByte(0) = &H81 'cmp edx,0x16A
    dByte(1) = &HFA
    dByte(2) = &H6A
    dByte(3) = &H1
    dByte(4) = &H0
    dByte(5) = &H0
    dByte(6) = &H90
    dByte(7) = &H90
    dByte(8) = &H90
    dByte(9) = &H90
    dByte(10) = &H90
    dByte(11) = &H90
    dByte(12) = &HEB 'jmp Xgame.6F87EB71
    dByte(13) = &H52
    WriteProcessMemory Handle, ByVal Wargamedll + &H87ED91, dByte(0), 14, 0& '10J
    
    dByte(0) = &H81 'cmp edx,0x16A
    dByte(1) = &HFA
    dByte(2) = &H6B
    dByte(3) = &H1
    dByte(4) = &H0
    dByte(5) = &H0
    dByte(6) = &H90
    dByte(7) = &H90
    dByte(8) = &H90
    dByte(9) = &H90
    dByte(10) = &H90
    dByte(11) = &H90
    dByte(12) = &HEB 'jmp Xgame.6F87EB71
    dByte(13) = &H52
    WriteProcessMemory Handle, ByVal Wargamedll + &H87EDF1, dByte(0), 14, 0& '10M
    
    dByte(0) = &H81 'cmp edx,0x16A
    dByte(1) = &HFA
    dByte(2) = &H92
    dByte(3) = &H1
    dByte(4) = &H0
    dByte(5) = &H0
    dByte(6) = &H90
    dByte(7) = &H90
    dByte(8) = &H90
    dByte(9) = &H90
    dByte(10) = &H90
    dByte(11) = &H90
    dByte(12) = &HEB 'jmp Xgame.6F87EB71
    dByte(13) = &H52
    WriteProcessMemory Handle, ByVal Wargamedll + &H87EE51, dByte(0), 14, 0& '11J
    
    dByte(0) = &H81 'cmp edx,0x16A
    dByte(1) = &HFA
    dByte(2) = &H93
    dByte(3) = &H1
    dByte(4) = &H0
    dByte(5) = &H0
    dByte(6) = &H90
    dByte(7) = &H90
    dByte(8) = &H90
    dByte(9) = &H90
    dByte(10) = &H90
    dByte(11) = &H90
    dByte(12) = &HEB 'jmp Xgame.6F87EB71
    dByte(13) = &H52
    WriteProcessMemory Handle, ByVal Wargamedll + &H87EEB1, dByte(0), 14, 0& '11M
    
    dByte(0) = &H81 'cmp edx,0x16A
    dByte(1) = &HFA
    dByte(2) = &HBA
    dByte(3) = &H1
    dByte(4) = &H0
    dByte(5) = &H0
    dByte(6) = &H90
    dByte(7) = &H90
    dByte(8) = &H90
    dByte(9) = &H90
    dByte(10) = &H90
    dByte(11) = &H90
    dByte(12) = &HEB 'jmp Xgame.6F87EB71
    dByte(13) = &HB
    WriteProcessMemory Handle, ByVal Wargamedll + &H87EF11, dByte(0), 14, 0& '12J
    
    dByte(0) = &HE9 'jmp game.6F87ECA1
    dByte(1) = &H8B
    dByte(2) = &H0
    dByte(3) = &H0
    dByte(4) = &H0
    WriteProcessMemory Handle, ByVal Wargamedll + &H87EF2A, dByte(0), 5, 0& '跳板
    
    dByte(0) = &H81 'cmp edx,0x16A
    dByte(1) = &HFA
    dByte(2) = &HBB
    dByte(3) = &H1
    dByte(4) = &H0
    dByte(5) = &H0
    dByte(6) = &H90
    dByte(7) = &H90    '12M
    dByte(8) = &H8B
    dByte(9) = &H54
    dByte(10) = &H24
    dByte(11) = &H4
    dByte(12) = &HE9
    dByte(13) = &H97
    dByte(14) = &H4C
    dByte(15) = &HBF
    dByte(16) = &HFF
    WriteProcessMemory Handle, ByVal Wargamedll + &H87EFBA, dByte(0), 17, 0&
    
    

    PLR() = Split(Text1.Text, " ")
    For i = 0 To UBound(PLR) - 1
        If Val(PLR(i)) = 1 Then
    dByte(0) = &HF
    dByte(1) = &H84
    dByte(2) = &H54
    dByte(3) = &HB3
    dByte(4) = &H40
    dByte(5) = &H0
    WriteProcessMemory Handle, ByVal Wargamedll + &H473C48, dByte(0), 6, 0& '1J
    dByte(0) = &HF
    dByte(1) = &H84
    dByte(2) = &H27
    dByte(3) = &HB3
    dByte(4) = &H40
    dByte(5) = &H0
    WriteProcessMemory Handle, ByVal Wargamedll + &H473C75, dByte(0), 6, 0& '1M

        ElseIf Val(PLR(i)) = 2 Then
    dByte(0) = &HF
    dByte(1) = &H84
    dByte(2) = &H15
    dByte(3) = &HB2
    dByte(4) = &H40
    dByte(5) = &H0
    WriteProcessMemory Handle, ByVal Wargamedll + &H473D87, dByte(0), 6, 0& '2J
    dByte(0) = &HF
    dByte(1) = &H84
    dByte(2) = &H4
    dByte(3) = &HB2
    dByte(4) = &H40
    dByte(5) = &H0
    WriteProcessMemory Handle, ByVal Wargamedll + &H473D98, dByte(0), 6, 0& '2M

         ElseIf Val(PLR(i)) = 3 Then
    dByte(0) = &HF
    dByte(1) = &H84
    dByte(2) = &H98
    dByte(3) = &HB1
    dByte(4) = &H40
    dByte(5) = &H0
    WriteProcessMemory Handle, ByVal Wargamedll + &H473E04, dByte(0), 6, 0& '3J
    dByte(0) = &HF
    dByte(1) = &H84
    dByte(2) = &H95
    dByte(3) = &HB0
    dByte(4) = &H40
    dByte(5) = &H0
    WriteProcessMemory Handle, ByVal Wargamedll + &H473F07, dByte(0), 6, 0& '3M

         ElseIf Val(PLR(i)) = 4 Then
    dByte(0) = &HF
    dByte(1) = &H84
    dByte(2) = &H37
    dByte(3) = &HB0
    dByte(4) = &H40
    dByte(5) = &H0
    WriteProcessMemory Handle, ByVal Wargamedll + &H473F65, dByte(0), 6, 0& '4J
    dByte(0) = &HF
    dByte(1) = &H84
    dByte(2) = &HD7
    dByte(3) = &HAF
    dByte(4) = &H40
    dByte(5) = &H0
    WriteProcessMemory Handle, ByVal Wargamedll + &H473FC5, dByte(0), 6, 0& '4M
         ElseIf Val(PLR(i)) = 5 Then
    dByte(0) = &HF
    dByte(1) = &H84
    dByte(2) = &H35
    dByte(3) = &H5
    dByte(4) = &H0
    dByte(5) = &H0
    WriteProcessMemory Handle, ByVal Wargamedll + &H87EA67, dByte(0), 6, 0& '5J
    dByte(0) = &HF
    dByte(1) = &H84
    dByte(2) = &H25
    dByte(3) = &H5
    dByte(4) = &H0
    dByte(5) = &H0
    WriteProcessMemory Handle, ByVal Wargamedll + &H87EA77, dByte(0), 6, 0& '5M
         
         ElseIf Val(PLR(i)) = 6 Then
    dByte(0) = &HF
    dByte(1) = &H84
    dByte(2) = &H15
    dByte(3) = &H5
    dByte(4) = &H0
    dByte(5) = &H0
    WriteProcessMemory Handle, ByVal Wargamedll + &H87EA87, dByte(0), 6, 0& '6J
    dByte(0) = &HF
    dByte(1) = &H84
    dByte(2) = &H5
    dByte(3) = &H5
    dByte(4) = &H0
    dByte(5) = &H0
    WriteProcessMemory Handle, ByVal Wargamedll + &H87EA97, dByte(0), 6, 0& '6M
         ElseIf Val(PLR(i)) = 7 Then
    dByte(0) = &HF
    dByte(1) = &H84
    dByte(2) = &HF5
    dByte(3) = &H4
    dByte(4) = &H0
    dByte(5) = &H0
    WriteProcessMemory Handle, ByVal Wargamedll + &H87EAA7, dByte(0), 6, 0& '7J
    dByte(0) = &HF
    dByte(1) = &H84
    dByte(2) = &HE5
    dByte(3) = &H4
    dByte(4) = &H0
    dByte(5) = &H0
    WriteProcessMemory Handle, ByVal Wargamedll + &H87EAB7, dByte(0), 6, 0& '7M
         ElseIf Val(PLR(i)) = 8 Then
    dByte(0) = &HF
    dByte(1) = &H84
    dByte(2) = &H85
    dByte(3) = &H4
    dByte(4) = &H0
    dByte(5) = &H0
    WriteProcessMemory Handle, ByVal Wargamedll + &H87EB17, dByte(0), 6, 0& '8J
    dByte(0) = &HF
    dByte(1) = &H84
    dByte(2) = &H25
    dByte(3) = &H4
    dByte(4) = &H0
    dByte(5) = &H0
    WriteProcessMemory Handle, ByVal Wargamedll + &H87EB77, dByte(0), 6, 0& '8M
         ElseIf Val(PLR(i)) = 9 Then
    dByte(0) = &HF
    dByte(1) = &H84
    dByte(2) = &HF5
    dByte(3) = &H2
    dByte(4) = &H0
    dByte(5) = &H0
    WriteProcessMemory Handle, ByVal Wargamedll + &H87ECA7, dByte(0), 6, 0& '9J
    dByte(0) = &HF
    dByte(1) = &H84
    dByte(2) = &H95
    dByte(3) = &H2
    dByte(4) = &H0
    dByte(5) = &H0
    WriteProcessMemory Handle, ByVal Wargamedll + &H87ED07, dByte(0), 6, 0& '9M
         ElseIf Val(PLR(i)) = 10 Then
    dByte(0) = &HF
    dByte(1) = &H84
    dByte(2) = &H5
    dByte(3) = &H2
    dByte(4) = &H0
    dByte(5) = &H0
    WriteProcessMemory Handle, ByVal Wargamedll + &H87ED97, dByte(0), 6, 0& '10J
    dByte(0) = &HF
    dByte(1) = &H84
    dByte(2) = &HA5
    dByte(3) = &H1
    dByte(4) = &H0
    dByte(5) = &H0
    WriteProcessMemory Handle, ByVal Wargamedll + &H87EDF7, dByte(0), 6, 0& '10M
         ElseIf Val(PLR(i)) = 11 Then
    dByte(0) = &HF
    dByte(1) = &H84
    dByte(2) = &H45
    dByte(3) = &H1
    dByte(4) = &H0
    dByte(5) = &H0
    WriteProcessMemory Handle, ByVal Wargamedll + &H87EE57, dByte(0), 6, 0&     '11J
    dByte(0) = &HF
    dByte(1) = &H84
    dByte(2) = &HE5
    dByte(3) = &H0
    dByte(4) = &H0
    dByte(5) = &H0
    WriteProcessMemory Handle, ByVal Wargamedll + &H87EEB7, dByte(0), 6, 0& '11M
         ElseIf Val(PLR(i)) = 12 Then
    dByte(0) = &HF
    dByte(1) = &H84
    dByte(2) = &H85
    dByte(3) = &H0
    dByte(4) = &H0
    dByte(5) = &H0
    WriteProcessMemory Handle, ByVal Wargamedll + &H87EF17, dByte(0), 6, 0&     '12J
    dByte(0) = &H74
    dByte(1) = &HE0
    WriteProcessMemory Handle, ByVal Wargamedll + &H87EFC0, dByte(0), 2, 0& '12M
         End If
    Next
End If
End Sub
