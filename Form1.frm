VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "War3 Assistant"
   ClientHeight    =   2580
   ClientLeft      =   7215
   ClientTop       =   3660
   ClientWidth     =   6180
   DrawMode        =   1  'Blackness
   ForeColor       =   &H00000000&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   6180
   ShowInTaskbar   =   0   'False
   Begin VB.Timer QD 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   3480
      Top             =   2280
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   2040
      Top             =   2280
   End
   Begin VB.Timer Timer4 
      Interval        =   100
      Left            =   2520
      Top             =   2280
   End
   Begin VB.CommandButton Command3 
      Caption         =   "作弊"
      Height          =   375
      Left            =   4560
      TabIndex        =   25
      ToolTipText     =   "打开改资源窗体"
      Top             =   2220
      Width           =   555
   End
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   4500
      Top             =   2340
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   3000
      Top             =   2280
   End
   Begin VB.CommandButton Command2 
      Caption         =   "退出"
      Height          =   375
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "退出程序"
      Top             =   2220
      Width           =   555
   End
   Begin VB.CommandButton Command1 
      Caption         =   "更多"
      Height          =   375
      Left            =   5100
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "显示更多"
      Top             =   2220
      Width           =   555
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   7740
      Top             =   2640
   End
   Begin VB.CheckBox CHKMH 
      Caption         =   "开图"
      Height          =   255
      Left            =   4560
      TabIndex        =   18
      ToolTipText     =   "快捷键：HOME开图，END取消开图"
      Top             =   180
      Width           =   675
   End
   Begin VB.Timer HhTimer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   7320
      Top             =   2640
   End
   Begin VB.Frame Frame1 
      Caption         =   "改键"
      Height          =   2235
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2835
      Begin VB.CheckBox Check1 
         Caption         =   "快速启动"
         Height          =   255
         Left            =   1680
         TabIndex        =   45
         ToolTipText     =   "勾选后同时按+-可启动游戏,请先在更多窗口处设置游戏路径"
         Top             =   180
         Width           =   1030
      End
      Begin VB.CheckBox CHKXX 
         Caption         =   "显血"
         Height          =   255
         Left            =   960
         TabIndex        =   2
         ToolTipText     =   "开启显血"
         Top             =   180
         Width           =   675
      End
      Begin VB.TextBox bKeycodeText 
         Height          =   270
         Index           =   10
         Left            =   2520
         TabIndex        =   16
         Top             =   1920
         Width           =   255
      End
      Begin VB.TextBox bKeycodeText 
         Height          =   270
         Index           =   9
         Left            =   1080
         TabIndex        =   14
         Top             =   1920
         Width           =   255
      End
      Begin VB.TextBox bKeycodeText 
         Height          =   270
         Index           =   7
         Left            =   1080
         TabIndex        =   10
         Top             =   1560
         Width           =   255
      End
      Begin VB.TextBox bKeycodeText 
         Height          =   270
         Index           =   8
         Left            =   2520
         TabIndex        =   12
         Top             =   1560
         Width           =   255
      End
      Begin VB.TextBox aKeycodeText 
         Height          =   270
         Index           =   10
         Left            =   1560
         TabIndex        =   15
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox aKeycodeText 
         Height          =   270
         Index           =   8
         Left            =   1560
         TabIndex        =   11
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox aKeycodeText 
         Height          =   270
         Index           =   9
         Left            =   120
         TabIndex        =   13
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox aKeycodeText 
         Height          =   270
         Index           =   7
         Left            =   120
         TabIndex        =   9
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox aKeycodeText 
         Height          =   270
         Index           =   6
         Left            =   1560
         TabIndex        =   8
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox aKeycodeText 
         Height          =   270
         Index           =   5
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox aKeycodeText 
         Height          =   270
         Index           =   4
         Left            =   1560
         TabIndex        =   6
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox aKeycodeText 
         Height          =   270
         Index           =   3
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox aKeycodeText 
         Height          =   270
         Index           =   2
         Left            =   1560
         TabIndex        =   4
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox aKeycodeText 
         Height          =   270
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   615
      End
      Begin VB.CheckBox CHKkey 
         Caption         =   "启用"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         ToolTipText     =   "开启改键"
         Top             =   180
         Width           =   735
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "="
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
         Index           =   3
         Left            =   2280
         TabIndex        =   41
         Top             =   1920
         Width           =   120
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "="
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
         Index           =   2
         Left            =   2280
         TabIndex        =   40
         Top             =   1560
         Width           =   120
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "="
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
         Index           =   1
         Left            =   840
         TabIndex        =   39
         Top             =   1920
         Width           =   120
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "="
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
         Index           =   0
         Left            =   840
         TabIndex        =   38
         Top             =   1560
         Width           =   120
      End
      Begin VB.Label Label4 
         Caption         =   "= 2"
         Height          =   255
         Index           =   5
         Left            =   2400
         TabIndex        =   37
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "= 1"
         Height          =   255
         Index           =   4
         Left            =   840
         TabIndex        =   36
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "= 5"
         Height          =   255
         Index           =   3
         Left            =   2400
         TabIndex        =   35
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "= 4"
         Height          =   255
         Index           =   2
         Left            =   840
         TabIndex        =   34
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "= 8"
         Height          =   255
         Index           =   1
         Left            =   2400
         TabIndex        =   33
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "= 7"
         Height          =   255
         Index           =   0
         Left            =   840
         TabIndex        =   32
         Top             =   480
         Width           =   375
      End
   End
   Begin VB.Timer GetTimeTimer 
      Interval        =   100
      Left            =   6960
      Top             =   2640
   End
   Begin VB.Frame Frame2 
      Caption         =   "喊话"
      Height          =   2235
      Left            =   2880
      TabIndex        =   28
      Top             =   0
      Width           =   3315
      Begin VB.CheckBox CHKXL 
         Caption         =   "显蓝"
         Height          =   255
         Left            =   840
         TabIndex        =   46
         ToolTipText     =   "此功能需要先在更多窗口设置游戏路径，并且在游戏启动前设置或设置后重启游戏才生效"
         Top             =   180
         Width           =   735
      End
      Begin RichTextLib.RichTextBox HhText4 
         Height          =   270
         Left            =   660
         TabIndex        =   23
         Top             =   1560
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   476
         _Version        =   393217
         BorderStyle     =   0
         MultiLine       =   0   'False
         OLEDragMode     =   0
         OLEDropMode     =   0
         TextRTF         =   $"Form1.frx":15162
      End
      Begin RichTextLib.RichTextBox HhText3 
         Height          =   270
         Left            =   660
         TabIndex        =   22
         Top             =   1200
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   476
         _Version        =   393217
         BorderStyle     =   0
         MultiLine       =   0   'False
         OLEDragMode     =   0
         OLEDropMode     =   0
         TextRTF         =   $"Form1.frx":151F1
      End
      Begin RichTextLib.RichTextBox HhText2 
         Height          =   270
         Left            =   660
         TabIndex        =   21
         Top             =   840
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   476
         _Version        =   393217
         BorderStyle     =   0
         MultiLine       =   0   'False
         OLEDragMode     =   0
         OLEDropMode     =   0
         TextRTF         =   $"Form1.frx":15280
      End
      Begin RichTextLib.RichTextBox HhText1 
         Height          =   270
         Left            =   660
         TabIndex        =   20
         Top             =   480
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   476
         _Version        =   393217
         BorderStyle     =   0
         MultiLine       =   0   'False
         OLEDragMode     =   0
         OLEDropMode     =   0
         TextRTF         =   $"Form1.frx":1530F
      End
      Begin VB.TextBox HhText5 
         Height          =   255
         Left            =   660
         TabIndex        =   24
         Top             =   1920
         Width           =   2475
      End
      Begin VB.CheckBox CHKHh 
         Caption         =   "启用"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         ToolTipText     =   "开启喊话"
         Top             =   180
         Width           =   735
      End
      Begin VB.CheckBox CK 
         Caption         =   "过-CK"
         Height          =   255
         Left            =   2400
         TabIndex        =   19
         ToolTipText     =   "过狠角色-CK，其它地图不建议启用"
         Top             =   180
         Width           =   745
      End
      Begin VB.Label Label3 
         Caption         =   "~ ="
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   44
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "9 ="
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   43
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "0 ="
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   31
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "8 ="
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   30
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "7 ="
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   29
         Top             =   480
         Width           =   495
      End
   End
   Begin VB.Image Image3 
      Height          =   240
      Left            =   7080
      Picture         =   "Form1.frx":1539E
      Top             =   3720
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image ImageC 
      Height          =   240
      Left            =   0
      Picture         =   "Form1.frx":15928
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image ImageD 
      Height          =   240
      Left            =   0
      Picture         =   "Form1.frx":15EB2
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   0
      Picture         =   "Form1.frx":1643C
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   2220
      Picture         =   "Form1.frx":169C6
      Top             =   1140
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImageB 
      Height          =   240
      Left            =   0
      Picture         =   "Form1.frx":17290
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image ImageA 
      Height          =   240
      Left            =   0
      Picture         =   "Form1.frx":1781A
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "游戏状态：未运行"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000011&
      Height          =   255
      Left            =   0
      TabIndex        =   42
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Menu file 
      Caption         =   "文件"
      Visible         =   0   'False
      Begin VB.Menu Show 
         Caption         =   "显示"
      End
      Begin VB.Menu Hide 
         Caption         =   "隐藏"
      End
      Begin VB.Menu SH1 
         Caption         =   "作弊"
      End
      Begin VB.Menu SH2 
         Caption         =   "更多"
      End
      Begin VB.Menu SH3 
         Caption         =   "联系作者"
      End
      Begin VB.Menu Exit 
         Caption         =   "退出"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal UFlags As Long) As Long
Dim ICOState As Integer
Dim GameState As Integer, CKSTATE As Integer
Dim LabelState As Integer
Dim StartState As Integer
Dim hHook As Long
Dim Yy() As Byte '保存声音文件
Dim Counter        As Long    '文件字节数
Dim PFdll() As Byte '保存皮肤dll
Dim PF() As Byte '保存皮肤文件
Dim ColorState1 As Integer
Dim ColorState2 As Integer
Dim ColorState3 As Integer
Dim ColorState4 As Integer

Private Sub aKeycodeText_Change(Index As Integer)                                            '如果是用鼠标右键删除文本,则aKeycode为0
    If aKeycodeText(Index) = "" Then aKeycode(Index) = 0
End Sub
Private Sub bKeycodeText_Change(Index As Integer)                                            '如果是用鼠标右键删除文本,则bKeycode为0
    If bKeycodeText(Index) = "" Then bKeycode(Index) = 0
End Sub
Private Sub aKeycodeText_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)       '设置aKeycode文本框内容
    If KeyCode > 47 And KeyCode < 58 Or KeyCode > 64 And KeyCode < 91 Then
        aKeycodeText(Index).MaxLength = 1
        aKeycodeText(Index).Text = Chr(KeyCode)
        aKeycode(Index) = KeyCode
Debug.Print aKeycode(Index)
    Else
        Select Case KeyCode
            Case 32
                aKeycodeText(Index).MaxLength = 5
                aKeycodeText(Index).Text = "Space"
                aKeycode(Index) = KeyCode
Debug.Print aKeycode(Index)
            Case 17
                aKeycodeText(Index).MaxLength = 4
                aKeycodeText(Index).Text = "CTRL"
                aKeycode(Index) = KeyCode
Debug.Print aKeycode(Index)
            Case 18
                aKeycodeText(Index).MaxLength = 3
                aKeycodeText(Index).Text = "ALT"
                aKeycode(Index) = KeyCode
Debug.Print aKeycode(Index)
            Case 192
                aKeycodeText(Index).MaxLength = 1
                aKeycodeText(Index).Text = "~"
                aKeycode(Index) = KeyCode
Debug.Print aKeycode(Index)
            Case 16
                aKeycodeText(Index).MaxLength = 5
                aKeycodeText(Index).Text = "SHIFT"
                aKeycode(Index) = KeyCode
Debug.Print aKeycode(Index)

            Case 189
                aKeycodeText(Index).MaxLength = 1
                aKeycodeText(Index).Text = "-"
                aKeycode(Index) = KeyCode
Debug.Print aKeycode(Index)
            Case 187
                aKeycodeText(Index).MaxLength = 1
                aKeycodeText(Index).Text = "="
                aKeycode(Index) = KeyCode
Debug.Print aKeycode(Index)
            Case 219
                aKeycodeText(Index).MaxLength = 1
                aKeycodeText(Index).Text = "["
                aKeycode(Index) = KeyCode
Debug.Print aKeycode(Index)
            Case 221
                aKeycodeText(Index).MaxLength = 1
                aKeycodeText(Index).Text = "]"
                aKeycode(Index) = KeyCode
Debug.Print aKeycode(Index)
            Case 220
                aKeycodeText(Index).MaxLength = 1
                aKeycodeText(Index).Text = "\"
                aKeycode(Index) = KeyCode
Debug.Print aKeycode(Index)
            Case 186
                aKeycodeText(Index).MaxLength = 1
                aKeycodeText(Index).Text = ";"
                aKeycode(Index) = KeyCode
Debug.Print aKeycode(Index)
            Case 222
                aKeycodeText(Index).MaxLength = 1
                aKeycodeText(Index).Text = "'"
                aKeycode(Index) = KeyCode
Debug.Print aKeycode(Index)
            Case 13
                aKeycodeText(Index).MaxLength = 5
                aKeycodeText(Index).Text = "ENTER"
                aKeycode(Index) = KeyCode
Debug.Print aKeycode(Index)
            Case 188
                aKeycodeText(Index).MaxLength = 1
                aKeycodeText(Index).Text = ","
                aKeycode(Index) = KeyCode
Debug.Print aKeycode(Index)
            Case 190
                aKeycodeText(Index).MaxLength = 1
                aKeycodeText(Index).Text = "."
                aKeycode(Index) = KeyCode
Debug.Print aKeycode(Index)
            Case 191
                aKeycodeText(Index).MaxLength = 1
                aKeycodeText(Index).Text = "/"
                aKeycode(Index) = KeyCode
Debug.Print aKeycode(Index)
            Case Else
                aKeycodeText(Index).MaxLength = 1
                aKeycodeText(Index).Text = ""
                aKeycode(Index) = 0
Debug.Print aKeycode(Index)
        End Select
    End If
End Sub

Private Sub bKeycodeText_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)                    '设置bKeycoed文本框内容
    If KeyCode > 47 And KeyCode < 58 Or KeyCode > 64 And KeyCode < 91 Then
        bKeycodeText(Index).MaxLength = 1
        bKeycodeText(Index).Text = Chr(KeyCode)
        bKeycode(Index) = KeyCode

    Else
        Select Case KeyCode
            Case 32
                bKeycodeText(Index).MaxLength = 5
                bKeycodeText(Index).Text = "Space"
                bKeycode(Index) = KeyCode

            Case 17
                bKeycodeText(Index).MaxLength = 4
                bKeycodeText(Index).Text = "CTRL"
                bKeycode(Index) = KeyCode

            Case 18
                bKeycodeText(Index).MaxLength = 3
                bKeycodeText(Index).Text = "ALT"
                bKeycode(Index) = KeyCode

            Case 192
                bKeycodeText(Index).MaxLength = 1
                bKeycodeText(Index).Text = "~"
                bKeycode(Index) = KeyCode

            Case 16
                bKeycodeText(Index).MaxLength = 5
                bKeycodeText(Index).Text = "SHIFT"
                bKeycode(Index) = KeyCode


            Case 189
                bKeycodeText(Index).MaxLength = 1
                bKeycodeText(Index).Text = "-"
                bKeycode(Index) = KeyCode

            Case 187
                bKeycodeText(Index).MaxLength = 1
                bKeycodeText(Index).Text = "="
                bKeycode(Index) = KeyCode

            Case 219
                bKeycodeText(Index).MaxLength = 1
                bKeycodeText(Index).Text = "["
                bKeycode(Index) = KeyCode

            Case 221
                bKeycodeText(Index).MaxLength = 1
                bKeycodeText(Index).Text = "]"
                bKeycode(Index) = KeyCode

            Case 220
                bKeycodeText(Index).MaxLength = 1
                bKeycodeText(Index).Text = "\"
                bKeycode(Index) = KeyCode

            Case 186
                bKeycodeText(Index).MaxLength = 1
                bKeycodeText(Index).Text = ";"
                bKeycode(Index) = KeyCode

            Case 222
                bKeycodeText(Index).MaxLength = 1
                bKeycodeText(Index).Text = "'"
                bKeycode(Index) = KeyCode

            Case 13
                bKeycodeText(Index).MaxLength = 5
                bKeycodeText(Index).Text = "ENTER"
                bKeycode(Index) = KeyCode

            Case 188
                bKeycodeText(Index).MaxLength = 1
                bKeycodeText(Index).Text = ","
                bKeycode(Index) = KeyCode

            Case 190
                bKeycodeText(Index).MaxLength = 1
                bKeycodeText(Index).Text = "."
                bKeycode(Index) = KeyCode

            Case 191
                bKeycodeText(Index).MaxLength = 1
                bKeycodeText(Index).Text = "/"
                bKeycode(Index) = KeyCode

            Case Else
                bKeycodeText(Index).MaxLength = 1
                bKeycodeText(Index).Text = ""
                bKeycode(Index) = 0
        End Select
    End If
End Sub

Private Sub Check1_Click()
If Check1.Value = 1 Then QD.Enabled = True
If Check1.Value = 0 Then QD.Enabled = False
End Sub

Private Sub CHKHh_Click()         '喊话的checkbox
    If CHKHh.Value = 1 Then
        HhTimer.Enabled = True
    Else
        HhTimer.Enabled = False
    End If
End Sub
Private Sub CHKMH_Click()
If CHKMH.Value = 0 Then
    Form3.Check2.Enabled = False
    Form3.Check3.Enabled = False
    Form3.Check4.Enabled = False
    Form3.Check5.Enabled = False
    Form3.Check6.Enabled = False
    Form3.Check7.Enabled = False
    If 获取魔兽版本 = "1.24E" Then Call 不改124E
    If 获取魔兽版本 = "1.24B" Then Call 不改124B
    If 获取魔兽版本 = "1.20E" Then Call 不改120E
Else
    Form3.Check2.Enabled = True
    Form3.Check3.Enabled = True
    Form3.Check4.Enabled = True
    Form3.Check5.Enabled = True
    Form3.Check6.Enabled = True
    Form3.Check7.Enabled = True
    GameState = 0
End If
End Sub

Private Sub CHKXL_Click() '显蓝
Dim Data() As Byte
If CHKXL.Value = 1 Then
    If Trim(Form2.Text5.Text) <> "" And Dir(Form2.Text5.Text) <> "" Then
        If Dir(Mid(Form2.Text5.Text, 1, Len(Form2.Text5.Text) - 8) & "wars.mixtape") = "" Then
            Data = LoadResData(139, "CUSTOM")
            Open Mid(Form2.Text5.Text, 1, Len(Form2.Text5.Text) - 8) & "wars.mixtape" For Binary As #1
            For Lon = 0 To UBound(Data)  '欲生成的文件大小
                Put #1, , Data(Lon) '释放文件
            Next
            Close #1
        End If
        
        If Dir(Mid(Form2.Text5.Text, 1, Len(Form2.Text5.Text) - 8) & "ManaColor_99uxi.com.txt") = "" Then
            Data = LoadResData(140, "CUSTOM")
            Open Mid(Form2.Text5.Text, 1, Len(Form2.Text5.Text) - 8) & "ManaColor_99uxi.com.txt" For Binary As #1
            For Lon = 0 To UBound(Data)  '欲生成的文件大小
                Put #1, , Data(Lon) '释放文件
            Next
            Close #1
        End If
    Else
        CHKXL.Value = 0
        MsgBox "显蓝功能需要在更多窗口设置游戏路径才能开启", vbInformation, "提示"
    End If
Else
    If Form2.Text5.Text <> "" And Dir(Form2.Text5.Text) <> "" Then
        If FindWindow(vbNullString, "Warcraft III") = 0 Then      '如果游戏没运行
            If Dir(Mid(Form2.Text5.Text, 1, Len(Form2.Text5.Text) - 8) & "wars.mixtape") <> "" Then Kill Mid(Form2.Text5.Text, 1, Len(Form2.Text5.Text) - 8) & "wars.mixtape"
        Else
            If Dir(Mid(Form2.Text5.Text, 1, Len(Form2.Text5.Text) - 8) & "wars.mixtape") <> "" Then
                MsgBox "游戏运行中,无法关闭显蓝功能", vbInformation, "提示"
                CHKXL.Value = 1
            End If
        End If
    End If
End If
End Sub

Private Sub CHKXX_Click()         '显血的CHECKBOX
    If CHKXX.Value = 1 Then
        If 获取魔兽版本 = "1.24E" Then Call 显血124E
        If 获取魔兽版本 = "1.24B" Then Call 显血124B
        If 获取魔兽版本 = "1.20E" Then Call 显血120E
    Else
        If 获取魔兽版本 = "1.24E" Then Call 不显血124E
        If 获取魔兽版本 = "1.24B" Then Call 不显血124B
        If 获取魔兽版本 = "1.20E" Then Call 不显血120E
    End If
End Sub

Private Sub CK_Click()
If CK.Value = 1 Then
    Timer5.Enabled = True
Else
    Timer5.Enabled = False
End If
End Sub

Private Sub Command2_Click() '退出按钮
    If hHook > 0 Then UnhookWindowsHookEx hHook           '程序退出时卸载钩子
    Call DelIcon                                          '删除系统托盘
    Open "C:\配置文件" For Output As #1               '保存当前设置
    For i = 1 To 10
        Print #1, aKeycodeText(i)
        Print #1, aKeycode(i)
    Next
    For i = 7 To 10
        Print #1, bKeycodeText(i)
        Print #1, bKeycode(i)
    Next
    Print #1, HhText1.Text
    Print #1, HhText2.Text
    Print #1, HhText3.Text
    Print #1, HhText4.Text
    Print #1, HhText5.Text
    Print #1, CHKkey.Value
    Print #1, CHKMH.Value
    Print #1, CHKHh.Value
    Print #1, CHKXX.Value
    Print #1, CK.Value
    Print #1, Check1.Value
    Print #1, CHKXL.Value
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

Private Sub Command3_Click()
Form3.Hide
Form3.Show
Form3.Timer1 = True
End Sub

Private Sub Exit_Click()                    '如果点击托盘菜单退出
    If hHook > 0 Then UnhookWindowsHookEx hHook           '程序退出时卸载钩子
    Call DelIcon                                          '删除系统托盘
    Open "C:\配置文件" For Output As #1               '保存当前设置
    For i = 1 To 10
        Print #1, aKeycodeText(i)
        Print #1, aKeycode(i)
    Next
    For i = 7 To 10
        Print #1, bKeycodeText(i)
        Print #1, bKeycode(i)
    Next
    Print #1, HhText1.Text
    Print #1, HhText2.Text
    Print #1, HhText3.Text
    Print #1, HhText4.Text
    Print #1, HhText5.Text
    Print #1, CHKkey.Value
    Print #1, CHKMH.Value
    Print #1, CHKHh.Value
    Print #1, CHKXX.Value
    Print #1, CK.Value
    Print #1, Check1.Value
    Print #1, CHKXL.Value
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



Private Sub Form_Activate()
Form1.CHKHh.Refresh
Form1.CHKkey.Refresh
Form1.CHKMH.Refresh
Form1.CHKXX.Refresh
Form3.Command1.Refresh
End Sub
Private Sub Form_Load()
Dim a As String, i As Integer, Lon As Long
StartState = 1
If App.PrevInstance = True Then               '如果程序已运行
    Open "C:\WINDOWS\system32\reinstant" For Output As #11
    Close #11
    End
End If
  
If Dir("C:\WINDOWS\system32\SkinH_VB6.dll") = "" Then    '如果皮肤dll文件不存在
        PFdll = LoadResData(102, "CUSTOM")
        Open "C:\WINDOWS\system32\SkinH_VB6.dll" For Binary As #8
        For Lon = 0 To 99250  '欲生成的文件大小，原文件99251字节
        Put #8, , PFdll(Lon) '释放文件
        Next
        Close #8
End If

If Dir("C:\WINDOWS\system32\PiFu.she") = "" Then    '如果皮肤文件不存在
        PF = LoadResData(103, "CUSTOM")
        Open "C:\WINDOWS\system32\PiFu.she" For Binary As #9
        For Lon = 0 To 30737  '欲生成的文件大小，原文件30738字节
        Put #9, , PF(Lon) '释放文件
        Next
        Close #9
End If

loadSkin  '加载皮肤

Form1.Icon = Image1.Picture
Form2.Icon = Image1.Picture
Form3.Icon = Image1.Picture
Form4.Icon = Image1.Picture
Form4.Caption = "版本:" & App.Major & "." & App.Minor & "." & App.Revision
Call TokenPrivileges                          '提权
Call AddIcon                                  '设置系统托盘
bKeycode(1) = 103   '小键盘数字键7
bKeycode(2) = 104   '小键盘数字键8
bKeycode(3) = 100   '小键盘数字键4
bKeycode(4) = 101   '小键盘数字键5
bKeycode(5) = 97    '小键盘数字键1
bKeycode(6) = 98    '小键盘数字键2

If Dir("C:\配置文件") = "" Then    '如果配置文件不存在
        PF = LoadResData(104, "CUSTOM")
        Open "C:\配置文件" For Binary As #9
        For Lon = 0 To UBound(PF)   '欲生成的文件大小
        Put #9, , PF(Lon) '释放文件
        Next
        Close #9
End If
If Dir("C:\配置文件") <> "" Then    '如果配置文件存在
    Open "C:\配置文件" For Input As #1
    For i = 1 To 10
        Line Input #1, a
On Error GoTo NewStart
        aKeycodeText(i).Text = a
        Line Input #1, a
On Error GoTo NewStart
        aKeycode(i) = Val(a)
    Next
    For i = 7 To 10
        Line Input #1, a
On Error GoTo NewStart
        bKeycodeText(i).Text = a
        Line Input #1, a
On Error GoTo NewStart
        bKeycode(i) = Val(a)
    Next
    Line Input #1, a
On Error GoTo NewStart
    HhText1.Text = a
    Line Input #1, a
On Error GoTo NewStart
    HhText2.Text = a
    Line Input #1, a
On Error GoTo NewStart
    HhText3.Text = a
    Line Input #1, a
On Error GoTo NewStart
    HhText4.Text = a
    Line Input #1, a
On Error GoTo NewStart
    HhText5.Text = a
    Line Input #1, a
On Error GoTo NewStart
    If Val(a) = 1 Then CHKkey.Value = 1                     '如果配置开启改键
    Line Input #1, a
On Error GoTo NewStart
    If Val(a) = 1 Then CHKMH.Value = 1                      '如果配置开图
    Call CHKMH_Click
    Line Input #1, a
On Error GoTo NewStart
    If Val(a) = 1 Then Call CHKHh_Click: CHKHh.Value = 1    '如果配置开启喊话
    Line Input #1, a
On Error GoTo NewStart
    If Val(a) = 1 Then CHKXX.Value = 1                      '如果配置开启显血
    Line Input #1, a
On Error GoTo NewStart
    If Val(a) = 1 Then CK.Value = 1: Timer5.Enabled = True  '如果配置开启过-CK
    Line Input #1, a
On Error GoTo NewStart
    If Val(a) = 1 Then Check1.Value = 1: QD.Enabled = True  '如果配置快速启动
    Line Input #1, a
On Error GoTo NewStart
    If Val(a) = 1 Then i = -2  '如果配置开启显蓝
    Line Input #1, a
On Error GoTo NewStart
    Form2.WidthTx.Text = a
    Line Input #1, a
On Error GoTo NewStart
    Form2.HeightTx.Text = a
 Line Input #1, a
On Error GoTo NewStart
    Form2.Text1.Text = a
 Line Input #1, a
On Error GoTo NewStart
    Form2.Text2.Text = a
 Line Input #1, a
On Error GoTo NewStart
    Form2.GMTX.Text = a
 Line Input #1, a
On Error GoTo NewStart
    Form2.Text5.Text = a
 Line Input #1, a
On Error GoTo NewStart
    Form3.Text1.Text = a
 Line Input #1, a
On Error GoTo NewStart
    Form3.Text2.Text = a
 Line Input #1, a
On Error GoTo NewStart
    Form3.Check2.Value = Val(a)
 Line Input #1, a
On Error GoTo NewStart
    Form3.Check3.Value = Val(a)
 Line Input #1, a
On Error GoTo NewStart
    Form3.Check4.Value = Val(a)
 Line Input #1, a
On Error GoTo NewStart
    Form3.Check5.Value = Val(a)
 Line Input #1, a
On Error GoTo NewStart
    Form3.Check6.Value = Val(a)
 Line Input #1, a
On Error GoTo NewStart
    Form3.Check7.Value = Val(a)
On Error GoTo NewStart
 Line Input #1, a
On Error GoTo NewStart
    Form1.top = Val(a)
 Line Input #1, a
On Error GoTo NewStart
    Form1.left = Val(a)
 Line Input #1, a
On Error GoTo NewStart
    Form2.top = Val(a)
 Line Input #1, a
On Error GoTo NewStart
    Form2.left = Val(a)
 Line Input #1, a
On Error GoTo NewStart
    Form3.top = Val(a)
 Line Input #1, a
On Error GoTo NewStart
    Form3.left = Val(a)
 Line Input #1, a
On Error GoTo NewStart
    Form4.top = Val(a)
 Line Input #1, a
On Error GoTo NewStart
    Form4.left = Val(a)
 Line Input #1, a
On Error GoTo NewStart
    Form4.Text2.Text = a
 Line Input #1, a
On Error GoTo NewStart
    Form4.Text3.Text = a
Close #1

End If

'为了显示完整字符
Me.Show
Form1.CHKHh.Refresh
Form1.CHKkey.Refresh
Form1.CHKMH.Refresh
Form1.CHKXX.Refresh
Form3.Command1.Refresh

If i = -2 Then '如果开启显蓝,i=-2
    CHKXL.Value = 1: Call CHKXL_Click    '如果配置开启显蓝
End If


    If i = -1 Then       'i不可能为-1
NewStart:
        Close #1
        For i = 1 To 10
            aKeycodeText(i).Text = ""
            aKeycode(i) = 0
        Next
        For i = 7 To 10
            bKeycodeText(i).Text = ""
            bKeycode(i) = 0
        Next
        HhText1.Text = ""
        HhText2.Text = ""
        HhText3.Text = ""
        HhText4.Text = ""
        HhText5.Text = ""
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Cancel = 0 Then
    Cancel = 1                            '点击窗体的关闭，则将窗体隐藏
    Form1.Hide
End If
End Sub


Private Sub GetTimeTimer_Timer()
If GetAsyncKeyState(13) <> 0 And GetAsyncKeyState(8) <> 0 And GetAsyncKeyState(46) <> 0 Then '如果按下回车、退格、DEL
    Shell "taskkill /f /im war3.exe", vbHide  '强制关闭进程
End If
If GetAsyncKeyState(vbKeyF5) <> 0 Then '监测F5键并报时
    GetWindowText GetForegroundWindow, WindowText, 255                  '获取前台窗体标题
    If left(WindowText, 12) = "Warcraft III" Then Call GetTime          '如果前台窗体标题为Warcraft III则报时
End If
If GetAsyncKeyState(36) <> 0 Then '如果按下HOME，则开图
    CHKMH.Value = 1
    Call CHKMH_Click
    Beep
End If
If GetAsyncKeyState(35) <> 0 Then '如果按下END，则取消开图
    CHKMH.Value = 0
    Call CHKMH_Click
    Beep
End If

End Sub

Private Sub HhText1_Change()
Dim OldSelStart As Long
If left(HhText1.Text, 3) = "SY|" Or left(HhText1.Text, 3) = "sy|" Or left(HhText1.Text, 3) = "Sy|" Or left(HhText1.Text, 3) = "sY|" Then
    If ColorState1 = 0 Then
        OldSelStart = HhText1.SelStart
        HhText1.Text = "SY|" & Mid(HhText1.Text, 4, Len(HhText1.Text) - 3)
        HhText1.SelStart = 0
        HhText1.SelLength = 3
        HhText1.SelColor = &H80000003
        HhText1.SelStart = OldSelStart
        HhText1.SelLength = 0
        HhText1.SelColor = 0
        ColorState1 = 1
    End If
Else
   If ColorState1 = 1 Then
       OldSelStart = HhText1.SelStart
       HhText1.SelStart = 0
       HhText1.SelLength = 3
       HhText1.SelColor = 0
       HhText1.SelStart = OldSelStart
       ColorState1 = 0
    End If
End If
End Sub

Private Sub HhText2_Change()
Dim OldSelStart As Long
If left(HhText2.Text, 3) = "SY|" Or left(HhText2.Text, 3) = "sy|" Or left(HhText2.Text, 3) = "Sy|" Or left(HhText2.Text, 3) = "sY|" Then
    If ColorState2 = 0 Then
        OldSelStart = HhText2.SelStart
        HhText2.Text = "SY|" & Mid(HhText2.Text, 4, Len(HhText2.Text) - 3)
        HhText2.SelStart = 0
        HhText2.SelLength = 3
        HhText2.SelColor = &H80000003
        HhText2.SelStart = OldSelStart
        HhText2.SelLength = 0
        HhText2.SelColor = 0
        ColorState2 = 1
    End If
Else
   If ColorState2 = 1 Then
       OldSelStart = HhText2.SelStart
       HhText2.SelStart = 0
       HhText2.SelLength = 3
       HhText2.SelColor = 0
       HhText2.SelStart = OldSelStart
       ColorState2 = 0
    End If
End If
End Sub

Private Sub HhText3_Change()
Dim OldSelStart As Long
If left(HhText3.Text, 3) = "SY|" Or left(HhText3.Text, 3) = "sy|" Or left(HhText3.Text, 3) = "Sy|" Or left(HhText3.Text, 3) = "sY|" Then
    If ColorState3 = 0 Then
        OldSelStart = HhText3.SelStart
        HhText3.Text = "SY|" & Mid(HhText3.Text, 4, Len(HhText3.Text) - 3)
        HhText3.SelStart = 0
        HhText3.SelLength = 3
        HhText3.SelColor = &H80000003
        HhText3.SelStart = OldSelStart
        HhText3.SelLength = 0
        HhText3.SelColor = 0
        ColorState3 = 1
    End If
Else
   If ColorState3 = 1 Then
       OldSelStart = HhText3.SelStart
       HhText3.SelStart = 0
       HhText3.SelLength = 3
       HhText3.SelColor = 0
       HhText3.SelStart = OldSelStart
       ColorState3 = 0
    End If
End If
End Sub
Private Sub HhText4_Change()
Dim OldSelStart As Long
If left(HhText4.Text, 3) = "SY|" Or left(HhText4.Text, 3) = "sy|" Or left(HhText4.Text, 3) = "Sy|" Or left(HhText4.Text, 3) = "sY|" Then
    If ColorState4 = 0 Then
        OldSelStart = HhText4.SelStart
        HhText4.Text = "SY|" & Mid(HhText4.Text, 4, Len(HhText4.Text) - 3)
        HhText4.SelStart = 0
        HhText4.SelLength = 3
        HhText4.SelColor = &H80000003
        HhText4.SelStart = OldSelStart
        HhText4.SelLength = 0
        HhText4.SelColor = 0
        ColorState4 = 1
    End If
Else
   If ColorState4 = 1 Then
       OldSelStart = HhText4.SelStart
       HhText4.SelStart = 0
       HhText4.SelLength = 3
       HhText4.SelColor = 0
       HhText4.SelStart = OldSelStart
       ColorState4 = 0
    End If
End If
End Sub

Private Sub HhText5_Change()
'If left(HhText5.Text, 3) = "sy|" Then HhText5.Text = "SY|" & Mid(HhText5.Text, 4, Len(HhText5.Text) - 3): HhText5.SelStart = Len(HhText5.Text)
'If left(HhText5.Text, 3) = "Sy|" Then HhText5.Text = "SY|" & Mid(HhText5.Text, 4, Len(HhText5.Text) - 3): HhText5.SelStart = Len(HhText5.Text)
'If left(HhText5.Text, 3) = "sY|" Then HhText5.Text = "SY|" & Mid(HhText5.Text, 4, Len(HhText5.Text) - 3): HhText5.SelStart = Len(HhText5.Text)
End Sub

Private Sub HhTimer_Timer()                                                                            '喊话相关Timer
    If GetAsyncKeyState(55) <> 0 Then   '按下7
       GetWindowText GetForegroundWindow, WindowText, 255              '获取前台窗体标题
        If left(WindowText, 12) = "Warcraft III" And ChatState = 0 Then         '如果前台窗体标题为Warcraft III并且不在聊天状态
            If HhText1.Text <> "" And 获取游戏状态 > 0 Then
                If left(HhText1.Text, 3) = "SY|" Then '如果对所有人喊话
                    SendString Mid(HhText1.Text, 4, Len(HhText1) - 3), 1
                Else '如果不对所有人喊话
                    SendString HhText1.Text
                End If
            End If
        ElseIf left(WindowText, 12) = "Warcraft III" And ChatState = 1 And 进入房间状态124E = 1 And 获取游戏状态 = 0 Then '如果前台窗体标题为Warcraft III并且在房间内
            If GetAsyncKeyState(18) <> 0 Then    'Alt
                If left(HhText1.Text, 3) = "SY|" Then SendString Mid(HhText1.Text, 4, Len(HhText1) - 3), 2
                If left(HhText1.Text, 3) <> "SY|" Then SendString HhText1.Text, 2
            End If
        End If
    End If
    If GetAsyncKeyState(56) <> 0 Then   '按下8
       GetWindowText GetForegroundWindow, WindowText, 255              '获取前台窗体标题
        If left(WindowText, 12) = "Warcraft III" And ChatState = 0 Then         '如果前台窗体标题为Warcraft III并且不在聊天状态
            If HhText2.Text <> "" And 获取游戏状态 > 0 Then
                If left(HhText2.Text, 3) = "SY|" Then '如果对所有人喊话
                    SendString Mid(HhText2.Text, 4, Len(HhText2) - 3), 1
                Else '如果不对所有人喊话
                    SendString HhText2.Text
                End If
            End If
        ElseIf left(WindowText, 12) = "Warcraft III" And ChatState = 1 And 进入房间状态124E = 1 And 获取游戏状态 = 0 Then '如果前台窗体标题为Warcraft III并且在房间内
            If GetAsyncKeyState(18) <> 0 Then    'Alt
                If left(HhText2.Text, 3) = "SY|" Then SendString Mid(HhText2.Text, 4, Len(HhText2) - 3), 2
                If left(HhText2.Text, 3) <> "SY|" Then SendString HhText2.Text, 2
            End If
        End If
    End If
    If GetAsyncKeyState(57) <> 0 Then   '按下9
       GetWindowText GetForegroundWindow, WindowText, 255              '获取前台窗体标题
        If left(WindowText, 12) = "Warcraft III" And ChatState = 0 Then         '如果前台窗体标题为Warcraft III并且不在聊天状态
            If HhText3.Text <> "" And 获取游戏状态 > 0 Then
                If left(HhText3.Text, 3) = "SY|" Then '如果对所有人喊话
                    SendString Mid(HhText3.Text, 4, Len(HhText3) - 3), 1
                Else '如果不对所有人喊话
                    SendString HhText3.Text
                End If
            End If
        ElseIf left(WindowText, 12) = "Warcraft III" And ChatState = 1 And 进入房间状态124E = 1 And 获取游戏状态 = 0 Then '如果前台窗体标题为Warcraft III并且在房间内
            If GetAsyncKeyState(18) <> 0 Then    'Alt
                If left(HhText3.Text, 3) = "SY|" Then SendString Mid(HhText3.Text, 4, Len(HhText3) - 3), 2
                If left(HhText3.Text, 3) <> "SY|" Then SendString HhText3.Text, 2
            End If
        End If
    End If
    If GetAsyncKeyState(48) <> 0 Then   '按下0
       GetWindowText GetForegroundWindow, WindowText, 255              '获取前台窗体标题
        If left(WindowText, 12) = "Warcraft III" And ChatState = 0 Then         '如果前台窗体标题为Warcraft III并且不在聊天状态
            If HhText4.Text <> "" And 获取游戏状态 > 0 Then
                If left(HhText4.Text, 3) = "SY|" Then '如果对所有人喊话
                    SendString Mid(HhText4.Text, 4, Len(HhText4) - 3), 1
                Else '如果不对所有人喊话
                    SendString HhText4.Text
                End If
            End If
        ElseIf left(WindowText, 12) = "Warcraft III" And ChatState = 1 And 进入房间状态124E = 1 And 获取游戏状态 = 0 Then '如果前台窗体标题为Warcraft III并且在房间内
            If GetAsyncKeyState(18) <> 0 Then    'Alt
                If left(HhText4.Text, 3) = "SY|" Then SendString Mid(HhText4.Text, 4, Len(HhText4) - 3), 2
                If left(HhText4.Text, 3) <> "SY|" Then SendString HhText4.Text, 2
            End If
        End If
    End If
     '按下~ 移到了改键模块
End Sub

Private Sub Hide_Click()
    Form1.Hide
    Form2.Hide
    Form3.Hide
    Form4.Hide
    Form3.Timer1 = False
End Sub

Private Sub Command1_Click()
    Form2.Hide
    Form2.Show
    'Form2.Label1.BackColor = RGB(255, 255, 125)
    'Form2.Label2.BackColor = RGB(255, 255, 125)
    'Form2.Label3.BackColor = RGB(255, 255, 125)
    'Form2.Label4.BackColor = RGB(255, 255, 125)
    'Form2.Label5.BackColor = RGB(255, 255, 125)
    'Form2.Label6.BackColor = RGB(255, 255, 125)
    
    'Form2.BackColor = RGB(255, 255, 125)
    'Form2.Command1.BackColor = RGB(255, 255, 125)
    'Form2.Command2.BackColor = RGB(255, 255, 125)
    'Form2.Command3.BackColor = RGB(255, 255, 125)
    'Form2.Command4.BackColor = RGB(255, 255, 125)
    'Form2.Command5.BackColor = RGB(255, 255, 125)
    'Form2.Command6.BackColor = RGB(255, 255, 125)
    'Form2.Command7.BackColor = RGB(255, 255, 125)
    'Form2.Command8.BackColor = RGB(255, 255, 125)
    'Form2.BackColor = RGB(255, 255, 125)
    'Form2.Label3.BackColor = RGB(255, 255, 125)
    Form2.WidthTx = Screen.Width / Screen.TwipsPerPixelX
    Form2.HeightTx = Screen.Height / Screen.TwipsPerPixelY
End Sub

Private Sub QD_Timer()
Dim hwnd, PID
If GetAsyncKeyState(187) < 0 And GetAsyncKeyState(189) < 0 Then  '按下-+
    hwnd = FindWindow(vbNullString, "Warcraft III")
    If hwnd = 0 Then '如果游戏没运行
        If UCase(right(Form2.Text5.Text, 8)) = "WAR3.EXE" Then
            Shell Form2.Text5.Text
        Else
            Form2.Hide
            Form2.Show
            MsgBox "请先获取游戏路径！", vbInformation, "提示"
        End If
    End If
End If

End Sub

Private Sub SH1_Click()
Form3.Hide
Form3.Show
Form1.CHKHh.Refresh
Form1.CHKkey.Refresh
Form1.CHKMH.Refresh
Form1.CHKXX.Refresh
Form3.Command1.Refresh
Form3.Timer1 = True
End Sub

Private Sub SH2_Click()
Form2.Hide
Form2.Show
Form1.CHKHh.Refresh
Form1.CHKkey.Refresh
Form1.CHKMH.Refresh
Form1.CHKXX.Refresh
Form3.Command1.Refresh
End Sub

Private Sub SH3_Click()
Form4.Hide
Form4.Show
If InternetGetConnectedState(0&, 0&) Then '如果网络已连接
    Form4.Text1.Text = GetUrlFile("http://faithdmc.host166.web522.com/war3GGB")
Else '如果网络未连接
    Form4.Text1.Text = "网络未连接"
End If
Form1.CHKHh.Refresh
Form1.CHKkey.Refresh
Form1.CHKMH.Refresh
Form1.CHKXX.Refresh
Form3.Command1.Refresh
End Sub
Private Sub Show_Click()
Form1.Hide
Form1.Show
Form1.CHKHh.Refresh
Form1.CHKkey.Refresh
Form1.CHKMH.Refresh
Form1.CHKXX.Refresh
Form3.Command1.Refresh
End Sub

Private Sub Timer1_Timer()                          '设置动态托盘
    If ICOState = 0 Then
        ChangeIcon Form1.ImageB.Picture.Handle
        ICOState = 1
    ElseIf ICOState = 1 Then
        ChangeIcon Form1.ImageC.Picture.Handle
        ICOState = 2
    ElseIf ICOState = 2 Then
        ChangeIcon Form1.ImageD.Picture.Handle
        ICOState = 3
    ElseIf ICOState = 3 Then
        ChangeIcon Form1.ImageA.Picture.Handle
        ICOState = 0
    End If
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)               '响应托盘事件
 Dim lMsg As Single
  lMsg = X / Screen.TwipsPerPixelX  '屏幕横向(X)分辨率
  If lMsg = WM_RBUTTONUP Then PopupMenu file   '右键点击则弹出菜单
  If lMsg = WM_LBUTTONUP Then Form1.Hide: Form1.Show      '左键点击则显示窗体
End Sub

Private Sub Timer2_Timer() '用于提示游戏开始
If FindWindow(vbNullString, "Warcraft III") > 0 Then      '如果游戏已运行
    If StartState = 0 Then
    If 获取游戏状态 > 0 Then
        If CHKXX.Value = 1 Then '显血
            Select Case 获取魔兽版本
            Case "1.24E": Call 显血124E
            Case "1.24B": Call 显血124B
            Case "1.20E": Call 显血120E
            End Select
        End If
        
        GetWindowText GetForegroundWindow, WindowText, 255                  '获取前台窗体标题
        If left(WindowText, 12) <> "Warcraft III" Then
        Call Remind '提醒游戏开始了
        Yy = LoadResData(101, "WAVE")
        Open "C:\游戏开始了.wav" For Binary As #2
        For Counter = 0 To 194859 '欲生成的文件大小，原文件194860字节
        Put #2, , Yy(Counter) '释放文件
        Next
        Close #2
        sndPlaySound "C:\游戏开始了.wav", 1 '播放声音文件
        StartState = 1
        Delay 3000 '延迟3秒，为了避免音频还在播放导致删除出错
        Kill "C:\游戏开始了.wav"
        Else: StartState = 1
        End If
    End If
    End If
    If 获取游戏状态 = 0 Then
        If StartState = 1 Then StartState = 0
        If Timer1.Enabled = False Then Timer1.Enabled = True
    End If
    If Timer1.Enabled = False Then
       GetWindowText GetForegroundWindow, WindowText, 255                  '获取前台窗体标题
       If left(WindowText, 12) = "Warcraft III" Then Timer1.Enabled = True
    End If
End If
End Sub

Private Sub Timer3_Timer() '间隔为1秒
If FindWindow(vbNullString, "Warcraft III") > 0 Then      '如果游戏已运行
    Form2.Command1.Enabled = False
    Form2.Command1.Caption = "游戏中，不能改"
    If GameState = 0 And CHKMH.Value = 1 Then            '重新启动游戏后，如果MH的CHECKBOX为1则开启全图
        If 获取魔兽版本 = "1.24E" Then Call 改124E
        If 获取魔兽版本 = "1.24B" Then Call 改124B
        If 获取魔兽版本 = "1.20E" Then Call 改120E
'        If 获取魔兽版本 = "1.21" Then Call 改121
        GameState = 1
    End If
    If LabelState = 0 Then
        Label6.Caption = "游戏版本：" & 获取魔兽版本
        Label6.ForeColor = &H80000001
        LabelState = 1
    End If
    If CHKkey.Value = 1 Or CHKHh.Value = 1 Then  '如果开启改键或喊话则挂载钩子
        If hHook = 0 Then hHook = SetWindowsHookEx(WH_KEYBOARD_LL, AddressOf MyKBHook, App.hInstance, 0)
    End If
    If CHKkey.Value = 0 And CHKHh.Value = 0 Then '如果取消改键并且取消喊话则卸载钩子
        If hHook > 0 Then UnhookWindowsHookEx hHook: hHook = 0
    End If
Else                                                  '如果游戏已关闭
    Form2.Command1.Enabled = True
    Form2.Command1.Caption = "设置魔兽分辨率"
    
    If LabelState = 1 Then
        Label6.Caption = "游戏状态：未运行"
        Label6.ForeColor = &H80000011
        LabelState = 0
    End If
    If GameState = 1 Then GameState = 0
    If hHook > 0 Then
        UnhookWindowsHookEx hHook
        hHook = 0
    End If
    If Timer1.Enabled = False Then Timer1.Enabled = True
End If
End Sub

Private Sub Timer4_Timer()
If Dir("C:\WINDOWS\system32\reinstant") <> "" Then    '如果reinstant文件存在,为了重复打开程序时显示已打开的窗体
    Form1.Show
    Kill "C:\WINDOWS\system32\reinstant"
End If
End Sub

Private Sub Timer5_Timer() '过-CK
Dim hwnd As Long, Handle As Long, PID As Long, Addr As Long
If CHKMH.Value = 1 Then
    hwnd = FindWindow(vbNullString, "Warcraft III")
    GetWindowThreadProcessId hwnd, PID
    Handle = OpenProcess(PROCESS_ALL_ACCESS, False, PID)
    ReadProcessMemory Handle, ByVal Wargamedll + &HA8C058, Addr, 4, 0&
    If CKSTATE = 0 And Addr = 1 Then
        Call 不改124E
        CKSTATE = 1
        Delay 6000
    End If
    If CKSTATE = 1 And Addr = 0 Then
        Call 改124E
        CKSTATE = 0
    End If
End If
End Sub
