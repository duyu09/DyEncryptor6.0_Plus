VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DyEncryptor6.0"
   ClientHeight    =   8970
   ClientLeft      =   45
   ClientTop       =   645
   ClientWidth     =   11310
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   8970
   ScaleWidth      =   11310
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer TimerGUI 
      Interval        =   1800
      Left            =   9120
      Top             =   5040
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   120
      Max             =   255
      MousePointer    =   1  'Arrow
      TabIndex        =   38
      Top             =   6120
      Value           =   255
      Width           =   2055
   End
   Begin VB.Timer TimerTemprt 
      Enabled         =   0   'False
      Interval        =   15
      Left            =   8160
      Top             =   5040
   End
   Begin VB.CommandButton Command10 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   8160
      TabIndex        =   37
      Top             =   6480
      Width           =   375
   End
   Begin VB.CommandButton Command9 
      Caption         =   "显示字符"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   7.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2760
      TabIndex        =   36
      Top             =   3960
      Width           =   840
   End
   Begin VB.CommandButton Command8 
      Caption         =   "显示字符"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   7.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2760
      TabIndex        =   35
      Top             =   2880
      Width           =   840
   End
   Begin VB.CommandButton Command6 
      Caption         =   "属性"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9840
      TabIndex        =   33
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton Command6_2 
      Caption         =   "Command6"
      Height          =   495
      Left            =   10080
      TabIndex        =   32
      Top             =   3960
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Timer Timer4 
      Interval        =   10
      Left            =   8640
      Top             =   5040
   End
   Begin DyEncryptor.MorphTextBox Text1 
      Height          =   375
      Left            =   480
      TabIndex        =   29
      Top             =   3240
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DefaultColor1   =   8421504
      DefaultColor2   =   16777215
      FocusColor2     =   16777152
      PasswordChar    =   "*"
   End
   Begin VB.TextBox Text4 
      Height          =   1935
      Left            =   2400
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      OLEDropMode     =   1  'Manual
      ScrollBars      =   3  'Both
      TabIndex        =   8
      ToolTipText     =   "加密系统运行日志"
      Top             =   6960
      Width           =   8775
   End
   Begin VB.CommandButton Command5 
      Caption         =   "历史记录"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2760
      TabIndex        =   25
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CheckBox Check2 
      Caption         =   "生成EXE解密文件"
      Height          =   300
      Left            =   120
      TabIndex        =   27
      ToolTipText     =   "生成EXE解密文件"
      Top             =   6960
      Width           =   2415
   End
   Begin VB.CheckBox Check1 
      Caption         =   "完成后关机"
      Height          =   300
      Left            =   120
      TabIndex        =   24
      ToolTipText     =   "完成后关闭计算机"
      Top             =   7320
      Width           =   2415
   End
   Begin VB.CommandButton Command4 
      Caption         =   "设置"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   120
      TabIndex        =   22
      ToolTipText     =   "导出日志到输出目录"
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "导出日志"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1440
      TabIndex        =   18
      ToolTipText     =   "导出日志到输出目录"
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   2400
      TabIndex        =   15
      ToolTipText     =   "输出目录选择"
      Top             =   5760
      Width           =   3015
      Begin VB.OptionButton Option4 
         Caption         =   "自定义"
         Height          =   300
         Left            =   1680
         TabIndex        =   17
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton Option3 
         Caption         =   "源目录"
         Height          =   300
         Left            =   240
         TabIndex        =   16
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.TextBox Text5 
      Enabled         =   0   'False
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   3600
      TabIndex        =   14
      Top             =   6480
      Width           =   4455
   End
   Begin VB.Timer Timer3 
      Interval        =   25
      Left            =   10080
      Top             =   5040
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   60
      Left            =   9600
      Top             =   5040
   End
   Begin VB.CommandButton Command2 
      Caption         =   "一键处理"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8640
      TabIndex        =   12
      Top             =   5880
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   5640
      TabIndex        =   9
      ToolTipText     =   "模式选择"
      Top             =   5760
      Width           =   2895
      Begin VB.OptionButton Option2 
         Caption         =   "解密"
         Height          =   300
         Left            =   1560
         TabIndex        =   11
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "加密"
         Height          =   300
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "浏览文件"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   9960
      TabIndex        =   5
      ToolTipText     =   "浏览以选择文件"
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   390
      IMEMode         =   3  'DISABLE
      Left            =   4080
      OLEDropMode     =   1  'Manual
      TabIndex        =   4
      Top             =   1200
      Width           =   5775
   End
   Begin VB.TextBox Text2m 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   4800
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   4200
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.TextBox Text1m 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   4920
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   4800
      Visible         =   0   'False
      Width           =   3135
   End
   Begin DyEncryptor.MorphTextBox Text2 
      Height          =   375
      Left            =   480
      TabIndex        =   30
      Top             =   4320
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DefaultColor1   =   8421504
      DefaultColor2   =   16777215
      PasswordChar    =   "*"
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "100.0%"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   8.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   40
      Top             =   5820
      Width           =   1215
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "透明度设置："
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   8.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   39
      Top             =   5820
      Width           =   2055
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "加密系统安全键盘已启动防护"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   360
      TabIndex        =   34
      Top             =   2160
      Width           =   3255
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "大写键已开启 "
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   360
      TabIndex        =   31
      Top             =   5040
      Width           =   3255
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright (C) 2016-2022 DUYU."
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   7.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7440
      TabIndex        =   19
      ToolTipText     =   "单击此处查看版权声明，齐鲁工业大学 软件工程（开发）21-1班 杜宇(202103180009) 保留所有权利。"
      Top             =   5400
      Width           =   3255
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   375
      Left            =   9600
      TabIndex        =   28
      Top             =   4440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "version"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   6.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   26
      ToolTipText     =   "单击此处查看版权声明，齐鲁工业大学 软件工程（开发）21-1班 杜宇(202103180009) 保留所有权利。"
      Top             =   8640
      Width           =   2415
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "销毁源文件"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   23
      Top             =   6480
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   0
      Picture         =   "Form1.frx":C84A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11295
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "您可以单击“浏览文件”按钮，也可以拖放文件到此处"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   5640
      OLEDropMode     =   1  'Manual
      TabIndex        =   21
      ToolTipText     =   "拖放文件到此处"
      Top             =   3360
      Width           =   4215
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   5  'Dash-Dot-Dot
      Height          =   3735
      Left            =   4200
      Shape           =   4  'Rounded Rectangle
      Top             =   1920
      Width           =   6855
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "111"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   20
      ToolTipText     =   "当前时间"
      Top             =   8280
      Width           =   3135
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "输出目录："
      Height          =   495
      Left            =   2400
      TabIndex        =   13
      Top             =   6480
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      Height          =   375
      Left            =   1320
      TabIndex        =   7
      Top             =   7680
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "已用时间："
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   7680
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "确认密码："
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "设置密码："
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Image Image2 
      Height          =   3780
      Left            =   120
      Picture         =   "Form1.frx":1813C
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   3855
   End
   Begin VB.Menu file 
      Caption         =   "文件(F)"
      Begin VB.Menu addf 
         Caption         =   "添加文件(A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu dest 
         Caption         =   "销毁文件(D)"
         Shortcut        =   ^D
      End
   End
   Begin VB.Menu setting 
      Caption         =   "设置(S)"
      Begin VB.Menu setcenter 
         Caption         =   "设置中心(C)"
         Shortcut        =   ^C
      End
      Begin VB.Menu wincol 
         Caption         =   "窗体颜色设置(W)"
         Shortcut        =   ^W
      End
   End
   Begin VB.Menu about 
      Caption         =   "关于(A)"
   End
   Begin VB.Menu quit 
      Caption         =   "退出(Q)"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 'Copyright (c) 2016-2022 DUYU.  杜宇保留所有权利。
 'All Rights Reserved.
 '2016-04-05.  version:1.0.
 '2018-10-13.  version:1.0_Plus.
 '2019-06-02.  version:2.0.
 '2020-02-26.  version:3.0.
 '2020-03-26.  version:4.0.
 '2021-08-29.  version:5.0.
 '2021-09-17.  version:5.0_Plus.
 '2021-10-28.  version:6.0.
 '2022-01-01.  version:6.0_Plus.
 '2016-2018 山东省济南实验初级中学（泉景校区）- 计算机社团小组 No.04
 '2018-2020 济南市历城第二中学 - 信息学奥赛训练队 No.5522007 No.5531028
 '2020-2022 齐鲁工业大学 - ICPC-ACM集训队 ， 橙果工作室 No.202103180009
 
 'Form1.frm
 '气泡提示效果下载自http://www.codefans.net，在此声明。
 '
 '说明：DyEnc6.0没有更新涉及文件附属数据的内容，故加密文件的文件头仍然采用Version=5
 
 '窗体透明
 Private Type POINTAPI
        x As Long
        y As Long
End Type
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const WS_EX_LAYERED = &H80000
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const LWA_COLORKEY = &H1
 
' 引用 CMultiToolTips 类
Dim TT As New CMultiToolTips
Private dlg As CCommonDialog
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long '判断键盘大写键是否开启
Const VK_CAPITAL = &H14
Function CapitalStatus() As Boolean '判断键盘大写键是否开启
    Dim tKeyboardMap(255) As Byte
    Dim bResult As Long
    bResult = GetKeyboardState(tKeyboardMap(0))
    If tKeyboardMap(VK_CAPITAL) And 1 = 1 Then CapitalStatus = True
End Function
Private Function Modd(a1 As Double, a2 As Double) As Double
Modd = a1 - a2 * Fix(a1 / a2)
End Function

Private Function GetFileName(FileAllName As String) As String
Dim a As String
a = FileAllName
For b = 1 To Len(a)
  If Mid(a, b, 1) = "\" Then
     C = b
  End If
Next b
GetFileName = Right(a, Len(a) - C)
End Function

Private Function GetFilePath(FileAllName As String) As String
Dim a As String, b As Integer, C As Integer
a = FileAllName
For b = 1 To Len(a)
   If Mid(a, b, 1) = "\" Then
   C = b
   End If
Next b
GetFilePath = Left(a, C - 1)
End Function

Private Function GetEncWord(OriWord As String) As String
Dim sT(0 To 100) As String, ia As Integer
Dim xpaa As String, jzz As Integer, izz As Integer

For ia = 1 To 94
sT(ia) = Chr(ia + 31)
Next
sT(95) = Chr(1)
sT(96) = Chr(2)
sT(97) = Chr(3)
sT(98) = Chr(4)
sT(99) = Chr(5)
sT(100) = Chr(6)
sT(0) = Chr(7)

For izz = 1 To Len(OriWord)
For jzz = 1 To 95
If sT(jzz) = Mid(OriWord, izz, 1) Then
Exit For
End If

Next jzz
Next izz
Dim sopz As String, stwz As String, az As Integer, bz As Integer, RAz As Integer, cz As String, dz As Integer, ez As String
For az = 1 To Len(OriWord)
For bz = 1 To 94
If sT(bz) = Mid(OriWord, az, 1) Then
Randomize
RAz = CInt(4.1 - 3 * Rnd)
sopz = sopz + sT(bz + RAz)
stwz = stwz + sT(RAz)
End If
Next bz
Next az
ez = CInt(9 - 9 * Rnd)
For dz = 1 To ez
Randomize
cz = cz + sT(CInt(90 - 89 * Rnd))
Next dz
xpaa = stwz + StrReverse(sopz) + cz + ez

GetEncWord = xpaa
End Function


Private Function UnEncWord(EncWord As String) As String
Dim sT(0 To 100) As String, ia As Long
Dim s As String, pas As String
Dim av As String, bpv As String, cpv As String, iav As Long, dv As Long, bmv As String, cmv As String, ev As Long, fv As Long
Dim a As String, bp As String, CP As String, d As Long, bm As String, cm As String, e As Long, f As Long

pas = EncWord

For ia = 1 To 94
sT(ia) = Chr(ia + 31)
Next
sT(95) = Chr(1)
sT(96) = Chr(2)
sT(97) = Chr(3)
sT(98) = Chr(4)
sT(99) = Chr(5)
sT(100) = Chr(6)
sT(0) = Chr(7)

av = Left(pas, Len(pas) - 1 - Val(Right(pas, 1)))
bpv = StrReverse(Right(av, Len(av) / 2))
cpv = Left(av, Len(av) \ 2)
For dv = 1 To Len(av) \ 2
For ev = 1 To 4
If sT(ev) = Mid(cpv, dv, 1) Then
For fv = 1 To 95
If sT(fv) = Mid(bpv, dv, 1) Then
bmv = bmv + sT(fv - ev)
Exit For
End If
Next
End If
Next
Next

UnEncWord = bmv
End Function

Private Sub about_Click()
Form2.Show
End Sub

Private Sub addf_Click()
Command1_Click
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
   MsgB ("声明：该功能正处于试验阶段，对于部分操作系统和机型可能不适用。生成自解密EXE的过程涉及数据压缩，WinRar提供技术支持（Copyright (c) 1993-2019 Alexander Roshal）")
End If
End Sub

Private Sub Command1_Click()
Dim OFN As OPENFILENAME
Dim rtn As String, fp As String, xnk As String

If Dir(selP, vbDirectory) <> "" Then
   xnk = selP
Else
   xnk = App.Path
End If
OFN.lStructSize = Len(OFN)
OFN.hwndOwner = Me.hwnd
OFN.hInstance = App.hInstance
OFN.lpstrFilter = "所有文件(*.*)"
OFN.lpstrFile = Space(254)
OFN.nMaxFile = 255
OFN.lpstrFileTitle = Space(254)
OFN.nMaxFileTitle = 255
OFN.lpstrInitialDir = xnk
OFN.lpstrTitle = "请选择被处理的文件 - DyEncryptor"
OFN.Flags = 6148
rtn = GetOpenFileName(OFN)
If rtn >= 1 Then
   Text3.Text = OFN.lpstrFile
   Open App.Path & "\DyEncGUI5.0.config" For Output As #10
        Print #10, GetFilePath(Text3.Text)
   Close #10
End If
End Sub

Private Sub Command10_Click()
If Option4.Value = False Then
   MsgB ("请首先选择自定义输出目录模式，再修改输出目录。" & vbCrLf & "您设置的自定义输出目录为：" & Text5.Text)
   Exit Sub
End If
Set sh = CreateObject("Shell.Application")
Set fd = sh.BrowseForFolder(Me.hwnd, "DyEncryptor" & App.Major & "." & App.Minor & " - 请选择输出目录", 0)
'Me.hWnd是“选择文件夹”对话框的父窗口句柄，不关闭对话框不能返回父窗口。改为0没这个效果。
'第三个参数0改为512不显示“新建文件夹”按钮
If TypeName(fd) = "Folder3" Then Text5.Text = fd.Self.Path
End Sub

Private Sub Command2_Click()
starttime = Now
Timer2.Enabled = True

DoEvents
Text4.Text = Text4.Text & "[" & Format(Hour(Now), "00") & ":" & Format(Minute(Now), "00") & ":" & Format(Second(Now), "00") & "] 定位文件：" & Text3.Text & vbCrLf
On Error Resume Next
If Dir(Text3.Text) = "" Then
  MsgB ("您选择的文件不存在。")
  Timer2.Enabled = False
  Text4.Text = ""
  Label5.Caption = "00:00:00"
  Command2.SetFocus
  Exit Sub
End If

DoEvents
If Text3.Text = "" Then
  MsgB ("请选择文件。")
  Timer2.Enabled = False
  Text4.Text = ""
  Label5.Caption = "00:00:00"
  Command1_Click
  Command2.SetFocus
  Exit Sub
End If

DoEvents
Dim judge1 As Boolean
 If Option4.Value = True Then
 Text4.Text = Text4.Text & "[" & Format(Hour(Now), "00") & ":" & Format(Minute(Now), "00") & ":" & Format(Second(Now), "00") & "] 定位输出目录：" & Text5.Text & vbCrLf
 OutputDir = Text5.Text
    If Dir(Text5.Text, vbDirectory) = "" Or Text5.Text = "" Then
       judge1 = False
    Else
       judge1 = True
    End If
 Else
 Text4.Text = Text4.Text & "[" & Format(Hour(Now), "00") & ":" & Format(Minute(Now), "00") & ":" & Format(Second(Now), "00") & "] 定位输出目录：" & GetFilePath(Text3.Text) & vbCrLf
 OutputDir = GetFilePath(Text3.Text)
    If Dir(GetFilePath(Text3.Text), vbDirectory) = "" Then
       judge1 = False
    Else
       judge1 = True
    End If
 End If
If judge1 = False Then
  MsgB ("您输入的目录路径不存在。")
  Timer2.Enabled = False
  Text4.Text = ""
  Label5.Caption = "00:00:00"
  Command2.SetFocus
  Exit Sub
End If

DoEvents
Text4.Text = Text4.Text & "[" & Format(Hour(Now), "00") & ":" & Format(Minute(Now), "00") & ":" & Format(Second(Now), "00") & "] 检查运行环境"
If exitproc("Dy_EncCore.exe") = True Then
  Text4.Text = Text4.Text & "：异常" & vbCrLf
  MsgB ("加密核心组件正在运行，请将其关闭。")
  Timer2.Enabled = False
  Text4.Text = ""
  Label5.Caption = "00:00:00"
  Command2.SetFocus
  Exit Sub
End If

DoEvents
If Text1.Text = "" Then
  Text4.Text = Text4.Text & "：异常" & vbCrLf
  MsgB ("请输入密码。")
  Timer2.Enabled = False
  Text4.Text = ""
  Label5.Caption = "00:00:00"
  Command2.SetFocus
  Exit Sub
End If

DoEvents
If Option1.Value = True Then
    If Text1.Text <> Text2.Text Then
      Text4.Text = Text4.Text & "：异常" & vbCrLf
      MsgB ("您设置的密码与确认密码不一致，请重新输入。")
      Text2.Text = ""
      Timer2.Enabled = False
      Text4.Text = ""
      Label5.Caption = "00:00:00"
      Command2.SetFocus
      Exit Sub
    End If
End If

DoEvents
If Len(Text1.Text) > 25 Then
  Text4.Text = Text4.Text & "：异常" & vbCrLf
  MsgB ("您最多可输入25位密码。")
  Timer2.Enabled = False
  Text4.Text = ""
  Label5.Caption = "00:00:00"
  Command2.SetFocus
  Exit Sub
End If

DoEvents
If Dir(App.Path & "\Dy_EncCore.exe") = "" Then
  Text4.Text = Text4.Text & "：异常" & vbCrLf
  MsgB ("核心组件丢失，无法执行操作。")
  Timer2.Enabled = False
  Text4.Text = ""
  Label5.Caption = "00:00:00"
  Command2.SetFocus
  Exit Sub
End If

DoEvents
If Dir(OutputDir & "\" & GetFileName(Text3.Text) & ".Dyenc_Output") <> "" Then
  Text4.Text = Text4.Text & "：异常" & vbCrLf
  MsgB ("已存在" & OutputDir & "\" & GetFileName(Text3.Text) & ".Dyenc_Output，请将其移动到其他目录")
  Timer2.Enabled = False
  Text4.Text = ""
  Label5.Caption = "00:00:00"
  Command2.SetFocus
  Exit Sub
End If

On Error Resume Next
Kill Environ("temp") & "\finish.dyenc"

DoEvents
Text4.Text = Text4.Text & "：正常" & vbCrLf

Timer3.Enabled = False
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text5.Enabled = False
Command1.Enabled = False
Command2.Enabled = False
Frame1.Enabled = False
Frame2.Enabled = False
Command2.SetFocus
'*********************************加密代码***************************************
'20字节：固定头（[DyEncFile]Version=5）
'1字节：随机头长度（取值范围1-255）
'1字节：密码密文长度（小于等于255）
'？字节：随机头
'1字节：密钥（key）
'？字节：密码密文
'？字节：加密文件数据
DoEvents
Text4.Text = Text4.Text & "[" & Format(Hour(Now), "00") & ":" & Format(Minute(Now), "00") & ":" & Format(Second(Now), "00") & "] 生成附属数据：完成 " & vbCrLf
If Option1.Value = True Then
   Dim CommandArg As String, key As Integer, LTotalhead As Long 'LTotalhead文件头总长度，LEncPass密码密文长度，LRndhead随机头长度
   Dim LEncPass As Integer, LRndhead As Integer, EncP As String 'EncP密码密文
   
   Randomize                            '生成密钥，随机头长度，密码密文
   key = Fix(1 + Rnd() * 255)
   Randomize
   LRndhead = Fix(1 + Rnd() * 255)
   EncP = GetEncWord(Text1.Text)
   LEncPass = Len(EncP)
   LTotalhead = 23 + LRndhead + LEncPass
   
   Dim HeadData(0 To 539) As Byte       '生成头数据于数组
   Dim a() As Byte
   a() = StrConv("[DyEncFile]Version=5", vbFromUnicode)
   For b = 0 To 19
       DoEvents
       HeadData(b) = a(b)
   Next b
   HeadData(20) = CByte(LRndhead)
   HeadData(21) = CByte(LEncPass)
   For C = 22 To 21 + LRndhead
       DoEvents
       Randomize
       HeadData(C) = CByte(Fix(1 + Rnd() * 255))
   Next C
   HeadData(22 + LRndhead) = CByte(key)
   Dim Esa As Long, ar As Byte
   Open Environ("temp") & "\TempFile.dyenc" For Output As #5
        Print #5, EncP
   Close #5
   Open Environ("temp") & "\TempFile.dyenc" For Binary As #6
        Esa = 1
        For d = 23 + LRndhead To 22 + LRndhead + LEncPass
            DoEvents
            Get #6, Esa, ar
            HeadData(d) = ar
            Esa = Esa + 1
        Next d
   Close #6
   On Error Resume Next
   Kill Environ("temp") & "\TempFile.dyenc"
   DoEvents
   Text4.Text = Text4.Text & "[" & Format(Hour(Now), "00") & ":" & Format(Minute(Now), "00") & ":" & Format(Second(Now), "00") & "] 开始处理：运行 " & App.Path & "\Dy_EncCore.exe" & vbCrLf
   CommandArg = Chr(34) & Text3.Text & Chr(34) & " " & Chr(34) & OutputDir & "\" & GetFileName(Text3.Text) & ".Dyenc_Output" & Chr(34) & " " & CStr(key) & " " & CStr(LTotalhead) & " " & "0"
   ShellExecute Me.hwnd, "open", "Dy_EncCore.exe", CommandArg, App.Path, 0

   DoEvents
   Sleep (75)
   DoEvents
   Dim Numtmp As Integer
   Numtmp = 0
   Do While Numtmp = 0
   DoEvents
   Sleep (70)
   DoEvents
      If exitproc("Dy_EncCore.exe") = False Then
         If Dir(Environ("temp") & "\finish.dyenc") <> "" Then
            Exit Do
         Else
            Timer2.Enabled = False
            Text4.Text = ""
            Label5.Caption = "00:00:00"
            Timer3.Enabled = True
            Text1.Enabled = True
            Text2.Enabled = True
            Text3.Enabled = True
            Text5.Enabled = True
            Command1.Enabled = True
            Command2.Enabled = True
            Frame1.Enabled = True
            Frame2.Enabled = True
            DoEvents
            Shell "taskkill /f /im Dy_EncCore.exe", vbHide
            On Error Resume Next
            Kill Environ("temp") & "\finish.dyenc"
            On Error Resume Next
            Kill OutputDir & "\" & GetFileName(Text3.Text) & ".Dyenc_Output"
            MsgB ("写文件失败。请检查是否在输出目录存在同名文件。")
            Exit Sub
         End If
      End If
   Loop
   
   DoEvents
   Text4.Text = Text4.Text & "[" & Format(Hour(Now), "00") & ":" & Format(Minute(Now), "00") & ":" & Format(Second(Now), "00") & "] 正在填充数据" & vbCrLf
   Dim Numtmp2 As Integer
   Open OutputDir & "\" & GetFileName(Text3.Text) & ".Dyenc_Output" For Binary As #2
   For Numtmp2 = 1 To LTotalhead
       DoEvents
       Put #2, Numtmp2, HeadData(Numtmp2 - 1)
   Next Numtmp2
   Close #2
   DoEvents
   Text4.Text = Text4.Text & "[" & Format(Hour(Now), "00") & ":" & Format(Minute(Now), "00") & ":" & Format(Second(Now), "00") & "] 完成处理，已生成文件：" & OutputDir & "\" & GetFileName(Text3.Text) & ".Dyenc_Output" & vbCrLf
            Timer2.Enabled = False
            Label5.Caption = "00:00:00"
            Timer3.Enabled = True
            Text1.Enabled = True
            Text2.Enabled = True
            Text3.Enabled = True
            Text5.Enabled = True
            Command1.Enabled = True
            Command2.Enabled = True
            Frame1.Enabled = True
            Frame2.Enabled = True
            On Error Resume Next
   If Form1.Check2.Value = 1 Then
      ShellExecute Me.hwnd, "open", "DyEnc_BulidEXE.exe", OutputDir & "\" & GetFileName(Text3.Text) & ".Dyenc_Output", App.Path, 0
      MsgB ("正在生成EXE自解密文件包，请耐心等待。")
   End If
   On Error Resume Next
   Open App.Path & "\DyEnc6.0.HISTORY" For Append As #21
        Print #21, "["; Format(Year(Now), "0000") & "-" & Format(Month(Now), "00") & "-" & Format(Day(Now), "00") & " " & Format(Hour(Now), "00") & ":" & Format(Minute(Now), "00") & ":" & Format(Second(Now), "00") & "] " & "加密：" & OutputDir & "\" & GetFileName(Text3.Text)
   Close #21
   If Check1.Value = 1 Then
      Shell "shutdown -s -t 0", vbHide
   End If
   MsgB ("完成处理，已生成文件：" & OutputDir & "\" & GetFileName(Text3.Text) & ".Dyenc_Output")
Else

'*********************************解密代码***************************************
'验证文件头（[DyEncFile]Version=5）
  DoEvents
  Text4.Text = Text4.Text & "[" & Format(Hour(Now), "00") & ":" & Format(Minute(Now), "00") & ":" & Format(Second(Now), "00") & "] 检查头数据" & vbCrLf
  Dim GuDing() As Byte, DuQuH(0 To 19) As Byte, JU As Boolean
  GuDing() = StrConv("[DyEncFile]Version=5", vbFromUnicode)
  Open Text3.Text For Binary As #3
  For I = 0 To 19
      DoEvents
      Get #3, I + 1, DuQuH(I)
  Next I
  Close #3
  JU = True
  For j = 0 To 19
      DoEvents
      If DuQuH(j) <> GuDing(j) Then
         JU = False
      End If
  Next j
  If JU = False Then
            Timer2.Enabled = False
            Text4.Text = ""
            Label5.Caption = "00:00:00"
            Timer3.Enabled = True
            Text1.Enabled = True
            Text2.Enabled = True
            Text3.Enabled = True
            Text5.Enabled = True
            Command1.Enabled = True
            Command2.Enabled = True
            Frame1.Enabled = True
            Frame2.Enabled = True
            DoEvents
            Shell "taskkill /f /im Dy_EncCore.exe", vbHide
            On Error Resume Next
            Kill Environ("temp") & "\finish.dyenc"
            On Error Resume Next
            Kill OutputDir & "\" & Left(GetFileName(Text3.Text), Len(GetFileName(Text3.Text)) - 13)
            MsgB ("该文件为DyEncryptor加密系统不可读取的文件，无法解密。")
            Command2.SetFocus
            Exit Sub
  End If

  If Dir(OutputDir & "\" & Left(GetFileName(Text3.Text), Len(GetFileName(Text3.Text)) - 13)) <> "" Then
     MsgB ("输出目录下已存在同名文件，请将其从" & OutputDir & "中移除")
            Timer2.Enabled = False
            Text4.Text = ""
            Label5.Caption = "00:00:00"
            Timer3.Enabled = True
            Text1.Enabled = True
            Text2.Enabled = True
            Text3.Enabled = True
            Text5.Enabled = True
            Command1.Enabled = True
            Command2.Enabled = True
            Frame1.Enabled = True
            Frame2.Enabled = True
            DoEvents
            Shell "taskkill /f /im Dy_EncCore.exe", vbHide
            On Error Resume Next
            Kill Environ("temp") & "\finish.dyenc"
     Exit Sub
  End If
  
  Dim DLRndH As Byte, DLEncw As Byte, DKey As Byte 'DLRndH为读取来的随机头长度，DLEncw为读取来的密码密文长度，DKey为读取来的密钥
  Dim DEncP As String, Temp3(0 To 399) As Byte, Mn As Long, OPassword As String '读取的密码密文
  Open Text3.Text For Binary As #4
    Get #4, 21, DLRndH
    Get #4, 22, DLEncw
    Get #4, 23 + CInt(DLRndH), DKey
    Mn = 0
    For kl = 24 + CInt(DLRndH) To 23 + CInt(DLRndH) + DLEncw
        DoEvents
        Get #4, kl, Temp3(Mn)
        Mn = Mn + 1
    Next kl
  Close #4
   
   Open Environ("temp") & "\TempFile2.dyenc" For Binary As #7
        For puy = 1 To DLEncw
            Put #7, puy, Temp3(puy - 1)
        Next puy
   Close #7
   Open Environ("temp") & "\TempFile2.dyenc" For Input As #8
        Line Input #8, DEncP
   Close #8
    OPassword = UnEncWord(DEncP)
    If Text1.Text <> OPassword Then
            Timer2.Enabled = False
            Text4.Text = ""
            Label5.Caption = "00:00:00"
            Timer3.Enabled = True
            Text1.Enabled = True
            Text2.Enabled = True
            Text3.Enabled = True
            Text5.Enabled = True
            Command1.Enabled = True
            Command2.Enabled = True
            Frame1.Enabled = True
            Frame2.Enabled = True
            MsgB ("您输入的密码错误！")
            Text1.Text = ""
            DoEvents
            Shell "taskkill /f /im Dy_EncCore.exe", vbHide
            On Error Resume Next
            Kill Environ("temp") & "\finish.dyenc"
            On Error Resume Next
            Kill OutputDir & "\" & Left(GetFileName(Text3.Text), Len(GetFileName(Text3.Text)) - 13)
            Command2.SetFocus
            Exit Sub
    End If
    
    Dim CommandArg2 As String
    DoEvents
    Text4.Text = Text4.Text & "[" & Format(Hour(Now), "00") & ":" & Format(Minute(Now), "00") & ":" & Format(Second(Now), "00") & "] 运行 " & App.Path & "\Dy_EncCore.exe" & vbCrLf
    CommandArg2 = Chr(34) & Text3.Text & Chr(34) & " " & Chr(34) & OutputDir & "\" & Left(GetFileName(Text3.Text), Len(GetFileName(Text3.Text)) - 13) & Chr(34) & " " & CStr(CInt(DKey)) & " " & CStr(23 + CInt(DLRndH) + DLEncw) & " " & "1"
    ShellExecute Me.hwnd, "open", "Dy_EncCore.exe", CommandArg2, App.Path, 0

   DoEvents
   Sleep (75)
   DoEvents
   Dim Numtmp3 As Integer
   Numtmp3 = 0
   Do While Numtmp3 = 0
   DoEvents
   Sleep (70)
   DoEvents
      If exitproc("Dy_EncCore.exe") = False Then
         If Dir(Environ("temp") & "\finish.dyenc") <> "" Then
            Exit Do
         Else
            Timer2.Enabled = False
            Text4.Text = ""
            Label5.Caption = "00:00:00"
            Timer3.Enabled = True
            Text1.Enabled = True
            Text2.Enabled = True
            Text3.Enabled = True
            Text5.Enabled = True
            Command1.Enabled = True
            Command2.Enabled = True
            Frame1.Enabled = True
            Frame2.Enabled = True
            DoEvents
            Shell "taskkill /f /im Dy_EncCore.exe", vbHide
            On Error Resume Next
            Kill Environ("temp") & "\finish.dyenc"
            On Error Resume Next
            Kill OutputDir & "\" & Left(GetFileName(Text3.Text), Len(GetFileName(Text3.Text)) - 13)
            MsgB ("写文件失败。请检查是否在输出目录存在同名文件。")
            Command2.SetFocus
            Exit Sub
         End If
      End If
   Loop
  Text4.Text = Text4.Text & "[" & Format(Hour(Now), "00") & ":" & Format(Minute(Now), "00") & ":" & Format(Second(Now), "00") & "] 完成处理，已生成文件：" & OutputDir & "\" & Left(GetFileName(Text3.Text), Len(GetFileName(Text3.Text)) - 13) & vbCrLf
            Timer2.Enabled = False
            Label5.Caption = "00:00:00"
            Timer3.Enabled = True
            Text1.Enabled = True
            Text2.Enabled = True
            Text3.Enabled = True
            Text5.Enabled = True
            Command1.Enabled = True
            Command2.Enabled = True
            Frame1.Enabled = True
            Frame2.Enabled = True
   On Error Resume Next
   Open App.Path & "\DyEnc6.0.HISTORY" For Append As #22
        Print #22, "["; Format(Year(Now), "0000") & "-" & Format(Month(Now), "00") & "-" & Format(Day(Now), "00") & " " & Format(Hour(Now), "00") & ":" & Format(Minute(Now), "00") & ":" & Format(Second(Now), "00") & "] " & "解密：" & OutputDir & "\" & Left(GetFileName(Text3.Text), Len(GetFileName(Text3.Text)) - 13)
   Close #22
   If Check1.Value = 1 Then
      Shell "shutdown -s -t 0", vbHide
   End If
   If OpenFN = 1 Then
      ShellExecute Me.hwnd, "open", Left(GetFileName(Text3.Text), Len(GetFileName(Text3.Text)) - 13), vbNullString, OutputDir, vbNormalFocus
   End If
   MsgB ("完成处理，已生成文件：" & OutputDir & "\" & Left(GetFileName(Text3.Text), Len(GetFileName(Text3.Text)) - 13))
End If
Command2.SetFocus
End Sub

Private Sub Command3_Click()
Form8.Show
End Sub

Private Sub Command4_Click()
Load Form4
Form4.Show
End Sub

Private Sub Command5_Click()
Form6.Show
End Sub

Private Sub Command6_2_Click()
Close_Window (15)
End Sub

Private Sub Command6_Click()
Dim sFiles() As String
    Dim filecount As Long
    Dim sDir As String
    Dim I As Long ' 循环变量
    Dim blnS As Boolean ' 是否成功打开？
    ' 文件类型过滤
    Const strSETFILTER As String = "文本文件(*.txt)|*.txt" & _
                    "|图片文件(*.bmp;*.cur;*.emf;*.gif;*.ico;*.jpg;*.jpeg;*.wmf)|*.bmp;*.cur;*.dib;*.emf;*.gif;*.ico;*.jpg;*.jpeg;*.wmf" & _
                    "|音频文件(*.wma;*.mp3;*.mp4;*.m4a;*.mid;*.midi)|*.aif;*.aiff;*.aifc;*.au;*.cda;*.wma;*.snd;*.voc;*.mp1;*.mp2;*.mp3;*.mp4;*.m4a;*.mid;*rmi;*.midi" & _
                    "|波形文件(*.wav)|*.wav" & _
                    "|所有文件(*.*)|*.*"
    Set dlg = New CCommonDialog
    With dlg
        .hwnd = Me.hwnd
        .DialogTitle = "DyEncryptor文件加密系统" ' 设置对话框标题
        ' 设置标志：是否隐藏只读、是否允许多选
        .Flags = OFN_EnableHook Or OFN_Explorer Or OFN_FileMustExist Or OFN_ShowHelp ' 使用回调！！！
        ' 设置对话框启动位置
        '.CancelError = True

                If Text3.Text = "" Then
                    MsgB "请您首先选择文件或输入文件名，再查看其属性。"
                Else
                    .ShowProperty Text3.Text
                End If
    End With
End Sub

Private Sub Command8_Click()
If Text1.PasswordChar = "*" Then
   Text1.PasswordChar = ""
   Command8.Caption = "屏蔽字符"
Else
   Text1.PasswordChar = "*"
   Command8.Caption = "显示字符"
   End If
End Sub

Private Sub Command9_Click()
If Text2.PasswordChar = "*" Then
   Text2.PasswordChar = ""
   Command9.Caption = "屏蔽字符"
Else
   Text2.PasswordChar = "*"
   Command9.Caption = "显示字符"
   End If
End Sub

Private Sub dest_Click()
Form5.Show
Form5.Text1.Text = Form1.Text3.Text
End Sub

Private Sub Form_Load()
Text4.Text = Text4.Text & "[" & Format(Hour(Now), "00") & ":" & Format(Minute(Now), "00") & ":" & Format(Second(Now), "00") & "] 加密系统加载完毕，工作目录：" & App.Path & vbCrLf
OutputDir = App.Path
rtn = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong Me.hwnd, GWL_EXSTYLE, rtn
b = SetLayeredWindowAttributes(Me.hwnd, 0, Me.HScroll1.Value, LWA_ALPHA)
Label11.Caption = "Version:" & App.Major & "." & App.Minor & "." & App.Revision

Load Form2
Load Form9

    TT.hwnd = Me.Command2.hwnd  ' 欲显示ToolTip的控件句柄（必须设置！！！）
    
    TT.ToolTipIcon = TTI_WARNING ' 图标
    TT.ToolTipTitle = "一键处理" ' 标题文字
    'TT.ToolTipText = "一键处理"
    
   TT.BackColor = vbWhite ' 背景色
   TT.ForeColor = vbMagenta ' 前景（字体）颜色

   TT.TimeToStay = 5210 ' ToolTip显示时间，停留的时间！单位：毫秒！
     'TT.TimeInterval = 3500 ' 显示ToolTip的时隔时间

   TT.TTStyle = TT_Balloon ' Tooltip 显示样式
    
    ' 创建 ToolTip 窗口（函数返回创建窗口的句柄）
    TT.CreateToolTip
    TT.TTStyle = TT_Standard
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
If Command2.Enabled = False Then
   Exit Sub
End If
Dim fileg
On Error Resume Next
For Each fileg In Data.Files
    If Err.Number > 0 Then
       MsgB (Err.Description)
       Exit Sub
    End If
    Text3.Text = fileg
Next
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Cancel = True
  If Label5.Caption <> "00:00:00" Then
     Dim a As Integer
     a = MsgBox("加密系统正在处理文件，您是否退出系统？", vbYesNo, "系统提示")
     If a = vbYes Then
        Shell "taskkill /f /im Dy_EncCore.exe", vbHide
        On Error Resume Next
        Kill Environ("temp") & "\finish.dyenc"
        On Error Resume Next
        Kill OutputDir & "\" & GetFileName(Text3.Text) & ".Dyenc_Output"
        On Error Resume Next
        Kill Environ("temp") & "\TempFile2.dyenc"
        On Error Resume Next
        Kill Environ("temp") & "\TempFile.dyenc"
        Set TT = Nothing
        Close_Window (15)
        'end
     Else
        Exit Sub
     End If
  Else
     Shell "taskkill /f /im Dy_EncCore.exe", vbHide
     On Error Resume Next
     Kill Environ("temp") & "\finish.dyenc"
        On Error Resume Next
        Kill Environ("temp") & "\TempFile2.dyenc"
        On Error Resume Next
        Kill Environ("temp") & "\TempFile.dyenc"
        Set TT = Nothing
        Close_Window (15)
     'End
  End If
  Set TT = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = True
  If Label5.Caption <> "00:00:00" Then
     Dim a As Integer
     a = MsgBox("加密系统正在处理文件，您是否退出系统？", vbYesNo, "系统提示")
     If a = vbYes Then
        Shell "taskkill /f /im Dy_EncCore.exe", vbHide
        On Error Resume Next
        Kill Environ("temp") & "\finish.dyenc"
        On Error Resume Next
        Kill OutputDir & "\" & GetFileName(Text3.Text) & ".Dyenc_Output"
        On Error Resume Next
        Kill Environ("temp") & "\TempFile2.dyenc"
        On Error Resume Next
        Kill Environ("temp") & "\TempFile.dyenc"
        Set TT = Nothing
        Close_Window (15)
        Set TT = Nothing
        'End
     Else
        Exit Sub
     End If
  Else
     Shell "taskkill /f /im Dy_EncCore.exe", vbHide
     On Error Resume Next
     Kill Environ("temp") & "\finish.dyenc"
        On Error Resume Next
        Kill Environ("temp") & "\TempFile2.dyenc"
        On Error Resume Next
        Kill Environ("temp") & "\TempFile.dyenc"
        Set TT = Nothing
        Close_Window (15)
        Set TT = Nothing
  End If
End Sub


Private Sub HScroll1_Change()
rtn = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong Me.hwnd, GWL_EXSTYLE, rtn
b = SetLayeredWindowAttributes(Me.hwnd, 0, Me.HScroll1.Value, LWA_ALPHA)

Label15.Caption = Format(Me.HScroll1.Value / 255, "0.0%")
End Sub

Private Sub Image1_DblClick()
Label7_Click
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Static ox As Integer, oy As Integer
  If Button = 1 Then
    Me.Left = Me.Left + x - ox
    Me.Top = Me.Top + y - oy
  Else
    ox = x
    oy = y
  End If
End Sub

Private Sub Label10_Click()
Form5.Show
End Sub

Private Sub Label7_Click()
Form2.Show
End Sub

Private Sub Label9_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
If Command2.Enabled = False Then
   Exit Sub
End If
On Error Resume Next
For Each fileg In Data.Files
    If Err.Number > 0 Then
       MsgB (Err.Description)
       Exit Sub
    End If
    Text3.Text = fileg
Next
End Sub

Private Sub quit_Click()
Cancel = True
  If Label5.Caption <> "00:00:00" Then
     Dim a As Integer
     a = MsgBox("加密系统正在处理文件，您是否退出系统？", vbYesNo, "系统提示")
     If a = vbYes Then
        Shell "taskkill /f /im Dy_EncCore.exe", vbHide
        On Error Resume Next
        Kill Environ("temp") & "\finish.dyenc"
        On Error Resume Next
        Kill OutputDir & "\" & GetFileName(Text3.Text) & ".Dyenc_Output"
        On Error Resume Next
        Kill Environ("temp") & "\TempFile2.dyenc"
        On Error Resume Next
        Kill Environ("temp") & "\TempFile.dyenc"
        Set TT = Nothing
        MCCloseForm Me, 4
        Set TT = Nothing
        'End
     Else
        Exit Sub
     End If
  Else
     Shell "taskkill /f /im Dy_EncCore.exe", vbHide
     On Error Resume Next
     Kill Environ("temp") & "\finish.dyenc"
        On Error Resume Next
        Kill Environ("temp") & "\TempFile2.dyenc"
        On Error Resume Next
        Kill Environ("temp") & "\TempFile.dyenc"
        Set TT = Nothing
        MCCloseForm Me, 4
        Set TT = Nothing
     'End
  End If
  Set TT = Nothing
End Sub

Private Sub setcenter_Click()
Form4.Show
End Sub

Private Sub Text3_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
If Command2.Enabled = False Then
   Exit Sub
End If
On Error Resume Next
For Each fileg In Data.Files
    If Err.Number > 0 Then
       MsgB (Err.Description)
       Exit Sub
    End If
    Text3.Text = fileg
Next
End Sub

Private Sub Text4_Change()
Text4.SelStart = Len(Text4.Text)
End Sub

Private Sub Text4_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
If Command2.Enabled = False Then
   Exit Sub
End If
On Error Resume Next
For Each fileg In Data.Files
    If Err.Number > 0 Then
       MsgB (Err.Description)
       Exit Sub
    End If
    Text3.Text = fileg
Next
End Sub

Private Sub Timer1_Timer()

End Sub

Private Sub Timer2_Timer()
Dim dif As Double
dif = DateDiff("s", starttime, Now)
Dim b As Double, C As Double, d As Double, e As Double
b = Modd(dif, 31536000)
b1 = Fix(b / 2678400)
C = Modd(b, 2678400)
c1 = Fix(C / 86400)
d = Modd(C, 86400)
d1 = Fix(d / 3600)
e = Modd(d, 3600)
e1 = Fix(e / 60)
f1 = Modd(e, 60)
Label5.Caption = Format(d1, "00") & ":" & Format(e1, "00") & ":" & Format(f1, "00")
Label8.Caption = Now()
End Sub

Private Sub Timer3_Timer()
If Option1.Value = True Then
  Label1.Caption = "设置密码："
  Text2.Enabled = True
  Text2.Visible = True
  Command9.Visible = True
  Label2.Caption = "确认密码："
  Me.Caption = "DyEncryptor" & App.Major & "." & App.Minor & " - 文件加密"
Else
  Label1.Caption = "输入密码："
  Text2.Enabled = False
  Text2.Visible = False
  Command9.Visible = False
  Label2.Caption = ""
  Me.Caption = "DyEncryptor" & App.Major & "." & App.Minor & " - 文件解密"
End If

If Option3.Value = True Then
  Text5.Enabled = False
Else
  Text5.Enabled = True
End If

If Text4.Text = "" Then
  Text4.Text = Text4.Text & "[" & Format(Hour(Now), "00") & ":" & Format(Minute(Now), "00") & ":" & Format(Second(Now), "00") & "] 加密系统加载完毕，工作目录：" & App.Path & vbCrLf
End If

Label8.Caption = Now()

End Sub

Private Sub Timer4_Timer()
If CapitalStatus() = True Then
   Label12.Caption = "温馨提示：大写键已开启!!"
Else
   Label12.Caption = ""
End If
Label8.ToolTipText = "当前时间：" & Now()
Label5.ToolTipText = "已用时间：" & Label5.Caption
End Sub

Private Sub TimerGUI_Timer()
   Me.Label13.fontname = Me.fontname
   Me.Command8.fontname = Me.fontname
   Me.Command9.fontname = Me.fontname
   Me.Command6.fontname = Me.fontname
   Me.Label14.fontname = Me.fontname
   Me.Label12.fontname = Me.fontname
End Sub

Private Sub TimerTemprt_Timer()
  Tempyui = Me.Height
Debug.Print 666666
  On Error Resume Next
  Me.Width = Me.Width - Form10.Tag
  On Error Resume Next
  Me.Height = Me.Height - TimerTemprt.Tag
  If (Me.Height = Tempyui) Then End
End Sub


Private Sub wincol_Click()
Form3.Show
End Sub

'End of Visual Basic 6.0 Source file of DyEncryptor6.0. Form1.frm.
