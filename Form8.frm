VERSION 5.00
Begin VB.Form Form8 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DyEncryptor - 导出运行记录"
   ClientHeight    =   4155
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4725
   Icon            =   "Form8.frx":0000
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   4725
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command2 
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
      Height          =   375
      Left            =   4200
      TabIndex        =   5
      Top             =   1320
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
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
      Left            =   3120
      TabIndex        =   4
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "完成后打开输出目录"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   2520
      Value           =   1  'Checked
      Width           =   4335
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "完成后打开文件"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1920
      Width           =   4335
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1320
      Width           =   3975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "输出目录设置："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   0
      Picture         =   "Form8.frx":048A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Sub Command1_Click()
On Error Resume Next
Open Text1.Text & "\DyEncryptor.log" For Append As #1
 If Err.Number > 0 Then
   MsgB ("生成日志失败：" & Err.Description)
   Exit Sub
 End If
Print #1, Now & " 生成日志"
Print #1, "----------------------------"
Print #1, Form1.Text4.Text
Close #1
If Check1.Value = 1 Then
   ShellExecute Me.hwnd, "open", Text1.Text & "\DyEncryptor.log", vbNullString, Text1.Text, vbNormalFocus
End If
If Check2.Value = 1 Then
   ShellExecute Me.hwnd, "open", Text1.Text, vbNullString, Text1.Text, vbNormalFocus
End If
MsgB ("已生成日志文件DyEncryptor.log于" & Text1.Text)
End Sub

Private Sub Command2_Click()
Set sh = CreateObject("Shell.Application")
Set fd = sh.BrowseForFolder(Me.hwnd, "DyEncryptor" & App.Major & "." & App.Minor & " - 请选择输出目录", 0)
'Me.hWnd是“选择文件夹”对话框的父窗口句柄，不关闭对话框不能返回父窗口。改为0没这个效果。
'第三个参数0改为512不显示“新建文件夹”按钮
If TypeName(fd) = "Folder3" Then Text1.Text = fd.Self.Path
End Sub

Private Sub Form_Load()
Dim opa As String
Text1.Text = OutputDir
Me.Text1.fontname = Form1.fontname
Me.Label1.fontname = Form1.fontname
Me.Check1.fontname = Form1.fontname
Me.Check2.fontname = Form1.fontname
Me.Command1.fontname = Form1.fontname
End Sub
