VERSION 5.00
Begin VB.Form Form6 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DyEncryptor - 加密系统运行历史记录"
   ClientHeight    =   6810
   ClientLeft      =   90
   ClientTop       =   420
   ClientWidth     =   8025
   Icon            =   "Form6.frx":0000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   8025
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command2 
      Caption         =   "打开历史记录文件"
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
      Left            =   4920
      TabIndex        =   2
      Top             =   6360
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
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
      Left            =   6840
      TabIndex        =   1
      Top             =   6360
      Width           =   975
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2580
      Left            =   0
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "清除历史记录"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   6360
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   0
      Picture         =   "Form6.frx":0442
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8055
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd _
As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Const WM_SYSCOMMAND = &H112&
Const SC_MONITORPOWER = &HF170&
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Shell "notepad.exe " & Chr(34) & App.Path & "\DyEnc6.0.HISTORY" & Chr(34), vbNormalFocus
End Sub


Private Sub Form_Load()
SendMessage List1.hwnd, &H194, 3000, ByVal 0
Dim ght(1 To 32766) As String, ser As Integer
Open App.Path & "\DyEnc6.0.HISTORY" For Input As #24
     Do While Not EOF(24)
        ser = ser + 1
        Line Input #24, ght(ser)
     Loop
Close #24
If NumOfHi > ser Then
   For a = 1 To ser
       List1.AddItem ght(a)
   Next a
Else
   For b = ser - NumOfHi + 1 To ser
       List1.AddItem ght(b)
   Next b
End If
End Sub

Private Sub Form_Resize()
List1.Width = Me.Width * 0.985
List1.Height = Me.Height * 0.88 - List1.Top
List1.fontname = Form1.fontname
End Sub


Private Sub Label6_Click()
xtu = MsgBox("您确定要清除历史记录吗？清除后将无法恢复。", vbYesNo)
If xtu = vbNo Then
   Exit Sub
End If
On Error Resume Next
Open App.Path & "\DyEnc6.0.HISTORY" For Output As #6
       If Err.Number > 0 Then
          MsgBox "历史记录清除失败。", 48
          Exit Sub
       End If
       Print #6, vbNullString
       Close #6
MsgB ("历史记录已清除。")
Me.List1.Clear
End Sub

Private Sub List1_DblClick()
MsgB (List1.Text)
End Sub
