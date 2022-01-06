VERSION 5.00
Begin VB.Form Form7 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "DyEncryptor - 加密系统温馨提示"
   ClientHeight    =   4455
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6465
   Icon            =   "Form7.frx":0000
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   6465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   4680
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "未曾忘记那位少年"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   1200
      Width           =   6255
   End
   Begin VB.Image Image4 
      Height          =   495
      Left            =   4800
      Top             =   240
      Width           =   495
   End
   Begin VB.Image Image3 
      Height          =   495
      Left            =   5880
      Top             =   240
      Width           =   375
   End
   Begin VB.Image Image5 
      Height          =   1215
      Left            =   3000
      Picture         =   "Form7.frx":C84A
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   3495
   End
   Begin VB.Image Image2 
      Height          =   855
      Left            =   3960
      Picture         =   "Form7.frx":1D02E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   120
      Picture         =   "Form7.frx":1DC85
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6375
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'圆角窗体
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Dim lw&, lh&
Private Sub Form_Load()
lw = Me.Width \ Screen.TwipsPerPixelX
lh = Me.Height \ Screen.TwipsPerPixelY
SetWindowRgn hwnd, CreateRoundRectRgn(0, 0, lw, lh, 90, 90), True
Text1.fontname = Form1.fontname
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Rtb = 1
Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Rtb = 1
Unload Me
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
Private Sub label1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Static ox As Integer, oy As Integer
  If Button = 1 Then
    Me.Left = Me.Left + x - ox
    Me.Top = Me.Top + y - oy
  Else
    ox = x
    oy = y
  End If
End Sub

Private Sub Image3_Click()
Rtb = 1
Unload Me
End Sub

Private Sub Image5_Click()
Rtb = 1
Unload Me
End Sub

Private Sub Text1_GotFocus()
Me.Check1.SetFocus
End Sub
