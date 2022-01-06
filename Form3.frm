VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DyEncryptor - ´°ÌåÑÕÉ«ÉèÖÃ"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "·ÂËÎ"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   4560
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1560
      Top             =   2520
   End
   Begin VB.Frame Frame1 
      Caption         =   "ÑÕÉ«Ô¤ÀÀ"
      Height          =   735
      Left            =   120
      TabIndex        =   11
      Top             =   1680
      Width           =   4215
      Begin VB.Label Label4 
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   3975
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "È·¶¨"
      Height          =   375
      Left            =   3240
      TabIndex        =   10
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "»Ö¸´Ä¬ÈÏ"
      Height          =   375
      Left            =   2040
      TabIndex        =   9
      Top             =   2520
      Width           =   1095
   End
   Begin VB.HScrollBar HScroll3 
      Height          =   255
      Left            =   1200
      Max             =   255
      TabIndex        =   8
      Top             =   1200
      Width           =   3255
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   255
      Left            =   1200
      Max             =   255
      TabIndex        =   7
      Top             =   720
      Width           =   3255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   1200
      Max             =   255
      TabIndex        =   6
      Top             =   240
      Width           =   3255
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "·ÂËÎ"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "0"
      Top             =   1200
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "·ÂËÎ"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "0"
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "·ÂËÎ"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "0"
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "À¶"
      BeginProperty Font 
         Name            =   "·ÂËÎ"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "ÂÌ"
      BeginProperty Font 
         Name            =   "·ÂËÎ"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "ºì"
      BeginProperty Font 
         Name            =   "·ÂËÎ"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   375
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Me.HScroll1.Value = 240
Me.HScroll2.Value = 240
Me.HScroll3.Value = 240
End Sub

Private Sub Command2_Click()
If Dir(App.Path & "\GUI_Color.config") <> "" Then
   On Error Resume Next
   Open App.Path & "\GUI_Color.config" For Output As #12
        Print #12, Text1.Text & vbCrLf & Text2.Text & vbCrLf & Text3.Text
   Close #12
End If
Form1.BackColor = RGB(Val(Text1.Text), Val(Text2.Text), Val(Text3.Text))
Form1.Frame1.BackColor = RGB(Val(Text1.Text), Val(Text2.Text), Val(Text3.Text))
Form1.Frame2.BackColor = RGB(Val(Text1.Text), Val(Text2.Text), Val(Text3.Text))
Form1.Option1.BackColor = RGB(Val(Text1.Text), Val(Text2.Text), Val(Text3.Text))
Form1.Option2.BackColor = RGB(Val(Text1.Text), Val(Text2.Text), Val(Text3.Text))
Form1.Option3.BackColor = RGB(Val(Text1.Text), Val(Text2.Text), Val(Text3.Text))
Form1.Option4.BackColor = RGB(Val(Text1.Text), Val(Text2.Text), Val(Text3.Text))
Form1.Check1.BackColor = RGB(Val(Text1.Text), Val(Text2.Text), Val(Text3.Text))
Form1.Check2.BackColor = RGB(Val(Text1.Text), Val(Text2.Text), Val(Text3.Text))
Me.Hide
End Sub

Private Sub Form_Load()
Dim rnm As String, gnm As String, bnm As String
If Dir(App.Path & "\GUI_Color.config") <> "" Then
   On Error Resume Next
   Open App.Path & "\GUI_Color.config" For Input As #11
        Line Input #11, rnm
        Line Input #11, gnm
        Line Input #11, bnm
   Close #11
   Text1.Text = rnm
   Text2.Text = gnm
   Text3.Text = bnm
   Me.HScroll1.Value = Text1.Text
   Me.HScroll2.Value = Text2.Text
   Me.HScroll3.Value = Text3.Text
End If
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
Label4.BackColor = RGB(Val(Text1.Text), Val(Text2.Text), Val(Text3.Text))
Text1.Text = Me.HScroll1.Value
Text2.Text = Me.HScroll2.Value
Text3.Text = Me.HScroll3.Value
End Sub
