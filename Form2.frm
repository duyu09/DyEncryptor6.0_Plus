VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DyEncryptor6.0 - ��Ȩ����"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   8040
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   8040
   StartUpPosition =   2  '��Ļ����
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4080
      Top             =   3600
   End
   Begin VB.CommandButton Command2 
      Caption         =   "�鿴ͼƬ"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   4
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ȷ��"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6960
      TabIndex        =   2
      Top             =   4080
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2415
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1320
      Width           =   7815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3840
      TabIndex        =   3
      ToolTipText     =   "��Ȩ��������³��ҵ��ѧ ������̿���1�� ����(202103180009) ��������Ȩ����"
      Top             =   3960
      Width           =   3135
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   0
      Picture         =   "Form2.frx":C84A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright (c) 2016-2022 Duyu"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "��Ȩ��������³��ҵ��ѧ ������̿���1�� ����(202103180009) ��������Ȩ����"
      Top             =   3960
      Width           =   4335
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Me.Hide
End Sub

Private Sub Command2_Click()
Form9.Show
End Sub

Private Sub Form_Load()
Dim a As String
a = "���Ȩ������" & vbCrLf & "���Ʒ���ƣ�DyEncryptor6.0" & vbCrLf & "�������Ʒ����ͬһ��װĿ¼�µ������ļ������ܺ������Dy_EncCore.exe�Լ����������ļ�������������Ȩ����" & vbCrLf & "����³��ҵ��ѧ ������̣����������21-01��  ���� "
Text1.Text = a
b = Text1.Text + vbCrLf & "�����ϵͳ������־��" & vbCrLf & "  2016-04-05.  version:1.0." & vbCrLf & "  2018-10-13.  version:1.0_Plus." & vbCrLf & "  2019-06-02.  version:2.0."
b = b & vbCrLf & "  2020-02-26.  version:3.0."
b = b & vbCrLf & "  2020-03-26.  version:4.0."
b = b & vbCrLf & "  2021-08-29.  version:5.0."
b = b & vbCrLf & "  2021-09-17.  version:5.0_Plus."
b = b & vbCrLf & "  2021-10-28.  version:6.0."
b = b & vbCrLf & "  2022-01-01.  version:6.0_Plus."
Text1.Text = b
Label2.Caption = "Version: " & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub Text1_GotFocus()
Command1.SetFocus
End Sub

Private Sub Timer1_Timer()
   Me.Text1.fontname = Form1.fontname
   Me.Command1.fontname = Form1.fontname
   Me.Command2.fontname = Form1.fontname
   Me.Label1.fontname = Form1.fontname
   Me.Label2.fontname = Form1.fontname
End Sub
