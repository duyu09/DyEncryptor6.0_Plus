VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DyEncryptor - �ļ����� (���������)"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7935
   BeginProperty Font 
      Name            =   "΢���ź�"
      Size            =   14.25
      Charset         =   134
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form5.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   3180
   ScaleWidth      =   7935
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   7320
      TabIndex        =   5
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   840
      Width           =   5175
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "---�����ļ����˴�---"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      OLEDropMode     =   1  'Manual
      TabIndex        =   4
      Top             =   2040
      Width           =   2655
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000013&
      BorderStyle     =   5  'Dash-Dot-Dot
      FillColor       =   &H8000000D&
      Height          =   1575
      Left            =   2040
      Shape           =   4  'Rounded Rectangle
      Top             =   1440
      Width           =   5775
   End
   Begin VB.Image Image1 
      Height          =   1455
      Left            =   240
      Picture         =   "Form5.frx":0582
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "�������ļ�����"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "�����˴���ʼ����"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   5160
      TabIndex        =   1
      Top             =   2520
      Width           =   2535
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "���棡�����ٵ��ļ����޷�ͨ���κη�ʽ�ָ��������������"
      ForeColor       =   &H000000FF&
      Height          =   390
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7695
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub Command1_Click()
Dim OFN As OPENFILENAME
Dim rtn As String
OFN.lStructSize = Len(OFN)
OFN.hwndOwner = Me.hwnd
OFN.hInstance = App.hInstance
OFN.lpstrFilter = "�����ļ�(*.*)"
OFN.lpstrFile = Space(254)
OFN.nMaxFile = 255
OFN.lpstrFileTitle = Space(254)
OFN.nMaxFileTitle = 255
OFN.lpstrInitialDir = xnk
OFN.lpstrTitle = "�����ѡ�񱻴�����ļ� - DyEncryptor"
OFN.Flags = 6148
rtn = GetOpenFileName(OFN)
If rtn >= 1 Then
   Text1.Text = OFN.lpstrFile
End If
End Sub

Private Sub Form_Load()
Text1.Text = Form1.Text3.Text
Me.BackColor = RGB(255, 230, 230)
Form5.Label1.fontname = Form1.fontname
Form5.Label2.fontname = Form1.fontname
Form5.Label3.fontname = Form1.fontname
Form5.Text1.fontname = Form1.fontname
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
For Each fileg In Data.Files
    If Err.Number > 0 Then
       MsgB (Err.Description)
       Exit Sub
    End If
    Text1.Text = fileg
Next
End Sub

Private Sub Label2_Click()
If Text1.Text = "" Then
   MsgB ("�������ļ�����·����")
   Exit Sub
End If
If Dir(App.Path & "\DyEnc_FileDestroyModule.exe") = "" Then
   MsgB ("�ļ����ٺ��������ʧ���޷��ļ�����ִ������")
   Exit Sub
End If
Dim tmpa As Integer
tmpa = MsgBox("���ٴ�ȷ���Ƿ�ʼ����", vbYesNo, "����ϸ�ٴκ˶� - DyEncryptor����ϵͳ��ʾ")
If tmpa = vbNo Then
   Exit Sub
End If
ShellExecute Me.hwnd, "open", "DyEnc_FileDestroyModule.exe", Chr(34) & Text1.Text & Chr(34) & " " & Len(Text1.Text), App.Path, 0
   DoEvents
   DoEvents
   Sleep (100)
   DoEvents
Do While exitproc("DyEnc_FileDestroyModule.exe")
   DoEvents
   Sleep (88)
   DoEvents
Loop
   DoEvents
   Sleep (88)
   DoEvents
On Error Resume Next
Kill Text1.Text
MsgB ("�ļ�������ϡ�")
Unload Me
End Sub

Private Sub Label4_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
For Each fileg In Data.Files
    If Err.Number > 0 Then
       MsgB (Err.Description)
       Exit Sub
    End If
    Text3.Text = fileg
Next
End Sub
