VERSION 5.00
Begin VB.Form Form10 
   BorderStyle     =   0  'None
   Caption         =   "Form10"
   ClientHeight    =   3750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6015
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   5160
      Top             =   1920
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   4320
      TabIndex        =   2
      Top             =   915
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   4200
      TabIndex        =   1
      Top             =   2640
      Width           =   615
   End
   Begin VB.Line Line4 
      X1              =   240
      X2              =   4080
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line3 
      X1              =   240
      X2              =   4080
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line2 
      X1              =   240
      X2              =   240
      Y1              =   2640
      Y2              =   2880
   End
   Begin VB.Line Line1 
      X1              =   4080
      X2              =   4080
      Y1              =   2640
      Y2              =   2880
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   2640
      Width           =   3855
   End
   Begin VB.Image Image1 
      Height          =   3960
      Left            =   0
      Picture         =   "Form10.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6015
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
'���ڽ�CreateRoundRectRgn������Բ�����򸳸�����
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
'���ڴ���һ��Բ�Ǿ��Σ��þ�����X1��Y1-X2��Y2ȷ��������X3��Y3ȷ������Բ����Բ�ǻ��ȡ�
'���� ���ͼ�˵����
'X1,Y1 Long���������Ͻǵ�X��Y����
'X2,Y2 Long���������½ǵ�X��Y����
'X3 Long��Բ����Բ�Ŀ��䷶Χ��0��û��Բ�ǣ������ο�ȫԲ��
'Y3 Long��Բ����Բ�ĸߡ��䷶Χ��0��û��Բ�ǣ������θߣ�ȫԲ��
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST& = -1
' �����������б�������λ���κ�������ڵ�ǰ��
Const HWND_NOTOPMOST = -2 'ȡ�����ϲ��趨
Private Const SWP_NOSIZE& = &H1
' ���ִ��ڴ�С
Private Const SWP_NOMOVE& = &H2
' ���ִ���λ��

'��CreateRoundRectRgn����������ɾ�������Ǳ�Ҫ�ģ����򲻱�Ҫ��ռ�õ����ڴ�
'����������һ��ȫ�ֱ���,�������������
Private Sub Form_Activate() '����Activate()�¼�
Call rgnform(Me, 70, 70) '�����ӹ���
End Sub
Private Sub rgnform(ByVal frmbox As Form, ByVal fw As Long, ByVal fh As Long) '�ӹ��̣��ı����fw��fh��ֵ��ʵ��Բ��
Dim w As Long, h As Long
w = frmbox.ScaleX(frmbox.Width, vbTwips, vbPixels)
h = frmbox.ScaleY(frmbox.Height, vbTwips, vbPixels)
outrgn = CreateRoundRectRgn(0, 0, w, h, fw, fh)
Call SetWindowRgn(frmbox.hwnd, outrgn, True)
End Sub

Private Sub Form_Load()
Label3.Caption = "v" & App.Major & "." & App.Minor & "." & App.Revision
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub

Private Sub Form_Unload(Cancel As Integer) '����Unload�¼�
DeleteObject outrgn '��Բ������ʹ�õ�����ϵͳ��Դ�ͷ�
End Sub

Private Sub Timer1_Timer()
Label2.Caption = Format(Label1.Width / Tw, "0%")
End Sub
