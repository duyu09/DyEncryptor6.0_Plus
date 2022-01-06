VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H80000016&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DyEncryptor - 设置中心"
   ClientHeight    =   8010
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5295
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   ScaleHeight     =   8010
   ScaleWidth      =   5295
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame4 
      BackColor       =   &H8000000B&
      BorderStyle     =   0  'None
      Caption         =   "历史记录"
      Height          =   1215
      Left            =   0
      TabIndex        =   13
      Top             =   5520
      Width           =   5415
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   15
         Text            =   "30"
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "历史记录"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   0
         Width           =   1095
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
         TabIndex        =   17
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "条历史记录"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   16
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "最多显示"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H8000000B&
      BorderStyle     =   0  'None
      Caption         =   "图标"
      Height          =   1935
      Left            =   0
      TabIndex        =   9
      Top             =   3480
      Width           =   5415
      Begin VB.CommandButton Command4 
         Caption         =   "..."
         Height          =   375
         Left            =   4560
         TabIndex        =   12
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   720
         Width           =   3135
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "图标"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "图标预览："
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   975
      End
      Begin VB.Image Image1 
         Height          =   975
         Left            =   240
         Stretch         =   -1  'True
         Top             =   720
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000B&
      BorderStyle     =   0  'None
      Caption         =   "字体"
      Height          =   1095
      Left            =   0
      TabIndex        =   6
      Top             =   2280
      Width           =   5415
      Begin VB.CommandButton Command5 
         Caption         =   "搜索所有字体"
         Height          =   375
         Left            =   3360
         TabIndex        =   24
         Top             =   600
         Width           =   1695
      End
      Begin VB.ComboBox Combo1 
         Height          =   375
         Left            =   3360
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "字体"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   0
         Width           =   615
      End
      Begin VB.Line Line6 
         X1              =   240
         X2              =   3120
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line Line7 
         X1              =   240
         X2              =   3120
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Line Line8 
         X1              =   3120
         X2              =   3120
         Y1              =   360
         Y2              =   840
      End
      Begin VB.Line Line5 
         X1              =   240
         X2              =   240
         Y1              =   360
         Y2              =   840
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "字体预览"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   480
         Width           =   2655
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000B&
      BorderStyle     =   0  'None
      Caption         =   "窗体颜色"
      Height          =   1095
      Left            =   0
      TabIndex        =   3
      Top             =   1080
      Width           =   5415
      Begin VB.CommandButton Command1 
         Caption         =   "设置窗体颜色"
         Height          =   495
         Left            =   3360
         TabIndex        =   4
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "窗体颜色"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   0
         Width           =   735
      End
      Begin VB.Line Line2 
         X1              =   240
         X2              =   3120
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line Line3 
         X1              =   240
         X2              =   3120
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Line Line1 
         X1              =   3120
         X2              =   3120
         Y1              =   360
         Y2              =   840
      End
      Begin VB.Line Line4 
         X1              =   240
         X2              =   240
         Y1              =   360
         Y2              =   840
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "窗体颜色"
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   480
         Width           =   2655
      End
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H8000000B&
      Caption         =   "解密完成后打开文件"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   7080
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
      Caption         =   "取消"
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   7560
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "确定"
      Height          =   375
      Left            =   4200
      TabIndex        =   0
      Top             =   7560
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4800
      Top             =   120
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H8000000B&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   -120
      TabIndex        =   22
      Top             =   6840
      Width           =   5655
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "后续任务"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.Image Image2 
      Height          =   960
      Left            =   0
      Picture         =   "Form4.frx":C84A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5340
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const LF_FACESIZE = 32
Private Const CF_PRINTERFONTS = &H2
Private Const CF_SCREENFONTS = &H1
Private Const CF_BOTH = (CF_SCREENFONTS Or CF_PRINTERFONTS)
Private Const CF_EFFECTS = &H100&
Private Const CF_FORCEFONTEXIST = &H10000
Private Const CF_INITTOLOGFONTSTRUCT = &H40&
Private Const CF_LIMITSIZE = &H2000&
Private Const REGULAR_FONTTYPE = &H400
'charset Constants
Private Const ANSI_CHARSET = 0
Private Const ARABIC_CHARSET = 178
Private Const BALTIC_CHARSET = 186
Private Const CHINESEBIG5_CHARSET = 136
Private Const DEFAULT_CHARSET = 1
Private Const EASTEUROPE_CHARSET = 238
Private Const GB2312_CHARSET = 134
Private Const GREEK_CHARSET = 161
Private Const HANGEUL_CHARSET = 129
Private Const HEBREW_CHARSET = 177
Private Const JOHAB_CHARSET = 130
Private Const MAC_CHARSET = 77
Private Const OEM_CHARSET = 255
Private Const RUSSIAN_CHARSET = 204
Private Const SHIFTJIS_CHARSET = 128
Private Const SYMBOL_CHARSET = 2
Private Const THAI_CHARSET = 222
Private Const TURKISH_CHARSET = 162
Private Type LOGFONT
lfHeight As Long
lfWidth As Long
lfEscapement As Long
lfOrientation As Long
lfWeight As Long
lfItalic As Byte
lfUnderline As Byte
lfStrikeOut As Byte
lfCharSet As Byte
lfOutPrecision As Byte
lfClipPrecision As Byte
lfQuality As Byte
lfPitchAndFamily As Byte
lfFaceName As String * 31
End Type
Private Declare Function ChooseFont Lib "comdlg32.dll" Alias "ChooseFontA" (ByRef pChoosefont As ChooseFont) As Long
Private Type ChooseFont
lStructSize As Long
hwndOwner As Long ' caller's window handle
hdc As Long ' printer DC/IC or NULL
lpLogFont As Long ' ptr. to a LOGFONT struct
iPointSize As Long ' 10 * size in points of selected font
Flags As Long ' enum. type flags
rgbColors As Long ' returned text color
lCustData As Long ' data passed to hook fn.
lpfnHook As Long ' ptr. to hook function
lpTemplateName As String ' custom template name
hInstance As Long ' instance handle of.EXE that
' contains cust. dlg. template
lpszStyle As String ' return the style field here
' must be LF_FACESIZE or bigger
nFontType As Integer ' same value reported to the EnumFonts
' call back with the extra FONTTYPE_
' bits added
MISSING_ALIGNMENT As Integer
nSizeMin As Long ' minimum pt size allowed &
nSizeMax As Long ' max pt size allowed if
' CF_LIMITSIZE is used
End Type

Private Sub Combo1_Click()
Label2.Font.Name = Combo1.Text
End Sub

Private Sub Command1_Click()
Form3.Show
End Sub

Private Sub Command2_Click()
Dim fnc2 As String
If Dir(Text1.Text) = "" Then
   MsgB ("您设置的图标不存在")
   Exit Sub
End If
fnc2 = Combo1.Text
SetFoName (fnc2)
Form1.Icon = LoadPicture(Text1.Text)
Form2.Icon = LoadPicture(Text1.Text)
Form3.Icon = LoadPicture(Text1.Text)
Form4.Icon = LoadPicture(Text1.Text)
Form6.Icon = LoadPicture(Text1.Text)
OpenFN = Check1.Value
NumOfHi = Val(Text2.Text)
On Error Resume Next
Open App.Path & "\DyEncGUI5.0.OtherSettings.config" For Output As #7
     If Err.Number > 0 Then
        MsgB ("写配置文件失败。")
        Exit Sub
     End If
     Print #7, fnc2
     Print #7, Text1.Text
     Print #7, Text2.Text
     Print #7, OpenFN
Close #7
DU_IconP = Text1.Text
Unload Me
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
Dim OFN As OPENFILENAME
Dim rtn As String, xnk As String
If Dir(selP, vbDirectory) <> "" Then
   xnk = selP
Else
   xnk = App.Path & "\DyEncGUI_IconLib"
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
OFN.lpstrTitle = "请选择ico格式图标 - DyEncryptor"
OFN.Flags = 6148
rtn = GetOpenFileName(OFN)
If rtn >= 1 Then
   Text1.Text = OFN.lpstrFile
End If
If Right(Text1.Text, 4) <> ".ico" Then
   MsgB ("请选择ico格式图标")
   Text1.Text = ""
   Exit Sub
End If
On Error Resume Next
Image1.Picture = LoadPicture(Text1.Text)
If Err.Number > 0 Then
   MsgB ("加载图标失败。")
   Image1.Picture = LoadPicture(App.Path & "\DyEncIcon.ico")
   Text1.Text = App.Path & "\DyEncIcon.ico"
   Exit Sub
End If
End Sub

Private Sub Command5_Click()
Dim CF As ChooseFont, lfont As LOGFONT
Dim fontname As String, ret As Long
CF.Flags = CF_BOTH Or CF_EFFECTS Or CF_FORCEFONTEXIST Or CF_INITTOLOGFONTSTRUCT Or CF_LIMITSIZE
CF.lpLogFont = VarPtr(lfont)
CF.lStructSize = LenB(CF)
'cf.lStructSize = Len(cf) ' size of structure
CF.hwndOwner = Me.hwnd ' window Form1 is opening this dialog box
'cf.hDC = Printer.hDC ' device context of default printer (using VB's mechanism)
CF.rgbColors = RGB(0, 0, 0) ' black
CF.nFontType = REGULAR_FONTTYPE ' regular font type i.e. not bold or anything
CF.nSizeMin = 10 ' minimum point size
CF.nSizeMax = 72 ' maximum point size
ret = ChooseFont(CF) 'brings up the font dialog
If ret <> 0 Then ' success
fontname = StrConv(lfont.lfFaceName, vbUnicode, &H804) 'Retrieve chinese font name in english version os
fontname = Left$(fontname, InStr(1, fontname, vbNullChar) - 1)
'Assign the font properties to text1

'.Charset = lfont.lfCharSet 'assign charset to font
Me.Combo1.Text = fontname
'.Size = cf.iPointSize / 10 'assign point size
End If
End Sub

Private Sub Form_Load()
Text1.Text = App.Path & "\DyEncIcon.ico"
Image1.Picture = Form1.Icon
Dim a As Integer
For a = 0 To Screen.FontCount - 1
    DoEvents
    Me.Combo1.AddItem Screen.Fonts(a)
Next a
Combo1.Text = Form1.fontname
Label2.fontname = Combo1.Text
Check1.Value = OpenFN
Text2.Text = NumOfHi
Text1.Text = DU_IconP
On Error Resume Next
Image1.Picture = LoadPicture(Text1.Text)

Dim fnc2 As String
fnc2 = Form1.fontname
Form4.fontname = fnc2
Form4.Check1.fontname = fnc2
Form4.Combo1.fontname = fnc2
Form4.Command1.fontname = fnc2
Form4.Command2.fontname = fnc2
Form4.Command3.fontname = fnc2
Form4.Command4.fontname = fnc2
Form4.Frame1.fontname = fnc2
Form4.Frame2.fontname = fnc2
Form4.Frame3.fontname = fnc2
Form4.Frame4.fontname = fnc2
Form4.Label1.fontname = fnc2
Form4.Label10.fontname = fnc2
Form4.Label2.fontname = fnc2
Form4.Label3.fontname = fnc2
Form4.Label4.fontname = fnc2
Form4.Label5.fontname = fnc2
Form4.Label6.fontname = fnc2
Form4.Label7.fontname = fnc2
Form4.Label8.fontname = fnc2
Form4.Label9.fontname = fnc2
Form4.Text1.fontname = fnc2
Form4.Text2.fontname = fnc2
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Static ox As Integer, oy As Integer
  If Button = 1 Then
    Me.Left = Me.Left + x - ox
    Me.Top = Me.Top + y - oy
  Else
    ox = x
    oy = y
  End If
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
       End If
       Print #6, vbNullString
       Close #6
MsgB ("历史记录已清除。")
End Sub

Private Sub Text2_Change()
If Val(Text2.Text) > 32766 Or Val(Text2.Text) < 0 Then
   MsgBox "可保留0~32766条历史记录。", 48
   Text2.Text = "30"
   Exit Sub
End If
End Sub

Private Sub Timer1_Timer()
Label1.BackColor = Form1.BackColor
End Sub
