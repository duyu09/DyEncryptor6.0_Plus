Attribute VB_Name = "MCDHook"
Option Explicit
' --- API 函数 申明
' 释放程序内存
Private Declare Function SetProcessWorkingSetSize Lib "kernel32" (ByVal hProcess As Long, ByVal dwMinimumWorkingSetSize As Long, ByVal dwMaximumWorkingSetSize As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long

' API call to alter the class data for this window
Private Declare Function SetWindowLong& Lib "user32" Alias "SetWindowLongA" (ByVal hWnd&, _
                                                              ByVal nIndex&, ByVal dwNewLong&)
Private Declare Function CallWindowProc& Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc&, _
                                                 ByVal hWnd&, ByVal msg&, ByVal wParam&, ByVal lParam&)

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetDlgItem Lib "user32" (ByVal hDlg As Long, ByVal nIDDlgItem As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
' 取得控件相对屏幕左上角的坐标值！（单位：像素？！）
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWndParent As Long, ByVal hWndChildAfter As Long, ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function UpdateWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, lpString As Any) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal e As Long, ByVal O As Long, ByVal W As Long, ByVal I As Long, ByVal u As Long, ByVal s As Long, ByVal C As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal f As String) As Long

' VB 取得图片大小
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long

' =========================================================================================
' ==== 声音文件的播放（可去掉）============================================================
' =========================================================================================
'API 申明 使用PlaySound函数播放声音
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, _
    ByVal hModule As String, ByVal dwFlags As Long) As Long
'API 申明 使用sndPlaySound函数播放声音，它是 PlaySound 函数的子集？！
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, _
    ByVal uFlags As Long) As Long
Private Declare Function sndStopSound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszNull As Long, ByVal uFlags As Long) As Long
'关闭声音
'sndPlaySound Null, SND_ASYNC
'PlaySound 0,0,0
' 高级媒体播放函数
Private Declare Function mciExecute Lib "winmm.dll" (ByVal lpstrCommand As String) As Long
' mciSendString 是用来播放多媒体文件的API指令，可以播放MPEG,AVI,WAV,MP3,等等
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
' Multimedia Command Strings: http://msdn.microsoft.com/en-us/library/ms712587.aspx
' MCI Command Strings:http://msdn.microsoft.com/en-us/library/ms710815(VS.85).aspx

' --- for PlaySound \ sndPlaySound
Private Const SND_ASYNC = &H1 ' play asynchronously 在播放的同时继续执行以后的语句
Private Const SND_FILENAME = &H20000 ' name is a file name
Private Const SND_LOOP = &H8 ' loop the sound until next sndPlaySound 一直重复播放声音，直到该函数开始播放第二个声音为止
Private Const SND_MEMORY = &H4         '  lpszSoundName points to a memory file 播放内存中的声音, 譬如资源文件中的声音
Private Const SND_NODEFAULT = &H2         '  silence not default, if sound not found
Private Const SND_NOSTOP = &H10        '  don't stop any currently playing sound
Private Const SND_NOWAIT = &H2000      '  don't wait if the driver is busy
Private Const SND_PURGE = &H40               '  purge non-static events for task
Private Const SND_RESERVED = &HFF000000  '  In particular these flags are reserved
Private Const SND_RESOURCE = &H40004     '  name is a resource name or atom
Private Const SND_SYNC = &H0         '  play synchronously (default) 播放完声音之后再执行后面的语句
Private Const SND_TYPE_MASK = &H170007
Private Const SND_VALID = &H1F        '  valid flags          / ;Internal /
Private Const SND_VALIDFLAGS = &H17201F    '  Set of valid flag bits.  Anything outside

Private Enum PlayStatus ' 声音播放状态！
    IsPlaying = 0
    IsPaused = 1
    IsStopped = 2
End Enum
' =========================================================================================
' ==== 声音文件的播放（可去掉）============================================================
' =========================================================================================


' =========================================================================================
' ==== 字体对话框 （单独）=================================================================
' =========================================================================================
Rem --------------------------------------------------------
Rem FONT STUFF
Rem --------------------------------------------------------
Public Const LF_FACESIZE = 32
Public Type LOGFONT
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
        lfFaceName(LF_FACESIZE) As Byte
        'lfFaceName As String * LF_FACESIZE
End Type
'Private lpLF As LOGFONT

Public Const LOGPIXELSY = 90    '  Logical pixels/inch in Y

Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long

Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetWindowDC Lib "user32.dll" (ByVal hWnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Public Declare Function lstrcpyANY Lib "kernel32" Alias "lstrcpyA" (p1 As Any, p2 As Any) As Long
Public Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Public Declare Function GetTextFace Lib "gdi32" Alias "GetTextFaceA" (ByVal hdc As Long, ByVal nCount As Long, ByVal lpFacename As String) As Long
Public Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
   
Private Declare Function GetTextColor Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
'Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
' hDC 很重要！
' 说明 设置当前文本颜色。这种颜色也称为“前景色” 返回值 Long，文本色的前一个RGB颜色设定。CLR_INVALID表示失败。

Rem --------------------------------------------------------

Rem --------------------------------------------------------
Rem ChooseFont structure and function declarations
Rem --------------------------------------------------------
Public Type ChooseFontType
    lStructSize As Long
    hwndOwner As Long           '  caller's window handle
    hdc As Long                 '  printer DC/IC or NULL
    lpLogFont As Long           '  ptr. to a LOGFONT struct - changed from old "lpLogFont As LOGFONT"
    iPointSize As Long          '  10 * size in points of selected font
    Flags As Long               '  enum. type flags
    rgbColors As Long           '  returned text color
    lCustData As Long           '  data passed to hook fn.
    lpfnHook As Long            '  ptr. to hook function
    lpTemplateName As String    '  custom template name
    hInstance As Long           '  instance handle of.EXE that contains cust. dlg. template
    lpszStyle As String         '  return the style field here must be LF_FACESIZE or bigger
    nFontType As Integer        '  same value reported to the EnumFonts call back with the extra FONTTYPE bits added
    MISSING_ALIGNMENT As Integer
    nSizeMin As Long            '  minimum pt size allowed &
    nSizeMax As Long            '  max pt size allowed if CF_LIMITSIZE is used
End Type

Public Declare Function ChooseFont Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As ChooseFontType) As Long
Private Declare Function SendDlgItemMessage Lib "user32.dll" Alias "SendDlgItemMessageA" (ByVal hDlg As Long, ByVal nIDDlgItem As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Const CB_GETCURSEL As Long = &H147
Private Const CB_GETITEMDATA As Long = &H150
Private Const CB_ERR As Long = (-1)
Private Const CB_RESETCONTENT As Long = &H14B

Public Enum CF_Flags
    CF_SCREENFONTS = &H1&
    CF_PRINTERFONTS = &H2&
    CF_BOTH = (CF_SCREENFONTS Or CF_PRINTERFONTS)
    CF_SHOWHELP = &H4&
    CF_ENABLEHOOK = &H8&
    CF_ENABLETEMPLATE = &H10&
    CF_ENABLETEMPLATEHANDLE = &H20&
    CF_INITTOLOGFONTSTRUCT = &H40&
    CF_USESTYLE = &H80&
    CF_EFFECTS = &H100&
    CF_APPLY = &H200&
    CF_ANSIONLY = &H400&
    CF_NOVECTORFONTS = &H800&
    CF_NOOEMFONTS = CF_NOVECTORFONTS
    CF_NOSIMULATIONS = &H1000&
    CF_LIMITSIZE = &H2000&
    CF_FIXEDPITCHONLY = &H4000&
    CF_WYSIWYG = &H8000&           'Must also have CF_SCREENFONTS and CF_PRINTERFONTS
    CF_FORCEFONTEXIST = &H1000&
    CF_SCALABLEONLY = &H2000&
    CF_TTONLY = &H4000&
    CF_NOFACESEL = &H8000&
    CF_NOSTYLESEL = &H100000
    CF_NOSIZESEL = &H200000
End Enum

Public Const SIMULATED_FONTTYPE = &H8000
Public Const PRINTER_FONTTYPE = &H4000
Public Const SCREEN_FONTTYPE = &H2000
Public Const BOLD_FONTTYPE = &H100
Public Const ITALIC_FONTTYPE = &H200
Public Const REGULAR_FONTTYPE = &H400

'public Const WM_CHOOSEFONT_GETLOGFONT = (&H400 + 1) 'WM_USER + 1

Public Const LBSELCHSTRING = "commdlg_LBSelChangedNotify"
Public Const SHAREVISTRING = "commdlg_ShareViolation"
Public Const FILEOKSTRING = "commdlg_FileNameOK"
Public Const COLOROKSTRING = "commdlg_ColorOK"
Public Const SETRGBSTRING = "commdlg_SetRGBColor"
Public Const FINDMSGSTRING = "commdlg_FindReplace"
Public Const HELPMSGSTRING = "commdlg_help"

Public Const CD_LBSELNOITEMS = -1
Public Const CD_LBSELCHANGE = 0
Public Const CD_LBSELSUB = 1
Public Const CD_LBSELADD = 2

Rem ------------------------------------------------------------
Rem per maggior praticit? ho enumerato tutti i controlli della
Rem finestra Carattere
Rem ------------------------------------------------------------
Public Enum enumFONT_CTL ' 字体对话框上的控件 ID
    stc_FontName = 1088 ' 字体(&F): 标签
    edt_FontName = 1001 ' 字体名称 文本框？？
    cbo_FontName = &H470  ' 字体名称 下拉框？？66672
    
    stc_BoldItalic = 1089 ' 字形(&Y): 标签
    edt_BoldItalic = 1001 ' 字形 文本框？？
    cbo_BoldItalic = &H471  ' 字形 下拉框？？66673
    
    stc_Size = 1090 ' 大小(&S): 标签
    edt_Size = 1001 ' 大小 文本框？？
    cbo_Size = &H472  ' 大小 下拉框？？66674
    
    btn_Ok = 1 ' 确定(&O) 按钮
    btn_Cancel = 2 ' 取消(&C) 按钮
    btn_Apply = 1026 ' 应用(&A) 按钮
    btn_Help = 1038 ' 帮助(&H) 按钮
    
    btn_Effects = 1072 ' 效果 组合框
    btn_Strikethrough = &H410 ' 删除线(&K) 按钮
    btn_Underline = &H411 ' 下划线(&U) 按钮
    stc_Color = &H443 ' 颜色(&C): 标签
    cbo_Color = &H473 ' 颜色 下拉框？？66675
    
    btn_Sample = 1073 ' 示例组合框
    stc_SampleText = &H444 ' 示例标签：微软中文软件
    
    stc_Charset = 1094 ' 字符集(&R): 标签
    cbo_Charset = &H474 ' 字符集下拉框
    stc_Description = 1093 ' 字体描述标签：该字体用于显示。打印时将使用最接近的匹配字体。

    ' Note: 'Axis' is a invisible groupbox with some controls
    btn_Axis = 1074     ' groupbox
    hsb_1 = 1168        ' horizontal scrollbar
    hsb_2 = 1169
    hsb_3 = 1170
    hsb_4 = 1171
    hsb_5 = 1172
    hsb_6 = 1173
    stc_1 = 1098        ' static
    stc_2 = 1099
    stc_3 = 1100
    stc_4 = 1101
    stc_5 = 1102
    stc_6 = 1103
    stc_7 = 1105
    stc_8 = 1106
    stc_9 = 1107
    stc_10 = 1108
    stc_11 = 1109
    stc_12 = 1110
    stc_13 = 1112
    stc_14 = 1113
    stc_15 = 1114
    stc_16 = 1115
    stc_17 = 1116
    stc_18 = 1118
    edt_1 = 1152        ' edit
    edt_2 = 1153
    edt_3 = 1154
    edt_4 = 1155
    edt_5 = 1156
    edt_6 = 1157
End Enum
' =========================================================================================
' ==== 字体对话框（单独）==================================================================
' =========================================================================================



' =========================================================================================
' ==== 颜色对话框 （单独）=================================================================
' =========================================================================================
Rem --------------------------------------------------------
Rem ChooseColor structure and function declarations
Rem --------------------------------------------------------
Public Type CHOOSECOLOR_TYPE
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As Long
    Flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Public Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As CHOOSECOLOR_TYPE) As Long
Public Enum CC_Flags
    CC_RGBINIT = &H1
    CC_FULLOPEN = &H2
    CC_PREVENTFULLOPEN = &H4
    CC_SHOWHELP = &H8
    CC_ENABLEHOOK = &H10
    CC_ENABLETEMPLATE = &H20
    CC_ENABLETEMPLATEHANDLE = &H40
End Enum

Rem --------------------------------------------------------
Rem Public MEMORY Stuff
Rem --------------------------------------------------------
Public Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Function GlobalFree Lib "kernel32" (ByVal hmem As Long) As Long
Public Declare Function GlobalLock Lib "kernel32" (ByVal hmem As Long) As Long
Public Declare Function GlobalUnlock Lib "kernel32" (ByVal hmem As Long) As Long
'Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, Source As Any, ByVal Lenght As Long)

Public Const GMEM_MOVEABLE = &H2
Public Const GMEM_ZEROINIT = &H40
Public Const GHND = (GMEM_MOVEABLE Or GMEM_ZEROINIT)
Rem --------------------------------------------------------

' =========================================================================================
' ==== 颜色对话框（单独）=================================================================
' =========================================================================================


' --- 常数 申明
' for Windows 消息 常数
Private Const WM_USER = &H400
Private Const WM_INITDIALOG = &H110
Private Const WM_NOTIFY = &H4E
Private Const WM_DESTROY = &H2
Private Const WM_COMMAND As Long = &H111
Private Const WM_GETDLGCODE = &H87
Private Const WM_SETREDRAW = &HB
Private Const WM_SHOWWINDOW = &H18
Private Const WM_WINDOWPOSCHANGING = &H46
Private Const WM_WINDOWPOSCHANGED = &H47
Private Const WM_NCCALCSIZE = &H83
Private Const WM_CHILDACTIVATE = &H22
Private Const WM_NCDESTROY = &H82
Private Const WM_MOUSEMOVE As Long = &H200
Private Const WM_LBUTTONDOWN As Long = &H201
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_CHOOSEFONT_GETLOGFONT As Long = (WM_USER + 1)
Private Const WM_CHOOSEFONT_SETFLAGS As Long = (WM_USER + 102)
Private Const WM_CHOOSEFONT_SETLOGFONT As Long = (WM_USER + 101)
Private Const WM_CTLCOLOREDIT = &H133
Private Const WM_CLOSE As Long = &H10
Private Const WM_GETFONT As Long = &H31
Private Const WM_SETFONT As Long = &H30
Private Const WM_SETTEXT = &HC
Private Const WM_PAINT As Long = &HF&
Private Const SW_NORMAL = 1
Private Const SW_HIDE = 0
Private Const SW_SHOW = 5
' for SetWindowLong&
Private Const GWL_WNDPROC As Long = (-4&)

' for 对话框上的消息
Private Const CDM_First = (WM_USER + 100)                   '/---
Private Const CDM_GetSpec = (CDM_First + &H0)               '取得文件名
Private Const CDM_GetFilePath = (CDM_First + &H1)           '取得文件名与目录
Private Const CDM_GetFolderPath = (CDM_First + &H2)         '取得路径
Private Const CDM_GetFolderIDList = (CDM_First + &H3)       '
Private Const CDM_SetControlText = (CDM_First + &H4)        '设置控件文本
Private Const CDM_HideControl = (CDM_First + &H5)           '隐藏控件
Private Const CDM_SetDefext = (CDM_First + &H6)             '
Private Const CDM_Last = (WM_USER + 200)                    '\---

Private Const CDN_First = (-601)                            '/---
Private Const CDN_InitDone = (CDN_First - &H0)              '初始化完成
Private Const CDN_SelChange = (CDN_First - &H1)             '选择文件改变
Private Const CDN_FolderChange = (CDN_First - &H2)          '目录改变
Private Const CDN_ShareViolation = (CDN_First - &H3)        '
Private Const CDN_Help = (CDN_First - &H4)                  '点了帮助
Private Const CDN_FileOK = (CDN_First - &H5)                '点了确定
Private Const CDN_TypeChange = (CDN_First - &H6)            '过滤类型改变
Private Const CDN_IncludeItem = (CDN_First - &H7)           '
Private Const CDN_Last = (-699)                             '\---
  
' for 对话框上控件的 ID
Private Const ID_FolderLabel   As Long = &H443              '“查找范围(&I)”标签
Private Const ID_FolderCombo   As Long = &H471              '目录下拉框
Private Const ID_ToolBar       As Long = &H440              '工具栏（特别注意：无法通过 CDMoveOriginControl 函数移动！）
Private Const ID_ToolBarWin2K  As Long = &H4A0              '快捷目录区（版本>=Win2K）

' 列表框（列出文件的最大区域）
Private Const ID_List0         As Long = &H460              ' 使用这个有效！？！
Private Const ID_List1         As Long = &H461
Private Const ID_List2         As Long = &H462

Private Const ID_OK            As Long = 1                  '“确定(&O)”按键
Private Const ID_Cancel        As Long = 2                  '“取消(&C)”按键
Private Const ID_Help          As Long = &H40E              '“帮助(&H)”按键
Private Const ID_ReadOnly      As Long = &H410              '“只读”多选框

Private Const ID_FileTypeLabel As Long = &H441              '“文件类型(&T)”标签
Private Const ID_FileNameLable As Long = &H442              '“文件名(&N)”标签
'“文件类型(&T)”下拉框
Private Const ID_FileTypeCombo0 As Long = &H470             ' 使用这个有效！？！
Private Const ID_FileTypeCombo1 As Long = &H471
Private Const ID_FileTypeCombo2 As Long = &H472
Private Const ID_FileTypeComboC As Long = &H47C             '“文件名(&N)”文本框
Private Const ID_FileNameText  As Long = &H480              '“文件名(&N)”文本框（新外观时不是它！）

' for SendMessage 取得复选框是否选中？
Private Const BM_GETCHECK = &HF0

' for CreateWindowEx 创建预览文本框
Private Const WS_EX_STATICEDGE = &H20000
Private Const WS_EX_TRANSPARENT = &H20&
Private Const WS_EX_TOPMOST = &H8&
Private Const WS_BORDER = &H800000
Private Const WS_CHILD = &H40000000
Private Const WS_HSCROLL = &H100000
Private Const WS_VISIBLE = &H10000000
Private Const WS_VSCROLL = &H200000
Private Const ES_AUTOHSCROLL = &H80&
Private Const ES_AUTOVSCROLL = &H40&
Private Const ES_LEFT = &H0&
Private Const ES_MULTILINE = &H4&       ' 文本允许多行
Private Const ES_READONLY = &H800&      ' 将编辑框设置成只读的
Private Const ES_CENTER = &H1&          ' 文本显示居中
Private Const ES_WANTRETURN = &H1000&   ' 使多行编辑器接收回车键输入并换行。如果不指定该风格，按回车键会选择缺省的命令按钮，这往往会导致对话框的关闭。

' --- for CreateFont 字体信息常数
Private Const CLIP_LH_ANGLES            As Long = 16 ' 字符旋转所需要的
Private Const PROOF_QUALITY             As Long = 2
Private Const TRUETYPE_FONTTYPE         As Long = &H4
Private Const ANTIALIASED_QUALITY       As Long = 4
Private Const DEFAULT_CHARSET           As Long = 1
Private Const FF_DONTCARE = 0    '  Don't care or don't know.
Private Const DEFAULT_PITCH = 0
Private Const OUT_DEFAULT_PRECIS = 0

' --- 枚举 申明
Public Enum PreviewPosition ' 预览图片框位置
    ppNone = -1 ' 设为此值时，不显示！
    ppTop = 0
    ppLeft = 1
    ppRight = 2
    ppBottom = 3
End Enum
Public Enum DialogStyle ' 对话框风格，打开？保存？字体？颜色？
    ssOpen = 0
    ssSave = 1
    ssFont = 2
    ssColor = 3
End Enum
Private Enum FileType
    ffText = 0      ' 文本 预览（默认值，任何文件可以以文本方式打开？！）
    ffPicture = 1   ' 图片 预览
    ffWave = 2      ' Wave 波形文件 预览，画出声音波形！！
    ffAudio = 3     ' 一般音频文件，添加播放、暂停、停止按钮，进行预览。API播放声音！！
End Enum

' --- 结构体 申明
' for CopyMemory 取得对话框哪些控件改变？
Private Type NMHDR
    hwndFrom   As Long
    idFrom   As Long
    code   As Long
End Type
' 坐标？
Private Type POINTAPI
  X As Long
  y As Long
End Type
Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type
' for GetObject
Private Type BITMAP ' 取得BITMAP结构体
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type
Private Type PicInfo ' 图片宽、高
    picWidth As Long
    picHeight As Long
End Type

' --- 私有变量 申明
Private procOld As Long ' 保存原 窗体属性的变量，其实是默认的 窗体函数 的地址
Private hWndTextView As Long ' 动态创建的预览文本框 句柄
Private hWndButtonPlay(0 To 2) As Long ' 3 个播放按钮 句柄
Private strSelFile As String ' 选中的文件路径

' ==== 字体对话框 （单独）=================================================================
Private hWndFontPreview As Long ' 字体预览文本框
' ==== 字体对话框 （单独）=================================================================

' --- 公共变量 申明...为 CCommonDialog 服务！！
Public IsReadOnlyChecked As Boolean ' 指示是否选定只读复选框
Public WhichStyle As DialogStyle ' 对话框风格，打开？保存？字体？颜色？

' 特别特别注意：图片框设计时必须有图片，否则第二次弹出对话框时图片框消失？！！！且窗体上要放两个空的图片框（不作任何事，当摆设！！！）
Public m_picLogoPicture As PictureBox ' 程序标志图片框图片
Public m_picPreviewPicture As PictureBox ' 预览图片框图片
Public m_ppLogoPosition As PreviewPosition ' 程序标志图片框位置
Public m_ppPreviewPosition As PreviewPosition ' 预览图片框位置
Public m_dlgStartUpPosition As StartUpPositionConstants ' 对话框启动位置？
Public m_blnHideControls(0 To 8) As Boolean ' 是否隐藏对话框上的控件？（可去掉）
Public m_strControlsCaption(0 To 6) As String ' 对话框上的控件的文字？（可去掉）

' ################################################################################################
' 回调函数，用来截取消息，让动态创建的控件可以响应消息！（注意：是截取对话框消息！）
' ################################################################################################
Private Function WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, _
                                              ByVal wParam As Long, ByVal lParam As Long) As Long
    ' 确定接收到的是什么消息
    Select Case iMsg
        Case WM_COMMAND ' 单击
            Dim I As Integer
            For I = 0 To 2
                If lParam = hWndButtonPlay(I) Then Call B3Button_Click(I)
            Next I
'        Case WM_LBUTTONDOWN ' 鼠标左键按下
'            Debug.Print "WM_LBUTTONDOWN " & lParam
    End Select
  
    ' 如果不是我们需要的消息，则传递给原来的窗体函数处理
    WindowProc = CallWindowProc(procOld, hWnd, iMsg, wParam, lParam)

End Function
' 设置开始和结束的两个过程！！！
Private Sub CDHook(ByVal hWnd As Long)
    ' 整个procOld变量用来存储窗口的原始参数，以便恢复
    ' 调用了 SetWindowLong 函数，它使用了 GWL_WNDPROC 索引来创建窗口类的子类，通过这样设置
    ' 操作系统发给窗体的消息将由回调函数 (WindowProc) 来截取， AddressOf是关键字取得函数地址
    procOld = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf WindowProc)
             ' AddressOf是一元运算符，它在过程地址传送到 API 过程之前，先得到该过程的地址
End Sub
Private Sub CDUnHook(ByVal hWnd As Long)
    ' 此句关键，把窗口（不是窗体，而是具有句柄的任一控件）的属性复原
    Call SetWindowLong(hWnd, GWL_WNDPROC, procOld)
End Sub
' ################################################################################################
' 回调函数，用来截取消息，让动态创建的控件可以响应消息！
' ################################################################################################


' 回调函数，对话框显示时要使用！！！
Public Function CDCallBackFun(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
On Error GoTo CDCallBack_Error
'    Debug.Print "&H" + Hex$(hWnd); ":",
    Dim retV As Long ' 函数返回值？！
    
    ' 取得父窗体句柄？（仅是打开、保存对话框句柄？！，字体时，hWnd 是对话框句柄！！）
    Dim hWndParent As Long: hWndParent = GetParent(hWnd)

    ' 判断消息，检测是否为需处理的消息
    Select Case uMsg
        Case WM_INITDIALOG ' 对话框初始化时，
            Debug.Print "WM_INITDIALOG", "&H" + Hex(wParam), "&H" + Hex(lParam)
            ' 私有变量初始化！
            procOld = 0: hWndTextView = 0: strSelFile = ""
            hWndButtonPlay(0) = 0: hWndButtonPlay(1) = 0: hWndButtonPlay(2) = 0
            ' 显示对话框之前。自定义字体对话框外观。
            CDHook hWndParent ' 回调函数，用来截取消息，让动态创建的控件可以响应消息！
            If WhichStyle = ssFont Then CustomizeFontDialog hWnd ' 初始化字体对话框
            If WhichStyle = ssColor Then setDlgStartUpPosition hWnd, hWndParent ' 初始化颜色对话框，只需改启动位置！
            ' 判断有没有设置两个图片框？？！！
            ' 修正了没有设置预览或程序标志图片框时，对话框位置无法调整的问题；
            If m_picLogoPicture Is Nothing Then m_ppLogoPosition = ppNone
            If m_picPreviewPicture Is Nothing Then m_ppPreviewPosition = ppNone
            
        Case WM_NOTIFY ' 对话框变化时，仅对打开/保存对话框！！！
            retV = CDNotify(hWndParent, lParam)
        Case WM_COMMAND ' 仅单击 字体、颜色对话框 上的控件？！
            'Debug.Print LOWORD(wParam); HIWORD(wParam)
            Dim L As Long: L = LOWORD(wParam)
            If WhichStyle = ssFont Then
                If L = enumFONT_CTL.btn_Apply _
                    Or L = enumFONT_CTL.cbo_FontName Or L = enumFONT_CTL.cbo_BoldItalic _
                    Or L = enumFONT_CTL.cbo_Size Or L = enumFONT_CTL.btn_Strikethrough _
                    Or L = enumFONT_CTL.btn_Underline Or L = enumFONT_CTL.cbo_Color _
                    Or L = enumFONT_CTL.cbo_Charset Then ' lParam 控件句柄？ wParam 参数=控件 ID ！！！
                    ' 设置字体对话框预览。（有些单击不管用要双击，前3个cbo！）
                    mSetFontPreview hWnd
                ElseIf L = enumFONT_CTL.btn_Help Then ' 字体对话框帮助！
                    MsgBox "字体对话框帮助！", vbInformation
                'Else ' 其他单击，发送消息单击 应用 按钮！还是不行！？
                '    SendMessage GetDlgItem(hWnd, enumFONT_CTL.btn_Apply), WM_LBUTTONDOWN, 0&, 0&
                End If
            ElseIf WhichStyle = ssColor Then
                If L = enumFONT_CTL.btn_Help Then ' 颜色对话框帮助！几个按钮通用一个ID值？！
                    MsgBox "颜色对话框帮助！", vbInformation
                End If
            End If
        Case WM_DESTROY ' 对话框销毁时，
            Debug.Print "WM_DESTROY", "&H" + Hex(wParam), "&H" + Hex(lParam)
            ' 取得 是否选定只读复选框
            Dim hWndButton As Long: hWndButton = GetDlgItem(hWndParent, ID_ReadOnly)
            IsReadOnlyChecked = SendMessage(hWndButton, BM_GETCHECK, ByVal 0&, ByVal 0&)
            
            ' 设置图片框到原来的父窗口，（桌面句柄=0，恢复到其原来的父，再次弹出对话框时图片消失了！！！）
            ' 如果窗体上有两个图片框，再次弹出时就能正常显示？！！！）
            If Not m_picLogoPicture Is Nothing Then ' 若不加判断，会引发错误！
                ShowWindow m_picLogoPicture.hWnd, SW_HIDE
                Call SetParent(m_picLogoPicture.hWnd, Val(m_picLogoPicture.Tag))
            End If
            If Not m_picPreviewPicture Is Nothing Then
                ShowWindow m_picPreviewPicture.hWnd, SW_HIDE
                Call SetParent(m_picPreviewPicture.hWnd, Val(m_picPreviewPicture.Tag))
            End If
            
            ' 停止声音 停止声音，因为前面可能在播放！！！！' 注意：这里不直接调用 B3Button_Click 2 ！
            PlayAudio strSelFile, 2
            sndPlaySound vbNullString, 0 ' 停止 Wave 文件播放。
            
            ' 销毁创建的控件
            If hWndTextView Then DestroyWindow hWndTextView
            If hWndButtonPlay(0) Then DestroyWindow hWndButtonPlay(0) ': hWndButtonPlay(0) = 0
            If hWndButtonPlay(1) Then DestroyWindow hWndButtonPlay(1) ': hWndButtonPlay(1) = 0
            If hWndButtonPlay(2) Then DestroyWindow hWndButtonPlay(2) ': hWndButtonPlay(2) = 0
            If hWndFontPreview Then DestroyWindow hWndFontPreview  ' 字体预览文本框
            ' 回调函数，用来截取消息，让动态创建的控件可以响应消息！
            CDUnHook hWndParent
            ' 释放物理内存！！！不知为什么，调出对话框后，程序占用的物理内存大增！！！
            SetProcessWorkingSetSize GetCurrentProcess(), -1&, -1&
'        Case Else
'            Debug.Print "Else ", "&H" + Hex(wParam), "&H" + Hex(lParam)
    End Select
    CDCallBackFun = retV ' 函数返回值？！
    
    On Error GoTo 0
    Exit Function

CDCallBack_Error:
    Debug.Print "CDCallBackFun Error " & Err.Number & " (" & Err.Description & ")"
    Resume Next
End Function
          
' 对话框变化时，进行调整。仅对打开/保存对话框！！！
Private Function CDNotify(ByVal hWndParent As Long, ByVal lParam As Long) As Long
    Dim hToolBar As Long    ' 对话框上工具栏句柄
    Dim rcTB As RECT        ' 工具栏矩形
    Dim pt As POINTAPI, W As Long, H As Long
    Dim rcDlg As RECT       ' 对话框矩形
    Dim picLeft As Long, picTop As Long ' 图片框位置坐标，两个图片框相互影响，一个移动时要判断另一个的位置！
    ' == 中间最大的列表框矩形，图片框 Left Top 位置的基准点。。。
    Dim hWndControl As Long, rcList0 As RECT, ptL As POINTAPI
    hWndControl = GetDlgItem(hWndParent, ID_List0) ' 根据ID取得控件句柄
    GetWindowRect hWndControl, rcList0 ' 取得控件矩形
    ptL.X = rcList0.Left: ptL.y = rcList0.Top
    ScreenToClient hWndParent, ptL ' ptL经过转化后才能得到想要的结果！
    
    Dim hdr     As NMHDR
    Call CopyMemory(hdr, ByVal lParam, LenB(hdr))
    Select Case hdr.code
        Case CDN_InitDone ' 初始化完成，对话框将要显示时，
            Debug.Print "InitDone"
                        
            ' ===== 判断程序标志图片框位置，以调整对话框外观（尺寸及其上的控件位置）！
            Dim OffSetX As Long, OffSetY As Long, stpX As Single, stpY As Single  ' 对话框大小、控件偏移量（像素！）
            stpX = Screen.TwipsPerPixelX: stpY = Screen.TwipsPerPixelY ' Twips 转化为 Pixels 要除以他们！
            If m_ppLogoPosition = ppNone Then GoTo NoLogo ' 判断有没有设置两个图片框？？！！
            OffSetX = m_picLogoPicture.Width \ stpX: OffSetY = m_picLogoPicture.Height \ stpY
            Dim ClientRect As RECT ' ppBottom 时！取得对话框矩形，与其他都不同！不知为什么只有这样才行！！！？？？
            Select Case m_ppLogoPosition
                Case ppNone ' 无程序标志图片，不操作！
                    OffSetX = 0: OffSetY = 0
                    picLeft = 0: picTop = 0
                Case ppLeft ' 程序标志图片 在左端，要移动对话框上原来的控件！
                    ' 对话框上所有原始控件右移
                    CDMoveOriginControl hWndParent, ID_OK, OffSetX
                    CDMoveOriginControl hWndParent, ID_Cancel, OffSetX
                    CDMoveOriginControl hWndParent, ID_Help, OffSetX
                    CDMoveOriginControl hWndParent, ID_ReadOnly, OffSetX
                                        
                    CDMoveOriginControl hWndParent, ID_FolderLabel, OffSetX
                    CDMoveOriginControl hWndParent, ID_FolderCombo, OffSetX
                    CDMoveOriginControl hWndParent, ID_ToolBarWin2K, OffSetX
                    CDMoveOriginControl hWndParent, ID_List0, OffSetX
                    
                    CDMoveOriginControl hWndParent, ID_FileNameLable, OffSetX
                    CDMoveOriginControl hWndParent, ID_FileTypeCombo0, OffSetX
                    CDMoveOriginControl hWndParent, ID_FileTypeLabel, OffSetX
                    CDMoveOriginControl hWndParent, ID_FileTypeComboC, OffSetX  ' 新外观！移动对话框上文件名文本框，
                    CDMoveOriginControl hWndParent, ID_FileNameText, OffSetX
                    
                    ' 移动工具栏！工具栏（特别注意：无法通过 CDMoveOriginControl 函数移动！）
                    hToolBar = CDGetToolBarHandle(hWndParent)
                    GetWindowRect hToolBar, rcTB
                    pt.X = rcTB.Left
                    pt.y = rcTB.Top
                    ScreenToClient hWndParent, pt
                    MoveWindow hToolBar, pt.X + OffSetX, pt.y, rcTB.Right - rcTB.Left, rcTB.Bottom - rcTB.Top, True
                    
                    ' 改变对话框大小，！
                    GetWindowRect hWndParent, rcDlg ' 取得对话框矩形，再移动（实际只改变宽度！）
                    MoveWindow hWndParent, rcDlg.Left, rcDlg.Top, rcDlg.Right - rcDlg.Left + OffSetX, rcDlg.Bottom - rcDlg.Top, True
                    ' 设置程序标志图片
                    ' 设置新的，并保存图片框原来的父窗口句柄？！
                    m_picLogoPicture.Tag = SetParent(m_picLogoPicture.hWnd, hWndParent)
                    ' 移动图片框，Top 位置固定，高度固定！
                    If m_ppPreviewPosition = ppLeft Then
                        picLeft = 2: picTop = 0
                    'ElseIf m_ppPreviewPosition = ppRight Then' 不需要判断！
                    ElseIf m_ppPreviewPosition = ppTop Then
                        picLeft = 2: picTop = m_picPreviewPicture.Height \ stpY
                    ElseIf m_ppPreviewPosition = ppBottom Then
                        picLeft = 2: picTop = m_picPreviewPicture.Height \ stpY
                    Else
                        picLeft = 2: picTop = 0
                    End If
                    MoveWindow m_picLogoPicture.hWnd, picLeft, 2, _
                        m_picLogoPicture.Width \ stpX, rcDlg.Bottom - rcDlg.Top + picTop - 29, True
                    ' 加载图片
                    'm_picLogoPicture.PaintPicture m_picLogoPicture.Picture, 0, 0, m_picLogoPicture.ScaleWidth * 100, m_picLogoPicture.ScaleHeight
                    ShowWindow m_picLogoPicture.hWnd, SW_SHOW ' 显示图片框

                Case ppRight
                    ' 改变对话框大小，！
                    GetWindowRect hWndParent, rcDlg ' 取得对话框矩形，再移动（实际只改变宽度！）
                    MoveWindow hWndParent, rcDlg.Left, rcDlg.Top, rcDlg.Right - rcDlg.Left + OffSetX, rcDlg.Bottom - rcDlg.Top, True
                    ' 设置程序标志图片
                    ' 设置新的，并保存图片框原来的父窗口句柄？！
                    m_picLogoPicture.Tag = SetParent(m_picLogoPicture.hWnd, hWndParent)
                    ' 移动图片框，，Top 位置固定，高度固定！
                    If m_ppPreviewPosition = ppLeft Then
                        picLeft = rcDlg.Right - rcDlg.Left + m_picPreviewPicture.Width \ stpX - 8: picTop = 0
                    ElseIf m_ppPreviewPosition = ppRight Then
                        picLeft = m_picPreviewPicture.Width \ stpX + rcDlg.Right - rcDlg.Left - 5: picTop = 0
                    ElseIf m_ppPreviewPosition = ppTop Then
                        picLeft = rcDlg.Right - rcDlg.Left - 8: picTop = m_picPreviewPicture.Height \ stpY
                    ElseIf m_ppPreviewPosition = ppBottom Then
                        picLeft = rcDlg.Right - rcDlg.Left - 8: picTop = m_picPreviewPicture.Height \ stpY
                    Else
                        picLeft = rcDlg.Right - rcDlg.Left - 8: picTop = 0
                    End If
                    MoveWindow m_picLogoPicture.hWnd, picLeft, 2, _
                        m_picLogoPicture.Width \ stpX, rcDlg.Bottom - rcDlg.Top + picTop - 29, True
                    ' 加载图片
                    'm_picLogoPicture.PaintPicture m_picLogoPicture.Picture, 0, 0, m_picLogoPicture.ScaleWidth, m_picLogoPicture.ScaleHeight
                    ShowWindow m_picLogoPicture.hWnd, SW_SHOW ' 显示图片框

                Case ppTop ' 程序标志图片 在顶端，要移动对话框上原来的控件！
                    ' 对话框上所有原始控件下移
                    CDMoveOriginControl hWndParent, ID_OK, , OffSetY
                    CDMoveOriginControl hWndParent, ID_Cancel, , OffSetY
                    CDMoveOriginControl hWndParent, ID_Help, , OffSetY
                    CDMoveOriginControl hWndParent, ID_ReadOnly, , OffSetY
                    
                    CDMoveOriginControl hWndParent, ID_FolderLabel, , OffSetY
                    CDMoveOriginControl hWndParent, ID_FolderCombo, , OffSetY
                    CDMoveOriginControl hWndParent, ID_ToolBarWin2K, , OffSetY
                    CDMoveOriginControl hWndParent, ID_List0, , OffSetY
                    
                    CDMoveOriginControl hWndParent, ID_FileNameLable, , OffSetY
                    CDMoveOriginControl hWndParent, ID_FileTypeCombo0, , OffSetY
                    CDMoveOriginControl hWndParent, ID_FileTypeLabel, , OffSetY
                    CDMoveOriginControl hWndParent, ID_FileTypeComboC, , OffSetY ' 新外观！移动对话框上文件名文本框，
                    CDMoveOriginControl hWndParent, ID_FileNameText, , OffSetY
                    
                    ' 移动工具栏！工具栏（特别注意：无法通过 CDMoveOriginControl 函数移动！）
                    hToolBar = CDGetToolBarHandle(hWndParent)
                    GetWindowRect hToolBar, rcTB
                    pt.X = rcTB.Left
                    pt.y = rcTB.Top
                    ScreenToClient hWndParent, pt
                    MoveWindow hToolBar, pt.X, pt.y + OffSetY, rcTB.Right - rcTB.Left, rcTB.Bottom - rcTB.Top, True
                    
                    ' 改变对话框大小，！
                    GetWindowRect hWndParent, rcDlg ' 取得对话框矩形，再移动（实际只改变高度！）
                    MoveWindow hWndParent, rcDlg.Left, rcDlg.Top, rcDlg.Right - rcDlg.Left, rcDlg.Bottom - rcDlg.Top + OffSetY, True
                    ' 设置程序标志图片
                    ' 设置新的，并保存图片框原来的父窗口句柄？！
                    m_picLogoPicture.Tag = SetParent(m_picLogoPicture.hWnd, hWndParent) ' GetParent(m_picLogoPicture.hwnd)
                    ' 移动图片框，Left 位置固定，宽度固定。picLeft + (rcDlg.Right - rcDlg.Left - m_picLogoPicture.Width \ stpX) \ 2 - 3
                    If m_ppPreviewPosition = ppLeft Then
                        picLeft = m_picPreviewPicture.Width \ stpX: picTop = 0
                    ElseIf m_ppPreviewPosition = ppRight Then ' 不需要判断！
                        picLeft = m_picPreviewPicture.Width \ stpX
                    'ElseIf m_ppPreviewPosition = ppTop Then
                    'ElseIf m_ppPreviewPosition = ppBottom Then
                    End If
                    MoveWindow m_picLogoPicture.hWnd, 5, picTop + 2, _
                        rcDlg.Right - rcDlg.Left + picLeft - 15, m_picLogoPicture.Height \ stpY, True
                    ' 加载图片
                    'm_picLogoPicture.PaintPicture m_picLogoPicture.Picture, 0, 0, m_picLogoPicture.ScaleWidth, m_picLogoPicture.ScaleHeight
                    ShowWindow m_picLogoPicture.hWnd, SW_SHOW ' 显示图片框

                Case ppBottom
                    ' 改变对话框大小，！
                    GetWindowRect hWndParent, rcDlg ' 取得对话框矩形，再移动（实际只改变高度！）
                    MoveWindow hWndParent, rcDlg.Left, rcDlg.Top, rcDlg.Right - rcDlg.Left, rcDlg.Bottom - rcDlg.Top + OffSetY, True
                    ' 设置程序标志图片
                    Call GetClientRect(hWndParent, ClientRect) ' 用 rcDlg.Bottom 不行！！！
                    ' 设置新的，并保存图片框原来的父窗口句柄？！
                    m_picLogoPicture.Tag = SetParent(m_picLogoPicture.hWnd, hWndParent)
                    ' 移动图片框，Left 位置固定，宽度固定。
                    If m_ppPreviewPosition = ppLeft Then
                        picLeft = m_picPreviewPicture.Width \ stpX: picTop = 0
                    ElseIf m_ppPreviewPosition = ppRight Then
                        picLeft = m_picPreviewPicture.Width \ stpX
                    ElseIf m_ppPreviewPosition = ppTop Then
                        picTop = m_picPreviewPicture.Height \ stpY: picLeft = 0
                    ElseIf m_ppPreviewPosition = ppBottom Then ' 这时，要移动标志到预览下面！
                        picTop = m_picPreviewPicture.Height \ stpY: picLeft = 0
                    End If
                    MoveWindow m_picLogoPicture.hWnd, 5, picTop + ClientRect.Bottom - OffSetY, _
                        rcDlg.Right - rcDlg.Left + picLeft - 15, m_picLogoPicture.Height \ stpY, True
                    ' 加载图片
                    'm_picLogoPicture.PaintPicture m_picLogoPicture.Picture, 0, 0, m_picLogoPicture.ScaleWidth, m_picLogoPicture.ScaleHeight
                    ShowWindow m_picLogoPicture.hWnd, SW_SHOW ' 显示图片框

            End Select
NoLogo:
' **********************************************************************************************************
            ' ===== 判断预览图片框位置，特别注意：要判断程序标志图片框位置？！！！！方法：？？？！！！
            ' 预览图片框位置固定一个值！！！大小在左右、上下分两种情况：分别固定高度、宽度！！！
            If m_ppPreviewPosition = ppNone Then GoTo NoPreview ' 判断有没有设置两个图片框？？！！
            OffSetX = m_picPreviewPicture.Width \ stpX: OffSetY = m_picPreviewPicture.Height \ stpY
            Select Case m_ppPreviewPosition
                Case ppNone
                    OffSetX = 0: OffSetY = 0
                    picLeft = 0: picTop = 0
                Case ppLeft ' 预览图片框 在左端，要移动对话框上原来的控件！
                    ' 重新取得 LIST控件矩形，可能在上面被移动了！
                    GetWindowRect hWndControl, rcList0
                    ptL.X = rcList0.Left: ptL.y = rcList0.Top
                    ScreenToClient hWndParent, ptL ' ptL经过转化后才能得到想要的结果！
                    
                    ' 对话框上所有原始控件右移
                    CDMoveOriginControl hWndParent, ID_OK, OffSetX
                    CDMoveOriginControl hWndParent, ID_Cancel, OffSetX
                    CDMoveOriginControl hWndParent, ID_Help, OffSetX
                    CDMoveOriginControl hWndParent, ID_ReadOnly, OffSetX
                                        
                    CDMoveOriginControl hWndParent, ID_FolderLabel, OffSetX
                    CDMoveOriginControl hWndParent, ID_FolderCombo, OffSetX
                    CDMoveOriginControl hWndParent, ID_ToolBarWin2K, OffSetX
                    CDMoveOriginControl hWndParent, ID_List0, OffSetX
                    
                    CDMoveOriginControl hWndParent, ID_FileNameLable, OffSetX
                    CDMoveOriginControl hWndParent, ID_FileTypeCombo0, OffSetX
                    CDMoveOriginControl hWndParent, ID_FileTypeLabel, OffSetX
                    CDMoveOriginControl hWndParent, ID_FileTypeComboC, OffSetX  ' 新外观！移动对话框上文件名文本框，
                    CDMoveOriginControl hWndParent, ID_FileNameText, OffSetX
                    
                    ' 移动工具栏！工具栏（特别注意：无法通过 CDMoveOriginControl 函数移动！）
                    hToolBar = CDGetToolBarHandle(hWndParent)
                    GetWindowRect hToolBar, rcTB
                    pt.X = rcTB.Left
                    pt.y = rcTB.Top
                    ScreenToClient hWndParent, pt
                    MoveWindow hToolBar, pt.X + OffSetX, pt.y, rcTB.Right - rcTB.Left, rcTB.Bottom - rcTB.Top, True
                    
                    ' 改变对话框大小，！
                    GetWindowRect hWndParent, rcDlg ' 取得对话框矩形，再移动（实际只改变宽度！）
                    MoveWindow hWndParent, rcDlg.Left, rcDlg.Top, rcDlg.Right - rcDlg.Left + OffSetX, rcDlg.Bottom - rcDlg.Top, True
                    ' 设置 预览图片框
                    ' 设置新的，并保存图片框原来的父窗口句柄？！
                    m_picPreviewPicture.Tag = SetParent(m_picPreviewPicture.hWnd, hWndParent)
                    ' 移动图片框，Top 位置固定，高度固定！
                    picLeft = 5: picTop = ptL.y: W = 5
                    If m_ppLogoPosition = ppLeft Then
                        picLeft = m_picLogoPicture.Width \ stpX + 5
                    'ElseIf m_ppLogoPosition = ppRight Then' 不需要判断！
                    'ElseIf m_ppLogoPosition = ppTop Then
                    'ElseIf m_ppLogoPosition = ppBottom Then
                    End If
                    MoveWindow m_picPreviewPicture.hWnd, picLeft, picTop, _
                        m_picPreviewPicture.Width \ stpX - W, rcList0.Bottom - rcList0.Top, True
                    ' 加载图片
                    'm_picPreviewPicture.PaintPicture m_picPreviewPicture.Picture, 0, 0, m_picPreviewPicture.ScaleWidth, m_picPreviewPicture.ScaleHeight
                    myPaintPicture m_picPreviewPicture, False
                    ShowWindow m_picPreviewPicture.hWnd, SW_SHOW ' 显示图片框

                Case ppRight
                    ' 重新取得 LIST控件矩形，可能在上面被移动了！
                    GetWindowRect hWndControl, rcList0
                    ptL.X = rcList0.Left: ptL.y = rcList0.Top
                    ScreenToClient hWndParent, ptL ' ptL经过转化后才能得到想要的结果！
                    
                    ' 改变对话框大小，！
                    GetWindowRect hWndParent, rcDlg ' 取得对话框矩形，再移动（实际只改变宽度！）
                    MoveWindow hWndParent, rcDlg.Left, rcDlg.Top, rcDlg.Right - rcDlg.Left + OffSetX, rcDlg.Bottom - rcDlg.Top, True
                    ' 设置 预览图片框
                    ' 设置新的，并保存图片框原来的父窗口句柄？！
                    m_picPreviewPicture.Tag = SetParent(m_picPreviewPicture.hWnd, hWndParent)
                    ' 移动图片框，Right 位置固定，高度固定！
                    If m_ppLogoPosition = ppRight Then
                        picLeft = rcDlg.Right - rcDlg.Left - 5 - m_picLogoPicture.Width \ stpX: picTop = 0
                    Else
                        picLeft = rcDlg.Right - rcDlg.Left - 8: picTop = 0
                    End If
                    MoveWindow m_picPreviewPicture.hWnd, picLeft, ptL.y, _
                        m_picPreviewPicture.Width \ stpX - 3, rcList0.Bottom - rcList0.Top, True
                    ' 加载图片
                    'm_picPreviewPicture.PaintPicture m_picPreviewPicture.Picture, 0, 0, m_picPreviewPicture.ScaleWidth, m_picPreviewPicture.ScaleHeight
                    myPaintPicture m_picPreviewPicture, False
                    ShowWindow m_picPreviewPicture.hWnd, SW_SHOW ' 显示图片框

                Case ppTop ' 预览图片框 在顶端，要移动对话框上原来的控件！
                    ' 重新取得 LIST控件矩形，可能在上面被移动了！
                    GetWindowRect hWndControl, rcList0
                    ptL.X = rcList0.Left: ptL.y = rcList0.Top
                    ScreenToClient hWndParent, ptL ' ptL经过转化后才能得到想要的结果！
                    
                    ' 对话框上所有原始控件下移
                    CDMoveOriginControl hWndParent, ID_OK, , OffSetY
                    CDMoveOriginControl hWndParent, ID_Cancel, , OffSetY
                    CDMoveOriginControl hWndParent, ID_Help, , OffSetY
                    CDMoveOriginControl hWndParent, ID_ReadOnly, , OffSetY

                    CDMoveOriginControl hWndParent, ID_FolderLabel, , OffSetY
                    CDMoveOriginControl hWndParent, ID_FolderCombo, , OffSetY
                    CDMoveOriginControl hWndParent, ID_ToolBarWin2K, , OffSetY
                    CDMoveOriginControl hWndParent, ID_List0, , OffSetY
                    
                    CDMoveOriginControl hWndParent, ID_FileNameLable, , OffSetY
                    CDMoveOriginControl hWndParent, ID_FileTypeCombo0, , OffSetY
                    CDMoveOriginControl hWndParent, ID_FileTypeLabel, , OffSetY
                    CDMoveOriginControl hWndParent, ID_FileTypeComboC, , OffSetY ' 新外观！移动对话框上文件名文本框，
                    CDMoveOriginControl hWndParent, ID_FileNameText, , OffSetY
                    
                    ' 移动工具栏！工具栏（特别注意：无法通过 CDMoveOriginControl 函数移动！）
                    hToolBar = CDGetToolBarHandle(hWndParent)
                    GetWindowRect hToolBar, rcTB
                    pt.X = rcTB.Left
                    pt.y = rcTB.Top
                    ScreenToClient hWndParent, pt
                    MoveWindow hToolBar, pt.X, pt.y + OffSetY, rcTB.Right - rcTB.Left, rcTB.Bottom - rcTB.Top, True

                    ' 改变对话框大小，！
                    GetWindowRect hWndParent, rcDlg ' 取得对话框矩形，再移动（实际只改变高度！）
                    MoveWindow hWndParent, rcDlg.Left, rcDlg.Top, rcDlg.Right - rcDlg.Left, rcDlg.Bottom - rcDlg.Top + OffSetY, True
                    ' 设置程序标志图片
                    ' 设置新的，并保存图片框原来的父窗口句柄？！
                    m_picPreviewPicture.Tag = SetParent(m_picPreviewPicture.hWnd, hWndParent)
                    ' 移动图片框，Left 位置固定，宽度固定！
                    If m_ppLogoPosition = ppLeft Then
                        picLeft = m_picLogoPicture.Width \ stpX + 5: picTop = 2: W = picLeft + 17
                    ElseIf m_ppLogoPosition = ppRight Then
                        picLeft = 5: picTop = 2: W = m_picLogoPicture.Width \ stpX + 19
                    ElseIf m_ppLogoPosition = ppTop Then
                        picLeft = 5: picTop = m_picLogoPicture.Height \ stpY + 2: W = 15
                    ElseIf m_ppLogoPosition = ppBottom Then
                        picLeft = 5: picTop = 2: W = 15
                    End If
                    MoveWindow m_picPreviewPicture.hWnd, picLeft, picTop, _
                        rcDlg.Right - rcDlg.Left - W, m_picPreviewPicture.Height \ stpY, True
                    ' 加载图片
                    'm_picPreviewPicture.PaintPicture m_picPreviewPicture.Picture, 0, 0, m_picPreviewPicture.ScaleWidth, m_picPreviewPicture.ScaleHeight
                    myPaintPicture m_picPreviewPicture, False
                    ShowWindow m_picPreviewPicture.hWnd, SW_SHOW ' 显示图片框
                    
                Case ppBottom
                    ' 重新取得 LIST控件矩形，可能在上面被移动了！
                    GetWindowRect hWndControl, rcList0
                    ptL.X = rcList0.Left: ptL.y = rcList0.Top
                    ScreenToClient hWndParent, ptL ' ptL经过转化后才能得到想要的结果！
                    
                    ' 改变对话框大小，！
                    GetWindowRect hWndParent, rcDlg ' 取得对话框矩形，再移动（实际只改变高度！）
                    MoveWindow hWndParent, rcDlg.Left, rcDlg.Top, rcDlg.Right - rcDlg.Left, rcDlg.Bottom - rcDlg.Top + OffSetY, True
                    ' 设置 预览图片框
                    Call GetClientRect(hWndParent, ClientRect) ' 用 rcDlg.Bottom 不行！！！
                    ' 设置新的，并保存图片框原来的父窗口句柄？！
                    m_picPreviewPicture.Tag = SetParent(m_picPreviewPicture.hWnd, hWndParent)
                    ' 移动图片框，Left 位置固定，宽度固定！
                    picTop = ClientRect.Bottom - OffSetY - 2
                    If m_ppLogoPosition = ppLeft Then
                        picLeft = m_picLogoPicture.Width \ stpX + 5:  W = picLeft + 7
                    ElseIf m_ppLogoPosition = ppRight Then
                        picLeft = 5:  W = m_picLogoPicture.Width \ stpX + 12
                    ElseIf m_ppLogoPosition = ppTop Then
                        picLeft = 5: W = 10
                    ElseIf m_ppLogoPosition = ppBottom Then
                        picLeft = 5: picTop = ClientRect.Bottom - OffSetY - m_picLogoPicture.Height \ stpY - 2
                        W = 10
                    End If
                    MoveWindow m_picPreviewPicture.hWnd, picLeft, picTop, _
                        ClientRect.Right - ClientRect.Left - W, m_picPreviewPicture.Height \ stpY, True
                    ' 加载图片
                    'm_picPreviewPicture.PaintPicture m_picPreviewPicture.Picture, 0, 0, m_picPreviewPicture.ScaleWidth, m_picPreviewPicture.ScaleHeight
                    myPaintPicture m_picPreviewPicture, False
                    ShowWindow m_picPreviewPicture.hWnd, SW_SHOW ' 显示图片框

            End Select
NoPreview:
            ' 设置对话框启动位置？只判断屏幕中心和所有者中心，其他不管！
            GetWindowRect hWndParent, rcDlg ' 取得对话框矩形
            W = rcDlg.Right - rcDlg.Left: H = rcDlg.Bottom - rcDlg.Top ' 对话框宽、高
            If m_dlgStartUpPosition = vbStartUpScreen Then ' 再移动。屏幕中心
                MoveWindow hWndParent, (Screen.Width \ stpX - W) \ 2, (Screen.Height \ stpY - H) \ 2, W, H, True
            ElseIf m_dlgStartUpPosition = vbStartUpOwner Then ' 所有者中心
                Dim rcOwner As RECT, T As Long ' 取得对话框的父窗口矩形
                GetWindowRect GetParent(hWndParent), rcOwner
                T = rcOwner.Top + (rcOwner.Bottom - rcOwner.Top - H) \ 2
                If T < 0 Then T = 0 ' 保证对话框不超过屏幕顶端！
                MoveWindow hWndParent, rcOwner.Left + (rcOwner.Right - rcOwner.Left - W) \ 2, T, W, H, True
            End If
            ' 启动位置，因为几种对话框都有这个属性！改用一个函数实现，效果不好！不用了！
            'setDlgStartUpPosition hWndParent, GetParent(hWndParent)
            
            ' 创建预览文本框，其大小和位置与预览图片框一样
            Dim rcP As RECT ' 决定预览文本框位置和大小。
            GetWindowRect m_picPreviewPicture.hWnd, rcP ' 没设置 m_picPreviewPicture ，这出错：(对象变量或 With 块变量未设置)
            ptL.X = rcP.Left: ptL.y = rcP.Top
            ScreenToClient hWndParent, ptL ' ptL经过转化后才能得到想要的结果！
            ' 开始创建预览文本框 去掉 Or WS_VISIBLE ，不在创建时显示！
            hWndTextView = CreateWindowEx(0, _
                "Edit", App.LegalCopyright, _
                WS_BORDER Or WS_CHILD Or WS_HSCROLL Or WS_VSCROLL Or ES_AUTOHSCROLL Or ES_AUTOHSCROLL Or ES_MULTILINE, _
                ptL.X, ptL.y, _
                rcP.Right - rcP.Left, rcP.Bottom - rcP.Top, _
                hWndParent, 0&, App.hInstance, 0&)
            ' 创建新的字体 Fixedsys - Times New Roman - MS Sans Serif
            Dim NewFont As Long
            NewFont = CreateFont(18, 0, 0, 0, _
                      366, False, False, False, _
                      DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_LH_ANGLES, _
                      ANTIALIASED_QUALITY Or PROOF_QUALITY, DEFAULT_PITCH Or FF_DONTCARE, _
                       "Times New Roman")
            ' 设置文本框字体
            SendMessage hWndTextView, WM_SETFONT, NewFont, 0
            ' 隐藏、显示对话框上的控件 m_blnHideControls(I) 的值决定是不是要隐藏！
            Call HideOrShowDlgControls(hWndParent)
            ' 设置对话框上的控件的文字，m_strControlsCaption(I) 决定其值！
            Call mSetDlgControlsCaption(hWndParent)
            
        Case CDN_SelChange ' 文件选择改变时，进行预览！
            strSelFile = SendMsgGetStr(hdr.hwndFrom, CDM_GetFilePath) ' 记录选中的文件路径
            Debug.Print "SelChange:"; strSelFile
            'Screen.MousePointer = vbHourglass ' 鼠标呈沙漏状
            If Not m_ppPreviewPosition = ppNone Then LoadPreview strSelFile, hWndParent  ' 调用函数，加载预览。
            'Screen.MousePointer = vbDefault ' 完成预览，鼠标恢复
        Case CDN_FolderChange
            Debug.Print "FolderChange:"; SendMsgGetStr(hdr.hwndFrom, CDM_GetFolderPath)
        Case CDN_ShareViolation
            Debug.Print "ShareViolation"
        Case CDN_Help
            Debug.Print "Help"
            If WhichStyle = ssOpen Then
                MsgBox "帮助：打开对话框 ！", vbInformation
            ElseIf WhichStyle = ssSave Then
                MsgBox "帮助：保存对话框 ！", vbInformation
            End If
        Case CDN_FileOK
            Debug.Print "FileOK:"; SendMsgGetStr(hdr.hwndFrom, CDM_GetFilePath)
        Case CDN_TypeChange
            Debug.Print "TypeChange"
        Case CDN_IncludeItem
            Debug.Print "IncludeItem"
        Case Else
            Debug.Print "WM_NOTIFY:   " + "&H" + Hex$(hdr.code)
    End Select
    ' 加载' 程序标志图片，非放在这里不可！！！否则效果不对！！！
    myPaintPicture m_picLogoPicture
End Function

' 移动 打开/保存 对话框上原有的控件' 若要移动对话框上原来没有的控件，需另外处理！
' 后面的可选参数设置为默认 -1 时，移动时不操作！
Private Sub CDMoveOriginControl(ByVal hWndCD As Long, ByVal ID As Long, _
    Optional ByVal X As Long = -1, Optional ByVal y As Long = -1, _
    Optional ByVal nWidth As Long = -1, Optional ByVal nHeight As Long = -1)
    
    Dim hWndControl As Long ' 控件句柄
    Dim rectControl As RECT ' 控件矩形
    Dim ptCtr As POINTAPI   ' 控件左上角在屏幕的坐标值
    Dim ptDlg As POINTAPI   ' 对话框左上角在屏幕的坐标值
    
    ' == 根据ID取得控件句柄
    hWndControl = GetDlgItem(hWndCD, ID)
    
    ' == 取得控件矩形（位置是相对对话框而言，即以对话框左上角那点为0点）并设置控件大小
    GetWindowRect hWndControl, rectControl

    ' ==取得对话框位置
    ScreenToClient hWndCD, ptDlg
    ' ==取得并设置控件位置
    ScreenToClient hWndControl, ptCtr
    ptCtr.X = rectControl.Left + ptDlg.X + IIf(X <> -1, X, 0)
    ptCtr.y = rectControl.Top + ptDlg.y + IIf(y <> -1, y, 0)
    X = ptCtr.X
    y = ptCtr.y
    nWidth = rectControl.Right - rectControl.Left + IIf(nWidth <> -1, nWidth, 0)
    nHeight = rectControl.Bottom - rectControl.Top + IIf(nHeight <> -1, nHeight, 0)
    
    ' 调用API移动控件！X，Y为相对屏幕左上角的坐标值！
    MoveWindow hWndControl, X, y, nWidth, nHeight, True
End Sub

' 移动对话框上工具栏时，取得工具栏句柄！
Private Function CDGetToolBarHandle(ByVal hDialog As Long) As Long
    CDGetToolBarHandle = FindWindowEx(hDialog, 0, "ToolBarWindow32", vbNullString)
End Function

' 对话框改变目录、选择文件时，取得 路径。
Private Function SendMsgGetStr(ByVal hWnd As Long, ByVal wMsg As Long, Optional ByVal DefLen As Long = 260) As String
    Dim TempLen As Long
    Dim TempStr As String
    Dim rc As Long
    TempLen = DefLen
    TempStr = String$(DefLen, 0)
    
    rc = SendMessage(hWnd, wMsg, TempLen, ByVal TempStr)
    If rc Then
        TempStr = StrConv(LeftB(StrConv(TempStr, vbFromUnicode), rc - 1), vbUnicode)
        'Debug.Print     TempStr   ',   Len(TempStr);
        'If   TempStr   <>   ""   Then   Debug.Print   Asc(Right(TempStr,   1))
        SendMsgGetStr = Replace$(TempStr, Chr$(0), "")
    End If
End Function


' ###############################################################################
' ### 以下为预览相关的重要函数 ##################################################
' 预览图片框图片的显示，用 PaintPicture 方法在图片框上画图片？
Private Sub myPaintPicture(picBox As PictureBox, Optional blnShowPictureOrNone As Boolean = True, Optional stdPic As StdPicture = Nothing)
    On Error Resume Next
    Dim tempPic As StdPicture
    If stdPic Is Nothing Then
        Set tempPic = picBox.Picture
    Else
        Set tempPic = stdPic
    End If
    picBox.AutoRedraw = True ' 让图片框自动刷新！
    If blnShowPictureOrNone Then
        picBox.PaintPicture tempPic, 0, 0, picBox.ScaleWidth, picBox.ScaleHeight
    Else ' 隐藏图片！
        Set picBox.Picture = Nothing
        ' 在图片框中心显示文字？有必要？！ =====================================
        Dim s As String: s = App.LegalCopyright
        picBox.ForeColor = vbBlack: picBox.FontSize = 18
        picBox.CurrentX = (picBox.ScaleWidth - picBox.TextWidth(s)) / 2
        picBox.CurrentY = (picBox.ScaleHeight - picBox.TextHeight(s)) / 2
        picBox.Print s ' =======================================================
    End If
End Sub

' 判断文件类型，以决定用什么方式预览！！！
Private Function getFileType(strFileName As String) As FileType
    Dim strExt As String    ' 文件后缀名 如 TXT
    ' 取得文件后缀名，并转化为大写！
    strExt = UCase$(Right$(strFileName, Len(strFileName) - InStrRev(strFileName, ".")))
    getFileType = ffText ' 设置函数返回默认类型！
    ' 判断文件后缀名，VB 图片框不能加载 PNG 图片 ANI 动画光标！JPE\JFIF\TIF\TIFF 未知！
    If strExt = "TXT" Then
        getFileType = ffText
    ElseIf strExt = "BMP" Or strExt = "CUR" Or strExt = "DIB" Or strExt = "EMF" Or strExt = "GIF" _
        Or strExt = "ICO" Or strExt = "JPEG" Or strExt = "JPG" Or strExt = "JPE" Or strExt = "JFIF" _
         Or strExt = "TIF" Or strExt = "TIFF" Or strExt = "WMF" Then
        getFileType = ffPicture
    ElseIf strExt = "WAV" Then
        getFileType = ffWave
    ElseIf strExt = "AIF" Or strExt = "AIFF" Or strExt = "AIFC" Or strExt = "AU" _
        Or strExt = "CDA" Or strExt = "WMA" Or strExt = "SND" Or strExt = "VOC" _
        Or strExt = "MP1" Or strExt = "MP2" Or strExt = "MP3" Or strExt = "MP4" Or strExt = "M4A" _
        Or strExt = "MID" Or strExt = "RMI" Or strExt = "MIDI" Then
        getFileType = ffAudio
    End If
End Function

' VB 取得图片大小
Private Function fncGetPicInfo(lsPicName As String, Optional hBitmap As Long = 0) As PicInfo
    Dim res As Long
    Dim bmp As BITMAP
    If hBitmap = 0 Then hBitmap = LoadPicture(lsPicName).Handle ' 函数要在窗体中，LoadPicture 才有效！
    res = GetObject(hBitmap, Len(bmp), bmp) '取得BITMAP的结构
    fncGetPicInfo.picWidth = bmp.bmWidth
    fncGetPicInfo.picHeight = bmp.bmHeight
End Function

' 加载预览，图片？文本文件？Wave 文件，音频文件？
Private Sub LoadPreview(strFileName As String, ByVal hWndParent As Long, Optional ShowFileSize As Long = &H1000)
    If FileLen(strFileName) = 0 Then Exit Sub
    On Error GoTo ErrLoad
    Dim I As Integer
    ' 根据文件类型，进行不同的预览操作。
    If getFileType(strFileName) = ffText Then ' 文本文件
        Dim FileNum      As Integer
        Dim FileSize     As Long
        Dim LoadBytes()  As Byte
        ' 读取文本文件
        FileNum = FreeFile
        Open strFileName For Binary Access Read Lock Write As FileNum
            FileSize = LOF(FileNum)
            If FileSize > ShowFileSize Then FileSize = ShowFileSize ' 预览文件的大小，可能不显示所有内容。
            ReDim LoadBytes(0 To FileSize - 1)
            Get #FileNum, 1, LoadBytes
        Close FileNum
        ' 在文本框显示文字
        Debug.Print SetTextColor(GetDC(hWndTextView), vbRed)
        Call SetWindowText(hWndTextView, LoadBytes(0))  ' 设置文字（文件本身的文字）
        Dim s As String: s = String(FileSize, 0)
        GetWindowText hWndTextView, s, FileSize
        Call SetWindowText(hWndTextView, ByVal Replace$(s, Chr$(0), "") _
            & vbCrLf & "（鹤望兰·流 省略 " & Format$(FileLen(strFileName) - FileSize, "###,###,###,##0") & " 字节）") ' 设置文字（+鹤望兰·流 标志）
        ShowWindow hWndTextView, SW_SHOW                ' 文本框可见
        ShowWindow m_picPreviewPicture.hWnd, SW_HIDE    ' 隐藏预览图片框
        For I = 0 To 2
            ShowWindow hWndButtonPlay(I), SW_HIDE
        Next I
        Exit Sub
    ElseIf getFileType(strFileName) = ffPicture Then ' 图片文件（图片预览按比例显示，并显示图片宽x高和预览比例。）
        Dim W As Long, H As Long, fWH As PicInfo ' 加载的图片宽x高
        Dim W0 As Long, H0 As Long, oldSM As ScaleModeConstants ' 图片框大小！
        Dim X As Long, y As Long, W1 As Long, H1 As Long, per As Integer, sP As StdPicture ' 按比例重新显示的图片位置，大小！
        oldSM = m_picPreviewPicture.ScaleMode: m_picPreviewPicture.ScaleMode = vbPixels ' 设置图片框ScaleMode为像素！
        W0 = m_picPreviewPicture.ScaleWidth: H0 = m_picPreviewPicture.ScaleHeight
        ' 加载图片到图片框，只为取得其尺寸！。
        Set sP = LoadPicture(strFileName)
        Set m_picPreviewPicture.Picture = sP ' LoadPicture(strFileName)
        fWH = fncGetPicInfo(strFileName, m_picPreviewPicture.Picture)
        W = fWH.picWidth: H = fWH.picHeight
        ' 判断图片框宽高比、加载的图片宽高比，其比值在相比。。。
        If (W0 / H0) / (W / H) < 1 Then ' 以' 图片框宽度为基准
            W1 = W0: H1 = W0 / W * H
            X = 0: y = (H0 - H1) / 2    ' 调整显示位置，使其居中！
            per = W0 / W * 100
        Else                            ' 以' 图片框高度为基准
            W1 = H0 / H * W: H1 = H0
            X = (W0 - W1) / 2: y = 0    ' 调整显示位置，使其居中！
            per = H0 / H * 100
        End If
        'Debug.Print "尺寸: " & W & " x " & H & " 图片框: " & W0 & " x " & H0
        Set m_picPreviewPicture.Picture = Nothing
        '' 再次加载图片到图片框，以显示。。。不再加载，否则会显示两个图（PaintPicture 方法也会显示图！）。。。
        'Set m_picPreviewPicture.Picture = sP ' LoadPicture(strFileName)
        'myPaintPicture m_picPreviewPicture
        m_picPreviewPicture.PaintPicture sP, X, y, W1, H1
        ' 图片框上显示文字
        Dim sT As String: sT = "尺寸: " & W & " x " & H & " (" & per & "%)"
        With m_picPreviewPicture
            .ForeColor = vbRed:   .FontSize = 9
            .CurrentX = 1 '(.ScaleWidth - .TextWidth(sT)) / 2
            .CurrentY = 1 '(.ScaleHeight - .TextHeight(sT)) / 2
            m_picPreviewPicture.Print sT
            '.ForeColor = vbBlue: m_picPreviewPicture.Print sT
        End With
        ShowWindow m_picPreviewPicture.hWnd, SW_SHOW
        ShowWindow hWndTextView, SW_HIDE
        For I = 0 To 2
            ShowWindow hWndButtonPlay(I), SW_HIDE
        Next I
        m_picPreviewPicture.ScaleMode = oldSM ' 还原图片框旧的 ScaleMode 值。
        Exit Sub
    ElseIf getFileType(strFileName) = ffWave Then ' 波形文件
        Create3Buttons hWndParent ' 创建 3 个播放控制按钮。
        ShowWindow m_picPreviewPicture.hWnd, SW_SHOW ' 显示预览图片框
        Set m_picPreviewPicture.Picture = Nothing
        ShowWindow hWndTextView, SW_HIDE             ' 隐藏文本框
        
        ' 先停止声音，因为前面可能播放过！！！！' 注意：这里不直接调用 B3Button_Click 2 ！
        PlayAudio strSelFile, 2
        sndPlaySound vbNullString, 0 ' 停止 Wave 文件播放。
        ' =========================================================================================
        ' ==== 画出Wave文件波形（可去掉）==========================================================
        ' =========================================================================================
        ' 画出波形！用一个模块完成，可以去掉此功能！！！！！！！！
        MDrawWaves.DrawWaves strFileName, m_picPreviewPicture  ' 为啥不行？原因：一定要设置 ScaleMode = vbTwips ！！！
        ' =========================================================================================
        ' ==== 画出Wave文件波形（可去掉）==========================================================
        ' =========================================================================================
        Exit Sub
    ElseIf getFileType(strFileName) = ffAudio Then ' 音频文件
        Create3Buttons hWndParent ' 创建 3 个播放控制按钮。
        ShowWindow m_picPreviewPicture.hWnd, SW_SHOW ' 显示预览图片框
        Set m_picPreviewPicture.Picture = Nothing
        ShowWindow hWndTextView, SW_HIDE             ' 隐藏文本框
        
        ' 先停止声音，因为前面可能播放过！！！！' 注意：这里不直接调用 B3Button_Click 2 ！
        PlayAudio strSelFile, 2
        sndPlaySound vbNullString, 0 ' 停止 Wave 文件播放。
        Exit Sub
    End If

ErrLoad:
    Debug.Print "LoadPreview Error " & Err.Number & ": " & Err.Description
End Sub
' 创建播放控制 3 个按钮
Private Sub Create3Buttons(ByVal hWndParent As Long)
    
    ' 不能重复创建，否则，响应消息时，hWndButtonPlay(0)。。。值变了，而 lParam 的值还是第一次创建时的，导致无法响应消息！
    If hWndButtonPlay(0) Then ' 要显示按钮！否则不见了？！
        ShowWindow hWndButtonPlay(0), SW_SHOW
        ShowWindow hWndButtonPlay(1), SW_SHOW
        ShowWindow hWndButtonPlay(2), SW_SHOW
        Exit Sub
    End If
    'MsgBox "创建播放控制 3 个按钮"
    ' 创建播放控制 3 个按钮，其位置由文件类型标签位置决定。大小固定。
    Dim rcP As RECT, ptL As POINTAPI ' 决定播放控制 3 个按钮位置和大小。
    Dim NewFont As Long
    GetWindowRect GetDlgItem(hWndParent, ID_FileTypeLabel), rcP
    ptL.X = rcP.Left: ptL.y = rcP.Top
    ScreenToClient hWndParent, ptL ' ptL经过转化后才能得到想要的结果！
    ' 开始创建 播放控制 3 个按钮 加 Or WS_VISIBLE ，在创建时显示！
    hWndButtonPlay(0) = CreateWindowEx(WS_EX_STATICEDGE Or WS_EX_TOPMOST, _
        "Button", "4", _
        WS_CHILD Or WS_VISIBLE, _
        ptL.X, ptL.y + rcP.Bottom - rcP.Top + 5, _
        25, 22, _
        hWndParent, 0&, App.hInstance, 0&)
    hWndButtonPlay(1) = CreateWindowEx(WS_EX_STATICEDGE Or WS_EX_TOPMOST, _
        "Button", ";", _
        WS_CHILD Or WS_VISIBLE, _
        ptL.X + 25, ptL.y + rcP.Bottom - rcP.Top + 5, _
        25, 22, _
        hWndParent, 0&, App.hInstance, 0&)
    hWndButtonPlay(2) = CreateWindowEx(WS_EX_STATICEDGE Or WS_EX_TOPMOST, _
        "Button", "<", _
        WS_CHILD Or WS_VISIBLE, _
        ptL.X + 50, ptL.y + rcP.Bottom - rcP.Top + 5, _
        25, 22, _
        hWndParent, 0&, App.hInstance, 0&)
    ' 创建新的字体 Webdings - Fixedsys - Times New Roman - MS Sans Serif
    NewFont = CreateFont(18, 0, 0, 0, _
              366, False, False, False, _
              DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_LH_ANGLES, _
              ANTIALIASED_QUALITY Or PROOF_QUALITY, DEFAULT_PITCH Or FF_DONTCARE, _
               "Webdings")
    ' 设置字体
    SendMessage hWndButtonPlay(0), WM_SETFONT, NewFont, 0
    SendMessage hWndButtonPlay(1), WM_SETFONT, NewFont, 0
    SendMessage hWndButtonPlay(2), WM_SETFONT, NewFont, 0
End Sub

' 注意：仅对颜色对话框。。。设置对话框启动位置？只判断屏幕中心和所有者中心，其他不管！
Private Sub setDlgStartUpPosition(ByVal hWndDlg As Long, ByVal hWndParent As Long)
    Dim W As Long, H As Long, T As Long
    Dim rcDlg As RECT, rcOwner As RECT
    GetWindowRect hWndDlg, rcDlg ' 取得对话框矩形
    W = rcDlg.Right - rcDlg.Left: H = rcDlg.Bottom - rcDlg.Top ' 对话框宽、高
    If m_dlgStartUpPosition = vbStartUpScreen Then ' 再移动。屏幕中心
        MoveWindow hWndDlg, (Screen.Width \ Screen.TwipsPerPixelX - W) \ 2, (Screen.Height \ Screen.TwipsPerPixelY - H) \ 2, W, H, True
    ElseIf m_dlgStartUpPosition = vbStartUpOwner Then ' 所有者中心
         ' 取得对话框的父窗口矩形
        GetWindowRect hWndParent, rcOwner
        T = rcOwner.Top + (rcOwner.Bottom - rcOwner.Top - H) \ 2
        If T < 0 Then T = 0 ' 保证对话框不超过屏幕顶端！
        MoveWindow hWndDlg, rcOwner.Left + (rcOwner.Right - rcOwner.Left - W) \ 2, T, W, H, True
    End If
End Sub


' =========================================================================================
' ==== 自定义对话框上的控件隐藏、显示，改变文字等！（可去掉）==============================
' =========================================================================================
' 隐藏、显示对话框上的控件 m_blnHideControls(I) 的值决定是不是要隐藏！而且，Flags 属性也有影响。
Private Sub HideOrShowDlgControls(ByVal hWndDlg As Long)
    Dim I As Integer, ctrl_ID As Variant ' 控件ID，转到一个数组，好操作！
    ' (0 to 12) 分别为：有些不能隐藏。用《》标出。只剩下 (0 to 8)
    ' “查找范围(&I)”标签      -- 目录下拉框                         -- 《工具栏》        1
    ' 《快捷目录区（版本>=Win2K）》 -- 《列表框（列出文件的最大区域）》                    2
    ' “文件名(&N)”标签        -- 《“文件名(&N)”文本框》           -- “确定(&O)”按键  1
    ' “文件类型(&T)”标签      -- “文件类型(&T)”下拉框（新外观时） -- “取消(&C)”按键  0
    ' “只读”多选框            -- “帮助(&H)”按键                                        0
    ctrl_ID = Array(ID_FolderLabel, ID_FolderCombo, _
                ID_FileNameLable, ID_OK, _
                ID_FileTypeLabel, ID_FileTypeCombo0, ID_Cancel, _
                ID_ReadOnly, ID_Help)
    ' 隐藏控件
    For I = 0 To 8
        If m_blnHideControls(I) Then Call SendMessage(hWndDlg, CDM_HideControl, ctrl_ID(I), ByVal 0&)
    Next I
    Set ctrl_ID = Nothing
End Sub
' 设置对话框上的控件的文字。m_strControlsCaption(I) 决定其值，默认为空，不改变原始值！
Private Sub mSetDlgControlsCaption(ByVal hWndDlg As Long)
    Dim I As Integer, ctrl_ID As Variant ' 控件ID，转到一个数组，好操作！
    ' (0 to 12) 分别为：有些不能设置文字。用《》标出。只剩下 (0 to 6)
    ' “查找范围(&I)”标签      -- 《目录下拉框》                        -- 《工具栏》        2
    ' 《快捷目录区（版本>=Win2K）》 -- 《列表框（列出文件的最大区域）》                       2
    ' “文件名(&N)”标签        -- 《“文件名(&N)”文本框》              -- “确定(&O)”按键  1
    ' “文件类型(&T)”标签      -- 《“文件类型(&T)”下拉框（新外观时）》-- “取消(&C)”按键  1
    ' “只读”多选框            -- “帮助(&H)”按键                                           0
    ctrl_ID = Array(ID_FolderLabel, _
                ID_FileNameLable, ID_OK, _
                ID_FileTypeLabel, ID_Cancel, _
                ID_ReadOnly, ID_Help)
    For I = 0 To 6
        If Len(m_strControlsCaption(I)) <> 0 Then Call SendMessage(hWndDlg, CDM_SetControlText, ctrl_ID(I), ByVal m_strControlsCaption(I))
    Next I
    Set ctrl_ID = Nothing
End Sub
' =========================================================================================
' ==== 自定义对话框上的控件隐藏、显示，改变文字等！（可去掉）==============================
' =========================================================================================



' =========================================================================================
' ==== 声音文件的播放（可去掉）============================================================
' =========================================================================================
' 三个播放按钮单击事件？
Private Sub B3Button_Click(Index As Integer)
'    PlayAudio strSelFile, Index' 都用 mciSendString 函数播放？！！！
    If Index = 0 Then ' 播放
        If getFileType(strSelFile) = ffWave Then ' Wave 文件
            sndPlaySound strSelFile, SND_FILENAME Or SND_ASYNC
        Else
            PlayAudio strSelFile, 0
        End If
    ElseIf Index = 1 Then ' 暂停
        If getFileType(strSelFile) = ffWave Then ' Wave 文件
            sndPlaySound vbNullString, 0 ' 不知道怎么暂停！这里停止！！！
        Else
            PlayAudio strSelFile, 1
        End If
    Else ' 停止
        If getFileType(strSelFile) = ffWave Then ' Wave 文件
            sndPlaySound vbNullString, 0
        Else
            PlayAudio strSelFile, 2
        End If
    End If
End Sub
' 高级媒体播放，播放音频文件。反应太慢了？！！！Wave 文件还是用其他函数波放？？？！！！
Private Sub PlayAudio(strFileName As String, Optional setStatus As Integer = 0)
    If Len(Dir$(strFileName)) = 0 Then Exit Sub
    Const ALIAS_NAME As String = "mySound"
    Dim rt As Long
    If setStatus = 0 Then ' 播放
        If getPalyStatus(ALIAS_NAME) = IsStopped Then
            ' 打开并从头开始播放。注意：给 strFileName 加双引号，否则有的文件名有空格，无法播放！
            rt = mciSendString("open " & """" & strFileName & """" & " alias " & ALIAS_NAME, vbNullString, 0, 0)
            rt = mciSendString("play " & ALIAS_NAME, vbNullString, 0, 0)
            ' 取得媒体文件长度？！
            Dim RefStr1 As String * 80
            mciSendString "status " & ALIAS_NAME & " length", RefStr1, Len(RefStr1), 0
            Debug.Print "总长度：" & Val(RefStr1)
        ElseIf getPalyStatus(ALIAS_NAME) = IsPaused Then
            ' 继续播放
            rt = mciSendString("resume " & ALIAS_NAME, vbNullString, 0, 0)
            ' 获取当前播放进度，相对文件长度而言？！
            Dim RefStr2 As String * 80
            mciSendString "status " & ALIAS_NAME & " position", RefStr2, Len(RefStr2), 0
            Debug.Print "已播放：" & Val(RefStr2)
        Else
            ' 停止播放并关闭声音！
            rt = mciSendString("stop " & ALIAS_NAME, vbNullString, 0, 0)
            rt = mciSendString("close " & ALIAS_NAME, vbNullString, 0, 0)
        End If
    ElseIf setStatus = 1 Then ' 暂停
        rt = mciSendString("pause " & ALIAS_NAME, vbNullString, 0, 0)
    Else ' 停止
        rt = mciSendString("stop " & ALIAS_NAME, vbNullString, 0, 0)
        rt = mciSendString("close " & ALIAS_NAME, vbNullString, 0, 0)
    End If
End Sub
' 获得当前媒体的状态。是在播放？暂停？停止？
Private Function getPalyStatus(Optional strAlias As String = "mySound") As PlayStatus
    Dim sl As String * 255
    mciSendString "status " & strAlias & " mode", sl, Len(sl), 0
    If UCase$(Left$(sl, 7)) = "PLAYING" Or Left$(sl, 2) = "播放" Then
        getPalyStatus = IsPlaying
    ElseIf UCase$(Left$(sl, 6)) = "PAUSED" Or Left$(sl, 2) = "暂停" Then
        getPalyStatus = IsPaused
    Else
        getPalyStatus = IsStopped
    End If
End Function
' =========================================================================================
' ==== 声音文件的播放（可去掉）============================================================
' =========================================================================================



' =========================================================================================
' ==== 字体对话框 （单独）=================================================================
' =========================================================================================
' 显示对话框之前。自定义字体对话框外观。
Private Sub CustomizeFontDialog(ByVal hWnd As Long)
    Dim rcDlg As RECT, hWndParent As Long
    Dim pt As POINTAPI, W As Long, H As Long
    ' 对话框的父窗口句柄。
    hWndParent = GetParent(hWnd)
    ' 设置对话框启动位置？只判断屏幕中心和所有者中心，其他不管！
    GetWindowRect hWnd, rcDlg ' 取得对话框矩形
    W = rcDlg.Right - rcDlg.Left: H = rcDlg.Bottom - rcDlg.Top + 120 ' 对话框宽、高（高度要加个预览文本框高！）
    If m_dlgStartUpPosition = vbStartUpScreen Then ' 再移动。屏幕中心
        MoveWindow hWnd, (Screen.Width \ Screen.TwipsPerPixelX - W) \ 2, (Screen.Height \ Screen.TwipsPerPixelY - H) \ 2, W, H, True
    ElseIf m_dlgStartUpPosition = vbStartUpOwner Then ' 所有者中心
        Dim rcOwner As RECT, T As Long ' 取得对话框的父窗口矩形
        GetWindowRect hWndParent, rcOwner
        T = rcOwner.Top + (rcOwner.Bottom - rcOwner.Top - H) \ 2
        If T < 0 Then T = 0 ' 保证对话框不超过屏幕顶端！
        MoveWindow hWnd, rcOwner.Left + (rcOwner.Right - rcOwner.Left - W) \ 2, T, W, H, True
    Else ' 这个特别，要处理。高度要变，否则，添加的文本框无法显示！
        MoveWindow hWnd, rcDlg.Left, rcDlg.Top, W, H, True
    End If
    ' 启动位置，因为几种对话框都有这个属性！改用一个函数实现，效果不好！不用了！
    'setDlgStartUpPosition hWndParent, GetParent(hWndParent)
    
    ' 创建预览文本框，其大小固定!！
    Dim rcP As RECT, ptL As POINTAPI  ' 决定预览文本框位置和大小，rcL 那个说明标签矩形！ID = 1093 ?。
    GetWindowRect GetDlgItem(hWnd, enumFONT_CTL.stc_Description), rcP
    ptL.X = rcP.Left: ptL.y = rcP.Top
    ScreenToClient hWnd, ptL ' ptL经过转化后才能得到想要的结果！
    ' 开始创建预览文本框  加 Or WS_VISIBLE ，在创建时显示！ Or ES_READONLY 设置为只读！
    Dim sT As String
    'Debug.Print " rcDlg.Bottom - rcP.Bottom + 20 " & rcDlg.Bottom - rcP.Bottom + 20 ' 文本框高度会变？！不正常？！
    sT = App.LegalCopyright & vbCrLf _
        & "一二三四五六七八九十" & vbCrLf & "壹贰叁肆伍陆柒捌玖拾" & vbCrLf _
        & "ABCDEFGHILMNOPQRSTUVWXYZ" & vbCrLf & "abcdefghilmnopqrstuvwxyz" & vbCrLf & "0123456789"
    hWndFontPreview = CreateWindowEx(WS_EX_STATICEDGE Or WS_EX_TOPMOST, _
        "Edit", sT, _
        WS_BORDER Or WS_CHILD Or WS_VISIBLE Or WS_HSCROLL Or WS_VSCROLL Or ES_AUTOHSCROLL Or ES_AUTOHSCROLL Or ES_MULTILINE Or ES_LEFT Or ES_WANTRETURN, _
        ptL.X, ptL.y + (rcP.Bottom - rcP.Top), _
        W - 20, 120, _
        hWnd, 0&, App.hInstance, &H520)
    ' 创建新的字体，默认字体！！！
    Dim NewFont As Long, lpLF As LOGFONT
    With lpLF
        .lfCharSet = 134
        '.lfFaceName = "宋体"
        .lfItalic = False
        .lfStrikeOut = False
        .lfUnderline = False
        .lfWeight = 520
    End With
    NewFont = CreateFontIndirect(lpLF)
    ' 设置文本框字体
    SendMessage hWndFontPreview, WM_SETFONT, NewFont, 0
End Sub
' 设置字体对话框预览。（文字效果动态变化。截取 WM_COMMAND 消息，判断字体格式有哪些变化！）
Private Sub mSetFontPreview(ByVal hWnd As Long, Optional wP As Long = 0&)
    Dim hFontToUse As Long, hdc As Long, RetValue As Long
    Dim lpLF As LOGFONT
    Dim tBuf As String * 80, sFontName As String
    Dim iIndex As Long, dwRGB As Long ' dwRGB 字体颜色值！
    
     ' 取得选择的字体信息
    SendMessage hWnd, WM_CHOOSEFONT_GETLOGFONT, wP, lpLF
    hFontToUse = CreateFontIndirect(lpLF): If hFontToUse = 0 Then Exit Sub
    hdc = GetDC(hWnd)
    SelectObject hdc, hFontToUse
    RetValue = GetTextFace(hdc, 79, tBuf)
    sFontName = Mid$(tBuf, 1, RetValue)

    ' 取得选择的字体颜色
    iIndex = SendDlgItemMessage(hWnd, enumFONT_CTL.cbo_Color, CB_GETCURSEL, 0&, 0&)    ' cmb4
    If iIndex <> CB_ERR Then
        dwRGB = SendDlgItemMessage(hWnd, enumFONT_CTL.cbo_Color, CB_GETITEMDATA, iIndex, 0&)
    End If
    ' 创建新的字体，颜色信息没有！！！
    Dim NewFont As Long
    NewFont = CreateFontIndirect(lpLF)
'    NewFont = CreateFont(Abs(lpLF.lfHeight * (72 / GetDeviceCaps(hDC, LOGPIXELSY))), 0, 0, 0, _
              lpLF.lfWeight, lpLF.lfItalic, lpLF.lfUnderline, lpLF.lfStrikeOut, _
              lpLF.lfCharSet, OUT_DEFAULT_PRECIS, CLIP_LH_ANGLES, _
              ANTIALIASED_QUALITY Or PROOF_QUALITY, DEFAULT_PITCH Or FF_DONTCARE, _
              sFontName)
    ' 设置文本框字体
    SendMessage hWndFontPreview, WM_SETFONT, NewFont, 0
    ' 设置文本框文字颜色（不知为什么，预览颜色改变。。。无法实现！！！）
'    SendMessage hWndFontPreview, 4103, 0, ByVal dwRGB
    If SetTextColor(GetDC(hWndFontPreview), dwRGB) = &HFFFF Then MsgBox "失败：设置文字颜色出错！", vbCritical
'    frmMain.txtNewCaption(2).ForeColor = dwRGB
'    frmMain.txtNewCaption(1).ForeColor = GetTextColor(GetDC(frmMain.txtNewCaption(2).hWnd))
'    Dim cDC As Long, chWnd As Long
'    chWnd = GetDlgItem(hWnd, enumFONT_CTL.btn_Apply)
'    cDC = GetDC(cDC)
'    Call SetTextColor(cDC, dwRGB)
'    Call SendMessage(chWnd, CDM_SetControlText, enumFONT_CTL.btn_Apply, ByVal "m_strControlsCaption(I)")
'    If dwRGB = GetTextColor(cDC) Then frmMain.BackColor = GetTextColor(cDC)
'    Dim sl As Long ' 文本框中文字个数，要取得，暂时固定一个值！。
'    sl = 1024
'    Dim s As String: s = String(sl, 0)
'    GetWindowText hWndFontPreview, s, sl
'    Debug.Print Replace$(s, Chr$(0), "")
'    ' 重新设置文字
'    SendMessage hWndFontPreview, WM_SETTEXT, -1, ByVal s & vbCrLf & App.LegalCopyright

    ' 释放资源
    ReleaseDC hWnd, hdc
End Sub

Private Function LOWORD(Param As Long) As Long
    LOWORD = Param And &HFFFF&
End Function
Private Function HIWORD(Param As Long) As Long
    HIWORD = Param \ &H10000 And &HFFFF&
End Function
' =========================================================================================
' ==== 字体对话框 （单独）=================================================================
' =========================================================================================
