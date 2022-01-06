Attribute VB_Name = "MCDHook"
Option Explicit
' --- API ���� ����
' �ͷų����ڴ�
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
' ȡ�ÿؼ������Ļ���Ͻǵ�����ֵ������λ�����أ�����
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWndParent As Long, ByVal hWndChildAfter As Long, ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function UpdateWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, lpString As Any) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal e As Long, ByVal O As Long, ByVal W As Long, ByVal I As Long, ByVal u As Long, ByVal s As Long, ByVal C As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal f As String) As Long

' VB ȡ��ͼƬ��С
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long

' =========================================================================================
' ==== �����ļ��Ĳ��ţ���ȥ����============================================================
' =========================================================================================
'API ���� ʹ��PlaySound������������
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, _
    ByVal hModule As String, ByVal dwFlags As Long) As Long
'API ���� ʹ��sndPlaySound������������������ PlaySound �������Ӽ�����
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, _
    ByVal uFlags As Long) As Long
Private Declare Function sndStopSound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszNull As Long, ByVal uFlags As Long) As Long
'�ر�����
'sndPlaySound Null, SND_ASYNC
'PlaySound 0,0,0
' �߼�ý�岥�ź���
Private Declare Function mciExecute Lib "winmm.dll" (ByVal lpstrCommand As String) As Long
' mciSendString ���������Ŷ�ý���ļ���APIָ����Բ���MPEG,AVI,WAV,MP3,�ȵ�
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
' Multimedia Command Strings: http://msdn.microsoft.com/en-us/library/ms712587.aspx
' MCI Command Strings:http://msdn.microsoft.com/en-us/library/ms710815(VS.85).aspx

' --- for PlaySound \ sndPlaySound
Private Const SND_ASYNC = &H1 ' play asynchronously �ڲ��ŵ�ͬʱ����ִ���Ժ�����
Private Const SND_FILENAME = &H20000 ' name is a file name
Private Const SND_LOOP = &H8 ' loop the sound until next sndPlaySound һֱ�ظ�����������ֱ���ú�����ʼ���ŵڶ�������Ϊֹ
Private Const SND_MEMORY = &H4         '  lpszSoundName points to a memory file �����ڴ��е�����, Ʃ����Դ�ļ��е�����
Private Const SND_NODEFAULT = &H2         '  silence not default, if sound not found
Private Const SND_NOSTOP = &H10        '  don't stop any currently playing sound
Private Const SND_NOWAIT = &H2000      '  don't wait if the driver is busy
Private Const SND_PURGE = &H40               '  purge non-static events for task
Private Const SND_RESERVED = &HFF000000  '  In particular these flags are reserved
Private Const SND_RESOURCE = &H40004     '  name is a resource name or atom
Private Const SND_SYNC = &H0         '  play synchronously (default) ����������֮����ִ�к�������
Private Const SND_TYPE_MASK = &H170007
Private Const SND_VALID = &H1F        '  valid flags          / ;Internal /
Private Const SND_VALIDFLAGS = &H17201F    '  Set of valid flag bits.  Anything outside

Private Enum PlayStatus ' ��������״̬��
    IsPlaying = 0
    IsPaused = 1
    IsStopped = 2
End Enum
' =========================================================================================
' ==== �����ļ��Ĳ��ţ���ȥ����============================================================
' =========================================================================================


' =========================================================================================
' ==== ����Ի��� ��������=================================================================
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
' hDC ����Ҫ��
' ˵�� ���õ�ǰ�ı���ɫ��������ɫҲ��Ϊ��ǰ��ɫ�� ����ֵ Long���ı�ɫ��ǰһ��RGB��ɫ�趨��CLR_INVALID��ʾʧ�ܡ�

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
Rem per maggior praticit� ho enumerato tutti i controlli della
Rem finestra Carattere
Rem ------------------------------------------------------------
Public Enum enumFONT_CTL ' ����Ի����ϵĿؼ� ID
    stc_FontName = 1088 ' ����(&F): ��ǩ
    edt_FontName = 1001 ' �������� �ı��򣿣�
    cbo_FontName = &H470  ' �������� �����򣿣�66672
    
    stc_BoldItalic = 1089 ' ����(&Y): ��ǩ
    edt_BoldItalic = 1001 ' ���� �ı��򣿣�
    cbo_BoldItalic = &H471  ' ���� �����򣿣�66673
    
    stc_Size = 1090 ' ��С(&S): ��ǩ
    edt_Size = 1001 ' ��С �ı��򣿣�
    cbo_Size = &H472  ' ��С �����򣿣�66674
    
    btn_Ok = 1 ' ȷ��(&O) ��ť
    btn_Cancel = 2 ' ȡ��(&C) ��ť
    btn_Apply = 1026 ' Ӧ��(&A) ��ť
    btn_Help = 1038 ' ����(&H) ��ť
    
    btn_Effects = 1072 ' Ч�� ��Ͽ�
    btn_Strikethrough = &H410 ' ɾ����(&K) ��ť
    btn_Underline = &H411 ' �»���(&U) ��ť
    stc_Color = &H443 ' ��ɫ(&C): ��ǩ
    cbo_Color = &H473 ' ��ɫ �����򣿣�66675
    
    btn_Sample = 1073 ' ʾ����Ͽ�
    stc_SampleText = &H444 ' ʾ����ǩ��΢���������
    
    stc_Charset = 1094 ' �ַ���(&R): ��ǩ
    cbo_Charset = &H474 ' �ַ���������
    stc_Description = 1093 ' ����������ǩ��������������ʾ����ӡʱ��ʹ����ӽ���ƥ�����塣

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
' ==== ����Ի��򣨵�����==================================================================
' =========================================================================================



' =========================================================================================
' ==== ��ɫ�Ի��� ��������=================================================================
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
' ==== ��ɫ�Ի��򣨵�����=================================================================
' =========================================================================================


' --- ���� ����
' for Windows ��Ϣ ����
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

' for �Ի����ϵ���Ϣ
Private Const CDM_First = (WM_USER + 100)                   '/---
Private Const CDM_GetSpec = (CDM_First + &H0)               'ȡ���ļ���
Private Const CDM_GetFilePath = (CDM_First + &H1)           'ȡ���ļ�����Ŀ¼
Private Const CDM_GetFolderPath = (CDM_First + &H2)         'ȡ��·��
Private Const CDM_GetFolderIDList = (CDM_First + &H3)       '
Private Const CDM_SetControlText = (CDM_First + &H4)        '���ÿؼ��ı�
Private Const CDM_HideControl = (CDM_First + &H5)           '���ؿؼ�
Private Const CDM_SetDefext = (CDM_First + &H6)             '
Private Const CDM_Last = (WM_USER + 200)                    '\---

Private Const CDN_First = (-601)                            '/---
Private Const CDN_InitDone = (CDN_First - &H0)              '��ʼ�����
Private Const CDN_SelChange = (CDN_First - &H1)             'ѡ���ļ��ı�
Private Const CDN_FolderChange = (CDN_First - &H2)          'Ŀ¼�ı�
Private Const CDN_ShareViolation = (CDN_First - &H3)        '
Private Const CDN_Help = (CDN_First - &H4)                  '���˰���
Private Const CDN_FileOK = (CDN_First - &H5)                '����ȷ��
Private Const CDN_TypeChange = (CDN_First - &H6)            '�������͸ı�
Private Const CDN_IncludeItem = (CDN_First - &H7)           '
Private Const CDN_Last = (-699)                             '\---
  
' for �Ի����Ͽؼ��� ID
Private Const ID_FolderLabel   As Long = &H443              '�����ҷ�Χ(&I)����ǩ
Private Const ID_FolderCombo   As Long = &H471              'Ŀ¼������
Private Const ID_ToolBar       As Long = &H440              '���������ر�ע�⣺�޷�ͨ�� CDMoveOriginControl �����ƶ�����
Private Const ID_ToolBarWin2K  As Long = &H4A0              '���Ŀ¼�����汾>=Win2K��

' �б���г��ļ����������
Private Const ID_List0         As Long = &H460              ' ʹ�������Ч������
Private Const ID_List1         As Long = &H461
Private Const ID_List2         As Long = &H462

Private Const ID_OK            As Long = 1                  '��ȷ��(&O)������
Private Const ID_Cancel        As Long = 2                  '��ȡ��(&C)������
Private Const ID_Help          As Long = &H40E              '������(&H)������
Private Const ID_ReadOnly      As Long = &H410              '��ֻ������ѡ��

Private Const ID_FileTypeLabel As Long = &H441              '���ļ�����(&T)����ǩ
Private Const ID_FileNameLable As Long = &H442              '���ļ���(&N)����ǩ
'���ļ�����(&T)��������
Private Const ID_FileTypeCombo0 As Long = &H470             ' ʹ�������Ч������
Private Const ID_FileTypeCombo1 As Long = &H471
Private Const ID_FileTypeCombo2 As Long = &H472
Private Const ID_FileTypeComboC As Long = &H47C             '���ļ���(&N)���ı���
Private Const ID_FileNameText  As Long = &H480              '���ļ���(&N)���ı��������ʱ����������

' for SendMessage ȡ�ø�ѡ���Ƿ�ѡ�У�
Private Const BM_GETCHECK = &HF0

' for CreateWindowEx ����Ԥ���ı���
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
Private Const ES_MULTILINE = &H4&       ' �ı��������
Private Const ES_READONLY = &H800&      ' ���༭�����ó�ֻ����
Private Const ES_CENTER = &H1&          ' �ı���ʾ����
Private Const ES_WANTRETURN = &H1000&   ' ʹ���б༭�����ջس������벢���С������ָ���÷�񣬰��س�����ѡ��ȱʡ�����ť���������ᵼ�¶Ի���Ĺرա�

' --- for CreateFont ������Ϣ����
Private Const CLIP_LH_ANGLES            As Long = 16 ' �ַ���ת����Ҫ��
Private Const PROOF_QUALITY             As Long = 2
Private Const TRUETYPE_FONTTYPE         As Long = &H4
Private Const ANTIALIASED_QUALITY       As Long = 4
Private Const DEFAULT_CHARSET           As Long = 1
Private Const FF_DONTCARE = 0    '  Don't care or don't know.
Private Const DEFAULT_PITCH = 0
Private Const OUT_DEFAULT_PRECIS = 0

' --- ö�� ����
Public Enum PreviewPosition ' Ԥ��ͼƬ��λ��
    ppNone = -1 ' ��Ϊ��ֵʱ������ʾ��
    ppTop = 0
    ppLeft = 1
    ppRight = 2
    ppBottom = 3
End Enum
Public Enum DialogStyle ' �Ի����񣬴򿪣����棿���壿��ɫ��
    ssOpen = 0
    ssSave = 1
    ssFont = 2
    ssColor = 3
End Enum
Private Enum FileType
    ffText = 0      ' �ı� Ԥ����Ĭ��ֵ���κ��ļ��������ı���ʽ�򿪣�����
    ffPicture = 1   ' ͼƬ Ԥ��
    ffWave = 2      ' Wave �����ļ� Ԥ���������������Σ���
    ffAudio = 3     ' һ����Ƶ�ļ�����Ӳ��š���ͣ��ֹͣ��ť������Ԥ����API������������
End Enum

' --- �ṹ�� ����
' for CopyMemory ȡ�öԻ�����Щ�ؼ��ı䣿
Private Type NMHDR
    hwndFrom   As Long
    idFrom   As Long
    code   As Long
End Type
' ���ꣿ
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
Private Type BITMAP ' ȡ��BITMAP�ṹ��
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type
Private Type PicInfo ' ͼƬ����
    picWidth As Long
    picHeight As Long
End Type

' --- ˽�б��� ����
Private procOld As Long ' ����ԭ �������Եı�������ʵ��Ĭ�ϵ� ���庯�� �ĵ�ַ
Private hWndTextView As Long ' ��̬������Ԥ���ı��� ���
Private hWndButtonPlay(0 To 2) As Long ' 3 �����Ű�ť ���
Private strSelFile As String ' ѡ�е��ļ�·��

' ==== ����Ի��� ��������=================================================================
Private hWndFontPreview As Long ' ����Ԥ���ı���
' ==== ����Ի��� ��������=================================================================

' --- �������� ����...Ϊ CCommonDialog ���񣡣�
Public IsReadOnlyChecked As Boolean ' ָʾ�Ƿ�ѡ��ֻ����ѡ��
Public WhichStyle As DialogStyle ' �Ի����񣬴򿪣����棿���壿��ɫ��

' �ر��ر�ע�⣺ͼƬ�����ʱ������ͼƬ������ڶ��ε����Ի���ʱͼƬ����ʧ���������Ҵ�����Ҫ�������յ�ͼƬ�򣨲����κ��£������裡������
Public m_picLogoPicture As PictureBox ' �����־ͼƬ��ͼƬ
Public m_picPreviewPicture As PictureBox ' Ԥ��ͼƬ��ͼƬ
Public m_ppLogoPosition As PreviewPosition ' �����־ͼƬ��λ��
Public m_ppPreviewPosition As PreviewPosition ' Ԥ��ͼƬ��λ��
Public m_dlgStartUpPosition As StartUpPositionConstants ' �Ի�������λ�ã�
Public m_blnHideControls(0 To 8) As Boolean ' �Ƿ����ضԻ����ϵĿؼ�������ȥ����
Public m_strControlsCaption(0 To 6) As String ' �Ի����ϵĿؼ������֣�����ȥ����

' ################################################################################################
' �ص�������������ȡ��Ϣ���ö�̬�����Ŀؼ�������Ӧ��Ϣ����ע�⣺�ǽ�ȡ�Ի�����Ϣ����
' ################################################################################################
Private Function WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, _
                                              ByVal wParam As Long, ByVal lParam As Long) As Long
    ' ȷ�����յ�����ʲô��Ϣ
    Select Case iMsg
        Case WM_COMMAND ' ����
            Dim I As Integer
            For I = 0 To 2
                If lParam = hWndButtonPlay(I) Then Call B3Button_Click(I)
            Next I
'        Case WM_LBUTTONDOWN ' ����������
'            Debug.Print "WM_LBUTTONDOWN " & lParam
    End Select
  
    ' �������������Ҫ����Ϣ���򴫵ݸ�ԭ���Ĵ��庯������
    WindowProc = CallWindowProc(procOld, hWnd, iMsg, wParam, lParam)

End Function
' ���ÿ�ʼ�ͽ������������̣�����
Private Sub CDHook(ByVal hWnd As Long)
    ' ����procOld���������洢���ڵ�ԭʼ�������Ա�ָ�
    ' ������ SetWindowLong ��������ʹ���� GWL_WNDPROC ��������������������࣬ͨ����������
    ' ����ϵͳ�����������Ϣ���ɻص����� (WindowProc) ����ȡ�� AddressOf�ǹؼ���ȡ�ú�����ַ
    procOld = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf WindowProc)
             ' AddressOf��һԪ����������ڹ��̵�ַ���͵� API ����֮ǰ���ȵõ��ù��̵ĵ�ַ
End Sub
Private Sub CDUnHook(ByVal hWnd As Long)
    ' �˾�ؼ����Ѵ��ڣ����Ǵ��壬���Ǿ��о������һ�ؼ��������Ը�ԭ
    Call SetWindowLong(hWnd, GWL_WNDPROC, procOld)
End Sub
' ################################################################################################
' �ص�������������ȡ��Ϣ���ö�̬�����Ŀؼ�������Ӧ��Ϣ��
' ################################################################################################


' �ص��������Ի�����ʾʱҪʹ�ã�����
Public Function CDCallBackFun(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
On Error GoTo CDCallBack_Error
'    Debug.Print "&H" + Hex$(hWnd); ":",
    Dim retV As Long ' ��������ֵ����
    
    ' ȡ�ø��������������Ǵ򿪡�����Ի���������������ʱ��hWnd �ǶԻ�����������
    Dim hWndParent As Long: hWndParent = GetParent(hWnd)

    ' �ж���Ϣ������Ƿ�Ϊ�账�����Ϣ
    Select Case uMsg
        Case WM_INITDIALOG ' �Ի����ʼ��ʱ��
            Debug.Print "WM_INITDIALOG", "&H" + Hex(wParam), "&H" + Hex(lParam)
            ' ˽�б�����ʼ����
            procOld = 0: hWndTextView = 0: strSelFile = ""
            hWndButtonPlay(0) = 0: hWndButtonPlay(1) = 0: hWndButtonPlay(2) = 0
            ' ��ʾ�Ի���֮ǰ���Զ�������Ի�����ۡ�
            CDHook hWndParent ' �ص�������������ȡ��Ϣ���ö�̬�����Ŀؼ�������Ӧ��Ϣ��
            If WhichStyle = ssFont Then CustomizeFontDialog hWnd ' ��ʼ������Ի���
            If WhichStyle = ssColor Then setDlgStartUpPosition hWnd, hWndParent ' ��ʼ����ɫ�Ի���ֻ�������λ�ã�
            ' �ж���û����������ͼƬ�򣿣�����
            ' ������û������Ԥ��������־ͼƬ��ʱ���Ի���λ���޷����������⣻
            If m_picLogoPicture Is Nothing Then m_ppLogoPosition = ppNone
            If m_picPreviewPicture Is Nothing Then m_ppPreviewPosition = ppNone
            
        Case WM_NOTIFY ' �Ի���仯ʱ�����Դ�/����Ի��򣡣���
            retV = CDNotify(hWndParent, lParam)
        Case WM_COMMAND ' ������ ���塢��ɫ�Ի��� �ϵĿؼ�����
            'Debug.Print LOWORD(wParam); HIWORD(wParam)
            Dim L As Long: L = LOWORD(wParam)
            If WhichStyle = ssFont Then
                If L = enumFONT_CTL.btn_Apply _
                    Or L = enumFONT_CTL.cbo_FontName Or L = enumFONT_CTL.cbo_BoldItalic _
                    Or L = enumFONT_CTL.cbo_Size Or L = enumFONT_CTL.btn_Strikethrough _
                    Or L = enumFONT_CTL.btn_Underline Or L = enumFONT_CTL.cbo_Color _
                    Or L = enumFONT_CTL.cbo_Charset Then ' lParam �ؼ������ wParam ����=�ؼ� ID ������
                    ' ��������Ի���Ԥ��������Щ����������Ҫ˫����ǰ3��cbo����
                    mSetFontPreview hWnd
                ElseIf L = enumFONT_CTL.btn_Help Then ' ����Ի��������
                    MsgBox "����Ի��������", vbInformation
                'Else ' ����������������Ϣ���� Ӧ�� ��ť�����ǲ��У���
                '    SendMessage GetDlgItem(hWnd, enumFONT_CTL.btn_Apply), WM_LBUTTONDOWN, 0&, 0&
                End If
            ElseIf WhichStyle = ssColor Then
                If L = enumFONT_CTL.btn_Help Then ' ��ɫ�Ի��������������ťͨ��һ��IDֵ����
                    MsgBox "��ɫ�Ի��������", vbInformation
                End If
            End If
        Case WM_DESTROY ' �Ի�������ʱ��
            Debug.Print "WM_DESTROY", "&H" + Hex(wParam), "&H" + Hex(lParam)
            ' ȡ�� �Ƿ�ѡ��ֻ����ѡ��
            Dim hWndButton As Long: hWndButton = GetDlgItem(hWndParent, ID_ReadOnly)
            IsReadOnlyChecked = SendMessage(hWndButton, BM_GETCHECK, ByVal 0&, ByVal 0&)
            
            ' ����ͼƬ��ԭ���ĸ����ڣ���������=0���ָ�����ԭ���ĸ����ٴε����Ի���ʱͼƬ��ʧ�ˣ�������
            ' ���������������ͼƬ���ٴε���ʱ����������ʾ����������
            If Not m_picLogoPicture Is Nothing Then ' �������жϣ�����������
                ShowWindow m_picLogoPicture.hWnd, SW_HIDE
                Call SetParent(m_picLogoPicture.hWnd, Val(m_picLogoPicture.Tag))
            End If
            If Not m_picPreviewPicture Is Nothing Then
                ShowWindow m_picPreviewPicture.hWnd, SW_HIDE
                Call SetParent(m_picPreviewPicture.hWnd, Val(m_picPreviewPicture.Tag))
            End If
            
            ' ֹͣ���� ֹͣ��������Ϊǰ������ڲ��ţ�������' ע�⣺���ﲻֱ�ӵ��� B3Button_Click 2 ��
            PlayAudio strSelFile, 2
            sndPlaySound vbNullString, 0 ' ֹͣ Wave �ļ����š�
            
            ' ���ٴ����Ŀؼ�
            If hWndTextView Then DestroyWindow hWndTextView
            If hWndButtonPlay(0) Then DestroyWindow hWndButtonPlay(0) ': hWndButtonPlay(0) = 0
            If hWndButtonPlay(1) Then DestroyWindow hWndButtonPlay(1) ': hWndButtonPlay(1) = 0
            If hWndButtonPlay(2) Then DestroyWindow hWndButtonPlay(2) ': hWndButtonPlay(2) = 0
            If hWndFontPreview Then DestroyWindow hWndFontPreview  ' ����Ԥ���ı���
            ' �ص�������������ȡ��Ϣ���ö�̬�����Ŀؼ�������Ӧ��Ϣ��
            CDUnHook hWndParent
            ' �ͷ������ڴ棡������֪Ϊʲô�������Ի���󣬳���ռ�õ������ڴ����������
            SetProcessWorkingSetSize GetCurrentProcess(), -1&, -1&
'        Case Else
'            Debug.Print "Else ", "&H" + Hex(wParam), "&H" + Hex(lParam)
    End Select
    CDCallBackFun = retV ' ��������ֵ����
    
    On Error GoTo 0
    Exit Function

CDCallBack_Error:
    Debug.Print "CDCallBackFun Error " & Err.Number & " (" & Err.Description & ")"
    Resume Next
End Function
          
' �Ի���仯ʱ�����е��������Դ�/����Ի��򣡣���
Private Function CDNotify(ByVal hWndParent As Long, ByVal lParam As Long) As Long
    Dim hToolBar As Long    ' �Ի����Ϲ��������
    Dim rcTB As RECT        ' ����������
    Dim pt As POINTAPI, W As Long, H As Long
    Dim rcDlg As RECT       ' �Ի������
    Dim picLeft As Long, picTop As Long ' ͼƬ��λ�����꣬����ͼƬ���໥Ӱ�죬һ���ƶ�ʱҪ�ж���һ����λ�ã�
    ' == �м������б����Σ�ͼƬ�� Left Top λ�õĻ�׼�㡣����
    Dim hWndControl As Long, rcList0 As RECT, ptL As POINTAPI
    hWndControl = GetDlgItem(hWndParent, ID_List0) ' ����IDȡ�ÿؼ����
    GetWindowRect hWndControl, rcList0 ' ȡ�ÿؼ�����
    ptL.X = rcList0.Left: ptL.y = rcList0.Top
    ScreenToClient hWndParent, ptL ' ptL����ת������ܵõ���Ҫ�Ľ����
    
    Dim hdr     As NMHDR
    Call CopyMemory(hdr, ByVal lParam, LenB(hdr))
    Select Case hdr.code
        Case CDN_InitDone ' ��ʼ����ɣ��Ի���Ҫ��ʾʱ��
            Debug.Print "InitDone"
                        
            ' ===== �жϳ����־ͼƬ��λ�ã��Ե����Ի�����ۣ��ߴ缰���ϵĿؼ�λ�ã���
            Dim OffSetX As Long, OffSetY As Long, stpX As Single, stpY As Single  ' �Ի����С���ؼ�ƫ���������أ���
            stpX = Screen.TwipsPerPixelX: stpY = Screen.TwipsPerPixelY ' Twips ת��Ϊ Pixels Ҫ�������ǣ�
            If m_ppLogoPosition = ppNone Then GoTo NoLogo ' �ж���û����������ͼƬ�򣿣�����
            OffSetX = m_picLogoPicture.Width \ stpX: OffSetY = m_picLogoPicture.Height \ stpY
            Dim ClientRect As RECT ' ppBottom ʱ��ȡ�öԻ�����Σ�����������ͬ����֪Ϊʲôֻ���������У�����������
            Select Case m_ppLogoPosition
                Case ppNone ' �޳����־ͼƬ����������
                    OffSetX = 0: OffSetY = 0
                    picLeft = 0: picTop = 0
                Case ppLeft ' �����־ͼƬ ����ˣ�Ҫ�ƶ��Ի�����ԭ���Ŀؼ���
                    ' �Ի���������ԭʼ�ؼ�����
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
                    CDMoveOriginControl hWndParent, ID_FileTypeComboC, OffSetX  ' ����ۣ��ƶ��Ի������ļ����ı���
                    CDMoveOriginControl hWndParent, ID_FileNameText, OffSetX
                    
                    ' �ƶ������������������ر�ע�⣺�޷�ͨ�� CDMoveOriginControl �����ƶ�����
                    hToolBar = CDGetToolBarHandle(hWndParent)
                    GetWindowRect hToolBar, rcTB
                    pt.X = rcTB.Left
                    pt.y = rcTB.Top
                    ScreenToClient hWndParent, pt
                    MoveWindow hToolBar, pt.X + OffSetX, pt.y, rcTB.Right - rcTB.Left, rcTB.Bottom - rcTB.Top, True
                    
                    ' �ı�Ի����С����
                    GetWindowRect hWndParent, rcDlg ' ȡ�öԻ�����Σ����ƶ���ʵ��ֻ�ı��ȣ���
                    MoveWindow hWndParent, rcDlg.Left, rcDlg.Top, rcDlg.Right - rcDlg.Left + OffSetX, rcDlg.Bottom - rcDlg.Top, True
                    ' ���ó����־ͼƬ
                    ' �����µģ�������ͼƬ��ԭ���ĸ����ھ������
                    m_picLogoPicture.Tag = SetParent(m_picLogoPicture.hWnd, hWndParent)
                    ' �ƶ�ͼƬ��Top λ�ù̶����߶ȹ̶���
                    If m_ppPreviewPosition = ppLeft Then
                        picLeft = 2: picTop = 0
                    'ElseIf m_ppPreviewPosition = ppRight Then' ����Ҫ�жϣ�
                    ElseIf m_ppPreviewPosition = ppTop Then
                        picLeft = 2: picTop = m_picPreviewPicture.Height \ stpY
                    ElseIf m_ppPreviewPosition = ppBottom Then
                        picLeft = 2: picTop = m_picPreviewPicture.Height \ stpY
                    Else
                        picLeft = 2: picTop = 0
                    End If
                    MoveWindow m_picLogoPicture.hWnd, picLeft, 2, _
                        m_picLogoPicture.Width \ stpX, rcDlg.Bottom - rcDlg.Top + picTop - 29, True
                    ' ����ͼƬ
                    'm_picLogoPicture.PaintPicture m_picLogoPicture.Picture, 0, 0, m_picLogoPicture.ScaleWidth * 100, m_picLogoPicture.ScaleHeight
                    ShowWindow m_picLogoPicture.hWnd, SW_SHOW ' ��ʾͼƬ��

                Case ppRight
                    ' �ı�Ի����С����
                    GetWindowRect hWndParent, rcDlg ' ȡ�öԻ�����Σ����ƶ���ʵ��ֻ�ı��ȣ���
                    MoveWindow hWndParent, rcDlg.Left, rcDlg.Top, rcDlg.Right - rcDlg.Left + OffSetX, rcDlg.Bottom - rcDlg.Top, True
                    ' ���ó����־ͼƬ
                    ' �����µģ�������ͼƬ��ԭ���ĸ����ھ������
                    m_picLogoPicture.Tag = SetParent(m_picLogoPicture.hWnd, hWndParent)
                    ' �ƶ�ͼƬ�򣬣�Top λ�ù̶����߶ȹ̶���
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
                    ' ����ͼƬ
                    'm_picLogoPicture.PaintPicture m_picLogoPicture.Picture, 0, 0, m_picLogoPicture.ScaleWidth, m_picLogoPicture.ScaleHeight
                    ShowWindow m_picLogoPicture.hWnd, SW_SHOW ' ��ʾͼƬ��

                Case ppTop ' �����־ͼƬ �ڶ��ˣ�Ҫ�ƶ��Ի�����ԭ���Ŀؼ���
                    ' �Ի���������ԭʼ�ؼ�����
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
                    CDMoveOriginControl hWndParent, ID_FileTypeComboC, , OffSetY ' ����ۣ��ƶ��Ի������ļ����ı���
                    CDMoveOriginControl hWndParent, ID_FileNameText, , OffSetY
                    
                    ' �ƶ������������������ر�ע�⣺�޷�ͨ�� CDMoveOriginControl �����ƶ�����
                    hToolBar = CDGetToolBarHandle(hWndParent)
                    GetWindowRect hToolBar, rcTB
                    pt.X = rcTB.Left
                    pt.y = rcTB.Top
                    ScreenToClient hWndParent, pt
                    MoveWindow hToolBar, pt.X, pt.y + OffSetY, rcTB.Right - rcTB.Left, rcTB.Bottom - rcTB.Top, True
                    
                    ' �ı�Ի����С����
                    GetWindowRect hWndParent, rcDlg ' ȡ�öԻ�����Σ����ƶ���ʵ��ֻ�ı�߶ȣ���
                    MoveWindow hWndParent, rcDlg.Left, rcDlg.Top, rcDlg.Right - rcDlg.Left, rcDlg.Bottom - rcDlg.Top + OffSetY, True
                    ' ���ó����־ͼƬ
                    ' �����µģ�������ͼƬ��ԭ���ĸ����ھ������
                    m_picLogoPicture.Tag = SetParent(m_picLogoPicture.hWnd, hWndParent) ' GetParent(m_picLogoPicture.hwnd)
                    ' �ƶ�ͼƬ��Left λ�ù̶�����ȹ̶���picLeft + (rcDlg.Right - rcDlg.Left - m_picLogoPicture.Width \ stpX) \ 2 - 3
                    If m_ppPreviewPosition = ppLeft Then
                        picLeft = m_picPreviewPicture.Width \ stpX: picTop = 0
                    ElseIf m_ppPreviewPosition = ppRight Then ' ����Ҫ�жϣ�
                        picLeft = m_picPreviewPicture.Width \ stpX
                    'ElseIf m_ppPreviewPosition = ppTop Then
                    'ElseIf m_ppPreviewPosition = ppBottom Then
                    End If
                    MoveWindow m_picLogoPicture.hWnd, 5, picTop + 2, _
                        rcDlg.Right - rcDlg.Left + picLeft - 15, m_picLogoPicture.Height \ stpY, True
                    ' ����ͼƬ
                    'm_picLogoPicture.PaintPicture m_picLogoPicture.Picture, 0, 0, m_picLogoPicture.ScaleWidth, m_picLogoPicture.ScaleHeight
                    ShowWindow m_picLogoPicture.hWnd, SW_SHOW ' ��ʾͼƬ��

                Case ppBottom
                    ' �ı�Ի����С����
                    GetWindowRect hWndParent, rcDlg ' ȡ�öԻ�����Σ����ƶ���ʵ��ֻ�ı�߶ȣ���
                    MoveWindow hWndParent, rcDlg.Left, rcDlg.Top, rcDlg.Right - rcDlg.Left, rcDlg.Bottom - rcDlg.Top + OffSetY, True
                    ' ���ó����־ͼƬ
                    Call GetClientRect(hWndParent, ClientRect) ' �� rcDlg.Bottom ���У�����
                    ' �����µģ�������ͼƬ��ԭ���ĸ����ھ������
                    m_picLogoPicture.Tag = SetParent(m_picLogoPicture.hWnd, hWndParent)
                    ' �ƶ�ͼƬ��Left λ�ù̶�����ȹ̶���
                    If m_ppPreviewPosition = ppLeft Then
                        picLeft = m_picPreviewPicture.Width \ stpX: picTop = 0
                    ElseIf m_ppPreviewPosition = ppRight Then
                        picLeft = m_picPreviewPicture.Width \ stpX
                    ElseIf m_ppPreviewPosition = ppTop Then
                        picTop = m_picPreviewPicture.Height \ stpY: picLeft = 0
                    ElseIf m_ppPreviewPosition = ppBottom Then ' ��ʱ��Ҫ�ƶ���־��Ԥ�����棡
                        picTop = m_picPreviewPicture.Height \ stpY: picLeft = 0
                    End If
                    MoveWindow m_picLogoPicture.hWnd, 5, picTop + ClientRect.Bottom - OffSetY, _
                        rcDlg.Right - rcDlg.Left + picLeft - 15, m_picLogoPicture.Height \ stpY, True
                    ' ����ͼƬ
                    'm_picLogoPicture.PaintPicture m_picLogoPicture.Picture, 0, 0, m_picLogoPicture.ScaleWidth, m_picLogoPicture.ScaleHeight
                    ShowWindow m_picLogoPicture.hWnd, SW_SHOW ' ��ʾͼƬ��

            End Select
NoLogo:
' **********************************************************************************************************
            ' ===== �ж�Ԥ��ͼƬ��λ�ã��ر�ע�⣺Ҫ�жϳ����־ͼƬ��λ�ã���������������������������
            ' Ԥ��ͼƬ��λ�ù̶�һ��ֵ��������С�����ҡ����·�����������ֱ�̶��߶ȡ���ȣ�����
            If m_ppPreviewPosition = ppNone Then GoTo NoPreview ' �ж���û����������ͼƬ�򣿣�����
            OffSetX = m_picPreviewPicture.Width \ stpX: OffSetY = m_picPreviewPicture.Height \ stpY
            Select Case m_ppPreviewPosition
                Case ppNone
                    OffSetX = 0: OffSetY = 0
                    picLeft = 0: picTop = 0
                Case ppLeft ' Ԥ��ͼƬ�� ����ˣ�Ҫ�ƶ��Ի�����ԭ���Ŀؼ���
                    ' ����ȡ�� LIST�ؼ����Σ����������汻�ƶ��ˣ�
                    GetWindowRect hWndControl, rcList0
                    ptL.X = rcList0.Left: ptL.y = rcList0.Top
                    ScreenToClient hWndParent, ptL ' ptL����ת������ܵõ���Ҫ�Ľ����
                    
                    ' �Ի���������ԭʼ�ؼ�����
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
                    CDMoveOriginControl hWndParent, ID_FileTypeComboC, OffSetX  ' ����ۣ��ƶ��Ի������ļ����ı���
                    CDMoveOriginControl hWndParent, ID_FileNameText, OffSetX
                    
                    ' �ƶ������������������ر�ע�⣺�޷�ͨ�� CDMoveOriginControl �����ƶ�����
                    hToolBar = CDGetToolBarHandle(hWndParent)
                    GetWindowRect hToolBar, rcTB
                    pt.X = rcTB.Left
                    pt.y = rcTB.Top
                    ScreenToClient hWndParent, pt
                    MoveWindow hToolBar, pt.X + OffSetX, pt.y, rcTB.Right - rcTB.Left, rcTB.Bottom - rcTB.Top, True
                    
                    ' �ı�Ի����С����
                    GetWindowRect hWndParent, rcDlg ' ȡ�öԻ�����Σ����ƶ���ʵ��ֻ�ı��ȣ���
                    MoveWindow hWndParent, rcDlg.Left, rcDlg.Top, rcDlg.Right - rcDlg.Left + OffSetX, rcDlg.Bottom - rcDlg.Top, True
                    ' ���� Ԥ��ͼƬ��
                    ' �����µģ�������ͼƬ��ԭ���ĸ����ھ������
                    m_picPreviewPicture.Tag = SetParent(m_picPreviewPicture.hWnd, hWndParent)
                    ' �ƶ�ͼƬ��Top λ�ù̶����߶ȹ̶���
                    picLeft = 5: picTop = ptL.y: W = 5
                    If m_ppLogoPosition = ppLeft Then
                        picLeft = m_picLogoPicture.Width \ stpX + 5
                    'ElseIf m_ppLogoPosition = ppRight Then' ����Ҫ�жϣ�
                    'ElseIf m_ppLogoPosition = ppTop Then
                    'ElseIf m_ppLogoPosition = ppBottom Then
                    End If
                    MoveWindow m_picPreviewPicture.hWnd, picLeft, picTop, _
                        m_picPreviewPicture.Width \ stpX - W, rcList0.Bottom - rcList0.Top, True
                    ' ����ͼƬ
                    'm_picPreviewPicture.PaintPicture m_picPreviewPicture.Picture, 0, 0, m_picPreviewPicture.ScaleWidth, m_picPreviewPicture.ScaleHeight
                    myPaintPicture m_picPreviewPicture, False
                    ShowWindow m_picPreviewPicture.hWnd, SW_SHOW ' ��ʾͼƬ��

                Case ppRight
                    ' ����ȡ�� LIST�ؼ����Σ����������汻�ƶ��ˣ�
                    GetWindowRect hWndControl, rcList0
                    ptL.X = rcList0.Left: ptL.y = rcList0.Top
                    ScreenToClient hWndParent, ptL ' ptL����ת������ܵõ���Ҫ�Ľ����
                    
                    ' �ı�Ի����С����
                    GetWindowRect hWndParent, rcDlg ' ȡ�öԻ�����Σ����ƶ���ʵ��ֻ�ı��ȣ���
                    MoveWindow hWndParent, rcDlg.Left, rcDlg.Top, rcDlg.Right - rcDlg.Left + OffSetX, rcDlg.Bottom - rcDlg.Top, True
                    ' ���� Ԥ��ͼƬ��
                    ' �����µģ�������ͼƬ��ԭ���ĸ����ھ������
                    m_picPreviewPicture.Tag = SetParent(m_picPreviewPicture.hWnd, hWndParent)
                    ' �ƶ�ͼƬ��Right λ�ù̶����߶ȹ̶���
                    If m_ppLogoPosition = ppRight Then
                        picLeft = rcDlg.Right - rcDlg.Left - 5 - m_picLogoPicture.Width \ stpX: picTop = 0
                    Else
                        picLeft = rcDlg.Right - rcDlg.Left - 8: picTop = 0
                    End If
                    MoveWindow m_picPreviewPicture.hWnd, picLeft, ptL.y, _
                        m_picPreviewPicture.Width \ stpX - 3, rcList0.Bottom - rcList0.Top, True
                    ' ����ͼƬ
                    'm_picPreviewPicture.PaintPicture m_picPreviewPicture.Picture, 0, 0, m_picPreviewPicture.ScaleWidth, m_picPreviewPicture.ScaleHeight
                    myPaintPicture m_picPreviewPicture, False
                    ShowWindow m_picPreviewPicture.hWnd, SW_SHOW ' ��ʾͼƬ��

                Case ppTop ' Ԥ��ͼƬ�� �ڶ��ˣ�Ҫ�ƶ��Ի�����ԭ���Ŀؼ���
                    ' ����ȡ�� LIST�ؼ����Σ����������汻�ƶ��ˣ�
                    GetWindowRect hWndControl, rcList0
                    ptL.X = rcList0.Left: ptL.y = rcList0.Top
                    ScreenToClient hWndParent, ptL ' ptL����ת������ܵõ���Ҫ�Ľ����
                    
                    ' �Ի���������ԭʼ�ؼ�����
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
                    CDMoveOriginControl hWndParent, ID_FileTypeComboC, , OffSetY ' ����ۣ��ƶ��Ի������ļ����ı���
                    CDMoveOriginControl hWndParent, ID_FileNameText, , OffSetY
                    
                    ' �ƶ������������������ر�ע�⣺�޷�ͨ�� CDMoveOriginControl �����ƶ�����
                    hToolBar = CDGetToolBarHandle(hWndParent)
                    GetWindowRect hToolBar, rcTB
                    pt.X = rcTB.Left
                    pt.y = rcTB.Top
                    ScreenToClient hWndParent, pt
                    MoveWindow hToolBar, pt.X, pt.y + OffSetY, rcTB.Right - rcTB.Left, rcTB.Bottom - rcTB.Top, True

                    ' �ı�Ի����С����
                    GetWindowRect hWndParent, rcDlg ' ȡ�öԻ�����Σ����ƶ���ʵ��ֻ�ı�߶ȣ���
                    MoveWindow hWndParent, rcDlg.Left, rcDlg.Top, rcDlg.Right - rcDlg.Left, rcDlg.Bottom - rcDlg.Top + OffSetY, True
                    ' ���ó����־ͼƬ
                    ' �����µģ�������ͼƬ��ԭ���ĸ����ھ������
                    m_picPreviewPicture.Tag = SetParent(m_picPreviewPicture.hWnd, hWndParent)
                    ' �ƶ�ͼƬ��Left λ�ù̶�����ȹ̶���
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
                    ' ����ͼƬ
                    'm_picPreviewPicture.PaintPicture m_picPreviewPicture.Picture, 0, 0, m_picPreviewPicture.ScaleWidth, m_picPreviewPicture.ScaleHeight
                    myPaintPicture m_picPreviewPicture, False
                    ShowWindow m_picPreviewPicture.hWnd, SW_SHOW ' ��ʾͼƬ��
                    
                Case ppBottom
                    ' ����ȡ�� LIST�ؼ����Σ����������汻�ƶ��ˣ�
                    GetWindowRect hWndControl, rcList0
                    ptL.X = rcList0.Left: ptL.y = rcList0.Top
                    ScreenToClient hWndParent, ptL ' ptL����ת������ܵõ���Ҫ�Ľ����
                    
                    ' �ı�Ի����С����
                    GetWindowRect hWndParent, rcDlg ' ȡ�öԻ�����Σ����ƶ���ʵ��ֻ�ı�߶ȣ���
                    MoveWindow hWndParent, rcDlg.Left, rcDlg.Top, rcDlg.Right - rcDlg.Left, rcDlg.Bottom - rcDlg.Top + OffSetY, True
                    ' ���� Ԥ��ͼƬ��
                    Call GetClientRect(hWndParent, ClientRect) ' �� rcDlg.Bottom ���У�����
                    ' �����µģ�������ͼƬ��ԭ���ĸ����ھ������
                    m_picPreviewPicture.Tag = SetParent(m_picPreviewPicture.hWnd, hWndParent)
                    ' �ƶ�ͼƬ��Left λ�ù̶�����ȹ̶���
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
                    ' ����ͼƬ
                    'm_picPreviewPicture.PaintPicture m_picPreviewPicture.Picture, 0, 0, m_picPreviewPicture.ScaleWidth, m_picPreviewPicture.ScaleHeight
                    myPaintPicture m_picPreviewPicture, False
                    ShowWindow m_picPreviewPicture.hWnd, SW_SHOW ' ��ʾͼƬ��

            End Select
NoPreview:
            ' ���öԻ�������λ�ã�ֻ�ж���Ļ���ĺ����������ģ��������ܣ�
            GetWindowRect hWndParent, rcDlg ' ȡ�öԻ������
            W = rcDlg.Right - rcDlg.Left: H = rcDlg.Bottom - rcDlg.Top ' �Ի������
            If m_dlgStartUpPosition = vbStartUpScreen Then ' ���ƶ�����Ļ����
                MoveWindow hWndParent, (Screen.Width \ stpX - W) \ 2, (Screen.Height \ stpY - H) \ 2, W, H, True
            ElseIf m_dlgStartUpPosition = vbStartUpOwner Then ' ����������
                Dim rcOwner As RECT, T As Long ' ȡ�öԻ���ĸ����ھ���
                GetWindowRect GetParent(hWndParent), rcOwner
                T = rcOwner.Top + (rcOwner.Bottom - rcOwner.Top - H) \ 2
                If T < 0 Then T = 0 ' ��֤�Ի��򲻳�����Ļ���ˣ�
                MoveWindow hWndParent, rcOwner.Left + (rcOwner.Right - rcOwner.Left - W) \ 2, T, W, H, True
            End If
            ' ����λ�ã���Ϊ���ֶԻ�����������ԣ�����һ������ʵ�֣�Ч�����ã������ˣ�
            'setDlgStartUpPosition hWndParent, GetParent(hWndParent)
            
            ' ����Ԥ���ı������С��λ����Ԥ��ͼƬ��һ��
            Dim rcP As RECT ' ����Ԥ���ı���λ�úʹ�С��
            GetWindowRect m_picPreviewPicture.hWnd, rcP ' û���� m_picPreviewPicture �������(��������� With �����δ����)
            ptL.X = rcP.Left: ptL.y = rcP.Top
            ScreenToClient hWndParent, ptL ' ptL����ת������ܵõ���Ҫ�Ľ����
            ' ��ʼ����Ԥ���ı��� ȥ�� Or WS_VISIBLE �����ڴ���ʱ��ʾ��
            hWndTextView = CreateWindowEx(0, _
                "Edit", App.LegalCopyright, _
                WS_BORDER Or WS_CHILD Or WS_HSCROLL Or WS_VSCROLL Or ES_AUTOHSCROLL Or ES_AUTOHSCROLL Or ES_MULTILINE, _
                ptL.X, ptL.y, _
                rcP.Right - rcP.Left, rcP.Bottom - rcP.Top, _
                hWndParent, 0&, App.hInstance, 0&)
            ' �����µ����� Fixedsys - Times New Roman - MS Sans Serif
            Dim NewFont As Long
            NewFont = CreateFont(18, 0, 0, 0, _
                      366, False, False, False, _
                      DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_LH_ANGLES, _
                      ANTIALIASED_QUALITY Or PROOF_QUALITY, DEFAULT_PITCH Or FF_DONTCARE, _
                       "Times New Roman")
            ' �����ı�������
            SendMessage hWndTextView, WM_SETFONT, NewFont, 0
            ' ���ء���ʾ�Ի����ϵĿؼ� m_blnHideControls(I) ��ֵ�����ǲ���Ҫ���أ�
            Call HideOrShowDlgControls(hWndParent)
            ' ���öԻ����ϵĿؼ������֣�m_strControlsCaption(I) ������ֵ��
            Call mSetDlgControlsCaption(hWndParent)
            
        Case CDN_SelChange ' �ļ�ѡ��ı�ʱ������Ԥ����
            strSelFile = SendMsgGetStr(hdr.hwndFrom, CDM_GetFilePath) ' ��¼ѡ�е��ļ�·��
            Debug.Print "SelChange:"; strSelFile
            'Screen.MousePointer = vbHourglass ' ����ɳ©״
            If Not m_ppPreviewPosition = ppNone Then LoadPreview strSelFile, hWndParent  ' ���ú���������Ԥ����
            'Screen.MousePointer = vbDefault ' ���Ԥ�������ָ�
        Case CDN_FolderChange
            Debug.Print "FolderChange:"; SendMsgGetStr(hdr.hwndFrom, CDM_GetFolderPath)
        Case CDN_ShareViolation
            Debug.Print "ShareViolation"
        Case CDN_Help
            Debug.Print "Help"
            If WhichStyle = ssOpen Then
                MsgBox "�������򿪶Ի��� ��", vbInformation
            ElseIf WhichStyle = ssSave Then
                MsgBox "����������Ի��� ��", vbInformation
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
    ' ����' �����־ͼƬ���Ƿ������ﲻ�ɣ���������Ч�����ԣ�����
    myPaintPicture m_picLogoPicture
End Function

' �ƶ� ��/���� �Ի�����ԭ�еĿؼ�' ��Ҫ�ƶ��Ի�����ԭ��û�еĿؼ��������⴦��
' ����Ŀ�ѡ��������ΪĬ�� -1 ʱ���ƶ�ʱ��������
Private Sub CDMoveOriginControl(ByVal hWndCD As Long, ByVal ID As Long, _
    Optional ByVal X As Long = -1, Optional ByVal y As Long = -1, _
    Optional ByVal nWidth As Long = -1, Optional ByVal nHeight As Long = -1)
    
    Dim hWndControl As Long ' �ؼ����
    Dim rectControl As RECT ' �ؼ�����
    Dim ptCtr As POINTAPI   ' �ؼ����Ͻ�����Ļ������ֵ
    Dim ptDlg As POINTAPI   ' �Ի������Ͻ�����Ļ������ֵ
    
    ' == ����IDȡ�ÿؼ����
    hWndControl = GetDlgItem(hWndCD, ID)
    
    ' == ȡ�ÿؼ����Σ�λ������ԶԻ�����ԣ����ԶԻ������Ͻ��ǵ�Ϊ0�㣩�����ÿؼ���С
    GetWindowRect hWndControl, rectControl

    ' ==ȡ�öԻ���λ��
    ScreenToClient hWndCD, ptDlg
    ' ==ȡ�ò����ÿؼ�λ��
    ScreenToClient hWndControl, ptCtr
    ptCtr.X = rectControl.Left + ptDlg.X + IIf(X <> -1, X, 0)
    ptCtr.y = rectControl.Top + ptDlg.y + IIf(y <> -1, y, 0)
    X = ptCtr.X
    y = ptCtr.y
    nWidth = rectControl.Right - rectControl.Left + IIf(nWidth <> -1, nWidth, 0)
    nHeight = rectControl.Bottom - rectControl.Top + IIf(nHeight <> -1, nHeight, 0)
    
    ' ����API�ƶ��ؼ���X��YΪ�����Ļ���Ͻǵ�����ֵ��
    MoveWindow hWndControl, X, y, nWidth, nHeight, True
End Sub

' �ƶ��Ի����Ϲ�����ʱ��ȡ�ù����������
Private Function CDGetToolBarHandle(ByVal hDialog As Long) As Long
    CDGetToolBarHandle = FindWindowEx(hDialog, 0, "ToolBarWindow32", vbNullString)
End Function

' �Ի���ı�Ŀ¼��ѡ���ļ�ʱ��ȡ�� ·����
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
' ### ����ΪԤ����ص���Ҫ���� ##################################################
' Ԥ��ͼƬ��ͼƬ����ʾ���� PaintPicture ������ͼƬ���ϻ�ͼƬ��
Private Sub myPaintPicture(picBox As PictureBox, Optional blnShowPictureOrNone As Boolean = True, Optional stdPic As StdPicture = Nothing)
    On Error Resume Next
    Dim tempPic As StdPicture
    If stdPic Is Nothing Then
        Set tempPic = picBox.Picture
    Else
        Set tempPic = stdPic
    End If
    picBox.AutoRedraw = True ' ��ͼƬ���Զ�ˢ�£�
    If blnShowPictureOrNone Then
        picBox.PaintPicture tempPic, 0, 0, picBox.ScaleWidth, picBox.ScaleHeight
    Else ' ����ͼƬ��
        Set picBox.Picture = Nothing
        ' ��ͼƬ��������ʾ���֣��б�Ҫ���� =====================================
        Dim s As String: s = App.LegalCopyright
        picBox.ForeColor = vbBlack: picBox.FontSize = 18
        picBox.CurrentX = (picBox.ScaleWidth - picBox.TextWidth(s)) / 2
        picBox.CurrentY = (picBox.ScaleHeight - picBox.TextHeight(s)) / 2
        picBox.Print s ' =======================================================
    End If
End Sub

' �ж��ļ����ͣ��Ծ�����ʲô��ʽԤ��������
Private Function getFileType(strFileName As String) As FileType
    Dim strExt As String    ' �ļ���׺�� �� TXT
    ' ȡ���ļ���׺������ת��Ϊ��д��
    strExt = UCase$(Right$(strFileName, Len(strFileName) - InStrRev(strFileName, ".")))
    getFileType = ffText ' ���ú�������Ĭ�����ͣ�
    ' �ж��ļ���׺����VB ͼƬ���ܼ��� PNG ͼƬ ANI ������꣡JPE\JFIF\TIF\TIFF δ֪��
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

' VB ȡ��ͼƬ��С
Private Function fncGetPicInfo(lsPicName As String, Optional hBitmap As Long = 0) As PicInfo
    Dim res As Long
    Dim bmp As BITMAP
    If hBitmap = 0 Then hBitmap = LoadPicture(lsPicName).Handle ' ����Ҫ�ڴ����У�LoadPicture ����Ч��
    res = GetObject(hBitmap, Len(bmp), bmp) 'ȡ��BITMAP�Ľṹ
    fncGetPicInfo.picWidth = bmp.bmWidth
    fncGetPicInfo.picHeight = bmp.bmHeight
End Function

' ����Ԥ����ͼƬ���ı��ļ���Wave �ļ�����Ƶ�ļ���
Private Sub LoadPreview(strFileName As String, ByVal hWndParent As Long, Optional ShowFileSize As Long = &H1000)
    If FileLen(strFileName) = 0 Then Exit Sub
    On Error GoTo ErrLoad
    Dim I As Integer
    ' �����ļ����ͣ����в�ͬ��Ԥ��������
    If getFileType(strFileName) = ffText Then ' �ı��ļ�
        Dim FileNum      As Integer
        Dim FileSize     As Long
        Dim LoadBytes()  As Byte
        ' ��ȡ�ı��ļ�
        FileNum = FreeFile
        Open strFileName For Binary Access Read Lock Write As FileNum
            FileSize = LOF(FileNum)
            If FileSize > ShowFileSize Then FileSize = ShowFileSize ' Ԥ���ļ��Ĵ�С�����ܲ���ʾ�������ݡ�
            ReDim LoadBytes(0 To FileSize - 1)
            Get #FileNum, 1, LoadBytes
        Close FileNum
        ' ���ı�����ʾ����
        Debug.Print SetTextColor(GetDC(hWndTextView), vbRed)
        Call SetWindowText(hWndTextView, LoadBytes(0))  ' �������֣��ļ���������֣�
        Dim s As String: s = String(FileSize, 0)
        GetWindowText hWndTextView, s, FileSize
        Call SetWindowText(hWndTextView, ByVal Replace$(s, Chr$(0), "") _
            & vbCrLf & "������������ ʡ�� " & Format$(FileLen(strFileName) - FileSize, "###,###,###,##0") & " �ֽڣ�") ' �������֣�+���������� ��־��
        ShowWindow hWndTextView, SW_SHOW                ' �ı���ɼ�
        ShowWindow m_picPreviewPicture.hWnd, SW_HIDE    ' ����Ԥ��ͼƬ��
        For I = 0 To 2
            ShowWindow hWndButtonPlay(I), SW_HIDE
        Next I
        Exit Sub
    ElseIf getFileType(strFileName) = ffPicture Then ' ͼƬ�ļ���ͼƬԤ����������ʾ������ʾͼƬ��x�ߺ�Ԥ����������
        Dim W As Long, H As Long, fWH As PicInfo ' ���ص�ͼƬ��x��
        Dim W0 As Long, H0 As Long, oldSM As ScaleModeConstants ' ͼƬ���С��
        Dim X As Long, y As Long, W1 As Long, H1 As Long, per As Integer, sP As StdPicture ' ������������ʾ��ͼƬλ�ã���С��
        oldSM = m_picPreviewPicture.ScaleMode: m_picPreviewPicture.ScaleMode = vbPixels ' ����ͼƬ��ScaleModeΪ���أ�
        W0 = m_picPreviewPicture.ScaleWidth: H0 = m_picPreviewPicture.ScaleHeight
        ' ����ͼƬ��ͼƬ��ֻΪȡ����ߴ磡��
        Set sP = LoadPicture(strFileName)
        Set m_picPreviewPicture.Picture = sP ' LoadPicture(strFileName)
        fWH = fncGetPicInfo(strFileName, m_picPreviewPicture.Picture)
        W = fWH.picWidth: H = fWH.picHeight
        ' �ж�ͼƬ���߱ȡ����ص�ͼƬ��߱ȣ����ֵ����ȡ�����
        If (W0 / H0) / (W / H) < 1 Then ' ��' ͼƬ����Ϊ��׼
            W1 = W0: H1 = W0 / W * H
            X = 0: y = (H0 - H1) / 2    ' ������ʾλ�ã�ʹ����У�
            per = W0 / W * 100
        Else                            ' ��' ͼƬ��߶�Ϊ��׼
            W1 = H0 / H * W: H1 = H0
            X = (W0 - W1) / 2: y = 0    ' ������ʾλ�ã�ʹ����У�
            per = H0 / H * 100
        End If
        'Debug.Print "�ߴ�: " & W & " x " & H & " ͼƬ��: " & W0 & " x " & H0
        Set m_picPreviewPicture.Picture = Nothing
        '' �ٴμ���ͼƬ��ͼƬ������ʾ���������ټ��أ��������ʾ����ͼ��PaintPicture ����Ҳ����ʾͼ����������
        'Set m_picPreviewPicture.Picture = sP ' LoadPicture(strFileName)
        'myPaintPicture m_picPreviewPicture
        m_picPreviewPicture.PaintPicture sP, X, y, W1, H1
        ' ͼƬ������ʾ����
        Dim sT As String: sT = "�ߴ�: " & W & " x " & H & " (" & per & "%)"
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
        m_picPreviewPicture.ScaleMode = oldSM ' ��ԭͼƬ��ɵ� ScaleMode ֵ��
        Exit Sub
    ElseIf getFileType(strFileName) = ffWave Then ' �����ļ�
        Create3Buttons hWndParent ' ���� 3 �����ſ��ư�ť��
        ShowWindow m_picPreviewPicture.hWnd, SW_SHOW ' ��ʾԤ��ͼƬ��
        Set m_picPreviewPicture.Picture = Nothing
        ShowWindow hWndTextView, SW_HIDE             ' �����ı���
        
        ' ��ֹͣ��������Ϊǰ����ܲ��Ź���������' ע�⣺���ﲻֱ�ӵ��� B3Button_Click 2 ��
        PlayAudio strSelFile, 2
        sndPlaySound vbNullString, 0 ' ֹͣ Wave �ļ����š�
        ' =========================================================================================
        ' ==== ����Wave�ļ����Σ���ȥ����==========================================================
        ' =========================================================================================
        ' �������Σ���һ��ģ����ɣ�����ȥ���˹��ܣ���������������
        MDrawWaves.DrawWaves strFileName, m_picPreviewPicture  ' Ϊɶ���У�ԭ��һ��Ҫ���� ScaleMode = vbTwips ������
        ' =========================================================================================
        ' ==== ����Wave�ļ����Σ���ȥ����==========================================================
        ' =========================================================================================
        Exit Sub
    ElseIf getFileType(strFileName) = ffAudio Then ' ��Ƶ�ļ�
        Create3Buttons hWndParent ' ���� 3 �����ſ��ư�ť��
        ShowWindow m_picPreviewPicture.hWnd, SW_SHOW ' ��ʾԤ��ͼƬ��
        Set m_picPreviewPicture.Picture = Nothing
        ShowWindow hWndTextView, SW_HIDE             ' �����ı���
        
        ' ��ֹͣ��������Ϊǰ����ܲ��Ź���������' ע�⣺���ﲻֱ�ӵ��� B3Button_Click 2 ��
        PlayAudio strSelFile, 2
        sndPlaySound vbNullString, 0 ' ֹͣ Wave �ļ����š�
        Exit Sub
    End If

ErrLoad:
    Debug.Print "LoadPreview Error " & Err.Number & ": " & Err.Description
End Sub
' �������ſ��� 3 ����ť
Private Sub Create3Buttons(ByVal hWndParent As Long)
    
    ' �����ظ�������������Ӧ��Ϣʱ��hWndButtonPlay(0)������ֵ���ˣ��� lParam ��ֵ���ǵ�һ�δ���ʱ�ģ������޷���Ӧ��Ϣ��
    If hWndButtonPlay(0) Then ' Ҫ��ʾ��ť�����򲻼��ˣ���
        ShowWindow hWndButtonPlay(0), SW_SHOW
        ShowWindow hWndButtonPlay(1), SW_SHOW
        ShowWindow hWndButtonPlay(2), SW_SHOW
        Exit Sub
    End If
    'MsgBox "�������ſ��� 3 ����ť"
    ' �������ſ��� 3 ����ť����λ�����ļ����ͱ�ǩλ�þ�������С�̶���
    Dim rcP As RECT, ptL As POINTAPI ' �������ſ��� 3 ����ťλ�úʹ�С��
    Dim NewFont As Long
    GetWindowRect GetDlgItem(hWndParent, ID_FileTypeLabel), rcP
    ptL.X = rcP.Left: ptL.y = rcP.Top
    ScreenToClient hWndParent, ptL ' ptL����ת������ܵõ���Ҫ�Ľ����
    ' ��ʼ���� ���ſ��� 3 ����ť �� Or WS_VISIBLE ���ڴ���ʱ��ʾ��
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
    ' �����µ����� Webdings - Fixedsys - Times New Roman - MS Sans Serif
    NewFont = CreateFont(18, 0, 0, 0, _
              366, False, False, False, _
              DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_LH_ANGLES, _
              ANTIALIASED_QUALITY Or PROOF_QUALITY, DEFAULT_PITCH Or FF_DONTCARE, _
               "Webdings")
    ' ��������
    SendMessage hWndButtonPlay(0), WM_SETFONT, NewFont, 0
    SendMessage hWndButtonPlay(1), WM_SETFONT, NewFont, 0
    SendMessage hWndButtonPlay(2), WM_SETFONT, NewFont, 0
End Sub

' ע�⣺������ɫ�Ի��򡣡������öԻ�������λ�ã�ֻ�ж���Ļ���ĺ����������ģ��������ܣ�
Private Sub setDlgStartUpPosition(ByVal hWndDlg As Long, ByVal hWndParent As Long)
    Dim W As Long, H As Long, T As Long
    Dim rcDlg As RECT, rcOwner As RECT
    GetWindowRect hWndDlg, rcDlg ' ȡ�öԻ������
    W = rcDlg.Right - rcDlg.Left: H = rcDlg.Bottom - rcDlg.Top ' �Ի������
    If m_dlgStartUpPosition = vbStartUpScreen Then ' ���ƶ�����Ļ����
        MoveWindow hWndDlg, (Screen.Width \ Screen.TwipsPerPixelX - W) \ 2, (Screen.Height \ Screen.TwipsPerPixelY - H) \ 2, W, H, True
    ElseIf m_dlgStartUpPosition = vbStartUpOwner Then ' ����������
         ' ȡ�öԻ���ĸ����ھ���
        GetWindowRect hWndParent, rcOwner
        T = rcOwner.Top + (rcOwner.Bottom - rcOwner.Top - H) \ 2
        If T < 0 Then T = 0 ' ��֤�Ի��򲻳�����Ļ���ˣ�
        MoveWindow hWndDlg, rcOwner.Left + (rcOwner.Right - rcOwner.Left - W) \ 2, T, W, H, True
    End If
End Sub


' =========================================================================================
' ==== �Զ���Ի����ϵĿؼ����ء���ʾ���ı����ֵȣ�����ȥ����==============================
' =========================================================================================
' ���ء���ʾ�Ի����ϵĿؼ� m_blnHideControls(I) ��ֵ�����ǲ���Ҫ���أ����ң�Flags ����Ҳ��Ӱ�졣
Private Sub HideOrShowDlgControls(ByVal hWndDlg As Long)
    Dim I As Integer, ctrl_ID As Variant ' �ؼ�ID��ת��һ�����飬�ò�����
    ' (0 to 12) �ֱ�Ϊ����Щ�������ء��á��������ֻʣ�� (0 to 8)
    ' �����ҷ�Χ(&I)����ǩ      -- Ŀ¼������                         -- ����������        1
    ' �����Ŀ¼�����汾>=Win2K���� -- ���б���г��ļ���������򣩡�                    2
    ' ���ļ���(&N)����ǩ        -- �����ļ���(&N)���ı���           -- ��ȷ��(&O)������  1
    ' ���ļ�����(&T)����ǩ      -- ���ļ�����(&T)�������������ʱ�� -- ��ȡ��(&C)������  0
    ' ��ֻ������ѡ��            -- ������(&H)������                                        0
    ctrl_ID = Array(ID_FolderLabel, ID_FolderCombo, _
                ID_FileNameLable, ID_OK, _
                ID_FileTypeLabel, ID_FileTypeCombo0, ID_Cancel, _
                ID_ReadOnly, ID_Help)
    ' ���ؿؼ�
    For I = 0 To 8
        If m_blnHideControls(I) Then Call SendMessage(hWndDlg, CDM_HideControl, ctrl_ID(I), ByVal 0&)
    Next I
    Set ctrl_ID = Nothing
End Sub
' ���öԻ����ϵĿؼ������֡�m_strControlsCaption(I) ������ֵ��Ĭ��Ϊ�գ����ı�ԭʼֵ��
Private Sub mSetDlgControlsCaption(ByVal hWndDlg As Long)
    Dim I As Integer, ctrl_ID As Variant ' �ؼ�ID��ת��һ�����飬�ò�����
    ' (0 to 12) �ֱ�Ϊ����Щ�����������֡��á��������ֻʣ�� (0 to 6)
    ' �����ҷ�Χ(&I)����ǩ      -- ��Ŀ¼������                        -- ����������        2
    ' �����Ŀ¼�����汾>=Win2K���� -- ���б���г��ļ���������򣩡�                       2
    ' ���ļ���(&N)����ǩ        -- �����ļ���(&N)���ı���              -- ��ȷ��(&O)������  1
    ' ���ļ�����(&T)����ǩ      -- �����ļ�����(&T)�������������ʱ����-- ��ȡ��(&C)������  1
    ' ��ֻ������ѡ��            -- ������(&H)������                                           0
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
' ==== �Զ���Ի����ϵĿؼ����ء���ʾ���ı����ֵȣ�����ȥ����==============================
' =========================================================================================



' =========================================================================================
' ==== �����ļ��Ĳ��ţ���ȥ����============================================================
' =========================================================================================
' �������Ű�ť�����¼���
Private Sub B3Button_Click(Index As Integer)
'    PlayAudio strSelFile, Index' ���� mciSendString �������ţ�������
    If Index = 0 Then ' ����
        If getFileType(strSelFile) = ffWave Then ' Wave �ļ�
            sndPlaySound strSelFile, SND_FILENAME Or SND_ASYNC
        Else
            PlayAudio strSelFile, 0
        End If
    ElseIf Index = 1 Then ' ��ͣ
        If getFileType(strSelFile) = ffWave Then ' Wave �ļ�
            sndPlaySound vbNullString, 0 ' ��֪����ô��ͣ������ֹͣ������
        Else
            PlayAudio strSelFile, 1
        End If
    Else ' ֹͣ
        If getFileType(strSelFile) = ffWave Then ' Wave �ļ�
            sndPlaySound vbNullString, 0
        Else
            PlayAudio strSelFile, 2
        End If
    End If
End Sub
' �߼�ý�岥�ţ�������Ƶ�ļ�����Ӧ̫���ˣ�������Wave �ļ������������������ţ�����������
Private Sub PlayAudio(strFileName As String, Optional setStatus As Integer = 0)
    If Len(Dir$(strFileName)) = 0 Then Exit Sub
    Const ALIAS_NAME As String = "mySound"
    Dim rt As Long
    If setStatus = 0 Then ' ����
        If getPalyStatus(ALIAS_NAME) = IsStopped Then
            ' �򿪲���ͷ��ʼ���š�ע�⣺�� strFileName ��˫���ţ������е��ļ����пո��޷����ţ�
            rt = mciSendString("open " & """" & strFileName & """" & " alias " & ALIAS_NAME, vbNullString, 0, 0)
            rt = mciSendString("play " & ALIAS_NAME, vbNullString, 0, 0)
            ' ȡ��ý���ļ����ȣ���
            Dim RefStr1 As String * 80
            mciSendString "status " & ALIAS_NAME & " length", RefStr1, Len(RefStr1), 0
            Debug.Print "�ܳ��ȣ�" & Val(RefStr1)
        ElseIf getPalyStatus(ALIAS_NAME) = IsPaused Then
            ' ��������
            rt = mciSendString("resume " & ALIAS_NAME, vbNullString, 0, 0)
            ' ��ȡ��ǰ���Ž��ȣ�����ļ����ȶ��ԣ���
            Dim RefStr2 As String * 80
            mciSendString "status " & ALIAS_NAME & " position", RefStr2, Len(RefStr2), 0
            Debug.Print "�Ѳ��ţ�" & Val(RefStr2)
        Else
            ' ֹͣ���Ų��ر�������
            rt = mciSendString("stop " & ALIAS_NAME, vbNullString, 0, 0)
            rt = mciSendString("close " & ALIAS_NAME, vbNullString, 0, 0)
        End If
    ElseIf setStatus = 1 Then ' ��ͣ
        rt = mciSendString("pause " & ALIAS_NAME, vbNullString, 0, 0)
    Else ' ֹͣ
        rt = mciSendString("stop " & ALIAS_NAME, vbNullString, 0, 0)
        rt = mciSendString("close " & ALIAS_NAME, vbNullString, 0, 0)
    End If
End Sub
' ��õ�ǰý���״̬�����ڲ��ţ���ͣ��ֹͣ��
Private Function getPalyStatus(Optional strAlias As String = "mySound") As PlayStatus
    Dim sl As String * 255
    mciSendString "status " & strAlias & " mode", sl, Len(sl), 0
    If UCase$(Left$(sl, 7)) = "PLAYING" Or Left$(sl, 2) = "����" Then
        getPalyStatus = IsPlaying
    ElseIf UCase$(Left$(sl, 6)) = "PAUSED" Or Left$(sl, 2) = "��ͣ" Then
        getPalyStatus = IsPaused
    Else
        getPalyStatus = IsStopped
    End If
End Function
' =========================================================================================
' ==== �����ļ��Ĳ��ţ���ȥ����============================================================
' =========================================================================================



' =========================================================================================
' ==== ����Ի��� ��������=================================================================
' =========================================================================================
' ��ʾ�Ի���֮ǰ���Զ�������Ի�����ۡ�
Private Sub CustomizeFontDialog(ByVal hWnd As Long)
    Dim rcDlg As RECT, hWndParent As Long
    Dim pt As POINTAPI, W As Long, H As Long
    ' �Ի���ĸ����ھ����
    hWndParent = GetParent(hWnd)
    ' ���öԻ�������λ�ã�ֻ�ж���Ļ���ĺ����������ģ��������ܣ�
    GetWindowRect hWnd, rcDlg ' ȡ�öԻ������
    W = rcDlg.Right - rcDlg.Left: H = rcDlg.Bottom - rcDlg.Top + 120 ' �Ի�����ߣ��߶�Ҫ�Ӹ�Ԥ���ı���ߣ���
    If m_dlgStartUpPosition = vbStartUpScreen Then ' ���ƶ�����Ļ����
        MoveWindow hWnd, (Screen.Width \ Screen.TwipsPerPixelX - W) \ 2, (Screen.Height \ Screen.TwipsPerPixelY - H) \ 2, W, H, True
    ElseIf m_dlgStartUpPosition = vbStartUpOwner Then ' ����������
        Dim rcOwner As RECT, T As Long ' ȡ�öԻ���ĸ����ھ���
        GetWindowRect hWndParent, rcOwner
        T = rcOwner.Top + (rcOwner.Bottom - rcOwner.Top - H) \ 2
        If T < 0 Then T = 0 ' ��֤�Ի��򲻳�����Ļ���ˣ�
        MoveWindow hWnd, rcOwner.Left + (rcOwner.Right - rcOwner.Left - W) \ 2, T, W, H, True
    Else ' ����ر�Ҫ�����߶�Ҫ�䣬������ӵ��ı����޷���ʾ��
        MoveWindow hWnd, rcDlg.Left, rcDlg.Top, W, H, True
    End If
    ' ����λ�ã���Ϊ���ֶԻ�����������ԣ�����һ������ʵ�֣�Ч�����ã������ˣ�
    'setDlgStartUpPosition hWndParent, GetParent(hWndParent)
    
    ' ����Ԥ���ı������С�̶�!��
    Dim rcP As RECT, ptL As POINTAPI  ' ����Ԥ���ı���λ�úʹ�С��rcL �Ǹ�˵����ǩ���Σ�ID = 1093 ?��
    GetWindowRect GetDlgItem(hWnd, enumFONT_CTL.stc_Description), rcP
    ptL.X = rcP.Left: ptL.y = rcP.Top
    ScreenToClient hWnd, ptL ' ptL����ת������ܵõ���Ҫ�Ľ����
    ' ��ʼ����Ԥ���ı���  �� Or WS_VISIBLE ���ڴ���ʱ��ʾ�� Or ES_READONLY ����Ϊֻ����
    Dim sT As String
    'Debug.Print " rcDlg.Bottom - rcP.Bottom + 20 " & rcDlg.Bottom - rcP.Bottom + 20 ' �ı���߶Ȼ�䣿������������
    sT = App.LegalCopyright & vbCrLf _
        & "һ�����������߰˾�ʮ" & vbCrLf & "Ҽ��������½��ƾ�ʰ" & vbCrLf _
        & "ABCDEFGHILMNOPQRSTUVWXYZ" & vbCrLf & "abcdefghilmnopqrstuvwxyz" & vbCrLf & "0123456789"
    hWndFontPreview = CreateWindowEx(WS_EX_STATICEDGE Or WS_EX_TOPMOST, _
        "Edit", sT, _
        WS_BORDER Or WS_CHILD Or WS_VISIBLE Or WS_HSCROLL Or WS_VSCROLL Or ES_AUTOHSCROLL Or ES_AUTOHSCROLL Or ES_MULTILINE Or ES_LEFT Or ES_WANTRETURN, _
        ptL.X, ptL.y + (rcP.Bottom - rcP.Top), _
        W - 20, 120, _
        hWnd, 0&, App.hInstance, &H520)
    ' �����µ����壬Ĭ�����壡����
    Dim NewFont As Long, lpLF As LOGFONT
    With lpLF
        .lfCharSet = 134
        '.lfFaceName = "����"
        .lfItalic = False
        .lfStrikeOut = False
        .lfUnderline = False
        .lfWeight = 520
    End With
    NewFont = CreateFontIndirect(lpLF)
    ' �����ı�������
    SendMessage hWndFontPreview, WM_SETFONT, NewFont, 0
End Sub
' ��������Ի���Ԥ����������Ч����̬�仯����ȡ WM_COMMAND ��Ϣ���ж������ʽ����Щ�仯����
Private Sub mSetFontPreview(ByVal hWnd As Long, Optional wP As Long = 0&)
    Dim hFontToUse As Long, hdc As Long, RetValue As Long
    Dim lpLF As LOGFONT
    Dim tBuf As String * 80, sFontName As String
    Dim iIndex As Long, dwRGB As Long ' dwRGB ������ɫֵ��
    
     ' ȡ��ѡ���������Ϣ
    SendMessage hWnd, WM_CHOOSEFONT_GETLOGFONT, wP, lpLF
    hFontToUse = CreateFontIndirect(lpLF): If hFontToUse = 0 Then Exit Sub
    hdc = GetDC(hWnd)
    SelectObject hdc, hFontToUse
    RetValue = GetTextFace(hdc, 79, tBuf)
    sFontName = Mid$(tBuf, 1, RetValue)

    ' ȡ��ѡ���������ɫ
    iIndex = SendDlgItemMessage(hWnd, enumFONT_CTL.cbo_Color, CB_GETCURSEL, 0&, 0&)    ' cmb4
    If iIndex <> CB_ERR Then
        dwRGB = SendDlgItemMessage(hWnd, enumFONT_CTL.cbo_Color, CB_GETITEMDATA, iIndex, 0&)
    End If
    ' �����µ����壬��ɫ��Ϣû�У�����
    Dim NewFont As Long
    NewFont = CreateFontIndirect(lpLF)
'    NewFont = CreateFont(Abs(lpLF.lfHeight * (72 / GetDeviceCaps(hDC, LOGPIXELSY))), 0, 0, 0, _
              lpLF.lfWeight, lpLF.lfItalic, lpLF.lfUnderline, lpLF.lfStrikeOut, _
              lpLF.lfCharSet, OUT_DEFAULT_PRECIS, CLIP_LH_ANGLES, _
              ANTIALIASED_QUALITY Or PROOF_QUALITY, DEFAULT_PITCH Or FF_DONTCARE, _
              sFontName)
    ' �����ı�������
    SendMessage hWndFontPreview, WM_SETFONT, NewFont, 0
    ' �����ı���������ɫ����֪Ϊʲô��Ԥ����ɫ�ı䡣�����޷�ʵ�֣�������
'    SendMessage hWndFontPreview, 4103, 0, ByVal dwRGB
    If SetTextColor(GetDC(hWndFontPreview), dwRGB) = &HFFFF Then MsgBox "ʧ�ܣ�����������ɫ����", vbCritical
'    frmMain.txtNewCaption(2).ForeColor = dwRGB
'    frmMain.txtNewCaption(1).ForeColor = GetTextColor(GetDC(frmMain.txtNewCaption(2).hWnd))
'    Dim cDC As Long, chWnd As Long
'    chWnd = GetDlgItem(hWnd, enumFONT_CTL.btn_Apply)
'    cDC = GetDC(cDC)
'    Call SetTextColor(cDC, dwRGB)
'    Call SendMessage(chWnd, CDM_SetControlText, enumFONT_CTL.btn_Apply, ByVal "m_strControlsCaption(I)")
'    If dwRGB = GetTextColor(cDC) Then frmMain.BackColor = GetTextColor(cDC)
'    Dim sl As Long ' �ı��������ָ�����Ҫȡ�ã���ʱ�̶�һ��ֵ����
'    sl = 1024
'    Dim s As String: s = String(sl, 0)
'    GetWindowText hWndFontPreview, s, sl
'    Debug.Print Replace$(s, Chr$(0), "")
'    ' ������������
'    SendMessage hWndFontPreview, WM_SETTEXT, -1, ByVal s & vbCrLf & App.LegalCopyright

    ' �ͷ���Դ
    ReleaseDC hWnd, hdc
End Sub

Private Function LOWORD(Param As Long) As Long
    LOWORD = Param And &HFFFF&
End Function
Private Function HIWORD(Param As Long) As Long
    HIWORD = Param \ &H10000 And &HFFFF&
End Function
' =========================================================================================
' ==== ����Ի��� ��������=================================================================
' =========================================================================================
