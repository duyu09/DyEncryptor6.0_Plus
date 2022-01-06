Attribute VB_Name = "MCDHook"
Option Explicit
' --- API º¯Êý ÉêÃ÷
' ÊÍ·Å³ÌÐòÄÚ´æ
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
' È¡µÃ¿Ø¼þÏà¶ÔÆÁÄ»×óÉÏ½ÇµÄ×ø±êÖµ£¡£¨µ¥Î»£ºÏñËØ£¿£¡£©
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWndParent As Long, ByVal hWndChildAfter As Long, ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function UpdateWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, lpString As Any) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal e As Long, ByVal O As Long, ByVal W As Long, ByVal I As Long, ByVal u As Long, ByVal s As Long, ByVal C As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal f As String) As Long

' VB È¡µÃÍ¼Æ¬´óÐ¡
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long

' =========================================================================================
' ==== ÉùÒôÎÄ¼þµÄ²¥·Å£¨¿ÉÈ¥µô£©============================================================
' =========================================================================================
'API ÉêÃ÷ Ê¹ÓÃPlaySoundº¯Êý²¥·ÅÉùÒô
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, _
    ByVal hModule As String, ByVal dwFlags As Long) As Long
'API ÉêÃ÷ Ê¹ÓÃsndPlaySoundº¯Êý²¥·ÅÉùÒô£¬ËüÊÇ PlaySound º¯ÊýµÄ×Ó¼¯£¿£¡
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, _
    ByVal uFlags As Long) As Long
Private Declare Function sndStopSound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszNull As Long, ByVal uFlags As Long) As Long
'¹Ø±ÕÉùÒô
'sndPlaySound Null, SND_ASYNC
'PlaySound 0,0,0
' ¸ß¼¶Ã½Ìå²¥·Åº¯Êý
Private Declare Function mciExecute Lib "winmm.dll" (ByVal lpstrCommand As String) As Long
' mciSendString ÊÇÓÃÀ´²¥·Å¶àÃ½ÌåÎÄ¼þµÄAPIÖ¸Áî£¬¿ÉÒÔ²¥·ÅMPEG,AVI,WAV,MP3,µÈµÈ
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
' Multimedia Command Strings: http://msdn.microsoft.com/en-us/library/ms712587.aspx
' MCI Command Strings:http://msdn.microsoft.com/en-us/library/ms710815(VS.85).aspx

' --- for PlaySound \ sndPlaySound
Private Const SND_ASYNC = &H1 ' play asynchronously ÔÚ²¥·ÅµÄÍ¬Ê±¼ÌÐøÖ´ÐÐÒÔºóµÄÓï¾ä
Private Const SND_FILENAME = &H20000 ' name is a file name
Private Const SND_LOOP = &H8 ' loop the sound until next sndPlaySound Ò»Ö±ÖØ¸´²¥·ÅÉùÒô£¬Ö±µ½¸Ãº¯Êý¿ªÊ¼²¥·ÅµÚ¶þ¸öÉùÒôÎªÖ¹
Private Const SND_MEMORY = &H4         '  lpszSoundName points to a memory file ²¥·ÅÄÚ´æÖÐµÄÉùÒô, Æ©Èç×ÊÔ´ÎÄ¼þÖÐµÄÉùÒô
Private Const SND_NODEFAULT = &H2         '  silence not default, if sound not found
Private Const SND_NOSTOP = &H10        '  don't stop any currently playing sound
Private Const SND_NOWAIT = &H2000      '  don't wait if the driver is busy
Private Const SND_PURGE = &H40               '  purge non-static events for task
Private Const SND_RESERVED = &HFF000000  '  In particular these flags are reserved
Private Const SND_RESOURCE = &H40004     '  name is a resource name or atom
Private Const SND_SYNC = &H0         '  play synchronously (default) ²¥·ÅÍêÉùÒôÖ®ºóÔÙÖ´ÐÐºóÃæµÄÓï¾ä
Private Const SND_TYPE_MASK = &H170007
Private Const SND_VALID = &H1F        '  valid flags          / ;Internal /
Private Const SND_VALIDFLAGS = &H17201F    '  Set of valid flag bits.  Anything outside

Private Enum PlayStatus ' ÉùÒô²¥·Å×´Ì¬£¡
    IsPlaying = 0
    IsPaused = 1
    IsStopped = 2
End Enum
' =========================================================================================
' ==== ÉùÒôÎÄ¼þµÄ²¥·Å£¨¿ÉÈ¥µô£©============================================================
' =========================================================================================


' =========================================================================================
' ==== ×ÖÌå¶Ô»°¿ò £¨µ¥¶À£©=================================================================
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
' hDC ºÜÖØÒª£¡
' ËµÃ÷ ÉèÖÃµ±Ç°ÎÄ±¾ÑÕÉ«¡£ÕâÖÖÑÕÉ«Ò²³ÆÎª¡°Ç°¾°É«¡± ·µ»ØÖµ Long£¬ÎÄ±¾É«µÄÇ°Ò»¸öRGBÑÕÉ«Éè¶¨¡£CLR_INVALID±íÊ¾Ê§°Ü¡£

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
Rem per maggior praticità ho enumerato tutti i controlli della
Rem finestra Carattere
Rem ------------------------------------------------------------
Public Enum enumFONT_CTL ' ×ÖÌå¶Ô»°¿òÉÏµÄ¿Ø¼þ ID
    stc_FontName = 1088 ' ×ÖÌå(&F): ±êÇ©
    edt_FontName = 1001 ' ×ÖÌåÃû³Æ ÎÄ±¾¿ò£¿£¿
    cbo_FontName = &H470  ' ×ÖÌåÃû³Æ ÏÂÀ­¿ò£¿£¿66672
    
    stc_BoldItalic = 1089 ' ×ÖÐÎ(&Y): ±êÇ©
    edt_BoldItalic = 1001 ' ×ÖÐÎ ÎÄ±¾¿ò£¿£¿
    cbo_BoldItalic = &H471  ' ×ÖÐÎ ÏÂÀ­¿ò£¿£¿66673
    
    stc_Size = 1090 ' ´óÐ¡(&S): ±êÇ©
    edt_Size = 1001 ' ´óÐ¡ ÎÄ±¾¿ò£¿£¿
    cbo_Size = &H472  ' ´óÐ¡ ÏÂÀ­¿ò£¿£¿66674
    
    btn_Ok = 1 ' È·¶¨(&O) °´Å¥
    btn_Cancel = 2 ' È¡Ïû(&C) °´Å¥
    btn_Apply = 1026 ' Ó¦ÓÃ(&A) °´Å¥
    btn_Help = 1038 ' °ïÖú(&H) °´Å¥
    
    btn_Effects = 1072 ' Ð§¹û ×éºÏ¿ò
    btn_Strikethrough = &H410 ' É¾³ýÏß(&K) °´Å¥
    btn_Underline = &H411 ' ÏÂ»®Ïß(&U) °´Å¥
    stc_Color = &H443 ' ÑÕÉ«(&C): ±êÇ©
    cbo_Color = &H473 ' ÑÕÉ« ÏÂÀ­¿ò£¿£¿66675
    
    btn_Sample = 1073 ' Ê¾Àý×éºÏ¿ò
    stc_SampleText = &H444 ' Ê¾Àý±êÇ©£ºÎ¢ÈíÖÐÎÄÈí¼þ
    
    stc_Charset = 1094 ' ×Ö·û¼¯(&R): ±êÇ©
    cbo_Charset = &H474 ' ×Ö·û¼¯ÏÂÀ­¿ò
    stc_Description = 1093 ' ×ÖÌåÃèÊö±êÇ©£º¸Ã×ÖÌåÓÃÓÚÏÔÊ¾¡£´òÓ¡Ê±½«Ê¹ÓÃ×î½Ó½üµÄÆ¥Åä×ÖÌå¡£

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
' ==== ×ÖÌå¶Ô»°¿ò£¨µ¥¶À£©==================================================================
' =========================================================================================



' =========================================================================================
' ==== ÑÕÉ«¶Ô»°¿ò £¨µ¥¶À£©=================================================================
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
' ==== ÑÕÉ«¶Ô»°¿ò£¨µ¥¶À£©=================================================================
' =========================================================================================


' --- ³£Êý ÉêÃ÷
' for Windows ÏûÏ¢ ³£Êý
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

' for ¶Ô»°¿òÉÏµÄÏûÏ¢
Private Const CDM_First = (WM_USER + 100)                   '/---
Private Const CDM_GetSpec = (CDM_First + &H0)               'È¡µÃÎÄ¼þÃû
Private Const CDM_GetFilePath = (CDM_First + &H1)           'È¡µÃÎÄ¼þÃûÓëÄ¿Â¼
Private Const CDM_GetFolderPath = (CDM_First + &H2)         'È¡µÃÂ·¾¶
Private Const CDM_GetFolderIDList = (CDM_First + &H3)       '
Private Const CDM_SetControlText = (CDM_First + &H4)        'ÉèÖÃ¿Ø¼þÎÄ±¾
Private Const CDM_HideControl = (CDM_First + &H5)           'Òþ²Ø¿Ø¼þ
Private Const CDM_SetDefext = (CDM_First + &H6)             '
Private Const CDM_Last = (WM_USER + 200)                    '\---

Private Const CDN_First = (-601)                            '/---
Private Const CDN_InitDone = (CDN_First - &H0)              '³õÊ¼»¯Íê³É
Private Const CDN_SelChange = (CDN_First - &H1)             'Ñ¡ÔñÎÄ¼þ¸Ä±ä
Private Const CDN_FolderChange = (CDN_First - &H2)          'Ä¿Â¼¸Ä±ä
Private Const CDN_ShareViolation = (CDN_First - &H3)        '
Private Const CDN_Help = (CDN_First - &H4)                  'µãÁË°ïÖú
Private Const CDN_FileOK = (CDN_First - &H5)                'µãÁËÈ·¶¨
Private Const CDN_TypeChange = (CDN_First - &H6)            '¹ýÂËÀàÐÍ¸Ä±ä
Private Const CDN_IncludeItem = (CDN_First - &H7)           '
Private Const CDN_Last = (-699)                             '\---
  
' for ¶Ô»°¿òÉÏ¿Ø¼þµÄ ID
Private Const ID_FolderLabel   As Long = &H443              '¡°²éÕÒ·¶Î§(&I)¡±±êÇ©
Private Const ID_FolderCombo   As Long = &H471              'Ä¿Â¼ÏÂÀ­¿ò
Private Const ID_ToolBar       As Long = &H440              '¹¤¾ßÀ¸£¨ÌØ±ð×¢Òâ£ºÎÞ·¨Í¨¹ý CDMoveOriginControl º¯ÊýÒÆ¶¯£¡£©
Private Const ID_ToolBarWin2K  As Long = &H4A0              '¿ì½ÝÄ¿Â¼Çø£¨°æ±¾>=Win2K£©

' ÁÐ±í¿ò£¨ÁÐ³öÎÄ¼þµÄ×î´óÇøÓò£©
Private Const ID_List0         As Long = &H460              ' Ê¹ÓÃÕâ¸öÓÐÐ§£¡£¿£¡
Private Const ID_List1         As Long = &H461
Private Const ID_List2         As Long = &H462

Private Const ID_OK            As Long = 1                  '¡°È·¶¨(&O)¡±°´¼ü
Private Const ID_Cancel        As Long = 2                  '¡°È¡Ïû(&C)¡±°´¼ü
Private Const ID_Help          As Long = &H40E              '¡°°ïÖú(&H)¡±°´¼ü
Private Const ID_ReadOnly      As Long = &H410              '¡°Ö»¶Á¡±¶àÑ¡¿ò

Private Const ID_FileTypeLabel As Long = &H441              '¡°ÎÄ¼þÀàÐÍ(&T)¡±±êÇ©
Private Const ID_FileNameLable As Long = &H442              '¡°ÎÄ¼þÃû(&N)¡±±êÇ©
'¡°ÎÄ¼þÀàÐÍ(&T)¡±ÏÂÀ­¿ò
Private Const ID_FileTypeCombo0 As Long = &H470             ' Ê¹ÓÃÕâ¸öÓÐÐ§£¡£¿£¡
Private Const ID_FileTypeCombo1 As Long = &H471
Private Const ID_FileTypeCombo2 As Long = &H472
Private Const ID_FileTypeComboC As Long = &H47C             '¡°ÎÄ¼þÃû(&N)¡±ÎÄ±¾¿ò
Private Const ID_FileNameText  As Long = &H480              '¡°ÎÄ¼þÃû(&N)¡±ÎÄ±¾¿ò£¨ÐÂÍâ¹ÛÊ±²»ÊÇËü£¡£©

' for SendMessage È¡µÃ¸´Ñ¡¿òÊÇ·ñÑ¡ÖÐ£¿
Private Const BM_GETCHECK = &HF0

' for CreateWindowEx ´´½¨Ô¤ÀÀÎÄ±¾¿ò
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
Private Const ES_MULTILINE = &H4&       ' ÎÄ±¾ÔÊÐí¶àÐÐ
Private Const ES_READONLY = &H800&      ' ½«±à¼­¿òÉèÖÃ³ÉÖ»¶ÁµÄ
Private Const ES_CENTER = &H1&          ' ÎÄ±¾ÏÔÊ¾¾ÓÖÐ
Private Const ES_WANTRETURN = &H1000&   ' Ê¹¶àÐÐ±à¼­Æ÷½ÓÊÕ»Ø³µ¼üÊäÈë²¢»»ÐÐ¡£Èç¹û²»Ö¸¶¨¸Ã·ç¸ñ£¬°´»Ø³µ¼ü»áÑ¡ÔñÈ±Ê¡µÄÃüÁî°´Å¥£¬ÕâÍùÍù»áµ¼ÖÂ¶Ô»°¿òµÄ¹Ø±Õ¡£

' --- for CreateFont ×ÖÌåÐÅÏ¢³£Êý
Private Const CLIP_LH_ANGLES            As Long = 16 ' ×Ö·ûÐý×ªËùÐèÒªµÄ
Private Const PROOF_QUALITY             As Long = 2
Private Const TRUETYPE_FONTTYPE         As Long = &H4
Private Const ANTIALIASED_QUALITY       As Long = 4
Private Const DEFAULT_CHARSET           As Long = 1
Private Const FF_DONTCARE = 0    '  Don't care or don't know.
Private Const DEFAULT_PITCH = 0
Private Const OUT_DEFAULT_PRECIS = 0

' --- Ã¶¾Ù ÉêÃ÷
Public Enum PreviewPosition ' Ô¤ÀÀÍ¼Æ¬¿òÎ»ÖÃ
    ppNone = -1 ' ÉèÎª´ËÖµÊ±£¬²»ÏÔÊ¾£¡
    ppTop = 0
    ppLeft = 1
    ppRight = 2
    ppBottom = 3
End Enum
Public Enum DialogStyle ' ¶Ô»°¿ò·ç¸ñ£¬´ò¿ª£¿±£´æ£¿×ÖÌå£¿ÑÕÉ«£¿
    ssOpen = 0
    ssSave = 1
    ssFont = 2
    ssColor = 3
End Enum
Private Enum FileType
    ffText = 0      ' ÎÄ±¾ Ô¤ÀÀ£¨Ä¬ÈÏÖµ£¬ÈÎºÎÎÄ¼þ¿ÉÒÔÒÔÎÄ±¾·½Ê½´ò¿ª£¿£¡£©
    ffPicture = 1   ' Í¼Æ¬ Ô¤ÀÀ
    ffWave = 2      ' Wave ²¨ÐÎÎÄ¼þ Ô¤ÀÀ£¬»­³öÉùÒô²¨ÐÎ£¡£¡
    ffAudio = 3     ' Ò»°ãÒôÆµÎÄ¼þ£¬Ìí¼Ó²¥·Å¡¢ÔÝÍ£¡¢Í£Ö¹°´Å¥£¬½øÐÐÔ¤ÀÀ¡£API²¥·ÅÉùÒô£¡£¡
End Enum

' --- ½á¹¹Ìå ÉêÃ÷
' for CopyMemory È¡µÃ¶Ô»°¿òÄÄÐ©¿Ø¼þ¸Ä±ä£¿
Private Type NMHDR
    hwndFrom   As Long
    idFrom   As Long
    code   As Long
End Type
' ×ø±ê£¿
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
Private Type BITMAP ' È¡µÃBITMAP½á¹¹Ìå
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type
Private Type PicInfo ' Í¼Æ¬¿í¡¢¸ß
    picWidth As Long
    picHeight As Long
End Type

' --- Ë½ÓÐ±äÁ¿ ÉêÃ÷
Private procOld As Long ' ±£´æÔ­ ´°ÌåÊôÐÔµÄ±äÁ¿£¬ÆäÊµÊÇÄ¬ÈÏµÄ ´°Ìåº¯Êý µÄµØÖ·
Private hWndTextView As Long ' ¶¯Ì¬´´½¨µÄÔ¤ÀÀÎÄ±¾¿ò ¾ä±ú
Private hWndButtonPlay(0 To 2) As Long ' 3 ¸ö²¥·Å°´Å¥ ¾ä±ú
Private strSelFile As String ' Ñ¡ÖÐµÄÎÄ¼þÂ·¾¶

' ==== ×ÖÌå¶Ô»°¿ò £¨µ¥¶À£©=================================================================
Private hWndFontPreview As Long ' ×ÖÌåÔ¤ÀÀÎÄ±¾¿ò
' ==== ×ÖÌå¶Ô»°¿ò £¨µ¥¶À£©=================================================================

' --- ¹«¹²±äÁ¿ ÉêÃ÷...Îª CCommonDialog ·þÎñ£¡£¡
Public IsReadOnlyChecked As Boolean ' Ö¸Ê¾ÊÇ·ñÑ¡¶¨Ö»¶Á¸´Ñ¡¿ò
Public WhichStyle As DialogStyle ' ¶Ô»°¿ò·ç¸ñ£¬´ò¿ª£¿±£´æ£¿×ÖÌå£¿ÑÕÉ«£¿

' ÌØ±ðÌØ±ð×¢Òâ£ºÍ¼Æ¬¿òÉè¼ÆÊ±±ØÐëÓÐÍ¼Æ¬£¬·ñÔòµÚ¶þ´Îµ¯³ö¶Ô»°¿òÊ±Í¼Æ¬¿òÏûÊ§£¿£¡£¡£¡ÇÒ´°ÌåÉÏÒª·ÅÁ½¸ö¿ÕµÄÍ¼Æ¬¿ò£¨²»×÷ÈÎºÎÊÂ£¬µ±°ÚÉè£¡£¡£¡£©
Public m_picLogoPicture As PictureBox ' ³ÌÐò±êÖ¾Í¼Æ¬¿òÍ¼Æ¬
Public m_picPreviewPicture As PictureBox ' Ô¤ÀÀÍ¼Æ¬¿òÍ¼Æ¬
Public m_ppLogoPosition As PreviewPosition ' ³ÌÐò±êÖ¾Í¼Æ¬¿òÎ»ÖÃ
Public m_ppPreviewPosition As PreviewPosition ' Ô¤ÀÀÍ¼Æ¬¿òÎ»ÖÃ
Public m_dlgStartUpPosition As StartUpPositionConstants ' ¶Ô»°¿òÆô¶¯Î»ÖÃ£¿
Public m_blnHideControls(0 To 8) As Boolean ' ÊÇ·ñÒþ²Ø¶Ô»°¿òÉÏµÄ¿Ø¼þ£¿£¨¿ÉÈ¥µô£©
Public m_strControlsCaption(0 To 6) As String ' ¶Ô»°¿òÉÏµÄ¿Ø¼þµÄÎÄ×Ö£¿£¨¿ÉÈ¥µô£©

' ################################################################################################
' »Øµ÷º¯Êý£¬ÓÃÀ´½ØÈ¡ÏûÏ¢£¬ÈÃ¶¯Ì¬´´½¨µÄ¿Ø¼þ¿ÉÒÔÏìÓ¦ÏûÏ¢£¡£¨×¢Òâ£ºÊÇ½ØÈ¡¶Ô»°¿òÏûÏ¢£¡£©
' ################################################################################################
Private Function WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, _
                                              ByVal wParam As Long, ByVal lParam As Long) As Long
    ' È·¶¨½ÓÊÕµ½µÄÊÇÊ²Ã´ÏûÏ¢
    Select Case iMsg
        Case WM_COMMAND ' µ¥»÷
            Dim I As Integer
            For I = 0 To 2
                If lParam = hWndButtonPlay(I) Then Call B3Button_Click(I)
            Next I
'        Case WM_LBUTTONDOWN ' Êó±ê×ó¼ü°´ÏÂ
'            Debug.Print "WM_LBUTTONDOWN " & lParam
    End Select
  
    ' Èç¹û²»ÊÇÎÒÃÇÐèÒªµÄÏûÏ¢£¬Ôò´«µÝ¸øÔ­À´µÄ´°Ìåº¯Êý´¦Àí
    WindowProc = CallWindowProc(procOld, hWnd, iMsg, wParam, lParam)

End Function
' ÉèÖÃ¿ªÊ¼ºÍ½áÊøµÄÁ½¸ö¹ý³Ì£¡£¡£¡
Private Sub CDHook(ByVal hWnd As Long)
    ' Õû¸öprocOld±äÁ¿ÓÃÀ´´æ´¢´°¿ÚµÄÔ­Ê¼²ÎÊý£¬ÒÔ±ã»Ö¸´
    ' µ÷ÓÃÁË SetWindowLong º¯Êý£¬ËüÊ¹ÓÃÁË GWL_WNDPROC Ë÷ÒýÀ´´´½¨´°¿ÚÀàµÄ×ÓÀà£¬Í¨¹ýÕâÑùÉèÖÃ
    ' ²Ù×÷ÏµÍ³·¢¸ø´°ÌåµÄÏûÏ¢½«ÓÉ»Øµ÷º¯Êý (WindowProc) À´½ØÈ¡£¬ AddressOfÊÇ¹Ø¼ü×ÖÈ¡µÃº¯ÊýµØÖ·
    procOld = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf WindowProc)
             ' AddressOfÊÇÒ»ÔªÔËËã·û£¬ËüÔÚ¹ý³ÌµØÖ·´«ËÍµ½ API ¹ý³ÌÖ®Ç°£¬ÏÈµÃµ½¸Ã¹ý³ÌµÄµØÖ·
End Sub
Private Sub CDUnHook(ByVal hWnd As Long)
    ' ´Ë¾ä¹Ø¼ü£¬°Ñ´°¿Ú£¨²»ÊÇ´°Ìå£¬¶øÊÇ¾ßÓÐ¾ä±úµÄÈÎÒ»¿Ø¼þ£©µÄÊôÐÔ¸´Ô­
    Call SetWindowLong(hWnd, GWL_WNDPROC, procOld)
End Sub
' ################################################################################################
' »Øµ÷º¯Êý£¬ÓÃÀ´½ØÈ¡ÏûÏ¢£¬ÈÃ¶¯Ì¬´´½¨µÄ¿Ø¼þ¿ÉÒÔÏìÓ¦ÏûÏ¢£¡
' ################################################################################################


' »Øµ÷º¯Êý£¬¶Ô»°¿òÏÔÊ¾Ê±ÒªÊ¹ÓÃ£¡£¡£¡
Public Function CDCallBackFun(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
On Error GoTo CDCallBack_Error
'    Debug.Print "&H" + Hex$(hWnd); ":",
    Dim retV As Long ' º¯Êý·µ»ØÖµ£¿£¡
    
    ' È¡µÃ¸¸´°Ìå¾ä±ú£¿£¨½öÊÇ´ò¿ª¡¢±£´æ¶Ô»°¿ò¾ä±ú£¿£¡£¬×ÖÌåÊ±£¬hWnd ÊÇ¶Ô»°¿ò¾ä±ú£¡£¡£©
    Dim hWndParent As Long: hWndParent = GetParent(hWnd)

    ' ÅÐ¶ÏÏûÏ¢£¬¼ì²âÊÇ·ñÎªÐè´¦ÀíµÄÏûÏ¢
    Select Case uMsg
        Case WM_INITDIALOG ' ¶Ô»°¿ò³õÊ¼»¯Ê±£¬
            Debug.Print "WM_INITDIALOG", "&H" + Hex(wParam), "&H" + Hex(lParam)
            ' Ë½ÓÐ±äÁ¿³õÊ¼»¯£¡
            procOld = 0: hWndTextView = 0: strSelFile = ""
            hWndButtonPlay(0) = 0: hWndButtonPlay(1) = 0: hWndButtonPlay(2) = 0
            ' ÏÔÊ¾¶Ô»°¿òÖ®Ç°¡£×Ô¶¨Òå×ÖÌå¶Ô»°¿òÍâ¹Û¡£
            CDHook hWndParent ' »Øµ÷º¯Êý£¬ÓÃÀ´½ØÈ¡ÏûÏ¢£¬ÈÃ¶¯Ì¬´´½¨µÄ¿Ø¼þ¿ÉÒÔÏìÓ¦ÏûÏ¢£¡
            If WhichStyle = ssFont Then CustomizeFontDialog hWnd ' ³õÊ¼»¯×ÖÌå¶Ô»°¿ò
            If WhichStyle = ssColor Then setDlgStartUpPosition hWnd, hWndParent ' ³õÊ¼»¯ÑÕÉ«¶Ô»°¿ò£¬Ö»Ðè¸ÄÆô¶¯Î»ÖÃ£¡
            ' ÅÐ¶ÏÓÐÃ»ÓÐÉèÖÃÁ½¸öÍ¼Æ¬¿ò£¿£¿£¡£¡
            ' ÐÞÕýÁËÃ»ÓÐÉèÖÃÔ¤ÀÀ»ò³ÌÐò±êÖ¾Í¼Æ¬¿òÊ±£¬¶Ô»°¿òÎ»ÖÃÎÞ·¨µ÷ÕûµÄÎÊÌâ£»
            If m_picLogoPicture Is Nothing Then m_ppLogoPosition = ppNone
            If m_picPreviewPicture Is Nothing Then m_ppPreviewPosition = ppNone
            
        Case WM_NOTIFY ' ¶Ô»°¿ò±ä»¯Ê±£¬½ö¶Ô´ò¿ª/±£´æ¶Ô»°¿ò£¡£¡£¡
            retV = CDNotify(hWndParent, lParam)
        Case WM_COMMAND ' ½öµ¥»÷ ×ÖÌå¡¢ÑÕÉ«¶Ô»°¿ò ÉÏµÄ¿Ø¼þ£¿£¡
            'Debug.Print LOWORD(wParam); HIWORD(wParam)
            Dim L As Long: L = LOWORD(wParam)
            If WhichStyle = ssFont Then
                If L = enumFONT_CTL.btn_Apply _
                    Or L = enumFONT_CTL.cbo_FontName Or L = enumFONT_CTL.cbo_BoldItalic _
                    Or L = enumFONT_CTL.cbo_Size Or L = enumFONT_CTL.btn_Strikethrough _
                    Or L = enumFONT_CTL.btn_Underline Or L = enumFONT_CTL.cbo_Color _
                    Or L = enumFONT_CTL.cbo_Charset Then ' lParam ¿Ø¼þ¾ä±ú£¿ wParam ²ÎÊý=¿Ø¼þ ID £¡£¡£¡
                    ' ÉèÖÃ×ÖÌå¶Ô»°¿òÔ¤ÀÀ¡££¨ÓÐÐ©µ¥»÷²»¹ÜÓÃÒªË«»÷£¬Ç°3¸öcbo£¡£©
                    mSetFontPreview hWnd
                ElseIf L = enumFONT_CTL.btn_Help Then ' ×ÖÌå¶Ô»°¿ò°ïÖú£¡
                    MsgBox "×ÖÌå¶Ô»°¿ò°ïÖú£¡", vbInformation
                'Else ' ÆäËûµ¥»÷£¬·¢ËÍÏûÏ¢µ¥»÷ Ó¦ÓÃ °´Å¥£¡»¹ÊÇ²»ÐÐ£¡£¿
                '    SendMessage GetDlgItem(hWnd, enumFONT_CTL.btn_Apply), WM_LBUTTONDOWN, 0&, 0&
                End If
            ElseIf WhichStyle = ssColor Then
                If L = enumFONT_CTL.btn_Help Then ' ÑÕÉ«¶Ô»°¿ò°ïÖú£¡¼¸¸ö°´Å¥Í¨ÓÃÒ»¸öIDÖµ£¿£¡
                    MsgBox "ÑÕÉ«¶Ô»°¿ò°ïÖú£¡", vbInformation
                End If
            End If
        Case WM_DESTROY ' ¶Ô»°¿òÏú»ÙÊ±£¬
            Debug.Print "WM_DESTROY", "&H" + Hex(wParam), "&H" + Hex(lParam)
            ' È¡µÃ ÊÇ·ñÑ¡¶¨Ö»¶Á¸´Ñ¡¿ò
            Dim hWndButton As Long: hWndButton = GetDlgItem(hWndParent, ID_ReadOnly)
            IsReadOnlyChecked = SendMessage(hWndButton, BM_GETCHECK, ByVal 0&, ByVal 0&)
            
            ' ÉèÖÃÍ¼Æ¬¿òµ½Ô­À´µÄ¸¸´°¿Ú£¬£¨×ÀÃæ¾ä±ú=0£¬»Ö¸´µ½ÆäÔ­À´µÄ¸¸£¬ÔÙ´Îµ¯³ö¶Ô»°¿òÊ±Í¼Æ¬ÏûÊ§ÁË£¡£¡£¡£©
            ' Èç¹û´°ÌåÉÏÓÐÁ½¸öÍ¼Æ¬¿ò£¬ÔÙ´Îµ¯³öÊ±¾ÍÄÜÕý³£ÏÔÊ¾£¿£¡£¡£¡£©
            If Not m_picLogoPicture Is Nothing Then ' Èô²»¼ÓÅÐ¶Ï£¬»áÒý·¢´íÎó£¡
                ShowWindow m_picLogoPicture.hWnd, SW_HIDE
                Call SetParent(m_picLogoPicture.hWnd, Val(m_picLogoPicture.Tag))
            End If
            If Not m_picPreviewPicture Is Nothing Then
                ShowWindow m_picPreviewPicture.hWnd, SW_HIDE
                Call SetParent(m_picPreviewPicture.hWnd, Val(m_picPreviewPicture.Tag))
            End If
            
            ' Í£Ö¹ÉùÒô Í£Ö¹ÉùÒô£¬ÒòÎªÇ°Ãæ¿ÉÄÜÔÚ²¥·Å£¡£¡£¡£¡' ×¢Òâ£ºÕâÀï²»Ö±½Óµ÷ÓÃ B3Button_Click 2 £¡
            PlayAudio strSelFile, 2
            sndPlaySound vbNullString, 0 ' Í£Ö¹ Wave ÎÄ¼þ²¥·Å¡£
            
            ' Ïú»Ù´´½¨µÄ¿Ø¼þ
            If hWndTextView Then DestroyWindow hWndTextView
            If hWndButtonPlay(0) Then DestroyWindow hWndButtonPlay(0) ': hWndButtonPlay(0) = 0
            If hWndButtonPlay(1) Then DestroyWindow hWndButtonPlay(1) ': hWndButtonPlay(1) = 0
            If hWndButtonPlay(2) Then DestroyWindow hWndButtonPlay(2) ': hWndButtonPlay(2) = 0
            If hWndFontPreview Then DestroyWindow hWndFontPreview  ' ×ÖÌåÔ¤ÀÀÎÄ±¾¿ò
            ' »Øµ÷º¯Êý£¬ÓÃÀ´½ØÈ¡ÏûÏ¢£¬ÈÃ¶¯Ì¬´´½¨µÄ¿Ø¼þ¿ÉÒÔÏìÓ¦ÏûÏ¢£¡
            CDUnHook hWndParent
            ' ÊÍ·ÅÎïÀíÄÚ´æ£¡£¡£¡²»ÖªÎªÊ²Ã´£¬µ÷³ö¶Ô»°¿òºó£¬³ÌÐòÕ¼ÓÃµÄÎïÀíÄÚ´æ´óÔö£¡£¡£¡
            SetProcessWorkingSetSize GetCurrentProcess(), -1&, -1&
'        Case Else
'            Debug.Print "Else ", "&H" + Hex(wParam), "&H" + Hex(lParam)
    End Select
    CDCallBackFun = retV ' º¯Êý·µ»ØÖµ£¿£¡
    
    On Error GoTo 0
    Exit Function

CDCallBack_Error:
    Debug.Print "CDCallBackFun Error " & Err.Number & " (" & Err.Description & ")"
    Resume Next
End Function
          
' ¶Ô»°¿ò±ä»¯Ê±£¬½øÐÐµ÷Õû¡£½ö¶Ô´ò¿ª/±£´æ¶Ô»°¿ò£¡£¡£¡
Private Function CDNotify(ByVal hWndParent As Long, ByVal lParam As Long) As Long
    Dim hToolBar As Long    ' ¶Ô»°¿òÉÏ¹¤¾ßÀ¸¾ä±ú
    Dim rcTB As RECT        ' ¹¤¾ßÀ¸¾ØÐÎ
    Dim pt As POINTAPI, W As Long, H As Long
    Dim rcDlg As RECT       ' ¶Ô»°¿ò¾ØÐÎ
    Dim picLeft As Long, picTop As Long ' Í¼Æ¬¿òÎ»ÖÃ×ø±ê£¬Á½¸öÍ¼Æ¬¿òÏà»¥Ó°Ïì£¬Ò»¸öÒÆ¶¯Ê±ÒªÅÐ¶ÏÁíÒ»¸öµÄÎ»ÖÃ£¡
    ' == ÖÐ¼ä×î´óµÄÁÐ±í¿ò¾ØÐÎ£¬Í¼Æ¬¿ò Left Top Î»ÖÃµÄ»ù×¼µã¡£¡£¡£
    Dim hWndControl As Long, rcList0 As RECT, ptL As POINTAPI
    hWndControl = GetDlgItem(hWndParent, ID_List0) ' ¸ù¾ÝIDÈ¡µÃ¿Ø¼þ¾ä±ú
    GetWindowRect hWndControl, rcList0 ' È¡µÃ¿Ø¼þ¾ØÐÎ
    ptL.X = rcList0.Left: ptL.y = rcList0.Top
    ScreenToClient hWndParent, ptL ' ptL¾­¹ý×ª»¯ºó²ÅÄÜµÃµ½ÏëÒªµÄ½á¹û£¡
    
    Dim hdr     As NMHDR
    Call CopyMemory(hdr, ByVal lParam, LenB(hdr))
    Select Case hdr.code
        Case CDN_InitDone ' ³õÊ¼»¯Íê³É£¬¶Ô»°¿ò½«ÒªÏÔÊ¾Ê±£¬
            Debug.Print "InitDone"
                        
            ' ===== ÅÐ¶Ï³ÌÐò±êÖ¾Í¼Æ¬¿òÎ»ÖÃ£¬ÒÔµ÷Õû¶Ô»°¿òÍâ¹Û£¨³ß´ç¼°ÆäÉÏµÄ¿Ø¼þÎ»ÖÃ£©£¡
            Dim OffSetX As Long, OffSetY As Long, stpX As Single, stpY As Single  ' ¶Ô»°¿ò´óÐ¡¡¢¿Ø¼þÆ«ÒÆÁ¿£¨ÏñËØ£¡£©
            stpX = Screen.TwipsPerPixelX: stpY = Screen.TwipsPerPixelY ' Twips ×ª»¯Îª Pixels Òª³ýÒÔËûÃÇ£¡
            If m_ppLogoPosition = ppNone Then GoTo NoLogo ' ÅÐ¶ÏÓÐÃ»ÓÐÉèÖÃÁ½¸öÍ¼Æ¬¿ò£¿£¿£¡£¡
            OffSetX = m_picLogoPicture.Width \ stpX: OffSetY = m_picLogoPicture.Height \ stpY
            Dim ClientRect As RECT ' ppBottom Ê±£¡È¡µÃ¶Ô»°¿ò¾ØÐÎ£¬ÓëÆäËû¶¼²»Í¬£¡²»ÖªÎªÊ²Ã´Ö»ÓÐÕâÑù²ÅÐÐ£¡£¡£¡£¿£¿£¿
            Select Case m_ppLogoPosition
                Case ppNone ' ÎÞ³ÌÐò±êÖ¾Í¼Æ¬£¬²»²Ù×÷£¡
                    OffSetX = 0: OffSetY = 0
                    picLeft = 0: picTop = 0
                Case ppLeft ' ³ÌÐò±êÖ¾Í¼Æ¬ ÔÚ×ó¶Ë£¬ÒªÒÆ¶¯¶Ô»°¿òÉÏÔ­À´µÄ¿Ø¼þ£¡
                    ' ¶Ô»°¿òÉÏËùÓÐÔ­Ê¼¿Ø¼þÓÒÒÆ
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
                    CDMoveOriginControl hWndParent, ID_FileTypeComboC, OffSetX  ' ÐÂÍâ¹Û£¡ÒÆ¶¯¶Ô»°¿òÉÏÎÄ¼þÃûÎÄ±¾¿ò£¬
                    CDMoveOriginControl hWndParent, ID_FileNameText, OffSetX
                    
                    ' ÒÆ¶¯¹¤¾ßÀ¸£¡¹¤¾ßÀ¸£¨ÌØ±ð×¢Òâ£ºÎÞ·¨Í¨¹ý CDMoveOriginControl º¯ÊýÒÆ¶¯£¡£©
                    hToolBar = CDGetToolBarHandle(hWndParent)
                    GetWindowRect hToolBar, rcTB
                    pt.X = rcTB.Left
                    pt.y = rcTB.Top
                    ScreenToClient hWndParent, pt
                    MoveWindow hToolBar, pt.X + OffSetX, pt.y, rcTB.Right - rcTB.Left, rcTB.Bottom - rcTB.Top, True
                    
                    ' ¸Ä±ä¶Ô»°¿ò´óÐ¡£¬£¡
                    GetWindowRect hWndParent, rcDlg ' È¡µÃ¶Ô»°¿ò¾ØÐÎ£¬ÔÙÒÆ¶¯£¨Êµ¼ÊÖ»¸Ä±ä¿í¶È£¡£©
                    MoveWindow hWndParent, rcDlg.Left, rcDlg.Top, rcDlg.Right - rcDlg.Left + OffSetX, rcDlg.Bottom - rcDlg.Top, True
                    ' ÉèÖÃ³ÌÐò±êÖ¾Í¼Æ¬
                    ' ÉèÖÃÐÂµÄ£¬²¢±£´æÍ¼Æ¬¿òÔ­À´µÄ¸¸´°¿Ú¾ä±ú£¿£¡
                    m_picLogoPicture.Tag = SetParent(m_picLogoPicture.hWnd, hWndParent)
                    ' ÒÆ¶¯Í¼Æ¬¿ò£¬Top Î»ÖÃ¹Ì¶¨£¬¸ß¶È¹Ì¶¨£¡
                    If m_ppPreviewPosition = ppLeft Then
                        picLeft = 2: picTop = 0
                    'ElseIf m_ppPreviewPosition = ppRight Then' ²»ÐèÒªÅÐ¶Ï£¡
                    ElseIf m_ppPreviewPosition = ppTop Then
                        picLeft = 2: picTop = m_picPreviewPicture.Height \ stpY
                    ElseIf m_ppPreviewPosition = ppBottom Then
                        picLeft = 2: picTop = m_picPreviewPicture.Height \ stpY
                    Else
                        picLeft = 2: picTop = 0
                    End If
                    MoveWindow m_picLogoPicture.hWnd, picLeft, 2, _
                        m_picLogoPicture.Width \ stpX, rcDlg.Bottom - rcDlg.Top + picTop - 29, True
                    ' ¼ÓÔØÍ¼Æ¬
                    'm_picLogoPicture.PaintPicture m_picLogoPicture.Picture, 0, 0, m_picLogoPicture.ScaleWidth * 100, m_picLogoPicture.ScaleHeight
                    ShowWindow m_picLogoPicture.hWnd, SW_SHOW ' ÏÔÊ¾Í¼Æ¬¿ò

                Case ppRight
                    ' ¸Ä±ä¶Ô»°¿ò´óÐ¡£¬£¡
                    GetWindowRect hWndParent, rcDlg ' È¡µÃ¶Ô»°¿ò¾ØÐÎ£¬ÔÙÒÆ¶¯£¨Êµ¼ÊÖ»¸Ä±ä¿í¶È£¡£©
                    MoveWindow hWndParent, rcDlg.Left, rcDlg.Top, rcDlg.Right - rcDlg.Left + OffSetX, rcDlg.Bottom - rcDlg.Top, True
                    ' ÉèÖÃ³ÌÐò±êÖ¾Í¼Æ¬
                    ' ÉèÖÃÐÂµÄ£¬²¢±£´æÍ¼Æ¬¿òÔ­À´µÄ¸¸´°¿Ú¾ä±ú£¿£¡
                    m_picLogoPicture.Tag = SetParent(m_picLogoPicture.hWnd, hWndParent)
                    ' ÒÆ¶¯Í¼Æ¬¿ò£¬£¬Top Î»ÖÃ¹Ì¶¨£¬¸ß¶È¹Ì¶¨£¡
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
                    ' ¼ÓÔØÍ¼Æ¬
                    'm_picLogoPicture.PaintPicture m_picLogoPicture.Picture, 0, 0, m_picLogoPicture.ScaleWidth, m_picLogoPicture.ScaleHeight
                    ShowWindow m_picLogoPicture.hWnd, SW_SHOW ' ÏÔÊ¾Í¼Æ¬¿ò

                Case ppTop ' ³ÌÐò±êÖ¾Í¼Æ¬ ÔÚ¶¥¶Ë£¬ÒªÒÆ¶¯¶Ô»°¿òÉÏÔ­À´µÄ¿Ø¼þ£¡
                    ' ¶Ô»°¿òÉÏËùÓÐÔ­Ê¼¿Ø¼þÏÂÒÆ
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
                    CDMoveOriginControl hWndParent, ID_FileTypeComboC, , OffSetY ' ÐÂÍâ¹Û£¡ÒÆ¶¯¶Ô»°¿òÉÏÎÄ¼þÃûÎÄ±¾¿ò£¬
                    CDMoveOriginControl hWndParent, ID_FileNameText, , OffSetY
                    
                    ' ÒÆ¶¯¹¤¾ßÀ¸£¡¹¤¾ßÀ¸£¨ÌØ±ð×¢Òâ£ºÎÞ·¨Í¨¹ý CDMoveOriginControl º¯ÊýÒÆ¶¯£¡£©
                    hToolBar = CDGetToolBarHandle(hWndParent)
                    GetWindowRect hToolBar, rcTB
                    pt.X = rcTB.Left
                    pt.y = rcTB.Top
                    ScreenToClient hWndParent, pt
                    MoveWindow hToolBar, pt.X, pt.y + OffSetY, rcTB.Right - rcTB.Left, rcTB.Bottom - rcTB.Top, True
                    
                    ' ¸Ä±ä¶Ô»°¿ò´óÐ¡£¬£¡
                    GetWindowRect hWndParent, rcDlg ' È¡µÃ¶Ô»°¿ò¾ØÐÎ£¬ÔÙÒÆ¶¯£¨Êµ¼ÊÖ»¸Ä±ä¸ß¶È£¡£©
                    MoveWindow hWndParent, rcDlg.Left, rcDlg.Top, rcDlg.Right - rcDlg.Left, rcDlg.Bottom - rcDlg.Top + OffSetY, True
                    ' ÉèÖÃ³ÌÐò±êÖ¾Í¼Æ¬
                    ' ÉèÖÃÐÂµÄ£¬²¢±£´æÍ¼Æ¬¿òÔ­À´µÄ¸¸´°¿Ú¾ä±ú£¿£¡
                    m_picLogoPicture.Tag = SetParent(m_picLogoPicture.hWnd, hWndParent) ' GetParent(m_picLogoPicture.hwnd)
                    ' ÒÆ¶¯Í¼Æ¬¿ò£¬Left Î»ÖÃ¹Ì¶¨£¬¿í¶È¹Ì¶¨¡£picLeft + (rcDlg.Right - rcDlg.Left - m_picLogoPicture.Width \ stpX) \ 2 - 3
                    If m_ppPreviewPosition = ppLeft Then
                        picLeft = m_picPreviewPicture.Width \ stpX: picTop = 0
                    ElseIf m_ppPreviewPosition = ppRight Then ' ²»ÐèÒªÅÐ¶Ï£¡
                        picLeft = m_picPreviewPicture.Width \ stpX
                    'ElseIf m_ppPreviewPosition = ppTop Then
                    'ElseIf m_ppPreviewPosition = ppBottom Then
                    End If
                    MoveWindow m_picLogoPicture.hWnd, 5, picTop + 2, _
                        rcDlg.Right - rcDlg.Left + picLeft - 15, m_picLogoPicture.Height \ stpY, True
                    ' ¼ÓÔØÍ¼Æ¬
                    'm_picLogoPicture.PaintPicture m_picLogoPicture.Picture, 0, 0, m_picLogoPicture.ScaleWidth, m_picLogoPicture.ScaleHeight
                    ShowWindow m_picLogoPicture.hWnd, SW_SHOW ' ÏÔÊ¾Í¼Æ¬¿ò

                Case ppBottom
                    ' ¸Ä±ä¶Ô»°¿ò´óÐ¡£¬£¡
                    GetWindowRect hWndParent, rcDlg ' È¡µÃ¶Ô»°¿ò¾ØÐÎ£¬ÔÙÒÆ¶¯£¨Êµ¼ÊÖ»¸Ä±ä¸ß¶È£¡£©
                    MoveWindow hWndParent, rcDlg.Left, rcDlg.Top, rcDlg.Right - rcDlg.Left, rcDlg.Bottom - rcDlg.Top + OffSetY, True
                    ' ÉèÖÃ³ÌÐò±êÖ¾Í¼Æ¬
                    Call GetClientRect(hWndParent, ClientRect) ' ÓÃ rcDlg.Bottom ²»ÐÐ£¡£¡£¡
                    ' ÉèÖÃÐÂµÄ£¬²¢±£´æÍ¼Æ¬¿òÔ­À´µÄ¸¸´°¿Ú¾ä±ú£¿£¡
                    m_picLogoPicture.Tag = SetParent(m_picLogoPicture.hWnd, hWndParent)
                    ' ÒÆ¶¯Í¼Æ¬¿ò£¬Left Î»ÖÃ¹Ì¶¨£¬¿í¶È¹Ì¶¨¡£
                    If m_ppPreviewPosition = ppLeft Then
                        picLeft = m_picPreviewPicture.Width \ stpX: picTop = 0
                    ElseIf m_ppPreviewPosition = ppRight Then
                        picLeft = m_picPreviewPicture.Width \ stpX
                    ElseIf m_ppPreviewPosition = ppTop Then
                        picTop = m_picPreviewPicture.Height \ stpY: picLeft = 0
                    ElseIf m_ppPreviewPosition = ppBottom Then ' ÕâÊ±£¬ÒªÒÆ¶¯±êÖ¾µ½Ô¤ÀÀÏÂÃæ£¡
                        picTop = m_picPreviewPicture.Height \ stpY: picLeft = 0
                    End If
                    MoveWindow m_picLogoPicture.hWnd, 5, picTop + ClientRect.Bottom - OffSetY, _
                        rcDlg.Right - rcDlg.Left + picLeft - 15, m_picLogoPicture.Height \ stpY, True
                    ' ¼ÓÔØÍ¼Æ¬
                    'm_picLogoPicture.PaintPicture m_picLogoPicture.Picture, 0, 0, m_picLogoPicture.ScaleWidth, m_picLogoPicture.ScaleHeight
                    ShowWindow m_picLogoPicture.hWnd, SW_SHOW ' ÏÔÊ¾Í¼Æ¬¿ò

            End Select
NoLogo:
' **********************************************************************************************************
            ' ===== ÅÐ¶ÏÔ¤ÀÀÍ¼Æ¬¿òÎ»ÖÃ£¬ÌØ±ð×¢Òâ£ºÒªÅÐ¶Ï³ÌÐò±êÖ¾Í¼Æ¬¿òÎ»ÖÃ£¿£¡£¡£¡£¡·½·¨£º£¿£¿£¿£¡£¡£¡
            ' Ô¤ÀÀÍ¼Æ¬¿òÎ»ÖÃ¹Ì¶¨Ò»¸öÖµ£¡£¡£¡´óÐ¡ÔÚ×óÓÒ¡¢ÉÏÏÂ·ÖÁ½ÖÖÇé¿ö£º·Ö±ð¹Ì¶¨¸ß¶È¡¢¿í¶È£¡£¡£¡
            If m_ppPreviewPosition = ppNone Then GoTo NoPreview ' ÅÐ¶ÏÓÐÃ»ÓÐÉèÖÃÁ½¸öÍ¼Æ¬¿ò£¿£¿£¡£¡
            OffSetX = m_picPreviewPicture.Width \ stpX: OffSetY = m_picPreviewPicture.Height \ stpY
            Select Case m_ppPreviewPosition
                Case ppNone
                    OffSetX = 0: OffSetY = 0
                    picLeft = 0: picTop = 0
                Case ppLeft ' Ô¤ÀÀÍ¼Æ¬¿ò ÔÚ×ó¶Ë£¬ÒªÒÆ¶¯¶Ô»°¿òÉÏÔ­À´µÄ¿Ø¼þ£¡
                    ' ÖØÐÂÈ¡µÃ LIST¿Ø¼þ¾ØÐÎ£¬¿ÉÄÜÔÚÉÏÃæ±»ÒÆ¶¯ÁË£¡
                    GetWindowRect hWndControl, rcList0
                    ptL.X = rcList0.Left: ptL.y = rcList0.Top
                    ScreenToClient hWndParent, ptL ' ptL¾­¹ý×ª»¯ºó²ÅÄÜµÃµ½ÏëÒªµÄ½á¹û£¡
                    
                    ' ¶Ô»°¿òÉÏËùÓÐÔ­Ê¼¿Ø¼þÓÒÒÆ
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
                    CDMoveOriginControl hWndParent, ID_FileTypeComboC, OffSetX  ' ÐÂÍâ¹Û£¡ÒÆ¶¯¶Ô»°¿òÉÏÎÄ¼þÃûÎÄ±¾¿ò£¬
                    CDMoveOriginControl hWndParent, ID_FileNameText, OffSetX
                    
                    ' ÒÆ¶¯¹¤¾ßÀ¸£¡¹¤¾ßÀ¸£¨ÌØ±ð×¢Òâ£ºÎÞ·¨Í¨¹ý CDMoveOriginControl º¯ÊýÒÆ¶¯£¡£©
                    hToolBar = CDGetToolBarHandle(hWndParent)
                    GetWindowRect hToolBar, rcTB
                    pt.X = rcTB.Left
                    pt.y = rcTB.Top
                    ScreenToClient hWndParent, pt
                    MoveWindow hToolBar, pt.X + OffSetX, pt.y, rcTB.Right - rcTB.Left, rcTB.Bottom - rcTB.Top, True
                    
                    ' ¸Ä±ä¶Ô»°¿ò´óÐ¡£¬£¡
                    GetWindowRect hWndParent, rcDlg ' È¡µÃ¶Ô»°¿ò¾ØÐÎ£¬ÔÙÒÆ¶¯£¨Êµ¼ÊÖ»¸Ä±ä¿í¶È£¡£©
                    MoveWindow hWndParent, rcDlg.Left, rcDlg.Top, rcDlg.Right - rcDlg.Left + OffSetX, rcDlg.Bottom - rcDlg.Top, True
                    ' ÉèÖÃ Ô¤ÀÀÍ¼Æ¬¿ò
                    ' ÉèÖÃÐÂµÄ£¬²¢±£´æÍ¼Æ¬¿òÔ­À´µÄ¸¸´°¿Ú¾ä±ú£¿£¡
                    m_picPreviewPicture.Tag = SetParent(m_picPreviewPicture.hWnd, hWndParent)
                    ' ÒÆ¶¯Í¼Æ¬¿ò£¬Top Î»ÖÃ¹Ì¶¨£¬¸ß¶È¹Ì¶¨£¡
                    picLeft = 5: picTop = ptL.y: W = 5
                    If m_ppLogoPosition = ppLeft Then
                        picLeft = m_picLogoPicture.Width \ stpX + 5
                    'ElseIf m_ppLogoPosition = ppRight Then' ²»ÐèÒªÅÐ¶Ï£¡
                    'ElseIf m_ppLogoPosition = ppTop Then
                    'ElseIf m_ppLogoPosition = ppBottom Then
                    End If
                    MoveWindow m_picPreviewPicture.hWnd, picLeft, picTop, _
                        m_picPreviewPicture.Width \ stpX - W, rcList0.Bottom - rcList0.Top, True
                    ' ¼ÓÔØÍ¼Æ¬
                    'm_picPreviewPicture.PaintPicture m_picPreviewPicture.Picture, 0, 0, m_picPreviewPicture.ScaleWidth, m_picPreviewPicture.ScaleHeight
                    myPaintPicture m_picPreviewPicture, False
                    ShowWindow m_picPreviewPicture.hWnd, SW_SHOW ' ÏÔÊ¾Í¼Æ¬¿ò

                Case ppRight
                    ' ÖØÐÂÈ¡µÃ LIST¿Ø¼þ¾ØÐÎ£¬¿ÉÄÜÔÚÉÏÃæ±»ÒÆ¶¯ÁË£¡
                    GetWindowRect hWndControl, rcList0
                    ptL.X = rcList0.Left: ptL.y = rcList0.Top
                    ScreenToClient hWndParent, ptL ' ptL¾­¹ý×ª»¯ºó²ÅÄÜµÃµ½ÏëÒªµÄ½á¹û£¡
                    
                    ' ¸Ä±ä¶Ô»°¿ò´óÐ¡£¬£¡
                    GetWindowRect hWndParent, rcDlg ' È¡µÃ¶Ô»°¿ò¾ØÐÎ£¬ÔÙÒÆ¶¯£¨Êµ¼ÊÖ»¸Ä±ä¿í¶È£¡£©
                    MoveWindow hWndParent, rcDlg.Left, rcDlg.Top, rcDlg.Right - rcDlg.Left + OffSetX, rcDlg.Bottom - rcDlg.Top, True
                    ' ÉèÖÃ Ô¤ÀÀÍ¼Æ¬¿ò
                    ' ÉèÖÃÐÂµÄ£¬²¢±£´æÍ¼Æ¬¿òÔ­À´µÄ¸¸´°¿Ú¾ä±ú£¿£¡
                    m_picPreviewPicture.Tag = SetParent(m_picPreviewPicture.hWnd, hWndParent)
                    ' ÒÆ¶¯Í¼Æ¬¿ò£¬Right Î»ÖÃ¹Ì¶¨£¬¸ß¶È¹Ì¶¨£¡
                    If m_ppLogoPosition = ppRight Then
                        picLeft = rcDlg.Right - rcDlg.Left - 5 - m_picLogoPicture.Width \ stpX: picTop = 0
                    Else
                        picLeft = rcDlg.Right - rcDlg.Left - 8: picTop = 0
                    End If
                    MoveWindow m_picPreviewPicture.hWnd, picLeft, ptL.y, _
                        m_picPreviewPicture.Width \ stpX - 3, rcList0.Bottom - rcList0.Top, True
                    ' ¼ÓÔØÍ¼Æ¬
                    'm_picPreviewPicture.PaintPicture m_picPreviewPicture.Picture, 0, 0, m_picPreviewPicture.ScaleWidth, m_picPreviewPicture.ScaleHeight
                    myPaintPicture m_picPreviewPicture, False
                    ShowWindow m_picPreviewPicture.hWnd, SW_SHOW ' ÏÔÊ¾Í¼Æ¬¿ò

                Case ppTop ' Ô¤ÀÀÍ¼Æ¬¿ò ÔÚ¶¥¶Ë£¬ÒªÒÆ¶¯¶Ô»°¿òÉÏÔ­À´µÄ¿Ø¼þ£¡
                    ' ÖØÐÂÈ¡µÃ LIST¿Ø¼þ¾ØÐÎ£¬¿ÉÄÜÔÚÉÏÃæ±»ÒÆ¶¯ÁË£¡
                    GetWindowRect hWndControl, rcList0
                    ptL.X = rcList0.Left: ptL.y = rcList0.Top
                    ScreenToClient hWndParent, ptL ' ptL¾­¹ý×ª»¯ºó²ÅÄÜµÃµ½ÏëÒªµÄ½á¹û£¡
                    
                    ' ¶Ô»°¿òÉÏËùÓÐÔ­Ê¼¿Ø¼þÏÂÒÆ
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
                    CDMoveOriginControl hWndParent, ID_FileTypeComboC, , OffSetY ' ÐÂÍâ¹Û£¡ÒÆ¶¯¶Ô»°¿òÉÏÎÄ¼þÃûÎÄ±¾¿ò£¬
                    CDMoveOriginControl hWndParent, ID_FileNameText, , OffSetY
                    
                    ' ÒÆ¶¯¹¤¾ßÀ¸£¡¹¤¾ßÀ¸£¨ÌØ±ð×¢Òâ£ºÎÞ·¨Í¨¹ý CDMoveOriginControl º¯ÊýÒÆ¶¯£¡£©
                    hToolBar = CDGetToolBarHandle(hWndParent)
                    GetWindowRect hToolBar, rcTB
                    pt.X = rcTB.Left
                    pt.y = rcTB.Top
                    ScreenToClient hWndParent, pt
                    MoveWindow hToolBar, pt.X, pt.y + OffSetY, rcTB.Right - rcTB.Left, rcTB.Bottom - rcTB.Top, True

                    ' ¸Ä±ä¶Ô»°¿ò´óÐ¡£¬£¡
                    GetWindowRect hWndParent, rcDlg ' È¡µÃ¶Ô»°¿ò¾ØÐÎ£¬ÔÙÒÆ¶¯£¨Êµ¼ÊÖ»¸Ä±ä¸ß¶È£¡£©
                    MoveWindow hWndParent, rcDlg.Left, rcDlg.Top, rcDlg.Right - rcDlg.Left, rcDlg.Bottom - rcDlg.Top + OffSetY, True
                    ' ÉèÖÃ³ÌÐò±êÖ¾Í¼Æ¬
                    ' ÉèÖÃÐÂµÄ£¬²¢±£´æÍ¼Æ¬¿òÔ­À´µÄ¸¸´°¿Ú¾ä±ú£¿£¡
                    m_picPreviewPicture.Tag = SetParent(m_picPreviewPicture.hWnd, hWndParent)
                    ' ÒÆ¶¯Í¼Æ¬¿ò£¬Left Î»ÖÃ¹Ì¶¨£¬¿í¶È¹Ì¶¨£¡
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
                    ' ¼ÓÔØÍ¼Æ¬
                    'm_picPreviewPicture.PaintPicture m_picPreviewPicture.Picture, 0, 0, m_picPreviewPicture.ScaleWidth, m_picPreviewPicture.ScaleHeight
                    myPaintPicture m_picPreviewPicture, False
                    ShowWindow m_picPreviewPicture.hWnd, SW_SHOW ' ÏÔÊ¾Í¼Æ¬¿ò
                    
                Case ppBottom
                    ' ÖØÐÂÈ¡µÃ LIST¿Ø¼þ¾ØÐÎ£¬¿ÉÄÜÔÚÉÏÃæ±»ÒÆ¶¯ÁË£¡
                    GetWindowRect hWndControl, rcList0
                    ptL.X = rcList0.Left: ptL.y = rcList0.Top
                    ScreenToClient hWndParent, ptL ' ptL¾­¹ý×ª»¯ºó²ÅÄÜµÃµ½ÏëÒªµÄ½á¹û£¡
                    
                    ' ¸Ä±ä¶Ô»°¿ò´óÐ¡£¬£¡
                    GetWindowRect hWndParent, rcDlg ' È¡µÃ¶Ô»°¿ò¾ØÐÎ£¬ÔÙÒÆ¶¯£¨Êµ¼ÊÖ»¸Ä±ä¸ß¶È£¡£©
                    MoveWindow hWndParent, rcDlg.Left, rcDlg.Top, rcDlg.Right - rcDlg.Left, rcDlg.Bottom - rcDlg.Top + OffSetY, True
                    ' ÉèÖÃ Ô¤ÀÀÍ¼Æ¬¿ò
                    Call GetClientRect(hWndParent, ClientRect) ' ÓÃ rcDlg.Bottom ²»ÐÐ£¡£¡£¡
                    ' ÉèÖÃÐÂµÄ£¬²¢±£´æÍ¼Æ¬¿òÔ­À´µÄ¸¸´°¿Ú¾ä±ú£¿£¡
                    m_picPreviewPicture.Tag = SetParent(m_picPreviewPicture.hWnd, hWndParent)
                    ' ÒÆ¶¯Í¼Æ¬¿ò£¬Left Î»ÖÃ¹Ì¶¨£¬¿í¶È¹Ì¶¨£¡
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
                    ' ¼ÓÔØÍ¼Æ¬
                    'm_picPreviewPicture.PaintPicture m_picPreviewPicture.Picture, 0, 0, m_picPreviewPicture.ScaleWidth, m_picPreviewPicture.ScaleHeight
                    myPaintPicture m_picPreviewPicture, False
                    ShowWindow m_picPreviewPicture.hWnd, SW_SHOW ' ÏÔÊ¾Í¼Æ¬¿ò

            End Select
NoPreview:
            ' ÉèÖÃ¶Ô»°¿òÆô¶¯Î»ÖÃ£¿Ö»ÅÐ¶ÏÆÁÄ»ÖÐÐÄºÍËùÓÐÕßÖÐÐÄ£¬ÆäËû²»¹Ü£¡
            GetWindowRect hWndParent, rcDlg ' È¡µÃ¶Ô»°¿ò¾ØÐÎ
            W = rcDlg.Right - rcDlg.Left: H = rcDlg.Bottom - rcDlg.Top ' ¶Ô»°¿ò¿í¡¢¸ß
            If m_dlgStartUpPosition = vbStartUpScreen Then ' ÔÙÒÆ¶¯¡£ÆÁÄ»ÖÐÐÄ
                MoveWindow hWndParent, (Screen.Width \ stpX - W) \ 2, (Screen.Height \ stpY - H) \ 2, W, H, True
            ElseIf m_dlgStartUpPosition = vbStartUpOwner Then ' ËùÓÐÕßÖÐÐÄ
                Dim rcOwner As RECT, T As Long ' È¡µÃ¶Ô»°¿òµÄ¸¸´°¿Ú¾ØÐÎ
                GetWindowRect GetParent(hWndParent), rcOwner
                T = rcOwner.Top + (rcOwner.Bottom - rcOwner.Top - H) \ 2
                If T < 0 Then T = 0 ' ±£Ö¤¶Ô»°¿ò²»³¬¹ýÆÁÄ»¶¥¶Ë£¡
                MoveWindow hWndParent, rcOwner.Left + (rcOwner.Right - rcOwner.Left - W) \ 2, T, W, H, True
            End If
            ' Æô¶¯Î»ÖÃ£¬ÒòÎª¼¸ÖÖ¶Ô»°¿ò¶¼ÓÐÕâ¸öÊôÐÔ£¡¸ÄÓÃÒ»¸öº¯ÊýÊµÏÖ£¬Ð§¹û²»ºÃ£¡²»ÓÃÁË£¡
            'setDlgStartUpPosition hWndParent, GetParent(hWndParent)
            
            ' ´´½¨Ô¤ÀÀÎÄ±¾¿ò£¬Æä´óÐ¡ºÍÎ»ÖÃÓëÔ¤ÀÀÍ¼Æ¬¿òÒ»Ñù
            Dim rcP As RECT ' ¾ö¶¨Ô¤ÀÀÎÄ±¾¿òÎ»ÖÃºÍ´óÐ¡¡£
            GetWindowRect m_picPreviewPicture.hWnd, rcP ' Ã»ÉèÖÃ m_picPreviewPicture £¬Õâ³ö´í£º(¶ÔÏó±äÁ¿»ò With ¿é±äÁ¿Î´ÉèÖÃ)
            ptL.X = rcP.Left: ptL.y = rcP.Top
            ScreenToClient hWndParent, ptL ' ptL¾­¹ý×ª»¯ºó²ÅÄÜµÃµ½ÏëÒªµÄ½á¹û£¡
            ' ¿ªÊ¼´´½¨Ô¤ÀÀÎÄ±¾¿ò È¥µô Or WS_VISIBLE £¬²»ÔÚ´´½¨Ê±ÏÔÊ¾£¡
            hWndTextView = CreateWindowEx(0, _
                "Edit", App.LegalCopyright, _
                WS_BORDER Or WS_CHILD Or WS_HSCROLL Or WS_VSCROLL Or ES_AUTOHSCROLL Or ES_AUTOHSCROLL Or ES_MULTILINE, _
                ptL.X, ptL.y, _
                rcP.Right - rcP.Left, rcP.Bottom - rcP.Top, _
                hWndParent, 0&, App.hInstance, 0&)
            ' ´´½¨ÐÂµÄ×ÖÌå Fixedsys - Times New Roman - MS Sans Serif
            Dim NewFont As Long
            NewFont = CreateFont(18, 0, 0, 0, _
                      366, False, False, False, _
                      DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_LH_ANGLES, _
                      ANTIALIASED_QUALITY Or PROOF_QUALITY, DEFAULT_PITCH Or FF_DONTCARE, _
                       "Times New Roman")
            ' ÉèÖÃÎÄ±¾¿ò×ÖÌå
            SendMessage hWndTextView, WM_SETFONT, NewFont, 0
            ' Òþ²Ø¡¢ÏÔÊ¾¶Ô»°¿òÉÏµÄ¿Ø¼þ m_blnHideControls(I) µÄÖµ¾ö¶¨ÊÇ²»ÊÇÒªÒþ²Ø£¡
            Call HideOrShowDlgControls(hWndParent)
            ' ÉèÖÃ¶Ô»°¿òÉÏµÄ¿Ø¼þµÄÎÄ×Ö£¬m_strControlsCaption(I) ¾ö¶¨ÆäÖµ£¡
            Call mSetDlgControlsCaption(hWndParent)
            
        Case CDN_SelChange ' ÎÄ¼þÑ¡Ôñ¸Ä±äÊ±£¬½øÐÐÔ¤ÀÀ£¡
            strSelFile = SendMsgGetStr(hdr.hwndFrom, CDM_GetFilePath) ' ¼ÇÂ¼Ñ¡ÖÐµÄÎÄ¼þÂ·¾¶
            Debug.Print "SelChange:"; strSelFile
            'Screen.MousePointer = vbHourglass ' Êó±ê³ÊÉ³Â©×´
            If Not m_ppPreviewPosition = ppNone Then LoadPreview strSelFile, hWndParent  ' µ÷ÓÃº¯Êý£¬¼ÓÔØÔ¤ÀÀ¡£
            'Screen.MousePointer = vbDefault ' Íê³ÉÔ¤ÀÀ£¬Êó±ê»Ö¸´
        Case CDN_FolderChange
            Debug.Print "FolderChange:"; SendMsgGetStr(hdr.hwndFrom, CDM_GetFolderPath)
        Case CDN_ShareViolation
            Debug.Print "ShareViolation"
        Case CDN_Help
            Debug.Print "Help"
            If WhichStyle = ssOpen Then
                MsgBox "°ïÖú£º´ò¿ª¶Ô»°¿ò £¡", vbInformation
            ElseIf WhichStyle = ssSave Then
                MsgBox "°ïÖú£º±£´æ¶Ô»°¿ò £¡", vbInformation
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
    ' ¼ÓÔØ' ³ÌÐò±êÖ¾Í¼Æ¬£¬·Ç·ÅÔÚÕâÀï²»¿É£¡£¡£¡·ñÔòÐ§¹û²»¶Ô£¡£¡£¡
    myPaintPicture m_picLogoPicture
End Function

' ÒÆ¶¯ ´ò¿ª/±£´æ ¶Ô»°¿òÉÏÔ­ÓÐµÄ¿Ø¼þ' ÈôÒªÒÆ¶¯¶Ô»°¿òÉÏÔ­À´Ã»ÓÐµÄ¿Ø¼þ£¬ÐèÁíÍâ´¦Àí£¡
' ºóÃæµÄ¿ÉÑ¡²ÎÊýÉèÖÃÎªÄ¬ÈÏ -1 Ê±£¬ÒÆ¶¯Ê±²»²Ù×÷£¡
Private Sub CDMoveOriginControl(ByVal hWndCD As Long, ByVal ID As Long, _
    Optional ByVal X As Long = -1, Optional ByVal y As Long = -1, _
    Optional ByVal nWidth As Long = -1, Optional ByVal nHeight As Long = -1)
    
    Dim hWndControl As Long ' ¿Ø¼þ¾ä±ú
    Dim rectControl As RECT ' ¿Ø¼þ¾ØÐÎ
    Dim ptCtr As POINTAPI   ' ¿Ø¼þ×óÉÏ½ÇÔÚÆÁÄ»µÄ×ø±êÖµ
    Dim ptDlg As POINTAPI   ' ¶Ô»°¿ò×óÉÏ½ÇÔÚÆÁÄ»µÄ×ø±êÖµ
    
    ' == ¸ù¾ÝIDÈ¡µÃ¿Ø¼þ¾ä±ú
    hWndControl = GetDlgItem(hWndCD, ID)
    
    ' == È¡µÃ¿Ø¼þ¾ØÐÎ£¨Î»ÖÃÊÇÏà¶Ô¶Ô»°¿ò¶øÑÔ£¬¼´ÒÔ¶Ô»°¿ò×óÉÏ½ÇÄÇµãÎª0µã£©²¢ÉèÖÃ¿Ø¼þ´óÐ¡
    GetWindowRect hWndControl, rectControl

    ' ==È¡µÃ¶Ô»°¿òÎ»ÖÃ
    ScreenToClient hWndCD, ptDlg
    ' ==È¡µÃ²¢ÉèÖÃ¿Ø¼þÎ»ÖÃ
    ScreenToClient hWndControl, ptCtr
    ptCtr.X = rectControl.Left + ptDlg.X + IIf(X <> -1, X, 0)
    ptCtr.y = rectControl.Top + ptDlg.y + IIf(y <> -1, y, 0)
    X = ptCtr.X
    y = ptCtr.y
    nWidth = rectControl.Right - rectControl.Left + IIf(nWidth <> -1, nWidth, 0)
    nHeight = rectControl.Bottom - rectControl.Top + IIf(nHeight <> -1, nHeight, 0)
    
    ' µ÷ÓÃAPIÒÆ¶¯¿Ø¼þ£¡X£¬YÎªÏà¶ÔÆÁÄ»×óÉÏ½ÇµÄ×ø±êÖµ£¡
    MoveWindow hWndControl, X, y, nWidth, nHeight, True
End Sub

' ÒÆ¶¯¶Ô»°¿òÉÏ¹¤¾ßÀ¸Ê±£¬È¡µÃ¹¤¾ßÀ¸¾ä±ú£¡
Private Function CDGetToolBarHandle(ByVal hDialog As Long) As Long
    CDGetToolBarHandle = FindWindowEx(hDialog, 0, "ToolBarWindow32", vbNullString)
End Function

' ¶Ô»°¿ò¸Ä±äÄ¿Â¼¡¢Ñ¡ÔñÎÄ¼þÊ±£¬È¡µÃ Â·¾¶¡£
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
' ### ÒÔÏÂÎªÔ¤ÀÀÏà¹ØµÄÖØÒªº¯Êý ##################################################
' Ô¤ÀÀÍ¼Æ¬¿òÍ¼Æ¬µÄÏÔÊ¾£¬ÓÃ PaintPicture ·½·¨ÔÚÍ¼Æ¬¿òÉÏ»­Í¼Æ¬£¿
Private Sub myPaintPicture(picBox As PictureBox, Optional blnShowPictureOrNone As Boolean = True, Optional stdPic As StdPicture = Nothing)
    On Error Resume Next
    Dim tempPic As StdPicture
    If stdPic Is Nothing Then
        Set tempPic = picBox.Picture
    Else
        Set tempPic = stdPic
    End If
    picBox.AutoRedraw = True ' ÈÃÍ¼Æ¬¿ò×Ô¶¯Ë¢ÐÂ£¡
    If blnShowPictureOrNone Then
        picBox.PaintPicture tempPic, 0, 0, picBox.ScaleWidth, picBox.ScaleHeight
    Else ' Òþ²ØÍ¼Æ¬£¡
        Set picBox.Picture = Nothing
        ' ÔÚÍ¼Æ¬¿òÖÐÐÄÏÔÊ¾ÎÄ×Ö£¿ÓÐ±ØÒª£¿£¡ =====================================
        Dim s As String: s = App.LegalCopyright
        picBox.ForeColor = vbBlack: picBox.FontSize = 18
        picBox.CurrentX = (picBox.ScaleWidth - picBox.TextWidth(s)) / 2
        picBox.CurrentY = (picBox.ScaleHeight - picBox.TextHeight(s)) / 2
        picBox.Print s ' =======================================================
    End If
End Sub

' ÅÐ¶ÏÎÄ¼þÀàÐÍ£¬ÒÔ¾ö¶¨ÓÃÊ²Ã´·½Ê½Ô¤ÀÀ£¡£¡£¡
Private Function getFileType(strFileName As String) As FileType
    Dim strExt As String    ' ÎÄ¼þºó×ºÃû Èç TXT
    ' È¡µÃÎÄ¼þºó×ºÃû£¬²¢×ª»¯Îª´óÐ´£¡
    strExt = UCase$(Right$(strFileName, Len(strFileName) - InStrRev(strFileName, ".")))
    getFileType = ffText ' ÉèÖÃº¯Êý·µ»ØÄ¬ÈÏÀàÐÍ£¡
    ' ÅÐ¶ÏÎÄ¼þºó×ºÃû£¬VB Í¼Æ¬¿ò²»ÄÜ¼ÓÔØ PNG Í¼Æ¬ ANI ¶¯»­¹â±ê£¡JPE\JFIF\TIF\TIFF Î´Öª£¡
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

' VB È¡µÃÍ¼Æ¬´óÐ¡
Private Function fncGetPicInfo(lsPicName As String, Optional hBitmap As Long = 0) As PicInfo
    Dim res As Long
    Dim bmp As BITMAP
    If hBitmap = 0 Then hBitmap = LoadPicture(lsPicName).Handle ' º¯ÊýÒªÔÚ´°ÌåÖÐ£¬LoadPicture ²ÅÓÐÐ§£¡
    res = GetObject(hBitmap, Len(bmp), bmp) 'È¡µÃBITMAPµÄ½á¹¹
    fncGetPicInfo.picWidth = bmp.bmWidth
    fncGetPicInfo.picHeight = bmp.bmHeight
End Function

' ¼ÓÔØÔ¤ÀÀ£¬Í¼Æ¬£¿ÎÄ±¾ÎÄ¼þ£¿Wave ÎÄ¼þ£¬ÒôÆµÎÄ¼þ£¿
Private Sub LoadPreview(strFileName As String, ByVal hWndParent As Long, Optional ShowFileSize As Long = &H1000)
    If FileLen(strFileName) = 0 Then Exit Sub
    On Error GoTo ErrLoad
    Dim I As Integer
    ' ¸ù¾ÝÎÄ¼þÀàÐÍ£¬½øÐÐ²»Í¬µÄÔ¤ÀÀ²Ù×÷¡£
    If getFileType(strFileName) = ffText Then ' ÎÄ±¾ÎÄ¼þ
        Dim FileNum      As Integer
        Dim FileSize     As Long
        Dim LoadBytes()  As Byte
        ' ¶ÁÈ¡ÎÄ±¾ÎÄ¼þ
        FileNum = FreeFile
        Open strFileName For Binary Access Read Lock Write As FileNum
            FileSize = LOF(FileNum)
            If FileSize > ShowFileSize Then FileSize = ShowFileSize ' Ô¤ÀÀÎÄ¼þµÄ´óÐ¡£¬¿ÉÄÜ²»ÏÔÊ¾ËùÓÐÄÚÈÝ¡£
            ReDim LoadBytes(0 To FileSize - 1)
            Get #FileNum, 1, LoadBytes
        Close FileNum
        ' ÔÚÎÄ±¾¿òÏÔÊ¾ÎÄ×Ö
        Debug.Print SetTextColor(GetDC(hWndTextView), vbRed)
        Call SetWindowText(hWndTextView, LoadBytes(0))  ' ÉèÖÃÎÄ×Ö£¨ÎÄ¼þ±¾ÉíµÄÎÄ×Ö£©
        Dim s As String: s = String(FileSize, 0)
        GetWindowText hWndTextView, s, FileSize
        Call SetWindowText(hWndTextView, ByVal Replace$(s, Chr$(0), "") _
            & vbCrLf & "£¨º×ÍûÀ¼¡¤Á÷ Ê¡ÂÔ " & Format$(FileLen(strFileName) - FileSize, "###,###,###,##0") & " ×Ö½Ú£©") ' ÉèÖÃÎÄ×Ö£¨+º×ÍûÀ¼¡¤Á÷ ±êÖ¾£©
        ShowWindow hWndTextView, SW_SHOW                ' ÎÄ±¾¿ò¿É¼û
        ShowWindow m_picPreviewPicture.hWnd, SW_HIDE    ' Òþ²ØÔ¤ÀÀÍ¼Æ¬¿ò
        For I = 0 To 2
            ShowWindow hWndButtonPlay(I), SW_HIDE
        Next I
        Exit Sub
    ElseIf getFileType(strFileName) = ffPicture Then ' Í¼Æ¬ÎÄ¼þ£¨Í¼Æ¬Ô¤ÀÀ°´±ÈÀýÏÔÊ¾£¬²¢ÏÔÊ¾Í¼Æ¬¿íx¸ßºÍÔ¤ÀÀ±ÈÀý¡££©
        Dim W As Long, H As Long, fWH As PicInfo ' ¼ÓÔØµÄÍ¼Æ¬¿íx¸ß
        Dim W0 As Long, H0 As Long, oldSM As ScaleModeConstants ' Í¼Æ¬¿ò´óÐ¡£¡
        Dim X As Long, y As Long, W1 As Long, H1 As Long, per As Integer, sP As StdPicture ' °´±ÈÀýÖØÐÂÏÔÊ¾µÄÍ¼Æ¬Î»ÖÃ£¬´óÐ¡£¡
        oldSM = m_picPreviewPicture.ScaleMode: m_picPreviewPicture.ScaleMode = vbPixels ' ÉèÖÃÍ¼Æ¬¿òScaleModeÎªÏñËØ£¡
        W0 = m_picPreviewPicture.ScaleWidth: H0 = m_picPreviewPicture.ScaleHeight
        ' ¼ÓÔØÍ¼Æ¬µ½Í¼Æ¬¿ò£¬Ö»ÎªÈ¡µÃÆä³ß´ç£¡¡£
        Set sP = LoadPicture(strFileName)
        Set m_picPreviewPicture.Picture = sP ' LoadPicture(strFileName)
        fWH = fncGetPicInfo(strFileName, m_picPreviewPicture.Picture)
        W = fWH.picWidth: H = fWH.picHeight
        ' ÅÐ¶ÏÍ¼Æ¬¿ò¿í¸ß±È¡¢¼ÓÔØµÄÍ¼Æ¬¿í¸ß±È£¬Æä±ÈÖµÔÚÏà±È¡£¡£¡£
        If (W0 / H0) / (W / H) < 1 Then ' ÒÔ' Í¼Æ¬¿ò¿í¶ÈÎª»ù×¼
            W1 = W0: H1 = W0 / W * H
            X = 0: y = (H0 - H1) / 2    ' µ÷ÕûÏÔÊ¾Î»ÖÃ£¬Ê¹Æä¾ÓÖÐ£¡
            per = W0 / W * 100
        Else                            ' ÒÔ' Í¼Æ¬¿ò¸ß¶ÈÎª»ù×¼
            W1 = H0 / H * W: H1 = H0
            X = (W0 - W1) / 2: y = 0    ' µ÷ÕûÏÔÊ¾Î»ÖÃ£¬Ê¹Æä¾ÓÖÐ£¡
            per = H0 / H * 100
        End If
        'Debug.Print "³ß´ç: " & W & " x " & H & " Í¼Æ¬¿ò: " & W0 & " x " & H0
        Set m_picPreviewPicture.Picture = Nothing
        '' ÔÙ´Î¼ÓÔØÍ¼Æ¬µ½Í¼Æ¬¿ò£¬ÒÔÏÔÊ¾¡£¡£¡£²»ÔÙ¼ÓÔØ£¬·ñÔò»áÏÔÊ¾Á½¸öÍ¼£¨PaintPicture ·½·¨Ò²»áÏÔÊ¾Í¼£¡£©¡£¡£¡£
        'Set m_picPreviewPicture.Picture = sP ' LoadPicture(strFileName)
        'myPaintPicture m_picPreviewPicture
        m_picPreviewPicture.PaintPicture sP, X, y, W1, H1
        ' Í¼Æ¬¿òÉÏÏÔÊ¾ÎÄ×Ö
        Dim sT As String: sT = "³ß´ç: " & W & " x " & H & " (" & per & "%)"
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
        m_picPreviewPicture.ScaleMode = oldSM ' »¹Ô­Í¼Æ¬¿ò¾ÉµÄ ScaleMode Öµ¡£
        Exit Sub
    ElseIf getFileType(strFileName) = ffWave Then ' ²¨ÐÎÎÄ¼þ
        Create3Buttons hWndParent ' ´´½¨ 3 ¸ö²¥·Å¿ØÖÆ°´Å¥¡£
        ShowWindow m_picPreviewPicture.hWnd, SW_SHOW ' ÏÔÊ¾Ô¤ÀÀÍ¼Æ¬¿ò
        Set m_picPreviewPicture.Picture = Nothing
        ShowWindow hWndTextView, SW_HIDE             ' Òþ²ØÎÄ±¾¿ò
        
        ' ÏÈÍ£Ö¹ÉùÒô£¬ÒòÎªÇ°Ãæ¿ÉÄÜ²¥·Å¹ý£¡£¡£¡£¡' ×¢Òâ£ºÕâÀï²»Ö±½Óµ÷ÓÃ B3Button_Click 2 £¡
        PlayAudio strSelFile, 2
        sndPlaySound vbNullString, 0 ' Í£Ö¹ Wave ÎÄ¼þ²¥·Å¡£
        ' =========================================================================================
        ' ==== »­³öWaveÎÄ¼þ²¨ÐÎ£¨¿ÉÈ¥µô£©==========================================================
        ' =========================================================================================
        ' »­³ö²¨ÐÎ£¡ÓÃÒ»¸öÄ£¿éÍê³É£¬¿ÉÒÔÈ¥µô´Ë¹¦ÄÜ£¡£¡£¡£¡£¡£¡£¡£¡
        MDrawWaves.DrawWaves strFileName, m_picPreviewPicture  ' ÎªÉ¶²»ÐÐ£¿Ô­Òò£ºÒ»¶¨ÒªÉèÖÃ ScaleMode = vbTwips £¡£¡£¡
        ' =========================================================================================
        ' ==== »­³öWaveÎÄ¼þ²¨ÐÎ£¨¿ÉÈ¥µô£©==========================================================
        ' =========================================================================================
        Exit Sub
    ElseIf getFileType(strFileName) = ffAudio Then ' ÒôÆµÎÄ¼þ
        Create3Buttons hWndParent ' ´´½¨ 3 ¸ö²¥·Å¿ØÖÆ°´Å¥¡£
        ShowWindow m_picPreviewPicture.hWnd, SW_SHOW ' ÏÔÊ¾Ô¤ÀÀÍ¼Æ¬¿ò
        Set m_picPreviewPicture.Picture = Nothing
        ShowWindow hWndTextView, SW_HIDE             ' Òþ²ØÎÄ±¾¿ò
        
        ' ÏÈÍ£Ö¹ÉùÒô£¬ÒòÎªÇ°Ãæ¿ÉÄÜ²¥·Å¹ý£¡£¡£¡£¡' ×¢Òâ£ºÕâÀï²»Ö±½Óµ÷ÓÃ B3Button_Click 2 £¡
        PlayAudio strSelFile, 2
        sndPlaySound vbNullString, 0 ' Í£Ö¹ Wave ÎÄ¼þ²¥·Å¡£
        Exit Sub
    End If

ErrLoad:
    Debug.Print "LoadPreview Error " & Err.Number & ": " & Err.Description
End Sub
' ´´½¨²¥·Å¿ØÖÆ 3 ¸ö°´Å¥
Private Sub Create3Buttons(ByVal hWndParent As Long)
    
    ' ²»ÄÜÖØ¸´´´½¨£¬·ñÔò£¬ÏìÓ¦ÏûÏ¢Ê±£¬hWndButtonPlay(0)¡£¡£¡£Öµ±äÁË£¬¶ø lParam µÄÖµ»¹ÊÇµÚÒ»´Î´´½¨Ê±µÄ£¬µ¼ÖÂÎÞ·¨ÏìÓ¦ÏûÏ¢£¡
    If hWndButtonPlay(0) Then ' ÒªÏÔÊ¾°´Å¥£¡·ñÔò²»¼ûÁË£¿£¡
        ShowWindow hWndButtonPlay(0), SW_SHOW
        ShowWindow hWndButtonPlay(1), SW_SHOW
        ShowWindow hWndButtonPlay(2), SW_SHOW
        Exit Sub
    End If
    'MsgBox "´´½¨²¥·Å¿ØÖÆ 3 ¸ö°´Å¥"
    ' ´´½¨²¥·Å¿ØÖÆ 3 ¸ö°´Å¥£¬ÆäÎ»ÖÃÓÉÎÄ¼þÀàÐÍ±êÇ©Î»ÖÃ¾ö¶¨¡£´óÐ¡¹Ì¶¨¡£
    Dim rcP As RECT, ptL As POINTAPI ' ¾ö¶¨²¥·Å¿ØÖÆ 3 ¸ö°´Å¥Î»ÖÃºÍ´óÐ¡¡£
    Dim NewFont As Long
    GetWindowRect GetDlgItem(hWndParent, ID_FileTypeLabel), rcP
    ptL.X = rcP.Left: ptL.y = rcP.Top
    ScreenToClient hWndParent, ptL ' ptL¾­¹ý×ª»¯ºó²ÅÄÜµÃµ½ÏëÒªµÄ½á¹û£¡
    ' ¿ªÊ¼´´½¨ ²¥·Å¿ØÖÆ 3 ¸ö°´Å¥ ¼Ó Or WS_VISIBLE £¬ÔÚ´´½¨Ê±ÏÔÊ¾£¡
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
    ' ´´½¨ÐÂµÄ×ÖÌå Webdings - Fixedsys - Times New Roman - MS Sans Serif
    NewFont = CreateFont(18, 0, 0, 0, _
              366, False, False, False, _
              DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_LH_ANGLES, _
              ANTIALIASED_QUALITY Or PROOF_QUALITY, DEFAULT_PITCH Or FF_DONTCARE, _
               "Webdings")
    ' ÉèÖÃ×ÖÌå
    SendMessage hWndButtonPlay(0), WM_SETFONT, NewFont, 0
    SendMessage hWndButtonPlay(1), WM_SETFONT, NewFont, 0
    SendMessage hWndButtonPlay(2), WM_SETFONT, NewFont, 0
End Sub

' ×¢Òâ£º½ö¶ÔÑÕÉ«¶Ô»°¿ò¡£¡£¡£ÉèÖÃ¶Ô»°¿òÆô¶¯Î»ÖÃ£¿Ö»ÅÐ¶ÏÆÁÄ»ÖÐÐÄºÍËùÓÐÕßÖÐÐÄ£¬ÆäËû²»¹Ü£¡
Private Sub setDlgStartUpPosition(ByVal hWndDlg As Long, ByVal hWndParent As Long)
    Dim W As Long, H As Long, T As Long
    Dim rcDlg As RECT, rcOwner As RECT
    GetWindowRect hWndDlg, rcDlg ' È¡µÃ¶Ô»°¿ò¾ØÐÎ
    W = rcDlg.Right - rcDlg.Left: H = rcDlg.Bottom - rcDlg.Top ' ¶Ô»°¿ò¿í¡¢¸ß
    If m_dlgStartUpPosition = vbStartUpScreen Then ' ÔÙÒÆ¶¯¡£ÆÁÄ»ÖÐÐÄ
        MoveWindow hWndDlg, (Screen.Width \ Screen.TwipsPerPixelX - W) \ 2, (Screen.Height \ Screen.TwipsPerPixelY - H) \ 2, W, H, True
    ElseIf m_dlgStartUpPosition = vbStartUpOwner Then ' ËùÓÐÕßÖÐÐÄ
         ' È¡µÃ¶Ô»°¿òµÄ¸¸´°¿Ú¾ØÐÎ
        GetWindowRect hWndParent, rcOwner
        T = rcOwner.Top + (rcOwner.Bottom - rcOwner.Top - H) \ 2
        If T < 0 Then T = 0 ' ±£Ö¤¶Ô»°¿ò²»³¬¹ýÆÁÄ»¶¥¶Ë£¡
        MoveWindow hWndDlg, rcOwner.Left + (rcOwner.Right - rcOwner.Left - W) \ 2, T, W, H, True
    End If
End Sub


' =========================================================================================
' ==== ×Ô¶¨Òå¶Ô»°¿òÉÏµÄ¿Ø¼þÒþ²Ø¡¢ÏÔÊ¾£¬¸Ä±äÎÄ×ÖµÈ£¡£¨¿ÉÈ¥µô£©==============================
' =========================================================================================
' Òþ²Ø¡¢ÏÔÊ¾¶Ô»°¿òÉÏµÄ¿Ø¼þ m_blnHideControls(I) µÄÖµ¾ö¶¨ÊÇ²»ÊÇÒªÒþ²Ø£¡¶øÇÒ£¬Flags ÊôÐÔÒ²ÓÐÓ°Ïì¡£
Private Sub HideOrShowDlgControls(ByVal hWndDlg As Long)
    Dim I As Integer, ctrl_ID As Variant ' ¿Ø¼þID£¬×ªµ½Ò»¸öÊý×é£¬ºÃ²Ù×÷£¡
    ' (0 to 12) ·Ö±ðÎª£ºÓÐÐ©²»ÄÜÒþ²Ø¡£ÓÃ¡¶¡·±ê³ö¡£Ö»Ê£ÏÂ (0 to 8)
    ' ¡°²éÕÒ·¶Î§(&I)¡±±êÇ©      -- Ä¿Â¼ÏÂÀ­¿ò                         -- ¡¶¹¤¾ßÀ¸¡·        1
    ' ¡¶¿ì½ÝÄ¿Â¼Çø£¨°æ±¾>=Win2K£©¡· -- ¡¶ÁÐ±í¿ò£¨ÁÐ³öÎÄ¼þµÄ×î´óÇøÓò£©¡·                    2
    ' ¡°ÎÄ¼þÃû(&N)¡±±êÇ©        -- ¡¶¡°ÎÄ¼þÃû(&N)¡±ÎÄ±¾¿ò¡·           -- ¡°È·¶¨(&O)¡±°´¼ü  1
    ' ¡°ÎÄ¼þÀàÐÍ(&T)¡±±êÇ©      -- ¡°ÎÄ¼þÀàÐÍ(&T)¡±ÏÂÀ­¿ò£¨ÐÂÍâ¹ÛÊ±£© -- ¡°È¡Ïû(&C)¡±°´¼ü  0
    ' ¡°Ö»¶Á¡±¶àÑ¡¿ò            -- ¡°°ïÖú(&H)¡±°´¼ü                                        0
    ctrl_ID = Array(ID_FolderLabel, ID_FolderCombo, _
                ID_FileNameLable, ID_OK, _
                ID_FileTypeLabel, ID_FileTypeCombo0, ID_Cancel, _
                ID_ReadOnly, ID_Help)
    ' Òþ²Ø¿Ø¼þ
    For I = 0 To 8
        If m_blnHideControls(I) Then Call SendMessage(hWndDlg, CDM_HideControl, ctrl_ID(I), ByVal 0&)
    Next I
    Set ctrl_ID = Nothing
End Sub
' ÉèÖÃ¶Ô»°¿òÉÏµÄ¿Ø¼þµÄÎÄ×Ö¡£m_strControlsCaption(I) ¾ö¶¨ÆäÖµ£¬Ä¬ÈÏÎª¿Õ£¬²»¸Ä±äÔ­Ê¼Öµ£¡
Private Sub mSetDlgControlsCaption(ByVal hWndDlg As Long)
    Dim I As Integer, ctrl_ID As Variant ' ¿Ø¼þID£¬×ªµ½Ò»¸öÊý×é£¬ºÃ²Ù×÷£¡
    ' (0 to 12) ·Ö±ðÎª£ºÓÐÐ©²»ÄÜÉèÖÃÎÄ×Ö¡£ÓÃ¡¶¡·±ê³ö¡£Ö»Ê£ÏÂ (0 to 6)
    ' ¡°²éÕÒ·¶Î§(&I)¡±±êÇ©      -- ¡¶Ä¿Â¼ÏÂÀ­¿ò¡·                        -- ¡¶¹¤¾ßÀ¸¡·        2
    ' ¡¶¿ì½ÝÄ¿Â¼Çø£¨°æ±¾>=Win2K£©¡· -- ¡¶ÁÐ±í¿ò£¨ÁÐ³öÎÄ¼þµÄ×î´óÇøÓò£©¡·                       2
    ' ¡°ÎÄ¼þÃû(&N)¡±±êÇ©        -- ¡¶¡°ÎÄ¼þÃû(&N)¡±ÎÄ±¾¿ò¡·              -- ¡°È·¶¨(&O)¡±°´¼ü  1
    ' ¡°ÎÄ¼þÀàÐÍ(&T)¡±±êÇ©      -- ¡¶¡°ÎÄ¼þÀàÐÍ(&T)¡±ÏÂÀ­¿ò£¨ÐÂÍâ¹ÛÊ±£©¡·-- ¡°È¡Ïû(&C)¡±°´¼ü  1
    ' ¡°Ö»¶Á¡±¶àÑ¡¿ò            -- ¡°°ïÖú(&H)¡±°´¼ü                                           0
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
' ==== ×Ô¶¨Òå¶Ô»°¿òÉÏµÄ¿Ø¼þÒþ²Ø¡¢ÏÔÊ¾£¬¸Ä±äÎÄ×ÖµÈ£¡£¨¿ÉÈ¥µô£©==============================
' =========================================================================================



' =========================================================================================
' ==== ÉùÒôÎÄ¼þµÄ²¥·Å£¨¿ÉÈ¥µô£©============================================================
' =========================================================================================
' Èý¸ö²¥·Å°´Å¥µ¥»÷ÊÂ¼þ£¿
Private Sub B3Button_Click(Index As Integer)
'    PlayAudio strSelFile, Index' ¶¼ÓÃ mciSendString º¯Êý²¥·Å£¿£¡£¡£¡
    If Index = 0 Then ' ²¥·Å
        If getFileType(strSelFile) = ffWave Then ' Wave ÎÄ¼þ
            sndPlaySound strSelFile, SND_FILENAME Or SND_ASYNC
        Else
            PlayAudio strSelFile, 0
        End If
    ElseIf Index = 1 Then ' ÔÝÍ£
        If getFileType(strSelFile) = ffWave Then ' Wave ÎÄ¼þ
            sndPlaySound vbNullString, 0 ' ²»ÖªµÀÔõÃ´ÔÝÍ££¡ÕâÀïÍ£Ö¹£¡£¡£¡
        Else
            PlayAudio strSelFile, 1
        End If
    Else ' Í£Ö¹
        If getFileType(strSelFile) = ffWave Then ' Wave ÎÄ¼þ
            sndPlaySound vbNullString, 0
        Else
            PlayAudio strSelFile, 2
        End If
    End If
End Sub
' ¸ß¼¶Ã½Ìå²¥·Å£¬²¥·ÅÒôÆµÎÄ¼þ¡£·´Ó¦Ì«ÂýÁË£¿£¡£¡£¡Wave ÎÄ¼þ»¹ÊÇÓÃÆäËûº¯Êý²¨·Å£¿£¿£¿£¡£¡£¡
Private Sub PlayAudio(strFileName As String, Optional setStatus As Integer = 0)
    If Len(Dir$(strFileName)) = 0 Then Exit Sub
    Const ALIAS_NAME As String = "mySound"
    Dim rt As Long
    If setStatus = 0 Then ' ²¥·Å
        If getPalyStatus(ALIAS_NAME) = IsStopped Then
            ' ´ò¿ª²¢´ÓÍ·¿ªÊ¼²¥·Å¡£×¢Òâ£º¸ø strFileName ¼ÓË«ÒýºÅ£¬·ñÔòÓÐµÄÎÄ¼þÃûÓÐ¿Õ¸ñ£¬ÎÞ·¨²¥·Å£¡
            rt = mciSendString("open " & """" & strFileName & """" & " alias " & ALIAS_NAME, vbNullString, 0, 0)
            rt = mciSendString("play " & ALIAS_NAME, vbNullString, 0, 0)
            ' È¡µÃÃ½ÌåÎÄ¼þ³¤¶È£¿£¡
            Dim RefStr1 As String * 80
            mciSendString "status " & ALIAS_NAME & " length", RefStr1, Len(RefStr1), 0
            Debug.Print "×Ü³¤¶È£º" & Val(RefStr1)
        ElseIf getPalyStatus(ALIAS_NAME) = IsPaused Then
            ' ¼ÌÐø²¥·Å
            rt = mciSendString("resume " & ALIAS_NAME, vbNullString, 0, 0)
            ' »ñÈ¡µ±Ç°²¥·Å½ø¶È£¬Ïà¶ÔÎÄ¼þ³¤¶È¶øÑÔ£¿£¡
            Dim RefStr2 As String * 80
            mciSendString "status " & ALIAS_NAME & " position", RefStr2, Len(RefStr2), 0
            Debug.Print "ÒÑ²¥·Å£º" & Val(RefStr2)
        Else
            ' Í£Ö¹²¥·Å²¢¹Ø±ÕÉùÒô£¡
            rt = mciSendString("stop " & ALIAS_NAME, vbNullString, 0, 0)
            rt = mciSendString("close " & ALIAS_NAME, vbNullString, 0, 0)
        End If
    ElseIf setStatus = 1 Then ' ÔÝÍ£
        rt = mciSendString("pause " & ALIAS_NAME, vbNullString, 0, 0)
    Else ' Í£Ö¹
        rt = mciSendString("stop " & ALIAS_NAME, vbNullString, 0, 0)
        rt = mciSendString("close " & ALIAS_NAME, vbNullString, 0, 0)
    End If
End Sub
' »ñµÃµ±Ç°Ã½ÌåµÄ×´Ì¬¡£ÊÇÔÚ²¥·Å£¿ÔÝÍ££¿Í£Ö¹£¿
Private Function getPalyStatus(Optional strAlias As String = "mySound") As PlayStatus
    Dim sl As String * 255
    mciSendString "status " & strAlias & " mode", sl, Len(sl), 0
    If UCase$(Left$(sl, 7)) = "PLAYING" Or Left$(sl, 2) = "²¥·Å" Then
        getPalyStatus = IsPlaying
    ElseIf UCase$(Left$(sl, 6)) = "PAUSED" Or Left$(sl, 2) = "ÔÝÍ£" Then
        getPalyStatus = IsPaused
    Else
        getPalyStatus = IsStopped
    End If
End Function
' =========================================================================================
' ==== ÉùÒôÎÄ¼þµÄ²¥·Å£¨¿ÉÈ¥µô£©============================================================
' =========================================================================================



' =========================================================================================
' ==== ×ÖÌå¶Ô»°¿ò £¨µ¥¶À£©=================================================================
' =========================================================================================
' ÏÔÊ¾¶Ô»°¿òÖ®Ç°¡£×Ô¶¨Òå×ÖÌå¶Ô»°¿òÍâ¹Û¡£
Private Sub CustomizeFontDialog(ByVal hWnd As Long)
    Dim rcDlg As RECT, hWndParent As Long
    Dim pt As POINTAPI, W As Long, H As Long
    ' ¶Ô»°¿òµÄ¸¸´°¿Ú¾ä±ú¡£
    hWndParent = GetParent(hWnd)
    ' ÉèÖÃ¶Ô»°¿òÆô¶¯Î»ÖÃ£¿Ö»ÅÐ¶ÏÆÁÄ»ÖÐÐÄºÍËùÓÐÕßÖÐÐÄ£¬ÆäËû²»¹Ü£¡
    GetWindowRect hWnd, rcDlg ' È¡µÃ¶Ô»°¿ò¾ØÐÎ
    W = rcDlg.Right - rcDlg.Left: H = rcDlg.Bottom - rcDlg.Top + 120 ' ¶Ô»°¿ò¿í¡¢¸ß£¨¸ß¶ÈÒª¼Ó¸öÔ¤ÀÀÎÄ±¾¿ò¸ß£¡£©
    If m_dlgStartUpPosition = vbStartUpScreen Then ' ÔÙÒÆ¶¯¡£ÆÁÄ»ÖÐÐÄ
        MoveWindow hWnd, (Screen.Width \ Screen.TwipsPerPixelX - W) \ 2, (Screen.Height \ Screen.TwipsPerPixelY - H) \ 2, W, H, True
    ElseIf m_dlgStartUpPosition = vbStartUpOwner Then ' ËùÓÐÕßÖÐÐÄ
        Dim rcOwner As RECT, T As Long ' È¡µÃ¶Ô»°¿òµÄ¸¸´°¿Ú¾ØÐÎ
        GetWindowRect hWndParent, rcOwner
        T = rcOwner.Top + (rcOwner.Bottom - rcOwner.Top - H) \ 2
        If T < 0 Then T = 0 ' ±£Ö¤¶Ô»°¿ò²»³¬¹ýÆÁÄ»¶¥¶Ë£¡
        MoveWindow hWnd, rcOwner.Left + (rcOwner.Right - rcOwner.Left - W) \ 2, T, W, H, True
    Else ' Õâ¸öÌØ±ð£¬Òª´¦Àí¡£¸ß¶ÈÒª±ä£¬·ñÔò£¬Ìí¼ÓµÄÎÄ±¾¿òÎÞ·¨ÏÔÊ¾£¡
        MoveWindow hWnd, rcDlg.Left, rcDlg.Top, W, H, True
    End If
    ' Æô¶¯Î»ÖÃ£¬ÒòÎª¼¸ÖÖ¶Ô»°¿ò¶¼ÓÐÕâ¸öÊôÐÔ£¡¸ÄÓÃÒ»¸öº¯ÊýÊµÏÖ£¬Ð§¹û²»ºÃ£¡²»ÓÃÁË£¡
    'setDlgStartUpPosition hWndParent, GetParent(hWndParent)
    
    ' ´´½¨Ô¤ÀÀÎÄ±¾¿ò£¬Æä´óÐ¡¹Ì¶¨!£¡
    Dim rcP As RECT, ptL As POINTAPI  ' ¾ö¶¨Ô¤ÀÀÎÄ±¾¿òÎ»ÖÃºÍ´óÐ¡£¬rcL ÄÇ¸öËµÃ÷±êÇ©¾ØÐÎ£¡ID = 1093 ?¡£
    GetWindowRect GetDlgItem(hWnd, enumFONT_CTL.stc_Description), rcP
    ptL.X = rcP.Left: ptL.y = rcP.Top
    ScreenToClient hWnd, ptL ' ptL¾­¹ý×ª»¯ºó²ÅÄÜµÃµ½ÏëÒªµÄ½á¹û£¡
    ' ¿ªÊ¼´´½¨Ô¤ÀÀÎÄ±¾¿ò  ¼Ó Or WS_VISIBLE £¬ÔÚ´´½¨Ê±ÏÔÊ¾£¡ Or ES_READONLY ÉèÖÃÎªÖ»¶Á£¡
    Dim sT As String
    'Debug.Print " rcDlg.Bottom - rcP.Bottom + 20 " & rcDlg.Bottom - rcP.Bottom + 20 ' ÎÄ±¾¿ò¸ß¶È»á±ä£¿£¡²»Õý³££¿£¡
    sT = App.LegalCopyright & vbCrLf _
        & "Ò»¶þÈýËÄÎåÁùÆß°Ë¾ÅÊ®" & vbCrLf & "Ò¼·¡ÈþËÁÎéÂ½Æâ°Æ¾ÁÊ°" & vbCrLf _
        & "ABCDEFGHILMNOPQRSTUVWXYZ" & vbCrLf & "abcdefghilmnopqrstuvwxyz" & vbCrLf & "0123456789"
    hWndFontPreview = CreateWindowEx(WS_EX_STATICEDGE Or WS_EX_TOPMOST, _
        "Edit", sT, _
        WS_BORDER Or WS_CHILD Or WS_VISIBLE Or WS_HSCROLL Or WS_VSCROLL Or ES_AUTOHSCROLL Or ES_AUTOHSCROLL Or ES_MULTILINE Or ES_LEFT Or ES_WANTRETURN, _
        ptL.X, ptL.y + (rcP.Bottom - rcP.Top), _
        W - 20, 120, _
        hWnd, 0&, App.hInstance, &H520)
    ' ´´½¨ÐÂµÄ×ÖÌå£¬Ä¬ÈÏ×ÖÌå£¡£¡£¡
    Dim NewFont As Long, lpLF As LOGFONT
    With lpLF
        .lfCharSet = 134
        '.lfFaceName = "ËÎÌå"
        .lfItalic = False
        .lfStrikeOut = False
        .lfUnderline = False
        .lfWeight = 520
    End With
    NewFont = CreateFontIndirect(lpLF)
    ' ÉèÖÃÎÄ±¾¿ò×ÖÌå
    SendMessage hWndFontPreview, WM_SETFONT, NewFont, 0
End Sub
' ÉèÖÃ×ÖÌå¶Ô»°¿òÔ¤ÀÀ¡££¨ÎÄ×ÖÐ§¹û¶¯Ì¬±ä»¯¡£½ØÈ¡ WM_COMMAND ÏûÏ¢£¬ÅÐ¶Ï×ÖÌå¸ñÊ½ÓÐÄÄÐ©±ä»¯£¡£©
Private Sub mSetFontPreview(ByVal hWnd As Long, Optional wP As Long = 0&)
    Dim hFontToUse As Long, hdc As Long, RetValue As Long
    Dim lpLF As LOGFONT
    Dim tBuf As String * 80, sFontName As String
    Dim iIndex As Long, dwRGB As Long ' dwRGB ×ÖÌåÑÕÉ«Öµ£¡
    
     ' È¡µÃÑ¡ÔñµÄ×ÖÌåÐÅÏ¢
    SendMessage hWnd, WM_CHOOSEFONT_GETLOGFONT, wP, lpLF
    hFontToUse = CreateFontIndirect(lpLF): If hFontToUse = 0 Then Exit Sub
    hdc = GetDC(hWnd)
    SelectObject hdc, hFontToUse
    RetValue = GetTextFace(hdc, 79, tBuf)
    sFontName = Mid$(tBuf, 1, RetValue)

    ' È¡µÃÑ¡ÔñµÄ×ÖÌåÑÕÉ«
    iIndex = SendDlgItemMessage(hWnd, enumFONT_CTL.cbo_Color, CB_GETCURSEL, 0&, 0&)    ' cmb4
    If iIndex <> CB_ERR Then
        dwRGB = SendDlgItemMessage(hWnd, enumFONT_CTL.cbo_Color, CB_GETITEMDATA, iIndex, 0&)
    End If
    ' ´´½¨ÐÂµÄ×ÖÌå£¬ÑÕÉ«ÐÅÏ¢Ã»ÓÐ£¡£¡£¡
    Dim NewFont As Long
    NewFont = CreateFontIndirect(lpLF)
'    NewFont = CreateFont(Abs(lpLF.lfHeight * (72 / GetDeviceCaps(hDC, LOGPIXELSY))), 0, 0, 0, _
              lpLF.lfWeight, lpLF.lfItalic, lpLF.lfUnderline, lpLF.lfStrikeOut, _
              lpLF.lfCharSet, OUT_DEFAULT_PRECIS, CLIP_LH_ANGLES, _
              ANTIALIASED_QUALITY Or PROOF_QUALITY, DEFAULT_PITCH Or FF_DONTCARE, _
              sFontName)
    ' ÉèÖÃÎÄ±¾¿ò×ÖÌå
    SendMessage hWndFontPreview, WM_SETFONT, NewFont, 0
    ' ÉèÖÃÎÄ±¾¿òÎÄ×ÖÑÕÉ«£¨²»ÖªÎªÊ²Ã´£¬Ô¤ÀÀÑÕÉ«¸Ä±ä¡£¡£¡£ÎÞ·¨ÊµÏÖ£¡£¡£¡£©
'    SendMessage hWndFontPreview, 4103, 0, ByVal dwRGB
    If SetTextColor(GetDC(hWndFontPreview), dwRGB) = &HFFFF Then MsgBox "Ê§°Ü£ºÉèÖÃÎÄ×ÖÑÕÉ«³ö´í£¡", vbCritical
'    frmMain.txtNewCaption(2).ForeColor = dwRGB
'    frmMain.txtNewCaption(1).ForeColor = GetTextColor(GetDC(frmMain.txtNewCaption(2).hWnd))
'    Dim cDC As Long, chWnd As Long
'    chWnd = GetDlgItem(hWnd, enumFONT_CTL.btn_Apply)
'    cDC = GetDC(cDC)
'    Call SetTextColor(cDC, dwRGB)
'    Call SendMessage(chWnd, CDM_SetControlText, enumFONT_CTL.btn_Apply, ByVal "m_strControlsCaption(I)")
'    If dwRGB = GetTextColor(cDC) Then frmMain.BackColor = GetTextColor(cDC)
'    Dim sl As Long ' ÎÄ±¾¿òÖÐÎÄ×Ö¸öÊý£¬ÒªÈ¡µÃ£¬ÔÝÊ±¹Ì¶¨Ò»¸öÖµ£¡¡£
'    sl = 1024
'    Dim s As String: s = String(sl, 0)
'    GetWindowText hWndFontPreview, s, sl
'    Debug.Print Replace$(s, Chr$(0), "")
'    ' ÖØÐÂÉèÖÃÎÄ×Ö
'    SendMessage hWndFontPreview, WM_SETTEXT, -1, ByVal s & vbCrLf & App.LegalCopyright

    ' ÊÍ·Å×ÊÔ´
    ReleaseDC hWnd, hdc
End Sub

Private Function LOWORD(Param As Long) As Long
    LOWORD = Param And &HFFFF&
End Function
Private Function HIWORD(Param As Long) As Long
    HIWORD = Param \ &H10000 And &HFFFF&
End Function
' =========================================================================================
' ==== ×ÖÌå¶Ô»°¿ò £¨µ¥¶À£©=================================================================
' =========================================================================================
