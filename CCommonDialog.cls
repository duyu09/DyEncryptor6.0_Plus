VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCommonDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ##############################################################################
' CCommonDialog.cls
' 功  能：使用 Windows 通用对话框，如下：
'         0、文件属性对话框
'         1、打开对话框（可以提供某些文件预览）
'         2、保存对话框
'         3、字体对话框（预览？）
'         4、颜色对话框
' 依  赖：（MCDHook.bas）
' 作  者：鹤望兰·流
' 版  权：鹤望兰·流
' 日  期：2008-09-16
' 网  站：http://hewanglan34512.cech.com.cn
' E-mail：hewanglan34512@163.com
'* (******* 复制请保留以上信息 *******)
' ##############################################################################
'Download by http://www.codefans.net
' --- API 函数 申明
' 取得给定文件 、文件夹的属性（显示属性对话框）
Private Declare Function ShellExecuteEx Lib "shell32.dll" (ByRef s As SHELLEXECUTEINFO) As Long ' for ShowProperty
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

Rem --------------------------------------------------------
Rem Get any errors during execution of common dialogs
Rem --------------------------------------------------------
Private Declare Function CommDlgExtendedError Lib "comdlg32.dll" () As Long

' --- 常数 申明
Private Const SW_SHOW = 5                       ' for ShowProperty
Private Const SEE_MASK_INVOKEIDLIST = &HC       ' for ShowProperty

' --- 枚举 申明
Public Enum DialogFlags
    OFN_SHAREFALLTHROUGH = 2
    OFN_SHARENOWARN = 1
    OFN_SHAREWARN = 0
    OFN_ReadOnly = &H1                      ' 只读选项为选中
    OFN_OverWritePrompt = &H2               ' 覆盖时提示
    OFN_HideReadOnly = &H4                  ' 隐藏只读选项，注意：保存对话框始终不显示！
    OFN_NoChangeDir = &H8
    OFN_ShowHelp = &H10
    OFN_EnableHook = &H20                   ' 使用回调函数，注意：使用后对话框变为老式外观！
    OFN_EnableTemplate = &H40
    OFN_EnableTemplateHandle = &H80
    OFN_NoValidate = &H100
    OFN_AllowMultiSelect = &H200           ' 允许多个选中
    OFN_ExtenSionDifferent = &H400
    OFN_PathMustExist = &H800
    OFN_FileMustExist = &H1000
    OFN_CreatePrompt = &H2000
    OFN_ShareAware = &H4000
    OFN_NoReadOnlyReturn = &H8000
    OFN_NoTestFileCreate = &H10000
    OFN_NoNetWorkButton = &H20000
    OFN_NoLongNames = &H40000               ' force no long names for 4.x modules
    OFN_Explorer = &H80000                  ' 资源管理器风格外观
    OFN_NoDereferenceLinks = &H100000
    OFN_LongNames = &H200000                ' force long names for 3.x modules
    OFN_EnableIncludeNotify = &H400000
    OFN_EnableSizing = &H800000
    OFN_UseMonikers = &H1000000
    OFN_DontAddToRecent = &H2000000
    OFN_ForceShowHidden = &H10000000        ' 强制显示隐藏文件、文件夹，不管用！！！？？？
    cdlCancel = 32755
    
'    PD_ALLPAGES = &H0
'    PD_COLLATE = &H10
'    PD_DISABLEPRINTTOFILE = &H80000
'    PD_ENABLEPRINTHOOK = &H1000
'    PD_ENABLEPRINTTEMPLATE = &H4000
'    PD_ENABLEPRINTTEMPLATEHANDLE = &H10000
'    PD_ENABLESETUPHOOK = &H2000
'    PD_ENABLESETUPTEMPLATE = &H8000
'    PD_ENABLESETUPTEMPLATEHANDLE = &H20000
'    PD_HIDEPRINTTOFILE = &H100000
'    PD_NONETWORKBUTTON = &H200000
'    PD_NOPAGENUMS = &H8
'    PD_NOSELECTION = &H4
'    PD_NOWARNING = &H80
'    PD_PAGENUMS = &H2
'    PD_PRINTSETUP = &H40
'    PD_PRINTTOFILE = &H20
'    PD_RETURNDC = &H100
'    PD_RETURNDEFAULT = &H400
'    PD_RETURNIC = &H200
'    PD_SELECTION = &H1
'    PD_SHOWHELP = &H800
'    PD_USEDEVMODECOPIES = &H40000
'    PD_USEDEVMODECOPIESANDCOLLATE = &H40000
End Enum
'Public Enum PreviewPosition ' 预览图片框位置
'    ppNone = -1
'    ppTop = 0
'    ppLeft = 1
'    ppRight = 2
'    ppBottom = 3
'End Enum

' --- 结构体 申明
Private Type SHELLEXECUTEINFO                   ' for ShowProperty
    cbSize As Long
    fMask As Long
    hWnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    ' optional fields
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type
Private Type OPENFILENAME                       ' for ShowOpen
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    lFlags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

' --- 成员变量 申明
' for ShowOpen
Private m_blnForceShowNewSkin As Boolean
Private m_blnCancelError As Boolean
Private m_strFileName As String
Private m_strFileTitle As String
Private m_strFilter As String
Private m_lngFilterIndex As Long
Private m_strDefaultExt As String
Private m_strDialogTitle As String
Private m_lngFlags As DialogFlags
Private m_lngHwnd As Long
Private m_lngInstance As Long
Private m_strInitDir As String
Private m_blnReadOnlyChecked As Boolean ' 指示是否选定只读复选框，在 MCDHook.bas 中取得！！

'Private m_picLogoPicture As PictureBox ' 程序标志图片框图片
'Private m_picPreviewPicture As PictureBox ' 预览图片框图片
'Private m_ppLogoPosition As PreviewPosition ' 程序标志图片框位置
'Private m_ppPreviewPosition As PreviewPosition ' 预览图片框位置

' ==== 字体对话框 （单独）=================================================================
' 字体对话框成员变量
Private m_FontFlags As CF_Flags ' 字体对话框标志
Private m_ReturnFont As StdFont ' 字体对象属性
' ==== 字体对话框 （单独）=================================================================

' ==== 颜色对话框 （单独）=================================================================
Private m_ColorFlags As CC_Flags    ' 颜色对话框标志
Private m_ReturnColor As Long       ' 返回颜色属性
' ==== 颜色对话框 （单独）=================================================================

' --- 私有变量 申明


' ==================================================================================
' === 公共属性（核心功能）==========================================================
' ==================================================================================
' 属性：点取消按钮是否引发错误？（特别注意：四个对话框共有！）
Public Property Get CancelError() As Boolean
    'return object's CancelError property
    CancelError = m_blnCancelError
End Property
Public Property Let CancelError(ByVal bCancelError As Boolean)
    'assign object's CancelError property
    m_blnCancelError = bCancelError
End Property

' 属性：是否强制显示新外观？？？！！！ 76 老式外观 88 新式外观（Win98可能导致无法显示对话框！）
Public Property Get ForceShowNewSkin() As Boolean
    ForceShowNewSkin = m_blnForceShowNewSkin
End Property
Public Property Let ForceShowNewSkin(ByVal vNewValue As Boolean)
    m_blnForceShowNewSkin = vNewValue
End Property

' 属性：对话框启动位置？（特别注意：四个对话框共有！）
Public Property Get dlgStartUpPosition() As StartUpPositionConstants
    dlgStartUpPosition = m_dlgStartUpPosition
End Property
Public Property Let dlgStartUpPosition(ByVal vNewValue As StartUpPositionConstants)
    m_dlgStartUpPosition = vNewValue
End Property
' 属性：文件名（完整路径）
Public Property Get FileName() As String
    FileName = m_strFileName
End Property
Public Property Let FileName(ByVal sFilename As String)
    m_strFileName = sFilename
End Property

' 属性：文件名（短文件名+后缀）设置了 OFN_AllowMultiSelect 后，无用！
Public Property Get FileTitle() As String
    FileTitle = m_strFileTitle
End Property
Public Property Let FileTitle(ByVal vNewValue As String)
    m_strFileTitle = vNewValue
End Property

' 属性：文件类型过滤下拉列表文字 "媒体文件(*.piano)|*.piano|所有文件(*.*)|*.*"
Public Property Get Filter() As String
    Filter = m_strFilter
End Property
Public Property Let Filter(ByVal sFilter As String)
    Dim s As String
    'm_strFilter = sFilter
    ' To make Windows-style filter, replace | and : with nulls
    Dim ch As String, I As Integer
    For I = 1 To Len(sFilter)
        ch = Mid$(sFilter, I, 1)
        If ch = "|" Or ch = ":" Then
            s = s & vbNullChar
        Else
            s = s & ch
        End If
    Next I
    m_strFilter = s & vbNullChar & vbNullChar
End Property

' 属性：文件类型过滤下拉列表序号，无用？！！！
Public Property Get FilterIndex() As Long
    FilterIndex = m_lngFilterIndex
End Property
Public Property Let FilterIndex(ByVal lFilterIndex As Long)
    m_lngFilterIndex = lFilterIndex
End Property

' 属性：保存对话框默认后缀！！！只在另存为所有文件时有效！！！
Public Property Get DefaultExt() As String
    DefaultExt = m_strDefaultExt
End Property
Public Property Let DefaultExt(ByVal sDefaultExt As String)
    m_strDefaultExt = sDefaultExt
End Property

' 属性：对话框标题（特别注意：四个对话框共有！）
Public Property Get DialogTitle() As String
    DialogTitle = m_strDialogTitle
End Property
Public Property Let DialogTitle(ByVal vNewValue As String)
    m_strDialogTitle = vNewValue
End Property

' 属性：对话框外观行为标志
Public Property Get Flags() As DialogFlags
    Flags = m_lngFlags
End Property
Public Property Let Flags(ByVal vNewValue As DialogFlags)
    m_lngFlags = vNewValue
End Property

' 属性：句柄，（特别注意：四个对话框共有！）特别注意：所有对话框都要设置这个属性！！！否则可能无法使用！！！
Public Property Get hWnd() As Long
    hWnd = m_lngHwnd
End Property
Public Property Let hWnd(ByVal vNewValue As Long)
    m_lngHwnd = vNewValue
End Property

' 属性：实例（特别注意：四个对话框共有！）
Public Property Get Instance() As Long
    Instance = m_lngInstance
End Property
Public Property Let Instance(ByVal vNewValue As Long)
    m_lngInstance = vNewValue
End Property

' 属性：对话框初始路径
Public Property Get InitDir() As String
    InitDir = m_strInitDir
End Property
Public Property Let InitDir(ByVal vNewValue As String)
    m_strInitDir = vNewValue
End Property

' 属性：指示是否选定只读复选框，默认为False
Public Property Get ReadOnlyChecked() As Boolean
    ReadOnlyChecked = MCDHook.IsReadOnlyChecked ' m_blnReadOnlyChecked
End Property
Public Property Let ReadOnlyChecked(ByVal vNewValue As Boolean)
    m_blnReadOnlyChecked = vNewValue
End Property

' 属性：程序标志图片框图片
Public Property Get LogoPicture() As PictureBox
    LogoPicture = m_picLogoPicture
End Property
Public Property Set LogoPicture(ByVal vNewValue As PictureBox)
    Set m_picLogoPicture = vNewValue
End Property
' 属性：程序标志图片框位置
Public Property Get ppLogoPosition() As PreviewPosition
    ppLogoPosition = m_ppLogoPosition
End Property
Public Property Let ppLogoPosition(ByVal vNewValue As PreviewPosition)
    m_ppLogoPosition = vNewValue
End Property

' 属性：预览图片框图片
Public Property Get PreviewPicture() As PictureBox
    PreviewPicture = m_picPreviewPicture
End Property
Public Property Set PreviewPicture(ByVal vNewValue As PictureBox)
    Set m_picPreviewPicture = vNewValue
End Property

' 属性：预览图片框位置
Public Property Get ppPreviewPosition() As PreviewPosition
    ppPreviewPosition = m_ppPreviewPosition
End Property
Public Property Let ppPreviewPosition(ByVal vNewValue As PreviewPosition)
    m_ppPreviewPosition = vNewValue
End Property

' ==== 字体对话框 （单独）=================================================================
' 属性：字体对话框返回值
Public Property Get getReturnFont() As StdFont
    Set getReturnFont = m_ReturnFont
End Property
Public Property Let getReturnFont(vNewValue As StdFont)
    Set m_ReturnFont = vNewValue
End Property
' 属性：对话框外观行为标志
Public Property Get FontFlags() As CF_Flags
    FontFlags = m_FontFlags
End Property
Public Property Let FontFlags(ByVal vNewValue As CF_Flags)
    m_FontFlags = vNewValue
End Property
' ==== 字体对话框 （单独）=================================================================

' ==== 颜色对话框 （单独）=================================================================
' 属性：返回选择的颜色
Public Property Get getReturnColor() As Long
    getReturnColor = m_ReturnColor
End Property
Public Property Let getReturnColor(vNewValue As Long)
    m_ReturnColor = vNewValue
End Property
' 属性：对话框外观行为标志
Public Property Get ColorFlags() As CC_Flags
    ColorFlags = m_ColorFlags
End Property
Public Property Let ColorFlags(ByVal vNewValue As CC_Flags)
    m_ColorFlags = vNewValue
End Property
' ==== 颜色对话框 （单独）=================================================================
' ==================================================================================
' === 公共属性（核心功能）==========================================================
' ==================================================================================


' ==================================================================================
' === 公共函数（核心功能）==========================================================
' ==================================================================================
Public Sub ShowProperty(strFilePath As String)
    ' ##############################################################################
    ' 功  能：取得给定文件 、文件夹的属性（显示属性对话框）
    ' 函数名：ShowProperty
    ' 参  数：strFilePath ，字符串：文件 、文件夹路径
    ' 返回值：无
    ' ##############################################################################
    Dim shInfo As SHELLEXECUTEINFO
    With shInfo
        .cbSize = LenB(shInfo)
        .lpFile = strFilePath
        .nShow = SW_SHOW
        .fMask = SEE_MASK_INVOKEIDLIST
        .lpVerb = "properties"
    End With
    ShellExecuteEx shInfo
End Sub

Public Function ShowOpen() As Boolean
    ' ##############################################################################
    ' 功  能：显示打开对话框
    ' 函数名：ShowOpen
    ' 参  数：无
    ' 返回值：逻辑值：是否成功返回（有没点取消按钮？！）
    ' ##############################################################################
    On Error Resume Next
    Dim OFN As OPENFILENAME
    Dim lFileSize As Long

    ' 初始化结构体 OFN
    With OFN
        .lpfnHook = VBGetProcAddress(AddressOf CDCallBackFun) ' 回调函数地址？？？！！！
        ' 是否强制显示新外观？？？！！！ 76 老式外观 88 新式外观（Win98可能导致无法显示对话框！）
        If m_blnForceShowNewSkin Then
            .lStructSize = 88
        Else
            .lStructSize = Len(OFN)
        End If
        .hwndOwner = m_lngHwnd
        .hInstance = m_lngInstance
        .lFlags = m_lngFlags
        .lpstrTitle = m_strDialogTitle
        ' 当允许多选时，必须设置 OFN_Explorer ，否则程序会崩溃！？？？
        ' 不会崩溃，但强制显示新外观时，多选由允许改为不允许后，无法显示对话框！！！
        If (.lFlags And OFN_AllowMultiSelect) = OFN_AllowMultiSelect Then
            .lFlags = .lFlags Or OFN_Explorer
            lFileSize = IIf(m_blnForceShowNewSkin, 1024, 8192)  ' 强制显示新外观惹的祸？？？！！！
            ' 8192 这个值？？？导致无法显示对话框！！！？？？1024时，图片框在底端无法定位！！！？？？
        Else
            lFileSize = 1024
        End If
        
        .lpstrFile = m_strFileName & String$(lFileSize - Len(m_strFileName), 0)
        .nMaxFile = lFileSize
        .lpstrFileTitle = m_strFileTitle & String$(lFileSize - Len(FileTitle), 0)
        .nMaxFileTitle = lFileSize
        
        .lpstrFilter = m_strFilter
        .nFilterIndex = m_lngFilterIndex
        .lpstrInitialDir = m_strInitDir & vbNullChar & vbNullChar

        ' 设置标志是否选定只读复选框
        If m_blnReadOnlyChecked Then .lFlags = .lFlags Or OFN_ReadOnly
        
        ' 显示打开对话框？！
        WhichStyle = ssOpen
        If GetOpenFileName(OFN) Then
            If (.lFlags And OFN_AllowMultiSelect) = OFN_AllowMultiSelect Then
                m_strFileName = .lpstrFile
            Else
                m_strFileName = TrimNull(.lpstrFile)
            End If
            m_strFileTitle = .lpstrFileTitle
            ShowOpen = True
        Else
            ShowOpen = False
            If m_blnCancelError Then
                Err.Raise cdlCancel, App.EXEName & ".CCommonDialog", "用户选择了取消。"
            End If
        End If
    End With
End Function
Public Function ParseMultiFileName(ByRef sDir As String, ByRef sFiles() As String, ByRef lFileCount As Long) As Boolean
    ' ##############################################################################
    ' 功  能：取得打开对话框选择多个文件时的所有文件名
    ' 函数名：ParseMultiFileName
    ' 参  数：sDir，字符串：文件目录；sFiles 字符数组：文件名；lFileCount 长整型：文件数目
    ' 返回值：逻辑值：是否选中多个文件（注意：参数也作返回值！）
    ' ##############################################################################
    Dim lPos As Long
    Dim lNextPos As Long
    Dim sAllFiles As String
    On Error Resume Next

    lPos = InStr(m_strFileName, vbNullChar & vbNullChar)
    sAllFiles = Left$(m_strFileName, lPos - 1)
    lNextPos = InStr(sAllFiles, vbNullChar)
    If lNextPos <> 0 Then ' 选了多个文件？
        ParseMultiFileName = True
        
        sDir = Mid$(sAllFiles, 1, lNextPos - 1)
        sDir = IIf(Right$(sDir, 1) <> "\", sDir & "\", sDir)
        
        Do While lNextPos <> 0
            lPos = lNextPos + 1
            lNextPos = InStr(lPos, sAllFiles, vbNullChar)
            lFileCount = lFileCount + 1
            ReDim Preserve sFiles(0 To lFileCount - 1) As String
            If lNextPos > 0 Then
                sFiles(lFileCount - 1) = Mid$(sAllFiles, lPos, lNextPos - lPos)
            Else
                sFiles(lFileCount - 1) = Mid$(sAllFiles, lPos)
            End If
        Loop
    Else ' 选单个文件
        ParseMultiFileName = False
        lFileCount = 1
        ReDim sFiles(0)
        lPos = InStrRev(m_strFileName, "\")
        If lPos > 0 Then
            sDir = Left$(m_strFileName, lPos)
            sFiles(0) = TrimNull(Right$(m_strFileName, Len(m_strFileName) - lPos))
        Else
            sDir = ""
            sFiles(0) = m_strFileName
        End If
    End If
End Function

Public Function ShowSave() As Boolean
    ' ##############################################################################
    ' 功  能：显示保存对话框
    ' 函数名：ShowSave
    ' 参  数：无
    ' 返回值：逻辑值：是否成功返回（有没点取消按钮？！）
    ' ##############################################################################
    On Error Resume Next
    Dim OFN As OPENFILENAME
    With OFN
        .lpfnHook = VBGetProcAddress(AddressOf CDCallBackFun) ' 回调函数地址？？？！！！
        ' 是否强制显示新外观？？？！！！ 76 老式外观 88 新式外观（Win98可能导致无法显示对话框！）
        If m_blnForceShowNewSkin Then
            .lStructSize = 88
        Else
            .lStructSize = Len(OFN)
        End If
        .hwndOwner = m_lngHwnd
        .hInstance = m_lngInstance
        .lFlags = m_lngFlags
        .lpstrTitle = m_strDialogTitle
        
        .lpstrFile = m_strFileName & String$(1024 - Len(m_strFileName), 0)
        .nMaxFile = 1024
        .lpstrFileTitle = m_strFileTitle & String$(1024 - Len(FileTitle), 0)
        .nMaxFileTitle = 1024
        
        .lpstrFilter = m_strFilter
        .nFilterIndex = m_lngFilterIndex
        .lpstrInitialDir = m_strInitDir & vbNullChar & vbNullChar
        .lpstrDefExt = m_strDefaultExt & vbNullChar & vbNullChar ' ShowOpen 不需要此值！

        ' 显示保存对话框？！
        WhichStyle = ssSave
        If GetSaveFileName(OFN) Then
            ShowSave = True
            m_strFileName = TrimNull(.lpstrFile)
            m_strFileTitle = .lpstrFileTitle
        Else
            ShowSave = False
            If m_blnCancelError Then
                Err.Raise cdlCancel, App.EXEName & ".CCommonDialog", "用户选择了取消。"
            End If
        End If
    End With
End Function

' =========================================================================================
' ==== 自定义对话框上的控件隐藏、显示，改变文字等！（可去掉）==============================
' =========================================================================================
' 隐藏、显示打开、保存对话框上的控件
Public Sub HideDlgControls(ByVal Index As Integer)
    m_blnHideControls(Index) = True
End Sub
Public Sub ShowDlgControls(ByVal Index As Integer)
    m_blnHideControls(Index) = False
End Sub
' 设置对话框上的控件的文字。m_strControlsCaption(I) 决定其值，默认为空，不改变原始值！
Public Sub SetDlgControlsCaption(ByVal Index As Integer, strNewCaption As String)
    m_strControlsCaption(Index) = strNewCaption
End Sub
' =========================================================================================
' ==== 自定义对话框上的控件隐藏、显示，改变文字等！（可去掉）==============================
' =========================================================================================


' ==== 字体对话框 （单独）=================================================================
Public Function ShowFont(Optional ByRef FontRGB As Long) As Boolean
' 注意：按地址传递参数，当返回值！可取得选择的字体颜色！！！
    Const FW_BOLD = 700
    Const cdlCFScreenFonts = &H1
    
    Dim CF As ChooseFontType
    Dim LF As LOGFONT
    Dim TempByteArray() As Byte
    Dim ByteArrayLimit As Long
    Dim OldhDC As Long
    Dim FontToUse As Long
    Dim tBuf As String * 80
    Dim X As Long
    Dim uFlag As Long
    
    Dim RetValue As Long

    ' 字体信息变量！
    Dim mRGBResult As Long      ' 文字颜色？！
    Dim mFontName As String     ' 字体名称
    Dim mItalic As Boolean      ' 是否斜体
    Dim mUnderline As Boolean   ' 是否下划线
    Dim mStrikethru As Boolean  ' 是否删除线
    Dim mFontSize As Long       ' 字体大小
    Dim mBold As Boolean        ' 是否加粗
    Dim mCharSet As Integer     ' 字符集（中文？西文？）
    Dim mWeight As Integer
    
    ' 设置对话框初始字体。
    mFontName = m_ReturnFont.Name
    mFontSize = m_ReturnFont.Size
    mBold = m_ReturnFont.Bold
    mCharSet = m_ReturnFont.Charset
    
    TempByteArray = StrConv(mFontName & vbNullChar, vbFromUnicode)
    ByteArrayLimit = UBound(TempByteArray)
    ' 设置两个结构体的值！
    With LF
        For X = 0 To ByteArrayLimit
            .lfFaceName(X) = TempByteArray(X)
        Next
            .lfHeight = mFontSize / 72 * GetDeviceCaps(GetDC(m_lngHwnd), LOGPIXELSY)
            .lfItalic = mItalic * -1
            .lfUnderline = mUnderline * -1
            .lfStrikeOut = mStrikethru * -1
            .lfCharSet = mCharSet
        If mBold Then .lfWeight = FW_BOLD
    End With
    With CF
        .lStructSize = Len(CF)
        .hwndOwner = m_lngHwnd
        .hdc = GetDC(m_lngHwnd)
        .lpLogFont = lstrcpyANY(LF, LF)
        If Not uFlag Then
            .Flags = CF_BOTH Or CF_WYSIWYG
        Else
            .Flags = uFlag Or CF_BOTH Or CF_WYSIWYG
        End If
        .Flags = .Flags Or m_FontFlags
        .rgbColors = mRGBResult
        .lCustData = 0
        .lpfnHook = VBGetProcAddress(AddressOf CDCallBackFun)
        .lpTemplateName = 0
        .hInstance = m_lngInstance
        .lpszStyle = 0
        .nFontType = SCREEN_FONTTYPE
        .nSizeMin = 0
        .nSizeMax = 0
        .iPointSize = 14
    End With
    
    ' 显示字体对话框？！
    WhichStyle = ssFont
    RetValue = ChooseFont(CF)
    
    ' 判断字体对话框返回值，是否成功？
    If RetValue Then
        ShowFont = True
        With LF
            mItalic = .lfItalic * -1
            mUnderline = .lfUnderline * -1
            mStrikethru = .lfStrikeOut * -1
            mCharSet = .lfCharSet
            mWeight = .lfWeight
        End With
        With CF
            mFontSize = .iPointSize \ 10
            mRGBResult = .rgbColors
            If .nFontType And BOLD_FONTTYPE Then
                mBold = True
            Else
                mBold = False
            End If
        End With
        
        FontToUse = CreateFontIndirect(LF)
        If FontToUse = 0 Then Exit Function
        OldhDC = SelectObject(CF.hdc, FontToUse)
        RetValue = GetTextFace(CF.hdc, 79, tBuf)
        mFontName = Mid$(tBuf, 1, RetValue)
        ' 返回字体信息到成员变量
        With m_ReturnFont
            .Bold = mBold
            .Charset = mCharSet
            .Italic = mItalic
            .Name = mFontName
            .Size = mFontSize
            .Strikethrough = mStrikethru
            .Underline = mUnderline
            .Weight = mWeight
        End With
        ' 注意：颜色！在标准字体结构体中没有？！
        FontRGB = mRGBResult
    Else
        ShowFont = False
        If m_blnCancelError Then
            Err.Raise cdlCancel, App.EXEName & ".CCommonDialog", "用户选择了取消。"
        End If
    End If
   
End Function
' ==== 字体对话框 （单独）=================================================================

' ==== 颜色对话框 （单独）=================================================================
Public Function ShowColor(Optional ByRef CError As Long, Optional ByRef ExCError As Long, Optional ByVal lInitColor As Long = vbBlue) As Boolean
' 注意：参数 CError 按地址传递参数，当返回值！返回错误号！？
' ShowColor 函数返回值：成功否？
' 默认返回的颜色值：初始化对话框时的颜色。蓝色！
    ShowColor = False: m_ReturnColor = lInitColor: CError& = 0

    If lInitColor < 0 Then lInitColor = 0

    Dim C As CHOOSECOLOR_TYPE
    Dim MemHandle As Long, OK As Long
    Dim Address As Long
    Dim wSize As Long
    Dim I As Long
    Dim result As Long

    ReDim ClrArray(15) As Long    ' for 16 custom colors
    wSize = Len(ClrArray(0)) * 16 ' block memory size

    Rem ----------------------------------------------------
    Rem  I prepare a block memory size to keep
    Rem  custom colors
    Rem ----------------------------------------------------
    MemHandle = GlobalAlloc(GHND, wSize)
    If MemHandle = 0 Then
        ExCError = 1 ' return error code
        Exit Function
    End If

    Address = GlobalLock(MemHandle)
    If Address = 0 Then
        ExCError = 2 ' return error code
        Exit Function
    End If

    Rem ----------------------------------------------------
    Rem Setall custom colors WHITE
    Rem ----------------------------------------------------
    For I& = 0 To UBound(ClrArray)
        ClrArray(I&) = &HFFFFFF
    Next

    Rem ----------------------------------------------------
    Rem copy custom colors to block memory
    Rem ----------------------------------------------------
    Call CopyMemory(ByVal Address, ClrArray(0), wSize)

    Rem ----------------------------------------------------
    Rem fill CHOOSECOLOR structure to open the
    Rem Colors dialog
    Rem ----------------------------------------------------
    C.lStructSize = Len(C)
    C.hwndOwner = m_lngHwnd
    C.hInstance = m_lngInstance
    C.lpCustColors = Address
    C.rgbResult = lInitColor
    C.Flags = m_ColorFlags
    C.lpfnHook = VBGetProcAddress(AddressOf CDCallBackFun)
    
    ' 显示颜色对话框？！
    WhichStyle = ssColor
    result = ChooseColor(C)
    CError = CommDlgExtendedError()

    ' 用户选择了取消
    If result = 0 Then
        If m_blnCancelError Then
            Err.Raise cdlCancel, App.EXEName & ".CCommonDialog", "用户选择了取消。"
        End If
        ExCError = 3 '  return error code
        Exit Function
    End If

    Rem ----------------------------------------------------
    Rem copy custom colors
    Rem ----------------------------------------------------
    Call CopyMemory(ClrArray(0), ByVal Address, wSize)
    Rem relelase resource
    OK = GlobalUnlock(MemHandle)
    OK = GlobalFree(MemHandle)

    Rem ----------------------------------------------------
    Rem Return color code selected
    Rem ----------------------------------------------------
    'retChooseColor& = C.rgbResult
    m_ReturnColor = C.rgbResult  ' 返回选择的颜色值！
    ShowColor = True ' 函数成功，返回！
    
    Rem ----------------------------------------------------
    Rem This is custom colors (not used here)
    Rem ----------------------------------------------------
    'For i& = 0 To UBound(ClrArray)
    '    Debug.Print "Custom Color"; Str$(i&); ":", Hex$(ClrArray(i&))
    'Next
    
End Function
' ==== 颜色对话框 （单独）=================================================================

' ==================================================================================
' === 公共函数（核心功能）==========================================================
' ==================================================================================


' ==================================================================================
' === 私有函数（辅助功能）==========================================================
' ==================================================================================
Private Function TrimNull(ByVal strS As String) As String
    ' ##############################################################################
    ' 功  能：去掉字符串尾 vbNullChar==Chr$(0)
    ' 函数名：TrimNull
    ' 参  数：strS ，字符串：预处理的字符串
    ' 返回值：字符串：去掉字符串尾 vbNullChar 的新字符串
    ' ##############################################################################
    Dim pos As Long
    pos = InStr(strS, Chr$(0))
    If pos Then strS = Left$(strS, pos - 1)
    TrimNull = strS
End Function

Private Function VBGetProcAddress(ByVal lpfnAddressOfX As Long) As Long
    ' ##############################################################################
    ' 功  能：取得函数的地址
    ' 函数名：VBGetProcAddress
    ' 参  数：AddressOf 与取得地址的函数名
    ' 返回值：长整型：函数的地址
    ' ##############################################################################
    VBGetProcAddress = lpfnAddressOfX
End Function
' ==================================================================================
' === 私有函数（辅助功能）==========================================================
' ==================================================================================



' ==================================================================================
' 类的初始化及销毁
Private Sub Class_Initialize()
    ' 成员变量（属性）初始值设定，有些使用默认值！
    m_blnForceShowNewSkin = True ' 是否强制显示新外观？
    m_lngFlags = OFN_Explorer Or OFN_EnableHook  ' 对话框风格 资源管理器风格且使用回调！！！
    m_strFilter = "文本文件(*.txt)|*.txt|所有文件(*.*)|*.*" ' 文件类型过滤下拉列表文字
    
    ' MCDHook 中的变量！
    Set m_picLogoPicture = Nothing ' 程序标志图片框图片
    Set m_picPreviewPicture = Nothing ' 预览图片框图片
    m_ppLogoPosition = ppTop ' 程序标志图片框位置
    m_ppPreviewPosition = ppBottom ' 预览图片框位置
    m_dlgStartUpPosition = vbStartUpScreen ' 对话框启动位置？屏幕中心
    
    ' 是否隐藏对话框上的控件？（可去掉）
    Dim I As Integer
    For I = 0 To 8
        m_blnHideControls(I) = False
    Next I
    ' 设置对话框上的控件的文字？
    For I = 0 To 6
        m_strControlsCaption(I) = ""
    Next I
    ' ==== 字体对话框 （单独）=================================================================
    Set m_ReturnFont = New StdFont ' 字体对象属性
    With m_ReturnFont
        .Bold = False
        .Charset = 134 ' 中文字符集
        .Italic = False
        .Name = "宋体"
        .Size = 12
        .Strikethrough = False
        .Underline = False
        .Weight = 520
    End With
    m_FontFlags = CF_INITTOLOGFONTSTRUCT Or CF_ENABLEHOOK Or CF_EFFECTS Or CF_APPLY 'Or CF_SHOWHELP ' 字体对话框标志
    ' ==== 字体对话框 （单独）=================================================================
    ' ==== 颜色对话框 （单独）=================================================================
    m_ColorFlags = CC_ENABLEHOOK Or CC_RGBINIT Or CC_FULLOPEN
    m_ReturnColor = vbBlue ' 默认返回的颜色值：蓝色！
    ' ==== 颜色对话框 （单独）=================================================================
End Sub
Private Sub Class_Terminate()
    Set m_picLogoPicture = Nothing ' 程序标志图片框图片
    Set m_picPreviewPicture = Nothing ' 预览图片框图片
    Set m_ReturnFont = Nothing
End Sub
' 类的初始化及销毁
' ==================================================================================
