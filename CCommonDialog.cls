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
' ��  �ܣ�ʹ�� Windows ͨ�öԻ������£�
'         0���ļ����ԶԻ���
'         1���򿪶Ի��򣨿����ṩĳЩ�ļ�Ԥ����
'         2������Ի���
'         3������Ի���Ԥ������
'         4����ɫ�Ի���
' ��  ������MCDHook.bas��
' ��  �ߣ�����������
' ��  Ȩ������������
' ��  �ڣ�2008-09-16
' ��  վ��http://hewanglan34512.cech.com.cn
' E-mail��hewanglan34512@163.com
'* (******* �����뱣��������Ϣ *******)
' ##############################################################################
'Download by http://www.codefans.net
' --- API ���� ����
' ȡ�ø����ļ� ���ļ��е����ԣ���ʾ���ԶԻ���
Private Declare Function ShellExecuteEx Lib "shell32.dll" (ByRef s As SHELLEXECUTEINFO) As Long ' for ShowProperty
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

Rem --------------------------------------------------------
Rem Get any errors during execution of common dialogs
Rem --------------------------------------------------------
Private Declare Function CommDlgExtendedError Lib "comdlg32.dll" () As Long

' --- ���� ����
Private Const SW_SHOW = 5                       ' for ShowProperty
Private Const SEE_MASK_INVOKEIDLIST = &HC       ' for ShowProperty

' --- ö�� ����
Public Enum DialogFlags
    OFN_SHAREFALLTHROUGH = 2
    OFN_SHARENOWARN = 1
    OFN_SHAREWARN = 0
    OFN_ReadOnly = &H1                      ' ֻ��ѡ��Ϊѡ��
    OFN_OverWritePrompt = &H2               ' ����ʱ��ʾ
    OFN_HideReadOnly = &H4                  ' ����ֻ��ѡ�ע�⣺����Ի���ʼ�ղ���ʾ��
    OFN_NoChangeDir = &H8
    OFN_ShowHelp = &H10
    OFN_EnableHook = &H20                   ' ʹ�ûص�������ע�⣺ʹ�ú�Ի����Ϊ��ʽ��ۣ�
    OFN_EnableTemplate = &H40
    OFN_EnableTemplateHandle = &H80
    OFN_NoValidate = &H100
    OFN_AllowMultiSelect = &H200           ' �������ѡ��
    OFN_ExtenSionDifferent = &H400
    OFN_PathMustExist = &H800
    OFN_FileMustExist = &H1000
    OFN_CreatePrompt = &H2000
    OFN_ShareAware = &H4000
    OFN_NoReadOnlyReturn = &H8000
    OFN_NoTestFileCreate = &H10000
    OFN_NoNetWorkButton = &H20000
    OFN_NoLongNames = &H40000               ' force no long names for 4.x modules
    OFN_Explorer = &H80000                  ' ��Դ������������
    OFN_NoDereferenceLinks = &H100000
    OFN_LongNames = &H200000                ' force long names for 3.x modules
    OFN_EnableIncludeNotify = &H400000
    OFN_EnableSizing = &H800000
    OFN_UseMonikers = &H1000000
    OFN_DontAddToRecent = &H2000000
    OFN_ForceShowHidden = &H10000000        ' ǿ����ʾ�����ļ����ļ��У������ã�����������
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
'Public Enum PreviewPosition ' Ԥ��ͼƬ��λ��
'    ppNone = -1
'    ppTop = 0
'    ppLeft = 1
'    ppRight = 2
'    ppBottom = 3
'End Enum

' --- �ṹ�� ����
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

' --- ��Ա���� ����
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
Private m_blnReadOnlyChecked As Boolean ' ָʾ�Ƿ�ѡ��ֻ����ѡ���� MCDHook.bas ��ȡ�ã���

'Private m_picLogoPicture As PictureBox ' �����־ͼƬ��ͼƬ
'Private m_picPreviewPicture As PictureBox ' Ԥ��ͼƬ��ͼƬ
'Private m_ppLogoPosition As PreviewPosition ' �����־ͼƬ��λ��
'Private m_ppPreviewPosition As PreviewPosition ' Ԥ��ͼƬ��λ��

' ==== ����Ի��� ��������=================================================================
' ����Ի����Ա����
Private m_FontFlags As CF_Flags ' ����Ի����־
Private m_ReturnFont As StdFont ' �����������
' ==== ����Ի��� ��������=================================================================

' ==== ��ɫ�Ի��� ��������=================================================================
Private m_ColorFlags As CC_Flags    ' ��ɫ�Ի����־
Private m_ReturnColor As Long       ' ������ɫ����
' ==== ��ɫ�Ի��� ��������=================================================================

' --- ˽�б��� ����


' ==================================================================================
' === �������ԣ����Ĺ��ܣ�==========================================================
' ==================================================================================
' ���ԣ���ȡ����ť�Ƿ��������󣿣��ر�ע�⣺�ĸ��Ի����У���
Public Property Get CancelError() As Boolean
    'return object's CancelError property
    CancelError = m_blnCancelError
End Property
Public Property Let CancelError(ByVal bCancelError As Boolean)
    'assign object's CancelError property
    m_blnCancelError = bCancelError
End Property

' ���ԣ��Ƿ�ǿ����ʾ����ۣ����������� 76 ��ʽ��� 88 ��ʽ��ۣ�Win98���ܵ����޷���ʾ�Ի��򣡣�
Public Property Get ForceShowNewSkin() As Boolean
    ForceShowNewSkin = m_blnForceShowNewSkin
End Property
Public Property Let ForceShowNewSkin(ByVal vNewValue As Boolean)
    m_blnForceShowNewSkin = vNewValue
End Property

' ���ԣ��Ի�������λ�ã����ر�ע�⣺�ĸ��Ի����У���
Public Property Get dlgStartUpPosition() As StartUpPositionConstants
    dlgStartUpPosition = m_dlgStartUpPosition
End Property
Public Property Let dlgStartUpPosition(ByVal vNewValue As StartUpPositionConstants)
    m_dlgStartUpPosition = vNewValue
End Property
' ���ԣ��ļ���������·����
Public Property Get FileName() As String
    FileName = m_strFileName
End Property
Public Property Let FileName(ByVal sFilename As String)
    m_strFileName = sFilename
End Property

' ���ԣ��ļ��������ļ���+��׺�������� OFN_AllowMultiSelect �����ã�
Public Property Get FileTitle() As String
    FileTitle = m_strFileTitle
End Property
Public Property Let FileTitle(ByVal vNewValue As String)
    m_strFileTitle = vNewValue
End Property

' ���ԣ��ļ����͹��������б����� "ý���ļ�(*.piano)|*.piano|�����ļ�(*.*)|*.*"
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

' ���ԣ��ļ����͹��������б���ţ����ã�������
Public Property Get FilterIndex() As Long
    FilterIndex = m_lngFilterIndex
End Property
Public Property Let FilterIndex(ByVal lFilterIndex As Long)
    m_lngFilterIndex = lFilterIndex
End Property

' ���ԣ�����Ի���Ĭ�Ϻ�׺������ֻ������Ϊ�����ļ�ʱ��Ч������
Public Property Get DefaultExt() As String
    DefaultExt = m_strDefaultExt
End Property
Public Property Let DefaultExt(ByVal sDefaultExt As String)
    m_strDefaultExt = sDefaultExt
End Property

' ���ԣ��Ի�����⣨�ر�ע�⣺�ĸ��Ի����У���
Public Property Get DialogTitle() As String
    DialogTitle = m_strDialogTitle
End Property
Public Property Let DialogTitle(ByVal vNewValue As String)
    m_strDialogTitle = vNewValue
End Property

' ���ԣ��Ի��������Ϊ��־
Public Property Get Flags() As DialogFlags
    Flags = m_lngFlags
End Property
Public Property Let Flags(ByVal vNewValue As DialogFlags)
    m_lngFlags = vNewValue
End Property

' ���ԣ���������ر�ע�⣺�ĸ��Ի����У����ر�ע�⣺���жԻ���Ҫ����������ԣ�������������޷�ʹ�ã�����
Public Property Get hWnd() As Long
    hWnd = m_lngHwnd
End Property
Public Property Let hWnd(ByVal vNewValue As Long)
    m_lngHwnd = vNewValue
End Property

' ���ԣ�ʵ�����ر�ע�⣺�ĸ��Ի����У���
Public Property Get Instance() As Long
    Instance = m_lngInstance
End Property
Public Property Let Instance(ByVal vNewValue As Long)
    m_lngInstance = vNewValue
End Property

' ���ԣ��Ի����ʼ·��
Public Property Get InitDir() As String
    InitDir = m_strInitDir
End Property
Public Property Let InitDir(ByVal vNewValue As String)
    m_strInitDir = vNewValue
End Property

' ���ԣ�ָʾ�Ƿ�ѡ��ֻ����ѡ��Ĭ��ΪFalse
Public Property Get ReadOnlyChecked() As Boolean
    ReadOnlyChecked = MCDHook.IsReadOnlyChecked ' m_blnReadOnlyChecked
End Property
Public Property Let ReadOnlyChecked(ByVal vNewValue As Boolean)
    m_blnReadOnlyChecked = vNewValue
End Property

' ���ԣ������־ͼƬ��ͼƬ
Public Property Get LogoPicture() As PictureBox
    LogoPicture = m_picLogoPicture
End Property
Public Property Set LogoPicture(ByVal vNewValue As PictureBox)
    Set m_picLogoPicture = vNewValue
End Property
' ���ԣ������־ͼƬ��λ��
Public Property Get ppLogoPosition() As PreviewPosition
    ppLogoPosition = m_ppLogoPosition
End Property
Public Property Let ppLogoPosition(ByVal vNewValue As PreviewPosition)
    m_ppLogoPosition = vNewValue
End Property

' ���ԣ�Ԥ��ͼƬ��ͼƬ
Public Property Get PreviewPicture() As PictureBox
    PreviewPicture = m_picPreviewPicture
End Property
Public Property Set PreviewPicture(ByVal vNewValue As PictureBox)
    Set m_picPreviewPicture = vNewValue
End Property

' ���ԣ�Ԥ��ͼƬ��λ��
Public Property Get ppPreviewPosition() As PreviewPosition
    ppPreviewPosition = m_ppPreviewPosition
End Property
Public Property Let ppPreviewPosition(ByVal vNewValue As PreviewPosition)
    m_ppPreviewPosition = vNewValue
End Property

' ==== ����Ի��� ��������=================================================================
' ���ԣ�����Ի��򷵻�ֵ
Public Property Get getReturnFont() As StdFont
    Set getReturnFont = m_ReturnFont
End Property
Public Property Let getReturnFont(vNewValue As StdFont)
    Set m_ReturnFont = vNewValue
End Property
' ���ԣ��Ի��������Ϊ��־
Public Property Get FontFlags() As CF_Flags
    FontFlags = m_FontFlags
End Property
Public Property Let FontFlags(ByVal vNewValue As CF_Flags)
    m_FontFlags = vNewValue
End Property
' ==== ����Ի��� ��������=================================================================

' ==== ��ɫ�Ի��� ��������=================================================================
' ���ԣ�����ѡ�����ɫ
Public Property Get getReturnColor() As Long
    getReturnColor = m_ReturnColor
End Property
Public Property Let getReturnColor(vNewValue As Long)
    m_ReturnColor = vNewValue
End Property
' ���ԣ��Ի��������Ϊ��־
Public Property Get ColorFlags() As CC_Flags
    ColorFlags = m_ColorFlags
End Property
Public Property Let ColorFlags(ByVal vNewValue As CC_Flags)
    m_ColorFlags = vNewValue
End Property
' ==== ��ɫ�Ի��� ��������=================================================================
' ==================================================================================
' === �������ԣ����Ĺ��ܣ�==========================================================
' ==================================================================================


' ==================================================================================
' === �������������Ĺ��ܣ�==========================================================
' ==================================================================================
Public Sub ShowProperty(strFilePath As String)
    ' ##############################################################################
    ' ��  �ܣ�ȡ�ø����ļ� ���ļ��е����ԣ���ʾ���ԶԻ���
    ' ��������ShowProperty
    ' ��  ����strFilePath ���ַ������ļ� ���ļ���·��
    ' ����ֵ����
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
    ' ��  �ܣ���ʾ�򿪶Ի���
    ' ��������ShowOpen
    ' ��  ������
    ' ����ֵ���߼�ֵ���Ƿ�ɹ����أ���û��ȡ����ť������
    ' ##############################################################################
    On Error Resume Next
    Dim OFN As OPENFILENAME
    Dim lFileSize As Long

    ' ��ʼ���ṹ�� OFN
    With OFN
        .lpfnHook = VBGetProcAddress(AddressOf CDCallBackFun) ' �ص�������ַ������������
        ' �Ƿ�ǿ����ʾ����ۣ����������� 76 ��ʽ��� 88 ��ʽ��ۣ�Win98���ܵ����޷���ʾ�Ի��򣡣�
        If m_blnForceShowNewSkin Then
            .lStructSize = 88
        Else
            .lStructSize = Len(OFN)
        End If
        .hwndOwner = m_lngHwnd
        .hInstance = m_lngInstance
        .lFlags = m_lngFlags
        .lpstrTitle = m_strDialogTitle
        ' ��������ѡʱ���������� OFN_Explorer ���������������������
        ' �����������ǿ����ʾ�����ʱ����ѡ��������Ϊ���������޷���ʾ�Ի��򣡣���
        If (.lFlags And OFN_AllowMultiSelect) = OFN_AllowMultiSelect Then
            .lFlags = .lFlags Or OFN_Explorer
            lFileSize = IIf(m_blnForceShowNewSkin, 1024, 8192)  ' ǿ����ʾ������ǵĻ�������������
            ' 8192 ���ֵ�����������޷���ʾ�Ի��򣡣���������1024ʱ��ͼƬ���ڵ׶��޷���λ������������
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

        ' ���ñ�־�Ƿ�ѡ��ֻ����ѡ��
        If m_blnReadOnlyChecked Then .lFlags = .lFlags Or OFN_ReadOnly
        
        ' ��ʾ�򿪶Ի��򣿣�
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
                Err.Raise cdlCancel, App.EXEName & ".CCommonDialog", "�û�ѡ����ȡ����"
            End If
        End If
    End With
End Function
Public Function ParseMultiFileName(ByRef sDir As String, ByRef sFiles() As String, ByRef lFileCount As Long) As Boolean
    ' ##############################################################################
    ' ��  �ܣ�ȡ�ô򿪶Ի���ѡ�����ļ�ʱ�������ļ���
    ' ��������ParseMultiFileName
    ' ��  ����sDir���ַ������ļ�Ŀ¼��sFiles �ַ����飺�ļ�����lFileCount �����ͣ��ļ���Ŀ
    ' ����ֵ���߼�ֵ���Ƿ�ѡ�ж���ļ���ע�⣺����Ҳ������ֵ����
    ' ##############################################################################
    Dim lPos As Long
    Dim lNextPos As Long
    Dim sAllFiles As String
    On Error Resume Next

    lPos = InStr(m_strFileName, vbNullChar & vbNullChar)
    sAllFiles = Left$(m_strFileName, lPos - 1)
    lNextPos = InStr(sAllFiles, vbNullChar)
    If lNextPos <> 0 Then ' ѡ�˶���ļ���
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
    Else ' ѡ�����ļ�
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
    ' ��  �ܣ���ʾ����Ի���
    ' ��������ShowSave
    ' ��  ������
    ' ����ֵ���߼�ֵ���Ƿ�ɹ����أ���û��ȡ����ť������
    ' ##############################################################################
    On Error Resume Next
    Dim OFN As OPENFILENAME
    With OFN
        .lpfnHook = VBGetProcAddress(AddressOf CDCallBackFun) ' �ص�������ַ������������
        ' �Ƿ�ǿ����ʾ����ۣ����������� 76 ��ʽ��� 88 ��ʽ��ۣ�Win98���ܵ����޷���ʾ�Ի��򣡣�
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
        .lpstrDefExt = m_strDefaultExt & vbNullChar & vbNullChar ' ShowOpen ����Ҫ��ֵ��

        ' ��ʾ����Ի��򣿣�
        WhichStyle = ssSave
        If GetSaveFileName(OFN) Then
            ShowSave = True
            m_strFileName = TrimNull(.lpstrFile)
            m_strFileTitle = .lpstrFileTitle
        Else
            ShowSave = False
            If m_blnCancelError Then
                Err.Raise cdlCancel, App.EXEName & ".CCommonDialog", "�û�ѡ����ȡ����"
            End If
        End If
    End With
End Function

' =========================================================================================
' ==== �Զ���Ի����ϵĿؼ����ء���ʾ���ı����ֵȣ�����ȥ����==============================
' =========================================================================================
' ���ء���ʾ�򿪡�����Ի����ϵĿؼ�
Public Sub HideDlgControls(ByVal Index As Integer)
    m_blnHideControls(Index) = True
End Sub
Public Sub ShowDlgControls(ByVal Index As Integer)
    m_blnHideControls(Index) = False
End Sub
' ���öԻ����ϵĿؼ������֡�m_strControlsCaption(I) ������ֵ��Ĭ��Ϊ�գ����ı�ԭʼֵ��
Public Sub SetDlgControlsCaption(ByVal Index As Integer, strNewCaption As String)
    m_strControlsCaption(Index) = strNewCaption
End Sub
' =========================================================================================
' ==== �Զ���Ի����ϵĿؼ����ء���ʾ���ı����ֵȣ�����ȥ����==============================
' =========================================================================================


' ==== ����Ի��� ��������=================================================================
Public Function ShowFont(Optional ByRef FontRGB As Long) As Boolean
' ע�⣺����ַ���ݲ�����������ֵ����ȡ��ѡ���������ɫ������
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

    ' ������Ϣ������
    Dim mRGBResult As Long      ' ������ɫ����
    Dim mFontName As String     ' ��������
    Dim mItalic As Boolean      ' �Ƿ�б��
    Dim mUnderline As Boolean   ' �Ƿ��»���
    Dim mStrikethru As Boolean  ' �Ƿ�ɾ����
    Dim mFontSize As Long       ' �����С
    Dim mBold As Boolean        ' �Ƿ�Ӵ�
    Dim mCharSet As Integer     ' �ַ��������ģ����ģ���
    Dim mWeight As Integer
    
    ' ���öԻ����ʼ���塣
    mFontName = m_ReturnFont.Name
    mFontSize = m_ReturnFont.Size
    mBold = m_ReturnFont.Bold
    mCharSet = m_ReturnFont.Charset
    
    TempByteArray = StrConv(mFontName & vbNullChar, vbFromUnicode)
    ByteArrayLimit = UBound(TempByteArray)
    ' ���������ṹ���ֵ��
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
    
    ' ��ʾ����Ի��򣿣�
    WhichStyle = ssFont
    RetValue = ChooseFont(CF)
    
    ' �ж�����Ի��򷵻�ֵ���Ƿ�ɹ���
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
        ' ����������Ϣ����Ա����
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
        ' ע�⣺��ɫ���ڱ�׼����ṹ����û�У���
        FontRGB = mRGBResult
    Else
        ShowFont = False
        If m_blnCancelError Then
            Err.Raise cdlCancel, App.EXEName & ".CCommonDialog", "�û�ѡ����ȡ����"
        End If
    End If
   
End Function
' ==== ����Ի��� ��������=================================================================

' ==== ��ɫ�Ի��� ��������=================================================================
Public Function ShowColor(Optional ByRef CError As Long, Optional ByRef ExCError As Long, Optional ByVal lInitColor As Long = vbBlue) As Boolean
' ע�⣺���� CError ����ַ���ݲ�����������ֵ�����ش���ţ���
' ShowColor ��������ֵ���ɹ���
' Ĭ�Ϸ��ص���ɫֵ����ʼ���Ի���ʱ����ɫ����ɫ��
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
    
    ' ��ʾ��ɫ�Ի��򣿣�
    WhichStyle = ssColor
    result = ChooseColor(C)
    CError = CommDlgExtendedError()

    ' �û�ѡ����ȡ��
    If result = 0 Then
        If m_blnCancelError Then
            Err.Raise cdlCancel, App.EXEName & ".CCommonDialog", "�û�ѡ����ȡ����"
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
    m_ReturnColor = C.rgbResult  ' ����ѡ�����ɫֵ��
    ShowColor = True ' �����ɹ������أ�
    
    Rem ----------------------------------------------------
    Rem This is custom colors (not used here)
    Rem ----------------------------------------------------
    'For i& = 0 To UBound(ClrArray)
    '    Debug.Print "Custom Color"; Str$(i&); ":", Hex$(ClrArray(i&))
    'Next
    
End Function
' ==== ��ɫ�Ի��� ��������=================================================================

' ==================================================================================
' === �������������Ĺ��ܣ�==========================================================
' ==================================================================================


' ==================================================================================
' === ˽�к������������ܣ�==========================================================
' ==================================================================================
Private Function TrimNull(ByVal strS As String) As String
    ' ##############################################################################
    ' ��  �ܣ�ȥ���ַ���β vbNullChar==Chr$(0)
    ' ��������TrimNull
    ' ��  ����strS ���ַ�����Ԥ�������ַ���
    ' ����ֵ���ַ�����ȥ���ַ���β vbNullChar �����ַ���
    ' ##############################################################################
    Dim pos As Long
    pos = InStr(strS, Chr$(0))
    If pos Then strS = Left$(strS, pos - 1)
    TrimNull = strS
End Function

Private Function VBGetProcAddress(ByVal lpfnAddressOfX As Long) As Long
    ' ##############################################################################
    ' ��  �ܣ�ȡ�ú����ĵ�ַ
    ' ��������VBGetProcAddress
    ' ��  ����AddressOf ��ȡ�õ�ַ�ĺ�����
    ' ����ֵ�������ͣ������ĵ�ַ
    ' ##############################################################################
    VBGetProcAddress = lpfnAddressOfX
End Function
' ==================================================================================
' === ˽�к������������ܣ�==========================================================
' ==================================================================================



' ==================================================================================
' ��ĳ�ʼ��������
Private Sub Class_Initialize()
    ' ��Ա���������ԣ���ʼֵ�趨����Щʹ��Ĭ��ֵ��
    m_blnForceShowNewSkin = True ' �Ƿ�ǿ����ʾ����ۣ�
    m_lngFlags = OFN_Explorer Or OFN_EnableHook  ' �Ի����� ��Դ�����������ʹ�ûص�������
    m_strFilter = "�ı��ļ�(*.txt)|*.txt|�����ļ�(*.*)|*.*" ' �ļ����͹��������б�����
    
    ' MCDHook �еı�����
    Set m_picLogoPicture = Nothing ' �����־ͼƬ��ͼƬ
    Set m_picPreviewPicture = Nothing ' Ԥ��ͼƬ��ͼƬ
    m_ppLogoPosition = ppTop ' �����־ͼƬ��λ��
    m_ppPreviewPosition = ppBottom ' Ԥ��ͼƬ��λ��
    m_dlgStartUpPosition = vbStartUpScreen ' �Ի�������λ�ã���Ļ����
    
    ' �Ƿ����ضԻ����ϵĿؼ�������ȥ����
    Dim I As Integer
    For I = 0 To 8
        m_blnHideControls(I) = False
    Next I
    ' ���öԻ����ϵĿؼ������֣�
    For I = 0 To 6
        m_strControlsCaption(I) = ""
    Next I
    ' ==== ����Ի��� ��������=================================================================
    Set m_ReturnFont = New StdFont ' �����������
    With m_ReturnFont
        .Bold = False
        .Charset = 134 ' �����ַ���
        .Italic = False
        .Name = "����"
        .Size = 12
        .Strikethrough = False
        .Underline = False
        .Weight = 520
    End With
    m_FontFlags = CF_INITTOLOGFONTSTRUCT Or CF_ENABLEHOOK Or CF_EFFECTS Or CF_APPLY 'Or CF_SHOWHELP ' ����Ի����־
    ' ==== ����Ի��� ��������=================================================================
    ' ==== ��ɫ�Ի��� ��������=================================================================
    m_ColorFlags = CC_ENABLEHOOK Or CC_RGBINIT Or CC_FULLOPEN
    m_ReturnColor = vbBlue ' Ĭ�Ϸ��ص���ɫֵ����ɫ��
    ' ==== ��ɫ�Ի��� ��������=================================================================
End Sub
Private Sub Class_Terminate()
    Set m_picLogoPicture = Nothing ' �����־ͼƬ��ͼƬ
    Set m_picPreviewPicture = Nothing ' Ԥ��ͼƬ��ͼƬ
    Set m_ReturnFont = Nothing
End Sub
' ��ĳ�ʼ��������
' ==================================================================================