VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMultiToolTips"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' CMultiToolTips.cls
' ��  �ܣ�Ϊ�ؼ����ö���������ʾ
' ʹ  �ã�1������ hWnd Ϊ����ʾ��ʾ�Ŀؼ������
'         2������ CreateToolTip ������������ TTStyle ���ԡ�
'         ע�⣺������������ڲ���2֮ǰ���ã���Ȼ������Ҳ���Ըı䣬���������������仯��
' --- API ��������
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long '��������
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long '������Ϣ
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()

' --- ���� ����
Private Const WM_USER = &H400
Private Const CW_USEDEFAULT = &H80000000
Private Const TTDT_AUTOPOP = 2
Private Const TTDT_INITIAL = 3

' ToolTip ����
Private Const TOOLTIPS_CLASSA = "tooltips_class32"

' Tooltip ����
Private Const TTS_ALWAYSTIP = &H1                   ' ����
Private Const TTS_NOPREFIX = &H2                    ' ��׼ ������ʾ
Private Const TTS_BALLOON = &H40                    ' ���� ������ʾ

'TooltipInfo Flags(���кܶ࣬һ��ֻ��Ҫ�������������㹻��)
Private Const TTF_CENTERTIP = &H2                   ' ����
Private Const TTF_SUBCLASS = &H10                   ' ��ʼ
Private Const TTF_TRANSPARENT = &H100               ' ͸��

' ToolTip ��Ϣ(ͨ��Sendmessage API���е���)
Private Const TTM_ACTIVATE = WM_USER + 1            ' ����
Private Const TTM_ADDTOOLA = (WM_USER + 4)          ' ����Tooltip
Private Const TTM_SETMAXTIPWIDTH = (WM_USER + 24)   ' ����������
Private Const TTM_SETTIPBKCOLOR = (WM_USER + 19)    ' ���ñ�����ɫ
Private Const TTM_SETTIPTEXTCOLOR = (WM_USER + 20)  ' �����ı���ɫ
Private Const TTM_SETTITLE = (WM_USER + 32)         ' ���ñ���
Private Const TTM_SETDELAYTIME = (WM_USER + 3)      ' ����Tooltipʱ��
Private Const TTM_UPDATETIPTEXTA = (WM_USER + 12)   ' ������ʾ�ı�

' --- ö�� ����
' Tooltip ��ʾ��ʽ
Public Enum ToolTipShowStyle
    TT_Standard = TTS_NOPREFIX              ' ��׼ ������ʾ
    TT_Balloon = TTS_BALLOON                ' ���� ������ʾ
End Enum

' Tooltipͼ����ʽ
Public Enum ToolTipIconStyle
    TTI_NONE = 0                           ' ��ͼ��
    TTI_INFO = 1                           ' ��Ϣ
    TTI_WARNING = 2                        ' ����
    TTI_ERROR = 3                          ' ����
    TTI_GUID = &H4                         ' ��֪��ʲô
End Enum

' --- �ṹ�� ����
' ����
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
' Tooltip ���ͣ�VB�Դ���API�����û�У���Commctl.h����ȡ����
Private Type TOOLINFO
    cbSize As Long
    lFlags As Long
    lHwnd As Long
    lId As Long
    lpRect As RECT
    hInstance As Long
    lpszText As String
    lParam As Long
End Type

' --- ��Ա���� ����
Private m_lngHwnd As Long                   ' ����ʾToolTip�Ŀؼ����

Private m_ToolTipIcon As ToolTipIconStyle   ' ͼ��
Private m_ToolTipTitle As String            ' ��������
Private m_ToolTipText As String             ' ��ʾ�ı�

Private m_BackColor As OLE_COLOR            ' ����ɫ
Private m_ForeColor As OLE_COLOR            ' ǰ�������壩��ɫ

Private m_TimeToStay As Long                ' ToolTip��ʾʱ�䣬ͣ����ʱ�䣡��λ�����룡
Private m_TimeInterval As Long              ' ��ʾToolTip��ʱ��ʱ��

Private m_TTStyle As ToolTipShowStyle       ' Tooltip ��ʾ��ʽ

' --- ˽�б��� ����
Private lngToolTipHwnd As Long              ' ���� ToolTip ���ڵľ��
' ToolTipText ����������ʾ�ı� ���������Ҫ�õ������������ĳ�ʼ���� CreateToolTip �����У�
' SendMessage lngToolTipHwnd, TTM_UPDATETIPTEXTA, m_ToolTipIcon, ByVal m_ToolTipText �޷����£�����
Private lpToolInfo As TOOLINFO


' ###################################################################################################
' ### �������� ###
' ###################################################################################################
' ���ԣ������ע�⣺Ҫ����������ԣ����������޷�ʹ�ã�����
Public Property Get hWnd() As Long
    hWnd = m_lngHwnd
End Property
Public Property Let hWnd(ByVal NewValue As Long)
    m_lngHwnd = NewValue
End Property

' ���ԣ�ToolTip ͼ��
Public Property Get ToolTipIcon() As ToolTipIconStyle
    ToolTipIcon = m_ToolTipIcon
End Property
Public Property Let ToolTipIcon(ByVal NewValue As ToolTipIconStyle)
    m_ToolTipIcon = NewValue
    ' ��������ͼ�꣡
    SendMessage lngToolTipHwnd, TTM_SETTITLE, m_ToolTipIcon, ByVal m_ToolTipTitle
End Property

' ���ԣ�tooltip�ı���
Public Property Get ToolTipTitle() As String
    ToolTipTitle = m_ToolTipTitle
End Property
Public Property Let ToolTipTitle(ByVal NewValue As String)
    m_ToolTipTitle = NewValue
    ' �������±��⣡
    SendMessage lngToolTipHwnd, TTM_SETTITLE, m_ToolTipIcon, ByVal m_ToolTipTitle
End Property

'���ԣ�tooltip���ı�(֧�ֶ���)
Public Property Get ToolTipText() As String
    ToolTipText = m_ToolTipText
End Property
Public Property Let ToolTipText(ByVal NewValue As String)
    m_ToolTipText = NewValue
    lpToolInfo.lpszText = NewValue ' �ر�ע�⣺�ı��ֵ�����ܸ��£�����
    ' ����������ʾ�ı���
    SendMessage lngToolTipHwnd, TTM_UPDATETIPTEXTA, m_ToolTipIcon, lpToolInfo ' ByVal m_ToolTipText
End Property

' ���ԣ�tooltip ����ɫ
Public Property Get BackColor() As OLE_COLOR
    BackColor = m_BackColor
End Property
Public Property Let BackColor(ByVal NewValue As OLE_COLOR)
    m_BackColor = NewValue
    ' �������±���ɫ��
    SendMessage lngToolTipHwnd, TTM_SETTIPBKCOLOR, m_BackColor, ByVal 0&            ' ������ɫ
End Property

' ���ԣ�tooltip ǰ�������壩��ɫ
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = m_ForeColor
End Property
Public Property Let ForeColor(ByVal NewValue As OLE_COLOR)
    m_ForeColor = NewValue
    ' �������±���ɫ��
    SendMessage lngToolTipHwnd, TTM_SETTIPTEXTCOLOR, m_ForeColor, ByVal 0&          ' ǰ�������壩��ɫ
End Property

' ���ԣ�ToolTip��ʾʱ�䣬ͣ����ʱ�䣡
Public Property Get TimeToStay() As Long
    TimeToStay = m_TimeToStay
End Property
Public Property Let TimeToStay(ByVal NewValue As Long)
    m_TimeToStay = NewValue
    ' �������� ToolTip��ʾʱ�䣬ͣ����ʱ�䣡��
    SendMessage lngToolTipHwnd, TTM_SETDELAYTIME, TTDT_AUTOPOP, ByVal m_TimeToStay    ' ����ToolTip��ʾʱ��
End Property

' ���ԣ���ʾToolTip��ʱ��ʱ��
Public Property Get TimeInterval() As Long
    TimeInterval = m_TimeInterval
End Property
Public Property Let TimeInterval(ByVal NewValue As Long)
    m_TimeInterval = NewValue
    ' �������� ��ʾToolTip��ʱ��ʱ�䣡
    SendMessage lngToolTipHwnd, TTM_SETDELAYTIME, TTDT_INITIAL, ByVal m_TimeInterval ' ������ʾToolTipʱ��ʱ��
End Property

' ���ԣ�Tooltip ��ʾ��ʽ
Public Property Get TTStyle() As ToolTipShowStyle
    TTStyle = m_TTStyle
End Property
Public Property Let TTStyle(ByVal NewValue As ToolTipShowStyle)
    m_TTStyle = NewValue
    ' �������� Tooltip ��ʾ��ʽ���ڴ˵��� ���� ToolTip ���� �������������Զ����£�����
    CreateToolTip m_TTStyle
End Property
' ###################################################################################################
' ### �������� ###
' ###################################################################################################


' ###################################################################################################
' ### �������� ###
' ###################################################################################################
' ���� ToolTip ���ڣ��������ش������ڵľ����
Public Function CreateToolTip(Optional ByVal opTTStyle As ToolTipShowStyle = TT_Balloon) As Long
    Dim lpRect As RECT
    'Dim lpToolInfo As TOOLINFO
    
    ' �����ٿ����еĴ���
    DestroyWindow lngToolTipHwnd
    
     ' �ڴ���tooltips_class32ʱ������ó�ʼ����
    ' �������Ҿ��ñ���ע�⣬��Ȼ�������API����ʼ��û��������ʱ����VBIDE�����е��Ե�ʱ��һ��������
    ' ���Ǳ����Ժ����ǳ��ֲ���ToolTip�������������˺ܾã�����������˹����ⷽ������ϣ�
    ' ��֪���ڴ���ToolTipʱ�������Comctl32.dll�е�InitCommonControl API�������г�ʼ�������ƣ������ڲŵ���ͨ����
    InitCommonControls
    
    ' ���ش��ڵľ��������ToolTip����
    lngToolTipHwnd = CreateWindowEx(0, TOOLTIPS_CLASSA, vbNullString, TTS_ALWAYSTIP Or opTTStyle, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, m_lngHwnd, 0, App.hInstance, ByVal 0&)
    CreateToolTip = lngToolTipHwnd
    
    ' ��ȡ��������
    GetClientRect m_lngHwnd, lpRect
    
    ' ��ʼ���ṹ��
    With lpToolInfo
        .cbSize = Len(lpToolInfo)       ' ��С
        .lFlags = TTF_SUBCLASS          ' Tooltip��ʽ
        .lHwnd = m_lngHwnd              ' ����ʾToolTip�Ŀؼ����
        .hInstance = App.hInstance      ' ���ʵ��
        .lpszText = m_ToolTipText       ' ToolTip�ı�
        .lId = 0                        ' idΪ��
        .lpRect = lpRect                ' ���ھ���
    End With
    
    ' ���� ToolTip
    SendMessage lngToolTipHwnd, TTM_ADDTOOLA, 0, lpToolInfo
    
    'ToolTip ��ɫ���ã�vbBlue��vbWhite��
    SendMessage lngToolTipHwnd, TTM_SETTIPTEXTCOLOR, m_ForeColor, ByVal 0&          ' ǰ�������壩��ɫ
    SendMessage lngToolTipHwnd, TTM_SETTIPBKCOLOR, m_BackColor, ByVal 0&            ' ������ɫ
    SendMessage lngToolTipHwnd, TTM_SETTITLE, m_ToolTipIcon, ByVal m_ToolTipTitle   ' ����
    
    ' ������ʾʱ��
    SendMessage lngToolTipHwnd, TTM_SETDELAYTIME, TTDT_AUTOPOP, ByVal m_TimeToStay  ' ����ToolTip��ʾʱ��
    SendMessage lngToolTipHwnd, TTM_SETDELAYTIME, TTDT_INITIAL, ByVal m_TimeInterval ' ������ʾToolTipʱ��ʱ��

End Function
' ###################################################################################################
' ### �������� ###
' ###################################################################################################



' ###################################################################################################
' ### ���˽�з��� ###
' ###################################################################################################
' ���ʼ��ʱ�����Ա����ĳ�ʼ��������Ĭ�����ԣ�
Private Sub Class_Initialize()
    m_lngHwnd = 0                                       ' ����ʾToolTip�Ŀؼ����
    m_ToolTipIcon = TTI_INFO                            ' ͼ��
    m_ToolTipTitle = "DyEncryptor"                 ' ��������
    m_ToolTipText = "DyEncryptor" ' ��ʾ�ı�
    
    m_BackColor = vbWhite                               ' ����ɫ
    m_ForeColor = vbBlue                                ' ǰ�������壩��ɫ
    
    m_TimeToStay = 8888&                                ' ToolTip��ʾʱ�䣬ͣ����ʱ�䣡��λ�����룡
    m_TimeInterval = 222&                               ' ��ʾToolTip��ʱ��ʱ��
    
    m_TTStyle = TT_Balloon                              ' Tooltip ��ʾ��ʽ
End Sub
' ������ʱ��
Private Sub Class_Terminate()
    ' ����õ�ʱ����������ٴ���
    DestroyWindow lngToolTipHwnd
End Sub
' ###################################################################################################
' ### ���˽�з��� ###
' ###################################################################################################
