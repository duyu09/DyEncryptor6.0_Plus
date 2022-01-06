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
' 功  能：为控件设置多行气泡提示
' 使  用：1、设置 hWnd 为欲显示提示的控件句柄；
'         2、调用 CreateToolTip 函数，或设置 TTStyle 属性。
'         注意：其他属性最好在步骤2之前设置！当然，后面也可以改变，并可以立即看到变化！
' --- API 函数申明
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long '创建窗口
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long '发出消息
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()

' --- 常数 申明
Private Const WM_USER = &H400
Private Const CW_USEDEFAULT = &H80000000
Private Const TTDT_AUTOPOP = 2
Private Const TTDT_INITIAL = 3

' ToolTip 类名
Private Const TOOLTIPS_CLASSA = "tooltips_class32"

' Tooltip 参数
Private Const TTS_ALWAYSTIP = &H1                   ' 永久
Private Const TTS_NOPREFIX = &H2                    ' 标准 矩形提示
Private Const TTS_BALLOON = &H40                    ' 球形 气泡提示

'TooltipInfo Flags(还有很多，一般只需要这两个函数就足够了)
Private Const TTF_CENTERTIP = &H2                   ' 居中
Private Const TTF_SUBCLASS = &H10                   ' 初始
Private Const TTF_TRANSPARENT = &H100               ' 透明

' ToolTip 消息(通过Sendmessage API进行调用)
Private Const TTM_ACTIVATE = WM_USER + 1            ' 激活
Private Const TTM_ADDTOOLA = (WM_USER + 4)          ' 添加Tooltip
Private Const TTM_SETMAXTIPWIDTH = (WM_USER + 24)   ' 设置最大宽度
Private Const TTM_SETTIPBKCOLOR = (WM_USER + 19)    ' 设置背景颜色
Private Const TTM_SETTIPTEXTCOLOR = (WM_USER + 20)  ' 设置文本颜色
Private Const TTM_SETTITLE = (WM_USER + 32)         ' 设置标题
Private Const TTM_SETDELAYTIME = (WM_USER + 3)      ' 设置Tooltip时间
Private Const TTM_UPDATETIPTEXTA = (WM_USER + 12)   ' 更新提示文本

' --- 枚举 申明
' Tooltip 显示样式
Public Enum ToolTipShowStyle
    TT_Standard = TTS_NOPREFIX              ' 标准 矩形提示
    TT_Balloon = TTS_BALLOON                ' 球形 气泡提示
End Enum

' Tooltip图标样式
Public Enum ToolTipIconStyle
    TTI_NONE = 0                           ' 无图标
    TTI_INFO = 1                           ' 信息
    TTI_WARNING = 2                        ' 警告
    TTI_ERROR = 3                          ' 错误
    TTI_GUID = &H4                         ' 不知是什么
End Enum

' --- 结构体 申明
' 矩形
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
' Tooltip 类型，VB自带的API浏览器没有，从Commctl.h中提取出来
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

' --- 成员变量 申明
Private m_lngHwnd As Long                   ' 欲显示ToolTip的控件句柄

Private m_ToolTipIcon As ToolTipIconStyle   ' 图标
Private m_ToolTipTitle As String            ' 标题文字
Private m_ToolTipText As String             ' 提示文本

Private m_BackColor As OLE_COLOR            ' 背景色
Private m_ForeColor As OLE_COLOR            ' 前景（字体）颜色

Private m_TimeToStay As Long                ' ToolTip显示时间，停留的时间！单位：毫秒！
Private m_TimeInterval As Long              ' 显示ToolTip的时隔时间

Private m_TTStyle As ToolTipShowStyle       ' Tooltip 显示样式

' --- 私有变量 声明
Private lngToolTipHwnd As Long              ' 创建 ToolTip 窗口的句柄
' ToolTipText 立即更新提示文本 这个属性需要用到它！！！它的初始化在 CreateToolTip 函数中！
' SendMessage lngToolTipHwnd, TTM_UPDATETIPTEXTA, m_ToolTipIcon, ByVal m_ToolTipText 无法更新！！！
Private lpToolInfo As TOOLINFO


' ###################################################################################################
' ### 公共属性 ###
' ###################################################################################################
' 属性：句柄，注意：要设置这个属性！！！否则无法使用！！！
Public Property Get hWnd() As Long
    hWnd = m_lngHwnd
End Property
Public Property Let hWnd(ByVal NewValue As Long)
    m_lngHwnd = NewValue
End Property

' 属性：ToolTip 图标
Public Property Get ToolTipIcon() As ToolTipIconStyle
    ToolTipIcon = m_ToolTipIcon
End Property
Public Property Let ToolTipIcon(ByVal NewValue As ToolTipIconStyle)
    m_ToolTipIcon = NewValue
    ' 立即更新图标！
    SendMessage lngToolTipHwnd, TTM_SETTITLE, m_ToolTipIcon, ByVal m_ToolTipTitle
End Property

' 属性：tooltip的标题
Public Property Get ToolTipTitle() As String
    ToolTipTitle = m_ToolTipTitle
End Property
Public Property Let ToolTipTitle(ByVal NewValue As String)
    m_ToolTipTitle = NewValue
    ' 立即更新标题！
    SendMessage lngToolTipHwnd, TTM_SETTITLE, m_ToolTipIcon, ByVal m_ToolTipTitle
End Property

'属性：tooltip的文本(支持多行)
Public Property Get ToolTipText() As String
    ToolTipText = m_ToolTipText
End Property
Public Property Let ToolTipText(ByVal NewValue As String)
    m_ToolTipText = NewValue
    lpToolInfo.lpszText = NewValue ' 特别注意：改变此值，才能更新！！！
    ' 立即更新提示文本！
    SendMessage lngToolTipHwnd, TTM_UPDATETIPTEXTA, m_ToolTipIcon, lpToolInfo ' ByVal m_ToolTipText
End Property

' 属性：tooltip 背景色
Public Property Get BackColor() As OLE_COLOR
    BackColor = m_BackColor
End Property
Public Property Let BackColor(ByVal NewValue As OLE_COLOR)
    m_BackColor = NewValue
    ' 立即更新背景色！
    SendMessage lngToolTipHwnd, TTM_SETTIPBKCOLOR, m_BackColor, ByVal 0&            ' 背景颜色
End Property

' 属性：tooltip 前景（字体）颜色
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = m_ForeColor
End Property
Public Property Let ForeColor(ByVal NewValue As OLE_COLOR)
    m_ForeColor = NewValue
    ' 立即更新背景色！
    SendMessage lngToolTipHwnd, TTM_SETTIPTEXTCOLOR, m_ForeColor, ByVal 0&          ' 前景（字体）颜色
End Property

' 属性：ToolTip显示时间，停留的时间！
Public Property Get TimeToStay() As Long
    TimeToStay = m_TimeToStay
End Property
Public Property Let TimeToStay(ByVal NewValue As Long)
    m_TimeToStay = NewValue
    ' 立即更新 ToolTip显示时间，停留的时间！！
    SendMessage lngToolTipHwnd, TTM_SETDELAYTIME, TTDT_AUTOPOP, ByVal m_TimeToStay    ' 设置ToolTip显示时间
End Property

' 属性：显示ToolTip的时隔时间
Public Property Get TimeInterval() As Long
    TimeInterval = m_TimeInterval
End Property
Public Property Let TimeInterval(ByVal NewValue As Long)
    m_TimeInterval = NewValue
    ' 立即更新 显示ToolTip的时隔时间！
    SendMessage lngToolTipHwnd, TTM_SETDELAYTIME, TTDT_INITIAL, ByVal m_TimeInterval ' 设置显示ToolTip时隔时间
End Property

' 属性：Tooltip 显示样式
Public Property Get TTStyle() As ToolTipShowStyle
    TTStyle = m_TTStyle
End Property
Public Property Let TTStyle(ByVal NewValue As ToolTipShowStyle)
    m_TTStyle = NewValue
    ' 立即更新 Tooltip 显示样式，在此调用 创建 ToolTip 窗口 函数，许多属性都更新！！！
    CreateToolTip m_TTStyle
End Property
' ###################################################################################################
' ### 公共属性 ###
' ###################################################################################################


' ###################################################################################################
' ### 公共方法 ###
' ###################################################################################################
' 创建 ToolTip 窗口（函数返回创建窗口的句柄）
Public Function CreateToolTip(Optional ByVal opTTStyle As ToolTipShowStyle = TT_Balloon) As Long
    Dim lpRect As RECT
    'Dim lpToolInfo As TOOLINFO
    
    ' 先销毁可能有的窗口
    DestroyWindow lngToolTipHwnd
    
     ' 在创建tooltips_class32时必须调用初始化。
    ' （这里我觉得必须注意，必然加上这个API，开始我没有添加上时，在VBIDE环境中调试的时候一切正常，
    ' 可是编译以后老是出现不了ToolTip，这里我郁闷了很久，最后到网上找了关于这方面的资料，
    ' 才知道在创建ToolTip时必须调用Comctl32.dll中的InitCommonControl API函数进行初始化。郁闷，到现在才调试通过）
    InitCommonControls
    
    ' 返回窗口的句柄并创建ToolTip窗口
    lngToolTipHwnd = CreateWindowEx(0, TOOLTIPS_CLASSA, vbNullString, TTS_ALWAYSTIP Or opTTStyle, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, m_lngHwnd, 0, App.hInstance, ByVal 0&)
    CreateToolTip = lngToolTipHwnd
    
    ' 获取窗户矩形
    GetClientRect m_lngHwnd, lpRect
    
    ' 初始化结构体
    With lpToolInfo
        .cbSize = Len(lpToolInfo)       ' 大小
        .lFlags = TTF_SUBCLASS          ' Tooltip样式
        .lHwnd = m_lngHwnd              ' 欲显示ToolTip的控件句柄
        .hInstance = App.hInstance      ' 句柄实例
        .lpszText = m_ToolTipText       ' ToolTip文本
        .lId = 0                        ' id为空
        .lpRect = lpRect                ' 窗口矩形
    End With
    
    ' 添加 ToolTip
    SendMessage lngToolTipHwnd, TTM_ADDTOOLA, 0, lpToolInfo
    
    'ToolTip 颜色设置（vbBlue和vbWhite）
    SendMessage lngToolTipHwnd, TTM_SETTIPTEXTCOLOR, m_ForeColor, ByVal 0&          ' 前景（字体）颜色
    SendMessage lngToolTipHwnd, TTM_SETTIPBKCOLOR, m_BackColor, ByVal 0&            ' 背景颜色
    SendMessage lngToolTipHwnd, TTM_SETTITLE, m_ToolTipIcon, ByVal m_ToolTipTitle   ' 标题
    
    ' 设置显示时间
    SendMessage lngToolTipHwnd, TTM_SETDELAYTIME, TTDT_AUTOPOP, ByVal m_TimeToStay  ' 设置ToolTip显示时间
    SendMessage lngToolTipHwnd, TTM_SETDELAYTIME, TTDT_INITIAL, ByVal m_TimeInterval ' 设置显示ToolTip时隔时间

End Function
' ###################################################################################################
' ### 公共方法 ###
' ###################################################################################################



' ###################################################################################################
' ### 类的私有方法 ###
' ###################################################################################################
' 类初始化时，属性变量的初始化，设置默认属性！
Private Sub Class_Initialize()
    m_lngHwnd = 0                                       ' 欲显示ToolTip的控件句柄
    m_ToolTipIcon = TTI_INFO                            ' 图标
    m_ToolTipTitle = "DyEncryptor"                 ' 标题文字
    m_ToolTipText = "DyEncryptor" ' 提示文本
    
    m_BackColor = vbWhite                               ' 背景色
    m_ForeColor = vbBlue                                ' 前景（字体）颜色
    
    m_TimeToStay = 8888&                                ' ToolTip显示时间，停留的时间！单位：毫秒！
    m_TimeInterval = 222&                               ' 显示ToolTip的时隔时间
    
    m_TTStyle = TT_Balloon                              ' Tooltip 显示样式
End Sub
' 类销毁时，
Private Sub Class_Terminate()
    ' 最后不用的时候别忘了销毁窗口
    DestroyWindow lngToolTipHwnd
End Sub
' ###################################################################################################
' ### 类的私有方法 ###
' ###################################################################################################

