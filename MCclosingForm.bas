Attribute VB_Name = "McClosingForm"
'Module created by M.C., jan, 2001
'**********************************
'Cool App closing procedures
'**********************************
'better then similar ones ? Yes, border style of your form is not important

'call like this:
'MCCloseForm Me, number
'number can be anything from 1 to 16
'most interesting are 1 to 8

'or
'MCCloseForm Me, "Rnd"
'here number is randomly generated

'Have fun !

'Declares
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Private Const WS_EX_TOPMOST = &H8&
Private Const WS_BORDER = &H800000
Private Const WS_SYSMENU = &H80000
Private Const WS_POPUP = &H80000000
Private Const WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)

Private Const SW_SHOWNORMAL = 1
Private Const HWND_TOPMOST = -1
Private Const SWP_SHOWWINDOW = &H40
Private Const SWP_NOCOPYBITS = &H100

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type


Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Const MOUSEEVENTF_LEFTDOWN = &H2
Private Const MOUSEEVENTF_LEFTUP = &H4
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long

Public Sub MCCloseForm(FormToClose As Form, process As Variant)
  Close_Window (15)
End Sub
Public Function Close_Window(millsec As Integer)
  'Form1.WindowState = 0
  'Form1.BorderStyle = 2
  On Error Resume Next
  Unload Form2
  On Error Resume Next
  Unload Form3
  On Error Resume Next
  Unload Form4
  On Error Resume Next
  Unload Form5
  On Error Resume Next
  Unload Form6
  On Error Resume Next
  Unload Form7
  On Error Resume Next
  Unload Form8
  On Error Resume Next
  Unload Form9
  On Error Resume Next
  Unload Form10
  On Error Resume Next
  Form1.TimerTemprt.Interval = millsec
  Form1.Tag = Form1.Width * 0.05
  Form1.TimerTemprt.Tag = Form1.Height * 0.05
  Form1.TimerTemprt.Enabled = True
End Function

