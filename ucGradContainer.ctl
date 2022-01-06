VERSION 5.00
Begin VB.UserControl ucGradContainer 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0C0&
   ClientHeight    =   2700
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3525
   ControlContainer=   -1  'True
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   180
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   235
   ToolboxBitmap   =   "ucGradContainer.ctx":0000
End
Attribute VB_Name = "ucGradContainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'Download by http://www.codefans.net
Option Explicit

Private Declare Sub RtlMoveMemory Lib "kernel32" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function MoveTo Lib "gdi32" Alias "MoveToEx" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As Any) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, pccolorref As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function DrawText Lib "user32.dll" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function FillRgn Lib "gdi32.dll" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function CombineRgn Lib "gdi32.dll" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long

Private Type RECT
    Left    As Long
    Top     As Long
    Right   As Long
    Bottom  As Long
End Type

Public Enum HeaderAngleValues
   [Horizontal] = 180
   [Vertical] = 90
End Enum

Public Enum IconSizeEnum
   [Display Full Size] = 0
   [Size To Header] = 1
End Enum

' [Property Variables / Constants]
Private m_HeaderVisible As Boolean
Private m_BackMiddleOut As Boolean
Private m_HeaderMiddleOut As Boolean
Private m_Enabled As Boolean
Private m_HeaderAngle       As HeaderAngleValues
Private m_BackAngle         As Integer ' background gradient display angle
Private m_Iconsize          As IconSizeEnum
Private m_HeaderColor2      As OLE_COLOR
Private m_HeaderColor1      As OLE_COLOR
Private m_BackColor2        As OLE_COLOR
Private m_BackColor1        As OLE_COLOR
Private m_BorderThickness   As Integer
Private m_BorderColor       As OLE_COLOR
Private m_CaptionColor      As OLE_COLOR
Private m_Caption           As String
Private m_HeaderHeight      As Long
Private m_CaptionFont       As StdFont
Private m_Alignment         As AlignmentConstants
Private m_Icon              As Picture
Private m_hMod              As Long
Private m_Curvature         As Long

Private Const m_def_HeaderVisible = True
Private Const m_def_BackMiddleOut = True
Private Const m_def_HeaderMiddleOut = True
Private Const m_def_Enabled = 0
Private Const m_def_HeaderAngle = 180     ' init to horizontal header gradient
Private Const m_def_BackAngle = 180       ' init to horizontal bg gradient
Private Const m_def_Iconsize = 1 ' size to header
Private Const m_DEF_HeaderColor2 = &HF7E0D3
Private Const m_DEF_HeaderColor1 = &HEDC5A7
Private Const m_DEF_BackColor2 = &HFCF4EF
Private Const m_DEF_BackColor1 = &HFAE8DC
Private Const M_DEF_Caption = "Gradient Container"
Private Const m_def_BorderThickness = 1
Private Const m_DEF_BorderColor = &HDCC1AD
Private Const m_DEF_Align = vbLeftJustify
Private Const m_DEF_CaptionColor = &H7B2D02
Private Const m_DEF_Curvature = 10
Private Const m_DEF_hHeight = 25

Event Resize()
'Event Declarations:
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Private bInitialized As Boolean
Private bDrawFirstTime As Boolean             ' for when you draw a new button in design mode.

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub RedrawControl()

   If bInitialized Or bDrawFirstTime Then
      UserControl.Cls
      SetBackGround
      If m_HeaderVisible Then
         SetHeader
      End If
      DrawBorder
   End If

End Sub

Private Function TranslateColor(ByVal oClr As OLE_COLOR, Optional hPal As Long = 0) As Long
   If OleTranslateColor(oClr, hPal, TranslateColor) Then
      TranslateColor = -1
   End If
End Function

Private Sub UserControl_Initialize()
   m_hMod = LoadLibrary("shell32.dll") ' Used to prevent crashes on Windows XP
End Sub

Private Sub UserControl_InitProperties()
   Set m_Icon = Nothing
   Set m_CaptionFont = Ambient.Font
   m_HeaderAngle = m_def_HeaderAngle
   m_BackAngle = m_def_BackAngle
   m_HeaderColor2 = m_DEF_HeaderColor2
   m_HeaderColor1 = m_DEF_HeaderColor1
   m_BackColor2 = m_DEF_BackColor2
   m_BackColor1 = m_DEF_BackColor1
   m_BorderColor = m_DEF_BorderColor
   m_CaptionColor = m_DEF_CaptionColor
   m_Caption = M_DEF_Caption
   m_Alignment = vbLeftJustify
   m_Curvature = m_DEF_Curvature
   m_HeaderHeight = m_DEF_hHeight
   m_Enabled = m_def_Enabled
   m_BorderThickness = m_def_BorderThickness
   m_BackMiddleOut = m_def_BackMiddleOut
   m_HeaderMiddleOut = m_def_HeaderMiddleOut
   m_HeaderVisible = m_def_HeaderVisible
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   Set m_Icon = PropBag.ReadProperty("HeaderIcon", Nothing)
   Set m_CaptionFont = PropBag.ReadProperty("CaptionFont", Ambient.Font)
   m_Iconsize = PropBag.ReadProperty("IconSize", m_def_Iconsize)
   m_HeaderAngle = PropBag.ReadProperty("HeaderAngle", m_def_HeaderAngle)
   m_BackAngle = PropBag.ReadProperty("BackAngle", m_def_BackAngle)
   m_HeaderColor2 = PropBag.ReadProperty("HeaderColor2", m_DEF_HeaderColor2)
   m_HeaderColor1 = PropBag.ReadProperty("HeaderColor1", m_DEF_HeaderColor1)
   m_BackColor2 = PropBag.ReadProperty("BackColor2", m_DEF_BackColor2)
   m_BackColor1 = PropBag.ReadProperty("BackColor1", m_DEF_BackColor1)
   m_BorderColor = PropBag.ReadProperty("BorderColor", m_DEF_BorderColor)
   m_CaptionColor = PropBag.ReadProperty("CaptionColor", m_DEF_CaptionColor)
   m_Caption = PropBag.ReadProperty("Caption", M_DEF_Caption)
   m_Curvature = PropBag.ReadProperty("Curvature", m_DEF_Curvature)
   m_Alignment = PropBag.ReadProperty("CaptionAlignment", m_DEF_Align)
   m_HeaderHeight = PropBag.ReadProperty("HeaderHeight", m_DEF_hHeight)
   m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
   m_BorderThickness = PropBag.ReadProperty("BorderThickness", m_def_BorderThickness)
   m_BackMiddleOut = PropBag.ReadProperty("BackMiddleOut", m_def_BackMiddleOut)
   m_HeaderMiddleOut = PropBag.ReadProperty("HeaderMiddleOut", m_def_HeaderMiddleOut)
   m_HeaderVisible = PropBag.ReadProperty("HeaderVisible", m_def_HeaderVisible)

   bInitialized = True

End Sub

Private Sub UserControl_Show()
   bDrawFirstTime = True
   RedrawControl '   UserControl_Resize
   bDrawFirstTime = False
End Sub

Private Sub UserControl_Terminate()
   FreeLibrary m_hMod ' Used to prevent crashes on Windows XP
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
   Call PropBag.WriteProperty("HeaderAngle", m_HeaderAngle, m_def_HeaderAngle)
   Call PropBag.WriteProperty("BackAngle", m_BackAngle, m_def_BackAngle)
   Call PropBag.WriteProperty("IconSize", m_Iconsize, m_def_Iconsize)
   Call PropBag.WriteProperty("HeaderColor2", m_HeaderColor2, m_DEF_HeaderColor2)
   Call PropBag.WriteProperty("HeaderColor1", m_HeaderColor1, m_DEF_HeaderColor1)
   Call PropBag.WriteProperty("BackColor2", m_BackColor2, m_DEF_BackColor2)
   Call PropBag.WriteProperty("BackColor1", m_BackColor1, m_DEF_BackColor1)
   Call PropBag.WriteProperty("BorderColor", m_BorderColor, m_DEF_BorderColor)
   Call PropBag.WriteProperty("CaptionColor", m_CaptionColor, m_DEF_CaptionColor)
   Call PropBag.WriteProperty("Caption", m_Caption, M_DEF_Caption)
   Call PropBag.WriteProperty("CaptionAlignment", m_Alignment, vbLeftJustify)
   Call PropBag.WriteProperty("HeaderHeight", m_HeaderHeight, m_DEF_hHeight)
   Call PropBag.WriteProperty("CaptionFont", m_CaptionFont, Ambient.Font)
   Call PropBag.WriteProperty("Curvature", m_Curvature, m_DEF_Curvature)
   Call PropBag.WriteProperty("HeaderIcon", m_Icon, Nothing)
   Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
   Call PropBag.WriteProperty("BorderThickness", m_BorderThickness, m_def_BorderThickness)
   Call PropBag.WriteProperty("BackMiddleOut", m_BackMiddleOut, m_def_BackMiddleOut)
   Call PropBag.WriteProperty("HeaderMiddleOut", m_HeaderMiddleOut, m_def_HeaderMiddleOut)
   Call PropBag.WriteProperty("HeaderVisible", m_HeaderVisible, m_def_HeaderVisible)
End Sub

Public Property Get HeaderVisible() As Boolean
   HeaderVisible = m_HeaderVisible
End Property

Public Property Let HeaderVisible(ByVal New_HeaderVisible As Boolean)
   m_HeaderVisible = New_HeaderVisible
   PropertyChanged "HeaderVisible"
   RedrawControl
End Property

Public Property Get BackMiddleOut() As Boolean
   BackMiddleOut = m_BackMiddleOut
End Property

Public Property Let BackMiddleOut(ByVal New_BackMiddleOut As Boolean)
   m_BackMiddleOut = New_BackMiddleOut
   PropertyChanged "BackMiddleOut"
   RedrawControl
End Property

Public Property Get HeaderMiddleOut() As Boolean
   HeaderMiddleOut = m_HeaderMiddleOut
End Property

Public Property Let HeaderMiddleOut(ByVal New_HeaderMiddleOut As Boolean)
   m_HeaderMiddleOut = New_HeaderMiddleOut
   PropertyChanged "HeaderMiddleOut"
   RedrawControl
End Property

Public Property Get HeaderAngle() As HeaderAngleValues
   HeaderAngle = m_HeaderAngle
End Property

Public Property Let HeaderAngle(ByVal New_HeaderAngle As HeaderAngleValues)
   m_HeaderAngle = New_HeaderAngle
   PropertyChanged "HeaderAngle"
   RedrawControl
End Property

Public Property Get BackAngle() As Integer
   BackAngle = m_BackAngle
End Property

Public Property Let BackAngle(ByVal New_BackAngle As Integer)
'  do some bounds checking.
   If New_BackAngle > 360 Then
      New_BackAngle = 360
   ElseIf New_BackAngle < 0 Then
      New_BackAngle = 0
   End If
   m_BackAngle = New_BackAngle
   PropertyChanged "BackAngle"
   RedrawControl
End Property

Public Property Get BackColor1() As OLE_COLOR
   BackColor1 = m_BackColor1
End Property

Public Property Let BackColor1(ByVal New_BackColor1 As OLE_COLOR)
   m_BackColor1 = New_BackColor1
   PropertyChanged "BackColor1"
   RedrawControl
End Property

Public Property Get BackColor2() As OLE_COLOR
   BackColor2 = m_BackColor2
End Property

Public Property Let BackColor2(ByVal New_BackColor2 As OLE_COLOR)
   m_BackColor2 = New_BackColor2
   PropertyChanged "BackColor2"
   RedrawControl
End Property

Public Property Get BorderThickness() As Integer
   BorderThickness = m_BorderThickness
End Property

Public Property Let BorderThickness(ByVal New_BorderThickness As Integer)
   m_BorderThickness = New_BorderThickness
   PropertyChanged "BorderThickness"
   RedrawControl
End Property

Public Property Get BorderColor() As OLE_COLOR
   BorderColor = m_BorderColor
End Property

Public Property Let BorderColor(ByVal New_BorderColor As OLE_COLOR)
   m_BorderColor = New_BorderColor
   PropertyChanged "BorderColor"
   RedrawControl
End Property

Public Property Get Caption() As String
   Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
   m_Caption = New_Caption
   PropertyChanged "Caption"
   RedrawControl
End Property

Public Property Get HeaderColor1() As OLE_COLOR
   HeaderColor1 = m_HeaderColor1
End Property

Public Property Let HeaderColor1(ByVal New_HeaderColor1 As OLE_COLOR)
   m_HeaderColor1 = New_HeaderColor1
   PropertyChanged "HeaderColor1"
   RedrawControl
End Property

Public Property Get HeaderColor2() As OLE_COLOR
   HeaderColor2 = m_HeaderColor2
End Property

Public Property Let HeaderColor2(ByVal New_HeaderColor2 As OLE_COLOR)
   m_HeaderColor2 = New_HeaderColor2
   PropertyChanged "HeaderColor2"
   RedrawControl
End Property

Public Property Get IconSize() As IconSizeEnum
   IconSize = m_Iconsize
End Property

Public Property Let IconSize(ByVal New_IconSize As IconSizeEnum)
   m_Iconsize = New_IconSize
   PropertyChanged "IconSize"
   RedrawControl
End Property

Public Property Get CaptionColor() As OLE_COLOR
   CaptionColor = m_CaptionColor
End Property

Public Property Let CaptionColor(ByVal New_CaptionColor As OLE_COLOR)
   m_CaptionColor = New_CaptionColor
   PropertyChanged "CaptionColor"
   RedrawControl
End Property

Public Property Get HeaderHeight() As Long
   HeaderHeight = m_HeaderHeight
End Property

Public Property Let HeaderHeight(ByVal vNewHeight As Long)
   m_HeaderHeight = vNewHeight
   PropertyChanged "HeaderHeight"
   RedrawControl
End Property

Public Property Get CaptionFont() As Font
   Set CaptionFont = m_CaptionFont
End Property

Public Property Set CaptionFont(ByVal vNewCaptionFont As Font)
   Set m_CaptionFont = vNewCaptionFont
   PropertyChanged "CaptionFont"
   RedrawControl
End Property

Public Property Get CaptionAlignment() As AlignmentConstants
   CaptionAlignment = m_Alignment
End Property

Public Property Let CaptionAlignment(ByVal vNewAlignment As AlignmentConstants)
   m_Alignment = vNewAlignment
   PropertyChanged "CaptionAlignment"
   RedrawControl
End Property

Public Property Get HeaderIcon() As Picture
   Set HeaderIcon = m_Icon
End Property

Public Property Set HeaderIcon(ByVal vNewIcon As Picture)
   Set m_Icon = vNewIcon
   PropertyChanged "HeaderIcon"
   RedrawControl
End Property

Public Property Get Curvature() As Long
   Curvature = m_Curvature
End Property

Public Property Let Curvature(ByVal vNewCurvature As Long)
   m_Curvature = vNewCurvature
   PropertyChanged "Curvature"
   RedrawControl
End Property

Public Property Get Enabled() As Boolean
   Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
   m_Enabled = New_Enabled
   PropertyChanged "Enabled"
End Property

Private Sub UserControl_Resize()
   bDrawFirstTime = True
   RedrawControl
   bDrawFirstTime = False
End Sub

Private Sub SetBackGround()

'*************************************************************************
'* displays the control's background gradient.                           *
'*************************************************************************

   On Error GoTo ErrHandler

   Dim lColor As Long, lColor2 As Long

'  Get the control colors
   lColor = TranslateColor(m_BackColor2)
   lColor2 = TranslateColor(m_BackColor1)
'  Apply the gradients
   DrawGradient hdc, UserControl.ScaleWidth, UserControl.ScaleHeight, lColor, lColor2, m_BackAngle, m_BackMiddleOut

ErrHandler:
   Exit Sub

End Sub

Private Sub DrawBorder()

'*************************************************************************
'* draws the border around the control, using appropriate curvature      *
'*************************************************************************

   On Error GoTo ErrHandler

   Dim BdrCol As Long
   Dim hBrush As Long
   Dim hrgn1 As Long
   Dim hrgn2 As Long

'  Get the border color
   BdrCol = TranslateColor(m_BorderColor)

'  Define the regions
   hrgn1 = CreateRoundRectRgn(0, 0, ScaleWidth, ScaleHeight, m_Curvature, m_Curvature)
   hrgn2 = CreateRoundRectRgn(m_BorderThickness, m_BorderThickness, ScaleWidth - m_BorderThickness, ScaleHeight - m_BorderThickness, m_Curvature, m_Curvature)
   CombineRgn hrgn2, hrgn1, hrgn2, 3

'  Create/Apply the ColorBrush
   hBrush = CreateSolidBrush(BdrCol)
   FillRgn hdc, hrgn2, hBrush

'  Set the control region
   SetWindowRgn hWnd, hrgn1, True

'  Free the memory
   DeleteObject hrgn2
   DeleteObject hBrush
   DeleteObject hrgn1

ErrHandler:
   Exit Sub

End Sub

Private Sub SetHeader()

'*************************************************************************
'* displays the header gradient, caption text and an icon if used        *
'*************************************************************************

   Dim lColor  As Long
   Dim lColor2 As Long

   If Not m_CaptionFont Is Nothing Then

'     Get color / Fill gradients
      lColor = TranslateColor(m_HeaderColor2)
      lColor2 = TranslateColor(m_HeaderColor1)
      DrawGradient hdc, UserControl.ScaleWidth, m_HeaderHeight, lColor, lColor2, m_HeaderAngle, m_HeaderMiddleOut

'     draw the caption
      Dim r           As RECT
      Dim tHeight     As Long
      Dim tWidth      As Long
      Dim Clearance   As Long

'     Apply the font/Forecolor
      Set UserControl.Font = m_CaptionFont
      tHeight = TextHeight(m_Caption)
      tWidth = TextWidth(m_Caption)
      UserControl.ForeColor = TranslateColor(m_CaptionColor)

'     make the left clearance one letter width.
      Clearance = TextWidth("A")
      With r 'Define the drawing rectangle size.
         If m_Alignment = vbCenter Then
            .Left = (ScaleWidth - TextWidth(m_Caption)) / 2
         ElseIf m_Alignment = vbLeftJustify Then
            If IsThere(m_Icon) Then
               .Left = TextWidth("A") + m_HeaderHeight
            Else
               .Left = Clearance
            End If
         Else
            .Left = (ScaleWidth - TextWidth(m_Caption)) - Clearance
         End If
         .Top = (m_HeaderHeight - TextHeight(m_Caption)) / 2
         .Bottom = r.Top + tHeight
         .Right = .Left + tWidth
      End With

'     Draw the caption.
      DrawText hdc, m_Caption, -1, r, 0

'     Draw the icon with the most simple method.
      If IsThere(m_Icon) Then
         If m_Iconsize = [Display Full Size] Then ' don't fit to header;, display full size
            PaintPicture m_Icon, m_BorderThickness + 3, 2
         Else
            PaintPicture m_Icon, m_BorderThickness + 3, 2, m_HeaderHeight - 2, m_HeaderHeight - 3 ' fit to header height
         End If
      End If

   End If

End Sub

Private Function IsThere(ByVal Pic As StdPicture) As Boolean

'*************************************************************************
'* checks for existence of a picture by checking dimensions.             *
'*************************************************************************

   If Not Pic Is Nothing Then
      If Pic.Height <> 0 Then
         IsThere = Pic.Width <> 0
      End If
   End If

End Function

Private Sub DrawGradient(ByVal hdc As Long, ByVal lWidth As Long, ByVal lHeight As Long, _
                         ByVal lCol1 As Long, ByVal lCol2 As Long, ByVal zAngle As Single, ByVal bMOut As Boolean)

'*************************************************************************
'* adapted version of redbird77's gradient generation routine that       *
'* supports middle-out gradients as well as regular.                     *
'*************************************************************************

   Dim xStart  As Long, yStart As Long
   Dim xEnd    As Long, yEnd   As Long
   Dim X1      As Long, Y1     As Long
   Dim X2      As Long, Y2     As Long
   Dim lRange  As Long
   Dim iQ      As Integer
   Dim bVert   As Boolean
   Dim lPtr    As Long, lInc   As Long
   Dim lCols() As Long, lCols2() As Long
   Dim hPO     As Long, hPN    As Long
   Dim r       As Long
   Dim X       As Long, xUp    As Long
   Dim b1(2)   As Byte, b2(2)  As Byte, b3(2) As Byte
   Dim p       As Single, ip   As Single
   Dim Y As Long

   lInc = 1
   xEnd = lWidth - 1
   yEnd = lHeight - 1

'  Positive angles are measured counter-clockwise; negative angles clockwise.
   zAngle = zAngle Mod 360
   If zAngle < 0 Then zAngle = 360 + zAngle

'  Get angle's quadrant (0 - 3).
   iQ = zAngle \ 90

'  Is angle more horizontal or vertical?
   bVert = ((iQ + 1) * 90) - zAngle > 45
   If (iQ Mod 2 = 0) Then bVert = Not bVert

'  Convert angle in degrees to radians.
   zAngle = zAngle * Atn(1) / 45

'  Get start and end y-positions (if vertical), x-positions (if horizontal).
   If bVert Then
      If zAngle Then xStart = lHeight / Abs(Tan(zAngle))
      lRange = lWidth + xStart - 1

      Y1 = IIf(iQ Mod 2, 0, yEnd)
      Y2 = IIf(Y1, -1, lHeight)
      If iQ > 1 Then
         lPtr = lRange: lInc = -1
      End If
   Else
      yStart = lWidth * Abs(Tan(zAngle))
      lRange = lHeight + yStart - 1

      X1 = IIf(iQ Mod 2, 0, xEnd)
      X2 = IIf(X1, -1, lWidth)

      If iQ = 1 Or iQ = 2 Then
         lPtr = lRange: lInc = -1
      End If
   End If

'  Fill in the color array with the interpolated color values.
   ReDim lCols(lRange)
   ReDim lCols2(lRange)

   ' Get the r, g, b components of each color.
   RtlMoveMemory b1(0), lCol1, 3
   RtlMoveMemory b2(0), lCol2, 3
   RtlMoveMemory b3(0), 0, 3
   xUp = UBound(lCols)

   If bMOut Then ' middle-out gradient desired.
'     get the full color array in lCols2.
      For X = 0 To xUp
         ' Get the position and the 1 - position.
         p = X / xUp
         ip = 1 - p
         ' Interpolate the value at the current position.
         lCols2(X) = RGB(b1(0) * ip + b2(0) * p, b1(1) * ip + b2(1) * p, b1(2) * ip + b2(2) * p)
      Next X
'     put the array in first half of lcols
      Y = 0
      For X = 0 To xUp Step 2
         lCols(Y) = lCols2(X)
         Y = Y + 1
      Next X
'     put the reverse of the array in the second half of lcols
      For X = xUp - 1 To 1 Step -2
         lCols(Y) = lCols2(X)
         Y = Y + 1
      Next X
   Else
'     get the full color array in lCols2.
      For X = 0 To xUp
         ' Get the position and the 1 - position.
         p = X / xUp
         ip = 1 - p
         ' Interpolate the value at the current position.
         lCols(X) = RGB(b1(0) * ip + b2(0) * p, b1(1) * ip + b2(1) * p, b1(2) * ip + b2(2) * p)
      Next X
   End If

   If bVert Then
      For X1 = -xStart To xEnd
         hPN = CreatePen(0, 1, lCols(lPtr))
         hPO = SelectObject(hdc, hPN)
         MoveTo hdc, X1, Y1, ByVal 0&
         LineTo hdc, X2, Y2
         r = SelectObject(hdc, hPO): r = DeleteObject(hPN)
         lPtr = lPtr + lInc
         X2 = X2 + 1
      Next
   Else
      For Y1 = -yStart To yEnd
         hPN = CreatePen(0, 1, lCols(lPtr))
         hPO = SelectObject(hdc, hPN)
         MoveTo hdc, X1, Y1, ByVal 0&
         LineTo hdc, X2, Y2
         r = SelectObject(hdc, hPO): r = DeleteObject(hPN)
         lPtr = lPtr + lInc
         Y2 = Y2 + 1
      Next
   End If

End Sub
