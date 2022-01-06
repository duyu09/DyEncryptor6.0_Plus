Attribute VB_Name = "Module2"
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public starttime As Date
Public OutputDir As String
Public selP As String
Public NumOfHi As Integer, OpenFN As Integer
Public InpWord As String '“系统温馨提示”公有变量
Public Rtb As Integer

Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Type OPENFILENAME
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
Flags As Long
nFileOffset As Integer
nFileExtension As Integer
lpstrDefExt As String
lCustData As Long
lpfnHook As Long
lpTemplateName As String
End Type

Public DU_FoN As String, DU_IconP As String
Public Const Tw = 3855
Sub Main()
''''
Load Form10
DoEvents
Form10.Label1.Width = 0 * Tw
Form10.Show
DoEvents
Sleep (35)
DoEvents
Form10.Label1.Width = 0.03 * Tw
For a11 = 0.03 To 0.09 Step 0.005
    Sleep (5)
    DoEvents
    Form10.Label1.Width = a11 * Tw
    DoEvents
Next
Form10.Label1.Width = 0.09 * Tw
DoEvents
Sleep (94)
DoEvents
Form10.Label1.Width = 0.175 * Tw
DoEvents
For qwera = 0.09 To 0.175 Step 0.009
   DoEvents
   Sleep (7)
   DoEvents
   Form10.Label1.Width = qwera * Tw
   DoEvents
Next
Dim RealCom As String
If Left(Command(), 1) = Chr(34) Then
   RealCom = Mid(Command(), 2, Len(Command()) - 2)
Else
   RealCom = Command()
End If
DoEvents
Form10.Label1.Width = 0.25 * Tw
Load Form1
Load Form2
Load Form9
DoEvents
Sleep (115)
DoEvents
DoEvents
Form10.Label1.Width = 0.335 * Tw
DoEvents
Sleep (120)
DoEvents
If Dir(App.Path & "\DyEncGUI5.0.config") <> "" Then
   On Error Resume Next
   Open App.Path & "\DyEncGUI5.0.config" For Input As #9
        Line Input #9, selP
   Close #9
Else
   selP = App.Path
End If
Form1.Text5.Text = selP
DoEvents
Form10.Label1.Width = 0.45 * Tw
DoEvents
Sleep (105)
Load Form3
For asdfh = 0.45 To 0.525 Step 0.045
    DoEvents
    Form10.Label1.Width = asdfh * Tw
    DoEvents
    Sleep (4)
Next
DoEvents
Form10.Label1.Width = 0.525 * Tw
DoEvents
Sleep (103)
DoEvents
Dim rnm As String, gnm As String, bnm As String
If Dir(App.Path & "\GUI_Color.config") <> "" Then
   On Error Resume Next
   Open App.Path & "\GUI_Color.config" For Input As #11
        Line Input #11, rnm
        Line Input #11, gnm
        Line Input #11, bnm
   Close #11
Form3.HScroll1.Value = Val(rnm)
Form3.HScroll2.Value = Val(gnm)
Form3.HScroll3.Value = Val(bnm)
DoEvents
Sleep (95)
DoEvents
Form10.Label1.Width = 0.585 * Tw
DoEvents
Form1.BackColor = RGB(Val(rnm), Val(gnm), Val(bnm))
Form1.Frame1.BackColor = RGB(Val(rnm), Val(gnm), Val(bnm))
DoEvents
Sleep (50)
DoEvents
Form10.Label1.Width = 0.59 * Tw
DoEvents
Form1.Frame2.BackColor = RGB(Val(rnm), Val(gnm), Val(bnm))
Form1.Option1.BackColor = RGB(Val(rnm), Val(gnm), Val(bnm))
DoEvents
Sleep (45)
DoEvents
Form10.Label1.Width = 0.6 * Tw
DoEvents
Form1.Option2.BackColor = RGB(Val(rnm), Val(gnm), Val(bnm))
Form1.Option3.BackColor = RGB(Val(rnm), Val(gnm), Val(bnm))
DoEvents
Sleep (50)
DoEvents
Form10.Label1.Width = 0.615 * Tw
DoEvents
Form1.Option4.BackColor = RGB(Val(rnm), Val(gnm), Val(bnm))
Form1.Check1.BackColor = RGB(Val(rnm), Val(gnm), Val(bnm))
DoEvents
Sleep (40)
DoEvents
Form10.Label1.Width = 0.625 * Tw
DoEvents
Form1.Check2.BackColor = RGB(Val(rnm), Val(gnm), Val(bnm))
End If
DoEvents
Sleep (55)
DoEvents
Form10.Label1.Width = 0.65 * Tw
DoEvents
NumOfHi = 30
On Error Resume Next
Open App.Path & "\DyEncGUI5.0.OtherSettings.config" For Input As #20
     If Err.Number > 0 Then
        MsgBox "读取" & App.Path & "\DyEncGUI5.0.OtherSettings.config配置文件失败。", 48
        Exit Sub
     End If
     DoEvents
     Sleep (80)
     DoEvents
     Form10.Label1.Width = 0.685 * Tw
     DoEvents
     Dim NumOfHi_s As String, OpenFN_s As String
     Line Input #20, DU_FoN
     Line Input #20, DU_IconP
     Line Input #20, NumOfHi_s
     Line Input #20, OpenFN_s
Close #20
DoEvents
Sleep (105)
DoEvents
Form10.Label1.Width = 0.725 * Tw
NumOfHi = Val(NumOfHi_s)
OpenFN = Val(OpenFN_s)
SetFoName (DU_FoN)
On Error Resume Next
Form1.Icon = LoadPicture(DU_IconP)
On Error Resume Next
Form2.Icon = LoadPicture(DU_IconP)
Sleep (70)
DoEvents
DoEvents
Form10.Label1.Width = 0.775 * Tw
DoEvents
On Error Resume Next
Form3.Icon = LoadPicture(DU_IconP)
On Error Resume Next
Form4.Icon = LoadPicture(DU_IconP)
DoEvents
Sleep (105)
DoEvents
Form10.Label1.Width = 0.825 * Tw
Dim deab As Integer, jsb2 As Boolean
For deab = 0 To Screen.FontCount - 1
    If Screen.Fonts(deab) = "未曾忘记那位少年" Then
       jsb2 = True
    End If
Next deab
If jsb2 = False Then
   On Error Resume Next
   AddFontResource App.Path & "\DyEncGUI_FontsLib\DyEncGUI_FirstFonts.ttf"
   SetFoName ("未曾忘记那位少年")
   DoEvents
   Sleep (75)
   DoEvents
   Form10.Label1.Width = 0.865 * Tw
   DoEvents
End If
DoEvents
Sleep (95)
DoEvents
Form10.Label1.Width = 0.9 * Tw
DoEvents
Form1.Show
DoEvents
On Error Resume Next
If StrConv(Right(RealCom, 13), vbLowerCase) = ".dyenc_output" Then
   Form1.Text3.Text = RealCom
   Form1.Option2.Value = True
   Form1.Text1.SetFocus
Else
   Form1.Text3.Text = RealCom
   Form1.Text1.SetFocus
End If
If Command() = "" Then
   Form1.Command1.SetFocus
End If
DoEvents
Sleep (75)
DoEvents
Form10.Label1.Width = 0.95 * Tw
DoEvents
DoEvents
Sleep (20)
DoEvents
Unload Form10
End Sub

Sub SetFoName(FoN As String)
Dim fnc2 As String
fnc2 = FoN
Form1.fontname = fnc2
Form1.Check2.fontname = fnc2
Form1.Check1.fontname = fnc2
Form1.Label1.fontname = fnc2
Form1.Label2.fontname = fnc2
Form1.Label3.fontname = fnc2
Form1.Label4.fontname = fnc2
Form1.Label5.fontname = fnc2
Form1.Label6.fontname = fnc2
Form1.Label7.fontname = fnc2
Form1.Label8.fontname = fnc2
Form1.Label9.fontname = fnc2
Form1.Label10.fontname = fnc2
Form1.Label11.fontname = fnc2
Form1.Option1.fontname = fnc2
Form1.Option2.fontname = fnc2
Form1.Option3.fontname = fnc2
Form1.Option4.fontname = fnc2
Form1.Command1.fontname = fnc2
Form1.Command2.fontname = fnc2
Form1.Command3.fontname = fnc2
Form1.Command4.fontname = fnc2
Form1.Command5.fontname = fnc2
Form1.Frame1.fontname = fnc2
Form1.Frame2.fontname = fnc2
Form1.Text1.Font = fnc2
Form1.Text2.Font = fnc2
Form1.Text3.fontname = fnc2
Form1.Text4.fontname = fnc2
Form1.Text5.fontname = fnc2
Form2.fontname = fnc2
Form2.Label1.fontname = fnc2
Form2.Text1.fontname = fnc2
Form2.Command1.fontname = fnc2
Form3.fontname = fnc2
Form3.Command1.fontname = fnc2
Form3.Command2.fontname = fnc2
Form3.Frame1.fontname = fnc2
Form3.Label1.fontname = fnc2
Form3.Label2.fontname = fnc2
Form3.Label3.fontname = fnc2
Form3.Label4.fontname = fnc2
Form3.Text1.fontname = fnc2
Form3.Text2.fontname = fnc2
Form3.Text3.fontname = fnc2
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
Form5.Label1.fontname = fnc2
Form5.Label2.fontname = fnc2
Form5.Label3.fontname = fnc2
Form5.Text1.fontname = fnc2
Form6.List1.fontname = fnc2
Form7.Text1.fontname = fnc2
End Sub

'*******************************************************

Public Sub MsgB(Word1 As String)
InpWord = Word1
Load Form7
Form7.Text1.Text = InpWord
If Len(Form7.Text1.Text) > 20 Then
   Form7.Text1.FontSize = 16
Else
   Form7.Text1.FontSize = 22
End If
Form7.Show (1)
InpWord = ""
Do While Rtb = 0
   DoEvents
Loop
Rtb = 0
End Sub
