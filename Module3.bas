Attribute VB_Name = "Module3"
Declare Function AddFontResource Lib "gdi32" Alias "AddFontResourceA" (ByVal lpFileName As String) As Long
Declare Function RemoveFontResource Lib "gdi32" Alias "RemoveFontResourceA" (ByVal lpFileName As String) As Long

Option Explicit

'******************************************************************
'根据.ttf字体文件，取得字体名称。
'转载注明来源 Http://Www.YuLv.Net/
'******************************************************************

'Api 声明
Declare Sub RtlMoveMemory Lib "kernel32" (dst As Any, src As Any, ByVal Length As Long)
Declare Function ntohl Lib "ws2_32.dll" (ByVal netlong As Long) As Long
Declare Function ntohs Lib "ws2_32.dll" (ByVal netshort As Integer) As Integer

'常量声明
Public Type OFFSET_TABLE
uMajorVersion As Integer
uMinorVersion As Integer
uNumOfTables As Integer
uSearchRange As Integer
uEntrySelector As Integer
uRangeShift As Integer
End Type

Public Type TABLE_DIRECTORY
szTag As String * 4
uCheckSum As Long
uOffset As Long
uLength As Long
End Type

Public Type NAME_TABLE_HEADER
uFSelector As Integer
uNRCount As Integer
uStorageOffset As Integer
End Type

Public Type NAME_RECORD
uPlatformID As Integer
uEncodingID As Integer
uLanguageID As Integer
uNameID As Integer
uStringLength As Integer
uStringOffset As Integer
End Type


'************************************************************
'转换字节顺序相关
'***********************************************************
Sub SwapLong(LongVal As Long)
LongVal = ntohl(LongVal)
End Sub

Sub SwapInt(IntVal As Integer)
IntVal = ntohs(IntVal)
End Sub


'************************************************************
'主要过程如下：
'***********************************************************
Function GetFontName(ByVal FontPath As String) As String

Dim TblDir As TABLE_DIRECTORY
Dim OffSetTbl As OFFSET_TABLE
Dim NameTblHdr As NAME_TABLE_HEADER
Dim NameRecord As NAME_RECORD
Dim FileNum As Integer
Dim lPosition As Long
Dim sFontTest As String
Dim X As Long
Dim I As Long

'以二进制的方式打开TTF文件
On Error GoTo Finished
FileNum = FreeFile
Open FontPath For Binary As FileNum

'读取第一个表头
Get #FileNum, , OffSetTbl

'检查版本是否为1.0
With OffSetTbl
SwapInt .uMajorVersion
SwapInt .uMinorVersion
SwapInt .uNumOfTables
If .uMajorVersion <> 1 Or .uMinorVersion <> 0 Then
Debug.Print FontPath & " -> 字体版本不正确, 无法取得字体名称!"
GoTo Finished
End If
End With

If OffSetTbl.uNumOfTables > 0 Then
For X = 0 To OffSetTbl.uNumOfTables - 1
Get #FileNum, , TblDir
If StrComp(TblDir.szTag, "name", vbTextCompare) = 0 Then
'如果找到了字体的名称偏移量则继续：
With TblDir
SwapLong .uLength
SwapLong .uOffset
If .uOffset Then
Get #FileNum, .uOffset + 1, NameTblHdr
SwapInt NameTblHdr.uNRCount
SwapInt NameTblHdr.uStorageOffset

For I = 0 To NameTblHdr.uNRCount - 1
Get #FileNum, , NameRecord
SwapInt NameRecord.uNameID

If NameRecord.uNameID = 1 Then
SwapInt NameRecord.uStringLength
SwapInt NameRecord.uStringOffset
lPosition = Loc(FileNum)

If NameRecord.uStringLength Then
sFontTest = Space$(NameRecord.uStringLength)
Get #FileNum, TblDir.uOffset + NameRecord.uStringOffset + NameTblHdr.uStorageOffset + 1, sFontTest
If Len(sFontTest) Then
GoTo Finished
End If
End If

'字符串为空，继续搜索。
Seek #FileNum, lPosition

End If
Next I
End If
End With
End If
Next X
End If


Finished:
Close #FileNum

Dim getfontnamet As String, dt As Integer, few As String
getfontnamet = sFontTest

For dt = 1 To Len(getfontnamet) / 2
    few = few + Mid(getfontnamet, dt * 2, 1)
Next dt
GetFontName = few
End Function

