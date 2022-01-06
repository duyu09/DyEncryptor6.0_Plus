Attribute VB_Name = "Module3"
Declare Function AddFontResource Lib "gdi32" Alias "AddFontResourceA" (ByVal lpFileName As String) As Long
Declare Function RemoveFontResource Lib "gdi32" Alias "RemoveFontResourceA" (ByVal lpFileName As String) As Long

Option Explicit

'******************************************************************
'����.ttf�����ļ���ȡ���������ơ�
'ת��ע����Դ Http://Www.YuLv.Net/
'******************************************************************

'Api ����
Declare Sub RtlMoveMemory Lib "kernel32" (dst As Any, src As Any, ByVal Length As Long)
Declare Function ntohl Lib "ws2_32.dll" (ByVal netlong As Long) As Long
Declare Function ntohs Lib "ws2_32.dll" (ByVal netshort As Integer) As Integer

'��������
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
'ת���ֽ�˳�����
'***********************************************************
Sub SwapLong(LongVal As Long)
LongVal = ntohl(LongVal)
End Sub

Sub SwapInt(IntVal As Integer)
IntVal = ntohs(IntVal)
End Sub


'************************************************************
'��Ҫ�������£�
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

'�Զ����Ƶķ�ʽ��TTF�ļ�
On Error GoTo Finished
FileNum = FreeFile
Open FontPath For Binary As FileNum

'��ȡ��һ����ͷ
Get #FileNum, , OffSetTbl

'���汾�Ƿ�Ϊ1.0
With OffSetTbl
SwapInt .uMajorVersion
SwapInt .uMinorVersion
SwapInt .uNumOfTables
If .uMajorVersion <> 1 Or .uMinorVersion <> 0 Then
Debug.Print FontPath & " -> ����汾����ȷ, �޷�ȡ����������!"
GoTo Finished
End If
End With

If OffSetTbl.uNumOfTables > 0 Then
For X = 0 To OffSetTbl.uNumOfTables - 1
Get #FileNum, , TblDir
If StrComp(TblDir.szTag, "name", vbTextCompare) = 0 Then
'����ҵ������������ƫ�����������
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

'�ַ���Ϊ�գ�����������
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

