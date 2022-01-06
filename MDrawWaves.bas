Attribute VB_Name = "MDrawWaves"
Option Explicit
'Download by http://www.codefans.net
Private Const CALLBACK_FUNCTION = &H30000
Private Const MMIO_READ = &H0
Private Const MMIO_FINDCHUNK = &H10
Private Const MMIO_FINDRIFF = &H20
Private Const MM_WOM_DONE = &H3BD

Private Type mmioinfo
   dwFlags As Long
   fccIOProc As Long
   pIOProc As Long
   wErrorRet As Long
   htask As Long
   cchBuffer As Long
   pchBuffer As String
   pchNext As String
   pchEndRead As String
   pchEndWrite As String
   lBufOffset As Long
   lDiskOffset As Long
   adwInfo(4) As Long
   dwReserved1 As Long
   dwReserved2 As Long
   hmmio As Long
End Type
   
Private Type WaveFormat
   wFormatTag As Integer
   nChannels As Integer
   nSamplesPerSec As Long
   nAvgBytesPerSec As Long
   nBlockAlign As Integer
   wBitsPerSample As Integer
   cbSize As Integer
End Type

Private Type MMCKINFO
    ckid As Long
    ckSize As Long
    fccType As Long
    dwDataOffset As Long
    dwFlags As Long
End Type

Private Declare Function mmioClose Lib "winmm.dll" (ByVal hmmio As Long, ByVal uFlags As Long) As Long
Private Declare Function mmioDescend Lib "winmm.dll" (ByVal hmmio As Long, lpck As MMCKINFO, lpckParent As MMCKINFO, ByVal uFlags As Long) As Long
Private Declare Function mmioDescendParent Lib "winmm.dll" Alias "mmioDescend" (ByVal hmmio As Long, lpck As MMCKINFO, ByVal X As Long, ByVal uFlags As Long) As Long
Private Declare Function mmioOpen Lib "winmm.dll" Alias "mmioOpenA" (ByVal szFileName As String, lpmmioinfo As mmioinfo, ByVal dwOpenFlags As Long) As Long
Private Declare Function mmioRead Lib "winmm.dll" (ByVal hmmio As Long, ByVal pch As Long, ByVal cch As Long) As Long
Private Declare Function mmioReadFormat Lib "winmm.dll" Alias "mmioRead" (ByVal hmmio As Long, ByRef pch As WaveFormat, ByVal cch As Long) As Long
Private Declare Function mmioStringToFOURCC Lib "winmm.dll" Alias "mmioStringToFOURCCA" (ByVal sz As String, ByVal uFlags As Long) As Long
Private Declare Function mmioAscend Lib "winmm.dll" (ByVal hmmio As Long, lpck As MMCKINFO, ByVal uFlags As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hmem As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hmem As Long) As Long ' �ͷ�ָ����ȫ���ڴ�顣
Private Declare Sub CopyStructFromPtr Lib "kernel32" Alias "RtlMoveMemory" (struct As Any, ByVal ptr As Long, ByVal cb As Long)

' variables for managing wave file
Private wFormat As WaveFormat
Private hmmioOut As Long
Private mmckinfoParentIn As MMCKINFO
Private mmckinfoSubchunkIn As MMCKINFO

Private bufferIn As Long
Private hmem As Long
Private numSamples As Long
Private drawFrom As Long
Private drawTo As Long

' ���� Wave �ļ����ڴ档
Private Function LoadFile(inFile As String) As Boolean
    LoadFile = False
    ' �жϴ�������Ƿ�Ϊ�գ�
    If (Len(inFile) = 0) Then: GlobalFree hmem: Exit Function
    
    Dim rc As Long
    Dim hmmioIn As Long
    Dim mmioinf As mmioinfo

    ' �� Wave �ļ�
    hmmioIn = mmioOpen(inFile, mmioinf, MMIO_READ)
    If hmmioIn = 0 Then
        MsgBox "���ļ�����rc = " & mmioinf.wErrorRet, vbCritical
        Exit Function
    End If
   
    ' ����ļ��Ƿ�ΪWave��ʽ
    mmckinfoParentIn.fccType = mmioStringToFOURCC("WAVE", 0)
    rc = mmioDescendParent(hmmioIn, mmckinfoParentIn, 0, MMIO_FINDRIFF)
    If (rc <> 0) Then
        rc = mmioClose(hmmioOut, 0)
        MsgBox "���󣺲�����Ч�� Wave ��ʽ�ļ�", vbCritical
        Exit Function
    End If
   
    ' ȡ���ļ��ṹ��Ϣ
    mmckinfoSubchunkIn.ckid = mmioStringToFOURCC("fmt", 0)
    rc = mmioDescend(hmmioIn, mmckinfoSubchunkIn, mmckinfoParentIn, MMIO_FINDCHUNK)
    If (rc <> 0) Then
        rc = mmioClose(hmmioOut, 0)
        MsgBox "���󣺲���ȡ���ļ���ʽ�飡", vbCritical
        Exit Function
    End If
    rc = mmioReadFormat(hmmioIn, wFormat, Len(wFormat))
    If (rc = -1) Then
       rc = mmioClose(hmmioOut, 0)
       MsgBox "��ȡ�ļ���ʽ��Ϣʧ�ܣ�", vbCritical
       Exit Function
    End If
    rc = mmioAscend(hmmioIn, mmckinfoSubchunkIn, 0)
   
    ' ȡ���ļ����ݿ�
    mmckinfoSubchunkIn.ckid = mmioStringToFOURCC("data", 0)
    rc = mmioDescend(hmmioIn, mmckinfoSubchunkIn, mmckinfoParentIn, MMIO_FINDCHUNK)
    If (rc <> 0) Then
       rc = mmioClose(hmmioOut, 0)
       MsgBox "�����޷�ȡ���ļ����ݿ飡", vbCritical
       Exit Function
    End If
   
    ' Allocate soundbuffer and read sound data
    GlobalFree hmem
    hmem = GlobalAlloc(&H40, mmckinfoSubchunkIn.ckSize)
    bufferIn = GlobalLock(hmem)
    rc = mmioRead(hmmioIn, bufferIn, mmckinfoSubchunkIn.ckSize)

    numSamples = mmckinfoSubchunkIn.ckSize / wFormat.nBlockAlign
   
    ' �ر��ļ�
    rc = mmioClose(hmmioOut, 0)
   
    LoadFile = True
    
End Function

Private Sub GetStereo16Sample(ByVal sample As Long, ByRef leftVol As Double, ByRef rightVol As Double)
    ' These subs obtain a PCM sample and converts it into volume levels from (-1 to 1)
   Dim sample16 As Integer, ptr As Long
   ptr = sample * wFormat.nBlockAlign + bufferIn
   CopyStructFromPtr sample16, ptr, 2
   leftVol = sample16 / 32768
   CopyStructFromPtr sample16, ptr + 2, 2
   rightVol = sample16 / 32768
End Sub

Private Sub GetStereo8Sample(ByVal sample As Long, ByRef leftVol As Double, ByRef rightVol As Double)
   Dim sample8 As Byte, ptr As Long
   ptr = sample * wFormat.nBlockAlign + bufferIn
   CopyStructFromPtr sample8, ptr, 1
   leftVol = (sample8 - 128) / 128
   CopyStructFromPtr sample8, ptr + 1, 1
   rightVol = (sample8 - 128) / 128
End Sub

Private Sub GetMono16Sample(ByVal sample As Long, ByRef leftVol As Double)
   Dim sample16 As Integer, ptr As Long
   ptr = sample * wFormat.nBlockAlign + bufferIn
   CopyStructFromPtr sample16, ptr, 2
   leftVol = sample16 / 32768
End Sub

Private Sub GetMono8Sample(ByVal sample As Long, ByRef leftVol As Double)
   Dim sample8 As Byte, ptr As Long
   ptr = sample * wFormat.nBlockAlign + bufferIn
   CopyStructFromPtr sample8, ptr, 1
   leftVol = (sample8 - 128) / 128
End Sub
' =====================================================================================
' Wave �ļ���������
' =====================================================================================
Public Sub DrawWaves(strFileName As String, picBox As PictureBox, Optional ByVal lineColor As OLE_COLOR = vbBlack)
   ' if no file is loaded, don't try to draw graph
   If Not LoadFile(strFileName) Then Exit Sub
   
    ' Graph the waveform
    Dim X As Long               ' current X position
    Dim leftYOffset As Long     ' Y offset for left channel graph
    Dim rightYOffset As Long    ' Y offset for right channel graph
    Dim curLeftY As Long        ' current left channel Y value
    Dim curRightY As Long       ' current right channel Y value
    Dim lastX As Long           ' last X position
    Dim lastLeftY As Long       ' last left channel Y value
    Dim lastRightY As Long      ' last right channel Y value
    Dim maxAmplitude As Long    ' the maximum amplitude for a wave graph on the form
    Dim leftVol As Double       ' buffer for retrieving the left volume level
    Dim rightVol As Double      ' buffer for retrieving the right volume level
    Dim scaleFactor As Double   ' samples per pixel on the wave graph
    Dim xStep As Double         ' pixels per sample on the wave graph
    Dim curSample As Long       ' current sample number
    Dim oldSM As ScaleModeConstants ' ͼƬ��ɵ� ScaleMode ֵ��
    Dim oldFC As OLE_COLOR ' �ɵ���ɫ
    
    ' clear the screen
    picBox.AutoRedraw = True: picBox.Cls
    ' ���û�ͼ������ɫ
    oldFC = picBox.ForeColor
    picBox.ForeColor = lineColor
    ' ScaleMode һ��Ҫ���ã����򣬻�ͼ���ԣ�����
    oldSM = picBox.ScaleMode
    picBox.ScaleMode = vbTwips ' �����µ� ScaleMode ֵ
    
    drawFrom = 0
    drawTo = numSamples
    
    ' calculate drawing parameters
    scaleFactor = (drawTo - drawFrom) / picBox.Width
    If (scaleFactor < 1) Then
        xStep = 1 / scaleFactor
    Else
        xStep = 1
    End If

    ' Draw the graph
    If (wFormat.nChannels = 2) Then ' �����˫����
        maxAmplitude = picBox.Height / 4
        leftYOffset = maxAmplitude
        rightYOffset = maxAmplitude * 3
         
        For X = 0 To picBox.Width Step xStep
            curSample = scaleFactor * X + drawFrom
            If (wFormat.wBitsPerSample = 16) Then
                GetStereo16Sample curSample, leftVol, rightVol
            Else
                GetStereo8Sample curSample, leftVol, rightVol
            End If
            curRightY = CLng(rightVol * maxAmplitude)
            curLeftY = CLng(leftVol * maxAmplitude)
            picBox.Line (lastX, leftYOffset + lastLeftY)-(X, curLeftY + leftYOffset)
            picBox.Line (lastX, rightYOffset + lastRightY)-(X, curRightY + rightYOffset)
            lastLeftY = curLeftY
            lastRightY = curRightY
            lastX = X
        Next
    Else ' ��������ֻ��Ҫ��һ��ͼ
        maxAmplitude = picBox.Height / 2
        leftYOffset = maxAmplitude
        
        For X = 0 To picBox.Width Step xStep
           curSample = scaleFactor * X + drawFrom
           If (wFormat.wBitsPerSample = 16) Then
               GetMono16Sample curSample, leftVol
           Else
               GetMono8Sample curSample, leftVol
           End If
           curLeftY = CLng(leftVol * maxAmplitude)
           picBox.Line (lastX, leftYOffset + lastLeftY)-(X, curLeftY + leftYOffset)
           lastLeftY = curLeftY
           lastX = X
        Next
    End If

    ' ��ԭͼƬ��ɵ� ScaleMode ֵ��
    picBox.ScaleMode = oldSM
    picBox.ForeColor = oldFC
End Sub
' =====================================================================================
' Wave �ļ���������
' =====================================================================================
