Attribute VB_Name = "HhAPi"

Public Declare Function GetLastError Lib "kernel32" () As Long

Public Declare Function LocalAlloc _
               Lib "kernel32" (ByVal uFlags As Long, _
                               ByVal uBytes As Long) As Long

Public Declare Function EmptyClipboard Lib "user32" () As Long

Public Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long

Public Declare Function CloseClipboard Lib "user32" () As Long

Public Declare Function SetClipboardData _
               Lib "user32" (ByVal wFormat As Long, _
                             ByVal hMem As Long) As Long


Public Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long

Public Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long

Public Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long

Public Declare Function GlobalAlloc _
               Lib "kernel32" (ByVal wFlags As Long, _
                               ByVal dwBytes As Long) As Long

Public Const CF_TEXT = 1

Public Const CF_UNICODETEXT = 13

Public Const GMEM_FIXED = &H0

Public Const GMEM_MOVEABLE = &H2

Public Const GMEM_ZEROINIT = &H40

Public Const GHND = (GMEM_MOVEABLE Or GMEM_ZEROINIT)

Public Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)

Public Declare Sub ZeroMemory _
               Lib "kernel32" _
               Alias "RtlZeroMemory" (ByVal Destination As Long, _
                                      ByVal cch As Long)
'---------------------------------------------------
'UTF8 ±àÂë
'--------------------APIÉùÃ÷²¿·Ö--------------------
Public Declare Function WideCharToMultiByte _
               Lib "kernel32" (ByVal CodePage As Long, _
                               ByVal dwFlags As Long, _
                               ByVal lpWideCharStr As Long, _
                               ByVal cchWideChar As Long, _
                               ByRef lpMultiByteStr As Any, _
                               ByVal cchMultiByte As Long, _
                               ByVal lpDefaultChar As String, _
                               ByVal lpUsedDefaultChar As Long) As Long

Public Const CP_UTF8 = 65001

Public Function UTF8_Encode(ByVal strUnicode As String) As Byte()   'UTF-8 ±àÂë

    Dim TLen          As Long

    Dim lngBufferSize As Long

    Dim lngResult     As Long

    Dim bytUtf8()     As Byte

    TLen = Len(strUnicode)

    If TLen = 0 Then Exit Function
    lngBufferSize = TLen * 3 + 1
    ReDim bytUtf8(lngBufferSize - 1)
    lngResult = WideCharToMultiByte(CP_UTF8, 0, StrPtr(strUnicode), TLen, bytUtf8(0), lngBufferSize, vbNullString, 0)

    If lngResult <> 0 Then
        lngResult = lngResult  '// -1
        ReDim Preserve bytUtf8(lngResult)
    End If

    UTF8_Encode = bytUtf8
End Function




Public Function ClipboardSetData(ByVal hWnd As Long, _
                                  ByRef NewData() As Byte, _
                                  ByVal DataType As Long) As Boolean

    On Error GoTo errhdl__

    Dim dwLength As Long

    Dim hGlobal  As Long, lpGlobal As Long

    dwLength = UBound(NewData) - LBound(NewData) + 1
    hGlobal = GlobalAlloc(GHND, dwLength)

    If (hGlobal = 0) Then ClipboardSetData = False: Exit Function
    lpGlobal = GlobalLock(hGlobal)

    If (lpGlobal = 0) Then

        GlobalFree (hGlobal)

        ClipboardSetData = False: Exit Function
    End If

    Call ZeroMemory(lpGlobal, dwLength)
    Call CopyMemory(ByVal lpGlobal, ByVal VarPtr(NewData(0)), dwLength)

    If (OpenClipboard(hWnd) = 0) Then

        GlobalUnlock (hGlobal)

        GlobalFree (hGlobal)

        MsgBox "unable to open clipboard"
        ClipboardSetData = False: Exit Function
    End If

    EmptyClipboard

    If (SetClipboardData(DataType, hGlobal) = 0) Then

        GlobalUnlock (hGlobal)

        CloseClipboard
        MsgBox "unable to set clipboard data " & dwLength & " " & hGlobal & " " & GetLastError & " " & Err.LastDllError
        ClipboardSetData = False: Exit Function
    End If

    GlobalUnlock (hGlobal)

    CloseClipboard
    ClipboardSetData = True

    Exit Function

errhdl__:
    CloseClipboard
    MsgBox Err.Description
End Function

Public Function ClipboardSetText(ByVal hWnd As Long, ByRef szText() As Byte) As Boolean
    ReDim Preserve szText(UBound(szText) + LenB(vbNullChar))
    ClipboardSetText = ClipboardSetData(hWnd, szText, CF_TEXT)
End Function

