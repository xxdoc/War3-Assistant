Attribute VB_Name = "收邮件并解码"
Option Explicit
Public Const cstBase64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
Public arrBase64() As String
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
Private Const CP_ACP = 0        ' default to ANSI code page
Private Const CP_UTF8 = 65001   ' default to UTF-8 code page
Public Function Base64Encode(strSource As String) As String '编码
On Error Resume Next
If UBound(arrBase64) = -1 Then
arrBase64 = Split(StrConv(cstBase64, vbUnicode), vbNullChar)
End If
Dim arrB() As Byte, bTmp(2)  As Byte, bT As Byte
Dim i As Long, J As Long
arrB = StrConv(strSource, vbFromUnicode)
J = UBound(arrB)
For i = 0 To J Step 3
Erase bTmp
bTmp(0) = arrB(i + 0)
bTmp(1) = arrB(i + 1)
bTmp(2) = arrB(i + 2)
bT = (bTmp(0) And 252) / 4
Base64Encode = Base64Encode & arrBase64(bT)
bT = (bTmp(0) And 3) * 16
bT = bT + bTmp(1) \ 16
Base64Encode = Base64Encode & arrBase64(bT)
bT = (bTmp(1) And 15) * 4
bT = bT + bTmp(2) \ 64
If i + 1 <= J Then
Base64Encode = Base64Encode & arrBase64(bT)
Else
Base64Encode = Base64Encode & "="
End If
bT = bTmp(2) And 63
If i + 2 <= J Then
Base64Encode = Base64Encode & arrBase64(bT)
Else
Base64Encode = Base64Encode & "="
End If
Next
End Function
Public Function Base64Decode(strEncoded As String) As String '解码
On Error Resume Next
Dim arrB() As Byte, bTmp(3)  As Byte, bT As Long, bRet() As Byte
Dim i As Long, J As Long
arrB = StrConv(strEncoded, vbFromUnicode)
J = InStr(strEncoded & "=", "=") - 2
ReDim bRet(J - J \ 4 - 1)
For i = 0 To J Step 4
Erase bTmp
bTmp(0) = (InStr(cstBase64, Chr(arrB(i))) - 1) And 63
bTmp(1) = (InStr(cstBase64, Chr(arrB(i + 1))) - 1) And 63
bTmp(2) = (InStr(cstBase64, Chr(arrB(i + 2))) - 1) And 63
bTmp(3) = (InStr(cstBase64, Chr(arrB(i + 3))) - 1) And 63
bT = bTmp(0) * 2 ^ 18 + bTmp(1) * 2 ^ 12 + bTmp(2) * 2 ^ 6 + bTmp(3)
bRet((i \ 4) * 3) = bT \ 65536
bRet((i \ 4) * 3 + 1) = (bT And 65280) \ 256
bRet((i \ 4) * 3 + 2) = bT And 255
Next
Base64Decode = StrConv(bRet, vbUnicode)
End Function
Function StrToBytes(ByVal Source As String) As Byte()
Dim bB64Str() As Byte
bB64Str = StrConv(Source, vbFromUnicode)
Dim lB64Len As Long
lB64Len = InStrB(bB64Str, ChrB$(Asc("="))) - 1
Dim lLenPad As Long
lLenPad = (4 - lB64Len Mod 4) Mod 4
Dim lLen As Long
lLen = ((lB64Len + lLenPad) \ 4) * 3
Dim bStr() As Byte
ReDim bStr(lLen - 1)
lLen = lLen - lLenPad
Dim i As Long
Dim lBuffer As Long
For i = 0 To lB64Len - 1 Step 4
lBuffer = DeB64CodeA(bB64Str(i + 0)) * &H40000 Or DeB64CodeA(bB64Str(i + 1)) * &H1000& _
Or DeB64CodeA(bB64Str(i + 2)) * &H40& Or DeB64CodeA(bB64Str(i + 3))
bStr((i \ 4) * 3 + 2) = lBuffer And &HFF&
lBuffer = lBuffer \ &H100&
bStr((i \ 4) * 3 + 1) = lBuffer And &HFF&
lBuffer = lBuffer \ &H100&
bStr((i \ 4) * 3 + 0) = lBuffer And &HFF&
lBuffer = lBuffer \ &H100&
Next
ReDim Preserve bStr(lLen - 1)
StrToBytes = bStr
End Function
Private Function DeB64CodeA(ByVal Char As Byte) As Byte
Select Case Char
Case Asc("A") To Asc("Z"): DeB64CodeA = Char - Asc("A")
Case Asc("a") To Asc("z"): DeB64CodeA = Char - Asc("a") + 26
Case Asc("0") To Asc("9"): DeB64CodeA = Char - Asc("0") + 52
Case Asc("+"): DeB64CodeA = 62
Case Asc("/"): DeB64CodeA = 63
Case Asc("="): DeB64CodeA = 64
End Select
End Function
Function Utf8ToUnicode(ByRef Utf() As Byte) As String
Dim lRet As Long
Dim lLength As Long
Dim lBufferSize As Long
lLength = UBound(Utf) - LBound(Utf) + 1
If lLength <= 0 Then Exit Function
lBufferSize = lLength * 2
Utf8ToUnicode = String$(lBufferSize, Chr(0))
lRet = MultiByteToWideChar(CP_UTF8, 0, VarPtr(Utf(0)), lLength, StrPtr(Utf8ToUnicode), lBufferSize)
If lRet <> 0 Then
Utf8ToUnicode = left(Utf8ToUnicode, lRet)
End If
End Function

Sub Rec()
Dim att As Object
Dim EmailMsg As Object
Dim atts As Object
Dim JMail As Object
Dim EmailCont$, Subject$, EmailID&, i As Integer
Dim X$()
    Dim J#
    Set JMail = CreateObject("JMail.POP3")
    JMail.Connect "faithdmc@163.com", "110120130", "pop.163.com", "110" 'JMail.Connect "邮箱名", "密码", "服务器" [,"端口号"]
    Debug.Print "你有" & JMail.Count & "封邮件"     '邮件数量
    For i = JMail.Count To 1 Step -1
Set EmailMsg = JMail.Messages.Item(i)       '取得一条邮件信息
        Subject = EmailMsg.Headers.GetHeader("Subject") '邮件标题，可正常解码，但UTF-8格式的标题取不全
        X = Split(Subject, "?")
        If UBound(X) > 3 Then
            If X(1) = "UTF-8" Then
                Subject = Utf8ToUnicode(StrToBytes(X(3)))
            Else
                Subject = Base64Decode(X(3))
            End If
        End If
        If Subject = "公告" Then EmailCont = EmailMsg.Body & vbCrLf & vbCrLf & EmailMsg.Date: Exit For
        DoEvents
    Next
    EmailCont = Replace(EmailCont, vbCrLf & " ", vbCrLf)
    If EmailCont = "" Then EmailCont = "暂无公告"
    Form4.Text1.Text = EmailCont
'------------------------------------------------------------------------------以下为各种参数设置
'        EmailMsg.Charset = "gb2312"'编码方式
'        EmailMsg.ContentTransferEncoding = "base64"'解码方式
'        EmailMsg.Encoding = "base64"
'        EmailMsg.ContentType = "multipart/mixed"   '发送邮件时
'        EmailMsg.ContentType = "text/html"         '接收邮件时
'        EmailMsg.ISOEncodeHeaders = False'True     '功能不清？
'-----------------------------------------------------------------------------可以取得的各元素
'        MsgBox EmailMsg.Priority                   '邮件的优先级，1-5，1最高，正常情况为3。
'        MsgBox EmailMsg.From                       '邮件的发送人的信箱地址
'        MsgBox EmailMsg.FromName                   '邮件的发送人
'        MsgBox EmailMsg.Date                       '邮件日期
'        MsgBox EmailMsg.Body                       '邮件内容
'        MsgBox EmailMsg.Size                       '邮件大小'----------------------------------------------------------------------------

End Sub




