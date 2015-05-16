Attribute VB_Name = "获取网页源码"
Public Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef dwFlags As Long, ByVal dwReserved As Long) As Long
'获取公网IP
Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" ( _
         ByVal sAgent As String, ByVal lAccessType As Long, _
         ByVal sProxyName As String, ByVal sProxyBypass As String, _
         ByVal lFlags As Long) As Long
Private Declare Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" ( _
         ByVal hInternetSession As Long, ByVal sUrl As String, _
         ByVal sHeaders As String, ByVal lHeadersLength As Long, _
         ByVal lFlags As Long, ByVal lContext As Long) As Long
Private Declare Function InternetReadFile Lib "wininet.dll" ( _
         ByVal hFile As Long, ByVal sBuffer As String, _
         ByVal lNumBytesToRead As Long, _
         lNumberOfBytesRead As Long) As Integer
Private Declare Function InternetCloseHandle Lib "wininet.dll" ( _
         ByVal hInet As Long) As Integer
Private Const INTERNET_FLAG_NO_CACHE_WRITE = &H4000000
Public Function GetUrlFile(stUrl As String) As String
     Dim lgInternet As Long, lgSession As Long
     Dim stBuf As String * 1024
     Dim inRes As Integer
     Dim lgRet As Long
     Dim stTotal As String
     stTotal = vbNullString
     lgSession = InternetOpen("VBTagEdit", 1, vbNullString, vbNullString, 0)
     If lgSession Then
         lgInternet = InternetOpenUrl(lgSession, stUrl, vbNullString, _
                 0, INTERNET_FLAG_NO_CACHE_WRITE, 0)
         If lgInternet Then
             Do
                 inRes = InternetReadFile(lgInternet, stBuf, 1024, lgRet)
                 stTotal = stTotal & Mid$(stBuf, 1, lgRet)
             Loop While (lgRet <> 0)
         End If
         inRes = InternetCloseHandle(lgInternet)
     End If
     GetUrlFile = stTotal
End Function

