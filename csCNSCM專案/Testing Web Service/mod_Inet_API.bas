Attribute VB_Name = "mod_Inet_API"
Option Explicit

Public Const MAX_PATH  As Long = 260
Public Const FILE_ATTRIBUTE_ARCHIVE = &H20

'use registry configuration
Public Const INTERNET_OPEN_TYPE_PRECONFIG = 0
'direct to net
Public Const INTERNET_OPEN_TYPE_DIRECT = 1
'via named proxy
Public Const INTERNET_OPEN_TYPE_PROXY = 3
'prevent using java/script/INS
Public Const INTERNET_OPEN_TYPE_PRECONFIG_WITH_NO_AUTOPROXY = 4
'used for FTP connections
Public Const INTERNET_FLAG_PASSIVE = &H8000000
Public Const INTERNET_FLAG_RELOAD = &H80000000

'additional cache flags
'don't write this item to the cache
Public Const INTERNET_FLAG_NO_CACHE_WRITE = &H4000000
Public Const INTERNET_FLAG_DONT_CACHE = INTERNET_FLAG_NO_CACHE_WRITE
'make this item persistent in cache
Public Const INTERNET_FLAG_MAKE_PERSISTENT = &H2000000
'use offline semantics
Public Const INTERNET_FLAG_FROM_CACHE = &H1000000
Public Const INTERNET_FLAG_OFFLINE = INTERNET_FLAG_FROM_CACHE

'additional flags
'use PCT/SSL if applicable (HTTP)
Public Const INTERNET_FLAG_SECURE = &H800000
'use keep-alive semantics
Public Const INTERNET_FLAG_KEEP_CONNECTION = &H400000
'don't handle redirections automatically
Public Const INTERNET_FLAG_NO_AUTO_REDIRECT = &H200000
'do background read prefetch
Public Const INTERNET_FLAG_READ_PREFETCH = &H100000
'no automatic cookie handling
Public Const INTERNET_FLAG_NO_COOKIES = &H80000
'no automatic authentication handling
Public Const INTERNET_FLAG_NO_AUTH = &H40000
'return cache file if net request fails
Public Const INTERNET_FLAG_CACHE_IF_NET_FAIL = &H10000
'return cache file if net request fails
Public Const INTERNET_DEFAULT_FTP_PORT = 21
'   "     "  gopher "
Public Const INTERNET_DEFAULT_GOPHER_PORT = 70
'   "     "  HTTP   "
Public Const INTERNET_DEFAULT_HTTP_PORT = 80
'   "     "  HTTPS  "
Public Const INTERNET_DEFAULT_HTTPS_PORT = 443
'default for SOCKS firewall servers.
Public Const INTERNET_DEFAULT_SOCKS_PORT = 1080
'FTP: use existing InternetConnect handle for server if possible
Public Const INTERNET_FLAG_EXISTING_CONNECT = &H20000000
Public Const INTERNET_SERVICE_FTP = 1
Public Const INTERNET_SERVICE_GOPHER = 2
Public Const INTERNET_SERVICE_HTTP = 3

'transfer flags
Public Const FTP_TRANSFER_TYPE_UNKNOWN = &H0
Public Const FTP_TRANSFER_TYPE_ASCII = &H1
Public Const FTP_TRANSFER_TYPE_BINARY = &H2
Public Const INTERNET_FLAG_TRANSFER_ASCII = FTP_TRANSFER_TYPE_ASCII
Public Const INTERNET_FLAG_TRANSFER_BINARY = FTP_TRANSFER_TYPE_BINARY
Public Const FTP_TRANSFER_TYPE_MASK = (FTP_TRANSFER_TYPE_ASCII Or _
                                       FTP_TRANSFER_TYPE_BINARY)

'internet error flags
Public Const INTERNET_ERROR_BASE = 12000
Public Const ERROR_INTERNET_OUT_OF_HANDLES = (INTERNET_ERROR_BASE + 1)
Public Const ERROR_INTERNET_TIMEOUT = (INTERNET_ERROR_BASE + 2)
Public Const ERROR_INTERNET_EXTENDED_ERROR = (INTERNET_ERROR_BASE + 3)
Public Const ERROR_INTERNET_INTERNAL_ERROR = (INTERNET_ERROR_BASE + 4)
Public Const ERROR_INTERNET_INVALID_URL = (INTERNET_ERROR_BASE + 5)
Public Const ERROR_INTERNET_UNRECOGNIZED_SCHEME = (INTERNET_ERROR_BASE + 6)
Public Const ERROR_INTERNET_NAME_NOT_RESOLVED = (INTERNET_ERROR_BASE + 7)
Public Const ERROR_INTERNET_PROTOCOL_NOT_FOUND = (INTERNET_ERROR_BASE + 8)
Public Const ERROR_INTERNET_INVALID_OPTION = (INTERNET_ERROR_BASE + 9)
Public Const ERROR_INTERNET_BAD_OPTION_LENGTH = (INTERNET_ERROR_BASE + 10)
Public Const ERROR_INTERNET_OPTION_NOT_SETTABLE = (INTERNET_ERROR_BASE + 11)
Public Const ERROR_INTERNET_SHUTDOWN = (INTERNET_ERROR_BASE + 12)
Public Const ERROR_INTERNET_INCORRECT_USER_NAME = (INTERNET_ERROR_BASE + 13)
Public Const ERROR_INTERNET_INCORRECT_PASSWORD = (INTERNET_ERROR_BASE + 14)
Public Const ERROR_INTERNET_LOGIN_FAILURE = (INTERNET_ERROR_BASE + 15)
Public Const ERROR_INTERNET_INVALID_OPERATION = (INTERNET_ERROR_BASE + 16)
Public Const ERROR_INTERNET_OPERATION_CANCELLED = (INTERNET_ERROR_BASE + 17)
Public Const ERROR_INTERNET_INCORRECT_HANDLE_TYPE = (INTERNET_ERROR_BASE + 18)
Public Const ERROR_INTERNET_INCORRECT_HANDLE_STATE = (INTERNET_ERROR_BASE + 19)
Public Const ERROR_INTERNET_NOT_PROXY_REQUEST = (INTERNET_ERROR_BASE + 20)
Public Const ERROR_INTERNET_REGISTRY_VALUE_NOT_FOUND = (INTERNET_ERROR_BASE + 21)
Public Const ERROR_INTERNET_BAD_REGISTRY_PARAMETER = (INTERNET_ERROR_BASE + 22)
Public Const ERROR_INTERNET_NO_DIRECT_ACCESS = (INTERNET_ERROR_BASE + 23)
Public Const ERROR_INTERNET_NO_CONTEXT = (INTERNET_ERROR_BASE + 24)
Public Const ERROR_INTERNET_NO_CALLBACK = (INTERNET_ERROR_BASE + 25)
Public Const ERROR_INTERNET_REQUEST_PENDING = (INTERNET_ERROR_BASE + 26)
Public Const ERROR_INTERNET_INCORRECT_FORMAT = (INTERNET_ERROR_BASE + 27)
Public Const ERROR_INTERNET_ITEM_NOT_FOUND = (INTERNET_ERROR_BASE + 28)
Public Const ERROR_INTERNET_CANNOT_CONNECT = (INTERNET_ERROR_BASE + 29)
Public Const ERROR_INTERNET_CONNECTION_ABORTED = (INTERNET_ERROR_BASE + 30)
Public Const ERROR_INTERNET_CONNECTION_RESET = (INTERNET_ERROR_BASE + 31)
Public Const ERROR_INTERNET_FORCE_RETRY = (INTERNET_ERROR_BASE + 32)
Public Const ERROR_INTERNET_INVALID_PROXY_REQUEST = (INTERNET_ERROR_BASE + 33)
Public Const ERROR_INTERNET_NEED_UI = (INTERNET_ERROR_BASE + 34)
Public Const ERROR_INTERNET_HANDLE_EXISTS = (INTERNET_ERROR_BASE + 36)
Public Const ERROR_INTERNET_SEC_CERT_DATE_INVALID = (INTERNET_ERROR_BASE + 37)
Public Const ERROR_INTERNET_SEC_CERT_CN_INVALID = (INTERNET_ERROR_BASE + 38)
Public Const ERROR_INTERNET_HTTP_TO_HTTPS_ON_REDIR = (INTERNET_ERROR_BASE + 39)
Public Const ERROR_INTERNET_HTTPS_TO_HTTP_ON_REDIR = (INTERNET_ERROR_BASE + 40)
Public Const ERROR_INTERNET_MIXED_SECURITY = (INTERNET_ERROR_BASE + 41)
Public Const ERROR_INTERNET_CHG_POST_IS_NON_SECURE = (INTERNET_ERROR_BASE + 42)
Public Const ERROR_INTERNET_POST_IS_NON_SECURE = (INTERNET_ERROR_BASE + 43)
Public Const ERROR_INTERNET_CLIENT_AUTH_CERT_NEEDED = (INTERNET_ERROR_BASE + 44)
Public Const ERROR_INTERNET_INVALID_CA = (INTERNET_ERROR_BASE + 45)
Public Const ERROR_INTERNET_CLIENT_AUTH_NOT_SETUP = (INTERNET_ERROR_BASE + 46)
Public Const ERROR_INTERNET_ASYNC_THREAD_FAILED = (INTERNET_ERROR_BASE + 47)
Public Const ERROR_INTERNET_REDIRECT_SCHEME_CHANGE = (INTERNET_ERROR_BASE + 48)
Public Const ERROR_INTERNET_DIALOG_PENDING = (INTERNET_ERROR_BASE + 49)
Public Const ERROR_INTERNET_RETRY_DIALOG = (INTERNET_ERROR_BASE + 50)
Public Const ERROR_INTERNET_HTTPS_HTTP_SUBMIT_REDIR = (INTERNET_ERROR_BASE + 52)
Public Const ERROR_INTERNET_INSERT_CDROM = (INTERNET_ERROR_BASE + 53)
Public Const ERROR_INTERNET_FORTEZZA_LOGIN_NEEDED = (INTERNET_ERROR_BASE + 54)
Public Const ERROR_INTERNET_SEC_CERT_ERRORS = (INTERNET_ERROR_BASE + 55)
Public Const ERROR_INTERNET_SEC_CERT_NO_REV = (INTERNET_ERROR_BASE + 56)
Public Const ERROR_INTERNET_SEC_CERT_REV_FAILED = (INTERNET_ERROR_BASE + 57)

'FTP API errors
Public Const ERROR_FTP_TRANSFER_IN_PROGRESS = (INTERNET_ERROR_BASE + 110)
Public Const ERROR_FTP_DROPPED = (INTERNET_ERROR_BASE + 111)
Public Const ERROR_FTP_NO_PASSIVE_MODE = (INTERNET_ERROR_BASE + 112)

'gopher API errors
Public Const ERROR_GOPHER_PROTOCOL_ERROR = (INTERNET_ERROR_BASE + 130)
Public Const ERROR_GOPHER_NOT_FILE = (INTERNET_ERROR_BASE + 131)
Public Const ERROR_GOPHER_DATA_ERROR = (INTERNET_ERROR_BASE + 132)
Public Const ERROR_GOPHER_END_OF_DATA = (INTERNET_ERROR_BASE + 133)
Public Const ERROR_GOPHER_INVALID_LOCATOR = (INTERNET_ERROR_BASE + 134)
Public Const ERROR_GOPHER_INCORRECT_LOCATOR_TYPE = (INTERNET_ERROR_BASE + 135)
Public Const ERROR_GOPHER_NOT_GOPHER_PLUS = (INTERNET_ERROR_BASE + 136)
Public Const ERROR_GOPHER_ATTRIBUTE_NOT_FOUND = (INTERNET_ERROR_BASE + 137)
Public Const ERROR_GOPHER_UNKNOWN_LOCATOR = (INTERNET_ERROR_BASE + 138)

'HTTP API errors
Public Const ERROR_HTTP_HEADER_NOT_FOUND = (INTERNET_ERROR_BASE + 150)
Public Const ERROR_HTTP_DOWNLEVEL_SERVER = (INTERNET_ERROR_BASE + 151)
Public Const ERROR_HTTP_INVALID_SERVER_RESPONSE = (INTERNET_ERROR_BASE + 152)
Public Const ERROR_HTTP_INVALID_HEADER = (INTERNET_ERROR_BASE + 153)
Public Const ERROR_HTTP_INVALID_QUERY_REQUEST = (INTERNET_ERROR_BASE + 154)
Public Const ERROR_HTTP_HEADER_ALREADY_EXISTS = (INTERNET_ERROR_BASE + 155)
Public Const ERROR_HTTP_REDIRECT_FAILED = (INTERNET_ERROR_BASE + 156)
Public Const ERROR_HTTP_NOT_REDIRECTED = (INTERNET_ERROR_BASE + 160)
Public Const ERROR_HTTP_COOKIE_NEEDS_CONFIRMATION = (INTERNET_ERROR_BASE + 161)
Public Const ERROR_HTTP_COOKIE_DECLINED = (INTERNET_ERROR_BASE + 162)
Public Const ERROR_HTTP_REDIRECT_NEEDS_CONFIRMATION = (INTERNET_ERROR_BASE + 168)

'additional Internet API error codes
Public Const ERROR_INTERNET_SECURITY_CHANNEL_ERROR = (INTERNET_ERROR_BASE + 157)
Public Const ERROR_INTERNET_UNABLE_TO_CACHE_FILE = (INTERNET_ERROR_BASE + 158)
Public Const ERROR_INTERNET_TCPIP_NOT_INSTALLED = (INTERNET_ERROR_BASE + 159)
Public Const ERROR_INTERNET_DISCONNECTED = (INTERNET_ERROR_BASE + 163)
Public Const ERROR_INTERNET_SERVER_UNREACHABLE = (INTERNET_ERROR_BASE + 164)
Public Const ERROR_INTERNET_PROXY_SERVER_UNREACHABLE = (INTERNET_ERROR_BASE + 165)
Public Const ERROR_INTERNET_BAD_AUTO_PROXY_SCRIPT = (INTERNET_ERROR_BASE + 166)
Public Const ERROR_INTERNET_UNABLE_TO_DOWNLOAD_SCRIPT = (INTERNET_ERROR_BASE + 167)
Public Const ERROR_INTERNET_SEC_INVALID_CERT = (INTERNET_ERROR_BASE + 169)
Public Const ERROR_INTERNET_SEC_CERT_REVOKED = (INTERNET_ERROR_BASE + 170)

'InternetAutodial specific errors
Public Const ERROR_INTERNET_FAILED_DUETOSECURITYCHECK = (INTERNET_ERROR_BASE + 171)
Public Const ERROR_INTERNET_NOT_INITIALIZED = (INTERNET_ERROR_BASE + 172)
Public Const ERROR_INTERNET_NEED_MSN_SSPI_PKG = (INTERNET_ERROR_BASE + 173)
Public Const ERROR_INTERNET_LOGIN_FAILURE_DISPLAY_ENTITY_BODY = (INTERNET_ERROR_BASE + 174)
Public Const INTERNET_ERROR_LAST = (INTERNET_ERROR_BASE + 174)

Public Type FILETIME
   dwLowDateTime As Long
   dwHighDateTime As Long
End Type


Public Type WIN32_FIND_DATA
   dwFileAttributes As Long
   ftCreationTime As FILETIME
   ftLastAccessTime As FILETIME
   ftLastWriteTime As FILETIME
   nFileSizeHigh As Long
   nFileSizeLow As Long
   dwReserved0 As Long
   dwReserved1 As Long
   cFileName As String * MAX_PATH
   cAlternate As String * 14
End Type

Public Declare Function FtpFindFirstFile Lib "wininet.dll" _
   Alias "FtpFindFirstFileA" _
  (ByVal hConnect As Long, _
   ByVal lpszSearchFile As String, _
   lpFindFileData As Any, _
   ByVal dwFlags As Long, _
   ByVal dwContext As Long) As Long

Public Declare Function InternetFindNextFile Lib "wininet.dll" Alias "InternetFindNextFileA" _
  (ByVal hFind As Long, _
   lpFindFileData As WIN32_FIND_DATA) As Long

Public Declare Function InternetGetLastResponseInfo Lib "wininet.dll" Alias "InternetGetLastResponseInfoA" _
  (lpdwError As Long, _
   ByVal lpszBuffer As String, _
    lpdwBufferLength As Long) As Long
   
Public Declare Function FtpGetCurrentDirectory Lib "wininet.dll" Alias "FtpGetCurrentDirectoryA" _
  (ByVal hConnect As Long, _
   ByVal lpszCurrentDirectory As String, _
    lpdwCurrentDirectory As Long) As Long

Public Declare Function FtpSetCurrentDirectory Lib "wininet.dll" Alias "FtpSetCurrentDirectoryA" _
  (ByVal hConnect As Long, _
   ByVal lpszDirectory As String) As Long

Public Declare Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" _
   (ByVal hInternetSession As Long, _
   ByVal sUrl As String, ByVal sHeaders As String, _
   ByVal lHeadersLength As Long, ByVal lFlags As Long, _
   ByVal lContext As Long) As Long
   
Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" _
   (ByVal lpszCallerName As String, _
    ByVal dwAccessType As Long, _
    ByVal lpszProxyName As String, _
    ByVal lpszProxyBypass As String, _
    ByVal dwFlags As Long) As Long

Private Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" _
   (ByVal hInternetSession As Long, _
    ByVal lpszServerName As String, _
    ByVal nProxyPort As Integer, _
    ByVal lpszUsername As String, _
    ByVal lpszPassword As String, _
    ByVal dwService As Long, _
    ByVal dwFlags As Long, _
    ByVal dwContext As Long) As Long

Private Declare Function InternetReadFile Lib "wininet.dll" _
   (ByVal hFile As Long, _
    ByVal sBuffer As String, _
    ByVal lNumBytesToRead As Long, _
    lNumberOfBytesRead As Long) As Integer

Private Declare Function HttpOpenRequest Lib "wininet.dll" Alias "HttpOpenRequestA" _
   (ByVal hInternetSession As Long, _
    ByVal lpszVerb As String, _
    ByVal lpszObjectName As String, _
    ByVal lpszVersion As String, _
    ByVal lpszReferer As String, _
    ByVal lpszAcceptTypes As Long, _
    ByVal dwFlags As Long, _
    ByVal dwContext As Long) As Long

Private Declare Function HttpSendRequest Lib "wininet.dll" Alias "HttpSendRequestA" _
   (ByVal hHttpRequest As Long, _
    ByVal sHeaders As String, _
    ByVal lHeadersLength As Long, _
    ByVal sOptional As String, _
    ByVal lOptionalLength As Long) As Boolean

Private Declare Function InternetCloseHandle Lib "wininet.dll" _
   (ByVal hInternetHandle As Long) As Boolean

Private Declare Function HttpAddRequestHeaders Lib "wininet.dll" Alias "HttpAddRequestHeadersA" _
   (ByVal hHttpRequest As Long, _
   ByVal sHeaders As String, _
   ByVal lHeadersLength As Long, _
   ByVal lModifiers As Long) As Integer

Public Function GetInetError(ByVal lErrorCode As Long) As String

'''''''''''''''''''''''''''''''''''''''''''''''''''''
' From       : MSDN
' Name       : GetInetError (Formerly TranslateErrorCode)
' Purpose    : Provides message for DLL error codes
' Parameters : The DLL error code
' Return val : String containing message
' Algorithm : Selects the appropriate string
''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim sBuffer As String
Dim nBuffer As Long

Select Case lErrorCode
   Case 12001: GetInetError = _
      "No more handles could be generated at this time"
   Case 12002: GetInetError = _
      "The request has timed out."
   Case 12003:
      'extended error. Retrieve the details using
      'the InternetGetLastResponseInfo API.
      
      sBuffer = Space$(256)
      nBuffer = Len(sBuffer)
      
      If InternetGetLastResponseInfo(lErrorCode, _
                                     sBuffer, _
                                     nBuffer) Then
         GetInetError = StripNull(sBuffer)
      Else
         GetInetError = "Extended error returned from server."
      End If
      
   Case 12004: GetInetError = _
      "An internal error has occurred."
   Case 12005: GetInetError = _
      "URL is invalid."
   Case 12006: GetInetError = _
      "URL scheme could not be recognized, or is not supported."
   Case 12007: GetInetError = _
      "Server name could not be resolved."
   Case 12008: GetInetError = _
      "Requested protocol could not be located."
   Case 12009: GetInetError = _
      "Request to InternetQueryOption or InternetSetOption" & _
      " specified an invalid option value."
   Case 12010: GetInetError = _
      "Length of an option supplied to InternetQueryOption or" & _
      " InternetSetOption is incorrect for the type of" & _
      " option specified."
   Case 12011: GetInetError = _
      "Request option can not be set, only queried. "
   Case 12012: GetInetError = _
      "Win32 Internet support is being shutdown or unloaded."
   Case 12013: GetInetError = _
      "Request to connect and login to an FTP server could not" & _
      " be completed because the supplied user name is incorrect."
   Case 12014: GetInetError = _
      "Request to connect and login to an FTP server could not" & _
      " be completed because the supplied password is incorrect. "
   Case 12015: GetInetError = _
      "Request to connect to and login to an FTP server failed."
   Case 12016: GetInetError = _
      "Requested operation is invalid. "
   Case 12017: GetInetError = _
      "Operation was canceled, usually because the handle on" & _
      " which the request was operating was closed before the" & _
      " operation completed."
   Case 12018: GetInetError = _
      "Type of handle supplied is incorrect for this operation."
   Case 12019: GetInetError = _
      "Requested operation can not be carried out because the" & _
      " handle supplied is not in the correct state."
   Case 12020: GetInetError = _
      "Request can not be made via a proxy."
   Case 12021: GetInetError = _
      "Required registry value could not be located. "
   Case 12022: GetInetError = _
      "Required registry value was located but is an incorrect" & _
      " type or has an invalid value."
   Case 12023: GetInetError = _
      "Direct network access cannot be made at this time. "
   Case 12024: GetInetError = _
      "Asynchronous request could not be made because a zero" & _
      " context value was supplied."
   Case 12025: GetInetError = _
      "Asynchronous request could not be made because a" & _
      " callback function has not been set."
   Case 12026: GetInetError = _
      "Required operation could not be completed because" & _
      " one or more requests are pending."
   Case 12027: GetInetError = _
      "Format of the request is invalid."
   Case 12028: GetInetError = _
      "Requested item could not be located."
   Case 12029: GetInetError = _
      "Attempt to connect to the server failed."
   Case 12030: GetInetError = _
      "Connection with the server has been terminated."
   Case 12031: GetInetError = _
      "Connection with the server has been reset."
   Case 12036: GetInetError = _
      "Request failed because the handle already exists."
   Case Else: GetInetError = _
      "Error details not available."
End Select

End Function

Function StripNull(item As String)

'Return a string without the chr$(0) terminator.
Dim pos As Integer
pos = InStr(item, Chr$(0))

If pos Then
      StripNull = Left$(item, pos - 1)
Else: StripNull = item
End If

End Function
   
Public Function HttpPost(ByVal URL, ByVal PostData As String) As String
'The API functions used here are more generic then those in HttpGet and
'can handle all kinds of internet protocols and verb. However, we use it
'for HTTP POST here.

Dim hInternetOpen As Long
Dim hInternetConnect As Long
Dim hHttpOpenRequest As Long
Dim bRet As Boolean
Dim strServer As String             'the URL's server
Dim intPort As Integer              'the URL's port
Dim strPath As String               'the URL's document path
Dim oUrl As cls_Url                    'the URL helper object

Const INTERNET_DEFAULT_HTTP_PORT = 80

#If 0 Then 'for testing
strServer = "cgi3.ebay.de"
strPath = "/aw-cgi/eBayISAPI.dll?TimeShow"
intPort = 80
#End If

'split URL because we need separate pieces here
Set oUrl = New cls_Url
oUrl.Href = URL
strServer = oUrl.Host
intPort = Val(oUrl.Port)
strPath = oUrl.PagePath

'prepare WinInet
hInternetOpen = 0
hInternetConnect = 0
hHttpOpenRequest = 0

'se registry access settings.
Const INTERNET_OPEN_TYPE_PRECONFIG = 0
hInternetOpen = InternetOpen("http generic", _
                INTERNET_OPEN_TYPE_PRECONFIG, _
                vbNullString, _
                vbNullString, _
                0)

If hInternetOpen <> 0 Then
   'Type of service to access.
   Const INTERNET_SERVICE_HTTP = 3
   'Change the server to your server name
   hInternetConnect = InternetConnect(hInternetOpen, _
                      strServer, _
                      intPort, _
                      vbNullString, _
                      "HTTP/1.0", _
                      INTERNET_SERVICE_HTTP, _
                      0, _
                      0)

   If hInternetConnect <> 0 Then
    'Brings the data across the wire even if it locally cached.
     Const INTERNET_FLAG_RELOAD = &H80000000
     hHttpOpenRequest = HttpOpenRequest(hInternetConnect, _
                         "POST", _
                         strPath, _
                         "HTTP/1.0", _
                         vbNullString, _
                         0, _
                         INTERNET_FLAG_RELOAD, _
                         0)

      If hHttpOpenRequest <> 0 Then
         Dim sHeader As String
         Const HTTP_ADDREQ_FLAG_ADD = &H20000000
         Const HTTP_ADDREQ_FLAG_REPLACE = &H80000000
         sHeader = "Content-Type: application/x-www-form-urlencoded" & vbCrLf
         bRet = HttpAddRequestHeaders(hHttpOpenRequest, _
            sHeader, Len(sHeader), HTTP_ADDREQ_FLAG_REPLACE _
            Or HTTP_ADDREQ_FLAG_ADD)

         Dim lPostDataLen As Long

         lPostDataLen = Len(PostData)
         bRet = HttpSendRequest(hHttpOpenRequest, _
                vbNullString, _
                0, _
                PostData, _
                lPostDataLen)

         Dim bDoLoop             As Boolean
         Dim sReadBuffer         As String * 2048
         Dim lNumberOfBytesRead  As Long
         Dim sBuffer             As String
         bDoLoop = True
         While bDoLoop
            sReadBuffer = vbNullString
            bDoLoop = InternetReadFile(hHttpOpenRequest, _
               sReadBuffer, Len(sReadBuffer), lNumberOfBytesRead)
            sBuffer = sBuffer & _
                 Left(sReadBuffer, lNumberOfBytesRead)
            If Not CBool(lNumberOfBytesRead) Then bDoLoop = False
         Wend
         HttpPost = sBuffer
         bRet = InternetCloseHandle(hHttpOpenRequest)
      End If
      bRet = InternetCloseHandle(hInternetConnect)
   End If
   bRet = InternetCloseHandle(hInternetOpen)
End If

End Function

Public Function HttpGet(strURL) As String
'The API functions used here can only to an HTTP GET
On Error GoTo Trap

Dim hInternetSession As Long
Dim hURLFile As Long
Dim sReadBuffer            As String * 4096     ' Grab 4k at a time
Dim sBuffer                As String
Dim lNumberOfBytesRead     As Long
Dim bDoLoop As Boolean
Dim hNewFile As Long
Dim lngTotalBytes As Long

hInternetSession = InternetOpen("HttpGet", INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)

If CBool(hInternetSession) Then
    hURLFile = InternetOpenUrl(hInternetSession, _
             strURL, vbNullString, 0, INTERNET_FLAG_RELOAD, 0)
    If CBool(hURLFile) Then
        bDoLoop = True
            While bDoLoop
                sReadBuffer = ""
                bDoLoop = InternetReadFile(hURLFile, sReadBuffer, _
                            Len(sReadBuffer), lNumberOfBytesRead)
                sBuffer = sBuffer & Left$(sReadBuffer, lNumberOfBytesRead)
                If Not CBool(lNumberOfBytesRead) Then bDoLoop = False
                DoEvents
                lngTotalBytes = lngTotalBytes + lNumberOfBytesRead
                Debug.Print lngTotalBytes / 1024
            Wend
            HttpGet = sBuffer
    End If
End If

InternetCloseHandle (hURLFile)
InternetCloseHandle (hInternetSession)

Exit Function

Trap:
MsgBox Err & " " & Err.Description, vbCritical
Exit Function

End Function




