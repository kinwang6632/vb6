VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGetIP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' PR : Hammer
' Start date : 2004/12/22
' Last Modify : 2004/12/22

Option Explicit

Private Const MAX_WSADescription = 256
Private Const MAX_WSASYSStatus = 128
Private Const ERROR_SUCCESS       As Long = 0
Private Const WS_VERSION_REQD     As Long = &H101
Private Const WS_VERSION_MAJOR    As Long = WS_VERSION_REQD \ &H100 And &HFF&
Private Const WS_VERSION_MINOR    As Long = WS_VERSION_REQD And &HFF&
Private Const MIN_SOCKETS_REQD    As Long = 1
Private Const SOCKET_ERROR        As Long = -1

Private Type HOSTENT
   hName      As Long
   hAliases   As Long
   hAddrType  As Integer
   hLen       As Integer
   hAddrList  As Long
End Type

Private Type WSADATA
   wVersion      As Integer
   wHighVersion  As Integer
   szDescription(0 To MAX_WSADescription)   As Byte
   szSystemStatus(0 To MAX_WSASYSStatus)    As Byte
   wMaxSockets   As Integer
   wMaxUDPDG     As Integer
   dwVendorInfo  As Long
End Type


Private Declare Function WSAGetLastError Lib "WSOCK32.DLL" () As Long
Private Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal wVersionRequired As Long, lpWSADATA As WSADATA) As Long
Private Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long
Private Declare Function gethostname Lib "WSOCK32.DLL" (ByVal szHost As String, ByVal dwHostLen As Long) As Long
Private Declare Function gethostbyname Lib "WSOCK32.DLL" (ByVal szHost As String) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)

Private blnSocketOK As Boolean

Public Function GetIPAddress(Optional ByVal strCptName As String = "") As String
  On Error GoTo ChkErr
    Dim strHostName As String * 256
    Dim lnglpHost As Long
    Dim HOST As HOSTENT
    Dim lngdwIPAddr  As Long
    Dim BytTmpIPAddr() As Byte
    Dim intLoop As Integer
    Dim strIPAddr  As String
    GetIPAddress = ""
    If Not blnSocketOK Then Exit Function
    If Len(strCptName) > 0 Then
        lnglpHost = gethostbyname(strCptName)
    Else
        If gethostname(strHostName, 256) = SOCKET_ERROR Then
           GetIPAddress = ""
           Debug.Print "Windows Sockets error " & CStr(WSAGetLastError()) & " has occurred. Unable to successfully get Host Name."
           SocketsCleanup
           Exit Function
        End If
        strHostName = Trim(strHostName)
        lnglpHost = gethostbyname(strHostName)
    End If
    If lnglpHost = 0 Then
        GetIPAddress = ""
        Debug.Print "Windows Sockets are not responding. Unable to successfully get Host Name."
        SocketsCleanup
        Exit Function
    End If
    CopyMemory HOST, lnglpHost, Len(HOST)
    CopyMemory lngdwIPAddr, HOST.hAddrList, 4
    ReDim BytTmpIPAddr(1 To HOST.hLen)
    CopyMemory BytTmpIPAddr(1), lngdwIPAddr, HOST.hLen
    For intLoop = 1 To HOST.hLen
       strIPAddr = strIPAddr & BytTmpIPAddr(intLoop) & "."
    Next
    GetIPAddress = Mid$(strIPAddr, 1, Len(strIPAddr) - 1)
    SocketsCleanup
  Exit Function
ChkErr:
    ErrHandle "clsGetIP", "Function : GetIPAddress"
End Function

Private Function HiByte(ByVal wParam As Integer)
    HiByte = wParam \ &H100 And &HFF&
End Function

Private Function LoByte(ByVal wParam As Integer)
    LoByte = wParam And &HFF&
End Function

Private Sub SocketsCleanup()
  On Error GoTo ChkErr
    If WSACleanup() <> ERROR_SUCCESS Then Debug.Print "Socket error occurred in Cleanup."
  Exit Sub
ChkErr:
    ErrHandle "clsGetIP", "Subroutine : SocketsCleanup"
End Sub

Private Function SocketsInitialize() As Boolean
  On Error GoTo ChkErr
   Dim WSAD As WSADATA
   Dim sLoByte As String
   Dim sHiByte As String
   If WSAStartup(WS_VERSION_REQD, WSAD) <> ERROR_SUCCESS Then
      Debug.Print "The 32-bit Windows Socket is not responding."
      SocketsInitialize = False
      Exit Function
   End If
   If WSAD.wMaxSockets < MIN_SOCKETS_REQD Then
        Debug.Print "This application requires a minimum of " & CStr(MIN_SOCKETS_REQD) & " supported sockets."
        SocketsInitialize = False
        Exit Function
   End If
   If LoByte(WSAD.wVersion) < WS_VERSION_MAJOR Or (LoByte(WSAD.wVersion) = WS_VERSION_MAJOR And HiByte(WSAD.wVersion) < WS_VERSION_MINOR) Then
      sHiByte = CStr(HiByte(WSAD.wVersion))
      sLoByte = CStr(LoByte(WSAD.wVersion))
      Debug.Print "Sockets version " & sLoByte & "." & sHiByte & " is not supported by 32-bit Windows Sockets."
      SocketsInitialize = False
      Exit Function
   End If
    SocketsInitialize = True
  Exit Function
ChkErr:
    ErrHandle "clsGetIP", "Function : SocketsInitialize"
End Function

Private Sub Class_Initialize()
  On Error GoTo ChkErr
    blnSocketOK = SocketsInitialize()
  Exit Sub
ChkErr:
    ErrHandle "clsGetIP", "Subroutine : Class_Initialize"
End Sub
