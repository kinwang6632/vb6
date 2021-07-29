VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EnableProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Function EnableProcess(ByVal IniFile As String, ByVal Section As String, _
                            ByVal Key As String, Optional ByVal blnAny As Boolean = True) As Boolean
  On Error GoTo ChkErr
    Dim strAry As Variant
    Dim strMac As String
    Dim aValue As String
    Dim aryMac As Variant
    Dim strTmp As String
    Dim blnRet As Boolean
    Dim i As Integer
    If IniFile = "" Then Exit Function
    If Not ChkDirExist(IniFile) Then Exit Function
    Dim s As String, length As Long
    s = String(1024, 0)
    length = GetPrivateProfileString(Section, Key, "", s, Len(s), IniFile)
    aValue = Left(s, length)
    aValue = CreateObject("DevPowerEncrypt.EnCrypt").Decrypt(aValue)
    strAry = Split(aValue, ";")
    If CDate(Format(Now, "yyyy/mm/dd")) > CDate(strAry(0)) Then Exit Function
    For i = 1 To UBound(strAry)
        strMac = IIf(strMac = "", strAry(i), strMac & ";" & strAry(i))
    Next i
    If Not blnAny Then
        If Replace(strMac, ":", "-") <> GetMacAddress Then
            blnRet = False
        Else
            blnRet = True
        End If
    Else
        aryMac = Split(strMac, ";")
        strTmp = GetMacAddress
        strTmp = Replace(strTmp, ":", "-")
        For i = LBound(aryMac) To UBound(aryMac)
            If InStr(1, strTmp, aryMac(i)) > 0 Then
                blnRet = True
                Exit For
            End If
        Next i
    End If
    EnableProcess = blnRet
    Exit Function
ChkErr:
  EnableProcess = False
End Function

Private Function ChkDirExist(FileName As String) As Boolean
  On Error GoTo ChkErr
    Dim varValue As Variant
    If Right(FileName, 1) = "\" Then FileName = Left(FileName, Len(FileName) - 1)
    FileName = UCase(FileName)
    varValue = Split(FileName, "\")
    If UCase(Dir(FileName, vbDirectory)) = varValue(UBound(varValue)) Then
        ChkDirExist = True
    Else
        If Dir(Trim(FileName)) = "" Then
            ChkDirExist = False
        Else
            ChkDirExist = True
        End If
    End If
  Exit Function
ChkErr:
    ChkDirExist = False
End Function
Private Function GetMacAddress() As String
  On Error GoTo ChkErr
    Dim objWMI As Object
    Dim objAdapters As Object
    Dim objMac As Object
    Dim strMac As String
    Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
    Set objAdapters = objWMI.ExecQuery _
        ("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")
    For Each objMac In objAdapters
        strMac = IIf(strMac = "", objMac.MacAddress, strMac & ";" & objMac.MacAddress)
    Next
    GetMacAddress = strMac
    Exit Function
ChkErr:
  GetMacAddress = ""
End Function
Public Function Encryption(ByVal aValue As String, Optional ByVal strMac As String = "") As String
  On Error GoTo ChkErr
    Dim strEncoding
    If strMac = "" Then
        strEncoding = aValue & ";" & GetMacAddress
    Else
        strEncoding = aValue & ";" & strMac
    End If
    strEncoding = CreateObject("DevPowerEncrypt.EnCrypt").Encrypt(strEncoding)
    Encryption = strEncoding
    Exit Function
ChkErr:
  Encryption = ""
End Function

