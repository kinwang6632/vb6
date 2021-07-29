Attribute VB_Name = "mod_Cvt_Code"
Option Explicit

Private Declare Function GetACP Lib "Kernel32" () As Long
Private Declare Function MultiByteToWideChar Lib "Kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Private Declare Function WideCharToMultiByte Lib "Kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, lpUsedDefaultChar As Long) As Long

Public Function Encode_2_UTF8(ByVal strValue As String) As String
  On Error GoTo ChkErr
    
    Encode_2_UTF8 = Empty
    If strValue = vbNullString Then Exit Function
    Encode_2_UTF8 = StrConv(Cvt_Cpg(strValue, 65001), vbUnicode)
  
  Exit Function
ChkErr:
    ErrHandle "mod_Cvt_Code", "Encode_2_UTF8"
End Function

Private Function Cvt_Cpg(ByVal strValue As String, Optional ByVal CodePage As Long = -1) As String
  On Error GoTo ChkErr
    
    Dim strBuffer As String
    Dim lngWC As Long
    Dim lngStrPtr As Long
    Dim lngStrBufferPtr As Long
    
    If CodePage = -1 Then CodePage = GetACP
    lngStrPtr = StrPtr(strValue)
    lngWC = WideCharToMultiByte(CodePage, 0, lngStrPtr, -1, 0&, 0&, ByVal 0&, ByVal 0&)
    strBuffer = String(lngWC + 1, vbNullChar)
    lngStrBufferPtr = StrPtr(strBuffer)
    lngWC = WideCharToMultiByte(CodePage, 0, lngStrPtr, -1, lngStrBufferPtr, Len(strBuffer), ByVal 0&, ByVal 0&)
    Cvt_Cpg = Left(strBuffer, lngWC - 1)
    
  Exit Function
ChkErr:
    ErrHandle "mod_Cvt_Code", "Cvt_Cpg"
End Function
