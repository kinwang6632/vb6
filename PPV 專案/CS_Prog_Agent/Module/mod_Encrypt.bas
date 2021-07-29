Attribute VB_Name = "mod_Encrypt"
' PR : Hammer
' Start date : 2004/12/22
' Last Modify : 2004/12/22

Option Explicit

' 將字串轉成位元陣列後再轉成16進位的特定字串格式
Private Function Convert_2_Define_String(ByVal strData As String) As String
  On Error GoTo ChkErr
    Dim bytAry() As Byte
    Dim strResult() As String
    Dim lngUB As Long
    Dim lngLoop As Long
    Convert_2_Define_String = ""
    If Len(strData) = 0 Then Exit Function
    bytAry = StrConv(strData, vbFromUnicode)
    lngUB = UBound(bytAry)
    ReDim strResult(lngUB) As String
    For lngLoop = 0 To lngUB
        strResult(lngLoop) = Hex(bytAry(lngLoop))
    Next
    Convert_2_Define_String = Join(strResult, Chr(255))
    Erase bytAry
    Erase strResult
  Exit Function
ChkErr:
    ErrHandle "mod_Encrypt", "Function : Convert_2_Define_String"
End Function

' 還原 Convert_2_Define_String 函數所產生的值
Private Function Revert_2_Original_String(ByVal strData As String) As String
  On Error GoTo ChkErr
    Dim varAry As Variant
    Dim varElement As Variant
    Dim bytAry() As Byte
    Dim lngLoop As Long
    Revert_2_Original_String = ""
    If Len(strData) = 0 Then Exit Function
    varAry = Split(strData, Chr(255))
    ReDim bytAry(UBound(varAry)) As Byte
    For Each varElement In varAry
        bytAry(lngLoop) = Val("&H" & varElement)
        lngLoop = lngLoop + 1
    Next
    Revert_2_Original_String = StrConv(bytAry, vbUnicode)
    Erase bytAry
  Exit Function
ChkErr:
    ErrHandle "mod_Encrypt", "Function : Revert_2_Original_String"
End Function

Public Function Encrypt(ByVal strData As String, Optional EncryptKey As String = "PowerHammer") As String ' 加密
  On Error GoTo ChkErr
    Dim strChar As String
    Dim strKey As String
    Dim strTmp As String
    Dim intPos As Integer
    Dim intLoop As Integer
    Dim strPart1 As String
    Dim strPart2 As String
    intPos = 1
    For intLoop = 1 To Len(strData)
        strChar = Mid(strData, intLoop, 1)
        strKey = Mid(EncryptKey, intPos, 1)
        strTmp = strTmp & Chr(Asc(strChar) Xor Asc(strKey))
        If intPos = Len(EncryptKey) Then intPos = 0
        intPos = intPos + 1
    Next intLoop
'    Debug.Print strTmp
    If Len(strTmp) Mod 2 = 0 Then
        strPart1 = StrReverse(Left(strTmp, (Len(strTmp) / 2)))
        strPart2 = StrReverse(Right(strTmp, (Len(strTmp) / 2)))
        strTmp = strPart1 & strPart2
    End If
'    Encrypt = strTmp
    Encrypt = Convert_2_Define_String(strTmp)
  Exit Function
ChkErr:
    ErrHandle "mod_Encrypt", "Function : Encrypt"
End Function

Public Function Decrypt(ByVal strData As String, Optional EncryptKey As String = "PowerHammer") As String ' 解密
  On Error GoTo ChkErr
    Dim strChar As String
    Dim strKey As String
    Dim strTmp As String
    Dim intPos As Integer
    Dim intLoop As Integer
    Dim strPart1 As String
    Dim strPart2 As String
    intPos = 1
    strData = Revert_2_Original_String(strData)
    If Len(strData) Mod 2 = 0 Then
        strPart1 = StrReverse(Left(strData, (Len(strData) / 2)))
        strPart2 = StrReverse(Right(strData, (Len(strData) / 2)))
        strData = strPart1 & strPart2
    End If
    For intLoop = 1 To Len(strData)
        strChar = Mid(strData, intLoop, 1)
        strKey = Mid(EncryptKey, intPos, 1)
        strTmp = strTmp & Chr(Asc(strChar) Xor Asc(strKey))
        If intPos = Len(EncryptKey) Then intPos = 0
        intPos = intPos + 1
    Next
    Decrypt = strTmp
  Exit Function
ChkErr:
    ErrHandle "mod_Encrypt", "Function : Decrypt"
End Function

