Attribute VB_Name = "mod_CharsSetCVT"
Option Explicit

Public Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByRef lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As String, ByVal lpUsedDefaultChar As Long) As Long
Public Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByRef lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long

Public sCodePage As Long
Public cnvUni2 As String
Public cnvUni As String

Public Function EncodeUTF8(ByVal cnvUni As String)
    If cnvUni = vbNullString Then Exit Function
    EncodeUTF8 = StrConv(WToA(cnvUni, sCodePage, 0), vbUnicode)
End Function

Public Function DecodeUTF8(ByVal cnvUni As String)
    If cnvUni = vbNullString Then Exit Function
    cnvUni2 = WToA(cnvUni, CP_ACP)
    DecodeUTF8 = AToW(cnvUni2, sCodePage)
End Function

Public Function CharSetConvert(ByRef strData As String, _
                                                        docCharset As String, _
                                                        Optional ReadFileName As String = Empty, _
                                                        Optional SaveFileName As String = Empty) As String
    Dim objStream As Object
    Set objStream = CreateObject("ADODB.Stream")
    If strData <> Empty Or ReadFileName <> Empty Then
        With objStream
                .Type = 2
                .Mode = 3
'                .Options = 1
                .Open
                 If ReadFileName <> Empty Then .LoadFromFile ReadFileName
                 If strData <> Empty Then .WriteText strData
                .Position = 0
                .Charset = docCharset
'                .SetEOS
                 CharSetConvert = .ReadText
                 If SaveFileName <> Empty Then .SaveToFile SaveFileName, 2
                .Close
        End With
    End If
End Function

'Purpose:Convert Utf8 to Unicode
Public Function UTF8_Decode(ByVal sUTF8 As String) As String

   Dim lngUtf8Size      As Long
   Dim strBuffer        As String
   Dim lngBufferSize    As Long
   Dim lngResult        As Long
   Dim bytUtf8()        As Byte
   Dim n                As Long

   If LenB(sUTF8) Then
      On Error GoTo EndFunction
      bytUtf8 = StrConv(sUTF8, vbFromUnicode)
      lngUtf8Size = UBound(bytUtf8) + 1
      On Error GoTo 0
      'Set buffer for longest possible string i.e. each byte is
      'ANSI<=&HFF, thus 1 unicode(2 bytes)for every utf-8 character.
      lngBufferSize = lngUtf8Size * 2
      strBuffer = String$(lngBufferSize, vbNullChar)
      'Translate using code page 65001(UTF-8)
      lngResult = MultiByteToWideChar(CP_UTF8, 0, bytUtf8(0), _
         lngUtf8Size, StrPtr(strBuffer), lngBufferSize)
      'Trim result to actual length
      If lngResult Then
         UTF8_Decode = Left$(strBuffer, lngResult)
         'Debug.Print UTF8_Decode
      End If
   End If

EndFunction:

End Function

Public Function Utf8Decode(ByRef bytUtf8() As Byte, Optional ByRef strDefaultChar As String) As String

Dim lngUtf8Size As Long
Dim lngDefaultCharLength As Long
Dim strBuffer As String
Dim lngWriteLength As Long
Dim bytUcs2Char(1) As Byte
Dim i As Long

    On Error GoTo ExitFunction
    lngUtf8Size = UBound(bytUtf8)
    On Error GoTo 0
    lngDefaultCharLength = Len(strDefaultChar)
    If lngDefaultCharLength Then
        strBuffer = String$((lngUtf8Size + 1) * lngDefaultCharLength, vbNullChar)
    Else
        strBuffer = String$(lngUtf8Size + 1, vbNullChar)
    End If
    lngWriteLength = 1
    
    Do While i <= lngUtf8Size
        If bytUtf8(i) <= &H7F Then
            Mid(strBuffer, lngWriteLength, 1) = ChrW$(bytUtf8(i))
            lngWriteLength = lngWriteLength + 1
            i = i + 1
        ElseIf (bytUtf8(i) >= &HC2 And bytUtf8(i) <= &HDF) Then
            If (i + 1) <= lngUtf8Size Then
                If (bytUtf8(i + 1) >= &H80 And bytUtf8(i + 1) <= &HBF) Then
                    bytUcs2Char(0) = ((bytUtf8(i) And &H3) * &H40) Or (bytUtf8(i + 1) And &H3F)
                    bytUcs2Char(1) = (bytUtf8(i) And &H1C) \ &H4
                    Mid(strBuffer, lngWriteLength, 1) = bytUcs2Char
                    lngWriteLength = lngWriteLength + 1
                    i = i + 2
                Else
                    If lngDefaultCharLength Then GoSub SetDefaultChar
                    i = i + 1
                End If
            Else
                If lngDefaultCharLength Then GoSub SetDefaultChar
                i = i + 1
            End If
        ElseIf (bytUtf8(i) >= &HE0 And bytUtf8(i) <= &HEF) Then
            If (i + 2) <= lngUtf8Size Then
                If (bytUtf8(i + 1) >= &H80 And bytUtf8(i + 1) <= &HBF) And _
                   (bytUtf8(i + 2) >= &H80 And bytUtf8(i + 2) <= &HBF) Then
                    bytUcs2Char(0) = ((bytUtf8(i + 1) And &H3) * &H40) Or (bytUtf8(i + 2) And &H3F)
                    bytUcs2Char(1) = ((bytUtf8(i) And &HF) * &H10) Or ((bytUtf8(i + 1) And &H3C) \ &H4)
                    Mid(strBuffer, lngWriteLength, 1) = bytUcs2Char
                    lngWriteLength = lngWriteLength + 1
                    i = i + 3
                Else
                    If lngDefaultCharLength Then GoSub SetDefaultChar
                    i = i + 1
                End If
            Else
                If lngDefaultCharLength Then GoSub SetDefaultChar
                i = i + 1
            End If
        Else
            If lngDefaultCharLength Then GoSub SetDefaultChar
            i = i + 1
        End If
    Loop

    Utf8Decode = Left$(strBuffer, lngWriteLength - 1)

    Exit Function
    
SetDefaultChar:
    Mid(strBuffer, lngWriteLength, lngDefaultCharLength) = strDefaultChar
    lngWriteLength = lngWriteLength + lngDefaultCharLength
    Return

ExitFunction:

End Function
