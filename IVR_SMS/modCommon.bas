Attribute VB_Name = "modCommon"
Option Explicit
Public Sub ErrHandle(Optional FormName As String = "frmMain", Optional ProcedureName As String = "") ' 錯誤掌控處理
    Dim strErr As String
    Dim ErrNum As Variant
    Dim ErrDesc As Variant
    Dim ErrSrc As Variant
    ErrNum = Err.Number:      ErrDesc = Err.Description:     ErrSrc = Err.Source
    
    On Error Resume Next
    Screen.MousePointer = vbDefault
    DoEvents
    Screen.ActiveForm.Enabled = True
    strErr = "發生錯誤時間 : " & CStr(Now) & vbCrLf & _
                "發生錯誤專案 : " & CStr(ErrSrc) & vbCrLf & _
                "發生錯誤表單 : " & FormName & vbCrLf & _
                "發生錯誤程序 : " & ProcedureName & vbCrLf & _
                "錯誤代碼 : " & CStr(ErrNum) & vbCrLf & _
                "錯誤原因 : " & CStr(ErrDesc)
    strRetMsg = strErr
    blnRetOK = False
    'MsgBox strErr, vbOKOnly + vbInformation, "系統執行錯誤!"
End Sub
Public Function MidMbcs(ByVal str As String, start, length) As String
  On Error GoTo ChkErr
    MidMbcs = StrConv(MidB(StrConv(str, vbFromUnicode), start, length), vbUnicode)
  Exit Function
ChkErr:
    ErrHandle "modCommon", "MidMbcs"
End Function
Public Function MsgToWide(ByVal sMsg As String) As String
  On Error GoTo ChkErr
    Dim strSpecial As String
    Dim i As Integer
    Dim strReturn As String
    strSpecial = "!@#$%^&*().,\~][" & "';:/?><`"
    strReturn = sMsg
    For i = 0 To Len(strSpecial) - 1
        strReturn = Replace(strReturn, Mid(strSpecial, i + 1, 1), StrConv(Mid(strSpecial, i + 1, 1), vbWide))
    Next i
    MsgToWide = strReturn
  Exit Function
ChkErr:
    ErrHandle "modCommon", "MsgToWide"
End Function
Public Function GetXMLobj(ByRef objXML As Object, Optional strErrMsg As String = "") As Boolean ' 建立 XML http 物件
  On Error Resume Next
    Dim strErrorMessage  As String
    GetXMLobj = True
    Set objXML = CreateObject("Msxml2.XMLHTTP.5.0")
    If Err.Number <> 0 Then
        Err.Clear
        Set objXML = CreateObject("Msxml2.XMLHTTP.4.0")
        If Err.Number <> 0 Then
            Err.Clear
            Set objXML = CreateObject("Msxml2.XMLHTTP.3.0")
            If Err.Number <> 0 Then
                Err.Clear
                Set objXML = CreateObject("Msxml2.XMLHTTP.2.6")
                If Err.Number <> 0 Then
                    Err.Clear
                    Set objXML = CreateObject("MSXML2.ServerXMLHTTP")
                    If Err.Number <> 0 Then
                        Err.Clear
                        Set objXML = CreateObject("MSXML2.XMLHTTP")
                        If Err.Number <> 0 Then
                            Err.Clear
                            Set objXML = CreateObject("Microsoft.XMLHTTP")
                            If Err.Number <> 0 Then
                                Err.Clear
                                GetXMLobj = False
                                strErrMsg = "無法建立 M$ XML HTTP 物件 !"
                                strErrorMessage = strErrMsg
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
End Function
