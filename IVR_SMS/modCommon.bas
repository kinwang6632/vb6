Attribute VB_Name = "modCommon"
Option Explicit
Public Sub ErrHandle(Optional FormName As String = "frmMain", Optional ProcedureName As String = "") ' ���~�x���B�z
    Dim strErr As String
    Dim ErrNum As Variant
    Dim ErrDesc As Variant
    Dim ErrSrc As Variant
    ErrNum = Err.Number:      ErrDesc = Err.Description:     ErrSrc = Err.Source
    
    On Error Resume Next
    Screen.MousePointer = vbDefault
    DoEvents
    Screen.ActiveForm.Enabled = True
    strErr = "�o�Ϳ��~�ɶ� : " & CStr(Now) & vbCrLf & _
                "�o�Ϳ��~�M�� : " & CStr(ErrSrc) & vbCrLf & _
                "�o�Ϳ��~��� : " & FormName & vbCrLf & _
                "�o�Ϳ��~�{�� : " & ProcedureName & vbCrLf & _
                "���~�N�X : " & CStr(ErrNum) & vbCrLf & _
                "���~��] : " & CStr(ErrDesc)
    strRetMsg = strErr
    blnRetOK = False
    'MsgBox strErr, vbOKOnly + vbInformation, "�t�ΰ�����~!"
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
Public Function GetXMLobj(ByRef objXML As Object, Optional strErrMsg As String = "") As Boolean ' �إ� XML http ����
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
                                strErrMsg = "�L�k�إ� M$ XML HTTP ���� !"
                                strErrorMessage = strErrMsg
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
End Function
