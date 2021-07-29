VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCNS_SMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const clsName As String = "clsCNS_SMS"

Private Sub SendSMS()
  On Error GoTo ChkErr
    Dim strPhoneNo As String
    Dim strUrl As String
    If GetXMLobj(objxmlHttp) Then
        If Get_SMS_Info(strSvr, strUser, strPwd) Then
            'strPhoneNo
            'strUrl = strSMSUrl & "?username=" & strUser & "&password=" & strPwd & "&message=" & Msg
            strUrl = strSMSUrl & "?id=" & strUrl & "&passwd=" & strPwd
                   
            For Each p In Split(strPhone, ",")
                If p <> "" Then
                    intIdx = intIdx + 1
                    strUrl = strUrl & "&target" & intIdx & "=" & p
                End If
            Next
        
            lblAction = funAction(3, "CircuitNo : " & CircuitNo & ",")
            DoEvents
            
            RetMsg = Inet.OpenURL(strUrl) ' Send SMS
    Else
    
    End If
  Exit Sub
ChkErr:
    ErrHandle clsName, "SendSMS"
End Sub
Public Function sBytesToBstr(ByVal vIn As String)
  Dim objStream
  Set objStream = Server.CreateObject("adodb.stream")
  objStream.Type = 1
  objStream.Mode = 3
  objStream.open
  objStream.Write vIn
    
  objStream.Position = 0
  objStream.Type = 2
  objStream.Charset = "big5"
  sBytesToBstr = objStream.ReadText
  objStream.Close
  Set objStream = Nothing
End Function


Function SaveTextAsUTF8(filePath, Text)
        Const adTypeText = 2
        Const adSaveCreateOverWrite = 2
        Dim x As New ADODB.Stream
        
        'Create Stream object
        Dim TextStream
        Set TextStream = CreateObject("ADODB.Stream")
        With TextStream
                .open
                .Charset = "Big5"
                .Position = TextStream.Size
                .WriteText Text
                .SaveToFile filePath, adSaveCreateOverWrite
                .Close
        End With
        Set TextStream = Nothing
        
End Function


