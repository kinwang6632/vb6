Attribute VB_Name = "mduControlPanelWeb"
Option Explicit

Private Const blnDebug = False

Public gCompFilterStr As String


Public gDefaultComp As String
Public objActForm As Form
Public gLogInID As String
Public gstrUserPremission As String
Public lngSeqNo As Long
Public strOwner As String
'Dim lngCMtimeout As Long
Public strCmdSeqNo As String
Public blnExeCommand As Boolean
Public strCMowner As String
Dim strErrorMessage As String
Public intCompCode As Integer
Public blnShowMsg As Boolean
Public strFaciSNO As String
Public strFaciCode As String
Public strResvTime As String
Public blnSendKey As Boolean
Public strReason As String
Public blnShowFaci As Boolean
Public str7SNO As String
Public str3MediaBillNo As String
Public strSeqNo As String
Public strReasonCode As String
Public strReasonName As String
Public strRetCmdSeqNo As String
Public lngCDRTimeOut As Long
Public blnRecordProcedure As Boolean
Public rs2 As New ADODB.Recordset
Public Declare Function GetTickCount Lib "kernel32" () As Long

Public Function GetModeTypeValue(ByVal strValue As String) As String
    On Error Resume Next
    Dim strReturn As String
        
        Select Case strValue
            Case "A1"
                strReturn = "�}��"
            Case "A2"
                strReturn = "����"
            Case "A3"
                strReturn = "����"
            Case "A4"
                strReturn = "���](Reset)"
            
            Case "B1"
                strReturn = "��CM"
            Case "B2"
                strReturn = "�ܧ�CM�t�v"
            Case "B3"
                strReturn = "Ping CM"
            Case "B4"
                strReturn = "����ʺAIP"
            Case "B5"
                strReturn = "ACS�b���K�X�����\��"
            
            Case "C1"
                strReturn = "CM ���A�d��"
            Case "C2"
                strReturn = "�d�� HFC �A�����O"
            Case "C3"
                strReturn = "�M��EMM"
            Case "C4"
                strReturn = "CM �s�u�����d��"
            Case "C5"
                strReturn = "PCIP�s�u�O���d��"
        End Select
        
        GetModeTypeValue = strReturn
End Function
Public Function GetChannelCitemCode(ByVal aChannelID As String, ByVal aCitemCode As String) As String
  On Error GoTo chkErr
    Dim rsCD019A As New ADODB.Recordset
    Dim aSQL As String
    If aCitemCode = "" Then
        GetChannelCitemCode = ""
        Exit Function
    End If
    aSQL = "SELECT A.CITEMCODE FROM " & GetOwner & "CD019A A," & GetOwner & "CD024 B " & _
            " WHERE B.CHANNELID ='" & aChannelID & "' AND A.CODENO = B.CODENO " & _
            " AND B.VODTYPE=2 AND B.STOPFLAG <> 1 " & _
            " AND A.CITEMCODE IN(" & Replace(aCitemCode, "'", "") & ")"
    If Not GetRS(rsCD019A, aSQL, gcnGi, adUseClient, adOpenKeyset) Then Exit Function
    If rsCD019A.EOF Then
        GetChannelCitemCode = ""
    Else
        GetChannelCitemCode = rsCD019A("CITEMCODE") & ""
    End If
    On Error Resume Next
    Call CloseRecordset(rsCD019A)
    Exit Function
chkErr:
  Call ErrSub(FormName, "GetChannelCitemCode")
End Function

'#5692 ���oSVOD�W�D(aStatusType=0 �q��SVOD) By Kin 2010/06/15
Public Function GetChannelID(ByVal aCitemCode As String, ByVal aCompCode As String, _
                            Optional ByVal aCustId As String = Empty, _
                            Optional ByVal aStatusType As Integer = 0, _
                            Optional ByVal aVodaccountId As String) As String
  On Error GoTo chkErr
    Dim rsCD024 As New ADODB.Recordset
    Dim aSQL As String
    Dim aRet As String
    aSQL = Empty
    aRet = Empty
    If aCitemCode = "" Then
        GetChannelID = ""
        Exit Function
    End If
    aCitemCode = Replace(aCitemCode, "'", "")
    aSQL = "SELECT DISTINCT CD024.ChannelID FROM " & GetOwner & "CD024 " & _
        " WHERE CD024.CodeNo IN( SELECT CODENO FROM " & GetOwner & "CD019A " & _
        " WHERE CD019A.CITEMCODE IN(" & aCitemCode & ")) " & _
        " AND CD024.VODType=2 AND CD024.STOPFLAG <> 1 AND CD024.COMPCODE=" & aCompCode
        
    If aStatusType = 0 Then '0�N��n�}�qSVOD 1�N��n��SVOD
        aSQL = aSQL & " AND CD024.CODENO NOT IN(SELECT ProductID FROM " & GetOwner & "SO555B " & _
                " WHERE NVL(STATUS,0)=1 AND VODACCOUNTID = " & aVodaccountId & ") "
    Else
        aSQL = aSQL & " AND CD024.CODENO NOT IN(SELECT ProductID FROM " & GetOwner & "SO555B " & _
                " WHERE NVL(STATUS,0)=0 AND VODACCOUNTID = " & aVodaccountId & ") "
    End If
    If Not GetRS(rsCD024, aSQL, gcnGi, adUseClient, adOpenKeyset) Then Exit Function
    Do While Not rsCD024.EOF
        If aRet = "" Then
            aRet = rsCD024("ChannelID")
        Else
            aRet = aRet & "," & rsCD024("ChannelID")
        End If
        rsCD024.MoveNext
    Loop
    GetChannelID = aRet
    On Error Resume Next
    Call CloseRecordset(rsCD024)
    Exit Function
chkErr:
  Call ErrSub(FormName, "GetChannelID")
End Function
Public Function Get_EMTA_Mac(strFaciSNO As String) As String
  On Error GoTo chkErr
    If strFaciSNO <> Empty Then
        Get_EMTA_Mac = GetRsValue("SELECT MTAMAC FROM " & GetOwner & "SO306" & _
                                                        " WHERE HFCMAC='" & strFaciSNO & "'") & ""
       
    End If
  Exit Function
  
chkErr:
    ErrSub FormName, "Get_EMTA_Mac"
End Function
Public Function Get_EMTA_Mode(strFaciSNO As String) As String
  On Error Resume Next
    Get_EMTA_Mode = GetRsValue("SELECT Nvl(ModeID,0) FROM " & GetOwner & "SO306" & _
                                                        " WHERE HFCMAC='" & strFaciSNO & "'") & ""
End Function
Public Function GetMacAddress(ByVal aFaciSNo As String) As String
  On Error GoTo chkErr
    Dim rsMac As New ADODB.Recordset
    Dim strMac As String
    Dim strTmp As String
    Dim i As Integer
    If aFaciSNo = "" Then
        strMac = ""
    Else
        If Not GetRS(rsMac, "Select EthernetMAC From " & GetOwner & "SOAC0201A Where " & _
                        "FaciSNo='" & aFaciSNo & "'", gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
        If rsMac.EOF Then
            strMac = ""
        Else
            If rsMac(0) & "" = "" Then
                strMac = ""
            Else
                strMac = rsMac(0).Value
                If InStr(1, strMac, ":") = 0 Then
                    For i = 2 To Len(strMac) Step 2
                        strTmp = strTmp & Mid(strMac, i - 1, 2) & ":"
                    Next
                    strMac = Mid(strTmp, 1, Len(strTmp) - 1)
                End If
            End If
        End If
    End If
    GetMacAddress = strMac
  On Error Resume Next
  Call CloseRecordset(rsMac)
  Exit Function
chkErr:
  ErrSub FormName, "GetMacAddress"
End Function

Public Function SetVodCmd(ByRef rsData As ADODB.Recordset, ByVal strCMD As String, _
                        Optional ByVal strExchTime As String = "", Optional ByVal strServiceProviderID As String = "", _
                        Optional ByVal strAccountNum As String, Optional ByVal strAccoutName As String, _
                        Optional ByVal strAddress As String = "", Optional ByVal strPostCode As String = "", _
                        Optional ByVal strLoginID As String = "", Optional ByVal strPin As String = "", _
                        Optional ByVal strSmartCardID As String = "", _
                        Optional ByVal strMacAddress As String = "", Optional ByVal strSerialNumber As String = "", _
                        Optional ByVal strCASNo As String = "", Optional ByVal strAccountUID As String, _
                        Optional ByVal strProductID As String = "", Optional ByVal strProductName As String, _
                        Optional ByVal strUpdTime As String = "", Optional ByVal strUpdEn As String = "", _
                        Optional ByRef strSTBUId As String = "", _
                        Optional ByVal strProcessingDate As String = "", Optional ByVal strSNO As String, _
                        Optional ByVal strMediaBillNo As String = "", Optional ByVal strReasonCode As String = "", _
                        Optional ByVal strReasonName As String = "", Optional ByVal strBillStartDate As String = "", _
                        Optional ByVal strBillEndDate As String = "", _
                        Optional ByRef rsResult As Object = Nothing, Optional strResult As String, _
                        Optional ByRef strRetErrMsg As String, Optional ByRef strFaultReason As String, _
                        Optional ByVal strICCUID As String, Optional ByVal strCitemStr As String = "") As Boolean
  On Error GoTo chkErr
    Dim strNowDate As String
    Dim varBookMark As Variant
    Dim strCmdName As String
    Dim strUpdDate As String
    Dim strUpdName As String
    Dim objVodCommand As New clsVodcommand
    'Dim rs555B As New ADODB.Recordset
    Dim rs555 As New ADODB.Recordset
    'Dim rs005B As New ADODB.Recordset
    Dim i As Integer
    Dim aCitemCode As String
    Dim aServiceType As String
    Dim strUPDSOAC0201B As String
    If strExchTime = "" Then
        strUpdDate = RightNow
    Else
        strUpdDate = strExchTime
    End If
    
    If strUpdTime = "" Then
        strUpdTime = strUpdDate
    End If
    If strUpdEn = "" Then
        strUpdName = GaryGi(1)
    Else
        strUpdName = strUpdEn
    End If
    intCompCode = rsData("CompCode")
    Call WriteRecordVodProcedure("2.2 �i�J�s�W����x��� --> " & strCMD)
    varBookMark = rsData.AbsolutePosition
    
    Select Case UCase(strCMD)
        Case "E1"
            strCmdName = "���]�q�ʱK�X"
        Case "C4"
            strCmdName = "�]�w/�ܧ�ϰ�"
        Case "B11"
            strCmdName = "�q��SVOD"
        Case "B12"
            strCmdName = "����SVOD"
        Case "Z2"
            strCmdName = "��QSP AccountUID�b�����"
        Case "Z3"
            strCmdName = "��QSP���q�O����d��b��"
        Case "Z4"
            strCmdName = "�إ�STB�]�Ʃ�QSP��"
        Case "Z5"
            strCmdName = "��QSP������STB�]��"
    End Select
    If blnExeCommand Then
        Call WriteRecordVodProcedure("2.2 �i�J�s�W����x��� --> " & strCMD)
        If Not objVodCommand.Vod_Command(gcnGi.ConnectionString, intCompCode & "", GetOwner, strCMD _
                                , , , , , strUpdDate, , , strServiceProviderID, strAccountNum, _
                                strAccoutName, , strAddress, strPostCode, _
                                strAccountUID, strLoginID, _
                                strPin, , , strSmartCardID, strMacAddress, _
                                strSerialNumber, , , , strCASNo _
                                , strProductID, strProductName, strUpdTime, strUpdName, _
                                strSTBUId, strProcessingDate, _
                                strSNO, strMediaBillNo, strReasonCode, strReasonName, GaryGi(0), _
                                str(intCompCode), strBillStartDate, strBillEndDate, _
                                IIf(blnShowMsg, frmSO1623A.pbr1, Nothing), rsResult, strResult, strRetErrMsg, _
                                strFaultReason, strICCUID, strCitemStr) Then GoTo 88
            
        Call WriteRecordVodProcedure("2.3 �����s�W����x��� --> " & strCMD)
    End If
    
    If strProcessingDate <> Empty And blnExeCommand Then
        strResult = "1"
        SetVodCmd = True
        Call WriteRecordVodProcedure("2.4 �w����� --> " & strCMD)
        GoTo 88
    End If
    rsData.AbsolutePosition = varBookMark
    Call WriteRecordVodProcedure("2.4.1 �}�l��s��� --> " & strCMD)
    Select Case UCase(strCMD)
        Case "B11"      '�H��JSO555����Ƭ��D,��JSO555B,�p�GSO555B���s�b�h�s�W,�s�b�hUpd
            Call GetRs555(rs555)
            If rs555 Is Nothing Then strRetErrMsg = "��SO555��ƿ��~" & vbCrLf & "SEQNO=" & strCmdSeqNo: GoTo 88
            If rs555.State <> adStateOpen Then strRetErrMsg = "SO555�}�ҥ��ѡI": GoTo 88
            If rs555.EOF Then strRetErrMsg = "�䤣��SO555 SeqNo=" & strCmdSeqNo & " ����ơI": GoTo 88
            If Not UpdB11Cmd(rs555) Then strRetErrMsg = "��sSO555B���ѡI": GoTo 88
        Case "B12"      '�H��JSO555����Ƭ��D,��sSO555B�����
            Call GetRs555(rs555)
            If rs555 Is Nothing Then strRetErrMsg = "��SO555��ƿ��~" & vbCrLf & "SEQNO=" & strCmdSeqNo: GoTo 88
            If rs555.State <> adStateOpen Then strRetErrMsg = "rs555�}�ҥ��ѡI": GoTo 88
            If rs555.EOF Then strRetErrMsg = "�䤣��SeqNo=" & strCmdSeqNo & " ����ơI": GoTo 88
            If Not UpdB12Cmd(rs555, strSNO) Then strRetErrMsg = "��sSO555B���ѡI": GoTo 88
        Case "Z4"
            If Not rsResult Is Nothing Then
                If rsResult.State = adStateClosed Then
                    strRetErrMsg = "�䤣��GateWay�^�񪺸�ơI"
                    GoTo 88
                Else
                    If rsResult.EOF Then
                        strRetErrMsg = "�䤣��GateWay�^�񪺸�ơI"
                        GoTo 88
                    Else
                        strSTBUId = rsResult("STBUID") & ""
                    End If
                End If
            Else
                strRetErrMsg = "�䤣��GateWay�^�񪺸�ơI"
                GoTo 88
            End If
        
    End Select
    strResult = "1"
    SetVodCmd = True
    Call WriteRecordVodProcedure("2.5 ������s��� --> " & strCMD)
88:
    If blnShowMsg Then
        ShowMsg "(1) " & strCMD & "  " & strCmdName, strResult, strFaultReason, strRetErrMsg
    End If
    On Error Resume Next
    'Call CloseRecordset(rs005)
    Call CloseRecordset(rs555)
    'Call CloseRecordset(rs005B)
    Exit Function
chkErr:
  Call ErrSub(FormName, "SetVod")
End Function
Private Function UpdB12Cmd(ByRef rs555 As ADODB.Recordset, Optional ByVal aPrSNo As String = Empty) As Boolean
  On Error GoTo chkErr
    Dim rs555B As New ADODB.Recordset
    UpdB12Cmd = False
    rs555.MoveFirst
    Do While Not rs555.EOF
        Call GetRs555B(rs555B, rs555("PID"), rs555("AccountNum"), rs555("AccountUID"), False)
        If rs555B.EOF Then Exit Function
        With rs555B
            .Fields("Status") = 0
            '.Fields("PRSNo") = aPrSNo
            .Fields("CloseTime") = rs555("ExchTime")
            .Update
        End With
        rs555.MoveNext
    Loop
    UpdB12Cmd = True
    On Error Resume Next
    Call CloseRecordset(rs555B)
    Exit Function
chkErr:
  Call ErrSub(FormName, "UpdB12Cmd")
End Function
Private Function UpdB11Cmd(ByRef rs555 As ADODB.Recordset) As Boolean
  On Error GoTo chkErr
    Dim aCitemCode As String
    Dim aServiceType As String
    Dim rs555B As New ADODB.Recordset
    UpdB11Cmd = False
    rs555.MoveFirst
    Do While Not rs555.EOF
        aServiceType = Empty
        aCitemCode = Empty
        Call GetRs555B(rs555B, rs555("PID"), rs555("AccountNum"), rs555("AccountUID"), True)
        With rs555B
            If .EOF Then
                .AddNew
                'Call GetCitemCode(rs555("PID"), aCitemCode, aServiceType)
                .Fields("VODAccountID") = rs555("AccountUID")
                .Fields("ProductID") = rs555("PID")
                .Fields("ProductName") = rs555("PName")
                .Fields("Status") = 1
                .Fields("OpenTime") = rs555("ExchTime")
                '.Fields("CitemCode") = aCitemCode
                .Fields("CitemCode") = NoZero(rs555("CitemCode") & "")
                .Update
            Else
                .Fields("Status") = 1
                .Fields("OpenTime") = rs555("ExchTime")
                .Update
            End If
        End With
        rs555.MoveNext
    Loop
    UpdB11Cmd = True
    On Error Resume Next
    Call CloseRecordset(rs555B)
    Exit Function
chkErr:
  Call ErrSub(FormName, "UpdB11Cmd")
End Function

Public Sub GetCitemCode(ByVal strCode As String, ByRef aCCode As String, ByRef aService As String)
  On Error GoTo chkErr
    Dim strSQL As String
    Dim rsCD019A As New ADODB.Recordset
    strSQL = "Select A.CitemCode,C.ServiceType From " & GetOwner & "CD019A A," & _
            GetOwner & "CD024 B," & GetOwner & "CD019 C " & _
            " Where A.CodeNo='" & strCode & "' And A.CodeNo=B.CodeNO " & _
            " And B.VodType=2 And B.StopFlag<> 1 And A.CitemCode=C.CodeNo"
    If Not GetRS(rsCD019A, strSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
    If Not rsCD019A.EOF Then
        aCCode = rsCD019A("CitemCode")
        aService = rsCD019A("ServiceType")
    End If
    Exit Sub
chkErr:
  Call ErrSub(FormName, "GetCitemCode")
End Sub
Public Sub GetRs555B(ByRef rs As ADODB.Recordset, ByVal strProductID As String, _
                    ByVal aCustId As String, ByVal aVodaccountId As String, _
                    ByVal aOrder As Boolean)
  On Error GoTo chkErr
    Dim strSQL As String
'    strSQL = "Select A.RowID,A.* From " & GetOwner & "SO555B A " & _
'            " Where CustId=" & aCustId & " And ProductID='" & strProductID & "'" & _
'            " AND VODAccountID=" & aVODAccountId & " And Status = " & IIf(aOrder, "1", "0")
            
    strSQL = "Select A.RowID,A.* From " & GetOwner & "SO555B A " & _
            " Where  ProductID='" & strProductID & "'" & _
            " AND VODAccountID=" & aVodaccountId & " And Status = " & IIf(aOrder, "0", "1")
    If Not GetRS(rs, strSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
  Exit Sub
chkErr:
  Call ErrSub(FormName, "GetRs005")
End Sub
Public Sub GetRs555(ByRef rs As ADODB.Recordset)
  On Error GoTo chkErr
    Dim strSQL As String
    strSQL = "Select A.*,B.ProductId PID,B.ProductName PName,C.ChannelID,B.CitemCode From " & _
                strCMowner & "SO555 A," & _
                strCMowner & "SO555A B," & strCMowner & "CD024 C " & _
                " Where A.SeqNo=" & strCmdSeqNo & " And A.SeqNo=B.SeqNo " & _
                " And C.CodeNo=B.ProductId And C.VodType=2 And C.StopFlag<>1"
  
    If Not GetRS(rs, strSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then
        Set rs = Nothing
    End If
    Exit Sub
chkErr:
  Call ErrSub(FormName, "GetRs555")
End Sub

'5696 �p�G�OA6�R�O�n�h��ProductID By Kin 2010/06/17
Public Function GetProductID(ByVal aVodAccount As String) As String
  On Error GoTo chkErr
    Dim aSQL As String
    Dim strCodeNo As String
    Dim rsSO555B As New ADODB.Recordset
    Dim rsCD024 As New ADODB.Recordset
    Dim aRet As String
    aSQL = "SELECT ProductID FROM " & GetOwner & "SO555B " & _
        " WHERE NVL(Status,0)=1 AND VODAccountID = " & aVodAccount
    If Not GetRS(rsSO555B, aSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
    If rsSO555B.EOF Then GetProductID = Empty: Exit Function
    Do While Not rsSO555B.EOF
        If strCodeNo = "" Then
            strCodeNo = "'" & rsSO555B("ProductID") & "'"
        Else
            strCodeNo = strCodeNo & ",'" & rsSO555B("ProductID") & "'"
        End If
        rsSO555B.MoveNext
    Loop
    aSQL = "SELECT ChannelID FROM " & GetOwner & "CD024 " & _
            " WHERE CodeNo IN(" & strCodeNo & ")"
    If Not GetRS(rsCD024, aSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
    Do While Not rsCD024.EOF
        If aRet = "" Then
            aRet = rsCD024("ChannelID")
        Else
            aRet = aRet & "," & rsCD024("ChannelID")
        End If
        rsCD024.MoveNext
    Loop
    GetProductID = aRet
    On Error Resume Next
    Call CloseRecordset(rsSO555B)
    Exit Function
chkErr:
  Call ErrSub(FormName, "GetProductID")
End Function

Public Function OpenVod(ByRef rsData As ADODB.Recordset, ByVal strCMD As String, _
                        Optional ByVal strExchTime As String = "", Optional ByVal strServiceProviderID As String = "", _
                        Optional ByVal strAccountNum As String, Optional ByVal strAccoutName As String, _
                        Optional ByVal strAddress As String = "", Optional ByVal strPostCode As String = "", _
                        Optional ByVal strLoginID As String = "", Optional ByVal strPin As String = "", _
                        Optional ByVal strSmartCardID As String = "", _
                        Optional ByVal strMacAddress As String = "", Optional ByVal strSerialNumber As String = "", _
                        Optional ByVal strCASNo As String = "", Optional ByVal strAccountUID As String, _
                        Optional ByVal strProductID As String = "", Optional ByVal strProductName As String, _
                        Optional ByVal strUpdTime As String = "", Optional ByVal strUpdEn As String = "", _
                        Optional ByRef strSTBUId As String = "", _
                        Optional ByVal strProcessingDate As String = "", Optional ByVal strSNO As String = "", _
                        Optional ByVal strMediaBillNo As String = "", Optional ByVal strReasonCode As String = "", _
                        Optional ByVal strReasonName As String = "", _
                        Optional ByVal strBillStartDate As String = "", Optional ByVal strBillEndDate As String = "", _
                        Optional ByRef rsResult As Object = Nothing, Optional strResult As String, _
                        Optional ByRef strRetErrMsg As String, Optional ByRef strFaultReason As String, _
                        Optional ByVal strICCUID As String = "") As Boolean
  On Error GoTo chkErr
    OpenVod = False
    Dim strUpdDate As String
    Dim strUpdName As String
    Dim varBookMark As Variant
    Dim strCmdName As String
    Dim strUPDSOAC0201B As String
    Dim rsSO004G As New ADODB.Recordset
    Dim objVodCommand As New clsVodcommand
    Dim aOLDICCUID As String
    Dim aNewICCUID As String
    If strExchTime = "" Then
        strUpdDate = RightNow
    Else
        strUpdDate = strExchTime
    End If
    If strUpdTime = "" Then
        strUpdTime = strUpdDate
    End If
    If strUpdEn = "" Then
        strUpdName = GaryGi(1)
    Else
        strUpdName = strUpdEn
    End If
    Call WriteRecordVodProcedure("2.1 �}�l�s�W����x��� --> " & strCMD)
    varBookMark = rsData.AbsolutePosition
    
    Select Case UCase(strCMD)
        Case "A1"
            strCmdName = "�Ұʭq���v"
        Case "A2"
            strCmdName = "�����q���v"
        Case "A3"
            strCmdName = "�Ȱ��q���v"
            
        Case "A4"
            strCmdName = "��_�q���v"
        Case "A6"
            strCmdName = "���״���(��մ�)"
        Case "A7"
            strCmdName = "�t��"
        Case "A8"
            strCmdName = "���״��� STB Swap"
        Case "A9"
            strCmdName = "���״��� Smart Card Swap"
    End Select
    
    
    intCompCode = rsData("CompCode")
    If blnExeCommand Then
        Call WriteRecordVodProcedure("2.2 �i�J�s�W����x��� --> " & strCMD)
        If Not objVodCommand.Vod_Command(gcnGi.ConnectionString, intCompCode & "", GetOwner, strCMD _
                                , , , , , strUpdDate, , , strServiceProviderID, strAccountNum, _
                                strAccoutName, , strAddress, strPostCode, strAccountUID, strLoginID, _
                                strPin, , , strSmartCardID, strMacAddress, strSerialNumber, , , , strCASNo _
                                 , strProductID, strProductName, strUpdTime, strUpdName, _
                                 strSTBUId, strProcessingDate, _
                                 strSNO, strMediaBillNo, strReasonCode, strReasonName, GaryGi(0), _
                                 str(intCompCode), strBillStartDate, strBillEndDate, _
                                 IIf(blnShowMsg, frmSO1623A.pbr1, Nothing), rsResult, strResult, _
                                strRetErrMsg, strFaultReason, strICCUID) Then GoTo 88
            
        Call WriteRecordVodProcedure("2.3 �����s�W����x��� --> " & strCMD)
    End If
    If strProcessingDate <> Empty And blnExeCommand Then
        strResult = "1"
        OpenVod = True
        Call WriteRecordVodProcedure("2.2 �w����� --> " & strCMD)
        GoTo 88
    End If
    Call WriteRecordVodProcedure("2.4 �i�J��s��� --> " & strCMD)
    rsData.AbsolutePosition = varBookMark
    Call WriteRecordVodProcedure("2.4.1 �}�l��s��� --> " & strCMD)
    Select Case UCase(strCMD)
        Case "A1"
            If rsResult Is Nothing Then
                strRetErrMsg = "�L�^��VOD�b��"
                strResult = "0"
                Call WriteRecordVodProcedure("-666 �L�^��VOD�b�� ")
                GoTo 88
            End If
            
            If Not GetRS(rsSO004G, "Select * From " & GetOwner & "SO004G Where " & _
                        " 1=0 ", gcnGi, _
                        adUseClient, adOpenKeyset, adLockOptimistic) Then GoTo 88
            
            rsSO004G.AddNew
            rsSO004G("VODCredit") = Val(GetRsValue("Select VODCredit From " & GetOwner & "SO041 Where SysId=" & intCompCode, gcnGi) & "")
            rsSO004G("UpdEn") = strUpdName
            rsSO004G("UpdTime") = GetDTString(strUpdDate)
            rsSO004G("PrePay") = 0
            rsSO004G("Present") = 0
            rsSO004G("SendCount") = 0
            rsSO004G("VODAccountId") = rsResult("AccountUID").Value
            rsSO004G.Update
            rsData("VODStatus") = 1
            rsData("VodAccountId") = rsResult("AccountUID").Value
            '#5603 �n�NSO555.ICCUID�^���SOAC0201B By Kin 2010/03/29
            If Val(rsResult("ICCUID") & "") > 0 Then
                strUPDSOAC0201B = "UpDate " & GetOwner & "SOAC0201B " & _
                    " SET ICCUID=" & rsResult("ICCUID") & _
                    " WHERE FACISNO='" & strSmartCardID & "'"
                gcnGi.Execute strUPDSOAC0201B
            End If
        Case "A2"
            '#5722 �e�R�O���ɭ�SO555�n��JSOAC0201B.ICCUID,���\��n�M��SOAC0201B.ICCUID By Kin 2010/07/21
            rsData("VODStatus") = 0
            strUPDSOAC0201B = "UPDATE " & GetOwner & "SOAC0201B " & _
                " SET ICCUID = NULL " & _
                " WHERE FaciSNo = '" & strSmartCardID & "' " & _
                " AND COMPCODE = " & intCompCode
            gcnGi.Execute strUPDSOAC0201B
        Case "A3"
            rsData("VODStatus") = 2
        Case "A4"
            rsData("VODStatus") = 1
            If Not GetRS(rsSO004G, "Select StopType From " & GetOwner & "SO004G Where " & _
                " VODAccountID = " & rsData("VODAccountID"), gcnGi, _
                adUseClient, adOpenKeyset, adLockOptimistic) Then GoTo 88
            If Not rsSO004G.EOF Then
                rsSO004G.Fields(0) = 0
                rsSO004G.Update
            End If
        '#5469 �p�G�O�����n�N�³]�Ƨאּ0 By Kin 2010/01/07
        Case "A6"
            '#5497 �n�N�s�]�ƪ����A�令�³]�ƪ����A By Kin 2010/03/04
            '#5277 �n�N�ª�smartcardno��SOAC0201B.ICCUID�M��,�N�s��smartcardno��Soac0201B.ICCUID�� upd���s�� By Kin 2010/07/21
            rs2("VodStatus") = Val(rsData("VODStatus") & "")
            rs2.Update
            rsData("VODStatus") = 0
            If Not rsResult Is Nothing Then
                If rsResult("ICCUID") & "" <> "" Then
                    If InStr(1, rsResult("ICCUID"), ",") > 0 Then
                        aOLDICCUID = Mid(rsResult("ICCUID"), 1, InStr(1, rsResult("ICCUID"), ",") - 1)
                        aNewICCUID = Mid(rsResult("ICCUID"), InStr(1, rsResult("ICCUID"), ",") + 1)
                        'aOLDICCUID = Split(rsResult("ICCUID"), ",")(1)
                        If aOLDICCUID <> "" Then
                            strUPDSOAC0201B = "UPDATE " & GetOwner & "SOAC0201B " & _
                                " SET ICCUID = NULL " & _
                                " WHERE FACISNO ='" & Split(strSmartCardID, ",")(0) & "'" & _
                                " AND COMPCODE = " & intCompCode
                            gcnGi.Execute strUPDSOAC0201B
                        End If
                        If aNewICCUID <> "" Then
                            strUPDSOAC0201B = "UPDATE " & GetOwner & "SOAC0201B " & _
                                " SET ICCUID = " & aNewICCUID & _
                                " WHERE FACISNO = '" & Split(strSmartCardID, ",")(1) & "'" & _
                                " AND COMPCODE = " & intCompCode
                            gcnGi.Execute strUPDSOAC0201B
                        End If
                    End If
                    
                End If
            End If
    End Select
    rsData("UpdEn") = strUpdName
    rsData("UpdTime") = GetDTString(strUpdDate)
    
    rsData.Update
    Call WriteRecordVodProcedure("2.5 ��s��Ƨ��� --> " & strCMD)
    strResult = "1"
    OpenVod = True
88:
    If blnShowMsg Then
        ShowMsg "(1) " & strCMD & "  " & strCmdName, strResult, strFaultReason, strRetErrMsg
    End If
    On Error Resume Next
    Call CloseRecordset(rsSO004G)
    Exit Function
chkErr:
  Call ErrSub(FormName, "OpenVod")
End Function
Public Function GetICCUID(ByVal aSmartCardNo As String, ByVal aCompCode As String) As String
  On Error GoTo chkErr
    Dim aSQL As String
    Dim aRet As String
    aSQL = "SELECT ICCUID FROM " & GetOwner & "SOAC0201B " & _
        " WHERE FaciSNo = '" & aSmartCardNo & "' " & _
        " AND COMPCODE = " & aCompCode
    
    If aSmartCardNo = "" Then
        GetICCUID = ""
        Exit Function
    End If
    aRet = GetRsValue(aSQL, gcnGi) & ""
    GetICCUID = aRet
    Exit Function
chkErr:
  Call ErrSub(FormName, "GetICCUID")
End Function
Private Function TurnNum(ByVal strChgNum) As Long
  On Error GoTo chkErr
    Dim i As Long
    Dim strNum As String
    For i = 0 To Len(strChgNum) - 1
        If IsNumeric(Mid(strChgNum, i + 1, 1)) Then
            strNum = strNum & Mid(strChgNum, i + 1, 1)
        End If
    Next i
    TurnNum = Val(strNum)
  Exit Function
chkErr:
    ErrSub FormName, "TurnNum"
End Function




Public Function GetStateData(objCn, ByVal intCompCode As Integer) As String
  On Error Resume Next
    GetStateData = objCn.Execute("SELECT CmdStatus FROM " & strCMowner & "SO555" & _
                                                        " WHERE  SEQNO='" & strCmdSeqNo & "'")(0) & ""
                                                        
                                                
End Function
Public Function GetErrMsg(objCn, ByVal intCompCode As Integer, ByRef strErrMsg As String, ByRef strFaultReason As String) As Boolean
    On Error Resume Next
    Dim varErrPara As String
    GetErrMsg = False
    varErrPara = objCn.Execute("Select RetMsg,SysMsg From " & strCMowner & "SO555" & _
                                                " Where  SEQNO='" & strCmdSeqNo & "'").GetString(, , ",")
    strErrMsg = Split(Replace(varErrPara, Chr(13), ""), ",")(1)
    strFaultReason = Split(varErrPara, ",")(0)
    If (strErrMsg = "") And (strFaultReason <> "") Then
        strErrMsg = strFaultReason
        strFaultReason = ""
    End If
    'If strErrMsg <> "" Then strErrMsg = "���~�N�X:" & strErrMsg
    If Trim(strErrMsg) <> "" Or Trim(strFaultReason) <> "" Then
        GetErrMsg = True
    End If
End Function
Private Sub ShowMsg(strCMD As String, strResult As String, strFaultReason As String, strErrMsg As String)
    On Error GoTo chkErr
        If strResult = "1" Then
            MsgBox "[ " & strCMD & " ] ����T���w�e�X���� !! ", vbInformation, gimsgPrompt
        Else
            If strFaultReason = Empty Then
                MsgBox "[ " & strCMD & " ] ���� !! " & vbCrLf & vbCrLf & _
                                "���~��] : " & strErrMsg, vbInformation, gimsgPrompt
            Else
                MsgBox "[ " & strCMD & " ] ���� !! " & vbCrLf & vbCrLf & _
                                "���~��] : " & strFaultReason & vbCrLf & vbCrLf & _
                                strErrMsg, vbInformation, gimsgPrompt
            End If
        End If
    Exit Sub
chkErr:
    ErrSub FormName, "subShowMessage"
End Sub
Public Function GetClassID(ByRef rs As ADODB.Recordset, Optional strBPcode As String = "") As String
    On Error Resume Next
        Dim strClassIdFld As String
        strClassIdFld = IIf(IsCMCP(rs), "EMCCMCP", "EMCCM")
        If strBPcode = Empty Then
            If rs("BPcode") & "" = "" Then Exit Function
            GetClassID = GetRsValue("SELECT " & strClassIdFld & " FROM " & GetOwner & "CD078D" & _
                                                        " WHERE BPCODE='" & rs("BPcode") & "'" & _
                                                        " AND SERVICETYPE='" & rs("ServiceType") & "'")
        Else
            GetClassID = GetRsValue("SELECT " & strClassIdFld & " FROM " & GetOwner & "CD078D" & _
                                                        " WHERE BPCODE = '" & strBPcode & "'" & _
                                                        " AND SERVICETYPE='" & rs("ServiceType") & "'")
        End If
        If Err.Number <> 0 Then GetClassID = Empty
End Function
Public Function IsCMCP(ByRef rs As ADODB.Recordset) As Boolean
  On Error GoTo chkErr
        IsCMCP = (GetRsValue("SELECT COUNT(*) FROM " & GetOwner & "SO004" & _
                                            " WHERE REFACISNO='" & rs("FaciSno") & "'" & _
                                            " AND COMPCODE=" & rs("CompCode") & _
                                            " AND CUSTID=" & rs("CustID") & _
                                            " AND SERVICETYPE='P'" & _
                                            " AND PRDATE IS NULL") > 0)

  Exit Function
chkErr:
    ErrSub FormName, "IsCMCP"
End Function
Public Function IsEMTA(ByVal strFaciCode As String) As Boolean
  On Error Resume Next
    Dim rs As New ADODB.Recordset
    IsEMTA = False
    Set rs = gcnGi.Execute("Select REFNO From " & GetOwner & "CD022 Where CODENO=" & strFaciCode)
    If Not rs.EOF Then
        If rs(0) = 5 Then IsEMTA = True
    End If
    CloseRecordset rs
    
End Function
Public Function GetInstAddNo(ByRef rs As ADODB.Recordset) As String
  On Error GoTo chkErr
    Dim strAddNoSQL As String
    Dim strReInsAddrNoSQL As String
    Dim strRet As String
    If rs.EOF Then Exit Function
    '#5149 �W�[�ˬd �O�_��CATV���������u�����˦a�}�i�� By Kin 2009/06/18
    strReInsAddrNoSQL = "Select ReInstAddrNo From " & GetOwner & "SO009 " & _
                " Where PRCODE IN (SELECT CODENO FROM CD007 WHERE REFNO = 3) " & _
                " And ServiceType='C'" & _
                " AND FINTIME IS NULL AND RETURNCODE IS NULL AND RETURNDESCCODE IS NULL " & _
                " AND CUSTID =" & rs("CustId")
    strRet = GetRsValue(strReInsAddrNoSQL, gcnGi) & ""
    If strRet = "" Then
        strAddNoSQL = "Select InstAddrNo From " & GetOwner & "SO001 Where Custid=" & rs("Custid") & " And CompCode=" & rs("CompCode")
        strRet = GetRsValue(strAddNoSQL, gcnGi) & ""
    End If
    GetInstAddNo = strRet & ""
    Exit Function
chkErr:
  ErrSub FormName, "GetInstAddNo"
End Function
'**********************************************************************************************************
'Corey �g��Function
'�Q�ζǤJ��ADDRNO(�@�w�n�˾��a�})��XSO014.ADDSORT�A�åBSO017C���[�\�a�}���]�t��ADDRSORT��SNO�C�X�ӡA
'�åB�Q��SNO��XSO017B.SwitchID�^�ǡA�γr���j�}�C��SwitchID
Public Function GetSwitchID(ByVal strInstAddrNO As String, ByRef strErrString As String) As String
On Error GoTo chkErr
    Dim rsTmp As New ADODB.Recordset
    Dim strAddrSort As String
    Dim strSNO As String
    Dim strSwitchID As String
    GetSwitchID = ""
    strSwitchID = ""
    strErrString = ""
    If Not GetRS(rsTmp, "Select AddrSort,FTTxFlag From " & GetOwner & "SO014 Where AddrNO=" & strInstAddrNO) Then Exit Function
    If rsTmp.RecordCount = 0 Then
        strErrString = "�䤣������a�}"
        Exit Function '�䤣��������a�} Addrsort
    End If
'    If Val(rsTmp("FTTxFlag") & "") = 0 Then
'        strErrString = "�Ӧa�}���b FTTX �[�\�d��"
'        Exit Function
'    End If

    strAddrSort = Mid(rsTmp("AddrSort") & "", 1, GetFldMaxLen("SO017C", "AddrSortA"))
    If Not GetRS(rsTmp, "Select * From " & GetOwner & "SO017C Where AddrSortA <= '" & strAddrSort & "' and AddrSortB >= '" & strAddrSort & "'") Then Exit Function
    '�X�G���󪺼ƶq�j��0�h�}�l��y�k�A�_�h�^�ǪŦr��
    If rsTmp.RecordCount > 0 Then
        '�����oSO017C��SNO�A�A�Q��SNO��XSO017B��SwitchID
        rsTmp.MoveFirst
        With rsTmp
            Do While Not .EOF
                strSNO = GetRsValue("Select SwitchID From " & GetOwner & "SO017B where SNO=" & rsTmp("SNO") & "") & ""
                strSwitchID = strSwitchID & "," & Trim(strSNO)
                .MoveNext
            Loop
            strSwitchID = Mid(strSwitchID, 2)
        End With
        GetSwitchID = Trim(strSwitchID)
        If strSwitchID = "" Then strErrString = "�䤣�������SwitchID"
    Else
         strErrString = "�䤣�������SwitchID"
    End If
    On Error Resume Next
    CloseRecordset rsTmp
Exit Function
chkErr:
    ErrSub FormName, "GetSwitchID"
End Function
Public Function GetFldMaxLen(ByVal uTableName As String, ByVal uColumn As String) As Double

  On Error GoTo chkErr
  
    '��Function�O���X��Ʈw��������SCHEMA�����j�p��
    '�åBSCHEMA��DATA_TYPE=NUMBER�A�íp��XNUMBER�̤j�Ȧ^��Double
    '�t�ƪ��ɭԬO�q -1.79769313486232E308 �� -4.94065645841247E-324
    '���ƪ��ɭԬO�q 4.94065645841247E-324 �� 1.79769313486232E308
    
    'uTableName=�ǤJTable�W��  uColumn=�ǤJ���W��

    Dim rsColumn As New ADODB.Recordset
    Dim strSQL As String
    
    GetFldMaxLen = 0

    '�S���Ǹ�ƶi�ӫh�������}�A�_�h���y�k�ɷ|�X���~�T��
    If IsNull(uTableName) Or IsNull(uColumn) Then Exit Function
    
    strSQL = "Select * From " & "User_Tab_Columns" & _
                " Where Table_Name='" & UCase(uTableName) & "'" & _
                " And Column_Name='" & UCase(uColumn) & "'"
    
    If Not GetRS(rsColumn, strSQL) Then Exit Function
    
    Select Case UCase(rsColumn("Data_Type") & "")
        Case "NUMBER"
            GetFldMaxLen = (10 ^ rsColumn("Data_Precision")) - 1
            If rsColumn("Data_Scale") > 0 Then '�P�_�p���I�O�_����
                GetFldMaxLen = GetFldMaxLen + (0.1 ^ rsColumn("Data_Scale")) * ((10 ^ rsColumn("Data_Scale")) - 1)
            End If
        Case "VARCHAR2", "VARCHAR", "CHAR", "CLOB", "BLOB", "ROWID"
            GetFldMaxLen = Val(rsColumn("Data_Length") & "")
        Case "LONG", "DATE", "UNDEFINED"
            GetFldMaxLen = 0
    End Select
    
    Call CloseRecordset(rsColumn)

  Exit Function
chkErr:
    ErrSub FormName, "GetFldMaxLen"
End Function
'*******************************************************************************************
Public Function chkUseFlag(ByVal strFaciSNO As String, ByVal intComp As Integer) As Boolean
  On Error GoTo chkErr:
    Dim rs As New ADODB.Recordset
    chkUseFlag = False
    'If blnUseFttx Then chkUseFlag = True: Exit Function
    
    If Not GetRS(rs, "Select Useflag From " & GetOwner & "SO306 Where HFCMAC='" & strFaciSNO & "'") Then Exit Function
    If rs.EOF Then Call CloseRecordset(rs): Exit Function
    '******************************************************************************************************
    '#3755 �p�GLinkmss��0���ˬdUseFlag By Kin 2008/2/01
    If Val(GetRsValue("Select Nvl(Linkmss,0) From " & GetOwner & "CD039 Where CodeNo=" & intComp)) = 0 Then
        On Error Resume Next
        chkUseFlag = True
        Call CloseRecordset(rs)
        Exit Function
    End If
    '******************************************************************************************************
    If rs(0).Value <> 1 Then Call CloseRecordset(rs): Exit Function
    chkUseFlag = True
    Call CloseRecordset(rs)
    Exit Function
chkErr:
    ErrSub FormName, "chkUseFlag"
End Function

Public Sub GetScmData(ByRef objCn As Object, ByRef rsResult As Object, ByVal strFldName As String)
  On Error Resume Next
    Dim strScmData As String
    'Dim strFldName As String
    
    strScmData = objCn.Execute("SELECT " & strFldName & " FROM " & strCMowner & "SO555" & _
                                           " WHERE  SEQNO='" & strCmdSeqNo & "'").GetString(, , ";") & ""
    
    If Right(strScmData, 1) = ";" Then
        strScmData = Mid(strScmData, 1, Len(strScmData) - 1)
    End If
    
    If strScmData <> Empty Then
     '   strFldName = "AccountUID"
        Call SetRsMemory(rsResult, strFldName, strScmData)
    Else
        Set rsResult = Nothing
    End If
End Sub
Public Sub SetRsMemory(ByRef rs As Object, ByVal strFldName As String, ByVal strFldValue As String)
  On Error GoTo chkErr
    Dim i As Long
    Dim strAryFldName() As String
    Dim strAryFldValue() As String
    Set rs = CreateObject("ADODB.Recordset")
    If rs.State = adStateOpen Then rs.Close
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenKeyset
    rs.LockType = adLockOptimistic
    strAryFldName = Split(strFldName, ",")
    If Left(strFldValue, 1) = "," Then
        strFldValue = Chr(0) & Right(strFldValue, Len(strFldValue) - 1)
    End If
    strAryFldValue = Split(strFldValue, ";")
    For i = LBound(strAryFldName) To UBound(strAryFldName)
        rs.Fields.Append strAryFldName(i), adBSTR, 100
    Next i
    rs.Open
    rs.AddNew
    For i = LBound(strAryFldName) To UBound(strAryFldName)
        If Left(strAryFldValue(i), 1) = Chr(0) Then
            strAryFldValue(i) = "," & Right(strAryFldValue(i), Len(strAryFldValue(i)) - 1)
        End If
        rs.Fields(i).Value = strAryFldValue(i) & ""
    Next i
    rs.Update
    Exit Sub
chkErr:
    ErrSub FormName, "SetRsMemory"
End Sub

Public Function GetVendCode(ByVal intComp As Integer) As String
  On Error Resume Next
    GetVendCode = gcnGi.Execute("Select VendorCode From " & GetOwner & "CM006 Where " & _
                                "CompCode=" & intComp & " And StopFlag=0")(0)
                        

End Function
Private Function UpdSO004F(ByRef rs As Recordset, ByVal strUpdTime As String) As Boolean
  On Error GoTo chkErr
    Dim strQrySQL As String
    Dim i As Integer
    Dim strDialAccount, strExtendAccount, strExtendPassword As String
    Dim rsQry As New ADODB.Recordset
    If rs(2).Value = "" Then
        UpdSO004F = True
        Exit Function
    End If
    strQrySQL = "Select * From " & GetOwner & "SO004F" & _
                " Where ExtendAccount='" & rs("FixIPAccountId") & "'" & _
                " And ExtendType=1 And PrDate is Null"
    If Not GetRS(rsQry, strQrySQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
    If Not rsQry.EOF Then
        strDialAccount = rsQry("DialAccount") & ""
        strExtendAccount = rsQry("ExtendAccount") & ""
        strExtendPassword = rsQry("ExtendPassword") & ""
        rsQry.MoveFirst
        Do While Not rsQry.EOF
            rsQry("PrDate") = Format(strUpdTime, "YYYY/MM/DD HH:MM:SS")
            rsQry("UpdEn") = GaryGi(1)
            rsQry("UpdTime") = GetDTString(strUpdTime)
            rsQry.Update
            rsQry.MoveNext
        Loop
    End If
    For i = 0 To rs.Fields.Count - 3
        With rsQry
            .AddNew
            .Fields("SeqNo") = Get_SO004F_Seq_No(gcnGi)
            .Fields("DialAccount") = strDialAccount
            .Fields("ExtendAccount") = strExtendAccount
            .Fields("ExtendPassword") = strExtendPassword
            .Fields("FixIPAddress") = rs(i + 2).Value & ""
            .Fields("UpdEn") = GaryGi(1)
            .Fields("UpdTime") = GetDTString(strUpdTime)
            .Update
        End With
    Next i
    UpdSO004F = True
    On Error Resume Next
    Call CloseRecordset(rsQry)
    
    Exit Function
chkErr:
  Call ErrSub(FormName, "UpdSO004F")
End Function
