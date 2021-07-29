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
Public IsA6Command As Boolean
Public IsRestCommand As Boolean
Public OldVodAccountId As String
Public strFaciSNo As String
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
Public blnReturnCmd As Boolean
Public blnHasRunReturn As Boolean
Public intRunRetunCount As Integer
Public Const strA1ErrCode = "120116"
Public Const strA2ErrCode = "120120"
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
  On Error GoTo ChkErr
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
ChkErr:
  Call ErrSub(FormName, "GetChannelCitemCode")
End Function

'#5692 ���oSVOD�W�D(aStatusType=0 �q��SVOD) By Kin 2010/06/15
Public Function GetChannelID(ByVal aCitemCode As String, ByVal aCompCode As String, _
                            Optional ByVal aCustId As String = Empty, _
                            Optional ByVal aStatusType As Integer = 0, _
                            Optional ByVal aVodAccountId As String) As String
  On Error GoTo ChkErr
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
    '�p�G�SVodAccountId�N���n��U�����y�k�F�A���M�|�� By Kin 2014/06/05 For Corey
    If Len(aVodAccountId) > 0 Then
        If aStatusType = 0 Then '0�N��n�}�qSVOD 1�N��n��SVOD
            aSQL = aSQL & " AND CD024.CODENO NOT IN(SELECT ProductID FROM " & GetOwner & "SO555B " & _
                    " WHERE NVL(STATUS,0)=1 AND VODACCOUNTID = " & aVodAccountId & ") "
        Else
            aSQL = aSQL & " AND CD024.CODENO NOT IN(SELECT ProductID FROM " & GetOwner & "SO555B " & _
                    " WHERE NVL(STATUS,0)=0 AND VODACCOUNTID = " & aVodAccountId & ") "
        End If
    End If
    If Not GetRS(rsCD024, aSQL, gcnGi, adUseClient, adOpenKeyset, , , , , False) Then Exit Function
    Do While Not rsCD024.EOF
        If aRet = "" Then
            aRet = rsCD024("ChannelID") & ""
        Else
            aRet = aRet & "," & rsCD024("ChannelID")
        End If
        rsCD024.MoveNext
    Loop
    GetChannelID = aRet
    On Error Resume Next
    Call CloseRecordset(rsCD024)
    Exit Function
ChkErr:
  Call ErrSub(FormName, "GetChannelID")
End Function
Public Function Get_EMTA_Mac(strFaciSNo As String) As String
  On Error GoTo ChkErr
    If strFaciSNo <> Empty Then
        Get_EMTA_Mac = GetRsValue("SELECT MTAMAC FROM " & GetOwner & "SO306" & _
                                                        " WHERE HFCMAC='" & strFaciSNo & "'") & ""
       
    End If
  Exit Function
  
ChkErr:
    ErrSub FormName, "Get_EMTA_Mac"
End Function
Public Function Get_EMTA_Mode(strFaciSNo As String) As String
  On Error Resume Next
    Get_EMTA_Mode = GetRsValue("SELECT Nvl(ModeID,0) FROM " & GetOwner & "SO306" & _
                                                        " WHERE HFCMAC='" & strFaciSNo & "'") & ""
End Function
Public Function GetMacAddress(ByVal aFaciSNo As String) As String
  On Error GoTo ChkErr
    Dim rsMac As New ADODB.Recordset
    Dim strMac As String
    Dim strTmp As String
    Dim i As Integer
    If aFaciSNo = "" Then
        strMac = ""
    Else
        If Not GetRS(rsMac, "Select EthernetMAC From " & GetOwner & "SOAC0201A Where " & _
                        "FaciSNo='" & aFaciSNo & "'", gcnGi, adUseClient, adOpenKeyset, adLockOptimistic, , , , False) Then Exit Function
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
ChkErr:
  ErrSub FormName, "GetMacAddress"
End Function
'#6752 �W�[7��Smit Vod�R�O By Kin 2014/04/08
Public Function SetSmitVodCmd(ByRef rsData As ADODB.Recordset, ByVal strCMD As String, _
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
                        Optional ByVal strICCUID As String, Optional ByVal strCitemStr As String = "", _
                        Optional ByVal strNotes As String = "", Optional strDVRQuota As String = "", Optional strTunerCount As String = "", _
                         Optional ByVal strCitemCodes As String = "", _
                        Optional ByRef rsResult As Object = Nothing, Optional strResult As String, _
                        Optional ByRef strRetErrMsg As String, Optional ByRef strFaultReason As String) As Boolean
    On Error GoTo ChkErr
       
        Dim varBookMark As Variant
        Dim strCmdName As String
        Dim strUpdDate As String
        Dim strUpdName As String
        Dim objVodCommand As New clsVodcommand
        Dim intPVRCMDType As Byte
        intPVRCMDType = 0
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
            Case "A7"
                strCmdName = "�]�w���W��Tuner�ƶq"
            Case "B13"
                strCmdName = "�]�w���v�e�q"
            Case "B14"
                strCmdName = "�������v�e�q"
            Case "B15"
                strCmdName = "���W���}�q�M�\Prod ID"
            Case "B16"
                strCmdName = "���W�������M�\Prod ID"
            Case "B17"
                strCmdName = "�}��CatchUp"
            Case "B18"
                strCmdName = "����CatchUp"
        End Select
        '�ϦV�Y�ɩR�O�p�GVodAccountId�O�ŭ�,�h���n�e�R�O,�����^���\ For Jacky By Kin 2014/0716
        If (IsCancelCmd) And (Len(strProcessingDate & "") = 0) And (Len(rsData("VodAccountId") & "") = 0) Then
            strResult = "1"
            SetSmitVodCmd = True
            Call WriteRecordVodProcedure("2.2 �ϦV�Y�ɦӥBVodAccountId�L�ȡA���e�R�O! --> " & strCMD)
            GoTo 88
        End If
        '�W�[�P�_Jacky jacky: CD043.PVRCMDType = 0 �~�e�R�O By Kin 2015/05/07
        intPVRCMDType = qryPVRCMDType(rsData("ModelCode") & "")
        If intPVRCMDType = 0 Then
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
                                        strICCUID, strCitemStr, , , strNotes, strDVRQuota, strTunerCount, strCitemCodes, _
                                        , , , GetPrbCtl, rsResult, strResult, strRetErrMsg, _
                                        strFaultReason) Then GoTo 88
    
                Call WriteRecordVodProcedure("2.3 �����s�W����x��� --> " & strCMD)
            End If
        Else
            blnShowMsg = False
        End If
        If strProcessingDate <> Empty And blnExeCommand Then
            strResult = "1"
            SetSmitVodCmd = True
            Call WriteRecordVodProcedure("2.4 �w����� --> " & strCMD)
            GoTo 88
        End If
        strResult = "1"
        SetSmitVodCmd = True
        Call WriteRecordVodProcedure("2.5 ������s��� --> " & strCMD)
88:
        If blnShowMsg Then
            If IsA6Command Then
            
            Else
                ShowMsg "(2) " & strCMD & "  " & strCmdName, strResult, strFaultReason, strRetErrMsg
            End If
        End If
        On Error Resume Next
        'Call CloseRecordset(rs555)
        Exit Function
     Exit Function
ChkErr:
  Call ErrSub(FormName, "SetSmitVodCmd")
End Function
Public Function QueryOTT(ByRef rsData As ADODB.Recordset, ByVal strCMD As String, _
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
                        Optional ByVal strICCUID As String, Optional ByVal strCitemStr As String = "", _
                        Optional ByVal strCitemCodes As String = "", _
                        Optional ByVal strOTTUserID As String = "", _
                        Optional ByVal striOSOrderNo As String = "", _
                        Optional ByVal strFVOrderNo As String = "", _
                        Optional ByRef rsResult As Object = Nothing, Optional strResult As String, _
                        Optional ByRef strRetErrMsg As String, Optional ByRef strFaultReason As String) As Boolean
                        
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
        Case "C2"
            strCmdName = "IOS��d���A�d��"
    End Select
    If (IsCancelCmd) And (Len(strProcessingDate & "") = 0) And (Len(rsData("VodAccountId") & "") = 0) Then
        strResult = "1"
        QueryOTT = True
        Call WriteRecordVodProcedure("2.2 �ϦV�Y�ɦӥBVodAccountId�L�ȡA���e�R�O! --> " & strCMD)
        GoTo 88
    End If
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
                                strICCUID, strCitemStr, , , , , , strCitemCodes, _
                                strOTTUserID, striOSOrderNo, strFVOrderNo, GetPrbCtl, rsResult, strResult, strRetErrMsg, _
                                strFaultReason) Then GoTo 88
            
        Call WriteRecordVodProcedure("2.3 �����s�W����x��� --> " & strCMD)
    End If
    If strProcessingDate <> Empty And blnExeCommand Then
        strResult = "1"
        QueryOTT = True
        Call WriteRecordVodProcedure("2.4 �w����� --> " & strCMD)
        GoTo 88
    End If
    rsData.AbsolutePosition = varBookMark
    strResult = "1"
    QueryOTT = True
    If blnShowMsg Then
         If Len(strRetErrMsg & "") > 0 Then
            MsgBox strRetErrMsg, vbInformation, "�T��"
        End If
    End If
   
88:
    If blnShowMsg Then
        If IsA6Command Or IsRestCommand Then
'            If UCase(strCMD) <> "B11" Then
'                ShowMsg "(1) " & strCMD & "  " & strCmdName, strResult, strFaultReason, strRetErrMsg
'            End If
        Else
            ShowMsg "(1) " & strCMD & "  " & strCmdName, strResult, strFaultReason, strRetErrMsg
        End If
    End If
    
    On Error Resume Next
    'Call CloseRecordset(rs005)
    'Call CloseRecordset(rs555)
    'Call CloseRecordset(rs005B)
    Exit Function
ChkErr:
  Call ErrSub(FormName, "QueryOTT")
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
                        Optional ByVal strICCUID As String, Optional ByVal strCitemStr As String = "", _
                        Optional ByVal strCitemCodes As String = "", _
                        Optional ByRef rsResult As Object = Nothing, Optional strResult As String, _
                        Optional ByRef strRetErrMsg As String, Optional ByRef strFaultReason As String) As Boolean
  On Error GoTo ChkErr
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
'    Dim objPbr1 As Object
'    If blnShowMsg Then
'        If frmSO1623A Is Nothing Then
'            Set objPbr1 = Nothing
'        Else
'            Set objPbr1 = frmSO1623A.pbr1
'        End If
'    Else
'        Set objPbr1 = Nothing
'    End If
    '�ϦV�Y�ɩR�O�p�GVodAccountId�O�ŭ�,�h���n�e�R�O,�����^���\ For Jacky By Kin 2014/0716
     If (IsCancelCmd) And (Len(strProcessingDate & "") = 0) And (Len(rsData("VodAccountId") & "") = 0) Then
        strResult = "1"
        SetVodCmd = True
        Call WriteRecordVodProcedure("2.2 �ϦV�Y�ɦӥBVodAccountId�L�ȡA���e�R�O! --> " & strCMD)
        GoTo 88
    End If
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
                                strICCUID, strCitemStr, , , , , , strCitemCodes, , , , GetPrbCtl, rsResult, strResult, strRetErrMsg, _
                                strFaultReason) Then GoTo 88
            
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
            If rs555 Is Nothing Then strRetErrMsg = "��SO555��ƿ��~" & vbCrLf & "SEQNO=" & strCmdSeqNo: strResult = "0": GoTo 88
            If rs555.State <> adStateOpen Then strRetErrMsg = "SO555�}�ҥ��ѡI": strResult = "0": GoTo 88
            If rs555.EOF Then strRetErrMsg = "�䤣��SO555 SeqNo=" & strCmdSeqNo & " ����ơI": strResult = "0": GoTo 88
            If Not UpdB11Cmd(rs555) Then strRetErrMsg = "��sSO555B���ѡI": strResult = "0": GoTo 88
        Case "B12"      '�H��JSO555����Ƭ��D,��sSO555B�����
            Call GetRs555(rs555)
            If rs555 Is Nothing Then strRetErrMsg = "��SO555��ƿ��~" & vbCrLf & "SEQNO=" & strCmdSeqNo: GoTo 88
            If rs555.State <> adStateOpen Then strRetErrMsg = "rs555�}�ҥ��ѡI": strResult = "0": GoTo 88
            If rs555.EOF Then strRetErrMsg = "�䤣��SeqNo=" & strCmdSeqNo & " ����ơI": strResult = "0": GoTo 88
            If Not UpdB12Cmd(rs555, strSNO) Then strRetErrMsg = "��sSO555B���ѡI": strResult = "0": GoTo 88
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
        If IsA6Command Or IsRestCommand Then
'            If UCase(strCMD) <> "B11" Then
'                ShowMsg "(1) " & strCMD & "  " & strCmdName, strResult, strFaultReason, strRetErrMsg
'            End If
        Else
            ShowMsg "(1) " & strCMD & "  " & strCmdName, strResult, strFaultReason, strRetErrMsg
        End If
    End If
    On Error Resume Next
    'Call CloseRecordset(rs005)
    Call CloseRecordset(rs555)
    'Call CloseRecordset(rs005B)
    Exit Function
ChkErr:
  Call ErrSub(FormName, "SetVod")
End Function
Private Function UpdB12Cmd(ByRef rs555 As ADODB.Recordset, Optional ByVal aPrSNo As String = Empty) As Boolean
  On Error GoTo ChkErr
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
ChkErr:
  Call ErrSub(FormName, "UpdB12Cmd")
End Function
Private Function UpdB11Cmd(ByRef rs555 As ADODB.Recordset) As Boolean
  On Error GoTo ChkErr
    Dim aCitemCode As String
    Dim aServiceType As String
    Dim rs555B As New ADODB.Recordset
    UpdB11Cmd = False
    rs555.MoveFirst
    Do While Not rs555.EOF
        aServiceType = Empty
        aCitemCode = Empty
        Call GetRs555B(rs555B, rs555("PID") & "", rs555("AccountNum") & "", rs555("AccountUID") & "", True)
        With rs555B
            If .EOF Then
                .AddNew
                'Call GetCitemCode(rs555("PID"), aCitemCode, aServiceType)
                .Fields("VODAccountID") = rs555("AccountUID") & ""
                .Fields("ProductID") = rs555("PID") & ""
                .Fields("ProductName") = rs555("PName") & ""
                .Fields("Status") = 1
                .Fields("OpenTime") = rs555("ExchTime")
                '.Fields("CitemCode") = aCitemCode
                .Fields("CitemCode") = NoZero(rs555("CitemCode") & "")
                .Update
            Else
                .Fields("Status") = 1
                .Fields("OpenTime") = rs555("ExchTime")
                .Fields("CloseTime") = Null
                .Update
            End If
        End With
        rs555.MoveNext
    Loop
    UpdB11Cmd = True
    On Error Resume Next
    Call CloseRecordset(rs555B)
    Exit Function
ChkErr:
  Call ErrSub(FormName, "UpdB11Cmd")
End Function

Public Sub GetCitemCode(ByVal strCode As String, ByRef aCCode As String, ByRef aService As String)
  On Error GoTo ChkErr
    Dim strSQL As String
    Dim rsCD019A As New ADODB.Recordset
    strSQL = "Select A.CitemCode,C.ServiceType From " & GetOwner & "CD019A A," & _
            GetOwner & "CD024 B," & GetOwner & "CD019 C " & _
            " Where A.CodeNo='" & strCode & "' And A.CodeNo=B.CodeNO " & _
            " And B.VodType=2 And B.StopFlag<> 1 And A.CitemCode=C.CodeNo"
    If Not GetRS(rsCD019A, strSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic, , , , False) Then Exit Sub
    If Not rsCD019A.EOF Then
        aCCode = rsCD019A("CitemCode")
        aService = rsCD019A("ServiceType")
    End If
    Exit Sub
ChkErr:
  Call ErrSub(FormName, "GetCitemCode")
End Sub
Public Sub GetRs555B(ByRef rs As ADODB.Recordset, ByVal strProductID As String, _
                    ByVal aCustId As String, ByVal aVodAccountId As String, _
                    ByVal aOrder As Boolean)
  On Error GoTo ChkErr
    Dim strSQL As String
'    strSQL = "Select A.RowID,A.* From " & GetOwner & "SO555B A " & _
'            " Where CustId=" & aCustId & " And ProductID='" & strProductID & "'" & _
'            " AND VODAccountID=" & aVODAccountId & " And Status = " & IIf(aOrder, "1", "0")
            
    strSQL = "Select A.RowID,A.* From " & GetOwner & "SO555B A " & _
            " Where  ProductID='" & strProductID & "'" & _
            " AND VODAccountID=" & aVodAccountId & " And Status = " & IIf(aOrder, "0", "1")
    If Not GetRS(rs, strSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic, , , , False) Then Exit Sub
  Exit Sub
ChkErr:
  Call ErrSub(FormName, "GetRs005")
End Sub
Public Sub GetRs555(ByRef rs As ADODB.Recordset)
  On Error GoTo ChkErr
    Dim strSQL As String
    strSQL = "Select A.*,B.ProductId PID,B.ProductName PName,C.ChannelID,B.CitemCode From " & _
                strCMowner & "SO555 A," & _
                strCMowner & "SO555A B," & strCMowner & "CD024 C " & _
                " Where A.SeqNo=" & strCmdSeqNo & " And A.SeqNo=B.SeqNo " & _
                " And C.CodeNo=B.ProductId And C.VodType=2 And C.StopFlag<>1"
  
    If Not GetRS(rs, strSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic, , , , False) Then
        Set rs = Nothing
    End If
    Exit Sub
ChkErr:
  Call ErrSub(FormName, "GetRs555")
End Sub
Public Function GetSO555BCitemCode(ByVal aVodAccountId As String, Optional ByVal aStatus As Integer = 1) As String
  On Error GoTo ChkErr
    Dim aSQL As String
    Dim aRet As String
    Dim rsTmp As New ADODB.Recordset
    aSQL = "SELECT CitemCode FROM " & strCMowner & "SO555B " & _
            " WHERE VODAccountID = " & aVodAccountId & _
            " AND Status = " & aStatus
    If Not GetRS(rsTmp, aSQL, gcnGi, adUseClient, adOpenKeyset, adLockPessimistic) Then Exit Function
    Do While Not rsTmp.EOF
        If Len(aRet) = 0 Then
            aRet = rsTmp("CitemCode") & ""
        Else
            aRet = aRet & "," & rsTmp("CitemCode") & ""
        End If
        rsTmp.MoveNext
    Loop
    GetSO555BCitemCode = aRet
    On Error Resume Next
    CloseRecordset rsTmp
  Exit Function
ChkErr:
  Call ErrSub(FormName, "GetProductID")
End Function

'5696 �p�G�OA6�R�O�n�h��ProductID By Kin 2010/06/17
Public Function GetProductId(ByVal aVodAccount As String, Optional ByVal aStatus As Integer = 1) As String
  On Error GoTo ChkErr
    Dim aSQL As String
    Dim strCodeNo As String
    Dim rsSO555B As New ADODB.Recordset
    Dim rsCD024 As New ADODB.Recordset
    Dim aRet As String
    
    aSQL = "SELECT ProductID FROM " & strCMowner & "SO555B " & _
        " WHERE NVL(Status,0)=" & aStatus & " AND VODAccountID = " & aVodAccount
    If Not GetRS(rsSO555B, aSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic, , , , False) Then Exit Function
    If rsSO555B.EOF Then GetProductId = Empty: Exit Function
    Do While Not rsSO555B.EOF
        If strCodeNo = "" Then
            strCodeNo = "'" & rsSO555B("ProductID") & "'"
        Else
            strCodeNo = strCodeNo & ",'" & rsSO555B("ProductID") & "'"
        End If
        rsSO555B.MoveNext
    Loop
    aSQL = "SELECT ChannelID FROM " & strCMowner & "CD024 " & _
            " WHERE CodeNo IN(" & strCodeNo & ")"
    If Not GetRS(rsCD024, aSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic, , , , False) Then Exit Function
    Do While Not rsCD024.EOF
        If aRet = "" Then
            aRet = rsCD024("ChannelID")
        Else
            aRet = aRet & "," & rsCD024("ChannelID")
        End If
        rsCD024.MoveNext
    Loop
    GetProductId = aRet
    On Error Resume Next
    Call CloseRecordset(rsSO555B)
    Exit Function
ChkErr:
  Call ErrSub(FormName, "GetProductID")
End Function
Public Function QueryAccInfo(ByRef rsData As ADODB.Recordset, ByVal strCMD As String, _
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
                        Optional ByVal strICCUID As String = "", Optional ByVal strOldSerialNumber As String = "", _
                        Optional ByVal strOldSmartCardNo As String = "", Optional strNotes As String = "", _
                        Optional strDVRQuota As String = "", Optional strTunerCount As String = "", _
                        Optional strCitemCodes As String = "", _
                        Optional ByRef rsResult As Object = Nothing, Optional strResult As String, _
                        Optional ByRef strRetErrMsg As String, Optional ByRef strFaultReason As String) As Boolean
                        
    On Error GoTo ChkErr
        Dim strCmdName As String
        Dim strUpdName As String
        Dim strUpdDate As String
        Dim intCompCode As Integer
        Dim aMsg As String
        Dim aCitemNames As String
        Dim objVodCommand As New clsVodcommand
         strCmdName = "�d�߱b���β��~ID"
        If strExchTime = "" Then
            strUpdDate = RightNow
        Else
             strUpdDate = strExchTime
        End If
        If strUpdEn = "" Then
             strUpdName = GaryGi(1)
        Else
            strUpdName = strUpdEn
        End If
         intCompCode = rsData("CompCode")
        If Not objVodCommand.Vod_Command(gcnGi.ConnectionString, intCompCode & "", GetOwner, strCMD _
                                    , , , , , strUpdDate, , , strServiceProviderID, strAccountNum, _
                                    strAccoutName, , strAddress, strPostCode, strAccountUID, strLoginID, _
                                    strPin, , , strSmartCardID, strMacAddress, strSerialNumber, , , , strCASNo _
                                     , strProductID, strProductName, strUpdTime, strUpdName, _
                                     strSTBUId, strProcessingDate, _
                                     strSNO, strMediaBillNo, strReasonCode, strReasonName, GaryGi(0), _
                                     str(intCompCode), strBillStartDate, strBillEndDate, _
                                       strICCUID, , strOldSerialNumber, strOldSmartCardNo, _
                                        strNotes, strDVRQuota, strTunerCount, _
                                         strCitemCodes, , , , _
                                       GetPrbCtl, rsResult, strResult, _
                                    strRetErrMsg, strFaultReason) Then
            ShowMsg strCmdName, strResult, strFaultReason, strRetErrMsg
        Else
            If Not rsResult Is Nothing Then
                If (rsResult.RecordCount > 0) Then
                    If Len(rsResult("SMARTCARDID") & "") > 0 Then
                        aMsg = "�d���G" & rsResult("SMARTCARDID")
                    End If
                    If Len(rsResult("AccountUID") & "") > 0 Then
                        aMsg = aMsg & "VOD�b���G" & rsResult("AccountUId") & vbCrLf
                    End If
                    aCitemNames = GetProductName(rsResult("Productid"))
                    If Len(aCitemNames & "") > 0 Then
                        aMsg = aMsg & "���O���ئW�١G" & aCitemNames
                    End If
                    If Len(rsResult("AccountUID") & "") > 0 Then
                        MsgBox aMsg, vbInformation, "���G"
                    Else
                        MsgBox "�d�L��ơI", vbCritical, "ĵ�i"
                    End If
                    
                Else
                    MsgBox "�d�L��ơI", vbCritical, "ĵ�i"
                End If
            Else
                MsgBox "�d�L��ơI", vbCritical, "ĵ�i"
            End If
        
        End If
    
    Exit Function
ChkErr:
Call ErrSub(FormName, "OpenVod")
End Function
Private Function GetProductName(ByVal productid As String) As String
  On Error GoTo ChkErr
    If Len(productid & "") = 0 Then Exit Function
    Dim aResult As String
    Dim aSQL As String
    Dim arrayProductid() As String
    Dim aProductid As String
    Dim i As Integer
    Dim rsProduct As New ADODB.Recordset
    arrayProductid = Split(productid, ",")
    For i = LBound(arrayProductid) To UBound(arrayProductid)
        If Len(aProductid & "") = 0 Then
            aProductid = "'" & Replace(arrayProductid(i), "'", "") & "'"
        Else
            aProductid = aProductid & ",'" & Replace(arrayProductid(i), "'", "") & "'"
        End If
    Next i
    aSQL = "select distinct c.description from " & GetOwner & "cd024 a , " & GetOwner & "cd019a b ," & GetOwner & "cd019 c " & _
                " where channelid IN (" & aProductid & ") and nvl(a.stopflag,0)=0 and a.codeno=b.codeno and b.citemcode=c.codeno " & _
                " and nvl(c.stopflag,0) = 0 "
    If Not GetRS(rsProduct, aSQL, gcnGi, adUseClient, adOpenKeyset, adLockPessimistic) Then Exit Function
    Do While Not rsProduct.EOF
        If Len(aResult & "") = 0 Then
            aResult = rsProduct("DESCRIPTION")
        Else
            aResult = aResult & "," & rsProduct("DESCRIPTION")
        End If
        rsProduct.MoveNext
    Loop
    On Error Resume Next
    GetProductName = aResult
    Call CloseRecordset(rsProduct)
  Exit Function
ChkErr:
  Call ErrSub(FormName, "GetProductName")
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
                        Optional ByVal strICCUID As String = "", Optional ByVal strOldSerialNumber As String = "", _
                        Optional ByVal strOldSmartCardNo As String = "", Optional strNotes As String = "", _
                        Optional strDVRQuota As String = "", Optional strTunerCount As String = "", _
                        Optional strCitemCodes As String = "", _
                        Optional ByRef rsResult As Object = Nothing, Optional strResult As String, _
                        Optional ByRef strRetErrMsg As String, Optional ByRef strFaultReason As String) As Boolean
  On Error GoTo ChkErr
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
    Dim rsOldSO004G As New ADODB.Recordset
    
    
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
    strResult = "0"
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
'        Case "A7"
'            strCmdName = "�t��"
        Case "A8"
            strCmdName = "���״��� STB Swap"
        Case "A9"
            strCmdName = "���״��� Smart Card Swap"
        Case "C1"
            strCmdName = "�d�߱b���β��~ID"
    End Select
'    Dim objPbr1 As Object
'    If blnShowMsg Then
'        If frmSO1623A Is Nothing Then
'            Set objPbr1 = Nothing
'        Else
'            Set objPbr1 = frmSO1623A.pbr1
'        End If
'    Else
'        Set objPbr1 = Nothing
'    End If
'
    intCompCode = rsData("CompCode")
    '�ϦV�Y�ɩR�O�p�GVodAccountId�O�ŭ�,�h���n�e�R�O,�����^���\ For Jacky By Kin 2014/0716
    If (IsCancelCmd) And (Len(strProcessingDate & "") = 0) And (Len(rsData("VodAccountId") & "") = 0) Then
        strResult = "1"
        OpenVod = True
        Call WriteRecordVodProcedure("2.2 �ϦV�Y�ɦӥBVodAccountId�L�ȡA���e�R�O! --> " & strCMD)
        GoTo 88
    End If
    
    If (blnExeCommand) Then
        Call WriteRecordVodProcedure("2.2 �i�J�s�W����x��� --> " & strCMD)
        If Not objVodCommand.Vod_Command(gcnGi.ConnectionString, intCompCode & "", GetOwner, strCMD _
                                , , , , , strUpdDate, , , strServiceProviderID, strAccountNum, _
                                strAccoutName, , strAddress, strPostCode, strAccountUID, strLoginID, _
                                strPin, , , strSmartCardID, strMacAddress, strSerialNumber, , , , strCASNo _
                                 , strProductID, strProductName, strUpdTime, strUpdName, _
                                 strSTBUId, strProcessingDate, _
                                 strSNO, strMediaBillNo, strReasonCode, strReasonName, GaryGi(0), _
                                 str(intCompCode), strBillStartDate, strBillEndDate, _
                                   strICCUID, , strOldSerialNumber, strOldSmartCardNo, _
                                    strNotes, strDVRQuota, strTunerCount, _
                                     strCitemCodes, , , , _
                                   GetPrbCtl, rsResult, strResult, _
                                strRetErrMsg, strFaultReason) Then GoTo 88
            
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
                        adUseClient, adOpenKeyset, adLockOptimistic, , , , False) Then GoTo 88
            
            rsSO004G.AddNew
            rsSO004G("VODCredit") = Val(GetRsValue("Select VODCredit From " & GetOwner & "SO041 Where SysId=" & intCompCode, gcnGi) & "")
            rsSO004G("UpdEn") = strUpdName
            rsSO004G("UpdTime") = GetDTString(strUpdDate)
            rsSO004G("PrePay") = 0
            rsSO004G("Present") = 0
            rsSO004G("SendCount") = 0
            rsSO004G("VODAccountId") = rsResult("AccountUID").Value
            rsSO004G("MVODID") = rsResult("AccountUID").Value
            rsSO004G("VodType") = GetVodType(rsData("ModelCode") & "")
            '#6617 �n�N�³]�ƪ�SO004G�ƨ�@���ܷs���]�ƤW By Kin 2013/10/01
            '#7068 �n�W�[VODType By Kin 2015/08/06
            If Len(OldVodAccountId) > 0 Then
                If Not GetRS(rsOldSO004G, "SELECT * FROM " & GetOwner & "SO004G " & _
                    " WHERE VODACCOUNTID = " & OldVodAccountId, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then GoTo 88
                If rsOldSO004G.RecordCount > 0 Then
                    rsSO004G("PrePay") = rsOldSO004G("PrePay")
                    rsSO004G("Present") = rsOldSO004G("Present")
                    rsSO004G("VODCredit") = rsOldSO004G("VODCredit")
                    rsSO004G("SalePointcode") = rsOldSO004G("SalePointcode")
                    rsSO004G("SalePointname") = rsOldSO004G("SalePointname")
                    rsSO004G("SendCount") = rsOldSO004G("SendCount")
                    rsSO004G("NoPayCredit") = rsOldSO004G("NoPayCredit")
                    rsSO004G("UpdEn") = rsOldSO004G("UpdEn")
                    rsSO004G("UpdTime") = rsOldSO004G("UpdTime")
                    rsSO004G("PrintDetail") = rsOldSO004G("PrintDetail")
                    rsSO004G("StopType") = rsOldSO004G("StopType")
                End If
            End If
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
            
            If Len(OldVodAccountId) > 0 Then
                gcnGi.Execute "Update " & GetOwner & "SO182 " & _
                                    " SET VODAccountID = " & rsResult("AccountUID") & _
                                    " WHERE VODACCOUNTID = " & OldVodAccountId & _
                                    " AND COMPCODE = " & intCompCode
                
                 gcnGi.Execute "Update " & GetOwner & "SO033VOD " & _
                                    " SET VODAccountID = " & rsResult("AccountUID") & _
                                    " WHERE VODACCOUNTID = " & OldVodAccountId & _
                                    " AND COMPCODE = " & intCompCode
                
                gcnGi.Execute "Update " & GetOwner & "SO555B " & _
                                    " SET VODAccountID = " & rsResult("AccountUID") & _
                                    " WHERE VODACCOUNTID = " & OldVodAccountId
                
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
                adUseClient, adOpenKeyset, adLockOptimistic, , , , False) Then GoTo 88
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
        Case "C1"
            
    End Select
    rsData("UpdEn") = strUpdName
    rsData("UpdTime") = GetDTString(strUpdDate)
    rsData.Update
    Call WriteRecordVodProcedure("2.5 ��s��Ƨ��� --> " & strCMD)
    strResult = "1"
    OpenVod = True
88:
    '#6140 A1�BA2�S�w���~�n����V���O By Kin 2011/11/08
    Select Case UCase(strCMD)
        Case "A1"
            blnReturnCmd = InStr(1, strRetErrMsg, strA1ErrCode) > 0
            If (intRunRetunCount > 0) And (strResult = "1") Then    '����Ĥ@����V���O
                blnReturnCmd = True
            End If
            If (Not blnReturnCmd) And (intRunRetunCount = 0) Then
                intRunRetunCount = 99
            End If
        Case "A2"
            blnReturnCmd = InStr(1, strRetErrMsg, strA2ErrCode) > 0
            If (intRunRetunCount > 0) And (strResult = "1") Then    '����Ĥ@����V���O
                blnReturnCmd = True
            End If
            If (Not blnReturnCmd) And (intRunRetunCount = 0) Then
                intRunRetunCount = 99
            End If
        Case Else
            blnReturnCmd = False
            intRunRetunCount = 99
    End Select
        
    If (intRunRetunCount > 0) And (strResult <> "1") Then '��ܤw�g�]�L�@��,�p�G���N�����^���~�F
        blnReturnCmd = False
        intRunRetunCount = 99
    End If
    If (strResult = "1") And (Not blnReturnCmd) Then intRunRetunCount = 99
    If (intRunRetunCount >= 2) Then blnReturnCmd = False: intRunRetunCount = 99 '�ϦV���O���]���F,���n�A�]�U�h�F By Kin 2011/11/08
    
    If IsA6Command Then
        blnReturnCmd = False
        intRunRetunCount = 99
    End If
    
    If (Not blnReturnCmd) And (intRunRetunCount > 0) Then
        intRunRetunCount = 0
        '#6573 A6�R�O�2��,�ҥH�P�_���өR�O�nShowMsg By Kin 2013/08/27
        If IsA6Command Or IsRestCommand Then
'            If (strCMD = UCase("A1")) Then
'                ShowMsg "(1) " & strCMD & "  " & strCmdName, strResult, strFaultReason, strRetErrMsg
'            Else
'                If strResult <> "1" Then
'                    ShowMsg "(1) " & strCMD & "  " & strCmdName, strResult, strFaultReason, strRetErrMsg
'                End If
'            End If
        Else
            If blnShowMsg Then
                ShowMsg "(1) " & strCMD & "  " & strCmdName, strResult, strFaultReason, strRetErrMsg
            End If
        End If
        
    Else
        intRunRetunCount = intRunRetunCount + 1
    End If
    On Error Resume Next
    
    Call CloseRecordset(rsSO004G)
    Call CloseRecordset(rsOldSO004G)
    Exit Function
ChkErr:
  Call ErrSub(FormName, "OpenVod")
End Function
Public Function GetVodType(ByVal ModelCode As String) As Integer
  On Error GoTo ChkErr
    Dim rsVodType As New ADODB.Recordset
    Dim aSQL As String
    GetVodType = 0
    aSQL = "Select Nvl(VodType,0) VodType From " & GetOwner & "CD043 Where CodeNo = " & ModelCode
    If Len(ModelCode) > 0 Then
        If Not GetRS(rsVodType, aSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
        If rsVodType.RecordCount > 0 Then
            GetVodType = Val(rsVodType("VodType") & "")
        End If
    End If
    On Error Resume Next
    Call CloseRecordset(rsVodType)
  Exit Function
ChkErr:
 Call ErrSub(FormName, "GetVodType")
End Function
Public Function GetICCUID(ByVal aSmartCardNo As String, ByVal aCompCode As String) As String
  On Error GoTo ChkErr
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
ChkErr:
  Call ErrSub(FormName, "GetICCUID")
End Function
Private Function TurnNum(ByVal strChgNum) As Long
  On Error GoTo ChkErr
    Dim i As Long
    Dim strNum As String
    For i = 0 To Len(strChgNum) - 1
        If IsNumeric(Mid(strChgNum, i + 1, 1)) Then
            strNum = strNum & Mid(strChgNum, i + 1, 1)
        End If
    Next i
    TurnNum = Val(strNum)
  Exit Function
ChkErr:
    ErrSub FormName, "TurnNum"
End Function




Public Function GetStateData(objCn, ByVal intCompCode As Integer) As String
  On Error Resume Next
    GetStateData = objCn.Execute("SELECT CmdStatus FROM " & strCMowner & "SO555" & _
                                                        " WHERE  SEQNO='" & strCmdSeqNo & "'")(0) & ""
                                                        
                                                
End Function
Public Function GetErrMsg(objCn, ByVal intCompCode As Integer, ByRef strErrMsg As String, ByRef strFaultReason As String) As Boolean
    On Error Resume Next
    'Dim varErrPara As String
    GetErrMsg = False
    Dim rsErr As New ADODB.Recordset
    
'    If Not GetRS(rsErr, "Select RetMsg,SysMsg From " & strCMowner & "SO555" & _
'                                                " Where  SEQNO='" & strCmdSeqNo & "'", objCn, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
'    varErrPara = objCn.Execute("Select RetMsg,SysMsg From " & strCMowner & "SO555" & _
'                                                " Where  SEQNO='" & strCmdSeqNo & "'").GetString(, , ",")
    '#6140 ���դ�OK,�]��ErrMsg���]�t�r�I,�ҥH��Split(,)�|�����D
    Set rsErr = objCn.Execute("Select RetMsg,SysMsg From " & strCMowner & "SO555" & _
                                                " Where  SEQNO='" & strCmdSeqNo & "'")
                                                
    strErrMsg = Replace(rsErr("RetMsg"), Chr(13), "")
    strFaultReason = Replace(rsErr("SysMsg"), Chr(13), "")
    'strErrMsg = Split(Replace(varErrPara, Chr(13), ""), ",")(1)
    'strFaultReason = Split(varErrPara, ",")(0)
    If (strErrMsg = "") And (strFaultReason <> "") Then
        strErrMsg = strFaultReason
        strFaultReason = ""
    End If
    'If strErrMsg <> "" Then strErrMsg = "���~�N�X:" & strErrMsg
    If Trim(strErrMsg) <> "" Or Trim(strFaultReason) <> "" Then
        GetErrMsg = True
    End If
    CloseRecordset rsErr
End Function
Public Sub ShowMsg(strCMD As String, strResult As String, strFaultReason As String, strErrMsg As String)
    On Error GoTo ChkErr
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
ChkErr:
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
  On Error GoTo ChkErr
        IsCMCP = (GetRsValue("SELECT COUNT(*) FROM " & GetOwner & "SO004" & _
                                            " WHERE REFACISNO='" & rs("FaciSno") & "'" & _
                                            " AND COMPCODE=" & rs("CompCode") & _
                                            " AND CUSTID=" & rs("CustID") & _
                                            " AND SERVICETYPE='P'" & _
                                            " AND PRDATE IS NULL") > 0)

  Exit Function
ChkErr:
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
  On Error GoTo ChkErr
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
ChkErr:
  ErrSub FormName, "GetInstAddNo"
End Function
'**********************************************************************************************************
'Corey �g��Function
'�Q�ζǤJ��ADDRNO(�@�w�n�˾��a�})��XSO014.ADDSORT�A�åBSO017C���[�\�a�}���]�t��ADDRSORT��SNO�C�X�ӡA
'�åB�Q��SNO��XSO017B.SwitchID�^�ǡA�γr���j�}�C��SwitchID
Public Function GetSwitchID(ByVal strInstAddrNO As String, ByRef strErrString As String) As String
On Error GoTo ChkErr
    Dim rsTmp As New ADODB.Recordset
    Dim strAddrSort As String
    Dim strSNO As String
    Dim strSwitchID As String
    GetSwitchID = ""
    strSwitchID = ""
    strErrString = ""
    If Not GetRS(rsTmp, "Select AddrSort,FTTxFlag From " & GetOwner & "SO014 Where AddrNO=" & strInstAddrNO, gcnGi, adUseClient, adOpenKeyset, adLockPessimistic, , , , False) Then Exit Function
    If rsTmp.RecordCount = 0 Then
        strErrString = "�䤣������a�}"
        Exit Function '�䤣��������a�} Addrsort
    End If
'    If Val(rsTmp("FTTxFlag") & "") = 0 Then
'        strErrString = "�Ӧa�}���b FTTX �[�\�d��"
'        Exit Function
'    End If

    strAddrSort = Mid(rsTmp("AddrSort") & "", 1, GetFldMaxLen("SO017C", "AddrSortA"))
    If Not GetRS(rsTmp, "Select * From " & GetOwner & "SO017C Where AddrSortA <= '" & strAddrSort & "' and AddrSortB >= '" & strAddrSort & "'", gcnGi, adUseClient, adOpenKeyset, adLockOptimistic, , , , False) Then Exit Function
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
ChkErr:
    ErrSub FormName, "GetSwitchID"
End Function
Public Function GetFldMaxLen(ByVal uTableName As String, ByVal uColumn As String) As Double

  On Error GoTo ChkErr
  
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
    
    If Not GetRS(rsColumn, strSQL, , , , , , , , False) Then Exit Function
    
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
ChkErr:
    ErrSub FormName, "GetFldMaxLen"
End Function
'*******************************************************************************************
Public Function chkUseFlag(ByVal strFaciSNo As String, ByVal intComp As Integer) As Boolean
  On Error GoTo ChkErr:
    Dim rs As New ADODB.Recordset
    chkUseFlag = False
    'If blnUseFttx Then chkUseFlag = True: Exit Function
    
    If Not GetRS(rs, "Select Useflag From " & GetOwner & "SO306 Where HFCMAC='" & strFaciSNo & "'") Then Exit Function
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
ChkErr:
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
  On Error GoTo ChkErr
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
ChkErr:
    ErrSub FormName, "SetRsMemory"
End Sub

Public Function GetVendCode(ByVal intComp As Integer) As String
  On Error Resume Next
    GetVendCode = gcnGi.Execute("Select VendorCode From " & GetOwner & "CM006 Where " & _
                                "CompCode=" & intComp & " And StopFlag=0")(0)
                        

End Function
Private Function UpdSO004F(ByRef rs As Recordset, ByVal strUpdTime As String) As Boolean
  On Error GoTo ChkErr
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
ChkErr:
  Call ErrSub(FormName, "UpdSO004F")
End Function
