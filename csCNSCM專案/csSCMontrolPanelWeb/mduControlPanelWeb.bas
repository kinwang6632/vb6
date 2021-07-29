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
Public strFaciSNo As String
Public strFaciCode As String
Public strResvTime As String
Public strBaudRate  As String
Public strVendor As String
Public strPassword As String
Public blnSendKey As Boolean
Public strReason As String
Public blnShowFaci As Boolean
Public strSNO As String
Public strSeqNo As String
Public rsUpdSo004 As New ADODB.Recordset
Public strDynIpCount As String
Public strFixIPCount  As String
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Function GetModeTypeValue(ByVal strValue As String) As String
    On Error Resume Next
    Dim strReturn As String
        Select Case strValue
            Case "A1"
                strReturn = "開機"
            Case "A2"
                strReturn = "關機"
            Case "A3"
                strReturn = "停機"
            Case "A4"
                strReturn = "重設(Reset)"
            
            Case "B1"
                strReturn = "換CM"
            Case "B2"
                strReturn = "變更CM速率"
            Case "B3"
                strReturn = "Ping CM"
            Case "B4"
                strReturn = "釋放動態IP"
            Case "B5"
                strReturn = "ACS帳號密碼相關功能"
            
            Case "C1"
                strReturn = "CM 狀態查詢"
            Case "C2"
                strReturn = "查詢 HFC 服務類別"
            Case "C3"
                strReturn = "清除EMM"
            Case "C4"
                strReturn = "CM 連線紀錄查詢"
            Case "C5"
                strReturn = "PCIP連線記錄查詢"
        End Select
        
        GetModeTypeValue = strReturn
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


'裝機[10]
Public Function InstCM(ByRef rsData As ADODB.Recordset, ByVal strID_Code As String, _
                                        Optional ByVal strCustName As String = Empty, _
                                        Optional ByVal strTel As String = Empty, _
                                        Optional ByVal strCustAddress As String = Empty, _
                                        Optional ByVal strClassID As String = Empty, _
                                        Optional ByVal strVendCode As String = Empty, _
                                        Optional ByVal strMTAMAC As String = Empty, _
                                        Optional ByVal strSNO As String = Empty, _
                                        Optional ByVal strMediaBillNo As String = Empty, _
                                        Optional ByVal strProcessingDate As String = Empty, _
                                        Optional ByRef rsResult As Object = Nothing, _
                                        Optional ByRef strResult As String = Empty, _
                                        Optional ByRef strFaultReason As String = Empty, _
                                        Optional ByRef strErrMsg As String = Empty) As Boolean
                                        
  On Error GoTo ChkErr
    Dim varBookMark As Variant
    Dim strScmData As String
    Dim varPara As Variant
    
    Dim rsCM006 As New ADODB.Recordset
    'Dim objScmCommand As Object
    'Set objScmCommand = CreateObject("csSCMcontrol.clsSCMcommand")
    Dim objScmCommand As New csSCMcontrol.clsSCMcommand
    '#4262 原本是用rsData.BookMark,但會使某些RecordSet找不到指標 By Kin 2008/12/03
    varBookMark = rsData.AbsolutePosition
    InstCM = False
    intCompCode = rsData("CompCode") & ""
    
    If blnExeCommand Then
        If Not objScmCommand.SCM_Command(gcnGi.ConnectionString, intCompCode, GetOwner, "10", GaryGi(0), rsData("CustID"), strCustName, _
                                        strTel, strCustAddress, , strClassID, strVendCode, rsData("ID") & "", _
                                        rsData("DeclarantName") & "", rsData("BirthDay") & "", rsData("FaciSNO") & "", _
                                        strMTAMAC, , , , strSNO, strMediaBillNo, strProcessingDate, Empty, Empty, _
                                        GetPrbCtl, rsResult, strResult, strFaultReason, strErrMsg) Then GoTo 66
    End If
    rsData.AbsolutePosition = varBookMark
    If strProcessingDate <> Empty And blnExeCommand Then
        strResult = "1"
        InstCM = True
        GoTo 66
    End If
    If Not blnExeCommand Then
        GetScmData gcnGi, rsResult
    End If
    If Not rsResult Is Nothing Then
        rsData("ServiceType") = rsResult("ServiceType")
        rsData("DialAccount") = rsResult("AccountId")
        rsData("DialPassWord") = rsResult("AccountPassword")
        
        '**********************************************************************************
        '#3772 增加回傳資料,要Update SO004.VendorCode與SO004.VendorName By Kin 2008/02/15
        If rsResult("SchemeCode2") & "" <> "" Then
            rsData("VendorCode") = rsResult("SchemeCode2")
            Set rsCM006 = gcnGi.Execute("Select VendorName From " & GetOwner & "CM006 Where " & _
                                      "VendorCode=" & rsResult("SchemeCode2") & _
                                      " And CompCode=" & intCompCode)
            If Not rsCM006.EOF Then
                If Not IsNull(rsCM006("VendorName")) Then rsData("VendorName") = rsCM006("VendorName") & ""
            End If
        End If
        '***********************************************************************************
        
        '#3778裝機成功改填開機時間不填裝機時間 By Kin 2008/02/22
        'rsData("InstDate") = Format(RightNow, "YYYY/MM/DD HH:MM:SS")
        '#6812 判斷設備正常再更新 By Kin 2014/06/18
        '判斷正常設又要拿掉 For 銘誠 By Kin 2014/10/16
        rsData("CMOpenDate") = Format(RightNow, "YYYY/MM/DD HH:MM:SS")
        rsData("UpdEn") = GaryGi(1)
        rsData("UpdTime") = GetDTString(RightNow)
        rsData.Update
        strResult = "1"
        InstCM = True
    Else
        If blnExeCommand Then
            strErrMsg = "處理成功，但無任何回傳值!!"
        Else
            InstCM = True
        End If
    End If
66:
    If blnShowMsg Then
        ShowMsg "(1) 建立CM服務帳號", strResult, strFaultReason, strErrMsg
    End If
    On Error Resume Next
    Call CloseRecordset(rsCM006)
    Exit Function
ChkErr:
    ErrSub FormName, "InstCM"
    On Error Resume Next
    Call CloseRecordset(rsCM006)
End Function
'#4276 重置PPPOE密碼 By Kin 2008/12/25
Public Function ResetPPPOEPassowrd(ByRef rsData As ADODB.Recordset, ByVal strID_Code As String, _
                                                        Optional ByVal strSNO As String = "", _
                                                        Optional ByVal strMediaBillNo As String = "", _
                                                        Optional ByVal strProcessingDate As String = "", _
                                                        Optional ByRef strResult As String, _
                                                        Optional ByRef strFaultReason As String, _
                                                        Optional ByRef strErrMsg As String) As Boolean
On Error GoTo ChkErr
    Dim varBookMark As Variant
    Dim objScmCommand As New csSCMcontrol.clsSCMcommand
    ResetPPPOEPassowrd = False
    varBookMark = rsData.AbsolutePosition
    intCompCode = rsData("CompCode") & ""
    
    If blnExeCommand Then
        If Not objScmCommand.SCM_Command(gcnGi.ConnectionString, str(intCompCode), GetOwner, strID_Code, _
                                            GaryGi(0), rsData("Custid") & "", , , , rsData("DialAccount") & "", , , , , , rsData("FaciSNo") & "", , , , _
                                            , strSNO, strMediaBillNo, strProcessingDate, Empty, Empty, _
                                            GetPrbCtl, , strResult, strFaultReason, strErrMsg) Then GoTo 66
    End If
    If strProcessingDate <> Empty And blnExeCommand Then
        strResult = "1"
        ResetPPPOEPassowrd = True
        GoTo 66
    End If
    strResult = "1"
    ResetPPPOEPassowrd = True
66:
    If blnShowMsg Then
        ShowMsg "(2) 重設PPPOE密碼 ", strResult, strFaultReason, strErrMsg
    End If
    Exit Function
ChkErr:
    ErrSub FormName, "ResetPPPOEPassowrd"

End Function
'#4276 取消固定IP By Kin 2008/12/24
Public Function ReleaseFixIP(ByRef rsData As ADODB.Recordset, ByVal strID_Code As String, _
                             Optional ByVal strSNO As String = "", _
                             Optional ByVal strMediaBillNo As String, _
                             Optional ByVal strProcessingDate As String = "", _
                             Optional ByRef rsResult As Object = Nothing, _
                             Optional ByRef strResult As String, _
                             Optional ByRef strFaultReason As String, _
                             Optional ByRef strErrMsg As String = "") As Boolean
    On Error GoTo ChkErr
    Dim strCmdName As String
    Dim varBookMark As Variant
    Dim strParamData As String
    Dim objScmCommand As New csSCMcontrol.clsSCMcommand
    Dim rsSO004F As New ADODB.Recordset
    Dim strSO004FSQL As String
    Dim i As Integer
    Dim strNowTime As String
    strNowTime = RightNow
    varBookMark = rsData.AbsolutePosition
    ReleaseFixIP = False
    strCmdName = "取消申請固定IP"
    intCompCode = rsData("CompCode") & ""
    strSO004FSQL = "Select * From " & GetOwner & "SO004F" & _
                 " Where DialAccount='" & rsData("DialAccount") & "'" & _
                 " And ExtendType=1 And PrDate is Null"
    If Not GetRS(rsSO004F, strSO004FSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
    If Not rsSO004F.EOF Then
        rsSO004F.MoveFirst
        Do While Not rsSO004F.EOF
            If i = 0 Then
                strParamData = "FixIPAccountId=" & rsSO004F("ExtendAccount") & ""
            End If
            strParamData = strParamData & Chr(13) & "FixIP" & i + 1 & "=" & rsSO004F("FixIPAddress")
'            If strProcessingDate = Empty Then
'                rsSO004F("PrDate") = Format(strNowTime, "YYYY/MM/DD HH:MM:SS")
'                rsSO004F("UpdEn") = GaryGi(1)
'                rsSO004F("UpdTime") = GetDTString(strNowTime)
'                rsSO004F.Update
'            End If
            rsSO004F.MoveNext
            i = i + 1
        Loop
    Else
        strFaultReason = "找不到任何帳號資料"
        GoTo 66
    End If
    If blnExeCommand Then
        If Not objScmCommand.SCM_Command(gcnGi.ConnectionString, str(intCompCode), GetOwner, _
                                         strID_Code, GaryGi(0), _
                                        rsData("CustId") & "", , , , , , , , , , rsData("FaciSNO") & "", _
                                        , , , , strSNO, strMediaBillNo, strProcessingDate, , , _
                                        GetPrbCtl, _
                                        , strResult, strFaultReason, strErrMsg, strParamData) Then GoTo 66



    End If
    If strProcessingDate <> Empty And blnExeCommand Then
        strResult = "1"
        ReleaseFixIP = True
        GoTo 66
    End If
    '*********************************************************************
    '指令成功要將SO004F Upd By Kin 2008/12/29
    rsSO004F.MoveFirst
    Do While Not rsSO004F.EOF
        rsSO004F("PrDate") = Format(strNowTime, "YYYY/MM/DD HH:MM:SS")
        rsSO004F("UpdEn") = GaryGi(1)
        rsSO004F("UpdTime") = GetDTString(strNowTime)
        rsSO004F.Update
        rsSO004F.MoveNext
    Loop
    strResult = "1"
    ReleaseFixIP = True
    '********************************************************************
66:
    If blnShowMsg Then
        ShowMsg "(2)  取消申請固定IP", strResult, strFaultReason, strErrMsg
    End If
    Exit Function
ChkErr:
    ErrSub FormName, "ReleaseFixIP"
End Function
'#4276 申請固IP By Kin 2008/12/24
Public Function RequireFixIP(ByRef rsData As ADODB.Recordset, ByVal strID_Code As String, _
                             Optional ByVal strSNO As String = "", _
                             Optional ByVal strMediaBillNo As String, _
                             Optional ByVal strProcessingDate As String = "", _
                             Optional ByVal strDynIpNum As String = "", _
                             Optional ByVal strFixIPNum As String = "", _
                             Optional ByRef rsResult As Object = Nothing, _
                             Optional ByRef strResult As String, _
                             Optional ByRef strFaultReason As String, _
                             Optional ByRef strErrMsg As String = "") As Boolean
    On Error GoTo ChkErr
    Dim strCmdName As String
    Dim varBookMark As Variant
    Dim objScmCommand As New csSCMcontrol.clsSCMcommand
    Dim rsSO004F As New ADODB.Recordset
    Dim i As Integer
    Dim strNowTime As String
    strNowTime = RightNow
    varBookMark = rsData.AbsolutePosition
    RequireFixIP = False
    strCmdName = "申請固定IP"
    intCompCode = rsData("CompCode") & ""
    If blnExeCommand Then
        If Not objScmCommand.SCM_Command(gcnGi.ConnectionString, str(intCompCode), GetOwner, _
                                         strID_Code, GaryGi(0), _
                                        rsData("CustId") & "", , , , rsData("DialAccount") & "", , , , , , rsData("FaciSNO") & "", _
                                        , , , , strSNO, strMediaBillNo, strProcessingDate, , , GetPrbCtl, _
                                        rsResult, strResult, strFaultReason, strErrMsg, , strDynIpNum, strFixIPCount) Then GoTo 66



    End If
    If strProcessingDate <> Empty And blnExeCommand Then
        strResult = "1"
        RequireFixIP = True
        GoTo 66
    End If
    If Not blnExeCommand Then
        GetScmData gcnGi, rsResult
    End If
    rsData.AbsolutePosition = varBookMark
    If Not rsResult Is Nothing Then
        If Not GetRS(rsSO004F, "Select * From " & GetOwner & "SO004F Where 1=0", gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
            For i = 0 To rsResult.Fields.Count - 5
                With rsSO004F
                    .AddNew
                    .Fields("SeqNo") = Get_SO004F_Seq_No(gcnGi)
                    .Fields("DialAccount") = rsData("DialAccount") & ""
                    .Fields("ExtendAccount") = rsResult("FixIPAccountId") & ""
                    .Fields("ExtendPassword") = rsResult("FixIPAccountPassword") & ""
                    .Fields("FixIPAddress") = rsResult(i + 4).Value & ""
                    .Fields("UpdEn") = GaryGi(1)
                    .Fields("UpdTime") = GetDTString(strNowTime)
                    .Update
                End With
            Next i
        strResult = "1"
        RequireFixIP = True
    Else
        If blnExeCommand Then
            strErrMsg = "處理成功，但無任何回傳值!!"
        Else
            RequireFixIP = True
        End If
    End If
66:
    If blnShowMsg Then
        ShowMsg "(2)  申請固定IP", strResult, strFaultReason, strErrMsg
    End If
    Exit Function
ChkErr:
    ErrSub FormName, "RequireFixIP"
End Function

'#4276 啟用PPPOE帳號 By Kin 2008/12/24
Public Function OpenPPPOEAccount(ByRef rsData As ADODB.Recordset, ByVal strID_Code As String, _
                                 Optional ByVal strCustName As String = Empty, _
                                 Optional ByVal strTel As String = Empty, _
                                 Optional ByVal strCustAddress As String = Empty, _
                                 Optional ByVal strProcessingDate As String = Empty, _
                                 Optional ByVal strSNO As String = Empty, _
                                 Optional ByVal strMediaBillNo As String = Empty, _
                                 Optional ByRef strResult As String, _
                                 Optional ByRef strFaultReason As String, _
                                 Optional ByRef strErrMsg As String = "") As Boolean
  On Error GoTo ChkErr
    Dim varBookMark As Variant
    Dim strScmData As String
    Dim varPara As Variant
    Dim rsCM006 As New ADODB.Recordset
    Dim strUpdTime As String
    Dim objScmCommand As New csSCMcontrol.clsSCMcommand
    varBookMark = rsData.AbsolutePosition
    OpenPPPOEAccount = False
    intCompCode = rsData("CompCode") & ""
    
    If blnExeCommand Then
        If Not objScmCommand.SCM_Command(gcnGi.ConnectionString, intCompCode, GetOwner, "10", GaryGi(0), rsData("CustID"), strCustName, _
                                        strTel, strCustAddress, rsData("DialAccount") & "", rsData("CMBaudNo") & "", rsData("VendorCode") & "", rsData("ID") & "", _
                                        rsData("DeclarantName") & "", rsData("BirthDay") & "", rsData("FaciSNO") & "", _
                                        , , , rsData("DialPassword") & "", strSNO, strMediaBillNo, strProcessingDate, Empty, _
                                        Empty, GetPrbCtl, , strResult, strFaultReason, strErrMsg) Then GoTo 66
    End If
    rsData.AbsolutePosition = varBookMark
    If strProcessingDate <> Empty And blnExeCommand Then
        strResult = "1"
        OpenPPPOEAccount = True
        GoTo 66
    End If
    strUpdTime = RightNow
    '#6812 判斷設備正常再更新 By Kin 2014/06/18
    '判斷正常設又要拿掉 For 銘誠 By Kin 2014/10/16
    rsData("CMOpenDate") = Format(strUpdTime, "YYYY/MM/DD HH:MM:SS")
    rsData("UpdEn") = GaryGi(1)
    rsData("UpdTime") = GetDTString(strUpdTime)
    rsData.Update
    strResult = "1"
    OpenPPPOEAccount = True
66:
    If blnShowMsg Then
        ShowMsg "(1) 啟用PPPOE帳號", strResult, strFaultReason, strErrMsg
    End If
    Exit Function

ChkErr:
  ErrSub FormName, "OpenPPPOEAccount"

End Function

'#4276 Fttx裝機退單 By Kin 2008/12/23
Public Function ReleasePPPOEAccount(ByRef rsData As ADODB.Recordset, ByVal strID_Code As String, _
                                    Optional ByVal strHFCMAC As String = "", _
                                    Optional ByVal strSNO As String = "", _
                                    Optional ByVal strMediaBillNo As String, _
                                    Optional ByVal strProcessingDate As String = "", _
                                    Optional ByRef strResult As String, _
                                    Optional ByRef strFaultReason As String, _
                                    Optional ByRef strErrMsg As String = "") As Boolean
  On Error GoTo ChkErr
    Dim strCmdName As String
    Dim varBookMark As Variant
    Dim objScmCommand As New csSCMcontrol.clsSCMcommand
    varBookMark = rsData.AbsolutePosition
    ReleasePPPOEAccount = False
    strCmdName = "取消預約FTTX裝機"
    intCompCode = rsData("CompCode") & ""
    If blnExeCommand Then
        If Not objScmCommand.SCM_Command(gcnGi.ConnectionString, str(intCompCode), GetOwner, _
                                         strID_Code, GaryGi(0), _
                                        rsData("CustId") & "", , , , rsData("DialAccount") & "", , , , , , strHFCMAC, _
                                        , , , , strSNO, strMediaBillNo, strProcessingDate, , , GetPrbCtl, _
                                        , strResult, strFaultReason, strErrMsg) Then GoTo 66



    End If
    If strProcessingDate <> Empty And blnExeCommand Then
        strResult = "1"
        ReleasePPPOEAccount = True
        GoTo 66
    End If
    rsData.AbsolutePosition = varBookMark
    strResult = "1"
    ReleasePPPOEAccount = True
  
66:
    If blnShowMsg Then
        ShowMsg "(1)" & strCmdName, strResult, strFaultReason, strErrMsg
    End If
    Exit Function
ChkErr:
    ErrSub FormName, "ReleasePPPOEAccount"
End Function


'#4276 預約FTTX裝機 By Kin 2008/12/23
Public Function RequirePPPOEAccount(ByRef rsData As ADODB.Recordset, ByVal strID_Code As String, _
                                    Optional ByVal strSNO As String = "", _
                                    Optional ByVal strMediaBillNo As String, _
                                    Optional ByVal strProcessingDate As String = "", _
                                    Optional ByVal strDynIpNum As String = "", _
                                    Optional ByVal strFixIPNum As String = "", _
                                    Optional ByRef rsResult As Object = Nothing, _
                                    Optional ByRef strResult As String = "", _
                                    Optional ByRef strFaultReason As String = "", _
                                    Optional ByRef strErrMsg As String = "") As Boolean
                                    
  On Error GoTo ChkErr
    Dim strCmdName As String
    Dim varBookMark As Variant
    Dim strParamData As String
    Dim aryParamData As Variant
    Dim strNowTime As String
    Dim i As Integer
    Dim rsSO004F As New ADODB.Recordset
    Dim InstAddrNo As String
    Dim strSwitchErr As String
    Dim objScmCommand As New csSCMcontrol.clsSCMcommand
    strNowTime = RightNow
    strSwitchErr = ""
    varBookMark = rsData.AbsolutePosition
    Select Case strID_Code
        Case "18"
            strCmdName = "預約FTTX裝機"
    End Select
    RequirePPPOEAccount = False
    intCompCode = rsData("CompCode") & ""
    aryParamData = Split(GetSwitchID(GetInstAddNo(rsData), strSwitchErr), ",")
    
    For i = LBound(aryParamData) To UBound(aryParamData)
        strParamData = IIf(strParamData = Empty, "", strParamData & Chr(13))
        strParamData = strParamData & "SWNo" & i + 1 & "=" & aryParamData(i)
    Next
    If strParamData = "" Then
        strResult = ""
        strErrMsg = strSwitchErr
        GoTo 66
    End If
    If blnExeCommand Then
        If Not objScmCommand.SCM_Command(gcnGi.ConnectionString, str(intCompCode), GetOwner, _
                                         strID_Code, GaryGi(0), _
                                        rsData("CustId") & "", , , , , rsData("CMBaudNo") & "", rsData("VendorCode") & "", , , , , _
                                        , , , , strSNO, strMediaBillNo, strProcessingDate, , , GetPrbCtl, rsResult _
                                        , strResult, strFaultReason, strErrMsg, strParamData, strDynIpNum, strFixIPCount) Then GoTo 66



    End If
    rsData.AbsolutePosition = varBookMark
    If strProcessingDate <> Empty And blnExeCommand Then
        strResult = "1"
        RequirePPPOEAccount = True
        GoTo 66
    End If
    If Not blnExeCommand Then
        GetScmData gcnGi, rsResult
    End If
    If Not rsResult Is Nothing Then
        If Not GetRS(rsSO004F, "Select * From " & GetOwner & "SO004F Where 1=0", gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
        With rsData
            .Fields("DialAccount") = rsResult("AccountId") & ""
            .Fields("DialPassWord") = rsResult("AccountPassword") & ""
            If Trim(rsResult("SchemeCode2")) <> "空" And Trim(rsResult("SchemeCode2")) <> "" Then
                .Fields("VendorCode") = rsResult("SchemeCode2") & ""
                .Fields("VendorName") = GetRsValue("Select VendorName From " & GetOwner & "CM006 Where VendorCode='" & rsResult("SchemeCode2") & "'" & _
                                                    " And CompCode=" & rsData("CompCode"), gcnGi) & ""

            End If
            .Fields("UpdEn") = GaryGi(1)
            .Fields("UpdTime") = GetDTString(strNowTime)
            .Update
        End With
        For i = 0 To rsResult.Fields.Count - 8
            rsSO004F.AddNew
            rsSO004F("SeqNo") = Get_SO004F_Seq_No(gcnGi)
            rsSO004F("DialAccount") = rsData("DialAccount") & ""
            rsSO004F("ExtendAccount") = rsResult("FixIPAccountId") & ""
            rsSO004F("ExtendPassword") = rsResult("FixIPAccountPassword") & ""
            rsSO004F("FixIpAddress") = rsResult("FixIP" & i + 1) & ""
            rsSO004F("UpdEn") = GaryGi(1)
            rsSO004F("UpdTime") = GetDTString(strNowTime)
            rsSO004F.Update
        Next i

        
        strResult = "1"
        RequirePPPOEAccount = True
    Else
        strResult = "1"
        RequirePPPOEAccount = True
    End If
66:
    If blnShowMsg Then
        ShowMsg "(1)" & strCmdName, strResult, strFaultReason, strErrMsg
    End If
    On Error Resume Next
    Call CloseRecordset(rsSO004F)
    Exit Function
ChkErr:
  ErrSub FormName, "RequirePPPOEAccount"
End Function
'開機、關機、停機、復機、拆機、維修換拆[11、12、13、14、15、16]
Public Function OpenCM(ByRef rsData As ADODB.Recordset, ByVal strID_Code As String, _
                                                        Optional ByVal strHFCMAC As String = "", _
                                                        Optional ByVal strMATMac As String = "", _
                                                        Optional ByVal strSNO As String = "", _
                                                        Optional ByVal strMediaBillNo As String, _
                                                        Optional ByVal strProcessingDate As String, _
                                                        Optional ByVal strReasonCode As String, _
                                                        Optional ByVal strReasonName As String, _
                                                        Optional ByRef strResult As String, _
                                                        Optional ByRef strFaultReason As String, _
                                                        Optional ByRef strErrMsg As String) As Boolean
  On Error GoTo ChkErr
    Dim strCmdName As String
    Dim varBookMark As Variant
    Dim rsSO004F As New ADODB.Recordset
    Dim strSchSO004F As String
    Dim strNowDate As String
    'Dim objScmCommand As Object
    'Set objScmCommand = CreateObject("csSCMcontrol.clsSCMcommand")
    Dim objScmCommand As New csSCMcontrol.clsSCMcommand
    Dim rsVirtual As New ADODB.Recordset
    'Set rsVirtual = Nothing
    '#4262 原本是用rsData.BookMark,但會使某些RecordSet找不到指標 By Kin 2008/12/03
    strNowDate = RightNow
    varBookMark = rsData.AbsolutePosition
    Select Case strID_Code
        Case "11"
            strCmdName = "開機"
        Case "12"
            strCmdName = "關機"
        Case "13"
            strCmdName = "停機"
        Case "14"
            strCmdName = "復機"
        Case "15"
            strCmdName = "終止服務帳號"
        Case "16"
            strCmdName = "設備更換"
        Case "17"
            strCmdName = "重置設備"
        Case "29"
            strCmdName = "FTTX同區移入"
        Case "33"
            strCmdName = "Release Port"
            
    End Select
    OpenCM = False
    
    intCompCode = rsData("CompCode") & ""
    Call WriteRecordVodProcedure("3.進入命令模式")
    '#3778 增加異動原因 By Kin 2008/02/26
    '#5165 指令11與12 要增加判斷是否有傳回ScmData值,如果有要upd SO004F By Kin 2009/07/03
    If blnExeCommand Then
        If Not objScmCommand.SCM_Command(gcnGi.ConnectionString, str(intCompCode), GetOwner, _
                                         strID_Code, GaryGi(0), _
                                        rsData("CustId") & "", , , , _
                                        IIf(strID_Code = 17, Empty, rsData("DialAccount") & ""), , , , , , _
                                        strHFCMAC, strMATMac, _
                                        , , , strSNO, strMediaBillNo, strProcessingDate, strReasonCode, strReasonName, _
                                        GetPrbCtl, rsVirtual _
                                        , strResult, strFaultReason, strErrMsg) Then GoTo 66
    End If
    Call WriteRecordVodProcedure("7.送出SO314成功 ")
    If strProcessingDate <> Empty And blnExeCommand Then
        strResult = "1"
        OpenCM = True
        GoTo 66
    End If
    
    rsData.AbsolutePosition = varBookMark
    Call WriteRecordVodProcedure("8.準備更新SO004")
    Select Case strID_Code
        Case "11"   '開機
            '#5165 如果有回傳ScmData要進行UPD SO004F By Kin 2009/07/03
            If Not rsVirtual Is Nothing Then
                If rsVirtual.State = adStateOpen Then
                    If Not UpdSO004F(rsVirtual, strNowDate) Then GoTo 66
                End If
            End If
           '#6812 判斷設備正常再更新 By Kin 2014/06/18
            '判斷正常設又要拿掉 For 銘誠 By Kin 2014/10/16
            rsData("CMOpenDate") = Format(strNowDate, "YYYY/MM/DD HH:MM:SS")
            rsData("UpdEn") = GaryGi(1)
            rsData("UpdTime") = GetDTString(strNowDate)
            rsData.Update
            
            Call WriteRecordVodProcedure("9.更新SO004成功")
        Case "12"   '關機
            '#5165 如果有回傳ScmData要進行UPD SO004F By Kin 2009/07/03
            If Not rsVirtual Is Nothing Then
                If rsVirtual.State = adStateOpen Then
                    If Not UpdSO004F(rsVirtual, strNowDate) Then GoTo 66
                End If
            End If
            '#6812 判斷設備正常再更新 By Kin 2014/06/18
            '判斷正常設又要拿掉 For 銘誠 By Kin 2014/10/16
            rsData("CMCloseDate") = Format(strNowDate, "YYYY/MM/DD HH:MM:SS")
            rsData("UpdEn") = GaryGi(1)
            rsData("UpdTime") = GetDTString(strNowDate)
            rsData.Update
            
        Case "13"   '停機
             '#5165 如果有回傳ScmData要進行UPD SO004F By Kin 2009/07/03
            If Len(rsData("InstDate") & "") > 0 And Len(rsData("PRDate") & "") = 0 Then
                rsData("FaciStatusCode") = 3
                '#3778 增加填入關機時間 By Kin 2008/02/22
                rsData("CMCloseDate") = Format(RightNow, "YYYY/MM/DD HH:MM:SS")
                rsData.Update
            End If
        Case "14"   '復機
            '#6812 判斷設備正常再更新 By Kin 2014/06/18
            '判斷正常設又要拿掉 For 銘誠 By Kin 2014/10/16
            rsData("FaciStatusCode") = 4
            '#3778 增加填入開機時間 By Kin 2008/02/22
            rsData("CMOpenDate") = Format(RightNow, "YYYY/MM/DD HH:MM:SS")
            rsData.Update
            
        Case "15"   '拆機
            '#3859 原本填PRDate日期,改填CMCloseDate By Kin 2008/05/07
            '#4276 要找是否有SO004F的資料,如果有要進行UPD(不管是不是FTTX) By Kin 2009/01/08
            '#4367 如果安裝日期是空值不進行任何更新的動作 By Kin 2009/02/16
            If rsData("InstDate") & "" <> "" Then
                If rsData("DialAccount") & "" <> "" Then
                    strSchSO004F = "Select * From " & GetOwner & "SO004F Where " & _
                                    " Prdate is null And SO004F.DialAccount='" & rsData("DialAccount") & "" & "'"
                    If Not GetRS(rsSO004F, strSchSO004F, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then GoTo 66
                    
                    Do While Not rsSO004F.EOF
                        rsSO004F("Prdate") = Format(strNowDate, "YYYY/MM/DD HH:MM:SS")
                        rsSO004F("UpdEn") = GaryGi(1)
                        rsSO004F("UpdTime") = GetDTString(strNowDate)
                        rsSO004F.Update
                        rsSO004F.MoveNext
                    Loop
                End If
                '#6812 判斷設備正常再更新 By Kin 2014/06/18
                '判斷正常設又要拿掉 For 銘誠 By Kin 2014/10/16
                rsData("CMCloseDate") = Format(strNowDate, "YYYY/MM/DD HH:MM:SS")
                rsData("UpdEn") = GaryGi(1)
                rsData("UpdTime") = GetDTString(strNowDate)
                rsData.Update
            End If
        Case "16"   '設備更換
            Dim strRightNow As String
            strRightNow = RightNow
            '#6812 判斷設備正常再更新 By Kin 2014/06/18
            '判斷正常設又要拿掉 For 銘誠 By Kin 2014/10/16
            
            rsData("CMCloseDate") = Format(strRightNow, "YYYY/MM/DD HH:MM:SS")
            rsData.Update
            
            If Not rsUpdSo004 Is Nothing Then
                If rsUpdSo004.State = adStateOpen Then
                    rsUpdSo004("CMOpenDate") = Format(strRightNow, "YYYY/MM/DD HH:MM:SS")
                    rsUpdSo004.Update
                End If
            End If
'            If strSeqNo <> Empty Then
'                gcnGi.Execute "Update " & GetOwner & "SO004 " & _
'                              "Set CMOpenDate=To_Date('" & Format(strRightNow, "YYYYMMDDHHMMSS") & "','YYYY/MM/DD HH24:MI:SS')" & _
'                              " Where SeqNo='" & strSeqNo & "'" & _
'                              " And CompCode=" & rsData("CompCode")
'
'            End If
        Case "17"   '重置設置備
            '#6812 判斷設備正常再更新 By Kin 2014/06/18
            '判斷正常設又要拿掉 For 銘誠 By Kin 2014/10/16
            '#7578 Cancel to update the fileld of cmopendate by Kin 2017/08/03
            'rsData("CMOpenDate") = Format(RightNow, "YYYY/MM/DD HH:MM:SS")
            rsData("UpdEn") = GaryGi(1)
            rsData("UpdTime") = GetDTString(RightNow)
            rsData.Update
            
        Case "29"
            '#6812 判斷設備正常再更新 By Kin 2014/06/18
            '判斷正常設又要拿掉 For 銘誠 By Kin 2014/10/16
            rsData("CMOpenDate") = Format(RightNow, "YYYY/MM/DD HH:MM:SS")
            rsData("UpdEn") = GaryGi(1)
            rsData("UpdTime") = GetDTString(RightNow)
            rsData.Update

    End Select
    OpenCM = True
    strResult = "1"
    GoTo 66
  
66:
    If blnShowMsg Then
        
        ShowMsg "(1) " & IIf(strID_Code = "29", "", "") & strCmdName, strResult, strFaultReason, strErrMsg
    End If
    On Error Resume Next
    Call CloseRecordset(rsSO004F)
    Call CloseRecordset(rsVirtual)
    Exit Function
ChkErr:
    ErrSub FormName, "OpenCM"
End Function
Public Function GetPrbCtl() As Object
  On Error GoTo ChkErr
    Dim objCtl As Object
    If blnShowMsg Then
        If frmSO1623A Is Nothing Then
            Set objCtl = Nothing
        Else
            Set objCtl = frmSO1623A.pbr1
        End If
    Else
        Set objCtl = Nothing
    End If
    Set GetPrbCtl = objCtl
    Exit Function
ChkErr:
  Call ErrSub(FormName, "GetPrbCtl")
End Function

'變更基本資料[20]
Public Function ChangeData(ByRef rsData As ADODB.Recordset, ByVal strID_Code As String, _
                                        Optional ByVal strCustName As String = Empty, _
                                        Optional ByVal strTel As String = Empty, _
                                        Optional ByVal strCustAddress As String = Empty, _
                                        Optional ByVal strClassID As String = Empty, _
                                        Optional ByVal strSNO As String = Empty, _
                                        Optional ByVal strMediaBillNo As String = Empty, _
                                        Optional ByVal strProcessingDate As String = Empty, _
                                        Optional ByRef strResult As String = Empty, _
                                        Optional ByRef strFaultReason As String = Empty, _
                                        Optional ByRef strErrMsg As String = Empty) As Boolean
                                        
  On Error GoTo ChkErr
    ChangeData = False
    'Dim objScmCommand As Object
    'Set objScmCommand = CreateObject("csSCMcontrol.clsSCMcommand")
    Dim objScmCommand As New csSCMcontrol.clsSCMcommand
    intCompCode = rsData("CompCode") & ""
    If blnExeCommand Then
        If Not objScmCommand.SCM_Command(gcnGi.ConnectionString, str(intCompCode), GetOwner, _
                                            strID_Code, GaryGi(0), rsData("CustID"), strCustName, _
                                            strTel, strCustAddress, rsData("DialAccount") & "", , , _
                                            , , , rsData("FaciSNo") & "", , , , , strSNO, strMediaBillNo, _
                                            strProcessingDate, , , GetPrbCtl, _
                                            , strResult, strFaultReason, strErrMsg) Then GoTo 66
    End If
    strResult = "1"
    ChangeData = True
66:
    If blnShowMsg Then
        ShowMsg "(2) 變更基本資料", strResult, strFaultReason, strErrMsg
    End If
    Exit Function
ChkErr:
    ErrSub FormName, "ChangeData"
End Function
'確認使用人[22]
Public Function ConfirmCust(ByRef rsData As ADODB.Recordset, _
                            ByVal strID_Code As String, _
                            Optional ByVal strCustName As String = Empty, _
                            Optional ByVal strTel As String = Empty, _
                            Optional ByVal strCustAddress As String = Empty, _
                            Optional ByVal strSNO As String = Empty, _
                            Optional ByVal strMediaBillNo As String = Empty, _
                            Optional ByVal strProcessingDate As String = Empty, _
                            Optional ByRef strResult As String = Empty, _
                            Optional ByRef strFaultReason As String = Empty, _
                            Optional ByRef strErrMsg As String = Empty) As Boolean
                            

  On Error GoTo ChkErr
    Dim varBookMark As Variant
    'Dim strScmData As String
    'Dim objScmCommand As Object
    'Set objScmCommand = CreateObject("csSCMcontrol.clsSCMcommand")
    Dim objScmCommand As New csSCMcontrol.clsSCMcommand
    '#4262 原本是用rsData.BookMark,但會使某些RecordSet找不到指標 By Kin 2008/12/03
    varBookMark = rsData.AbsolutePosition
    ConfirmCust = False
    intCompCode = rsData("CompCode") & ""
    '#4276測試不OK,如果是FTTX要傳FaciSno By Kin 2009/01/08
    If blnExeCommand Then
        If Not objScmCommand.SCM_Command(gcnGi.ConnectionString, str(intCompCode), _
                                         GetOwner, strID_Code, _
                                         GaryGi(0), rsData("CustID"), strCustName, _
                                         strTel, strCustAddress, _
                                         rsData("DialAccount") & "", , , _
                                         rsData("ID") & "", _
                                         rsData("DeclarantName") & "", _
                                         rsData("BirthDay") & "", IIf(blnUseFttx, rsData("FaciSNo") & "", Empty), , , , , _
                                         strSNO, strMediaBillNo, strProcessingDate, Empty, Empty, _
                                         GetPrbCtl, _
                                         , strResult, strFaultReason, strErrMsg) Then GoTo 66
    End If
    If strProcessingDate <> Empty And blnExeCommand Then
        strResult = "1"
        ConfirmCust = True
        GoTo 66
    End If
    strResult = 1
    ConfirmCust = True

    
66:
    If blnShowMsg Then
        ShowMsg "(2) 確認申請人", strResult, strFaultReason, strErrMsg
    End If
    Exit Function
ChkErr:
    ErrSub FormName, "ChangeCust"

End Function

'變更申請人[21]
Public Function ChangeCust(ByRef rsData As ADODB.Recordset, ByVal strID_Code As String, _
                                        Optional ByVal strCustName As String = Empty, _
                                        Optional ByVal strTel As String = Empty, _
                                        Optional ByVal strCustAddress As String = Empty, _
                                        Optional ByVal strClassID As String = Empty, _
                                        Optional ByVal strVendCode As String = Empty, _
                                        Optional ByVal strMTAMAC As String = Empty, _
                                        Optional ByVal strSNO As String = Empty, _
                                        Optional ByVal strMediaBillNo As String = Empty, _
                                        Optional ByVal strProcessingDate As String = Empty, _
                                        Optional ByVal strAccountID As String = Empty, _
                                        Optional ByVal strPassword As String = Empty, _
                                        Optional ByVal strUserIdNo As String = Empty, _
                                        Optional ByVal strUserName As String = Empty, _
                                        Optional ByVal strUserBirthday As String = Empty, _
                                        Optional ByRef rsResult As Object = Nothing, _
                                        Optional ByRef strResult As String = Empty, _
                                        Optional ByRef strFaultReason As String = Empty, _
                                        Optional ByRef strErrMsg As String = Empty) As Boolean
  On Error GoTo ChkErr
    Dim varPara As Variant
    Dim varBookMark As Variant
    Dim strScmData As String
    Dim strUpdSO004F As String
    Dim strNowTime As String
    Dim rsTmpSO004F As New ADODB.Recordset
    Dim strOtherSQL As String
    Dim lngRcd As Long
    Dim objScmCommand As New csSCMcontrol.clsSCMcommand
    strNowTime = RightNow
    varBookMark = rsData.AbsolutePosition
    ChangeCust = False
    intCompCode = rsData("CompCode") & ""
    
    If blnExeCommand Then
        If Not objScmCommand.SCM_Command(gcnGi.ConnectionString, str(intCompCode), GetOwner, strID_Code, _
                                            GaryGi(0), rsData("CustID"), strCustName, strTel, strCustAddress, _
                                            strAccountID, IIf(blnUseFttx, Empty, strClassID), IIf(blnUseFttx, Empty, strVendCode), strUserIdNo, _
                                            strUserName, strUserBirthday, rsData("FaciSNO") & "", _
                                            IIf(blnUseFttx, Empty, strMTAMAC), , , strPassword, strSNO, strMediaBillNo, strProcessingDate, Empty, Empty, _
                                            GetPrbCtl, rsResult, strResult, strFaultReason, strErrMsg) Then GoTo 66
    End If
    If strProcessingDate <> Empty And blnExeCommand Then
        strResult = "1"
        ChangeCust = True
        GoTo 66
    End If
    If Not blnExeCommand Then
        GetScmData gcnGi, rsResult
    End If
    rsData.AbsolutePosition = varBookMark
    If Not rsResult Is Nothing Then
        '**********************************************************************************************************************
        '#4330 當原帳號與SO314.ScmData不一樣時,要將SO004F押上拆機日期 By Kin 2009/01/14
        '#4357 原本是找PrDate is Not Null 改為 PrDate is Null By Kin 2009/02/10
        If rsData("DialAccount") & "" <> rsResult("AccountId") & "" Then
            strUpdSO004F = "Select * From " & GetOwner & "SO004F Where " & _
                         " PrDate is Null And DialAccount='" & rsData("DialAccount") & "'"
            If Not GetRS(rsTmpSO004F, strUpdSO004F, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
            Do While Not rsTmpSO004F.EOF
                rsTmpSO004F("PrDate") = Format(strNowTime, "YYYY/MM/DD HH:MM:SS")
                rsTmpSO004F("UpdEn") = GaryGi(1)
                rsTmpSO004F("UpdTime") = GetDTString(strNowTime)
                rsTmpSO004F.Update
                rsTmpSO004F.MoveNext
            Loop
        End If
        '#4330 要將其它的指令的AccountID Upd成為ScmData回傳的值 By Kin 2009/01/14
        strOtherSQL = "Update " & GetOwner & "SO314 Set AccountID='" & rsResult("AccountId") & "'" & _
                    " Where CmdStatus='W' And ProcessingDate is Not Null" & _
                    " And CustId=" & rsData("CustId") & " And DeviceSNo1='" & rsData("FaciSNo") & "'" & _
                    " And StopFlag<>1 And AccountId is Not Null"
        gcnGi.Execute strOtherSQL, lngRcd
        Debug.Print lngRcd
        '**********************************************************************************************************************
        If Not blnUseFttx Then
            rsData("ServiceType") = rsResult("ServiceType")
            '------------------------------------------------------------------------------------------------------------------'
            '#7065 改由控制台回填GW回填的相關資料 By Kin 2015/08/14
            '#7213 Don't change cmopendate field by kin 2016/03/15
            'rsData("CMOpenDate") = RightNow
            '------------------------------------------------------------------------------------------------------------------'
        End If
        '------------------------------------------------------------------------------------------------------------------'
        '#7065 改由控制台回填GW回填的相關資料 By Kin 2015/08/14
        rsData("FaciStatusCode") = 1
        
        '-------------------------------------------------------------------------------------------------------------------
        rsData("DialAccount") = rsResult("AccountId")
        rsData("DialPassWord") = rsResult("AccountPassword")
        '#3941 安裝日期不可以Update By Kin 2008/05/21
        'rsData("InstDate") = Format(RightNow, "YYYY/MM/DD HH:MM:SS")
        rsData("UpdEn") = GaryGi(1)
        rsData("UpdTime") = GetDTString(strNowTime)
        rsData.Update
        strResult = "1"
        ChangeCust = True
    Else
        If blnExeCommand Then
            strErrMsg = "處理成功，但無任何回傳值!!"
        Else
            ChangeCust = True
        End If
    End If

    
66:
    If blnShowMsg Then
        ShowMsg "(1) 申請人變更", strResult, strFaultReason, strErrMsg
    End If
    On Error Resume Next
    Call CloseRecordset(rsTmpSO004F)
    Exit Function
ChkErr:
    ErrSub FormName, "ChangeCust"
End Function
'變更速率、路由[22、23]
Public Function ChangeRate(ByRef rsData As ADODB.Recordset, ByVal strID_Code As String, _
                                        Optional ByVal strClassID As String = Empty, _
                                        Optional ByVal strVendCode As String = Empty, _
                                        Optional ByVal strVendName As String = Empty, _
                                        Optional ByVal strCMBaudNo As String = Empty, _
                                        Optional ByVal strCMBaudRate As String = Empty, _
                                        Optional ByVal strSNO As String = Empty, _
                                        Optional ByVal strMediaBillNo As String = Empty, _
                                        Optional ByVal strProcessingDate As String = Empty, _
                                        Optional ByRef rsVirtual As Object = Nothing, _
                                        Optional ByRef strResult As String = Empty, _
                                        Optional ByRef strFaultReason As String = Empty, _
                                        Optional ByRef strErrMsg As String = Empty) As Boolean
   On Error GoTo ChkErr
    Dim strCmdName As String
    Dim varBookMark As Variant
    Dim objScmCommand As New csSCMcontrol.clsSCMcommand
    Dim i As Integer
    Dim strNowTime As String
    Dim rsSO004F As New ADODB.Recordset
    Dim strDialAccount As String
    Dim strExtendAccount As String
    Dim strExtendPassword As String
    Dim strSO004FSQL As String
    ChangeRate = False
    '#4262 原本是用rsData.BookMark,但會使某些RecordSet找不到指標 By Kin 2008/12/03
    varBookMark = rsData.AbsolutePosition
    intCompCode = rsData("CompCode") & ""
    strNowTime = RightNow
    Select Case strID_Code
        Case "22"
            strCmdName = "變更速率"
        Case "23"
            strCmdName = "變更路由"
    End Select
    '#4209 (1)當指令為22變更速率，只填Schemacode1，不填Schemacode2(2)當指令為23變更速率，不填Schemacode1，只填Schemacode2 By Kin 2008/11/06
    If blnExeCommand Then
        If Not objScmCommand.SCM_Command(gcnGi.ConnectionString, str(intCompCode), GetOwner, strID_Code, _
                                            GaryGi(0), rsData("Custid") & "", , , , rsData("DialAccount") & "", IIf(strID_Code = "22", strClassID, Empty), _
                                            IIf(strID_Code = "22", Empty, strVendCode), _
                                            , , , rsData("FaciSNo") & "", , , , , strSNO, strMediaBillNo, _
                                            strProcessingDate, Empty, Empty, _
                                            GetPrbCtl, _
                                            rsVirtual, strResult, strFaultReason, strErrMsg) Then GoTo 66
    End If
    If strProcessingDate <> Empty And blnExeCommand Then
        strResult = "1"
        ChangeRate = True
        GoTo 66
    End If
    
    rsData.AbsolutePosition = varBookMark
    
    Select Case strID_Code
        Case "22"
            rsData("CMBaudNo") = IIf(strCMBaudNo <> "", strCMBaudNo, Null)
            rsData("CMBaudRate") = IIf(strCMBaudRate <> "", strCMBaudRate, Null)
        Case "23"
            '#4276 如果是Fttx,而且是固定IP會有回傳值 By Kin 2008/12/26
            'If blnUseFttx Then
            If Not rsVirtual Is Nothing Then
                '#5165 將下段程式Mark,改用UpdSO004F下去跑 By Kin 2009/07/03
                If Not UpdSO004F(rsVirtual, strNowTime) Then GoTo 66
'                strSO004FSQL = "Select * From " & GetOwner & "SO004F" & _
'                            " Where ExtendAccount='" & rsVirtual("FixIPAccountId") & "'" & _
'                            " And ExtendType=1 And PrDate is Null"
'                If Not GetRS(rsSO004F, strSO004FSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
'                If Not rsSO004F.EOF Then
'                    strDialAccount = rsSO004F("DialAccount") & ""
'                    strExtendAccount = rsSO004F("ExtendAccount") & ""
'                    strExtendPassword = rsSO004F("ExtendPassword") & ""
'                    rsSO004F.MoveFirst
'                    Do While Not rsSO004F.EOF
'                        rsSO004F("PrDate") = Format(strNowTime, "YYYY/MM/DD HH:MM:SS")
'                        rsSO004F("UpdEn") = GaryGi(1)
'                        rsSO004F("UpdTime") = GetDTString(strNowTime)
'                        rsSO004F.Update
'                        rsSO004F.MoveNext
'                    Loop
'                End If
'                For i = 0 To rsVirtual.Fields.Count - 3
'                    With rsSO004F
'                        .AddNew
'                        .Fields("SeqNo") = Get_SO004F_Seq_No(gcnGi)
'                        .Fields("DialAccount") = strDialAccount
'                        .Fields("ExtendAccount") = strExtendAccount
'                        .Fields("ExtendPassword") = strExtendPassword
'                        .Fields("FixIPAddress") = rsVirtual(i + 2).Value & ""
'                        .Fields("UpdEn") = GaryGi(1)
'                        .Fields("UpdTime") = GetDTString(strNowTime)
'                        .Update
'                    End With
'                Next i
            End If
            
            'End If
            rsData("VendorCode") = IIf(strVendCode <> "", strVendCode, Null)
            rsData("VendorName") = IIf(strVendName <> "", strVendName, Null)
    End Select
    If Not blnUseFttx Then
        rsData("UpdEn") = GaryGi(1)
        rsData("UpdTime") = GetDTString(strNowTime)
        rsData.Update
    End If
    strResult = "1"
    ChangeRate = True


66:
    If blnShowMsg Then
        ShowMsg "(2) " & strCmdName, strResult, strFaultReason, strErrMsg
    End If
    On Error Resume Next
    Call CloseRecordset(rsSO004F)
    Exit Function
ChkErr:
    ErrSub FormName, "ChangeRate"
End Function
'變更密碼[24]
Public Function ChangePassword(ByRef rsData As ADODB.Recordset, ByVal strID_Code As String, _
                                                        Optional ByVal strPassword As String = "", _
                                                        Optional ByVal strSNO As String = "", _
                                                        Optional ByVal strMediaBillNo As String = "", _
                                                        Optional ByVal strProcessingDate As String = "", _
                                                        Optional ByRef strResult As String, _
                                                        Optional ByRef strFaultReason As String, _
                                                        Optional ByRef strErrMsg As String) As Boolean

  On Error GoTo ChkErr
    Dim varBookMark As Variant
    'Dim objScmCommand As Object
    'Set objScmCommand = CreateObject("csSCMcontrol.clsSCMcommand")
    Dim objScmCommand As New csSCMcontrol.clsSCMcommand
    ChangePassword = False
    '#4262 原本是用rsData.BookMark,但會使某些RecordSet找不到指標 By Kin 2008/12/03
    varBookMark = rsData.AbsolutePosition
    intCompCode = rsData("CompCode") & ""
    
    If blnExeCommand Then
        If Not objScmCommand.SCM_Command(gcnGi.ConnectionString, str(intCompCode), GetOwner, strID_Code, _
                                            GaryGi(0), rsData("Custid") & "", , , , rsData("DialAccount") & "", , , , , , rsData("FaciSNo") & "", , , , _
                                            strPassword, strSNO, strMediaBillNo, strProcessingDate, Empty, Empty, _
                                            GetPrbCtl, , strResult, strFaultReason, strErrMsg) Then GoTo 66
    End If
    If strProcessingDate <> Empty And blnExeCommand Then
        strResult = "1"
        ChangePassword = True
        GoTo 66
    End If
                                        
    rsData.AbsolutePosition = varBookMark
    rsData("DialPassword").Value = IIf(strPassword <> "", strPassword, Null)
    rsData("UpdEn") = GaryGi(1)
    rsData("UpdTime") = GetDTString(RightNow)
    rsData.Update
    strResult = "1"
    ChangePassword = True
66:
    If blnShowMsg Then
        ShowMsg "(2) 變更密碼 ", strResult, strFaultReason, strErrMsg
    End If
    Exit Function
ChkErr:
    ErrSub FormName, "ChangeRate"
End Function
'查詢CM資訊、查詢帳號資訊[30、31]
Public Function QueryCMStatus(rsData As ADODB.Recordset, ByVal strID_Code As String, _
                                Optional ByVal strModeID As String = Empty, _
                                Optional ByVal strHFCMAC As String = Empty, _
                                Optional ByVal strAccountID As String = Empty, _
                                Optional ByVal strProcessingDate As String = Empty, _
                                Optional ByRef rsResult As ADODB.Recordset, _
                                Optional ByRef strResult As String = Empty, _
                                Optional ByRef strFaultReason As String = Empty, _
                                Optional ByRef strErrMsg As String = Empty) As Boolean
  On Error GoTo ChkErr
    Dim strCmdName As String
    Dim varArray As Variant
    Dim strScmData As String
    'Dim objScmCommand As Object
    'Set objScmCommand = CreateObject("csSCMcontrol.clsSCMcommand")
    Dim objScmCommand As New csSCMcontrol.clsSCMcommand
    Select Case strID_Code
        Case 30
            strCmdName = "查詢CM資訊"
        Case 31
            strCmdName = "查詢帳號資訊"
        Case Else
    End Select
    QueryCMStatus = False
    intCompCode = rsData("CompCode") & ""
    If blnExeCommand Then
        If Not objScmCommand.SCM_Command(gcnGi.ConnectionString, str(intCompCode), GetOwner, strID_Code, _
                                            GaryGi(0), rsData("Custid") & "", , , , strAccountID, , , , , , strHFCMAC, , _
                                            strModeID, , , , , strProcessingDate, Empty, Empty, _
                                            GetPrbCtl, _
                                            rsResult, strResult, strFaultReason, strErrMsg) Then GoTo 66
    End If
    If strProcessingDate <> Empty Then
        strResult = "1"
        QueryCMStatus = True
        GoTo 66
    End If
    If Not rsResult Is Nothing Then
        strResult = "1"
        QueryCMStatus = True
    Else
        strErrMsg = "查詢成功，但無任何傳回傳!"
    End If
    

66:
    If blnShowMsg Then
        ShowMsg "(3) " & strCmdName, strResult, strFaultReason, strErrMsg
    End If
    Exit Function
ChkErr:
    ErrSub FormName, "QueryCMStatus"
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
Public Function RequireFreeSWPort(rsData As ADODB.Recordset, ByVal strID_Code As String, _
                                    Optional ByRef strResult As String = Empty, _
                                    Optional ByRef strFaultReason As String = Empty, _
                                    Optional ByRef strErrMsg As String = Empty) As Boolean
    On Error GoTo ChkErr
    Dim strCmdName As String
    Dim strCustID As String
    Dim rsReturn As New ADODB.Recordset
    Dim strParamData As String
    Dim aryParamData As Variant
    Dim strSwitchErr As String
    Dim objScmCommand As New csSCMcontrol.clsSCMcommand
    Dim i As Integer
    Dim j As Integer
    Dim lngNum As Long
    Dim strNum As String
    RequireFreeSWPort = False
    strSwitchErr = ""
    intCompCode = rsData("CompCode") & ""
    aryParamData = Split(GetSwitchID(GetInstAddNo(rsData), strSwitchErr), ",")
    For i = LBound(aryParamData) To UBound(aryParamData)
        strParamData = IIf(strParamData = Empty, "", strParamData & Chr(13))
        strParamData = strParamData & "SWNo" & i + 1 & "=" & aryParamData(i)
    Next
    
    If blnShowMsg Then
        strCustID = rsData("CustId") & ""
    End If
    If strParamData = "" Then
        strResult = ""
        strErrMsg = strSwitchErr
        GoTo 66
    End If
    If blnExeCommand Then
        If Not objScmCommand.SCM_Command(gcnGi.ConnectionString, str(intCompCode), GetOwner, strID_Code, _
                                            GaryGi(0), strCustID, , , , , , , , , , _
                                            , , , , , , , , , _
                                            , GetPrbCtl, rsReturn, strResult, strFaultReason, strErrMsg, strParamData) Then GoTo 66
    End If
    If Not rsReturn Is Nothing Then
        '#4367 原本資訊不夠明確,回來的東西再加工處理一次 By Kin 2009/02/16
        '#4367 測試不OK,會多出一個剩餘Port數,所以多加判斷 By Kin 2009/02/24
        If Not rsReturn.EOF Then strResult = "公司別 : " & intCompCode
        For i = 0 To rsReturn.Fields.Count - 1
            If InStr(1, rsReturn.Fields(i).Name, "SWNo") > 0 Then
                strResult = strResult & "可用之Switch ID : " & rsReturn.Fields(i).Value & ""
            Else
                If InStr(1, rsReturn.Fields(i).Name, "SWPort") > 0 Then
                    strResult = strResult & "剩餘PORT數 : " & TurnNum(rsReturn.Fields(i).Value & "") & ""
                End If
            End If
            'strResult = strResult & rsReturn.Fields(i).Name & "=" & rsReturn.Fields(i).Value
            If i <> rsReturn.Fields.Count - 1 Then
                strResult = strResult & vbCrLf
            End If
        Next
        RequireFreeSWPort = True
    Else
        RequireFreeSWPort = False
    End If
    
66:
    If blnShowMsg Then
        ShowMsg "(3) 查詢FreePort", strResult, strFaultReason, strErrMsg
    End If
    Exit Function
ChkErr:
    ErrSub FormName, "RequireFreeSWPort"
End Function


'設備出入庫[40、41]
Public Function StockCM(rsData As ADODB.Recordset, ByVal strID_Code As String, _
                                Optional ByVal strModeID As String = Empty, _
                                Optional ByVal strModeCode As String = Empty, _
                                Optional ByVal strHFCMAC As String = Empty, _
                                Optional ByVal strMTAMAC As String = Empty, _
                                Optional ByVal strSNO As String = Empty, _
                                Optional ByVal strMediaBillNo As String = Empty, _
                                Optional ByVal strProcessingDate As String = Empty, _
                                Optional ByRef strResult As String = Empty, _
                                Optional ByRef strFaultReason As String = Empty, _
                                Optional ByRef strErrMsg As String = Empty) As Boolean
  On Error GoTo ChkErr
    Dim strCmdName As String
    Dim strCustID As String
    'Dim objScmCommand As Object
    'Set objScmCommand = CreateObject("csSCMcontrol.clsSCMcommand")
    Dim objScmCommand As New csSCMcontrol.clsSCMcommand
    strCustID = Empty
    Select Case strID_Code
        Case 40
            strCmdName = "設備入庫"
        Case 41
            strCmdName = "設備出庫"
        Case Else
    End Select
    StockCM = False
    intCompCode = rsData("CompCode") & ""
    
    If blnShowMsg Then
        strCustID = rsData("CustId") & ""
    End If
    If blnExeCommand Then
        If Not objScmCommand.SCM_Command(gcnGi.ConnectionString, str(intCompCode), GetOwner, strID_Code, _
                                            GaryGi(0), strCustID, , , , , , , , , , _
                                            strHFCMAC, strMTAMAC, strModeID, strModeCode, strSNO, strMediaBillNo, strProcessingDate, Empty, Empty, _
                                            , GetPrbCtl, , strResult, strFaultReason, strErrMsg) Then GoTo 66
            
    End If
    StockCM = True
66:
    If blnShowMsg Then
        ShowMsg "(4) " & strCmdName, strResult, strFaultReason, strErrMsg
    End If
    Exit Function
ChkErr:
    ErrSub FormName, "StockCM"
End Function


Public Function GetStateData(objCn, ByVal intCompCode As Integer) As String
  On Error Resume Next
    Dim s As String
    GetStateData = objCn.Execute("SELECT CmdStatus FROM " & strCMowner & "SO314" & _
                                                        " WHERE  CompCode=" & intCompCode & " AND " & _
                                                        " CMDSEQNO='" & strCmdSeqNo & "'")(0) & ""
                                                        
                                                
End Function
Public Function GetErrMsg(objCn, ByVal intCompCode As Integer, ByRef strErrMsg As String, ByRef strFaultReason As String) As Boolean
    On Error Resume Next
    Dim varErrPara As String
    GetErrMsg = False
    varErrPara = objCn.Execute("Select ErrCode,ErrMsg From " & strCMowner & "SO314" & _
                                                " Where CompCode=" & intCompCode & " And " & _
                                                " CMDSEQNO='" & strCmdSeqNo & "'").GetString(, , ",")
    strErrMsg = Split(varErrPara, ",")(1)
    strFaultReason = Split(varErrPara, ",")(0)
    'If strErrMsg <> "" Then strErrMsg = "錯誤代碼:" & strErrMsg
    If Trim(strErrMsg) <> "" Or Trim(strFaultReason) <> "" Then
        GetErrMsg = True
    End If
End Function
Private Sub ShowMsg(strCMD As String, strResult As String, strFaultReason As String, strErrMsg As String)
    On Error GoTo ChkErr
        If strResult = "1" Then
            MsgBox "[ " & strCMD & " ] 控制訊號已送出完成 !! ", vbInformation, gimsgPrompt
        Else
            If strFaultReason = Empty Then
                MsgBox "[ " & strCMD & " ] 失敗 !! " & vbCrLf & vbCrLf & _
                                "錯誤原因 : " & strErrMsg, vbInformation, gimsgPrompt
            Else
                MsgBox "[ " & strCMD & " ] 失敗 !! " & vbCrLf & vbCrLf & _
                                "錯誤原因 : " & strFaultReason & vbCrLf & vbCrLf & _
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
    '#5149 增加檢查 是否有CATV移機未完工的移裝地址可用 By Kin 2009/06/18
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
'Corey 寫的Function
'利用傳入的ADDRNO(一定要裝機地址)找出SO014.ADDSORT，並且SO017C的涵蓋地址有包含此ADDRSORT的SNO列出來，
'並且利用SNO找出SO017B.SwitchID回傳，用逗號隔開每個SwitchID
Public Function GetSwitchID(ByVal strInstAddrNO As String, ByRef strErrString As String) As String
On Error GoTo ChkErr
    Dim rsTmp As New ADODB.Recordset
    Dim strAddrSort As String
    Dim strSNO As String
    Dim strSwitchID As String
    Dim intAddr As Long
    GetSwitchID = ""
    strSwitchID = ""
    strErrString = ""
    If Not GetRS(rsTmp, "Select AddrSort,FTTxFlag,NO1 From " & GetOwner & "SO014 Where AddrNO=" & strInstAddrNO) Then Exit Function
    If rsTmp.RecordCount = 0 Then
        strErrString = "找不到對應地址"
        Exit Function '找不到對應的地址 Addrsort
    End If
'    If Val(rsTmp("FTTxFlag") & "") = 0 Then
'        strErrString = "該地址不在 FTTX 涵蓋範圍"
'        Exit Function
'    End If
    '舊地址格式要再減6 For Corey By Kin 2010/02/25
    
    If gcnGi.Execute("Select Nvl(AddrType,0) From " & GetOwner & "SO041 Where SysId=" & intCompCode)(0) = 0 Then
        intAddr = 6
    Else
        intAddr = 0
    End If
    strAddrSort = Mid(rsTmp("AddrSort") & "", 1, GetFldMaxLen("SO017C", "AddrSortA") - intAddr)
    'If Not GetRS(rsTmp, "Select * From " & GetOwner & "SO017C Where AddrSortA <= '" & strAddrSort & "' and AddrSortB >= '" & strAddrSort & "'") Then Exit Function
    '#5634 2010.04.30 by Corey 增加判斷地址號碼，單號或雙號判斷串語法找尋SO017C的地址資料
    If Not GetRS(rsTmp, "Select * From " & GetOwner & "SO017C Where AddrSortA <= '" & strAddrSort & "' and AddrSortB >= '" & strAddrSort & "'" & _
                        " And (NOE Is Null Or NOE=0 Or NOE=" & IIf(rsTmp("NO1") Mod 2 = 1, "1", "2") & ")") Then Exit Function
    '合乎條件的數量大於0則開始串語法，否則回傳空字串
    If rsTmp.RecordCount > 0 Then
        '先取得SO017C的SNO，再利用SNO找出SO017B的SwitchID
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
        If strSwitchID = "" Then strErrString = "找不到對應的SwitchID"
    Else
         strErrString = "找不到對應的SwitchID"
    End If
    On Error Resume Next
    CloseRecordset rsTmp
Exit Function
ChkErr:
    ErrSub FormName, "GetSwitchID"
End Function
Public Function GetFldMaxLen(ByVal uTableName As String, ByVal uColumn As String) As Double

  On Error GoTo ChkErr
  
    '此Function是取出資料庫內本身的SCHEMA的欄位大小值
    '並且SCHEMA的DATA_TYPE=NUMBER，並計算出NUMBER最大值回傳Double
    '負數的時候是從 -1.79769313486232E308 到 -4.94065645841247E-324
    '正數的時候是從 4.94065645841247E-324 到 1.79769313486232E308
    
    'uTableName=傳入Table名稱  uColumn=傳入欄位名稱

    Dim rsColumn As New ADODB.Recordset
    Dim strSQL As String
    
    GetFldMaxLen = 0

    '沒有傳資料進來則直接離開，否則都語法時會出錯誤訊息
    If IsNull(uTableName) Or IsNull(uColumn) Then Exit Function
    
    strSQL = "Select * From " & "User_Tab_Columns" & _
                " Where Table_Name='" & UCase(uTableName) & "'" & _
                " And Column_Name='" & UCase(uColumn) & "'"
    
    If Not GetRS(rsColumn, strSQL) Then Exit Function
    
    Select Case UCase(rsColumn("Data_Type") & "")
        Case "NUMBER"
            GetFldMaxLen = (10 ^ rsColumn("Data_Precision")) - 1
            If rsColumn("Data_Scale") > 0 Then '判斷小數點是否有值
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
    If blnUseFttx Then chkUseFlag = True: Exit Function
    
    If Not GetRS(rs, "Select Useflag From " & GetOwner & "SO306 Where HFCMAC='" & strFaciSNo & "'") Then Exit Function
    If rs.EOF Then Call CloseRecordset(rs): Exit Function
    '******************************************************************************************************
    '#3755 如果Linkmss為0不檢查UseFlag By Kin 2008/2/01
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

Public Sub GetScmData(ByRef objCn As Object, ByRef rsResult As Object)
  On Error Resume Next
    Dim strScmData As String
    Dim strTmpName As String
    Dim strTmpValue As String
    Dim strFldName() As String
    Dim strFldValue() As String
    Dim varArray As Variant
    Dim i As Long
    strScmData = objCn.Execute("SELECT ScmData FROM " & strCMowner & "SO314" & _
                                           " WHERE  CompCode=" & intCompCode & " AND " & _
                                           " CMDSEQNO='" & strCmdSeqNo & "'")(0) & ""
    If strScmData <> Empty Then
        strScmData = Replace(strScmData, Chr(10), "")
        varArray = Split(strScmData, Chr(13))
        ReDim strFldName(UBound(varArray))
        ReDim strFldValue(UBound(varArray))
        For i = 0 To UBound(varArray)
            strTmpName = Replace(Replace(Mid(varArray(i), 1, InStr(1, varArray(i), "=") - 1), Chr(10), ""), Chr(13), "")
            strTmpValue = Replace(Replace(Mid(varArray(i), InStr(1, varArray(i), "=") + 1), Chr(10), ""), Chr(13), "")
            If Trim(strTmpName) <> "" Then strFldName(i) = strTmpName
            If Trim(strTmpValue) <> "" Then strFldValue(i) = strTmpValue
            If Trim(strFldValue(i)) = "空" Then strFldValue(i) = ""
            strTmpName = ""
            strTmpValue = ""
        Next i
        Call SetRsMemory(rsResult, strFldName, strFldValue)
    Else
        Set rsResult = Nothing
    End If
End Sub
Public Sub SetRsMemory(ByRef rs As Object, ByRef arrFldName() As String, ByRef arrFldValue() As String)
  On Error GoTo ChkErr
    Dim i As Long
    Set rs = CreateObject("ADODB.Recordset")
    If rs.State = adStateOpen Then rs.Close
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenKeyset
    rs.LockType = adLockOptimistic
    For i = 0 To UBound(arrFldName)
        If arrFldName(i) <> "" Then
            rs.Fields.Append arrFldName(i), adBSTR, 100
        End If
    Next i
    rs.Open
    rs.AddNew
    For i = 0 To UBound(arrFldName)
        If arrFldName(i) <> "" Then
            rs.Fields(arrFldName(i)).Value = arrFldValue(i) & ""
        End If
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
