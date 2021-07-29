VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsExeCommand"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Function ExeCMCommand(ByRef cn As ADODB.Connection, ByRef rs As ADODB.Recordset, _
                             ByVal strCMD As String, ByRef GaryGi() As String, _
                             Optional ByVal strFaciCode As String = Empty, _
                             Optional ByVal strFaciSNo As String = Empty, _
                             Optional ByVal strSNO As String = Empty, _
                             Optional ByVal strMediaBillNo As String = Empty, _
                             Optional ByVal strProcessingDate As String = Empty, _
                             Optional ByRef rs2 As ADODB.Recordset, _
                             Optional ByRef strResult As String, _
                             Optional ByRef strFaultReason As String, _
                             Optional ByRef strErrMsg As String, _
                             Optional ByVal blnExe As Boolean = True, _
                             Optional ByRef strCmdSeq As String = Empty, _
                             Optional ByVal strBaudRate As String = Empty, _
                             Optional ByVal strBaudName As String = Empty, _
                             Optional ByVal strVendorCode As String = Empty, _
                             Optional ByVal strVendorName As String = Empty, _
                             Optional ByVal strDevice2 As String, _
                             Optional ByVal strReasonCode As String, _
                             Optional ByVal strReasonName As String, _
                             Optional ByVal strDynIpNum As String = Empty, _
                             Optional ByVal strFixIPNum As String = Empty) As Boolean
    On Error Resume Next
    blnShowMsg = False
    blnNoShowMsg = True
    ExeCMCommand = False
    blnExeCommand = blnExe '判斷是否執行命令再更新SO004
    Dim strArry As String
    Dim strSQLRef As String
    Set frmSO1623A = Nothing
    strArry = Join(GaryGi, Chr(9))
    Set gcnGi = cn
    
    Call GetGlobal(strArry)
    gErrLogPath = ReadGICMIS1("ErrLogPath")
    blnRecordProcedure = IsRecordProcedure
    Call WriteRecordVodProcedure("1.開始執行無UI命令")
    If rs Is Nothing Then Call WriteRecordVodProcedure("SO004 Is Nothing")
    If rs.EOF Then Call WriteRecordVodProcedure("無任何SO004資料傳入!!!")
    '************************************************
    '#4276 判斷是否使用FTTX By Kin 2008/12/25
    '#4357 設備出入庫(40、41)一定不是FTTX設備,而且改命令Miggie沒有傳FaciCode,所以不能判斷 By Kin 2009/02/10
    If strCMD <> 32 Or strCMD <> 40 Or strCMD <> 41 Then
        strSQLRef = "Select Count(*) From " & GetOwner & "CD022 Where RefNo=8 And CodeNo=" & rs("FaciCode")
        WriteRecordVodProcedure ("1-1.判斷是否為FTTX")
        If gcnGi.Execute(strSQLRef)(0) > 0 Then
            blnUseFttx = True
        Else
            blnUseFttx = False
        End If
    End If
    '************************************************
    '******************************************************************************************************************************
    '#3353 監控中心所用，當blnExeCommand=False時，則不執行命令，直接進入命令裡的SO004更新 By Kin 2008/02/01
    If Not blnExe Then
        strCMowner = gcnGi.Execute("Select CMOwner From " & GetOwner & "SO041 Where CompCode=" & rs("CompCode"))(0) & ""
        strCmdSeqNo = strCmdSeq
        If strCMowner <> Empty Then
            strCMowner = strCMowner & "."
        End If
        'strCMowner = GetOwner
    End If
    '*****************************************************************************************************************************
    If rs Is Nothing Then Exit Function
    If rs.EOF Then Exit Function
    WriteRecordVodProcedure ("2.準備進入執行命令 > " & strCMD)
    Select Case strCMD
        Case "10"   '裝機
            '#4276 增加啟用PPPOE帳號 By Kin 2008/12/25
            If blnUseFttx Then
                ExeCMCommand = ExeOpenPPPOEAccount(rs, "10", strSNO, strMediaBillNo, strProcessingDate, strResult, strFaultReason, strErrMsg)
            Else
                ExeCMCommand = ExeInstallCM(rs, "10", strSNO, strMediaBillNo, strProcessingDate, strResult, strFaultReason, strErrMsg)
            End If
        Case "11"   '開機
            ExeCMCommand = ExeOpenCM(rs, "11", strSNO, strMediaBillNo, strProcessingDate, strResult, strFaultReason, strErrMsg)
        Case "12"   '關機
            ExeCMCommand = ExeCloseCM(rs, "12", strSNO, strMediaBillNo, strProcessingDate, strReasonCode, strReasonName, strResult, strFaultReason, strErrMsg)
        Case "13"   '停機
            ExeCMCommand = ExeStopCM(rs, "13", strSNO, strMediaBillNo, strProcessingDate, strReasonCode, strReasonName, strResult, strFaultReason, strErrMsg)
        Case "14"   '復機
            ExeCMCommand = ExeResumeCM(rs, "14", strSNO, strMediaBillNo, strProcessingDate, strResult, strFaultReason, strErrMsg)
        Case "15"   '拆機
            ExeCMCommand = ExeDownCM(rs, "15", strSNO, strMediaBillNo, strProcessingDate, strReasonCode, strReasonName, strResult, strFaultReason, strErrMsg)
        Case "16"   '維修換拆
            If blnExe Then  '執行工單流程(要先執行拆機才可維修換拆)
                If rs2 Is Nothing Then Exit Function
                If Not ExeDownCM(rs2, "15", strSNO, strMediaBillNo, strProcessingDate, strResult, strFaultReason, strErrMsg) Then Exit Function
                ExeCMCommand = ExeRepairCM(rs, "16", strFaciCode, strFaciSNo, strSNO, strMediaBillNo, strProcessingDate, strResult, strFaultReason, strErrMsg)
            Else
                ExeCMCommand = ExeRepairCM(rs, "16", strFaciCode, strFaciSNo, strSNO, strMediaBillNo, strProcessingDate, strResult, strFaultReason, strErrMsg)
            End If
        '***************************************************************************************************************************
        '#3755 增加此命令 By Kin 2008/02/01
        Case "17"   '重置設備
            ExeCMCommand = ExeRest(rs, "17", strSNO, strMediaBillNo, strProcessingDate, strResult, strFaultReason, strErrMsg)
        '#4276 增加FTTX預約裝機 By Kin 2008/12/29
        Case "18"
            ExeCMCommand = ExeRequirePPPOEAccount(rs, "18", strSNO, strMediaBillNo, strProcessingDate, strDynIpNum, strFixIPNum, strResult, strFaultReason, strErrMsg)
        '****************************************************************************************************************************
        '****************************************************************************************************************************
        '#4276 增加FTTX命令 By Kin 2008/12/25
        Case "19" '取消預約FTTX裝機
            If Not blnUseFttx Then
                strErrMsg = "設備不為FTTX"
                ExeCMCommand = False
                Exit Function
            End If
            ExeCMCommand = ExeReleasePPPOEAccount(rs, "19", strSNO, strMediaBillNo, strProcessingDate, strResult, strFaultReason, strErrMsg)
        '******************************************************************************************************************************
        Case "20"   '變更基本資料
            ExeCMCommand = ExeChangeData(rs, "20", strSNO, strMediaBillNo, strProcessingDate, strResult, strFaultReason, strErrMsg)
        Case "21"   '變更申請人
            ExeCMCommand = ExeChangeCust(rs, "21", strSNO, strMediaBillNo, strProcessingDate, strResult, strFaultReason, strErrMsg)
        Case "22"   '變更速率
            ExeCMCommand = ExeChangeRate(rs, "22", strBaudRate, strBaudName, , , strSNO, strMediaBillNo, strProcessingDate, strResult, strFaultReason, strErrMsg)
        Case "23"   '變更路由
            ExeCMCommand = ExeChangeRate(rs, "23", , , strVendorCode, strVendorName, strSNO, strMediaBillNo, strProcessingDate, strResult, strFaultReason, strErrMsg)
        Case "24"   '變更密碼
            ExeCMCommand = ExeChangePassword(rs, "24", strPassword, strSNO, strMediaBillNo, strProcessingDate, strResult, strFaultReason, strErrMsg)
        '*****************************************************************************************************************************
        '#3755 增加此命令 By Kin 2008/02/01
        Case "25"   '確認使用人
            ExeCMCommand = ExeConfirmCust(rs, "25", strSNO, strMediaBillNo, strProcessingDate, strResult, strFaultReason, strErrMsg)
        '*****************************************************************************************************************************
        '*****************************************************************************************************************************
        '#4276 增加申請固定IP By Kin 2008/12/26
        Case "26"
            ExeCMCommand = ExeRequireFixIP(rs, "26", strSNO, strMediaBillNo, strProcessingDate, strDynIpNum, strFixIPNum, strResult, strFaultReason, strErrMsg)
        '#4276 增加取消固定IP By Kin 2008/12/29
        Case "27"
            ExeCMCommand = ExeReleaseFixIP(rs, "27", strSNO, strMediaBillNo, strProcessingDate, strResult, strFaultReason, strErrMsg)
        '#4276 增加重設PPPOE密碼 By Kin 2008/12/29
        Case "28"
            ExeCMCommand = ExeResetPPPOEPassowrd(rs, "28", strSNO, strMediaBillNo, strProcessingDate, strResult, strFaultReason, strErrMsg)
        '#4276 增加同區移入 By Kin 2008/12/30
        Case "29"
            ExeCMCommand = ExeMoveFaci(rs, strSNO, strMediaBillNo, strProcessingDate, strResult, strFaultReason, strErrMsg)
        '*****************************************************************************************************************************
        Case "30"
            ExeCMCommand = True
        Case "31"
            ExeCMCommand = True
        '#4276 查詢FreePort By Kin 2008/12/29
        Case "32"
            ExeCMCommand = ExeRequireFreeSWPort(rs, strResult, strFaultReason, strErrMsg)
        '#4384 增加Release Port 命令 By Kin 2009/02/26
        Case "33"
            ExeCMCommand = ExeReleasePort(rs, "33", strSNO, strMediaBillNo, strProcessingDate, strReasonCode, strReasonName, strResult, strFaultReason, strErrMsg)
        
        Case "40"   '設備入庫
            blnUseFttx = False  '#4357 一定不是FTTX設備 By Kin 2009/02/10
            ExeCMCommand = ExeStockCM(rs, "40", strDevice2, strProcessingDate, strResult, strFaultReason, strErrMsg)
        Case "41"   '設備出庫
            blnUseFttx = False  '#4357 一定不是FTTX設備 By Kin 2009/02/10
            ExeCMCommand = ExeStockCM(rs, "41", strDevice2, strProcessingDate, strResult, strFaultReason, strErrMsg)
        Case Else
            strErrMsg = "找不到命令"
            ExeCMCommand = False
    End Select
    strCmdSeq = strCmdSeqNo
End Function
'#4276 增加同區移入 By Kin 2008/12/30
Private Function ExeMoveFaci(ByRef rs As ADODB.Recordset, _
                             Optional ByVal strSNO As String = Empty, _
                             Optional ByVal strMediaBillNo As String = Empty, _
                             Optional ByVal strProcessingDate As String = Empty, _
                             Optional ByRef strResult As String, _
                             Optional ByRef strFaultReason As String, _
                             Optional ByRef strErrMsg As String) As Boolean
On Error GoTo ChkErr
    blnShowMsg = False
    blnNoShowMsg = True
    ExeMoveFaci = False
    
    'If Not chkUseFlag(strFaciSNo, Val(rs("CompCode") & "")) Then Exit Function
    ExeMoveFaci = OpenCM(rs, "29", rs("FaciSNo") & "", IIf(IsEMTA(rs("FaciCode") & ""), Get_EMTA_Mac(rs("FaciSNo") & ""), ""), _
                         strSNO, strMediaBillNo, strProcessingDate, strResult, strFaultReason, strErrMsg)
    Exit Function
ChkErr:
    ErrSub "clsExeCommand", "ExeMoveFaci"

End Function
'#4276 查詢FreePort By Kin 2008/12/29
Private Function ExeRequireFreeSWPort(ByRef rs As ADODB.Recordset, _
                                      Optional ByRef strResult As String, _
                                      Optional ByRef strFaultReason As String, _
                                      Optional ByRef strErrMsg As String) As Boolean
    On Error GoTo ChkErr
    blnShowMsg = False
    blnNoShowMsg = True
    ExeRequireFreeSWPort = False
    blnUseFttx = True
'    If Not blnUseFttx Then
'        strResult = Empty
'        strErrMsg = "不是FTTX設備"
'        Exit Function
'    End If
    ExeRequireFreeSWPort = RequireFreeSWPort(rs, "32", strResult, strFaultReason, strErrMsg)
    Exit Function
ChkErr:
    ErrSub "clsExeCommand", "ExeRequireFreeSWPort"

End Function

'#4276 增加重置PPPOE密碼 By Kin 2008/12/29
Private Function ExeResetPPPOEPassowrd(ByRef rs As ADODB.Recordset, _
                                 ByVal strCMD As String, _
                                 Optional ByVal strSNO As String = Empty, _
                                 Optional ByVal strMediaBillNo As String = Empty, _
                                 Optional ByVal strProcessingDate As String = Empty, _
                                 Optional ByRef strResult As String, _
                                 Optional ByRef strFaultReason As String, _
                                 Optional ByRef strErrMsg As String) As Boolean
    On Error GoTo ChkErr
    blnShowMsg = False
    blnNoShowMsg = True
    ExeResetPPPOEPassowrd = False
    
    ExeResetPPPOEPassowrd = ResetPPPOEPassowrd(rs, "27", strSNO, strMediaBillNo, strProcessingDate, strResult, strFaultReason, strErrMsg)
    
    Exit Function
ChkErr:
    ErrSub "clsExeCommand", "ExeResetPPPOEPassowrd"

End Function
'#4276 增加取消固定IP By Kin 2008/12/29
Private Function ExeReleaseFixIP(ByRef rs As ADODB.Recordset, _
                                 ByVal strCMD As String, _
                                 Optional ByVal strSNO As String = Empty, _
                                 Optional ByVal strMediaBillNo As String = Empty, _
                                 Optional ByVal strProcessingDate As String = Empty, _
                                 Optional ByRef strResult As String, _
                                 Optional ByRef strFaultReason As String, _
                                 Optional ByRef strErrMsg As String) As Boolean
    On Error GoTo ChkErr
    blnShowMsg = False
    blnNoShowMsg = True
    ExeReleaseFixIP = False
    
    ExeReleaseFixIP = ReleaseFixIP(rs, "27", strSNO, strMediaBillNo, strProcessingDate, , strResult, strFaultReason, strErrMsg)
    
    Exit Function
ChkErr:
    ErrSub "clsExeCommand", "ExeReleaseFixIP"

End Function

'#4276 申請固定IP By Kin 2008/12/26
Private Function ExeRequireFixIP(ByRef rs As ADODB.Recordset, _
                                 ByVal strCMD As String, _
                                 Optional ByVal strSNO As String = Empty, _
                                 Optional ByVal strMediaBillNo As String = Empty, _
                                 Optional ByVal strProcessingDate As String = Empty, _
                                 Optional ByVal strDynIpNum As String = Empty, _
                                 Optional ByVal strFixIPNum As String = Empty, _
                                 Optional ByRef strResult As String, _
                                 Optional ByRef strFaultReason As String, _
                                 Optional ByRef strErrMsg As String)
    On Error GoTo ChkErr
    blnShowMsg = False
    blnNoShowMsg = True
    ExeRequireFixIP = False
    Dim rsReturn As New ADODB.Recordset
    If Val(strFixIPNum) <= 0 Then
        strErrMsg = "固定IP數為0"
        Exit Function
    End If
    ExeRequireFixIP = RequireFixIP(rs, "26", strSNO, strMediaBillNo, strProcessingDate, strDynIpNum, strFixIPNum, rsReturn, strResult, strFaultReason, strErrMsg)
    On Error Resume Next
    Call CloseRecordset(rsReturn)
    Exit Function
ChkErr:
    ErrSub "clsExeCommand", "ExeRequireFixIP"
End Function

'#4276 取消預約FTTX裝機 By Kin 2008/12/25
Private Function ExeReleasePPPOEAccount(ByRef rs As ADODB.Recordset, _
                                        ByVal strCMD As String, _
                                        Optional ByVal strSNO As String = Empty, _
                                        Optional ByVal strMediaBillNo As String = Empty, _
                                        Optional ByVal strProcessingDate As String = Empty, _
                                        Optional ByRef strResult As String, _
                                        Optional ByRef strFaultReason As String, _
                                        Optional ByRef strErrMsg As String) As Boolean
  On Error GoTo ChkErr
    blnShowMsg = False
    blnNoShowMsg = True
    ExeReleasePPPOEAccount = False
    ExeReleasePPPOEAccount = ReleasePPPOEAccount(rs, strCMD, rs("FaciSNo") & "", strSNO, strMediaBillNo, strProcessingDate, strResult, strFaultReason, strErrMsg)
    Exit Function
ChkErr:
    ErrSub "clsExeCommand", "ExeReleasePPPOEAccount"


End Function
 '出庫入庫
Private Function ExeStockCM(ByRef rs As ADODB.Recordset, _
                           ByVal strCMD As String, _
                           Optional ByVal strDevice2 As String, _
                           Optional ByVal strProcessingDate As String = Empty, _
                           Optional ByRef strResult As String, _
                           Optional ByRef strFaultReason As String, _
                           Optional ByRef strErrMsg As String) As Boolean
  On Error GoTo ChkErr
    Dim strArry As String
    blnShowMsg = False
    blnNoShowMsg = True
    ExeStockCM = False
    'strArry = Join(GaryGi, Chr(9))
    'Set gcnGi = cn
    'Call GetGlobal(strArry)
    Select Case strCMD
        Case "40"
            'rs("ModelId") & ""
            ExeStockCM = StockCM(rs, strCMD, rs("ModelId") & "", rs("ModelCode") & "", rs("FaciSNo") & "", _
                                strDevice2, , , strProcessingDate, strResult, strFaultReason, strErrMsg)
        Case "41"
            ExeStockCM = StockCM(rs, strCMD, , , rs("FaciSNo") & "", , , , strProcessingDate, strResult, strFaultReason, strErrMsg)
    End Select
    
    On Error Resume Next
    
    Exit Function
ChkErr:
  ErrSub "clsExeCommand", "ExeStockCM"
End Function
'開機
Private Function ExeOpenCM(ByRef rs As ADODB.Recordset, _
                          ByVal strCMD As String, _
                          Optional ByVal strSNO As String = Empty, _
                          Optional ByVal strMediaBillNo As String = Empty, _
                          Optional ByVal strProcessingDate As String = Empty, _
                          Optional ByRef strResult As String, _
                          Optional ByRef strFaultReason As String, _
                          Optional ByRef strErrMsg As String) As Boolean
  On Error GoTo ChkErr
    blnShowMsg = False
    blnNoShowMsg = True
    ExeOpenCM = False
    
    ExeOpenCM = OpenCM(rs, strCMD, rs("FaciSNo") & "", , strSNO, strMediaBillNo, strProcessingDate, , , strResult, strFaultReason, strErrMsg)
    If ExeOpenCM Then
        Call WriteRecordVodProcedure("命令完成", True)
    End If
    Exit Function
ChkErr:
    ErrSub "clsExeCommand", "ExeOpenCM"
End Function
'關機
Private Function ExeCloseCM(ByRef rs As ADODB.Recordset, _
                          ByVal strCMD As String, _
                          Optional ByVal strSNO As String = Empty, _
                          Optional ByVal strMediaBillNo As String = Empty, _
                          Optional ByVal strProcessingDate As String = Empty, _
                          Optional ByVal strReasonCode As String = Empty, _
                          Optional ByVal strReasonName As String = Empty, _
                          Optional ByRef strResult As String, _
                          Optional ByRef strFaultReason As String, _
                          Optional ByRef strErrMsg As String) As Boolean
  On Error GoTo ChkErr
    blnShowMsg = False
    blnNoShowMsg = True
    ExeCloseCM = False
    ExeCloseCM = OpenCM(rs, strCMD, rs("FaciSNo") & "", , strSNO, strMediaBillNo, strProcessingDate, strReasonCode, strReasonName, strResult, strFaultReason, strErrMsg)
    Exit Function
ChkErr:
    ErrSub "clsExeCommand", "ExeCloseCM"
End Function
'停機
Private Function ExeStopCM(ByRef rs As ADODB.Recordset, _
                          ByVal strCMD As String, _
                          Optional ByVal strSNO As String = Empty, _
                          Optional ByVal strMediaBillNo As String = Empty, _
                          Optional ByVal strProcessingDate As String = Empty, _
                          Optional ByVal strReasonCode As String = Empty, _
                          Optional ByVal strReasonName As String = Empty, _
                          Optional ByRef strResult As String, _
                          Optional ByRef strFaultReason As String, _
                          Optional ByRef strErrMsg As String) As Boolean
  On Error GoTo ChkErr
    'Dim strArry As String
    blnShowMsg = False
    blnNoShowMsg = True
    ExeStopCM = False
    'strArry = Join(GaryGi, Chr(9))
    'Set gcnGi = cn
    'Call GetGlobal(strArry)
    'If Trim(strCMD) <> "13" Then Exit Function
    ExeStopCM = OpenCM(rs, strCMD, rs("FaciSNo") & "", , strSNO, strMediaBillNo, strProcessingDate, strResult, strFaultReason, strErrMsg, strReasonCode, strReasonName)
    Exit Function
ChkErr:
    ErrSub "clsExeCommand", "ExeStopCM"
End Function
'復機
Private Function ExeResumeCM(ByRef rs As ADODB.Recordset, _
                          ByVal strCMD As String, _
                          Optional ByVal strSNO As String = Empty, _
                          Optional ByVal strMediaBillNo As String = Empty, _
                          Optional ByVal strProcessingDate As String = Empty, _
                          Optional ByRef strResult As String, _
                          Optional ByRef strFaultReason As String, _
                          Optional ByRef strErrMsg As String) As Boolean
  On Error GoTo ChkErr
    'Dim strArry As String
    blnShowMsg = False
    blnNoShowMsg = True
    ExeResumeCM = False
    'strArry = Join(GaryGi, Chr(9))
    'Set gcnGi = cn
    'Call GetGlobal(strArry)
    'If Trim(strCMD) <> "14" Then Exit Function
    ExeResumeCM = OpenCM(rs, strCMD, rs("FaciSNo") & "", , strSNO, strMediaBillNo, strProcessingDate, strResult, strFaultReason, strErrMsg)
    Exit Function
ChkErr:
    ErrSub "clsExeCommand", "ExeResumeCM"
End Function
Private Function ExeReleasePort(ByRef rs As ADODB.Recordset, _
                                 ByVal strCMD As String, _
                                 Optional ByVal strSNO As String = Empty, _
                                Optional ByVal strMediaBillNo As String = Empty, _
                                Optional ByVal strProcessingDate As String = Empty, _
                                Optional ByVal strReasonCode As String = Empty, _
                                Optional ByVal strReasonName As String = Empty, _
                                Optional ByRef strResult As String, _
                                Optional ByRef strFaultReason As String, _
                                Optional ByRef strErrMsg As String) As Boolean
    On Error GoTo ChkErr
    blnShowMsg = False
    blnNoShowMsg = True
    ExeReleasePort = False
    ExeReleasePort = OpenCM(rs, strCMD, rs("FaciSNo") & "", IIf(IsCMCP(rs), Get_EMTA_Mac(rs("FaciSNo") & ""), ""), _
                        strSNO, strMediaBillNo, strProcessingDate, strReasonCode, strReasonName, strResult, strFaultReason, strErrMsg)
    Exit Function
ChkErr:
    ErrSub "clsExeCommand", "ExeReleasePort"
                                
End Function




'拆機
Private Function ExeDownCM(ByRef rs As ADODB.Recordset, _
                          ByVal strCMD As String, _
                          Optional ByVal strSNO As String = Empty, _
                          Optional ByVal strMediaBillNo As String = Empty, _
                          Optional ByVal strProcessingDate As String = Empty, _
                          Optional ByVal strReasonCode As String = Empty, _
                          Optional ByVal strReasonName As String = Empty, _
                          Optional ByRef strResult As String, _
                          Optional ByRef strFaultReason As String, _
                          Optional ByRef strErrMsg As String) As Boolean
  On Error GoTo ChkErr
    'Dim strArry As String
    blnShowMsg = False
    blnNoShowMsg = True
    ExeDownCM = False
    'strArry = Join(GaryGi, Chr(9))
    'Set gcnGi = cn
    'Call GetGlobal(strArry)
    'If Trim(strCMD) <> "15" Then Exit Function
    '#4263 參數傳錯,導致ReasonCode與ReasonName沒有填入SO314 By Kin 2008/12/03
    ExeDownCM = OpenCM(rs, strCMD, rs("FaciSNo") & "", IIf(IsCMCP(rs), Get_EMTA_Mac(rs("FaciSNo") & ""), ""), _
                        strSNO, strMediaBillNo, strProcessingDate, strReasonCode, strReasonName, strResult, strFaultReason, strErrMsg)
    Exit Function
ChkErr:
    ErrSub "clsExeCommand", "ExeDownCM"
End Function

'維修換拆
Private Function ExeRepairCM(ByRef rs As ADODB.Recordset, _
                          ByVal strCMD As String, _
                          ByVal strFaciCode As String, _
                          ByVal strFaciSNo As String, _
                          Optional ByVal strSNO As String = Empty, _
                          Optional ByVal strMediaBillNo As String = Empty, _
                          Optional ByVal strProcessingDate As String = Empty, _
                          Optional ByRef strResult As String, _
                          Optional ByRef strFaultReason As String, _
                          Optional ByRef strErrMsg As String) As Boolean
  On Error GoTo ChkErr
    blnShowMsg = False
    blnNoShowMsg = True
    ExeRepairCM = False
    If Trim(strFaciSNo) = "" Then Exit Function
    If Trim(strFaciCode) = "" Then Exit Function
    
    '#3755 chkUseFlag多增公司別參數，因為CD039.Linkmss=0時不判斷UseFlag一定要為1 By Kin 2008/02/01
    If Not chkUseFlag(strFaciSNo, Val(rs("CompCode") & "")) Then Exit Function
    ExeRepairCM = OpenCM(rs, strCMD, strFaciSNo, IIf(IsEMTA(strFaciCode), Get_EMTA_Mac(rs("FaciSNo") & ""), ""), _
                         strSNO, strMediaBillNo, strProcessingDate, strResult, strFaultReason, strErrMsg)
    Exit Function
ChkErr:
    ErrSub "clsExeCommand", "ExeRepairCM"
End Function
'#4276 啟用PPPOE帳號 By Kin 2008/12/25
Private Function ExeOpenPPPOEAccount(ByRef rs As ADODB.Recordset, _
                                  ByVal strCMD As String, _
                                  Optional ByVal strSNO As String = Empty, _
                                  Optional ByVal strMediaBillNo As String = Empty, _
                                  Optional ByVal strProcessingDate As String = Empty, _
                                  Optional ByRef strResult As String, _
                                  Optional ByRef strFaultReason As String, _
                                  Optional ByRef strErrMsg As String) As Boolean
On Error GoTo ChkErr
    Dim rsSO001 As New ADODB.Recordset
    Dim strQry As String
    Dim strCustName As String
    Dim strInstAddress As String
    Dim strCustTel As String
    
    blnShowMsg = False
    blnNoShowMsg = True
    ExeOpenPPPOEAccount = False
    strCustName = Empty
    strInstAddress = Empty
    strCustTel = Empty
    strQry = "Select CustName,InstAddress,TELH1,Tel1 From " & GetOwner & "SO001" & _
            " Where CustId=" & rs("CustId") & " And CompCode=" & rs("CompCode")
    Set rsSO001 = gcnGi.Execute(strQry)
    If Not rsSO001.EOF Then
        strCustName = rsSO001("CustName") & ""
        strInstAddress = rsSO001("InstAddress") & ""
        strCustTel = IIf(rsSO001("TELH1") & "" = "", "", rsSO001("TELH1") & "-") & rsSO001("Tel1") & ""
    End If
    If Trim(strCMD) <> "10" Then Exit Function
    ExeOpenPPPOEAccount = OpenPPPOEAccount(rs, strCMD, strCustName, strCustTel, strInstAddress, _
                                            strProcessingDate, strSNO, strMediaBillNo, strResult, _
                                            strFaultReason, strErrMsg)
    On Error Resume Next
    Call CloseRecordset(rsSO001)
    Exit Function
ChkErr:
    ErrSub "clsExeCommand", "ExeOpenPPPOEAccount"
    
    
End Function

'裝機
Private Function ExeInstallCM(ByRef rs As ADODB.Recordset, _
                             ByVal strCMD As String, _
                             Optional ByVal strSNO As String = Empty, _
                             Optional ByVal strMediaBillNo As String = Empty, _
                             Optional ByVal strProcessingDate As String = Empty, _
                             Optional ByRef strResult As String, _
                             Optional ByRef strFaultReason As String, _
                             Optional ByRef strErrMsg As String) As Boolean
  On Error GoTo ChkErr
    Dim rsSO001 As New ADODB.Recordset
    Dim rsReturn As New ADODB.Recordset
    Dim strQry As String
    Dim strCustName As String
    Dim strInstAddress As String
    Dim strCustTel As String
    
    blnShowMsg = False
    blnNoShowMsg = True
    ExeInstallCM = False
    strCustName = Empty
    strInstAddress = Empty
    strCustTel = Empty
    'strArry = Join(GaryGi, Chr(9))
    'Set gcnGi = cn
    'Call GetGlobal(strArry)
    'If Trim(strCMD) <> "10" Then Exit Function
    strQry = "Select CustName,InstAddress,TELH1,Tel1 From " & GetOwner & "SO001" & _
            " Where CustId=" & rs("CustId") & " And CompCode=" & rs("CompCode")
    Set rsSO001 = gcnGi.Execute(strQry)
    If Not rsSO001.EOF Then
        strCustName = rsSO001("CustName") & ""
        strInstAddress = rsSO001("InstAddress") & ""
        strCustTel = IIf(rsSO001("TELH1") & "" = "", "", rsSO001("TELH1") & "-") & rsSO001("Tel1") & ""
        'strCustTel = rsSO001("Tel1") & ""
    End If
    If Trim(strCMD) <> "10" Then Exit Function
    ExeInstallCM = InstCM(rs, strCMD, strCustName, strCustTel, strInstAddress, _
                    rs("CMBaudNo") & "", rs("VendorCode") & "", _
                    IIf(IsCMCP(rs), Get_EMTA_Mac(rs("FaciSNo") & ""), ""), strSNO, strMediaBillNo, _
                    strProcessingDate, rsReturn, strResult, strFaultReason, strErrMsg)
    On Error Resume Next
    Call CloseRecordset(rsSO001)
    Call CloseRecordset(rsReturn)
    Exit Function
ChkErr:
    ErrSub "clsExeCommand", "ExeInstallCM"
End Function

'變更申請人[21]
Private Function ExeChangeCust(ByRef rs As ADODB.Recordset, _
                             ByVal strCMD As String, _
                             Optional ByVal strSNO As String = Empty, _
                             Optional ByVal strMediaBillNo As String = Empty, _
                             Optional ByVal strProcessingDate As String = Empty, _
                             Optional ByRef strResult As String, _
                             Optional ByRef strFaultReason As String, _
                             Optional ByRef strErrMsg As String) As Boolean
  On Error GoTo ChkErr
    Dim rsSO001 As New ADODB.Recordset
    Dim rsReturn As New ADODB.Recordset
    Dim strQry As String
    Dim strCustName As String
    Dim strInstAddress As String
    Dim strCustTel As String
    blnShowMsg = False
    blnNoShowMsg = True
    ExeChangeCust = False
    strCustName = Empty
    strInstAddress = Empty
    strCustTel = Empty
    strQry = "Select CustName,InstAddress,TELH1,Tel1 From " & GetOwner & "SO001" & _
            " Where CustId=" & rs("CustId") & " And CompCode=" & rs("CompCode")
    Set rsSO001 = gcnGi.Execute(strQry)
    If Not rsSO001.EOF Then
        strCustName = rsSO001("CustName") & ""
        strInstAddress = rsSO001("InstAddress") & ""
        'strCustTel = rsSO001("Tel1") & ""
        strCustTel = IIf(rsSO001("TELH1") & "" = "", "", rsSO001("TELH1") & "-") & rsSO001("Tel1") & ""
    End If
    ExeChangeCust = False
    ExeChangeCust = ChangeCust(rs, "21", strCustName, strCustTel, strInstAddress, IIf(blnUseFttx, Empty, rs("CMBaudNo") & ""), _
                    IIf(blnUseFttx, Empty, rs("VendorCode") & ""), IIf(blnUseFttx, Empty, IIf(IsCMCP(rs), Get_EMTA_Mac(rs("FaciSNo")), "")), _
                    strSNO, strMediaBillNo, strProcessingDate, rs("DialAccount") & "", rs("DialPassword") & "", _
                    rs("ID") & "", rs("DeclarantName") & "", rs("Birthday") & "", rsReturn, strResult, strFaultReason, strErrMsg)
    Call CloseRecordset(rsSO001)
    Call CloseRecordset(rsReturn)
    Exit Function
ChkErr:
    ErrSub "clsExeCommand", "ExeChangeCust"
End Function
'#4276 預約FTTX裝機 By Kin 2008/12/29
Private Function ExeRequirePPPOEAccount(ByRef rs As ADODB.Recordset, _
                                        ByVal strCMD As String, _
                                        Optional ByVal strSNO As String = Empty, _
                                        Optional ByVal strMediaBillNo As String = Empty, _
                                        Optional ByVal strProcessingDate As String = Empty, _
                                        Optional ByVal strDynIpNum As String = Empty, _
                                        Optional ByVal strFixIPNum As String = Empty, _
                                        Optional ByRef strResult As String, _
                                        Optional ByRef strFaultReason As String, _
                                        Optional ByRef strErrMsg As String) As Boolean
                                        
  On Error GoTo ChkErr
    Dim rsReturn As New ADODB.Recordset
    blnShowMsg = False
    blnNoShowMsg = True
    ExeRequirePPPOEAccount = False
    ExeRequirePPPOEAccount = RequirePPPOEAccount(rs, "18", strSNO, strMediaBillNo, strProcessingDate, strDynIpNum, strFixIPCount, rsReturn, strResult, strFaultReason, strErrMsg)
    On Error Resume Next
    Call CloseRecordset(rsReturn)
    Exit Function
ChkErr:
  ErrSub "clsExeCommand", "ExeRequirePPPOEAccount"
End Function
'變更速率、變更路由
Private Function ExeChangeRate(ByRef rs As ADODB.Recordset, _
                             ByVal strCMD As String, Optional ByVal strBaudRate As String = Empty, _
                             Optional ByVal strBaudName As String = Empty, _
                             Optional ByVal strVendorCode As String = Empty, _
                             Optional ByVal strVendorName As String = Empty, _
                             Optional ByVal strSNO As String = Empty, _
                             Optional ByVal strMediaBillNo As String = Empty, _
                             Optional ByVal strProcessingDate As String = Empty, _
                             Optional ByRef strResult As String, _
                             Optional ByRef strFaultReason As String, _
                             Optional ByRef strErrMsg As String) As Boolean
  On Error GoTo ChkErr
    Dim rsReturn As New ADODB.Recordset
    blnShowMsg = False
    blnNoShowMsg = True
    ExeChangeRate = False
    Select Case strCMD
        Case "22"
            If Not blnExeCommand Then
                If Not blnExeCommand Then
                    strBaudRate = gcnGi.Execute("Select SchemaCode1 From " & strCMowner & "SO314 Where CmdSeqNo=" & strCmdSeqNo & _
                                                " And CompCode=" & rs("CompCode"))(0) & ""
                    If strBaudRate <> "" Then
                        strBaudName = gcnGi.Execute("Select Description From " & GetOwner & "CD064 Where CodeNo=" & strBaudRate)(0) & ""
                    End If
                End If
            End If
            ExeChangeRate = ChangeRate(rs, "22", strBaudRate, rs("VendorCode") & "", , strBaudRate, _
                                            strBaudName, strSNO, strMediaBillNo, strProcessingDate, , _
                                            strResult, strFaultReason, strErrMsg)
        Case "23"
            If Not blnExeCommand Then
                If Not blnExeCommand Then
                    strVendorCode = gcnGi.Execute("Select SchemaCode2 From " & strCMowner & "SO314 Where CmdSeqNo=" & strCmdSeqNo & _
                                                    " And CompCode=" & rs("CompCode"))(0) & ""
                    If strVendorCode <> Empty Then
                        strVendorName = gcnGi.Execute("Select VendorName From " & GetOwner & "CM006 Where VendorCode='" & strVendorCode & "'" & _
                                                    " And CompCode=" & rs("CompCode"))(0) & ""
                    End If
                End If
            End If
            ExeChangeRate = ChangeRate(rs, strCMD, rs("CMBaudNo") & "", strVendorCode, strVendorName, , , strSNO, strMediaBillNo, strProcessingDate, IIf(blnUseFttx, rsReturn, Nothing), strResult, strFaultReason, strErrMsg)
    End Select
    On Error Resume Next
    Call CloseRecordset(rsReturn)
    Exit Function
ChkErr:
    ErrSub "clsExeCommand", "ExeChangeRate"
End Function

'變更基本資料
Private Function ExeChangeData(ByRef rs As ADODB.Recordset, _
                             ByVal strCMD As String, _
                             Optional ByVal strSNO As String = Empty, _
                             Optional ByVal strMediaBillNo As String = Empty, _
                             Optional ByVal strProcessingDate As String = Empty, _
                             Optional ByRef strResult As String, _
                             Optional ByRef strFaultReason As String, _
                             Optional ByRef strErrMsg As String) As Boolean
  On Error GoTo ChkErr
    Dim rsSO001 As New ADODB.Recordset
    Dim strQry As String
    Dim strCustName As String
    Dim strInstAddress As String
    Dim strCustTel As String
    blnShowMsg = False
    blnNoShowMsg = True
    ExeChangeData = False
    strCustName = Empty
    strInstAddress = Empty
    strCustTel = Empty
    strQry = "Select CustName,InstAddress,TELH1,Tel1 From " & GetOwner & "SO001" & _
            " Where CustId=" & rs("CustId") & " And CompCode=" & rs("CompCode")
    Set rsSO001 = gcnGi.Execute(strQry)
    If Not rsSO001.EOF Then
        strCustName = rsSO001("CustName") & ""
        strInstAddress = rsSO001("InstAddress") & ""
        strCustTel = IIf(rsSO001("TELH1") & "" = "", "", rsSO001("TELH1") & "-") & rsSO001("Tel1") & ""
        'strCustTel = rsSO001("Tel1") & ""
    End If
    ExeChangeData = ChangeData(rs, "20", strCustName, strCustName, strInstAddress, , strSNO, strMediaBillNo, strProcessingDate, strResult, strFaultReason, strErrMsg)
    Call CloseRecordset(rsSO001)
    Exit Function
ChkErr:
  ErrSub "clsExeCommand", "ChangeData"
End Function

'變更密碼
Private Function ExeChangePassword(ByRef rs As ADODB.Recordset, _
                             ByVal strCMD As String, _
                             ByVal strPassword As String, _
                             Optional ByVal strSNO As String = Empty, _
                             Optional ByVal strMediaBillNo As String = Empty, _
                             Optional ByVal strProcessingDate As String = Empty, _
                             Optional ByRef strResult As String, _
                             Optional ByRef strFaultReason As String, _
                             Optional ByRef strErrMsg As String) As Boolean
  On Error GoTo ChkErr
    blnShowMsg = False
    blnNoShowMsg = True
    ExeChangePassword = False
    If Not blnExeCommand Then
        strPassword = gcnGi.Execute("Select AccountPassword From " & strCMowner & "SO314 Where CmdSeqNo=" & strCmdSeqNo & _
                                                " And CompCode=" & rs("CompCode"))(0) & ""
    End If
    ExeChangePassword = ChangePassword(rs, "24", strPassword, strSNO, strMediaBillNo, strProcessingDate, strResult, strFaultReason, strErrMsg)
    
    
    Exit Function
ChkErr:
  ErrSub "clsExeCommand", "ExeChangePassword"
End Function

'重置設備
Private Function ExeRest(ByRef rs As ADODB.Recordset, _
                         ByVal strCMD As String, _
                         Optional ByVal strSNO As String = Empty, _
                         Optional ByVal strMediaBillNo As String = Empty, _
                         Optional ByVal strProcessingDate As String = Empty, _
                         Optional ByRef strResult As String, _
                         Optional ByRef strFaultReason As String, _
                         Optional ByRef strErrMsg As String) As Boolean
  On Error GoTo ChkErr
    blnShowMsg = False
    blnNoShowMsg = True
    ExeRest = False
    ExeRest = OpenCM(rs, strCMD, rs("FaciSNo") & "", , strSNO, strMediaBillNo, strProcessingDate, strResult, strFaultReason, strErrMsg)
    Exit Function
ChkErr:
    ErrSub "clsExeCommand", "ExeRest"
End Function

'確認使用人
Private Function ExeConfirmCust(ByRef rs As ADODB.Recordset, _
                                ByVal strCMD As String, _
                                Optional ByVal strSNO As String = Empty, _
                                Optional ByVal strMediaBillNo As String = Empty, _
                                Optional ByVal strProcessingDate As String = Empty, _
                                Optional ByRef strResult As String, _
                                Optional ByRef strFaultReason As String, _
                                Optional ByRef strErrMsg As String) As Boolean

 On Error GoTo ChkErr
    Dim strInstAddress As String
    Dim strCustTel As String
    Dim strCustName As String
    Dim strQry As String
    Dim rsSO001 As ADODB.Recordset
    blnShowMsg = False
    blnNoShowMsg = True
    ExeConfirmCust = False
    strInstAddress = Empty
    strCustName = Empty
    strCustTel = Empty
    strQry = Empty
    strQry = "Select CustName,InstAddress,TELH1,Tel1 From " & GetOwner & "SO001" & _
            " Where CustId=" & rs("CustId") & " And CompCode=" & rs("CompCode")
    Set rsSO001 = gcnGi.Execute(strQry)
    If Not rsSO001.EOF Then
        strCustName = rsSO001("CustName") & ""
        strInstAddress = rsSO001("InstAddress") & ""
        strCustTel = IIf(rsSO001("TELH1") & "" = "", "", rsSO001("TELH1") & "-") & rsSO001("Tel1") & ""
        'strCustTel = rsSO001("Tel1") & ""
    End If
    ExeConfirmCust = ConfirmCust(rs, strCMD, strCustName, strCustTel, strInstAddress, strSNO, strMediaBillNo, strProcessingDate, strResult, strFaultReason, strErrMsg)
    Call CloseRecordset(rsSO001)
    Exit Function
ChkErr:
    ErrSub "clsExeCommand", "ExeConfirmCust"
End Function


Private Sub Class_Terminate()
  On Error Resume Next
    objRecordWrite.Close
    Set objRecordWrite = Nothing
    Set objFileSystem = Nothing
End Sub
