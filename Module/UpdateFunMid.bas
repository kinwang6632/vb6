Attribute VB_Name = "UpdateFunMid"
Option Explicit
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Const gMsgIsDataOK = "此為必要欄位,須有值 !! "
Public lngCount As Long           '記錄成功筆數
Public lngErrCount As Long        '記錄錯誤筆數
Public lngStatusCount As Long     '記錄非正常戶筆數
Public lngTime As Long             '記錄所花時間
Public strNowTime As String        '記錄目前時間
Public FSO As New FileSystemObject
Public File As TextStream
Public ErrFile As TextStream     '錯誤資料log檔
Private intPara3 As Integer
Private intPara5 As Integer
Private intPara8 As Integer
Private intPara9 As Integer
Public strUserName As String
Private strFilePath As String
Public objStorePara As clsStoreParameter
Public strOwnerName As String
Public blnFailUpd As Boolean
Public strUCCode As String
Public strUCName As String
Public blnCrossCustCombine  As Boolean
'Public GetOwner As String

'接受上層傳的檔案路徑
Public Function GetFilePath(ByVal vFilePath As String)
  On Error GoTo ChkErr
     strFilePath = vFilePath
  Exit Function
ChkErr:
  ErrSub "DealFun", "GetFilePath"
End Function

Public Sub ErrLog(ByVal ErrString As String)
    On Error Resume Next
    Dim LogFile As String
    Dim fnum As Integer
    
    'LogFile = ReadFromINI(GetIniPath1, "GICMISPath", "ErrLogPath")
    LogFile = strFilePath
    If Not ChkDirExist(LogFile) Then MkDir LogFile
    
    If Right(LogFile, 1) <> "\" Then LogFile = LogFile & "\"
    LogFile = LogFile & "SYSERR.TXT"
    
    ErrString = Replace(ErrString, vbCrLf & vbCrLf, vbCrLf)
    ErrString = "發生錯誤的系統 : 冠崴客服營運管理系統" & vbCrLf & _
                "發生錯誤的電腦 : " & CreateObject("WScript.NetWork").ComputerName & vbCrLf & _
                "使用者名稱 : " & strUserName & vbCrLf & ErrString
    
    fnum = FreeFile
    Open LogFile For Append As fnum
    Print #fnum, ErrString & vbCrLf & String(66, "-") & vbCrLf
    Close fnum
End Sub

Public Function ChkDirExist(FileName As String) As Boolean
  On Error GoTo ChkErr
    Dim varValue As Variant
    If Right(FileName, 1) = "\" Then FileName = Left(FileName, Len(FileName) - 1)
    FileName = UCase(FileName)
    varValue = Split(FileName, "\")
    If UCase(Dir(FileName, vbDirectory)) = varValue(UBound(varValue)) Then
        ChkDirExist = True
    Else
        ChkDirExist = False
    End If
  Exit Function
ChkErr:
    Call ErrSub("UpdateFunMid", "Function ChkDirExist")
End Function

Public Function Decrypt(EncryptionString As String) As String
  On Error GoTo ChkErr
    Decrypt = CreateObject("DevPowerEncrypt.EnCrypt").Decrypt(EncryptionString)
  Exit Function
ChkErr:
    Call ErrSub("UpdateFunMid", "Function Decrypt")
    
End Function

'#3527 增加參數intRefNo(CD018.REFNO)，如果為2，則使用POS3新的流程，目前只給mediaPost3使用 By Kin 2007/10/05
'#3417 增加參數vBankId，如果收費人員有值，則SO033的備註欄需填入銀行代碼
Public Function ChkData(ByVal vBillNo As String, ByVal vRealAmt As String, _
                        ByVal vClctEn As String, ByVal vClctName As String, _
                        ByVal vRealDate As String, ByVal vUpdEn As String, _
                        ByVal vCMCode As String, ByVal vCMName As String, _
                        ByVal vPTCode As String, ByVal vPTName As String, _
                        ByVal vCM As String, _
                        Optional blnBillNO As Boolean = False, _
                        Optional strCitemQry As String, _
                        Optional strMedia As String, _
                        Optional intRefNo As Integer = 0, _
                        Optional ByVal vBankId As String = Empty) As Boolean
  On Error GoTo ChkErr
  Dim strSQL As String
  Dim rsChk As New ADODB.Recordset
  Dim rsCust As New ADODB.Recordset
  Dim strSQL2 As String
  Dim strBillNo_Old As String
  Dim rsTmp As New ADODB.Recordset
  Dim strReturn As String
  Dim strWhere As String, lngBudgetPeriod As Long
  Dim strFailUpdSQL As String
  Dim strFail As String
  Dim strCD013RefNo As String
  Dim lngTotalRealAmt As Long
  Dim lngTotalShouldAmt As Long
  Dim strCustId As String
    '#5564 CD013.PAYOK=1或CD013.REFNO=3,7都算櫃台已收 By Kin 2010/05/19
    strCD013RefNo = " MAX(DECODE(NVL(CD013.PAYOK,0),1,3," & _
                "DECODE(NVL(CD013.REFNO,0),3,3,DECODE(NVL(CD013.REFNO,0),7,3,0)))) REFNO "
    strFail = Empty
    '收費項目
    If strCitemQry <> "" Then strSQL2 = " And SO033.CitemCode In (" & strCitemQry & ")"
    '7179
    blnCrossCustCombine = Val(GetRsValue("Select Nvl(CrossCustCombine,0) From " & strOwnerName & "SO041", gcnGi) & "")
    
    '***********************************************************************************************************************************************************************************************************
    '#3527 判斷是否使用參考號為2，如果為2則使用新的POS3流程 By Kin 2007/10/05
    If intRefNo <> 2 Then
        strWhere = " SO033.UCCODE IS NOT NULL AND " & IIf(blnBillNO, "SO033.BillNo", "SO033.MediaBillNo") & " = '" & vBillNo & "' And SO033.CancelFlag <> 1 " & strSQL2
        
    Else
        strWhere = " SO033.UCCODE IS NOT NULL AND SO033.CitibankATM='" & vBillNo & "' And SO033.CancelFlag <> 1 " & strSQL2
    End If
    '***********************************************************************************************************************************************************************************************************
    
    '***************************************************************************************************************
    '#3417 失敗時要更新的語法 By Kin 2007/12/10
    '#4388 增加更新異動人員與異動時間 By Kin 2009/05/04
    If blnFailUpd Then
        strFailUpdSQL = "Update " & strOwnerName & "SO033 Set UCCode=" & strUCCode & _
                        ",UPDTime='" & objStorePara.uUPDTime & "',UPDEN='" & objStorePara.UpdEn & "'" & _
                        ",UCName='" & strUCName & "',Note=Note"
                                                                                  
    End If
    '***************************************************************************************************************
    '#4010 如果UCcode的參考號為3要視為異常 By Kin 2008/07/29
    '#5564 CD013.PAYOK=1或CD013.REFNO=3,7都算櫃台已收 By Kin 2010/05/19
    If intRefNo <> 2 Then
        If strMedia = 0 Then
            If blnBillNO Then
                strSQL = "Select Nvl(sum(ShouldAmt),0) as TotalShouldAmt, Nvl(sum(RealAmt),0) as TotalRealAmt," & _
                        "SO033.RealDate, SO033.CustId, SO033.ServiceType," & strCD013RefNo & " From " & strOwnerName & "SO033," & strOwnerName & "CD013 " & _
                        " Where SO033.UCCODE IS NOT NULL AND SO033.BillNo = '" & vBillNo & "' And SO033.CancelFlag <> 1 And SO033.UCCode=CD013.CodeNo " & _
                        strSQL2 & " Group By SO033.CustId,SO033.RealDate,SO033.ServiceType "
            Else
                strSQL = "Select Nvl(sum(SO033.ShouldAmt),0) as TotalShouldAmt, Nvl(sum(SO033.RealAmt),0) as TotalRealAmt," & _
                         "SO033.RealDate,SO033.CustId,SO033.ServiceType," & strCD013RefNo & " From " & strOwnerName & "SO033," & strOwnerName & "CD013 " & _
                         " Where SO033.UCCODE IS NOT NULL AND SO033.MediaBillNo = '" & vBillNo & "' And SO033.CancelFlag <> 1 And SO033.UCCode=CD013.CodeNo " & _
                         strSQL2 & " Group By SO033.CustId,SO033.RealDate,SO033.ServiceType "
                
            End If
        Else
        
            If blnBillNO Then
                strSQL = "Select Nvl(sum(SO033.ShouldAmt),0) as TotalShouldAmt, Nvl(sum(SO033.RealAmt),0) as TotalRealAmt," & _
                         "SO033.RealDate, SO033.CustId," & strCD013RefNo & " From " & strOwnerName & "SO033," & strOwnerName & "CD013" & _
                         " Where SO033.UCCODE IS NOT NULL AND SO033.BillNo = '" & vBillNo & "' And SO033.CancelFlag <> 1 And SO033.UCCode=CD013.CodeNo " & _
                         strSQL2 & " Group By SO033.CustId,SO033.RealDate"
            Else
                strSQL = "Select Nvl(sum(SO033.ShouldAmt),0) as TotalShouldAmt, Nvl(sum(SO033.RealAmt),0) as TotalRealAmt," & _
                         "SO033.RealDate,SO033.CustId,SO033.MediaBillNo," & strCD013RefNo & " From " & strOwnerName & "SO033," & strOwnerName & "CD013" & _
                         " Where SO033.UCCODE IS NOT NULL AND SO033.MediaBillNo = '" & vBillNo & "' And SO033.CancelFlag <> 1 And SO033.UCcode=CD013.CodeNo " & _
                         strSQL2 & " Group By SO033.CustId,SO033.RealDate,SO033.MediaBillNo"
                 
            End If
        End If
    Else
        If vBillNo <> "" Then
            strSQL = "Select Nvl(sum(SO033.ShouldAmt),0) as TotalShouldAmt, Nvl(sum(SO033.RealAmt),0) as TotalRealAmt," & _
                     "SO033.RealDate,SO033.CustId," & strCD013RefNo & " From " & strOwnerName & "SO033," & strOwnerName & "CD013" & _
                     " Where SO033.UCCODE IS NOT NULL AND SO033.CitibankATM = '" & vBillNo & "' And SO033.CancelFlag <> 1 And SO033.UCCode=CD013.CodeNo " & strSQL2 & _
                     " Group By SO033.CustId,SO033.RealDate"
        Else
            ErrFile.Write ("虛擬帳號不正確！！")
            lngErrCount = lngErrCount + 1
            ChkData = False
            Exit Function
        End If
    End If
    '7179
    If blnCrossCustCombine Then
        strSQL = "Select A.*,SO001.AMduid, (Select MainCustId From " & strOwnerName & "SO202 Where MduId = SO001.AmduId) MainCustId " & _
                             " From (" & strSQL & " ) A," & strOwnerName & "SO001 " & _
                             " Where A.CustId = SO001.CustId "
    End If
    '*********************************************************************************************************************************************************************************************************************************************************************************************************************
    If rsChk.State = adStateOpen Then rsChk.Close
  
    If Not GetRS(rsChk, strSQL, gcnGi) Then
       ChkData = False
       Exit Function
    End If
    
    If rsChk.EOF Or rsChk.BOF Then
       ErrFile.WriteLine ("單據編號： " & vBillNo & " 查無此單號！" & "  代收收費處： " & vCM)
'       strFail = strFailUpdSQL & "||'單據編號： " & vBillNo & _
'                                 "  查無此單號！" & _
'                                 "  代收收費處： " & vCM & "'" & _
'                                 " Where " & strWhere
'       gcnGi.Execute strFail
       lngErrCount = lngErrCount + 1
       ChkData = True
       Exit Function
    Else
        '7179*********************************************************************************************
        lngTotalRealAmt = 0
        lngTotalShouldAmt = 0
        If blnCrossCustCombine Then
            rsChk.MoveFirst
            If Len(rsChk("MainCustId") & "") > 0 Then
                strCustId = rsChk("MainCustId") & ""
            Else
                strCustId = rsChk("CustId") & ""
            End If
            Do While Not rsChk.EOF
                lngTotalRealAmt = lngTotalRealAmt + Val(rsChk("TotalRealAmt") & "")
                lngTotalShouldAmt = lngTotalShouldAmt + Val(rsChk("TotalShouldAmt") & "")
                rsChk.MoveNext
            Loop
            rsChk.MoveFirst
        Else
            lngTotalRealAmt = Val(rsChk("TotalRealAmt") & "")
            lngTotalShouldAmt = Val(rsChk("TotalShouldAmt") & "")
            strCustId = rsChk("CustId") & ""
        End If
        '*************************************************************************************************
        '***********************************************************************
        '#5386 判斷是否有不同的客編但媒体單號相同 By Kin 2009/12/04
        If rsChk.RecordCount > 1 Then
            Dim strRetMsg As String
            strRetMsg = ChkDiffCust(rsChk)  '檢核不同的客編
            If strRetMsg <> "" Then
                ErrFile.WriteLine ("單據編號： " & vBillNo & " 相同單據編號但客戶編號不同！ 客編：" & strRetMsg)
                If blnFailUpd Then
                    strFail = strFailUpdSQL & "||'單據編號：" & vBillNo & _
                            " 相同單據編號但客戶編號不同！ 客編：" & strRetMsg & "'" & _
                            " Where " & strWhere
                    gcnGi.Execute strFail
                
                End If
                lngErrCount = lngErrCount + 1
                ChkData = True
                Exit Function
            End If
        End If
        '***********************************************************************
        
        '**************************************************************************************************
        '#4010 UCCode的參考號為3要視為櫃臺已收 By Kin 2007/07/29
        If Val(rsChk("RefNo")) = 3 Then
            ErrFile.WriteLine ("單據編號： " & vBillNo & " 櫃臺已收請查核！" & "  代收收費處： " & vCM)
            If blnFailUpd Then
                strFail = strFailUpdSQL & "||'單據編號： " & vBillNo & _
                                " 櫃臺已收請查核！ " & _
                                "  代收收費處： " & vCM & "'" & _
                                " Where " & strWhere
                
                gcnGi.Execute strFail
            End If
            lngErrCount = lngErrCount + 1
            ChkData = True
            Exit Function
        End If
        '***************************************************************************************************
        '*********************************************************************************************************************************************
        '#3527 判斷是否使用參考號為2，如果為2則使用新的POS3流程 By Kin 2007/10/05
        If intRefNo <> 2 Then
          '找出舊單號 2003/08/27
            '7179
            If blnBillNO = False Then
               strBillNo_Old = GetRsValue("Select BillNo From " & strOwnerName & "SO033 Where MediaBillNO='" & vBillNo & "'" & _
                                           IIf(blnCrossCustCombine, " And Custid = " & strCustId, ""), gcnGi) & ""
            Else
               strBillNo_Old = vBillNo
            End If
        Else
            '7179
            strBillNo_Old = GetRsValue("Select BillNo From " & strOwnerName & "SO033 Where CitibankATM='" & vBillNo & "'" & _
                              IIf(blnCrossCustCombine, " And Custid = " & strCustId, ""), gcnGi) & ""
        End If
        '**********************************************************************************************************************************************
        
       If rsChk("RealDate").Value & "" <> "" And lngTotalRealAmt > 0 Then
            'ErrFile.WriteLine ("單據編號： " & vBillNo & " 檔案內已實收金額$ " & rsChk("TotalRealAmt").Value & " ，請查核！" & "  代收收費處： " & vCM)
            ErrFile.WriteLine ("單據編號： " & vBillNo & " 檔案內已實收金額$ " & lngTotalRealAmt & " ，請查核！" & "  代收收費處： " & vCM)
            
            '********************************************************************************************************************************************************************************
            '#3417 錯誤時更新UCCode和UCName By Kin 2007/12/10
            If blnFailUpd Then
                
'                strFail = strFailUpdSQL & "||'單據編號： " & vBillNo & _
'                                " 檔案內已實收金額$ " & rsChk("TotalRealAmt").Value & _
'                                " ，請查核！" & "  代收收費處： " & vCM & "'" & _
'                                " Where " & strWhere
                 strFail = strFailUpdSQL & "||'單據編號： " & vBillNo & _
                                " 檔案內已實收金額$ " & lngTotalRealAmt & _
                                " ，請查核！" & "  代收收費處： " & vCM & "'" & _
                                " Where " & strWhere
                gcnGi.Execute strFail
            End If
            '*******************************************************************************************************************************************************************************
            
            lngErrCount = lngErrCount + 1
            ChkData = True
            Exit Function
       Else
          '#3122 判斷扣款日期是否大於日結日期
          If Not ChkOverCloseDate(Format(vRealDate, "####/##/##"), objStorePara.CompCode, objStorePara.ServiceType, strReturn) Then
                ErrFile.WriteLine ("單據編號：" & vBillNo & "  實際扣款日：" & Format(str(Val(vRealDate) - 19110000), "##/##/##") & "  錯誤原因：" & strReturn)
                
                '***************************************************************************************************************************************************************************************
                '#3417 錯誤時更新UCCode和UCName By Kin 2007/12/10
                If blnFailUpd Then
                    
                    strFail = strFailUpdSQL & "||'單據編號：" & vBillNo & _
                                    "  實際扣款日：" & Format(str(Val(vRealDate) - 19110000), "##/##/##") & _
                                    "  錯誤原因：" & strReturn & "'" & _
                                    " Where " & strWhere
                    gcnGi.Execute strFail
                End If
                '****************************************************************************************************************************************************************************************
                
                lngErrCount = lngErrCount + 1
                ChkData = True
                Exit Function
          End If
          If lngTotalShouldAmt <> Val(vRealAmt) Then
                'ErrFile.WriteLine ("單據編號： " & vBillNo & " 金額不合，檔案應收金額$ " & rsChk("TotalShouldAmt").Value & " ，磁片已收金額$ " & Val(vRealAmt) & "  代收收費處： " & vCM & "  收費單號：" & strBillNo_Old)
                ErrFile.WriteLine ("單據編號： " & vBillNo & " 金額不合，檔案應收金額$ " & lngTotalShouldAmt & " ，磁片已收金額$ " & Val(vRealAmt) & "  代收收費處： " & vCM & "  收費單號：" & strBillNo_Old)
                '*********************************************************************************************************************************************************************************************************************************************
                '#3417 錯誤時更新UCCode和UCName By Kin 2007/12/10
                If blnFailUpd Then
                    
'                    strFail = strFailUpdSQL & "||'單據編號： " & vBillNo & _
'                                    " 金額不合，檔案應收金額$ " & rsChk("TotalShouldAmt").Value & _
'                                    " ，磁片已收金額$ " & Val(vRealAmt) & "  代收收費處： " & vCM & _
'                                    "  收費單號：" & strBillNo_Old & "'" & _
'                                    " Where " & strWhere
                    strFail = strFailUpdSQL & "||'單據編號： " & vBillNo & _
                                    " 金額不合，檔案應收金額$ " & lngTotalShouldAmt & _
                                    " ，磁片已收金額$ " & Val(vRealAmt) & "  代收收費處： " & vCM & _
                                    "  收費單號：" & strBillNo_Old & "'" & _
                                    " Where " & strWhere
                    gcnGi.Execute strFail
                End If
                '*********************************************************************************************************************************************************************************************************************************************
                
                lngErrCount = lngErrCount + 1
                ChkData = True
                Exit Function
          Else
                If rsCust.State = adStateOpen Then rsCust.Close
                
                '***************************************************************************************************************************************************************************************************************************************************************************
                '#3527 判斷是否使用參考號為2，如果為2則使用新的POS3流程 By Kin 2007/10/05
                If intRefNo <> 2 Then
                    If strMedia = 0 Then
                        strSQL = "Select CustStatusCode,WipCode3,ServiceType From " & strOwnerName & "So002 Where CustId = " & rsChk("CustId").Value & " And ServiceType= '" & rsChk("ServiceType") & "'"
                    Else
                        If blnBillNO Then
                            strSQL = "Select CustStatusCode,WipCode3,ServiceType From " & strOwnerName & "So002 Where CustId = " & rsChk("CustId").Value & " And ServiceType= '" & rsChk("ServiceType") & "'"
                        Else
                            strSQL = "Select SO002.CustStatusCode,SO002.WipCode3,So002.ServiceType From " & strOwnerName & "So002," & strOwnerName & "SO033 Where SO002.ServiceType=SO033.ServiceType And SO002.CustId = SO033.CustId And MediaBillNo = '" & rsChk("MediaBillNo").Value & "'"
                        End If
                    End If
                Else
                    strSQL = "Select SO002.CustStatusCode,SO002.WipCode3,So002.ServiceType From " & strOwnerName & "So002," & strOwnerName & "SO033 Where SO002.ServiceType=SO033.ServiceType And SO002.CustId = SO033.CustId And CitibankATM = '" & vBillNo & "'"
                End If
                '***************************************************************************************************************************************************************************************************************************************************************************
                '7179
                If blnCrossCustCombine Then
                    'strSQL = strSQL & " And SO033.Custid = " & strCustId
                End If
                Call GetRS(rsCust, strSQL, gcnGi)
                While Not rsCust.EOF
                    '非正常客戶則先作log再入帳
                    If rsCust("CustStatusCode").Value & "" <> 1 Then
                        ErrFile.WriteLine ("客戶編號: " & rsChk("CustId").Value & " 單據編號： " & vBillNo & " 為非正常客戶 " & " 客戶狀態: " & rsCust("CustStatusCode").Value & " 代收收費處: " & vCM & " 服務別: " & rsCust("ServiceType").Value)
                        lngStatusCount = lngStatusCount + 1
                    '正常客戶且CM關機中
                    ElseIf rsCust("CustStatusCode").Value & "" = 1 And rsCust("WipCode3") = 16 Then
                        ErrFile.WriteLine ("客戶編號: " & rsChk("CustId").Value & " 單據編號： " & vBillNo & " 為CM關機中客戶 " & " 代收收費處: " & vCM)
                        lngStatusCount = lngStatusCount + 1
                    End If
                    rsCust.MoveNext
                Wend
                '7179
                If Not GetRS(rsTmp, "Select RealPeriod,CustId,CompCode,ServiceType,CitemCode,FaciSeqNo From " & strOwnerName & "So033 " & _
                                " Where " & strWhere) Then Exit Function
                '94/09/30 1805 Jacky
                    Do While Not rsTmp.EOF
                        lngBudgetPeriod = Val(GetRsValue("Select Nvl(BudgetPeriod,0) From " & strOwnerName & "SO003 Where Custid = " & rsTmp("CustId") & " And CompCode = " & rsTmp("CompCode") & " And ServiceType = '" & rsTmp("ServiceType") & "' " & _
                                        " And CitemCode = " & rsTmp("CitemCode") & " And FaciSeqNo = '" & rsTmp("FaciSeqNo") & "'") & "")
                        If rsTmp("RealPeriod") > lngBudgetPeriod And lngBudgetPeriod > 0 Then
                            Call ErrFile.WriteLine("入帳期數大於分期剩餘期數,無法入帳!!")
                            
                            '***************************************************************************************
                            '#3417 錯誤時更新UCCode和UCName By Kin 2007/12/10
                            If blnFailUpd Then
                                
                                strFail = strFailUpdSQL & "||'入帳期數大於分期剩餘期數,無法入帳!!'" & _
                                                " Where CustId=" & rsTmp("CustId") & _
                                                " And CompCode=" & rsTmp("CompCode") & _
                                                " And ServiceType = '" & rsTmp("ServiceType") & "'" & _
                                                " And CitemCode = " & rsTmp("CitemCode") & "" & _
                                                " And FaciSeqNo = '" & rsTmp("FaciSeqNo") & "" & "'" & _
                                                " And " & strWhere
                                gcnGi.Execute strFail
                            End If
                            '****************************************************************************************
                            
                            lngErrCount = lngErrCount + 1
                            ChkData = True
                            Exit Function
                        End If
                        rsTmp.MoveNext
                    Loop
                    
            '#3527 多傳一個intRefNo變數 By Kin 2007/10/05
            '執行更新SO033
                If UpdAccount(vBillNo, vRealAmt, vClctEn, vClctName, vRealDate, vUpdEn, vCMCode, vCMName, vPTCode, vPTName, blnBillNO, strCitemQry, intRefNo, vBankId) = False Then
                    ChkData = False
                    Exit Function
                End If
          End If
       End If
    End If
    If rsChk.State = adStateOpen Then rsChk.Close
    Set rsChk = Nothing
    If rsCust.State = adStateOpen Then rsCust.Close
    Set rsCust = Nothing
    
    ChkData = True
  Exit Function
ChkErr:
  ErrSub "UpdateFunMid", "ChkData"
End Function
'#5386 判斷那幾筆客編不相同 By Kin 2009/12/04
Private Function ChkDiffCust(ByRef rs As ADODB.Recordset) As String
  On Error GoTo ChkErr
    Dim strRetMsg As String
    Dim strDiffCust As String
    Dim strRealDate As String
    Dim strRetCustId As String
    Dim strRetRealDate As String
    '7179
    If blnCrossCustCombine Then
        ChkDiffCust = ""
        Exit Function
    End If
    rs.MoveFirst
    Do While Not rs.EOF
        If strDiffCust <> rs("CustId") & "" Then
            strDiffCust = rs("CustId") & ""
            strRetCustId = IIf(strRetCustId = "", strDiffCust, strRetCustId & ";" & strDiffCust)
        End If
        rs.MoveNext
    Loop
    rs.MoveFirst
    If InStr(1, strRetCustId, ";") > 0 Then
        strRetMsg = strRetCustId
    End If
    ChkDiffCust = strRetMsg
    Exit Function
ChkErr:
  ErrSub "UpdateFunMid", "ChkDiffCust"
End Function

'#3527 增加參數intRefNo(CD018.REFNO)，如果為2，則使用POS3新的流程，目前只給mediaPost3使用 By Kin 2007/10/05
'#3417 增加參數 vBankId，如果有選擇收費人員則SO033的備註欄需填入銀行代碼
Public Function UpdAccount(ByVal vBillNo As String, ByVal vRealAmt As String, _
                           ByVal vClctEn As String, ByVal vClctName As String, _
                           ByVal vRealDate As String, ByVal vUpdEn As String, _
                           ByVal vCMCode As String, ByVal vCMName As String, _
                           ByVal vPTCode As String, ByVal vPTName As String, _
                           Optional blnBillNO As Boolean = False, Optional strCitemQry As String, _
                           Optional intRefNo As Integer = 0, _
                           Optional vBankId As String = Empty) As Boolean
    On Error GoTo ChkErr
    Dim rsAccount As New ADODB.Recordset
    Dim strSQL As String, strSeqQry As String, lngAffected As Long
    Dim lngPeriod As Long, lngAmount As Long
    Dim strSo003 As String, strUpdSql As String
    Dim strSQL2 As String, rsTmp As New ADODB.Recordset
        '收費項目
        If strCitemQry <> "" Then strSQL2 = " And CitemCode In (" & strCitemQry & ")"
        If intRefNo <> 2 Then
            If blnBillNO Then
                strSQL = "Select RowId, So033.* From " & strOwnerName & "So033 Where UCCODE IS NOT NULL  And BillNo ='" & vBillNo & "' AND CancelFlag <> 1" & strSQL2
            Else
                strSQL = "Select RowId, So033.* From " & strOwnerName & "So033 Where UCCODE IS NOT NULL  And MediaBillNo ='" & vBillNo & "' AND CancelFlag <> 1" & strSQL2
            End If
        Else
            strSQL = "Select RowId, So033.* From " & strOwnerName & "So033 Where UCCODE IS NOT NULL  And CitibankATM ='" & vBillNo & "' AND CancelFlag <> 1" & strSQL2
        End If
        If Not GetRS(rsAccount, strSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then
            UpdAccount = False
            Exit Function
        End If
        
        If rsAccount.RecordCount > 0 Then rsAccount.MoveFirst
        While Not rsAccount.EOF
            With rsAccount
                .Fields("RealAmt") = rsAccount("ShouldAmt")
                .Fields("RealDate") = NoZero(Format(vRealDate, "####/##/##"))
                .Fields("Updtime") = GetDTString(strNowTime)
                .Fields("UpdEn") = vUpdEn
                .Fields("UCCode") = Null
                .Fields("UCName") = Null
                .Fields("CMCode") = IIf(vCMCode = "", rsAccount("CMCode"), vCMCode)
                .Fields("CMName") = IIf(vCMName = "", rsAccount("CMName"), vCMName)
                '#7033 如果有輸入付款種類，則以輸入為主，否則不回填null By Kin 2015/04/28
                If Len(vPTCode) > 0 Then
                    .Fields("PTCode") = IIf(vPTCode = "", Null, vPTCode)
                    .Fields("PTName") = NoZero(vPTName)
                End If
'                .Fields("PTCode") = IIf(vPTCode = "", Null, vPTCode)
'                .Fields("PTName") = NoZero(vPTName)
                .Fields("ClctEn") = IIf(Trim(vClctEn) = "", rsAccount("ClctEn"), vClctEn)
                .Fields("ClctName") = IIf(Trim(vClctName) = "", rsAccount("ClctName"), vClctName)
                .Fields("FirstTime") = GetDTString(strNowTime)
                
                '*************************************************************************
                '#3417 如果有選收費人員則SO033.Note要填入銀行代碼 By Kin 2008/01/22
                If vClctEn <> Empty Then
                    If vBankId <> Empty Then
                        .Fields("Note") = .Fields("Note") & " " & "銀行代碼:" & vBankId
                    End If
                End If
                '*************************************************************************
                 On Error Resume Next
                Call AlterSO003(rsAccount)
                .Update
            End With
        
'            strSQL = "Update " & strOwnerName & "So033 Set RealAmt =" & rsAccount("ShouldAmt") & _
'                    ",ClctEn = '" & IIf(Trim(vClctEn) = "", rsAccount("ClctEn"), vClctEn) & "'" & _
'                    ",ClctName = '" & IIf(Trim(vClctName) = "", rsAccount("ClctName"), vClctName) & "'" & _
'                    ",RealDate = To_date('" & vRealDate & "','YYYYMMDD')" & _
'                    ",Updtime = '" & GetDTString(strNowTime) & "'" & _
'                    ",UpdEn = '" & vUpdEn & "'" & _
'                    ",UCCode = Null, UCName = Null" & _
'                    ", CMCode = '" & IIf(vCMCode = "", rsAccount("CMCode"), vCMCode) & "'" & _
'                    ", CMName = '" & IIf(vCMName = "", rsAccount("CMName"), vCMName) & "'" & _
'                    ", PTCode = " & GetNullString(vPTCode) & _
'                    ", PTName = " & GetNullString(vPTName) & _
'                    ",FirstTime = '" & GetDTString(strNowTime) & "'" & _
'                    " Where Rowid = '" & rsAccount("Rowid") & "'"
'            gcnGi.Execute strSQL
'
'            If Not GetRS(rsTmp, "Select * From " & strOwnerName & "SO033 Where RowId = '" & rsAccount("RowId") & "'") Then Exit Function
'            If AlterSO003(rsTmp) = False Then
'               UpdAccount = False
'               Exit Function
'            End If
'
lLoop:
            rsAccount.MoveNext
        Wend
            
            
'            If intPara8 >= 1 Then
'                lngPeriod = NoZero(rsAccount("RealPeriod"), True)
'                If intPara9 = 1 Then
'                    lngAmount = NoZero(rsAccount("OldAmt"), True)
'                ElseIf intPara9 = 2 Then
'                    lngAmount = NoZero(rsAccount("ShouldAmt"), True)
'                Else
'                    lngAmount = NoZero(rsAccount("ShouldAmt"), True)
'                End If
'                strUpdSql = ",Amount = " & lngAmount & ",Period = " & lngPeriod
'            End If
'
'
'            If rsAccount("SeqNo") Then strSeqQry = " And SeqNo = " & rsAccount("SeqNo")
'
'            If Val(rsAccount("BudgetPeriod") & "") > 0 Then
'                If Val(rsAccount("BudgetPeriod") & "") = 1 Then
'                    If Not ExecuteCommand("Delete From " & strOwnerName & "SO003 Where CustId = '" & rsAccount("CustId") & "' And CitemCode = '" & rsAccount("CitemCode") & "' And StopDate < " & _
'                        GetNullString(rsAccount("RealStopDate"), giDateV, giOracle) & strSeqQry) Then Exit Function
'                    GoTo lLoop
'                End If
'                strSo003 = ",BudgetPeriod = " & (rsAccount("BudgetPeriod") - 1)
'            End If
'            strSQL = "Update " & strOwnerName & "So003 Set StartDate = " & GetNullString(rsAccount("RealStartDate"), giDateV, giOracle) & _
'                    ",StopDate = " & GetNullString(rsAccount("RealStopDate"), giDateV, giOracle) & _
'                    ",ClctDate = " & GetNullString(rsAccount("RealStopDate"), giDateV, giOracle) & " + " & intPara5 & _
'                    strUpdSql & ",UpdEn = '" & vUpdEn & "',UpdTime = '" & GetDTString(strNowTime) & "'" & strSo003 & _
'                    " Where CustId = '" & rsAccount("CustId") & "' And CitemCode = '" & rsAccount("CitemCode") & "' And StopDate < " & _
'                    GetNullString(rsAccount("RealStopDate"), giDateV, giOracle)
'
'            gcnGi.Execute strSQL & strSeqQry, lngAffected
'            If strSeqQry <> "" And lngAffected = 0 Then
'                 gcnGi.Execute strSQL
'            End If
        lngCount = lngCount + 1
        Call CloseRecordset(rsAccount)
        UpdAccount = True
    Exit Function
ChkErr:
    UpdAccount = False
End Function

Public Function InitData(Optional strCompCode As String = "", Optional strServiceType As String = "") As Boolean
On Error GoTo ChkErr
   Dim rsTmp As New ADODB.Recordset
   Dim strSQL As String
   
   strSQL = "Select Para3, Para5, Para8, Para9 From " & strOwnerName & "So043"
   If Not GetRS(rsTmp, strSQL, gcnGi) Then InitData = False: Exit Function
   If rsTmp.EOF = False Then
      intPara3 = rsTmp("Para3")
      intPara5 = rsTmp("Para5")
      intPara8 = rsTmp("Para8")
      intPara9 = rsTmp("Para9")
   End If
   Set rsTmp = Nothing
   InitData = True
 Exit Function
ChkErr:
  ErrSub "UpdateFunMid", "InitData"
  InitData = False
End Function

Public Function MsgResult()
  On Error GoTo ChkErr
    lngTime = Timer - lngTime
    MsgBox "轉帳入帳筆數共：" & lngCount + lngErrCount & "筆" & vbCrLf & _
                   "成功共：" & lngCount & "筆" & vbCrLf & _
                   "失敗共：" & lngErrCount & "筆" & vbCrLf & _
                   "非正常戶共：" & lngStatusCount & "筆" & vbCrLf & _
                   "共花費：" & lngTime & "秒", vbInformation, "提示"
  Exit Function
ChkErr:
  ErrSub "UpdateFunMid", "MsgResult"
End Function

Public Sub SetLst(objLst As Object, _
                  strFldName1 As String, _
                  strFldName2 As String, _
                  Optional lngFldLen1 As Long, _
                  Optional lngFldLen2 As Long, _
                  Optional lngWidth1 As Long, _
                  Optional lngWidth2 As Long, _
                  Optional StrTableName As String, _
                  Optional strWhere As String, Optional blnStopFlag As Boolean = False)
  On Error GoTo ChkErr
    objLst.SetFldName1 strFldName1
    objLst.SetFldName2 strFldName2
    If lngFldLen1 > 0 Then objLst.FldLen1 = lngFldLen1
    If lngFldLen2 > 0 Then objLst.FldLen2 = lngFldLen2
    If lngWidth1 > 0 Then objLst.FldWidth1 = lngWidth1
    If lngWidth2 > 0 Then objLst.FldWidth2 = lngWidth2
    objLst.SetTableName strOwnerName & StrTableName
    objLst.SendConn gcnGi
    objLst.Filter = strWhere
  Exit Sub
ChkErr:
    Call ErrSub("UpdateFunMid", "SetLst")
End Sub

Public Function ChkGiList(Optional KeyCode As Integer) As Boolean
  On Error GoTo ChkErr
    With Screen
        If TypeOf .ActiveControl Is GiList Then
            If Len(Trim(.ActiveControl.GetCodeNo & "")) = 0 And Len(Trim(.ActiveControl.GetDescription & "")) = 0 Then ChkGiList = True: Exit Function
            ChkGiList = True
            If .ActiveControl.F5Corresponding Then
                If KeyCode <> vbKeyF5 Then
                   ChkGiList = True
                Else
                   ChkGiList = .ActiveControl.DataMatch
                   If KeyCode = vbKeyF5 Then Exit Function
                End If
            End If
            If .ActiveControl.F3Corresponding Then
                If KeyCode <> vbKeyF3 Then
                   ChkGiList = True
                Else
                   ChkGiList = .ActiveControl.DataMatch
                   If KeyCode = vbKeyF3 Then Exit Function
                End If
            End If
            If .ActiveControl.F2Corresponding Then
                If KeyCode <> vbKeyF2 Then
                   ChkGiList = True
                Else
                   ChkGiList = .ActiveControl.DataMatch
                   If KeyCode = vbKeyF2 Then Exit Function
                End If
            End If
        Else
            ChkGiList = True
        End If
    End With
  Exit Function
ChkErr:
    ChkGiList = True
    Resume Next
'   ErrSub "ExtraModule", "ChkGiList"
End Function

Public Function GetBillNo_New(ByVal strNo As String) As String
    On Error Resume Next
    Dim strType As String
    Dim strYM As String
    Dim strMM As String
    
'        strNo = Mid(strReadLine, 65, 11)
        
        strYM = Left(strNo, 2)
        strMM = Mid(strNo, 3, 1)
        If strMM > "9" Then
            strYM = strYM & Format(Trim(Asc(strMM) - Asc("A") + 10), "00")
        Else
            strYM = strYM & Format(strMM, "00")
        End If
        Select Case Mid(strNo, 4, 1)
            Case 1
                strType = "BC"
            Case 2
                strType = "TC"
            Case 3
                strType = "BI"
            Case 4
                strType = "TI"
        End Select
        
        GetBillNo_New = Trim(CStr(Left(strYM, 4) + 191100)) & _
                strType & Format(Right(strNo, 7), "0000000")
End Function

Public Function GetBillNo_Old(ByVal strNo As String, ByVal strServiceType) As String
  On Error Resume Next
    
    GetBillNo_Old = Trim(CStr(Val(Left(strNo, 4)) + 191100)) & _
            Mid(strNo, 5, 1) & strServiceType & Format(Right(strNo, 6), "0000000")
End Function
'#3122 檢核是否有超過日結日期 By Kin Add 2007/03/28
Public Function ChkOverCloseDate(ByVal strCloseDate As String, _
    ByVal strCompCode As String, ByVal strServiceType As String, strResult) As Boolean
    On Error GoTo ChkErr
    Dim intDayCut As Integer
    Dim strTranDate As String
    Dim intPara6 As Integer
        ChkOverCloseDate = False
        intDayCut = Val(GetSystemParaItem("DayCut", strCompCode, "", "SO041") & "")
        strTranDate = GetRsValue("Select tranDate from " & GetOwner & "So062 Where CompCode= " & strCompCode & " And Type=1 Order By TranDate Desc") & ""
        '收費參數檔（SO043)，取收費日登錄期限<Para6>，供後續使用！
        intPara6 = Val(GetSystemParaItem("Para6", strCompCode, strServiceType, "SO043") & "")
    
        If strCloseDate = "" Then strCloseDate = RightDate
        
        If Not IsDate(strCloseDate) Then
            strResult = "日期不合法！"
            Exit Function
        End If
        
        If CDate(strCloseDate) > RightDate Then strResult = "此日期超過今天日期": Exit Function

        If (RightDate - CDate(strCloseDate)) > intPara6 Then
            strResult = "此日期已超過系統設定的安全期限！"
            Exit Function
        End If
        If CDate(strCloseDate) <= IIf(strTranDate <> "", CDate(strTranDate), 0) Then
            If intDayCut = 1 And CDate(strCloseDate) = IIf(strTranDate <> "", CDate(strTranDate), 0) Then
                ChkOverCloseDate = True
                Exit Function
            End If
            strResult = " 已過曰結日期不可入帳！"
            Exit Function
        End If
        ChkOverCloseDate = True
    Exit Function
ChkErr:
    ErrSub "UpdateFunMid", "ChkOverCloseDate"
End Function



