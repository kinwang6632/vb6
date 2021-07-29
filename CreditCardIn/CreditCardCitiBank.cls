VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CreditCardCitiBank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const FormName = "CreditCardCitiBank"
Dim strFilePath As String
Dim strRealDate As String
Dim strSourcePath As String
Dim strUpdEn As String
Dim strUPDName As String
Dim strUPDTime As String
Dim strPrgName As String
Dim strPTCode As String
Dim strPTName As String
Dim strServiceType As String
Dim strCompCode As String
Dim lngTotalCount As Long, lngTotalAmt As Long
Dim lngErrCount As Long, lngErrAmt As Long
Dim lngPara24 As Long
Dim strCitemQry As String
Dim strClctEn As String, strClctName As String
Dim strCMCode As String, strCMName As String
Dim strErrFileName As String
Dim lngExportFormat As Long
Private strBankId As String
Public Function Action() As Boolean
    On Error GoTo ChkErr
    Dim rsCD013 As New ADODB.Recordset
    blnFailUpd = False
    strUCCode = Empty
    strUCName = Empty
    '********************************************************************************************************************************************************
    '#3417 電子檔匯出時,要填入未收原因(RefNo=4) By Kin 2007/12/13
    If Not GetRS(rsCD013, "Select * From " & strOwnerName & "CD013 Where RefNo=6 And StopFlag<>1 Order By CodeNo Desc", gcnGi, adUseClient, adOpenKeyset) Then Exit Function
    If rsCD013.EOF Then
        blnFailUpd = False
    Else
        blnFailUpd = True
        strUCCode = rsCD013("CodeNo") & ""
        strUCName = rsCD013("Description") & ""
    End If
    '*********************************************************************************************************************************************************
    If Not InitData Then GoTo KillFile
    If Not OpenData Then GoTo KillFile
    If Not BeginTrans Then GoTo KillFile
    CloseRecordset rsCD013
    Action = True
KillFile:
    CloseFS
    Exit Function
ChkErr:
    CloseRecordset rsCD013
    ErrSub FormName, "Action"
End Function

Private Sub CloseFS()
    On Error Resume Next
        File.Close
        Set File = Nothing
        ErrFile.Close
        Set ErrFile = Nothing
        Set fso = Nothing
    
End Sub

Private Function InitData() As Boolean
    On Error GoTo ChkErr
        With objStorePara
            strFilePath = .FilePath
            strRealDate = .RealDate
            strSourcePath = .SourcePath
            strUpdEn = .UpdEn
            strUPDName = .UpdName
            strUPDTime = GetDTString(RightNow)
            strPrgName = .PrgName
            strPTCode = .PTCode
            strPTName = .PTName
            strServiceType = .ServiceType
            strCompCode = .CompCode
            strCitemQry = .uCitemQry
            strClctEn = .uClctEn
            strClctName = .uClctName
            strCMCode = .uCMCode
            strCMName = .uCMName
            strPTCode = .PTCode
            strPTName = .PTName
            strBankId = .BankId
            lngExportFormat = .ExportFormat  '判斷是不是使用批次格式
        End With
        lngPara24 = Val(GetRsValue("Select Para24 From " & GetOwner & "So043 ") & "")
        InitData = True
    Exit Function
ChkErr:
    ErrSub FormName, "InitData"
End Function

Private Function OpenData() As Boolean
    On Error GoTo ChkErr
    Dim strFileName As String
        
        strErrFileName = strFilePath & "\" & strPrgName & "Err.log"
        strFileName = strSourcePath
        
        Set File = fso.OpenTextFile(strFileName, ForReading)
        Set ErrFile = fso.CreateTextFile(strErrFileName, True)
        
        OpenData = True
    Exit Function
ChkErr:
    ErrSub FormName, "OpenData"
End Function

Private Function BeginTrans() As Boolean
    On Error GoTo ChkErr
    Dim strReadLine As String
    Dim lngTime As Long
        lngCount = 0
        lngErrCount = 0
        lngTime = Timer
        'strUpdTime = RightNow
        gcnGi.BeginTrans
        Do While Not File.AtEndOfLine
        
            '****************************************************************************************
            '#3417 失敗更新UCCode與UCName的SQL基本語法 By Kin 2007/12/13
            '#4388 增加更新異動人員與異動時間 By Kin 2009/05/04
            strFailUpdSQL = Empty
            strFailUpdSQL = "UPDATE " & strOwnerName & "SO033 Set UCCode=" & strUCCode & _
                            ",UCName='" & strUCName & "',UPDEN='" & strUPDName & "'" & _
                            ",UPDTime='" & strUPDTime & "',Note=Note||"
            '****************************************************************************************
            
            strReadLine = File.ReadLine
            If ChkDataOk(strReadLine) Then
                If Not UpdateToSO033(strReadLine) Then gcnGi.RollbackTrans: Exit Function
            End If
        Loop
        gcnGi.CommitTrans
        If lngErrCount > 0 Then
            Shell "NotePad " & strErrFileName, vbNormalFocus
        Else
            objStorePara.UpdState = True
        End If
        MsgBox "成功筆數: " & lngTotalCount & " ,成功金額: " & lngTotalAmt & vbCrLf & vbCrLf & _
                    "失敗筆數: " & lngErrCount & " ,失敗金額: " & lngErrAmt & vbCrLf & vbCrLf & _
                    "共花費 : " & Round(Timer - lngTime) & " 秒", vbInformation, "提示"
        BeginTrans = True
    Exit Function
ChkErr:
    ErrSub FormName, "BeginTrans"
End Function

Private Function ChkDataOk(strReadLine As String, Optional ByVal strBankName As String) As Boolean
    On Error GoTo ChkErr
    Dim rsTmp As New ADODB.Recordset
    Dim strBillNo As String, strErrDescription As String
    Dim lngBudgetPeriod As Long
    Dim strChkRealDate As String
    Dim strResult As String
    Dim strErrCode As String
    Dim strCD013RefNo As String
    blnCrossCustCombine = Val(GetRsValue("Select Nvl(CrossCustCombine,0) From " & strOwnerName & "SO041", gcnGi) & "")
    strChkRealDate = GetRealDate(strReadLine)
    strCD013RefNo = " MAX(DECODE(NVL(CD013.PAYOK,0),1,3," & _
                "DECODE(NVL(CD013.REFNO,0),3,3,DECODE(NVL(CD013.REFNO,0),7,3,0)))) REFNO "
        If lngExportFormat = 0 Then  'HSB格式
            Select Case Left(strReadLine, 1)
                Case "D"
                    Select Case MidMbcs(strReadLine, 82, 2)
                        Case "00" '                00 = 轉帳成功
                            'lngTotalAmt = lngTotalAmt + GetRealAmt(strReadLine)
                            'lngTotalCount = lngTotalCount + 1
                            strErrDescription = ""
                        Case "01" '                01 = Refer to card issuer
                            strErrDescription = "Refer to card issuer !! "
                            strErrCode = "01"
                        Case "02" '                02 = Refer to card issuer's for special condition
                            strErrDescription = "Refer to card issuers for special condition !! "
                            strErrCode = "02"
                        Case "03" '                03 = Invalid merchant
                            strErrDescription = "Invalid merchant !! "
                            strErrCode = "03"
                        Case "04" '                04 = Pick-up Card
                            strErrDescription = "Pick-up Card !! "
                            strErrCode = "04"
                        Case "05" '                05 = Do not honor
                            strErrDescription = "Do not honor !! "
                            strErrCode = "05"
                        Case "07" '                07 = HOLD CALL, Pick-up card, Special condition
                            strErrDescription = "HOLD CALL, Pick-up card, Special condition !! "
                            strErrCode = "07"
                        Case "11" '                11 = Approved (VIP)
                            strErrDescription = "Approved (VIP) !! "
                            strErrCode = "11"
                        Case "12" '                12 = Invalid Transaction Type
                            strErrDescription = "Invalid Transaction Type !! "
                            strErrCode = "12"
                        Case "13" '                13 = Invalid Amount
                            strErrDescription = "Invalid Amount !! "
                            strErrCode = "13"
                        Case "14" '                14 = Invalid card number (no such number)
                            strErrDescription = "Invalid card number (no such number) !! "
                            strErrCode = "14"
                        Case "19" '                19 = Re-Enter Transaction
                            strErrDescription = "Re-Enter Transaction !! "
                            strErrCode = "19"
                            
                        Case "21" '                21 = NO ACTION TAKEN
                            strErrDescription = "NO ACTION TAKEN !! "
                            strErrCode = "21"
                            
                        Case "30" '                30 = Format error
                            strErrDescription = "Format error !! "
                            strErrCode = "30"
                        Case "31" '                31 = Bank not supported by switch
                            strErrDescription = "Bank not supported by switch !! "
                            strErrCode = "31"
                        Case "41" '                41 = Pick-up Card for Lost card
                            strErrDescription = "Pick-up Card for Lost card !! "
                            strErrCode = "41"
                        Case "43" '                43 = Pick-up Card for Stolen card, Pickup
                            strErrDescription = "Pick-up Card for Stolen card, Pickup !! "
                            strErrCode = "43"
                            
                        Case "51" '                51 = Not sufficient funds
                            strErrDescription = "Not sufficient funds !! "
                            strErrCode = "51"
                        Case "54" '                54 = Expired card or card expire date error
                            strErrDescription = "Expired card or card expire date error !! "
                            strErrCode = "54"
                        Case "58" '                58 = Card has not been activated or others
                            strErrDescription = "Card has not been activated or others !! "
                            strErrCode = "58"
                            
                        Case "61" '                61 = Exceeds withdrawal amount limit
                            strErrDescription = "Exceeds withdrawal amount limit !! "
                            strErrCode = "61"
                        Case "63" '                63 = CVC error
                            strErrDescription = "CVC error !! "
                            strErrCode = "63"
                            
                        Case "75" '                75 = Allowable number of PIN tries exceeded
                            strErrDescription = "Allowable number of PIN tries exceeded !! "
                            strErrCode = "75"
                        Case "77" '                77 = Reversal data are inconsistent with Original Message
                            strErrDescription = "Reversal data are inconsistent with Original Message !! "
                            strErrCode = "77"
                        Case "76" '                76 = REFERENCE ERROR or 產品代碼錯誤(product error)
                            strErrDescription = "REFERENCE ERROR or 產品代碼錯誤(product error)  !! "
                            strErrCode = "77"
                        Case "79" '                79 = ALREADY REVERSED
                            strErrDescription = "ALREADY REVERSED !! "
                            strErrCode = "79"
                            
                        Case "85" '                85 = Card ok, declined by issue bank
                        
                            strErrDescription = "Card ok, declined by issue bank !! "
                            strErrCode = "85"
                        Case "89" '                89 = System error (reserved for BASE24 use)
                            strErrDescription = "System error (reserved for BASE24 use) !! "
                            strErrCode = "89"
                            
                        Case "90" '                90 = Cut-off is in process (switch ending a day's business and starting Then ext. Transaction can be sent again in a few minutes)
                            strErrDescription = "Cut-off is in process (switch ending a days business and starting Then ext. Transaction can be sent again in a few minutes) !! "
                            strErrCode = "90"
                        Case "91"
                            strErrDescription = "Issuer or switch is inoperative (發卡行系統錯誤) !! "
                            strErrCode = "91"
                        Case "92"
                            strErrDescription = "BIN Error (錯誤分期卡別) !! "
                            strErrCode = "92"
                        Case "94"
                            strErrDescription = "Duplicate transaction !! "
                            strErrCode = "94"
                        Case "96"
                            strErrDescription = "System Malfunction (發卡行系統錯誤)料 !! "
                            strErrCode = "96"
                        Case Else
                            strErrDescription = "其他不成功原因 !! "
                            strErrCode = MidMbcs(strReadLine, 82, 2)
                    End Select
                    '銀行之錯誤
                    If strErrDescription <> "" Then
                        Call InsertToErr(strReadLine, strErrDescription, strBankName, strErrCode)
                        Exit Function
                    End If
                    '程式檢查的錯誤
                    strBillNo = GetBillNo(strReadLine)
                    '#4010 如果UCCode的參考號為3要視為異常資料 By Kin 2008/07/29
                    '#5564 CD013.PAYOK=1或CD013.REFNO=3,7都算櫃台已收 By Kin 2010/05/19
                    If Not GetRS(rsTmp, "Select Nvl(Sum(SO033.ShouldAmt),0) as TotalAmt," & strCD013RefNo & " From " & _
                                    strOwnerName & "So033," & strOwnerName & "CD013 " & GetWhereCondition(strBillNo, False)) Then
                        Exit Function
                    End If
                    If rsTmp.RecordCount = 0 Then
                        Call InsertToErr(strReadLine, "無此單據編號或此單號已收款 !! ", strBankName)
                        Exit Function
                    End If
                    If rsTmp("TotalAmt") = 0 Then
                        Call InsertToErr(strReadLine, "無此單據編號或此單號已收款 !! ", strBankName)
                        Exit Function
                    End If
                    If Val(rsTmp("TotalAmt") & "") <> GetRealAmt(strReadLine) Or Val(rsTmp("TotalAmt") & "") = 0 Then
                        Call InsertToErr(strReadLine, "銀行轉帳金額與應收金額不符 !! ", strBankName)
                        Exit Function
                    End If
                    '#3122 檢核是否有超過日結日期 By Kin Add 2007/03/29
                    If Not ChkOverCloseDate(Format(strChkRealDate, "####/##/##"), strCompCode, strServiceType, strResult) Then
                        Call InsertToErr(strReadLine, strResult, strBankName)
                        Exit Function
                    End If
                    If Val(rsTmp("RefNo") & "") = 3 Then
                        Call InsertToErr(strReadLine, "櫃臺已收!", strBankName)
                        Exit Function
                    End If
                    If Not GetRS(rsTmp, "Select RealPeriod,CustId,CompCode,ServiceType,CitemCode,FaciSeqNo From " & strOwnerName & "So033 " & _
                                    GetWhereCondition(strBillNo, False, False)) Then Exit Function
                    '94/09/30 1805 Jacky
                    Do While Not rsTmp.EOF
                        lngBudgetPeriod = Val(GetRsValue("Select Nvl(BudgetPeriod,0) From " & strOwnerName & "SO003 Where Custid = " & rsTmp("CustId") & " And CompCode = " & rsTmp("CompCode") & " And ServiceType = '" & rsTmp("ServiceType") & "' " & _
                            " And CitemCode = " & rsTmp("CitemCode") & " And FaciSeqNo = '" & rsTmp("FaciSeqNo") & "'") & "")
                        If rsTmp("RealPeriod") > lngBudgetPeriod And lngBudgetPeriod > 0 Then
                            Call InsertToErr(strReadLine, "入帳期數大於分期剩餘期數,無法入帳!!", strBankName)
                            Exit Function
                        End If
                        rsTmp.MoveNext
                    Loop
                    If GetAuthorizeNo(strReadLine) = "" Then
                        Call InsertToErr(strReadLine, "無核准號碼 !! ", strBankName)
                        Exit Function
                    End If
                Case "H"
                    Exit Function
                Case Else
                    Call InsertToErr(strReadLine, "格式有問題,無法入帳!!", strBankName)
                    Exit Function
            End Select
        Else
            '花旗批次格式
            Select Case UCase(Trim(MidMbcs(strReadLine, 2, 1)))
                Case "H"
                    Exit Function
                Case "D"
                    '#3305 將原本成功狀態其中之一的"SETTLE"改為SALE By Kin 2007/07/26
                    If Trim(UCase(MidMbcs(strReadLine, 42, 12))) = "AUTH" Or Trim(UCase(MidMbcs(strReadLine, 42, 12))) = "CAP" Or Trim(UCase(MidMbcs(strReadLine, 42, 12))) = "SALE" Then
                        strErrDescription = ""
                    Else
                        strErrDescription = "請款或授權失敗!!" & _
                                            "  執行結果碼： " & Right("00" & Trim(MidMbcs(strReadLine, 147, 2)), 2) & _
                                            "  錯誤回覆碼： " & Right("00" & Trim(MidMbcs(strReadLine, 149, 2)), 2)
                        Call InsertToErr(strReadLine, strErrDescription, strBankName)
                        Exit Function
                    End If
                    '檢查在應收費項目裡是否有此一帳單
                    strBillNo = GetBillNo(strReadLine)
                    
                    If Not GetRS(rsTmp, "Select * From " & strOwnerName & "So033 " & GetWhereCondition(strBillNo, False, False) & " And AccountNo='" & GetAccountNo(strReadLine) & "'") Then
                        Exit Function
                    End If
                    If rsTmp.RecordCount = 0 Then
                    '將出Log的訊息明確表達(原來都是秀無此單據編號或此單號已收款!!) For Jacy 2006/10/12
'                        Call InsertToErr(strReadLine, "無此單據編號或此單號已收款 !! ", strBankName)
                        Call InsertToErr(strReadLine, "無此單據編號 !! ", strBankName)
                        Exit Function
                    End If
                    '判斷是否已收款
                    '#4010 如果UCcode的參考號為3要視為異常資料 By Kin 2008/07/30
                    '#5564 CD013.PAYOK=1或CD013.REFNO=3,7都算櫃台已收 By Kin 2010/05/19
                    If Not GetRS(rsTmp, "Select Nvl(Sum(SO033.ShouldAmt),0) as TotalAmt," & strCD013RefNo & " From " & _
                                        strOwnerName & "So033," & strOwnerName & "CD013 " & GetWhereCondition(strBillNo, False, True) & " And AccountNo='" & GetAccountNo(strReadLine) & "'") Then
                        Exit Function
                    End If
                    If rsTmp("TotalAmt") = 0 Then
'                        Call InsertToErr(strReadLine, "無此單據編號或此單號已收款 !! ", strBankName)
                        Call InsertToErr(strReadLine, "此單號已收款 !! ", strBankName)
                        Exit Function
                    End If
                    '判斷金額是否相符
                    If Val(rsTmp("TotalAmt") & "") <> GetRealAmt(strReadLine) Or Val(rsTmp("TotalAmt") & "") = 0 Then
                        Call InsertToErr(strReadLine, "銀行轉帳金額與應收金額不符 !! ", strBankName)
                        Exit Function
                    End If
                    '#3122 檢核是否有超過日結日期 By Kin Add 2007/03/29
                    If Not ChkOverCloseDate(Format(strChkRealDate, "####/##/##"), strCompCode, strServiceType, strResult) Then
                        Call InsertToErr(strReadLine, strResult, strBankName)
                        Exit Function
                    End If
                    '#4010 UCCode的參考號為3要視為異常資料 By Kin 2008/07/30
                    If Val(rsTmp("RefNo") & "") = 3 Then
                        Call InsertToErr(strReadLine, "櫃臺已收!!", strBankName)
                        Exit Function
                    End If
                    If Not GetRS(rsTmp, "Select RealPeriod,CustId,CompCode,ServiceType,CitemCode,FaciSeqNo From " & strOwnerName & "So033 " & GetWhereCondition(strBillNo, False, False)) Then Exit Function
                    Do While Not rsTmp.EOF
                        lngBudgetPeriod = Val(GetRsValue("Select Nvl(BudgetPeriod,0) From " & strOwnerName & "SO003 Where Custid = " & rsTmp("CustId") & " And CompCode = " & rsTmp("CompCode") & " And ServiceType = '" & rsTmp("ServiceType") & "' " & _
                            " And CitemCode = " & rsTmp("CitemCode") & " And FaciSeqNo = '" & rsTmp("FaciSeqNo") & "'") & "")
                        If rsTmp("RealPeriod") > lngBudgetPeriod And lngBudgetPeriod > 0 Then
                            Call InsertToErr(strReadLine, "入帳期數大於分期剩餘期數,無法入帳!!", strBankName)
                            Exit Function
                        End If
                        rsTmp.MoveNext
                    Loop
                    If GetAuthorizeNo(strReadLine) = "" Then
                        Call InsertToErr(strReadLine, "無核准號碼 !! ", strBankName)
                        Exit Function
                    End If
                Case Else
                    Call InsertToErr(strReadLine, "格式有問題,無法入帳!!", strBankName)
                    Exit Function
            End Select
        End If
            
        Call CloseRecordset(rsTmp)
        ChkDataOk = True
    Exit Function
ChkErr:
    ErrSub FormName, "ChkDataOk"
End Function

Private Function GetAuthorizeNo(strReadLine As String) As String
    On Error Resume Next
    Select Case lngExportFormat
        Case 0
            GetAuthorizeNo = Trim(MidMbcs(strReadLine, 76, 6))
        Case 1
            GetAuthorizeNo = Trim(MidMbcs(strReadLine, 141, 6))
    End Select
End Function

Private Function GetBillNo(strReadLine As String) As String
    On Error Resume Next
    Select Case lngExportFormat
        Case 0
            GetBillNo = Trim(MidMbcs(strReadLine, 104, 20))
        Case 1
            GetBillNo = Trim(MidMbcs(strReadLine, 19, 19))
    End Select
End Function

Private Function GetRealAmt(strReadLine As String) As Long
    On Error Resume Next
    Select Case lngExportFormat
        Case 0
            GetRealAmt = Val(MidMbcs(strReadLine, 64, 10))
        Case 1
            GetRealAmt = Val(MidMbcs(strReadLine, 54, 11) & ".00")
    End Select
End Function

Private Function GetRealDate(strReadLine As String) As String
    On Error Resume Next
        GetRealDate = strRealDate
        If GetRealDate = "" Then GetRealDate = Format(RightDate, "YYYYMMDD")
End Function

Private Function GetAccountNo(strReadLine As String) As String
    On Error Resume Next
    Select Case lngExportFormat
        Case 0
            GetAccountNo = Trim(MidMbcs(strReadLine, 38, 19))
        Case 1
            GetAccountNo = Trim(MidMbcs(strReadLine, 67, 16))
    End Select
End Function

Private Function GetWhereCondition(strBillNo As String, blnUpd As Boolean, Optional ByVal blnCD013 As Boolean = True) As String
    On Error Resume Next
    
    Dim strSQL As String
        If lngPara24 = 1 Then
            strSQL = "SO033.MediaBillNo = '" & strBillNo & "'"
        Else
            strSQL = "SO033.BillNo = '" & strBillNo & "'"
        End If
        If strCitemQry <> "" Then strSQL = strSQL & IIf(strSQL = "", "", " And ") & " SO033.CitemCode In (" & strCitemQry & ")"
        '#4010 如果UCcode的參考號為3要視為異常資料 By Kin 2008/07/30
        If blnCD013 Then
            strSQL = IIf(strSQL = "", " SO033.UCCode=CD013.CodeNo ", strSQL & " And SO033.UCCode=CD013.CodeNo ")
        End If
        If Not blnUpd Then
            GetWhereCondition = " Where " & strSQL & IIf(strSQL = "", "", " And ") & "SO033.RealDate Is Null And SO033.CancelFlag=0 And SO033.UCCode is Not Null "
        Else
            GetWhereCondition = " Where " & strSQL
        End If
End Function

Private Function UpdateToSO033(strReadLine As String) As Boolean
    On Error GoTo ChkErr
    Dim strBillNo As String
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim lngAffected As Long
        strBillNo = GetBillNo(strReadLine)
        
        strSQL = "Update " & strOwnerName & "So033 Set RealAmt = ShouldAmt" & _
                ",UpdTime = '" & strUPDTime & "'" & _
                ",FirstTime = '" & strUPDTime & "'" & _
                ",UpdEn = '" & strUPDName & "'" & _
                ",UCCode = Null, UCName = Null"
        If strRealDate = "" Then
            strRealDate = GetRealDate(strReadLine)
        End If
        strSQL = strSQL & ",RealDate = To_date('" & strRealDate & "','YYYYMMDD')"
        '收費人員
        If strClctEn <> "" Then
            strSQL = strSQL & ",ClctEn = '" & strClctEn & "'" & ",ClctName = '" & strClctName & "'"
            
            '******************************************************************************************
            '#3417 測試不OK,當有輸入收費人員,SO033.Note要填入銀行代號,要在這裡Update By Kin 2008/01/28
            strSQL = strSQL & ",Note=Note ||'銀行代碼:" & strBankId & "'"
            '******************************************************************************************
        Else
            strSQL = strSQL & ",ClctEn = OldClctEn ,ClctName = OldClctName "
        End If
        '收費方式
        If strCMCode <> "" Then strSQL = strSQL & ",CMCode = " & strCMCode & ", CMName = '" & strCMName & "'"
        '付款種類
        If strPTCode <> "" Then strSQL = strSQL & ",PTCode = " & strPTCode & ", PTName = '" & strPTName & "'"
        '核准號碼
        strSQL = strSQL & ",AuthorizeNo = '" & GetAuthorizeNo(strReadLine) & "'"
        
        If Not GetRS(rsTmp, "Select RowId,SO033.* From " & strOwnerName & "So033 " & GetWhereCondition(strBillNo, True, False), gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
        Set rsTmp.ActiveConnection = Nothing
        
        Do While Not rsTmp.EOF
            rsTmp("RealAmt") = rsTmp("ShouldAmt")
            rsTmp("UpdTime") = strUPDTime
            rsTmp("FirstTime") = strUPDTime
            rsTmp("UpdEn") = strUPDName
            rsTmp("UCCode") = Null
            rsTmp("UCName") = Null
            rsTmp("RealDate") = Format(strRealDate, "####/##/##")
            If strClctEn <> "" Then
                rsTmp("ClctEn") = strClctEn
                rsTmp("ClctName") = strClctName
                
                '*********************************************************************************************************************
                '#3417 如果有輸入收費人員,SO033.NOTE要填入銀行代號 By Kin 2008/01/22
                If strBankId <> Empty Then
                    rsTmp("NOTE") = IIf(IsNull(rsTmp("NOTE")), "銀行代碼:" & strBankId, rsTmp("NOTE") & " " & "銀行代碼:" & strBankId)
                End If
                '*********************************************************************************************************************
            Else
                rsTmp("ClctEn") = rsTmp("OldClctEn")
                rsTmp("ClctName") = rsTmp("OldClctName")
            End If
            If strCMCode <> "" Then
                rsTmp("CMCode") = strCMCode
                rsTmp("CMName") = strCMName
            End If
            If strPTCode <> "" Then
                rsTmp("PTCode") = strPTCode
                rsTmp("PTName") = strPTName
            End If
            
            If Not AlterSO003(rsTmp) Then Exit Function
            rsTmp.CancelUpdate
            rsTmp.MoveNext
        Loop
        
        If Not ExecuteCommand(strSQL & GetWhereCondition(strBillNo, True, False), gcnGi, lngAffected) Then Exit Function
        
        lngTotalAmt = lngTotalAmt + GetRealAmt(strReadLine)
        lngTotalCount = lngTotalCount + 1
        UpdateToSO033 = True
    Exit Function
ChkErr:
    ErrSub FormName, "UpdateToSO033"
End Function

Private Function InsertToErr(strReadLine As String, strErrDescription As String, strBankName As String, Optional strErrCode As String) As Boolean
    On Error Resume Next
    Dim strBillNo As String
    Dim strAccountNo As String
        strBillNo = GetBillNo(strReadLine)
        strAccountNo = GetAccountNo(strReadLine)
        
        WriteTextLine ErrFile, "收費單號:" & strBillNo & "; 信用卡卡號:" & strAccountNo & "; 錯誤名稱:" & strErrDescription
        
        '****************************************************************************************************************************
        '#3417 如果失敗則更新UCCode與UCName By Kin 2007/12/13
        If blnFailUpd Then
            
            strFailUpdSQL = strFailUpdSQL & IIf(strErrCode <> "", " '錯誤代碼:" & strErrCode & " ", "'") & _
                            "收費單號:" & strBillNo & "; 信用卡卡號:" & strAccountNo & "; 錯誤名稱:" & strErrDescription & "' " & _
                            GetWhereCondition(strBillNo, False, False)
            gcnGi.Execute strFailUpdSQL
        End If
        '*****************************************************************************************************************************
        
        lngErrCount = lngErrCount + 1
        lngErrAmt = lngErrAmt + GetRealAmt(strReadLine)
        InsertToErr = True
End Function

Public Function WriteTextLine(objFile As Object, _
        ByVal strData As String) As Boolean
    On Error GoTo ChkErr
        objFile.WriteLine strData
        WriteTextLine = True
    Exit Function
ChkErr:
    Call ErrSub("DealFun", "WriteTextLine")
End Function

