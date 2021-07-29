Attribute VB_Name = "UpdateFun"
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
Public blnSO3314 As Boolean
'Public GetOwner As String
Public blnCrossCustCombine  As Boolean
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
    Call ErrSub("DealFun", "Function ChkDirExist")
End Function

Public Function Decrypt(EncryptionString As String) As String
  On Error GoTo ChkErr
    Decrypt = CreateObject("DevPowerEncrypt.EnCrypt").Decrypt(EncryptionString)
  Exit Function
ChkErr:
    Call ErrSub("DealFun", "Function Decrypt")
    
End Function

Public Function ChkData(ByVal vBillNo As String, ByVal vRealAmt As String, _
                        ByVal vClctEn As String, ByVal vClctName As String, _
                        ByVal vRealDate As String, ByVal vUpdEn As String, _
                        ByVal vCMCode As String, ByVal vCMName As String, _
                        ByVal vPTCode As String, ByVal vPTName As String, _
                        ByVal vCM As String, _
                        Optional blnBillNo As Boolean = False, _
                        Optional blnMediaBillNo As Boolean = False, _
                        Optional strCitemQry As String, _
                        Optional intpara24 As Integer, _
                        Optional intAccLong As Integer, _
                        Optional strOld_billno As String, Optional strErrorMsg As String) As Boolean
  On Error GoTo ChkErr
    Dim strSQL As String
    Dim rsChk As New ADODB.Recordset
    Dim rsCust As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim strErr As String
    Dim strSQL2 As String
    Dim strBillNo_Old As String
    Dim strRealDate As String
    Dim strReturn As String
    Dim strCD013Ref As String
    Dim strCustId As String
    Dim lngTotalShouldAmt As Long
    Dim lngTotalRealAmt As Long
    strCD013Ref = " Decode(Nvl(CD013.RefNo,0),7,3,Decode(Nvl(CD013.RefNo,0),8,3," & _
            "Decode(Nvl(CD013.RefNo,0),3,3,0))) RefNo "
    
    '收費項目
        If strCitemQry <> "" Then strSQL2 = " And SO033.CitemCode In (" & strCitemQry & ")"
        If intpara24 = 0 Then
            '#4010 UCCode的參考號為3要視為異常 By Kin 2008/07/30
            If blnBillNo Then
                strSQL = "Select Nvl(sum(SO033.ShouldAmt),0) as TotalShouldAmt, Nvl(sum(SO033.RealAmt),0) as TotalRealAmt," & _
                         "SO033.RealDate,SO033.CustId,SO033.ServiceType,Max(Decode(CD013.RefNo,3,3,0)) RefNo From " & strOwnerName & "SO033," & strOwnerName & "CD013 " & _
                         " Where SO033.UCCODE IS NOT NULL AND SO033.BillNo = '" & vBillNo & "' And SO033.CancelFlag <> 1 And SO033.UCCode=CD013.CodeNo " & strSQL2 & " Group By SO033.CustId,SO033.RealDate,SO033.ServiceType"
            Else
            '問題集3008 將Like %拿掉,以增加效率,By Kin 2007/1/22
                'strSQL = "Select Nvl(sum(ShouldAmt),0) as TotalShouldAmt, Nvl(sum(RealAmt),0) as TotalRealAmt, RealDate, CustId, ServiceType From " & strOwnerName & "SO033 Where UCCODE IS NOT NULL AND CitiBankATM like '" & Mid(vBillNo, 1, 15) & "%' And CancelFlag <> 1 " & strSQL2 & " Group By CustId,RealDate,ServiceType"
                strSQL = "Select Nvl(sum(SO033.ShouldAmt),0) as TotalShouldAmt, Nvl(sum(SO033.RealAmt),0) as TotalRealAmt," & _
                         "SO033.RealDate,SO033.CustId,SO033.ServiceType,Max(Decode(CD013.RefNo,3,3,0)) RefNo" & _
                         " From " & strOwnerName & "SO033," & strOwnerName & "CD013 " & _
                         " Where SO033.UCCODE IS NOT NULL AND SO033.CitiBankATM = '" & vBillNo & "' And SO033.CancelFlag <> 1 And SO033.UCCode=CD013.CodeNo " & strSQL2 & _
                         " Group By SO033.CustId,SO033.RealDate,SO033.ServiceType"
            End If
        Else
            If blnBillNo Then
                If blnMediaBillNo Then
                    strSQL = "Select Nvl(sum(SO033.ShouldAmt),0) as TotalShouldAmt, Nvl(sum(SO033.RealAmt),0) as TotalRealAmt," & _
                            "SO033.RealDate,SO033.CustId,SO033.MediabillNo,Max(Decode(CD013.RefNo,3,3,0)) RefNo " & _
                            " From " & strOwnerName & "SO033," & strOwnerName & "CD013 " & _
                            " Where SO033.UCCODE IS NOT NULL AND SO033.MediaBillNo = '" & vBillNo & "' And SO033.CancelFlag <> 1  And SO033.UCCode=CD013.CodeNo " & strSQL2 & _
                            " Group By SO033.CustId,SO033.RealDate,SO033.MediabillNo"
                Else
                    strSQL = "Select Nvl(sum(SO033.ShouldAmt),0) as TotalShouldAmt, Nvl(sum(SO033.RealAmt),0) as TotalRealAmt," & _
                             "SO033.RealDate,SO033.CustId,Max(Decode(CD013.RefNo,3,3,0)) RefNo From " & strOwnerName & "SO033," & strOwnerName & "CD013 " & _
                             " Where SO033.UCCODE IS NOT NULL AND SO033.BillNo = '" & vBillNo & "' And SO033.CancelFlag <> 1 And SO033.UCCode=CD013.CodeNo " & strSQL2 & _
                             " Group By SO033.CustId,SO033.RealDate"
                End If
            Else
                '問題集3008 將Like %拿掉,以增加效率,By Kin 2007/1/22
                'strSQL = "Select Nvl(sum(ShouldAmt),0) as TotalShouldAmt, Nvl(sum(RealAmt),0) as TotalRealAmt, RealDate, CustId From " & strOwnerName & "SO033 Where UCCODE IS NOT NULL AND CitiBankATM like '" & Mid(vBillNo, 1, 15) & "%' And CancelFlag <> 1 " & strSQL2 & " Group By CustId,RealDate"
                strSQL = "Select Nvl(sum(SO033.ShouldAmt),0) as TotalShouldAmt, Nvl(sum(SO033.RealAmt),0) as TotalRealAmt," & _
                        "SO033.RealDate,SO033.CustId,Max(Decode(CD013.RefNo,3,3,0)) RefNo From " & strOwnerName & "SO033," & strOwnerName & "CD013 " & _
                        " Where SO033.UCCODE IS NOT NULL AND SO033.CitiBankATM = '" & vBillNo & "' And SO033.CancelFlag <> 1 And SO033.UCCode=CD013.CodeNo " & strSQL2 & _
                        " Group By SO033.CustId,SO033.RealDate"
            End If
        End If
          '7179
        If blnCrossCustCombine Then
             strSQL = "Select A.*,SO001.AMduid, (Select MainCustId From " & strOwnerName & "SO202 Where MduId = SO001.AmduId) MainCustId " & _
                                 " From (" & strSQL & " ) A," & strOwnerName & "SO001 " & _
                                 " Where A.CustId = SO001.CustId "
        End If
        If Not GetRS(rsChk, strSQL, gcnGi) Then
           ChkData = False
           Exit Function
        End If
        '如果是中國信託入帳Log的單據編號為原始入帳號碼 921104 Liga
        '查無單號時,如果已經實收Log實收日期
        If rsChk.EOF Or rsChk.BOF Then
              blnSO3314 = False
              If intpara24 = 0 Then
                 If blnBillNo Then
                    strRealDate = GetRsValue("SELECT RealDate FROM " & strOwnerName & "SO033 WHERE BillNo='" & strOld_billno & "'") & ""
                 Else
                    '問題集3008 將Like %拿掉,以增加效率,By Kin 2007/1/22
                    'strRealDate = GetRsValue("SELECT RealDate FROM " & strOwnerName & "SO033 WHERE CitiBankATM like '%" & Mid(strOld_billno, 1, 15) & "%'") & ""
                    strRealDate = GetRsValue("SELECT RealDate FROM " & strOwnerName & "SO033 WHERE CitiBankATM = '" & strOld_billno & "'") & ""
                 End If
              Else
                 If blnBillNo Then
                    If blnMediaBillNo Then
                       strRealDate = GetRsValue("SELECT RealDate FROM " & strOwnerName & "SO033 WHERE MediaBillNo='" & strOld_billno & "'") & ""
                    Else
                       strRealDate = GetRsValue("SELECT RealDate FROM " & strOwnerName & "SO033 WHERE BillNo='" & strOld_billno & "'") & ""
                    End If
                 Else
                    '問題集3008 將Like %拿掉,以增加效率,By Kin 2007/1/22
                    'strRealDate = GetRsValue("SELECT RealDate FROM " & strOwnerName & "SO033 WHERE CitiBankATM like'%" & Mid(strOld_billno, 1, 15) & "%'") & ""
                    strRealDate = GetRsValue("SELECT RealDate FROM " & strOwnerName & "SO033 WHERE CitiBankATM ='" & strOld_billno & "'") & ""
                 End If
              End If

           If strRealDate = "" Then
                '950630 在SO3314入帳需回傳strErrorMsg,故此Function中全部的ErrFile.WriteLine皆須改寫
                'ErrFile.WriteLine ("單據編號： " & strOld_billno & " 查無此單號！" & "  代收收費處： " & vCM)
                strErrorMsg = "單據編號： " & strOld_billno & " 查無此單號！" & "  代收收費處： " & vCM
                ErrFile.WriteLine strErrorMsg
           Else
                strErrorMsg = "單據編號： " & strOld_billno & " 實收日期：" & strRealDate
                ErrFile.WriteLine strErrorMsg
           End If
           lngErrCount = lngErrCount + 1
           blnSO3314 = True
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
                    strErrorMsg = "單據編號： " & strOld_billno & " 相同單據編號但客戶編號不同！ 客編：" & strRetMsg
                    ErrFile.WriteLine (strErrorMsg)
                    
                    lngErrCount = lngErrCount + 1
                    blnSO3314 = True
                    GoTo XX
                End If
            End If
            '***********************************************************************
        
            '#3122 檢核是否有超過日結日期 By Kin Add 2007/03/29
            If Not ChkOverCloseDate(Format(vRealDate, "####/##/##"), objStorePara.CompCode, objStorePara.ServiceType, strReturn) Then
                strErrorMsg = strReturn
                ErrFile.WriteLine ("單據編號：" & vBillNo & "  實際扣款日：" & Format(str(Val(vRealDate) - 19110000), "##/##/##") & "  錯誤原因：" & strReturn)
                lngErrCount = lngErrCount + 1
                blnSO3314 = True
                GoTo XX
            End If
            '********************************************************************************************************************************
            '#4010 UCCode的參考號為3要視為異常資料 By Kin 2008/07/30
            If Val(rsChk("RefNo") & "") = 3 Then
                strErrorMsg = "單據編號： " & strOld_billno & " 客戶編號：" & rsChk("CustId") & " 櫃臺已收！請查核  代收收費處： " & vCM
                ErrFile.WriteLine strErrorMsg
                lngErrCount = lngErrCount + 1
                blnSO3314 = True
                GoTo XX
            End If
            '********************************************************************************************************************************
            If rsChk("RealDate").Value & "" <> "" And lngTotalRealAmt > 0 Then
                strErrorMsg = "單據編號： " & strOld_billno & " 客戶編號：" & strCustId & " 檔案內已實收金額$ " & lngTotalRealAmt & " ，請查核！" & "  代收收費處： " & vCM
                ErrFile.WriteLine strErrorMsg
                lngErrCount = lngErrCount + 1
                blnSO3314 = True
            Else
                If lngTotalShouldAmt <> Val(vRealAmt) Then
                    'strErrorMsg = "單據編號： " & strOld_billno & " 客戶編號：" & rsChk("CustId") & " 金額不合，檔案應收金額$ " & rsChk("TotalShouldAmt").Value & " ，磁片已收金額$ " & Val(vRealAmt) & "  代收收費處： " & vCM & "  收費單號：" & strBillNo_Old
                    strErrorMsg = "單據編號： " & strOld_billno & " 客戶編號：" & strCustId & " 金額不合，檔案應收金額$ " & lngTotalShouldAmt & " ，磁片已收金額$ " & Val(vRealAmt) & "  代收收費處： " & vCM & "  收費單號：" & strBillNo_Old
                    ErrFile.WriteLine strErrorMsg
                    lngErrCount = lngErrCount + 1
                    blnSO3314 = True
                Else
                    strSQL = "Select SO002.CustStatusCode,SO002.WipCode3,So002.ServiceType,SO033.BudgetPeriod," & _
                             "SO033.CitemCode,SO033.CompCode,SO033.FaciSeqNo,SO033.RealPeriod,SO033.CustId " & _
                             " From " & strOwnerName & "So002," & strOwnerName & "SO033 " & _
                             " Where SO002.CustId = SO033.CustId And SO002.CompCode = SO033.CompCode " & _
                             " And SO002.ServiceType=SO033.ServiceType "
                    
                    If intpara24 = 0 Then
                        If blnBillNo Then
                           strSQL = strSQL & " And BillNo='" & strOld_billno & "'"
                        Else
                        '問題集3008 將Like %拿掉,以增加效率,By Kin 2007/1/22
                           'strSQL = strSQL & " And CitiBankATM like '%" & Mid(strOld_billno, 1, 15) & "%'"
                           strSQL = strSQL & " And CitiBankATM = '" & strOld_billno & "'"
                        End If
                        
                        'strSQL = "Select CustStatusCode,WipCode3,ServiceType From " & strOwnerName & "So002 Where CustId = " & rsChk("CustId").Value & " And ServiceType= '" & rsChk("ServiceType") & "'"
                    Else
                        If blnBillNo Then
                            If blnMediaBillNo Then
                                strSQL = strSQL & " And SO033.MediaBillNo = '" & vBillNo & "'"
                            Else
                                '7179
                                strSQL = strSQL & " And SO033.BillNo = '" & vBillNo & "'" & IIf(blnCrossCustCombine, " And SO033.CustId = " & strCustId, "")
                            End If
                        Else
                            '問題集3008 將Like %拿掉,以增加效率,By Kin 2007/1/22
                            'strSQL = strSQL & " And SO033.CitiBankATM like '" & Mid(vBillNo, 1, 15) & "%'"
                            '7179
                            strSQL = strSQL & " And SO033.CitiBankATM = '" & vBillNo & "'" & IIf(blnCrossCustCombine, " And SO033.CustId = " & strCustId, "")
                        End If
                    End If
                    If Not GetRS(rsCust, strSQL, gcnGi) Then Exit Function
                    Do While Not rsCust.EOF
                        '非正常客戶則先作log再入帳
                        If rsCust("CustStatusCode").Value & "" <> 1 Then
                            'strErrorMsg = "客戶編號: " & rsChk("CustId").Value & " 單據編號： " & strOld_billno & " 為非正常客戶 " & " 客戶狀態: " & rsCust("CustStatusCode").Value & " 代收收費處: " & vCM & " 服務別: " & rsCust("ServiceType").Value
                            strErrorMsg = "客戶編號: " & strCustId & " 單據編號： " & strOld_billno & " 為非正常客戶 " & " 客戶狀態: " & rsCust("CustStatusCode").Value & " 代收收費處: " & vCM & " 服務別: " & rsCust("ServiceType").Value
                            ErrFile.WriteLine strErrorMsg
                            lngStatusCount = lngStatusCount + 1
                        '正常客戶且CM關機中
                        ElseIf rsCust("CustStatusCode").Value & "" = 1 And rsCust("WipCode3") = 16 Then
                            'strErrorMsg = "客戶編號: " & rsChk("CustId").Value & " 單據編號： " & strOld_billno & " 為CM關機中客戶 " & " 代收收費處: " & vCM & " 服務別: " & rsCust("ServiceType").Value
                            strErrorMsg = "客戶編號: " & strCustId & " 單據編號： " & strOld_billno & " 為CM關機中客戶 " & " 代收收費處: " & vCM & " 服務別: " & rsCust("ServiceType").Value
                            ErrFile.WriteLine strErrorMsg
                            lngStatusCount = lngStatusCount + 1
                        End If
'                        If Not GetRS(rsTmp, "Select BudgetPeriod From " & strOwnerName & "SO003 Where Custid = " & rsCust("CustId") & " And CompCode = " & rsCust("CompCode") & " And ServiceType = '" & rsCust("ServiceType") & "' " & _
'                            " And CitemCode = " & rsCust("CitemCode") & " And FaciSeqNo = '" & rsCust("FaciSeqNo") & "'") Then Exit Function
'                        If Not rsTmp.EOF Then
'                            If rsCust("RealPeriod") > rsTmp.Fields("BudgetPeriod") And rsTmp.Fields("BudgetPeriod") > 0 Then
'                                strErrorMsg = "客戶編號: " & rsChk("CustId").Value & " 單據編號： " & strOld_billno & "入帳期數大於分期剩餘期數,無法入帳"
'                                ErrFile.WriteLine strErrorMsg
'                                lngErrCount = lngErrCount + 1
'                                blnSO3314 = True
'                                GoTo XX
'                            End If
'                        End If
                        
                        rsCust.MoveNext
                    Loop
                    '執行更新SO033
                    If UpdAccount(vBillNo, vRealAmt, vClctEn, vClctName, vRealDate, vUpdEn, vCMCode, vCMName, vPTCode, vPTName, blnBillNo, blnMediaBillNo, strCitemQry) = False Then
                        ChkData = False
                        Exit Function
                    End If
                End If
           End If
        End If
XX:
        If rsChk.State = adStateOpen Then rsChk.Close
        Set rsChk = Nothing
        If rsCust.State = adStateOpen Then rsCust.Close
        Set rsCust = Nothing
        
        ChkData = True
    Exit Function
ChkErr:
    ErrSub "DealFun", "ChkData"
End Function
'#5386 判斷那幾筆客編不相同 By Kin 2009/12/04
Private Function ChkDiffCust(ByRef rs As ADODB.Recordset) As String
  On Error GoTo ChkErr
    Dim strRetMsg As String
    Dim strDiffCust As String
    Dim strRealDate As String
    Dim strRetCustId As String
    Dim strRetRealDate As String
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
  ErrSub "UpdateFun", "ChkDiffCust"
End Function
Public Function UpdAccount(ByVal vBillNo As String, ByVal vRealAmt As String, _
                           ByVal vClctEn As String, ByVal vClctName As String, _
                           ByVal vRealDate As String, ByVal vUpdEn As String, _
                           ByVal vCMCode As String, ByVal vCMName As String, _
                           ByVal vPTCode As String, ByVal vPTName As String, _
                           Optional blnBillNo As Boolean = False, _
                           Optional blnMediaBillNo As Boolean = False, Optional strCitemQry As String) As Boolean
    On Error GoTo ChkErr
    Dim rsAccount As New ADODB.Recordset
    Dim strSQL As String, strSeqQry As String, lngAffected As Long
    Dim lngPeriod As Long, lngAmount As Long
    Dim strSo003 As String, strUpdSql As String
    Dim strSQL2 As String, rsTmp As New ADODB.Recordset
        '收費項目
        If strCitemQry <> "" Then strSQL2 = " And CitemCode In (" & strCitemQry & ")"
        If blnBillNo Then
            If blnMediaBillNo Then
             strSQL = "Select RowId, So033.* From " & strOwnerName & "So033 Where UCCODE IS NOT NULL  And MediaBillNo ='" & vBillNo & "' AND CancelFlag <> 1" & strSQL2
'            strSQL = "Select RowId, CustId, ShouldAmt, CitemCode, RealStartDate, RealStopDate, RealPeriod, OldPeriod, OldAmt, ClctEn, ClctName, CMCode, CMName, SeqNo, PackageNo, PackageName, BudgetPeriod From " & strOwnerName & "So033 Where UCCODE IS NOT NULL  And MediaBillNo ='" & vBillNo & "' AND CancelFlag <> 1" & strSQL2
            Else
                strSQL = "Select RowId,So033.* From " & strOwnerName & "So033 Where UCCODE IS NOT NULL  And BillNo ='" & vBillNo & "' AND CancelFlag <> 1" & strSQL2
            End If
        Else
            '問題集3146 原Update有Like，現在改為=  by Kin 2007/3/28
            strSQL = "Select RowId,So033.* From " & strOwnerName & "So033 Where UCCODE IS NOT NULL  And CitiBankATM = '" & vBillNo & "' AND CancelFlag <> 1" & strSQL2
            'strSQL = "Select RowId,So033.* From " & strOwnerName & "So033 Where UCCODE IS NOT NULL  And CitiBankATM like '" & Mid(vBillNo, 1, 15) & "%' AND CancelFlag <> 1" & strSQL2
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
                .Fields("PTCode") = IIf(vPTCode = "", Null, vPTCode)
                .Fields("PTName") = NoZero(vPTName)
                .Fields("ClctEn") = IIf(Trim(vClctEn) = "", rsAccount("ClctEn"), vClctEn)
                .Fields("ClctName") = IIf(Trim(vClctName) = "", rsAccount("ClctName"), vClctName)
               .Fields("FirstTime") = GetDTString(strNowTime)
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
            
'            If Not GetRS(rsTmp, "Select * From " & strOwnerName & "SO033 Where RowId = '" & rsAccount("RowId") & "'") Then Exit Function
'            If AlterSO003(rsTmp) = False Then
'                   UpdAccount = False
'                   Exit Function
'            End If
lLoop:
            rsAccount.MoveNext
        Wend
            
            
'            If intPara8 >= 1 Then
'                lngPeriod = NoZero(rsAccount("RealPeriod"), True)
'                If intPara9 = 1 Then
'                    lngAmount = NoZero(rsAccount("OldAmt"), True)
'                    lngPeriod = NoZero(rsAccount("OldPeriod"), True)
'                ElseIf intPara9 = 2 Then
'                    lngAmount = NoZero(rsAccount("ShouldAmt"), True)
'                Else
'                    lngAmount = NoZero(rsAccount("ShouldAmt"), True)
'                End If
'                strUpdSql = ",Amount = " & lngAmount & ",Period = " & lngPeriod
'            End If
'
'
'            If rsAccount("SeqNo") & "" <> "" Then strSeqQry = " And SeqNo = " & rsAccount("SeqNo")
'
'            If Val(rsAccount("BudgetPeriod") & "") > 0 Then
'                If Val(rsAccount("BudgetPeriod") & "") = 1 Then
'                    If Not ExecuteCommand("Delete From SO003 Where CustId = '" & rsAccount("CustId") & "' And CitemCode = '" & rsAccount("CitemCode") & "' And StopDate < " & _
'                        GetNullString(rsAccount("RealStopDate"), giDateV, giOracle) & strSeqQry) Then Exit Function
'                    GoTo lLoop
'                End If
'                strSo003 = ",BudgetPeriod = " & (rsAccount("BudgetPeriod") - 1)
'            End If
'
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
        Call CloseRecordset(rsTmp)
        UpdAccount = True
    Exit Function
ChkErr:
    UpdAccount = False
End Function

Public Function InitData() As Boolean
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
  ErrSub "DealFun", "InitData"
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
  ErrSub "DealFun", "MsgResult"
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
    objLst.FilterStop = blnStopFlag
    objLst.SendConn gcnGi
    objLst.Filter = strWhere
  Exit Sub
ChkErr:
    Call ErrSub("DealFun", "SetLst")
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

Public Function GetBillNo_New(ByVal strNo As String, Optional blnMediaBillNo As Boolean) As String
    On Error Resume Next
    Dim strType As String
    Dim strYM As String
    Dim strMM As String
    
'        strNo = Mid(strReadLine, 65, 11)
        If blnMediaBillNo Then
            GetBillNo_New = strNo
        Else
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
        End If
End Function

Public Function GetBillNo_Old(ByVal strNo As String, ByVal strServiceType, _
    Optional blnMediaBillNo As Boolean) As String
    On Error Resume Next
        If blnMediaBillNo Then
            GetBillNo_Old = strNo
        Else
            GetBillNo_Old = Trim(CStr(Val(Left(strNo, 4)) + 191100)) & _
                    Mid(strNo, 5, 1) & strServiceType & Format(Right(strNo, 6), "0000000")
        End If
End Function

Public Function GetBankATM(strBillNo As String, lngCustId As Long, strBankCode As String, _
    strShouldDate As String, lngShouldAmt As Long, Optional ByRef clsBankATMNo As Object, _
    Optional ByRef strCorpId As String) As String
    On Error Resume Next
    Dim rsTmp As New ADODB.Recordset
    Dim strShDate As String
        If clsBankATMNo Is Nothing Then
            If strBankCode = "" Or strBillNo = "" Or lngCustId = 0 Then Exit Function
            If Not GetRS(rsTmp, "Select FileName2,CorpID From " & strOwnerName & "CD018 Where CodeNo = " & strBankCode) Then Exit Function
            If Len(rsTmp("FileName2") & "") = 0 Or Len(rsTmp("CorpId") & "") = 0 Then MsgBox "銀行代碼檔中無定義轉帳輸入檔名或收件單位代號", vbCritical, "警告": Exit Function
            strCorpId = rsTmp("CorpId") & ""
            Set clsBankATMNo = CreateObject("csBankATMNo.cls" & rsTmp("FileName2"))
            If Err.Number = 429 Then MsgBox "銀行代碼檔中定義之轉帳輸入檔名 : [" & rsTmp("FileName2") & "] 不正確,請確認!!", vbCritical, "警告": Exit Function
        End If
        strShDate = Format(strShouldDate, "yyyy/mm/dd")
        GetBankATM = clsBankATMNo.GetATMNo(strBillNo, lngCustId, strCorpId, strShDate, lngShouldAmt, gcnGi)
End Function

'#3122 檢核是否有超過日結日期 By Kin Add 2007/03/29
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

