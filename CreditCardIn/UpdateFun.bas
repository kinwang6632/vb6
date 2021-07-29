Attribute VB_Name = "UpdateFun"
Option Explicit
Private Const FormName = "UpdateFun"
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Const gMsgIsDataOK = "此為必要欄位,須有值 !! "
Public lngCount As Long           '記錄成功筆數
Public lngErrCount As Long        '記錄錯誤筆數
Public lngStatusCount As Long     '記錄非正常戶筆數
Public lngTime As Long             '記錄所花時間
Public strNowTime As String        '記錄目前時間
Public fso As New FileSystemObject
Public File As TextStream
Public ErrFile As TextStream     '錯誤資料log檔
Private intPara3 As Integer
Private intPara5 As Integer
Private intPara8 As Integer
Public gcnGi As New ADODB.Connection
Public strUserName As String
Private strFilePath As String
Public objStorePara As clsStoreParameter
Public Enum giArrow
    giLeft = 0
    giRight = 1
End Enum
Public Enum ChangeDB
    giAccessDb = 0 'Access Table
    giOracle = 1    'Oracle Table
End Enum
Enum giVType
    giLongV = 0
    giStringV = 1
    giDateV = 2
End Enum

Public blnMultiAccount As Boolean     ''  記錄 SO043 的 SO043.para24  ，以判斷是否用多帳戶設定
Public strOwnerName As String
Public Enum mCreditCardType
    lngUnionChinese = 0
    lngChinatrust = 1
    lngHSBC = 2
End Enum

Public GetOwner As String


'接受上層傳的檔案路徑
Public Function GetFilePath(ByVal vFilePath As String)
  On Error GoTo ChkErr
     strFilePath = vFilePath
  Exit Function
ChkErr:
  ErrSub "DealFun", "GetFilePath"
End Function

Public Sub ErrSub(FormName As String, ProcedureName As String)
    Dim strErr As String
    Dim aryErr As Variant
    Dim strCrLf As String
    Dim RetVal As Variant
    DoEvents
    Screen.MousePointer = vbDefault
    aryErr = Array("發生錯誤時間 : ", "發生錯誤專案 : ", "發生錯誤表單 : ", "發生錯誤程序 : ", "錯誤代碼 : ", "錯誤原因 : ")
    strCrLf = vbCrLf + vbCrLf
    strErr = aryErr(0) + CStr(Now) + strCrLf + _
             aryErr(1) + CStr(Err.Source) + strCrLf + _
             aryErr(2) + FormName + strCrLf + _
             aryErr(3) + ProcedureName + strCrLf + _
             aryErr(4) + CStr(Err.Number) + strCrLf + _
             aryErr(5) + CStr(Err.Description)
    ErrLog strErr
    RetVal = MsgBox(strErr, vbYesNo, "系統執行錯誤..(按 <是> 列印錯誤訊息)!!")
    If RetVal = vbYes Then
       With Printer
           .FontSize = 14
            Printer.Print "": Printer.Print "": Printer.Print ""
            Printer.Print strErr
           .NewPage
       End With
    End If

End Sub

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

Public Function GetRsValue(strSQL As String, Optional CnnObj As ADODB.Connection, Optional MsgNoData As String) As Variant
  On Error GoTo ChkErr
    Dim rs As New ADODB.Recordset
    If CnnObj Is Nothing Then Set CnnObj = gcnGi
    If Not GetRS(rs, strSQL, CnnObj) Then Exit Function
    If rs.RecordCount = 0 Then
        If Len(MsgNoData) > 0 Then
            MsgBox MsgNoData, vbExclamation, "提示!"
        End If
        GetRsValue = Null
    Else
        GetRsValue = rs(0)
    End If
  Exit Function
ChkErr:
    ErrSub "DealFun", "GetRsValue"
End Function

Public Function GetString(ByVal strS As Variant, ByVal intLength As Integer, Optional ByVal lngArrow As giArrow = giLeft, Optional ByVal InsertZero As Boolean = False) As String
  On Error GoTo ChkErr
    Dim strCH As String
    strCH = CStr(strS & "")
    If lngArrow = giLeft Then
        strCH = StrConv(LeftB(StrConv(strCH & Space(intLength), vbFromUnicode), intLength), vbUnicode)
    Else
        strCH = StrConv(RightB(StrConv(Space(intLength) & strCH, vbFromUnicode), intLength), vbUnicode)
    End If
    If InsertZero Then
    If Trim(strCH) = "" Then strCH = "0"
        If lngArrow = giLeft Then
            strCH = RTrim(strCH) & String(intLength - Len(RTrim(strCH)), "0")
        Else
            strCH = String(intLength - Len(LTrim(strCH)), "0") & LTrim(strCH)
        End If
    End If
    strCH = Replace(strCH, Chr(0), "")
    GetString = strCH
  Exit Function
ChkErr:
    Call ErrSub("Sys_Lib", "GetString")
End Function

Public Function ChkData(ByVal vBillNo As String, ByVal vRealAmt As String, _
                        ByVal vRealDate As String, ByVal vUpdEn As String, _
                        ByVal vCMCode As String, ByVal vCMName As String, _
                        ByVal vPTCode As String, ByVal vPTName As String, _
                        Optional pEmpNO As String = "", Optional pEmpName As String = "", _
                        Optional pAuthorizeNo As String = " ", Optional p_frmCreditCardType As mCreditCardType = lngUnionChinese, _
                        Optional pCitemCodeChkData As String, Optional pstrServiceType As String) As Boolean
                        
  On Error GoTo ChkErr
  Dim strSQL As String
  Dim rsChk As New ADODB.Recordset
  Dim rsCust As New ADODB.Recordset
  Dim rsTmp As New ADODB.Recordset
  Dim strBillField As String
  Dim strWhere As String
  Dim rsCheckBillNoCancelFlag As New ADODB.Recordset
  Dim lngBudgetPeriod As Long
  
  '''以下的 BillNo 欄位 改成 MediaBillNo
  
  If blnMultiAccount = True Then
       strBillField = "MediaBillno"
  Else
       strBillField = "Billno"
  End If
  
     strSQL = "Select " & _
                 "Nvl(sum(ShouldAmt),0) as TotalShouldAmt, Nvl(sum(RealAmt),0) as TotalRealAmt, RealDate, CustId " & _
                 "From " & _
                 frmUnion.TableOwnerName & "SO033 " & _
                 "Where UCCODE IS NOT NULL AND " & _
                 strBillField & " = '" & vBillNo & "' And CancelFlag <> 1 Group By CustId,RealDate"
    
    If rsChk.State = adStateOpen Then rsChk.Close
  
    If Not GetRS(rsChk, strSQL, gcnGi) Then
       ChkData = False
       Exit Function
    End If
   
    If rsChk.EOF Or rsChk.BOF Then
                       
                        ''2004/09  加入UCCODE 已收的判斷式 以提供更明確的訊息
                         strSQL = "Select count(CustId) C " & _
                                        "From " & _
                                        frmUnion.TableOwnerName & "SO033 " & _
                                        "Where " & _
                                        strBillField & " = '" & vBillNo & "' And CancelFlag <> 1 Group By CustId,RealDate"
                       ''  2004/09/13 加入多一層的判斷式 除止 bof eof
                        If GetRS(rsCheckBillNoCancelFlag, strSQL, gcnGi, adUseClient, adOpenDynamic) = True Then
                             If rsCheckBillNoCancelFlag.EOF Or rsCheckBillNoCancelFlag.BOF Then
                                     ErrFile.WriteLine ("單據編號： " & vBillNo & " 查無此單號！")
                             Else
                                     ErrFile.WriteLine ("單據編號： " & vBillNo & " 此單號的單據已收！")
                             End If
                        Else
                              ErrFile.WriteLine ("單據編號： " & vBillNo & " 查無此單號！")
                        End If
                        
''                        If RPxx(gcnGi.Execute(strsql).GetString) > 0 Then   '' 這一行如果是emc會出問題所以上面再加一個eof、bof的判斷
''                                    ErrFile.WriteLine ("單據編號： " & vBillNo & " 此單號的單據已收！")
''                        Else
''                                    ErrFile.WriteLine ("單據編號： " & vBillNo & " 查無此單號！")
''                        End If
                        lngErrCount = lngErrCount + 1
                        
    Else
                        If rsChk("RealDate").Value & "" <> "" And rsChk("TotalRealAmt").Value > 0 Then
                                ErrFile.WriteLine ("客戶編號: " & rsChk("CustId").Value & "單據編號： " & vBillNo & " 檔案內已包含實收金額$ " & rsChk("TotalRealAmt").Value & " ，請查核！")
                                lngErrCount = lngErrCount + 1
                        Else
                                  ''   2003/10/13 中國信託信用卡
                                 If rsChk("TotalShouldAmt").Value <> Val(vRealAmt) And p_frmCreditCardType <> 1 Then
                                                 ErrFile.WriteLine ("客戶編號: " & rsChk("CustId").Value & "單據編號： " & vBillNo & " 金額不合，檔案應收金額$ " & rsChk("TotalShouldAmt").Value & " ，磁片已收金額$ " & Val(vRealAmt))
                                                 lngErrCount = lngErrCount + 1
                                 Else
                                                  If rsCust.State = adStateOpen Then rsCust.Close
                                                                
                                                                strSQL = "Select CustStatusCode,WipCode3 ,CustStatusName From " & frmUnion.TableOwnerName & "So002 Where CustId = " & rsChk("CustId").Value & "  AND SERVICETYPE  = '" & pstrServiceType & "'"
                                                                Call GetRS(rsCust, strSQL, gcnGi)
                                                                '非正常客戶則先作log再入帳
                                                                If rsCust("CustStatusCode").Value & "" <> 1 Then
                                                                                   ErrFile.WriteLine ("客戶編號: " & rsChk("CustId").Value & " 單據編號： " & vBillNo & " 為非正常客戶 " & " 客戶狀態: " & rsCust("CustStatusName").Value)
                                                                                   lngStatusCount = lngStatusCount + 1
                                                                '正常客戶且CM關機中
                                                                ElseIf rsCust("CustStatusCode").Value & "" = 1 And rsCust("WipCode3") = 16 Then
                                                                                   ErrFile.WriteLine ("客戶編號: " & rsChk("CustId").Value & " 單據編號： " & vBillNo & " 為CM關機中客戶 ")
                                                                                   lngStatusCount = lngStatusCount + 1
                                                                End If
                                                                If Not GetRS(rsTmp, "Select RealPeriod,CustId,CompCode,ServiceType,CitemCode,FaciSeqNo From " & strOwnerName & "So033 Where UCCODE IS NOT NULL AND " & _
                                                                        strBillField & " = '" & vBillNo & "' And CancelFlag <> 1 ") Then Exit Function
                                                                '94/09/30 1805 Jacky
                                                                Do While Not rsTmp.EOF
                                                                    lngBudgetPeriod = Val(GetRsValue("Select Nvl(BudgetPeriod,0) From " & strOwnerName & "SO003 Where Custid = " & rsTmp("CustId") & " And CompCode = " & rsTmp("CompCode") & " And ServiceType = '" & rsTmp("ServiceType") & "' " & _
                                                                        " And CitemCode = " & rsTmp("CitemCode") & " And FaciSeqNo = '" & rsTmp("FaciSeqNo") & "'") & "")
                                                                    If rsTmp("RealPeriod") > lngBudgetPeriod And lngBudgetPeriod > 0 Then
                                                                        Call ErrFile.WriteLine("入帳期數大於分期剩餘期數,無法入帳!!")
                                                                        lngErrCount = lngErrCount + 1
                                                                        ChkData = True
                                                                        Exit Function
                                                                    End If
                                                                    rsTmp.MoveNext
                                                                Loop
                                         
                                                                '執行更新SO033
                                                                If UpdAccount( _
                                                                                        vBillNo, vRealAmt, vRealDate, vUpdEn, vCMCode, vCMName, _
                                                                                        vPTCode, vPTName, pEmpNO, pEmpName, pAuthorizeNo, _
                                                                                        pCitemCodeChkData) = False Then
                                                                                   ChkData = False
                                                                                   Exit Function
                                                                End If  '' If UpdAccount
                                                   End If
                                 End If
    End If
    If rsChk.State = adStateOpen Then rsChk.Close
    Set rsChk = Nothing
    If rsCust.State = adStateOpen Then rsCust.Close
    Set rsCust = Nothing
    Call CloseRecordset(rsTmp)
    ChkData = True
  Exit Function
ChkErr:
  ErrSub "DealFun", "ChkData"
End Function

Public Function UpdAccount(ByVal vBillNo As String, ByVal vRealAmt As String, _
                           ByVal vRealDate As String, ByVal vUpdEn As String, _
                           ByVal vCMCode As String, ByVal vCMName As String, _
                           ByVal vPTCode As String, ByVal vPTName As String, _
                           Optional pEmpNO As String = "", Optional pEmpName As String = "", _
                           Optional pAuthorizeNo As String, Optional pCitemCoddSql As String = "") As Boolean

 '' vBillNo 單據編號
 '' vClctEn 收費人員
 '' vRealDate 實收日期
 '' vCMCode 收費方式代碼  ref CD031''
 '' vPTCode 付款種類  ref CD032''
On Error GoTo ChkErr

 Dim rsAccount As New ADODB.Recordset
 Dim strSQL As String
 Dim strUpdateEmpSql As String
 Dim strBillField As String
 Dim intBudgetPeriod As Integer
 Dim strBudgetPeriodSQL As String
 ''  2003/08/18  這個變數定義收集目前調整的這一筆SO033  ，將其傳給貴仔的FUNCTION AlterSO003
 Dim rsAlterSO003 As New ADODB.Recordset
 Dim strWhereCitemCode As String
 
strWhereCitemCode = IIf(Len(pCitemCoddSql) > 0, " AND CitemCode " & pCitemCoddSql & "  ", " ")
 If blnMultiAccount = True Then
      strBillField = "MediaBillno"
 Else
      strBillField = "Billno"
 End If
    '' 這一行的WHERE  後面的 Billno   改成MediaBillno 以適合多帳號功能
   ' strsql = "Select RowId, CustId, ShouldAmt, CitemCode, RealStartDate, RealStopDate, RealPeriod, ClctEn, ClctName, CMCode, CMName,BudgetPeriod,AccountNO, REALAMT From  " & frmUnion.TableOwnerName & "So033    Where UCCODE IS NOT NULL  And " & strBillField & " ='" & vBillNo & "' AND CancelFlag <> 1 " & strWhereCitemCode
        strSQL = "SELECT  *   From  " & frmUnion.TableOwnerName & "So033    Where UCCODE IS NOT NULL  And " & strBillField & " ='" & vBillNo & "' AND CancelFlag <> 1 " & strWhereCitemCode
    
    If Not GetRS(rsAccount, strSQL, gcnGi, adUseClient, adOpenDynamic, adLockOptimistic) Then
       UpdAccount = False
       Exit Function
    End If
    If rsAccount.RecordCount > 0 Then rsAccount.MoveFirst
    
    '' """"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
    '' // 2002/11/21       -- 新增內容--
    '' // 如果有傳入收費人員
    '' // 則將收費人員的資料一併更新所需的 SQL 語法串起來  --
    '' -------------------------------------------------------------------------------------------------------------
    If Len(pEmpNO) <> 0 Then
         strUpdateEmpSql = ",ClctEn = '" & pEmpNO & "', ClctName ='" & pEmpName & "'"
    End If
    '' """"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""

    While Not rsAccount.EOF

       
        With rsAccount
            .Fields("RealAmt") = rsAccount("ShouldAmt")
            .Fields("RealDate") = NoZero(Format(vRealDate, "####/##/##"))
            .Fields("Updtime") = GetDTString(strNowTime)
            .Fields("UpdEn") = vUpdEn
            .Fields("UCCode") = Null
            .Fields("UCName") = Null
            .Fields("CMCode") = vCMCode
            .Fields("CMName") = vCMName
            .Fields("PTCode") = vPTCode
            .Fields("PTName") = vPTName
             If Len(pEmpNO) <> 0 Then
                   .Fields("ClctEn") = pEmpNO
                   .Fields("ClctName") = pEmpName
            End If
           .Fields("FirstTime") = GetDTString(strNowTime)
           .Fields("AuthorizeNo") = pAuthorizeNo
        
            Call AlterSO003(rsAccount)
            .Update
        End With
'         strsql = "Update " & frmUnion.TableOwnerName & "So033 Set RealAmt =" & rsAccount("ShouldAmt") & _
'                                 ",RealDate = To_date('" & vRealDate & "','YYYYMMDD')" & _
'                                  ",Updtime = '" & GetDTString(strNowTime) & "'" & _
'                                  ",UpdEn = '" & vUpdEn & "'" & _
'                                  ",UCCode = Null, UCName = Null" & _
'                                  ", CMCode = '" & vCMCode & "'" & _
'                                  ", CMName = '" & vCMName & "'" & _
'                                  ", PTCode = '" & vPTCode & "'" & _
'                                  ", PTName = '" & vPTName & "'" & _
'                                  strUpdateEmpSql & _
'                                  ",FirstTime = '" & GetDTString(strNowTime) & "'," & _
'                                  "AuthorizeNo='" & pAuthorizeNo & "' " & _
'                                  "Where Rowid = '" & rsAccount("Rowid") & "'"
'          gcnGi.Execute strsql
           rsAccount.MoveNext
    Wend

    lngCount = lngCount + 1
    Set rsAccount = Nothing
    UpdAccount = True
 Exit Function
ChkErr:
   UpdAccount = False
   ErrSub "UpdateFun", "UpdAccount"
End Function


'根據傳來之系統內定時間字串轉成固定西曆時間字串
'Parameter1: DateTime
'Return: Format DateTime String
Public Function GetDTString(dtm As String) As String
  On Error GoTo ChkErr
    GetDTString = Format(dtm, "ee/mm/dd hh:mm:ss")
  Exit Function
ChkErr:
   ErrSub "DealFun", "GetDTString"
End Function

Public Function GetNullString(varValue As Variant, _
                     Optional lngVType As giVType = giStringV, _
                     Optional DBFlag As ChangeDB = giOracle, _
                     Optional blnAddTime As Boolean = False) As Variant
    On Error GoTo ChkErr
        If lngVType = giDateV Then
            If Len(varValue & "") = 0 Then
                GetNullString = "NULL"
            Else
                If DBFlag = giOracle Then
                    If blnAddTime Then
                       GetNullString = "To_Date('" & Format(varValue, "YYYYMMDDHHMMSS") & "','YYYYMMDDHH24MISS')"
                    Else
                       GetNullString = "To_Date('" & Format(varValue, "YYYYMMDD") & "','YYYYMMDD')"
                    End If
                Else
                    If blnAddTime Then
                        GetNullString = "#" & Format(varValue, "YYYY/MM/DD HH:MM:SS") & "#"
                    Else
                        GetNullString = "#" & Format(varValue, "YYYY/MM/DD") & "#"
                    End If
                End If
            End If
        ElseIf lngVType = giLongV Then
            If Len(varValue & "") = 0 Then
                GetNullString = "NULL"
            Else
                GetNullString = varValue
            End If
        ElseIf lngVType = giStringV Then
            If Len(varValue & "") = 0 Then
                GetNullString = "NULL"
            Else
                varValue = Replace(varValue, "'", "''")
                GetNullString = "'" & varValue & "'"
            End If
        End If
    Exit Function
ChkErr:
  Call ErrSub("Utility5", "GetNullString")
End Function

Public Function InitData() As Boolean
On Error GoTo ChkErr
   Dim rsTmp As New ADODB.Recordset
   Dim strSQL As String
   
   strSQL = "Select Para3, Para5, Para8 From " & frmUnion.TableOwnerName & "So043"
   If Not GetRS(rsTmp, strSQL, gcnGi) Then InitData = False: Exit Function
   If rsTmp.EOF = False Then
      intPara3 = rsTmp("Para3")
      intPara5 = rsTmp("Para5")
      intPara8 = rsTmp("Para8")
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
                  Optional strTableName As String, _
                  Optional strWhere As String)
  On Error GoTo ChkErr
    objLst.SetFldName1 strFldName1
    objLst.SetFldName2 strFldName2
    If lngFldLen1 > 0 Then objLst.FldLen1 = lngFldLen1
    If lngFldLen2 > 0 Then objLst.FldLen2 = lngFldLen2
    If lngWidth1 > 0 Then objLst.FldWidth1 = lngWidth1
    If lngWidth2 > 0 Then objLst.FldWidth2 = lngWidth2
    objLst.SetTableName strTableName
    objLst.SendConn gcnGi
    objLst.Filter = strWhere
  Exit Sub
ChkErr:
    Call ErrSub("DealFun", "SetLst")
End Sub

Public Function GetRS(ByRef rs As ADODB.Recordset, _
                      strSQL As String, _
                      Optional Conn As ADODB.Connection, _
                      Optional CursorLocation As CursorLocationEnum = adUseClient, _
                      Optional CursorType As CursorTypeEnum = adOpenForwardOnly, _
                      Optional LockType As LockTypeEnum = adLockReadOnly, _
                      Optional ErrMessage As String, _
                      Optional Subroutune As String, _
                      Optional DirectUnload As Boolean = False, _
                      Optional MaxRecords As String) As Boolean
  On Error GoTo ChkErr
    If rs.State = adStateOpen Then rs.Close
    rs.CursorLocation = CursorLocation
    If Val(MaxRecords) > 0 Then rs.MaxRecords = Val(MaxRecords)
    If Conn Is Nothing Then Set Conn = gcnGi
    rs.Open strSQL, Conn, CursorType, LockType
    GetRS = True
  Exit Function
ChkErr:
    GetRS = False
    MsgBox strSQL
    ErrSub "UpdateFun", "GetRs"
    On Error Resume Next
    If DirectUnload Then Unload Screen.ActiveForm
End Function

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
Public Function RPxx(strValue As String) As String
  On Error Resume Next
    strValue = Trim(Replace(strValue, Chr(13), "", 1, , vbTextCompare))
    strValue = Trim(Replace(strValue, Chr(10), "", 1, , vbTextCompare))
    RPxx = Trim(Replace(strValue, Chr(9), "", 1, , vbTextCompare))
End Function
Public Function MidMbcs(ByVal str As String, start, Length)
  On Error GoTo ChkErr
    MidMbcs = StrConv(MidB(StrConv(str, vbFromUnicode), start, Length), vbUnicode)
  Exit Function
ChkErr:
    ErrSub "Sys_Lib", "MidMbcs"
End Function

''2003/08/18 這一段新增週期性收費項目

'Public Function AlterSO003(rsSO033 As ADODB.Recordset, Optional ByVal blnAddSo003 As Boolean = True, _
'                Optional ByVal blnUpdAmt As Boolean = True, Optional ByVal blnUpdClctDate As Boolean = True, _
'                Optional ByVal blnChkMustAnswer As Boolean = False, Optional ByRef strSeqNo As String, _
'                Optional ByVal blnMultiCount As Boolean = False) As Boolean
'    On Error GoTo ChkErr
'
'    Dim rsAlter As New ADODB.Recordset
'    Dim rsTmp As New ADODB.Recordset
'    Dim lngPeriodFlag As Boolean
'    Dim Ypara5 As Long, Ypara8 As Long
'    Dim Ypara9 As Long, YPara24 As Long
'    Dim lngCompCode As Long, lngBad As Long
'    Dim blnFlag As Boolean
'    Dim blnDiscount As Boolean
'
'        '判斷是否為週期性收費項目
'
'        lngPeriodFlag = Val(GetRsValue("Select PeriodFlag From " & frmUnion.TableOwnerName & "Cd019 Where CodeNo = " & rsSO033("CitemCode")) & "")
'        If Len(rsSO033("STCode") & "") > 0 Then lngBad = Val(GetRsValue("Select RefNo From " & frmUnion.TableOwnerName & "CD016 Where CodeNo = " & rsSO033("STCode")) & "")
'        If lngPeriodFlag = 0 Or (lngBad = 1) Then
'            AlterSO003 = True
'            Exit Function
'        End If
'
'        If Not GetSystemPara(rsTmp, rsSO033("CompCode") & "", rsSO033("ServiceType"), "SO043", "Para5,Para8,Para9,Para24") Then Exit Function
'        If rsTmp.RecordCount = 0 Then Exit Function
'        Ypara5 = rsTmp("Para5")
'        Ypara8 = rsTmp("Para8")
'        Ypara9 = rsTmp("Para9")
'        YPara24 = rsTmp("Para24")
'        Call CloseRecordset(rsTmp)
'
'        lngCompCode = Val(GetRsValue("Select CompCode From " & frmUnion.TableOwnerName & "So001 Where CustId = " & rsSO033("CustId")) & "")
'
''        If rsSO033("OrderNo") & "" <> "" Then
'            If Val(rsSO033("SeqNo") & "") = 0 Then
'                If Not GetRS(rsAlter, "Select RowId,SO003.* From " & frmUnion.TableOwnerName & "So003 Where 1 = 0 ", gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
'            Else
'                If Not GetRS(rsAlter, "Select RowId,SO003.* From " & frmUnion.TableOwnerName & "So003 Where CustId = " & rsSO033("CustId") & " And CitemCode = " & rsSO033("CitemCode") & " And SeqNo = " & rsSO033("SeqNo"), gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
'            End If
''        End If
'        With rsAlter
'            If .RecordCount = 0 Then
'                '檢查是不要新增週期性收費項目
'                If blnAddSo003 Then
'                    If blnChkMustAnswer Then
'                        If MsgBox("無週期性收費項目,是否要新增??", vbQuestion + vbYesNo + vbDefaultButton2, "提示") = vbNo Then
'                            AlterSO003 = True
'                            Exit Function
'                        End If
'                    End If
'                    .AddNew
'                    If Val(rsSO033("SeqNo") & "") = 0 Then
'                        .Fields("SeqNO") = GetInvoiceNo3("SO003") & ""
'                    Else
'                        .Fields("SeqNO") = rsSO033("SeqNo")
'                    End If
'                    blnFlag = True
'                    '首次才自動幫客戶更新到A項 92/03/25 與 Lawrence 討論之結果 92/08/12 又改
'                    If YPara24 = 1 Then
'                        If rsSO033("AccountNO") & "" <> "" Then
'                            If Not GetRS(rsTmp, "Select Nvl(AddCitemAccount,0) AddCitemAccount,SnactionDate From " & frmUnion.TableOwnerName & "SO106 Where CustId = " & rsSO033("CustId") & _
'                                " And AccountID = '" & rsSO033("AccountNo") & "' And Nvl(StopFlag,0)= 0") Then Exit Function
'                            If rsTmp.RecordCount > 0 Then If rsTmp("AddCitemAccount") = 0 Or rsTmp("SnactionDate") & "" = "" Then blnFlag = False
'                        End If
'                    End If
'                    If blnFlag Then
'                        .Fields("CMCode") = NoZero(rsSO033("CMCode"))
'                        .Fields("CMName") = NoZero(rsSO033("CMName"))
'                        .Fields("BankCode") = NoZero(rsSO033("BankCode"))
'                        .Fields("BankName") = NoZero(rsSO033("BankName"))
'                        .Fields("AccountNo") = NoZero(rsSO033("AccountNo"))
'                    Else
'                        If Not GetRS(rsTmp, "Select CodeNo,Description From " & frmUnion.TableOwnerName & "CD031 Where RefNo=1", gcnGi) Then Exit Function
'                        .Fields("CMCode") = NoZero(rsTmp("CodeNo"))
'                        .Fields("CMName") = NoZero(rsTmp("Description"))
'                    End If
'
'                    .Fields("OrderNo") = NoZero(rsSO033("OrderNo"))
'                    .Fields("PTCode") = NoZero(rsSO033("PTCode"))
'                    .Fields("PTName") = NoZero(rsSO033("PTName"))
'                Else
'                    AlterSO003 = True
'                    Exit Function
'                End If
'            Else
'                blnAddSo003 = False
'            End If
'
'            If .Fields("StopDate") >= rsSO033("RealStopDate") Then AlterSO003 = True: Exit Function
'            If rsSO033("DiscountCode") & "" <> "" Then
'                '93/04/20 Jacky 優待辦法調整
'                If Format(rsSO033("RealStartDate") & "", "yyyymmdd") <= Format(.Fields("DiscountDate2") & "", "yyyymmdd") Or blnAddSo003 Then
'                    blnDiscount = True
'                    '.Fields("Amount") = NoZero(rsSO033("OldAmt"), True)
'                    .Fields("DiscountCode") = NoZero(rsSO033("DiscountCode"))
'                    .Fields("DiscountName") = NoZero(rsSO033("DiscountName"))
'                    .Fields("DiscountDate1") = NoZero(rsSO033("DiscountDate1"))
'                    .Fields("DiscountDate2") = NoZero(rsSO033("DiscountDate2"))
'                    .Fields("DiscountAmt") = NoZero(rsSO033("DiscountAmt"))
'                    If blnAddSo003 Then
'                        .Fields("Period").Value = GetFieldValue(rsSO033, "OldPeriod")
'                        .Fields("Amount").Value = GetFieldValue(rsSO033, "OldAmt")
'                    End If
'                End If
'            End If
'            If blnAddSo003 Or (blnUpdAmt And Ypara8 >= 1) Then
'                '若<Para9>=1, 則 1: Amount = OldAmt, 2: Amount = ShouldAmt , 3: Amount = RealAmt
'                If Not blnDiscount Then
'                    .Fields("Period").Value = GetFieldValue(rsSO033, "RealPeriod")
'                    If Ypara9 = 1 Then
'                        .Fields("Period").Value = GetFieldValue(rsSO033, "OldPeriod")
'                        .Fields("Amount").Value = GetFieldValue(rsSO033, "OldAmt")
'                    ElseIf Ypara9 = 2 Then
'                        '.Fields("Period").Value = GetFieldValue(rsSo033, "RealPeriod")
'                        .Fields("Amount").Value = GetFieldValue(rsSO033, "ShouldAmt")
'                    Else
'                        '.Fields("Period").Value = GetFieldValue(rsSo033, "RealPeriod")
'                        .Fields("Amount").Value = GetFieldValue(rsSO033, "RealAmt")
'                    End If
'                End If
'            End If
'            If rsSO033("BudgetPeriod") > 0 Then
'                .Fields("BudgetPeriod") = rsSO033("BudgetPeriod") - rsSO033("RealPeriod")
'                If .Fields("BudgetPeriod") <= 0 Then
'                     If Not ExecuteCommand("Delete From " & frmUnion.TableOwnerName & "SO003 Where RowId= '" & .Fields("RowId") & "'") Then Exit Function
'                    AlterSO003 = True
'                    Exit Function
'                End If
'            End If
'            .Fields("CustId").Value = GetFieldValue(rsSO033, "CustID")
'            .Fields("CitemCode").Value = GetFieldValue(rsSO033, "CitemCode")
'            .Fields("CitemName").Value = GetFieldValue(rsSO033, "CitemName")
'            .Fields("StartDate").Value = GetFieldValue(rsSO033, "RealStartDate")
'            .Fields("StopDate").Value = GetFieldValue(rsSO033, "RealStopDate")
'            .Fields("UpdTime").Value = GetFieldValue(rsSO033, "UpdTime")
'            .Fields("UpdEn").Value = GetFieldValue(rsSO033, "UpdEn")
'            .Fields("ClctDate").Value = GetFieldValue(rsSO033, "RealStopDate") + Ypara5
'            .Fields("LastDate").Value = GetFieldValue(rsSO033, "RealStopDate") + Ypara5
'            .Fields("CompCode").Value = lngCompCode
'            .Fields("ServiceType").Value = GetFieldValue(rsSO033, "ServiceType")
'            .Fields("PackageNo").Value = NoZero(rsSO033("PackageNo"))
'            .Fields("PackageName").Value = NoZero(rsSO033("PackageName"))
'            .Update
'            strSeqNo = .Fields("SeqNo") & ""
'        End With
'        AlterSO003 = True
'    Exit Function
'ChkErr:
'    ErrSub "UpdateFun", "AlterSo003"
'End Function

Public Function GetSystemPara(ByRef rs As ADODB.Recordset, ByVal strCompCode As String, _
    ByVal strServiceType As String, ByVal strTable As String, Optional ByVal strParaField As String = "", _
    Optional blnShowMsg As Boolean = True, Optional blnServiceTypeUseLike As Boolean = False) As Boolean
    On Error GoTo ChkErr
    Dim strSQL As String
    Dim blnFlag As Boolean
        If strParaField = "" Then strParaField = "*"
        'If rs Is Nothing Then Set rs = New ADODB.Recordset
        If blnServiceTypeUseLike Then
            strSQL = " Like '%" & strServiceType & "%'"
        Else
            strSQL = " = '" & strServiceType & "'"
        End If
        '先找CompCode 及ServiceType 都有的
        strTable = frmUnion.TableOwnerName & strTable
        
        If strCompCode <> "" And strServiceType <> "" Then
            If Not GetRS(rs, "Select " & strParaField & " From " & strTable & " Where CompCode = " & strCompCode & " And ServiceType " & strSQL) Then Exit Function
            blnFlag = Not rs.EOF
        End If
        If Not blnFlag Then
            '再找CompCode 有的
            If strCompCode <> "" Then
                If Not GetRS(rs, "Select " & strParaField & " From " & strTable & " Where CompCode = " & strCompCode) Then Exit Function
                blnFlag = Not rs.EOF
            End If
            '最後找只要有這個 參數檔的
            If Not blnFlag Then
                If Not GetRS(rs, "Select " & strParaField & " From " & strTable) Then Exit Function
                blnFlag = Not rs.EOF
            End If
        End If
        If blnShowMsg And Not blnFlag Then MsgBox "在系統參數檔[" & strTable & "] 找不到 公司別代號為 [" & strCompCode & "] ,且服務類別為: [" & strServiceType & "]的資料,請查證!!", vbCritical, "訊息"
        GetSystemPara = True
    Exit Function
ChkErr:
    ErrSub "UpdateFun", "GetSystemPara"
End Function

Public Function GetFieldValue(rs As ADODB.Recordset, _
        strFieldName As String, Optional intType As Integer = 0)
    On Error GoTo ChkErr
    Dim varField As Variant
        'intType 0:Value  1:OriginalValue
        If intType = 0 Then
            varField = rs.Fields(strFieldName).Value
        ElseIf intType = 1 Then
            varField = rs.Fields(strFieldName).OriginalValue
        End If
        
        If rs.Fields(strFieldName).Type = 131 Or rs.Fields(strFieldName).Type = 139 Then
            If intType = 0 Then
                If Len(rs.Fields(strFieldName).Value & "") > 0 Then
                    varField = Val(rs.Fields(strFieldName).Value)
                End If
            ElseIf intType = 1 Then
                If Len(rs.Fields(strFieldName).OriginalValue & "") > 0 Then
                    varField = Val(rs.Fields(strFieldName).OriginalValue)
                End If
            End If
        End If
        GetFieldValue = varField
    Exit Function
ChkErr:
    ErrSub "UpdateFun", "GetFiledValue"
End Function

Public Function NoZero(strVal As Variant, Optional ZeroFlag As Boolean = False) As Variant
  On Error GoTo ChkErr
    strVal = strVal & ""
    If Trim(strVal) = "" Then
        If ZeroFlag = True Then
            NoZero = 0
        Else
            NoZero = Null
        End If
    Else
        NoZero = strVal
    End If
  Exit Function
ChkErr:
    ErrSub "Sys_Lib", "NoZero"
End Function


Public Sub CloseRecordset(rs As ADODB.Recordset)
  On Error Resume Next
    rs.CancelUpdate
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
End Sub
Public Function GetInvoiceNo3(strTableName As String, Optional Conn As ADODB.Connection) As String
On Error GoTo ChkErr
Dim strSeq As String
Dim strInv
strSeq = strOwnerName & "S_" & strTableName
If Conn Is Nothing Then Set Conn = gcnGi
strInv = GetRsValue("SELECT Ltrim(To_Char(" & strSeq & ".NextVal, '09999999')) FROM Dual", Conn)
GetInvoiceNo3 = strInv
Exit Function
ChkErr:
ErrSub "Sys_Lib", "GetInvoiceNo3"
End Function

Public Function ExecuteCommand(strSQL As String, Optional Conn As ADODB.Connection, Optional ByRef lngAffected As Long) As Boolean
    On Error GoTo ChkErr
        If Conn Is Nothing Then Set Conn = gcnGi
        Conn.Execute strSQL, lngAffected
        ExecuteCommand = True
    Exit Function
ChkErr:
    ErrSub FormName, "ExecuteCommand, SQL : [" & strSQL & "] "
End Function


Public Function RightNow() As Date
  On Error Resume Next
    RightNow = gcnGi.Execute("SELECT TO_CHAR(SYSDATE,'YYYY/MM/DD HH24:MI:SS') FROM DUAL").GetString & ""
    If Err.Number <> 0 Then RightNow = Now
End Function

Public Function RightDate() As Date
  On Error Resume Next
    RightDate = Format(RightNow, "YYYY/MM/DD")
    If Err.Number <> 0 Then RightDate = Date
End Function

