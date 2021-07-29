VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "csACHTranReferOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const FormName = "clsACHTranRefer"

'Dim clsStoreParameter As clsStoreParameter

Private rsCD018 As New ADODB.Recordset
Private rsCD018Total As New ADODB.Recordset
Private DataFile(30) As TextStream
Private ErrFile As TextStream
Private MemoFile As TextStream

Private lngBankCount(30) As Long
Private lngBankAmt(30) As Long

Private lngCountBillNo As Long
Private lngErrCount As Long

Private lngPercent As Long
Private strInvoiceId As String
Private strReturnSQL As String


Private strChooseSQL As Variant       'INI檔案路徑
Private strErrPath As Variant       'INI檔案路徑
Private strRealDate As Variant       '入帳日期
Private strUpdEn As Variant          '記錄異動人員
Private strUpdName As Variant          '記錄異動人員
Private strPTCode As Variant         '付款種類
Private strServiceType As Variant    '服務類別
Private strBankSQL As Variant
Private strCompCode As Variant       '公司別
Private strBillNo As String                 'BillNo
Private blnBeOne As Boolean
Private strBankHand  As String
Private strChoose33 As String
Private lngPara24 As Long
Private strFlowId As Integer

Public Function Action(Optional ByRef lngSecond As Long) As Boolean
    On Error GoTo ChkErr
    Dim lngTime As Long
        lngTime = Timer
        If Not InitData Then Exit Function
        If Not OpenData Then Exit Function
        If Not BeginTran Then Exit Function
        
        lngSecond = Round(Timer - lngTime)
        MsgBox "共花費 : " & IIf(lngSecond > 60, lngSecond \ 60 & " 分 ", "") & (lngSecond Mod 60) & " 秒 !!", vbInformation, gimsgPrompt
        Action = True
    Exit Function
ChkErr:
    ErrSub FormName, "Action"
End Function

Private Function InitData() As Boolean
    On Error GoTo ChkErr
        With clsStoreParameter
            blnBeOne = .uBeOne
            strBankHand = .uBankHand
            strBankSQL = .uBankSQL
            strChooseSQL = .uChooseSQL
            strCompCode = .uCompCode
            strErrPath = .uErrPath & IIf(Right(.uErrPath, 1) = "\", "", "\")
            strPTCode = .uPTCode
            strRealDate = .uRealDate
            strServiceType = .uServiceType
            strUpdEn = .uUpdEn
            strUpdName = .uUpdName
            strChoose33 = .uChoose33
            strFlowId = .FlowId
        End With
        lngPara24 = GetSystemParaItem("Para24", CStr(strCompCode), CStr(strServiceType), "SO043")
        InitData = True
    Exit Function
ChkErr:
    Call ErrSub(FormName, "InitData")
End Function

Private Function GetRsTmp(ByRef rs As ADODB.Recordset) As Boolean
    On Error GoTo ChkErr
    Dim strFrom  As String, strWhere As String
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset

       'EMC
        If Not GetRS(rsTmp, "Select A.BankCode as BankCode,A.BankName as BankName,A.RealDate " & _
          " as TranDate,A.AccountNo,A.ShouldAmt as ShouldAmt,A.BillNo as BillNo_Old,A.BillNo as BillNo_New From " & strOwnerName & "So033 A Where RowId=''") Then Exit Function
        
        If Not CreateMDBTable(rsTmp, GetMDBTableName, cnn) Then Exit Function
        
        If lngPara24 = 0 Then
'            strFrom = " " & strOwnerName & "SO033 A, " & strOwnerName & "SO002A D, " & strOwnerName & "So001 B "
'            strWhere = " And A.CustId = D.CustId And A.ServiceType = D.ServiceType And A.CustId=B.CustId "
            strFrom = " " & strOwnerName & "SO033 A "
            
            If strChooseSQL <> "" Then
                If InStr(1, strChooseSQL, "C.") Then
                    strFrom = strFrom & "," & strOwnerName & "So014 C "
                    strWhere = strWhere & IIf(strWhere = "", "", " And ") & " A.AddrNo=C.AddrNo "
                End If
                If InStr(1, strChooseSQL, "E.") Then
                    strFrom = strFrom & "," & strOwnerName & "So002 E "
                    strWhere = strWhere & IIf(strWhere = "", "", " And ") & " A.ServiceType = E.ServiceType And A.CustId=E.CustId And A.CompCode=E.CompCode "
                End If
            End If
            strWhere = strWhere & IIf(strWhere = "", "", " And ") & strChooseSQL & " And A.CancelFlag=0 And A.UCCode is Not Null And A.ShouldAmt > 0"
'GetUseIndexStr
            strSQL = "SELECT " & GetUseIndexStr("So033", "ShouldDate") & " DISTINCT A.BillNo,A.CustId,A.AccountNo,A.BankCode,A.AddrNo,A.MediaBillNo, Sum(ShouldAmt) ShouldAmt From " & _
                        strFrom & " Where " & _
                        strWhere & " Group By A.BillNo,A.CustId,A.AccountNo,A.BankCode,A.AddrNo,A.MediaBillNo "
         Else
            strFrom = " " & strOwnerName & "SO033 A "
                If strChooseSQL <> "" Then
                If InStr(1, strChooseSQL, "C.") Then
                    strFrom = strFrom & "," & strOwnerName & "So014 C "
                    strWhere = strWhere & IIf(strWhere = "", "", " And ") & " A.AddrNo=C.AddrNo "
                End If
                If InStr(1, strChooseSQL, "E.") Then
                    strFrom = strFrom & "," & strOwnerName & "So002 E "
                    strWhere = strWhere & IIf(strWhere = "", "", " And ") & " A.ServiceType = E.ServiceType And A.CustId=E.CustId And A.CompCode=E.CompCode "
                End If
            End If
            strWhere = strWhere & IIf(strWhere = "", "", " And ") & strChooseSQL & " And A.CancelFlag=0 And A.UCCode is Not Null And A.ShouldAmt > 0"
'GetUseIndexStr
            strSQL = "SELECT " & GetUseIndexStr("So033", "ShouldDate") & " DISTINCT A.MediaBillNo BillNo,A.AccountNo,A.BankCode,A.AddrNo, Sum(ShouldAmt) ShouldAmt  From " & _
                        strFrom & " Where " & _
                        strWhere & " Group By A.MediaBillNo,A.AccountNo,A.BankCode,A.AddrNo "
         
         End If
        If Not GetRS(rs, strSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic, , , True) Then Exit Function
        strReturnSQL = rs.Source
    GetRsTmp = True
    Exit Function
ChkErr:
    ErrSub FormName, "GetRsTmp"
End Function

Private Function GetRsUpd(ByRef rs As ADODB.Recordset) As Boolean
    On Error GoTo ChkErr
    Dim strFrom  As String, strWhere As String
    Dim strSQL As String
    Dim rsUpd As New ADODB.Recordset
            strFrom = " " & strOwnerName & "SO033 A "
            If strChooseSQL <> "" Then
                If InStr(1, strChooseSQL, "C.") Then
                    strFrom = strFrom & "," & strOwnerName & "So014 C "
                    strWhere = " A.AddrNo=C.AddrNo "
                End If
                strWhere = strWhere & " And " & strChoose33 & " And A.MediaBillNo Is Null "
            End If
            
            strWhere = strWhere & " And A.CancelFlag=0 And A.UCCode is Not Null And A.ShouldAmt > 0"
            strWhere = Mid(strWhere, 5)
'GetUseIndexStr
            strSQL = "SELECT " & GetUseIndexStr("So033", "ShouldDate") & " A.CustId, A.BillNo, A.MediaBillNo,A.AccountNo,A.BankCode,A.AddrNo From " & _
                        strFrom & " Where " & strWhere
         
       
        If Not GetRS(rs, strSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic, , , True) Then Exit Function
        strReturnSQL = rs.Source
    GetRsUpd = True
    Exit Function
ChkErr:
    ErrSub FormName, "GetRsUpd"
End Function

Private Function InsertHead(rs As ADODB.Recordset) As Boolean
    On Error GoTo ChkErr
    Dim intReserve As Integer
        '區別碼1 , 發件單位8, 收件單位8, 轉帳類別3, 入 / 扣帳日6, 性質別1, 保留欄93
        rsCD018.MoveFirst
        If blnBeOne Then rsCD018.Filter = "CodeNo = " & strBankHand
        Do While Not rsCD018.EOF
            intReserve = 93
            If Val(rsCD018("TextLength") & "") > 0 Then intReserve = 93 - (120 - Val(rsCD018("TextLength") & ""))
            WriteTextLine DataFile(rsCD018.AbsolutePosition), "1" & GetString(rsCD018("BankId"), 8, giLeft) & _
                    GetString(rsCD018("CorpId") & "", 8, giLeft) & _
                    GetString(rsCD018("Class") & "", 3) & _
                    GetString(strRealDate, 6) & _
                    GetString(1, 1) & _
                    GetString("", intReserve)
            rsCD018.MoveNext
        Loop
        rsCD018.Filter = ""
        InsertHead = True
    Exit Function
ChkErr:
    ErrSub FormName, "InsertHead"
End Function

Private Function InsertDetail(rs As ADODB.Recordset) As Boolean
    On Error GoTo ChkErr
    Dim lngAmount As Long, lngOldCustId As Long
    Dim strWriteLine As String, strLimitation As String, strWhere As String
    Dim rsTmp As New ADODB.Recordset
    Dim rsSo106 As New ADODB.Recordset
    Dim rsSo002a As New ADODB.Recordset
    Dim rsInv As New ADODB.Recordset
    Dim strAccNameID As String
    Dim strSQL As String, strInvNo As String
        If rs.RecordCount > 0 Then rs.MoveFirst
        cnn.BeginTrans
        While Not rs.EOF
            If Not FindCD018(rs) Then
                lngAmount = 0
                lngErrCount = lngErrCount + 1: GoTo lNext
            End If
            'If Not ChkErrCustId(rs, lngErrCount) Then
            '    lngAmount = 0
            '    GoTo lNext
            'End If
            '寫到文字檔
            If lngPara24 = 1 Then
                strWhere = " Where MediaBillNo='" & rs("BillNo") & "' And Nvl(CancelFlag,0)=0 And UCCode is Not Null " & IIf(Len(strChoose33) > 0, "And ", "") & strChoose33
            Else
                strWhere = " Where BillNo='" & rs("BillNo") & "' And Nvl(CancelFlag,0)=0 And UCCode is Not Null " & IIf(Len(strChoose33) > 0, "And ", "") & strChoose33
            End If
            
            If rsCD018("RefNo") & "" = 4 Then
                '吉元聯合中心的才取基本台訊號有效期限
                If Not GetRS(rsTmp, "Select RealStartDate,RealStopDate From " & strOwnerName & "So033 A " & strWhere & " And CitemCode in (Select CodeNo From " & strOwnerName & "CD019 Where RefNo = 1)", gcnGi, , , , , 1) Then Exit Function
                If Not rsTmp.EOF Then
                    strLimitation = Replace(GetDT(rsTmp("RealStartDate"), GiDate) & "-" & GetDT(rsTmp("RealStopDate"), GiDate), "/", "")
                Else
                    strLimitation = ""
                End If
            End If

'            lngAmount = Val(GetRsValue("Select Sum(ShouldAmt) From " & strOwnerName & "So033 A " & strWhere) & "")
            lngAmount = rs("ShouldAmt") & ""
            If rs("AccountNo") & "" <> "" Then
                strSQL = "Select CUSTID,AccountNameID From " & strOwnerName & "So106 Where AccountID='" & rs("AccountNo") & "" & "'" _
                          & " And StopDate Is Null And SnactionDate Is Not Null"
                If Not GetRS(rsSo106, strSQL, gcnGi, , , , , 1) Then Exit Function
                If rsSo106.EOF Then
                    strAccNameID = ""
                Else
                    strAccNameID = rsSo106("AccountNameId") & ""
                End If
            Else
                strAccNameID = ""
            End If
'            strSQL = "Select CUSTID,Id,InvNo From " & strOwnerName & "So002 Where AccountNo='" & rs("AccountNo") & "'" & " And Rownum=1"
'            If Not GetRS(rsInv, strSQL, gcnGi) Then Exit Function
'            If rsSo106.EOF = False Then
'                If rsInv("ID") & "" = "" Then
'                strsql = "Select Id From " & strOwnerName & "So001 Where Custid = (SELECT A.Custid From " & strOwnerName & "So002a A , " & _
'                         strOwnerName & "So033 B Where A.AccountNo=B.AccountNo And A.CustId=B.CustId And B.AccountNo='" & rs("AccountNo") & "" & "' And Rownum=1)"
'                 If Not GetRS(rsSo002a, strsql, gcnGi) Then Exit Function
'                    strAccNameID = rsSo106("AccountNameId") & ""
'                Else
'                    strAccNameID = rsInv("ID") & ""
'                End If
'            Else
'                strAccNameID = rsInv("ID") & ""
'            End If
'            20031006 Jim 提不用用到InvNo 此欄位,所以把它拿掉
'            If rs("InvNo") & "" = "" Then
'                 strInvNo = GetRsValue("Select InvNo From " & strOwnerName & "So002 Where AccountNo ='" & rs("AccountNo") & "" & "'") & ""
'            Else
'                strInvNo = rs("InvNo") & ""
'            End If
            If Not WriteDetailToText(rsCD018("BankId") & "", rsCD018("CorpId") & "", rsCD018("Class") & "", rs("AccountNo") & "" _
                    , lngAmount, strInvoiceId, rs("BillNo"), strAccNameID & "", rsCD018("RefNo") & "", rs("BankCode") & "", rsCD018("PrgName") & "" _
                    , rs("AddrNo") & "", strLimitation, Val(rsCD018("TextLength") & ""), strAccNameID, rs("BillNo") & "") Then Exit Function
            
            If blnBeOne Then
            '存總計筆數及金額
                lngBankCount(1) = lngBankCount(1) + 1
                lngBankAmt(1) = lngBankAmt(1) + lngAmount
            Else
                '存各銀行小計及金額
                lngBankCount(rsCD018.AbsolutePosition) = lngBankCount(rsCD018.AbsolutePosition) + 1
                lngBankAmt(rsCD018.AbsolutePosition) = lngBankAmt(rsCD018.AbsolutePosition) + lngAmount
            End If
lNext:
'            lngOldCustId = rs("BillNo")
            rs.MoveNext
            DoEvents
        Wend
        cnn.CommitTrans
        InsertDetail = True
    Exit Function
ChkErr:
    cnn.RollbackTrans
    ErrSub FormName, "InsertDetail"
End Function

Private Function GetOldValue(rs As ADODB.Recordset, ByRef strOldBankId As String, _
                ByRef strOldCorpId As String, ByRef strOldClass As String, _
                ByRef strOldAccountNo As String, ByRef lngAmount As Long, _
                ByRef strInvoiceId As String, ByRef strOldBillNo As String, _
                ByRef lngOldCustId As Long, ByRef strOldID As String, _
                ByRef strBankRefNo As String, ByRef strBankCode As String, _
                ByRef strPrgName As String) As Boolean
    On Error GoTo ChkErr
        strOldBillNo = rs("BillNo") & ""
        lngOldCustId = rs("CustId")
        strBankCode = rs("BankCode") & ""
        strOldBankId = rsCD018("BankId") & ""
        strOldCorpId = rsCD018("CorpId") & ""
        strOldClass = rsCD018("Class") & ""
        strBankRefNo = rsCD018("RefNo") & ""
        strPrgName = rsCD018("PrgName") & ""
        strOldAccountNo = rs("AccountNo") & ""
        strOldID = rs("ID") & ""
        GetOldValue = True
    Exit Function
ChkErr:
    ErrSub FormName, "GetOldValue"
End Function

Private Function WriteDetailToText(strOldBankId As String, _
                 strOldCorpId As String, strOldClass As String, _
                 strOldAccountNo As String, lngAmount As Long, _
                 strInvoiceId As String, strOldBillNo As String, _
                 strID As String, strBankRefNo As String, _
                 strOldBankCode As String, strPrgName As String, _
                 strAddrNo As String, strLimitation As String, _
                 intTextLength As Integer, strInvNo As String, _
                 strMediaBillNo As String) As Boolean
    On Error GoTo ChkErr
    Dim strWriteLine As String
    Dim strSpecialBillNo As String, strBillNoOut As String
    Dim lngBR As Long, intReserve As Integer
        '區別碼1,發件單位8,收件單位8,轉帳類別3,入/扣帳日6
        '帳號14,交易金額14(2),營利事業統一編號8,狀況代號2(99)
        
        
        '花旗銀行的   專用資料區44 銀行 (特殊單號+檢查碼+單號)
        '  or  專用資料區44 郵局 (身份證字號ID+單號)
        '其他家銀行 單號 BillNO(15)
        
        strBankRefNo = Val(strBankRefNo)
        
        '保留欄12
        intReserve = 12
        If intTextLength > 0 Then intReserve = 12 - (120 - intTextLength)
        strWriteLine = "2" & _
            GetString(strOldBankId & "", 8, giLeft) & _
            GetString(strOldCorpId & "", 8, giLeft) & _
            GetString(strOldClass & "", 3) & _
            GetString(strRealDate, 6) & _
            GetString(strOldAccountNo & "", 14, giRight, True) & _
            GetString(lngAmount, 12, giRight, True) & "00" & _
            GetString(strInvoiceId, 8) & "99" & _
            GetUserAreaStr(strOldBillNo, strPrgName, strBankRefNo, strID, strAddrNo, strLimitation, strBillNoOut, strInvNo, strMediaBillNo)
            '如果為郵局格式則在109的位址上加註 P
            If strBankRefNo = 6 Then
                strWriteLine = strWriteLine & GetString("P", 109)
            End If
            strWriteLine = strWriteLine & GetString("", intReserve)
        lngBR = rsCD018.AbsolutePosition
        Call FindCD018(, , strOldBankCode)
        If blnBeOne Then
            If Not WriteTextLine(DataFile(1), strWriteLine) Then Exit Function
        Else
            If Not WriteTextLine(DataFile(rsCD018.AbsolutePosition), strWriteLine) Then Exit Function
        End If
        '填入暫存檔
        If Not ExecuteCommand("Insert Into " & GetMDBTableName & " (BankCode,BankName,TranDate,AccountNo,ShouldAmt,BillNo_Old,BillNo_New) Values (" & _
                GetNullString(strOldBankCode, giLongV) & "," & GetNullString(GetBankDescription(strOldBankCode) & "") & ",'" & GetRealDateTran(CStr(strRealDate)) & "'," & _
                GetNullString(strOldAccountNo) & "," & lngAmount & ",'" & strBillNoOut & "','" & strOldBillNo & "')", cnn) Then Exit Function
        If rsCD018.AbsolutePosition <> lngBR Then
            'MsgBox "x"
        End If
        If lngBR > 0 Then rsCD018.AbsolutePosition = lngBR
        WriteDetailToText = True
    Exit Function
ChkErr:
    ErrSub FormName, "WriteDetailToText"
End Function

Private Function GetBankDescription(strBankCode As String) As String
    On Error GoTo ChkErr
        rsCD018.Filter = "CodeNo = " & strBankCode
        GetBankDescription = rsCD018("Description") & ""
        rsCD018.Filter = ""
    Exit Function
ChkErr:
    ErrSub FormName, "GetBankDescription"
End Function

Private Function FindCD018(Optional ByRef rs As ADODB.Recordset, _
        Optional lngOldCustId As Long, Optional strBankCode As String = "") As Boolean
    On Error GoTo ChkErr
    Dim lngCustId As Long
        'If lngOldCustId <> rs("CustId") And rsCD018.RecordCount > 0 Then
            lngCustId = lngOldCustId
            If strBankCode = "" Then strBankCode = rs("BankCode")
            If strBankCode = "" Then Exit Function
            
            If strBankCode = "" Then strBankCode = rs("BankCode")
            If strBankCode = "" Then Exit Function
            rsCD018.MoveFirst
            If blnBeOne Then
                rsCD018.Find "CodeNo = " & strBankHand
            Else
                rsCD018.Find " CodeNo = " & strBankCode
            End If
            If rsCD018.EOF Then
                Exit Function
            End If
        'End If
        FindCD018 = True
    Exit Function
ChkErr:
    ErrSub FormName, "FindCD018"
End Function

'尾筆
Private Function InsertFinal(rs As ADODB.Recordset) As Boolean
    On Error GoTo ChkErr
    Dim intReserve As Integer
        rsCD018.MoveFirst
        If blnBeOne Then rsCD018.Filter = "CodeNo = " & strBankHand
        '區別碼1,發件
        Do While Not rsCD018.EOF
            intReserve = 42
            If Val(rsCD018("TextLength") & "") > 0 Then intReserve = 42 - (120 - Val(rsCD018("TextLength") & ""))
            
            WriteTextLine DataFile(rsCD018.AbsolutePosition), "3" & GetString(rsCD018("BankId") & "", 8, giLeft) & _
                    GetString(rsCD018("CorpId") & "", 8, giLeft) & _
                    GetString(rsCD018("Class") & "", 3) & _
                    GetString(strRealDate, 6) & _
                    GetString(lngBankAmt(rsCD018.AbsolutePosition), 14, giRight, True) & _
                    "00" & _
                    GetString(lngBankCount(rsCD018.AbsolutePosition), 10, giRight, True) & _
                    String(26, "0") & _
                    GetString("", intReserve)
            rsCD018.MoveNext
        Loop
        rsCD018.Filter = ""
        InsertFinal = True
    Exit Function
ChkErr:
    ErrSub FormName, "InsertFinal"
End Function

Private Function BeginTran() As Boolean
    On Error GoTo ChkErr
    Dim rsTmp As New ADODB.Recordset
    Dim strBankId As String, strSQL As String
    Dim strOldBillNo As String
    Dim lngCountCustId As Long, lngErrCount As Long
    Dim lngPara24 As Long
    Dim rsUpd As New ADODB.Recordset
        '取資料
        If Not GetRsTmp(rsTmp) Then Exit Function
        
'        Select Case strFlowId
'        Case 0
            '檢查是否有用多帳號功能
            lngPara24 = GetSystemParaItem("Para24", CStr(strCompCode), CStr(strServiceType), "SO043")
            If lngPara24 = 1 Then
                If Not GetRsUpd(rsUpd) Then Exit Function
                If Not BatchUpdateMediaBillNo(rsUpd, strChoose33, gcnGi) Then Exit Function
            End If
'        End Select
        If rsTmp.RecordCount = 0 Then
            MsgBox "無資料！請重新操作！", vbInformation, "訊息！"
            GoTo KillFile
        Else
            If Not InsertHead(rsTmp) Then GoTo KillFile
            If Not InsertDetail(rsTmp) Then GoTo KillFile
            If Not InsertFinal(rsTmp) Then GoTo KillFile
        End If
        Call CloseFS(False)
        BeginTran = True
        On Error Resume Next
        rsTmp.Close
        Set rsTmp = Nothing
    Exit Function
KillFile:
    '刪除檔案..
    Call CloseFS(True)
    Exit Function
ChkErr:
    Call ErrSub(FormName, "BeginTran")
End Function

Private Function ChkErrCustId(rs As ADODB.Recordset, _
        ByRef lngErrCount As Long) As Boolean
    On Error GoTo ChkErr
    'Dim rsTmp As New ADODB.Recordset
        'If Not GetRS(rsTmp, "Select CustName,MailAddress From So001 Where CustId = " & rs("CustId")) Then Exit Function
        ChkErrCustId = True
        Exit Function
        If InStr(1, rs("MailAddress") & "", "號") = 0 Then
           WriteTextLine ErrFile, rs("CustId") & rs("CustName") & rs("BillNo") & rs("MailAddress")
           lngErrCount = lngErrCount + 1
           Exit Function
           '問題檔
        End If
        ChkErrCustId = True
    Exit Function
ChkErr:
    ErrSub FormName, "ChkErrCustId"
End Function

Private Function OpenData() As Boolean
    On Error GoTo ChkErr
    Dim lngCount As Long
    Dim strFileName As String
        If Not GetRS(rsCD018, "Select * From " & strOwnerName & "CD018 Where CodeNo  " & strBankSQL & " order by CodeNo", gcnGi) Then Exit Function
        
        If rsCD018.RecordCount = 0 Then Exit Function
'        Select Case strFlowId
'        Case 0  'EMC
            strInvoiceId = GetSystemParaItem("InvoiceId", strCompCode & "", strServiceType & "", "SO041")
'        Case 1  'TBC
'            strInvoiceId = GetRsValue("Select InvoiceId From " & strOwnerName & "So041 Where ServiceType ='" & strServiceType & "'") & ""
            If strInvoiceId = "" Then MsgBox "服務類別代碼為 :[ " & strServiceType & " ] 且公司別代碼為 : [" & strCompCode & "] 無對應之InvoiceID !! 請查證.. ": Exit Function
'        End Select
        rsCD018.MoveFirst
        If blnBeOne Then rsCD018.Filter = "CodeNo = " & strBankHand
        Do While Not rsCD018.EOF
            lngCount = rsCD018.AbsolutePosition
            strFileName = Left(rsCD018("FileName1") & "", InStrRev(rsCD018("FileName1") & "", "."))
            If strFileName = "" Then strFileName = rsCD018("FileName1") & ""
            If strFileName = "" Then MsgBox "請設定銀行 : [" & rsCD018("CodeNO") & rsCD018("Description") & " ] 之轉帳輸出檔名 ": Exit Function
            
            If Not OpenFile(DataFile(lngCount), rsCD018("FileName1") & "", True) Then Exit Function
            rsCD018.MoveNext
        Loop
        rsCD018.Filter = ""
        
        If Not OpenFile(ErrFile, strErrPath & "ACHTranReferErr.Txt", True) Then Exit Function
        If Not OpenFile(MemoFile, strErrPath & "ACHTranReferMem.Txt", True) Then Exit Function
        OpenData = True
    Exit Function
ChkErr:
    ErrSub FormName, "OpenData"
End Function

Private Function CloseFS(Optional KillFile As Boolean = False) As Boolean
    On Error Resume Next
    Dim lngLoop As Long
    Dim strFileName As String
        ErrFile.Close
        MemoFile.Close
        rsCD018.MoveFirst
        If blnBeOne Then rsCD018.Filter = "CodeNo = " & strBankHand
        Do While Not rsCD018.EOF
            DataFile(rsCD018.AbsolutePosition).Close
            If KillFile Then
                strFileName = Left(rsCD018("FileName1"), InStrRev(rsCD018("FileName1"), "."))
                If strFileName = "" Then strFileName = rsCD018("FileName1") & ""
                If strFileName <> "" Then
                    Kill rsCD018("FileName1") & ""
                    Kill strErrPath & "ACHTranReferErr.Txt"
                    Kill strErrPath & "ACHTranReferMem.Txt"
                End If
            End If
            rsCD018.MoveNext
        Loop
        rsCD018.Filter = ""

    Exit Function
ChkErr:
    ErrSub FormName, "CloseFS"
End Function

Public Property Get uReturnSQL() As String
    On Error Resume Next
        uReturnSQL = strReturnSQL
End Property

Private Function GetSpecialBillNo(strBillNo As String, _
        Optional strRefNo As String = "", _
        Optional strID As String, _
        Optional strAddrNo As String) As String
    On Error Resume Next
    Dim strSpecialBillNo As String
    Dim strCircuitNo As String
        '等於 網路編號有'1',否:'2' +Mid(舊單號,2,3) +六碼客編+1碼檢查碼
        If strAddrNo <> "" Then strCircuitNo = GetRsValue("Select CircuitNo From " & strOwnerName & "So014 Where AddrNo = " & strAddrNo) & ""
        strSpecialBillNo = IIf(strCircuitNo = "", 1, 2) & Mid(GetBillNo_Old(strBillNo, , True), 2, 3) & Right(strBillNo, 6)
        strSpecialBillNo = strSpecialBillNo & GetChkCode(strSpecialBillNo)
        GetSpecialBillNo = strSpecialBillNo
End Function

Private Function GetBillNo_New(strBillNo As String, Optional strMediaBillNo As String) As String
    On Error Resume Next
    Dim strType As String
    Dim strYM As String
        If lngPara24 = 1 Then
            GetBillNo_New = strMediaBillNo
        Else
            Select Case Mid(strBillNo, 7, 2)
                Case "BC"
                    strType = 1
                Case "TC"
                    strType = 2
                Case "BI"
                    strType = 3
                Case "TI"
                    strType = 4
            End Select
            strYM = Format(Format(Left(strBillNo, 6) & "01", "####/##/##"), "EEMM")
            If Val(Mid(strBillNo, 5, 2)) < 10 Then
                strYM = Left(strYM, 2) & Right(strYM, 1)
            Else
                strYM = Left(strYM, 2) & Chr(Asc("A") + Mid(strYM, 3, 2) - 10)
            End If
            '改成民國年11 碼單號
            GetBillNo_New = strYM & strType & Right(strBillNo, 7)
        
        End If
End Function

Private Function GetBillNo_Old(ByVal strBillNo As String, Optional strMediaBillNo As String, _
    Optional blnRealBillNo As Boolean = False) As String
    On Error Resume Next
        If lngPara24 = 1 And Not blnRealBillNo Then
            GetBillNo_Old = strMediaBillNo
        Else
            GetBillNo_Old = Trim(CStr(Val(Left(strBillNo, 6)) - 191100)) & _
                Mid(strBillNo, 7, 1) & Format(Right(strBillNo, 6), "000000")
        End If
End Function

Private Function GetMDBTableName() As String
    On Error Resume Next
        GetMDBTableName = UCase("csACHTranReferOut")
End Function

Private Function GetUserAreaStr(strBillNo As String, strPrgName As String, _
    strRefNo As String, strID As String, strAddrNo As String, strLimitation As String, _
    ByRef strBillNoOut As String, strInvNo As String, strMediaBillNo As String) As String
    On Error Resume Next
    '花旗銀行的   專用資料區44 銀行 (特殊單號+檢查碼+單號)
    '  or  專用資料區44 郵局 (身份證字號ID+單號)
    '其他家銀行 單號 BillNO(15)
    '6 是郵局的
    Dim strUserArea As String
        If strRefNo = 1 Or strRefNo = 6 Then
            '花旗的格式
            strUserArea = GetString(strID, 10) & GetBillNo_Old(strBillNo, strMediaBillNo)
            strBillNoOut = GetBillNo_Old(strBillNo, strMediaBillNo)
        ElseIf strRefNo = 2 Then
            '真正的金資
            strUserArea = GetSpecialBillNo(strBillNo, strRefNo, strID, strAddrNo) & GetBillNo_New(strBillNo, strMediaBillNo)
            strBillNoOut = GetBillNo_New(strBillNo, strMediaBillNo)
        ElseIf strRefNo = 3 Then
            '台灣銀行用(吉元)
            strUserArea = GetBillNo_New(strBillNo, strMediaBillNo) & Space(9) & GetString(strID, 10)
            strBillNoOut = GetBillNo_New(strBillNo, strMediaBillNo)
        ElseIf strRefNo = 4 Then
            '農會聯合中心(吉元)
            strUserArea = GetString(strLimitation, 13) & GetBillNo_Old(strBillNo, strMediaBillNo)
            strBillNoOut = GetBillNo_Old(strBillNo, strMediaBillNo)
        ElseIf strRefNo = 5 Then
            '土地銀行(吉元)
'           以身份證字號傳入---for Lydia
'            If strInvNo <> "" Then
'                strUserArea = GetString(strInvNo, 8) & Space(8) & GetBillNo_Old(strBillNo, strMediaBillNo)
'            Else
                strUserArea = GetString(strInvNo, 10) & IIf(Len(strInvNo) <= 8, Space(2), "") & Space(6) & GetBillNo_Old(strBillNo, strMediaBillNo)
'            End If
            strBillNoOut = GetBillNo_Old(strBillNo, strMediaBillNo)
        Else
            '其他家的金資格式
            strUserArea = GetBillNo_New(strBillNo, strMediaBillNo)
            strBillNoOut = GetBillNo_New(strBillNo, strMediaBillNo)
        End If
        GetUserAreaStr = strUserArea & GetString(" ", 44 - Len(strUserArea))
End Function

Private Function GetRealDateTran(strRealDate As String) As String
    On Error Resume Next
        GetRealDateTran = Format(Val(strRealDate) + 19110000, "####/##/##")
End Function

Private Function GetChkCode(strBillNo As String) As String
  On Error GoTo ChkErr
  Dim lngLoop As Long
  Dim lngMath As Long
  Dim lngValue As Long
   '取單據編號之檢查碼
   If strBillNo = "" Then Exit Function
    For lngLoop = 1 To Len(strBillNo)
        Select Case lngLoop Mod 3
            Case 0
               lngMath = 1
            Case 1
               lngMath = 3
            Case 2
               lngMath = 7
        End Select
            lngValue = lngValue + Val(Mid(strBillNo, lngLoop, 1)) * lngMath
    Next
    GetChkCode = Trim(Right(10 - Right(lngValue & "", 1), 1))
  Exit Function
ChkErr:
    Call ErrSub(FormName, "GetChkCode")
End Function

