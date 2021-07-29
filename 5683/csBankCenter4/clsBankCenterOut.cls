VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBankCenterOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const FormName = "clsBankCenter"
Private rsCD018 As New ADODB.Recordset
Private rsCD018Total As New ADODB.Recordset
Private DataFile(30) As TextStream
Private strDataFileName(30) As String
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
Private blnNT As Boolean
Private strBankHand  As String
Private strChoose33 As String
Private lngPara24 As Long
Private strFlowId As Integer
Private FSO As New FileSystemObject
Private Tmp(30) As String
Private blnCkZero As Boolean
Private intStyle As Integer '格式
Private strBatch As String  '批號
Private blnNew96 As Boolean '中信96新格式
Private strFrom  As String
Private strWhere As String
Private strUpdFields As String
Private strUpdWhere As String
Private strUpdOldWhere As String
Private blnUpdUCCode As Boolean
Private strUCCode As String
Private strUCName As String
Private strBillNoType As String
Private strChkFile As String
Private strChkRefNo As String
Private blnChkProc As Boolean
Private strHavSave As String
Private Const strMinRealStopDate As String = " MIN(DECODE(NVL(A.PAYKIND,0),1, " & _
                "DECODE(RealStopDate,NULL,TO_DATE('99991231','YYYYMMDD'),REALSTOPDATE)," & _
                "TO_DATE('99991231','YYYYMMDD'))) REALSTOPDATE "
Private Const strMinPayKind As String = " MAX(Nvl(A.PAYKIND,0)) PAYKIND "
Private strCutDate As String
Public Enum ChkPayKind
    giSingle = 0
    giAll = 1
End Enum
Public Function Action(Optional ByRef lngSecond As Long) As Boolean
    On Error GoTo chkErr
    Dim lngTime As Long
    Dim i As Long
    Dim objChkErr  As Object
    Dim blnErrNum As Boolean
    Dim blnShowErr As Boolean
    Dim strAryFile() As String
    strHavSave = Empty
        lngTime = Timer
        If Not InitData Then Exit Function
        If Not OpenData Then Exit Function
        If Not BeginTran Then Exit Function
                         
        lngSecond = Round(Timer - lngTime)
        
        MsgBox "共花費 : " & IIf(lngSecond > 60, lngSecond \ 60 & " 分 ", "") & (lngSecond Mod 60) & " 秒 !!" & _
                         IIf(lngErrCount > 0, vbCrLf & "錯誤筆數 :" & lngErrCount & "筆", ""), vbInformation, gimsgPrompt
                    
        strAryFile = Split(strChkFile, ",")
        blnShowErr = False
        If blnChkProc Then
            
            For i = 0 To UBound(strAryFile)
                If Not OpenFile(ErrFile, strErrPath & "BankCenterErr.Txt", False, 2) Then Exit Function
                ErrFile.WriteLine "--------------------------- clsBankCenterOut_" & Split(strAryFile(i), "@")(1) & " --------------------------------------"
                ErrFile.Close
                Set objChkErr = CreateObject("CheckMediaTEXT.clsExechkMediaText")
                With objChkErr
                    .ugiOpenFileType = 1 'Error文字檔開啟方式  0=Create  1=Append
                    .uClassName = "clsBankCenterOut_" & Split(strAryFile(i), "@")(1)
                    .uPathFile = Split(strAryFile(i), "@")(0) '資料路徑檔 路徑+檔名
                    .uErrPathFile = strErrPath & "BankCenterErr.Txt" '錯誤檔案 路徑+檔名
                    .uBankCode = Split(strAryFile(i), "@")(2)
                    .ugcnGi = gcnGi
                    .uOwner = strOwnerName
                    .uCompCode = strCompCode
                    .CheckMediaTextShow
                    'blnReturn = .uReturnOK '回傳是否已經結束 TRUE=執行完畢沒有問題  FALSE=執行中或有問題(需要驗證)
                    blnErrNum = .ublnError '回傳是否有ERROR的資料 TRUE=有錯誤資料  FALSE=無錯誤資料
                End With
                If blnErrNum Then
                    blnShowErr = True
                End If
            Next i
        End If
        If lngErrCount > 0 Or blnShowErr Then
            Shell "notepad " & strErrPath & "BankCenterErr.Txt", vbNormalNoFocus
        End If
        Action = True
        On Error Resume Next
        Set objChkErr = Nothing
    Exit Function
chkErr:
    ErrSub FormName, "Action"
End Function
Private Sub FormatCutDate()
  On Error GoTo chkErr
    Dim aTmp As String
    aTmp = Trim(Replace(Replace(clsStoreParameter.uRealDate, "/", ""), " ", ""))
    If Len(aTmp) = 6 Then
        aTmp = str(Val(Left(aTmp, 2)) + 1911) & Right(aTmp, 4)
    End If
    aTmp = Trim(aTmp)
    aTmp = Left(aTmp, 4) & "/" & Mid(aTmp, 5, 2) & "/" & Right(aTmp, 2)
    strCutDate = aTmp
    Exit Sub
chkErr:
  ErrSub FormName, "FormatCutDate"
End Sub
Private Function InitData() As Boolean
    On Error GoTo chkErr
        With clsStoreParameter
            blnBeOne = .uBeOne
            blnNT = .uNT
            blnCkZero = .uCkZero
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
            intStyle = .uStyle
            strBatch = .uBatch
            blnChkProc = .uChkProc
        End With
        
        lngPara24 = GetSystemParaItem("Para24", CStr(strCompCode), CStr(strServiceType), "SO043")
        InitData = True
    Exit Function
chkErr:
    Call ErrSub(FormName, "InitData")
End Function

Private Function GetRsTmp(ByRef rs As ADODB.Recordset) As Boolean
    On Error GoTo chkErr
    
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
        strFrom = Empty
        strWhere = Empty
        strUpdWhere = Empty
        strUpdFields = Empty
        blnNew96 = False
        '#3571 判斷主要銀行參考號是否為13，如果為13，單號要取虛擬帳號 By Kin 2007/10/29
'        If strBankHand <> "" Then
'            If GetRsValue("Select Nvl(RefNo,0) From " & strOwnerName & "CD018 Where CodeNo=" & strBankHand) = 13 Then
'                blnNew96 = True
'            End If
'        End If
'        If gcnGi.Execute("Select Count(*) From " & strOwnerName & "CD018 Where CodeNo " & strBankSQL & " And RefNo=13")(0) > 0 Then
'            blnNew96 = True
'        End If
        If intStyle = 1 Then blnNew96 = True
        If Not GetRS(rsTmp, "Select A.BankCode as BankCode,A.BankName as BankName,A.RealDate " & _
          " as TranDate,A.AccountNo,A.ShouldAmt as ShouldAmt,A.BillNo as BillNo_Old,A.BillNo as BillNo_New From " & strOwnerName & "So033 A Where RowId=''") Then Exit Function
        If Not CreateMDBTable(rsTmp, GetMDBTableName, cnn) Then Exit Function
        
        '#3571 使用中信新格式 By Kin 2007/10/30
        '**************************************************************************************************************************************************************************
        '#5055 沒有提到,但順便修改,CD013.REFNO=3不再出帳 By Kin 2009/04/30
        
        If blnNew96 Then
            strFrom = " " & strOwnerName & "SO033 A, " & strOwnerName & "CD013"
            If strChooseSQL <> "" Then
                If InStr(1, strChooseSQL, "C.") Then
                    strFrom = strFrom & "," & strOwnerName & "So014 C "
                    '#3417 更新UCCode與UCName的 Table與 條件 By Kin 2007/12/06
                    strUpdFields = strOwnerName & "SO014 C"
                  '  strUpdWhere = strUpdWhere & IIf(strWhere = "", "", " And ") & " A.AddrNo=C.AddrNo "
                    
                    strWhere = strWhere & IIf(strWhere = "", "", " And ") & " A.AddrNo=C.AddrNo "
                End If
                If InStr(1, strChooseSQL, "E.") Then
                    strFrom = strFrom & "," & strOwnerName & "So002 E "
                    
                    '#3417 更新UCCode與UCName的 Table與 條件 By Kin 2007/12/06
                    strUpdFields = IIf(Len(strUpdFields) = 0, "", strUpdFields & ",") & strOwnerName & "SO002 E"
                   ' strUpdWhere = strUpdWhere & IIf(strWhere = "", "", " And ") & " A.ServiceType = E.ServiceType And A.CustId=E.CustId And A.CompCode=E.CompCode "
                    
                    strWhere = strWhere & IIf(strWhere = "", "", " And ") & " A.ServiceType = E.ServiceType And A.CustId=E.CustId And A.CompCode=E.CompCode "
                End If
                
                '#3627 要以銀行別分檔案，所以要多加A.BankCode By Kin 2007/11/21
                '#5055 沒提到,但順便修改,CD013.REFNO=3不再出帳 By Kin 2009/04/30
                '#5218 櫃台已收要用Nvl的方式 By Kin 2009/08/05
                '#5564 新增參考7,8與CD013.PayOK=1代表已收 By Kin 2010/05/17
                '#5683 增加RealStopdate與PayKind By Kin 2010/08/10
                strWhere = strWhere & IIf(strWhere = "", "", " And ") & strChooseSQL & " And A.CancelFlag=0 And A.UCCode is Not Null " & _
                            IIf(blnCkZero = False, "And A.ShouldAmt > 0", "")
                strSQL = "SELECT " & GetUseIndexStr("So033", "ShouldDate") & " DISTINCT A.CitiBankATM BillNo," & _
                        " A.CustId,A.AccountNo, " & strMinPayKind & " , " & strMinRealStopDate & " , " & _
                        " A.MediaBillNo,A.AddrNo,A.BankCode,Sum(ShouldAmt) ShouldAmt" & IIf(blnNew96, ",CitiBankATM OldBillNo", "") & " From " & _
                        strFrom & " Where " & _
                        strWhere & " And A.UCCode=CD013.CodeNo And Nvl(CD013.REFNO,0) Not In(3,7) " & _
                        " AND Nvl(CD013.PayOK,0) = 0 " & _
                        " Group By A.CitiBankATM,A.CustId,A.AccountNo,A.MediaBillNo,A.AddrNo,A.BankCode "
                strBillNoType = "CitiBankATM"
            End If
        '**************************************************************************************************************************************************************************
        Else
            '#5055 沒提到,但順便修改,當CD013.REFNO=3不再出帳 By Kin 2009/04/30
            If lngPara24 = 0 Then
                strFrom = " " & strOwnerName & "SO033 A, " & strOwnerName & "CD013 "
                If strChooseSQL <> "" Then
                    If InStr(1, strChooseSQL, "C.") Then
                        strFrom = strFrom & "," & strOwnerName & "So014 C "
                        
                        '#3417 更新UCCode與UCName的 Table與 條件 By Kin 2007/12/06
                        strUpdFields = strOwnerName & "SO014 C"
                        
                        strWhere = strWhere & IIf(strWhere = "", "", " And ") & " A.AddrNo=C.AddrNo "
                    End If
                    If InStr(1, strChooseSQL, "E.") Then
                        strFrom = strFrom & "," & strOwnerName & "So002 E "
                        
                        '#3417 更新UCCode與UCName的 Table與 條件 By Kin 2007/12/06
                        strUpdFields = IIf(Len(strUpdFields) = 0, "", strUpdFields & ",") & strOwnerName & "SO002 E"
                        
                        strWhere = strWhere & IIf(strWhere = "", "", " And ") & " A.ServiceType = E.ServiceType And A.CustId=E.CustId And A.CompCode=E.CompCode "
                    End If
                End If
                '#5055 沒提到,順便修改,CD013.REFNO=3不再出帳 By Kin 2009/04/30
                '#5218 櫃台已收要用Nvl的方式 By Kin 2009/08/05
                '#5564 新增參考7,8與CD013.PayOK=1代表已收 By Kin 2010/05/17
                '#5683 增加RealStopdate與PayKind By Kin 2010/08/10
                strWhere = strWhere & IIf(strWhere = "", "", " And ") & strChooseSQL & " And A.CancelFlag=0 And A.UCCode is Not Null " & _
                            IIf(blnCkZero = False, "And A.ShouldAmt > 0", "")
                strSQL = "SELECT " & GetUseIndexStr("So033", "ShouldDate") & " DISTINCT A.BillNo," & _
                         " A.CustId,A.AccountNo,A.BankCode,A.AddrNo," & strMinPayKind & " , " & strMinRealStopDate & " , " & _
                         " A.MediaBillNo, Sum(ShouldAmt) ShouldAmt From " & _
                            strFrom & " Where " & _
                            strWhere & " And A.UCCode=CD013.CodeNo And Nvl(CD013.REFNO,0) Not In (3,7) " & _
                            " AND Nvl(CD013.PayOK,0) = 0 " & _
                            " Group By A.BillNo,A.CustId,A.AccountNo,A.BankCode,A.AddrNo,A.MediaBillNo "
                strBillNoType = "BillNo"
             Else
                '#5055 沒提到,順便修改,CD013.REFNO=3不再出帳 By Kin 2009/04/30
                strFrom = " " & strOwnerName & "SO033 A, " & strOwnerName & "CD013 "
                    If strChooseSQL <> "" Then
                    If InStr(1, strChooseSQL, "C.") Then
                        strFrom = strFrom & "," & strOwnerName & "So014 C "
                        
                        '#3417 更新UCCode與UCName的 Table與 條件 By Kin 2007/12/06
                        strUpdFields = strOwnerName & "SO014 C"
                        
                        strWhere = strWhere & IIf(strWhere = "", "", " And ") & " A.AddrNo=C.AddrNo "
                    End If
                    If InStr(1, strChooseSQL, "E.") Then
                        strFrom = strFrom & "," & strOwnerName & "So002 E "
                        
                        '#3417 更新UCCode與UCName的 Table與 條件 By Kin 2007/12/06
                        strUpdFields = IIf(Len(strUpdFields) = 0, "", strUpdFields & ",") & strOwnerName & "SO002 E"
                        
                        strWhere = strWhere & IIf(strWhere = "", "", " And ") & " A.ServiceType = E.ServiceType And A.CustId=E.CustId And A.CompCode=E.CompCode "
                    End If
                End If
                '#5218 櫃台已收要用Nvl的方式 By Kin 2009/08/05
                '#5564 新增參考7,8與CD013.PayOK=1代表已收 By Kin 2010/05/17
                '#5683 增加RealStopdate與PayKind By Kin 2010/08/10
                strWhere = strWhere & IIf(strWhere = "", "", " And ") & strChooseSQL & " And A.CancelFlag=0 And A.UCCode is Not Null " & _
                                    IIf(blnCkZero = False, "And A.ShouldAmt > 0", "")
                strSQL = "SELECT " & GetUseIndexStr("So033", "ShouldDate") & " DISTINCT A.Custid," & _
                        " A.MediaBillNo BillNo,A.AccountNo," & _
                        " A.BankCode,A.AddrNo, " & strMinPayKind & " , " & strMinRealStopDate & " , " & _
                        " Sum(ShouldAmt) ShouldAmt From " & strFrom & " Where " & _
                         strWhere & " And A.UCCode=CD013.CodeNo And Nvl(CD013.REFNO,0) Not In(3,7) " & _
                         " AND Nvl(CD013.PayOk,0) = 0 " & _
                         " Group By A.Custid,A.MediaBillNo,A.AccountNo,A.BankCode,A.AddrNo "
                strBillNoType = "MediaBillNo"
                
             End If
        End If
        strUpdOldWhere = strWhere
        If Not GetRS(rs, strSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic, , , True) Then Exit Function
        strReturnSQL = rs.Source
    GetRsTmp = True
    Exit Function
chkErr:
    ErrSub FormName, "GetRsTmp"
End Function

Private Function GetRsUpd(ByRef rs As ADODB.Recordset) As Boolean
    On Error GoTo chkErr
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
            strSQL = "SELECT " & GetUseIndexStr("So033", "ShouldDate") & " A.CustId, A.BillNo, A.MediaBillNo,A.AccountNo,A.BankCode,A.AddrNo From " & _
                        strFrom & " Where " & strWhere
        If Not GetRS(rs, strSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic, , , True) Then Exit Function
        strReturnSQL = rs.Source
    GetRsUpd = True
    Exit Function
chkErr:
    ErrSub FormName, "GetRsUpd"
End Function

Private Function InsertHead(rs As ADODB.Recordset) As Boolean
    On Error GoTo chkErr
    Dim intReserve As Integer
        '區別碼1 , 發件單位8, 收件單位8, 轉帳類別3, 入 / 扣帳日6, 性質別1, 保留欄93
        rsCD018.MoveFirst
        
        If blnBeOne Then rsCD018.Filter = "CodeNo = " & strBankHand
        Do While Not rsCD018.EOF
        '#3627 中信96新格式資料調整 By Kin 2007/11/21
            If blnNew96 Then
                intReserve = 93
                If Val(rsCD018("TextLength") & "") > 0 Then intReserve = 93 - (120 - Val(rsCD018("TextLength") & ""))
                WriteTextLine DataFile(rsCD018.AbsolutePosition), "1" & GetString(strInvoiceId, 8, giLeft) & _
                        GetString(Left(rsCD018("BankID2") & "", 3), 8, giLeft) & _
                        GetString(rsCD018("Class") & "", 3) & _
                        GetString(strRealDate, 6) & _
                        GetString(1, 1) & _
                        GetString("", intReserve)
                rsCD018.MoveNext
            
            Else
                intReserve = 93
                '5596 土銀格式改變(CD018.REFNO=14) By Kin 2010/04/07
                If Val(rsCD018("RefNo") & "") = 14 Then
                    intReserve = 169
                    WriteTextLine DataFile(rsCD018.AbsolutePosition), "1" & GetString(rsCD018("BankId"), 8, giLeft) & _
                            GetString(rsCD018("CorpId") & "", 8, giLeft) & _
                            GetString(rsCD018("Class") & "", 5) & _
                            GetString(strRealDate, 8, giRight, True) & _
                            GetString(1, 1) & _
                            GetString("", intReserve)
                Else
                    '#5596 REFNO=15代表渣打銀行,日期要補0
                    If Val(rsCD018("TextLength") & "") > 0 Then intReserve = 93 - (120 - Val(rsCD018("TextLength") & ""))
                    WriteTextLine DataFile(rsCD018.AbsolutePosition), "1" & GetString(rsCD018("BankId"), 8, giLeft) & _
                            GetString(rsCD018("CorpId") & "", 8, giLeft) & _
                            GetString(rsCD018("Class") & "", 3) & _
                            IIf(Val(rsCD018("RefNo") & "") = 15, GetString(strRealDate, 7, giRight, True), GetString(strRealDate, 6)) & _
                            GetString(1, 1) & _
                            GetString("", intReserve)
                End If
                rsCD018.MoveNext
            End If
        Loop
        rsCD018.Filter = ""
        InsertHead = True
    Exit Function
chkErr:
    ErrSub FormName, "InsertHead"
End Function

Private Function InsertDetail(rs As ADODB.Recordset, Optional ByRef rsOld As ADODB.Recordset) As Boolean
    On Error GoTo chkErr
    Dim lngAmount As Long, lngOldCustId As Long
    Dim strWriteLine As String, strLimitation As String, strWhere As String
    Dim rsTmp As New ADODB.Recordset
    Dim rsSo106 As New ADODB.Recordset
    Dim rsSo002a As New ADODB.Recordset
    Dim rsInv As New ADODB.Recordset
    Dim strAccNameID As String
    Dim strSQL As String, strInvNo As String
    Dim strCustId As String
    Dim strWErr As String
    Dim strUpdTime As String
    strUpdTime = GetDTString(RightNow)
        If rs.RecordCount > 0 Then rs.MoveFirst
        cnn.BeginTrans
        While Not rs.EOF
            strWErr = ""
            If rs("ShouldAmt") > 0 Then
                '#3571 如果不是中信96新格式才需要找銀行
                If Not blnNew96 Then
                    If Not FindCD018(rs) Then
                        lngAmount = 0
                        lngErrCount = lngErrCount + 1: GoTo lNext
                    End If
                End If
                If Not IsPayKindOK(rs, Replace(strCutDate, "/", ""), giAll, lngPara24) Then
                    strWErr = "客戶編號:" & rs("Custid") & "  帳號:" & rs("AccountNo") & " 收視截止日不正確" & _
                            "  收視截止日期: " & rs("RealStopDate") & "  單據編號:" & rs("BillNo") & "  繳付類別：預付制"
                    lngErrCount = lngErrCount + 1
                    GoTo lNext
                End If
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
                lngAmount = rs("ShouldAmt") & ""
                
                If rs("AccountNo") & "" <> "" Then
                    strSQL = "Select CUSTID,AccountNameID From " & strOwnerName & "So106 Where AccountID='" & rs("AccountNo") & "" & "'" _
                              & " And StopDate Is Null And SnactionDate Is Not Null"
                    If Not GetRS(rsSo106, strSQL, gcnGi, , , , , 1) Then Exit Function
                    If rsSo106.EOF Then
                        strAccNameID = ""
                        strWErr = "客戶編號:" & rs("Custid") & "  帳號:" & rs("AccountNo") & "對應不到SO106信用卡及轉帳申請記錄檔"
                    Else
                        strAccNameID = rsSo106("AccountNameId") & ""
                    End If
                Else
                    strAccNameID = ""
                    strWErr = "客戶編號:" & rs("Custid") & "  帳號為空值"

                End If
                
                '#3627 中國信託新格式，要使用SO106.CustID，必須一定要有值 By Kin 2007/11/21
                '*****************************************************************************************************************************
                If blnNew96 Then
                    If rs("AccountNo") & "" = "" Then
                        strWErr = "客戶編號:" & rs("Custid") & "  帳號為空值"
                        lngErrCount = lngErrCount + 1
                        GoTo lNext
                    End If
                    If Not rsSo106.EOF Then
                        strCustId = rsSo106("CustId") & ""
                    Else
                        If strAccNameID = "" Then
                            strWErr = "客戶編號:" & rs("Custid") & "  帳號:" & rs("AccountNo") & "對應不到SO106信用卡及轉帳申請記錄檔"
                            lngErrCount = lngErrCount + 1
                            GoTo lNext
                        End If
                        
                    End If
                Else
                    strCustId = rs("Custid") & ""
                End If
                '****************************************************************************************************************************
                If blnNew96 Then
                    '#3627 如果使用中國信託96新格式，判斷是否有使用合併銀行資料
                    If blnBeOne Then
                        rsCD018.Filter = "CodeNo = " & strBankHand
                    Else
                        rsCD018.Find "CodeNo = " & rs("BankCode")
                    End If
                    '#3910 將strMediaBillNo 代入rsOld("BillNo") By Kin 2008/04/29
                    If Not WriteDetailToText(strInvoiceId & "", Left(rsCD018("BankID2") & "", 3), rsCD018("Class") & "", rs("AccountNo") & "" _
                            , lngAmount, strInvoiceId, rs("BillNo") & "", strAccNameID & "", rsCD018("RefNo") & "", rsCD018("CodeNo") & "", rsCD018("PrgName") & "" _
                            , rs("AddrNo") & "", strLimitation, Val(rsCD018("TextLength") & ""), strAccNameID, IIf(rsOld Is Nothing, rs("BillNo") & "", rsOld("BillNo") & ""), strCustId, rsCD018("ActNo") & "") Then Exit Function

                Else
                    If Not WriteDetailToText(rsCD018("BankId") & "", rsCD018("CorpId") & "", rsCD018("Class") & "", rs("AccountNo") & "" _
                            , lngAmount, strInvoiceId, rs("BillNo") & "", strAccNameID & "", rsCD018("RefNo") & "", rs("BankCode") & "", rsCD018("PrgName") & "" _
                            , rs("AddrNo") & "", strLimitation, Val(rsCD018("TextLength") & ""), strAccNameID, rs("BillNo") & "", strCustId, rsCD018("ActNo") & "") Then Exit Function
                End If
                If blnBeOne Then
                '存總計筆數及金額
                    lngBankCount(1) = lngBankCount(1) + 1
                    lngBankAmt(1) = lngBankAmt(1) + lngAmount
                Else
                    '存各銀行小計及金額
                    lngBankCount(rsCD018.AbsolutePosition) = lngBankCount(rsCD018.AbsolutePosition) + 1
                    lngBankAmt(rsCD018.AbsolutePosition) = lngBankAmt(rsCD018.AbsolutePosition) + lngAmount
                End If
            Else
                 strWErr = "客戶編號:" & rs("Custid") & "  應收金額為" & rs("ShouldAmt") & "不予產出"
            End If
            '#3417 更新UCCode與UCName的語法 By Kin 2007/12/06
            '#4388 增加更新異動人員與異動時間 By Kin 2009/04/30
            '#5055 順便修改,櫃台已收不更新未收原因 By Kin 2009/04/30
            '#5218 櫃台已收要用Nvl的方式 By Kin 2009/08/05
            '#5564 新增參考7,8與CD013.PayOK=1代表已收 By Kin 2010/05/17
            If blnUpdUCCode Then
              On Error GoTo lNext
                strUpdWhere = Empty
                strUpdWhere = "UPDATE " & strOwnerName & " SO033 A Set UCCode=" & strUCCode & ",UCName='" & strUCName & "' " & _
                                ",UPDEN='" & strUpdName & "',UPDTIME='" & strUpdTime & "' " & _
                                " Where RowID IN(Select A.RowID From " & strUpdFields & _
                                IIf(strUpdFields <> "", ",", "") & strOwnerName & "CD013, " & strOwnerName & "SO033 A " & _
                                " Where " & strUpdOldWhere & _
                                " And A." & strBillNoType & "='" & rs("BillNo") & "'" & _
                                " And A.ACCOUNTNO='" & rs("ACCOUNTNO") & "'" & _
                                " And A.BankCode=" & rs("BankCode") & _
                                " And A.UCCode=CD013.CodeNo And Nvl(CD013.REFNO,0) Not IN (3,7) " & _
                                " AND NVL(CD013.PAYOK,0) = 0 )"
                gcnGi.Execute strUpdWhere
            End If
            
lNext:
            If Not WriteTextLine(ErrFile, strWErr) Then Exit Function
'            lngOldCustId = rs("BillNo")
            rsCD018.Filter = ""
            rs.MoveNext
            On Error Resume Next
            If Not rsOld Is Nothing Then
                If rsOld.State = adStateOpen Then
                    rsOld.MoveNext
                End If
            End If
            DoEvents
        Wend
        cnn.CommitTrans
        InsertDetail = True
    Exit Function
chkErr:
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
    On Error GoTo chkErr
        strOldBillNo = rs("BillNo") & ""
        lngOldCustId = rs("CustId")
        strBankCode = rs("BankCode") & ""
        strOldBankId = rsCD018("BankId") & ""
        strOldCorpId = rsCD018("CorpId") & ""
        strOldClass = rsCD018("Class") & ""
        strBankRefNo = rsCD018("RefNo") & ""
        strPrgName = rsCD018("PrgName") & ""
'        strActNo = rsCD018("Actno") & ""
        strOldAccountNo = rs("AccountNo") & ""
        strOldID = rs("ID") & ""
        GetOldValue = True
    Exit Function
chkErr:
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
                 strMediaBillNo As String, strCustId As String, _
                 strActNo As String) As Boolean
    On Error GoTo chkErr
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
        '#5596 增銀格式有改變(CD018.REFNO=5) By Kin 2010/04/07
        If Val(strBankRefNo) = 14 Then
            intReserve = 22
        Else
            intReserve = 12
        End If
        If intTextLength > 0 Then intReserve = 12 - (120 - intTextLength)
        '#5596 土銀格式有改變(CD018.REFNO=5) By Kin 2010/04/07
        '#5596 REFNO=15 渣打銀行日期要補0
        strWriteLine = "2" & _
            GetString(strOldBankId & "", 8, giLeft) & _
            GetString(strOldCorpId & "", 8, giLeft) & _
            GetString(strOldClass & "", IIf(Val(strBankRefNo) = 14, 5, 3)) & _
            IIf(Val(strBankRefNo) = 14, GetString(strRealDate, 8, giRight, True), _
                IIf(strBankRefNo = 15, GetString(strRealDate, 7, giRight, True), GetString(strRealDate, 6)))
            '參考號10  for 觀昇 Penny 提出
            strChkRefNo = ""
            If strBankRefNo = 10 Then
                strWriteLine = strWriteLine & GetString(strOldAccountNo, 14, giLeft, True)
            Else
                '#5596 土銀格式有改變(CD018.REFNO=5) By Kin 2010/04/07
                If Val(strBankRefNo) = 14 Then
                    strWriteLine = strWriteLine & GetString(strOldAccountNo, 20, giRight, True)
                Else
                    strWriteLine = strWriteLine & GetString(strOldAccountNo, 14, giRight, True)
                End If
            End If
            strWriteLine = strWriteLine & GetString(lngAmount, 12, giRight, True) & "00"
            '參考號11 問題集1192:小杜-20041011 Edit:Crystal-20041013
            If strBankRefNo = 11 Then
                strWriteLine = strWriteLine & GetString(strOldCorpId, 8)
            Else
                strWriteLine = strWriteLine & GetString(strInvoiceId, 8)
            End If
            '#5596 土銀格式有改變(CD018.REFNO=5) By Kin 2010/04/07
            If Val(strBankRefNo) = 14 Then
                strWriteLine = strWriteLine & "9999" & GetUserAreaStr(strOldBillNo, strPrgName, strBankRefNo, strID, strAddrNo, strLimitation, strBillNoOut, strInvNo, strMediaBillNo, strCustId, strActNo)
            Else
                strWriteLine = strWriteLine & "99" & GetUserAreaStr(strOldBillNo, strPrgName, strBankRefNo, strID, strAddrNo, strLimitation, strBillNoOut, strInvNo, strMediaBillNo, strCustId, strActNo)
            End If
            '如果為郵局格式則在109的位址上加註 P
            If strBankRefNo = 6 Then
                Dim strLenSQL As String
                Dim rsLenAct As New ADODB.Recordset
              '問題集2830 增加判斷CD018 ACTLENGTH長度，分別填入P，G
                strLenSQL = "Select ActLength From " & GetOwner & " CD018 Where CodeNo=" & strOldBankCode
                Set rsLenAct = gcnGi.Execute(strLenSQL)
                If rsLenAct(0) = 14 Then
                    strWriteLine = strWriteLine & GetString("P", 1)
                Else
                    strWriteLine = strWriteLine & GetString("G", 1)
                End If
                Set rsLenAct = Nothing
                strWriteLine = strWriteLine & GetString("", intReserve - 1)
            ElseIf strBankRefNo = 9 Then
                strWriteLine = strWriteLine & GetString(strInvNo, 12)
            '#5596 土銀格式改變,要多填一個帳戶ID
            ElseIf strBankRefNo = 14 Then
                strWriteLine = strWriteLine & GetString(strID, 10) & Space(2) & GetString("", intReserve)
            Else
                strWriteLine = strWriteLine & GetString("", intReserve)
            End If
            
        lngBR = rsCD018.AbsolutePosition
        Call FindCD018(, , strOldBankCode)
        If blnBeOne Then
            If Not WriteTextLine(DataFile(1), strWriteLine) Then Exit Function
            strChkFile = strDataFileName(1) & "@" & strChkRefNo & "@" & rsCD018("CodeNo") & ""
        Else
            If Not WriteTextLine(DataFile(rsCD018.AbsolutePosition), strWriteLine) Then Exit Function
            If InStr(1, strHavSave, rsCD018.AbsolutePosition) <= 0 Then
                strChkFile = IIf(strChkFile = Empty, strDataFileName((rsCD018.AbsolutePosition)) & "@" & strChkRefNo & "@" & rsCD018("CodeNo") & "", strChkFile & "," & strDataFileName((rsCD018.AbsolutePosition)) & "@" & strChkRefNo & "@" & rsCD018("CodeNo") & "")
            End If
            strHavSave = IIf(strHavSave = Empty, rsCD018.AbsolutePosition & "", strHavSave & "," & rsCD018.AbsolutePosition & "")
        End If
        
        '填入暫存檔
        If Not ExecuteCommand("Insert Into " & GetMDBTableName & " (BankCode,BankName,TranDate,AccountNo,ShouldAmt,BillNo_Old,BillNo_New) Values (" & _
                GetNullString(strOldBankCode, giLongV) & "," & GetNullString(GetBankDescription(strOldBankCode) & "") & ",'" & GetRealDateTran(CStr(strRealDate)) & "'," & _
                GetNullString(strOldAccountNo) & "," & lngAmount & ",'" & strBillNoOut & "','" & strOldBillNo & "')", cnn) Then Exit Function
        If rsCD018.AbsolutePosition <> lngBR Then
        End If
        If lngBR > 0 Then rsCD018.AbsolutePosition = lngBR
        WriteDetailToText = True
    Exit Function
chkErr:
    ErrSub FormName, "WriteDetailToText"
End Function

Private Function GetBankDescription(strBankCode As String) As String
    On Error GoTo chkErr
        rsCD018.Filter = "CodeNo = " & strBankCode
        GetBankDescription = rsCD018("Description") & ""
        rsCD018.Filter = ""
    Exit Function
chkErr:
    ErrSub FormName, "GetBankDescription"
End Function

Private Function FindCD018(Optional ByRef rs As ADODB.Recordset, _
        Optional lngOldCustId As Long, Optional strBankCode As String = "") As Boolean
    On Error GoTo chkErr
    Dim lngCustId As Long
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
chkErr:
    ErrSub FormName, "FindCD018"
End Function

'尾筆
Private Function InsertFinal(rs As ADODB.Recordset) As Boolean
    On Error GoTo chkErr
    Dim intReserve As Integer
        rsCD018.MoveFirst
        If blnBeOne Then rsCD018.Filter = "CodeNo = " & strBankHand
        '區別碼1,發件
        Do While Not rsCD018.EOF
            intReserve = 42
            If Val(rsCD018("TextLength") & "") > 0 Then intReserve = 42 - (120 - Val(rsCD018("TextLength") & ""))
            '#3627 使用中國信託96新格式 By Kin 2007/11/21
            
            If blnNew96 Then
                WriteTextLine DataFile(rsCD018.AbsolutePosition), "3" & GetString(strInvoiceId, 8, giLeft) & _
                        GetString(Left(rsCD018("BankID2") & "", 3), 8, giLeft) & _
                        GetString(rsCD018("Class") & "", 3) & _
                        GetString(strRealDate, 6) & _
                        GetString(lngBankAmt(rsCD018.AbsolutePosition), 14, giRight, True) & _
                        "00" & _
                        GetString(lngBankCount(rsCD018.AbsolutePosition), 10, giRight, True) & _
                        String(26, "0") & _
                        GetString("", intReserve)
            Else
                '#5596 土銀格式改變 By Kin 2010/04/09
                If Val(rsCD018("RefNo") & "") = 14 Then
                    intReserve = 118
                    WriteTextLine DataFile(rsCD018.AbsolutePosition), "3" & GetString(rsCD018("BankId") & "", 8, giLeft) & _
                                GetString(rsCD018("CorpId") & "", 8, giLeft) & _
                                GetString(rsCD018("Class") & "", 5) & _
                                GetString(strRealDate, 8, giRight, True) & _
                                GetString(lngBankAmt(rsCD018.AbsolutePosition), 14, giRight, True) & _
                                "00" & _
                                GetString(lngBankCount(rsCD018.AbsolutePosition), 10, giRight, True) & _
                                String(26, "0") & _
                                GetString("", intReserve)
                Else
                    '#5596 渣打銀行(REFNO=15)日期要改7碼 By Kin 2010/04/09
                    WriteTextLine DataFile(rsCD018.AbsolutePosition), "3" & GetString(rsCD018("BankId") & "", 8, giLeft) & _
                            GetString(rsCD018("CorpId") & "", 8, giLeft) & _
                            GetString(rsCD018("Class") & "", 3) & _
                            IIf(Val(rsCD018("RefNo") & "") = 15, GetString(strRealDate, 7, giRight, True), GetString(strRealDate, 6)) & _
                            GetString(lngBankAmt(rsCD018.AbsolutePosition), 14, giRight, True) & _
                            "00" & _
                            GetString(lngBankCount(rsCD018.AbsolutePosition), 10, giRight, True) & _
                            String(26, "0") & _
                            GetString("", intReserve)
                End If
            End If
            rsCD018.MoveNext
        Loop
        rsCD018.Filter = ""
        InsertFinal = True
    Exit Function
chkErr:
    ErrSub FormName, "InsertFinal"
End Function
Private Function HaveDetail(ByRef rs As ADODB.Recordset) As Boolean
  On Error GoTo chkErr
    
    Exit Function
chkErr:
  ErrSub forname, "HaveDetail"
End Function
Private Function BeginTran() As Boolean
    On Error GoTo chkErr
    Dim rsTmp As New ADODB.Recordset
    Dim strBankId As String, strSQL As String
    Dim strOldBillNo As String
    Dim lngCountCustId As Long, lngErrCount As Long
    Dim lngPara24 As Long
    Dim rsUpd As New ADODB.Recordset
    Dim rsCD013 As New ADODB.Recordset
    Dim strKillFileName As String
    Dim strExtName As String
    
    '********************************************************************************************************************************************************
    '#3417 電子檔匯出時,要填入未收原因(RefNo=4) By Kin 2007/12/06
    If Not GetRS(rsCD013, "Select * From " & strOwnerName & "CD013 Where RefNo=4 And StopFlag<>1 Order By CodeNo Desc", gcnGi, adUseClient, adOpenKeyset) Then Exit Function
    If rsCD013.EOF Then
        blnUpdUCCode = False
    Else
        blnUpdUCCode = True
        strUCCode = rsCD013("CodeNo") & ""
        strUCName = rsCD013("Description") & ""
    End If
    '*********************************************************************************************************************************************************
        '#5683 取得轉帳扣款日期 By Kin 2010/08/10
        Call FormatCutDate
        '取資料
        If Not GetRsTmp(rsTmp) Then Exit Function
         '檢查是否有用多帳號功能
          lngPara24 = GetSystemParaItem("Para24", CStr(strCompCode), CStr(strServiceType), "SO043")
        If lngPara24 = 1 Then
            If Not GetRsUpd(rsUpd) Then Exit Function
            '#5055 沒有提到,但順便修改,櫃台已收則不進行更新 By Kin 2009/04/30
            If Not BatchUpdateMediaBillNo(rsUpd, strChoose33 & " And A.UCCode NOT IN (Select CodeNO From " & strOwnerName & "CD013 Where REFNO=3)", gcnGi) Then Exit Function
        End If
        '#3627 測試不OK,當完全沒有資料時,也要清空檔案 By Kin 2007/11/26
        If rsTmp.RecordCount = 0 Then
            If blnNew96 Then
                On Error Resume Next
                Call CloseFS(False)
                rsCD018.MoveFirst
                Do While Not rsCD018.EOF
                    strKillFileName = Mid(rsCD018("FileName1") & "", 1, InStrRev(rsCD018("FileName1") & "", "\"))
                    strExtName = Left(rsCD018("BankId2") & "", 3)
                    strKillFileName = strKillFileName & "FISCP_" & Left(Trim(strInvoiceId), 8) & "_" & _
                                    str(Val(Left(strRealDate, 2)) + 1911) & Right(strRealDate, 4) & "_" & _
                                    strBatch & "." & strExtName
                    strKillFileName = Replace(strKillFileName, " ", "")
                    Kill strKillFileName
                    rsCD018.MoveNext
                Loop
            End If
            MsgBox "無資料！請重新操作！", vbInformation, "訊息！"
            GoTo KillFile
        Else
            If blnNew96 Then
                '#3910 以rsOld Keep 舊的CitiBankATm By Kin 2008/04/29
                Dim rsOld As New ADODB.Recordset
                Set rsOld = rsTmp.Clone
                Set rsOld.ActiveConnection = Nothing
                If Not BatchUpdateCitiBankATM(strWhere, strFrom, gcnGi) Then Exit Function
                '*******************************************************************************************************
                '#3910 因為CitiBankAMT重新產生所以RecordSet需重新尋找一次資料 By Kin 2008/04/29
                If Not GetRS(rsTmp, rsTmp.Source, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
                '*******************************************************************************************************
            End If
            If Not InsertHead(rsTmp) Then GoTo KillFile
            If Not InsertDetail(rsTmp, rsOld) Then GoTo KillFile
            If Not InsertFinal(rsTmp) Then GoTo KillFile
        End If
        Call CloseFS(False)
        
        '#3627 測試不OK,如果挑選的銀行無資料,則不要產生檔案,在此段將資料刪除 By Kin 2007/11/22
        On Error Resume Next
        If blnNew96 Then
            rsCD018.MoveFirst
            rsTmp.MoveLast
            Do While Not rsCD018.EOF
                rsTmp.Filter = ""
                rsTmp.Find "BankCode =" & rsCD018("CodeNo") & "", , adSearchForward, 1
                If rsTmp.EOF Then
                    strKillFileName = Mid(rsCD018("FileName1") & "", 1, InStrRev(rsCD018("FileName1") & "", "\"))
                    strExtName = Left(rsCD018("BankId2") & "", 3)
                    strKillFileName = strKillFileName & "FISCP_" & Left(Trim(strInvoiceId), 8) & "_" & _
                                    str(Val(Left(strRealDate, 2)) + 1911) & Right(strRealDate, 4) & "_" & _
                                    strBatch & "." & strExtName
                    strKillFileName = Replace(strKillFileName, " ", "")
                    Kill strKillFileName
                Else
                    rsTmp.Filter = "PAYKIND = 1 OR REALSTOPDATE <= '" & strCutDate & "'"
                    If rsTmp.RecordCount = 0 Then
                        strKillFileName = Mid(rsCD018("FileName1") & "", 1, InStrRev(rsCD018("FileName1") & "", "\"))
                        strExtName = Left(rsCD018("BankId2") & "", 3)
                        strKillFileName = strKillFileName & "FISCP_" & Left(Trim(strInvoiceId), 8) & "_" & _
                                    str(Val(Left(strRealDate, 2)) + 1911) & Right(strRealDate, 4) & "_" & _
                                    strBatch & "." & strExtName
                        strKillFileName = Replace(strKillFileName, " ", "")
                        Kill strKillFileName
                    End If
                End If
                rsCD018.MoveNext
            Loop
            
        End If
        BeginTran = True
        On Error Resume Next
        rsTmp.Close
        Set rsTmp = Nothing
        CloseRecordset rsCD013
        CloseRecordset rsOld
    Exit Function
KillFile:
    '刪除檔案..
    Call CloseFS(True)
    Exit Function
chkErr:
    Resume 0
    Call ErrSub(FormName, "BeginTran")
End Function

Private Function ChkErrCustId(rs As ADODB.Recordset, _
        ByRef lngErrCount As Long) As Boolean
    On Error GoTo chkErr
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
chkErr:
    ErrSub FormName, "ChkErrCustId"
End Function

Private Function OpenData() As Boolean
    On Error GoTo chkErr
    Dim lngCount As Long
    Dim strFileName As String
    Dim strExtName As String
    Dim strFileNameDate As String
        If Not GetRS(rsCD018, "Select * From " & strOwnerName & "CD018 Where CodeNo  " & strBankSQL & " order by CodeNo", gcnGi) Then Exit Function
        
        If rsCD018.RecordCount = 0 Then Exit Function
        strInvoiceId = GetSystemParaItem("InvoiceId", strCompCode & "", strServiceType & "", "SO041")
        If strInvoiceId = "" Then MsgBox "服務類別代碼為 :[ " & strServiceType & " ] 且公司別代碼為 : [" & strCompCode & "] 無對應之InvoiceID !! 請查證.. ": Exit Function
        rsCD018.MoveFirst
        If blnBeOne Then rsCD018.Filter = "CodeNo = " & strBankHand
        Do While Not rsCD018.EOF
            lngCount = rsCD018.AbsolutePosition
            strFileName = Left(rsCD018("FileName1") & "", InStrRev(rsCD018("FileName1") & "", "."))
            
            '#3571 如果為中信96年格式(Style=1 使用預設檔名) By Kin
            'If intStyle = 0 Then
            If intStyle = 0 Then
                If strFileName = "" Then strFileName = rsCD018("FileName1") & ""
                If strFileName = "" Then MsgBox "請設定銀行 : [" & rsCD018("CodeNO") & rsCD018("Description") & " ] 之轉帳輸出檔名 ": Exit Function
            Else
                strFileName = Mid(rsCD018("FileName1") & "", 1, InStrRev(rsCD018("FileName1") & "", "\"))
                If strFileName = "" Or rsCD018("FileName1") & "" = "" Then
                    MsgBox "請設定銀行 : [" & rsCD018("CodeNO") & rsCD018("Description") & " ] 之轉帳輸出路徑 "
                    Exit Function
                '#3571 檔名：FISCP_公司統編8碼_西元年日期8碼_批號3碼.銀行金資代號3碼 By Kin 2007/10/22
                Else
                    If rsCD018("CorpId") & "" = "" Then MsgBox "請設定銀行 : [" & rsCD018("CodeNO") & rsCD018("Description") & " ] 之收件單位", vbInformation, "訊息": Exit Function
                    If rsCD018("BankId2") & "" = "" Then MsgBox "請設定銀行 : [" & rsCD018("CodeNo") & rsCD018("Description") & " ] 之金融機構代號", vbInformation, "訊息": Exit Function
                    
                    '#3627副檔名改為取BankID2前三碼 By Kin 2007/11/21
                    '#5596 超過民國100年檔名會取錯 By Kin 2010/04/08
                    strExtName = Left(rsCD018("BankId2") & "", 3)
                    strFileNameDate = Right("0" & strRealDate, 7)
                    strFileName = strFileName & "FISCP_" & Left(Trim(strInvoiceId), 8) & "_" & _
                                    str(Val(Left(strFileNameDate, 3)) + 1911) & Right(strRealDate, 4) & "_" & _
                                    strBatch & "." & strExtName
                    strFileName = Replace(strFileName, " ", "")
                    
                End If
            End If
            
            If Not OpenFile(DataFile(lngCount), IIf(Val(rsCD018("RefNo") & "") <> 13, rsCD018("FileName1") & "", strFileName), True) Then Exit Function
            strDataFileName(lngCount) = IIf(Val(rsCD018("RefNo") & "") <> 13, rsCD018("FileName1") & "", strFileName)
            rsCD018.MoveNext
        Loop
        rsCD018.Filter = ""
        
        If Not OpenFile(ErrFile, strErrPath & "BankCenterErr.Txt", True) Then Exit Function
        If Not OpenFile(MemoFile, strErrPath & "BankCenterMem.Txt", True) Then Exit Function
        OpenData = True
    Exit Function
chkErr:
    ErrSub FormName, "OpenData"
End Function

Private Function CloseFS(Optional KillFile As Boolean = False) As Boolean
    On Error Resume Next
    Dim lngLoop As Long
    Dim strFileName As String
    Dim strTmpFile As String
    Dim LogFile(30) As Variant
    Dim File As Object
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
                    Kill strErrPath & "BankCenterErr.Txt"
                    Kill strErrPath & "BankCenterMem.Txt"
                End If
            End If
            If blnNT Then
                If rsCD018.AbsolutePosition <> 1 Then
                    Set LogFile(rsCD018.AbsolutePosition) = FSO.OpenTextFile(rsCD018("FileName1"), ForReading)
                    Tmp(rsCD018.AbsolutePosition) = LogFile(rsCD018.AbsolutePosition).ReadAll
                    File.Write Tmp(rsCD018.AbsolutePosition)
                    LogFile(rsCD018.AbsolutePosition).Close
                    Kill rsCD018("FileName1")
                Else
                    strTmpFile = rsCD018("FileName1")
                    Set File = FSO.OpenTextFile(strTmpFile, ForAppending, False, TristateFalse)
                End If
            End If
            rsCD018.MoveNext
        Loop
        File.Close
        rsCD018.Filter = ""
    Exit Function
chkErr:
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
Private Function GetCitiBankNo(strBillNo As String, Optional strMediaBillNo As String) As String
  On Error Resume Next
    
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
        GetMDBTableName = UCase("BankCenterOut")
End Function

Private Function GetUserAreaStr(strBillNo As String, strPrgName As String, _
    strRefNo As String, strID As String, strAddrNo As String, strLimitation As String, _
    ByRef strBillNoOut As String, strInvNo As String, strMediaBillNo As String, _
    strCustId As String, strActNo As String) As String
    On Error Resume Next
    '花旗銀行的   專用資料區44 銀行 (特殊單號+檢查碼+單號)65位元開始
    '  or  專用資料區44 郵局 (身份證字號ID+單號)
    '其他家銀行 單號 BillNO(15)
    '6 是郵局的
    Dim strUserArea As String
        If strRefNo = 1 Or strRefNo = 6 Or strRefNo = 10 Then
            '花旗的格式
            strUserArea = GetString(strID, 10) & GetBillNo_Old(strBillNo, strMediaBillNo)
            strBillNoOut = GetBillNo_Old(strBillNo, strMediaBillNo)
            If strRefNo = 6 Then
                strChkRefNo = "12"
            ElseIf strRefNo = 10 Then
                strChkRefNo = "10"
            Else
                strChkRefNo = "0"
            End If
        ElseIf strRefNo = 2 Or strRefNo = 9 Or strRefNo = 11 Then
            '真正的金資
            strUserArea = GetSpecialBillNo(strBillNo, strRefNo, strID, strAddrNo) & GetBillNo_New(strBillNo, strMediaBillNo)
            strBillNoOut = GetBillNo_New(strBillNo, strMediaBillNo)
            If strRefNo = 9 Then
                strChkRefNo = "13"
            ElseIf strRefNo = 11 Then
                strChkRefNo = "11"
            Else
                strChkRefNo = "1"
            End If
        ElseIf strRefNo = 3 Then
            '台灣銀行用(吉元)
            strUserArea = GetBillNo_New(strBillNo, strMediaBillNo) & Space(9) & GetString(strID, 10)
            strBillNoOut = GetBillNo_New(strBillNo, strMediaBillNo)
            strChkRefNo = "2"
        ElseIf strRefNo = 4 Then
            '農會聯合中心(吉元)
            strUserArea = GetString(strLimitation, 13) & GetBillNo_Old(strBillNo, strMediaBillNo)
            strBillNoOut = GetBillNo_Old(strBillNo, strMediaBillNo)
            strChkRefNo = "3"
        ElseIf strRefNo = 5 Then
            '土地銀行(吉元)
'           以身份證字號傳入---for Lydia
            'strUserArea = GetString(strInvNo, 10) & Space(6) & GetBillNo_Old(strBillNo, strMediaBillNo)
            strUserArea = GetString(strInvNo, 10) & Space(6) & GetBillNo_Old(strBillNo, strMediaBillNo)
            strBillNoOut = GetBillNo_Old(strBillNo, strMediaBillNo)
            strChkRefNo = "4"
        ElseIf strRefNo = 7 Then
            '群健
            strUserArea = GetString(strCustId, 6, giRight, True) & Space(14) & GetBillNo_New(strBillNo, strMediaBillNo)
            strBillNoOut = GetBillNo_New(strBillNo, strMediaBillNo)
            strChkRefNo = "5"
        ElseIf strRefNo = 8 Then
            '一銀
            strUserArea = GetSpecialBillNo(strBillNo, strRefNo, strID, strAddrNo) & GetBillNo_New(strBillNo, strMediaBillNo) & Space(8) & GetString(strActNo, 14)
            strBillNoOut = GetBillNo_New(strBillNo, strMediaBillNo)
            strChkRefNo = "6"
        ElseIf strRefNo = 12 Then
            '問題集1192:小杜-20041011 Edit:Crystal-20041013
            '台新
            strUserArea = GetSpecialBillNo(strBillNo, strRefNo, strID, strAddrNo) & GetBillNo_New(strBillNo, strMediaBillNo) & GetString(strID, 10)
            strBillNoOut = GetBillNo_New(strBillNo, strMediaBillNo)
            strChkRefNo = "7"
        ElseIf strRefNo = 13 Then
'            strUserArea = GetBillNo_New(strBillNo, strMediabillNo) & Space(9) & GetString(strID, 10) & GetString(rsCD018("BankID2") & "", 14, giLeft)
'            strBillNoOut = GetBillNo_New(strBillNo, strMediabillNo)
            '#3627 格式變動 By Kin 2007/11/21
            strUserArea = GetString(strBillNo, 16, giLeft) & Space(4) & GetString(strID, 10) & GetString(rsCD018("CorpID") & "", 6, giLeft) & _
                            GetString(strCustId, 8, giRight, True)
                            
            '#3910 以strMediabillNo 代入strBillNoOut By Kin 2008/04/29
            strBillNoOut = GetBillNo_New(strBillNo, strMediaBillNo)
            '#3571 增加中信96格式 By Kin 2007/10/22
            strChkRefNo = "8"
        ElseIf strRefNo = "14" Then
            '#5596 測試不OK,土銀的媒體單號要往後再退3碼 By Kin 2010/05/26
            strUserArea = GetString(strInvNo, 10) & Space(6) & GetString("", 4) & GetBillNo_Old(strBillNo, strMediaBillNo)
            strBillNoOut = GetBillNo_Old(strBillNo, strMediaBillNo)
            strChkRefNo = "4"
        Else
            '其他家的金資格式
            strUserArea = GetBillNo_New(strBillNo, strMediaBillNo)
            strBillNoOut = GetBillNo_New(strBillNo, strMediaBillNo)
            If strRefNo <> 9 Then
                strChkRefNo = "9"
            Else
                strChkRefNo = "13"
            End If
        End If
        '#5596土地銀行格式改變
        If Val(strRefNo) = 14 Then
            GetUserAreaStr = strUserArea & GetString(" ", 90 - Len(strUserArea))
        Else
            GetUserAreaStr = strUserArea & GetString(" ", 44 - Len(strUserArea))
        End If
End Function

Private Function GetRealDateTran(strRealDate As String) As String
    On Error Resume Next
        GetRealDateTran = Format(Val(strRealDate) + 19110000, "####/##/##")
End Function

Private Function GetChkCode(strBillNo As String) As String
  On Error GoTo chkErr
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
chkErr:
    Call ErrSub(FormName, "GetChkCode")
End Function

Public Function BatchUpdateCitiBankATM(ByVal strChoose33 As String, _
                                        Optional ByVal strTablName As String = "", _
                                        Optional cn As ADODB.Connection) As Boolean
  On Error GoTo chkErr
    'Dim rsChkAtm As New ADODB.Recordset
    Dim strCorpID As String, strSFString As String
    Dim strBankId As String, strBillNo As String
    Dim strTmpViewName As String, strViewSQL As String
    Dim strSQL As String, strField As String
    Dim lngCount As Long, lngMedia As Long
    strBankId = Empty
    lngMedia = 0
    strField = Empty
    strBankId = strBankHand
    If strTablName <> "" And InStr(1, strTablName, "SO033") = 0 Then
        strTablName = strOwnerName & "SO033 A" & strTablName
    End If
    strSQL = "Select nvl(Para24,0) From " & strOwnerName & "SO043 Where Rownum=1"
    lngMedia = Val(cn.Execute(strSQL)(0))
    If lngMedia = 1 Then
        strField = "MediaBillNo"
    Else
        strField = "BillNo"
    End If
    Dim rsChkATM As New ADODB.Recordset
    If Not GetRS(rsChkATM, "SELECT CheckATM,FileName2,CorpID From " & strOwnerName & "CD018 Where CodeNo=" & strBankId, cn) Then GoTo chkErr
    'If rsChkAtm("CheckATM") = 1 Then
        '***************************************************************************************
        '#3276 取用虛擬帳號改用此方法 By Kin 2008/04/02
        'strSFString = GetBankATM("20" & strBillNo, -1, strBankId, "", 0)
'        GetBankATM "20" & strBillNo, -1, strBankId, "", 0, , strCorpID, strSFString, cn
        '#3276 測試不OK 多增加此Function判斷後端是否存在 By Kin 2008/04/14
        If Not ChkHaveSFObj(rsChkATM("FileName2") & "") Then Exit Function
        strSFString = "SF_" & rsChkATM("FileName2") & ""
        strCorpID = rsChkATM("CorpID") & ""
        '***************************************************************************************
        '取得View暫存檔名稱
        strTmpViewName = GetTmpViewName(cn)
        If strTmpViewName = "" Then Exit Function
        If lngMedia = 1 Then
            strViewSQL = "Create View " & strOwnerName & strTmpViewName & " as " & _
                         " Select MediaBillNo BillNo," & strOwnerName & strSFString & _
                         "('" & strCorpID & "','20'||A.MediaBillNo,Substr(A.MediaBillNo,5)," & _
                         "Sum(A.Shouldamt)) CitiBankATM From " & strTablName & IIf(strChoose33 <> "", " Where " & strChoose33, "") & _
                         " Group By MediaBillNo"
        Else
            strViewSQL = "Create View " & strOwnerName & strTmpViewName & " as " & _
                        " Select BillNo," & strOwnerName & strSFString & "('" & strCorpID & _
                        "',A.BillNo,Substr(A.BillNo,9),Sum(A.Shouldamt)) CitiBankATM From " & _
                        strTablName & IIf(strChoose33 <> "", " Where " & strChoose33, "") & _
                         " Group By BillNo"
        End If
        If Not ExecuteCommand(strViewSQL, cn, lngCount) Then Exit Function
            If cn.Execute("Select Count(*) CNT From " & strOwnerName & strTmpViewName)(0) <= 0 Then BatchUpdateCitiBankATM = True: Exit Function
            cn.BeginTrans
            If Not ExecuteCommand("Update " & strOwnerName & "SO033 A Set CitiBankATM = " & _
                                  "(Select CitiBankATM From " & strOwnerName & strTmpViewName & _
                                  " B Where A." & strField & "= B.BillNo) Where A." & strField & " In (" & _
                                  "Select A." & strField & " From " & strTablName & " Where " & strChoose33 & ")", cn, lngCount) Then cn.RollbackTrans: GoTo chkErr
                                  
            cn.CommitTrans
        On Error Resume Next
        cn.Execute "Drop View " & strOwnerName & strTmpViewName
    'Else
        BatchUpdateCitiBankATM = True
    'End If
    On Error Resume Next
    Call CloseRecordset(rsChkATM)
    Exit Function
chkErr:
    ErrSub FormName, "BatchUpdateCitiBankATM"
End Function
'#3583 取得View的暫存檔名稱 By Kin 2007/10/26
Public Function GetTmpViewName(cn As ADODB.Connection) As String
  On Error Resume Next
    Dim strTmpViewName As String
    Dim strGetOwner As String
    strGetOwner = strOwnerName
    strTmpViewName = "TMP_" & GetRsValue("SELECT trim(To_Char(" & strGetOwner & "S_TMPRPT_ViewName.NextVal, '0999999')) FROM Dual", cn)
    cn.Execute "Drop Table " & strTmpViewName
    cn.Execute "Drop View " & strTmpViewName
    GetTmpViewName = strTmpViewName
End Function
Public Function IsPayKindOK(ByRef rsSource As ADODB.Recordset, _
    ByVal strRealStopDate As String, ByVal aChkPayKind As ChkPayKind, _
    Optional ByVal aMedia As Integer = 1) As Boolean
  On Error GoTo chkErr
    Dim aQry As String
    Dim aPayKind As Integer
    Dim aRsTmp As New ADODB.Recordset
    Dim aBillField As String
    Select Case aMedia
        Case 0
            aBillField = " BILLNO "
        Case 1
            aBillField = " MEDIABILLNO "
        Case Else
            aBillField = " CitibankATM "
    End Select
    If Not GetPaynowFlag Then
        IsPayKindOK = True
        GoTo lEnd
    End If
    Select Case aChkPayKind
        Case giSingle
            aQry = "SELECT MAX(NVL(A.PAYKIND,0)) PAYKIND, " & _
                        strMinRealStopDate & _
                        " FROM " & GetOwner & "SO033 A WHERE " & _
                        aBillField & "='" & rsSource("BILLNO") & "'"
            Set aRsTmp = gcnGi.Execute(aQry)
            aPayKind = Val(aRsTmp("PAYKIND") & "")
            If aPayKind = 0 Then
                IsPayKindOK = True
                GoTo lEnd
            Else
                If Val(Format(aRsTmp("REALSTOPDATE") & "", "YYYYMMDD")) < Val(strRealStopDate & "") Then
                    IsPayKindOK = False
                Else
                    IsPayKindOK = True
                End If
            End If
        Case giAll
            If Val(rsSource("PayKind") & "") = 0 Then
                IsPayKindOK = True
                GoTo lEnd
            End If
            If Val(Format(rsSource("REALSTOPDATE") & "", "YYYYMMDD")) < Val(strRealStopDate & "") Then
                IsPayKindOK = False
            Else
                IsPayKindOK = True
            End If
    End Select
lEnd:
    On Error Resume Next
    Call CloseRecordset(aRsTmp)
    Exit Function
chkErr:
  Call ErrSub("FormName", "IsPayKindOK")
End Function
Public Function GetPaynowFlag() As Boolean
  On Error Resume Next
     GetPaynowFlag = Val(GetRsValue("SELECT PaynowFlag FROM " & GetOwner & "SO041") & "") = 1
End Function

''#3583增加strSFString(後端名稱)與Cn(Connect) By Kin 2007/10/26
'Public Function GetBankATM(strBillNo As String, lngCustId As Long, strBankCode As String, _
'    strShouldDate As String, lngShouldAmt As Long, Optional ByRef clsBankATMNo As Object, _
'    Optional ByRef strCorpID As String, Optional ByRef strSFString As String, _
'    Optional ByRef cn As ADODB.Connection) As String
'    On Error Resume Next
'    Dim rsTmp As New ADODB.Recordset
'    Dim strShDate As String
'        If clsBankATMNo Is Nothing Then
'
'            If strBankCode = "" Or strBillNo = "" Or lngCustId = 0 Then Exit Function
'            If Not GetRS(rsTmp, "Select FileName2,CorpID From " & strOwnerName & "CD018 Where CodeNo = " & strBankCode, cn) Then Exit Function
'            If Len(rsTmp("FileName2") & "") = 0 Or Len(rsTmp("CorpId") & "") = 0 Then MsgBox "銀行代碼檔中無定義轉帳輸入檔名或收件單位代號", vbCritical, "警告": Exit Function
'            strCorpID = rsTmp("CorpId") & ""
'            Set clsBankATMNo = CreateObject("csBankATMNo.cls" & rsTmp("FileName2"))
'            If Err.Number = 429 Then MsgBox "銀行代碼檔中定義之轉帳輸入檔名 : [" & rsTmp("FileName2") & "] 不正確,請確認!!", vbCritical, "警告": Exit Function
'            strSFString = "SF_" & rsTmp("FileName2")
'            If lngCustId = -1 Then Exit Function
'        End If
'        strShDate = Format(strShouldDate, "yyyy/mm/dd")
'        GetBankATM = clsBankATMNo.GetATMNo(strBillNo, lngCustId, strCorpID, strShDate, lngShouldAmt)
'End Function


