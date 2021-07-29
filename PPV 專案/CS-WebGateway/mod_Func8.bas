Attribute VB_Name = "mod_Func8"
Option Explicit

Private cn As Object
Private rs As Object
Private rs105E As Object
Private rs105F As Object
Private rs033 As Object
Private strOwner As String
Private strCardBillNo As String
Private strOrderNo As String
Private strAuthenticID As String
Private strAuthenticCode As String
Private strName As String
Private strCardExpDate As String
Private lngCustID As Long
Private strUpdTime As String

Public Function JustDoIt(ByRef objWebPool As Object, _
                                        ByRef strPara As String) As String
  On Error GoTo ChkErr
  
    Dim objPayment As Object
    
    Dim varPara As Variant, varDetail As Variant
    
    Dim strServiceType As String
    Dim strBillData As String
    
    Dim lngDepositFlag As Long
    Dim lngCardCode As Long
    Dim sglTotalAmt As Single
    Dim strAccountNo As String
    
    Dim strCMCode As String, strCMName As String
    Dim strPTCode As String, strPTName As String
    Dim lngReturn As Long
    
    Dim strCitemName As String
    Dim lngAmount As Long
    Dim intCreditPoint As Integer
    Dim strBillNo As String
    Dim intLoop As Integer
    
    Set cn = Get_General_CN
    
    If cn.State <= 0 Then
        Set cn = Get_General_CN
        If cn.State <= 0 Then
            strErr = "無法連線資料庫!"
            JustDoIt = "-99,資料庫連線失敗!"
            GoTo 66
        End If
    End If
    
'    8、會員代號、認證ID、卡號、信用卡有效期限、信用卡背面授權碼、信用卡別(折行)

    varPara = Split(strPara, vbCrLf)
    '*************************************************************************
    '#3018 判斷傳入的個數是否正確 By Kin 2007/08/31
    If UBound(varPara) < 2 Then
        JustDoIt = "-99,參數資料不正確 ! 請確認 !"
        GoTo 66
    End If
    
    varDetail = Split(varPara(0), ",")
    If UBound(varDetail) < 6 Then
        JustDoIt = "-99,參數不足!"
        GoTo 66
    End If
'    For intLoop = 1 To UBound(varPara)
'        If varPara(intLoop) = Empty Then
'            JustDoIt = "-99,參數不足!"
'            GoTo 66
'        End If
'    Next intLoop
    '*************************************************************************
    strAuthenticID = varDetail(2)
    bytComp = Val(Left(strAuthenticID, 3))
'    lngCustID = Val(Right(strAuthenticID, 7))
    strAccountNo = varDetail(3) '   strAccountNo 信用卡卡號(WEB傳入)
    strCardExpDate = varDetail(4) & "" '   strCardExpDate 信用卡有效年月(WEB傳入)
    lngCardCode = Val(varDetail(6) & "") '   lngCardCode 信用卡別(WEB傳入)
    strAuthenticCode = varDetail(5)
    strOwner = GetOwner(cn)
    
'   Select count(*) as Cnt from so002b where MemberId=[會員代號] and AuthenticId=[認證ID]
    Set rs = cn.Execute("SELECT COMPCODE,CUSTID,NAME FROM " & strOwner & "SO002B" & Get_DB_Link(cn) & _
                                    " WHERE MEMBERID='" & varDetail(1) & "' AND AUTHENTICID='" & strAuthenticID & "'")
    
    If rs.AbsolutePage <= 0 Then '    If Cnt = 0 Then
        JustDoIt = "-2,訂購PPV認證錯誤" '        回傳(-2,訂購PPV認證錯誤)
        GoTo 66
    End If '    End If
    
'   strCardBillNo   信用卡訂單編號(由API以Sequence Object產生)
    strCardBillNo = GetValue(cn, "SELECT " & strOwner & "S_SO033_CARDBILLNO.NEXTVAL FROM DUAL") & ""
    strCardBillNo = Trim(strCardBillNo)
    
    If strCardBillNo = Empty Then
        JustDoIt = "-99,[S_SO033_CARDBILLNO] Sequence 循序物件不存在 !"
        GoTo 66
    End If
    
    strCardBillNo = "2" & Right("00000000" & strCardBillNo, 8) '    產生一信用卡訂單單號(9碼).
    '   原則 :
    '       第1碼為來源識別碼. [2]表WEB呼叫;
    '       第2碼開始為流水號. 呼叫 Sequence Object S_SO033_CardBillNo 取得最新號. 8碼右靠左補零.
    '       Ex : 來源識別碼 + 流水號 → 200000001
    
    bytComp = Val(rs("COMPCODE") & "")
    lngCustID = Val(rs("CUSTID") & "")
    strName = rs("NAME") & ""
    
'   訂單單號產生: 西元年月+ Sequence Object S_SO105E_OrderNo Ex: '2005030000001'
'    strOrderNo = GetValue(cn, "SELECT " & strOwner & "S_SO105E_ORDERNO.NEXTVAL FROM DUAL") & ""
'    strOrderNo = Trim(strOrderNo)
'
'    If strOrderNo = Empty Then
'        JustDoIt = "-99,[S_SO105E_ORDERNO] Sequence 循序物件不存在 !"
'        GoTo 66
'    End If
'    strOrderNo = Format(Date, "YYYYMM") & Right("0000000" & strOrderNo, 7)
    
    '   取訂單編號
    If Not GetOrderNo(cn, strOwner, strOrderNo, JustDoIt) Then GoTo 66

    sglTotalAmt = 0
    Get_Def_CM cn, strOwner, strCMCode, strCMName '    strCMcode = "預設收費方式代碼" , strCMname = "預設收費方式名稱"
    Get_Def_PT cn, strOwner, strPTCode, strPTName '    strPTcode = "預設付款種類代碼" , strPTname = "預設付款種類名稱"
    strUpdTime = "" ' 異動時間 Null
   
    cn.BeginTrans
    
    strUpdTime = Format(RightNow(cn), "YYYY/MM/DD HH:MM:SS")
    
    If Not Ins_2_105E Then
        JustDoIt = "-99,[ PPV 訂單資料主檔 ] 新增失敗 !"
        GoTo RollBack
    End If
    
'    Loop處理 (第二行參數開始)
'        新增訂單單身於SO105F中(訂單單號, 收費項目代號, 收費項目名稱, 金額, iPPV/PPV點數, 業者服務類別)
'        新增資料至SO033中.以T單方式產生並填入信用卡訂單單號
'    End Loop

    '    收費項目代號 (折行)
    '    收費項目代號 (折行)
    '    ...
    
    '   取單據編號
    strBillNo = GetBillNo(cn, strOwner, JustDoIt)
    
    If strBillNo = Empty Then GoTo RollBack
    
    If GetRS(cn, rs105F, "SELECT ORDERNO,CITEMCODE,CITEMNAME,AMOUNT" & _
                                        ",CREDITPOINT,SERVICETYPE,PREPAYFLAG,BILLNO,ITEM" & _
                                        " FROM " & strOwner & "SO105F" & Get_DB_Link(cn) & _
                                        " WHERE ROWID=''") Then
                                        '" WHERE BILLNO='' AND ORDERNO=''"
    
        For intLoop = 1 To UBound(varPara)
            strBillData = Trim(varPara(intLoop) & "")
            If strBillData <> Empty Then
                If GetChargeInfo(strBillData, strCitemName, lngAmount, intCreditPoint) Then
                    With rs105F
                            .AddNew
                            .Fields("OrderNo") = strOrderNo
                            .Fields("CitemCode") = strBillData
                            .Fields("CitemName") = strCitemName
                            .Fields("Amount") = lngAmount
                            .Fields("CreditPoint") = intCreditPoint
                            .Fields("ServiceType") = "C"
                            .Fields("BillNo") = strBillNo
                            .Fields("Item") = .RecordCount
                            .Fields("PrePayFlag") = 1
                            .Update
                            If Not Ins_2_033(strBillData, strBillNo, .Fields("Item"), lngAmount, strCardBillNo, _
                                                        strPTCode, strPTName, strCMCode, strCMName) Then
                                If strErr <> Empty Then
                                    JustDoIt = "-99," & strErr
                                Else
                                    JustDoIt = "-99,[ 正式應收資料檔 ] 新增失敗 !"
                                End If
                                GoTo RollBack
                            End If
                    End With
                Else
                    JustDoIt = "-99,[ 收費項目代碼檔 ] 對應不到代碼 [" & strBillData & "]"
                    GoTo RollBack
                End If
            End If
        Next
    
    Else
        JustDoIt = "-99,[ iPPV/PPV點數訂購檔 ] 開啟失敗 !"
        GoTo RollBack
    End If
    
'   依傳來參數呼叫Payment Gateway 模組之取授權(Approve)並請款(Deposit)之二合一API
    
'   PS: 總計金額請由Web端自行加總處理
    
    lngDepositFlag = GetValue(cn, "SELECT PARA32 FROM " & strOwner & "SO043" & Get_DB_Link(cn) & _
                                                        " WHERE COMPCODE=" & bytComp & _
                                                        " AND SERVICETYPE='C'")
    
    Set objPayment = CreateObject("csPaymentClass.Approve")
    
    With objPayment '   呼叫Payment Gateway 模組之取授權(Approve)並請款(Deposit)之二合一API
            Set .objConn = cn
            .uOwnerName = strOwner
            lngReturn = .Approve(Val(bytComp), strCardBillNo, lngCardCode, strAccountNo, _
                                    strCardExpDate, Val(sglTotalAmt), "", strCMCode, _
                                    strCMName, strPTCode, strPTName, "WEB", "網站", _
                                    "", "", "WEB", lngDepositFlag, strAuthenticCode)
            If lngReturn = 0 Then
                cn.CommitTrans
                JustDoIt = "0,成功," & strCardBillNo '       回傳(0,成功,信用卡訂單編號)
                GoTo 66
            Else '    If 授權結果不ok Then
                JustDoIt = "-7,信用授權有誤" '       回傳(-7,信用授權有誤)
                JustDoIt = JustDoIt & " " & CStr(lngReturn) & " " & .getErrorCode & " " & .getErrorMsg
                GoTo RollBack
            End If
    End With

'    結果碼、結果訊息字串(折行)
'    信用卡訂單編號

RollBack:
    cn.RollbackTrans
  
66:
    On Error Resume Next
    Rlx rs
    Rlx rs105E
    Rlx rs105F
    Rlx rs033
    Rlx cn
    Rlx strName
    Rlx objPayment
    Rlx varPara
    Rlx varDetail
    Rlx strAuthenticID
    Rlx strAuthenticCode
    Rlx strServiceType
    Rlx strBillData
    Rlx lngDepositFlag
    Rlx lngCardCode
    Rlx sglTotalAmt
    Rlx strCardBillNo
    Rlx strOrderNo
    Rlx strCardExpDate
    Rlx strAccountNo
    Rlx strCMCode
    Rlx strCMName
    Rlx strPTCode
    Rlx strPTName
    Rlx strUpdTime
    Rlx intLoop
    Rlx lngCustID
    Rlx strOwner
    Rlx strCitemName
    Rlx lngAmount
    Rlx intCreditPoint
    Rlx strBillNo
    
    Set cn = Nothing
    Set rs = Nothing
    Set rs105E = Nothing
    Set rs105F = Nothing
    Set rs033 = Nothing
  Exit Function
ChkErr:
    ErrHandle "mod_Func8", "JustDoIt"
    cn.RollbackTrans
    JustDoIt = "-99," & strErr
End Function

Private Function GetChargeInfo(strCitemCode As String, _
                                                    ByRef strCitemName As String, _
                                                    ByRef lngAmount As Long, _
                                                    ByRef intCreditPoint As Integer) As Boolean
  On Error Resume Next
    If GetRS(cn, rs, "SELECT DESCRIPTION,AMOUNT,CREDITPOINT" & _
                            " FROM " & strOwner & "CD019" & Get_DB_Link(cn) & _
                            " WHERE CODENO=" & strCitemCode) Then
        strCitemName = rs("DESCRIPTION") & ""
        lngAmount = Val(rs("AMOUNT") & "")
        intCreditPoint = Val(rs("CREDITPOINT") & "")
        GetChargeInfo = True
    Else
        strCitemName = ""
        lngAmount = ""
        intCreditPoint = ""
    End If
End Function

Private Function Ins_2_105E() As Boolean
  On Error GoTo ChkErr
'    新增一筆訂單單頭於SO105E中
'    (訂單單號, 公司別, 客戶編號, 客戶名稱, PPV訂購方式代碼, PPV訂購方式, 認證編號, 姓名, 卡號,
'        信用卡有效期限, 受理時間, 受理人員代號, 受理人員名稱, 業者服務類別, 信用卡訂單編號)
    Dim strCustName As String
    strCustName = RPxx(cn.Execute("SELECT CUSTNAME FROM " & strOwner & "SO001" & Get_DB_Link(cn) & _
                                                        " WHERE CUSTID=" & lngCustID & _
                                                        " AND COMPCODE=" & bytComp).GetString(2, 1, "", "", "") & "")
    
    If GetRS(cn, rs105E, "SELECT ORDERNO,COMPCODE,CUSTID,CUSTNAME," & _
                                        "PPVORDERMODENO,PPVORDERMODE,AUTHENTICID," & _
                                        "NAME,ACCOUNTNO,CARDEXPDATE,ACCEPTTIME,ACCEPTEN,ACCEPTNAME," & _
                                        "SERVICETYPE,CARDBILLNO,UPDTIME,UPDEN" & _
                                        " FROM " & strOwner & "SO105E" & Get_DB_Link(cn) & _
                                        " WHERE ROWID=''") Then
                                        '" WHERE COMPCODE=" & bytComp & " AND ORDERNO=''"
        With rs105E
                .AddNew
                .Fields("ORDERNO") = strOrderNo
                .Fields("COMPCODE") = Val(bytComp)
                .Fields("CUSTID") = lngCustID
                .Fields("CUSTNAME") = strCustName
                .Fields("PPVORDERMODENO") = 2
                .Fields("PPVORDERMODE") = "網站"
                .Fields("AUTHENTICID") = strAuthenticID
                .Fields("NAME") = strName
                .Fields("CARDEXPDATE") = Val(strCardExpDate)
                .Fields("ACCEPTTIME") = strUpdTime
                .Fields("ACCEPTEN") = "WEB"
                .Fields("UPDTIME") = strUpdTime
                .Fields("UPDEN") = "WEB"
                .Fields("ACCEPTNAME") = "網站"
                .Fields("SERVICETYPE") = "C"
                .Fields("CARDBILLNO") = strCardBillNo
                .Update
        End With
    End If
    Ins_2_105E = True
    On Error Resume Next
    Rlx strCustName
  Exit Function
ChkErr:
    ErrHandle "mod_Func8", "Ins_2_105E"
    Ins_2_105E = False
End Function

Private Function Ins_2_033(ByVal strCitemCode As String, ByVal strBillNo As String, ByVal lngItem As Long, _
                                            ByVal lngAmount As Long, ByVal strCardBillNo As String, _
                                            ByVal strPTCode As String, ByVal strPTName As String, _
                                            ByVal strCMCode As String, ByVal strCMName As String) As Boolean
  On Error GoTo ChkErr
    Ins_2_033 = False
    If Not GetRS(cn, rs033, "SELECT * FROM " & strOwner & "SO033" & Get_DB_Link(cn) & " WHERE ROWID =''") Then
        'WHERE BILLNO ='' AND ITEM = -1
        strErr = "-99,[ 正式應收資料檔 ] 開啟失敗 ! GetRS(rs033) "
        Exit Function
    End If
    With rs033
            Set .ActiveConnection = Nothing
'            strErr = strErr & lngCustID & "," & bytComp & "," & strBillNo & "," & lngItem & strCitemCode
            If Ins_Chg_Fld(Val(lngCustID), Val(bytComp), strBillNo, lngItem, strCitemCode, Format(strUpdTime, "YYYY/MM/DD")) Then
                .Fields("RealDate") = Format(strUpdTime, "YYYY/MM/DD")
                If strCMCode & "" <> "" Then .Fields("CMCode") = strCMCode
                If strCMName & "" <> "" Then .Fields("CMName") = strCMName
                If strPTCode & "" <> "" Then .Fields("PTCode") = strPTCode
                If strPTName & "" <> "" Then .Fields("PTName") = strPTName
'                .Fields("ShouldAmt") = lngAmount
                .Fields("OrderNo") = strOrderNo
                .Fields("CardBillNo") = NoZero(strCardBillNo)
                .Fields("AuthenticId") = NoZero(strAuthenticID)
                .Fields("ServiceType") = "C"
                .Fields("CreateTime") = strUpdTime
                .Fields("UpdTime") = strUpdTime
                .Fields("UpdEn") = "WEB"
                 Ins_2_033 = InsertToOracle(cn, strOwner & "SO033", rs033, " WHERE BILLNO = '' AND ITEM = -1")
                If Not Ins_2_033 Then strErr = "-99,[ 正式應收資料檔 ] 開啟失敗 ! InsertToOracle(rs033) " & strErr
            Else
                strErr = "-99,填資料至 [ 資料錄SO033 ] ( Ins_Chg_Fld ) 失敗 !" & strErr
            End If
            .CancelUpdate
    End With
  Exit Function
ChkErr:
    ErrHandle "mod_Func8", "Ins_2_033"
End Function

Private Function Ins_Chg_Fld(ByVal lngCustID As Long, _
                                                ByVal lngCompCode As Long, _
                                                ByVal strBillNo As String, _
                                                ByVal lngItem As Long, _
                                                ByVal strCitemCode As String, _
                                                ByVal strShouldDate As String) As Boolean
  On Error GoTo ChkErr
    Dim objINC As Object
'    Dim objINC As New csAlterCharge4.clsInsertNewCharge
    Set objINC = CreateObject("csAlterCharge4.clsInsertNewCharge")
    With objINC
        .uOwnerName = strOwner
        .uErrPath = Environ("Temp")
        Set .uConn = cn
        Ins_Chg_Fld = .InsertChargeField(rs033, lngCustID, lngCompCode, "C", _
                                                                strBillNo, lngItem, strCitemCode, strShouldDate, _
                                                                "WEB", "網站", , , True, False)
    End With
    On Error Resume Next
    Set objINC = Nothing
  Exit Function
ChkErr:
    ErrHandle "mod_Func8", "Ins_Chg_Fld"
    Ins_Chg_Fld = False
End Function
'PS:
'   訂單單號產生: 西元年月+ Sequence Object S_SO105E_OrderNo Ex: '2005030000001'
'   授權金額: 由參數傳入之所有收費項目代號讀取cd019.Amount做加總產生
'   收費方式代碼, 收費方式名稱: select CodeNo,Description from cd031 where RefNo=4 and StopFlag=0 and rownum=1
'   付款種類代碼, 付款種類名稱: select CodeNo,Description from cd032 where RefNo=4 and StopFlag=0 and rownum=1
'   收費人員代碼: 'WEB'
'   收費人員名稱: '網站'
'   公司別代碼: so002b.CompCode
'   客戶編號: so002b.CustId
'   客戶名稱: so001.CustName
'   PPV訂購方式代碼: 2
'   PPV訂購方式: 'WEB'
'   姓名: so002b.Name
'   受理時間: 由程式產生
'   受理人員代號: 'WEB'
'   受理人員名稱: '網站'
'   業者服務類別: 'C'
'   信用卡訂單編號產生: 參考註一
'   收費項目名稱:  CD019.Description
'   金額:  CD019.Amount
'   iPPV/PPV點數: CD019.CreditPoint

'csPaymentClass.Approve
'
'Public Function Approve(ByVal lngCompCode As Long, ByVal strCardBillNo As String, _
'        ByVal lngCardCode As Long, ByVal strAccountNo As String, _
'        ByVal strCardExpDate As String, ByVal lngAmount As Long, ByVal strRealDate As String, _
'        ByVal strCMCode As String, ByVal strCMName As String, _
'        ByVal strPTCode As String, ByVal strPTName As String, _
'        ByVal strClctEn As String, ByVal strClctName As String, _
'        ByVal strUpdTime As String, ByVal strUpdEn As String, ByVal strEntryEn As String, _
'        ByVal lngDepositFlag As Long) As Long
'
'   objConn
'   uOwnerName
'
'   GetErrorCode
'   GetErrorMsg
