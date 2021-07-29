Attribute VB_Name = "mod_Func7"
Option Explicit

Private cn As Object
Private rs As Object
Private strOwner As String

Public Function JustDoIt(ByRef objWebPool As Object, _
                                        ByRef strPara As String) As String
  On Error GoTo ChkErr
  
    Dim objPayment As Object
    
    Dim varPara As Variant
    Dim varDetail As Variant
    
    Dim strAuthenticID As String
    Dim strAuthenticCode As String
    Dim strServiceType As String
    Dim strBillData As String
'    Dim lngCustID As Long
    
    Dim lngDepositFlag As Long
    Dim lngCardCode As Long
    Dim sglTotalAmt As Single
    Dim strCardBillNo As String
    Dim strCardExpDate As String
    Dim strAccountNo As String
    
    Dim strRealDate As String
    Dim strCMCode As String
    Dim strCMName As String
    Dim strPTCode As String
    Dim strPTName As String
    Dim strClctEn As String
    Dim strClctName As String
    Dim strUpdTime As String
    Dim strUpdEn As String
    Dim strEntryEn As String
    Dim lngReturn As Long
   
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
    
'   7、會員代號、認證ID、卡號、信用卡有效期限、信用卡背面授權碼、信用卡別(折行)
'   單據編號、項次(折行)
'   單據編號、項次(折行)
'   ….

    varPara = Split(strPara, vbCrLf)
    '***********************************************************************************
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
    For intLoop = 1 To UBound(varPara)
        If (UBound(Split(varPara(intLoop), ",")) <> 1) And (varPara(intLoop) <> Empty) Then
            JustDoIt = "-99,參數不足!"
            GoTo 66
        End If
    Next intLoop
    '***********************************************************************************
    strAuthenticID = varDetail(2)
    bytComp = Val(Left(strAuthenticID, 3)) '   公司別代碼: SO033.CompCode
'    lngCustID = Val(Right(strAuthenticID, 7))
    
    strOwner = GetOwner(cn)

'   Select count(*) as Cnt from so002b where MemberId=[會員代號] and AuthenticId=[認證ID]
    Set rs = cn.Execute("SELECT COUNT(*) AS CNT FROM " & strOwner & "SO002B" & Get_DB_Link(cn) & _
                                    "WHERE MEMBERID='" & varDetail(1) & "' AND AUTHENTICID='" & strAuthenticID & "'")
    
    If rs("CNT") <= 0 Then '    If Cnt = 0 Then
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
    
    strCardBillNo = "2" & Right("00000000" & strCardBillNo, 8)
    '   產生信用卡訂單單號(9碼). 碼原則 :
    '       第1碼為來源識別碼. [2]表WEB呼叫;
    '       第2碼開始為流水號. 呼叫 Sequence Object S_SO033_CardBillNo 取得最新號. 8碼右靠左補零.
    '       Ex : 來源識別碼 + 流水號 → 200000001
    
    strAccountNo = varDetail(3) '   strAccountNo 信用卡卡號(WEB傳入)
    strCardExpDate = varDetail(4) & "" '   strCardExpDate 信用卡有效年月(WEB傳入)
    strAuthenticCode = varDetail(5)
    lngCardCode = Val(varDetail(6) & "") '   lngCardCode 信用卡別(WEB傳入)
'   lngDepositFlag = -1
    sglTotalAmt = 0
    
'   依傳來參數呼叫Payment Gateway 模組之取授權(Approve)並請款(Deposit)之二合一API

    cn.BeginTrans
    
    For intLoop = 1 To UBound(varPara)
        
        strBillData = varPara(intLoop) & ""
        
        If strBillData <> Empty Then
            varDetail = Split(strBillData, ",")
        '    單據編號、項次(折行)
        '    單據編號、項次(折行)
        '    ….
            If GetRS(cn, rs, "SELECT CARDBILLNO,SHOULDAMT,SERVICETYPE FROM " & _
                                        strOwner & "SO033" & Get_DB_Link(cn) & _
                                        " WHERE BILLNO='" & Trim(varDetail(0)) & "'" & _
                                        " AND ITEM=" & Trim(varDetail(1))) Then
                                            
                If rs.AbsolutePosition > 0 Then
                    rs!CardBillNo = strCardBillNo
                    rs.Update
                    sglTotalAmt = sglTotalAmt + Val(rs!ShouldAmt & "") '   授權金額: 由(單據編號、項次) 讀取 SO033 做加總產生
                    If strServiceType = Empty Then strServiceType = rs!ServiceType & "" '   lngAmount 授權金額(由API讀取資料加總)
                End If
                
            End If
            
        End If
        
    Next
    
    '   PS: 總計金額請由Web端自行加總處理
    
    lngDepositFlag = Val(GetValue(cn, "SELECT PARA32 FROM " & strOwner & "SO043" & Get_DB_Link(cn) & _
                                                            " WHERE COMPCODE=" & bytComp & _
                                                            " AND SERVICETYPE='" & strServiceType & "'") & "")
    
    Set objPayment = CreateObject("csPaymentClass.Approve")
    
    Get_Def_CM cn, strOwner, strCMCode, strCMName '    strCMcode = "預設收費方式代碼" , strCMname = "預設收費方式名稱"
    Get_Def_PT cn, strOwner, strPTCode, strPTName '    strPTcode = "預設付款種類代碼" , strPTname = "預設付款種類名稱"
    strRealDate = "" ' 收款日期(空白以系統日期為主) Null
    strClctEn = "WEB" '   收費人員代碼: 'WEB'
    strClctName = "網站" '   收費人員名稱: '網站'
    strUpdTime = "" ' 異動時間 Null
    strUpdEn = "" ' 異動人員 Null
    strEntryEn = "WEB" ' 操作者代號填入 'WEB'
    '#3599 多增加一個 strAuthenticID參數
    With objPayment '   呼叫Payment Gateway 模組之取授權(Approve)並請款(Deposit)之二合一API
            Set .objConn = cn
            .uOwnerName = strOwner
            lngReturn = .Approve(Val(bytComp), strCardBillNo, lngCardCode, strAccountNo, _
                                    strCardExpDate, Val(sglTotalAmt), strRealDate, strCMCode, _
                                    strCMName, strPTCode, strPTName, strClctEn, strClctName, _
                                    strUpdTime, strUpdEn, strEntryEn, lngDepositFlag, strAuthenticCode)
            
            If lngReturn = 0 Then
                cn.CommitTrans
                JustDoIt = "0,成功," & strCardBillNo '       回傳(0,成功,信用卡訂單編號)
            Else '    If 授權結果不ok Then
                JustDoIt = "-7,信用授權有誤" '       回傳(-7,信用授權有誤)
                cn.RollbackTrans
                JustDoIt = JustDoIt & " " & CStr(lngReturn) & " " & .getErrorCode & " " & .getErrorMsg
            End If
    End With
    
'    結果碼、結果訊息字串(折行)
'    信用卡訂單編號
  
66:
    On Error Resume Next
    Rlx rs
    Rlx cn
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
    Rlx strCardExpDate
    Rlx strAccountNo
    Rlx strRealDate
    Rlx strCMCode
    Rlx strCMName
    Rlx strPTCode
    Rlx strPTName
    Rlx strClctEn
    Rlx strClctName
    Rlx strUpdTime
    Rlx strUpdEn
    Rlx strEntryEn
    Rlx intLoop
    Rlx strOwner
'    Rlx lngCustID
    Set cn = Nothing
  Exit Function
ChkErr:
    ErrHandle "mod_Func7", "JustDoIt"
    cn.RollbackTrans
    JustDoIt = "-99," & strErr
End Function

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

