Attribute VB_Name = "mod_Func2"
Option Explicit

Private strAuthenticID As String
'Private objCnPool As Object
Private cn As Object

Public Function JustDoIt(ByRef objWebPool As Object, _
                                        ByRef strPara As String) As String
  On Error GoTo ChkErr
    Dim strQry As String
    Dim varPara As Variant
    Dim rs As Object
    Dim lngCustID As Long
    Dim objCnPool As Object
    Dim strOwner As String
    Dim strAccBefore As String
    Dim strAccBalance As String
    Dim strSubQry As String
    
    Set objCnPool = objWebPool
    Set cn = objCnPool.GetConnection
    
    If cn.State <= 0 Then
        Set cn = objCnPool.GetConnection
        If cn.State <= 0 Then
            strErr = "無法連線資料庫!"
            JustDoIt = "-99,資料庫連線失敗!"
            GoTo 66
        End If
    End If
    
    '  2、會員代號、認證ID、PPV後付訂購密碼
    
    varPara = Split(strPara, ",")
    '******************************************************************
    '#3018 增加判斷傳入的個數是否正確 By Kin 2007/08/30
    If UBound(varPara) <> 3 Then
        If UBound(varPara) <> 41 Then
            JustDoIt = "-99,參數不足!"
            Exit Function
        End If
    End If
    '******************************************************************
    strAuthenticID = varPara(2)
    bytComp = Val(Left(strAuthenticID, 3))
    
    strOwner = GetOwner(cn)
    
    ' 舊 : Select PPVOrderPwd,CompCode,CustId from so002b where MemberId=[會員代號] and AuthenticId=[認證ID]
    'strQry = "SELECT PPVORDERPWD,COMPCODE,CUSTID FROM " & strOwner & "SO002B" & Get_DB_Link(cn) & _
                    "WHERE MEMBERID='" & varPara(1) & "' AND AUTHENTICID='" & strAuthenticID & "'"
    
    
    ' 舊 : Select PPVOrderPwd,CompCode,CustId,FaciSNo,AccountingBefore, AccountingBalance from so002b
    '       where MemberId=[會員代號] and AuthenticId=[認證ID]
    ' Q : 2302 , Doc : Web API說明_20060413_Liga.doc
    
    ' 新 : Select PPVOrderPwd,CompCode,CustId,FaciSNo,AccountingBefore, AccountingBalance, AuthenticID
    '       from so002b where MemberId=[會員代號] and AuthenticId=[認證ID]   --2006/04/11 Modify
    ' Q : 2378 , Doc : Web API說明_20060428_Liga.doc
    
    strQry = "SELECT PPVORDERPWD,COMPCODE,CUSTID,FACISNO" & _
                    ",ACCOUNTINGBEFORE,ACCOUNTINGBALANCE,AUTHENTICID" & _
                    " FROM " & strOwner & "SO002B" & Get_DB_Link(cn) & _
                    " WHERE MEMBERID='" & varPara(1) & "' AND AUTHENTICID='" & strAuthenticID & "'"

    Set rs = cn.Execute(strQry)
    
    '   If RecordCount=0 then
    '       回傳(-2,訂購PPV認證錯誤)
    '   else
    '       if PPVOrderPwd<>[ PPV後付訂購密碼] then
    '            回傳(-4,未啟用PPV訂購權)
    '   end if
    
    ' QQ : 2577
    ' Doc : Web API說明_20060703_Liga.doc
    ' API-2規格異動
    ' 取消PPV後付訂購密碼檢核，參數保留不動，僅修改檢核語法，取消該段語法即可。
    ' else　if　PPVOrderPwd<>[　PPV後付訂購密碼]　then
    '  回傳(-4,未啟用PPV訂購權)

    ' 調整該API之源由:
    ' 原API-2規格設計，僅FOR意藍訂戶會員使用，即需檢核PPV後付訂購密碼；
    ' 該支API除了用來LOGIN檢核外，亦可做為客戶基本資料更新使用，現況訂戶會員並不定一開始都會有PPV訂購密碼，
    ' 故調整API-2，不作PPV訂購密碼的檢核。
    
    
    If rs.RecordCount <= 0 Then
        JustDoIt = "-2,訂購PPV認證錯誤"
'    ElseIf StrComp(CStr(rs("PPVORDERPWD") & ""), CStr(varPara(3)), vbBinaryCompare) = 0 Then ' By Hammer 2006/07/06
    Else
        bytComp = rs("COMPCODE")
        lngCustID = rs("CUSTID")
        strAccBefore = rs("ACCOUNTINGBEFORE") & ""
        strAccBalance = rs("ACCOUNTINGBALANCE") & ""
        
        ' select count(*) as Cnt from so004 where
        '   CompCode=[so002b.CompCode] and CustId=[so002b.CustId] and
        '   InstDate is not null and prdate is null
        '   PS : 消費上限,欠費金額,信用卡付費上限待欄位定義後加入回傳值.
        
'        strQry = "SELECT COUNT(*) AS RCDCNT FROM " & strOwner & "SO004" & Get_DB_Link(cn) & _
                        "WHERE COMPCODE=" & bytComp & " AND CUSTID=" & lngCustID & _
                        " AND INSTDATE IS NOT NULL AND PRDATE IS NULL"
        
        'Doc : Web API說明_20060413_Liga.doc
        'select SEQNo, FaciSNo, SmartCardNo from so004 where
        '   CompCode=[so002b.CompCode] and CustId=[so002b.CustId] and
        '   SeqNo in [so002b.FaciSNo] and InstDate is not null and prdate is null
        '   and facicode in (select codeno from cd022 where refno=3) ;
        
        strSubQry = "SELECT FACISNO FROM " & strOwner & "SO002B" & Get_DB_Link(cn) & _
                                " WHERE MEMBERID='" & varPara(1) & "' AND AUTHENTICID='" & strAuthenticID & "'" & _
                                " AND FACISNO IS NOT NULL"
                                
        On Error Resume Next
        strSubQry = RPxx(cn.Execute(strSubQry).GetString(2, , "", ",", ""))
        If Err = 3021 Then
            Err.Clear
            strSubQry = ""
        Else
            strSubQry = Left(strSubQry, Len(strSubQry) - 1)
        End If
        
        On Error GoTo ChkErr
        If Len(strSubQry) > 0 Then
            'Doc : Web API說明_20060428_Liga.doc
            'select SEQNo, FaciSNo, SmartCardNo from so004 where
            '   CompCode=[so002b.CompCode] and CustId=[so002b.CustId] and
            '   InstDate is not null and prdate is null
            '   and facicode in (select codeno from cd022 where refno=3)
            '   and SeqNo in [so002b.FaciSNo] ;  --2006/04/11 Modify
            
            '2006/06/05 Modify
            '之前API-2僅回傳該認證ID可使用的DSTB資料，現調整規格，只要CD022.REFNO>0的，皆需回傳XML給意藍
            'QQ:2513 , doc : Web API說明_20060605_Liga.doc
            strQry = "SELECT SEQNO,FACISNO,SMARTCARDNO FROM " & strOwner & "SO004" & Get_DB_Link(cn) & _
                            " WHERE COMPCODE=" & bytComp & " AND CUSTID=" & lngCustID & _
                            " AND SEQNO IN (" & strSubQry & ")" & _
                            " AND INSTDATE IS NOT NULL AND PRDATE IS NULL AND FACICODE IN" & _
                            " (SELECT CODENO FROM " & strOwner & "CD022" & Get_DB_Link(cn) & " WHERE REFNO > 0)"
            
'            strQry = "SELECT SEQNO,FACISNO,SMARTCARDNO FROM " & strOwner & "SO004" & Get_DB_Link(cn) & _
                            " WHERE COMPCODE=" & bytComp & " AND CUSTID=" & lngCustID & _
                            " AND SEQNO IN (" & strSubQry & ")" & _
                            " AND INSTDATE IS NOT NULL AND PRDATE IS NULL AND FACICODE IN" & _
                            " (SELECT CODENO FROM " & strOwner & "CD022" & Get_DB_Link(cn) & " WHERE REFNO=3)"
            
            '   if Cnt=0 then
            '       回傳(-14, 查無安裝STB設備資料)
            '   else
            '       如果參數數量超過4個時(即>4),則表示有客戶認證資訊 [Web額外傳入之資料]
            '       需增加update to so002b命令.
            '       回傳(0,成功,消費上限,欠費金額,信用卡付費上限).
            '   end if
            
            Set rs = cn.Execute(strQry)
'            If rs("RCDCNT") = 0 Then
            If rs.EOF Then
                If UBound(varPara) + 1 > 4 Then
                    If Update_2_SO002B(varPara, strOwner) Then
                        ' Delete from so002blog where MemberId=[會員代號] and AuthenticId=[認證ID]
                        cn.Execute "DELETE FROM " & strOwner & "SO002BLOG" & Get_DB_Link(cn) & _
                                            " WHERE MEMBERID='" & varPara(1) & "' AND" & _
                                            " AUTHENTICID='" & strAuthenticID & "'"
                        JustDoIt = "-14, 查無安裝STB設備資料"
                    Else
                        strErr = "-99," & "新增客戶資料至 [客戶認證資訊檔] 失敗!" & IIf(strErr = Empty, "", strCrLf & strErr)
                    End If
                Else
                    JustDoIt = "-14, 查無安裝STB設備資料"
                End If
            Else
                If UBound(varPara) + 1 > 4 Then
    '                需增加update to so002b命令.
                    If Update_2_SO002B(varPara, strOwner) Then
                        ' Delete from so002blog where MemberId=[會員代號] and AuthenticId=[認證ID]
                        cn.Execute "DELETE FROM " & strOwner & "SO002BLOG" & Get_DB_Link(cn) & _
                                            " WHERE MEMBERID='" & varPara(1) & "' AND" & _
                                            " AUTHENTICID='" & strAuthenticID & "'"
                        ' 回傳(0,成功)及 消費上限,欠費金額,信用卡付費上限 及 可使用DSTB之XML檔
                        JustDoIt = "0,成功" & vbCrLf & _
                                            strAccBefore & "," & strAccBalance & "," & _
                                            "," & GetCG1(strOwner, lngCustID) & _
                                            "," & GetCG2(strOwner) & vbCrLf & _
                                            ExpXML(rs)
                    Else
                        If strErr = Empty Then
                            strErr = "新增客戶資料至 [客戶認證資訊檔] 失敗!"
                        Else
                            strErr = "新增客戶資料至 [客戶認證資訊檔] 失敗!" & strCrLf & strErr
                        End If
                        JustDoIt = "-99," & strErr
                    End If
                Else
                    ' 舊
                    'PS1 : 消費上限,欠費金額,信用卡付費上限，請依該認證ID所select出來的值回傳。
                    '信用卡付費上限 (尚未定義暫傳空值)
                    ' 新
                    '--2006/04/28 Add
                    'PS1:  消費上限 , 欠費金額上限, 信用卡付費上限, 目前欠費金額, 未結算訂單金額
                    '請依該認證ID所select出來的值回傳。信用卡付費上限(尚未定義暫傳空值)
                    
                    JustDoIt = "0,成功" & vbCrLf & _
                                        strAccBefore & "," & strAccBalance & "," & _
                                        "," & GetCG1(strOwner, lngCustID) & _
                                        "," & GetCG2(strOwner) & vbCrLf & _
                                        ExpXML(rs)
                    'PS2 : 該認證ID可使用DSTB之XML(設備流水號、設備序號、智慧卡卡號)
                End If
            End If
        Else
            JustDoIt = "-23,該認證ID查無認證設備或服務"
'            JustDoIt = "-99,[客戶認證資訊檔] 查無 [認證設備序號] 資料!"
'            JustDoIt = "-4,未啟用PPV訂購權"
        End If
'    Else
'        JustDoIt = "-99,[客戶認證資訊檔] 查無 [認證設備序號] 資料!"
'        JustDoIt = "-4,未啟用PPV訂購權"
    End If
    
    
    '   結果碼、結果訊息字串(折行)
    '   消費上限、欠費金額 (該欄待定義)、信用卡付費上限
    
    '   0:   成功
    '   -2 : 訂購PPV認證錯誤
    '   -3 : 此機上盒已停用
    '   -4 : 未啟用PPV訂購權
       
66:
    On Error Resume Next
    Rlx rs
    Rlx strQry
    Rlx varPara
    Rlx strOwner
    Rlx lngCustID
    Rlx strAccBefore
    Rlx strAccBalance
    Rlx strAuthenticID
    Set cn = Nothing
    Set objCnPool = Nothing
  Exit Function
ChkErr:
    ErrHandle "mod_Func2", "JustDoIt"
    JustDoIt = "-99," & strErr
End Function

Private Function GetCG1(strOwner As String, lngCustID As Long) As String
  On Error GoTo ChkErr
    '--2006/04/28 Add  count該認證ID已結算的節目費與點數欠費金額
    'select sum(shouldamt) as charge1 from so033 where compcode=[CustId] and authenticid=[AuthenticID]
    'and citemcode in (select codeno from cd019 where refno in (3,17)) and uccode is not null and cancelflag!=1
    GetCG1 = GetValue(cn, "SELECT SUM(SHOULDAMT) AS CHARGE1 FROM " & strOwner & "SO033" & Get_DB_Link(cn) & _
                                            " WHERE CUSTID=" & lngCustID & " AND AUTHENTICID='" & strAuthenticID & "'" & _
                                            " AND CITEMCODE IN (SELECT CODENO FROM " & strOwner & "CD019" & Get_DB_Link(cn) & _
                                            " WHERE REFNO IN (3,17) AND UCCODE IS NOT NULL AND CANCELFLAG<>1)")
  Exit Function
ChkErr:
    ErrHandle "mod_Func2", "GetCG1"
End Function

Private Function GetCG2(strOwner As String) As String
  On Error GoTo ChkErr
    '--2006/04/28 Add count該認證ID尚未結算的訂單金額
    'select sum(b.amount) as charge2 from so105e a, so105g b Where a.AuthenticID = [AuthenticID]
    'And a.orderno = b.orderno And b.accounting! = 1 And b.canceltime Is Null
    GetCG2 = GetValue(cn, "SELECT SUM(B.AMOUNT) AS CHARGE2 FROM " & _
                                            strOwner & "SO105E" & Get_DB_Link(cn) & " A," & strOwner & "SO105G" & Get_DB_Link(cn) & " B" & _
                                            " WHERE AUTHENTICID='" & strAuthenticID & "' AND A.ORDERNO=B.ORDERNO" & _
                                            " AND B.ACCOUNTING <> 1 AND B.CANCELTIME IS NULL")
  Exit Function
ChkErr:
    ErrHandle "mod_Func2", "GetCG2"
End Function

Private Function Update_2_SO002B(ByRef varPara As Variant, ByRef strOwner As String) As Boolean
  On Error GoTo ChkErr
'    Dim cmd As Object
    Dim intRet As Integer
    '    Set cmd = objCnPool.GetCommand
    '2、會員代號、認證ID、PPV後付訂購密碼
    '會員名稱、身分證字號、出生年月日、聯絡電話、手機號碼、性別、
    '聯絡地址、Email、PPV後付訂購密碼、登入密碼(Web有會員之料被異動時需傳)
    '    With cmd
    '         Set .ActiveConnection = cn
    '        .CommandType = 1
    '        .CommandText = "UPDATE " & strOwner & "SO002B " & _
    '                                    "SET MEMBERID=?,NAME=?,ID=?," & _
    '                                    "BIRTHDAY=TO_DATE(?, 'YYYY/MM/DD')," & _
    '                                    "TEL=?,MOBILE=?,SEXTUAL=?," & _
    '                                    "ADDRESS=?,EMAIL=?," & _
    '                                    "PPVORDERPWD=?,LOGINPASSWORD=? " & _
    '                                    "WHERE ( AUTHENTICID=? )"
    '        .Execute intRet, Array(varPara(1), varPara(4), varPara(5), _
    '                                    varPara(6), varPara(7), varPara(8), varPara(9), _
    '                                    varPara(10), varPara(11), varPara(12), _
    '                                    varPara(13), strAuthenticID), 1
    '        Update_2_SO002B = (intRet > 0)
    '        Set .ActiveConnection = Nothing
    '    End With

    '   新參數 : 客服網站會員代號、會員名稱、身分證字號、出生日期-年、出生日期-月、出生日期-日、
    '   聯絡電話(宅) 、聯絡電話(公)、行動電話、性別、通訊地址-縣市、通訊地址-鄉鎮市區、通訊地址-路段街巷弄巷號樓、
    '   Email、客服網站登入帳號、PPV訂購密碼、客服網站會員密碼、密碼提示問題、密碼提示答案、
    '   學歷、願意收到有線電視服務及頻道節目等相關訊息、婚姻狀況、同住人數、同住成員、職業別、
    '   個人年收入、身份別、帳號狀態、申請日期時間、啟用日期、最後登入日期、最後異動日期、
    '   系統台代碼、發票抬頭、發票種類(二聯或三聯)、發票地址、統一編號(當發票種類為三聯時須為必填) 、設備流水號
    '   (Web有會員之料被異動時需傳)
    Dim strSQL As String
    
    varPara(41) = Replace(varPara(41), "#", ",", 1)
    varPara(41) = Replace(varPara(41), "'", "''", 1)
    strSQL = "UPDATE " & strOwner & "SO002B" & Get_DB_Link(cn) & _
                    "SET MEMBERID=" & PrcType(varPara(4)) & ",NAME=" & PrcType(varPara(5)) & _
                    ",ID=" & PrcType(varPara(6)) & "," & "BIRTHDAYYYYY=" & PrcType(varPara(7)) & _
                    ",BIRTHDAYMM=" & PrcType(varPara(8)) & ",BIRTHDAYDD=" & PrcType(varPara(9)) & _
                    ",TEL=" & PrcType(varPara(10)) & ",TEL2=" & PrcType(varPara(11)) & _
                    ",MOBILE=" & PrcType(varPara(12)) & "," & "SEXTUAL=" & PrcType(varPara(13)) & _
                    ",CITY=" & PrcType(varPara(14)) & ",TOWN=" & PrcType(varPara(15)) & "," & _
                    "ADDRESS=" & PrcType(varPara(16)) & ",EMAIL=" & PrcType(varPara(17)) & _
                    ",LOGINID=" & PrcType(varPara(18)) & "," & "PPVORDERPWD=" & PrcType(varPara(19)) & _
                    ",LOGINPASSWORD=" & PrcType(varPara(20)) & "," & "HINTNAME=" & PrcType(varPara(21)) & _
                    ",HINTANSWER=" & PrcType(varPara(22)) & "," & "EDUCATION=" & PrcType(varPara(23)) & _
                    ",MESSAGE=" & PrcType(varPara(24)) & "," & "MARRIED=" & PrcType(varPara(25)) & _
                    ",BODYCOUNT=" & PrcType(varPara(26), FldNum) & "," & "MEMBERAGE=" & PrcType(varPara(27)) & _
                    ",JOB=" & PrcType(varPara(28)) & "," & "RENTE=" & PrcType(varPara(29)) & _
                    ",CLASSID=" & PrcType(varPara(30)) & ",STATUS=" & PrcType(varPara(31)) & _
                    ",CREATEDATE=" & PrcType(varPara(32), FldDate) & "," & "USEDATE=" & PrcType(varPara(33), FldDate) & _
                    ",LASTLOGINDATE=" & PrcType(varPara(34), FldDate) & ",UPDTIME=" & PrcType(varPara(35)) & _
                    ",COMPCODE=" & PrcType(varPara(36), FldNum) & ",INVTITLE=" & PrcType(varPara(37)) & _
                    ",INVOICETYPE=" & IIf(varPara(38) = Empty, "NULL", Val(varPara(38))) & _
                    ",INVADDRESS=" & PrcType(varPara(39)) & _
                    ",INVNO=" & PrcType(varPara(40)) & ",FACISNO=" & PrcType(varPara(41)) & _
                    " WHERE ( AUTHENTICID='" & strAuthenticID & "')"

'PrcType(varPara(38), FldNum)

'    strSQL = "UPDATE " & strOwner & "SO002B " & _
'                    "SET MEMBERID='" & varPara(1) & "'," & _
'                    "NAME='" & varPara(4) & "'," & _
'                    "ID='" & varPara(5) & "'," & _
'                    "BIRTHDAY=TO_DATE('" & varPara(6) & "', 'YYYY/MM/DD')," & _
'                    "TEL='" & varPara(7) & "'," & _
'                    "MOBILE='" & varPara(8) & "'," & _
'                    "SEXTUAL=" & varPara(9) & "," & _
'                    "ADDRESS='" & varPara(10) & "'," & _
'                    "EMAIL='" & varPara(11) & "'," & _
'                    "PPVORDERPWD='" & varPara(12) & "'," & _
'                    "LOGINPASSWORD='" & varPara(13) & "' " & _
'                    "WHERE ( AUTHENTICID='" & strAuthenticID & "')"
    cn.Execute strSQL, intRet
    Update_2_SO002B = (intRet > 0)
    strSQL = Empty
  Exit Function
ChkErr:
    ErrHandle "mod_Func2", "Update_2_SO002B"
    Update_2_SO002B = False
End Function
