Attribute VB_Name = "mod_Func6"
Option Explicit

Private strAuthenticID As String
Private cn As Object

Public Function JustDoIt(ByRef objWebPool As Object, _
                                        ByRef strPara As String) As String
  On Error GoTo ChkErr
    Dim strQry As String, strSubQry1 As String, strSubQry2 As String
    Dim varPara As Variant
    Dim rs As Object
    Dim strOwner As String
    Dim objCnPool As Object
    
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
    
    '   客戶有線電視收視費轉帳作業/計次付費預付金額轉帳作業(信用卡小額付款機制)--線上付款授權--收查詢
    '   6、會員代號、認證ID
    '   XML(單據編號、項次、收費項目名稱、應收日期、應收金額、收費期數、有效起始日期、有效截止日期、未收原因、服務類別);
    '   過濾未收且未指定付款之資料及線上付款請授權通過部分
    '   0:  成功
    '   -6 : 查無帳務資料
    
    varPara = Split(strPara, ",")
    '*****************************************************
    '#3018 增加判斷傳入的個數是否正確 By Kin 2007/08/30
    If UBound(varPara) <> 2 Then
        JustDoIt = "-99,參數不足!"
        GoTo 66
    End If
    '*****************************************************

    strAuthenticID = varPara(2)
    bytComp = Val(Left(strAuthenticID, 3))
    strOwner = GetOwner(cn)
    
    '    Select count(*) as Cnt from so002b where MemberId=[會員代號] and AuthenticId=[認證ID]
    '    Select FaciSNo from so002b where MemberId=[會員代號] and AuthenticId=[認證ID]
    strQry = "SELECT COUNT(*) AS CNT FROM " & strOwner & "SO002B" & Get_DB_Link(cn) & _
                    "WHERE MEMBERID='" & varPara(1) & "' AND AUTHENTICID='" & strAuthenticID & "'"
    Set rs = cn.Execute(strQry)
    
    '    If Cnt = 0 Then
    '      回傳(-2,訂購PPV認證錯誤)
    '    End If
    
'    If rs.RecordCount <= 0 Then
    If rs("CNT") <= 0 Then
        JustDoIt = "-2,訂購PPV認證錯誤"
        GoTo 66
    End If
    
    '    舊 :
    '    Select 單據編號、項次、收費項目名稱、應收日期、應收金額、收費期數、有效起始日期、有效截止日期、未收原因、服務類別
    '    from so033 where AuthenticId=[認證ID] and 未收(有值且參考號<>3)
    '    且未指定付款之資料(AccountNo is null) and realdate is null
    '    order by ShouldDate Desc, BillNo, ItemNo
    
    '   新 :
    '   Select單據編號、項次、收費項目名稱、應收日期、應收金額、收費期數、有效起始日期、有效截止日期、未收原因、服務類別
    '   from so033 where FaciSeqNo in [FaciSNo] and 未收(有值且參考號<>3)
    '   且未指定付款之資料(AccountNo is null) and realdate is null
    '   order by ShouldDate Desc, BillNo, ItemNo ;
    
    strSubQry1 = "SELECT FACISNO FROM " & strOwner & "SO002B" & Get_DB_Link(cn) & _
                            " WHERE MEMBERID='" & varPara(1) & "' AND AUTHENTICID='" & strAuthenticID & "'" & _
                            " AND FACISNO IS NOT NULL"
                            
    On Error Resume Next
    strSubQry1 = RPxx(cn.Execute(strSubQry1).GetString(2, , "", ",", ""))
    If Err = 3021 Then
        Err.Clear
        strSubQry1 = ""
    End If
    
    On Error GoTo ChkErr
    If Len(strSubQry1) > 0 Then
        strSubQry2 = "SELECT CODENO FROM " & strOwner & "CD013" & Get_DB_Link(cn) & " WHERE NVL(REFNO,0) <> 3"
        strQry = "SELECT BILLNO,ITEM,CITEMNAME,SHOULDDATE,SHOULDAMT,REALPERIOD,REALSTARTDATE" & _
                        ",REALSTOPDATE,UCNAME,SERVICETYPE FROM " & strOwner & "SO033" & Get_DB_Link(cn) & _
                        " WHERE FACISEQNO IN (" & Left(strSubQry1, Len(strSubQry1) - 1) & ")" & _
                        " AND UCCODE IN (" & strSubQry2 & ")" & _
                        " AND ACCOUNTNO IS NULL" & _
                        " AND REALDATE IS NULL" & _
                        " ORDER BY SHOULDDATE DESC,BILLNO,ITEM"
        Set rs = cn.Execute(strQry)
        If rs.RecordCount <= 0 Then
            JustDoIt = "-6,查無帳務資料"
        Else
            JustDoIt = "0,成功" & vbCrLf & ExpXML(rs)
        End If
    Else
'        JustDoIt = "-99,[客戶認證資訊檔] 查無 [認證設備序號] 資料!"
        JustDoIt = "-23,該認證ID查無認證設備或服務"
    End If
    
    '   If RecordCount = 0 Then
    '       回傳(-6,查無帳務資料)
    '   Else
    '       產生XML檔.(依select出來之資料)
    '       回傳(0,成功)及XML檔
    '   End If
    
66:
    On Error Resume Next
    Rlx rs
    Rlx strQry
    Rlx strSubQry1
    Rlx strSubQry2
    Rlx varPara
    Rlx strOwner
    Rlx strAuthenticID
    Set cn = Nothing
    Set objCnPool = Nothing
  Exit Function
ChkErr:
    ErrHandle "mod_Func6", "JustDoIt"
    JustDoIt = "-99," & strErr
End Function





