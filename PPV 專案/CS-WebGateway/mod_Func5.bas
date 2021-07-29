Attribute VB_Name = "mod_Func5"
Option Explicit

Private strAuthenticID As String
Private cn As Object

Public Function JustDoIt(ByRef objWebPool As Object, _
                                        ByRef strPara As String) As String
  On Error GoTo ChkErr
    Dim strQry As String
    Dim varPara As Variant
    Dim rs As Object
    Dim lngCustID As Long
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
    
    '   客戶帳務資料查詢(日期區間內得歷次帳單及繳款記錄)(個人)
    '   5、會員代號、認證ID、日期起日、日期迄日
    
    varPara = Split(strPara, ",")
    '*****************************************************
    '#3018 增加判斷傳入的個數是否正確 By Kin 2007/08/30
    If UBound(varPara) <> 4 Then
        JustDoIt = "-99,參數不足!"
        GoTo 66
    End If
    '*****************************************************

    
    strAuthenticID = varPara(2)
    bytComp = Val(Left(strAuthenticID, 3))
    strOwner = GetOwner(cn)
    
    '   Select count(*) as Cnt from so002b where MemberId=[會員代號] and AuthenticId=[認證ID]
    strQry = "SELECT COUNT(*) AS CNT FROM " & strOwner & "SO002B" & Get_DB_Link(cn) & _
                    "WHERE MEMBERID='" & varPara(1) & "' AND AUTHENTICID='" & strAuthenticID & "'"
    Set rs = cn.Execute(strQry)
    
    '   If Cnt = 0 Then
    '       回傳(-2,訂購PPV認證錯誤)
    '   End If
    
    If rs.RecordCount <= 0 Then
        JustDoIt = "-2,訂購PPV認證錯誤"
        GoTo 66
    End If
    
    '   5、會員代號、認證ID、日期起日、日期迄日

    '   Select 單據編號、項次、收費項目名稱、應收日、實收日、應收金額、實收金額、
    '   收費期數、有效起始日期、有效截止日期、收費人員姓名、收費方式、未收原因、
    '   短收原因、作廢、作廢原因、服務類別
    '   from so033與so034 where AuthenticId=[認證ID] and ShouldDate>=[日期起日] and
    '   ShouldDate<=[日期迄日] order by ShouldDate Desc, BillNo, ItemNo
    
    '   XML
    '   (單據編號、項次、收費項目名稱、應收日、實收日、應收金額、實收金額、收費期數、有效起始日期、
    '   有效截止日期、收費人員姓名、收費方式、未收原因、短收原因、作廢、作廢原因、服務類別)
   
    '   0:  成功
    '   -6 : 查無帳務資料
    
    strQry = "SELECT * FROM (" & _
                    "SELECT BILLNO,ITEM,CITEMNAME,SHOULDDATE,REALDATE,SHOULDAMT,REALAMT," & _
                    "REALPERIOD,REALSTARTDATE,REALSTOPDATE,CLCTNAME,CMNAME,UCNAME,STNAME," & _
                    "DECODE(CANCELFLAG,1,'是','否') CANCELFLAG,CANCELNAME,SERVICETYPE,GUINO" & _
                    " FROM " & strOwner & "SO033" & Get_DB_Link(cn) & _
                    "WHERE AUTHENTICID='" & strAuthenticID & "'" & _
                    " AND SHOULDDATE >= " & PrcType(varPara(3), FldDate) & _
                    " AND SHOULDDATE < " & PrcType(varPara(4), FldDate) & "+1" & _
                    " UNION ALL " & _
                    "SELECT BILLNO,ITEM,CITEMNAME,SHOULDDATE,REALDATE,SHOULDAMT,REALAMT," & _
                    "REALPERIOD,REALSTARTDATE,REALSTOPDATE,CLCTNAME,CMNAME,UCNAME,STNAME," & _
                    "DECODE(CANCELFLAG,1,'是','否') CANCELFLAG,CANCELNAME,SERVICETYPE,GUINO" & _
                    " FROM " & strOwner & "SO034" & Get_DB_Link(cn) & _
                    "WHERE AUTHENTICID='" & strAuthenticID & "'" & _
                    " AND SHOULDDATE >= " & PrcType(varPara(3), FldDate) & _
                    " AND SHOULDDATE < " & PrcType(varPara(4), FldDate) & "+1" & _
                    " ) ORDER BY SHOULDDATE DESC,BILLNO,ITEM"

    Set rs = cn.Execute(strQry)
    
    '   If RecordCount = 0 Then
    '       回傳(-6,查無帳務資料)
    '   Else
    '       產生XML檔.(依select出來之資料)
    '       回傳(0,成功)及XML檔
    '   End If
    
    If rs.RecordCount <= 0 Then
        JustDoIt = "-6,查無帳務資料"
    Else
        JustDoIt = "0,成功" & vbCrLf & ExpXML(rs)
    End If
    
    '   PS:  XML產出之欄位說明
    '   收費方式: CMName
    '   未收原因: UCName
    '   短收原因: STName
    '   作廢: 0→否 ,1→是
    '   作廢原因: CancelName
   
66:
    On Error Resume Next
    Rlx rs
    Rlx strQry
    Rlx varPara
    Rlx strOwner
    Rlx lngCustID
    Rlx strAuthenticID
    Set cn = Nothing
    Set objCnPool = Nothing
  Exit Function
ChkErr:
    ErrHandle "mod_Func5", "JustDoIt"
    JustDoIt = "-99," & strErr
End Function



