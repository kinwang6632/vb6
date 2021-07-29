Attribute VB_Name = "mod_Func20"
Option Explicit

Private cn As Object
Private strAuthenticID As String

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
    
    ' 欠費客戶名單查詢 ( 命令 20 )
    '   20、會員代號、認證ID
    '   0:  成功
    '   -17 : 查無欠費資料
    
    varPara = Split(strPara, ",")
    '*****************************************************
    '#3018 增加判斷傳入的個數是否正確 By Kin 2007/09/03
    If UBound(varPara) < 2 Then
        JustDoIt = "-99,參數不足!"
        GoTo 66
    End If
    '*****************************************************
    
    strAuthenticID = varPara(2)
    bytComp = Val(Left(strAuthenticID, 3))
    strOwner = GetOwner(cn)
    
    'Select count(*) as Cnt from so002b where MemberId=[會員代號] and AuthenticId=[認證ID]
    strQry = "SELECT COUNT(*) AS CNT FROM " & strOwner & "SO002B" & Get_DB_Link(cn) & _
                    "WHERE MEMBERID='" & varPara(1) & "' AND AUTHENTICID='" & strAuthenticID & "'"
    Set rs = cn.Execute(strQry)
    
    'If Cnt = 0 Then
    '  回傳(-2,訂購PPV認證錯誤)
    'End If
    
    If rs.RecordCount <= 0 Then
        JustDoIt = "-2,訂購PPV認證錯誤"
        GoTo 66
    End If
    
    'select BillNo, Item, CitemName, ShouldDate, ShouldAmt, RealPeriod, RealStartDate, RealStopDate,
    ' UCName, ServiceType from so033 where AuthenticId=[認證ID] and
    ' UCCode is not null and CancelFlag=0 Order by BillNo, Item, ServiceType

    strQry = "SELECT BILLNO,ITEM,CITEMNAME,SHOULDDATE,SHOULDAMT,REALPERIOD,REALSTARTDATE," & _
                    "REALSTOPDATE,UCNAME,SERVICETYPE FROM " & strOwner & "SO033" & Get_DB_Link(cn) & _
                    " WHERE AUTHENTICID='" & strAuthenticID & "' AND" & _
                    " UCCODE IS NOT NULL AND CANCELFLAG=0" & _
                    " ORDER BY BILLNO,ITEM,SERVICETYPE"
    
    Set rs = cn.Execute(strQry)
    
    'If RecordSet筆數 > 0 Then
    '  產生XML檔.(依select出來之資料)
    '  回傳(0,成功)及XML檔
    'Else
    '   回傳(-17,查無欠費資料)
    'End If
    
    If rs.RecordCount <= 0 Then
        JustDoIt = "-17,查無欠費資料"
    Else
        JustDoIt = "0,成功" & vbCrLf & ExpXML(rs)
    End If
    
    'PS:
    'XML(單據編號、項次、收費項目名稱、應收日期、應收金額、收費期數、
    ' 有效起始日期、有效截止日期、未收原因、服務類別);未收資料
   
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
    ErrHandle "mod_Func20", "JustDoIt"
    JustDoIt = "-99," & strErr
End Function

