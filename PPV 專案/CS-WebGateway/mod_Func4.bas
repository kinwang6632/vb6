Attribute VB_Name = "mod_Func4"
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
    
    'XML(訂單單號、訂單項次、節目名稱、設備流水號、STB序號、智慧卡序號、受理日期、結帳、
    '   預付節目、訂購方式、金額、點數、取消日期、取消人員、信用卡卡號)
    
    'Select a.訂單單號, a.訂單項次, a.節目名稱, a.設備流水號, a.STB序號, a.智慧卡序號, b.受理日期, a.結帳,
    '   a.預付節目, a.訂購方式, a.金額, a.點數, a.取消日期, a.取消人員, b.信用卡卡號
    '   from so105g a, so105e b
    '   where b.OrderNo=a.OrderNo and b.AuthenticId=[認證ID] and
    '    b.AcceptTime >=[訂購起日] and b.AcceptTime<=[訂購迄日]
    '   order by b.AcceptTime Desc, a.設備流水號, a.訂單單號, a.訂單項次
    
    ' 4、會員代號、認證ID、訂購起日、訂購迄日
    
    '預付節目: 0→後付 , 1→預付
    '信用卡卡號: 前四碼與後四碼內容正常, 中間八碼部分則傳XXXXXXXX.
    
    strQry = "SELECT A.ORDERNO,A.ORDERITEM ITEM,A.PRODUCTNAME,A.FACISEQNO," & _
                    "A.FACISNO,A.SMARTCARDNO,B.ACCEPTTIME," & _
                    "DECODE(A.ACCOUNTING,1,'是','否') ACCOUNTING," & _
                    "DECODE(A.PREPAY,1,'預付','後付') PREPAY,A.PPVORDERMODE," & _
                    "A.AMOUNT,A.CREDITAMT,A.CANCELTIME,A.CANCELNAME," & _
                    "DECODE(B.ACCOUNTNO,NULL,''," & _
                    "SUBSTR(B.ACCOUNTNO,1,4) || 'XXXXXXXX' || " & _
                    "SUBSTR(B.ACCOUNTNO,13)) CREDITCARD," & _
                    "C.EVENTBEGINTIME DEVENTBEGINTIME,A.PRODUCTID" & _
                    " FROM " & _
                     strOwner & "SO105G" & Get_DB_Link(cn) & "A," & _
                     strOwner & "SO105E" & Get_DB_Link(cn) & "B," & _
                     strOwner & "SO155" & Get_DB_Link(cn) & "C " & _
                    "WHERE B.ORDERNO=A.ORDERNO AND" & _
                    " B.AUTHENTICID='" & strAuthenticID & "' AND" & _
                    " B.ACCEPTTIME >= " & PrcType(varPara(3), FldTime) & " AND" & _
                    " B.ACCEPTTIME <= " & PrcType(varPara(4), FldTime) & "+1" & _
                    " AND A.PRODUCTID=C.PRODUCTID(+)" & _
                    " ORDER BY B.ACCEPTTIME DESC,A.FACISEQNO,A.ORDERNO,A.ORDERITEM"
                    
    '    SELECT A.ORDERNO,A.ORDERITEM,A.PRODUCTNAME,A.FACISEQNO,A.FACISNO,A.SMARTCARDNO,B.ACCEPTTIME,
    '        DECODE(A.ACCOUNTING,1,'是','否') ACCOUNTING,
    '        DECODE(A.PREPAY,1,'預付','後付') PREPAY,
    '        A.ORDERMODE,A.AMOUNT,A.CANCELTIME,A.CANCELNAME,
    '        SUBSTR(B.ACCOUNTNO,1,4) || 'XXXXXXXX' || SUBSTR(B.ACCOUNTNO,13) ACCOUNTNO
    '        FROM GICMIS5.SO105G A,GICMIS5.SO105E B
    '        WHERE B.ORDERNO=A.ORDERNO AND B.AUTHENTICID='0010000062' AND
    '        B.ACCEPTTIME >= TO_DATE('20050410' ,'YYYYMMDD')
    '        AND B.ACCEPTTIME <= TO_DATE('20050412' ,'YYYYMMDD')+1
    '        ORDER BY B.ACCEPTTIME DESC,A.FACISEQNO,A.ORDERNO,A.ORDERITEM
    
    '   If RecordCount = 0 Then
    '       回傳(-5,查無訂購資料)
    '   Else
    '       產生XML檔.(依select出來之資料)
    '       回傳(0,成功)及XML檔
    '   End If
    
    Set rs = cn.Execute(strQry)
    
    If rs.RecordCount <= 0 Then
        JustDoIt = "-5,查無訂購資料"
    Else
        JustDoIt = "0,成功" & vbCrLf & ExpXML(rs)
    End If
    
    'PS:  XML產出之欄位轉換說明
    '結帳: 0→否 , 1→是
    '預付節目: 0→後付 , 1→預付
    '信用卡卡號: 前四碼與後四碼內容正常, 中間八碼部分則傳XXXXXXXX.
    'Ex: 4311XXXXXXXX1234
   
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
    ErrHandle "mod_Func4", "JustDoIt"
    JustDoIt = "-99," & strErr
End Function

