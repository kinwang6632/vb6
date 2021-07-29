Attribute VB_Name = "mod_Func19"
Option Explicit

Private cn As Object

Public Function JustDoIt(ByRef objWebPool As Object, _
                                        ByRef strPara As String) As String
  On Error GoTo ChkErr
    Dim rs As Object
    Dim strQry As String
    Dim varPara As Variant
    Dim strOwner As String
    Dim objCnPool As Object
    Dim strAuthenticID As String
    
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
    '#3018 增加判斷傳入的個數是否正確 By Kin 2007/09/03
    If UBound(varPara) < 2 Then
        JustDoIt = "-99,參數不足!"
        GoTo 66
    End If
    '*****************************************************

    strAuthenticID = varPara(2)
    bytComp = Val(Left(strAuthenticID, 3))
    strOwner = GetOwner(cn)
    
    ' 傳回SMS會員資料給WEB Update
    ' 19、會員代號、認證ID
    
    ' XML
    ' (客服網站會員代號、會員名稱、身分證字號、出生日期-年、出生日期-月、出生日期-日、
    ' 聯絡電話(宅) 、聯絡電話(公)、行動電話、性別、通訊地址-縣市、通訊地址-鄉鎮市區、
    ' 通訊地址-路段街巷弄巷號樓、Email、客服網站登入帳號、PPV訂購密碼、客服網站會員密碼、
    ' 密碼提示問題、密碼提示答案、學歷、願意收到有線電視服務及頻道節目等相關訊息、婚姻狀況、
    ' 同住人數、同住成員、職業別、個人年收入、身份別、帳號狀態、申請日期時間、啟用日期、
    ' 最後登入日期、最後異動日期、系統台代碼)
    
    ' 2006/04/12 Add -- 發票抬頭、發票種類、發票地址、統一編號
    
    strQry = "SELECT MEMBERID,NAME,ID,BIRTHDAYYYYY,BIRTHDAYMM,BIRTHDAYDD," & _
                    "TEL,TEL2,MOBILE,SEXTUAL,CITY,TOWN,ADDRESS,EMAIL,LOGINID," & _
                    "PPVORDERPWD,LOGINPASSWORD, HINTNAME,HINTANSWER,EDUCATION," & _
                    "MESSAGE,MARRIED,BODYCOUNT, MEMBERAGE,JOB,RENTE,CLASSID,STATUS," & _
                    "CREATEDATE,USEDATE,LASTLOGINDATE,UPDTIME,COMPCODE," & _
                    "INVTITLE,INVOICETYPE,INVADDRESS,INVNO,FACISNO" & _
                    " FROM " & strOwner & "SO002BLOG" & Get_DB_Link(cn) & _
                    "WHERE MEMBERID='" & varPara(1) & "' AND AUTHENTICID='" & strAuthenticID & "'"

    ' Select * from so002blog where MemberId=[會員代號] and AuthenticId=[認證ID]
    Set rs = cn.Execute(strQry)
    
    ' If RecordSet筆數 > 0 Then
    '   產生XML檔.(依select出來之資料)
    '   回傳(0,成功)及XML檔
    ' Else
    '   回傳(-16,查無傳回SMS會員資料)
    ' End If
    
    ' 0:  成功
    ' -16 : 查無傳回SMS會員資料
    
    ' -16, 該會員資料於SMS無異動
    If rs.RecordCount <= 0 Then
        JustDoIt = "-16,該會員資料於SMS無異動"
'        JustDoIt = "-16,查無傳回SMS會員資料"
    Else
        JustDoIt = "0,成功" & vbCrLf & ExpXML(rs)
    End If
    
66:
    On Error Resume Next
    Rlx rs
    Rlx strQry
    Rlx strOwner
    Set cn = Nothing
    Set objCnPool = Nothing
  Exit Function
ChkErr:
    ErrHandle "mod_Func19", "JustDoIt"
    JustDoIt = "-99," & strErr
End Function
