Attribute VB_Name = "mod_Func17"
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
    Dim strFaciSno As String
    Dim objCnPool As Object
    Dim strOwner As String
    
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
    lngCustID = Val(Right(strAuthenticID, 7))
    strOwner = GetOwner(cn)
    
    '客戶安裝STB設備資料查詢(已安裝STB型號,編號及位置)
    '17、會員代號、認證ID
    'XML
    '   ----------- 舊的 -----------
    '(設備流水號、STB序號、智慧卡序號、母機序號、裝置點、型號名稱、回傳模組、安裝日期、PPV使用權、iPPV使用權、節目等級)
    '   ----------- 新的 -----------
    '(設備流水號、STB序號、智慧卡序號、母機序號、裝置點、型號名稱、回傳模組、安裝日期、PPV使用權、iPPV使用權、節目等級、節目等級名稱)
    '0 : 成功
    '-14 : 查無安裝STB設備資料
    
    '   Select CustId as nCustId from so002b where MemberId=[會員代號] and AuthenticId=[認證ID]
    '   If nCustId Is Null Then
    '       回傳(-2,訂購PPV認證錯誤)
    '   End If
    
'    strQry = "SELECT COUNT(*) AS RCDCNT FROM " & strOwner & "SO002B" & Get_DB_Link(cn) & _
                    "WHERE MEMBERID='" & varPara(1) & "' AND AUTHENTICID='" & strAuthenticID & "'"
    
'    If Val(GetValue(cn, strQry)) = 0 Then
'        JustDoIt = "-2,訂購PPV認證錯誤"
'        GoTo 66
'    End If
    
'    strQry = "SELECT CUSTID FROM " & strOwner & "SO002B" & Get_DB_Link(cn) & _
'                    "WHERE MEMBERID='" & varPara(1) & "' AND AUTHENTICID='" & strAuthenticID & "'"
'
'    lngCustID = Val(GetValue(cn, strQry) & "")
'
'    If lngCustID = 0 Then
'        JustDoIt = "-2,訂購PPV認證錯誤"
'        GoTo 66
'    End If
    
    strQry = "SELECT CUSTID,FACISNO FROM " & strOwner & "SO002B" & Get_DB_Link(cn) & _
                    " WHERE MEMBERID='" & varPara(1) & "' AND AUTHENTICID='" & strAuthenticID & "'"

    Set rs = cn.Execute(strQry)
    
    If rs.EOF Then
        JustDoIt = "-2,訂購PPV認證錯誤"
        GoTo 66
    Else
        lngCustID = Val(rs!CustID & "")
        strFaciSno = rs!FaciSno & ""
        ' Else If  [SO002B.FaciSNo] = null  THEN   --2006/07/12 Add
        If strFaciSno = Empty Then
            JustDoIt = "-23,該認證ID查無認證設備或服務"
            GoTo 66
            Exit Function
        End If
    End If
    
    '   ----------- 舊的 -----------
    '   Select a.SEQNo, a.FaciSNO, a.SmartCardNo, a.ReFaciSNO, c.Description, a.ModelName, b.FuncFlag01,
    '   a.InstDate, a.PPVright, a.PPVright, a.PgNo from so004 a, cd043 b, cd056 c, cd022 d where
    '   a.CustIt=[nCustId] and a.InitPlaceNo =c.CodeNo and a.ModelCode=b.CodeNo and
    '   a.InstDate is not null and a.PRDate is null and a.FaciCode=d.CodeNo and d.RefNo=3
    '   Order By a.InstDate, a.SEQNo
    
    '   ----------- 新的 -----------
    '   Select a.SEQNo, a.FaciSNO, a.SmartCardNo, a.ReFaciSNO, c.Description, a.ModelName, b.FuncFlag01,
    '   a.InstDate, a.PPVright, a.PPVright, a.PgNo, e.Description from so004 a, cd043 b, cd056 c, cd022 d, cd029 e
    '   where a.CustIt=[nCustId] and a.InitPlaceNo =c.CodeNo and a.ModelCode=b.CodeNo and a.InstDate is not null
    '   and a.PRDate is null and a.FaciCode=d.CodeNo and d.RefNo=3 and a.PgNo=e.CodeNo
    '   Order By a.InstDate, a.SEQNo
    '   文件裡的 PPVright 寫了兩次
    
'   1913
'   Debby
'    程式日期：CS_Web.dll　2005/11/17　下午6:16
'    【API-17】客戶安裝STB設備資料查詢(已安裝STB型號,編號及位置)
'    1．與意藍測試資料接收成功，但回傳欄位個數不符：文件上為回傳12　個，實際回傳11個，其中PPVRIGHT重覆送2次。請修改…
'    (1)少給節目等級名稱；應依規格串回CD029取得並回傳。
'    (2)PPVRIGHT多送一個，應傳予IPPVRIGHT才對。
'
'    2．API-17規格是會關聯至CD056、CD043，有對應到了才算找到，但因SO004設備資料檔有些裝置點、設備型號無值，
'    以致該戶有設備但會回應查無設備，故異動規格為……
'    只要於該戶SO004有找到正常使用之盒子，其裝置點、設備型號若無值，亦需為有找到
    
'                    strOwner & "CD022" & Get_DB_Link(cn) & "D,"

'API -17
'1． 其回傳回來的內容，其DESCRIPTION欄位名稱重覆，請調整回傳回來欄位名稱…
'(1) 裝置點：目前回傳名稱為Description ' 修改給予別名為 InitPlace
'(2) 節目名稱：目前回傳名稱為Description ' 修改給予別名為 PgName

    strQry = "SELECT A.SEQNO,A.FACISNO,A.SMARTCARDNO,A.REFACISNO,C.DESCRIPTION INITPLACE,A.MODELNAME," & _
                    "B.FUNCFLAG01,A.INSTDATE,A.PPVRIGHT,A.IPPVRIGHT,A.PGNO,E.DESCRIPTION PGNAME " & _
                    "FROM " & _
                    strOwner & "SO004" & Get_DB_Link(cn) & "A," & _
                    strOwner & "CD043" & Get_DB_Link(cn) & "B," & _
                    strOwner & "CD056" & Get_DB_Link(cn) & "C," & _
                    strOwner & "CD029" & Get_DB_Link(cn) & "E " & _
                    " WHERE A.CUSTID=" & lngCustID & _
                    " AND A.INITPLACENO=C.CODENO(+)" & _
                    " AND A.MODELCODE=B.CODENO(+)" & _
                    " AND A.INSTDATE IS NOT NULL AND A.PRDATE IS NULL" & _
                    " AND A.PGNO=E.CODENO(+)" & _
                    " AND A.FACICODE IN (SELECT CODENO FROM " & strOwner & "CD022" & Get_DB_Link(cn) & "WHERE REFNO=3)" & _
                    " AND A.SEQNO IN (" & strFaciSno & ")" & _
                    " ORDER BY A.INSTDATE,A.SEQNO"
    
    '   ----------- 舊的 -----------
'    strQry = "SELECT A.SEQNO,A.FACISNO,A.SMARTCARDNO,A.REFACISNO,C.DESCRIPTION,A.MODELNAME," & _
                    "B.FUNCFLAG01,A.INSTDATE,A.PPVRIGHT,A.PPVRIGHT,A.PGNO " & _
                    "FROM " & _
                    strOwner & "SO004" & Get_DB_Link(cn) & "A," & _
                    strOwner & "CD043" & Get_DB_Link(cn) & "B," & _
                    strOwner & "CD056" & Get_DB_Link(cn) & "C " & _
                    " WHERE A.CUSTID=" & lngCustID & _
                    " AND A.INITPLACENO=C.CODENO " & _
                    " AND A.MODELCODE=B.CODENO" & _
                    " AND A.INSTDATE IS NOT NULL AND A.PRDATE IS NULL" & _
                    " AND A.FACICODE IN " & _
                    "(SELECT CODENO FROM " & strOwner & "CD022" & Get_DB_Link(cn) & "WHERE REFNO=3)" & _
                    " ORDER BY A.INSTDATE,A.SEQNO"
    
    '   ----------- 舊的 -----------
    '    strQry = "SELECT A.SEQNO,A.FACISNO,A.SMARTCARDNO,A.REFACISNO,C.DESCRIPTION,A.MODELNAME," & _
    '                    "B.FUNCFLAG01,A.INSTDATE,A.PPVRIGHT,A.PPVRIGHT,A.PGNO " & _
    '                    "FROM " & strOwner & "SO004 A," & _
    '                    strOwner & "CD043 B," & _
    '                    strOwner & "CD056 C," & _
    '                    strOwner & "CD022 D " & _
    '                    "WHERE A.CUSTID=" & lngCustId & _
    '                    " AND A.INITPLACENO=C.CODENO " & _
    '                    " AND A.MODELCODE=B.CODENO" & _
    '                    " AND A.INSTDATE IS NOT NULL AND A.PRDATE IS NULL" & _
    '                    " AND A.FACICODE=D.CODENO AND D.REFNO=3" & _
    '                    " ORDER BY A.INSTDATE,A.SEQNO"
    
    Set rs = cn.Execute(strQry)
    
    '    If RecordSet筆數 > 0 Then
    '      產生XML檔.(依select出來之資料)
    '      回傳(0,成功)及XML檔
    '    Else
    '       回傳(-14,查無安裝STB設備資料)
    '    End If
    
    If rs.RecordCount <= 0 Then
        JustDoIt = "-14,查無安裝STB設備資料"
    Else
        JustDoIt = "0,成功" & vbCrLf & ExpXML(rs)
    End If
    
    '   PS:
    '   ----------- 舊的 -----------
    '    XML(設備流水號、STB序號、智慧卡序號、母機序號、裝置點、型號名稱、回傳模組、安裝日期、PPV使用權、iPPV使用權、節目等級)
    '   ----------- 新的 -----------
    '    XML(設備流水號、STB序號、智慧卡序號、母機序號、裝置點、型號名稱、回傳模組、安裝日期、PPV使用權、iPPV使用權、節目等級、節目等級名稱)
    
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
    ErrHandle "mod_Func17", "JustDoIt"
    JustDoIt = "-99," & strErr
End Function




