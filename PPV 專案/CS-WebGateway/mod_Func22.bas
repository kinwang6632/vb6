Attribute VB_Name = "mod_Func22"
Option Explicit

Private cn As Object
Private strOwner As String

' 22、會員代號、認證ID、種類(1=查帳務資訊、2=查設備資訊)、起始日期、截止日期
' 日期顯示格式 20060401

' 此支更改為多服務設備與收費資料查詢，XML有回傳 "SERVICETYPE"，以供區分分類。
' 分類如下:
' C 表CATV&DSTB
' I 表CM
' P 表CP

' 1． 查詢種類1回傳之 XML
'    (客編、單據編號、項次、收費項目名稱、應收日期、應收金額、實收日期、實收金額、實收期數
'    、有效起始日期、有效截止日期、服務類別、設備序號)

' 2． 查詢種類2回傳之 XML
'    (設備名稱、買賣方式、安裝日期、拆除日期、設備型號、CM速率、動態IP數、CM開機日期、CM關機日期、CM帳號啟用日期
'    、CM帳號停用日期、設備保証金、CM序號、Port數、方案代碼、方案名稱、綁約到期日、EBT合約編號、服務別、設備序號)

' 0: 成功
' -6 :查無帳務資料
' -14:查無安裝STB設備資料

Public Function JustDoIt(ByRef objWebPool As Object, _
                                        ByRef strPara As String) As String
  On Error GoTo ChkErr
  
    Dim strQry As String
    Dim varPara As Variant
    Dim rs As Object
    Dim objCnPool As Object
    Dim strAuthenticID As String
    Dim strStartDate As String
    Dim strStopDate As String
    Dim strCustID As String
    Dim strFaciSno As String
    Dim strBillCloseDate As String
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
    
    ' CATV&DSTB相關資訊查詢
    ' 依認證ID查詢所屬設備及收費資料
    ' 2006/04/11 改以認證ID查詢，請保留原始API-22之程式碼內容
    ' 傳入參數 [22, 會員代號, 認證ID, 種類(1=查帳務資訊、2=查設備資訊),起始日期,截止日期]

    varPara = Split(strPara, ",")
    '*****************************************************
    '#3018 增加判斷傳入的個數是否正確 By Kin 2007/09/03
    If UBound(varPara) < 3 Then
        JustDoIt = "-99,參數不足!"
        GoTo 66
    End If
    '*****************************************************

    strAuthenticID = varPara(2)
    bytComp = Val(Left(strAuthenticID, 3))
    
    On Error Resume Next
    strStartDate = varPara(4)
    strStopDate = varPara(5)
    
    On Error GoTo ChkErr
    strOwner = GetOwner(cn)
    
    If Val(varPara(3)) > 2 Or Val(varPara(3)) < 1 Then
            strErr = "傳入參數 [種類] 錯誤!!"
            JustDoIt = "-99,參數 [種類] 錯誤!"
        GoTo 66
    End If
    
    ' select Custid, FaciSNo from SO002B where MemberId=[會員代號] and AuthenticId=[認證帳號] ;
    
    strQry = "SELECT CUSTID,FACISNO FROM " & strOwner & "SO002B" & Get_DB_Link(cn) & _
                    " WHERE MEMBERID='" & varPara(1) & "' AND AUTHENTICID='" & strAuthenticID & "'"

    Set rs = cn.Execute(strQry)
    
    ' If RecordSet筆數 = 0 Then
    '    回傳(-1,帳戶認證錯誤)
    ' End If

    If rs.EOF Then
        JustDoIt = "-1,帳戶認證錯誤"
    Else
        JustDoIt = Empty
        strCustID = rs!CustID & ""
        strFaciSno = rs!FaciSno & ""
'        Else If  [SO002B.FaciSNo] = null  THEN   --2006/07/12 Add
'           回傳(-23, 該認證ID查無認證設備或服務)
        If strFaciSno = Empty Then
            JustDoIt = "-23,該認證ID查無認證設備或服務"
            GoTo 66
            Exit Function
        End If
        Select Case Val(varPara(3) & "")
                Case 1 ' CASE WHEN [種類]=1 THEN
                        '#3018 當查詢 種類 為1時，回傳之 XML，請在 [未收原因] 後 增加[媒體單號] ( MediaBillNo ) 及 [繳款截止日] ( BillCloseDate ) By Kin 2007/08/21
                        '      回傳收費資料的 [代收單繳費截止日] (BillCloseDate)，若該欄無值，則傳入 [應收日期] (ShouldDate)。
                        strBillCloseDate = "Decode(BillCloseDate,Null,ShouldDate,BillCloseDate) as BillCloseDate"
                        strQry = "SELECT CUSTID,BILLNO,ITEM,CITEMNAME,SHOULDDATE,SHOULDAMT,REALDATE,REALAMT" & _
                                        ",REALPERIOD,REALSTARTDATE,REALSTOPDATE,SERVICETYPE,FACISNO,UCNAME,MediaBillNo," & strBillCloseDate & _
                                        " FROM " & strOwner & "SO033" & Get_DB_Link(cn) & _
                                        " WHERE CUSTID=" & strCustID & _
                                        " AND ((FACISEQNO IN (" & strFaciSno & ")) OR AUTHENTICID='" & strAuthenticID & "')" & _
                                        " AND SHOULDDATE >= " & PrcType(strStartDate, FldDate) & _
                                        " AND SHOULDDATE < " & PrcType(strStopDate, FldDate) & "+1" & _
                                        " AND CANCELFLAG <> 1 UNION ALL " & _
                                         "SELECT CUSTID,BILLNO,ITEM,CITEMNAME,SHOULDDATE,SHOULDAMT,REALDATE,REALAMT" & _
                                         ",REALPERIOD,REALSTARTDATE,REALSTOPDATE,SERVICETYPE,FACISNO,UCNAME,MediaBillNo," & strBillCloseDate & _
                                        " FROM " & strOwner & "SO034" & Get_DB_Link(cn) & _
                                        " WHERE CUSTID=" & strCustID & _
                                        " AND ((FACISEQNO IN (" & strFaciSno & ")) OR AUTHENTICID='" & strAuthenticID & "')" & _
                                        " AND SHOULDDATE >= " & PrcType(strStartDate, FldDate) & _
                                        " AND SHOULDDATE < " & PrcType(strStopDate, FldDate) & "+1"
                                        
                        strQry = "SELECT * FROM (" & strQry & ") ORDER BY BILLNO DESC,ITEM"
                            
                        Set rs = cn.Execute(strQry)
                        
                        If rs.EOF Then ' If RecordSet筆數 > 0 Then
                            JustDoIt = "-6, 查無帳務資料" ' 回傳(-6, 查無帳務資料)
                        Else ' Else
                            ' 產生XML檔.(依select出來之資料)
                            JustDoIt = "0,成功" & vbCrLf & ExpXML(rs) ' 回傳(0, 成功)及XML檔
                        End If ' End If
'                       PS:
'                       XML(客編、單據編號、項次、收費項目名稱、應收日期、應收金額、實收日期、實收金額、
'                               實收期數、有效起始日期、有效截止日期、服務類別、設備序號)

                Case 2 ' CASE WHEN [種類]=2 THEN
                
'                        Select FaciName, BuyName, OldInstDate,PRDate,ModelName,CMBaudRate,DynIPCount,
'                            CMOpenDate,CMCloseDate,EnableAccount, DisableAccount,Deposit,ReFaciSNo,PortNo,
'                            BPCode,BPName,ContractDate,EBTContNo,ServiceType, FaciSNO from SO004
'                            where Custid=[SO002B.Custid] and SeqNo in [SO002B.FaciSNo]
'                            Union All
'                        Select FaciName, BuyName, OldInstDate,PRDate,ModelName,CMBaudRate,DynIPCount,
'                            CMOpenDate,CMCloseDate,EnableAccount, DisableAccount,Deposit,ReFaciSNo,PortNo,
'                            BPCode,BPName,ContractDate,EBTContNo,ServiceType, FaciSNO from SO004
'                            where ServiceType='P' and Custid=[SO002B.Custid] and ReFaciSNo in
'                            (Select FaciSNo from SO004 where Custid=[SO002B.Custid] and SeqNo=[SO002B.FaciSNo] ) ;

                        strQry = "SELECT FACINAME, BUYNAME, OLDINSTDATE,PRDATE,MODELNAME,CMBAUDRATE,DYNIPCOUNT" & _
                                        ",CMOPENDATE,CMCLOSEDATE,ENABLEACCOUNT, DISABLEACCOUNT,DEPOSIT,REFACISNO,PORTNO" & _
                                        ",BPCODE,BPNAME,CONTRACTDATE,EBTCONTNO,SERVICETYPE, FACISNO" & _
                                        ",SMARTCARDNO,DECLARANTNAME,DESCRIPTION INITPLACE,SEQNO" & _
                                        " FROM " & strOwner & "SO004" & Get_DB_Link(cn) & _
                                        "," & strOwner & "CD056" & Get_DB_Link(cn) & _
                                        " WHERE CUSTID=" & strCustID & " AND SEQNO IN (" & strFaciSno & ")" & _
                                        " AND SO004.INITPLACENO=CD056.CODENO(+)"
                                        
                        strQry = strQry & " UNION ALL " & _
                                        "SELECT FACINAME, BUYNAME, OLDINSTDATE,PRDATE,MODELNAME,CMBAUDRATE,DYNIPCOUNT" & _
                                        ",CMOPENDATE,CMCLOSEDATE,ENABLEACCOUNT, DISABLEACCOUNT,DEPOSIT,REFACISNO,PORTNO" & _
                                        ",BPCODE,BPNAME,CONTRACTDATE,EBTCONTNO,SERVICETYPE, FACISNO" & _
                                        ",SMARTCARDNO,DECLARANTNAME,DESCRIPTION INITPLACE,SEQNO" & _
                                        " FROM " & strOwner & "SO004" & Get_DB_Link(cn) & _
                                        "," & strOwner & "CD056" & Get_DB_Link(cn) & _
                                        " WHERE SERVICETYPE='P' AND CUSTID=" & strCustID & _
                                        " AND REFACISNO IN (SELECT FACISNO FROM " & strOwner & "SO004" & Get_DB_Link(cn) & _
                                        " WHERE CUSTID=" & strCustID & " AND SEQNO IN (" & strFaciSno & "))" & _
                                        " AND SO004.INITPLACENO=CD056.CODENO(+)"
                                        
                        strQry = strQry & " UNION ALL " & _
                                        "SELECT 'CATV服務', NULL BUYNAME, B.INSTTIME OLDINSTDATE,B.PRTIME PRDATE" & _
                                        ",NULL MODELNAME,NULL CMBAUDRATE,NULL DYNIPCOUNT,NULL CMOPENDATE" & _
                                        ",NULL CMCLOSEDATE,NULL ENABLEACCOUNT,NULL DISABLEACCOUNT,NULL DEPOSIT" & _
                                        ",NULL REFACISNO,NULL PORTNO,NULL BPCODE,NULL BPNAME,NULL CONTRACTDATE" & _
                                        ",NULL EBTCONTNO,B.SERVICETYPE,TO_CHAR(A.CUSTID) FACISNO,NULL SMARTCARDNO" & _
                                        ",A.CUSTNAME DECLARANTNAME,NULL INITPLACE,TO_CHAR(A.CUSTID) SEQNO" & _
                                        " FROM " & strOwner & "SO001" & Get_DB_Link(cn) & " A" & _
                                        "," & strOwner & "SO002" & Get_DB_Link(cn) & " B" & _
                                        " WHERE A.CUSTID IN (" & strFaciSno & ")" & _
                                        " AND A.CUSTID=B.CUSTID AND B.SERVICETYPE='C'"
                                        
'                        Union All
'                        Select 'CATV服務', null BuyName, InstTime OldInstDate,PRTime PRDate,null ModelName,
'                        null CMBaudRate,null DynIPCount,null CMOpenDate,null CMCloseDate,null EnableAccount,
'                        null DisableAccount,null Deposit,null ReFaciSNo,null PortNo,null BPCode,null BPName,
'                        null ContractDate,null EBTContNo,b.ServiceType, to_char(a.custid) FaciSNO, null SmartCard,
'                        a.CustName DeclarantName, null InitPlace
'                        from SO001 a, SO002 b
'                        where a.Custid in [SO002B.FaciSNo] and a.custid=b.custid and b.servicetype='C' ;
                                        
                        Set rs = cn.Execute(strQry)
                        
                        If rs.EOF Then ' If RecordSet筆數 > 0 Then
                            JustDoIt = "-14, 查無安裝STB設備資料" ' 回傳(-14, 查無安裝STB設備資料)
                        Else ' Else
                            ' 產生XML檔.(依select出來之資料)
                            JustDoIt = "0,成功" & vbCrLf & ExpXML(rs) ' 回傳(0, 成功)及XML檔
                        End If ' End If
'                       PS:
'                       XML(設備名稱、買賣方式、安裝日期、拆除日期、設備型號、CM速率、動態IP數、CM開機日期、CM關機日期、
'                               CM帳號啟用日期、CM帳號停用日期、設備保証金、CM序號、Port數、方案代碼、方案名稱、綁約到期日、
'                               EBT合約編號、服務類別、設備序號、設備流水號)
        End Select
    End If

66:
    On Error Resume Next
    Rlx varPara
    Rlx strCustID
    Rlx strFaciSno
    Rlx objCnPool
    Rlx strStartDate
    Rlx strStopDate
    Rlx strAuthenticID
    Set cn = Nothing
    Set objCnPool = Nothing
  Exit Function

ChkErr:
    ErrHandle "mod_Func22", "JustDoIt"
    JustDoIt = "-99," & strErr
End Function

'Private cn As Object
'Private strOwner As String
'
'Public Function JustDoIt(ByRef objWebPool As Object, _
'                                        ByRef strPara As String) As String
'  On Error GoTo ChkErr
'
'    Dim strQry As String
'    Dim varPara As Variant
'    Dim rs As Object
'    Dim objCnPool As Object
'
'    Set objCnPool = objWebPool
'    Set cn = objCnPool.GetConnection
'
'    If cn.state <= 0 Then
'        Set cn = objCnPool.GetConnection
'        If cn.state <= 0 Then
'            strErr = "無法連線資料庫!"
'            JustDoIt = "-99,資料庫連線失敗!"
'            GoTo 66
'        End If
'    End If
'
'    ' CATV&DSTB相關資訊查詢
'    ' 傳入參數 [22, 公司別,申請人, 身份証字號, 電話, 種類(1=查帳務資訊、2=查設備資訊),起始日期,截止日期]
'
'    varPara = Split(strPara, ",")
'
'    bytComp = varPara(1)
'    strOwner = GetOwner(cn)
'
'    If Val(varPara(5)) > 2 Or Val(varPara(5)) < 1 Then
'            strErr = "傳入參數 [種類] 錯誤!!"
'            JustDoIt = "-99,參數 [種類] 錯誤!"
'        GoTo 66
'    End If
'
'    ' ---- 舊 的 ----
'    '   Select distinct Custid, Facisno from (
'    '       Select Custid, null FaciSNO from SO001 where Compcode=[公司別] and CustName=[申請人]
'    '           and ID=[身份証字號] and Tel1=[電話]
'    '       Union All
'    '       Select Custid, Facisno from SO004 where Compcode=[公司別] and DeclarantName=[申請人]
'    '           and ID=[身份証字號] and Contel=[電話]
'    '           and facicode in (select codeno from cd022 where refno>0) and ServiceType='C' );
'    '
'    '   If RecordSet筆數 = 0 Then
'    '       回傳(-20, 查無申請人資料)
'    '   End If
'
'    ' ---- 新 的 ----
'    '    Select distinct Custid, Facisno from (
'    '       Select A.Custid, null FaciSNO from SO001 A, SO002 B where A.Compcode=[公司別]
'    '           and A.CUSTID=B.CUSTID AND B.SERVICETYPE='C' AND A.CustName=[申請人]
'    '           and B.ID=[身份証字號] and A.Tel1=[電話]
'    '       Union All
'    '       Select Custid, Facisno from SO004 where Compcode=[公司別] and DeclarantName=[申請人]
'    '           and ID=[身份証字號] and ConTtel=[電話]
'    '           and facicode in (select codeno from cd022 where refno>0) and ServiceType='C' );
'
'
'    strQry = "SELECT DISTINCT CUSTID,FACISNO FROM (" & _
'                    " SELECT A.CUSTID,NULL FACISNO" & _
'                    " FROM " & strOwner & "SO001" & Get_DB_Link(cn) & " A," & strOwner & "SO002" & Get_DB_Link(cn) & " B" & _
'                    " WHERE A.COMPCODE=" & varPara(1) & _
'                    " AND A.CUSTID=B.CUSTID" & _
'                    " AND B.SERVICETYPE='C'" & _
'                    " AND A.CUSTNAME='" & varPara(2) & "'" & _
'                    " AND B.ID='" & varPara(3) & "'" & _
'                    " AND A.TEL1='" & varPara(4) & "'" & _
'                    " UNION ALL " & _
'                    " SELECT CUSTID,FACISNO FROM " & strOwner & "SO004" & Get_DB_Link(cn) & _
'                    " WHERE COMPCODE=" & varPara(1) & _
'                    " AND DECLARANTNAME='" & varPara(2) & "'" & _
'                    " AND ID='" & varPara(3) & "'" & _
'                    " AND CONTTEL='" & varPara(4) & "'" & _
'                    " AND FACICODE IN " & _
'                    " (SELECT CODENO FROM " & strOwner & "CD022" & Get_DB_Link(cn) & " WHERE REFNO > 0)" & _
'                    " AND SERVICETYPE='C')"
'
'' 舊的
''    strQry = "SELECT DISTINCT CUSTID,FACISNO FROM (" & _
'                    " SELECT CUSTID,NULL FACISNO FROM " & strOwner & "SO001" & Get_DB_Link(cn) & _
'                    " WHERE COMPCODE=" & varPara(1) & _
'                    " AND CUSTNAME='" & varPara(2) & "'" & _
'                    " AND ID='" & varPara(3) & "'" & _
'                    " AND TEL1='" & varPara(4) & "'" & _
'                    " UNION ALL " & _
'                    " SELECT CUSTID,FACISNO FROM " & strOwner & "SO004" & Get_DB_Link(cn) & _
'                    " WHERE COMPCODE=" & varPara(1) & _
'                    " AND DECLARANTNAME='" & varPara(2) & "'" & _
'                    " AND ID='" & varPara(3) & "'" & _
'                    " AND CONTTEL='" & varPara(4) & "'" & _
'                    " AND FACICODE IN " & _
'                    " (SELECT CODENO FROM " & strOwner & "CD022" & Get_DB_Link(cn) & " WHERE REFNO > 0)" & _
'                    " AND SERVICETYPE='C')"
'
'    Set rs = cn.Execute(strQry)
'
'    '   If RecordSet筆數 = 0 Then
'    '       回傳(-20, 查無申請人資料)
'    '   End If
'
'    If rs.RecordCount <= 0 Then
'        JustDoIt = "-20, 查無申請人資料"
'    Else
'        JustDoIt = Empty
'        With rs ' LOOP 處理 (依申請人撈取出有符合的CustId, FaciSno)
''           [記錄select出來的資料]
'            If Val(varPara(5)) = 1 Then
'                While Not .EOF
'                    JustDoIt = JustDoIt & ExpXML(cn.Execute(GetTypeSQL(Val(varPara(5) & ""), !CustID & "", !FaciSno & "", _
'                                                                        varPara(6) & "", varPara(7) & "")), True)
'                    .MoveNext
'                Wend
'            Else
'                While Not .EOF
'                    JustDoIt = JustDoIt & ExpXML(cn.Execute(GetTypeSQL(Val(varPara(5) & ""), !CustID & "", !FaciSno & "")), True)
'                    .MoveNext
'                Wend
'            End If
'        End With ' END LOOP
'
'        If JustDoIt = Empty Then
'            If Val(varPara(5)) = 1 Then '   CASE WHEN [種類]=1 THEN
'                JustDoIt = "-6, 查無帳務資料" ' 回傳(-6, 查無帳務資料)
'            Else '   CASE WHEN [種類]=2 THEN
'                JustDoIt = "-14, 查無安裝STB設備資料" ' 回傳(-14, 查無安裝STB設備資料)
'            End If
'        Else '   If RecordSet筆數 > 0 Then
'            ' 產生XML檔.(依select出來之資料)
'            JustDoIt = "0,成功" & vbCrLf & _
'                                "<?xml version=""1.0"" encoding=""BIG5"" standalone=""yes"" ?>" & vbCrLf & _
'                                "<DataSet>" & vbCrLf & _
'                                "  <DataTable>" & vbCrLf & _
'                                JustDoIt & _
'                                "  </DataTable>" & vbCrLf & _
'                                "</DataSet>" '       回傳(0, 成功)及XML檔
'        End If
'    End If
'
'66:
'    On Error Resume Next
'    Rlx varPara
'    Rlx objCnPool
'    Set cn = Nothing
'    Set objCnPool = Nothing
'  Exit Function
'
'ChkErr:
'    ErrHandle "mod_Func22", "JustDoIt"
'    JustDoIt = "-99," & strErr
'End Function
'
'Private Function GetTypeSQL(bytType As Byte, _
'                                                strCustID As String, _
'                                                ByVal strFaciSNo As String, _
'                                                Optional strStartDate As String = "", _
'                                                Optional strStopDate As String = "") As String
'  On Error GoTo ChkErr
'
'    If bytType = 1 Then
'
'        'PS:
'        'XML(客編、單據編號、項次、收費項目名稱、應收日期、應收金額、實收日期、實收金額、實收期數、
'        '       有效起始日期、有效截止日期、未收原因、短收原因、服務類別、設備序號)
'
'        strFaciSNo = IIf(strFaciSNo = Empty, "FACISNO IS NULL", "FACISNO='" & strFaciSNo & "'")
'
'        GetTypeSQL = "SELECT CUSTID,BILLNO,ITEM,CITEMNAME,SHOULDDATE,SHOULDAMT,REALDATE,REALAMT" & _
'                                ",REALPERIOD,REALSTARTDATE,REALSTOPDATE,UCNAME,STNAME,SERVICETYPE,FACISNO" & _
'                                " FROM " & strOwner & "SO033" & Get_DB_Link(cn) & _
'                                " WHERE CUSTID=" & strCustID & " AND " & strFaciSNo & _
'                                " AND SHOULDDATE >= " & PrcType(strStartDate, FldDate) & _
'                                " AND SHOULDDATE < " & PrcType(strStopDate, FldDate) & "+1" & _
'                                " AND SERVICETYPE = 'C'" & _
'                                " UNION ALL " & _
'                                "SELECT CUSTID,BILLNO,ITEM,CITEMNAME,SHOULDDATE,SHOULDAMT,REALDATE,REALAMT" & _
'                                ",REALPERIOD,REALSTARTDATE,REALSTOPDATE,UCNAME,STNAME,SERVICETYPE,FACISNO" & _
'                                " FROM " & strOwner & "SO034" & Get_DB_Link(cn) & _
'                                " WHERE CUSTID=" & strCustID & " AND " & strFaciSNo & _
'                                " AND SHOULDDATE >= " & PrcType(strStartDate, FldDate) & _
'                                " AND SHOULDDATE < " & PrcType(strStopDate, FldDate) & "+1" & _
'                                " AND SERVICETYPE = 'C'"
'
''        If [FaciSno] Is Null Then
''            Select Custid, BillNo, Item, CitemName, ShouldDate, ShouldAmt, RealDate, RealAmt, RealPeriod,
''                RealStartDate, RealStopDate, UCName, STName, Servicetype, FaciSno from SO033
''                Where Custid = [Custid] And FaciSno Is Null And ShouldDate >= [起始日期] And ShouldDate < [截止日期] + 1
''                and ServiceType='C'
''                Union All
''            Select Custid, BillNo, Item, CitemName, ShouldDate, ShouldAmt, RealDate, RealAmt, RealPeriod,
''                RealStartDate, RealStopDate, UCName, STName, Servicetype FaciSno from SO034
''                Where Custid = [Custid] And FaciSno Is Null And ShouldDate >= [起始日期] And ShouldDate < [截止日期] + 1
''                and ServiceType='C' ;
''        Else
''            Select Custid, BillNo, Item, CitemName, ShouldDate, ShouldAmt, RealDate, RealAmt, RealPeriod,
''                RealStartDate, RealStopDate, UCName, STName, Servicetype, FaciSno from SO033
''                Where Custid = [Custid] And FaciSno = [FaciSno] And ShouldDate >= [起始日期] And ShouldDate < [截止日期] + 1
''                and ServiceType='C'
''                Union All
''            Select Custid, BillNo, Item, CitemName, ShouldDate, ShouldAmt, RealDate, RealAmt, RealPeriod,
''                RealStartDate, RealStopDate, UCName, STName, Servicetype, FaciSno from SO034
''                Where Custid = [Custid] And FaciSno = [FaciSno] And ShouldDate >= [起始日期] And ShouldDate < [截止日期] + 1
''                and ServiceType='C' ;
''        End If
'
'    Else
'
'        'PS:
'        'XML(客編、設備名稱、設備序號、買賣方式、安裝日期、拆除日期、設備型號、設備保証金、智慧卡序號、收費項目、有效起始日、有效截止日)
'
'        GetTypeSQL = "SELECT A.CUSTID,A.FACINAME,A.FACISNO,A.BUYNAME,A.OLDINSTDATE,A.PRDATE" & _
'                                ",A.MODELNAME,A.DEPOSIT,A.SMARTCARDNO,B.CITEMNAME,B.STARTDATE,B.STOPDATE" & _
'                                " FROM " & strOwner & "SO004" & Get_DB_Link(cn) & " A," & strOwner & "SO003" & Get_DB_Link(cn) & " B" & _
'                                " WHERE A.CUSTID=" & strCustID & " AND A.FACISNO='" & strFaciSNo & "'" & _
'                                " AND A.CUSTID=B.CUSTID(+) AND A.FACISNO=B.FACISNO(+)"
'
''        If [FaciSno] Is Not Null Then
''            Select a.Custid,a.FaciName,a.FaciSNO,a.BuyName,a.OldInstDate,a.PRDate,a.ModelName,a.Deposit,
''                   a.SmartCardNo,b.citemname,b.startdate,b.stopdate from SO004 a, so003 b
''                   where a.Custid=[Custid] and a.FaciSno=[FaciSno] and a.custid=b.custid(+) and a.facisno=b.facisno(+) ;
''        End If
'    End If
'
'  Exit Function
'ChkErr:
'    ErrHandle "mod_Func22", "GetTypeSQL"
'    GetTypeSQL = "-99," & strErr
'End Function
'
