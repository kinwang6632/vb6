Attribute VB_Name = "mod_Func23"
Option Explicit
    
Private cn As Object
Private strOwner As String

Public Function JustDoIt(ByRef objWebPool As Object, _
                                        ByRef strPara As String) As String
  On Error GoTo ChkErr
  
    Dim strQry As String
    Dim varPara As Variant
    Dim rs As Object
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
    
    ' CMCP相關資訊查詢
    ' 傳入參數 [23, 公司別, 申請人, 身份証字號, 生日, 種類(1=查帳務資訊、2=查設備資訊), 起始日期, 截止日期]
    
    varPara = Split(strPara, ",")

    bytComp = varPara(1)
    strOwner = GetOwner(cn)
    
    If Val(varPara(5)) > 2 Or Val(varPara(5)) < 1 Then
            strErr = "傳入參數 [種類] 錯誤!!"
            JustDoIt = "-99,參數 [種類] 錯誤!"
        GoTo 66
    End If
    
    '   Select CUSTID,FACISNO From SO004
    '       Where ServiceType='I' And DeclarantName=[申請人]
    '       And ID=[身份証字號] And Birthday=[生日];
    
    strQry = "SELECT CUSTID,FACISNO FROM " & strOwner & "SO004" & Get_DB_Link(cn) & _
                    " WHERE SERVICETYPE='I'" & _
                    " AND DECLARANTNAME='" & varPara(2) & "'" & _
                    " AND ID='" & varPara(3) & "'" & _
                    " AND BIRTHDAY=" & PrcType(varPara(4), FldDate) & _
                    " AND COMPCODE=" & varPara(1)
    
    Set rs = cn.Execute(strQry)
    
    '   If RecordSet筆數 = 0 Then
    '       回傳(-20, 查無申請人資料)
    '   End If
    
    If rs.RecordCount <= 0 Then
        JustDoIt = "-20, 查無申請人資料"
    Else
        JustDoIt = Empty
        With rs ' LOOP 處理 (依申請人撈取出有符合的CustId, FaciSno)
            If Val(varPara(5)) = 1 Then
                While Not .EOF
                    JustDoIt = JustDoIt & ExpXML(cn.Execute(GetTypeSQL(Val(varPara(5) & ""), !CustID & "", !FaciSno & "", _
                                                                        varPara(6) & "", varPara(7) & "")), True)
                    .MoveNext
                Wend
            Else
                While Not .EOF
                    JustDoIt = JustDoIt & ExpXML(cn.Execute(GetTypeSQL(Val(varPara(5) & ""), !CustID & "", !FaciSno) & ""), True)
                    .MoveNext
                Wend
            End If
        End With ' END LOOP
        
        If JustDoIt = Empty Then
            If Val(varPara(5)) = 1 Then '   CASE WHEN [種類]=1 THEN
                JustDoIt = "-6, 查無帳務資料" ' 回傳(-6, 查無帳務資料)
            Else '   CASE WHEN [種類]=2 THEN
                JustDoIt = "-21, 查無安裝CMCP設備資料" ' 回傳(-21, 查無安裝CMCP設備資料)
            End If
        Else
            JustDoIt = "0,成功" & vbCrLf & _
                                "<?xml version=""1.0"" encoding=""BIG5"" standalone=""yes"" ?>" & vbCrLf & _
                                "<DataSet>" & vbCrLf & _
                                "  <DataTable>" & vbCrLf & _
                                JustDoIt & _
                                "  </DataTable>" & vbCrLf & _
                                "</DataSet>" '       回傳(0, 成功)及XML檔
        End If
    End If

66:
    On Error Resume Next
    Rlx varPara
    Rlx objCnPool
    Set cn = Nothing
    Set objCnPool = Nothing
  Exit Function

ChkErr:
    ErrHandle "mod_Func23", "JustDoIt"
    JustDoIt = "-99," & strErr
End Function

Private Function GetTypeSQL(bytType As Byte, _
                                                strCustID As String, _
                                                strFaciSno As String, _
                                                Optional strStartDate As String = "", _
                                                Optional strStopDate As String = "") As String
  On Error GoTo ChkErr
  
    If bytType = 1 Then
        '   PS:
        '   XML(單據編號、項次、收費項目名稱、應收日期、應收金額、實收日期、實收金額、實收期數
        '           、有效起始日期、有效截止日期、未收原因、短收原因、作廢否、服務類別)
        GetTypeSQL = "SELECT BILLNO,ITEM,CITEMNAME,SHOULDDATE,SHOULDAMT" & _
                                ",REALDATE,REALAMT,REALPERIOD,REALSTARTDATE,REALSTOPDATE" & _
                                ",UCNAME,STNAME,DECODE(CANCELFLAG,1,'是','否') CANCELFLAG" & _
                                ",SERVICETYPE,FACISNO" & _
                                " FROM " & strOwner & "SO033" & Get_DB_Link(cn) & _
                                " WHERE CUSTID=" & strCustID & _
                                " AND FACISNO='" & strFaciSno & "'" & _
                                " AND SERVICETYPE IN ('I','P')" & _
                                " AND SHOULDDATE >= " & PrcType(strStartDate, FldDate) & _
                                " AND SHOULDDATE < " & PrcType(strStopDate, FldDate) & "+1" & _
                                " UNION ALL " & _
                                "SELECT BILLNO,ITEM,CITEMNAME,SHOULDDATE,SHOULDAMT" & _
                                ",REALDATE,REALAMT,REALPERIOD,REALSTARTDATE,REALSTOPDATE" & _
                                ",UCNAME,STNAME,DECODE(CANCELFLAG,1,'是','否') CANCELFLAG" & _
                                ",SERVICETYPE,FACISNO" & _
                                " FROM " & strOwner & "SO034" & Get_DB_Link(cn) & _
                                " WHERE CUSTID=" & strCustID & _
                                " AND FACISNO='" & strFaciSno & "'" & _
                                " AND SERVICETYPE IN ('I','P')" & _
                                " AND SHOULDDATE >= " & PrcType(strStartDate, FldDate) & _
                                " AND SHOULDDATE < " & PrcType(strStopDate, FldDate) & "+1"
                                
        '      Select BillNo, Item, CitemName, ShouldDate, ShouldAmt, RealDate, RealAmt, RealPeriod,
        '           RealStartDate, RealStopDate, UCName, STName, CancelFlag, ServiceType From SO033
        '           Where CustId=[CustId] And FaciSno=[FaciSno] And ServiceType In ('I','P')
        '           And ShouldDate>=[起始日期] And ShouldDate < [截止日期]+1
        '      Union All
        '          Select BillNo, Item, CitemName, ShouldDate, ShouldAmt, RealDate, RealAmt, RealPeriod,
        '           RealStartDate, RealStopDate, UCName, STName, CancelFlag, ServiceType From SO034
        '           Where CustId=[CustId] And FaciSno=[FaciSno] And ServiceType In ('I','P')
        '           And ShouldDate>=[起始日期] And ShouldDate < [截止日期]+1
    Else
        '   PS:
        '   XML(設備名稱、買賣方式、安裝日期、拆除日期、設備型號、CM速率、動態IP數、CM開機日期、CM關機日期
        '           、CM帳號啟用日期、CM帳號停用日期、設備保証金、CM序號、Port數、方案代碼、方案名稱
        '           、綁約到期日、EBT合約編號、服務別)
        
        GetTypeSQL = "SELECT FACINAME,BUYNAME,OLDINSTDATE,PRDATE,MODELNAME" & _
                                ",CMBAUDRATE,DYNIPCOUNT,CMOPENDATE,CMCLOSEDATE" & _
                                ",ENABLEACCOUNT,DISABLEACCOUNT,DEPOSIT,REFACISNO,PORTNO" & _
                                ",BPCODE,BPNAME,CONTRACTDATE,EBTCONTNO,SERVICETYPE,FACISNO" & _
                                " FROM " & strOwner & "SO004" & Get_DB_Link(cn) & _
                                " WHERE CUSTID=" & strCustID & _
                                " AND FACISNO='" & strFaciSno & "'" & _
                                " AND SERVICETYPE='I' " & _
                                " UNION ALL " & _
                                "SELECT FACINAME,BUYNAME,OLDINSTDATE,PRDATE,MODELNAME" & _
                                ",CMBAUDRATE,DYNIPCOUNT,CMOPENDATE,CMCLOSEDATE" & _
                                ",ENABLEACCOUNT,DISABLEACCOUNT,DEPOSIT,REFACISNO,PORTNO" & _
                                ",BPCODE,BPNAME,CONTRACTDATE,EBTCONTNO,SERVICETYPE,FACISNO" & _
                                " FROM " & strOwner & "SO004" & Get_DB_Link(cn) & _
                                " WHERE CUSTID=" & strCustID & _
                                " AND REFACISNO='" & strFaciSno & "'" & _
                                " AND SERVICETYPE='P' "

            '    Select FaciName, BuyName, OldInstDate,PRDate,ModelName,CMBaudRate,DynIPCount,CMOpenDate,
            '       CMCloseDate,EnableAccount, DisableAccount,Deposit,ReFaciSNo,PortNo,BPCode,BPName,
            '       ContractDate,EBTContNo,ServiceType, FaciSNO from SO004
            '       where ServiceType='I' and Custid=[Custid] and FaciSno=[FaciSNo]
            '    Union All
            '       Select FaciName, BuyName, OldInstDate,PRDate,ModelName,CMBaudRate,DynIPCount,CMOpenDate,
            '       CMCloseDate,EnableAccount, DisableAccount,Deposit,ReFaciSNo,PortNo,BPCode,BPName,
            '       ContractDate,EBTContNo,ServiceType, FaciSNO from SO004
            '       where ServiceType='P' and Custid=[Custid] and ReFaciSno=[FaciSNo]
                                    
    End If
      
  Exit Function
ChkErr:
    ErrHandle "mod_Func23", "GetTypeSQL"
    GetTypeSQL = "-99," & strErr
End Function

'        GetTypeSQL = "SELECT A.FACINAME,A.BUYNAME,A.OLDINSTDATE,A.PRDATE,A.MODELNAME" & _
'                                ",B.CMBAUDRATE,B.DYNIPCOUNT,A.CMOPENDATE,A.CMCLOSEDATE" & _
'                                ",A.ENABLEACCOUNT,A.DISABLEACCOUNT,A.DEPOSIT,A.REFACISNO" & _
'                                ",A.PORTNO,C.BPCODE,C.BPNAME,A.CONTRACTDATE,A.EBTCONTNO" & _
'                                ",A.SERVICETYPE,A.FACISNO" & _
'                                " FROM " & strOwner & "SO004" & Get_DB_Link(cn) & " A " & _
'                                " ," & strOwner & "CD078D" & Get_DB_Link(cn) & " B " & _
'                                " ," & strOwner & "SO003" & Get_DB_Link(cn) & " C " & _
'                                " WHERE A.CUSTID=" & strCustID & _
'                                " AND A.FACISNO='" & strFaciSno & "'" & _
'                                " AND A.SERVICETYPE='I' " & _
'                                " AND A.SERVICETYPE=C.SERVICETYPE" & _
'                                " AND A.FACISNO=C.FACISNO" & _
'                                " AND C.BPCODE=B.BPCODE" & _
'                                " UNION ALL " & _
'                                "SELECT FACINAME,BUYNAME,OLDINSTDATE,PRDATE,MODELNAME" & _
'                                ",CMBAUDRATE,DYNIPCOUNT,CMOPENDATE,CMCLOSEDATE" & _
'                                ",ENABLEACCOUNT,DISABLEACCOUNT,DEPOSIT,REFACISNO" & _
'                                ",PORTNO,BPCODE,BPNAME,CONTRACTDATE,EBTCONTNO" & _
'                                ",SERVICETYPE,FACISNO" & _
'                                " FROM " & strOwner & "SO004" & Get_DB_Link(cn) & _
'                                " WHERE CUSTID=" & strCustID & _
'                                " AND REFACISNO='" & strFaciSno & "'" & _
'                                " AND SERVICETYPE='P' "
                                    
    '       Select a.FaciName, a.BuyName, OldInstDate,PRDate,a.ModelName,b.CMBaudRate,
    '       b.DynIPCount,CMOpenDate,CMCloseDate,EnableAccount, DisableAccount,Deposit,
    '       ReFaciSNo,PortNo,c.BPCode,c.BPName,ContractDate,EBTContNo,a.ServiceType
    '       From SO004 a, CD078D b, SO003 c
    '       Where a.ServiceType='I' And a.Custid=[Custid] And a.FaciSno=[FaciSNo] And
    '       a.ServiceType=c.ServiceType And a.FaciSno=c.FaciSNo And c.bpcode=b.bpcode
    '       Union All
    '       Select FaciName, BuyName, OldInstDate,PRDate,ModelName,CMBaudRate,
    '       DynIPCount,CMOpenDate,CMCloseDate,EnableAccount, DisableAccount,Deposit,
    '       ReFaciSNo,PortNo,BPCode,BPName,ContractDate,EBTContNo,ServiceType
    '       From SO004
    '       Where ServiceType='P' And Custid=[Custid] And ReFaciSno=[FaciSNo]


'        If varPara(5) = 1 Then '   CASE WHEN [種類]=1 THEN
'
'            JustDoIt = Empty
'
'            With rs
'                While Not .EOF ' LOOP 處理 (依申請人撈取出有符合的CustId, FaciSno)
'                    JustDoIt = JustDoIt & _
'                                        ExpXML(cn.Execute(GetType1SQL(!CustID, !FaciSno, CStr(varPara(6)), CStr(varPara(7)))), True)
'                    .MoveNext
'                Wend ' END LOOP
'            End With
'
'            ' [記錄select出來的資料]
'
'            '   If RecordSet筆數 > 0 Then
'            '       產生XML檔.(依select出來之資料)
'            '       回傳(0, 成功)及XML檔
'            '   Else
'            '       回傳(-6, 查無帳務資料)
'            '   End If
'
'            If JustDoIt = Empty Then
'                JustDoIt = "-6, 查無帳務資料"
'            Else
'                JustDoIt = "0,成功" & vbCrLf & _
'                                    "<?xml version=""1.0"" encoding=""BIG5"" standalone=""yes"" ?>" & vbCrLf & _
'                                    "<DataSet>" & vbCrLf & _
'                                    "  <DataTable>" & vbCrLf & _
'                                    JustDoIt & _
'                                    "  </DataTable>" & vbCrLf & _
'                                    "</DataSet>"
'            End If
'            '   PS:
'            '   XML(單據編號、項次、收費項目名稱、應收日期、應收金額、實收日期、實收金額、實收期數
'            '           、有效起始日期、有效截止日期、未收原因、短收原因、作廢否、服務類別)
'
'        ElseIf varPara(5) = 2 Then '   CASE WHEN [種類]=2 THEN
'
'            JustDoIt = Empty
'            With rs
'                While Not .EOF ' LOOP 處理 (依申請人撈取出有符合的Custid, Facisno)
'                    JustDoIt = JustDoIt & _
'                                        ExpXML(cn.Execute(GetType2SQL(!CustID, !FaciSno, "", "")), True)
'                    .MoveNext
'                Wend ' END LOOP
'            End With
'
'            ' [記錄select出來的資料]
'
'            '   If RecordSet筆數 > 0 Then
'            '       產生XML檔.(依select出來之資料)
'            '       回傳(0, 成功)及XML檔
'            '   Else
'            '       回傳(-21, 查無安裝CMCP設備資料)
'            '   End If
'
'            If JustDoIt = Empty Then
'                JustDoIt = "-21, 查無安裝CMCP設備資料"
'            Else
'                JustDoIt = "0,成功" & vbCrLf & _
'                                    "<?xml version=""1.0"" encoding=""BIG5"" standalone=""yes"" ?>" & vbCrLf & _
'                                    "<DataSet>" & vbCrLf & _
'                                    "  <DataTable>" & vbCrLf & _
'                                    JustDoIt & _
'                                    "  </DataTable>" & vbCrLf & _
'                                    "</DataSet>"
'            End If
'
'            '   PS:
'            '   XML(設備名稱、買賣方式、安裝日期、拆除日期、設備型號、CM速率、動態IP數、CM開機日期、CM關機日期
'            '           、CM帳號啟用日期、CM帳號停用日期、設備保証金、CM序號、Port數、方案代碼、方案名稱
'            '           、綁約到期日、EBT合約編號、服務別)
'
'        End If


'Private Function GetType1SQL(strCustID As String, _
'                                                    strFaciSno As String, _
'                                                    strStartDate As String, _
'                                                    strStopDate As String) As String
'  On Error GoTo ChkErr
'
'
'    GetType1SQL = "SELECT BILLNO,ITEM,CITEMNAME,SHOULDDATE,SHOULDAMT" & _
'                                ",REALDATE,REALAMT,REALPERIOD,REALSTARTDATE,REALSTOPDATE" & _
'                                ",UCNAME,STNAME,DECODE(CANCELFLAG,1,'是','否') CANCELFLAG" & _
'                                ",SERVICETYPE,FACISNO" & _
'                                " FROM " & strOwner & "SO033" & Get_DB_Link(cn) & _
'                                " WHERE CUSTID=" & strCustID & _
'                                " AND FACISNO='" & strFaciSno & "'" & _
'                                " AND SERVICETYPE IN ('I','P')" & _
'                                " AND SHOULDDATE >= " & PrcType(strStartDate, FldDate) & _
'                                " AND SHOULDDATE < " & PrcType(strStopDate, FldDate) & "+1" & _
'                                " UNION ALL " & _
'                                "SELECT BILLNO,ITEM,CITEMNAME,SHOULDDATE,SHOULDAMT" & _
'                                ",REALDATE,REALAMT,REALPERIOD,REALSTARTDATE,REALSTOPDATE" & _
'                                ",UCNAME,STNAME,DECODE(CANCELFLAG,1,'是','否') CANCELFLAG" & _
'                                ",SERVICETYPE,FACISNO" & _
'                                " FROM " & strOwner & "SO034" & Get_DB_Link(cn) & _
'                                " WHERE CUSTID=" & strCustID & _
'                                " AND FACISNO='" & strFaciSno & "'" & _
'                                " AND SERVICETYPE IN ('I','P')" & _
'                                " AND SHOULDDATE >= " & PrcType(strStartDate, FldDate) & _
'                                " AND SHOULDDATE < " & PrcType(strStopDate, FldDate) & "+1"
'
'    '      Select BillNo, Item, CitemName, ShouldDate, ShouldAmt, RealDate, RealAmt, RealPeriod,
'    '           RealStartDate, RealStopDate, UCName, STName, CancelFlag, ServiceType From SO033
'    '           Where CustId=[CustId] And FaciSno=[FaciSno] And ServiceType In ('I','P')
'    '           And ShouldDate>=[起始日期] And ShouldDate < [截止日期]+1
'    '      Union All
'    '          Select BillNo, Item, CitemName, ShouldDate, ShouldAmt, RealDate, RealAmt, RealPeriod,
'    '           RealStartDate, RealStopDate, UCName, STName, CancelFlag, ServiceType From SO034
'    '           Where CustId=[CustId] And FaciSno=[FaciSno] And ServiceType In ('I','P')
'    '           And ShouldDate>=[起始日期] And ShouldDate < [截止日期]+1
'
'  Exit Function
'ChkErr:
'    ErrHandle "mod_Func23", "GetType1SQL"
'    GetType1SQL = "-99," & strErr
'End Function
'
'Private Function GetType2SQL(strCustID As String, _
'                                                    strFaciSno As String, _
'                                                    strStartDate As String, _
'                                                    strStopDate As String) As String
'  On Error GoTo ChkErr
'
'
'    GetType2SQL = "SELECT A.FACINAME,A.BUYNAME,A.OLDINSTDATE,A.PRDATE,A.MODELNAME" & _
'                                ",B.CMBAUDRATE,B.DYNIPCOUNT,A.CMOPENDATE,A.CMCLOSEDATE" & _
'                                ",A.ENABLEACCOUNT,A.DISABLEACCOUNT,A.DEPOSIT,A.REFACISNO" & _
'                                ",A.PORTNO,C.BPCODE,C.BPNAME,A.CONTRACTDATE,A.EBTCONTNO" & _
'                                ",A.SERVICETYPE,A.FACISNO" & _
'                                " FROM " & strOwner & "SO004" & Get_DB_Link(cn) & " A " & _
'                                " ," & strOwner & "CD078D" & Get_DB_Link(cn) & " B " & _
'                                " ," & strOwner & "SO003" & Get_DB_Link(cn) & " C " & _
'                                " WHERE A.CUSTID=" & strCustID & _
'                                " AND A.FACISNO='" & strFaciSno & "'" & _
'                                " AND A.SERVICETYPE='I' " & _
'                                " AND A.SERVICETYPE=C.SERVICETYPE" & _
'                                " AND A.FACISNO=C.FACISNO" & _
'                                " AND C.BPCODE=B.BPCODE" & _
'                                " UNION ALL " & _
'                                "SELECT FACINAME,BUYNAME,OLDINSTDATE,PRDATE,MODELNAME" & _
'                                ",CMBAUDRATE,DYNIPCOUNT,CMOPENDATE,CMCLOSEDATE" & _
'                                ",ENABLEACCOUNT,DISABLEACCOUNT,DEPOSIT,REFACISNO" & _
'                                ",PORTNO,BPCODE,BPNAME,CONTRACTDATE,EBTCONTNO" & _
'                                ",SERVICETYPE,FACISNO" & _
'                                " FROM " & strOwner & "SO004" & Get_DB_Link(cn) & _
'                                " WHERE CUSTID=" & strCustID & _
'                                " AND REFACISNO='" & strFaciSno & "'" & _
'                                " AND SERVICETYPE='P' "
'
'    '       Select a.FaciName, a.BuyName, OldInstDate,PRDate,a.ModelName,b.CMBaudRate,
'    '       b.DynIPCount,CMOpenDate,CMCloseDate,EnableAccount, DisableAccount,Deposit,
'    '       ReFaciSNo,PortNo,c.BPCode,c.BPName,ContractDate,EBTContNo,a.ServiceType
'    '       From SO004 a, CD078D b, SO003 c
'    '       Where a.ServiceType='I' And a.Custid=[Custid] And a.FaciSno=[FaciSNo] And
'    '       a.ServiceType=c.ServiceType And a.FaciSno=c.FaciSNo And c.bpcode=b.bpcode
'    '       Union All
'    '       Select FaciName, BuyName, OldInstDate,PRDate,ModelName,CMBaudRate,
'    '       DynIPCount,CMOpenDate,CMCloseDate,EnableAccount, DisableAccount,Deposit,
'    '       ReFaciSNo,PortNo,BPCode,BPName,ContractDate,EBTContNo,ServiceType
'    '       From SO004
'    '       Where ServiceType='P' And Custid=[Custid] And ReFaciSno=[FaciSNo]
'
'  Exit Function
'ChkErr:
'    ErrHandle "mod_Func23", "GetType2SQL"
'    GetType2SQL = "-99," & strErr
'End Function
'
