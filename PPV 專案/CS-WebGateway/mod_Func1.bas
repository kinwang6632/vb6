Attribute VB_Name = "mod_Func1"
Option Explicit

Private strAuthenticID As String
Private cn As Object
Private strOwner As String

'select * from so002b where facisno like '%' || (select facisno from so002b where facisno like '%''200302270069925''%') || '%'
'select * from so002b where facisno like '%' || (
'select facisno from so002b where facisno like '%''200302270069925''%' union all
'select facisno from so002b where facisno like '%''200302270069925''%' )  || '%'

' 有線電視客戶身份SMS帳號密碼認證申請
' 1、公司別、客戶編號、種類、認證ID -- (當種類=0或2時須傳認證ID) 2007/04/02 liga add (2.2版本)
' C、用戶姓名、電話(折行)
' D、DSTB設備序號、智慧卡卡號(折行)
' I、申請人、身份證、生日(折行)、(新增) CM合約編號 (折行)
' ……

' 第二行以下之參數，無限定認證的服務種類順序，依網站上所傳進來的服務不同，
' 而傳不同的參數即可，每增加一服務認證申請則往下新增一行。
' EX:
'  1,16,123
'  D , 119110074891#, 116010001031#
'  I , 王大明, J123456789, 19770301
'  I , 王美麗, H223456789, 19890607
'  C , LIGA, 3010001

' 結果碼、結果訊息字串(折行)
' 認證ID (折行)
' XML (顯示逐筆認證成功 / 失敗狀態)

' 0:  成功
' -22 : 該客編或設備資料於此系統台不存在

Public Function JustDoIt(ByRef objWebPool As Object, _
                                        ByRef strPara As String) As String
  On Error GoTo ChkErr
    Dim rs As Object
    Dim strQry As String
    Dim strSubQry As String
    Dim strCustID As String
    Dim strItem As String
    Dim intLoop As Integer
    Dim varPara As Variant
    Dim varDetail As Variant
    Dim objCnPool As Object
    Dim strFaciSeqNo() As String
    Dim strFaciSno As String
    Dim strStatus() As String
    Dim bytLoop As Byte
    Dim bytLoop2 As Byte
    Dim bytUB As Byte
    Dim bytType As Byte
    Dim blnFlag As Boolean
    Dim strOriAuthenticID As String

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
    
    varPara = Split(strPara, vbCrLf)
    
    If UBound(varPara) = 0 Then
        JustDoIt = "-99,參數資料不正確 ! 請確認 !"
        GoTo 66
    End If
    '****************************************************************************************************************
    '#3018 增加判斷傳入的個數是否正確 By Kin 2007/09/10
    If UBound(Split(varPara(0), ",")) < 41 Then
        JustDoIt = "-99,參數不足!"
        GoTo 66
    End If
    If UBound(varPara) < 2 Then
        JustDoIt = "-99,參數不足!"
        GoTo 66
    Else
        For intLoop = 1 To UBound(varPara)
            If varPara(intLoop) <> Empty Then
                Select Case Split(varPara(intLoop), ",")(0)
                    Case "C"
                        If UBound(Split(varPara(intLoop), ",")) < 2 Then JustDoIt = "-99,參數不足!": GoTo 66
                    Case "D"
                        If UBound(Split(varPara(intLoop), ",")) < 2 Then JustDoIt = "-99,參數不足!": GoTo 66
                    Case "I"
                        If UBound(Split(varPara(intLoop), ",")) < 3 Then JustDoIt = "-99,參數不足!": GoTo 66
                End Select
                
            End If
        Next intLoop
    End If
    '******************************************************************************************************************
'    If UBound(varPara) = 0 Then
'        varPara = Split(strPara, ",")
'        If UBound(varPara) = 0 Then
'            JustDoIt = "-99,參數資料不正確 ! 請確認 !"
'            GoTo 66
'        Else
'            bytComp = Val(varPara(1))
'            strCustID = varPara(2)
'            bytType = varPara(3)
'            On Error Resume Next
'            strOriAuthenticID = varPara(4)
'        End If
'    Else
'        varDetail = Split(varPara(0), ",")
'        bytComp = Val(varDetail(1))
'        strCustID = varDetail(2)
'        bytType = varDetail(3)
'        On Error Resume Next
'        strOriAuthenticID = varDetail(4)
'    End If
        
    varDetail = Split(varPara(0), ",")
    bytComp = Val(varDetail(1))
    strCustID = varDetail(2)
    bytType = varDetail(3)
    On Error Resume Next
    strOriAuthenticID = varDetail(4)
    
    On Error GoTo ChkErr
    strOwner = GetOwner(cn)
    
'    2006/05/04 Add
'    第1行參數，增加"種類"以利區別：
'    種類=1  新增認證帳號
'    種類=0  新增booking設備使用
    
    If bytType < 0 Or bytType > 2 Then
        JustDoIt = "-99,傳入參數[種類]錯誤!"
        GoTo 66
    End If
    
     ' (當種類=0或2時須傳認證ID) 2007/04/02 liga add (2.2版本)
     If bytType = 0 Or bytType = 2 Then
        If Trim(strOriAuthenticID) = Empty Then
            JustDoIt = "-99,當種類 = 0 或 2 時須傳 認證ID!"
            GoTo 66
        End If
     End If
    
    'Select count(*) as Cnt from SO001 where Compcode=[公司別] and Custid=[客編] ;
    strQry = "SELECT COUNT(*) CNT FROM " & strOwner & "SO001" & Get_DB_Link(cn) & _
                    "WHERE COMPCODE=" & bytComp & " AND CUSTID=" & strCustID
                
'    Set rs = cn.Execute(strQry)
'    If Not GetRS(cn, rs, strQry) Then
'        JustDoIt = "-99,查詢失敗! SQL:" & strQry
'        GoTo 66
'    End If
    
'    If rs.EOF Then '    If Cnt = 0 Then
    If Val(GetValue(cn, strQry) & "") = 0 Then
        JustDoIt = "-22, 該客編或設備資料於此系統台不存在" '      回傳(-22, 該客編或設備資料於此系統台不存在)
        GoTo 66
    End If '    End If
    
    bytUB = UBound(varPara)
    ReDim strStatus(Val(bytUB) - 1) As String
    ReDim strFaciSeqNo(Val(bytUB) - 1) As String
    
    bytLoop = 0
    bytLoop2 = 0
    
    For intLoop = 1 To bytUB '    Loop處理 (第二行參數開始)
        strItem = Trim(varPara(intLoop) & "")
        If strItem <> Empty Then
            varDetail = Split(varPara(intLoop), ",")
            Select Case Trim(UCase(varDetail(0)))
                Case "C" '    CASE WHEN [種類]='C' THEN
                    ' C、用戶姓名、電話(折行)
                    ' Select 0 as Flag, null as FaciSEQNo, 'CATV用戶 客編'||Custid||' 已被認證(取消)使用' as Status
                    ' from SO002B where Compcode=[公司別] and Custid=[客編] and FaciSNo like [客編] ;
                    ' strQry = "SELECT 0 AS FLAG,NULL AS FACISEQNO,'CATV用戶 客編'||CUSTID||' 已被認證(取消)使用' AS STATUS"
                    strQry = "SELECT 'CATV用戶 客編'||CUSTID||' 已被認證(取消)使用' AS STATUS" & _
                                    " FROM " & strOwner & "SO002B" & Get_DB_Link(cn) & " WHERE COMPCODE=" & bytComp & _
                                    " AND CUSTID=" & strCustID & " AND FACISNO LIKE '%''" & strCustID & "''%'"
                    'cable_liga: 結果碼、結果訊息字串(折行)
                    '認證ID (折行)
                    'XML(Flag、設備序號、認證狀態)
                    '
                    'Hammer: <Flag>值</Flag><設備序號>值</設備序號>
                    'cable_liga: Select count as Cnt from SO001 where Compcode=[公司別] and Custid=[客編] ;
                    'If Cnt = 0 Then
                    '  回傳(-22, 該客編或設備資料於此系統台不存在)
                    'End If
'                    If GetRS(cn, rs, strQry) Then ' rs > 0
                    Set rs = cn.Execute(strQry)
                    If Err = 0 Then
                        If Not rs.EOF Then ' If Status Is Not Null Then
                            ' [記錄select出來的Status]
                            strStatus(bytLoop) = MakeXML("0", "", rs!Status & "")
                            bytLoop = bytLoop + 1
                            If bytType = 2 Then
                                strQry = "SELECT A.CUSTID AS FACISEQNO" & _
                                                " FROM " & strOwner & "SO001" & Get_DB_Link(cn) & " A" & _
                                                "," & strOwner & "SO002" & Get_DB_Link(cn) & " B" & _
                                                " WHERE A.CUSTID=B.CUSTID AND A.CUSTID=" & strCustID & _
                                                " AND B.CUSTSTATUSCODE=1 AND B.SERVICETYPE='C'" & _
                                                " AND A.COMPCODE=" & bytComp & " AND" & _
                                                " (A.TEL1='" & varDetail(2) & "' OR A.TEL2='" & varDetail(2) & "' OR A.TEL3='" & varDetail(2) & "')" & _
                                                " AND A.CUSTNAME='" & varDetail(1) & "'"
                                strFaciSeqNo(bytLoop2) = "'" & GetValue(cn, strQry) & "" & "'"
                                bytLoop2 = bytLoop2 + 1
                            End If
'                            strFaciSeqNo = strFaciSeqNo & ",'" & rs!FaciSeqNo & "'"
                        Else  ' Else
                            ' Select 1 as Flag, a.CustID as FaciSEQNo, 'CATV用戶 客編'|| a.CustID||' 申請認證成功' as Status
                            ' from SO001 a, SO002 b where a.Custid=b.Custid and a.Custid=[客編] and
                            ' b.custstatuscode=1 and b.ServiceType='C' and a.Compcode=[公司別] and a.Tel1=[電話];
                            strQry = "SELECT 1 AS FLAG,A.CUSTID AS FACISEQNO,'CATV用戶 客編 ' ||A.CUSTID|| ' 申請認證成功' AS STATUS" & _
                                            " FROM " & strOwner & "SO001" & Get_DB_Link(cn) & " A" & _
                                            "," & strOwner & "SO002" & Get_DB_Link(cn) & " B" & _
                                            " WHERE A.CUSTID=B.CUSTID AND A.CUSTID=" & strCustID & _
                                            " AND B.CUSTSTATUSCODE=1 AND B.SERVICETYPE='C'" & _
                                            " AND A.COMPCODE=" & bytComp & " AND" & _
                                            " (A.TEL1='" & varDetail(2) & "' OR A.TEL2='" & varDetail(2) & "' OR A.TEL3='" & varDetail(2) & "')" & _
                                            " AND A.CUSTNAME='" & varDetail(1) & "'"
'                            If GetRS(cn, rs, strQry) Then
                            Set rs = cn.Execute(strQry)
                            If Err = 0 Then
                                If Not rs.EOF Then
                                    ' [記錄select出來的Status]
                                    blnFlag = True
                                    strStatus(bytLoop) = MakeXML("1", rs!FaciSeqNo & "", rs!Status & "")
                                    bytLoop = bytLoop + 1
                                    If rs!FaciSeqNo & "" <> Empty Then
                                        strFaciSeqNo(bytLoop2) = "'" & rs!FaciSeqNo & "'"
                                        bytLoop2 = bytLoop2 + 1
                                    End If
                                Else ' else    --2006/05/17 Add
                                    ' Select 2 as Flag, null as FaciSEQNo, '查無CATV用戶'||[用戶姓名]||'相關資料' as Status from dual;
                                    ' [記錄select出來的Status]
                                    strStatus(bytLoop) = MakeXML("2", "", "查無CATV用戶" & varDetail(1) & "相關資料")
                                    bytLoop = bytLoop + 1
                                End If
                            Else
                                Err.Clear
                                JustDoIt = "-99,查詢失敗! SQL:" & strQry
                            End If
                        End If  ' End If
                    Else
                        Err.Clear
                        JustDoIt = "-99,查詢失敗! SQL:" & strQry
                    End If
                Case "D" '    CASE WHEN [種類]='D' THEN
                    ' D、DSTB設備序號、智慧卡卡號(折行)
                    ' Select 0 as Flag, null as FaciSEQNo, '數位服務 DSTB序號'||[DSTB設備序號]||' 已被認證(取消)使用' as Status
                    ' from SO002B where Compcode=[公司別] and Custid=[客編] and FaciSNo like
                    ' (Select SeqNo from SO004 where Compcode=[公司別] and Custid=[客編]
                    ' and FaciSNo=[DSTB設備序號] and SmartCardNo=[智慧卡卡號] ) ;
                    
                    ' strQry = "SELECT 0 AS FLAG,NULL AS FACISEQNO,'數位服務 DSTB序號'||[" & varDetail(1) & "]||' 已被認證(取消)使用' AS STATUS"
                    
'                    JustDoIt = JustDoIt & "跑到D了"
                    
                    strQry = "SELECT '數位服務 DSTB序號 " & Trim(varDetail(1)) & " 已被認證(取消)使用' AS STATUS" & _
                                    " FROM " & strOwner & "SO002B" & Get_DB_Link(cn) & _
                                    " WHERE COMPCODE=" & bytComp & _
                                    " AND CUSTID=" & strCustID & " AND FACISNO LIKE '%''' || " & _
                                    " (SELECT SEQNO FROM " & strOwner & "SO004" & Get_DB_Link(cn) & _
                                    " WHERE COMPCODE=" & bytComp & " AND CUSTID=" & strCustID & _
                                    " AND FACISNO='" & varDetail(1) & "' AND SMARTCARDNO='" & varDetail(2) & "'" & _
                                    " AND INSTDATE IS NOT NULL AND PRDATE IS NULL) || '''%'"
                                    ' and InstDate is not null and PRDate is null
'                                    " AND INSTDATE IS NOT NULL AND PRDATE IS NULL" & ") || '''%'"

'                    JustDoIt = JustDoIt & vbCrLf & "SQL1 : " & strQry ' Debug

'                    If GetRS(cn, rs, strQry) Then
                    Set rs = cn.Execute(strQry)
                    If Err = 0 Then
                        If Not rs.EOF Then ' If Status Is Not Null Then
                            ' [記錄select出來的Status]
                            strStatus(bytLoop) = MakeXML("0", "", rs!Status & "")
'                            JustDoIt = JustDoIt & vbCrLf & "SQL1 : 有找到資料"  ' Debug
                            bytLoop = bytLoop + 1
                            If bytType = 2 Then
                                strQry = "SELECT SEQNO AS FACISEQNO" & _
                                                " FROM " & strOwner & "SO004" & Get_DB_Link(cn) & _
                                                " WHERE COMPCODE=" & bytComp & " AND CUSTID=" & strCustID & _
                                                " AND FACISNO='" & varDetail(1) & "' AND SMARTCARDNO='" & varDetail(2) & "'" & _
                                                " AND INSTDATE IS NOT NULL AND PRDATE IS NULL"
                                strFaciSeqNo(bytLoop2) = "'" & GetValue(cn, strQry) & "" & "'"
                                bytLoop2 = bytLoop2 + 1
'                                JustDoIt = JustDoIt & vbCrLf & "種類2" & " SeqNo : " & strFaciSeqNo(bytLoop2) ' Debug 用
                            End If
                        Else  ' Else
                            ' Select 1 as Flag, SeqNo as FaciSEQNo, '數位服務 DSTB序號' ||FaciSNo||' 申請認證成功' as Status
                            ' from SO004 where Compcode=[公司別] and Custid=[客編]
                            ' and FaciSNo=[DSTB設備序號] and SmartCardNo=[智慧卡卡號] ;
                            
'                            JustDoIt = JustDoIt & vbCrLf & "SQL1 : 沒找到資料"  ' Debug
                            
                            strQry = "SELECT 1 AS FLAG,SEQNO AS FACISEQNO,'數位服務 DSTB序號 ' ||FACISNO||' 申請認證成功' AS STATUS" & _
                                            " FROM " & strOwner & "SO004" & Get_DB_Link(cn) & _
                                            " WHERE COMPCODE=" & bytComp & " AND CUSTID=" & strCustID & _
                                            " AND FACISNO='" & varDetail(1) & "' AND SMARTCARDNO='" & varDetail(2) & "'" & _
                                            " AND INSTDATE IS NOT NULL AND PRDATE IS NULL"
'                            If GetRS(cn, rs, strQry) Then
                            
'                            JustDoIt = JustDoIt & vbCrLf & "SQL2 : " & strQry  ' Debug
                            
                            Set rs = cn.Execute(strQry)
                            If Err = 0 Then
                                If Not rs.EOF Then
                                    ' [記錄select出來的Status]
'                                    JustDoIt = JustDoIt & vbCrLf & "SQL2 : 有找到資料 FaciSeqNo : " & rs!FaciSeqNo & ""   ' Debug
                                    blnFlag = True
                                    strStatus(bytLoop) = MakeXML("1", rs!FaciSeqNo & "", rs!Status & "")
                                    bytLoop = bytLoop + 1
                                    If rs!FaciSeqNo & "" <> Empty Then
                                        strFaciSeqNo(bytLoop2) = "'" & rs!FaciSeqNo & "'"
                                        bytLoop2 = bytLoop2 + 1
                                    End If
                                Else ' else     --2006/05/17 Add
                                    ' Select 2 as Flag, null as FaciSEQNo, '查無數位服務 DSTB序號'||[DSTB設備序號]||'相關資料' as Status from dual;
                                    ' [記錄select出來的Status]
                                    
'                                    JustDoIt = JustDoIt & vbCrLf & "SQL2 : 沒找到資料"   ' Debug
                                    
                                    strStatus(bytLoop) = MakeXML("2", "", "查無數位服務 DSTB序號" & varDetail(1) & "相關資料")
                                    bytLoop = bytLoop + 1
                                End If
                            Else
                                Err.Clear
                                JustDoIt = "-99,查詢失敗! SQL:" & strQry
                            End If
                        End If  ' End If
                    Else
                        Err.Clear
                        JustDoIt = "-99,查詢失敗! SQL:" & strQry
                    End If
                Case "I" '    CASE WHEN [種類]='I' THEN
                    ' I、申請人、身份證、生日(折行)
                    ' Select 0 as Flag, null as FaciSEQNO, '寬頻服務 '||[申請人]|| '申請之CM序號已被認證(取消)使用' as Status
                    ' from SO002B where Compcode=[公司別] and Custid=[客編] and FaciSNo like
                    ' (Select SeqNo from SO004 where Compcode=[公司別] and Custid=[客編]
                    ' and DeclarantName=[申請人] and ID=[身份證] and Birthday=[生日]) ;
                    ' 後來增加 " and EBTContNo=[CM合約編號] "
                    ' strQry = "SELECT 0 AS FLAG,NULL AS FACISEQNO,'寬頻服務 '||[申請人]|| '申請之CM序號已被認證(取消)使用' AS STATUS"
                    strQry = "SELECT '寬頻服務 " & varDetail(1) & " 申請之CM序號已被認證(取消)使用' AS STATUS" & _
                                    " FROM " & strOwner & "SO002B" & Get_DB_Link(cn) & _
                                    " WHERE COMPCODE=" & bytComp & _
                                    " AND CUSTID=" & strCustID & " AND FACISNO LIKE '%''' || " & _
                                    " (SELECT SEQNO FROM " & strOwner & "SO004" & Get_DB_Link(cn) & _
                                    " WHERE COMPCODE=" & bytComp & " AND CUSTID=" & strCustID & _
                                    " AND DECLARANTNAME='" & varDetail(1) & "' AND ID='" & varDetail(2) & "'" & _
                                    " AND BIRTHDAY=" & PrcType(varDetail(3), FldDate) & _
                                    " AND EBTCONTNO='" & varDetail(4) & "'" & _
                                    " AND INSTDATE IS NOT NULL AND PRDATE IS NULL) || '''%'"
'                                    " AND INSTDATE IS NOT NULL AND PRDATE IS NULL" & ") || '''%'"
'                    If GetRS(cn, rs, strQry) Then
                    Set rs = cn.Execute(strQry)
                    If Err = 0 Then
                        If Not rs.EOF Then ' If Status Is Not Null Then
                            ' [記錄select出來的Status]
                            strStatus(bytLoop) = MakeXML("0", "", rs!Status & "")
                            bytLoop = bytLoop + 1
                            If bytType = 2 Then
                                strQry = "SELECT SEQNO AS FACISEQNO" & _
                                                " FROM " & strOwner & "SO004" & Get_DB_Link(cn) & _
                                                " WHERE COMPCODE=" & bytComp & " AND CUSTID=" & strCustID & _
                                                " AND DECLARANTNAME='" & varDetail(1) & "' AND ID='" & varDetail(2) & "'" & _
                                                " AND BIRTHDAY=" & PrcType(varDetail(3), FldDate) & _
                                                " AND INSTDATE IS NOT NULL AND PRDATE IS NULL" & _
                                                " AND EBTCONTNO='" & varDetail(4) & "'"
                                strFaciSeqNo(bytLoop2) = "'" & GetValue(cn, strQry) & "" & "'"
                                bytLoop2 = bytLoop2 + 1
                            End If
                        Else  ' Else
                            ' Select 1 as Flag, SeqNo as FaciSEQNo , '寬頻服務 CM序號'||FaciSNo||' 申請認證成功' as Status
                            ' from SO004 where Compcode=[公司別] and Custid=[客編]
                            ' and DeclarantName=[申請人] and ID=[身份證] and Birthday=[生日] ;
                            strQry = "SELECT 1 AS FLAG,SEQNO AS FACISEQNO,'寬頻服務 CM序號 '|| FACISNO ||' 申請認證成功' AS STATUS" & _
                                            " FROM " & strOwner & "SO004" & Get_DB_Link(cn) & _
                                            " WHERE COMPCODE=" & bytComp & " AND CUSTID=" & strCustID & _
                                            " AND DECLARANTNAME='" & varDetail(1) & "' AND ID='" & varDetail(2) & "'" & _
                                            " AND BIRTHDAY=" & PrcType(varDetail(3), FldDate) & _
                                            " AND INSTDATE IS NOT NULL AND PRDATE IS NULL" & _
                                            " AND EBTCONTNO='" & varDetail(4) & "'"
'                            If GetRS(cn, rs, strQry) Then
                            Set rs = cn.Execute(strQry)
                            If Err = 0 Then
                                If Not rs.EOF Then
                                    ' [記錄select出來的Status]
                                    blnFlag = True
                                    strStatus(bytLoop) = MakeXML("1", rs!FaciSeqNo & "", rs!Status & "")
                                    bytLoop = bytLoop + 1
                                    If rs!FaciSeqNo & "" <> Empty Then
                                        strFaciSeqNo(bytLoop2) = "'" & rs!FaciSeqNo & "'"
                                        bytLoop2 = bytLoop2 + 1
                                    End If
                                Else ' else     --2006/05/17 Add
                                    ' Select 2 as Flag, null as FaciSEQNo, '查無寬頻服務 CM用戶'||[申請人]||'相關資料' as Status from dual;
                                    ' [記錄select出來的Status]
                                    strStatus(bytLoop) = MakeXML("2", "", "查無寬頻服務 CM用戶" & varDetail(1) & "相關資料")
                                    bytLoop = bytLoop + 1
                                End If
                            Else
                                Err.Clear
                                JustDoIt = "-99,查詢失敗! SQL:" & strQry
                            End If
                        End If  ' End If
                    Else
                        Err.Clear
                        JustDoIt = "-99,查詢失敗! SQL:" & strQry
                    End If
            End Select
        End If
    Next
    
    If Len(strStatus(0)) = 0 And Len(strFaciSeqNo(0)) = 0 Then
        JustDoIt = "-22, 該客編或設備資料於此系統台不存在"
        GoTo 66
'    Else
'        If UBound(strFaciSeqNo) = 0 And Len(strFaciSeqNo(0)) = 0 Then strFaciSeqNo(0) = vbCrLf
'        If UBound(strStatus) = 0 And Len(strStatus(0)) = 0 Then strStatus(0) = vbCrLf
    End If
    
    If bytLoop2 > 0 Then ReDim Preserve strFaciSeqNo(bytLoop2 - 1) As String
    If bytLoop > 0 Then ReDim Preserve strStatus(bytLoop - 1) As String
    
    If Not blnFlag Then
        If bytType = 2 Then ' Else If RecordSet(Flag=0的)>0 and [種類]=2 then
            GoTo 99
        Else
            JustDoIt = "-22, 該客編或設備資料於此系統台不存在" & vbCrLf & _
                                strOriAuthenticID & vbCrLf & _
                                "<?xml version=""1.0"" encoding=""BIG5"" standalone=""yes"" ?>" & vbCrLf & _
                                "<DataSet>" & vbCrLf & _
                                "  <DataTable>" & vbCrLf & _
                                Join(strStatus, vbCrLf) & vbCrLf & _
                                "  </DataTable>" & vbCrLf & _
                                "</DataSet>" & vbCrLf
            GoTo 66
        End If
    End If
    
    If bytType = 1 Then ' If Recordset(Flag=1的)>0 and [種類]=1 then
'        If blnFlag Then

        strFaciSno = Join(strFaciSeqNo, ",")
        strFaciSno = Replace(strFaciSno, ",,", ",", 1)
        If Left(strFaciSno, 1) = "," Then strFaciSno = Mid(strFaciSno, 2)
        If Right(strFaciSno, 1) = "," Then strFaciSno = Left(strFaciSno, Len(strFaciSno) - 1)
        
'        JustDoIt = JustDoIt & vbCrLf & "判斷 Flag 前 , strFaciSno : " & strFaciSno ' Debug
        
        If Insert_2_SO002B(CStr(varPara(0)), strFaciSno) Then ' .[Web額外傳入之資料] 資料
    '           .將select出來之[FaciSeqNo]組串起來並insert SO002B.FaciSNo
            JustDoIt = "0,成功" & vbCrLf & strAuthenticID & vbCrLf & _
                                "<?xml version=""1.0"" encoding=""BIG5"" standalone=""yes"" ?>" & vbCrLf & _
                                "<DataSet>" & vbCrLf & _
                                "  <DataTable>" & vbCrLf & _
                                Join(strStatus, vbCrLf) & vbCrLf & _
                                "  </DataTable>" & vbCrLf & _
                                "</DataSet>" & vbCrLf
'           .取SO041.AccountingBefore , AccountingBalance參數值回填SO002B.AccountingBefore, AccountingBalance欄位
'               回傳(0,成功,AuthenticId值)及依select出來之資料產生XML檔
        Else
            If strErr = Empty Then strErr = "新增客戶資料至 [客戶認證資訊檔] 失敗!"
            JustDoIt = "-99," & strErr
        End If
'        Else
'            JustDoIt = "0,成功" & vbCrLf & "" & vbCrLf & _
'                                "<?xml version=""1.0"" encoding=""BIG5"" standalone=""yes"" ?>" & vbCrLf & _
'                                "<DataSet>" & vbCrLf & _
'                                "  <DataTable>" & vbCrLf & _
'                                Join(strStatus, vbCrLf) & vbCrLf & _
'                                "  </DataTable>" & vbCrLf & _
'                                "</DataSet>" & vbCrLf
'        End If
    Else '    Else
    
99:
        strFaciSno = GetValue(cn, "SELECT FACISNO FROM " & strOwner & "SO002B" & Get_DB_Link(cn) & _
                                                " WHERE AUTHENTICID='" & strOriAuthenticID & "'") & ""
                                                
        If bytType = 0 Then '    If RecordSet(Flag=1的)>0 and [種類]=0 then
            '     ．將select出來之[FaciSeqNo]組串起來並update SO002B.FaciSNo
            '        update SO002B set FaciSNo=FaciSNo || [FaciSeqNo] (append上去，而非覆蓋該欄位)
            '        where custid=[客編] and authenticid=[認證ID];
'            JustDoIt = JustDoIt & vbCrLf & "種類0"  ' Debug 用
            strFaciSno = IIf(strFaciSno <> Empty, strFaciSno & ",", "") & Join(strFaciSeqNo, ",")
        ElseIf bytType = 2 Then
            '    Else If RecordSet(Flag=1的)>0 and [種類]=2 then
            '     ．將select出來之[FaciSeqNo]將其replace為空值，並update回SO002B.FaciSNo
            '      回傳(0,成功,AuthenticId值)及依select出來之資料產生XML檔
'            JustDoIt = JustDoIt & vbCrLf & "99種類2"  ' Debug 用
'            JustDoIt = JustDoIt & vbCrLf & "維度 : " & UBound(strFaciSeqNo) ' Debug 用
'            JustDoIt = JustDoIt & vbCrLf & strFaciSno & "-" & Join(strFaciSeqNo, ",")  ' Debug 用
            strFaciSno = Minus2str(strFaciSno, Join(strFaciSeqNo, ","))
        End If
        
        strFaciSno = Replace(strFaciSno, ",,", ",", 1)
        If Left(strFaciSno, 1) = "," Then strFaciSno = Mid(strFaciSno, 2)
        If Right(strFaciSno, 1) = "," Then strFaciSno = Left(strFaciSno, Len(strFaciSno) - 1)
        strFaciSno = Replace(strFaciSno, "'", "''")
        
'        JustDoIt = JustDoIt & vbCrLf & strFaciSno ' Debug 用
        
         ' Debug 用
'        JustDoIt = JustDoIt & vbCrLf & _
                            "UPDATE " & strOwner & "SO002B" & Get_DB_Link(cn) & _
                            " SET FACISNO='" & strFaciSno & "' WHERE AUTHENTICID='" & strOriAuthenticID & "'"
        
        cn.Execute "UPDATE " & strOwner & "SO002B" & Get_DB_Link(cn) & _
                            " SET FACISNO='" & strFaciSno & "' WHERE AUTHENTICID='" & strOriAuthenticID & "'"
        
'        回傳(0,成功,AuthenticId值)及依select出來之資料產生XML檔
        JustDoIt = "0,成功" & vbCrLf & strOriAuthenticID & vbCrLf & _
                            "<?xml version=""1.0"" encoding=""BIG5"" standalone=""yes"" ?>" & vbCrLf & _
                            "<DataSet>" & vbCrLf & _
                            "  <DataTable>" & vbCrLf & _
                            Join(strStatus, vbCrLf) & vbCrLf & _
                            "  </DataTable>" & vbCrLf & _
                            "</DataSet>" & vbCrLf
    End If

'    Else
'        回傳(-22, 該客編或設備資料於此系統台不存在)
'    End If
    
66:
    On Error Resume Next
    Rlx rs
    Rlx strQry
    Rlx blnFlag
    Rlx strSubQry
    Rlx strItem
    Rlx strCustID
    Rlx intLoop
    Rlx varPara
    Rlx varDetail
    Rlx strOwner
    Rlx strAuthenticID
    Rlx strOriAuthenticID
    Rlx bytUB
    Rlx bytLoop
    Rlx bytType
    Rlx strFaciSno
    Erase strStatus
    Erase strFaciSeqNo
    Set cn = Nothing
    Set objCnPool = Nothing
  Exit Function
ChkErr:
    ErrHandle "mod_Func1", "JustDoIt"
    JustDoIt = "-99," & strErr
End Function

Private Function MakeXML(ByVal strFlag As String, _
                                    ByVal strFaciSno As String, _
                                    ByVal strStatus As String) As String
  On Error GoTo ChkErr
    MakeXML = "    <DataRow FLAG=""" & strFlag & """" & _
                            " FACISEQNO=""" & strFaciSno & """" & _
                            " STATUS=""" & strStatus & """/>"
  Exit Function
ChkErr:
    ErrHandle "mod_Func1", "MakeXML"
End Function

Private Function GetAccPara(strAccBefore As String, strAccBalance As String) As Boolean
  On Error GoTo ChkErr
    Dim rs041 As Object
    GetAccPara = False
'    If GetRS(cn, rs041, "SELECT ACCOUNTINGBEFORE,ACCOUNTINGBALANCE FROM " & _
                                    strOwner & "SO043" & Get_DB_Link(cn) & " WHERE COMPCODE=" & bytComp) Then
    
    Set rs041 = cn.Execute("SELECT ACCOUNTINGBEFORE,ACCOUNTINGBALANCE FROM " & _
                                    strOwner & "SO043" & Get_DB_Link(cn) & " WHERE COMPCODE=" & bytComp)
    If Err.Number = 0 Then
        With rs041
            If Not .EOF Then
                strAccBefore = !AccountingBefore & ""
                strAccBalance = !ACCOUNTINGBALANCE & ""
                GetAccPara = True
            End If
        End With
    Else
        Err.Clear
    End If
    On Error Resume Next
    Rlx rs041
  Exit Function
ChkErr:
    ErrHandle "mod_Func1", "GetAccPara"
End Function

' 新增客戶資料至 [客戶認證資訊檔] SO002B
Private Function Insert_2_SO002B(ByVal strPara As String, strFaciSno As String) As Boolean
  On Error GoTo ChkErr
    Dim intRet As Integer
    Dim strSQL As String
    Dim strInvTitle As String
    Dim strInvoiceType As String
    Dim strInvAddress As String
    Dim strInvNo As String
    Dim strAccBefore As String
    Dim strAccBalance As String
    Dim varP As Variant
    
    varP = Split(strPara, ",")
    
    strAuthenticID = GetID '        新增so002b(客戶認證資訊檔) 並產生AuthenticId值
    
    GetAccPara strAccBefore, strAccBalance
    
    On Error Resume Next
    strInvTitle = varP(38) & ""
    strInvoiceType = varP(39) & ""
    If strInvoiceType = Empty Then strInvoiceType = "NULL"
    strInvAddress = varP(40) & ""
    strInvNo = varP(41) & ""
 
    ' 1、公司別、客戶編號、種類、認證ID
    
'   客服網站會員代號、會員名稱、身分證字號、出生日期-年、出生日期-月、出生日期-日、
'   聯絡電話(宅) 、聯絡電話(公)、行動電話、性別、通訊地址-縣市、通訊地址-鄉鎮市區、通訊地址-路段街巷弄巷號樓、
'   Email、客服網站登入帳號、PPV訂購密碼、客服網站會員密碼、密碼提示問題、密碼提示答案、學歷、
'   願意收到有線電視服務及頻道節目等相關訊息、婚姻狀況、同住人數、同住成員、職業別、個人年收入、
'   身份別、帳號狀態、申請日期時間、啟用日期、最後登入日期、最後異動日期、系統台代碼、發票抬頭、發票種類(二聯或三聯)、發票地址、統一編號(當發票種類為三聯時須為必填)
    
    On Error GoTo ChkErr
    strSQL = "INSERT INTO " & strOwner & "SO002B" & Get_DB_Link(cn) & _
                    "(MEMBERID,NAME,ID,BIRTHDAYYYYY,BIRTHDAYMM,BIRTHDAYDD," & _
                    "TEL,TEL2,MOBILE,SEXTUAL,CITY,TOWN,ADDRESS,EMAIL,LOGINID," & _
                    "PPVORDERPWD,LOGINPASSWORD,HINTNAME,HINTANSWER,EDUCATION," & _
                    "MESSAGE,MARRIED,BODYCOUNT,MEMBERAGE,JOB,RENTE,CLASSID," & _
                    "STATUS,CREATEDATE,USEDATE,LASTLOGINDATE,UPDTIME," & _
                    "COMPCODE,CUSTID,FACISNO,AUTHENTICID," & _
                    "INVTITLE,INVOICETYPE,INVADDRESS,INVNO," & _
                    "ACCOUNTINGBEFORE,ACCOUNTINGBALANCE) VALUES (" & _
                    PrcType(varP(5)) & "," & PrcType(varP(6)) & "," & PrcType(varP(7)) & "," & _
                    PrcType(varP(8)) & "," & PrcType(varP(9)) & "," & PrcType(varP(10)) & "," & _
                    PrcType(varP(11)) & "," & PrcType(varP(12)) & "," & PrcType(varP(13)) & "," & _
                    PrcType(varP(14)) & "," & PrcType(varP(15)) & "," & PrcType(varP(16)) & "," & _
                    PrcType(varP(17)) & "," & PrcType(varP(18)) & "," & PrcType(varP(19)) & "," & _
                    PrcType(varP(20)) & "," & PrcType(varP(21)) & "," & PrcType(varP(22)) & "," & _
                    PrcType(varP(23)) & "," & PrcType(varP(24)) & "," & PrcType(varP(25)) & "," & _
                    PrcType(varP(26)) & "," & PrcType(varP(27), FldNum) & "," & PrcType(varP(28)) & "," & _
                    PrcType(varP(29)) & "," & PrcType(varP(30)) & "," & PrcType(varP(31)) & "," & _
                    PrcType(varP(32)) & "," & PrcType(varP(33), FldDate) & "," & _
                    PrcType(varP(34), FldDate) & "," & PrcType(varP(35), FldDate) & "," & _
                    PrcType(varP(36)) & "," & varP(37) & "," & varP(2) & ",'" & Replace(strFaciSno, "'", "''", 1) & "','" & strAuthenticID & "'," & _
                    "'" & strInvTitle & "'," & strInvoiceType & ",'" & strInvAddress & "','" & strInvNo & "'," & _
                    "'" & strAccBefore & "','" & strAccBalance & "')"
    
    cn.Execute strSQL, intRet
    Insert_2_SO002B = (intRet > 0)
    
    On Error Resume Next
    strSQL = Empty
    Rlx intRet
    Rlx strSQL
    Rlx strInvNo
    Rlx strInvTitle
    Rlx strAccBefore
    Rlx strAccBalance
    Rlx strInvAddress
    Rlx strInvoiceType
  Exit Function
ChkErr:
    ErrHandle "mod_Func1", "Insert_2_SO002B"
    Insert_2_SO002B = False
End Function

' 取得 AuthenticID
Private Function GetID() As String
  On Error GoTo ChkErr
    GetID = GetValue(cn, "SELECT " & strOwner & "S_SO002B_AUTHENTICID.NEXTVAL" & Get_DB_Link(cn) & " FROM DUAL")
'   PS : so002b Sequence Object: S_SO002B_AuthenticId
'   認證編號的產生!(公司別3碼+流水號7碼) 右靠左補零 , Ex: 公司別=1, 流水號為123; 則認證編號='0010000123'
    GetID = Right("0000000" & GetID, 7)
    GetID = Right("000" & bytComp, 3) & GetID
  Exit Function
ChkErr:
    ErrHandle "mod_Func1", "GetID"
End Function



'Private strAuthenticID As String
'Private cn As Object
'Private strOwner As String
'
'' 有線電視客戶身份SMS帳號密碼認證申請
'
'' 1、公司別、客戶編號、STB序號
'
'' 舊的參數 : 會員代號、會員名稱、身分證字號、出生年月日、聯絡電話、手機號碼、性別、聯絡地址、Email、PPV後付訂購密碼、登入密碼
'
'' 新的參數 : 客服網站會員代號、會員名稱、身分證字號、出生日期-年、出生日期-月、出生日期-日、
''   聯絡電話(宅) 、聯絡電話(公)、行動電話、性別、通訊地址-縣市、通訊地址-鄉鎮市區、通訊地址-路段街巷弄巷號樓、
''   Email、客服網站登入帳號、PPV訂購密碼、客服網站會員密碼、密碼提示問題、密碼提示答案、學歷、
''   願意收到有線電視服務及頻道節目等相關訊息、婚姻狀況、同住人數、同住成員、職業別、個人年收入、
''   身份別、帳號狀態、申請日期時間、啟用日期、最後登入日期、最後異動日期、系統台代碼
'
'' 結果碼、結果訊息字串(折行)
'' 認證ID
'
'' 0:  成功
'' -1 : 帳戶認證錯誤
'' -3 : 此機上盒已停用
'
'Public Function JustDoIt(ByRef objWebPool As Object, _
'                                        ByRef strPara As String) As String
'  On Error GoTo ChkErr
'    Dim strQry As String
'    Dim varPara As Variant
'    Dim rsSO004 As Object
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
'    varPara = Split(strPara, ",")
'    bytComp = Val(varPara(1))
'    strOwner = GetOwner(cn)
'
'    'select prdate from so004 where CompCode=[公司別] and CustId=[客戶編號] and
'    ' FaciSNO=[STB序號] order by Instdate desc(取第一筆)
'
'    strQry = "SELECT PRDATE FROM " & strOwner & "SO004" & Get_DB_Link(cn) & _
'                    "WHERE COMPCODE=" & bytComp & _
'                    " AND CUSTID=" & varPara(2) & _
'                    " AND FACISNO='" & varPara(3) & "'" & _
'                    " ORDER BY INSTDATE DESC"
'
'    Set rsSO004 = cn.Execute(strQry)
'
'    '   If RecordCount = 0 Then
'    '       回傳(-1,帳戶認證錯誤)
'    '   else if prdate is not null then
'    '       回傳(-3,此機上盒已停用)
'    '   Else
'    '       新增so002b(客戶認證資訊檔)[Web額外傳入之資料]資料並產生AuthenticId值
'    '       回傳(0,成功,AuthenticId值)
'    '   End If
'
'    If rsSO004.RecordCount <= 0 Then
'        JustDoIt = "-1,帳戶認證錯誤"
'    ElseIf Not IsNull(rsSO004("PRDATE")) Then
'        JustDoIt = "-3,此機上盒已停用"
'    Else
'        If Insert_2_SO002B(varPara) Then
'            JustDoIt = "0,成功" & vbCrLf & strAuthenticID
'        Else
'            If strErr = Empty Then strErr = "新增客戶資料至 [客戶認證資訊檔] 失敗!"
'            JustDoIt = "-99," & strErr
'        End If
'    End If
'
'66:
'    On Error Resume Next
'    Rlx strQry
'    Rlx varPara
'    Rlx rsSO004
'    Rlx strOwner
'    Rlx strAuthenticID
'    Set cn = Nothing
'    Set objCnPool = Nothing
'  Exit Function
'ChkErr:
'    ErrHandle "mod_Func1", "JustDoIt"
'    JustDoIt = "-99," & strErr
'End Function
'
'' 新增客戶資料至 [客戶認證資訊檔] SO002B
'Private Function Insert_2_SO002B(ByRef varPara As Variant) As Boolean
'  On Error GoTo ChkErr
'    'Dim cmd As Object
'    Dim intRet As Integer
''    Set cmd = objCnPool.GetCommand
''    With cmd
''         Set .ActiveConnection = cn
''        .CommandType = 1
''        .CommandText = "INSERT INTO " & strOwner & "SO002B " & _
''                                "(AUTHENTICID,NAME,TEL,MOBILE,BIRTHDAY,ID,SEXTUAL," & _
''                                "ADDRESS,COMPCODE,CUSTID,FACISNO,EMAIL," & _
''                                "MEMBERID,PPVORDERPWD,LOGINPASSWORD)" & _
''                                " VALUES (?,?,?,?,TO_DATE(?, 'YYYY/MM/DD'),?,?,?,?,?,?,?,?,?,?)"
''         strAuthenticID = GetID
''        .Execute intRet, Array(strAuthenticID, varPara(5), varPara(8), varPara(9), varPara(7), _
''                                varPara(6), varPara(10), varPara(11), bytComp, varPara(2), _
''                                varPara(3), varPara(12), varPara(4), varPara(13), varPara(14)), 1
''        Insert_2_SO002B = (intRet > 0)
''         Set .ActiveConnection = Nothing
''    End With
'
''   新的參數 : 客服網站會員代號、會員名稱、身分證字號、出生日期-年、出生日期-月、出生日期-日、
''   聯絡電話(宅) 、聯絡電話(公)、行動電話、性別、通訊地址-縣市、通訊地址-鄉鎮市區、通訊地址-路段街巷弄巷號樓、
''   Email、客服網站登入帳號、PPV訂購密碼、客服網站會員密碼、密碼提示問題、密碼提示答案、學歷、
''   願意收到有線電視服務及頻道節目等相關訊息、婚姻狀況、同住人數、同住成員、職業別、個人年收入、
''   身份別、帳號狀態、申請日期時間、啟用日期、最後登入日期、最後異動日期、系統台代碼
'
'    Dim strSQL As String
'    strAuthenticID = GetID
'
'
''   3．結算前交易限額(ACCOUNTINGBEFORE)、結算後交易限額( ACCOUNTINGBALANCE)，現回傳皆為0，
''       此於SO043已有增加這2個參數，應於意藍回傳時，亦需至SO043讀取這2個參數並回填SO002B。
''    txtAccBefore = Val(GetSystemParaItem("AccountingBefore", gCompCode, gServiceType, "SO043", , True) & Empty)
''    txtAccBalance = Val(GetSystemParaItem("AccountingBalance", gCompCode, gServiceType, "SO043", , True) & Empty)
'
'    strSQL = "INSERT INTO " & strOwner & "SO002B" & Get_DB_Link(cn) & _
'                                "(MEMBERID,NAME,ID,BIRTHDAYYYYY,BIRTHDAYMM,BIRTHDAYDD," & _
'                                "TEL,TEL2,MOBILE,SEXTUAL,CITY,TOWN,ADDRESS,EMAIL,LOGINID," & _
'                                "PPVORDERPWD,LOGINPASSWORD,HINTNAME,HINTANSWER,EDUCATION," & _
'                                "MESSAGE,MARRIED,BODYCOUNT,MEMBERAGE,JOB,RENTE,CLASSID," & _
'                                "STATUS,CREATEDATE,USEDATE,LASTLOGINDATE,UPDTIME," & _
'                                "COMPCODE,CUSTID,FACISNO,AUTHENTICID) VALUES (" & _
'                                PrcType(varPara(4)) & "," & PrcType(varPara(5)) & "," & PrcType(varPara(6)) & "," & _
'                                PrcType(varPara(7)) & "," & PrcType(varPara(8)) & "," & PrcType(varPara(9)) & "," & _
'                                PrcType(varPara(10)) & "," & PrcType(varPara(11)) & "," & PrcType(varPara(12)) & "," & _
'                                PrcType(varPara(13)) & "," & PrcType(varPara(14)) & "," & PrcType(varPara(15)) & "," & _
'                                PrcType(varPara(16)) & "," & PrcType(varPara(17)) & "," & PrcType(varPara(18)) & "," & _
'                                PrcType(varPara(19)) & "," & PrcType(varPara(20)) & "," & PrcType(varPara(21)) & "," & _
'                                PrcType(varPara(22)) & "," & PrcType(varPara(23)) & "," & PrcType(varPara(24)) & "," & _
'                                PrcType(varPara(25)) & "," & PrcType(varPara(26), FldNum) & "," & PrcType(varPara(27)) & "," & _
'                                PrcType(varPara(28)) & "," & PrcType(varPara(29)) & "," & PrcType(varPara(30)) & "," & _
'                                PrcType(varPara(31)) & "," & PrcType(varPara(32), FldDate) & "," & _
'                                PrcType(varPara(33), FldDate) & "," & PrcType(varPara(34), FldDate) & "," & _
'                                PrcType(varPara(35)) & "," & bytComp & "," & varPara(2) & "," & _
'                                "'" & varPara(3) & "','" & strAuthenticID & "')"
'
''    strSQL = "INSERT INTO " & strOwner & "SO002B " & _
''                                "(AUTHENTICID,NAME,TEL,MOBILE,BIRTHDAY,ID,SEXTUAL," & _
''                                "ADDRESS,COMPCODE,CUSTID,FACISNO,EMAIL," & _
''                                "MEMBERID,PPVORDERPWD,LOGINPASSWORD)" & _
''                                " VALUES ('" & strAuthenticID & "','" & varPara(5) & "'," & _
''                                "'" & varPara(8) & "','" & varPara(9) & "'," & _
''                                "TO_DATE('" & varPara(7) & "', 'YYYY/MM/DD')," & _
''                                "'" & varPara(6) & "'," & varPara(10) & "," & _
''                                "'" & varPara(11) & "'," & bytComp & "," & _
''                                varPara(2) & ",'" & varPara(3) & "'," & _
''                                "'" & varPara(12) & "','" & varPara(4) & "'," & _
''                                "'" & varPara(13) & "','" & varPara(14) & "')"
'    cn.Execute strSQL, intRet
'    Insert_2_SO002B = (intRet > 0)
'    strSQL = Empty
'  Exit Function
'ChkErr:
'    ErrHandle "mod_Func1", "Insert_2_SO002B"
'    Insert_2_SO002B = False
'End Function
'
'' 取得 AuthenticID
'Private Function GetID() As String
'  On Error GoTo ChkErr
''    GetID = RPxx(cn.Execute("SELECT " & strOwner & "S_SO002B_AUTHENTICID.NEXTVAL AUTHENTICID FROM DUAL").GetString(2, 1, "", "", "") & "")
'    GetID = GetValue(cn, "SELECT " & strOwner & "S_SO002B_AUTHENTICID.NEXTVAL AUTHENTICID FROM DUAL")
'
''   PS : so002b Sequence Object: S_SO002B_AuthenticId
''   認證編號的產生!(公司別3碼+流水號7碼) 右靠左補零
''   Ex: 公司別=1, 流水號為123; 則認證編號='0010000123'
'
'    GetID = Right("0000000" & GetID, 7)
'    GetID = Right("000" & bytComp, 3) & GetID
'  Exit Function
'ChkErr:
'    ErrHandle "mod_Func1", "GetID"
'End Function
'
                    
'                    strSubQry = "SELECT SEQNO FROM " & strOwner & "SO004" & Get_DB_Link(cn) & _
'                                        " WHERE COMPCODE=" & bytComp & " AND CUSTID=" & strCustID & _
'                                        " AND FACISNO='" & varDetail(1) & "' AND SMARTCARDNO='" & varDetail(2) & "'" & _
'                                        " AND INSTDATE IS NOT NULL AND PRDATE IS NULL"
'
'                    On Error Resume Next
'                    strSubQry = cn.Execute(strSubQry).GetString(2, , "", ",", "")
'                    If Err = 3021 Then
'                        JustDoIt = "-99,查詢失敗! SQL:" & strSubQry
'                        GoTo 66
'                    End If
'                    strSubQry = Left(strSubQry, Len(strSubQry) - 1)


