Attribute VB_Name = "RunExeCmd"
Option Explicit
Private Const frmName As String = "RunExeCmd"
Public fOrderDate1 As String
Public fOrderDate2 As String
Public fCustId As String
Public fOrderWayCode As String
Public fOrderEn As String
Public fCanSelect As Boolean
Private rsSO033Pvod As New ADODB.Recordset
Private rsSO004 As New ADODB.Recordset
Private rsSO563 As New ADODB.Recordset
Public cnCom As ADODB.Connection
Public fBillNo As String
Public fBillNoItem As String
Public fBillNoAndItem As String
Public fSTBSeqNos As String
Public fSTBSeqNo As String
Public fShowHistory As Boolean
Public fCitemCode As String
Public fProductType As Integer '0.包月 1.論片制
Private Const strComSectionName As String = "PVODCOM"
Public fSO033PvodField As String
Public Sub InitiComCn()
  On Error GoTo ChkErr:
    If cnCom Is Nothing Then
            Dim ComConString As String
            Dim pstrDataSource As String
            Dim pstrUserId As String
            Dim pstrPassword As String
            Dim SysPath As String
            
            SysPath = garyGi(11)
            pstrDataSource = ReadFROMINI(SysPath, strComSectionName, "1", True)
            pstrUserId = ReadFROMINI(SysPath, strComSectionName, "2", True)
            pstrPassword = ReadFROMINI(SysPath, strComSectionName, "3", True)
            ComConString = "Provider=MSDAORA.1;" & _
                        "Password=" & pstrPassword & ";" & _
                        "User ID=" & pstrUserId & ";" & _
                        "Data Source=" & pstrDataSource & ";" & _
                        "Persist Security Info=True"
                        
            Set cnCom = New ADODB.Connection
            cnCom.ConnectionString = ComConString
            cnCom.CursorLocation = adUseClient
            cnCom.Open
            
    End If
    
    Exit Sub
ChkErr:
    Call ErrSub(frmName, "InitiComCn")
End Sub
'取出裝罝地點名稱
Public Function GetInitPlaceName(ByVal aStbSeqNo As String, ByVal aCustId As String) As String
  On Error GoTo ChkErr
    Dim aQry As String
    Dim aRet As String
    If aStbSeqNo = Empty Then
        GetInitPlaceName = Empty
        Exit Function
    End If
    aQry = "SELECT Description FROM " & GetOwner & "CD056 " & _
            " WHERE CODENO = (SELECT InitPlaceNo FROM " & GetOwner & "SO004 " & _
            " WHERE SEQNO = '" & aStbSeqNo & "' AND CUSTID = " & aCustId & _
            " AND COMPCODE = " & gCompCode & " )"
    aRet = GetRsValue(aQry, gcnGi) & ""
    GetInitPlaceName = aRet
  Exit Function
ChkErr:
  Call ErrSub(frmName, "GetInitPlaceName")
End Function
Private Function IsRecordEmpty(ByRef aRs As ADODB.Recordset) As Boolean
  On Error GoTo ChkErr
    IsRecordEmpty = True
    If aRs Is Nothing Then
        Exit Function
    End If
    If aRs.State = adStateClosed Then
        Exit Function
    End If
    If aRs.EOF Or aRs.BOF Then
        Exit Function
    End If
    IsRecordEmpty = False
  Exit Function
ChkErr:
  Call ErrSub(frmName, "IsRecordEmpty")
End Function
'取出授權狀態
Public Function GetAuthStatus(ByRef aRs As ADODB.Recordset) As String
  On Error GoTo ChkErr
    Dim aRet As String
    Dim aSQL As String
'    If aRs Is Nothing Then
'      GetAuthStatus = Empty
'      Exit Function
'    End If
'    If aRs.EOF Or aRs.BOF Then
'        GetAuthStatus = Empty
'        Exit Function
'    End If
    If IsRecordEmpty(aRs) Then
        GetAuthStatus = Empty
    End If
    If Len(aRs("AuthStopTime") & "") = 0 Then
        GetAuthStatus = Empty
        Exit Function
    End If
    If Not IsDate(aRs("AuthStopTime") & "") Then
        GetAuthStatus = Empty
        Exit Function
    End If
    If CDate(aRs("AuthStopTime") & "") >= RightNow Then
        aRet = "收視中"
    Else
        aRet = "已過期"
    End If
    '#6115 增加 尚未授權 By Kin 2011/09/15
    aSQL = "SELECT COUNT(*) CNT FROM SO033PVOD2 " & _
            " WHERE SEQNO = " & aRs("SEQNO") & _
            " AND CMDSeqNo IS NULL "
    If Len(aRs("CMDSEQNO") & "") = 0 Then
        aRet = "尚未授權"
    End If
    GetAuthStatus = aRet
    Exit Function
ChkErr:
  Call ErrSub(frmName, "GetAuthStatus")
End Function
'取出授權完成時間
Public Function GetUpdAuthDate(ByVal aCMDSEQNO As String) As String
  On Error GoTo ChkErr
    Dim aRet As String
    Dim aSQL As String
    aRet = Empty
    If aCMDSEQNO = Empty Then
        GetUpdAuthDate = Empty
        Exit Function
    End If
    aSQL = "SELECT AuthTime FROM SO561A WHERE SEQNO = " & aCMDSEQNO
    aRet = GetRsValue(aSQL, cnCom) & ""
    If Len(aRet) = 0 Then
        aSQL = "SELECT AUTHTIME FROM SO560 WHERE SEQNO = " & aCMDSEQNO
        aRet = GetRsValue(aSQL, cnCom) & ""
    End If
    GetUpdAuthDate = aRet
    Exit Function
ChkErr:
  Call ErrSub(frmName, "GetUpdAuthDate")
End Function
Public Function GetAmount(ByVal aSEQNO As String) As String
  On Error GoTo ChkErr
  
  Exit Function
ChkErr:
  Call ErrSub(frmName, "GetAmount")
End Function
'取出繳費狀態
Public Function GetBillStatus(ByRef aRs As ADODB.Recordset) As String
  On Error GoTo ChkErr
    Dim aRet As String
    Dim aSQL As String
    Dim aBillNo As String
    Dim aBillItem  As String
    Dim aRsSO033 As New ADODB.Recordset
    If IsRecordEmpty(aRs) Then
        GetBillStatus = Empty
        Exit Function
    End If
    aBillNo = aRs("BillNo") & ""
    aBillItem = aRs("ITEM") & ""
    aRet = Empty
    '#6033 測試不OK,如果是作廢要顯示作廢 By Kin 2011/07/22
    '#6033 測試不OK,如果是包月制直接用STBSeqNo、CitemCode對回SO033取最大ShouldDate By Kin 2011/08/11
    Dim aCitemStr As String
    If (Val(aRs("ProductType") & "") = 0) And Len(aRs("CitemCode") & "") = 0 And Val(aRs("Amount")) = 0 Then
        aCitemStr = "-1"
    Else
        aCitemStr = aRs("CitemCode")
    End If
    
    
    If (Val(aRs("ProductType") & "") = 0) And (Val(aRs("OrderKind") & "") = 0) Then
        aSQL = "SELECT BillNo,Item FROM " & GetOwner & "SO033 " & _
                " Where SO033.CUSTID = " & aRs("CUSTID") & _
                " AND SO033.FaciSeqno = '" & aRs("STBSeqNo") & "' " & _
                " AND SO033.CITEMCODE = " & aCitemStr & _
                " ORDER BY SO033.SHOULDDATE DESC "
        aSQL = "SELECT * FROM ( " & aSQL & " ) WHERE ROWNUM = 1"
        If Not GetRS(aRsSO033, aSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
        If aRsSO033.RecordCount > 0 Then
            aBillNo = aRsSO033("BillNo") & ""
            aBillItem = aRsSO033("Item") & ""
        End If
    End If
    If Len(aBillNo) = 0 Then
        GetBillStatus = "無單據"
        Exit Function
    End If
    aSQL = "SELECT COUNT(*) CNT FROM " & GetOwner & "SO033 " & _
        " WHERE SO033.CANCELFLAG = 1 " & _
        " AND SO033.CUSTID = " & aRs("CUSTID") & _
        " AND SO033.BILLNO = '" & aBillNo & "' "
    
    If Len(aBillItem) <> 0 Then
        aSQL = aSQL & " AND SO033.ITEM = " & aBillItem
    End If
    
    If Val(GetRsValue(aSQL, gcnGi) & "") = 1 Then
        aRet = "作廢"
    Else
        aSQL = "SELECT COUNT(*) CNT FROM " & GetOwner & "SO033," & GetOwner & "CD013 " & _
                " WHERE SO033.UCCODE IS NOT NULL AND CUSTID = " & aRs("CUSTID") & _
                " AND SO033.UCCODE = CD013.CODENO AND CD013.REFNO NOT IN (3,8) " & _
                " AND NVL(CD013.PAYOK,0) = 0 AND SO033.BILLNO = '" & aBillNo & "' "
        If Len(aRs("ITEM") & "") <> 0 Then
            aSQL = aSQL & " AND SO033.ITEM = " & aBillItem
        End If
        If Val(GetRsValue(aSQL, gcnGi) & "") = 0 Then
            aRet = "已繳費"
        Else
            aRet = "未繳費"
        End If
    End If
    On Error Resume Next
    Call CloseRecordset(aRsSO033)
    GetBillStatus = aRet
  Exit Function
ChkErr:
  Call ErrSub(frmName, "GetBillStatus")
End Function
'取出訂購狀態
Public Function GetOrderStatus(ByRef aRs As ADODB.Recordset) As String
  On Error GoTo ChkErr
    Dim aRet As String
'    If aRs Is Nothing Then
'      GetOrderStatus = Empty
'      Exit Function
'    End If
'    If aRs.EOF Or aRs.BOF Then
'        GetOrderStatus = Empty
'        Exit Function
'    End If
    If IsRecordEmpty(aRs) Then
        GetOrderStatus = Empty
    End If
    '#6115 改成判斷CancelFlag By Kin 2011/09/15
    If Val(aRs("CancelFlag") & "") > 0 Then
        aRet = "退租"
    Else
        aRet = "有效"
    End If
    GetOrderStatus = aRet
  Exit Function
ChkErr:
  Call ErrSub(frmName, "GetOrderStatus")
End Function
Public Sub ShowDataView()
  On Error GoTo ChkErr
    If rsSO033Pvod Is Nothing Then
         MsgBox "無任何訂購資料明細", vbInformation, "訊息"
         Exit Sub
    End If
    If rsSO033Pvod.State = adStateClosed Then
        MsgBox "無任何訂購資料明細", vbInformation, "訊息"
         Exit Sub
    End If
    
    If rsSO033Pvod.RecordCount <= 0 Then
        MsgBox "無任何訂購資料明細", vbInformation, "訊息"
        Exit Sub
    End If
    If rsSO004 Is Nothing Then
        MsgBox "無任何設備資料", vbInformation, "訊息"
        Exit Sub
    End If
    If rsSO004.State = adStateClosed Then
        MsgBox "無任何設備資料", vbInformation, "訊息"
        Exit Sub
    End If
    If rsSO004.RecordCount <= 0 Then
        MsgBox "無任何設備資料", vbInformation, "訊息"
        Exit Sub
    End If
    With frmBrower
        .uCanSelect = fCanSelect
        Set frmBrower.uSourceRs = rsSO033Pvod
        Set frmBrower.uSO004Rs = rsSO004
        Set frmBrower.uSO563Rs = rsSO563
        .Show vbModal
    End With
  Exit Sub
ChkErr:
  Call ErrSub(frmName, "ShowDataView")
End Sub
Public Function GetQryData() As ADODB.Recordset
  On Error GoTo ChkErr
    Dim aSQL As String
    Dim aSQL2 As String
    Dim aSQL3 As String
    Dim aSO563SQL As String
    Dim aSO004SQL As String
    Dim aStbSeqNo As String
    Dim rsSTBSEQNO As New ADODB.Recordset
    Dim rsField As New ADODB.Recordset
    Dim i As Long
    
    fSO033PvodField = Empty
    If Not GetRS(rsField, " SELECT * FROM SO033PVODOLD WHERE 1 = 0 ", cnCom, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
    For i = 0 To rsField.Fields.Count - 1
        If fSO033PvodField = Empty Then
            fSO033PvodField = "SO033PVOD." & rsField.Fields(i).Name
        Else
            fSO033PvodField = fSO033PvodField & ",SO033PVOD." & rsField.Fields(i).Name
        End If
    Next i
    
    
    
    aSQL = GetSQL & " UNION ALL " & GetSQL2       '取出資料的語法
    aSQL = " SELECT DISTINCT * FROM ( " & aSQL & " ) ORDER BY ORDFLAG,AuthStartTime "
    Call ClearRs
    If Not GetRS(rsSO033Pvod, aSQL, cnCom, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
    If Not rsSO033Pvod Is Nothing Then
        If Not rsSO033Pvod.EOF Then
            If rsSO033Pvod.RecordCount > 0 Then
                aSQL2 = "SELECT DISTINCT STBSEQNO FROM ( " & aSQL & " )"
                aSQL3 = "SELECT DISTINCT SEQNO FROM ( " & aSQL & " )"
                If Not GetRS(rsSTBSEQNO, aSQL, cnCom, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
                If Not rsSTBSEQNO Is Nothing Then
                    If Not rsSTBSEQNO.EOF Then
                        rsSTBSEQNO.MoveFirst
                        Do While Not rsSTBSEQNO.EOF
                            If Len(rsSTBSEQNO("STBSEQNO") & "") > 0 Then
                                If aStbSeqNo = Empty Then
                                    aStbSeqNo = "'" & rsSTBSEQNO("STBSEQNO") & "'"
                                Else
                                    aStbSeqNo = aStbSeqNo & "," & "'" & rsSTBSEQNO("STBSEQNO") & "'"
                                End If
                            End If
                            rsSTBSEQNO.MoveNext
                        Loop
                    End If
                End If
'                rsSO033Pvod.MoveFirst
'                Do While Not rsSO033Pvod.EOF
'                    If Len(rsSO033Pvod("STBSEQNO") & "") > 0 Then
'                        If aStbSeqNo = Empty Then
'                            aStbSeqNo = "'" & rsSO033Pvod("STBSEQNO") & "'"
'                        Else
'                            aStbSeqNo = aStbSeqNo & "," & "'" & rsSO033Pvod("STBSEQNO") & "'"
'                        End If
'                    End If
'                    rsSO033Pvod.MoveNext
'                Loop
            End If
        End If
    End If
    If aStbSeqNo <> Empty Then
        aSO004SQL = "SELECT RowId,A.* FROM " & GetOwner & "SO004 A " & _
                    " WHERE SEQNO IN ( " & aStbSeqNo & ") " & _
                    " AND CUSTID = " & gCustId & _
                    " AND COMPCODE = " & gCompCode
                    
        If Not GetRS(rsSO004, aSO004SQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
        
    End If
    
    If aSQL3 <> Empty Then
    
'        aSO563SQL = "SELECT SO563.*,SO033PVOD2.AuthStartTime,SO033PVOD2.AuthStopTime, " & _
'                    "SO033PVOD.SEQNO " & _
'                    " FROM SO563,SO562A,SO562,SO033PVOD2,SO033PVOD " & _
'                    " WHERE SO562.ID = SO562A.PRODUCTID " & _
'                    " AND SO562A.VODITEMID = SO563.ID " & _
'                    " AND SO563.TYPE = 5 " & _
'                    " AND SO033PVOD.SEQNO IN ( " & aSQL3 & " )" & _
'                    " AND SO033PVOD2.SEQNO = SO033PVOD.SEQNO " & _
'                    " AND SO033PVOD2.CASID = SO562.CASID "
'
        aSO563SQL = "SELECT SO563.*,SO033PVOD2.AuthStartTime,SO033PVOD2.AuthStopTime, " & _
                    "SO033PVOD.SEQNO " & _
                    " FROM SO563,SO562A,SO562,SO033PVOD2,SO033PVOD,( " & aSQL & " ) A " & _
                    " WHERE SO562.ID = SO562A.PRODUCTID " & _
                    " AND SO562A.VODITEMID = SO563.ID " & _
                    " AND SO563.TYPE = 5 " & _
                    " AND NVL(SO563.LOCALLANGUAGE,0) = 0 " & _
                    " AND SO033PVOD.SEQNO = A.SEQNO " & _
                    " AND SO033PVOD2.SEQNO = SO033PVOD.SEQNO " & _
                    " AND SO033PVOD2.CASID = SO562.CASID " & _
                    " AND TO_CHAR(SO033PVOD2.AUTHSTOPTIME,'YYYYMM')= TO_CHAR(A.AUTHSTOPTIME,'YYYYMM')"
                    
        aSO563SQL = "SELECT DISTINCT * FROM ( " & aSO563SQL & " )"
        If Not GetRS(rsSO563, aSO563SQL, cnCom, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
    End If
    'select C.* from SO562 A,SO562A B,SO563 C " & _ "Where A.ID = B.ProductId And B.VoditemID = C.ID And C.Type = 5 And A.CASID = " & strCASID
    
    Set GetQryData = rsSO033Pvod.Clone
    On Error Resume Next
    CloseRecordset rsSTBSEQNO
    CloseRecordset rsField
    Exit Function
ChkErr:
  Call ErrSub(frmName, "GetQryData")
End Function
Public Sub ClearRs()
  On Error Resume Next
    CloseRecordset rsSO033Pvod
    CloseRecordset rsSO004
    Set rsSO033Pvod = Nothing
    Set rsSO004 = Nothing
End Sub


'包月制的語法
Public Function GetSQL2() As String
  On Error GoTo ChkErr
    Dim aFields As String
    Dim aFields2 As String
    Dim aSQL As String
    Dim aHavingSQL As String
    
    Dim aTable As String
    Dim i As Integer
    Dim intCnt As Integer
    If Not fShowHistory Then
        aHavingSQL = " HAVING MAX(SO033PVOD2.AUTHSTOPTIME) >= TO_DATE('" & Format(RightNow, "YYYYMMDD") & "','YYYYMMDD') "
        intCnt = 1
    Else
        aHavingSQL = Empty
        intCnt = 2
    End If
'    aFields = " SO563.TITLE,SO562.RENTALDURATION, " & _
'                 fSO033PvodField & " ,SO033PVOD2.CMDSEQNO "
    aFields = " SO563.TITLE,SO562.RENTALDURATION, " & _
                 fSO033PvodField
                
    aFields2 = ", MIN(SO033PVOD2.AuthStartTime) AuthStartTime, " & _
                " MAX(SO033PVOD2.AuthStopTime) AuthStopTime, " & _
                "DECODE(NVL(SO033PVOD.ReturnCode,0),0,0,1) ORDFLAG "
    For i = 1 To intCnt
        Select Case i
            Case 1
                aTable = " SO033PVOD SO033PVOD "
            Case 2
                aTable = " SO033PVODOLD SO033PVOD "
                aSQL = aSQL & " UNION ALL "
        End Select
        aSQL = aSQL & "SELECT " & aFields & aFields2 & ",MAX(SO033PVOD2.CMDSEQNO) CMDSEQNO " & _
                    " From " & aTable & " , " & _
                    " SO562,SO563,SO033PVOD2  " & GetSQLWhere & _
                    " AND SO033PVOD.PRODUCTTYPE = 0 " & _
                    " AND SO033PVOD.SEQNO  = SO033PVOD2.SEQNO " & _
                    aHavingSQL & _
                    " GROUP BY " & aFields
                 
       
    Next i
    aSQL = "SELECT DISTINCT * FROM (" & aSQL & ") "
    GetSQL2 = aSQL
  Exit Function
ChkErr:
  Call ErrSub(frmName, "GetSQL2")
End Function

'論片制的語法
Public Function GetSQL() As String
  On Error GoTo ChkErr
    Dim aFields As String
    Dim aSQL As String
    Dim aFields2 As String
    Dim aHavingSQL As String
    Dim aTable As String
    Dim i As Integer
    Dim intCnt As Integer
    '如果沒有顯示歷史資料要過濾授權終止日期>現在時間
    If Not fShowHistory Then
        aHavingSQL = " HAVING MAX(SO033PVOD2.AUTHSTOPTIME) >= TO_DATE('" & Format(RightNow, "YYYYMMDD") & "','YYYYMMDD') "
        intCnt = 1
    Else
        aHavingSQL = Empty
        intCnt = 2
    End If
'    aFields = " SO563.TITLE,SO562.RENTALDURATION, " & _
'                fSO033PvodField & " ,SO033PVOD2.CMDSEQNO "
     aFields = " SO563.TITLE,SO562.RENTALDURATION, " & _
                fSO033PvodField
               
    aFields2 = ", MIN(SO033PVOD2.AuthStartTime) AuthStartTime, " & _
                " MAX(SO033PVOD2.AuthStopTime) AuthStopTime, " & _
                "DECODE(NVL(SO033PVOD.ReturnCode,0),0,0,1) ORDFLAG"
                
    For i = 1 To intCnt
        Select Case i
            Case 1
                aTable = " SO033PVOD SO033PVOD "
            Case 2
                aTable = " SO033PVODOLD SO033PVOD "
                aSQL = aSQL & " UNION ALL "
        End Select
        
        aSQL = aSQL & "SELECT " & aFields & aFields2 & ",MAX(SO033PVOD2.CMDSEQNO) CMDSEQNO " & _
                    " From " & aTable & " , " & _
                    " SO562,SO563,SO033PVOD2  " & GetSQLWhere & getOrdDateWhere & _
                    " AND SO033PVOD.PRODUCTTYPE IN (1,2) " & _
                    " AND SO033PVOD.SEQNO  = SO033PVOD2.SEQNO " & _
                    aHavingSQL & _
                    " GROUP BY " & aFields
    Next
    GetSQL = aSQL
    
  Exit Function
ChkErr:
  Call ErrSub(frmName, "GetSQL")
End Function
Public Function getOrdDateWhere() As String
  On Error GoTo ChkErr
    Dim aRetString As String
    If fOrderDate1 <> Empty Then
        aRetString = aRetString & " AND SO033PVOD.ORDERDATE >= TO_DATE('" & _
                    fOrderDate1 & "','YYYYMMDD')"
    End If
    If fOrderDate1 <> Empty Then
        aRetString = aRetString & " AND SO033PVOD.ORDERDATE <= TO_DATE('" & _
                    fOrderDate2 & "','YYYYMMDD')"
    End If
    getOrdDateWhere = aRetString
    Exit Function
ChkErr:
  Call ErrSub(frmName, "getOrdDateWhere")
End Function

Public Function GetSQLWhere() As String
  On Error GoTo ChkErr
    Dim aRetString As String
    aRetString = " WHERE 1 = 1 "
    aRetString = aRetString & " AND SO033PVOD2.CASID = SO562.CASID "
    aRetString = aRetString & " AND SO033PVOD.COMPCODE = " & gCompCode
    aRetString = aRetString & " AND SO562.ID = SO563.ID AND SO563.TYPE = 1 AND NVL(SO563.LOCALLANGUAGE,0) = 0 "
    
    '沒有顯示歷史資料ReturnCode一定要為Null By Kin 2011/06/08
    If Not fShowHistory Then
        aRetString = aRetString & " AND SO033PVOD.ReturnCode IS NULL "
    End If
    
    'aRetString = aRetString & " AND NVL(SO033PVOD.ORDERWAYCODE,0) = CD114.CODENO "
    'aRetString = aRetString & " AND CD114.STOPFLAG <> 1 "
    
    If Val(gCustId & "") > 0 Then
        aRetString = aRetString & " AND SO033PVOD.CUSTID = " & gCustId
    End If
    
    
    If fOrderWayCode <> Empty Then
        aRetString = aRetString & " AND SO033PVOD.ORDERWAYCODE In(" & fOrderWayCode & ")"
    End If
    
    If fOrderEn <> Empty Then
        aRetString = aRetString & " AND SO033PVOD.ORDEREN In (" & fOrderEn & ")"
    End If
    
    If (fBillNo <> Empty) And (fProductType = 1 Or fProductType = 99) Then
        aRetString = aRetString & " AND SO033PVOD.BILLNO = '" & fBillNo & "'"
    End If
    If (fBillNoItem <> Empty) And (fProductType = 1 Or fProductType = 99) Then
        aRetString = aRetString & " AND SO033PVOD.ITEM = " & fBillNoItem
    End If
    If (fBillNoAndItem <> Empty) And (fProductType = 1 Or fProductType = 99) Then
        aRetString = aRetString & " And SO033PVOD.BillNo || SO033PVOD.Item In ('" & fBillNoAndItem & "')"
    End If
    If (fProductType = 99 Or fProductType = 0) Then
        If fSTBSeqNos <> Empty Then
            aRetString = aRetString & " AND SO033PVOD.STBSEQNO IN (" & fSTBSeqNos & ") "
        End If
        If fSTBSeqNo <> Empty Then
            aRetString = aRetString & " AND SO033PVOD.STBSEQNO = '" & fSTBSeqNo & "' "
        End If
        If fCitemCode <> Empty Then
            aRetString = aRetString & " AND SO033PVOD.CITEMCODE = " & fCitemCode
        End If
    End If
    
    GetSQLWhere = aRetString
    Exit Function
ChkErr:
  Call ErrSub(frmName, "GetSQLWhere")
End Function
Public Sub SetGimQry(objMs As Object, qryStr As String, _
                    Optional FldCaption1 As String = "", _
                    Optional FldCaption2 As String = "", _
                    Optional ColumnCaption As String = "", _
                    Optional ColumnWidth As String = "", _
                    Optional WideScreen As Boolean = False, _
                    Optional FldCaption3 As String = "", Optional FldCaption4 As String = "", Optional FldCaption5 As String = "", Optional FldCaption6 As String = "", Optional FldCaption7 As String = "", _
                    Optional FldCaption8 As String = "", Optional FldCaption9 As String = "", Optional FldCaption10 As String = "", Optional FldCaption11 As String = "", Optional FldCaption12 As String = "", _
                    Optional Fld1Width As String = "", Optional Fld2Width As String = "", Optional Fld3Width As String = "", Optional Fld4Width As String = "", Optional Fld5Width As String = "", Optional Fld6Width As String = "", _
                    Optional Fld7Width As String = "", Optional Fld8Width As String = "", Optional Fld9Width As String = "", Optional Fld10Width As String = "", Optional Fld11Width As String = "", Optional Fld12Width As String = "", _
                    Optional Fld1Fmt As String = "", Optional Fld2Fmt As String = "", Optional Fld3Fmt As String = "", Optional Fld4Fmt As String = "", Optional Fld5Fmt As String = "", Optional Fld6Fmt As String = "", _
                    Optional Fld7Fmt As String = "", Optional Fld8Fmt As String = "", Optional Fld9Fmt As String = "", Optional Fld10Fmt As String = "", Optional Fld11Fmt As String = "", Optional Fld12Fmt As String = "", _
                    Optional FldIdx As Integer = 1)
  On Error GoTo ChkErr
    With objMs
        .SendConn gcnGi
        .QueryString = qryStr
         If TypeOf objMs Is CSmulti Then
            .SetFldIdx = FldIdx
            .WideScreen = WideScreen
         End If
         If Len(FldCaption1) > 0 Then .FldCaption1 = FldCaption1
         If Len(FldCaption2) > 0 Then .FldCaption2 = FldCaption2
         If Len(FldCaption3) > 0 Then .FldCaption3 = FldCaption3
         If Len(FldCaption4) > 0 Then .FldCaption4 = FldCaption4
         If Len(FldCaption5) > 0 Then .FldCaption5 = FldCaption5
         If Len(FldCaption6) > 0 Then .FldCaption6 = FldCaption6
         If Len(FldCaption7) > 0 Then .FldCaption7 = FldCaption7
         If Len(FldCaption8) > 0 Then .FldCaption8 = FldCaption8
         If Len(FldCaption9) > 0 Then .FldCaption9 = FldCaption9
         If Len(FldCaption10) > 0 Then .FldCaption10 = FldCaption10
         If Len(FldCaption11) > 0 Then .FldCaption11 = FldCaption11
         If Len(FldCaption12) > 0 Then .FldCaption12 = FldCaption12
         If Len(ColumnCaption) > 0 Then .ColumnCaption = ColumnCaption
         If Len(ColumnWidth) > 0 Then .ColumnWidth = ColumnWidth
        
         If Len(Fld1Width) > 0 Then .Fld1Width = Fld1Width
         If Len(Fld2Width) > 0 Then .Fld2Width = Fld2Width
         If Len(Fld3Width) > 0 Then .Fld3Width = Fld3Width
         If Len(Fld4Width) > 0 Then .Fld4Width = Fld4Width
         If Len(Fld5Width) > 0 Then .Fld5Width = Fld5Width
         If Len(Fld6Width) > 0 Then .Fld6Width = Fld6Width
         If Len(Fld7Width) > 0 Then .Fld7Width = Fld7Width
         If Len(Fld8Width) > 0 Then .Fld8Width = Fld8Width
         If Len(Fld9Width) > 0 Then .Fld9Width = Fld9Width
         If Len(Fld10Width) > 0 Then .Fld10Width = Fld10Width
         If Len(Fld11Width) > 0 Then .Fld11Width = Fld11Width
         If Len(Fld12Width) > 0 Then .Fld12Width = Fld12Width
    
         If Len(Fld1Fmt) > 0 Then .FldFormat1 = Fld1Fmt
         If Len(Fld2Fmt) > 0 Then .FldFormat2 = Fld2Fmt
         If Len(Fld3Fmt) > 0 Then .FldFormat3 = Fld3Fmt
         If Len(Fld4Fmt) > 0 Then .FldFormat4 = Fld4Fmt
         If Len(Fld5Fmt) > 0 Then .FldFormat5 = Fld5Fmt
         If Len(Fld6Fmt) > 0 Then .FldFormat6 = Fld6Fmt
         If Len(Fld7Fmt) > 0 Then .FldFormat7 = Fld7Fmt
         If Len(Fld8Fmt) > 0 Then .FldFormat8 = Fld8Fmt
         If Len(Fld9Fmt) > 0 Then .FldFormat9 = Fld9Fmt
         If Len(Fld10Fmt) > 0 Then .FldFormat10 = Fld10Fmt
         If Len(Fld11Fmt) > 0 Then .FldFormat11 = Fld11Fmt
         If Len(Fld12Fmt) > 0 Then .FldFormat12 = Fld12Fmt
    End With
  Exit Sub
ChkErr:
    ErrSub frmName, "SetGimQry"
End Sub
Public Function GetDVRSno(ByVal lngDVRCustID As Long, ByVal strSTBSNo As String, _
                          Optional ByRef strDVRFaciCode As String = "", Optional ByRef strDVRHDDCode As String = "", _
                          Optional ByRef strSTBSeqNo As String = "", Optional ByRef blnClearPrSNO As Boolean = False, _
                          Optional ByRef strDVRHDDSizeCode As String = "", Optional ByVal strPRSNO As String = "") As String
On Error GoTo ChkErr
    Dim rsDVR As New ADODB.Recordset
    Dim strDVRSNOfnG As String
    Dim intLoop As Integer
    strDVRFaciCode = ""
    strDVRHDDCode = ""
    strDVRHDDSizeCode = ""
                
    If strSTBSeqNo = "" Then
        strSTBSeqNo = GetRsValue("Select SeqNo From " & GetOwner & "SO004 Where FaciSno='" & strSTBSNo & "' and Custid=" & lngDVRCustID & IIf(blnClearPrSNO, "", " and PrSNO Is Null")) & ""
    End If
    
    If Not GetRS(rsDVR, "Select FaciSNO,FaciCode,DVRAuthSizeCode,DVRSizeCode From " & GetOwner & "SO004 Where STBSNO='" & strSTBSeqNo & "' and Custid=" & lngDVRCustID & IIf(Len(strPRSNO) > 0, " and PRSNo = '" & strPRSNO & "'", IIf(blnClearPrSNO, "", " and PrSNO Is Null"))) Then Exit Function
    If rsDVR.RecordCount > 0 Then rsDVR.MoveFirst
    Do While Not rsDVR.EOF
        strDVRSNOfnG = strDVRSNOfnG & "/" & rsDVR("FaciSNO") & ""
        strDVRFaciCode = strDVRFaciCode & "/" & rsDVR("FaciCode") & ""
        strDVRHDDCode = strDVRHDDCode & "/" & rsDVR("DVRAuthSizeCode") & ""
        strDVRHDDSizeCode = strDVRHDDSizeCode & "/" & rsDVR("DVRSizeCode") & ""
        rsDVR.MoveNext
    Loop
    If Len(strDVRSNOfnG) > 0 Then GetDVRSno = Mid(strDVRSNOfnG, 2)
    If Len(strDVRFaciCode) > 0 Then strDVRFaciCode = Mid(strDVRFaciCode, 2)
    If Len(strDVRHDDCode) > 0 Then strDVRHDDCode = Mid(strDVRHDDCode, 2)
    If Len(strDVRHDDSizeCode) > 0 Then strDVRHDDSizeCode = Mid(strDVRHDDSizeCode, 2)
    
    CloseRecordset rsDVR
    Exit Function
ChkErr:
    ErrSub frmName, "GetDVRSno"
End Function


Public Function GetDVRHDD(ByVal lngDVRCustID As Long, ByVal strSTBSNo As String, _
                          Optional ByRef strDvrHDDData As String, Optional ByVal strSTBSeqNo As String, _
                          Optional ByVal blnAuth As Boolean = False, Optional ByVal blnCancel As Boolean) As String
On Error GoTo ChkErr
    Dim strDVRSNOfnHD As String, intDVRCount As Integer
    Dim strDSNO() As String, strDvrHDDSno As String
    Dim rsDVR As New ADODB.Recordset
    Dim strDvrSizeName As String
    
    If blnAuth Then
        strDvrSizeName = "DVRAuthSizeCode"
    Else
        strDvrSizeName = "DVRSizeCode"
    End If
    
    strDVRSNOfnHD = GetDVRSno(lngDVRCustID, strSTBSNo, , , strSTBSeqNo, blnCancel)
    If Len(strDVRSNOfnHD) > 0 Then
        strDSNO = Split(strDVRSNOfnHD, "/")
        For intDVRCount = 0 To UBound(strDSNO)
            '#5265 2009.08.26 by Corey 程式增加PRDATE　IS　NULL的條件，並且如果有兩台相同FACISNO則取SEQNO最大值來取得DVR授權容量。
            If Not GetRS(rsDVR, "Select * From " & GetOwner & "CD102 Where CodeNo in " & _
                        "(Select " & strDvrSizeName & " FROM (Select " & strDvrSizeName & ",(RANK() OVER ( ORDER BY SeqNo Desc)) RankX From " & _
                                                      GetOwner & "SO004 Where FaciSno='" & strDSNO(intDVRCount) & "'" & _
                                                      " and Custid=" & lngDVRCustID & " and STBSNO='" & strSTBSeqNo & "' and PRDATE IS NULL)" & _
                        " WHERE RankX = 1) ") Then Exit Function
            If rsDVR.RecordCount > 0 Then
                strDvrHDDSno = strDvrHDDSno & "/" & Trim(str(Val(rsDVR("DVRSize") & "")))
                strDvrHDDData = strDvrHDDData & "/" & Trim(str(Val(rsDVR("DVRSize") & "")))
            End If
        Next
        If Len(strDvrHDDSno) > 0 Then GetDVRHDD = Mid(strDvrHDDSno, 2)
        If Len(strDvrHDDData) > 0 Then strDvrHDDData = Mid(strDvrHDDData, 2)
    End If
    CloseRecordset rsDVR
    Exit Function
ChkErr:
    GetDVRHDD = "0"
    ErrSub frmName, "GetDVRSno"
End Function


