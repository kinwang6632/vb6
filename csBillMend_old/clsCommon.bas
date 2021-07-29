Attribute VB_Name = "clsComm"
Option Explicit
Private Const FormName = "clsModule"
Public blnNoShowMsg As Boolean
Public rsSO0304 As New ADODB.Recordset
Public blnNoDefine As Boolean       '外部自己傳來的,就不需自己找資料,直接以傳來的為主
Public rsTmp As ADODB.Recordset
Public rsCopyClone As Recordset '外部自己傳來的RecordSet,rsSO0304,要以這個為主
Public blnTrans  As Boolean     '是否做Trans False(代表內部程式自己做) True(代表外部程式自己做)
Public blnSelectAll As Boolean '是否全選資料
Public strCloseDate As String  '結清日期
Public strRetBillNo As String
Public rsSO003 As New ADODB.Recordset
Public strNewPeriod As String
Public strNewClctStopDate As String
Public strAryName As String
Public strAryName2 As String
Public strTmpSO032 As String
Public strRowIds As String
Public Function DefineRs(Optional ByVal aCitemCodes As String = Empty) As ADODB.Recordset
  On Error GoTo ChkErr
    Dim strQry As String
    Dim i As Long
    Dim strWhere As String
    
    Set rsTmp = New ADODB.Recordset
    strWhere = " SO003.CUSTID=" & gCustId & _
             " And SO003.COMPCODE=" & gCompCode & _
             " And SO003.CUSTID=SO004.CUSTID(+) And SO003.COMPCODE=SO004.COMPCODE(+)" & _
             " And SO003.FaciSeqNo=SO004.SeqNo(+)" & _
             " And SO003.StopFlag<>1  And SO003.CitemCode=CD019.CodeNo And CD019.NoShowCitem=0"
    If aCitemCodes <> Empty Then
        strWhere = strWhere & " AND SO003.CITEMCODE IN (" & aCitemCodes & ") "
    End If
    strQry = "SELECT SO003.OrderNo,SO004.DeclarantName,SO003.CITEMCODE,SO003.CITEMNAME,SO003.PERIOD,SO003.AMOUNT," & _
             "SO003.ACCOUNTNO,SO003.STARTDATE,SO003.StopDate,SO003.CLCTSTOPDATE," & _
             "SO003.SEQNO,SO003.FaciSNo,SO004.FaciName,SO003.CLCTDATE,FaciSeqNo,SO003.RowId,SO003.CustId FROM " & _
             GetOwner & "SO003," & GetOwner & "SO004, " & GetOwner & "CD019"
    '如果對方有傳RecordSet直接使用傳進來的RecordSet
    If Not blnNoDefine Then
        If Not GetRS(rsSO0304, strQry & " Where 1=0", gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
    Else
        Set rsSO0304 = rsCopyClone
    End If
    
    With rsTmp
        If .State = adStateOpen Then .Close
        .CursorLocation = adUseClient
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Fields.Append ("Choice"), adBSTR, 1, adFldIsNullable
        For i = 0 To rsSO0304.Fields.Count - 1
            .Fields.Append rsSO0304.Fields(i).Name, adBSTR, rsSO0304.Fields(i).DefinedSize, adFldIsNullable
        Next i
    End With
    
    rsTmp.Open
    If Not blnNoDefine Then
        If Not GetRS(rsSO0304, strQry & " Where " & strWhere, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
    End If
    If rsSO0304.RecordCount > 0 Then
        rsSO0304.MoveFirst
    End If
    Do While Not rsSO0304.EOF
        rsTmp.AddNew
        For i = 0 To rsSO0304.Fields.Count - 1
            If Not IsNull(rsSO0304.Fields(i).Value) Then
                rsTmp.Fields(rsSO0304.Fields(i).Name).Value = rsSO0304.Fields(i).Value & ""
            End If
        Next i
        If blnNoDefine Or blnSelectAll Then rsTmp("Choice").Value = "1"
        rsTmp.Update
        rsSO0304.MoveNext
    Loop
    Set DefineRs = rsTmp
    On Error Resume Next
    
    Exit Function
ChkErr:
  ErrSub FormName, "OpenData"
End Function
'取得挑選的RowId與設備流水號
Public Sub GetRowIdAndFaci(ByRef aRs As ADODB.Recordset, _
        ByRef aRowid As String, ByRef aFaciSeqNo As String)
  On Error GoTo ChkErr
    aRowid = Empty
    aFaciSeqNo = Empty
    If aRs.RecordCount > 0 Then aRs.MoveFirst
    Do While Not aRs.EOF
        If aRs("Choice") = "1" Then
            aRowid = IIf(aRowid <> "", aRowid & ",'" & aRs("RowId") & "'", "'" & aRs("RowId") & "'")
            '#5569 要把設備挑出來傳給後端 By Kin 2010/03/03
            If aRs("FaciSeqNo") & "" <> "" Then
                If aFaciSeqNo = Empty Then
                    aFaciSeqNo = "'" & aRs("FaciSeqNo") & "'"
                Else
                    If InStr(1, aFaciSeqNo, aRs("FaciSeqNo") & "") <= 0 Then
                        aFaciSeqNo = aFaciSeqNo & ",'" & aRs("FaciSeqNo") & "'"
                    End If
                End If
            End If
        End If
        aRs.MoveNext
    Loop
    If aFaciSeqNo <> Empty Then
        If InStr(1, aFaciSeqNo, ",") > 0 Then
            aFaciSeqNo = "In(" & aFaciSeqNo & ")"
        Else
            aFaciSeqNo = "=" & aFaciSeqNo
        End If
    End If
    Exit Sub
ChkErr:
 ErrSub FormName, "GetRowIdAndFaci"
End Sub
'取得挑選的SO003
Public Function GetRs003(ByVal aRowIds As String) As ADODB.Recordset
  On Error GoTo ChkErr
    Dim strQry As String
    Dim rs03 As New ADODB.Recordset
    strQry = "Select Period,Amount,CustAllot,CLCTSTOPDATE,CLCTDATE,RowId," & _
            "StopDate,CitemCode,FaciSeqNo,Nvl(PayKind,0) PayKind " & _
            " From " & GetOwner & "SO003" & _
            " Where RowId In (" & aRowIds & ")"
    If Not GetRS(rs03, strQry, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then
        Exit Function
    End If
    Set GetRs003 = rs03
    Exit Function
ChkErr:
  Call ErrSub(FormName, "GetRs003")
End Function
'取得應收日期的起迄,要傳入給後端的參數
Public Sub GetShouldDate(ByVal aRowIds As String, ByRef aShouldDate1 As String, ByRef aShouldDate2 As String)
  On Error GoTo ChkErr
    Dim rsDate As New ADODB.Recordset
    Dim strQry As String
    strQry = "Select Min(ClctDate) ShouldDate1,Max(ClctDate) ShouldDate2" & _
            " From " & GetOwner & "SO003" & _
            " Where RowId In (" & aRowIds & ")"
    If Not GetRS(rsDate, strQry, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
    aShouldDate1 = Format(rsDate("ShouldDate1"), "yyyymmdd")
    aShouldDate2 = Format(rsDate("ShouldDate2"), "yyyymmdd")
    
    On Error Resume Next
    Call CloseRecordset(rsDate)
    Exit Sub
ChkErr:
    ErrSub FormName, "GetShouldDate"
End Sub
Public Function GetServiceTypes(ByVal aRowIds As String) As String
   On Error GoTo ChkErr
    Dim strQry As String
    Dim strServiceTypes As String
    strQry = "Select Min(ServiceType) ServiceType From " & GetOwner & "SO003" & _
            " Where RowId In (" & aRowIds & ")" & _
            " Group By ServiceType"
    strServiceTypes = gcnGi.Execute(strQry).GetString(, , , ",")
    strServiceTypes = Mid(strServiceTypes, 1, Len(strServiceTypes) - 1)
    GetServiceTypes = strServiceTypes
    Exit Function
ChkErr:
    ErrSub FormName, "GetServiceType"
End Function
Public Function GetCitemCodes(ByVal aRowIds As String) As String
  On Error GoTo ChkErr
    Dim strQry As String
    Dim strCitemCodes As String
    strQry = "Select CitemCode From " & GetOwner & "SO003" & _
            " Where RowId In (" & aRowIds & ")"
    strCitemCodes = gcnGi.Execute(strQry).GetString(, , , ",")
    strCitemCodes = Mid(strCitemCodes, 1, Len(strCitemCodes) - 1)
    GetCitemCodes = strCitemCodes
    Exit Function
ChkErr:
    ErrSub FormName, "GetCitemCode"
End Function

Public Sub GetUseDate(ByVal aCloseDate As String, ByVal aCitemCodes As String, _
    ByRef aShouldDate2 As String, ByRef rs03 As ADODB.Recordset)
  On Error GoTo ChkErr
    Dim rsDate As New ADODB.Recordset
    Dim aMaxShouldDate2 As String
    Dim strSO033SQL As String
    If rs03.RecordCount > 0 Then rs03.MoveFirst
    If aCloseDate <> Empty Or 1 = 1 Then
         Do While Not rs03.EOF
            '#5815 只有現付制才需要找出最大值 By Kin 2010/10/21
            If Val(rs03("PayKind") & "") = 1 Then
                strSO033SQL = "Select Max(RealStopDate) MaxDate,Count(*) Cnt From " & GetOwner & "SO033" & _
                        " Where CustId=" & gCustId & _
                        " And UCcode is Not Null " & _
                        " And CitemCode =" & rs03("CitemCode") & "" & _
                        " And FaciSeqNo='" & rs03("FaciSeqNo") & "'"
                If Not GetRS(rsDate, strSO033SQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
                '找出每筆的RealStopDate,並找出其中一筆最大值 By Kin 2010/09/16
                If Not rsDate.EOF Then
                    If Val(rsDate("Cnt") & "") > 0 Then
                        If aMaxShouldDate2 <> "" Then
                            If CDate(aMaxShouldDate2) < rsDate("MaxDate") Then aMaxShouldDate2 = rsDate("MaxDate")
                        Else
                            aMaxShouldDate2 = rsDate("MaxDate")
                        End If
                        '將so003的ClctDate與StopDate UPDATE成SO033的日期 By Kin 2010/09/16
                        rs03("ClctDate") = CDate(rsDate("MaxDate")) + 1
                        rs03("StopDate") = rsDate("MaxDate") & ""
                        rs03.Update
                    End If
                End If
                
            End If
            rs03.MoveNext
        Loop
        If aMaxShouldDate2 <> Empty Then
            If aCloseDate <> Empty Then
                If StrToDate(aMaxShouldDate2) > StrToDate(aCloseDate) Then
                    aShouldDate2 = Format(aCloseDate & "", "yyyymmdd")
                Else
                    aShouldDate2 = Format(CDate(aMaxShouldDate2) + 1 & "", "yyyymmdd")
                End If
            Else
                 aShouldDate2 = Format(CDate(aMaxShouldDate2) + 1 & "", "yyyymmdd")
            End If
        End If
    End If
    On Error Resume Next
    CloseRecordset rsDate
    Exit Sub
'    If aCloseDate <> Empty Then
'        strSO033SQL = "Select Max(RealStopDate) MaxDate,Count(*) Cnt From " & GetOwner & "SO033" & _
'                    " Where CustId=" & gCustId & _
'                    " And UCcode is Not Null " & _
'                    " And CitemCode In(" & aCitemCodes & ")"
'
'        If Not GetRS(rsDate, strSO033SQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
'        If Not rsDate.EOF Then
'            If Val(rsDate("Cnt")) > 0 Then
'                If CDate(rsDate(0).Value) > aCloseDate Then
'                    aShouldDate2 = Format(aCloseDate & "", "yyyymmdd")
'                Else
'                    aShouldDate2 = Format(CDate(rsDate(0) + 1) & "", "yyyymmdd")
'                End If
'            End If
'        End If
'
'    End If
'    On Error Resume Next
'    CloseRecordset rsDate
'    Exit Sub
ChkErr:
    ErrSub FormName, "GetUseDate"
End Sub
Public Sub Update03Period(ByRef rs03 As ADODB.Recordset, _
    ByVal aPeriod As String, ByVal aClctDate As String, ByVal blnReturn As Boolean)
    Dim blnCallExecuteSingleTransfer As Boolean
    blnCallExecuteSingleTransfer = False
    Dim RetMsg As String
  On Error GoTo ChkErr
    If blnReturn Then
        If rs03.State = adStateClosed Then Exit Sub
        rs03.MoveFirst
        Do While Not rs03.EOF
            rsSO0304.Find "ROWID='" & rs03("RowId") & "'", , adSearchForward, 1
           
            If aPeriod <> "" Then
                rs03("Period") = rsSO0304("Period")
                rs03("Amount") = rsSO0304("Amount")
            End If
            If Len(aClctDate) > 0 Then
                If Len(rsSO0304("CLCTSTOPDATE") & "") = 0 Then
                    rs03("CLCTSTOPDATE") = Null
                Else
                    rs03("CLCTSTOPDATE") = rsSO0304("CLCTSTOPDATE")
                End If
            End If
            
            rs03("ClctDate") = rsSO0304("ClctDate") & ""
            rs03("StopDate") = rsSO0304("StopDate") & ""
            rs03.Update
            '#6943 如果有異動期數或收費截止日再判斷是否長繳別,如果是要呼叫Stan的後端去異動SO003A By Kin 2014/12/03
            If Val(GetRsValue("Select Nvl(LONGPAYflag,0) LONGPAYflag From " & GetOwner & "SO003 Where RowId ='" & rs03("RowId") & "'", gcnGi)) = 1 Then
                Call ExecuteSingleTransfer(garyGi(1), gCustId, rs03("CitemCode"), rs03("FaciSeqNo"), RetMsg)
            End If
            
            rs03.MoveNext
        Loop
    Else
        rs03.MoveFirst
        Do While Not rs03.EOF
            blnCallExecuteSingleTransfer = False
            RetMsg = ""
            'If aPeriod <> "" Or Len(aClctDate) > 0 Then blnCallExecuteSingleTransfer = True
            
            If aPeriod <> "" Then
                blnCallExecuteSingleTransfer = True
                If Val(rs03("CustAllot")) <> 1 Then
                    If Val(rs03("Amount")) <> 0 Then
                        rs03("Amount") = (Val(rs03("Amount")) / Val(rs03("Period"))) * Val(aPeriod)
                    End If
                    rs03("Period") = aPeriod
                End If
            End If
            If Len(aClctDate) > 0 Then
                blnCallExecuteSingleTransfer = True
                rs03("CLCTSTOPDATE") = aClctDate & ""
            End If
            rs03.Update
            '#6943 如果有異動期數或收費截止日再判斷是否長繳別,如果是要呼叫Stan的後端去異動SO003A By Kin 2014/12/03
            If blnCallExecuteSingleTransfer Then
                If Val(GetRsValue("Select Nvl(LONGPAYflag,0) LONGPAYflag From " & GetOwner & "SO003 Where RowId ='" & rs03("RowId") & "'", gcnGi)) = 1 Then
                    Call ExecuteSingleTransfer(garyGi(1), gCustId, rs03("CitemCode"), rs03("FaciSeqNo"), RetMsg)
                End If
            End If
            rs03.MoveNext
        Loop
    End If
    Exit Sub
ChkErr:
    ErrSub FormName, "Update03Period"
End Sub
Public Function GetTmpFiles2(ByVal intTables As Integer) As String
  On Error GoTo ChkErr
    Dim i As Integer
    Dim aAryName2 As String
    Dim strTmp As String
    aAryName2 = Empty
    strTmp = Empty
    For i = 0 To intTables - 1
        strTmp = GetTmpViewName
        aAryName2 = IIf(aAryName2 = Empty, strTmp, aAryName2 & "," & strTmp)
    Next i
    GetTmpFiles2 = aAryName2
  Exit Function
ChkErr:
    ErrSub FormName, "GetTmpFiles2"
End Function
Public Function GetTmpFiles(ByVal intTables As Integer, ByRef aTmpSO032 As String) As String
  On Error GoTo ChkErr
    Dim i As Long
    Dim strTmp As String
    Dim aAryName As String
    aAryName = Empty
    strTmp = Empty
    aTmpSO032 = Empty
    For i = 0 To intTables - 1
        Select Case i
            Case 5
                strTmp = "TMP017"
            Case 8, 9, 10, 11
                strTmp = "TMP0" & i + 5
            Case Else
                strTmp = GetTmpViewName
        End Select
        
        If i = 5 Then aTmpSO032 = strTmp
        aAryName = IIf(aAryName = Empty, strTmp, aAryName & "," & strTmp)
    Next i
    GetTmpFiles = aAryName
    Exit Function
ChkErr:
    ErrSub FormName, "GetTmpFiles"
End Function
Public Function Run_SF_ACCOUNTING(ByVal cnConn As ADODB.Connection, ByVal p_YM As String, ByVal P_DAY1 As Long, _
ByVal P_DAY2 As Long, ByVal p_CompSQL As String, ByVal p_CitemSQL As String, _
ByVal p_ClassSQL As String, ByVal P_AREASQL As String, ByVal p_ServSQL As String, _
ByVal p_ClctAreaSQL As String, ByVal P_MDUSQL As String, ByVal p_StrtSQL As String, _
ByVal p_AMduSQL As String, p_ProSQL, _
ByVal P_ADDRTYPE As Long, ByVal p_GENOVERDUE As Long, ByVal p_GENPRCUST As Long, _
ByVal P_TOENDDATE As Long, ByVal P_CUSTIDLIST As String, ByVal P_UPDEN As String, _
ByVal P_UPDNAME As String, ByVal p_SERVICETYPE As String, ByVal P_BANKCODE As String, _
ByVal p_CMSQL As String, ByVal p_StopPRCust As Long, ByVal p_FileName As String, _
ByVal p_FaciSeqNoSQL, ByVal p_PayKind As String, ByRef p_RetMsg As String) As Integer
On Error GoTo ChkErr
        Dim ComSF_ACCOUNTING As New ADODB.Command
        If cnConn Is Nothing Then
                'MsgBox vmsgSendConn, vbCritical, vmsgPrompt
                Exit Function
        End If

        With ComSF_ACCOUNTING
                .Parameters.Append .CreateParameter("RETURN_VALUE", adVarNumeric, adParamReturnValue)
                .Parameters.Append .CreateParameter("P_YM", adVarChar, adParamInput, 2000, p_YM)
                .Parameters.Append .CreateParameter("P_DAY1", adVarNumeric, adParamInput, , P_DAY1)
                .Parameters.Append .CreateParameter("P_DAY2", adVarNumeric, adParamInput, , P_DAY2)
                .Parameters.Append .CreateParameter("P_COMPSQL", adVarChar, adParamInput, 2000, p_CompSQL)
                .Parameters.Append .CreateParameter("P_CITEMSQL", adVarChar, adParamInput, 2000, p_CitemSQL)
                .Parameters.Append .CreateParameter("P_CLASSSQL", adVarChar, adParamInput, 2000, p_ClassSQL)
                .Parameters.Append .CreateParameter("P_AREASQL", adVarChar, adParamInput, 2000, P_AREASQL)
                .Parameters.Append .CreateParameter("P_SERVSQL", adVarChar, adParamInput, 2000, p_ServSQL)
                .Parameters.Append .CreateParameter("P_CLCTAREASQL", adVarChar, adParamInput, 2000, p_ClctAreaSQL)
                .Parameters.Append .CreateParameter("P_MDUSQL", adVarChar, adParamInput, 2000, P_MDUSQL)
                .Parameters.Append .CreateParameter("P_STRTSQL", adVarChar, adParamInput, 2000, p_StrtSQL)
                .Parameters.Append .CreateParameter("P_AMDUSQL", adVarChar, adParamInput, 2000, p_AMduSQL)
                .Parameters.Append .CreateParameter("p_PROSQL", adVarChar, adParamInput, 2000, p_ProSQL)
                .Parameters.Append .CreateParameter("P_ADDRTYPE", adVarNumeric, adParamInput, , P_ADDRTYPE)
                .Parameters.Append .CreateParameter("P_GENOVERDUE", adVarNumeric, adParamInput, , p_GENOVERDUE)
                .Parameters.Append .CreateParameter("P_GENPRCUST", adVarNumeric, adParamInput, , p_GENPRCUST)
                .Parameters.Append .CreateParameter("P_TOENDDATE", adVarNumeric, adParamInput, , P_TOENDDATE)
                .Parameters.Append .CreateParameter("P_CUSTIDLIST", adVarChar, adParamInput, 2000, P_CUSTIDLIST)
                .Parameters.Append .CreateParameter("P_UPDEN", adVarChar, adParamInput, 2000, P_UPDEN)
                .Parameters.Append .CreateParameter("P_UPDNAME", adVarChar, adParamInput, 2000, P_UPDNAME)
                .Parameters.Append .CreateParameter("P_SERVICETYPE", adVarChar, adParamInput, 2000, p_SERVICETYPE)
                .Parameters.Append .CreateParameter("P_BANKCODE", adVarChar, adParamInput, 2000, P_BANKCODE)
                .Parameters.Append .CreateParameter("P_CMSQL", adVarChar, adParamInput, 2000, p_CMSQL)
                .Parameters.Append .CreateParameter("P_STOPPRCUST", adNumeric, adParamInput, , p_StopPRCust)
                .Parameters.Append .CreateParameter("p_FileName", adVarChar, adParamInput, 2000, p_FileName)
                .Parameters.Append .CreateParameter("p_FaciSeqNoSQL", adVarChar, adParamInput, 4000, p_FaciSeqNoSQL)
                .Parameters.Append .CreateParameter("p_PayKindSql", adVarChar, adParamInput, 2000, p_PayKind)
                .Parameters.Append .CreateParameter("P_RETMSG", adVarChar, adParamOutput, 2000, p_RetMsg)
                Set .ActiveConnection = cnConn
                .CommandText = "SF_ACCOUNTING"
                .CommandType = adCmdStoredProc
                .Execute
                p_RetMsg = .Parameters("P_RETMSG").Value & ""
                Run_SF_ACCOUNTING = Val(.Parameters("Return_value").Value & "")
        End With
        Exit Function
ChkErr:
    ErrSub FormName, "Run_SF_ACCOUNTING"
End Function
Public Function ReplaceT(ByVal aTmpSO032 As String) As String
    On Error GoTo ChkErr
    Dim rsBill As New ADODB.Recordset
    Dim rsServiceType As New ADODB.Recordset
    Dim strSeqNo As String
    Dim strBillNo As String
    Dim strRetString As String
    strRetBillNo = Empty
    '******************************************************************
    '#4158 同一個服務別產生出來的T單要一樣 By Kin 2008/10/21
    '#6132 tmp檔不串Owner會出錯 By Kin 2011/09/30
    If Not GetRS(rsServiceType, "Select ServiceType From " & GetOwner & aTmpSO032 & " Group By ServiceType", gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
    Do While Not rsServiceType.EOF
        strBillNo = Empty
        strBillNo = GetInvoiceNo("T", rsServiceType("ServiceType") & "")
        '#6132 tmp檔不串Owner會出錯 By Kin 2011/09/30
        If Not GetRS(rsBill, "Select * From " & GetOwner & aTmpSO032 & " Where ServiceType='" & rsServiceType("ServiceType") & "'", gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
        Do While Not rsBill.EOF
            rsBill("BillNo") = strBillNo
            rsBill.Update
            rsBill.MoveNext
        Loop
        strRetString = IIf(strRetString = Empty, "'" & strBillNo & "'", strRetString & ",'" & strBillNo & "'")
        rsServiceType.MoveNext
    Loop
    '******************************************************************
    ReplaceT = strRetString
    On Error Resume Next
    CloseRecordset rsBill
    CloseRecordset rsServiceType
    Exit Function
ChkErr:
  Call ErrSub(FormName, "ReplaceT")
End Function
Public Sub DelTmp(ByVal aAryName As String, ByVal aAryName2 As String)
    Dim i As Integer
    Dim j As Integer
    Dim strTmpFiles() As String
    For j = 0 To 1
        Select Case j
            Case 0
                strTmpFiles = Split(aAryName, ",")
            Case 1
                strTmpFiles = Split(aAryName2, ",")
        End Select
        For i = LBound(strTmpFiles) To UBound(strTmpFiles)
          On Error Resume Next
            Select Case UCase(strTmpFiles(i))
                Case "TMP017", "TMP013", "TMP014", "TMP015", "TMP016"
                Case Else
                    gcnGi.Execute "Drop Table " & strTmpFiles(i)
            End Select
        Next i
    Next j
    Exit Sub
ChkErr:
  Call ErrSub(FormName, "DelTmp")
End Sub
Public Function SF_TMPTOC1(ByVal cnConn As ADODB.Connection, ByVal p_UserId As String, _
    ByVal p_CompCode As Long, ByVal p_SERVICETYPE As String, ByVal p_FileName As String, ByRef p_RetMsg As String) As Long
  On Error GoTo ChkErr
    Dim ComSF_TMPTOC1 As New ADODB.Command
    If cnConn Is Nothing Then
            'MsgBox vmsgSendConn, vbCritical, vmsgPrompt
            Exit Function
    End If

    With ComSF_TMPTOC1
            .Parameters.Append .CreateParameter("RETURN_VALUE", adVarNumeric, adParamReturnValue)
            .Parameters.Append .CreateParameter("P_USERID", adVarChar, adParamInput, 2000, p_UserId)
            .Parameters.Append .CreateParameter("P_COMPCODE", adVarNumeric, adParamInput, 2000, p_CompCode)
            .Parameters.Append .CreateParameter("P_SERVICETYPE", adVarChar, adParamInput, 2000, p_SERVICETYPE)
            .Parameters.Append .CreateParameter("p_FileName", adVarChar, adParamInput, 2000, p_FileName)
            .Parameters.Append .CreateParameter("P_RETMSG", adVarChar, adParamOutput, 2000, p_RetMsg)
            
            Set .ActiveConnection = cnConn
            .CommandText = "SF_TMPTOC1"
            .CommandType = adCmdStoredProc
            .Execute
            p_RetMsg = .Parameters("P_RETMSG").Value & ""
            SF_TMPTOC1 = Val(.Parameters("Return_value").Value & "")
    End With
    Exit Function
ChkErr:
        ErrSub FormName, "SF_TMPTOC1"
End Function
Public Function RunCmdAccount(ByRef aSource As ADODB.Recordset, ByRef aMsg As String, _
            ByRef aTmpSO032 As String, Optional intGenOverdue As Integer = 0, _
            Optional ByVal strClctStopDate As String = Empty, _
            Optional ByVal strPeriod As String = Empty) As Boolean
  
  On Error GoTo ChkErr
    
    Dim rsClone As New ADODB.Recordset
    Dim strQry As String
    Dim lngRetCode As Long
    Dim strRetMsg As String
    Dim strTmpFiles() As String
    
    Dim strFaciSeqNo As String
    Dim strShouldDate1 As String
    Dim strShouldDate2 As String
    Dim strServiceTypes As String
    Dim strCitemCodes As String
    
    Dim i As Long
    Dim j As Long
    
    
    Dim blnOK As Boolean
    Dim blnSuccess As Boolean
    strRowIds = Empty
    strRetBillNo = Empty
    strFaciSeqNo = Empty
    strNewClctStopDate = Empty
    strNewPeriod = Empty
    'blnOK = False
    'blnSuccess = False
    'blnTrans = obj.uTrans
    strNewClctStopDate = strClctStopDate
    strNewPeriod = strPeriod
    If aSource.RecordCount = 0 Then Exit Function
    
    'Screen.MousePointer = vbHourglass
    strRowIds = Empty
    strRetMsg = Empty
    lngRetCode = 0
    strFaciSeqNo = Empty
    aMsg = Empty
    aTmpSO032 = Empty
    Set rsClone = aSource.Clone
    rsClone.MoveFirst
    '挑選選擇資料的RowId與設備流水號
    Call GetRowIdAndFaci(rsClone, strRowIds, strFaciSeqNo)

    'If strRowIds = "" Then MsgBox "請選擇資料！", vbInformation, "訊息": GoTo Fin
    If strRowIds = "" Then aMsg = "請選擇資料！": GoTo Fin
    Set rsSO003 = GetRs003(strRowIds)
    Call GetShouldDate(strRowIds, strShouldDate1, strShouldDate2) '截取選取收費項目的ClctDate最大值與最小值,帶入後端的ShouldDate1與ShouldDate2
    strServiceTypes = GetServiceTypes(strRowIds) '截取選取收費項目的服務別
    strCitemCodes = GetCitemCodes(strRowIds)  '截取選取的收費項目
    '****************************************************************
    '#4306 判斷截止是要用結清日、待收日或下收日 By Kin 2009/03/25
    '#5815 從結清進來才進行滾帳 By Kin 2010/10/21
    If strCloseDate <> Empty Then
        Call GetUseDate(strCloseDate, strCitemCodes, strShouldDate2, rsSO003)
    End If
    '****************************************************************
    'If Not blnTrans Then gcnGi.BeginTrans
    '如果有輸入期數,要先進行UPDATE,等後端做完後再將資料UPDATE回來 By Kin 2008/08/06
    '#5220 增加一個收費截止日 By Kin 2009/10/06
    If strPeriod <> "" Or strClctStopDate <> "" Or strCloseDate <> "" Then
        Call Update03Period(rsSO003, strPeriod, strClctStopDate, False)
    End If
    '呼叫出帳後端
    strAryName = GetTmpFiles(20, strTmpSO032)
    'Call GetTmpFiles(20)
    '*****************************************
    '#4142 過入要多加三個Table By Kin 2008/12/04 For Lawrence
    strAryName2 = GetTmpFiles2(3)
    'Call GetTmpFiles2(3)
    '*****************************************
    '#6883 後端參數地址依據原本固定傳入1,改為抓取參數值 By Kin 2014/10/24
    Dim intAddrType As Integer
    Dim aSQL As String
    aSQL = "Select Nvl(para14,1) para14 From " & GetOwner & "SO043  Where RowNum = 1"
    intAddrType = Val(GetRsValue(aSQL, gcnGi))
    
    lngRetCode = Run_SF_ACCOUNTING(gcnGi, Format(RightNow, "yyyymm"), Val(strShouldDate1), Val(strShouldDate2), _
                             " In (" & gCompCode & " )", strCitemCodes, "", "", "", "", "", "", "", "", intAddrType _
                             , intGenOverdue, 0, 0, str(gCustId), garyGi(0), garyGi(1), _
                             strServiceTypes, "", "", 0, strAryName, strFaciSeqNo, "", strRetMsg)


    If lngRetCode < 0 Then
        'If Not blnTrans Then gcnGi.RollbackTrans
        aMsg = strRetMsg
        RunCmdAccount = False
        GoTo Fin
    Else
        aMsg = strRetMsg
        aTmpSO032 = strTmpSO032
        RunCmdAccount = True
    End If
    
Fin:
    On Error Resume Next
    'CloseRecordset rsSO003
    'CloseRecordset rsClone
    
    '要連過入的動態檔名一起刪除 By Kin 2008/12/04
    'Call DelTmp(strAryName, strAryName2)
    'blnSuccess = True
    'Screen.MousePointer = vbDefault
    
    If strCloseDate <> Empty Then blnNoDefine = True
    Exit Function
ChkErr:
    blnOK = False
    blnSuccess = False
    'gcnGi.RollbackTrans
    'Screen.MousePointer = vbDefault
    ErrSub FormName, "RunCmdAccount"
  Exit Function
End Function
'#6943 有修改期數,要呼叫後端異動SO003A的期數 By Kin 2014/12/03
Private Function ExecuteSingleTransfer(ByVal strUpdEn As String, ByVal strCustId As String, _
                        ByVal strCitemCode As String, ByVal strFaciSeqNo As String, ByRef RetMsg As String) As Boolean
    On Error GoTo ChkErr

    Dim p_RetMsg As String
    Dim ComSingleTransfer As New ADODB.Command
    With ComSingleTransfer
        
       .Parameters.Append .CreateParameter("RETURN_VALUE", adVarNumeric, adParamReturnValue)
       .Parameters.Append .CreateParameter("P_UPDEN", adVarChar, adParamInput, 2000, strUpdEn)
       .Parameters.Append .CreateParameter("P_CUSTID", adVarNumeric, adParamInput, , Val(strCustId))
        .Parameters.Append .CreateParameter("P_CITEMCODE", adVarNumeric, adParamInput, , Val(strCitemCode))
        .Parameters.Append .CreateParameter("P_FACISEQNO", adVarChar, adParamInput, 2000, strFaciSeqNo)
       .Parameters.Append .CreateParameter("P_RETMSG", adVarChar, adParamOutput, 2000, p_RetMsg)
         Set .ActiveConnection = gcnGi
        .CommandText = "PK_LongPayTransfer.ExecuteSingleTransfer"
        .CommandType = adCmdStoredProc
        .Execute
        p_RetMsg = .Parameters("P_RETMSG").Value & ""
'        If Val(.Parameters("Return_value").Value & "") < 0 Then
'            MsgBox p_RetMsg, vbCritical, "錯誤"
'            ExecuteSingleTransfer = False
'            Exit Function
'        Else
'            '#6865 如果有回傳訊息，Show出回傳訊息 By Kin 2014/09/12
'            If Len(p_RetMsg & "") > 0 Then
'                MsgBox p_RetMsg, vbInformation, "訊息"
'            End If
'           ExecuteSingleTransfer = True
'        End If
    End With
    ExecuteSingleTransfer = True
    Exit Function
ChkErr:
    ErrSub "clsComm", "ExecuteSingleTransfer"
End Function

Public Function ShowSO1100U2(ByVal aTmp32 As String) As Boolean
    '#6132 tmp檔不串Owner會出錯 By Kin 2011/09/30
    If gcnGi.Execute("SELECT count(*) FROM " & GetOwner & aTmp32)(0) = 0 Then
        ShowSO1100U2 = False
    Else
        strRetBillNo = ReplaceT(aTmp32)
        ShowSO1100U2 = True
    End If
End Function
Public Function EXE_SF_TMPTOC1(ByVal aTmpSO032 As String, ByRef aMsg As String) As Boolean
  On Error GoTo ChkErr
    Dim lngRetCode As Long
    Dim strRetMsg As String
    
    lngRetCode = 0
    aMsg = Empty
    strRetMsg = Empty
    lngRetCode = SF_TMPTOC1(gcnGi, garyGi(1), gCompCode, "", aTmpSO032 & "," & strAryName2, strRetMsg)
    aMsg = strRetMsg
    If lngRetCode < 0 Then
        'If Not blnTrans Then gcnGi.RollbackTrans
        EXE_SF_TMPTOC1 = False
    Else
        EXE_SF_TMPTOC1 = True
    End If
    Exit Function
ChkErr:
    ErrSub FormName, "EXE_SF_TMPTOC1"
End Function
Private Function ReturnSO003() As Boolean
  On Error GoTo ChkErr
    Call Update03Period(rsSO003, strNewPeriod, strNewClctStopDate, True)
    ReturnSO003 = True
    Exit Function
ChkErr:
    Call ErrSub(FormName, "ReturnSO003")
End Function
Public Function DelTmpTable() As Boolean
  On Error Resume Next
    Call DelTmp(strAryName, strAryName2)
    DelTmpTable = True
End Function
Public Function FinProcess() As Boolean
  On Error GoTo ChkErr
    Call ReturnSO003
    FinProcess = True
    Exit Function
ChkErr:
  Call ErrSub(FormName, "FinProcess")
End Function

Public Function GetRetRecordSet(ByVal aSourceTable As String) As ADODB.Recordset
  On Error GoTo ChkErr
    Dim rsReturn As New ADODB.Recordset
    Set GetRetRecordSet = Nothing
    '#6132 tmp檔不串Owner會出錯 By Kin 2011/09/30
    If Not GetRS(rsReturn, "SELECT * FROM " & GetOwner & aSourceTable, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
    Set GetRetRecordSet = rsReturn
    Exit Function
ChkErr:
  Call ErrSub(FormName, "GetRetRecordSet")
End Function
