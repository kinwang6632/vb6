Attribute VB_Name = "mduVodTools"
Option Explicit
Public strRecordProcedureFile As String
Public objFileSystem As Scripting.FileSystemObject
Public Const strMoveDir As String = "SucessVOD"
Public Const strRecordName As String = "VODRecord_Kin.TXT"
Public objRecordWrite As TextStream
Public IsCancelCmd As Boolean
Private Declare Function CoCreateGuid Lib "OLE32.DLL" (pGuid As GUID) As Long
Public FPausetime As Integer
Public intClctStopDateType As Integer
Public strOTTOwner As String
Public intOTTCompID As Integer
Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type
'#7031 增加是否要過濾B15命令資料 By Kin 2015/04/28
Public Sub GetClctStopDateType()
  On Error GoTo ChkErr
     intClctStopDateType = Val(GetRsValue("Select Nvl(ClctStopDateType,0) From " & GetOwner & "SO041"))
  Exit Sub
ChkErr:
   ErrSub FormName, "GetClctStopDateType"
End Sub
'判斷PVRCMDType,如果不是0則smith的命令不發送 By Kin 2015/05/11
Public Function qryPVRCMDType(ByVal strModelCode As String) As Integer
  On Error GoTo ChkErr
    Dim rsPVRCMDType As New ADODB.Recordset
    If Len(strModelCode) = 0 Then qryPVRCMDType = 0: GoTo 88
    If Not GetRS(rsPVRCMDType, "Select Nvl(PVRCMDType,0) PVRCMDType From " & GetOwner & "CD043 Where CodeNo = " & strModelCode, _
                            gcnGi, adUseClient, adOpenKeyset, adLockPessimistic) Then qryPVRCMDType = 0
    If rsPVRCMDType.RecordCount = 0 Then
        qryPVRCMDType = 0
    Else
        qryPVRCMDType = Val(rsPVRCMDType("PVRCMDType") & "")
    End If
88:
    On Error Resume Next
    Call CloseRecordset(rsPVRCMDType)
    Exit Function
ChkErr:
    ErrSub FormName, "qryPVRCMDType"
End Function
'功能說明: 利用STBNO OR ICCNO取得設備的MODELCODE，判別設備是否為NSTV
'    備註: 回傳 1=NSTV 0=NAGRA
Public Function GetDTVCAType(ByVal strSTB As String, ByVal strICC As String) As Integer
    On Error GoTo ChkErr
    Dim strModelCode As String
    If strSTB = "" And strICC <> "" Then
        strSTB = GetRsValue("Select STBSNO From " & GetOwner & "SOAC0201C Where SMARTCARDNO='" & strICC & "'") & ""
    End If
    If strSTB = "" Then GetDTVCAType = 0: Exit Function
    
    strModelCode = GetRsValue("Select ModelCode From " & GetOwner & "SOAC0201A Where FACISNO='" & strSTB & "'") & ""
    If strModelCode = "" Then GetDTVCAType = 0: Exit Function
    If Val(GetRsValue("Select FuncFlag08 From " & GetOwner & "CD043 Where CodeNO=" & strModelCode) & "") = 1 Then
    'GetDTVCAType = Val(GetRsValue("Select FuncFlag08 From " & GetOwner & "CD043 Where CodeNO=" & strModelCode) & "")
        GetDTVCAType = Val(GetRsValue("Select Nvl(VodType,0) From " & GetOwner & "CD043 Where CodeNO=" & strModelCode) & "")
    Else
        GetDTVCAType = 0
    End If
    Exit Function
ChkErr:
    ErrSub FormName, "GetDTVCAType"
End Function
Public Function GetTuner(ByVal strSTB As String, ByVal strICC As String) As Integer
    On Error GoTo ChkErr
    Dim strModelCode As String
    If strSTB = "" And strICC <> "" Then
        strSTB = GetRsValue("Select STBSNO From " & GetOwner & "SOAC0201C Where SMARTCARDNO='" & strICC & "'") & ""
    End If
    If strSTB = "" Then GetTuner = 0: Exit Function
    
    strModelCode = GetRsValue("Select ModelCode From " & GetOwner & "SOAC0201A Where FACISNO='" & strSTB & "'") & ""
    If strModelCode = "" Then GetTuner = 0: Exit Function
    
    GetTuner = Val(GetRsValue("Select Nvl(TunerCount,0) From " & GetOwner & "CD043 Where CodeNO=" & strModelCode) & "")
    Exit Function
ChkErr:
    ErrSub FormName, "GetTuner"
End Function
Public Function GetMaxSizeCodeCitemCode(ByVal CitemCodes As String) As String
  On Error GoTo ChkErr
    Dim aSQL As String
    Dim rsCitemCode As New ADODB.Recordset
    aSQL = "Select CodeNo  From " & GetOwner & "CD019 Where DVRSizeCode In ( " & _
                    " Select CodeNo From ( Select CodeNo From " & GetOwner & "CD102  " & _
                    " Where Codeno in ( Select Nvl(DVRSizeCode,-1) SizeCode From " & GetOwner & "CD019 " & _
                    " Where CodeNo in (" & CitemCodes & " )) Order by DVRSize Desc) where Rownum<=1) " & _
                    " And Codeno in (" & CitemCodes & ") And RowNum = 1"
    If Not GetRS(rsCitemCode, aSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
    If rsCitemCode.RecordCount > 0 Then
        GetMaxSizeCodeCitemCode = rsCitemCode(0) & ""
    Else
        GetMaxSizeCodeCitemCode = ""
    End If
    On Error Resume Next
    Call CloseRecordset(rsCitemCode)
  Exit Function
ChkErr:
  Call ErrSub(FormName, "GetMaxSizeCodeCitemCode")
End Function
Public Function GetMaxDVRSIZECODE(ByRef rsSO004 As ADODB.Recordset) As String
  On Error GoTo ChkErr
    Dim aSQL As String
    Dim rsMax As New ADODB.Recordset
    aSQL = GetChargeRecordCitemCodeSQL(rsSO004)
    aSQL = "Select CodeNo From " & GetOwner & "CD019 " & _
            " Where CodeNo In (" & aSQL & ") " & _
            " Order By DVRSIZECODE Desc "
    If Not GetRS(rsMax, aSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
    If rsMax.RecordCount = 0 Then
        GetMaxDVRSIZECODE = ""
    Else
        GetMaxDVRSIZECODE = rsMax("CodeNo")
    End If
    On Error Resume Next
    Call CloseRecordset(rsMax)

'      aSQL = "Select Nvl(DVRSize*1024,0) MaxSize  From " & GetOwner & "CD102 " & _
'                " Where Type=1 and CodeNo In ( " & _
'                " Select Distinct CodeNo  From " & GetOwner & "CD024 " & _
'                    " Where IsDVR=1 and nvl(StopFlag,0)=0 and CodeNo In (" & _
'                    " Select CodeNo From " & GetOwner & "CD019A where CitemCode In (" & aSQL & ") ) " & _
'                    " Order By DVRSize Desc "
'    Dim rsRecordCapacity As New ADODB.Recordset
'    If Not GetRS(rsRecordCapacity, aSQL, gcnGi, adUseClient, adOpenKeyset) Then Exit Function
'    If rsRecordCapacity.RecordCount = 0 Then
'        GetRecordCapacity = 0
'    Else
'        rsRecordCapacity.MoveFirst
'        GetRecordCapacity = Val(rsRecordCapacity(0))
'    End If
'
'    On Error Resume Next
'    Call CloseRecordset(rsRecordCapacity)
    
    
  Exit Function
ChkErr:
  Call ErrSub(FormName, "GetMaxDVRSIZECODE")
End Function
Public Function GetChargeRecordCitemCodeSQL(ByRef rsSO004 As ADODB.Recordset) As String
  On Error GoTo ChkErr
    Dim aSQL As String
    Dim aBillNo As String
    aBillNo = str7SNO
    If aBillNo = "" Then
        aBillNo = "X"
    End If
    aSQL = "Select Distinct SO033.CitemCode From " & GetOwner & "SO033," & GetOwner & "CD019," & GetOwner & "CD019A," & _
            GetOwner & "CD024 " & _
            " WHERE SO033.CUSTID = " & rsSO004("CustId") & _
            " And SO033.BillNo='" & aBillNo & "' And SO033.CompCode = " & rsSO004("CompCode") & _
            " AND SO033.CITEMCODE = CD019.CODENO AND NVL(CD019.STOPFLAG,0) = 0 AND CD019.DVRSIZECODE IS NOT NULL " & _
            " AND CD019.CODENO = CD019A.CITEMCODE AND CD019A.CODENO = CD024.CODENO AND NVL(CD024.STOPFLAG,0) = 0 " & _
            " AND CD024.ISDVR = 1 "
    
    GetChargeRecordCitemCodeSQL = aSQL
  Exit Function
ChkErr:
 Call ErrSub(FormName, "GetChargeRecordCitemCode")
End Function
Public Function IsCatchUpCitemCode(ByVal CitemCodes As String) As Boolean
  On Error GoTo ChkErr
    Dim aSQL As String
'    aSQL = "Select Count(1) Cnt From " & GetOwner & "CD024," & GetOwner & "CD019," & GetOwner & "CD019A " & _
'            " Where CD019.CodeNo In ( " & CitemCodes & ") " & _
'            " And CD019.CodeNo = CD019A.CitemCode And Nvl(CD019.StopFlag,0) = 0 And CD019.DVRSizeCode Is Not Null " & _
'            " And CD019A.CodeNo = CD024.CodeNo And Nvl(CD024.StopFlag,0) = 0 " & _
'            " And CD024.IsCatchUp = 1 "
      aSQL = "Select Count(1) Cnt From " & GetOwner & "CD024," & GetOwner & "CD019," & GetOwner & "CD019A " & _
            " Where CD019.CodeNo In ( " & CitemCodes & ") " & _
            " And CD019.CodeNo = CD019A.CitemCode And Nvl(CD019.StopFlag,0) = 0 " & _
            " And CD019A.CodeNo = CD024.CodeNo And Nvl(CD024.StopFlag,0) = 0 " & _
            " And CD024.IsCatchUp = 1 "
    If Val(GetRsValue(aSQL, gcnGi) & "") > 0 Then
        IsCatchUpCitemCode = True
    Else
        IsCatchUpCitemCode = False
    End If
  Exit Function
ChkErr:
    Call ErrSub(FormName, "IsCatchUpCitemCode")
End Function
Public Function IsCatchUp(ByRef rsSO004 As ADODB.Recordset) As Boolean
On Error GoTo ChkErr
    Dim aSQL As String
    
    aSQL = "Select Distinct SO003.CitemCode From " & GetOwner & "SO003," & GetOwner & "CD019," & GetOwner & "CD019A," & _
            GetOwner & "CD024 " & _
            " WHERE SO003.CUSTID = " & rsSO004("CustId") & _
            " AND NVL(SO003.STOPFLAG,0) = 0 AND SO003.FACISEQNO = '" & rsSO004("SeqNo") & "' AND SO003.SERVICETYPE = 'D' " & _
            " AND SO003.CITEMCODE = CD019.CODENO AND NVL(CD019.STOPFLAG,0) = 0 " & _
            " AND CD019.CODENO = CD019A.CITEMCODE AND CD019A.CODENO = CD024.CODENO AND NVL(CD024.STOPFLAG,0) = 0 " & _
            " And SO003.CompCode=" & rsSO004("CompCode") & _
            " AND CD024.IsCatchUp = 1 " & _
            " UNION ALL " & _
            "Select Distinct SO033.CitemCode From " & GetOwner & "SO033," & GetOwner & "CD019," & GetOwner & "CD019A," & _
            GetOwner & "CD024 " & _
            " WHERE SO033.CUSTID = " & rsSO004("CustId") & _
            " AND SO033.FACISEQNO = '" & rsSO004("SeqNo") & "' AND SO033.SERVICETYPE = 'D' " & _
            " AND SO033.CITEMCODE = CD019.CODENO AND NVL(CD019.STOPFLAG,0) = 0 " & _
            " AND CD019.CODENO = CD019A.CITEMCODE AND CD019A.CODENO = CD024.CODENO AND NVL(CD024.STOPFLAG,0) = 0 " & _
            " And SO033.CompCode=" & rsSO004("CompCode") & _
            " AND CD024.IsCatchUp = 1 " & _
            IIf(Len(str7SNO) > 0, " AND BillNo = '" & str7SNO & "' ", "")
    aSQL = "Select Count(1) Cnt From (" & aSQL & ")"
     If Val(GetRsValue(aSQL, gcnGi)) > 0 Then
        IsCatchUp = True
     Else
        IsCatchUp = False
     End If
    
  Exit Function
ChkErr:
  Call ErrSub(FormName, "IsCatchUp")
End Function
Public Sub SetOTTgiList(objGilist As Object, strFldName1 As String, strFldName2 As String, _
            strTableName As String, Optional lngTop As Long, Optional lngLeft As Long, Optional lngWidth1 As Long, _
            Optional lngWidth2 As Long, Optional lngFldLen1 As Long, Optional lngFldLen2 As Long, _
            Optional strWhere As String, Optional blnStopFlag As Boolean = False)
 '欄位1,欄位2,表格名稱,Top,Left,Width1,Width2,FldLen1,FldLen2
  On Error GoTo ChkErr
    objGilist.SetFldName1 strFldName1
    objGilist.SetFldName2 strFldName2
    objGilist.SetTableName strOTTOwner & strTableName
    objGilist.Filter = strWhere
    If lngTop > 0 Then objGilist.Top = lngTop
    If lngLeft > 0 Then objGilist.Left = lngLeft
    If lngWidth1 > 0 Then objGilist.FldWidth1 = lngWidth1
    If lngWidth2 > 0 Then objGilist.FldWidth2 = lngWidth2
    If lngFldLen1 > 0 Then objGilist.FldLen1 = lngFldLen1
    If lngFldLen2 > 0 Then objGilist.FldLen2 = lngFldLen2
    objGilist.FilterStop = blnStopFlag
    Call objGilist.SendConn(gcnGi)
    
  Exit Sub
ChkErr:
    Call ErrSub(FormName, "SetgiList")
End Sub
Public Function GetRecordCitemCodeSQL(ByRef rsSO004 As ADODB.Recordset) As String
  On Error GoTo ChkErr
    Dim aSQL As String
    aSQL = "Select Distinct SO003.CitemCode From " & GetOwner & "SO003," & GetOwner & "CD019," & GetOwner & "CD019A," & _
            GetOwner & "CD024 " & _
            " WHERE SO003.CUSTID = " & rsSO004("CustId") & _
            " AND NVL(SO003.STOPFLAG,0) = 0 AND SO003.FACISEQNO = '" & rsSO004("SeqNo") & "' AND SO003.SERVICETYPE = 'D' " & _
            " AND SO003.CITEMCODE = CD019.CODENO AND NVL(CD019.STOPFLAG,0) = 0 AND CD019.DVRSIZECODE IS NOT NULL " & _
            " AND CD019.CODENO = CD019A.CITEMCODE AND CD019A.CODENO = CD024.CODENO AND NVL(CD024.STOPFLAG,0) = 0 " & _
            " And SO003.CompCode=" & rsSO004("CompCode") & _
            " AND CD024.ISDVR = 1 " & _
            " UNION ALL " & _
            "Select Distinct SO033.CitemCode From " & GetOwner & "SO033," & GetOwner & "CD019," & GetOwner & "CD019A," & _
            GetOwner & "CD024 " & _
            " WHERE SO033.CUSTID = " & rsSO004("CustId") & _
            " AND SO033.FACISEQNO = '" & rsSO004("SeqNo") & "' AND SO033.SERVICETYPE = 'D' " & _
            " AND SO033.CITEMCODE = CD019.CODENO AND NVL(CD019.STOPFLAG,0) = 0 AND CD019.DVRSIZECODE IS NOT NULL " & _
            " AND CD019.CODENO = CD019A.CITEMCODE AND CD019A.CODENO = CD024.CODENO AND NVL(CD024.STOPFLAG,0) = 0 " & _
            " And SO033.CompCode=" & rsSO004("CompCode") & _
            " AND CD024.ISDVR = 1 "
    
    GetRecordCitemCodeSQL = aSQL
 
  Exit Function
ChkErr:
  Call ErrSub(FormName, "GetRecordCitemCodeSQL")
End Function

Public Function GetChargeRecordCapacity(ByRef rsSO004 As ADODB.Recordset) As String
  On Error GoTo ChkErr
    Dim aSQL As String
    Dim rsRecordCapacity As New ADODB.Recordset
    aSQL = GetChargeRecordCitemCodeSQL(rsSO004)
    aSQL = "Select Nvl(DVRSize*1024,0) DVRSize  From " & GetOwner & "CD102 " & _
                " Where Type=1 and CodeNo In ( " & _
                " Select Distinct CodeNo  From " & GetOwner & "CD024 " & _
                    " Where IsDVR=1 and nvl(StopFlag,0)=0 and CodeNo In (" & _
                    " Select CodeNo From " & GetOwner & "CD019A where CitemCode IN (" & aSQL & ") ) " & _
                    " Order By DVRSize Desc "
    
    If Not GetRS(rsRecordCapacity, aSQL, gcnGi, adUseClient, adOpenKeyset) Then Exit Function
    If rsRecordCapacity.RecordCount = 0 Then
        GetChargeRecordCapacity = 0
    Else
        rsRecordCapacity.MoveFirst
        GetChargeRecordCapacity = Val(rsRecordCapacity(0))
    End If
    
    On Error Resume Next
    Call CloseRecordset(rsRecordCapacity)
    
'    aSQL = "Select CodeNo From " & GetOwner & "CD019 " & _
'            " Where CodeNo In (" & aSQL & ") " & _
'            " Order By DVRSIZECODE Desc "
'    If Not GetRS(rsMax, aSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
'    If rsMax.RecordCount = 0 Then
'        GetMaxDVRSIZECODE = ""
'    Else
'        GetMaxDVRSIZECODE = rsMax("CodeNo")
'    End If
'    On Error Resume Next
'    Call CloseRecordset(rsMax)
  Exit Function
ChkErr:
  Call ErrSub(FormName, "GetChargeRecordCapacity")
End Function
Public Function GetRecordCapacity(ByVal CitemCode As String) As Double
  On Error GoTo ChkErr
'   GetRecordCapacity = Val(GetRsValue("Select DVRSize From " & GetOwner & "CD102 " & _
'                                    " Where Type = 1 AND CodeNo = (select DVRSizeCode From " & GetOwner & "CD019 " & _
'                                                    " Where CodeNo = " & CitemCode & ")") & "")
    Dim aSQL As String
    
'    aSQL = "Select Nvl(DVRSize*1024,0) DVRSize  From " & GetOwner & "CD102 " & _
'                " Where Type=1 and CodeNo In ( " & _
'                " Select Distinct CodeNo  From " & GetOwner & "CD024 " & _
'                    " Where IsDVR=1 and nvl(StopFlag,0)=0 and CodeNo In (" & _
'                    " Select CodeNo From " & GetOwner & "CD019A where CitemCode=" & CitemCode & ") ) " & _
'                    " Order By DVRSize Desc "
    aSQL = "Select ( Nvl(DVRSize,0) * 1024 ) DVRSize From " & GetOwner & "CD102 " & _
            " Where CodeNo = ( Select DVRSizeCode From " & GetOwner & "CD019 " & _
            " Where CodeNo In ( " & CitemCode & ")) Order By DVRSize Desc"
'    aSQL = "Select (Nvl(DVRSizeCode,0) * 1024) DVRSize From " & GetOwner & "CD019 " & _
'            " Where CodeNo = " & CitemCode
            
    Dim rsRecordCapacity As New ADODB.Recordset
    If Not GetRS(rsRecordCapacity, aSQL, gcnGi, adUseClient, adOpenKeyset) Then Exit Function
    If rsRecordCapacity.RecordCount = 0 Then
        GetRecordCapacity = 0
    Else
        rsRecordCapacity.MoveFirst
        GetRecordCapacity = Val(rsRecordCapacity(0))
    End If
    
    On Error Resume Next
    Call CloseRecordset(rsRecordCapacity)
      
  Exit Function
ChkErr:
  ErrSub FormName, "GetRecordCapacity"
End Function

'Public Function GetSmitOrdCitemFilter(ByVal CustId As String, ByVal FaciSNo As String, ByVal CompCode As String) As String
'  On Error GoTo ChkErr
'    Dim aSQL As String
''    aSQL = "Select Distinct CitemCode  From " & GetOwner & "SO003 " & _
''                " Where CustId=" & CustId & _
''                " And Nvl(StopFlag,0)=0 And FaciSEQNO='" & FaciSeqNo & "'" & _
''                " And ServiceType='D' " & _
''                " Minus " & _
''                " Select Distinct CitemCode From " & GetOwner & "SO005 " & _
''                " Where CustId = " & CustId & _
''                " And CvtSNo = '" & FaciSNo & "'" & _
''                " And DueDate > SysDate "
'    aSQL = "SELECT Distinct SO005.CitemCode From " & GetOwner & "SO005," & GetOwner & "CD024 " & _
'            " WHERE SO005.CUSTID = " & CustId & " AND SO005.COMPCODE = " & CompCode & _
'            " AND SO005.CvtSNo = '" & FaciSNo & "' " & _
'            " AND SO005.DueDate > SysDate AND SO005.ChCode = CD024.CODENO " & _
'            " AND CD024.CanRetWatch = 1 AND CD024.CompCode = SO005.CompCode " & _
'            " AND NVL(STOPFLAG,0) = 0  "
'    GetSmitOrdCitemFilter = aSQL
'  Exit Function
'ChkErr:
'  Call ErrSub(FormName, "GetSmitOrdCitemFilter")
'End Function


Public Function GetSmitCitemFilter(ByVal CustId As String, ByVal FaciSNo As String, _
                                    ByVal CompCode As String, ByVal IsCanRetWatch As Boolean) As String
  On Error GoTo ChkErr
     Dim aSQL As String


    aSQL = "SELECT Distinct SO005.CitemCode From " & GetOwner & "SO005," & GetOwner & "CD024 " & _
            " WHERE SO005.CUSTID = " & CustId & " AND SO005.COMPCODE = " & CompCode & _
            " AND SO005.CvtSNo = '" & FaciSNo & "' " & _
            " AND SO005.DueDate > SysDate AND SO005.ChCode = CD024.CODENO " & _
             IIf(IsCanRetWatch, " AND CD024.CanRetWatch = 1", " AND CD024.IsCatchUp = 1 ") & _
             "  AND CD024.CompCode = SO005.CompCode " & _
            " AND NVL(STOPFLAG,0) = 0  "
    GetSmitCitemFilter = aSQL
  Exit Function
ChkErr:
  Call ErrSub(FormName, "GetSmitCitemFilter")
End Function
Public Function GetSmitProductId(ByVal CitemCodes As String, ByVal CustId As String, _
                                        ByVal FaciSNo As String, _
                                        ByVal CompCode As String, ByVal IsCanRetWatch As Boolean, _
                                        ByVal FilterNormalProd As Boolean) As String
  On Error GoTo ChkErr
    Dim aSQL As String
    Dim aSQL2 As String
    Dim aCanRetWatchWhere As String
    Dim RetString As String
    Dim i As Integer
    Dim rsRet As New ADODB.Recordset
    aCanRetWatchWhere = " 1=1 "
    If IsCanRetWatch Then
        aCanRetWatchWhere = " CD024.CanRetWatch = 1 "
    End If
    aSQL = "Select Distinct Decode(VOD_PID,Null,ChannelID,VOD_PID) ChannelID From " & GetOwner & "CD024 " & _
            " Where Nvl(StopFlag,0)=0 " & _
            " And CodeNo in ( Select CodeNo From " & GetOwner & "CD019A Where CitemCode In (" & CitemCodes & "))" & _
            " And " & aCanRetWatchWhere
            
    If FilterNormalProd Then
       aSQL2 = "SELECT Distinct SO005.CitemCode From " & GetOwner & "SO005," & GetOwner & "CD024 " & _
            " WHERE SO005.CUSTID = " & CustId & " AND SO005.COMPCODE = " & CompCode & _
            " AND SO005.CvtSNo = '" & FaciSNo & "' " & _
            " AND SO005.DueDate > SysDate AND SO005.ChCode = CD024.CODENO " & _
             IIf(IsCanRetWatch, " AND CD024.CanRetWatch = 1", " AND CD024.IsCatchUp = 1 ") & _
             "  AND CD024.CompCode = SO005.CompCode " & _
            " AND NVL(STOPFLAG,0) = 0  And SO005.CitemCode Not In (" & CitemCodes & ") "
         
        
    End If
    If Len(aSQL2) > 0 Then
        aSQL = aSQL & " Minus " & _
            "Select Distinct Decode(VOD_PID,Null,ChannelID,VOD_PID) ChannelID From " & GetOwner & "CD024 " & _
            " Where Nvl(StopFlag,0)=0 " & _
            " And CodeNo in ( Select CodeNo From " & GetOwner & "CD019A Where CitemCode In (" & aSQL2 & ")) " & _
            " And " & aCanRetWatchWhere
    End If
    If Not GetRS(rsRet, aSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
    If rsRet.RecordCount = 0 Then
        RetString = Empty
    Else
        rsRet.MoveFirst
        Do While Not rsRet.EOF
            If Len(RetString) = 0 Then
                RetString = rsRet(0)
            Else
                RetString = RetString & "," & rsRet(0)
            End If
            rsRet.MoveNext
        Loop
    End If
    On Error Resume Next
    CloseRecordset rsRet
'    RetString = gcnGi.Execute(aSQL2).GetString(, , , ",")
'    If Right(RetString, 1) = "," Then
'        RetString = Mid(RetString, 1, Len(RetString) - 1)
'    End If
'    If RetString = "" Then RetString = "-1"
    GetSmitProductId = RetString
  Exit Function
ChkErr:
  Call ErrSub(FormName, "GetSmitProductId")
End Function

Public Function GetSmitChargeCitemCode(ByRef rsSO004 As ADODB.Recordset, ByVal OrdProd As Boolean, ByVal IsCanRetWatch As Boolean, ByVal OtherCitemCodes As String) As String
  On Error GoTo ChkErr
    
    Dim RetString As String
    Dim rsRet As New ADODB.Recordset
    If Not GetRS(rsRet, GetSmitChargeCitemCodeSQL(rsSO004, OrdProd, IsCanRetWatch, OtherCitemCodes), gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
    If rsRet.RecordCount = 0 Then
        RetString = "-1"
    Else
        rsRet.MoveFirst
        Do While Not rsRet.EOF
            If Len(RetString) = 0 Then
                RetString = rsRet(0)
            Else
                RetString = RetString & "," & rsRet(0)
            End If
            rsRet.MoveNext
        Loop
    End If
    'RetString = gcnGi.Execute(GetSmitChargeCitemCodeSQL(rsSO004, OrdProd, IsCanRetWatch)).GetString(, , , ",")
'    If Right(RetString, 1) = "," Then
'        RetString = Mid(RetString, 1, Len(RetString) - 1)
'    End If
'    If RetString = "" Then RetString = "-1"
    GetSmitChargeCitemCode = RetString
    On Error Resume Next
    CloseRecordset rsRet
  Exit Function
ChkErr:
  Call ErrSub(FormName, "GetSmitChargeCitemCode")
End Function


Public Function GetSmitChargeCitemCodeSQL(ByRef rsSO004 As ADODB.Recordset, _
                                            ByVal OrdProd As Boolean, ByVal IsCanRetWatch As Boolean, ByVal OtherCitemCodes As String) As String
  On Error GoTo ChkErr
    Dim aSQL As String
    Dim aBillNo As String
    aBillNo = str7SNO
    Dim aFilterClctStopDateType As String
    aFilterClctStopDateType = ""
    '#7031 B項已停用  &  SO003.停用日期(CeaseDate) < TRUNC( SYSDATE) 者,該PACK不能執行重開 By Kin 2015/04/28
    If intClctStopDateType = 1 Then
        aFilterClctStopDateType = " Not IN ( " & _
                        " Select CitemCode From " & GetOwner & "SO003 " & _
                            " Where Nvl(SO003.StopFlag,0) = 1 And Nvl(SO003.CeaseDate,TRUNC(To_Date('29991231','yyyymmdd'))) < TRUNC( SYSDATE) " & _
                            " And CustId = " & rsSO004("CustId") & " And FaciSeqNo = '" & rsSO004("SeqNo") & "' )"
    End If
'    If aBillNo = "" Then
'        aBillNo = "X"
'    End If
    '#7031 B項已停用  &  SO003.停用日期(CeaseDate) < TRUNC( SYSDATE) 者,該PACK不能執行重開 By Kin 2015/04/28
    If IsA6Command Or True Then aBillNo = ""
    aSQL = "Select Distinct SO003.CitemCode From " & GetOwner & "SO003," & _
                    GetOwner & "CD019, " & GetOwner & "CD019A," & _
                    GetOwner & "CD024 " & _
                " WHERE SO003.CUSTID = " & rsSO004("CustId") & "  AND SO003.FACISEQNO = '" & rsSO004("SeqNo") & "' " & _
                " AND SO003.SERVICETYPE = 'D'  AND SO003.CITEMCODE = CD019.CODENO AND NVL(CD019.STOPFLAG,0) = 0 " & _
                " AND CD019.CODENO = CD019A.CITEMCODE AND CD019A.CODENO = CD024.CODENO " & _
                " AND NVL(CD024.STOPFLAG,0) = 0  And SO003.CompCode = " & rsSO004("CompCode") & _
                IIf(IsCanRetWatch, "  AND CD024.CANRETWATCH = 1 ", " ") & _
                IIf(Len(aFilterClctStopDateType) > 0, " And SO003.CitemCode " & aFilterClctStopDateType, "") & _
                " Union All " & _
                " Select Distinct SO033.CitemCode From " & GetOwner & "SO033," & _
                  GetOwner & "CD019, " & GetOwner & "CD019A," & GetOwner & "CD024 " & _
                " WHERE SO033.CUSTID =" & rsSO004("CustId") & " AND SO033.FACISEQNO = '" & rsSO004("SeqNo") & "' " & _
                " AND SO033.SERVICETYPE = 'D'  AND SO033.CITEMCODE = CD019.CODENO AND NVL(CD019.STOPFLAG,0) = 0 " & _
                "  AND CD019.CODENO = CD019A.CITEMCODE AND CD019A.CODENO = CD024.CODENO " & _
                " And SO033.SBillNo Is NULL " & _
                " AND NVL(CD024.STOPFLAG,0) = 0  And SO033.CompCode= " & rsSO004("CompCode") & _
                IIf(IsCanRetWatch, " AND CD024.CANRETWATCH = 1 ", " ") & " AND Nvl(SO033.CHEVEN,0)=0 " & _
                IIf(Len(aBillNo) > 0, " And SO033.BillNo='" & aBillNo & "' ", "") & _
                IIf(Len(aFilterClctStopDateType) > 0, " And SO033.CitemCode " & aFilterClctStopDateType, "")
    '#6803 B15命令一律抓所有收費項目，如果派工流程有傳收費項目，則要加串 By Kin 2014/05/30
    If Len(OtherCitemCodes) > 0 Then
        aSQL = aSQL & " Union All " & _
            " Select CodeNo CitemCode From " & GetOwner & "CD019 " & _
            " Where CodeNo In (" & OtherCitemCodes & " )"
    End If
    aSQL = "Select Distinct CitemCode From (" & aSQL & ") "
    
    GetSmitChargeCitemCodeSQL = aSQL
        
  Exit Function
ChkErr:
  Call ErrSub(FormName, "GetSmitChargeCitemCode")
End Function


Public Function ReadGICMIS1(strKey As String) As String
  On Error GoTo ChkErr
    Dim strINI As String
    strINI = GaryGi(10)
    If strINI = "" Then
        ReadGICMIS1 = Empty
        Exit Function
    End If
    
    ReadGICMIS1 = ReadFROMINI(strINI, "GICMISPath", strKey)
  Exit Function
ChkErr:
    Call ErrSub(FormName, "Function ReadGICMIS1")
End Function
Public Function ReadFROMINI(IniFileName As String, Section As String, Key As String, Optional DecryptFlag As Boolean) As String
  On Error GoTo ChkErr
    If Not ChkDirExist(IniFileName) Then MsgBox "INI之檔名或路徑不正確!!", vbInformation, "訊息..": Exit Function
    Dim s As String, length As Long
    s = String(1024, 0)
    length = GetPrivateProfileString(Section, Key, "", s, Len(s), IniFileName)
    ReadFROMINI = IIf(DecryptFlag, Decrypt(Left(s, length)), Left(s, length))
  Exit Function
ChkErr:
    Call ErrSub(FormName, "Function ReadFROMINI")
End Function
'判斷是否要記錄LOG
Public Function IsRecordProcedure() As Boolean
  On Error GoTo ChkErr
    Dim aFile As String
    gErrLogPath = GaryGi(10)
    
    If gErrLogPath = Empty Then
        IsRecordProcedure = False
        Exit Function
        'gErrLogPath = ReadGICMIS1("ErrLogPath")
    End If
    
    aFile = gErrLogPath
    If Right(aFile, 1) <> "\" Then
        aFile = gErrLogPath & "\"
    End If
    'aFile = aFile & "VOD_Record\"
    aFile = aFile & strRecordName
    'aFile = aFile & GetGUID & "_" & Format(Now, "yyyymmddhhmmss") & ".TXT"
    If Dir(aFile) = "" Then
        IsRecordProcedure = False
        Exit Function
    End If
    Set objFileSystem = New Scripting.FileSystemObject
    strRecordProcedureFile = gErrLogPath
    
    If Right(strRecordProcedureFile, 1) <> "\" Then
        strRecordProcedureFile = strRecordProcedureFile & "\"
    End If
    'strRecordProcedureFile = strRecordProcedureFile & "VOD_Record\"
    If Not objFileSystem.FolderExists(strRecordProcedureFile) Then
        objFileSystem.CreateFolder (strRecordProcedureFile)
    End If
    strRecordProcedureFile = strRecordProcedureFile & "VOD_" & Format(Now, "yyyymmddhhmmss") & "_" & _
            GetGUID & ".TXT"
    
    Set objRecordWrite = objFileSystem.CreateTextFile(strRecordProcedureFile, True)
    IsRecordProcedure = True
    Exit Function
ChkErr:
  IsRecordProcedure = False
End Function
'將程式訊息寫入記錄檔
Public Sub WriteRecordVodProcedure(ByVal aMsg As String, Optional ByVal aBlnMoveDir As Boolean = False)
  On Error Resume Next
    If Not blnRecordProcedure Then Exit Sub
    Dim strMsg As String
    strMsg = Now & " > " & aMsg
    objRecordWrite.WriteLine strMsg
    If aBlnMoveDir Then
        
        objRecordWrite.Close
        Dim aMoveFolder As String
        Dim aMoveFileName As String
        aMoveFolder = gErrLogPath
        If Right(aMoveFolder, 1) <> "\" Then
            aMoveFolder = aMoveFolder & "\"
        End If
        aMoveFileName = objFileSystem.GetFileName(strRecordProcedureFile)
        If Not objFileSystem.FolderExists(aMoveFolder & strMoveDir) Then
            objFileSystem.CreateFolder aMoveFolder & strMoveDir
            
        End If
        objFileSystem.MoveFile strRecordProcedureFile, aMoveFolder & strMoveDir & "\" & aMoveFileName
    End If
    
End Sub

Public Function GetGUID() As String
  On Error Resume Next
    Dim udtGUID As GUID
    If (CoCreateGuid(udtGUID) = 0) Then
        GetGUID = _
        String(8 - Len(Hex$(udtGUID.Data1)), "0") & Hex$(udtGUID.Data1) & _
        String(4 - Len(Hex$(udtGUID.Data2)), "0") & Hex$(udtGUID.Data2) & _
        String(4 - Len(Hex$(udtGUID.Data3)), "0") & Hex$(udtGUID.Data3) & _
        IIf((udtGUID.Data4(0) < &H10), "0", "") & Hex$(udtGUID.Data4(0)) & _
        IIf((udtGUID.Data4(1) < &H10), "0", "") & Hex$(udtGUID.Data4(1)) & _
        IIf((udtGUID.Data4(2) < &H10), "0", "") & Hex$(udtGUID.Data4(2)) & _
        IIf((udtGUID.Data4(3) < &H10), "0", "") & Hex$(udtGUID.Data4(3)) & _
        IIf((udtGUID.Data4(4) < &H10), "0", "") & Hex$(udtGUID.Data4(4)) & _
        IIf((udtGUID.Data4(5) < &H10), "0", "") & Hex$(udtGUID.Data4(5)) & _
        IIf((udtGUID.Data4(6) < &H10), "0", "") & Hex$(udtGUID.Data4(6)) & _
        IIf((udtGUID.Data4(7) < &H10), "0", "") & Hex$(udtGUID.Data4(7))
    End If
End Function
Public Function GetPrbCtl() As Object
  On Error GoTo ChkErr
    Dim objCtl As Object
    If blnShowMsg Then
        If frmSO1623A Is Nothing Then
            Set objCtl = Nothing
        Else
            Set objCtl = frmSO1623A.pbr1
        End If
        
    Else
        Set objCtl = Nothing
    End If
    Set GetPrbCtl = objCtl
    Exit Function
ChkErr:
  Call ErrSub(FormName, "GetPrbCtl")
End Function


