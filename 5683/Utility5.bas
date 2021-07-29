Attribute VB_Name = "Utility5"
Public strChoose As String
Public strSQL As String
Public strChooseString As String
Public strViewName As String
Public strGroupName As String
Public cnn As New ADODB.Connection
Public strErrFile As String '錯誤欄位名稱
Public strSubViewName(10)

Enum giVType
    giLongV = 0
    giStringV = 1
    giDateV = 2
End Enum

'Public Enum ChangeDB
'    giAccessDb = 0 'Access Table
'    giOracle = 1    'Oracle Table
'End Enum
Public Function GetPaynowFlag() As Boolean
  On Error Resume Next
     GetPaynowFlag = Val(GetRsValue("SELECT PaynowFlag FROM " & GetOwner & "SO041") & "") = 1
End Function
Public Function subSpace(str As String) As String
  On Error GoTo chkErr
   If str = "" Then
       subSpace = ""
   Else
       subSpace = str
   End If
  Exit Function
chkErr:
    Call ErrSub("Utility5", "subSpace")
End Function
'*******************************************************************************
'SO5610~SO5630使用   2007-10-16 BY TOMMY
'當該工單有對應到SO004且SO004有對應到SO004D時，若SO004.DECLARANTNAME有值，
'則工單的CUSTNAME取SO004.DECLARANTNAME，工單的TEL1取SO004.CONTTEL，
'作為報表顯示欄位的資料來源，反之則取原工單的資料
'strTable為SO007~SO009
'*******************************************************************************
Public Function fnGetDecode(strTable As String)
    Dim strDecode  As String
    On Error GoTo chkErr
    
    strDecode = "DECODE(DECODE(SO004D.SNO,NULL," & strTable & ".CUSTNAME,SO004.DECLARANTNAME),NULL," & strTable & ".CUSTNAME,DECODE(SO004D.SNO,NULL," & strTable & ".CUSTNAME,SO004.DECLARANTNAME)) CUSTNAME,"
    strDecode = strDecode & _
            "DECODE(DECODE(SO004D.SNO,NULL," & strTable & ".TEL1,SO004.CONTTEL),NULL," & strTable & ".TEL1,DECODE(SO004D.SNO,NULL," & strTable & ".TEL1,SO004.CONTTEL)) TEL1,"
    fnGetDecode = strDecode
    Exit Function
chkErr:
    Call ErrSub("Utility5", "fnGetDecode")
End Function

Public Sub SetgiList(objGilist As Object, strFldName1 As String, strFldName2 As String, _
            StrTableName As String, Optional lngTop As Long, Optional lngLeft As Long, Optional lngWidth1 As Long, _
            Optional lngWidth2 As Long, Optional lngFldLen1 As Long, Optional lngFldLen2 As Long, _
            Optional strWhere As String, Optional blnStopFlag As Boolean = False, Optional cn As ADODB.Connection)
 '欄位1,欄位2,表格名稱,Top,Left,Width1,Width2,FldLen1,FldLen2
  On Error GoTo chkErr
    If UCase(StrTableName) <> "CD039" Then objGilist.Clear
    objGilist.SetFldName1 (strFldName1)
    objGilist.SetFldName2 (strFldName2)
    'objGilist.SetTableName (StrTableName)
    objGilist.SetTableName GetOwner & StrTableName
    objGilist.Filter = strWhere
    If lngTop > 0 Then objGilist.Top = lngTop
    If lngLeft > 0 Then objGilist.Left = lngLeft
    If lngWidth1 > 0 Then objGilist.FldWidth1 = lngWidth1
    If lngWidth2 > 0 Then objGilist.FldWidth2 = lngWidth2
    If lngFldLen1 > 0 Then objGilist.FldLen1 = lngFldLen1
    If lngFldLen2 > 0 Then objGilist.FldLen2 = lngFldLen2
    objGilist.FilterStop = blnStopFlag
    If cn Is Nothing Then Set cn = gcnGi
    Call objGilist.SendConn(cn)
  Exit Sub
chkErr:
    Call ErrSub("Utility5", "SetgiList")
End Sub
'Public Sub SetgiList(objGilist As Object, strFldName1 As String, strFldName2 As String, _
'            StrTableName As String, Optional lngTop As Long, Optional lngLeft As Long, Optional lngWidth1 As Long, _
'            Optional lngWidth2 As Long, Optional lngFldLen1 As Long, Optional lngFldLen2 As Long, _
'            Optional strWhere As String, Optional blnStopFlag As Boolean = False)
' '欄位1,欄位2,表格名稱,Top,Left,Width1,Width2,FldLen1,FldLen2
'  On Error GoTo ChkErr
'    If UCase(StrTableName) <> "CD039" Then objGilist.Clear
'    objGilist.SetFldName1 (strFldName1)
'    objGilist.SetFldName2 (strFldName2)
'    'objGilist.SetTableName (StrTableName)
'    objGilist.SetTableName GetOwner & StrTableName
'    objGilist.Filter = strWhere
'    If lngTop > 0 Then objGilist.Top = lngTop
'    If lngLeft > 0 Then objGilist.Left = lngLeft
'    If lngWidth1 > 0 Then objGilist.FldWidth1 = lngWidth1
'    If lngWidth2 > 0 Then objGilist.FldWidth2 = lngWidth2
'    If lngFldLen1 > 0 Then objGilist.FldLen1 = lngFldLen1
'    If lngFldLen2 > 0 Then objGilist.FldLen2 = lngFldLen2
'    objGilist.FilterStop = blnStopFlag
'    Call objGilist.SendConn(gcnGi)
'  Exit Sub
'ChkErr:
'    Call ErrSub("Utility5", "SetgiList")
'End Sub

Public Function SelectServType(ByVal strNo As String, ByVal objServiceType As Object) As String
  On Error Resume Next
    Dim rsSO041 As New ADODB.Recordset
    Dim strType As String
    Dim varType() As String
    Dim intLoop As Integer
    Dim strType1 As String
    If strNo = "" Then Exit Function
    If Not GetRS(rsSO041, "Select ServiceType,MainServiceType From " & GetOwner & "SO041 Where CompCode =" & strNo, gcnGi, adUseClient, adOpenDynamic, adLockOptimistic) Then Exit Function
    If rsSO041.RecordCount > 0 Then
        strType = Replace(rsSO041("ServiceType"), Chr(13), "")
          If Err.Number = 3021 Then
              strType = "Where CodeNo=''"
          Else
            On Error GoTo chkErr
              strType = Replace(strType, ",", "")
              For intLoop = 1 To Len(strType)
                strType1 = strType1 & ",'" & Mid(strType, intLoop, 1) & "'"
              Next
              strType = Mid(strType1, 2)
              strType = "Where CodeNo in (" & strType & ")"
          End If
        If TypeName(objServiceType) = "GiList" Then
            Call SetgiList(objServiceType, "CodeNo", "Description", "CD046", , , , , , , strType)
            objServiceType.SetCodeNo ""
            objServiceType.Query_Description
            objServiceType.ListIndex = 1
        Else
            Call SetgiMulti(objServiceType, "CodeNo", "Description", "CD046", "服務類別代碼", "服務類別名稱")
            objServiceType.SetQueryCode rsSO041("MainServiceType") & ""
        End If
        
    End If
    CloseRecordset rsSO041
  Exit Function
chkErr:
  Call ErrSub("Utility5", "SelectServType")
End Function

'Public Function GetServiceType(ByVal intCompCode As Integer) As String
'  On Error Resume Next
'    Dim strType As String
'    Dim varType() As String
'    Dim intLoop As Integer
'    Dim strType1 As String
'    If strNo = "" Then Exit Function
'    strType = Replace(gcnGi.Execute("Select ServiceType From SO041 Where CompCode =" & intCompCode).GetString, Chr(13), "")
'      If Err.Number = 3021 Then
'          GetServiceType = ""
'      Else
'        On Error GoTo ChkErr
'          strType = Replace(strType, "'", "")
'          varType = Split(strType, ",")
'          GetServiceType = varType(0)
'      End If
'  Exit Function
'ChkErr:
'  Call ErrSub("Utility5", "GetServiceType")
'End Function
Public Sub SetgiMulti(objgiMulti As Object, strFldName1 As String, _
        strFldName2 As String, StrTableName As String, strFldCaption1 As String, _
        strFldCaption2 As String, Optional strFilter As String, _
        Optional blnStopFlag As Boolean = False, Optional cn As ADODB.Connection)
 '欄位1,欄位2,Table,中文1,中文2
  On Error GoTo chkErr
    If cn Is Nothing Then Set cn = gcnGi
    Call objgiMulti.SendConn(cn)
    objgiMulti.FldName1 = strFldName1
    objgiMulti.FldName2 = strFldName2
    objgiMulti.TableName = GetOwner & StrTableName
    objgiMulti.FldCaption1 = strFldCaption1
    objgiMulti.FldCaption2 = strFldCaption2
    objgiMulti.FilterStop = blnStopFlag
    If strFilter <> "" Then objgiMulti.Filter = strFilter
  Exit Sub
chkErr:
    Call ErrSub("Utility5", "SetgiMulti")
End Sub

'Public Sub SetgiMulti(objgiMulti As Object, strFldName1 As String, _
'        strFldName2 As String, StrTableName As String, strFldCaption1 As String, _
'        strFldCaption2 As String, Optional strFilter As String, _
'        Optional blnStopFlag As Boolean = False)
' '欄位1,欄位2,Table,中文1,中文2
'  On Error GoTo ChkErr
'    Call objgiMulti.SendConn(gcnGi)
'    objgiMulti.FldName1 = strFldName1
'    objgiMulti.FldName2 = strFldName2
'    objgiMulti.TableName = GetOwner & StrTableName
'    objgiMulti.FldCaption1 = strFldCaption1
'    objgiMulti.FldCaption2 = strFldCaption2
'    objgiMulti.FilterStop = blnStopFlag
'    If strFilter <> "" Then objgiMulti.Filter = strFilter
'  Exit Sub
'ChkErr:
'    Call ErrSub("Utility5", "SetgiMulti")
'End Sub

Public Sub SetgiMultiAddItem(objgiMulti As GiMulti, strCodeNo As String, strDescription As String, Optional strFldCaption1 As String, Optional strFldCaption2 As String)
  On Error GoTo chkErr
   Dim varCodeNo As Variant
   Dim varDescription As Variant
   Dim varValue As Variant
   Dim intValue As Integer
   objgiMulti.Reset
   objgiMulti.FldCaption1 = strFldCaption1
   objgiMulti.FldCaption2 = strFldCaption2
   objgiMulti.FldName1 = "CodeNo"
   objgiMulti.FldName2 = "Description"
   objgiMulti.DIY = True
   intValue = 0
   varCodeNo = Split(strCodeNo, ",")
   varDescription = Split(strDescription, ",")
   For Each varValue In varCodeNo
       objgiMulti.AddItem varValue, varDescription(intValue)
       intValue = intValue + 1
   Next
  Exit Sub
chkErr:
    Call ErrSub("Utility5", "SetgiMultiAddItem")
End Sub

Public Sub subAnd(strCH As String)
  On Error GoTo chkErr
    If strChoose <> "" Then
        strChoose = strChoose & " And " & strCH
    Else
        strChoose = strCH
    End If
  Exit Sub
chkErr:
    Call ErrSub("Utility5", "subAnd")
End Sub

Public Function subAnd2(strChoose As String, strAnd As String) As String
  On Error GoTo chkErr
    If strChoose <> "" Then
        strChoose = strChoose & " And " & strAnd
    Else
        strChoose = strAnd
    End If
    subAnd2 = strChoose
  Exit Function
chkErr:
    Call ErrSub("Utility5", "subAnd2")
End Function

Public Function GetNullString(varValue As Variant, _
                     Optional lngVType As giVType = giStringV, _
                     Optional DBFlag As ChangeDB = giOracle, _
                     Optional blnAddTime As Boolean = False) As Variant
    On Error GoTo chkErr
        If lngVType = giDateV Then
            If Len(varValue & "") = 0 Then
                GetNullString = "NULL"
            Else
                If DBFlag = giOracle Then
                    If blnAddTime Then
                       GetNullString = "To_Date('" & Format(varValue, "YYYYMMDDHHMMSS") & "','YYYYMMDDHH24MISS')"
                    Else
                       GetNullString = "To_Date('" & Format(varValue, "YYYYMMDD") & "','YYYYMMDD')"
                    End If
                Else
                    If blnAddTime Then
                        GetNullString = "#" & Format(varValue, "YYYY/MM/DD HH:MM:SS") & "#"
                    Else
                        GetNullString = "#" & Format(varValue, "YYYY/MM/DD") & "#"
                    End If
                End If
            End If
        ElseIf lngVType = giLongV Then
            If Len(varValue & "") = 0 Then
                GetNullString = "NULL"
            Else
                GetNullString = varValue
            End If
        ElseIf lngVType = giStringV Then
            If Len(varValue & "") = 0 Then
                GetNullString = "NULL"
            Else
                varValue = Replace(varValue, "'", "''")
                GetNullString = "'" & varValue & "'"
            End If
        End If
    Exit Function
chkErr:
  Call ErrSub("Utility5", "GetNullString")
End Function

Public Function ChkDate2(objDate1 As Object, objDate2 As Object) As Boolean
  On Error GoTo chkErr
    If objDate1.GetValue <> "" And objDate2.GetValue < objDate1.GetValue Then
       MsgDateRangeX
       objDate2.SetValue GetDayLastTime(objDate1.GetValue(True))
       ChkDate2 = True
    End If
  Exit Function
chkErr:
  Call ErrSub("Utility5", "chkDate2")
End Function

Public Function GetLastDate(ByVal YYYYMM As String) As String '取當月最後一天
  On Error GoTo chkErr
    If Not IsDate(YYYYMM) Then Exit Function
    If InStr(1, YYYYMM, "/") > 0 Then
        GetLastDate = (DateAdd("M", 1, Format(YYYYMM, "YYYY/MM") & "/01") - 1)
    Else
        GetLastDate = (DateAdd("M", 1, Format(YYYYMM, "####/##") & "/01") - 1)
    End If
  Exit Function
chkErr:
  Call ErrSub("Utility5", "GetLastDate")
End Function

Public Sub FunctionKey(KeyCode As Integer, Shift As Integer, strFormName As Form)
  On Error GoTo chkErr
    Select Case KeyCode
           '----------------------------------------------------
           Case vbKeyF5 '   F5: 列印 相當於按下cmdPrint
                        If Not ChkGiList(KeyCode) Then Exit Sub
                        If strFormName.cmdPrint.Enabled = True Then strFormName.cmdPrint.Value = True
           Case vbKeyEscape '   Esc
                        If strFormName.cmdExit.Enabled = True Then strFormName.cmdExit.Value = True
           '----------------------------------------------------
    End Select
  Exit Sub
chkErr:
  Call ErrSub("Utility5", "FunctionKey")
End Sub

Public Function GetCode(ByVal strTable As String, _
                        Optional ByVal intRefNo As String, _
                        Optional ByVal strServiceType As String = "", _
                        Optional ByVal strCompCode As String = "", _
                        Optional ByVal strType As giVType = giLongV, _
                        Optional ByVal blnAdd As Boolean = True, _
                        Optional ByVal strWhere As String = "") As String
  On Error GoTo chkErr
    Dim rsCode As New ADODB.Recordset
    Dim strSC As String
    Dim strOne As String
    Dim strMulti1 As String
    Dim strMulti2 As String
    If rsCode.State = adStateOpen Then rsCode.Close
    rsCode.CursorLocation = adUseClient
    If strWhere = "" Then
        strSC = " Where RefNo " & intRefNo
    Else
        strSC = " Where " & strWhere
    End If
    If strServiceType <> "" Then strSC = strSC & " And Nvl(ServiceType,'" & Trim(strServiceType) & "') = '" & Trim(strServiceType) & "'"
    If strCompCode <> "" Then strSC = strSC & " And Nvl(CompCode," & Trim(strCompCode) & ") = " & Trim(strCompCode)
    rsCode.Open "Select CodeNo From " & GetOwner & strTable & strSC, gcnGi, adOpenDynamic, adLockReadOnly
    If rsCode.EOF Then
        GetCode = ""
    Else
        If blnAdd = True Then
            strOne = "="
            strMulti1 = " IN("
            strMulti2 = ")"
        Else
            strOne = ""
            strMulti1 = ""
            strMulti2 = ""
        End If
        If rsCode.RecordCount = 1 Then
            If strType = giStringV Then
                GetCode = strOne & "'" & rsCode("CodeNo") & "'"
            Else
                GetCode = strOne & rsCode("CodeNo")
            End If
        Else
            If strType = giStringV Then
                GetCode = "'" & rsCode.GetString(, , , "','") & "'"
                GetCode = strMulti1 & "'" & Left(GetCode, Len(GetCode) - 1) & "'" & strMulti2
            Else
                GetCode = rsCode.GetString(, , , ",")
                GetCode = strMulti1 & Left(GetCode, Len(GetCode) - 1) & strMulti2
            End If
        End If
    End If
    rsCode.Close
  Exit Function
chkErr:
  Call ErrSub("Utility5", "GetCode")
End Function

Public Function ReturnAndString(ByRef str1 As String, strAnd As String, strDivide As String) As Boolean
  On Error GoTo chkErr
    str1 = Trim(str1)
    If str1 = "" Then
        str1 = strAnd
    Else
        str1 = str1 & strDivide & strAnd
    End If
    ReturnAndString = True
  Exit Function
chkErr:
  Call ErrSub("Utility5", "ReturnAndString")
End Function

Public Function CreateSubView2(strsubsql() As String) As Boolean
  On Error Resume Next
    Dim intFor As Integer
    Dim intLoop As Integer
      For intFor = 0 To UBound(strsubsql)
          gcnGi.Execute "Drop View " & strSubViewName(intFor)
      Next
  On Error GoTo chkErr
      For intFor = 0 To UBound(strsubsql)
          If strsubsql(intFor) <> "" Then If Not ExeView(intFor, strsubsql(intFor)) Then Exit Function
          SendSQL strsubsql(intFor)
      Next
    CreateSubView2 = True
  Exit Function
chkErr:
  Call ErrSub("Utility5", "CreatSubView2")
End Function

'Public Function CreateSubView(ByVal strSubSQL1 As String, Optional strSubSQL2 As String = "", _
'                              Optional strSubSQL3 As String = "", Optional strSubSQL4 As String = "", _
'                              Optional strSubSQL5 As String = "", Optional strSubSQL6 As String = "", _
'                              Optional strSubSQL7 As String = "", Optional strSubSQL8 As String = "", _
'                              Optional strSubSQL9 As String = "", Optional strSubSQL10 As String = "") As Boolean
'  On Error Resume Next
'    Dim intFor As Integer
'    Dim intLoop As Integer
'      For intFor = 1 To 10
'          gcnGi.Execute "Drop View " & strSubViewName(intFor)
'      Next
'  On Error GoTo ChkErr
'    If Not ExeView(1, strSubSQL1) Then Exit Function
'    If strSubSQL2 <> "" Then If Not ExeView(2, strSubSQL2) Then Exit Function
'    If strSubSQL3 <> "" Then If Not ExeView(3, strSubSQL3) Then Exit Function
'    If strSubSQL4 <> "" Then If Not ExeView(4, strSubSQL4) Then Exit Function
'    If strSubSQL5 <> "" Then If Not ExeView(5, strSubSQL5) Then Exit Function
'    If strSubSQL6 <> "" Then If Not ExeView(6, strSubSQL6) Then Exit Function
'    If strSubSQL7 <> "" Then If Not ExeView(7, strSubSQL7) Then Exit Function
'    If strSubSQL8 <> "" Then If Not ExeView(8, strSubSQL8) Then Exit Function
'    If strSubSQL9 <> "" Then If Not ExeView(9, strSubSQL9) Then Exit Function
'    If strSubSQL10 <> "" Then If Not ExeView(10, strSubSQL10) Then Exit Function
'
'    CreateSubView = True
'  Exit Function
'ChkErr:
'  Call ErrSub("Utility5", "CreatSubView")
'End Function

Private Function ExeView(ByVal intIndex As Integer, ByVal SubSQLName As String) As Boolean
  On Error Resume Next
    strSubViewName(intIndex) = GetTmpViewName
    gcnGi.Execute "Drop View " & strSubViewName(intIndex)
  On Error GoTo chkErr
    gcnGi.Execute "Create View " & strSubViewName(intIndex) & " as (" & SubSQLName & ")"
    ExeView = True
  Exit Function
chkErr:
  If Err.Number = -2147217900 Then
    If intLoop < 3 Then
      intLoop = intLoop + 1
      gcnGi.Execute "Drop View " & strSubViewName(intIndex)
      Resume 0
    Else
      MsgBox "此一 View (" & strSubViewName(intIndex) & ") 已被一個物件使用中,無法刪除成功!!"
      Exit Function
    End If
  Else
      Call ErrSub("Utility5", "ExeView")
  End If
End Function

Public Function GetSubSQL(ByVal FieldName1 As String, ByVal FieldName2 As String, ByVal strSQL As String, ByVal TableName As String, Optional ByVal blnCust As Boolean = True, Optional ByVal SumCust As Long = 0) As String
  On Error GoTo chkErr
  'If IsMissing(SumCust) Then SumCust = Null
  If blnCust Then
      'GetSubSQL = "SELECT " & FieldName1 & " as FileCode," & FieldName2 & " as FieldName," & _
                  "Count(Distinct " & TableName & ".SNo) CountSNo," & _
                  "Count(Distinct " & TableName & ".Custid) CountCustid," & _
                  "(Count(Distinct " & TableName & ".SNo)-Count(Distinct DeCode(Nvl(length(SubStr(" & TableName & ".ReturnCode,1,1)),0),1," & TableName & ".SNo,Null))) as CountAccept," & _
                  "Count(Distinct DeCode(Nvl(length(SubStr(" & TableName & ".ReturnCode,1,1)),0),1," & TableName & ".SNo,Null)) as CountReturn," & _
                  SumCust & " as SumCustid " & strSql & _
                  " Group BY " & FieldName1 & "," & FieldName2
      GetSubSQL = "SELECT " & FieldName1 & " as FileCode," & FieldName2 & " as FieldName," & _
                  "Count(Distinct " & TableName & ".SNo) CountSNo," & _
                  "Count(Distinct " & TableName & ".Custid) CountCustid," & _
                  "(Count(Distinct " & TableName & ".SNo)-Count(Distinct DeCode(" & TableName & ".ReturnCode,Null,Null," & TableName & ".SNo))) as CountAccept," & _
                  "Count(Distinct DeCode(" & TableName & ".ReturnCode,Null,Null," & TableName & ".SNo)) as CountReturn," & _
                  SumCust & " as SumCustid " & strSQL & _
                  " Group BY " & FieldName1 & "," & FieldName2
  Else
      GetSubSQL = "SELECT " & FieldName1 & " as FileCode," & FieldName2 & " as FieldName," & _
                "Count(Distinct " & TableName & ".SNo) CountSNo," & _
                "Count(Distinct " & TableName & ".Custid) CountCustid," & _
                SumCust & " as SumCustid " & strSQL & _
                " Group BY " & FieldName1 & "," & FieldName2
  End If
  Exit Function
chkErr:
  Call ErrSub("Utility5", "GetSubSQL")
End Function

Public Function ChkMaxLengthOK(objText As Object) As Boolean
  On Error GoTo chkErr
    Dim strChar As String
    Dim strChar2 As String
    Dim intLength  As Integer
    Dim intLength2 As Integer
        ChkMaxLengthOK = True
        strChar = objText.Text
        Do
            intLength = InStr(1, objText.Text, ",", vbTextCompare)
            strChar = Mid(strChar, intLength + 1)
            If InStr(1, strChar, ",", vbTextCompare) = 0 Then
                If Len(Trim(strChar)) >= 8 Then
                  strChar2 = strChar
                  Do
                    intLength2 = InStr(1, strChar2, "-", vbTextCompare)
                    strChar2 = Mid(strChar2, intLength2 + 1)
                    If InStr(1, strChar2, "-", vbTextCompare) = 0 Then
                      If Len(Trim(strChar2)) >= 8 Then ChkMaxLengthOK = False
                      Exit Do
                    End If
                  Loop
                End If
               Exit Do
            End If
        Loop
    Exit Function
chkErr:
  Call ErrSub("Utility5", "ChkMaxLengthOk")
End Function

Public Function TimetxtCustId(objText As Object, Optional strTable As String, Optional strField As String) As Boolean
  On Error GoTo chkErr
    Dim arrCustid()
    Dim strCustid As String
      TimetxtCustId = False
        Call ParseWord(objText.Text, arrCustid)
        strCustid = Join(arrCustid, ",")
        If strCustid = "" Then Exit Function
        If strTable = "" Then strTable = "A"
        If strField = "" Then strField = "CustId"
        If InStr(1, strCustid, ",") = 0 Then
            Call subAnd(strTable & "." & strField & " =" & strCustid)
        Else
            Call subAnd(strTable & "." & strField & " In (" & strCustid & ")")
        End If
      TimetxtCustId = True
  Exit Function
chkErr:
  Call ErrSub("Utility5", "TimetxtCustId")
End Function

Public Function GetDayLastTime(strDate As Variant) As String
  On Error Resume Next
    'GetDayLastTime = DateAdd("n", "-1", CDate(strDate) + 1)
    GetDayLastTime = Format(DateAdd("n", "-1", CDate(strDate) + 1), "YYYY/MM/DD HH:MM:SS")
End Function

Public Function GetUseIndexStr(StrTableName As String, strColumnName As String) As String
  On Error Resume Next
    '先拿掉指定索引
    'GetUseIndexStr = UCase(" /*+ Index(" & GetOwner & StrTableName & " I_" & StrTableName & "_" & strColumnName & ") */ ")
    GetUseIndexStr = ""
End Function

Public Function CreateTmpIndex(StrTableName As String, strColumnName As String, Optional strColumn As String = "") As Boolean
  On Error Resume Next
    If strColumn = "" Then strColumn = strColumnName
    gcnGi.Execute "Drop Index " & GetOwner & "I_" & StrTableName & "_" & strColumnName
    gcnGi.Execute "Create Index " & GetOwner & "I_" & StrTableName & "_" & strColumnName & " On " & GetOwner & StrTableName & " (" & strColumn & ")"
    CreateTmpIndex = True
End Function

Public Function DropTmpIndex(StrTableName As String, strColumnName As String) As Boolean
  On Error Resume Next
    gcnGi.Execute "Drop Index " & GetOwner & "I_" & StrTableName & "_" & strColumnName
    DropTmpIndex = True
End Function

Public Function GetTmpMDBExecuteStr(rs As ADODB.Recordset, StrTableName As String) As String
  On Error GoTo chkErr
    Dim lngLoop As Long
    Dim strField As String
    Dim strValue As String
    Dim strVal
        For lngLoop = 0 To rs.Fields.Count - 1
            strField = strField & ",[" & rs.Fields(lngLoop).Name & "]"
            If rs.Fields(lngLoop).Type = adDBTimeStamp Then
                strVal = GetNullString(rs.Fields(lngLoop).Value, giDateV, giAccessDb, True)
            Else
                strVal = GetNullString(rs.Fields(lngLoop).Value)
            End If
            strValue = strValue & "," & strVal
        Next
        strField = Mid(strField, 2)
        strValue = Mid(strValue, 2)
        GetTmpMDBExecuteStr = "Insert Into " & StrTableName & " ( " & strField & ") Values (" & strValue & ")"
    Exit Function
chkErr:
  Call ErrSub("Utility5", "GetTmpMDBExecuteStr")
End Function

Public Function CreateMDBTable(rs As ADODB.Recordset, StrTableName As String, cnn As ADODB.Connection) As Boolean
  On Error Resume Next
    Dim lngLoop As Long
    Dim strField As String
    Dim strValue As String
    Dim strMessage As String
      cnn.Execute "Drop Table " & StrTableName
        If Err.Number = -2147217900 Then MsgBox "無法刪除〔" & StrTableName & "] 資料表 ", vbCritical, "錯誤..": Exit Function
        'If ChkTmpMdbTableExist(StrTableName) = True Then MsgBox "無法刪除〔" & StrTableName & "] 資料表 ", vbCritical, "錯誤..": Exit Function
        On Error GoTo chkErr
            CreateMDBTable = False
            If rs.State = 2 Then Exit Function
            For lngLoop = 0 To rs.Fields.Count - 1
                strValue = strValue & ",[" & rs.Fields(lngLoop).Name & "]"
                Select Case rs.Fields(lngLoop).Type
                       Case adBigInt, adVarNumeric, adNumeric, adDecimal, adSingle, adDouble, adInteger '"Double"
                            strValue = strValue & " Double"
                       Case adBoolean '"Logical"
                            strValue = strValue & " Logical"
                       Case adBSTR, adChar, adChapter, adVarChar, adVariant, adVarWChar, adLongVarChar '"Text("
                            If rs.Fields(lngLoop).DefinedSize > 512 Then
                                strValue = strValue & " Memo"
                            ElseIf rs.Fields(lngLoop).DefinedSize * 2 > 255 Then
                                strValue = strValue & " Memo"
                            Else
                                strValue = strValue & " Text(" & rs.Fields(lngLoop).DefinedSize * 2 & ")"
                            End If
                       Case adDate, adDBDate, adDBTimeStamp, adDBTime '"Date"
                            strValue = strValue & " Date"
                       Case Else
                            strMessage = rs.Fields(lngLoop).Type
                End Select
            Next
        strValue = Mid(strValue, 2)
      cnn.Execute "Create Table " & StrTableName & " (" & strValue & ")"
      CreateMDBTable = True
  Exit Function
chkErr:
  Call ErrSub("Utility5", "CreateMDBTable : Type : (" & strMessage & "), SQL:(" & strValue & ")")
End Function

Public Sub DropTMPVIEW(strViewName As String, strSubViewName() As Variant)
  On Error Resume Next
    Dim lngLoop As Long
      gcnGi.Execute "Drop View " & GetOwner & strViewName
      For lngLoop = 1 To UBound(strSubViewName)
          If Len(strSubViewName(lngLoop) & "") > 0 Then gcnGi.Execute "Drop View " & GetOwner & strSubViewName(lngLoop)
      Next
End Sub

Public Function SplitWipCode(ByVal strWipCode As String, ByVal StrTableName As String) As String
  On Error GoTo chkErr
    Dim arrWipCode() As String
    Dim intFor As Integer
    Dim strWipCode1 As String
    Dim strWipCode2 As String
    Dim strWipCode3 As String
    
    strWipCode1 = ""
    strWipCode2 = ""
    strWipCode3 = ""
    
    arrWipCode = Split(Replace(strWipCode, "'", ""), ",")
    For intFor = 0 To UBound(arrWipCode)
      If Len(arrWipCode(intFor)) = 1 Then
          strWipCode1 = strWipCode1 & "," & arrWipCode(intFor)
      Else
          If Left(arrWipCode(intFor), 1) = 2 Then
              strWipCode2 = strWipCode2 & "," & arrWipCode(intFor)
          Else
              strWipCode3 = strWipCode3 & "," & arrWipCode(intFor)
          End If
      End If
    Next
    If Len(strWipCode1) > 0 Then
        If InStr(2, strWipCode1, ",") = 0 Then
            SplitWipCode = StrTableName & ".WipCode1 =" & Mid(strWipCode1, 2)
        Else
            SplitWipCode = StrTableName & ".WipCode1 in (" & Mid(strWipCode1, 2) & ")"
        End If
    End If
    If Len(strWipCode2) > 0 Then
        If InStr(2, strWipCode2, ",") = 0 Then
            SplitWipCode = IIf(SplitWipCode = "", "", SplitWipCode & " Or ") & StrTableName & ".WipCode2 =" & Mid(strWipCode2, 2)
        Else
            SplitWipCode = IIf(SplitWipCode = "", "", SplitWipCode & " Or ") & StrTableName & ".WipCode2 in (" & Mid(strWipCode2, 2) & ")"
        End If
    End If
    If Len(strWipCode3) > 0 Then
        If InStr(2, strWipCode3, ",") = 0 Then
            SplitWipCode = IIf(SplitWipCode = "", "", SplitWipCode & " Or ") & StrTableName & ".WipCode3 =" & Mid(strWipCode3, 2)
        Else
            SplitWipCode = IIf(SplitWipCode = "", "", SplitWipCode & " Or ") & StrTableName & ".WipCode3 in (" & Mid(strWipCode3, 2) & ")"
        End If
    End If
  Exit Function
chkErr:
  Call ErrSub("Utility5", "SplitWipCode")
End Function

Public Sub CloseRecordset(rs As ADODB.Recordset)
  On Error Resume Next
    rs.CancelUpdate
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
End Sub

Public Sub ChangIntroVisible(ByVal funFrom As Object, ByVal funMediaCode As String, ByVal funCompCode As String)
  On Error GoTo chkErr
    Dim rsCD009 As New ADODB.Recordset
    Dim strSQL As String
    strSQL = "Select RefNo From " & GetOwner & "CD009 Where CodeNo" & funMediaCode
    Call GetRS(rsCD009, strSQL)
    Select Case rsCD009("refno")
           Case 1
                funFrom.lblCustid.Visible = True
                funFrom.txtCustId.Visible = True
                funFrom.gimIntroId.Visible = False
           Case 2
                Call SetgiMulti(funFrom.gimIntroId, "EmpNo", "EmpName", "CM003", "介紹人代碼", "介紹人姓名", "Where CompCode=" & funCompCode)
                funFrom.lblCustid.Visible = False
                funFrom.txtCustId.Visible = False
                funFrom.gimIntroId.Visible = True
           Case 3
                Call SetgiMulti(funFrom.gimIntroId, "CodeNo", "Description", "CD010", "介紹人代碼", "介紹人姓名", "Where CompCode=" & funCompCode)
                funFrom.lblCustid.Visible = False
                funFrom.txtCustId.Visible = False
                funFrom.gimIntroId.Visible = True
           Case Else
                funFrom.lblCustid.Visible = False
                funFrom.txtCustId.Visible = False
                funFrom.gimIntroId.Visible = False
    End Select
    CloseRecordset rsCD009
  Exit Sub
chkErr:
  Call ErrSub("Utility5", "ChangIntroVisible")
End Sub

Public Function AddMailType(ByVal objcboMailType As ComboBox, ByVal blnException As Boolean, Optional ByVal rs As ADODB.Recordset) As Boolean
  On Error GoTo chkErr
    Dim strSQL As String
        strSQL = "Select * From So052"
        If Not blnException Then strSQL = strSQL & " Where LabelRptName is not null"
        If GetRS(rs, strSQL, gcnGi, adUseClient, adOpenForwardOnly, adLockReadOnly) Then
           While Not rs.EOF
              objcboMailType.AddItem rs("DocName").Value
              rs.MoveNext
           Wend
        End If
        AddMailType = True
   Exit Function
chkErr:
  Call ErrSub("Utility5", "AddMailType")
End Function

'第一個傳條件說明,第二個傳條件內容,第三個傳是否加分號(; 93/11/24
'Ex: GetSelectConditionStr("實收日期","93/11/24~93/11/31",False)
'結果為: "實收日期:93/11/24~93/11/31;"
Public Function GetSelectConditionStr(ByVal strCaption As String, ByVal strSELECT As String, _
Optional ByVal blnSemicolon As Boolean = True) As String
Dim strSelectConditionStr As String
If strSELECT <> "" Then
strSelectConditionStr = strCaption & ":" & strSELECT
If blnSemicolon Then strSelectConditionStr = strSelectConditionStr & ";"
End If
GetSelectConditionStr = strSelectConditionStr
End Function

