Attribute VB_Name = "clsModule"
Option Explicit
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Const FormName = "clsModule"
'自訂列舉值
Public Const gimsgTranErr = "的資料時發生錯誤: 請查明原因!!"
Public Const vmsgSendConn = "請先傳入連線!!"
Public Const gimsgWarning = "警告!"
Public Const gimsgPrompt = "提示"
Public Const gimsgNoRcd = "查無資料"
Public Const gimsgPleaseSendSNo = "請先傳入單據編號"
Public Const gimsgPleaseSetSO011 = "請先設定預約時段[SO011]"
Public Const gMsgIsDataOK = "此為必要欄位,須有值 !! "
Public strErrPath As Variant
Public gcnGi As New ADODB.Connection
Public cnn As New ADODB.Connection
Public clsStoreParameter As Object
Public gErrLogPath As String
Public garyGi(20) As String
Public fso As New FileSystemObject
Public file(2) As TextStream
'Public objStorePara As clsStoreParameter

'Public Function msgResult(lngTotalCount As Long, lngErrCount As Long, lngTime As Long)
'    MsgBox "已完成資料筆數共" & lngTotalCount & "筆," & vbCrLf & vbCrLf & _
'           "問題筆數共" & lngErrCount & "筆," & vbCrLf & vbCrLf & _
'           "共花費:" & (Timer - lngTime) \ 1 & "秒"
'    objStorePara.uTime = (Timer - lngTime) \ 1
'End Function

''讀取 GICMIS1.INI 檔路徑
'Public Function GetIniPath1() As String
'  On Error GoTo ChkErr
'    Dim strINI As String
'    strINI = IIf(Right(Environ("WinDir"), 1) <> "\", Environ("WinDir") & "\", Environ("WinDir")) & "GICMIS1.INI"
'    If Not ChkDirExist(strINI) Then
'        'strINI = App.Path & "\GICMIS1.INI"
'        strINI = objStorePara.iniPath
'    End If
'    GetIniPath1 = strINI
'  Exit Function
'ChkErr:
'    Call ErrSub("DealFun", "Function GetIniPath1")
'End Function

Public Function PutGlobal()
  On Error GoTo ChkErr
    PutGlobal = Join(garyGi, Chr(9))
  Exit Function
ChkErr:
    Call ErrSub(FormName, "Function PutGlobal")
End Function

Public Function GetGlobal(strArray As String)
    On Error GoTo ChkErr
    Dim strParams() As String
    Dim varElement  As Variant
    Dim intLoop As Integer
        strParams = Split(strArray, Chr(9))
        ReDim Preserve strParams(20) As String
        For Each varElement In strParams
           garyGi(intLoop) = varElement
           intLoop = intLoop + 1
        Next
    Exit Function
ChkErr:
    Call ErrSub(FormName, "Function GetGlobal")
End Function

Public Function GetDT(dtm As String, DTtype As GiDTtpye) As String
  On Error GoTo ChkErr
    If Len(dtm) = 0 Then Exit Function
    If InStr(dtm, "/") = 0 Then MsgBox "傳進 GetDT 之日期時間必須有遮罩!!", vbInformation, "訊息.."
    Select Case DTtype
           Case GiDate
                GetDT = Format(dtm, "ee/mm/dd")
           Case GiTime
                GetDT = Format(dtm, "hh:mm")
           Case GiYM
                GetDT = Format(dtm, "ee/mm")
           Case GiYear
                GetDT = Format(dtm, "ee")
           Case GiMonth
                GetDT = Month(dtm)
           Case GiDay
                GetDT = Day(dtm)
           Case GiHour
                GetDT = Hour(dtm)
           Case GiMinute
                GetDT = Minute(dtm)
           Case GiSecond
                GetDT = Second(dtm)
           Case GiMD
                GetDT = Format(dtm, "mm/dd")
    End Select
  Exit Function
ChkErr:
    Call ErrSub(FormName, "Function GetDT")
End Function

Public Function csmsgTranErr(Optional strmsg As String)
    MsgBox " [" & strmsg & "]" & gimsgTranErr, vbCritical, gimsgWarning
End Function

Public Function GetMClass(intIndex As Integer, intCodeno As Variant) As Variant
    On Error GoTo ChkErr
        GetMClass = Val(GetRsValue("Select GroupNo From " & strOwnerName & "CD00" & 4 + intIndex & " Where CodeNo = '" & intCodeno & "'") & "")
    Exit Function
ChkErr:
    ErrSub FormName, "GetMClass"
End Function

Public Function ExecuteCommand(strSQL As String, Optional Conn As ADODB.Connection, Optional ByRef lngAffected As Long) As Boolean
    On Error GoTo ChkErr
        If Conn Is Nothing Then Set Conn = gcnGi
        Conn.Execute strSQL, lngAffected
        ExecuteCommand = True
    Exit Function
ChkErr:
    ErrSub FormName, "ExecuteCommand, SQL : [" & strSQL & "]", True
End Function

Public Function GetRS(ByRef rs As ADODB.Recordset, strSQL As String, Optional Conn As ADODB.Connection, _
                    Optional CursorLocation As CursorLocationEnum = adUseClient, _
                    Optional CursorType As CursorTypeEnum = adOpenForwardOnly, _
                    Optional LockType As LockTypeEnum = adLockReadOnly, Optional strmsg As String, _
                    Optional MaxRecords As Long, Optional blnCutConnection As Boolean, Optional blnExecute As Boolean) As Boolean
  On Error GoTo ChkErr
  Dim intItem As String
    If Conn Is Nothing Then Set Conn = gcnGi
    intItem = 1
    If rs.State = adStateOpen Then
        intItem = 11
        If rs.RecordCount > 0 Then rs.CancelUpdate
        intItem = 12
        rs.Close
    End If
    intItem = 2
    rs.MaxRecords = MaxRecords
    rs.CursorLocation = CursorLocation
    intItem = 3
    If blnExecute Then
        intItem = 31
        Set rs = Conn.Execute(strSQL)
    Else
        intItem = 32
        rs.Open strSQL, Conn, CursorType, LockType
    End If
    intItem = 4
    If strmsg <> "" And rs.RecordCount = 0 Then
        MsgBox strmsg, vbCritical, gimsgPrompt
    End If
    intItem = 5
    
    If blnCutConnection Then Set rs.ActiveConnection = Nothing
    GetRS = True
  Exit Function
ChkErr:
    GetRS = False
    Call ErrSub(FormName, "GetRS;Item(" & intItem & ");SQL: " & strSQL & "]")
End Function

Public Function GetRsValue(strSQL As String, Optional cnnObj As ADODB.Connection, Optional MsgNoData As String) As Variant
    On Error GoTo ChkErr
    Dim rs As New ADODB.Recordset
        If cnnObj Is Nothing Then Set cnnObj = gcnGi
        If Not GetRS(rs, strSQL, cnnObj) Then Exit Function
        If rs.RecordCount = 0 Then
            If Len(MsgNoData) > 0 Then
                MsgBox MsgNoData, vbExclamation, gimsgPrompt
            End If
            GetRsValue = Null
        Else
            GetRsValue = rs(0)
        End If
        Call CloseRecordset(rs)
    Exit Function
ChkErr:
    ErrSub FormName, "GetRsValue"
End Function

Public Sub ErrSub(FormName As String, ProcedureName As String, Optional blnADOErr As Boolean)
    Dim strErr As String
    Dim aryErr As Variant
    Dim strCrLf As String
    Dim retval As Variant
    Dim PrintErrFlag As Boolean
'    DoEvents
    
    Screen.MousePointer = vbDefault
    aryErr = Array("發生錯誤時間 : ", "發生錯誤專案 : ", "發生錯誤表單 : ", "發生錯誤程序 : ", "錯誤代碼 : ", "錯誤原因 : ")
    strCrLf = vbCrLf + vbCrLf
    If blnADOErr Then
        strErr = aryErr(0) + CStr(RightNow) + strCrLf + _
                 aryErr(1) + CStr(Err.Source) + strCrLf + _
                 aryErr(2) + FormName + strCrLf + _
                 aryErr(4) + CStr(Err.Number) + strCrLf + _
                 aryErr(5) + CStr(Err.Description) + strCrLf + _
                 aryErr(3) + ProcedureName
    Else
        strErr = aryErr(0) + CStr(RightNow) + strCrLf + _
                 aryErr(1) + CStr(Err.Source) + strCrLf + _
                 aryErr(2) + FormName + strCrLf + _
                 aryErr(3) + ProcedureName + strCrLf + _
                 aryErr(4) + CStr(Err.Number) + strCrLf + _
                 aryErr(5) + CStr(Err.Description)
    End If
            
    On Error Resume Next
'    clsStoreParameter.uErr = Err.Number
    MsgBox strErr, vbCritical, gimsgWarning
    'ErrLog strErr
'    RetVal = MsgBox(strErr, vbYesNo + vbDefaultButton2, "系統執行錯誤..(按 <是> 列印錯誤訊息)!!")
'    If RetVal = vbYes Then
'       With Printer
'           .FontSize = 14
'            Printer.Print "": Printer.Print "": Printer.Print ""
'            Printer.Print strErr
'           .NewPage
'       End With
'    End If
End Sub

Public Sub ErrLog(ByVal ErrString As String)
  On Error Resume Next
    Dim LogFile As String
    
    Dim Fnum As Integer
    
    LogFile = gErrLogPath
    
    If Dir(LogFile) = "" Then Exit Sub
    
    If Right(LogFile, 1) <> "\" Then LogFile = LogFile & "\"
    LogFile = LogFile & "SYSERR.TXT"
    
    ErrString = Replace(ErrString, vbCrLf & vbCrLf, vbCrLf)
    ErrString = "發生錯誤的系統 : 開博客服營運管理系統" & vbCrLf & _
                "發生錯誤的電腦 : " & CreateObject("WScript.NetWork").ComputerName & vbCrLf & _
                "使用者名稱 : " & garyGi(1) & vbCrLf & ErrString
    
    Fnum = FreeFile
    Open LogFile For Append As Fnum
    Print #Fnum, ErrString & vbCrLf & String(66, "-") & vbCrLf
    Close Fnum
End Sub

'
'Private Sub PrintOutX(ErrorString As String)
'  On Error Resume Next
'    Dim objPrinter As Printer
'    Dim PrinterName As String
'    Dim PrintOutFlag As Boolean
'    PrinterName = UCase(GetPrinterName(7))
'    For Each objPrinter In Printers
'        If UCase(objPrinter.DeviceName) = PrinterName Then
'            PrintOutFlag = True
'            Set Printer = objPrinter
'            Printer.FontSize = 14
'            Printer.Print "": Printer.Print "": Printer.Print ""
'            Printer.Print ErrorString
'            Printer.NewPage
'            Exit For
'        End If
'    Next
'    If Not PrintOutFlag Then
'        Printer.FontSize = 14
'        Printer.Print "": Printer.Print "": Printer.Print ""
'        Printer.Print ErrorString
'        Printer.NewPage
'    End If
'End Sub
'
''Get PrinterName
''Function GetPrinterName
''Parameter1: PrinterID
''Return: Printer Name
'Public Function GetPrinterName(PrinterID) As String
'  On Error GoTo ChkErr
'
'    Dim rsPrt As New ADODB.Recordset
'    Dim PrinterIndex As String
'    Dim PrinterName As String
'
'    rsPrt.CursorLocation = adUseServer
'    PrinterIndex = Choose(PrinterID, "WorkPrinter1", "WorkPrinter2", "WorkPrinter3", "BillPrinter", "RptPrinter1", "RptPrinter2", "RptPrinter3")
'
'    rsPrt.Open UCase("SELECT " & PrinterIndex & " FROM SO045"), gcnGi, adOpenForwardOnly, adLockReadOnly
'
''    Debug.Print "PrinterID : " & PrinterID
''    MsgBox rsPrt(PrinterIndex) & ""
'
'    PrinterName = rsPrt(PrinterIndex) & ""
'
'    If Len(Environ("OS")) = 0 Then
'        If Left(PrinterName, 2) = "\\" Then
'            PrinterName = Mid(PrinterName, (InStr(3, PrinterName, "\")) + 1)
'        End If
'        GetPrinterName = PrinterName
'    Else
'        GetPrinterName = PrinterName
'    End If
'
''   Debug.Print "PrinterName : " & GetPrinterName
'
'    rsPrt.Close
'    Set rsPrt = Nothing
'
'  Exit Function
'ChkErr:
'    Call ErrSub("PrintModule", "Function GetPrinterName")
'End Function
'
Public Function GetNullString(varValue As Variant, _
    Optional lngVType As giVType = giStringV, _
    Optional DBFlag As ChangeDB = giOracle, _
    Optional blnAddTime As Boolean = False) As Variant
    On Error GoTo ChkErr
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
                GetNullString = "'" & varValue & "'"
            End If
        End If
    Exit Function
ChkErr:
    ErrSub "Utility5", "GetNullString"
End Function

Public Function GetFieldValue(rs As ADODB.Recordset, strFieldName As String)
    On Error GoTo ChkErr
    Dim varField As Variant
        varField = rs.Fields(strFieldName).Value
        If rs.Fields(strFieldName).Type = 131 Or rs.Fields(strFieldName).Type = 139 Then
            If Len(rs.Fields(strFieldName).Value & "") > 0 Then
                varField = Val(rs.Fields(strFieldName).Value)
            End If
        End If
        GetFieldValue = varField
    Exit Function
ChkErr:
    ErrSub FormName, "GetFiledValue"
End Function

Public Function GetInvoiceNo(strInvType As String, strServiceType As String) As String
    On Error GoTo ChkErr
    Dim rsInv As New ADODB.Recordset
    Dim strSeq As String
    Dim strSType As String
        '取業者服務種類
        Select Case strServiceType
            Case "C"
                strSType = "CATV"
            Case "I"
                strSType = "ISP"
            Case ""
                strSType = "CATV"
        End Select
        '取Seqence Object
        Select Case UCase(strInvType)
            Case "I"
                strSeq = strOwnerName & "S_SO007_" & strSType
            Case "M"
                strSeq = strOwnerName & "S_SO008_" & strSType
            Case "P"
                strSeq = strOwnerName & "S_SO009_" & strSType
            Case "T"
                strSeq = strOwnerName & "S_SO033_" & strSType
            Case Else
                MsgBox "請傳入正確的單據類別", vbCritical, gimsgWarning
                Exit Function
        End Select
        'YYYYMMXx00000000
        'X:單據類別,x:服務類別
        If GetRS(rsInv, "SELECT '" & Format(Date, "YYYYMM") & _
        strInvType & strServiceType & "' || Ltrim(To_Char(" & strSeq & ".NextVal, '0999999')) FROM Dual", gcnGi) = False Then Exit Function
        GetInvoiceNo = rsInv(0).Value
        Set rsInv = Nothing
    Exit Function
ChkErr:
    ErrSub FormName, "GetInvoiceNo"
End Function

Public Function ChangeToDate(ByVal strDate As Variant, Optional blnAddTime As Boolean) As String
    On Error Resume Next
        If strDate & "" = "" Then Exit Function
        If blnAddTime Then
            ChangeToDate = Format(strDate, "yyyy/mm/dd hh:mm:ss")
        Else
            ChangeToDate = Format(strDate, "yyyy/mm/dd")
        End If
        
End Function

'將日期字串轉換成日期格式
Public Function StrToDate(ByVal strDate As String) As String
  On Error GoTo ChkErr
    StrToDate = Format(strDate, "####/##/##")
  Exit Function
ChkErr:
    ErrSub FormName, "StrToDate"
End Function

Public Function InsertToOracle(strTable As String, _
            rsSox As ADODB.Recordset, _
            Optional strWhere As String, _
            Optional lngAffected As Long = 0, _
            Optional blnAddNew As Boolean = True) As Boolean
    On Error GoTo ChkErr
    Dim intLoop As Integer
    Dim rsTmp As New ADODB.Recordset
    Dim strField As String
    Dim strValue As String
    Dim strSQL As String
    Dim strFullField As String
'    MsgBox strTable
        If rsSox.State = adStateClosed Then Exit Function
        If strWhere = "" Then strWhere = " Where 1 = 0 "
        If Not GetRS(rsTmp, "Select * From " & strTable & " " & strWhere, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
        Call ChkHaveField(strFullField, "", rsSox)
        If rsTmp.RecordCount = 0 And blnAddNew Then
            For intLoop = 0 To rsTmp.Fields.Count - 1
                '檢查是不有該欄位
                If ChkHaveField(strFullField, rsTmp(intLoop).Name, rsSox) Then
                    If Len(rsSox(rsTmp(intLoop).Name) & "") > 0 Then
                        strField = strField & "," & rsTmp(intLoop).Name
                        If rsTmp(intLoop).Type = adDBTimeStamp Then
                            strValue = strValue & "," & GetNullString(rsSox(rsTmp(intLoop).Name).Value, giDateV, giOracle, True)
                        Else
                            strValue = strValue & "," & GetNullString(ReplaceSpecialChar(rsSox(rsTmp(intLoop).Name).Value), giStringV, giOracle)
                        End If
                    End If
                End If
            Next
            strValue = "Insert Into " & strTable & " (" & Mid(strField, 2) & ") Values (" & Mid(strValue, 2) & ")"
        Else
            For intLoop = 0 To rsTmp.Fields.Count - 1
                If ChkHaveField(strFullField, rsTmp(intLoop).Name, rsSox) Then
                    If rsTmp(intLoop).Type = adDBTimeStamp Then
                        strField = strField & "," & rsTmp(intLoop).Name & "=" & GetNullString(rsSox(rsTmp(intLoop).Name).Value, giDateV, giOracle, True)
                    Else
                        strField = strField & "," & rsTmp(intLoop).Name & "=" & GetNullString(ReplaceSpecialChar(rsSox(rsTmp(intLoop).Name).Value), giStringV, giOracle)
                    End If
                End If
            Next
            strValue = "Update " & strTable & " Set " & Mid(strField, 2) & " " & strWhere
        End If
        If Not ExecuteCommand(strValue, gcnGi, lngAffected) Then Exit Function
'        If rsTmp.RecordCount = 0 Then rsTmp.AddNew
'        For intLoop = 0 To rsTmp.Fields.Count - 1
'             rsTmp(intLoop) = GetFieldValue(rsSox, rsTmp(intLoop).Name)
'        Next
'        rsTmp.Update
        'lngAffected = 1
        Call CloseRecordset(rsTmp)
        InsertToOracle = True
    Exit Function
ChkErr:
    ErrSub FormName, "InsertToOracle"
End Function

Public Function ReplaceSpecialChar(ByVal strValue As Variant) As String
    On Error Resume Next
    Dim strChar As String
        strValue = strValue & ""
        strChar = Replace(strValue, ",", "'||chr(44)||'")
        strChar = Replace(strValue, "'", "'||chr(39)||'")
        strChar = Replace(strValue, Chr(34), "'||Chr(34)||'")
        
End Function


Public Function ChkHaveField(ByRef strFullField As String, ByVal strField As String, _
    rsSource As ADODB.Recordset) As Boolean
    On Error GoTo ChkErr
    Dim intLoop As Integer
        If strFullField = "" Then
            For intLoop = 0 To rsSource.Fields.Count - 1
                strFullField = strFullField & ",'" & rsSource(intLoop).Name & "'"
            Next
            strFullField = Mid(strFullField, 2)
        Else
            If strField = "" Then Exit Function
            If InStr(1, strFullField, "'" & strField & "'") = 0 Then Exit Function
        End If
        ChkHaveField = True
    Exit Function
ChkErr:
    ChkHaveField = False
End Function

Public Function GetParaValue(strTable As String, strPara As String, _
            ServiceType As String, CompCode As String, ByRef lngParaValue As Long) As Boolean
        On Error GoTo ChkErr
        Dim rs As New ADODB.Recordset
        Dim strWhere As String
        If Len(ServiceType) > 0 Then
            strWhere = " ServiceType = '" & ServiceType & "'"
        End If
        If Len(CompCode) > 0 Then
            strWhere = strWhere & IIf(strWhere = "", "", " And ") & " CompCode = " & CompCode
        End If
        If strWhere <> "" Then strWhere = " Where " & strWhere
        
        lngParaValue = Val(GetRsValue("Select " & strPara & " from " & strOwnerName & strTable & strWhere) & "")
        On Error Resume Next
        Call CloseRecordset(rs)
        GetParaValue = True
        Exit Function
ChkErr:
    ErrSub FormName, "GetParaValue"
End Function

Public Function GetRefCode(ByRef strCodeNo As String, _
            ByRef strDescription As String, StrTableName As String, _
            Optional lngRefNo As Long) As Boolean
    On Error GoTo ChkErr
        Dim rsTmp As New ADODB.Recordset
        If Len(StrTableName) > 0 Then
            If Not GetRS(rsTmp, "Select CodeNo ,Description From " & StrTableName & " Where RefNo = " & lngRefNo, , , , , , 1) Then Exit Function
            strCodeNo = rsTmp("CodeNo")
            strDescription = rsTmp("Description")
        End If
        Set rsTmp = Nothing
    Exit Function
ChkErr:
    ErrSub FormName, "GetRefCode"
End Function


Public Function SFGetAmount(ByVal blnChoose As Boolean, ByVal P_OPTION As Integer, _
                ByVal P_CUSTID As Long, _
                ByVal P_CITEMCODE As Long, _
                ByVal P_PERIOD As Integer, _
                ByVal P_REALSTARTDATE As String, _
                ByVal p_SERVICETYPE As String, _
                ByRef P_REALSTOPDATE As String, _
                ByRef P_SHOULDAMT As Long, _
                ByRef P_RETMSG As String, _
                ByRef P_RealPeriod As Integer) As Integer
    On Error GoTo ChkErr
    Dim comSF_GETAMOUNT As New ADODB.Command
    
    With comSF_GETAMOUNT
            .Parameters.Append .CreateParameter("RETURN_VALUE", adVarNumeric, adParamReturnValue)
            .Parameters.Append .CreateParameter("P_OPTION", adVarNumeric, adParamInput, , P_OPTION)
            .Parameters.Append .CreateParameter("P_CUSTID", adVarNumeric, adParamInput, , P_CUSTID)
            .Parameters.Append .CreateParameter("P_CITEMCODE", adVarNumeric, adParamInput, , P_CITEMCODE)
            .Parameters.Append .CreateParameter("P_PERIOD", adVarNumeric, adParamInput, , P_PERIOD)
            .Parameters.Append .CreateParameter("P_REALSTARTDATE", adVarChar, adParamInput, 2000, P_REALSTARTDATE)
            If blnChoose Then
                'SF_GetAmount2
                .Parameters.Append .CreateParameter("P_REALSTOPDATE", adVarChar, adParamInput, 2000, P_REALSTOPDATE)
                .Parameters.Append .CreateParameter("P_SERVICETYPE", adVarChar, adParamInput, 2000, p_SERVICETYPE)
            Else
                'SF_GetAmount
                .Parameters.Append .CreateParameter("P_SERVICETYPE", adVarChar, adParamInput, 2000, p_SERVICETYPE)
                .Parameters.Append .CreateParameter("P_REALSTOPDATE", adVarChar, adParamOutput, 2000, P_REALSTOPDATE)
            End If
            .Parameters.Append .CreateParameter("P_SHOULDAMT", adVarNumeric, adParamOutput, , P_SHOULDAMT)
            If Not blnChoose Then
                .Parameters.Append .CreateParameter("p_RealPeriod", adVarNumeric, adParamOutput, , P_RealPeriod)
            End If
            .Parameters.Append .CreateParameter("P_RETMSG", adVarChar, adParamOutput, 2000, P_RETMSG)
            Set .ActiveConnection = gcnGi
            .CommandText = IIf(blnChoose, "SF_GETAMOUNT2", "SF_GETAMOUNT")
            .CommandType = adCmdStoredProc
            .Execute
            If Not blnChoose Then
                P_REALSTOPDATE = .Parameters("P_REALSTOPDATE").Value & ""
                P_RealPeriod = .Parameters("p_RealPeriod").Value & ""
            End If
            P_SHOULDAMT = Val(.Parameters("P_SHOULDAMT").Value & "")
            P_RETMSG = .Parameters("P_RETMSG").Value & ""
            SFGetAmount = Val(.Parameters("Return_value").Value & "")
    End With
    Exit Function
ChkErr:
    Call ErrSub(FormName, "SFGetAmount")
End Function

Public Function GetDTString(dtm As String) As String
  On Error GoTo ChkErr
    GetDTString = Format(dtm, "ee/mm/dd hh:mm:ss")
  Exit Function
ChkErr:
    Call ErrSub(FormName, "Function GetDTString")
End Function

Public Function NoZero(strVal As Variant, Optional ZeroFlag As Boolean = False) As Variant
  On Error GoTo ChkErr
    strVal = strVal & ""
    If Trim(strVal) = "" Then
        If ZeroFlag = True Then
            NoZero = 0
        Else
            NoZero = Null
        End If
    Else
        NoZero = strVal
    End If
  Exit Function
ChkErr:
    ErrSub FormName, "NoZero"
End Function

Public Function GetString(ByVal strS As Variant, ByVal intLength As Integer, Optional ByVal lngArrow As giArrow = giLeft, Optional ByVal InsertZero As Boolean = False) As String
  On Error GoTo ChkErr
    Dim strCh As String
    strCh = CStr(strS & "")
    If lngArrow = giLeft Then
        strCh = StrConv(LeftB(StrConv(strCh & Space(intLength), vbFromUnicode), intLength), vbUnicode)
    Else
        strCh = StrConv(RightB(StrConv(Space(intLength) & strCh, vbFromUnicode), intLength), vbUnicode)
    End If
    If InsertZero Then
    If Trim(strCh) = "" Then strCh = "0"
        If lngArrow = giLeft Then
            strCh = RTrim(strCh) & String(intLength - Len(RTrim(strCh)), "0")
        Else
            strCh = String(intLength - Len(LTrim(strCh)), "0") & LTrim(strCh)
        End If
    End If
    strCh = Replace(strCh, Chr(0), "")
    GetString = strCh
  Exit Function
ChkErr:
    Call ErrSub(FormName, "GetString")
End Function

Public Function OpenFile(ByRef objFile As Object, _
                        ByVal FilePath As String, _
                        ByVal intCount As Integer, _
                        Optional OverWrite As Boolean = True, _
                        Optional OpenFileType As giOpenFileType = giCreateTEXT) As Boolean
On Error GoTo ChkErr
    On Error GoTo ChkErr
    Dim fso As Object
        OpenFile = False
        Set fso = CreateObject("Scripting.FileSystemObject")
        If Not fso.FolderExists(Left(FilePath, InStrRev(FilePath, "\"))) Then
            MsgBox "路徑不正確!", vbExclamation, "警告..."
            OpenFile = False
            Exit Function
        End If

        If fso.FileExists(FilePath) Then
            'ReadOnly=1 是否為唯讀
            If fso.GetFile(FilePath).Attributes = 1 Then
               MsgBox "使用者沒有 " & FilePath & " 的使用權限", vbExclamation, "警告"
               OpenFile = False
               Exit Function
            End If
        End If
        If OpenFileType = giCreateTEXT Then
            Set objFile = fso.CreateTextFile(FilePath, OverWrite)
        Else
            On Error GoTo Msg
            Set objFile = fso.OpenTextFile(FilePath)
        End If
        OpenFile = True
        Set fso = Nothing
Exit Function
Msg:
    MsgBox "找不到檔案！！", vbExclamation, "錯誤！"
    OpenFile = False
    Exit Function
    
ChkErr:
    Call ErrSub("clsModule", "OpenFile")
End Function

Public Function WriteTextLine(objFile As Object, _
        ByVal strData As String) As Boolean
On Error GoTo ChkErr
    objFile.WriteLine strData
    WriteTextLine = True
Exit Function
ChkErr:
    Call ErrSub("DealFun", "WriteTextLine")
End Function

Public Sub CloseRecordset(rs As ADODB.Recordset)
    On Error Resume Next
        rs.CancelUpdate
        If rs.State = adStateOpen Then rs.Close
        Set rs = Nothing
End Sub

Public Function WriteLineData(ByVal vData As String, ByVal objFile As Object)
On Error GoTo ChkErr
    Call objFile.WriteLine(vData)
'    Call file(intCount).WriteLine(vData)
Exit Function
ChkErr:
    Call ErrSub("clsModule", "WriteLineData")
End Function

Public Function GetUseIndexStr(StrTableName As String, strColumnName As String) As String
    On Error Resume Next
        GetUseIndexStr = UCase(" /*+ Index(" & StrTableName & " I_" & StrTableName & "_" & strColumnName & ") */ ")
End Function

Public Function GetTmpMDBExecuteStr(rs As ADODB.Recordset, StrTableName As String) As String
    On Error GoTo ChkErr
    Dim lngLoop As Long
    Dim strField As String
    Dim strValue As String
    Dim strVal
        For lngLoop = 0 To rs.Fields.Count - 1
            strField = strField & "," & rs.Fields(lngLoop).Name
            If rs.Fields(lngLoop).Type = adDBTimeStamp Then
                strVal = GetNullString(rs.Fields(lngLoop).Value, giDateV, giAccessDb, True)
            Else
                strVal = GetNullString(rs.Fields(lngLoop).Value)
            End If
            strValue = strValue & "," & strVal
        Next
        strField = Mid(strField, 2)
        strValue = Mid(strValue, 2)
        GetTmpMDBExecuteStr = "Insert Into " & StrTableName & " ( " & strField & ") Values (" & _
            strValue & ")"
    Exit Function
ChkErr:
    Call ErrSub(FormName, "GetTmpMDBExecuteStr")
End Function

Public Function GetMDBConnection(strPath As String) As Connection
    On Error GoTo ChkErr
        Dim cnTmpMDB As New ADODB.Connection
        cnTmpMDB.CursorLocation = adUseClient
        cnTmpMDB.Open UCase("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strPath & ";Persist Security Info=False")
        Set GetMDBConnection = cnTmpMDB
    Exit Function
ChkErr:
    Call ErrSub(FormName, "Function GetMDBConnection")
End Function

Public Function CreateMDBTable(rs As ADODB.Recordset, StrTableName As String, cnn As ADODB.Connection) As Boolean
    On Error Resume Next
    Dim lngLoop As Long
    Dim strField As String
    Dim strValue As String
      cnn.Execute "Drop Table " & StrTableName
    On Error GoTo ChkErr
      CreateMDBTable = False
      If rs.State = 2 Then Exit Function
      
      For lngLoop = 0 To rs.Fields.Count - 1
          strValue = strValue & "," & rs.Fields(lngLoop).Name
          Select Case rs.Fields(lngLoop).Type
                 Case adBigInt, adVarNumeric, adNumeric
                 '"Double"
                      strValue = strValue & " Double"
                 Case adBoolean
                 '"Logical"
                      strValue = strValue & " Logical"
                 Case adBSTR, adChar, adChapter, adVarChar, adVariant, adVarWChar
                 '"Text("
                      strValue = strValue & " Text(" & rs.Fields(lngLoop).DefinedSize * 2 & ")"
                 Case adDate, adDBDate, adDBTimeStamp, adDBTime
                 '"Date"
                      strValue = strValue & " Date"
          End Select
      Next
      strValue = Mid(strValue, 2)
      cnn.Execute "Create Table " & StrTableName & " (" & strValue & ")"
      CreateMDBTable = True
  Exit Function
ChkErr:
  Call ErrSub(FormName, "GetTmpMDBExecuteStr")
End Function

Public Function GetSystemPara(ByRef rs As ADODB.Recordset, strCompCode As String, _
    ByVal strServiceType As String, strTable As String, Optional strParaField As String = "", _
    Optional blnShowMsg As Boolean = True, Optional blnServiceTypeUseLike As Boolean = False) As Boolean
    On Error GoTo ChkErr
    Dim strSQL As String
    Dim blnFlag As Boolean
        If strParaField = "" Then strParaField = "*"
        'If rs Is Nothing Then Set rs = New ADODB.Recordset
        If blnServiceTypeUseLike Then
            strSQL = " Like '%" & strServiceType & "%'"
        Else
            strSQL = " = '" & strServiceType & "'"
        End If
        '先找CompCode 及ServiceType 都有的
        If strCompCode <> "" And strServiceType <> "" Then
            If Not GetRS(rs, "Select " & strParaField & " From " & strTable & " Where CompCode = " & strCompCode & " And ServiceType " & strSQL) Then Exit Function
            blnFlag = Not rs.EOF
        End If
        If Not blnFlag Then
            '再找CompCode 有的
            If strCompCode <> "" Then
                If Not GetRS(rs, "Select " & strParaField & " From " & strTable & " Where CompCode = " & strCompCode) Then Exit Function
                blnFlag = Not rs.EOF
            End If
            '最後找只要有這個 參數檔的
            If Not blnFlag Then
                If Not GetRS(rs, "Select " & strParaField & " From " & strTable) Then Exit Function
                blnFlag = Not rs.EOF
            End If
        End If
        If blnShowMsg And Not blnFlag Then MsgBox "在系統參數檔[" & strTable & "] 找不到 公司別代號為 [" & strCompCode & "] ,且服務類別為: [" & strServiceType & "]的資料,請查證!!", vbCritical, gimsgWarning
        GetSystemPara = True
    Exit Function
ChkErr:
    ErrSub FormName, "GetSystemPara"
End Function

Public Function GetSystemParaItem(strField As String, _
    strCompCode As String, strServiceType As String, strTable As String, _
    Optional blnShowMsg As Boolean = False, Optional blnServiceTypeUseLike As Boolean = False) As Variant
    On Error Resume Next
    Dim rs As New ADODB.Recordset
        If Not GetSystemPara(rs, strCompCode, strServiceType, strOwnerName & strTable, strField, False, blnServiceTypeUseLike) Then Exit Function
        GetSystemParaItem = rs(0)
End Function

Public Function Decrypt(EncryptionString As String) As String
  On Error GoTo ChkErr
    Decrypt = CreateObject("DevPowerEncrypt.EnCrypt").Decrypt(EncryptionString)
  Exit Function
ChkErr:
    Call ErrSub("ExtraModule", "Function Decrypt")
End Function

Public Function GetCommonConnection()
    On Error GoTo ChkErr
    Dim strSysPath As String
    Dim pstrDataSource As String
    Dim pstrUserId As String
    Dim pstrPassword As String
    Dim strConnection As String
        strSysPath = garyGi(11)
        pstrDataSource = GetCommonINI(strSysPath, "1")
        pstrUserId = GetCommonINI(strSysPath, "2")
        pstrPassword = GetCommonINI(strSysPath, "3")
        strConnection = "Provider=MSDAORA.1;Password=" & pstrPassword & ";" & _
            "User ID=" & pstrUserId & ";Data Source=" & pstrDataSource & ";" & _
            "Persist Security Info=True"
        GetCommonConnection = strConnection
Exit Function
ChkErr:
    ErrSub "Sys_Lib", "GetCommonConnection"
End Function

Private Function GetCommonINI(IniFileName As String, Key As String) As String
    On Error GoTo ChkErr
        Dim Length As Long, S As String
        S = String(1024, 0)
        Length = GetPrivateProfileString("Common", Key, "", S, Len(S), IniFileName)
        GetCommonINI = Decrypt(Left(S, Length))
    Exit Function
ChkErr:
    ErrSub "Sys_Lib", "GetCommonINI"
End Function

Public Function GetJoinSQL(strPara As String, ByRef strChoose As String, _
    Optional PutWhere As giArrow = giRight) As String
    On Error Resume Next
        If strChoose = "" Then
            strChoose = " " & strPara & " "
        Else
            If PutWhere = giRight Then
                strChoose = strChoose & " And " & strPara & " "
            Else
                strChoose = " " & strPara & " And " & strChoose
            End If
        End If
End Function

Public Sub ReleaseCOM(FrmName As Object)
  On Error Resume Next
    Dim o As Object
    For Each o In FrmName
        Set o = Nothing
    Next
End Sub

Public Function GetInvoiceNo3(StrTableName As String, Optional Conn As ADODB.Connection) As String
  On Error GoTo ChkErr
    Dim strSeq As String
    Dim strInv
    strSeq = strOwnerName & "S_" & StrTableName
    If Conn Is Nothing Then Set Conn = gcnGi
    strInv = GetRsValue("SELECT Ltrim(To_Char(" & strSeq & ".NextVal, '09999999')) FROM Dual", Conn)
    GetInvoiceNo3 = strInv
  Exit Function
ChkErr:
    ErrSub "Sys_Lib", "GetInvoiceNo3"
End Function

Public Function MidMbcs(ByVal str As String, start, Length)
   MidMbcs = StrConv(MidB(StrConv(str, vbFromUnicode), start, Length), vbUnicode)
End Function

Public Sub SetCombolListIndex(objCombol As Object, strValue As String)
    On Error GoTo ChkErr
    Dim lngLoop As Long
        If objCombol.ListCount > 0 Then
            For lngLoop = 0 To objCombol.ListCount - 1
                objCombol.ListIndex = lngLoop
                If objCombol.Text = strValue Then Exit Sub
            Next
        End If
        objCombol.ListIndex = -1
    Exit Sub
ChkErr:
    ErrSub FormName, "SetCombolListIndex"
End Sub

Public Function GetINIrs() As Object
  On Error GoTo ChkErr
    Dim INIrs As New ADODB.Recordset
    Dim S As String, Length As Long
    Dim intLoop As Integer
    Dim strINIname As String
    S = String(1024, 0)
    strINIname = garyGi(11)
    With INIrs
         If .State = 1 Then .Close
        .CursorLocation = adUseClient
        .Fields.Append "CC", adBSTR, 38
        .Fields.Append "F1", adBSTR, 38
        .Fields.Append "F2", adBSTR, 38
        .Fields.Append "F3", adBSTR, 38
        .Fields.Append "F4", adBSTR, 38
        .Fields.Append "F5", adBSTR, 38
        .Open
         For intLoop = 1 To 99
            If GetPrivateProfileString(intLoop, "1", "", S, Len(S), strINIname) <> 0 Then
               .AddNew
               .Fields("CC") = intLoop
                Length = GetPrivateProfileString(intLoop, "1", "", S, Len(S), strINIname)
               .Fields("F1") = Decrypt(Left(S, Length))
                Length = GetPrivateProfileString(intLoop, "2", "", S, Len(S), strINIname)
               .Fields("F2") = Decrypt(Left(S, Length))
                Length = GetPrivateProfileString(intLoop, "3", "", S, Len(S), strINIname)
               .Fields("F3") = Decrypt(Left(S, Length))
                Length = GetPrivateProfileString(intLoop, "4", "", S, Len(S), strINIname)
               .Fields("F4") = Decrypt(Left(S, Length))
                Length = GetPrivateProfileString(intLoop, "5", "", S, Len(S), strINIname)
               .Fields("F5") = Decrypt(Left(S, Length))
               .Update
            End If
         Next
    End With
    If INIrs.RecordCount > 0 Then INIrs.AbsolutePosition = 1
    Set GetINIrs = INIrs
  Exit Function
ChkErr:
    Call ErrSub("Sys_Lib", "GetINIrs")
End Function

Public Sub MsgNoRcd(Optional AdditionalStr As String)
  On Error Resume Next
    MsgBox gimsgNoRcd & AdditionalStr, vbExclamation, "提示"
End Sub

Public Function GetMediaBillNo(ByRef lngCustId As Long, ByRef strMediaBillNo As String, cn As ADODB.Connection) As Boolean
    On Error GoTo ChkErr
    
        lngCustId = Val(GetInvoiceNo3("SO033_MediaBillNo", cn) & "")
        strMediaBillNo = Format(Date, "yymm") & Format(lngCustId, "0000000")
        GetMediaBillNo = True
    Exit Function
ChkErr:
    ErrSub FormName, "GetMediaBillNo"
End Function

Public Function BatchUpdateMediaBillNo(rsSource As ADODB.Recordset, strChoose33 As String, _
    cn As ADODB.Connection) As Boolean
    On Error GoTo ChkErr
    Dim lngAffected As Long
    Dim strMediaBillNo As String
        With rsSource
            If .RecordCount > 0 Then .MoveFirst
            Do While Not .EOF
                '已產生媒體單號後不再重新產生      ----Crystal Edit
                If Trim(rsSource("MediaBillNo")) & "" = "" Then
                    strMediaBillNo = ""
                    If Not GetMediaBillNo(rsSource("CustId"), strMediaBillNo, cn) Then Exit Function
                    If Not ExecuteCommand("Update " & strOwnerName & "SO033 A Set MediaBillNo = '" & strMediaBillNo & "' Where BillNo = '" & rsSource("BillNo") & "' And CustId = " & rsSource("CustId") & IIf(Len(strChoose33) = 0, "", " And ") & strChoose33, cn, lngAffected) Then Exit Function
                    On Error Resume Next
                    .Fields("MediaBillNo") = strMediaBillNo
                    .Update
                End If
                On Error GoTo ChkErr
                .MoveNext
            Loop
        End With
        BatchUpdateMediaBillNo = True
    Exit Function
ChkErr:
    ErrSub FormName, "BatchUpdateMediaBillNo"
End Function

Public Function GetLastDate(ByVal YYYYMM As String) As String '取當月最後一天
  On Error GoTo ChkErr
    If Not IsDate(YYYYMM) Then Exit Function
    If InStr(1, YYYYMM, "/") > 0 Then
        GetLastDate = (DateAdd("M", 1, Format(YYYYMM, "YYYY/MM") & "/01") - 1)
    Else
        GetLastDate = (DateAdd("M", 1, Format(YYYYMM, "####/##") & "/01") - 1)
    End If
  Exit Function
ChkErr:
    Call ErrSub("Utility5", "GetLastDate")
End Function

Public Sub MsgMustBe(Optional AdditionalStr As String)
  On Error Resume Next
'    MsgBox gMsgIsDataOK & AdditionalStr, vbExclamation, "訊息.."
    Dim strmsg As String
    If AdditionalStr = "" Then
        strmsg = gMsgIsDataOK
    Else
        strmsg = "[" & AdditionalStr & "] 為必要欄位,須有值 !!"
    End If
    MsgBox strmsg, vbInformation, "訊息"
End Sub

Public Function GetZero(Optional intCount As Integer)
  On Error Resume Next
  Dim i As Integer
  Dim strZero As String
        For i = 1 To intCount
            strZero = strZero & "0"
        Next
  GetZero = strZero
End Function

Public Function RPxx(strValue As String) As String
  On Error Resume Next
    strValue = Trim(Replace(strValue, Chr(13), "", 1, , vbTextCompare))
    strValue = Trim(Replace(strValue, Chr(10), "", 1, , vbTextCompare))
    RPxx = Trim(Replace(strValue, Chr(9), "", 1, , vbTextCompare))
End Function

Public Function RightNow() As String
  On Error GoTo ChkErr
    If Val(gcnGi.Execute("SELECT SYSTIME FROM " & GetOwner & "SO041 WHERE ROWNUM <=1 ORDER BY COMPCODE").GetString(adClipString, , , , 1) & "") = 1 Then
        RightNow = gcnGi.Execute("SELECT TO_CHAR(SYSDATE,'YYYY/MM/DD') FROM DUAL").GetString & ""
        RightNow = RPxx(RightNow) & " " & RPxx(gcnGi.Execute("SELECT TO_CHAR(SYSDATE,'HH24:MI:SS') FROM DUAL").GetString & "")
    Else
        RightNow = Now
    End If
  Exit Function
ChkErr:
    ErrSub "SysLib", "RightNow"
End Function

Public Function GetBillNo_Old(strBillNo As String) As String
  On Error Resume Next
    Dim strYM As String
    
    '將西元年改成民國年(yymm)
    strYM = Format(Format(Left(strBillNo, 6) & "01", "####/##/##"), "EEMM")
    GetBillNo_Old = strYM & Mid(strBillNo, 7, 1) & Right(strBillNo, 6)
  Exit Function
End Function
