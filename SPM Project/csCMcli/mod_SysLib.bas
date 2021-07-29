Attribute VB_Name = "mod_SysLib"
Option Explicit

Public Declare Function GetComputerNameA Lib "kernel32" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long

Public Enum ChangeDB
    giAccessDb = 0 'Access Table
    giOracle = 1    'Oracle Table
End Enum

Public Enum giVType
    giLongV = 0
    giStringV = 1
    giDateV = 2
End Enum

Public Enum giArrow
    giLeft = 0
    giRight = 1
    giMiddle = 2
End Enum

Public Enum giOpenFileType
    giCreateTEXT = 0
    giOpenTEXT = 1
End Enum

Public Enum giEditModeEnu
    giEditModeView = 0      ' 0=顯示模式
    giEditModeEdit = 1      ' 1.編輯模式(EditMode)
    giEditModeInsert = 2    ' 2=新增模式
    giEditModeDelete = 3
End Enum

Public Enum GiDTtpye
    GiDate = 1
    GiTime = 2
    GiYM = 3
    GiYear = 4
    GiMonth = 5
    GiDay = 6
    GiHour = 7
    GiMinute = 8
    GiSecond = 9
    GiMD = 10
End Enum

Public Const gimsgTranErr = "的資料時發生錯誤: 請查明原因!!"
Public Const vmsgSendConn = "請先傳入連線!!"
Public Const gimsgWarning = "警告!"
Public Const gimsgPrompt = "提示"
Public Const gimsgNoRcd = "查無資料"
Public Const gimsgPleaseSendSNo = "請先傳入單據編號"
Public Const gimsgPleaseSetSO011 = "請先設定預約時段[SO011]"
Public Const gMsgIsDataOK = "此為必要欄位,須有值 !! "

Public gcnGi As New ADODB.Connection
Public cnn As New ADODB.Connection
'Public clsStoreParameter As Object

Public blnNoShowMsg As Boolean

Public garyGi(20) As String
Public gErrLogPath As String
Public strErrorMessage As String
Public UI_Mode As Boolean
Public strRetMsg As String

Public Const strCrLf = vbCrLf & vbCrLf
Public Const strYMDHMS = "YYYY/MM/DD HH:MM:SS"
Public Const strYMD = "YYYY/MM/DD"
Private Const Mod_Name = "mod_SysLib"

Public Sub AddRetMsg(ByVal Msg As String)
  On Error GoTo ChkErr
    If Len(Trim(strRetMsg)) > 0 Then
        If Len(Trim(Msg)) > 0 Then strRetMsg = strRetMsg & vbCrLf & Msg
    Else
        strRetMsg = Msg
    End If
  Exit Sub
ChkErr:
    ErrSub Mod_Name, "AddRetMsg"
End Sub

Public Function GetOwner(Optional strOwnerName As String) As String
  On Error Resume Next
    If strOwnerName = "" Then strOwnerName = garyGi(16)
    If strOwnerName <> "" Then GetOwner = strOwnerName
    If strOwnerName <> "" And InStr(1, strOwnerName, ".") = 0 Then GetOwner = strOwnerName & "."
End Function

Public Function PutGlobal()
  On Error GoTo ChkErr
    PutGlobal = Join(garyGi, Chr(9))
  Exit Function
ChkErr:
    ErrSub Mod_Name, "Function PutGlobal"
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
    ErrSub Mod_Name, "Function GetGlobal"
End Function

Public Function GetRS(ByRef rs As ADODB.Recordset, strSQL As String, Optional Conn As ADODB.Connection, _
                    Optional CursorLocation As CursorLocationEnum = adUseClient, _
                    Optional CursorType As CursorTypeEnum = adOpenForwardOnly, _
                    Optional LockType As LockTypeEnum = adLockReadOnly, Optional strMsg As String, _
                    Optional MaxRecords As Long, Optional blnCutConnection As Boolean, Optional blnExecute As Boolean) As Boolean
  On Error GoTo ChkErr
  Dim intItem As String
    If Conn Is Nothing Then Set Conn = gcnGi
    intItem = 1
    If rs.State = adStateOpen Then
        intItem = 11
        If Not (rs.EOF Or rs.BOF) Then rs.CancelUpdate
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
    If strMsg <> "" And rs.RecordCount = 0 Then
        strErrorMessage = strMsg
        AddRetMsg strMsg
        If Not blnNoShowMsg Then MsgBox strMsg, vbCritical, gimsgPrompt
    End If
    intItem = 5
    
    If blnCutConnection Then Set rs.ActiveConnection = Nothing
    GetRS = True
  Exit Function
ChkErr:
    GetRS = False
    ErrSub Mod_Name, "GetRS;Item(" & intItem & ");SQL: " & strSQL & "]"
End Function

'將日期字串轉換成日期格式
Public Function StrToDate(ByVal strDate As String, Optional ByVal blnAddTime As Boolean = False, _
    Optional ByVal blnAddMask As Boolean = True) As String
    On Error GoTo ChkErr
    Dim strMask As String
        If strDate = "" Then Exit Function
        If IsDate(strDate) Then
            If blnAddMask Then
                If blnAddTime Then
                    strMask = "yyyy/mm/dd hh:mm:ss"
                Else
                    strMask = "yyyy/mm/dd"
                End If
            Else
                If blnAddTime Then
                    strMask = "yyyymmddhhmmss"
                Else
                    strMask = "yyyymmdd"
                End If
            End If
        Else
            If blnAddMask Then
                If blnAddTime Then
                    strMask = "####/##/## ##:##:##"
                Else
                    strMask = "####/##/##"
                End If
            Else
                If blnAddTime Then
                    strMask = "##############"
                Else
                    strMask = "########"
                End If
            End If
        End If
        StrToDate = Format(strDate, strMask)
    Exit Function
ChkErr:
    ErrSub Mod_Name, "StrToDate"
End Function

Public Function GetRsValue(strSQL As String, Optional CnnObj As ADODB.Connection, _
    Optional MsgNoData As String, Optional MultiData As Boolean = False) As Variant
    On Error GoTo ChkErr
    Dim rs As New ADODB.Recordset
        If CnnObj Is Nothing Then Set CnnObj = gcnGi
        If Not GetRS(rs, strSQL, CnnObj) Then Exit Function
        If rs.RecordCount = 0 Then
            If Len(MsgNoData) > 0 Then
                AddRetMsg MsgNoData
                If Not blnNoShowMsg Then MsgBox MsgNoData, vbExclamation, gimsgPrompt
            End If
            GetRsValue = Null
        Else
            GetRsValue = rs(0)
            If MultiData Then GetRsValue = rs.GetString(, , ",", ",")
        End If
        Call CloseRecordset(rs)
    Exit Function
ChkErr:
    ErrSub Mod_Name, "GetRsValue"
End Function

Public Function GetSystemParaItem(strField As String, _
    strCompCode As String, strServiceType As String, strTable As String, _
    Optional blnShowMsg As Boolean = False, Optional blnServiceTypeUseLike As Boolean = False) As Variant
    On Error Resume Next
    Dim rs As New ADODB.Recordset
        If Not GetSystemPara(rs, strCompCode, strServiceType, strTable, strField, False, blnServiceTypeUseLike) Then Exit Function
        GetSystemParaItem = rs(0)
End Function

Public Function RightNow() As Date
  On Error Resume Next
    RightNow = gcnGi.Execute("SELECT TO_CHAR(SYSDATE,'YYYY/MM/DD HH24:MI:SS') FROM DUAL").GetString & ""
    If Err.Number <> 0 Then RightNow = Now
End Function

Public Function RightDate() As Date
  On Error Resume Next
    RightDate = Format(RightNow, "YYYY/MM/DD")
    If Err.Number <> 0 Then RightDate = Date
End Function

Public Sub MsgMustBe(Optional AdditionalStr As String)
  On Error Resume Next
'    MsgBox gMsgIsDataOK & AdditionalStr, vbExclamation, "訊息.."
    Dim strMsg As String
    If AdditionalStr = "" Then
        strMsg = gMsgIsDataOK
    Else
        strMsg = "[" & AdditionalStr & "] 為必要欄位,須有值 !!"
        AddRetMsg strMsg
    End If
    If Not blnNoShowMsg Then MsgBox strMsg, vbInformation, "訊息"
End Sub

Public Function GetSystemPara(ByRef rs As ADODB.Recordset, strCompCode As String, _
    ByVal strServiceType As String, ByVal strTable As String, Optional strParaField As String = "", _
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
            If Not GetRS(rs, "Select " & strParaField & " From " & GetOwner & strTable & " Where CompCode = " & strCompCode & " And ServiceType " & strSQL) Then Exit Function
            blnFlag = Not rs.EOF
        End If
        If Not blnFlag Then
            '再找CompCode 有的
            If strCompCode <> "" Then
                If Not GetRS(rs, "Select " & strParaField & " From " & GetOwner & strTable & " Where CompCode = " & strCompCode) Then Exit Function
                blnFlag = Not rs.EOF
            End If
            '最後找只要有這個 參數檔的
            If Not blnFlag Then
                If Not GetRS(rs, "Select " & strParaField & " From " & GetOwner & strTable) Then Exit Function
                blnFlag = Not rs.EOF
            End If
        End If
        If Not blnFlag Then AddRetMsg "在系統參數檔 [" & strTable & "] 找不到 公司別代號為 [" & strCompCode & "] , 且服務類別為: [" & strServiceType & "] 的資料 , 請查證 !!"
        If blnShowMsg And Not blnFlag Then MsgBox "在系統參數檔[" & strTable & "] 找不到 公司別代號為 [" & strCompCode & "] ,且服務類別為: [" & strServiceType & "]的資料,請查證!!", vbCritical, gimsgWarning
        GetSystemPara = True
    Exit Function
ChkErr:
    ErrSub Mod_Name, "GetSystemPara"
End Function

Public Function O2Date(varDT As Variant, Optional OnlyYMD As Boolean = False) As String
  On Error GoTo ChkErr
    If IsNull(varDT) Or varDT = Empty Or IsMissing(varDT) Then
        O2Date = "NULL"
    Else
        If OnlyYMD Then
            O2Date = "TO_DATE('" & Format(varDT, strYMD) & "','YYYY/MM/DD')"
        Else
            O2Date = "TO_DATE('" & Format(varDT, strYMDHMS) & "','YYYY/MM/DD HH24:MI:SS')"
        End If
    End If
  Exit Function
ChkErr:
    ErrSub Mod_Name, "O2Date"
End Function

Public Function CDT(varFld As Variant) As Date
  On Error GoTo ChkErr
    CDT = CDate(Format(varFld, strYMDHMS))
  Exit Function
ChkErr:
    ErrSub Mod_Name, "CDT"
End Function

Public Sub OnlyNum(KeyAscii As Integer, Optional AllowDot As Boolean = False)
  On Error Resume Next
    If KeyAscii = 8 Then Exit Sub
    If AllowDot Then
        If KeyAscii = 46 Then Exit Sub
    End If
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Public Function Get_Computer_Name() As String ' 取得電腦名稱 2005/02/26 By Hammer
  On Error GoTo ChkErr
    Dim lngLen As Long
    Dim strName As String
    lngLen = 32
    strName = String(lngLen, Chr(0))
    GetComputerNameA strName, lngLen
    Get_Computer_Name = Left(strName, lngLen)
    If Len(Get_Computer_Name) = 0 Then Get_Computer_Name = Environ("COMPUTERNAME")
    If Len(Get_Computer_Name) = 0 Then Get_Computer_Name = CreateObject("WScript.Network").ComputerName
    Get_Computer_Name = RPxx(Get_Computer_Name)
    Get_Computer_Name = Replace(Get_Computer_Name, Chr(0), "", 1)
  Exit Function
ChkErr:
    ErrSub Mod_Name, "Get_Computer_Name"
End Function

Public Function GetDT(dtm As String, DTtype As GiDTtpye) As String
  On Error GoTo ChkErr
    If Len(dtm) = 0 Then Exit Function
    If InStr(dtm, "/") = 0 Then
        If Not blnNoShowMsg Then strErrorMessage = "傳進 GetDT 之日期時間必須有遮罩!!": MsgBox strErrorMessage, vbInformation, "訊息.."
        AddRetMsg strErrorMessage
    End If
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
    ErrSub Mod_Name, "Function GetDT"
End Function

Public Function GetDTString(dtm As String) As String
  On Error GoTo ChkErr
    GetDTString = Format(dtm, "ee/mm/dd hh:mm:ss")
  Exit Function
ChkErr:
    ErrSub Mod_Name, "Function GetDTString"
End Function

Public Function RPxx(strValue As String) As String
  On Error Resume Next
    strValue = Trim(Replace(strValue, Chr(13), "", 1, , vbTextCompare))
    strValue = Trim(Replace(strValue, Chr(10), "", 1, , vbTextCompare))
    RPxx = Trim(Replace(strValue, Chr(9), "", 1, , vbTextCompare))
End Function

'drop sequence S_IPCMD_SEQNO;
'create sequence S_IPCMD_SEQNO
'   MINVALUE 1
'   MAXVALUE 999999999
'   INCREMENT BY 1
'   START WITH 1
'   NOCACHE
'   NOCYCLE;

Public Function Get_Cmd_Sno() As String
  On Error GoTo ChkErr
    ' SEQ OBJ 格式為 CD052.EDESCRIPTION(取三碼)+YYYYMMDD+流水序號9碼，總長度20
'    Get_Cmd_Sno = RPxx(gcnGi.Execute("SELECT " & GetOwner & "S_IPCMD_SEQNO.NEXTVAL FROM DUAL").GetString & "")
    Get_Cmd_Sno = RPxx(gcnGi.Execute("SELECT " & GetOwner & "S_CMDSEQNO.NEXTVAL FROM DUAL").GetString & "")
    Get_Cmd_Sno = Format(gCompCode, "000") & _
                                Format(RightDate, "yyyymmdd") & _
                                Format(Get_Cmd_Sno, "000000000")
  Exit Function
ChkErr:
    ErrSub Mod_Name, "Get_Cmd_Sno"
End Function

Public Function ChkDTok() As Boolean
'  On Error GoTo ChkErr
  On Error Resume Next
   'If Not ChkDTok Then Exit Sub
    If TypeOf Screen.ActiveControl Is GiDate Or _
        TypeOf Screen.ActiveControl Is GiTime Or _
        TypeOf Screen.ActiveControl Is GiYM Then
        If Len(Trim(Screen.ActiveControl.GetValue)) = 0 Then
            ChkDTok = True
        Else
            ChkDTok = Screen.ActiveControl.RaiseValidateEvent
        End If
    Else
        ChkDTok = True
    End If
'  Exit Function
'ChkErr:
'    ErrSub "Sys_Lib", "ChkDTok"
End Function
  
Public Sub SetgiList(objGilist As Object, strFldName1 As String, strFldName2 As String, _
            strTableName As String, Optional lngTop As Long, Optional lngLeft As Long, Optional lngWidth1 As Long, _
            Optional lngWidth2 As Long, Optional lngFldLen1 As Long, Optional lngFldLen2 As Long, _
            Optional strWhere As String, Optional blnStopFlag As Boolean = False)
 '欄位1,欄位2,表格名稱,Top,Left,Width1,Width2,FldLen1,FldLen2
  On Error GoTo ChkErr
    objGilist.SetFldName1 strFldName1
    objGilist.SetFldName2 strFldName2
    objGilist.SetTableName GetOwner & strTableName
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
    ErrSub Mod_Name, "SetgiList"
End Sub

Public Sub SetgiMulti(objgiMulti As Object, strFldName1 As String, _
        strFldName2 As String, strTableName As String, strFldCaption1 As String, _
        strFldCaption2 As String, Optional strFilter As String, _
        Optional blnStopFlag As Boolean = False)
 '欄位1,欄位2,Table,中文1,中文2
  On Error GoTo ChkErr
    Call objgiMulti.SendConn(gcnGi)
    objgiMulti.FldName1 = strFldName1
    objgiMulti.FldName2 = strFldName2
    objgiMulti.TableName = GetOwner & strTableName
    objgiMulti.FldCaption1 = strFldCaption1
    objgiMulti.FldCaption2 = strFldCaption2
    objgiMulti.FilterStop = blnStopFlag
    If strFilter <> "" Then objgiMulti.Filter = strFilter
  Exit Sub
ChkErr:
    ErrSub Mod_Name, "SetgiMulti"
End Sub

Public Sub SetMsQry(objMs As Object, qryStr As String, _
                    Optional FldCaption1 As String = "", _
                    Optional FldCaption2 As String = "", _
                    Optional ColumnCaption As String = "", _
                    Optional ColumnWidth As String = "", _
                    Optional blnStopFlag As Boolean)
  On Error GoTo ChkErr
    With objMs
        .SendConn gcnGi
        .QueryString = qryStr
        .FilterStop = blnStopFlag
         If Len(FldCaption1) > 0 Then .FldCaption1 = FldCaption1
         If Len(FldCaption2) > 0 Then .FldCaption2 = FldCaption2
         If Len(ColumnCaption) > 0 Then .ColumnCaption = ColumnCaption
         If Len(ColumnWidth) > 0 Then .ColumnWidth = ColumnWidth
    End With
  Exit Sub
ChkErr:
    ErrSub Mod_Name, "SetMsQry"
End Sub

Public Function GetCompCodeFilter(objControl As Object, Optional strFilterStr As String) As String
  On Error Resume Next
    If strFilterStr = "" Then Exit Function
    If objControl.Filter = "" Then
        GetCompCodeFilter = "Where CodeNo in (" & strFilterStr & ")"
    Else
        GetCompCodeFilter = objControl.Filter & " And CodeNo in (" & strFilterStr & ")"
    End If
End Function

Public Function GetPermission(strMID As String) As Boolean
  On Error GoTo ChkErr
    Dim rs As New ADODB.Recordset
        If Val(garyGi(4)) = 0 Then GetPermission = True: Exit Function
        If gstrUserPremission = "" Then
            If Not GetRS(rs, "Select Mid From " & GetOwner & "So029 Where Link like 'SO1172%' And Group" & garyGi(4) & " = 1") Then Exit Function
            If Not rs.EOF Then gstrUserPremission = UCase(rs.GetString(, , ",", ",", ""))
        End If
        GetPermission = InStr(1, gstrUserPremission, UCase(strMID)) > 0
    Exit Function
ChkErr:
    ErrSub Mod_Name, "GetPermission"
End Function

Public Sub ReleaseCOM(FrmName As Object)
  On Error Resume Next
    Dim O As Object
    For Each O In FrmName
        Set O = Nothing
    Next
End Sub

Public Sub CloseRecordset(rs As ADODB.Recordset)
    On Error Resume Next
        rs.CancelUpdate
        If rs.State = adStateOpen Then rs.Close
        Set rs = Nothing
End Sub

Public Sub ErrSub(Mod_Name As String, ProcedureName As String, Optional blnADOErr As Boolean)
    Dim strErr As String
    Dim aryErr As Variant
    'Dim RetVal As Variant
    'Dim PrintErrFlag As Boolean
    
    Screen.MousePointer = vbDefault
    aryErr = Array("發生錯誤時間 : ", "發生錯誤專案 : ", "發生錯誤表單 : ", "發生錯誤程序 : ", "錯誤代碼 : ", "錯誤原因 : ")
    If blnADOErr Then
        strErr = aryErr(0) & CStr(Now) & strCrLf & _
                 aryErr(1) & CStr(Err.Source) & strCrLf & _
                 aryErr(2) & Mod_Name & strCrLf & _
                 aryErr(4) & CStr(Err.Number) & strCrLf & _
                 aryErr(5) & CStr(Err.Description) & strCrLf & _
                 aryErr(3) & ProcedureName
    Else
        strErr = aryErr(0) & CStr(Now) & strCrLf & _
                 aryErr(1) & CStr(Err.Source) & strCrLf & _
                 aryErr(2) & Mod_Name & strCrLf & _
                 aryErr(3) & ProcedureName & strCrLf & _
                 aryErr(4) & CStr(Err.Number) & strCrLf & _
                 aryErr(5) & CStr(Err.Description)
    End If
            
    On Error Resume Next
'    clsStoreParameter.uErr = Err.Number
    If Not blnNoShowMsg Then MsgBox strErr, vbCritical, gimsgWarning
    AddRetMsg strErr
    'ErrLog strErr
'    RetVal = MsgBox(strErr, vbYesNo & vbDefaultButton2, "系統執行錯誤..(按 <是> 列印錯誤訊息)!!")
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
    Dim fnum As Integer
    
    LogFile = gErrLogPath
    If Dir(LogFile) = "" Then Exit Sub
    
    If Right(LogFile, 1) <> "\" Then LogFile = LogFile & "\"
    LogFile = LogFile & "SYSERR.TXT"
    
    ErrString = Replace(ErrString, vbCrLf & vbCrLf, vbCrLf)
    ErrString = "發生錯誤的系統 : 開博客服營運管理系統" & vbCrLf & _
                "發生錯誤的電腦 : " & CreateObject("WScript.NetWork").ComputerName & vbCrLf & _
                "使用者名稱 : " & garyGi(1) & vbCrLf & ErrString
    
    fnum = FreeFile
    Open LogFile For Append As fnum
    Print #fnum, ErrString & vbCrLf & String(66, "-") & vbCrLf
    Close fnum
End Sub

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

'Public Sub SetgiMultiAddItem(objgiMulti As Object, strCodeNo As String, strDescription As String, Optional strFldCaption1 As String, Optional strFldCaption2 As String)
'  On Error GoTo ChkErr
'   Dim varCodeNo As Variant
'   Dim varDescription As Variant
'   Dim varValue As Variant
'   Dim intValue As Integer
'   objgiMulti.FldCaption1 = strFldCaption1
'   objgiMulti.FldCaption2 = strFldCaption2
'   objgiMulti.DIY = True
'   intValue = 0
'   varCodeNo = Split(strCodeNo, ",")
'   varDescription = Split(strDescription, ",")
'   For Each varValue In varCodeNo
'       objgiMulti.AddItem varValue, varDescription(intValue)
'       intValue = intValue + 1
'   Next
'  Exit Sub
'ChkErr:
'    ErrSub Mod_Name, "SetgiMultiAddItem")
'End Sub
'
'Public Sub GiMultiFilter(objMulti As Object, Optional strServiceType As String = "", Optional strCompCode As String = "")
'    On Error GoTo ChkErr
'    Dim strFilter As String
'      objMulti.SetDispStr ""
'        If TypeOf objMulti Is GiMulti Then
'            If strServiceType <> "" Then
'                strFilter = strFilter & " WHERE Nvl(SERVICETYPE,'" & Trim(strServiceType) & "') = '" & Trim(strServiceType) & "'"
'            End If
'
'            If strCompCode <> "" Then
'                strFilter = strFilter & IIf(strFilter = "", " WHERE ", " And ") & " Nvl(COMPCODE," & Trim(strCompCode) & ") = " & Trim(strCompCode) & ""
'            End If
'            objMulti.Filter = strFilter
'        End If
'
'    Exit Sub
'ChkErr:
'    ErrSub "Sys_Lib", "GiMultiFilter"
'End Sub

'Public Function MustExist(ctl, Optional Flag, Optional strMsg As String, Optional ByRef TabObj As Object, Optional ByVal TabIdx As Integer) As Boolean
'  On Error GoTo ChkErr
'    MustExist = True
'    Flag = IIf(IsMissing(Flag), 0, Flag)
'    If strMsg = "" Then
'        strMsg = gMsgIsDataOK
'    Else
'        strMsg = "[" & strMsg & "] 為必要欄位,須有值 !!"
'    End If
'    ' GiDate
'    If Flag = 1 Then
'        If ctl.GetValue = "" Then
'            If ctl.Enabled Then ctl.SetFocus
'            MsgBox strMsg, vbInformation, "訊息"
'            MustExist = False
'        End If
'    ' GiList
'    ElseIf Flag = 2 Then
'        If ctl.GetCodeNo = "" Or ctl.GetDescription = "" Then
'            If ctl.Enabled Then ctl.SetFocus
'            MsgBox strMsg, vbInformation, "訊息"
'            MustExist = False
'        End If
'    ' GiAddress2
'    ElseIf Flag = 3 Then
'        If ctl.IsDataOk = False Or Len(Trim(ctl.GetCodeNo)) = 0 Then
'            If ctl.Enabled Then ctl.SetFocus
'            MsgBox strMsg, vbInformation, "訊息"
'            MustExist = False
'        End If
'    ' GiAddress1
'    ElseIf Flag = 4 Then
'        If ctl.AddrNo = 0 Then
'            If ctl.Enabled Then
'                ctl.SetFocus
'            End If
'            MsgBox strMsg, vbInformation, "訊息"
'            MustExist = False
'        End If
'    ' Gimulti
'    ElseIf Flag = 5 Then
'        If ctl.GetDispStr = "" Then
'            If ctl.Enabled Then
'                ctl.SetFocus
'            End If
'            MsgBox strMsg, vbInformation, "訊息"
'            MustExist = False
'        End If
'    Else
'        If ctl.Text = "" Then
'            If Not TabObj Is Nothing Then
'                TabObj.Tab = TabIdx
'                If ctl.Enabled Then ctl.SetFocus
'                MsgBox strMsg, vbInformation, "訊息"
'                On Error GoTo 0
'                MustExist = False
'            Else
'                If ctl.Enabled Then ctl.SetFocus
'                MsgBox strMsg, vbInformation, "訊息"
'                On Error GoTo 0
'                MustExist = False
'            End If
'        End If
'    End If
'  Exit Function
'ChkErr:
'    If Err.Number = 5 Then Resume Next: Exit Function
'    ErrSub Mod_Name, "MustExist"
'End Function

'Public Sub objSelectAll(obj As Object)
'    On Error GoTo ChkErr
'        With obj
'            .SelStart = 0
'            .SelLength = Len(.Text)
'        End With
'    Exit Sub
'ChkErr:
'    ErrSub Mod_Name, "objSelectAll"
'End Sub

'Private Function InsertToOracle(strTable As String, _
'            rsSox As ADODB.Recordset, _
'            Optional strWhere As String, _
'            Optional lngAffected As Long = 0) As Boolean
'    On Error GoTo ChkErr
'    Dim intLoop As Integer
'    Dim rsTmp As New ADODB.Recordset
'    Dim strField As String
'    Dim strValue As String
'    Dim strSQL As String
'    Dim strFullField As String
'        If rsSox.State = adStateClosed Then Exit Function
'        If strWhere = "" Then strWhere = " Where 1 = 0 "
'        If Not GetRS(rsTmp, "Select * From " & strTable & " " & strWhere, , adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
'        Call ChkHaveField(strFullField, "", rsSox)
'        If rsTmp.RecordCount = 0 Then
'            For intLoop = 0 To rsTmp.Fields.Count - 1
'                '檢查是不有該欄位
'                If ChkHaveField(strFullField, rsTmp(intLoop).Name, rsSox) Then
'                    If Not IsNull(rsSox(rsTmp(intLoop).Name)) Then
'                        strField = strField & "," & rsTmp(intLoop).Name
'                        If rsTmp(intLoop).Type = adDBTimeStamp Then
'                            strValue = strValue & "," & GetNullString(rsSox(rsTmp(intLoop).Name).Value, giDateV, giOracle, True)
'                        Else
'                            strValue = strValue & "," & GetNullString(rsSox(rsTmp(intLoop).Name).Value, giStringV, giOracle)
'                        End If
'                    End If
'                End If
'            Next
'            strValue = "Insert Into " & strTable & " (" & Mid(strField, 2) & ") Values (" & Mid(strValue, 2) & ")"
'        Else
'            For intLoop = 0 To rsTmp.Fields.Count - 1
'                If ChkHaveField(strFullField, rsTmp(intLoop).Name, rsSox) Then
'                    If rsTmp(intLoop).Type = adDBTimeStamp Then
'                        strField = strField & "," & rsTmp(intLoop).Name & "=" & GetNullString(rsSox(rsTmp(intLoop).Name).Value, giDateV, giOracle, True)
'                    Else
'                        strField = strField & "," & rsTmp(intLoop).Name & "=" & Replace(GetNullString(rsSox(rsTmp(intLoop).Name).Value, giStringV, giOracle), ",", "'||chr(44)||'")
'                    End If
'                End If
'            Next
'            strValue = "Update " & strTable & " Set " & Mid(strField, 2) & " " & strWhere
'        End If
'        If Not ExecuteCommand(strValue, gcnGi, lngAffected) Then Exit Function
''        If rsTmp.RecordCount = 0 Then rsTmp.AddNew
''        For intLoop = 0 To rsTmp.Fields.Count - 1
''             rsTmp(intLoop) = GetFieldValue(rsSox, rsTmp(intLoop).Name)
''        Next
''        rsTmp.Update
'        lngAffected = 1
'        Call CloseRecordset(rsTmp)
'        InsertToOracle = True
'    Exit Function
'ChkErr:
'    ErrSub Mod_Name, "InsertToOracle"
'End Function

'Public Sub subShowMessage(strSetType As String, strCmdStatus As String, _
'    strErrorCode As String, strErrorMsg As String)
'    On Error GoTo ChkErr
'        If strCmdStatus = "C" Then
'            MsgBox "[" & strSetType & "] 控制訊號已送出完成 !! ", vbInformation, gimsgPrompt
'        ElseIf strCmdStatus = "E" Then
'            MsgBox strSetType & "失敗!! " & vbCrLf & "錯誤代碼:" & strErrorCode & " ,錯誤原因:" & strErrorMsg, vbInformation, gimsgPrompt
'        End If
'    Exit Sub
'ChkErr:
'    ErrSub Mod_Name, "subShowMessage"
'End Sub

'Public Sub ToolBarButtonClick(ButtonKey As String, _
'        objToolBar As Toolbar)
'    On Error Resume Next
'    Dim strActMod_Name As String
'        If objToolBar.Buttons(ButtonKey).Enabled = False Then Exit Sub
'        If objActForm Is Nothing Then Exit Sub
'        strActMod_Name = objActForm.Name
'        Select Case ButtonKey
'            Case "AddNew"
'                Call objActForm.AddNewGo
'            Case "Edit"
'                Call objActForm.EditGo
'            Case "Delete"
'                Call objActForm.DeleteGo
'            Case "Find"
'                Call objActForm.FindGo
'            Case "Print"
'                Call objActForm.PrintGO
'            Case "Save"
'                Call objActForm.UpdateGo
'            Case "First"
'                Call objActForm.FirstGo
'            Case "Previous"
'                Call objActForm.PreviousGo
'            Case "Next"
'                Call objActForm.NextGo
'            Case "Last"
'                Call objActForm.LastGo
'            Case "Cancel"
'                Call objActForm.CancelGo
'        End Select
'        If ChkFormLoad(strActMod_Name) Then If objActForm.Visible And objActForm.Enabled Then objActForm.SetFocus
'End Sub

'Public Function ChkFormLoad(strForm As String) As Boolean
'  On Error Resume Next
'    Dim lngLoop As Long
'        ChkFormLoad = True
'        For lngLoop = 0 To Forms.Count - 1
'            If UCase(Forms(lngLoop).Name) = UCase(strForm) Then Exit Function
'        Next
'        ChkFormLoad = False
'End Function

'Public Sub FunctionKeyGo(KeyCode As Integer, Shift As Integer, _
'        objToolBar As Toolbar)
'    On Error GoTo ChkErr
'    Dim ButtonKey As String
'        If Shift = 0 Then
'            Select Case KeyCode
'                Case vbKeyF2
'                    ButtonKey = "Update"
'                Case vbKeyF3
'                    ButtonKey = "Find"
'                Case vbKeyF6
'                    ButtonKey = "AddNew"
'                Case vbKeyF11
'                    ButtonKey = "Edit"
'                Case vbKeyF10
'                    ButtonKey = "Delete"
'                Case vbKeyF5
'                    ButtonKey = "Print"
'                Case vbKeyPageUp
'                    'ButtonKey = "Previous"
'                Case vbKeyPageDown
'                    'ButtonKey = "Next"
'                Case vbKeyEscape
'                    ButtonKey = "Cancel"
'            End Select
'        ElseIf Shift = 1 Then
'            If KeyCode = vbKeyX Then
'                ButtonKey = "Cancel"
'            End If
'        ElseIf Shift = 2 Then
'            Select Case KeyCode
'                Case vbKeyPageUp
'                    ButtonKey = "First"
'                Case vbKeyPageDown
'                    ButtonKey = "Last"
'                Case vbKeyF
'                    ButtonKey = "Find"
'            End Select
'        End If
'        If ButtonKey <> "" Then Call ToolBarButtonClick(ButtonKey, objToolBar)
'    Exit Sub
'ChkErr:
'    ErrSub Mod_Name, "FunctionKeyGo"
'End Sub

