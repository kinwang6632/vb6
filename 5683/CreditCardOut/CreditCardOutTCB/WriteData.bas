Attribute VB_Name = "DealFun"
Option Explicit

Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public objStorePara As clsStoreParameter
Public fso As New FileSystemObject
''Public errlogfile As New TextStream
Public file(10) As TextStream
'Public txtFile As TextStream
Public strUserName As String
Public objAction As Object
Public rsTmp1 As New ADODB.Recordset
Public Const gMsgIsDataOK = "此為必要欄位,須有值 !! "
Public blnNodata As Boolean           '判斷是否有資料
Public Const strMinRealStopDate As String = " MIN(DECODE(NVL(A.PAYKIND,0),1, " & _
                "DECODE(RealStopDate,NULL,TO_DATE('99991231','YYYYMMDD'),REALSTOPDATE)," & _
                "TO_DATE('99991231','YYYYMMDD'))) REALSTOPDATE "

Public Const strMinPayKind As String = " Max(Nvl(A.PayKind,0)) PayKind "
Public Enum giArrow
    giLeft = 0
    GIRIGHT = 1
End Enum

Enum giVType
    giLongV = 0
    giStringV = 1
    giDateV = 2
End Enum

Public Enum ChangeDB
    giAccessDb = 0 'Access Table
    giOracle = 1    'Oracle Table
End Enum
Public Enum ChkPayKind
    giSingle = 0
    giAll = 1
End Enum
Public TableOwnerName As String

Public Sub DefineRs()
    If rsTmp1.State = adStateOpen Then Exit Sub
    With rsTmp1
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Fields.Append "Billno", adBSTR, 15
        .Fields.Append "CustId", adBigInt
        .Fields.Append "Amount", adBigInt
        .Fields.Append "Content", adBSTR, 400
        .Open
    End With
End Sub

Public Sub CloseFS(ByVal intIndex As Integer)
   Dim intLoop As Integer
   For intLoop = 0 To intIndex
     ''     If IsObject(file(intLoop)) Then file(intLoop).Close
   Next
   Set fso = Nothing
End Sub

Public Function WriteLineData(ByVal vData As String, ByVal intCount As Integer)
On Error GoTo ChkErr
    Call file(intCount).WriteLine(vData)
Exit Function
ChkErr:
    Call ErrSub("DealFun", "WriteLineData")
End Function

Public Function OpenFile(ByVal FilePath As String, _
                         ByVal intCount As Integer, _
                         Optional OverWrite As Boolean = True, _
                         Optional ByVal intStateType As Integer = -1, _
                         Optional ByVal strMerchantName, _
                         Optional ByVal strSpcNo) As Boolean
On Error GoTo ChkErr
Dim strFileName As String
    
    If intStateType <> -1 Then
         strFileName = Left(strMerchantName, 3) & strSpcNo & Format(Date, "MMDD")
         Select Case intStateType
'         Case 0
'            strFileName = strFileName & ".01" & "_" & strSpcNo
'         Case 1
'            strFileName = strFileName & ".02" & "_" & strSpcNo
'         Case 2
'            strFileName = strFileName & ".03" & "_" & strSpcNo
'         Case 3
'            strFileName = strFileName & ".04" & "_" & strSpcNo
                     
                     Case 0
                       strFileName = strFileName & ".01"
                    Case 1
                       strFileName = strFileName & ".02"
                    Case 2
                       strFileName = strFileName & ".03"
                    Case 3
                       strFileName = strFileName & ".04"
         End Select
       
    End If
      OpenFile = False
    'If Not fso.FolderExists(Left(FilePath, InStrRev(FilePath, "\"))) Then

    
    If intStateType <> -1 Then
        
         If Not fso.FolderExists(FilePath) Then
           MsgBox "路徑不正確!", vbExclamation, "警告..."
           OpenFile = False
           Exit Function
         End If
          FilePath = IIf(Right(FilePath, 1) <> "\", FilePath & "\", FilePath) & strFileName
       ''   Set file(intCount) = fso.CreateTextFile(FilePath, OverWrite)
    End If
    
'       If fso.FileExists(FilePath) Then
'
'             If intStateType = -1 Then
'               Set file(intCount) = fso.OpenTextFile(FilePath, ForWriting)
'            Else
'               frmCitiBank.blnSameFile = True
'               Set file(intCount) = fso.OpenTextFile(FilePath, ForAppending)
'            End If
'        Else
'            frmCitiBank.blnSameFile = False
            Set file(intCount) = fso.CreateTextFile(FilePath, OverWrite)
 '       End If
 objStorePara.uProcText = FilePath
 OpenFile = True
Exit Function
ChkErr:
    MsgBox "檔案 " & FilePath & "處理發生問題 !!", vbOKOnly + vbCritical, "訊息"
    Call ErrSub("DealFun", "OpenFile")
End Function

Public Function InstLogDB(ByVal cn As ADODB.Connection, ByVal BankCode As Integer, _
        ByVal TranType As String, ByVal BillNO As String, CustId As Integer, ByVal Amount As Integer, _
        ByVal ChkLog As Integer, ByVal TranTime As String, ByVal Rectype As Integer, _
        ByVal Content As String)
On Error GoTo ChkErr
    Dim strsql As String
    strsql = "Insert Into So037 (BankCode,TranType,BillNo,Custid,Amount,ChkLog,tranTime,RecType,Content) values("
    strsql = strsql & BankCode & ",'1','" & BillNO & "'," & CustId & "," & Amount & ",0,to_date('" & TranTime & "','YYYY/MM/DD HH24:MI:SS')," & Rectype & ",'" & Content & "')"
    cn.Execute strsql
Exit Function
ChkErr:
    Call ErrSub("DealFun", "InstlogDB")
End Function

Public Sub ErrSub(FormName As String, ProcedureName As String)
    Dim strErr As String
    Dim aryErr As Variant
    Dim strCrLf As String
    Dim retval As Variant
    DoEvents
    Screen.MousePointer = vbDefault
    aryErr = Array("發生錯誤時間 : ", "發生錯誤專案 : ", "發生錯誤表單 : ", "發生錯誤程序 : ", "錯誤代碼 : ", "錯誤原因 : ")
    strCrLf = vbCrLf + vbCrLf
    strErr = aryErr(0) + CStr(Now) + strCrLf + _
             aryErr(1) + CStr(Err.Source) + strCrLf + _
             aryErr(2) + FormName + strCrLf + _
             aryErr(3) + ProcedureName + strCrLf + _
             aryErr(4) + CStr(Err.Number) + strCrLf + _
             aryErr(5) + CStr(Err.Description)
    ErrLog strErr
    retval = MsgBox(strErr, vbYesNo, "系統執行錯誤..(按 <是> 列印錯誤訊息)!!")
    If retval = vbYes Then
       With Printer
           .FontSize = 14
            Printer.Print "": Printer.Print "": Printer.Print ""
            Printer.Print strErr
           .NewPage
       End With
    End If
End Sub

Public Sub ErrLog(ByVal ErrString As String)
    On Error Resume Next
    Dim LogFile As String
    Dim Fnum As Integer
    
    LogFile = ReadFromINI(GetIniPath1, "GICMISPath", "ErrLogPath")
    
    If Not ChkDirExist(LogFile) Then MkDir LogFile
    
    If Right(LogFile, 1) <> "\" Then LogFile = LogFile & "\"
    LogFile = LogFile & "SYSERR.TXT"
    
    ErrString = Replace(ErrString, vbCrLf & vbCrLf, vbCrLf)
    ErrString = "發生錯誤的系統 : 冠崴客服營運管理系統" & vbCrLf & _
                "發生錯誤的電腦 : " & CreateObject("WScript.NetWork").ComputerName & vbCrLf & _
                "使用者名稱 : " & strUserName & vbCrLf & ErrString
    
    Fnum = FreeFile
    Open LogFile For Append As Fnum
    Print #Fnum, ErrString & vbCrLf & String(66, "-") & vbCrLf
    Close Fnum
End Sub

Public Function msgResult(lngTotalCount As Long, lngErrCount As Long, lngTime As Long)
   Dim strMsg As String
   strMsg = "已完成資料筆數共 : " & lngTotalCount + lngErrCount & "筆," & vbCrLf & vbCrLf & _
            "成功寫入筆數 : " & lngTotalCount & "筆," & vbCrLf & vbCrLf & _
           "問題筆數共 : " & lngErrCount & "筆," & vbCrLf & vbCrLf & _
           "共花費:" & (Timer - lngTime) \ 1 & "秒"
'    strMsg = "已完成資料筆數共" & lngTotalCount & "筆," & vbCrLf & vbCrLf & _
'             "共花費:" & (Timer - lngTime) \ 1 & "秒"
    MsgBox strMsg
    objStorePara.uTime = (Timer - lngTime) \ 1
End Function

'讀取 GICMIS1.INI 檔路徑
Public Function GetIniPath1() As String
  On Error GoTo ChkErr
    Dim strINI As String
    strINI = IIf(Right(Environ("WinDir"), 1) <> "\", Environ("WinDir") & "\", Environ("WinDir")) & "GICMIS1.INI"
    If Not ChkDirExist(strINI) Then
        'strINI = App.Path & "\GICMIS1.INI"
        strINI = objStorePara.Inipath
    End If
    GetIniPath1 = strINI
  Exit Function
ChkErr:
    Call ErrSub("DealFun", "Function GetIniPath1")
End Function

Public Function ChkDirExist(FileName As String) As Boolean
  On Error GoTo ChkErr
    Dim varValue As Variant
    'If Right(FileName, 1) = "\" Then FileName = Left(FileName, Len(FileName) - 1)
    FileName = UCase(FileName)
    varValue = Split(FileName, "\")
    If UCase(Dir(FileName, vbDirectory)) = varValue(UBound(varValue)) Then
        ChkDirExist = True
    Else
        ChkDirExist = False
    End If
  Exit Function
ChkErr:
    Call ErrSub("DealFun", "Function ChkDirExist")
End Function

Public Function ReadFromINI(IniFileName As String, Section As String, Key As String, Optional DecryptFlag As Boolean) As String
  On Error GoTo ChkErr
    If Not ChkDirExist(IniFileName) Then MsgBox "INI之檔名或路徑不正確!!", vbInformation, "訊息..": Exit Function
    Dim s As String, Length As Long
    s = String(1024, 0)
    Length = GetPrivateProfileString(Section, Key, "", s, Len(s), IniFileName)
    ReadFromINI = IIf(DecryptFlag, Decrypt(Left(s, Length)), Left(s, Length))
  Exit Function
ChkErr:
    Call ErrSub("DealFun", "Function ReadFromINI")
End Function

Public Function Decrypt(EncryptionString As String) As String
  On Error GoTo ChkErr
    Decrypt = CreateObject("DevPowerEncrypt.EnCrypt").Decrypt(EncryptionString)
  Exit Function
ChkErr:
    Call ErrSub("DealFun", "Function Decrypt")
    
End Function

Public Function GetString(ByVal strS As Variant, ByVal intlength As Integer, Optional ByVal lngArrow As giArrow = giLeft, Optional ByVal InsertZero As Boolean = False) As String
  On Error GoTo ChkErr
    Dim strCh As String
    strCh = CStr(strS & "")
    If lngArrow = giLeft Then
        strCh = StrConv(LeftB(StrConv(strCh & Space(intlength), vbFromUnicode), intlength), vbUnicode)
    Else
        strCh = StrConv(RightB(StrConv(Space(intlength) & strCh, vbFromUnicode), intlength), vbUnicode)
    End If
    If InsertZero Then
    If Trim(strCh) = "" Then strCh = "0"
        If lngArrow = giLeft Then
            strCh = RTrim(strCh) & String(intlength - Len(RTrim(strCh)), "0")
        Else
            strCh = String(intlength - Len(LTrim(strCh)), "0") & LTrim(strCh)
        End If
    End If
    strCh = Replace(strCh, Chr(0), "")
    GetString = strCh
  Exit Function
ChkErr:
    Call ErrSub("DealFun", "GetString")
End Function

Public Function GetFieldValue(RS As ADODB.Recordset, strFieldName As String)
    On Error GoTo ChkErr
    Dim varField As Variant
        varField = RS.Fields(strFieldName).Value
        If RS.Fields(strFieldName).Type = 131 Or RS.Fields(strFieldName).Type = 139 Then
            If Len(RS.Fields(strFieldName).Value & "") > 0 Then
                varField = Val(RS.Fields(strFieldName).Value)
            End If
        End If
        GetFieldValue = varField
    Exit Function
ChkErr:
    ErrSub "DealFun", "GetFiledValue"
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
'    Call ErrSub("Sys_Lib", "ChkDTok")
End Function

Public Function GetRS(ByRef RS As ADODB.Recordset, _
                      strsql As String, _
                      Optional Conn As ADODB.Connection, _
                      Optional CursorLocation As CursorLocationEnum = adUseClient, _
                      Optional CursorType As CursorTypeEnum = adOpenForwardOnly, _
                      Optional LockType As LockTypeEnum = adLockReadOnly, _
                      Optional ErrMessage As String, _
                      Optional Subroutune As String, _
                      Optional DirectUnload As Boolean = False, _
                      Optional MaxRecords As String) As Boolean
  On Error GoTo ChkErr
    If RS.State = adStateOpen Then RS.Close
    RS.CursorLocation = CursorLocation
    If Val(MaxRecords) > 0 Then RS.MaxRecords = Val(MaxRecords)
    'If Conn Is Nothing Then Set Conn = gcnGi
    RS.Open strsql, Conn, CursorType, LockType
    GetRS = True
  Exit Function
ChkErr:
    GetRS = False
    MsgBox "開啟 " & ErrMessage & " 時發生錯誤: " & Err.Description & " 這個程式即將關閉!" & vbCrLf & strsql, vbExclamation, "警告..."
    On Error Resume Next
    If DirectUnload Then Unload Screen.ActiveForm
End Function

Public Function GetRsValue(strsql As String, Optional cnnObj As ADODB.Connection, Optional MsgNoData As String) As Variant
  On Error GoTo ChkErr
    Dim RS As New ADODB.Recordset
    'If cnnObj Is Nothing Then Set cnnObj = gcnGi
'   If rs.State = 1 Then rs.Close
'   rs.CursorLocation = adUseClient
'   rs.Open strSQL, cnnObj, adOpenForwardOnly, adLockReadOnly
    If Not GetRS(RS, strsql, cnnObj) Then Exit Function
    If RS.RecordCount = 0 Then
        If Len(MsgNoData) > 0 Then
            MsgBox MsgNoData, vbExclamation, "提示"
        End If
        GetRsValue = Null
    Else
        If IsNull(RS(0)) Then
            GetRsValue = Null
        Else
            GetRsValue = Trim(Replace(RS(0), Chr(13), ""))
        End If
    End If
  Exit Function
ChkErr:
    ErrSub "DealFun", "GetRsValue"
End Function

Public Function GetUseIndexStr(StrTableName As String, strColumnName As String) As String
    On Error Resume Next
        GetUseIndexStr = UCase(" /*+ Index(" & StrTableName & " I_" & StrTableName & "_" & strColumnName & ") */ ")
End Function

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
            If VarType(varValue) = vbNull Or VarType(varValue) = vbEmpty Then
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
    ErrSub "DealFun", "GetNullString"
End Function

Public Function GetBillNo_Old(strBillNO As String) As String
  On Error Resume Next
    Dim strYM As String
    
    '將西元年改成民國年(yymm)
    strYM = Format(Format(Left(strBillNO, 6) & "01", "####/##/##"), "EEMM")
    GetBillNo_Old = strYM & Mid(strBillNO, 7, 1) & Right(strBillNO, 6)
  Exit Function
End Function

Public Function GetBillNo_New(strBillNO As String) As String
    On Error Resume Next
    Dim strType As String
    Dim strYM As String
        Select Case Mid(strBillNO, 7, 2)
            Case "BC"
                strType = 1
            Case "TC"
                strType = 2
            Case "BI"
                strType = 3
            Case "TI"
                strType = 4
        End Select
        strYM = Format(Format(Left(strBillNO, 6) & "01", "####/##/##"), "EEMM")
        If Val(Mid(strBillNO, 5, 2)) < 10 Then
            strYM = Left(strYM, 2) & Right(strYM, 1)
        Else
            strYM = Left(strYM, 2) & Chr(Asc("A") + Mid(strYM, 3, 2) - 10)
        End If
        '改成民國年11 碼單號
        GetBillNo_New = strYM & strType & Right(strBillNO, 7)
End Function

Public Function FileDialog(Title As String, _
                           Filter As String, _
                           Optional DefaultDir As String = "", _
                           Optional FilterIndex As Integer = 1, _
                           Optional FileName As String = "", _
                           Optional OpenMode As Integer = 1, _
                           Optional DefaultExt As String = "") As String
    
  On Error GoTo ChkErr
  
    Dim ComDlg As Object
    Dim Result As String
    Set ComDlg = CreateObject("Common.Dialog")
    Result = ComDlg.FileDialog(Title, Filter, DefaultDir, FilterIndex, FileName, DefaultExt, OpenMode)
    Set ComDlg = Nothing
    FileDialog = Result
    
  Exit Function
ChkErr:
    ErrSub "DealFun", "FileDialog"
End Function

Public Function FolderDialog(Title As String) As String
  On Error GoTo ChkErr
    Dim ComDlg As Object
    Dim Result As String
    Set ComDlg = CreateObject("Common.Dialog")
    Result = ComDlg.FolderDialog(Title)
    Set ComDlg = Nothing
    FolderDialog = Result
  Exit Function
ChkErr:
    ErrSub "DealFun", "FolderDialog"
End Function

Public Function RPxx(strValue As String) As String
  On Error Resume Next
    strValue = Trim(Replace(strValue, Chr(13), "", 1, , vbTextCompare))
    strValue = Trim(Replace(strValue, Chr(10), "", 1, , vbTextCompare))
    RPxx = Trim(Replace(strValue, Chr(9), "", 1, , vbTextCompare))
End Function

Public Function GetInvoiceNo2(StrTableName As String, ByVal cn As ADODB.Connection) As String
  On Error GoTo ChkErr
    Dim rsInv As New ADODB.Recordset
    Dim strSeq As String
    strSeq = "S_" & StrTableName
    'YYYYMMDD00000000
    If GetRS(rsInv, "SELECT '" & Right(Format(Date, "YYYYMM"), 4) & _
        "' || Ltrim(To_Char(" & TableOwnerName & strSeq & ".NextVal, '0999999')) FROM Dual", cn) = False Then Exit Function
    GetInvoiceNo2 = rsInv(0).Value
    Set rsInv = Nothing
  Exit Function
ChkErr:
    ErrSub "DealFun", "GetInvoiceNo2"
End Function

Public Function IsPayKindOK(ByRef rsSource As ADODB.Recordset, _
    ByVal strRealStopDate As String, ByVal aChkPayKind As ChkPayKind, _
    Optional ByVal aMedia As Integer = 1) As Boolean
  On Error GoTo ChkErr
    Dim aQry As String
    Dim aPayKind As Integer
    Dim aRsTmp As New ADODB.Recordset
    Dim aBillField As String
    Select Case aMedia
        Case 0
            aBillField = " BILLNO "
        Case 1
            aBillField = " MEDIABILLNO "
        Case Else
            aBillField = " CitibankATM "
    End Select
    
    Select Case aChkPayKind
        Case giSingle
            aQry = "SELECT MAX(NVL(A.PAYKIND,0)) PAYKIND, " & _
                        strMinRealStopDate & _
                        " FROM " & TableOwnerName & "SO033 A WHERE " & _
                        aBillField & "='" & rsSource("BILLNO") & "'"
            Set aRsTmp = objStorePara.Connection.Execute(aQry)
            aPayKind = Val(aRsTmp("PAYKIND") & "")
            If aPayKind = 0 Then
                IsPayKindOK = True
                GoTo lEnd
                
            Else
                If Val(Format(aRsTmp("REALSTOPDATE") & "", "YYYYMMDD")) < Val(strRealStopDate & "") Then
                    IsPayKindOK = False
                Else
                    IsPayKindOK = True
                End If
            End If
        Case giAll
            If Val(rsSource("PayKind") & "") = 0 Then
                IsPayKindOK = True
                GoTo lEnd
            End If
            If Val(Format(rsSource("REALSTOPDATE") & "", "YYYYMMDD")) < Val(strRealStopDate & "") Then
                IsPayKindOK = False
            Else
                IsPayKindOK = True
            End If
    End Select
lEnd:
    On Error Resume Next
    If Not aRsTmp Is Nothing Then
        If aRsTmp.State = adStateOpen Then
            aRsTmp.Close
        End If
        Set aRsTmp = Nothing
    End If
    'Call CloseRecordset(aRsTmp)
    Exit Function
ChkErr:
  Call ErrSub("DealFun", "IsPayKindOK")
End Function


