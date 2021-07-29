Attribute VB_Name = "DealFun"
Option Explicit
Private Const FormName = "DealFun"

Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public objStorePara As clsStoreParameter
Public objInterface As clsInterface
Public fso As New FileSystemObject
Public file(2) As TextStream
Public strUserName As String
Public objAction As Object
Public rsTmp1 As New ADODB.Recordset
Public Const gMsgIsDataOK = "此為必要欄位,須有值 !! "
Public Const gimsgPrompt = "提示!"
Public Const gimsgWarning = "警告"
Public blnNodata As Boolean           '判斷是否有資料
Public gcnGi As New ADODB.Connection
Public strOwnerName As String
Public blnChkPayKind As Boolean
Public isCrossCustCombine  As Boolean
Public sqlTmpViewName As String
Public Const strMinRealStopDate As String = " MIN(DECODE(NVL(A.PAYKIND,0),1, " & _
                "DECODE(RealStopDate,NULL,TO_DATE('10000101','YYYYMMDD'),REALSTOPDATE)," & _
                "TO_DATE('10000101','YYYYMMDD'))) REALSTOPDATE "
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
Public Function GetDTString(dtm As String) As String
  On Error GoTo ChkErr
    'GetDTString = Format(dtm, gDTfmt)
    GetDTString = AnnoDominiToTaiwanCalendar(dtm, GiYear) & Format(dtm, "/MM/DD HH:MM:SS")
  Exit Function
ChkErr:
    Call ErrSub("DealFun", "Function GetDTString")
End Function
Public Function AnnoDominiToTaiwanCalendar(adDateTime As String, Optional dtType As GiDTtpye) As String
  On Error GoTo ChkErr
    AnnoDominiToTaiwanCalendar = adDateTime
    '----------------------------------------------------------------------------------------------
    If adDateTime <> Empty Then
        '----------------------------------------------------------------------------------------------
        Dim twCal As String
        Dim adYear As String
        Dim mthDay As String
        Dim mth As String
        Dim ds As String
        '----------------------------------------------------------------------------------------------
        adYear = Left(adDateTime, 4)
        adYear = CStr(Int(adYear) - 1911)
        '----------------------------------------------------------------------------------------------
        If Len(adDateTime) > 10 Then
            mthDay = Format(adDateTime, "/mm/dd HH:MM:SS")
        Else
            mthDay = Format(adDateTime, "/mm/dd")
        End If
        '----------------------------------------------------------------------------------------------
        'mthDay = Mid(adDateTime, 5)
        '----------------------------------------------------------------------------------------------
        mth = Format(adDateTime, "mm")
        ds = Format(adDateTime, "dd")
        '----------------------------------------------------------------------------------------------
        If IsMissing(dtType) Or dtType = 0 Then
            twCal = adYear & mthDay
            'twCal = Format(adDateTime, "ee/mm/dd")
        Else
            Select Case dtType
                Case GiDate
                    twCal = adYear & mthDay
                    'twCal = Format(adDateTime, "ee/mm/dd")
                Case GiYM
                    twCal = adYear & "/" & mth
                    'twCal = Format(adDateTime, "ee/mm")
                Case GiYear
                    twCal = adYear
                    'twCal = Format(adDateTime, "ee")
            End Select
        End If
        '----------------------------------------------------------------------------------------------
        AnnoDominiToTaiwanCalendar = twCal
        '----------------------------------------------------------------------------------------------
    End If
    '----------------------------------------------------------------------------------------------
  Exit Function
ChkErr:
    ErrSub "DealFun", "AnnoDominiToTaiwanCalendar"
End Function
Public Function GetChkPayKind() As Boolean
    On Error Resume Next
    Dim aRet As Boolean
    aRet = False
    aRet = Val(gcnGi.Execute("SELECT NVL(PayNowChkStopDate,0) FROM " & objStorePara.uGetOwner & "SO041")(0)) = 1
    GetChkPayKind = aRet
    
End Function
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

Public Sub CloseFS()
    On Error Resume Next
    file(0).Close
    file(1).Close
    file(2).Close
    Set fso = Nothing
End Sub

Public Function WriteLineData(ByVal vData As String, ByVal intCount As Integer)
On Error GoTo ChkErr
    Call file(intCount).WriteLine(vData)
Exit Function
ChkErr:
    Call ErrSub("DealFun", "WriteLineData")
End Function
Public Sub WriteSO3271Err(ByVal aCustid As String, ByVal aBillNo As String, ByVal aRealStopDate As String)
  On Error GoTo ChkErr
    Dim strErrMsg  As String
    Dim strPayKind As String
    Dim aSQL As String
    aSQL = "SELECT Description FROM " & objStorePara.uGetOwner & "CD112 WHERE CODENO=1"
    strPayKind = GetRsValue(aSQL, gcnGi) & ""
    strErrMsg = "客編：" & aCustid & "、單據編號：" & aBillNo & _
                            " 、原因: 收視截止日不正確-->收視截止日大於扣款處理日、收視截止日：" & aRealStopDate & "、繳付類別: " & strPayKind
    WriteLineData strErrMsg, 2
    Exit Sub
ChkErr:
    Call ErrSub("DealFun", "WriteSO3271Err")
End Sub
Public Function GetPayKindCustId(ByRef rsSource As ADODB.Recordset, ByVal aMedia As Integer) As String
  On Error GoTo ChkErr
    Dim aQry As String
    Dim aBillField As String
    Select Case aMedia
        Case 0
            aBillField = " BILLNO  "
        Case 1
            aBillField = " MEDIABILLNO  "
        Case Else
            aBillField = " CitibankATM  "
    End Select
    aQry = "SELECT CUSTID FROM " & objStorePara.uGetOwner & "SO033 WHERE " & _
        aBillField & "='" & rsSource("BILLNO") & "' AND ROWNUM =1 "
    GetPayKindCustId = GetRsValue(aQry, gcnGi) & ""
    
    Exit Function
ChkErr:
    Call ErrSub("DealFun", "GetPayKindCustId")
End Function
Public Function GetLiteraUpdateSQL(ByVal aMedia As Integer, ByVal aTables As String, _
    ByVal aWhere As String, ByVal aChoose As String, ByVal blnOutZero As Boolean, _
    ByVal aBillNo As String) As String
  On Error GoTo ChkErr
    Dim aRet As String
    Dim aBillField As String
    Dim aBillType As String
    Select Case aMedia
        Case 0
            aBillField = " A.BILLNO BILLNO "
            aBillType = " A.BILLNO "
        Case 1
            aBillField = " A.MEDIABILLNO BILLNO "
            aBillType = " A.MEDIABILLNO "
        Case Else
            aBillField = " A.CitibankATM BILLNO "
            aBillType = " A.CitibankATM "
    End Select
    aRet = "Select A.UCCode,A.UCName,A.UpdEn,A.UpdTime,Nvl(A.PayKind,0) PayKind, " & aBillField & " From " & _
                objStorePara.uGetOwner & "SO033 A" & aTables & _
              " Where " & aWhere & _
              IIf(Len(aWhere) = 0, "", " And ") & aChoose & " And " & aBillType & "='" & aBillNo & "'"
    If Not blnOutZero Then
'            aRet = "Select A.UCCode,A.UCName,A.UpdEn,A.UpdTime,Nvl(A.PayKind,0) PayKind, " & aBillField & " From " & _
'                objStorePara.uGetOwner & "SO033 A" & aTables & _
'              " Where " & aWhere & _
'              IIf(Len(aWhere) = 0, "", " And ") & aChoose & _
'               " And " & aBillType & "='" & aBillNo & "' " & _
'              " Group By " & Replace(aBillField, " BILLNO", "") & _
'              " Having Sum(A.ShouldAmt)>0 "

        aRet = "Select " & aBillField & " From " & _
                objStorePara.uGetOwner & "SO033 A" & aTables & _
              " Where " & aWhere & _
              IIf(Len(aWhere) = 0, "", " And ") & aChoose & _
               " And " & aBillType & "='" & aBillNo & "' " & _
              " Group By " & Replace(aBillField, " BILLNO", "") & _
              " Having Sum(A.ShouldAmt)>0 "

        aRet = "Select A.UCCode,A.UCName,A.UpdEn,A.UpdTime,Nvl(A.PayKind,0) PayKind, " & aBillField & " From " & _
                objStorePara.uGetOwner & "SO033 A" & aTables & _
              " Where " & aWhere & _
              IIf(Len(aWhere) = 0, "", " And ") & aChoose & _
              IIf(Len(aWhere) = 0, "", " And ") & Replace(aBillField, " BILLNO", "") & " IN (" & aRet & ")"
                
    End If
    GetLiteraUpdateSQL = aRet
    Exit Function
ChkErr:
  Call ErrSub("DealFun", "GetLiteraUpdateSQL")
End Function


Public Function GetUpdateSQL(ByVal aMedia As Integer, ByVal aTables As String, _
    ByVal aWhere As String, ByVal aChoose As String, ByVal blnOutZero As Boolean, _
    Optional filterCrossCustCombine As String = "") As String
  On Error GoTo ChkErr
    Dim aRet As String
    Dim aBillField As String
    Dim aSelectBillField As String
    Dim strCrossCustCombineSQL As String
    
    Select Case aMedia
        Case 0
            aBillField = " A.BILLNO BILLNO "
            aSelectBillField = "A.BillNO"
        Case 1
            aBillField = " A.MEDIABILLNO BILLNO "
            aSelectBillField = "A.MEDIABILLNO"
        Case Else
            aBillField = " A.CitibankATM BILLNO "
            aSelectBillField = "A.CitibankATM"
    End Select
    If Len(filterCrossCustCombine) > 0 Then
        aChoose = aChoose & " And " & aSelectBillField & " Not In (Select BillNo From ( " & filterCrossCustCombine & ") A Where ( Custid < 0  Or CustIdExistFlag = 0 ) ) "
    End If
    aRet = "Select A.UCCode,A.UCName,A.UpdEn,A.UpdTime,Nvl(A.PayKind,0) PayKind, " & aBillField & " From " & _
                objStorePara.uGetOwner & "SO033 A" & aTables & _
              " Where " & aWhere & _
              IIf(Len(aWhere) = 0, "", " And ") & aChoose
    If Not blnOutZero Then
    
        aRet = "Select " & aBillField & " From " & _
                objStorePara.uGetOwner & "SO033 A" & aTables & _
              " Where " & aWhere & _
              IIf(Len(aWhere) = 0, "", " And ") & aChoose & _
              " Group By " & Replace(aBillField, " BILLNO", "") & _
              " Having Sum(A.ShouldAmt)>0 "
              
        aRet = "Select A.UCCode,A.UCName,A.UpdEn,A.UpdTime,Nvl(A.PayKind,0) PayKind, " & aBillField & " From " & _
                objStorePara.uGetOwner & "SO033 A" & aTables & _
              " Where " & aWhere & _
              IIf(Len(aWhere) = 0, "", " And ") & aChoose & _
              IIf(Len(aWhere) = 0, "", " And ") & Replace(aBillField, " BILLNO", "") & " IN (" & aRet & ")"
              
              
'        aRet = "Select A.UCCode,A.UCName,A.UpdEn,A.UpdTime,Nvl(A.PayKind,0) PayKind , " & aBillField & " From " & _
'                objStorePara.uGetOwner & "SO033 A " & _
'                " WHERE " & Replace(aBillField, " BILLNO", "") & " IN (" & aRet & ")"

'        aRet = "Select A.UCCode,A.UCName,A.UpdEn,A.UpdTime,Nvl(A.PayKind,0) PayKind, " & aBillField & " From " & _
'                objStorePara.uGetOwner & "SO033 A" & aTables & _
'              " Where " & aWhere & _
'              IIf(Len(aWhere) = 0, "", " And ") & aChoose & _
'              " Group By A.UCCode,A.UCName,A.UpdEn,A.UpdTime,A.PayKind, " & Replace(aBillField, " BILLNO", "") & _
'              " Having Sum(A.ShouldAmt)>0 "
    End If
    GetUpdateSQL = aRet
    Exit Function
ChkErr:
  Call ErrSub("DealFun", "GetUpdateSQL")
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
                        " FROM " & objStorePara.uGetOwner & "SO033 A WHERE " & _
                        aBillField & "='" & rsSource("BILLNO") & "'"
            Set aRsTmp = gcnGi.Execute(aQry)
            aPayKind = Val(aRsTmp("PAYKIND") & "")
            If aPayKind = 0 Then
                IsPayKindOK = True
                GoTo lEnd
            Else
                If Val(Format(aRsTmp("REALSTOPDATE") & "", "YYYYMMDD")) <= Val(strRealStopDate & "") Then
                    IsPayKindOK = True
                Else
                    IsPayKindOK = False
                End If
            End If
        Case giAll
            If Val(rsSource("PayKind") & "") = 0 Then
                IsPayKindOK = True
                GoTo lEnd
            End If
            If Val(Format(rsSource("REALSTOPDATE") & "", "YYYYMMDD")) <= Val(strRealStopDate & "") Then
                IsPayKindOK = True
            Else
                IsPayKindOK = False
            End If
    End Select
lEnd:
    On Error Resume Next
    Call CloseRecordset(aRsTmp)
    Exit Function
ChkErr:
  Call ErrSub("DealFun", "IsPayKindOK")
End Function
Public Function OpenFile(ByVal FilePath As String, ByVal intCount As Integer, Optional OverWrite As Boolean = True) As Boolean
On Error GoTo ChkErr
    OpenFile = False
    
    If Not fso.FolderExists(Left(FilePath, InStrRev(FilePath, "\"))) Then
        MsgBox "路徑不正確!", vbExclamation, "警告..."
        OpenFile = False
        Exit Function
    End If
    
    If fso.FileExists(FilePath) Then
        If fso.GetFile(FilePath).Attributes = ReadOnly Then
           MsgBox "使用者沒有 " & FilePath & " 的使用權限", vbExclamation, "警告"
           OpenFile = False
           Exit Function
        End If
    End If
    Set file(intCount) = fso.CreateTextFile(FilePath, OverWrite)
    OpenFile = True
Exit Function
ChkErr:
    Call ErrSub("DealFun", "OpenFile")
End Function

Public Function InstLogDB(ByVal cn As ADODB.Connection, ByVal BankCode As Integer, _
        ByVal TranType As String, ByVal BillNO As String, CustId As Integer, ByVal Amount As Integer, _
        ByVal ChkLog As Integer, ByVal TranTime As String, ByVal Rectype As Integer, _
        ByVal Content As String)
On Error GoTo ChkErr
    Dim strSQL As String
    strSQL = "Insert Into " & strOwnerName & "So037 (BankCode,TranType,BillNo,Custid,Amount,ChkLog,tranTime,RecType,Content) values("
    strSQL = strSQL & BankCode & ",'1','" & BillNO & "'," & CustId & "," & Amount & ",0,to_date('" & TranTime & "','YYYY/MM/DD HH24:MI:SS')," & Rectype & ",'" & Content & "')"
    cn.Execute strSQL
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
    strErr = aryErr(0) + CStr(RightNow) + strCrLf + _
             aryErr(1) + CStr(Err.Source) + strCrLf + _
             aryErr(2) + FormName + strCrLf + _
             aryErr(3) + ProcedureName + strCrLf + _
             aryErr(4) + CStr(Err.Number) + strCrLf + _
             aryErr(5) + CStr(Err.Description)
    'ErrLog strErr
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
    MsgBox "已完成資料筆數共" & lngTotalCount & "筆," & vbCrLf & vbCrLf & _
           "問題筆數共" & lngErrCount & "筆," & vbCrLf & vbCrLf & _
           "共花費:" & (Timer - lngTime) \ 1 & "秒"
    objStorePara.uTime = (Timer - lngTime) \ 1
End Function

'讀取 GICMIS1.INI 檔路徑
Public Function GetIniPath1() As String
  On Error GoTo ChkErr
    Dim strINI As String
'    strINI = IIf(Right(Environ("WinDir"), 1) <> "\", Environ("WinDir") & "\", Environ("WinDir")) & "GICMIS1.INI"
'    If Not ChkDirExist(strINI) Then
'        strINI = objStorePara.Inipath
'    End If
    strINI = objStorePara.Inipath
    GetIniPath1 = strINI
  Exit Function
ChkErr:
    Call ErrSub("DealFun", "Function GetIniPath1")
End Function

Public Function ChkDirExist(FileName As String) As Boolean
  On Error GoTo ChkErr
    Dim varValue As Variant
    If Right(FileName, 1) = "\" Then FileName = Left(FileName, Len(FileName) - 1)
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
    Call ErrSub("DealFun", "GetString")
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

Public Function GetRS(ByRef rs As ADODB.Recordset, _
                      strSQL As String, _
                      Optional Conn As ADODB.Connection, _
                      Optional CursorLocation As CursorLocationEnum = adUseClient, _
                      Optional CursorType As CursorTypeEnum = adOpenForwardOnly, _
                      Optional LockType As LockTypeEnum = adLockReadOnly, _
                      Optional ErrMessage As String, _
                      Optional Subroutune As String, _
                      Optional DirectUnload As Boolean = False, _
                      Optional MaxRecords As String) As Boolean
  On Error GoTo ChkErr
    
    If rs.State = adStateOpen Then rs.Close
    rs.CursorLocation = CursorLocation
    If Val(MaxRecords) > 0 Then rs.MaxRecords = Val(MaxRecords)
    If Conn Is Nothing Then Set Conn = gcnGi
    rs.Open strSQL, Conn, CursorType, LockType
    GetRS = True
  Exit Function
ChkErr:
    GetRS = False
    MsgBox "開啟 " & ErrMessage & " 時發生錯誤: " & Err.Description & " 這個程式即將關閉!" & vbCrLf & strSQL, vbExclamation, "警告..."
    On Error Resume Next
    If DirectUnload Then Unload Screen.ActiveForm
End Function

Public Function GetRsValue(strSQL As String, Optional cnnObj As ADODB.Connection, Optional MsgNoData As String) As Variant
  On Error GoTo ChkErr
    Dim rs As New ADODB.Recordset
    'If cnnObj Is Nothing Then Set cnnObj = gcnGi
'   If rs.State = 1 Then rs.Close
'   rs.CursorLocation = adUseClient
'   rs.Open strSQL, cnnObj, adOpenForwardOnly, adLockReadOnly
    If Not GetRS(rs, strSQL, cnnObj) Then Exit Function
    If rs.RecordCount = 0 Then
        If Len(MsgNoData) > 0 Then
            MsgBox MsgNoData, vbExclamation, "提示"
        End If
        GetRsValue = Null
    Else
        If IsNull(rs(0)) Then
            GetRsValue = Null
        Else
            GetRsValue = Trim(Replace(rs(0), Chr(13), ""))
        End If
    End If
  Exit Function
ChkErr:
    ErrSub "DealFun", "GetRsValue"
End Function

Public Function GetUseIndexStr(StrTableName As String, strColumnName As String) As String
    On Error Resume Next
'        GetUseIndexStr = UCase(" /*+ Index(" & StrTableName & " I_" & StrTableName & "_" & strColumnName & ") */ ")
        GetUseIndexStr = ""
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

Public Function GetBillNo_Old(strBillNo As String) As String
  On Error Resume Next
    Dim strYM As String
    
    '將西元年改成民國年(yymm)
    strYM = Format(Format(Left(strBillNo, 6) & "01", "####/##/##"), "EEMM")
    GetBillNo_Old = strYM & Mid(strBillNo, 7, 1) & Right(strBillNo, 6)
  Exit Function
End Function

Public Function GetBillNo_New(strBillNo As String) As String
    On Error Resume Next
    Dim strType As String
    Dim strYM As String
        Select Case Mid(strBillNo, 7, 2)
            Case "BC"
                strType = 1
            Case "TC"
                strType = 2
            Case "BI"
                strType = 3
            Case "TI"
                strType = 4
        End Select
        strYM = Format(Format(Left(strBillNo, 6) & "01", "####/##/##"), "EEMM")
        If Val(Mid(strBillNo, 5, 2)) < 10 Then
            strYM = Left(strYM, 2) & Right(strYM, 1)
        Else
            strYM = Left(strYM, 2) & Chr(Asc("A") + Mid(strYM, 3, 2) - 10)
        End If
        '改成民國年11 碼單號
        GetBillNo_New = strYM & strType & Right(strBillNo, 7)
End Function

Public Sub CloseRecordset(rs As ADODB.Recordset)
    On Error Resume Next
        rs.CancelUpdate
        If rs.State = adStateOpen Then rs.Close
        Set rs = Nothing
End Sub

Public Function PutIntoSO075(strSQL As String, Conn As ADODB.Connection, ByRef lngBatchNo As Long) As Boolean
    On Error GoTo ChkErr
        If Not ExecuteCommand("INSERT INTO " & strOwnerName & "SO075 (BillNo, CustID, BatchNO,PrtSNO) " & strSQL, Conn) Then Exit Function
        PutIntoSO075 = True
    Exit Function
ChkErr:
    ErrSub FormName, "PutIntoSO075"
End Function

Public Function ExecuteCommand(strSQL As String, Conn As ADODB.Connection, Optional ByRef lngAffected As Long) As Boolean
    On Error GoTo ChkErr
        If Conn Is Nothing Then Exit Function
        Conn.Execute strSQL, lngAffected
        ExecuteCommand = True
    Exit Function
ChkErr:
    ErrSub FormName, "ExecuteCommand, SQL : [" & strSQL & "]"
End Function
'#3583增加strSFString(後端名稱)與Cn(Connect) By Kin 2007/10/26
Public Function GetBankATM(strBillNo As String, lngCustId As Long, strBankCode As String, _
    strShouldDate As String, lngShouldAmt As Long, Optional ByRef clsBankATMNo As Object, _
    Optional ByRef strCorpID As String, Optional ByRef strSFString As String, _
    Optional ByRef cn As ADODB.Connection) As String
    On Error Resume Next
    Dim rsTmp As New ADODB.Recordset
    Dim strShDate As String
        If clsBankATMNo Is Nothing Then
        
            If strBankCode = "" Or strBillNo = "" Or lngCustId = 0 Then Exit Function
            If Not GetRS(rsTmp, "Select FileName2,CorpID From " & strOwnerName & "CD018 Where CodeNo = " & strBankCode, cn) Then Exit Function
            If Len(rsTmp("FileName2") & "") = 0 Or Len(rsTmp("CorpId") & "") = 0 Then MsgBox "銀行代碼檔中無定義轉帳輸入檔名或收件單位代號", vbCritical, "警告": Exit Function
            strCorpID = rsTmp("CorpId") & ""
            Set clsBankATMNo = CreateObject("csBankATMNo.cls" & rsTmp("FileName2"))
            If Err.Number = 429 Then MsgBox "銀行代碼檔中定義之轉帳輸入檔名 : [" & rsTmp("FileName2") & "] 不正確,請確認!!", vbCritical, "警告": Exit Function
            strSFString = "SF_" & rsTmp("FileName2")
            If lngCustId = -1 Then Exit Function
        End If
        strShDate = Format(strShouldDate, "yyyy/mm/dd")
        GetBankATM = clsBankATMNo.GetATMNo(strBillNo, lngCustId, strCorpID, strShDate, lngShouldAmt)
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
            If Not GetRS(rs, "Select " & strParaField & " From " & strOwnerName & strTable & " Where CompCode = " & strCompCode & " And ServiceType " & strSQL) Then Exit Function
            blnFlag = Not rs.EOF
        End If
        If Not blnFlag Then
            '再找CompCode 有的
            If strCompCode <> "" Then
                If Not GetRS(rs, "Select " & strParaField & " From " & strOwnerName & strTable & " Where CompCode = " & strCompCode) Then Exit Function
                blnFlag = Not rs.EOF
            End If
            '最後找只要有這個 參數檔的
            If Not blnFlag Then
                If Not GetRS(rs, "Select " & strParaField & " From " & strOwnerName & strTable) Then Exit Function
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
        If Not GetSystemPara(rs, strCompCode, strServiceType, strTable, strField, False, blnServiceTypeUseLike) Then Exit Function
        GetSystemParaItem = rs(0)
End Function

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
                If Trim(rsSource("MediaBillNo")) & "" = "" Then
                    strMediaBillNo = ""
                    If Not GetMediaBillNo(rsSource("CustId"), strMediaBillNo, cn) Then Exit Function
                    If Not ExecuteCommand("Update " & strOwnerName & "SO033 A Set MediaBillNo = '" & strMediaBillNo & "' Where BillNo = '" & rsSource("BillNo") & "' And CustId = " & rsSource("CustId") & IIf(Len(strChoose33) = 0, "", " And ") & strChoose33, cn, lngAffected) Then Exit Function
                    On Error Resume Next
'                     .Fields("MediaBillNo") = strMediaBillNo
'                    .Update
                Else
                    WriteLineData "客編：" & rsSource("CustId") & " 已產生過媒體單號：" & rsSource("MediaBillNo"), 2
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
'#3583批次更新虛擬帳號CitiBankATM By Kin 2007/10/26
Public Function BatchUpdateCitiBankATM(ByVal strChoose33 As String, _
                                        Optional ByVal strTablName As String = "", _
                                        Optional cn As ADODB.Connection) As Boolean
  On Error GoTo ChkErr
    'Dim rsChkAtm As New ADODB.Recordset
    Dim strCorpID As String, strSFString As String
    Dim strBankId As String, strBillNo As String
    Dim strTmpViewName As String, strViewSQL As String
    Dim strSQL As String, strField As String
    Dim lngCount As Long, lngMedia As Long
    strBankId = Empty
    lngMedia = 0
    strField = Empty
    strBankId = objStorePara.BankId
    strTablName = strOwnerName & "SO033 A" & strTablName
    strSQL = "Select nvl(Para24,0) From " & strOwnerName & "SO043 Where Rownum=1"
    lngMedia = Val(cn.Execute(strSQL)(0))
    If lngMedia = 1 Then
        strField = "MediaBillNo"
    Else
        strField = "BillNo"
    End If
    Dim rsChkAtm As New ADODB.Recordset
    If Not GetRS(rsChkAtm, "SELECT CheckATM,FileName2,CorpID From " & strOwnerName & "CD018 Where CodeNo=" & strBankId, cn) Then GoTo ChkErr
    '******************************************************************************************
        '#3726 改變取用虛擬帳號 By Kin 2008/04/02
        'GetBankATM "20" & strBillNo, -1, strBankId, "", 0, , strCorpID, strSFString, cn
        '3726 測試不OK 多增加判斷後端是否存在 By Kin 2008/04/14
        If Not ChkHaveSFObj(rsChkAtm("FileName2") & "", cn) Then Exit Function
        strSFString = "SF_" & rsChkAtm("FileName2") & ""
        strCorpID = rsChkAtm("CorpID") & ""
    '******************************************************************************************
        '取得View暫存檔名稱
        
        strTmpViewName = GetTmpViewName(cn)
        If strTmpViewName = "" Then Exit Function
        If lngMedia = 1 Then
            strViewSQL = "Create View " & strOwnerName & strTmpViewName & " as " & _
                         " Select MediaBillNo BillNo," & strOwnerName & strSFString & _
                         "('" & strCorpID & "','20'||A.MediaBillNo,Substr(A.MediaBillNo,5)," & _
                         "Sum(A.Shouldamt)) CitiBankATM From " & strTablName & IIf(strChoose33 <> "", " Where " & strChoose33, "") & _
                         " Group By MediaBillNo"
        Else
            strViewSQL = "Create View " & strOwnerName & strTmpViewName & " as " & _
                        " Select BillNo," & strOwnerName & strSFString & "('" & strCorpID & _
                        "',A.BillNo,Substr(A.BillNo,9),Sum(A.Shouldamt)) CitiBankATM From " & _
                        strTablName & IIf(strChoose33 <> "", " Where " & strChoose33, "") & _
                         " Group By BillNo"
        End If
        On Error Resume Next
        If Not ExecuteCommand(strViewSQL, cn, lngCount) Then Exit Function
            If cn.Execute("Select Count(*) CNT From " & strOwnerName & strTmpViewName)(0) <= 0 Then BatchUpdateCitiBankATM = True: Exit Function
            cn.BeginTrans
            
            If Not ExecuteCommand("Update " & strOwnerName & "SO033 A Set CitiBankATM = " & _
                                  "(Select CitiBankATM From " & strOwnerName & strTmpViewName & _
                                  " B Where A." & strField & "= B.BillNo) Where A." & strField & " In (" & _
                                  "Select A." & strField & " From " & strTablName & " Where " & strChoose33 & ")", cn, lngCount) Then cn.RollbackTrans: GoTo ChkErr
                                  
            cn.CommitTrans
        On Error Resume Next
        cn.Execute "Drop View " & strOwnerName & strTmpViewName
    'Else
        BatchUpdateCitiBankATM = True
    'End If
    On Error Resume Next
    Call CloseRecordset(rsChkAtm)
    Exit Function
ChkErr:
    ErrSub FormName, "BatchUpdateCitiBankATM"
End Function
'#3583 取得View的暫存檔名稱 By Kin 2007/10/26
Public Function GetTmpViewName(cn As ADODB.Connection) As String
  On Error Resume Next
    Dim strTmpViewName As String
    Dim strGetOwner As String
    strGetOwner = objStorePara.uGetOwner
    strTmpViewName = "TMP_" & GetRsValue("SELECT trim(To_Char(" & strGetOwner & "S_TMPRPT_ViewName.NextVal, '0999999')) FROM Dual", cn)
    cn.Execute "Drop Table " & strTmpViewName
    cn.Execute "Drop View " & strTmpViewName
    GetTmpViewName = strTmpViewName
End Function
Public Function GetInvoiceNo3(StrTableName As String, Optional Conn As ADODB.Connection) As String
On Error GoTo ChkErr
Dim strSeq As String
Dim strInv

   strSeq = "S_" & StrTableName
   If Conn Is Nothing Then Set Conn = gcnGi
   strInv = GetRsValue("SELECT Ltrim(To_Char(" & strOwnerName & strSeq & ".NextVal, '09999999')) FROM Dual", Conn)
   GetInvoiceNo3 = strInv
   
  Exit Function
ChkErr:
  ErrSub "Sys_Lib", "GetInvoiceNo3"
End Function
Public Function RightNow() As String
  On Error GoTo ChkErr
  
    If Val(gcnGi.Execute("SELECT SYSTIME FROM " & strOwnerName & "SO041 WHERE ROWNUM <=1 ORDER BY COMPCODE").GetString(adClipString, , , , 1) & "") = 1 Then
        RightNow = gcnGi.Execute("SELECT TO_CHAR(SYSDATE,'YYYY/MM/DD') FROM DUAL").GetString & ""
        RightNow = RPxx(RightNow) & " " & RPxx(gcnGi.Execute("SELECT TO_CHAR(SYSDATE,'HH24:MI:SS') FROM DUAL").GetString & "")
    Else
        RightNow = Now
    End If
  Exit Function
ChkErr:
    ErrSub "SysLib", "RightNow"
End Function

Public Function RPxx(strValue As String) As String
  On Error Resume Next
    strValue = Trim(Replace(strValue, Chr(13), "", 1, , vbTextCompare))
    strValue = Trim(Replace(strValue, Chr(10), "", 1, , vbTextCompare))
    RPxx = Trim(Replace(strValue, Chr(9), "", 1, , vbTextCompare))
End Function

Public Function ChkHaveSFObj(strObj As String, cn As ADODB.Connection) As Boolean
    On Error Resume Next
        ChkHaveSFObj = Val(GetRsValue("Select Count(*) From All_OBJECTS where Owner = '" & Replace(strOwnerName, ".", "") & "' And object_type = 'FUNCTION' And Object_Name = '" & UCase("SF_" & strObj) & "'", cn) & "") > 0
        If Not ChkHaveSFObj Then MsgBox "銀行代碼檔中定義之轉帳輸入檔名 : [" & strObj & "] 不正確,請確認!!", vbCritical, gimsgWarning: Exit Function

End Function

