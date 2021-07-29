Attribute VB_Name = "CDRGlobleModule"
Option Explicit

Public strSysLog As String
Public strCDRIniFile As String
Public strVODCmdSeqNo As String
Public lngCmdTimeOut As Long
Public strRunTime As String
Public strCreditTypeCode As String
Public strCreditTypeName As String
Public blnExitApp As Boolean
Private Const moduleName As String = "CDRGlobleModule"
Public FMaxIncurredDate As String
Public fSmsConnectString As String
Public fSmsErrCount As Long
Public fSmsTels As String
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Sub ErrCDRSub(FormName As String, ProcedureName As String)
    Dim ErrNum As Long
    Dim ErrDesc As String
    Dim ErrSrc As String
    Dim aryErr As Variant
    Dim strErr As String
    ErrNum = Err.Number
    ErrDesc = Err.Description
    ErrSrc = Err.Source
    Err.Number = ErrNum
    Err.Description = ErrDesc
    If (ErrDesc = "") And (ErrNum = -2147467259) Then
        ErrDesc = "ORACLE 中斷連線"
    End If
    '#5667 錯誤增加公司別 By Kin 2010/05/27
    aryErr = Array("發生錯誤時間：", "發生錯誤專案 : ", "發生錯誤表單 : ", "發生錯誤程序 : ", "錯誤代碼 : ", "錯誤原因 : ", "錯誤公司 : ")
    
    strErr = aryErr(0) & Now & vbCrLf & _
                aryErr(1) & "csCallExeCDR" & vbCrLf & _
                aryErr(2) & FormName & vbCrLf & _
                aryErr(3) & ProcedureName & vbCrLf & _
                aryErr(4) & CStr(ErrNum) & vbCrLf & _
                aryErr(5) & CStr(ErrDesc) & vbCrLf & _
                aryErr(6) & CStr(gCompCode)
    
    
    On Error Resume Next
    Screen.ActiveForm.Enabled = True
    If ErrNum = -2147467259 Then
        ErrLog strErr
        Exit Sub
    End If
    If IsConnect(gcnGi) Then
        If blnExitApp Then
            Call ErrSub(FormName, ProcedureName)
        Else
            ErrLog strErr
        End If
    Else
        ErrLog aryErr
    End If
    Err.Clear
End Sub
Public Sub WriteErr(ByVal aMsg As String, ByVal afrmName As String, ByVal aProcedureName As String)
    Dim aryErr As Variant
    Dim strErr As String
    If Not IsConnect(gcnGi) Then
        aMsg = "與 Oracle 伺服器斷線"
    End If
    On Error Resume Next
    '#5667 錯誤訊息增加公司別 By Kin 2010/05/27
    aryErr = Array("發生錯誤時間：", "發生錯誤專案 : ", "發生錯誤表單 : ", "發生錯誤程序 : ", "錯誤代碼 : ", "錯誤原因 : ", "錯誤公司 : ")
    strErr = aryErr(0) & Now & vbCrLf & _
            aryErr(1) & "csCallExeCDR" & vbCrLf & _
            aryErr(2) & afrmName & vbCrLf & _
            aryErr(3) & aProcedureName & vbCrLf & _
            aryErr(4) & "-99自訂錯誤" & vbCrLf & _
            aryErr(5) & aMsg & vbCrLf & _
            aryErr(6) & CStr(gCompCode)
    ErrLog strErr
End Sub
Public Function IsSMSConnect(ByRef cn As ADODB.Connection) As Boolean
    IsSMSConnect = False
    On Error GoTo ResumeConnect
    If cn.State = adStateOpen Then
        If cn.Execute("Select SysDate From Dual")(0).Value & "" <> "" Then
            IsSMSConnect = True
            Exit Function
        End If
    End If
ResumeConnect:
    On Error GoTo ChkErr
'    If cn.State <> adStateClosed Then
'        cn.Close
'    End If
    Set cn = Nothing
    Dim objCN As Object
    Set objCN = CreateObject("ADODB.Connection")
    'Set cn = New ADODB.Connection
    objCN.ConnectionString = fSmsConnectString
    objCN.CursorLocation = adUseClient
    objCN.Open
    If objCN.Execute("Select SysDate From Dual")(0).Value & "" <> "" Then
        Set cn = objCN
        IsSMSConnect = True
    Else
        IsSMSConnect = False
        Err.Raise -2147467259
    End If
    Exit Function
ChkErr:
  Err.Raise -2147467259
End Function


Public Function IsConnect(ByRef cn As ADODB.Connection) As Boolean
    IsConnect = False
    On Error GoTo ResumeConnect
    If cn.State = adStateOpen Then
        If cn.Execute("Select SysDate From Dual")(0).Value & "" <> "" Then
            IsConnect = True
            Exit Function
        End If
    End If
ResumeConnect:
    On Error GoTo ChkErr
'    If cn.State <> adStateClosed Then
'        cn.Close
'    End If
    Set cn = Nothing
    Dim objCN As Object
    Set objCN = CreateObject("ADODB.Connection")
    'Set cn = New ADODB.Connection
    objCN.ConnectionString = garyGi(3)
    objCN.CursorLocation = adUseClient
    objCN.Open
    If objCN.Execute("Select SysDate From Dual")(0).Value & "" <> "" Then
        Set cn = objCN
        IsConnect = True
    Else
        IsConnect = False
        Err.Raise -2147467259
    End If
    Exit Function
ChkErr:
  Err.Raise -2147467259
End Function
Public Function GetCDRRs(ByRef rs As ADODB.Recordset, _
                      strSQL As String, _
                      Optional Conn As ADODB.Connection, _
                      Optional CursorLocation As CursorLocationEnum = adUseClient, _
                      Optional CursorType As CursorTypeEnum = adOpenForwardOnly, _
                      Optional LockType As LockTypeEnum = adLockReadOnly, _
                      Optional ErrMessage As String, _
                      Optional Subroutune As String, _
                      Optional DirectUnload As Boolean = False, _
                      Optional MaxRecords As String) As Boolean
  On Err GoTo ChkErr
    If rs.State = adStateOpen Then
        On Error Resume Next
        rs.CancelUpdate
        rs.Close
    End If
    rs.CursorLocation = CursorLocation
    If Val(MaxRecords) > 0 Then rs.MaxRecords = Val(MaxRecords)
    If Conn Is Nothing Then Set Conn = gcnGi
    rs.Open strSQL, Conn, CursorType, LockType
    GetCDRRs = True
    Exit Function
ChkErr:
    GetCDRRs = False
    ErrCDRSub moduleName, "GetCDRRs"
End Function
Public Function GetRightNow() As Date
  On Error GoTo ChkErr
    If Val(gcnGi.Execute("SELECT SYSTIME FROM " & GetOwner & "SO041 WHERE ROWNUM <=1 ORDER BY COMPCODE").GetString(adClipString, , , , 1) & "") = 1 Then
        GetRightNow = gcnGi.Execute("SELECT TO_CHAR(SYSDATE,'YYYY/MM/DD') FROM DUAL").GetString & ""
        GetRightNow = RPxx(CStr(GetRightNow)) & " " & RPxx(gcnGi.Execute("SELECT TO_CHAR(SYSDATE,'HH24:MI:SS') FROM DUAL").GetString & "")
    Else
        GetRightNow = Now
    End If
    Exit Function
ChkErr:
    GetRightNow = Now
End Function
'取出要傳送的起始日期
Public Function GetSendTime() As String
  On Error GoTo ChkErr:
    Dim strTime As String
    Dim ablnInitial As Boolean
    ablnInitial = ReadINIData(strCDRIniFile, CStr(gCompCode), "Initial") = 0
    strTime = Empty
    If ablnInitial Then
        strTime = MaxVodTime
    Else
        strTime = ReadINIData(strCDRIniFile, CStr(gCompCode), "SendCmdTime")
    End If
    If strTime = "" Then
        strTime = ReadINIData(strCDRIniFile, CStr(gCompCode), "SendCmdTime")
    End If
    GetSendTime = strTime
    
    Exit Function
ChkErr:
  GetSendTime = ""
End Function
'取出SO033VO時間最大的值
Private Function MaxVodTime() As String
  On Error GoTo ChkErr
    Dim rsTime As New ADODB.Recordset
    Dim strQry As String
    strQry = "Select Max(incurredDate) RDATE From " & GetOwner & "SO033VOD "
    If Not GetCDRRs(rsTime, strQry, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
    If rsTime.EOF Then
        MaxVodTime = ""
    Else
        If IsDate(rsTime(0) & "") Then
            MaxVodTime = Format(rsTime(0) & "", "yyyymmddhhmmss")
        Else
            MaxVodTime = ""
        End If
    End If
    On Error Resume Next
    Call CloseRecordset(rsTime)
    Exit Function
ChkErr:
    MaxVodTime = ""
    ErrCDRSub moduleName, "MaxVodTime"
End Function
'讀取INI檔的資料
Public Function ReadINIData(IniFileName As String, Section As String, Key As String, Optional DecryptFlag As Boolean) As String
  On Error GoTo ChkErr
    Dim S As String, Length As Long
    S = String(1024, 0)
    Length = GetPrivateProfileString(Section, Key, "", S, Len(S), IniFileName)
    ReadINIData = IIf(DecryptFlag, Decrypt(Left(S, Length)), Left(S, Length))
  Exit Function
ChkErr:
    Call ErrCDRSub(moduleName, "ReadINIData")
End Function
'寫入資料至Ini
Public Sub WriteCDRParaIni()
  On Error Resume Next
  WritePrivateProfileString CStr(gCompCode), "Initial", ByVal CStr("1"), strCDRIniFile
  WritePrivateProfileString CStr(gCompCode), "SendCmdTime", _
                            ByVal CStr(Format(FMaxIncurredDate, "yyyymmddhhmmss")), _
                            strCDRIniFile
End Sub

