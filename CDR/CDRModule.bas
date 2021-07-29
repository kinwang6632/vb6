Attribute VB_Name = "CDRModule"
Option Explicit
Public intAutoCDR As Integer            '�C�X��������@��
Public intAutoConnect As Integer        '�C�X���_�u�۰ʭ��s
Public strSendCmdTime As String         '�ǰeZ3���_�l�ɶ�
Public lngCmdTimeOut As Long            '�R�OTimeOut�ɶ�
Public lngErrExitCount As Long          '���~�X���n���}
Public strConnectString As String
Public scr As New Scripting.FileSystemObject
'Public FileStream As TextStream
'Public fsError As TextStream
Public rsReadFile As New ADODB.Recordset
Public Const strConnectFileName As String = "Connect.Dat"
Public blnExitApp As Boolean
Private strConnectFile As String
Private strLine As String
Private Const strTestSQL As String = "Select SysDate From Dual"
Private gstrStationCount As String
Private strCDRErrFile As String
Public varOwner() As String
Public strVODCmdSeqNo As String
Private Const moduleName As String = "CDRModule"
Private Const blnChkOwner = False
Private SysPath As String
Private Const aSection As String = "CDRCenter"
Public rsSO555C As New ADODB.Recordset
Public rsSO033Vod As New ADODB.Recordset
Private FCustId As String
Private FFaciSNO As String
Private FFaciSeqNo As String
Private FDeclarantName As String
Private FSO004RowId As String
Public MAXIncurredDate As String
Public strCreditTypeCode As String
Public strCreditTypeName As String
Public strRunTime As String
Public strErrlog As String
Public strLock As String
Public blnShowSecond As Boolean
Public strUpdEn As String
Public strAppTitle As String

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
'Private Function GetCreditType() As Boolean
'  On Error Resume Next
'    Dim rsTmp As New ADODB.Recordset
'    If Not GetCDRRs(rsTmp, "Select CodeNo,Description From " & GetOwner & "CD108" & _
'                      " WHERE REFNO = 3 AND NVL(STOPFLAG,0)=0", gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
'    If Not rsTmp.EOF Then
'        strCreditTypeCode = rsTmp(0) & ""
'        strCreditTypeName = rsTmp(1) & ""
'    End If
'    GetCreditType = True
'    Call CloseRecordset(rsTmp)
'End Function
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
        ErrDesc = "ORACLE ���_�s�u"
    End If
    aryErr = Array("�o�Ϳ��~�ɶ� : ", "�o�Ϳ��~�M�� : ", "�o�Ϳ��~��� : ", "�o�Ϳ��~�{�� : ", "���~�N�X : ", "���~��] : ")
                strErr = aryErr(0) & CStr(Now) & strCrLf & _
                aryErr(1) & CStr(ErrSrc) & strCrLf & _
                aryErr(2) & FormName & strCrLf & _
                aryErr(3) & ProcedureName & strCrLf & _
                aryErr(4) & CStr(ErrNum) & strCrLf & _
                aryErr(5) & CStr(ErrDesc)
    
    
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
        'frmMain.tmrRun.Enabled = False
        'frmMain.tmrConnect.Enabled = True
    End If
    Err.Clear
End Sub
'Public Function SendA3Cmd(ByVal aVodAccountId As String) As Boolean
'  On Error GoTo ChkErr
'    Dim blnRet As Boolean
'    Dim rsVir As New ADODB.Recordset
'    Dim strQry As String
'    Dim obj As Object
'    Dim strRetErrMsg As String
'    Set obj = CreateObject("csVODcontrol.clsExeCommand")
'    strQry = "SELECT A.ROWID,A.* FROM " & GetOwner & "SO004 A " & _
'          " WHERE A.VODACCOUNTID = " & aVodAccountId
'    If Not GetCDRRs(rsVir, strQry, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
'    If rsVir.EOF Then
'        Call WriteErr("�ǰeA3�R�O�䤣��SO004���", moduleName, "SendA3Cmd")
'    End If
'    blnRet = obj.ExeVodCommand(gcnGi, rsVir, "A3", garyGi(), _
'                                , , strRunTime, , , , , , , , , , , , , strRetErrMsg)
'    If Not blnRet Then
'        Call WriteErr(strRetErrMsg, moduleName, "SendA3Cmd")
'        Exit Function
'    End If
'    SendA3Cmd = blnRet
'    Exit Function
'
'ChkErr:
'  Call ErrCDRSub(moduleName, "SendA3Cmd")
'End Function
'�ǰeZ3���R�O
'Public Function SendZ3Cmd(ByVal strBillStartDate As String) As Boolean
'  On Error GoTo ChkErr
'    Dim obj As Object
'    Dim blnRet As Boolean
'    Dim rsVir As New ADODB.Recordset
'    Dim aMsg As String
'    strVODCmdSeqNo = "-1"
'    SetRsField rsVir    '�]�w�@�ӵ�����,�H�K�i�H�ǵ�����x
'    If Not IsDate(strBillStartDate) Then
'        Call WriteErr("�ǰeZ3�R�O��������T�I�ǰe�Ȭ�:" & strBillStartDate, moduleName, "SendZ3Cmd")
'        Exit Function
'    Else
'        strBillStartDate = DateAdd("n", -1, CDate(strBillStartDate))
'    End If
'    Set obj = CreateObject("csVODcontrol.clsExeCommand")
'    blnRet = obj.ExeVodCommand(gcnGi, rsVir, "Z3", garyGi, , , strRunTime, "EPG", _
'                , , , , , strBillStartDate, _
'                GetRightNow, , , , , aMsg, lngCmdTimeOut, strVODCmdSeqNo)
'
'    If Not blnRet Then
'        If aMsg = "" Then aMsg = "�������~"
'        Call WriteErr("�ǰeZ3�R�O���~�I " & aMsg, moduleName, "SendZ3Cmd")
'        Exit Function
'    End If
'    If strVODCmdSeqNo < 0 Then
'        aMsg = "���oSO555 SEQNO���~�I"
'        Call WriteErr(aMsg, moduleName, "SendZ3Cmd")
'    End If
'    SendZ3Cmd = blnRet
'    On Error Resume Next
'    Call CloseRecordset(rsVir)
'    Exit Function
'ChkErr:
'    Set obj = Nothing
'    Call ErrCDRSub(moduleName, "SendZ3Cmd")
'End Function
'Public Function GetSO004Data(ByVal aVodAccountId As String) As Boolean
'  On Error GoTo ChkErr
'    Dim rsSO004 As New ADODB.Recordset
'    If Not GetCDRRs(rsSO004, "SELECT A.RowId,A.* FROM " & GetOwner & "SO004 A " & _
'                    " WHERE VODAccountId=" & aVodAccountId, gcnGi, adUseClient, _
'                    adOpenKeyset, adLockOptimistic) Then Exit Function
'    FCustId = Empty
'    FFaciSNO = Empty
'    FFaciSeqNo = Empty
'    FDeclarantName = Empty
'    FSO004RowId = Empty
'    If Not rsSO004.EOF Then
'        FCustId = rsSO004("CustID") & ""
'        FFaciSNO = rsSO004("FaciSNO") & ""
'        FFaciSeqNo = rsSO004("SeqNo") & ""
'        FDeclarantName = rsSO004("DeclarantName") & ""
'        FSO004RowId = rsSO004("RowId") & ""
'    Else
'        Exit Function
'    End If
'    GetSO004Data = True
'    On Error Resume Next
'    Call CloseRecordset(rsSO004)
'    Exit Function
'ChkErr:
'  Call ErrCDRSub(moduleName, "GetSO004Data")
'End Function
'Public Function GetSO555C(ByVal aCmdSeqNo As String, ByRef rs As ADODB.Recordset) As Boolean
'  On Error GoTo ChkErr
'    Dim strSQL As String
'    Dim aMsg As String
'    If Val(aCmdSeqNo) = 0 Then
'        aMsg = "SEQNO���ŭ�"
'        Call WriteErr(aMsg, moduleName, "GetSO555C")
'        Exit Function
'    End If
'    strSQL = "Select * From " & GetOwner & "SO555C Where SeqNo = " & Val(aCmdSeqNo)
'    If Not GetCDRRs(rs, strSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
'    If Not GetCreditType Then GoTo ChkErr
'    GetSO555C = True
'    Exit Function
'ChkErr:
'  Call ErrCDRSub(moduleName, "GetSO555C")
'End Function
'Public Function InsSO033VOD(ByRef rs As ADODB.Recordset, _
'                            Optional ByRef blnDouble As Boolean) As Boolean
'
'  On Error GoTo ChkErr
'    If Not rs.EOF Then
'        If Not GetCDRRs(rsSO033Vod, "Select * From " & GetOwner & "SO033VOD " & _
'                    " WHERE 1=0", gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
'
'
'        If Not GetSO004Data(rs("VODAccountID") & "") Then Exit Function
'        If gcnGi.Execute("Select Count(CustId) From " & GetOwner & "SO033Vod " & _
'                    " WHERE VODUID = " & rs("VODUID"))(0) <= 0 Then
'            With rsSO033Vod
'                .AddNew
'                .Fields("CompCode") = gCompCode
'                .Fields("CustId") = FCustId
'                .Fields("VODAccountID") = rs("VODAccountID") & ""
'                .Fields("itemUID") = rs("itemUID") & ""
'                .Fields("itemName") = rs("itemName") & ""
'                If IsDate(rs("bincurredDate")) Then
'                    .Fields("incurredDate") = rs("BincurredDate")
'                End If
'                .Fields("Value") = Val(rs("Value") & "")
'                .Fields("frequencyType") = rs("frequencyType") & ""
'                .Fields("itemType") = rs("itemType") & ""
'                .Fields("SeqNo") = Get_SO033VOD_Seq
'                .Fields("VODUID") = rs("VODUID")
'                .Update
'            End With
'        Else
'            InsSO033VOD = False
'            blnDouble = True
'            Exit Function
'        End If
'
'    End If
'    InsSO033VOD = True
'    Exit Function
'ChkErr:
'  Call ErrCDRSub(moduleName, "InsSO033Vod")
'End Function
'Private Function Get_SO182_Seq() As String
'  On Error GoTo ChkErr
'    Get_SO182_Seq = RPxx(gcnGi.Execute("SELECT " & GetOwner & "S_SO182_SEQNO.NEXTVAL FROM DUAL").GetString & "")
'    Exit Function
'ChkErr:
'  Call ErrSub(moduleName, "Get_SO182_Seq")
'End Function
'Private Function Get_SO033VOD_Seq() As String
'On Error GoTo ChkErr
'    Get_SO033VOD_Seq = RPxx(gcnGi.Execute("SELECT " & GetOwner & "S_SO033VOD_SEQNO.NEXTVAL FROM DUAL").GetString & "")
'    Get_SO033VOD_Seq = Format(GetRightNow, "yyyymmdd") & _
'                        Right("000" & FCustId, 3) & _
'                        Format(Get_SO033VOD_Seq, "000000000")
'  Exit Function
'ChkErr:
'    Call ErrSub(moduleName, "Get_SO033VOD_Seq")
'End Function
'Public Sub SetRsField(ByRef rs As ADODB.Recordset)
'  On Error GoTo ChkErr
'    With rs
'        If .State = adStateOpen Then .Close
'        .CursorLocation = adUseClient
'        .CursorType = adOpenKeyset
'        .LockType = adLockOptimistic
'        .Fields.Append "CustId", adBSTR, 10, adFldIsNullable
'        .Fields.Append "FaciSNo", adBSTR, 10, adFldIsNullable
'        .Fields.Append "SmartCardNo", adBSTR, 15, adFldIsNullable
'        .Fields.Append "VODAccountId", adBSTR, 20, adFldIsNullable
'        .Fields.Append "CompCode", adBSTR, 10, adFldIsNullable
'    End With
'    rs.Open
'    rs.AddNew
'    rs("CompCode") = gCompCode
'    rs.Update
'    Exit Sub
'ChkErr:
'  Call ErrCDRSub(moduleName, "SetRsField")
'End Sub

'Public Function InsSO182(ByVal aMinusType As Integer, ByVal aValue As String, _
'                            ByRef rs033VOD As ADODB.Recordset) As Boolean
'  On Error GoTo ChkErr:
'    Dim strSQL As String
'    Dim rsSO182 As New ADODB.Recordset
'    If Not GetCDRRs(rsSO182, "Select * From " & GetOwner & "SO182 Where 1=0", _
'                gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
'    With rsSO182
'        .AddNew
'        .Fields("CompCode") = gCompCode
'        .Fields("ServiceType") = "D"
'        .Fields("RecTime") = rs033VOD("IncurredDate")
'        .Fields("RecEn") = garyGi(0)
'        .Fields("FaciSeqNo") = FFaciSeqNo
'        .Fields("UpdEn") = garyGi(0)
'        .Fields("UpdTime") = Format(strRunTime, gDTfmt)
'        .Fields("MinusType") = aMinusType
'        .Fields("CreditTypeCode") = strCreditTypeCode
'        .Fields("CreditTypeName") = strCreditTypeName
'        .Fields("MinusCredit") = aValue
'        .Fields("VODAccountID") = rs033VOD("VODACCOUNTID")
'        .Fields("CreditSeqNo") = rs033VOD("SEQNO")
'        .Fields("SeqNo") = Get_SO182_Seq
'        .Fields("STOPFLAG") = 0
'        .Update
'    End With
'    InsSO182 = True
'    On Error Resume Next
'    Call CloseRecordset(rsSO182)
'    Exit Function
'ChkErr:
'  Call ErrCDRSub(moduleName, "InsSO182")
'End Function

'Public Function UpdSO004G(ByVal aVodAccount As String, _
'                            ByVal aValue As String, _
'                            Optional ByVal blnSendA3Cmd As Boolean = True, _
'                            Optional ByVal blnAddSO182 As Boolean = True, _
'                            Optional ByRef rs033VOD As ADODB.Recordset = Nothing) As Boolean
' On Error GoTo ChkErr
'    Dim strQry As String
'    Dim rsSO004G As New ADODB.Recordset
'    Dim lngRetValue As Long
'    Dim lngUpdValue As Long
'    Dim lngMinusValue As Long
'    If (blnAddSO182) And (rsSO033Vod Is Nothing) Then
'        Exit Function
'    End If
'    strQry = "SELECT * FROM " & GetOwner & "SO004G " & _
'                " WHERE VODAccountID='" & aVodAccount & "'"
'    If Not GetCDRRs(rsSO004G, strQry, gcnGi, adUseClient, _
'                    adOpenKeyset, adLockOptimistic) Then Exit Function
'
'    If Not rsSO004G.EOF Then
'        '�p�G�x�s�I�Ʀ���,�������x�s�I��
'        If Val(rsSO004G("PrePay") & "") > 0 Then
'            lngRetValue = 0
'            lngUpdValue = CalculPoint(Val(rsSO004G("PrePay")), aValue, lngRetValue)
'            lngMinusValue = Val(rsSO004G("PrePay")) - lngUpdValue   '�쥻����-�nUpd����=�����h���I(�n�s�W��SO182��)
'            rsSO004G("PrePay").Value = lngUpdValue
'            If blnAddSO182 Then '�s�W�@�����ںطЬ� "�x��" ����SO182
'                If Not InsSO182(1, lngMinusValue, rs033VOD) Then Exit Function
'            End If
'        Else
'            lngRetValue = aValue
'        End If
'        '�p�G�ذe�I�Ʀ���,�A�����ذe�I��
'        If Val(rsSO004G("Present") & "") > 0 Then
'            If (lngRetValue > 0) Then
'                aValue = lngRetValue
'                lngRetValue = 0
'                lngMinusValue = 0
'                lngUpdValue = CalculPoint(Val(rsSO004G("Present") & ""), aValue, lngRetValue)
'                lngMinusValue = rsSO004G("Present") - lngUpdValue
'                rsSO004G("Present") = lngUpdValue
'                If blnAddSO182 Then '�s�W�@�����ں������ذe����SO182
'                    If Not InsSO182(2, lngMinusValue, rs033VOD) Then Exit Function
'                End If
'            End If
'        Else
'            lngRetValue = aValue
'        End If
'        '�������n�[�b���I�I��
'        If lngRetValue > 0 Then
'            lngMinusValue = Abs(lngRetValue)
'            rsSO004G("NoPayCredit") = Val(rsSO004G("NoPayCredit")) + Abs(lngRetValue)
'            If blnAddSO182 Then '�s�W�@�����ں������w�ɪ���SO182
'                If Not InsSO182(3, lngMinusValue, rs033VOD) Then Exit Function
'            End If
'        End If
'        '���I�O�I��>=�w���I�� �n�ǰeA3�R�O������x
'        If blnSendA3Cmd Then
'            If Val(rsSO004G("NoPayCredit") & "") >= Val(rsSO004G("VODCredit") & "") Then
'                If Not SendA3Cmd(aVodAccount) Then Exit Function
'            End If
'        End If
'        rsSO004G.Update
'    End If
'    UpdSO004G = True
'    Exit Function
'ChkErr:
'  Call ErrCDRSub(moduleName, "UpdSO004G")
'End Function
''�^�ǭnUpd SO004G���I�ƭ�
''aPointValue(SO004G����)
''aValue(�n�����I�ƭ�)
''aRetValue���������I��
'Private Function CalculPoint(ByVal aPointValue As Long, _
'                            ByVal aValue As Long, ByRef aRetValue As Long) As Long
'    Dim lngAnsValue As Long
'    lngAnsValue = aPointValue - aValue
'    If lngAnsValue >= 0 Then
'        aRetValue = 0
'        CalculPoint = lngAnsValue
'    Else
'        aRetValue = Abs(lngAnsValue)
'        CalculPoint = 0
'    End If
'End Function
'�]�w�U���q�O�POwner
Public Sub SetPubOwner(ByVal aIndex As Integer)
  On Error Resume Next
    Dim varAry As Variant
    varAry = Split(varOwner(aIndex), ",")
    gCompCode = varAry(0)
    garyGi(9) = gCompCode
    garyGi(16) = varAry(1)
End Sub
'�N�U�Ӥ��q�O������ܦ� Array�r��...�� 5,Emcncc;6,EMCFM
Public Function SetCompCodeAry() As String()
  On Error GoTo ChkErr
    Dim i As Long
    Dim sTmp As String
    Dim rsComp As New ADODB.Recordset
    If Not GetRS(rsComp, "Select CodeNo,TABLEOWNER FROM CD039", gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
    rsComp.MoveFirst
    Do While Not rsComp.EOF
        sTmp = sTmp & rsComp(0) & "," & rsComp(1) & ";"
        rsComp.MoveNext
    Loop
    sTmp = Mid(sTmp, 1, Len(sTmp) - 1)
    varOwner = Split(sTmp, ";")
    SetCompCodeAry = varOwner
    On Error Resume Next
    Call CloseRecordset(rsComp)
    Exit Function
ChkErr:
  Call ErrSub(moduleName, "SetCompCodeAry")
End Function
'Public Sub ReadParaData(ByRef fs As TextStream)
'    Dim i As Long
'    Dim Xdecrypt As Variant
'    Do While Not fs.AtEndOfStream
'        strLine = Empty
'        strLine = fs.ReadLine
'        Xdecrypt = CreateObject("Encryption.Password").Decrypt(strLine, "CS")
'        Dim arydata() As String
'        arydata = Split(Xdecrypt, Chr(0))
'        rsReadFile.AddNew
'        For i = LBound(arydata) To UBound(arydata)
'            rsReadFile(i) = arydata(i)
'        Next i
'        rsReadFile.Update
'    Loop
'End Sub
'Public Sub GetPara()
'    Dim i As Long
'    Dim strTmp As String
'    If rsReadFile.RecordCount > 0 Then rsReadFile.MoveFirst
'    Do While Not rsReadFile.EOF
'        strConnectString = " Provider=MSDAORA.1;Password=" & rsReadFile("LOGINPASS") & _
'                " ;User ID=" & rsReadFile("LOGINUSER") & _
'                " ;Data Source=" & rsReadFile("DATABASE") & _
'                " ;Persist Security Info=True"
'        intAutoCDR = rsReadFile("AUTORUN")
'        intAutoConnect = rsReadFile("AUTOCONN")
'        rsReadFile.MoveNext
'    Loop
'End Sub
Public Function WriteSetPara() As Boolean
  On Error Resume Next
'    Dim LogFile As String
'    LogFile = ReadFROMINI(GetIniPath1, "GICMISPath", "ErrLogPath")
'    If Right(LogFile, 1) <> "\" Then LogFile = LogFile & "\"
'    LogFile = LogFile & "CDRPara.TXT"
'    Set FileStream = scr.CreateTextFile(LogFile, True)
'    FileStream.WriteLine strCDRPara
'    FileStream.Close
'    Set FileStream = Nothing
End Function
Public Function GetSetPara() As Boolean
  On Error Resume Next
'    Dim LogFile As String
'    LogFile = ReadFROMINI(GetIniPath1, "GICMISPath", "ErrLogPath")
'    If Right(LogFile, 1) <> "\" Then LogFile = LogFile & "\"
'    LogFile = LogFile & "CDRPara.TXT"
'    If scr.FileExists(LogFile) Then
'        Set FileStream = scr.OpenTextFile(LogFile, ForReading, False)
'        Do While Not FileStream.AtEndOfStream
'            strCDRPara = FileStream.ReadLine
'        Loop
'        strCDRPara = Trim(Replace(strCDRPara, vbCrLf, ""))
'    Else
'        strCDRPara = ""
'    End If
'    FileStream.Close
'    Set FileStream = Nothing
    GetSetPara = True
End Function
'Public Sub DataInitialize()
'    'Call SetRsField
'    If Right(App.Path, 1) = "\" Then
'        strConnectFile = App.Path & strConnectFileName
'    Else
'        strConnectFile = App.Path & "\" & strConnectFileName
'    End If
'    'Call SetRsField
'    If scr.FileExists(strConnectFile) Then
'        Set FileStream = scr.OpenTextFile(strConnectFile, ForReading, False)
'        'Call ReadParaData(FileStream)
'        FileStream.Close
'    End If
'    'Call GetPara
'End Sub
'�P�_�O�_�s�u
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
'���Ҫ�l��,Connect,errlog path,user,database,password.....
Public Function ReadCDRINI(frm As Object)  ' �� GICMIS2.INI Ū���]�w���
  On Error GoTo ChkErr
    Dim SysPath As String
    Dim strTmp As String
    Dim lngLen As Long
    Dim pstrDataSource As String
    Dim pstrUserId As String
    Dim pstrMDBPath As String
    Dim pstrPassword As String
    Dim cpCode As String
    Dim abc As String
    SysPath = GetIniPath2
    
    If Dir(SysPath) = "" Then ' �p�G�٬O�䤣�� GICMIS.INI �A�h�����t��
        MsgBox "�t�����ҰѼƳ]�w���~�A�L�k�}��" & SysPath & " !", vbCritical, "���~"
        Unload frm
        End
    End If
    
    If Command <> Empty Then
        cpCode = Command()
        cpCode = Replace(cpCode, "CRM", "", 1, , vbTextCompare)
        If cpCode <> Empty Then
            cpCode = Val(Trim(cpCode))
            If cpCode = 0 Then cpCode = ReadFROMINI(SysPath, "DEFAULT", "DEFAULTDB", False) ' �� GICMIS2.INI Ū�� �]�w��
        Else
            cpCode = ReadFROMINI(SysPath, "DEFAULT", "DEFAULTDB", False) ' �� GICMIS2.INI Ū�� �]�w��
        End If
    Else
        cpCode = ReadFROMINI(SysPath, "DEFAULT", "DEFAULTDB", False)    ' �� GICMIS2.INI Ū�� �]�w��
    End If
    
    garyGi(14) = cpCode
    pstrDataSource = ReadFROMINI(SysPath, cpCode, "1", True)
    pstrUserId = ReadFROMINI(SysPath, cpCode, "2", True)
    pstrPassword = ReadFROMINI(SysPath, cpCode, "3", True)
    gstrStationCount = ReadFROMINI(SysPath, cpCode, "4", True) ' ���\�h�֤u�@���n�J
    
    
    If Len(pstrUserId) = 0 Or Len(pstrPassword) = 0 Or Len(gstrStationCount) = 0 Then
        MsgBox "�t �� �� �� �� �� �] �w �� �~ !!", , "���~�T��..!!"
        Unload frm ' �p�G ��ƨӷ� �� �ϥΪ̥N�X �ťաA
        End
    End If
   
    If Alfa Then
        GetGlobal
    Else
        garyGi(3) = "Provider=MSDAORA.1;" & _
                     IIf(pstrDataSource = "", "", "Data Source=" + pstrDataSource + ";") & _
                    "Password=" & pstrPassword + ";" & _
                    "User ID=" & pstrUserId + ";" & _
                    "Persist Security Info=True"
    End If
    If Not Alfa Then
        gcnGi.ConnectionString = garyGi(3)
        gcnGi.CursorLocation = adUseClient
        gcnGi.Open
        garyGi(9) = Command
        gCompCode = garyGi(9)
    Else
        gCompCode = garyGi(9)
    End If
    garyGi(0) = "EPG"   '���ʤH���N���O�T�w����
    garyGi(1) = "EPG"   '���ʤH���W�٦W�٬O�T�w����
    Exit Function
ChkErr:
    If ErrHandle(, True, , "ReadINI") Then Resume 0
    Unload frm
    End
End Function
'�g�J�ۭq�����~
Public Sub WriteErr(ByVal aMsg As String, ByVal afrmName As String, ByVal aProcedureName As String)
    Dim aryErr As Variant
    Dim strErr As String
    If Not IsConnect(gcnGi) Then
        aMsg = "Oracle �_�u"
    End If
    On Error Resume Next
    aryErr = Array("�o�Ϳ��~�ɶ� : ", "�o�Ϳ��~�M�� : ", "�o�Ϳ��~��� : ", "�o�Ϳ��~�{�� : ", "���~�N�X : ", "���~��] : ")
    strErr = aryErr(0) & CStr(Now) & strCrLf & _
            aryErr(1) & "prjAutoCDR" & strCrLf & _
            aryErr(2) & afrmName & strCrLf & _
            aryErr(3) & aProcedureName & strCrLf & _
            aryErr(4) & "�ۭq���~" & strCrLf & _
            aryErr(5) & aMsg
    ErrLog strErr
End Sub
'Ū��CDRCenter.ini�����|�ɦW
Public Function GetCDRIniPath(Optional IniPath1 As String) As String
  On Error GoTo ChkErr
    Static strINI As String
    If Len(strINI) > 0 Then GoTo 66
    If Len(IniPath1) > 0 Then
        strINI = IniPath1
        GetCDRIniPath = IniPath1
        Exit Function
    End If
    strINI = IIf(Right(Environ("WinDir"), 1) <> "\", Environ("WinDir") & "\", Environ("WinDir")) & "CDRCenter.INI"
   If Not ChkDirExist(strINI) Then
        If Alfa Then
            strINI = CurPath & IIf(Right(CurPath, 1) = "\", "", "\") & "CSMISV40\BIN\CDRCenter.INI"
        Else
            strINI = App.Path & "\CDRCenter.INI"
        End If
    End If
66:
    GetCDRIniPath = strINI
  Exit Function
ChkErr:
    GetCDRIniPath = ""
    'ErrSub FormName, "GetIniPath"
End Function

'�]�wCDR���ҰѼ� �۰ʰ���ɶ�,�ǰe�R�O�ɶ�,�R�OTimeOut�ɶ�....
Public Sub GetCDRParaIni()
  On Error GoTo ChkErr
    SysPath = GetCDRIniPath
    Dim aryOwner() As String
    Dim i As Long
    If Dir(SysPath) = "" Then ' �p�G�S��CDRCenter.Ini,�����إ�,�õ������w��
        'WritePrivateProfileString "CDRCenter", "Initial", ByVal CStr("1"), SysPath
        WritePrivateProfileString "CDRCenter", "AutoRun", ByVal CStr("1"), SysPath
        WritePrivateProfileString "CDRCenter", "CmdTimeOut", ByVal CStr("90"), SysPath
        WritePrivateProfileString "CDRCenter", "ErrExitCount", ByVal CStr("3"), SysPath
        WritePrivateProfileString "CDRCenter", "UpdEn", ByVal CStr("Auto"), SysPath
        aryOwner = SetCompCodeAry
        For i = LBound(aryOwner) To UBound(aryOwner)
            Call SetPubOwner(i)
            WritePrivateProfileString CStr(gCompCode), "Initial", ByVal CStr(0), SysPath
            'WritePrivateProfileString "SendCmdTime", "Initial", ByVal CStr("0"), SysPath
            WritePrivateProfileString CStr(gCompCode), "SendCmdTime", ByVal CStr(Format(Now, "yyyymmddhhmmss")), SysPath
        Next i
        'blnInitial = True
        intAutoCDR = 1
        lngCmdTimeOut = 90
        lngErrExitCount = 3
    Else
        'blnInitial = ReadFROMINI(SysPath, aSection, "Initial") = "0"
        intAutoCDR = Val(ReadFROMINI(SysPath, aSection, "AutoRun"))
        lngCmdTimeOut = Val(ReadFROMINI(SysPath, aSection, "CmdTimeOut"))
        lngErrExitCount = Val(ReadFROMINI(SysPath, aSection, "ErrExitCount"))
        blnShowSecond = Val(ReadFROMINI(SysPath, "CDRCenter", "Second")) = 1
        strUpdEn = ReadFROMINI(SysPath, "CDRCenter", "UpdEn") & ""
        strAppTitle = ReadFROMINI(SysPath, aSection, "AppTitle") & ""
        
        
'        strLock = ReadFROMINI(SysPath, "Lock", "Over")
    End If
    Exit Sub
ChkErr:
    MsgBox "Ū���]�w�ɦ��~,���ˬd " & SysPath & " �]�w�I"
    End
End Sub
'�g�JCDR���ҰѼ�
Public Sub WriteCDRParaIni()
  On Error Resume Next
  WritePrivateProfileString CStr(gCompCode), "Initial", ByVal CStr("1"), SysPath
  WritePrivateProfileString CStr(gCompCode), "SendCmdTime", _
                            ByVal CStr(Format(MAXIncurredDate, "yyyymmddhhmmss")), _
                            SysPath
End Sub

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
    If InStr(1, strSQL, garyGi(16) & ".") = 0 And blnChkOwner And Alfa Then
        MsgBox "�ӻy�k�S�����Owner Name, �Ь��{���]�p�v!!", vbCritical, gimsgPrompt
        Exit Function
        
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
'���X�n�ǰe���_�l���
Public Function GetSendTime() As String
  On Error GoTo ChkErr:
    Dim strTime As String
    Dim ablnInitial As Boolean
    ablnInitial = ReadFROMINI(SysPath, CStr(gCompCode), "Initial") = 0
    strTime = Empty
    If ablnInitial Then
        strTime = MaxVodTime
    Else
        strTime = ReadFROMINI(SysPath, CStr(gCompCode), "SendCmdTime")
    End If
    If strTime = "" Then
        strTime = ReadFROMINI(SysPath, CStr(gCompCode), "SendCmdTime")
    End If
    GetSendTime = strTime
    Exit Function
ChkErr:
  ErrCDRSub moduleName, "GetSendTime"
End Function
'���XSO033VOD IncurredDate �̤j�����@��
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
'���X�n�g�J���~�T�����ɮ�
Public Function GetErrFile() As String
  On Error GoTo ChkErr
    Dim LogFile As String
    LogFile = ReadFROMINI(GetIniPath1, "GICMISPath", "ErrLogPath")
    If Not ChkDirExist(LogFile) Then MkDir LogFile
    If Right(LogFile, 1) <> "\" Then LogFile = LogFile & "\"
    GetErrFile = LogFile
    Exit Function
ChkErr:
  Call ErrSub(moduleName, "GetErrFile")
End Function
