VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "csACHTranReferIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const FormName = "clsACHTranRefer"
Private rsCD018 As New ADODB.Recordset
Private DataFile As TextStream
Private ErrFile As TextStream
Private MemoFile As TextStream
Dim lngTotalCount As Long, lngTotalAmt As Long
Dim lngErrCount As Long, lngErrAmt As Long
Dim lngBankTotalAmt As Long, lngBankTotalCount As Long
Dim lngRefNo As Long
Private strChooseSQL As Variant     'SQL
Private strErrPath As Variant       'INI�ɮ׸��|
Private strFilePath As Variant      '�ɮ׸��|
Private strRealDate As Variant      '�J�b���
Private strUpdEn As Variant         '�O�����ʤH��
Private strUpdName As Variant       '�O�����ʤH��
Private strUpdTime As String
Private strPTCode As Variant        '�I�ں���
Private strPTName As Variant
Private strServiceType As Variant   '�A�����O
Private strBankSQL As Variant
Private strCompCode As Variant      '���q�O
Private strClctEn As Variant        '���O�H��
Private strClctName As Variant
Private strCMCode As Variant        '���O�覡
Private strCMName As Variant
Private lngPara24 As Long
Private strCitemQry As String       '���O����
Private intFlowId As Integer         '�y�{����
Private FSO As New FileSystemObject
Private blnCkZero As Boolean        '�t�����B�O�_����
Private rsCD013 As New ADODB.Recordset
Private strFailUpdSQL As String
Private strUCCode As String
Private strUCName As String
Private blnFailUpd As Boolean
Private strBankIdH As String
Private isCrossCustCombine  As Boolean
Public Function Action() As Boolean
    On Error GoTo ChkErr
    blnFailUpd = False
    
    '*************************************************************************************************************************************
    '#3417 Ū��������]��5,�p�G�����,��s���Ѯɭn�^�gUCCode�PUCName By Kin 2007/12/11
    If Not GetRS(rsCD013, "Select CodeNo,Description From " & GetOwner & "CD013 Where Nvl(RefNo,0)=5 And StopFlag<>1 Order by CodeNo Desc", _
                  gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
    If rsCD013.EOF Then
        blnFailUpd = False
    Else
        strUCCode = rsCD013("CodeNo") & ""
        strUCName = rsCD013("Description") & ""
        blnFailUpd = True
    End If
    '*************************************************************************************************************************************
    If Not InitData Then Exit Function
    If Not OpenData Then Exit Function
    If Not BeginTran Then Exit Function
    Action = True
        
    Exit Function
ChkErr:
    ErrSub FormName, "Action"
End Function

Private Function InitData() As Boolean
    On Error GoTo ChkErr
        With clsStoreParameter
'            blnBeOne = .uBeOne
            strBankSQL = .uBankSQL
            strChooseSQL = .uChooseSQL
            strCompCode = .uCompCode
            strErrPath = .uErrPath & IIf(Right(.uErrPath, 1) = "\", "", "\")
            strFilePath = .uFilePath
            strPTCode = .uPTCode
            strPTName = .uPTName
            strClctEn = .uClctEn
            strClctName = .uClctName
            strRealDate = .uRealDate
            strServiceType = .uServiceType
            strUpdEn = .uUpdEn
            strUpdName = .uUpdName
            strUpdTime = GetDTString(RightNow)
            strCMCode = .uCMCode
            strCMName = .uCMName
            strCitemQry = .uCitemQry
            intFlowId = .FlowId
            blnCkZero = .uCkZero
        End With
        lngPara24 = Val(GetSystemParaItem("Para24", CStr(strCompCode), CStr(strServiceType), "SO043") & "")
        InitData = True
    Exit Function
ChkErr:
    Call ErrSub(FormName, "InitData")
End Function

Private Function BeginTran() As Boolean
    On Error GoTo ChkErr
    Dim lngTime As Long
    Dim strReadLine As String
    Dim blnBigError As Boolean, blnTotalLog As Boolean
    Dim rs As New ADODB.Recordset
    Dim DefClctEn As String, DefClctName As String
    Dim strBankId As String, strBankName As String
    Dim rsTotal As New ADODB.Recordset
    Dim DefBankCode As String, DefBankName As String
    isCrossCustCombine = Val(GetRsValue("Select Nvl(CrossCustCombine,0) From " & GetOwner & "SO041", gcnGi) & "") = 1
        lngTime = Timer
        lngBankTotalCount = 0: lngBankTotalAmt = 0
        lngErrCount = 0: lngErrAmt = 0
        On Error GoTo KillFile
        '�w�]���O�H��
        
        DefClctEn = strClctEn
        DefClctName = strClctName
        DefBankCode = GetRsValue("Select CodeNo From " & GetOwner & "CD018 Where CodeNo " & strBankSQL) & ""
        DefBankName = GetRsValue("Select Description From " & GetOwner & "CD018 Where CodeNo " & strBankSQL) & ""
        strBankIdH = DefBankCode
        gcnGi.BeginTrans
        If Not DefineTotalLogRs(rsTotal) Then Exit Function
        Do While Not DataFile.AtEndOfLine
            With DataFile
                strReadLine = .ReadLine
                If Not .AtEndOfLine Then
                    If ChkDataOk(strReadLine, blnBigError, blnTotalLog, DefBankName) Then
                        '�p�X���ߤJ�b�ɪ��ʧ@
                            strBankId = GetAccountNo(strReadLine)
                        If strBankId <> "" Then
                            If Not GetRS(rs, "Select A.EmpNo,B.EmpName From " & GetOwner & "CD018 A," & GetOwner & "CM003 B Where A.EmpNo = B.EmpNo And CodeNo =" & strBankId, gcnGi) Then Exit Function
                                If rs.EOF Then
                                    strClctEn = DefClctEn
                                    strClctName = DefClctName
                                Else
                                    strClctEn = rs("EmpNo") & ""
                                    strClctName = rs("EmpName") & ""
                                End If
'                            End If
                        Else
                            strBankId = DefBankCode
                            strBankName = DefBankName
                        End If
                        
                        If Not UpdateToSO033(strReadLine) Then
                            MsgBox "��s���� !! �Ьd�� !! ", vbCritical, gimsgWarning
                            GoTo KillFile
                        End If
                        '���T��....
                        If Not UpdateBankTotal(rsTotal, strBankId, GetRealAmt(strReadLine), 0) Then GoTo KillFile
                    Else
                        '���~��....
                        If blnTotalLog Then If Not UpdateBankTotal(rsTotal, strBankId, GetRealAmt(strReadLine), -1) Then GoTo KillFile
                    End If
                    If blnBigError Then GoTo KillFile
                End If
            End With
        Loop
        gcnGi.CommitTrans
'        gcnGi.RollbackTrans
        Call CloseFS(False)
        If lngErrCount > 0 Then
            Shell "NotePad " & GetErrFileName(strErrPath), vbNormalFocus
        End If
        MsgBox "���\����: " & lngBankTotalCount & " ,���\���B: " & lngBankTotalAmt & vbCrLf & vbCrLf & _
                    "���ѵ���: " & lngErrCount & " ,���Ѫ��B: " & lngErrAmt & vbCrLf & vbCrLf & _
                    "�@��O : " & Round(Timer - lngTime) & " ��", vbInformation, gimsgPrompt
        Call CloseRecordset(rs)
        BeginTran = True
    Exit Function
KillFile:
    gcnGi.RollbackTrans
    '�R���ɮ�..
    Call CloseFS(False)
    Exit Function
ChkErr:
    Call ErrSub(FormName, "BeginTran")
End Function

Private Function UpdateToSO033(strReadLine As String) As Boolean
    On Error GoTo ChkErr
    Dim strBillNo As String
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim lngAffected As Long
    Dim aRealDate As String
        strBillNo = GetBillNo(strReadLine)
        strSQL = "Update " & GetOwner & "So033 Set RealAmt = ShouldAmt" & _
                ",UpdTime = '" & strUpdTime & "'" & _
                ",FirstTime = '" & strUpdTime & "'" & _
                ",UpdEn = '" & strUpdName & "'" & _
                ",UCCode = Null, UCName = Null"
        If strRealDate = "" Then
            strRealDate = GetRealDate(strReadLine)
        End If
        If Len(strRealDate) > 0 Then
            'aRealDate = Left(strRealDate, 4) & "/" & Mid(strReadLine, 5, 2) & "/" & Right(strRealDate, 2)
            aRealDate = Left(strRealDate, 4) & "/" & Mid(strRealDate, 5, 2) & "/" & Right(strRealDate, 2)
        End If
        strSQL = strSQL & ",RealDate = To_date('" & strRealDate & "','YYYYMMDD')"
        '���O�H��
        If strClctEn <> "" Then
            strSQL = strSQL & ",ClctEn = '" & strClctEn & "'" & ",ClctName = '" & strClctName & "'"
            '****************************************************************************
            '#3417 �p�G����J���O�H��SO033.Note�n��Ȧ�N�X By Kin 2008/01/22
            If strBankIdH <> Empty Then
                strSQL = strSQL & ",Note=Note || '�Ȧ�N�X:" & strBankIdH & "'"
            End If
            '****************************************************************************
        Else
            strSQL = strSQL & ",ClctEn = OldClctEn ,ClctName = OldClctName "
        End If
        '���O�覡
        If strCMCode <> "" Then strSQL = strSQL & ",CMCode = " & strCMCode & ", CMName = '" & strCMName & "'"
        '�I�ں���
        If strPTCode <> "" Then strSQL = strSQL & ",PTCode = " & strPTCode & ", PTName = '" & strPTName & "'"
        '#4010 ����GetWhereCondition�ǤJ���Ѽ�(blnCD013) By Kin 2008/07/30
        '#6307 �n�W�[CancelFalg,GetWhereCondition�ĤG�ӶǤJ���޼ƭn�אּFalse By Kin 2012/09/12
        '#6307 ���դ�OK,�n��SO033�Ҧ���Field���զX�X��,���M��FCD013��Recordset��Custid�PBillNo�ܩ_�Ƿ|�۰ʳQ�I�� By Kin 2102/09/18
        Dim aSQL As String
        aSQL = "SELECT * FROM SO033 WHERE 1=0"
        If Not GetRS(rsTmp, aSQL, _
                    gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
        Dim aFields As String
        Dim i As Long
        For i = 0 To rsTmp.Fields.Count - 1
            If i = 0 Then
                aFields = "SO033." & rsTmp.Fields(0).Name
            Else
                aFields = aFields & ",SO033." & rsTmp.Fields(i).Name
            End If
        Next i
        aSQL = "Select " & aFields & " From " & GetOwner & "So033," & GetOwner & "CD013 " & GetWhereCondition(strBillNo, False)
            
        If Not GetRS(rsTmp, aSQL, _
                    gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
        Set rsTmp.ActiveConnection = Nothing
        Do While Not rsTmp.EOF
            rsTmp("RealAmt") = rsTmp("ShouldAmt")
            rsTmp("UpdTime") = rsTmp("UpdTime")
            rsTmp("FirstTime") = rsTmp("UpdTime")
            rsTmp("UpdEn") = rsTmp("UpdEn")
            rsTmp("UCCode") = Null
            rsTmp("UCName") = Null
            'rsTmp("RealDate") = rsTmp("RealDate")
            rsTmp("RealDate") = aRealDate
            rsTmp("ClctEn") = rsTmp("ClctEn")
            rsTmp("ClctName") = rsTmp("ClctName")
            rsTmp("CMCode") = rsTmp("CMCode")
            rsTmp("CMName") = rsTmp("CMName")
            rsTmp("PTCode") = rsTmp("PTCode")
            rsTmp("PTName") = rsTmp("PTName")
            If Not AlterSO003(rsTmp) Then Exit Function
            rsTmp.CancelUpdate
            rsTmp.MoveNext
        Loop
        '#4010 ����GetWhereCondition�ǤJ���Ѽ�(blnCD013) By Kin 2008/07/30
        If Not ExecuteCommand(strSQL & GetWhereCondition(strBillNo, False, False), gcnGi, lngAffected) Then Exit Function
'        Do While Not rsTmp.EOF
'            If Not AlterSO003(rsTmp) Then Exit Function
'            rsTmp.MoveNext
'        Loop
        lngBankTotalAmt = lngBankTotalAmt + GetRealAmt(strReadLine)
        lngBankTotalCount = lngBankTotalCount + 1
'        lngTotalAmt = lngTotalAmt + GetRealAmt(strReadLine)
'        lngTotalCount = lngTotalCount + 1
        UpdateToSO033 = True
    Exit Function
ChkErr:
    ErrSub FormName, "UpdateToSO033"
End Function

Private Function ChkDataOk(strReadLine As String, _
        Optional BigError As Boolean = False, Optional TotalLog As Boolean = True, _
        Optional strBankName As String) As Boolean
    On Error GoTo ChkErr
    Dim rsTmp As New ADODB.Recordset
    Dim strBillNo As String, strErrDescription As String
    Dim lngBudgetPeriod As Long
    Dim strResult As String
    Dim strChkDate As String
    Dim strErrCode As String
    Dim strCD013RefNo As String
    If strRealDate = "" Then
        strChkDate = GetRealDate(strReadLine)
    Else
        strChkDate = strRealDate
    End If
    strCD013RefNo = " MAX(DECODE(NVL(CD013.PAYOK,0),1,3," & _
                "DECODE(NVL(CD013.REFNO,0),3,3,DECODE(NVL(CD013.REFNO,0),7,3,0)))) REFNO "
    Select Case Left(strReadLine, 3)
        Case "BOF"
            Exit Function
        Case "EOF"
            TotalLog = False
            If lngBankTotalAmt <> GetTotalAmtOk(strReadLine) Then
                MsgBox "���������`���B�P���\���Ƥ��`���B���P", vbCritical, gimsgWarning
                BigError = True
            End If
            If lngBankTotalCount <> GetTotalCountOk(strReadLine) Then
                MsgBox "���������`���ƻP���\���Ƥ��P", vbCritical, gimsgWarning
                BigError = True
            End If
            Exit Function
        Case Else
            strErrCode = Empty
            strErrDescription = Empty
            Select Case Mid(strReadLine, 65, 2)
                Case "00" '                00 = ��b���\
                    strErrDescription = ""
                Case "01" '                01 = �s�ڤ���
                    strErrDescription = "�s�ڤ��� !! "
                    strErrCode = "01"
                Case "02" '                02 = �D�e�U�Nú�N�o��
                    strErrDescription = "�D�e�U�Nú�N�o�� !! "
                    strErrCode = "02"
                Case "03" '                03 = �w�פ�e�U�Nú��
                    strErrDescription = "�w�פ�e�U�Nú�� !! "
                    strErrCode = "03"
                Case "04" '                04 = �L���b��
                    strErrDescription = "�L���b�� !! "
                    strErrCode = "04"
                Case "05" '                05 = �����̲νs���~
                    strErrDescription = "�����̲νs���~ !! "
                    strErrCode = "05"
                Case "06" '                06 = �L���Τḹ�X
                    strErrDescription = "�L���Τḹ�X !! "
                    strErrCode = "06"
                Case "07" '                07 = �Τḹ�X����
                    strErrDescription = "�Τḹ�X���� !! "
                    strErrCode = "07"
                Case "22" '                22 = �b��w���M
                    strErrDescription = "�b���w���M !! "
                    strErrCode = "22"
                Case "23"
                    strErrDescription = "�R��� !!"     '#4429�W�[�^���� By Kin 2009/05/19
                    strErrCode = "23"
                Case "24"
                    strErrDescription = "�ᵲ�� !!"     '#4429�W�[�^���� By Kin 2009/05/19
                    strErrCode = "24"
                Case "25"
                    strErrDescription = "�b��s�ھD�k�|�j����� !!"     '#4429�W�[�^���� By Kin 2009/05/19
                    strErrCode = "25"
                Case "99"
                    strErrDescription = "��L�����\��] !! "
                    strErrCode = "99"
                Case Else
                    strErrDescription = "��L�����\��] !! "
                    strErrCode = Mid(strReadLine, 65, 2)
            End Select
            '�Ȧ椧���~
            If strErrDescription <> "" Then
                Call InsertToErr(strReadLine, strErrDescription, strBankName, strErrCode)
                Exit Function
            End If
            
            '�{���ˬd�����~
            strBillNo = GetBillNo(strReadLine)
            '*****************************************************************************
            '#5386 �P�_�O�_���ۦP��ڸ����h���Ƚs By Kin 2009/12/09
            If Not isCrossCustCombine Then
                If Not GetRS(rsTmp, "Select CustId From " & strOwnerName & "SO033," & strOwnerName & "CD013 " & _
                            GetWhereCondition(strBillNo, False) & " Group By CustId") Then
                     BigError = True
                     Exit Function
                End If
                
                If Not rsTmp.EOF Then
                    If rsTmp.RecordCount > 1 Then
                        Dim strDiffCust As String
                        rsTmp.MoveFirst
                        Do While Not rsTmp.EOF
                            strDiffCust = IIf(strDiffCust = "", rsTmp("CustId"), strDiffCust & ";" & rsTmp("CustId"))
                            rsTmp.MoveNext
                        Loop
                        Call InsertToErr(strReadLine, "�ۦP��ڽs�����Ȥ�s�����P�I �Ƚs�G" & strDiffCust, strBankName)
                        Exit Function
                    End If
                End If
            End If
            '**********************************************************************************

            
            
            '#4010 UCCode�ѦҸ���3�n�������`��� By Kin 2008/07/30
            '#5564 CD013.PAYOK=1��CD013.REFNO=3,7�����d�x�w�� By Kin 2010/05/20
            If Not GetRS(rsTmp, "Select Nvl(Sum(SO033.ShouldAmt),0) as TotalAmt," & strCD013RefNo & " From " & GetOwner & "So033," & GetOwner & "CD013 " & GetWhereCondition(strBillNo, False)) Then
                BigError = True
                Exit Function
            End If
            If rsTmp.RecordCount = 0 Then
                Call InsertToErr(strReadLine, "�L����ڽs���Φ��渹�w���� !! ", strBankName)
                Exit Function
            End If
            If Val(rsTmp("RefNo") & "") = 3 Then
                Call InsertToErr(strReadLine, "�d�O�w�� !!", strBankName)
                Exit Function
            End If
            If rsTmp("TotalAmt") = 0 Then
                Call InsertToErr(strReadLine, "�L����ڽs���Φ��渹�w���� !! ", strBankName)
                Exit Function
            End If
            If Val(rsTmp("TotalAmt") & "") <> GetRealAmt(strReadLine) Or Val(rsTmp("TotalAmt") & "") = 0 Then
                Call InsertToErr(strReadLine, "�Ȧ���b���B�P�������B���� !! ", strBankName)
                Exit Function
            End If
            '#3122 �ˮ֬O�_���W�L�鵲��� By Kin Add 2007/03/29
            If Not ChkOverCloseDate(Format(strChkDate, "####/##/##"), strCompCode, strServiceType, strResult) Then
                Call InsertToErr(strReadLine, strResult, strBankName)
                strResult = ""
                Exit Function
            End If
            '#4010 ����GetWhereCondition�ǤJ���Ѽ�(blnCD013) By Kin 2008/07/30
            If Not GetRS(rsTmp, "Select RealPeriod,CustId,CompCode,ServiceType,CitemCode,FaciSeqNo From " & GetOwner & "So033 " & GetWhereCondition(strBillNo, False, False)) Then Exit Function
            '94/09/30 1805 Jacky
            Do While Not rsTmp.EOF
                lngBudgetPeriod = Val(GetRsValue("Select Nvl(BudgetPeriod,0) From " & GetOwner & "SO003 Where Custid = " & rsTmp("CustId") & " And CompCode = " & rsTmp("CompCode") & " And ServiceType = '" & rsTmp("ServiceType") & "' " & _
                    " And CitemCode = " & rsTmp("CitemCode") & " And FaciSeqNo = '" & rsTmp("FaciSeqNo") & "'") & "")
                If rsTmp("RealPeriod") > lngBudgetPeriod And lngBudgetPeriod > 0 Then
                    Call InsertToErr(strReadLine, "�J�b���Ƥj������Ѿl����,�L�k�J�b!!", strBankName)
                    Exit Function
                End If
                rsTmp.MoveNext
            Loop
    End Select
    Call CloseRecordset(rsTmp)
    ChkDataOk = True
    Exit Function
ChkErr:
    ErrSub "csACHTranReferIn", "ChkDataOk"
End Function
Private Function InsertToErr(strReadLine As String, strErrDescription As String, strBankName As String, Optional strErrCode As String = Empty) As Boolean
  On Error GoTo ChkErr
    Dim strBillNo As String
    Dim strAccountNo As String
    strBillNo = GetBillNo(strReadLine)
    strAccountNo = GetAccountNo(strReadLine)
'        WriteTextLine ErrFile, strBillNo & " " & strErrDescription
    WriteTextLine ErrFile, "�Ȧ�W��: " & strBankName & "; ���O�渹:" & strBillNo & "; ���~�W��:" & strErrDescription & "; �b��:" & strAccountNo
    
    '*****************************************************************************************************************************************************
    '#3417 �Y�����~��sUCCode�PUCName By Kin 2007/12/11
    If blnFailUpd Then
        '#4010 ����GetWhereCondition�ǤJ�Ѽ�(blnCD013) By Kin 2008/07/30
        '#4388 ��[��s���ʤH���P���ʮɶ� By Kin 2009/04/30
        '#52056 �p�G�d�x�w��,����s�d�x�w������� By Kin 2009/04/30
        strFailUpdSQL = "UPDATE " & GetOwner & "SO033 Set UCCode=" & strUCCode & _
                        ",UCName='" & strUCName & "'" & _
                        ",UPDEN='" & garyGi(1) & "'" & _
                        ",UPDTime='" & strUpdTime & "'" & _
                        ",Note=Note||'" & IIf(strErrCode <> Empty, " ���ѥN�X:" & strErrCode, "") & _
                        " �Ȧ�W��: " & strBankName & "; ���O�渹:" & strBillNo & _
                        "; ���~�W��:" & strErrDescription & "; �b��:" & strAccountNo & "'" & _
                         GetWhereCondition(strBillNo, False, False) & _
                         " And UCCode NOT IN (SELECT CodeNo From CD013 Where CD013.REFNO=3)"
        
        gcnGi.Execute strFailUpdSQL
    End If
    '*****************************************************************************************************************************************************
    
    lngErrCount = lngErrCount + 1
    lngErrAmt = lngErrAmt + GetRealAmt(strReadLine)
    InsertToErr = True
  Exit Function
ChkErr:
  Call ErrSub("csACHTranReferIn", "InsertToErr")
End Function

Private Function ChkErrCustId(rs As ADODB.Recordset, _
        ByRef lngErrCount As Long) As Boolean
    On Error GoTo ChkErr
    Dim rsTmp As New ADODB.Recordset
        If Not GetRS(rsTmp, "Select CustName,MailAddress From " & GetOwner & "So001 Where CustId = " & rs("CustId")) Then Exit Function
        If InStr(1, rsTmp("MailAddress") & "", "��") = 0 Then
           WriteTextLine ErrFile, rs("CustId") & rsTmp("CustName") & rs("BillNo") & rsTmp("MailAddress")
           lngErrCount = lngErrCount + 1
           Exit Function
           '���D��
        End If
        ChkErrCustId = True
    Exit Function
ChkErr:
    ErrSub "csACHTranReferIn", "ChkErrCustId"
End Function

Private Function OpenData() As Boolean
    On Error GoTo ChkErr
        If Not OpenFile(DataFile, GetDataFileName(strFilePath), True, giOpenTEXT) Then Exit Function
        If Not OpenFile(ErrFile, GetErrFileName(strErrPath), True) Then Exit Function
        If Not OpenFile(MemoFile, GetMemoFileName(strErrPath), True) Then Exit Function
        lngRefNo = Val(GetRsValue("Select RefNo From " & GetOwner & "CD018 Where CodeNo " & strBankSQL) & "")
        OpenData = True
    Exit Function
ChkErr:
    ErrSub FormName, "OpenData"
End Function

Private Function CloseFS(Optional KillFile As Boolean = False) As Boolean
    On Error Resume Next
        ErrFile.Close
        MemoFile.Close
        DataFile.Close
    Set FSO = Nothing
        If KillFile Then
            Kill GetDataFileName(strFilePath)
            Kill GetErrFileName(strErrPath)
            Kill GetMemoFileName(strErrPath)
        End If
'        fso = Nothing
    Exit Function
ChkErr:
    ErrSub "csACHTranReferIn", "CloseFS"
End Function

Public Property Let uClctEn(ByVal vData As Variant)
    On Error Resume Next
        strClctEn = vData
End Property

Public Property Let uClctName(ByVal vData As Variant)
    On Error Resume Next
        strClctName = vData
End Property

'�M�θ�ư�44(�S���渹+�ˬd�X+�渹),�O�d��12
'or �M�θ�ư�44(�����Ҧr��ID+�渹),�O�d��12
Private Function GetBillNo(strReadLine As String) As String
    On Error Resume Next
    Dim strBillNo As String
    Dim strType As String
    Dim strYM As String
    Dim strMM As String
        
        If lngRefNo = 1 Then
            '��X�榡
            strBillNo = Mid(strReadLine, 75, 11)
            GetBillNo = GetBillNo_Old(strBillNo)
            Exit Function
        ElseIf lngRefNo = 2 Then
            '�u��������
            strBillNo = Mid(strReadLine, 76, 11)
        ElseIf lngRefNo = 3 Then
            '�x�W�Ȧ��(�N��)
            strBillNo = Mid(strReadLine, 65, 11)
        ElseIf lngRefNo = 4 Then
            '�p�X���ߥ�(�N��)
            strBillNo = Mid(strReadLine, 78, 11)
            GetBillNo = GetBillNo_Old(strBillNo)
            Exit Function
        ElseIf lngRefNo = 5 Then
            '�g�a�Ȧ��(�N��)
            strBillNo = Mid(strReadLine, 81, 11)
            GetBillNo = GetBillNo_Old(strBillNo)
            Exit Function
        Else
            '��L�a������榡
            strBillNo = Mid(strReadLine, 129, 11)
        End If
        
        '�h�b���X�b
        If lngPara24 = 1 Then
            GetBillNo = strBillNo
        Else
    
            '1=BC,2=TC,3=BI,4=TI
            strYM = Left(strBillNo, 2)
            strMM = Mid(strBillNo, 3, 1)
            If strMM > "9" Then
                strYM = strYM & Format(Trim(Asc(strMM) - Asc("A") + 10), "00")
            Else
                strYM = strYM & Format(strMM, "00")
            End If
            Select Case Mid(strBillNo, 4, 1)
                Case 1
                    strType = "BC"
                Case 2
                    strType = "TC"
                Case 3
                    strType = "BI"
                Case 4
                    strType = "TI"
                Case 5
                    strType = "BD"
                Case 6
                    strType = "TD"
                Case 7
                    strType = "BP"
                Case 8
                    strType = "TP"
            End Select
            
            GetBillNo = Trim(CStr(Left(strYM, 4) + 191100)) & _
                    strType & Format(Right(strBillNo, 7), "0000000")
        End If
End Function

Private Function GetBillNo_Old(strBillNo As String) As String
    On Error Resume Next
        If lngPara24 = 1 Then
            GetBillNo_Old = strBillNo
        Else
            GetBillNo_Old = Trim(Val(Mid(strBillNo, 1, 4)) + 191100) & Mid(strBillNo, 5, 1) & strServiceType & Format(Right(strBillNo, 6), "0000000")
        End If

End Function

Private Function GetRealAmt(strReadLine As String) As Long
    On Error Resume Next
        GetRealAmt = Val(Mid(strReadLine, 55, 10))
End Function

Private Function GetTotalAmtOk(strReadLine As String) As Long
    On Error Resume Next
        GetTotalAmtOk = Val(Mid(strReadLine, 40, 16))
End Function

Private Function GetTotalCountOk(strReadLine As String) As Long
    On Error Resume Next
        GetTotalCountOk = Val(Mid(strReadLine, 32, 8))
End Function

Private Function GetTotalAmtFail(strReadLine As String) As Long
    On Error Resume Next
        GetTotalAmtFail = Val(Mid(strReadLine, 53, 14))
End Function

Private Function GetTotalCountFail(strReadLine As String) As Long
    On Error Resume Next
        GetTotalCountFail = Val(Mid(strReadLine, 69, 10))
End Function

Private Function GetRealDate(strReadLine As String) As String
    On Error Resume Next
        GetRealDate = Format(Trim(CStr(Val(Mid(strReadLine, 21, 6)) + 19110000)), "########")
End Function

Private Function GetWhereCondition(strBillNo As String, blnUpd As Boolean, Optional ByVal blnCD013 As Boolean = True) As String
    On Error Resume Next
    Dim strSQL As String

        If lngPara24 = 1 Then
            strSQL = "SO033.MediaBillNo = '" & strBillNo & "'"
        Else
            strSQL = "SO033.BillNo = '" & strBillNo & "'"
        End If
        'If strCitemQry <> "" Then strSQL = strSQL & IIf(strSQL = "", "", " And ") & " SO033.CitemCode In (" & strCitemQry & ")"
         If strCitemQry <> "" Then strSQL = strSQL & IIf(strSQL = "", "", " And " & _
                        " SO033.CitemCode In (Select CitemCode From " & GetOwner & "CD068A Where BillHeadFmt ='" & strCitemQry & "'  )")
        '#4010 UCCode�ѦҸ���3�n�������`��� By Kin 2008/07/30
        If blnCD013 Then
            strSQL = IIf(strSQL = "", " SO033.UCCode=CD013.CodeNo ", strSQL & " And SO033.UCCode=CD013.CodeNo ")
        End If
        If Not blnUpd Then
            GetWhereCondition = " Where " & strSQL & IIf(strSQL = "", "", " And ") & " SO033.RealDate Is Null And SO033.CancelFlag=0 And SO033.UCCode is Not Null " & IIf(blnCkZero = False, "And SO033.ShouldAmt > 0", "")
        Else
            GetWhereCondition = " Where " & strSQL
        End If
'ChkErr:
'    ErrSub FormName, "GetWhereCondition"
End Function

Private Function GetDataFileName(strPath As Variant) As String
    On Error Resume Next
        GetDataFileName = strPath
End Function

Private Function GetErrFileName(strPath As Variant) As String
    On Error Resume Next
        GetErrFileName = strPath & "ACHTranReferErr.Txt"
End Function

Private Function GetMemoFileName(strPath As Variant) As String
    On Error Resume Next
        GetMemoFileName = strPath & "ACHTranReferMem.Txt"
End Function

'�p�X���ߤJ�b�ɭn���b�����Ȧ�N�X
'�� GetBankId
Private Function GetAccountNo(strReadLine As String) As String
    On Error Resume Next
    Dim intLen As Integer
        intLen = Val(GetRsValue("Select ActLength From " & GetOwner & "CD018" & " Where BankId2 = '" & GetBankCode(strReadLine) & "'") & "")
        GetAccountNo = Trim(Right(Mid(strReadLine, 41, 14), intLen))
End Function

Private Function GetBankCode(strReadLine As String) As String
    On Error Resume Next
        GetBankCode = Trim(Mid(strReadLine, 34, 7))
End Function
Private Function GetBankId(strBillNo As String, ByRef BankName As String) As String
    On Error Resume Next
    Dim lngCustId As Long
    Dim rs As New ADODB.Recordset
'        lngCustId = GetRS(rs, "Select CustId From " & strOwnerName & "So033 Where BillNo = '" & strBillNo & "'") & ""
'        If Not GetRS(rs, "Select BankCode,BankName From " & strOwnerName & "So002 Where CustId = " & lngCustId) Then Exit Function

        If Not GetRS(rs, "Select CustId,BankCode,BankName From " & GetOwner & "So033 Where BillNo = '" & strBillNo & "'") Then Exit Function
        GetBankId = rs("BankCode") & ""
        BankName = rs("BankName") & ""
End Function

Private Function UpdateBankTotal(rs As ADODB.Recordset, _
    strBankCode As String, lngTotalAmt As Long, lngType As Long) As Boolean
    On Error GoTo ChkErr
        With rs
            If strBankCode = "" Then UpdateBankTotal = True: Exit Function
            If Not (.BOF And .EOF) Then .MoveFirst
            
            .Find "BankCode = " & strBankCode
            If .EOF Then .AddNew
            .Fields("BankCode") = strBankCode
            If lngType = 0 Then
                .Fields("IntCount") = Val(.Fields("IntCount") & "") + 1
                .Fields("TotalAmt") = Val(.Fields("TotalAmt") & "") + lngTotalAmt
            Else
                .Fields("ErrCount") = Val(.Fields("ErrCount") & "") + 1
                .Fields("ErrTotalAmt") = Val(.Fields("ErrTotalAmt") & "") + lngTotalAmt
            End If
        End With
        UpdateBankTotal = True
    Exit Function
ChkErr:
    ErrSub FormName, "UpdateBankTotal"
End Function

Private Function DefineTotalLogRs(ByRef rs As ADODB.Recordset) As Boolean
    On Error GoTo ChkErr
        Set rs = New ADODB.Recordset
        With rs
            .CursorLocation = adUseClient
            .CursorType = adOpenKeyset
            .LockType = adLockOptimistic
            .Fields.Append "BankCode", adSingle
            .Fields.Append "IntCount", adSingle
            .Fields.Append "TotalAmt", adSingle
            .Fields.Append "ErrCount", adSingle
            .Fields.Append "ErrTotalAmt", adSingle
            .Open
        End With
        DefineTotalLogRs = True
    Exit Function
ChkErr:
    ErrSub FormName, "DefineTotalLogRs"
End Function
'#3122 �ˮ֬O�_���W�L�鵲��� By Kin Add 2007/03/29
Public Function ChkOverCloseDate(ByVal strCloseDate As String, _
    ByVal strCompCode As String, ByVal strServiceType As String, strResult) As Boolean
    On Error GoTo ChkErr
    Dim intDayCut As Integer
    Dim strTranDate As String
    Dim intPara6 As Integer
        ChkOverCloseDate = False
        intDayCut = Val(GetSystemParaItem("DayCut", strCompCode, "", "SO041") & "")
        strTranDate = GetRsValue("Select tranDate from " & GetOwner & "So062 Where CompCode= " & strCompCode & " And Type=1 Order By TranDate Desc") & ""
        '���O�Ѽ��ɡ]SO043)�A�����O��n������<Para6>�A�ѫ���ϥΡI
        intPara6 = Val(GetSystemParaItem("Para6", strCompCode, strServiceType, "SO043") & "")
    
        If strCloseDate = "" Then strCloseDate = RightDate
        
        If Not IsDate(strCloseDate) Then
            strResult = "������X�k�I"
            Exit Function
        End If
        
        If CDate(strCloseDate) > RightDate Then strResult = "������W�L���Ѥ��": Exit Function

        If (RightDate - CDate(strCloseDate)) > intPara6 Then
            strResult = "������w�W�L�t�γ]�w���w�������I"
            Exit Function
        End If
        If CDate(strCloseDate) <= IIf(strTranDate <> "", CDate(strTranDate), 0) Then
            If intDayCut = 1 And CDate(strCloseDate) = IIf(strTranDate <> "", CDate(strTranDate), 0) Then
                ChkOverCloseDate = True
                Exit Function
            End If
            strResult = " �w�L�굲������i�J�b�I"
            Exit Function
        End If
        ChkOverCloseDate = True
    Exit Function
ChkErr:
    ErrSub "csACHTranReferIn", "ChkOverCloseDate"
End Function

