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

'Dim clsStoreParameter As clsStoreParameter

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
Private blnBeOne As Boolean         '�X�֤J�b
Private lngPara24 As Long
Private strCitemQry As String       '���O����
Private intFlowId As Integer         '�y�{����

Public Function Action() As Boolean
    On Error GoTo ChkErr
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
            blnBeOne = .uBeOne
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
        lngTime = Timer
        On Error GoTo KillFile
        '�w�]���O�H��
        DefClctEn = strClctEn
        DefClctName = strClctName
        DefBankCode = GetRsValue("Select CodeNo From " & strOwnerName & "CD018 Where CodeNo " & strBankSQL) & ""
        DefBankName = GetRsValue("Select Description From " & strOwnerName & "CD018 Where CodeNo " & strBankSQL) & ""
        gcnGi.BeginTrans
        If Not DefineTotalLogRs(rsTotal) Then Exit Function
        Do While Not DataFile.AtEndOfLine
            With DataFile
                strReadLine = .ReadLine
                If Not .AtEndOfLine Then
                    If ChkDataOk(strReadLine, blnBigError, blnTotalLog, DefBankName) Then
                        '�p�X���ߤJ�b�ɪ��ʧ@
                        If blnBeOne Then
                            Select Case intFlowId
                            Case 1
                                'TBC
                                strBankId = GetBankId(GetBillNo(strReadLine), DefBankName)
                            Case 0
                                'EMC
                                strBankId = GetAccountNo(strReadLine)
                            End Select
                            If strBankId <> "" Then
                                If Not GetRS(rs, "Select A.EmpNo,B.EmpName From " & strOwnerName & "CD018 A," & strOwnerName & "CM003 B Where A.EmpNo = B.EmpNo And CodeNo =" & strBankId, gcnGi) Then Exit Function
                                If rs.EOF Then
                                    strClctEn = DefClctEn
                                    strClctName = DefClctName
                                Else
                                    strClctEn = rs("EmpNo") & ""
                                    strClctName = rs("EmpName") & ""
                                End If
                            End If
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
        Call CloseFS(False)
        If lngErrCount > 0 Then
            Shell "NotePad " & GetErrFileName(strErrPath), vbNormalFocus
        End If
        MsgBox "���\����: " & lngTotalCount & " ,���\���B: " & lngTotalAmt & vbCrLf & vbCrLf & _
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
        strBillNo = GetBillNo(strReadLine)
        strSQL = "Update " & strOwnerName & "So033 Set RealAmt = ShouldAmt" & _
                ",UpdTime = '" & strUpdTime & "'" & _
                ",FirstTime = '" & strUpdTime & "'" & _
                ",UpdEn = '" & strUpdName & "'" & _
                ",UCCode = Null, UCName = Null"
        If strRealDate = "" Then
            strRealDate = GetRealDate(strReadLine)
        End If
        strSQL = strSQL & ",RealDate = To_date('" & strRealDate & "','YYYYMMDD')"
        '���O�H��
        If strClctEn <> "" Then
            strSQL = strSQL & ",ClctEn = '" & strClctEn & "'" & ",ClctName = '" & strClctName & "'"
        Else
            strSQL = strSQL & ",ClctEn = OldClctEn ,ClctName = OldClctName "
        End If
        '���O�覡
        If strCMCode <> "" Then strSQL = strSQL & ",CMCode = " & strCMCode & ", CMName = '" & strCMName & "'"
        '�I�ں���
        If strPTCode <> "" Then strSQL = strSQL & ",PTCode = " & strPTCode & ", PTName = '" & strPTName & "'"
        
        If Not GetRS(rsTmp, "Select * From " & strOwnerName & "So033 " & GetWhereCondition(strBillNo), gcnGi) Then Exit Function
        If Not ExecuteCommand(strSQL & GetWhereCondition(strBillNo), gcnGi, lngAffected) Then Exit Function
        Do While Not rsTmp.EOF
            If Not AlterSO003(rsTmp) Then Exit Function
            rsTmp.MoveNext
        Loop
        lngTotalAmt = lngTotalAmt + GetRealAmt(strReadLine)
        lngTotalCount = lngTotalCount + 1
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
        Select Case Left(strReadLine, 3)
            Case "NSD"
                Select Case Mid(strReadLine, 65, 2)
                    Case "00" '                00 = ��b���\
                        lngBankTotalAmt = lngBankTotalAmt + GetRealAmt(strReadLine)
                        lngBankTotalCount = lngBankTotalCount + 1
                        strErrDescription = ""
                    Case "01" '                01 = �s�ڤ���
                        strErrDescription = "�s�ڤ��� !! "
                    Case "02" '                02 = �D�e�U�Nú�N�o��
                        strErrDescription = "�D�e�U�Nú�N�o�� !! "
                    Case "03" '                03 = �w�פ�e�U�Nú��
                        strErrDescription = "�w�פ�e�U�Nú�� !! "
                    Case "04" '                04 = �L���b��
                        strErrDescription = "�L���b�� !! "
                    Case "05" '                05 = �����̲νs���~
                        strErrDescription = "�����̲νs���~ !! "
                    Case "06" '                06 = �L���Τḹ�X
                        strErrDescription = "�L���Τḹ�X !! "
                    Case "07" '                07 = �Τḹ�X����
                        strErrDescription = "�Τḹ�X���� !! "
                    Case "22" '                22 = �b��w���M
                        strErrDescription = "�b���w���M !! "
                    Case "99"
                        strErrDescription = "��L�����\��] !! "
                End Select
                '�Ȧ椧���~
                If strErrDescription <> "" Then
                    Call InsertToErr(strReadLine, strErrDescription, strBankName)
                    Exit Function
                End If
                
                '�{���ˬd�����~
                strBillNo = GetBillNo(strReadLine)
'                If lngPara24 = 0 Then
'                    If strServiceType <> Mid(strBillNo, 8, 1) Then
'                        Call InsertToErr(strReadLine, "��J(���G)��ƫD���A�����O !! ", strBankName)
'                        Exit Function
'                    End If
'                End If
                If Not GetRS(rsTmp, "Select Nvl(Sum(ShouldAmt),0) as TotalAmt From " & strOwnerName & "So033 " & GetWhereCondition(strBillNo)) Then
                    BigError = True
                    Exit Function
                End If
                If rsTmp.RecordCount = 0 Then
                    Call InsertToErr(strReadLine, "�L����ڽs���Φ��渹�w���� !! ", strBankName)
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
            Case "BOF"
'                TotalLog = False
'                If Mid(strReadLine, 27, 1) <> 2 Then
'                    MsgBox "����Ƥ�����J(���G)���榡 !! ", vbCritical, gimsgWarning
'                    BigError = True
'                End If
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
                MsgBox "�榡�����D!!"
                Exit Function
        End Select
        ChkDataOk = True
    Exit Function
ChkErr:
    ErrSub FormName, "ChkDataOk"
End Function

Private Function InsertToErr(strReadLine As String, strErrDescription As String, strBankName As String) As Boolean
    On Error Resume Next
    Dim strBillNo As String
    Dim strAccountNo As String
        strBillNo = GetBillNo(strReadLine)
        strAccountNo = GetAccountNo(strReadLine)
'        WriteTextLine ErrFile, strBillNo & " " & strErrDescription
        WriteTextLine ErrFile, "�Ȧ�W��: " & strBankName & "; ���O�渹:" & strBillNo & "; ���~�W��:" & strErrDescription & "; �b��:" & strAccountNo
        lngErrCount = lngErrCount + 1
        lngErrAmt = lngErrAmt + GetRealAmt(strReadLine)
        InsertToErr = True
End Function

Private Function ChkErrCustId(rs As ADODB.Recordset, _
        ByRef lngErrCount As Long) As Boolean
    On Error GoTo ChkErr
    Dim rsTmp As New ADODB.Recordset
        If Not GetRS(rsTmp, "Select CustName,MailAddress From " & strOwnerName & "So001 Where CustId = " & rs("CustId")) Then Exit Function
        If InStr(1, rsTmp("MailAddress") & "", "��") = 0 Then
           WriteTextLine ErrFile, rs("CustId") & rsTmp("CustName") & rs("BillNo") & rsTmp("MailAddress")
           lngErrCount = lngErrCount + 1
           Exit Function
           '���D��
        End If
        ChkErrCustId = True
    Exit Function
ChkErr:
    ErrSub FormName, "ChkErrCustId"
End Function

Private Function OpenData() As Boolean
    On Error GoTo ChkErr
        If Not OpenFile(DataFile, GetDataFileName(strFilePath), 0, True, giOpenTEXT) Then Exit Function
        If Not OpenFile(ErrFile, GetErrFileName(strErrPath), 1, True) Then Exit Function
        If Not OpenFile(MemoFile, GetMemoFileName(strErrPath), 2, True) Then Exit Function
        lngRefNo = Val(GetRsValue("Select RefNo From " & strOwnerName & "CD018 Where CodeNo " & strBankSQL) & "")
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

'    file(0).Close
'    file(1).Close
    Set fso = Nothing
        If KillFile Then
            Kill GetDataFileName(strFilePath)
            Kill GetErrFileName(strErrPath)
            Kill GetMemoFileName(strErrPath)
        End If
'        fso = Nothing
    Exit Function
ChkErr:
    ErrSub FormName, "CloseFS"
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

Private Function GetWhereCondition(strBillNo As String) As String
    On Error Resume Next

    Dim strSQL As String
        If lngPara24 = 1 Then
            strSQL = "MediaBillNo = '" & strBillNo & "'"
        Else
            strSQL = "BillNo = '" & strBillNo & "'"
        End If
        If strCitemQry <> "" Then strSQL = strSQL & IIf(strSQL = "", "", " And ") & " CitemCode In (" & strCitemQry & ")"
        GetWhereCondition = " Where " & strSQL & IIf(strSQL = "", "", " And ") & "RealDate Is Null And CancelFlag=0 And UCCode is Not Null And ShouldAmt > 0"
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
        GetAccountNo = Mid(strReadLine, 27, 14)
End Function

Private Function GetBankId(strBillNo As String, ByRef BankName As String) As String
    On Error Resume Next
    Dim lngCustId As Long
    Dim rs As New ADODB.Recordset
'        lngCustId = GetRS(rs, "Select CustId From " & strOwnerName & "So033 Where BillNo = '" & strBillNo & "'") & ""
'        If Not GetRS(rs, "Select BankCode,BankName From " & strOwnerName & "So002 Where CustId = " & lngCustId) Then Exit Function

        If Not GetRS(rs, "Select CustId,BankCode,BankName From " & strOwnerName & "So033 Where BillNo = '" & strBillNo & "'") Then Exit Function
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
