Attribute VB_Name = "UpdateFun"
Option Explicit
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Const gMsgIsDataOK = "�������n���,������ !! "
Public lngCount As Long           '�O�����\����
Public lngErrCount As Long        '�O�����~����
Public lngStatusCount As Long     '�O���D���`�ᵧ��
Public lngTime As Long             '�O���Ҫ�ɶ�
Public strNowTime As String        '�O���ثe�ɶ�
Public fso As New FileSystemObject
Public File As TextStream
Public ErrFile As TextStream     '���~���log��
Private intPara3 As Integer
Private intPara5 As Integer
Private intPara8 As Integer
Private intPara9 As Integer
Public strUserName As String
Private strFilePath As String
Public objStorePara As clsStoreParameter
Public strOwnerName As String
'Public GetOwner As String

'�����W�h�Ǫ��ɮ׸��|
Public Function GetFilePath(ByVal vFilePath As String)
  On Error GoTo ChkErr
     strFilePath = vFilePath
  Exit Function
ChkErr:
  ErrSub "DealFun", "GetFilePath"
End Function

Public Sub ErrLog(ByVal ErrString As String)
    On Error Resume Next
    Dim LogFile As String
    Dim fnum As Integer
    
    'LogFile = ReadFromINI(GetIniPath1, "GICMISPath", "ErrLogPath")
    LogFile = strFilePath
    If Not ChkDirExist(LogFile) Then MkDir LogFile
    
    If Right(LogFile, 1) <> "\" Then LogFile = LogFile & "\"
    LogFile = LogFile & "SYSERR.TXT"
    
    ErrString = Replace(ErrString, vbCrLf & vbCrLf, vbCrLf)
    ErrString = "�o�Ϳ��~���t�� : �a�Q�ȪA��B�޲z�t��" & vbCrLf & _
                "�o�Ϳ��~���q�� : " & CreateObject("WScript.NetWork").ComputerName & vbCrLf & _
                "�ϥΪ̦W�� : " & strUserName & vbCrLf & ErrString
    
    fnum = FreeFile
    Open LogFile For Append As fnum
    Print #fnum, ErrString & vbCrLf & String(66, "-") & vbCrLf
    Close fnum
End Sub

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

Public Function Decrypt(EncryptionString As String) As String
  On Error GoTo ChkErr
    Decrypt = CreateObject("DevPowerEncrypt.EnCrypt").Decrypt(EncryptionString)
  Exit Function
ChkErr:
    Call ErrSub("DealFun", "Function Decrypt")
    
End Function

Public Function ChkData(ByVal vBillNo As String, ByVal vRealAmt As String, _
                        ByVal vClctEn As String, ByVal vClctName As String, _
                        ByVal vRealDate As String, ByVal vUpdEn As String, _
                        ByVal vCMCode As String, ByVal vCMName As String, _
                        ByVal vPTCode As String, ByVal vPTName As String, _
                        ByVal vCM As String, _
                        Optional blnBillNo As Boolean = False, _
                        Optional blnMediaBillNo As Boolean = False, _
                        Optional strCitemQry As String, _
                        Optional intpara24 As Integer, _
                        Optional intAccLong As Integer, _
                        Optional strOld_billno As String, _
                        Optional lngCustId As Long, Optional strCustName As String, Optional strTel1 As String, Optional strRealAmt As String) As Boolean
  On Error GoTo ChkErr
    Dim strSQL As String
    Dim rsChk As New ADODB.Recordset
    Dim rsCust As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim strErr As String
    Dim strSQL2 As String
    Dim strRealDate As String
    '���O����
        If strCitemQry <> "" Then strSQL2 = " And CitemCode In (" & strCitemQry & ")"
        If intpara24 = 0 Then
            If blnBillNo Then
                strSQL = "Select Nvl(sum(ShouldAmt),0) as TotalShouldAmt, Nvl(sum(RealAmt),0) as TotalRealAmt, RealDate, CustId, ServiceType From " & strOwnerName & "SO033 Where UCCODE IS NOT NULL AND BillNo = '" & vBillNo & "' And CancelFlag <> 1 " & strSQL2 & " Group By CustId,RealDate,ServiceType"
            Else
                strSQL = "Select Nvl(sum(ShouldAmt),0) as TotalShouldAmt, Nvl(sum(RealAmt),0) as TotalRealAmt, RealDate, CustId, ServiceType From " & strOwnerName & "SO033 Where UCCODE IS NOT NULL AND CitiBankATM like '" & Mid(vBillNo, 1, 15) & "%' And CancelFlag <> 1 " & strSQL2 & " Group By CustId,RealDate,ServiceType"
            End If
        Else
            If blnBillNo Then
                If blnMediaBillNo Then
                    strSQL = "Select Nvl(sum(ShouldAmt),0) as TotalShouldAmt, Nvl(sum(RealAmt),0) as TotalRealAmt, RealDate, CustId ,MediabillNo From " & strOwnerName & "SO033 Where UCCODE IS NOT NULL AND MediaBillNo = '" & vBillNo & "' And CancelFlag <> 1 " & strSQL2 & " Group By CustId,RealDate,MediabillNo"
                Else
                    strSQL = "Select Nvl(sum(ShouldAmt),0) as TotalShouldAmt, Nvl(sum(RealAmt),0) as TotalRealAmt, RealDate, CustId From " & strOwnerName & "SO033 Where UCCODE IS NOT NULL AND BillNo = '" & vBillNo & "' And CancelFlag <> 1 " & strSQL2 & " Group By CustId,RealDate"
                End If
            Else
                strSQL = "Select Nvl(sum(ShouldAmt),0) as TotalShouldAmt, Nvl(sum(RealAmt),0) as TotalRealAmt, RealDate, CustId From " & strOwnerName & "SO033 Where UCCODE IS NOT NULL AND CitiBankATM like '" & Mid(vBillNo, 1, 15) & "%' And CancelFlag <> 1 " & strSQL2 & " Group By CustId,RealDate"
            End If
        End If
        
        If Not GetRS(rsChk, strSQL, gcnGi) Then
           ChkData = False
           Exit Function
        End If
        
        '�p�G�O����H�U�J�bLog����ڽs������l�J�b���X 921104 Liga
        '�d�L�渹��,�p�G�w�g�ꦬLog�ꦬ���
        If rsChk.EOF Or rsChk.BOF Then
              If intpara24 = 0 Then
                 If blnBillNo Then
                    strRealDate = GetRsValue("SELECT RealDate FROM " & strOwnerName & "SO033 WHERE BillNo='" & strOld_billno & "'") & ""
                 Else
                    strRealDate = GetRsValue("SELECT RealDate FROM " & strOwnerName & "SO033 WHERE CitiBankATM like '%" & Mid(strOld_billno, 1, 15) & "%'") & ""
                 End If
              Else
                 If blnBillNo Then
                    If blnMediaBillNo Then
                       strRealDate = GetRsValue("SELECT RealDate FROM " & strOwnerName & "SO033 WHERE MediaBillNo='" & strOld_billno & "'") & ""
                    Else
                       strRealDate = GetRsValue("SELECT RealDate FROM " & strOwnerName & "SO033 WHERE BillNo='" & strOld_billno & "'") & ""
                    End If
                 Else
                    strRealDate = GetRsValue("SELECT RealDate FROM " & strOwnerName & "SO033 WHERE CitiBankATM like'%" & Mid(strOld_billno, 1, 15) & "%'") & ""
                 End If
              End If

           If strRealDate = "" Then
'                ErrFile.WriteLine ("��ڽs���G " & strOld_billno & " �d�L���渹�I" & "  �N�����O�B�G " & vCM)
               '#2318 950419 �վ�log�W��
               ErrFile.WriteLine ("" & strOld_billno & vbTab & lngCustId & vbTab & strCustName & vbTab & strTel1 & vbTab & strRealDate & vbTab & strRealAmt & vbTab & " �d�L���渹�I& �N�����O�B�G " & vCM)
           Else
'                ErrFile.WriteLine ("��ڽs���G " & strOld_billno & " �ꦬ����G" & strRealDate)
               '#2318 950419 �վ�log�W��
               ErrFile.WriteLine ("" & strOld_billno & vbTab & lngCustId & vbTab & strCustName & vbTab & strTel1 & vbTab & strRealDate & vbTab & strRealAmt & "")
           End If
           lngErrCount = lngErrCount + 1
        Else
            If rsChk("RealDate").Value & "" <> "" And rsChk("TotalRealAmt").Value > 0 Then
'                ErrFile.WriteLine ("��ڽs���G " & strOld_billno & " �Ȥ�s���G" & rsChk("CustId") & " �ɮפ��w�ꦬ���B$ " & rsChk("TotalRealAmt").Value & " �A�Ьd�֡I" & "  �N�����O�B�G " & vCM)
               '#2318 950419 �վ�log�W��
               ErrFile.WriteLine ("" & strOld_billno & vbTab & lngCustId & vbTab & strCustName & vbTab & strTel1 & vbTab & strRealDate & vbTab & strRealAmt & vbTab & " �ɮפ��w�ꦬ���B$ " & rsChk("TotalRealAmt").Value & " �A�Ьd�֡I" & "  �N�����O�B�G " & vCM)
                lngErrCount = lngErrCount + 1
            Else
                If rsChk("TotalShouldAmt").Value <> Val(vRealAmt) Then
'                    ErrFile.WriteLine ("��ڽs���G " & strOld_billno & " �Ȥ�s���G" & rsChk("CustId") & " ���B���X�A�ɮ��������B$ " & rsChk("TotalShouldAmt").Value & " �A�Ϥ��w�����B$ " & Val(vRealAmt) & "  �N�����O�B�G " & vCM & "  ���O�渹�G" & strBillNo_Old)
                    '#2318 950419 �վ�log�W��
                    ErrFile.WriteLine ("" & strOld_billno & vbTab & lngCustId & vbTab & strCustName & vbTab & strTel1 & vbTab & strRealDate & vbTab & strRealAmt & vbTab & "���B���X�A�ɮ��������B$ " & rsChk("TotalShouldAmt").Value & " �A�Ϥ��w�����B$ " & Val(vRealAmt) & "  �N�����O�B�G " & vCM)
                    lngErrCount = lngErrCount + 1
                Else
                    strSQL = "Select SO002.CustStatusCode,SO002.WipCode3,So002.ServiceType,SO033.BudgetPeriod,SO033.CitemCode,SO033.CompCode,SO033.FaciSeqNo,SO033.RealPeriod,SO033.CustId From " & strOwnerName & "So002," & strOwnerName & "SO033 Where SO002.CustId = SO033.CustId And SO002.CompCode = SO033.CompCode And SO002.ServiceType=SO033.ServiceType "
                    
                    If intpara24 = 0 Then
                        If blnBillNo Then
                           strSQL = strSQL & " And BillNo='" & strOld_billno & "'"
                        Else
                           strSQL = strSQL & " And CitiBankATM like '%" & Mid(strOld_billno, 1, 15) & "%'"
                        End If
                        
                        'strSQL = "Select CustStatusCode,WipCode3,ServiceType From " & strOwnerName & "So002 Where CustId = " & rsChk("CustId").Value & " And ServiceType= '" & rsChk("ServiceType") & "'"
                    Else
                        If blnBillNo Then
                            If blnMediaBillNo Then
                                strSQL = strSQL & " And SO033.MediaBillNo = '" & vBillNo & "'"
                            Else
                                strSQL = strSQL & " And SO033.BillNo = '" & vBillNo & "'"
                            End If
                        Else
                            strSQL = strSQL & " And SO033.CitiBankATM like '" & Mid(vBillNo, 1, 15) & "%'"
                        End If
                    End If
                    If Not GetRS(rsCust, strSQL, gcnGi) Then Exit Function
                    Do While Not rsCust.EOF
                        '�D���`�Ȥ�h���@log�A�J�b
                        If rsCust("CustStatusCode").Value & "" <> 1 Then
'                            ErrFile.WriteLine ("�Ȥ�s��: " & rsChk("CustId").Value & " ��ڽs���G " & strOld_billno & " ���D���`�Ȥ� " & " �Ȥ᪬�A: " & rsCust("CustStatusCode").Value & " �N�����O�B: " & vCM & " �A�ȧO: " & rsCust("ServiceType").Value)
                            '#2318 950419 �վ�log�W��
                            ErrFile.WriteLine ("" & strOld_billno & vbTab & lngCustId & vbTab & strCustName & vbTab & strTel1 & vbTab & strRealDate & vbTab & strRealAmt & vbTab & " ���D���`�Ȥ� " & " �Ȥ᪬�A: " & rsCust("CustStatusCode").Value & " �N�����O�B: " & vCM & " �A�ȧO: " & rsCust("ServiceType").Value)
                            lngStatusCount = lngStatusCount + 1
                        '���`�Ȥ�BCM������
                        ElseIf rsCust("CustStatusCode").Value & "" = 1 And rsCust("WipCode3") = 16 Then
'                            ErrFile.WriteLine ("�Ȥ�s��: " & rsChk("CustId").Value & " ��ڽs���G " & strOld_billno & " ��CM�������Ȥ� " & " �N�����O�B: " & vCM & " �A�ȧO: " & rsCust("ServiceType").Value)
                            '#2318 950419 �վ�log�W��
                            ErrFile.WriteLine ("" & strOld_billno & vbTab & lngCustId & vbTab & strCustName & vbTab & strTel1 & vbTab & strRealDate & vbTab & strRealAmt & vbTab & " ��CM�������Ȥ� " & " �N�����O�B: " & vCM & " �A�ȧO: " & rsCust("ServiceType").Value)
                            lngStatusCount = lngStatusCount + 1
                        End If
                        If Not GetRS(rsTmp, "Select BudgetPeriod From " & strOwnerName & "SO003 Where Custid = " & rsCust("CustId") & " And CompCode = " & rsCust("CompCode") & " And ServiceType = '" & rsCust("ServiceType") & "' " & _
                            " And CitemCode = " & rsCust("CitemCode") & " And FaciSeqNo = '" & rsCust("FaciSeqNo") & "'") Then Exit Function
                        If Not rsTmp.EOF Then
                            If rsCust("RealPeriod") > rsTmp.Fields("BudgetPeriod") And rsTmp.Fields("BudgetPeriod") > 0 Then
'                                ErrFile.WriteLine "�Ȥ�s��: " & rsChk("CustId").Value & " ��ڽs���G " & strOld_billno & "�J�b���Ƥj������Ѿl����,�L�k�J�b"
                                '#2318 950419 �վ�log�W��
                                ErrFile.WriteLine ("" & strOld_billno & vbTab & lngCustId & vbTab & strCustName & vbTab & strTel1 & vbTab & strRealDate & vbTab & strRealAmt & vbTab & " �J�b���Ƥj������Ѿl����,�L�k�J�b")
                                lngErrCount = lngErrCount + 1
                                GoTo XX
                            End If
                        End If
                        
                        rsCust.MoveNext
                    Loop
                    '�����sSO033
                    If UpdAccount(vBillNo, vRealAmt, vClctEn, vClctName, vRealDate, vUpdEn, vCMCode, vCMName, vPTCode, vPTName, blnBillNo, blnMediaBillNo, strCitemQry) = False Then
                        ChkData = False
                        Exit Function
                    End If
                End If
           End If
        End If
XX:
        If rsChk.State = adStateOpen Then rsChk.Close
        Set rsChk = Nothing
        If rsCust.State = adStateOpen Then rsCust.Close
        Set rsCust = Nothing
        
        ChkData = True
    Exit Function
ChkErr:
    ErrSub "DealFun", "ChkData"
End Function

Public Function UpdAccount(ByVal vBillNo As String, ByVal vRealAmt As String, _
                           ByVal vClctEn As String, ByVal vClctName As String, _
                           ByVal vRealDate As String, ByVal vUpdEn As String, _
                           ByVal vCMCode As String, ByVal vCMName As String, _
                           ByVal vPTCode As String, ByVal vPTName As String, _
                           Optional blnBillNo As Boolean = False, _
                           Optional blnMediaBillNo As Boolean = False, Optional strCitemQry As String) As Boolean
    On Error GoTo ChkErr
    Dim rsAccount As New ADODB.Recordset
    Dim strSQL As String, strSeqQry As String, lngAffected As Long
    Dim lngPeriod As Long, lngAmount As Long
    Dim strSo003 As String, strUpdSql As String
    Dim strSQL2 As String, rsTmp As New ADODB.Recordset
        '���O����
        If strCitemQry <> "" Then strSQL2 = " And CitemCode In (" & strCitemQry & ")"
        If blnBillNo Then
            If blnMediaBillNo Then
             strSQL = "Select RowId, So033.* From " & strOwnerName & "So033 Where UCCODE IS NOT NULL  And MediaBillNo ='" & vBillNo & "' AND CancelFlag <> 1" & strSQL2
'            strSQL = "Select RowId, CustId, ShouldAmt, CitemCode, RealStartDate, RealStopDate, RealPeriod, OldPeriod, OldAmt, ClctEn, ClctName, CMCode, CMName, SeqNo, PackageNo, PackageName, BudgetPeriod From " & strOwnerName & "So033 Where UCCODE IS NOT NULL  And MediaBillNo ='" & vBillNo & "' AND CancelFlag <> 1" & strSQL2
            Else
                strSQL = "Select RowId,So033.* From " & strOwnerName & "So033 Where UCCODE IS NOT NULL  And BillNo ='" & vBillNo & "' AND CancelFlag <> 1" & strSQL2
            End If
        Else
            
            strSQL = "Select RowId,So033.* From " & strOwnerName & "So033 Where UCCODE IS NOT NULL  And CitiBankATM like '" & Mid(vBillNo, 1, 15) & "%' AND CancelFlag <> 1" & strSQL2
        End If
        
        If Not GetRS(rsAccount, strSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then
            UpdAccount = False
            Exit Function
        End If
        
        If rsAccount.RecordCount > 0 Then rsAccount.MoveFirst
        While Not rsAccount.EOF
            With rsAccount
                .Fields("RealAmt") = rsAccount("ShouldAmt")
                .Fields("RealDate") = NoZero(Format(vRealDate, "####/##/##"))
                .Fields("Updtime") = GetDTString(strNowTime)
                .Fields("UpdEn") = vUpdEn
                .Fields("UCCode") = Null
                .Fields("UCName") = Null
                .Fields("CMCode") = IIf(vCMCode = "", rsAccount("CMCode"), vCMCode)
                .Fields("CMName") = IIf(vCMName = "", rsAccount("CMName"), vCMName)
                .Fields("PTCode") = IIf(vPTCode = "", Null, vPTCode)
                .Fields("PTName") = NoZero(vPTName)
                .Fields("ClctEn") = IIf(Trim(vClctEn) = "", rsAccount("ClctEn"), vClctEn)
                .Fields("ClctName") = IIf(Trim(vClctName) = "", rsAccount("ClctName"), vClctName)
               .Fields("FirstTime") = GetDTString(strNowTime)
                On Error Resume Next
                Call AlterSO003(rsAccount)
                .Update
            End With
        
'            strSQL = "Update " & strOwnerName & "So033 Set RealAmt =" & rsAccount("ShouldAmt") & _
'                    ",ClctEn = '" & IIf(Trim(vClctEn) = "", rsAccount("ClctEn"), vClctEn) & "'" & _
'                    ",ClctName = '" & IIf(Trim(vClctName) = "", rsAccount("ClctName"), vClctName) & "'" & _
'                    ",RealDate = To_date('" & vRealDate & "','YYYYMMDD')" & _
'                    ",Updtime = '" & GetDTString(strNowTime) & "'" & _
'                    ",UpdEn = '" & vUpdEn & "'" & _
'                    ",UCCode = Null, UCName = Null" & _
'                    ", CMCode = '" & IIf(vCMCode = "", rsAccount("CMCode"), vCMCode) & "'" & _
'                    ", CMName = '" & IIf(vCMName = "", rsAccount("CMName"), vCMName) & "'" & _
'                    ", PTCode = " & GetNullString(vPTCode) & _
'                    ", PTName = " & GetNullString(vPTName) & _
'                    ",FirstTime = '" & GetDTString(strNowTime) & "'" & _
'                    " Where Rowid = '" & rsAccount("Rowid") & "'"
'            gcnGi.Execute strSQL
            
'            If Not GetRS(rsTmp, "Select * From " & strOwnerName & "SO033 Where RowId = '" & rsAccount("RowId") & "'") Then Exit Function
'            If AlterSO003(rsTmp) = False Then
'                   UpdAccount = False
'                   Exit Function
'            End If
lLoop:
            rsAccount.MoveNext
        Wend
            
            
'            If intPara8 >= 1 Then
'                lngPeriod = NoZero(rsAccount("RealPeriod"), True)
'                If intPara9 = 1 Then
'                    lngAmount = NoZero(rsAccount("OldAmt"), True)
'                    lngPeriod = NoZero(rsAccount("OldPeriod"), True)
'                ElseIf intPara9 = 2 Then
'                    lngAmount = NoZero(rsAccount("ShouldAmt"), True)
'                Else
'                    lngAmount = NoZero(rsAccount("ShouldAmt"), True)
'                End If
'                strUpdSql = ",Amount = " & lngAmount & ",Period = " & lngPeriod
'            End If
'
'
'            If rsAccount("SeqNo") & "" <> "" Then strSeqQry = " And SeqNo = " & rsAccount("SeqNo")
'
'            If Val(rsAccount("BudgetPeriod") & "") > 0 Then
'                If Val(rsAccount("BudgetPeriod") & "") = 1 Then
'                    If Not ExecuteCommand("Delete From SO003 Where CustId = '" & rsAccount("CustId") & "' And CitemCode = '" & rsAccount("CitemCode") & "' And StopDate < " & _
'                        GetNullString(rsAccount("RealStopDate"), giDateV, giOracle) & strSeqQry) Then Exit Function
'                    GoTo lLoop
'                End If
'                strSo003 = ",BudgetPeriod = " & (rsAccount("BudgetPeriod") - 1)
'            End If
'
'            strSQL = "Update " & strOwnerName & "So003 Set StartDate = " & GetNullString(rsAccount("RealStartDate"), giDateV, giOracle) & _
'                    ",StopDate = " & GetNullString(rsAccount("RealStopDate"), giDateV, giOracle) & _
'                    ",ClctDate = " & GetNullString(rsAccount("RealStopDate"), giDateV, giOracle) & " + " & intPara5 & _
'                    strUpdSql & ",UpdEn = '" & vUpdEn & "',UpdTime = '" & GetDTString(strNowTime) & "'" & strSo003 & _
'                    " Where CustId = '" & rsAccount("CustId") & "' And CitemCode = '" & rsAccount("CitemCode") & "' And StopDate < " & _
'                    GetNullString(rsAccount("RealStopDate"), giDateV, giOracle)
'
'            gcnGi.Execute strSQL & strSeqQry, lngAffected
'            If strSeqQry <> "" And lngAffected = 0 Then
'                 gcnGi.Execute strSQL
'            End If

        lngCount = lngCount + 1
        Call CloseRecordset(rsAccount)
        Call CloseRecordset(rsTmp)
        UpdAccount = True
    Exit Function
ChkErr:
    UpdAccount = False
End Function

Public Function InitData() As Boolean
On Error GoTo ChkErr
   Dim rsTmp As New ADODB.Recordset
   Dim strSQL As String
   
   strSQL = "Select Para3, Para5, Para8, Para9 From " & strOwnerName & "So043"
   If Not GetRS(rsTmp, strSQL, gcnGi) Then InitData = False: Exit Function
   If rsTmp.EOF = False Then
      intPara3 = rsTmp("Para3")
      intPara5 = rsTmp("Para5")
      intPara8 = rsTmp("Para8")
      intPara9 = rsTmp("Para9")
   End If
   Set rsTmp = Nothing
   InitData = True
 Exit Function
ChkErr:
  ErrSub "DealFun", "InitData"
  InitData = False
End Function

Public Function MsgResult()
  On Error GoTo ChkErr
    lngTime = Timer - lngTime
    MsgBox "��b�J�b���Ʀ@�G" & lngCount + lngErrCount & "��" & vbCrLf & _
                   "���\�@�G" & lngCount & "��" & vbCrLf & _
                   "���Ѧ@�G" & lngErrCount & "��" & vbCrLf & _
                   "�D���`��@�G" & lngStatusCount & "��" & vbCrLf & _
                   "�@��O�G" & lngTime & "��", vbInformation, "����"
  Exit Function
ChkErr:
  ErrSub "DealFun", "MsgResult"
End Function

Public Sub SetLst(objLst As Object, _
                  strFldName1 As String, _
                  strFldName2 As String, _
                  Optional lngFldLen1 As Long, _
                  Optional lngFldLen2 As Long, _
                  Optional lngWidth1 As Long, _
                  Optional lngWidth2 As Long, _
                  Optional StrTableName As String, _
                  Optional strWhere As String, Optional blnStopFlag As Boolean = False)
  On Error GoTo ChkErr
    objLst.SetFldName1 strFldName1
    objLst.SetFldName2 strFldName2
    If lngFldLen1 > 0 Then objLst.FldLen1 = lngFldLen1
    If lngFldLen2 > 0 Then objLst.FldLen2 = lngFldLen2
    If lngWidth1 > 0 Then objLst.FldWidth1 = lngWidth1
    If lngWidth2 > 0 Then objLst.FldWidth2 = lngWidth2
    objLst.SetTableName strOwnerName & StrTableName
    objLst.FilterStop = blnStopFlag
    objLst.SendConn gcnGi
    objLst.Filter = strWhere
  Exit Sub
ChkErr:
    Call ErrSub("DealFun", "SetLst")
End Sub

Public Function ChkGiList(Optional KeyCode As Integer) As Boolean
  On Error GoTo ChkErr
    With Screen
        If TypeOf .ActiveControl Is GiList Then
            If Len(Trim(.ActiveControl.GetCodeNo & "")) = 0 And Len(Trim(.ActiveControl.GetDescription & "")) = 0 Then ChkGiList = True: Exit Function
            ChkGiList = True
            If .ActiveControl.F5Corresponding Then
                If KeyCode <> vbKeyF5 Then
                   ChkGiList = True
                Else
                   ChkGiList = .ActiveControl.DataMatch
                   If KeyCode = vbKeyF5 Then Exit Function
                End If
            End If
            If .ActiveControl.F3Corresponding Then
                If KeyCode <> vbKeyF3 Then
                   ChkGiList = True
                Else
                   ChkGiList = .ActiveControl.DataMatch
                   If KeyCode = vbKeyF3 Then Exit Function
                End If
            End If
            If .ActiveControl.F2Corresponding Then
                If KeyCode <> vbKeyF2 Then
                   ChkGiList = True
                Else
                   ChkGiList = .ActiveControl.DataMatch
                   If KeyCode = vbKeyF2 Then Exit Function
                End If
            End If
        Else
            ChkGiList = True
        End If
    End With
  Exit Function
ChkErr:
    ChkGiList = True
    Resume Next
'   ErrSub "ExtraModule", "ChkGiList"
End Function

Public Function GetBillNo_New(ByVal strNo As String, Optional blnMediaBillNo As Boolean) As String
    On Error Resume Next
    Dim strType As String
    Dim strYM As String
    Dim strMM As String
    
'        strNo = Mid(strReadLine, 65, 11)
        If blnMediaBillNo Then
            GetBillNo_New = strNo
        Else
            strYM = Left(strNo, 2)
            strMM = Mid(strNo, 3, 1)
            If strMM > "9" Then
                strYM = strYM & Format(Trim(Asc(strMM) - Asc("A") + 10), "00")
            Else
                strYM = strYM & Format(strMM, "00")
            End If
            Select Case Mid(strNo, 4, 1)
                Case 1
                    strType = "BC"
                Case 2
                    strType = "TC"
                Case 3
                    strType = "BI"
                Case 4
                    strType = "TI"
            End Select
            
            GetBillNo_New = Trim(CStr(Left(strYM, 4) + 191100)) & _
                    strType & Format(Right(strNo, 7), "0000000")
        End If
End Function

Public Function GetBillNo_Old(ByVal strNo As String, ByVal strServiceType, _
    Optional blnMediaBillNo As Boolean) As String
    On Error Resume Next
        If blnMediaBillNo Then
            GetBillNo_Old = strNo
        Else
            GetBillNo_Old = Trim(CStr(Val(Left(strNo, 4)) + 191100)) & _
                    Mid(strNo, 5, 1) & strServiceType & Format(Right(strNo, 6), "0000000")
        End If
End Function

Public Function GetBankATM(strBillNo As String, lngCustId As Long, strBankCode As String, _
    strShouldDate As String, lngShouldAmt As Long, Optional ByRef clsBankATMNo As Object, _
    Optional ByRef strCorpId As String) As String
    On Error Resume Next
    Dim rsTmp As New ADODB.Recordset
    Dim strShDate As String
        If clsBankATMNo Is Nothing Then
            If strBankCode = "" Or strBillNo = "" Or lngCustId = 0 Then Exit Function
            If Not GetRS(rsTmp, "Select FileName2,CorpID From " & strOwnerName & "CD018 Where CodeNo = " & strBankCode) Then Exit Function
            If Len(rsTmp("FileName2") & "") = 0 Or Len(rsTmp("CorpId") & "") = 0 Then MsgBox "�Ȧ�N�X�ɤ��L�w�q��b��J�ɦW�Φ�����N��", vbCritical, "ĵ�i": Exit Function
            strCorpId = rsTmp("CorpId") & ""
            Set clsBankATMNo = CreateObject("csBankATMNo.cls" & rsTmp("FileName2"))
            If Err.Number = 429 Then MsgBox "�Ȧ�N�X�ɤ��w�q����b��J�ɦW : [" & rsTmp("FileName2") & "] �����T,�нT�{!!", vbCritical, "ĵ�i": Exit Function
        End If
        strShDate = Format(strShouldDate, "yyyy/mm/dd")
        GetBankATM = clsBankATMNo.GetATMNo(strBillNo, lngCustId, strCorpId, strShDate, lngShouldAmt, gcnGi)
End Function

'Public Function AlterSO003(rsSO033 As ADODB.Recordset, Optional ByVal blnAddSo003 As Boolean = True, _
'                Optional ByVal blnUpdAmt As Boolean = True, Optional ByVal blnUpdClctDate As Boolean = True, _
'                Optional ByVal blnChkMustAnswer As Boolean = False, Optional ByRef strSeqNo As String, _
'                Optional ByVal blnMultiCount As Boolean = False) As Boolean
'    On Error GoTo ChkErr
'    Dim rsAlter As New ADODB.Recordset
'    Dim rsTmp As New ADODB.Recordset
'    Dim lngPeriodFlag As Boolean
'    Dim Ypara5 As Long, Ypara8 As Long
'    Dim Ypara9 As Long, YPara24 As Long
'    Dim lngCompCode As Long, lngBad As Long
'    Dim blnFlag As Boolean
'    Dim blnDiscount As Boolean
'        '�P�_�O�_���g���ʦ��O����
'        lngPeriodFlag = Val(GetRsValue("Select PeriodFlag From " & strOwnerName & "Cd019 Where CodeNo = " & rsSO033("CitemCode")) & "")
'        If Len(rsSO033("STCode") & "") > 0 Then lngBad = Val(GetRsValue("Select RefNo From " & strOwnerName & "CD016 Where CodeNo = " & rsSO033("STCode")) & "")
'        If lngPeriodFlag = 0 Or (lngBad = 1) Then
'            AlterSO003 = True
'            Exit Function
'        End If
'
'        If Not GetSystemPara(rsTmp, rsSO033("CompCode") & "", rsSO033("ServiceType"), "SO043", "Para5,Para8,Para9,Para24") Then Exit Function
'        If rsTmp.RecordCount = 0 Then Exit Function
'        Ypara5 = rsTmp("Para5")
'        Ypara8 = rsTmp("Para8")
'        Ypara9 = rsTmp("Para9")
'        YPara24 = rsTmp("Para24")
'        Call CloseRecordset(rsTmp)
'
'        lngCompCode = Val(GetRsValue("Select CompCode From " & strOwnerName & "So001 Where CustId = " & rsSO033("CustId")) & "")
'
''        If rsSO033("OrderNo") & "" <> "" Then
'            If Val(rsSO033("SeqNo") & "") = 0 Then
'                If Not GetRS(rsAlter, "Select RowId,SO003.* From " & strOwnerName & "So003 Where 1 = 0 ", gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
'            Else
'                If Not GetRS(rsAlter, "Select RowId,SO003.* From " & strOwnerName & "So003 Where CustId = " & rsSO033("CustId") & " And CitemCode = " & rsSO033("CitemCode") & " And SeqNo = " & rsSO033("SeqNo"), gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
'            End If
''        End If
'        With rsAlter
'            If .RecordCount = 0 Then
'                '�ˬd�O���n�s�W�g���ʦ��O����
'                If blnAddSo003 Then
'                    If blnChkMustAnswer Then
'                        If MsgBox("�L�g���ʦ��O����,�O�_�n�s�W??", vbQuestion + vbYesNo + vbDefaultButton2, gimsgPrompt) = vbNo Then
'                            AlterSO003 = True
'                            Exit Function
'                        End If
'                    End If
'                    .AddNew
'                    If Val(rsSO033("SeqNo") & "") = 0 Then
'                        .Fields("SeqNO") = GetInvoiceNo3("SO003") & ""
'                    Else
'                        .Fields("SeqNO") = rsSO033("SeqNo")
'                    End If
'                    blnFlag = True
'                    '�����~�۰����Ȥ��s��A�� 92/03/25 �P Lawrence �Q�פ����G 92/08/12 �S��
'                    If YPara24 = 1 Then
'                        If rsSO033("AccountNO") & "" <> "" Then
'                            If Not GetRS(rsTmp, "Select Nvl(AddCitemAccount,0) AddCitemAccount,SnactionDate From " & strOwnerName & "SO106 Where CustId = " & rsSO033("CustId") & _
'                                " And AccountID = '" & rsSO033("AccountNo") & "' And Nvl(StopFlag,0)= 0") Then Exit Function
'                            If rsTmp.RecordCount > 0 Then If rsTmp("AddCitemAccount") = 0 Or rsTmp("SnactionDate") & "" = "" Then blnFlag = False
'                        End If
'                    End If
'                    If blnFlag Then
'                        .Fields("CMCode") = NoZero(rsSO033("CMCode"))
'                        .Fields("CMName") = NoZero(rsSO033("CMName"))
'                        .Fields("BankCode") = NoZero(rsSO033("BankCode"))
'                        .Fields("BankName") = NoZero(rsSO033("BankName"))
'                        .Fields("AccountNo") = NoZero(rsSO033("AccountNo"))
'                    Else
'                        If Not GetRS(rsTmp, "Select CodeNo,Description From " & strOwnerName & "CD031 Where RefNo=1", gcnGi) Then Exit Function
'                        .Fields("CMCode") = NoZero(rsTmp("CodeNo"))
'                        .Fields("CMName") = NoZero(rsTmp("Description"))
'                    End If
'
'                    .Fields("OrderNo") = NoZero(rsSO033("OrderNo"))
'                    .Fields("PTCode") = NoZero(rsSO033("PTCode"))
'                    .Fields("PTName") = NoZero(rsSO033("PTName"))
'                Else
'                    AlterSO003 = True
'                    Exit Function
'                End If
'            Else
'                blnAddSo003 = False
'            End If
'            If .Fields("StopDate") >= rsSO033("RealStopDate") Then AlterSO003 = True: Exit Function
'            If rsSO033("DiscountCode") & "" <> "" Then
'                '93/04/20 �令���u�ݬ���...�ӥB�u�ݴ����O�j��SO003���O�_�l�������SO003����
'                If Format(rsSO033("RealStartDate") & "", "yyyymmdd") <= Format(.Fields("DiscountDate2") & "", "yyyymmdd") Or blnAddSo003 Then
'                    '.Fields("Amount") = NoZero(rsSO033("OldAmt"), True)
'                    blnDiscount = True
'                    .Fields("DiscountCode") = NoZero(rsSO033("DiscountCode"))
'                    .Fields("DiscountName") = NoZero(rsSO033("DiscountName"))
'                    .Fields("DiscountDate1") = NoZero(rsSO033("DiscountDate1"))
'                    .Fields("DiscountDate2") = NoZero(rsSO033("DiscountDate2"))
'                    .Fields("DiscountAmt") = NoZero(rsSO033("DiscountAmt"))
'                    If blnAddSo003 Then
'                        .Fields("Period").Value = GetFieldValue(rsSO033, "OldPeriod")
'                        .Fields("Amount").Value = GetFieldValue(rsSO033, "OldAmt")
'                    End If
'                End If
'            End If
'
'            If blnAddSo003 Or (blnUpdAmt And Ypara8 >= 1) Then
'                If Not blnDiscount Then
'                    '�Y<Para9>=1, �h 1: Amount = OldAmt, 2: Amount = ShouldAmt , 3: Amount = RealAmt
'                    .Fields("Period").Value = GetFieldValue(rsSO033, "RealPeriod")
'                    If Ypara9 = 1 Then
'                        .Fields("Period").Value = GetFieldValue(rsSO033, "OldPeriod")
'                        .Fields("Amount").Value = GetFieldValue(rsSO033, "OldAmt")
'                    ElseIf Ypara9 = 2 Then
'                        '.Fields("Period").Value = GetFieldValue(rsSo033, "RealPeriod")
'                        .Fields("Amount").Value = GetFieldValue(rsSO033, "ShouldAmt")
'                    Else
'                        '.Fields("Period").Value = GetFieldValue(rsSo033, "RealPeriod")
'                        .Fields("Amount").Value = GetFieldValue(rsSO033, "RealAmt")
'                    End If
'                End If
'            End If
'            If rsSO033("BudgetPeriod") > 0 Then
'                .Fields("BudgetPeriod") = rsSO033("BudgetPeriod") - rsSO033("RealPeriod")
'                If .Fields("BudgetPeriod") <= 0 Then
'                    If Not ExecuteCommand("Delete From " & strOwnerName & "SO003 Where RowId= '" & .Fields("RowId") & "'") Then Exit Function
'                    AlterSO003 = True
'                    Exit Function
'                End If
'            End If
'            .Fields("CustId").Value = GetFieldValue(rsSO033, "CustID")
'            .Fields("CitemCode").Value = GetFieldValue(rsSO033, "CitemCode")
'            .Fields("CitemName").Value = GetFieldValue(rsSO033, "CitemName")
'            .Fields("StartDate").Value = GetFieldValue(rsSO033, "RealStartDate")
'            .Fields("StopDate").Value = GetFieldValue(rsSO033, "RealStopDate")
'            .Fields("UpdTime").Value = GetFieldValue(rsSO033, "UpdTime")
'            .Fields("UpdEn").Value = GetFieldValue(rsSO033, "UpdEn")
'            .Fields("ClctDate").Value = GetFieldValue(rsSO033, "RealStopDate") + Ypara5
'            .Fields("LastDate").Value = GetFieldValue(rsSO033, "RealStopDate") + Ypara5
'            .Fields("CompCode").Value = lngCompCode
'            .Fields("ServiceType").Value = GetFieldValue(rsSO033, "ServiceType")
'            .Fields("PackageNo").Value = NoZero(rsSO033("PackageNo"))
'            .Fields("PackageName").Value = NoZero(rsSO033("PackageName"))
'            .Update
'            strSeqNo = .Fields("SeqNo") & ""
'        End With
'        AlterSO003 = True
'    Exit Function
''ChkErr:
''    ErrSub FormName, "AlterSo003"
''End Function
'
'ChkErr:
'    AlterSO003 = False
'End Function

