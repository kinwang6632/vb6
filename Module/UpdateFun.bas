Attribute VB_Name = "UpdateFun"
Option Explicit
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Const gMsgIsDataOK = "�������n���,������ !! "
Public lngCount As Long           '�O�����\����
Public lngErrCount As Long        '�O�����~����
Public lngStatusCount As Long     '�O���D���`�ᵧ��
Public lngTime As Long             '�O���Ҫ�ɶ�
Public strNowTime As String        '�O���ثe�ɶ�
Public FSO As New FileSystemObject
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
Public blnSO3314 As Boolean
'Public GetOwner As String
Public blnCrossCustCombine  As Boolean
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
                        Optional strOld_billno As String, Optional strErrorMsg As String) As Boolean
  On Error GoTo ChkErr
    Dim strSQL As String
    Dim rsChk As New ADODB.Recordset
    Dim rsCust As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim strErr As String
    Dim strSQL2 As String
    Dim strBillNo_Old As String
    Dim strRealDate As String
    Dim strReturn As String
    Dim strCD013Ref As String
    Dim strCustId As String
    Dim lngTotalShouldAmt As Long
    Dim lngTotalRealAmt As Long
    strCD013Ref = " Decode(Nvl(CD013.RefNo,0),7,3,Decode(Nvl(CD013.RefNo,0),8,3," & _
            "Decode(Nvl(CD013.RefNo,0),3,3,0))) RefNo "
    
    '���O����
        If strCitemQry <> "" Then strSQL2 = " And SO033.CitemCode In (" & strCitemQry & ")"
        If intpara24 = 0 Then
            '#4010 UCCode���ѦҸ���3�n�������` By Kin 2008/07/30
            If blnBillNo Then
                strSQL = "Select Nvl(sum(SO033.ShouldAmt),0) as TotalShouldAmt, Nvl(sum(SO033.RealAmt),0) as TotalRealAmt," & _
                         "SO033.RealDate,SO033.CustId,SO033.ServiceType,Max(Decode(CD013.RefNo,3,3,0)) RefNo From " & strOwnerName & "SO033," & strOwnerName & "CD013 " & _
                         " Where SO033.UCCODE IS NOT NULL AND SO033.BillNo = '" & vBillNo & "' And SO033.CancelFlag <> 1 And SO033.UCCode=CD013.CodeNo " & strSQL2 & " Group By SO033.CustId,SO033.RealDate,SO033.ServiceType"
            Else
            '���D��3008 �NLike %����,�H�W�[�Ĳv,By Kin 2007/1/22
                'strSQL = "Select Nvl(sum(ShouldAmt),0) as TotalShouldAmt, Nvl(sum(RealAmt),0) as TotalRealAmt, RealDate, CustId, ServiceType From " & strOwnerName & "SO033 Where UCCODE IS NOT NULL AND CitiBankATM like '" & Mid(vBillNo, 1, 15) & "%' And CancelFlag <> 1 " & strSQL2 & " Group By CustId,RealDate,ServiceType"
                strSQL = "Select Nvl(sum(SO033.ShouldAmt),0) as TotalShouldAmt, Nvl(sum(SO033.RealAmt),0) as TotalRealAmt," & _
                         "SO033.RealDate,SO033.CustId,SO033.ServiceType,Max(Decode(CD013.RefNo,3,3,0)) RefNo" & _
                         " From " & strOwnerName & "SO033," & strOwnerName & "CD013 " & _
                         " Where SO033.UCCODE IS NOT NULL AND SO033.CitiBankATM = '" & vBillNo & "' And SO033.CancelFlag <> 1 And SO033.UCCode=CD013.CodeNo " & strSQL2 & _
                         " Group By SO033.CustId,SO033.RealDate,SO033.ServiceType"
            End If
        Else
            If blnBillNo Then
                If blnMediaBillNo Then
                    strSQL = "Select Nvl(sum(SO033.ShouldAmt),0) as TotalShouldAmt, Nvl(sum(SO033.RealAmt),0) as TotalRealAmt," & _
                            "SO033.RealDate,SO033.CustId,SO033.MediabillNo,Max(Decode(CD013.RefNo,3,3,0)) RefNo " & _
                            " From " & strOwnerName & "SO033," & strOwnerName & "CD013 " & _
                            " Where SO033.UCCODE IS NOT NULL AND SO033.MediaBillNo = '" & vBillNo & "' And SO033.CancelFlag <> 1  And SO033.UCCode=CD013.CodeNo " & strSQL2 & _
                            " Group By SO033.CustId,SO033.RealDate,SO033.MediabillNo"
                Else
                    strSQL = "Select Nvl(sum(SO033.ShouldAmt),0) as TotalShouldAmt, Nvl(sum(SO033.RealAmt),0) as TotalRealAmt," & _
                             "SO033.RealDate,SO033.CustId,Max(Decode(CD013.RefNo,3,3,0)) RefNo From " & strOwnerName & "SO033," & strOwnerName & "CD013 " & _
                             " Where SO033.UCCODE IS NOT NULL AND SO033.BillNo = '" & vBillNo & "' And SO033.CancelFlag <> 1 And SO033.UCCode=CD013.CodeNo " & strSQL2 & _
                             " Group By SO033.CustId,SO033.RealDate"
                End If
            Else
                '���D��3008 �NLike %����,�H�W�[�Ĳv,By Kin 2007/1/22
                'strSQL = "Select Nvl(sum(ShouldAmt),0) as TotalShouldAmt, Nvl(sum(RealAmt),0) as TotalRealAmt, RealDate, CustId From " & strOwnerName & "SO033 Where UCCODE IS NOT NULL AND CitiBankATM like '" & Mid(vBillNo, 1, 15) & "%' And CancelFlag <> 1 " & strSQL2 & " Group By CustId,RealDate"
                strSQL = "Select Nvl(sum(SO033.ShouldAmt),0) as TotalShouldAmt, Nvl(sum(SO033.RealAmt),0) as TotalRealAmt," & _
                        "SO033.RealDate,SO033.CustId,Max(Decode(CD013.RefNo,3,3,0)) RefNo From " & strOwnerName & "SO033," & strOwnerName & "CD013 " & _
                        " Where SO033.UCCODE IS NOT NULL AND SO033.CitiBankATM = '" & vBillNo & "' And SO033.CancelFlag <> 1 And SO033.UCCode=CD013.CodeNo " & strSQL2 & _
                        " Group By SO033.CustId,SO033.RealDate"
            End If
        End If
          '7179
        If blnCrossCustCombine Then
             strSQL = "Select A.*,SO001.AMduid, (Select MainCustId From " & strOwnerName & "SO202 Where MduId = SO001.AmduId) MainCustId " & _
                                 " From (" & strSQL & " ) A," & strOwnerName & "SO001 " & _
                                 " Where A.CustId = SO001.CustId "
        End If
        If Not GetRS(rsChk, strSQL, gcnGi) Then
           ChkData = False
           Exit Function
        End If
        '�p�G�O����H�U�J�bLog����ڽs������l�J�b���X 921104 Liga
        '�d�L�渹��,�p�G�w�g�ꦬLog�ꦬ���
        If rsChk.EOF Or rsChk.BOF Then
              blnSO3314 = False
              If intpara24 = 0 Then
                 If blnBillNo Then
                    strRealDate = GetRsValue("SELECT RealDate FROM " & strOwnerName & "SO033 WHERE BillNo='" & strOld_billno & "'") & ""
                 Else
                    '���D��3008 �NLike %����,�H�W�[�Ĳv,By Kin 2007/1/22
                    'strRealDate = GetRsValue("SELECT RealDate FROM " & strOwnerName & "SO033 WHERE CitiBankATM like '%" & Mid(strOld_billno, 1, 15) & "%'") & ""
                    strRealDate = GetRsValue("SELECT RealDate FROM " & strOwnerName & "SO033 WHERE CitiBankATM = '" & strOld_billno & "'") & ""
                 End If
              Else
                 If blnBillNo Then
                    If blnMediaBillNo Then
                       strRealDate = GetRsValue("SELECT RealDate FROM " & strOwnerName & "SO033 WHERE MediaBillNo='" & strOld_billno & "'") & ""
                    Else
                       strRealDate = GetRsValue("SELECT RealDate FROM " & strOwnerName & "SO033 WHERE BillNo='" & strOld_billno & "'") & ""
                    End If
                 Else
                    '���D��3008 �NLike %����,�H�W�[�Ĳv,By Kin 2007/1/22
                    'strRealDate = GetRsValue("SELECT RealDate FROM " & strOwnerName & "SO033 WHERE CitiBankATM like'%" & Mid(strOld_billno, 1, 15) & "%'") & ""
                    strRealDate = GetRsValue("SELECT RealDate FROM " & strOwnerName & "SO033 WHERE CitiBankATM ='" & strOld_billno & "'") & ""
                 End If
              End If

           If strRealDate = "" Then
                '950630 �bSO3314�J�b�ݦ^��strErrorMsg,�G��Function��������ErrFile.WriteLine�Ҷ���g
                'ErrFile.WriteLine ("��ڽs���G " & strOld_billno & " �d�L���渹�I" & "  �N�����O�B�G " & vCM)
                strErrorMsg = "��ڽs���G " & strOld_billno & " �d�L���渹�I" & "  �N�����O�B�G " & vCM
                ErrFile.WriteLine strErrorMsg
           Else
                strErrorMsg = "��ڽs���G " & strOld_billno & " �ꦬ����G" & strRealDate
                ErrFile.WriteLine strErrorMsg
           End If
           lngErrCount = lngErrCount + 1
           blnSO3314 = True
        Else
              '7179*********************************************************************************************
            lngTotalRealAmt = 0
            lngTotalShouldAmt = 0
            
            If blnCrossCustCombine Then
                rsChk.MoveFirst
                If Len(rsChk("MainCustId") & "") > 0 Then
                    strCustId = rsChk("MainCustId") & ""
                Else
                    strCustId = rsChk("CustId") & ""
                End If
                Do While Not rsChk.EOF
                    lngTotalRealAmt = lngTotalRealAmt + Val(rsChk("TotalRealAmt") & "")
                    lngTotalShouldAmt = lngTotalShouldAmt + Val(rsChk("TotalShouldAmt") & "")
                    rsChk.MoveNext
                Loop
                rsChk.MoveFirst
            Else
                lngTotalRealAmt = Val(rsChk("TotalRealAmt") & "")
                lngTotalShouldAmt = Val(rsChk("TotalShouldAmt") & "")
                strCustId = rsChk("CustId") & ""
            End If
            '*************************************************************************************************
            '***********************************************************************
            '#5386 �P�_�O�_�����P���Ƚs���C�^�渹�ۦP By Kin 2009/12/04
            If rsChk.RecordCount > 1 Then
                Dim strRetMsg As String
                strRetMsg = ChkDiffCust(rsChk)  '�ˮ֤��P���Ƚs
                If strRetMsg <> "" Then
                    strErrorMsg = "��ڽs���G " & strOld_billno & " �ۦP��ڽs�����Ȥ�s�����P�I �Ƚs�G" & strRetMsg
                    ErrFile.WriteLine (strErrorMsg)
                    
                    lngErrCount = lngErrCount + 1
                    blnSO3314 = True
                    GoTo XX
                End If
            End If
            '***********************************************************************
        
            '#3122 �ˮ֬O�_���W�L�鵲��� By Kin Add 2007/03/29
            If Not ChkOverCloseDate(Format(vRealDate, "####/##/##"), objStorePara.CompCode, objStorePara.ServiceType, strReturn) Then
                strErrorMsg = strReturn
                ErrFile.WriteLine ("��ڽs���G" & vBillNo & "  ��ڦ��ڤ�G" & Format(str(Val(vRealDate) - 19110000), "##/##/##") & "  ���~��]�G" & strReturn)
                lngErrCount = lngErrCount + 1
                blnSO3314 = True
                GoTo XX
            End If
            '********************************************************************************************************************************
            '#4010 UCCode���ѦҸ���3�n�������`��� By Kin 2008/07/30
            If Val(rsChk("RefNo") & "") = 3 Then
                strErrorMsg = "��ڽs���G " & strOld_billno & " �Ȥ�s���G" & rsChk("CustId") & " �d�O�w���I�Ьd��  �N�����O�B�G " & vCM
                ErrFile.WriteLine strErrorMsg
                lngErrCount = lngErrCount + 1
                blnSO3314 = True
                GoTo XX
            End If
            '********************************************************************************************************************************
            If rsChk("RealDate").Value & "" <> "" And lngTotalRealAmt > 0 Then
                strErrorMsg = "��ڽs���G " & strOld_billno & " �Ȥ�s���G" & strCustId & " �ɮפ��w�ꦬ���B$ " & lngTotalRealAmt & " �A�Ьd�֡I" & "  �N�����O�B�G " & vCM
                ErrFile.WriteLine strErrorMsg
                lngErrCount = lngErrCount + 1
                blnSO3314 = True
            Else
                If lngTotalShouldAmt <> Val(vRealAmt) Then
                    'strErrorMsg = "��ڽs���G " & strOld_billno & " �Ȥ�s���G" & rsChk("CustId") & " ���B���X�A�ɮ��������B$ " & rsChk("TotalShouldAmt").Value & " �A�Ϥ��w�����B$ " & Val(vRealAmt) & "  �N�����O�B�G " & vCM & "  ���O�渹�G" & strBillNo_Old
                    strErrorMsg = "��ڽs���G " & strOld_billno & " �Ȥ�s���G" & strCustId & " ���B���X�A�ɮ��������B$ " & lngTotalShouldAmt & " �A�Ϥ��w�����B$ " & Val(vRealAmt) & "  �N�����O�B�G " & vCM & "  ���O�渹�G" & strBillNo_Old
                    ErrFile.WriteLine strErrorMsg
                    lngErrCount = lngErrCount + 1
                    blnSO3314 = True
                Else
                    strSQL = "Select SO002.CustStatusCode,SO002.WipCode3,So002.ServiceType,SO033.BudgetPeriod," & _
                             "SO033.CitemCode,SO033.CompCode,SO033.FaciSeqNo,SO033.RealPeriod,SO033.CustId " & _
                             " From " & strOwnerName & "So002," & strOwnerName & "SO033 " & _
                             " Where SO002.CustId = SO033.CustId And SO002.CompCode = SO033.CompCode " & _
                             " And SO002.ServiceType=SO033.ServiceType "
                    
                    If intpara24 = 0 Then
                        If blnBillNo Then
                           strSQL = strSQL & " And BillNo='" & strOld_billno & "'"
                        Else
                        '���D��3008 �NLike %����,�H�W�[�Ĳv,By Kin 2007/1/22
                           'strSQL = strSQL & " And CitiBankATM like '%" & Mid(strOld_billno, 1, 15) & "%'"
                           strSQL = strSQL & " And CitiBankATM = '" & strOld_billno & "'"
                        End If
                        
                        'strSQL = "Select CustStatusCode,WipCode3,ServiceType From " & strOwnerName & "So002 Where CustId = " & rsChk("CustId").Value & " And ServiceType= '" & rsChk("ServiceType") & "'"
                    Else
                        If blnBillNo Then
                            If blnMediaBillNo Then
                                strSQL = strSQL & " And SO033.MediaBillNo = '" & vBillNo & "'"
                            Else
                                '7179
                                strSQL = strSQL & " And SO033.BillNo = '" & vBillNo & "'" & IIf(blnCrossCustCombine, " And SO033.CustId = " & strCustId, "")
                            End If
                        Else
                            '���D��3008 �NLike %����,�H�W�[�Ĳv,By Kin 2007/1/22
                            'strSQL = strSQL & " And SO033.CitiBankATM like '" & Mid(vBillNo, 1, 15) & "%'"
                            '7179
                            strSQL = strSQL & " And SO033.CitiBankATM = '" & vBillNo & "'" & IIf(blnCrossCustCombine, " And SO033.CustId = " & strCustId, "")
                        End If
                    End If
                    If Not GetRS(rsCust, strSQL, gcnGi) Then Exit Function
                    Do While Not rsCust.EOF
                        '�D���`�Ȥ�h���@log�A�J�b
                        If rsCust("CustStatusCode").Value & "" <> 1 Then
                            'strErrorMsg = "�Ȥ�s��: " & rsChk("CustId").Value & " ��ڽs���G " & strOld_billno & " ���D���`�Ȥ� " & " �Ȥ᪬�A: " & rsCust("CustStatusCode").Value & " �N�����O�B: " & vCM & " �A�ȧO: " & rsCust("ServiceType").Value
                            strErrorMsg = "�Ȥ�s��: " & strCustId & " ��ڽs���G " & strOld_billno & " ���D���`�Ȥ� " & " �Ȥ᪬�A: " & rsCust("CustStatusCode").Value & " �N�����O�B: " & vCM & " �A�ȧO: " & rsCust("ServiceType").Value
                            ErrFile.WriteLine strErrorMsg
                            lngStatusCount = lngStatusCount + 1
                        '���`�Ȥ�BCM������
                        ElseIf rsCust("CustStatusCode").Value & "" = 1 And rsCust("WipCode3") = 16 Then
                            'strErrorMsg = "�Ȥ�s��: " & rsChk("CustId").Value & " ��ڽs���G " & strOld_billno & " ��CM�������Ȥ� " & " �N�����O�B: " & vCM & " �A�ȧO: " & rsCust("ServiceType").Value
                            strErrorMsg = "�Ȥ�s��: " & strCustId & " ��ڽs���G " & strOld_billno & " ��CM�������Ȥ� " & " �N�����O�B: " & vCM & " �A�ȧO: " & rsCust("ServiceType").Value
                            ErrFile.WriteLine strErrorMsg
                            lngStatusCount = lngStatusCount + 1
                        End If
'                        If Not GetRS(rsTmp, "Select BudgetPeriod From " & strOwnerName & "SO003 Where Custid = " & rsCust("CustId") & " And CompCode = " & rsCust("CompCode") & " And ServiceType = '" & rsCust("ServiceType") & "' " & _
'                            " And CitemCode = " & rsCust("CitemCode") & " And FaciSeqNo = '" & rsCust("FaciSeqNo") & "'") Then Exit Function
'                        If Not rsTmp.EOF Then
'                            If rsCust("RealPeriod") > rsTmp.Fields("BudgetPeriod") And rsTmp.Fields("BudgetPeriod") > 0 Then
'                                strErrorMsg = "�Ȥ�s��: " & rsChk("CustId").Value & " ��ڽs���G " & strOld_billno & "�J�b���Ƥj������Ѿl����,�L�k�J�b"
'                                ErrFile.WriteLine strErrorMsg
'                                lngErrCount = lngErrCount + 1
'                                blnSO3314 = True
'                                GoTo XX
'                            End If
'                        End If
                        
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
'#5386 �P�_���X���Ƚs���ۦP By Kin 2009/12/04
Private Function ChkDiffCust(ByRef rs As ADODB.Recordset) As String
  On Error GoTo ChkErr
    Dim strRetMsg As String
    Dim strDiffCust As String
    Dim strRealDate As String
    Dim strRetCustId As String
    Dim strRetRealDate As String
    If blnCrossCustCombine Then
        ChkDiffCust = ""
        Exit Function
    End If
    rs.MoveFirst
    Do While Not rs.EOF
        If strDiffCust <> rs("CustId") & "" Then
            strDiffCust = rs("CustId") & ""
            strRetCustId = IIf(strRetCustId = "", strDiffCust, strRetCustId & ";" & strDiffCust)
        End If
        rs.MoveNext
    Loop
    rs.MoveFirst
    If InStr(1, strRetCustId, ";") > 0 Then
        strRetMsg = strRetCustId
    End If
    ChkDiffCust = strRetMsg
    Exit Function
ChkErr:
  ErrSub "UpdateFun", "ChkDiffCust"
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
            '���D��3146 ��Update��Like�A�{�b�אּ=  by Kin 2007/3/28
            strSQL = "Select RowId,So033.* From " & strOwnerName & "So033 Where UCCODE IS NOT NULL  And CitiBankATM = '" & vBillNo & "' AND CancelFlag <> 1" & strSQL2
            'strSQL = "Select RowId,So033.* From " & strOwnerName & "So033 Where UCCODE IS NOT NULL  And CitiBankATM like '" & Mid(vBillNo, 1, 15) & "%' AND CancelFlag <> 1" & strSQL2
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
    ErrSub "UpdateFunMid", "ChkOverCloseDate"
End Function

