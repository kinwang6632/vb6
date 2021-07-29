Attribute VB_Name = "UpdateFunMid"
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
Public blnFailUpd As Boolean
Public strUCCode As String
Public strUCName As String
Public blnCrossCustCombine  As Boolean
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
    Call ErrSub("UpdateFunMid", "Function ChkDirExist")
End Function

Public Function Decrypt(EncryptionString As String) As String
  On Error GoTo ChkErr
    Decrypt = CreateObject("DevPowerEncrypt.EnCrypt").Decrypt(EncryptionString)
  Exit Function
ChkErr:
    Call ErrSub("UpdateFunMid", "Function Decrypt")
    
End Function

'#3527 �W�[�Ѽ�intRefNo(CD018.REFNO)�A�p�G��2�A�h�ϥ�POS3�s���y�{�A�ثe�u��mediaPost3�ϥ� By Kin 2007/10/05
'#3417 �W�[�Ѽ�vBankId�A�p�G���O�H�����ȡA�hSO033���Ƶ���ݶ�J�Ȧ�N�X
Public Function ChkData(ByVal vBillNo As String, ByVal vRealAmt As String, _
                        ByVal vClctEn As String, ByVal vClctName As String, _
                        ByVal vRealDate As String, ByVal vUpdEn As String, _
                        ByVal vCMCode As String, ByVal vCMName As String, _
                        ByVal vPTCode As String, ByVal vPTName As String, _
                        ByVal vCM As String, _
                        Optional blnBillNO As Boolean = False, _
                        Optional strCitemQry As String, _
                        Optional strMedia As String, _
                        Optional intRefNo As Integer = 0, _
                        Optional ByVal vBankId As String = Empty) As Boolean
  On Error GoTo ChkErr
  Dim strSQL As String
  Dim rsChk As New ADODB.Recordset
  Dim rsCust As New ADODB.Recordset
  Dim strSQL2 As String
  Dim strBillNo_Old As String
  Dim rsTmp As New ADODB.Recordset
  Dim strReturn As String
  Dim strWhere As String, lngBudgetPeriod As Long
  Dim strFailUpdSQL As String
  Dim strFail As String
  Dim strCD013RefNo As String
  Dim lngTotalRealAmt As Long
  Dim lngTotalShouldAmt As Long
  Dim strCustId As String
    '#5564 CD013.PAYOK=1��CD013.REFNO=3,7�����d�x�w�� By Kin 2010/05/19
    strCD013RefNo = " MAX(DECODE(NVL(CD013.PAYOK,0),1,3," & _
                "DECODE(NVL(CD013.REFNO,0),3,3,DECODE(NVL(CD013.REFNO,0),7,3,0)))) REFNO "
    strFail = Empty
    '���O����
    If strCitemQry <> "" Then strSQL2 = " And SO033.CitemCode In (" & strCitemQry & ")"
    '7179
    blnCrossCustCombine = Val(GetRsValue("Select Nvl(CrossCustCombine,0) From " & strOwnerName & "SO041", gcnGi) & "")
    
    '***********************************************************************************************************************************************************************************************************
    '#3527 �P�_�O�_�ϥΰѦҸ���2�A�p�G��2�h�ϥηs��POS3�y�{ By Kin 2007/10/05
    If intRefNo <> 2 Then
        strWhere = " SO033.UCCODE IS NOT NULL AND " & IIf(blnBillNO, "SO033.BillNo", "SO033.MediaBillNo") & " = '" & vBillNo & "' And SO033.CancelFlag <> 1 " & strSQL2
        
    Else
        strWhere = " SO033.UCCODE IS NOT NULL AND SO033.CitibankATM='" & vBillNo & "' And SO033.CancelFlag <> 1 " & strSQL2
    End If
    '***********************************************************************************************************************************************************************************************************
    
    '***************************************************************************************************************
    '#3417 ���Ѯɭn��s���y�k By Kin 2007/12/10
    '#4388 �W�[��s���ʤH���P���ʮɶ� By Kin 2009/05/04
    If blnFailUpd Then
        strFailUpdSQL = "Update " & strOwnerName & "SO033 Set UCCode=" & strUCCode & _
                        ",UPDTime='" & objStorePara.uUPDTime & "',UPDEN='" & objStorePara.UpdEn & "'" & _
                        ",UCName='" & strUCName & "',Note=Note"
                                                                                  
    End If
    '***************************************************************************************************************
    '#4010 �p�GUCcode���ѦҸ���3�n�������` By Kin 2008/07/29
    '#5564 CD013.PAYOK=1��CD013.REFNO=3,7�����d�x�w�� By Kin 2010/05/19
    If intRefNo <> 2 Then
        If strMedia = 0 Then
            If blnBillNO Then
                strSQL = "Select Nvl(sum(ShouldAmt),0) as TotalShouldAmt, Nvl(sum(RealAmt),0) as TotalRealAmt," & _
                        "SO033.RealDate, SO033.CustId, SO033.ServiceType," & strCD013RefNo & " From " & strOwnerName & "SO033," & strOwnerName & "CD013 " & _
                        " Where SO033.UCCODE IS NOT NULL AND SO033.BillNo = '" & vBillNo & "' And SO033.CancelFlag <> 1 And SO033.UCCode=CD013.CodeNo " & _
                        strSQL2 & " Group By SO033.CustId,SO033.RealDate,SO033.ServiceType "
            Else
                strSQL = "Select Nvl(sum(SO033.ShouldAmt),0) as TotalShouldAmt, Nvl(sum(SO033.RealAmt),0) as TotalRealAmt," & _
                         "SO033.RealDate,SO033.CustId,SO033.ServiceType," & strCD013RefNo & " From " & strOwnerName & "SO033," & strOwnerName & "CD013 " & _
                         " Where SO033.UCCODE IS NOT NULL AND SO033.MediaBillNo = '" & vBillNo & "' And SO033.CancelFlag <> 1 And SO033.UCCode=CD013.CodeNo " & _
                         strSQL2 & " Group By SO033.CustId,SO033.RealDate,SO033.ServiceType "
                
            End If
        Else
        
            If blnBillNO Then
                strSQL = "Select Nvl(sum(SO033.ShouldAmt),0) as TotalShouldAmt, Nvl(sum(SO033.RealAmt),0) as TotalRealAmt," & _
                         "SO033.RealDate, SO033.CustId," & strCD013RefNo & " From " & strOwnerName & "SO033," & strOwnerName & "CD013" & _
                         " Where SO033.UCCODE IS NOT NULL AND SO033.BillNo = '" & vBillNo & "' And SO033.CancelFlag <> 1 And SO033.UCCode=CD013.CodeNo " & _
                         strSQL2 & " Group By SO033.CustId,SO033.RealDate"
            Else
                strSQL = "Select Nvl(sum(SO033.ShouldAmt),0) as TotalShouldAmt, Nvl(sum(SO033.RealAmt),0) as TotalRealAmt," & _
                         "SO033.RealDate,SO033.CustId,SO033.MediaBillNo," & strCD013RefNo & " From " & strOwnerName & "SO033," & strOwnerName & "CD013" & _
                         " Where SO033.UCCODE IS NOT NULL AND SO033.MediaBillNo = '" & vBillNo & "' And SO033.CancelFlag <> 1 And SO033.UCcode=CD013.CodeNo " & _
                         strSQL2 & " Group By SO033.CustId,SO033.RealDate,SO033.MediaBillNo"
                 
            End If
        End If
    Else
        If vBillNo <> "" Then
            strSQL = "Select Nvl(sum(SO033.ShouldAmt),0) as TotalShouldAmt, Nvl(sum(SO033.RealAmt),0) as TotalRealAmt," & _
                     "SO033.RealDate,SO033.CustId," & strCD013RefNo & " From " & strOwnerName & "SO033," & strOwnerName & "CD013" & _
                     " Where SO033.UCCODE IS NOT NULL AND SO033.CitibankATM = '" & vBillNo & "' And SO033.CancelFlag <> 1 And SO033.UCCode=CD013.CodeNo " & strSQL2 & _
                     " Group By SO033.CustId,SO033.RealDate"
        Else
            ErrFile.Write ("�����b�������T�I�I")
            lngErrCount = lngErrCount + 1
            ChkData = False
            Exit Function
        End If
    End If
    '7179
    If blnCrossCustCombine Then
        strSQL = "Select A.*,SO001.AMduid, (Select MainCustId From " & strOwnerName & "SO202 Where MduId = SO001.AmduId) MainCustId " & _
                             " From (" & strSQL & " ) A," & strOwnerName & "SO001 " & _
                             " Where A.CustId = SO001.CustId "
    End If
    '*********************************************************************************************************************************************************************************************************************************************************************************************************************
    If rsChk.State = adStateOpen Then rsChk.Close
  
    If Not GetRS(rsChk, strSQL, gcnGi) Then
       ChkData = False
       Exit Function
    End If
    
    If rsChk.EOF Or rsChk.BOF Then
       ErrFile.WriteLine ("��ڽs���G " & vBillNo & " �d�L���渹�I" & "  �N�����O�B�G " & vCM)
'       strFail = strFailUpdSQL & "||'��ڽs���G " & vBillNo & _
'                                 "  �d�L���渹�I" & _
'                                 "  �N�����O�B�G " & vCM & "'" & _
'                                 " Where " & strWhere
'       gcnGi.Execute strFail
       lngErrCount = lngErrCount + 1
       ChkData = True
       Exit Function
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
                ErrFile.WriteLine ("��ڽs���G " & vBillNo & " �ۦP��ڽs�����Ȥ�s�����P�I �Ƚs�G" & strRetMsg)
                If blnFailUpd Then
                    strFail = strFailUpdSQL & "||'��ڽs���G" & vBillNo & _
                            " �ۦP��ڽs�����Ȥ�s�����P�I �Ƚs�G" & strRetMsg & "'" & _
                            " Where " & strWhere
                    gcnGi.Execute strFail
                
                End If
                lngErrCount = lngErrCount + 1
                ChkData = True
                Exit Function
            End If
        End If
        '***********************************************************************
        
        '**************************************************************************************************
        '#4010 UCCode���ѦҸ���3�n�����d�O�w�� By Kin 2007/07/29
        If Val(rsChk("RefNo")) = 3 Then
            ErrFile.WriteLine ("��ڽs���G " & vBillNo & " �d�O�w���Ьd�֡I" & "  �N�����O�B�G " & vCM)
            If blnFailUpd Then
                strFail = strFailUpdSQL & "||'��ڽs���G " & vBillNo & _
                                " �d�O�w���Ьd�֡I " & _
                                "  �N�����O�B�G " & vCM & "'" & _
                                " Where " & strWhere
                
                gcnGi.Execute strFail
            End If
            lngErrCount = lngErrCount + 1
            ChkData = True
            Exit Function
        End If
        '***************************************************************************************************
        '*********************************************************************************************************************************************
        '#3527 �P�_�O�_�ϥΰѦҸ���2�A�p�G��2�h�ϥηs��POS3�y�{ By Kin 2007/10/05
        If intRefNo <> 2 Then
          '��X�³渹 2003/08/27
            '7179
            If blnBillNO = False Then
               strBillNo_Old = GetRsValue("Select BillNo From " & strOwnerName & "SO033 Where MediaBillNO='" & vBillNo & "'" & _
                                           IIf(blnCrossCustCombine, " And Custid = " & strCustId, ""), gcnGi) & ""
            Else
               strBillNo_Old = vBillNo
            End If
        Else
            '7179
            strBillNo_Old = GetRsValue("Select BillNo From " & strOwnerName & "SO033 Where CitibankATM='" & vBillNo & "'" & _
                              IIf(blnCrossCustCombine, " And Custid = " & strCustId, ""), gcnGi) & ""
        End If
        '**********************************************************************************************************************************************
        
       If rsChk("RealDate").Value & "" <> "" And lngTotalRealAmt > 0 Then
            'ErrFile.WriteLine ("��ڽs���G " & vBillNo & " �ɮפ��w�ꦬ���B$ " & rsChk("TotalRealAmt").Value & " �A�Ьd�֡I" & "  �N�����O�B�G " & vCM)
            ErrFile.WriteLine ("��ڽs���G " & vBillNo & " �ɮפ��w�ꦬ���B$ " & lngTotalRealAmt & " �A�Ьd�֡I" & "  �N�����O�B�G " & vCM)
            
            '********************************************************************************************************************************************************************************
            '#3417 ���~�ɧ�sUCCode�MUCName By Kin 2007/12/10
            If blnFailUpd Then
                
'                strFail = strFailUpdSQL & "||'��ڽs���G " & vBillNo & _
'                                " �ɮפ��w�ꦬ���B$ " & rsChk("TotalRealAmt").Value & _
'                                " �A�Ьd�֡I" & "  �N�����O�B�G " & vCM & "'" & _
'                                " Where " & strWhere
                 strFail = strFailUpdSQL & "||'��ڽs���G " & vBillNo & _
                                " �ɮפ��w�ꦬ���B$ " & lngTotalRealAmt & _
                                " �A�Ьd�֡I" & "  �N�����O�B�G " & vCM & "'" & _
                                " Where " & strWhere
                gcnGi.Execute strFail
            End If
            '*******************************************************************************************************************************************************************************
            
            lngErrCount = lngErrCount + 1
            ChkData = True
            Exit Function
       Else
          '#3122 �P�_���ڤ���O�_�j��鵲���
          If Not ChkOverCloseDate(Format(vRealDate, "####/##/##"), objStorePara.CompCode, objStorePara.ServiceType, strReturn) Then
                ErrFile.WriteLine ("��ڽs���G" & vBillNo & "  ��ڦ��ڤ�G" & Format(str(Val(vRealDate) - 19110000), "##/##/##") & "  ���~��]�G" & strReturn)
                
                '***************************************************************************************************************************************************************************************
                '#3417 ���~�ɧ�sUCCode�MUCName By Kin 2007/12/10
                If blnFailUpd Then
                    
                    strFail = strFailUpdSQL & "||'��ڽs���G" & vBillNo & _
                                    "  ��ڦ��ڤ�G" & Format(str(Val(vRealDate) - 19110000), "##/##/##") & _
                                    "  ���~��]�G" & strReturn & "'" & _
                                    " Where " & strWhere
                    gcnGi.Execute strFail
                End If
                '****************************************************************************************************************************************************************************************
                
                lngErrCount = lngErrCount + 1
                ChkData = True
                Exit Function
          End If
          If lngTotalShouldAmt <> Val(vRealAmt) Then
                'ErrFile.WriteLine ("��ڽs���G " & vBillNo & " ���B���X�A�ɮ��������B$ " & rsChk("TotalShouldAmt").Value & " �A�Ϥ��w�����B$ " & Val(vRealAmt) & "  �N�����O�B�G " & vCM & "  ���O�渹�G" & strBillNo_Old)
                ErrFile.WriteLine ("��ڽs���G " & vBillNo & " ���B���X�A�ɮ��������B$ " & lngTotalShouldAmt & " �A�Ϥ��w�����B$ " & Val(vRealAmt) & "  �N�����O�B�G " & vCM & "  ���O�渹�G" & strBillNo_Old)
                '*********************************************************************************************************************************************************************************************************************************************
                '#3417 ���~�ɧ�sUCCode�MUCName By Kin 2007/12/10
                If blnFailUpd Then
                    
'                    strFail = strFailUpdSQL & "||'��ڽs���G " & vBillNo & _
'                                    " ���B���X�A�ɮ��������B$ " & rsChk("TotalShouldAmt").Value & _
'                                    " �A�Ϥ��w�����B$ " & Val(vRealAmt) & "  �N�����O�B�G " & vCM & _
'                                    "  ���O�渹�G" & strBillNo_Old & "'" & _
'                                    " Where " & strWhere
                    strFail = strFailUpdSQL & "||'��ڽs���G " & vBillNo & _
                                    " ���B���X�A�ɮ��������B$ " & lngTotalShouldAmt & _
                                    " �A�Ϥ��w�����B$ " & Val(vRealAmt) & "  �N�����O�B�G " & vCM & _
                                    "  ���O�渹�G" & strBillNo_Old & "'" & _
                                    " Where " & strWhere
                    gcnGi.Execute strFail
                End If
                '*********************************************************************************************************************************************************************************************************************************************
                
                lngErrCount = lngErrCount + 1
                ChkData = True
                Exit Function
          Else
                If rsCust.State = adStateOpen Then rsCust.Close
                
                '***************************************************************************************************************************************************************************************************************************************************************************
                '#3527 �P�_�O�_�ϥΰѦҸ���2�A�p�G��2�h�ϥηs��POS3�y�{ By Kin 2007/10/05
                If intRefNo <> 2 Then
                    If strMedia = 0 Then
                        strSQL = "Select CustStatusCode,WipCode3,ServiceType From " & strOwnerName & "So002 Where CustId = " & rsChk("CustId").Value & " And ServiceType= '" & rsChk("ServiceType") & "'"
                    Else
                        If blnBillNO Then
                            strSQL = "Select CustStatusCode,WipCode3,ServiceType From " & strOwnerName & "So002 Where CustId = " & rsChk("CustId").Value & " And ServiceType= '" & rsChk("ServiceType") & "'"
                        Else
                            strSQL = "Select SO002.CustStatusCode,SO002.WipCode3,So002.ServiceType From " & strOwnerName & "So002," & strOwnerName & "SO033 Where SO002.ServiceType=SO033.ServiceType And SO002.CustId = SO033.CustId And MediaBillNo = '" & rsChk("MediaBillNo").Value & "'"
                        End If
                    End If
                Else
                    strSQL = "Select SO002.CustStatusCode,SO002.WipCode3,So002.ServiceType From " & strOwnerName & "So002," & strOwnerName & "SO033 Where SO002.ServiceType=SO033.ServiceType And SO002.CustId = SO033.CustId And CitibankATM = '" & vBillNo & "'"
                End If
                '***************************************************************************************************************************************************************************************************************************************************************************
                '7179
                If blnCrossCustCombine Then
                    'strSQL = strSQL & " And SO033.Custid = " & strCustId
                End If
                Call GetRS(rsCust, strSQL, gcnGi)
                While Not rsCust.EOF
                    '�D���`�Ȥ�h���@log�A�J�b
                    If rsCust("CustStatusCode").Value & "" <> 1 Then
                        ErrFile.WriteLine ("�Ȥ�s��: " & rsChk("CustId").Value & " ��ڽs���G " & vBillNo & " ���D���`�Ȥ� " & " �Ȥ᪬�A: " & rsCust("CustStatusCode").Value & " �N�����O�B: " & vCM & " �A�ȧO: " & rsCust("ServiceType").Value)
                        lngStatusCount = lngStatusCount + 1
                    '���`�Ȥ�BCM������
                    ElseIf rsCust("CustStatusCode").Value & "" = 1 And rsCust("WipCode3") = 16 Then
                        ErrFile.WriteLine ("�Ȥ�s��: " & rsChk("CustId").Value & " ��ڽs���G " & vBillNo & " ��CM�������Ȥ� " & " �N�����O�B: " & vCM)
                        lngStatusCount = lngStatusCount + 1
                    End If
                    rsCust.MoveNext
                Wend
                '7179
                If Not GetRS(rsTmp, "Select RealPeriod,CustId,CompCode,ServiceType,CitemCode,FaciSeqNo From " & strOwnerName & "So033 " & _
                                " Where " & strWhere) Then Exit Function
                '94/09/30 1805 Jacky
                    Do While Not rsTmp.EOF
                        lngBudgetPeriod = Val(GetRsValue("Select Nvl(BudgetPeriod,0) From " & strOwnerName & "SO003 Where Custid = " & rsTmp("CustId") & " And CompCode = " & rsTmp("CompCode") & " And ServiceType = '" & rsTmp("ServiceType") & "' " & _
                                        " And CitemCode = " & rsTmp("CitemCode") & " And FaciSeqNo = '" & rsTmp("FaciSeqNo") & "'") & "")
                        If rsTmp("RealPeriod") > lngBudgetPeriod And lngBudgetPeriod > 0 Then
                            Call ErrFile.WriteLine("�J�b���Ƥj������Ѿl����,�L�k�J�b!!")
                            
                            '***************************************************************************************
                            '#3417 ���~�ɧ�sUCCode�MUCName By Kin 2007/12/10
                            If blnFailUpd Then
                                
                                strFail = strFailUpdSQL & "||'�J�b���Ƥj������Ѿl����,�L�k�J�b!!'" & _
                                                " Where CustId=" & rsTmp("CustId") & _
                                                " And CompCode=" & rsTmp("CompCode") & _
                                                " And ServiceType = '" & rsTmp("ServiceType") & "'" & _
                                                " And CitemCode = " & rsTmp("CitemCode") & "" & _
                                                " And FaciSeqNo = '" & rsTmp("FaciSeqNo") & "" & "'" & _
                                                " And " & strWhere
                                gcnGi.Execute strFail
                            End If
                            '****************************************************************************************
                            
                            lngErrCount = lngErrCount + 1
                            ChkData = True
                            Exit Function
                        End If
                        rsTmp.MoveNext
                    Loop
                    
            '#3527 �h�Ǥ@��intRefNo�ܼ� By Kin 2007/10/05
            '�����sSO033
                If UpdAccount(vBillNo, vRealAmt, vClctEn, vClctName, vRealDate, vUpdEn, vCMCode, vCMName, vPTCode, vPTName, blnBillNO, strCitemQry, intRefNo, vBankId) = False Then
                    ChkData = False
                    Exit Function
                End If
          End If
       End If
    End If
    If rsChk.State = adStateOpen Then rsChk.Close
    Set rsChk = Nothing
    If rsCust.State = adStateOpen Then rsCust.Close
    Set rsCust = Nothing
    
    ChkData = True
  Exit Function
ChkErr:
  ErrSub "UpdateFunMid", "ChkData"
End Function
'#5386 �P�_���X���Ƚs���ۦP By Kin 2009/12/04
Private Function ChkDiffCust(ByRef rs As ADODB.Recordset) As String
  On Error GoTo ChkErr
    Dim strRetMsg As String
    Dim strDiffCust As String
    Dim strRealDate As String
    Dim strRetCustId As String
    Dim strRetRealDate As String
    '7179
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
  ErrSub "UpdateFunMid", "ChkDiffCust"
End Function

'#3527 �W�[�Ѽ�intRefNo(CD018.REFNO)�A�p�G��2�A�h�ϥ�POS3�s���y�{�A�ثe�u��mediaPost3�ϥ� By Kin 2007/10/05
'#3417 �W�[�Ѽ� vBankId�A�p�G����ܦ��O�H���hSO033���Ƶ���ݶ�J�Ȧ�N�X
Public Function UpdAccount(ByVal vBillNo As String, ByVal vRealAmt As String, _
                           ByVal vClctEn As String, ByVal vClctName As String, _
                           ByVal vRealDate As String, ByVal vUpdEn As String, _
                           ByVal vCMCode As String, ByVal vCMName As String, _
                           ByVal vPTCode As String, ByVal vPTName As String, _
                           Optional blnBillNO As Boolean = False, Optional strCitemQry As String, _
                           Optional intRefNo As Integer = 0, _
                           Optional vBankId As String = Empty) As Boolean
    On Error GoTo ChkErr
    Dim rsAccount As New ADODB.Recordset
    Dim strSQL As String, strSeqQry As String, lngAffected As Long
    Dim lngPeriod As Long, lngAmount As Long
    Dim strSo003 As String, strUpdSql As String
    Dim strSQL2 As String, rsTmp As New ADODB.Recordset
        '���O����
        If strCitemQry <> "" Then strSQL2 = " And CitemCode In (" & strCitemQry & ")"
        If intRefNo <> 2 Then
            If blnBillNO Then
                strSQL = "Select RowId, So033.* From " & strOwnerName & "So033 Where UCCODE IS NOT NULL  And BillNo ='" & vBillNo & "' AND CancelFlag <> 1" & strSQL2
            Else
                strSQL = "Select RowId, So033.* From " & strOwnerName & "So033 Where UCCODE IS NOT NULL  And MediaBillNo ='" & vBillNo & "' AND CancelFlag <> 1" & strSQL2
            End If
        Else
            strSQL = "Select RowId, So033.* From " & strOwnerName & "So033 Where UCCODE IS NOT NULL  And CitibankATM ='" & vBillNo & "' AND CancelFlag <> 1" & strSQL2
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
                '#7033 �p�G����J�I�ں����A�h�H��J���D�A�_�h���^��null By Kin 2015/04/28
                If Len(vPTCode) > 0 Then
                    .Fields("PTCode") = IIf(vPTCode = "", Null, vPTCode)
                    .Fields("PTName") = NoZero(vPTName)
                End If
'                .Fields("PTCode") = IIf(vPTCode = "", Null, vPTCode)
'                .Fields("PTName") = NoZero(vPTName)
                .Fields("ClctEn") = IIf(Trim(vClctEn) = "", rsAccount("ClctEn"), vClctEn)
                .Fields("ClctName") = IIf(Trim(vClctName) = "", rsAccount("ClctName"), vClctName)
                .Fields("FirstTime") = GetDTString(strNowTime)
                
                '*************************************************************************
                '#3417 �p�G���怜�O�H���hSO033.Note�n��J�Ȧ�N�X By Kin 2008/01/22
                If vClctEn <> Empty Then
                    If vBankId <> Empty Then
                        .Fields("Note") = .Fields("Note") & " " & "�Ȧ�N�X:" & vBankId
                    End If
                End If
                '*************************************************************************
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
'
'            If Not GetRS(rsTmp, "Select * From " & strOwnerName & "SO033 Where RowId = '" & rsAccount("RowId") & "'") Then Exit Function
'            If AlterSO003(rsTmp) = False Then
'               UpdAccount = False
'               Exit Function
'            End If
'
lLoop:
            rsAccount.MoveNext
        Wend
            
            
'            If intPara8 >= 1 Then
'                lngPeriod = NoZero(rsAccount("RealPeriod"), True)
'                If intPara9 = 1 Then
'                    lngAmount = NoZero(rsAccount("OldAmt"), True)
'                ElseIf intPara9 = 2 Then
'                    lngAmount = NoZero(rsAccount("ShouldAmt"), True)
'                Else
'                    lngAmount = NoZero(rsAccount("ShouldAmt"), True)
'                End If
'                strUpdSql = ",Amount = " & lngAmount & ",Period = " & lngPeriod
'            End If
'
'
'            If rsAccount("SeqNo") Then strSeqQry = " And SeqNo = " & rsAccount("SeqNo")
'
'            If Val(rsAccount("BudgetPeriod") & "") > 0 Then
'                If Val(rsAccount("BudgetPeriod") & "") = 1 Then
'                    If Not ExecuteCommand("Delete From " & strOwnerName & "SO003 Where CustId = '" & rsAccount("CustId") & "' And CitemCode = '" & rsAccount("CitemCode") & "' And StopDate < " & _
'                        GetNullString(rsAccount("RealStopDate"), giDateV, giOracle) & strSeqQry) Then Exit Function
'                    GoTo lLoop
'                End If
'                strSo003 = ",BudgetPeriod = " & (rsAccount("BudgetPeriod") - 1)
'            End If
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
        UpdAccount = True
    Exit Function
ChkErr:
    UpdAccount = False
End Function

Public Function InitData(Optional strCompCode As String = "", Optional strServiceType As String = "") As Boolean
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
  ErrSub "UpdateFunMid", "InitData"
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
  ErrSub "UpdateFunMid", "MsgResult"
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
    objLst.SendConn gcnGi
    objLst.Filter = strWhere
  Exit Sub
ChkErr:
    Call ErrSub("UpdateFunMid", "SetLst")
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

Public Function GetBillNo_New(ByVal strNo As String) As String
    On Error Resume Next
    Dim strType As String
    Dim strYM As String
    Dim strMM As String
    
'        strNo = Mid(strReadLine, 65, 11)
        
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
End Function

Public Function GetBillNo_Old(ByVal strNo As String, ByVal strServiceType) As String
  On Error Resume Next
    
    GetBillNo_Old = Trim(CStr(Val(Left(strNo, 4)) + 191100)) & _
            Mid(strNo, 5, 1) & strServiceType & Format(Right(strNo, 6), "0000000")
End Function
'#3122 �ˮ֬O�_���W�L�鵲��� By Kin Add 2007/03/28
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



