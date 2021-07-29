Attribute VB_Name = "UpdateFun3313"
Option Explicit
Private Const FormName = "UpdateFun"
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Const gMsgIsDataOK = "�������n���,������ !! "
Public lngCount As Long           '�O�����\����
Public lngErrCount As Long        '�O�����~����
Public lngStatusCount As Long     '�O���D���`�ᵧ��
Public lngAmtTotal As Long        '�O�����\���B
Public lngTime As Long             '�O���Ҫ�ɶ�
Public strNowTime As String        '�O���ثe�ɶ�
Public fso As New FileSystemObject
Public File As TextStream
Public ErrFile As TextStream     '���~���log��
Private intPara3 As Integer
Private intPara5 As Integer
Private intPara8 As Integer
Public gcnGi As New ADODB.Connection
Public strUserName As String
Private strFilePath As String
Public objStorePara As clsStoreParameter
Public blnFailUpd As Boolean
Public strUCCode As String
Public strUCName As String
Public strFailUpdSQL As String
Public Enum giArrow
    giLeft = 0
    giRight = 1
End Enum
Public Enum ChangeDB
    giAccessDb = 0 'Access Table
    giOracle = 1    'Oracle Table
End Enum
Enum giVType
    giLongV = 0
    giStringV = 1
    giDateV = 2
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
Public blnMultiAccount As Boolean     ''  �O�� SO043 �� SO043.para24  �A�H�P�_�O�_�Φh�b��]�w
Public strOwnerName As String

'#3697 �W�[����H�U�s�榡(lngChinatrustNew) By Kin 2007/12/31
Public Enum mCreditCardType
    lngUnionChinese = 0
    lngChinatrust = 1
    lngHSBC = 2
    lngChinatrustNew = 3
    lngNewWeb = 4
    lngUnionBath = 5
    lngFubon = 6
End Enum

Public GetOwner As String

Public blnCrossCustCombine  As Boolean
'�����W�h�Ǫ��ɮ׸��|
Public Function GetFilePath(ByVal vFilePath As String)
  On Error GoTo ChkErr
     strFilePath = vFilePath
  Exit Function
ChkErr:
  ErrSub "DealFun", "GetFilePath"
End Function

Public Sub ErrSub(FormName As String, ProcedureName As String)
    Dim strErr As String
    Dim aryErr As Variant
    Dim strCrLf As String
    Dim RetVal As Variant
    DoEvents
    Screen.MousePointer = vbDefault
    aryErr = Array("�o�Ϳ��~�ɶ� : ", "�o�Ϳ��~�M�� : ", "�o�Ϳ��~��� : ", "�o�Ϳ��~�{�� : ", "���~�N�X : ", "���~��] : ")
    strCrLf = vbCrLf + vbCrLf
    strErr = aryErr(0) + CStr(Now) + strCrLf + _
             aryErr(1) + CStr(Err.Source) + strCrLf + _
             aryErr(2) + FormName + strCrLf + _
             aryErr(3) + ProcedureName + strCrLf + _
             aryErr(4) + CStr(Err.Number) + strCrLf + _
             aryErr(5) + CStr(Err.Description)
    ErrLog strErr
    RetVal = MsgBox(strErr, vbYesNo, "�t�ΰ�����~..(�� <�O> �C�L���~�T��)!!")
    If RetVal = vbYes Then
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
    Dim fnum As Integer
    
    'LogFile = ReadFromINI(GetIniPath1, "GICMISPath", "ErrLogPath")
    LogFile = strFilePath
    If Not ChkDirExist(LogFile) Then MkDir LogFile
    
    If Right(LogFile, 1) <> "\" Then LogFile = LogFile & "\"
    LogFile = LogFile & "SYSERR.TXT"
    
    ErrString = Replace(ErrString, vbCrLf & vbCrLf, vbCrLf)
    ErrString = "�o�Ϳ��~���t�� : �}�իȪA��B�޲z�t��" & vbCrLf & _
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

Public Function GetRsValue(strSQL As String, Optional CnnObj As ADODB.Connection, Optional MsgNoData As String) As Variant
  On Error GoTo ChkErr
    Dim rs As New ADODB.Recordset
    If CnnObj Is Nothing Then Set CnnObj = gcnGi
    If Not GetRS(rs, strSQL, CnnObj) Then Exit Function
    If rs.RecordCount = 0 Then
        If Len(MsgNoData) > 0 Then
            MsgBox MsgNoData, vbExclamation, "����!"
        End If
        GetRsValue = Null
    Else
        GetRsValue = rs(0)
    End If
On Error Resume Next
  Set rs = Nothing
  Exit Function
ChkErr:
    ErrSub "DealFun", "GetRsValue"
End Function

Public Function GetString(ByVal strS As Variant, ByVal intLength As Integer, Optional ByVal lngArrow As giArrow = giLeft, Optional ByVal InsertZero As Boolean = False) As String
  On Error GoTo ChkErr
    Dim strCH As String
    strCH = CStr(strS & "")
    If lngArrow = giLeft Then
        strCH = StrConv(LeftB(StrConv(strCH & Space(intLength), vbFromUnicode), intLength), vbUnicode)
    Else
        strCH = StrConv(RightB(StrConv(Space(intLength) & strCH, vbFromUnicode), intLength), vbUnicode)
    End If
    If InsertZero Then
    If Trim(strCH) = "" Then strCH = "0"
        If lngArrow = giLeft Then
            strCH = RTrim(strCH) & String(intLength - Len(RTrim(strCH)), "0")
        Else
            strCH = String(intLength - Len(LTrim(strCH)), "0") & LTrim(strCH)
        End If
    End If
    strCH = Replace(strCH, Chr(0), "")
    GetString = strCH
  Exit Function
ChkErr:
    Call ErrSub("Sys_Lib", "GetString")
End Function

Public Function ChkData(ByVal vBillNo As String, ByVal vRealAmt As String, _
                        ByVal vRealDate As String, ByVal vUpdEn As String, _
                        ByVal vCMCode As String, ByVal vCMName As String, _
                        ByVal vPTCode As String, ByVal vPTName As String, _
                        Optional pEmpNO As String = "", Optional pEmpName As String = "", _
                        Optional pAuthorizeNo As String = " ", Optional p_frmCreditCardType As mCreditCardType = lngUnionChinese, _
                        Optional pCitemCodeChkData As String, Optional pstrServiceType As String) As Boolean
                        
  On Error GoTo ChkErr
    Dim strSQL As String
    Dim rsChk As New ADODB.Recordset
    Dim rsCust As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim strBillField As String
    Dim strWhere As String
    Dim rsCheckBillNoCancelFlag As New ADODB.Recordset
    Dim lngBudgetPeriod As Long
  
  '''�H�U�� BillNo ��� �令 MediaBillNo
  
    If blnMultiAccount = True Then
         strBillField = "MediaBillno"
    Else
         strBillField = "Billno"
    End If
    strSQL = "Select " & _
                "Nvl(sum(ShouldAmt),0) as TotalShouldAmt, Nvl(sum(RealAmt),0) as TotalRealAmt, RealDate, CustId " & _
                "From " & _
                frmUnion.TableOwnerName & "SO033 " & _
                "Where UCCODE IS NOT NULL AND " & _
                strBillField & " = '" & vBillNo & "' And CancelFlag <> 1 Group By CustId,RealDate"
    
    If rsChk.State = adStateOpen Then rsChk.Close
    If Not GetRS(rsChk, strSQL, gcnGi) Then
       ChkData = False
       Exit Function
    End If
    If rsChk.EOF Or rsChk.BOF Then
         ''2004/09  �[�JUCCODE �w�����P�_�� �H���ѧ���T���T��
        strSQL = "Select count(CustId) C " & _
                "From " & _
                frmUnion.TableOwnerName & "SO033 " & _
                "Where " & _
                strBillField & " = '" & vBillNo & "' And CancelFlag <> 1 Group By CustId,RealDate"
        ''  2004/09/13 �[�J�h�@�h���P�_�� ���� bof eof
        If GetRS(rsCheckBillNoCancelFlag, strSQL, gcnGi, adUseClient, adOpenDynamic) = True Then
            If rsCheckBillNoCancelFlag.EOF Or rsCheckBillNoCancelFlag.BOF Then
                ErrFile.WriteLine ("��ڽs���G " & vBillNo & " �d�L���渹�I")
            Else
                ErrFile.WriteLine ("��ڽs���G " & vBillNo & " ���渹����ڤw���I")
            End If
        Else
            ErrFile.WriteLine ("��ڽs���G " & vBillNo & " �d�L���渹�I")
        End If
        lngErrCount = lngErrCount + 1
    Else
        If rsChk("RealDate").Value & "" <> "" And rsChk("TotalRealAmt").Value > 0 Then
            ErrFile.WriteLine ("�Ȥ�s��: " & rsChk("CustId").Value & "��ڽs���G " & vBillNo & " �ɮפ��w�]�t�ꦬ���B$ " & rsChk("TotalRealAmt").Value & " �A�Ьd�֡I")
            lngErrCount = lngErrCount + 1
        Else
              ''   2003/10/13 ����H�U�H�Υd
            If rsChk("TotalShouldAmt").Value <> Val(vRealAmt) And p_frmCreditCardType <> 1 Then
                ErrFile.WriteLine ("�Ȥ�s��: " & rsChk("CustId").Value & "��ڽs���G " & vBillNo & " ���B���X�A�ɮ��������B$ " & rsChk("TotalShouldAmt").Value & " �A�Ϥ��w�����B$ " & Val(vRealAmt))
                lngErrCount = lngErrCount + 1
            Else
                If rsCust.State = adStateOpen Then rsCust.Close
                    strSQL = "Select CustStatusCode,WipCode3 ,CustStatusName From " & frmUnion.TableOwnerName & "So002 Where CustId = " & rsChk("CustId").Value & "  AND SERVICETYPE  = '" & pstrServiceType & "'"
                    Call GetRS(rsCust, strSQL, gcnGi)
                    '�D���`�Ȥ�h���@log�A�J�b
                If rsCust("CustStatusCode").Value & "" <> 1 Then
                                   ErrFile.WriteLine ("�Ȥ�s��: " & rsChk("CustId").Value & " ��ڽs���G " & vBillNo & " ���D���`�Ȥ� " & " �Ȥ᪬�A: " & rsCust("CustStatusName").Value)
                    lngStatusCount = lngStatusCount + 1
                '���`�Ȥ�BCM������
                ElseIf rsCust("CustStatusCode").Value & "" = 1 And rsCust("WipCode3") = 16 Then
                                   ErrFile.WriteLine ("�Ȥ�s��: " & rsChk("CustId").Value & " ��ڽs���G " & vBillNo & " ��CM�������Ȥ� ")
                    lngStatusCount = lngStatusCount + 1
                End If
'                If Not GetRS(rsTmp, "Select RealPeriod,CustId,CompCode,ServiceType,CitemCode,FaciSeqNo From " & strOwnerName & "So033 Where UCCODE IS NOT NULL AND " & _
'                        strBillField & " = '" & vBillNo & "' And CancelFlag <> 1 ") Then Exit Function
'                '94/09/30 1805 Jacky
'                Do While Not rsTmp.EOF
'                    lngBudgetPeriod = Val(GetRsValue("Select Nvl(BudgetPeriod,0) From " & strOwnerName & "SO003 Where Custid = " & rsTmp("CustId") & " And CompCode = " & rsTmp("CompCode") & " And ServiceType = '" & rsTmp("ServiceType") & "' " & _
'                        " And CitemCode = " & rsTmp("CitemCode") & " And FaciSeqNo = '" & rsTmp("FaciSeqNo") & "'") & "")
'                    If rsTmp("RealPeriod") > lngBudgetPeriod And lngBudgetPeriod > 0 Then
'                        Call ErrFile.WriteLine("�J�b���Ƥj������Ѿl����,�L�k�J�b!!")
'                        lngErrCount = lngErrCount + 1
'                        ChkData = True
'                        Exit Function
'                    End If
'                    rsTmp.MoveNext
'                Loop
                '�����sSO033
                If UpdAccount(vBillNo, vRealAmt, vRealDate, vUpdEn, vCMCode, vCMName, _
                               vPTCode, vPTName, pEmpNO, pEmpName, pAuthorizeNo, _
                               pCitemCodeChkData) = False Then
                    ChkData = False
                    Exit Function
                End If  '' If UpdAccount
            End If
        End If
    End If
    If rsChk.State = adStateOpen Then rsChk.Close
    Set rsChk = Nothing
    If rsCust.State = adStateOpen Then rsCust.Close
    Set rsCust = Nothing
    Call CloseRecordset(rsTmp)
    ChkData = True
  Exit Function
ChkErr:
  ErrSub "DealFun", "ChkData"
End Function

Public Function UpdAccount(ByVal vBillNo As String, ByVal vRealAmt As String, _
                           ByVal vRealDate As String, ByVal vUpdEn As String, _
                           ByVal vCMCode As String, ByVal vCMName As String, _
                           ByVal vPTCode As String, ByVal vPTName As String, _
                           Optional pEmpNO As String = "", _
                           Optional pEmpName As String = "", _
                           Optional pAuthorizeNo As String, _
                           Optional pCitemCoddSql As String = "", _
                           Optional vBankId As String = Empty, _
                           Optional vNote As String = Empty) As Boolean

 '' vBillNo ��ڽs��
 '' vClctEn ���O�H��
 '' vRealDate �ꦬ���
 '' vCMCode ���O�覡�N�X  ref CD031''
 '' vPTCode �I�ں���  ref CD032''
  On Error GoTo ChkErr

    Dim rsAccount As New ADODB.Recordset
    Dim strSQL As String
    Dim strUpdateEmpSql As String
    Dim strBillField As String
    Dim intBudgetPeriod As Integer
    Dim strBudgetPeriodSQL As String
    ''  2003/08/18  �o���ܼƩw�q�����ثe�վ㪺�o�@��SO033  �A�N��ǵ��Q�J��FUNCTION AlterSO003
    Dim rsAlterSO003 As New ADODB.Recordset
    Dim strWhereCitemCode As String
 
    strWhereCitemCode = IIf(Len(pCitemCoddSql) > 0, " AND CitemCode " & pCitemCoddSql & "  ", " ")
    If blnMultiAccount = True Then
        strBillField = "MediaBillno"
    Else
        strBillField = "Billno"
    End If
    '' �o�@�檺WHERE  �᭱�� Billno   �令MediaBillno �H�A�X�h�b���\��
   ' strsql = "Select RowId, CustId, ShouldAmt, CitemCode, RealStartDate, RealStopDate, RealPeriod, ClctEn, ClctName, CMCode, CMName,BudgetPeriod,AccountNO, REALAMT From  " & frmUnion.TableOwnerName & "So033    Where UCCODE IS NOT NULL  And " & strBillField & " ='" & vBillNo & "' AND CancelFlag <> 1 " & strWhereCitemCode
    strSQL = "SELECT * From " & frmUnion.TableOwnerName & _
             "So033 Where UCCODE IS NOT NULL  And " & strBillField & " ='" & vBillNo & "'" & _
             " AND CancelFlag <> 1 " & strWhereCitemCode
    
    If Not GetRS(rsAccount, strSQL, gcnGi, adUseClient, adOpenDynamic, adLockOptimistic) Then
        UpdAccount = False
        Exit Function
    End If
    If rsAccount.RecordCount > 0 Then rsAccount.MoveFirst
    
    '' """"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
    '' // 2002/11/21       -- �s�W���e--
    '' // �p�G���ǤJ���O�H��
    '' // �h�N���O�H������Ƥ@�֧�s�һݪ� SQL �y�k��_��  --
    '' -------------------------------------------------------------------------------------------------------------
    If Len(pEmpNO) <> 0 Then
        strUpdateEmpSql = ",ClctEn = '" & pEmpNO & "', ClctName ='" & pEmpName & "'"
    End If
    '' """"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
    While Not rsAccount.EOF
        With rsAccount
            .Fields("RealAmt") = rsAccount("ShouldAmt")
            .Fields("RealDate") = NoZero(Format(vRealDate, "####/##/##"))
            .Fields("Updtime") = GetDTString(strNowTime)
            .Fields("UpdEn") = vUpdEn
            .Fields("UCCode") = Null
            .Fields("UCName") = Null
            '#8555 don't update the value when fubon isn't integrate by kin 2020/01/13
            If Len(vCMCode) > 0 Then
                .Fields("CMCode") = vCMCode
                .Fields("CMName") = vCMName
            End If
            
            .Fields("PTCode") = vPTCode
            .Fields("PTName") = vPTName
            If Len(pEmpNO) <> 0 Then
                .Fields("ClctEn") = pEmpNO
                .Fields("ClctName") = pEmpName
                
                '************************************************************************************************************************
                '#3417 �p�G���O�H������,�hSO033.NOTE�n�[�J�Ȧ�N�X By Kin 2008/01/22
                If Len(vBankId) <> 0 Then
                    .Fields("Note") = IIf(IsNull(.Fields("Note")), "�Ȧ�N�X:" & vBankId, .Fields("Note") & " " & "�Ȧ�N�X:" & vBankId)
                End If
                '************************************************************************************************************************
            End If
            '#5494 �ŷs��޳Ƶ���n��32~40�r�� By Kin 2010/02/03
            If vNote <> Empty Then
                .Fields("Note") = IIf(IsNull(.Fields("Note")), vNote, .Fields("Note") & ";" & vNote)
            End If
            .Fields("FirstTime") = GetDTString(strNowTime)
            .Fields("AuthorizeNo") = pAuthorizeNo
             Call AlterSO003(rsAccount)
            .Update
        End With
        '*******************************************************
        '#3697 �p�⦨�\�����B By Kin 2007/01/02
        lngAmtTotal = lngAmtTotal + rsAccount("ShouldAmt")
        '*******************************************************
        rsAccount.MoveNext
    Wend
    lngCount = lngCount + 1
    Set rsAccount = Nothing
    UpdAccount = True
 Exit Function
ChkErr:
   UpdAccount = False
   lngAmtTotal = 0
   ErrSub "UpdateFun", "UpdAccount"
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
'�ھڶǨӤ��t�Τ��w�ɶ��r���ন�T�w���ɶ��r��
'Parameter1: DateTime
'Return: Format DateTime String
Public Function GetDTString(dtm As String) As String
  On Error GoTo ChkErr
    'GetDTString = Format(dtm, "ee/mm/dd hh:mm:ss")
     GetDTString = AnnoDominiToTaiwanCalendar(dtm, GiYear) & Format(dtm, "/MM/DD HH:MM:SS")
  Exit Function
ChkErr:
   ErrSub "DealFun", "GetDTString"
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
        If Len(varValue & "") = 0 Then
            GetNullString = "NULL"
        Else
            GetNullString = varValue
        End If
    ElseIf lngVType = giStringV Then
        If Len(varValue & "") = 0 Then
            GetNullString = "NULL"
        Else
            varValue = Replace(varValue, "'", "''")
            GetNullString = "'" & varValue & "'"
        End If
    End If
    Exit Function
ChkErr:
  Call ErrSub("Utility5", "GetNullString")
End Function

Public Function InitData() As Boolean
On Error GoTo ChkErr
   Dim rsTmp As New ADODB.Recordset
   Dim strSQL As String
   
   strSQL = "Select Para3, Para5, Para8 From " & frmUnion.TableOwnerName & "So043"
   If Not GetRS(rsTmp, strSQL, gcnGi) Then InitData = False: Exit Function
   If rsTmp.EOF = False Then
      intPara3 = rsTmp("Para3")
      intPara5 = rsTmp("Para5")
      intPara8 = rsTmp("Para8")
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
                   "�J�b�`���B�G" & lngAmtTotal & "��" & vbCrLf & _
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
                  Optional strTableName As String, _
                  Optional strWhere As String)
  On Error GoTo ChkErr
    objLst.SetFldName1 strFldName1
    objLst.SetFldName2 strFldName2
    If lngFldLen1 > 0 Then objLst.FldLen1 = lngFldLen1
    If lngFldLen2 > 0 Then objLst.FldLen2 = lngFldLen2
    If lngWidth1 > 0 Then objLst.FldWidth1 = lngWidth1
    If lngWidth2 > 0 Then objLst.FldWidth2 = lngWidth2
    objLst.SetTableName strTableName
    objLst.SendConn gcnGi
    objLst.Filter = strWhere
  Exit Sub
ChkErr:
    Call ErrSub("DealFun", "SetLst")
End Sub

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
    MsgBox strSQL
    ErrSub "UpdateFun", "GetRs"
    On Error Resume Next
    If DirectUnload Then Unload Screen.ActiveForm
End Function

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
Public Function RPxx(strValue As String) As String
  On Error Resume Next
    strValue = Trim(Replace(strValue, Chr(13), "", 1, , vbTextCompare))
    strValue = Trim(Replace(strValue, Chr(10), "", 1, , vbTextCompare))
    RPxx = Trim(Replace(strValue, Chr(9), "", 1, , vbTextCompare))
End Function
Public Function MidMbcs(ByVal str As String, start, Length)
  On Error GoTo ChkErr
    MidMbcs = StrConv(MidB(StrConv(str, vbFromUnicode), start, Length), vbUnicode)
  Exit Function
ChkErr:
    ErrSub "Sys_Lib", "MidMbcs"
End Function

''2003/08/18 �o�@�q�s�W�g���ʦ��O����

'Public Function AlterSO003(rsSO033 As ADODB.Recordset, Optional ByVal blnAddSo003 As Boolean = True, _
'                Optional ByVal blnUpdAmt As Boolean = True, Optional ByVal blnUpdClctDate As Boolean = True, _
'                Optional ByVal blnChkMustAnswer As Boolean = False, Optional ByRef strSeqNo As String, _
'                Optional ByVal blnMultiCount As Boolean = False) As Boolean
'    On Error GoTo ChkErr
'
'    Dim rsAlter As New ADODB.Recordset
'    Dim rsTmp As New ADODB.Recordset
'    Dim lngPeriodFlag As Boolean
'    Dim Ypara5 As Long, Ypara8 As Long
'    Dim Ypara9 As Long, YPara24 As Long
'    Dim lngCompCode As Long, lngBad As Long
'    Dim blnFlag As Boolean
'    Dim blnDiscount As Boolean
'
'        '�P�_�O�_���g���ʦ��O����
'
'        lngPeriodFlag = Val(GetRsValue("Select PeriodFlag From " & frmUnion.TableOwnerName & "Cd019 Where CodeNo = " & rsSO033("CitemCode")) & "")
'        If Len(rsSO033("STCode") & "") > 0 Then lngBad = Val(GetRsValue("Select RefNo From " & frmUnion.TableOwnerName & "CD016 Where CodeNo = " & rsSO033("STCode")) & "")
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
'        lngCompCode = Val(GetRsValue("Select CompCode From " & frmUnion.TableOwnerName & "So001 Where CustId = " & rsSO033("CustId")) & "")
'
''        If rsSO033("OrderNo") & "" <> "" Then
'            If Val(rsSO033("SeqNo") & "") = 0 Then
'                If Not GetRS(rsAlter, "Select RowId,SO003.* From " & frmUnion.TableOwnerName & "So003 Where 1 = 0 ", gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
'            Else
'                If Not GetRS(rsAlter, "Select RowId,SO003.* From " & frmUnion.TableOwnerName & "So003 Where CustId = " & rsSO033("CustId") & " And CitemCode = " & rsSO033("CitemCode") & " And SeqNo = " & rsSO033("SeqNo"), gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
'            End If
''        End If
'        With rsAlter
'            If .RecordCount = 0 Then
'                '�ˬd�O���n�s�W�g���ʦ��O����
'                If blnAddSo003 Then
'                    If blnChkMustAnswer Then
'                        If MsgBox("�L�g���ʦ��O����,�O�_�n�s�W??", vbQuestion + vbYesNo + vbDefaultButton2, "����") = vbNo Then
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
'                            If Not GetRS(rsTmp, "Select Nvl(AddCitemAccount,0) AddCitemAccount,SnactionDate From " & frmUnion.TableOwnerName & "SO106 Where CustId = " & rsSO033("CustId") & _
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
'                        If Not GetRS(rsTmp, "Select CodeNo,Description From " & frmUnion.TableOwnerName & "CD031 Where RefNo=1", gcnGi) Then Exit Function
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
'
'            If .Fields("StopDate") >= rsSO033("RealStopDate") Then AlterSO003 = True: Exit Function
'            If rsSO033("DiscountCode") & "" <> "" Then
'                '93/04/20 Jacky �u�ݿ�k�վ�
'                If Format(rsSO033("RealStartDate") & "", "yyyymmdd") <= Format(.Fields("DiscountDate2") & "", "yyyymmdd") Or blnAddSo003 Then
'                    blnDiscount = True
'                    '.Fields("Amount") = NoZero(rsSO033("OldAmt"), True)
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
'            If blnAddSo003 Or (blnUpdAmt And Ypara8 >= 1) Then
'                '�Y<Para9>=1, �h 1: Amount = OldAmt, 2: Amount = ShouldAmt , 3: Amount = RealAmt
'                If Not blnDiscount Then
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
'                     If Not ExecuteCommand("Delete From " & frmUnion.TableOwnerName & "SO003 Where RowId= '" & .Fields("RowId") & "'") Then Exit Function
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
'ChkErr:
'    ErrSub "UpdateFun", "AlterSo003"
'End Function

Public Function GetSystemPara(ByRef rs As ADODB.Recordset, ByVal strCompCode As String, _
    ByVal strServiceType As String, ByVal strTable As String, Optional ByVal strParaField As String = "", _
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
        '����CompCode ��ServiceType ������
        strTable = frmUnion.TableOwnerName & strTable
        
        If strCompCode <> "" And strServiceType <> "" Then
            If Not GetRS(rs, "Select " & strParaField & " From " & strTable & " Where CompCode = " & strCompCode & " And ServiceType " & strSQL) Then Exit Function
            blnFlag = Not rs.EOF
        End If
        If Not blnFlag Then
            '�A��CompCode ����
            If strCompCode <> "" Then
                If Not GetRS(rs, "Select " & strParaField & " From " & strTable & " Where CompCode = " & strCompCode) Then Exit Function
                blnFlag = Not rs.EOF
            End If
            '�̫��u�n���o�� �Ѽ��ɪ�
            If Not blnFlag Then
                If Not GetRS(rs, "Select " & strParaField & " From " & strTable) Then Exit Function
                blnFlag = Not rs.EOF
            End If
        End If
        If blnShowMsg And Not blnFlag Then MsgBox "�b�t�ΰѼ���[" & strTable & "] �䤣�� ���q�O�N���� [" & strCompCode & "] ,�B�A�����O��: [" & strServiceType & "]�����,�Ьd��!!", vbCritical, "�T��"
        GetSystemPara = True
    Exit Function
ChkErr:
    ErrSub "UpdateFun", "GetSystemPara"
End Function

Public Function GetFieldValue(rs As ADODB.Recordset, _
        strFieldName As String, Optional intType As Integer = 0)
    On Error GoTo ChkErr
    Dim varField As Variant
        'intType 0:Value  1:OriginalValue
        If intType = 0 Then
            varField = rs.Fields(strFieldName).Value
        ElseIf intType = 1 Then
            varField = rs.Fields(strFieldName).OriginalValue
        End If
        
        If rs.Fields(strFieldName).Type = 131 Or rs.Fields(strFieldName).Type = 139 Then
            If intType = 0 Then
                If Len(rs.Fields(strFieldName).Value & "") > 0 Then
                    varField = Val(rs.Fields(strFieldName).Value)
                End If
            ElseIf intType = 1 Then
                If Len(rs.Fields(strFieldName).OriginalValue & "") > 0 Then
                    varField = Val(rs.Fields(strFieldName).OriginalValue)
                End If
            End If
        End If
        GetFieldValue = varField
    Exit Function
ChkErr:
    ErrSub "UpdateFun", "GetFiledValue"
End Function

Public Function NoZero(strVal As Variant, Optional ZeroFlag As Boolean = False) As Variant
  On Error GoTo ChkErr
    strVal = strVal & ""
    If Trim(strVal) = "" Then
        If ZeroFlag = True Then
            NoZero = 0
        Else
            NoZero = Null
        End If
    Else
        NoZero = strVal
    End If
  Exit Function
ChkErr:
    ErrSub "Sys_Lib", "NoZero"
End Function


Public Sub CloseRecordset(rs As ADODB.Recordset)
  On Error Resume Next
    rs.CancelUpdate
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
End Sub
Public Function GetInvoiceNo3(strTableName As String, Optional Conn As ADODB.Connection) As String
On Error GoTo ChkErr
    Dim strSeq As String
    Dim strInv
    strSeq = strOwnerName & "S_" & strTableName
    If Conn Is Nothing Then Set Conn = gcnGi
    strInv = GetRsValue("SELECT Ltrim(To_Char(" & strSeq & ".NextVal, '09999999')) FROM Dual", Conn)
    GetInvoiceNo3 = strInv
    Exit Function
ChkErr:
  ErrSub "Sys_Lib", "GetInvoiceNo3"
End Function

Public Function ExecuteCommand(strSQL As String, Optional Conn As ADODB.Connection, Optional ByRef lngAffected As Long) As Boolean
    On Error GoTo ChkErr
        If Conn Is Nothing Then Set Conn = gcnGi
        Conn.Execute strSQL, lngAffected
        ExecuteCommand = True
    Exit Function
ChkErr:
    ErrSub FormName, "ExecuteCommand, SQL : [" & strSQL & "] "
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
'#3049 �]������H�U�אּ�n�P�_�^���ɪ��B�A�ҥH�N��Function�W�ߥX�ӡA�H�K�䥦�{���y�����D
'#3417 �W�[�Ѽ�vBankId,���p����J���O�H��,�hSO033.NOTE�n��J�Ȧ�N�X
Public Function ChkDataTo3313(ByVal vBillNo As String, ByVal vRealAmt As String, _
                        ByVal vRealDate As String, ByVal vUpdEn As String, _
                        ByVal vCMCode As String, ByVal vCMName As String, _
                        ByVal vPTCode As String, ByVal vPTName As String, _
                        Optional pEmpNO As String = "", Optional pEmpName As String = "", _
                        Optional pAuthorizeNo As String = " ", _
                        Optional p_frmCreditCardType As mCreditCardType = lngUnionChinese, _
                        Optional pCitemCodeChkData As String, _
                        Optional pstrServiceType As String, _
                        Optional ByVal vBankId As String = Empty, _
                        Optional ByVal vNote As String = Empty) As Boolean
                        
  On Error GoTo ChkErr
    Dim strSQL As String
    Dim strSQL2 As String
    Dim rsChk As New ADODB.Recordset
    Dim rsCust As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim strBillField As String
    Dim strWhere As String
    Dim rsCheckBillNoCancelFlag As New ADODB.Recordset
    Dim lngBudgetPeriod As Long
    Dim strResult As String
    Dim strCD013RefNo As String
    Dim strCustId As String
    Dim lngTotalShouldAmt As Long
    Dim lngTotalRealAmt As Long
  '''�H�U�� BillNo ��� �令 MediaBillNo
    strCD013RefNo = " MAX(DECODE(NVL(CD013.PAYOK,0),1,3," & _
                "DECODE(NVL(CD013.REFNO,0),3,3,DECODE(NVL(CD013.REFNO,0),7,3,0)))) REFNO "
    If blnMultiAccount = True Then
        strBillField = "SO033.MediaBillno"
    Else
        strBillField = "SO033.Billno"
    End If
    If pCitemCodeChkData <> Empty Then strSQL2 = " And SO033.CitemCode " & pCitemCodeChkData
   
    '****************************************************************************************************************************************************************************
    '#4010 UCCode���ѦҸ���3��,�n�������` By Kin 2008/07/29
    '#5564 CD013.PAYOK=1��CD013.REFNO=3,7�����d�x�w�� By Kin 2010/05/19
    strSQL = "Select " & _
            "Nvl(sum(SO033.ShouldAmt),0) as TotalShouldAmt, Nvl(sum(SO033.RealAmt),0) as TotalRealAmt, SO033.RealDate, SO033.CustId," & strCD013RefNo & _
            " From " & _
             frmUnion.TableOwnerName & "SO033," & frmUnion.TableOwnerName & "CD013 " & _
             "Where SO033.UCCODE IS NOT NULL AND " & _
             strBillField & " = '" & vBillNo & "' And SO033.CancelFlag <> 1 And SO033.UCCode=CD013.CodeNo" & strSQL2 & " Group By SO033.CustId,SO033.RealDate"
    '***************************************************************************************************************************************************************************
    '7179
    If blnCrossCustCombine Then
         strSQL = "Select A.*,SO001.AMduid, (Select MainCustId From " & frmUnion.TableOwnerName & "SO202 Where MduId = SO001.AmduId) MainCustId " & _
                             " From (" & strSQL & " ) A," & frmUnion.TableOwnerName & "SO001 " & _
                             " Where A.CustId = SO001.CustId "
    End If
    If rsChk.State = adStateOpen Then rsChk.Close
    If Not GetRS(rsChk, strSQL, gcnGi) Then
       ChkDataTo3313 = False
       Exit Function
    End If
    If rsChk.EOF Or rsChk.BOF Then
        strSQL = "Select count(CustId) C " & _
                  "From " & _
                  frmUnion.TableOwnerName & "SO033 " & _
                  "Where " & _
                  strBillField & " = '" & vBillNo & "' And CancelFlag <> 1 Group By CustId,RealDate"
        If GetRS(rsCheckBillNoCancelFlag, strSQL, gcnGi, adUseClient, adOpenDynamic) = True Then
            If rsCheckBillNoCancelFlag.EOF Or rsCheckBillNoCancelFlag.BOF Then
                ErrFile.WriteLine ("��ڽs���G " & vBillNo & " �d�L���渹�I")
            Else
                ErrFile.WriteLine ("��ڽs���G " & vBillNo & " ���渹����ڤw���I")
                '*****************************************************************************************************************************
                '#3417 ���Ѯɭn��sUCCode�PUCName By Kin 2007/12/13
                '#3850 ��sUCCode��,�n�h�W�[ UCCode Is Not NUll By Kin 2008/04/08
                If blnFailUpd Then
                    strFailUpdSQL = strFailUpdSQL & "'��ڽs���G " & vBillNo & " ���渹����ڤw���I'" & _
                                  " Where " & IIf(blnMultiAccount, "MediaBillNo='", "BillNo='") & vBillNo & "'" & _
                                   " And CancelFlag <> 1 And UCCODE IS NOT NULL "

                    gcnGi.Execute strFailUpdSQL
                End If
                '*****************************************************************************************************************************
            End If
        Else
            ErrFile.WriteLine ("��ڽs���G " & vBillNo & " �d�L���渹�I")
        End If
        lngErrCount = lngErrCount + 1
    Else
        '7179*********************************************************************************************
        lngTotalRealAmt = 0
        lngTotalShouldAmt = 0
        
        If blnCrossCustCombine Then
            rsChk.MoveFirst
            If rsChk.RecordCount = 1 Then
                If rsChk("CustId") <> rsChk("MainCustId") Then
                    strCustId = rsChk("CustId")
                End If
            Else
                If Len(rsChk("MainCustId") & "") > 0 Then
                    strCustId = rsChk("MainCustId") & ""
                Else
                    strCustId = rsChk("CustId") & ""
                End If
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
            strRetMsg = ChkDiffCust(rsChk) & ""  '�ˮ֤��P���Ƚs
            If strRetMsg <> "" Then
                ErrFile.WriteLine ("��ڽs���G " & vBillNo & " �ۦP��ڽs�����Ȥ�s�����P�I �Ƚs�G" & strRetMsg)
                If blnFailUpd Then
                    strFailUpdSQL = strFailUpdSQL & "'��ڽs���G" & vBillNo & _
                            " �ۦP��ڽs�����Ȥ�s�����P�I �Ƚs�G" & strRetMsg & "'" & _
                            " Where " & IIf(blnMultiAccount, "MediaBillNo='", "BillNo='") & vBillNo & "'" & _
                               " And CancelFlag <> 1 And UCCode is Not Null"
                    gcnGi.Execute strFailUpdSQL
                
                End If
                lngErrCount = lngErrCount + 1
                ChkDataTo3313 = True
                Exit Function
            End If
        End If
        '***********************************************************************
        
        If rsChk("RealDate").Value & "" <> "" And lngTotalRealAmt > 0 Then
            'ErrFile.WriteLine ("�Ȥ�s��: " & rsChk("CustId").Value & "��ڽs���G " & vBillNo & " �ɮפ��w�]�t�ꦬ���B$ " & rsChk("TotalRealAmt").Value & " �A�Ьd�֡I")
            ErrFile.WriteLine ("�Ȥ�s��: " & rsChk("CustId").Value & "��ڽs���G " & vBillNo & " �ɮפ��w�]�t�ꦬ���B$ " & lngTotalRealAmt & " �A�Ьd�֡I")
            '*****************************************************************************************************************************
            '#3417 ���Ѯɭn��sUCCode�PUCName By Kin 2007/12/13
            If blnFailUpd Then
                strFailUpdSQL = strFailUpdSQL & "'�Ȥ�s��: " & rsChk("CustId").Value & "��ڽs���G " & vBillNo & _
                                " �ɮפ��w�]�t�ꦬ���B$ " & rsChk("TotalRealAmt").Value & " �A�Ьd�֡I'" & _
                              " Where " & IIf(blnMultiAccount, "MediaBillNo='", "BillNo='") & vBillNo & "'" & _
                               " And CancelFlag <> 1 And UCCode is Not Null"
                gcnGi.Execute strFailUpdSQL
            End If
            '*****************************************************************************************************************************
            
            lngErrCount = lngErrCount + 1
        Else
            '#3122 �ˮ֬O�_���W�L�鵲��� By Kin Add 2007/03/29
            If Not ChkOverCloseDate(Format(vRealDate, "####/##/##"), objStorePara.CompCode, objStorePara.ServiceType, strResult) Then
                'ErrFile.WriteLine ("�Ȥ�s��:  " & rsChk("CustId").Value & "��ڽs���G " & vBillNo & "  ��ڦ��ڤ�G" & Format(str(Val(vRealDate) - 19110000), "##/##/##") & " ���~��]�G" & strResult)
                ErrFile.WriteLine ("�Ȥ�s��:  " & strCustId & "��ڽs���G " & vBillNo & "  ��ڦ��ڤ�G" & Format(str(Val(vRealDate) - 19110000), "##/##/##") & " ���~��]�G" & strResult)
                
                '*****************************************************************************************************************************
                '#3417 ���Ѯɭn��sUCCode�PUCName By Kin 2007/12/13
                If blnFailUpd Then
'                    strFailUpdSQL = strFailUpdSQL & "'�Ȥ�s��: " & rsChk("CustId").Value & "��ڽs���G " & vBillNo & _
'                                    "  ��ڦ��ڤ�G" & Format(str(Val(vRealDate) - 19110000), "##/##/##") & " ���~��]�G" & strResult & "'" & _
'                                  " Where " & IIf(blnMultiAccount, "MediaBillNo='", "BillNo='") & vBillNo & "'" & _
'                                   " And CancelFlag <> 1 And UCCode is Not Null" & strSQL2
                    strFailUpdSQL = strFailUpdSQL & "'�Ȥ�s��: " & strCustId & "��ڽs���G " & vBillNo & _
                                    "  ��ڦ��ڤ�G" & Format(str(Val(vRealDate) - 19110000), "##/##/##") & " ���~��]�G" & strResult & "'" & _
                                  " Where " & IIf(blnMultiAccount, "MediaBillNo='", "BillNo='") & vBillNo & "'" & _
                                   " And CancelFlag <> 1 And UCCode is Not Null" & strSQL2
                    gcnGi.Execute strFailUpdSQL
                End If
                '*****************************************************************************************************************************

                lngErrCount = lngErrCount + 1
                ChkDataTo3313 = True
                Exit Function
            End If
            '#4010 UCCode�ѦҬ�3�ɭn�������` By Kin 2008/07/29
            '#5564 CD013.PAYOK=1��CD013.REFNO=3,7�����d�x�w�� By Kin 2010/05/19
            If Val(rsChk("RefNo")) = 3 Then
                'ErrFile.WriteLine ("�Ƚs�s��: " & rsChk("CustId").Value & "��ڽs���G " & vBillNo & " �d�O�w���I")
                ErrFile.WriteLine ("�Ƚs�s��: " & strCustId & "��ڽs���G " & vBillNo & " �d�O�w���I")
                If blnFailUpd Then
'                    strFailUpdSQL = strFailUpdSQL & "'�Ȥ�s��: " & rsChk("CustId").Value & "��ڽs���G " & vBillNo & _
'                                    " �d�O�w���I'" & _
'                                  " Where " & IIf(blnMultiAccount, "MediaBillNo='", "BillNo='") & vBillNo & "'" & _
'                                   " And CancelFlag <> 1 And UCCode is Not Null"
                    strFailUpdSQL = strFailUpdSQL & "'�Ȥ�s��: " & strCustId & "��ڽs���G " & vBillNo & _
                                    " �d�O�w���I'" & _
                                  " Where " & IIf(blnMultiAccount, "MediaBillNo='", "BillNo='") & vBillNo & "'" & _
                                   " And CancelFlag <> 1 And UCCode is Not Null"
                    gcnGi.Execute strFailUpdSQL
                End If
                lngErrCount = lngErrCount + 1
                ChkDataTo3313 = True
                Exit Function
            End If
            
            If lngTotalShouldAmt <> Val(vRealAmt) Then
'                ErrFile.WriteLine ("�Ȥ�s��: " & rsChk("CustId").Value & "��ڽs���G " & vBillNo & " ���B���X�A�ɮ��������B$ " & rsChk("TotalShouldAmt").Value & " �A�Ϥ��w�����B$ " & Val(vRealAmt))
                ErrFile.WriteLine ("�Ȥ�s��: " & strCustId & "��ڽs���G " & vBillNo & " ���B���X�A�ɮ��������B$ " & lngTotalShouldAmt & " �A�Ϥ��w�����B$ " & Val(vRealAmt))
                
                '*****************************************************************************************************************************
                '#3417 ���Ѯɭn��sUCCode�PUCName By Kin 2007/12/13
                If blnFailUpd Then
'                    strFailUpdSQL = strFailUpdSQL & "'�Ȥ�s��: " & rsChk("CustId").Value & "��ڽs���G " & vBillNo & _
'                                    " ���B���X�A�ɮ��������B$ " & rsChk("TotalShouldAmt").Value & " �A�Ϥ��w�����B$ " & Val(vRealAmt) & "'" & _
'                                  " Where " & IIf(blnMultiAccount, "MediaBillNo='", "BillNo='") & vBillNo & "'" & _
'                                   " And CancelFlag <> 1 And UCCode is Not Null"
                     strFailUpdSQL = strFailUpdSQL & "'�Ȥ�s��: " & strCustId & "��ڽs���G " & vBillNo & _
                                    " ���B���X�A�ɮ��������B$ " & lngTotalShouldAmt & " �A�Ϥ��w�����B$ " & Val(vRealAmt) & "'" & _
                                  " Where " & IIf(blnMultiAccount, "MediaBillNo='", "BillNo='") & vBillNo & "'" & _
                                   " And CancelFlag <> 1 And UCCode is Not Null"
                    gcnGi.Execute strFailUpdSQL
                End If
                '*****************************************************************************************************************************
                
                lngErrCount = lngErrCount + 1
            Else
                If rsCust.State = adStateOpen Then rsCust.Close
                '#4022 �ҥΦh�C�^�P�_���P By Kin 2008/07/22
                If Not blnMultiAccount Then
                    strSQL = "Select CustStatusCode,WipCode3 ,CustStatusName From " & frmUnion.TableOwnerName & "So002 Where CustId = " & rsChk("CustId").Value & "  AND SERVICETYPE  = '" & pstrServiceType & "'"
                Else
                    'strSQL = "Select CustStatusCode,WipCode3 ,CustStatusName From " & frmUnion.TableOwnerName & "So002 Where CustId = " & rsChk("CustId").Value & "  AND MediabillNo = '" & vBillNo & "'"
                    '7179
                    strSQL = "Select distinct SO002.CustStatusCode,SO002.WipCode3,SO002.CustStatusName,So002.ServiceType From " & frmUnion.TableOwnerName & "So002," & frmUnion.TableOwnerName & "SO033 " & _
                                    " Where SO002.ServiceType=SO033.ServiceType And SO002.CustId = SO033.CustId And MediaBillNo = '" & vBillNo & "'" & _
                                    IIf(blnCrossCustCombine, " And 1=1 ", "")
                End If
                Call GetRS(rsCust, strSQL, gcnGi)
               '�D���`�Ȥ�h���@log�A�J�b
               Do While Not rsCust.EOF
                    If rsCust("CustStatusCode").Value & "" <> 1 Then
                                       ErrFile.WriteLine ("�Ȥ�s��: " & rsChk("CustId").Value & " ��ڽs���G " & vBillNo & " ���D���`�Ȥ� " & " �Ȥ᪬�A: " & rsCust("CustStatusName").Value)
                                       lngStatusCount = lngStatusCount + 1
                    '���`�Ȥ�BCM������
                    ElseIf rsCust("CustStatusCode").Value & "" = 1 And rsCust("WipCode3") = 16 Then
                                       ErrFile.WriteLine ("�Ȥ�s��: " & rsChk("CustId").Value & " ��ڽs���G " & vBillNo & " ��CM�������Ȥ� ")
                                       lngStatusCount = lngStatusCount + 1
                    End If
                    rsCust.MoveNext
                Loop
                If Not GetRS(rsTmp, "Select RealPeriod,CustId,CompCode,ServiceType,CitemCode,FaciSeqNo From " & strOwnerName & "So033 Where UCCODE IS NOT NULL AND " & _
                       strBillField & " = '" & vBillNo & "' And CancelFlag <> 1 ") Then Exit Function
                Do While Not rsTmp.EOF
                     lngBudgetPeriod = Val(GetRsValue("Select Nvl(BudgetPeriod,0) From " & strOwnerName & "SO003 Where Custid = " & rsTmp("CustId") & " And CompCode = " & rsTmp("CompCode") & " And ServiceType = '" & rsTmp("ServiceType") & "' " & _
                         " And CitemCode = " & rsTmp("CitemCode") & " And FaciSeqNo = '" & rsTmp("FaciSeqNo") & "'") & "")
                     If rsTmp("RealPeriod") > lngBudgetPeriod And lngBudgetPeriod > 0 Then
                         Call ErrFile.WriteLine("�J�b���Ƥj������Ѿl����,�L�k�J�b!!")
                         
                        '*****************************************************************************************************************************
                        '#3417 ���Ѯɭn��sUCCode�PUCName By Kin 2007/12/13
                        If blnFailUpd Then
                            strFailUpdSQL = strFailUpdSQL & "'�J�b���Ƥj������Ѿl����,�L�k�J�b!!'" & _
                                          " Where " & IIf(blnMultiAccount, "MediaBillNo='", "BillNo='") & vBillNo & "'" & _
                                          " And CancelFlag <> 1 And UCCode is Not Null"
                            gcnGi.Execute strFailUpdSQL
                        End If
                        '*****************************************************************************************************************************
                        
                         lngErrCount = lngErrCount + 1
                         ChkDataTo3313 = True
                         Exit Function
                     End If
                     rsTmp.MoveNext
                Loop
               '�����sSO033
               '#5494 �ŷs���32~40�r���n��J�Ƶ���,�ҥH�W�[vNote���Ѽ� By Kin 2010/02/03
                If UpdAccount(vBillNo, vRealAmt, vRealDate, vUpdEn, vCMCode, vCMName, _
                              vPTCode, vPTName, pEmpNO, pEmpName, pAuthorizeNo, _
                              pCitemCodeChkData, , vNote) = False Then
                    ChkDataTo3313 = False
                    Exit Function
                End If  '' If UpdAccount
            End If
        End If
    End If
    If rsChk.State = adStateOpen Then rsChk.Close
    Set rsChk = Nothing
    If rsCust.State = adStateOpen Then rsCust.Close
    Set rsCust = Nothing
    Call CloseRecordset(rsTmp)
    ChkDataTo3313 = True
    Exit Function
ChkErr:
    ErrSub "UpdateFun3313", "ChkDataTo3313"
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
  ErrSub "UpdateFun3313", "ChkDiffCust"
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
    ErrSub "UpdateFun3313", "ChkOverCloseDate"
End Function
'#3122 ���o�t�ΰѼƪ��� By Kin 2007/3/29 Add
Public Function GetSystemParaItem(strField As String, _
    strCompCode As String, strServiceType As String, strTable As String, _
    Optional blnShowMsg As Boolean = False, Optional blnServiceTypeUseLike As Boolean = False) As Variant
    On Error Resume Next
    Dim rs As New ADODB.Recordset
        If Not GetSystemPara(rs, strCompCode, strServiceType, strTable, strField, False, blnServiceTypeUseLike) Then Exit Function
        GetSystemParaItem = rs(0)
        CloseRecordset rs
End Function

