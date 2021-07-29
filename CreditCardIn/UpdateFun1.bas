Attribute VB_Name = "UpdateFun"
Option Explicit
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Const gMsgIsDataOK = "�������n���,������ !! "
Public lngCount As Long           '�O�����\����
Public lngErrCount As Long        '�O�����~����
Public lngStatusCount As Long     '�O���D���`�ᵧ��
Public lngTime As Long             '�O���Ҫ�ɶ�
Public strNowTime As String        '�O���ثe�ɶ�
Public Fso As New FileSystemObject
Public File As TextStream
Public ErrFile As TextStream     '���~���log��
Private intPara3 As Integer
Private intPara5 As Integer
Private intPara8 As Integer
Public gcnGi As New ADODB.Connection
Public strUserName As String
Private strFilePath As String
Public objStorePara As clsStoreParameter
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

Public blnMultiAccount As Boolean     ''  �O�� SO043 �� SO043.para24  �A�H�P�_�O�_�Φh�b��]�w
Public strOwnerName As String

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
    Dim retval As Variant
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
    retval = MsgBox(strErr, vbYesNo, "�t�ΰ�����~..(�� <�O> �C�L���~�T��)!!")
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
    
    'LogFile = ReadFromINI(GetIniPath1, "GICMISPath", "ErrLogPath")
    LogFile = strFilePath
    If Not ChkDirExist(LogFile) Then MkDir LogFile
    
    If Right(LogFile, 1) <> "\" Then LogFile = LogFile & "\"
    LogFile = LogFile & "SYSERR.TXT"
    
    ErrString = Replace(ErrString, vbCrLf & vbCrLf, vbCrLf)
    ErrString = "�o�Ϳ��~���t�� : �a�Q�ȪA��B�޲z�t��" & vbCrLf & _
                "�o�Ϳ��~���q�� : " & CreateObject("WScript.NetWork").ComputerName & vbCrLf & _
                "�ϥΪ̦W�� : " & strUserName & vbCrLf & ErrString
    
    Fnum = FreeFile
    Open LogFile For Append As Fnum
    Print #Fnum, ErrString & vbCrLf & String(66, "-") & vbCrLf
    Close Fnum
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

Public Function GetRsValue(strsql As String, Optional cnnObj As ADODB.Connection, Optional MsgNoData As String) As Variant
  On Error GoTo ChkErr
    Dim rs As New ADODB.Recordset
    If cnnObj Is Nothing Then Set cnnObj = gcnGi
    If Not GetRS(rs, strsql, cnnObj) Then Exit Function
    If rs.RecordCount = 0 Then
        If Len(MsgNoData) > 0 Then
            MsgBox MsgNoData, vbExclamation, "����!"
        End If
        GetRsValue = Null
    Else
        GetRsValue = rs(0)
    End If
  Exit Function
ChkErr:
    ErrSub "DealFun", "GetRsValue"
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
    Call ErrSub("Sys_Lib", "GetString")
End Function

Public Function ChkData(ByVal vBillNo As String, ByVal vRealAmt As String, _
                        ByVal vRealDate As String, ByVal vUpdEn As String, _
                        ByVal vCMCode As String, ByVal vCMName As String, _
                        ByVal vPTCode As String, ByVal vPTName As String, _
                        Optional pEmpNO As String = "", Optional pEmpName As String = "", _
                        Optional pAuthorizeNo As String = " ") As Boolean
                        
  On Error GoTo ChkErr
  Dim strsql As String
  Dim rsChk As New ADODB.Recordset
  Dim rsCust As New ADODB.Recordset
  Dim strBillField As String
  
  '''�H�U�� BillNo ��� �令 MediaBillNo
  
  If blnMultiAccount = True Then
       strBillField = "MediaBillno"
  Else
       strBillField = "Billno"
  End If
  
     strsql = "Select " & _
                 "Nvl(sum(ShouldAmt),0) as TotalShouldAmt, Nvl(sum(RealAmt),0) as TotalRealAmt, RealDate, CustId " & _
                 "From " & _
                 frmUnion.TableOwnerName & "SO033 " & _
                 "Where UCCODE IS NOT NULL AND " & _
                 strBillField & " = '" & vBillNo & "' And CancelFlag <> 1 Group By CustId,RealDate"
    
    If rsChk.State = adStateOpen Then rsChk.Close
  
    If Not GetRS(rsChk, strsql, gcnGi) Then
       ChkData = False
       Exit Function
    End If
   
    If rsChk.EOF Or rsChk.BOF Then
       ErrFile.WriteLine ("��ڽs���G " & vBillNo & " �d�L���渹�I")
       lngErrCount = lngErrCount + 1
    Else
       If rsChk("RealDate").Value & "" <> "" And rsChk("TotalRealAmt").Value > 0 Then
          ErrFile.WriteLine ("�Ȥ�s��: " & rsChk("CustId").Value & "��ڽs���G " & vBillNo & " �ɮפ��w�]�t�ꦬ���B$ " & rsChk("TotalRealAmt").Value & " �A�Ьd�֡I")
          lngErrCount = lngErrCount + 1
       Else
            If rsChk("TotalShouldAmt").Value <> Val(vRealAmt) Then
               ErrFile.WriteLine ("�Ȥ�s��: " & rsChk("CustId").Value & "��ڽs���G " & vBillNo & " ���B���X�A�ɮ��������B$ " & rsChk("TotalShouldAmt").Value & " �A�Ϥ��w�����B$ " & Val(vRealAmt))
               lngErrCount = lngErrCount + 1
            Else
                    If rsCust.State = adStateOpen Then rsCust.Close
                          strsql = "Select CustStatusCode,WipCode3 From " & frmUnion.TableOwnerName & "So002 Where CustId = " & rsChk("CustId").Value
                          Call GetRS(rsCust, strsql, gcnGi)
                    
                    '�D���`�Ȥ�h���@log�A�J�b
                    If rsCust("CustStatusCode").Value & "" <> 1 Then
                                       ErrFile.WriteLine ("�Ȥ�s��: " & rsChk("CustId").Value & " ��ڽs���G " & vBillNo & " ���D���`�Ȥ� " & " �Ȥ᪬�A: " & rsCust("CustStatusCode").Value)
                                       lngStatusCount = lngStatusCount + 1
                    '���`�Ȥ�BCM������
                    ElseIf rsCust("CustStatusCode").Value & "" = 1 And rsCust("WipCode3") = 16 Then
                                       ErrFile.WriteLine ("�Ȥ�s��: " & rsChk("CustId").Value & " ��ڽs���G " & vBillNo & " ��CM�������Ȥ� ")
                                       lngStatusCount = lngStatusCount + 1
                    End If
           
            '�����sSO033
                    If UpdAccount( _
                                            vBillNo, vRealAmt, vRealDate, vUpdEn, vCMCode, vCMName, _
                                            vPTCode, vPTName, pEmpNO, pEmpName, pAuthorizeNo _
                                            ) = False Then
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
    
    ChkData = True
  Exit Function
ChkErr:
  ErrSub "DealFun", "ChkData"
End Function

Public Function UpdAccount(ByVal vBillNo As String, ByVal vRealAmt As String, _
                           ByVal vRealDate As String, ByVal vUpdEn As String, _
                           ByVal vCMCode As String, ByVal vCMName As String, _
                           ByVal vPTCode As String, ByVal vPTName As String, _
                           Optional pEmpNO As String = "", Optional pEmpName As String = "", _
                           Optional pAuthorizeNo As String) As Boolean

 '' vBillNo ��ڽs��
 '' vClctEn ���O�H��
 '' vRealDate �ꦬ���
 '' vCMCode ���O�覡�N�X  ref CD031''
 '' vPTCode �I�ں���  ref CD032''
On Error GoTo ChkErr

 Dim rsAccount As New ADODB.Recordset
 Dim strsql As String
 Dim strUpdateEmpSql As String
 Dim strBillField As String
 Dim intBudgetPeriod As Integer
 Dim strBudgetPeriodSQL As String
 ''  2003/08/18  �o���ܼƩw�q�����ثe�վ㪺�o�@��SO033  �A�N��ǵ��Q�J��FUNCTION AlterSO003
 Dim rsAlterSO003 As New ADODB.Recordset
 
 If blnMultiAccount = True Then
      strBillField = "MediaBillno"
 Else
     strBillField = "Billno"
 End If
    '' �o�@�檺WHERE  �᭱�� Billno   �令MediaBillno �H�A�X�h�b���\��
    strsql = "Select RowId, CustId, ShouldAmt, CitemCode, RealStartDate, RealStopDate, RealPeriod, ClctEn, ClctName, CMCode, CMName,BudgetPeriod,AccountNO From  " & frmUnion.TableOwnerName & "So033    Where UCCODE IS NOT NULL  And " & strBillField & " ='" & vBillNo & "' AND CancelFlag <> 1"
        
    If Not GetRS(rsAccount, strsql, gcnGi) Then
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
       
         strsql = "Update " & frmUnion.TableOwnerName & "So033 Set RealAmt =" & rsAccount("ShouldAmt") & _
                                 ",RealDate = To_date('" & vRealDate & "','YYYYMMDD')" & _
                                  ",Updtime = '" & GetDTString(strNowTime) & "'" & _
                                  ",UpdEn = '" & vUpdEn & "'" & _
                                  ",UCCode = Null, UCName = Null" & _
                                  ", CMCode = '" & vCMCode & "'" & _
                                  ", CMName = '" & vCMName & "'" & _
                                  ", PTCode = '" & vPTCode & "'" & _
                                  ", PTName = '" & vPTName & "'" & _
                                  strUpdateEmpSql & _
                                  ",FirstTime = '" & GetDTString(strNowTime) & "'," & _
                                  "AuthorizeNo='" & pAuthorizeNo & "' " & _
                                  "Where Rowid = '" & rsAccount("Rowid") & "'"
          gcnGi.Execute strsql
           ''  �H�o�@�q �е����Ojacky���{���X  , �Ѧҳo�@�q�ӽվ㪺�{���X�b�U��
           ''  �� �Ѿl���Ƥj��0���ɭ�  ���ݱN SO033 ���Ѿl���Ʀ����@�A�A��J
           '' SO003
           
'            If rsSO033("BudgetPeriod") > 0 Then
'                    .Fields("BudgetPeriod") = rsSO033("BudgetPeriod") - 1
'                    If .Fields("BudgetPeriod") = 0 Then
'                    If Not ExecuteCommand("Delete From " & GetOwner & "SO003 Where RowId= '" & .Fields("RowId") & "'") Then Exit Function
'                            AlterSO003 = True
'                            Exit Function
'                    End If
'            End If
'
'
''' 2003/08/18  �H�U�o�@�q �аO�� �{���X����O�����{�b�令�H �Q�J��FUNCTION �ӧ@
'' **************************************************************************************************
'                                                        strBudgetPeriodSQL = "    "
'                                                        If rsAccount("BudgetPeriod") > 0 Then
'                                                                 intBudgetPeriod = rsAccount("BudgetPeriod") - 1
'                                                                 If intBudgetPeriod = 0 Then
'                                                                             gcnGi.Execute ("Delete  " & frmUnion.TableOwnerName & "SO003 Where CustId = '" & rsAccount("CustId") & "' And CitemCode = '" & rsAccount("CitemCode") & "'")
'                                                                 Else
'                                                                             strBudgetPeriodSQL = ",BudgetPeriod = " & intBudgetPeriod & " "
'
'                                                                 End If
'
'                                                         End If
'
'                                                         strsql = "Update " & frmUnion.TableOwnerName & "So003 Set StartDate = " & GetNullString(rsAccount("RealStartDate"), giDateV, giOracle) & _
'                                                                                        ",StopDate = " & GetNullString(rsAccount("RealStopDate"), giDateV, giOracle) & " + " & intPara3 & _
'                                                                                        ",ClctDate = " & GetNullString(rsAccount("RealStopDate"), giDateV, giOracle) & " + " & intPara5 & _
'                                                                                        ",UpdEn = '" & vUpdEn & "',UpdTime = '" & GetDTString(strNowTime) & "'" & _
'                                                                                        strBudgetPeriodSQL & _
'                                                                                        "Where CustId = '" & rsAccount("CustId") & "' And CitemCode = '" & rsAccount("CitemCode") & "' AND AccountNO = '" & rsAccount("AccountNO") & "'"
'
'                                                         gcnGi.Execute strsql
  '' **************************************************************************************************
          
'            Call GetRS(rsAlterSO003, _
'                             "SELECT * FROM SO003 WHERE RowID ='" & rsAccount("Rowid") & "'", gcnGi)
'
'           Call AlterSO003(rsAlterSO003)  ''�վ�SO003
                             
            rsAccount.MoveNext
           
    Wend
    lngCount = lngCount + 1
    Set rsAccount = Nothing
    UpdAccount = True
 Exit Function
ChkErr:
   UpdAccount = False
   ErrSub "UpdateFun", "UpdAccount"
End Function


'�ھڶǨӤ��t�Τ��w�ɶ��r���ন�T�w���ɶ��r��
'Parameter1: DateTime
'Return: Format DateTime String
Public Function GetDTString(dtm As String) As String
  On Error GoTo ChkErr
    GetDTString = Format(dtm, "ee/mm/dd hh:mm:ss")
  Exit Function
ChkErr:
   ErrSub "DealFun", "GetDTString"
End Function

Public Function GetNullString(varValue As Variant, Optional lngVType As giVType = giStringV, Optional DBFlag As ChangeDB = giOracle) As Variant
    On Error GoTo ChkErr
        If lngVType = giDateV Then
            If Len(varValue & "") = 0 Then
                GetNullString = "NULL"
            Else
                If DBFlag = giOracle Then
                    GetNullString = "To_Date('" & Format(varValue, "YYYYMMDD") & "','YYYYMMDD')"
                Else
                    GetNullString = "#" & Format(varValue, "YYYY/MM/DD") & "#"
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

Public Function InitData() As Boolean
On Error GoTo ChkErr
   Dim rsTmp As New ADODB.Recordset
   Dim strsql As String
   
   strsql = "Select Para3, Para5, Para8 From " & frmUnion.TableOwnerName & "So043"
   If Not GetRS(rsTmp, strsql, gcnGi) Then InitData = False: Exit Function
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
                  Optional strWhere As String)
  On Error GoTo ChkErr
    objLst.SetFldName1 strFldName1
    objLst.SetFldName2 strFldName2
    If lngFldLen1 > 0 Then objLst.FldLen1 = lngFldLen1
    If lngFldLen2 > 0 Then objLst.FldLen2 = lngFldLen2
    If lngWidth1 > 0 Then objLst.FldWidth1 = lngWidth1
    If lngWidth2 > 0 Then objLst.FldWidth2 = lngWidth2
    objLst.SetTableName StrTableName
    objLst.SendConn gcnGi
    objLst.Filter = strWhere
  Exit Sub
ChkErr:
    Call ErrSub("DealFun", "SetLst")
End Sub

Public Function GetRS(ByRef rs As ADODB.Recordset, _
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
    If rs.State = adStateOpen Then rs.Close
    rs.CursorLocation = CursorLocation
    If Val(MaxRecords) > 0 Then rs.MaxRecords = Val(MaxRecords)
    If Conn Is Nothing Then Set Conn = gcnGi
    rs.Open strsql, Conn, CursorType, LockType
    GetRS = True
  Exit Function
ChkErr:
    GetRS = False
    MsgBox strsql
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

Public Function AlterSO003(rsSO033 As ADODB.Recordset, Optional ByVal blnAddSo003 As Boolean = True, _
                Optional ByVal blnUpdAmt As Boolean = True, Optional ByVal blnUpdClctDate As Boolean = True, _
                Optional ByVal blnChkMustAnswer As Boolean = False, Optional ByRef strSeqNo As String, _
                Optional ByVal blnMultiCount As Boolean = False) As Boolean
    On Error GoTo ChkErr
    
    Dim rsAlter As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim lngPeriodFlag As Boolean
    Dim Ypara5 As Long, Ypara8 As Long
    Dim Ypara9 As Long, YPara24 As Long
    Dim lngCompCode As Long, lngBad As Long
    Dim blnFlag As Boolean
    
        '�P�_�O�_���g���ʦ��O����
        
        lngPeriodFlag = Val(GetRsValue("Select PeriodFlag From " & GetOwner & "Cd019 Where CodeNo = " & rsSO033("CitemCode")) & "")
        If Len(rsSO033("STCode") & "") > 0 Then lngBad = Val(GetRsValue("Select RefNo From " & GetOwner & "CD016 Where CodeNo = " & rsSO033("STCode")) & "")
        If lngPeriodFlag = 0 Or (lngBad = 1) Then
            AlterSO003 = True
            Exit Function
        End If
        
        If Not GetSystemPara(rsTmp, rsSO033("CompCode") & "", rsSO033("ServiceType"), "SO043", "Para5,Para8,Para9,Para24") Then Exit Function
        If rsTmp.RecordCount = 0 Then Exit Function
        Ypara5 = rsTmp("Para5")
        Ypara8 = rsTmp("Para8")
        Ypara9 = rsTmp("Para9")
        YPara24 = rsTmp("Para24")
        Call CloseRecordset(rsTmp)
        
        lngCompCode = Val(GetRsValue("Select CompCode From " & GetOwner & "So001 Where CustId = " & rsSO033("CustId")) & "")
        
'        If rsSO033("OrderNo") & "" <> "" Then
            If Val(rsSO033("SeqNo") & "") = 0 Then
                If Not GetRS(rsAlter, "Select RowId,SO003.* From " & GetOwner & "So003 Where 1 = 0 ", gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
            Else
                If Not GetRS(rsAlter, "Select RowId,SO003.* From " & GetOwner & "So003 Where CustId = " & rsSO033("CustId") & " And CitemCode = " & rsSO033("CitemCode") & " And SeqNo = " & rsSO033("SeqNo"), gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
            End If
'        End If
        With rsAlter
            If .RecordCount = 0 Then
                '�ˬd�O���n�s�W�g���ʦ��O����
                If blnAddSo003 Then
                    If blnChkMustAnswer Then
                        If MsgBox("�L�g���ʦ��O����,�O�_�n�s�W??", vbQuestion + vbYesNo + vbDefaultButton2, gimsgPrompt) = vbNo Then
                            AlterSO003 = True
                            Exit Function
                        End If
                    End If
                    .AddNew
                    If Val(rsSO033("SeqNo") & "") = 0 Then
                        .Fields("SeqNO") = GetInvoiceNo3("SO003") & ""
                    Else
                        .Fields("SeqNO") = rsSO033("SeqNo")
                    End If
                    blnFlag = True
                    '�����~�۰����Ȥ��s��A�� 92/03/25 �P Lawrence �Q�פ����G 92/08/12 �S��
                    If YPara24 = 1 Then
                        If rsSO033("AccountNO") & "" <> "" Then
                            If Not GetRS(rsTmp, "Select Nvl(AddCitemAccount,0) AddCitemAccount,SnactionDate From " & GetOwner & "SO106 Where CustId = " & rsSO033("CustId") & _
                                " And AccountID = '" & rsSO033("AccountNo") & "' And Nvl(StopFlag,0)= 0") Then Exit Function
                            If rsTmp.RecordCount > 0 Then If rsTmp("AddCitemAccount") = 0 Or rsTmp("SnactionDate") & "" = "" Then blnFlag = False
                        End If
                    End If
                    If blnFlag Then
                        .Fields("CMCode") = NoZero(rsSO033("CMCode"))
                        .Fields("CMName") = NoZero(rsSO033("CMName"))
                        .Fields("BankCode") = NoZero(rsSO033("BankCode"))
                        .Fields("BankName") = NoZero(rsSO033("BankName"))
                        .Fields("AccountNo") = NoZero(rsSO033("AccountNo"))
                    Else
                        If Not GetRS(rsTmp, "Select CodeNo,Description From " & GetOwner & "CD031 Where RefNo=1", gcnGi) Then Exit Function
                        .Fields("CMCode") = NoZero(rsTmp("CodeNo"))
                        .Fields("CMName") = NoZero(rsTmp("Description"))
                    End If
                    
                    .Fields("OrderNo") = NoZero(rsSO033("OrderNo"))
                    .Fields("PTCode") = NoZero(rsSO033("PTCode"))
                    .Fields("PTName") = NoZero(rsSO033("PTName"))
                Else
                    AlterSO003 = True
                    Exit Function
                End If
            Else
                blnAddSo003 = False
            End If

            If .Fields("StopDate") >= rsSO033("RealStopDate") Then AlterSO003 = True: Exit Function
            If blnAddSo003 Or (blnUpdAmt And Ypara8 >= 1) Then
                '�Y<Para9>=1, �h 1: Amount = OldAmt, 2: Amount = ShouldAmt , 3: Amount = RealAmt
                .Fields("Period").Value = GetFieldValue(rsSO033, "RealPeriod")
                If Ypara9 = 1 Then
                    .Fields("Period").Value = GetFieldValue(rsSO033, "OldPeriod")
                    .Fields("Amount").Value = GetFieldValue(rsSO033, "OldAmt")
                ElseIf Ypara9 = 2 Then
                    '.Fields("Period").Value = GetFieldValue(rsSo033, "RealPeriod")
                    .Fields("Amount").Value = GetFieldValue(rsSO033, "ShouldAmt")
                Else
                    '.Fields("Period").Value = GetFieldValue(rsSo033, "RealPeriod")
                    .Fields("Amount").Value = GetFieldValue(rsSO033, "RealAmt")
                End If
            End If
            If rsSO033("BudgetPeriod") > 0 Then
                .Fields("BudgetPeriod") = rsSO033("BudgetPeriod") - rsSO033("RealPeriod") / .Fields("Period")
                If .Fields("BudgetPeriod") = 0 Then
                    If Not ExecuteCommand("Delete From " & GetOwner & "SO003 Where RowId= '" & .Fields("RowId") & "'") Then Exit Function
                    AlterSO003 = True
                    Exit Function
                End If
            End If
            .Fields("CustId").Value = GetFieldValue(rsSO033, "CustID")
            .Fields("CitemCode").Value = GetFieldValue(rsSO033, "CitemCode")
            .Fields("CitemName").Value = GetFieldValue(rsSO033, "CitemName")
            .Fields("StartDate").Value = GetFieldValue(rsSO033, "RealStartDate")
            .Fields("StopDate").Value = GetFieldValue(rsSO033, "RealStopDate")
            .Fields("UpdTime").Value = GetFieldValue(rsSO033, "UpdTime")
            .Fields("UpdEn").Value = GetFieldValue(rsSO033, "UpdEn")
            .Fields("ClctDate").Value = GetFieldValue(rsSO033, "RealStopDate") + Ypara5
            .Fields("LastDate").Value = GetFieldValue(rsSO033, "RealStopDate") + Ypara5
            .Fields("CompCode").Value = lngCompCode
            .Fields("ServiceType").Value = GetFieldValue(rsSO033, "ServiceType")
            .Fields("PackageNo").Value = NoZero(rsSO033("PackageNo"))
            .Fields("PackageName").Value = NoZero(rsSO033("PackageName"))
            If rsSO033("DiscountCode") & "" <> "" Then
                .Fields("Amount") = NoZero(rsSO033("OldAmt"), True)
                .Fields("DiscountCode") = NoZero(rsSO033("DiscountCode"))
                .Fields("DiscountName") = NoZero(rsSO033("DiscountName"))
                .Fields("DiscountDate1") = NoZero(rsSO033("DiscountDate1"))
                .Fields("DiscountDate2") = NoZero(rsSO033("DiscountDate2"))
                .Fields("DiscountAmt") = NoZero(rsSO033("DiscountAmt"))
            End If
            .Update
            strSeqNo = .Fields("SeqNo") & ""
        End With
        AlterSO003 = True
    Exit Function
ChkErr:
    ErrSub FormName, "AlterSo003"
End Function

Public Function GetSystemParaItem(strField As String, _
    strCompCode As String, strServiceType As String, strTable As String, _
    Optional blnShowMsg As Boolean = False, Optional blnServiceTypeUseLike As Boolean = False) As Variant
    On Error Resume Next
    Dim rs As New ADODB.Recordset
        If Not GetSystemPara(rs, strCompCode, strServiceType, strTable, strField, False, blnServiceTypeUseLike) Then Exit Function
        GetSystemParaItem = rs(0)
End Function


