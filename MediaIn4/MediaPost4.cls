VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mediaPost4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'******************************************************************************************'
'����#2825�s�W���@�\��mediaPost3 By Kin 2006/11/30 Release
'******************************************************************************************'
Private blnUpdate As Boolean        '��s�{���O�_���\
Private TmpFilePath As String
Private ErrFilePath As String
Private strFilePath As String       'INI�ɮ׸��|
Private strRealDate As String       '�J�b���
Private strSourcePath As String     '��l�ɮ׸��|
Private strUpdEn As String          '�O�����ʤH��
Private strPrgName As String        '��b�{���W��
Private strCMCode As String         '���O�覡�N�X
Private strCMName As String         '���O�覡�W��
Private strClctEn As String         '���O�H���N�X
Private strClctName As String       '���O�H���W��
Private strPTCode As String         '�I�ں���
Private strPTName As String         '�I�ں����W��
Private strServiceType As String    '�A�����O
Private strBankName As String       '�Ȧ�W��
Private strBankId As String         '�Ȧ�N�X
Private strCompCode As String       '���q�O
Private strOption As String         '�P�_���eoption�����A
Private strCitemQry As String
Private rsTmp As New ADODB.Recordset
Private intRefNo As Integer
Private rsCD013 As New ADODB.Recordset
Private strFailUpdSQL As String
Private strUPDTime As String

Public Sub Action()
  On Error GoTo ChkErr
    blnFailUpd = False
    '*************************************************************************************************************************************
    '#3417 Ū��������]��5,�p�G�����,��s���Ѯɭn�^�gUCCode�PUCName By Kin 2007/12/11
    If Not GetRS(rsCD013, "Select CodeNo,Description From " & strOwnerName & "CD013 Where Nvl(RefNo,0)=5 And StopFlag<>1 Order by CodeNo Desc", _
                  gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
    If rsCD013.EOF Then
        blnFailUpd = False
    Else
        strUCCode = rsCD013("CodeNo") & ""
        strUCName = rsCD013("Description") & ""
        blnFailUpd = True
    End If
    '*************************************************************************************************************************************

    Call IntoData
    If Not IntoAccount Then
        objStorePara.UpdState = False
        Exit Sub
    Else
        objStorePara.UpdState = False
    End If
     
    Call MsgResult       '��ܵ��G�T��
    '�Y���~���Ƥj��0�h��NotePad������ɮפ��e
    If lngErrCount > 0 Then Call Shell("NotePad " & ErrFilePath, vbNormalFocus)
    File.Close
    ErrFile.Close
    Set FSO = Nothing
    CloseRecordset rsCD013
   Exit Sub
ChkErr:
    CloseRecordset rsCD013
    ErrSub "mediaCShing", "Action"
End Sub

Private Sub IntoData()
On Error GoTo ChkErr

    With objStorePara
        strFilePath = .FilePath
        strRealDate = .RealDate
        strSourcePath = .SourcePath
        strUpdEn = .UpdEn
        strPrgName = .PrgName
        strClctEn = .uClctEn
        strClctName = .uClctName
        strCMCode = .uCMCode
        strCMName = .uCMName
        strServiceType = .ServiceType
        strCompCode = .CompCode
        strBankName = .BankName
        strBankId = .BankId
        strCitemQry = .uCitemQry
        strUPDTime = .uUPDTime
        '#3527 �P�_�ѦҸ��A�p�G�O2�HCityBankAtm������ڨ̾� By Kin 2007/10/04
        intRefNo = .uRefNo
        '#7033 �W�[�I�ں��� By Kin 2015/04/28
        strPTCode = .uPTCode
        strPTName = .uPTName
    End With
   
    If InitData = False Then
        File.Close
        ErrFile.Close
        Set FSO = Nothing
        Exit Sub
    End If
  Exit Sub
ChkErr:
  ErrSub "mediaCShing", "IntoData"
End Sub

Private Function IntoAccount() As Boolean
  On Error GoTo ChkErr
    Dim strBillNo As String
    Dim strRealAmt As String
    Dim strData As String
    Dim strYYMM As String
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim rsPara As New ADODB.Recordset
    Dim strMedia As String
    Dim strFailUpdWhere As String
    Dim strErrCode As String
    Dim strErrDescript As String
    Dim strFail As String
    lngCount = 0
    lngErrCount = 0
    TmpFilePath = strFilePath & "\Tmp" & strPrgName & "In.log"
    ErrFilePath = strFilePath & "\" & strPrgName & "Err.log"
    Call GetFilePath(strFilePath)
    FSO.CopyFile strSourcePath, TmpFilePath
    Set File = FSO.OpenTextFile(TmpFilePath, ForReading)
    Set ErrFile = FSO.CreateTextFile(ErrFilePath)
    
    strNowTime = RightNow
    strSQL = "Select Para24 From " & strOwnerName & "SO043 Where Rownum=1"
    If Not GetRS(rsPara, strSQL, gcnGi) Then Exit Function
    strMedia = GetFieldValue(rsPara, "Para24") & ""
    
    '******************************************************************************************
    '#3417 ���Ѯɭn��s���y�k By Kin 2007/12/11
    '#4388 �W�[��s���ʤH���P���ʮɶ� By Kin 2009/05/04
    strFailUpdSQL = Empty
    If blnFailUpd Then
        strFailUpdSQL = "Update " & strOwnerName & "SO033 Set UCCode=" & strUCCode & _
                        ",UPDTime='" & strUPDTime & "',UPDEN='" & strUpdEn & "'" & _
                        ",UCName='" & strUCName & "',Note=Note"
                                                                                  
    End If
    '******************************************************************************************

    lngTime = Timer
    gcnGi.BeginTrans
    While Not File.AtEndOfLine   'Ū����b���
        strData = File.ReadLine   'Ū���@�C���
        '���b���D��7535
        '*************************************************************************************************
        If Mid(strData, 1, 1) = 1 Then
            '*********************************************************************************************
            strBillNo = Trim(MidMbcs(strData, 91, 110))
            strFailUpdWhere = " CitiBankATM='" & strBillNo & "'" & _
                               " And UCCode is Not Null" & _
                               " And CancelFlag<>1 " & _
                               IIf(strCitemQry <> "", " And CitemCode In (" & strCitemQry & ")", "")
            
            '*********************************************************************************************
         
            
            strErrCode = Empty
            strErrDescript = Empty
            '#4388 ���դ�OK UPD��,��]�n�b�e�� ��ڽs���b�᭱ By Kin 2009/05/12
            Select Case MidMbcs(strData, 79, 2)
                Case "01"
                    ErrFile.WriteLine ("��ڽs���G " & strBillNo & " ���s����!")
                    strErrCode = "01"
                    strErrDescript = " ���s����! ;��ڽs���G " & strBillNo
                    lngErrCount = lngErrCount + 1
                Case "02"
                    ErrFile.WriteLine ("��ڽs���G " & strBillNo & " ���ӽХNú�ΥΤḹ�X���~!")
                    strErrCode = "02"
                    strErrDescript = " ���ӽХNú�ΥΤḹ�X���~! ;��ڽs���G " & strBillNo
                    lngErrCount = lngErrCount + 1
                Case "03"
                    ErrFile.WriteLine ("��ڽs���G " & strBillNo & " �w�פ�Nú!")
                    strErrCode = "03"
                    strErrDescript = " �w�פ�Nú! ;��ڽs���G " & strBillNo
                    lngErrCount = lngErrCount + 1
                Case "04"
                    ErrFile.WriteLine ("��ڽs���G " & strBillNo & " �ڵ����Ӥ�!")
                    strErrCode = "04"
                    strErrDescript = " �ڵ����Ӥ�! ;��ڽs���G " & strBillNo
                    lngErrCount = lngErrCount + 1
                Case "06"
                    ErrFile.WriteLine ("��ڽs���G " & strBillNo & " �ᵲ���ĵ�ܤ�!")
                    strErrCode = "06"
                    strErrDescript = " �ᵲ���ĵ�ܤ�! ;��ڽs���G " & strBillNo
                    lngErrCount = lngErrCount + 1
                Case "07"
                    ErrFile.WriteLine ("��ڽs���G " & strBillNo & " �~�Ȥ䲼�M��!")
                    strErrCode = "07"
                    strErrDescript = " �~�Ȥ䲼�M��! ;��ڽs���G " & strBillNo
                    lngErrCount = lngErrCount + 1
                Case "08"
                    ErrFile.WriteLine ("��ڽs���G " & strBillNo & " �b�����~!")
                    strErrCode = "08"
                    strErrDescript = " �b�����~! ;��ڽs���G " & strBillNo
                    lngErrCount = lngErrCount + 1
                Case "09"
                    ErrFile.WriteLine ("��ڽs���G " & strBillNo & " �b��w�פ�!")
                    strErrCode = "09"
                    strErrDescript = " �b��w�פ�! ;��ڽs���G " & strBillNo
                    lngErrCount = lngErrCount + 1
                Case "10"
                    ErrFile.WriteLine ("��ڽs���G " & strBillNo & " �����Ҹ�����!")
                    strErrCode = "10"
                    strErrDescript = " �����Ҹ�����! ;��ڽs���G " & strBillNo
                    lngErrCount = lngErrCount + 1
                Case "15"
                    ErrFile.WriteLine ("��ڽs���G " & strBillNo & " �W�L���ڤW��!")
                    strErrCode = "15"
                    strErrDescript = " �W�L���ڤW��! ;��ڽs���G " & strBillNo
                    lngErrCount = lngErrCount + 1
                Case "18"
                    ErrFile.WriteLine ("��ڽs���G " & strBillNo & " �˸����~!")
                    strErrCode = "18"
                    strErrDescript = " �˸����~! ;��ڽs���G " & strBillNo
                    lngErrCount = lngErrCount + 1
                Case "19"
                    ErrFile.WriteLine ("��ڽs���G " & strBillNo & " ��L!")
                    strErrCode = "19"
                    strErrDescript = " ��L! ;��ڽs���G " & strBillNo
                    lngErrCount = lngErrCount + 1
                Case Else   '���J�b���\ ���D��2869 �h�W�[�P�_�O���O�٦��䥦���ѥN�X 2008/11/16
                    If MidMbcs(strData, 79, 2) = "  " Then
                        '���ӵ��J/���b���B(B44~B52���X���p��)
                        strRealAmt = MidMbcs(strData, 44, 9)
                        '�Y�J�b������ťաA�h��(B9~B14)��������
                        If Len(strRealDate) = 0 Then
                            '������������~
                            strRealDate = Val(MidMbcs(strData, 10, 7)) + 19110000
                        End If
                    '�YPara24=1,�h�δC��渹�J�b
                        If ChkData(strBillNo, strRealAmt, strClctEn, strClctName, strRealDate, strUpdEn, strCMCode, strCMName, strPTCode, strPTName, strBankName, False, strCitemQry, "1", 2, strBankId) = False Then
                           IntoAccount = False
                           GoTo Rollback
                        End If
                    Else
                        ErrFile.WriteLine ("��ڽs���G " & strBillNo & " ����!   �N�X�G" & MidMbcs(strData, 79, 2))
                        strErrCode = MidMbcs(strData, 79, 2)
                        strErrDescript = " ����! ;��ڽs���G " & strBillNo
                        lngErrCount = lngErrCount + 1
                    End If
            End Select
            
            '******************************************************************************************************************
            '#3417 ���p�����~�A��sUCCode�PUCName By Kin 2007/12/11
            If strErrDescript <> Empty Then
                If blnFailUpd Then
                    '#3913 �u���Ĥ@���|���\�A�]���쥻�g�k��strFailUpdSQL=strFailUpdSQL & "..."�A�o�˲ĤG����|�s�Ĥ@������i�h�A�h�[�@���ܼ��x�s
                    strFail = Empty
                    strFail = strFailUpdSQL & "||'���ѥN�X:" & strErrCode & ";���ѭ�]:" & strErrDescript & _
                                    "; ������]:" & strUCCode & _
                                    "; ������]�W��:" & strUCName & "'" & _
                                    " Where " & strFailUpdWhere
    
                    gcnGi.Execute strFail
                End If
            End If
            '******************************************************************************************************************
            
        End If
Nextloop:
        Wend
        gcnGi.CommitTrans
        IntoAccount = True
        On Error Resume Next
        rsTmp.Close
        Set rsTmp = Nothing
        rsPara.Close
        Set rsPara = Nothing
  Exit Function
Rollback:
        gcnGi.RollbackTrans
        MsgBox "��s����!", vbExclamation, "ĵ�i..."
        File.Close
        ErrFile.Close
        Set FSO = Nothing
  Exit Function
ChkErr:
  ErrSub "mediaCShing", "IntoAccount"
End Function



