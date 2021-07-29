VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mediaEsun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
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
Private rsCD013 As New ADODB.Recordset
Private strFailUpdSQL As String
Private strUPDTime As String
Public Sub Action()
  On Error GoTo ChkErr
    
    blnFailUpd = False
    
    '*************************************************************************************************************************************
    '#3417 Ū��������]��5,�p�G�����,��s���Ѯɭn�^�gUCCode�PUCName By Kin 2007/12/07
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
    
    lngTime = Timer
    gcnGi.BeginTrans
    While Not File.AtEndOfLine   'Ū����b���
        strData = File.ReadLine   'Ū���@�C���
        
        '******************************************************************************************
        '#3417 ���Ѯɭn��s���y�k
        '#4388 �W�[��s���ʤH���P���ʮɶ� By Kin 2009/05/04
        strFailUpdSQL = Empty
        If blnFailUpd Then
            strFailUpdSQL = "Update " & strOwnerName & "SO033 Set UCCode=" & strUCCode & _
                            ",UPDTime='" & strUPDTime & "',UPDEN='" & strUpdEn & "'" & _
                            ",UCName='" & strUCName & "',Note=Note"
                                                                                      
        End If
        '******************************************************************************************

        If strMedia = 1 Then
            strBillNo = MidMbcs(strData, 92, 11)
        Else
            strBillNo = GetBillNo_New(MidMbcs(strData, 92, 11))
        End If
        
        '********************************************************************************************
        '#3417 �n��sUCCode�PUCName�� Where ���� By Kin 2007/12/11
        strFailUpdWhere = Empty
        strFailUpdWhere = IIf(strMedia = "1", " MediaBillNo='", " BillNo='") & strBillNo & "'" & _
                        " And UCCode is Not Null" & _
                        " And CancelFlag<>1 " & _
                         IIf(strCitemQry <> "", " And CitemCode In (" & strCitemQry & ")", "")
        '*********************************************************************************************
    
     '����b���G(B67~70),�Y���b�����\�A�h�g�J���~log�ɡA���~���ƥ[��
     '#4388 ���դ�OK UPD�ɭn�N���ѭ�]��b�e��,��ڽs���b�᭱ By Kin 2009/05/13
        Select Case MidMbcs(strData, 67, 4)
            Case "1014"
                ErrFile.WriteLine ("��ڽs���G " & strBillNo & " �l�B����!")
                
                '#3417 �p�G���ڤ����\Update������]��SO033 By Kin 2007/12/10
                If blnFailUpd Then
                    '#3913 �u���Ĥ@���|���\�A�]���쥻�g�k��strFailUpdSQL=strFailUpdSQL & "..."�A�o�˲ĤG����|�s�Ĥ@������i�h�A�h�[�@���ܼ��x�s
                    strFail = Empty
                    strFail = strFailUpdSQL & "||'���ѥN�X:1014;���ѭ�]:�l�B����!;��ڽs��: " & _
                                    strBillNo & " ; ������]:" & _
                                    strUCCode & "; ������]�W��:" & strUCName & "'" & _
                                    " Where " & strFailUpdWhere

                    gcnGi.Execute strFail
                End If

                lngErrCount = lngErrCount + 1
            Case "0901"
                ErrFile.WriteLine ("��ڽs���G " & strBillNo & " ��Ʀ��~!")
                
                '#3417 �p�G���ڤ����\Update������]��SO033 By Kin 2007/12/10
                If blnFailUpd Then
                    '#3913 �u���Ĥ@���|���\�A�]���쥻�g�k��strFailUpdSQL=strFailUpdSQL & "..."�A�o�˲ĤG����|�s�Ĥ@������i�h�A�h�[�@���ܼ��x�s
                    strFail = Empty
                    strFail = strFailUpdSQL & "||'���ѥN�X:0901;���ѭ�]:��Ʀ��~!;��ڽs��: " & _
                                    strBillNo & " ; ������]:" & _
                                    strUCCode & "; ������]�W��:" & strUCName & "'" & _
                                    " Where " & strFailUpdWhere

                    gcnGi.Execute strFail
                End If

                lngErrCount = lngErrCount + 1
            Case "0922"
                ErrFile.WriteLine ("��ڽs���G " & strBillNo & " ���b��w�P��(���M)!")
                
                '#3417 �p�G���ڤ����\Update������]��SO033 By Kin 2007/12/10
                If blnFailUpd Then
                    '#3913 �u���Ĥ@���|���\�A�]���쥻�g�k��strFailUpdSQL=strFailUpdSQL & "..."�A�o�˲ĤG����|�s�Ĥ@������i�h�A�h�[�@���ܼ��x�s
                    strFail = Empty
                    strFail = strFailUpdSQL & "||'���ѥN�X:0922;���ѭ�]:���b��w�P��(���M)!;��ڽs��: " & _
                                    strBillNo & " ; ������]:" & _
                                    strUCCode & "; ������]�W��:" & strUCName & "'" & _
                                    " Where " & strFailUpdWhere

                    gcnGi.Execute strFail
                End If

                lngErrCount = lngErrCount + 1
            Case "1013"
                ErrFile.WriteLine ("��ڽs���G " & strBillNo & " �s�ᤧ�����w������!")
                
                '#3417 �p�G���ڤ����\Update������]��SO033 By Kin 2007/12/10
                If blnFailUpd Then
                    '#3913 �u���Ĥ@���|���\�A�]���쥻�g�k��strFailUpdSQL=strFailUpdSQL & "..."�A�o�˲ĤG����|�s�Ĥ@������i�h�A�h�[�@���ܼ��x�s
                    strFail = Empty
                    strFail = strFailUpdSQL & "||'���ѥN�X:1013;���ѭ�]:�s�ᤧ�����w������!��ڽs��: " & _
                                    strBillNo & " ; ������]:" & _
                                    strUCCode & "; ������]�W��:" & strUCName & "'" & _
                                    " Where " & strFailUpdWhere

                    gcnGi.Execute strFail
                End If

                lngErrCount = lngErrCount + 1
            Case "0967"
                ErrFile.WriteLine ("��ڽs���G " & strBillNo & " �����Ҹ��X���~!")
                
                '#3417 �p�G���ڤ����\Update������]��SO033 By Kin 2007/12/10
                If blnFailUpd Then
                    '#3913 �u���Ĥ@���|���\�A�]���쥻�g�k��strFailUpdSQL=strFailUpdSQL & "..."�A�o�˲ĤG����|�s�Ĥ@������i�h�A�h�[�@���ܼ��x�s
                    strFail = Empty
                    strFail = strFailUpdSQL & "||'���ѥN�X:0967;���ѭ�]:�����Ҹ��X���~!;��ڽs��: " & _
                                    strBillNo & " ; ������]:" & _
                                    strUCCode & "; ������]�W��:" & strUCName & "'" & _
                                    " Where " & strFailUpdWhere

                    gcnGi.Execute strFail
                End If

                lngErrCount = lngErrCount + 1
            Case "    "   '���J�b���\
              '���ӵ��J/���b���B(B14~B28���X���p��)
                strRealAmt = MidMbcs(strData, 14, 13)
              
                If ChkData(strBillNo, strRealAmt, strClctEn, strClctName, strRealDate, strUpdEn, strCMCode, strCMName, strPTCode, strPTName, strBankName, IIf(strMedia = 1, False, True), strCitemQry, strMedia, , strBankId) = False Then
                    IntoAccount = False
                    GoTo Rollback
                End If
        
        End Select
           
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