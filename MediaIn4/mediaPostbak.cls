VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mediaPost3"
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
Private strCompCode As String       '���q�O
Private strOption As String         '�P�_���eoption�����A
Private strCitemQry As String
Private rsTmp As New ADODB.Recordset

Public Sub Action()
  On Error GoTo ChkErr
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
      Set fso = Nothing
      
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
     strCitemQry = .uCitemQry
   End With
   
   If InitData = False Then
      File.Close
      ErrFile.Close
      Set fso = Nothing
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

        lngCount = 0
        lngErrCount = 0
        TmpFilePath = strFilePath & "\Tmp" & strPrgName & "In.log"
        ErrFilePath = strFilePath & "\" & strPrgName & "Err.log"
        Call GetFilePath(strFilePath)
        fso.CopyFile strSourcePath, TmpFilePath
        Set File = fso.OpenTextFile(TmpFilePath, ForReading)
        Set ErrFile = fso.CreateTextFile(ErrFilePath)
        
        strNowTime = RightNow
        strSQL = "Select Para24 From " & strOwnerName & "SO043 Where Rownum=1"
        If Not GetRS(rsPara, strSQL, gcnGi) Then Exit Function
        strMedia = GetFieldValue(rsPara, "Para24") & ""
        
        lngTime = Timer
        gcnGi.BeginTrans
        While Not File.AtEndOfLine   'Ū����b���
            strData = File.ReadLine   'Ū���@�C���
'            strCMCode = ""
'            strCMName = ""
            'para24=1,���C��渹(B44~B54)�ӤJ�b
            If strMedia = 1 Then
               strBillNo = MidMbcs(strData, 44, 11)
            'para24=0,����ڽs��(B44~B54),�å�GetBillNo_New()�ഫ��15�X��s
            Else
               strBillNo = GetBillNo_Old(MidMbcs(strData, 44, 11), strServiceType)
            End If
            
            '���ӵ��J/���b���B(B33~B43���X���p��)
            strRealAmt = MidMbcs(strData, 33, 11)
            
            '�Y�J�b������ťաA�h��(B9~B14)��������
            If Len(strRealDate) = 0 Then
                '������������~
                strRealDate = Val(MidMbcs(strData, 9, 6)) + 19110000
            End If
            
            '�YPara24=1,�h�δC��渹�J�b
        If ChkData(strBillNo, strRealAmt, strClctEn, strClctName, strRealDate, strUpdEn, strCMCode, strCMName, strPTCode, strPTName, strBankName, IIf(strMedia = 1, False, True), strCitemQry, strMedia) = False Then
           IntoAccount = False
           GoTo Rollback
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
        Set fso = Nothing
  Exit Function
ChkErr:
  ErrSub "mediaCShing", "IntoAccount"
End Function
