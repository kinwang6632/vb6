VERSION 5.00
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Begin VB.Form frmUnion 
   BorderStyle     =   1  '��u�T�w
   ClientHeight    =   1995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6450
   Icon            =   "frmUnion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   6450
   StartUpPosition =   2  '�ù�����
   Begin VB.CommandButton cmdOK 
      Caption         =   "F2.�T�w"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   570
      TabIndex        =   4
      Top             =   1515
      Width           =   1410
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "����(&X)"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   4365
      TabIndex        =   5
      Top             =   1515
      Width           =   1410
   End
   Begin prjGiList.GiList gilCMCode 
      Height          =   315
      Left            =   2205
      TabIndex        =   1
      Top             =   2205
      Visible         =   0   'False
      Width           =   2940
      _ExtentX        =   5186
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�s�ө���"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      F2Corresponding =   -1  'True
   End
   Begin prjGiList.GiList gilClctEn 
      Height          =   315
      Left            =   2205
      TabIndex        =   2
      Top             =   675
      Width           =   2940
      _ExtentX        =   5186
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�s�ө���"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      F2Corresponding =   -1  'True
   End
   Begin prjGiList.GiList gilEmpNO 
      Height          =   315
      Left            =   2205
      TabIndex        =   0
      Top             =   270
      Width           =   2940
      _ExtentX        =   5186
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�s�ө���"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      F2Corresponding =   -1  'True
   End
   Begin Gi_Date.GiDate gdtRealDate 
      Height          =   330
      Left            =   2205
      TabIndex        =   3
      Top             =   1065
      Visible         =   0   'False
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   582
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblRealDate 
      Caption         =   "�J�b���"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1140
      TabIndex        =   9
      Top             =   1125
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "���O�H��"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1140
      TabIndex        =   8
      Top             =   360
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "���O�覡"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1140
      TabIndex        =   7
      Top             =   2280
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Label lblClctEn 
      AutoSize        =   -1  'True
      Caption         =   "�I�ں���"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1140
      TabIndex        =   6
      Top             =   750
      Width           =   780
   End
End
Attribute VB_Name = "frmUnion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
Private strOption As String          '�P�_���eoption�����A
Private strUPDName As String        '4388 �W�[��s���ʤH�� By Kin 2009/05/04
Private rsTmp As New ADODB.Recordset
Public IsTcb As Boolean
Public isTcbCard As Boolean

Public TableOwnerName As String
Public frmCreditCardType As mCreditCardType    ''  2003/10/30 �o�ӭȥN��ҭn�ϥΪ��H�Υd����
Public pstrCitemCode As String
Public pstrRealDatefrm As String
Private strErrCode As String
Private strErrDescription As String
Private blnFuBonIntegrate As Boolean

Private Sub IntoData()
On Error GoTo ChkErr

   With objStorePara
            strFilePath = .FilePath
            strRealDate = .RealDate
            strSourcePath = .SourcePath
            strUpdEn = .UpdEn
            strPrgName = .PrgName
            strPTCode = .PTCode
            strPTName = .PTName
            strServiceType = .ServiceType
            strCompCode = .CompCode
            strBankId = .BankId
            strUPDName = .UpdName
             GetOwner = TableOwnerName
   End With
   
   If InitData = False Then
            File.Close
            ErrFile.Close
            Set fso = Nothing
            Exit Sub
   End If
  Exit Sub
ChkErr:
  ErrSub Me.Name, "IntoData"
End Sub

Private Function IntoAccount() As Boolean

  On Error GoTo ChkErr

    Dim strBillNo As String
    Dim strRealAmt As String
    Dim strData As String
    Dim strYYMM As String
    Dim rsTmp1 As New ADODB.Recordset
    Dim strSQL As String
    Dim strCM As String
    
    Dim strEmpNO  As String
    Dim strEmpName As String
    Dim strUPDTime As String
 ''�H�U�o����ܼơA�ΥH�x�s���b���ѮɡA���o���Ȥ�s��
    Dim lngErrCust   As Long
    Dim strErrCust As String
    strFailUpdSQL = Empty
    strUPDTime = GetDTString(RightNow)
    lngCount = 0
    lngErrCount = 0
    
    TmpFilePath = strFilePath & "\Tmp" & strPrgName & "In.log"
    ErrFilePath = strFilePath & "\" & strPrgName & "Err.log"
    
    Call GetFilePath(strFilePath)
    
    fso.CopyFile strSourcePath, TmpFilePath
    Set File = fso.OpenTextFile(TmpFilePath, ForReading)
    Set ErrFile = fso.CreateTextFile(ErrFilePath)
    
    strNowTime = Now
    lngTime = Timer
        
    If Len(gilEmpNO.GetCodeNo) = 0 Then
        strEmpNO = ""
        strEmpName = ""
    Else
        strEmpNO = gilEmpNO.GetCodeNo
        strEmpName = gilEmpNO.GetDescription
    End If
    
        ''  �H�U�o�ӧP�_���ϥΩ� TCB �L��d���ɭԡA�ΥH���o�̫�@���Ʀ檺�H�Υd�B�z��
        
    If isTcbCard = False And IsTcb = True Then
        Dim lastFile As TextStream
        Dim lastLine As String
        Dim strNoTcbCardDate As String
                
        Set lastFile = fso.OpenTextFile(TmpFilePath, ForReading)
        While Not lastFile.AtEndOfLine   'Ū����b���
            lastLine = ""
            lastLine = lastFile.ReadLine
            ''  strNoTcbCardDate = Mid(lastLine, 20, 8)
            ''  2003/07/17  �令�q 16 �X���o�H�Υd�B�z��
            strNoTcbCardDate = Mid(lastLine, 18, 8)
        Wend
         lastFile.Close
    End If
    
    
    gcnGi.BeginTrans
        ''  20040827 �[�W����H�U���P�_
        '#5494 �A�W�[�ŷs��ު��P�_,�D�n�O���n�P�_���Y����� By Kin 2010/02/02
        '#5609 �W�[�p�X�H�Υd�妸 By Kin 2010/06/07
        '#7765 ignore the headline by kin 2018/05/09
    If (frmCreditCardType = lngHSBC) Or (frmCreditCardType = lngChinatrust) Or (frmCreditCardType = lngNewWeb) Or (frmCreditCardType = lngFubon) Then
        strData = File.ReadLine
    End If
    Dim aRealDate As String
    While Not File.AtEndOfLine   'Ū����b���
        strData = File.ReadLine
        aRealDate = ""
        Dim strAuthorizeNo As String  '' �o���ܼưO�����v�X
        Dim strNote As String
        strAuthorizeNo = Empty
        strNote = Empty
        '****************************************************************************************
        '#3417 ���ѧ�sUCCode�PUCName��SQL�򥻻y�k By Kin 2007/12/13
        '#4388 �W�[��s���ʤH���P���ʮɶ� By Kin 2009/05/04
        strFailUpdSQL = Empty
        strFailUpdSQL = "UPDATE " & TableOwnerName & "SO033 Set UCCode=" & strUCCode & _
                    ",UCName='" & strUCName & "',UPDEN='" & strUPDName & "'" & _
                    ",UPDTime='" & strUPDTime & "',Note=Note||"
        '****************************************************************************************
        
        strAuthorizeNo = ""
       ''  2003/10/30   �H�U�s�W�s���i����ӻȪ��P�_
        
        Select Case frmCreditCardType
            Case 0
                If IsTcb = False Then     ''  �p�G�O���ذӻȪ��榡�A�h���t�~�@�Ӱj��
                                 '' �J�b�渹 ���^�n�渹 (B097~B112)
                    strBillNo = Trim(Mid(strData, 97, 16))
                      ''  """"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
                      '' �ˬd���v�X(B46~B53)  �A�Y���b�����\�A�h�g�J���~log�ɡA���~���ƥ[��
                      '' ���v�X�O�Ʀr�~�i�H
                       ''---------------------------------------------------------------------------------------------------------
                       ''   2002/11/21   -- �վ�o�@�q �ˬd���v�X���W�h --
                       ''  bytes 46~53�����v�Xalltrim��,���� >= 4�X(���׬��Ʀr�Φr��),
                       ''  �Bbytes 69~84�e8�X='APPROVED',�N��@���\,����J�t�Φ��O��Ƥ�.
                      ''---------------------------------------------------------------------------------------------------------
                    Dim blnApproved  As Boolean
                    '#5395 APPROVED�쥻�u�P�_�j�g,�{�b�j�p�g���P�_
                    blnApproved = Len(Trim(Mid(strData, 46, 8))) >= 4 And UCase(Mid(strData, 69, 8)) = "APPROVED"
                       ''If Not (Len(Trim(Mid(strData, 46, 8))) >= 4 And Mid(strData, 46, 8) = "APPROVED") Then
                    If blnApproved = False Then
                                        ''  2003/07/10  �s�W�H�U�{���X���o���b���Ѥ��Ȥ�s��
                        lngErrCust = GetfFailCustID(strBillNo)
                        If lngErrCust = -1 Then
                            strErrCust = "�d���즹��ڽs��"
                        Else
                            '' strErrCust = CStr(strBillNo)
                            strErrCust = lngErrCust
                        End If
                         
                        ErrFile.WriteLine ("��ڽs���G " & strBillNo & " �H�Υd���b����" & "�F�Ȥ�s���G" & strErrCust)
                        
                        '*****************************************************************************************************************************
                        '#3417 ���Ѯɭn��sUCCode�PUCName By Kin 2007/12/13
                        If blnFailUpd Then
                            strFailUpdSQL = strFailUpdSQL & "'��ڽs���G " & strBillNo & " �H�Υd���b����" & "�F�Ȥ�s���G" & strErrCust & "'" & _
                                          " Where " & IIf(blnMultiAccount, "MediaBillNo='", "BillNo='") & strBillNo & "'" & _
                                           " And CancelFlag <> 1 And UCCode Is Not Null"
                            gcnGi.Execute strFailUpdSQL
                        End If
                        '*****************************************************************************************************************************
                        
                        lngErrCount = lngErrCount + 1
                        GoTo Nextloop
                    End If
                    If Not IsNumeric(Mid(strData, 38, 8)) Then
                        strRealAmt = 0
                    Else
                        strRealAmt = CStr(Val(Mid(strData, 38, 8)))
                    End If
                    strRealDate = Mid(strData, 85, 6)
                    '#5902 �p�G�J�����100,�h�e���n��1 By Kin 2011/01/11
                    If Left(strRealDate, 1) = "0" Then
                        strRealDate = "1" & strRealDate
                    Else
                        strRealDate = "0" & strRealDate
                    End If
                    '#5902 �쥻�u��6�X,�ҥH�~��2��,�{�b�j���̫e����1��,�ҥH�n��3�X By Kin 2011/01/11
                    strRealDate = Trim(CStr(2000 + Val(Mid(strRealDate, 1, 3))) & Mid(strRealDate, 3, 4))
                    strAuthorizeNo = Mid(strData, 46, 8)
                Else    ''  �H�U�o�ӬO���ذӻȪ��榡��
                    If isTcbCard = True Then
                        If blnMultiAccount = True Then
                            strBillNo = Mid(strData, 34, 11)   '' �O�h�b�����N�������K�O�F�A�]�����O�����s�i�h��
                        Else
                            strBillNo = GetBillNo_New(Trim(Mid(strData, 34, 11))) '' ���ন  16 �X���榡
                        End If
                       ''  �ثe
                        If Trim(MidMbcs(strData, 139, 1)) <> "Y" Then
                           ''  2003/07/10  �s�W�H�U�{���X���o���b���Ѥ��Ȥ�s��
                            lngErrCust = GetfFailCustID(strBillNo)
                            If lngErrCust = -1 Then
                                strErrCust = "�d���즹��ڽs��"
                            Else
                                strErrCust = lngErrCust
                            End If
                            ErrFile.WriteLine ("��ڽs���G " & strBillNo & " �H�Υd���b����" & "�F�Ȥ�s���G" & strErrCust)
                            
                            '*****************************************************************************************************************************
                            '#3417 ���Ѯɭn��sUCCode�PUCName By Kin 2007/12/13
                            If blnFailUpd Then
                                strFailUpdSQL = strFailUpdSQL & "'��ڽs���G " & strBillNo & " �H�Υd���b����" & "�F�Ȥ�s���G" & strErrCust & "'" & _
                                              " Where " & IIf(blnMultiAccount, "MediaBillNo='", "BillNo='") & strBillNo & "'" & _
                                               " And CancelFlag <> 1 And UCCode Is Not Null"
                                gcnGi.Execute strFailUpdSQL
                            End If
                            '*****************************************************************************************************************************
                            
                            lngErrCount = lngErrCount + 1
                            GoTo Nextloop
                        End If
                        If Not IsNumeric(Mid(strData, 66, 7)) Then
                            strRealAmt = 0
                        Else
                            strRealAmt = CStr(Val(Mid(strData, 66, 7)))
                        End If
                        strRealDate = Mid(strData, 45, 6)
                        strRealDate = Trim(CStr(1911 + Val(Mid(strRealDate, 1, 2))) & Mid(strRealDate, 3, 4))
                        strAuthorizeNo = Mid(strData, 51, 6)
                    Else
                        strBillNo = Trim(Mid(strData, 17, 15))
                        If Trim(Mid(strData, 96, 2)) <> "00" Then
                            If Left(strData, 1) = "T" Then GoTo Nextloop
                            ''  2003/07/10  �s�W�H�U�{���X���o���b���Ѥ��Ȥ�s��
                            lngErrCust = GetfFailCustID(strBillNo)
                            If lngErrCust = -1 Then
                                strErrCust = "�d���즹��ڽs��"
                            Else
                                strErrCust = lngErrCust
                            End If
                            ErrFile.WriteLine ("��ڽs���G " & strBillNo & " �H�Υd���b����" & "�F�Ȥ�s���G" & strErrCust)
                            
                            '*****************************************************************************************************************************
                            '#3417 ���Ѯɭn��sUCCode�PUCName By Kin 2007/12/13
                            If blnFailUpd Then
                                strFailUpdSQL = strFailUpdSQL & "'��ڽs���G " & strBillNo & " �H�Υd���b����" & "�F�Ȥ�s���G" & strErrCust & "'" & _
                                              " Where " & IIf(blnMultiAccount, "MediaBillNo='", "BillNo='") & strBillNo & "'" & _
                                               " And CancelFlag <> 1 And UCCode Is Not Null"
                                gcnGi.Execute strFailUpdSQL
                            End If
                            '*****************************************************************************************************************************
                            
                            lngErrCount = lngErrCount + 1
                            GoTo Nextloop
                        End If
                         ''  --2003/07/17--   �t�X�X��W��վ�A�o�ӥ��ݦ^����Ӧ줸�A�쥻��  58, 12
                        If Not IsNumeric(Mid(strData, 58, 10)) Then
                            strRealAmt = 0
                        Else
                            strRealAmt = CStr(Val(Mid(strData, 58, 10)))
                        End If
                        '�J�b���
                        strRealDate = strNoTcbCardDate
                        strAuthorizeNo = Mid(strData, 98, 6)
                    End If
                End If     ''  �p�G�O���ذӻȪ��榡�A�h���t�~�@�Ӱj����  if  �P�_����
                                         
            Case 1    ''  2003/10/09  �H�U�]����ӻȮ榡
                  '����H�U�榡   2007/2/8  By Kin 2007/2/9 ����
        '' *************************************************************************************************************
                If blnMultiAccount = True Then
                    strBillNo = (Mid(strData, 49, 11))
                Else
                    strBillNo = GetBillNo_New((Mid(strData, 50, 11)))
                End If
           ''�P�_44~45������X
                Dim intTranCode As Integer
                Dim blnApprovedChina  As Boolean
                intTranCode = Val(Trim(Mid(strData, 44, 2)))
         '' 20040906 �令 �� 75 ��쪺8  �X
                '#5395 APPROVED�쥻�u�P�_�j�g,�{�b�j�p�g���P�_
                blnApprovedChina = UCase(Mid(strData, 75, 8)) = "APPROVED"
           
                If blnApprovedChina = False Then
                    lngErrCust = GetfFailCustID(strBillNo)
                    If lngErrCust = -1 Then
                        strErrCust = "�d���즹��ڽs��"
                    Else
                        strErrCust = lngErrCust
                    End If
                    ErrFile.WriteLine ("��ڽs���G " & strBillNo & " �H�Υd���b����" & "�F�Ȥ�s���G" & strErrCust)
                    '*****************************************************************************************************************************
                    '#3417 ���Ѯɭn��sUCCode�PUCName By Kin 2007/12/13
                    If blnFailUpd Then
                        strFailUpdSQL = strFailUpdSQL & "'��ڽs���G " & strBillNo & " �H�Υd���b����" & "�F�Ȥ�s���G" & strErrCust & "'" & _
                                      " Where " & IIf(blnMultiAccount, "MediaBillNo='", "BillNo='") & strBillNo & "'" & _
                                       " And CancelFlag <> 1 And UCCode Is Not Null"
                        gcnGi.Execute strFailUpdSQL
                    End If
                    '*****************************************************************************************************************************
                    
                    lngErrCount = lngErrCount + 1
                    GoTo Nextloop
                End If
'�@�@�@�@�@**********************************************************************************************************
            '#3049����H�U�쥻���P�_�^�¹q�l�ɪ����B�A�אּ�P�_�^�¹q�l�ɪ����B By Kin 2007/2/9
             ''2003/10/13  �o�@�q�����κޥL�F�A�]���S���^�J�b���B�Ajacy�^�Цp�O��(����\��#3049�A�N���^�_)
                If Not IsNumeric(Mid(strData, 30, 10)) Then
                    strRealAmt = 0
                Else
                    strRealAmt = CStr(Val(Mid(strData, 30, 10)))
                End If
   '�@�@�@�@************************************************************************************************************
           '�J�b���
                strRealDate = Mid(strData, 91, 6)
                strRealDate = Trim(CStr(2000 + Val(Mid(strRealDate, 1, 2))) & Mid(strRealDate, 3, 4))
                '' 20040906 �令 �� 60��쪺6�X
                strAuthorizeNo = Mid(strData, 60, 6)
        '' * ************************************************************************************************************
            Case 2     ''  2003/12/02 �H�U�]�s���׻Ȧ�榡
                
                Dim blnApprovedHSBC  As Boolean
                If blnMultiAccount = True Then   ''  so043.para24=1�h���C��渹11�X
                    strBillNo = (Mid(strData, 53, 11))
                Else                                             ''  so043.para24=0 �h���C��渹15�X
                    strBillNo = GetBillNo_New((Mid(strData, 53, 11)))
                End If
              
                If Trim(Mid(strData, 50, 2)) <> "00" Then
                    blnApprovedHSBC = False
                    ErrFile.WriteLine (GetHSBCResponseMessage(Trim(Mid(strData, 50, 2))))
                    
                    '*****************************************************************************************************************************
                    '#3417 ���Ѯɭn��sUCCode�PUCName By Kin 2007/12/13
                    If blnFailUpd Then
                        strFailUpdSQL = strFailUpdSQL & "';�^���X:" & Trim(Mid(strData, 50, 2)) & ";�T��:" & _
                                        GetHSBCResponseMessage(Trim(Mid(strData, 50, 2))) & "'" & _
                                      " Where " & IIf(blnMultiAccount, "MediaBillNo='", "BillNo='") & strBillNo & "'" & _
                                       " And CancelFlag <> 1 And UCCode Is Not Null"
                        gcnGi.Execute strFailUpdSQL
                    End If
                    '*****************************************************************************************************************************
                Else
                    blnApprovedHSBC = True
                End If
    
                If blnApprovedHSBC = False Then
                    lngErrCust = GetfFailCustID(strBillNo)
                    If lngErrCust = -1 Then
                        strErrCust = "�d���즹��ڽs��"
                    Else
                        strErrCust = lngErrCust
                    End If
                    ErrFile.WriteLine ("��ڽs���G " & strBillNo & " �H�Υd���b����" & "�F�Ȥ�s���G" & strErrCust)
                    
                    '*****************************************************************************************************************************
                    '#3417 ���Ѯɭn��sUCCode�PUCName By Kin 2007/12/13
                    If blnFailUpd Then
                        strFailUpdSQL = strFailUpdSQL & "'��ڽs���G " & strBillNo & " �H�Υd���b����" & "�F�Ȥ�s���G" & strErrCust & "'" & _
                                      " Where " & IIf(blnMultiAccount, "MediaBillNo='", "BillNo='") & strBillNo & "'" & _
                                       " And CancelFlag <> 1 And UCCode Is Not Null"
                        gcnGi.Execute strFailUpdSQL
                    End If
                    '*****************************************************************************************************************************
                    
                    lngErrCount = lngErrCount + 1
                    GoTo Nextloop
                End If
                If Not IsNumeric(Mid(strData, 30, 10)) Then
                    strRealAmt = 0
                Else
                    strRealAmt = CStr(Val(Mid(strData, 30, 10)))  ''���X�O�����A�]�������u�� 10 �X
                End If
             '' 204/02/12 �s�W�@�ӤJ�b������A�p�G�o����즳��ȡA�H�o����쬰��
                If Len(gdtRealDate.Text) > 0 Then
                    strRealDate = gdtRealDate.ClipText
                Else
                    strRealDate = Mid(strData, 1, 6)
                    strRealDate = Trim(CStr(1911 + Val(Mid(strRealDate, 1, 2))) & Mid(strRealDate, 3, 4))
                End If
            ''  �H�U�����v�X �A���ݶ�J SO033   AuthorizeNo  ��� .....
                strAuthorizeNo = Trim(Mid(strData, 43, 6))
                
                
            '#3697 �W�[�@�դ���H�U�s�榡 By Kin 2007/12/31
            Case 3
                strErrCode = ""
                If Mid(strData, 1, 1) = "H" Then GoTo Nextloop
                If Mid(strData, 1, 1) = "R" Then GoTo Nextloop
                If Mid(strData, 1, 1) = "" Then GoTo Nextloop
                If Mid(strData, 1, 1) = "T" Then
                    strErrCode = MidMbcs(strData, 43, 2)
                    strBillNo = Trim(Mid(strData, 45, 15))
                    If strErrCode <> "00" Then
                        '#3947 ���~�T���W�[��ڽs���P�Ȥ�s�� By Kin 2008/05/26
                        lngErrCust = GetfFailCustID(strBillNo)
                        If lngErrCust = -1 Then
                            strErrCust = "�d���즹��ڽs��"
                        Else
                            strErrCust = lngErrCust
                        End If
                        ErrFile.WriteLine ("��ڽs��: " & strBillNo & " �Ȥ�s��: " & strErrCust & " �J�b���ѡI ���~�N�X�G" & strErrCode & " ���~��]�G" & GetChinatrusNewMessage(strErrCode))
                        '*****************************************************************************************************************************
                        '#3417 ���Ѯɭn��sUCCode�PUCName By Kin 2007/12/31
                        If blnFailUpd Then
                            strFailUpdSQL = strFailUpdSQL & "'�J�b���ѡI ���~�N�X�G" & strErrCode & " ���~��]�G" & GetChinatrusNewMessage(strErrCode) & "'" & _
                                          " Where " & IIf(blnMultiAccount, "MediaBillNo='", "BillNo='") & strBillNo & "'" & _
                                           " And CancelFlag <> 1 And UCCode Is Not Null"
                            gcnGi.Execute strFailUpdSQL
                        End If
                        '*****************************************************************************************************************************
    
                        lngErrCount = lngErrCount + 1
                        GoTo Nextloop
                    Else
                        lngErrCust = GetfFailCustID(strBillNo)
                        If lngErrCust = -1 Then
                            strErrCust = "��ڡG" & strBillNo & "�@���s�b�I�Ьd��"
                            ErrFile.WriteLine (strErrCust)
                            '*****************************************************************************************************************************
                            '#3417 ���Ѯɭn��sUCCode�PUCName By Kin 2007/12/31
                            If blnFailUpd Then
                                strFailUpdSQL = strFailUpdSQL & "'" & strErrCust & "'" & _
                                              " Where " & IIf(blnMultiAccount, "MediaBillNo='", "BillNo='") & strBillNo & "'" & _
                                               " And CancelFlag <> 1 And UCCode Is Not Null"
                                gcnGi.Execute strFailUpdSQL
                            End If
                            '*****************************************************************************************************************************
                            lngErrCount = lngErrCount + 1
                            GoTo Nextloop
                        Else
                            If Not IsNumeric(Mid(strData, 30, 13)) Then
                                strRealAmt = 0
                            Else
                                strRealAmt = CStr(Val(Mid(strData, 30, 11)))  ''���X�O�����A�]�������u�� 10 �X
                            End If
                            If Len(gdtRealDate.Text) > 0 Then
                                strRealDate = gdtRealDate.GetValue
                            Else
                                
                            End If
                            
                            strAuthorizeNo = Trim(MidMbcs(strData, 67, 6))

                        End If
    
                    End If
                End If
            '#5494�W�[�@���ŷs��� By Kin 2010/02/02
            Case 4
                strBillNo = Trim(Replace((Mid(strData, 2, 30)), " ", ""))
                If Mid(strData, 98, 2) <> "00" Then
                    ErrFile.WriteLine ("��ڽs��: " & strBillNo & _
                        " ;�^���X:" & Mid(strData, 98, 2) & " ;�^���T��:" & GetNewWebMessage(Trim(Mid(strData, 98, 2))))
                    
                    If blnFailUpd Then
                        strFailUpdSQL = strFailUpdSQL & "';�^���X:" & Trim(Mid(strData, 98, 2)) & ";�T��:" & _
                                        GetNewWebMessage(Trim(Mid(strData, 98, 2))) & "'" & _
                                      " Where " & IIf(blnMultiAccount, "MediaBillNo='", "BillNo='") & strBillNo & "'" & _
                                       " And CancelFlag <> 1 And UCCode Is Not Null"
                        gcnGi.Execute strFailUpdSQL
                    End If
                    lngErrCount = lngErrCount + 1
                    GoTo Nextloop
                Else
                    lngErrCust = GetfFailCustID(strBillNo)
                    If lngErrCust = -1 Then
                        strErrCust = "��ڡG" & strBillNo & "�@���s�b�I�Ьd��"
                        ErrFile.WriteLine (strErrCust)
                        
                        If blnFailUpd Then
                            strFailUpdSQL = strFailUpdSQL & "'" & strErrCust & "'" & _
                                          " Where " & IIf(blnMultiAccount, "MediaBillNo='", "BillNo='") & strBillNo & "'" & _
                                           " And CancelFlag <> 1 And UCCode Is Not Null"
                            gcnGi.Execute strFailUpdSQL
                        End If
                        
                        lngErrCount = lngErrCount + 1
                        GoTo Nextloop
                    Else
                        strRealAmt = CStr(Val(GetAmt(Mid(strData, 44, 13))))
                        If Len(gdtRealDate.Text) > 0 Then
                            strRealDate = gdtRealDate.GetValue
                        End If
                        '�^�нX�n�^��SO033 By Kin 2010/02/03
                        strAuthorizeNo = Trim(MidMbcs(strData, 82, 6))
                    
                        If Trim(MidMbcs(strData, 32, 9)) <> "" Then
                            strNote = "�ŷs�q��s��:" & Trim(MidMbcs(strData, 32, 9))
                        End If
                    End If
                End If
            Case 5
                strBillNo = Trim(Replace((Mid(strData, 97, 16)), " ", ""))
                If Val(Mid(strData, 66, 3) & "") > 9 Then
                    ErrFile.WriteLine ("��ڽs��: " & strBillNo & _
                        " ;�^���X:" & Mid(strData, 66, 3) & " ;�^���T��:" & GetUnionBathMessage(Mid(strData, 66, 3), strData))
                        
                        'GetUnionBathMessage(Trim(Mid(strData, 66, 3))))
                    
                    If blnFailUpd Then
                        strFailUpdSQL = strFailUpdSQL & "';�^���X:" & Trim(Mid(strData, 66, 3)) & ";�T��:" & _
                                        GetUnionBathMessage(Mid(strData, 66, 3), strData) & "'" & _
                                      " Where " & IIf(blnMultiAccount, "MediaBillNo='", "BillNo='") & strBillNo & "'" & _
                                       " And CancelFlag <> 1 And UCCode Is Not Null"
                        gcnGi.Execute strFailUpdSQL
                    End If
                    lngErrCount = lngErrCount + 1
                    GoTo Nextloop
                Else
                    lngErrCust = GetfFailCustID(strBillNo)
                    If lngErrCust = -1 Then
                        strErrCust = "��ڡG" & strBillNo & "�@���s�b�I�Ьd��"
                        ErrFile.WriteLine (strErrCust)
                        
                        If blnFailUpd Then
                            strFailUpdSQL = strFailUpdSQL & "'" & strErrCust & "'" & _
                                          " Where " & IIf(blnMultiAccount, "MediaBillNo='", "BillNo='") & strBillNo & "'" & _
                                           " And CancelFlag <> 1 And UCCode Is Not Null"
                            gcnGi.Execute strFailUpdSQL
                        End If
                        lngErrCount = lngErrCount + 1
                        GoTo Nextloop
                    Else
                        strRealAmt = CStr(Val(Mid(strData, 38, 8) & ""))
                        If Len(gdtRealDate.Text) > 0 Then
                            strRealDate = gdtRealDate.GetValue
                        Else
                            
                            aRealDate = Mid(strData, 113, 6)
                            If Left(aRealDate, 1) = "0" Then
                                aRealDate = "1" & aRealDate
                            End If
                            strRealDate = CStr(Val(aRealDate) + 19110000)
                        End If
                        strAuthorizeNo = Replace(Mid(strData, 46, 8), " ", "")
                    End If
                End If
            '#7765 By Kin 2018/05/09
            Case 6
                If Mid(strData, 1, 1) = "T" Then GoTo Nextloop
                If LenB(strData) < 300 Then
                    ErrFile.WriteLine ("�^���ɮ榡�����T")
                    lngErrCount = lngErrCount + 1
                    GoTo Nextloop
                End If
                strBillNo = Trim(Replace((Mid(strData, 18, 25)), " ", ""))
                '7916 using table to write the description of the response by kin 2018/10/05
                Dim StrResponCode As String
                Dim strResponDes As String
                Dim intReplyLen As Byte
                Dim strSuccess1 As String
                Dim strSuccess2 As String
                '#8528 By Kin 2019/12/03
                intReplyLen = 3
                If blnFuBonIntegrate Then intReplyLen = 4
                '8575 stop checking the bank by kin 2020/02/21
                intReplyLen = 3
                strSuccess1 = Right("0000", intReplyLen)
                strSuccess2 = Right("0001", intReplyLen)
                StrResponCode = Replace(Mid(strData, 199, intReplyLen), " ", "")
                strResponDes = GetBankRespons("CREDITCARDFUBON", StrResponCode)
                'CREDITCARDFUBON
                '8575 stop checking the bank by kin 2020/02/21
                If Right("0000" & StrResponCode, 4) = "0006" And blnFuBonIntegrate And 1 = 0 Then
                    
                End If
                If StrResponCode <> strSuccess1 And StrResponCode <> strSuccess2 Then
                '8575 stop checking the bank by kin 2020/02/21
                    If Not blnFuBonIntegrate And Right("0000" & StrResponCode, 4) = "0006" And 1 = 0 Then
                    
                    Else
                        ErrFile.WriteLine ("��ڽs��: " & strBillNo & _
                        " ;�^���X:" & StrResponCode & ";�T��:" & strResponDes)
                        
                        If blnFailUpd Then
                            strFailUpdSQL = strFailUpdSQL & "'�J�b���ѡI�^���X:" & StrResponCode & ";�T��:" & _
                                           strResponDes & "'" & _
                                          " Where " & IIf(blnMultiAccount, "MediaBillNo='", "BillNo='") & strBillNo & "'" & _
                                           " And CancelFlag <> 1 And UCCode Is Not Null"
                            gcnGi.Execute strFailUpdSQL
                        End If
                        lngErrCount = lngErrCount + 1
                        GoTo Nextloop
                    End If
                    
                End If
                
                If Len(Replace(Mid(strData, 216, 8), " ", "")) = 0 Then
                    ErrFile.WriteLine ("��ڽs��: " & strBillNo & _
                        " ;�^���X:" & StrResponCode & " ;�T��:���v�X�L��")
                                                
                    
                    If blnFailUpd Then
                        strFailUpdSQL = strFailUpdSQL & "'�J�b���ѡI�^���X:" & StrResponCode & ";�T��:" & _
                                        "���v�X�L��" & "'" & _
                                      " Where " & IIf(blnMultiAccount, "MediaBillNo='", "BillNo='") & strBillNo & "'" & _
                                       " And CancelFlag <> 1 And UCCode Is Not Null"
                        gcnGi.Execute strFailUpdSQL
                    End If
                    lngErrCount = lngErrCount + 1
                    GoTo Nextloop
                Else
                    lngErrCust = GetfFailCustID(strBillNo)
                    If lngErrCust = -1 Then
                        strErrCust = "��ڡG" & strBillNo & "�@���s�b�I�Ьd��"
                        ErrFile.WriteLine (strErrCust)
                        
                        If blnFailUpd Then
                            strFailUpdSQL = strFailUpdSQL & "'" & strErrCust & "'" & _
                                          " Where " & IIf(blnMultiAccount, "MediaBillNo='", "BillNo='") & strBillNo & "'" & _
                                           " And CancelFlag <> 1 And UCCode Is Not Null"
                            gcnGi.Execute strFailUpdSQL
                        End If
                        lngErrCount = lngErrCount + 1
                        GoTo Nextloop
                    Else
                        strRealAmt = CStr(Val(Mid(strData, 46, 11) & ""))
                        If Len(gdtRealDate.Text) > 0 Then
                            strRealDate = gdtRealDate.GetValue
                        Else
                            aRealDate = Mid(strData, 145, 8)
                            If Len(Replace(aRealDate, " ", "")) = 0 Then
                                aRealDate = Format(Now, "yyyyMMdd")
                            End If
                        End If
                        If Len(aRealDate) > 0 Then
                            strRealDate = aRealDate
                        End If
                        strAuthorizeNo = StrResponCode
                        'strAuthorizeNo = Replace(Mid(strData, 46, 8), " ", "")
                    End If
                End If
        End Select
       
        strClctEn = gilClctEn.GetCodeNo
        strClctName = gilClctEn.GetDescription
        If Len(gilCMCode.GetCodeNo & "") > 0 Then
            strCMCode = gilCMCode.GetCodeNo
            strCMName = gilCMCode.GetDescription
        End If
        '' 20040107 �s�W�Y�O�e���W���ڤJ�b������ȡA�h�H��Ƕi�Ӫ�����
        If Len(pstrRealDatefrm) > 0 Then strRealDate = pstrRealDatefrm
        '*****************************************************************************************
        '#3049�쥻�ˮ֪�Function��ChkData�A�{�אּChkDataTo3313�A�NFunction�W�ߡA�H�K�v�Q�䥦�{��
        '#3417 �h�ǻȦ�N�X,�p�G���O�H�����Ȫ���,SO033.NOTE�n��J�Ȧ�N�X By Kin 2008/01/22
        '#5275 ��s���ʤH���OName,���OCode,�쥻�ѼƬ� strUpdEn,�{�b�אּstrUPDName By Kin 2009/09/07
        '#5494 �ŷs���32~40�r���n��J�Ƶ���,�ҥH�W�[vNote���Ѽ� By Kin 2010/02/03
        If ChkDataTo3313(strBillNo, strRealAmt, strRealDate, strUPDName, _
                        strCMCode, strCMName, strClctEn, strClctName, _
                        strEmpNO, strEmpName, strAuthorizeNo, frmCreditCardType, _
                        pstrCitemCode, strServiceType, strBankId, strNote) = False Then
            IntoAccount = False
            GoTo Roolback
        End If
        '******************************************************************************************
Nextloop:
    Wend
    gcnGi.CommitTrans
    IntoAccount = True
    On Error Resume Next
    rsTmp1.Close
    Set rsTmp1 = Nothing
    
  Exit Function
Roolback:
    gcnGi.RollbackTrans
    Me.MousePointer = vbDefault
    MsgBox "��s����!", vbExclamation, "ĵ�i..."
    File.Close
    ErrFile.Close
    Set fso = Nothing
    Exit Function
ChkErr:
    ErrSub "frmUnion", "IntoAccount," & strBillNo & "," & strData
End Function
Private Function GetBankRespons(ByVal PrgName As String, ByVal ErrCode As String) As String
  On Error Resume Next
    Dim ret  As String
    Dim aSQL As String
    Dim aWhere As String
    '8575 stop checking the bank by kin 2020/02/21
    If blnFuBonIntegrate And 1 = 0 Then
        aWhere = " And  nvl(PlatForm,0) = 1 "
    Else
        aWhere = " And  nvl(PlatForm,0) = 0 "
    End If
    aSQL = "Select ErrDescription From " & GetOwner & "SO033G Where upper(PrgName)=upper('" & PrgName & "') " & _
            " And ErrorCode = '" & ErrCode & "' " & aWhere
    ret = GetRsValue(aSQL, gcnGi) & ""
    If Len(ret) = 0 Then
        ret = "�Ȧ�^�Х���(�䤣��^����)"
    End If
    GetBankRespons = ret
End Function
'#5494 �ŷs�榡�����B,�̫�2�X���p���I By Kin 2010/02/03
Private Function GetAmt(ByVal aData As String) As String
  On Error GoTo ChkErr
    Dim retAmt As String
    retAmt = Empty
    retAmt = CStr(Val(Mid(aData, 1, Len(aData) - 2))) & "." & Right(aData, 2)
    GetAmt = retAmt
    Exit Function
ChkErr:
  Call ErrSub("frmUnion", "GetAmt")
End Function
Private Sub cmdCancel_Click()
On Error Resume Next
   Unload Me
   
End Sub

Private Sub cmdOK_Click()
  On Error GoTo ChkErr
    Dim rsCD013 As New ADODB.Recordset
    blnFailUpd = False
    strUCCode = Empty
    strUCName = Empty
    If Not IsDataOk Then Exit Sub
    lngErrCount = 0
    lngAmtTotal = 0
    lngStatusCount = 0
    Me.MousePointer = vbHourglass
    '********************************************************************************************************************************************************
    '#3417 �q�l�ɶץX��,�n��J������](RefNo=4) By Kin 2007/12/13
    If Not GetRS(rsCD013, "Select * From " & TableOwnerName & "CD013 Where RefNo=6 And StopFlag<>1 Order By CodeNo Desc", gcnGi, adUseClient, adOpenKeyset) Then Exit Sub
    If rsCD013.EOF Then
        blnFailUpd = False
    Else
        blnFailUpd = True
        strUCCode = rsCD013("CodeNo") & ""
        strUCName = rsCD013("Description") & ""
    End If
    '*********************************************************************************************************************************************************
     '7179
    blnCrossCustCombine = Val(GetRsValue("Select Nvl(CrossCustCombine,0) From " & frmUnion.TableOwnerName & "SO041", gcnGi) & "")
    blnUpdate = IntoAccount '�}�l�J�b
    If blnUpdate = False Then  '��s���ѡA�N�����ɮ�����
        objStorePara.UpdState = False
        Exit Sub
    Else
       objStorePara.UpdState = True
    End If
    
    Call DefineRS
    Call ScrToRcd
    
    Call MsgResult       '��ܵ��G�T��
    '�Y���~���Ƥj��0�h��NotePad������ɮפ��e
    If lngErrCount > 0 Then Call Shell("NotePad " & ErrFilePath, vbNormalFocus)
    If lngStatusCount > 0 Then Call Shell("NotePad " & ErrFilePath, vbNormalFocus)
    Me.MousePointer = vbDefault
    File.Close
    ErrFile.Close
    Set fso = Nothing
    
    Exit Sub
ChkErr:
   ErrSub Me.Name, "cmdOK_Click"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error Resume Next
        If Not ChkGiList(KeyCode) Then Exit Sub
        If KeyCode = vbKeyEscape Then Unload Me
        If KeyCode = vbKeyF2 Then If cmdOK.Enabled = True Then cmdOK.Value = True
End Sub

Private Sub Form_Load()
  On Error Resume Next
     Dim strCMCodeRefNo5 As String
     Dim strCMCodeRefNo4 As String
     Dim strPTCodeRefNo4 As String
     Me.Caption = objStorePara.BankName & ""
     Call IntoData
     Call subGil
     Call RcdToScr
    '#4022 �p�G�Ұʦh�C�^���P�_�A�ȧO By Kin 2008/07/22
    If strServiceType = Empty Then
        blnMultiAccount = IIf(Val(RPxx(gcnGi.Execute("SELECT para24 FROM " & TableOwnerName & "SO043 WHERE CompCode =" & strCompCode)(0))) = 1, True, False)
    Else
        blnMultiAccount = IIf(Val(RPxx(gcnGi.Execute("SELECT para24 FROM " & TableOwnerName & "SO043 WHERE CompCode =" & strCompCode & " AND ServiceType = '" & strServiceType & "'").GetString)) = 1, True, False)
    End If
    blnFuBonIntegrate = objStorePara.uFubonIntegrate
    
    If blnFuBonIntegrate And frmCreditCardType = lngFubon Then
        '#8555 Cancel to set the value of default and disable  the control by Kin 2020/01/13
        'strCMCodeRefNo5 = gcnGi.Execute("Select CodeNo From " & TableOwnerName & "CD031 Where Nvl(StopFlag,0) = 0 And RefNo= 5 And RowNum = 1 ")(0) & ""
        gilCMCode.Enabled = False
'        If Len(strCMCodeRefNo5) & "" <> "" Then
'            gilCMCode.SetCodeNo strCMCodeRefNo5
'            gilCMCode.Query_Description
'            'gilCMCode.Enabled = False
'        End If
    Else
     '#8555 Setting the value of default when not being fubonintegetrate by kin 2020/01/13
     '#8555 hidde the control
  '      strCMCodeRefNo4 = gcnGi.Execute("Select CodeNo From " & TableOwnerName & "CD031 Where Nvl(StopFlag,0) = 0 And RefNo= 4 And RowNum = 1  order by CodeNo")(0) & ""
  '      If Len(strCMCodeRefNo4 & "") > 0 Then
  '          gilCMCode.SetCodeNo strCMCodeRefNo4
  '          gilCMCode.Query_Description
  '      End If
    End If
    '#8555 Setting the value of default By Kin 2020/01/13
    strPTCodeRefNo4 = gcnGi.Execute("Select CodeNo From " & TableOwnerName & "CD032 Where Nvl(StopFlag,0) = 0 And RefNo= 4 And RowNum = 1 ")(0) & ""
    If Len(strPTCodeRefNo4 & "") > 0 Then
        gilClctEn.SetCodeNo strPTCodeRefNo4
        gilClctEn.Query_Description
    End If
'   gilEmpNO.SetCodeNo objStorePara.UpdEn
'   gilEmpNO.Query_Description
'   If Len(gilEmpNO.GetDescription & "") = 0 Then
'        gilEmpNO.Clear
'   End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
      File.Close
      ErrFile.Close
      Set fso = Nothing
End Sub

Private Sub OpenRec()
On Error GoTo ChkErr
  Dim strSQL As String
  Dim rsData As New ADODB.Recordset
  
     If rsData.State = adStateOpen Then rsData.Close
     strSQL = _
          "Select " & _
          "A.CodeNo,A.Description,A.ModeNo,B.Description ModeName " & _
          "From " & TableOwnerName & _
          "CD031A A," & TableOwnerName & "CD031 B " & _
          "Where " & _
          "A.ModeNo=B.CodeNo And A.StopFlag=0 And A.CompCode = " & strCompCode & " And (A.ServiceType = '" & strServiceType & "' Or A.ServiceType Is Null) Order By A.CodeNo"
  
     If rsData.State = adStateOpen Then
        rsData.Close
        Set rsData = Nothing
     End If
   Exit Sub
ChkErr:
   ErrSub Me.Name, "OpenRec"
End Sub
Private Sub subGil()
On Error GoTo ChkErr
    '#4022 �p�G�Ұʦh�C�^���P�_�A�ȧO By Kin 2008/07/22
    If strServiceType <> Empty Then
        SetLst gilClctEn, "CodeNo", "Description", 10, 20, , , "CD032", " Where (ServiceType Is Null Or ServiceType = '" & strServiceType & "'" & ")  AND Nvl(StopFlag,0) = 0   "
        SetLst gilCMCode, "CodeNo", "Description", 10, 20, , , "CD031", " Where (ServiceType Is Null Or ServiceType = '" & strServiceType & "'" & ")  AND Nvl(StopFlag,0) = 0  And RefNo = 4 "
    Else
        SetLst gilClctEn, "CodeNo", "Description", 10, 20, , , "CD032", " Where Nvl(StopFlag,0) = 0   "
        SetLst gilCMCode, "CodeNo", "Description", 10, 20, , , "CD031", " Where Nvl(StopFlag,0) = 0 And RefNo = 4   "
    End If
    SetLst gilEmpNO, "EmpNo", "EmpName", 10, 20, , , "CM003", " Where CompCode = " & strCompCode & "  AND Nvl(StopFlag,0)  = 0  "
Exit Sub
ChkErr:
   ErrSub Me.Name, "subGil"
End Sub

Private Sub RcdToScr()
On Error GoTo ChkErr
    
    If Dir(strFilePath & "\" & strPrgName & "In2.log") = "" Then
       
       gilClctEn.SetCodeNo ""
       gilClctEn.SetDescription ""
       gilCMCode.SetCodeNo ""
       gilCMCode.SetDescription ""
    Else
       With rsTmp
            If .State = adStateOpen Then .Close
            .Open strFilePath & "\" & strPrgName & "In2.log"
               gilClctEn.SetCodeNo .Fields("ClctEn") & ""
               gilClctEn.Query_Description
               gilCMCode.SetCodeNo .Fields("CMCode") & ""
               gilCMCode.Query_Description
       End With
    End If
   Exit Sub
ChkErr:
   ErrSub Me.Name, "RcdToScr"
End Sub

Private Sub ScrToRcd()
On Error GoTo ChkErr
    With rsTmp
       .AddNew
       .Fields("Option") = strOption
       If strOption = "Screen" Then
          .Fields("ClctEn") = gilClctEn.GetCodeNo
          .Fields("ClctName") = gilClctEn.GetDescription
          .Fields("CMCode") = gilCMCode.GetCodeNo
          .Fields("CMName") = gilCMCode.GetDescription
       Else
          .Fields("ClctEn") = ""
          .Fields("ClctName") = ""
          .Fields("CMCode") = ""
          .Fields("CMName") = ""
       End If
       .Save strFilePath & "\" & strPrgName & "In2.log"
    End With
  Exit Sub
ChkErr:
  ErrSub Me.Name, "ScrToRcd"
End Sub

Private Sub DefineRS()
On Error GoTo ChkErr
  If rsTmp.State = adStateOpen Then rsTmp.Close
    With rsTmp
       .LockType = adLockOptimistic
       .Fields.Append "Option", adBSTR, 12
       .Fields.Append "ClctEn", adBSTR, 3
       .Fields.Append "ClctName", adBSTR, 12
       .Fields.Append "CMCode", adBSTR, 3
       .Fields.Append "CMName", adBSTR, 12
       .Open
    End With
    
  Exit Sub
ChkErr:
  ErrSub Me.Name, "DefineRS"
End Sub

Private Function IsDataOk() As Boolean
  On Error GoTo ChkErr
    IsDataOk = False
    If Len(gilCMCode.GetCodeNo & "") = 0 Then
        If blnFuBonIntegrate And frmCreditCardType = lngFubon Then
        Else
         '   gilCMCode.SetFocus
         '   MsgBox "�п�ܦ��O�覡!", vbExclamation, "ĵ�i..."
        End If
    End If
    If Len(gilClctEn.GetCodeNo & "") = 0 Then
       gilClctEn.SetFocus
       MsgBox "�п�ܥI�ں���!", vbExclamation, "ĵ�i..."
       Exit Function
    End If
    IsDataOk = True
   Exit Function
   
ChkErr:
   ErrSub Me.Name, "IsDataOk"
End Function

Private Function GetfFailCustID(ByVal strID As String) As Long
    Dim strSQL As String
    Dim rs As New ADODB.Recordset
    Dim lngCustID As Long
    Dim strBillField As String
  
  On Error GoTo ChkErr
  
  '''�H�U�� BillNo ��� �令 MediaBillNo
  
    If blnMultiAccount = True Then
        strBillField = "MediaBillno"
    Else
        strBillField = "Billno"
    End If

    strSQL = "SELECT CUSTID FROM " & frmUnion.TableOwnerName & "SO033 WHERE " & strBillField & " ='" & strID & "'"
    '7179
    If blnCrossCustCombine Then
        strSQL = "Select A.*,(Select MainCustId From " & frmUnion.TableOwnerName & "SO202 Where MduId=SO001.AMduId) MainCustId " & _
                " From (" & strSQL & ") A," & frmUnion.TableOwnerName & "SO001 Where A.CustId=SO001.CustId "
    End If
    rs.Open strSQL, gcnGi, adOpenKeyset, adLockReadOnly
    
    If rs.EOF Or rs.BOF Then
        lngCustID = -1
    Else
        '7179
        If blnCrossCustCombine Then
            If Len(rs("MainCustId") & "") > 0 Then
                lngCustID = rs("MainCustId")
            Else
                lngCustID = rs("CustID")
            End If
        Else
            lngCustID = rs("CustID")
        End If
    End If
    GetfFailCustID = lngCustID
    rs.Close
    Set rs = Nothing
    Exit Function
ChkErr:
   ErrSub Me.Name, "GetfFailCustID"
End Function

'#3697 ����H�U�s�榡�Ǧ^�Ӫ����~ By Kin 2007/12/31
Private Function GetChinatrusNewMessage(ByVal pstrMessage As String) As String
    Dim strMessage As String
    '#5052 �^�ХN�X�����վ� By Kin 2009/05/14
    strMessage = ""
    Select Case UCase(pstrMessage)
        Case ""
            strMessage = "�L�^�ХN�X"
            
        Case "12"
            strMessage = "���}�d"
            strErrCode = "12"
        Case "14"
            'strMessage = "����ڵ�-�d����"         '#5052
            strMessage = "�d�����~"
            strErrCode = "14"
        Case "30"
            strMessage = "����ڵ�-format error"    '#5052
            strErrCode = "30"
        Case "41"
            'strMessage = "�S���d"                  '#5052
            strMessage = "�S���d����ڵ�"
            strErrCode = "41"
        Case "43"
            'strMessage = "�S���d"                  '#5052
            strMessage = "�S���d����ڵ�"
            strErrCode = "43"
        Case "51"
            strMessage = "�B�פ���/�W�B"
            strErrCode = "51"
        Case "54"
            strMessage = "�d�����Ĵ������~"
            strErrCode = "54"
        Case "56"
            'strMessage = "����ڵ�-CAF"
            strMessage = "����ڵ�-�d�����`"
            strErrCode = "56"
        Case "57"
            strMessage = "����ڵ�-���}�d"
            strErrCode = "57"
        Case "61"
            strMessage = "����ڵ�-���o�d��"
            strErrCode = "61"
        Case "65"
            strMessage = "����ڵ�-���o�d��"    '#5052
            strErrCode = "65"
        Case "75"
            strMessage = "����ڵ�-�K�X���~"    '#5052
            strErrCode = "75"
        Case "83"
            strMessage = "����ڵ�- �L���d��"   '#5052
            strErrCode = "83"
        Case "86"
            strMessage = "����ڵ�-�������~"
            strErrCode = "86"
        Case "87"
            'strMessage = "����ڵ�-CVV"
            strMessage = "CVV�ˮֿ��~"          '#5052
            strErrCode = "87"
        Case "94"
            strMessage = "�������"             '#5052
            strErrCode = "94"
        Case "02"
            strMessage = "�ӷ|�o�d��"
            strErrCode = "02"
        Case "03"
            strMessage = "�ө��N�����~/�ި�ө�"
            strErrCode = "03"
        Case "06"
            strMessage = "�t�β��`"
            strErrCode = "06"
        Case "07"
            strMessage = "�S���d"
            strErrCode = "07"
        Case "N8"
            strMessage = "����ڵ�-�W�L���B"
            strErrCode = "N8"
        Case "O6"
            strMessage = "�d�����~-�d���~"
            strErrCode = "O6"
        Case "N8"
            strMessage = "����ڵ�-�W�L���B"    '#5052
            strErrCode = "N8"
        Case "P9"
            strMessage = "�O�N��魭�B/��"      '#5052
            strErrCode = "P9"
        Case "Q0"
            strMessage = "���������~"         '#5052
            strErrCode = "Q0"
        Case "T1"
            strMessage = "����ڵ�-���B���~"
            strErrCode = "T1"
        Case "T3"
            strMessage = "�ө����������d���H�Υd"
            strErrCode = "T3"
        Case "T8"
            strMessage = "�d�����~"
            strErrCode = "T8"
        Case Else
            strMessage = "����ڵ�"
            strErrCode = pstrMessage
    End Select
        
    GetChinatrusNewMessage = strMessage

End Function
Private Function GetUnionBathMessage(ByVal pstrMessage As String, ByVal aData As String) As String
  On Error Resume Next
    Dim strMessage As String
    strErrCode = UCase(pstrMessage)
    strMessage = Trim(Mid(aData, 69, 16))
    If strMessage <> "" Then
        GetUnionBathMessage = strMessage
        Exit Function
    End If
    Select Case UCase(pstrMessage)
        Case "086"
            strMessage = "VISA LINE DOWN"
        Case "087"
            strMessage = "HOST DOWN"
        Case "088"
            strMessage = "TIME OUT"
        Case "555"
            strMessage = "SYSTEM TIME_OUT"
        Case "653"
            strMessage = "TC ERROR"
        Case "650"
            strMessage = "REJECT CRD_ERR"
        Case "651"
            strMessage = "SYSTEM ERROR"
        Case "088", "101"
            strMessage = "CALL BANK"
        Case "050", "076"
            strMessage = "REJECT"
        Case "051", "651", "652"
            strMessage = "EXPIRED CARD"
        Case "675"
            strMessage = "SYSTEM REJECT"
        Case Else
            strMessage = "�䥦�T��"
    End Select
    If Val(UCase(pstrMessage)) >= 900 And Val(UCase(pstrMessage)) < 910 Then
        strMessage = "PICK UP CARD"
    End If
    GetUnionBathMessage = strMessage
End Function
'#5494 �ŷs��޻Ȧ�^�Ъ��N�X By Kin 2010/02/02
Private Function GetNewWebMessage(ByVal pstrMessage As String) As String
  On Error Resume Next
  Dim strMessage As String
  strMessage = Empty
  strErrCode = UCase(pstrMessage)
  Select Case UCase(pstrMessage)
    Case "00"
        strMessage = "OK"
    Case "01", "02"
        strMessage = "�Ьd�߻Ȧ�"
    Case "03"
        strMessage = "�L�Ĥ��S��"
    Case "04", "07"
        strMessage = "�S���d"
    Case "05"
        strMessage = "�ڵ����"
    Case "06"
        strMessage = "���~"
    Case "12"
        strMessage = "�L�ĥ��"
    Case "13"
        strMessage = "���ŵo�d��]�w�����B"
    Case "14"
        strMessage = "�d�����s�b"
    Case "15"
        strMessage = "�o�d�椣�s�b"
    Case "19"
        strMessage = "�A��J��������A"
    Case "21"
        strMessage = "�L�k�N������������"
    Case "25"
        strMessage = "�n����VISA �����t�Τ���ƵL�k����"
    Case "28"
        strMessage = "�d�ߩ�n����VISA�����t�Τ���ƵL�k����"
    Case "30"
        strMessage = "�榡���~"
    Case "33", "54"
        strMessage = "�L���d"
    Case "41"
        strMessage = "�򥢥d"
    Case "43"
        strMessage = "���ѥd�A���H�S��"
    Case "51"
        strMessage = "�i�ξl�B����"
    Case "55"
        strMessage = "�����T"
    Case "56", "57"
        strMessage = "�o�d�楼���ѫ��d�H�Ӷ�������\��"
    Case "61"
        strMessage = "�W�L�o�d��]�w��VISA�t�ά����ɤ����B"
    Case "62"
        strMessage = "�ӥd���o�d��w����"
    Case "65"
        strMessage = "�W�L�o�d��]�w��VISA�t�ά����ɤ�����"
    Case "75"
        strMessage = "PIN ��J���~�W�L����"
    Case "80"
        strMessage = "����������"
    Case "81"
        strMessage = "PIN �K�X��CVV �Ȥ����T"
    Case "82"
        strMessage = "CVV �����T"
    Case "85"
        strMessage = "�g�d�ҫ�,�d���õL��"
    Case "91"
        strMessage = "�o�d��L�k�����γB�z�ӰT��"
    Case "93"
        strMessage = "����L�k����,�����\�����"
    Case "94"
        strMessage = "���Х��"
    Case "96"
        strMessage = "�t�Υ\�ಧ�`"
    Case "99"
        strMessage = "�u���c��"
    Case "IE"
        strMessage = "ID���~"
    Case Else
        strMessage = "�����T���^���X !!"
  End Select
  GetNewWebMessage = strMessage
End Function


Private Function GetHSBCResponseMessage(ByVal pstrMessage As String) As String

    Dim strMessage As String

    strMessage = ""
    Select Case pstrMessage
        Case "00"
            strMessage = "APPROVAL xxxxxx 00  APPROVED or COMPLETE SUCCESSFULLY(�֭�)"
            strErrCode = "00"
        Case "01"
            strMessage = "CALL    01  REFER TO CARD ISSUER(�Ьd�߻Ȧ�)"
            strErrCode = "01"
        Case "03"
            strMessage = "INVALID MERCHANT    03  INVALID MERCHANT(�ө��N�����~)"
            strErrCode = "03"
        Case "04"
            strMessage = "'*HOLD CARD  04  PICK UP CARD(�S���d�D�лP�Ȧ��p��)"
            strErrCode = "04"
        Case "05"
            strMessage = "DO NOT HONOR    05  DO NOT HONOR(�W�B�ΰө��Ȯɰ��d)"
            strErrCode = "05"
        Case "12"
            strMessage = "SERV NOT ALLOWED    12  INVALID TRANSACTION(�o�d���_�u�Ψt�γ]�w���~�D�йq�ܱ��v)"
            strErrCode = "12"
        Case "13"
            strMessage = "AMOUNT ERROR    13  INVALID AMOUNT(���B���~)"
            strErrCode = "13"
        Case "14"
            strMessage = "CARD NO. ERROR  14  INVALID CARD NUMBER(NO SUCH NUMBER) (�d�����~)"
            strErrCode = "14"
        Case "19"
            strMessage = "RE ENTER    19  RE-ENTER TRANSACTION(�@�~�O��)"
            strErrCode = "19"
        Case "41"
            strMessage = "*HOLD-CALL  41  PICK UP CARD(LOSE CARD) (�򥢥d)"
            strErrCode = "41"
        Case "43"
            strMessage = "*HOLD CALL  43  PICK UP CARD(STOLEN CARD) (�򥢥d�Υ��ѥd)"
            strErrCode = "43"
        Case "51"
            strMessage = "DECLINE 51  NOT SUFFICIFENT FUNDS(�W�B)"
            strErrCode = "51"
        Case "91"
            strMessage = "NO REPLY    91  SWITCH OR ISSUER SYSTEM INOPERATIVE(���}�d)"
            strErrCode = "91"
        Case "54"
            strMessage = "*EXPIRED CARD   54  EXPIRED CARD(�L���d)"
            strErrCode = "54"
        Case Else
            strMessage = "�����T���^���X !!"
        End Select

    GetHSBCResponseMessage = strMessage

End Function

