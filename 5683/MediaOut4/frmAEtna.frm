VERSION 5.00
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAEtna 
   ClientHeight    =   6810
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11190
   Icon            =   "frmAEtna.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6810
   ScaleWidth      =   11190
   StartUpPosition =   1  '���ݵ�������
   Begin MSComDlg.CommonDialog cnd1 
      Left            =   10530
      Top             =   6240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "��X��"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   5745
      Left            =   240
      TabIndex        =   11
      Top             =   180
      Width           =   10695
      Begin VB.Frame frmData 
         Caption         =   "����ɦ�m"
         Height          =   1875
         Left            =   450
         TabIndex        =   12
         Top             =   1680
         Width           =   9075
         Begin VB.CommandButton cmdErrPath 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   14.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   7530
            TabIndex        =   24
            Top             =   1320
            Width           =   375
         End
         Begin VB.CommandButton cmdMemoPath 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   14.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   7530
            TabIndex        =   23
            Top             =   810
            Width           =   375
         End
         Begin VB.CommandButton cmdDataPath 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   14.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   7530
            TabIndex        =   22
            Top             =   270
            Width           =   375
         End
         Begin VB.TextBox txtDataPath 
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3240
            TabIndex        =   4
            ToolTipText     =   "�п�J�r�ɤ����|���ɦW�I"
            Top             =   270
            Width           =   4095
         End
         Begin VB.TextBox txtMemoPath 
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3240
            TabIndex        =   5
            Top             =   795
            Width           =   4095
         End
         Begin VB.TextBox txtErrPath 
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3240
            TabIndex        =   6
            Top             =   1320
            Width           =   4095
         End
         Begin VB.Label lblDatapath 
            AutoSize        =   -1  'True
            Caption         =   "����ɦ�m�]���| + �W�١^"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   330
            TabIndex        =   15
            Top             =   390
            Width           =   2340
         End
         Begin VB.Label lblMemopath 
            AutoSize        =   -1  'True
            Caption         =   "�Ƶ��ɦ�m�]���| + �W�١^"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   330
            TabIndex        =   14
            Top             =   885
            Width           =   2340
         End
         Begin VB.Label lblErrpath 
            AutoSize        =   -1  'True
            Caption         =   "���D�Ѧ��ɦ�m�]���| + �W�١^"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   330
            TabIndex        =   13
            Top             =   1380
            Width           =   2730
         End
      End
      Begin VB.TextBox txtMemo1 
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1755
         Left            =   1800
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   3690
         Width           =   3255
      End
      Begin VB.TextBox txtMemo2 
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1755
         Left            =   6510
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   3720
         Width           =   3015
      End
      Begin VB.TextBox txtInVoice 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6660
         MaxLength       =   8
         TabIndex        =   3
         Top             =   1050
         Width           =   1125
      End
      Begin Gi_Date.GiDate gdaTransfer 
         Height          =   375
         Left            =   2280
         TabIndex        =   2
         Top             =   1050
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
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
      Begin Gi_Date.GiDate gdaCredit 
         Height          =   375
         Left            =   6660
         TabIndex        =   1
         Top             =   480
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
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
      Begin Gi_Date.GiDate gdaAgency 
         Height          =   375
         Left            =   2280
         TabIndex        =   0
         Top             =   480
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
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
      Begin VB.Label lblagency 
         AutoSize        =   -1  'True
         Caption         =   "�{�d�N��ú�O����"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   540
         TabIndex        =   21
         Top             =   570
         Width           =   1560
      End
      Begin VB.Label lblCredit 
         AutoSize        =   -1  'True
         Caption         =   "�H�Υd����ú�O����"
         Height          =   180
         Left            =   4890
         TabIndex        =   20
         Top             =   570
         Width           =   1620
      End
      Begin VB.Label lblTransfer 
         AutoSize        =   -1  'True
         Caption         =   "�۰���b���ڤ��"
         Height          =   180
         Left            =   570
         TabIndex        =   19
         Top             =   1170
         Width           =   1440
      End
      Begin VB.Label lblMemo1 
         AutoSize        =   -1  'True
         Caption         =   "���O��Ƶ��@�G"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   480
         TabIndex        =   18
         Top             =   3750
         Width           =   1365
      End
      Begin VB.Label lblMemo2 
         AutoSize        =   -1  'True
         Caption         =   "���O��Ƶ��G�G"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5160
         TabIndex        =   17
         Top             =   3750
         Width           =   1365
      End
      Begin VB.Label lblInVoice 
         AutoSize        =   -1  'True
         Caption         =   "�Τ@�s��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4890
         TabIndex        =   16
         Top             =   1170
         Width           =   780
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "F2.�}�l"
      Height          =   405
      Left            =   3570
      TabIndex        =   9
      Top             =   6180
      Width           =   1275
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "����(&X)"
      Height          =   405
      Left            =   5940
      TabIndex        =   10
      Top             =   6180
      Width           =   1275
   End
End
Attribute VB_Name = "frmAEtna"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private cnn As New ADODB.Connection
Private rsDefTmp As New ADODB.Recordset

'�J�b�b��
Private strActNo As String
'�Ʒ~���N��
Private strBankId As String
'�Ʒ~���W��
Private strBankName As String
'������N��
Private strCorpID As String
'Where ����
Private strChoose As String
'�{���W��
Private strPrgName As String
'���q�O
Private strCompCode As String
'�A�����O
Private strServiceType As String

'�H�U�OWriteLine �b�Ϊ�

'GICMIS1.INI���|
Private strPath As String
'ErrLog ���|
Private strErrPath As String
'�ɮצW��
Private strDataName As String
Private strMemoName As String
Private strErrName As String
Private strTrType As String
'��b�էO
Private strStopDate As String  '��������I���
'�C��h�b��B�z (0=�_ , 1=�O)
Private strMedia As String
Private strChoose33 As String
Private strGetOwner As String       'OwnerName
Private strFlowId As String         '�y�{����

Private Sub InitData()
  On Error GoTo ChkErr
    With objStorePara
        strActNo = .ActNo
        Set gcnGi = .Connection
        Set cnn = .MDBConn
        strBankId = .BankId
        strBankName = .BankName
        strCorpID = .CorpID
        strChoose = .ChooseStr
        strPath = .Inipath
        strErrPath = .ErrPath
        strCompCode = .uCompCode
        strServiceType = .uServiceType
        strStopDate = .uStopDate
        strChoose33 = .uChoose33
        strGetOwner = .uGetOwner
        strFlowId = .FlowId
        '�w�q�O����recordset
        Call DefineRs
    End With
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "InitData")
End Sub

Private Function BeginTran() As Boolean
  Dim strSQL As String
  Dim strOldBillNo As String
  Dim lngTime As Long
  Dim lngErrCount As Long
  Dim lngCount As Long
  Dim lngShouldAmt  As Long
  Dim lngTotAmt As Long
  Dim strBillNo As String, strBillNo_Old As String
  Dim rsTmp As New ADODB.Recordset
  Dim rsBank As New ADODB.Recordset
  Dim rsPara As New ADODB.Recordset
  Dim rsUpd As New ADODB.Recordset
  Dim rsCD013 As New ADODB.Recordset
  Dim StrTableName As String
  Dim strWhere As String
  Dim strSQLA As String
  Dim strUcCode As String
  Dim strUcName As String
  Dim blnUpdUcCode As Boolean
  Dim strUPDUCCode As String
  Dim strCD013 As String
  Dim strUpdTime As String
  Dim strErrMsg As String
  Dim strCUSTID As String
  On Error Resume Next
     cnn.BeginTrans
     cnn.Execute ("Delete From SO3271A")
  On Error GoTo ChkErr
    
    BeginTran = False
    lngTime = Timer
    lngErrCount = 0
    strUpdTime = Format(RightNow, "EE/MM/DD HH:MM:SS")
    
    
    '********************************************************************************************************************************************************
    '#3417 �q�l�ɶץX��,�n��J������](RefNo=4) By Kin 2007/12/04
    If Not GetRS(rsCD013, "Select * From " & strGetOwner & "CD013 Where RefNo=4 And StopFlag<>1 Order By CodeNo Desc", gcnGi, adUseClient, adOpenKeyset) Then Exit Function
    If rsCD013.EOF Then
        blnUpdUcCode = False
    Else
        blnUpdUcCode = True
        strUcCode = rsCD013("CodeNo") & ""
        strUcName = rsCD013("Description") & ""
    End If
    '*********************************************************************************************************************************************************
    
    strSQL = "Select Para24 From " & strGetOwner & "SO043 Where Rownum=1"
    If Not GetRS(rsPara, strSQL, gcnGi) Then Exit Function
    strMedia = GetFieldValue(rsPara, "Para24") & ""
    
    ' 2003/3/26 Crystal Edit
    If InStr(1, strChoose, "C.") Then
        StrTableName = "," & strGetOwner & "SO002 C"
        strWhere = strWhere & " And A.CustId = C.CustId And A.Servicetype = C.ServiceType "
    End If
    If InStr(1, strChoose, "B.") Then
        StrTableName = StrTableName & "," & strGetOwner & "SO001 B"
        strWhere = strWhere & " And A.CustId = B.CustId"
    End If
    '******************************************************************
    '#5055 �p�GCD013.REFNO=3(�d�x�w��)�h���A�X�b By Kin 2009/04/28
    '#5218 �d�x�w���n��Nvl���覡 By Kin 2009/08/05
    '#5564 �W�[�Ѧ�7,8�N��w��,PayOK,�]�N��O�w�� By Kin 2010/05/17
    StrTableName = StrTableName & "," & strGetOwner & "CD013"
    strWhere = strWhere & " And A.CancelFlag<> 1 And UCCode Is Not Null And A.UCCode=CD013.CodeNo " & _
            " And Nvl(CD013.REFNO,0) NOT IN(3,7) AND NVL(CD013.PAYOK,0)=0 "
    '******************************************************************
    If strWhere <> "" Then strWhere = Mid(strWhere, 5)
    '#5683 �W�[ú�����O,�ìD��̤p�������I��� By Kin 2010/08/03
    If strMedia = 0 Then
        strSQL = "SELECT " & GetUseIndexStr("SO033", "ShouldDate") & " A.BillNo,Sum(A.ShouldAmt) ShouldAmt, " & _
                " A.AccountNo,A.CustId,A.MediaBillNo,Max(Nvl(A.PayKind,0)) PAYKIND, " & _
                strMinRealStopDate & _
                 " From " & strGetOwner & "SO033 A " & StrTableName & _
                 " Where " & strWhere & _
                 IIf(Len(strWhere) = 0, "", " And ") & strChoose & " Group By A.BillNo,A.AccountNo,A.CustId,A.MediaBillNo "
        If Not GetRS(rsTmp, strSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
        If Not BatchUpdateMediaBillNo(rsTmp, strChoose33, gcnGi) Then Exit Function
    Else
        '#5055 �p�GCD013.REFNO=3(�d�x�w��)��UPD MediaBillNo By Kin 2009/04/28
        '#5218 �d�x�w���n��Nvl���覡 By Kin 2009/08/05
        '#5564 �W�[�Ѧ�7,8�N��w��,PayOK,�]�N��O�w�� By Kin 2010/05/17
        strSQLA = "SELECT " & GetUseIndexStr("SO033", "ShouldDate") & "A.BillNo ,A.MediaBillNo,A.AccountNo,A.CustId " & _
                 "From " & strGetOwner & "SO033 A,CD013 " & " Where " & strChoose33 & " And A.MediaBillNo Is Null" & _
                 " And A.UCCode=CD013.CodeNo And Nvl(CD013.REFNO,0) NOT IN(3,7)" & _
                 " AND NVL(CD013.PAYOK,0) = 0 "

        '#5683 �W�[ú�����O,�ìD��̤p�������I��� By Kin 2010/08/03
        strSQL = "SELECT " & GetUseIndexStr("SO033", "ShouldDate") & " A.MediaBillNo BillNo,Sum(A.ShouldAmt) ShouldAmt," & _
                " A.AccountNo,Max(NVL(A.PayKind,0)) PAYKIND, " & _
                strMinRealStopDate & _
                 " From " & strGetOwner & "SO033 A " & StrTableName & _
                 " Where " & strWhere & _
                 IIf(Len(strWhere) = 0, "", " And ") & strChoose & " Group By A.AccountNo,A.MediaBillNo,A.PayKind "
                
        If Not GetRS(rsUpd, strSQLA, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
        '�妸��sMediaBillNo���
        If Not BatchUpdateMediaBillNo(rsUpd, strChoose33, gcnGi) Then Exit Function
        If Not GetRS(rsTmp, strSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
    End If
    
    '**************************************************************************************************
    '#3417 ��sUCCode�PUCName���y�k By Kin 2007/12/05
    '#4388 �@�ֲ��ʲ��ʤH���P���ʮɶ� By Kin 2009/04/29
'    strUPDUCCode = "Select A.UCCode,A.UCName,A.UpdEn,A.UpdTime," & IIf(strMedia = "0", "BILLNO", "MEDIABILLNO") & " BILLNO From " & _
'                strGetOwner & "SO033 A" & StrTableName & _
'              " Where " & strWhere & _
'              IIf(Len(strWhere) = 0, "", " And ") & strChoose
    strUPDUCCode = GetUpdateSQL(Val(strMedia & ""), StrTableName, strWhere, strChoose)
    '**************************************************************************************************

    Set rsTmp.ActiveConnection = Nothing
    
'    If Not GetRS(rsTmp, strSQL, gcnGi) Then Exit Function
    strSQL = "Select Class From " & strGetOwner & "CD018 Where CodeNo ='" & strBankId & "'"
    If Not GetRS(rsBank, strSQL, gcnGi) Then Exit Function
    
    
    If rsTmp.RecordCount > 0 Then
       rsTmp.MoveFirst
    End If
    
    If rsTmp.BOF Or rsTmp.EOF Then
       MsgBox "�L��ơI�Э��s�ާ@�I", vbInformation, "�T���I"
       blnNodata = True
       Exit Function
    Else
        blnNodata = False
        strOldBillNo = "-1"
        
        '��b�էO
        strTrType = "CM" & GetString(GetFieldValue(rsBank, "Class") & "", 3) & GetString(Mid(strActNo, 6, 6), 6)
        While Not rsTmp.EOF
           '�p���h���PBillNo �����
           '#5683 �p�G�O���]�t�w�I��n�ˬd���O�I��� By Kin 2010/08/03
'           If rsTmp("PayKind") = 0 Then
'                If Val(Format(rsTmp("REALSTOPDATE") & "", "YYYYMMDD")) < Val(gdaTransfer.GetValue & "") Then
'                    'LOG Format ���y�Ƚs�z�B�y���O���ئW�١z�B�y���O�_�㨴����z�B�yú�I���O�z
'                     strCUSTID = gcnGi.Execute("SELECT CUSTID FROM " & strGetOwner & "SO033 WHERE " & _
'                            IIf(strMedia = "0", "BILLNO = '", "MEDIABILLNO = '") & rsTmp("BILLNO") & "' AND ROWNUM =1 ")(0)
'                     strErrMsg = "�Ƚs�G" & strCUSTID & "�B��ڽs���G" & rsTmp("BILLNO") & _
'                            "�B���~�����I���G" & rsTmp("REALSTOPDATE") & "�Bú�I���O: �w�I�� """
'                      WriteLineData strErrMsg, 2
'                      lngErrCount = lngErrCount + 1
'                      GoTo lNextRcd
'                End If
'           End If
            If Not IsPayKindOK(rsTmp, gdaTransfer.GetValue & "", giAll, Val(strMedia)) Then
                Call WriteSO3271Err(GetPayKindCustId(rsTmp, Val(strMedia)), rsTmp("BillNo"), rsTmp("REALSTOPDATE"))
                lngErrCount = lngErrCount + 1
                GoTo lNextRcd
            End If
            If strMedia = 0 Then
                lngShouldAmt = Val(GetRsValue("Select Sum(A.ShouldAmt) From " & strGetOwner & "So033 A Where A.BillNo = '" & rsTmp("BillNo") & "' And A.CancelFlag = 0 And A.UCCode Is Not Null AND " & strChoose33, gcnGi) & "")
            Else
                lngShouldAmt = Val(GetRsValue("Select Sum(A.ShouldAmt) From " & strGetOwner & "So033 A Where A.MediaBillNo = '" & rsTmp("BillNo") & "' And A.AccountNo='" & GetFieldValue(rsTmp, "AccountNo") & "' And A.CancelFlag = 0 And A.UCCode Is Not Null AND " & strChoose33, gcnGi) & "")
            End If
            If rsTmp("BillNo") = strOldBillNo & "" Then
              'lngShouldAmt = lngShouldAmt + GetFieldValue(rsTmp, "ShouldAmt") + 0
                GoTo lNextRcd
            End If
                
            If strMedia = 1 Then
                strBillNo = GetFieldValue(rsTmp, "BillNo") & ""
                strBillNo_Old = GetRsValue("Select BillNo From SO033 Where MediaBillNO='" & GetFieldValue(rsTmp, "BillNo") & "'", gcnGi) & ""
            Else
                strBillNo = GetBillNo_New(GetFieldValue(rsTmp, "BillNo") & "")
                strBillNo_Old = GetBillNo_Old(rsTmp("BillNo") & "")
            End If
                                                                  
           '�p�⦩�b�`���B
            lngCount = lngCount + 1
            lngTotAmt = lngTotAmt + lngShouldAmt
            With rsDefTmp
                 .AddNew
                 .Fields("Code") = "2"
                 '�ϧO�X 1-1
                 .Fields("TransType") = strTrType
                 '��b�էO 2-12
                 .Fields("TransDate") = Replace(gdaTransfer.GetOriginalValue, "/", "")
                 '��b��� 13-19
                 .Fields("AccountNo") = Mid(GetFieldValue(rsTmp, "AccountNo") & "", 1, 14)
                 '���b�b�� 20-33
                 .Fields("Amt") = lngShouldAmt
                 '���b���B 34-47
                 .Fields("Type") = ""
                 '�J���b�X 48-48
                 .Fields("TransStatus") = "99"
                 '��b���p 49-50
                 .Fields("Note") = strBillNo
                 '�O�d���(11�X�s��ڽs��+3�Ӫť�) 51-64
                 
                 '#3417 �p�GCD013.RefNo=4 ���ȫh��sUCCode�PUCName By Kin 2007/12/4
                 'Insert Into MDB(SO3271A)
            End With
             

              
'            strUPDUCCode = "UPDATE " & strGetOwner & "SO033 SET " & _
'                    " UCCODE = " & strUcCode & ",UCNAME = '" & strUcName & "'," & _
'                    " UPDEN = '" & objStorePara.uUpdEn & "'," & _
'                    " UPDTIME = '" & strUpdTime & "' " & _
'                    " WHERE " & IIf(strMedia = "0", " BILLNO = ", " MEDIABILLNO = ") & "='" & rsTmp("BILLNO") & "'"
                    
            
            cnn.Execute Replace("Insert Into SO3271A (AccountNo, TransDate, ShouldAmt, BillNo_Old, BillNo_New) Values (" & _
                 GetNullString(rsTmp("AccountNo")) & "," & GetNullString(gdaTransfer.GetValue(True), giDateV, giAccessDb) & "," & _
                 GetNullString(lngShouldAmt) & "," & GetNullString(strBillNo_Old) & "," & _
                 GetNullString(strBillNo) & ")", Chr(0), "")
           
lNextRcd:
           strOldBillNo = rsTmp("BillNo")
           rsTmp.MoveNext
           DoEvents
        Wend
        
        '*************************************************************************************************************************
        '#3417 ��sUCCode�PUCName By Kin 2007/12/04
        '#4388 �W�[��s���ʮɶ��P���ʤH�� By Kin 2009/04/29
        '#5683 �n�L�o��PAYKIND=0 AND REALSTOPDATE<�e������ By Kin 2010/08/03
        If blnUpdUcCode Then
            Dim rsUpdUCCode As New ADODB.Recordset
'            Dim rsChk As New ADODB.Recordset
'            Dim aQry As String
            
            
            If Not GetRS(rsUpdUCCode, strUPDUCCode, gcnGi, adUseClient, adOpenKeyset, adLockPessimistic) Then Exit Function
            
            
            Do While Not rsUpdUCCode.EOF

'                aQry = "SELECT MIN(NVL(PAYKIND,0)) PAYKIND, " & _
'                        strMinRealStopDate & _
'                        " FROM " & strGetOwner & "SO033 WHERE " & _
'                        IIf(strMedia = "0", " BILLNO='", " MediabillNo='") & rsUpdUCCode("BILLNO") & "'"
'                If Not GetRS(rsChk, aQry, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
                If IsPayKindOK(rsUpdUCCode, gdaTransfer.GetValue & "", giSingle, Val(strMedia)) Then
                    rsUpdUCCode("UPDTIME") = strUpdTime
                    rsUpdUCCode("UPDEN") = objStorePara.uUpdEn
                    rsUpdUCCode("UCcode") = strUcCode
                    rsUpdUCCode("UCName") = strUcName
                    rsUpdUCCode.Update
                End If
'                If Val(rsChk("PAYKIND") & "") = 0 Then
'                    If Val(Format(rsChk("REALSTOPDATE") & "", "YYYYMMDD")) > Val(gdaTransfer.GetValue & "") Then
'                        rsUpdUCCode("UPDTIME") = strUpdTime
'                        rsUpdUCCode("UPDEN") = objStorePara.uUpdEn
'                        rsUpdUCCode("UCcode") = strUcCode
'                        rsUpdUCCode("UCName") = strUcName
'                        rsUpdUCCode.Update
'                    End If
'                Else
'                    rsUpdUCCode("UPDTIME") = strUpdTime
'                    rsUpdUCCode("UPDEN") = objStorePara.uUpdEn
'                    rsUpdUCCode("UCcode") = strUcCode
'                    rsUpdUCCode("UCName") = strUcName
'                    rsUpdUCCode.Update
'                End If
                
                rsUpdUCCode.MoveNext
            Loop
            CloseRecordset rsUpdUCCode
        End If
        '*************************************************************************************************************************
        cnn.CommitTrans
        
        
        
        On Error Resume Next
        rsTmp.Close
        Set rsTmp = Nothing
        rsBank.Close
        Set rsBank = Nothing
        If strFlowId = 0 Then
            rsPara.Close
            Set rsPara = Nothing
        End If
        If Not subWriteLine(lngCount, lngTotAmt) Then
            msgResult lngCount, lngErrCount, lngTime      '��ܰ��浲�G
            CloseFS
            Exit Function      '�Ȥ᤺�e
        End If
        WriteLineData txtMemo1 & vbCrLf & txtMemo2, 1  '�Ƶ���
        msgResult lngCount, lngErrCount, lngTime      '��ܰ��浲�G
        BeginTran = True
        CloseFS
     End If
     CloseRecordset rsCD013
Exit Function
ChkErr:
    CloseFS
    Call ErrSub(Me.Name, "BeginTran")
End Function

Private Function subWriteLine(ByVal lngCount As Long, ByVal lngTotAmt As Long) As Boolean
  On Error GoTo ChkErr
    Dim varData As Variant
    Dim strData As String
    Dim lngloop As Long
    
    subWriteLine = False
    With rsDefTmp
         If .BOF Or .EOF Then Exit Function
         '����
         WriteLineData 1 & strTrType & GetString(Replace(gdaTransfer.GetOriginalValue, "/", ""), 7, GIRIGHT, True) & 1 & GetString(strActNo, 14), 0
         .MoveFirst
         varData = .GetRows
         '���e
         For lngloop = 0 To .RecordCount - 1
            WriteLineData GetString(varData(0, lngloop), 1) & GetString(varData(1, lngloop), 11) & _
                GetString(varData(2, lngloop), 7, GIRIGHT, True) & GetString(varData(3, lngloop), 14, GIRIGHT, True) & _
                GetString(varData(4, lngloop), 12, GIRIGHT, True) & "00" & GetString(varData(5, lngloop), 1) & _
                GetString(varData(6, lngloop), 2) & GetString(varData(7, lngloop), 11), 0
             DoEvents
         Next
         '����
         WriteLineData 3 & strTrType & GetString(Replace(gdaTransfer.GetOriginalValue, "/", ""), 7, GIRIGHT, True) & GetString(lngCount, 5, GIRIGHT, True) & GetString(lngTotAmt, 12, GIRIGHT, True) & "00", 0
    End With
    subWriteLine = True
    On Error Resume Next
    rsDefTmp.Close
    Set rsDefTmp = Nothing
    Exit Function
ChkErr:
    subWriteLine = False
    Call ErrSub(Me.Name, "subWriteLine")
End Function

'�w�q�O����Recordset
Private Sub DefineRs()
    On Error GoTo ChkErr
    With rsDefTmp
        If .State = adStateOpen Then .Close
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        
        .Fields.Append ("Code"), adBSTR, 1, adFldIsNullable         '�ϧO�X
        .Fields.Append ("TransType"), adBSTR, 11, adFldIsNullable   '��b�էO
        .Fields.Append ("TransDate"), adBSTR, 7, adFldIsNullable    '��b���
        .Fields.Append ("AccountNo"), adBSTR, 14, adFldIsNullable   '�}��H�b��
        .Fields.Append ("Amt"), adBSTR, 12, adFldIsNullable         '��b���B
        .Fields.Append ("Type"), adBSTR, 1, adFldIsNullable         '�J���b�X
        .Fields.Append ("TransStatus"), adBSTR, 2, adFldIsNullable  '��b���p
        .Fields.Append ("Note"), adBSTR, 11, adFldIsNullable        '�O�d���(11�X�s��ڽs��+3�Ӫť�)
        .Open
    End With
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "DefineRs")
End Sub

Private Sub cmdCancel_Click()
    On Error Resume Next
    Unload Me
End Sub

Private Sub cmdDataPath_Click()
    On Error GoTo ChkErr
        With cnd1
                .FileName = txtDataPath.Text
                .Filter = "��r��|*.txt"
                .InitDir = strErrPath
                .Action = 1
                txtDataPath = .FileName
        End With
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "cmdDataPath_Click")
End Sub

Private Sub cmdErrPath_Click()
    On Error GoTo ChkErr
        With cnd1
                .FileName = txtErrPath.Text
                .Filter = "��r��|*.txt"
                .InitDir = strErrPath
                .Action = 1
                txtErrPath = .FileName
        End With
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "cmdErrPath_Click")
End Sub

Private Sub cmdMemoPath_Click()
    On Error GoTo ChkErr
        With cnd1
                .FileName = txtMemoPath.Text
                .Filter = "��r��|*.txt"
                .InitDir = strErrPath
                .Action = 1
                txtMemoPath = .FileName
        End With
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "cmdMemoPath_Click")
End Sub

Private Sub cmdOK_Click()
    On Error GoTo ChkErr
    
    If Not ChkDTok Then Exit Sub
    If Not IsDataOk Then Exit Sub
    If OpenFile(txtDataPath, 0, True) = False Then txtDataPath.SetFocus: Exit Sub
    If OpenFile(txtMemoPath, 1, True) = False Then txtMemoPath.SetFocus: Exit Sub
    If OpenFile(txtErrPath, 2, True) = False Then txtErrPath.SetFocus: Exit Sub
    Screen.MousePointer = vbHourglass
    objStorePara.uProcText = txtDataPath.Text
    objStorePara.uProcErrText = txtErrPath.Text
    Call ScrToRcd         '�N�e����J����g��Log��
    If Not BeginTran Then
       If Not blnNodata Then
            objStorePara.uUpdate = False
            MsgBox "��b��ƿ�X����!", vbExclamation, "ĵ�i..."
            CloseFS
            Screen.MousePointer = vbDefault
            Unload Me
            Exit Sub
       Else
            objStorePara.uNoData = True
            CloseFS
            Screen.MousePointer = vbDefault
            Unload Me
            Exit Sub
       End If
    End If


    objStorePara.uUpdate = True
    Screen.MousePointer = vbDefault
    Unload Me
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "cmdOK_Click")
End Sub

Private Function IsDataOk() As Boolean
On Error GoTo ChkErr
Dim strErrMsg As String

        If gdaTransfer.GetValue = "" Then strErrMsg = "�۰���b���ڤ��": gdaTransfer.SetFocus: GoTo Warning
        If txtDataPath.Text = "" Then strErrMsg = "����ɦ�m": txtDataPath.SetFocus: GoTo Warning
        If txtMemoPath.Text = "" Then strErrMsg = "�Ƶ��ɦ�m": txtMemoPath.SetFocus: GoTo Warning
        If txtErrPath.Text = "" Then strErrMsg = "���D�Ѧ��ɦ�m": txtErrPath.SetFocus: GoTo Warning
        
        strDataName = Mid(txtDataPath.Text, InStrRev(txtDataPath.Text, "\") + 1)
        strMemoName = Mid(txtMemoPath.Text, InStrRev(txtMemoPath.Text, "\") + 1)
        strErrName = Mid(txtErrPath.Text, InStrRev(txtErrPath.Text, "\") + 1)
        If (strDataName = strMemoName) Or (strMemoName = strErrName) Or (strErrName = strDataName) Then
           MsgBox "�ɮצW�٤��i����!", vbExclamation, "ĵ�i"
           Exit Function
        End If
        
    IsDataOk = True
    Exit Function
Warning:
    MsgBox strErrMsg & "  " & gMsgIsDataOK, vbExclamation, "�T���I"
    Exit Function
ChkErr:
    Call ErrSub(Me.Name, "IsDataOk")
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   On Error GoTo ChkErr
    Select Case KeyCode
             Case vbKeyEscape
                     Unload Me
             Case vbKeyF2
                     If cmdOK.Enabled = True Then cmdOK.Value = True
    End Select
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "Form_KeyDown")
End Sub

Private Sub Form_Load()
    On Error Resume Next
      Me.Caption = objStorePara.BankName & ""
      Call InitData
      RcdToScr
End Sub

Private Sub RcdToScr()
    On Error GoTo ChkErr
    Dim LogFile As TextStream
    Dim rsInv As New ADODB.Recordset
    Dim strSQL As String
    
        strSQL = "Select InvoiceId From " & strGetOwner & "SO041 Where CompCode =" & strCompCode
        If Not GetRS(rsInv, strSQL, gcnGi) Then Exit Sub
        txtInVoice.Text = GetFieldValue(rsInv, "InvoiceId") & ""
        
        rsInv.Close
        Set rsInv = Nothing

        If Dir(strErrPath & "\" & strPrgName & "Out.log") = "" Then Exit Sub
            Set LogFile = fso.OpenTextFile(strErrPath & "\" & strPrgName & "Out.log", ForReading)
            With LogFile
                    If Not .AtEndOfStream Then gdaAgency.Text = .ReadLine & ""
                       '�{�d�N��ú�O����
                    If Not .AtEndOfStream Then gdaCredit.Text = .ReadLine & ""
                        '�H�Υd����ú�O����
                    If Not .AtEndOfStream Then gdaTransfer.Text = .ReadLine & ""
                        '�۰���bú�O����
                    If Not .AtEndOfStream Then txtDataPath = .ReadLine & ""
                        '�����
                    If Not .AtEndOfStream Then txtErrPath = .ReadLine & ""
                        '���D��
                    If Not .AtEndOfStream Then txtMemoPath = .ReadLine & ""
                        '�Ƶ���
                    If Not .AtEndOfStream Then txtMemo1 = .ReadLine & ""
                        '�Ƶ�1
                    If Not .AtEndOfStream Then txtMemo2 = .ReadLine & ""
                        '�Ƶ�2
            End With
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "RcdToScr")
End Sub

Private Sub ScrToRcd()
    On Error GoTo ChkErr
    Dim LogFile As TextStream
    
        Set LogFile = fso.CreateTextFile(strErrPath & "\" & strPrgName & "Out.log", True)
         With LogFile
               Call .WriteLine(gdaAgency.Text)       '�{�d�N��ú�O����
                .WriteLine (gdaCredit.Text)       '�H�Υd����ú�O����
                .WriteLine (gdaTransfer.Text)    '�۰���bú�O����
                .WriteLine (txtDataPath)            '�����
                .WriteLine (txtErrPath)              '���D��
                .WriteLine (txtMemoPath)          '�Ƶ���
                .WriteLine (txtMemo1)               '�Ƶ�1
                .WriteLine (txtMemo2)               '�Ƶ�2
        End With
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "ScrToRcd")
End Sub

Public Property Get PrgName() As String
    PrgName = strPrgName
End Property

Public Property Let PrgName(ByVal vData As String)
    strPrgName = vData
End Property
