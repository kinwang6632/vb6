VERSION 5.00
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmChangWa 
   ClientHeight    =   6810
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11190
   Icon            =   "frmChangWa.frx":0000
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
      TabIndex        =   13
      Top             =   180
      Width           =   10695
      Begin VB.TextBox txtGotSecId 
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5625
         MaxLength       =   8
         TabIndex        =   2
         Top             =   900
         Width           =   1830
      End
      Begin VB.TextBox txtSendSpcId 
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5625
         MaxLength       =   8
         TabIndex        =   1
         Top             =   360
         Width           =   1830
      End
      Begin VB.Frame frmData 
         Caption         =   "����ɦ�m"
         Height          =   1875
         Left            =   450
         TabIndex        =   14
         Top             =   1500
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
            TabIndex        =   8
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
            TabIndex        =   6
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
            TabIndex        =   4
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
            TabIndex        =   3
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
            TabIndex        =   7
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
            TabIndex        =   17
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
            TabIndex        =   16
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
            TabIndex        =   15
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
         TabIndex        =   9
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
         TabIndex        =   10
         Top             =   3720
         Width           =   3015
      End
      Begin Gi_Date.GiDate gdaTransfer 
         Height          =   375
         Left            =   2265
         TabIndex        =   0
         Top             =   360
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
      Begin VB.Label Label1 
         Alignment       =   1  '�a�k���
         Caption         =   "������N��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4125
         TabIndex        =   22
         Top             =   960
         Width           =   1425
      End
      Begin VB.Label Label7 
         Alignment       =   1  '�a�k���
         Caption         =   "�o����N��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3825
         TabIndex        =   21
         Top             =   435
         Width           =   1725
      End
      Begin VB.Label lblTransfer 
         AutoSize        =   -1  'True
         Caption         =   "�۰���b���ڤ��"
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
         Left            =   555
         TabIndex        =   20
         Top             =   435
         Width           =   1560
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
         TabIndex        =   19
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
         TabIndex        =   18
         Top             =   3750
         Width           =   1365
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "F2.�}�l"
      Height          =   405
      Left            =   3570
      TabIndex        =   11
      Top             =   6180
      Width           =   1275
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "����(&X)"
      Height          =   405
      Left            =   5940
      TabIndex        =   12
      Top             =   6180
      Width           =   1275
   End
End
Attribute VB_Name = "frmChangWa"
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
Private strHavingOutZero As String
Private blnHavingOutZero As Boolean
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
         '#6441 ��i�b��X�p�`�B<=0�O�_���� By Kin 2013/05/13
         strHavingOutZero = ""
         blnHavingOutZero = True
        If Not .uOutZero Then
            strHavingOutZero = " Having Sum(A.ShouldAmt)>0 "
            strHavingOutZero = ""
            blnHavingOutZero = False
        End If
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
    Dim strA As String, dblC As Double
    Dim dblB As Double, dblBsub As Double
    Dim dblTCheck As Double, dblBtmp As Double
    Dim strTxt As String
    Dim strUcCode As String
    Dim strUcName As String
    Dim blnUpdUcCode As Boolean
    Dim strUPDUCCode As String
    Dim strCD013 As String
    Dim strUpdTime As String
    Dim intCrossCustCombine  As Integer
    Dim rsUpdUCCode As New ADODB.Recordset
    dblB = 0: dblBsub = 0: dblTCheck = 0
    On Error Resume Next
        cnn.BeginTrans
        cnn.Execute ("Delete From SO3271A")
         '#7179
        If Len(sqlTmpViewName) > 0 Then
            gcnGi.Execute "Drop View " & strGetOwner & sqlTmpViewName
            sqlTmpViewName = Empty
        End If
    On Error GoTo ChkErr
    
    BeginTran = False
    lngTime = Timer
'    strUpdTime = Format(RightNow, "EE/MM/DD HH:MM:SS")
    strUpdTime = GetDTString(RightNow)
    '#5843 �W�[�ˮְѼ� By Kin 2010/11/24
    blnChkPayKind = GetChkPayKind
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
    '#7179
    isCrossCustCombine = Val(GetRsValue("Select Nvl(CrossCustCombine,0) From " & strGetOwner & "SO041", gcnGi) & "")
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
    '*************************************************************************************************
    '#5055 �p�GCD013.REFNO=3(�d�x�w��)���A�X�b By Kin 2009/04/28
    '#5218 �d�x�w���n��Nvl���覡 By Kin 2009/08/05
    '#5564 �W�[�Ѧ�7,8�N��w��,PayOK,�]�N��O�w�� By Kin 2010/05/17
    StrTableName = StrTableName & "," & strGetOwner & "CD013"
    strWhere = strWhere & " And A.CancelFlag<> 1 And " & _
            " A.UCCode Is Not Null AND A.UCCode=CD013.CodeNo And Nvl(CD013.REFNO,0) NOT IN(3,7) " & _
            " AND NVL(CD013.PAYOK,0) = 0 "
    '*************************************************************************************************
    
    If strWhere <> "" Then strWhere = Mid(strWhere, 5)
    '#5055 �p�GCD013.REFNO=3(�d�x�w��)���A�X�b By Kin 2009/04/28
    '#5218 �d�x�w���n��Nvl���覡 By Kin 2009/08/05
    '#5564 �W�[�Ѧ�7,8�N��w��,PayOK,�]�N��O�w�� By Kin 2010/05/17
    '#5683 �W�[���XRealStopDate�PPAYKIND By Kin 2010/08/04
    '#6441 ��i�b��X�p�`�B<=0�O�_���� By Kin 2013/05/13
    If strMedia = 0 Then
        strSQL = "SELECT " & GetUseIndexStr("SO033", "ShouldDate") & _
                " A.BillNo,Sum(A.ShouldAmt) ShouldAmt,A.AccountNo," & _
                " A.CustId,A.MediaBillNo,Max(Nvl(A.PAYKIND,0)) PAYKIND, " & _
                 strMinRealStopDate & _
                " From " & strGetOwner & "SO033 A " & StrTableName & _
                 " Where " & strWhere & _
                 IIf(Len(strWhere) = 0, "", " And ") & strChoose & strHavingOutZero & _
                 " Group By A.BillNo,A.AccountNo,A.CustId,A.MediaBillNo"
         strSQL = "Select * From ( " & strSQL & " ) "
        If Not blnHavingOutZero Then
            strSQL = strSQL & " Where ShouldAmt > 0 "
        End If
        If Not GetRS(rsTmp, strSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
        If Not BatchUpdateMediaBillNo(rsTmp, strChoose33, gcnGi) Then Exit Function
        
    Else
        strSQLA = "SELECT " & GetUseIndexStr("SO033", "ShouldDate") & "A.BillNo ,A.MediaBillNo,A.AccountNo,A.CustId " & _
                 "From " & strGetOwner & "SO033 A," & strGetOwner & "CD013 " & _
                 " Where " & strChoose33 & " And A.MediaBillNo Is Null And A.UCCode=CD013.CodeNo " & _
                 " And Nvl(CD013.REFNO,0) NOT IN(3,7) AND NVL(CD013.PAYOK,0)=0 "
        '#6441 ��i�b��X�p�`�B<=0�O�_���� By Kin 2013/05/13
        strSQL = "SELECT " & GetUseIndexStr("SO033", "ShouldDate") & _
                " A.MediaBillNo BillNo,Sum(A.ShouldAmt) ShouldAmt,A.AccountNo, " & _
                 IIf(isCrossCustCombine, "A.CustId,", "1 CustId,") & _
                "Max(NVL(A.PAYKIND,0)) PAYKIND," & strMinRealStopDate & _
                 " From " & strGetOwner & "SO033 A " & StrTableName & _
                 " Where " & strWhere & _
                 IIf(Len(strWhere) = 0, "", " And ") & strChoose & strHavingOutZero & _
                 " Group By A.AccountNo,A.MediaBillNo" & IIf(isCrossCustCombine, ",A.CustId", "")
        strSQL = "Select * From ( " & strSQL & " ) "
         If isCrossCustCombine Then
            '#7179
            intCrossCustCombine = 1
            sqlTmpViewName = GetTmpViewName(gcnGi)
            '�N�M�лP�M�ХD�Ƚs��X��,�p�G�L�M�Ыh�D�Ƚs��custid�N�J
            strSQL = "Select A.*," & _
                        " (Case " & intCrossCustCombine & " When 1 then Nvl(SO001.AMduId,Null)   else Null End ) AMduId, " & _
                        " (Case " & intCrossCustCombine & " When 1 Then " & _
                                       " ( Case  When AMDUID Is Null Then A.CustId Else Nvl((Select MainCustId From " & strGetOwner & "SO202 Where SO001.AMduId = SO202.MduId  ),-1)  End ) " & _
                        "  Else A.CustId  End ) MainCustId " & _
                        " From (" & strSQL & ") A," & strGetOwner & "SO001 " & _
                " Where A.CustId=SO001.CustId"
            strSQL = "Create View " & strGetOwner & sqlTmpViewName & " As (" & strSQL & ")"
            If Not ExecuteCommand(strSQL, gcnGi, lngCount) Then Exit Function
            '�P�_�M�ХD�Ƚs�O�_���X�{�bSO033.CUSTID
'            strSQL = " select A.*,(select (case " & _
'                                  "                when amduid is null then 1 Else " & _
'                                                          " ( select  count(1) from " & strGetOwner & sqlTmpViewName & _
'                                                          " Where A.MainCustId = Custid And A.AmdUid = AmduId And Nvl(A.BillNo,'X')= Nvl(BillNo,'X') " & _
'                                                           " And A.AccountNo = AccountNo ) " & _
'                                                  " End)  from dual) CustIdExistFlag  from " & strGetOwner & sqlTmpViewName & " A"
            strSQL = " select A.*,(select (case " & _
                                  "                when amduid is null then 1 Else " & _
                                                          " ( select  count(1) from " & strGetOwner & sqlTmpViewName & _
                                                          " Where A.MainCustId = Custid  ) " & _
                                                  " End)  from dual) CustIdExistFlag  from " & strGetOwner & sqlTmpViewName & " A"
            strSQL = "Select Decode(CustIdExistFlag,0,Custid,MainCustId)  CustId,Nvl(BillNo,'X') BillNo,AccountNo,Max(PayKind) PayKind,Min(RealStopDate) RealStopDate," & _
                            " Sum(ShouldAmt) ShouldAmt,AMDUID,CustIdExistFlag " & _
                    " From ( " & strSQL & " ) A Group By Decode(CustIdExistFlag,0,Custid,MainCustId) ,BillNo,AccountNo,AMDUID,CustIdExistFlag "
            '�N�³渹�P�������X��
            strSQL = "Select distinct A.* " & _
                        " From (" & strSQL & ") A"
                       
        End If
                         
        If Not GetRS(rsUpd, strSQLA, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
        '�妸��sMediaBillNo���
        If Not BatchUpdateMediaBillNo(rsUpd, strChoose33, gcnGi) Then Exit Function
         If Not blnHavingOutZero Then
            strSQL = strSQL & " Where ShouldAmt > 0 "
        End If
        If Not GetRS(rsTmp, strSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
    End If
    
    '**************************************************************************************************
    '#3417 ��sUCCode�PUCName���y�k By Kin 2007/12/05
    '#5055 �p�G��CD013.REFNO=3�����i�H�Q���� By Kin 2009/04/28
    '#4388 �W�[��s���ʤH���P���ʮɶ� By Kin 2009/04/29
'    strUPDUCCode = "Select A.UCCode,A.UCName,A.UPDEN,A.UpdTime " & _
'                " From " & strGetOwner & "SO033 A" & StrTableName & _
'              " Where " & strWhere & _
'              IIf(Len(strWhere) = 0, "", " And ") & strChoose
    '7179
    strUPDUCCode = GetUpdateSQL(Val(strMedia & ""), StrTableName, strWhere, _
                        strChoose, blnHavingOutZero, IIf(isCrossCustCombine, strSQL, ""))
    '**************************************************************************************************
        
    Set rsTmp.ActiveConnection = Nothing
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
        While Not rsTmp.EOF
            '#5683 �p�G�����I���p��e��������n�g��ErrorLog By Kin 2010/08/04
            '#5843 �p�G�]�w���ˮ֤��,�h�j���X�b By Kin 2010/11/24
            If (Not IsPayKindOK(rsTmp, gdaTransfer.GetValue & "", giAll, Val(strMedia))) And (blnChkPayKind) Then
                Call WriteSO3271Err(GetPayKindCustId(rsTmp, Val(strMedia)), rsTmp("BillNo"), rsTmp("REALSTOPDATE") & "")
                lngErrCount = lngErrCount + 1
                GoTo lNextRcd
            End If
             '#7179
            If isCrossCustCombine Then
                If Val(rsTmp("CustId") & "") = -1 Then
                        WriteLineData "��ڽs���G " & rsTmp("BillNo") & "�M�Ч䤣��Φ���Ȥ�s���A�г]�w�Φ���Ȥ�s��", 2
                        lngErrCount = lngErrCount + 1
                        GoTo lNextRcd
                End If
'                If Val(rsTmp("CustIdExistFlag") & "") = 0 Then
'                    WriteLineData "��ڽs���G[ " & rsTmp("BillNo") & " ]�@�X�b��ƨS���]�t�Φ���Ȥ�s���G" & rsTmp("CustId"), 2
'                    lngErrCount = lngErrCount + 1
'                    GoTo lNextRcd
'                End If
            End If
           '�p���h���PBillNo �����
           lngCount = lngCount + 1
           If strMedia = 0 Then
                lngShouldAmt = Val(GetRsValue("Select Sum(A.ShouldAmt) From " & strGetOwner & "So033 A Where A.BillNo = '" & rsTmp("BillNo") & "' And A.CancelFlag = 0 And A.UCCode Is Not Null AND " & strChoose33, gcnGi) & "")
           Else
                lngShouldAmt = Val(GetRsValue("Select Sum(A.ShouldAmt) From " & strGetOwner & "So033 A Where A.MediaBillNo = '" & rsTmp("BillNo") & "' And A.AccountNo='" & GetFieldValue(rsTmp, "AccountNo") & "' And A.CancelFlag = 0 And A.UCCode Is Not Null AND " & strChoose33, gcnGi) & "")
                If isCrossCustCombine Then
                    lngShouldAmt = rsTmp("ShouldAmt")
                End If
           End If
           If rsTmp("BillNo") = strOldBillNo & "" Then
              GoTo lNextRcd
           End If
                
           If strMedia = 1 Then
              strBillNo = GetFieldValue(rsTmp, "BillNo") & ""
              strBillNo_Old = GetRsValue("Select BillNo From SO033 Where " & _
                                    " 1=1 " & IIf(isCrossCustCombine, " And Custid=" & rsTmp("CustId"), "") & _
                                    "  And MediaBillNO='" & GetFieldValue(rsTmp, "BillNo") & "'", gcnGi) & ""
           Else
              strBillNo = GetBillNo_New(GetFieldValue(rsTmp, "BillNo") & "")
              strBillNo_Old = GetBillNo_Old(rsTmp("BillNo") & "")
           End If
                                                                  
           '�p�⦩�b�`���B
           lngTotAmt = lngTotAmt + lngShouldAmt
           With rsDefTmp
                .AddNew
                '�ϧO�X 1-1
                .Fields("Code") = "2"
                '�o���� 2-9
                .Fields("SendSpcId") = txtSendSpcId
                '������ 10-17
                .Fields("GotSecId") = txtGotSecId
                '��b���O 18-20
                .Fields("Type") = "160"
                '��b��� 21-26
                .Fields("TransDate") = Replace(gdaTransfer.GetOriginalValue, "/", "")
                '�b�� 27-40
                .Fields("AccountNo") = Mid(GetFieldValue(rsTmp, "AccountNo") & "", 1, 14)
                '������B 41-54
                .Fields("Amt") = lngShouldAmt & "00"
                '�Τ@�s�� 55-62
                .Fields("InVoice") = ""
                '���p�N�� 63-64
                .Fields("Statu") = "99"
                '�C��渹 65-75
                .Fields("MediaBillno") = rsTmp("BillNo") & ""
                strA = Replace(gdaTransfer.GetOriginalValue, "/", "")
                dblC = dblC + Val(Mid(GetFieldValue(rsTmp, "AccountNo") & "", 7, 6))
                dblBtmp = lngShouldAmt
                If dblBtmp > dblBsub Then
                   dblBsub = dblBtmp
                   dblB = Val(Mid(GetFieldValue(rsTmp, "AccountNo") & "", 7, 6))
                End If
           End With
                cnn.Execute Replace("Insert Into SO3271A (AccountNo, TransDate, ShouldAmt, BillNo_Old, BillNo_New) Values (" & _
                     GetNullString(rsTmp("AccountNo")) & "," & GetNullString(gdaTransfer.GetValue(True), giDateV, giAccessDb) & "," & _
                     GetNullString(lngShouldAmt) & "," & GetNullString(strBillNo_Old) & "," & _
                     GetNullString(strBillNo) & ")", Chr(0), "")
                  '7179 ����o�̳�C����
                strUPDUCCode = GetLiteraUpdateSQL(Val(strMedia), StrTableName, strWhere, _
                                 strChoose, IIf(Len(strHavingOutZero) = 0, True, False), rsTmp("BillNo"))
                
                If blnUpdUcCode Then
                    
                    If Not GetRS(rsUpdUCCode, strUPDUCCode, gcnGi, adUseClient, adOpenKeyset, adLockPessimistic) Then Exit Function
                     Do While Not rsUpdUCCode.EOF
                        If (IsPayKindOK(rsUpdUCCode, gdaTransfer.GetValue & "", giSingle, Val(strMedia))) _
                            Or (Not blnChkPayKind) Then
                            rsUpdUCCode("UPDEN") = objStorePara.uUpdEn
                            rsUpdUCCode("UPDTIME") = strUpdTime
                            rsUpdUCCode("UCcode") = strUcCode
                            rsUpdUCCode("UCName") = strUcName
                            rsUpdUCCode.Update
                        End If
                        rsUpdUCCode.MoveNext
                    Loop
                    CloseRecordset rsUpdUCCode
                 End If
lNextRcd:
           strOldBillNo = rsTmp("BillNo")
           rsTmp.MoveNext
           DoEvents
        Wend
          dblTCheck = Val(strA) + dblB + dblBsub + dblC
          strTxt = "A = " & strA & vbCrLf & "B = " & dblB & vbCrLf & _
                   "B'= " & dblBsub & vbCrLf & "C = " & dblC & vbCrLf & "�����`�� = " & Right(dblTCheck, 10)
          txtDataPath = Left(txtDataPath, InStrRev(txtDataPath, "\")) & "�����`��.txt"
          Call CreateObject("Scripting.FileSystemObject").CreateTextFile(txtDataPath, True).Write(strTxt)
        
        cnn.CommitTrans
        '**************************************************************************************************************************
        '#3417 ��sUCCode�PUCName By Kin 2007/12/04
        '#4388 �W�[��s���ʤH���P���ʮɶ� By Kin 2009/04/29
        If blnUpdUcCode And 1 = 0 Then
            
            If Not GetRS(rsUpdUCCode, strUPDUCCode, gcnGi, adUseClient, adOpenKeyset, adLockPessimistic) Then Exit Function
            Do While Not rsUpdUCCode.EOF
                '#5683 �p�G�����I���p��e����������nUPD By Kin 2010/08/04
                '#5843 �p�G�]�w���ˮ�,�h�j���X�b By Kin 2010/11/24
                If (IsPayKindOK(rsUpdUCCode, gdaTransfer.GetValue, giSingle, Val(strMedia & ""))) Or (Not blnChkPayKind) Then
                
                    rsUpdUCCode("UPDEN") = objStorePara.uUpdEn
                    rsUpdUCCode("UPDTime") = strUpdTime
                    rsUpdUCCode("UCcode") = strUcCode
                    rsUpdUCCode("UCName") = strUcName
                    rsUpdUCCode.Update
                End If
                rsUpdUCCode.MoveNext
            Loop
            CloseRecordset rsUpdUCCode
        End If
        '**************************************************************************************************************************
        
        rsTmp.Close
        Set rsTmp = Nothing
        rsBank.Close
        Set rsBank = Nothing
        If strFlowId = 0 Then
            rsPara.Close
            Set rsPara = Nothing
        End If
        CloseRecordset rsCD013
        CloseRecordset rsUpdUCCode
        If Not subWriteLine(lngCount, lngTotAmt) Then
            msgResult lngCount, lngErrCount, lngTime      '��ܰ��浲�G
            CloseFS
            Exit Function      '�Ȥ᤺�e
        End If
        WriteLineData txtMemo1 & vbCrLf & txtMemo2, 1 '�Ƶ���
        msgResult lngCount, lngErrCount, lngTime      '��ܰ��浲�G
        BeginTran = True
        CloseFS
     End If
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
         WriteLineData 1 & GetString(txtSendSpcId, 8) & GetString(txtGotSecId, 8) & _
            "160" & GetString(Replace(gdaTransfer.GetOriginalValue, "/", ""), 6) & "1", 0
         .MoveFirst
         varData = .GetRows
         '���e
         For lngloop = 0 To .RecordCount - 1
            WriteLineData GetString(varData(0, lngloop), 1) & GetString(varData(1, lngloop), 8) & _
                GetString(varData(2, lngloop), 8) & GetString(varData(3, lngloop), 3) & _
                GetString(varData(4, lngloop), 6) & GetString(varData(5, lngloop), 14) & _
                GetString(varData(6, lngloop), 14, GIRIGHT, True) & GetString(varData(7, lngloop), 8) & _
                GetString(varData(8, lngloop), 2) & GetString(varData(9, lngloop), 11), 0
             DoEvents
         Next
         '����
         WriteLineData 3 & GetString(txtSendSpcId, 8) & GetString(txtGotSecId, 8) & _
            "160" & GetString(Replace(gdaTransfer.GetOriginalValue, "/", ""), 6) & _
            GetString(lngTotAmt, 14, GIRIGHT, True) & "00" & GetString(lngCount, 10, GIRIGHT, True) & GetString("", 26, , True), 0
         
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
        .Fields.Append ("SendSpcId"), adBSTR, 8, adFldIsNullable    '�o����
        .Fields.Append ("GotSecId"), adBSTR, 8, adFldIsNullable     '������
        .Fields.Append ("Type"), adBSTR, 3, adFldIsNullable         '��b���O
        .Fields.Append ("TransDate"), adBSTR, 6, adFldIsNullable    '���b���
        .Fields.Append ("AccountNo"), adBSTR, 14, adFldIsNullable   '�b��
        .Fields.Append ("Amt"), adBSTR, 14, adFldIsNullable         '������B
        .Fields.Append ("InVoice"), adBSTR, 8, adFldIsNullable      '�Τ@�s��
        .Fields.Append ("Statu"), adBSTR, 2, adFldIsNullable        '���p�N��
        .Fields.Append ("MediaBillno"), adBSTR, 11, adFldIsNullable '�C��渹
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

        Set rsInv = Nothing

        If Dir(strErrPath & "\" & strPrgName & "Out.log") = "" Then Exit Sub
            Set LogFile = fso.OpenTextFile(strErrPath & "\" & strPrgName & "Out.log", ForReading)
            With LogFile
                    If Not .AtEndOfStream Then gdaTransfer.Text = .ReadLine & ""
                        '�۰���bú�O����
                    If Not .AtEndOfStream Then txtSendSpcId.Text = .ReadLine & ""
                        '�o����N��
                    If Not .AtEndOfStream Then txtGotSecId.Text = .ReadLine & ""
                        '������N��
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
                .WriteLine (gdaTransfer.Text)    '�۰���bú�O����
                .WriteLine (txtSendSpcId.Text)   '�o����N��
                .WriteLine (txtGotSecId.Text)    '������N��
                .WriteLine (txtDataPath)         '�����
                .WriteLine (txtErrPath)          '���D��
                .WriteLine (txtMemoPath)         '�Ƶ���
                .WriteLine (txtMemo1)            '�Ƶ�1
                .WriteLine (txtMemo2)            '�Ƶ�2
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

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    If Len(sqlTmpViewName) > 0 Then
        gcnGi.Execute "Drop View " & strGetOwner & sqlTmpViewName
        sqlTmpViewName = Empty
    End If
End Sub
