VERSION 5.00
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPost4 
   ClientHeight    =   6810
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11190
   Icon            =   "frmPost4.frx":0000
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
      TabIndex        =   12
      Top             =   180
      Width           =   10695
      Begin VB.Frame frmData 
         Caption         =   "����ɦ�m"
         Height          =   1875
         Left            =   450
         TabIndex        =   13
         Top             =   1680
         Width           =   9075
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
            TabIndex        =   2
            ToolTipText     =   "�п�J�r�ɤ����|���ɦW�I"
            Top             =   360
            Width           =   4095
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
            Left            =   7440
            TabIndex        =   3
            Top             =   360
            Width           =   375
         End
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
            Left            =   7440
            TabIndex        =   7
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
            Left            =   7440
            TabIndex        =   5
            Top             =   810
            Width           =   375
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
            TabIndex        =   4
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
            Top             =   1200
            Width           =   4095
         End
         Begin VB.Label lblDatapath 
            AutoSize        =   -1  'True
            Caption         =   "����ɦ�m�]���|�^"
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
            Left            =   360
            TabIndex        =   19
            Top             =   480
            Width           =   1755
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
            Left            =   360
            TabIndex        =   15
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
            Left            =   360
            TabIndex        =   14
            Top             =   1320
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
         TabIndex        =   8
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
         TabIndex        =   9
         Top             =   3720
         Width           =   3015
      End
      Begin Gi_Date.GiDate gdaAgency 
         Height          =   375
         Left            =   2160
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
      Begin Gi_Date.GiDate gdaTransfer 
         Height          =   375
         Left            =   2160
         TabIndex        =   1
         Top             =   1020
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
      Begin VB.Label lblTransfer 
         AutoSize        =   -1  'True
         Caption         =   "�۰���b���ڤ��"
         Height          =   180
         Left            =   540
         TabIndex        =   20
         Top             =   1110
         Width           =   1440
      End
      Begin VB.Label lblagency 
         AutoSize        =   -1  'True
         Caption         =   "ú�O����"
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
         TabIndex        =   18
         Top             =   570
         Width           =   780
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
         TabIndex        =   17
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
         TabIndex        =   16
         Top             =   3750
         Width           =   1365
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "F2.�}�l"
      Height          =   405
      Left            =   3570
      TabIndex        =   10
      Top             =   6180
      Width           =   1275
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "����(&X)"
      Height          =   405
      Left            =   5940
      TabIndex        =   11
      Top             =   6180
      Width           =   1275
   End
End
Attribute VB_Name = "frmPost4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As New ADODB.Connection
Private rsDefTmp As New ADODB.Recordset
Private rsBank As New ADODB.Recordset

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
Private strBillHeadFmt As String    '�h�b�Უ�ͨ̾ڳ]�w

Dim strPostUnit As String           'PostUnit �l���o����
Dim intRefNo As Integer
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
        strBillHeadFmt = .uBillHeadFmt '�h�b�Უ�ͨ̾ڳ]�w PostUnit �l���o����
        intRefNo = .uRefNo
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
    Dim strAMduSQL As String
    Dim strOldBillNo As String
    Dim strOldAccountId As String
    Dim lngTime As Long
    Dim lngErrCount As Long
    Dim lngCount As Long
    Dim lngShouldAmt  As Long
    Dim lngTotAmt As Long
    Dim strBillNo As String, strBillNo_Old As String
    Dim rsTmp As New ADODB.Recordset
    
    Dim rsUpd As New ADODB.Recordset
    Dim rsCD013 As New ADODB.Recordset
    Dim StrTableName As String
    Dim strWhere As String
    Dim strSQLA As String
    Dim strAccNameID As String
    Dim strCitibankATM As String
    Dim strUcCode As String
    Dim strUcName As String
    Dim blnUpdUcCode As Boolean
    Dim strUPDUCCode As String
    Dim strUpdTime As String
    Dim intCrossCustCombine As Integer
    Dim rsUpdUCCode As New ADODB.Recordset
    Dim rsCustId As New ADODB.Recordset
    Dim aBillNoType As String
    Dim aSO106Where As String
  On Error Resume Next
     cnn.BeginTrans
     cnn.Execute ("Delete From SO3271A")
  On Error GoTo ChkErr

    BeginTran = False
    If Len(sqlTmpViewName) > 0 Then
        gcnGi.Execute "Drop View " & strGetOwner & sqlTmpViewName
        sqlTmpViewName = Empty
    End If
    
    lngTime = Timer
    aSO106Where = " And exists (select * from  " & strGetOwner & "so106  " & _
                            " where so106.AccountID = A.AccountNo And SO106.stopFlag <> 1 and SO106.SnactionDate is not null and SO106.ACHCustId is not null and SO106.achtno is not null) "
                                
    'strUpdTime = Format(RightNow, "EE/MM/DD HH:MM:SS")
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
    'If Not GetRS(rsPara, strSQL, gcnGi) Then Exit Function
    'strMedia = GetFieldValue(rsPara, "Para24") & ""
    strMedia = "1"
    
    If InStr(1, strChoose, "C.") Then
        StrTableName = "," & strGetOwner & "SO002 C"
        strWhere = strWhere & " And A.CustId = C.CustId And A.Servicetype = C.ServiceType "
    End If
    If InStr(1, strChoose, "B.") Then
        StrTableName = StrTableName & "," & strGetOwner & "SO001 B"
        strWhere = strWhere & " And A.CustId = B.CustId"
    End If
    
    '************************************************************************
    '#5055 �p�GCD013.REFNO=3(�d�x�w��)���A�X�b By Kin 2009/04/29
    '#5218 �d�x�w���n��Nvl���覡 By Kin 2009/08/05
    '#5564 �W�[�Ѧ�7,8�N��w��,PayOK,�]�N��O�w�� By Kin 2010/05/17
    StrTableName = StrTableName & "," & strGetOwner & "CD013"
    strWhere = strWhere & " And A.CancelFlag<> 1 And UCCode Is Not Null " & _
                " And A.UCCode=CD013.CodeNo And Nvl(CD013.REFNO,0) NOT IN(3,7) " & _
                " AND NVL(CD013.PAYOK,0) = 0 "
    '************************************************************************
    If strWhere <> "" Then strWhere = Mid(strWhere, 5)
    '************************************************************************************************************************************************************************************************************
    strSQLA = "SELECT " & GetUseIndexStr("SO033", "ShouldDate") & _
            " A.BillNo ,A.MediaBillNo,A.AccountNo,A.CustId " & _
            " From " & strGetOwner & "SO033 A " & " Where " & strChoose33 & " And A.MediaBillNo Is Null"
    If Not GetRS(rsUpd, strSQLA, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
    '�妸��sMediaBillNo���
'    If Not BatchUpdateMediaBillNo(rsUpd, strChoose33, gcnGi) Then Exit Function
    strChoose = strChoose & aSO106Where
    aBillNoType = "CitibankATM"
    strSQL = "SELECT " & GetUseIndexStr("SO033", "ShouldDate") & _
            " A.CitibankATM BillNo,Sum(A.ShouldAmt) ShouldAmt,A.AccountNo,A.CustId," & _
            " TO_CHAR(A.ShouldDate,'YYYYMM') as ShouldDate,A.BillNo OldBillNo, " & _
            "(Select ACHCustId From  " & strOwnerName & "SO106 Where SO106.AccountID = A.AccountNo And SO106.StopFlag <> 1 And SO106.SnactionDate is not null  And RowNum = 1 ) ACHCustId, " & _
            " Max(Nvl(A.PayKind,0)) PayKind," & strMinRealStopDate & _
           "From " & strGetOwner & "SO033 A " & StrTableName & _
           " Where " & strWhere & _
           IIf(Len(strWhere) = 0, "", " And ") & strChoose & strHavingOutZero & _
           " Group By A.AccountNo,A.CitibankATM,A.ShouldDate,BillNo,A.CustId Order by BillNo"
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
        strSQL = "Create View " & strOwnerName & sqlTmpViewName & " As (" & strSQL & ")"
        If Not ExecuteCommand(strSQL, gcnGi, lngCount) Then Exit Function
        '�P�_�M�ХD�Ƚs�O�_���X�{�bSO033.CUSTID
        strSQL = " select A.*,(select (case " & _
                              "                when amduid is null then 1 Else " & _
                                                      " ( select  count(1) from " & strGetOwner & sqlTmpViewName & _
                                                      " Where A.MainCustId = Custid  ) " & _
                                              " End)  from dual) CustIdExistFlag  from " & strGetOwner & sqlTmpViewName & " A"
        strSQL = "Select Decode(CustIdExistFlag,0,Custid,MainCustId)  CustId,Nvl(BillNo,'X') BillNo,AccountNo,Max(PayKind) PayKind,Min(RealStopDate) RealStopDate," & _
                        " Sum(ShouldAmt) ShouldAmt,AMDUID,CustIdExistFlag,AchCustId " & _
                " From ( " & strSQL & " ) A Group By Decode(CustIdExistFlag,0,Custid,MainCustId) ,BillNo,AccountNo,AMDUID,CustIdExistFlag,AchCustId "
        '�N�³渹�P�������X��
        strSQL = "Select distinct A.*, " & _
                    " (select OldBillNo from " & strGetOwner & sqlTmpViewName & _
                            " where a.billno=billno and a.custid=maincustid and A.AccountNo = AccountNo And rownum<=1) OldBillNo, " & _
                    " (select ShouldDate from " & strGetOwner & sqlTmpViewName & _
                            " where a.billno=billno and  A.AccountNo = AccountNo And rownum<=1) ShouldDate " & _
                    " From (" & strSQL & ") A"
                   
    End If
    If Not BatchUpdateCitiBankATM(strWhere & IIf(Len(strWhere) = 0, "", " And ") & strChoose, StrTableName, gcnGi) Then Exit Function
    If Not blnHavingOutZero Then
        strSQL = strSQL & " Where ShouldAmt > 0 "
    End If
    If Not GetRS(rsTmp, strSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function

    
    
    
    '************************************************************************************************************************************************************************************************************
    
    '**************************************************************************************************
    intRefNo = 2
    '#7179 �ǤJSQL�y�k�A���M�|Update���~
    strUPDUCCode = GetUpdateSQL(2, StrTableName, strWhere, _
                     strChoose, blnHavingOutZero, IIf(isCrossCustCombine, strSQL, ""))
    '**************************************************************************************************
    Set rsTmp.ActiveConnection = Nothing
    
'    If Not GetRS(rsTmp, strSQL, gcnGi) Then Exit Function

    strSQL = "Select * From " & strGetOwner & "CD018 Where CodeNo ='" & strBankId & "'"
    If Not GetRS(rsBank, strSQL, gcnGi) Then Exit Function
    
    strPostUnit = GetRsValue("Select PostUnit From " & strGetOwner & "CD068 Where BillHeadFmt='" & strBillHeadFmt & "'", gcnGi) & ""
    
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
        strOldAccountId = "-1"
        While Not rsTmp.EOF
            intRefNo = 2
            '#5683 �p�GRealStopDate<�e������,���n�X�b By Kin 2010/08/05
            '#5843 �p�G�]�w���ˮ֫h�j���X�b By Kin 2010/11/24
            If (Not IsPayKindOK(rsTmp, gdaTransfer.GetValue & "", giAll, 2)) And (blnChkPayKind) Then
                Call WriteSO3271Err(GetPayKindCustId(rsTmp, IIf(intRefNo = 2, 3, Val(strMedia))), rsTmp("BillNo"), rsTmp("REALSTOPDATE"))
                lngErrCount = lngErrCount + 1
                GoTo lNextRcd
            End If
           '�p���h���PBillNo �����
           'lngCount = lngCount + 1
           '************************************************************************************************************************************************************************************************************************************************************************************************
           '#3527�@�ѦҸ���2�A�ϥηs��POS�y�{ By Kin 2007/10/04
           lngShouldAmt = Val(GetRsValue("Select Sum(A.ShouldAmt) From " & strGetOwner & "So033 A Where A.CitibankATM = '" & rsTmp("BillNo") & "' And A.AccountNo='" & GetFieldValue(rsTmp, "AccountNo") & "' And A.CancelFlag = 0 And A.UCCode Is Not Null AND " & strChoose33, gcnGi) & "")
            '7179
            If isCrossCustCombine Then
                lngShouldAmt = rsTmp("ShouldAmt") & ""
            End If
            '***********************************************************************************************************************************************************************************************************************************************************************************************
            
            If (rsTmp("BillNo") = strOldBillNo & "") And (rsTmp("AccountNo") = strOldAccountId) Then
               'lngShouldAmt = lngShouldAmt + GetFieldValue(rsTmp, "ShouldAmt") + 0
               GoTo lNextRcd
            End If
            
           '********************************************************************************************************************************************************
           
            strBillNo = GetFieldValue(rsTmp, "BillNo") & ""
            If strBillNo <> "" Then
                strBillNo_Old = GetRsValue("Select BillNo From " & strGetOwner & "SO033 Where CitibankATM='" & GetFieldValue(rsTmp, "BillNo") & "'", gcnGi) & ""
            Else
                WriteLineData "��ڽs���G " & rsTmp("OldBillNo") & "" & " �����b����Ƥ����T�I�I", 2
                lngErrCount = lngErrCount + 1
                GoTo lNextRcd
            End If
          
            '********************************************************************************************************************************************************
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
            lngCount = lngCount + 1
            '�p�⦩�b�`���B
             lngTotAmt = lngTotAmt + lngShouldAmt
              '���D��2386   Lydia�� Edit by Crystal 2006/06/1
              With rsDefTmp
                    '7179 �Ƚs�ݭ��s�P�_,�p�G�Ƚs���h���N�����l��A�p�G�u���@���N��N�O�l��@�� By Kin 2016/03/29
                    Dim aCustid As String
                    aCustid = rsTmp("CustId")
                    If isCrossCustCombine Then
                       
                       If Not GetRS(rsCustId, "Select Distinct Custid From " & strGetOwner & "SO033 Where " & aBillNoType & "='" & rsTmp("BillNO") & "'", _
                                    gcnGi, adUseClient, adOpenKeyset) Then
                                    Exit Function
                        Else
                            If rsCustId.RecordCount = 1 Then
                                aCustid = rsCustId(0)
                            End If
                        End If
                    End If
                   .AddNew
                   '��ƧO
                   .Fields("DataNo") = "1"
                   '�s�ڧO 2-2
                   .Fields("Case") = IIf(Left(GetFieldValue(rsTmp, "AccountNo") & "", 6) = "000000", "G", "P")
                   '�Ʒ~���N�� 3-5
                   .Fields("Comid") = Left(GetFieldValue(rsBank, "BankId") & "" & Space(3), 3)
                   '�ϳB���ҥN�� 6-9
                   .Fields("Citem") = GetString(strPostUnit & "", 4, GIRIGHT)
                   '��b��� 10- 16
                   .Fields("expiryDate") = Format(Trim(Val(gdaAgency.GetValue) - 19110000), "0000000")
                   '�֦L���O 17
                   .Fields("chkNote") = "S"
                   '�l���ϥ���1 18-19
                   .Fields("PosUseOnly") = Space(2)
                   '�x���b�� 20-33
                   .Fields("Acctno") = Right("000000" & Mid(GetFieldValue(rsTmp, "AccountNo") & "", 1, 14), 14)
                   '�O�d�� 34-43
                   .Fields("Reserve1") = Space(10)
                   'ú�O���B 44-54
                   .Fields("ShouldAmt") = Right("000000000" & Mid(lngShouldAmt & "", 1, 9), 9) & "00"
                   '�Τ�s�� 55-74
                   .Fields("ACHCustId") = Right(Space(20) & rsTmp("ACHCustId") & "", 20)
                   '�C�L�Τ�s���O��  75
                   .Fields("ACHCustIdPrint") = "0"
                   '�l���ϥ���2  76
                   .Fields("PosUseOnly2") = Space(1)
                   '���X���O 77
                   .Fields("HideFlag") = "0"
                   '�ܧ�sï�����O�� 78
                   .Fields("ChangePos") = Space(1)
                   '���p�N�� 79-80
                   .Fields("IsOK") = Space(2)
                   'ú�O��� 81-85
                   .Fields("PayDate") = Val(rsTmp("ShouldDate") & "") - 191100
                   '�l���ϥ���3 86-90
                   .Fields("PosUseOnly3") = Space(5)
                   '�e�U���c�ϥ��� 91-110
                   .Fields("CitibankATM") = Left(rsTmp("BillNo") & "" & Space(20), 20)
                   '�O�d�� 111-120
                   .Fields("Reserve2") = Space(10)
                   
                  
                 
              End With
                cnn.Execute Replace("Insert Into SO3271A (AccountNo, TransDate, ShouldAmt, BillNo_Old, BillNo_New) Values (" & _
                     GetNullString(rsTmp("AccountNo")) & "," & GetNullString(gdaTransfer.GetValue(True), giDateV, giAccessDb) & "," & _
                     GetNullString(lngShouldAmt) & "," & GetNullString(strBillNo_Old) & "," & _
                     GetNullString(strBillNo) & ")", Chr(0), "")
                strUPDUCCode = GetLiteraUpdateSQL(IIf(intRefNo = 2, 3, Val(strMedia)), StrTableName, strWhere, _
                                 strChoose, IIf(Len(strHavingOutZero) = 0, True, False), rsTmp("BillNo"))
                
                If blnUpdUcCode Then
                    
                    If Not GetRS(rsUpdUCCode, strUPDUCCode, gcnGi, adUseClient, adOpenKeyset, adLockPessimistic) Then Exit Function
                     Do While Not rsUpdUCCode.EOF
           
                        If (IsPayKindOK(rsUpdUCCode, gdaTransfer.GetValue & "", giSingle, 2)) _
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
           strOldBillNo = rsTmp("BillNo") & ""
           strOldAccountId = rsTmp("AccountNo")
           rsTmp.MoveNext
           DoEvents
        Wend
        cnn.CommitTrans
        
        '**************************************************************************************************************************
        '#3417 ��sUCCode�PUCName By Kin 2007/12/04
        '#4388 �W�[��s���ʤH���P���ʮɶ� By Kin 2009/04/29
        If blnUpdUcCode And 1 = 0 Then
            
            If Not GetRS(rsUpdUCCode, strUPDUCCode, gcnGi, adUseClient, adOpenKeyset, adLockPessimistic) Then Exit Function
            'Set rsUpdUCCode.ActiveConnection = Nothing
            Do While Not rsUpdUCCode.EOF
                '#5683 �p�G�����I���p��e�����󤣭nUPD By Kin 2010/08/05
                '#5843 �p�G�]�w���ˮ�,�h�j���X�b By Kin 2010/11/24
                If (IsPayKindOK(rsUpdUCCode, gdaTransfer.GetValue & "", giSingle, IIf(intRefNo = 2, 3, Val(strMedia)))) _
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
        '**************************************************************************************************************************

        rsTmp.Close
        Set rsTmp = Nothing
        CloseRecordset rsUpdUCCode
        If strFlowId = 0 Then
            'rsPara.Close
            'Set rsPara = Nothing
        End If
        CloseRecordset rsCD013
        CloseRecordset rsCustId
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
    Dim i As Long
    
    Dim strYMdata As String
    Dim strLastData As String
    subWriteLine = False
    With rsDefTmp
         If .BOF Or .EOF Then Exit Function
         .MoveFirst
         varData = .GetRows
         '���e
         For lngloop = 0 To .RecordCount - 1
                     
            WriteLineData varData(0, lngloop) & _
                    varData(1, lngloop) & _
                    varData(2, lngloop) & _
                    varData(3, lngloop) & _
                     varData(4, lngloop) & _
                    varData(5, lngloop) & _
                    varData(6, lngloop) & _
                     varData(7, lngloop) & _
                    varData(8, lngloop) & _
                    varData(9, lngloop) & _
                     varData(10, lngloop) & _
                    varData(11, lngloop) & _
                    varData(12, lngloop) & _
                     varData(13, lngloop) & _
                    varData(14, lngloop) & _
                    varData(15, lngloop) & _
                     varData(16, lngloop) & _
                    varData(17, lngloop) & _
                    varData(18, lngloop) & _
                    varData(19, lngloop), 0

           
             DoEvents
         Next
'         '����
                
        WriteLineData "2" & GetString("", 1) & _
                   GetString(rsBank("BankId") & "", 3) & _
                   GetString(strPostUnit, 4, GIRIGHT) & _
                   Format(Trim(Val(gdaAgency.GetValue) - 19110000), "0000000") & _
                   Space(3) & _
                   GetString(lngCount, 7, GIRIGHT, True) & _
                   GetString(lngTotAmt, 11, GIRIGHT, True) & "00" & _
                   Space(16) & _
                   String(7, "0") & _
                   String(13, "0") & _
                   Space(45), 0
                    
    End With
    subWriteLine = True
    On Error Resume Next
    rsDefTmp.Close
    Set rsDefTmp = Nothing
    rsBank.Close
    Set rsBank = Nothing
    Exit Function
ChkErr:
    subWriteLine = False
    Call ErrSub(Me.Name, "subWriteLine")
    Resume 0
End Function

'�w�q�O����Recordset
Private Sub DefineRs()
    On Error GoTo ChkErr
    With rsDefTmp
        If .State = adStateOpen Then .Close
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        
        .Fields.Append ("DataNo"), adBSTR, 1, adFldIsNullable       '��ƧO
        .Fields.Append ("Case"), adBSTR, 1, adFldIsNullable         '�s�ڧO
        .Fields.Append ("Comid"), adBSTR, 3, adFldIsNullable        '�Ʒ~���N��
        .Fields.Append ("Citem"), adBSTR, 4, adFldIsNullable        '�ϳB���ҥN��
        .Fields.Append ("expiryDate"), adBSTR, 6, adFldIsNullable    'ú�O����
        .Fields.Append ("chkNote"), adBSTR, 1, adFldIsNullable '�֦L���O
        .Fields.Append ("PosUseOnly"), adBSTR, adFldIsNullable '�l���ϥ���1
        .Fields.Append ("Acctno"), adBSTR, 14, adFldIsNullable      'ú�O�b��
        .Fields.Append ("Reserve1"), adBSTR, 10, adFldIsNullable '�O�d��
        .Fields.Append ("ShouldAmt"), adBSTR, 10, adFldIsNullable 'ú�O���B
        .Fields.Append ("ACHCustId"), adBSTR, 20, adFldIsNullable '�Τ�s��
        .Fields.Append ("ACHCustIdPrint"), adBSTR, 1, adFldIsNullable '�C�L�Τ�s���O��
        .Fields.Append ("PosUseOnly2"), adBSTR, 1, adFldIsNullable '�l���ϥ���2
        .Fields.Append ("HideFlag"), adBSTR, 1, adFldIsNullable '���X���O
        .Fields.Append ("ChangePos"), adBSTR, 1, adFldIsNullable '�ܧ�sï�����O��
        .Fields.Append ("IsOK"), adBSTR, 2, adFldIsNullable '���p�N��
        .Fields.Append ("PayDate"), adBSTR, 5, adFldIsNullable 'ú�O���
        .Fields.Append ("PosUseOnly3"), adBSTR, 5, adFldIsNullable '�l���ϥ���3
        .Fields.Append ("CitibankATM"), adBSTR, 20, adFldIsNullable '�e�U���c�ϥ���
        .Fields.Append ("Reserve2"), adBSTR, 10, adFldIsNullable '�O�d��
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
        txtDataPath = FolderDialog("�}�l")
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "cmdDataPath_Click")
End Sub
Public Function FolderDialog(Title As String) As String
  On Error GoTo ChkErr
    Dim ComDlg As Object
    Dim Result As String
    Set ComDlg = CreateObject("Common.Dialog")
    Result = ComDlg.FolderDialog(Title)
    Set ComDlg = Nothing
    FolderDialog = Result
  Exit Function
ChkErr:
    ErrSub Me.Name, "FolderDialog"
End Function

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
    Dim strFileName1 As String
    Dim lngFilePath As Long
    Dim strFileName As String
    If Not ChkDTok Then Exit Sub
    If Not IsDataOk Then Exit Sub
    '��X�ץX���ɮצW��
    strFileName1 = GetRsValue("Select FileName1 From " & strGetOwner & "CD018 Where CodeNo=" & strBankId, gcnGi) & ""
    '�ק��x�s�b�e�����O��,�u���x�s��ƪ����|,����s�ɦW���x,���M�|�PstrFileName1�Ĭ�
    If InStr(txtDataPath, strFileName1) = 0 Then
      lngFilePath = InStrRev(strFileName1, "\")
      If lngFilePath <> 0 Then strFileName1 = Mid(strFileName1, lngFilePath + 1, 9999)
        strFileName = txtDataPath & IIf(Right(txtDataPath, 1) = "\", "", "\") & strFileName1
    End If
    If strFileName1 = "" Then MsgBox "�L�]�w����ɦW��!!", vbExclamation, gimsgPrompt: txtDataPath.SetFocus: Exit Sub
    '#5596 ���դ�OK �n����ɦW By Kin 2010/04/30
    If gdaTransfer.GetValue <> "" Then
        strFileName = Replace(strFileName, "YYYYMMDD", gdaTransfer.GetValue)
    End If
'    If OpenFile(txtDataPath, 0, True) = False Then txtDataPath.SetFocus: CloseFS: Exit Sub
    If OpenFile(strFileName, 0, True) = False Then txtDataPath.SetFocus: CloseFS: Exit Sub
    If OpenFile(txtMemoPath, 1, True) = False Then txtMemoPath.SetFocus: CloseFS: Exit Sub
    If OpenFile(txtErrPath, 2, True) = False Then txtErrPath.SetFocus: CloseFS: Exit Sub
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
        If gdaAgency.GetValue = "" Then strErrMsg = "ú�O����": gdaAgency.SetFocus: GoTo Warning
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
'        txtInVoice.Text = GetFieldValue(rsInv, "InvoiceId") & ""
        
        rsInv.Close
        Set rsInv = Nothing

        If Dir(strErrPath & "\" & strPrgName & "Out.log") = "" Then Exit Sub
            Set LogFile = fso.OpenTextFile(strErrPath & "\" & strPrgName & "Out.log", ForReading)
            With LogFile
                    If Not .AtEndOfStream Then gdaAgency.Text = .ReadLine & ""
                       '�{�d�N��ú�O����
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
'    Dim strFilePath As String
'    strFilePath = txtDataPath.Text
'    strFilePath = Left(strFilePath, InStrRev(strFilePath, "\")) - 1
    
        Set LogFile = fso.CreateTextFile(strErrPath & "\" & strPrgName & "Out.log", True)
         With LogFile
               Call .WriteLine(gdaAgency.Text)   '�{�d�N��ú�O����
                .WriteLine (gdaTransfer.Text)    '�۰���bú�O����
                .WriteLine (txtDataPath)     '�����
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
