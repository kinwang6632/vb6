VERSION 5.00
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{A0C51418-555F-11D3-BC10-0050BA01A3A9}#9.9#0"; "GiMulti.ocx"
Object = "{074771B7-6ADE-11D5-B178-00485456B028}#2.1#0"; "csMulti.ocx"
Begin VB.Form frmSO3277A 
   AutoRedraw      =   -1  'True
   Caption         =   "ACH��b���ڸ�Ʋ��ͧ@�~ [SO3277A]"
   ClientHeight    =   7980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11190
   BeginProperty Font 
      Name            =   "�s�ө���"
      Size            =   9.75
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SO3277A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7980
   ScaleWidth      =   11190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '�ù�����
   Begin VB.TextBox txtSQL 
      Height          =   1155
      Left            =   3090
      TabIndex        =   25
      Top             =   3330
      Visible         =   0   'False
      Width           =   5085
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "����(&X)"
      Height          =   405
      Left            =   6360
      TabIndex        =   19
      Top             =   7530
      Width           =   1275
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "F2.�T�w"
      Default         =   -1  'True
      Height          =   405
      Left            =   3270
      TabIndex        =   18
      Top             =   7530
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      Caption         =   "�d�߱���"
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
      Height          =   7275
      Left            =   150
      TabIndex        =   20
      Top             =   195
      Width           =   10905
      Begin CS_Multi.CSmulti gimCitemCode 
         Height          =   390
         Left            =   210
         TabIndex        =   36
         Top             =   1680
         Width           =   10605
         _ExtentX        =   18706
         _ExtentY        =   688
         ButtonCaption   =   "�� �O �� ��"
         IsReadOnly      =   -1  'True
      End
      Begin VB.TextBox txtCustId 
         Height          =   360
         Left            =   2160
         TabIndex        =   32
         Top             =   5850
         Width           =   5925
      End
      Begin VB.CheckBox chkErr 
         Caption         =   "�ˮֲ��X�q�l��"
         Height          =   315
         Left            =   1680
         TabIndex        =   30
         Top             =   1320
         Width           =   1695
      End
      Begin VB.CheckBox chkZero 
         Caption         =   "��i�b��X�p�`�B<=0 �O�_�n����"
         Height          =   195
         Left            =   5490
         TabIndex        =   29
         Top             =   1290
         Width           =   3315
      End
      Begin VB.ComboBox cboACHbankType 
         Height          =   315
         ItemData        =   "SO3277A.frx":0442
         Left            =   5430
         List            =   "SO3277A.frx":0444
         Style           =   2  '��¤U�Ԧ�
         TabIndex        =   4
         Top             =   825
         Width           =   2925
      End
      Begin VB.TextBox txtACHTNo 
         Enabled         =   0   'False
         Height          =   330
         Left            =   9705
         TabIndex        =   5
         Top             =   795
         Width           =   1035
      End
      Begin VB.ComboBox cboBillHeadFmt 
         Height          =   315
         ItemData        =   "SO3277A.frx":0446
         Left            =   6390
         List            =   "SO3277A.frx":0448
         Style           =   2  '��¤U�Ԧ�
         TabIndex        =   3
         Top             =   345
         Width           =   4365
      End
      Begin VB.CheckBox chkIncludeNormal 
         Caption         =   "�]�t�@���"
         Height          =   195
         Left            =   300
         TabIndex        =   6
         Top             =   1380
         Value           =   1  '�֨�
         Width           =   1335
      End
      Begin Gi_Multi.GiMulti gimCustStatusCode 
         Height          =   405
         Left            =   210
         TabIndex        =   11
         Top             =   3480
         Width           =   10605
         _ExtentX        =   18706
         _ExtentY        =   714
         ButtonCaption   =   "��  ��  ��  �A"
         FontSize        =   9
         FontName        =   "�s�ө���"
      End
      Begin VB.Frame Frame2 
         Caption         =   " �j�Ө̾�"
         Height          =   1005
         Left            =   240
         TabIndex        =   23
         Top             =   6210
         Width           =   10545
         Begin VB.OptionButton optMduYes 
            Caption         =   "&A.�n����"
            Height          =   195
            Left            =   510
            TabIndex        =   16
            Top             =   720
            Value           =   -1  'True
            Width           =   1065
         End
         Begin VB.OptionButton optMduNo 
            Caption         =   "&B.������"
            Height          =   195
            Left            =   1800
            TabIndex        =   17
            Top             =   720
            Width           =   1095
         End
         Begin CS_Multi.CSmulti gmsMduId 
            Height          =   405
            Left            =   90
            TabIndex        =   15
            Top             =   240
            Width           =   10395
            _ExtentX        =   18336
            _ExtentY        =   714
            ButtonCaption   =   "�j�ӦW��"
         End
      End
      Begin Gi_Multi.GiMulti gmdBillType 
         Height          =   375
         Left            =   210
         TabIndex        =   14
         Top             =   4560
         Width           =   10605
         _ExtentX        =   18706
         _ExtentY        =   661
         ButtonCaption   =   "��  ��  ��  �O"
         FontSize        =   9
         FontName        =   "�s�ө���"
      End
      Begin Gi_Multi.GiMulti gmdServCode 
         Height          =   465
         Left            =   210
         TabIndex        =   10
         Top             =   3120
         Width           =   10605
         _ExtentX        =   18706
         _ExtentY        =   820
         ButtonCaption   =   "�A     ��     ��"
         FontSize        =   9
         FontName        =   "�s�ө���"
      End
      Begin Gi_Multi.GiMulti gmdAreaCode 
         Height          =   465
         Left            =   210
         TabIndex        =   9
         Top             =   2760
         Width           =   10605
         _ExtentX        =   18706
         _ExtentY        =   820
         ButtonCaption   =   "��     �F     ��"
         FontSize        =   9
         FontName        =   "�s�ө���"
      End
      Begin Gi_Multi.GiMulti gmdCMCode 
         Height          =   465
         Left            =   210
         TabIndex        =   8
         Top             =   2400
         Width           =   10605
         _ExtentX        =   18706
         _ExtentY        =   820
         ButtonCaption   =   "��  �O  ��  ��"
         FontSize        =   9
         FontName        =   "�s�ө���"
      End
      Begin Gi_Date.GiDate gdaShouldDate2 
         Height          =   345
         Left            =   3120
         TabIndex        =   2
         Top             =   840
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   609
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
      Begin Gi_Date.GiDate gdaShouldDate1 
         Height          =   345
         Left            =   1530
         TabIndex        =   1
         Top             =   840
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   609
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
      Begin Gi_Multi.GiMulti gimBankId 
         Height          =   465
         Left            =   210
         TabIndex        =   7
         Top             =   2040
         Width           =   10605
         _ExtentX        =   18706
         _ExtentY        =   820
         ButtonCaption   =   "��b�Ȧ�W��"
         DataType        =   2
         FontSize        =   9
         FontName        =   "�s�ө���"
      End
      Begin prjGiList.GiList gilCompCode 
         Height          =   315
         Left            =   1515
         TabIndex        =   0
         Top             =   360
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
      Begin CS_Multi.CSmulti gmdClctEn 
         Height          =   345
         Left            =   210
         TabIndex        =   13
         Top             =   4200
         Width           =   10605
         _ExtentX        =   18706
         _ExtentY        =   609
         ButtonCaption   =   "��  �O  �H  ��"
      End
      Begin CS_Multi.CSmulti gmdCustClass 
         Height          =   345
         Left            =   210
         TabIndex        =   12
         Top             =   3840
         Width           =   10605
         _ExtentX        =   18706
         _ExtentY        =   609
         ButtonCaption   =   "��  ��  ��  �O"
      End
      Begin Gi_Multi.GiMulti gimUCCode 
         Height          =   465
         Left            =   210
         TabIndex        =   31
         Top             =   4920
         Width           =   10605
         _ExtentX        =   18706
         _ExtentY        =   820
         ButtonCaption   =   "������]"
         FontSize        =   9
         FontName        =   "�s�ө���"
      End
      Begin Gi_Multi.GiMulti gmdPayType 
         Height          =   390
         Left            =   210
         TabIndex        =   35
         Top             =   5340
         Width           =   10605
         _ExtentX        =   18706
         _ExtentY        =   688
         ButtonCaption   =   "ú  �I  ��  �O"
         FontSize        =   9
         FontName        =   "�s�ө���"
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "�ХΡu,�v���} Exp 67,82,199"
         Height          =   315
         Left            =   8130
         TabIndex        =   34
         Top             =   5940
         Width           =   2535
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "���ͫȽs"
         Height          =   195
         Left            =   1230
         TabIndex        =   33
         Top             =   5940
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "ACH����N��"
         Height          =   195
         Left            =   8445
         TabIndex        =   28
         Top             =   870
         Width           =   1170
      End
      Begin VB.Label lblBillHeadFmt 
         AutoSize        =   -1  'True
         Caption         =   "�h�b�Უ�ͨ̾ڳ]�w"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   4560
         TabIndex        =   27
         Top             =   420
         Width           =   1755
      End
      Begin VB.Label lblBankHand 
         AutoSize        =   -1  'True
         Caption         =   "�ӿ�Ȧ�"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   4545
         TabIndex        =   26
         Top             =   885
         Width           =   780
      End
      Begin VB.Label lblCompCode 
         AutoSize        =   -1  'True
         Caption         =   "���q�O"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   300
         TabIndex        =   24
         Top             =   450
         Width           =   585
      End
      Begin VB.Label lblReadAmt 
         AutoSize        =   -1  'True
         Caption         =   "��������d��"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   270
         TabIndex        =   22
         Top             =   930
         Width           =   1170
      End
      Begin VB.Label lbluntil 
         AutoSize        =   -1  'True
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   2760
         TabIndex        =   21
         Top             =   900
         Width           =   180
      End
   End
End
Attribute VB_Name = "frmSO3277A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public strChoose As String
Public strPrgName As String
Private strChooseString As String
Private strChoose33 As String
Dim rsBillHeadFmt As New ADODB.Recordset
Dim intPara24 As Integer
Private rsBankData As New ADODB.Recordset
Public Property Get uOutZero() As Boolean
    uOutZero = chkZero.Value = 1
End Property
Private Function SetACHbankTypeCbo() As Boolean
  Dim strSQL As String
  Dim rsTmp As New ADODB.Recordset
  Dim lbn
  Dim lngIndex As Long
  On Error GoTo chkErr
  
  For lngIndex = 0 To cboACHbankType.ListCount - 1
       cboACHbankType.RemoveItem (0)
  Next
    lngIndex = 0
    strSQL = "SELECT DISTINCT PrgName FROM " & GetOwner & "CD018 WHERE UPPER(PrgName) like '%" & "ACH" & "%'  AND COMPCODE =" & gilCompCode.GetCodeNo & "  AND STOPFLAG = 0 "
    Call GetRS(rsTmp, strSQL, gcnGi, adUseClient, adOpenDynamic)
    Dim blnAdd As Boolean   ''  �P�_�p�G���[�JACH���زM�檺�ܡA�h�N��]��flase �A�o�˩��U��lngIndex�~�i�H�[ 1  ���M�U����ƪ��j��|�X�{ ���޶W�X�}�C�d��
    blnAdd = False
  
    Do While Not (rsTmp.EOF Or rsTmp.BOF)
        Select Case UCase(rsTmp("PrgName"))
            Case "ACHTRANREFER"
                 cboACHbankType.AddItem (" �Ĥ@�Ȧ�ACH���X�榡 ")
                 cboACHbankType.ItemData(lngIndex) = 0
                 blnAdd = True
          End Select
          If blnAdd = True Then lngIndex = lngIndex + 1
          blnAdd = False
         rsTmp.MoveNext
     Loop
    rsTmp.Close
    Set rsTmp = Nothing
      If lngIndex = 0 Then
         MsgBox "�Ȧ�O��ƥ��]�A�г]�w�����Ȧ�O�N�X����A���s����  !!"
         SetACHbankTypeCbo = False
    Else
         cboACHbankType.ListIndex = 0
          SetACHbankTypeCbo = True
    End If
    Screen.MousePointer = vbDefault
   
Exit Function
chkErr:
  ErrSub Me.Name, "SetACHbankTypeCbo"
End Function

Private Sub cboACHbankType_Click()
  Select Case cboACHbankType.ItemData(cboACHbankType.ListIndex)
       Case 0
           Call SetgiMulti(gimBankId, "CodeNo", "Description", "CD018", "�N�X", "�W��", _
                                   " Where UPPER(PrgName) = 'ACHTRANREFER'  AND  COMPCODE =" & _
                                   gilCompCode.GetCodeNo)
    End Select
End Sub

Private Sub CboAddItem()
    On Error Resume Next
        '#4106 �W�[�P�_ACHType=0 By Kin 2008/09/22
        If Not GetRS(rsBillHeadFmt, "Select BillHeadFmt,CitemCodeStr From " & GetOwner & "CD068 WHERE ACHTNo IS NOT NULL ") Then Exit Sub
        Do While Not rsBillHeadFmt.EOF
            cboBillHeadFmt.AddItem rsBillHeadFmt("BillHeadFmt") & ""
            rsBillHeadFmt.MoveNext
        Loop
End Sub

Private Sub cboBillHeadFmt_Click()
    On Error Resume Next
    Dim strCitemCode As String
    Dim rsCD019 As New ADODB.Recordset
    Dim rsCD068 As New ADODB.Recordset
    Dim strCD019SQL As String
        If rsBillHeadFmt.RecordCount = 0 Then Exit Sub
        rsBillHeadFmt.AbsolutePosition = cboBillHeadFmt.ListIndex + 1
        '#7049 ���CD068A.CitemCode By Kin 2015/07/14
        strCD019SQL = "Select CodeNo From " & GetOwner & "CD019 " & _
                            " Where CodeNo In ( Select CitemCode From " & GetOwner & "CD068A " & _
                            "         Where Exists (Select * From " & GetOwner & "CD068 Where CD068.BillHeadFmt = CD068A.BillHeadFmt " & _
                                                    " And CD068.BillHeadFmt = '" & rsBillHeadFmt("BillHeadFmt") & "' )) And StopFlag <>1"
'        If Not GetRS(rsCD019, "Select CodeNo From " & GetOwner & "CD019 Where CodeNo In (" & rsBillHeadFmt("CitemCodeStr") & ") And StopFlag=0") Then Exit Sub
    If Not GetRS(rsCD019, strCD019SQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
    If rsCD019.EOF Then Exit Sub
    
    If Not GetRS(rsCD068, "Select ACHTNO,PostUnit From " & GetOwner & "CD068 Where BillHeadFmt='" & cboBillHeadFmt.Text & "' And ACHTNO is Not Null") Then Exit Sub
    txtACHTNo = rsCD068("ACHTNO") & ""
    txtACHTNo.Tag = rsCD068("PostUnit") & ""
    strCitemCode = rsCD019.GetString(, , , ",")
    strCitemCode = Left(strCitemCode, Len(strCitemCode) - 1)
    gimCitemCode.Filter = gimCitemCode.Filter
    gimCitemCode.SetQueryCode strCitemCode
        'gimCitemCode.ShowMulti
End Sub

Private Sub cmdCancel_Click()
  On Error GoTo chkErr
    Unload Me
  Exit Sub
chkErr:
    Call ErrSub(Me.Name, "cmdCancel")
End Sub
Private Function GetBankData() As Boolean
On Error GoTo chkErr
Dim strWhere As String
Dim strSQL  As String
   
   If Len(gimBankId.GetQueryCode) <> 0 Then
      strWhere = " AND CodeNo In(" & gimBankId.GetQueryCode & ")"
   End If
   strSQL = "Select CodeNO,Description,PrgName,ActNo,BankId,CorpID,Spcno From " & GetOwner & "CD018 where " & " UPPER(PrgName) like '%" & "ACHTRANREFER" & "%'" & strWhere
   Set rsBankData = gcnGi.Execute(strSQL)
   If rsBankData.EOF Or rsBankData.BOF Then MsgBox "�Ȧ��Ʀ��~�I���ˬd�Ȧ�N�X�ɡI", vbExclamation, "�T���I": Exit Function
   GetBankData = True
Exit Function
chkErr:
    Call ErrSub(Me.Name, "GetBankData")
End Function

Private Sub cmdOk_Click()
  On Error GoTo chkErr
    Dim objAction As Object
    Dim objInterface As Object
    Dim lngSecond As Long
    Dim strBank As String
    
        If IsDataOk = False Then Exit Sub
        Set objInterface = New clsInterface
        Call subChoose          '��Where ���O
        If Not GetBankData Then Exit Sub
            
            If Len(rsBankData("PrgName") & "") = 0 Then
                MsgBox "�]�w�{���W�ٵL�]�εL�ϥ��v���I�I", vbExclamation, "����"
            Else
                If InStr(1, gimBankId.GetQueryCode, ",") > 0 Then
                    strBank = Mid(gimBankId.GetQueryCode, 1, InStr(1, gimBankId.GetQueryCode, ",") - 1)
                Else
                    strBank = gimBankId.GetQueryCode
                End If
                If rsBankData("Prgname").Value & "" = "ACHTRANREFER" Then
                 strPrgName = rsBankData("Prgname").Value & ""
                 With frmACHTranRefer
                    .uPostUnit = txtACHTNo.Tag
                    .uChkProc = chkErr.Value
                    .uBank = strBank
                    .uChooseString = strChooseString
                    .Show 1
                 End With
                End If
            End If
            
    Exit Sub
chkErr:
    If Err.Number = 20 Then Resume Next
    If Err.Number = 91 Then
       MsgBox "��b�{���W�ٿ��~�Ϊ̸ӻȦ�S����b��Ʋ��ͥ\��!", vbExclamation, "ĵ�i..."
       Exit Sub
    End If
    Call ErrSub(Me.Name, "cmdOk_Click")
End Sub

Private Function InsertToLog(lngSecond As Long)
    On Error Resume Next
    Dim strSQL As String
        strSQL = "INSERT INTO " & GetOwner & "SO097 (EMPNO,PRGNAME,RUNDATETIME,SECOND,SELECTION) " & _
                       "VALUES ('" & garyGi(0) & "','SO3274A',SysDate,'" & _
                       lngSecond & "��','" & Replace(strChooseString, "'", "") & "')"
        gcnGi.Execute strSQL
End Function

Private Sub subChoose()
    On Error GoTo chkErr
    Dim strAddr As String
    Dim strMduChoose As String
        strChoose = ""
        strChooseString = ""
           
        If gdaShouldDate1.GetValue <> "" Then Call subAnd("A.ShouldDate >= To_date(" & gdaShouldDate1.GetValue & ",'YYYYMMDD')")
        If gdaShouldDate2.GetValue <> "" Then Call subAnd("A.ShouldDate < To_date(" & gdaShouldDate2.GetValue & ",'YYYYMMDD') +1")
        If gmdAreaCode.GetQryStr <> "" Then Call subAnd("A.AreaCode " & gmdAreaCode.GetQryStr)
        If gimBankId.GetQueryCode <> "" Then
            If gimBankId.GetSelectCount = 1 Then
                Call subAnd("A.BankCode = " & gimBankId.GetQueryCode)
            Else
                Call subAnd("A.BankCode In (" & gimBankId.GetQueryCode & ")")
            End If
        End If
        
        If gmdBillType.GetQryStr <> "" Then
            If gmdBillType.GetSelectCount = 1 Then
                subAnd "SubStr(A.BillNo,7,1) = " & gmdBillType.GetQueryCode
            Else
                subAnd "SubStr(A.BillNo,7,1) In (" & gmdBillType.GetQueryCode & ")"
            End If
        Else
            Call subAnd("SubStr(A.BillNo,7,1) In ('B','T','I','M','P')")
        End If
        '#4388 �W�[������] By Kin 2009/04/30
        If gimUCCode.GetQryStr <> "" Then subAnd ("A.UCCODE " & gimUCCode.GetQryStr)
        '#4388 �W�[�Ƚs���� By Kin 2009/04/29
        If txtCustId.Text <> "" Then
            Call TimetxtCustId(txtCustId)
        End If
        '#5683 �W�[ú�I���O By Kin 2010/08/11
        If gmdPayType.GetQryStr <> "" Then
            subAnd ("A.PAYKIND " & gmdPayType.GetQryStr)
        End If
        '92/05/29 �[���O����
'        If gimCitemCode.GetQryStr <> "" Then subAnd ("A.CitemCode " & gimCitemCode.GetQryStr)
        '#7049 �ϥ�CD068A.CitemCode By Kin 2015/07/14
        If Len(gimCitemCode.GetQryStr & "") > 0 Then
            subAnd ("A.CitemCode In (Select CitemCode From " & GetOwner & "CD068A " & _
                                                                    " Where BillHeadFmt = '" & rsBillHeadFmt("BillHeadFmt") & "')")
        End If
        If gmdClctEn.GetQryStr <> "" Then Call subAnd("A.OldClctEn " & gmdClctEn.GetQryStr)
        If gmdCMCode.GetQryStr <> "" Then Call subAnd("A.CMCode " & gmdCMCode.GetQryStr)
        If gmdCustClass.GetQryStr <> "" Then Call subAnd("A.ClassCode " & gmdCustClass.GetQryStr)
'        Select Case garyGi(17)
'        Case 1  'TBC
'            If gimCustStatusCode.GetQryStr <> "" Then Call subAnd("B.CustStatusCode " & gimCustStatusCode.GetQryStr)
'        End Select
        '���ͤj��
        If optMduYes Then
            '���ͤ@���
            If chkIncludeNormal.Value = 1 Then
               If gmsMduId.GetQryStr <> "" Then subAnd ("(A.MduId " & gmsMduId.GetQryStr & " Or A.MduId Is Null) ")
            Else
            '�����ͤ@���
               If gmsMduId.GetQryStr <> "" Then
                  subAnd ("A.MduId " & gmsMduId.GetQryStr)
               Else
                  subAnd ("A.MduId Is Not Null")
               End If
            End If
        '�����ͤj��
        Else
            '���ͤ@���
            If chkIncludeNormal.Value = 1 Then
                If gmsMduId.GetQryStr <> "" Then
                   subAnd ("(Not A.MduId " & gmsMduId.GetQryStr & " Or A.MduId Is Null) ")
                Else
                   subAnd ("A.MduId Is Null")
                End If
            Else
                subAnd " 1 =0 "
            End If
        End If
        
        If gmdServCode.GetQryStr <> "" Then subAnd ("A.ServCode " & gmdServCode.GetQryStr)
        
        'If chkZero.Value = 0 Then subAnd "A.ShouldAmt > 0"
        strChoose33 = strChoose
        
        If gimCustStatusCode.GetQryStr <> "" Then Call subAnd("E.CustStatusCode " & gimCustStatusCode.GetQryStr)
        
        strChooseString = "���q�O�@    :" & subSpace(gilCompCode.GetDescription) & ";" & _
                          "�������  :" & subSpace(gdaShouldDate1.GetOriginalValue) & "~" & subSpace(gdaShouldDate2.GetOriginalValue) & ";" & _
                          "�ӿ�Ȧ�:" & subSpace(cboACHbankType.Text) & ";" & _
                          "ACH����N��:" & subSpace(txtACHTNo.Text) & ";" & _
                          "���O����:" & subSpace(gimCitemCode.GetDispStr) & ";" & _
                          "��b�Ȧ�W��:" & subSpace(gimBankId.GetDispStr) & ";" & _
                          "���O�覡  :" & subSpace(gmdCMCode.GetDispStr) & ";" & _
                          "��F��      :" & subSpace(gmdAreaCode.GetDispStr) & ";" & _
                          "�A�Ȱ�      :" & subSpace(gmdServCode.GetDispStr) & ";" & _
                          "�Ȥ᪬�A  :" & subSpace(gimCustStatusCode.GetDispStr) & ";" & _
                          "�Ȥ����O  :" & subSpace(gmdCustClass.GetDispStr) & ";" & _
                          "���O�H��      :" & subSpace(gmdClctEn.GetDispStr) & ";" & _
                          "������O  :" & subSpace(gmdBillType.GetDispStr) & ";" & _
                          "������]  :" & subSpace(gimUCCode.GetDispStr) & ";" & _
                          "ú�I���O  :" & subSpace(gmdPayType.GetDispStr) & ";" & _
                          "���ͫȽs  :" & subSpace(txtCustId.Text) & ";" & _
                          "�j�ӦW��  :" & subSpace(gmsMduId.GetDispStr)
        
        
    Exit Sub
chkErr:
    Call ErrSub(Me.Name, "subChoose")
End Sub

Private Sub Form_Activate()
    On Error Resume Next
        Screen.MousePointer = vbDefault
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If Shift = 2 Then
        If KeyCode = vbKeyF Then
            KeyCode = 0
            txtSQL.Visible = True
            txtSQL = strChoose
        ElseIf KeyCode = vbKeyX Then
            txtSQL.Visible = False
        End If
    End If
    If KeyCode = vbKeyF2 Then Call cmdOk_Click: KeyCode = 0
    If KeyCode = vbKeyEscape Then Unload Me: KeyCode = 0
End Sub

Private Sub Form_Load()
    On Error GoTo chkErr
        'Call subGim
        Call subGil
        Call DefaultValue
        CboAddItem
Exit Sub
chkErr:
    Call ErrSub(Me.Name, "Form_Load")
End Sub

Private Sub DefaultValue()
    On Error Resume Next
    Dim intPara23 As Integer
        gilCompCode.SetCodeNo garyGi(9)
        gilCompCode.Query_Description
        chkErr.Enabled = gcnGi.Execute("Select CheckText From " & GetOwner & "SO041  WHERE CompCode = " & garyGi(9))(0)
        If chkErr.Enabled Then chkErr.Value = 1 Else chkErr.Value = 0
End Sub

Private Sub subGil()
  On Error GoTo chkErr
    Call SetgiList(gilCompCode, "CodeNo", "Description", "CD039", , , , , , , GetCompCodeFilter(gilCompCode))
  Exit Sub
chkErr:
    Call ErrSub(Me.Name, "subGil")
End Sub

Private Sub subGim()
    On Error GoTo chkErr
        
        SetgiMulti gmdCMCode, "CodeNo", "Description", "CD031", "�N�X", "�W��", , True
        SetgiMulti gmdAreaCode, "CodeNo", "Description", "CD001", "�N�X", "�W��", , True
        SetgiMulti gmdServCode, "CodeNo", "Description", "CD002", "�N�X", "�W��", , True
        SetgiMulti gmdCustClass, "CodeNo", "Description", "CD004", "�N�X", "�W��", , True
        SetgiMulti gimCustStatusCode, "CodeNo", "Description", "CD035", "�N�X", "�W��"
        gimCustStatusCode.SetQueryCode "1"
        SetgiMulti gmdClctEn, "EmpNO", "EmpName", "CM003", "�N�X", "�W��", , True
        SetgiMulti gmsMduId, "Mduid", "Name", "SO017", "�N�X", "�W��"
        '#4388 �W�[������] By Kin 2009/04/30
        '#5218 �n��Nvl���覡 By Kin 2009/08/05
        SetgiMulti gimUCCode, "CodeNo", "Description", "CD013", "�N�X", "�W��", " Where Nvl(REFNO,0) NOT IN(3,7) AND NVL(PAYOK,0)=0 ", True
        If gmdBillType.GetQueryCode = "" Then
            SetgiMultiAddItem gmdBillType, "B,T,I,M,P", "���O��,�{�ɦ��O��,�˾���,���׳�,�������", "�N�X", "�W��"
        End If
        gmdBillType.SetQueryCode ("B,T")
        '���O����
        SetgiMulti gimCitemCode, "CodeNo", "Description", "CD019", "�N�X", "�W��", , True
        gmdPayType.Clear
        '#5683 �W�[ú�I���O By Kin 2010/08/11
        SetgiMulti gmdPayType, "CodeNo", "Description", "CD112", "�N�X", "�W��"
        'Call SetgiMultiAddItem(gmdPayType, "0,1", "�w�I��,�{�I��", "�N�X", "�W��")
        If GetPaynowFlag Then
            '#5925 ú�O���O�W�[�w�]�� By Kin 2011/04/12
            With gmdPayType
                Select Case Val(GetRsValue("SELECT NVL(PayKindDefault,0) FROM " & GetOwner & "SO041", gcnGi) & "")
                    Case 0
                        .SelectAll
                    Case 1
                        .SetQueryCode "0"
                    Case 2
                        .SetQueryCode "1"
                End Select
            End With
        Else
            gmdPayType.Clear
            gmdPayType.Enabled = False
        End If
Exit Sub
chkErr:
    Call ErrSub(Me.Name, "subGim")
End Sub
Private Function GetPaynowFlag() As Boolean
  On Error Resume Next
     GetPaynowFlag = Val(GetRsValue("SELECT PaynowFlag FROM " & GetOwner & "SO041") & "") = 1
End Function
Private Function IsDataOk() As Boolean
On Error GoTo chkErr
        If Not ChkDTok Then Exit Function
        If Not MustExist(gilCompCode, 2, "���q�O") Then Exit Function
        If gimBankId.GetDispStr = "" Then Call MsgMustBe("�Ȧ�O"): gimBankId.SetFocus: Exit Function
        If Not MustExist(gdaShouldDate1, 1, "���O����_�l��") Then Exit Function
        If Not MustExist(gdaShouldDate2, 1, "���O����I���") Then Exit Function
        If DateDiff("d", gdaShouldDate1.Text, gdaShouldDate2.Text) < 0 Then Call MsgDateRangeX("�������"): Exit Function
        
        IsDataOk = True
    Exit Function
chkErr:
    Call ErrSub(Me.Name, "IsDataOk")
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo chkErr
        Call CloseRecordset(rsBillHeadFmt)
        Call CloseRecordset(rsBankData)
    Exit Sub
chkErr:
    Call ErrSub(Me.Name, "Form_QueryUnload")
End Sub


Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
        ReleaseCOM frmSO3277A

End Sub

Private Sub gdaShouldDate1_Validate(Cancel As Boolean)
    On Error GoTo chkErr
        If Not IsDate(gdaShouldDate1.Text) Then Exit Sub
        If gdaShouldDate1.GetValue <> "" Then
            If gdaShouldDate2.GetValue = "" Then gdaShouldDate2.SetValue GetLastDate(gdaShouldDate1.GetValue(True))
        End If
    Exit Sub
chkErr:
    Call ErrSub(Me.Name, "gdaShouldDate1_Validate")
End Sub

Private Sub gdaShouldDate2_GotFocus()
    On Error Resume Next
        If gdaShouldDate2.GetValue <> "" Then Exit Sub
        gdaShouldDate2.SetValue GetLastDate(Date)
End Sub

Private Sub gdaShouldDate2_Validate(Cancel As Boolean)
    On Error Resume Next
        If gdaShouldDate1.GetValue = "" Or gdaShouldDate2.GetValue = "" Then Exit Sub
        If Not IsDate(gdaShouldDate2.Text) Then Exit Sub
        If DateDiff("d", gdaShouldDate1.Text, gdaShouldDate2.Text) < 0 Then
            MsgBox "���O�I��餣�o�p�󦬶O�_�l��I", vbExclamation, "�T���I"
            gdaShouldDate2.SetValue gdaShouldDate1.GetValue
            Cancel = True
        End If
End Sub

Private Sub gilCompCode_Change()
    On Error Resume Next
        garyGi(16) = gilCompCode.GetOwner
        Call subGim
        Call subGil
        Call GiMultiFilter(gmdAreaCode, , gilCompCode.GetCodeNo)
        Call GiMultiFilter(gmsMduId, , gilCompCode.GetCodeNo)
        Call GiMultiFilter(gmdServCode, , gilCompCode.GetCodeNo)
       If SetACHbankTypeCbo = False Then cmdOK.Enabled = False

End Sub

Public Sub subAnd(strCH As String)
  On Error GoTo chkErr
    If strChoose <> "" Then
        strChoose = strChoose & " And " & strCH
    Else
        strChoose = strCH
    End If
  Exit Sub
chkErr:
    Call ErrSub(Me.Name, "subAnd")
End Sub

Public Function GetOwner(Optional strOwner As String) As String
  On Error Resume Next
    If strOwner = "" Then strOwner = garyGi(16)
    If strOwner <> "" Then GetOwner = strOwner & "."
End Function

Private Sub txtCustId_KeyPress(keyAscii As Integer)
  On Error Resume Next
    If keyAscii >= 48 And keyAscii <= 57 Or keyAscii = 44 Or keyAscii = 8 Or keyAscii = 45 Then
        If keyAscii = 44 Or keyAscii = 45 Then
            If Asc(Right(" " & txtCustId.Text, 1)) = 44 Or Asc(Right(" " & txtCustId.Text, 1)) = 45 Or txtCustId = "" Then keyAscii = 9
        End If
        If Not ChkMaxLengthOK(txtCustId) Then
           If Not (keyAscii = 44 Or keyAscii = 8 Or keyAscii = 45) Then keyAscii = 9
        End If
    Else
        keyAscii = 9
    End If
End Sub
Private Function ChkMaxLengthOK(objText As Object) As Boolean
  On Error GoTo chkErr
    Dim strChar As String
    Dim strChar2 As String
    Dim intLength  As Integer
    Dim intLength2 As Integer
        ChkMaxLengthOK = True
        strChar = objText.Text
        Do
            intLength = InStr(1, objText.Text, ",", vbTextCompare)
            strChar = Mid(strChar, intLength + 1)
            If InStr(1, strChar, ",", vbTextCompare) = 0 Then
                If Len(Trim(strChar)) >= 8 Then
                  strChar2 = strChar
                  Do
                    intLength2 = InStr(1, strChar2, "-", vbTextCompare)
                    strChar2 = Mid(strChar2, intLength2 + 1)
                    If InStr(1, strChar2, "-", vbTextCompare) = 0 Then
                      If Len(Trim(strChar2)) >= 8 Then ChkMaxLengthOK = False
                      Exit Do
                    End If
                  Loop
                End If
               Exit Do
            End If
        Loop
    Exit Function
chkErr:
  Call ErrSub(Me.Name, "ChkMaxLengthOk")
End Function
Private Function TimetxtCustId(objText As Object, Optional strTable As String, Optional strField As String) As Boolean
  On Error GoTo chkErr
    Dim arrCustid()
    Dim strCustId As String
      TimetxtCustId = False
        Call ParseWord(objText.Text, arrCustid)
        strCustId = Join(arrCustid, ",")
        If strCustId = "" Then Exit Function
        If strTable = "" Then strTable = "A"
        If strField = "" Then strField = "CustId"
        If InStr(1, strCustId, ",") = 0 Then
            Call subAnd(strTable & "." & strField & " =" & strCustId)
        Else
            Call subAnd(strTable & "." & strField & " In (" & strCustId & ")")
        End If
      TimetxtCustId = True
  Exit Function
chkErr:
  Call ErrSub(Me.Name, "TimetxtCustId")
End Function
Private Sub ParseWord(ByVal strTxt As String, ary)
  On Error GoTo chkErr
    Dim lngN As Long
    Dim lngX As Long
    Dim lngZ As Long
    Dim strSub As String
    Dim lngCountAry As Long
    Dim strOut As String
    Dim lngFROM As Long
    Dim lngTo As Long
    strTxt = strTxt & IIf(Right(strTxt, 1) <> ",", ",", "")
    lngCountAry = -1
    Do While InStr(strTxt, ",") <> 0
        strSub = Left(strTxt, InStr(strTxt, ",") - 1)
        strTxt = Trim(Mid(strTxt, InStr(strTxt, ",") + 1))
        If InStr(strSub, "-") <> 0 Then
            strOut = ""
            If Len(Left(strSub, InStr(strSub, "-"))) > 8 Or Len(Mid(strSub, InStr(strSub, "-") + 1)) > 8 Then
                MsgBox "���w���d�� [" & strSub & "] �Ӥj�A�������N�|�Q���L�C", vbExclamation, "ĵ�i"
            Else
                lngFROM = Val(Left(strSub, InStr(strSub, "-")))
                lngTo = Val(Mid(strSub, InStr(strSub, "-") + 1))
                If lngTo - lngFROM > 200 Then
                    MsgBox "���w���d�� [" & strSub & "] �Ӥj�A�������N�|�Q���L�C", vbExclamation, "ĵ�i"
                Else
                    For lngX = lngFROM To lngTo
                        strOut = strOut & str(lngX) & ","
                        
                    Next
                    strTxt = strOut & strTxt
                End If
            End If
        Else
            If strTxt = "," Then
                Exit Do
            End If
            lngN = Val(strSub)
            lngCountAry = lngCountAry + 1
            ReDim Preserve ary(lngCountAry)
            ary(lngCountAry) = lngN
        End If
    Loop
  Exit Sub
chkErr:
    ErrSub Me.Name, "ParseWord"
End Sub
