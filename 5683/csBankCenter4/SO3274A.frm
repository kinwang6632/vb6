VERSION 5.00
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{A0C51418-555F-11D3-BC10-0050BA01A3A9}#9.8#0"; "GiMulti.ocx"
Object = "{074771B7-6ADE-11D5-B178-00485456B028}#1.17#0"; "csMulti.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmSO3274A 
   AutoRedraw      =   -1  'True
   Caption         =   "�����঩�b/�N����Ʋ��ͧ@�~ [SO3274A]"
   ClientHeight    =   7155
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
   Icon            =   "SO3274A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7155
   ScaleWidth      =   11190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '�ù�����
   Begin VB.TextBox txtSQL 
      Height          =   1155
      Left            =   3090
      TabIndex        =   33
      Top             =   3330
      Visible         =   0   'False
      Width           =   5085
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "����(&X)"
      Height          =   405
      Left            =   6360
      TabIndex        =   25
      Top             =   6675
      Width           =   1275
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "F2.�T�w"
      Default         =   -1  'True
      Height          =   405
      Left            =   3270
      TabIndex        =   24
      Top             =   6675
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      Caption         =   "�d�߱���"
      ForeColor       =   &H00C00000&
      Height          =   6315
      Left            =   225
      TabIndex        =   26
      Top             =   270
      Width           =   10695
      Begin MSMask.MaskEdBox mskBatchNo 
         Height          =   345
         Left            =   4950
         TabIndex        =   4
         Top             =   870
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   609
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   3
         Mask            =   "###"
         PromptChar      =   "_"
      End
      Begin VB.CheckBox chkZero 
         Caption         =   "���O���ت��B<=0�O�_����"
         Height          =   375
         Left            =   8760
         TabIndex        =   6
         Top             =   840
         Value           =   1  '�֨�
         Width           =   1875
      End
      Begin VB.CheckBox chkNT 
         Caption         =   "�n��M��"
         Height          =   315
         Left            =   4470
         TabIndex        =   9
         Top             =   1320
         Width           =   1110
      End
      Begin VB.ComboBox cboBillHeadFmt 
         Height          =   315
         Left            =   7560
         Style           =   2  '��¤U�Ԧ�
         TabIndex        =   35
         Top             =   360
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.CheckBox chkBeOne 
         Caption         =   "�N�U�Ȧ��ƦX�ֲ���"
         Height          =   195
         Left            =   1800
         TabIndex        =   8
         Top             =   1380
         Width           =   2595
      End
      Begin VB.CheckBox chkIncludeNormal 
         Caption         =   "�]�t�@���"
         Height          =   195
         Left            =   270
         TabIndex        =   7
         Top             =   1380
         Value           =   1  '�֨�
         Width           =   1335
      End
      Begin Gi_Multi.GiMulti gimCustStatusCode 
         Height          =   360
         Left            =   210
         TabIndex        =   16
         Top             =   3300
         Width           =   10035
         _ExtentX        =   17701
         _ExtentY        =   635
         ButtonCaption   =   "��  ��  ��  �A"
      End
      Begin VB.Frame Frame2 
         Caption         =   " �j�Ө̾�"
         Height          =   1035
         Left            =   195
         TabIndex        =   21
         Top             =   5175
         Width           =   10185
         Begin VB.OptionButton optMduYes 
            Caption         =   "&A.�n����"
            Height          =   195
            Left            =   510
            TabIndex        =   22
            Top             =   720
            Value           =   -1  'True
            Width           =   1065
         End
         Begin VB.OptionButton optMduNo 
            Caption         =   "&B.������"
            Height          =   195
            Left            =   1800
            TabIndex        =   23
            Top             =   720
            Width           =   1095
         End
         Begin CS_Multi.CSmulti gmsMduId 
            Height          =   405
            Left            =   90
            TabIndex        =   32
            Top             =   240
            Width           =   10035
            _ExtentX        =   17701
            _ExtentY        =   714
            ButtonCaption   =   "�j�ӦW��"
         End
      End
      Begin Gi_Multi.GiMulti gmdBillType 
         Height          =   390
         Left            =   210
         TabIndex        =   19
         Top             =   4380
         Width           =   10035
         _ExtentX        =   17701
         _ExtentY        =   688
         ButtonCaption   =   "��  ��  ��  �O"
      End
      Begin Gi_Multi.GiMulti gmdServCode 
         Height          =   465
         Left            =   210
         TabIndex        =   15
         Top             =   2980
         Width           =   10035
         _ExtentX        =   17701
         _ExtentY        =   820
         ButtonCaption   =   "�A     ��     ��"
      End
      Begin Gi_Multi.GiMulti gmdAreaCode 
         Height          =   465
         Left            =   210
         TabIndex        =   14
         Top             =   2660
         Width           =   10035
         _ExtentX        =   17701
         _ExtentY        =   820
         ButtonCaption   =   "��     �F     ��"
      End
      Begin Gi_Multi.GiMulti gmdCMCode 
         Height          =   465
         Left            =   210
         TabIndex        =   13
         Top             =   2320
         Width           =   10035
         _ExtentX        =   17701
         _ExtentY        =   820
         ButtonCaption   =   "��  �O  ��  ��"
      End
      Begin Gi_Date.GiDate gdaShouldDate2 
         Height          =   345
         Left            =   2910
         TabIndex        =   3
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
      Begin Gi_Multi.GiMulti gimBankId 
         Height          =   465
         Left            =   210
         TabIndex        =   12
         Top             =   2000
         Width           =   10035
         _ExtentX        =   17701
         _ExtentY        =   820
         ButtonCaption   =   "��b�Ȧ�W��"
         DataType        =   2
      End
      Begin Gi_Date.GiDate gdaCutDate 
         Height          =   345
         Left            =   7545
         TabIndex        =   5
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
      Begin prjGiList.GiList gilServiceType 
         Height          =   315
         Left            =   7560
         TabIndex        =   1
         Top             =   375
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
         TabIndex        =   18
         Top             =   4020
         Width           =   10035
         _ExtentX        =   17701
         _ExtentY        =   609
         ButtonCaption   =   "��  �O  �H  ��"
      End
      Begin CS_Multi.CSmulti gmdCustClass 
         Height          =   345
         Left            =   210
         TabIndex        =   17
         Top             =   3660
         Width           =   10035
         _ExtentX        =   17701
         _ExtentY        =   609
         ButtonCaption   =   "��  ��  ��  �O"
      End
      Begin Gi_Multi.GiMulti gimCitemCode 
         Height          =   345
         Left            =   210
         TabIndex        =   11
         Top             =   1680
         Width           =   10035
         _ExtentX        =   17701
         _ExtentY        =   609
         ButtonCaption   =   "�� �O �� ��"
      End
      Begin prjGiList.GiList gilBankHand 
         Height          =   315
         Left            =   6615
         TabIndex        =   10
         Top             =   1320
         Width           =   2940
         _ExtentX        =   5186
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Gi_Multi.GiMulti gimUCCode 
         Height          =   345
         Left            =   210
         TabIndex        =   20
         Top             =   4740
         Width           =   10035
         _ExtentX        =   17701
         _ExtentY        =   609
         ButtonCaption   =   "��  ��  ��  �]"
      End
      Begin VB.Label lblBatchNo 
         AutoSize        =   -1  'True
         Caption         =   "�帹"
         ForeColor       =   &H00C000C0&
         Height          =   195
         Left            =   4470
         TabIndex        =   37
         Top             =   930
         Width           =   390
      End
      Begin VB.Label lblBillHeadFmt 
         AutoSize        =   -1  'True
         Caption         =   "�h�b�Უ�ͨ̾ڳ]�w"
         ForeColor       =   &H00C000C0&
         Height          =   195
         Left            =   5745
         TabIndex        =   36
         Top             =   420
         Width           =   1755
      End
      Begin VB.Label lblBankHand 
         AutoSize        =   -1  'True
         Caption         =   "�N��Ȧ�"
         ForeColor       =   &H00C000C0&
         Height          =   195
         Left            =   5775
         TabIndex        =   34
         Top             =   1380
         Width           =   780
      End
      Begin VB.Label lblServiceType 
         AutoSize        =   -1  'True
         Caption         =   "�A�����O"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   6270
         TabIndex        =   31
         Top             =   420
         Width           =   780
      End
      Begin VB.Label lblCompCode 
         AutoSize        =   -1  'True
         Caption         =   "���q�O"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   300
         TabIndex        =   30
         Top             =   450
         Width           =   585
      End
      Begin VB.Label lblCutDate 
         AutoSize        =   -1  'True
         Caption         =   "��b���ڤ��"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   6315
         TabIndex        =   29
         Top             =   930
         Width           =   1170
      End
      Begin VB.Label lblReadAmt 
         AutoSize        =   -1  'True
         Caption         =   "��������d��"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   270
         TabIndex        =   28
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
         Left            =   2700
         TabIndex        =   27
         Top             =   930
         Width           =   180
      End
   End
End
Attribute VB_Name = "frmSO3274A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private strChoose As String
Private strChooseString As String
Private strChoose33 As String
Dim rsBillHeadFmt As New ADODB.Recordset
Dim intPara24 As Integer
'Dim intStyle As Integer
Dim blnStyle As Boolean
Private Sub CboAddItem()
    On Error Resume Next
        If Not GetRS(rsBillHeadFmt, "Select BillHeadFmt,CitemCodeStr From " & GetOwner & "CD068") Then Exit Sub
        Do While Not rsBillHeadFmt.EOF
            cboBillHeadFmt.AddItem rsBillHeadFmt("BillHeadFmt") & ""
            rsBillHeadFmt.MoveNext
        Loop
End Sub

Private Sub cboBillHeadFmt_Click()
    On Error Resume Next
    Dim strCitemCode As String
    Dim rsCD019  As New ADODB.Recordset
        If rsBillHeadFmt.RecordCount = 0 Then Exit Sub
        rsBillHeadFmt.AbsolutePosition = cboBillHeadFmt.ListIndex + 1
        If Not GetRS(rsCD019, "Select CodeNo From " & GetOwner & "CD019 Where CodeNo In (" & rsBillHeadFmt("CitemCodeStr") & ") And StopFlag=0") Then Exit Sub
        If rsCD019.EOF Then Exit Sub
        strCitemCode = rsCD019.GetString(, , , ",")
        strCitemCode = Left(strCitemCode, Len(strCitemCode) - 1)
        gimCitemCode.Filter = gimCitemCode.Filter
        gimCitemCode.SetQueryCode strCitemCode
        'gimCitemCode.ShowMulti
End Sub


Private Sub chkBeOne_Click()
    If blnStyle Then chkBeOne.Value = 1
    If chkBeOne.Value = 1 Then
        chkNT.Value = 0
    End If

End Sub

Private Sub chkNT_Click()
    If chkNT.Value = 1 Then
        chkBeOne.Value = 0
    End If
End Sub

Private Sub cmdCancel_Click()
  On Error GoTo ChkErr
    Unload Me
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "GetBankData")
End Sub

Private Sub cmdOk_Click()
  On Error GoTo ChkErr
    Dim objAction As Object
    Dim objInterface As Object
    Dim lngSecond As Long
        If IsDataOk = False Then Exit Sub
        Set objInterface = New clsInterface
        Call subChoose          '��Where ���O
        Screen.MousePointer = vbHourglass
        With objInterface
            If gimBankId.GetSelectCount = 1 Then
                .uBankSQL = " = " & gimBankId.GetQueryCode & " "
            Else
                .uBankSQL = " In (" & gimBankId.GetQueryCode & ") "
            End If
            .uBankHand = gilBankHand.GetCodeNo
            .uChooseSQL = strChoose
            .uChoose33 = strChoose33
            .uCompCode = gilCompCode.GetCodeNo
            .uServiceType = gilServiceType.GetCodeNo
            .uErrPath = ReadGICMIS1("ErrLogPath")
            Set .ugcnGi = gcnGi
            Set .uCnn = GetTmpMdbCn
            .uPTCode = Null
            .uRealDate = Replace(gdaCutDate.GetOriginalValue, "/", "")
            .uUpdEn = garyGi(0)
            .uUpdName = garyGi(1)
            .uBeOne = IIf(chkBeOne.Value = 1, True, False)
            .uNT = IIf(chkNT.Value = 1, True, False)
            .uGetOwner = GetOwner
            .FlowId = garyGi(17)
            .uCkZero = IIf(chkZero.Value = 1, True, False) '�t�����B�O�_����
            '.uStyle = intStyle          '#3571�@�ϥΦ�خ榡
            .uBatch = mskBatchNo.Text   '#3571  �帹
            Set objAction = New clsBankCenterOut
            If objAction.Action(lngSecond) Then
                Call InsertToLog(lngSecond)
                ReadyGoPrint
                Call PrintRpt(GetPrinterName(5), "SO3274A.RPT", , Me.Caption, "Select * From BankCenterOut", , , True, "TMP0000")
            End If
            strChoose = objAction.uReturnSQL
        End With
        Set objAction = Nothing
        Set objInterface = Nothing
        Screen.MousePointer = vbDefault
    Exit Sub
ChkErr:
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
    On Error GoTo ChkErr
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
            Call subAnd("SubStr(A.BillNo,7,1) In ('B','T')")
        End If
        '92/05/29 �[���O����
        If gimCitemCode.GetQryStr <> "" Then subAnd ("A.CitemCode " & gimCitemCode.GetQryStr)
        If gmdClctEn.GetQryStr <> "" Then Call subAnd("A.OldClctEn " & gmdClctEn.GetQryStr)
        If gmdCMCode.GetQryStr <> "" Then Call subAnd("A.CMCode " & gmdCMCode.GetQryStr)
        If gmdCustClass.GetQryStr <> "" Then Call subAnd("A.ClassCode " & gmdCustClass.GetQryStr)
        If gimUCCode.GetQryStr <> "" Then Call subAnd("A.UCCode " & gimUCCode.GetQryStr)
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
        strChoose33 = strChoose
        
        If gimCustStatusCode.GetQryStr <> "" Then Call subAnd("E.CustStatusCode " & gimCustStatusCode.GetQryStr)
        
        
        strChooseString = "���q�O�@    :" & subSpace(gilCompCode.GetDescription) & ";" & _
                          "�A�����O  :" & subSpace(gilServiceType.GetDescription) & ";" & _
                          "���O���  :" & subSpace(gdaShouldDate1.GetOriginalValue) & "~" & subSpace(gdaShouldDate2.GetOriginalValue) & ";" & _
                          "��b���ڤ��:" & subSpace(gdaCutDate.GetOriginalValue) & ";" & _
                          "��b�Ȧ�W��:" & subSpace(gimBankId.GetDescStr) & ";" & _
                          "���O�覡  :" & subSpace(gmdCMCode.GetDispStr) & ";" & _
                          "��F��      :" & subSpace(gmdAreaCode.GetDispStr) & ";" & _
                          "�A�Ȱ�      :" & subSpace(gmdServCode.GetDispStr) & ";" & _
                          "�Ȥ����O  :" & subSpace(gmdCustClass.GetDispStr) & ";" & _
                          "���O��      :" & subSpace(gmdClctEn.GetDescStr) & ";" & _
                          "������O  :" & subSpace(gmdBillType.GetDescStr) & ";" & _
                          "�j�ӦW��  :" & subSpace(gmsMduId.GetDescStr)
    Exit Sub
ChkErr:
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
    On Error GoTo ChkErr
        Call subGim
        Call subGil
        Call DefaultValue
        CboAddItem
Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "Form_Load")
End Sub

Private Sub DefaultValue()
    On Error Resume Next
    Dim intPara23 As Integer
        gilCompCode.SetCodeNo garyGi(9)
        gilCompCode.Query_Description
        intPara24 = GetSystemParaItem("Para24", gilCompCode.GetCodeNo, gilServiceType.GetCodeNo, "SO043")
        intPara23 = Val(GetSystemParaItem("Para23", gilCompCode.GetCodeNo, gilServiceType.GetCodeNo, "SO043") & "")
        'cboStyle.ListIndex = 0
        mskBatchNo.Text = "001"
        If intPara24 = 1 Then
            lblBillHeadFmt.Visible = True
            cboBillHeadFmt.Visible = True
            '  �δC��J�b�~Lock
            If intPara23 = 2 Then gimCitemCode.Enabled = False
            gilServiceType.SetCodeNo ""
            gilServiceType.Query_Description
            gilServiceType.Enabled = False
            gilServiceType.Visible = False
            lblServiceType.Visible = False
        Else
            lblBillHeadFmt.Visible = False
        End If
        
End Sub

Private Sub subGil()
  On Error GoTo ChkErr
    Call SetgiList(gilCompCode, "CodeNo", "Description", "CD039", , , , , , , GetCompCodeFilter(gilCompCode))
    Call SetgiList(gilServiceType, "CodeNo", "Description", "CD046")
    Call SetgiList(gilBankHand, "CodeNo", "Description", "CD018")
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "subGil")
End Sub

Private Sub subGim()
    On Error GoTo ChkErr
        
        SetgiMulti gimBankId, "CodeNo", "Description", "CD018", "�N�X", "�W��", , True
        SetgiMulti gmdCMCode, "CodeNo", "Description", "CD031", "�N�X", "�W��", , True
        SetgiMulti gmdAreaCode, "CodeNo", "Description", "CD001", "�N�X", "�W��", , True
        SetgiMulti gmdServCode, "CodeNo", "Description", "CD002", "�N�X", "�W��", , True
        SetgiMulti gmdCustClass, "CodeNo", "Description", "CD004", "�N�X", "�W��", , True
        SetgiMulti gimCustStatusCode, "CodeNo", "Description", "CD035", "�N�X", "�W��"
        gimCustStatusCode.SetQueryCode "1"
        SetgiMulti gmdClctEn, "EmpNO", "EmpName", "CM003", "�N�X", "�W��", , True
        SetgiMulti gmsMduId, "Mduid", "Name", "SO017", "�N�X", "�W��"
        SetgiMultiAddItem gmdBillType, "B,T", "���O��,�{�ɦ��O��", "�N�X", "�W��"
        '���O����
        SetgiMulti gimCitemCode, "CodeNo", "Description", "CD019", "�N�X", "�W��", , True
        '������] 930130_Liga
        SetgiMulti gimUCCode, "CodeNo", "Description", "CD013", "������]�N�X", "������]�W��", , True
        
Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "subGim")
End Sub

Private Function IsDataOk() As Boolean
On Error GoTo ChkErr
        If Not ChkDTok Then Exit Function
        If intPara24 = 0 Then If Not MustExist(gilServiceType, 2, "�A�����O") Then Exit Function
        If Not MustExist(gilCompCode, 2, "���q�O") Then Exit Function
        If Not MustExist(gdaCutDate, 1, "��b���ڤ��") Then Exit Function
        If gimBankId.GetDispStr = "" Then Call MsgMustBe("�Ȧ�O"): gimBankId.SetFocus: Exit Function
        If chkBeOne And gilBankHand.GetCodeNo = "" Then MsgBox "�X�ֲ��ͻݿ�ܥN��Ȧ�", vbExclamation, gimsgPrompt: gilBankHand.SetFocus: Exit Function
        If Not MustExist(gdaShouldDate1, 1, "���O����_�l��") Then Exit Function
        If Not MustExist(gdaShouldDate2, 1, "���O����I���") Then Exit Function
        If DateDiff("d", gdaShouldDate1.Text, gdaShouldDate2.Text) < 0 Then Call MsgDateRangeX("�������"): Exit Function
        
        '3571 �p�G�O���H96�榡�A�帹�����n�� By Kin 2007/10/22
        If blnStyle Then
            If Len(mskBatchNo.Text) <> 3 Then MsgBox "�帹�]�w���~ !", vbInformation, "�T��": mskBatchNo.SetFocus: Exit Function
        End If
        
        IsDataOk = True
    Exit Function
ChkErr:
    Call ErrSub(Me.Name, "IsDataOk")
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo ChkErr
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "Form_QueryUnload")
End Sub


Private Sub gdaCutDate_GotFocus()
    If gdaCutDate.GetValue <> "" Or gdaShouldDate2.GetValue = "" Then Exit Sub
    If Not IsDate(gdaShouldDate2.GetValue(True)) Then Exit Sub
    gdaCutDate.SetValue CDate(gdaShouldDate2.GetValue(True)) + 15
End Sub

Private Sub gdaShouldDate1_Validate(Cancel As Boolean)
    On Error GoTo ChkErr
        If Not IsDate(gdaShouldDate1.Text) Then Exit Sub
        If gdaShouldDate1.GetValue <> "" Then
            If gdaShouldDate2.GetValue = "" Then gdaShouldDate2.SetValue GetLastDate(gdaShouldDate1.GetValue(True))
        End If
    Exit Sub
ChkErr:
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

'#3571 �P�_�D�n�Ȧ�O�_���ѦҸ�13�A�p�G���n�����H96�s�榡�y�{ By Kin 2007/10/22
Private Sub gilBankHand_Change()
'  On Error GoTo ChkErr
'    intStyle = 0
'    If gilBankHand.GetCodeNo & "" <> "" Then
'        If gcnGi.Execute("Select Nvl(RefNo,0) From " & GetOwner & _
'                            "CD018 Where CodeNo=" & gilBankHand.GetCodeNo)(0) = 13 Then
'
'            intStyle = 1
'        Else
'            intStyle = 0
'        End If
'    End If
'    Exit Sub
'ChkErr:
'    Call ErrSub(Me.Name, "gilBankHand_Change")
End Sub

Private Sub gilCompCode_Change()
    On Error Resume Next
        garyGi(16) = gilCompCode.GetOwner
        Call subGim
        Call subGil
        Call SelectServType(gilCompCode.GetCodeNo, gilServiceType)
        gilServiceType.ListIndex = 1
        'MsgBox GetOwner
        Call GiMultiFilter(gimBankId, , gilCompCode.GetCodeNo)
        gimBankId.Filter = gimBankId.Filter & IIf(gimBankId.Filter = "", " Where ", " And ") & " SubStr(Upper(PrgName),1,10) = 'BANKCENTER'"
        Call GiMultiFilter(gmdAreaCode, , gilCompCode.GetCodeNo)
        Call GiMultiFilter(gmsMduId, , gilCompCode.GetCodeNo)
        Call GiMultiFilter(gmdServCode, , gilCompCode.GetCodeNo)

End Sub

Private Sub gilServiceType_Change()
    On Error Resume Next
        Call GiMultiFilter(gmdCMCode, gilServiceType.GetCodeNo)
        Call GiMultiFilter(gmdCustClass, gilServiceType.GetCodeNo)
End Sub

Public Sub subAnd(strCh As String)
  On Error GoTo ChkErr
    If strChoose <> "" Then
        strChoose = strChoose & " And " & strCh
    Else
        strChoose = strCh
    End If
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "subAnd")
End Sub

Public Function GetOwner(Optional strOwner As String) As String
  On Error Resume Next
    If strOwner = "" Then strOwner = garyGi(16)
    If strOwner <> "" Then GetOwner = strOwner & "."
End Function

'#3571 �P�_�Ȧ�̬O�_���]�tRefNo=13����ơA�p�G���ݱNChkBeOne���� By Kin 2007/10/23
Private Sub gimBankId_Change()
    On Error Resume Next
        blnStyle = False
        If gimBankId.GetQueryCode = "" Then
            gilBankHand.Filter = "Where 1 = 0"
        Else
            Call GiListFilter(gilBankHand, , gilCompCode.GetCodeNo)
            gilBankHand.Filter = gilBankHand.Filter & IIf(gilBankHand.Filter = "", "Where ", " And ") & " CodeNo in (" & gimBankId.GetQueryCode & ")"
        End If
        
        '****************************************************************************************************
        '#3571 �W�[���H96�s�榡 By Kin 2007/10/23
        If gimBankId.GetQueryCode <> "" Then
            If gcnGi.Execute("Select Count(*) From " & GetOwner & "CD018 Where " & _
                                "RefNo=13 And CodeNo in (" & gimBankId.GetQueryCode & ")")(0) > 0 Then
                blnStyle = True
                Call chkBeOne_Click
            Else
                blnStyle = False
            End If
        '****************************************************************************************************
        End If
End Sub

'Private Sub txtBatch_Validate(Cancel As Boolean)
'    On Error Resume Next
'    If Val(txtBatch.Text) = 0 Then MsgBox "�帹�����T !", vbInformation, "�T��": Cancel = True
'End Sub

Private Sub mskBatchNo_Change()

End Sub

Private Sub mskBatchNo_Validate(Cancel As Boolean)
  On Error Resume Next
    If blnStyle Then
        If Len(mskBatchNo.Text) <> 3 Then
            MsgBox "�帹�]�w���~ !", vbInformation, "�T��"
            Cancel = True
        End If
    End If
End Sub
