VERSION 5.00
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{A0C51418-555F-11D3-BC10-0050BA01A3A9}#9.9#0"; "GiMulti.ocx"
Object = "{074771B7-6ADE-11D5-B178-00485456B028}#2.1#0"; "csMulti.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmSO3274A 
   AutoRedraw      =   -1  'True
   Caption         =   "�����঩�b/�N����Ʋ��ͧ@�~ [SO3274A]"
   ClientHeight    =   7830
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
   ScaleHeight     =   7830
   ScaleWidth      =   11190
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
      Top             =   7335
      Width           =   1275
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "F2.�T�w"
      Default         =   -1  'True
      Height          =   405
      Left            =   3270
      TabIndex        =   24
      Top             =   7335
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      Caption         =   "�d�߱���"
      ForeColor       =   &H00C00000&
      Height          =   7155
      Left            =   225
      TabIndex        =   26
      Top             =   120
      Width           =   10845
      Begin CS_Multi.CSmulti gimCitemCode 
         Height          =   315
         Left            =   210
         TabIndex        =   45
         Top             =   1650
         Width           =   10035
         _ExtentX        =   17701
         _ExtentY        =   556
         ButtonCaption   =   "�� �O �� ��"
      End
      Begin VB.TextBox txtCustId 
         Height          =   360
         Left            =   2160
         TabIndex        =   41
         Top             =   5580
         Width           =   5925
      End
      Begin VB.ComboBox cboStyle 
         Height          =   315
         ItemData        =   "SO3274A.frx":0442
         Left            =   4710
         List            =   "SO3274A.frx":044C
         Style           =   2  '��¤U�Ԧ�
         TabIndex        =   39
         Top             =   870
         Width           =   1515
      End
      Begin MSMask.MaskEdBox mskBatchNo 
         Height          =   345
         Left            =   4980
         TabIndex        =   4
         Top             =   360
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
         Caption         =   "��i�b��X�p�`�B<=0 �O�_�n����"
         Height          =   405
         Left            =   8760
         TabIndex        =   6
         Top             =   840
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
         Left            =   7440
         Style           =   2  '��¤U�Ԧ�
         TabIndex        =   35
         Top             =   360
         Visible         =   0   'False
         Width           =   3315
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
         FontSize        =   9
         FontName        =   "�s�ө���"
      End
      Begin VB.Frame Frame2 
         Caption         =   " �j�Ө̾�"
         Height          =   1035
         Left            =   195
         TabIndex        =   21
         Top             =   6015
         Width           =   10185
         Begin VB.CheckBox chkErr 
            Caption         =   "�ˮֲ��X�q�l��"
            Height          =   315
            Left            =   3030
            TabIndex        =   40
            Top             =   660
            Width           =   1695
         End
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
         DIY             =   -1  'True
         FontSize        =   9
         FontName        =   "�s�ө���"
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
         FontSize        =   9
         FontName        =   "�s�ө���"
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
         FontSize        =   9
         FontName        =   "�s�ө���"
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
         FontSize        =   9
         FontName        =   "�s�ө���"
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
         FontSize        =   9
         FontName        =   "�s�ө���"
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
         Left            =   7740
         TabIndex        =   1
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
      Begin Gi_Multi.GiMulti gimCitemCodeX 
         Height          =   345
         Left            =   450
         TabIndex        =   11
         Top             =   5670
         Visible         =   0   'False
         Width           =   10035
         _ExtentX        =   17701
         _ExtentY        =   609
         ButtonCaption   =   "�� �O �� ��"
         FontSize        =   9
         FontName        =   "�s�ө���"
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
         FontSize        =   9
         FontName        =   "�s�ө���"
      End
      Begin Gi_Multi.GiMulti gmdPayType 
         Height          =   390
         Left            =   240
         TabIndex        =   44
         Top             =   5130
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   688
         ButtonCaption   =   "ú  �I  ��  �O"
         FontSize        =   9
         FontName        =   "�s�ө���"
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "���ͫȽs"
         Height          =   195
         Left            =   1290
         TabIndex        =   43
         Top             =   5640
         Width           =   855
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "�ХΡu,�v���} Exp 67,82,199"
         Height          =   195
         Left            =   8130
         TabIndex        =   42
         Top             =   5640
         Width           =   2535
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "�榡"
         ForeColor       =   &H00C000C0&
         Height          =   195
         Left            =   4170
         TabIndex        =   38
         Top             =   930
         Width           =   390
      End
      Begin VB.Label lblBatchNo 
         AutoSize        =   -1  'True
         Caption         =   "�帹"
         ForeColor       =   &H00C000C0&
         Height          =   195
         Left            =   4500
         TabIndex        =   37
         Top             =   420
         Width           =   390
      End
      Begin VB.Label lblBillHeadFmt 
         AutoSize        =   -1  'True
         Caption         =   "�h�b�Უ�ͨ̾ڳ]�w"
         ForeColor       =   &H00C000C0&
         Height          =   195
         Left            =   5655
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
Private blnPaynowFlag As Boolean
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
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
    Dim strCD019SQL As String
        If rsBillHeadFmt.RecordCount = 0 Then Exit Sub
        rsBillHeadFmt.AbsolutePosition = cboBillHeadFmt.ListIndex + 1
        strCD019SQL = "Select CodeNo From " & GetOwner & "CD019 " & _
                            " Where CodeNo In ( Select CitemCode From " & GetOwner & "CD068A " & _
                            "         Where Exists (Select * From " & GetOwner & "CD068 Where CD068.BillHeadFmt = CD068A.BillHeadFmt " & _
                                                    " And CD068.BillHeadFmt = '" & rsBillHeadFmt("BillHeadFmt") & "' )) And StopFlag <>1"
        
'        If Not GetRS(rsCD019, "Select CodeNo From " & GetOwner & "CD019 Where CodeNo In (" & rsBillHeadFmt("CitemCodeStr") & ") And StopFlag=0") Then Exit Sub
        '#7049 �ϥ�CD068A.CitemCode By Kin 2015/07/13
        If Not GetRS(rsCD019, strCD019SQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
        If rsCD019.EOF Then Exit Sub
        strCitemCode = rsCD019.GetString(, , , ",")
        strCitemCode = Left(strCitemCode, Len(strCitemCode) - 1)
        gimCitemCode.Filter = gimCitemCode.Filter
        gimCitemCode.SetQueryCode strCitemCode
        'gimCitemCode.ShowMulti
End Sub


Private Sub cboStyle_LostFocus()
    On Error Resume Next
        '#3627 ��ܷs�榡�ɡA�n�L�oRefNo=13 By Kin 2007/11/21
        Call GiMultiFilter(gimBankId, , gilCompCode.GetCodeNo)
        gimBankId.Filter = gimBankId.Filter & IIf(gimBankId.Filter = "", " Where ", " And ") & " SubStr(Upper(PrgName),1,10) = 'BANKCENTER'"
        If cboStyle.ListIndex = 1 Then
            gimBankId.Filter = gimBankId.Filter & IIf(gimBankId.Filter = "", "Where ", " And ") & " Nvl(RefNo,0)=13"
        Else
            gimBankId.Filter = gimBankId.Filter & IIf(gimBankId.Filter = "", "Where ", " And ") & " Nvl(RefNo,0)<>13"

        End If
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
  On Error GoTo chkErr
    Unload Me
  Exit Sub
chkErr:
    Call ErrSub(Me.Name, "GetBankData")
End Sub

Private Sub cmdOk_Click()
  On Error GoTo chkErr
    Dim objAction As Object
    Dim objInterface As Object
    Dim lngSecond As Long
    Dim LastTime As Single
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
            '.uRealDate = Replace(gdaCutDate.GetOriginalValue, "/", "")
            .uRealDate = Val(Replace(gdaCutDate.GetValue, "/", "")) - 19110000
            .uUpdEn = garyGi(0)
            .uUpdName = garyGi(1)
            .uBeOne = IIf(chkBeOne.Value = 1, True, False)
            .uNT = IIf(chkNT.Value = 1, True, False)
            .uGetOwner = GetOwner
            .FlowId = garyGi(17)
            .uCkZero = IIf(chkZero.Value = 1, True, False) '�t�����B�O�_����
            .uStyle = cboStyle.ListIndex           '#3571�@�ϥΦ�خ榡
            .uBatch = mskBatchNo.Text   '#3571  �帹
            .uChkProc = chkErr.Value
            Set objAction = New clsBankCenterOut
            'ReadyGoPrint
            '#7207
             NowTime = RightNow
            StartTime = timeGetTime \ 1000
            If objAction.Action(lngSecond) Then
                Call InsertToLog(lngSecond)
'                ReadyGoPrint
                Call PrintRpt(GetPrinterName(5), "SO3274A.RPT", , Me.Caption, "Select * From BankCenterOut", strChooseString, , True, "TMP0000.MDB")
            End If
            strChoose = objAction.uReturnSQL
        End With

        Set objAction = Nothing
        Set objInterface = Nothing
        Screen.MousePointer = vbDefault
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
    Dim strsql As String
        strsql = "INSERT INTO " & GetOwner & "SO097 (EMPNO,PRGNAME,RUNDATETIME,SECOND,SELECTION) " & _
                       "VALUES ('" & garyGi(0) & "','SO3274A',SysDate,'" & _
                       lngSecond & "��','" & Replace(strChooseString, "'", "") & "')"
        gcnGi.Execute strsql
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
'        Else
'            Call subAnd("SubStr(A.BillNo,7,1) In ('B','T')")
        End If
        '92/05/29 �[���O����
        '#7049 ���CD068A.CitemCode By Kin 2015/07/13
        If cboBillHeadFmt.Visible Then
            If Len(gimCitemCode.GetQryStr & "") > 0 Then
                subAnd ("A.CitemCode In (Select CitemCode From " & GetOwner & "CD068A " & _
                                                                    " Where BillHeadFmt = '" & rsBillHeadFmt("BillHeadFmt") & "')")
            End If
        Else
            If gimCitemCode.GetQryStr <> "" Then subAnd ("A.CitemCode " & gimCitemCode.GetQryStr)
        End If
        
        If gmdClctEn.GetQryStr <> "" Then Call subAnd("A.OldClctEn " & gmdClctEn.GetQryStr)
        If gmdCMCode.GetQryStr <> "" Then Call subAnd("A.CMCode " & gmdCMCode.GetQryStr)
        If gmdCustClass.GetQryStr <> "" Then Call subAnd("A.ClassCode " & gmdCustClass.GetQryStr)
        If gimUCCode.GetQryStr <> "" Then Call subAnd("A.UCCode " & gimUCCode.GetQryStr)
        '#4388 �W�[�Ƚs���� By Kin 2009/04/30
        If txtCustId.Text <> "" Then
            Call TimetxtCustId(txtCustId)
        End If
        '#5683 �W�[ú�I���O By Kin 2010/08/10
        If gmdPayType.GetQryStr <> "" Then subAnd ("A.PayKind " & gmdPayType.GetQryStr)
        '*************************************************************************************************************************
        '#3758 �H��ڦ��O,SO033.ServiceType�|�S��� By Kin 2008/01/30
        If intPara24 = 0 And gilServiceType.GetCodeNo <> "" Then subAnd ("A.ServiceType='" & gilServiceType.GetCodeNo & "'")
        '*************************************************************************************************************************
        
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
                          "��b�Ȧ�W��:" & subSpace(gimBankId.GetDispStr) & ";" & _
                          "���O�覡  :" & subSpace(gmdCMCode.GetDispStr) & ";" & _
                          "��F��      :" & subSpace(gmdAreaCode.GetDispStr) & ";" & _
                          "�A�Ȱ�      :" & subSpace(gmdServCode.GetDispStr) & ";" & _
                          "�Ȥ����O  :" & subSpace(gmdCustClass.GetDispStr) & ";" & _
                          "���O��      :" & subSpace(gmdClctEn.GetDispStr) & ";" & _
                          "������O  :" & subSpace(gmdBillType.GetDispStr) & ";" & _
                          "�j�ӦW��  :" & subSpace(gmsMduId.GetDispStr)
    Exit Sub
chkErr:
    Call ErrSub(Me.Name, "subChoose")
End Sub

Private Sub Form_Activate()
    On Error Resume Next
        Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Initialize()
  On Error GoTo chkErr
    SetgiMultiAddItem gmdBillType, "B,I,P,M,T", "���O��,�˾���,�������,���׳�,�{�ɦ��O��"
    gmdBillType.SetQueryCode "B,T"
    Exit Sub
chkErr:
  Call ErrSub(Me.Name, "Form_Initialize")
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
        '#4388 ���դ�OK,�ƥ�|Ĳ�o�⦸,�ܦ�BTBT,�åB�h�W�[IPM,�쥻�u��BT�Ӥw For Debby By Kin 2009/05/12
        

Exit Sub
chkErr:
    Call ErrSub(Me.Name, "Form_Load")
End Sub

Private Sub DefaultValue()
    On Error Resume Next
    Dim intPara23 As Integer
        gilCompCode.SetCodeNo garyGi(9)
        gilCompCode.Query_Description
        intPara24 = GetSystemParaItem("Para24", gilCompCode.GetCodeNo, gilServiceType.GetCodeNo, "SO043")
        intPara23 = Val(GetSystemParaItem("Para23", gilCompCode.GetCodeNo, gilServiceType.GetCodeNo, "SO043") & "")
        
        '*****************************************************************************************************************
        '#3627 �s�W����H�U96�s�榡,�ùL�o�Ȧ�ORefNo�O�_����13 By Kin 2007/11/21
        cboStyle.ListIndex = 0
        gimBankId.Filter = gimBankId.Filter & IIf(gimBankId.Filter = "", "Where ", " And ") & " Nvl(RefNo,0)<>13"
        mskBatchNo.Text = "001"
        '*****************************************************************************************************************
        If intPara24 = 1 Then
            lblBillHeadFmt.Visible = True
            cboBillHeadFmt.Visible = True
            '  �δC��J�b�~Lock
             '#7049 �P�_CitemCode�O�_�i�H��� By Kin 2015/07/13
'             gimCitemCode.Enabled = False
            gimCitemCode.IsReadOnly = True
            gilServiceType.SetCodeNo ""
            gilServiceType.Query_Description
            gilServiceType.Enabled = False
            gilServiceType.Visible = False
            lblServiceType.Visible = False
        Else
            lblBillHeadFmt.Visible = False
        End If
        chkErr.Enabled = gcnGi.Execute("Select CheckText From " & GetOwner & "SO041  WHERE CompCode = " & garyGi(9))(0)
       
        'gimCitemCode.Enabled = Not cboBillHeadFmt.Visible
        If chkErr.Enabled Then chkErr.Value = 1 Else chkErr.Value = 0
        blnPaynowFlag = GetPaynowFlag
End Sub

Private Sub subGil()
  On Error GoTo chkErr
    Call SetgiList(gilCompCode, "CodeNo", "Description", "CD039", , , , , , , GetCompCodeFilter(gilCompCode))
    Call SetgiList(gilServiceType, "CodeNo", "Description", "CD046")
    Call SetgiList(gilBankHand, "CodeNo", "Description", "CD018")
  Exit Sub
chkErr:
    Call ErrSub(Me.Name, "subGil")
End Sub

Private Sub subGim()
    On Error GoTo chkErr
        
        SetgiMulti gimBankId, "CodeNo", "Description", "CD018", "�N�X", "�W��", , True
        SetgiMulti gmdCMCode, "CodeNo", "Description", "CD031", "�N�X", "�W��", , True
        SetgiMulti gmdAreaCode, "CodeNo", "Description", "CD001", "�N�X", "�W��", , True
        SetgiMulti gmdServCode, "CodeNo", "Description", "CD002", "�N�X", "�W��", , True
        SetgiMulti gmdCustClass, "CodeNo", "Description", "CD004", "�N�X", "�W��", , True
        SetgiMulti gimCustStatusCode, "CodeNo", "Description", "CD035", "�N�X", "�W��"
        gimCustStatusCode.SetQueryCode "1"
        SetgiMulti gmdClctEn, "EmpNO", "EmpName", "CM003", "�N�X", "�W��", , True
        SetgiMulti gmsMduId, "Mduid", "Name", "SO017", "�N�X", "�W��"
        '#5683 �W�[ú�I���O By Kin 2010/08/10
        SetgiMulti gmdPayType, "CodeNo", "Description", "CD112", "�N�X", "�W��"
        'Call SetgiMultiAddItem(gmdPayType, "0,1", "�w�I��,�{�I��", "�N�X", "�W��")
        gmdBillType.Clear
        gmdBillType.SetQueryCode "B,T"
        '#5925 ú�O���O�W�[�w�]�� By Kin 2011/04/12
        If GetPaynowFlag Then
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
        'SetgiMultiAddItem gmdBillType, "B,T", "���O��,�{�ɦ��O��", "�N�X", "�W��"
        '���O����
        SetgiMulti gimCitemCode, "CodeNo", "Description", "CD019", "�N�X", "�W��", , True
        '������] 930130_Liga
        SetgiMulti gimUCCode, "CodeNo", "Description", "CD013", "������]�N�X", "������]�W��", " Where Nvl(REFNO,0) NOT IN(3,7) AND NVL(PAYOK,0)=0 ", True
        
Exit Sub
chkErr:
    Call ErrSub(Me.Name, "subGim")
End Sub

Private Function IsDataOk() As Boolean
On Error GoTo chkErr
        If Not ChkDTok Then Exit Function
        If intPara24 = 0 Then If Not MustExist(gilServiceType, 2, "�A�����O") Then Exit Function
        If Not MustExist(gilCompCode, 2, "���q�O") Then Exit Function
        If Not MustExist(gdaCutDate, 1, "��b���ڤ��") Then Exit Function
        If gimBankId.GetDispStr = "" Then Call MsgMustBe("�Ȧ�O"): gimBankId.SetFocus: Exit Function
        If chkBeOne And gilBankHand.GetCodeNo = "" Then MsgBox "�X�ֲ��ͻݿ�ܥN��Ȧ�", vbExclamation, gimsgPrompt: gilBankHand.SetFocus: Exit Function
        If Not MustExist(gdaShouldDate1, 1, "���O����_�l��") Then Exit Function
        If Not MustExist(gdaShouldDate2, 1, "���O����I���") Then Exit Function
        If DateDiff("d", gdaShouldDate1.Text, gdaShouldDate2.Text) < 0 Then Call MsgDateRangeX("�������"): Exit Function
        
        If cboStyle.ListIndex = 1 Then
            If Len(mskBatchNo.Text) <> 3 Then MsgBox "�帹�]�w���~�I", vbInformation, "�T��": mskBatchNo.SetFocus: Exit Function
            If gilBankHand.GetCodeNo & "" = "" Then MsgBox "�N��Ȧ欰����I", vbInformation, "�T��": gilBankHand.SetFocus: Exit Function
        End If
        
        '3571 �p�G�O���H96�榡�A�帹�����n�� By Kin 2007/10/22
        If blnStyle Then
            If Len(mskBatchNo.Text) <> 3 Then MsgBox "�帹�]�w���~ !", vbInformation, "�T��": mskBatchNo.SetFocus: Exit Function
        End If
        
        IsDataOk = True
    Exit Function
chkErr:
    Call ErrSub(Me.Name, "IsDataOk")
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo chkErr
    Exit Sub
chkErr:
    Call ErrSub(Me.Name, "Form_QueryUnload")
End Sub


Private Sub gdaCutDate_GotFocus()
    If gdaCutDate.GetValue <> "" Or gdaShouldDate2.GetValue = "" Then Exit Sub
    If Not IsDate(gdaShouldDate2.GetValue(True)) Then Exit Sub
    gdaCutDate.SetValue CDate(gdaShouldDate2.GetValue(True)) + 15
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
        'SetgiMultiAddItem gmdBillType, "B,T", "���O��,�{�ɦ��O��", "�N�X", "�W��"

End Sub

Private Sub gilServiceType_Change()
    On Error Resume Next
        Call GiMultiFilter(gmdCMCode, gilServiceType.GetCodeNo)
        Call GiMultiFilter(gmdCustClass, gilServiceType.GetCodeNo)
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
'        If gimBankId.GetQueryCode <> "" Then
'            If gcnGi.Execute("Select Count(*) From " & GetOwner & "CD018 Where " & _
'                                "RefNo=13 And CodeNo in (" & gimBankId.GetQueryCode & ")")(0) > 0 Then
'                blnStyle = True
'                Call chkBeOne_Click
'            Else
'                blnStyle = False
'            End If
'        '****************************************************************************************************
'        End If
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
    Dim strCustid As String
      TimetxtCustId = False
        Call ParseWord(objText.Text, arrCustid)
        strCustid = Join(arrCustid, ",")
        If strCustid = "" Then Exit Function
        If strTable = "" Then strTable = "A"
        If strField = "" Then strField = "CustId"
        If InStr(1, strCustid, ",") = 0 Then
            Call subAnd(strTable & "." & strField & " =" & strCustid)
        Else
            Call subAnd(strTable & "." & strField & " In (" & strCustid & ")")
        End If
      TimetxtCustId = True
  Exit Function
chkErr:
  Call ErrSub(Me.Name, "TimetxtCustId")
End Function
Public Sub ParseWord(ByVal strTxt As String, ary)
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
Public Function GetPaynowFlag() As Boolean
  On Error Resume Next
     GetPaynowFlag = Val(GetRsValue("SELECT PaynowFlag FROM " & GetOwner & "SO041") & "") = 1
End Function
