VERSION 5.00
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{A0C51418-555F-11D3-BC10-0050BA01A3A9}#9.9#0"; "GiMulti.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{074771B7-6ADE-11D5-B178-00485456B028}#2.1#0"; "csMulti.ocx"
Begin VB.Form frmSO3292A 
   BorderStyle     =   1  '��u�T�w
   Caption         =   "ACH ���ڱ��v��Ƨ@�~ [SO3292A]"
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7875
   Icon            =   "SO3292A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   7875
   StartUpPosition =   2  '�ù�����
   Begin VB.CheckBox chkStop 
      Caption         =   "�O�_�N���v���Ѱ���"
      Height          =   180
      Left            =   3240
      TabIndex        =   39
      Top             =   5700
      Width           =   2055
   End
   Begin VB.TextBox txtSQL 
      Height          =   525
      Left            =   3330
      MultiLine       =   -1  'True
      ScrollBars      =   2  '�������b
      TabIndex        =   38
      Top             =   3060
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Frame Frame3 
      Height          =   1605
      Left            =   165
      TabIndex        =   26
      Top             =   90
      Width           =   7605
      Begin VB.ComboBox CmbAutoType 
         Height          =   300
         ItemData        =   "SO3292A.frx":0442
         Left            =   1020
         List            =   "SO3292A.frx":0444
         TabIndex        =   0
         Top             =   210
         Width           =   2640
      End
      Begin Gi_Date.GiDate gdaPropDate2 
         Height          =   345
         Left            =   2580
         TabIndex        =   2
         Top             =   600
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
      Begin Gi_Date.GiDate gdaPropDate1 
         Height          =   345
         Left            =   1035
         TabIndex        =   1
         Top             =   600
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
      Begin Gi_Date.GiDate gdaStopDate2 
         Height          =   345
         Left            =   2580
         TabIndex        =   4
         Top             =   1080
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
      Begin Gi_Date.GiDate gdaStopDate1 
         Height          =   345
         Left            =   1035
         TabIndex        =   3
         Top             =   1080
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
         Left            =   4605
         TabIndex        =   5
         Top             =   210
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
      Begin CS_Multi.CSmulti gimBillHeadFmt_1 
         Height          =   375
         Left            =   4230
         TabIndex        =   41
         Top             =   990
         Visible         =   0   'False
         Width           =   2850
         _ExtentX        =   5027
         _ExtentY        =   661
         ButtonCaption   =   "�h�b�Უ�ͨ̾�"
      End
      Begin Gi_Multi.GiMulti gimBankId_1 
         Height          =   375
         Left            =   4110
         TabIndex        =   43
         Top             =   630
         Visible         =   0   'False
         Width           =   3060
         _ExtentX        =   5398
         _ExtentY        =   661
         ButtonCaption   =   "��b�Ȧ�W��"
         FontSize        =   9
         FontName        =   "�s�ө���"
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "���q�O"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   3870
         TabIndex        =   37
         Top             =   270
         Width           =   585
      End
      Begin VB.Label lblstoptit 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "���Τ��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   165
         TabIndex        =   31
         Top             =   1155
         Width           =   780
      End
      Begin VB.Label lbluntiltit 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�ӽФ��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   165
         TabIndex        =   30
         Top             =   720
         Width           =   780
      End
      Begin VB.Label lblStop 
         AutoSize        =   -1  'True
         Caption         =   "��"
         Height          =   195
         Left            =   2235
         TabIndex        =   29
         Top             =   1125
         Width           =   195
      End
      Begin VB.Label lbluntil 
         AutoSize        =   -1  'True
         Caption         =   "��"
         Height          =   195
         Left            =   2250
         TabIndex        =   28
         Top             =   630
         Width           =   195
      End
      Begin VB.Label lblComp 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "���v���O"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   165
         TabIndex        =   27
         Top             =   255
         Width           =   780
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1920
      Left            =   165
      TabIndex        =   23
      Top             =   3630
      Width           =   7590
      Begin VB.TextBox txtInvoiceId 
         Height          =   345
         Left            =   5385
         TabIndex        =   13
         Top             =   630
         Width           =   1830
      End
      Begin VB.TextBox txtPutId 
         Height          =   345
         Left            =   1500
         TabIndex        =   11
         Top             =   630
         Width           =   1830
      End
      Begin VB.TextBox txtSendSpcId 
         Height          =   345
         Left            =   1500
         TabIndex        =   10
         Top             =   210
         Width           =   1830
      End
      Begin VB.TextBox txtDataPath 
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1800
         TabIndex        =   14
         ToolTipText     =   "�п�J�r�ɤ����|���ɦW�I"
         Top             =   1050
         Width           =   5145
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
         Height          =   285
         Left            =   7005
         TabIndex        =   15
         Top             =   1065
         Width           =   450
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
         Height          =   285
         Left            =   7005
         TabIndex        =   17
         Top             =   1485
         Width           =   450
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
         Height          =   315
         Left            =   1800
         TabIndex        =   16
         Top             =   1455
         Width           =   5145
      End
      Begin Gi_Date.GiDate gdaSendDate 
         Height          =   345
         Left            =   5385
         TabIndex        =   12
         Top             =   225
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
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "�e����"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4440
         TabIndex        =   36
         Top             =   315
         Width           =   765
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "�o�ʪ̲Τ@�s��(SO)"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3540
         TabIndex        =   35
         Top             =   675
         Width           =   1725
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "�o�ʦ�N��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   210
         TabIndex        =   34
         Top             =   690
         Width           =   975
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "�o�e���N��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   210
         TabIndex        =   33
         Top             =   300
         Width           =   1170
      End
      Begin VB.Label lblDatapath 
         AutoSize        =   -1  'True
         Caption         =   "����ɸ��|���ɦW"
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
         Left            =   180
         TabIndex        =   25
         Top             =   1110
         Width           =   1560
      End
      Begin VB.Label lblErrpath 
         AutoSize        =   -1  'True
         Caption         =   "���D�ɸ��|���ɦW"
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
         Left            =   180
         TabIndex        =   24
         Top             =   1485
         Width           =   1560
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1785
      Left            =   165
      TabIndex        =   21
      Top             =   1740
      Width           =   7605
      Begin VB.ComboBox cboACHbankType 
         Height          =   300
         ItemData        =   "SO3292A.frx":0446
         Left            =   1650
         List            =   "SO3292A.frx":0448
         Style           =   2  '��¤U�Ԧ�
         TabIndex        =   6
         Top             =   195
         Width           =   2730
      End
      Begin VB.TextBox txtACHTNo 
         Height          =   345
         Left            =   5490
         TabIndex        =   8
         Top             =   180
         Width           =   1830
      End
      Begin VB.ComboBox cboBillHeadFmt 
         Height          =   300
         Left            =   270
         Style           =   2  '��¤U�Ԧ�
         TabIndex        =   7
         Top             =   0
         Visible         =   0   'False
         Width           =   1725
      End
      Begin Gi_Multi.GiMulti gimCitemCode_1 
         Height          =   375
         Left            =   60
         TabIndex        =   9
         Top             =   1650
         Visible         =   0   'False
         Width           =   7260
         _ExtentX        =   12806
         _ExtentY        =   661
         ButtonCaption   =   "�� �O �� ��_1"
         DataType        =   2
         Enabled         =   0   'False
         FontSize        =   9
         FontName        =   "�s�ө���"
      End
      Begin CS_Multi.CSmulti gimCitemCode 
         Height          =   375
         Left            =   120
         TabIndex        =   40
         Top             =   1290
         Width           =   7260
         _ExtentX        =   12806
         _ExtentY        =   661
         ButtonCaption   =   "�� �O �� ��"
      End
      Begin Gi_Multi.GiMulti gimBillHeadFmt 
         Height          =   375
         Left            =   120
         TabIndex        =   42
         Top             =   930
         Width           =   7245
         _ExtentX        =   12779
         _ExtentY        =   661
         ButtonCaption   =   "�h�b�Უ�ͨ̾�"
         DataType        =   2
         FontSize        =   9
         FontName        =   "�s�ө���"
      End
      Begin CS_Multi.CSmulti gimBankId 
         Height          =   375
         Left            =   120
         TabIndex        =   44
         Top             =   570
         Width           =   7260
         _ExtentX        =   12806
         _ExtentY        =   661
         ButtonCaption   =   "��b�Ȧ�W��"
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "����N�X"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4545
         TabIndex        =   32
         Top             =   240
         Width           =   780
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "�� �� �� ��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   165
         TabIndex        =   22
         Top             =   270
         Width           =   915
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "���v���ڴ��X"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   0
      Left            =   180
      TabIndex        =   18
      Top             =   5595
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
      Height          =   465
      Left            =   6255
      TabIndex        =   20
      Top             =   5625
      Width           =   1410
   End
   Begin MSComDlg.CommonDialog cnd1 
      Left            =   5235
      Top             =   5565
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "���v���ڴ��^"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   1
      Left            =   1665
      TabIndex        =   19
      Top             =   5595
      Width           =   1410
   End
End
Attribute VB_Name = "frmSO3292A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FSO As New FileSystemObject
Private strPath As String
Private strPrgName As String
Private strChoose As String
Private strWhere As String
Private strUpdEn As String
Private strUpdName As String
Private strUpdTime As String
Private strNowTime As String
Private DataPath As TextStream
Private ErrPath As TextStream
Private lngErrCount As Long
Private intSeq As Integer
Private strAchCid As String
Private rsTmp As New ADODB.Recordset
Dim rsBillHeadFmt As New ADODB.Recordset
Dim rsSO106 As New ADODB.Recordset
Dim strReturnSQL As String
Dim strDate As String
Dim strFrom  As String
Dim strErrAuth As String
Dim strINCD008Where As String
Dim blnSnactionNotNull As Boolean
Dim strACHTNO As String
Dim strACHDesc As String
Dim strBillHeadFmt As String
Dim strInACHTNO As String
Dim strInBillHeadFmt As String
Dim strInACHDesc As String
Dim strCancelWhere As String
Private blnHasUpd As Boolean
Private blnStopAll As Boolean
Private strNewRowId As String
Private strMasterId As String
Private blnCopy As Boolean
Private strUpdRowIds As String

Private Sub cboACHbankType_Click()
  Select Case cboACHbankType.ItemData(cboACHbankType.ListIndex)
       Case 0
'           e_CreditCardType = CardType_default
'           chkChinaBank.Visible = True
'           chkChinaBank.Value = 0
'           fraTCB.Visible = True
            
           Call SetgiMulti(gimBankId, "CodeNo", "Description", "CD018", "�N�X", "�W��", _
                                   " Where UPPER(PrgName) = 'ACHTRANREFER'  AND  COMPCODE =" & _
                                   gilCompCode.GetCodeNo & " And StopFlag=0")
    End Select
End Sub

Private Sub cboBillHeadFmt_Click()
    On Error Resume Next
    Dim strCitemCode As String
    Dim rsCD019  As New ADODB.Recordset
        If rsBillHeadFmt.RecordCount = 0 Then Exit Sub
        rsBillHeadFmt.AbsolutePosition = cboBillHeadFmt.ListIndex + 1
        If Not GetRS(rsCD019, "Select CodeNo From " & GetOwner & "CD019 Where CodeNo In (" & rsBillHeadFmt("CitemCodeStr") & ") And StopFlag=0") Then Exit Sub
        If rsCD019.EOF Then Exit Sub
        txtACHTNo = GetRsValue("Select ACHTNO From " & GetOwner & "CD068 Where BillHeadFmt='" & cboBillHeadFmt.Text & "'") & ""
        strCitemCode = rsCD019.GetString(, , , ",")
        strCitemCode = Left(strCitemCode, Len(strCitemCode) - 1)
        gimCitemCode.Filter = gimCitemCode.Filter
        gimCitemCode.SetQueryCode strCitemCode
        'gimCitemCode.ShowMulti
End Sub

Private Sub CmbAutoType_Click()
    On Error Resume Next
    gdaPropDate1.Text = "": gdaPropDate2.Text = ""
    gdaStopDate1.Text = "": gdaStopDate2.Text = ""
    Select Case CmbAutoType.ListIndex
    Case 0
        gdaStopDate1.Enabled = False
        gdaStopDate2.Enabled = False
        lblstoptit.Enabled = False
        lblStop.Enabled = False
        gdaPropDate1.Enabled = True
        gdaPropDate2.Enabled = True
        lbluntil.Enabled = True
        lbluntiltit.Enabled = True
    Case 1
        gdaStopDate1.Enabled = True
        gdaStopDate2.Enabled = True
        lblstoptit.Enabled = True
        lblStop.Enabled = True
        gdaPropDate1.Enabled = False
        gdaPropDate2.Enabled = False
        lbluntil.Enabled = False
        lbluntiltit.Enabled = False
    Case 2
        gdaStopDate1.Enabled = False
        gdaStopDate2.Enabled = False
        lblstoptit.Enabled = False
        lblStop.Enabled = False
        gdaPropDate1.Enabled = False
        gdaPropDate2.Enabled = False
        lbluntil.Enabled = False
        lbluntiltit.Enabled = False
    Case Else
        gdaStopDate1.Enabled = True
        gdaStopDate2.Enabled = True
        lblstoptit.Enabled = True
        lblStop.Enabled = True
        gdaPropDate1.Enabled = True
        gdaPropDate2.Enabled = True
        lbluntil.Enabled = True
        lbluntiltit.Enabled = True
    End Select

End Sub

Private Sub cmdCancel_Click()
  On Error GoTo ChkErr
    Unload Me
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "cmdCancel_Click")
End Sub

Public Sub CloseFS()
    On Error Resume Next
    DataPath.Close
    ErrPath.Close
    Set FSO = Nothing
End Sub

Private Function IsDataOk(intCmd As Integer) As Boolean
On Error GoTo ChkErr
    Dim strErrMsg As String
    Dim S As String
    Dim Z As String
    IsDataOk = False
    Select Case intCmd
    Case 0
        If CmbAutoType = "" Then strErrMsg = "���v���O": GoTo Warning
        If Not MustExist(gilCompCode, 2, "���q�O") Then Exit Function
        If cboACHbankType = "" Then strErrMsg = "�ӿ�Ȧ�": GoTo Warning
        If gimBankId.GetQueryCode = "" Then strErrMsg = "��b�Ȧ�W��": GoTo Warning
        'If cboBillHeadFmt = "" Then strErrMsg = "�h�b�Უ�ͨ̾ڳ]�w": GoTo Warning
        'If txtACHTNo.Text = "" Then strErrMsg = "����N�X": txtACHTNo.SetFocus: GoTo Warning
        Select Case CmbAutoType.ListIndex
        Case 0
            If gdaPropDate1.GetValue = "" Then strErrMsg = "�ӽФ��": gdaPropDate1.SetFocus: GoTo Warning
            If gdaPropDate2.GetValue = "" Then strErrMsg = "�ӽФ��": gdaPropDate2.SetFocus: GoTo Warning
        Case 1
            If gdaStopDate1.GetValue = "" Then strErrMsg = "���Τ��": gdaStopDate1.SetFocus: GoTo Warning
            If gdaStopDate2.GetValue = "" Then strErrMsg = "���Τ��": gdaStopDate2.SetFocus: GoTo Warning
        Case 2
        
        End Select
    Case 1
    
    End Select
'    If gdaHandleDate.GetValue = "" Then strErrMsg = "��X���": gdaHandleDate.SetFocus: GoTo Warning
    If txtDataPath.Text = "" Then strErrMsg = "����ɦ�m": txtDataPath.SetFocus: GoTo Warning
    If txtErrPath.Text = "" Then strErrMsg = "����ɦ�m": txtErrPath.SetFocus: GoTo Warning
    S = Left(txtDataPath, InStrRev(txtDataPath, "\"))
    If S = "C:\" Or S = "c:\" Then
    Else
        If Not ChkDirExist(S) Then MsgBox "���| " & S & " ���s�b!", vbExclamation: Exit Function
    End If
    Z = Left(txtErrPath, InStrRev(txtErrPath, "\"))
    If Z = "C:\" Or Z = "c:\" Then
    Else
        If Not ChkDirExist(Z) Then MsgBox "���| " & Z & " ���s�b!", vbExclamation: Exit Function
    End If
    IsDataOk = True
    Exit Function
Warning:
    Call MsgMustBe(strErrMsg)
    Exit Function
ChkErr:
    Call ErrSub(Me.Name, "IsDataOk")
End Function

Private Sub cmdDataPath_Click()
        On Error GoTo ChkErr
        With cnd1
                .FileName = txtDataPath.Text
                .Filter = "��r�� (*.txt;*.dat)|*.txt;*.dat|�Ҧ��ɮ� (*.*)|*.*"
                .InitDir = ReadGICMIS1("ErrLogPath")
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
                .FileName = txtDataPath.Text
                .Filter = "��r�� (*.txt;*.dat)|*.txt;*.dat|�Ҧ��ɮ� (*.*)|*.*"
                .InitDir = ReadGICMIS1("ErrLogPath")
                .Action = 1
                txtErrPath = .FileName
        End With
    Exit Sub
    
ChkErr:
    Call ErrSub(Me.Name, "cmdErrPath_Click")
End Sub
Private Function OpenData(intCmd As Integer) As Boolean
    On Error GoTo ChkErr
    If intCmd = 0 Then
        If OpenFile(DataPath, txtDataPath, True) = False Then Exit Function
    Else
        If OpenFile(DataPath, txtDataPath, True, giOpenTEXT) = False Then Exit Function
    End If
        If OpenFile(ErrPath, txtErrPath, True) = False Then Exit Function
        OpenData = True
    Exit Function
ChkErr:
    Call ErrSub(Me.Name, "OpenData")
End Function

Private Sub cmdOk_Click(intCmd As Integer)
    On Error GoTo ChkErr
    Dim rsCD068 As New ADODB.Recordset
    Dim i As Long
    Dim aryACHTNO() As String
    strInACHDesc = Empty
    strInACHTNO = Empty
    strInBillHeadFmt = Empty
    If Not IsDataOk(intCmd) Then Exit Sub
'    If Not OpenData(intCmd) Then Exit Sub
    Screen.MousePointer = vbHourglass
    '*************************************************************************************************
    '#3728 ACH��Ѧh�����X�P���^ By Kin 2008/03/10
    If gimBillHeadFmt.GetDescStr = Empty Or gimBillHeadFmt.GetDescStr = "(����)" Then
        If txtACHTNo.Text <> Empty Then
            aryACHTNO = Split(txtACHTNo.Text, ",")
            For i = LBound(aryACHTNO) To UBound(aryACHTNO)
                strInACHTNO = strInACHTNO & "'" & aryACHTNO(i) & "',"
            Next i
            
        '#3946 CD068�n�L�o ACHTNO Is Not Null������ By Kin 2008/05/30
        Else
            '#4094 �L�o���έn���� By Kin 2008/09/17
            '#4106 �W�[�P�_ACHTYPE�Ѽ� BY Kin 2008/09/22
            strInACHTNO = gcnGi.Execute("Select ''''||ACHTNO||'''' From " & GetOwner & "CD068 Where ACHTNO is Not Null And ACHTDesc Is Not Null And ACHType=1").GetString(, , , ",")
            'strInACHTNO = gcnGi.Execute("Select ''''||ACHTNO||'''' From " & GetOwner & "CD068 Where StopFlag<>1 And ACHTNO is Not Null And ACHTDesc Is Not Null").GetString(, , , ",")
        End If
        '#4094 �L�o���έn���� By Kin 2008/09/17
        'strInACHDesc = gcnGi.Execute("Select ''''||ACHTDESC||'''' From " & GetOwner & "CD068 Where StopFlag<>1 And ACHTNO is Not Null").GetString(, , , ",")
        'strInBillHeadFmt = gcnGi.Execute("Select ''''||BillHeadFmt||'''' From " & GetOwner & "CD068 Where StopFlag<>1 And ACHTNO is Not Null And ACHTDesc is Not Null").GetString(, , , ",")
        '#4106 �W�[�P�_ACHTYPE�Ѽ� BY Kin 2008/09/22
        strInACHDesc = gcnGi.Execute("Select ''''||ACHTDESC||'''' From " & GetOwner & "CD068 Where  ACHTNO is Not Null And ACHType=1").GetString(, , , ",")
        strInBillHeadFmt = gcnGi.Execute("Select ''''||BillHeadFmt||'''' From " & GetOwner & "CD068 Where ACHTNO is Not Null And ACHTDesc is Not Null And ACHType=1").GetString(, , , ",")
    Else
        If txtACHTNo.Text <> Empty Then
            aryACHTNO = Split(txtACHTNo.Text, ",")
            For i = LBound(aryACHTNO) To UBound(aryACHTNO)
                strInACHTNO = strInACHTNO & "'" & aryACHTNO(i) & "',"
            Next i
        Else
            '#4094 �L�o���έn���� By Kin 2008/09/17
            'strInACHTNO = gcnGi.Execute("Select ''''||ACHTNO||'''' From " & GetOwner & "CD068 Where StopFlag<>1 And BillHeadFmt In(" & gimBillHeadFmt.GetQueryCode & ")").GetString(, , , ",")
            '#4106 �W�[�P�_ACHTYPE�Ѽ� BY Kin 2008/09/22
            strInACHTNO = gcnGi.Execute("Select ''''||ACHTNO||'''' From " & GetOwner & "CD068 Where  BillHeadFmt In(" & gimBillHeadFmt.GetQueryCode & ") And ACHType=1").GetString(, , , ",")
        End If
        '#4094 �L�o���έn���� By Kin 2008/09/17
'        strInACHDesc = gcnGi.Execute("Select ''''||ACHTDESC||'''' From " & GetOwner & "CD068 Where StopFlag<>1 And BillHeadFmt In(" & gimBillHeadFmt.GetQueryCode & ")").GetString(, , , ",")
'        strInBillHeadFmt = gcnGi.Execute("Select ''''||BillHeadFmt||'''' From " & GetOwner & "CD068 Where StopFlag<>1 And BillHeadFmt In(" & gimBillHeadFmt.GetQueryCode & ")").GetString(, , , ",")
        '#4094 �쥻��BillHeadFmt �{�b�אּACHTDESC By Kin 2008/09/17
        '#4106 �W�[�P�_ACHTYPE�Ѽ� BY Kin 2008/09/22
        strInACHDesc = gcnGi.Execute("Select ''''||ACHTDESC||'''' From " & GetOwner & "CD068 Where  ACHTDESC In(" & gimBillHeadFmt.GetQueryCode & ") And ACHType=1").GetString(, , , ",")
        strInBillHeadFmt = gcnGi.Execute("Select ''''||BillHeadFmt||'''' From " & GetOwner & "CD068 Where  ACHTDESC In(" & gimBillHeadFmt.GetQueryCode & ") And ACHType=1").GetString(, , , ",")
    End If
    strInACHTNO = Left(strInACHTNO, Len(strInACHTNO) - 1)
    strInACHDesc = Left(strInACHDesc, Len(strInACHDesc) - 1)
    strInBillHeadFmt = Left(strInBillHeadFmt, Len(strInBillHeadFmt) - 1)
    '*************************************************************************************************
    Call ScrToRcd
    strACHTNO = Empty
    strACHDesc = Empty
    strBillHeadFmt = Empty
    Call subChoose          '��Where ���O
    If Not BeginTran(intCmd) Then
            CloseFS
            Screen.MousePointer = vbDefault
            cmdOK(0).Enabled = True: cmdOK(1).Enabled = True: cmdCancel.Enabled = True
            Screen.MousePointer = vbDefault
'            Exit Sub
    End If
    Screen.MousePointer = vbDefault
    On Error Resume Next
    Call CloseRecordset(rsCD068)
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "cmdOk_Click")
End Sub

Private Function AlterSO106(cn As ADODB.Connection, Optional rs As ADODB.Recordset, _
                            Optional strState As String, Optional strData As String, _
                            Optional strErrDes As String, Optional str106MasterId As String = Empty) As Boolean
        On Error GoTo ChkErr
    Dim strSQL As String
    Dim lngAffected As Long
    Dim strCitemStr As String
    Dim rsTmp As New ADODB.Recordset
    Dim strRowId As String
    Dim strReStatus As String
    Dim rsNew As New ADODB.Recordset
    '94/08/10 Jacky�� Jim ��1715 ���D
    strReStatus = GetReStatus(strData)
    strCitemStr = Replace(gimCitemCode.GetQryStr, ",", "','")
    strCitemStr = Replace(strCitemStr, "(", "('")
    strCitemStr = Replace(strCitemStr, ")", "')")
    strWhere = strWhere ' & " And CitemStr " & strCitemStr
    strUpdTime = Empty
    strNowTime = Empty
    AlterSO106 = False
    If strState = "" Then
'    strAchCid = txtACHTNo & GetString(rs("Custid") & "", 8, giRight, True)
    '���ͱ��v���X�ɫ�
    
    '***************************************************************************************************************
        '#3585 �P�_ACHCUSTID�O�_���ȡA�p�G���Ȥ�Update ACHCUSTID By Kin 2007/10/26
        '#3728 ���դ�OK,Update SO106�ɷ|�N�ۦP�Ƚs���PACH����Ƥ]�@�_Update��,�h�W�[�@��str106RowId��PK�� By Kin 2008/04/08
        '#3946 ���MasterId��PK By Kin 2008/05/30
        If rs("ACHCUSTID") & "" <> "" Then
            strSQL = "Update " & GetOwner & "SO106 A Set " & _
                    "SendDate = To_Date('" & gdaSendDate.GetValue & "','YYYYMMDD')," & _
                    "UpdEn='" & strUpdName & "',UpdTime='" & GetDTString(Now) & "' Where" & _
                    " AccountId='" & rs("AccountId") & "'" & _
                    IIf(str106MasterId <> Empty, " And A.MasterId=" & str106MasterId, "") & _
                    " And CustId=" & rs("Custid") & " And " & strWhere

        Else
            strSQL = "Update " & GetOwner & "SO106 A Set ACHCustid='" & strAchCid & _
                     "' , SendDate = To_Date('" & gdaSendDate.GetValue & "','YYYYMMDD')," & _
                     "UpdEn='" & strUpdName & "',UpdTime='" & GetDTString(Now) & "' Where" & _
                     " AccountId='" & rs("AccountId") & "'" & _
                     IIf(str106MasterId <> Empty, " And A.MasterId=" & str106MasterId, "") & _
                     " And CustId=" & rs("Custid") & " And " & strWhere
        End If
    '***************************************************************************************************************
    ElseIf strState = "E" Then
    '�Ȧ�Ǧ^�аT���D���\��
    '���D��2519  �ʸ۴� �Y�q�l�ɦ^�аT�����OR0,�h�бNSO106�� SENDDATE��ACHCUSID�M�� Null ---for 2006/06/26 Crystal Edit
    '���D2630   ���v���Ѱ��F�N�e�����M�Ť��~�мW�[���^�ɬO�_�N���v���Ѥ@�ְ��� ---2006/07/24 Crystal Edit
    '���D2699   �NAuthInOk���ͪ����~�T����JReAuthorizeStatus----for Jim    2006/08/29
'        strSQL = "Update " & GetOwner & "SO106 A Set SendDate='' Where" & _
                 " AccountId='" & GetAccID(strData) & "'" & " And CustId=" & GetCustID(strData) & " And " & strWhere
        '#3799 ���v���Ѱ��αb��]�n��14�X By Kin 2008/03/11
        '#3728 �P�_�O�_���������v����,�p�G���������v����,�~�i�氱��
        '#5354 ACH���v���ڴ��^�A�Y�q�l�ɥ��ѡA���n�M��SO106�Τ��ѧO�X(ACHCUSTID) By Kin 2012/09/03
        If blnStopAll Then
            If chkStop.Value = 1 Then
                strSQL = "Update " & GetOwner & "SO106 A Set SendDate='',ACHCustId=ACHCustId,UpdEn='" & strUpdName & "',UpdTime='" & GetDTString(Now) & "'," & _
                         "StopFlag=1,StopDate=To_Date('" & Format(Now, "YYYYMMDD") & "','YYYYMMDD'),ReAuthorizeStatus='" & strErrDes & "' Where" & _
                         " LPAD(AccountId,14,'0')='" & GetString(Mid(strData, 27, 14), 14, giRight, True) & "'" & " And ACHCustId='" & GetACHCustID(strData) & "' And SnactionDate is Null And nvl(StopFlag,0) = 0 "
                         
                '" AccountId='" & GetAccID(strData) & "'" & " And ACHCustId=" & GetACHCustID(strData) & " And SnactionDate is Null And nvl(StopFlag,0) = 0 "
            Else
                strSQL = "Update " & GetOwner & "SO106 A Set SendDate='',ACHCustId=ACHCustId,UpdEn='" & strUpdName & "',UpdTime='" & GetDTString(Now) & "',ReAuthorizeStatus='" & strErrDes & "' Where" & _
                         " LPAD(AccountId,14,'0')='" & GetString(Mid(strData, 27, 14), 14, giRight, True) & "'" & " And ACHCustId='" & GetACHCustID(strData) & "' And SnactionDate is Null And nvl(StopFlag,0) = 0 "
            End If
        End If
    ElseIf strState = "D" Then
    '�^���ɬ�D��(�������v)
        '#3728 �P�_�O�_���������Υ��������v���� By Kin 2008/03/14
        If blnStopAll Then
            '#3806 �������v�ɱb���]�n��14�X��� By Kin 2008/03/14
            strSQL = "Update " & GetOwner & "SO106 A Set AuthorizeStatus=2,ReAuthorizeStatus='" & strReStatus & "',UpdEn='" & strUpdName & "',UpdTime='" & GetDTString(Now) & "' Where" & _
                     " ACHCustId='" & GetACHCustID(strData) & "' And LPAD(AccountId,14,'0')='" & GetString(Mid(strData, 27, 14), 14, giRight, True) & "'"
        End If
    'O��(�s�W�¦��wñ���eú����)
    ElseIf strState = "O" Then
        '�s�W�@�����
        'UPDATE�֭���=so106.�e�X����B�Ȧ�֦L���= so106.�e�X����B���z�ɶ�= so106.�e�X������v���A�AStatus=1 (�w���v���\)
        '#3676 ACH���^,�b����Ѩ���r�ɧ��㪺14�X�b�PSO106.ACCOUNTID����0��� By Kin 2007/12/12
        If Not GetRS(rsTmp, "Select A.RowId,A.* From " & GetOwner & "SO106 A Where " & _
                    " LPAD(AccountID,14,'0')='" & GetString(Mid(strData, 27, 14), 14, giRight, True) & "'" & _
                    " And ACHCustId='" & GetACHCustID(strData) & "'" & _
                    " And SnactionDate Is not null and Propdate is not null" & _
                    " And OldACH=0 And nvl(StopFlag,0) = 0", , adUseClient, adOpenKeyset, adLockOptimistic, , , True) Then Exit Function
        Set rsTmp.ActiveConnection = Nothing
        With rsTmp
            If Not .EOF Then
                strNowTime = RightNow
                strUpdTime = GetDTString(strNowTime)
                strRowId = .Fields("RowId")
                .Fields("AuthorizeStatus") = 1
                '��sso106���ַǤ��;���D��1957  --------for Crystal 95/02/22
                 .Fields("SnactionDate") = CDate(Format(strDate, "####/##/##"))
'                .Fields("SnactionDate") = .Fields("SendDate")
                .Fields("AcceptTime") = .Fields("SendDate")
                .Fields("PropDate") = .Fields("SendDate")
                .Fields("OldACH") = 1
                .Fields("ReAuthorizeStatus") = NoZero(strReStatus)
                .Fields("UpdEn") = strUpdName
                .Fields("UpdTime") = strUpdTime
                '3946 �n�h�W�[MasterId By Kin 2008/05/30
                .Fields("MasterId") = Get_MasterID_Seq
                If Not InsertToOracle("SO106", rsTmp) Then Exit Function
            End If
            
        End With
        '*******************************************************************************
        '#3728 �M��s�W��RowId,�H�K�i�H�s�W��SO106A By Kin 2008/03/14
            strNewRowId = Empty
            strNewRowId = Empty
            strSQL = "Select A.RowId,A.MasterId From " & GetOwner & "SO106 A" & _
                     " Where OldACH=1" & _
                     " And UpdEn='" & strUpdName & "'" & _
                     " And UpdTime='" & strUpdTime & "'" & _
                     " And StopFlag<>1"
            Set rsNew = gcnGi.Execute(strSQL)
            If Not rsNew.EOF Then
                strNewRowId = rsNew("RowId")
                strMasterId = rsNew("MasterId")
            End If
            strUpdRowIds = IIf(strUpdRowIds = Empty, "'" & strRowId & "'", strUpdRowIds & ",'" & strRowId & "'")
            strSQL = Empty

        '********************************************************************************
        
        
    Else
    '�^���ɬ�A��(�s�W���v)
        '��sso106���ַǤ��;���D��1957  --------for Crystal 95/02/22
        '���D��2776 ���v�^�n�ɭYso106.stopflag�@!=0�@�ɤ�����s������� ------jim 95/9/26��
        '#2699 ���դ�OK,Where����SO106.CitemStr���ݭn�ŦXCD008�̪�CitmCodeStr By Kin 2007/10/24
        
        '*****************************************************************************************************************************************************************************
        '#3676 ACH���^,�b����Ѩ���r�ɧ��㪺14�X�b�PSO106.ACCOUNTID����0��� By Kin 2007/12/12
        '#4153 �]��ACHCustID�i�୫�ƩҥH�W�[�P�_�q�l�ɪ�ACHTNO By Kin 2008/10/21
        strSQL = "Update " & GetOwner & "SO106 A Set AuthorizeStatus=1" & _
                    ",SnactionDate=To_Date('" & strDate & "','YYYYMMDD'),ReAuthorizeStatus='" & _
                    strReStatus & "',UpdEn='" & strUpdName & "',UpdTime='" & GetDTString(RightNow) & _
                "' Where" & _
                    " ACHCustId='" & GetACHCustID(strData) & "'" & _
                    " And SendDate=To_Date('" & GetTxtDate(strData) & "','YYYYMMDD')" & _
                    " And LPAD(AccountID,14,'0')='" & GetString(Mid(strData, 27, 14), 14, giRight, True) & "'" & _
                    " And nvl(StopFlag,0) = 0 And SnactionDate is Null And " & strINCD008Where & _
                    " And InStr(',' || A.ACHTNO || ',' ,chr(39)|| '" & GetString(Mid(strData, 7, 3), 3, giRight, False) & "'||chr(39))>0"
        '*****************************************************************************************************************************************************************************
 
    End If
    '#3728 �P�_�O�_�������������v�Υ������v���Ѥ~��s By Kin 2008/03/11
    Select Case strState
        Case "O"
            If strSQL <> Empty Then
                If Not ExecuteCommand(strSQL, cn, lngAffected) Then GoTo Finall
            Else
                AlterSO106 = True
                GoTo Finall
            End If
        Case "E", "D"
            If blnStopAll Then
                If Not ExecuteCommand(strSQL, cn, lngAffected) Then GoTo Finall
            Else
                AlterSO106 = True
                GoTo Finall
            End If
        Case Else
            If Not ExecuteCommand(strSQL, cn, lngAffected) Then GoTo Finall
    End Select
    

    If lngAffected <> 0 Then
        AlterSO106 = True
    Else
        If strState = "A" Then
           strSQL = "select count(*) from " & GetOwner & "SO106A where masterid in ( " & _
                        " select masterid from " & GetOwner & "SO106 A Where " & _
                        "  ACHCustId='" & GetACHCustID(strData) & "'" & _
                        " And SendDate=To_Date('" & GetTxtDate(strData) & "','YYYYMMDD') " & _
                    " And LPAD(A.AccountID,14,'0')='" & GetString(Mid(strData, 27, 14), 14, giRight, True) & "'" & _
                    " And nvl(StopFlag,0) = 0 And SnactionDate is not Null And " & strINCD008Where & _
                    " And InStr(',' || A.ACHTNO || ',' ,chr(39)|| '" & GetString(Mid(strData, 7, 3), 3, giRight, False) & "'||chr(39))>0) " & _
                    " And AuthorizeStatus is null And ACHTNO = '" & GetString(Mid(strData, 7, 3), 3, giRight, False) & "'"
           If Val(GetRsValue(strSQL) & "") = 1 Then
                AlterSO106 = True
           End If
                    
        End If
    End If
Finall:
    On Error Resume Next
    Call CloseRecordset(rsNew)
    Call CloseRecordset(rsTmp)
    Exit Function
ChkErr:
    AlterSO106 = False
End Function

Private Function GetReStatus(strData) As String
    On Error Resume Next
        Select Case Mid(strData, 106, 1)
            Case "P"
                GetReStatus = "P:�w�o�e���v�Ѥα��v������"
            Case "R"
                GetReStatus = "R:������^�аT������������v��"
            Case "Y"
                GetReStatus = "Y:������^�аT���᦬����v��"
            Case "M"
                GetReStatus = "M:��������v�Ѧ�������^�аT��"
            Case "S"
                GetReStatus = "S:��������v�ѫ᦬��^�аT��"
            Case "C"
                GetReStatus = "C:�w�����¥����ɦ^�аT��"
            Case "D"
                GetReStatus = "D:�w����������v���ڦ^�аT��"
        End Select
End Function

Private Function InsSO002A(cn As ADODB.Connection, Optional rs As ADODB.Recordset, Optional strData As String, Optional strCustid As String) As Boolean
    On Error GoTo ChkErr
    Dim strSQL As String
    Dim lngAffected As Long
    Dim bytCnt As Integer
    
    strWhere = strWhere & " And CitemStr " & gimCitemCode.GetQryStr
    
    '*********************************************************************************************************
    '#3676 ACH���^,�b����Ѩ���r�ɧ��㪺14�X�b�PSO106.ACCOUNTID����0��� By Kin 2007/12/12
    bytCnt = Val(RPxx(gcnGi.Execute("select count(*) from " & GetOwner & "so002a where " & _
                     "LPAD(AccountNo,14,'0')='" & GetString(Mid(strData, 27, 14), 14, giRight, True) & "'" & _
                     " And CustId=" & strCustid & _
                     " And CompCode=" & gilCompCode.GetCodeNo).GetString))
                       
    strSQL = "Select DISTINCT A.CUSTID,A.BANKCODE,A.BANKNAME,A.AccountID," & _
             "A.AccountName,B.ChargeAddrNo,B.ChargeAddress,B.MailAddrNo," & _
             "B.MailAddress,C.InvNo,C.InvTitle,C.InvAddress,C.InvoiceType,A.InvSeqNo " & _
             " From " & GetOwner & "SO106 A," & GetOwner & "So001 B," & GetOwner & "So002 C" & _
             " Where A.CustId=B.Custid And A.Custid=C.Custid And A.Custid=" & strCustid & _
             " And LPAD(A.AccountID,14,'0')='" & GetString(Mid(strData, 27, 14), 14, giRight, True) & "'" & _
             " And A.Compcode=" & gilCompCode.GetCodeNo
    '***********************************************************************************************************
    
    If Not GetRS(rsSO106, strSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic, , , True) Then Exit Function
    If Not rsSO106.EOF Then
       If bytCnt = 0 Then
         strSQL = "INSERT INTO " & GetOwner & "SO002A " & _
                  "(CUSTID,COMPCODE,BANKCODE,BANKNAME,ID,ACCOUNTNO," & _
                  "CHARGEADDRNO,CHARGEADDRESS," & _
                  "MAILADDRNO,MAILADDRESS,CHARGETITLE," & _
                  "INVNO,INVTITLE,INVADDRESS,INVOICETYPE)" & _
                  " VALUES (" & _
                  strCustid & "," & _
                  gilCompCode.GetCodeNo & "," & _
                  rsSO106("BankCode") & ",'" & _
                  rsSO106("BankName") & "'," & _
                  0 & ",'" & rsSO106("AccountID") & "','" & _
                  rsSO106("ChargeAddrNo") & "','" & _
                  rsSO106("ChargeAddress") & "'," & _
                  rsSO106("MailAddrNo") & ",'" & _
                  rsSO106("MailAddress") & "','" & _
                  rsSO106("AccountName") & "','" & _
                  rsSO106("InvNo") & "','" & _
                  rsSO106("InvTitle") & "','" & _
                  rsSO106("InvAddress") & "'," & _
                  rsSO106("InvoiceType") & ")"
          
        
       ElseIf bytCnt = 1 Then
         strSQL = "UPDATE " & GetOwner & "SO002A SET " & _
                 "BANKCODE=" & rsSO106("BankCode") & _
                  ",BANKNAME='" & rsSO106("BankName") & _
                  "',ID=0,ACCOUNTNO='" & rsSO106("AccountID") & _
                  "',CHARGEADDRNO=" & rsSO106("ChargeAddrNo") & _
                  ",CHARGEADDRESS='" & rsSO106("ChargeAddress") & _
                  "',MAILADDRNO=" & rsSO106("MailAddrNo") & _
                  ",MAILADDRESS='" & rsSO106("MailAddress") & _
                  "',STOPFLAG=0,STOPDATE=NULL " & _
                  " WHERE CUSTID=" & strCustid & _
                  " AND LPAD(AccountNo,14,'0')='" & GetString(Mid(strData, 27, 14), 14, giRight, True) & _
                  "' AND COMPCODE=" & gilCompCode.GetCodeNo
        End If
        If Not ExecuteCommand(strSQL, cn, lngAffected) Then Exit Function
        '#3728�o�{�����D,�s�WSO002A��,�p�GSO002AD�S����Ƥ]�n�s�W�@�� By Kin 2008/03/14
        bytCnt = gcnGi.Execute("Select Count(*) From " & GetOwner & "SO002AD" & _
                                " Where CustId=" & strCustid & _
                                " And AccountNo='" & rsSO106("AccountID") & "'" & _
                                " And COMPCODE=" & gilCompCode.GetCodeNo & _
                                " AND INVSEQNO=" & rsSO106("INVSEQNO"))(0)
        If bytCnt = 0 Then
            strSQL = "Insert Into " & GetOwner & "SO002AD " & _
                     "(AccountNo,CompCode,CustId,InvSeqNo)" & _
                     " Values(" & _
                     "'" & rsSO106("AccountId") & "'," & _
                     gilCompCode.GetCodeNo & "," & _
                     strCustid & "," & rsSO106("InvSeqNo") & ")"
            gcnGi.Execute strSQL
        End If
                                
                                
    End If
'    rsSO106.Close
    
'    End If
    InsSO002A = True
    Exit Function
ChkErr:
    InsSO002A = False
End Function

Private Function AlterSO003(cn As ADODB.Connection, Optional rs As ADODB.Recordset, Optional strData As String, Optional strCustid As String) As Boolean
    On Error GoTo ChkErr
    Dim strSQL As String
    Dim lngAffected As Long
    Dim rsSO106 As New ADODB.Recordset
    Dim strCitem As String, i As Integer
    Dim aryCitem() As String
    
    '*****************************************************************************************************************
'    strSQL = "SELECT Citemstr,AccountID,BankCode,BankName,PTCode,PTName,CMCode,CMName FROM " & GetOwner & _
        "SO106 WHERE CUSTID=" & GetCustID(strData) & " AND AccountID='" & GetAccID(strData) & "'"
    '94/11/18 Jacky ��
    '#3481 ��SO106���O���خ�,�h�W�[ACHCustId����,�H�����s���~ By Kin 2007/08/27
    '#3481 Penny�n�h�[����Nvl(StopFlag,0)=0 And SendDate By Kin 2007/08/28
    '#3676 ACH���^,�b����Ѩ���r�ɧ��㪺14�X�b�PSO106.ACCOUNTID����0��� By Kin 2007/12/12
    strSQL = "SELECT Citemstr,AccountID,BankCode,BankName,PTCode,PTName,CMCode,CMName " & _
            "FROM " & GetOwner & "SO106 WHERE CUSTID=" & strCustid & _
            " AND LPAD(AccountId,14,'0')='" & GetString(Mid(strData, 27, 14), 14, giRight, True) & "'" & _
            " And ACHCustId='" & GetACHCustID(strData) & "'" & _
            " AND NVL(StopFlag,0)=0 AND SendDate=To_Date('" & GetTxtDate(strData) & "','YYYYMMDD')"
    '*****************************************************************************************************************
    
    If txtACHTNo <> "" Then strSQL = strSQL & " And instr(ACHTNo,chr(39)||'" & txtACHTNo & "'||chr(39))>0"
    
    If Not GetRS(rsSO106, strSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic, , , True) Then Exit Function
    If rsSO106.EOF Then Exit Function
    If rsSO106("Citemstr") & "" <> "" Then
        aryCitem = Split(rsSO106("Citemstr") & "", ",")
        For i = 0 To UBound(aryCitem)
            '94/11/17 �ק�Bug?? SeqNo => CitemCode
            '#3481 PENNY�����n�h��CeaseDate By Kin 2007/08/27
            '#3481 Penny��CeaseDate����S�n�����F 2007/08/28
            strSQL = "Update " & GetOwner & "SO003 SET BANKCODE='" & _
                     rsSO106("BankCode") & _
                     "',BANKNAME='" & rsSO106("BankName") & _
                     "',ACCOUNTNO='" & rsSO106("AccountID") & _
                     "',PTCode='" & rsSO106("PTCode") & _
                     "',PTName='" & rsSO106("PTName") & _
                     "',CMCode='" & rsSO106("CMCode") & _
                     "',CMName='" & rsSO106("CMName") & _
                     "' WHERE CUSTID=" & strCustid & _
                     " AND SeqNO =" & aryCitem(i)
            If Not ExecuteCommand(strSQL, cn, lngAffected) Then Exit Function
        Next
    End If
    AlterSO003 = True
    Exit Function
ChkErr:
    Call ErrSub(Me.Name, "AlterSO003")
End Function
Private Function AlterSO0032(cn As ADODB.Connection, rsSO106A As ADODB.Recordset, Optional rs As ADODB.Recordset, Optional strData As String, Optional strCustid As String) As Boolean
    On Error GoTo ChkErr
    Dim strSQL As String
    Dim lngAffected As Long
    Dim rsSO106 As New ADODB.Recordset
    Dim strCitem As String, i As Integer
    Dim aryCitem() As String
    strSQL = "SELECT Citemstr,AccountID,BankCode,BankName,PTCode,PTName,CMCode,CMName,InvSeqNo " & _
            "FROM " & GetOwner & "SO106 WHERE CUSTID=" & strCustid & _
            " AND LPAD(AccountId,14,'0')='" & GetString(Mid(strData, 27, 14), 14, giRight, True) & "'" & _
            " And ACHCustId='" & GetACHCustID(strData) & "'" & _
            " AND NVL(StopFlag,0)=0 AND SendDate=To_Date('" & GetTxtDate(strData) & "','YYYYMMDD')"
    '*****************************************************************************************************************
    
    strSQL = strSQL & " And instr(ACHTNo,chr(39)||'" & rsSO106A("ACHTNO") & "'||chr(39))>0"
    
    If Not GetRS(rsSO106, strSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic, , , True) Then Exit Function
    If rsSO106.EOF Then Exit Function
    strSQL = "Update " & GetOwner & "SO003 Set BankCode='" & rsSO106("BankCode") & "'," & _
            "BANKNAME='" & rsSO106("BANKNAME") & "'," & _
            "ACCOUNTNO='" & rsSO106("ACCOUNTID") & "'," & _
            "PTCode='" & rsSO106("PTCode") & "'," & _
            "PTName='" & rsSO106("PTName") & "'," & _
            "CMCode='" & rsSO106("CMCode") & "'," & _
            "CMName='" & rsSO106("CMName") & "'," & _
            "InvSeqNo=" & rsSO106("InvSeqNo") & _
            " WHERE CUSTID=" & strCustid & _
            " And CitemCode In(" & rsSO106A("CitemCodeStr") & ")"
     If Not ExecuteCommand(strSQL, cn, lngAffected) Then Exit Function
    AlterSO0032 = True
    Exit Function
ChkErr:
    Call ErrSub(Me.Name, "AlterSO003")
End Function
Private Function BeginTran(intCmd As Integer) As Boolean
    On Error GoTo ChkErr
'    Dim rsTmp As New ADODB.Recordset
    Dim strBankId As String, strSQL As String
    Dim strOldBillNo As String
    Dim lngCountCustId As Long
    Dim lngTime As Long
    Dim strData As String
    Dim blnBigError As Boolean, blnTotalLog As Boolean
    Dim rsUpd As New ADODB.Recordset
    Dim rsCustid As New ADODB.Recordset
    Dim strErrDescription As String
    Dim strCustid As String
    Dim rsSO106A As New ADODB.Recordset
    Dim rsClone As New ADODB.Recordset
    Dim strQrySO106A As String
    Dim strSO106RowId As String
    Dim strSO106ARowId As String
    Dim blnUpdTypeO As Boolean
    Dim strUpdTypeO As String
    
    lngTime = Timer
    blnUpdTypeO = False
    strUpdTypeO = " And SnactionDate Is not null and Propdate is not null" & _
                  " And OldACH=0 And nvl(StopFlag,0) = 0"
    
    If intCmd = 0 Then
'        ���v���ڴ��X
        Screen.MousePointer = vbHourglass
        cmdOK(0).Enabled = False: cmdOK(1).Enabled = False: cmdCancel.Enabled = False
        If Not GetRsTmp(rsTmp) Then Exit Function

        If rsTmp.RecordCount = 0 Then
            MsgBox "�L��ơI�Э��s�ާ@�I", vbInformation, "�T���I"
            GoTo KillFile
        Else
            If Not OpenData(intCmd) Then Exit Function
            
            If Not InsertHead(rsTmp) Then GoTo KillFile
            If Not InsertDetail(rsTmp) Then GoTo KillFile
            If Not InsertFinal(rsTmp) Then GoTo KillFile
            MsgBox "�ݳB�z���Ʀ@" & intSeq + lngErrCount & "��," & vbCrLf & vbCrLf & _
            "�w������Ƶ��Ʀ@" & intSeq & "��," & vbCrLf & vbCrLf & _
            "���D���Ʀ@" & lngErrCount & "��," & vbCrLf & vbCrLf & _
            "�@��O:" & (Timer - lngTime) \ 1 & "��", vbInformation, "�T��"
'            msgResult rsTmp.RecordCount, lngErrCount, lngTime       '��ܰ��浲�G
        End If
        Screen.MousePointer = vbDefault
        cmdOK(0).Enabled = True: cmdOK(1).Enabled = True: cmdCancel.Enabled = True
        Call CloseFS
        BeginTran = True
        On Error Resume Next
        rsTmp.Close
        Set rsTmp = Nothing
    Else
'   ���v���ڴ��^
        If Not OpenData(intCmd) Then Exit Function
        gcnGi.BeginTrans
        intSeq = 0: lngErrCount = 0
        While Not DataPath.AtEndOfLine   'Ū����b���
            With DataPath
                strData = .ReadLine     'Ū���@�C���
                strSO106RowId = Empty   '�O��SO106.RowId�H�K�ݬO�_�٭n��sSO106 By Kin 2008/03/12
                If Not .AtEndOfLine Then
                  '�Y�D���ӦC�����H�U�ʧ@
                '                If Mid(strData, 1, 3) = "BOF" Then strDate = CStr(Val(Mid(strData, 10, 8)) + 19110000)
                  'strDate �����פJ�ɮ׭�����ƪ�����æ��褸�~,��sso106���ַǤ��;���D��1957  --------for Crystal 95/02/22
                    If Mid(strData, 1, 3) = "BOF" Or Mid(strData, 1, 3) = "EOF" Then strDate = CStr(Val(Mid(strData, 10, 8)) + 19110000): GoTo Nextloop
                  '107�r������h���^�Ц��ڴ��^�ɮ�
                    If Mid(strData, 107, 1) = "R" Then
                    '2699 �H ACHCustid �� Custid ,�]ACHCUSTID�H�b���ӽs�X,�ɮפ��N�S�����ͫȽs�����
                    '�ҥH�H�U�ΫȽs�ӷ�D��Ȫ��a��,����ACHCUSTID�hSO106����-----------for Liga
                    '#2699���դ�OK,���^�ɭn�h�[��AccoundID For Liga By Kin 2007/10/24
                    
                    
                        '#3676 ACH���^,�b����Ѩ���r�ɧ��㪺14�X�b�PSO106.ACCOUNTID����0��� By Kin 2007/12/12
                        'strSQL = "Select CUSTID From " & GetOwner & "SO106 Where ACHCUSTID='" & GetACHCustID(strData) & "' And AccountID='" & GetAccID(strData) & "'"
                        '#3739 �W�[RowId���A�H�K�i�H�w��SO106���Y�@����sNOTE By Kin 2008/01/22
                        '#3946 �n���MasterId��PK,�ҥH�n�h�D��MasterId��� By Kin 2008/05/30
                        '#4153 �]��ACHCustID�i�୫�ƩҥH�W�[�P�_�q�l�ɪ�ACHTNO By Kin 2008/10/21
                        '#4316 �W�[�L�oStopFlag<>1 By Kin 2009/01/07 For Taco
                        '#5027 �������v(D),���P�_StopFlag<>1 By Kin 2009/04/17
                        strSQL = "Select RowId,CUSTID,MasterId From " & GetOwner & "SO106 A Where ACHCUSTID='" & GetACHCustID(strData) & _
                                "' And LPAD(AccountID,14,'0')='" & GetString(Mid(strData, 27, 14), 14, giRight, True) & "'" & _
                                " And InStr(',' || ACHTNO || ',' ,chr(39)|| '" & GetString(Mid(strData, 7, 3), 3, giRight, False) & "'||chr(39))>0" & _
                                IIf(UCase(Left(CmbAutoType.Text, 1)) = "D", "", " And StopFlag<>1")
                        
                        '*****************************************************************************************************************
                        '#3728 �p�G�O�s�W�¦��e�U(O) SQL�n���S�O���B�z�A�]��AlertSO106�ɡA�|�s�W�@��SO106,�ҥH�|������� By Kin 2008/03/18
                        If Mid(strData, 71, 1) = "O" Then
                            strSQL = strSQL & strUpdTypeO
                        End If
                        '******************************************************************************************************************
                        
                        If Not GetRS(rsCustid, strSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic, , , True) Then Exit Function
                        If Not rsCustid.EOF Or Not rsCustid.BOF Then    '�bSO106���S��ACHCUSTID�ɷ|�XLOG
                            strCustid = rsCustid("Custid") & ""
                            Set rsClone = rsCustid.Clone
                            '*******************************************
                            '#3946 ���MasterId By Kin 2008/05/30
                            'strSO106RowId = rsClone("RowId")
                            strSO106RowId = rsClone("MasterId")
                            '*******************************************
                            '#3728 �N�۹�����SO106A��X�� By Kin 2008/03/12
                            '#3946 ���MasterId By Kin 2008/05/30
                            If Not GetRS(rsSO106A, GetSO106A(UCase(Mid(strData, 71, 1)), rsClone("MasterId"), strData), gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
                            '*************************************************************************************
                            '���Ҧ^�аT���O�_�����\
                            '#3739 �h�Ǥ@��RecordSet �p�G�ӵ��O�^�Х��ѡA�hSO106.NOTE��J���ѭ�] By Kin 2008/01/22
                            If AuthInOk(strData, blnBigError, blnTotalLog, , strCustid, rsClone) Then
                            '**************************************************************************************
                                '�^�Ф����v���O
                                If Mid(strData, 71, 1) = "A" Then
                                    If rsSO106A.RecordCount > 0 Then
                                        '�p�G���P�e�@����Ƥ��P�h�ӥH�e�{��
                                        If strSO106RowId <> strSO106ARowId Then
                                            '#3946 ���MasterId��PK By Kin 2008/05/30
                                            strSO106ARowId = rsSO106A("MasterId")
                                            If AlterSO106(gcnGi, , "A", strData) Then
                                                If InsSO002A(gcnGi, , strData, strCustid) = False Then
                                                    strErrDescription = "�����\�s�WSO002A "
                                                    Call InsertToErr(strData, strErrDescription, strCustid, GetAccID(strData))
                                                End If
                                                '���D��2218 �����\�s�WSO002A�ɤ]�n�s�WSO003 --------For Crystal by Jim 2006/03/20
    '                                            If AlterSO003(gcnGi, , strData, strCustid) = False Then
    '                                                strErrDescription = "�����\��sSO003 "
    '                                                Call InsertToErr(strData, strErrDescription, strCustid, GetAccID(strData))
    '                                            '                                GoTo Nextloop
    '                                            Else
                                                '#3728 ��sSO003��ηs��Function By Kin 2008/03/13
                                                If AlterSO0032(gcnGi, rsSO106A, , strData, strCustid) = False Then
                                                    strErrDescription = "�����\��sSO003 "
                                                    Call InsertToErr(strData, strErrDescription, strCustid, GetAccID(strData))
                                                Else
                                                    '#��sSO106A,�p�G�sSO106A�]��s���\�ɫh�N���\����+1
                                                    If UpdateSO106A(rsSO106A, "A") Then
                                                        intSeq = intSeq + 1
                                                    Else
                                                        strErrDescription = "�����\��sSO106A "
                                                        Call InsertToErr(strData, strErrDescription, strCustid, GetAccID(strData))
                                                    End If
                                                End If
                                            Else
                                                strErrDescription = "�����\��sSO106! �i�ର�ӵ���Ƥw���֭���..�нT�{�Τ��ѧO�X�B�b��B�֭����ΰe�X����O�_���~�I"
                                                Call InsertToErr(strData, strErrDescription, strCustid, GetAccID(strData))
                                            '                            GoTo Nextloop
                                            End If
                                        Else
                                            '�p�G���P�e���ۦP�h�u��sSO106A�PSO003
                                            If AlterSO0032(gcnGi, rsSO106A, , strData, strCustid) = False Then
                                                strErrDescription = "�����\��sSO003 "
                                                Call InsertToErr(strData, strErrDescription, strCustid, GetAccID(strData))
                                            Else
                                            
                                                If UpdateSO106A(rsSO106A, "A") Then
                                                    intSeq = intSeq + 1
                                                    '#3946 ���MasterId ��PK By Kin 2008/05/30
                                                    strSO106ARowId = rsSO106A("MasterId")
                                                Else
                                                    strErrDescription = "�����\��sSO106A "
                                                    Call InsertToErr(strData, strErrDescription, strCustid, GetAccID(strData))
                                                End If
    
                                            End If
                                            '#3946 ���MasterId ��PK By Kin 2008/05/30
                                            strSO106ARowId = rsSO106A("MasterId")
                                        End If
                                    Else
                                        strErrDescription = "�䤣��SO106A�������,�ӵ���ƥi��w���v�Ψ������v "
                                        Call InsertToErr(strData, strErrDescription, strCustid, GetAccID(strData))
                                    End If
                                ElseIf Mid(strData, 71, 1) = "D" Then
                                    If rsSO106A.RecordCount > 0 Then
                                        If UpdateSO106A(rsSO106A, "D") Then
                                            If AlterSO106(gcnGi, , "D", strData) Then
                                                If StopSO003(rsClone("RowId") & "", rsSO106A) Then
                                                    intSeq = intSeq + 1
                                                Else
                                                    strErrDescription = "����SO003��ƿ��~�I"
                                                    Call InsertToErr(strData, strErrDescription, strCustid, GetAccID(strData))
                                                End If
                                            Else
                                                strErrDescription = "�����\��sSO106�I�нT�{�Τ��ѧO�X�B�b��B�֭����ΰe�X����O�_���~�I"
                                                Call InsertToErr(strData, strErrDescription, strCustid, GetAccID(strData))
                                            End If
                                        Else
                                            strErrDescription = "�����\��sSO106A�I"
                                            Call InsertToErr(strData, strErrDescription, strCustid, GetAccID(strData))
                                        End If
                                    Else
                                        strErrDescription = "�䤣��SO106A�������,�ӵ���ƥi��w���v�Ψ������v "
                                        Call InsertToErr(strData, strErrDescription, strCustid, GetAccID(strData))
                                    End If
                                ElseIf Mid(strData, 71, 1) = "O" Then
                                    If rsSO106A.RecordCount > 0 Then
                                        '#3946 ���MasterId��PK By Kin 2008/05/30
                                        If strSO106RowId <> strSO106ARowId Then
                                            If blnUpdTypeO And strUpdRowIds <> Empty Then
                                                UpdateOType strUpdRowIds
'                                                strSql = "Update " & GetOwner & "SO106 A Set StopFlag=1," & _
'                                                    "Stopdate=A.SendDate,OldACH =1 Where RowId In(" & strUpdRowIds & ")"
'                                                gcnGi.Execute strSql
                                                strUpdRowIds = Empty
                                            End If
                                            strSO106ARowId = rsSO106A("MasterId")
                                            If AlterSO106(gcnGi, , "O", strData) Then
                                                blnUpdTypeO = True
                                                intSeq = intSeq + 1
                                            Else
                                                strErrDescription = "�����\��sSO106�I�нT�{�Τ��ѧO�X�B�b��B�֭����ΰe�X����O�_���~�I"
                                                Call InsertToErr(strData, strErrDescription, strCustid, GetAccID(strData))
                                            End If
    
                                        Else
                                            strSO106ARowId = rsSO106A("MasterRowId")
                                        End If
                                        If Not UpdateSO106A(rsSO106A, "O") Then
                                            strErrDescription = "�䤣��SO106A�������,�ӵ���ƥi��w���v�Ψ������v "
                                            Call InsertToErr(strData, strErrDescription, strCustid, GetAccID(strData))
                                        Else
                                            'intSeq = intSeq + 1
                                        End If
                                    Else
                                        strErrDescription = "�䤣��SO106A�������,�ӵ���ƥi��w���v�Ψ������v "
                                        Call InsertToErr(strData, strErrDescription, strCustid, GetAccID(strData))
                                    End If
                                End If
                            Else
                                Call UpdateSO106A(rsSO106A, "E")
                                '�p�G�^�аT����0�H�~����L���~�T��
                                If AlterSO106(gcnGi, , "E", strData, strErrAuth) = False Then lngErrCount = lngErrCount
                                '#3278 �p�G���v����,�h�n��sSO106�̪�ACH���ءB�g���ʶ���
                                If Mid(strData, 71, 1) = "A" Then
                                    If Not UpdateSO106(rsClone("RowId") & "", rsSO106A) Then
                                        strErrDescription = "����SO106 ACH���ؿ��~�I"
                                        Call InsertToErr(strData, strErrDescription, strCustid, GetAccID(strData))
                                        lngErrCount = lngErrCount - 1
                                    End If
                                End If
                            End If
                        Else
                            strCustid = ""
                            strErrDescription = "����ACHCustid ���s�b���Ʈw "
                            Call InsertToErr(strData, strErrDescription, strCustid, GetAccID(strData))
                        End If
                    Else
                        MsgBox "�榡���~�I�I": gcnGi.RollbackTrans: Exit Function
                    End If
                End If
            End With
    
Nextloop:
        Wend
        If blnUpdTypeO And strUpdRowIds <> Empty Then
            UpdateOType strUpdRowIds
            strUpdRowIds = Empty
'            strSql = "Update " & GetOwner & "SO106 A Set StopFlag=1," & _
'                     "Stopdate=A.SendDate,OldACH =1 Where RowId In(" & strUpdRowIds & ")"
'            gcnGi.Execute strSql
        End If
        gcnGi.CommitTrans
        Call CloseRecordset(rsCustid)
        Call CloseRecordset(rsClone)
        Call CloseRecordset(rsUpd)
        Call CloseRecordset(rsSO106A)
          MsgBox "�w������Ƶ��Ʀ@" & intSeq & "��," & vbCrLf & vbCrLf & _
          "���D���Ʀ@" & lngErrCount & "��," & vbCrLf & vbCrLf & _
          "�@��O:" & (Timer - lngTime) \ 1 & "��"
    End If
    Exit Function
KillFile:
    
'    gcnGi.RollbackTrans
'    �R���ɮ�..
'    CloseFS
    Exit Function
    
ChkErr:
    
    gcnGi.RollbackTrans
    Call ErrSub(Me.Name, "BeginTran")
End Function
Private Function UpdateOType(ByVal aRowid As String) As Boolean
  On Error GoTo ChkErr
    Dim aRet As Boolean
    Dim aSQL As String
    
    If aRowid = Empty Then
        aRet = True
    Else
        aSQL = "Update " & GetOwner & "SO106 A Set StopFlag=1," & _
                     "Stopdate=A.SendDate,OldACH =1 Where RowId In(" & aRowid & ")"
        gcnGi.Execute aSQL
        aRet = True
    End If
    UpdateOType = aRet
    Exit Function
ChkErr:
  Call ErrSub(Me.Name, "UpdateOType")
End Function


Private Function AuthInOk(strReadLine As String, _
        Optional BigError As Boolean = False, Optional TotalLog As Boolean = True, _
        Optional strBankName As String, Optional strCustid As String, _
        Optional rs As ADODB.Recordset = Nothing) As Boolean
    On Error GoTo ChkErr
    Dim rsTmp As New ADODB.Recordset
    Dim strBillNo As String
    Dim strAccID As String
    Dim strUpdSQL As String
'    strCustid = GetCustID(strReadLine)
    strAccID = GetAccID(strReadLine)
     '#4429 �W�[�Ȧ���~�^�ХN�X By Kin 2009/05/07
     Select Case UCase(Mid(strReadLine, 108, 1))
     Case "0" '                0 = ���\�s�W�Ψ������v
'          lngBankTotalAmt = lngBankTotalAmt + GetRealAmt(strReadLine)
'          lngBankTotalCount = lngBankTotalCount + 1
          strErrAuth = ""
     Case "1" '                1 =�LŲ����
          strErrAuth = " �LŲ����!! "
     Case "2" '                2 = �L���b��
          strErrAuth = "�L���b�� !! "
     Case "3" '                3 = �eú��νs���s�b
          strErrAuth = "�eú��νs���s�b !! "
     Case "4" '                4 = ��ƭ���
          strErrAuth = "��ƭ��� !! "
     Case "5" '                5 = �������s�b
          strErrAuth = "�������s�b !! "
     Case "6" '                6 = �q�l��ƻP�®Ѥ��e����
          strErrAuth = "�q�l��ƻP���v�Ѥ��e���� !! "
     Case "7" '                7 = �b��w���M
          strErrAuth = "�b��w���M !! "
     Case "8" '                8 = �LŲ���M
          strErrAuth = "�LŲ���M !! "
     Case "A"
          strErrAuth = "��������v�� !!"
     Case "B"
          strErrAuth = "�Τḹ�X���~ !!"
     Case "C"
          strErrAuth = "�R��� !!"
     Case "D"
          strErrAuth = "�������n���� !!"
     Case Else '                9 = ��L
          strErrAuth = "��L�����\��] !! "
     End Select
        '�Ȧ椧���~
     If strErrAuth <> "" Then
        '****************************************************************************************************
        '#3739 ���v���Ѱ��Φ���ܮɡA�n�N���~��]��JSO106.NOTE By Kin 2008/01/22
'        If chkStop.Value = 1 Then
'            rs.MoveFirst
'            Do While Not rs.EOF
'                strUpdSQL = "Update SO106 Set NOTE='" & strErrAuth & "' Where RowId='" & rs("RowId") & "'"
'                gcnGi.Execute strUpdSQL
'                rs.MoveNext
'            Loop
'        End If
        '****************************************************************************************************
        Call InsertToErr(strReadLine, strErrAuth, strCustid, strAccID)
        Exit Function
     End If

        AuthInOk = True
    Exit Function
ChkErr:
    ErrSub Me.Name, "AuthInOk"
End Function

Private Function InsertToErr(strReadLine As String, strErrDescription As String, strCustid As String, strAccID As String) As Boolean
    On Error Resume Next
        WriteTextLine ErrPath, " �Ƚs: " & strCustid & "; �b��:" & strAccID & "; ���ѭ�]:" & strErrDescription & ""
        lngErrCount = lngErrCount + 1
        InsertToErr = True
End Function


Private Function GetRealDateTran(strRealDate As String) As String
    On Error Resume Next
        GetRealDateTran = Val(Replace(strRealDate, "/", "")) - 19110000
End Function

Private Function InsertDetail(rs As ADODB.Recordset) As Boolean
  On Error GoTo ChkErr
    Dim strData As Variant
    Dim rsCD018 As New ADODB.Recordset
    Dim rsAchCid As New ADODB.Recordset
    Dim strSQL As String
    Dim strErrlog As String
    Dim strBankId2 As String
    Dim strSqlCid As String
    Dim GetACHCustID As Integer
    Dim blnNoAchReceptor As Boolean
    Dim rsSO106A As New ADODB.Recordset
    Dim strQrySO106A As String
    Dim strUpdSO106A As String
    intSeq = 0: lngErrCount = 0
    blnNoAchReceptor = False
    If rs.RecordCount > 0 Then
        rs.MoveFirst
    End If

    '*********************************************************************************************************************************************************
    '#3253 �P�_87-106�줸�n��ťթζ]�¬y�{ By Kin 2007/12/19
'    If cboBillHeadFmt.Text <> "" Then
'        blnNoAchReceptor = gcnGi.Execute("Select Nvl(NoAchReceptor,0) From " & strOwnerName & "CD068 Where BillHeadFmt='" & strBillHeadFmt & "'")(0) = 1
'    End If
    '*********************************************************************************************************************************************************
    
    Do While Not rs.EOF
        '**********************************************************************************************************
        '#3728 �NSO106����ƹ�����SO106A,�HSO106A��ư����X���ʧ@ By Kin 2008/03/11
        strQrySO106A = Empty
        Select Case CmbAutoType.ListIndex
            Case 0
                '#3946 ���MasterId��PK By Kin 2008/05/30
                '#3946 �쥻AuthorizeState��Null �אּ4
                strQrySO106A = "Select RowID,A.* From " & GetOwner & "SO106A A " & _
                               "Where MasterId=" & rs("MasterId") & _
                               " And ACHTNO In(" & strInACHTNO & ")" & _
                               " And ACHDesc In(" & strInACHDesc & ")" & _
                               " And RecordType=0 And AuthorizeStatus=4 " & _
                               " And StopFlag<>1 And StopDate is Null Order By ACHTNO"
                '#3946 ���MasterId��PK By Kin 2008/05/30
                '#3946 �쥻AuthorizeState��Null �אּ4
            Case 1
                strQrySO106A = "Select ROWID,A.* From " & GetOwner & "SO106A A " & _
                               "Where MasterId=" & rs("MasterId") & _
                               " And ACHTNO In(" & strInACHTNO & ")" & _
                               " And ACHDesc In(" & strInACHDesc & ")" & _
                               " And RecordType=1 And AuthorizeStatus=4 "
                               
                '#3946 ���MasterId��PK By Kin 2008/05/30
            Case 2
                strQrySO106A = "Select RowID,A.* From " & GetOwner & "SO106A A " & _
                               "Where MasterId=" & rs("MasterId") & _
                               " And ACHTNO In(" & strInACHTNO & ")" & _
                               " And ACHDesc In(" & strInACHDesc & ")" & _
                               " And AuthorizeStatus=1 Order By ACHTNO"
        End Select
        If Not GetRS(rsSO106A, strQrySO106A, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
        If rsSO106A.RecordCount > 0 Then
            rsSO106A.MoveFirst
        End If
        '**********************************************************************************************************
        blnHasUpd = False
        '#3728 �쥻Ū��SO106,�{�b�אּŪ���ŦXSO106A���󪺸�� By Kin 2008/03/11
        Do While Not rsSO106A.EOF
            blnNoAchReceptor = gcnGi.Execute("Select Nvl(NoAchReceptor,0) From " & strOwnerName & "CD068 Where ACHTNO='" & rsSO106A("ACHTNO") & "' And ACHTDESC='" & rsSO106A("ACHDesc") & "'")(0) = 1
            'If rs("Citemstr") & "" <> "" And ChkInCitem(gimCitemCode.GetQueryCode, rs("Citemstr") & "") Or Asc(CmbAutoType) = 79 Then
            If rsSO106A("CitemCodeStr") & "" <> "" And ChkInSO106ACitem(gimCitemCode.GetQueryCode, rsSO106A("CitemCodeStr") & "", IIf(gimCitemCode.GetQueryCode = Empty, True, False)) Or Asc(CmbAutoType) = 79 Then
        '        if rs("CitemStr")&""
                intSeq = intSeq + 1
                strData = GetString(intSeq, 6, giRight, True)
                '����Ǹ�(1-6)
                '#3728 ��ŪSO106A�����
                'strData = strData & GetString(txtACHTNo, 3, giLeft)
                strData = strData & GetString(rsSO106A("ACHTNO"), 3, giLeft)
                '����N��(7-9)
                strData = strData & GetString(txtInvoiceId, 10, giLeft)
                '�o�ʪ̲Τ@�s��(10-19)
                strSQL = "SELECT BankID2 FROM " & GetOwner & "CD018 WHERE CodeNo='" & rs("BankCode") & _
                         "' And StopFlag=0 AND CompCode=" & gilCompCode.GetCodeNo
                If Not GetRS(rsCD018, strSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic, , , True) Then Exit Function
                strBankId2 = rsCD018("BankID2") & ""
                strData = strData & GetString(strBankId2 & "", 7, giLeft)
                '������N�� (20 - 26)
                strData = strData & GetString(rs("AccountId") & "", 14, giRight, True)
                '�eú��b��(27-40)
                strData = strData & GetString(rs("AccountNameId") & "", 10, giLeft)
                '�eú��Τ@�s��(41-50)
                '���D��2874 �������v���ڤ���̽s�X��h�A���ݪ������SO106 ACHCustId��� 2006/12/04
                '���D��2874 ���դ�OK�A�S���P�_�۲Ū��Ƚs�A�[�J�P�_Custid=rs("Custid") 2006/12/18
                If CmbAutoType.ListIndex = 1 Then
                    '#3728 �������v��Where���󦳧��ܦh��@��strCancelWhere
                    '#3946 ��Ū��MasterId By Kin 2008/05/30
                    strSqlCid = "Select RowId,ACHCustid From " & GetOwner & "SO106 A Where " & _
                                 strWhere & strCancelWhere & " And Custid=" & rs("Custid") & _
                                " And MasterId='" & rsSO106A("MasterId") & "'"
                    If Not GetRS(rsAchCid, strSqlCid, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic, , , True) Then Exit Function
                    If IsNull(rsAchCid("ACHCustid")) Then
                        lngErrCount = lngErrCount + 1
                        strErrlog = " �Ƚs : " & rs("Custid") & " �b�� : " & rs("AccountID") & " �S��ACH�Τḹ�X"
                        WriteLineData strErrlog, ErrPath
                        GoTo GotoLoop
                    Else
                        strAchCid = rsAchCid("ACHCustid")
                    End If
                Else
                    '���D2699 SO041.ACHCustID=1�ɡA�Ш�3�X����N�X+����6�X���ڱb����̤j���Τḹ�X�Ǹ� -----FOR Liga
                    GetACHCustID = GetRsValue("Select ACHCustID From " & GetOwner & "SO041 Where SysId = '" & gilCompCode.GetCodeNo & "'")
                    If GetACHCustID = 1 Then
                        '**************************************************************************************************************************************************************************************
                        '#3585 �p�GACHCUSTID�w�g���ȡA���έ��s����ACHCUSTID By Kin 2007/10/26
                        If rs("ACHCUSTID") & "" <> "" Then
                            strAchCid = Trim(rs("ACHCUSTID") & "")
                        Else
                            '#3728 ���rsSO106A("ACHTNO")���� By Kin 2008/03/12
                            strSqlCid = "SELECT COUNT(*) AS COUNTS FROM " & strFrom & _
                                " Where ACCOUNTID='" & rs("AccountId") & "" & "' AND SUBSTR(ACHCUSTID,4,6) ='" & Right(rs("AccountId") & "", 6) & "'"
                            If Not GetRS(rsAchCid, strSqlCid, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic, , , True) Then Exit Function
                            strAchCid = rsSO106A("ACHTNO") & "" & GetString(rs("AccountId") & "", 6, giRight, True) & Right("00" & CStr(Trim(rsAchCid("Counts")) + 1), 2)
                        End If
                        '**************************************************************************************************************************************************************************************
                    Else
                        strAchCid = rsSO106A("ACHTNO") & "" & GetString(rs("Custid") & "", 8, giRight, True)
                    End If
                End If
                strData = strData & GetString(strAchCid, 20, giLeft)
                '�Τḹ�X(51-70)
                If Asc(CmbAutoType) = 65 Then
                    strData = strData & GetString("A", 1, giLeft)
                ElseIf Asc(CmbAutoType) = 68 Then
                    strData = strData & GetString("D", 1, giLeft)
                ElseIf Asc(CmbAutoType) = 79 Then
                    strData = strData & GetString("O", 1, giLeft)
                End If
                '�s�W�Ψ���(71-71)
                strData = strData & GetString(GetRealDateTran(gdaSendDate.GetValue), 8, giRight, True)
                '���(72-79)
                strData = strData & GetString(txtPutId, 7, giLeft)
                '�o�ʦ�N��(80-86)
                
                '********************************************************************************************
                '#3253 �P�_87-106�줸�O�n��ťթζ]�¬y�{ By Kin 2007/12/19
                If blnNoAchReceptor = False Then
                    '94/08/10 Jacky�� Jim ��1715 ���D
                    strData = strData & GetString(rs("ACHSN"), 12)
                    '�o�ʪ̱M�ΰ�(���v�ѧǸ�)(87-98)
                    'strData = strData & GetString(GetPostUnit(rs("ACHTNo") & ""), 3)
                    '#3728 rs("ACHTNO")�אּrsSO106A("ACHTNO")
                    strData = strData & GetString(GetPostUnit(rsSO106A("ACHTNo") & "", strInACHTNO), 3)
                    '�o�ʪ̱M�ΰ�(�����N�X)(99-101)
                    strData = strData & GetString("", 4, giLeft)
                    '�o�ʪ̱M�ΰ�(102-105)
                    strData = strData & GetString("", 1, giLeft)
                    '�o�ʪ̱M�ΰ�(������v�Ѧ^�Ъ��A)(106-106)
                Else
                    strData = strData & Space(20)
                End If
                '********************************************************************************************
                
                strData = strData & GetString("N", 1, giLeft)
                '������A(107-107)
                    strData = strData & GetString("", 1, giLeft)
                '�^�аT��(108-108)
                 strData = strData & GetString("", 12, giLeft)
                '�ƥ�(109-120)
                WriteLineData strData, DataPath
                Select Case CmbAutoType.ListIndex
                    Case 0, 2
                        If Not blnHasUpd Then
                            '#3946 ���MasterId��PK By Kin 2008/05/30
                            If AlterSO106(gcnGi, rsTmp, , , , rs("MasterId")) = False Then
                                Exit Function
                            Else
                                blnHasUpd = True    '�N��w��s�L,���Φb��s�@��SO106
                            End If
                        End If
                End Select
                '#3946 �p�G���A���ݴ��X�hAuthorizeStatus�אּNull
                If rsSO106A("AuthorizeStatus") & "" = "4" Then
                    rsSO106A("AuthorizeStatus") = Null
                    rsSO106A.Update
                End If
            Else
                '�N�ŦX������S�����w���ڶg�����O���ؤ���b��ƥXlog�H�i��USER
                lngErrCount = lngErrCount + 1
                strErrlog = " �Ƚs : " & rs("Custid") & " �b�� : " & rs("AccountID") & " �S�����w���ڶg�����O����"
                WriteLineData strErrlog, ErrPath
            End If
            rsSO106A.MoveNext
        Loop
GotoLoop:
        blnHasUpd = False
        rs.MoveNext
        DoEvents
    Loop
     InsertDetail = True
     blnHasUpd = False
     On Error Resume Next
     Call CloseRecordset(rsSO106A)
     Call CloseRecordset(rsCD018)
     Call CloseRecordset(rsAchCid)
  Exit Function
ChkErr:
'    cnn.RollbackTrans
    
    ErrSub Me.Name, "InsertDetail"
End Function

Private Function GetPostUnit(ByVal strACHTNO As String, ByVal strACHTNOstr As String) As String
    On Error Resume Next
        If InStr(1, strACHTNOstr, "'" & strACHTNO & "'") > 0 Then
            GetPostUnit = GetRsValue("Select PostUnit From " & GetOwner & "CD068 Where AchtNo = '" & strACHTNO & "'")
        End If
End Function

Private Function WriteLineData(ByVal vData As String, ByVal objFile As Object)
    On Error GoTo ChkErr
        Call objFile.WriteLine(vData)
    Exit Function
ChkErr:
    Call ErrSub("clsModule", "WriteLineData")
End Function
'#2699 ���դ�OK,�ˬdSO106.CitemStr�O�_���]�t�b���w��CD068�̭� By Kin
Private Function ChkInCD068(ByVal strCitemStr As String) As Boolean
  On Error GoTo ChkErr
    Dim rsSO003 As New ADODB.Recordset
    Dim rsCD068 As New ADODB.Recordset
    Dim strCD068SQL As String
    strCitemStr = Replace(strCitemStr, "'", "")
    If strCitemStr <> "" Then
        If Not GetRS(rsSO003, "Select CitemCode From " & GetOwner & "SO003 Where SeqNo IN (" & _
                                strCitemStr & ") And Custid=" & rsTmp("CustID")) Then Exit Function
        Do While Not rsSO003.EOF
            strCD068SQL = "Select BillHeadFmt From " & GetOwner & "CD068 Where " & _
                        "InStr(',' || CitemCodeStr || ',','," & rsSO003("CitemCode") & ",')>0"
            If Not GetRS(rsCD068, strCD068SQL) Then Exit Function
            If Not rsCD068.EOF Then
                If Trim(rsCD068("BillHeadFmt")) & "" = Trim(cboBillHeadFmt.Text) Then
                    ChkInCD068 = True
                    GoTo EndLoop
                End If
            End If
            rsSO003.MoveNext
        Loop
    End If
    ChkInCD068 = False
EndLoop:
    Call CloseRecordset(rsSO003)
    Call CloseRecordset(rsCD068)
    Exit Function
ChkErr:
    ErrSub Me.Name, "ChkInCD068"
End Function
Private Function ChkInCitem(strCode As String, strCitemStr As String) As Boolean
    On Error GoTo ChkErr
    Dim aryCitem() As String
    Dim i As Integer
    Dim rsCit As New ADODB.Recordset
    ChkInCitem = False
    If strCitemStr <> "" Then
        If Not GetRS(rsCit, "Select CitemCode From " & GetOwner & "SO003 WHERE SeqNo IN(" & strCitemStr & ")") Then Exit Function
        Do While Not rsCit.EOF
             If InStr(1, strCode, rsCit("CitemCode") & "") <> 0 Then
                 ChkInCitem = True
             End If
             rsCit.MoveNext
        Loop
    End If
'    aryCitem = Split(strCitemStr, ",")
'    For i = 0 To UBound(aryCitem)
'        If InStr(1, strCode, aryCitem(i)) <> 0 Then
'            ChkInCitem = True
'        End If
'    Next
    
    Exit Function
ChkErr:
    ErrSub Me.Name, "ChkInCitem"
End Function
'#3728 �ˬdSO106A�����O���جO�_���ŦXCD068�̪����O���� By Kin 2008/03/11
Private Function ChkInSO106ACitem(strCode As String, strCitemStr As String, ByVal blnAll As Boolean) As Boolean
    On Error GoTo ChkErr
    Dim i As Integer
    Dim strCD068 As String
    Dim aryCitemStr() As String
    Dim rsCD068Temp As New ADODB.Recordset
    'blnAll=Ture�N�����ΨS����ܦh�b��̾�,�h�n�D��CD068�̪����O����
    If blnAll Then
        'strCD068 = gcnGi.Execute("Select CitemCodeStr From " & strOwnerName & "CD068 Where BillHeadFmt In (" & strInBillHeadFmt & ") And StopFlag<>1").GetString(, , , ",")
        '#4094 �L�o���έn���� By Kin 2008/09/17
        '#4106 �W�[�P�_ACHTYPE�Ѽ� BY Kin 2008/09/22
'        strCD068 = gcnGi.Execute("Select CitemCodeStr From " & strOwnerName & "CD068 Where BillHeadFmt In (" & strInBillHeadFmt & ") And ACHType=1").GetString(, , , ",")
'        If Len(strCD068) > 0 Then
'            strCD068 = Left(strCD068, Len(strCD068) - 1)
'            strCode = strCD068
'        End If
        '#7049 ���CD068A.CitemCode By Kin 2015/07/08
        If Not GetRS(rsCD068Temp, "Select distinct CitemCode From " & strOwnerName & "CD068A " & _
                        " Where Exists (Select * From " & strOwnerName & "CD068 Where CD068.BillHeadFmt = CD068A.BillHeadFmt " & _
                                            " And CD068.ACHType = 1 And BillHeadFmt In (" & strInBillHeadFmt & "))", gcnGi, adUseClient, adOpenKeyset) Then Exit Function
        strCD068 = rsCD068Temp.GetString(, , , ",")
        strCD068 = Left(strCD068, Len(strCD068) - 1)
        strCode = strCD068
    End If
    ChkInSO106ACitem = False
    If strCitemStr <> "" Then
        aryCitemStr = Split(strCitemStr, ",")
        For i = LBound(aryCitemStr) To UBound(aryCitemStr)
            If InStr(1, strCode, aryCitemStr(i) & "") <> 0 Then
                ChkInSO106ACitem = True
            Else
                ChkInSO106ACitem = False
                Exit For
            End If
        Next i
    End If
    On Error Resume Next
    CloseRecordset rsCD068Temp
    Exit Function
ChkErr:
    ErrSub Me.Name, "ChkInSO106ACitem"
End Function

Private Function InsertHead(rs As ADODB.Recordset) As Boolean
    On Error GoTo ChkErr
    Dim intReserve As Integer
    Dim strData As Variant
        '�ϧO�X(���O) , ��ƥN��, ������, �o�e���N��, �ƥ�
        strData = ""
        strData = strData & "BOF" & "ACHP02"
        '�����O(1-3)��ƥN��(4-9)
        strData = strData & GetString(GetRealDateTran(gdaSendDate.Text), 8, giRight, True)
        '�B�z���(10-17)
        strData = strData & GetString(txtSendSpcId, 7, giLeft)
        '�o�e���N��(18-24)
        strData = strData & GetString("", 96, giLeft)
        '�ƥ�(25-120)
        WriteLineData strData, DataPath
        InsertHead = True
    Exit Function
ChkErr:
    ErrSub Me.Name, "InsertHead"

End Function

Private Function GetRsTmp(ByRef rs As ADODB.Recordset) As Boolean
    On Error GoTo ChkErr
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset

       'EMC
    strFrom = " " & GetOwner & "SO106 A"
'    if instr(CmbAutoType.Text,"O")
            
            strWhere = strWhere & IIf(strWhere = "", "", " And ") & strChoose
            
            '#3585 �h�D��ACHCUSTID���,�P�_ACHCUSTID�O�_���ȡA�p�G���Ȥ��i��Update By Kin 2007/10/26
            '#3946 �h�W�[MasterId ��� By Kin 2008/05/30
            strSQL = "SELECT RowID,Custid,BankCode,AccountID,AccountNameId,CitemStr,ACHTNo,ACHSN,ACHCUSTID,MasterId From " & _
                        strFrom & " Where " & strWhere & strCancelWhere
         
        If Not GetRS(rs, strSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic, , , True) Then Exit Function
        strReturnSQL = rs.Source
    GetRsTmp = True
    Exit Function
ChkErr:
     ErrSub Me.Name, "GetRsTmp"
End Function

Private Sub Form_Activate()
    On Error Resume Next
        Screen.MousePointer = vbDefault
End Sub
Private Sub DefaultValue()
    On Error Resume Next
'    Dim intPara23 As Integer
        strPrgName = "SO3292A"
        gilCompCode.SetCodeNo garyGi(9)
        gilCompCode.Query_Description
        gdaSendDate.SetValue (RightNow)
        strUpdEn = garyGi(0)
        strUpdName = garyGi(1)
        strErrPath = ReadGICMIS1("ErrLogPath")
        txtInvoiceId = GetRsValue("Select InvoiceId From " & GetOwner & "So041") & ""
        If CmbAutoType.ListIndex = 0 Then
            gdaStopDate1.Enabled = False
            gdaStopDate2.Enabled = False
            lblstoptit.Enabled = False
            lblStop.Enabled = False
        End If
Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "DefaultValue")
End Sub
Private Sub Form_Load()
    On Error GoTo ChkErr
        Call subGim
        Call subGil
        CboAddItem
        Call DefaultValue
        RcdToScr
Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "Form_Load")
End Sub

Private Sub RcdToScr()
    On Error GoTo ChkErr
    Dim LogFile As TextStream
    If Dir(strErrPath & "\" & strPrgName & "Out.log") = "" Then Exit Sub
            On Error Resume Next
            Set LogFile = FSO.OpenTextFile(strErrPath & "\" & strPrgName & "Out.log", ForReading)
            With LogFile
                    If Not .AtEndOfStream Then txtSendSpcId.Text = .ReadLine & ""
                       '�o�e���N��
                    If Not .AtEndOfStream Then gdaSendDate.Text = .ReadLine & ""
                       '�e�X���
                    If Not .AtEndOfStream Then txtPutId.Text = .ReadLine & ""
                       '�o�ʪ̥N��
                    If Not .AtEndOfStream Then txtInvoiceId.Text = .ReadLine & ""
                       '�o�ʪ̲Τ@�s��
                    If Not .AtEndOfStream Then txtDataPath.Text = .ReadLine & ""
                       '����ɦ�m
                    If Not .AtEndOfStream Then txtErrPath.Text = .ReadLine & ""
                       '���D�ɦ�m
                    If Not .AtEndOfStream Then chkStop.Value = .ReadLine & ""
                       '�O�_�N���v���Ѱ���
            End With
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "RcdToScr")
End Sub
Private Sub ScrToRcd()
    On Error GoTo ChkErr
    Dim LogFile As TextStream
        Set LogFile = FSO.CreateTextFile(strErrPath & "\" & strPrgName & "Out.log", True)
        With LogFile
                .WriteLine (txtSendSpcId.Text)         '�o�e���N��
                .WriteLine (gdaSendDate.Text)          '�e�X���
                .WriteLine (txtPutId.Text)             '�o�ʪ̥N��
                .WriteLine (txtInvoiceId.Text)         '�o�ʪ̲Τ@�s��
                .WriteLine (txtDataPath.Text)          '����ɦ�m
                .WriteLine (txtErrPath.Text)           '���D�ɦ�m
                .WriteLine (chkStop.Value)             '�O�_�N���v���Ѱ���
        End With
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "ScrToRcd")
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If Shift = 2 Then
        If KeyCode = vbKeyF Then
            KeyCode = 0
            txtSQL.Visible = True
            txtSQL = strReturnSQL
            txtSQL.Move 0, 0, Me.Width, Me.Height / 2
        ElseIf KeyCode = vbKeyX Then
            txtSQL.Visible = False
        End If
    End If
'    If KeyCode = vbKeyF2 Then Call cmdOk_Click: KeyCode = 0
    If KeyCode = vbKeyEscape Then Unload Me: KeyCode = 0
End Sub

Private Sub subGim()
    On Error GoTo ChkErr
        SetgiMulti gimBankId, "CodeNo", "Description", "CD018", "�N�X", "�W��", , True
        SetgiMulti gimCitemCode, "CodeNo", "Description", "CD019", "�N�X", "�W��", , True
        '#3946 �n�L�oACHTNO ��Null ����� By Kin 2008/05/30
        'SetgiMulti gimBillHeadFmt, "BillHeadFmt", "ACHTNO", "CD068", "��ڮ榡", "����N�X", " Where ACHTNO Is Not Null And ACHTDesc Is not Null"
        '#4016 �W�[�@��ACHType�Ѽ� By Kin 2008/09/22
        SetgiMulti gimBillHeadFmt, "ACHTDESC", "ACHTNO", "CD068", "ACH�������", "ACH����N�X", " Where ACHTNO Is Not Null And ACHTDesc Is not Null And ACHType=1"
Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "subGim")
End Sub

Private Sub subGil()
  On Error GoTo ChkErr
    Call SetgiList(gilCompCode, "CodeNo", "Description", "CD039", , , , , , , GetCompCodeFilter(gilCompCode))
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "subGil")
End Sub

Private Sub CboAddItem()
    On Error Resume Next
    With CmbAutoType
        .AddItem "A-�s�W���v����", 0
        .AddItem "D-�������v����", 1
        .AddItem "O-�s�W�¦��wñ���eú����", 2
    End With
    If Not GetRS(rsBillHeadFmt, "Select BillHeadFmt,CitemCodeStr From " & GetOwner & "CD068 WHERE ACHTNo IS NOT NULL And ACHTDesc is Not Null") Then Exit Sub
'#3728 �N�����אּgimBillHeadFmt By Kin 2008/03/10
'        If Not GetRS(rsBillHeadFmt, "Select BillHeadFmt,CitemCodeStr From " & GetOwner & "CD068 WHERE ACHTNo IS NOT NULL") Then Exit Sub
'        Do While Not rsBillHeadFmt.EOF
'            cboBillHeadFmt.AddItem rsBillHeadFmt("BillHeadFmt") & ""
'            rsBillHeadFmt.MoveNext
'        Loop
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    Call CloseRecordset(rsBillHeadFmt)
    Call CloseRecordset(rsSO106)
    Call CloseRecordset(rsTmp)
End Sub

Private Sub gdaPropDate1_GotFocus()
    On Error GoTo ChkErr
    gdaPropDate1.SetValue (RightNow)
Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "gdaPropDate1_GotFocus")
End Sub

Private Sub gdaPropDate1_Validate(Cancel As Boolean)
    On Error GoTo ChkErr
        If Not IsDate(gdaPropDate1.Text) Then Exit Sub
        If gdaPropDate1.GetValue <> "" Then
            If gdaPropDate2.GetValue = "" Then gdaPropDate2.SetValue GetLastDate(gdaPropDate1.GetValue(True))
        End If
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "gdaPropDate1_Validate")
End Sub

Private Sub gdaPropDate2_GotFocus()
    On Error Resume Next
        If gdaPropDate2.GetValue <> "" Then Exit Sub
        gdaPropDate2.SetValue GetLastDate(Date)
End Sub

Private Sub gdaPropDate2_Validate(Cancel As Boolean)
    On Error Resume Next
        If gdaPropDate1.GetValue = "" Or gdaPropDate2.GetValue = "" Then Exit Sub
        If Not IsDate(gdaPropDate2.Text) Then Exit Sub
        If DateDiff("d", gdaPropDate1.Text, gdaPropDate2.Text) < 0 Then
            MsgBox "���O�I��餣�o�p�󦬶O�_�l��I", vbExclamation, "�T���I"
            gdaPropDate2.SetValue gdaPropDate1.GetValue
            Cancel = True
        End If
End Sub

Private Sub gdaSendDate_GotFocus()
    On Error GoTo ChkErr
    If gdaSendDate.Text = "" Then gdaSendDate.SetValue (RightNow)
Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "gdaSendDate_GotFocus")
End Sub

Private Sub gdaStopDate1_GotFocus()
    On Error GoTo ChkErr
    gdaStopDate1.SetValue (RightNow)
Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "gdaStopDate1_GotFocus")
End Sub

Private Sub gdaStopDate1_Validate(Cancel As Boolean)
    On Error GoTo ChkErr
        If Not IsDate(gdaStopDate1.Text) Then Exit Sub
        If gdaStopDate1.GetValue <> "" Then
            If gdaStopDate2.GetValue = "" Then gdaStopDate2.SetValue GetLastDate(gdaStopDate1.GetValue(True))
        End If
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "gdaStopDate1_Validate")
End Sub

Private Sub gdaStopDate2_GotFocus()
    On Error Resume Next
        If gdaStopDate2.GetValue <> "" Then Exit Sub
        gdaStopDate2.SetValue GetLastDate(Date)
End Sub

Private Sub gdaStopDate2_Validate(Cancel As Boolean)
    On Error Resume Next
        If gdaStopDate1.GetValue = "" Or gdaStopDate2.GetValue = "" Then Exit Sub
        If Not IsDate(gdaStopDate2.Text) Then Exit Sub
        If DateDiff("d", gdaStopDate1.Text, gdaStopDate2.Text) < 0 Then
            MsgBox "���O�I��餣�o�p�󦬶O�_�l��I", vbExclamation, "�T���I"
            gdaStopDate2.SetValue gdaStopDate1.GetValue
            Cancel = True
        End If
End Sub

Private Sub gilCompCode_Change()
    On Error Resume Next
        garyGi(16) = gilCompCode.GetOwner
        Call subGim
        Call subGil
        Call DefaultValue
       If SetACHbankTypeCbo = False Then cmdOK(0).Visible = False: cmdOK(1).Visible = False

End Sub

Private Function SetACHbankTypeCbo() As Boolean

  Dim strSQL As String
  Dim rsTmp As New ADODB.Recordset
  Dim lbn
  Dim lngIndex As Long
  On Error GoTo ChkErr
  
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
                 cboACHbankType.AddItem ("ACH���X�榡 ")
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
ChkErr:
  ErrSub Me.Name, "SetACHbankTypeCbo"
End Function

Private Sub subChoose()
    On Error GoTo ChkErr
    Dim strAddr As String
    Dim strMduChoose As String
    Dim aryACHID() As String
    Dim strWhereACHID As String
    Dim i As Long
    strChoose = ""
    strWhere = ""
    strCancelWhere = ""
    If gdaPropDate1.GetValue <> "" Then Call subAnd("PropDate >= To_date('" & gdaPropDate1.GetValue & "','YYYYMMDD')")
    If gdaPropDate2.GetValue <> "" Then Call subAnd("PropDate < To_date('" & gdaPropDate2.GetValue & "','YYYYMMDD') +1")
    '************************************************************************************************************************
    '#3728 �p�G�O���X���v�����h���Τ�����bSO106A By Kin 2008/03/11
    Select Case CmbAutoType.ListIndex
        Case 0, 2
            If gdaStopDate1.GetValue <> "" Then Call subAnd("StopDate >= To_date('" & gdaStopDate1.GetValue & "','YYYYMMDD')")
            If gdaStopDate2.GetValue <> "" Then Call subAnd("StopDate < To_date('" & gdaStopDate2.GetValue & "','YYYYMMDD') +1")
        Case 1
            If gdaStopDate1.GetValue <> "" Then strCancelWhere = IIf(strCancelWhere = Empty, "StopDate >= To_date('" & gdaStopDate1.GetValue & "','YYYYMMDD')", strCancelWhere & " And StopDate >= To_date('" & gdaStopDate1.GetValue & "','YYYYMMDD')")
            If gdaStopDate2.GetValue <> "" Then strCancelWhere = IIf(strCancelWhere = Empty, "StopDate < To_date('" & gdaStopDate2.GetValue & "','YYYYMMDD')+1", strCancelWhere & " And StopDate < To_date('" & gdaStopDate2.GetValue & "','YYYYMMDD')+1")
    End Select
    '***************************************************************************************************************************
    If gimBankId.GetQueryCode <> "" Then
        Call subAnd("BankCode  " & gimBankId.GetQryStr)
'        If gimBankId.GetSelectCount = 1 Then
'            Call subAnd("BankCode = " & gimBankId.GetQueryCode)
'        Else
'            Call subAnd("BankCode In (" & gimBankId.GetQueryCode & ")")
'        End If
    End If
'        If gimCitemCode.GetQryStr <> "" Then
'            subAnd ("(ACHTNo is not null And CitemStr " & gimCitemCode.GetQryStr & " or ACHTNo is null) ")
'        End If
        '94/11/18 Jacky ��
        '*************************************************************
        '#3728 ACHTNO�אּ�h�� By Kin 2008/03/11
        aryACHID = Split(txtACHTNo, ",")
        For i = LBound(aryACHID) To UBound(aryACHID)
            If aryACHID(i) <> Empty Then strWhereACHID = IIf(strWhereACHID <> Empty, strWhereACHID & " or Instr(ACHTNo,chr(39)||'" & aryACHID(i) & "'||chr(39)) > 0", "Instr(ACHTNo,chr(39)||'" & aryACHID(i) & "'||chr(39)) > 0")
        Next i
        '*************************************************************
        
        '*********************************************************************
        '#3728 �p�G�O�������v���δM��ACHTNO By Kin 2008/03/11
        If CmbAutoType.ListIndex = 0 Or CmbAutoType.ListIndex = 2 Then
            If strWhereACHID <> Empty Then subAnd "(" & strWhereACHID & ")"
        End If
        '*********************************************************************
'        If txtACHTNo <> "" Then subAnd "Instr(ACHTNo,chr(39)||'" & strInACHTNO & "'||chr(39)) > 0"
        
        '#2699 ���դ�OK,Where����SO106.CitemStr���ݭn�ŦXCD008�̪�CitmCodeStr By Kin 2007/10/24
        '#4106 �W�[�P�_ACHType�Ѽ� By Kin 2008/09/22
'        strINCD008Where = "Exists(Select CitemCode From " & GetOwner & "SO003 B Where " & _
'                  "B.Custid = A.Custid " & _
'                  "And B.CompCode=A.CompCode " & _
'                  "And instr(','||A.Citemstr||',',','||Chr(39)||B.Seqno||Chr(39)||',')>0 " & _
'                  "And Exists(Select * From " & GetOwner & "CD068 C Where " & _
'                  "instr(','||C.Citemcodestr||',',','||B.CitemCode||',')>0 " & _
'                  " And C.BillHeadFmt In(" & strInBillHeadFmt & ") And C.ACHType=1 ))"
        '#7049 ���CD068A.CitemCode By Kin 2015/07/08
        strINCD008Where = "Exists(Select CitemCode From " & GetOwner & "SO003 B Where " & _
                " B.Custid = A.Custid " & _
                " And B.CompCode=A.CompCode " & _
                " And instr(','||A.Citemstr||',',','||Chr(39)||B.Seqno||Chr(39)||',')>0 " & _
                " And B.CitemCode In (Select CitemCode From " & GetOwner & "CD068A " & _
                " Where Exists (Select * From CD068 Where CD068.BillHeadFmt=CD068A.BillHeadFmt " & _
                                    " And CD068.BillHeadFmt In (" & strInBillHeadFmt & ") And  CD068.ACHType = 1 )))"
                
        '�̱��v���O�ꤣ�P���z�����
        Select Case CmbAutoType.ListIndex
            '#3946 ���MasterId���M�� By Kin 2008/05/30
            Case 0
                'A �s�W���v
                strWhere = " SnactionDate Is Null And SendDate Is Null And nvl(StopFlag,0) = 0 And " & strINCD008Where & _
                            " And Exists(Select * From " & strOwnerName & " SO106A B" & _
                            " Where A.MasterId=B.MasterId And B.StopFlag<>1 And B.StopDate is Null And ACHTNO IN(" & strInACHTNO & "))"
                            
                            
            Case 1
                '���D��2835 �W�[�P�_�p�G���U���Τ������,StopFlag=1�A�p�G�S�U���ܡAStopFlga=0
                'D �������v
                If gdaStopDate1.GetValue = "" And gdaStopDate2.GetValue = "" Then
                    'strWhere = " DeAuthorize=1 And nvl(StopFlag,0) = 0 "
                    strCancelWhere = IIf(strCancelWhere = Empty, "Nvl(StopFlag,0)=0", strCancelWhere & " And Nvl(StopFlag,0)=0")
                Else
                    'strWhere = " DeAuthorize=1 And StopFlag=1 "
                    strCancelWhere = IIf(strCancelWhere = Empty, "Nvl(StopFlag,0)=1", strCancelWhere & " And Nvl(StopFlag,0)=1")
                End If
                '#3946 ���MasterId���M�� By Kin 2008/05/30
                If strCancelWhere <> Empty Then strCancelWhere = " And Exists(Select * From " & strOwnerName & "SO106A B " & _
                                                                " Where A.MasterId=B.MasterId And ACHTNO IN(" & strInACHTNO & ") And " & strCancelWhere & ")"
            Case 2
                'O
                strWhere = " SnactionDate Is not null and Propdate is not null and OldACH=0 And nvl(StopFlag,0) = 0"
        End Select
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "subChoose")
End Sub

Public Sub subAnd(strCH As String)
  On Error GoTo ChkErr
    If strChoose <> "" Then
        strChoose = strChoose & " And " & strCH
    Else
        strChoose = strCH
    End If
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "subAnd")
End Sub

Private Function InsertFinal(rs As ADODB.Recordset) As Boolean
    On Error GoTo ChkErr
    Dim intReserve As Integer
    Dim strData As Variant
        strData = "EOF"
        '�����O(1-3)
        strData = strData & GetString(intSeq, 8, giRight, True)
        '�`����(4-11)
        strData = strData & GetString("", 109, giLeft)
        '�ƥ�(12-120)
        WriteLineData strData, DataPath
        InsertFinal = True
    Exit Function
ChkErr:
    ErrSub Me.Name, "InsertFinal"
End Function

Private Function GetCustID(strReadLine As String) As Long
    On Error Resume Next
    
        GetCustID = Trim(Val(Mid(strReadLine, 54, 20)))
End Function

Private Function GetACHCustID(strReadLine As String) As String
    On Error Resume Next
        GetACHCustID = Trim(Mid(strReadLine, 51, 20))
End Function

Private Function GetAccID(strReadLine As String) As String
    On Error Resume Next
    Dim intLen As Integer
        intLen = Val(GetRsValue("Select ActLength From " & GetOwner & "CD018" & " Where BankId2 = '" & GetBankCode(strReadLine) & "'") & "")
        GetAccID = Trim(Right(Mid(strReadLine, 27, 14), intLen))
End Function

Private Function GetBankCode(strReadLine As String) As String
    On Error Resume Next
        GetBankCode = Trim(Mid(strReadLine, 20, 7))
End Function

Private Function GetTxtDate(strReadLine As String) As String
    On Error Resume Next
        GetTxtDate = CStr(Val(Mid(strReadLine, 72, 8)) + 19110000)
End Function
'#3728 ���^�ɧ�X�۹�����SO106A By Kin 2008/03/12
'�Ѽ� 1.strType:���v���O 2.strRowId:SO106.RowId  3.strData:���^�q�l�ɥثeŪ����Record
Private Function GetSO106A(ByVal strType As String, ByVal strRowId As String, ByVal strData As String) As String
  On Error GoTo ChkErr
    Select Case strType
        Case "A"
            '#3946 ���MasterId��PK By Kin 2008/05/30
            GetSO106A = "Select A.RowId,A.* From " & GetOwner & "SO106A A " & _
                        "Where A.MasterId='" & strRowId & "'" & _
                        " And A.ACHTNO='" & Trim(MidMbcs(strData, 7, 3)) & "'" & _
                        " And A.ACHTNO In(" & strInACHTNO & ")" & _
                        " And A.RecordType=0" & _
                        " And A.AuthorizeStatus is null" & _
                        " And A.StopFlag<>1"
        Case "D"
            GetSO106A = "Select A.RowId,A.* From " & GetOwner & "SO106A A " & _
                        "Where A.MasterId='" & strRowId & "'" & _
                        " And A.ACHTNO='" & Trim(MidMbcs(strData, 7, 3)) & "'" & _
                        " And A.ACHTNO In(" & strInACHTNO & ")" & _
                        " And A.RecordType=1" & _
                        " And A.AuthorizeStatus is null" & _
                        " And A.StopFlag=1" & _
                        IIf(gdaStopDate1.GetValue <> Empty, " And A.StopDate>=To_Date('" & gdaStopDate1.GetValue & "','YYYYMMDD')", "") & _
                        IIf(gdaStopDate1.GetValue <> Empty, " And A.StopDate<To_Date('" & gdaStopDate1.GetValue & "','YYYYMMDD')+1", "")
        Case "O"
            GetSO106A = "Select A.RowId,A.* From " & GetOwner & "SO106A A" & _
                        " Where A.MasterId='" & strRowId & "'" & _
                        " And A.ACHTNO='" & Trim(MidMbcs(strData, 7, 3)) & "'" & _
                        " And A.ACHTNO In(" & strInACHTNO & ")" & _
                        " And A.RecordType=0" & _
                        " And A.AuthorizeStatus=1" & _
                        " And A.StopFlag<>1"

    End Select
  
    Exit Function
ChkErr:
    ErrSub Me.Name, "GetSO106A"
End Function
'#3728 �������v���\�ɭn�sSO003�@�_����
Private Function StopSO003(ByVal strRowId As String, ByRef rsSO106A As ADODB.Recordset) As Boolean
  On Error GoTo ChkErr
    Dim strSQL As String
    Dim rsQrySO106 As New ADODB.Recordset
    If strRowId = Empty Then Exit Function
    If rsSO106A.State = adStateClosed Then Exit Function
    If rsSO106A.RecordCount = 0 Then Exit Function
    strSQL = "Select * From " & strOwnerName & "SO106 Where RowId='" & strRowId & "'"
    If Not GetRS(rsQrySO106, strSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
    If Not rsQrySO106.EOF Then
        strSQL = "Update " & strOwnerName & "SO003 Set " & _
                 "BANKCODE=NULL" & _
                 ",BANKNAME=NULL" & _
                 ",ACCOUNTNO=NULL" & _
                 ",InvSeqNO=NULL" & _
                 " Where CompCode=" & gilCompCode.GetCodeNo & _
                 " And CustId=" & rsQrySO106("CustId") & _
                 " And CitemCode In(" & rsSO106A("CitemCodeStr") & ")" & _
                 " And ACCOUNTNO='" & rsQrySO106("ACCOUNTID") & "'"
        gcnGi.Execute strSQL
        StopSO003 = True
    End If
    On Error Resume Next
    Call CloseRecordset(rsQrySO106)
    Exit Function
ChkErr:
    ErrSub Me.Name, "StopSO003"
End Function

'#3278 ���v���ѩΨ������v��,�n�N������ACH�B���O���ص����� By Kin 2008/03/18
Private Function UpdateSO106(ByVal strRowId As String, ByRef rsSO106A As ADODB.Recordset) As Boolean
  On Error GoTo ChkErr
    Dim strQry As String
    Dim rsUpdSO106 As New ADODB.Recordset
    Dim rsSO003 As New ADODB.Recordset
    Dim strOldACHTNO As String
    Dim strOldACHTDesc As String
    Dim strOldCitemStr As String
    Dim strQrySeqNo As String
    If strRowId = Empty Then Exit Function
    If rsSO106A.State = adStateClosed Then Exit Function
    If rsSO106A.RecordCount = 0 Then Exit Function
    strQry = "Select A.RowId,A.* From " & strOwnerName & "SO106 A Where RowId='" & strRowId & "'"
    If Not GetRS(rsUpdSO106, strQry, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
    If Not rsUpdSO106.EOF Then
        With rsUpdSO106
            If Not IsNull(.Fields("ACHTNO")) Then strOldACHTNO = .Fields("ACHTNO")
            If Not IsNull(.Fields("ACHTDesc")) Then strOldACHTDesc = .Fields("ACHTDesc")
            If Not IsNull(.Fields("CitemStr")) Then strOldCitemStr = .Fields("CitemStr")
            strQry = "Select * From " & strOwnerName & "SO003" & _
                    " Where CustId=" & .Fields("CustId") & _
                    " And CompCode=" & gilCompCode.GetCodeNo & _
                    " And CitemCode IN(" & rsSO106A("CitemCodeStr") & ")"
            If Not GetRS(rsSO003, strQry, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
            If Not rsSO003.EOF Then
                rsSO003.MoveFirst
                Do While Not rsSO003.EOF
                    strQrySeqNo = "'" & rsSO003("SeqNo") & "'"
                    strOldCitemStr = Replace(Replace(strOldCitemStr, "," & strQrySeqNo, ""), strQrySeqNo, "")
                    rsSO003.MoveNext
                Loop
            End If
            If Left(strOldCitemStr, 1) = "," Then
                strOldCitemStr = Mid(strOldCitemStr, 2)
            End If
            If Right(strOldCitemStr, 1) = "," Then
                strOldCitemStr = Mid(strOldCitemStr, 1, Len(strOldCitemStr) - 1)
            End If
            
            strOldACHTNO = Replace(Replace(strOldACHTNO, ",'" & rsSO106A("ACHTNO") & "'", ""), "'" & rsSO106A("ACHTNO") & "'", "")
            strOldACHTDesc = Replace(Replace(strOldACHTDesc, ",'" & rsSO106A("ACHDesc") & "'", ""), "'" & rsSO106A("ACHDesc") & "'", "")
            If Left(strOldACHTNO, 1) = "," Then
                strOldACHTNO = Mid(strOldACHTNO, 2)
            End If
            If Right(strOldACHTNO, 1) = "," Then
                strOldACHTNO = Mid(strOldACHTNO, 1, Len(strOldACHTNO) - 1)
            End If
            If Left(strOldACHTDesc, 1) = "," Then
                strOldACHTDesc = Mid(strOldACHTDesc, 2)
            End If
            If Right(strOldACHTDesc, 1) = "," Then
                strOldACHTDesc = Mid(strOldACHTDesc, 1, Len(strOldACHTDesc) - 1)
            End If

'            strQry = "Update " & strOwnerName & "SO106 Set " & _
'                    IIf(strOldACHTNO = Empty, "ACHTNO=NULL", "ACHTNO='" & strOldACHTNO & "'") & "," & _
'                    IIf(strOldACHTDesc = Empty, "ACHTDesc=NULL", "ACHTDesc='" & strOldACHTDesc & "'") & "," & _
'                    IIf(strOldCitemStr = Empty, "CitemStr=NULL", "CitemStr='" & strOldCitemStr & "'") & _
'                    " Where RowId='" & strRowId & "'"
'            gcnGi.Execute strQry
            rsUpdSO106("ACHTNO") = IIf(strOldACHTNO = "", Null, strOldACHTNO)
            rsUpdSO106("ACHTDesc") = IIf(strOldACHTDesc = "", Null, strOldACHTDesc)
            rsUpdSO106("CitemStr") = IIf(strOldCitemStr = "", Null, strOldCitemStr)
            rsUpdSO106.Update
        End With
    End If
    UpdateSO106 = True
    On Error Resume Next
    Call CloseRecordset(rsUpdSO106)
    Call CloseRecordset(rsSO003)
  Exit Function
ChkErr:
    ErrSub Me.Name, "UpdateSO106"
End Function

Private Function UpdateSO106A(ByVal rsSO106A As ADODB.Recordset, ByVal strType As String) As Boolean
  On Error GoTo ChkErr
    Dim strChkSQL As String
    Dim strSO106Err As String
    Dim strTime As String
    Dim rsErr As New ADODB.Recordset
    If rsSO106A.RecordCount > 0 Then
        Select Case strType
            Case "A"
                rsSO106A("AuthorizeStatus") = 1
                rsSO106A.Update
            Case "D"    '�������v
                rsSO106A("AuthorizeStatus") = 2
                rsSO106A.Update
                '#3946 ���MasterId
                strChkSQL = "Select Count(Decode(AuthorizeStatus,Null,Null)) A," & _
                                    "Count(Decode(AuthorizeStatus,1,1)) B," & _
                                    "Count(Decode(AuthorizeStatus,2,2)) C," & _
                                    "Count(Decode(AuthorizeStatus,3,3)) D" & _
                            " From " & GetOwner & "SO106A" & _
                            " Where MasterId='" & rsSO106A("MasterId") & "'"
                            
                Set rsErr = gcnGi.Execute(strChkSQL)
                If rsErr("A") = 0 Then
                    If rsErr("C") = rsErr("B") Then
                        blnStopAll = True
                    Else
                        blnStopAll = False
                    End If
                Else
                    blnStopAll = False
                End If
                
            Case "E"        '���v����,�ˬd���v��
                rsSO106A("AuthorizeStatus") = 3
                rsSO106A("UpdTime") = GetDTString(RightNow)
                rsSO106A("UpdEn") = garyGi(1)
                '#3946 ���ѮɳƵ��h�W�[�t�Τ��
                rsSO106A("Notes") = strErrAuth & "   ���Ѥ��:" & Format(Now, "yyyy/mm/dd")
                rsSO106A.Update
                '#3946 ���MasterId��PK By Kin 2008/05/30
                strChkSQL = "Select Count(*) From " & GetOwner & "SO106A" & _
                            " Where AchtNO<>'" & rsSO106A("ACHTNO") & "'" & _
                            " And ACHDesc<>'" & rsSO106A("ACHDesc") & "'" & _
                            " And MasterId=" & rsSO106A("MasterId") & _
                            " And (AuthorizeStatus=1 or AuthorizeStatus is Null)"
                If gcnGi.Execute(strChkSQL)(0) = 0 Then
                    blnStopAll = True
                    '#3946 ���MasterId��PK By Kin 2008/05/30
                    strChkSQL = "Select * From " & GetOwner & "SO106A " & _
                                " Where MasterId=" & rsSO106A("MasterId") & _
                                " And AuthorizeStatus=3"
                    If Not GetRS(rsErr, strChkSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
                    strSO106Err = Empty
                    Do While Not rsErr.EOF
                        strSO106Err = IIf(strSO106Err = Empty, rsErr("ACHDesc") & ":" & rsErr("Notes"), _
                                            strSO106Err & vbCrLf & rsErr("ACHDesc") & ":" & rsErr("Notes"))
                    
                        rsErr.MoveNext
                    Loop
                    If strSO106Err <> Empty Then strSO106Err = Left(strSO106Err, 125)
                    '#3946 ���MasterId��PK By Kin 2008/05/30
                    'Call CloseRecordset(rsErr)
                    '***************************************************************************************************************************************************************************
                    '#4126 �쥻�N�Ƶ�����Update��,���ݨD�n�D�n�ꤧ�e��Note���e By Kin 2008/10/09
                    If Not GetRS(rsErr, "Select Note From " & GetOwner & "SO106 Where Masterid=" & rsSO106A("MasterId"), gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
                    'Set rsErr = gcnGi.Execute("Select Note From " & GetOwner & "SO106 Where Masterid=" & rsSO106A("MasterId"))
                    If Not rsErr.EOF Then
                        strSO106Err = MidMbcs(IIf(rsErr("Note") & "" <> "", rsErr("Note") & vbCrLf & strSO106Err, strSO106Err), 1, rsErr.Fields(0).DefinedSize)
                        If strSO106Err <> Empty Then rsErr("Note").Value = strSO106Err: rsErr.Update
                    End If
                    '***************************************************************************************************************************************************************************
                    'gcnGi.Execute ("Update " & GetOwner & "SO106 Set Note=Note||Chr(13)||'" & strSO106Err & "' Where MasterId=" & rsSO106A("MasterId"))
                Else
                    blnStopAll = False
                End If
                
            
            Case "O"
                '#6119 CreateTime�쥻�ϥ�strTime�|�����D,�{�b�אּStrNowTime(RightNow) By Kin 2011/09/02
                If strNewRowId <> Empty Then
                    'strTime = Trim(str(Val(Mid(strUpdTime, 1, 2) + 1911)) & Mid(strUpdTime, 3))
                    strChkSQL = "Insert Into " & GetOwner & "SO106A" & _
                                "(MasterRowID,ACHTNO,Notes,CitemCodeStr,CitemNameStr," & _
                                " StopFlag,StopDate,UpdEn,UpdTime,CreateTime,CreateEn," & _
                                " RecordType,AuthorizeStatus,AchDesc,MasterId) Values" & _
                                "('" & strNewRowId & "'," & _
                                "'" & rsSO106A("ACHTNO") & "'," & _
                                IIf(IsNull(rsSO106A("Notes")), "Null", "'" & rsSO106A("Notes") & "'") & "," & _
                                "'" & rsSO106A("CitemCodeStr") & "'," & _
                                "'" & rsSO106A("CitemNameStr") & "'," & _
                                IIf(IsNull(rsSO106A("StopFlag")), "Null", rsSO106A("StopFlag")) & "," & _
                                IIf(IsNull(rsSO106A("StopDate")), "Null", "To_Date('" & rsSO106A("StopDate") & "','YYYY/MM/DD HH24:MI:SS')") & "," & _
                                "'" & strUpdName & "'," & _
                                "'" & strUpdTime & "'," & _
                                "To_Date('" & Format(strNowTime, "yyyymmddhhmmss") & "','YYYYMMDDHH24MISS')," & _
                                "'" & strUpdName & "'," & _
                                IIf(IsNull(rsSO106A("RecordType")), "Null", rsSO106A("RecordType")) & "," & _
                                IIf(IsNull(rsSO106A("AuthorizeStatus")), "Null", rsSO106A("AuthorizeStatus")) & "," & _
                                "'" & rsSO106A("AchDesc") & "'," & IIf(strMasterId <> Empty, strMasterId, "Null") & ")"
                                
                    gcnGi.Execute strChkSQL
                End If
            
        End Select
    End If
    UpdateSO106A = True
    On Error Resume Next
    Call CloseRecordset(rsErr)
    Exit Function
ChkErr:
    ErrSub Me.Name, "UpdateSO106A"
End Function

'#3728 �h�b��̾ڳ]�w�אּ�ƿ� By Kin 2008/03/10
Private Sub gimBillHeadFmt_Change()
    On Error Resume Next
    Dim strCitemCode As String
    Dim strCodeNo As String
    Dim rsCD019  As New ADODB.Recordset
    Dim strQry As String
    Dim strACHQry As String
    If gimBillHeadFmt.GetDispStr <> "(����)" Then
        gimBillHeadFmt.SetDispStr Replace(gimBillHeadFmt.GetQueryCode, "'", "")

    End If
    If gimBillHeadFmt.GetDispStr = Empty Then
        gimCitemCode.Clear
        txtACHTNo.Text = Empty
    End If
    If gimBillHeadFmt.GetDispStr = "(����)" Then
        '#3946 �n�L�o ACHTNO Is Not Null By Kin 2008/05/30
        '#4094 �L�o���έn���� By Kin 2008/09/17
        '#4106 �W�[�P�_ACHTYPE�Ѽ� BY Kin 2008/09/22
        strQry = "Select CitemCodeStr From " & GetOwner & "CD068 " & _
                 "Where ACHTNO Is Not Null And ACHTDesc Is Not Null And ACHType=1"
        '#7049 ���CD068A.CitemCode By Kin 2015/07/08
        strQry = "Select CitemCode From " & GetOwner & "CD068A " & _
                " Where  Exists (Select * From " & GetOwner & "CD068 " & _
                        "Where ACHTNO Is Not Null And ACHTDesc Is Not Null And ACHType=1 And CD068.BillHeadFmt = CD068A.BillHeadFmt)"
    Else
        '#4094 �L�o���έn���� By Kin 2008/09/17
        '#4094 �쥻�O�NBillHeadFmt �{�b�אּACHTDESC By Kin 2008/09/17
        '#4106 �W�[�P�_ACHTYPE�Ѽ� BY Kin 2008/09/22
        strQry = "Select CitemCodeStr From " & GetOwner & "CD068 " & _
                 "Where ACHTDESC In(" & gimBillHeadFmt.GetQueryCode & ") And ACHType=1 "
         '#7049 ���CD068A.CitemCode By Kin 2015/07/08
        strQry = "Select CitemCode From " & GetOwner & "CD068A " & _
                " Where  Exists (Select * From " & GetOwner & "CD068 " & _
                        "Where ACHTDESC In(" & gimBillHeadFmt.GetQueryCode & ")   And ACHType=1 And CD068.BillHeadFmt = CD068A.BillHeadFmt)"
    End If
    strCodeNo = gcnGi.Execute(strQry).GetString(, , , ",")
    If Len(strCodeNo) > 0 Then
       strCodeNo = Mid(strCodeNo, 1, Len(strCodeNo) - 1)
    Else
       Exit Sub
    End If
'    If Not GetRS(rsCD019, "Select CodeNo From " & GetOwner & "CD019 Where CodeNo In (" & strCodeNo & ") And StopFlag=0") Then Exit Sub
    '#7219
    If Not GetRS(rsCD019, "Select CodeNo From " & GetOwner & "CD019 Where CodeNo In (" & strQry & ") And StopFlag=0") Then Exit Sub
    If rsCD019.EOF Then Exit Sub
    If gimBillHeadFmt.GetDispStr = "(����)" Then
        '#3946 �n�L�o ACHTNO Is Not Null By Kin 2008/05/30
        '#4094 �L�o���έn���� By Kin 2008/09/17
        '#4106 �W�[�P�_ACHTYPE�Ѽ� BY Kin 2008/09/22
        strACHQry = "Select ACHTNO From " & GetOwner & "CD068 Where  ACHTNO Is Not NULL And ACHTDesc is Not Null And ACHType=1"
    Else
        '#4094 �쥻�O�NBillHeadFmt �{�b�אּACHTDESC By Kin 2008/09/17
        '#4106 �W�[�P�_ACHTYPE�Ѽ� BY Kin 2008/09/22
        strACHQry = "Select ACHTNO From " & GetOwner & "CD068 Where ACHTDESC IN(" & gimBillHeadFmt.GetQueryCode & ") And ACHType=1 "
    End If
    strACHQry = gcnGi.Execute(strACHQry).GetString(, , , ",")
    txtACHTNo.Text = Left(strACHQry, Len(strACHQry) - 1)
    strCitemCode = rsCD019.GetString(, , , ",")
    strCitemCode = Left(strCitemCode, Len(strCitemCode) - 1)
    gimCitemCode.Filter = gimCitemCode.Filter
    gimCitemCode.SetQueryCode strCitemCode

End Sub
'#3946 ���oSO106.MasterId��Sequence By Kin 2008/05/30
Private Function Get_MasterID_Seq() As String
  On Error GoTo ChkErr
    Get_MasterID_Seq = RPxx(gcnGi.Execute("SELECT " & GetOwner & "S_SO106_MasterId.NEXTVAL FROM DUAL").GetString & "")
    'Get_Cmd_Seq_No = Format(Get_Cmd_Seq_No, "00000000")
  Exit Function
ChkErr:
    ErrSub "mod_SysLib", "Get_MasterID_Seq"
End Function


