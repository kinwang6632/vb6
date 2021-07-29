VERSION 5.00
Object = "{3BB22E57-824A-11D3-BD70-0080C8F80BC4}#2.2#0"; "GiGridR.ocx"
Object = "{92575699-8EB8-11D3-B585-0080C8F80BC4}#3.3#0"; "GiNumber.ocx"
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#3.5#0"; "GiList.ocx"
Begin VB.Form frmSO3311F 
   BorderStyle     =   1  '��u�T�w
   Caption         =   "���O��ƽs��  [SO3311F]"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   3495
   ClientWidth     =   11910
   BeginProperty Font 
      Name            =   "�s�ө���"
      Size            =   9.75
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SO3311F.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   11910
   Begin VB.CommandButton cmdShowRate 
      Caption         =   "F7.��ܶO�v��"
      Height          =   375
      Left            =   7350
      TabIndex        =   20
      Top             =   4620
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "F2. �s��"
      Height          =   375
      Left            =   270
      TabIndex        =   17
      Top             =   4620
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "����(&X)"
      Height          =   375
      Left            =   8970
      TabIndex        =   19
      Top             =   4620
      Width           =   1215
   End
   Begin VB.Frame fraData 
      Height          =   4515
      Left            =   300
      TabIndex        =   18
      Top             =   30
      Width           =   11475
      Begin VB.TextBox txtBillNo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1110
         MaxLength       =   15
         TabIndex        =   0
         Top             =   210
         Width           =   1830
      End
      Begin prjGiGridR.GiGridR ggrMduRate 
         Height          =   2025
         Left            =   8640
         TabIndex        =   22
         Top             =   690
         Visible         =   0   'False
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   3572
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin prjGiGridR.GiGridR ggrCustRate 
         Height          =   2025
         Left            =   5760
         TabIndex        =   21
         Top             =   690
         Visible         =   0   'False
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   3572
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin prjNumber.GiNumber gnbQuantity 
         Height          =   285
         Left            =   5070
         TabIndex        =   14
         Top             =   3270
         Visible         =   0   'False
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         MaxLength       =   3
      End
      Begin prjGiList.GiList gilPTCode 
         Height          =   285
         Left            =   1110
         TabIndex        =   13
         Top             =   4140
         Width           =   2940
         _ExtentX        =   5186
         _ExtentY        =   503
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
      Begin prjGiList.GiList gilClctEn 
         Height          =   285
         Left            =   1110
         TabIndex        =   12
         Top             =   3747
         Width           =   2940
         _ExtentX        =   5186
         _ExtentY        =   503
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
      Begin prjGiList.GiList gilSTCode 
         Height          =   285
         Left            =   1110
         TabIndex        =   11
         Top             =   3354
         Width           =   2940
         _ExtentX        =   5186
         _ExtentY        =   503
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
      Begin prjGiList.GiList gilUCCode 
         Height          =   285
         Left            =   1110
         TabIndex        =   10
         Top             =   2961
         Visible         =   0   'False
         Width           =   2940
         _ExtentX        =   5186
         _ExtentY        =   503
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
      Begin prjGiList.GiList gilCMCode 
         Height          =   285
         Left            =   1110
         TabIndex        =   8
         Top             =   2175
         Width           =   2940
         _ExtentX        =   5186
         _ExtentY        =   503
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
      Begin prjGiList.GiList gilCitemCode 
         Height          =   285
         Left            =   1110
         TabIndex        =   1
         Top             =   603
         Width           =   2940
         _ExtentX        =   5186
         _ExtentY        =   503
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
      Begin Gi_Date.GiDate gdaRealStopDate 
         Height          =   285
         Left            =   3300
         TabIndex        =   5
         Top             =   1383
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
      End
      Begin Gi_Date.GiDate gdaRealDate 
         Height          =   285
         Left            =   1110
         TabIndex        =   9
         Top             =   2568
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
      End
      Begin Gi_Date.GiDate gdaRealStartDate 
         Height          =   285
         Left            =   1110
         TabIndex        =   4
         Top             =   1389
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
      End
      Begin VB.TextBox txtManualNo 
         Height          =   285
         Left            =   5040
         MaxLength       =   10
         TabIndex        =   15
         Top             =   3675
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.TextBox txtNote 
         Height          =   285
         Left            =   5040
         MaxLength       =   250
         TabIndex        =   16
         Top             =   4035
         Width           =   5445
      End
      Begin prjNumber.GiNumber gnbRealPeriod 
         Height          =   285
         Left            =   3300
         TabIndex        =   3
         Top             =   996
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   503
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         MaxLength       =   2
      End
      Begin Gi_Date.GiDate gdaShouldDate 
         Height          =   285
         Left            =   1110
         TabIndex        =   2
         Top             =   996
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   503
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
      Begin prjNumber.GiNumber gnbShouldAmt 
         Height          =   285
         Left            =   1110
         TabIndex        =   6
         Top             =   1782
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         MaxLength       =   8
      End
      Begin prjNumber.GiNumber gnbRealAmt 
         Height          =   285
         Left            =   3300
         TabIndex        =   7
         Top             =   1770
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         MaxLength       =   8
      End
      Begin VB.Label lblQuantity 
         AutoSize        =   -1  'True
         Caption         =   "�ƶq"
         Height          =   195
         Left            =   4170
         TabIndex        =   47
         Top             =   3270
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.Label lblManualNo 
         AutoSize        =   -1  'True
         Caption         =   "��}�渹"
         Height          =   195
         Left            =   4125
         TabIndex        =   46
         Top             =   3675
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label lblNote 
         AutoSize        =   -1  'True
         Caption         =   "�Ƶ�"
         Height          =   195
         Left            =   4125
         TabIndex        =   45
         Top             =   4095
         Width           =   390
      End
      Begin VB.Label lblCustRate 
         AutoSize        =   -1  'True
         Caption         =   "�Ȥ����O�O�v��"
         Height          =   195
         Left            =   5790
         TabIndex        =   44
         Top             =   390
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.Label lblMduRate 
         AutoSize        =   -1  'True
         Caption         =   "�j�ӶO�v��"
         Height          =   195
         Left            =   8700
         TabIndex        =   43
         Top             =   390
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lblCustClass 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   7380
         TabIndex        =   42
         Top             =   390
         Width           =   45
      End
      Begin VB.Label lblMduName 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   9750
         TabIndex        =   41
         Top             =   390
         Width           =   45
      End
      Begin VB.Label lblRealPeriod 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   195
         Left            =   2370
         TabIndex        =   40
         Top             =   1020
         Width           =   390
      End
      Begin VB.Label lblRealAmt 
         AutoSize        =   -1  'True
         Caption         =   "�ꦬ���B"
         Height          =   195
         Left            =   2370
         TabIndex        =   39
         Top             =   1830
         Width           =   780
      End
      Begin VB.Label lblShouldAmt 
         AutoSize        =   -1  'True
         Caption         =   "�X����B"
         Height          =   195
         Left            =   210
         TabIndex        =   38
         Top             =   1800
         Width           =   780
      End
      Begin VB.Label lblUpdTime 
         Caption         =   " yyy/mm/dd hh:mm:ss"
         Height          =   285
         Left            =   7035
         TabIndex        =   37
         Top             =   3675
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.Label lblUpdEn 
         Caption         =   "XXXXXXXX"
         Height          =   315
         Left            =   9765
         TabIndex        =   36
         Top             =   3675
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.Label lblUTime 
         Caption         =   "���ʮɶ�:"
         Height          =   375
         Left            =   6105
         TabIndex        =   35
         Top             =   3675
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.Label lblUEn 
         Caption         =   "���ʤH��:"
         Height          =   375
         Left            =   8865
         TabIndex        =   34
         Top             =   3675
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label lblBillNo 
         AutoSize        =   -1  'True
         Caption         =   "��ڽs��"
         Height          =   195
         Left            =   210
         TabIndex        =   33
         Top             =   240
         Width           =   780
      End
      Begin VB.Label lblRealDate 
         AutoSize        =   -1  'True
         Caption         =   "�J�b���"
         Height          =   195
         Left            =   210
         TabIndex        =   32
         Top             =   2580
         Width           =   780
      End
      Begin VB.Label lblClctEn 
         AutoSize        =   -1  'True
         Caption         =   "���O�H��"
         Height          =   195
         Left            =   210
         TabIndex        =   31
         Top             =   3750
         Width           =   780
      End
      Begin VB.Label lblSTCode 
         AutoSize        =   -1  'True
         Caption         =   "�u����]"
         Height          =   195
         Left            =   210
         TabIndex        =   30
         Top             =   3360
         Width           =   780
      End
      Begin VB.Label lblUCCode 
         AutoSize        =   -1  'True
         Caption         =   "������]"
         Height          =   195
         Left            =   210
         TabIndex        =   29
         Top             =   2970
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label lblCMCode 
         AutoSize        =   -1  'True
         Caption         =   "���O�覡"
         Height          =   195
         Left            =   210
         TabIndex        =   28
         Top             =   2190
         Width           =   780
      End
      Begin VB.Label lblRealStopDate 
         AutoSize        =   -1  'True
         Caption         =   "��"
         Height          =   195
         Left            =   2370
         TabIndex        =   27
         Top             =   1440
         Width           =   195
      End
      Begin VB.Label lblRealStartDate 
         AutoSize        =   -1  'True
         Caption         =   "���Ĵ���"
         Height          =   195
         Left            =   210
         TabIndex        =   26
         Top             =   1410
         Width           =   780
      End
      Begin VB.Label lblShouldDate 
         AutoSize        =   -1  'True
         Caption         =   "�������"
         Height          =   195
         Left            =   210
         TabIndex        =   25
         Top             =   1020
         Width           =   780
      End
      Begin VB.Label lblCitemCode 
         AutoSize        =   -1  'True
         Caption         =   "���O����"
         Height          =   195
         Left            =   210
         TabIndex        =   24
         Top             =   630
         Width           =   780
      End
      Begin VB.Label lblPTCode 
         AutoSize        =   -1  'True
         Caption         =   "�I�ں���"
         Height          =   195
         Left            =   210
         TabIndex        =   23
         Top             =   4140
         Width           =   780
      End
   End
   Begin VB.Label lblEditMode 
      Alignment       =   2  '�m�����
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  '��u�T�w
      Caption         =   "���"
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   10530
      TabIndex        =   48
      Top             =   4650
      Width           =   675
   End
End
Attribute VB_Name = "frmSO3311F"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Bruce
'2000/01/18
Option Explicit
Private lngEditMode As giEditModeEnu

'�_�l�Ѽ���
Private Ypara7 As Integer  '���Ĵ���ĵ�i���
Private Ypara10 As Integer '�ϥζO�v��
Private Ypara12 As Integer '���ݿ�J�u����]
Private Ypara14 As Integer '�a�}�̾�
Private Ypara5 As Integer '������(1999/12/21)
Private Ypara8 As Integer '�H���O��Ƨ���(1999/12/21)
Private Ypara9 As Integer '�H���O��Ƨ���(1999/12/21)

'���O���ءA���ơA���İ_�l��A���ĺI���B���O�鵲�I��餧���ʼȦs�ܼ�
'�ɶU�Ÿ�Sign
Private intXCitemCode As Integer
Private intXRealPeriod As Integer
Private strXRealStartDate As String
Private strXRealStopDate As String
Private strXLastDate As String
Private strXSign As String '�ɶU�Ÿ�

Private intRefNo As Integer  '�����O���ؤ��ѦҸ�

'�ŧi�@�O�����X�ѥ����ϥ�
Private rsTmpRec As New ADODB.Recordset

'�������u��ǨӤ�Connection
Public Conn As New ADODB.Connection

'�P�_�O�_���g���ʶO��
Private blnPeriodFlag As Boolean

Private lngCustID As Long '��frmSo3311B�ǨӤ��Ȥ�s��
Private strCustName As String '��frmso3311B�ǨӤ��Ȥ�m�W
Private intBillMode As Integer '��So3311B�ǨӤ���ڪ��A 1:��ڽs��
Private strNote As String '��frmso3311B�ǨӤ��Ƶ����
Private strBillNo As String '�b�s�W�ɥ�frmso3311b�ǨӤ���ڽs��
Private strPrtSno As String '�L��Ǹ�
Private intItem As Integer '�Y�ק�ɵL�k���oRowId�Y���ϥ�BillNo and Item

Private strRowId As String '��frmSo3311B�ǨӤ�RowId�ק�ɨϥ�
Private strMdbSql As String '�O���bmdb����So074�����O
Private rsMdbRec As New ADODB.Recordset '�O���b�ק�Ҧ��U��RecordSet
Private strServiceType As String
Private strCompCode As String

Private Sub cmdCancel_Click()
    If rsTmpRec.State = adStateOpen Then rsTmpRec.Close: Set rsTmpRec = Nothing
    Unload Me
End Sub

Private Sub cmdsave_Click()
On Error GoTo ChkErr
    If IsDataOk Then
        If ScrToRcd Then MsgBox "�s�ɥ��ѡI�Э��s�ާ@�I", vbExclamation, "�T���I"
'        If Not ChkData Then MsgBox "�s�ɥ��ѡI�Э��s�ާ@�I", vbExclamation, "�T���I"
        cmdCancel.Value = True
    End If
Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "cmdSave_Click")
End Sub

Private Sub cmdShowRate_Click()
On Error GoTo ChkErr
    Dim rsCustRate As New ADODB.Recordset
    Dim rsMduRate As New ADODB.Recordset
    Dim MfldsCustRate As New GiGridFlds
    Dim MfldsMduRate As New GiGridFlds
    Dim strCustRateSQL   As String
    Dim strMduRateSQL As String, strMudId As String
    
    ' �]�wggrCustRate �Ȥ�O�v�� ��ܦ��O���A���ơA���B
    strCustRateSQL = "Select A.CustName,C.Description,B.Period,B.Amount "
    strCustRateSQL = strCustRateSQL & "From So001 A,CD019CD004 B,CD019 C Where  A.CustID=" & lngCustID & " and B.ClassCode=A.ClassCode1 and C.CodeNo=B.CitemCode"
    rsCustRate.CursorLocation = adUseClient
    rsCustRate.Open strCustRateSQL, gcnGi, adOpenForwardOnly, adLockReadOnly
    
    If Not rsCustRate.EOF Then
        MfldsCustRate.Add "Description", , , , False, "  ���O����  ", vbLeftJustify
        MfldsCustRate.Add "Period", , , , False, "����", vbLeftJustify
        MfldsCustRate.Add "Amount", , , , False, "���B", vbLeftJustify
    Else
        If rsCustRate.State = adStateOpen Then rsCustRate.Close
        strCustRateSQL = "Select Description,PeriodFlag,Amount From CD019 Where PeriodFlag =1"
        rsCustRate.Open strCustRateSQL, gcnGi, adOpenForwardOnly, adLockReadOnly
        MfldsCustRate.Add "Description", , , , False, "  ���O����  ", vbLeftJustify
        MfldsCustRate.Add "PeriodFlag", , , , False, "����", vbLeftJustify
        MfldsCustRate.Add "Amount", , , , False, "���B", vbLeftJustify
    End If
    ggrCustRate.AllFields = MfldsCustRate
    ggrCustRate.SetHead
    lblCustClass.Caption = "(" & GetValueFromRs2("Select CustName From So001 Where Custid=" & lngCustID) & ")"
    Set ggrCustRate.Recordset = rsCustRate
    ggrCustRate.Refresh
    '�]�wggrMduRate �j�ӶO�v��
    MfldsMduRate.Add "Description", , , , False, "  ���O����  ", vbLeftJustify
    MfldsMduRate.Add "Period", , , , False, "����", vbLeftJustify
    MfldsMduRate.Add "Amount", , , , False, "���B", vbLeftJustify
    ggrMduRate.AllFields = MfldsMduRate
    ggrMduRate.SetHead
    strMudId = GetValueFromRs2("Select Mduid From So001 Where Custid=" & lngCustID) & ""
    If strMudId <> "" Then
        rsMduRate.CursorLocation = adUseClient
        strMduRateSQL = "Select B.Description ,A.Period,A.Amount From CD019SO017 A,CD019 B Where A.MduId='" & strMudId & "' and B.CodeNo=A.CitemCode"
        rsMduRate.Open strMduRateSQL, gcnGi, adOpenForwardOnly, adLockReadOnly
        lblMduName = "(" & GetValueFromRs2("Select Name From So017 Where MduId='" & strMudId & "'") & ")"
        Set ggrMduRate.Recordset = rsMduRate
        ggrMduRate.Refresh
    End If
    ggrCustRate.Visible = True
    ggrMduRate.Visible = True
    lblCustRate.Visible = True
    lblMduRate.Visible = True
Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "cmdShowrate_Click")
End Sub

Private Sub Form_Activate()
On Error GoTo ChkErr
'�]�w��l�ɪ�Focus
        Select Case lngEditMode
            Case giEditModeInsert
                gilCitemCode.SetFocus
            Case giEditModeEdit
                gnbShouldAmt.SetFocus
            Case giEditModeView
                cmdCancel.SetFocus
        End Select
Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "Form_Activeate")
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ChkErr
    'F2�s�� F7��ܶO�v��
    If KeyCode = vbKeyF2 And cmdSave.Enabled Then Call cmdsave_Click: Exit Sub
    If KeyCode = vbKeyF7 And cmdShowRate.Enabled Then Call cmdShowRate_Click: Exit Sub
    'KeyCode = 0
Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "Form_KeyDown")
End Sub

Private Sub Form_Load()
    On Error GoTo ChkErr
    Dim strTmpSql As String
        If Not800600 Then
            Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2 + 1200
        End If
        
        If frmSO3311B.uSelBill = 1 Then
            lblBillNo.Caption = "��ڽs��"
        Else
            lblBillNo.Caption = "�L��Ǹ�"
        End If
        '    ���o�t�ΰѼ�
        '     Call GetGlobal
        
        strMdbSql = "Select * "
    
        '�̪ܳ�鵲�ѼưO����(So062)�����O�鵲�I���
        'frmSo3311 �b�ϥ�
        If lngEditMode <> giEditModeView Then
            strTmpSql = "Select tranDate from So062 Where Type=1"
            rsTmpRec.CursorLocation = adUseClient
            rsTmpRec.Open strTmpSql, gcnGi, adOpenForwardOnly, adLockReadOnly
            If Not rsTmpRec.EOF Then strXLastDate = rsTmpRec(0).Value & ""
            rsTmpRec.Close
        End If
        
        '������So3311B�ǨӤ��Ȥ�s���M�Ȥ�m�W�I
        lngCustID = frmSO3311B.uCustID
        strCustName = frmSO3311B.uCustName
        strNote = frmSO3311B.uNote
        
        '�ܦ��O�Ѽ���(So043)���U���ѼơI
        Call GetParameter
    
        '�]�wGiList���
        Call subGil
        
        '�w�]���g���ʶO�ΡA�]�����O���عw�]�Ȭ��Ĥ@��<1999/12/22>
        blnPeriodFlag = True '<1999/12/22>
    
        '�]�w���ʼҦ�
        Call ChangeMode(lngEditMode)
        PrcList Me
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "Form_Load")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ChkErr
    '�P�_�ϥΪ̬O�_�b���ʪ����A���I
        If cmdCancel.Caption = "����(&X)" Or lngEditMode = giEditModeView Then
            If rsTmpRec.State = adStateOpen Then rsTmpRec.Close: Set rsTmpRec = Nothing
            Set frmSO3311F = Nothing
            Unload Me
            Exit Sub
        End If
            MsgBox "�Х��s�ɩΨ�����A���}�I", vbExclamation, "�T��"
            Cancel = True
Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "Form_QueryUnload")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
        Call ReleaseCOM(Me)

End Sub

Private Sub gdaRealStartDate_Validate(Cancel As Boolean)
On Error GoTo ChkErr
Dim lngChk As Long
If Not IsDateOk(gdaRealStartDate.GetValue) Then gdaRealStartDate.SetFocus: Exit Sub
'�O��SFGetAmount�Ǧ^����
Dim strStopDate As String, lngRealAmount As Long, strMsg As String
'�O��RealStartDate ,RealStopDate �����
Dim strRealStartDate As String, strRealStopDate As String
strRealStartDate = CDate(StrToDate(gdaRealStartDate.GetValue & ""))
strRealStopDate = gdaRealStopDate.GetValue & ""

    If strRealStartDate = "" Or CDate(strRealStartDate) > CDate(StrToDate(strRealStopDate)) Then MsgBox "�����ݦ��ȥB<=�I����": gdaRealStartDate.SetFocus: Exit Sub
    If CDate(StrToDate(strXRealStartDate)) <> CDate(strRealStartDate) Then
        lngChk = SFGetAmount(False, 2, lngCustID, gilCitemCode.GetCodeNo, gnbRealPeriod.Text, gdaRealStartDate.GetValue, gServiceType, strStopDate, lngRealAmount, strMsg)
        If lngChk < 0 Then
            MsgBox strMsg, vbExclamation, "�T��"
            gdaRealStartDate.SetFocus
        Else
            gnbShouldAmt.Text = lngRealAmount
            Call gdaRealStopDate.SetValue(strStopDate)
            '�]�w�U���ܼƥH�ѤU�C�{���ϥ�
            strXRealStartDate = gdaRealStartDate.GetValue
            strXRealStopDate = gdaRealStopDate.GetValue
        End If
    End If
Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "gdaRealStartDate_VAlidate")
End Sub

Private Sub gdaRealStopDate_Validate(Cancel As Boolean)
'On Error GoTo ChkErr
Dim lngChk As Long, strVBYesNo As String
If Not IsDateOk(gdaRealStopDate.GetValue) Then gdaRealStopDate.SetFocus: Exit Sub
'�O��RealStartDate ,RealStopDate �����
Dim strRealStartDate As String, strRealStopDate As String
strRealStartDate = StrToDate(gdaRealStartDate.GetValue & "") 'StrToDate YYYY/MM/DD
strRealStopDate = gdaRealStopDate.GetValue & ""
'�O��SFGetAmount�Ǧ^����
Dim strStopDate As String, lngRealAmount As Long, strMsg As String

    If strRealStopDate = "" Or CDate(strRealStartDate) > CDate(StrToDate(strRealStopDate)) Then MsgBox "�����ݦ��ȥB<=�I����": gdaRealStopDate.SetFocus: Exit Sub
    If CDate(StrToDate(strXRealStopDate)) <> CDate(StrToDate(strRealStopDate)) Then
        '�Y��������h���İ_�l����j��intYPara7�A�h��ܽT�{����
        If DateDiff("d", CDate(strRealStartDate), CDate(StrToDate(strRealStopDate))) > Ypara7 Then
                If vbNo = MsgBox("���Ĵ����W�L�t�ιw�]�ѼơA�O�_����������I", vbExclamation, "�T��") Then gdaRealStopDate.SetFocus: Exit Sub
        End If
        lngChk = SFGetAmount(True, 2, lngCustID, gilCitemCode.GetCodeNo, gnbRealPeriod.Text, gdaRealStartDate.GetValue, gServiceType, strRealStopDate, lngRealAmount, strMsg)
        If lngChk < 0 Then
            MsgBox strMsg, vbExclamation, "�T��"
            gdaRealStopDate.SetFocus
        Else
            If GetNumberType(lngRealAmount & "") <> 0 Then
                 gnbShouldAmt.Text = Val(lngRealAmount & "")
                '�]�w�U���ܼƥH�ѤU�C�{���ϥ�
                strXRealStartDate = gdaRealStartDate.GetValue & ""
                strXRealStopDate = gdaRealStopDate.GetValue & ""
            End If
        End If
    End If
Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "gdaRealStopDate_VAlidate")
End Sub

Private Sub gdaShouldDate_Validate(Cancel As Boolean)
    If gdaShouldDate.Text = "" Then Exit Sub
'    If Not IsDateOk(gdaShouldDate.GetValue) Then MsgBox "������X�k�I": gdaShouldDate.SetFocus: Exit Sub
End Sub

Private Sub ggrCustRate_ShowCellData(ByVal Col As Long, ByVal Row As Long, ByVal giFld As prjGiGridR.GiGridFld, ByVal Fld As ADODB.Field, Value As Variant)
On Error GoTo ChkErr
    If Fld.Name = "AMOUNT" Then
        Value = Format(Value, "###,###,###")
    End If
Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "ggrCustRate_ShowCellDate")
End Sub

Private Sub ggrMduRate_ShowCellData(ByVal Col As Long, ByVal Row As Long, ByVal giFld As prjGiGridR.GiGridFld, ByVal Fld As ADODB.Field, Value As Variant)
On Error GoTo ChkErr
    If Fld.Name = "AMOUNT" Then
        Value = Format(Value, "###,###,###")
    End If
Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "ggrMduRate_ShowCellDate")
End Sub

Private Sub gilCitemCode_Validate(Cancel As Boolean)
'On Error GoTo ChkErr
Dim strTmpSql As String, intAmount As Long, intPeriod As Integer
Dim blnPeriodFlagMode As Boolean
Dim rsTmpCitemCode As New ADODB.Recordset
'�O��SFGetAmount �Ǧ^���������
Dim strStopDate As String, lngRealAmount As Long, strMsg As String
Dim lngChk As Long
    
    If gilCitemCode.GetCodeNo = "" Or gilCitemCode.GetDescription = "" Then MsgBox "����쥲�ݦ���", vbExclamation, "�T��": gilCitemCode.SetFocus: Exit Sub
    '�Y���ȥB������
    If intXCitemCode <> gilCitemCode.GetCodeNo Then
        '�ܦ��O���إN�X�ɨ��]�O�_�g���ʶO�Ρ^�A�]�w�]���B�^�A(�w�]���ơ^<<1999/12/22>>�A�]�ѦҸ��^<<1999/12/22>>,(�ɶU�Ÿ�,Sign�^<<2000/03/10>>
        strTmpSql = "Select PeriodFlag,Amount,Period,RefNo,Sign From CD019 Where CodeNo=" & gilCitemCode.GetCodeNo
        
        '�}�ҳs�u�I
        If OpenRecordset(rsTmpCitemCode, strTmpSql, gcnGi, adOpenForwardOnly, adLockReadOnly, adUseServer, False, False) <> giOK Then MsgBox "�s�u���ѡI�Э��s�ާ@�I", vbExclamation, "�T���I"
        
        blnPeriodFlag = rsTmpCitemCode("PeriodFlag").Value
        strXSign = rsTmpCitemCode("Sign").Value
        intAmount = GetNumberType(rsTmpCitemCode("Amount").Value & "")
        intPeriod = Val(rsTmpCitemCode("Period").Value & "")      '<<1999/12/22>>�w�]����
        intRefNo = Val(rsTmpCitemCode("RefNO").Value & "")        '<<1999/12/22>>�ѦҸ�
        rsTmpCitemCode.Close
        Set rsTmpCitemCode = Nothing
        '�Y�]�O�_�g���ʶO�O�Ρ^=�_ �h�G
        '������=0�A���İ_�l���/�I����=�L�A�X����B=�]�w�]���B�^
        '��Disable ���ơA���İ_���/�I����
        If blnPeriodFlag = False Then
            blnPeriodFlagMode = False
            gnbRealPeriod.Text = ""
            Call gdaRealStopDate.SetValue("")
            Call gdaRealStartDate.SetValue("")
            gnbShouldAmt.Text = intAmount
        Else
            gnbRealPeriod.Text = intPeriod '<<1999/12/22>> Old gnbRealPeriod.Text =1
            
           Call gdaRealStartDate.SetValue(ReturnRealStartDate(gilCitemCode.GetCodeNo, lngCustID)) '<<1999/12/21>>
           gdaShouldDate.SetValue gdaRealStartDate.GetValue
            lngChk = SFGetAmount(False, 2, lngCustID, gilCitemCode.GetCodeNo, gnbRealPeriod.Text, gdaRealStartDate.GetValue, gServiceType, strStopDate, lngRealAmount, strMsg)
            If lngChk < 0 Then
                MsgBox strMsg, vbExclamation, "�T��"
                On Error Resume Next
                gilCitemCode.SetFocus
                On Error GoTo ChkErr
            Else
                gnbShouldAmt.Text = lngRealAmount
                Call gdaRealStopDate.SetValue(strStopDate & "")
            End If
        End If
        gnbRealAmt.Text = gnbShouldAmt.Value
        
        gnbRealPeriod.Enabled = blnPeriodFlagMode
        gdaRealStartDate.Enabled = blnPeriodFlagMode
        gdaRealStopDate.Enabled = blnPeriodFlagMode
        intXCitemCode = Val(gilCitemCode.GetCodeNo & "")
        intXRealPeriod = Val(gnbRealPeriod.Text & "")
        strXRealStartDate = gdaRealStartDate.GetValue & ""
        strXRealStopDate = gdaRealStopDate.GetValue & ""
     End If
Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "gilCitemCode_Validate")
End Sub

Private Sub gnbQuantity_GotFocus()
    With gnbQuantity
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub gnbRealAmt_GotFocus()
    With gnbRealAmt
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub gnbRealAmt_Validate(Cancel As Boolean)
On Error GoTo ChkErr
Dim intShouldAmt As Long
Dim intRealAmt As Long
intShouldAmt = gnbShouldAmt.Value
intRealAmt = gnbRealAmt.Value
'�Y���ȡA�h�M��������]
'�Y���ȥB�X����B���ȡG
'   '�B�Y�G�̤��۵��A�hĵ�i�G"�ꦬ�P�X����B���۵�",�ö�J�u����]���1�ι����W��
    '�B�Y�G�̬۵��A�h�M���u����]
    
    If gnbRealAmt.Value < 0 Then
            gnbRealAmt.Text = (-1) * gnbRealAmt.Value
    End If
    If intRealAmt > 0 Then
        gilUCCode.Clear
    Else
        gilSTCode.Clear
        gilUCCode.ListIndex = 1
    End If

    If intShouldAmt > 0 And intRealAmt > 0 Then
        If intShouldAmt <> intRealAmt Then
            MsgBox "�ꦬ�P�X����B���۵�,�u����]������"
            If intRealAmt > intShouldAmt Then gnbRealAmt.Text = 0
            'gilSTCode.ListIndex = 1
            gilSTCode.SetFocus
        ElseIf intShouldAmt = intRealAmt Then
            gilSTCode.Clear
        End If
    End If
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "gnbRealAmt_VailDate")
End Sub

Private Sub gnbRealPeriod_GotFocus()
    With gnbRealPeriod
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub gnbRealPeriod_Validate(Cancel As Boolean)
On Error GoTo ChkErr
Dim lngChk As Long, strStopDate As String
Dim lngRealAmount As Long, strMsg As String
    '�Y���ƦbEnable�����A�U������
If Not gnbRealPeriod.Enabled Then Exit Sub

    '�Y�L�ȩ�<=0�h���~:"����쥲�ݬ����ȡI"�B���}
    If gnbRealPeriod.Text <= 0 Or str(gnbRealPeriod.Text) = "" Then MsgBox "����쥲�ݦ��ȥB�ݬ�����", vbExclamation, "�T��": Exit Sub
    
    If intXRealPeriod <> gnbRealPeriod.Text Then
    lngChk = SFGetAmount(False, 2, lngCustID, gilCitemCode.GetCodeNo, gnbRealPeriod.Text, gdaRealStartDate.GetValue, gServiceType, strStopDate, lngRealAmount, strMsg)
        If lngChk < 0 Then
            MsgBox strMsg, vbExclamation, "�T��"
            gnbRealPeriod.SetFocus
        Else
            If lngRealAmount = 0 Then lngChk = SFGetAmount(True, 2, lngCustID, gilCitemCode.GetCodeNo, gnbRealPeriod.Text, gdaRealStartDate.GetValue, gServiceType, gdaRealStopDate.GetValue, lngRealAmount, strMsg)
            If lngChk < 0 Then
                MsgBox strMsg, vbExclamation, "�T��"
                gnbRealPeriod.SetFocus
                Exit Sub
            End If
            gnbShouldAmt.Text = lngRealAmount
            Call gdaRealStopDate.SetValue(strStopDate)
            '�]�w�U���ܼƥH�ѤU�C�{���ϥ�
            intXRealPeriod = gnbRealPeriod.Text
            strXRealStartDate = gdaRealStartDate.GetValue
            strXRealStopDate = gdaRealStopDate.GetValue
        End If
    End If
Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "gnbRealPeriod_Validate")
End Sub

Private Sub gnbShouldAmt_GotFocus()
    With gnbShouldAmt
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub gnbShouldAmt_Validate(Cancel As Boolean)
On Error GoTo ChkErr
Dim intShouldAmt As Long
Dim intRealAmt As Long
intShouldAmt = gnbShouldAmt.Value
intRealAmt = gnbRealAmt.Value
    '�X����B�Y���ȥB�ꦬ���B���ȡG
        '�Y�G�̤��۵��A�hĵ�i
        '�Y�G�̬۵��A�h�M���u����]<1999/12/27>
        If gnbShouldAmt.Value < 0 Then
            gnbShouldAmt.Text = (-1) * gnbShouldAmt.Value
        End If
        If intRealAmt > 0 And intShouldAmt > 0 Then
            If intRealAmt <> intShouldAmt Then
                MsgBox "�ꦬ�P�X����B���۵�,�u����]������"
                If intRealAmt > intShouldAmt Then
                        gnbRealAmt.Text = 0
                ElseIf intRealAmt = 0 Then
                        gilSTCode.Clear
                        gilUCCode.ListIndex = 1
                    Else
                        'gilSTCode.ListIndex = 1
                    End If
            ElseIf intRealAmt = intShouldAmt Then
                gilSTCode.Clear
            End If
        End If
Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "gnbShouldAmt_Validate")
End Sub

Private Sub txtManualNo_GotFocus()
    With txtManualNo
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtNote_GotFocus()
    With txtNote
        .IMEMode = 1
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub ChangeMode(ByVal lngMode As giEditModeEnu)
On Error GoTo ChkErr
    Dim blnFlag As Boolean  ' �O���O�_�b����s���Ҧ��A
    lngEditMode = lngMode
    
    Select Case lngEditMode
      Case giEditModeInsert
           blnFlag = True
           lblEditMode.Caption = "�s�W"
           cmdCancel.Caption = "����(&X)"
           Call NewRec
      Case giEditModeEdit
           blnFlag = False
           lblEditMode.Caption = "�ק�"
           cmdCancel.Caption = "����(&X)"
           Call RcdToScr
    End Select
    '�]�w�U�Ӫ���Enabled/Disabled
    gilCitemCode.Enabled = blnFlag  '���O����
    gdaShouldDate.Enabled = blnFlag '�������
    
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "ChangeMode")
End Sub

' �ۭq�ݩʡGEditMode
' �O���ثe�b�s��B�s�W���˵��Ҧ�
' giEditModeEnu(�ۭq�C�|�ȡA�]�w��Sys_Lib)
Public Property Get EditMode() As giEditModeEnu
  On Error GoTo ChkErr
    ' ���ثe�s��Ҧ�
    EditMode = lngEditMode
  Exit Property
ChkErr:
   Call ErrSub(Me.Name, "Get EditMode")
End Property

Public Property Let EditMode(ByVal vNewValue As giEditModeEnu)
  On Error GoTo ChkErr
    '�]�w�s��Ҧ�
    lngEditMode = vNewValue
  Exit Property
ChkErr:
    Call ErrSub(Me.Name, "Let EditMode")
End Property

Private Sub NewRec()
On Error GoTo ChkErr
    Dim strShouldDate As String
    
    '�L��Ǹ�
    txtBillNo.Text = IIf(frmSO3311B.uSelBill = 2, strPrtSno, strBillNo)
'    gilCitemCode
    '���O����
    Call gilCMCode.SetCodeNo("")
    Call gilCMCode.SetDescription("")
    '�J�b���
    Call gdaRealDate.SetValue(IIf(frmSO3311B.uRealDate <> "", frmSO3311B.uRealDate, ""))
    '���O�覡
     gilCitemCode.ListIndex = 1
    'Call gilCitemCode.SetDescription(IIf(frmSO3311B.uCMName <> "", frmSO3311B.uCMName, ""))
    
    '���O�H��
    Call gilClctEn.SetCodeNo(IIf(frmSO3311B.uClctEn <> "", frmSO3311B.uClctEn, ""))
    Call gilClctEn.SetDescription(IIf(frmSO3311B.uClctName <> "", frmSO3311B.uClctName, ""))
    '�I�ں���
    Call gilPTCode.SetCodeNo(IIf(frmSO3311B.uPTCode <> "", frmSO3311B.uPTCode, ""))
    Call gilPTCode.SetDescription(IIf(frmSO3311B.uPTName <> "", frmSO3311B.uPTName, ""))
    '���ʮɶ�
    lblUpdTime.Caption = GetDTString(Now)
    '���ʤH��
    lblUpdEn.Caption = garyGi(1)
    '������]
    gilUCCode.Clear
    '�u����]
    gilSTCode.Clear
    '�������
    gdaShouldDate.SetValue (IIf(frmSO3311B.uRealDate <> "", frmSO3311B.uRealDate, ""))
    '���İ_�l��
    gdaRealStartDate.SetValue ("")
    '���ĺI���
    gdaRealStopDate.SetValue ("")
    '����
    gnbRealPeriod.Text = ""
    '�ꦬ���B
    gnbRealAmt.Text = ""
    '�������B
    gnbShouldAmt.Text = ""
    '�ƶq
    gnbQuantity.Text = ""
    '��l��
    intXCitemCode = 0
    intXRealPeriod = 0
    strXSign = "+"
Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "NewRec")
End Sub

''�����W�hRecordSET rsSo1132A_View V_Charge
'Public Property Get rsInputRec() As ADODB.Recordset
'On Error GoTo ChkErr
'    Set AddBillNo = rsSo1132ARec
'Exit Property
'ChkErr:
'    Call ErrSub(Me.Name, "prtItem")
'End Property
'
''�]�w�W�hRecordSET rsSo1132A_View V_Charge
'Public Property Let rsInputRec(ByVal rsNewValue As ADODB.Recordset)
'On Error GoTo ChkErr
'    Set rsSo1132ARec = rsNewValue
'    strBillNo = rsSo1132ARec("BillNo").Value & ""
'    intItem = Val(rsSo1132ARec("Item").Value & "")
''    rsSo1132ARec.Close
'Exit Property
'ChkErr:
'    Call ErrSub(Me.Name, "proItem")
'End Property

Private Sub GetParameter()
On Error GoTo ChkErr
    Dim rsSO043 As New ADODB.Recordset
    Dim strSo043Sql As String
    
    strSo043Sql = "Select Para7,Para10,Para12,Para14,Para5,Para8,Para9 From So043"
    rsSO043.Open strSo043Sql, gcnGi, adOpenForwardOnly, adLockReadOnly
        Ypara5 = rsSO043("Para5").Value 'Update(199/12/21)
        Ypara8 = rsSO043("Para8").Value '(199/12/21)
        Ypara9 = rsSO043("Para9").Value '(199/12/21)
        Ypara7 = rsSO043("Para7").Value
        Ypara10 = rsSO043("Para10").Value
        Ypara12 = rsSO043("Para12").Value
        Ypara14 = rsSO043("Para14").Value
    rsSO043.Close
Exit Sub
ChkErr:
    rsSO043.Close
    Call ErrSub(Me.Name, "GetParameter")
End Sub

Private Sub RcdToScr()
On Error GoTo ChkErr
Dim strSelSql As String
    
    If rsMdbRec.State = adStateOpen Then rsMdbRec.Close
    If strRowId <> "" Then
        strSelSql = strMdbSql & " From So3311 Where Rowid='" & strRowId & "'"
    Else
        strSelSql = strMdbSql & " From So3311 Where Billno='" & strBillNo & "' And Item = " & intItem
    End If
    If Not GetRS(rsMdbRec, strSelSql, Conn, adUseClient, adOpenKeyset, adLockOptimistic) Then
        MsgBox "�s�u���ѡI�Э��s�ާ@�I", vbExclamation, "�T���I": Unload Me
    End If
    txtBillNo.Text = rsMdbRec("BillNo").Value                               '��ڽs��
    Call gilCitemCode.SetCodeNo(rsMdbRec("CitemCode").Value & "")           '���O���إN��
    Call gilCitemCode.SetDescription(rsMdbRec("CitemName").Value & "")      '���O���ئW��
    Call gdaShouldDate.SetValue(Format(rsMdbRec("ShouldDate").Value & "", "YYYYMMDD"))  '�������
    
     gnbShouldAmt.Text = IIf(GetNumberType(rsMdbRec("ShouldAmt").Value & "") = 0, "", GetNumberType(rsMdbRec("ShouldAmt").Value & ""))              '�X����B
     If gnbShouldAmt.Value < 0 Then
        strXSign = "-"
     End If
     gnbRealAmt.Text = GetNumberType(rsMdbRec("RealAmt").Value & "")     '������B
     '-----------------------------------------------------------------------------------------
    '�Y���ק�Ҧ��G�Y���� >0 �A�hEnable ����,���İ_�l���I���,�Ϥ��hdisable
    If rsMdbRec("RealPeriod").Value > 0 Then
        blnPeriodFlag = True
    Else
        blnPeriodFlag = False
    End If
    gnbRealPeriod.Enabled = blnPeriodFlag
    gdaRealStartDate.Enabled = blnPeriodFlag
    gdaRealStopDate.Enabled = blnPeriodFlag
    If rsMdbRec("RealPeriod").Value > 0 Then
            gnbRealPeriod.Text = IIf(GetNumberType(rsMdbRec("RealPeriod").Value) = 0, "", GetNumberType(rsMdbRec("RealPeriod").Value))        '����
            Call gdaRealStartDate.SetValue(Format(rsMdbRec("RealStartDate") & "", "YYYYMMDD"))      '���İ_�l��
            Call gdaRealStopDate.SetValue(Format(rsMdbRec("RealStopDate") & "", "YYYYMMDD"))     '���ĺI���
            '�]�w�Ȧs�ܼƤ��e
            intXCitemCode = GetNumberType(rsMdbRec("CitemCode").Value & "")
            intXRealPeriod = rsMdbRec("RealPeriod").Value
            strXRealStartDate = Format(rsMdbRec("RealStartDate").Value & "", "YYYYMMDD")
    End If
     
'-----------------------------------------------------------------------------------------
     Call gilCMCode.SetCodeNo(rsMdbRec("CMCode").Value & "")              '���O�N�X
     Call gilCMCode.SetDescription(rsMdbRec("CMName").Value & "")                   '�W��
     Call gdaRealDate.SetValue(Format((rsMdbRec("RealDate").Value & ""), "YYYYMMDD"))         '�ꦬ���
     'Call gilUCCode.SetCodeNo(rsMdbRec("UCCode").Value & "")                        '������]�N�X
     'Call gilUCCode.SetDescription(rsMdbRec("UCName").Value & "")                   '������]�W��
     Call gilSTCode.SetCodeNo(rsMdbRec("STCode").Value & "")                        '�u����]�N�X
     Call gilSTCode.SetDescription(rsMdbRec("STName").Value & "")                   '�u����]�W��
     Call gilClctEn.SetCodeNo(rsMdbRec("ClctEn").Value & "")                        '���O�H���N�X
     Call gilClctEn.SetDescription(rsMdbRec("ClctName").Value & "")                 '���O�H���W��
     Call gilPTCode.SetCodeNo(rsMdbRec("PTCode").Value & "")                        '�I�ں����N�X
     Call gilPTCode.SetDescription(rsMdbRec("PTName").Value & "")                   '�I�ں����W��
     txtManualNo.Text = frmSO3311B.uManualNo                         '��}�渹
     txtNote.Text = rsMdbRec("Note").Value & ""                                      '�Ƶ�
     'strSign = GetRsValue("Select Sign From CD019 Where CodeNo = " & gilCitemCode.GetCodeNo) & ""
     
     '�P�_�p�G�O��ܼҦ��N�NrsTmpRec�����A�_�h�Y���s��Ҧ����A���Ѧs�ɮɨϥ�
     
Exit Sub
ChkErr:
     Call ErrSub(Me.Name, "RcdToScr")
End Sub

'�s�W�ɨ��o�b��s��
Private Function GetBillNo(ByVal strDate As Date, ByVal strBillNo As String) As String
On Error Resume Next
Dim str_BillNo As String
    str_BillNo = Year(strDate) & String(2 - Len(Month(strDate)), "0") & Month(strDate)
    str_BillNo = str_BillNo & "T" & String(8 - Len(strBillNo), "0") & strBillNo
    GetBillNo = str_BillNo
End Function

Private Function IsDataOk() As Boolean
On Error GoTo ChkErr
      IsDataOk = False
      '�X����B���o�p�󵥩�0�A�ΪŭȡI
      
'      If gnbShouldAmt.Value <= 0 Or gnbShouldAmt.Text = "" Then MsgBox "�X����B���o�p�󵥩�s�I", vbExclamation, "�T���I": Exit Function
      
      '���n���G��ڽs���B���ءA���O���إN�X�A��������A�ƶq�A���O�覡�N�X
      If gdaShouldDate.GetValue = "" Then MsgBox gMsgIsDataOK, vbExclamation, "�T��": gdaShouldDate.SetFocus: Exit Function
      If gilCitemCode.GetCodeNo = "" Then MsgBox gMsgIsDataOK, vbExclamation, "�T��": gilCitemCode.SetFocus: Exit Function
      If gilCMCode.GetCodeNo = "" Then MsgBox gMsgIsDataOK, vbExclamation, "�T��": gilCMCode.SetFocus: Exit Function
     
      '�Y����>0�A�h���İ_�l��P�I���ݦ��ȡA�B�_�l��<=�I���
      If GetNumberType(gnbRealPeriod.Text & "") > 0 Then
          If gdaRealStartDate.Text = "" Then MsgBox "���İ_�l�鶷���ȡI", vbExclamation, "�T���I": gdaRealStartDate.SetFocus: Exit Function
          If gdaRealStopDate.Text = "" Then MsgBox "���ĺI��鶷���ȡI", vbExclamation, "�T���I": gdaRealStopDate.SetFocus: Exit Function
          If (CDate(StrToDate(gdaRealStartDate.GetValue & "")) > CDate(StrToDate(gdaRealStopDate.GetValue))) Then
                  MsgBox "���İ_�l��P�I���ݦ��ȡA�B�_�l�鶷<=�I���"
                  Exit Function
           End If
      End If
      
      '�Y�X����B>0�A�B�X����B!=�ꦬ���B�A�h�u����]�ݦ���(OLD)
      '�Y�X����B <> 0�A�B�X����B!=�ꦬ���B�A�h�u����]�ݦ���<2000/03/10>
      
       If Val(gnbShouldAmt.Text & "") <> 0 And (Val(gnbRealAmt.Text & "") <> Val(gnbShouldAmt.Text & "")) Then
           If gilSTCode.GetCodeNo = "" Then
                MsgBox "�X����B������ꦬ���B�ݦ��u����]�I", vbExclamation, "�T��"
                gilSTCode.SetFocus
                Exit Function
            End If
       End If
     
    IsDataOk = True
Exit Function
ChkErr:
    Call ErrSub(Me.Name, "IsDataOK")
End Function

Private Function ReturnDate(ByVal strDate As String) As Date
    On Error GoTo ChkErr
        If strDate <> "" Then
            ReturnDate = CDate(strDate) + Ypara5
        End If
    Exit Function
ChkErr:
        Call ErrSub(Me.Name, "Returndate")
End Function

Private Sub GetPeriod_Amt(ByRef intReturnPeriod As Integer, ByRef intReturnAmt As Long)
    On Error GoTo ChkErr
        '�Y�H���O��Ƨ���2�]YPara9)��1�]�������B�^�A�h�G
        '   ����=<�ӵ��즬�O����>�A���B=<�ӵ����������B>
        '�YYPara9�� 2�]�X����B�^�A�h�G
        '   ����=<�ӵ����O����>�A���B=<�ӵ��X����B>
    Dim rsTmp As New ADODB.Recordset
    Dim strsql As String
        rsTmp.MaxRecords = 1
        strsql = "Select OldAmt,OldPeriod From So033 Where BillNo='" & strBillNo & "' and Item =" & intItem
        If Ypara9 = 1 Then
                If OpenRecordset(rsTmp, strsql, gcnGi, adOpenForwardOnly, adLockReadOnly, adUseClient, False, False) <> giOK Then MsgBox "�s�����ѡI�Э��s�ާ@�I", vbExclamation, "�T���I"
                intReturnPeriod = rsTmp("Oldperiod").Value
                intReturnAmt = rsTmp("OldAmt").Value
                rsTmp.Close
                Set rsTmp = Nothing
        Else
                intReturnPeriod = Val(gnbRealPeriod.Text)
                intReturnAmt = Val(gnbShouldAmt.Text)
        End If
        
    Exit Sub
ChkErr:
        Call ErrSub(Me.Name, "GetPeriod_Amt")
End Sub

Private Function ScrToRcd() As Boolean
    On Error GoTo ChkErr
    Dim rsScr As New ADODB.Recordset '�bScrtoRcd �ϥΪ�Recordset
    Dim strsql As String
    Dim lngItem As Long '�O���n�x�s�ɪ����ؽs��
        If Not IsDataOk Then Exit Function
        If frmSO3311B.uSelBill = 1 Then
            strsql = "Select Max(Item)+1 as MaxItem From So033 Where Billno='" & strBillNo & "'"
        Else
            strsql = "Select Max(Item)+1 as MaxItem From So033 Where PrtSno='" & strPrtSno & "'"
        End If
        If rsMdbRec.State = adStateClosed Then If Not GetRS(rsMdbRec, "Select * From So3311", Conn, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
        
        If OpenRecordset(rsScr, strsql, gcnGi, adOpenForwardOnly, adLockReadOnly, adUseClient, False) <> giOK Then MsgBox "�s�u���ѡI�Э��s�ާ@�I", vbExclamation, "�T���I": Exit Function
        '�O�����ؽs��
        
        If Not (rsScr.EOF Or IsNull(rsScr("MaxItem").Value)) Then
            lngItem = rsScr("MaxItem").Value
        Else
            lngItem = 1
        End If
        rsScr.Close
        Set rsScr = Nothing
        
        
        With rsMdbRec
        
        If lngEditMode = giEditModeInsert Then
            .AddNew
        End If
            .Fields("BillNo").Value = strBillNo
            .Fields("PrtSno").Value = strPrtSno
            .Fields("Item").Value = lngItem
            .Fields("CustName").Value = strCustName
            .Fields("CustId").Value = lngCustID
            .Fields("CItemCode").Value = gilCitemCode.GetCodeNo
            .Fields("CItemName").Value = gilCitemCode.GetDescription
            .Fields("ShouldAmt").Value = NoZero(gnbShouldAmt.Value, True)
            
            .Fields("ShouldDate").Value = Format(gdaShouldDate.GetValue, "####/##/##")
            If gnbRealPeriod.Value > 0 Then
                .Fields("RealPeriod").Value = NoZero(gnbRealPeriod.Value, True)
                .Fields("RealStartDate").Value = Format(gdaRealStartDate.GetValue, "####/##/##")
                .Fields("RealStopDate").Value = Format(gdaRealStopDate.GetValue, "####/##/##")
            Else
                .Fields("RealPeriod").Value = 0
                .Fields("RealStartDate").Value = Null
                .Fields("RealStopDate").Value = Null
            End If
            
            .Fields("EntryEn").Value = garyGi(0)
            .Fields("Note").Value = txtNote & ""
            .Fields("CancelFlag").Value = 0
            
            If gilClctEn.GetDescription <> "" Then
                .Fields("ClctEn").Value = gilClctEn.GetCodeNo
                .Fields("ClctName").Value = gilClctEn.GetDescription
            Else
                .Fields("ClctEn").Value = Null
                .Fields("ClctName").Value = Null
            End If
            
            .Fields("CmCode").Value = gilCMCode.GetCodeNo
            .Fields("CMName").Value = gilCMCode.GetDescription
                            
            If gilPTCode.GetDescription <> "" Then
                .Fields("PTCode").Value = gilPTCode.GetCodeNo
                .Fields("PTName").Value = gilPTCode.GetDescription
            Else
                .Fields("PTCode").Value = Null
                .Fields("PTName").Value = Null
            End If
            
            .Fields("RowID").Value = strRowId
            
            If gnbRealAmt.Value > 0 Then
                .Fields("RealAmt").Value = NoZero(gnbRealAmt.Value, True)
            Else
                .Fields("RealAmt").Value = 0
            End If
            
            If gilSTCode.GetDescription <> "" Then
                .Fields("STCode").Value = gilSTCode.GetCodeNo
                .Fields("STName").Value = gilSTCode.GetDescription
            Else
                .Fields("STCode").Value = Null
                .Fields("STName").Value = Null
            End If
            '.Fields("ManualNo") = NoZero(txtManualNo)
            .Fields("ServiceType") = gServiceType
            .Fields("CompCode") = strCompCode
            .Update
        End With
        rsMdbRec.Close
        Set rsMdbRec = Nothing
    Exit Function
ChkErr:
    Call ErrSub(Me.Name, "ScrToRcd")
End Function


Private Function GetNumberType(ByVal strAmount As String) As Long
    On Error Resume Next
        GetNumberType = Val(Format(strAmount, "00000000"))
End Function

Private Sub txtNote_Validate(Cancel As Boolean)
    On Error Resume Next
        txtNote.IMEMode = 2
End Sub

Private Sub subGil()
    On Error GoTo ChkErr
        SetgiList gilCitemCode, "CodeNo", "Description", "CD019", , , , , 3, 12, "Where ServiceType ='" & gServiceType & "'"
        gilCitemCode.Filter = "Where ServiceType ='" & gServiceType & "'"
        gilCitemCode.ListIndex = 1
        SetgiList gilCMCode, "CodeNo", "Description", "CD031", , , , , 3, 12
        SetgiList gilUCCode, "CodeNo", "Description", "CD013", , , , , 3, 12
        SetgiList gilSTCode, "CodeNo", "Description", "CD016", , , , , 3, 12
        SetgiList gilClctEn, "EmpNo", "EmpName", "CM003", , , , , 10, 20
        SetgiList gilPTCode, "CodeNo", "Description", "CD032", , , , , 3, 12
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "subGil")
End Sub

'������SO3311b�ǨӤ�ROWID �ק�ɨϥΡI
Public Property Let uRowId(ByVal vData As String)
    strRowId = vData
End Property

'������So3311B�ǨӤ���ڽs���ΦL��Ǹ��I�s�W�ɨϥΡI
Public Property Let uBillNo(ByVal vData As String)
    strBillNo = vData
End Property

'������So3311b�ǨӤ��L��Ǹ�
Public Property Let uPrtSNo(ByVal vData As String)
    strPrtSno = vData
End Property

'��So3311B�ǨӤ����ؽs��
Public Property Let uItem(ByVal vData As Integer)
    intItem = vData
End Property

'������So3311b�ǨӤ��A�����O
Public Property Let uServiceType(ByVal vData As String)
    strServiceType = vData
End Property

'������So3311b�ǨӤ����q�O
Public Property Let uCompCode(ByVal vData As String)
    strCompCode = vData
End Property
