VERSION 5.00
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{3BB22E57-824A-11D3-BD70-0080C8F80BC4}#3.4#0"; "GiGridR.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Begin VB.Form frmSO1148A 
   BorderStyle     =   1  '��u�T�w
   Caption         =   "�Ȥ��ث~�O�� [SO1148A]"
   ClientHeight    =   6810
   ClientLeft      =   45
   ClientTop       =   1530
   ClientWidth     =   11880
   Icon            =   "SO1148A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraData 
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3030
      Left            =   150
      TabIndex        =   16
      Top             =   30
      Width           =   11595
      Begin VB.CommandButton cmdHouse 
         Caption         =   "�d�߳]�Ʃ���"
         Height          =   315
         Left            =   3360
         TabIndex        =   33
         Top             =   1770
         Width           =   1305
      End
      Begin VB.TextBox txtFaciSNo 
         Height          =   315
         Left            =   1740
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   32
         Top             =   1770
         Width           =   1575
      End
      Begin VB.CheckBox chkBack 
         Caption         =   "�h�^�O"
         Height          =   180
         Left            =   6690
         TabIndex        =   13
         Top             =   2205
         Width           =   885
      End
      Begin VB.TextBox txtCVSno 
         Height          =   315
         Left            =   7740
         MaxLength       =   15
         TabIndex        =   10
         Top             =   1020
         Width           =   2940
      End
      Begin VB.TextBox txtOrderNo 
         Height          =   315
         Left            =   1740
         MaxLength       =   15
         TabIndex        =   1
         Top             =   270
         Width           =   2940
      End
      Begin VB.TextBox txtSumAmount 
         Height          =   315
         Left            =   1740
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   5
         Top             =   1410
         Width           =   1560
      End
      Begin VB.TextBox txtAmount 
         Height          =   315
         Left            =   3450
         MaxLength       =   3
         TabIndex        =   4
         Top             =   1020
         Width           =   1230
      End
      Begin VB.TextBox txtPrice 
         Height          =   315
         Left            =   1740
         MaxLength       =   8
         TabIndex        =   3
         Top             =   1020
         Width           =   1125
      End
      Begin VB.TextBox txtClass 
         Height          =   315
         Left            =   7740
         MaxLength       =   10
         TabIndex        =   8
         Top             =   270
         Width           =   2940
      End
      Begin VB.TextBox txtMake 
         Height          =   315
         Left            =   7740
         MaxLength       =   15
         TabIndex        =   9
         Top             =   645
         Width           =   2940
      End
      Begin prjGiList.GiList gilArticle 
         Height          =   315
         Left            =   1740
         TabIndex        =   2
         Top             =   645
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
      End
      Begin prjGiList.GiList gilProcEn 
         Height          =   315
         Left            =   7740
         TabIndex        =   12
         Top             =   1770
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
      End
      Begin Gi_Date.GiDate gdaCVSdate 
         Height          =   315
         Left            =   7740
         TabIndex        =   11
         Top             =   1395
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
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
      Begin Gi_Date.GiDate gdaBackDate 
         Height          =   315
         Left            =   9600
         TabIndex        =   14
         Top             =   2145
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
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
      Begin prjGiList.GiList gilCPCode 
         Height          =   315
         Left            =   1740
         TabIndex        =   7
         Top             =   2520
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
      Begin prjGiList.GiList gilPackage 
         Height          =   315
         Left            =   1740
         TabIndex        =   6
         Top             =   2145
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
      End
      Begin prjGiList.GiList gilBackEn 
         Height          =   315
         Left            =   7740
         TabIndex        =   15
         Top             =   2520
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
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "�]�ƧǸ�"
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
         Left            =   735
         TabIndex        =   31
         Top             =   1830
         Width           =   780
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "�p�p"
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
         Left            =   1110
         TabIndex        =   30
         Top             =   1485
         Width           =   405
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "���"
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
         Left            =   1110
         TabIndex        =   29
         Top             =   1095
         Width           =   405
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "�ƶq"
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
         Left            =   2910
         TabIndex        =   28
         Top             =   1095
         Width           =   405
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "�g��h�^�H��"
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
         Left            =   6360
         TabIndex        =   27
         Top             =   2580
         Width           =   1185
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "�h�^���"
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
         Left            =   8610
         TabIndex        =   26
         Top             =   2205
         Width           =   780
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "�I���鸹"
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
         Left            =   6750
         TabIndex        =   25
         Top             =   1095
         Width           =   795
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "�M�\"
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
         Left            =   1110
         TabIndex        =   24
         Top             =   2205
         Width           =   405
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "�u�ݿ�k"
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
         Left            =   720
         TabIndex        =   23
         Top             =   2580
         Width           =   795
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "�I�����"
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
         Left            =   6750
         TabIndex        =   22
         Top             =   1470
         Width           =   795
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "�B�z�H��"
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
         Left            =   6750
         TabIndex        =   21
         Top             =   1845
         Width           =   795
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "�q��渹"
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
         Left            =   720
         TabIndex        =   20
         Top             =   360
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "�ӫ~"
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
         Left            =   1110
         TabIndex        =   19
         Top             =   735
         Width           =   405
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "���O"
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
         Left            =   7140
         TabIndex        =   18
         Top             =   360
         Width           =   405
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "����"
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
         Left            =   7140
         TabIndex        =   17
         Top             =   735
         Width           =   405
      End
   End
   Begin prjGiGridR.GiGridR ggrData 
      Height          =   3540
      Left            =   135
      TabIndex        =   0
      ToolTipText     =   "�Ы�����⦸,����Ȥ���"
      Top             =   3150
      Width           =   11610
      _ExtentX        =   20479
      _ExtentY        =   6244
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
End
Attribute VB_Name = "frmSO1148A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Programmer: Power Hammer
' Last Modify: 2002/09/04

Option Explicit
Private lngEditMode As giEditModeEnu ' �O���ثe�b�s��B�s�W���˵��Ҧ�
Private WithEvents rsSO105C As ADODB.Recordset
Attribute rsSO105C.VB_VarHelpID = -1
Dim rsCD059 As New ADODB.Recordset

'�O���ثe�ϥΪ��v��
Private blnUserPrivAddNew   As Boolean   '�s�W
Private blnUserPrivUpdate   As Boolean  '�ק�

Private Sub cmdHouse_Click()
  On Error GoTo ChkErr
    With frmSO1131E
        .uParentForm = Me
        If Not rsSO105C.EOF Then
            .uCustId = rsSO105C("Custid")
        Else
            .uCustId = gCustId
        End If
        .uEditMode = lngEditMode
        .uDefFaciSeqNo = txtFaciSNo.Tag
        .uPRCanClick = True
        If Not rsSO105C.EOF Then
            .uServiceType = rsSO105C("ServiceType")
        Else
            .uServiceType = gServiceType
        End If
        .Show vbModal
        If lngEditMode <> giEditModeView Then
            If .uFaciSeqNo <> "" Then
                txtFaciSNo.Tag = .uFaciSeqNo
                txtFaciSNo.Text = .uFaciSNo
                txtFaciSNo.ToolTipText = "�]�Ƭy����:" & txtFaciSNo.Text
            End If
        End If
    End With
    Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdHouse_Click")
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
    ReleaseCOM Me
End Sub

Private Sub gilArticle_Change()
  On Error GoTo ChkErr
    If Len(gilArticle.GetCodeNo) > 0 Then
        With rsCD059
            .Filter = "CodeNo=" & gilArticle.GetCodeNo
            txtPrice.Text = .Fields("UnitPrice")
            txtAmount.Text = .Fields("Quantity")
            txtClass.Text = .Fields("Class") & ""
            txtMake.Text = .Fields("Make") & ""
        End With
    End If
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "gilArticle_Change")
End Sub

Private Sub RcdToScr()
  On Error GoTo ChkErr
    With rsSO105C
        txtOrderNo = .Fields("OrderNo") & ""
        gilArticle.SetCodeNo .Fields("ArticleNo") & ""
        gilArticle.SetDescription .Fields("ArticleName") & ""
        txtPrice.Text = .Fields("Price") & ""
        txtAmount.Text = .Fields("Amount") & ""
        txtSumAmount.Text = .Fields("SubTotal") & ""
        gilPackage.SetCodeNo .Fields("PackageNo") & ""
        gilPackage.SetDescription .Fields("PackageName") & ""
        txtClass.Text = .Fields("Class") & ""
        txtMake.Text = .Fields("Make") & ""
        txtCVSno = .Fields("ConversionNo") & ""
        gdaCVSdate.Text = .Fields("ConversionDate") & ""
        gilProcEn.SetDescription .Fields("ProcEn") & ""
        gilProcEn.Query_CodeNo
        gdaBackDate.Text = .Fields("BackDate") & ""
        gilBackEn.SetDescription .Fields("BackEn") & ""
        gilBackEn.Query_CodeNo
        gilCPCode.SetCodeNo .Fields("CPCode") & ""
        gilCPCode.SetDescription .Fields("CPName") & ""
        chkBack.Value = Val(.Fields("BackGift"))
        '#6118 �W�[�]�ƧǸ� By Kin 2011/09/27
        txtFaciSNo.Tag = rsSO105C("FaciSeqNo") & ""
        txtFaciSNo.Text = GetFaciSNo(rsSO105C("FaciSeqNo") & "")
    End With
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "RcdToScr")
End Sub

Private Function ScrToRcd() As Boolean
  On Error GoTo ChkErr
    With rsSO105C
        If EditMode = giEditModeInsert Then .AddNew
        .Fields("OrderNo") = txtOrderNo
        .Fields("ArticleNo") = gilArticle.GetCodeNo2
        .Fields("ConversionNo") = txtCVSno.Text
        .Fields("ConversionDate") = IIf(gdaCVSdate.Text = "", Null, gdaCVSdate.Text)
        .Fields("ArticleName") = gilArticle.GetDescription2
        .Fields("PackageNo") = gilPackage.GetCodeNo2
        .Fields("PackageName") = gilPackage.GetDescription2
        .Fields("ArticleNo") = gilArticle.GetCodeNo2
        .Fields("ArticleName") = gilArticle.GetDescription2
        .Fields("Price") = CInt(Trim(txtPrice.Text))
        .Fields("Amount") = Val(Trim(txtAmount.Text))
        .Fields("SubTotal") = GetSum
        .Fields("FreeGift") = IIf(Trim(txtPrice.Text) = 0, 1, 0)
        .Fields("Class") = Trim(txtClass)
        .Fields("Make") = Trim(txtMake)
        .Fields("ProcEn") = gilProcEn.GetDescription2
        .Fields("BackDate") = IIf(gdaBackDate.Text = "", Null, gdaBackDate.Text)
        .Fields("BackEn") = gilBackEn.GetDescription2
        .Fields("CPCode") = gilCPCode.GetCodeNo2
        .Fields("CPName") = gilCPCode.GetDescription2
        .Fields("CustId") = gCustId
        .Fields("FREEGIFT") = 1
        .Fields("ServiceType") = gServiceType
        .Fields("BackGift") = chkBack.Value
        '#6118 �W�[�]�ƧǸ���� By Kin 2011/09/27
        If txtFaciSNo.Text <> "" Then
            .Fields("FaciSeqNo") = txtFaciSNo.Tag
        Else
            .Fields("FaciSeqNo") = Null
        End If
        .Update
    End With
    ScrToRcd = True
  Exit Function
ChkErr:
  ScrToRcd = False
  gcnGi.RollbackTrans
  Call ErrSub(Me.Name, "ScrToRcd")
End Function

Private Function GetSum() As Long
  On Error GoTo ChkErr
    If Len(Trim(txtPrice.Text)) > 0 And Len(Trim(txtAmount.Text)) > 0 Then
        GetSum = Trim(txtPrice.Text) * Trim(txtAmount.Text)
    Else
        GetSum = 0
    End If
  Exit Function
ChkErr:
  Call ErrSub(Me.Name, "GetSum")
End Function

Public Sub FunctionKey(KeyCode As Integer, Shift As Integer)
  On Error GoTo ChkErr
    Select Case KeyCode
           '----------------------------------------------------
        Case vbKeyF2  '   F3:�s��, �۷����Ucmdsave
            If Not ChkGiList(KeyCode) Then Exit Sub
            UpdateGo
           '----------------------------------------------------
        Case vbKeyEscape
            CancelGo
    End Select
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "FunctionKey")
End Sub

'Private Function chkOnly() As Boolean
'  On Error GoTo ChkErr
'  Dim lngAbs As Long
'    With rsArticle
'        If .RecordCount > 0 Then
'            lngAbs = .AbsolutePosition
'            .MoveFirst
'            While Not .EOF
'                If .Fields("ArticleNo") = gilArticle.GetCodeNo Then Exit Function
'                .MoveNext
'            Wend
'            .AbsolutePosition = lngAbs
'        End If
'    End With
'    chkOnly = True
'  Exit Function
'ChkErr:
'  Call ErrSub(Me.Name, "IsDataOK")
'End Function

Private Function IsDataOk() As Boolean
  On Error GoTo ChkErr
    If gilArticle.GetCodeNo = "" Then gilArticle.SetFocus: strErrFile = "���~": GoTo 66
    If txtPrice.Text = "" Then txtPrice.SetFocus: strErrFile = "����": GoTo 66
    If txtAmount.Text = "" Then txtAmount.SetFocus: strErrFile = "���B": GoTo 66
    IsDataOk = True
  Exit Function
66:
  MsgMustBe (strErrFile)
  Exit Function
ChkErr:
  Call ErrSub(Me.Name, "IsDataOK")
End Function

Private Sub NewRcd()
  On Error Resume Next
    Dim ctl As Control
    For Each ctl In Me
        If TypeOf ctl Is TextBox Then ctl = ""
        If TypeOf ctl Is GiDate Then Call ctl.SetValue("")
        If TypeOf ctl Is GiTime Then ctl.Text = ""
        If TypeOf ctl Is GiList Then ctl.Clear
        If TypeOf ctl Is CheckBox Then ctl.Value = 0
        If TypeOf ctl Is MaskEdBox Then ctl.Text = ""
    Next
    txtOrderNo.SetFocus
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "NewRcd")
End Sub

Public Sub AddNewGo()
  On Error GoTo ChkErr
    ChangeMode (giEditModeInsert)
    NewRcd
  Exit Sub
ChkErr:
   Call ErrSub(Me.Name, "cmdAdd_Click")
End Sub

Public Sub UpdateGo()
  On Error GoTo ChkErr
    Dim strSQL As String
    If IsDataOk() = False Then Exit Sub
    gcnGi.BeginTrans
    
    If Not ScrToRcd Then Exit Sub           ' �ⱱ������ȡA�s���Ʈw��
   
   'rsSO105C.Requery
    If rsSO105C.State = adStateOpen Then
        strSQL = rsSO105C.Source
        rsSO105C.Close
        rsSO105C.Open strSQL, gcnGi, adOpenKeyset, adLockOptimistic
        Set ggrData.Recordset = rsSO105C
    End If
    
    ggrData.Refresh
    gcnGi.CommitTrans
    If EditMode = giEditModeInsert Then '�~��s�W
        Call ChangeMode(giEditModeInsert)
        AddNewGo
    Else '�i�J��ܼҦ�
        Call ChangeMode(giEditModeView)
        RcdToScr
    End If
   Exit Sub
ChkErr:
    gcnGi.RollbackTrans
    Call ErrSub(Me.Name, "UpdateGo")
End Sub

Private Sub Form_Activate()
  On Error Resume Next
    Screen.MousePointer = vbDefault
    Set ObjActiveForm = Me
    strActFrmName = "frmSO1148A"
End Sub

Private Sub Form_KeyPress(keyAscii As Integer)
  On Error Resume Next
   If keyAscii = Asc("'") Then keyAscii = 0
End Sub

Private Sub Form_Load()
  On Error GoTo ChkErr
    Dim mFlds As New GiGridFlds
   
    InitializingListOcx
    If Crm Then
        PolyFrmFunction Me
        Me.Move frmSO1100BMDI.Left + ((frmSO1100BMDI.Width - Me.Width) / 2), frmSO1100BMDI.Top + ((frmSO1100BMDI.Height - Me.Height) / 2)
    Else
        If Not800600 Then Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2 + 300
    End If
    
    GetRS rsCD059, "Select CodeNo,Description,Class,Make,Quantity,UnitPrice From " & GetOwner & "CD059 Where CompCode=" & gCompCode, gcnGi, adUseClient
    
    Set rsSO105C = New ADODB.Recordset
    With rsSO105C
         If .State = 1 Then
            .Close
         Else
            .CursorLocation = adUseClient
         End If
        .Open "SELECT ROWID,SO105C.* FROM " & GetOwner & "SO105C WHERE CUSTID= " & gCustId & " AND SERVICETYPE='" & gServiceType & "' AND FREEGIFT=1 ORDER BY ARTICLENO,ORDERNO", gcnGi, adOpenKeyset, adLockOptimistic
    End With
    
    With mFlds
        .Add "OrderNo", , , , , "�q��渹           "
        .Add "ArticleName", , , , , "�ث~�W��  "
        .Add "PackageName", , , , , "�M�\�W��  "
        .Add "CPName", , , , , "�u�ݦW�� "
        .Add "Class", , , , , "���O    "
        .Add "Make", , , , , "����    "
        .Add "ConversionNo", , , , , "�I������    "
        .Add "ConversionDate", giControlTypeDate, , , , "  �I�����  "
        .Add "ProcEn", , , , , "�B�z�H��"
        .Add "BackGift", , , , , "�h�^�O "
        .Add "BackDate", giControlTypeDate, , , , "  �h�^���  "
        .Add "BackEn", , , , , "�g��h�^�H��"
    End With
         
    ggrData.AllFields = mFlds
    ggrData.SetHead
    Set ggrData.Recordset = rsSO105C
    ggrData.Refresh
   'ggrData.Blank = True
    Call ChangeMode(giEditModeView)
    
'    Dim strPermis As String
'    Call GetPermission(garyGi(4), "SO1162", strPermis)
'    Call SetPermission(strPermis, blnAdd:=blnUserPrivAddNew, AddKey:="SO11621" _
'                                , blnUpdate:=blnUserPrivUpdate, UpdKey:="SO11622")
                                    
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "Form_Load")
End Sub
'Private Function GetFaciSNo(ByVal aSeqNo As String) As String
'  On Error GoTo ChkErr
'    Dim aSQL As String
'    aSQL = "SELECT FACISNO FROM " & GetOwner & "SO004 WHERE SEQNO = '" & aSeqNo & "'"
'    GetFaciSNo = GetRsValue(aSQL, gcnGi) & ""
'    Exit Function
'ChkErr:
'  Call ErrSub(Me, "GetFaciSNo")
'End Function
Private Sub InitializingListOcx()
  On Error GoTo ChkErr
    Call SetgiList(gilArticle, "CodeNo", "Description", "CD059", , , , , , , " Where ServiceType='" & gServiceType & "'")
    Call SetgiList(gilCPCode, "CodeNo", "Description", "CD048")
    Call SetgiList(gilProcEn, "EmpNo", "EmpName", "CM003")
    Call SetgiList(gilBackEn, "EmpNo", "EmpName", "CM003")
    Call SetgiList(gilPackage, "CodeNo", "Description", "CD027")
  Exit Sub
ChkErr:
    ErrSub Me.Name, "InitializingListOcx"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error GoTo ChkErr
    Call FunctionKey(KeyCode, Shift)
  Exit Sub
ChkErr:
       Call ErrSub(Me.Name, "Form_KeyDown")
End Sub
 
Public Sub EditGo()
  On Error GoTo ChkErr
    Call ChangeMode(giEditModeEdit)
    On Error Resume Next
    txtOrderNo.SetFocus
  Exit Sub
ChkErr:
   Call ErrSub(Me.Name, "cmdEdit_Click")
End Sub

Public Sub CancelGo()
   On Error GoTo ChkErr
   If EditMode = giEditModeView Then
        Unload Me
        Exit Sub
   Else
      If EditMode = giEditModeEdit Then
         RevertRcd
      Else
         If Not rsSO105C.EOF Then
            rsSO105C.MoveFirst
            RcdToScr
         Else
         End If
      End If
      Call ChangeMode(giEditModeView)
   End If
 Exit Sub
ChkErr:
   Call ErrSub(Me.Name, "CancelGo")
End Sub

Private Sub RevertRcd() '�٭��Ƥ��e
 On Error GoTo ChkErr
   RcdToScr
 Exit Sub
ChkErr:
   Call ErrSub(Me.Name, "RevertRcd")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo ChkErr
        If EditMode <> giEditModeView Then
            If giMsgNotSave = vbNo Then
                Cancel = True: Exit Sub
            Else
                CancelGo
            End If
        End If
        rsSO105C.Close
        Set rsSO105C = Nothing
        Call FormQueryUnload
    Exit Sub
ChkErr:
    ErrSub Me.Name, "Form_QueryUnload"
End Sub

Private Sub ggrData_ShowCellData(ByVal Col As Long, ByVal Row As Long, ByVal giFld As prjGiGridR.GiGridFld, ByVal Fld As ADODB.Field, Value As Variant)
  On Error GoTo ChkErr
    If Fld.Name = UCase("BACKGIFT") Then Value = IIf(Value = 0, "�_", "�O")
  Exit Sub
ChkErr:
    ErrSub Me.Name, "ggrData_ShowCellData"
End Sub

Private Sub rsSO105C_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  On Error GoTo ChkErr
    If adReason = 10 Then
        If Not rsSO105C.EOF And Not rsSO105C.BOF And lngEditMode = giEditModeView Then
            If Not ggrData.Enabled Then ggrData.Enabled = True
            RcdToScr
        Else
            ggrData.Enabled = False
        End If
    End If
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "rsSO105C_MoveComplete")
End Sub

Public Property Get EditMode() As giEditModeEnu
  On Error GoTo ChkErr
    EditMode = lngEditMode    ' ���ثe�s��Ҧ�
  Exit Property
ChkErr:
    Call ErrSub(Me.Name, "Get EditMode")
End Property

Public Property Let EditMode(ByVal vNewValue As giEditModeEnu)
  On Error GoTo ChkErr
    lngEditMode = vNewValue '�]�w�s��Ҧ�
  Exit Property
ChkErr:
    Call ErrSub(Me.Name, "Let EditMode")
End Property

Private Sub ChangeMode(ByVal lngMode As giEditModeEnu) '�ھڶǨӤ��s��Ҧ�, �]�w�U�����ݩ�
  On Error GoTo ChkErr
    Dim blnFlag As Boolean  ' �O���O�_�b����s���Ҧ��A
    lngEditMode = lngMode
    Select Case lngMode
           Case giEditModeInsert
                blnFlag = False
                MenuEnabled , , False, True, False, True
           Case giEditModeEdit
                blnFlag = False
                MenuEnabled , , False, True, False, True
           Case giEditModeView
                blnFlag = True
                MenuEnabled True, rsSO105C.RecordCount > 0, False, , False, True
    End Select
    SetStatusBar lngEditMode
    fraData.Enabled = Not blnFlag
    ggrData.Enabled = blnFlag
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "ChangeMode")
End Sub

Private Sub txtClass_Change()
  On Error Resume Next
    CML
End Sub

Private Sub txtCVSno_Change()
  On Error Resume Next
    CML
End Sub

Private Sub txtMake_Change()
  On Error Resume Next
    CML
End Sub

Private Sub txtOrderNo_Change()
  On Error Resume Next
    CML
End Sub
