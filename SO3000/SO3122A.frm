VERSION 5.00
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmSO3122A 
   BorderStyle     =   1  '單線固定
   Caption         =   "應收彙總 [SO3122A]"
   ClientHeight    =   3945
   ClientLeft      =   2220
   ClientTop       =   2250
   ClientWidth     =   6525
   BeginProperty Font 
      Name            =   "新細明體"
      Size            =   9.75
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SO3122A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   6525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '螢幕中央
   Begin VB.ComboBox cboSort 
      Height          =   315
      ItemData        =   "SO3122A.frx":0442
      Left            =   1140
      List            =   "SO3122A.frx":044C
      Style           =   2  '單純下拉式
      TabIndex        =   17
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton cmdPrevRpt 
      Caption         =   "上次統計結果"
      Height          =   495
      Left            =   2610
      TabIndex        =   15
      Top             =   3390
      Width           =   1395
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      Height          =   495
      Left            =   5160
      TabIndex        =   14
      Top             =   3390
      Width           =   1275
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "F5:列印"
      Height          =   495
      Left            =   180
      TabIndex        =   13
      Top             =   3390
      Width           =   1275
   End
   Begin MSMask.MaskEdBox mskBillNo2 
      Height          =   345
      Left            =   3180
      TabIndex        =   1
      ToolTipText     =   "YYYYMMBxxxxxxxxx"
      Top             =   915
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   609
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "######AA#######"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskBillNo1 
      Height          =   345
      Left            =   1200
      TabIndex        =   0
      ToolTipText     =   "YYYYMMBxxxxxxxxx"
      Top             =   915
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   609
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "######AA#######"
      PromptChar      =   "_"
   End
   Begin VB.Frame fraStatus 
      Caption         =   "客戶類別"
      Height          =   615
      Left            =   2880
      TabIndex        =   8
      Top             =   2520
      Width           =   3525
      Begin VB.OptionButton optNormal 
         Caption         =   "一般戶"
         Height          =   195
         Left            =   210
         TabIndex        =   2
         Top             =   300
         Width           =   975
      End
      Begin VB.OptionButton optMdu 
         Caption         =   "大樓戶"
         Height          =   195
         Left            =   1380
         TabIndex        =   3
         Top             =   300
         Width           =   975
      End
      Begin VB.OptionButton optAll 
         Caption         =   "全部"
         Height          =   195
         Left            =   2580
         TabIndex        =   4
         Top             =   300
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin prjGiList.GiList gilStrtCode 
      Height          =   315
      Left            =   1200
      TabIndex        =   5
      Top             =   1350
      Width           =   2940
      _ExtentX        =   5186
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      F5Corresponding =   -1  'True
   End
   Begin prjGiList.GiList gilClctEn 
      Height          =   315
      Left            =   1200
      TabIndex        =   6
      Top             =   1770
      Width           =   2940
      _ExtentX        =   5186
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      F5Corresponding =   -1  'True
   End
   Begin prjGiList.GiList gilCompCode 
      Height          =   315
      Left            =   1200
      TabIndex        =   18
      Top             =   90
      Width           =   2940
      _ExtentX        =   5186
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.75
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
      Left            =   1200
      TabIndex        =   19
      Top             =   495
      Width           =   2940
      _ExtentX        =   5186
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      F2Corresponding =   -1  'True
   End
   Begin prjGiList.GiList giloldClctEn 
      Height          =   315
      Left            =   1200
      TabIndex        =   7
      Top             =   2160
      Width           =   2940
      _ExtentX        =   5186
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      F5Corresponding =   -1  'True
   End
   Begin VB.Label lbloldClctEn 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "原收費人員"
      Height          =   195
      Left            =   120
      TabIndex        =   22
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label lblCompCode 
      AutoSize        =   -1  'True
      Caption         =   "公司別"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   210
      TabIndex        =   21
      Top             =   180
      Width           =   585
   End
   Begin VB.Label lblServiceType 
      AutoSize        =   -1  'True
      Caption         =   "服務類別"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   210
      TabIndex        =   20
      Top             =   600
      Width           =   780
   End
   Begin VB.Label lblSort 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "排序方式"
      Height          =   195
      Left            =   90
      TabIndex        =   16
      Top             =   2850
      Width           =   780
   End
   Begin VB.Label lblBillNo 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "單號編號"
      Height          =   195
      Left            =   210
      TabIndex        =   12
      Top             =   1035
      Width           =   780
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "---"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   2940
      TabIndex        =   11
      Top             =   1020
      Width           =   180
   End
   Begin VB.Label lblStrt 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "街         道"
      Height          =   195
      Left            =   210
      TabIndex        =   10
      Top             =   1455
      Width           =   795
   End
   Begin VB.Label lblClctEn 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "收費人員"
      Height          =   195
      Left            =   210
      TabIndex        =   9
      Top             =   1890
      Width           =   780
   End
End
Attribute VB_Name = "frmSO3122A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private cnn As New ADODB.Connection
Private strChooseString As String
Private strChoose As String
Private strForm As String
Private strFormula As String

Private Sub cmdExit_Click()
  On Error GoTo ChkErr
   Unload Me
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "cmdExit_Click")
End Sub

Private Sub cmdPrevRpt_Click()
    On Error GoTo ChkErr
        Screen.MousePointer = vbHourglass
        Call PreviousRpt(GetPrinterName(5), "SO3122A.Rpt", Me.Caption)
        Screen.MousePointer = vbDefault
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "cmdPrevRpt_Click")
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
  On Error GoTo ChkErr
   Set cnn = GetTmpMdbCn
   Call subGil
   cboSort.ListIndex = 0
    gilCompCode.SetCodeNo garyGi(9)
    gilCompCode.Query_Description
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "Form_Load")
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo ChkErr
        If Not ChkGiList(KeyCode) Then Exit Sub
        Call FunctionKey(KeyCode, Shift)
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "FunctionKey")
End Sub

Public Sub FunctionKey(KeyCode As Integer, Shift As Integer)
  On Error GoTo ChkErr
    Select Case KeyCode
           '----------------------------------------------------
           Case vbKeyF5 '   F5: 列印 相當於按下cmdPrint
                        If cmdPrint.Enabled = True Then cmdPrint.Value = True
           Case vbKeyEscape
                        If cmdExit.Enabled = True Then cmdExit.Value = True
           '----------------------------------------------------
    End Select
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "FunctionKey")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
        Call ReleaseCOM(Me)

End Sub

Private Sub gilCompCode_Change()
    On Error Resume Next
        If Not ChgComp(gilCompCode, "SO3120", "SO3122") Then Exit Sub
        Call subGil
        Call SelectServType(gilCompCode.GetCodeNo, gilServiceType)
        gilServiceType.SetCodeNo ""
        gilServiceType.Query_Description
        gilServiceType.ListIndex = 1
        Call GiListFilter(gilStrtCode, , gilCompCode)
        
End Sub

Private Sub mskBillNo1_GotFocus()
  On Error GoTo ChkErr
   mskBillNo1.SelStart = 0
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "mskBillNo1_GotFocus")
End Sub

Private Sub mskBillNo1_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub mskBillNo2_GotFocus()
  On Error GoTo ChkErr
   mskBillNo2.SelStart = 0
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "mskBillNo2_GotFocus")
End Sub

Private Sub mskBillNo2_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo ChkErr
    Dim strSQL As String
    Dim lngAffected As Long
        If Not IsDataOk Then Exit Sub
        If Not ChkDTok Then Exit Sub
        Screen.MousePointer = vbHourglass
        cmdExit.SetFocus
        ReadyGoPrint
        Call subChoose
        If Not InsertToMDB(lngAffected) Then Exit Sub
        If lngAffected = 0 Then
            MsgNoRcd
        Else
            strSQL = "Select * from So3122"
            Call PrintRpt(GetPrinterName(5), "SO3122A.rpt", "SO3122", Me.Caption, strSQL, strChooseString, cboSort.Text, True, "Tmp0000.MDB", , strFormula)
        End If
        Screen.MousePointer = vbDefault
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "cmdPrint_Click")
End Sub

Private Function IsDataOk() As Boolean
    On Error GoTo ChkErr
        If Not MustExist(gilCompCode, 2, "公司別") Then Exit Function
        If Not MustExist(gilServiceType, 2, "服務類別") Then Exit Function
        IsDataOk = True
    Exit Function
ChkErr:
    ErrSub Me.Name, "IsDataOk"
End Function


Private Function InsertToMDB(lngAffected As Long) As Boolean
    On Error GoTo ChkErr
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
        cnn.BeginTrans
        strSQL = "SELECT SO032.CustId,SO032.SHOULDDATE,So032.ShouldAmt,So032.AddrNo,SO032.CLCTEN," & _
            "SO032.CLCTNAME,SO001.CUSTNAME,SO001.TEL1,SO014.Address,So032.StrtCode,CD017.Description StrtName " & strChoose
        If Not GetRS(rsTmp, strSQL) Then Exit Function
        If Not CreateMDBTable(rsTmp, "SO3122", cnn) Then Exit Function
        Do While Not rsTmp.EOF
            cnn.Execute GetTmpMDBExecuteStr(rsTmp, "SO3122")
            rsTmp.MoveNext
            DoEvents
        Loop
        cnn.CommitTrans
        lngAffected = rsTmp.RecordCount
        InsertToMDB = True
    Exit Function
ChkErr:
    cnn.RollbackTrans
    ErrSub Me.Name, "InsertToMdb"
End Function

Private Sub subChoose()
  '串 SelectCondition
    On Error GoTo ChkErr
    Dim strCustStatus As String
    strChoose = " From " & GetOwner & "SO001 SO001," & GetOwner & "SO014 SO014, " & GetOwner & _
        IIf(strForm = "SO3122", "SO032 ", "SO033 ") & " SO032," & GetOwner & "CD017 CD017 Where " & _
        "SO032.CUSTID=SO001.CUSTID And SO032.CompCode=SO001.CompCode AND SO032.ADDRNO=SO014.ADDRNO And " & _
        "SO032.CompCode=SO014.CompCode And So032.StrtCode=CD017.CodeNo And ClctYM is Not Null "
    
    If mskBillNo1 <> "" Then Call subAnd("SO032.BillNo>='" & mskBillNo1 & "'")
    If mskBillNo1 <> "" Then Call subAnd("SO032.BillNo<='" & mskBillNo2 & "'")
    If optNormal.Value Then Call subAnd("SO032.MduId is Null")
    If optMdu.Value Then Call subAnd("SO032.MduId is not Null")
    '#1659  950517原"收費人員"改為"原收費人員"，另新增一"收費人員"
    If gilClctEn.GetCodeNo <> "" Then Call subAnd("SO032.ClctEn='" & gilClctEn.GetCodeNo & "'")
    If gilOldClctEn.GetCodeNo <> "" Then Call subAnd("SO032.oldClctEn='" & gilOldClctEn.GetCodeNo & "'")
    If gilStrtCode.GetCodeNo <> "" Then Call subAnd("SO014.StrtCode='" & gilStrtCode.GetCodeNo & "'")
    If gilCompCode.GetCodeNo <> "" Then Call subAnd("SO032.CompCode =" & gilCompCode.GetCodeNo)
    If gilServiceType.GetCodeNo <> "" Then Call subAnd("SO032.ServiceType ='" & gilServiceType.GetCodeNo & "'")
    
    If cboSort.ListIndex = 0 Then
        strFormula = "GroupName1={SO3122.STRTCODE};GroupName2={SO3122.ClctEn}"
    ElseIf cboSort.ListIndex = 1 Then
        strFormula = "GroupName1={SO3122.ClctEn};GroupName2={SO3122.STRTCODE}"
    End If
    If optNormal.Value Then
       strCustStatus = "一般戶"
    Else
       If optMdu.Value Then
          strCustStatus = "大樓戶"
       Else
          strCustStatus = "全  部"
       End If
    End If
    strChooseString = "收集單號: " & subSpace(Me.mskBillNo1) & "~" & subSpace(Me.mskBillNo2) & ";" & _
                      "客戶類別: " & strCustStatus & ";" & _
                      "街道　　: " & subSpace(gilStrtCode.GetDescription) & ";" & _
                      "收費員: " & subSpace(gilClctEn.GetDescription) & ";" & _
                      "原收費員: " & subSpace(gilOldClctEn.GetDescription)
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "subChoose")
End Sub

Private Sub subAnd(strCH As String)
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

Private Sub subGil()
  On Error GoTo ChkErr
    Call SetgiList(gilClctEn, "EmpNo", "EmpName", "CM003", , , , , 10, 20, , True)
    Call SetgiList(gilOldClctEn, "EmpNo", "EmpName", "CM003", , , , , 10, 20, , True)
    Call SetgiList(gilStrtCode, "CodeNo", "Description", "CD017", , , , , , , , True)
    SetgiList gilServiceType, "CodeNo", "Description", "CD046", , , , , 1, 12
    SetgiList gilCompCode, "CodeNo", "Description", "CD039", , , , , 3, 12, GetCompCodeFilter(gilCompCode)
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "subGil")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
        Set cnn = Nothing
End Sub

Public Property Get uForm() As String
    On Error GoTo ChkErr
        uForm = strForm
    Exit Property
ChkErr:
    ErrSub Me.Name, "Get uForm"
End Property

Public Property Let uForm(ByVal vForm As String)
    On Error GoTo ChkErr
        strForm = vForm
    Exit Property
ChkErr:
    ErrSub Me.Name, "Let uForm"
End Property

