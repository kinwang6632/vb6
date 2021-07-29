VERSION 5.00
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmSO3124A 
   BorderStyle     =   1  '單線固定
   Caption         =   "應收期數及金額分析 [SO3124A]"
   ClientHeight    =   4110
   ClientLeft      =   2550
   ClientTop       =   2730
   ClientWidth     =   6855
   BeginProperty Font 
      Name            =   "新細明體"
      Size            =   9.75
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SO3124A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton cmdPrevRpt 
      Caption         =   "上次統計結果"
      Height          =   525
      Left            =   2550
      TabIndex        =   17
      Top             =   3420
      Width           =   1545
   End
   Begin MSMask.MaskEdBox mskPeriod1 
      Height          =   315
      Left            =   1320
      TabIndex        =   15
      Top             =   1740
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   2
      Mask            =   "99"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "F5:列印"
      Height          =   525
      Left            =   240
      TabIndex        =   7
      Top             =   3420
      Width           =   1275
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      Height          =   525
      Left            =   5340
      TabIndex        =   8
      Top             =   3420
      Width           =   1305
   End
   Begin VB.Frame fraDate 
      Caption         =   "期間天數"
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   3900
      TabIndex        =   14
      Top             =   2190
      Width           =   2775
      Begin VB.OptionButton optRealDate 
         Caption         =   "實際日"
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1440
         TabIndex        =   6
         Top             =   300
         Width           =   1065
      End
      Begin VB.OptionButton optStandardDate 
         Caption         =   "標準日"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   210
         TabIndex        =   5
         Top             =   300
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.Frame fraStatus 
      Caption         =   "客戶類別"
      Height          =   615
      Left            =   60
      TabIndex        =   13
      Top             =   2190
      Width           =   3735
      Begin VB.OptionButton optAll 
         Caption         =   "全部"
         Height          =   195
         Left            =   2640
         TabIndex        =   4
         Top             =   270
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton optMdu 
         Caption         =   "大樓戶"
         Height          =   195
         Left            =   1380
         TabIndex        =   3
         Top             =   270
         Width           =   1245
      End
      Begin VB.OptionButton optNormal 
         Caption         =   "一般戶"
         Height          =   195
         Left            =   210
         TabIndex        =   2
         Top             =   270
         Width           =   975
      End
   End
   Begin prjGiList.GiList gilClctEn 
      Height          =   315
      Left            =   1290
      TabIndex        =   1
      Top             =   1290
      Width           =   3000
      _ExtentX        =   5292
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
      FldWidth1       =   800
      F5Corresponding =   -1  'True
   End
   Begin prjGiList.GiList gilAreaCode 
      Height          =   315
      Left            =   1290
      TabIndex        =   0
      Top             =   860
      Width           =   3000
      _ExtentX        =   5292
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
      FldWidth1       =   800
      F5Corresponding =   -1  'True
   End
   Begin MSMask.MaskEdBox mskPeriod2 
      Height          =   315
      Left            =   2220
      TabIndex        =   16
      Top             =   1740
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "99"
      PromptChar      =   "_"
   End
   Begin prjGiList.GiList gilCompCode 
      Height          =   315
      Left            =   1290
      TabIndex        =   18
      Top             =   0
      Width           =   3000
      _ExtentX        =   5292
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
      FldWidth1       =   800
      F2Corresponding =   -1  'True
   End
   Begin prjGiList.GiList gilServiceType 
      Height          =   315
      Left            =   1290
      TabIndex        =   19
      Top             =   435
      Width           =   3000
      _ExtentX        =   5292
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
      FldWidth1       =   800
      F2Corresponding =   -1  'True
   End
   Begin VB.Label lblCompCode 
      AutoSize        =   -1  'True
      Caption         =   "公司別"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   210
      TabIndex        =   21
      Top             =   90
      Width           =   585
   End
   Begin VB.Label lblServiceType 
      AutoSize        =   -1  'True
      Caption         =   "服務類別"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   210
      TabIndex        =   20
      Top             =   510
      Width           =   780
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "~"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1920
      TabIndex        =   12
      Top             =   1770
      Width           =   120
   End
   Begin VB.Label lblPeriod 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "期        數 :"
      Height          =   195
      Left            =   210
      TabIndex        =   11
      Top             =   1800
      Width           =   840
   End
   Begin VB.Label lblClctEn 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "收費人員"
      Height          =   195
      Left            =   210
      TabIndex        =   10
      Top             =   1350
      Width           =   780
   End
   Begin VB.Label lblAreaCode 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "行  政  區 :"
      Height          =   195
      Left            =   210
      TabIndex        =   9
      Top             =   930
      Width           =   855
   End
End
Attribute VB_Name = "frmSO3124A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private cnn As New ADODB.Connection
Private strChoose As String
Private strChooseString As String
Private strForm As String
Private strFormula As String

Private Sub cmdPrevRpt_Click()
    On Error GoTo ChkErr
        Screen.MousePointer = vbHourglass
        Call PreviousRpt(GetPrinterName(5), "SO3124A.Rpt", "本期應收期數及金額分析 [SO3124A]")
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

Private Sub cmdExit_Click()
  On Error GoTo ChkErr
   Unload Me
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "cmdExit_Click")
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo ChkErr
    Dim lngAffected As Long
    
        If Not ChkDTok Then Exit Sub
        cmdExit.SetFocus
        ReadyGoPrint
        If Not subChoose Then Exit Sub
        Screen.MousePointer = vbHourglass
        If Not subInsertMDB(lngAffected) Then Exit Sub
        If lngAffected = 0 Then
            MsgNoRcd
        Else
            strSQL = "SELECT * FROM SO3124A"
            Call PrintRpt(GetPrinterName(5), "SO3124A.rpt", "SO3124", "本期應收區域及道路分析 [SO3124A]", strSQL, strChooseString, " 收費人員", True, "Tmp0000.MDB", , strFormula)
        End If
        Screen.MousePointer = vbDefault
    Exit Sub
ChkErr:
    Screen.MousePointer = vbDefault
    Call ErrSub(Me.Name, "cmdPrint_Click")
End Sub

Private Function subInsertMDB(ByRef lngAffected As Long) As Boolean
    On Error GoTo ChkErr
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
        cnn.BeginTrans
        strSQL = "Select ClctYM,ClctEn,ClctName,RealPeriod,ShouldAmt,Decode(MduId,Null,'一般戶','大樓戶') CustType, " & _
              "Sum(ShouldAmt) TotalShould,Count(*) intCount From " & GetOwner & "SO032 " & strChoose & " Group By " & _
              "ClctYM,ClctEn,ClctName,RealPeriod,ShouldAmt,Decode(MduId,Null,'一般戶','大樓戶') "
        SendSQL strSQL, True
        If Not GetRS(rsTmp, strSQL) Then Exit Function
        If Not CreateMDBTable(rsTmp, "SO3124A", cnn) Then Exit Function
        Do While Not rsTmp.EOF
            cnn.Execute GetTmpMDBExecuteStr(rsTmp, "SO3124A")
            rsTmp.MoveNext
            DoEvents
        Loop
        lngAffected = rsTmp.RecordCount
        cnn.CommitTrans
        Call Close3Recordset(rsTmp)
        
        subInsertMDB = True
    Exit Function
ChkErr:
    cnn.RollbackTrans
    ErrSub Me.Name, "subInsertMDB"
End Function

Private Function subChoose() As Boolean
    '串 SelectCondition
    On Error GoTo ChkErr
    Dim strCustStatus As String
    Dim strDateType As String
        strChoose = " ClctYM is Not Null "
        strChooseString = ""
        If mskPeriod1 <> "" Then Call subAnd("OldPeriod>=" & Val(mskPeriod1))
        If mskPeriod2 <> "" Then Call subAnd("OldPeriod<=" & Val(mskPeriod2))
        If gilAreaCode.GetCodeNo <> "" Then Call subAnd("AreaCode='" & gilAreaCode.GetCodeNo & "'")
        If gilClctEn.GetCodeNo <> "" Then Call subAnd("ClctEn='" & gilClctEn.GetCodeNo & "'")
        If gilCompCode.GetCodeNo <> "" Then Call subAnd2(strChoose, "CompCode =" & gilCompCode.GetCodeNo)
        If gilServiceType.GetCodeNo <> "" Then Call subAnd2(strChoose, "ServiceType ='" & gilServiceType.GetCodeNo & "'")
        If optNormal.Value Then Call subAnd("MduId Is Null")
        If optMdu.Value Then Call subAnd("MduId Is Not Null")
        strChoose = IIf(strChoose = "", "", " WHERE " & strChoose)
        If optNormal.Value Then
           strCustStatus = "一般戶"
        Else
           If optMdu.Value Then
              strCustStatus = "大樓戶"
           Else
              strCustStatus = "全  部"
           End If
        End If
        If optStandardDate.Value Then
           strDateType = "標準日"
           strFormula = "DayType=1"
        Else
           strDateType = "實際日"
           strFormula = "DayType=2"
        End If
        strChooseString = "行政區　: " & subSpace(gilAreaCode.GetDescription) & ";" & _
                          "收費人員: " & subSpace(gilClctEn.GetDescription) & ";" & _
                          "期數　　: " & subSpace(mskPeriod1) & "~" & subSpace(mskPeriod2) & ";" & _
                          "客戶類別: " & strCustStatus & ";" & _
                          "期間天數: " & strDateType
        subChoose = True
    Exit Function
ChkErr:
    Call ErrSub(Me.Name, "subChoose")
End Function

Private Sub subAnd(strCh As String)
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

Private Sub subGil()
  On Error GoTo ChkErr
    SetgiList gilClctEn, "EmpNo", "EmpName", "CM003", , , , , 10, 20, , True
    SetgiList gilAreaCode, "CodeNo", "Description", "CD001", , , , , 3, 12, , True
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

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
        Call ReleaseCOM(Me)
End Sub

Private Sub gilCompCode_Change()
    On Error Resume Next
        If Not ChgComp(gilCompCode, "SO3120", "SO3124") Then Exit Sub
        Call subGil
        Call SelectServType(gilCompCode.GetCodeNo, gilServiceType)
        gilServiceType.SetCodeNo ""
        gilServiceType.Query_Description
        gilServiceType.ListIndex = 1
        Call GiListFilter(gilAreaCode, , gilCompCode.GetCodeNo)
        Call GiListFilter(gilClctEn, , gilCompCode.GetCodeNo)
End Sub

