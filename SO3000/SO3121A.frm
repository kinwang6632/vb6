VERSION 5.00
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSO3121A 
   BorderStyle     =   1  '單線固定
   Caption         =   "應收明細 [SO3121A]"
   ClientHeight    =   4275
   ClientLeft      =   2205
   ClientTop       =   1755
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
   Icon            =   "SO3121A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '螢幕中央
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   990
      Left            =   870
      ScaleHeight     =   930
      ScaleWidth      =   5055
      TabIndex        =   14
      Top             =   1470
      Visible         =   0   'False
      Width           =   5115
      Begin MSComctlLib.ProgressBar pgbBar 
         Height          =   375
         Left            =   165
         TabIndex        =   15
         Top             =   390
         Width           =   4845
         _ExtentX        =   8546
         _ExtentY        =   661
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Label lblPmt 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "程序處理中，請稍後 ..."
         Height          =   180
         Left            =   210
         TabIndex        =   16
         Top             =   135
         Width           =   1800
      End
   End
   Begin VB.ComboBox cboOrder 
      Height          =   315
      ItemData        =   "SO3121A.frx":0442
      Left            =   990
      List            =   "SO3121A.frx":044F
      Style           =   2  '單純下拉式
      TabIndex        =   13
      Top             =   3000
      Width           =   2055
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "F5:列印"
      Default         =   -1  'True
      Height          =   435
      Left            =   150
      MaskColor       =   &H8000000F&
      TabIndex        =   5
      Top             =   3750
      Width           =   1245
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      Height          =   435
      Left            =   5460
      TabIndex        =   4
      Top             =   3750
      Width           =   1275
   End
   Begin VB.CommandButton cmdPrevRpt 
      Caption         =   "上次統計結果"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1650
      TabIndex        =   3
      Top             =   3750
      Width           =   1395
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2475
      Left            =   0
      TabIndex        =   0
      Top             =   420
      Width           =   6825
      Begin VB.OptionButton optAll 
         Caption         =   "全部查詢"
         Height          =   345
         Left            =   300
         TabIndex        =   2
         Top             =   2040
         Width           =   1275
      End
      Begin VB.OptionButton optAbnormal 
         Caption         =   "異常查詢"
         Height          =   225
         Left            =   300
         TabIndex        =   1
         Top             =   300
         Value           =   -1  'True
         Width           =   1245
      End
      Begin VB.Label lbl5 
         Appearance      =   0  '平面
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '透明
         Caption         =   "5.收費人員 = {空白} 或 {未設}"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   0
         Left            =   2730
         TabIndex        =   17
         Top             =   1590
         Width           =   2535
      End
      Begin VB.Label lbl1 
         Appearance      =   0  '平面
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  '透明
         Caption         =   "異常條件 :"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   0
         Left            =   1590
         TabIndex        =   11
         Top             =   300
         Width           =   975
      End
      Begin VB.Label lbl2 
         Appearance      =   0  '平面
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '透明
         Caption         =   "1.應收金額 = 0 "
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   1
         Left            =   2730
         TabIndex        =   10
         Top             =   300
         Width           =   1245
      End
      Begin VB.Label lbl5 
         Appearance      =   0  '平面
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '透明
         Caption         =   "4.應收日期 = {空白}"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   2
         Left            =   2730
         TabIndex        =   9
         Top             =   1260
         Width           =   1680
      End
      Begin VB.Label lbl3 
         Appearance      =   0  '平面
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '透明
         Caption         =   "2.裝機數量 = 0 "
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   3
         Left            =   2730
         TabIndex        =   8
         Top             =   600
         Width           =   1245
      End
      Begin VB.Label lbl4 
         Appearance      =   0  '平面
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '透明
         Caption         =   "3.收費期間 = 0"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   4
         Left            =   2730
         TabIndex        =   7
         Top             =   930
         Width           =   1200
      End
      Begin VB.Label Label1 
         Alignment       =   2  '置中對齊
         Appearance      =   0  '平面
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '透明
         Caption         =   "以上條件有一項成立者為異常"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   195
         Index           =   5
         Left            =   2670
         TabIndex        =   6
         Top             =   1920
         Width           =   2775
      End
   End
   Begin prjGiList.GiList gilCompCode 
      Height          =   315
      Left            =   720
      TabIndex        =   18
      Top             =   60
      Width           =   2535
      _ExtentX        =   4471
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
      FldWidth1       =   600
      FldWidth2       =   1600
      F2Corresponding =   -1  'True
   End
   Begin prjGiList.GiList gilServiceType 
      Height          =   315
      Left            =   4260
      TabIndex        =   19
      Top             =   75
      Width           =   2535
      _ExtentX        =   4471
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
      FldWidth1       =   600
      FldWidth2       =   1600
      F2Corresponding =   -1  'True
   End
   Begin VB.Label lblServiceType 
      AutoSize        =   -1  'True
      Caption         =   "服務類別"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   3420
      TabIndex        =   21
      Top             =   150
      Width           =   780
   End
   Begin VB.Label lblCompCode 
      AutoSize        =   -1  'True
      Caption         =   "公司別"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   90
      TabIndex        =   20
      Top             =   120
      Width           =   585
   End
   Begin VB.Label lblOrder 
      AutoSize        =   -1  'True
      Caption         =   "排序條件"
      Height          =   195
      Left            =   60
      TabIndex        =   12
      Top             =   3090
      Width           =   780
   End
End
Attribute VB_Name = "frmSO3121A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private cnn As New ADODB.Connection
Private strSQL As String
Private strChoose As String
Private strChooseString As String
Private strForm As String
Private strOrder As String

Private Sub cmdExit_Click()
  On Error GoTo chkErr
   Unload Me
  Exit Sub
chkErr:
    Call ErrSub(Me.Name, "cmdExit_Click")
End Sub

Private Sub cmdPrevRpt_Click()
    On Error GoTo chkErr
        Screen.MousePointer = vbHourglass
        Call PreviousRpt(GetPrinterName(5), "SO3121A.Rpt", Me.Caption)
        Screen.MousePointer = vbDefault
    Exit Sub
chkErr:
    Call ErrSub(Me.Name, "cmdPrevRpt_Click")
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo chkErr
    Dim rs As New ADODB.Recordset
        Screen.MousePointer = vbHourglass
        If Not ChkDTok Then Exit Sub
        If Not IsDataOK Then Exit Sub
        Me.Enabled = False
        ReadyGoPrint
        subChoose
        If Not GetRS(rs, "SELECT COUNT(*) as intCount FROM " & GetOwner & IIf(strForm = "SO3121", " SO032 ", "SO033 ") & " A " & IIf(strChoose = "", "", " Where ") & strChoose) Then Me.Enabled = True: Exit Sub
        If rs("intCount") = 0 Then
            MsgNoRcd
        Else
            Picture1.Visible = True
            If Not subInsertMDB Then Me.Enabled = True: Picture1.Visible = False: Exit Sub
            Picture1.Visible = False
            strSQL = "Select * From SO3121A"
            Call PrintRpt(GetPrinterName(5), "SO3121A.rpt", "SO3121", Me.Caption, strSQL, strChooseString, strOrder, True, "TMP0000.MDB", , , GiPaperLandscape)
        End If
        On Error Resume Next
        Screen.MousePointer = vbDefault
        Me.Enabled = True
    Exit Sub
chkErr:
    Me.Enabled = True
    Call ErrSub(Me.Name, "cmdPrint_Click")
End Sub

Private Function IsDataOK() As Boolean
    On Error GoTo chkErr
        If Not MustExist(gilCompCode, 2, "公司別") Then Exit Function
        If Not MustExist(gilServiceType, 2, "服務類別") Then Exit Function
        IsDataOK = True
    Exit Function
chkErr:
    ErrSub Me.Name, "IsDataOk"
End Function

Private Sub subChoose()
    On Error GoTo chkErr
        strChoose = " A.ClctYM is Not Null "
        If gilCompCode.GetCodeNo <> "" Then Call subAnd2(strChoose, "A.CompCode =" & gilCompCode.GetCodeNo)
        If gilServiceType.GetCodeNo <> "" Then Call subAnd2(strChoose, "A.ServiceType ='" & gilServiceType.GetCodeNo & "'")
        If optAbnormal.Value Then
            strChoose = strChoose & "  and (A.ShouldAmt=0 or A.RealPeriod=0 or " & _
            "A.RealStartDate is Null or A.RealStopDate is Null or A.ShouldDate is Null Or " & _
            "A.ClctEn Is Null Or  A.ClctEn ='???') "
            strChooseString = "異常資料"
        Else
            strChooseString = "全部資料"
        End If
        Select Case cboOrder.ListIndex
            Case 0
                strChoose = strChoose & " Order By A.CustId "
                strOrder = "客戶編號"
            Case 1
                strChoose = strChoose & " Order By A.ClctEn,A.CustId "
                strOrder = "收費人員"
            Case 2
                strChoose = strChoose & " Order By A.StrtCode,A.CustId "
                strOrder = "街道編號"
            Case Else
                strChoose = strChoose & " Order By A.CustId "
                strOrder = "客戶編號"
        End Select
    Exit Sub
chkErr:
    ErrSub Me.Name, "subChoose"
End Sub

Private Sub subGil()
    On Error Resume Next
        SetgiList gilServiceType, "CodeNo", "Description", "CD046", , , , , 1, 12
        SetgiList gilCompCode, "CodeNo", "Description", "CD039", , , , , 3, 12, GetCompCodeFilter(gilCompCode)

End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
  On Error GoTo chkErr
    Set cnn = GetTmpMdbCn
    subGil
    gilCompCode.SetCodeNo garyGi(9)
    gilCompCode.Query_Description
    cboOrder.ListIndex = 0
  Exit Sub
chkErr:
    Call ErrSub(Me.Name, "cmdPrint_Click")
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error GoTo chkErr
    Call FunctionKey(KeyCode, Shift)
  Exit Sub
chkErr:
    Call ErrSub(Me.Name, "FunctionKey")
End Sub

Public Sub FunctionKey(KeyCode As Integer, Shift As Integer)
  On Error GoTo chkErr
    Select Case KeyCode
           '----------------------------------------------------
           Case vbKeyF5 '   F5: 列印 相當於按下cmdPrint
                        If cmdPrint.Enabled = True Then cmdPrint.Value = True
           '----------------------------------------------------
           Case vbKeyEscape
                        If cmdExit.Enabled = True Then cmdExit.Value = True
    End Select
  Exit Sub
chkErr:
    Call ErrSub(Me.Name, "FunctionKey")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
        Set cnn = Nothing
End Sub

Public Property Get uForm() As String
    On Error GoTo chkErr
        uForm = strForm
    Exit Property
chkErr:
    ErrSub Me.Name, "Get uForm"
End Property

Public Property Let uForm(ByVal vForm As String)
    On Error GoTo chkErr
        strForm = vForm
    Exit Property
chkErr:
    ErrSub Me.Name, "Let uForm"
End Property

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
        Call ReleaseCOM(Me)

End Sub

Private Function subInsertMDB() As Boolean
    On Error GoTo chkErr
    Dim strSQL As String
    Dim strMduName As String
    Dim strClctAreaName As String
    Dim strOldSQL As String
    Dim rsTmp As New ADODB.Recordset
        '#4030 增加期數 By Kin 2008/07/31
        '#4258 要增加BpCode與BpName By Kin 2008/12/22
        strOldSQL = "Select A.CustId,A.BillNO,A.CitemName,A.ShouldAmt,A.RealStartDate,A.RealStopDate, " & _
            "A.ShouldDate,A.ClctName,A.AddrNo,A.MduId,B.Tel1,B.CustName,C.Address,A.RealPeriod,A.BpCode,A.BpName "
        strSQL = strOldSQL & ",D.NAME MduName,E.Description ClctAreaName From " & GetOwner & "SO032 A," & GetOwner & "SO001 B," & GetOwner & "So014 C," & GetOwner & "SO017 D," & GetOwner & "CD040 E Where 1=0"
        '新增一個TMP 檔
        If Not GetRS(rsTmp, strSQL) Then Exit Function
        If Not CreateMDBTable(rsTmp, "SO3121A", cnn) Then Exit Function
        
        strSQL = strOldSQL & ",A.ClctAreaCode From " & GetOwner & "So032 A," & GetOwner & "So001 B," & GetOwner & "So014 C " & _
            " Where A.CustId = B.CustId And A.CompCode = B.CompCode And A.AddrNo= C.AddrNo And A.CompCode = C.CompCode " & IIf(strChoose = "", "", " And ") & strChoose
        
        cnn.BeginTrans
        If Not GetRS(rsTmp, strSQL) Then Exit Function
        SendSQL strSQL, True
        Do While Not rsTmp.EOF
            If Len(rsTmp("MduId") & "") > 0 Then
                strMduName = GetRsValue("Select Name From " & GetOwner & "So017 Where MduId = '" & rsTmp("MduId") & "'") & ""
            Else
                strMduName = ""
            End If
            If Len(rsTmp("ClctAreaCode") & "") > 0 Then
                strClctAreaName = GetRsValue("Select Description From " & GetOwner & "CD040 Where CodeNo = '" & rsTmp("ClctAreaCode") & "'") & ""
            Else
                strClctAreaName = ""
            End If
            '#4030 增加期數 By Kin 2008/07/31
            '#4258 測試不OK,BpCode,BpName沒有加進去 By Kin 2009/02/02
            cnn.Execute "Insert Into So3121A (CustId,BillNO,CitemName,ShouldAmt,RealStartDate,RealStopDate," & _
                    "ShouldDate,ClctName,AddrNo,MduId,Tel1,CustName,Address,MduName,ClctAreaName,RealPeriod,BpCode,BpName) Values (" & _
                    GetNullString(rsTmp(0), giLongV) & "," & GetNullString(rsTmp(1)) & "," & _
                    GetNullString(rsTmp(2)) & "," & GetNullString(rsTmp(3)) & "," & _
                    GetNullString(rsTmp(4), giDateV, giAccessDb) & "," & GetNullString(rsTmp(5), giDateV, giAccessDb) & "," & _
                    GetNullString(rsTmp(6), giDateV, giAccessDb) & "," & GetNullString(rsTmp(7)) & "," & _
                    GetNullString(rsTmp(8)) & "," & GetNullString(rsTmp(9)) & "," & _
                    GetNullString(rsTmp(10)) & "," & GetNullString(rsTmp(11)) & "," & _
                    GetNullString(rsTmp(12)) & "," & GetNullString(strMduName) & "," & _
                    GetNullString(strClctAreaName) & "," & GetNullString(rsTmp(13)) & "," & _
                    GetNullString(rsTmp("BpCode")) & "," & GetNullString(rsTmp("BpName")) & ")"
            If (rsTmp.AbsolutePosition / rsTmp.RecordCount * 100) Mod 5 = 0 Then
                pgbBar.Value = rsTmp.AbsolutePosition / rsTmp.RecordCount * 100
                DoEvents
            End If
            rsTmp.MoveNext
        Loop
        cnn.CommitTrans
        subInsertMDB = True
    Exit Function
chkErr:
    cnn.RollbackTrans
    Picture1.Visible = False
    ErrSub Me.Name, "subInsertMDB"
End Function

Private Sub gilCompCode_Change()
    On Error Resume Next
        If Not ChgComp(gilCompCode, "SO3120", "SO3121") Then Exit Sub
        Call subGil
        Call SelectServType(gilCompCode.GetCodeNo, gilServiceType)
        gilServiceType.SetCodeNo ""
        gilServiceType.Query_Description
        gilServiceType.ListIndex = 1
        
End Sub
