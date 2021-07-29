VERSION 5.00
Object = "{A0C51418-555F-11D3-BC10-0050BA01A3A9}#9.8#0"; "GiMulti.ocx"
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{074771B7-6ADE-11D5-B178-00485456B028}#1.17#0"; "csMulti.ocx"
Begin VB.Form frmSO52B0A 
   BorderStyle     =   1  '單線固定
   Caption         =   "收費金額資料統計[ SO52B0A ]"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7680
   Icon            =   "SO52B0A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   7680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '螢幕中央
   Begin VB.Frame Frame4 
      Caption         =   "網路編號"
      Height          =   735
      Left            =   0
      TabIndex        =   26
      Top             =   4530
      Width           =   7665
      Begin VB.TextBox mskCircuitNo 
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   90
         TabIndex        =   14
         Top             =   240
         Width           =   4605
      End
      Begin VB.OptionButton optAll 
         Caption         =   "全部"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   6660
         TabIndex        =   16
         Top             =   270
         Value           =   -1  'True
         Width           =   675
      End
      Begin VB.OptionButton optCircuitNo 
         Caption         =   "只印空白網路編號"
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   4830
         TabIndex        =   15
         Top             =   240
         Width           =   1845
      End
   End
   Begin CS_Multi.CSmulti gimStrtCode 
      Height          =   405
      Left            =   0
      TabIndex        =   12
      Top             =   3720
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   714
      ButtonCaption   =   "街  道  範  圍"
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "期數算法"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   705
      Left            =   3960
      TabIndex        =   23
      Top             =   120
      Width           =   3525
      Begin VB.OptionButton optPeriod 
         BackColor       =   &H00E0E0E0&
         Caption         =   "實際收費期間期數"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   1440
         TabIndex        =   5
         Top             =   300
         Value           =   -1  'True
         Width           =   1875
      End
      Begin VB.OptionButton optOldPeriod 
         BackColor       =   &H00E0E0E0&
         Caption         =   "收費單期數"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   300
         Width           =   1335
      End
   End
   Begin VB.Frame fraReportType 
      BackColor       =   &H00E0E0E0&
      Caption         =   "報表型式"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   705
      Left            =   3960
      TabIndex        =   20
      Top             =   960
      Width           =   3525
      Begin VB.OptionButton optSummaric1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "試算表"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   195
         Left            =   150
         TabIndex        =   6
         Top             =   300
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optDetail 
         BackColor       =   &H00E0E0E0&
         Caption         =   "條列表"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   195
         Left            =   1440
         TabIndex        =   7
         Top             =   300
         Width           =   915
      End
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "F5:列印"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   120
      TabIndex        =   17
      Top             =   5580
      Width           =   1245
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   6270
      TabIndex        =   19
      Top             =   5580
      Width           =   1245
   End
   Begin VB.CommandButton cmdPrevRpt 
      Caption         =   "上次統計結果"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   1800
      TabIndex        =   18
      Top             =   5580
      Width           =   1395
   End
   Begin Gi_Multi.GiMulti gimBillType 
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   2550
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   661
      ButtonCaption   =   "單  據  類  別"
      DataType        =   2
   End
   Begin Gi_Multi.GiMulti gimCMCode 
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   2160
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   661
      ButtonCaption   =   "收  費  方  式"
   End
   Begin Gi_Date.GiDate gdaRealDate2 
      Height          =   345
      Left            =   2670
      TabIndex        =   1
      Top             =   120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
      ForeColor       =   16711680
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
   Begin Gi_Date.GiDate gdaRealDate1 
      Height          =   345
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
      ForeColor       =   16711680
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
   Begin Gi_Multi.GiMulti gimAreaCode 
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   3330
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   661
      ButtonCaption   =   "行     政     區"
   End
   Begin Gi_Multi.GiMulti gimServCode 
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   2940
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   661
      ButtonCaption   =   "服     務     區"
   End
   Begin prjGiList.GiList gilServiceType 
      Height          =   315
      Left            =   870
      TabIndex        =   3
      Top             =   1050
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
      Left            =   870
      TabIndex        =   2
      Top             =   630
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
   Begin CS_Multi.CSmulti gimMduId 
      Height          =   375
      Left            =   0
      TabIndex        =   13
      Top             =   4140
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   661
      ButtonCaption   =   "大  樓  編  號"
   End
   Begin CS_Multi.CSmulti gimCitemCode 
      Height          =   375
      Left            =   0
      TabIndex        =   27
      Top             =   1770
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   661
      ButtonCaption   =   "收  費  項  目"
   End
   Begin VB.Label lblCompCode 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BackStyle       =   0  '透明
      Caption         =   "公  司  別"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   90
      TabIndex        =   25
      Top             =   690
      Width           =   765
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BackStyle       =   0  '透明
      Caption         =   "服務類別"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   90
      TabIndex        =   24
      Top             =   1110
      Width           =   780
   End
   Begin VB.Label lblShouldDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "實收日期範圍"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   90
      TabIndex        =   22
      Top             =   210
      Width           =   1170
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "~"
      ForeColor       =   &H00FF0000&
      Height          =   165
      Left            =   2520
      TabIndex        =   21
      Top             =   240
      Width           =   105
   End
End
Attribute VB_Name = "frmSO52B0A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Table=SO033 Or SO034 A
Option Explicit
Dim strPeriod As String
Dim strAmt As String
Dim TableFieldAmt As String
Dim TableFieldCount As String
Dim strFieldAmt  As String
Dim strFieldCount As String

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
    Select Case True
           Case optSummaric1.Value '試算表
                Call PreviousRpt(GetPrinterName(5), RptName("SO52B0", "1"), Me.Caption)
           Case optDetail.Value '條列表
                Call PreviousRpt(GetPrinterName(5), RptName("SO52B0", "2"), Me.Caption)
    End Select
  Screen.MousePointer = vbDefault
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdPrevRpt_Click")
End Sub

Private Sub cmdPrint_Click()
  On Error GoTo ChkErr
    Dim rsTmp As New ADODB.Recordset
       
      If Not ChkDTok Then Exit Sub
      If Not IsDataOk Then Exit Sub
        Screen.MousePointer = vbHourglass
          cmdExit.SetFocus
          Me.Enabled = False
            ReadyGoPrint
            Call subChoose
            subCreateView
            CreateTable
            Set rsTmp = cnn.Execute("SELECT Count(*) as intCount From SO52B0A2")
            If rsTmp("intCount") = 0 Then
                MsgNoRcd
                SendSQL , , True
            Else
              strsql = "Select * From SP52B0A2"
              Select Case True
                Case optSummaric1.Value '試算表
                     
                     Call PrintRpt(GetPrinterName(5), RptName("SO52B0", "1"), , Me.Caption, strsql, strChooseString, , True, "Tmp0000.MDB", , , GiPaperLandscape)
                Case optDetail.Value '條列表
                     
                     Call PrintRpt(GetPrinterName(5), RptName("SO52B0", "2"), , Me.Caption, strsql, strChooseString, , True, "Tmp0000.MDB", , , GiPaperPortrait)
    
              End Select
    
            End If
          Me.Enabled = True
        Screen.MousePointer = vbDefault
        CloseRecordset rsTmp
    Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdPrint_Click")
End Sub

Private Function IsDataOk() As Boolean
  On Error GoTo ChkErr
    If gilCompCode.GetCodeNo = "" Then gilCompCode.SetFocus: strErrFile = "公司別": GoTo 66
    IsDataOk = True
  Exit Function
66:
    MsgMustBe (strErrFile)
  Exit Function
ChkErr:
  Call ErrSub(Me.Name, "IsDataOk")
End Function

Private Sub subChoose()
  On Error GoTo ChkErr
  Dim strRptType As String
  Dim strPeriodtype As String
  Dim strGroupString As String
    strChoose = ""
    strChooseString = ""

    If gdaRealDate1.GetValue <> "" Then Call subAnd("A.RealDate >= To_Date('" & gdaRealDate1.GetValue & "','YYYYMMDD')")
    If gdaRealDate2.GetValue <> "" Then Call subAnd("A.RealDate < To_Date('" & gdaRealDate2.GetValue & "','YYYYMMDD')+1")
    'If gimCompCode.GetQryStr <> "" Then Call subAnd("A.CompCode" & gimCompCode.GetQryStr)
    If gimCitemCode.GetQryStr <> "" Then Call subAnd("A.CitemCode " & gimCitemCode.GetQryStr)
    If gimCMCode.GetQryStr <> "" Then Call subAnd("A.CMCode " & gimCMCode.GetQryStr)
    If gimBillType.GetQryStr <> "" Then Call subAnd("SubStr(A.BillNo,7,1) " & gimBillType.GetQryStr)
    If gimServCode.GetQryStr <> "" Then Call subAnd("A.ServCode " & gimServCode.GetQryStr)
    If gimAreaCode.GetQryStr <> "" Then Call subAnd("A.AreaCode " & gimAreaCode.GetQryStr)
    If gimStrtCode.GetQryStr <> "" Then Call subAnd("A.StrtCode " & gimStrtCode.GetQryStr)
    If gimMduId.GetQryStr <> "" Then
        Call subAnd("A.MduId " & gimMduId.GetQryStr)
    Else
        If gimMduId.GetDispStr <> "" Then subAnd ("A.MduId is Not Null")
    End If

  'GiList
    If gilCompCode.GetCodeNo <> "" Then Call subAnd("A.CompCode=" & gilCompCode.GetCodeNo)
    If gilServiceType.GetCodeNo <> "" Then Call subAnd("A.ServiceType='" & gilServiceType.GetCodeNo & "'")

    If strChoose <> "" Then
        strChoose = " Where RealDate IS Not Null And A.cancelFlag=0 And A.UCcode is null And " & strChoose
    Else
        strChoose = " Where RealDate IS Not Null And A.cancelFlag=0 And A.UCcode is null "
    End If

    Select Case True
           Case optOldPeriod.Value
                strPeriod = "OldPeriod"
                strAmt = "OldAmt"
                strPeriodtype = "收費單期數"
           Case optPeriod.Value
                strPeriod = "RealPeriod"
                strAmt = "ShouldAmt"
                strPeriodtype = "實際收費期間期數"
    End Select
     
    Select Case True
           Case optSummaric1
                strRptType = "試算表"
           Case optDetail
                strRptType = "條列表"
    End Select

    strChooseString = "收費日期: " & subSpace(gdaRealDate1.GetValue(True)) & "~" & subSpace(gdaRealDate2.GetValue(True)) & ";" & _
                      "公司別　: " & subSpace(gilCompCode.GetDescription) & ";" & _
                      "服務類別: " & subSpace(gilServiceType.GetDescription) & ";" & _
                      "報表型式: " & strRptType & ";" & _
                      "期數算法: " & strPeriodtype & ";" & _
                      "收費項目: " & subSpace(gimCitemCode.GetDispStr) & ";" & _
                      "收費方式: " & subSpace(gimCMCode.GetDispStr) & ";" & _
                      "單據編號: " & subSpace(gimBillType.GetDispStr) & ";" & _
                      "服務區　: " & subSpace(gimServCode.GetDispStr) & ";" & _
                      "行政區　: " & subSpace(gimAreaCode.GetDispStr) & ";" & _
                      "街道範圍: " & subSpace(gimStrtCode.GetDispStr) & ";" & _
                      "大樓編號: " & subSpace(gimMduId.GetDispStr)


  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "subChoose")
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error GoTo ChkErr
    Call FunctionKey(KeyCode, Shift, frmSO52B0A)
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "Form_KeyDown")
End Sub

Private Sub Form_Load()
  On Error GoTo ChkErr
    If Alfa2 Then Call GetGlobal
    Call subGim
    Call subGil
    'gilCompCode.ListIndex = 1
    gilCompCode.SetCodeNo garyGi(9)
    gilCompCode.Query_Description
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "Form_Load")
End Sub

Private Sub subGim()
  On Error GoTo ChkErr
    Call SetgiMulti(gimCitemCode, "CodeNo", "Description", "CD019", "收費項目代碼", "收費項目名稱")
    Call SetgiMulti(gimCMCode, "CodeNo", "Description", "CD031", "收費方式代碼", "收費方式名稱")
    Call SetgiMultiAddItem(gimBillType, "B,I,P,M,T", "收費單,裝機單,停拆移機單,維修單,臨時收費單")
    Call SetgiMulti(gimServCode, "CodeNo", "Description", "CD002", "服務區代碼", "服務區名稱")
    Call SetgiMulti(gimAreaCode, "CodeNo", "Description", "CD001", "行政區代碼", "行政區名稱")
    Call SetgiMulti(gimStrtCode, "CodeNo", "Description", "CD017", "街道代碼", "街道名稱")
    Call SetgiMulti(gimMduId, "MduId", "Name", "SO017", "大樓編號", "大樓名稱")
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "subGim")
End Sub

Private Sub subGil()
  On Error GoTo ChkErr
    'Call SetgiList(gilCompCode, "CodeNo", "Description", "CD039")
    Call SetgiList(gilCompCode, "CodeNo", "Description", "CD039", , , , , , , GetCompCodeFilter(gilCompCode))
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "subGil")
End Sub

Private Function subCreateView() As Boolean
  Dim strView As String
  On Error Resume Next
    gcnGi.Execute "Drop View " & strViewName
  On Error GoTo ChkErr
    strViewName = GetTmpViewName
    strView = "Create View " & strViewName & " as (SELECT A.BillNo,A.RealDate," & _
              "A.CitemCode,A.CitemName,A.RealPeriod,A.OldPeriod,A.ShouldAmt,A.OldAmt,A.RealAmt " & _
              "From SO033 A " & strChoose & " Union All " & _
              "SELECT A.BillNo,A.RealDate,A.CitemCode,A.CitemName,A.RealPeriod,A.OldPeriod," & _
              "A.ShouldAmt,A.OldAmt,A.RealAmt From SO034 A " & strChoose & ")"
    gcnGi.Execute strView
    SendSQL strView, True
  Exit Function
ChkErr:
  Call ErrSub(Me.Name, "subCreateView")
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  On Error Resume Next
  gcnGi.Execute "Drop View " & strViewName
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
  ReleaseCOM frmSO52B0A
End Sub

Private Sub gdaRealDate1_GotFocus()
  On Error Resume Next
  If gdaRealDate1.GetValue = "" Then gdaRealDate1.SetValue (RightDate)
End Sub

Private Sub gdaRealDate2_GotFocus()
  On Error Resume Next
    If gdaRealDate1.GetValue = "" Or gdaRealDate2.GetValue = "" Then gdaRealDate2.SetValue gdaRealDate1.GetValue
End Sub

Private Sub gdaRealDate2_Validate(Cancel As Boolean)
  On Error Resume Next
  Cancel = ChkDate2(gdaRealDate1, gdaRealDate2)
End Sub

Private Sub gilCompCode_Change()
  On Error GoTo ChkErr
    Call SelectServType(gilCompCode.GetCodeNo, gilServiceType)
    gilServiceType.ListIndex = 1

    GiMultiFilter gimServCode, , gilCompCode.GetCodeNo
    GiMultiFilter gimAreaCode, , gilCompCode.GetCodeNo
    GiMultiFilter gimStrtCode, , gilCompCode.GetCodeNo
    GiMultiFilter gimMduId, , gilCompCode.GetCodeNo
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "gilCompCode_Change")
End Sub

Private Sub gilCompCode_Validate(Cancel As Boolean)
  On Error GoTo ChkErr
    If gilCompCode.GetCodeNo = "" Then
        MsgMustBe ("公司別")
        Cancel = True
    End If
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "gilCompCode_Validate")
End Sub

Private Sub gilServiceType_Change()
  On Error GoTo ChkErr
    GiMultiFilter gimCitemCode, gilServiceType.GetCodeNo
    GiMultiFilter gimCMCode, gilServiceType.GetCodeNo
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "gilCompCode_Validate")
End Sub

Private Sub mskCircuitNo_Validate(Cancel As Boolean)
  If mskCircuitNo.Text = "" Then
      optCircuitNo.Enabled = True
      optAll.Enabled = True
  Else
      optCircuitNo.Enabled = False
      optAll.Enabled = False
  End If
End Sub

Private Sub CreateTable()
  On Error Resume Next
    cnn.Execute "Drop Table SO52B0A1"
    cnn.Execute "Drop Table So52B0A2"
  On Error GoTo ChkErr
    FieldName
    'Insert 1-31日金額,張數
    Select Case True
      Case optSummaric1.Value
           cnn.Execute "Create Table SO52B0A1 (CitemCode Long,CitemName Text(40), Period Long,Amt Long" & TableFieldAmt & TableFieldCount & ")"
           Call subInsertMDB1("Select Count (Distinct BillNo) as CountBillNo,CitemCode,CitemName," & strPeriod & " as Period," & strAmt & " as Amt,To_Char(RealDate,'DD') as RealDay,Sum(" & strAmt & ") as SAmt From " & strViewName & _
                              " Group by CitemCode,CitemName," & strPeriod & "," & strAmt & ",To_Char(RealDate,'DD')")
           SumMDB
      Case optDetail.Value
           cnn.Execute "Create Table SO52B0A2 (CitemCode Long,CitemName Text(40), Period Long,Amt Long,RealDate Date,SumAmt Long,SumCount Long)"
           Call subInsertMDB2("Select Count (Distinct BillNo) as CountBillNo,CitemCode,CitemName," & strPeriod & " as Period," & strAmt & " as Amt,RealDate,Sum(" & strAmt & ") as SAmt From " & strViewName & _
                              " Group by CitemCode,CitemName," & strPeriod & "," & strAmt & ",RealDate")
                   
    End Select
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "CreateTabel")
End Sub

Private Sub subInsertMDB1(strsql As String)
  On Error GoTo ChkErr
  Dim rsTmp As New ADODB.Recordset
    Set rsTmp = gcnGi.Execute(strsql)
    SendSQL strsql
    cnn.BeginTrans
    While Not rsTmp.EOF
      cnn.Execute "Insert Into SO52B0A1 (CitemCode,CitemName,Period,Amt,Amt" & Format(rsTmp("RealDay"), "00") & ",Count" & Format(rsTmp("RealDay"), "00") & _
                  ") Values(" & _
                  GetNullString(rsTmp("CitemCode")) & "," & _
                  GetNullString(rsTmp("CitemName")) & "," & _
                  GetNullString(rsTmp("Period")) & "," & _
                  GetNullString(rsTmp("Amt")) & "," & _
                  GetNullString(rsTmp("SAmt")) & "," & _
                  GetNullString(rsTmp("CountBillno")) & ")"
     rsTmp.MoveNext
    Wend
    cnn.CommitTrans
    CloseRecordset rsTmp
  Exit Sub
ChkErr:
  cnn.RollbackTrans
  Call ErrSub(Me.Name, "subInsertMDB1")
End Sub

Private Sub subInsertMDB2(strsql As String)
  On Error GoTo ChkErr
    Dim rsTmp As New ADODB.Recordset
    Set rsTmp = gcnGi.Execute(strsql)
    SendSQL strsql
    cnn.BeginTrans
      While Not rsTmp.EOF
      
        cnn.Execute "Insert Into SO52B0A2 (" & _
                    "CitemCode,CitemName,Period,Amt,RealDate,SumAmt,SumCount) Values (" & _
                    GetNullString(rsTmp("Citemcode")) & "," & _
                    GetNullString(rsTmp("CitemName")) & "," & _
                    GetNullString(rsTmp("Period")) & "," & _
                    GetNullString(rsTmp("Amt")) & "," & _
                    GetNullString(rsTmp("RealDate"), giDateV, giAccessDb) & "," & _
                    GetNullString(rsTmp("SAmt")) & "," & _
                    GetNullString(rsTmp("CountBillNo")) & ")"
        rsTmp.MoveNext
      Wend
    cnn.CommitTrans
  Exit Sub
ChkErr:
  cnn.RollbackTrans
  Call ErrSub(Me.Name, "subInsertMDB2")
End Sub

Private Sub subUpdateMDB(strsql As String)
  On Error GoTo ChkErr
    Dim rsTmp As New ADODB.Recordset
    Set rsTmp = gcnGi.Execute(strsql)
    cnn.BeginTrans
      While Not rsTmp.EOF
        cnn.Execute "Update SO52B0A1 set SumCount=" & rsTmp("SAmt") & _
                    " Where CitemCode=" & rsTmp("CitemCode") & _
                    " And CitemName='" & rsTmp("CitemName") & "'" & _
                    " And Period=" & rsTmp("Period") & _
                    " And Amt=" & rsTmp("Amt")
        rsTmp.MoveNext
      Wend
    cnn.CommitTrans
  Exit Sub
ChkErr:
  cnn.RollbackTrans
  Call ErrSub(Me.Name, "UpdataMDB")
End Sub

Private Sub SumMDB()
  On Error GoTo ChkErr
  Dim strSQL1 As String
    cnn.BeginTrans
      strSQL1 = "Select CitemCode,CitemName,Period,Amt" & strFieldAmt & strFieldCount & _
                  " Into SO52B0A2 From SO52B0A1 Group by " & _
                  " CitemCode,CitemName,Period,Amt"
      cnn.Execute strSQL1
    cnn.CommitTrans
    SendSQL strSQL1, True
  Exit Sub
ChkErr:
  cnn.RollbackTrans
  Call ErrSub(Me.Name, "SumMDB")
End Sub

Private Sub FieldName()
  On Error GoTo ChkErr
    Dim intFor As Integer
    TableFieldAmt = ""
    TableFieldCount = ""
    strFieldAmt = ""
    strFieldCount = ""
    For intFor = 1 To 31
      TableFieldAmt = TableFieldAmt & ",Amt" & Format(intFor, "00") & " Long"
      TableFieldCount = TableFieldCount & ",Count" & Format(intFor, "00") & " Long"
      strFieldAmt = strFieldAmt & ",iif(isnull(sum(Amt" & Format(intFor, "00") & ")),0,sum(Amt" & Format(intFor, "00") & ")) as Amt" & Format(intFor, "00")
      strFieldCount = strFieldCount & ",iif(isnull(sum(Count" & Format(intFor, "00") & ")),0,sum(Count" & Format(intFor, "00") & ")) as Count" & Format(intFor, "00")
    Next
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "FieldName")
End Sub
