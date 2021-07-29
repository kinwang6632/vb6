VERSION 5.00
Object = "{0CE6E756-92AE-11D3-8311-0080C8453DDF}#5.0#0"; "GiYM.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSO5510B 
   BorderStyle     =   1  '單線固定
   Caption         =   "Night-run 報表"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   6555
   Icon            =   "SO5510B.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   6555
   StartUpPosition =   1  '所屬視窗中央
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
      Height          =   435
      Left            =   4320
      TabIndex        =   5
      Top             =   1920
      Width           =   1200
   End
   Begin VB.CommandButton cmdSummary 
      Caption         =   "彙總表"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   600
      TabIndex        =   3
      Top             =   1920
      Width           =   1200
   End
   Begin VB.CommandButton cmdDetail 
      Caption         =   "匯出明細表"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2520
      TabIndex        =   4
      Top             =   1920
      Width           =   1200
   End
   Begin prjGiList.GiList gilServiceType 
      Height          =   315
      Left            =   1680
      TabIndex        =   0
      Top             =   300
      Width           =   3075
      _ExtentX        =   5424
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
      FldWidth1       =   750
      FldWidth2       =   2000
      F3Corresponding =   -1  'True
   End
   Begin Gi_YM.GiYM GiDataYm 
      Height          =   330
      Left            =   1680
      TabIndex        =   1
      Top             =   750
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
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
      Left            =   1680
      TabIndex        =   2
      Top             =   1200
      Visible         =   0   'False
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
   Begin MSComDlg.CommonDialog ComdPath 
      Left            =   5400
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "資料年月"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   720
      TabIndex        =   7
      Top             =   840
      Width           =   780
   End
   Begin VB.Label lblComp 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "服  務  別"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   735
      TabIndex        =   6
      Top             =   360
      Width           =   765
   End
End
Attribute VB_Name = "frmSO5510B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
Private objParentForm As Object
Private strServiceType As String
Private strViewName As String
Private strSqlString As String

Private Sub cmdDetail_Click()
On Error GoTo ChkErr
  Dim lngIndex, lngDataCount As Long
  Dim strSoName, strRptYm, StrTableName, strFileName, strLineData As String
  Dim rsTmp As New ADODB.Recordset
  Dim FSO As New FileSystemObject
  Dim strmTextFile1 As TextStream
  
  '匯出明細表
  '依公司別代號找出公司名稱
  Set rsTmp = gcnGi.Execute("Select Description From So507 Where CodeNo='" & gilCompCode.GetCodeNo & "'")
  strSoName = rsTmp("Description")
  Set rsTmp = Nothing
  
  strRptYm = GiDataYm.GetValue
  StrTableName = strSoName & "_SO001ALL_" & strRptYm
  
  '選擇匯出的檔名及路徑
  With ComdPath
    .FileName = StrTableName & ".TXT"
    .Filter = "文字檔 (*.txt;*.dat)|*.txt;*.dat|所有檔案 (*.*)|*.*"
    .Action = 1
    strFileName = .FileName
  End With
  
  '匯出
  Screen.MousePointer = vbHourglass
  Set strmTextFile1 = FSO.CreateTextFile(strFileName, True)
  Set rsTmp = gcnGi.Execute("Select * From " & StrTableName & _
                            " Where CustStatusCode=1 " & _
                            " And ClassCode1>=11010 " & _
                            " And ClassCode1<=91010 " & _
                            " And ClassCode1 Not In (14990,15990) ")
  '欄位名稱
  strLineData = ""
  For lngIndex = 0 To rsTmp.Fields.Count - 1
    strLineData = strLineData & rsTmp.Fields(lngIndex).Name & vbTab
  Next
  strmTextFile1.Write strLineData & vbCrLf
  
  lngDataCount = 0
  While Not rsTmp.EOF
    strLineData = ""
    For lngIndex = 0 To rsTmp.Fields.Count - 1
      strLineData = strLineData & rsTmp.Fields(lngIndex) & vbTab
    Next
    lngDataCount = lngDataCount + 1
    strmTextFile1.Write strLineData & vbCrLf
    DoEvents
    rsTmp.MoveNext
  Wend
  Screen.MousePointer = vbDefault
  Set rsTmp = Nothing
  Set strmTextFile1 = Nothing
  Set FSO = Nothing
  MsgBox "匯出完成，筆數：" & lngDataCount, vbInformation, "訊息"
ChkErr:
  Set strmTextFile1 = Nothing
  Call ErrSub(Me.Name, "cmdDetail_Click")
End Sub

Private Sub cmdExit_Click()
On Error GoTo ChkErr
  Unload Me
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdExit_Click")
End Sub

Private Sub cmdSummary_Click()
On Error GoTo ChkErr
  Dim blnGoOn As Boolean
  Dim rsTmp As New ADODB.Recordset
  
  'INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName) VALUES ( 'SO19J0', null, 'CP 帳務資料拋轉作業', '', 'SO1900', Null);
  
  blnGoOn = ChkDTok And IsDataOk
  If blnGoOn Then
  
    '視窗暫時 disabled，以免 USer 重複按
    Screen.MousePointer = vbHourglass
    Me.Enabled = False
    
    Call ReadyGoPrint
    Call subChoose
    Call subCreateView
    
    Set rsTmp = gcnGi.Execute("SELECT Count(*) As DataCount From " & strViewName)
    If rsTmp("DataCount") = 0 Then
      Call MsgNoRcd
      SendSQL , , True
    Else
      strsql = "Select * From " & strViewName & " V"
      Call PrintRpt2(GetPrinterName(5), RptName("SO5510", 1), , "現有各類客戶數統計 [SO5510B]", strsql, strChooseString, , True, , , , GiPaperPortrait)
    End If
    Set rsTmp = Nothing
    
    '視窗恢復 enabled
    Me.Enabled = True
    Screen.MousePointer = vbDefault
  End If
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdSummary_Click")
End Sub

Private Sub Form_Load()
On Error GoTo ChkErr

  '公司別，隱藏，報表顯示用
  Call SetgiList(gilCompCode, "CodeNo", "Description", "CD039", , , , , , , GetCompCodeFilter(gilCompCode))
  gilCompCode.SetCodeNo garyGi(9)
  gilCompCode.Query_Description
  
  '服務類別
  Call SelectServType(garyGi(9), gilServiceType)
  gilServiceType.SetCodeNo strServiceType
  gilServiceType.Query_Description
  
  '資料年月
  GiDataYm.SetValue (RightDate)
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "Form_Load")
End Sub

Private Function IsDataOk() As Boolean
On Error GoTo ChkErr
  IsDataOk = True
  If gilServiceType.GetCodeNo = "" Then
    IsDataOk = False
    MsgBox "請先指定服務別。", vbExclamation, "警告"
    gilServiceType.SetFocus
  ElseIf GiDataYm.Text = "" Then
    IsDataOk = False
    MsgBox "請先指定資料年月。", vbExclamation, "警告"
    GiDataYm.SetFocus
  End If
  Exit Function
ChkErr:
  Call ErrSub(Me.Name, "IsDataOk")
End Function

Private Sub subChoose()
  On Error GoTo ChkErr
    Dim lngClassCode, lngYear, lngMonth, lngYear2, lngMonth2 As Long
    Dim strsql, strSQL1, strSQL2, strSQL3, strSQL4, RptYm, RptYm2, strCompCode As String
    Dim strRowName(6)
    Dim rsTmp As New ADODB.Recordset
    
    RptYm = GiDataYm.GetValue
    lngYear = RptYm \ 100
    lngMonth = (RptYm Mod 100) - 1
    If lngMonth < 1 Then
      lngYear = lngYear - 1
      lngMonth = 12
    End If
    RptYm2 = lngYear * 100 + lngMonth
    
    lngYear2 = RptYm \ 100
    lngMonth2 = (RptYm Mod 100) + 1
    If lngMonth2 > 12 Then
      lngYear2 = lngYear2 + 1
      lngMonth2 = 1
    End If
    
    strCompCode = gilCompCode.GetCodeNo
    
    strSqlString = ""
    strRowName(1) = "一般戶"
    strRowName(2) = "優待戶"
    strRowName(3) = "免費戶"
    strRowName(4) = "大樓統收"
    strRowName(5) = "大樓個收"
    strRowName(6) = "九太公播戶"
    
    For lngClassCode = 1 To 6
      strSQL1 = "Select Nvl(Sum(CustCount),0) As CustCount From So511A " & _
                "Where (RptYm=" & RptYm & ") " & _
                "And (DataType=0) " & _
                "And (NextYm>" & RptYm & ") " & _
                "And (CompCode=" & strCompCode & ") " & _
                "And (ClassCode=" & lngClassCode & ") "
      strSQL2 = "SELECT Nvl(Sum(CustCount),0) As CustCount From So511A " & _
                "Where (RptYm=" & RptYm & ") " & _
                "And (DataType=0) " & _
                "And (NextYm=" & RptYm & ") " & _
                "And (CompCode=" & strCompCode & ") " & _
                "And (ClassCode=" & lngClassCode & ") "
      strSQL3 = "SELECT Nvl(Sum(CustCount),0) As CustCount From So511A " & _
                "Where (RptYm=" & RptYm & ") " & _
                "And (DataType=-2) " & _
                "And (NextYm=" & RptYm2 & ") " & _
                "And (CompCode=" & strCompCode & ") " & _
                "And (ClassCode=" & lngClassCode & ") "
      strSQL4 = "SELECT Nvl(Sum(CustCount),0) As CustCount From So511A " & _
                "Where (RptYm=" & RptYm & ") " & _
                "And (DataType=-2) " & _
                "And (NextYm<" & RptYm2 & ") " & _
                "And (CompCode=" & strCompCode & ") " & _
                "And (ClassCode=" & lngClassCode & ") "
      strSqlString = strSqlString & "Select '" & strRowName(lngClassCode) & "' RowName, " & _
                                    "(" & strSQL1 & ") A, " & _
                                    "(" & strSQL2 & ") B, " & _
                                    "(" & strSQL3 & ") C, " & _
                                    "(" & strSQL4 & ") D " & _
                                    " From Dual "
      If lngClassCode < 6 Then
        strSqlString = strSqlString & "Union All "
      End If
    Next
  '選取的條件
    Set rsTmp = gcnGi.Execute("Select Description From Cd019 Where (RefNo=1) And (ServiceType='C')")
    RptYm = GiDataYm.GetValue
    lngYear = RptYm \ 100
    lngMonth = RptYm Mod 100
    strChooseString = "公司別　: " & subSpace(gilCompCode.GetDescription) & ";" & _
                      "服務類別: " & subSpace(gilServiceType.GetDescription) & ";" & _
                      "收費項目: " & subSpace(rsTmp("Description")) & ";" & _
                      "大樓統收子戶是否列入統計: 是;" & _
                      "同一大樓戶視為同一客戶: 否;" & _
                      "客戶狀態:   正常;" & _
                      "基準日期: " & lngYear2 & "/" & lngMonth2 & "/3;" & _
                      "歸屬月份: " & lngYear & "/" & lngMonth
    Set rsTmp = Nothing
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "subChoose")
End Sub

Private Sub subCreateView()
  Dim strView As String
  
  On Error Resume Next
  gcnGi.Execute "Drop Table " & strViewName
  
On Error GoTo ChkErr
  strViewName = GetTmpViewName
  strView = "Create Table " & strViewName & " As (" & strSqlString & ")"
  gcnGi.Execute strView
  SendSQL strView, True
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "subCreateView")
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
  If strViewName <> "" Then
    gcnGi.Execute "Drop Table " & strViewName
  End If
End Sub

Public Property Get uParentForm() As Object
  Set uParentForm = objParentForm
End Property

Public Property Set uParentForm(ByVal vData As Object)
  Set objParentForm = vData
End Property

Public Property Get uServiceType() As String
  uServiceType = strServiceType
End Property

Public Property Let uServiceType(ByVal vData As String)
  strServiceType = vData
End Property
