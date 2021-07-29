VERSION 5.00
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSO1193A 
   Caption         =   "VOD結算作業[SO32E02A]"
   ClientHeight    =   6555
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5190
   Icon            =   "SO1193A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   5190
   StartUpPosition =   2  '螢幕中央
   Begin MSComDlg.CommonDialog comdPath 
      Left            =   0
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "...."
      Height          =   315
      Left            =   3450
      TabIndex        =   9
      Top             =   5640
      Width           =   465
   End
   Begin VB.TextBox txtXLSFile 
      Height          =   315
      Left            =   1050
      TabIndex        =   8
      Top             =   5640
      Width           =   2355
   End
   Begin VB.CommandButton cmdSTB 
      Caption         =   "設備"
      Height          =   315
      Left            =   3570
      TabIndex        =   2
      Top             =   3210
      Width           =   705
   End
   Begin VB.TextBox txtSrNO 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2160
      TabIndex        =   1
      Top             =   3210
      Width           =   1335
   End
   Begin VB.TextBox txtCustid 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2160
      TabIndex        =   0
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "F2. 確定"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   90
      TabIndex        =   10
      Top             =   6150
      Width           =   1365
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&X)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1875
      TabIndex        =   11
      Top             =   6135
      Width           =   1365
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
      Height          =   375
      Left            =   3615
      TabIndex        =   12
      Top             =   6135
      Width           =   1425
   End
   Begin VB.Frame Frame1 
      Height          =   1305
      Left            =   240
      TabIndex        =   20
      Top             =   4260
      Width           =   2745
      Begin VB.CheckBox chkRun 
         Caption         =   "執行結算作業"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   180
         TabIndex        =   7
         Top             =   840
         Width           =   1635
      End
      Begin VB.CheckBox chkTran2 
         Caption         =   "產生未執行結算明細報表"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   180
         TabIndex        =   6
         Top             =   480
         Width           =   2475
      End
      Begin VB.CheckBox chkTran1 
         Caption         =   "產生預計結算明細報表"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   180
         TabIndex        =   5
         Top             =   150
         Value           =   1  '核取
         Width           =   2385
      End
   End
   Begin VB.Frame fraLast 
      Caption         =   "上次結算記錄"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1290
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   4950
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         Caption         =   "結算截止日:"
         Height          =   180
         Left            =   825
         TabIndex        =   19
         Top             =   300
         Width           =   945
      End
      Begin VB.Label lblTranDate 
         BackColor       =   &H00E0E0E0&
         Caption         =   "YY/MM/DD"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1875
         TabIndex        =   18
         Top             =   285
         Width           =   1050
      End
      Begin VB.Label lbl2 
         AutoSize        =   -1  'True
         Caption         =   "上次執行結算時間:"
         Height          =   180
         Left            =   240
         TabIndex        =   17
         Top             =   630
         Width           =   1485
      End
      Begin VB.Label lblUpdTime 
         BackColor       =   &H00E0E0E0&
         Caption         =   "YY/MM/DD HH:MM:SS"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1875
         TabIndex        =   16
         Top             =   645
         Width           =   1980
      End
      Begin VB.Label lbl3 
         AutoSize        =   -1  'True
         Caption         =   "上次執行結算人員: "
         Height          =   180
         Left            =   240
         TabIndex        =   15
         Top             =   960
         Width           =   1530
      End
      Begin VB.Label lblUpdEn 
         BackColor       =   &H00E0E0E0&
         Caption         =   "XXXXXXXXXXXXXXXXXXXX"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1875
         TabIndex        =   14
         Top             =   975
         Width           =   2760
      End
   End
   Begin prjGiList.GiList gilCompCode 
      Height          =   300
      Left            =   1920
      TabIndex        =   31
      Top             =   1440
      Width           =   3090
      _ExtentX        =   5450
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FldWidth1       =   900
      FldWidth2       =   1855
      F2Corresponding =   -1  'True
   End
   Begin prjGiList.GiList gilServiceType 
      Height          =   315
      Left            =   1920
      TabIndex        =   32
      Top             =   1770
      Width           =   3090
      _ExtentX        =   5450
      _ExtentY        =   556
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FldWidth1       =   900
      FldWidth2       =   1855
      F2Corresponding =   -1  'True
   End
   Begin Gi_Date.GiDate gdaEndDate 
      Height          =   270
      Left            =   2160
      TabIndex        =   3
      Top             =   3540
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   476
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ReplaceText     =   -1  'True
   End
   Begin Gi_Date.GiDate gdaShouldDate 
      Height          =   270
      Left            =   2160
      TabIndex        =   4
      Top             =   3870
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   476
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ReplaceText     =   -1  'True
   End
   Begin VB.Label lblXLSPath 
      Caption         =   "檔案路徑"
      Height          =   255
      Left            =   240
      TabIndex        =   34
      Top             =   5700
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "應收日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1290
      TabIndex        =   33
      Top             =   3900
      Width           =   780
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "(含此日之前的資料將被結算)"
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
      Height          =   390
      Left            =   3480
      TabIndex        =   30
      Top             =   3540
      Width           =   1665
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblEndDate 
      AutoSize        =   -1  'True
      Caption         =   "結算日期期限"
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
      Height          =   255
      Left            =   900
      TabIndex        =   29
      Top             =   3540
      Width           =   1200
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblCustid 
      AutoSize        =   -1  'True
      Caption         =   "客編"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1680
      TabIndex        =   28
      Top             =   2880
      Width           =   375
   End
   Begin VB.Label lblPara36 
      BackColor       =   &H00E0E0E0&
      Caption         =   "XXXXXXXXX"
      Height          =   255
      Left            =   2160
      TabIndex        =   27
      Top             =   2550
      Width           =   1215
   End
   Begin VB.Label lblPara35 
      BackColor       =   &H00E0E0E0&
      Caption         =   "XXXXXXXXX"
      Height          =   195
      Left            =   2160
      TabIndex        =   26
      Top             =   2190
      Width           =   1200
   End
   Begin VB.Label lbl5 
      AutoSize        =   -1  'True
      Caption         =   "VOD結算作業月數限額"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   25
      Top             =   2550
      Width           =   1965
   End
   Begin VB.Label lbl4 
      AutoSize        =   -1  'True
      Caption         =   "VOD結算作業金額限額"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   24
      Top             =   2190
      Width           =   1965
   End
   Begin VB.Label lblCompCode 
      AutoSize        =   -1  'True
      Caption         =   "公司別"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1260
      TabIndex        =   23
      Top             =   1455
      Width           =   585
   End
   Begin VB.Label lblSrNO 
      AutoSize        =   -1  'True
      Caption         =   "STB設備流水號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   780
      TabIndex        =   22
      Top             =   3210
      Width           =   1305
   End
   Begin VB.Label lblServicetype 
      AutoSize        =   -1  'True
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
      Height          =   195
      Left            =   1080
      TabIndex        =   21
      Top             =   1830
      Width           =   780
   End
End
Attribute VB_Name = "frmSO1193A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'#5327 需求 By Kin 2009/11/06
Option Explicit
Dim lngTranDate As Long             ' 結算截止日
Dim rsCD039 As New ADODB.Recordset  ' 公司別代碼檔
Dim intPara35 As Integer, intPara36 As Integer
Dim strPara35In As String, strPara36In As String
Dim strSQL As String
Dim strChooseString As String
Dim strFormula As String
Dim strCompCode As String
Dim strServiceType As String
Dim strCustid As String
Dim strSrNO As String
Dim strEndDate As String
Dim strMsg As String
Dim blnDO As Boolean
Dim blnGive As Boolean
Private blnTrans As Boolean
Private strCitemCode As String
Private strCitemName As String
Private strRunTime As String
Private blnBeginTrans As Boolean
Private strCMCode As String
Private strCMName As String
Private strUCCode As String
Private strUCName As String
Private strPTCode As String
Private strPTName As String
Private intPara14 As Integer
Private strSalePointCode As String
Private strSalePointName As String
Private rs014 As New ADODB.Recordset
Private blnNOUpdate As Boolean
Private blnNoData As Boolean
Private rs033Clone As ADODB.Recordset
Private rs033VodClone As ADODB.Recordset
Private rs182Clone As ADODB.Recordset
Private strFaciSeqNo As String
Private strFaciSNo As String
Private blnAutoEXE As Boolean
Private strViewXLSName As String
Private Sub cmdCancel_Click()
  On Error Resume Next
    Unload Me
End Sub

Private Sub cmdOpen_Click()
   On Error GoTo ChkErr
        With comdPath
            .DialogTitle = "選擇輸出路徑"
            .Filter = "Microsoft Excel 活頁簿 (*.xls)|*.xls"
            .Action = 1
             If .FileName <> "" Then txtXLSFile = .FileName
        End With
    Exit Sub
ChkErr:
   ErrSub Me.Name, "cmdOpen_Click"
End Sub

Private Sub cmdPrevRpt_Click()
  On Error GoTo ChkErr
    Screen.MousePointer = vbHourglass
      Call PreviousRpt(GetPrinterName(5), RptName("SO32E02"), Me.Caption)
    Screen.MousePointer = vbDefault
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdPrevRpt_Click")
End Sub

Private Sub cmdSTB_Click()
  On Error GoTo ChkErr
    If txtCustID.Text = "" Then
        MsgBox "請輸入客戶編號再做選取設備動作！", vbInformation, "訊息"
        Exit Sub
    End If
    With frmSO1131E
        .uParentForm = Me
        .uEditMode = giEditModeEdit
        .uCustId = txtCustID.Text
        
        .uServiceType = gilServiceType.GetCodeNo
        .uEditMode = giEditModeEdit
        .uRefNo = 3
        .Move (Screen.Width - frmSO1131E.Width) / 2, (Screen.Height - Me.Height) / 2 + cmdSTB.Top + cmdSTB.Height + 400
        .Show 1, Me
        txtSrNO.Tag = .uFaciSeqNo
        txtSrNO.Text = .uFaciSNo
        On Error Resume Next
        SetFocus
    End With
  Exit Sub
ChkErr:
    ErrSub Name, "cmdSTB_Click"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error Resume Next
    If KeyCode = vbKeyF2 Then cmdOK.Value = True
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
    strServiceType = "": strCompCode = "": strEndDate = "": strCustid = "": strSrNO = ""
    strPara35In = "": strPara36In = "": strFaciSeqNo = "": strFaciSNo = ""
    blnGive = False
    SendSQL , , True
    'Call ReleaseCOM(Me)
End Sub
Private Function FillXlsData(ByVal aFile As String, ByRef aMsg As String) As Boolean
  On Error GoTo ChkErr
    Dim rsXls As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Set rsXls = GetXlsRs(aFile, "Sheet1", " * ")
    If rsXls Is Nothing Then Exit Function
    If rsXls.State = adStateClosed Then Exit Function
    If Not GetRS(rsTmp, "SELECT * FROM " & GetOwner & strViewXLSName & " WHERE 1=0 ", gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
    If Not rsXls.EOF Then rsXls.MoveFirst
    Do While Not rsXls.EOF
        rsTmp.AddNew
        If rsXls("客編") & "" = "" Then aMsg = "Excel含有空白資料！請檢查Excel": Exit Function
        If rsXls("服務帳號") & "" = "" Then aMsg = "Excel含有空白資料！請檢查Excel": Exit Function
        If rsXls("設備序號") & "" = "" Then aMsg = "Excel含有空白資料！請檢查Excel": Exit Function
        rsTmp("CUSTID") = rsXls("客編") & ""
        rsTmp("VODACCOUNTID") = rsXls("服務帳號") & ""
        rsTmp("FACISNO") = rsXls("設備序號")
        'rsTmp("FACISEQNO") = rsXls("FACISEQNO") & ""
        rsTmp.Update
        rsXls.MoveNext
    Loop
    FillXlsData = True
    Exit Function
ChkErr:
  Call ErrSub(Me.Name, "FillXlsData")
End Function
Private Function CrtTmpTable() As Boolean
  On Error GoTo ChkErr
    Dim strString As String
    strViewXLSName = GetTmpViewName
    On Error Resume Next
        strString = "Create Table " & GetOwner & strViewXLSName & "(CUSTID Number(8), VODACCOUNTID Number(15), FACISNO Varchar2(20), FACISEQNO Varchar2(15))"
        gcnGi.Execute strString
        CrtTmpTable = True
    Exit Function
ChkErr:
  Call ErrSub(Me.Name, "CrtTmpTable")
End Function
Private Sub cmdOK_Click()
  On Error GoTo ChkErr
    'If gdaEndDate.GetValue = "" Then MsgBox "請填入結算日期期限!", vbInformation, "訊息!": Exit Sub
    blnNoData = False
    If Not IsDataOk Then Exit Sub
    strServiceType = gilServiceType.GetCodeNo
    strCompCode = gilCompCode.GetCodeNo
    strEndDate = gdaEndDate.GetValue
    strCustid = txtCustID
    strSrNO = txtSrNO
    strRunTime = RightNow
    If Not subChoose Then Exit Sub
    If Not blnAutoEXE Then cmdCancel.SetFocus
    Me.Enabled = False
    
    
    If chkTran1.Value = 1 Then
        'If Not subChoose(1, strCompCode, strServiceType, strEndDate, , strCustid, strSrNO) Then Exit Sub
        Call subPrint(0)
    End If
    If chkTran2.Value = 1 Then
        'If Not subChoose(2, strCompCode, strServiceType, strEndDate, , strCustid, strSrNO) Then Exit Sub
        Call subPrint(1)
    End If
    If chkRun.Value = 1 Then
        Screen.MousePointer = vbHourglass
        
        If Not blnTrans Then gcnGi.BeginTrans
        If Not CreateBill(blnNoData) Then
            If Not blnTrans Then gcnGi.RollbackTrans
            GoTo 88
        End If
        If blnGive Then
            If Not NewSO062(blnNoData, strMsg) Then
                If Not blnTrans Then gcnGi.RollbackTrans
                GoTo 88
            Else
                blnDO = True
            End If
        Else
            If blnNoData Then
                strMsg = "結算作業完成！"
            Else
                strMsg = "無結算資料！"
            End If
            blnDO = True
        End If
        If (blnNOUpdate) And (Not blnTrans) Then
            gcnGi.RollbackTrans
        Else
            If Not blnTrans Then gcnGi.CommitTrans
        End If
    End If
    If Not blnAutoEXE Then
        If blnGive = True And chkRun.Value = 1 Then
            Screen.MousePointer = vbDefault
            MsgBox strMsg, vbInformation, "訊息"
            'Unload Me
        Else
            If (Not blnGive) And (chkRun.Value = 1) And (Not blnNOUpdate) Then MsgBox strMsg, vbInformation, "訊息"
        End If
    End If
   
88:
'    If Not blnTrans Then
'        If blnNoData Then
'            gcnGi.RollbackTrans
'        Else
'            gcnGi.CommitTrans
'        End If
'    End If
    Me.Enabled = True
    Screen.MousePointer = vbDefault
    If blnAutoEXE Then Unload Me
    On Error Resume Next
    gcnGi.Execute "Drop Table " & GetOwner & strViewXLSName
    Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdOk_Click")
End Sub
Private Sub subPrint(ByVal aEXETYPE As Integer)
  On Error GoTo ChkErr
    Dim rsTmp As New ADODB.Recordset
    Dim strMsg As String
    Dim strPrintSQL As String
    Screen.MousePointer = vbHourglass
    Me.Enabled = False
    ReadyGoPrint
    Call subCreateView(aEXETYPE)
    DoEvents
    Set rsTmp = gcnGi.Execute("SELECT Count(*) as intCount From " & strViewName & " Where RowNum=1")
    If rsTmp("intCount") = 0 Then
        If aEXETYPE = 0 Then
            strMsg = "無預計執行結算資料"
        Else
            strMsg = "無未執行結算資料"
        End If
        MsgBox strMsg, vbExclamation, "提示"
        SendSQL , , True
    Else
        strPrintSQL = "SELECT * From " & strViewName & " V"
        strFormula = IIf(strFormula = "", "Sort0={V.CustId}", strFormula & ";Sort0={V.CUSTID}")
        strFormula = strFormula & ";Sort1={V.FACISNO};Sort2={V.VODACCOUNTID};Type=" & aEXETYPE + 1
        Call PrintRpt(GetPrinterName(5), "SO32E2A.RPT", , "預計結算明細報表 [SO32E2A]", strPrintSQL, strChooseString, , True, , , strFormula)
        SendSQL , , True
    End If
    Me.Enabled = True
    Screen.MousePointer = vbDefault
    CloseRecordset rsTmp
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "subPrint")
End Sub
'aEXETYPE = 0 代表預計執行 aEXETYPE =1 代表未執行
Private Function subCreateView(ByVal aEXETYPE As Integer) As Boolean
    Dim strView As String
    On Error Resume Next
      gcnGi.Execute "Drop View " & strViewName
    On Error GoTo ChkErr
    Dim strCrtSQL As String
    strViewName = GetTmpViewName
    '#5521 產生出來的金額要再扣掉SUM3 By Kin 2010/02/03
'    strCrtSQL = "SELECT DISTINCT CUSTID,VODACCOUNTID," & _
'                "FACISNO,SUM(MINUSCREDIT) MINUSCREDIT,SUM(AddCreditit) ADDCREDIT," & _
'                "ABS(SUM(MINUSCREDIT)-SUM(ADDCREDITIT)-SUM3) TOTAL  FROM (" & _
'                strSQL & ") WHERE EXETYPE=" & aEXETYPE & _
'                " GROUP BY CUSTID,VODACCOUNTID,FACISNO,SUM3"
    strCrtSQL = "SELECT DISTINCT CUSTID,VODACCOUNTID," & _
                "FACISNO,SUM(MINUSCREDIT) MINUSCREDIT,SUM(AddCreditit) ADDCREDIT," & _
                "ABS(SUM1-SUM2-SUM3) TOTAL  FROM (" & _
                strSQL & ") WHERE EXETYPE=" & aEXETYPE & _
                " GROUP BY CUSTID,VODACCOUNTID,FACISNO,SUM1,SUM2,SUM3"

'    strSQL = "SELECT * FROM (" & strSQL & ") WHERE EXETYPE=" & aEXETYPE
    strView = "Create View " & strViewName & " as ( " & strCrtSQL & ")"
    gcnGi.Execute strView
    SendSQL strView, True
    Exit Function
ChkErr:
  Call ErrSub(Me.Name, "subCreateView")
End Function
Private Function IsDataOk() As Boolean
  On Error GoTo ChkErr
    If gdaEndDate.GetValue & "" = "" Then
        MsgBox "請填入結算日期期限!", vbInformation, "訊息!"
        If gdaEndDate.Enabled Then gdaEndDate.SetFocus
        Exit Function
    End If
    If chkTran1.Value = 1 Or chkTran2.Value = 1 Or chkRun.Value = 1 Then
        'IsDataOk = True
    Else
        MsgBox "請選擇要執行的項目！", vbInformation, "訊息"
        Exit Function
    End If
    If gilCompCode.GetCodeNo = "" Then
        MsgBox "請輸入公司別！", vbInformation, "訊息"
        Exit Function
    End If
    '#5488 如果畫面有輸入EXCEL,則要建立暫存檔,並把EXCEL資料匯進來 By Kin 2010/01/20
    If txtXLSFile.Text <> "" Then
        If Not CrtTmpTable Then
            MsgBox "建立EXCEL暫存資料表錯誤！", vbInformation, "訊息"
            Exit Function
        End If
        Dim aMsg As String
        If Not FillXlsData(txtXLSFile.Text, aMsg) Then
            If aMsg <> "" Then
                MsgBox aMsg, vbInformation, "訊息"
                Exit Function
            Else
                MsgBox "EXCEL暫存資料建立失敗！", vbInformation, "訊息"
                Exit Function
            End If
            
        End If
    End If
    
    IsDataOk = True
    Exit Function
ChkErr:
  Call ErrSub(Me.Name, "IsDataOK")
End Function
Private Function NewSO062(ByVal aHaveData As Boolean, Optional ByRef aMsg As String) As Boolean
On Error GoTo ChkErr
    Dim strCH As String
    Dim strQry As String
    Dim rs062 As New ADODB.Recordset
    strQry = "SELECT * FROM " & GetOwner & "SO062 " & _
            " WHERE TYPE =4 AND COMPCODE=" & gilCompCode.GetCodeNo & _
            " AND SERVICETYPE = '" & gilServiceType.GetCodeNo & "'"
    If Not GetRS(rs062, strQry, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
    strCH = "公司別=" & gilCompCode.GetCodeNo & _
            " ;服務類別=" & gilServiceType.GetCodeNo & _
            " ; 結算日期期限=" & gdaEndDate.GetValue(True) & _
            " ; 客編=" & txtCustID.Text & "; STB設備序號=" & txtSrNO.Text & ""
    If aHaveData Then
        If rs062.EOF Then
            rs062.AddNew
        End If
        With rs062
            .Fields("Type") = 4
            .Fields("TranDate") = gdaEndDate.GetValue(True)
            .Fields("UpdEn") = garyGi(1)
            .Fields("UpdTime") = GetDTString(strRunTime)
            .Fields("Para") = strCH
            .Fields("CompCode") = gilCompCode.GetCodeNo
            .Fields("ServiceType") = gilServiceType.GetCodeNo
            .Update
        End With
        aMsg = "結算作業完成！"
        
    Else
        aMsg = "無結算資料！"
    End If
    

    NewSO062 = True
Exit Function
ChkErr:
    Call ErrSub(Me.Name, "NewSO062")
End Function
Private Function GetSQL2(ByVal aType As Integer, ByVal aWhere As String) As String
    On Error GoTo ChkErr
    Dim strWhere As String
    Dim strWhere2 As String
    Dim strWhere3 As String
    Dim strTotalQry As String
    Dim strSum1Qry As String
    Dim strSum2Qry As String
    Dim strSum3Qry As String
    Dim strSubQry As String
    Dim strRetSQL As String
    Dim strRunSQL As String
    Dim strFields As String
    Dim str004Where As String
    Dim strSubExcel As String
    '拆除,SO004要過濾拆除的資料
    'str004Where = "(A.INSTDATE IS NOT NULL AND A.PRDATE IS NOT NULL AND A.PRDATE > A.INSTDATE)"

    
    
    '****************************期間內總訂購點數的 Where 條件 *******************************
'    strWhere = " WHERE A.CompCode = B.CompCode And A.ServiceType = B.ServiceType " & _
'                " And CREDITTYPECODE In (SELECT CODENO FROM " & GetOwner & "CD108 " & _
'                " WHERE CD108.REFNO = 3 AND NVL(CD108.STOPFLAG,0) = 0 AND B.SERVICETYPE=CD108.SERVICETYPE)" & _
'                " AND NVL(B.CloseFlag,0)=0 AND NVL(B.STOPFLAG,0)= 0 AND A.SEQNO = B.FACISEQNO "
                
    strWhere = " WHERE A.CompCode = B.CompCode  " & _
                " AND NVL(B.STOPFLAG,0)= 0 AND A.SEQNO = B.FACISEQNO " & _
                " And A.VodAccountid=B.VodAccountId "
    '******************************************************************************************
    
    '****************************可使用點數的 Where 條件 *****************************************
'    strWhere2 = " WHERE A.CompCode = B.CompCode And A.ServiceType = B.ServiceType " & _
'               " And CREDITTYPECODE In (SELECT CODENO FROM " & GetOwner & "CD108 " & _
'               " WHERE CD108.REFNO IN(1,2) AND NVL(CD108.STOPFLAG,0) = 0 AND B.SERVICETYPE=CD108.SERVICETYPE)" & _
'               " AND NVL(B.STOPFLAG,0)=0 AND A.SEQNO=B.FACISEQNO "
               
    strWhere2 = " WHERE A.CompCode = B.CompCode  " & _
               " AND NVL(B.STOPFLAG,0)=0 AND A.SEQNO=B.FACISEQNO And A.VodAccountid=B.VodAccountId"
    '**********************************************************************************************
    
    '****************************已結算過期間內總訂購點數的 Where 條件 *******************************
    '這一段語法不要了,所以加一個1=0
    '#5432 把這段加進去,SUM(SO033.SHOULDAMT) By Kin 2009/12/09
    strWhere3 = " WHERE A.CompCode = B.CompCode  " & _
                " AND NVL(B.STOPFLAG,0)=0 AND A.SEQNO=B.FACISEQNO AND A.VODACCOUNTID=B.VODACCOUNTID "
    '**************************************************************************************************
    If aWhere <> Empty Then
        strWhere = strWhere & " AND " & aWhere
        '#5551 結清日期除了SUM1要串其它不要串 By Kin 2010/02/22
        If gdaEndDate.GetValue <> "" Then
            Call subAnd2(strWhere, " To_Date(To_Char(B.RecTime,'YYYYMMDD'),'YYYYMMDD') <= " & _
                        "To_Date('" & gdaEndDate.GetValue & "','YYYYMMDD')")
            
        End If
        strWhere2 = strWhere2 & " AND " & aWhere
        strWhere3 = strWhere3 & " AND " & aWhere
    End If
    '****************************要結算的SQL語法*****************************************************
    
    '****************************算出期間內總訂購數之後還要判斷合計或未結算次數的條件*****************
    '#5488 EXCLE有資料的話,強迫將結算次數設定成參數+1，以便可以強迫出帳 By Kin 2010/01/20
    If strViewXLSName <> "" And txtXLSFile.Text <> "" Then
        strSubExcel = " Union " & _
                "SELECT B.VODACCOUNTID,0 MINUSTOTAL," & intPara36 + 1 & " MAXNOCLOSETIME, " & _
                "0 ADDTOTAL " & _
                " FROM " & GetOwner & " SO004 A," & GetOwner & "SO182 B," & GetOwner & strViewXLSName & " C " & _
                strWhere & " AND A.VODACCOUNTID=C.VODACCOUNTID AND A.CUSTID=C.CUSTID " & _
                " AND NVL(A.FACISNO,'X')=Nvl(C.FACISNO,'X') " & _
                " GROUP BY B.VODACCOUNTID "
    End If
    
    strSubQry = "SELECT B.VODACCOUNTID,SUM(B.MINUSCREDIT) MINUSTOTAL,Max(Nvl(B.NOCLOSETIME,0)) MAXNOCLOSETIME, " & _
                    "SUM(B.AddCredit) ADDTOTAL " & _
                    " FROM " & GetOwner & "SO004 A," & GetOwner & "SO182 B" & strWhere & _
                    " GROUP BY B.VODACCOUNTID " & strSubExcel

    strSubQry = "SELECT VODACCOUNTID,SUM(MINUSTOTAL) MINUSTOTAL,SUM(MAXNOCLOSETIME) MAXNOCLOSETIME, " & _
            "SUM(ADDTOTAL) ADDTOTAL FROM (" & strSubQry & ") GROUP BY VODACCOUNTID "
    '**************************************************************************************
    strSum1Qry = "SELECT B.VODACCOUNTID,B.NOCLOSETIME,C.MINUSTOTAL,B.MINUSCREDIT " & _
                        " FROM " & GetOwner & "SO004 A," & GetOwner & "SO182 B,(" & strSubQry & ") C " & _
                        strWhere & " AND B.VODACCOUNTID=C.VODACCOUNTID "
                        

                         
    ' " AND (B.NOCLOSETIME > " & intPara36 & " OR C.TOTAL > " & intPara35 & ")"
                         
    '算出期間內總訂購數
    strSum1Qry = " SELECT VODACCOUNTID,SUM(MINUSCREDIT) SUM1,0 SUM2,0 SUM3 " & _
                 " FROM (" & strSum1Qry & ") GROUP BY VODACCOUNTID "
                 
    '算出可使用點數
    strSum2Qry = "SELECT B.VODACCOUNTID,0 SUM1,SUM(NVL(B.AddCredit,0)) SUM2,0 SUM3 " & _
                " FROM " & GetOwner & "SO004 A," & GetOwner & "SO182 B " & _
                strWhere2 & " GROUP BY B.VODACCOUNTID "
                
    '***************************************************************************
    '算出已結算過期間內總訂購點數
    '#5432 strSum3Qry 去Sum SO033.ShouldAmt By Kin 2009/12/09
    strSum3Qry = "SELECT A.CUSTID,B.VODACCOUNTID,A.SEQNO " & _
                " FROM " & GetOwner & "SO004 A," & GetOwner & "SO182 B" & _
                strWhere3 & " GROUP BY A.CUSTID,B.VODACCOUNTID,A.SEQNO"
                
    strSum3Qry = "SELECT A.VODACCOUNTID,0 SUM1,0 SUM2,SUM(B.SHOULDAMT) SUM3 " & _
                " FROM (" & strSum3Qry & ") A," & GetOwner & "SO033 B," & _
                GetOwner & "CD019 C, " & GetOwner & "CD013 D " & _
                " WHERE B.CUSTID=A.CUSTID AND B.FACISEQNO=A.SEQNO AND B.UCCODE IS NOT NULL " & _
                " AND B.CITEMCODE=C.CODENO AND NVL(C.REFNO,0)=21  " & _
                " AND B.UCCODE=D.CODENO AND NVL(D.REFNO,0) Not In (3,7) " & _
                " AND B.BILLNO || B.ITEM NOT IN(SELECT NVL(B.BILLNO,'X')||NVL(B.ITEM,0) FROM " & _
                GetOwner & "SO004 A," & GetOwner & "SO182 B " & strWhere3 & ") " & _
                " GROUP BY A.VODACCOUNTID "
    '***************************************************************************
    
'    strSum3Qry = "SELECT B.VODACCOUNTID,0 SUM1,0 SUM2,SUM(NVL(B.MinusCredit,0)) SUM3 " & _
'                " FROM " & GetOwner & "SO004 A," & GetOwner & "SO182 B" & _
'                strWhere3 & " GROUP BY B.VODACCOUNTID "
    
    '將資料合併
    strTotalQry = "SELECT VODACCOUNTID,SUM(SUM1) SUM1, " & _
                " SUM(SUM2) SUM2,SUM(SUM3) SUM3 FROM (" & _
                strSum1Qry & " UNION " & strSum2Qry & " UNION " & strSum3Qry & _
                " ) GROUP BY VODACCOUNTID "
                
    '#5432 Add+SO33.ShouldAmt要<Minus 才需要結清 By Kin 2009/12/10
    strRunSQL = " FROM " & GetOwner & "SO004 A," & GetOwner & "SO182 B ,(" & strSubQry & ") C, " & _
                "(" & strTotalQry & ") D " & _
                strWhere & " AND B.VODACCOUNTID=C.VODACCOUNTID " & _
                " AND B.VODACCOUNTID=D.VODACCOUNTID " & _
                " AND D.SUM2+SUM3-SUM1 < 0 "
    '最後所要挑選的欄位
    strFields = "SELECT A.CUSTID,A.DeclarantName,B.VodAccountid,B.FaciSeqNo," & _
                " B.NoCloseTime,C.MINUSTOTAL,C.ADDTOTAL,B.OrderNo,NVL(B.MinusCredit,0) MINUSCREDIT, " & _
                " NVL(B.AddCredit,0) ADDCREDITIT,B.RecTime,A.FACISNO,B.ROWID ID,C.MAXNOCLOSETIME,Nvl(D.SUM1,0) SUM1, " & _
                " NVL(D.SUM2,0) SUM2,NVL(D.SUM3,0) SUM3 "

'                " AND ( B.NOCLOSETIME > " & intPara36 & " OR C.TOTAL > " & intPara35 & ")"
    '********************************************************************************************************
    Select Case aType
        Case 1
            strRetSQL = strFields & ",0 EXETYPE " & strRunSQL & _
                        " AND ( Nvl(C.MAXNOCLOSETIME,0) >= " & intPara36 & " OR C.MINUSTOTAL >= " & intPara35 & ")"
        Case 2
            strRetSQL = strFields & ",1 EXETYPE " & strRunSQL & _
                        " AND ( Nvl(C.MAXNOCLOSETIME,0) < " & intPara36 & " AND C.MINUSTOTAL < " & intPara35 & ")"
        Case 3
            '#5521 計算費用原本只判斷SUM1-SUM2,此規格要多扣掉SUM3的金額 By Kin 2010/02/03
'            strRetSQL = strFields & ",0 EXETYPE " & strRunSQL & _
'                        " AND ( Nvl(C.MAXNOCLOSETIME,0) >= " & intPara36 & " OR ABS(Nvl(C.MINUSTOTAL,0)-Nvl(ADDTOTAL,0)-Nvl(SUM3,0)) >= " & intPara35 & ")" & _
'                        " UNION " & _
'                        strFields & ",1 EXETYPE " & strRunSQL & _
'                        " AND ( Nvl(C.MAXNOCLOSETIME,0) < " & intPara36 & " AND ABS(Nvl(C.MINUSTOTAL,0)-Nvl(C.ADDTOTAL,0)-Nvl(SUM3,0)) < " & intPara35 & ")"

            strRetSQL = strFields & ",0 EXETYPE " & strRunSQL & _
                        " AND ( Nvl(C.MAXNOCLOSETIME,0) >= " & intPara36 & " OR ABS(Nvl(SUM1,0)-Nvl(SUM2,0)-Nvl(SUM3,0)) >= " & intPara35 & ")" & _
                        " UNION " & _
                        strFields & ",1 EXETYPE " & strRunSQL & _
                        " AND ( Nvl(C.MAXNOCLOSETIME,0) < " & intPara36 & " AND ABS(Nvl(SUM1,0)-Nvl(SUM2,0)-Nvl(SUM3,0)) < " & intPara35 & ")"

    End Select
    GetSQL2 = strRetSQL
    Exit Function
ChkErr:
  Call ErrSub(Me.Name, "GetSQL2")
End Function


'抓取欲結算之VOD訂購資料
Private Function subChoose() As Boolean
  On Error GoTo ChkErr
    strChoose = Empty
    '************************畫面上的條件********************************************
    If gilCompCode.GetCodeNo <> "" Then
        Call subAnd2(strChoose, " A.CompCode = " & gilCompCode.GetCodeNo)
    End If
    If gilServiceType.GetCodeNo <> "" Then
        Call subAnd2(strChoose, " A.ServiceType='" & gilServiceType.GetCodeNo & "'")
    End If
    If txtCustID.Text <> "" Then
        Call subAnd2(strChoose, " A.CustId=" & txtCustID.Text)
    End If
    If txtSrNO.Tag <> "" Then
        Call subAnd2(strChoose, " A.SEQNO ='" & txtSrNO.Tag & "'")
    Else
        If txtSrNO.Text <> "" Then
            Call subAnd2(strChoose, "A.FaciSNO='" & txtSrNO.Text & "'")
        End If
    End If
'    If gdaEndDate.GetValue <> "" Then
'        Call subAnd2(strChoose, " To_Date(To_Char(B.RecTime,'YYYYMMDD'),'YYYYMMDD') <= " & _
'                        "To_Date('" & gdaEndDate.GetValue & "','YYYYMMDD')")
'    End If
   
    '************************************************************************************
     '*************************SO004的設備一定要參考號為3而且不能有拆機日期********************
    Call subAnd2(strChoose, "A.FaciCode In ( SELECT CODENO FROM " & GetOwner & "CD022 " & _
                        " WHERE NVL(StopFlag,0)=0 AND NVL(REFNO,0)= 3 )")
    If Not blnAutoEXE Then
        Call subAnd2(strChoose, "A.PRDATE IS NULL ")
    End If
    '*****************************************************************************************
    strSQL = GetSQL2(3, strChoose)

    strChooseString = "公司別　: " & subSpace(strCompCode) & ";" & _
                      "服務類別: " & subSpace(strServiceType) & ";" & _
                      "結算日期期限　: " & subSpace(strEndDate) & ";" & _
                      "PPV結算作業金額限額: " & intPara35 & ";" & _
                      "PPV結算作業月數限額: " & intPara36 & _
                      IIf(txtCustID.Text <> "", ";客戶編號: " & txtCustID.Text, "") & _
                      IIf(txtSrNO.Text <> "", ";STB設備流水號:" & txtSrNO.Text, "")
                      
    subChoose = True
  Exit Function
ChkErr:
  Call ErrSub(Me.Name, "subChoose")
End Function

Private Sub CreateTable()
  On Error Resume Next
    cnn.Execute "Drop Table SO1156"
  On Error GoTo ChkErr
    cnn.Execute "Create Table SO1156 (CompCode text(3),Custid Text(8),AuthenticId text(10),Amount Long,48CodeNo text(10))"
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "CreateTable")
End Sub

Private Function CreateBill(ByRef aHaveData As Boolean) As Boolean
  On Error GoTo ChkErr
    
    Dim rs182 As New ADODB.Recordset
    
    Dim strUpdSql As String
    Dim aVodAccountId As String
    Dim strBillNo As String
    Dim strUPD182 As String
    Dim strUPD033Vod As String
    aVodAccountId = Empty

    If Not GetRS(rs182, strSQL & " Order By VodAccountId,EXETYPE ", gcnGi, _
                adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
    If rs182.EOF Then
        aHaveData = False
        CreateBill = True
        strMsg = "無結算資料！"
        Exit Function
    End If
    strUpdSql = "UPDATE " & GetOwner & "SO182 Set NoCloseTime=Nvl(NoCloseTime,0)+1," & _
                "UpdEn='" & garyGi(1) & "',UPDTime='" & GetDTString(strRunTime) & "' "
    rs182.MoveFirst
    
     
    Do While Not rs182.EOF
        If Val(rs182("EXETYPE")) = 0 Then
            If aVodAccountId <> rs182("VODACCOUNTID") Then
                If Not Insert033(strBillNo, rs182) Then Exit Function
            End If
            '將有結算的資料CloseFlag UPD成1,並填入BillNO By Kin
            '如果BillNo已經有值的話就用原本的值 By Kin 2009/11/11
            '#5327測試不OK BillNo改成不要填了 By Kin 2009/11/18
            
            
            
            strUPD182 = "UPDATE " & GetOwner & "SO182 Set " & _
                        "CloseTime=To_Date('" & Format(strRunTime, "YYYYMMDDHHMMSS") & "','yyyymmddhh24miss')," & _
                        "CloseEN='" & garyGi(0) & "',CloseName='" & garyGi(1) & "'," & _
                        "UpdEn='" & garyGi(1) & "',UPDTime='" & GetDTString(strRunTime) & "', " & _
                        "CloseBillNo='" & strBillNo & "'," & _
                        "CloseItem=1,CloseFlag = 1" & _
                        " Where RowId = '" & rs182("ID") & "' AND NVL(AddCredit,0)= 0 " & _
                        " AND CLOSEBILLNO IS NULL "
    
            '以CreditSeqNo填入SO033VOD的CloseBillNo
            
            
            strUPD033Vod = " UPDATE " & GetOwner & "SO033VOD SET " & _
                            " CloseBillNo='" & strBillNo & "'," & _
                            " CloseItem=1" & _
                            " WHERE SEQNO=(SELECT CreditSeqNo FROM " & GetOwner & "SO182 " & _
                                            " WHERE ROWID='" & rs182("ID") & "' AND Nvl(ADDCREDIT,0)= 0 )" & _
                            " And CloseBillNo IS NULL "
                                            
            gcnGi.Execute (strUPD182)
            gcnGi.Execute (strUPD033Vod)
            
            
        Else
            '如果是未執行,要將次數加1
            Dim strQry As String
            strQry = strUpdSql & " WHERE ROWID='" & rs182("ID") & "'"
            gcnGi.Execute (strUpdSql & " WHERE ROWID='" & rs182("ID") & "' AND NVL(AddCredit,0)= 0")
        End If
        If blnNOUpdate Or blnAutoEXE Then
            If Not InsCloneRs(rs182("ID") & "", "SO182", strBillNo, rs182Clone) Then Exit Function
            If Not InsCloneRs(rs182("ID") & "", "SO033VOD", strBillNo, rs033VodClone) Then Exit Function
        End If
        aVodAccountId = rs182("VODAccountID")
        rs182.MoveNext
    Loop
  
    CreateBill = True
    On Error Resume Next
    aHaveData = True
    Call CloseRecordset(rs182)
    Exit Function
ChkErr:
   gcnGi.RollbackTrans
  Call ErrSub(Me.Name, "CreateBill")
End Function
Private Function InsCloneRs(ByVal aRowId As String, ByVal aTableName As String, ByVal strBillNo As String, ByRef rsClone As ADODB.Recordset) As Boolean
  On Error GoTo ChkErr
    Dim strQry As String
    Dim rsTmp As New ADODB.Recordset
    Dim i As Long
    '#5521 測試不OK,不能直接用Where條件取資料,因為之前已經Upd過了,所以改用strBillNo就一定是對的 By Kin 2010/02/04
    If UCase(aTableName) = "SO033VOD" Then
        '#5521 多增加 CloseBillNo Is Null的判斷 By Kin 2010/02/03
        strQry = "SELECT A.ROWID,A.* FROM " & GetOwner & aTableName & " A " & _
                " WHERE A.SEQNO IN (SELECT CREDITSEQNO FROM " & GetOwner & "SO182 " & _
                                    " WHERE ROWID='" & aRowId & "'" & _
                                    " AND CLOSEBILLNO='" & strBillNo & "')"
    Else
        
        strQry = "SELECT A.ROWID,A.* FROM " & GetOwner & aTableName & " A WHERE ROWID='" & aRowId & "'" & _
                " AND A.CLOSEBILLNO ='" & strBillNo & "'"
    End If
    If Not GetRS(rsTmp, strQry, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
    If Not rsTmp.EOF Then
        rsTmp.MoveFirst
        Do While Not rsTmp.EOF
            If rsClone.RecordCount = 0 Then
                rsClone.AddNew
            Else
                rsClone.Find "RowId='" & rsTmp("RowId") & "" & "'", , adSearchForward, 1
                If rsClone.EOF Then
                    rsClone.AddNew
                Else
                    GoTo lNext
                End If
            End If
            For i = 0 To rsTmp.Fields.Count - 1
                If Len(rsTmp.Fields(i).Value & "") > 0 Then
                    rsClone.Fields(rsTmp.Fields(i).Name) = rsTmp.Fields(i).Value & ""
                End If
            Next i
            rsClone.Update
lNext:
            rsTmp.MoveNext
        Loop
    End If
    On Error Resume Next
    Call CloseRecordset(rsTmp)
    InsCloneRs = True
    Exit Function
ChkErr:
  Call ErrSub(Me.Name, "InsCloneRs")
End Function

Private Function Insert033(ByRef aBillNo As String, ByRef rsSO182 As Recordset) As Boolean
  On Error GoTo ChkErr
    Dim rs033 As New ADODB.Recordset
    Dim i As Long
    If Not GetRS(rs033, "SELECT A.RowId,A.* FROM " & GetOwner & "SO033 A WHERE 1=0 ", gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
    If Not Get014Data(rsSO182("CustId")) Then Exit Function
    With rs033
        .AddNew
        .Fields("CustId") = rsSO182("CustID")
        .Fields("CompCode") = gilCompCode.GetCodeNo
        aBillNo = GetInvoiceNo("T", strServiceType)
        .Fields("BillNo") = aBillNo
        .Fields("Item") = 1
        .Fields("CitemCode").Value = strCitemCode
        .Fields("CitemName").Value = strCitemName
        '#5488 如果畫面上有下應收日期,則以畫面條件為主
        If gdaShouldDate.GetValue & "" <> "" Then
            .Fields("ShouldDate").Value = gdaShouldDate.GetValue(True)
        Else
            .Fields("ShouldDate").Value = GetADdate(strRunTime)
        End If
        '#5521 產出的費用原本為SUM1-SUM2,改為SUM1-(SUM2+SUM3) By Kin 2010/02/03
'        .Fields("OldAmt").Value = Abs(Val(rsSO182.Fields("MINUSTotal") & "") - Val(rsSO182.Fields("ADDTOTAL") & "") - Val(rsSO182.Fields("SUM3") & ""))
'        .Fields("ShouldAmt").Value = Abs(Val(rsSO182.Fields("MINUSTotal") & "") - Val(rsSO182.Fields("ADDTOTAL") & "") - Val(rsSO182.Fields("SUM3") & ""))
        
        .Fields("OldAmt").Value = Abs(Val(rsSO182.Fields("SUM1") & "") - Val(rsSO182.Fields("SUM2") & "") - Val(rsSO182.Fields("SUM3") & ""))
        .Fields("ShouldAmt").Value = Abs(Val(rsSO182.Fields("SUM1") & "") - Val(rsSO182.Fields("SUM2") & "") - Val(rsSO182.Fields("SUM3") & ""))
        .Fields("OldPeriod").Value = 0
        .Fields("RealPeriod").Value = 0
        .Fields("CMCODE").Value = IIf(strCMCode = "", Null, strCMCode)
        .Fields("CMName").Value = IIf(strCMName = "", Null, strCMName)
        .Fields("UCCODE").Value = IIf(strUCCode = "", Null, strUCCode)
        .Fields("UCName").Value = IIf(strUCName = "", Null, strUCName)
        .Fields("PTCODE").Value = 1
        .Fields("PTName").Value = "現金"
        .Fields("ClassCode") = rs014("ClassCode1") & ""
        .Fields("FaciSeqNo").Value = rsSO182("FaciSeqNo") & ""
        .Fields("FaciSNO").Value = rsSO182("FaciSNO") & ""
        .Fields("ServiceType").Value = gilServiceType.GetCodeNo
        .Fields("Note").Value = "VOD費用結算產生的收費資料"
        .Fields("CreateTime").Value = strRunTime
        .Fields("UpdTime").Value = GetDTString(strRunTime)
        .Fields("CreateEn").Value = garyGi(0)
        .Fields("UpdEn").Value = garyGi(1)
        .Fields("AddrNo").Value = rs014("AddrNo") & ""
        .Fields("StrtCode").Value = rs014("StrtCode") & ""
        .Fields("MduId").Value = rs014("MduId") & ""
        .Fields("ServCode").Value = rs014("ServCode") & ""
        .Fields("ClctAreaCode").Value = rs014("ClctAreaCode") & ""
        .Fields("OldClctEn").Value = rs014("ClctEn") & ""
        .Fields("OldClctName").Value = rs014("ClctName") & ""
        .Fields("AreaCode").Value = rs014("AreaCode") & ""
        .Fields("ClctEn").Value = rs014("ClctEn") & ""
        .Fields("ClctName").Value = rs014("ClctName") & ""
        .Fields("SalePointCode").Value = IIf(strSalePointCode = "", Null, strSalePointCode)
        .Fields("SalePointName").Value = IIf(strSalePointName = "", Null, strSalePointName)
        '#5488 增加結算截止日 By Kin 2010/01/21
        .Fields("CloseStopDate").Value = gdaEndDate.GetValue(True)
        .Update
    End With
    If blnNOUpdate Or blnAutoEXE Then
        rs033Clone.AddNew
        For i = 0 To rs033.Fields.Count - 1
            If Len(rs033.Fields(i).Value & "") > 0 Then
                rs033Clone.Fields(rs033.Fields(i).Name).Value = rs033.Fields(i).Value & ""
            End If
        Next i
        rs033Clone.Update
    End If
    Insert033 = True
    On Error Resume Next
    Call CloseRecordset(rs033)
    Exit Function
ChkErr:
  Call ErrSub(Me.Name, "Insert033")
End Function
Private Function Get014Data(ByVal aCustId As String) As Boolean
  On Error GoTo ChkErr
    Dim strQry As String
    strQry = "SELECT A.*,B.CLASSCODE1 FROM " & GetOwner & "SO014 A," & GetOwner & "SO001 B " & _
            " WHERE B.CUSTID=" & aCustId & " AND " & IIf(intPara14 = 1, "B.ChargeAddrNo", "B.InstAddrNo") & "=A.AddrNo" & _
            " AND B.COMPCODE=" & gilCompCode.GetCodeNo & " AND A.COMPCODE=B.COMPCODE"
    If Not GetRS(rs014, strQry, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
    Get014Data = True
    Exit Function
ChkErr:
  Call ErrSub(Me.Name, "Get014Data")
End Function






Public Sub subAnd(strCH As String)
  On Error GoTo ChkErr
    If strChoose <> "" Then
        strChoose = strChoose & " And " & strCH
    Else
        strChoose = strCH
    End If
  Exit Sub
ChkErr:
    Call ErrSub("Utility5", "subAnd")
End Sub

Private Sub Form_Load()
  On Error GoTo ChkErr
    strServiceType = "D"
    Call SubGil
    Call DefaultVal
    gilServiceType.Query_Description
    If blnNOUpdate Or blnAutoEXE Then
        If Not DefineRs(rs033Clone, "SO033") Then Exit Sub
        If Not DefineRs(rs033VodClone, "SO033Vod") Then Exit Sub
        If Not DefineRs(rs182Clone, "SO182") Then Exit Sub
    End If
    
    'If blnAutoEXE And strEndDate <> "" Then cmdOk.Value = True
    
  Exit Sub
ChkErr:
    ErrSub Me.Name, "Form_Load"
End Sub
Private Sub Form_Activate()
    On Error Resume Next
        Screen.MousePointer = vbDefault
        If blnAutoEXE And strEndDate <> "" Then cmdOK.Value = True
End Sub
Private Sub SubGil()
  On Error GoTo ChkErr
    SetgiList gilServiceType, "CodeNo", "Description", "CD046"
    SetgiList gilCompCode, "CodeNo", "Description", "CD039", , , , , , , GetCompCodeFilter(gilCompCode)
    If strCompCode = "" Then gilCompCode.SetCodeNo garyGi(9) Else gilCompCode.SetCodeNo (strCompCode)
    gilCompCode.Query_Description
    If strServiceType = "" Then Call SelectServType(gilCompCode.GetCodeNo, gilServiceType) Else gilServiceType.SetCodeNo (strServiceType)
     gilServiceType.GetDescription
  Exit Sub
ChkErr:
    ErrSub Me.Name, "subGil"
End Sub
Private Sub DefaultVal()
  On Error GoTo ChkErr
    Dim rsSO062 As New ADODB.Recordset
    Dim rsSO043 As New ADODB.Recordset  ' 收費參數
    Dim rsCD019 As New ADODB.Recordset
    Dim rs044 As New ADODB.Recordset
    Dim rs107 As New ADODB.Recordset
    Dim strSQL As String, strSO043 As String
    Dim strCD019 As String, strSO044 As String
    Dim strCD017 As String
    'If strServiceType <> "" Then gilServiceType.SetCodeNo strServiceType
    
    strSQL = "SELECT * FROM " & GetOwner & "SO062 WHERE TYPE = 4 And CompCode=" & gilCompCode.GetCodeNo & " And ServiceType='" & gilServiceType.GetCodeNo & "'"
    If OpenRecordset(rsSO062, strSQL, gcnGi, adOpenForwardOnly, adLockReadOnly, adUseClient, False) <> giOK Then
        ErrHandle "開啟 [最近日結參數記錄檔](SO062) 時發生錯誤: " & Err.Description & " 這個程式即將關閉。", False, 2, "Form_Load"
        Unload Me
        Exit Sub
    End If
    
    '***************************************************************************************
    '收費參數資料
    strSO043 = "Select para35,para36,Para14 from " & GetOwner & "SO043 Where CompCode=" & gilCompCode.GetCodeNo & " And ServiceType='" & gilServiceType.GetCodeNo & "'"
    If Not GetRS(rsSO043, strSO043, gcnGi, , adOpenForwardOnly, adLockReadOnly, adUseClient, False) Then Exit Sub
    
    If Not rsSO043.EOF Then
        intPara14 = Val(rsSO043("Para14") & "")
    Else
        intPara14 = 0
    End If
    '***************************************************************************************
    
    '***************************************************************************
    '新增至SO033的收費項目
    strCD019 = "SELECT CodeNo,Description FROM " & GetOwner & "CD019 " & _
                " WHERE NVL(STOPFLAG,0)=0 AND REFNO = 21 " & _
                " AND SERVICETYPE ='" & gilServiceType.GetCodeNo & "'" & _
                " AND SIGN = '+'"
    If Not GetRS(rsCD019, strCD019, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
    If Not rsCD019.EOF Then
        strCitemCode = rsCD019("CODENO") & ""
        strCitemName = rsCD019("DESCRIPTION") & ""
    Else
        strCitemCode = ""
        strCitemName = ""
    End If
    '*****************************************************************************
    
    '***********************SO044的預設值****************************************
    strSO044 = "SELECT A.CMCode,B.Description CMNAME FROM " & GetOwner & "SO044 A," & GetOwner & "CD031 B " & _
                " WHERE A.CMCODE = B.CODENO AND A.SERVICETYPE='" & gilServiceType.GetCodeNo & "'"
    If Not GetRS(rs044, strSO044, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
    If Not rs044.EOF Then
        strCMCode = rs044("CMCODE") & ""
        strCMName = rs044("CMNAME") & ""
    Else
        strCMCode = "": strCMName = ""
    End If
    
    strSO044 = "SELECT A.UCCODE,B.Description UCNAME FROM " & GetOwner & "SO044 A," & GetOwner & "CD013 B " & _
                " WHERE A.UCCODE = B.CODENO AND A.SERVICETYPE='" & gilServiceType.GetCodeNo & "'"
    If Not GetRS(rs044, strSO044, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
    If Not rs044.EOF Then
        strUCCode = rs044("UCCODE") & ""
        strUCName = rs044("UCNAME") & ""
    Else
        strUCCode = "": strUCName = ""
    End If
    strPTCode = "1"
    strPTName = "現金"
    
    '****************************************************************************
    '****************************CD107的預設值***********************************
    strCD017 = "SELECT CodeNo,Description FROM " & GetOwner & "CD107 " & _
            " WHERE NVL(STOPFLAG,0)=0 AND DefaultValue=1 AND COMPCODE=" & gilCompCode.GetCodeNo
    
    If Not GetRS(rs107, strCD017, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
    If Not rs107.EOF Then
        strSalePointCode = rs107("CODENO") & ""
        strSalePointName = rs107("DESCRIPTION") & ""
    Else
        strSalePointCode = ""
        strSalePointName = ""
    End If
    
    '****************************************************************************
    
    
    ' 若SO062無資料 或無傳入值
    If rsSO062.EOF Then
        ' 則各欄位為空白或初值
        lblTranDate.Caption = ""
        lblUpdTime.Caption = ""
        lblUpdEn.Caption = ""
        lblPara35.Caption = "0"
        lblPara36.Caption = "0"
        gdaEndDate.SetValue ""
        If Not rsSO043.EOF Then
            intPara35 = rsSO043("Para35").Value & ""
            lblPara35.Caption = rsSO043("Para35").Value & ""
            intPara36 = rsSO043("Para36").Value & ""
            lblPara36.Caption = rsSO043("Para36").Value & ""
        End If
    ' 若有資料
    Else
        '結算截止日
'        lblTranDate.Caption = Format(rsSO062("TranDate").Value & "", "ee/MM/DD")
        lblTranDate.Caption = GetDTString(rsSO062("TranDate").Value & "")
        lngTranDate = Val(rsSO062("TranDate").Value & "")
        '上次執行結算時間
        lblUpdTime.Caption = rsSO062("UpdTime").Value & ""
        '上次執行結算人員
        lblUpdEn.Caption = rsSO062("UpdEn").Value & ""
        'PPV結算作業金額限額 PPV結算作業月數限額
        intPara35 = rsSO043("Para35").Value & ""
        lblPara35.Caption = rsSO043("Para35").Value & ""
        intPara36 = rsSO043("Para36").Value & ""
        lblPara36.Caption = rsSO043("Para36").Value & ""
        '結算日期為上次結算日期加1天
        '#5343 要把預設值拿掉 By Kin 2009/11/27
        'gdaEndDate.SetValue DateAdd("d", 1, rsSO062("TranDate").Value & "")
        
    End If
    
    '判斷是否由別的Form進來的
    txtCustID.Text = strCustid
    If strCustid <> "" Then txtCustID.Enabled = False
    txtSrNO.Text = strSrNO
    txtXLSFile.Visible = False
    cmdOpen.Visible = False
    lblXLSPath.Visible = False
    If blnGive Then
        txtCustID.Enabled = False
        txtSrNO.Enabled = False
        cmdSTB.Enabled = False
        txtXLSFile.Visible = True
        lblXLSPath.Visible = True
        cmdOpen.Visible = True
        If strEndDate <> "" Then gdaEndDate.SetValue strEndDate & ""
        If strEndDate <> "" Then gdaEndDate.Enabled = False
        If strPara35In <> "" Then lblPara35.Caption = strPara35In Else lblPara35.Caption = intPara35
        If strPara36In <> "" Then lblPara36.Caption = strPara36In Else lblPara36.Caption = intPara36
        chkTran1.Value = 1
        chkTran2.Value = 1
        chkRun.Value = 1
        If strCompCode <> "" Then gilCompCode.Enabled = False
        If strServiceType <> "" Then gilServiceType.Enabled = False
    End If
    If blnAutoEXE Then
        chkTran1.Value = 0
        chkTran2.Value = 0
        chkRun.Value = 1
    End If
    If strFaciSNo <> "" Then
        txtSrNO.Text = strFaciSNo
    End If
    If strFaciSeqNo <> "" Then
        txtSrNO.Tag = strFaciSeqNo
    End If
    If strEndDate <> "" Then gdaEndDate.SetValue strEndDate & ""
    If strEndDate <> "" Then gdaEndDate.Enabled = False
    If strPara35In <> "" Then
        intPara35 = Val(strPara35In)
        lblPara35.Caption = intPara35
    End If
    If strPara36In <> "" Then
        intPara36 = Val(strPara36In)
        lblPara36.Caption = intPara36
    End If
    '#5343 測試不OK，執行結算要有權限控制 By Kin 2009/11/25
    If Not blnAutoEXE Then
        chkRun.Enabled = GetUserPriv("SO32E02", "(SO32E021)")
    End If
    If Not chkRun.Enabled Then
        chkRun.Value = 0
    End If
    On Error Resume Next
    Call CloseRecordset(rsSO062)
    Call CloseRecordset(rsCD019)
    Call CloseRecordset(rs107)
    Exit Sub
ChkErr:
   Resume 0
   ErrSub Me.Name, "DefaultVal"
  
End Sub

Private Sub gilServiceType_Change()
  On Error GoTo ChkErr
    DefaultVal
  Exit Sub
ChkErr:
   ErrSub Me.Name, "gilServiceType_Change"
End Sub

Private Sub gilCompCode_Change()
  On Error GoTo ChkErr
    Dim strSQL As String

    If Len(gilCompCode.GetCodeNo & "") = 0 Then Exit Sub
    ' 公司別：若有值且離開此欄位時
    If rsCD039.State = adStateOpen Then rsCD039.Close
    strSQL = "SELECT VouPath,CODENO,EmcCompCode,AccountVer FROM " & GetOwner & "CD039 Where CODENO = " & gilCompCode.GetCodeNo & " ORDER BY CODENO"
    If OpenRecordset(rsCD039, strSQL, gcnGi, adOpenStatic, adLockReadOnly, adUseClient, False) <> giOK Then
        ErrHandle "在開啟 [公司別代碼檔](CD039) 時發生錯誤: " & Err.Description & " 這個程式即將關閉!", False, 2, "gilCompCode_Validate"
        End
        Exit Sub
    End If
     Call SelectServType(gilCompCode.GetCodeNo, gilServiceType)
     DefaultVal
  Exit Sub
ChkErr:
    ErrSub Me.Name, "gilCompCode_Change"
End Sub

Private Sub gilCompCode_Validate(Cancel As Boolean)
  On Error GoTo ChkErr
    If gilCompCode.GetCodeNo = "" Then MsgMustBe ("公司別"): Cancel = True
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "gilCompCode_Validate")
End Sub
Private Function DefineRs(ByRef rsClone As Recordset, ByVal sTableName As String) As Boolean
  On Error GoTo ChkErr
    Dim strQry As String
    Dim i As Long
    Dim rsTmp As New ADODB.Recordset
    
    strQry = "SELECT A.ROWID,A.* FROM " & GetOwner & sTableName & " A WHERE 1=0"
    If Not GetRS(rsTmp, strQry, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
    With rsClone
        If .State = adStateOpen Then .Close
        .CursorLocation = adUseClient
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        For i = 0 To rsTmp.Fields.Count - 1
            .Fields.Append rsTmp.Fields(i).Name, adBSTR, rsTmp.Fields(i).DefinedSize, adFldIsNullable
        Next i
    End With
    rsClone.Open
    
    DefineRs = True
    On Error Resume Next
    Call CloseRecordset(rsTmp)
    Exit Function
ChkErr:
  ErrSub Me.Name, "DefineRs"
End Function
Public Property Let uFromBat(ByVal vData As Boolean)
    blnGive = vData
End Property
Public Property Let uCompCode(ByVal vData As String)
    strCompCode = vData
    'blnGive = True
End Property

Public Property Let uServiceType(ByVal vData As String)
    strServiceType = vData
    strServiceType = "D"    '#5327 測試不OK,要固定D服務
    'gilServiceType.SetCodeNo vData
    'blnGive = True
End Property

Public Property Let uCustId(ByVal vData As String)
    strCustid = vData
    'txtCustid.Text = vData
    'blnGive = True
End Property

'STB設備序號
Public Property Let uSrNO(ByVal vData As String)
    strSrNO = vData
    txtSrNO.Text = vData
    'txtSrNO.Text = vData
End Property
Public Property Let uFaciSeqNo(ByVal vData As String)
  On Error Resume Next
    strFaciSeqNo = vData
    'txtSrNO.Tag = strFaciSeqNo
End Property
Public Property Let uFaciSNo(ByVal vData As String)
  On Error Resume Next
    strFaciSNo = vData
    'txtSrNO.Text = strFaciSNO
End Property

'結算日期期限
Public Property Let uEndDate(ByVal vData As String)
    strEndDate = vData
End Property

'PPV結算作業金額限額
Public Property Let uPara35(ByVal vData As Integer)
    strPara35In = vData
End Property

'PPV結算作業月數限額
Public Property Let uPara36(ByVal vData As Integer)
    strPara36In = vData
End Property
'回傳執行的Msg
Public Property Get uMsg() As Variant
    uMsg = strMsg
End Property
'回傳所有新增SO033的資料
Public Property Get uRs033() As ADODB.Recordset
  On Error Resume Next
    Set uRs033 = rs033Clone
End Property
'回傳所有更新SO033VOD的資料
Public Property Get uRs033Vod() As ADODB.Recordset
  On Error Resume Next
    Set uRs033Vod = rs033VodClone
End Property
'回傳所有更新SO182的資料
Public Property Get uRs182() As ADODB.Recordset
  On Error Resume Next
    Set uRs182 = rs182Clone
End Property
Public Property Set uRs182(ByRef rs As ADODB.Recordset)
    Set rs182Clone = rs
End Property
Public Property Set uRs033(ByRef rs As ADODB.Recordset)
    Set rs033Clone = rs
End Property
Public Property Set uRs033Vod(ByRef rs As ADODB.Recordset)
    Set rs033VodClone = rs
End Property

'由Miggie結清過來的,不允許任何資料Update,到最後要RollbackTrans
Public Property Let uNoUpdate(ByVal vData As Boolean)
  On Error Resume Next
    blnNOUpdate = vData
End Property
'是否正常執行完程式
Public Property Get uDo() As Boolean
    uDo = blnDO
End Property
'是否有執行的資料
Public Property Get uHaveData() As Boolean
  On Error Resume Next
    uHaveData = blnNoData
End Property
Public Property Let uAutoExe(ByVal vData As Boolean)
  On Error Resume Next
    blnAutoEXE = vData
End Property

Public Property Let uTrans(ByVal vData As Boolean)
  On Error Resume Next
    blnTrans = vData
End Property
Private Sub txtSrNO_KeyPress(keyAscii As Integer)
  On Error Resume Next
    txtSrNO.Tag = ""
End Sub

