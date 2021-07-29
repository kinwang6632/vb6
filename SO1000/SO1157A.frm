VERSION 5.00
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Begin VB.Form frmSO1157A 
   Caption         =   "VOD結算作業[SO1157A]"
   ClientHeight    =   6360
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5190
   Icon            =   "SO1157A.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   5190
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton cmdSTB 
      Caption         =   "設備"
      Height          =   315
      Left            =   3570
      TabIndex        =   29
      Top             =   3570
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
      Top             =   3570
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
      Top             =   3240
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
      Left            =   60
      TabIndex        =   5
      Top             =   5790
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
      TabIndex        =   6
      Top             =   5775
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
      TabIndex        =   7
      Top             =   5775
      Width           =   1425
   End
   Begin VB.Frame Frame1 
      Height          =   1305
      Left            =   240
      TabIndex        =   15
      Top             =   4320
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
         TabIndex        =   4
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
         TabIndex        =   3
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
         TabIndex        =   2
         Top             =   120
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
      TabIndex        =   8
      Top             =   120
      Width           =   4950
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         Caption         =   "結算截止日:"
         Height          =   180
         Left            =   825
         TabIndex        =   14
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
         TabIndex        =   13
         Top             =   285
         Width           =   1050
      End
      Begin VB.Label lbl2 
         AutoSize        =   -1  'True
         Caption         =   "上次執行結算時間:"
         Height          =   180
         Left            =   240
         TabIndex        =   12
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
         TabIndex        =   11
         Top             =   645
         Width           =   1980
      End
      Begin VB.Label lbl3 
         AutoSize        =   -1  'True
         Caption         =   "上次執行結算人員: "
         Height          =   180
         Left            =   240
         TabIndex        =   10
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
         TabIndex        =   9
         Top             =   975
         Width           =   2760
      End
   End
   Begin prjGiList.GiList gilCompCode 
      Height          =   300
      Left            =   1920
      TabIndex        =   26
      Top             =   1590
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
      TabIndex        =   27
      Top             =   2010
      Width           =   3090
      _ExtentX        =   5450
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
      FldWidth1       =   900
      FldWidth2       =   1855
      F2Corresponding =   -1  'True
   End
   Begin Gi_Date.GiDate gdaEndDate 
      Height          =   270
      Left            =   2160
      TabIndex        =   28
      Top             =   3900
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
      TabIndex        =   25
      Top             =   3900
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
      TabIndex        =   24
      Top             =   3900
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
      TabIndex        =   23
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label lblPara36 
      BackColor       =   &H00E0E0E0&
      Caption         =   "XXXXXXXXX"
      Height          =   255
      Left            =   2160
      TabIndex        =   22
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label lblPara35 
      BackColor       =   &H00E0E0E0&
      Caption         =   "XXXXXXXXX"
      Height          =   195
      Left            =   2160
      TabIndex        =   21
      Top             =   2520
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
      TabIndex        =   20
      Top             =   2880
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
      TabIndex        =   19
      Top             =   2520
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
      TabIndex        =   18
      Top             =   1605
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
      TabIndex        =   17
      Top             =   3570
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
      TabIndex        =   16
      Top             =   2070
      Width           =   780
   End
End
Attribute VB_Name = "frmSO1157A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'#5327 需求 By Kin 2009/11/06
Option Explicit
Dim lngTranDate As Long             ' 結算截止日
Dim rsCD039 As New ADODB.Recordset  ' 公司別代碼檔
Dim intPara35 As Integer, intPara36 As Integer
Dim intPara35In As Integer, intPara36In As Integer
Dim strSQL As String
Dim strChooseString As String
Dim strFormula As String
Dim strCompCode As String
Dim strServiceType As String
Dim strCustId As String
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
Private strCMname As String
Private strUCCode As String
Private strUCName As String
Private strPTCode As String
Private strPTName As String
Private intPara14 As Integer
Private strSalePointCode As String
Private strSalePointName As String
Private rs014 As New ADODB.Recordset

Private Sub cmdCancel_Click()
  On Error Resume Next
    Unload Me
End Sub

Private Sub cmdPrevRpt_Click()
  On Error GoTo chkErr
    Screen.MousePointer = vbHourglass
      Call PreviousRpt(GetPrinterName(5), RptName("SO1157"), Me.Caption)
    Screen.MousePointer = vbDefault
  Exit Sub
chkErr:
  Call ErrSub(Me.Name, "cmdPrevRpt_Click")
End Sub

Private Sub cmdSTB_Click()
  On Error GoTo chkErr
    If txtCustId.Text = "" Then
        MsgBox "請輸入客戶編號再做選取設備動作！", vbInformation, "訊息"
        Exit Sub
    End If
    With frmSO1131E
        .uParentForm = Me
        .uEditMode = giEditModeEdit
        .uCustId = txtCustId.Text
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
chkErr:
    ErrSub Name, "cmdSTB_Click"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error Resume Next
    If KeyCode = vbKeyF2 Then cmdOk.Value = True
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
    strServiceType = "": strCompCode = "": strEndDate = "": strCustId = "": strSrNO = ""
    blnGive = False
    SendSQL , , True
    Call ReleaseCOM(Me)
End Sub

Private Sub cmdOK_Click()
  On Error GoTo chkErr
    'If gdaEndDate.GetValue = "" Then MsgBox "請填入結算日期期限!", vbInformation, "訊息!": Exit Sub
    If Not IsDataOK Then Exit Sub
    strServiceType = gilServiceType.GetCodeNo
    strCompCode = gilCompCode.GetCodeNo
    strEndDate = gdaEndDate.GetValue
    strCustId = txtCustId
    strSrNO = txtSrNO
    strRunTime = RightNow
    If Not subChoose Then Exit Sub
    cmdCancel.SetFocus
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
        Dim blnNoData As Boolean
        If Not blnTrans Then gcnGi.BeginTrans
        If Not CreateBill(blnNoData) Then
            If Not blnTrans Then gcnGi.RollbackTrans
            GoTo 88
        End If
        If Not NewSO062(blnNoData, strMsg) Then
            If Not blnTrans Then gcnGi.RollbackTrans
            GoTo 88
        End If
        If Not blnTrans Then gcnGi.CommitTrans
    End If
    If blnGive = True And chkRun.Value = 1 Then
        Screen.MousePointer = vbDefault
        MsgBox strMsg, vbInformation, "訊息"
        Unload Me
    Else
        If Not blnGive And chkRun.Value = 1 Then MsgBox strMsg, vbInformation, "訊息"
    End If
88:
    Me.Enabled = True
    Screen.MousePointer = vbDefault
    Exit Sub
chkErr:
  Call ErrSub(Me.Name, "cmdOk_Click")
End Sub
Private Sub subPrint(ByVal aEXETYPE As Integer)
  On Error GoTo chkErr
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
        Call PrintRpt(GetPrinterName(5), "SO1157A.rpt", , "預計結算明細報表 [SO1157A]", strPrintSQL, strChooseString, , True, , , strFormula)
        SendSQL , , True
    End If
    Me.Enabled = True
    Screen.MousePointer = vbDefault
    CloseRecordset rsTmp
  Exit Sub
chkErr:
  Call ErrSub(Me.Name, "subPrint")
End Sub
'aEXETYPE = 0 代表預計執行 aEXETYPE =1 代表未執行
Private Function subCreateView(ByVal aEXETYPE As Integer) As Boolean
    Dim strView As String
    On Error Resume Next
      gcnGi.Execute "Drop View " & strViewName
    On Error GoTo chkErr
    Dim strCrtSQL As String
    strViewName = GetTmpViewName
    strCrtSQL = "SELECT DISTINCT CUSTID,VODACCOUNTID," & _
                "FACISNO,SUM(MINUSCREDIT) MINUSCREDIT FROM (" & _
                strSQL & ") WHERE EXETYPE=" & aEXETYPE & _
                " GROUP BY CUSTID,VODACCOUNTID,FACISNO"
'    strSQL = "SELECT * FROM (" & strSQL & ") WHERE EXETYPE=" & aEXETYPE
    strView = "Create View " & strViewName & " as ( " & strCrtSQL & ")"
    gcnGi.Execute strView
    SendSQL strView, True
    Exit Function
chkErr:
  Call ErrSub(Me.Name, "subCreateView")
End Function
Private Function IsDataOK() As Boolean
  On Error GoTo chkErr
    If gdaEndDate.GetValue & "" = "" Then
        MsgBox "請填入結算日期期限!", vbInformation, "訊息!"
        Exit Function
    End If
    If chkTran1.Value = 1 Or chkTran2.Value = 1 Or chkRun.Value = 1 Then
        IsDataOK = True
    Else
        MsgBox "請選擇要執行的項目！", vbInformation, "訊息"
        Exit Function
    End If
    If gilCompCode.GetCodeNo = "" Then
        MsgBox "請輸入公司別！", vbInformation, "訊息"
        Exit Function
    End If
    IsDataOK = True
    Exit Function
chkErr:
  Call ErrSub(Me.Name, "IsDataOK")
End Function
Private Function NewSO062(ByVal aHaveData As Boolean, Optional ByRef aMsg As String) As Boolean
On Error GoTo chkErr
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
            " ; 客編=" & txtCustId.Text & "; STB設備序號=" & txtSrNO.Text & ""
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
    blnDO = True

    NewSO062 = True
Exit Function
chkErr:
    Call ErrSub(Me.Name, "NewSO062")
End Function
Private Function GetSQL2(ByVal aType As Integer, ByVal aWhere As String) As String
    On Error GoTo chkErr
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
    
    '拆除,SO004要過濾拆除的資料
    'str004Where = "(A.INSTDATE IS NOT NULL AND A.PRDATE IS NOT NULL AND A.PRDATE > A.INSTDATE)"

    
    
    '****************************期間內總訂購點數的 Where 條件 *******************************
'    strWhere = " WHERE A.CompCode = B.CompCode And A.ServiceType = B.ServiceType " & _
'                " And CREDITTYPECODE In (SELECT CODENO FROM " & GetOwner & "CD108 " & _
'                " WHERE CD108.REFNO = 3 AND NVL(CD108.STOPFLAG,0) = 0 AND B.SERVICETYPE=CD108.SERVICETYPE)" & _
'                " AND NVL(B.CloseFlag,0)=0 AND NVL(B.STOPFLAG,0)= 0 AND A.SEQNO = B.FACISEQNO "
                
    strWhere = " WHERE A.CompCode = B.CompCode And A.ServiceType = B.ServiceType " & _
                " AND NVL(B.CloseFlag,0)=0 AND NVL(B.STOPFLAG,0)= 0 AND A.SEQNO = B.FACISEQNO "
    '******************************************************************************************
    
    '****************************可使用點數的 Where 條件 *****************************************
'    strWhere2 = " WHERE A.CompCode = B.CompCode And A.ServiceType = B.ServiceType " & _
'               " And CREDITTYPECODE In (SELECT CODENO FROM " & GetOwner & "CD108 " & _
'               " WHERE CD108.REFNO IN(1,2) AND NVL(CD108.STOPFLAG,0) = 0 AND B.SERVICETYPE=CD108.SERVICETYPE)" & _
'               " AND NVL(B.STOPFLAG,0)=0 AND A.SEQNO=B.FACISEQNO "
               
    strWhere2 = " WHERE A.CompCode = B.CompCode And A.ServiceType = B.ServiceType " & _
               " AND NVL(B.STOPFLAG,0)=0 AND A.SEQNO=B.FACISEQNO "
    '**********************************************************************************************
    
    '****************************已結算過期間內總訂購點數的 Where 條件 *******************************
    '這一段語法不要了,所以加一個1=0
    strWhere3 = " WHERE A.CompCode = B.CompCode And A.ServiceType = B.ServiceType " & _
                " And CREDITTYPECODE In (SELECT CODENO FROM " & GetOwner & "CD108 " & _
                " WHERE CD108.REFNO = 3 AND NVL(CD108.STOPFLAG,0) = 0 AND B.SERVICETYPE=CD108.SERVICETYPE)" & _
                " AND NVL(B.CloseFlag,0)=1 AND NVL(B.STOPFLAG,0)=0 AND 1=0 "
    '**************************************************************************************************
    If aWhere <> Empty Then
        strWhere = strWhere & " AND " & aWhere
        strWhere2 = strWhere2 & " AND " & aWhere
        strWhere3 = strWhere3 & " AND " & aWhere
    End If
    '****************************要結算的SQL語法*****************************************************
    
    '****************************算出期間內總訂購數之後還要判斷合計或未結算次數的條件*****************
    strSubQry = "SELECT B.VODACCOUNTID,SUM(B.MINUSCREDIT) TOTAL,Max(Nvl(B.NOCLOSETIME,0)) MAXNOCLOSETIME " & _
                    " FROM " & GetOwner & "SO004 A," & GetOwner & "SO182 B" & strWhere & _
                    " GROUP BY B.VODACCOUNTID "
    '**************************************************************************************
    strSum1Qry = "SELECT B.VODACCOUNTID,B.NOCLOSETIME,C.TOTAL,B.MINUSCREDIT " & _
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
    '算出已結算過期間內總訂購點數
    strSum3Qry = "SELECT B.VODACCOUNTID,0 SUM1,0 SUM2,SUM(NVL(B.MinusCredit,0)) SUM3 " & _
                " FROM " & GetOwner & "SO004 A," & GetOwner & "SO182 B" & _
                strWhere3 & " GROUP BY B.VODACCOUNTID "
    '將資料合併
    strTotalQry = "SELECT VODACCOUNTID,SUM(SUM1) SUM1, " & _
                " SUM(SUM2) SUM2,SUM(SUM3) SUM3 FROM (" & _
                strSum1Qry & " UNION " & strSum2Qry & " UNION " & strSum3Qry & _
                " ) GROUP BY VODACCOUNTID "
    
    strRunSQL = " FROM " & GetOwner & "SO004 A," & GetOwner & "SO182 B ,(" & strSubQry & ") C, " & _
                "(" & strTotalQry & ") D " & _
                strWhere & " AND B.VODACCOUNTID=C.VODACCOUNTID " & _
                " AND B.VODACCOUNTID=D.VODACCOUNTID " & _
                " AND D.SUM2-SUM3-SUM1 < 0 "
    '最後所要挑選的欄位
    strFields = "SELECT A.CUSTID,A.DeclarantName,B.VodAccountid,B.FaciSeqNo," & _
                " B.NoCloseTime,C.TOTAL,B.OrderNo,NVL(B.MinusCredit,0) MINUSCREDIT, " & _
                " NVL(B.AddCredit,0) ADDCREDITIT,B.RecTime,A.FACISNO,B.ROWID ID "

'                " AND ( B.NOCLOSETIME > " & intPara36 & " OR C.TOTAL > " & intPara35 & ")"
    '********************************************************************************************************
    Select Case aType
        Case 1
            strRetSQL = strFields & ",0 EXETYPE " & strRunSQL & _
                        " AND ( Nvl(C.MAXNOCLOSETIME,0) >= " & intPara36 & " OR C.TOTAL >= " & intPara35 & ")"
        Case 2
            strRetSQL = strFields & ",1 EXETYPE " & strRunSQL & _
                        " AND ( Nvl(C.MAXNOCLOSETIME,0) < " & intPara36 & " AND C.TOTAL < " & intPara35 & ")"
        Case 3
            strRetSQL = strFields & ",0 EXETYPE " & strRunSQL & _
                        " AND ( Nvl(C.MAXNOCLOSETIME,0) >= " & intPara36 & " OR C.TOTAL >= " & intPara35 & ")" & _
                        " UNION " & _
                        strFields & ",1 EXETYPE " & strRunSQL & _
                        " AND ( Nvl(C.MAXNOCLOSETIME,0) < " & intPara36 & " AND C.TOTAL < " & intPara35 & ")"

    End Select
    GetSQL2 = strRetSQL
    Exit Function
chkErr:
  Call ErrSub(Me.Name, "GetSQL2")
End Function

Private Function GetSQL(ByVal aType As Integer, ByVal aWhere As String) As String
  On Error GoTo chkErr
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
    
    '拆除,SO004要過濾拆除的資料
    'str004Where = "(A.INSTDATE IS NOT NULL AND A.PRDATE IS NOT NULL AND A.PRDATE > A.INSTDATE)"

    
    
    '****************************期間內總訂購點數的 Where 條件 *******************************
    strWhere = " WHERE A.CompCode = B.CompCode And A.ServiceType = B.ServiceType " & _
                " And CREDITTYPECODE In (SELECT CODENO FROM " & GetOwner & "CD108 " & _
                " WHERE CD108.REFNO = 3 AND NVL(CD108.STOPFLAG,0) = 0 AND B.SERVICETYPE=CD108.SERVICETYPE)" & _
                " AND NVL(B.CloseFlag,0)=0 AND NVL(B.STOPFLAG,0)= 0 AND A.SEQNO = B.FACISEQNO "
    '******************************************************************************************
    
    '****************************可使用點數的 Where 條件 *****************************************
    strWhere2 = " WHERE A.CompCode = B.CompCode And A.ServiceType = B.ServiceType " & _
               " And CREDITTYPECODE In (SELECT CODENO FROM " & GetOwner & "CD108 " & _
               " WHERE CD108.REFNO IN(1,2) AND NVL(CD108.STOPFLAG,0) = 0 AND B.SERVICETYPE=CD108.SERVICETYPE)" & _
               " AND NVL(B.STOPFLAG,0)=0 AND A.SEQNO=B.FACISEQNO "
    '**********************************************************************************************
    
    '****************************已結算過期間內總訂購點數的 Where 條件 *******************************
    strWhere3 = " WHERE A.CompCode = B.CompCode And A.ServiceType = B.ServiceType " & _
                " And CREDITTYPECODE In (SELECT CODENO FROM " & GetOwner & "CD108 " & _
                " WHERE CD108.REFNO = 3 AND NVL(CD108.STOPFLAG,0) = 0 AND B.SERVICETYPE=CD108.SERVICETYPE)" & _
                " AND NVL(B.CloseFlag,0)=1 AND NVL(B.STOPFLAG,0)=0 "
    '**************************************************************************************************
    If aWhere <> Empty Then
        strWhere = strWhere & " AND " & aWhere
        strWhere2 = strWhere2 & " AND " & aWhere
        strWhere3 = strWhere3 & " AND " & aWhere
    End If
    '****************************要結算的SQL語法*****************************************************
    
    '****************************算出期間內總訂購數之後還要判斷合計或未結算次數的條件*****************
    strSubQry = "SELECT B.VODACCOUNTID,SUM(B.MINUSCREDIT) TOTAL " & _
                    " FROM " & GetOwner & "SO004 A," & GetOwner & "SO182 B" & strWhere & _
                    " GROUP BY B.VODACCOUNTID "
    '**************************************************************************************
    strSum1Qry = "SELECT B.VODACCOUNTID,B.NOCLOSETIME,C.TOTAL,B.MINUSCREDIT " & _
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
    '算出已結算過期間內總訂購點數
    strSum3Qry = "SELECT B.VODACCOUNTID,0 SUM1,0 SUM2,SUM(NVL(B.MinusCredit,0)) SUM3 " & _
                " FROM " & GetOwner & "SO004 A," & GetOwner & "SO182 B" & _
                strWhere3 & " GROUP BY B.VODACCOUNTID "
    '將資料合併
    strTotalQry = "SELECT VODACCOUNTID,SUM(SUM1) SUM1, " & _
                " SUM(SUM2) SUM2,SUM(SUM3) SUM3 FROM (" & _
                strSum1Qry & " UNION " & strSum2Qry & " UNION " & strSum3Qry & _
                " ) GROUP BY VODACCOUNTID "
    
    strRunSQL = " FROM " & GetOwner & "SO004 A," & GetOwner & "SO182 B ,(" & strSubQry & ") C, " & _
                "(" & strTotalQry & ") D " & _
                strWhere & " AND B.VODACCOUNTID=C.VODACCOUNTID " & _
                " AND B.VODACCOUNTID=D.VODACCOUNTID " & _
                " AND D.SUM2-SUM3-SUM1 < 0 "
    '最後所要挑選的欄位
    strFields = "SELECT A.CUSTID,A.DeclarantName,B.VodAccountid,B.FaciSeqNo," & _
                " B.NoCloseTime,C.TOTAL,B.OrderNo,NVL(B.MinusCredit,0) MINUSCREDIT, " & _
                " NVL(B.AddCredit,0) ADDCREDITIT,B.RecTime,A.FACISNO,B.ROWID ID "

'                " AND ( B.NOCLOSETIME > " & intPara36 & " OR C.TOTAL > " & intPara35 & ")"
    '********************************************************************************************************
    Select Case aType
        Case 1
            strRetSQL = strFields & ",0 EXETYPE " & strRunSQL & _
                        " AND ( Nvl(B.NOCLOSETIME,0) >= " & intPara36 & " OR C.TOTAL >= " & intPara35 & ")"
        Case 2
            strRetSQL = strFields & ",1 EXETYPE " & strRunSQL & _
                        " AND ( Nvl(B.NOCLOSETIME,0) < " & intPara36 & " AND C.TOTAL < " & intPara35 & ")"
        Case 3
            strRetSQL = strFields & ",0 EXETYPE " & strRunSQL & _
                        " AND ( Nvl(B.NOCLOSETIME,0) >= " & intPara36 & " OR C.TOTAL >= " & intPara35 & ")" & _
                        " UNION " & _
                        strFields & ",1 EXETYPE " & strRunSQL & _
                        " AND ( Nvl(B.NOCLOSETIME,0) < " & intPara36 & " AND C.TOTAL < " & intPara35 & ")"

    End Select
    GetSQL = strRetSQL
    Exit Function
chkErr:
  Call ErrSub(Me.Name, "GetSQL")
End Function
'抓取欲結算之VOD訂購資料
Private Function subChoose() As Boolean
  On Error GoTo chkErr
    strChoose = Empty
    '************************畫面上的條件********************************************
    If gilCompCode.GetCodeNo <> "" Then
        Call subAnd2(strChoose, " A.CompCode = " & gilCompCode.GetCodeNo)
    End If
    If gilServiceType.GetCodeNo <> "" Then
        Call subAnd2(strChoose, " A.ServiceType='" & gilServiceType.GetCodeNo & "'")
    End If
    If txtCustId.Text <> "" Then
        Call subAnd2(strChoose, " A.CustId=" & txtCustId.Text)
    End If
    If txtSrNO.Tag <> "" Then
        Call subAnd2(strChoose, " A.SEQNO ='" & txtSrNO.Tag & "'")
    Else
        If txtSrNO.Text <> "" Then
            Call subAnd2(strChoose, "A.FaciSNO='" & txtSrNO.Text & "'")
        End If
    End If
    If gdaEndDate.GetValue <> "" Then
        Call subAnd2(strChoose, " To_Date(To_Char(B.RecTime,'YYYYMMDD'),'YYYYMMDD') <= " & _
                        "To_Date('" & gdaEndDate.GetValue & "','YYYYMMDD')")
    End If
   
    '************************************************************************************
     '*************************SO004的設備一定要參考號為3而且不能有拆機日期********************
    Call subAnd2(strChoose, "A.FaciCode In ( SELECT CODENO FROM " & GetOwner & "CD022 " & _
                        " WHERE NVL(StopFlag,0)=0 AND NVL(REFNO,0)= 3 )")

    Call subAnd2(strChoose, "A.PRDATE IS NULL ")
    '*****************************************************************************************
    strSQL = GetSQL2(3, strChoose)

    strChooseString = "公司別　: " & subSpace(strCompCode) & ";" & _
                      "服務類別: " & subSpace(strServiceType) & ";" & _
                      "結算日期期限　: " & subSpace(strEndDate) & ";" & _
                      "PPV結算作業金額限額: " & intPara35 & ";" & _
                      "PPV結算作業月數限額: " & intPara36 & _
                      IIf(txtCustId.Text <> "", ";客戶編號: " & txtCustId.Text, "") & _
                      IIf(txtSrNO.Text <> "", ";STB設備流水號:" & txtSrNO.Text, "")
                      
    subChoose = True
  Exit Function
chkErr:
  Call ErrSub(Me.Name, "subChoose")
End Function

Private Sub CreateTable()
  On Error Resume Next
    cnn.Execute "Drop Table SO1156"
  On Error GoTo chkErr
    cnn.Execute "Create Table SO1156 (CompCode text(3),Custid Text(8),AuthenticId text(10),Amount Long,48CodeNo text(10))"
  Exit Sub
chkErr:
  Call ErrSub(Me.Name, "CreateTable")
End Sub

Private Function CreateBill(ByRef aHaveData As Boolean) As Boolean
  On Error GoTo chkErr
    
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
            strUPD182 = " UPDATE " & GetOwner & "SO182 Set " & _
                        "BillNo = Decode(BillNo,Null," & "'" & strBillNo & "',BillNo)," & _
                        "Item = 1,CloseFlag = 1," & _
                        "CloseTime=To_Date('" & Format(strRunTime, "YYYYMMDDHHMMSS") & "','yyyymmddhh24miss')," & _
                        "CloseEN='" & garyGi(0) & "',CloseName='" & garyGi(1) & "'," & _
                        "UpdEn='" & garyGi(1) & "',UPDTime='" & GetDTString(strRunTime) & "', " & _
                        "CloseBillNo='" & strBillNo & "'," & _
                        "CloseItem=1" & _
                        " Where RowId = '" & rs182("ID") & "' AND AddCredit IS NULL"
    
            '以CreditSeqNo填入SO033VOD的CloseBillNo
            strUPD033Vod = " UPDATE " & GetOwner & "SO033VOD SET " & _
                            " CloseBillNo='" & strBillNo & "'," & _
                            " CloseItem=1" & _
                            " WHERE SEQNO=(SELECT CreditSeqNo FROM " & GetOwner & "SO182 " & _
                                            " WHERE ROWID='" & rs182("ID") & "' AND ADDCREDIT IS NULL )"
            gcnGi.Execute (strUPD182)
            gcnGi.Execute (strUPD033Vod)
        Else
            '如果是未執行,要將次數加1
            gcnGi.Execute (strUpdSql & " WHERE ROWID='" & rs182("ID") & "'")
        End If
        aVodAccountId = rs182("VODAccountID")
        rs182.MoveNext
    Loop
  
    CreateBill = True
    On Error Resume Next
    aHaveData = True
    Call CloseRecordset(rs182)
    Exit Function
chkErr:
   gcnGi.RollbackTrans
  Call ErrSub(Me.Name, "CreateBill")
End Function
Private Function Insert033(ByRef aBillNo As String, ByRef rsSO182 As Recordset) As Boolean
  On Error GoTo chkErr
    Dim rs033 As New ADODB.Recordset
    If Not GetRS(rs033, "SELECT * FROM " & GetOwner & "SO033 WHERE 1=0 ", gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
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
        .Fields("ShouldDate").Value = GetADdate(strRunTime)
        .Fields("OldAmt").Value = rsSO182.Fields("Total")
        .Fields("ShouldAmt").Value = rsSO182.Fields("Total")
        .Fields("OldPeriod").Value = 0
        .Fields("RealPeriod").Value = 0
        .Fields("CMCODE").Value = IIf(strCMCode = "", Null, strCMCode)
        .Fields("CMName").Value = IIf(strCMname = "", Null, strCMname)
        .Fields("UCCODE").Value = IIf(strUCCode = "", Null, strUCCode)
        .Fields("UCName").Value = IIf(strUCName = "", Null, strUCName)
        .Fields("PTCODE").Value = 1
        .Fields("PTName").Value = "現金"
        .Fields("ClassCode") = rs014("ClassCode1") & ""
        .Fields("FaciSeqNo").Value = rsSO182("FaciSeqNo") & ""
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
        .Update
    End With
    
    Insert033 = True
    On Error Resume Next
    Call CloseRecordset(rs033)
    Exit Function
chkErr:
  Call ErrSub(Me.Name, "Insert033")
End Function
Private Function Get014Data(ByVal aCustId As String) As Boolean
  On Error GoTo chkErr
    Dim strQry As String
    strQry = "SELECT A.*,B.CLASSCODE1 FROM " & GetOwner & "SO014 A," & GetOwner & "SO001 B " & _
            " WHERE B.CUSTID=" & aCustId & " AND " & IIf(intPara14 = 1, "B.ChargeAddrNo", "B.InstAddrNo") & "=A.AddrNo" & _
            " AND B.COMPCODE=" & gilCompCode.GetCodeNo & " AND A.COMPCODE=B.COMPCODE"
    If Not GetRS(rs014, strQry, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
    Get014Data = True
chkErr:
  Call ErrSub(Me.Name, "Get014Data")
End Function






Public Sub subAnd(strCH As String)
  On Error GoTo chkErr
    If strChoose <> "" Then
        strChoose = strChoose & " And " & strCH
    Else
        strChoose = strCH
    End If
  Exit Sub
chkErr:
    Call ErrSub("Utility5", "subAnd")
End Sub

Private Sub Form_Load()
  On Error GoTo chkErr
    Call subGil
    Call DefaultVal
    gilServiceType.Query_Description
  Exit Sub
chkErr:
    ErrSub Me.Name, "Form_Load"
End Sub
Private Sub Form_Activate()
    On Error Resume Next
        Screen.MousePointer = vbDefault
End Sub
Private Sub subGil()
  On Error GoTo chkErr
    SetgiList gilServiceType, "CodeNo", "Description", "CD046"
    SetgiList gilCompCode, "CodeNo", "Description", "CD039", , , , , , , GetCompCodeFilter(gilCompCode)
    If strCompCode = "" Then gilCompCode.SetCodeNo garyGi(9) Else gilCompCode.SetCodeNo (strCompCode)
    gilCompCode.Query_Description
    If strServiceType = "" Then Call SelectServType(gilCompCode.GetCodeNo, gilServiceType) Else gilServiceType.SetCodeNo (strServiceType)
     gilServiceType.GetDescription
  Exit Sub
chkErr:
    ErrSub Me.Name, "subGil"
End Sub
Private Sub DefaultVal()
  On Error GoTo chkErr
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
                " AND SERVICETYPE ='" & gilServiceType.GetCodeNo & "'"
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
        strCMname = rs044("CMNAME") & ""
    Else
        strCMCode = "": strCMname = ""
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
        lblTranDate.Caption = Format(rsSO062("TranDate").Value & "", "ee/MM/DD")
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
        gdaEndDate.SetValue DateAdd("d", 1, rsSO062("TranDate").Value & "")
        
    End If
    
    '判斷是否由別的Form進來的
    txtCustId.Text = strCustId
    If strCustId <> "" Then txtCustId.Enabled = False
    txtSrNO.Text = strSrNO
    If blnGive Then
        txtCustId.Enabled = False
        txtSrNO.Enabled = False
        cmdSTB.Enabled = False
        'gdaEndDate.SetValue strEndDate
        If strEndDate <> "" Then gdaEndDate.Enabled = False
        If intPara35In <> 0 Then lblPara35.Caption = intPara35In Else lblPara35.Caption = intPara35
        If intPara36In <> 0 Then lblPara36.Caption = intPara36In Else lblPara36.Caption = intPara36
        chkTran1.Value = 1
        chkTran2.Value = 1
        chkRun.Value = 1
        If strCompCode <> "" Then gilCompCode.Enabled = False
        If strServiceType <> "" Then gilServiceType.Enabled = False
    End If
    On Error Resume Next
    Call CloseRecordset(rsSO062)
    Call CloseRecordset(rsCD019)
    Call CloseRecordset(rs107)
    Exit Sub
chkErr:
   ErrSub Me.Name, "DefaultVal"
End Sub

Private Sub gilServiceType_Change()
  On Error GoTo chkErr
    DefaultVal
  Exit Sub
chkErr:
   ErrSub Me.Name, "gilServiceType_Change"
End Sub

Private Sub gilCompCode_Change()
  On Error GoTo chkErr
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
chkErr:
    ErrSub Me.Name, "gilCompCode_Change"
End Sub

Private Sub gilCompCode_Validate(Cancel As Boolean)
  On Error GoTo chkErr
    If gilCompCode.GetCodeNo = "" Then MsgMustBe ("公司別"): Cancel = True
  Exit Sub
chkErr:
  Call ErrSub(Me.Name, "gilCompCode_Validate")
End Sub
Public Property Let uFromBat(ByVal vData As Boolean)
    blnGive = vData
End Property
Public Property Let uCompCode(ByVal vData As String)
    strCompCode = vData
    'blnGive = True
End Property

Public Property Let uServiceType(ByVal vData As String)
    strServiceType = vData
    'gilServiceType.SetCodeNo vData
    'blnGive = True
End Property

Public Property Let uCustId(ByVal vData As String)
    strCustId = vData
    'txtCustid.Text = vData
    'blnGive = True
End Property

'STB設備序號
Public Property Let uSrNO(ByVal vData As String)
    strSrNO = vData
    'txtSrNO.Text = vData
End Property


'結算日期期限
Public Property Let uEndDate(ByVal vData As String)
    strEndDate = vData
End Property

'PPV結算作業金額限額
Public Property Let uPara35(ByVal vData As Integer)
    intPara35In = vData
End Property

'PPV結算作業月數限額
Public Property Let uPara36(ByVal vData As Integer)
    intPara36In = vData
End Property

Public Property Get uMsg() As Variant
    uMsg = strMsg
End Property

Public Property Get uDo() As Boolean
    uDo = blnDO
End Property

Private Sub txtSrNO_KeyPress(KeyAscii As Integer)
  On Error Resume Next
    txtSrNO.Tag = ""
End Sub
