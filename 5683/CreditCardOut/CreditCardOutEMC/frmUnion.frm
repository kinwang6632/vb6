VERSION 5.00
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmUnion 
   BorderStyle     =   1  '單線固定
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9075
   BeginProperty Font 
      Name            =   "新細明體"
      Size            =   9.75
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUnion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   9075
   StartUpPosition =   1  '所屬視窗中央
   Begin MSComDlg.CommonDialog cnd1 
      Left            =   6840
      Top             =   4245
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   4020
      Left            =   195
      TabIndex        =   14
      Top             =   135
      Width           =   8730
      Begin VB.ComboBox cboMonthDay 
         Height          =   315
         ItemData        =   "frmUnion.frx":0442
         Left            =   7590
         List            =   "frmUnion.frx":049A
         Style           =   2  '單純下拉式
         TabIndex        =   29
         Top             =   330
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.CheckBox chkDuteDate 
         Caption         =   "信用卡過期資料一併產生"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   300
         TabIndex        =   27
         Top             =   2010
         Width           =   2535
      End
      Begin VB.TextBox txtBillMemo 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   1800
         MaxLength       =   19
         TabIndex        =   25
         Top             =   3930
         Visible         =   0   'False
         Width           =   6645
      End
      Begin VB.Frame fraOtherBank 
         Caption         =   "聯合信用卡專用欄位"
         ForeColor       =   &H00FF0000&
         Height          =   1200
         Left            =   300
         TabIndex        =   20
         Top             =   750
         Width           =   8160
         Begin VB.TextBox txtAuthorizationBatch 
            Height          =   300
            Left            =   1170
            MaxLength       =   8
            TabIndex        =   24
            Text            =   "00001"
            Top             =   285
            Width           =   1575
         End
         Begin VB.OptionButton opyymm 
            Caption         =   "YYMM"
            Height          =   195
            Left            =   6435
            TabIndex        =   7
            Top             =   930
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton optmmyy 
            Caption         =   "MMYY"
            Height          =   195
            Left            =   6435
            TabIndex        =   6
            Top             =   660
            Width           =   1095
         End
         Begin VB.ComboBox cboMore 
            Height          =   315
            ItemData        =   "frmUnion.frx":0505
            Left            =   1170
            List            =   "frmUnion.frx":0512
            TabIndex        =   4
            Top             =   675
            Width           =   2460
         End
         Begin VB.CheckBox chkCancel 
            Caption         =   "取消"
            Height          =   300
            Left            =   5085
            TabIndex        =   5
            Top             =   660
            Width           =   750
         End
         Begin VB.ComboBox cboSet 
            Height          =   315
            ItemData        =   "frmUnion.frx":0556
            Left            =   1155
            List            =   "frmUnion.frx":0563
            TabIndex        =   2
            Top             =   270
            Width           =   2460
         End
         Begin VB.TextBox txtMerchantName 
            Height          =   300
            Left            =   5085
            MaxLength       =   8
            TabIndex        =   3
            Top             =   255
            Width           =   2460
         End
         Begin VB.Label Label5 
            Alignment       =   1  '靠右對齊
            Caption         =   "交易旗號"
            Height          =   285
            Left            =   225
            TabIndex        =   23
            Top             =   750
            Width           =   795
         End
         Begin VB.Label Label1 
            Alignment       =   1  '靠右對齊
            Caption         =   "交易碼"
            Height          =   285
            Left            =   90
            TabIndex        =   22
            Top             =   345
            Width           =   915
         End
         Begin VB.Label Label8 
            Alignment       =   1  '靠右對齊
            Caption         =   "端末機代號"
            Height          =   285
            Left            =   3900
            TabIndex        =   21
            Top             =   315
            Width           =   1080
         End
      End
      Begin Gi_Date.GiDate GiDAccept 
         Height          =   285
         Left            =   5145
         TabIndex        =   1
         Top             =   360
         Width           =   1140
         _ExtentX        =   2011
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
      Begin VB.TextBox txtSpcNO 
         Height          =   300
         Left            =   1455
         TabIndex        =   0
         Top             =   352
         Width           =   2475
      End
      Begin VB.Frame frmData 
         Caption         =   "資料檔位置"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1560
         HelpContextID   =   2
         Left            =   285
         TabIndex        =   15
         Top             =   2310
         Width           =   8190
         Begin VB.TextBox txtDataPath 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2625
            TabIndex        =   8
            ToolTipText     =   "請輸入字檔之路徑及檔名！"
            Top             =   480
            Width           =   4725
         End
         Begin VB.CommandButton cmdErrPath 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   14.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   7620
            TabIndex        =   11
            Top             =   960
            Width           =   375
         End
         Begin VB.CommandButton cmdDataPath 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   14.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   7605
            TabIndex        =   9
            Top             =   495
            Width           =   375
         End
         Begin VB.TextBox txtErrPath 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2625
            TabIndex        =   10
            Top             =   945
            Width           =   4725
         End
         Begin VB.Label lblDatapath 
            AutoSize        =   -1  'True
            Caption         =   "資料檔位置(路徑)"
            Height          =   195
            Left            =   120
            TabIndex        =   17
            Top             =   555
            Width           =   1485
         End
         Begin VB.Label lblErrpath 
            AutoSize        =   -1  'True
            Caption         =   "問題參考檔位置(路徑 + 名稱)"
            Height          =   195
            Left            =   105
            TabIndex        =   16
            Top             =   990
            Width           =   2460
         End
      End
      Begin VB.Label lblMonthDay 
         Caption         =   "每月付款日"
         Height          =   210
         Left            =   6480
         TabIndex        =   28
         Top             =   390
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label lblBillMessage 
         AutoSize        =   -1  'True
         Caption         =   "帳單提列訊息"
         Height          =   195
         Left            =   450
         TabIndex        =   26
         Top             =   4110
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.Label Label7 
         Alignment       =   1  '靠右對齊
         Caption         =   "商店代號"
         Height          =   285
         Left            =   360
         TabIndex        =   19
         Top             =   360
         Width           =   945
      End
      Begin VB.Label Label2 
         Caption         =   "送件日期"
         Height          =   210
         Left            =   4290
         TabIndex        =   18
         Top             =   390
         Width           =   780
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "F2.開始"
      Height          =   375
      Left            =   480
      TabIndex        =   12
      Top             =   4320
      Width           =   1875
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&X)"
      Height          =   375
      Left            =   7410
      TabIndex        =   13
      Top             =   4320
      Width           =   1275
   End
End
Attribute VB_Name = "frmUnion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
''
Private cn As New ADODB.Connection
Private rsTmp As New ADODB.Recordset     '' 取出實體資料庫 的資料
Private rsDefTmp As New ADODB.Recordset    ''存文字檔的記憶體 recordset

Private strPrgName As String  ''程式名稱
Private strPath As String     ''GICMIS1.INI路徑
Private strDataName As String  ''檔案名稱
Private strHeader As String   ''記錄標頭檔的文字
Private strChoose As String  ''
Private lngErrCount As Long   '' 記錄錯誤筆數

Public IsTCB As Boolean
Public IsTCBCard As Boolean
Public strChooseMultiAcc As String   '' 這個參數用來承接篩選 SO033 所需的收費資料
Public blnfrmChina  As Boolean ' 2003/10/08   由 SO3272A  傳入，指示其 為 中國信託信用卡格式的
Dim strBankSQL As String
Dim strBankId As String
Public frmCreditCardType As mCreditCardType     ''  2003/10/28 這個值代表所要使用的信用卡種類
Dim intBatchNumber As Integer  '' 記錄 HSBC今日出過的批次
Dim strBatchDate As String     '' 記錄 HSBC出過批次的日期

Private Sub chkDuteDate_Click()
  On Error Resume Next
    '#4090 增加信用卡過期要產生Log By Kin 2008/09/19
    If chkDuteDate.Value = 1 Then
        Call MsgBox("當您選擇勾選這個選項，信用卡過期資料除了被記錄之外，" & vbCrLf & _
                             "亦將被一併產生，若是您不打算這麼作，請取消這個勾選。", vbInformation, "過期資料產生")
    End If

End Sub

Private Sub cmdCancel_Click()
On Error Resume Next
    Unload Me
End Sub

Private Sub cmdOK_Click()

 '''If Not ChkDTok Then Exit Sub
    If Not IsDataOk Then Exit Sub
    Call DefineRs   ''建立記憶體 recordset 結構
 
    Dim sFile As String
    Dim sPath As String
    sFile = txtDataPath.Text
    If frmCreditCardType = lngHSBC Then
        sPath = txtDataPath.Text
        If Right(sPath, 1) <> "\" Then sPath = sPath & "\"
        sFile = sPath & "\" & txtAuthorizationBatch.Text & Right(GiDAccept.GetValue, 6)
    ElseIf frmCreditCardType = lngNewWeb Then   '#5494 增加藍新科技檔名 By Kin 2010/02/03
        sPath = txtDataPath.Text
        If Right(sPath, 1) <> "\" Then sPath = sPath & "\"
        sFile = sPath & txtSpcNO.Text & "_" & GiDAccept.GetValue & "." & txtAuthorizationBatch.Text
    ElseIf frmCreditCardType = lngUnionBatch Then
    
        sPath = txtDataPath.Text
        Dim iSch As Long
        
        iSch = InStrRev(sPath, "\")
        sPath = Mid(sPath, 1, iSch)
        txtDataPath.Text = sPath
        If Right(sPath, 1) <> "\" Then sPath = sPath & "\"
        sFile = sPath & txtMerchantName.Text & ".TXT"
    End If
    
    If OpenFile(sFile, 0) = False Then Exit Sub
    If OpenFile(txtErrPath.Text, 1) = False Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    objStorePara.uProcText = txtDataPath.Text
    objStorePara.uProcErrText = txtErrPath.Text
    Call ScrToRcd   '#4018 按執行時就要將資料Keep起來 By Kin 2008/07/18
    If Not BeginTran Then
        If Not blnNodata Then
            objStorePara.uUpdate = False
            MsgBox "轉帳資料輸出失敗!", vbExclamation, "警告..."
            CloseFS
            Screen.MousePointer = vbDefault
            Unload Me
            Exit Sub
        Else
            objStorePara.uNoData = True
            CloseFS
            Screen.MousePointer = vbDefault
            Unload Me
            Exit Sub
        End If
    End If
    'Call ScrToRcd
    objStorePara.uUpdate = True
    Unload Me
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "cmdOK_Click")
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error Resume Next
    If KeyCode = vbKeyF2 Then
        cmdOK.Value = cmdOK.Enabled
    End If
  
End Sub

Private Sub Form_Load()
    On Error Resume Next
    
      Me.Caption = objStorePara.BankName & ""
      cboMore.ListIndex = 2
      GiDAccept.Text = Format(Date, "eemmdd")
      Call InitData
      RcdToScr
      If IsTCB = True Then
           Dim rsCompName As New ADODB.Recordset
        '            rsCompName.CursorLocation = adUseClient
        '            rsCompName.Open "SELECT DESCRIPTION FROM " & TableOwnerName & "CD039", cn, adOpenKeyset, adLockReadOnly
           If Not GetRS(rsCompName, "SELECT DESCRIPTION FROM " & TableOwnerName & "CD039", cn) Then Exit Sub
           txtSpcNO.Text = rsCompName(0)
           rsCompName.Close
           Set rsCompName = Nothing
           fraOtherBank.Enabled = False
        ''   txtSpcNO.Locked = True
            
      End If
      
'      If blnfrmChina = False Then
'                 fraOtherBank.Enabled = False
'      End If
      
End Sub
Private Sub RcdToScr()
    On Error GoTo ChkErr
    Dim LogFile As TextStream
    Dim interr As Integer
    
    interr = 0
    Set LogFile = Nothing
    interr = 1
    If Dir(strPath & "\" & strPrgName & "Out.log") = "" Then Exit Sub
    interr = 2
    Set LogFile = fso.OpenTextFile(strPath & "\" & strPrgName & "Out.log", ForReading)
    interr = 3
    '#5494 多增加讀取每月付款日 By Kin 2010/02/03
    With LogFile
        If Not .AtEndOfStream Then txtSpcNO = .ReadLine & "":      interr = 4
        If Not .AtEndOfStream Then txtMerchantName = .ReadLine & "":       interr = 5
        If Not .AtEndOfStream Then txtDataPath = .ReadLine & "":        interr = 6 '資料檔
        If Not .AtEndOfStream Then txtErrPath = .ReadLine & "":        interr = 7 '問題檔
        'If Not .AtEndOfStream Then cboMonthDay.Text = .ReadLine & "":    interr = 8
        
    End With
    interr = 71
    If frmCreditCardType = 2 Then
        interr = 72
        If Len(txtDataPath.Text) = 0 Then
            intBatchNumber = 0: interr = 8
            strBatchDate = Format(Date, "eemmdd"): interr = 9
        Else
            interr = 91
            Dim aPath() As String
            Dim varAnayPath
            Dim strAnayPath As String
            aPath = Split(txtDataPath.Text, "\"): interr = 92
            For Each varAnayPath In aPath
                strAnayPath = varAnayPath: interr = 94
            Next
             '' 2004/03/04 今天加入以下的判斷式，因為若是strAnayPath是空的，會出現錯誤
            If Len(strAnayPath) < 1 Then
                strBatchDate = Format(Date, "eemmdd"): interr = 941
                intBatchNumber = 0: interr = 942
            Else
                aPath = Split(strAnayPath, "."): interr = 95
                strAnayPath = aPath(0): interr = 96
                If Right(strAnayPath, 6) <> Format(Date, "eemmdd") Then
                    strBatchDate = Format(Date, "eemmdd"): interr = 97
                    intBatchNumber = 0: interr = 98
                Else
                    strBatchDate = Right(strAnayPath, 6): interr = 99
                    intBatchNumber = CInt(Mid(strAnayPath, 2, Len(strAnayPath) - 7)): interr = 991
                End If
                interr = 10
            End If
        End If
    End If
    interr = 11
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "RcdToScr" & "   errcode :" & interr)
End Sub

Private Sub InitData()
  On Error GoTo ChkErr
'  Dim o As Object
'  Set o = CreateObject("CreditCardOut.clsInterface")
    With objStorePara
        Set cn = .Connection
        strBankId = .BankId
        '#3531 多增加銀行條件(strBankId條件寫得太死) By Kin 2007/11/02
        strBankSQL = .BankStr
        'strBankName = .BankName
        'strCorpId = .CorpID
        strChoose = .ChooseStr
        strPath = .ErrPath
        txtSpcNO.Text = .uSpcNO
    End With
    If frmCreditCardType <> lngNewWeb Then
        If frmCreditCardType <> lngHSBC Then lblDatapath.Caption = "資料檔位置(路徑 + 名稱)"
    End If
    
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "InitData")
End Sub

Public Property Get PrgName() As String
    PrgName = strPrgName
End Property

Public Property Let PrgName(ByVal vData As String)
    strPrgName = vData
End Property


Private Sub cmdDataPath_Click()
  On Error GoTo ChkErr
    '#5494 增加藍新科技的判斷 By Kin 2010/02/03
    If (frmCreditCardType = lngHSBC) Or (frmCreditCardType = lngNewWeb) Or (frmCreditCardType = lngUnionBatch) Then
        txtDataPath.Text = FolderDialog("開啟路徑")
'        If frmCreditCardType = lngUnionBatch Then
'            If txtMerchantName.Text <> "" Then
'                If Right(txtDataPath.Text, 1) = "\" Then
'                    txtDataPath.Text = txtDataPath & txtMerchantName.Text & ".TXT"
'                Else
'                    txtDataPath.Text = txtDataPath & "\" & txtMerchantName.Text & ".TXT"
'                End If
'            End If
'        End If
    Else
        With cnd1
            .FileName = txtErrPath.Text
            .Filter = "文字檔|*.txt"
            .InitDir = strPath
            .Action = 1
            txtDataPath = .FileName
        End With
    End If
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "cmdDataPath_Click")
End Sub

Private Sub cmdErrPath_Click()
    On Error GoTo ChkErr
    With cnd1
        .FileName = txtErrPath.Text
        .Filter = "文字檔|*.txt"
        .InitDir = strPath
        .Action = 1
        txtErrPath = .FileName
    End With
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "cmdErrPath_Click")
End Sub

Private Sub DefineRs()
    On Error GoTo ChkErr
    With rsDefTmp
        If .State = adStateOpen Then .Close
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        Select Case frmCreditCardType
        Case 0
            If IsTCB = False Then
            ''"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
                .Fields.Append ("F110"), adBSTR, 10, adFldIsNullable
                .Fields.Append ("F1118"), adBSTR, 8, adFldIsNullable
                .Fields.Append ("F1937"), adBSTR, 19, adFldIsNullable
                .Fields.Append ("F3845"), adBSTR, 8, adFldIsNullable
                .Fields.Append ("F4653"), adBSTR, 8, adFldIsNullable
                
                .Fields.Append ("F5455"), adBSTR, 2, adFldIsNullable
                .Fields.Append ("F5661"), adBSTR, 6, adFldIsNullable
                .Fields.Append ("F6265"), adBSTR, 4, adFldIsNullable
                .Fields.Append ("F6668"), adBSTR, 3, adFldIsNullable
                .Fields.Append ("F6984"), adBSTR, 16, adFldIsNullable
                
                .Fields.Append ("F8590"), adBSTR, 6, adFldIsNullable
                .Fields.Append ("F9196"), adBSTR, 6, adFldIsNullable
                .Fields.Append ("F97112"), adBSTR, 16, adFldIsNullable
                .Fields.Append ("F113118"), adBSTR, 6, adFldIsNullable
                .Fields.Append ("F119"), adBSTR, 1, adFldIsNullable
                
                .Fields.Append ("F120"), adBSTR, 1, adFldIsNullable
                    
            Else   '' // 以下為中華商銀的格式
                If IsTCBCard = True Then
                    .Fields.Append ("F06"), adBSTR, 6, adFldIsNullable   '' 請款日期
                    .Fields.Append ("F10"), adBSTR, 4, adFldIsNullable   '' 序號
                    .Fields.Append ("F26"), adBSTR, 16, adFldIsNullable  '' 卡號
                    .Fields.Append ("F29"), adBSTR, 3, adFldIsNullable   '' CVC2
                    .Fields.Append ("F33"), adBSTR, 4, adFldIsNullable   '' 有效年月
                    .Fields.Append ("F44"), adBSTR, 11, adFldIsNullable  ''訂單編號
                       
                    .Fields.Append ("F50"), adBSTR, 6, adFldIsNullable   ''交易日期
                    .Fields.Append ("F56"), adBSTR, 6, adFldIsNullable   '' 授權碼
                    .Fields.Append ("F58"), adBSTR, 2, adFldIsNullable   '' 期數
                    .Fields.Append ("F65"), adBSTR, 7, adFldIsNullable   '' 分期金額
                    .Fields.Append ("F72"), adBSTR, 7, adFldIsNullable    '' 總金額
                    .Fields.Append ("F79"), adBSTR, 7, adFldIsNullable    '' 頭 款
                    
                    .Fields.Append ("F86"), adBSTR, 7, adFldIsNullable  ''  尾款
                    .Fields.Append ("F91"), adBSTR, 5, adFldIsNullable   '' 特站代號
                    .Fields.Append ("F116"), adBSTR, 25, adFldIsNullable  '' 特店名稱
                    .Fields.Append ("F136"), adBSTR, 20, adFldIsNullable  '' 特店名稱
                    .Fields.Append ("F138"), adBSTR, 2, adFldIsNullable   ''  交易類別
                    
                    .Fields.Append ("F180"), adBSTR, 62, adFldIsNullable   ''  空白
                Else
                    .Fields.Append ("F15"), adBSTR, 16, adFldIsNullable   '' 訂單編號
                    .Fields.Append ("F16"), adBSTR, 15, adFldIsNullable   '' 銷貨編號
                    .Fields.Append ("F31"), adBSTR, 19, adFldIsNullable  '' 卡號
                    .Fields.Append ("F50"), adBSTR, 3, adFldIsNullable   '' CVC2
                    .Fields.Append ("F53"), adBSTR, 4, adFldIsNullable   '' 有效年月(YYMM)
                    .Fields.Append ("F57"), adBSTR, 10, adFldIsNullable  ''交易金額
                    .Fields.Append ("F67"), adBSTR, 2, adFldIsNullable  ''長設值
                    .Fields.Append ("F69"), adBSTR, 2, adFldIsNullable   ''期數
                    .Fields.Append ("F71"), adBSTR, 12, adFldIsNullable   '' 頭款
                    .Fields.Append ("F83"), adBSTR, 12, adFldIsNullable  ''尾款
                End If
            End If
        Case 1
           ''  2003/10/08   如果是中國商銀的改走以下這一段
    
            ''''特店代號：依業務有不同代號 , 尚未確定
            ''''卡號：聯合信用卡時卡號只有11位: 先放4位 '4000',再放11位卡號,再放1位'0',組成16位 ; 其他卡皆16位
            ''''交易碼 : 原始 ‘0 ‘+空白 , 經授權處理後會變為 ‘1‘+空白 , 強迫授權(二次授權)後會變為 ‘2‘+空白
            ''''流水號:  該筆明細在檔案之第幾筆 , 絕對不能重復
            ''''回應碼 : ‘1+空白+空白’及‘ 5+空白+空白’屬成功交易,其對應之回應訊息必定為APPROVED ; 除此之外之回應碼均屬授權失敗交易
            ''''回應訊息 : ‘APPROVED’代表成功 ; ‘REFERRAL’代表CALL BANK,其他則代表失敗,最常見為TRANSACTION ERRO’代表資料錯誤 ; ‘INVALID CAPTURE’代表其他失敗 ; ‘DECLINE’代表其他失敗,詳細原因煩請參照”批次回應碼”一表

            ''  以下標示為　"**"  表示必需由程式填入，其他則是銀行回應碼
            
            .Fields.Append ("F113"), adBSTR, 13, adFldIsNullable   ''  ** 特店代號(類似聯合:端末機代號)
            .Fields.Append ("F1429"), adBSTR, 16, adFldIsNullable  ''  ** 卡號
            .Fields.Append ("F3039"), adBSTR, 10, adFldIsNullable  ''  ** 金額
            .Fields.Append ("F4043"), adBSTR, 4, adFldIsNullable   ''   ** 有效月年 (MMYY) (西元年)
            .Fields.Append ("F4445"), adBSTR, 2, adFldIsNullable   '''   ** 交易碼   (0+空白)
            .Fields.Append ("F4659"), adBSTR, 14, adFldIsNullable  '''   ** 流水號(絕對不可重複)  (不足補0)
             '' 以下為銀行回覆用的
            .Fields.Append ("F6065"), adBSTR, 6, adFldIsNullable  ''  APPROVED CODE
            .Fields.Append ("F6671"), adBSTR, 6, adFldIsNullable   ''   APPROVED DATE (YYMMDD) (西元)
            .Fields.Append ("F7274"), adBSTR, 3, adFldIsNullable   ''  回應碼
            .Fields.Append ("F7590"), adBSTR, 16, adFldIsNullable  '''  回應訊息
            
            .Fields.Append ("F9196"), adBSTR, 6, adFldIsNullable   ''  授權日期(APPROVED DATE)
            .Fields.Append ("F97102"), adBSTR, 6, adFldIsNullable   ''  授權時間(時/分/秒)
        Case 2
                ''  2003/11/20  以下匯豐銀行信用卡格式
            .Fields.Append ("F116"), adBSTR, 16, adFldIsNullable   ''  ** 卡號  不足，依系統規定補足16碼space
            .Fields.Append ("F1717"), adBSTR, 1, adFldIsNullable  ''  ** 空白
            .Fields.Append ("F1821"), adBSTR, 4, adFldIsNullable  ''  ** 有效期限  MBF,MMYY(西元年)
            .Fields.Append ("F2222"), adBSTR, 1, adFldIsNullable   ''   ** " " or "-"
            .Fields.Append ("F2334"), adBSTR, 12, adFldIsNullable   '''   金額  不足補'0'後兩碼為角分
            .Fields.Append ("F3535"), adBSTR, 1, adFldIsNullable  '''  空白
        
            .Fields.Append ("F3641"), adBSTR, 6, adFldIsNullable  ''   核淮號碼   輸入'999999'or approvaled code
            .Fields.Append ("F4242"), adBSTR, 1, adFldIsNullable   ''   空白
            .Fields.Append ("F4344"), adBSTR, 2, adFldIsNullable   ''  回復碼   輸入'99'or '0'(if this is a refund)
            .Fields.Append ("F4545"), adBSTR, 1, adFldIsNullable  '''  空白
            .Fields.Append ("F4685"), adBSTR, 40, adFldIsNullable   ''  商店自行運用 (so043.para24=1則取媒體單號
        Case 3
            .Fields.Append ("SPCNO"), adBSTR, 5, adFldIsNullable
            .Fields.Append ("ACCOUNTID"), adBSTR, 19, adFldIsNullable
            .Fields.Append ("STOPYM"), adBSTR, 4, adFldIsNullable
            .Fields.Append ("SHOULDAMT"), adBSTR, 13, adFldIsNullable
            .Fields.Append ("BILLNO"), adBSTR, 15, adFldIsNullable
            .Fields.Append ("TRADETYPE"), adBSTR, 1, adFldIsNullable
            .Fields.Append ("CHINESE"), adBSTR, 40, adFldIsNullable
        '#5494 增加藍新科技格式 By Kin 2010/01/27
        Case 5
            .Fields.Append "RcdType", adBSTR, 10, adFldIsNullable   '記錄型別,固定填D
            .Fields.Append "OrderNo", adBSTR, 30, adFldIsNullable   '商店訂單單號
            .Fields.Append "OrdItem", adBSTR, 9, adFldIsNullable    '訂單編號 (交易序號)
            .Fields.Append "Period", adBSTR, 2, adFldIsNullable     '分期期數
            .Fields.Append "Status", adBSTR, 1, adFldIsNullable     '交易狀態
            .Fields.Append "Amt", adBSTR, 13, adFldIsNullable       '訂單金額(含2位小數點)
            .Fields.Append "CardNo", adBSTR, 16, adFldIsNullable    '信用卡卡號
            .Fields.Append "CVC2", adBSTR, 3, adFldIsNullable       'CVC2
            .Fields.Append "ValidDate", adBSTR, 6, adFldIsNullable  '有效年月
            .Fields.Append "AuthNo", adBSTR, 6, adFldIsNullable     '授權碼
            .Fields.Append "Prc", adBSTR, 5, adFldIsNullable        '主回覆碼
            .Fields.Append "Src", adBSTR, 5, adFldIsNullable        '次回覆碼
            .Fields.Append "RetBank", adBSTR, 2, adFldIsNullable    '銀行回覆碼
            .Fields.Append "BatNum", adBSTR, 9, adFldIsNullable     '批次號碼
            .Fields.Append "AutoFlag", adBSTR, 1, adFldIsNullable   '自動請款旗標
            .Fields.Append "hold", adBSTR, 91, adFldIsNullable      '保留欄位
            
        '#5609 增加聯合信用卡批次 By Kin 2010/06/03
        Case 6
            Call DefRs_6(rsDefTmp)
        End Select
        .Open
    End With
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "DefineRs")
End Sub
Private Sub DefRs_6(ByRef rsDef As Recordset)
  On Error GoTo ChkErr
    With rsDef
        .Fields.Append "SPCNO", adBSTR, 10, adFldIsNullable         '商店代碼
        .Fields.Append "MERCHANTNAME", adBSTR, 8, adFldIsNullable   '端末機代號
        .Fields.Append "ACCOUNTNO", adBSTR, 16, adFldIsNullable     '卡號
        .Fields.Append "ACCOUNTNOEX", adBSTR, 3, adFldIsNullable    '卡號延伸碼
        .Fields.Append "STOPYM", adBSTR, 4, adFldIsNullable         '有效期限
        .Fields.Append "AMT", adBSTR, 10, adFldIsNullable           '交易金額
        .Fields.Append "AUTHNO", adBSTR, 8, adFldIsNullable         '授權碼
        .Fields.Append "SETNO", adBSTR, 2, adFldIsNullable          '交易碼
        .Fields.Append "ACCEPTDATE", adBSTR, 6, adFldIsNullable     '送件日期
        .Fields.Append "BILLNO", adBSTR, 16, adFldIsNullable        'USER DEFINE
        .Fields.Append "ADDITION", adBSTR, 30, adFldIsNullable      'Additional field
        .Fields.Append "ACCINFO", adBSTR, 40, adFldIsNullable       '卡人資訊
    End With
    Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "DefRs_6")
End Sub
Private Function BeginTran() As Boolean
  On Error GoTo ChkErr
    Dim strsql As String
    Dim strOldBillNo As String
    Dim blnFirst As Boolean
    
    Dim lngErrCount As Long
    Dim lngCount As Long
    Dim lngTime As Long
    Dim lngIndex As Long
    
    Dim strSequenceNumber As String   '' 取得對應到此單號的物件序號
    Dim intPara24 As Integer  ''  記錄是否使用媒體多帳戶處理
    Dim strMediaBillNOField As String   '' 記錄目前為 BillNO 或是  MediaBillNO
    Dim rsMedioBillnoNull As New Recordset   '' 這個參數用來承接篩選 SO033裏 mediobillno 是null 的資料
    Dim strMedioBillnoNull As String    '' 這個參數用來承接篩選 SO033裏 mediobillno 是null 的資料
    Dim rsCustIDErr As New ADODB.Recordset  ''2003/08/17
    Dim lngTotalHSBC As Long: lngTotalHSBC = 0
    Dim lngToatalCountHSBC As Long: lngToatalCountHSBC = 0
    Dim lngTotalHSBCN As Long:   lngTotalHSBCN = 0 ''記錄金額為負的總數
    Dim lngTotalCountHSBCN As Long:   lngTotalCountHSBCN = 0  ''記錄金額為負的筆數
    Dim begintrnsErrCode As Double
    Dim strStopYmSQL As String
    Dim lngShouldAmt As Long
    Dim rsSO106 As New ADODB.Recordset
    Dim rsCD013 As New ADODB.Recordset
    Dim rsUpdUCCode As New ADODB.Recordset
    Dim strUCCode As String
    Dim strUCName As String
    Dim strUpdUCCode As String
    Dim blnUpdUCCode As Boolean
    Dim strUCCodeWhere As String
    Dim strStopYmWhere As String
    begintrnsErrCode = 0
    lngShouldAmt = 0
    
    '********************************************************************************************************************************************************
    '#3417 電子檔匯出時,要填入未收原因(RefNo=4) By Kin 2007/12/05
    If Not GetRS(rsCD013, "Select * From " & TableOwnerName & "CD013 Where RefNo=4 And StopFlag<>1 Order By CodeNo Desc", cn, adUseClient, adOpenKeyset) Then Exit Function
    If rsCD013.EOF Then
        blnUpdUCCode = False
    Else
        blnUpdUCCode = True
        strUCCode = rsCD013("CodeNo") & ""
        strUCName = rsCD013("Description") & ""
    End If
    '*********************************************************************************************************************************************************
    strStopYmWhere = " Max(Decode(Length(C.StopYm),5," & _
                " Substr(C.StopYm,2,4) || '0' || Substr(C.StopYm,1,1), " & _
                " Substr(C.StopYm,3,4) || Substr(C.StopYm,1,2))) CardExpDate"

    '#3728 強迫以多媒体出帳 By Kin 2008/03/25
    'intPara24 = CInt(RPxx(cn.Execute("SELECT Para24 FROM " & TableOwnerName & "SO043 Where Rownum=1").GetString))
    intPara24 = 1
    If intPara24 = 0 Then
        strMediaBillNOField = "BillNO"
    Else
        strMediaBillNOField = "MediaBillNO"
    End If
     BeginTran = False
     lngTime = Timer
     strsql = ""
        
 '' 這個欄位   "TCBbudget "  用以取得分期期數
    strsql = ""
 
 ''2003/08/18  將其中的CUSTID 拿掉 ，由底下重新取得
    '#5267 原本取SO002A現在改取SO106 By Kin 2010/04/20
    If IsTCB = False Then    '' 如果是中華商銀的格式，SQL 語法還要將其中的 CitemCode 這個項目群組出來
    ''  匯豐要依卡號作排序 ，所以新增以的的程式碼
        Dim strOrderbyAccountNO As String
        strOrderbyAccountNO = strMediaBillNOField
        If frmCreditCardType = lngHSBC Then
             strOrderbyAccountNO = "AccountNO  DESC  "
        End If
    '' errcode**********
    '' 20050203 移除And A.ShouldAmt > 0的條件限制
        begintrnsErrCode = 1
        '#5055 雖然沒有提到,但是順便將櫃台已收過濾掉 By Kin 2009/04/30
        '#5218 櫃台已收要用Nvl(CD013.REFNO,0) By Kin 2009/08/05
        '#5564 參考號7,8與CD013.PAYOK=1都是代表已收 By Kin 2010/05/17
        '#5683 增加挑選RealStopDate 與 PayKind By Kin 2010/08/06
        strsql = "SELECT " & GetUseIndexStr("A", "ShouldAmt") & " " & _
            " A." & strMediaBillNOField & ", sum(A.ShouldAmt)  ShouldAmt," & _
            " A.AccountNO,B.CUSTNAME,B.CUSTID," & _
            " SUM(A.TCBbudget) TCBbudget , MAX(A.PTCODE)  PtCode," & _
            " MAX(A.ServiceType) ServiceType,  " & _
            strMinPayKind & " , " & strMinRealStopDate & _
            " FROM " & TableOwnerName & "SO033 A," & TableOwnerName & "SO001 B ," & TableOwnerName & "SO002 D " & _
            "," & TableOwnerName & "CD013 " & _
            "Where " & strBankId & IIf(Len(strBankId) > 0, " AND ", " ") & _
            "A.CustID  = D.CUSTID AND " & _
            "A.CustId=B.CustId And A.CancelFlag = 0 And " & _
            "A.UCCode Is Not Null  AND " & _
            "A.UCCode=CD013.CodeNo And Nvl(CD013.REFNO,0) NOT IN(3,7) AND " & _
            "NVL(CD013.PAYOK,0)= 0 " & _
            IIf(Len(strChoose) = 0, "", " And ") & strChoose & " AND " & _
            "A.SERVICETYPE = D.SERVICETYPE   " & _
            "Group by A." & strMediaBillNOField & " ,  A.AccountNO  , B.CUSTNAME,B.CUSTID  ORDER BY A." & strOrderbyAccountNO
                
        strsql = "SELECT A." & strMediaBillNOField & " BILLNO ,ShouldAmt," & _
            " A.AccountNO,A.CUSTNAME,A.CUSTID,Min(A.RealStopDate) RealStopDate," & _
            " Min(A.PayKind) PayKind , " & _
            " A.TCBbudget , A.PTCODE,A.ServiceType,C.CVC2, " & strStopYmWhere & _
            " FROM (" & strsql & ") A, " & TableOwnerName & "SO106 C " & _
            " WHERE A.ACCOUNTNO=C.ACCOUNTID AND A.CUSTID=C.CUSTID " & _
            " AND C.STOPFLAG<>1 AND C.SnactionDate IS NOT NULL " & _
            " GROUP BY A." & strMediaBillNOField & " ,ShouldAmt,A.AccountNO," & _
              " A.CUSTNAME,A.CUSTID,A.TCBbudget,A.PTCODE,A.ServiceType,C.CVC2 " & _
            " ORDER BY A." & strOrderbyAccountNO
            
            
            
    Else

        '#5564 參考號7,8與CD013.PAYOK=1都是代表已收 By Kin 2010/05/17
        strsql = "SELECT " & GetUseIndexStr("A", "ShouldAmt") & " " & _
                   " A." & strMediaBillNOField & ", sum(A.ShouldAmt)  ShouldAmt," & _
                   " A.AccountNO,B.CUSTNAME,B.CUSTID," & _
                   " A.CitemCode," & strMinPayKind & " , " & strMinRealStopDate & " , " & _
                   " MAX(A.PTCODE)  PtCode,MAX(A.ServiceType) ServiceType," & _
                   " MAX(A.TCBbudget) TCBbudget  " & _
                   " FROM " & TableOwnerName & "SO033 A," & TableOwnerName & "SO001 B ," & TableOwnerName & "SO002 D " & _
                   "," & TableOwnerName & "CD013 " & _
                   " Where " & strBankId & IIf(Len(strBankId) > 0, " AND ", " ") & _
                   " A.CustID  = D.CUSTID AND " & _
                   " A.CustId=B.CustId And A.CancelFlag = 0 And " & _
                   " A.UCCode Is Not Null  And " & _
                   " A.UCCode=CD013.CodeNo And Nvl(CD013.REFNO,0) NOT IN(3,7) AND  " & _
                   " NVL(CD013.PAYOK,0) = 0 " & _
                   IIf(Len(strChoose) = 0, "", " And ") & strChoose & " AND " & _
                   " A.SERVICETYPE = D.SERVICETYPE " & _
                   " Group by A." & strMediaBillNOField & " ,  A.AccountNO ,  B.CUSTNAME , A.CItemCode,B.CUSTID  " & _
                   " ORDER BY A." & strMediaBillNOField
                   
        strsql = "SELECT A." & strMediaBillNOField & " BILLNO ,A.ShouldAmt," & _
                " A.AccountNO,A.CUSTNAME,Min(A.PAYKIND) PAYKIND,MIN(REALSTOPDATE) REALSTOPDATE ," & _
                " A.CUSTID,A.CitemCode,A.PtCode,A.ServiceType,TCBbudget,C.CVC2, " & strStopYmWhere & _
                " From (" & strsql & ") A," & TableOwnerName & "SO106 C " & _
                " WHERE A.ACCOUNTNO=C.ACCOUNTID AND A.CUSTID=C.CUSTID " & _
                " AND C.STOPFLAG<>1 AND C.SnactionDate IS NOT NULL " & _
                " GROUP BY A." & strMediaBillNOField & " ,ShouldAmt,A.AccountNO," & _
                " A.CUSTNAME,A.CUSTID,A.CitemCode,A.PtCode,A.ServiceType,A.TCBbudget,C.CVC2 " & _
                " ORDER BY A." & strMediaBillNOField
        
    End If
    
    '*********************************************************************************************
    '#3417 更新UCCode與UCName的基本Where 條件 By Kin 2007/12/05
    '#5267 C.原本是SO002A,改用SO106
    strUCCodeWhere = strBankId & IIf(Len(strBankId) > 0, " AND ", " ") & _
                      "A.CustID  = C.CUSTID AND " & _
                      "A.CustId=B.CustId And A.CancelFlag = 0 And " & _
                      "A.UCCode Is Not Null  " & _
                      IIf(Len(strChoose) = 0, "", " And ") & strChoose & " AND " & _
                      "A.AccountNO = C.AccountID AND A.CustID = D.CustID  AND " & _
                      "A.SERVICETYPE = D.SERVICETYPE " & _
                      " AND C.StopFlag<> 1 AND C.SnactionDate IS NOT NULL "
    '*********************************************************************************************

    ''  2003/08/14 改成 再將 mediabillno   資料取出 這裏這一次多取一個 BillNO 然後填入當mediabillno為null 值的時候，填入媒體單號
    '' 2003/08/17  必需再作一次，單純取出mediaBillno是null的值，並且將其填入新填
    '' errcode**********
    begintrnsErrCode = 2
    If intPara24 = 1 Then
        ''這一段直接取前述的主 SQL (strsql)  拿掉 group  ，加上 MediaBillno
        '' errcode**********
        '' 20050203 移除And A.ShouldAmt > 0的條件限制
        '#5055 沒提到,但順便一起改,過濾櫃台已收的資料 By Kin 2009/04/30
        '#5218 櫃台已收要用Nvl By Kin 2009/08/05
        '#5564 參考號7,8與CD013.PAYOK=1都是代表已收 By Kin 2010/05/17
        '#5683 增加挑選RealStopDate 與 PayKind By Kin 2010/08/06
        begintrnsErrCode = 2
        If IsTCB = False Then    '
            strMedioBillnoNull = "SELECT A.BillNO, A.MediaBillNO,A.AccountNO," & _
                        strMinPayKind & " , " & strMinRealStopDate & " , " & strStopYmWhere & _
                        " FROM " & TableOwnerName & "SO033 A," & TableOwnerName & "SO001 B ," & TableOwnerName & "SO002 D," & TableOwnerName & "SO106 C  " & _
                        "," & TableOwnerName & "CD013 " & _
                        " Where " & strBankId & IIf(Len(strBankId) > 0, " AND ", " ") & _
                        " A.CustID  = D.CUSTID AND " & _
                        " A.CustId=B.CustId And A.CancelFlag = 0 And " & _
                        " A.UCCode Is Not Null  And " & _
                        " A.UCCode=CD013.CodeNo And Nvl(CD013.REFNO,0) NOT IN(3,7) AND  " & _
                        " NVL(CD013.PAYOK,0) = 0 " & _
                        IIf(Len(strChoose) = 0, "", " And ") & strChoose & " AND " & _
                        " A.SERVICETYPE =D.SERVICETYPE  AND   " & _
                        " A.AccountNO = C.AccountID AND A.CustID = C.CustID  AND A.MediaBillNO IS NULL " & _
                        " AND C.StopFlag <> 1 AND C.SnactionDate Is Not Null " & _
                        " Group by A.BillNo,A.MediaBillNo,A.AccountNo "
        Else
            strMedioBillnoNull = "SELECT A.BillNO,A.MediaBillNO,A.AccountNO," & _
                        strMinPayKind & " , " & strMinRealStopDate & " , " & strStopYmWhere & _
                       " FROM " & TableOwnerName & "SO033 A," & TableOwnerName & "SO001 B ," & TableOwnerName & "SO002 D," & TableOwnerName & "SO106 C  " & _
                       "," & TableOwnerName & "CD013 " & _
                       " Where " & strBankId & IIf(Len(strBankId) > 0, " AND ", " ") & _
                       " A.CustID  = D.CUSTID AND " & _
                       " A.CustId=B.CustId And A.CancelFlag = 0 And " & _
                       " A.UCCode Is Not Null  And " & _
                       " A.UCCode=CD013.CodeNo And Nvl(CD013.REFNO,0) NOT INT (3,7) AND  " & _
                       " NVL(CD013.PAYOK,0) = 0 " & _
                       IIf(Len(strChoose) = 0, "", " And ") & strChoose & " AND " & _
                       " A.SERVICETYPE =D.SERVICETYPE  AND   " & _
                       " A.AccountNO = C.AccountID AND A.CustID = C.CustID  AND A.MediaBillNO IS NULL " & _
                       " AND C.StopFlag <> 1 AND C.SnactionDate Is Not Null " & _
                       " Group by A.BillNo,A.MediaBillNo,A.AccountNo "
        End If
        '' errcode**********
        '' 20050203 移除And A.ShouldAmt > 0的條件限制
        
        begintrnsErrCode = 3
        strChooseMultiAcc = "  A.CancelFlag = 0 AND " & _
                            "A.UCCode Is Not Null   " & _
                             IIf(Len(strChooseMultiAcc) = 0, "", " And ") & strChooseMultiAcc
        '' errcode**********
        begintrnsErrCode = 4
        If GetRS(rsMedioBillnoNull, strMedioBillnoNull, cn) Then
            Do While Not rsMedioBillnoNull.EOF
                
                If IsNull(rsMedioBillnoNull("MediaBillNo")) Then
                     strSequenceNumber = GetInvoiceNo2("SO033_MediaBillNo", cn)
                     cn.Execute ( _
                                    "UPDATE " & TableOwnerName & "SO033 A SET " & _
                                    "A.MediaBillNo  ='" & strSequenceNumber & "'  WHERE   " & _
                                    "A.BillNo = '" & rsMedioBillnoNull("BillNO") & "'   AND  " & strChooseMultiAcc & " AND " & _
                                    "A.AccountNO ='" & rsMedioBillnoNull("AccountNO") & "' ")
                   
                End If
                rsMedioBillnoNull.MoveNext
             Loop
        End If
    '' errcode**********
        begintrnsErrCode = 5
        rsMedioBillnoNull.Close
        Set rsMedioBillnoNull = Nothing
    End If
    


    ''********************************************************************************************
  '' errcode**********
    begintrnsErrCode = 6
    Set rsTmp = cn.Execute(strsql)
    lngIndex = 1
    blnFirst = True
  '' errcode**********
    begintrnsErrCode = 7
    ''這一段取得只有 SO033 的篩選資料  ........
    '' 20050203 移除And A.ShouldAmt > 0的條件限制
    If intPara24 = 0 Then
        strChooseMultiAcc = "  A.CancelFlag = 0 AND " & _
                            "A.UCCode Is Not Null   " & _
                             IIf(Len(strChooseMultiAcc) = 0, "", " And ") & strChooseMultiAcc
    End If
    '' errcode**********
    begintrnsErrCode = 8
    If rsTmp.BOF Or rsTmp.EOF Then
        MsgBox "無資料！請重新操作！", vbInformation, "訊息！"
        blnNodata = True
        Unload Me
    Else
        blnNodata = False
        Dim strSN As String  '' 指定序列號
        Dim lngtotal As Long '' 累計標頭檔 Total Amount 欄位總出單金額
        lngCount = 0
        lngtotal = 0
        Dim lngDoubleMediaNO As Long '' 這個值記錄重覆出單的單子件數
        Dim rsDoubleMediaNO As New ADODB.Recordset    ' 這個   rs 用以 check  MediaBillNO是不是有值
        lngDoubleMediaNO = 0
 '' errcode**********
        begintrnsErrCode = 9
           ''開始逐筆收集單資料
           ''**********************************************************************************
        While Not rsTmp.EOF
            Dim strCardExpDate As String
            Dim iCED As Integer
            Dim dteCED As Date
            Dim strErrMessage As String
            'rsUpdUCCode.Filter = "BillNo='" & rsTmp("BillNo") & "' And Account"
       '' 這一段補上其他的篩選條件值，因為會有多筆 SO033  同編號  2003/07/21
                
      '' 2003/08/17  由於主sql段的語法裏面只能Group 客戶的帳號以及單據編號，
      '' 因此 重新再取一次客編以及相關的資訊
'' errcode**********
            begintrnsErrCode = 10
            
    ' Jacky 94/10/05
            Call GetRS(rsCustIDErr, "SELECT  SO001.CUSTNAME , SO033.CUSTID FROM " & TableOwnerName & "SO001," & TableOwnerName & "SO033 " & _
                        "WHERE SO033.CUSTID = SO001.CUSTID And SO033.CompCode = SO001.CompCode And SO033." & strMediaBillNOField & "='" & rsTmp("BillNO") & "' ", cn)
            '#5683 如果收視截止日小於畫面條件不要出帳 By Kin 2010/08/06
            If Not IsPayKindOK(rsTmp, GiDAccept.GetValue & "", giAll, 1) Then
                lngErrCount = lngErrCount + 1
                strErrMessage = "收視截止日不正確：單號　" & rsTmp("BillNo") & " 客編：" & rsCustIDErr("CUSTID") & _
                            " 客戶姓名　" & rsCustIDErr("custname") & _
                            " 收視截止日：" & rsTmp("RealStopDate") & " 繳付類別：現付制 "
                file(1).WriteLine strErrMessage
                GoTo NextLoop
                            
            End If
'' errcode**********
            begintrnsErrCode = 11
            If IsNull(rsTmp("AccountNO")) Then
                lngErrCount = lngErrCount + 1
                strErrMessage = "信用卡卡號空白 : 單號　" & rsTmp("BillNO") & " 客戶姓名 " & rsCustIDErr("custname")
                file(1).WriteLine strErrMessage
                GoTo NextLoop
            End If
'' errcode**********
            begintrnsErrCode = 12
            '#3531 抓取SO106的StopYM(最小值) By Kin 2007/11/01
            '#4018 改取SO106的最大值(增加Desc) By Kin 2008/07/18
            '#5494 藍新科技也抓SO106 By Kin 2010/02/02
            If frmCreditCardType = 3 Or frmCreditCardType = 5 Then
                strStopYmSQL = "Select Decode(Length(A.StopYm),5,Substr(A.StopYm,2,4)||'0'||Substr(A.StopYm,1,1)," & _
                                             "Substr(A.StopYm,3,4)||Substr(A.StopYm,1,2)) As CardExpDate " & _
                                        " From " & TableOwnerName & "SO106 A " & _
                                        "Where A.AccountID='" & rsTmp("AccountNo") & "' " & _
                                        "And A.CustId=" & rsCustIDErr("Custid") & " And A.StopFlag=0 " & _
                                        IIf(strBankId <> "", " And " & strBankId, "") & _
                                        " Order By CardExpDate Desc"
                '#5267 已經改用SO106了,所以不用在重取一次 By Kin 2010/04/20
'                If Not GetRS(rsSO106, strStopYmSQL, cn) Then Exit Function
'                If rsSO106.EOF Then
'                    strCardExpDate = "NULL"
'                Else
'                    strCardExpDate = Left(rsSO106("CardExpDate"), 4) & "/" & Right(rsSO106("CardExpDate"), 2)
'                End If
                strCardExpDate = Empty
                If rsTmp("CardExpDate") & "" <> "" Then
                    strCardExpDate = Left(rsTmp("CardExpDate"), 4) & "/" & Right(rsTmp("CardExpDate"), 2)
                End If
            Else
                strCardExpDate = Empty
                strCardExpDate = rsTmp("CardExpDate") & ""
                'strCardExpDate =  GetNullString(rsTmp("CardExpDate"), giStringV, giAccessDb)
                If UCase(strCardExpDate) = "NULL" Then strCardExpDate = ""
                If strCardExpDate <> "" Then
                    strCardExpDate = Right(strCardExpDate, 2) & Left(strCardExpDate, 4)
                End If
            End If
            If UCase(strCardExpDate) = "NULL" Or strCardExpDate = Empty Then
                lngErrCount = lngErrCount + 1
                strErrMessage = "信用卡資料不正確 : 單號　" & rsTmp("BillNO") & " 客戶姓名 " & rsCustIDErr("custname")
                file(1).WriteLine strErrMessage
                GoTo NextLoop
            End If
            
'' errcode**********
            begintrnsErrCode = 12.5
            '#3531 如果是新格式使用SO106.StopYm,不是則使用原來的資料 By Kin 2007/11/01
            If frmCreditCardType <> 3 And frmCreditCardType <> 5 Then
                strCardExpDate = rsTmp("CardExpDate")
                begintrnsErrCode = 12.51
                iCED = IIf(Len(strCardExpDate) = 5, 1, 2)
                strCardExpDate = Left(strCardExpDate, 4) & "/" & Right(strCardExpDate, iCED)
                'strCardExpDate = Right(strCardExpDate, 4) & "/" & Mid(strCardExpDate, 1, iCED)
            End If
            dteCED = CDate(strCardExpDate)
            begintrnsErrCode = 12.52
            If dteCED < CDate(Format(GiDAccept.Text, "YYYY/MM")) Then
                lngErrCount = lngErrCount + 1
                begintrnsErrCode = 12.53
                On Error Resume Next
                If Not rsCustIDErr.EOF Then
                    strErrMessage = "信用卡過期 : 單號　" & rsTmp.Fields("Billno") & " 客戶姓名　" & rsCustIDErr("custname") & " 信用卡號 " & rsTmp.Fields("AccountNO") & " 到期日　" & _
                                    IIf(frmCreditCardType <> 3, rsTmp.Fields("CardExpDate"), strCardExpDate)
                                    '#3531 判斷使用那一種格式 By Kin 2007/11/01
                Else
                    strErrMessage = rsTmp.Source
                End If
                begintrnsErrCode = 12.54
                file(1).WriteLine strErrMessage
                '#4090 如果有勾選也要產生資料，如果沒有勾選則直接跳至下一筆 By Kin 2008/09/19
                If chkDuteDate.Value = 0 Then GoTo NextLoop
            End If
'' errcode**********
            begintrnsErrCode = 12.6
            '#3531 測試不OK，金額為0的也要算進去 By Kin 2008/05/22
            If rsTmp("ShouldAmt") <= 0 Then
                'strErrMessage = "負項金額 : 單號　" & rsTmp.Fields("Billno") & " 客戶姓名　" & rsCustIDErr("custname") & "    金額　" & rsTmp("ShouldAmt")
                strErrMessage = "負項或是金額等於零 : 單號　" & rsTmp.Fields("Billno") & " 客戶姓名　" & rsCustIDErr("custname") & "    金額　" & rsTmp("ShouldAmt")
                file(1).WriteLine strErrMessage
                lngErrCount = lngErrCount + 1
                GoTo NextLoop
            End If

            begintrnsErrCode = 14
            If IsTCB = True Then
                Dim rsBudget As ADODB.Recordset
                Set rsBudget = GetBugetData(rsTmp("CitemCode"), rsTmp("ServiceType"), rsTmp("PTCode"), rsTmp("ShouldAmt"), IIf(IsNull(rsTmp("TCBbudget")), 0, rsTmp("TCBbudget")))
                If rsBudget.EOF Or rsBudget.BOF Then
                    lngErrCount = lngErrCount + 1
                    strErrMessage = "TCB業務相關對應收費檔  分期資料未設  " & _
                                              "Citemcode =" & rsTmp("CitemCode") & "、" & _
                                              "ServiceType=" & rsTmp("ServiceType") & "、" & _
                                              "PTCode=" & rsTmp("PTCode") & "、" & _
                                              "ShouldAmt=" & rsTmp("ShouldAmt") & "、" & _
                                              "TCBbudget=" & rsTmp("TCBbudget") & " --   " & _
                                              "單號　" & rsTmp.Fields("Billno") & " 客戶姓名　" & rsCustIDErr("custname") & " 信用卡號 " & rsTmp.Fields("AccountNO")
                      file(1).WriteLine strErrMessage
                      GoTo NextLoop
                 End If
            End If
'' errcode**********
            begintrnsErrCode = 15
            lngCount = lngCount + 1
            '#3531 計算總金額 By Kin 2007/11/02
            lngShouldAmt = lngShouldAmt + Val(rsTmp("ShouldAmt"))
            With rsDefTmp
                .AddNew
                 ''  2003/10/08  如果是中國商銀的話改走另一個迴路
                Select Case frmCreditCardType
                Case 0
                    If IsTCB = False Then
                        .Fields("F110") = GetString(txtSpcNO.Text, 10, giLeft, False)
                        .Fields("F1118") = GetString(txtMerchantName.Text, 8, giLeft, False)
                        .Fields("F1937") = GetString(rsTmp("AccountNO"), 19, giLeft, True)
                        .Fields("F3845") = GetString(rsTmp("ShouldAmt"), 8, GIRIGHT, True)
                         lngtotal = lngtotal + rsTmp("ShouldAmt")
                        .Fields("F4653") = Space(8)
                        .Fields("F5455") = "0" & cboSet.ListIndex    ''  交易碼
                        .Fields("F5661") = Format(GiDAccept.Text, "EEMMDD")
                        If optmmyy.Value = True Then
                            '.Fields("F6265") = IIf(Len(rsTmp("CardExpDate")) = 6, Left(rsTmp("CardExpDate"), 2), "0" & Left(rsTmp("CardExpDate"), 1)) & Right(rsTmp("CardExpDate"), 2)
                            .Fields("F6265") = Right(rsTmp("CardExpDate"), 2) & Mid(strCardExpDate, 3, 2)
                        Else
                            .Fields("F6265") = Mid(strCardExpDate, 3, 2) & Right(rsTmp("CardExpDate"), 2)
                            '.Fields("F6265") = Right(rsTmp("CardExpDate"), 2) & IIf(Len(rsTmp("CardExpDate")) = 6, Left(rsTmp("CardExpDate"), 2), "0" & Left(rsTmp("CardExpDate"), 1))
                        End If
                        .Fields("F6668") = Space(3)
                        .Fields("F6984") = Space(16)
                        .Fields("F8590") = Space(6)
                        .Fields("F9196") = Space(6)
                       
                       
                        'If intPara24 = 0 Then
                        .Fields("F97112") = GetString(rsTmp("BillNO"), 16, giLeft, False)
                        
                       .Fields("F113118") = Space(6)
                       .Fields("F119") = chkCancel.Value
                       .Fields("F120") = cboMore.ListIndex
                    Else
                        Dim strTwDate As String
                        strTwDate = ""
                        strTwDate = Right((CInt(Right(rsTmp("CardExpDate"), 4)) - 1911), 2) & IIf(Len(rsTmp("CardExpDate")) = 6, Left(rsTmp("CardExpDate"), 2), "0" & Left(rsTmp("CardExpDate"), 1))
                        If IsTCBCard = True Then
                            .Fields("F06") = Format(GiDAccept.Text, "EEMMDD")
                            .Fields("F10") = GetString(lngCount, 4, GIRIGHT, True)
                            .Fields("F26") = GetString(rsTmp("AccountNO"), 16, giLeft, False)
                            .Fields("F29") = GetString(rsTmp("CVC2") & "", 3, GIRIGHT, True)
                            .Fields("F33") = strTwDate
                            If intPara24 = 0 Then
                                .Fields("F44") = GetString(GetBillNo_New(rsTmp("BillNO")), 11, giLeft, False)
                            Else
                                .Fields("F44") = GetString(rsTmp("BillNo"), 11, giLeft, False)
                            End If
                            .Fields("F50") = Format(GiDAccept.Text, "EEMMDD")
                            .Fields("F56") = Space(6)
                            .Fields("F58") = GetString(rsBudget("Period"), 2, GIRIGHT, True) '' 期數
                            .Fields("F65") = GetString(rsBudget("BudgetAmount"), 7, GIRIGHT, True) ''分期金額
                            .Fields("F72") = GetString(rsBudget("Amount"), 7, GIRIGHT, True)  ''總金額
                            .Fields("F79") = GetString(rsBudget("BudgetAmount"), 7, GIRIGHT, True)       ''頭款
                            .Fields("F86") = GetString(rsBudget("LastAmount"), 7, GIRIGHT, True)       ''尾款
                            .Fields("F91") = "10000"
                            Dim strMerENG As String
                            Dim strMerCHI As String
                            strMerENG = txtSpcNO.Text & Space(50)
                            strMerCHI = txtSpcNO.Text & Space(80)
                            .Fields("F116") = StrConv(LeftB(StrConv(StrConv(strMerENG, vbWide), vbFromUnicode), 25), vbUnicode)
                            .Fields("F136") = StrConv(LeftB(StrConv(StrConv(strMerCHI, vbWide), vbFromUnicode), 20), vbUnicode)
                            
                            .Fields("F138") = "05"
                            .Fields("F180") = Space(62)
                            
                             lngtotal = lngtotal + rsTmp("ShouldAmt")   ' '  記錄總金額  填在檔尾
                        Else
                            .Fields("F15") = GetString(rsTmp("BillNO"), 16, giLeft, False)
                            .Fields("F16") = GetString(rsTmp("BillNO"), 15, giLeft, False)
                            .Fields("F31") = GetString(rsTmp("AccountNO"), 19, giLeft, False)
                            .Fields("F50") = GetString(rsTmp("CVC2") & "", 3, giLeft, False)
                                         ''改為西元年      2003/05/05
                                         ''.Fields("F53") = strTwDate    ''到期日
                            '.Fields("F53") = Right(rsTmp("CardExpDate"), 2) & IIf(Len(rsTmp("CardExpDate")) = 6, Left(rsTmp("CardExpDate"), 2), "0" & Left(rsTmp("CardExpDate"), 1))
                             .Fields("F53") = Right(rsTmp("CardExpDate"), 2) & Mid(strCardExpDate, 3, 2)
                                         '' 總金額改為  10 碼
                                         '' .Fields("F57") = GetString(rsBudget("Amount"), 12, GIRIGHT, True) '' 總金額
                            .Fields("F57") = GetString(rsBudget("Amount"), 10, GIRIGHT, True) '' 總金額
                                          ''期數預設改為   '00'
                            .Fields("F67") = "00"    '' 常設值
                            .Fields("F69") = GetString(rsBudget("Period"), 2, GIRIGHT, True)  '' 期數
                            .Fields("F71") = GetString(rsBudget("BudgetAmount"), 12, GIRIGHT, True)       ''頭款
                            .Fields("F83") = GetString(rsBudget("LastAmount"), 12, GIRIGHT, True)       ''尾款
                             lngtotal = lngtotal + rsBudget("Amount")   ' '  記錄總金額  填在檔尾
                        End If
                    End If
                Case 1    '' 2003/10/24 以下這一段是中國商銀的迴路
                            ''''特店代號：依業務有不同代號 , 尚未確定
                            ''''卡號：聯合信用卡時卡號只有11位: 先放4位 '4000',再放11位卡號,再放1位'0',組成16位 ; 其他卡皆16位
                            ''''交易碼 : 原始 ‘0 ‘+空白 , 經授權處理後會變為 ‘1‘+空白 , 強迫授權(二次授權)後會變為 ‘2‘+空白
                            ''''流水號:  該筆明細在檔案之第幾筆 , 絕對不能重復
                            ''''回應碼 : ‘1+空白+空白’及‘ 5+空白+空白’屬成功交易,其對應之回應訊息必定為APPROVED ; 除此之外之回應碼均屬授權失敗交易
                            ''''回應訊息 : ‘APPROVED’代表成功 ; ‘REFERRAL’代表CALL BANK,其他則代表失敗,最常見為TRANSACTION ERRO’代表資料錯誤 ; ‘INVALID CAPTURE’代表其他失敗 ; ‘DECLINE’代表其他失敗,詳細原因煩請參照”批次回應碼”一表
                    begintrnsErrCode = 16.1
                    .Fields("F113") = GetString(txtMerchantName.Text, 13, giLeft, False)  ''  ** 13特店代號(類似聯合:端末機代號)
                      ''** 16卡號
                    If Len(Trim(rsTmp("AccountNO"))) = 11 Then
                        .Fields("F1429") = "4000" & rsTmp("AccountNO") & "0":          begintrnsErrCode = 16.2
                    Else
                        .Fields("F1429") = GetString(rsTmp("AccountNO"), 16, giLeft, True):           begintrnsErrCode = 16.3
                    End If
                    .Fields("F3039") = GetString(rsTmp("ShouldAmt"), 10, GIRIGHT, True)  ''  ** 10金額
'                    .Fields("F4043") = IIf(Len(rsTmp("CardExpDate")) = 6, Left(rsTmp("CardExpDate"), 2),  "0" & Left(rsTmp("CardExpDate"), 1)) & Right(rsTmp("CardExpDate"), 2 )   ''   ** 4有效月年 (MMYY) (西元年)
                    .Fields("F4043") = Right(rsTmp("CardExpDate"), 2) & Mid(strCardExpDate, 3, 2)
                    .Fields("F4445") = "0" & Space(1)  '' ** 2交易碼   (0+空白)
                    ''                       .Fields("F4659") = GetString(lngCount, 14, GIRIGHT, True) ''** 14流水號(絕對不可重複)  (不足補0)
                    If intPara24 = 0 Then
                        .Fields("F4659") = GetString(GetBillNo_New(rsTmp("BillNO")), 14, GIRIGHT, True)  ''** 2003/10/13 這個欄位改成填單號
                    Else
                        .Fields("F4659") = GetString(rsTmp("BillNO"), 14, GIRIGHT, True)  ''** 2003/10/13 這個欄位改成填單號
                    End If
                    '' 以下為銀行回覆用的
                    .Fields("F6065") = Space(6)
                    .Fields("F6671") = Space(6)
                    .Fields("F7274") = Space(3)
                    .Fields("F7590") = Space(16)
                    .Fields("F9196") = Space(6)
                    .Fields("F97102") = Space(6)
                    lngtotal = lngtotal + rsTmp("ShouldAmt")
                Case 2    '' 2003/11/20 以下這一段是匯豐銀行的迴路
:                           begintrnsErrCode = 16.4
                    .Fields("F116") = GetString(rsTmp("AccountNO"), 16, giLeft, False)    ''卡號  不足，依系統規定補足16碼space
                    .Fields("F1717") = Space(1)
                    begintrnsErrCode = 16.5
                    .Fields("F1821") = IIf(Len(rsTmp("CardExpDate")) = 6, Left(rsTmp("CardExpDate"), 2), "0" & Left(rsTmp("CardExpDate"), 1)) & Right(rsTmp("CardExpDate"), 2)      ''  ** 有效期限  MBF,MMYY
                    .Fields("F2222") = Space(1):   begintrnsErrCode = 16.6
                    .Fields("F2334") = GetString(rsTmp("ShouldAmt") & "00", 12, GIRIGHT, True)   ''  ** 金額  不足補'0'後兩碼為角分
                    .Fields("F3535") = Space(1): begintrnsErrCode = 16.7
                    
                    .Fields("F3641") = "999999"  ''   核淮號碼   輸入'999999'or approvaled code
                    .Fields("F4242") = Space(1): begintrnsErrCode = 16.8
                    .Fields("F4344") = "99"   ''  回復碼   輸入'99'or '0'(if this is a refund)
                    .Fields("F4545") = Space(1): begintrnsErrCode = 16.9
                    .Fields("F4685") = GetString(rsTmp("BillNO"), 40, giLeft, False)     '' 商店自行運用 (so043.para24=1則取媒體單號
                    '' 計算相關筆數以及收費金額
                    If rsTmp("ShouldAmt") > 0 Then
                        lngTotalHSBC = lngTotalHSBC + rsTmp("ShouldAmt"): begintrnsErrCode = 17.3
                        lngToatalCountHSBC = lngToatalCountHSBC + 1: begintrnsErrCode = 17.2
                    Else
                        lngTotalHSBCN = lngTotalHSBCN + Abs(rsTmp("ShouldAmt")): begintrnsErrCode = 17
                        lngTotalCountHSBCN = lngTotalCountHSBCN + 1: begintrnsErrCode = 17.1
                    End If
                '3531 使用新格式 By Kin 2007/11/01
                Case 3
                    .Fields("SPCNO") = txtAuthorizationBatch.Text
                    .Fields("ACCOUNTID") = rsTmp("ACCountNO")
                    '取年之後兩碼例10/2010 取1010
                    .Fields("STOPYM") = Trim(Right(strCardExpDate, 2)) & Mid(strCardExpDate, 3, 2)      ' Trim(Str(Val(Left(strCardExpDate, 4)) - 1911))
                    .Fields("SHOULDAMT") = rsTmp("ShouldAmt")
                    .Fields("BILLNO") = rsTmp("BillNo")
                    .Fields("TRADETYPE") = Left(cboMore.Text, 1)
                    '#4033 原本取101~140全型字形,改為101~138後面補2個空白 By Kin 2008/08/14
                    .Fields("CHINESE") = GetString(StrConv(txtBillMemo.Text, vbWide) & StrConv(Space(38), vbWide), 38, giLeft) & Space(2)
                    
                '#5494 增加藍新科技格式 By Kin 2010/01/27
                Case 5
                    .Fields("RcdType") = "D"
                    .Fields("OrderNo") = GetString(rsTmp("BillNo") & "", 30, giLeft)
                    .Fields("OrdItem") = Space(9)
                    '原則上分期期數一定會是空白
                    If rsTmp("TCBbudget") & "" = "" Then
                        .Fields("Period") = Space(2)
                    Else
                        .Fields("Period") = GetString(rsTmp("TCBbudget") & "", 2, GIRIGHT, True)
                    End If
                    .Fields("Status") = Space(1)
                    .Fields("Amt") = GetString(FmtAmt(rsTmp("ShouldAmt") & ""), 13, GIRIGHT, True)
                    .Fields("CardNo") = GetString(rsTmp("AccountNo") & "", 16, GIRIGHT, True)
                    If .Fields("CVC2") & "" <> "" Then
                        .Fields("CVC2") = GetString(rsTmp("CVC2") & "", 3, GIRIGHT, True)
                    Else
                        .Fields("CVC2") = Space(3)
                    End If
                    .Fields("ValidDate") = GetString(Replace(strCardExpDate, "/", ""), 6, GIRIGHT, True)
                    .Fields("AuthNo") = Space(6)
                    .Fields("Prc") = Space(5)
                    .Fields("Src") = Space(5)
                    .Fields("RetBank") = Space(2)
                    .Fields("BatNum") = Space(9)
                    .Fields("AutoFlag") = "1"
                    .Fields("hold") = Space("91")
                '#5609 增加聯合信用卡批次格式 By Kin 2010/06/07
                Case 6
'                    .Fields.Append "SPCNO", adBSTR, 10, adFldIsNullable         '商店代碼
'                    .Fields.Append "MERCHANTNAME", adBSTR, 8, adFldIsNullable   '端末機代號
'                    .Fields.Append "ACCOUNTNO", adBSTR, 16, adFldIsNullable     '卡號
'                    .Fields.Append "ACCOUNTNOEX", adBSTR, 3, adFldIsNullable    '卡號延伸碼
'                    .Fields.Append "STOPYM", adBSTR, 4, adFldIsNullable         '有效期限
'                    .Fields.Append "AMT", adBSTR, 10, adFldIsNullable           '交易金額
'                    .Fields.Append "AUTHNO", adBSTR, 8, adFldIsNullable         '授權碼
'                    .Fields.Append "SETNO", adBSTR, 2, adFldIsNullable          '交易碼
'                    .Fields.Append "ACCEPTDATE", adBSTR, 6, adFldIsNullable     '送件日期
'                    .Fields.Append "BILLNO", adBSTR, 16, adFldIsNullable        'USER DEFINE
'                    .Fields.Append "ADDITION", adBSTR, 30, adFldIsNullable      'Additional field
'                    .Fields.Append "ACCINFO", adBSTR, 40, adFldIsNullable       '卡人資訊
                    .Fields("SPCNO") = GetString(txtSpcNO, 10, giLeft)
                    .Fields("MERCHANTNAME") = GetString(txtMerchantName.Text, 8, giLeft, False)
                    .Fields("ACCOUNTNO") = GetString(rsTmp("AccountNO"), 16, giLeft, True)
                    .Fields("ACCOUNTNOEX") = "000"
'                    .Fields("STOPYM") = Right(rsTmp("CardExpDate"), 2) & IIf(Len(rsTmp("CardExpDate")) = 6, Left(rsTmp("CardExpDate"), 2), "0" & Left(rsTmp("CardExpDate"), 1))
                    .Fields("STOPYM") = Right(Left(rsTmp("CardExpDate") & "", 4), 2) & Right(rsTmp("CardExpDate") & "", 2)
                    .Fields("AMT") = GetString(rsTmp("ShouldAmt"), 10, GIRIGHT, True)
                    .Fields("AUTHNO") = GetString(rsTmp("CVC2") & "", 8, giLeft)
                    .Fields("SETNO") = Left(cboSet.Text, 2)
                    .Fields("ACCEPTDATE") = Format(GiDAccept.Text, "EEMMDD")
                    .Fields("BILLNO") = GetString(rsTmp("BillNo") & "", 16, giLeft)
                    .Fields("ADDITION") = Space(30)
                    If txtBillMemo.Text = "" Then
                        .Fields("ACCINFO") = String(40, StrConv(" ", vbWide))
                    Else
                        .Fields("ACCINFO") = Left(StrConv(txtBillMemo.Text, vbWide) & _
                                        String(40, StrConv(" ", vbWide)), 40)
                    End If
                        
                End Select
                    .Update
            End With
                        
            '**********************************************************************************************************************************
            '#3417 如果CD013.REFNO有包含到4,則更新UCCode與UCName By Kin 2007/12/05
            '#4388 增加更新異動人員與異動時間 By Kin 2009/04/30
            '#5218 櫃台已收要用Nvl By Kin 2009/08/05
            If blnUpdUCCode Then
              
                '#3892 更新速度變慢,調整語法,原本用Exists,現在改用= By Kin 2008/06/03
                '#4037 會將所有相符的媒体單號都Upd掉,改成RowId In(...) By Kin 2008/09/11
                '#5267 WHERE條件SO002A改用SO106 BY Kin 2010/04/20
                '#5564 參考號7,8與CD013.PAYOK=1都是代表已收 By Kin 2010/05/17
                strUpdUCCode = "UPDATE " & TableOwnerName & "SO033  Set UCCode=" & strUCCode & ",UCName='" & strUCName & "'" & _
                                ",UPDEN='" & objStorePara.uUPDEN & "',UPDTIME='" & objStorePara.uUPDTime & "' " & _
                            " Where RowId  In (Select A.RowId From " & _
                                            TableOwnerName & "SO001 B ," & _
                                            TableOwnerName & "SO002 D," & _
                                            TableOwnerName & "SO106 C," & _
                                            TableOwnerName & "SO033 A," & _
                                            TableOwnerName & "CD013 " & _
                                            " Where " & strUCCodeWhere & _
                                            " And A." & strMediaBillNOField & "='" & rsTmp("BillNo") & "'" & _
                                            " And A.AccountNO='" & rsTmp("AccountNO") & "'" & _
                                            " And A.UCCode=CD013.CodeNo And Nvl(CD013.REFNO,0) NOT IN (3,7) AND  " & _
                                            " NVL(CD013.PAYOK,0) = 0 " & _
                                            IIf(rsTmp("CardExpDate") & "" <> "", " And " & _
                                            "DECODE(LENGTH(C.STOPYM),5,SUBSTR(C.STOPYM,2,4) || '0' || SUBSTR(C.STOPYM,1,1),SUBSTR(C.STOPYM,3,4)||SUBSTR(C.STOPYM,1,2))  = " & rsTmp("CardExpDate"), "") & _
                                            IIf(rsTmp("CVC2") & "" <> "", " And C.CVC2=" & rsTmp("CVC2"), "") & ")"
                cn.Execute strUpdUCCode
            End If
            '**********************************************************************************************************************************
            
NextLoop:
            rsTmp.MoveNext
        Wend
        '' 以下取得 header 字串
        Select Case frmCreditCardType
        Case 0
'            strHeader = "Header"
'            strHeader = strHeader & GetString(lngCount, 8, GIRIGHT, True)
'            strHeader = strHeader & Space(106)
            strHeader = ""
        Case 1
            ''' 2003/10/08 以下這一段是中國商銀的迴路
            strHeader = ""
            strHeader = GetString(txtSpcNO.Text, 6, giLeft, False) & Format(Date, "eemmdd") & _
                        GetString(lngCount, 10, GIRIGHT, True) & GetString(lngtotal, 12, GIRIGHT, True) & _
                        Space(10) & Space(12) & Space(10) & Space(12)
        Case 2
            begintrnsErrCode = 17.5
            strHeader = "": begintrnsErrCode = 17.6
            strHeader = GetString(Left(txtMerchantName.Text, 8), 8, giLeft, False) & Space(1) & _
                        GetString(Left(txtSpcNO.Text, 15), 15, giLeft, False) & Space(1) & _
                        GetString(lngTotalCountHSBCN + lngToatalCountHSBC, 5, GIRIGHT, True) & Space(1) & _
                        GetString((lngTotalHSBC + lngTotalHSBCN) & "00", 12, GIRIGHT, True) & Space(1) & _
                        GetString(lngTotalCountHSBCN, 5, GIRIGHT, True) & Space(1) & _
                        GetString((lngTotalHSBCN) & "00", 12, GIRIGHT, True) & Space(1) & _
                        GetString((lngTotalHSBC - lngTotalHSBCN) & "00", 12, GIRIGHT, True)
        '#3531 填入表頭 By Kin 2007/11/01
        Case 3
            '#3531 測試不OK,要填入183個0，不能是空白 By Kin 2007/11/22
            strHeader = ""
            strHeader = "H" & GetString(txtSpcNO, 15) & _
                        "E" & Space(183)
        '#5609 增加聯合信用卡批次格式 By Kin 2010/06/03
        Case 6
            strHeader = ""
            strHeader = "HEADER" & Format(GiDAccept.Text, "EEMMDD") & _
                        GetString(lngCount, 10, GIRIGHT, True) & _
                        GetString(lngtotal, 12, GIRIGHT, True) & _
                        Space(46)
            
        End Select
        '' 2003/01/23  如果是中華商銀的 在這一行取檔尾的記錄  ........
        If IsTCB = True Then
             Dim strEndLine As String
             strEndLine = "T"
             strEndLine = strEndLine & GetString(lngCount, 6, GIRIGHT, True)
             
             ''  總金額如果是他行卡，改為 十 碼 ，因此以下這一行分別依長度不同改為在    If IsTCBCard = True Then  的區塊裏頭作
            ''strEndLine = strEndLine & GetString(lngtotal, 12, GIRIGHT, True)
            If IsTCBCard = True Then
                strEndLine = strEndLine & GetString(lngtotal, 12, GIRIGHT, True)
                strEndLine = strEndLine & Format(Date, "EEMMDD")
            Else
                strEndLine = strEndLine & GetString(lngtotal, 10, GIRIGHT, True)
                strEndLine = strEndLine & Format(Date, "YYYYMMDD")
            End If
            strEndLine = strEndLine & Format(Now, "HHMMSS")
            If IsTCBCard = False Then
                Dim rsCD018Tmp As New ADODB.Recordset
                Dim strNN As String
                strNN = ""
                '                                                     rsCD018Tmp.CursorLocation = adUseClient
                '                                                     rsCD018Tmp.Open "SELECT  *  FROM " & TableOwnerName & "CD018Tmp   ORDER BY BATCHNODATE  DESC , BATCHNO DESC  ", cn, adOpenKeyset, adLockReadOnly
                If Not GetRS(rsCD018Tmp, "SELECT  *  FROM " & TableOwnerName & "CD018Tmp   ORDER BY BATCHNODATE  DESC , BATCHNO DESC  ", cn) Then Exit Function
                If rsCD018Tmp.EOF Or rsCD018Tmp.BOF Then
                    strNN = "01"
                Else
                    If Format(Date, "yyyymmdd") = rsCD018Tmp("BatchNODate") Then
                        If rsCD018Tmp("BatchNO") < 9 Then
                            strNN = "0" & CStr(rsCD018Tmp("BatchNO") + 1)
                        Else
                            strNN = CStr(rsCD018Tmp("BatchNO") + 1)
                        End If
                    Else
                        strNN = "01"
                    End If
                End If
                rsCD018Tmp.Close
                Set rsCD018Tmp = Nothing
                strEndLine = strEndLine & Left(txtSpcNO.Text, 4) & "AU" & Format(Date, "YYMMDD") & strNN
                If lngCount > 0 Then
                    cn.Execute "INSERT INTO CD018Tmp (BatchNO, BatchNODate) values (" & CInt(strNN) & ",'" & Format(Date, "yyyymmdd") & "')"
                End If
            End If
        End If
        '#3531 中國信託新格式需要填入新格式 By Kin 2007/11/02
        If frmCreditCardType = lngCtrust2 Then
            Dim intSpcCount As Integer
            Dim intKeep As Integer
            strEndLine = ""
            strEndLine = "R001"
            intSpcCount = 18
            intKeep = 142 '保留欄位
            Select Case Left(cboMore.Text, 1)
                Case "S", "O"
                    strEndLine = strEndLine & GetString(CStr(lngCount), 5, GIRIGHT, True) & _
                                GetString(CStr(lngShouldAmt), 11, GIRIGHT, True) & "00" & String(36, "0") & Space(intKeep)
                Case "A"
                    strEndLine = strEndLine & String(intSpcCount, "0") & _
                                GetString(CStr(lngCount), 5, GIRIGHT, True) & _
                                GetString(CStr(lngShouldAmt), 11, GIRIGHT, True) & "00" & String(intSpcCount, "0") & Space(intKeep)
                Case "R"
                    strEndLine = strEndLine & String(intSpcCount * 2, "0") & _
                                GetString(CStr(lngCount), 5, GIRIGHT, True) & _
                                GetString(CStr(lngShouldAmt), 11, GIRIGHT, True) & "00" & Space(intKeep)
            End Select
        End If
        '#5494 藍新科技的表頭 By Kin 2010/02/02
        If frmCreditCardType = lngNewWeb Then
            strHeader = "H" & GetString(GiDAccept.GetValue, 8, GIRIGHT, True) & _
                GetString(txtAuthorizationBatch.Text, 2, GIRIGHT, True) & _
                GetString(rsDefTmp.RecordCount, 8, GIRIGHT, True) & _
                GetString(FmtAmt(CStr(lngShouldAmt)), 15, GIRIGHT, True) & Space(166)
                


        
        End If
    End If
        begintrnsErrCode = 17.7
        If IsTCB = True Then
            Call subWriteLine(strEndLine)
        Else
            Call subWriteLine(strEndLine)
        End If
            begintrnsErrCode = 17.8
        msgResult lngCount, lngErrCount, lngTime      '顯示執行結果
       
        BeginTran = True
        On Error Resume Next
        If rsSO106.State = adStateOpen Then rsSO106.Close
        If rsCD013.State = adStateOpen Then rsCD013.Close
        If rsUpdUCCode.State = adStateOpen Then rsUpdUCCode.Close
        rsTmp.Close
        Set rsTmp = Nothing
        Set rsSO106 = Nothing
        Set rsUpdUCCode = Nothing
        Set rsCD013 = Nothing
        
        
        
        CloseFS
     
    If lngDoubleMediaNO > 0 Then
    '     Call MsgBox("出單的資料有些是重覆的，請查閱記錄檔確認所需的資料內容，避免重覆出單  !!", vbExclamation + vbOKOnly, "重覆出單警示")
    '     Call Shell("NotePad " & txtErrPath.Text, vbNormalFocus)
    End If
    
    Exit Function
ChkErr:
    Call ErrSub(Me.Name, "BeginTran ：errCode = " & begintrnsErrCode & "---" & strBankId & ";ErrNumber:" & Err.Number & ";ErrDescription:" & Err.Description)
    CloseFS
'    MsgBox begintrnsErrCode
'    MsgBox "單號　" & rsTmp.Fields("Billno") & " 客戶姓名　" & rsCustIDErr("custname") & " 信用卡號 " & rsTmp.Fields("AccountNO")
    
End Function
'#5494 藍新格式的金額後2碼是小數點
Private Function FmtAmt(ByVal aShouldAmt As String) As String
  On Error GoTo ChkErr
    Dim sTmp As String
    If InStr(1, aShouldAmt, ".") Then
        sTmp = Left(aShouldAmt, InStr(1, aShouldAmt, "."))
        sTmp = sTmp & GetString(Mid(aShouldAmt, InStr(1, aShouldAmt, ".") + 1, 2), 2, giLeft, True)
        sTmp = Replace(sTmp, ".", "")
    Else
        sTmp = aShouldAmt & "00"
    End If
    FmtAmt = sTmp
    Exit Function
ChkErr:
    Call ErrSub(Me.Name, "FmtAmt")
End Function

Private Function subWriteLine(Optional pEndLine As String = "") As Boolean

  On Error GoTo ChkErr
  
    Dim varData As Variant
    Dim strData As String
    Dim lngLoop As Long
    Dim lngLoopi As Long
    Dim strContent As String

    subWriteLine = False
    
    With rsDefTmp
        If .BOF Or .EOF Then Exit Function
            .MoveFirst
         
            If IsTCB = False Then
                If strHeader <> Empty Then
                    file(0).WriteLine (strHeader)    ''於文字檔填入首筆
                End If
            End If
         
            For lngLoop = 0 To .RecordCount - 1
                 strContent = ""
                 Dim i As Long
                 
                 ''2003/10/24  根據信用卡別，則設定為為其格式
                Select Case frmCreditCardType
                    Case 0
                        If IsTCB = False Then
                            i = 15
                        Else
                            If IsTCBCard = True Then i = 17 Else i = 8
                        End If
                    Case 1
                        i = 11
                    Case 2
                        i = 10
                    '#5494 藍新科技 By Kin 2010/02/03
                    '#5609 聯合信用卡批次格式 By Kin 2010/06/07
                    Case 5, 6
                        i = rsDefTmp.Fields.Count - 1
                    
                End Select
                '#3531 新增中國信託新格式(明細) By Kin 2007/11/02
                If frmCreditCardType = lngCtrust2 Then
                    strContent = "T" & GetString(.Fields("SPCNO") & "", 5, GIRIGHT, True) & _
                                    GetString(.Fields("ACCOUNTID") & "", 19) & .Fields("STOPYM") & _
                                    GetString(.Fields("ShouldAMT") & "", 11, GIRIGHT, True) & "00" & _
                                    Space(2) & GetString(.Fields("BillNo") & "", 15) & Space(6) & _
                                    .Fields("TRADETYPE") & "" & Space(34) & .Fields("CHINESE") & Space(60)
                                        
                Else
                    For lngLoopi = 0 To i
                        strContent = strContent & rsDefTmp(lngLoopi)
                    Next
                End If
                file(0).WriteLine (strContent)
                rsDefTmp.MoveNext
                DoEvents
            Next
            '#3531 如果是中國信託新格式也要寫入表尾 By Kin 2007/11/02
            If IsTCB = True Or frmCreditCardType = lngCtrust2 Then
                file(0).WriteLine (pEndLine)
            End If
            
    End With
    subWriteLine = True
    On Error Resume Next
    rsDefTmp.Close
    Set rsDefTmp = Nothing
    
    Exit Function
ChkErr:
    subWriteLine = False
    Call ErrSub(Me.Name, "subWriteLine")
End Function

Private Sub ScrToRcd()
    On Error GoTo ChkErr
    Dim LogFile As TextStream
    Set LogFile = fso.CreateTextFile(strPath & "\" & strPrgName & "Out.log", True)
    '#5494 多增加每月扣款日 By Kin 2010/02/03
    With LogFile
        .WriteLine (txtSpcNO.Text) '信用卡扣帳特店名稱
        .WriteLine (txtMerchantName.Text) '信用卡扣帳特店名稱
        .WriteLine (txtDataPath)         '資料檔路徑
        .WriteLine (txtErrPath.Text)     ' 問題參考檔路徑
        .WriteLine (cboMonthDay.Text)
    End With
        
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "ScrToRcd")
End Sub
Private Function IsDataOk() As Boolean
    Dim strerrmsg As String
    
    If IsTCB = False Then
        Select Case frmCreditCardType
            Case 0
                If ChkFields(True, True, True, True) = False Then Exit Function
            Case 1
                If ChkFields(True, True, False, False) = False Then Exit Function
            Case 2
                If ChkFields(True, True, False, False, True) = False Then Exit Function
            Case 3
                If ChkFields(True, False, False, True, True) = False Then Exit Function
            '#5494 增加藍新科技的判斷 By Kin 2010/02/03
            Case 5
                If ChkFields(True, False, False, False, False, True) = False Then Exit Function
            '#5609 增加聯合信用卡批次檢查 By Kin 2010/06/03
            Case 6
                If ChkFields(True, True, True, True) = False Then Exit Function
        End Select
    End If

    If Len(GiDAccept.Text) = 0 Then strerrmsg = "送件日期": GiDAccept.SetFocus: GoTo Warning
    If Len(txtDataPath.Text) = 0 Then strerrmsg = "資料檔路徑": txtDataPath.SetFocus: GoTo Warning
    If Len(txtErrPath.Text) = 0 Then strerrmsg = "問題參考檔路徑": txtErrPath.SetFocus: GoTo Warning
    IsDataOk = True
    
    Exit Function
Warning:
    MsgBox "欄位 " & strerrmsg & " 必需有值", vbOKOnly + vbInformation, "訊息"
    Exit Function
ChkErr:
    Call ErrSub(Me.Name, "IsDataOk")
End Function
Private Function ChkFields(ByVal blnSpcNO As Boolean, ByVal blnMerchantName As Boolean, _
                           ByVal blnSet As Boolean, ByVal blnMore As Boolean, _
                           Optional blnAuthorizationBatch As Boolean = False, _
                           Optional ByVal blnAcceptDate As Boolean = False) As Boolean
                                   
    Dim strerrmsg As String
    ChkFields = False
    If blnSpcNO = True Then
        If Len(txtSpcNO.Text & "") = 0 Then strerrmsg = "商店代號": txtSpcNO.SetFocus: GoTo Warning
    End If
    If blnMerchantName = True Then
        If Len(txtMerchantName.Text & "") = 0 Then strerrmsg = "端未機代號": txtMerchantName.SetFocus: GoTo Warning
    End If
    If blnSet = True Then
        If Len(cboSet.Text) = 0 Then strerrmsg = "交易碼": cboSet.SetFocus: GoTo Warning
    End If
    If blnMore = True Then
        If Len(cboMore.Text) = 0 Then strerrmsg = "交易旗號": cboMore.SetFocus: GoTo Warning
    End If
    If blnAuthorizationBatch = True Then
        If Len(txtAuthorizationBatch.Text) = 0 Then strerrmsg = "授權批次": txtAuthorizationBatch.SetFocus: GoTo Warning
    End If
    If blnAcceptDate Then
        If Len(GiDAccept.GetValue) = 0 Then strerrmsg = "送件日期": GiDAccept.SetFocus: GoTo Warning
    End If
    
    ChkFields = True

Exit Function
Warning:
    MsgBox "欄位 " & strerrmsg & " 必需有值", vbOKOnly + vbInformation, "訊息"
    Exit Function
ChkErr:
    Call ErrSub(Me.Name, "IsDataOk")
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call CloseFS
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
        rsTmp.Close
        Set rsTmp = Nothing
End Sub

Private Sub GiDAccept_GotFocus()
    If Len(GiDAccept.Text) = 0 Then GiDAccept.Text = Format(Date, "EE/MM/DD")
End Sub

Private Sub txtAuthorizationBatch_KeyPress(KeyAscii As Integer)
  On Error Resume Next
    If KeyAscii = 8 Then Exit Sub
    If frmCreditCardType = 3 Or frmCreditCardType = 5 Then
        If KeyAscii >= 48 And KeyAscii <= 57 Then Exit Sub
    End If
    KeyAscii = 0
End Sub

Private Sub txtMerchantName_KeyPress(KeyAscii As Integer)
  On Error Resume Next
    
End Sub

Public Function GetInvoiceNo3(StrTableName As String, Optional Conn As ADODB.Connection) As String
  On Error GoTo ChkErr
    Dim strSeq As String
    Dim strInv
    strSeq = "S_" & StrTableName
    ''If Conn Is Nothing Then Set Conn = gcnGi
    strInv = GetRsValue("SELECT Ltrim(To_Char(" & strSeq & ".NextVal, '09999999')) FROM Dual", Conn)
    GetInvoiceNo3 = strInv
    Exit Function
ChkErr:
  ErrSub Me.Name, "GetInvoiceNo3"
End Function
Public Function BatchUpdateMediaBillNo(rsSource As ADODB.Recordset, strChoose33 As String, _
cn As ADODB.Connection) As Boolean
'On Error GoTo ChkErr
'Dim lngAffected As Long
'Dim strMediaBillNo As String
'        With rsSource
'            If .RecordCount > 0 Then .MoveFirst
'            Do While Not .EOF
'                    strMediaBillNo = ""
'                    If Not GetMediaBillNo(rsSource("CustId"), strMediaBillNo, cn) Then Exit Function
'                    If Not ExecuteCommand("Update SO033 A Set MediaBillNo = '" & strMediaBillNo & "' Where BillNo = '" & rsSource("BillNo") & "' And CustId = " & rsSource("CustId") & IIf(Len(strChoose33) = 0, "", " And ") & strChoose33, cn, lngAffected) Then Exit Function
'                    On Error Resume Next
'                    .Fields("MediaBillNo") = strMediaBillNo
'                    .Update
'                    On Error GoTo ChkErr
'                    .MoveNext
'            Loop
'        End With
'      BatchUpdateMediaBillNo = True
'Exit Function
'ChkErr:
'            ErrSub FormName, "BatchUpdateMediaBillNo"
End Function

Public Function GetMediaBillNo(ByRef lngCustId As Long, ByRef strMediaBillNo As String, cn As ADODB.Connection) As Boolean
'''On Error GoTo ChkErr

'lngCustId = Val(GetInvoiceNo3("SO033_MediaBillNo", cn) & "")
'strMediaBillNo = Format(Date, "yymm") & Format(lngCustId, "0000000")
'GetMediaBillNo = True
'Exit Function
'ChkErr:
'ErrSub FormName, "GetMediaBillNo"
End Function

Private Function GetBugetData(ByVal intCitemCODE As Integer, _
                                                  ByVal strServiceType As String, _
                                                  ByVal intPTCode As Integer, _
                                                  ByVal intAmount As Integer, _
                                                  ByVal intPeriod As Integer) As ADODB.Recordset

    Dim strsql As String
    Dim rs As New ADODB.Recordset
  On Error GoTo ChkErr
        strsql = "SELECT * From " & TableOwnerName & "CD019B Where " & _
                    "Citemcode=" & intCitemCODE & " AND  ServiceType='" & strServiceType & "' AND " & _
                    "PTCode =" & intPTCode & " AND  Amount = " & intAmount & " AND Period = " & intPeriod
        If rs.State = 1 Then rs.Close
    '    rs.CursorLocation = adUseClient
    '    rs.Open strsql, cn, adOpenKeyset, adLockReadOnly
        If Not GetRS(rs, strsql, cn) Then Exit Function
        Set GetBugetData = rs
Exit Function
ChkErr:
    ErrSub Me.Name, "GetBugetData"
End Function

Private Function CheckOutDate() As Boolean

    Dim aPath() As String
    Dim varAnayPath
    Dim strAnayPath As String

  On Error GoTo ChkErr

    If frmCreditCardType = 2 Then
        aPath = Split(txtDataPath.Text, "\")
        For Each varAnayPath In aPath
            strAnayPath = varAnayPath
        Next
        aPath = Split(strAnayPath, ".")
        strAnayPath = aPath(0)
        CheckOutDate = True
        If Right(strAnayPath, 6) <> Format(Date, "eemmdd") Then
            MsgBox "輸出檔名的日期非今日，請進行調整    !!   "
            CheckOutDate = False
        Else
            If intBatchNumber = CInt(Mid(strAnayPath, 2, Len(strAnayPath) - 7)) Then
                MsgBox "輸出檔批次重覆，請調整檔名批次號碼    !!   "
                CheckOutDate = False
            End If
        End If
    Else
        CheckOutDate = True
    End If

    Exit Function
ChkErr:
    ErrSub Me.Name, "CheckOutDate"
End Function

