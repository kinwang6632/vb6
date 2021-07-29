VERSION 5.00
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCitiBank 
   BorderStyle     =   1  '單線固定
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9240
   BeginProperty Font 
      Name            =   "新細明體"
      Size            =   9.75
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCitiBank.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   9240
   StartUpPosition =   1  '所屬視窗中央
   Begin MSComDlg.CommonDialog cnd1 
      Left            =   5520
      Top             =   180
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
      Height          =   3975
      Left            =   165
      TabIndex        =   16
      Top             =   15
      Width           =   8895
      Begin VB.TextBox txtBatch 
         Height          =   345
         Left            =   3570
         MaxLength       =   5
         TabIndex        =   5
         Text            =   "01"
         Top             =   1410
         Width           =   900
      End
      Begin VB.ComboBox cboBillType 
         ForeColor       =   &H000000FF&
         Height          =   315
         ItemData        =   "frmCitiBank.frx":0442
         Left            =   2040
         List            =   "frmCitiBank.frx":044F
         Style           =   2  '單純下拉式
         TabIndex        =   0
         Top             =   240
         Width           =   1665
      End
      Begin VB.CheckBox chkDuteDate 
         Caption         =   "信用卡過期資料一併產生"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   2040
         TabIndex        =   9
         Top             =   2265
         Width           =   2535
      End
      Begin VB.TextBox txtSpcNO 
         Height          =   300
         Left            =   2025
         MaxLength       =   15
         TabIndex        =   1
         Top             =   705
         Width           =   1830
      End
      Begin VB.TextBox txtMerchantName 
         Height          =   300
         Left            =   6060
         MaxLength       =   8
         TabIndex        =   2
         Text            =   "TBGB"
         Top             =   705
         Width           =   2460
      End
      Begin VB.TextBox txtStatement 
         Height          =   345
         Left            =   2025
         TabIndex        =   8
         Top             =   1800
         Width           =   6465
      End
      Begin VB.TextBox txtDiscountRate 
         Height          =   345
         Left            =   2025
         MaxLength       =   5
         TabIndex        =   4
         Text            =   "0"
         Top             =   1410
         Width           =   900
      End
      Begin VB.TextBox txtPayDate 
         Height          =   345
         Left            =   6075
         MaxLength       =   2
         TabIndex        =   7
         Text            =   "00"
         Top             =   1410
         Width           =   330
      End
      Begin VB.ComboBox cboSet 
         Height          =   315
         ItemData        =   "frmCitiBank.frx":047D
         Left            =   2025
         List            =   "frmCitiBank.frx":048D
         TabIndex        =   3
         Top             =   1050
         Width           =   2460
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
         Height          =   1275
         HelpContextID   =   2
         Left            =   285
         TabIndex        =   17
         Top             =   2475
         Width           =   8160
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
            Left            =   3015
            TabIndex        =   10
            ToolTipText     =   "請輸入字檔之路徑及檔名！"
            Top             =   300
            Width           =   4095
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
            Left            =   7320
            TabIndex        =   13
            Top             =   780
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
            Left            =   7320
            TabIndex        =   11
            Top             =   315
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
            Left            =   3015
            TabIndex        =   12
            Top             =   765
            Width           =   4095
         End
         Begin VB.Label lblDatapath 
            AutoSize        =   -1  'True
            Caption         =   "資料檔位置（路徑 )"
            Height          =   195
            Left            =   1290
            TabIndex        =   19
            Top             =   375
            Width           =   1665
         End
         Begin VB.Label lblErrpath 
            AutoSize        =   -1  'True
            Caption         =   "問題參考檔位置（路徑 + 名稱）"
            Height          =   195
            Left            =   330
            TabIndex        =   18
            Top             =   810
            Width           =   2640
         End
      End
      Begin Gi_Date.GiDate gdaSystemDate 
         Height          =   345
         Left            =   6060
         TabIndex        =   6
         Top             =   1035
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   609
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
      Begin VB.Label lblBatch 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "批號"
         Height          =   195
         Left            =   3015
         TabIndex        =   29
         Top             =   1500
         Width           =   390
      End
      Begin VB.Label Label6 
         Caption         =   "輸出格式"
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   1170
         TabIndex        =   28
         Top             =   330
         Width           =   825
      End
      Begin VB.Label lblSpcNO 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "商店代號"
         Height          =   195
         Left            =   1155
         TabIndex        =   27
         Top             =   780
         Width           =   780
      End
      Begin VB.Label Label5 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "系統日期"
         Height          =   195
         Left            =   5190
         TabIndex        =   26
         Top             =   1155
         Width           =   780
      End
      Begin VB.Label Label9 
         Caption         =   "(DD)"
         Height          =   240
         Left            =   6465
         TabIndex        =   25
         Top             =   1440
         Width           =   1080
      End
      Begin VB.Label lblMerchantName 
         AutoSize        =   -1  'True
         Caption         =   "端末機代號"
         Height          =   195
         Left            =   5025
         TabIndex        =   24
         Top             =   780
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "客戶帳單內容"
         Height          =   195
         Left            =   750
         TabIndex        =   23
         Top             =   1890
         Width           =   1170
      End
      Begin VB.Label Label3 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "折扣率"
         Height          =   195
         Left            =   1350
         TabIndex        =   22
         Top             =   1515
         Width           =   585
      End
      Begin VB.Label Label2 
         Caption         =   "請款日期"
         Height          =   210
         Left            =   5190
         TabIndex        =   21
         Top             =   1515
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "結帳種類"
         Height          =   195
         Left            =   1140
         TabIndex        =   20
         Top             =   1155
         Width           =   780
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "F2.開始"
      Height          =   375
      Left            =   315
      TabIndex        =   14
      Top             =   4140
      Width           =   1275
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&X)"
      Height          =   375
      Left            =   7665
      TabIndex        =   15
      Top             =   4125
      Width           =   1275
   End
End
Attribute VB_Name = "frmCitiBank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
''
Private cn As New ADODB.Connection
Private rsTmp As New ADODB.Recordset     '' 取出實體資料庫 的資料
Private rsDefTmp As New ADODB.Recordset    ''存文字檔的記憶體 recordset
Private rsCD013 As New ADODB.Recordset
Private strPrgName As String  ''程式名稱
Private strPath As String     ''GICMIS1.INI路徑
Private strDataName As String  ''檔案名稱
Private strHeader As String   ''記錄標頭檔的文字
Private strChoose As String  ''
Private lngErrCount As Long   '' 記錄錯誤筆數
Public blnSameFile As Boolean   '' 記錄不同的卡別是不是同一個特店名稱與代碼，是的話則在同一個文字檔  append 進去
Dim strBankId As String
Dim lngCount As Long       '' 記錄所有文字檔輸出的總筆數

Public intPara24 As Integer   ''  記錄是否使用媒體多帳戶處理
Public strChooseMultiAcc As String   '' 這個參數用來承接篩選 SO033 所需的收費資料
Dim rsMedioBillnoNull As New Recordset   '' 這個參數用來承接篩選 SO033裏 mediobillno 是null 的資料
Private blnUpdUcCode As Boolean
Private strUCCode As String
Private strUCName As String



Private Sub cboBillType_Click()
    If cboBillType.ListIndex = 2 Then
        cboSet.ListIndex = 0
        txtBatch.Enabled = True
        txtBatch.Text = "01"
    Else
        cboSet.Text = ""
        txtBatch.Enabled = False
        txtBatch.Text = ""
    End If

End Sub

Private Sub chkDuteDate_Click()
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
    On Error GoTo ChkErr
    Dim s As String
    Dim intTxtFileIndex  As Integer
    Dim lngTime As Long
        lngTime = Timer
        If Not IsDataOk Then Exit Sub
        
        '********************************************************************************************************************************************************
        '#3417 電子檔匯出時,要填入未收原因(RefNo=4) By Kin 2007/12/05
        If Not GetRS(rsCD013, "Select * From " & TableOwnerName & "CD013 Where RefNo=4 And StopFlag<>1 Order By CodeNo Desc", cn, adUseClient, adOpenKeyset) Then Exit Sub
        If rsCD013.EOF Then
            blnUpdUcCode = False
        Else
            blnUpdUcCode = True
            strUCCode = rsCD013("CodeNo") & ""
            strUCName = rsCD013("Description") & ""
        End If
        '*********************************************************************************************************************************************************
        
        
        '' Call DefineRs   ''建立記憶體 recordset 結構
        s = txtDataPath.Text
        If Not CreateObject("Scripting.FileSystemObject").FolderExists(s) Then MsgBox "路徑 " & s & " 不存在!", vbExclamation: Exit Sub
        '錯誤的Log
        If OpenFile(txtErrPath, 0, True) = False Then Exit Sub
        '輸出格式
        'objStorePara.uProcText = txtDataPath.Text
        objStorePara.uProcErrText = txtErrPath.Text
        objStorePara.uChkType = cboBillType.ListIndex
        If cboBillType.ListIndex = 0 Then
            If Not subCitiBankFlow Then file(0).Close: Exit Sub
        ElseIf cboBillType.ListIndex = 1 Then
            If Not subHSBCFlow Then file(0).Close: Exit Sub
        '問題集2763  新增花旗批次格式
        ElseIf cboBillType.ListIndex = 2 Then
            If Not subCitiBankBatch Then file(0).Close: Exit Sub
        End If
        '' 將收集所得的資料填入文字檔
        file(0).Close
        Call ScrToRcd
        objStorePara.uUpdate = True
        
        If lngErrCount > 0 Then Call Shell("NotePad " & txtErrPath.Text, vbNormalFocus)
        Call msgResult(lngCount, lngErrCount, lngTime)      '顯示執行結果
        Unload Me

    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "cmdOK_Click")
End Sub

Private Function subCitiBankFlow() As Boolean
    On Error GoTo ChkErr
    Dim rsCD037SpcNO  As New ADODB.Recordset  '' 取得 SPCNO 為唯一值的記錄，以同樣SPCNO的卡別一次取出
    Dim rsCD037 As New ADODB.Recordset
    Dim s As String
    Dim intTxtFileIndex  As Integer
        ''""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
        ' 取得 SPCNO 為唯一值的記錄，以同樣SPCNO的卡別一次取出
        ''""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
        '信用卡種類代碼,取出扣帳特店代碼
        

        
        If Not GetRS(rsCD037SpcNO, "SELECT DISTINCT SPCNO FROM " & TableOwnerName & "CD037  WHERE SpcNO IS NOT NULL  ", cn, adUseClient) Then Exit Function
        If rsCD037SpcNO.EOF Or rsCD037SpcNO.BOF Then
            MsgBox "信用卡種類的扣帳特店代碼必需設值 "
            rsCD037SpcNO.Close
            Exit Function
        End If
        ' ''""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
        '信用卡種類代碼,找出扣帳特店名稱<>NULL
        If Not GetRS(rsCD037, "SELECT * FROM " & TableOwnerName & "CD037 " & _
                "WHERE " & _
                "Spcno IS NOT NULL AND MerchantName IS NOT NULL " & _
                "ORDER BY CodeNO", cn, adUseClient) Then Exit Function
        intTxtFileIndex = 1
        ' ''""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
        Screen.MousePointer = vbHourglass
        Dim lngSubCount As Long        '' 儲存每一次不同的扣帳特店代號的處理資料筆數
        Dim lngTotalAmount As Long   '' 儲存每一次不同特扣帳特店代的處理資料金額總合
        Dim blnNewRec As Boolean     '' 記錄這是否為一筆新的
        Do While Not rsCD037SpcNO.EOF
            rsCD037.Filter = ""
            rsCD037.Filter = "SPCNO = " & rsCD037SpcNO(0)
            lngSubCount = 0
            lngTotalAmount = 0
            blnNewRec = False
            Do While Not rsCD037.EOF
                If blnNewRec = False Then
                       Call DefineRs   ''建立記憶體 recordset 結構
                End If
                 
                If Not BeginTran(rsCD037("Description"), rsCD037("SPCNO"), rsCD037("MerchantName"), intTxtFileIndex, lngSubCount, lngTotalAmount, blnNewRec) Then
                End If
                rsCD037.MoveNext
                blnNewRec = True
                DoEvents
            Loop
           
           '' 將收集所得的資料填入文字檔
           
            If lngSubCount > 0 Then
                rsCD037.MoveFirst
                ''       If OpenFile(txtDataPath, intTxtFileIndex, True, cboSet.ListIndex, rsCD037("MerchantName"), rsCD037("SPCNO")) = False Then
                
                If OpenFile(txtDataPath, intTxtFileIndex, True, cboSet.ListIndex, rsCD037("MerchantName"), rsCD037("CardID")) = False Then
                    MsgBox "無法建立登錄特店代碼為 " & rsCD037("SPCNO") & " 的文字檔 !!", vbOKOnly + vbCritical, "訊息 "
                    Exit Function
                End If
                
                '' 以下取得 header 字串
                
                strHeader = rsCD037("MerchantName") & Space((9 - Len(rsCD037("MerchantName"))))
                strHeader = strHeader & Format(gdaSystemDate.Text, "yyyymmdd") & Format(lngSubCount, "000000")
                strHeader = strHeader & Format(lngTotalAmount, "0000000000") & cboSet.ListIndex + 1
                strHeader = strHeader & "autobilling" & Space(133) & "00"
                Call subWriteLine(intTxtFileIndex)   ''''' 將取得的資料寫入文字檔
            End If
            rsCD037SpcNO.MoveNext
            intTxtFileIndex = intTxtFileIndex + 1
        Loop
        
        
        rsCD037.Close
        Set rsCD037 = Nothing
        rsCD037SpcNO.Close
        Set rsCD037SpcNO = Nothing
        subCitiBankFlow = True
    Exit Function
ChkErr:
    Call ErrSub(Me.Name, "subCitiBankFlow")
End Function

Private Function subHSBCFlow() As Boolean
    On Error GoTo ChkErr
    Dim intTxtFileIndex As Integer
    Dim lngSubCount As Long
    Dim lngTotalAmount As Long
    Dim strFileName As String
        lngSubCount = 0
        lngTotalAmount = 0
        intTxtFileIndex = 1
        Call DefineRs2   ''建立記憶體 recordset 結構
        If Not BeginTran("", txtSpcNO, txtMerchantName, intTxtFileIndex, lngSubCount, lngTotalAmount, True) Then
            'Exit Function
        End If
           
        '' 將收集所得的資料填入文字檔
        
        If lngSubCount > 0 Then
            strFileName = txtDataPath & "\" & txtMerchantName & "." & Format(Date, "yymmdd") & "." & Format(intTxtFileIndex, "00") & ".in"
            objStorePara.uProcText = strFileName
            Set file(intTxtFileIndex) = fso.CreateTextFile(strFileName, True)
            
            '' 以下取得 header 字串
            '1   記錄識別碼 1     'H'
            '2   終端機號碼  8   輸入0 至 9 數字(畫面上端末機代號)
            '3   商店代號    15  輸入 0 至 9數字， 靠左不足補空白(畫面上商店代號)
            '4   幣別    1   T':TWD, 'U':USD(預設T)
            '5   總筆數  6   輸入 000001-999999 數字
            '6   授權不清算交易(A)筆數   6   輸入 000001-999999 數字(結帳種類1 . 授權作計算)
            '7   授權不清算交易(A)金額   12  小數2位，靠右不足前面補 0(結帳種類1 . 授權作計算)
            '8   授權後清算交易(S)筆數   6   輸入 000001-999999 數字(結帳種類3 . 授權並請款作計算)
            '9   授權後清算交易(S)金額   12  小數2位，靠右不足前面補 0(結帳種類3 . 授權並請款作計算)
            '10  補登清算交易(O)筆數 6   輸入 000001-999999 數字(結帳種類2 . 請款作計算)
            '11  補登清算交易(O)金額 12  小數2位，靠右不足前面補 0(結帳種類2 . 請款作計算)
            '12  退貨清算交易(R)筆數 6   輸入 000001-999999 數字(結帳種類4 . 退款作計算)
            '13  退貨清算交易(R)金額 12  小數2位，靠右不足前面補 0(結帳種類4 . 退款作計算)
            '14  保留欄位    151 Space
            Dim lngSetCount(4) As Long, lngSetTotalAmt(4) As Long
            strHeader = "H" & GetString(txtMerchantName, 8) & GetString(txtSpcNO, 15)
            strHeader = strHeader & "T" & Format(lngSubCount, "000000")
            
            lngSetCount(cboSet.ListIndex + 1) = lngSubCount
            lngSetTotalAmt(cboSet.ListIndex + 1) = lngTotalAmount
            
            strHeader = strHeader & GetString(lngSetCount(1), 6, GIRIGHT, True) & GetString(lngSetTotalAmt(1), 10, GIRIGHT, True) & "00"
            strHeader = strHeader & GetString(lngSetCount(3), 6, GIRIGHT, True) & GetString(lngSetTotalAmt(3), 10, GIRIGHT, True) & "00"
            strHeader = strHeader & GetString(lngSetCount(2), 6, GIRIGHT, True) & GetString(lngSetTotalAmt(2), 10, GIRIGHT, True) & "00"
            strHeader = strHeader & GetString(lngSetCount(4), 6, GIRIGHT, True) & GetString(lngSetTotalAmt(4), 10, GIRIGHT, True) & "00"
            strHeader = strHeader & Space(151)
            Call subWriteLine(intTxtFileIndex)   ''''' 將取得的資料寫入文字檔
        End If
        subHSBCFlow = True
    Exit Function
ChkErr:
    Call ErrSub(Me.Name, "subHSBCFlow")
End Function

Private Sub Form_Load()
    On Error Resume Next
        blnSameFile = False
        lngErrCount = 0
        lngCount = 0
        cboBillType.ListIndex = 0
        Me.Caption = objStorePara.BankName & ""
        Call InitData
        RcdToScr
      ''取得是否使用多媒體
     
End Sub
Private Sub RcdToScr()
    On Error GoTo ChkErr
    Dim LogFile As TextStream
      txtPayDate.Text = Format(Date, "dd")
        If Dir(strPath & "\" & strPrgName & "Out.log") = "" Then Exit Sub
            Set LogFile = fso.OpenTextFile(strPath & "\" & strPrgName & "Out.log", ForReading)
            With LogFile
            
                'If Not .AtEndOfStream Then txtPayDate.Text = .ReadLine & ""   '請款日期
                If Not .AtEndOfStream Then txtDiscountRate.Text = .ReadLine & ""   '折扣率
                If Not .AtEndOfStream Then txtStatement.Text = .ReadLine & ""  '客戶帳單內容
                '              If Not .AtEndOfStream Then txtInVoice.Text = .ReadLine & ""   '統一編號
                
                If Not .AtEndOfStream Then txtDataPath = .ReadLine & ""   '資料檔
                If Not .AtEndOfStream Then txtErrPath = .ReadLine & ""   '問題檔
                If Not .AtEndOfStream Then txtSpcNO = .ReadLine & ""   '端末機名稱
                If Not .AtEndOfStream Then txtMerchantName = .ReadLine & ""   '商店名稱
                  
            End With
            LogFile.Close
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "RcdToScr")
End Sub

Private Sub InitData()
  On Error GoTo ChkErr
'  Dim o As Object
'  Set o = CreateObject("CreditCardOut.clsInterface")
    With objStorePara
      Set cn = .Connection
      strBankId = .BankId
      'strBankName = .BankName
      'strCorpId = .CorpID
      gdaSystemDate.SetValue (Date)
      strChoose = .ChooseStr
      strPath = .ErrPath
   ''   txtSpcNO.Text = .uSpcNo
    End With
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
'        With cnd1
'                .FileName = txtDataPath.Text
'                .Filter = "文字檔|*.txt"
'                .InitDir = strPath
'                .Action = 1
'                txtDataPath = .FileName
'        End With
    txtDataPath.Text = FolderDialog("資料檔路徑 ")
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
        
        ''"""""""""""""""""""""""""""""""""
        .Fields.Append ("SequenceNo"), adBSTR, 6, adFldIsNullable
        .Fields.Append ("MerchantNo"), adBSTR, 9, adFldIsNullable
        .Fields.Append ("DiscountRate"), adBSTR, 5, adFldIsNullable
        .Fields.Append ("CardNo"), adBSTR, 16, adFldIsNullable
        .Fields.Append ("ExpiryDate"), adBSTR, 6, adFldIsNullable
        
        .Fields.Append ("SequenceNo2"), adBSTR, 15, adFldIsNullable
        .Fields.Append ("Filler1"), adBSTR, 8, adFldIsNullable
        .Fields.Append ("Amount"), adBSTR, 7, adFldIsNullable
        .Fields.Append ("StatementDescription"), adBSTR, 38, adFldIsNullable
        .Fields.Append ("Filler2"), adBSTR, 2, adFldIsNullable
        
        .Fields.Append ("MerchantDescription"), adBSTR, 20, adFldIsNullable
        .Fields.Append ("StatusCode"), adBSTR, 2, adFldIsNullable
        .Fields.Append ("ApprovalCode"), adBSTR, 6, adFldIsNullable
        .Fields.Append ("AuthorizationDate"), adBSTR, 8, adFldIsNullable
        .Fields.Append ("Filler3"), adBSTR, 16, adFldIsNullable
        
        .Fields.Append ("Filler4"), adBSTR, 6, adFldIsNullable
        .Fields.Append ("PaymentFlag"), adBSTR, 1, adFldIsNullable
        .Fields.Append ("ResultDescription"), adBSTR, 3, adFldIsNullable
        .Fields.Append ("BillingDate"), adBSTR, 2, adFldIsNullable
        .Fields.Append ("Filler5"), adBSTR, 2, adFldIsNullable
        
        .Fields.Append ("HeaderIndicator"), adBSTR, 2, adFldIsNullable
        .Open
    End With
    
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "DefineRs")
End Sub

Private Sub DefineRs2()
    On Error GoTo ChkErr
    With rsDefTmp
        If .State = adStateOpen Then .Close
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        ''"""""""""""""""""""""""""""""""""
        .Fields.Append ("F01"), adVarChar, 1, adFldIsNullable
        .Fields.Append ("F02"), adVarChar, 6, adFldIsNullable
        .Fields.Append ("F03"), adVarChar, 8, adFldIsNullable
        .Fields.Append ("F04"), adVarChar, 15, adFldIsNullable
        .Fields.Append ("F05"), adVarChar, 6, adFldIsNullable
        
        .Fields.Append ("F06"), adVarChar, 1, adFldIsNullable
        .Fields.Append ("F07"), adVarChar, 19, adFldIsNullable
        .Fields.Append ("F08"), adVarChar, 4, adFldIsNullable
        .Fields.Append ("F09"), adVarChar, 3, adFldIsNullable
        .Fields.Append ("F10"), adVarChar, 12, adFldIsNullable
        
        .Fields.Append ("F11"), adVarChar, 6, adFldIsNullable
        .Fields.Append ("F12"), adVarChar, 2, adFldIsNullable
        .Fields.Append ("F13"), adVarChar, 20, adFldIsNullable
        .Fields.Append ("F14"), adVarChar, 20, adFldIsNullable
        .Fields.Append ("F15"), adVarChar, 40, adFldIsNullable
        
        .Fields.Append ("F16"), adVarChar, 1, adFldIsNullable
        .Fields.Append ("F17"), adVarChar, 10, adFldIsNullable
        .Fields.Append ("F18"), adVarChar, 3, adFldIsNullable
        .Fields.Append ("F19"), adVarChar, 9, adFldIsNullable
        .Fields.Append ("F20"), adVarChar, 9, adFldIsNullable
        
        .Fields.Append ("F21"), adVarChar, 7, adFldIsNullable
        .Fields.Append ("F22"), adVarChar, 16, adFldIsNullable
        .Fields.Append ("F23"), adVarChar, 36, adFldIsNullable
        .Open
        Dim intLoop As Integer
        Dim intMin As Long, intMax As Long
        intMin = 1
        For intLoop = 0 To .Fields.Count - 1
            intMax = intMin + .Fields(intLoop).DefinedSize - 1
            Debug.Print GetString((intLoop + 1), 2) & " " & GetString(.Fields(intLoop).DefinedSize, 3, GIRIGHT) & " " & GetString(intMin, 3, GIRIGHT) & "-" & GetString(intMax, 3, GIRIGHT)
            intMin = intMax + 1
        Next
    End With
        
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "DefineRs2")
End Sub

Private Function BeginTran(ByVal strCardName As String, ByVal strSpcNo As String, _
                                          ByVal strMerchantName As String, ByVal intIndexTxt As Integer, _
                                          ByRef pSubCount As Long, ByRef pTotalAmount As Long, ByVal pNewRec As Boolean) As Boolean

  On Error GoTo ChkErr
  
  Dim strsql As String
  Dim strOldBillNo As String
  Dim lngIndex As Long
  Dim lngSubCount As Long    '' 記錄這個卡別是不是有相同正確的資料被輸出，有再建立文字檔
  
  Dim set34 As String   '' 如果是授權與請款 則取正項的 SHOULDAMT
  Dim strSequenceNumber As String   '' 記錄  多媒體單號
  Dim strSQLGroup   ''92/08/14   調整規格 群組的語法根據是否啟重 intPara24 作改變 所以定義這個變數
  Dim strMediaBillNONull  As String
  Dim lngCustIDErr As Long  '' 取得某筆資料在SO002A的客編
  Dim rsCustIDErr As New ADODB.Recordset
  Dim rsSpcNo As New ADODB.Recordset
  Dim strSelectSPECNO As String
  ''2005/01/10  加入判斷式  取信用卡有效年月最大值
  Dim strAccountID As String
  Dim lngCustID As Long
  Dim CountSameCustIDBillNO As Integer
  
  Dim theSingalAmount As Long  '' 2005/02/02 用來儲存單筆SO033資料，
  
  Dim strCardNameSQL As String
  Dim lngerror As Long:    lngerror = 0
  
  Dim strWhereUCCode As String
  Dim strUpdUCCode As String
  Dim strStopYmWhere As String
  strAccountID = ""
  lngCustID = 0
  CountSameCustIDBillNO = 0
  BeginTran = False
    
'' 2005/02/02 不管正負項了  一律都出來
   set34 = "  "
   
   ''' 2003/07/23 新增客戶編號(CustID)以及媒體單號(MediaBillNo)名稱
   ''' 2003/08/14 改成SO002A 為資料來源，並且將GROUP BY 的字句割開，分成使用媒體單號與不使用的
   '''使用以MediaBillNO 為主，沒有以 Billno為主 ，全部的客戶編號資訊 則以  重新取得的SO002A 為準，CustStatusCode  在 SO002A 沒有 因此要將 SO002 串進來
'' *****************************************************************************************************
    ''  2004/09/24 加入SO106.StopFlag = 0 的條件值 防止出現兩筆
    ''  2004/09/24 還是將整段mark掉好了，直接取SO106就好了
    ''  以下的兩段SQL 將其中的 A.ShouldAmt > 0  的條件拿掉
    
    strStopYmWhere = "MAX(DECODE(LENGTH(STOPYM),5,SUBSTR(STOPYM,2,4) || '0' || SUBSTR(STOPYM,1,1), " & _
                        "SUBSTR(STOPYM,3,4) || SUBSTR(STOPYM,1,2)))"
    
    If strCardName <> "" Then strCardNameSQL = " And C.CARDNAME  ='" & strCardName & "'"
    '#5055 雖然沒提到,但順便一起改,櫃台已收則不再出帳 By Kin 2009/04/30
    '#5218 櫃台已收要用Nvl的方式 By Kin 2009/08/05
    '#5267 改用SO106去串資料 BY Kin 2010/04/21
    If intPara24 = 0 Then

        '#5564 增加參考號7,8與CD013.PAYOK代表已收 By Kin 2010/05/17
        '#5683 增加選取RealStopDate 與 PayKind By Kin 2010/08/06
        strsql = "SELECT " & GetUseIndexStr("A", "ShouldAmt") & " " & _
                " A.BillNO, sum(A.ShouldAmt)  ShouldAmt,  " & _
                " A.AccountNo,A.CUSTID, " & _
                strMinPayKind & " , " & strMinRealStopDate & _
                " From " & _
                TableOwnerName & "SO033 A ," & _
                TableOwnerName & "SO001 B ," & _
                TableOwnerName & "SO002 D,  " & _
                TableOwnerName & "CD013 " & _
                "Where " & strBankId & " AND " & _
                " A.CustID  = D.CUSTID AND " & _
                " A.CustId=B.CustId And A.CancelFlag = 0 And " & _
                " A.UCCode Is Not Null  AND  " & _
                " A.UCCode=CD013.CodeNo And Nvl(CD013.REFNO,0) NOT IN(3,7) AND  " & _
                " NVL(CD013.PAYOK,0) = 0 AND " & _
                " A.SERVICETYPE = D.SERVICETYPE " & _
                IIf(Len(strChoose) = 0, "", " And ") & strChoose & set34 & _
                " Group by A.BillNO,A.AccountNo,A.CUSTID ORDER BY  A.BillNO DESC  "

        strsql = "SELECT A.BILLNO,A.SHOULDAMT,A.ACCOUNTNO," & _
            " Min(A.PayKind) PayKind,Min(A.RealStopDate) RealStopDate, " & _
            strStopYmWhere & " CardExpDate " & _
            " FROM (" & strsql & ") A," & TableOwnerName & "SO106 C " & _
            " WHERE A.ACCOUNTNO=C.ACCOUNTID AND C.STOPFLAG <> 1 " & _
            " AND SnactionDate IS NOT NULL AND A.CUSTID=C.CUSTID " & strCardNameSQL & _
            " GROUP BY A.BILLNO,A.SHOULDAMT,A.ACCOUNTNO " & _
            " ORDER BY A.BILLNO DESC "
    Else
     'So033:正式應收資料,So001:客戶資料檔,So002A:客戶帳戶資料,So002:客戶服務基本資料,So106:信用卡及轉帳申請記錄
     '找出客戶正式應收總金額
        '#5267 改用SO106去串資料 By Kin 2010/04/21
        '#5564 增加參考號7,8與CD013.PAYOK代表已收 By Kin 2010/05/17
        strsql = "SELECT " & GetUseIndexStr("A", "ShouldAmt") & " " & _
                "sum(A.ShouldAmt)  ShouldAmt,  " & _
                "A.AccountNo,A.MediaBillNo,A.CustId, " & _
                strMinPayKind & " , " & strMinRealStopDate & _
                " From " & _
                TableOwnerName & "SO033 A ," & _
                TableOwnerName & "SO001 B ," & _
                TableOwnerName & "SO002 D,  " & _
                TableOwnerName & "CD013 " & _
                " Where " & strBankId & " AND " & _
                " A.CustID  = D.CUSTID AND " & _
                " A.CustId=B.CustId And A.CancelFlag = 0 And " & _
                " A.UCCode Is Not Null AND  " & _
                " A.UCCode=CD013.CodeNo And Nvl(CD013.REFNO,0) NOT IN(3,7) And " & _
                " NVL(CD013.PAYOK,0) = 0 AND " & _
                " A.SERVICETYPE = D.SERVICETYPE  " & _
                IIf(Len(strChoose) = 0, "", " And ") & strChoose & set34 & _
                " Group by A.AccountNO,A.MediaBillNo,A.CUSTID "
                
        strsql = "SELECT A.SHOULDAMT,A.ACCOUNTNO," & _
                    " A.MEDIABILLNO," & strStopYmWhere & " CardExpDate, " & _
                    " Min(PayKind) PayKind,Min(RealStopDate) RealStopDate " & _
            " FROM (" & strsql & ") A," & TableOwnerName & "SO106 C " & _
            " WHERE A.ACCOUNTNO=C.ACCOUNTID AND C.STOPFLAG <> 1 " & _
            " AND SnactionDate IS NOT NULL AND A.CUSTID=C.CUSTID " & strCardNameSQL & _
            " GROUP BY A.SHOULDAMT,A.ACCOUNTNO,A.MEDIABILLNO " & _
            " ORDER BY A.MEDIABILLNO DESC "
            
    End If
    
    '**************************************************************************************************************
    '#3417 更新UCCode與UCName基本的 Where條件 By Kin 2007/12/05
    strWhereUCCode = strBankId & strCardNameSQL & " AND " & _
                    "A.CustID  = C.CUSTID AND " & _
                    "A.CustID  = D.CUSTID AND " & _
                    "A.CustId=B.CustId And A.CancelFlag = 0 And " & _
                    "A.UCCode Is Not Null AND  " & _
                    "A.SERVICETYPE = D.SERVICETYPE   AND A.AccountNo = C.AccountID  AND C.StopFlag = 0  " & _
                    IIf(Len(strChoose) = 0, "", " And ") & strChoose
    '***************************************************************************************************************

 ''  2003/08/14 改成 再將 mediabillno   資料取出 這裏這一次多取一個 BillNO 然後填入當mediabillno為null 值的時候，填入媒體單號
    '#5055 沒提到,但順便一起改,要過濾掉櫃台已收 By Kin 2009/04/30
    '#5218 櫃台已收要用Nvl方式 By Kin 2009/08/05
    If intPara24 = 1 Then
    
                        ''這一段直接取前述的主 SQL (strsql)  拿掉 group  ，加上 Billno
        '#5564 增加參考號7,8與CD013.PAYOK代表已收 By Kin 2010/05/17
        strMediaBillNONull = "SELECT " & GetUseIndexStr("A", "ShouldAmt") & " " & _
                             "C.AccountID,A.MediaBillNo,A.BillNO,MAX(A.MEDIABILLNO)   " & _
                             "From " & _
                              TableOwnerName & "SO033 A ," & _
                              TableOwnerName & "SO001 B ," & _
                              TableOwnerName & "SO106 C,  " & _
                              TableOwnerName & "SO002 D,  " & _
                              TableOwnerName & "CD013 " & _
                             "Where " & strBankId & strCardNameSQL & " AND " & _
                             "A.CustID  = C.CUSTID AND " & _
                             "A.CustID  = D.CUSTID AND " & _
                             "A.CustId=B.CustId And A.CancelFlag = 0 And " & _
                             "A.UCCode Is Not Null  AND " & _
                             "A.UCCode=CD013.CodeNo And Nvl(CD013.REFNO,0) NOT IN(3,7) And " & _
                             "NVL(CD013.PAYOK,0) = 0 AND " & _
                             "A.SERVICETYPE=D.SERVICETYPE  AND A.AccountNo = C.AccountID  " & _
                             IIf(Len(strChoose) = 0, "", " And ") & strChoose & set34 & "  AND A.MEDIABILLNO IS NULL" & _
                             " GROUP BY C.ACCOUNTID,A.MEDIABILLNO,A.BILLNO "
                                            
                                            
        strChooseMultiAcc = " A.CancelFlag = 0 AND " & _
                            " A.UCCode Is Not Null  " & _
                            IIf(Len(strChooseMultiAcc) = 0, "", " And ") & strChooseMultiAcc
                
        If GetRS(rsMedioBillnoNull, strMediaBillNONull, cn) Then
            Do While Not rsMedioBillnoNull.EOF
                If IsNull(rsMedioBillnoNull("MediaBillNo")) Then
                     strSequenceNumber = GetInvoiceNo2("SO033_MediaBillNo", cn)
                     cn.Execute ( _
                                    "UPDATE " & TableOwnerName & "SO033 A SET " & _
                                    "A.MediaBillNo  ='" & strSequenceNumber & "'  WHERE   " & _
                                    "A.BillNo = '" & rsMedioBillnoNull("BillNO") & "'   AND  " & strChooseMultiAcc & " AND " & _
                                    "A.AccountNO ='" & rsMedioBillnoNull("AccountID") & "' ")
                    rsMedioBillnoNull.MoveNext
                End If
            Loop
                     
        End If
        rsMedioBillnoNull.Close
        Set rsMedioBillnoNull = Nothing
    End If
    
    If intPara24 = 0 Then
        strChooseMultiAcc = " A.CancelFlag = 0 AND " & _
                            "A.UCCode Is Not Null   " & _
                             IIf(Len(strChooseMultiAcc) = 0, "", " And ") & strChooseMultiAcc
    End If
    
    Set rsTmp = cn.Execute(strsql)
    
     ''  2003/07/23    這一段取得只有 SO033 的篩選資料  ........
    
    If rsTmp.BOF Or rsTmp.EOF Then
        blnNodata = True
        objStorePara.uNoData = True
    Else
        blnNodata = False
        Dim strSN As String  '' 指定序列號
        Dim lngTotal As Long '' 累計標頭檔 Total Amount 欄位總出單金額
        Dim lngDoubleMediaNO As Long '' 這個值記錄重覆出單的單子件數

        lngDoubleMediaNO = 0
        If pNewRec = False Then
             lngTotal = 0
             lngSubCount = 0
        Else
            lngTotal = pTotalAmount
            lngSubCount = pSubCount
        End If
                                  
        Do While Not rsTmp.EOF
            Dim strCardExpDate As String
            Dim iCED As Integer
            Dim dteCED As Date
            Dim strErrMessage As String
            Dim strBillNOField As String
            
            If intPara24 = 0 Then
                 strBillNOField = "BillNO"
            Else
                 strBillNOField = "MediaBillNO"
            End If
                 
           ''  20050110 加入斷式 取同帳號信用卡裏面到期年月最大的 由於SQL已經DESC排序
           ''  因此如果單號是相同的, 表示取出的是相同的兩筆，就不要繼續往下了，
           '' 直到接下來的 一筆是不一樣的再出單
            '' ************************************************************************
            Dim strSopYMMaxString As String: lngerror = 1
            If strAccountID = rsTmp(strBillNOField) Then
                lngerror = 2
                GoTo NextLoop
            Else
                lngerror = 3
                strAccountID = rsTmp(strBillNOField)
                lngerror = 4
                '' 20050205 調整  取第一筆最大的 因為原來的不能用是文字的排出來會出問題
                
                ''   strSopYMMaxString = GetStopYMDateString(rsTmp(strBillNOField), strCardName, set34)
                '#5267 已經改用Max(SO106.StopYm),所以不需要此段了 By Kin 2010/04/21
                'strSopYMMaxString = GetStopYMDateString(strAccountID, strCardName, set34)
                strSopYMMaxString = rsTmp("CardExpDate") & ""
                lngerror = 45
            End If
            '' ************************************************************************
      '     ''2005/01/26將整理好的資料  依照單據排序之後  將同單號正負項合併
           '''Call LogSameNOPMAmount(rsTmp(strBillNOField), strCardName, set34)
           ''2005/02/02 將整理好的資料  依照單據排序之後  將同單號正負項合併
           
           lngerror = 5
           '問題集2868 秀出Log時，客編如果有多帳戶會判斷錯誤，多增加以SO033的客編判斷 2006/11/15 for Jacy
            Call GetRS(rsCustIDErr, "SELECT  distinct SO001.CUSTNAME , SO106.CUSTID FROM " & TableOwnerName & "SO106," & TableOwnerName & "SO001  " & _
                                 "WHERE " & _
                                 "SO106.AccountID ='" & rsTmp("AccountNO") & "'  AND  SO106.CUSTID = SO001.CUSTID And SO106.CUSTID IN (" & _
                                 "Select DISTINCT CUSTID From " & TableOwnerName & "SO033 Where " & strBillNOField & "='" & _
                                  rsTmp(strBillNOField) & "' And AccountNo='" & rsTmp("AccountNO") & "')", cn)
            If rsCustIDErr.RecordCount <= 0 Then
                Call GetRS(rsCustIDErr, "Select SO001.CUSTNAME,SO033.CUSTID FROM " & TableOwnerName & "SO001," & TableOwnerName & "SO033 " & _
                                       "Where SO001.CUSTID=SO033.CUSTID And SO033.AccountNo='" & rsTmp("AccountNO") & "' And " & strBillNOField & "='" & _
                                       rsTmp(strBillNOField) & "'", cn)
                                       
                lngErrCount = lngErrCount + 1: lngerror = 6.5
                strErrMessage = "客戶帳戶資料檔 :單號  " & rsTmp(strBillNOField) & " 客戶姓名 " & rsCustIDErr("CustName") & " 客戶編號  " & rsCustIDErr("CustID")
                file(0).WriteLine strErrMessage: lngerror = 6.8
                GoTo NextLoop
            End If
            '#5683 如果收視截止日小於畫面條件不要出帳 By Kin 2010/08/06
            If Not IsPayKindOK(rsTmp, gdaSystemDate.GetValue & "", giAll, intPara24) Then
                lngErrCount = lngErrCount + 1
                strErrMessage = "收視截止日錯誤 :單號  " & rsTmp(strBillNOField) & " 客戶姓名 " & rsCustIDErr("CustName") & _
                        " 客戶編號  " & rsCustIDErr("CustID") & " 收視截止日 : " & rsTmp("RealStopDate") & _
                        " 繳付類別 : 現付制 "
                file(0).WriteLine strErrMessage: lngerror = 6.9
                GoTo NextLoop
            End If
            
   lngerror = 6
            If IsNull(rsTmp("AccountNO")) Then
                lngErrCount = lngErrCount + 1:    lngerror = 7
                strErrMessage = "信用卡卡號空白 : 單號　" & rsTmp(strBillNOField) & " 客戶姓名 " & rsCustIDErr("custname") & " 客戶編號  " & rsCustIDErr("CustID")
                file(0).WriteLine strErrMessage:     lngerror = 8
                GoTo NextLoop
            End If
           
           
           ''20050205 這一行是不是也要改掉 呢   ************************************
            '***********************************
           ' ************************************
         ''  strCardExpDate = GetNullString(rsTmp("CardExpDate"), giStringV, giAccessDb)
                      '' 20050205加入這一行改成抓GetStopYMDateString取得的最大值
                      
            strCardExpDate = strSopYMMaxString
            If strCardExpDate = "NULL" Then
                lngErrCount = lngErrCount + 1
                strErrMessage = "信用卡資料不正確 : 單號　" & rsTmp(strBillNOField) & " 客戶姓名 " & rsCustIDErr("custname") & " 客戶編號  " & rsCustIDErr("CustID")
                file(0).WriteLine strErrMessage
                GoTo NextLoop
            End If
            lngerror = 9
            '#5267 取出時已經格式好了,不需要再轉了 By Kin 2010/04/21
            'strCardExpDate = Right(strSopYMMaxString, 4) & "/" & Replace(strSopYMMaxString, Right(strSopYMMaxString, 4), "")
            strCardExpDate = Left(strSopYMMaxString, 4) & "/" & Right(strSopYMMaxString, 2)
            lngerror = 10
           
            dteCED = CDate(strCardExpDate)
            If dteCED < CDate(Format(gdaSystemDate.Text, "YYYY/MM")) Then
                lngErrCount = lngErrCount + 1
                strErrMessage = "信用卡過期 : 單號　" & rsTmp.Fields(strBillNOField) & " 客戶姓名　" & rsCustIDErr("custname") & " 信用卡號 " & rsTmp.Fields("AccountNO") & " 到期日　" & Format(rsTmp.Fields("CardExpDate"), "000000") & " 客戶編號  " & rsCustIDErr("CustID")
                file(0).WriteLine strErrMessage
                If chkDuteDate.Value = 0 Then GoTo NextLoop
            End If
          
'           strCardExpDate = rsTmp("CardExpDate")
'           iCED = IIf(Len(strCardExpDate) = 5, 1, 2)
'           strCardExpDate = Right(strCardExpDate, 4) & "/" & Mid(strCardExpDate, 1, iCED)
          
          '' 2005/02/02 如果加總金額小於零 作LOG不出單
          lngerror = 11
          Dim strtest As String
          strtest = rsTmp(strBillNOField)
          lngerror = 111
          theSingalAmount = GetSumAmout(rsTmp(strBillNOField), strCardName, set34)
          lngerror = 112
          If theSingalAmount < 0 Or theSingalAmount = 0 Then
                    strErrMessage = " 負項或是金額等於零：單號　" & rsTmp(strBillNOField)
                               lngerror = 12

                    lngErrCount = lngErrCount + 1
                    file(0).WriteLine strErrMessage
                               lngerror = 12

                    GoTo NextLoop
          End If
            lngerror = 13
        
        '問題集 2852 檢查特扣店是否正確如果不正確，成功筆數會出錯(此段本來在下面,現在提到這邊來處理) For Jacy 2006/11/08
        If cboBillType.ListIndex = 2 Then
            strSpcNo = "Select SPCNO From CD037 WHERE CARDID IN (Select DISTINCT CARDCODE FROM " _
                       & TableOwnerName & " So106 A," & _
                        TableOwnerName & " So002A B Where A.ACCOUNTID='" & rsTmp("AccountNO") & "'" & _
                        " AND A.CUSTID=B.CUSTID AND A.ACCOUNTID=B.ACCOUNTNO AND A.CUSTID=" & rsCustIDErr("CustID") & " ) "
            If Not GetRS(rsSpcNo, strSpcNo, cn) Then Exit Do
            If rsSpcNo.RecordCount <= 0 Then
                strErrMessage = " 扣帳特店代碼不正確：單號　" & rsTmp(strBillNOField) & " 客戶姓名　" & rsCustIDErr("custname") & "  客戶編號  " & rsCustIDErr("CustID")
                lngErrCount = lngErrCount + 1
                file(0).WriteLine strErrMessage
                GoTo NextLoop
            Else
                strSpcNo = Trim(GetString(rsSpcNo(0), 15))
            End If
        End If
        lngCount = lngCount + 1
        lngSubCount = lngSubCount + 1
        If cboBillType.ListIndex = 0 Then
            With rsDefTmp
                 .AddNew
                 .Fields("SequenceNo") = Format(lngSubCount, "000000")
                 '' SequenceNo(1-6)
                 '#3201 因為CD037.SPCNO長度變為15會影嚮花旗原格式，所以原格式取SPCNO時，多增加Mid函數，強迫取前面九碼 By Kin 2007/08/03
                 .Fields("MerchantNo") = Mid(Format(strSpcNo, "000000000"), 1, 9)
                 ''MerchantNo(7-15)
                 .Fields("DiscountRate") = Format(txtDiscountRate.Text, "00000")
                 ''DiscountRate(16-20)
                 .Fields("CardNo") = IIf(Len(rsTmp("AccountNO")) > 0, Format(rsTmp("AccountNO"), "0000000000000000"), "0000000000000000")
                 ''CardNo(21-36)
                 '' 20050205 MARK以上的那一行，改成以下的
                 strCardExpDate = Replace(strCardExpDate, "/", "")
                 If Len(strCardExpDate) = 5 Then
                     strCardExpDate = Left(strCardExpDate, 4) & "0" & Right(strCardExpDate, 1)
                 End If
                 .Fields("ExpiryDate") = Right(strCardExpDate, 2) & Left(strCardExpDate, 4)
                 ''ExpiryDate(37-42)
                 .Fields("SequenceNo2") = Format(lngSubCount, "000000000000000")
                 ''SequenceNo2(43-57)
                 .Fields("Filler1") = Space(8)
                 '' Filler(58-65)
                 ''2005/02/02 以下金額欄位值改成重取SO033的金額 避免程式取得重覆金額  !!
                 .Fields("Amount") = Format(theSingalAmount, "0000000")
                 lngTotal = lngTotal + theSingalAmount
                 '' Amount(66-72)
                  Dim strSD As String
                  Dim intlength As Integer
                  
                  '""轉換成為正確長度的全形字形'''''''''''''''''''''''''
                  strSD = StrConv(txtStatement.Text, vbFromUnicode)
                  intlength = LenB(strSD)
                  strSD = txtStatement.Text & Space(38)
                   ''""""""""""""""""""""""""""""""""""""""""""""""""""""
                 .Fields("StatementDescription") = StrConv(LeftB(StrConv(StrConv(strSD, vbWide), vbFromUnicode), 38), vbUnicode)
                 
                 .Fields("Filler2") = Space(2)
                 
                '' 2003/07/23 以下進行單號填值的操作，判斷是否以媒體單號的功能入單
                '' 如果是則以媒體單號入單，否則轉換成舊單號之後再入
                '' 2003/08/14 以下的內容，取得單號內容，如果啟用媒體的功能，就直接取MediaBillNO
                '問題集2861 如果是使用媒体單號，匯出時會媒体單號會秀不出來，因為Else下面那行，不小心被拿掉了
                 If intPara24 = 0 Then
                      Dim strBillNOOld As String
                      strBillNOOld = Trim(CStr(Val(Left(rsTmp("BillNO"), 6)) - 191100)) & _
                                     Mid(rsTmp("BillNO"), 7, 1) & Format(Right(rsTmp("BillNO"), 6), "000000")
                     .Fields("MerchantDescription") = strBillNOOld & Space(20 - Len(strBillNOOld))
                 Else
                     .Fields("MerchantDescription") = GetString(rsTmp("MediaBillNo"), 20, giLeft, False)
                 End If
                 ''MerchantDescriptione(113-132)
                 .Fields("StatusCode") = Space(2)
                 ''StatusCode(133-134)
                 .Fields("ApprovalCode") = Space(6)
                 ''ApprovalCode(135-140)
                 .Fields("AuthorizationDate") = Space(8)
                  ''AuthorizationDate(141-148)
                 .Fields("Filler3") = Space(16)
                  ''Filler(149-164)
                 .Fields("Filler4") = Space(6)
                  ''Filler(165-170)
                 .Fields("PaymentFlag") = Space(1)
                 .Fields("ResultDescription") = Space(3)
                 .Fields("BillingDate") = Format(txtPayDate.Text, "00")
                 .Fields("Filler5") = Space(2)
                 .Fields("HeaderIndicator") = "01"
                 .Update
            End With
         Else
            'HSB格式
            If cboBillType.ListIndex = 1 Then
                If Not BeginTran2(lngSubCount, lngTotal, strCardExpDate, theSingalAmount) Then Exit Do
            Else
                If Not BeginTran3(lngSubCount, lngTotal, strCardExpDate, theSingalAmount, strSpcNo) Then Exit Do
                rsSpcNo.Close
                Set rsSpcNo = Nothing
            End If
         End If
         
'         TableOwnerName & "SO033 A ," & _
'                TableOwnerName & "SO001 B ," & _
'                TableOwnerName & "SO002A C,  " & _
'                TableOwnerName & "SO002 D,  " & _
'                TableOwnerName & "SO106 E  " & _

        '**********************************************************************************************************************************
        '#3417 如果CD013.REFNO有包含到4,則更新UCCode與UCName By Kin 2007/12/06
        If blnUpdUcCode Then
           '#3892 更新速度變慢,調整語法,原本用Exists,現在改用= By Kin 2008/06/03
           '#4388 增加更新異動人員與異動時間 By Kin 2009/04/30
           '#5218 櫃台已收要用Nvl的方式 By Kin 2009/08/05
           '#5564 增加參考號7,8與CD013.PAYOK代表已收 By Kin 2010/05/17
            strUpdUCCode = "UPDATE " & TableOwnerName & "SO033 Set UCCode=" & strUCCode & ",UCName='" & strUCName & "'" & _
                           ",UPDEN='" & objStorePara.uUPDEN & "',UPDTime='" & objStorePara.uUPDTime & "' " & _
                        " Where " & strBillNOField & " In (Select A." & strBillNOField & " From " & _
                                        TableOwnerName & "SO001 B ," & _
                                        TableOwnerName & "SO106 C," & _
                                        TableOwnerName & "SO002 D," & _
                                        TableOwnerName & "SO033 A," & _
                                        TableOwnerName & "CD013 " & _
                                        " Where " & strWhereUCCode & _
                                        " And A." & strBillNOField & "='" & rsTmp(strBillNOField) & "'" & _
                                        " And C.AccountID='" & rsTmp("AccountNO") & "'" & _
                                        " And A.UCCode=CD013.CodeNo And Nvl(CD013.REFNO,0) NOT IN(3,7) " & _
                                        " AND NVL(CD013.PAYOK,0) =0 " & _
                                        IIf(rsTmp("CardExpDate") & "" <> "", _
                                        " And Decode(Length(C.StopYM),5,Substr(C.StopYM,2,4)||'0'||Substr(C.StopYM,1,1),Substr(C.StopYM,3,4)||Substr(C.StopYM,1,2))=" & rsTmp("CardExpDate"), "") & ")"
                                    
            cn.Execute strUpdUCCode
        End If
        '**********************************************************************************************************************************
NextLoop:
        rsTmp.MoveNext
        Loop
        pSubCount = lngSubCount
        pTotalAmount = lngTotal
        
        BeginTran = True
        On Error Resume Next
        rsTmp.Close
        Set rsTmp = Nothing
  End If
  
Exit Function
ChkErr:
    Call CloseFS(intIndexTxt)
    Call ErrSub(Me.Name, "BeginTran" & lngerror & "-" & strAccountID)
End Function

Private Function BeginTran2(ByRef lngSubCount As Long, ByRef lngTotal As Long, _
    ByVal strCardExpDate As String, theSingalAmount As Long) As Boolean
    On Error GoTo ChkErr
    
        With rsDefTmp
            .AddNew
            '1   記錄識別碼 1     'D'
            .Fields("F01") = "D"
            '2   交易序號    6   輸入 0 至 9數字，靠右不足前面補 0(流水號)
            .Fields("F02") = Format(lngSubCount, "000000")
            '3   終端機號碼  8   輸入0 至 9 數字字(畫面上端末機代號)
            .Fields("F03") = GetString(txtMerchantName, 8)
            '4   商店代號    15  輸入 0 至 9數字， 靠左不足補空白(畫面上商店代號)
            .Fields("F04") = GetString(txtSpcNO, 15)
            '5   交易日  6   YYMMDD(西元年) (畫面上系統年月+請款日)
            .Fields("F05") = Format(Date, "YYMM") & txtPayDate
            '6   授權清算方式    1   A:只授權不清算(1 . 授權), S:授權後清算 (3 . 授權並請款),O:補登清算(2 . 請款),  R:退貨清算(4 . 退款)(畫面結帳種類選項決定)
            Select Case cboSet.ListIndex
                Case 0          '1 .授權
                    .Fields("F06") = "A"
                Case 1          '2 .請款
                    .Fields("F06") = "O"
                Case 2          '3 .授權並請款
                    .Fields("F06") = "S"
                Case 3          '4 .退款
                    .Fields("F06") = "R"
            End Select
            '7   卡    號    19  靠左,不足右補空白
            .Fields("F07") = GetString(rsTmp("AccountNo"), 19)
            '8   有效期限    4   MMYY, MBF(Must be full)(西元年)
            strCardExpDate = Format(strCardExpDate, "MMYY")
            .Fields("F08") = Left(strCardExpDate, 2) & Right(strCardExpDate, 2)
            '9   CVC2    3   填入 CVC2 or Space(預設Spac)
            .Fields("F09") = Space(3)
            '10  金    額    12  靠右左補最後兩碼小數位數
            .Fields("F10") = GetString(theSingalAmount, 10, GIRIGHT, True) & "00"
            lngTotal = lngTotal + theSingalAmount
            '11  核准號碼    6   未授權填入'999999'or 已取得授權碼則直接填入授權碼(先授權後請款用)
            .Fields("F11") = "999999"
            '12  回 復 碼    2   未授權填入'99'or 已授權填入'00'(入帳程式是別為00時才可以入賬)
            .Fields("F12") = "99"
             '13  持卡人帳單列印資料  20  此欄位資料可印在持卡人對帳單上(畫面上帳單內容過長請截掉)
            .Fields("F13") = GetString(txtStatement, 20)
            '14  商店保留欄位    20  商店自行運用(Space)(媒體單號客服系統入帳依據)
            If intPara24 = 0 Then
                .Fields("F14") = GetString(rsTmp("BillNo"), 20)
            Else
                .Fields("F14") = GetString(rsTmp("MediaBillNo"), 20)
            End If
            '15  人工授權資料    40  Context Length  Note
            '            Name    8   靠左不足補空白
            '            ID  10  靠左不足補空白
            '            DOB 7   YYYMMDD(民國年),靠右左不足補空白
            '            Tel 15  靠左不足補空白
            '            此欄位在二次人工授權時使用 (目前系統無使用填空白)
            .Fields("F15") = Space(40)
            '16  分期註記    1   1:分期交易 , 0:非分期交易(目前系統無使用分期填0)
            .Fields("F16") = "0"
            '17  產品代碼    10  靠左不足補 '0'(目前系統無使用填空白)
            .Fields("F17") = Space(10)
            '18  期數    3   分期付款用, Right justified, Start with zeros(主機回傳) (目前系統無使用填空白)
            .Fields("F18") = Space(3)
            '19  每期金額    9   分期付款用, Right justified, Start with zeros(主機回傳) (目前系統無使用填空白)
            .Fields("F19") = Space(9)
            '20  第一期金額  9   分期付款用, Right justified, Start with zeros(主機回傳) (目前系統無使用填空白)
            .Fields("F20") = Space(9)
            '21  手續費率    7   分期付款用, Right justified, Start with zeros(主機回傳) (目前系統無使用填空白)
            .Fields("F21") = Space(7)
            '22  晶片碼(TC)  16  晶片交易請款用(目前系統無使用填空白)
            .Fields("F22") = Space(16)
            '23  保留欄位    36  Space(目前系統無使用填空白)
            .Fields("F23") = Space(36)
            '24  CR+LF   2   Chr(13)+Chr(10)
            '.Fields("F24") = ""
        End With
        BeginTran2 = True
    Exit Function
ChkErr:
    ErrSub Me.Name, "BeginTran2"
End Function

Private Function subWriteLine(ByVal intIndexTxt As Integer) As Boolean

  On Error GoTo ChkErr
  
    Dim varData As Variant
    Dim strData As String
    Dim lngLoop As Long
    Dim lngLoopi As Long
    Dim strContent As String  '' 每一筆寫入的資料

    subWriteLine = False
    
    With rsDefTmp
        
         If .BOF Or .EOF Then Exit Function
         .MoveFirst
         If blnSameFile = False Then
              file(intIndexTxt).WriteLine strHeader     ''於文字檔填入首筆
         End If
         For lngLoop = 0 To .RecordCount - 1
            strContent = ""
            For lngLoopi = 0 To rsDefTmp.Fields.Count - 1
                strContent = strContent & rsDefTmp(lngLoopi)
            Next
            file(intIndexTxt).WriteLine strContent
            rsDefTmp.MoveNext
            DoEvents
         Next
    End With
    subWriteLine = True
    On Error Resume Next
    rsDefTmp.Close
    Set rsDefTmp = Nothing
    file(intIndexTxt).Close
    Exit Function
ChkErr:
    subWriteLine = False
    Call ErrSub(Me.Name, "subWriteLine")
End Function

Private Sub ScrToRcd()
    On Error GoTo ChkErr
    Dim LogFile As TextStream
        
        Set LogFile = fso.CreateTextFile(strPath & "\" & strPrgName & "Out.log", True)
        With LogFile
                '.WriteLine (txtPayDate.Text)      '請款日期
                .WriteLine (txtDiscountRate.Text) '折扣率
                .WriteLine (txtStatement.Text)    '客戶帳單內容
                '.WriteLine (txtInVoice.Text)     '統一編號
                .WriteLine (txtDataPath)         '資料檔路徑
                .WriteLine (txtErrPath.Text)     ' 問題參考檔路徑
                .WriteLine (txtSpcNO.Text)  '端末機名稱
                .WriteLine (txtMerchantName.Text) '端末機名稱
        End With
        LogFile.Close
        Set LogFile = Nothing
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "ScrToRcd")
End Sub

Private Function IsDataOk() As Boolean
    Dim strErrMsg As String
        '如果不是選擇花旗批次，則結帳種類1.授權與2.請款無法使用
        If cboBillType.ListIndex <> 2 Then
            If cboSet.ListIndex = 0 Or cboSet.ListIndex = 1 Then
                MsgBox "結帳種類 (1.授權 2.請款) 目前無法使用 !!", vbOKOnly + vbInformation, "訊息　"
                cboSet.SetFocus
                Exit Function
            End If
        End If
        If Len(cboSet.Text) = 0 Then
            MsgBox "請選擇一種結帳種類 !!", vbOKOnly + vbInformation, "訊息 "
            cboSet.SetFocus
            Exit Function
        End If
        
        If Not IsNumeric(txtPayDate.Text) Then
            MsgBox " 請款日期格式不正確　!! ", vbOKOnly + vbInformation, "訊息　"
            txtPayDate.SetFocus
            Exit Function
        End If
        
        If Not IsDate(Format(Date, "yyyy/mm") & "/" & txtPayDate.Text) Then
            MsgBox " 請款日期格式不正確　!! ", vbOKOnly + vbInformation, "訊息　"
            txtPayDate.SetFocus
            Exit Function
        End If
        
        If Len(gdaSystemDate.Text) = 0 Then
            MsgBox " 系統日期必需有值　!! ", vbOKOnly + vbInformation, "訊息　"
            gdaSystemDate.SetFocus
            Exit Function
        End If
         
        If Not IsNumeric(txtDiscountRate.Text) Then
            MsgBox " 折扣率必需為數字　!! ", vbOKOnly + vbInformation, "訊息　"
            txtDiscountRate.SetFocus
            Exit Function
        End If
  
        If cboBillType.ListIndex <> 0 Then
            If Len(txtSpcNO.Text & "") = 0 Then strErrMsg = "商店代號": txtSpcNO.SetFocus: GoTo Warning
            If Len(txtMerchantName.Text & "") = 0 Then strErrMsg = "端末機代號": txtMerchantName.SetFocus: GoTo Warning
        End If
        '如果選擇花旗批次，則結帳種類只能選授權
        If cboBillType.ListIndex = 2 Then
            If cboSet.ListIndex <> 0 Then
                MsgBox "結帳種類只能選擇 1.授權", vbOKOnly + vbInformation, "訊息"
                cboSet.ListIndex = 0
                cboSet.SetFocus
                Exit Function
            End If
            If Trim(txtBatch) = "" Then
                strErrMsg = "批號"
                txtBatch.SetFocus
                GoTo Warning
            End If
        End If
        If Len(cboSet.Text) = 0 Then strErrMsg = "結帳種類": cboSet.SetFocus: GoTo Warning
        If Len(txtPayDate.Text) = 0 Then strErrMsg = "請款日期": txtPayDate.SetFocus: GoTo Warning
        If Len(txtDiscountRate.Text) = 0 Then strErrMsg = "折扣率": txtDiscountRate.SetFocus: GoTo Warning
        If Len(txtDataPath.Text) = 0 Then strErrMsg = "資料檔路徑": txtDataPath.SetFocus: GoTo Warning
        If Len(txtErrPath.Text) = 0 Then strErrMsg = "問題參考檔路徑": txtErrPath.SetFocus: GoTo Warning
        IsDataOk = True
    
    Exit Function
Warning:
        MsgBox "欄位 " & strErrMsg & " 必需有值", vbOKOnly + vbInformation, "訊息"
    Exit Function
ChkErr:
    Call ErrSub(Me.Name, "IsDataOk")
End Function

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Set rsTmp = Nothing
    If rsCD013.State = adStateOpen Then rsCD013.Close
    Set rsCD013 = Nothing
    file(0).Close
    Set fso = Nothing
End Sub

Private Sub gdaSystemDate_LostFocus()
   Dim intDD As Integer
   If Len(gdaSystemDate.Text) <> 0 Then
   intDD = CInt(Format(gdaSystemDate.Text, "DD"))
   txtPayDate.Text = Format(intDD, "00")
   End If
End Sub

''2005/01/26將整理好的資料  依照單據排序之後  將同單號正負項合併
Private Sub LogSameNOPMAmount(ByVal strMNO As String, ByVal CardName As String, ByVal set34sql As String)

    Dim checkSameNOSQL As String
    Dim snoType As String
    Dim rscheckSameNOSQL As New ADODB.Recordset
    
    If intPara24 = 0 Then
         snoType = "BillNO"
    Else
         snoType = "MediaBillNo"
    End If
    
checkSameNOSQL = "SELECT Count(A." & snoType & ")   " & _
                    "From " & _
                    TableOwnerName & "SO033 A ," & _
                    TableOwnerName & "SO001 B ," & _
                    TableOwnerName & "SO002A C,  " & _
                    TableOwnerName & "SO002 D,  " & _
                    TableOwnerName & "SO106 E  " & _
                    "Where " & strBankId & " AND  " & _
                    "C.CARDNAME  ='" & CardName & "' AND " & _
                    "A.CustID  = C.CUSTID AND " & _
                    "A.CustID  = D.CUSTID AND " & _
                    "A.CustID  = E.CUSTID AND C.AccountNO  = E.AccountID  AND  " & _
                    "A.CustId=B.CustId And A.CancelFlag = 0 And " & _
                    "A.UCCode Is Not Null AND " & _
                    "A.SERVICETYPE = D.SERVICETYPE   AND A.AccountNo = C.AccountNo  AND E.StopFlag = 0  " & _
                    IIf(Len(strChoose) = 0, "", " And ") & strChoose & set34sql & "  AND A." & snoType & " = '" & strMNO & "'  AND A.ShouldAmt< 0   "
    
If GetRS(rscheckSameNOSQL, checkSameNOSQL, cn, adUseClient, adOpenKeyset) = True Then
      If rscheckSameNOSQL(0) > 0 Then
                '' 進行  log 記錄
                file(0).WriteLine " 負項金額 : 單號　" & strMNO
      End If
End If

rscheckSameNOSQL.Close
Set rscheckSameNOSQL = Nothing

Exit Sub

ChkErr:
    Call ErrSub(Me.Name, "LogSameNOPMAmount")
End Sub

Private Function GetSumAmout(ByVal strMNO As String, ByVal CardName As String, ByVal set34sql As String)
    Dim SingalSumAmount As Long: SingalSumAmount = 0
    Dim checkSameNOSQL As String
    Dim snoType As String
    Dim rscheckSameNOSQL As New ADODB.Recordset
    
        If intPara24 = 0 Then
             snoType = "BillNO"
        Else
             snoType = "MediaBillNo"
        End If
         '' 將 C.StopFlag = 0  這個拿掉，避免搜不到資料
        '#4040 多串SO001 不然選擇客戶類別會出錯 By Kin 2008/08/13
        checkSameNOSQL = "SELECT SUM(A.ShouldAmt)  ShouldAmt  " & _
                            "From " & _
                            TableOwnerName & "SO033 A," & _
                            TableOwnerName & "SO002 D," & _
                            TableOwnerName & "SO001 B " & _
                            "Where " & _
                            "A.CancelFlag = 0 And " & _
                            "A.CustID  = D.CUSTID AND " & _
                            "A.SERVICETYPE = D.SERVICETYPE AND " & _
                            "A.UCCode Is Not Null And " & _
                            "A.CustId=B.CustId " & _
                            IIf(Len(strChoose) = 0, "", " And ") & strChoose & set34sql & "  AND A." & snoType & " = '" & strMNO & "'   " & _
                            " Group by A." & snoType
        
        If GetRS(rscheckSameNOSQL, checkSameNOSQL, cn, adUseClient, adOpenKeyset) = True Then
            SingalSumAmount = rscheckSameNOSQL(0)
        Else
            file(0).WriteLine " 問題金額 : 單號　" & strMNO
        End If
        
        rscheckSameNOSQL.Close
        Set rscheckSameNOSQL = Nothing
        GetSumAmout = SingalSumAmount
    Exit Function

ChkErr:
    Call ErrSub(Me.Name, "GetSumAmout")
End Function

Private Function GetStopYMDateString(ByVal strMNO As String, ByVal CardName As String, ByVal set34sql As String) As String
    Dim rsStopYMDateString As New ADODB.Recordset
    Dim rsStopYMDateSQL As String
    Dim strCardNameSQL As String
        If CardName <> "" Then strCardNameSQL = " And C.CARDNAME  ='" & CardName & "'"
        
        If intPara24 = 0 Then
           rsStopYMDateSQL = "SELECT E.StopYM   " & _
                    "From " & _
                    TableOwnerName & "SO033 A ," & _
                    TableOwnerName & "SO001 B ," & _
                    TableOwnerName & "SO002A C,  " & _
                    TableOwnerName & "SO002 D,  " & _
                    TableOwnerName & "SO106 E  " & _
                    "Where " & strBankId & strCardNameSQL & " AND " & _
                    "A.CustID  = C.CUSTID AND " & _
                    "A.CustID  = D.CUSTID AND " & _
                    "A.CustID  = E.CUSTID AND C.AccountNO  = E.AccountID  AND  " & _
                    "A.CustId=B.CustId And A.CancelFlag = 0 And " & _
                    "A.UCCode Is Not Null  AND  " & _
                    "A.SERVICETYPE = D.SERVICETYPE   AND A.AccountNo = C.AccountNo AND E.StopFlag = 0 " & _
                    IIf(Len(strChoose) = 0, "", " And ") & strChoose & set34sql & "  AND A.BillNO = '" & strMNO & "'   " & _
                    " Group by A.BillNO , E.AccountID, E.StopYM  ORDER BY  A.BillNO,SUBSTR(LPAD(E.STOPYM,6,'0'),3,4) || SUBSTR(LPAD(E.STOPYM,6,'0'),1,2) DESC  "
    
        Else
          rsStopYMDateSQL = "SELECT  E.StopYM    " & _
                    "From " & _
                    TableOwnerName & "SO033 A ," & _
                    TableOwnerName & "SO001 B ," & _
                    TableOwnerName & "SO002A C,  " & _
                    TableOwnerName & "SO002 D,  " & _
                    TableOwnerName & "SO106 E  " & _
                    "Where " & strBankId & strCardNameSQL & " AND " & _
                    "A.CustID  = C.CUSTID AND " & _
                    "A.CustID  = D.CUSTID AND " & _
                    "A.CustID  = E.CUSTID AND C.AccountNO  = E.AccountID  AND  " & _
                    "A.CustId=B.CustId And A.CancelFlag = 0 And " & _
                    "A.UCCode Is Not Null AND  " & _
                    "A.SERVICETYPE = D.SERVICETYPE   AND A.AccountNo = C.AccountNo  AND E.StopFlag = 0  " & _
                    IIf(Len(strChoose) = 0, "", " And ") & strChoose & set34sql & "  AND A.MediaBillNo= '" & strMNO & "'   " & _
                    " Group by  E.AccountID , E.StopYM ,  A.MediaBillNo  ORDER BY  A.MediaBillNo ,SUBSTR(LPAD(E.STOPYM,6,'0'),3,4) || SUBSTR(LPAD(E.STOPYM,6,'0'),1,2) DESC  "
    
                    
        End If
        
        If GetRS(rsStopYMDateString, rsStopYMDateSQL, cn, adUseClient, adOpenDynamic) = True Then
            If IsNull(rsStopYMDateString(0)) Then
                GetStopYMDateString = "NULL"
            Else
                GetStopYMDateString = rsStopYMDateString(0)
            End If
        Else
            GetStopYMDateString = "NULL"
        End If
    
    Exit Function
ChkErr:
    Call ErrSub(Me.Name, "GetStopYMDateString")
End Function
Function subCitiBankBatch() As Boolean
    On Error GoTo ChkErr
    Dim intTxtFileIndex As Integer
    Dim lngSubCount As Long
    Dim lngTotalAmount As Long
    Dim strFileName As String
        lngSubCount = 0
        lngTotalAmount = 0
        intTxtFileIndex = 1
        Call DefineRs3    ''建立記憶體 recordset 結構
        If Not BeginTran("", txtSpcNO, txtMerchantName, intTxtFileIndex, lngSubCount, lngTotalAmount, True) Then
            'Exit Function
        End If
        If lngSubCount > 0 Then
            strFileName = txtDataPath & "\" & txtSpcNO & Format(Date, "yymmdd") & "." & txtBatch
            objStorePara.uProcText = strFileName
            Set file(intTxtFileIndex) = fso.CreateTextFile(strFileName, True)
            Dim lngSetCount(4) As Long, lngSetTotalAmt(4) As Long
            strHeader = "FHD"
            strHeader = strHeader & Right(Format(Date, "YYYYMMDD"), 6)
            strHeader = strHeader & GetString(txtBatch, 2, GIRIGHT, True)
            strHeader = strHeader & GetString(lngSubCount, 8, GIRIGHT, True)
            strHeader = strHeader & GetString(lngTotalAmount, 13, GIRIGHT, True) & "00"
            '*****************************************************************************
            '#3305 將Head保留欄位長度由原來的116改為183 By Kin 2007/07/26
            strHeader = strHeader & Space(183)
            '*****************************************************************************
            Call subWriteLine(intTxtFileIndex)   ''''' 將取得的資料寫入文字檔
        End If
        subCitiBankBatch = True
    Exit Function
ChkErr:
    Call ErrSub(Me.Name, "subCitiBankBatch")
End Function
Private Sub DefineRs3()
    On Error GoTo ChkErr
    With rsDefTmp
        If .State = adStateOpen Then .Close
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        
        ''"""""""""""""""""""""""""""""""""
        .Fields.Append ("F01"), adBSTR, 3, adFldIsNullable
        .Fields.Append ("F02"), adBSTR, 3, adFldIsNullable
        .Fields.Append ("F03"), adBSTR, 12, adFldIsNullable
        .Fields.Append ("F04"), adBSTR, 19, adFldIsNullable
        .Fields.Append ("F05"), adBSTR, 2, adFldIsNullable
        
        .Fields.Append ("F06"), adBSTR, 2, adFldIsNullable
        .Fields.Append ("F07"), adBSTR, 12, adFldIsNullable
        .Fields.Append ("F08"), adBSTR, 13, adFldIsNullable
        .Fields.Append ("F09"), adBSTR, 16, adFldIsNullable
        .Fields.Append ("F10"), adBSTR, 3, adFldIsNullable
        
        .Fields.Append ("F11"), adBSTR, 6, adFldIsNullable
        .Fields.Append ("F12"), adBSTR, 12, adFldIsNullable
        .Fields.Append ("F13"), adBSTR, 2, adFldIsNullable
        .Fields.Append ("F14"), adBSTR, 4, adFldIsNullable
        .Fields.Append ("F15"), adBSTR, 8, adFldIsNullable
        
        .Fields.Append ("F16"), adBSTR, 8, adFldIsNullable
        .Fields.Append ("F17"), adBSTR, 15, adFldIsNullable
        .Fields.Append ("F18"), adBSTR, 6, adFldIsNullable
        .Fields.Append ("F19"), adBSTR, 2, adFldIsNullable
        .Fields.Append ("F20"), adBSTR, 2, adFldIsNullable
        '*********************************************************************
        '#3305 新增以1個保留欄位長度67 By Kin 2007/07/26
        .Fields.Append ("F21"), adBSTR, 67, adFldIsNullable
        '*********************************************************************
        .Open
    End With
    
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "DefineRs3")
End Sub
Private Function BeginTran3(ByRef lngSubCount As Long, ByRef lngTotal As Long, _
    ByVal strCardExpDate As String, theSingalAmount As Long, ByVal strSpcNo As String) As Boolean
    
    On Error GoTo ChkErr
    Dim strBillNO As String
    If intPara24 = 0 Then
         strBillNO = "BillNO"
    Else
         strBillNO = "MediaBillNO"
    End If
    With rsDefTmp
        .AddNew
        'A Detail Record固定為"FDT"
        .Fields("F01").Value = "FDT"
        'N 收單行庫代號，MerchantID(依各筆扣款資料之SO106.CardCode對應至CD037.CodeNo取CD037.SPCNO)前三碼
        .Fields("F02").Value = Right("000" & Trim(GetString(strSpcNo, 3)), 3)
        '空白
        .Fields("F03").Value = Space(12)
        '訂單編號(媒體單號或收費單號依據收費參數是否啟動'媒體多帳戶處理')
        .Fields("F04").Value = GetString(rsTmp(strBillNO), 19)
        'N 分期期數，不作分期則放空白(目前放空白)
        .Fields("F05").Value = Space(2)
        '空白
        .Fields("F06").Value = Space(2)
        '值域：auth: 授權成功 authfail: 授權失敗 sale: 購物成功 salefail: 購物失敗 cap: 請款成功 capfail: 請款失敗 settle: 請款結帳
        .Fields("F07").Value = Space(12)
        'N 含二位小數位(SO033.SHOULDAMT)
        .Fields("F08").Value = GetString(theSingalAmount, 11, GIRIGHT, True) & "00"
        lngTotal = lngTotal + theSingalAmount  '總金額=金額累加
        .Fields("F09").Value = GetString(rsTmp("AccountNo"), 16, GIRIGHT)
        '**************************************************************
        '#3305 第10個欄位原本是放000，現在改為3個空白 By Kin 2007/07/26
        .Fields("F10").Value = Space(3)
        '**************************************************************
        .Fields("F11").Value = Format(strCardExpDate, "YYYYMM")
        .Fields("F12").Value = Space(12)
        .Fields("F13").Value = Space(2)
        .Fields("F14").Value = Space(4)
        .Fields("F15").Value = Space(8)
        .Fields("F16").Value = GetString(Trim(txtMerchantName), 8, GIRIGHT, True)
        .Fields("F17").Value = GetString(strSpcNo, 15)
        .Fields("F18").Value = Space(6)
        .Fields("F19").Value = Space(2)
        .Fields("F20").Value = Space(2)
        '*****************************************************
        '#3305增加一個保留欄位，長度67 By Kin 2007/07/26
        .Fields("F21").Value = Space(67)
        '*****************************************************
        .Update
    End With
    BeginTran3 = True
    Exit Function
ChkErr:
    ErrSub Me.Name, "BeginTran3"
End Function


