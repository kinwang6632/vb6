VERSION 5.00
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCitiBank 
   BorderStyle     =   1  '單線固定
   ClientHeight    =   5115
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
   ScaleHeight     =   5115
   ScaleWidth      =   9240
   StartUpPosition =   1  '所屬視窗中央
   Begin MSComDlg.CommonDialog cnd1 
      Left            =   7920
      Top             =   915
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
      TabIndex        =   13
      Top             =   345
      Width           =   8895
      Begin VB.CheckBox chkDuteDate 
         Caption         =   "信用卡過期資料一併產生"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   2040
         TabIndex        =   5
         Top             =   1995
         Width           =   2535
      End
      Begin VB.TextBox txtSpcNO 
         Height          =   300
         Left            =   2025
         MaxLength       =   9
         TabIndex        =   24
         Top             =   435
         Visible         =   0   'False
         Width           =   1830
      End
      Begin VB.TextBox txtMerchantName 
         Height          =   300
         Left            =   6060
         MaxLength       =   9
         TabIndex        =   12
         Text            =   "TBGB"
         Top             =   390
         Visible         =   0   'False
         Width           =   2460
      End
      Begin VB.TextBox txtStatement 
         Height          =   345
         Left            =   2025
         TabIndex        =   4
         Top             =   1530
         Width           =   6465
      End
      Begin VB.TextBox txtDiscountRate 
         Height          =   345
         Left            =   2025
         MaxLength       =   5
         TabIndex        =   1
         Text            =   "0"
         Top             =   1140
         Width           =   900
      End
      Begin VB.TextBox txtPayDate 
         Height          =   345
         Left            =   6075
         MaxLength       =   2
         TabIndex        =   3
         Text            =   "00"
         Top             =   1140
         Width           =   330
      End
      Begin VB.ComboBox cboSet 
         Height          =   315
         ItemData        =   "frmCitiBank.frx":0442
         Left            =   2025
         List            =   "frmCitiBank.frx":0452
         TabIndex        =   0
         Top             =   780
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
         TabIndex        =   14
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
            TabIndex        =   6
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
            Left            =   7335
            TabIndex        =   9
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
            TabIndex        =   7
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
            TabIndex        =   8
            Top             =   765
            Width           =   4095
         End
         Begin VB.Label lblDatapath 
            AutoSize        =   -1  'True
            Caption         =   "資料檔位置（路徑 )"
            Height          =   195
            Left            =   1290
            TabIndex        =   16
            Top             =   375
            Width           =   1665
         End
         Begin VB.Label lblErrpath 
            AutoSize        =   -1  'True
            Caption         =   "問題參考檔位置（路徑 + 名稱）"
            Height          =   195
            Left            =   330
            TabIndex        =   15
            Top             =   810
            Width           =   2640
         End
      End
      Begin Gi_Date.GiDate gdaSystemDate 
         Height          =   345
         Left            =   6060
         TabIndex        =   2
         Top             =   720
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
      Begin VB.Label Label7 
         Alignment       =   1  '靠右對齊
         Caption         =   "信用卡扣帳特店代碼"
         Height          =   285
         Left            =   45
         TabIndex        =   25
         Top             =   510
         Visible         =   0   'False
         Width           =   1890
      End
      Begin VB.Label Label5 
         Alignment       =   1  '靠右對齊
         Caption         =   "系統日期"
         Height          =   285
         Left            =   4215
         TabIndex        =   23
         Top             =   780
         Width           =   1755
      End
      Begin VB.Label Label9 
         Caption         =   "(DD)"
         Height          =   240
         Left            =   6465
         TabIndex        =   22
         Top             =   1170
         Width           =   1080
      End
      Begin VB.Label Label8 
         Caption         =   "信用卡扣帳特店名稱"
         Height          =   285
         Left            =   4215
         TabIndex        =   21
         Top             =   435
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.Label Label4 
         Caption         =   "客戶帳單內容"
         Height          =   210
         Left            =   750
         TabIndex        =   20
         Top             =   1620
         Width           =   1185
      End
      Begin VB.Label Label3 
         Alignment       =   1  '靠右對齊
         Caption         =   "折扣率"
         Height          =   210
         Left            =   1215
         TabIndex        =   19
         Top             =   1245
         Width           =   720
      End
      Begin VB.Label Label2 
         Caption         =   "請款日期"
         Height          =   210
         Left            =   5190
         TabIndex        =   18
         Top             =   1185
         Width           =   780
      End
      Begin VB.Label Label1 
         Caption         =   "結帳種類"
         Height          =   285
         Left            =   1140
         TabIndex        =   17
         Top             =   870
         Width           =   795
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "F2.開始"
      Height          =   375
      Left            =   315
      TabIndex        =   10
      Top             =   4470
      Width           =   1275
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&X)"
      Height          =   375
      Left            =   7665
      TabIndex        =   11
      Top             =   4455
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

Dim rsCD037SpcNO  As New ADODB.Recordset  '' 取得 SPCNO 為唯一值的記錄，以同樣SPCNO的卡別一次取出
Dim rsCD037 As New ADODB.Recordset
Dim s As String
Dim intTxtFileIndex  As Integer
Dim lngTime As Long


lngTime = Timer
 If Not IsDataOk Then Exit Sub
 '' Call DefineRs   ''建立記憶體 recordset 結構
 
  s = txtDataPath.Text
  If Not CreateObject("Scripting.FileSystemObject").FolderExists(s) Then MsgBox "路徑 " & s & " 不存在!", vbExclamation: Exit Sub

  If OpenFile(txtErrPath, 0, True) = False Then Exit Sub
  
  ''""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
  ' 取得 SPCNO 為唯一值的記錄，以同樣SPCNO的卡別一次取出
  ''""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
  rsCD037SpcNO.CursorLocation = adUseClient
  rsCD037SpcNO.Open _
      "SELECT DISTINCT SPCNO FROM " & TableOwnerName & "CD037  WHERE SpcNO IS NOT NULL  ", _
      cn, adOpenKeyset, adLockReadOnly
  If rsCD037SpcNO.EOF Or rsCD037SpcNO.BOF Then
        MsgBox "信用卡種類的扣帳特店代碼必需設值 "
        rsCD037SpcNO.Close
        Exit Sub
  End If

 ' ''""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
 
  rsCD037.CursorLocation = adUseClient
  
  rsCD037.Open _
          "SELECT * FROM " & TableOwnerName & "CD037 " & _
          "WHERE " & _
          "Spcno IS NOT NULL AND MerchantName IS NOT NULL " & _
          "ORDER BY CodeNO", _
          cn, adOpenKeyset, adLockReadOnly
          
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
                     ''MsgBox " 信用卡別為 " & rsCD037("Description") & "　資料轉出　!!", vbOKOnly + vbCritical, "訊息"
                End If
                rsCD037.MoveNext
                blnNewRec = True
     Loop
     
     '' 將收集所得的資料填入文字檔
     
     If lngSubCount > 0 Then
           rsCD037.MoveFirst
 ''       If OpenFile(txtDataPath, intTxtFileIndex, True, cboSet.ListIndex, rsCD037("MerchantName"), rsCD037("SPCNO")) = False Then

            If OpenFile(txtDataPath, intTxtFileIndex, True, cboSet.ListIndex, rsCD037("MerchantName"), rsCD037("CardID")) = False Then
                MsgBox "無法建立登錄特店代碼為 " & rsCD037("SPCNO") & " 的文字檔 !!", vbOKOnly + vbCritical, "訊息 "
                Exit Sub
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

Call CloseFS(intTxtFileIndex)
Call ScrToRcd
objStorePara.uUpdate = True

rsCD037.Close
Set rsCD037 = Nothing
rsCD037SpcNO.Close
Set rsCD037SpcNO = Nothing
If lngErrCount > 0 Then Call Shell("NotePad " & txtErrPath.Text, vbNormalFocus)
Call msgResult(lngCount, lngErrCount, lngTime)      '顯示執行結果
Unload Me

Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "cmdOK_Click")
End Sub

Private Sub Form_Load()

    On Error Resume Next
    blnSameFile = False
    lngErrCount = 0
    lngCount = 0
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
            
              '' If Not .AtEndOfStream Then txtMerchantName.Text = .ReadLine & ""   '信用卡扣帳特店名稱
              '' If Not .AtEndOfStream Then txtPayDate.Text = .ReadLine & ""   '請款日期
               If Not .AtEndOfStream Then txtDiscountRate.Text = .ReadLine & ""   '折扣率
               If Not .AtEndOfStream Then txtStatement.Text = .ReadLine & ""  '客戶帳單內容
'              If Not .AtEndOfStream Then txtInVoice.Text = .ReadLine & ""   '統一編號

               If Not .AtEndOfStream Then txtDataPath = .ReadLine & ""   '資料檔
               If Not .AtEndOfStream Then txtErrPath = .ReadLine & ""   '問題檔
                       
                  
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
  
  
  ''2005/01/10  加入判斷式  取信用卡有效年月最大值
  Dim strAccountID As String
  Dim lngCustID As Long
  Dim CountSameCustIDBillNO As Integer
  
  Dim theSingalAmount As Long  '' 2005/02/02 用來儲存單筆SO033資料，
  
  
  Dim lngerror As Long:    lngerror = 0
  
  strAccountID = ""
  lngCustID = 0
  CountSameCustIDBillNO = 0
  
  
     
  BeginTran = False
    
'' 2005/02/02 不管正負項了  一律都出來
    
'   If cboSet.ListIndex = 2 Then
'       set34 = "  AND A.ShouldAmt > 0 "
'   Else
'        set34 = "  AND A.ShouldAmt < 0 "
'   End If
   set34 = "  "
   
   ''' 2003/07/23 新增客戶編號(CustID)以及媒體單號(MediaBillNo)名稱
   ''' 2003/08/14 改成SO002A 為資料來源，並且將GROUP BY 的字句割開，分成使用媒體單號與不使用的
   '''使用以MediaBillNO 為主，沒有以 Billno為主 ，全部的客戶編號資訊 則以  重新取得的SO002A 為準，CustStatusCode  在 SO002A 沒有 因此要將 SO002 串進來

'       strsql = "SELECT " & GetUseIndexStr("A", "ShouldAmt") & " " & _
'                "A.BillNO, sum(A.ShouldAmt)  ShouldAmt,  " & _
'                "C.AccountNO,C.CardExpDate,B.CUSTNAME ,A.MediaBillNo,A.CUSTID   " & _
'                "From " & _
'                TableOwnerName & "SO033 A ," & _
'                TableOwnerName & "SO001 B ," & _
'                TableOwnerName & "SO002A C  " & _
'                "Where " & strBankId & " AND  " & _
'                "C.CARDNAME  ='" & strCardName & "' AND " & _
'                "A.CustID  = C.CUSTID AND " & _
'                "A.CustId=B.CustId And A.CancelFlag = 0 And " & _
'                "A.UCCode Is Not Null And A.ShouldAmt > 0 " & _
'                IIf(Len(strChoose) = 0, "", " And ") & strChoose & set34 & _
'                " Group by A.BillNO , C.AccountNO , C.CardExpDate , B.CUSTNAME, A.MediaBillNo,A.CustID  "

''  此段於 2004/08/12 保留，其中的 C.CardExpDate  改成取 SO106 的 StopYM
'' *****************************************************************************************************
'  If intPara24 = 0 Then
'       strsql = "SELECT " & GetUseIndexStr("A", "ShouldAmt") & " " & _
'                "A.BillNO, sum(A.ShouldAmt)  ShouldAmt,  " & _
'                "C.AccountNO,C.CardExpDate   " & _
'                "From " & _
'                TableOwnerName & "SO033 A ," & _
'                TableOwnerName & "SO001 B ," & _
'                TableOwnerName & "SO002A C,  " & _
'                TableOwnerName & "SO002 D  " & _
'                "Where " & strBankId & " AND  " & _
'                "C.CARDNAME  ='" & strCardName & "' AND " & _
'                "A.CustID  = C.CUSTID AND " & _
'                "A.CustID  = D.CUSTID AND " & _
'                "A.CustId=B.CustId And A.CancelFlag = 0 And " & _
'                "A.UCCode Is Not Null And A.ShouldAmt > 0  AND " & _
'                "A.SERVICETYPE = D.SERVICETYPE   AND A.AccountNo = C.AccountNo  " & _
'                IIf(Len(strChoose) = 0, "", " And ") & strChoose & set34 & _
'                " Group by A.BillNO , C.AccountNO , C.CardExpDate  "
'    Else
'      strsql = "SELECT " & GetUseIndexStr("A", "ShouldAmt") & " " & _
'                "sum(A.ShouldAmt)  ShouldAmt,  " & _
'                "C.AccountNO,C.CardExpDate ,A.MediaBillNo   " & _
'                "From " & _
'                TableOwnerName & "SO033 A ," & _
'                TableOwnerName & "SO001 B ," & _
'                TableOwnerName & "SO002A C,  " & _
'                TableOwnerName & "SO002 D  " & _
'                "Where " & strBankId & " AND  " & _
'                "C.CARDNAME  ='" & strCardName & "' AND " & _
'                "A.CustID  = C.CUSTID AND " & _
'                "A.CustID  = D.CUSTID AND " & _
'                "A.CustId=B.CustId And A.CancelFlag = 0 And " & _
'                "A.UCCode Is Not Null And A.ShouldAmt > 0  AND " & _
'                "A.SERVICETYPE = D.SERVICETYPE   AND A.AccountNo = C.AccountNo  " & _
'                IIf(Len(strChoose) = 0, "", " And ") & strChoose & set34 & _
'                " Group by  C.AccountNO , C.CardExpDate ,  A.MediaBillNo "
'    End If

'' *****************************************************************************************************
    ''  2004/09/24 加入SO106.StopFlag = 0 的條件值 防止出現兩筆
    ''  2004/09/24 還是將整段mark掉好了，直接取SO106就好了
    ''  以下的兩段SQL 將其中的 A.ShouldAmt > 0  的條件拿掉
    
    If intPara24 = 0 Then
       strsql = "SELECT " & GetUseIndexStr("A", "ShouldAmt") & " " & _
                "A.BillNO, sum(A.ShouldAmt)  ShouldAmt,  " & _
                "E.AccountID  AccountNO,E.StopYM CardExpDate  " & _
                "From " & _
                TableOwnerName & "SO033 A ," & _
                TableOwnerName & "SO001 B ," & _
                TableOwnerName & "SO002A C,  " & _
                TableOwnerName & "SO002 D,  " & _
                TableOwnerName & "SO106 E  " & _
                "Where " & strBankId & " AND  " & _
                "C.CARDNAME  ='" & strCardName & "' AND " & _
                "A.CustID  = C.CUSTID AND " & _
                "A.CustID  = D.CUSTID AND " & _
                "A.CustID  = E.CUSTID AND C.AccountNO  = E.AccountID  AND  " & _
                "A.CustId=B.CustId And A.CancelFlag = 0 And " & _
                "A.UCCode Is Not Null  AND  " & _
                "A.SERVICETYPE = D.SERVICETYPE   AND A.AccountNo = C.AccountNo AND E.StopFlag = 0 " & _
                IIf(Len(strChoose) = 0, "", " And ") & strChoose & set34 & _
                " Group by A.BillNO , E.AccountID, E.StopYM  ORDER BY  A.BillNO,E.StopYM DESC  "
        '        " Group by A.BillNO , E.AccountID, E.StopYM  ORDER BY  A.BillNO,SUBSTR(LPAD(E.STOPYM,6,'0'),3,4) || SUBSTR(LPAD(E.STOPYM,6,'0'),1,2) DESC  "
          
               ' " Group by A.BillNO , C.AccountNO , E.StopYM  "  ''這一段20050110 調掉的為了取SO106 的資料
    Else
      strsql = "SELECT " & GetUseIndexStr("A", "ShouldAmt") & " " & _
                "sum(A.ShouldAmt)  ShouldAmt,  " & _
                "E.AccountID  AccountNO,E.StopYM CardExpDate  ,A.MediaBillNo   " & _
                "From " & _
                TableOwnerName & "SO033 A ," & _
                TableOwnerName & "SO001 B ," & _
                TableOwnerName & "SO002A C,  " & _
                TableOwnerName & "SO002 D,  " & _
                TableOwnerName & "SO106 E  " & _
                "Where " & strBankId & " AND  " & _
                "C.CARDNAME  ='" & strCardName & "' AND " & _
                "A.CustID  = C.CUSTID AND " & _
                "A.CustID  = D.CUSTID AND " & _
                "A.CustID  = E.CUSTID AND C.AccountNO  = E.AccountID  AND  " & _
                "A.CustId=B.CustId And A.CancelFlag = 0 And " & _
                "A.UCCode Is Not Null AND  " & _
                "A.SERVICETYPE = D.SERVICETYPE   AND A.AccountNo = C.AccountNo  AND E.StopFlag = 0  " & _
                IIf(Len(strChoose) = 0, "", " And ") & strChoose & set34 & _
                " Group by  E.AccountID , E.StopYM ,  A.MediaBillNo  ORDER BY  A.MediaBillNo ,E.StopYM DESC  "
          '     " Group by  E.AccountID , E.StopYM ,  A.MediaBillNo  ORDER BY  A.MediaBillNo ,SUBSTR(LPAD(E.STOPYM,6,'0'),3,4) || SUBSTR(LPAD(E.STOPYM,6,'0'),1,2) DESC  "
           '
               ' " Group by  C.AccountNO , E.StopYM ,  A.MediaBillNo  "  ''這一段20050110 調掉的為了取SO106 的資料
                
    End If
    
    
    
 ''  2003/08/14 改成 再將 mediabillno   資料取出 這裏這一次多取一個 BillNO 然後填入當mediabillno為null 值的時候，填入媒體單號
    If intPara24 = 1 Then
    
                        ''這一段直接取前述的主 SQL (strsql)  拿掉 group  ，加上 Billno
                        
                        strMediaBillNONull = "SELECT " & GetUseIndexStr("A", "ShouldAmt") & " " & _
                                                           "C.AccountNO,A.MediaBillNo,A.BillNO   " & _
                                                            "From " & _
                                                            TableOwnerName & "SO033 A ," & _
                                                            TableOwnerName & "SO001 B ," & _
                                                            TableOwnerName & "SO002A C,  " & _
                                                            TableOwnerName & "SO002 D  " & _
                                                            "Where " & strBankId & " AND  " & _
                                                            "C.CARDNAME  ='" & strCardName & "' AND " & _
                                                            "A.CustID  = C.CUSTID AND " & _
                                                            "A.CustID  = D.CUSTID AND " & _
                                                            "A.CustId=B.CustId And A.CancelFlag = 0 And " & _
                                                            "A.UCCode Is Not Null  AND " & _
                                                            "A.SERVICETYPE=D.SERVICETYPE  AND A.AccountNo = C.AccountNo  " & _
                                                            IIf(Len(strChoose) = 0, "", " And ") & strChoose & set34 & "  AND A.MEDIABILLNO IS NULL"
                                                            
                                                            
                            strChooseMultiAcc = " A.CancelFlag = 0 AND " & _
                                                              "A.UCCode Is Not Null  " & _
                                                              IIf(Len(strChooseMultiAcc) = 0, "", " And ") & strChooseMultiAcc
                
                        If GetRS(rsMedioBillnoNull, strMediaBillNONull, cn) Then
                                     Do While Not rsMedioBillnoNull.EOF
                                          If IsNull(rsMedioBillnoNull("MediaBillNo")) Then
                                                  strSequenceNumber = GetInvoiceNo2("SO033_MediaBillNo", cn)
                                                  cn.Execute ( _
                                                                            "UPDATE " & TableOwnerName & "SO033 A SET " & _
                                                                            "A.MediaBillNo  ='" & strSequenceNumber & "'  WHERE   " & _
                                                                            "A.BillNo = '" & rsMedioBillnoNull("BillNO") & "'   AND  " & strChooseMultiAcc & " AND " & _
                                                                            "A.AccountNO ='" & rsMedioBillnoNull("AccountNO") & "' ")
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
    Else
        blnNodata = False
        Dim strSN As String  '' 指定序列號
        Dim lngtotal As Long '' 累計標頭檔 Total Amount 欄位總出單金額
        Dim lngDoubleMediaNO As Long '' 這個值記錄重覆出單的單子件數

        lngDoubleMediaNO = 0
        If pNewRec = False Then
             lngtotal = 0
             lngSubCount = 0
        Else
            lngtotal = pTotalAmount
            lngSubCount = pSubCount
        End If
                                  
        While Not rsTmp.EOF
        
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
                strSopYMMaxString = GetStopYMDateString(strAccountID, strCardName, set34)
                lngerror = 45
           End If
            '' ************************************************************************
      '     ''2005/01/26將整理好的資料  依照單據排序之後  將同單號正負項合併
           '''Call LogSameNOPMAmount(rsTmp(strBillNOField), strCardName, set34)
           ''2005/02/02 將整理好的資料  依照單據排序之後  將同單號正負項合併
           
           lngerror = 5
             Call GetRS(rsCustIDErr, "SELECT  SO001.CUSTNAME , SO002A.CUSTID FROM " & TableOwnerName & "SO002A," & TableOwnerName & "SO001  " & _
                                  "WHERE " & _
                                  "SO002A.AccountNO ='" & rsTmp("AccountNO") & "'  AND  SO002A.CUSTID = SO001.CUSTID ", cn)
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
             strCardExpDate = Right(strSopYMMaxString, 4) & "/" & Replace(strSopYMMaxString, Right(strSopYMMaxString, 4), "")
             lngerror = 10
           
           dteCED = CDate(strCardExpDate)
           If dteCED < CDate(Format(gdaSystemDate.Text, "YYYY/MM")) Then
               lngErrCount = lngErrCount + 1
               strErrMessage = "信用卡過期 : 單號　" & rsTmp.Fields(strBillNOField) & " 客戶姓名　" & rsCustIDErr("custname") & " 信用卡號 " & rsTmp.Fields("AccountNO") & " 到期日　" & rsTmp.Fields("CardExpDate") & " 客戶編號  " & rsCustIDErr("CustID")
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
        ''   Else
               lngCount = lngCount + 1
'               lngtotal = lngtotal + rsTmp.Fields("ShouldAmt")
               lngSubCount = lngSubCount + 1
               With rsDefTmp
                    .AddNew
                    .Fields("SequenceNo") = Format(lngSubCount, "000000")
                    '' SequenceNo(1-6)
                    '' .Fields("MerchantNo") = Format(txtSpcNO.Text, "000000000")
                    .Fields("MerchantNo") = Format(strSpcNo, "000000000")
                    ''MerchantNo(7-15)
                    .Fields("DiscountRate") = Format(txtDiscountRate.Text, "00000")
                    ''DiscountRate(16-20)
                    .Fields("CardNo") = IIf(Len(rsTmp("AccountNO")) > 0, Format(rsTmp("AccountNO"), "0000000000000000"), "0000000000000000")
                    ''CardNo(21-36)
                  ''  .Fields("ExpiryDate") = IIf(IsNull(rsTmp("CardExpDate")), "000000", Format(rsTmp("CardExpDate"), "000000"))
                  '' 20050205 MARK以上的那一行，改成以下的
                  strCardExpDate = Replace(strCardExpDate, "/", "")
                  If Len(strCardExpDate) = 5 Then
                      strCardExpDate = Left(strCardExpDate, 4) & "0" & Right(strCardExpDate, 1)
                  End If
                  ''.Fields("ExpiryDate") = Replace(strCardExpDate, "/", "")
                  .Fields("ExpiryDate") = Right(strCardExpDate, 2) & Left(strCardExpDate, 4)
                    ''ExpiryDate(37-42)
                    .Fields("SequenceNo2") = Format(lngSubCount, "000000000000000")
                    ''SequenceNo2(43-57)
                    .Fields("Filler1") = Space(8)
                    '' Filler(58-65)
                    ''2005/02/02 以下金額欄位值改成重取SO033的金額 避免程式取得重覆金額  !!
                    ''.Fields("Amount") = Format(rsTmp("ShouldAmt"), "0000000")
                    .Fields("Amount") = Format(theSingalAmount, "0000000")
                    lngtotal = lngtotal + theSingalAmount
                    '' Amount(66-72)
                     Dim strSD As String
                     Dim intlength As Integer
                     
                     '""轉換成為正確長度的全形字形'''''''''''''''''''''''''
                     strSD = StrConv(txtStatement.Text, vbFromUnicode)
                     intlength = LenB(strSD)
                     'strSD = txtStatement.Text & IIf(intlength < 38, Space(38 - intlength), "")
                     strSD = txtStatement.Text & Space(38)
                   '  MsgBox LenB(StrConv(strSD, vbFromUnicode))
                     
                 ''  strSD = StrConv(LeftB(StrConv(strSD, vbFromUnicode), 38), vbUnicode)
                     ''""""""""""""""""""""""""""""""""""""""""""""""""""""
                   .Fields("StatementDescription") = StrConv(LeftB(StrConv(StrConv(strSD, vbWide), vbFromUnicode), 38), vbUnicode)
                    
                    '.Fields("StatementDescription") = txtStatement.Text & Space(38 - Len(txtStatement.Text))
                    ''StatementDescription(73-110)
                    .Fields("Filler2") = Space(2)
                    
                   '' 2003/07/23 以下進行單號填值的操作，判斷是否以媒體單號的功能入單
                   '' 如果是則以媒體單號入單，否則轉換成舊單號之後再入
                   '' 2003/08/14 以下的內容，取得單號內容，如果啟用媒體的功能，就直接取MediaBillNO
                   
                   If intPara24 = 0 Then
                             Dim strBillNOOld As String
                             strBillNOOld = Trim(CStr(Val(Left(rsTmp("BillNO"), 6)) - 191100)) & _
                                            Mid(rsTmp("BillNO"), 7, 1) & Format(Right(rsTmp("BillNO"), 6), "000000")
                            .Fields("MerchantDescription") = strBillNOOld & Space(20 - Len(strBillNOOld))
                   Else
'                          If Not IsNull(rsTmp("MediaBillNo")) Then
'                          lngDoubleMediaNO = lngDoubleMediaNO + 1
'                             .Fields("MerchantDescription") = GetString(rsTmp("MediaBillNo"), 20, giLeft, False)
                             
                       ''     2003/08/06   此部份的程式碼不再需要使用了，因為一定會媒體單號
                       ''      file(0).WriteLine "單號   [" & rsTmp("BillNo") & "]  重覆出單；" & _
                                                       "媒體單號 [" & rsTmp("MediaBillNo") & "]；" & _
                                                       "客戶編號 [" & rsTmp("CustID") & "]  "
'                          Else
'                              strSequenceNumber = GetInvoiceNo2("SO033_MediaBillNo", cn)
                          ''    .Fields("MerchantDescription") = GetString(strSequenceNumber, 20, giLeft, False)
                          '' 直接取MediaBillNO
                           .Fields("MerchantDescription") = GetString(rsTmp("MediaBillNo"), 20, giLeft, False)
                    ''         cn.Execute ("UPDATE SO033 A SET MediaBillNo  ='" & strSequenceNumber & "'  WHERE   BillNo = '" & rsTmp("BillNO") & "'   AND  " & strChooseMultiAcc & " AND CustID = " & rsTmp("CUSTID") & " AND AccountNO ='" & rsTmp("AccountNO") & "'")
                   ''    End If
                   End If
                   
                   '' 這一段註解是emc 的 貼來參考用
'                                     If intPara24 = 0 Then
'                                             .Fields("F97112") = GetString(rsTmp("BillNO"), 16, giLeft, False)
'                                    Else
'                                         If Not IsNull(rsDoubleMediaNO("MediaBillNo")) Then
'                                            lngDoubleMediaNO = lngDoubleMediaNO + 1
'                                            .Fields("F97112") = GetString(rsDoubleMediaNO("MediaBillNo"), 16, giLeft, False)
'                                             file(1).WriteLine "單號   [" & rsTmp("BillNo") & _
'                                             "]  重覆出單，媒體單號 [" & rsDoubleMediaNO("MediaBillNo") & _
'                                             "] 客戶編號 [" & rsTmp("CustID") & "]  "
'                                        Else
'                                             strSequenceNumber = GetInvoiceNo2("SO033_MediaBillNo", cn)
'                                            .Fields("F97112") = GetString(strSequenceNumber, 16, giLeft, False)
'                                             cn.Execute ("UPDATE SO033 A SET MediaBillNo  ='" & strSequenceNumber & "'  WHERE   BillNo = '" & rsTmp("BillNO") & "'   AND " & strChooseMultiAcc & " AND CustID = " & rsTmp("CUSTID") & " AND AccountNO ='" & rsTmp("AccountNO") & "'")
'                                         End If
'                                    End If
'
                  ''    ''""""""""""""""""""""""""""""""""""""""""""""""""""""     ''""""""""""""""""""""""""""""""""""""""""""""""""""""
                  ''    ''""""""""""""""""""""""""""""""""""""""""""""""""""""     ''""""""""""""""""""""""""""""""""""""""""""""""""""""
                   
                   '' .Fields("MerchantDescription") = rsTmp("BillNO") & Space((20 - Len(rsTmp("BillNO"))))
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
         ''  End If
NextLoop:
        rsTmp.MoveNext
        Wend
        
        
        pSubCount = lngSubCount
        pTotalAmount = lngtotal
        
        BeginTran = True
        On Error Resume Next
        rsTmp.Close
        Set rsTmp = Nothing
        '' 203/08/05 這一段不需要了因為一定會有重覆出單的
'         If lngDoubleMediaNO > 0 Then
'                    Call MsgBox("出單的資料有一些是重覆的，請查閱記錄檔確認所需的資料內容，避免重覆出單  !!", vbExclamation + vbOKOnly, "重覆出單警示")
'                    Call Shell("NotePad " & txtErrPath.Text, vbNormalFocus)
'         End If
  End If
  
Exit Function
ChkErr:
    Call CloseFS(intIndexTxt)
    Call ErrSub(Me.Name, "BeginTran" & lngerror & "-" & strAccountID)
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
              For lngLoopi = 0 To 20
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
              ''  .WriteLine (txtMerchantName.Text) '信用卡扣帳特店名稱
              '' .WriteLine (txtPayDate.Text)      '請款日期
                .WriteLine (txtDiscountRate.Text) '折扣率
                .WriteLine (txtStatement.Text)    '客戶帳單內容
                '.WriteLine (txtInVoice.Text)     '統一編號
                .WriteLine (txtDataPath)         '資料檔路徑
                .WriteLine (txtErrPath.Text)     ' 問題參考檔路徑
        End With
        LogFile.Close
        Set LogFile = Nothing
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "ScrToRcd")
End Sub

Private Function IsDataOk() As Boolean
Dim strErrMsg As String

  If cboSet.ListIndex = 0 Or cboSet.ListIndex = 1 Then
        MsgBox "結帳種類 (1.授權 2.請款) 目前無法使用 !!", vbOKOnly + vbInformation, "訊息　"
        cboSet.SetFocus
        Exit Function
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
  
'  If Not IsNumeric(txtSpcNO.Text) Then
'      MsgBox " 信用卡扣帳特店代碼必需為數字　!! ", vbOKOnly + vbInformation, "訊息　"
'      txtPayDate.SetFocus
'      Exit Function
'  End If
  
 '' If Len(txtSpcNO.Text & "") = 0 Then strErrMsg = "信用卡扣帳特店代碼": txtSpcNO.SetFocus: GoTo Warning
 '' If Len(txtMerchantName.Text & "") = 0 Then strErrMsg = "信用卡扣帳特店名稱": txtMerchantName.SetFocus: GoTo Warning
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
    
'checkSameNOSQL = "SELECT SUM(A.ShouldAmt)  ShouldAmt  " & _
'                                        "From " & _
'                                        TableOwnerName & "SO033 A ," & _
'                                        TableOwnerName & "SO001 B ," & _
'                                        TableOwnerName & "SO002A C,  " & _
'                                        TableOwnerName & "SO002 D  " & _
'                                        "Where " & strBankId & " AND  " & _
'                                        "C.CARDNAME  ='" & CardName & "' AND " & _
'                                        "A.CustID  = C.CUSTID AND " & _
'                                        "A.CustID  = D.CUSTID AND " & _
'                                        "A.CustId=B.CustId And A.CancelFlag = 0 And " & _
'                                        "A.UCCode Is Not Null AND " & _
'                                        "A.SERVICETYPE = D.SERVICETYPE   AND A.AccountNo = C.AccountNo  AND C.StopFlag = 0  " & _
'                                        IIf(Len(strChoose) = 0, "", " And ") & strChoose & set34sql & "  AND A." & snoType & " = '" & strMNO & "'   " & _
'                                        " Group by A." & snoType & " , C.AccountNO  "
                                
 '' 將 C.StopFlag = 0  這個拿掉，避免搜不到資料
                                        
checkSameNOSQL = "SELECT SUM(A.ShouldAmt)  ShouldAmt  " & _
                                        "From " & _
                                        TableOwnerName & "SO033 A,  " & _
                                        TableOwnerName & "SO002 D  " & _
                                        "Where " & _
                                        "A.CancelFlag = 0 And " & _
                                        "A.CustID  = D.CUSTID AND " & _
                                        "A.SERVICETYPE = D.SERVICETYPE   AND   " & _
                                        "A.UCCode Is Not Null  " & _
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
 If intPara24 = 0 Then
       rsStopYMDateSQL = "SELECT E.StopYM   " & _
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
                "Where " & strBankId & " AND  " & _
                "C.CARDNAME  ='" & CardName & "' AND " & _
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
        GetStopYMDateString = rsStopYMDateString(0)
    Else
        GetStopYMDateString = ""
    End If
    
    Exit Function
ChkErr:
    Call ErrSub(Me.Name, "GetStopYMDateString")
End Function
