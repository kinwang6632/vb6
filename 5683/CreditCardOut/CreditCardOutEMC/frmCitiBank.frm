VERSION 5.00
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCitiBank 
   BorderStyle     =   1  '單線固定
   ClientHeight    =   4575
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
   ScaleHeight     =   4575
   ScaleWidth      =   9240
   StartUpPosition =   1  '所屬視窗中央
   Begin MSComDlg.CommonDialog cnd1 
      Left            =   7950
      Top             =   1905
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
      Height          =   3825
      Left            =   180
      TabIndex        =   5
      Top             =   120
      Width           =   8895
      Begin VB.CheckBox chkCancel 
         Caption         =   "取消"
         Height          =   345
         Left            =   6075
         TabIndex        =   20
         Top             =   1230
         Width           =   2310
      End
      Begin VB.ComboBox cboMore 
         Height          =   315
         ItemData        =   "frmCitiBank.frx":0442
         Left            =   2010
         List            =   "frmCitiBank.frx":044F
         TabIndex        =   18
         Top             =   1320
         Width           =   2460
      End
      Begin Gi_Date.GiDate GiDAccept 
         Height          =   285
         Left            =   6075
         TabIndex        =   17
         Top             =   825
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
      Begin VB.TextBox txtMerchantName 
         Height          =   300
         Left            =   6060
         MaxLength       =   4
         TabIndex        =   15
         Top             =   420
         Width           =   2460
      End
      Begin VB.TextBox txtSpcNO 
         Height          =   300
         Left            =   2010
         TabIndex        =   13
         Top             =   435
         Width           =   1830
      End
      Begin VB.ComboBox cboSet 
         Height          =   315
         ItemData        =   "frmCitiBank.frx":0493
         Left            =   2010
         List            =   "frmCitiBank.frx":04A0
         TabIndex        =   11
         Top             =   870
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
         Height          =   1560
         HelpContextID   =   2
         Left            =   285
         TabIndex        =   6
         Top             =   2010
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
            Locked          =   -1  'True
            TabIndex        =   9
            ToolTipText     =   "請輸入字檔之路徑及檔名！"
            Top             =   480
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
            TabIndex        =   2
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
            Left            =   7320
            TabIndex        =   0
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
            Left            =   3015
            Locked          =   -1  'True
            TabIndex        =   1
            Top             =   945
            Width           =   4095
         End
         Begin VB.Label lblDatapath 
            AutoSize        =   -1  'True
            Caption         =   "資料檔位置（路徑 )"
            Height          =   195
            Left            =   1290
            TabIndex        =   8
            Top             =   555
            Width           =   1665
         End
         Begin VB.Label lblErrpath 
            AutoSize        =   -1  'True
            Caption         =   "問題參考檔位置（路徑 + 名稱）"
            Height          =   195
            Left            =   330
            TabIndex        =   7
            Top             =   990
            Width           =   2640
         End
      End
      Begin VB.Label Label5 
         Alignment       =   1  '靠右對齊
         Caption         =   "交易旗號"
         Height          =   285
         Left            =   1080
         TabIndex        =   19
         Top             =   1365
         Width           =   795
      End
      Begin VB.Label Label8 
         Alignment       =   1  '靠右對齊
         Caption         =   "端未機代號"
         Height          =   285
         Left            =   4215
         TabIndex        =   16
         Top             =   480
         Width           =   1755
      End
      Begin VB.Label Label7 
         Alignment       =   1  '靠右對齊
         Caption         =   "商店代號"
         Height          =   285
         Left            =   405
         TabIndex        =   14
         Top             =   510
         Width           =   1470
      End
      Begin VB.Label Label2 
         Caption         =   "送件日期"
         Height          =   210
         Left            =   5205
         TabIndex        =   12
         Top             =   900
         Width           =   780
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         Caption         =   "交易碼"
         Height          =   285
         Left            =   1065
         TabIndex        =   10
         Top             =   945
         Width           =   795
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "F2.開始"
      Height          =   375
      Left            =   255
      TabIndex        =   3
      Top             =   4065
      Width           =   1275
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&X)"
      Height          =   375
      Left            =   7605
      TabIndex        =   4
      Top             =   4050
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
Dim strBankId As String

Private Sub cmdCancel_Click()
On Error Resume Next
    Unload Me
End Sub

Private Sub cmdOK_Click()

 '''If Not ChkDTok Then Exit Sub
 If Not IsDataOk Then Exit Sub
 Call DefineRs   ''建立記憶體 recordset 結構
 
  Dim s As String
  s = txtDataPath.Text
  ''If Not ChkDirExist(s) Then MsgBox "路徑 " & s & " 不存在!", vbExclamation: Exit Sub
  If OpenFile(txtDataPath.Text, 0) = False Then Exit Sub
  If OpenFile(txtErrPath.Text, 1) = False Then Exit Sub
  
  Screen.MousePointer = vbHourglass
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
    Call ScrToRcd
    objStorePara.uUpdate = True
    Unload Me
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "cmdOK_Click")
End Sub

Private Sub Form_Load()
    On Error Resume Next
      Me.Caption = objStorePara.BankName & ""
      cboMore.ListIndex = 2
      GiDAccept.Text = Format(Date, "eemmdd")
      Call InitData
      RcdToScr
End Sub
Private Sub RcdToScr()
    On Error GoTo ChkErr
    Dim LogFile As TextStream
 
        If Dir(strPath & "\" & strPrgName & "Out.log") = "" Then Exit Sub
            Set LogFile = fso.OpenTextFile(strPath & "\" & strPrgName & "Out.log", ForReading)
            With LogFile
               If Not .AtEndOfStream Then txtSpcNO = .ReadLine & ""
               If Not .AtEndOfStream Then txtMerchantName = .ReadLine & ""
               If Not .AtEndOfStream Then txtDataPath = .ReadLine & ""   '資料檔
               If Not .AtEndOfStream Then txtErrPath = .ReadLine & ""   '問題檔
            End With
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
      strChoose = .ChooseStr
      strPath = .ErrPath
      txtSpcNO.Text = .uSpcNO
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
        With cnd1
                .FileName = txtErrPath.Text
                .Filter = "文字檔|*.txt"
                .InitDir = strPath
                .Action = 1
                txtDataPath = .FileName
        End With
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
      
        .Open
    End With
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "DefineRs")
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

  
    BeginTran = False
   lngTime = Timer
             
        strsql = "SELECT " & GetUseIndexStr("A", "ShouldAmt") & " " & _
                "A.BillNO, sum(A.ShouldAmt)  ShouldAmt,  " & _
                "C.AccountNO,C.CardExpDate,B.CUSTNAME  " & _
                "From SO033 A,SO001 B ,SO002 C  " & _
                "Where " & strBankId & " AND  " & _
                "A.CustID  = C.CUSTID AND " & _
                "A.CustId=B.CustId And A.CancelFlag = 0 And " & _
                "A.UCCode Is Not Null And A.ShouldAmt > 0 " & _
                IIf(Len(strChoose) = 0, "", " And ") & strChoose & _
                " Group by A.BillNO , C.AccountNO , C.CardExpDate , B.CUSTNAME "
             
                
    Set rsTmp = cn.Execute(strsql)
    lngIndex = 1
    blnFirst = True
    
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
 
        While Not rsTmp.EOF
        
           Dim strCardExpDate As String
           Dim iCED As Integer
           Dim dteCED As Date
           Dim strErrMessage As String
        
            If IsNull(rsTmp("AccountNO")) Then
              lngErrCount = lngErrCount + 1
              strErrMessage = "信用卡卡號空白 : 單號　" & rsTmp("BillNO") & " 客戶姓名 " & rsTmp.Fields("custname")
              file(1).WriteLine strErrMessage
              GoTo NextLoop
           End If
           
                      
           strCardExpDate = GetNullString(rsTmp("CardExpDate"), giStringV, giAccessDb)
           
           If strCardExpDate = "NULL" Then
               lngErrCount = lngErrCount + 1
               strErrMessage = "信用卡資料不正確 : 單號　" & rsTmp("BillNO") & " 客戶姓名 " & rsTmp.Fields("custname")
               file(1).WriteLine strErrMessage
               GoTo NextLoop
           End If
        
           strCardExpDate = rsTmp("CardExpDate")
           iCED = IIf(Len(strCardExpDate) = 5, 1, 2)
           strCardExpDate = Right(strCardExpDate, 4) & "/" & Mid(strCardExpDate, 1, iCED)
           dteCED = CDate(strCardExpDate)
           If dteCED < CDate(Format(GiDAccept.Text, "YYYY/MM")) Then
               lngErrCount = lngErrCount + 1
               strErrMessage = "信用卡過期 : 單號　" & rsTmp.Fields("Billno") & " 客戶姓名　" & rsTmp.Fields("custname") & " 信用卡號 " & rsTmp.Fields("AccountNO") & " 到期日　" & rsTmp.Fields("CardExpDate")
               file(1).WriteLine strErrMessage
                GoTo NextLoop
           End If
                      lngCount = lngCount + 1
                      With rsDefTmp
                           .AddNew
                           
                           .Fields("F110") = GetString(txtSpcNO.Text, 10, giLeft, False)
                           .Fields("F1118") = GetString(txtMerchantName.Text, 8, giLeft, False)
                           .Fields("F1937") = GetString(rsTmp("AccountNO"), 19, giLeft, True)
                           .Fields("F3845") = GetString(rsTmp("ShouldAmt"), 8, GIRIGHT, True)
                    
                           .Fields("F4653") = Space(8)
                           .Fields("F5455") = "0" & cboSet.ListIndex    ''  交易碼
                           .Fields("F5661") = Format(GiDAccept.Text, "EEMMDD")
                          
                           .Fields("F6265") = Right(rsTmp("CardExpDate"), 2) & IIf(Len(rsTmp("CardExpDate")) = 6, Left(rsTmp("CardExpDate"), 2), "0" & Left(rsTmp("CardExpDate"), 1))
                           .Fields("F6668") = Space(3)
                           .Fields("F6984") = Space(16)
                           .Fields("F8590") = Space(6)
                           .Fields("F9196") = Space(6)
                           .Fields("F97112") = GetString(rsTmp("BillNO"), 16, giLeft, False)
                           .Fields("F113118") = Space(6)
                           .Fields("F119") = chkCancel.Value
                           .Fields("F120") = cboMore.ListIndex
                     
                           .Update
                      End With
NextLoop:
                        rsTmp.MoveNext
        Wend
        
        '' 以下取得 header 字串
        
        strHeader = "Header"
        strHeader = strHeader & GetString(lngCount, 8, GIRIGHT, True)
        strHeader = strHeader & Space(106)
        
        subWriteLine
        
        msgResult lngCount, lngErrCount, lngTime      '顯示執行結果
        BeginTran = True
        On Error Resume Next
        rsTmp.Close
        Set rsTmp = Nothing
        CloseFS
     End If
Exit Function
ChkErr:
    CloseFS
    Call ErrSub(Me.Name, "BeginTran")
End Function

Private Function subWriteLine() As Boolean

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
         file(0).WriteLine (strHeader)    ''於文字檔填入首筆
         For lngLoop = 0 To .RecordCount - 1
               strContent = ""
              For lngLoopi = 0 To 15
                  strContent = strContent & rsDefTmp(lngLoopi)
              Next
              file(0).WriteLine (strContent)
              rsDefTmp.MoveNext
              DoEvents
         Next
         
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
        With LogFile
                .WriteLine (txtSpcNO.Text) '信用卡扣帳特店名稱
                .WriteLine (txtMerchantName.Text) '信用卡扣帳特店名稱
                .WriteLine (txtDataPath)         '資料檔路徑
                .WriteLine (txtErrPath.Text)     ' 問題參考檔路徑
        End With
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "ScrToRcd")
End Sub

Private Function IsDataOk() As Boolean
Dim strErrMsg As String

  If Len(txtSpcNO.Text & "") = 0 Then strErrMsg = "商站代號": txtSpcNO.SetFocus: GoTo Warning
  If Len(txtMerchantName.Text & "") = 0 Then strErrMsg = "端未機代號": txtMerchantName.SetFocus: GoTo Warning
  If Len(cboSet.Text) = 0 Then strErrMsg = "交易碼": cboSet.SetFocus: GoTo Warning
  If Len(GiDAccept.Text) = 0 Then strErrMsg = "送件日期": GiDAccept.SetFocus: GoTo Warning
If Len(cboMore.Text) = 0 Then strErrMsg = "交易旗號": cboMore.SetFocus: GoTo Warning

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
    Set rsTmp = Nothing
End Sub

Private Sub GiDAccept_GotFocus()
    If Len(GiDAccept.Text) = 0 Then GiDAccept.Text = Format(Date, "EE/MM/DD")
End Sub

