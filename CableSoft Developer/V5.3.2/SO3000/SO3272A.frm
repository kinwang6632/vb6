VERSION 5.00
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{A0C51418-555F-11D3-BC10-0050BA01A3A9}#9.8#0"; "GiMulti.ocx"
Object = "{6DDA0569-6454-11D3-9612-0048544B09DF}#4.1#0"; "GiTime.ocx"
Object = "{074771B7-6ADE-11D5-B178-00485456B028}#2.0#0"; "csMulti.ocx"
Begin VB.Form frmSO3272A 
   BorderStyle     =   1  '單線固定
   Caption         =   "信用卡扣款資料產生作業 [SO3272A]"
   ClientHeight    =   7245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11070
   Icon            =   "SO3272A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   11070
   StartUpPosition =   1  '所屬視窗中央
   Begin VB.Frame Frame1 
      Caption         =   "查詢條件"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   6675
      Left            =   150
      TabIndex        =   31
      Top             =   45
      Width           =   10665
      Begin VB.ComboBox cboBillHeadFmt 
         Height          =   300
         ItemData        =   "SO3272A.frx":0442
         Left            =   2310
         List            =   "SO3272A.frx":0444
         Style           =   2  '單純下拉式
         TabIndex        =   1
         Top             =   720
         Width           =   1905
      End
      Begin VB.ComboBox cboCreditCardType 
         Height          =   300
         ItemData        =   "SO3272A.frx":0446
         Left            =   5550
         List            =   "SO3272A.frx":0448
         Style           =   2  '單純下拉式
         TabIndex        =   2
         Top             =   360
         Width           =   2715
      End
      Begin VB.Frame fraTCB 
         Enabled         =   0   'False
         Height          =   855
         Left            =   8475
         TabIndex        =   36
         Top             =   555
         Width           =   1875
         Begin VB.OptionButton optNotTCB 
            Caption         =   "他行卡"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   165
            TabIndex        =   13
            Top             =   555
            Width           =   1335
         End
         Begin VB.OptionButton optTCB 
            Caption         =   "中華商銀自行卡"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   150
            TabIndex        =   12
            Top             =   270
            Value           =   -1  'True
            Width           =   1590
         End
      End
      Begin VB.CheckBox chkChinaBank 
         Caption         =   "以中華商銀格式出單"
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   8475
         TabIndex        =   11
         Top             =   360
         Width           =   1965
      End
      Begin VB.Frame Frame2 
         Caption         =   " 大樓依據"
         Height          =   1005
         Left            =   195
         TabIndex        =   32
         Top             =   5600
         Width           =   10185
         Begin VB.OptionButton optMduYes 
            Caption         =   "&A.要產生"
            Height          =   255
            Left            =   450
            TabIndex        =   27
            Top             =   690
            Value           =   -1  'True
            Width           =   1065
         End
         Begin VB.OptionButton optMduNo 
            Caption         =   "&B.不產生"
            Height          =   285
            Left            =   1830
            TabIndex        =   28
            Top             =   660
            Width           =   1125
         End
         Begin CS_Multi.CSmulti gmdMduid 
            Height          =   405
            Left            =   150
            TabIndex        =   26
            Top             =   240
            Width           =   9855
            _ExtentX        =   17383
            _ExtentY        =   714
            ButtonCaption   =   "大 樓 名 稱"
         End
      End
      Begin VB.CheckBox chkOther 
         Caption         =   "包含一般戶"
         Height          =   225
         Left            =   5565
         TabIndex        =   10
         Top             =   1155
         Value           =   1  '核取
         Width           =   2145
      End
      Begin Gi_Multi.GiMulti gmdBillType 
         Height          =   360
         Left            =   165
         TabIndex        =   25
         Top             =   5250
         Width           =   10200
         _ExtentX        =   17992
         _ExtentY        =   635
         ButtonCaption   =   "單 據 類 別"
      End
      Begin Gi_Multi.GiMulti gmdClctEn 
         Height          =   360
         Left            =   165
         TabIndex        =   20
         Top             =   3500
         Width           =   10200
         _ExtentX        =   17992
         _ExtentY        =   635
         ButtonCaption   =   "收    費    員"
      End
      Begin Gi_Multi.GiMulti gmdCustClass 
         Height          =   360
         Left            =   165
         TabIndex        =   19
         Top             =   3150
         Width           =   10200
         _ExtentX        =   17992
         _ExtentY        =   635
         ButtonCaption   =   "客 戶 類 別"
      End
      Begin Gi_Multi.GiMulti gmdServCode 
         Height          =   360
         Left            =   165
         TabIndex        =   17
         Top             =   2490
         Width           =   10200
         _ExtentX        =   17992
         _ExtentY        =   635
         ButtonCaption   =   "服    務    區"
      End
      Begin Gi_Multi.GiMulti gmdAreaCode 
         Height          =   360
         Left            =   165
         TabIndex        =   16
         Top             =   2150
         Width           =   10200
         _ExtentX        =   17992
         _ExtentY        =   635
         ButtonCaption   =   "行    政    區"
      End
      Begin Gi_Multi.GiMulti gmdCMCode 
         Height          =   360
         Left            =   165
         TabIndex        =   15
         Top             =   1820
         Width           =   10200
         _ExtentX        =   17992
         _ExtentY        =   635
         ButtonCaption   =   "收 費 方 式"
      End
      Begin Gi_Multi.GiMulti gimCustStatus 
         Height          =   360
         Left            =   165
         TabIndex        =   18
         Top             =   2800
         Width           =   10200
         _ExtentX        =   17992
         _ExtentY        =   635
         ButtonCaption   =   "客  戶 狀 態"
      End
      Begin Gi_Multi.GiMulti gimCreateEn 
         Height          =   360
         Left            =   165
         TabIndex        =   22
         Top             =   4180
         Width           =   10200
         _ExtentX        =   17992
         _ExtentY        =   635
         ButtonCaption   =   "產 生 人 員"
      End
      Begin prjGiList.GiList gilServiceType 
         Height          =   330
         Left            =   6675
         TabIndex        =   9
         Top             =   -30
         Visible         =   0   'False
         Width           =   2940
         _ExtentX        =   5186
         _ExtentY        =   582
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         F2Corresponding =   -1  'True
      End
      Begin Gi_Date.GiDate gdaShouldDate2 
         Height          =   360
         Left            =   7140
         TabIndex        =   6
         Top             =   750
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   635
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
      Begin Gi_Date.GiDate gdaShouldDate1 
         Height          =   360
         Left            =   5550
         TabIndex        =   5
         Top             =   750
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   635
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
      Begin Gi_Multi.GiMulti gmBank 
         Height          =   360
         Left            =   165
         TabIndex        =   14
         Top             =   1485
         Width           =   10200
         _ExtentX        =   17992
         _ExtentY        =   635
         ButtonCaption   =   "銀 行 名 稱"
      End
      Begin Gi_Multi.GiMulti gimUCCode 
         Height          =   360
         Left            =   165
         TabIndex        =   23
         Top             =   4550
         Width           =   10200
         _ExtentX        =   17992
         _ExtentY        =   635
         ButtonCaption   =   "未  收  原  因"
      End
      Begin CS_Multi.CSmulti gimCitemCode 
         Height          =   360
         Left            =   165
         TabIndex        =   24
         Top             =   4900
         Width           =   10200
         _ExtentX        =   17992
         _ExtentY        =   635
         ButtonCaption   =   "收  費  項  目"
      End
      Begin prjGiList.GiList gilCompCode 
         Height          =   330
         Left            =   1275
         TabIndex        =   0
         Top             =   360
         Width           =   2940
         _ExtentX        =   5186
         _ExtentY        =   582
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         F2Corresponding =   -1  'True
      End
      Begin Gi_Date.GiDate gdaCreateTime2 
         Height          =   330
         Left            =   4500
         TabIndex        =   4
         Top             =   -60
         Visible         =   0   'False
         Width           =   1125
         _ExtentX        =   1984
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
      Begin Gi_Date.GiDate gdaCreateTime1 
         Height          =   330
         Left            =   3180
         TabIndex        =   3
         Top             =   -60
         Visible         =   0   'False
         Width           =   1125
         _ExtentX        =   1984
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
      Begin Gi_Multi.GiMulti gmdOldClctEn 
         Height          =   360
         Left            =   160
         TabIndex        =   21
         Top             =   3840
         Width           =   10200
         _ExtentX        =   17992
         _ExtentY        =   635
         ButtonCaption   =   "原 收 費  員"
      End
      Begin Gi_Time.GiTime gitCreateTime1 
         Height          =   345
         Left            =   1290
         TabIndex        =   7
         Top             =   1110
         Width           =   1935
         _ExtentX        =   3413
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
      Begin Gi_Time.GiTime gitCreateTime2 
         Height          =   345
         Left            =   3630
         TabIndex        =   8
         Top             =   1110
         Width           =   1935
         _ExtentX        =   3413
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
      Begin VB.Label lblBillHeadFmt 
         AutoSize        =   -1  'True
         Caption         =   "多帳戶產生依據設定"
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   420
         TabIndex        =   41
         Top             =   810
         Width           =   1620
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "承辦銀行"
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
         Left            =   4740
         TabIndex        =   40
         Top             =   450
         Width           =   780
      End
      Begin VB.Label lblCreateTime 
         AutoSize        =   -1  'True
         Caption         =   "產生時間"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   390
         TabIndex        =   39
         Top             =   1170
         Width           =   780
      End
      Begin VB.Label lblunti2 
         AutoSize        =   -1  'True
         Caption         =   "至"
         Height          =   195
         Left            =   3330
         TabIndex        =   38
         Top             =   1200
         Width           =   195
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
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   615
         TabIndex        =   37
         Top             =   465
         Width           =   585
      End
      Begin VB.Label lbluntil 
         AutoSize        =   -1  'True
         Caption         =   "至"
         Height          =   210
         Left            =   6840
         TabIndex        =   35
         Top             =   855
         Width           =   195
      End
      Begin VB.Label lblReadAmt 
         AutoSize        =   -1  'True
         Caption         =   "應收日期範圍"
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
         Height          =   210
         Left            =   4350
         TabIndex        =   34
         Top             =   840
         Width           =   1170
      End
      Begin VB.Label lblServiceType 
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
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   5820
         TabIndex        =   33
         Top             =   45
         Visible         =   0   'False
         Width           =   780
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "F2.確定"
      Default         =   -1  'True
      Height          =   405
      Left            =   600
      TabIndex        =   29
      Top             =   6780
      Width           =   1275
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&X)"
      Height          =   405
      Left            =   9090
      TabIndex        =   30
      Top             =   6780
      Width           =   1275
   End
End
Attribute VB_Name = "frmSO3272A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private objAgency As Object
'Private objAgency As New clsInterface
Private objAction As Object
Private rsBankData As New ADODB.Recordset
Private rsDefVal As New ADODB.Recordset
Private strChoose As String
Private blnUnload As Boolean
Private strINIpath1 As String
Private strErrPath As String
Private strChooseString As String
Dim pChooseMultiAcc As String
Private intPara24 As Integer
Private blnChina As Boolean  ''  2003/10/07   判斷是否為新視波的中國商銀
Private rsBillHeadFmt As New ADODB.Recordset
'#3531 多新增一個中國信託(新)流程 By Kin 2007/10/31
Public Enum CreditCardType     ''  2003/10/27  判斷所用的是那一定的銀行
      CardType_default = 0  '' 預設為 聯合與中華  東森用的
      CardType_chinatrust = 1   '' 中國信託 目前是 新視波
      CardType_hsbc = 2  '' 匯豐銀行信用卡
      CardType_newChina = 3
End Enum
Public e_CreditCardType As CreditCardType

Private Sub cboBillHeadFmt_Click()
    On Error Resume Next
    Dim strCitemCode As String
    Dim rsCD019  As New ADODB.Recordset
        If rsBillHeadFmt.RecordCount = 0 Then Exit Sub
        rsBillHeadFmt.AbsolutePosition = cboBillHeadFmt.ListIndex + 1
        If Not GetRS(rsCD019, "Select CodeNo From " & GetOwner & "CD019 Where CodeNo In (" & rsBillHeadFmt("CitemCodeStr") & ") And StopFlag=0") Then Exit Sub
        If rsCD019.EOF Then Exit Sub
        strCitemCode = rsCD019.GetString(, , , ",")
        strCitemCode = Left(strCitemCode, Len(strCitemCode) - 1)
        gimCitemCode.Filter = gimCitemCode.Filter
        gimCitemCode.SetQueryCode strCitemCode

End Sub

Private Sub cboCreditCardType_Click()
  Select Case cboCreditCardType.ItemData(cboCreditCardType.ListIndex)
       Case 0
           e_CreditCardType = CardType_default
           chkChinaBank.Visible = True
           chkChinaBank.Value = 0
           fraTCB.Visible = True
           Call SetgiMulti(gmBank, "CodeNo", "Description", "CD018", "代碼", "名稱", _
                                   " Where UPPER(PrgName) = 'CREDITCARDUNION'  AND  COMPCODE =" & _
                                   gilCompCode.GetCodeNo & " AND " & "(BANKID <> 'TCB' OR BANKID IS NULL)")
       Case 1
           e_CreditCardType = CardType_chinatrust
           chkChinaBank.Visible = False
           chkChinaBank.Value = 0
           fraTCB.Visible = False
            Call SetgiMulti(gmBank, "CodeNo", "Description", "CD018", "代碼", "名稱", _
                                                  " Where UPPER(PrgName) = '" & "CREDITCARDCHINATRUST" & "' AND  COMPCODE =" & gilCompCode.GetCodeNo)
        Case 2
            e_CreditCardType = CardType_hsbc
            chkChinaBank.Visible = False
            chkChinaBank.Value = 0
            fraTCB.Visible = False
            Call SetgiMulti(gmBank, "CodeNo", "Description", "CD018", "代碼", "名稱", _
                                                  " Where UPPER(PrgName) = '" & "CREDITCARDHSBC" & "' AND  COMPCODE =" & gilCompCode.GetCodeNo)
        '#3531 新增中國信託(新)格式
        Case 3
            e_CreditCardType = CardType_newChina
            chkChinaBank.Visible = False
            chkChinaBank.Value = 0
            fraTCB.Visible = False
            Call SetgiMulti(gmBank, "CodeNo", "Description", "CD018", "代碼", "名稱", _
                                                  " Where UPPER(PrgName) = '" & "CREDITCARDTRUSTNEW" & "' AND  COMPCODE =" & gilCompCode.GetCodeNo)

            
    End Select
End Sub

Private Sub chkChinaBank_Click()
If chkChinaBank.Value = 1 Then
   gmBank.Clear
   gmBank.Enabled = 0
   fraTCB.Enabled = True
Else
   gmBank.Enabled = 1
   fraTCB.Enabled = False
End If
End Sub

Private Sub cmdCancel_Click()
  On Error GoTo ChkErr
    Unload Me
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "GetBankData")
End Sub

Private Function GetBankData() As Boolean
On Error GoTo ChkErr
Dim strWhere As String
 
   Dim strSQL  As String
   
   If Len(gmBank.GetQryStr) <> 0 Then
      strWhere = " AND CodeNo " & gmBank.GetQryStr
   End If
   
'   If chkChinaBank.Value = 1 Then
'       strWhere = " AND BANKID = 'TCB' "
'   End If

   strSQL = "Select CodeNO,Description,PrgName,ActNo,BankId,CorpID,Spcno From " & GetOwner & "CD018 where " & " UPPER(SUBSTR(PrgName,1,10)) = '" & "CREDITCARD" & "'" & strWhere
   Set rsBankData = gcnGi.Execute(strSQL)
   If rsBankData.EOF Or rsBankData.BOF Then MsgBox "銀行資料有誤！請檢查銀行代碼檔！", vbExclamation, "訊息！": Exit Function
   
   GetBankData = True
Exit Function
ChkErr:
    Call ErrSub(Me.Name, "GetBankData")
End Function

Private Sub cmdOK_Click()
  On Error GoTo ChkErr
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
      
    If Not ChkDTok Then Exit Sub
    If Not IsDataOk Then Exit Sub
    blnUnload = False
    If Not subChoose Then GoTo lExit
    Set objAgency = CreateObject("CreditCardOutEmc.clsInterface")
    If gmdCustClass.GetQryStr <> "" Then
        strSQL = " From " & GetOwner & "SO033 A," & GetOwner & "SO001 B Where " & strChoose
    Else
        strSQL = " From " & GetOwner & "SO033 A Where " & strChoose
    End If
       
      '' 2003/10/08  如果是中國商銀暫時不管他的 PTCODE
       
    If e_CreditCardType = CardType_default Then
        If chkChinaBank.Value = 1 Then
            If optTCB.Value = True Then
                 strSQL = strSQL & "  AND A.PTCode = 3 "
            Else
                 strSQL = strSQL & "  AND A.PTCode = 5"
            End If
        Else
            strSQL = strSQL & "  AND (A.PTCode <>  3 AND A.PTCODE <> 5  OR A.PTCode IS  NULL) "
        End If
    End If
      
    If Not GetBankData Then Exit Sub
    With objAgency
          .errPath = strErrPath
          .iniPath = strINIpath1
          .ChooseStr = strChoose
        If Len(gmBank.GetQryStr) <> 0 Then
                 '' 以下改為以 SO033 的BANKCODE 作判斷 而不是 SO002 .BankId = " C.BANKCODE " & gmBank.GetQryStr
                 '' 2003/10/08  如果是中國商銀暫時不管他的 PTCODE
            Select Case e_CreditCardType
                Case 0
                    .BankId = " A.BANKCODE " & gmBank.GetQryStr & " AND (A.PTCODE NOT IN (3,5)   OR A.PTCode IS  NULL) "
                Case 1
                    .BankId = " A.BANKCODE " & gmBank.GetQryStr
                Case 2
                    .BankId = " A.BANKCODE " & gmBank.GetQryStr
                '#3531 增加中國信託新格式 By Kin 2007/11/02
                Case 3
                    .BankId = " A.BANKCODE " & gmBank.GetQryStr
                    .Bankstr = gmBank.GetQryStr
            End Select
        Else
            If chkChinaBank.Value = 1 Then
                If optTCB.Value = True Then
                    .BankId = " A.PTCode = 3  "
                Else
                    .BankId = " A.PTCode = 5  "
                End If
            Else
                Select Case e_CreditCardType
                    Case 0
                        .BankId = " A.BANKCODE IN (" & gmBank.GetQueryCode & ")  AND (A.PTCODE NOT IN (3,5)  OR A.PTCODE IS NULL) "
                    Case Else
                        .BankId = " A.BANKCODE IN(" & gmBank.GetQueryCode & ")"
                End Select
            End If
        End If
       ''  2003/10/08  若是中國商銀則  e_CreditCardType = CardType_chinatrust
        Select Case e_CreditCardType
            Case 0
                If chkChinaBank.Value = 0 Then
                    .BankName = Me.Caption & "  聯合信用卡 "
                    .blnIsTCB = False
                    .blnIsTCBCard = False
                Else
                    If optTCB.Value = True Then
                        .BankName = Me.Caption & "  中華商銀自行卡  "
                        .blnIsTCB = True
                        .blnIsTCBCard = True
                    Else
                        .BankName = Me.Caption & "  中華商銀他行卡 "
                        .blnIsTCB = True
                        .blnIsTCBCard = False
                    End If
                 End If
            Case 1
                .BankName = Me.Caption & "  中國信託信用卡 "
                .blnIsTCB = False
                .blnIsTCBCard = False
            Case 2
                .BankName = Me.Caption & "  匯豐銀行信用卡 "
            '#3531 增加中國信託新格式
            Case 3
                .BankName = Me.Caption & "  中國信託信用卡 "
                .blnIsTCB = False
                .blnIsTCBCard = False
        End Select
        .pTableOwnerName = GetOwner
        .PrgName = rsBankData("PrgName") & ""
        .uSpcNO = rsBankData("SpcNO") & ""
        Set .Connection = gcnGi
        If Len(rsBankData("PrgName") & "") = 0 Then
            MsgBox "設定程式名稱無效或無使用權限！！", vbExclamation, "提示"
        Else
            '' 2003/10/29  由於目前改成不是直接由CD018 的PRGNAME作元件名稱，所以只有直接命名
            '' Set objAction = .InitObject(rsBankData("Prgname") & "")
             Set objAction = .InitObject("CREDITCARDUNION")
             Call objAction.Action(.blnIsTCB, pChooseMultiAcc, .blnIsTCBCard, CLng(e_CreditCardType))
             Set objAction = Nothing
        End If
    
    End With
    '判斷是否有資料
      If objAgency.unodata Then
          Set objAgency = Nothing
          Screen.MousePointer = vbDefault
          blnUnload = True
          On Error Resume Next
          gilCompCode.SetFocus
          Exit Sub
      Else
          '判斷是否執行成功
          If objAgency.uupdate Then
              Call subInsert
              blnUnload = True
          End If
      End If
lExit:
      blnUnload = True
      Set objAgency = Nothing
      DoEvents
      Screen.MousePointer = vbDefault
     
  Exit Sub
ChkErr:
    blnUnload = True
    If Err.Number = 20 Then Resume Next
    If Err.Number = 91 Then
       MsgBox "轉帳程式名稱錯誤或者該銀行沒有轉帳資料產生功能!", vbExclamation, "警告..."
       Exit Sub
    End If
    Call ErrSub(Me.Name, "cmdOk_Click")
End Sub

Private Function subChoose() As Boolean
  On Error GoTo ChkErr
  Dim strAddr As String

   subChoose = False
   strChoose = ""
    If gilCompCode.GetCodeNo <> "" Then subAnd ("A.CompCode=" & gilCompCode.GetCodeNo)
    '#3728
    'If gilServiceType.GetCodeNo <> "" Then subAnd ("A.ServiceType='" & gilServiceType.GetCodeNo & "'")
    If gdaShouldDate1.GetValue <> "" Then subAnd ("A.ShouldDate >= To_Date(" & gdaShouldDate1.GetValue & ",'YYYYMMDD')")
    If gdaShouldDate2.GetValue <> "" Then subAnd ("A.ShouldDate < To_Date(" & gdaShouldDate2.GetValue & ",'YYYYMMDD') +1")
    '*************************************************************************************************************************
    '#2672 產生時間改為到時分秒 By Kin 2008/05/12
'    If gdaCreateTime1.GetValue <> "" Then subAnd ("A.CreateTime >= To_Date(" & gdaCreateTime1.GetValue & ",'YYYYMMDD')")
'    If gdaCreateTime2.GetValue <> "" Then subAnd ("A.CreateTime < To_Date(" & gdaCreateTime2.GetValue & ",'YYYYMMDD') +1")
    If gitCreateTime1.GetValue <> "" Then subAnd ("A.CreateTime>=To_Date(" & gitCreateTime1.GetValue & ",'YYYYMMDDHH24MISS')")
    If gitCreateTime2.GetValue <> "" Then subAnd ("A.CreateTime<To_Date(" & gitCreateTime2.GetValue & ",'YYYYMMDDHH24MISS')+1")
    '*************************************************************************************************************************
    If gmdCMCode.GetQryStr <> "" Then subAnd ("A.CMCode " & gmdCMCode.GetQryStr)
    If gmdAreaCode.GetQryStr <> "" Then subAnd ("A.AreaCode " & gmdAreaCode.GetQryStr)
    If gmdServCode.GetQryStr <> "" Then subAnd ("A.ServCode " & gmdServCode.GetQryStr)
    '#1659  950517原"收費人員"改為"原收費人員"，另新增一"收費人員"
    If gmdClctEn.GetQryStr <> "" Then subAnd ("A.ClctEn " & gmdClctEn.GetQryStr)
    If gmdOldClctEn.GetQryStr <> "" Then subAnd ("A.OldClctEn " & gmdOldClctEn.GetQryStr)

    If gmdBillType.GetQryStr <> "" Then
       subAnd ("SubStr(A.BillNo,7,1) In (" & gmdBillType.GetQueryCode & ")")
    Else
       subAnd ("SubStr(A.BillNo,7,1) In ('B','T','I','M','P')")
       
    End If
    If gimCreateEn.GetQryStr <> "" Then subAnd ("A.CreateEn " & gimCreateEn.GetQryStr)
    
    '產生大樓
    If optMduYes Then
       '產生一般戶
       If chkOther.Value = 1 Then
          If gmdMduid.GetQryStr <> "" Then subAnd ("(A.MduId " & gmdMduid.GetQryStr & " Or A.MduId Is Null) ")
       Else
       '不產生一般戶
          If gmdMduid.GetQryStr <> "" Then
             subAnd ("A.MduId " & gmdMduid.GetQryStr)
          Else
             subAnd ("A.MduId Is Not Null")
          End If
       End If
    '不產生大樓
    Else
       '產生一般戶
       If chkOther.Value = 1 Then
          If gmdMduid.GetQryStr <> "" Then
             subAnd ("(Not A.MduId " & gmdMduid.GetQryStr & " Or A.MduId Is Null) ")
          Else
             subAnd ("A.MduId Is Null")
          End If
       Else
       '不產生一般戶
          If gmdMduid.GetQryStr <> "" Then
             subAnd ("Not A.MduId " & gmdMduid.GetQryStr)
          Else
             MsgBox "此條件查無資料!!", vbExclamation, "訊息"
             Exit Function
          End If
       End If
    End If
    
    If gimUCCode.GetQryStr <> "" Then subAnd ("A.UCCODE " & gimUCCode.GetQryStr)
    If gimCitemCode.GetQryStr <> "" Then subAnd ("A.CitemCode " & gimCitemCode.GetQryStr)
    
    
    pChooseMultiAcc = strChoose   '' 到目前為止為 SO033 的條件值
    
    If gimCustStatus.GetQryStr <> "" Then subAnd ("C.CustStatusCode " & gimCustStatus.GetQryStr)
    If gmdCustClass.GetQryStr <> "" Then subAnd ("B.ClassCode1 " & gmdCustClass.GetQryStr)
    
'    If optChargeAddr.Value Then
'       strAddr = "收費地址"
'    Else
'       strAddr = "裝機地址"
'    End If
    subChoose = True
    strChooseString = "公司別　    :" & subSpace(gilCompCode.GetDescription) & ";" & _
                      "服務類別    :" & subSpace(gilServiceType.GetDescription) & ";" & _
                      "承辦銀行    :" & subSpace(gmBank.GetDispStr) & ";" & _
                      "收費日期　　:" & subSpace(gdaShouldDate1.GetValue(True)) & "~" & subSpace(gdaShouldDate2.GetValue(True)) & ";" & _
                      "收費方式    :" & subSpace(gmdCMCode.GetDispStr) & ";" & _
                      "行政區      :" & subSpace(gmdAreaCode.GetDispStr) & ";" & _
                      "服務區      :" & subSpace(gmdServCode.GetDispStr) & ";" & _
                      "客戶狀態    :" & subSpace(gimCustStatus.GetDispStr) & ";" & _
                      "客戶類別　　:" & subSpace(gmdCustClass.GetDispStr) & ";" & _
                      "收費員      :" & subSpace(gmdClctEn.GetDispStr) & ";" & _
                      "原收費員    :" & subSpace(gmdOldClctEn.GetDispStr) & ";" & _
                      "單據類別    :" & subSpace(gmdBillType.GetDispStr) & ";" & _
                      "產生人員    :" & subSpace(gimCreateEn.GetDispStr) & ";" & _
                      "大樓名稱    :" & subSpace(gmdMduid.GetDispStr)
  
  Exit Function
ChkErr:
    Call ErrSub(Me.Name, "subChoose")
End Function

Private Sub subAnd(strCH As String)
  On Error GoTo ChkErr
    '是空白加And
    If strChoose <> "" Then
       strChoose = strChoose & " And " & strCH
    Else
       strChoose = strCH
    End If
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "subAnd")
End Sub

Private Function AddChar(ByVal vStr As String, ByVal vChar As String, Optional OverWrite As Boolean = True) As String
On Error GoTo ChkErr
    Static strChar As String
        If OverWrite Then strChar = ""
        If InStr(1, vStr, vChar) > 0 Then
            strChar = strChar & IIf(strChar <> "", ",'", "'") & vChar & "'"
        End If
        AddChar = strChar
Exit Function
ChkErr:
    Call ErrSub(Me.Name, "ADDChar")
End Function
Private Sub Form_Activate()
On Error GoTo ChkErr
  Dim strSQL As String
  Dim rsTmp As New ADODB.Recordset
  
    strSQL = "Select PrgName From " & GetOwner & "CD018 Where UPPER(SUBSTR(PrgName,1,10)) = '" & "CREDITCARD" & "'"
    If rsTmp.State = adStateOpen Then rsTmp.Close
'    rsTmp.CursorLocation = adUseClient
'    rsTmp.Open strSQL, gcnGi, adOpenForwardOnly, adLockReadOnly
    If Not GetRS(rsTmp, strSQL, gcnGi, adUseClient, adOpenForwardOnly, adLockReadOnly) Then Exit Sub
    If rsTmp.EOF Then
       MsgBox "請設定銀行代碼中轉帳程式名稱!", vbExclamation, "警告"
       Unload Me
    End If
    rsTmp.Close
    Set rsTmp = Nothing
    Screen.MousePointer = vbDefault
  Exit Sub
ChkErr:
  ErrSub Me.Name, "Form_Activate"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If Not ChkGiList(KeyCode) Then Exit Sub
    If KeyCode = vbKeyF2 Then cmdOK.Value = True: KeyCode = 0
    If KeyCode = vbKeyEscape Then Unload Me: KeyCode = 0
End Sub

Private Sub Form_Load()
    On Error GoTo ChkErr
    
    '' 2003/10/07 這一段暫時作測試用 設為true，以後看如果判斷是新視波所要中國商銀
        gitCreateTime1.ShowSecond
        gitCreateTime2.ShowSecond
        If Alfa2 Then
            Call GetGlobal
        End If
        blnUnload = True
        strINIpath1 = GetIniPath1
        strErrPath = ReadGICMIS1("ErrLogPath")
        Call subGil
        Call subGim
        Call getDefault
        chkChinaBank.Enabled = GetUserPriv("SO3272", "SO32721")
        '*************************************************************************************************
        '#3728 強迫格式以多媒体出帳方式 By Kin 2008/03/25
        'intPara24 = GetSystemParaItem("Para24", gilCompCode.GetCodeNo, gilServiceType.GetCodeNo, "SO043")
        cboBillHeadFmt.Enabled = True
        'gilServiceType.Enabled = False
        'lblServiceType.ForeColor = &H808080
        CboAddItem
        '*************************************************************************************************

   Exit Sub
ChkErr:
   Call ErrSub(Me.Name, "Form_Load")
End Sub

Private Sub subGil()
  On Error GoTo ChkErr
  
        SetgiList gilCompCode, "CodeNo", "Description", "CD039", , , , , 3, 12
        SetgiList gilServiceType, "CodeNo", "Description", "CD046", , , , , 1, 12
        
        
        'SetgiList gilBank, "CodeNo", "Description", "CD018", , , , , 3, 24, " Where UPPER(SUBSTR(PrgName,1,10)) = '" & "CREDITCARD" & "'"
   Exit Sub
ChkErr:
   ErrSub Me.Name, "subGil"
End Sub

Private Sub subGim()

Dim strWhereBank As String

  On Error GoTo ChkErr
  
  ''2003/10/29  以下不用作 gmBank 的設定
  
'    Select Case e_CreditCardType
'        Case 0  '' 預設
'             strWhereBank = "CREDITCARD"
'        Case 1  '' 中國信託
'             strWhereBank = "CREDITCARDChiatrust"
'     End Select
'    Call SetgiMulti(gmBank, "CodeNo", "Description", "CD018", "代碼", "名稱", " Where UPPER(SUBSTR(PrgName,1,10)) = '" & "CREDITCARD" & "'")
    
    Call SetgiMulti(gmdCMCode, "CodeNo", "Description", "CD031", "代碼", "名稱")
    Call SetgiMulti(gmdAreaCode, "CodeNo", "Description", "CD001", "代碼", "名稱")
    Call SetgiMulti(gmdServCode, "CodeNo", "Description", "CD002", "代碼", "名稱")
    Call SetgiMulti(gimCustStatus, "CodeNo", "Description", "CD035", "代碼", "名稱")
    Call SetgiMulti(gmdCustClass, "CodeNo", "Description", "CD004", "代碼", "名稱")
    Call SetgiMulti(gmdClctEn, "EmpNo", "EmpName", "CM003", "代碼", "名稱")
    Call SetgiMulti(gmdOldClctEn, "EmpNo", "EmpName", "CM003", "代碼", "名稱")
    Call SetgiMultiAddItem(gmdBillType, "B,T,I,M,P", "收費單,臨時收費單,裝機單,維修單,停拆移機單", "代碼", "名稱")
   
    Call SetgiMulti(gimUCCode, "CodeNo", "Description", "CD013", "代碼", "名稱")   '' 未收原因
    Call SetgiMulti(gimCitemCode, "CodeNo", "Description", "CD019", "代碼", "名稱")  ' '收費項目

   With gmdBillType
          .SetDispStr "收費單,臨時收費單"
          .SetQueryCode "B,T"
'          .SetDispStr "收費單,臨時收費單,裝機單,維修單,停拆移機單"
'          .SetQueryCode "B,T,I,M,P"
    End With
    
    Call SetgiMulti(gimCreateEn, "EmpNo", "EmpName", "CM003", "代碼", "名稱")
    Call SetgiMulti(gmdMduid, "Mduid", "Name", "SO017", "代碼", "名稱")
    gimCustStatus.SetQueryCode ("1")
   ''   gmdBillType.SelectAll
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "subGim")
End Sub

Private Sub getDefault()
  On Error GoTo ChkErr
     gilCompCode.SetCodeNo garyGi(9)
     gilCompCode.Query_Description
  Exit Sub
ChkErr:
  ErrSub Me.Name, "getDefault"
End Sub

Private Function IsDataOk() As Boolean
On Error GoTo ChkErr
Dim strErrMsg As String


    IsDataOk = False
        If Len(gilCompCode.GetCodeNo & "") = 0 Then strErrMsg = "公司別": gilCompCode.GetCodeNo: GoTo Warning
        '#3728 強迫多媒体出帳,所以取消服務別選項 By Kin 2008/03/25
        'If Len(gilServiceType.GetCodeNo & "") = 0 Then strErrMsg = "服務類別": gilServiceType.GetCodeNo: GoTo Warning
     ''   If gilBank.GetDescription = "" Then strErrMsg = "承辦銀行": gilBank.SetFocus: GoTo Warning
        If chkChinaBank.Value = 0 Then
             If gmBank.GetQueryCode = "" Then strErrMsg = "承辦銀行": gmBank.SetFocus: GoTo Warning
        End If
        If gdaShouldDate1.GetValue = "" Then strErrMsg = "收費日期起始日": gdaShouldDate1.SetFocus: GoTo Warning
        If gdaShouldDate2.GetValue = "" Then strErrMsg = "收費日期截止日": gdaShouldDate2.SetFocus: GoTo Warning
        If DateDiff("d", gdaShouldDate1.Text, gdaShouldDate2.Text) < 0 Then MsgBox "收費截止日不得小於收費起始日！", vbExclamation, "訊息！": Exit Function
        '#3728 強迫以多媒体出帳 By Kin 2008/03/25
        If cboBillHeadFmt.Text = "" Then strErrMsg = "多帳戶產生依據設定": cboBillHeadFmt.SetFocus: GoTo Warning
        
    IsDataOk = True
    Exit Function
Warning:
    Call MsgMustBe(strErrMsg)
    Exit Function
ChkErr:
    Call ErrSub(Me.Name, "IsDataOk")
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo ChkErr
    If blnUnload Then
       Set objAgency = Nothing
    Else
       Cancel = True
    End If
    Call CloseRecordset(rsBankData)
    Call CloseRecordset(rsBillHeadFmt)
    Screen.MousePointer = vbDefault
    
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "Form_QueryUnload")
End Sub

Private Sub Frame4_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub gdaCreateTime1_GotFocus()
  On Error Resume Next
    If gdaCreateTime1.GetValue <> "" Then Exit Sub
    gdaCreateTime1.SetValue Date
End Sub

Private Sub gdaCreateTime1_Validate(Cancel As Boolean)
On Error GoTo ChkErr
   If Not IsDate(gdaCreateTime1.Text) Then Exit Sub
   If gdaCreateTime1.GetValue <> "" Then
      If gdaCreateTime2.GetValue = "" Then gdaCreateTime2.SetValue GetLastDate(gdaCreateTime1.GetValue(True))
   End If
Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "gdaCreateTime1_Validate")
End Sub

Private Sub gdaCreateTime2_GotFocus()
  On Error Resume Next
    If gdaCreateTime2.GetValue <> "" Then Exit Sub
    gdaCreateTime2.SetValue GetLastDate(Date)
    
End Sub

Private Sub gdaCreateTime2_Validate(Cancel As Boolean)
   On Error Resume Next
     If gdaCreateTime1.GetValue = "" Or gdaCreateTime2.GetValue = "" Then Exit Sub
     If Not IsDate(gdaCreateTime2.Text) Then Exit Sub
     If DateDiff("d", gdaCreateTime1.GetValue(True), gdaCreateTime2.GetValue(True)) < 0 Then
        MsgBox "產生時間截止日不可小於產生時間起始日!", vbExclamation, "警告"
        gdaCreateTime2.SetValue gdaCreateTime1.GetValue
        Cancel = True
     End If
End Sub

Private Sub gdaShouldDate1_GotFocus()
  On Error Resume Next
    If gdaShouldDate1.GetValue <> "" Then Exit Sub
    gdaShouldDate1.SetValue Date
End Sub

Private Sub gdaShouldDate1_Validate(Cancel As Boolean)
On Error GoTo ChkErr
     
   If Not IsDate(gdaShouldDate1.Text) Then Exit Sub
   If gdaShouldDate1.GetValue <> "" Then
      If gdaShouldDate2.GetValue = "" Then gdaShouldDate2.SetValue GetLastDate(gdaShouldDate1.GetValue(True))
   End If
Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "gdaShouldDate1_Validate")
End Sub

Private Sub gdaShouldDate2_GotFocus()
  On Error Resume Next
    If gdaShouldDate2.GetValue <> "" Then Exit Sub
    gdaShouldDate2.SetValue GetLastDate(Date)
End Sub

Private Sub gdaShouldDate2_Validate(Cancel As Boolean)
   On Error Resume Next
     If gdaShouldDate1.GetValue = "" Or gdaShouldDate2.GetValue = "" Then Exit Sub
     If Not IsDate(gdaShouldDate2.Text) Then Exit Sub
     If DateDiff("d", gdaShouldDate1.GetValue(True), gdaShouldDate2.GetValue(True)) < 0 Then
        MsgBox "應收截止日不可小於應收起始日!", vbExclamation, "警告"
        gdaShouldDate2.SetValue gdaShouldDate1.GetValue
        Cancel = True
     End If
End Sub

Private Sub subInsert()
On Error GoTo ChkErr
  Dim strSQL As String
  
    strSQL = "INSERT INTO " & GetOwner & "SO097 (EMPNO,PRGNAME,RUNDATETIME,SECOND,SELECTION) " & _
                   "VALUES ('" & garyGi(0) & "','SO3273A',sysdate,'" & _
                   objAgency.utime & "秒','" & Replace(strChooseString, "'", "") & "')"
    gcnGi.Execute (strSQL)
   Exit Sub
ChkErr:
   ErrSub Me.Name, "subInsert"
End Sub

Private Sub gilCompCode_Change()
    On Error Resume Next
    
If ChgComp(gilCompCode, "SO3270", "SO3272") = False Then Exit Sub
    
    Call subGil
    Call subGim
    Call SelectServType(gilCompCode.GetCodeNo, gilServiceType)
'    gilServiceType.ListIndex = 1
    
    Call GiMultiFilter(gmBank, , gilCompCode.GetCodeNo)
''    gmBank.Filter = gmBank.Filter & IIf(gmBank.Filter = "", " Where ", " And ") & " UPPER(SUBSTR(PrgName,1,10)) = '" & "CREDITCARD" & "'  AND (BANKID  <>   '" & "TCB" & "'   OR BANKID IS NULL) "

    Call GiMultiFilter(gmdAreaCode, , gilCompCode.GetCodeNo)
    Call GiMultiFilter(gmdMduid, , gilCompCode.GetCodeNo)
    Call GiMultiFilter(gmdServCode, , gilCompCode.GetCodeNo)
    
   If SetCreditCardTypeCbo = False Then cmdOK.Enabled = False

End Sub

Private Sub gilServiceType_Change()
    On Error Resume Next
        Call GiMultiFilter(gmdCMCode, gilServiceType.GetCodeNo)
        Call GiMultiFilter(gmdCustClass, gilServiceType.GetCodeNo)
End Sub

Private Function SetCreditCardTypeCbo() As Boolean

  Dim strSQL As String
  Dim rsTmp As New ADODB.Recordset
  Dim lbn
  Dim lngIndex As Long
  On Error GoTo ChkErr
  
  For lngIndex = 0 To cboCreditCardType.ListCount - 1
       cboCreditCardType.RemoveItem (0)
  Next
    lngIndex = 0
    strSQL = "SELECT DISTINCT PrgName FROM " & GetOwner & "CD018 WHERE UPPER(SUBSTR(PrgName,1,10)) = '" & "CREDITCARD" & "'  AND COMPCODE =" & gilCompCode.GetCodeNo & "  AND STOPFLAG = 0 "
    Call GetRS(rsTmp, strSQL, gcnGi, adUseClient, adOpenDynamic)
    Dim blnAdd As Boolean   ''  判斷如果有加入信用卡項目清單的話，則將其設為flase ，這樣底下的lngIndex才可以加 1  不然下筆資料的迴圈會出現 索引超出陣列範圍
    blnAdd = False
  
    Do While Not (rsTmp.EOF Or rsTmp.BOF)
        Select Case UCase(rsTmp("PrgName"))
            Case "CREDITCARDUNION"
                 cboCreditCardType.AddItem (" 聯合/中華商銀 信用卡")
                 cboCreditCardType.ItemData(lngIndex) = 0
                 blnAdd = True
            Case "CREDITCARDCHINATRUST"
                 cboCreditCardType.AddItem (" 中國信託 信用卡")
                 cboCreditCardType.ItemData(lngIndex) = 1
                blnAdd = True
            Case "CREDITCARDHSBC"
                 cboCreditCardType.AddItem (" 匯豐銀行 信用卡")
                 cboCreditCardType.ItemData(lngIndex) = 2
                 blnAdd = True
            '#3531 多新增一個中國信託(新) By Kin 2007/10/31
            Case "CREDITCARDTRUSTNEW"
                cboCreditCardType.AddItem ("中國信託 信用卡(新)")
                cboCreditCardType.ItemData(lngIndex) = 3
                blnAdd = True
          End Select
          If blnAdd = True Then lngIndex = lngIndex + 1
          blnAdd = False
         rsTmp.MoveNext
     Loop
    rsTmp.Close
    Set rsTmp = Nothing
      If lngIndex = 0 Then
         MsgBox "銀行別資料未設，請設定相關銀行別代碼之後，重新執行  !!"
         SetCreditCardTypeCbo = False
    Else
         cboCreditCardType.ListIndex = 0
          SetCreditCardTypeCbo = True
    End If
    Screen.MousePointer = vbDefault
   
Exit Function
ChkErr:
  ErrSub Me.Name, "SetCreditCardTypeCbo"
End Function
'#2672 產生時間改為到時分秒 By Kin 2008/05/12
Private Sub gitCreateTime1_GotFocus()
  On Error Resume Next
    If gitCreateTime1.GetValue <> "" Then Exit Sub
    gitCreateTime1.SetValue Now
    
End Sub
'#2672 產生時間改為到時分秒 By Kin 2008/05/12
Private Sub gitCreateTime1_Validate(Cancel As Boolean)
  On Error GoTo ChkErr
   If Not IsDate(gitCreateTime1.Text) Then Exit Sub
   If gitCreateTime1.GetValue <> "" Then
      If gitCreateTime2.GetValue = "" Then gitCreateTime2.SetValue GetLastDate(gitCreateTime1.GetValue(True)) & "235959"
   End If
Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "gitCreateTime1_Validate")
End Sub

'#2672 產生時間改為到時分秒 By Kin 2008/05/12
Private Sub gitCreateTime2_GotFocus()
  On Error Resume Next
    If gitCreateTime2.GetValue <> "" Then Exit Sub
    gitCreateTime2.SetValue GetLastDate(Date) & "235959"
    
End Sub
'#2672 產生時間改為到時分秒 By Kin 2008/05/12
Private Sub gitCreateTime2_Validate(Cancel As Boolean)
  On Error Resume Next
    If gitCreateTime1.GetValue = "" Or gitCreateTime2.GetValue = "" Then Exit Sub
    If Not IsDate(gitCreateTime2.Text) Then Exit Sub
    If DateDiff("s", gitCreateTime1.GetValue(True), gitCreateTime2.GetValue(True)) < 0 Then
        MsgBox "產生時間截止日不可小於產生時間起始日!", vbExclamation, "警告"
        Cancel = True
    End If
End Sub

Private Sub gmBank_Change()
Call SetgiMulti(gmdCMCode, "CodeNo", "Description", "CD031", "代碼", "名稱", "")

End Sub

Private Sub CboAddItem()
    On Error Resume Next
        'If cboBillHeadFmt.ListCount > 0 Then Exit Sub
        If Not GetRS(rsBillHeadFmt, "Select BillHeadFmt,CitemCodeStr From " & GetOwner & "CD068") Then Exit Sub
        Do While Not rsBillHeadFmt.EOF
            cboBillHeadFmt.AddItem rsBillHeadFmt("BillHeadFmt") & ""
            rsBillHeadFmt.MoveNext
        Loop
End Sub

