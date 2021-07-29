VERSION 5.00
Object = "{3BB22E57-824A-11D3-BD70-0080C8F80BC4}#2.2#0"; "GiGridR.ocx"
Object = "{92575699-8EB8-11D3-B585-0080C8F80BC4}#3.3#0"; "GiNumber.ocx"
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#3.5#0"; "GiList.ocx"
Begin VB.Form frmSO3311F 
   BorderStyle     =   1  '單線固定
   Caption         =   "收費資料編輯  [SO3311F]"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   3495
   ClientWidth     =   11910
   BeginProperty Font 
      Name            =   "新細明體"
      Size            =   9.75
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SO3311F.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   11910
   Begin VB.CommandButton cmdShowRate 
      Caption         =   "F7.顯示費率表"
      Height          =   375
      Left            =   7350
      TabIndex        =   20
      Top             =   4620
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "F2. 存檔"
      Height          =   375
      Left            =   270
      TabIndex        =   17
      Top             =   4620
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&X)"
      Height          =   375
      Left            =   8970
      TabIndex        =   19
      Top             =   4620
      Width           =   1215
   End
   Begin VB.Frame fraData 
      Height          =   4515
      Left            =   300
      TabIndex        =   18
      Top             =   30
      Width           =   11475
      Begin VB.TextBox txtBillNo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1110
         MaxLength       =   15
         TabIndex        =   0
         Top             =   210
         Width           =   1830
      End
      Begin prjGiGridR.GiGridR ggrMduRate 
         Height          =   2025
         Left            =   8640
         TabIndex        =   22
         Top             =   690
         Visible         =   0   'False
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   3572
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin prjGiGridR.GiGridR ggrCustRate 
         Height          =   2025
         Left            =   5760
         TabIndex        =   21
         Top             =   690
         Visible         =   0   'False
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   3572
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin prjNumber.GiNumber gnbQuantity 
         Height          =   285
         Left            =   5070
         TabIndex        =   14
         Top             =   3270
         Visible         =   0   'False
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         MaxLength       =   3
      End
      Begin prjGiList.GiList gilPTCode 
         Height          =   285
         Left            =   1110
         TabIndex        =   13
         Top             =   4140
         Width           =   2940
         _ExtentX        =   5186
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin prjGiList.GiList gilClctEn 
         Height          =   285
         Left            =   1110
         TabIndex        =   12
         Top             =   3747
         Width           =   2940
         _ExtentX        =   5186
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin prjGiList.GiList gilSTCode 
         Height          =   285
         Left            =   1110
         TabIndex        =   11
         Top             =   3354
         Width           =   2940
         _ExtentX        =   5186
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin prjGiList.GiList gilUCCode 
         Height          =   285
         Left            =   1110
         TabIndex        =   10
         Top             =   2961
         Visible         =   0   'False
         Width           =   2940
         _ExtentX        =   5186
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin prjGiList.GiList gilCMCode 
         Height          =   285
         Left            =   1110
         TabIndex        =   8
         Top             =   2175
         Width           =   2940
         _ExtentX        =   5186
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin prjGiList.GiList gilCitemCode 
         Height          =   285
         Left            =   1110
         TabIndex        =   1
         Top             =   603
         Width           =   2940
         _ExtentX        =   5186
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Gi_Date.GiDate gdaRealStopDate 
         Height          =   285
         Left            =   3300
         TabIndex        =   5
         Top             =   1383
         Width           =   1125
         _ExtentX        =   1984
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
         Enabled         =   0   'False
      End
      Begin Gi_Date.GiDate gdaRealDate 
         Height          =   285
         Left            =   1110
         TabIndex        =   9
         Top             =   2568
         Width           =   1305
         _ExtentX        =   2302
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
         Enabled         =   0   'False
      End
      Begin Gi_Date.GiDate gdaRealStartDate 
         Height          =   285
         Left            =   1110
         TabIndex        =   4
         Top             =   1389
         Width           =   1125
         _ExtentX        =   1984
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
         Enabled         =   0   'False
      End
      Begin VB.TextBox txtManualNo 
         Height          =   285
         Left            =   5040
         MaxLength       =   10
         TabIndex        =   15
         Top             =   3675
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.TextBox txtNote 
         Height          =   285
         Left            =   5040
         MaxLength       =   250
         TabIndex        =   16
         Top             =   4035
         Width           =   5445
      End
      Begin prjNumber.GiNumber gnbRealPeriod 
         Height          =   285
         Left            =   3300
         TabIndex        =   3
         Top             =   996
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   503
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         MaxLength       =   2
      End
      Begin Gi_Date.GiDate gdaShouldDate 
         Height          =   285
         Left            =   1110
         TabIndex        =   2
         Top             =   996
         Width           =   1125
         _ExtentX        =   1984
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
      Begin prjNumber.GiNumber gnbShouldAmt 
         Height          =   285
         Left            =   1110
         TabIndex        =   6
         Top             =   1782
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         MaxLength       =   8
      End
      Begin prjNumber.GiNumber gnbRealAmt 
         Height          =   285
         Left            =   3300
         TabIndex        =   7
         Top             =   1770
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         MaxLength       =   8
      End
      Begin VB.Label lblQuantity 
         AutoSize        =   -1  'True
         Caption         =   "數量"
         Height          =   195
         Left            =   4170
         TabIndex        =   47
         Top             =   3270
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.Label lblManualNo 
         AutoSize        =   -1  'True
         Caption         =   "手開單號"
         Height          =   195
         Left            =   4125
         TabIndex        =   46
         Top             =   3675
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label lblNote 
         AutoSize        =   -1  'True
         Caption         =   "備註"
         Height          =   195
         Left            =   4125
         TabIndex        =   45
         Top             =   4095
         Width           =   390
      End
      Begin VB.Label lblCustRate 
         AutoSize        =   -1  'True
         Caption         =   "客戶類別費率表"
         Height          =   195
         Left            =   5790
         TabIndex        =   44
         Top             =   390
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.Label lblMduRate 
         AutoSize        =   -1  'True
         Caption         =   "大樓費率表"
         Height          =   195
         Left            =   8700
         TabIndex        =   43
         Top             =   390
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lblCustClass 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   7380
         TabIndex        =   42
         Top             =   390
         Width           =   45
      End
      Begin VB.Label lblMduName 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   9750
         TabIndex        =   41
         Top             =   390
         Width           =   45
      End
      Begin VB.Label lblRealPeriod 
         AutoSize        =   -1  'True
         Caption         =   "期數"
         Height          =   195
         Left            =   2370
         TabIndex        =   40
         Top             =   1020
         Width           =   390
      End
      Begin VB.Label lblRealAmt 
         AutoSize        =   -1  'True
         Caption         =   "實收金額"
         Height          =   195
         Left            =   2370
         TabIndex        =   39
         Top             =   1830
         Width           =   780
      End
      Begin VB.Label lblShouldAmt 
         AutoSize        =   -1  'True
         Caption         =   "出單金額"
         Height          =   195
         Left            =   210
         TabIndex        =   38
         Top             =   1800
         Width           =   780
      End
      Begin VB.Label lblUpdTime 
         Caption         =   " yyy/mm/dd hh:mm:ss"
         Height          =   285
         Left            =   7035
         TabIndex        =   37
         Top             =   3675
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.Label lblUpdEn 
         Caption         =   "XXXXXXXX"
         Height          =   315
         Left            =   9765
         TabIndex        =   36
         Top             =   3675
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.Label lblUTime 
         Caption         =   "異動時間:"
         Height          =   375
         Left            =   6105
         TabIndex        =   35
         Top             =   3675
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.Label lblUEn 
         Caption         =   "異動人員:"
         Height          =   375
         Left            =   8865
         TabIndex        =   34
         Top             =   3675
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label lblBillNo 
         AutoSize        =   -1  'True
         Caption         =   "單據編號"
         Height          =   195
         Left            =   210
         TabIndex        =   33
         Top             =   240
         Width           =   780
      End
      Begin VB.Label lblRealDate 
         AutoSize        =   -1  'True
         Caption         =   "入帳日期"
         Height          =   195
         Left            =   210
         TabIndex        =   32
         Top             =   2580
         Width           =   780
      End
      Begin VB.Label lblClctEn 
         AutoSize        =   -1  'True
         Caption         =   "收費人員"
         Height          =   195
         Left            =   210
         TabIndex        =   31
         Top             =   3750
         Width           =   780
      End
      Begin VB.Label lblSTCode 
         AutoSize        =   -1  'True
         Caption         =   "短收原因"
         Height          =   195
         Left            =   210
         TabIndex        =   30
         Top             =   3360
         Width           =   780
      End
      Begin VB.Label lblUCCode 
         AutoSize        =   -1  'True
         Caption         =   "未收原因"
         Height          =   195
         Left            =   210
         TabIndex        =   29
         Top             =   2970
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label lblCMCode 
         AutoSize        =   -1  'True
         Caption         =   "收費方式"
         Height          =   195
         Left            =   210
         TabIndex        =   28
         Top             =   2190
         Width           =   780
      End
      Begin VB.Label lblRealStopDate 
         AutoSize        =   -1  'True
         Caption         =   "至"
         Height          =   195
         Left            =   2370
         TabIndex        =   27
         Top             =   1440
         Width           =   195
      End
      Begin VB.Label lblRealStartDate 
         AutoSize        =   -1  'True
         Caption         =   "有效期限"
         Height          =   195
         Left            =   210
         TabIndex        =   26
         Top             =   1410
         Width           =   780
      End
      Begin VB.Label lblShouldDate 
         AutoSize        =   -1  'True
         Caption         =   "應收日期"
         Height          =   195
         Left            =   210
         TabIndex        =   25
         Top             =   1020
         Width           =   780
      End
      Begin VB.Label lblCitemCode 
         AutoSize        =   -1  'True
         Caption         =   "收費項目"
         Height          =   195
         Left            =   210
         TabIndex        =   24
         Top             =   630
         Width           =   780
      End
      Begin VB.Label lblPTCode 
         AutoSize        =   -1  'True
         Caption         =   "付款種類"
         Height          =   195
         Left            =   210
         TabIndex        =   23
         Top             =   4140
         Width           =   780
      End
   End
   Begin VB.Label lblEditMode 
      Alignment       =   2  '置中對齊
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  '單線固定
      Caption         =   "顯示"
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   10530
      TabIndex        =   48
      Top             =   4650
      Width           =   675
   End
End
Attribute VB_Name = "frmSO3311F"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Bruce
'2000/01/18
Option Explicit
Private lngEditMode As giEditModeEnu

'起始參數檔
Private Ypara7 As Integer  '有效期限警告日數
Private Ypara10 As Integer '使用費率表
Private Ypara12 As Integer '必需輸入短收原因
Private Ypara14 As Integer '地址依據
Private Ypara5 As Integer '次收日(1999/12/21)
Private Ypara8 As Integer '隨收費資料改變(1999/12/21)
Private Ypara9 As Integer '隨收費資料改變(1999/12/21)

'收費項目，期數，有效起始日，有效截止日、收費日結截止日之異動暫存變數
'借貸符號Sign
Private intXCitemCode As Integer
Private intXRealPeriod As Integer
Private strXRealStartDate As String
Private strXRealStopDate As String
Private strXLastDate As String
Private strXSign As String '借貸符號

Private intRefNo As Integer  '為收費項目之參考號

'宣告一記錄集合供本表單使用
Private rsTmpRec As New ADODB.Recordset

'接收派工單傳來之Connection
Public Conn As New ADODB.Connection

'判斷是否為週期性費用
Private blnPeriodFlag As Boolean

Private lngCustID As Long '由frmSo3311B傳來之客戶編號
Private strCustName As String '由frmso3311B傳來之客戶姓名
Private intBillMode As Integer '由So3311B傳來之單據狀態 1:單據編號
Private strNote As String '由frmso3311B傳來之備註資料
Private strBillNo As String '在新增時由frmso3311b傳來之單據編號
Private strPrtSno As String '印單序號
Private intItem As Integer '若修改時無法取得RowId即須使用BillNo and Item

Private strRowId As String '由frmSo3311B傳來之RowId修改時使用
Private strMdbSql As String '記錄在mdb中的So074欄位指令
Private rsMdbRec As New ADODB.Recordset '記錄在修改模式下的RecordSet
Private strServiceType As String
Private strCompCode As String

Private Sub cmdCancel_Click()
    If rsTmpRec.State = adStateOpen Then rsTmpRec.Close: Set rsTmpRec = Nothing
    Unload Me
End Sub

Private Sub cmdsave_Click()
On Error GoTo ChkErr
    If IsDataOk Then
        If ScrToRcd Then MsgBox "存檔失敗！請重新操作！", vbExclamation, "訊息！"
'        If Not ChkData Then MsgBox "存檔失敗！請重新操作！", vbExclamation, "訊息！"
        cmdCancel.Value = True
    End If
Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "cmdSave_Click")
End Sub

Private Sub cmdShowRate_Click()
On Error GoTo ChkErr
    Dim rsCustRate As New ADODB.Recordset
    Dim rsMduRate As New ADODB.Recordset
    Dim MfldsCustRate As New GiGridFlds
    Dim MfldsMduRate As New GiGridFlds
    Dim strCustRateSQL   As String
    Dim strMduRateSQL As String, strMudId As String
    
    ' 設定ggrCustRate 客戶費率檔 顯示收費項，期數，金額
    strCustRateSQL = "Select A.CustName,C.Description,B.Period,B.Amount "
    strCustRateSQL = strCustRateSQL & "From So001 A,CD019CD004 B,CD019 C Where  A.CustID=" & lngCustID & " and B.ClassCode=A.ClassCode1 and C.CodeNo=B.CitemCode"
    rsCustRate.CursorLocation = adUseClient
    rsCustRate.Open strCustRateSQL, gcnGi, adOpenForwardOnly, adLockReadOnly
    
    If Not rsCustRate.EOF Then
        MfldsCustRate.Add "Description", , , , False, "  收費項目  ", vbLeftJustify
        MfldsCustRate.Add "Period", , , , False, "期數", vbLeftJustify
        MfldsCustRate.Add "Amount", , , , False, "金額", vbLeftJustify
    Else
        If rsCustRate.State = adStateOpen Then rsCustRate.Close
        strCustRateSQL = "Select Description,PeriodFlag,Amount From CD019 Where PeriodFlag =1"
        rsCustRate.Open strCustRateSQL, gcnGi, adOpenForwardOnly, adLockReadOnly
        MfldsCustRate.Add "Description", , , , False, "  收費項目  ", vbLeftJustify
        MfldsCustRate.Add "PeriodFlag", , , , False, "期數", vbLeftJustify
        MfldsCustRate.Add "Amount", , , , False, "金額", vbLeftJustify
    End If
    ggrCustRate.AllFields = MfldsCustRate
    ggrCustRate.SetHead
    lblCustClass.Caption = "(" & GetValueFromRs2("Select CustName From So001 Where Custid=" & lngCustID) & ")"
    Set ggrCustRate.Recordset = rsCustRate
    ggrCustRate.Refresh
    '設定ggrMduRate 大樓費率檔
    MfldsMduRate.Add "Description", , , , False, "  收費項目  ", vbLeftJustify
    MfldsMduRate.Add "Period", , , , False, "期數", vbLeftJustify
    MfldsMduRate.Add "Amount", , , , False, "金額", vbLeftJustify
    ggrMduRate.AllFields = MfldsMduRate
    ggrMduRate.SetHead
    strMudId = GetValueFromRs2("Select Mduid From So001 Where Custid=" & lngCustID) & ""
    If strMudId <> "" Then
        rsMduRate.CursorLocation = adUseClient
        strMduRateSQL = "Select B.Description ,A.Period,A.Amount From CD019SO017 A,CD019 B Where A.MduId='" & strMudId & "' and B.CodeNo=A.CitemCode"
        rsMduRate.Open strMduRateSQL, gcnGi, adOpenForwardOnly, adLockReadOnly
        lblMduName = "(" & GetValueFromRs2("Select Name From So017 Where MduId='" & strMudId & "'") & ")"
        Set ggrMduRate.Recordset = rsMduRate
        ggrMduRate.Refresh
    End If
    ggrCustRate.Visible = True
    ggrMduRate.Visible = True
    lblCustRate.Visible = True
    lblMduRate.Visible = True
Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "cmdShowrate_Click")
End Sub

Private Sub Form_Activate()
On Error GoTo ChkErr
'設定初始時的Focus
        Select Case lngEditMode
            Case giEditModeInsert
                gilCitemCode.SetFocus
            Case giEditModeEdit
                gnbShouldAmt.SetFocus
            Case giEditModeView
                cmdCancel.SetFocus
        End Select
Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "Form_Activeate")
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ChkErr
    'F2存檔 F7顯示費率表
    If KeyCode = vbKeyF2 And cmdSave.Enabled Then Call cmdsave_Click: Exit Sub
    If KeyCode = vbKeyF7 And cmdShowRate.Enabled Then Call cmdShowRate_Click: Exit Sub
    'KeyCode = 0
Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "Form_KeyDown")
End Sub

Private Sub Form_Load()
    On Error GoTo ChkErr
    Dim strTmpSql As String
        If Not800600 Then
            Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2 + 1200
        End If
        
        If frmSO3311B.uSelBill = 1 Then
            lblBillNo.Caption = "單據編號"
        Else
            lblBillNo.Caption = "印單序號"
        End If
        '    取得系統參數
        '     Call GetGlobal
        
        strMdbSql = "Select * "
    
        '至最近日結參數記錄檔(So062)取收費日結截止日
        'frmSo3311 在使用
        If lngEditMode <> giEditModeView Then
            strTmpSql = "Select tranDate from So062 Where Type=1"
            rsTmpRec.CursorLocation = adUseClient
            rsTmpRec.Open strTmpSql, gcnGi, adOpenForwardOnly, adLockReadOnly
            If Not rsTmpRec.EOF Then strXLastDate = rsTmpRec(0).Value & ""
            rsTmpRec.Close
        End If
        
        '接收由So3311B傳來之客戶編號和客戶姓名！
        lngCustID = frmSO3311B.uCustID
        strCustName = frmSO3311B.uCustName
        strNote = frmSO3311B.uNote
        
        '至收費參數檔(So043)取各項參數！
        Call GetParameter
    
        '設定GiList控制項
        Call subGil
        
        '預設為週期性費用，因為收費項目預設值為第一值<1999/12/22>
        blnPeriodFlag = True '<1999/12/22>
    
        '設定異動模式
        Call ChangeMode(lngEditMode)
        PrcList Me
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "Form_Load")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ChkErr
    '判斷使用者是否在異動的狀態中！
        If cmdCancel.Caption = "取消(&X)" Or lngEditMode = giEditModeView Then
            If rsTmpRec.State = adStateOpen Then rsTmpRec.Close: Set rsTmpRec = Nothing
            Set frmSO3311F = Nothing
            Unload Me
            Exit Sub
        End If
            MsgBox "請先存檔或取消後再離開！", vbExclamation, "訊息"
            Cancel = True
Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "Form_QueryUnload")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
        Call ReleaseCOM(Me)

End Sub

Private Sub gdaRealStartDate_Validate(Cancel As Boolean)
On Error GoTo ChkErr
Dim lngChk As Long
If Not IsDateOk(gdaRealStartDate.GetValue) Then gdaRealStartDate.SetFocus: Exit Sub
'記錄SFGetAmount傳回之值
Dim strStopDate As String, lngRealAmount As Long, strMsg As String
'記錄RealStartDate ,RealStopDate 之資料
Dim strRealStartDate As String, strRealStopDate As String
strRealStartDate = CDate(StrToDate(gdaRealStartDate.GetValue & ""))
strRealStopDate = gdaRealStopDate.GetValue & ""

    If strRealStartDate = "" Or CDate(strRealStartDate) > CDate(StrToDate(strRealStopDate)) Then MsgBox "此欄位需有值且<=截止日期": gdaRealStartDate.SetFocus: Exit Sub
    If CDate(StrToDate(strXRealStartDate)) <> CDate(strRealStartDate) Then
        lngChk = SFGetAmount(False, 2, lngCustID, gilCitemCode.GetCodeNo, gnbRealPeriod.Text, gdaRealStartDate.GetValue, gServiceType, strStopDate, lngRealAmount, strMsg)
        If lngChk < 0 Then
            MsgBox strMsg, vbExclamation, "訊息"
            gdaRealStartDate.SetFocus
        Else
            gnbShouldAmt.Text = lngRealAmount
            Call gdaRealStopDate.SetValue(strStopDate)
            '設定各個變數以供下列程式使用
            strXRealStartDate = gdaRealStartDate.GetValue
            strXRealStopDate = gdaRealStopDate.GetValue
        End If
    End If
Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "gdaRealStartDate_VAlidate")
End Sub

Private Sub gdaRealStopDate_Validate(Cancel As Boolean)
'On Error GoTo ChkErr
Dim lngChk As Long, strVBYesNo As String
If Not IsDateOk(gdaRealStopDate.GetValue) Then gdaRealStopDate.SetFocus: Exit Sub
'記錄RealStartDate ,RealStopDate 之資料
Dim strRealStartDate As String, strRealStopDate As String
strRealStartDate = StrToDate(gdaRealStartDate.GetValue & "") 'StrToDate YYYY/MM/DD
strRealStopDate = gdaRealStopDate.GetValue & ""
'記錄SFGetAmount傳回之值
Dim strStopDate As String, lngRealAmount As Long, strMsg As String

    If strRealStopDate = "" Or CDate(strRealStartDate) > CDate(StrToDate(strRealStopDate)) Then MsgBox "此欄位需有值且<=截止日期": gdaRealStopDate.SetFocus: Exit Sub
    If CDate(StrToDate(strXRealStopDate)) <> CDate(StrToDate(strRealStopDate)) Then
        '若本欄日期減去有效起始日期大於intYPara7，則顯示確認視窗
        If DateDiff("d", CDate(strRealStartDate), CDate(StrToDate(strRealStopDate))) > Ypara7 Then
                If vbNo = MsgBox("有效期限超過系統預設天數，是否接受此日期！", vbExclamation, "訊息") Then gdaRealStopDate.SetFocus: Exit Sub
        End If
        lngChk = SFGetAmount(True, 2, lngCustID, gilCitemCode.GetCodeNo, gnbRealPeriod.Text, gdaRealStartDate.GetValue, gServiceType, strRealStopDate, lngRealAmount, strMsg)
        If lngChk < 0 Then
            MsgBox strMsg, vbExclamation, "訊息"
            gdaRealStopDate.SetFocus
        Else
            If GetNumberType(lngRealAmount & "") <> 0 Then
                 gnbShouldAmt.Text = Val(lngRealAmount & "")
                '設定各個變數以供下列程式使用
                strXRealStartDate = gdaRealStartDate.GetValue & ""
                strXRealStopDate = gdaRealStopDate.GetValue & ""
            End If
        End If
    End If
Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "gdaRealStopDate_VAlidate")
End Sub

Private Sub gdaShouldDate_Validate(Cancel As Boolean)
    If gdaShouldDate.Text = "" Then Exit Sub
'    If Not IsDateOk(gdaShouldDate.GetValue) Then MsgBox "日期不合法！": gdaShouldDate.SetFocus: Exit Sub
End Sub

Private Sub ggrCustRate_ShowCellData(ByVal Col As Long, ByVal Row As Long, ByVal giFld As prjGiGridR.GiGridFld, ByVal Fld As ADODB.Field, Value As Variant)
On Error GoTo ChkErr
    If Fld.Name = "AMOUNT" Then
        Value = Format(Value, "###,###,###")
    End If
Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "ggrCustRate_ShowCellDate")
End Sub

Private Sub ggrMduRate_ShowCellData(ByVal Col As Long, ByVal Row As Long, ByVal giFld As prjGiGridR.GiGridFld, ByVal Fld As ADODB.Field, Value As Variant)
On Error GoTo ChkErr
    If Fld.Name = "AMOUNT" Then
        Value = Format(Value, "###,###,###")
    End If
Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "ggrMduRate_ShowCellDate")
End Sub

Private Sub gilCitemCode_Validate(Cancel As Boolean)
'On Error GoTo ChkErr
Dim strTmpSql As String, intAmount As Long, intPeriod As Integer
Dim blnPeriodFlagMode As Boolean
Dim rsTmpCitemCode As New ADODB.Recordset
'記錄SFGetAmount 傳回的相關資料
Dim strStopDate As String, lngRealAmount As Long, strMsg As String
Dim lngChk As Long
    
    If gilCitemCode.GetCodeNo = "" Or gilCitemCode.GetDescription = "" Then MsgBox "此欄位必需有值", vbExclamation, "訊息": gilCitemCode.SetFocus: Exit Sub
    '若有值且有改變
    If intXCitemCode <> gilCitemCode.GetCodeNo Then
        '至收費項目代碼檔取（是否週期性費用），（預設金額），(預設期數）<<1999/12/22>>，（參考號）<<1999/12/22>>,(借貸符號,Sign）<<2000/03/10>>
        strTmpSql = "Select PeriodFlag,Amount,Period,RefNo,Sign From CD019 Where CodeNo=" & gilCitemCode.GetCodeNo
        
        '開啟連線！
        If OpenRecordset(rsTmpCitemCode, strTmpSql, gcnGi, adOpenForwardOnly, adLockReadOnly, adUseServer, False, False) <> giOK Then MsgBox "連線失敗！請重新操作！", vbExclamation, "訊息！"
        
        blnPeriodFlag = rsTmpCitemCode("PeriodFlag").Value
        strXSign = rsTmpCitemCode("Sign").Value
        intAmount = GetNumberType(rsTmpCitemCode("Amount").Value & "")
        intPeriod = Val(rsTmpCitemCode("Period").Value & "")      '<<1999/12/22>>預設期數
        intRefNo = Val(rsTmpCitemCode("RefNO").Value & "")        '<<1999/12/22>>參考號
        rsTmpCitemCode.Close
        Set rsTmpCitemCode = Nothing
        '若（是否週期性費費用）=否 則：
        '▲期數=0，有效起始日期/截止日期=無，出單金額=（預設金額）
        '▲Disable 期數，有效起日期/截止日期
        If blnPeriodFlag = False Then
            blnPeriodFlagMode = False
            gnbRealPeriod.Text = ""
            Call gdaRealStopDate.SetValue("")
            Call gdaRealStartDate.SetValue("")
            gnbShouldAmt.Text = intAmount
        Else
            gnbRealPeriod.Text = intPeriod '<<1999/12/22>> Old gnbRealPeriod.Text =1
            
           Call gdaRealStartDate.SetValue(ReturnRealStartDate(gilCitemCode.GetCodeNo, lngCustID)) '<<1999/12/21>>
           gdaShouldDate.SetValue gdaRealStartDate.GetValue
            lngChk = SFGetAmount(False, 2, lngCustID, gilCitemCode.GetCodeNo, gnbRealPeriod.Text, gdaRealStartDate.GetValue, gServiceType, strStopDate, lngRealAmount, strMsg)
            If lngChk < 0 Then
                MsgBox strMsg, vbExclamation, "訊息"
                On Error Resume Next
                gilCitemCode.SetFocus
                On Error GoTo ChkErr
            Else
                gnbShouldAmt.Text = lngRealAmount
                Call gdaRealStopDate.SetValue(strStopDate & "")
            End If
        End If
        gnbRealAmt.Text = gnbShouldAmt.Value
        
        gnbRealPeriod.Enabled = blnPeriodFlagMode
        gdaRealStartDate.Enabled = blnPeriodFlagMode
        gdaRealStopDate.Enabled = blnPeriodFlagMode
        intXCitemCode = Val(gilCitemCode.GetCodeNo & "")
        intXRealPeriod = Val(gnbRealPeriod.Text & "")
        strXRealStartDate = gdaRealStartDate.GetValue & ""
        strXRealStopDate = gdaRealStopDate.GetValue & ""
     End If
Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "gilCitemCode_Validate")
End Sub

Private Sub gnbQuantity_GotFocus()
    With gnbQuantity
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub gnbRealAmt_GotFocus()
    With gnbRealAmt
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub gnbRealAmt_Validate(Cancel As Boolean)
On Error GoTo ChkErr
Dim intShouldAmt As Long
Dim intRealAmt As Long
intShouldAmt = gnbShouldAmt.Value
intRealAmt = gnbRealAmt.Value
'若有值，則清除未收原因
'若有值且出單金額有值：
'   '、若二者不相等，則警告："實收與出單金額不相等",並填入短收原因初值1及對應名稱
    '、若二者相等，則清除短收原因
    
    If gnbRealAmt.Value < 0 Then
            gnbRealAmt.Text = (-1) * gnbRealAmt.Value
    End If
    If intRealAmt > 0 Then
        gilUCCode.Clear
    Else
        gilSTCode.Clear
        gilUCCode.ListIndex = 1
    End If

    If intShouldAmt > 0 And intRealAmt > 0 Then
        If intShouldAmt <> intRealAmt Then
            MsgBox "實收與出單金額不相等,短收原因須有值"
            If intRealAmt > intShouldAmt Then gnbRealAmt.Text = 0
            'gilSTCode.ListIndex = 1
            gilSTCode.SetFocus
        ElseIf intShouldAmt = intRealAmt Then
            gilSTCode.Clear
        End If
    End If
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "gnbRealAmt_VailDate")
End Sub

Private Sub gnbRealPeriod_GotFocus()
    With gnbRealPeriod
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub gnbRealPeriod_Validate(Cancel As Boolean)
On Error GoTo ChkErr
Dim lngChk As Long, strStopDate As String
Dim lngRealAmount As Long, strMsg As String
    '若期數在Enable的狀態下須有值
If Not gnbRealPeriod.Enabled Then Exit Sub

    '若無值或<=0則錯誤:"此欄位必需為正值！"且離開
    If gnbRealPeriod.Text <= 0 Or str(gnbRealPeriod.Text) = "" Then MsgBox "此欄位必需有值且需為正值", vbExclamation, "訊息": Exit Sub
    
    If intXRealPeriod <> gnbRealPeriod.Text Then
    lngChk = SFGetAmount(False, 2, lngCustID, gilCitemCode.GetCodeNo, gnbRealPeriod.Text, gdaRealStartDate.GetValue, gServiceType, strStopDate, lngRealAmount, strMsg)
        If lngChk < 0 Then
            MsgBox strMsg, vbExclamation, "訊息"
            gnbRealPeriod.SetFocus
        Else
            If lngRealAmount = 0 Then lngChk = SFGetAmount(True, 2, lngCustID, gilCitemCode.GetCodeNo, gnbRealPeriod.Text, gdaRealStartDate.GetValue, gServiceType, gdaRealStopDate.GetValue, lngRealAmount, strMsg)
            If lngChk < 0 Then
                MsgBox strMsg, vbExclamation, "訊息"
                gnbRealPeriod.SetFocus
                Exit Sub
            End If
            gnbShouldAmt.Text = lngRealAmount
            Call gdaRealStopDate.SetValue(strStopDate)
            '設定各個變數以供下列程式使用
            intXRealPeriod = gnbRealPeriod.Text
            strXRealStartDate = gdaRealStartDate.GetValue
            strXRealStopDate = gdaRealStopDate.GetValue
        End If
    End If
Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "gnbRealPeriod_Validate")
End Sub

Private Sub gnbShouldAmt_GotFocus()
    With gnbShouldAmt
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub gnbShouldAmt_Validate(Cancel As Boolean)
On Error GoTo ChkErr
Dim intShouldAmt As Long
Dim intRealAmt As Long
intShouldAmt = gnbShouldAmt.Value
intRealAmt = gnbRealAmt.Value
    '出單金額若有值且實收金額有值：
        '若二者不相等，則警告
        '若二者相等，則清除短收原因<1999/12/27>
        If gnbShouldAmt.Value < 0 Then
            gnbShouldAmt.Text = (-1) * gnbShouldAmt.Value
        End If
        If intRealAmt > 0 And intShouldAmt > 0 Then
            If intRealAmt <> intShouldAmt Then
                MsgBox "實收與出單金額不相等,短收原因須有值"
                If intRealAmt > intShouldAmt Then
                        gnbRealAmt.Text = 0
                ElseIf intRealAmt = 0 Then
                        gilSTCode.Clear
                        gilUCCode.ListIndex = 1
                    Else
                        'gilSTCode.ListIndex = 1
                    End If
            ElseIf intRealAmt = intShouldAmt Then
                gilSTCode.Clear
            End If
        End If
Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "gnbShouldAmt_Validate")
End Sub

Private Sub txtManualNo_GotFocus()
    With txtManualNo
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtNote_GotFocus()
    With txtNote
        .IMEMode = 1
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub ChangeMode(ByVal lngMode As giEditModeEnu)
On Error GoTo ChkErr
    Dim blnFlag As Boolean  ' 記錄是否在資料瀏覽模式，
    lngEditMode = lngMode
    
    Select Case lngEditMode
      Case giEditModeInsert
           blnFlag = True
           lblEditMode.Caption = "新增"
           cmdCancel.Caption = "取消(&X)"
           Call NewRec
      Case giEditModeEdit
           blnFlag = False
           lblEditMode.Caption = "修改"
           cmdCancel.Caption = "取消(&X)"
           Call RcdToScr
    End Select
    '設定各個物件的Enabled/Disabled
    gilCitemCode.Enabled = blnFlag  '收費項目
    gdaShouldDate.Enabled = blnFlag '應收日期
    
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "ChangeMode")
End Sub

' 自訂屬性：EditMode
' 記錄目前在編輯、新增或檢視模式
' giEditModeEnu(自訂列舉值，設定於Sys_Lib)
Public Property Get EditMode() As giEditModeEnu
  On Error GoTo ChkErr
    ' 取目前編輯模式
    EditMode = lngEditMode
  Exit Property
ChkErr:
   Call ErrSub(Me.Name, "Get EditMode")
End Property

Public Property Let EditMode(ByVal vNewValue As giEditModeEnu)
  On Error GoTo ChkErr
    '設定編輯模式
    lngEditMode = vNewValue
  Exit Property
ChkErr:
    Call ErrSub(Me.Name, "Let EditMode")
End Property

Private Sub NewRec()
On Error GoTo ChkErr
    Dim strShouldDate As String
    
    '印單序號
    txtBillNo.Text = IIf(frmSO3311B.uSelBill = 2, strPrtSno, strBillNo)
'    gilCitemCode
    '收費項目
    Call gilCMCode.SetCodeNo("")
    Call gilCMCode.SetDescription("")
    '入帳日期
    Call gdaRealDate.SetValue(IIf(frmSO3311B.uRealDate <> "", frmSO3311B.uRealDate, ""))
    '收費方式
     gilCitemCode.ListIndex = 1
    'Call gilCitemCode.SetDescription(IIf(frmSO3311B.uCMName <> "", frmSO3311B.uCMName, ""))
    
    '收費人員
    Call gilClctEn.SetCodeNo(IIf(frmSO3311B.uClctEn <> "", frmSO3311B.uClctEn, ""))
    Call gilClctEn.SetDescription(IIf(frmSO3311B.uClctName <> "", frmSO3311B.uClctName, ""))
    '付款種類
    Call gilPTCode.SetCodeNo(IIf(frmSO3311B.uPTCode <> "", frmSO3311B.uPTCode, ""))
    Call gilPTCode.SetDescription(IIf(frmSO3311B.uPTName <> "", frmSO3311B.uPTName, ""))
    '異動時間
    lblUpdTime.Caption = GetDTString(Now)
    '異動人員
    lblUpdEn.Caption = garyGi(1)
    '未收原因
    gilUCCode.Clear
    '短收原因
    gilSTCode.Clear
    '應收日期
    gdaShouldDate.SetValue (IIf(frmSO3311B.uRealDate <> "", frmSO3311B.uRealDate, ""))
    '有效起始日
    gdaRealStartDate.SetValue ("")
    '有效截止日
    gdaRealStopDate.SetValue ("")
    '期數
    gnbRealPeriod.Text = ""
    '實收金額
    gnbRealAmt.Text = ""
    '應收金額
    gnbShouldAmt.Text = ""
    '數量
    gnbQuantity.Text = ""
    '初始值
    intXCitemCode = 0
    intXRealPeriod = 0
    strXSign = "+"
Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "NewRec")
End Sub

''接收上層RecordSET rsSo1132A_View V_Charge
'Public Property Get rsInputRec() As ADODB.Recordset
'On Error GoTo ChkErr
'    Set AddBillNo = rsSo1132ARec
'Exit Property
'ChkErr:
'    Call ErrSub(Me.Name, "prtItem")
'End Property
'
''設定上層RecordSET rsSo1132A_View V_Charge
'Public Property Let rsInputRec(ByVal rsNewValue As ADODB.Recordset)
'On Error GoTo ChkErr
'    Set rsSo1132ARec = rsNewValue
'    strBillNo = rsSo1132ARec("BillNo").Value & ""
'    intItem = Val(rsSo1132ARec("Item").Value & "")
''    rsSo1132ARec.Close
'Exit Property
'ChkErr:
'    Call ErrSub(Me.Name, "proItem")
'End Property

Private Sub GetParameter()
On Error GoTo ChkErr
    Dim rsSO043 As New ADODB.Recordset
    Dim strSo043Sql As String
    
    strSo043Sql = "Select Para7,Para10,Para12,Para14,Para5,Para8,Para9 From So043"
    rsSO043.Open strSo043Sql, gcnGi, adOpenForwardOnly, adLockReadOnly
        Ypara5 = rsSO043("Para5").Value 'Update(199/12/21)
        Ypara8 = rsSO043("Para8").Value '(199/12/21)
        Ypara9 = rsSO043("Para9").Value '(199/12/21)
        Ypara7 = rsSO043("Para7").Value
        Ypara10 = rsSO043("Para10").Value
        Ypara12 = rsSO043("Para12").Value
        Ypara14 = rsSO043("Para14").Value
    rsSO043.Close
Exit Sub
ChkErr:
    rsSO043.Close
    Call ErrSub(Me.Name, "GetParameter")
End Sub

Private Sub RcdToScr()
On Error GoTo ChkErr
Dim strSelSql As String
    
    If rsMdbRec.State = adStateOpen Then rsMdbRec.Close
    If strRowId <> "" Then
        strSelSql = strMdbSql & " From So3311 Where Rowid='" & strRowId & "'"
    Else
        strSelSql = strMdbSql & " From So3311 Where Billno='" & strBillNo & "' And Item = " & intItem
    End If
    If Not GetRS(rsMdbRec, strSelSql, Conn, adUseClient, adOpenKeyset, adLockOptimistic) Then
        MsgBox "連線失敗！請重新操作！", vbExclamation, "訊息！": Unload Me
    End If
    txtBillNo.Text = rsMdbRec("BillNo").Value                               '單據編號
    Call gilCitemCode.SetCodeNo(rsMdbRec("CitemCode").Value & "")           '收費項目代號
    Call gilCitemCode.SetDescription(rsMdbRec("CitemName").Value & "")      '收費項目名稱
    Call gdaShouldDate.SetValue(Format(rsMdbRec("ShouldDate").Value & "", "YYYYMMDD"))  '應收日期
    
     gnbShouldAmt.Text = IIf(GetNumberType(rsMdbRec("ShouldAmt").Value & "") = 0, "", GetNumberType(rsMdbRec("ShouldAmt").Value & ""))              '出單金額
     If gnbShouldAmt.Value < 0 Then
        strXSign = "-"
     End If
     gnbRealAmt.Text = GetNumberType(rsMdbRec("RealAmt").Value & "")     '實金金額
     '-----------------------------------------------------------------------------------------
    '若為修改模式：若期數 >0 ，則Enable 期數,有效起始／截止日,反之則disable
    If rsMdbRec("RealPeriod").Value > 0 Then
        blnPeriodFlag = True
    Else
        blnPeriodFlag = False
    End If
    gnbRealPeriod.Enabled = blnPeriodFlag
    gdaRealStartDate.Enabled = blnPeriodFlag
    gdaRealStopDate.Enabled = blnPeriodFlag
    If rsMdbRec("RealPeriod").Value > 0 Then
            gnbRealPeriod.Text = IIf(GetNumberType(rsMdbRec("RealPeriod").Value) = 0, "", GetNumberType(rsMdbRec("RealPeriod").Value))        '期數
            Call gdaRealStartDate.SetValue(Format(rsMdbRec("RealStartDate") & "", "YYYYMMDD"))      '有效起始日
            Call gdaRealStopDate.SetValue(Format(rsMdbRec("RealStopDate") & "", "YYYYMMDD"))     '有效截止日
            '設定暫存變數內容
            intXCitemCode = GetNumberType(rsMdbRec("CitemCode").Value & "")
            intXRealPeriod = rsMdbRec("RealPeriod").Value
            strXRealStartDate = Format(rsMdbRec("RealStartDate").Value & "", "YYYYMMDD")
    End If
     
'-----------------------------------------------------------------------------------------
     Call gilCMCode.SetCodeNo(rsMdbRec("CMCode").Value & "")              '收費代碼
     Call gilCMCode.SetDescription(rsMdbRec("CMName").Value & "")                   '名稱
     Call gdaRealDate.SetValue(Format((rsMdbRec("RealDate").Value & ""), "YYYYMMDD"))         '實收日期
     'Call gilUCCode.SetCodeNo(rsMdbRec("UCCode").Value & "")                        '未收原因代碼
     'Call gilUCCode.SetDescription(rsMdbRec("UCName").Value & "")                   '未收原因名稱
     Call gilSTCode.SetCodeNo(rsMdbRec("STCode").Value & "")                        '短收原因代碼
     Call gilSTCode.SetDescription(rsMdbRec("STName").Value & "")                   '短收原因名稱
     Call gilClctEn.SetCodeNo(rsMdbRec("ClctEn").Value & "")                        '收費人員代碼
     Call gilClctEn.SetDescription(rsMdbRec("ClctName").Value & "")                 '收費人員名稱
     Call gilPTCode.SetCodeNo(rsMdbRec("PTCode").Value & "")                        '付款種類代碼
     Call gilPTCode.SetDescription(rsMdbRec("PTName").Value & "")                   '付款種類名稱
     txtManualNo.Text = frmSO3311B.uManualNo                         '手開單號
     txtNote.Text = rsMdbRec("Note").Value & ""                                      '備註
     'strSign = GetRsValue("Select Sign From CD019 Where CodeNo = " & gilCitemCode.GetCodeNo) & ""
     
     '判斷如果是顯示模式就將rsTmpRec結束，否則即為編輯模式須再提供存檔時使用
     
Exit Sub
ChkErr:
     Call ErrSub(Me.Name, "RcdToScr")
End Sub

'新增時取得帳單編號
Private Function GetBillNo(ByVal strDate As Date, ByVal strBillNo As String) As String
On Error Resume Next
Dim str_BillNo As String
    str_BillNo = Year(strDate) & String(2 - Len(Month(strDate)), "0") & Month(strDate)
    str_BillNo = str_BillNo & "T" & String(8 - Len(strBillNo), "0") & strBillNo
    GetBillNo = str_BillNo
End Function

Private Function IsDataOk() As Boolean
On Error GoTo ChkErr
      IsDataOk = False
      '出單金額不得小於等於0，或空值！
      
'      If gnbShouldAmt.Value <= 0 Or gnbShouldAmt.Text = "" Then MsgBox "出單金額不得小於等於零！", vbExclamation, "訊息！": Exit Function
      
      '必要欄位：單據編號、項目，收費項目代碼，應收日期，數量，收費方式代碼
      If gdaShouldDate.GetValue = "" Then MsgBox gMsgIsDataOK, vbExclamation, "訊息": gdaShouldDate.SetFocus: Exit Function
      If gilCitemCode.GetCodeNo = "" Then MsgBox gMsgIsDataOK, vbExclamation, "訊息": gilCitemCode.SetFocus: Exit Function
      If gilCMCode.GetCodeNo = "" Then MsgBox gMsgIsDataOK, vbExclamation, "訊息": gilCMCode.SetFocus: Exit Function
     
      '若期數>0，則有效起始日與截止日需有值，且起始日<=截止日
      If GetNumberType(gnbRealPeriod.Text & "") > 0 Then
          If gdaRealStartDate.Text = "" Then MsgBox "有效起始日須有值！", vbExclamation, "訊息！": gdaRealStartDate.SetFocus: Exit Function
          If gdaRealStopDate.Text = "" Then MsgBox "有效截止日須有值！", vbExclamation, "訊息！": gdaRealStopDate.SetFocus: Exit Function
          If (CDate(StrToDate(gdaRealStartDate.GetValue & "")) > CDate(StrToDate(gdaRealStopDate.GetValue))) Then
                  MsgBox "有效起始日與截止日需有值，且起始日須<=截止日"
                  Exit Function
           End If
      End If
      
      '若出單金額>0，且出單金額!=實收金額，則短收原因需有值(OLD)
      '若出單金額 <> 0，且出單金額!=實收金額，則短收原因需有值<2000/03/10>
      
       If Val(gnbShouldAmt.Text & "") <> 0 And (Val(gnbRealAmt.Text & "") <> Val(gnbShouldAmt.Text & "")) Then
           If gilSTCode.GetCodeNo = "" Then
                MsgBox "出單金額不等於實收金額需有短收原因！", vbExclamation, "訊息"
                gilSTCode.SetFocus
                Exit Function
            End If
       End If
     
    IsDataOk = True
Exit Function
ChkErr:
    Call ErrSub(Me.Name, "IsDataOK")
End Function

Private Function ReturnDate(ByVal strDate As String) As Date
    On Error GoTo ChkErr
        If strDate <> "" Then
            ReturnDate = CDate(strDate) + Ypara5
        End If
    Exit Function
ChkErr:
        Call ErrSub(Me.Name, "Returndate")
End Function

Private Sub GetPeriod_Amt(ByRef intReturnPeriod As Integer, ByRef intReturnAmt As Long)
    On Error GoTo ChkErr
        '若隨收費資料改變2（YPara9)為1（原應收額），則：
        '   期數=<該筆原收費期數>，金額=<該筆原應收金額>
        '若YPara9為 2（出單金額），則：
        '   期數=<該筆收費期數>，金額=<該筆出單金額>
    Dim rsTmp As New ADODB.Recordset
    Dim strsql As String
        rsTmp.MaxRecords = 1
        strsql = "Select OldAmt,OldPeriod From So033 Where BillNo='" & strBillNo & "' and Item =" & intItem
        If Ypara9 = 1 Then
                If OpenRecordset(rsTmp, strsql, gcnGi, adOpenForwardOnly, adLockReadOnly, adUseClient, False, False) <> giOK Then MsgBox "連接失敗！請重新操作！", vbExclamation, "訊息！"
                intReturnPeriod = rsTmp("Oldperiod").Value
                intReturnAmt = rsTmp("OldAmt").Value
                rsTmp.Close
                Set rsTmp = Nothing
        Else
                intReturnPeriod = Val(gnbRealPeriod.Text)
                intReturnAmt = Val(gnbShouldAmt.Text)
        End If
        
    Exit Sub
ChkErr:
        Call ErrSub(Me.Name, "GetPeriod_Amt")
End Sub

Private Function ScrToRcd() As Boolean
    On Error GoTo ChkErr
    Dim rsScr As New ADODB.Recordset '在ScrtoRcd 使用的Recordset
    Dim strsql As String
    Dim lngItem As Long '記錄要儲存時的項目編號
        If Not IsDataOk Then Exit Function
        If frmSO3311B.uSelBill = 1 Then
            strsql = "Select Max(Item)+1 as MaxItem From So033 Where Billno='" & strBillNo & "'"
        Else
            strsql = "Select Max(Item)+1 as MaxItem From So033 Where PrtSno='" & strPrtSno & "'"
        End If
        If rsMdbRec.State = adStateClosed Then If Not GetRS(rsMdbRec, "Select * From So3311", Conn, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
        
        If OpenRecordset(rsScr, strsql, gcnGi, adOpenForwardOnly, adLockReadOnly, adUseClient, False) <> giOK Then MsgBox "連線失敗！請重新操作！", vbExclamation, "訊息！": Exit Function
        '記錄項目編號
        
        If Not (rsScr.EOF Or IsNull(rsScr("MaxItem").Value)) Then
            lngItem = rsScr("MaxItem").Value
        Else
            lngItem = 1
        End If
        rsScr.Close
        Set rsScr = Nothing
        
        
        With rsMdbRec
        
        If lngEditMode = giEditModeInsert Then
            .AddNew
        End If
            .Fields("BillNo").Value = strBillNo
            .Fields("PrtSno").Value = strPrtSno
            .Fields("Item").Value = lngItem
            .Fields("CustName").Value = strCustName
            .Fields("CustId").Value = lngCustID
            .Fields("CItemCode").Value = gilCitemCode.GetCodeNo
            .Fields("CItemName").Value = gilCitemCode.GetDescription
            .Fields("ShouldAmt").Value = NoZero(gnbShouldAmt.Value, True)
            
            .Fields("ShouldDate").Value = Format(gdaShouldDate.GetValue, "####/##/##")
            If gnbRealPeriod.Value > 0 Then
                .Fields("RealPeriod").Value = NoZero(gnbRealPeriod.Value, True)
                .Fields("RealStartDate").Value = Format(gdaRealStartDate.GetValue, "####/##/##")
                .Fields("RealStopDate").Value = Format(gdaRealStopDate.GetValue, "####/##/##")
            Else
                .Fields("RealPeriod").Value = 0
                .Fields("RealStartDate").Value = Null
                .Fields("RealStopDate").Value = Null
            End If
            
            .Fields("EntryEn").Value = garyGi(0)
            .Fields("Note").Value = txtNote & ""
            .Fields("CancelFlag").Value = 0
            
            If gilClctEn.GetDescription <> "" Then
                .Fields("ClctEn").Value = gilClctEn.GetCodeNo
                .Fields("ClctName").Value = gilClctEn.GetDescription
            Else
                .Fields("ClctEn").Value = Null
                .Fields("ClctName").Value = Null
            End If
            
            .Fields("CmCode").Value = gilCMCode.GetCodeNo
            .Fields("CMName").Value = gilCMCode.GetDescription
                            
            If gilPTCode.GetDescription <> "" Then
                .Fields("PTCode").Value = gilPTCode.GetCodeNo
                .Fields("PTName").Value = gilPTCode.GetDescription
            Else
                .Fields("PTCode").Value = Null
                .Fields("PTName").Value = Null
            End If
            
            .Fields("RowID").Value = strRowId
            
            If gnbRealAmt.Value > 0 Then
                .Fields("RealAmt").Value = NoZero(gnbRealAmt.Value, True)
            Else
                .Fields("RealAmt").Value = 0
            End If
            
            If gilSTCode.GetDescription <> "" Then
                .Fields("STCode").Value = gilSTCode.GetCodeNo
                .Fields("STName").Value = gilSTCode.GetDescription
            Else
                .Fields("STCode").Value = Null
                .Fields("STName").Value = Null
            End If
            '.Fields("ManualNo") = NoZero(txtManualNo)
            .Fields("ServiceType") = gServiceType
            .Fields("CompCode") = strCompCode
            .Update
        End With
        rsMdbRec.Close
        Set rsMdbRec = Nothing
    Exit Function
ChkErr:
    Call ErrSub(Me.Name, "ScrToRcd")
End Function


Private Function GetNumberType(ByVal strAmount As String) As Long
    On Error Resume Next
        GetNumberType = Val(Format(strAmount, "00000000"))
End Function

Private Sub txtNote_Validate(Cancel As Boolean)
    On Error Resume Next
        txtNote.IMEMode = 2
End Sub

Private Sub subGil()
    On Error GoTo ChkErr
        SetgiList gilCitemCode, "CodeNo", "Description", "CD019", , , , , 3, 12, "Where ServiceType ='" & gServiceType & "'"
        gilCitemCode.Filter = "Where ServiceType ='" & gServiceType & "'"
        gilCitemCode.ListIndex = 1
        SetgiList gilCMCode, "CodeNo", "Description", "CD031", , , , , 3, 12
        SetgiList gilUCCode, "CodeNo", "Description", "CD013", , , , , 3, 12
        SetgiList gilSTCode, "CodeNo", "Description", "CD016", , , , , 3, 12
        SetgiList gilClctEn, "EmpNo", "EmpName", "CM003", , , , , 10, 20
        SetgiList gilPTCode, "CodeNo", "Description", "CD032", , , , , 3, 12
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "subGil")
End Sub

'接收由SO3311b傳來之ROWID 修改時使用！
Public Property Let uRowId(ByVal vData As String)
    strRowId = vData
End Property

'接收由So3311B傳來之單據編號或印單序號！新增時使用！
Public Property Let uBillNo(ByVal vData As String)
    strBillNo = vData
End Property

'接收由So3311b傳來之印單序號
Public Property Let uPrtSNo(ByVal vData As String)
    strPrtSno = vData
End Property

'由So3311B傳來之項目編號
Public Property Let uItem(ByVal vData As Integer)
    intItem = vData
End Property

'接收由So3311b傳來之服務類別
Public Property Let uServiceType(ByVal vData As String)
    strServiceType = vData
End Property

'接收由So3311b傳來之公司別
Public Property Let uCompCode(ByVal vData As String)
    strCompCode = vData
End Property
