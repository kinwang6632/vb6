VERSION 5.00
Object = "{92575699-8EB8-11D3-B585-0080C8F80BC4}#4.0#0"; "GiNumber.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{0CE6E756-92AE-11D3-8311-0080C8453DDF}#5.0#0"; "GiYM.ocx"
Object = "{A0C51418-555F-11D3-BC10-0050BA01A3A9}#9.8#0"; "GiMulti.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{074771B7-6ADE-11D5-B178-00485456B028}#2.0#0"; "csMulti.ocx"
Begin VB.Form frmSO3251A 
   BorderStyle     =   1  '單線固定
   Caption         =   "印單金額調整 [SO3251A]"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "新細明體"
      Size            =   9.75
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SO3251A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '螢幕中央
   Begin VB.TextBox txtSQL 
      Height          =   525
      Left            =   5340
      TabIndex        =   37
      Top             =   2820
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "F2. 確定"
      Height          =   375
      Left            =   7920
      TabIndex        =   19
      Top             =   5460
      Width           =   1365
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&X)"
      Height          =   375
      Left            =   10230
      TabIndex        =   20
      Top             =   5460
      Width           =   1365
   End
   Begin VB.Frame fraData 
      Height          =   5985
      Left            =   120
      TabIndex        =   21
      Top             =   30
      Width           =   11655
      Begin VB.Frame fraCustAllot 
         Caption         =   "客戶指定種類"
         ForeColor       =   &H00800080&
         Height          =   495
         Left            =   4530
         TabIndex        =   36
         Top             =   960
         Width           =   3825
         Begin VB.OptionButton optCustAllotAll 
            Caption         =   "全部"
            Height          =   195
            Left            =   2790
            TabIndex        =   6
            Top             =   240
            Value           =   -1  'True
            Width           =   915
         End
         Begin VB.OptionButton optCustAllotNo 
            Caption         =   "不指定"
            Height          =   195
            Left            =   1590
            TabIndex        =   5
            Top             =   240
            Width           =   915
         End
         Begin VB.OptionButton optCustAllotYes 
            Caption         =   "指定"
            Height          =   195
            Left            =   390
            TabIndex        =   4
            Top             =   240
            Width           =   915
         End
      End
      Begin CS_Multi.CSmulti gmsServEmp 
         Height          =   345
         Left            =   90
         TabIndex        =   9
         Top             =   2610
         Width           =   11475
         _ExtentX        =   20241
         _ExtentY        =   609
         ButtonCaption   =   "收費人員"
      End
      Begin Gi_Multi.GiMulti gimCMCode 
         Height          =   375
         Left            =   90
         TabIndex        =   31
         Top             =   2250
         Width           =   11475
         _ExtentX        =   20241
         _ExtentY        =   661
         ButtonCaption   =   "收費方式"
      End
      Begin Gi_Multi.GiMulti gmsClctArea 
         Height          =   375
         Left            =   90
         TabIndex        =   8
         Top             =   1905
         Width           =   11475
         _ExtentX        =   20241
         _ExtentY        =   661
         ButtonCaption   =   "收費區"
      End
      Begin Gi_Multi.GiMulti gmsServArea 
         Height          =   375
         Left            =   90
         TabIndex        =   7
         Top             =   1560
         Width           =   11475
         _ExtentX        =   20241
         _ExtentY        =   661
         ButtonCaption   =   "服務區"
      End
      Begin CS_Multi.CSmulti gmsStrtCode 
         Height          =   405
         Left            =   90
         TabIndex        =   12
         Top             =   3690
         Width           =   11475
         _ExtentX        =   20241
         _ExtentY        =   714
         ButtonCaption   =   "街道範圍"
      End
      Begin CS_Multi.CSmulti gmsMduId 
         Height          =   405
         Left            =   90
         TabIndex        =   11
         Top             =   3345
         Width           =   11475
         _ExtentX        =   20241
         _ExtentY        =   714
         ButtonCaption   =   "大樓名稱"
      End
      Begin VB.CommandButton cmdView 
         Caption         =   "瀏覽選取資料"
         Height          =   375
         Left            =   5940
         TabIndex        =   18
         Top             =   5430
         Visible         =   0   'False
         Width           =   1455
      End
      Begin MSMask.MaskEdBox mskPrtSNo2 
         Height          =   345
         Left            =   3120
         TabIndex        =   1
         ToolTipText     =   "YYYYMMxxxxxx"
         Top             =   630
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   609
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   12
         Mask            =   "999999999999"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskPrtSNo1 
         Height          =   345
         Left            =   1290
         TabIndex        =   0
         ToolTipText     =   "YYYYMMxxxxxx"
         Top             =   630
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   609
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   12
         Mask            =   "999999999999"
         PromptChar      =   "_"
      End
      Begin prjNumber.GiNumber gnbNewAmt 
         Height          =   315
         Left            =   3780
         TabIndex        =   16
         Top             =   4710
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
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
         MaxLength       =   10
         AutoSelect      =   -1  'True
         AllowMinus      =   -1  'True
         AllowZero       =   0   'False
      End
      Begin prjNumber.GiNumber gnbOldAmt 
         Height          =   315
         Left            =   1350
         TabIndex        =   14
         Top             =   4710
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
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
         MaxLength       =   10
         AutoSelect      =   -1  'True
         AllowMinus      =   -1  'True
         AllowZero       =   0   'False
      End
      Begin prjGiList.GiList gilCitemCode 
         Height          =   315
         Left            =   1350
         TabIndex        =   13
         Top             =   4230
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
         F2Corresponding =   -1  'True
      End
      Begin prjNumber.GiNumber gnbOldPeriod 
         Height          =   315
         Left            =   1350
         TabIndex        =   15
         Top             =   5190
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
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
         AutoSelect      =   -1  'True
      End
      Begin prjNumber.GiNumber gnbNewPeriod 
         Height          =   315
         Left            =   3780
         TabIndex        =   17
         Top             =   5190
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
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
         AutoSelect      =   -1  'True
      End
      Begin Gi_YM.GiYM gdaClctYM2 
         Height          =   375
         Left            =   2700
         TabIndex        =   3
         Top             =   1050
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   661
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
      Begin Gi_YM.GiYM gdaClctYM1 
         Height          =   375
         Left            =   1290
         TabIndex        =   2
         Top             =   1050
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   661
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
         Left            =   1290
         TabIndex        =   32
         Top             =   210
         Width           =   2925
         _ExtentX        =   5159
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
         FldWidth1       =   600
         FldWidth2       =   2000
         F5Corresponding =   -1  'True
      End
      Begin CS_Multi.CSmulti gimClassCode 
         Height          =   345
         Left            =   90
         TabIndex        =   10
         Top             =   2970
         Width           =   11475
         _ExtentX        =   20241
         _ExtentY        =   609
         ButtonCaption   =   "客戶類別"
      End
      Begin prjGiList.GiList gilServiceType 
         Height          =   315
         Left            =   5400
         TabIndex        =   34
         Top             =   210
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
         F3Corresponding =   -1  'True
      End
      Begin VB.Label lblServiceType 
         AutoSize        =   -1  'True
         Caption         =   "服務類別"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   4500
         TabIndex        =   35
         Top             =   270
         Width           =   780
      End
      Begin VB.Label lblComp 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "公  司  別"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   330
         TabIndex        =   33
         Top             =   300
         Width           =   765
      End
      Begin VB.Label lblNewPeriod 
         AutoSize        =   -1  'True
         Caption         =   "新出單期數"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   2730
         TabIndex        =   30
         Top             =   5280
         Width           =   975
      End
      Begin VB.Label lblNewAmt 
         AutoSize        =   -1  'True
         Caption         =   "新出單金額"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   2730
         TabIndex        =   29
         Top             =   4830
         Width           =   975
      End
      Begin VB.Label lblNote2A 
         AutoSize        =   -1  'True
         Caption         =   "收集年月"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   330
         TabIndex        =   28
         Top             =   1200
         Width           =   780
      End
      Begin VB.Label lblNote2B 
         AutoSize        =   -1  'True
         Caption         =   "至"
         Height          =   195
         Left            =   2340
         TabIndex        =   27
         Top             =   1170
         Width           =   195
      End
      Begin VB.Label lblOldPeriod 
         AutoSize        =   -1  'True
         Caption         =   "出單期數"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   330
         TabIndex        =   26
         Top             =   5280
         Width           =   780
      End
      Begin VB.Label lblOldAmt 
         AutoSize        =   -1  'True
         Caption         =   "出單金額"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   330
         TabIndex        =   25
         Top             =   4830
         Width           =   780
      End
      Begin VB.Label lblNote1B 
         AutoSize        =   -1  'True
         Caption         =   "至"
         Height          =   195
         Left            =   2760
         TabIndex        =   24
         Top             =   780
         Width           =   195
      End
      Begin VB.Label lblNote1A 
         AutoSize        =   -1  'True
         Caption         =   "印單序號"
         Height          =   195
         Left            =   330
         TabIndex        =   23
         Top             =   750
         Width           =   780
      End
      Begin VB.Label lblCitemCode 
         AutoSize        =   -1  'True
         Caption         =   "調整項目"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   330
         TabIndex        =   22
         Top             =   4365
         Width           =   780
      End
   End
End
Attribute VB_Name = "frmSO3251A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'本功能可將正式應收資料檔中(SO033), 符合條件的應收資料, 將其出單金額做批次調整
Option Explicit

Private Sub cmdCancel_Click()
    On Error GoTo ChkErr
        Unload Me
    Exit Sub
ChkErr:
    ErrSub Me.Name, "cmdCancel_Click"
End Sub

'‧  依據所下的參數, 串成SQL指令至後端做批次調整:
'‧  若原/新費用不相等且原/新期數不相等, SQL:
'    "update SO033 set ShouldAmt = <新費用>, OldPeriod = <新期數> where <各參數條件> and RealDate is null "
'‧  若只有原/新費用不相等, SQL:
'    "update SO033 set ShouldAmt = <新費用> where <各參數條件> and RealDate is null "
'‧  若只有原/新期數不相等, SQL:
'    "update SO033 set OldPeriod = <新期數> where <各參數條件> and RealDate is null"
'‧  若成功, 則顯示"調整完畢, 筆數=<被異動的資料筆數>"

Private Sub cmdOK_Click()
    On Error GoTo ChkErr
    Dim lngAffectCount As Long
    Dim strSQL As String
    Dim strUpdTime As String
        If Not IsDataOK Then Exit Sub
        strSQL = GetParaStr
        txtSQL = strSQL
        If MsgBox("確定調整?", vbQuestion + vbYesNo, "提示") = vbNo Then Exit Sub
        strUpdTime = GetDTString(RightNow)
        Screen.MousePointer = vbHourglass
        gcnGi.BeginTrans
        If ExecuteSQL("update " & GetOwner & "SO033 A set RealStopDate = Add_Months(RealStopDate," & (gnbNewPeriod.Value - gnbOldPeriod.Value) & ") ,UpdTime='" & strUpdTime & _
                "',Upden='" & garyGi(1) & "', ShouldAmt =" & NoZero(gnbNewAmt.Value, True) & "  , RealPeriod =" & NoZero(gnbNewPeriod.Value, True) & _
                "  where " & strSQL & "  and RealDate is null ", gcnGi, lngAffectCount) <> giOK Then MsgBox "連線失敗！請重新操作！", vbExclamation, "訊息！": GoTo ChkErr
        
        
'        If gnbOldAmt.Value <> gnbNewAmt.Value And gnbOldPeriod.Value <> gnbNewPeriod.Value Then
'            If ExecuteSQL("update SO033 set UpdTime='" & strUpdTime & "',Upden='" & garyGi(1) & "', ShouldAmt =" & NoZero(gnbNewAmt.Value, True) & "  , RealPeriod =" & NoZero(gnbNewPeriod.Value, True) & "  where " & strSQL & "  and RealDate is null ", gcnGi, lngAffectCount) <> giOK Then MsgBox "連線失敗！請重新操作！", vbExclamation, "訊息！"
'        ElseIf gnbOldAmt.Value <> gnbNewAmt.Value And gnbOldPeriod.Value = gnbNewPeriod.Value Then
'                    If ExecuteSQL("update SO033 set UpdTime='" & strUpdTime & "',Upden='" & garyGi(1) & "', ShouldAmt =" & NoZero(gnbNewAmt.Value, True) & " where " & strSQL & " and RealDate is null ", gcnGi, lngAffectCount) <> giOK Then MsgBox "連線失敗！請重新操作！", vbExclamation, "訊息！"
'            ElseIf gnbOldAmt.Value = gnbNewAmt.Value And gnbOldPeriod.Value <> gnbNewPeriod.Value Then
'                             If ExecuteSQL("update SO033 set UpdTime='" & strUpdTime & "',Upden='" & garyGi(1) & "', RealPeriod =" & NoZero(gnbNewPeriod.Value, True) & " where " & strSQL & " and RealDate is null", gcnGi, lngAffectCount) <> giOK Then MsgBox "連線失敗！請重新操作！", vbExclamation, "訊息！"
'        End If
        Screen.MousePointer = vbDefault
        gcnGi.CommitTrans
         MsgBox "調整完畢，筆數=" & IIf(lngAffectCount <= 0, 0, lngAffectCount), vbInformation, " 訊息！"
    Exit Sub
ChkErr:
    Screen.MousePointer = vbDefault
    gcnGi.RollbackTrans
    MsgBox "異動失敗！請重新操作！", vbExclamation, "訊息！"
    Unload Me
End Sub

Private Sub cmdView_Click()
'保留
    On Error GoTo ChkErr
    Dim rsView As New ADODB.Recordset
    Dim strSQL As String
        If Not IsDataOK Then Exit Sub
        strSQL = GetParaStr
'        rsView.CursorLocation = adUseClient
'        rsView.Open "Select CustId,CustName,Tel1,ChargeAddress,CustStatusName From " & GetOwner & "SO001 Where " & strSQL & " and RealDate is null", gcnGi, adOpenForwardOnly, adLockReadOnly
        If Not GetRS(rsView, "Select CustId,CustName,Tel1,ChargeAddress,CustStatusName From " & GetOwner & "SO001 Where " & strSQL & " and RealDate is null", gcnGi, adUseClient, adOpenForwardOnly, adLockReadOnly) Then Exit Sub
        With ViewForm
                .uRecordset = rsView
                .Show vbModal
        End With
    Exit Sub
ChkErr:
    ErrSub Me.Name, "cmdView_Click"
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo ChkErr
        If Not ChkGiList(KeyCode) Then Exit Sub
        If KeyCode = vbKeyF2 Then
            Call cmdOK_Click: Exit Sub
        ElseIf Shift = 2 Then
            If KeyCode = vbKeyF Then
                txtSQL.Move 0, 0, Me.Width, Me.Height / 2
                txtSQL.Visible = Not txtSQL.Visible
            End If
        End If
    Exit Sub
ChkErr:
    ErrSub Me.Name, "Form_KeyDown"
End Sub

Private Sub Form_Load()
    On Error Resume Next
        If Alfa2 Then
            Call GetGlobal
        End If
        Call subGim
        Call subGil
End Sub

Private Sub subGim()
    On Error Resume Next
    '收費方式
        SetgiMulti gimCMCode, "CodeNo", "Description", "CD031", "代碼", "名稱", , True
    '服務區
        SetgiMulti gmsServArea, "CodeNo", "Description", "CD002", "代碼", "名稱", , True
    '收費區
        SetgiMulti gmsClctArea, "CodeNo", "Description", "CD040", "代碼", "名稱", , True
        gimCMCode.Clear
    '客戶類別
        SetgiMulti gimClassCode, "CodeNo", "Description", "CD004", "代碼", "名稱", , True
        gimClassCode.Clear
    '收費人員
        SetgiMulti gmsServEmp, "EmpNo", "EmpName", "CM003", "代碼", "名稱", , True
    '大樓
        SetgiMulti gmsMduId, "Mduid", "Name", "SO017", "大樓編號", "大樓名稱"
    '街道
        SetgiMulti gmsStrtCode, "CodeNo", "Description", "CD017", "代碼", "名稱", , True
        
End Sub

Private Sub subGil()
    On Error Resume Next
        If Val(GetSystemParaItem("SecondDiscount", gCompCode, "", "SO041") & "") = 0 Then
            SetgiList gilCitemCode, "CodeNo", "Description", "CD019", , , , , 3, 12, " Where PeriodFlag=1", True
        Else
            '#XXXX 判斷是否有優惠組合 By Kin 2009/04/01
            SetgiList gilCitemCode, "CodeNo", "Description", "CD019", , , , , 3, 12, "Where Nvl(ServiceType,'" & gServiceType & "' ) ='" & gServiceType & "' And Nvl(RefNo,0) Not In (19,20) ", True
        End If
        gilCitemCode.Clear
        SetgiList gilCompCode, "CodeNo", "Description", "CD039", , , , , 3, 12, GetCompCodeFilter(gilCompCode)
        gilCompCode.SetCodeNo garyGi(9)
        gilCompCode.Query_Description
End Sub

Private Function IsDataOK() As Boolean
'2.  必要欄位: 收集年月範圍, 調整項目, 原/新費用, 原/新期數
'(註:原費用=新費用, 且原期數=新期數, 則為不過)
    On Error GoTo ChkErr
    IsDataOK = False
    Dim objText As Object
        If Not ChkDTok Then Exit Function
        If Not MustExist(gilCompCode, 2, "公司別") Then Exit Function
        If Not MustExist(gilServiceType, 2, "服務類別") Then Exit Function
        If Not MustExist(gdaClctYM1, 1, "收集年月起始日") Then Exit Function
        If Not MustExist(gdaClctYM2, 1, "收集年月截止日") Then Exit Function
        
        
        If gdaClctYM1.Text > gdaClctYM2.Text Then Call MsgDateRangeX("收集年月"): gdaClctYM1.SetFocus: Exit Function
        If Not MustExist(gilCitemCode, 2, "調整項目") Then Exit Function
        
        If Not MustExist(gnbOldAmt, , "原出單金額") Then Exit Function
        If Not MustExist(gnbNewAmt, , "新出單金額") Then Exit Function
        
        If Not MustExist(gnbOldPeriod, , "原出單期數") Then Exit Function
        If Not MustExist(gnbNewPeriod, , "新出單期數") Then Exit Function
        
        If gnbNewPeriod.Value > 12 Then MsgBox "此欄位不得大於12個月！", vbExclamation, "訊息！": gnbNewPeriod.SetFocus: Exit Function
        If gnbOldPeriod.Value > 12 Then MsgBox "此欄位不得大於12個月！", vbExclamation, "訊息！": gnbOldPeriod.SetFocus: Exit Function
        
        If gnbOldPeriod.Value = gnbNewPeriod.Value And gnbNewAmt.Value = gnbOldAmt.Value Then MsgBox "期數與金額皆相同無法調整！", vbExclamation, "訊息！": gnbNewPeriod.SetFocus: Exit Function
        IsDataOK = True
    Exit Function
ChkErr:
    Call ErrSub(Me.Name, "IsDataOk")
End Function

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
        Call ReleaseCOM(Me)

End Sub

Private Sub gdaClctYM1_Validate(Cancel As Boolean)
    If gdaClctYM1.Text <> "" Then
        Call gdaClctYM2.SetValue(gdaClctYM1.GetValue)
    End If
End Sub

Private Sub gilCompCode_Change()
    On Error Resume Next
        If Not ChgComp(gilCompCode, "SO3250", "SO3251") Then Exit Sub
        subGim
        subGil
        Call SelectServType(gilCompCode.GetCodeNo, gilServiceType)
        gilServiceType.SetCodeNo ""
        gilServiceType.Query_Description
        gilServiceType.ListIndex = 1
        Call GiMultiFilter(gmsServArea, , gilCompCode.GetCodeNo)
        Call GiMultiFilter(gmsClctArea, , gilCompCode.GetCodeNo)
        Call GiMultiFilter(gmsServEmp, , gilCompCode.GetCodeNo)
        Call GiMultiFilter(gmsMduId, , gilCompCode.GetCodeNo)
        Call GiMultiFilter(gmsStrtCode, , gilCompCode.GetCodeNo)
        

End Sub

Private Sub gilServiceType_Change()
    On Error Resume Next
        Call GiListFilter(gilCitemCode, gilServiceType.GetCodeNo)
        gilCitemCode.Filter = gilCitemCode.Filter & IIf(Trim(gilCitemCode.Filter) = "", " Where ", " And ") & " PeriodFlag = 1"
        Call GiMultiFilter(gimClassCode, gilServiceType.GetCodeNo)
        Call GiMultiFilter(gimCMCode, gilServiceType.GetCodeNo)
    
End Sub

Private Sub mskprtSNo1_GotFocus()
    mskPrtSNo1.SelStart = 0
    mskPrtSNo1.SelLength = Len(mskPrtSNo1.Text)
End Sub

Private Sub mskprtSNo2_GotFocus()
    mskPrtSNo2 = mskPrtSNo1
    mskPrtSNo2.SelStart = 0
    mskPrtSNo2.SelLength = Len(mskPrtSNo2.Text)
End Sub

Private Function GetParaStr() As String
    On Error GoTo ChkErr
    Dim strSQL As String
        If mskPrtSNo1.Text <> "" And mskPrtSNo2.Text <> "" Then
            strSQL = " And A.PrtSNo >='" & mskPrtSNo1.Text & "' And A.Prtsno  <='" & mskPrtSNo2.Text & "'"
        ElseIf mskPrtSNo1.Text <> "" And mskPrtSNo2.Text = "" Then
                strSQL = " And A.PrtSNo ='" & mskPrtSNo1.Text & "'"
            ElseIf mskPrtSNo1.Text = "" And mskPrtSNo2.Text <> "" Then
                strSQL = " And A.PrtSNo='" & mskPrtSNo2.Text & "' "
        End If
        
            
        strSQL = strSQL & " And ( A.ClctYM >=" & gdaClctYM1.Text & " And A.ClctYM <=" & gdaClctYM2.Text & " )"
        
        If gilCompCode.GetCodeNo <> "" Then
            strSQL = strSQL & " And  A.CompCode = " & gilCompCode.GetCodeNo
        End If
        
        If gilServiceType.GetCodeNo <> "" Then
            strSQL = strSQL & " And  A.ServiceType = '" & gilServiceType.GetCodeNo & "'"
        End If
        
        If gmsServArea.GetQryStr <> "" Then
            strSQL = strSQL & " And  A.ServCode " & gmsServArea.GetQryStr
        End If
        
        If gmsClctArea.GetQryStr <> "" Then
            strSQL = strSQL & " And A.ClctAreaCode " & gmsClctArea.GetQryStr
        End If
        
        If gimCMCode.GetQryStr <> "" Then
            strSQL = strSQL & " And A.CMCode " & gimCMCode.GetQryStr
        End If
        
        If gimClassCode.GetQryStr <> "" Then
            strSQL = strSQL & " And A.ClassCode " & gimClassCode.GetQryStr
        End If
        
        If gmsServEmp.GetQryStr <> "" Then
            strSQL = strSQL & " And A.ClctEn " & gmsServEmp.GetQryStr
        End If
        
        If gmsMduId.GetQryStr <> "" Then
            strSQL = strSQL & "  And A.MduId " & gmsMduId.GetQryStr
        End If
        
        If gmsStrtCode.GetQryStr <> "" Then
            strSQL = strSQL & " And A.StrtCode " & gmsStrtCode.GetQryStr
        End If
        
        If gilCitemCode.GetCodeNo <> "" Then
            strSQL = strSQL & " And A.CitemCode=" & gilCitemCode.GetCodeNo
        End If
        
        If gnbOldAmt.Text <> "" Then strSQL = strSQL & " And A.ShouldAmt=" & gnbOldAmt.Value
        
        If gnbOldPeriod.Text <> "" Then strSQL = strSQL & " And A.RealPeriod=" & gnbOldPeriod.Value
        
        '客戶指定有選 93/12/02 Jacky
        If optCustAllotAll.Value = False Then
            If optCustAllotYes Then strSQL = strSQL & " And B.CustAllot = 1 "
            If optCustAllotNo Then strSQL = strSQL & " And B.CustAllot = 0 "
            
            strSQL = " And A.RowId In (Select A.RowId From " & GetOwner & "SO033 A," & GetOwner & "SO003 B Where A.CompCode = B.CompCode And A.ServiceType = B.ServiceType And A.CustId = B.CustId And A.CitemCode=B.CitemCode And A.FaciSeqNo = B.FaciSeqNo " & strSQL & ")"
        End If
        
        strSQL = Mid(strSQL, 5)
        GetParaStr = strSQL
    Exit Function
ChkErr:
    Call ErrSub(Me.Name, "GetParaStr")
End Function
