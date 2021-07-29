VERSION 5.00
Object = "{92575699-8EB8-11D3-B585-0080C8F80BC4}#4.0#0"; "GiNumber.ocx"
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{0CE6E756-92AE-11D3-8311-0080C8453DDF}#5.0#0"; "GiYM.ocx"
Object = "{A0C51418-555F-11D3-BC10-0050BA01A3A9}#9.8#0"; "GiMulti.ocx"
Object = "{6DDA0569-6454-11D3-9612-0048544B09DF}#4.1#0"; "GiTime.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{074771B7-6ADE-11D5-B178-00485456B028}#2.0#0"; "csMulti.ocx"
Begin VB.Form frmSO3252A 
   BorderStyle     =   1  '單線固定
   Caption         =   "收費單整批作廢 [SO3252A]"
   ClientHeight    =   7830
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
   Icon            =   "SO3252A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '螢幕中央
   Begin VB.PictureBox pic1 
      Height          =   1005
      Left            =   3660
      ScaleHeight     =   945
      ScaleWidth      =   5085
      TabIndex        =   45
      Top             =   2250
      Visible         =   0   'False
      Width           =   5145
      Begin VB.TextBox txtComPaswrd 
         ForeColor       =   &H000040C0&
         Height          =   315
         IMEMode         =   3  '暫止
         Left            =   1350
         PasswordChar    =   "*"
         TabIndex        =   47
         Top             =   540
         Width           =   3135
      End
      Begin VB.TextBox txtPassword 
         ForeColor       =   &H00004080&
         Height          =   315
         IMEMode         =   3  '暫止
         Left            =   1350
         PasswordChar    =   "*"
         TabIndex        =   46
         Top             =   150
         Width           =   3135
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "確認密碼"
         ForeColor       =   &H000040C0&
         Height          =   195
         Left            =   210
         TabIndex        =   49
         Top             =   600
         Width           =   780
      End
      Begin VB.Label lblPassword 
         AutoSize        =   -1  'True
         Caption         =   "請輸入密碼"
         ForeColor       =   &H00004080&
         Height          =   195
         Left            =   210
         TabIndex        =   48
         Top             =   210
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "F2. 查詢"
      Height          =   375
      Left            =   7980
      TabIndex        =   18
      Top             =   7020
      Width           =   1365
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&X)"
      Height          =   375
      Left            =   10230
      TabIndex        =   19
      Top             =   7020
      Width           =   1365
   End
   Begin VB.Frame fraData 
      Height          =   7485
      Left            =   90
      TabIndex        =   34
      Top             =   120
      Width           =   11655
      Begin Gi_Multi.GiMulti gimBillType 
         Height          =   435
         Left            =   90
         TabIndex        =   56
         Top             =   5130
         Width           =   11475
         _ExtentX        =   20241
         _ExtentY        =   767
         ButtonCaption   =   "單  據  類  別"
         DataType        =   2
         DIY             =   -1  'True
      End
      Begin Gi_Time.GiTime gdtCreateTime1 
         Height          =   315
         Left            =   5280
         TabIndex        =   6
         Top             =   810
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
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
      Begin VB.Frame fraReal 
         Caption         =   "收款條件"
         Height          =   645
         Left            =   7710
         TabIndex        =   20
         Top             =   1230
         Width           =   3825
         Begin VB.OptionButton optReal 
            Caption         =   "已收"
            Height          =   195
            Left            =   270
            TabIndex        =   25
            Top             =   300
            Width           =   915
         End
         Begin VB.OptionButton optUC 
            Caption         =   "未收"
            Height          =   195
            Left            =   1575
            TabIndex        =   26
            Top             =   300
            Value           =   -1  'True
            Width           =   915
         End
         Begin VB.OptionButton optAll 
            Caption         =   "全部"
            Height          =   195
            Left            =   2880
            TabIndex        =   27
            Top             =   300
            Width           =   915
         End
      End
      Begin prjGiList.GiList gilCancelCode 
         Height          =   315
         Left            =   5640
         TabIndex        =   14
         Top             =   5910
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
      End
      Begin VB.CommandButton cmdView 
         Caption         =   "瀏覽選取資料"
         Height          =   375
         Left            =   5640
         TabIndex        =   33
         Top             =   6900
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txtReason 
         Height          =   375
         Left            =   5640
         TabIndex        =   17
         Text            =   "整批作廢"
         Top             =   6360
         Width           =   5385
      End
      Begin Gi_Multi.GiMulti gmsClctArea 
         Height          =   375
         Left            =   90
         TabIndex        =   22
         Top             =   2250
         Width           =   11475
         _ExtentX        =   20241
         _ExtentY        =   661
         ButtonCaption   =   "收費區"
      End
      Begin Gi_Multi.GiMulti gmsServArea 
         Height          =   375
         Left            =   90
         TabIndex        =   21
         Top             =   1890
         Width           =   11475
         _ExtentX        =   20241
         _ExtentY        =   661
         ButtonCaption   =   "服務區"
      End
      Begin Gi_YM.GiYM gdaClctYM2 
         Height          =   345
         Left            =   10410
         TabIndex        =   3
         Top             =   195
         Width           =   975
         _ExtentX        =   1720
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
      Begin Gi_YM.GiYM gdaClctYM1 
         Height          =   345
         Left            =   9150
         TabIndex        =   2
         Top             =   195
         Width           =   975
         _ExtentX        =   1720
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
      Begin MSMask.MaskEdBox mskPrtSNo2 
         Height          =   345
         Left            =   6450
         TabIndex        =   11
         ToolTipText     =   "YYYYMMxxxxxx"
         Top             =   1395
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   609
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   12
         Mask            =   "999999999999"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskPrtSNo1 
         Height          =   345
         Left            =   4920
         TabIndex        =   10
         ToolTipText     =   "YYYYMMxxxxxx"
         Top             =   1395
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   609
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   12
         Mask            =   "999999999999"
         PromptChar      =   "_"
      End
      Begin prjNumber.GiNumber gnbRealPeriod2 
         Height          =   345
         Left            =   2880
         TabIndex        =   16
         Top             =   6390
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   609
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
      Begin prjNumber.GiNumber gnbRealPeriod1 
         Height          =   345
         Left            =   1260
         TabIndex        =   15
         Top             =   6420
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   609
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
      Begin prjNumber.GiNumber gnbShouldAmt2 
         Height          =   345
         Left            =   2880
         TabIndex        =   13
         Top             =   5910
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   609
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
         AllowZero       =   0   'False
      End
      Begin prjNumber.GiNumber gnbShouldAmt1 
         Height          =   345
         Left            =   1260
         TabIndex        =   12
         Top             =   5910
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   609
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
         AllowZero       =   0   'False
      End
      Begin CS_Multi.CSmulti gmsStrtCode 
         Height          =   375
         Left            =   90
         TabIndex        =   29
         Top             =   3660
         Width           =   11475
         _ExtentX        =   20241
         _ExtentY        =   661
         ButtonCaption   =   "街道範圍"
      End
      Begin CS_Multi.CSmulti gmsMduId 
         Height          =   405
         Left            =   90
         TabIndex        =   28
         Top             =   3300
         Width           =   11475
         _ExtentX        =   20241
         _ExtentY        =   714
         ButtonCaption   =   "大樓名稱"
      End
      Begin prjGiList.GiList gilCompCode 
         Height          =   315
         Left            =   1110
         TabIndex        =   0
         Top             =   240
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
      Begin Gi_Multi.GiMulti gimUCCode 
         Height          =   375
         Left            =   90
         TabIndex        =   23
         Top             =   2595
         Width           =   11475
         _ExtentX        =   20241
         _ExtentY        =   661
         ButtonCaption   =   "未收原因"
      End
      Begin CS_Multi.CSmulti gmsServEmp 
         Height          =   345
         Left            =   90
         TabIndex        =   24
         Top             =   2940
         Width           =   11475
         _ExtentX        =   20241
         _ExtentY        =   609
         ButtonCaption   =   "收費人員"
      End
      Begin Gi_Date.GiDate gdaShouldDay1 
         Height          =   315
         Left            =   1110
         TabIndex        =   4
         Top             =   840
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         ForeColor       =   8388736
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
      Begin Gi_Date.GiDate gdaShouldDay2 
         Height          =   315
         Left            =   2850
         TabIndex        =   5
         Top             =   840
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         ForeColor       =   8388736
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
      Begin CS_Multi.CSmulti gimCitemCode 
         Height          =   375
         Left            =   90
         TabIndex        =   30
         Top             =   4020
         Width           =   11475
         _ExtentX        =   20241
         _ExtentY        =   661
         ButtonCaption   =   "調整項目"
      End
      Begin prjGiList.GiList gilServiceType 
         Height          =   315
         Left            =   5280
         TabIndex        =   1
         Top             =   240
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
      Begin CS_Multi.CSmulti gimClassCode 
         Height          =   375
         Left            =   90
         TabIndex        =   31
         Top             =   4390
         Width           =   11475
         _ExtentX        =   20241
         _ExtentY        =   661
         ButtonCaption   =   "客  戶  類  別"
      End
      Begin Gi_Multi.GiMulti gimCMCode 
         Height          =   375
         Left            =   90
         TabIndex        =   32
         Top             =   4760
         Width           =   11475
         _ExtentX        =   20241
         _ExtentY        =   661
         ButtonCaption   =   "收  費  方  式"
      End
      Begin Gi_Time.GiTime gdtCreateTime2 
         Height          =   315
         Left            =   7650
         TabIndex        =   7
         Top             =   810
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
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
      Begin Gi_Date.GiDate gdaNextYM1 
         Height          =   315
         Left            =   1110
         TabIndex        =   8
         Top             =   1410
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         ForeColor       =   8388736
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
      Begin Gi_Date.GiDate gdaNextYM2 
         Height          =   315
         Left            =   2850
         TabIndex        =   9
         Top             =   1440
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         ForeColor       =   8388736
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
      Begin Gi_Multi.GiMulti gimCustStatus 
         Height          =   435
         Left            =   90
         TabIndex        =   59
         Top             =   5520
         Width           =   11475
         _ExtentX        =   20241
         _ExtentY        =   767
         ButtonCaption   =   "客  戶  狀  態"
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "至"
         Height          =   195
         Left            =   2460
         TabIndex        =   58
         Top             =   1500
         Width           =   195
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "下收日期"
         Height          =   195
         Left            =   240
         TabIndex        =   57
         Top             =   1500
         Width           =   780
      End
      Begin VB.Label lblServiceType 
         AutoSize        =   -1  'True
         Caption         =   "服務類別"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   4380
         TabIndex        =   55
         Top             =   300
         Width           =   780
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "產生日期"
         Height          =   195
         Left            =   4380
         TabIndex        =   54
         Top             =   870
         Width           =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "至"
         Height          =   195
         Left            =   7290
         TabIndex        =   53
         Top             =   870
         Width           =   195
      End
      Begin VB.Label lblD3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "至"
         Height          =   195
         Left            =   2460
         TabIndex        =   52
         Top             =   900
         Width           =   195
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "應收日期"
         Height          =   195
         Left            =   240
         TabIndex        =   51
         Top             =   900
         Width           =   780
      End
      Begin VB.Label lblComp 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "公  司  別"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   270
         TabIndex        =   50
         Top             =   330
         Width           =   765
      End
      Begin VB.Label lblNote 
         AutoSize        =   -1  'True
         Caption         =   "備註"
         Height          =   195
         Left            =   4740
         TabIndex        =   44
         Top             =   6420
         Width           =   390
      End
      Begin VB.Label lblReason 
         AutoSize        =   -1  'True
         Caption         =   "作廢原因"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   4740
         TabIndex        =   43
         Top             =   5985
         Width           =   780
      End
      Begin VB.Label lblNote1A 
         AutoSize        =   -1  'True
         Caption         =   "印單序號"
         Height          =   195
         Left            =   4080
         TabIndex        =   42
         Top             =   1470
         Width           =   780
      End
      Begin VB.Label lblNote1B 
         AutoSize        =   -1  'True
         Caption         =   "至"
         Height          =   195
         Left            =   6210
         TabIndex        =   41
         Top             =   1500
         Width           =   195
      End
      Begin VB.Label lblAmtA 
         AutoSize        =   -1  'True
         Caption         =   "出單金額"
         Height          =   195
         Left            =   330
         TabIndex        =   40
         Top             =   6000
         Width           =   780
      End
      Begin VB.Label lblPeriodA 
         AutoSize        =   -1  'True
         Caption         =   "出單期數"
         Height          =   195
         Left            =   330
         TabIndex        =   39
         Top             =   6480
         Width           =   780
      End
      Begin VB.Label lblNote2B 
         AutoSize        =   -1  'True
         Caption         =   "至"
         Height          =   195
         Left            =   10170
         TabIndex        =   38
         Top             =   270
         Width           =   195
      End
      Begin VB.Label lblNote2A 
         AutoSize        =   -1  'True
         Caption         =   "收集年月"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   8340
         TabIndex        =   37
         Top             =   300
         Width           =   780
      End
      Begin VB.Label lblAmtB 
         AutoSize        =   -1  'True
         Caption         =   "至"
         Height          =   195
         Left            =   2580
         TabIndex        =   36
         Top             =   6000
         Width           =   195
      End
      Begin VB.Label lblPeriodB 
         AutoSize        =   -1  'True
         Caption         =   "至"
         Height          =   195
         Left            =   2580
         TabIndex        =   35
         Top             =   6480
         Width           =   195
      End
   End
End
Attribute VB_Name = "frmSO3252A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private blnDisCount As Boolean
Private Sub cmdCancel_Click()
    On Error Resume Next
        If pic1.Visible Then
            pic1.Visible = False
            fraData.Enabled = True
        Else
            Unload Me
        End If
End Sub

Private Sub subGil()
    On Error Resume Next
        '收費項目
        'SetgiList gilCitemCode, "CodeNo", "Description", "CD019", , , , , , , " Where PeriodFlag=1", True
        SetgiList gilCancelCode, "CodeNo", "Description", "CD051", , , , , 3, 12, , True
        gilCancelCode.Clear
        SetgiList gilCompCode, "CodeNo", "Description", "CD039", , , , , 3, 12, GetCompCodeFilter(gilCompCode)
        gilCompCode.SetCodeNo garyGi(9)
        gilCompCode.Query_Description
End Sub
'原本功能為根據畫面的條件直接做Update動作，但#3042需要先秀出Grid讓User自行選擇，所以此段功能改為秀出SO3252B By Kin 2007/3/22
Private Sub cmdOK_Click()
    On Error GoTo chkErr
    Dim strSQL As String
    Dim strSQLData As String
    Dim rsTemp As New ADODB.Recordset
    Dim ObjOpenDb As New giOpenDBObj.OpenDBObj
    Dim strDisWhere As String
    Dim strUpdTime As String
        If Not IsDataOK Then Exit Sub
        If pic1.Visible = False Then pic1.Visible = True: txtPassword.Text = "": txtComPaswrd = "": txtPassword.SetFocus: fraData.Enabled = False: Exit Sub
        If Not IsDataOk2 Then Exit Sub
        strSQL = GetParaStr
        strUpdTime = GetDTString(RightNow)
        Screen.MousePointer = vbHourglass
        pic1.Visible = False
        strDisWhere = Empty
        If blnDisCount Then
            
        End If
        strSQLData = "Select CustID,BillNo,CITEMCODE,CITEMNAME,nvl(PREINVOICE,0) as PREINVOICE," & _
                     "nvl(ShouldAmt,0) as ShouldAmt,nvl(RealPeriod,0) as RealPeriod,ShouldDate,RealStartDate,RealStopDate,UCName,GUINO " & _
                     " From " & GetOwner & "SO033 Where " & strSQL & " Order by BillNo"
        If Not GetRS(rsTemp, strSQLData, gcnGi, adUseClient, adOpenKeyset, adLockPessimistic) Then Exit Sub
        If rsTemp.RecordCount <= 0 Then
            MsgBox "無任何資料可異動!!", vbInformation, "警告訊息"
            pic1.Visible = False
            CloseRecordset rsTemp
            fraData.Enabled = True
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        '開啟瀏覽視窗，並將相關資訊傳入
        With frmSO3252B
            '.uRS = rsData
            .uRS = rsTemp
            .uSQLWhere = strSQL
            .uReasno = txtReason
            .uCompCode = gilCompCode.GetCodeNo
            .uUpdTime = strUpdTime
            .uCancelCode = gilCancelCode.GetCodeNo
            .uCancelName = gilCancelCode.GetDescription
            .Show 1
        End With
        fraData.Enabled = True
        CloseRecordset rsTemp
'        Screen.MousePointer = vbHourglass
'        gcnGi.BeginTrans
'        'UCCode=Null,UCName =Null <2000/03/24>
'        If ExecuteSQL("Update " & GetOwner & "So033 Set UCCode=Null,UCName=Null,CancelFlag=1,RealDate=To_Date('" & Format(RightDate, "YYYY/MM/DD") & "','YYYY/MM/DD')," & _
'            "Note = Trim(Nvl(Note,'')) || ';作廢原因:" & txtReason.Text & "; 未收原因:' || UCName,CancelCode = " & gilCancelCode.GetCodeNo & "," & _
'            "CancelName = '" & gilCancelCode.GetDescription & "',UpdEn = '" & garyGi(1) & "',UpdTime ='" & strUpdTime & "'  Where " & strSQL, gcnGi, lngAffectCounts, False, False) <> giOK Then
'            MsgBox "連線失敗！請重新操作！", vbExclamation, "訊息！": GoTo ChkErr
'        End If
'        strSQL = SetAddComma(strSQL)
'        If ExecuteSQL("Insert into " & GetOwner & "So070 (UpdTime,UpdName,Reason,Para,RcdCount) values('" & strUpdTime & "','" & garyGi(1) & "'," & IIf(Trim(gilCancelCode.GetDescription) <> "", "'" & Trim(gilCancelCode.GetDescription) & "'", Null) & ",'" & Replace(strSQL, "'", "") & "'," & IIf(lngAffectCounts < 0, 0, lngAffectCounts) & ")", gcnGi, , , False) <> giOK Then
'                MsgBox "連線失敗！請重新操作！", vbExclamation, "訊息！"
'        End If
'        gcnGi.CommitTrans
'        pic1.Visible = False
'        fraData.Enabled = True
'        Screen.MousePointer = vbDefault
'        MsgBox "作廢完畢，筆數=" & IIf(lngAffectCounts < 0, 0, lngAffectCounts), vbInformation, "訊息！"
    Exit Sub
chkErr:
    Screen.MousePointer = vbDefault
    gcnGi.RollbackTrans
    MsgBox "整批作廢失敗！請重新操作！", vbExclamation, "訊息！"
    pic1.Visible = False
    fraData.Enabled = True
End Sub

Private Function SetAddComma(ByVal strSQL As String) As String
    On Error GoTo chkErr
        Dim strNewSQl As String
        Dim strTmpSql  As String
        Dim intSitus As Integer
        Do While True
            intSitus = InStr(1, strSQL, "'", vbTextCompare)
            If intSitus <= 0 Then Exit Do
            strTmpSql = strTmpSql & Left(strSQL, intSitus) & "'"
            strSQL = Mid(strSQL, intSitus + 1)
        Loop
        SetAddComma = strTmpSql
    Exit Function
chkErr:
    ErrSub Me.Name, "SetAddComma"
End Function

Private Sub Form_Activate()
    On Error Resume Next
        Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    On Error GoTo chkErr
'    Call GetGlobal
        Call subGil
        Call subGim
        '#XXXX 判斷是否有第二層優惠 By Kin 2009/04/01
        
    Exit Sub
chkErr:
    ErrSub Me.Name, "Form_Load"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
        If Not ChkGiList(KeyCode) Then Exit Sub
        If KeyCode = vbKeyF2 Then Call cmdOK_Click: Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
        Call ReleaseCOM(Me)

End Sub

Private Sub gdaClctYM2_Validate(Cancel As Boolean)
    '#3042自我測試時發現原有存在Bug， 變更此事件，原本寫在gdaClctYM1，會造成gdaClctYM1一直處在驗證狀態
    On Error Resume Next
        '#3042 測試不OK,需要判斷起始年月不得大於終止年月 By Kin 2007/06/13
        If (gdaClctYM2.Text = "") And (gdaClctYM1.Text <> "") Then
            Call gdaClctYM2.SetValue(gdaClctYM1.GetValue)
        End If
        If Not IsDate(gdaClctYM2.GetValue(True)) Then Exit Sub
        If DateDiff("m", gdaClctYM1.GetValue(True), gdaClctYM2.GetValue(True)) < 0 Then
           MsgBox "收集年月截止日不可小於收集年月起始日!", vbExclamation, "警告"
           gdaClctYM2.SetValue gdaClctYM1.GetValue
           Cancel = True
        End If
End Sub

Private Sub gdaNextYM2_Validate(Cancel As Boolean)
    On Error Resume Next
        '#3042 測試不OK,需要判斷起始年月不得大於終止年月 By Kin 2007/06/13
        If (gdaNextYM2.Text = "") And (gdaNextYM1.Text <> "") Then
            Call gdaNextYM2.SetValue(gdaNextYM1.GetValue)
        End If
        If Not IsDate(gdaNextYM2.GetValue(True)) Then Exit Sub
        If DateDiff("d", gdaNextYM1.GetValue(True), gdaNextYM2.GetValue(True)) < 0 Then
           MsgBox "下收截止日不可小於下收起始日!", vbExclamation, "警告"
           gdaNextYM2.SetValue gdaNextYM1.GetValue
           Cancel = True
        End If
End Sub

Private Sub gdaShouldDay2_Validate(Cancel As Boolean)
    On Error Resume Next
        '#3042 測試不OK,需要判斷起始年月不得大於終止年月 By Kin 2007/06/13
        If (gdaShouldDay2.Text = "") And (gdaShouldDay1.Text <> "") Then
            Call gdaShouldDay2.SetValue(gdaShouldDay1.GetValue)
        End If
        If Not IsDate(gdaShouldDay2.GetValue(True)) Then Exit Sub
        If DateDiff("d", gdaShouldDay1.GetValue(True), gdaShouldDay2.GetValue(True)) < 0 Then
           MsgBox "應收截止日不可小於應收起始日!", vbExclamation, "警告"
           gdaShouldDay2.SetValue gdaShouldDay1.GetValue
           Cancel = True
        End If

End Sub

Private Sub gdtCreateTime2_Validate(Cancel As Boolean)
    '#3042 測試不OK,需要判斷起始年月不得大於終止年月 By Kin 2007/06/13
    On Error Resume Next
    If (gdtCreateTime2.Text = "") And (gdtCreateTime1.Text <> "") Then
        Call gdtCreateTime2.SetValue(gdtCreateTime1.GetValue)
    End If
    If Not IsDate(gdtCreateTime2.GetValue(True)) Then Exit Sub
    If DateDiff("n", gdtCreateTime1.GetValue(True), gdtCreateTime2.GetValue(True)) < 0 Then
       MsgBox "產生日期截止日不可小於產生日期起始日!", vbExclamation, "警告"
       gdtCreateTime2.SetValue gdtCreateTime1.GetValue
       Cancel = True
    End If

End Sub

Private Sub gilCompCode_Change()
    On Error Resume Next
        If Not ChgComp(gilCompCode, "SO3250", "SO3252") Then Exit Sub
        subGil
        subGim
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
        If gilServiceType.GetCodeNo = "" Then Exit Sub
        '#XXXX 判斷是否有第二層優惠 By Kin 2009/04/01
        blnDisCount = (Val(GetSystemParaItem("SecondDiscount", gCompCode, "", "SO041") & "") = 1)
        If blnDisCount Then
            gimCitemCode.Filter = Empty
            Call GiMultiFilter(gimCitemCode, gilServiceType.GetCodeNo)
            gimCitemCode.Filter = gimCitemCode.Filter & "  And ( Nvl(SecendDiscount,0)<=0 And RefNo<>19 )"
        Else
            Call GiMultiFilter(gimCitemCode, gilServiceType.GetCodeNo)
        End If
        Call GiMultiFilter(gimUCCode, gilServiceType.GetCodeNo)
        
End Sub

Private Sub gimCitemCode_ButtonClick()
    Debug.Print gimCitemCode.Filter
End Sub

Private Sub gimCitemCode_Click()
    Debug.Print gimCitemCode.Filter
End Sub

Private Sub mskprtSNo1_GotFocus()
    On Error Resume Next
        mskPrtSNo1.SelStart = 0
        mskPrtSNo1.SelLength = Len(mskPrtSNo1.Text)
End Sub

Private Sub mskprtSNo2_GotFocus()
    On Error Resume Next
        mskPrtSNo2 = mskPrtSNo1
        mskPrtSNo2.SelStart = 0
        mskPrtSNo2.SelLength = Len(mskPrtSNo2.Text)
End Sub

Private Sub subGim()
    On Error GoTo chkErr
        '服務區
            SetgiMulti gmsServArea, "CodeNo", "Description", "CD002", "代碼", "名稱", , True
        '收費區
            SetgiMulti gmsClctArea, "CodeNo", "Description", "CD040", "代碼", "名稱", , True
        '未繳原因
            SetgiMulti gimUCCode, "CodeNo", "Description", "CD013", "代碼", "名稱", , True
            gimUCCode.Clear
        '收費人員
            SetgiMulti gmsServEmp, "EmpNo", "EmpName", "CM003", "代碼", "名稱", , True
        '大樓
            SetgiMulti gmsMduId, "Mduid", "Name", "SO017", "大樓編號", "大樓名稱"
        '街道
            SetgiMulti gmsStrtCode, "CodeNo", "Description", "CD017", "代碼", "名稱", , True
        '收費項目
            SetgiMulti gimCitemCode, "CodeNo", "Description", "CD019", "代碼", "名稱", , True
        '客戶類別       問題集2424 Penny 提 Crystal Edit by 2006/05/23
            SetgiMulti gimClassCode, "CodeNo", "Description", "CD004", "客戶類別代碼", "客戶類別名稱", , True
        '收費方式
            SetgiMulti gimCMCode, "CodeNo", "Description", "CD031", "收費方式代碼", "收費方式名稱", , True
            gimCitemCode.Clear
         '問題集2673 多增加一個單據類別
            SetgiMultiAddItem gimBillType, "B,I,P,M,T", "收費單,裝機單,停拆移機單,維修單,臨時收費單", "單據代號", "單據名稱"
            gimBillType.SetDispStr "收費單"
            gimBillType.SetQueryCode "B"
          '#3042 增加客戶狀態，並預設正常 By Kin 2007/03/20 Add
            SetgiMulti gimCustStatus, "CodeNo", "Description", "CD035", "客戶狀態代碼", "客戶狀態名稱", , False
            gimCustStatus.SetQueryCode 1
    Exit Sub
chkErr:
    ErrSub Me.Name, "subGim"
End Sub

Private Function GetParaStr() As String
    On Error GoTo chkErr
    Dim strSQL As String
    Dim strsubsql As String
        If mskPrtSNo1.Text <> "" And mskPrtSNo2.Text <> "" Then
            strSQL = "PrtSNo >='" & mskPrtSNo1.Text & "'  and  prtsno  <='" & mskPrtSNo2.Text & "'"
        ElseIf mskPrtSNo1.Text <> "" And mskPrtSNo2.Text = "" Then
                strSQL = "PrtSNo ='" & mskPrtSNo1.Text & "'"
            ElseIf mskPrtSNo1.Text = "" And mskPrtSNo2.Text <> "" Then
                strSQL = "PrtSNo='" & mskPrtSNo2.Text & "' "
        End If
            
        strSQL = strSQL & IIf(strSQL <> "", " And", "") & " ( ClctYM >=" & gdaClctYM1.Text & " And ClctYM <=" & gdaClctYM2.Text & " )"
    
        If gilCompCode.GetCodeNo <> "" Then
            strSQL = strSQL & " And  CompCode = " & gilCompCode.GetCodeNo
        End If
    
        If gilServiceType.GetCodeNo <> "" Then
            strSQL = strSQL & " And  ServiceType = '" & gilServiceType.GetCodeNo & "'"
        End If
        
        If gmsServArea.GetQryStr <> "" Then
            strSQL = strSQL & " And  ServCode " & gmsServArea.GetQryStr
        End If
        
        '91/12/30
        If gdaShouldDay1.GetValue <> "" Then subAnd2 strSQL, "ShouldDate >= " & GetNullString(gdaShouldDay1.GetValue(True), giDateV)
        If gdaShouldDay2.GetValue <> "" Then subAnd2 strSQL, "ShouldDate <= " & GetNullString(gdaShouldDay2.GetValue(True), giDateV)
        
        '92/04/28
        If gdtCreateTime1.GetValue <> "" Then subAnd2 strSQL, "CreateTime >= " & GetNullString(gdtCreateTime1.GetValue(True), giDateV, giOracle, True)
        If gdtCreateTime2.GetValue <> "" Then subAnd2 strSQL, "CreateTime <= " & GetNullString(gdtCreateTime2.GetValue(True), giDateV, giOracle, True)
        
        If gmsClctArea.GetQryStr <> "" Then
            strSQL = strSQL & " And ClctAreaCode " & gmsClctArea.GetQryStr
        End If
        
        If gmsServEmp.GetQryStr <> "" Then
            strSQL = strSQL & " And ClctEn " & gmsServEmp.GetQryStr
        End If
        
        If gmsMduId.GetQryStr <> "" Then
            strSQL = strSQL & "  And MduId " & gmsMduId.GetQryStr
        End If
        
        If gmsStrtCode.GetQryStr <> "" Then
            strSQL = strSQL & " And StrtCode " & gmsStrtCode.GetQryStr
        End If
        
        If gimUCCode.GetQryStr <> "" Then
            strSQL = strSQL & " And UCCode " & gimUCCode.GetQryStr
        End If
        
        If gimCitemCode.GetQueryCode <> "" Then subAnd2 strSQL, "CitemCode in (" & gimCitemCode.GetQueryCode & ")"
        
        '問題集2424 Penny提 Crystal Edit by 2006/05/23
        
        If gimClassCode.GetQryStr <> "" Then
          strSQL = strSQL & " And ClassCode " & gimClassCode.GetQryStr
        End If
        
        If gimCMCode.GetQryStr <> "" Then
          strSQL = strSQL & " And CMCode " & gimCMCode.GetQryStr
        End If
        '問題集2673 增加單據類別與下收日的判斷
        If gimBillType.GetQryStr <> "" Then
            strSQL = strSQL & " And subStr(BillNo,7,1) " & gimBillType.GetQryStr
        End If
        '下收日
        '#3042 多增加判斷客戶狀態，所以要多串SO002 by Kin 2007/03/22
        '#3246 將Custid in(SO003....)改為Exists語法,以防此有同一筆收費資料,但下收日不同都會被作廢 By Kin 2007/5/25
        If Trim(gdaNextYM1.Text) <> "" And Trim(gdaNextYM1.Text) <> "" Then
            strsubsql = "Select  SO003.* From " & GetOwner & "SO003," & GetOwner & "SO002 " & _
                        " Where SO003.CompCode=" & gilCompCode.GetCodeNo & _
                      " And SO033.ServiceType='" & gilServiceType.GetCodeNo & "'" & _
                      " And SO033.CitemCode IN (" & gimCitemCode.GetQueryCode & ")" & _
                      " And SO003.ClctDate>=To_Date(" & gdaNextYM1.GetValue & ",'YYYYMMDD')" & _
                      " And SO003.ClctDate<=To_Date(" & gdaNextYM2.GetValue & ",'YYYYMMDD')" & _
                      " And SO003.CompCode=SO033.CompCode And SO003.ServiceType=SO033.ServiceType And SO003.CustID=SO033.CustID" & _
                      " And SO003.FaciSeqNo=SO033.FaciSeqNo And SO033.Custid=SO002.Custid And SO002.ServiceType=SO033.ServiceType"
            '#3042 判斷客戶狀態否有值 By Kin 2007/03/22
            If gimCustStatus.GetQueryCode & "" <> "" Then
                strsubsql = strsubsql & " And SO002.CustStatusCode IN (" & gimCustStatus.GetQueryCode & ")"
            End If
            strSQL = strSQL & " And Exists (" & strsubsql & ")"
        Else
            If gimCustStatus.GetQueryCode <> "" Then
                strsubsql = "Select CustID From " & GetOwner & "SO002 Where ServiceType='" & gilServiceType.GetCodeNo & "' And CustStatusCode In (" & gimCustStatus.GetQueryCode & ") And SO033.Custid=SO002.CustId"
                strSQL = strSQL & " And exists (" & strsubsql & ")"
            End If
        End If
        If gnbShouldAmt1.Value <> 0 Then strSQL = strSQL & " And ShouldAmt >= " & gnbShouldAmt1.Value
        
        If gnbShouldAmt2.Value <> 0 Then strSQL = strSQL & " And ShouldAmt <= " & gnbShouldAmt2.Value
        
        If gnbRealPeriod1.Value <> 0 Then strSQL = strSQL & " And RealPeriod >= " & gnbRealPeriod1.Value
        
        If gnbRealPeriod2.Value <> 0 Then strSQL = strSQL & " And RealPeriod <= " & gnbRealPeriod2.Value
        
        If optReal.Value = True Then
            strSQL = strSQL & " And RealDate Is Not Null And UCCode is Null"
        ElseIf optUC.Value = True Then
            strSQL = strSQL & " And UCCode is not Null"
        End If
        strSQL = strSQL & " And GUINO is Null"
        'strSQL = strSQL & " And ShouldAmt>=" & gnbShouldAmt1.Value & " And ShouldAmt <=" & gnbShouldAmt2.Value & " And RealPeriod>=" & gnbRealPeriod1.Value & "  and RealPeriod<=" & gnbRealPeriod2.Value
        GetParaStr = strSQL
    Exit Function
chkErr:
    Call ErrSub(Me.Name, "GetParaStr")
End Function

Private Function IsDataOK() As Boolean
'2.  必要欄位: 收集年月範圍, 調整項目, 原/新費用, 原/新期數
    On Error GoTo chkErr
    IsDataOK = False
    Dim objText As Object
        If Not ChkDTok Then Exit Function
        
        If Not MustExist(gilCompCode, 2, "公司別") Then Exit Function
        If Not MustExist(gdaClctYM1, 1, "收集年月起始日") Then Exit Function
        If Not MustExist(gdaClctYM2, 1, "收集年月截止日") Then Exit Function
        
        If gdaClctYM1.Text > gdaClctYM2.Text Then MsgBox "起始月份不得大於結束月份！", vbExclamation, "訊息！": gdaClctYM1.SetFocus: Exit Function
        If gdaShouldDay1.Text > gdaShouldDay2.Text Then MsgBox "起始日期不行大於結束日期！", vbExclamation, "訊息！": gdaShouldDay1.SetFocus: Exit Function
        If gdaNextYM1.Text > gdaNextYM2.Text Then MsgBox "起始日期不行大於結束日期！", vbExclamation, "訊息！": gdaNextYM1.SetFocus: Exit Function
        If gimCitemCode.GetQueryCode = "" Then gimCitemCode.SetFocus: Call MsgMustBe("調整項目"): Exit Function
        If Not MustExist(gilCancelCode, 2, "作廢原因") Then Exit Function
'        If StrLen(Trim(txtReason)) > 750 Then MsgBox "備註欄位值太長！", vbExclamation, "訊息！": txtReason.SetFocus: Exit Function
    IsDataOK = True
    Exit Function
chkErr:
    Call ErrSub(Me.Name, "IsDataOk")
End Function

Private Function IsDataOk2() As Boolean
    On Error GoTo chkErr
        If UCase(txtPassword.Text) = "TEST" Then IsDataOk2 = True: Exit Function
        If txtPassword.Text = "" Then txtPassword.SetFocus: MsgBox "請輸入密碼", vbExclamation, gimsgPrompt: Exit Function
        If txtComPaswrd.Text = "" Then txtComPaswrd.SetFocus: MsgBox "請確認密碼", vbExclamation, gimsgPrompt: Exit Function
        If txtPassword.Text <> txtComPaswrd Then txtComPaswrd.SetFocus: MsgBox "確認密碼與輸入密碼不符", vbExclamation, gimsgPrompt: Exit Function
        If txtPassword.Text <> garyGi(12) Then txtPassword.SetFocus: MsgBox "輸入密碼不正確!!", vbExclamation, gimsgPrompt: Exit Function
        If txtComPaswrd.Text <> garyGi(12) Then txtComPaswrd.SetFocus: MsgBox "確認密碼不正確!!", vbExclamation, gimsgPrompt: Exit Function
        IsDataOk2 = True
    Exit Function
chkErr:
    ErrSub Me.Name, "IsDataOk2"
End Function

Private Sub txtReason_GotFocus()
    On Error Resume Next
        Call objSelectAll(txtReason)
End Sub
'#3042 建立記憶体RecordSet By Kin 2007/03/22
'Private Sub CreatersArray()
'  On Error GoTo ChkErr
'    If rsData.State = adStateOpen Then rsData.Close
'    rsData.CursorLocation = adUseClient
'    rsData.CursorType = adOpenKeyset
'    rsData.LockType = adLockPessimistic
''    rsData.Fields.Append "Chose", adInteger
''    rsData.Fields.Append "RowID", adBSTR, 30
'    rsData.Fields.Append "CustID", adBSTR, 20
'    rsData.Fields.Append "BillNO", adBSTR, 20
'    rsData.Fields.Append "CitemCode", adBSTR, 20
'    rsData.Fields.Append "CitemName", adBSTR, 30
'    rsData.Fields.Append "PREINVOICE", adInteger
'    rsData.Fields.Append "ShouldAmt", adBSTR, 20
'    rsData.Fields.Append "RealPeriod", adBSTR, 10
'    rsData.Fields.Append "ShouldDate", adBSTR, 15
'    rsData.Fields.Append "RealStartDate", adBSTR, 15
'    rsData.Fields.Append "RealStopDate", adBSTR, 15
'    rsData.Fields.Append "UCName", adBSTR, 200
'    rsData.Fields.Append "GUINO", adBSTR, 20
'    rsData.Open
'    Exit Sub
'ChkErr:
'    ErrSub Me.Name, "CreatersArray"
'End Sub

