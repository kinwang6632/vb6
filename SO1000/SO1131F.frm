VERSION 5.00
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{92575699-8EB8-11D3-B585-0080C8F80BC4}#4.0#0"; "GiNumber.ocx"
Object = "{3BB22E57-824A-11D3-BD70-0080C8F80BC4}#3.4#0"; "GiGridR.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Begin VB.Form frmSO1131F 
   BorderStyle     =   1  '單線固定
   Caption         =   "階段式優惠資料 [SO1131F]"
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   4035
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
   Icon            =   "SO1131F.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCopyIns 
      Caption         =   "F12.新增複製"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4620
      TabIndex        =   63
      Top             =   7470
      Width           =   1275
   End
   Begin VB.CommandButton cmdLog 
      Caption         =   "修改歷程"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5910
      TabIndex        =   62
      Top             =   7470
      Width           =   1095
   End
   Begin VB.TextBox txtNote 
      Height          =   1515
      Left            =   1770
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  '垂直捲軸
      TabIndex        =   15
      Top             =   4020
      Width           =   9195
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1995
      Left            =   5910
      TabIndex        =   48
      Top             =   4530
      Visible         =   0   'False
      Width           =   4755
      Begin VB.ComboBox cboPenalType 
         Height          =   315
         ItemData        =   "SO1131F.frx":0442
         Left            =   1650
         List            =   "SO1131F.frx":0444
         Locked          =   -1  'True
         TabIndex        =   50
         Top             =   360
         Visible         =   0   'False
         Width           =   3525
      End
      Begin VB.ComboBox cboExpireType 
         Height          =   315
         ItemData        =   "SO1131F.frx":0446
         Left            =   1650
         List            =   "SO1131F.frx":0448
         Locked          =   -1  'True
         TabIndex        =   49
         Top             =   720
         Visible         =   0   'False
         Width           =   3525
      End
      Begin prjNumber.GiNumber ginDeposit 
         Height          =   315
         Left            =   1650
         TabIndex        =   51
         Top             =   0
         Visible         =   0   'False
         Width           =   1155
         _ExtentX        =   2037
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
      End
      Begin prjNumber.GiNumber ginPenalAmt 
         Height          =   315
         Left            =   4020
         TabIndex        =   52
         Top             =   0
         Visible         =   0   'False
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         Alignment       =   0
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
      End
      Begin VB.Label Label13 
         Caption         =   "違約時之計價依據"
         Height          =   195
         Left            =   0
         TabIndex        =   56
         Top             =   420
         Visible         =   0   'False
         Width           =   1560
      End
      Begin VB.Label Label14 
         Caption         =   "優惠到期計價依據"
         Height          =   195
         Left            =   30
         TabIndex        =   55
         Top             =   900
         Visible         =   0   'False
         Width           =   1560
      End
      Begin VB.Label Label4 
         Caption         =   "履約保證金"
         Height          =   195
         Left            =   585
         TabIndex        =   54
         Top             =   60
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "違約金"
         Height          =   195
         Left            =   3180
         TabIndex        =   53
         Top             =   60
         Visible         =   0   'False
         Width           =   585
      End
   End
   Begin VB.TextBox txtPromtion 
      Height          =   1515
      Left            =   1770
      MultiLine       =   -1  'True
      ScrollBars      =   2  '垂直捲軸
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2460
      Width           =   4065
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "F2. 存檔..."
      Height          =   375
      Left            =   180
      TabIndex        =   20
      Top             =   7470
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "F11.修改"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3510
      TabIndex        =   23
      Top             =   7470
      Width           =   1095
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "F10. 停用"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2400
      TabIndex        =   22
      Top             =   7470
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "F6. 新增..."
      Height          =   375
      Left            =   1290
      TabIndex        =   21
      Top             =   7470
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&X)"
      Height          =   345
      Left            =   9630
      TabIndex        =   24
      Top             =   7500
      Width           =   1215
   End
   Begin VB.Frame fraData 
      Enabled         =   0   'False
      Height          =   5685
      Left            =   150
      TabIndex        =   25
      Top             =   -30
      Width           =   11625
      Begin VB.Frame fraUpdEn 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1635
         Left            =   5760
         TabIndex        =   43
         Top             =   2370
         Visible         =   0   'False
         Width           =   5190
         Begin VB.TextBox txtNReqNotes 
            Height          =   480
            Left            =   1575
            MaxLength       =   500
            MultiLine       =   -1  'True
            ScrollBars      =   2  '垂直捲軸
            TabIndex        =   18
            Top             =   1020
            Width           =   3510
         End
         Begin VB.TextBox txtNReqDep 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1575
            MaxLength       =   20
            TabIndex        =   16
            Top             =   255
            Width           =   2220
         End
         Begin prjGiList.GiList gilNReqEn 
            Height          =   315
            Left            =   1575
            TabIndex        =   17
            Top             =   630
            Width           =   2235
            _ExtentX        =   3942
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
            FldWidth2       =   1300
            F2Corresponding =   -1  'True
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "修改原因"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   735
            TabIndex        =   46
            Top             =   1050
            Width           =   780
         End
         Begin VB.Label Label18 
            Alignment       =   1  '靠右對齊
            AutoSize        =   -1  'True
            Caption         =   "要求修改人員"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   345
            TabIndex        =   45
            Top             =   705
            Width           =   1170
         End
         Begin VB.Label Label17 
            Alignment       =   1  '靠右對齊
            AutoSize        =   -1  'True
            Caption         =   "要求修改單位"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   345
            TabIndex        =   44
            Top             =   330
            Width           =   1170
         End
      End
      Begin VB.TextBox txtContMonth 
         Alignment       =   1  '靠右對齊
         Height          =   315
         Left            =   7320
         MaxLength       =   2
         TabIndex        =   11
         Top             =   600
         Width           =   405
      End
      Begin VB.CheckBox chkPunish 
         Caption         =   "違約處份"
         Enabled         =   0   'False
         Height          =   195
         Left            =   1620
         TabIndex        =   9
         Top             =   2130
         Width           =   1095
      End
      Begin VB.CheckBox chkStopFlag 
         Caption         =   "停用"
         Height          =   195
         Left            =   8070
         TabIndex        =   26
         Top             =   660
         Width           =   735
      End
      Begin VB.TextBox txtPeriod 
         Alignment       =   1  '靠右對齊
         Height          =   315
         Left            =   3870
         MaxLength       =   2
         TabIndex        =   4
         Top             =   1350
         Width           =   405
      End
      Begin VB.TextBox txtContNo 
         Height          =   315
         Left            =   7320
         TabIndex        =   10
         Top             =   240
         Width           =   3525
      End
      Begin Gi_Date.GiDate gdaDiscountDate2 
         Height          =   315
         Left            =   9690
         TabIndex        =   6
         Top             =   1350
         Width           =   1155
         _ExtentX        =   2037
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
      Begin Gi_Date.GiDate gdaDiscountDate1 
         Height          =   315
         Left            =   7320
         TabIndex        =   5
         Top             =   1350
         Width           =   1155
         _ExtentX        =   2037
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
      Begin prjGiList.GiList gilDiscountCode 
         Height          =   315
         Left            =   1620
         TabIndex        =   2
         Top             =   990
         Width           =   3465
         _ExtentX        =   6112
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
         FldWidth2       =   2400
         TableName       =   "CD019"
         F2Corresponding =   -1  'True
      End
      Begin prjNumber.GiNumber ginDiscountAmt 
         Height          =   315
         Left            =   1620
         TabIndex        =   3
         Top             =   1350
         Width           =   1155
         _ExtentX        =   2037
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
      End
      Begin prjNumber.GiNumber ginMonthAmt 
         Height          =   315
         Left            =   1620
         TabIndex        =   7
         Top             =   1710
         Width           =   1155
         _ExtentX        =   2037
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
         MaxLength       =   14
         AutoSelect      =   -1  'True
         Decima          =   3
         AllowMinus      =   -1  'True
      End
      Begin prjNumber.GiNumber ginDayAmt 
         Height          =   315
         Left            =   3870
         TabIndex        =   8
         Top             =   1710
         Width           =   1155
         _ExtentX        =   2037
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
         MaxLength       =   9
         AutoSelect      =   -1  'True
         Decima          =   3
         AllowMinus      =   -1  'True
      End
      Begin Gi_Date.GiDate gdaStopDate 
         Height          =   315
         Left            =   9690
         TabIndex        =   28
         Top             =   600
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         ForeColor       =   -2147483630
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
      Begin prjGiList.GiList gilPromCode 
         Height          =   315
         Left            =   1620
         TabIndex        =   0
         Top             =   240
         Width           =   3465
         _ExtentX        =   6112
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
         FldWidth2       =   2400
         TableName       =   "CD019"
         F2Corresponding =   -1  'True
      End
      Begin Gi_Date.GiDate gdaContStopDate 
         Height          =   315
         Left            =   9690
         TabIndex        =   13
         Top             =   960
         Width           =   1155
         _ExtentX        =   2037
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
      Begin Gi_Date.GiDate gdaContStartDate 
         Height          =   315
         Left            =   7320
         TabIndex        =   12
         Top             =   960
         Width           =   1155
         _ExtentX        =   2037
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
      Begin prjGiList.GiList gilBpCode 
         Height          =   315
         Left            =   1620
         TabIndex        =   1
         Top             =   630
         Width           =   3465
         _ExtentX        =   6112
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
         FldWidth2       =   2400
         TableName       =   "CD019"
         F2Corresponding =   -1  'True
      End
      Begin Gi_Date.GiDate gdaOldContStopDate1 
         Height          =   315
         Left            =   7320
         TabIndex        =   58
         Top             =   1740
         Width           =   1155
         _ExtentX        =   2037
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
         Enabled         =   0   'False
      End
      Begin Gi_Date.GiDate gdaOldDiscountDate2 
         Height          =   315
         Left            =   7320
         TabIndex        =   59
         Top             =   2100
         Width           =   1155
         _ExtentX        =   2037
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
         Enabled         =   0   'False
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "活動說明"
         Height          =   195
         Left            =   450
         TabIndex        =   61
         Top             =   4170
         Width           =   780
      End
      Begin VB.Label lblOldDiscountDate2 
         AutoSize        =   -1  'True
         Caption         =   "原始優惠到期日"
         Height          =   195
         Left            =   5850
         TabIndex        =   60
         Top             =   2160
         Width           =   1365
      End
      Begin VB.Label lblOldContStopDate1 
         AutoSize        =   -1  'True
         Caption         =   "原始合約到期日"
         Height          =   195
         Left            =   5850
         TabIndex        =   57
         Top             =   1800
         Width           =   1365
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "優惠組合"
         Height          =   195
         Left            =   780
         TabIndex        =   47
         Top             =   690
         Width           =   780
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "合約月數"
         Height          =   195
         Left            =   6450
         TabIndex        =   42
         Top             =   660
         Width           =   780
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "優惠期限"
         Height          =   195
         Left            =   6450
         TabIndex        =   41
         Top             =   1410
         Width           =   780
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "優待辦法"
         Height          =   195
         Left            =   780
         TabIndex        =   40
         Top             =   1050
         Width           =   780
      End
      Begin VB.Label lblAmount 
         AutoSize        =   -1  'True
         Caption         =   "優待金額"
         Height          =   195
         Left            =   780
         TabIndex        =   39
         Top             =   1410
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "單月金額"
         Height          =   195
         Left            =   780
         TabIndex        =   38
         Top             =   1800
         Width           =   780
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "促銷方案"
         Height          =   195
         Left            =   780
         TabIndex        =   37
         Top             =   300
         Width           =   780
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "促銷方案說明"
         Height          =   195
         Left            =   390
         TabIndex        =   36
         Top             =   2520
         Width           =   1170
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "停用日期"
         Height          =   195
         Left            =   8850
         TabIndex        =   34
         Top             =   660
         Width           =   780
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "至"
         Height          =   195
         Left            =   8850
         TabIndex        =   33
         Top             =   1020
         Width           =   195
      End
      Begin VB.Label Label6 
         Caption         =   "合約期限"
         Height          =   195
         Left            =   6450
         TabIndex        =   32
         Top             =   1020
         Width           =   780
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "單日金額"
         Height          =   195
         Left            =   2910
         TabIndex        =   31
         Top             =   1770
         Width           =   780
      End
      Begin VB.Label lblPeriod 
         AutoSize        =   -1  'True
         Caption         =   "每次期數"
         Height          =   195
         Left            =   2910
         TabIndex        =   30
         Top             =   1410
         Width           =   780
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "至"
         Height          =   195
         Left            =   8850
         TabIndex        =   29
         Top             =   1410
         Width           =   195
      End
      Begin VB.Label Label7 
         Caption         =   "合約編號"
         Height          =   195
         Left            =   6450
         TabIndex        =   27
         Top             =   300
         Width           =   780
      End
   End
   Begin prjGiGridR.GiGridR ggrData 
      Height          =   1635
      Left            =   150
      TabIndex        =   19
      Top             =   5760
      Width           =   11625
      _ExtentX        =   20505
      _ExtentY        =   2884
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
   Begin VB.Label lblEditMode 
      Alignment       =   2  '置中對齊
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  '單線固定
      Caption         =   "顯示"
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   11100
      TabIndex        =   35
      Top             =   7500
      Width           =   675
   End
End
Attribute VB_Name = "frmSO1131F"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lngCustId As Long
Dim WithEvents rsData As ADODB.Recordset
Attribute rsData.VB_VarHelpID = -1
Dim objParentForm As Form
Dim strCitemCode As String
Dim strContNo As String
Dim strFaciSeqNo As String
Dim lngEditMode As giEditModeEnu
Dim blnReadRS As Boolean
Dim strClctDate As String
'Dim rsTmp As New ADODB.Recordset
Private strPenalType As String
Private strExpiretype As String
Private strFaciSNo As String
Private lngAbs As Long
Private strServiceType As String
Private strCustid As String
Private strSQLWhere As String
Private blnCopyIns As Boolean
Public Property Let uCustId(ByVal vData As String)
    lngCustId = vData
    
End Property

Public Property Let uClctDate(ByVal vData As String)
    strClctDate = vData
End Property

Public Property Let uCitemCode(ByVal vData As String)
    strCitemCode = vData
End Property

Public Property Let uContNo(ByVal vData As String)
    strContNo = vData
End Property
'#4165 要儲存FaciSNo By Kin 2008/10/23
Public Property Let uFaciSNo(ByVal vData As String)
    strFaciSNo = vData
End Property
Public Property Let uFaciSeqNo(ByVal vData As String)
    strFaciSeqNo = vData
End Property

Public Property Let uParentForm(ByVal vData As Form)
    Set objParentForm = vData
End Property
Public Property Let uPenalType(ByVal vData As String)
    strPenalType = vData
End Property
Public Property Let uExpiretype(ByVal vData As String)
    strExpiretype = vData
End Property
Public Property Let uServiceType(ByVal vData As String)
    strServiceType = vData
End Property
Private Sub ChangeMode(lngMode As giEditModeEnu)
    On Error GoTo ChkErr
    Dim blnFlag As Boolean
        lngEditMode = lngMode
        Select Case lngMode
            Case giEditModeInsert
                blnFlag = True
                lblEditMode.Caption = "新增"
                Call NewRcd
            Case giEditModeEdit
                blnFlag = True
                lblEditMode.Caption = "修改"
                Call RcdToScr
            Case giEditModeView
                blnFlag = False
                lblEditMode.Caption = "顯示"
                Call RcdToScr
        End Select
        If lngEditMode = giEditModeEdit Then
            fraUpdEn.Visible = True
            txtPromtion.Width = 4050
        Else
            fraUpdEn.Visible = False
            txtPromtion.Width = 9200
        End If
        fraData.Enabled = blnFlag
        ggrData.Enabled = Not blnFlag
        cmdAdd.Enabled = Not blnFlag And GetUserPriv("SO11319", "(SO113191)")
        '#7213
        cmdCopyIns.Enabled = cmdAdd.Enabled
        cmdDelete.Enabled = Not blnFlag And GetUserPriv("SO11319", "(SO113193)")
        cmdEdit.Enabled = Not blnFlag And GetUserPriv("SO11319", "(SO113192)")
        cmdSave.Enabled = blnFlag
        cmdLog.Enabled = Not blnFlag
        '#7287 to avoid cleaning data of gilDiscountCode By kin 2016/08/04
        If blnCopyIns Then
            gdaDiscountDate1.SetFocus
        End If
        blnCopyIns = False
    Exit Sub
ChkErr:
    ErrSub Me.Name, "ChangeMode"
End Sub

Private Sub RcdToScr()
    On Error GoTo ChkErr
        With rsData
            If .BOF Or .EOF Then Exit Sub
            blnReadRS = True
            gilPromCode.SetCodeNo .Fields("PromCode") & ""
            gilPromCode.SetDescription .Fields("PromName") & ""
            gilDiscountCode.SetCodeNo .Fields("DiscountCode") & ""
            gilDiscountCode.SetDescription .Fields("DiscountName") & ""
            gdaDiscountDate1.SetValue .Fields("DiscountDate1") & ""
            gdaDiscountDate2.SetValue .Fields("DiscountDate2") & ""
            txtPeriod.Text = Val(.Fields("Period") & "")
            ginDiscountAmt.Text = Val(.Fields("DiscountAmt") & "")
            ginMonthAmt.Text = Val(.Fields("MonthAmt") & "")
            ginDayAmt.Text = Val(.Fields("DayAmt") & "")
            chkPunish.Value = Val(.Fields("Punish") & "")
            txtContNo.Text = .Fields("ContNo") & ""
            '#4121 將Deposit欄位取消 By Kin 2008/10/01
            'ginDeposit.Text = Val(.Fields("Deposit") & "")
            '#4121 將PenalAmt欄位取消 By Kin 2008/10/01
            'ginPenalAmt.Text = Val(.Fields("PenalAmt") & "")
            txtContMonth.Text = Val(.Fields("ContMonth") & "")
            gdaContStartDate.SetValue .Fields("ContStartDate") & ""
            gdaContStopDate.SetValue .Fields("ContStopDate") & ""
            '#3588 Show欄位時要減1
            '#4121 將PenalType欄位取消 By Kin 2008/10/01
            'cboPenalType.ListIndex = Val(.Fields("PenalType") & "") - 1
            '#4121 將ExpireType欄位取消 By Kin 2008/10/01
            'cboExpireType.ListIndex = Val(.Fields("expiretype") & "") - 1
            chkStopFlag.Value = Val(.Fields("StopFlag") & "")
            gdaStopDate.SetValue .Fields("StopDate") & ""
            gilBpCode.SetCodeNo .Fields("Bpcode") & ""
            gilBpCode.SetDescription .Fields("BpName") & ""
            
            '#4306 增加原始合約到期日起迄 By Kin 2009/03/25
            '#4306 測試不OK,欄位有值時才要顯示 By Kin 2009/05/06
'            If .Fields("OldContStopDate") & "" <> "" Then
'                gdaOldContStopDate1.SetValue .Fields("OldContStopDate") & ""
'            End If
'            If .Fields("OldDiscountDate2") & "" <> "" Then
'                gdaOldDiscountDate2.SetValue .Fields("OldDiscountDate2") & ""
'            End If
            gdaOldContStopDate1.SetValue .Fields("OldContStopDate") & ""
            gdaOldDiscountDate2.SetValue .Fields("OldDiscountDate2") & ""
            SetPromtion False
            blnReadRS = False
'            txtNReqDep.Text = ""
'            txtNReqNotes.Text = ""
'            gilNReqEn.Clear
            
        End With
    Exit Sub
ChkErr:
    blnReadRS = False
    ErrSub Me.Name, "RcdToScr"
End Sub

Private Sub ScrToRcd()
    On Error GoTo ChkErr
        'CompCode,ServiceType,CustId,FaciSeqNo,CitemCode,ContNo,DiscountDate1
        Dim updSynchronize As Boolean
        updSynchronize = False
        gcnGi.BeginTrans
        With rsData
            If lngEditMode = giEditModeInsert Then
                .AddNew
                .Fields("FaciSeqNo") = strFaciSeqNo
                .Fields("CustId") = gCustId
                .Fields("CompCode") = gCompCode
                .Fields("ServiceType") = strServiceType
                .Fields("CitemCode") = strCitemCode
                '*************************************************************************
                '#4121 將SO003的資料直接代入 By Kin 2008/10/01
                .Fields("PenalType") = IIf(strPenalType = Empty, Null, strPenalType)
                .Fields("Expiretype") = IIf(strExpiretype = Empty, Null, strExpiretype)
                '*************************************************************************
                '#4165 要儲存FaciSno By Kin 2008/10/23
                .Fields("FaciSNo") = IIf(strFaciSNo = Empty, Null, strFaciSNo)
            End If
            If lngEditMode = giEditModeEdit Then
                If (.Fields("ContStartDate").OriginalValue <> gdaContStartDate.GetValue(True)) Or _
                    (.Fields("ContStopDate").OriginalValue <> gdaContStopDate.GetValue(True)) Then
                    updSynchronize = True
                End If
            End If
            .Fields("PromCode") = gilPromCode.GetCodeNo2
            .Fields("PromName") = gilPromCode.GetDescription2
            .Fields("DiscountCode") = gilDiscountCode.GetCodeNo2
            .Fields("DiscountName") = gilDiscountCode.GetDescription2
            .Fields("DiscountDate1") = NoZero(gdaDiscountDate1.GetValue(True))
            .Fields("DiscountDate2") = NoZero(gdaDiscountDate2.GetValue(True))
            .Fields("Period") = NoZero(txtPeriod.Text, True)
            .Fields("DiscountAmt") = NoZero(ginDiscountAmt.Value, True)
            .Fields("MonthAmt") = NoZero(ginMonthAmt.Value, True)
            .Fields("DayAmt") = NoZero(ginDayAmt.Value, True)
            .Fields("Punish") = chkPunish.Value
            .Fields("ContNo") = NoZero(txtContNo.Text)
            '#4121 將Deposit欄位取消,直接填0 By Kin 2008/10/01
            '.Fields("Deposit") = NoZero(ginDeposit.Value, True)
            .Fields("Deposit") = 0
            '#4121 將PenalAmt欄位取消,直接填0 By Kin 2008/10/01
            '.Fields("PenalAmt") = NoZero(ginPenalAmt.Value, True)
            .Fields("PenalAmt") = 0
            .Fields("ContMonth") = NoZero(txtContMonth.Text, True)
            .Fields("ContStartDate") = NoZero(gdaContStartDate.GetValue(True))
            .Fields("ContStopDate") = NoZero(gdaContStopDate.GetValue(True))
            '#3588 儲存時要加1
            '#4121 將PenalType欄位取消 By Kin 2008/10/01
            '.Fields("PenalType") = cboPenalType.ListIndex + 1
            '#4121 將ExpireType欄位取消 By Kin 2008/10/01
            '.Fields("Expiretype") = cboExpireType.ListIndex + 1
            .Fields("StopFlag") = chkStopFlag.Value
            .Fields("StopDate") = NoZero(gdaStopDate.GetValue(True))
            .Fields("UpdTime") = GetDTString(RightNow)
            .Fields("UpdEn") = garyGi(1)
            .Fields("BpCode") = gilBpCode.GetCodeNo2
            .Fields("BpName") = gilBpCode.GetDescription2
            If Not InsertToLog Then Exit Sub
            .Update
            '#7263 update ContStartDate field and ContStopDate field to same ContNo field when they have been updated by Kin 2016/07/05
             If updSynchronize Then
                Dim updSql As String
                updSql = "Update " & GetOwner & "SO003A Set ContStartDate =To_Date('" & gdaContStartDate.GetValue & "','yyyymmdd')," & _
                                "ContStopDate = To_Date('" & gdaContStopDate.GetValue & "','yyyymmdd') " & _
                                " Where CustId = " & lngCustId & " And CompCode = " & gCompCode & _
                                    " And ServiceType = '" & strServiceType & "' " & _
                                    " And ContNo = '" & txtContNo.Text & "' " & _
                                    " And FaciSeqNo = '" & strFaciSeqNo & "'"
                gcnGi.Execute updSql
                updSql = "update " & GetOwner & "SO003 Set ContStartDate =To_Date('" & gdaContStartDate.GetValue & "','yyyymmdd')," & _
                            "ContStopDate = To_Date('" & gdaContStopDate.GetValue & "','yyyymmdd') " & _
                             " Where CustId = " & lngCustId & " And CompCode = " & gCompCode & _
                                " And ServiceType = '" & strServiceType & "' " & _
                                " And ContNo = '" & txtContNo.Text & "' " & _
                                " And FaciSeqNo = '" & strFaciSeqNo & "'"
                gcnGi.Execute updSql
            End If
            
        End With
        gcnGi.CommitTrans
    Exit Sub
ChkErr:
    gcnGi.RollbackTrans
    ErrSub Me.Name, "ScrToRcd"
End Sub

Private Function InsertToLog() As Boolean
    On Error GoTo ChkErr
    Dim rs As New ADODB.Recordset
    Dim intLoop As Integer
        If Not GetRS(rs, "Select * From " & GetOwner & "SO003ALog Where CompCode = -1 And ServiceType = '' And CustId = -1", , adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
        With rs
            .AddNew
            On Error Resume Next
            For intLoop = 0 To rs.Fields.Count - 1
                .Fields(rsData(intLoop).Name) = rsData(intLoop).OriginalValue
            Next
            On Error GoTo ChkErr
            If lngEditMode = giEditModeInsert Then
                .Fields("FaciSeqNo") = strFaciSeqNo
                .Fields("CustId") = gCustId
                .Fields("CompCode") = gCompCode
                .Fields("ServiceType") = strServiceType
                .Fields("CitemCode") = strCitemCode
            End If
            .Fields("NCONTNO") = txtContNo
            .Fields("NDISCOUNTDATE1") = NoZero(gdaDiscountDate1.GetValue(True))
            .Fields("NDISCOUNTDATE2") = NoZero(gdaDiscountDate2.GetValue(True))
            .Fields("NPERIOD") = Val(txtPeriod.Text)
            .Fields("NDISCOUNTAMT") = NoZero(ginDiscountAmt.Value, True)
            .Fields("NMONTHAMT") = NoZero(ginMonthAmt.Value, True)
            .Fields("NDAYAMT") = NoZero(ginDayAmt.Value, True)
            .Fields("NPUNISH") = chkPunish.Value
            .Fields("NDISCOUNTCODE") = gilDiscountCode.GetCodeNo2
            .Fields("NDISCOUNTNAME") = gilDiscountCode.GetDescription2
            .Fields("NPROMCODE") = gilPromCode.GetCodeNo2
            .Fields("NPROMNAME") = gilPromCode.GetDescription2
            '#4121 將Deposit欄位取消,直接填0 By Kin 2008/10/01
            '.Fields("NDEPOSIT") = NoZero(ginDeposit.Value, True)
            .Fields("NDEPOSIT") = 0
            '#4121 將PenalAmt欄位取消,直接填0 By Kin 2008/10/01
            '.Fields("NPENALAMT") = NoZero(ginPenalAmt.Value, True)
            .Fields("NPENALAMT") = 0
            .Fields("NCONTSTARTDATE") = NoZero(gdaContStartDate.GetValue(True))
            .Fields("NCONTSTOPDATE") = NoZero(gdaContStopDate.GetValue(True))
            '#4121 將PenalType欄位取消,代入SO003的值 By Kin 2008/10/01
            '.Fields("NPENALTYPE") = cboPenalType.ListIndex
            .Fields("NPENALTYPE") = IIf(strPenalType = Empty, Null, strPenalType)
            '#4121 將ExpireType欄位取消,代入SO003的值 By Kin 2008/10/01
            '.Fields("NEXPIRETYPE") = cboExpireType.ListIndex
            .Fields("NEXPIRETYPE") = IIf(strExpiretype = Empty, Null, strExpiretype)
            .Fields("NCONTMONTH") = NoZero(txtContMonth.Text, True)
            .Fields("NewUpdTime") = RightNow
            If fraUpdEn.Visible Then
                .Fields("ModifyDep") = txtNReqDep.Text
                .Fields("ModifyEn") = gilNReqEn.GetCodeNo & ""
                .Fields("ModifyName") = gilNReqEn.GetDescription & ""
                .Fields("ModifyNotes") = txtNReqNotes.Text & ""
            End If
            If gilBpCode.GetCodeNo2 & "" <> "" Then
                .Fields("BpCode") = gilBpCode.GetCodeNo2 & ""
                .Fields("BpName") = gilBpCode.GetDescription2 & ""
            End If
            .Update
        End With
        InsertToLog = True
    Exit Function
ChkErr:
    ErrSub Me.Name, "InsertToLog"
End Function

Private Sub NewRcd()
    On Error Resume Next
        '#7213
        If blnCopyIns Then
            Exit Sub
        End If
        gilPromCode.Clear
        gilDiscountCode.Clear
        gdaDiscountDate1.SetValue ""
        gdaDiscountDate2.SetValue ""
        txtPeriod.Text = 0
        ginDiscountAmt.Text = 0
        ginMonthAmt.Text = 0
        ginDayAmt.Text = "0"
        chkPunish.Value = 0
        txtContNo.Text = ""
        '#4121 將Deposit欄位取消 By Kin 2008/10/01
        'ginDeposit.Text = 0
        '#4121 將PenalAmt欄位取消 By Kin 2008/10/01
        'ginPenalAmt.Text = 0
        gdaContStartDate.SetValue ""
        gdaContStopDate.SetValue ""
        '#4121 將PenalType欄位取消 By Kin 2008/10/01
        'cboPenalType.ListIndex = 0
        '#4121 將ExpireType欄位取消 By Kin 2008/10/01
        'cboExpireType.ListIndex = 0
        chkStopFlag.Value = 0
        gdaStopDate.SetValue ""
        txtPromtion.Text = ""
        txtNote.Text = ""
        txtNReqDep.Text = ""
        txtNReqNotes.Text = ""
        gilNReqEn.Clear
        gilBpCode.Clear
        gilPromCode.SetFocus
        
End Sub

Private Sub chkStopFlag_Click()
    On Error Resume Next
        If chkStopFlag.Value = 1 Then
            gdaStopDate.SetValue Date
        Else
            gdaStopDate.SetValue ""
        End If
End Sub

Private Sub cmdAdd_Click()
    txtContNo.Enabled = True
    Call ChangeMode(giEditModeInsert)
End Sub

Private Sub cmdCancel_Click()
    On Error Resume Next
        If lngEditMode = giEditModeView Then
            Unload Me
        Else
            Call ChangeMode(giEditModeView)
            txtContNo.Enabled = True
        End If
        
End Sub

Private Sub cmdCopyIns_Click()
    txtContNo.Enabled = True
    blnCopyIns = True
    Call ChangeMode(giEditModeInsert)
End Sub

Private Sub cmdDelete_Click()
    On Error GoTo ChkErr
        If rsData.BOF Or rsData.EOF Then Exit Sub
        If MsgBox("確定要停用??", vbQuestion + vbYesNo, gimsgPrompt) = vbNo Then Exit Sub
        With rsData
            .Fields("StopDate") = Date
            .Fields("StopFlag") = 1
            .Fields("UpdTime") = GetDTString(RightNow)
            .Fields("UpdEn") = garyGi(1)
            .Update
        End With
        frmSO1131C.uSO1131FSave = True
        Call ChangeMode(giEditModeView)
    Exit Sub
ChkErr:
    ErrSub Me.Name, "cmdDelete_Click"
End Sub

Private Sub cmdEdit_Click()
    If rsData.RecordCount = 1 Then rsData.AbsolutePosition = 1
    txtContNo.Enabled = False
    Call ChangeMode(giEditModeEdit)
End Sub

Private Sub cmdLog_Click()
  On Error GoTo ChkErr
  '#7213
   With frmSO1131F2
        .uCustId = lngCustId
        .uCitemCode = strCitemCode
        .uServiceType = strServiceType
        .uContNo = txtContNo
        .uFaciSNo = strFaciSNo
        .Show vbModal
    End With
    Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdLog_Click")
End Sub

Private Sub cmdSave_Click()
    If Not IsDataOk Then Exit Sub
    '#3168 如果資料重複，秀出警告訊息 By Kin 2007/06/20
    If rsData.RecordCount > 1 Then
        If chkSaveDouble = True Then MsgBox "資料已存在，請重新輸入!!", vbInformation, "警告訊息": Exit Sub
    End If
    ScrToRcd
    Call ChangeMode(giEditModeView)
    txtContNo.Enabled = True
    frmSO1131C.uSO1131FSave = True
End Sub

Private Function IsDataOk() As Boolean
    On Error GoTo ChkErr
        If chkStopFlag.Value = 1 Then If Not MustExist(gdaStopDate, 1, "停用日期") Then Exit Function
        If lngEditMode = giEditModeEdit Then
            If Not MustExist(txtNReqDep, , "要求修改單位") Then Exit Function
            If Not MustExist(gilNReqEn, 2, "要求修改人員") Then Exit Function
            If Not MustExist(txtNReqNotes, , "修改原因") Then Exit Function
        End If
        If Not MustExist(txtContNo, , "合約編號") Then Exit Function
        If Not MustExist(gdaDiscountDate1, 1, "優惠期限起始日") Then Exit Function
        If Not MustExist(gdaDiscountDate2, 1, "優惠期限截止日") Then Exit Function
        '#4258 每次期數必須大於0 By Kin 2008/12/22
        If txtPeriod.Text <= 0 Then MsgBox "每次期數必須大於0", vbInformation, "警告訊息": Exit Function
        '#4176 要自動判斷正負數 By Kin 2008/10/27
        If Not ChkSign Then Exit Function
        IsDataOk = True
    Exit Function
ChkErr:
    ErrSub Me.Name, "IsDataOk"
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
        Select Case KeyCode
            Case vbKeyF2
                If cmdSave.Enabled Then cmdSave.Value = True
            Case vbKeyF6
                If cmdAdd.Enabled Then cmdAdd.Value = True
            Case vbKeyF10
                If cmdDelete.Enabled Then cmdDelete.Value = True
            Case vbKeyF11
                If cmdEdit.Enabled Then cmdEdit.Value = True
            Case vbKeyF12
                If cmdCopyIns.Enabled Then cmdCopyIns.Value = True
            Case vbKeyEscape
                If cmdCancel.Enabled Then cmdCancel.Value = True
        End Select
        
End Sub

Private Sub Form_Load()
    On Error Resume Next
        Call InitData
        Call subGil
        Call SetCombol
        Call subGrd
        Call OpenData
        Call ChangeMode(giEditModeView)
        Call DefaultValue
        txtContNo.MaxLength = rsData.Fields("ContNo").DefinedSize
End Sub

Private Sub DefaultValue()
    With rsData
        If .EOF Or .BOF Then Exit Sub
        Do While Not .EOF
            If StrToDate(strClctDate) >= StrToDate(.Fields("DiscountDate1") & "") And StrToDate(strClctDate) <= StrToDate(.Fields("DiscountDate2") & "") Then Exit Sub
            .MoveNext
        Loop
        .MoveFirst
    End With
End Sub

Private Sub subGil()
    On Error Resume Next
        Call SetgiList(gilDiscountCode, "CodeNo", "Description", "CD048", , , , , 3, 30, , True)
        SetgiList gilPromCode, "CodeNo", "Description", "CD042", , , , , 3, 12, " Where to_char(LastOrderDate,'yyyymmdd') >= to_char(sysdate,'yyyymmdd') ", True
        SetgiList gilBpCode, "CodeNo", "Description", "CD078", , , , , 3, 12, , True
        Call SetgiList(gilNReqEn, "EmpNo", "EmpName", "CM003", , , , , 10, 20, , True)

End Sub

Private Sub SetCombol()
    On Error Resume Next
        '#4121 將PenalType欄位取消 By Kin 2008/10/01
'        cboPenalType.AddItem "1.回溯牌價"
'        cboPenalType.AddItem "2.優惠價"
'        cboPenalType.AddItem "3.優惠平均價"
        '#4121 將ExpireType欄位取消 By Kin 2008/10/01
'        cboExpireType.AddItem "1.恢復最新公告牌價"
'        cboExpireType.AddItem "2.以最後一階繼續優惠"
'        '************************************************
'        '#3732 多增加以下兩個選項 By Kin 2008/06/09
'        cboExpireType.AddItem "3.到期不出"
'        cboExpireType.AddItem "4.接新約"
        '************************************************
End Sub

Private Sub InitData()
    On Error Resume Next
        Me.Move objParentForm.Left, objParentForm.Top + objParentForm.Height - Me.Height
End Sub

Private Sub OpenData()
    On Error GoTo ChkErr
    Dim strSQL As String
        Set rsData = New ADODB.Recordset
        '#5072 服務別與客編改成根據傳來的RecordSet為主 By Kin 2009/05/05
        strSQL = "Select * From " & GetOwner & "SO003A Where CustId = " & _
            lngCustId & " And CompCode = " & gCompCode & " And ServiceType = '" & _
            strServiceType & "' And CitemCode = " & IIf(strCitemCode = Empty, 0, strCitemCode) & " And FaciSeqNo = '" & strFaciSeqNo & "'"
            
        If strContNo <> "" Then strSQL = strSQL & " And ContNo = '" & strContNo & "'"
        strSQL = strSQL & " Order By DiscountDate1"
        
        If Not GetRS(rsData, strSQL, , adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
        'Set rsTmp = rsData.Clone
        Set ggrData.Recordset = rsData
        ggrData.Refresh
        If rsData.RecordCount > 0 Then
            rsData.AbsolutePosition = 1
        End If
    Exit Sub
ChkErr:
    ErrSub Me.Name, "OpenData"
End Sub

Private Sub subGrd()
    On Error Resume Next
    Dim mFlds As New GiGridFlds
        mFlds.Add "DiscountDate1", giControlTypeDate, , , , "優待起始日"
        mFlds.Add "DiscountDate2", giControlTypeDate, , , , "優待截止日"
        mFlds.Add "Period", , , , , "每次期數"
        mFlds.Add "DiscountAmt", , , , , "優待金額"
        mFlds.Add "MonthAmt", , , , , "單月金額"
        mFlds.Add "DayAmt", , , , , "單日金額"
        mFlds.Add "Punish", , , , , "違約處份"
        mFlds.Add "StopFlag", , , , , "停用"
        mFlds.Add "StopDate", giControlTypeDate, , , , "  停用日期  "
        mFlds.Add "UpdTime", , , , , "異動時間" & Space(10)
        mFlds.Add "UpdEn", , , , , "異動人員"
        ggrData.AllFields = mFlds

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
        If lngEditMode <> giEditModeView Then
            If giMsgNotSave = vbNo Then Cancel = True: Exit Sub
        End If
        Call CloseRecordset(rsData)
        'Call CloseRecordset(rsTmp)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
        'ReleaseCOM frmSO1131F

End Sub

Private Sub gdaContStartDate_Change()
    On Error Resume Next
        If blnReadRS Then Exit Sub
        AutoDate gdaContStartDate, gdaContStopDate, txtContMonth
End Sub

Private Sub gdaDiscountDate1_Change()
    On Error Resume Next
        If blnReadRS Then Exit Sub
        '#4150 將優惠期間連動取消 By Kin 2008/10/17
        'AutoDate gdaDiscountDate1, gdaDiscountDate2, txtPeriod
    
End Sub

Private Sub gdaStopDate_Change()
    If blnReadRS Then Exit Sub
    If gdaStopDate.GetValue = "" Then
        chkStopFlag = 0
    Else
        If InStr(gdaStopDate.GetOriginalValue, "_") > 0 Then Exit Sub
        If Not gdaStopDate.RaiseValidateEvent Then Exit Sub
        chkStopFlag = 1
    End If
End Sub

Private Sub gilPromCode_Validate(Cancel As Boolean)
    SetPromtion True
End Sub

Private Sub SetPromtion(blnSetDisCount As Boolean)
    On Error GoTo ChkErr
    Dim rs As New ADODB.Recordset
        If gilPromCode.GetCodeNo = "" Then Exit Sub
        If Not GetRS(rs, "Select PromotionCont,DiscountCode,DiscountName,Note From " & GetOwner & "CD042 Where CodeNo = " & gilPromCode.GetCodeNo) Then Exit Sub
        If rs.EOF Or rs.BOF Then Exit Sub
        txtPromtion.Text = rs("PromotionCont") & ""
        txtNote.Text = rs("Note") & ""
        If blnSetDisCount Then
            gilDiscountCode.SetCodeNo rs("DiscountCode") & ""
            gilDiscountCode.SetDescription rs("DiscountName") & ""
        End If
        Call CloseRecordset(rs)
    Exit Sub
ChkErr:
    ErrSub Me.Name, "SetPromtion"
End Sub

Private Sub AutoDate(gdaStartDate As GiDate, gdaStopDate As GiDate, txtPeriod As Object)
    On Error GoTo ChkErr
    Dim strStopDate As String
        If InStr(gdaStartDate.GetOriginalValue, "_") > 0 Then Exit Sub
        If Not ChkDTok Then Exit Sub
        GetAutoStopDate Val(txtPeriod.Text), gdaStartDate.GetValue(True), strStopDate, ""
        gdaStopDate.SetValue strStopDate
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "AutoDate")
End Sub



Private Sub ginDiscountAmt_Validate(Cancel As Boolean)
  On Error GoTo ChkErr
    
    '***************************************************************************
    '#4121 自動輸入單月金額、單日金額 By Kin 2008/10/01
    If Val(txtPeriod.Text) <> 0 And Val(ginDiscountAmt.Text) <> 0 Then
        ginMonthAmt.Text = Val(ginDiscountAmt.Value) / Val(txtPeriod.Text)
        ginDayAmt.Text = Val(ginDiscountAmt.Value) / Val(txtPeriod.Text) / 30
    Else
        ginMonthAmt.Text = "0"
        ginDayAmt.Text = "0"
    End If
    '**************************************************************************
    
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "ginDiscountAmt_Validate")
End Sub

Private Sub rsData_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
      If lngEditMode = giEditModeView Then Call RcdToScr
End Sub

Private Sub txtContMonth_Change()
        If blnReadRS Then Exit Sub
        AutoDate gdaContStartDate, gdaContStopDate, txtContMonth

End Sub

Private Sub txtPeriod_Change()
  On Error GoTo ChkErr
    If blnReadRS Then Exit Sub
    '#4150 將優惠期間連動取消 By Kin 2008/10/17
    'AutoDate gdaDiscountDate1, gdaDiscountDate2, txtPeriod
    
    '***********************************************************************
    '#4121 自動輸入單月金額、單日金額 By Kin 2008/10/01
    If Val(txtPeriod.Text) <> 0 And Val(ginDiscountAmt.Text) <> 0 Then
        ginMonthAmt.Text = Val(ginDiscountAmt.Value) / Val(txtPeriod.Text)
        ginDayAmt.Text = Val(ginDiscountAmt.Value) / Val(txtPeriod.Text) / 30
    Else
        ginMonthAmt.Text = "0"
        ginDayAmt.Text = "0"
    End If
    '**************************************************************************
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "txtPeriod_Change")
End Sub


Private Sub txtPeriod_KeyPress(keyAscii As Integer)
  On Error Resume Next
    If keyAscii = 8 Then Exit Sub
    If keyAscii >= 48 And keyAscii <= 57 Then Exit Sub
    keyAscii = 0
End Sub

Private Sub txtPromtion_KeyPress(keyAscii As Integer)
    keyAscii = 0
End Sub
'#3168 儲存時多增加判斷資料是否重複 By Kin 2007/06/20
Private Function chkSaveDouble() As Boolean
On Error GoTo ChkErr
    Dim strChkSQL As String
    Dim blnChkDouble As Boolean
    blnChkDouble = False
    '#3334 判斷是否在編輯模式，如果是編輯模式，而且候選鍵又沒被異動的話，不檢查資料是否有重複 By Kin 2007/07/23
    If lngEditMode = giEditModeEdit Then
        If rsData("CompCode") <> gCompCode Then blnChkDouble = True
        If rsData("ServiceType") <> strServiceType Then blnChkDouble = True
        If rsData("CustId") <> gCustId Then blnChkDouble = True
        If rsData("FaciSeqNo") <> strFaciSeqNo Then blnChkDouble = True
        If rsData("CitemCode") <> strCitemCode Then blnChkDouble = True
        If rsData("ContNo") <> Trim(txtContNo.Text) Then blnChkDouble = True
        If rsData("DisCountDate1") <> gdaDiscountDate1.GetValue(True) Then blnChkDouble = True
    Else
        blnChkDouble = True
    End If
    '檢查資料是否重複
    If blnChkDouble Then
        strChkSQL = "Select Count(*) From " & GetOwner & "SO003A Where CompCode=" & gCompCode & _
                    " And ServiceType='" & strServiceType & "' And Custid=" & gCustId & _
                    " And FaciSeqNo='" & strFaciSeqNo & "' And CitemCode=" & strCitemCode & _
                    " And ContNo='" & Trim(txtContNo.Text) & "' And DisCountDate1=to_Date(" & gdaDiscountDate1.GetValue & " ,'yyyymmdd')"
        If gcnGi.Execute(strChkSQL)(0) > 0 Then chkSaveDouble = True: Exit Function
        chkSaveDouble = False
    Else
        chkSaveDouble = False
    End If
    Exit Function
ChkErr:
    Call ErrSub(Me.Name, "chkSaveDouble")
End Function
Private Function ChkSign() As Boolean
    On Error GoTo ChkErr
    Dim strMsg As String
    Dim strSign As String
    Dim blnGetFocus As Boolean
    Dim blnError As Boolean
    blnGetFocus = False
    blnError = False
        strSign = GetRsValue("Select Sign From " & GetOwner & "CD019 Where CodeNo = " & strCitemCode) & ""
        strMsg = "借貸符號為 '" & strSign & "',只能輸入'" & strSign & "' 金額"
        If Val(ginDiscountAmt.Value) <> Abs(Val(ginDiscountAmt.Value)) * Val(strSign & "1") Then
            ginDiscountAmt.Text = Abs(Val(ginDiscountAmt.Value)) * Val(strSign & "1")
            ginDiscountAmt.SetFocus
            blnGetFocus = True
            blnError = True
        End If
        If Val(ginMonthAmt.Value) <> Abs(Val(ginMonthAmt.Value)) * Val(strSign & "1") Then
            ginMonthAmt.Text = Abs(Val(ginMonthAmt.Value)) * Val(strSign & "1")
            If Not blnGetFocus Then ginMonthAmt.SetFocus
            blnError = True
        End If
        If Val(ginDayAmt.Value) <> Abs(Val(ginDayAmt.Value)) * Val(strSign & "1") Then
            ginDayAmt.Text = Abs(Val(ginDayAmt.Value)) * Val(strSign & "1")
            If Not blnGetFocus Then ginDayAmt.SetFocus
            blnError = True
        End If
        If blnError Then MsgBox strMsg, vbExclamation, gimsgPrompt: Exit Function
        ChkSign = True
    Exit Function
ChkErr:
    ErrSub Me.Name, "ChkSign"
End Function
