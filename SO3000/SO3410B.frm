VERSION 5.00
Object = "{3BB22E57-824A-11D3-BD70-0080C8F80BC4}#3.4#0"; "GiGridR.ocx"
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{A0C51418-555F-11D3-BC10-0050BA01A3A9}#9.9#0"; "GiMulti.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{6DDA0569-6454-11D3-9612-0048544B09DF}#4.1#0"; "GiTime.ocx"
Object = "{074771B7-6ADE-11D5-B178-00485456B028}#2.1#0"; "csMulti.ocx"
Begin VB.Form frmSO3410B 
   BorderStyle     =   1  '單線固定
   Caption         =   "收費單未收原因登錄 [SO3410B]"
   ClientHeight    =   7575
   ClientLeft      =   45
   ClientTop       =   330
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
   Icon            =   "SO3410B.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   11910
   StartUpPosition =   2  '螢幕中央
   Begin VB.TextBox txtSQL 
      Height          =   525
      Left            =   5460
      MultiLine       =   -1  'True
      ScrollBars      =   2  '垂直捲軸
      TabIndex        =   40
      Top             =   3270
      Visible         =   0   'False
      Width           =   1245
   End
   Begin TabDlg.SSTab sstData 
      Height          =   6435
      Left            =   120
      TabIndex        =   9
      Top             =   1050
      Width           =   11685
      _ExtentX        =   20611
      _ExtentY        =   11351
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      ForeColor       =   16711680
      TabCaption(0)   =   "單筆登錄"
      TabPicture(0)   =   "SO3410B.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "chkUseOldBillNo"
      Tab(0).Control(1)=   "chkAutoSave"
      Tab(0).Control(2)=   "txtBillNo"
      Tab(0).Control(3)=   "txtNote"
      Tab(0).Control(4)=   "cboBillType"
      Tab(0).Control(5)=   "ggrData"
      Tab(0).Control(6)=   "Label1"
      Tab(0).Control(7)=   "lblCustId"
      Tab(0).Control(8)=   "lbl3"
      Tab(0).Control(9)=   "lbl4"
      Tab(0).Control(10)=   "lblName"
      Tab(0).Control(11)=   "Label2"
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "整批登錄"
      TabPicture(1)   =   "SO3410B.frx":045E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lblD3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblShouldDay"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lblPrtSNo2"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label3"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "lblBillHeadFmt"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label4"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label5"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "gdtCreatTime2"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "gdaShouldDate2"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "gdaShouldDate1"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "gimCitemCode"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "gimClassCode"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "gimCMCode"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "gimBillType"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "gimMduId"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "fraClose"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "txtPrtSNo1"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "txtPrtSNo2"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "cboBillHeadFmt"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "cmdSave"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "gdtCreatTime1"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).ControlCount=   21
      Begin Gi_Time.GiTime gdtCreatTime1 
         Height          =   315
         Left            =   4650
         TabIndex        =   12
         Top             =   510
         Width           =   1905
         _ExtentX        =   3360
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
      Begin VB.CheckBox chkUseOldBillNo 
         Caption         =   "是否使用舊版單號"
         Height          =   195
         Left            =   -71850
         TabIndex        =   48
         Top             =   570
         Width           =   1875
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "F2. 存檔"
         Height          =   375
         Left            =   10170
         TabIndex        =   47
         Top             =   1710
         Width           =   1185
      End
      Begin VB.CheckBox chkAutoSave 
         Caption         =   "自動存檔"
         Height          =   195
         Left            =   -71850
         TabIndex        =   7
         Top             =   570
         Value           =   1  '核取
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.ComboBox cboBillHeadFmt 
         Height          =   315
         Left            =   6840
         Style           =   2  '單純下拉式
         TabIndex        =   16
         Top             =   1020
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox txtPrtSNo2 
         Height          =   300
         Left            =   3150
         MaxLength       =   12
         TabIndex        =   15
         Top             =   1020
         Width           =   1650
      End
      Begin VB.TextBox txtPrtSNo1 
         Height          =   300
         Left            =   1260
         MaxLength       =   12
         TabIndex        =   14
         Top             =   1020
         Width           =   1650
      End
      Begin VB.Frame fraClose 
         Caption         =   "是否已給號"
         ForeColor       =   &H00C00000&
         Height          =   615
         Left            =   300
         TabIndex        =   32
         Top             =   1440
         Width           =   4485
         Begin VB.OptionButton optYES 
            Caption         =   "是"
            Height          =   195
            Left            =   360
            TabIndex        =   17
            Top             =   300
            Width           =   975
         End
         Begin VB.OptionButton optNo 
            Caption         =   "否"
            Height          =   195
            Left            =   1755
            TabIndex        =   18
            Top             =   300
            Width           =   1065
         End
         Begin VB.OptionButton optAll 
            Caption         =   "全部"
            Height          =   195
            Left            =   3180
            TabIndex        =   19
            Top             =   300
            Value           =   -1  'True
            Width           =   735
         End
      End
      Begin VB.TextBox txtBillNo 
         Height          =   375
         Left            =   -73515
         MaxLength       =   15
         TabIndex        =   6
         Top             =   480
         Width           =   1605
      End
      Begin VB.TextBox txtNote 
         Height          =   765
         Left            =   -73515
         MaxLength       =   15
         MultiLine       =   -1  'True
         ScrollBars      =   2  '垂直捲軸
         TabIndex        =   8
         Top             =   990
         Width           =   9885
      End
      Begin VB.ComboBox cboBillType 
         Height          =   315
         ItemData        =   "SO3410B.frx":047A
         Left            =   -74700
         List            =   "SO3410B.frx":0487
         Style           =   2  '單純下拉式
         TabIndex        =   5
         Top             =   510
         Width           =   1185
      End
      Begin CS_Multi.CSmulti gimMduId 
         Height          =   375
         Left            =   270
         TabIndex        =   24
         Top             =   3915
         Width           =   11145
         _ExtentX        =   19659
         _ExtentY        =   661
         ButtonCaption   =   "大樓編號"
      End
      Begin Gi_Multi.GiMulti gimBillType 
         Height          =   375
         Left            =   270
         TabIndex        =   21
         Top             =   2625
         Width           =   11145
         _ExtentX        =   19659
         _ExtentY        =   661
         ButtonCaption   =   "單據類別"
         DataType        =   2
         FldCaption1     =   "單據類別代碼"
         FldCaption2     =   "單據類別名稱"
         DIY             =   -1  'True
         Exception       =   -1  'True
         FontSize        =   9
         FontName        =   "新細明體"
      End
      Begin Gi_Multi.GiMulti gimCMCode 
         Height          =   375
         Left            =   270
         TabIndex        =   23
         Top             =   3480
         Width           =   11145
         _ExtentX        =   19659
         _ExtentY        =   661
         ButtonCaption   =   "收費方式"
         FontSize        =   9
         FontName        =   "新細明體"
      End
      Begin CS_Multi.CSmulti gimClassCode 
         Height          =   375
         Left            =   270
         TabIndex        =   22
         Top             =   3060
         Width           =   11145
         _ExtentX        =   19659
         _ExtentY        =   661
         ButtonCaption   =   "客戶類別"
      End
      Begin CS_Multi.CSmulti gimCitemCode 
         Height          =   375
         Left            =   270
         TabIndex        =   20
         Top             =   2190
         Width           =   11145
         _ExtentX        =   19659
         _ExtentY        =   661
         ButtonCaption   =   "收費項目"
      End
      Begin Gi_Date.GiDate gdaShouldDate1 
         Height          =   315
         Left            =   1260
         TabIndex        =   10
         Top             =   510
         Width           =   1095
         _ExtentX        =   1931
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
      Begin Gi_Date.GiDate gdaShouldDate2 
         Height          =   315
         Left            =   2640
         TabIndex        =   11
         Top             =   510
         Width           =   1095
         _ExtentX        =   1931
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
      Begin prjGiGridR.GiGridR ggrData 
         Height          =   4395
         Left            =   -74760
         TabIndex        =   45
         Top             =   1860
         Width           =   11145
         _ExtentX        =   19659
         _ExtentY        =   7752
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
      Begin Gi_Time.GiTime gdtCreatTime2 
         Height          =   315
         Left            =   6840
         TabIndex        =   13
         Top             =   510
         Width           =   1905
         _ExtentX        =   3360
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
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "至"
         ForeColor       =   &H00800080&
         Height          =   195
         Left            =   6600
         TabIndex        =   50
         Top             =   570
         Width           =   195
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "產單日期"
         ForeColor       =   &H00800080&
         Height          =   195
         Left            =   3810
         TabIndex        =   49
         Top             =   570
         Width           =   780
      End
      Begin VB.Label Label1 
         Caption         =   "(收費項目、應收金額、應收期數、實收金額、實收期數、本期有效期限、作廢、備註)"
         Height          =   315
         Left            =   -74460
         TabIndex        =   46
         Top             =   1830
         Visible         =   0   'False
         Width           =   8565
      End
      Begin VB.Label lblCustId 
         AutoSize        =   -1  'True
         Caption         =   "99999999"
         Height          =   195
         Left            =   -65400
         TabIndex        =   44
         Top             =   570
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label lbl3 
         AutoSize        =   -1  'True
         Caption         =   "姓名:"
         Height          =   195
         Left            =   -69750
         TabIndex        =   43
         Top             =   570
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Label lbl4 
         AutoSize        =   -1  'True
         Caption         =   "客戶編號:"
         Height          =   195
         Left            =   -66540
         TabIndex        =   42
         Top             =   570
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Caption         =   "XXXXXXXXXXXXXXXXXX"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   -69150
         TabIndex        =   41
         Top             =   570
         Visible         =   0   'False
         Width           =   2430
      End
      Begin VB.Label lblBillHeadFmt 
         AutoSize        =   -1  'True
         Caption         =   "多帳戶產生依據設定"
         Height          =   195
         Left            =   5040
         TabIndex        =   35
         Top             =   1080
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "印單序號"
         ForeColor       =   &H00800080&
         Height          =   195
         Left            =   420
         TabIndex        =   34
         Top             =   1050
         Width           =   780
      End
      Begin VB.Label lblPrtSNo2 
         AutoSize        =   -1  'True
         Caption         =   "至"
         ForeColor       =   &H00800080&
         Height          =   180
         Left            =   2940
         TabIndex        =   33
         Top             =   1080
         Width           =   195
      End
      Begin VB.Label lblShouldDay 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "應收日期"
         ForeColor       =   &H00800080&
         Height          =   195
         Left            =   420
         TabIndex        =   31
         Top             =   570
         Width           =   780
      End
      Begin VB.Label lblD3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "至"
         ForeColor       =   &H00800080&
         Height          =   195
         Left            =   2400
         TabIndex        =   30
         Top             =   600
         Width           =   195
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "收費備註"
         Height          =   195
         Left            =   -74325
         TabIndex        =   29
         Top             =   1080
         Width           =   780
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&X)"
      Height          =   375
      Left            =   10320
      TabIndex        =   25
      Top             =   30
      Width           =   1335
   End
   Begin prjGiList.GiList gilCompCode 
      Height          =   315
      Left            =   1110
      TabIndex        =   0
      Top             =   90
      Width           =   2430
      _ExtentX        =   4286
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
      Locked          =   -1  'True
      FldWidth1       =   600
      FldWidth2       =   1500
      F5Corresponding =   -1  'True
   End
   Begin prjGiList.GiList gilUCCode 
      Height          =   315
      Left            =   9525
      TabIndex        =   4
      Top             =   630
      Width           =   2130
      _ExtentX        =   3757
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
      FldWidth2       =   1200
      F2Corresponding =   -1  'True
   End
   Begin prjGiList.GiList gilCMCode 
      Height          =   315
      Left            =   6375
      TabIndex        =   3
      Top             =   630
      Width           =   2130
      _ExtentX        =   3757
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
      ListType        =   0
      FldWidth1       =   600
      FldWidth2       =   1200
      F2Corresponding =   -1  'True
   End
   Begin Gi_Date.GiDate gdaRealDate 
      Height          =   315
      Left            =   4245
      TabIndex        =   2
      Top             =   630
      Width           =   1185
      _ExtentX        =   2090
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
   Begin prjGiList.GiList gilClctEn 
      Height          =   315
      Left            =   1110
      TabIndex        =   1
      Top             =   630
      Width           =   2130
      _ExtentX        =   3757
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
      FldWidth2       =   1200
      F2Corresponding =   -1  'True
   End
   Begin VB.Label lblCMCode 
      AutoSize        =   -1  'True
      Caption         =   "收費方式"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   5565
      TabIndex        =   39
      Top             =   690
      Width           =   780
   End
   Begin VB.Label lblClctEn 
      AutoSize        =   -1  'True
      Caption         =   "收費人員"
      Height          =   195
      Left            =   210
      TabIndex        =   38
      Top             =   690
      Width           =   780
   End
   Begin VB.Label lblPTCode 
      AutoSize        =   -1  'True
      Caption         =   "未收原因"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   8685
      TabIndex        =   37
      Top             =   690
      Width           =   780
   End
   Begin VB.Label lblRealDate 
      AutoSize        =   -1  'True
      Caption         =   "登錄日期"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   3435
      TabIndex        =   36
      Top             =   690
      Width           =   780
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      Caption         =   "登錄單據數:"
      Height          =   195
      Left            =   5010
      TabIndex        =   28
      Top             =   180
      Width           =   1020
   End
   Begin VB.Label lblBillCnt 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Left            =   6120
      TabIndex        =   27
      Top             =   180
      Width           =   90
   End
   Begin VB.Label lblComp 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "公  司  別"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   210
      TabIndex        =   26
      Top             =   180
      Width           =   765
   End
End
Attribute VB_Name = "frmSO3410B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Private ObjopenDb As New giOpenDBObj.OpenDBObj
Private rs As New ADODB.Recordset
Dim strSQL As String

Private strDate As String '登錄日期
Private lngPara6 As Long '參數6
Private strUCCode As String  '未收原因
Private strUCName  As String  '未收原因
Private strCMCode As String '收費方式
Private strCMname As String '收費方式
Private strClctEn As String '收費人員
Private strClctName As String  '收費人員
Private strOldBillNo As String
Private strCompCode As String
Private intPara23 As Integer
Private intPara24 As Integer
Private rsBillHeadFmt As New ADODB.Recordset
Private rsBill As New ADODB.Recordset
Dim lngBillcount As Long

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
        'gimCitemCode.ShowMulti
End Sub

Private Sub cboBillType_Click()
    On Error Resume Next
        txtBillNo.Text = ""
        Select Case cboBillType.ListIndex
            Case 0
                txtBillNo.MaxLength = 15
            Case 1
                txtBillNo.MaxLength = 12
            Case 2
                txtBillNo.MaxLength = 11
            
        End Select
        txtBillNo.SetFocus
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Function IsDataOk() As Boolean
    On Error GoTo chkErr
        If Not ChkDTok Then Exit Function
'        If Not MustExist(gilClctEn, 2, "收費人員") Then Exit Function
        'If Not MustExist(gdaRealDate, 1, "登錄日期") Then Exit Function
        If Not MustExist(gilCMCode, 2, "收費方式") Then Exit Function
        If Not MustExist(gilUCCode, 2, "未收原因") Then Exit Function
        IsDataOk = True
    Exit Function
66:
    MsgMustBe
    Exit Function
chkErr:
    ErrSub Me.Name, "IsDataOk"
End Function

Private Function subChoose() As Boolean
    On Error GoTo chkErr
        strChoose = ""
        If gdaShouldDate1.GetValue <> "" Then subAnd "SO033.ShouldDate >=" & GetNullString(gdaShouldDate1.GetValue(True), giDateV)
        If gdaShouldDate2.GetValue <> "" Then subAnd "SO033.ShouldDate <" & GetNullString(gdaShouldDate2.GetValue(True), giDateV) & " + 1"
        If txtPrtSNo1 <> "" Then subAnd "PrtSNo >= '" & txtPrtSNo1 & "'"
        If txtPrtSNo2 <> "" Then subAnd "PrtSNo <= '" & txtPrtSNo2 & "'"
        
        If gilCompCode.GetCodeNo <> "" Then subAnd gilCompCode.GetCodeNo & " =SO033.CompCode"
        
        If gimCMCode.GetQryStr <> "" Then subAnd "SO033.CMCode " & gimCMCode.GetQryStr
        If gimCitemCode.GetQryStr <> "" Then subAnd "SO033.CitemCode " & gimCitemCode.GetQryStr
        
        If gimClassCode.GetQryStr <> "" Then subAnd "SO033.ClassCode " & gimClassCode.GetQryStr
        '#2908整批登錄增加產單範圍條件 By Kin 2007/04/02
        If gdtCreatTime1.GetValue <> "" Then subAnd "SO033.CreateTime>=To_Date(" & gdtCreatTime1.GetValue & ",'yyyymmddhh24mi')"
        If gdtCreatTime2.GetValue <> "" Then subAnd "SO033.CreateTime<=To_Date(" & gdtCreatTime2.GetValue & ",'yyyymmddhh24mi')"
        
        If gimBillType.GetQryStr <> "" Then subAnd "substr(SO033.BillNo,7,1) " & gimBillType.GetQryStr
        
        If optYES Then subAnd "PrtSNo Is not null"
        If optNo Then subAnd "PrtSNo Is null"
        
        If gimMduId.GetQryStr <> "" Then
            subAnd "SO033.MduId " & gimMduId.GetQryStr
        Else
            If gimMduId.GetDispStr <> "" Then subAnd "So033.MduId Is Not Null"
        End If
        subAnd "UCCode is not null"
        subChoose = True
    Exit Function
chkErr:
    ErrSub Me.Name, "subChoose"
End Function

Private Sub cmdSave_Click()
    On Error GoTo chkErr
    Dim strUpdate As String
    Dim lngAffectTotal As Long
        If Not IsDataOk Then Exit Sub
        '#3869 登錄日期有值才填入實收日期 By Kin 2008/06/04
        strUpdate = "Update " & GetOwner & "So033 Set  " & _
                                    strUpdate & "CMCode=" & GetNullString(gilCMCode.GetCodeNo) & _
                                   ",CMName=" & GetNullString(gilCMCode.GetDescription) & _
                                   ",UCCode=" & GetNullString(gilUCCode.GetCodeNo) & _
                                   ",UCName=" & GetNullString(gilUCCode.GetDescription) & _
                                   IIf(gdaRealDate.GetValue = "", "", ",RealDate=" & GetNullString(gdaRealDate.GetValue(True), giDateV)) & _
                                   ",UpdEn='" & garyGi(1) & "'" & _
                                   ",UpdTime='" & GetDTString(RightNow) & "'"

        Call subChoose
        txtSQL = strUpdate & " Where " & strChoose
        gcnGi.BeginTrans
        If ExecuteSQL(strUpdate & " Where " & strChoose, gcnGi, lngAffectTotal) <> giOK Then MsgBox "異動失敗！請重新操作！", vbExclamation, "訊息！": GoTo ErrTrans
        gcnGi.CommitTrans
        
        lblBillCnt.Caption = lblBillCnt.Caption + lngAffectTotal
        MsgBox "產生完畢！成功筆數: " & lngAffectTotal, vbInformation, gimsgPrompt
        
        Call DefaultValue
        Screen.MousePointer = vbDefault
    Exit Sub
ErrTrans:
        gcnGi.RollbackTrans
        Call DefaultValue
    Exit Sub
chkErr:
    ErrSub Me.Name, "cmdSave_Click"
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
        If Not ChkGiList(KeyCode) Then Exit Sub
        If Shift = 0 Then
            If sstData.Tab = 1 Then
                If KeyCode = vbKeyF2 Then
                    If Not ChkCloseDate(gdaRealDate.GetValue(True), gilCompCode.GetCodeNo, "") Then
                        gdaRealDate.SetFocus
                        Exit Sub
                    End If
                    cmdSave.Value = True
                End If
            End If
        ElseIf Shift = 2 Then
            If KeyCode = vbKeyF Then
                txtSQL.Move 0, 0, Me.Width, Me.Height / 2
                txtSQL.Visible = Not txtSQL.Visible
            End If
                
        End If
End Sub

Private Sub Form_Load()
    On Error Resume Next
        Call subGrd
        Call subGil
        Call subGim
        Call DefaultValue
        Call OpenData
        Call CboAddItem
End Sub

Private Sub OpenData()
    On Error Resume Next
        If Not GetRS(rsBill, strSQL & " From " & GetOwner & "SO033 Where RowId= ''", , adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
        Set rsBill.ActiveConnection = Nothing
        Set ggrData.Recordset = rsBill
        ggrData.Refresh

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

Private Sub subGrd()
    On Error GoTo chkErr
    Dim mFlds As New prjGiGridR.GiGridFlds
        With mFlds
            .Add "EntryNo", , , , , "序號 ", vbLeftJustify
            .Add "CustId", , , , , "客戶編號", vbLeftJustify
            .Add "CitemName", , , , , "項目名稱         ", vbLeftJustify
            .Add "ShouldAmt", , , , , "應收金額", vbRightJustify
            .Add "OldPeriod", , , , , "應收期數", vbLeftJustify
            .Add "RealAmt", , , , , "實收金額", vbRightJustify
            .Add "RealPeriod", , , , , "實收期數", vbLeftJustify
            .Add "RealStartDate", , , , , "有效起始日", vbLeftJustify
            .Add "RealStopDate", , , , , "有效截止日", vbLeftJustify
            .Add "BillNo", , , , , "單據編號          ", vbLeftJustify
            .Add "MediaBillNo", , , , , "媒體編號   ", vbLeftJustify
            .Add "PrtSNo", , , , , "印單序號    ", vbLeftJustify
            .Add "CancelFlag", , , , , "作廢", vbLeftJustify
            .Add "RealDate", giControlTypeDate, , , , "登錄時間", vbLeftJustify
            .Add "CMName", , , , , "收費方式", vbLeftJustify
            .Add "UCName", , , , , "未收原因", vbLeftJustify
            .Add "Note", , , , , "備註" & Space(20), vbLeftJustify
        End With
        ggrData.AllFields = mFlds
        ggrData.SetHead
    Exit Sub
chkErr:
    ErrSub Me.Name, "subGrd"
End Sub

Private Sub subGil()
    On Error Resume Next
        '預設收費方式
        SetgiList gilCMCode, "CodeNO", "Description", "CD031", , , , , , , , True
        gilCMCode.ListIndex = 1
        
        '預設收費人員
        SetgiList gilClctEn, "EmpNO", "EmpName", "CM003", , , , , , , , True
        
        '預設未收原因
        '#3869 過濾只有RefNo=1 Or RefNo is Null By Kin 2008/06/04
        SetgiList gilUCCode, "CodeNo", "Description", "CD013", , , , , , , " Where (RefNO=1 or RefNO is Null) ", True
        '公司別
        SetgiList gilCompCode, "CodeNo", "Description", "CD039", , , , , , , GetCompCodeFilter(gilCompCode)
        
End Sub

Private Sub subGim()
    On Error GoTo chkErr
        SetgiMulti gimCMCode, "CodeNo", "Description", "CD031", "收費方式代碼", "收費方式名稱"
        SetgiMulti gimCitemCode, "CodeNo", "Description", "CD019", "收費項目代碼", "收費項目名稱"
        SetgiMulti gimClassCode, "CodeNo", "Description", "CD004", "客戶類別代碼", "客戶類別名稱"
        SetgiMulti gimMduId, "MduId", "Name", "SO017", "大樓編號", "大樓名稱"
        SetgiMultiAddItem gimBillType, "B,I,P,M,T", "收費單,裝機單,停拆移機單,維修單,臨時收費單"
    Exit Sub
chkErr:
    Call ErrSub(Me.Name, "subGim")
End Sub

 '未收原因
Public Property Let uUCCode(ByVal vData As String)
        strUCCode = vData
End Property

 '未收原因
Public Property Let uUCname(ByVal vData As String)
        strUCName = vData
End Property

'登錄日期！
Public Property Let uDate(ByVal vData As String)
    strDate = vData
End Property

'參數6
Public Property Let uPara6(ByVal vData As Long)
    lngPara6 = vData
End Property

'收費方式
Public Property Let uCMCode(ByVal vData As String)
    strCMCode = vData
End Property

'收費方式
Public Property Let uCMName(ByVal vData As String)
    strCMname = vData
End Property

'收費人員
Public Property Let uClctEn(ByVal vData As String)
    strClctEn = vData
End Property

'收費人員
Public Property Let uClctName(ByVal vData As String)
    strClctName = vData
End Property

'公司別
Public Property Let uCompCode(ByVal vData As String)
    strCompCode = vData
End Property

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If rs.State = adStateOpen Then
        rs.Close
        Set rs = Nothing
    End If
    On Error Resume Next
    'frmSO3410A.blnCanBeClose = True
    'Set ObjopenDb = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
        Call ReleaseCOM(frmSO3410B)
        
End Sub

Private Sub gdaRealDate_Validate(Cancel As Boolean)
    On Error Resume Next
        
        If gdaRealDate.GetValue = "" Then Exit Sub
        Cancel = Not ChkCloseDate(gdaRealDate.GetValue(True), gilCompCode.GetCodeNo, "")
        
        
'        If gdaRealDate.Text = "" Then
'            If vbNo = MsgBox("本欄位是否為空值！", vbExclamation, "訊息！") Then
'                Cancel = True
'            End If
'    End If
'    If Not IsDate(gdaRealDate.Text) Then
'        MsgBox "日期不合法！", vbExclamation, "訊息！"
'        Exit Sub
'    End If
'    If CDate(gdaRealDate.Text) > RightDate Then MsgBox "此日期超過今天日期！", vbExclamation, "訊息！": Exit Sub
'
'    If (RightDate - CDate(gdaRealDate.Text)) > lngPara6 Then MsgBox "此日期已超過系統設定的安全期限！", vbExclamation, "訊息！"
End Sub

Private Sub gdaShouldDate2_GotFocus()
    If Not IsDate(gdaShouldDate1.GetValue) Then Exit Sub
    If gdaShouldDate2.GetValue = "" Then gdaShouldDate2.SetValue gdaShouldDate1.GetValue(True)
End Sub

Private Sub gdtCreatTime1_Validate(Cancel As Boolean)
    On Error Resume Next
    If gdtCreatTime1.GetValue & "" <> "" Then
        If gdtCreatTime2.GetValue & "" = "" Then
            gdtCreatTime2.SetValue gdtCreatTime1.GetDate & "2359"
        Else
            If gdtCreatTime1.GetValue > gdtCreatTime2.GetValue Then
                MsgBox "起始日期不得大於終止日期!!", vbInformation, "警告訊息"
                Cancel = True
            End If
        End If
    Else
        If gdtCreatTime2.GetValue & "" <> "" Then
            MsgBox "終止日期有值，起始日期也必須有值!!", vbInformation, "警告訊息"
            Cancel = True
        End If
    End If
End Sub

Private Sub gdtCreatTime2_Validate(Cancel As Boolean)
    On Error Resume Next
    If gdtCreatTime2.GetValue & "" = "" Then
        If gdtCreatTime1.GetValue & "" <> "" Then
            MsgBox "起始日期有值，終止日期也必須有值!!", vbInformation, "警告訊息"
            Cancel = True
        End If
    Else
        If gdtCreatTime1.GetValue & "" = "" Then
            MsgBox "起始日期有值，終止日期也必須有值!!", vbInformation, "警告訊息"
            gdtCreatTime1.SetFocus
            Exit Sub
        End If
        If gdtCreatTime1.GetValue > gdtCreatTime2.GetValue Then
            MsgBox "起始日期不得大於終止日期!!", vbInformation, "警告訊息"
            Cancel = True
        End If
    End If
End Sub

Private Sub ggrData_ShowCellData(ByVal Col As Long, ByVal Row As Long, ByVal giFld As prjGiGridR.GiGridFld, ByVal Fld As ADODB.Field, Value As Variant)
    Select Case UCase(Fld.Name)
        Case "REALAMT", "SHOULDAMT"
            Value = Format(Value, "###,###,###")
        Case "REALSTARTDATE", "REALSTOPDATE"
            'Value = Format(Value, "EE/MM/DD")
            Dim dtString As String
            dtString = Value
            Value = GetDTString(dtString)
    End Select
End Sub

Private Sub gilCompCode_Change()
    On Error Resume Next
        If Not ChgComp(gilCompCode, "SO3400", "SO3410") Then Exit Sub
        Call subGil
        'Call GiListFilter(gilUCCode, , gilCompCode.GetCodeNo)
End Sub

Private Sub optAll_Click()
    txtPrtSNo1.Enabled = True
    txtPrtSNo2.Enabled = True

End Sub

Private Sub optNo_Click()
    txtPrtSNo1.Text = ""
    txtPrtSNo2.Text = ""
    txtPrtSNo1.Enabled = False
    txtPrtSNo2.Enabled = False
End Sub

Private Sub optYES_Click()
    txtPrtSNo1.Enabled = True
    txtPrtSNo2.Enabled = True

End Sub

Private Sub sstData_Click(PreviousTab As Integer)
    If sstData.Tab = 0 Then
        cmdSave.Enabled = False
    Else
        cmdSave.Enabled = True
    End If
End Sub
'
'Private Sub txtBillNo_Change()
'    On Error GoTo ChkErr
'    Dim rsTmp As New ADODB.Recordset
'    Dim strTmp As String
'        If Len(txtBillNo.Text) < txtBillNo.MaxLength Then Exit Sub
'        Select Case cboBillType.ListIndex
'            Case 0
'            If Not GetRS(rsBill, strSQL & " From " & GetOwner & "So033 Where Billno='" & Trim(txtBillNo.Text) & "' And CompCode = " & gilCompCode.GetCodeNo & " And UCCode is Not Null", gcnGi, adUseClient) Then MsgBox "連線失敗！請重新操作！", vbExclamation, "訊息！": Exit Sub
'            Case 1
'                If Not GetRS(rsBill, strSQL & " From " & GetOwner & "So033 Where PrtSNo='" & Trim(txtBillNo.Text) & "' And CompCode = " & gilCompCode.GetCodeNo & " And UCCode is Not Null", gcnGi, adUseClient) Then MsgBox "連線失敗！請重新操作！", vbExclamation, "訊息！": Exit Sub
'            Case 2
'                If Not GetRS(rsBill, strSQL & " From " & GetOwner & "So033 Where MediaBillNo='" & Trim(txtBillNo.Text) & "' And CompCode = " & gilCompCode.GetCodeNo & " And UCCode is Not Null", gcnGi, adUseClient) Then MsgBox "連線失敗！請重新操作！", vbExclamation, "訊息！": Exit Sub
'        End Select
'        If rsBill.EOF Then
'            MsgBox "無此單據編號或此單據已繳費，請核對！", vbExclamation, "訊息！"
'            rsBill.Close
'            txtBillNo.Text = ""
'            txtBillNo.SetFocus
'            Exit Sub
'        End If
'        strTmp = "Select CustName From " & GetOwner & "So001 Where Custid=" & rsBill("Custid").Value
'        If Not GetRS(rsTmp, strTmp, gcnGi, adUseClient, adOpenForwardOnly, adLockReadOnly) Then MsgBox "連線失敗！請重新操作！", vbExclamation, "訊息！"
'        If rsTmp.EOF Then
'            MsgBox "無此單據之客戶資料！", vbExclamation, "訊息！"
'            txtBillNo.Text = ""
'            txtBillNo.SetFocus
'            rsTmp.Close
'            Set rsTmp = Nothing
'            Exit Sub
'        End If
'        txtNote = rsBill("Note") & ""
'        lblName.Caption = rsTmp("CustName").Value
'        lblCustId.Caption = rsBill("CustID").Value
'        cmdSave.Enabled = True
'        Set ggrData.Recordset = rsBill
'        ggrData.Refresh
'        If chkAutoSave.Value = 1 Then cmdSave.Value = True
'    Exit Sub
'ChkErr:
'    ErrSub Me.Name, "txtBillNo_Change"
'End Sub

Private Sub txtBillNo_Change()
    On Error GoTo chkErr
    Dim rsTmp As New ADODB.Recordset
    Dim strTime As String
    Dim lngAffectRow As Long, lngAffectTotal As Long
    Dim strUpdate As String
    Dim rsSO074A As New ADODB.Recordset
        If txtBillNo = "" Then Exit Sub
        If Len(txtBillNo.Text) < txtBillNo.MaxLength And Not (chkUseOldBillNo.Value = 1 And Len(txtBillNo) = 11) Then Exit Sub
        If Not IsDataOk Then Exit Sub
        Select Case cboBillType.ListIndex
            Case 0
                If Len(txtBillNo) = 11 And chkUseOldBillNo.Value = 1 Then
                    If InStr("BT", Mid(txtBillNo, 5, 1)) > 0 Then
                        txtBillNo = GetADString(Left(txtBillNo, 4), False) & Mid(txtBillNo, 5, 1) & "C" & Format(Mid(txtBillNo, 6), "0000000")
                        Exit Sub
                    End If
                End If
                If Not GetRS(rsTmp, strSQL & " From " & GetOwner & "So033 Where Billno='" & Trim(txtBillNo.Text) & "' And CompCode = " & gilCompCode.GetCodeNo & " And UCCode is Not Null ", gcnGi, adUseClient) Then MsgBox "連線失敗！請重新操作！", vbExclamation, "訊息！": Exit Sub
            Case 1
                If Not GetRS(rsTmp, strSQL & " From " & GetOwner & "So033 Where PrtSNo='" & Trim(txtBillNo.Text) & "' And CompCode = " & gilCompCode.GetCodeNo & " And UCCode is Not Null", gcnGi, adUseClient) Then MsgBox "連線失敗！請重新操作！", vbExclamation, "訊息！": Exit Sub
            Case 2
                If Not GetRS(rsTmp, strSQL & " From " & GetOwner & "So033 Where MediaBillNo='" & Trim(txtBillNo.Text) & "' And CompCode = " & gilCompCode.GetCodeNo & " And UCCode is Not Null", gcnGi, adUseClient) Then MsgBox "連線失敗！請重新操作！", vbExclamation, "訊息！": Exit Sub
        End Select
        If rsTmp.EOF Then
            MsgBox "無此單據編號或此單據已繳費，請核對！", vbExclamation, "訊息！"
            rsTmp.Close
            txtBillNo.Text = ""
            txtBillNo.SetFocus
            Exit Sub
        Else
            If GetRefNoByCode("CD013", rsTmp("UCCode")) = 3 Then
                If Not GetRS(rsSO074A, "Select EntryEn,RealDate From " & GetOwner & "SO074A Where BillNo = '" & rsTmp("BillNo") & "' And Item = " & rsTmp("Item")) Then Exit Sub
                If Not rsSO074A.EOF Then
                    MsgBox "本單據已於服務櫃檯入帳成功,入帳日期:" & GetDT(rsSO074A("RealDate"), GiDate) & ",入帳人員:" & rsSO074A("EntryEn"), vbExclamation, gimsgPrompt
                    
                    'MsgBox "此單據已櫃台繳費，請核對！", vbExclamation, "訊息！"
                    rsTmp.Close
                    txtBillNo.Text = ""
                    txtBillNo.SetFocus
                    Exit Sub
                End If
            End If
        End If
        strTime = GetDTString(RightNow)
        If gilClctEn.GetCodeNo <> "" Then
            strUpdate = "ClctEn='" & gilClctEn.GetCodeNo & "',ClctName='" & gilClctEn.GetDescription & "',"
        End If
        
        '#3869 當登錄日期有值才填入實收日期 By Kin 2008/06/04
        strUpdate = "Update " & GetOwner & "So033 Set  " & _
                                    strUpdate & "CMCode=" & GetNullString(gilCMCode.GetCodeNo) & _
                                   ",CMName=" & GetNullString(gilCMCode.GetDescription) & _
                                   ",UCCode=" & GetNullString(gilUCCode.GetCodeNo) & _
                                   ",UCName=" & GetNullString(gilUCCode.GetDescription) & _
                                   IIf(gdaRealDate.GetValue = "", "", ",RealDate=" & GetNullString(gdaRealDate.GetValue(True), giDateV)) & _
                                   ",UpdEn='" & garyGi(1) & "'" & _
                                   ",UpdTime='" & strTime & "'"
        On Error GoTo ErrTrans
        gcnGi.BeginTrans
        rsTmp.MoveFirst
        While Not rsTmp.EOF
                If ExecuteSQL(strUpdate & ",Note = Note || '" & txtNote & "' Where Rowid='" & rsTmp("RowID").Value & "'", gcnGi, lngAffectRow, , False) <> giOK Then MsgBox "異動失敗！請重新操作！", vbExclamation, "訊息！": GoTo ErrTrans
                               lngAffectTotal = lngAffectTotal + 1
            rsTmp.MoveNext
        Wend
        lngBillcount = lngBillcount + 1
        rsTmp.Requery
        Do While Not rsTmp.EOF
            CopyRecordset rsTmp, rsBill, False, False
            rsBill("EntryNo") = lngBillcount
            rsTmp.MoveNext
        Loop
        txtNote = rsBill("Note") & ""
        lblCustId.Caption = rsBill("CustID").Value
        
        Set ggrData.Recordset = rsBill
        ggrData.Refresh
        ggrData.Sort "EntryNo Desc"
        
        '#2908 資料輸入後,將指標移至Grid最上方 By Kin 2007/04/02
        ggrData.GotoRow 1
        
        gcnGi.CommitTrans
        
        lblBillCnt.Caption = lblBillCnt.Caption + lngAffectTotal
        'If chkAutoSave.Value = 0 Then MsgBox "產生完畢！成功筆數: " & lngAffectTotal, vbInformation, gimsgPrompt
        
        Call DefaultValue
        Call CloseRecordset(rsTmp)
        Screen.MousePointer = vbDefault
    Exit Sub
ErrTrans:
        gcnGi.RollbackTrans
        Call DefaultValue
    Exit Sub
chkErr:
    ErrSub Me.Name, "txtBillNo_Change"
End Sub

Private Sub txtBillNo_GotFocus()
    txtBillNo.SelStart = 0
    txtBillNo.SelLength = Len(txtBillNo.Text)
End Sub

Private Sub txtBillNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtBillNo_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub DefaultValue()
    On Error Resume Next
        strSQL = "Select RowId,CitemName,ShouldAmt,OldPeriod,RealAmt,RealPeriod,RealStartDate,RealStopDate,CancelFlag,RealDate," & _
                        "Note,CustId,UCCode,UCName,CmCode,CMName,ClctEn,ClctName,CustId,BillNo,Item,PrtSNo,MediaBillno,CustId EntryNo"
        intPara23 = Val(GetSystemParaItem("Para23", gilCompCode.GetCodeNo, "", "SO043") & "")
        intPara24 = Val(GetSystemParaItem("Para24", gilCompCode.GetCodeNo, "", "SO043") & "")
        lblCustId.Caption = ""
        lblName.Caption = ""
        gilCompCode.SetCodeNo strCompCode
        gilCompCode.Query_Description
        '收費人員 gilClctEn' strClctEn,strClctName
        gilClctEn.SetCodeNo (strClctEn)
        gilClctEn.SetDescription (strClctName)
        '登錄日期 gdaRealDate strDate
        gdaRealDate.SetValue (strDate)
        '收費方式 gilCMCode strCMCode,strCMName
        gilCMCode.SetCodeNo (strCMCode)
        gilCMCode.SetDescription (strCMname)
        '未收原因 gilUCCode strUCCode,strUCName
        gilUCCode.SetCodeNo (strUCCode)
        gilUCCode.SetDescription (strUCName)
        strOldBillNo = ""
        txtBillNo.Text = ""
        txtNote.Text = ""
        txtBillNo.SetFocus
        sstData.Tab = 0
        If intPara23 = 1 Then
            cboBillType.ListIndex = 1
            txtBillNo.MaxLength = 12
            'If cboBillType.ListCount > 2 Then cboBillType.RemoveItem 2
        ElseIf intPara23 = 2 Then
            cboBillType.ListIndex = 2
            txtBillNo.MaxLength = 11
            'cboBillType.Visible = False
            'txtBillNo.Move gilClctEn.Left, txtBillNo.Top
        Else
            cboBillType.ListIndex = 0
            txtBillNo.MaxLength = 15
            'If cboBillType.ListCount > 2 Then cboBillType.RemoveItem 2
        End If
        
        If intPara24 = 1 Then
            lblBillHeadFmt.Visible = True
            cboBillHeadFmt.Visible = True
            '  用媒體入帳才Lock
            If intPara23 = 2 Then gimCitemCode.Enabled = False
        End If
        
        cmdSave.Enabled = False
End Sub

Private Function ChkDupData(ByVal strRowId As String) As Boolean
Dim rsDup1 As New ADODB.Recordset
ChkDupData = True
    If rs.RecordCount > 0 Then
        Set rsDup1 = rs.Clone
        rsDup1.MoveFirst
        While Not rsDup1.EOF
            If rsDup1("RowId").Value = strRowId Then
                rsDup1.Close: Set rsDup1 = Nothing
                Exit Function
            End If
            rsDup1.MoveNext
        Wend
    End If
    If rsDup1.State = adStateOpen Then rsDup1.Close
    Set rsDup1 = Nothing
ChkDupData = False
End Function
