VERSION 5.00
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{0CE6E756-92AE-11D3-8311-0080C8453DDF}#5.0#0"; "GiYM.ocx"
Object = "{A0C51418-555F-11D3-BC10-0050BA01A3A9}#9.9#0"; "GiMulti.ocx"
Object = "{6DDA0569-6454-11D3-9612-0048544B09DF}#4.1#0"; "GiTime.ocx"
Object = "{074771B7-6ADE-11D5-B178-00485456B028}#2.1#0"; "csMulti.ocx"
Begin VB.Form frmSO3220A 
   BorderStyle     =   1  '單線固定
   Caption         =   "印單序號給號作業 [SO3220A]"
   ClientHeight    =   7425
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
   Icon            =   "SO3220A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7425
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton cmdOk 
      Caption         =   "F2. 確定"
      Default         =   -1  'True
      Height          =   375
      Left            =   8790
      TabIndex        =   24
      Top             =   6960
      Width           =   1365
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&X)"
      Height          =   375
      Left            =   10410
      TabIndex        =   25
      Top             =   6960
      Width           =   1365
   End
   Begin VB.Frame fraData 
      Height          =   6600
      Left            =   180
      TabIndex        =   26
      Top             =   120
      Width           =   11655
      Begin CS_Multi.CSmulti gimCitemCode 
         Height          =   315
         Left            =   180
         TabIndex        =   47
         Top             =   1680
         Width           =   11355
         _ExtentX        =   20029
         _ExtentY        =   556
         ButtonCaption   =   "收費項目"
         IsReadOnly      =   -1  'True
      End
      Begin Gi_Multi.GiMulti gimBillType 
         Height          =   345
         Left            =   180
         TabIndex        =   46
         Top             =   5280
         Width           =   11355
         _ExtentX        =   20029
         _ExtentY        =   609
         ButtonCaption   =   "單據類別"
         DataType        =   2
         DIY             =   -1  'True
         FontSize        =   9
         FontName        =   "新細明體"
      End
      Begin VB.ComboBox cboBillHeadFmt 
         Height          =   315
         Left            =   6180
         Style           =   2  '單純下拉式
         TabIndex        =   3
         Top             =   900
         Visible         =   0   'False
         Width           =   5295
      End
      Begin CS_Multi.CSmulti gimClassCode 
         Height          =   375
         Left            =   180
         TabIndex        =   10
         Top             =   2040
         Width           =   11355
         _ExtentX        =   20029
         _ExtentY        =   661
         ButtonCaption   =   "客戶類別"
      End
      Begin CS_Multi.CSmulti gimCreateEn 
         Height          =   375
         Left            =   180
         TabIndex        =   16
         Top             =   4200
         Width           =   11355
         _ExtentX        =   20029
         _ExtentY        =   661
         ButtonCaption   =   "產生人員"
      End
      Begin CS_Multi.CSmulti gmsClctEn 
         Height          =   375
         Left            =   180
         TabIndex        =   11
         Top             =   2400
         Width           =   11355
         _ExtentX        =   20029
         _ExtentY        =   661
         ButtonCaption   =   "收費人員"
      End
      Begin VB.TextBox txtOrder 
         Height          =   285
         Left            =   7740
         MaxLength       =   6
         TabIndex        =   20
         Text            =   "1432"
         Top             =   5640
         Width           =   630
      End
      Begin Gi_Multi.GiMulti gimCMCode 
         Height          =   375
         Left            =   180
         TabIndex        =   15
         Top             =   3840
         Width           =   11355
         _ExtentX        =   20029
         _ExtentY        =   661
         ButtonCaption   =   "收費方式"
         FontSize        =   9
         FontName        =   "新細明體"
      End
      Begin prjGiList.GiList gilServiceType 
         Height          =   315
         Left            =   5220
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
         F2Corresponding =   -1  'True
      End
      Begin prjGiList.GiList gilCompCode 
         Height          =   315
         Left            =   1140
         TabIndex        =   0
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
         F2Corresponding =   -1  'True
      End
      Begin VB.ComboBox cboRange 
         Height          =   315
         ItemData        =   "SO3220A.frx":0442
         Left            =   765
         List            =   "SO3220A.frx":0455
         Style           =   2  '單純下拉式
         TabIndex        =   21
         Top             =   6075
         Width           =   1935
      End
      Begin VB.CheckBox chkClctType2 
         Caption         =   "&1.待拆客戶需要給號"
         Height          =   195
         Left            =   3390
         TabIndex        =   22
         Top             =   6150
         Width           =   2310
      End
      Begin VB.CheckBox chkClctType1 
         Caption         =   "&2.前期欠費客戶需要給號"
         Height          =   300
         Left            =   6915
         TabIndex        =   23
         Top             =   6090
         Width           =   2490
      End
      Begin VB.ComboBox cboNoe1 
         Height          =   315
         ItemData        =   "SO3220A.frx":0494
         Left            =   765
         List            =   "SO3220A.frx":04A7
         Style           =   2  '單純下拉式
         TabIndex        =   19
         Top             =   5640
         Width           =   1125
      End
      Begin Gi_YM.GiYM gymClctYM1 
         Height          =   330
         Left            =   930
         TabIndex        =   4
         Top             =   1260
         Width           =   1020
         _ExtentX        =   1799
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
      Begin Gi_YM.GiYM gymPrtYM 
         Height          =   300
         Left            =   3045
         TabIndex        =   2
         Top             =   870
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   529
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
      Begin Gi_YM.GiYM gymClctYM2 
         Height          =   330
         Left            =   2340
         TabIndex        =   5
         Top             =   1260
         Width           =   990
         _ExtentX        =   1746
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
      Begin CS_Multi.CSmulti gmsStrtCode 
         Height          =   375
         Left            =   180
         TabIndex        =   18
         Top             =   4920
         Width           =   11355
         _ExtentX        =   20029
         _ExtentY        =   661
         ButtonCaption   =   "街道範圍"
      End
      Begin Gi_Multi.GiMulti gmsServCode 
         Height          =   375
         Left            =   180
         TabIndex        =   12
         Top             =   2760
         Width           =   11355
         _ExtentX        =   20029
         _ExtentY        =   661
         ButtonCaption   =   "服  務  區"
         FontSize        =   9
         FontName        =   "新細明體"
      End
      Begin Gi_Multi.GiMulti gmsClctAreaCode 
         Height          =   375
         Left            =   180
         TabIndex        =   13
         Top             =   3120
         Width           =   11355
         _ExtentX        =   20029
         _ExtentY        =   661
         ButtonCaption   =   "收  費  區"
         FontSize        =   9
         FontName        =   "新細明體"
      End
      Begin CS_Multi.CSmulti gmsMduId 
         Height          =   375
         Left            =   180
         TabIndex        =   17
         Top             =   4560
         Width           =   11355
         _ExtentX        =   20029
         _ExtentY        =   661
         ButtonCaption   =   "大樓名稱"
      End
      Begin Gi_Date.GiDate gdaShouldDay1 
         Height          =   315
         Left            =   4140
         TabIndex        =   6
         Top             =   1290
         Width           =   1065
         _ExtentX        =   1879
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
      Begin Gi_Date.GiDate gdaShouldDay2 
         Height          =   315
         Left            =   5430
         TabIndex        =   7
         Top             =   1290
         Width           =   1065
         _ExtentX        =   1879
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
      Begin Gi_Multi.GiMulti gimUCCode 
         Height          =   375
         Left            =   180
         TabIndex        =   14
         Top             =   3480
         Width           =   11355
         _ExtentX        =   20029
         _ExtentY        =   661
         ButtonCaption   =   "未收原因"
         FontSize        =   9
         FontName        =   "新細明體"
      End
      Begin Gi_Time.GiTime gdaCreateTime1 
         Height          =   315
         Left            =   7290
         TabIndex        =   8
         Top             =   1290
         Width           =   1965
         _ExtentX        =   3466
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
      Begin Gi_Time.GiTime gdaCreateTime2 
         Height          =   315
         Left            =   9510
         TabIndex        =   9
         Top             =   1290
         Width           =   1965
         _ExtentX        =   3466
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
      Begin VB.Label lblShouldDay 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "應收日期"
         Height          =   195
         Left            =   3360
         TabIndex        =   45
         Top             =   1350
         Width           =   780
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "至"
         Height          =   195
         Left            =   5220
         TabIndex        =   44
         Top             =   1365
         Width           =   195
      End
      Begin VB.Label lblBillHeadFmt 
         AutoSize        =   -1  'True
         Caption         =   "多帳戶產生依據設定"
         Height          =   195
         Left            =   4380
         TabIndex        =   43
         Top             =   930
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.Label lblD3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "至"
         Height          =   195
         Left            =   9270
         TabIndex        =   39
         Top             =   1350
         Width           =   195
      End
      Begin VB.Label lblCreateTime 
         AutoSize        =   -1  'True
         Caption         =   "產生日期"
         Height          =   195
         Left            =   6510
         TabIndex        =   38
         Top             =   1350
         Width           =   780
      End
      Begin VB.Label lblServiceType 
         AutoSize        =   -1  'True
         Caption         =   "服務類別"
         Height          =   195
         Left            =   4380
         TabIndex        =   37
         Top             =   300
         Width           =   780
      End
      Begin VB.Label lblCompCode 
         AutoSize        =   -1  'True
         Caption         =   "公司別"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   330
         TabIndex        =   36
         Top             =   330
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "範圍"
         Height          =   195
         Left            =   330
         TabIndex        =   32
         Top             =   6120
         Width           =   390
      End
      Begin VB.Label lblOrder 
         AutoSize        =   -1  'True
         Caption         =   "排序(1.收費員 2.街道 3.服務區 4.大樓編號 5.收費區 6.網路編號)"
         Height          =   195
         Left            =   2430
         TabIndex        =   31
         Top             =   5685
         Width           =   5250
      End
      Begin VB.Label lblNoe1 
         AutoSize        =   -1  'True
         Caption         =   "門號"
         Height          =   195
         Left            =   330
         TabIndex        =   30
         Top             =   5715
         Width           =   390
      End
      Begin VB.Label lblNote2B 
         AutoSize        =   -1  'True
         Caption         =   "至"
         Height          =   195
         Left            =   2100
         TabIndex        =   29
         Top             =   1350
         Width           =   195
      End
      Begin VB.Label lblNote2A 
         AutoSize        =   -1  'True
         Caption         =   "收集年月"
         Height          =   195
         Left            =   90
         TabIndex        =   28
         Top             =   1350
         Width           =   780
      End
      Begin VB.Label lblNote1 
         AutoSize        =   -1  'True
         Caption         =   "印單年月(處理本年月未編號部分)"
         Height          =   195
         Left            =   60
         TabIndex        =   27
         Top             =   930
         Width           =   2850
      End
   End
   Begin VB.Label lblPrtSno1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  '單線固定
      Height          =   285
      Left            =   1920
      TabIndex        =   42
      Top             =   6990
      Width           =   1320
   End
   Begin VB.Label lblNote5B 
      Caption         =   "至"
      Height          =   195
      Left            =   3315
      TabIndex        =   41
      Top             =   7020
      Width           =   195
   End
   Begin VB.Label lblPrtSno2 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  '單線固定
      Height          =   285
      Left            =   3570
      TabIndex        =   40
      Top             =   6990
      Width           =   1335
   End
   Begin VB.Label lblSNoCnt 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  '單線固定
      Height          =   255
      Left            =   7080
      TabIndex        =   35
      Top             =   6990
      Width           =   780
   End
   Begin VB.Label lblNote6 
      Caption         =   "合計張數："
      Height          =   195
      Left            =   6090
      TabIndex        =   34
      Top             =   7050
      Width           =   975
   End
   Begin VB.Label lblNote5 
      AutoSize        =   -1  'True
      Caption         =   "產生之印單序號："
      Height          =   195
      Left            =   270
      TabIndex        =   33
      Top             =   7050
      Width           =   1560
   End
End
Attribute VB_Name = "frmSO3220A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Programmer:  Corey HSU
' Last Modify: 2000/1/3
Option Explicit
Dim rsBillHeadFmt As New ADODB.Recordset
Dim intPara24 As Integer
Dim strOrderPath As String
Dim strOrderFileName As String
Private Sub cboBillHeadFmt_Click()
    On Error Resume Next
    Dim strCitemCode As String
    Dim rsCD019  As New ADODB.Recordset
    Dim strCD019SQL As String
        If rsBillHeadFmt.RecordCount = 0 Then Exit Sub
        rsBillHeadFmt.AbsolutePosition = cboBillHeadFmt.ListIndex + 1
        strCD019SQL = "Select CodeNo From " & GetOwner & "CD019 " & _
                            " Where CodeNo In ( Select CitemCode From " & GetOwner & "CD068A " & _
                            "         Where Exists (Select * From " & GetOwner & "CD068 Where CD068.BillHeadFmt = CD068A.BillHeadFmt " & _
                                                    " And CD068.BillHeadFmt = '" & rsBillHeadFmt("BillHeadFmt") & "' )) And StopFlag <>1"
        'If Not GetRS(rsCD019, "Select CodeNo From " & GetOwner & "CD019 Where CodeNo In (" & rsBillHeadFmt("CitemCodeStr") & ") And StopFlag=0") Then Exit Sub
        If Not GetRS(rsCD019, strCD019SQL) Then Exit Sub
        If rsCD019.EOF Then Exit Sub
        strCitemCode = rsCD019.GetString(, , , ",")
        strCitemCode = Left(strCitemCode, Len(strCitemCode) - 1)
        'gimCitemCode.Filter = gimCitemCode.Filter
        gimCitemCode.SetQueryCode strCitemCode
        'gimCitemCode.ShowMulti
End Sub

Private Sub cmdCancel_Click()
    On Error GoTo ChkErr
        Unload Me
    Exit Sub
ChkErr:
    ErrSub Me.Name, "cmdCancel_Click"
End Sub

Private Sub Form_Activate()
    On Error GoTo ChkErr
        Screen.MousePointer = vbDefault
    Exit Sub
ChkErr:
    ErrSub Me.Name, "Form_Activate"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo ChkErr
        If Not ChkGiList(KeyCode) Then Exit Sub
        If KeyCode = vbKeyF2 Then
            If cmdOk.Enabled Then Call cmdOk_Click
            KeyCode = 0
        End If
    Exit Sub
ChkErr:
    ErrSub Me.Name, "Form_KeyDown"
End Sub

Private Sub Form_Load()
    On Error GoTo ER
        strOrderFileName = "SO3220RC.TXT"
        strOrderPath = ReadGICMIS1("ErrLogPath") & ""
        '*****************************************************
        '#2672 產生時間要到時分秒 By Kin 2008/05/12
        gdaCreateTime1.ShowSecond
        gdaCreateTime2.ShowSecond
        '******************************************************
        subGim
        subGil
        DefaultValue
        CboAddItem
    Exit Sub
ER:
    If ErrHandle(, True, , "Form_Load") Then Resume 0
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

Private Sub subGil()
    On Error GoTo ChkErr
        SetgiList gilServiceType, "CodeNo", "Description", "CD046", , , , , 1, 12
        SetgiList gilCompCode, "CodeNo", "Description", "CD039", , , , , 3, 12, GetCompCodeFilter(gilCompCode)
    Exit Sub
ChkErr:
    ErrSub Me.Name, "subGil"
End Sub

Private Sub DefaultValue()
    On Error GoTo ChkErr
    Dim intPara23 As Integer
        gilCompCode.SetCodeNo garyGi(9)
        gilCompCode.Query_Description
        ' 門號 = 0(所有)
        cboNoe1.ListIndex = 0
        ' 範圍 = 0(所有)
        cboRange.ListIndex = 0
        
        ' 出單年月 = 當日之年月
        gymPrtYM.SetValue (Format(RightNow, "YYYYMM"))
        
        ' 收集年月範圍 = 當日之年月至當日之年月
        gymClctYM1.SetValue (Format(RightNow, "YYYYMM"))
        gymClctYM2.SetValue (Format(RightNow, "YYYYMM"))
        
        ' 排序 = '132'
        'txtOrder.Text = "1432"
        '#2923將記錄顯示於TxtOrder
        txtOrder = Replace(RcdToTxt, vbCrLf, "")
        
        ' 前期欠費是否重編:      Checked
        chkClctType1.Value = 1
        
        ' 待拆客戶是否重編:  not checked
        chkClctType2.Value = 0
        
        ' 產生之印單序號範圍 , 合計張數: 皆為空白
        lblPrtSno1.Caption = ""
        lblPrtSno2.Caption = ""
        lblSNoCnt.Caption = ""
        intPara24 = GetSystemParaItem("Para24", gilCompCode.GetCodeNo, gilServiceType.GetCodeNo, "SO043")
        intPara23 = Val(GetSystemParaItem("Para23", gilCompCode.GetCodeNo, gilServiceType.GetCodeNo, "SO043") & "")
        If intPara24 = 1 Then
            lblBillHeadFmt.Visible = True
            cboBillHeadFmt.Visible = True
            '  用媒體入帳才Lock
            'If intPara23 = 2 Then gimCitemCode.Enabled = False
            gilServiceType.SetCodeNo ""
            gilServiceType.Query_Description
            gilServiceType.Enabled = False
            gilServiceType.Visible = False
            lblServiceType.Visible = False
        End If
    Exit Sub
ChkErr:
    ErrSub Me.Name, "DefalutValue"
End Sub

Private Sub subGim()
    On Error GoTo ChkErr
'        SetgiMulti gmsCompCode, "CodeNo", "Description", "CD039", "公司代碼", "公司名稱"
        SetgiMulti gimCitemCode, "CodeNo", "Description", "CD019", "收費項目代碼", "收費項目名稱", , True
        SetgiMulti gimCMCode, "CodeNo", "Description", "CD031", "收費方式代碼", "收費方式名稱", , True
        SetgiMulti gimUCCode, "CodeNo", "Description", "CD013", "代碼", "名稱", , True
        SetgiMulti gimClassCode, "CodeNo", "Description", "CD004", "客戶類別代碼", "客戶類別名稱", , True
        SetgiMulti gmsClctEn, "EmpNo", "EmpName", "CM003", "人員代碼", "收費人員", , True
        SetgiMulti gimCreateEn, "EmpNo", "EmpName", "CM003", "人員代碼", "產生人員", , True
        SetgiMulti gmsServCode, "CodeNo", "Description", "CD002", "服務區代碼", "服務區", , True
        SetgiMulti gmsClctAreaCode, "CodeNo", "Description", "CD040", "收費區代碼", "收費區", , True
        SetgiMulti gmsMduId, "MduID", "Name", "SO017", "大樓編號", "大樓名稱"
        SetgiMulti gmsStrtCode, "CodeNo", "Description", "CD017", "街道代碼", "街道名稱", , True
        '問題集2673 增加單據類別
        SetgiMultiAddItem gimBillType, "B,I,P,M,T", "收費單,裝機單,停拆移機單,維修單,臨時收費單", "單據代號", "單據名稱"
        gimBillType.SetDispStr "收費單"
        gimBillType.SetQueryCode "B"
    Exit Sub
ChkErr:
    ErrSub Me.Name, "subGim"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
        Call CloseRecordset(rsBillHeadFmt)
        Call ReleaseCOM(Me)
End Sub

Private Sub gilCompCode_Change()
    On Error Resume Next
         If Not ChgComp(gilCompCode, "SO3200", "SO3220") Then Exit Sub
        Call subGil
        Call subGim
        Call SelectServType(gilCompCode.GetCodeNo, gilServiceType)
        gilServiceType.SetCodeNo ""
        gilServiceType.Query_Description
        gilServiceType.ListIndex = 1
        
'        gimBankId.Filter = gimBankId.Filter & IIf(gimBankId.Filter = "", " Where ", " And ") & " Upper(PrgName) = 'BANKCENTER'"
        Call GiMultiFilter(gmsMduId, , gilCompCode.GetCodeNo)
        Call GiMultiFilter(gmsServCode, , gilCompCode.GetCodeNo)
        Call GiMultiFilter(gmsClctAreaCode, , gilCompCode.GetCodeNo)
        Call GiMultiFilter(gmsStrtCode, , gilCompCode.GetCodeNo)
End Sub

Private Sub gilServiceType_Change()
    On Error Resume Next
        'If gilServiceType.GetCodeNo = "" Then Exit Sub
        Call GiMultiFilter(gimCitemCode, gilServiceType.GetCodeNo)
        'gimCitemCode.Filter = gimCitemCode.Filter & IIf(Trim(gimCitemCode.Filter) = "", " Where ", " And ") & " PeriodFlag = 1"
        Call GiMultiFilter(gimClassCode, gilServiceType.GetCodeNo)
        Call GiMultiFilter(gimCMCode, gilServiceType.GetCodeNo)
        
End Sub

Private Sub gymClctYM1_Validate(Cancel As Boolean)
    On Error GoTo ChkErr
        Call CheckYM
    Exit Sub
ChkErr:
    ErrSub Me.Name, "gymClctYM1_Validate"
End Sub

Private Sub gymClctYM2_Validate(Cancel As Boolean)
    On Error GoTo ChkErr
        Call CheckYM
    Exit Sub
ChkErr:
    ErrSub Me.Name, "gymClctYM2_Validate"
End Sub

Private Sub txtOrder_GotFocus()
    On Error GoTo ChkErr
        txtOrder.SelStart = 0
        txtOrder.SelLength = Len(txtOrder.Text)
    Exit Sub
ChkErr:
    ErrSub Me.Name, "txtOrder_GotFocus"
End Sub

Private Sub txtOrder_KeyPress(KeyAscii As Integer)
    On Error GoTo ChkErr
        If KeyAscii >= Asc("1") And KeyAscii <= Asc("6") Then
            If txtOrder.SelLength > 0 Then Exit Sub
            If Len(txtOrder.Text) >= 6 Then KeyAscii = 0
        Else
            KeyAscii = 0
        End If
    Exit Sub
ChkErr:
    ErrSub Me.Name, "txtOrder_KeyPress"
End Sub
'
'Private Sub txtOrder_Validate(Cancel As Boolean)
'    On Error GoTo ChkErr
'    Dim lngX As Long
'        txtOrder.Text = Left(LTrim(txtOrder), 4)
'        For lngX = 1 To 4
'            If Val(Mid(txtOrder, lngX, 1)) < 1 Or Val(Mid(txtOrder, lngX, 1)) > 4 Then
'                MsgBox "輸入的排序字串不合法。", vbInformation, "訊息"
'                Cancel = True
'                txtOrder.Text = "1432"
'                Exit Sub
'            End If
'        Next
'    Exit Sub
'ChkErr:
'    ErrSub Me.Name, "txtOrder_Validate"
'End Sub

Private Sub CheckYM()
    On Error GoTo ChkErr
        If gymClctYM1.GetValue <> "" And gymClctYM2.GetValue <> "" Then
            If gymClctYM2.GetValue < gymClctYM1.GetValue Then
                gymClctYM2.SetValue (gymClctYM1.GetValue)
            End If
        End If
        If gymClctYM1.GetValue <> "" And gymClctYM2.GetValue = "" Then
            gymClctYM2.SetValue (gymClctYM1.GetValue)
        End If
        If gymClctYM2.GetValue <> "" And gymClctYM1.GetValue = "" Then
            gymClctYM1.SetValue (gymClctYM2.GetValue)
        End If
    Exit Sub
ChkErr:
    ErrSub Me.Name, "CheckYM"
End Sub

Private Function IsDataOk()
    On Error GoTo ChkErr
        IsDataOk = False
        If Not ChkDTok Then Exit Function
        If Not MustExist(gymPrtYM, 1, "印單年月") Then Exit Function
        If Not MustExist(gymClctYM1, 1, "收集年月起始日") Then Exit Function
        If Not MustExist(gymClctYM2, 1, "收集年月截止日") Then Exit Function
        If intPara24 = 0 Then If Not MustExist(gilServiceType, 2, "服務類別") Then Exit Function
        If Not MustExist(gilCompCode, 2, "公司別") Then Exit Function
        If Not MustExist(txtOrder, "排序") Then Exit Function
        If Len(gimCitemCode.GetQueryCode & "") = 0 Then
            MsgBox "收費項目需有值！", vbInformation, "訊息"
            Exit Function
        End If
        
        IsDataOk = True
    Exit Function
ChkErr:
    ErrSub Me.Name, "IsDataOk"
End Function

' 3.  確定鍵或F2按下:
Private Sub cmdOk_Click()
    On Error GoTo ChkErr
    Dim strmsg As String
    Dim strStartSNo As String
    Dim strEndSNo As String
    Dim strRcdCnt As String
    Dim rsTmp As New ADODB.Recordset
        '#2923 將TxtOrder的值寫入到文字檔,以便給予下次登入的值
        If Not TxttoRcd Then Exit Sub
        Call CheckYM
        '‧  必要欄位檢查, 若不過不可繼續。 除了複選元件外, 所有欄位皆為必要
        If Not IsDataOk Then Exit Sub
'        '    取該年月之印單序號存放於<目前編號>中. 若不存在, 則新增一筆資料,
'        '    其中: 出單年月=西曆之<出單年月>, 目前已用序號=0
        Screen.MousePointer = vbHourglass
        gcnGi.BeginTrans
        '修改資料不同收費人員,地址...
'        If SF_CHANPRTSNO(gymPrtYM.GetValue, gymClctYM1.GetValue, gymClctYM2.GetValue, _
'            cboNoe1.ListIndex, txtOrder.Text, chkClctType1.Value, chkClctType2.Value, cboRange.ListIndex, _
'            gilCompCode.GetCodeNo, gilServiceType.GetCodeNo, gimCitemCode.GetQryStr, gimClassCode.GetQryStr, gmsClctEn.GetQryStr, _
'            gmsServCode.GetQryStr, gmsClctAreaCode.GetQryStr, gmsMduId.GetQryStr, gmsStrtCode.GetQryStr, gimCMCode.GetQryStr, _
'            gimCreateEn.GetQryStr, gdaCreateTime1.GetValue, gdaCreateTime2.GetValue, strmsg, strStartSNo, strEndSNo, strRcdCnt) <> 0 Then
'            MsgBox strmsg, vbExclamation, "警告"
'            gcnGi.RollbackTrans
'            Exit Sub
'        Else
            '正式給號
            '#3678 產生日期要判斷到秒，所以起始與終止加串秒數 By Kin 2007/12/18
            '#2672 產生日期要判斷到秒 By Kin 2008/05/12
            Dim aCitemCode As String
            'aCitemCode = " In (" & gimCitemCode.GetQueryCode & ")"
            'change the paraterment of citemcode  to BillHeadFmt of cd068a field to avoid over length  by kin 2016/06/30
            aCitemCode = rsBillHeadFmt("BillHeadFmt") & ""
            If SF_GENPRTSNO(gymPrtYM.GetValue, gymClctYM1.GetValue, gymClctYM2.GetValue, _
                cboNoe1.ListIndex, txtOrder.Text, chkClctType1.Value, chkClctType2.Value, cboRange.ListIndex, _
                gilCompCode.GetCodeNo, gilServiceType.GetCodeNo, aCitemCode, gimClassCode.GetQryStr, gmsClctEn.GetQryStr, _
                gmsServCode.GetQryStr, gmsClctAreaCode.GetQryStr, gmsMduId.GetQryStr, gmsStrtCode.GetQryStr, gimCMCode.GetQryStr, _
                gimCreateEn.GetQryStr, gdaCreateTime1.GetValue & "", gdaCreateTime2.GetValue & "", _
                gdaShouldDay1.GetValue, gdaShouldDay2.GetValue, gimUCCode.GetQryStr, gimBillType.GetQryStr, strmsg, strStartSNo, strEndSNo, strRcdCnt) <> 0 Then
                MsgBox strmsg, vbExclamation, "警告"
                gcnGi.RollbackTrans
                Screen.MousePointer = vbDefault
                Exit Sub
            Else
                If strRcdCnt > 0 Then
'                    lblPrtSno1.Caption = gymPrtYM.GetValue & Format(strStartSNo, "000000")
'                    lblPrtSno2.Caption = gymPrtYM.GetValue & Format(strEndSNo, "000000")
                    lblPrtSno1.Caption = strStartSNo
                    lblPrtSno2.Caption = strEndSNo
                    lblSNoCnt.Caption = strRcdCnt & ""
                    Screen.MousePointer = vbDefault
                    If Not GetRS(rsTmp, "Select * From " & GetOwner & "tmp007", gcnGi) Then Exit Sub
                    If rsTmp.RecordCount > 0 Then
                        With ViewForm
                            .uIsSo3220A = True
                            .uRecordset = rsTmp
                            .Show vbModal
                        End With
                    End If
                    MsgBox "執行成功" & vbCrLf & _
                           "起始序號: " & gymPrtYM.GetValue & Format(strStartSNo, "000000") & vbCrLf & _
                           "截止序號: " & gymPrtYM.GetValue & Format(strEndSNo, "000000") & vbCrLf & vbCrLf & _
                           "共處理 " & strRcdCnt & " 筆資料。" & vbCrLf & vbCrLf & _
                           strmsg, vbInformation, "訊息"
                Else
                    MsgBox "執行成功,筆數 0筆"
                End If
            End If
'        End If
        gcnGi.CommitTrans
        Screen.MousePointer = vbDefault
    Exit Sub
ChkErr:
    ErrSub Me.Name, "cmdOk_Click"
End Sub

Public Function SF_GENPRTSNO(ByVal p_YM As String, ByVal p_YM1 As String, _
ByVal p_YM2 As String, ByVal p_NOE As Long, ByVal p_ORDER As String, _
ByVal p_GENOVERDUE As Long, ByVal p_GENPRCUST As Long, ByVal p_RANGE As Long, _
ByVal p_CompCode As String, ByVal p_SERVICETYPE As String, ByVal p_CitemSQL As String, _
ByVal p_ClassSQL As String, ByVal p_ClctEnSQL As String, ByVal p_ServSQL As String, _
ByVal p_ClctAreaSQL As String, ByVal p_MduIdSQL As String, ByVal p_StrtSQL As String, _
ByVal p_CMSQL As String, ByVal P_CreateEnSQL As String, _
ByVal p_CreateTime1 As String, ByVal p_CreateTime2 As String, _
ByVal p_ShouldDate1 As String, ByVal p_ShouldDate2 As String, ByVal P_UCCodeSQL As String, _
ByVal p_BillType As String, ByRef p_RetMsg As String, ByRef P_STARTSNO As String, ByRef p_ENDSNO As String, _
ByRef p_RCDCNT As String) As Long
On Error GoTo ChkErr
        Dim ComSF_GENPRTSNO As New ADODB.Command
'        If cnConn Is Nothing Then
'                MsgBox vmsgSendConn, vbCritical, vmsgPrompt
'                Exit Function
'        End If

        With ComSF_GENPRTSNO
                .Parameters.Append .CreateParameter("RETURN_VALUE", adVarNumeric, adParamReturnValue)
                .Parameters.Append .CreateParameter("P_YM", adVarChar, adParamInput, 2000, p_YM)
                .Parameters.Append .CreateParameter("P_YM1", adVarChar, adParamInput, 2000, p_YM1)
                .Parameters.Append .CreateParameter("P_YM2", adVarChar, adParamInput, 2000, p_YM2)
                .Parameters.Append .CreateParameter("P_NOE", adVarNumeric, adParamInput, , p_NOE)
                .Parameters.Append .CreateParameter("P_ORDER", adVarChar, adParamInput, 2000, p_ORDER)
                .Parameters.Append .CreateParameter("P_GENOVERDUE", adVarNumeric, adParamInput, , p_GENOVERDUE)
                .Parameters.Append .CreateParameter("P_GENPRCUST", adVarNumeric, adParamInput, , p_GENPRCUST)
                .Parameters.Append .CreateParameter("P_RANGE", adVarNumeric, adParamInput, , p_RANGE)
                .Parameters.Append .CreateParameter("P_COMPCODE", adVarChar, adParamInput, 2000, p_CompCode)
                .Parameters.Append .CreateParameter("P_SERVICETYPE", adVarChar, adParamInput, 2000, p_SERVICETYPE)
                .Parameters.Append .CreateParameter("P_CITEMSQL", adVarChar, adParamInput, 5000, p_CitemSQL)
                .Parameters.Append .CreateParameter("P_CLASSSQL", adVarChar, adParamInput, 2000, p_ClassSQL)
                .Parameters.Append .CreateParameter("P_CLCTENSQL", adVarChar, adParamInput, 2000, p_ClctEnSQL)
                .Parameters.Append .CreateParameter("P_SERVSQL", adVarChar, adParamInput, 2000, p_ServSQL)
                .Parameters.Append .CreateParameter("P_CLCTAREASQL", adVarChar, adParamInput, 2000, p_ClctAreaSQL)
                .Parameters.Append .CreateParameter("P_MDUIDSQL", adVarChar, adParamInput, 2000, p_MduIdSQL)
                .Parameters.Append .CreateParameter("P_STRTSQL", adVarChar, adParamInput, 2000, p_StrtSQL)
                .Parameters.Append .CreateParameter("P_CMSQL", adVarChar, adParamInput, 2000, p_CMSQL)
                .Parameters.Append .CreateParameter("p_CreateEnSQL", adVarChar, adParamInput, 2000, P_CreateEnSQL)
                .Parameters.Append .CreateParameter("p_CreateTime1", adVarChar, adParamInput, 2000, p_CreateTime1)
                .Parameters.Append .CreateParameter("p_CreateTime2", adVarChar, adParamInput, 2000, p_CreateTime2)
                .Parameters.Append .CreateParameter("p_ShouldDate1", adVarChar, adParamInput, 2000, p_ShouldDate1)
                .Parameters.Append .CreateParameter("p_ShouldDate2", adVarChar, adParamInput, 2000, p_ShouldDate2)
                .Parameters.Append .CreateParameter("P_UCCodeSQL", adVarChar, adParamInput, 2000, P_UCCodeSQL)
                .Parameters.Append .CreateParameter("P_BillType", adVarChar, adParamInput, 2000, p_BillType)
                .Parameters.Append .CreateParameter("P_RETMSG", adVarChar, adParamOutput, 2000, p_RetMsg)
                .Parameters.Append .CreateParameter("P_STARTSNO", adVarChar, adParamOutput, 2000, P_STARTSNO)
                .Parameters.Append .CreateParameter("P_ENDSNO", adVarChar, adParamOutput, 2000, p_ENDSNO)
                .Parameters.Append .CreateParameter("P_RCDCNT", adVarChar, adParamOutput, 2000, p_RCDCNT)
                Set .ActiveConnection = gcnGi
                .CommandText = "SF_GENPRTSNO"
                .CommandType = adCmdStoredProc
                .Execute
                p_RetMsg = .Parameters("P_RETMSG").Value & ""
                P_STARTSNO = .Parameters("P_STARTSNO").Value & ""
                p_ENDSNO = .Parameters("P_ENDSNO").Value & ""
                p_RCDCNT = .Parameters("P_RCDCNT").Value & ""
                SF_GENPRTSNO = Val(.Parameters("Return_value").Value & "")
        End With
        Exit Function
ChkErr:
        SF_GENPRTSNO = -99: p_RetMsg = "程式產生:其他錯誤"
        ErrSub Me.Name, "SF_GENPRTSNO"
End Function


Public Function SF_CHANPRTSNO(ByVal p_YM As String, ByVal p_YM1 As String, _
ByVal p_YM2 As String, ByVal p_NOE As Long, ByVal p_ORDER As String, _
ByVal p_GENOVERDUE As Long, ByVal p_GENPRCUST As Long, ByVal p_RANGE As Long, _
ByVal p_CompCode As String, ByVal p_SERVICETYPE As String, ByVal p_CitemSQL As String, _
ByVal p_ClassSQL As String, ByVal p_ClctEnSQL As String, ByVal p_ServSQL As String, _
ByVal p_ClctAreaSQL As String, ByVal p_MduIdSQL As String, ByVal p_StrtSQL As String, _
ByVal p_CMSQL As String, ByVal P_CreateEnSQL As String, _
ByVal p_CreateTime1 As String, ByVal p_CreateTime2 As String, _
ByRef p_RetMsg As String, ByRef P_STARTSNO As String, ByRef p_ENDSNO As String, _
ByRef p_RCDCNT As String) As Long
On Error GoTo ChkErr
        Dim ComSF_CHANPRTSNO As New ADODB.Command
'        If cnConn Is Nothing Then
'                MsgBox vmsgSendConn, vbCritical, vmsgPrompt
'                Exit Function
'        End If

        With ComSF_CHANPRTSNO
                .Parameters.Append .CreateParameter("RETURN_VALUE", adVarNumeric, adParamReturnValue)
                .Parameters.Append .CreateParameter("P_YM", adVarChar, adParamInput, 2000, p_YM)
                .Parameters.Append .CreateParameter("P_YM1", adVarChar, adParamInput, 2000, p_YM1)
                .Parameters.Append .CreateParameter("P_YM2", adVarChar, adParamInput, 2000, p_YM2)
                .Parameters.Append .CreateParameter("P_NOE", adVarNumeric, adParamInput, , p_NOE)
                .Parameters.Append .CreateParameter("P_ORDER", adVarChar, adParamInput, 2000, p_ORDER)
                .Parameters.Append .CreateParameter("P_GENOVERDUE", adVarNumeric, adParamInput, , p_GENOVERDUE)
                .Parameters.Append .CreateParameter("P_GENPRCUST", adVarNumeric, adParamInput, , p_GENPRCUST)
                .Parameters.Append .CreateParameter("P_RANGE", adVarNumeric, adParamInput, , p_RANGE)
                .Parameters.Append .CreateParameter("P_COMPCODE", adVarChar, adParamInput, 2000, p_CompCode)
                .Parameters.Append .CreateParameter("P_SERVICETYPE", adVarChar, adParamInput, 2000, p_SERVICETYPE)
                .Parameters.Append .CreateParameter("P_CITEMSQL", adVarChar, adParamInput, 2000, p_CitemSQL)
                .Parameters.Append .CreateParameter("P_CLASSSQL", adVarChar, adParamInput, 2000, p_ClassSQL)
                .Parameters.Append .CreateParameter("P_CLCTENSQL", adVarChar, adParamInput, 2000, p_ClctEnSQL)
                .Parameters.Append .CreateParameter("P_SERVSQL", adVarChar, adParamInput, 2000, p_ServSQL)
                .Parameters.Append .CreateParameter("P_CLCTAREASQL", adVarChar, adParamInput, 2000, p_ClctAreaSQL)
                .Parameters.Append .CreateParameter("P_MDUIDSQL", adVarChar, adParamInput, 2000, p_MduIdSQL)
                .Parameters.Append .CreateParameter("P_STRTSQL", adVarChar, adParamInput, 2000, p_StrtSQL)
                .Parameters.Append .CreateParameter("P_CMSQL", adVarChar, adParamInput, 2000, p_CMSQL)
                .Parameters.Append .CreateParameter("p_CreateEnSQL", adVarChar, adParamInput, 2000, P_CreateEnSQL)
                .Parameters.Append .CreateParameter("p_CreateTime1", adVarChar, adParamInput, 2000, p_CreateTime1)
                .Parameters.Append .CreateParameter("p_CreateTime2", adVarChar, adParamInput, 2000, p_CreateTime2)
                .Parameters.Append .CreateParameter("P_RETMSG", adVarChar, adParamOutput, 2000, p_RetMsg)
                Set .ActiveConnection = gcnGi
                .CommandText = "SF_CHANPRTSNO"
                .CommandType = adCmdStoredProc
                .Execute
                p_RetMsg = .Parameters("P_RETMSG").Value & ""
                SF_CHANPRTSNO = Val(.Parameters("Return_value").Value & "")
        End With
        Exit Function
ChkErr:
        SF_CHANPRTSNO = -99: p_RetMsg = "程式產生:其他錯誤"
        ErrSub Me.Name, "SF_CHANPRTSNO"
End Function
'#2923 將TxtOrder記錄於文字檔 By Kin 2007/04/02
Function TxttoRcd() As Boolean
On Error GoTo ChkErr:
    Dim fsoTxt As TextStream
    TxttoRcd = False
    If strOrderPath = "" Then
        MsgBox "路徑找不到", vbInformation, "訊息"
        Exit Function
    Else
        If Right(strOrderPath, 1) = "\" Then
            OpenFile fsoTxt, strOrderPath & strOrderFileName
        Else
            OpenFile fsoTxt, strOrderPath & "\" & strOrderFileName
        End If
    End If
    If txtOrder <> "" Then
        fsoTxt.WriteLine txtOrder
    End If
    fsoTxt.Close
    Set fsoTxt = Nothing
    TxttoRcd = True
    Exit Function
ChkErr:
    ErrSub Me.Name, "TxttoRcd"
End Function
'#2923 將記錄顯示於TxtOrder by Kin 2007/04/02
Function RcdToTxt() As String
On Error GoTo ChkErr:
    Dim strSO3220RC As String
    Dim fsoFile As New Scripting.FileSystemObject
    If strOrderPath = "" Then
        RcdToTxt = "1432"
        Exit Function
    Else
        If Right(strOrderPath, 1) = "\" Then
            If fsoFile.FileExists(strOrderPath & strOrderFileName) Then
                strSO3220RC = ReadFromFile(strOrderPath & strOrderFileName)
            Else
                RcdToTxt = "1432"
                Exit Function
            End If
        Else
            If fsoFile.FileExists(strOrderPath & "\" & strOrderFileName) Then
                strSO3220RC = ReadFromFile(strOrderPath & "\" & strOrderFileName)
            Else
                RcdToTxt = "1432"
                Exit Function
            End If
        End If
    End If
    If strSO3220RC = "" Then
        RcdToTxt = "1432"
    Else
        RcdToTxt = Trim(strSO3220RC)
    End If
    Set fsoFile = Nothing
    Exit Function
ChkErr:
    ErrSub Me.Name, "RcdToTxt"
End Function


