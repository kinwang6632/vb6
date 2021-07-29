VERSION 5.00
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{A0C51418-555F-11D3-BC10-0050BA01A3A9}#9.8#0"; "GiMulti.ocx"
Object = "{6DDA0569-6454-11D3-9612-0048544B09DF}#4.1#0"; "GiTime.ocx"
Object = "{074771B7-6ADE-11D5-B178-00485456B028}#2.0#0"; "csMulti.ocx"
Begin VB.Form frmSO3430A 
   BorderStyle     =   1  '單線固定
   Caption         =   "整批入短收作業 [SO3430A]"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9270
   Icon            =   "SO3430A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   9270
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton cmdOK 
      Caption         =   "F2.確定"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2400
      TabIndex        =   16
      Top             =   4740
      Width           =   1410
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&X)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   5400
      TabIndex        =   17
      Top             =   4740
      Width           =   1410
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3795
      Left            =   120
      TabIndex        =   20
      Top             =   870
      Width           =   9015
      Begin Gi_Time.GiTime gdtCreateTime1 
         Height          =   315
         Left            =   4680
         TabIndex        =   5
         Top             =   240
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   556
         ForeColor       =   16711680
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
         Height          =   315
         Left            =   1050
         TabIndex        =   3
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         ForeColor       =   16711680
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
         Left            =   2430
         TabIndex        =   4
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         ForeColor       =   16711680
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
      Begin Gi_Date.GiDate gdaRealDate1 
         Height          =   315
         Left            =   1050
         TabIndex        =   7
         Top             =   600
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         ForeColor       =   16711680
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
      Begin Gi_Date.GiDate gdaRealDate2 
         Height          =   315
         Left            =   2430
         TabIndex        =   8
         Top             =   600
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         ForeColor       =   16711680
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
         Height          =   345
         Left            =   120
         TabIndex        =   9
         Top             =   1050
         Width           =   8745
         _ExtentX        =   15425
         _ExtentY        =   609
         ButtonCaption   =   "收費項目"
      End
      Begin Gi_Multi.GiMulti gimBillType 
         Height          =   345
         Left            =   120
         TabIndex        =   10
         Top             =   1410
         Width           =   8745
         _ExtentX        =   15425
         _ExtentY        =   609
         ButtonCaption   =   "單據類別"
         DataType        =   2
         FldCaption1     =   "單據類別代碼"
         FldCaption2     =   "單據類別名稱"
         DIY             =   -1  'True
         Exception       =   -1  'True
      End
      Begin Gi_Multi.GiMulti gimUCCode 
         Height          =   345
         Left            =   120
         TabIndex        =   11
         Top             =   1770
         Width           =   8745
         _ExtentX        =   15425
         _ExtentY        =   609
         ButtonCaption   =   "未收原因"
      End
      Begin Gi_Multi.GiMulti gimCMCode 
         Height          =   345
         Left            =   120
         TabIndex        =   12
         Top             =   2130
         Width           =   8745
         _ExtentX        =   15425
         _ExtentY        =   609
         ButtonCaption   =   "收費方式"
      End
      Begin CS_Multi.CSmulti gimClassCode 
         Height          =   345
         Left            =   120
         TabIndex        =   13
         Top             =   2490
         Width           =   8745
         _ExtentX        =   15425
         _ExtentY        =   609
         ButtonCaption   =   "客戶類別"
      End
      Begin CS_Multi.CSmulti gimMduId 
         Height          =   345
         Left            =   120
         TabIndex        =   14
         Top             =   2850
         Width           =   8745
         _ExtentX        =   15425
         _ExtentY        =   609
         ButtonCaption   =   "大樓名稱"
      End
      Begin Gi_Multi.GiMulti gimCustStatus 
         Height          =   345
         Left            =   120
         TabIndex        =   15
         Top             =   3210
         Width           =   8745
         _ExtentX        =   15425
         _ExtentY        =   609
         ButtonCaption   =   "客  戶  狀  態"
      End
      Begin Gi_Time.GiTime gdtCreateTime2 
         Height          =   315
         Left            =   6600
         TabIndex        =   6
         Top             =   240
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   556
         ForeColor       =   16711680
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "至"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   6390
         TabIndex        =   28
         Top             =   300
         Width           =   195
      End
      Begin VB.Label lblCreateTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "產單日期"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   3780
         TabIndex        =   27
         Top             =   300
         Width           =   720
      End
      Begin VB.Label lblD3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "至"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   2190
         TabIndex        =   24
         Top             =   330
         Width           =   195
      End
      Begin VB.Label lblShouldDay 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "應收日期"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   180
         TabIndex        =   23
         Top             =   300
         Width           =   780
      End
      Begin VB.Label lblD1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "至"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   2190
         TabIndex        =   22
         Top             =   690
         Width           =   195
      End
      Begin VB.Label lblRealDay 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "實收日期"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   180
         TabIndex        =   21
         Top             =   690
         Width           =   780
      End
   End
   Begin prjGiList.GiList gilCompCode 
      Height          =   315
      Left            =   1200
      TabIndex        =   0
      Top             =   120
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
   Begin Gi_Date.GiDate gdaRealDate 
      Height          =   315
      Left            =   1200
      TabIndex        =   1
      Top             =   480
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   556
      ForeColor       =   255
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
   Begin prjGiList.GiList gilSTCode 
      Height          =   315
      Left            =   3840
      TabIndex        =   2
      Top             =   480
      Width           =   2325
      _ExtentX        =   4101
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
      FldWidth2       =   1400
      F2Corresponding =   -1  'True
   End
   Begin VB.Label Label3 
      Caption         =   "本功能僅提供入不更新週期性收費項目之短收原因"
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
      Height          =   375
      Left            =   6360
      TabIndex        =   26
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label lblSTCode 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "短收原因"
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
      Left            =   3000
      TabIndex        =   25
      Top             =   600
      Width           =   780
   End
   Begin VB.Label lblRealDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "入帳日期"
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
      Left            =   240
      TabIndex        =   19
      Top             =   600
      Width           =   780
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
      Height          =   195
      Left            =   240
      TabIndex        =   18
      Top             =   180
      Width           =   585
   End
End
Attribute VB_Name = "frmSO3430A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private blnDisCount As Boolean
Private Sub cmdOK_Click()
    On Error GoTo chkErr

    If Not IsDateOk Then Exit Sub
    Call subChoose
    If Not UpdSO033 Then Exit Sub

    Exit Sub
chkErr:
    ErrSub Me.Name, "cmdOK_Click"
End Sub

Private Function IsDateOk() As Boolean
  On Error Resume Next
        IsDateOk = False
        If Not ChkDTok Then Exit Function
        If Not MustExist(gilCompCode, 2, "公司別") Then Exit Function
        If Not MustExist(gilSTCode, 2, "短收原因") Then Exit Function
        If Not MustExist(gdaRealDate, 1, "入帳日期") Then Exit Function
        
        IsDateOk = True
End Function
Private Sub Form_Activate()
    On Error Resume Next
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 On Error Resume Next
        If Not ChkGiList(KeyCode) Then Exit Sub
        If KeyCode = vbKeyEscape Then Unload Me
        If KeyCode = vbKeyF2 Then
            If Not ChkCloseDate(gdaRealDate.GetValue(True), gilCompCode.GetCodeNo, "") Then
                gdaRealDate.SetFocus
                Exit Sub
            End If
            If cmdOK.Enabled = True Then cmdOK.Value = True
        End If
End Sub

Private Sub Form_Load()
  On Error Resume Next
    
    gdaRealDate.SetValue Date
    Call subGil
    Call subGim
End Sub
Private Sub subGil()
    On Error GoTo chkErr
        '公司別
        SetgiList gilCompCode, "CodeNo", "Description", "CD039", , , , , 1, 12, GetCompCodeFilter(gilCompCode)
        gilCompCode.SetCodeNo garyGi(9)
        gilCompCode.Query_Description
        '結案短收原因：
        SetgiList gilSTCode, "CodeNo", "Description", "CD016", , , , , 3, 12, "Where RefNo = 1", True
    Exit Sub
chkErr:
    ErrSub Me.Name, "subGil"
End Sub
Private Sub subGim()
    On Error GoTo chkErr
        '大樓
        SetgiMulti gimMduId, "MduId", "Name", "SO017", "編號", "名稱"
        '#XXXX 判斷是否有第二層優惠 By Kin 2009/04/01
        blnDisCount = Val(GetSystemParaItem("SecondDiscount", gCompCode, "", "SO041") & "") = 1
        If blnDisCount Then
            SetgiMulti gimCitemCode, "CodeNo", "Description", "CD019", "代碼", "名稱", "Where Nvl(RefNo,0) Not In (19,20) And Nvl(SecendDiscount,0) = 0  ", True
        Else
            SetgiMulti gimCitemCode, "CodeNo", "Description", "CD019", "代碼", "名稱", , True
        End If
        SetgiMultiAddItem gimBillType, "B,I,P,M,T", "收費單,裝機單,停拆移機單,維修單,臨時收費單"
        SetgiMulti gimUCCode, "CodeNo", "Description", "CD013", "代碼", "名稱", , True
        SetgiMulti gimCMCode, "CodeNo", "Description", "CD031", "代碼", "名稱", , True
        SetgiMulti gimClassCode, "CodeNo", "Description", "CD004", "代碼", "名稱", , True
        '#3323 增加客戶狀態讓客戶選取預設狀態為正常 By Kin 2007/07/18
        SetgiMulti gimCustStatus, "CodeNo", "Description", "CD035", "客戶狀態代碼", "客戶狀態名稱", , False
        gimCustStatus.SetQueryCode 1
    Exit Sub
chkErr:
    ErrSub Me.Name, "subGim"
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub
            
Private Sub subChoose()
    On Error GoTo chkErr
        Dim strCustStatusSQl As String
        strChoose = ""
    'GiDate
        If gdaShouldDate1.GetValue <> "" Then Call subAnd("ShouldDate >=  To_date('" & gdaShouldDate1.GetValue(True) & "','yyyy/mm/dd')")
        If gdaShouldDate2.GetValue <> "" Then Call subAnd("ShouldDate <= To_date('" & gdaShouldDate2.GetValue(True) & "','yyyy/mm/dd')")
        If gdaRealDate1.GetValue <> "" Then Call subAnd("RealDate >= To_date('" & gdaRealDate1.GetValue(True) & "','yyyy/mm/dd')")
        If gdaRealDate2.GetValue <> "" Then Call subAnd("RealDate <=  To_date('" & gdaRealDate2.GetValue(True) & "','yyyy/mm/dd')")
    'GiMulti
        If gimCitemCode.GetQryStr <> "" Then Call subAnd("CitemCode " & gimCitemCode.GetQryStr)
        If gimBillType.GetQryStr <> "" Then Call subAnd("Substr(BillNo,7,1)  " & gimBillType.GetQryStr)
        If gimUCCode.GetQryStr <> "" Then Call subAnd("UCCode " & gimUCCode.GetQryStr)
        If gimCMCode.GetQryStr <> "" Then Call subAnd("CMCode " & gimCMCode.GetQryStr)
        If gimClassCode.GetQryStr <> "" Then Call subAnd("ClassCode " & gimClassCode.GetQryStr)
        If gimMduId.GetQryStr <> "" Then Call subAnd("MduId " & gimMduId.GetQryStr)
    '#3323 增加公司別條件 By Kin 2007/07/18
        If gilCompCode.GetCodeNo <> "" Then Call subAnd("CompCode=" & gilCompCode.GetCodeNo)
    '#3323 增加產生日期的條件 By Kin 2007/07/18
        If gdtCreateTime1.GetValue <> "" Then subAnd ("CreateTime>=" & GetNullString(gdtCreateTime1.GetValue(True), giDateV, giOracle, True))
        If gdtCreateTime2.GetValue <> "" Then subAnd ("CreateTime<=" & GetNullString(gdtCreateTime2.GetValue(True), giDateV, giOracle, True))
    '#3323 增加判斷客戶狀態，要多判斷SO002 By Kin 2007/07/18
        If gimCustStatus.GetQueryCode <> "" Then
            strCustStatusSQl = "Exists (Select * From " & GetOwner & "SO002 Where SO002.Custid=SO033.Custid " & _
                               " And SO002.CompCode=SO033.CompCode And SO002.ServiceType=SO033.ServiceType " & _
                               " And SO002.CustStatusCode IN (" & gimCustStatus.GetQueryCode & "))"
            strChoose = IIf(strChoose & "" <> "", strChoose & " And " & strCustStatusSQl, strCustStatusSQl)
        End If
    Exit Sub
chkErr:
    ErrSub Me.Name, "subChoose"
End Sub

Private Function UpdSO033() As Boolean
    On Error GoTo chkErr
    Dim strSQL As String
    Dim intCount As Integer
    Dim strMustSQL As String
    '#3323 增加必要條件，以免Update錯資料 By Kin 2007/07/18
    '#XXXX 判斷是否有第二層優惠 By Kin 2009/04/01
    strMustSQL = "UcCode is Not Null And CancelFlag=0 And " & _
               " Exists(Select * From CD013 Where RefNo<>3 And SO033.UcCode=CD013.CodeNo And CD013.StopFlag=0)" & _
               " And Exists(Select * From CD019 Where CD019.StopFlag<>1 And Nvl(RefNo,0) Not IN(19,20)" & _
               " And CD019.CodeNO=SO033.CitemCode And Nvl(SecendDiscount,0) = 0) "
    '#3323 Where條件要注意是否有值，無值時，前面不能加"And" By Kin 2007/07/18
    strSQL = "Update " & GetOwner & "so033 SET " & _
             "RealDate=To_Date('" & gdaRealDate.GetValue(True) & "','YYYY/MM/DD') ," & _
             "StCode='" & gilSTCode.GetCodeNo & "'," & _
             "StName='" & gilSTCode.GetDescription & "'," & _
             "Uccode='',Ucname=''," & _
             "UpdTime ='" & GetDTString(RightNow) & "'," & _
             "Upden='" & garyGi(1) & "' " & _
             " Where " & strChoose & IIf(strChoose & "" <> "", " And " & strMustSQL, strMustSQL)
    gcnGi.BeginTrans
    Screen.MousePointer = vbHourglass
    gcnGi.Execute strSQL, intCount
    gcnGi.CommitTrans
    If intCount > 0 Then
        MsgBox "執行完畢!共更新 " & intCount & " 筆資料", vbInformation, "訊息"
    Else
        MsgBox "執行完畢!無資料! ", vbInformation, "訊息"
    End If
    
    Screen.MousePointer = vbDefault
    Exit Function
chkErr:
    gcnGi.RollbackTrans
    ErrSub Me.Name, "UpdSO033"
End Function

Private Sub gdaRealDate_Validate(Cancel As Boolean)
 '#3122 判斷入帳日期是否小於日結日期
    On Error Resume Next
        Cancel = Not ChkCloseDate(gdaRealDate.GetValue(True), gilCompCode.GetCodeNo, "")
End Sub

Private Sub gdaRealDate2_Validate(Cancel As Boolean)
 On Error Resume Next
    If gdaRealDate1.GetValue & "" <> "" Then
        If CDate(gdaRealDate1.GetValue(True)) > CDate(gdaRealDate2.GetValue(True)) Then
            MsgBox "實起始日期不得大於實收終止日", vbInformation, "警告訊息"
            gdaRealDate2.SetValue gdaRealDate1.GetValue
            Cancel = True
        End If
    End If
End Sub

Private Sub gdaShouldDate2_Validate(Cancel As Boolean)
    On Error Resume Next
    If gdaShouldDate1.GetValue & "" <> "" Then
        If CDate(gdaShouldDate1.GetValue(True)) > CDate(gdaShouldDate2.GetValue(True)) Then
            MsgBox "應收起始日期不得大於應收終止日", vbInformation, "警告訊息"
            gdaShouldDate2.SetValue gdaShouldDate1.GetValue
            Cancel = True
        End If
    End If
End Sub

Private Sub gdtCreateTime2_Validate(Cancel As Boolean)
    On Error Resume Next
    If (gdtCreateTime2.Text = "") And (gdtCreateTime1.Text <> "") Then
        Call gdtCreateTime2.SetValue(gdtCreateTime1.GetDate & "2359")
    End If
    If Not IsDate(gdtCreateTime2.GetValue(True)) Then Exit Sub
    If DateDiff("n", gdtCreateTime1.GetValue(True), gdtCreateTime2.GetValue(True)) < 0 Then
       MsgBox "產生日期截止日不可小於產生日期起始日!", vbExclamation, "警告訊息"
       gdtCreateTime2.SetValue gdtCreateTime1.GetDate & "2359"
       Cancel = True
    End If
End Sub

Private Sub gilCompCode_Change()
    On Error Resume Next
        Call subGil
        Call subGim
End Sub

