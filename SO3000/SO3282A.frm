VERSION 5.00
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{A0C51418-555F-11D3-BC10-0050BA01A3A9}#9.8#0"; "GiMulti.ocx"
Object = "{6DDA0569-6454-11D3-9612-0048544B09DF}#4.1#0"; "GiTime.ocx"
Object = "{074771B7-6ADE-11D5-B178-00485456B028}#2.0#0"; "csMulti.ocx"
Begin VB.Form frmSO3282A 
   BorderStyle     =   1  '單線固定
   Caption         =   "收費單號調整(新) [SO3282A]"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7395
   BeginProperty Font 
      Name            =   "新細明體"
      Size            =   9.75
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SO3282A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   7395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '螢幕中央
   Begin VB.Frame Frame1 
      Caption         =   "單據合併調整項目"
      ForeColor       =   &H00FF0000&
      Height          =   645
      Left            =   120
      TabIndex        =   29
      Top             =   4110
      Width           =   7155
      Begin VB.CheckBox chkUpdCreateEn 
         Caption         =   "產單人員"
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   5310
         TabIndex        =   16
         Top             =   240
         Width           =   1545
      End
      Begin VB.CheckBox chkCmCode 
         Caption         =   "收費方式"
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   3690
         TabIndex        =   15
         Top             =   240
         Width           =   1185
      End
      Begin VB.CheckBox chkCreateTime 
         Caption         =   "產生日期"
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2010
         TabIndex        =   14
         Top             =   240
         Width           =   1185
      End
      Begin VB.CheckBox chkShouldDate 
         Caption         =   "應收日期"
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   300
         TabIndex        =   13
         Top             =   240
         Width           =   1185
      End
   End
   Begin VB.CommandButton cmdPrevRpt 
      Caption         =   "上次統計結果"
      Height          =   375
      Left            =   2130
      TabIndex        =   18
      Top             =   4920
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "F2. 確定"
      Default         =   -1  'True
      Height          =   375
      Left            =   240
      TabIndex        =   17
      Top             =   4920
      Width           =   1365
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&X)"
      Height          =   375
      Left            =   5850
      TabIndex        =   19
      Top             =   4920
      Width           =   1365
   End
   Begin VB.Frame fraData 
      Height          =   4020
      Left            =   120
      TabIndex        =   20
      ToolTipText     =   "本功能只針對未收之T單做合併"
      Top             =   30
      Width           =   7155
      Begin VB.PictureBox pic1 
         Height          =   1605
         Left            =   120
         ScaleHeight     =   1545
         ScaleWidth      =   6885
         TabIndex        =   31
         Top             =   2280
         Width           =   6945
         Begin VB.VScrollBar vsl1 
            Height          =   1515
            Left            =   6600
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   0
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Frame fraMulti 
            BorderStyle     =   0  '沒有框線
            Height          =   1545
            Left            =   -30
            TabIndex        =   32
            Top             =   0
            Width           =   6885
            Begin CS_Multi.CSmulti gimUCCode 
               Height          =   345
               Left            =   30
               TabIndex        =   10
               Top             =   390
               Width           =   6880
               _ExtentX        =   12144
               _ExtentY        =   609
               ButtonCaption   =   "未 收 費 原 因"
            End
            Begin Gi_Multi.GiMulti gimBillType 
               Height          =   375
               Left            =   30
               TabIndex        =   11
               Top             =   780
               Width           =   6880
               _ExtentX        =   12144
               _ExtentY        =   661
               ButtonCaption   =   "單  據  類  別"
               DataType        =   2
               DIY             =   -1  'True
            End
            Begin Gi_Multi.GiMulti gimPayType 
               Height          =   375
               Left            =   30
               TabIndex        =   12
               Top             =   1140
               Width           =   6880
               _ExtentX        =   12144
               _ExtentY        =   661
               ButtonCaption   =   "繳  付  類  別"
               DataType        =   2
            End
            Begin CS_Multi.CSmulti gimCitemCode 
               Height          =   345
               Left            =   30
               TabIndex        =   9
               Top             =   30
               Width           =   6885
               _ExtentX        =   12144
               _ExtentY        =   609
               ButtonCaption   =   "收 費 項 目"
            End
         End
      End
      Begin VB.TextBox txtCustId 
         Height          =   315
         Left            =   2760
         MaxLength       =   200
         TabIndex        =   4
         Top             =   1080
         Width           =   4230
      End
      Begin Gi_Time.GiTime gdtCreateTime2 
         Height          =   345
         Left            =   3240
         TabIndex        =   8
         Top             =   1860
         Width           =   1845
         _ExtentX        =   3254
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
      Begin Gi_Time.GiTime gdtCreateTime1 
         Height          =   315
         Left            =   1020
         TabIndex        =   7
         Top             =   1890
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
      Begin VB.ComboBox cboBillHeadFmt 
         Height          =   315
         Left            =   1980
         Style           =   2  '單純下拉式
         TabIndex        =   2
         Top             =   720
         Visible         =   0   'False
         Width           =   2025
      End
      Begin Gi_Date.GiDate gdaProcessDate 
         Height          =   315
         Left            =   4965
         TabIndex        =   3
         Top             =   705
         Visible         =   0   'False
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
      Begin prjGiList.GiList gilServiceType 
         Height          =   315
         Left            =   1065
         TabIndex        =   1
         Top             =   665
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
         Left            =   1065
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
      Begin Gi_Date.GiDate gdaShouldDate1 
         Height          =   345
         Left            =   1020
         TabIndex        =   5
         Top             =   1470
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
      Begin Gi_Date.GiDate gdaShouldDate2 
         Height          =   345
         Left            =   2760
         TabIndex        =   6
         Top             =   1470
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
         AutoSize        =   -1  'True
         Caption         =   "客戶編號(以,相隔或以-取範圍)"
         Height          =   195
         Left            =   150
         TabIndex        =   30
         Top             =   1170
         Width           =   2565
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "至"
         Height          =   195
         Left            =   2970
         TabIndex        =   28
         Top             =   1950
         Width           =   195
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "產生時間"
         Height          =   195
         Left            =   150
         TabIndex        =   27
         Top             =   1950
         Width           =   780
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "至"
         Height          =   195
         Left            =   2370
         TabIndex        =   26
         Top             =   1530
         Width           =   195
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "應收日期"
         Height          =   195
         Left            =   150
         TabIndex        =   25
         Top             =   1560
         Width           =   780
      End
      Begin VB.Label lblBillHeadFmt 
         AutoSize        =   -1  'True
         Caption         =   "多帳戶產生依據設定"
         Height          =   195
         Left            =   150
         TabIndex        =   24
         Top             =   780
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.Label lblServiceType 
         AutoSize        =   -1  'True
         Caption         =   "服務類別"
         Height          =   195
         Left            =   180
         TabIndex        =   23
         Top             =   720
         Width           =   780
      End
      Begin VB.Label lblCompCode 
         AutoSize        =   -1  'True
         Caption         =   "公司別"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   180
         TabIndex        =   22
         Top             =   330
         Width           =   585
      End
      Begin VB.Label lblProcessDate 
         AutoSize        =   -1  'True
         Caption         =   "處理日期"
         Height          =   195
         Left            =   4080
         TabIndex        =   21
         Top             =   780
         Visible         =   0   'False
         Width           =   780
      End
   End
End
Attribute VB_Name = "frmSO3282A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private intPara24 As Integer
Private rsBillHeadFmt As New ADODB.Recordset
Private strPBillType As String   '是否啟用跨服務
Private strPCitemStr As String  '非週期性項目與P服務
Private blnIFlag As Boolean     '啟動跨服務後，收費項目裡是否有包含I服務
Private strCustIdStr As String
'Private blnShouldDateMerge As Boolean
Private Sub cboBillHeadFmt_Click()
  On Error GoTo chkErr
    Dim strCitemCode As String
    Dim rsCD019  As New ADODB.Recordset
    Dim rsCD019Mutli As New ADODB.Recordset
    Dim strCD019SQL As String
    If rsBillHeadFmt.RecordCount = 0 Then Exit Sub
    rsBillHeadFmt.AbsolutePosition = cboBillHeadFmt.ListIndex + 1
     '#7049 改用CD068A.CitemCode By Kin 2015/07/14
    strCD019SQL = "Select CodeNo From " & GetOwner & "CD019 " & _
                        " Where CodeNo In ( Select CitemCode From " & GetOwner & "CD068A " & _
                        "         Where Exists (Select * From " & GetOwner & "CD068 Where CD068.BillHeadFmt = CD068A.BillHeadFmt " & _
                                                " And CD068.BillHeadFmt = '" & rsBillHeadFmt("BillHeadFmt") & "' )) And StopFlag <>1"
    'If Not GetRS(rsCD019, "Select CodeNo From " & GetOwner & "CD019 Where CodeNo In (" & rsBillHeadFmt("CitemCodeStr") & ") And StopFlag=0") Then Exit Sub
    If Not GetRS(rsCD019, strCD019SQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
    If rsCD019.EOF Then Exit Sub
    strCitemCode = rsCD019.GetString(, , , ",")
    strCitemCode = Left(strCitemCode, Len(strCitemCode) - 1)
    'gimCitemCode.Filter = "Where PeriodFlag = 1"
    
'    If strCitemCode <> "" Then
'        gimCitemCode.Filter = "Where CodeNo In (" & strCitemCode & ") And PeriodFlag = 1"
'
'    End If
     '#7049 改變畫面的收費項目 By Kin 2015/07/15
    Dim strCD068ACitem As String
    strCD068ACitem = "Select CitemCode From " & GetOwner & "CD068A Where  CD068A.BillHeadFmt = '" & cboBillHeadFmt.Text & "'"
    If strCitemCode <> "" Then
        'gimCitemCode.Filter = "Where CodeNo In (" & strCitemCode & ") "
        gimCitemCode.Filter = "Where CodeNo In (" & strCD068ACitem & ") "
    End If
    
    If rsBillHeadFmt("MutiServiceUnion") = 1 Then '啟動
        '#7049 改用CD068A.CitemCode By Kin 2015/07/14
'        blnIFlag = GetRsValue("Select Count(*) From " & GetOwner & "CD019 " & _
'                    " Where CodeNo In (" & strCitemCode & ") And ServiceType='I'", gcnGi)
         blnIFlag = GetRsValue("Select Count(*) From " & GetOwner & "CD019 " & _
                    " Where CodeNo In (" & strCD068ACitem & ") And ServiceType='I'", gcnGi)
'        If Not GetRS(rsCD019Mutli, "Select CodeNo From " & GetOwner & "CD019 " & _
'                " Where CodeNo In (" & strCitemCode & ") And PeriodFlag=0 " & _
'                " And ServiceType='P'", , adUseClient, adOpenKeyset) Then Exit Sub
        If Not GetRS(rsCD019Mutli, "Select CodeNo From " & GetOwner & "CD019 " & _
                " Where CodeNo In (" & strCD068ACitem & ") And PeriodFlag=0 " & _
                " And ServiceType='P'", , adUseClient, adOpenKeyset) Then Exit Sub
        If rsCD019Mutli.RecordCount > 0 Then
            strPBillType = "'B'"
            strPCitemStr = rsCD019Mutli.GetString(, , , ",")
            strPCitemStr = " IN (" & Left(strPCitemStr, Len(strPCitemStr) - 1) & ")"
        Else
            strPBillType = Empty
            strPCitemStr = Empty
            blnIFlag = False
        End If
    Else
        strPBillType = Empty
        strPCitemStr = Empty
        blnIFlag = False
    End If
    gimCitemCode.SetQueryCode strCitemCode
    On Error Resume Next
    CloseRecordset rsCD019
    CloseRecordset rsCD019Mutli
    Exit Sub
chkErr:
  Call ErrSub(Me.Name, "cboBillHeadFmt_Click")
End Sub

Private Sub cmdCancel_Click()
    On Error GoTo chkErr
        Unload Me
    Exit Sub
chkErr:
    ErrSub Me.Name, "cmdCancel_Click"
End Sub

Private Sub cmdOK_Click()
  On Error GoTo chkErr
    Dim strmsg As String
    Dim rsTmp As New ADODB.Recordset
    Dim intCount As Long
    Dim intError As Integer
    Dim aryCustId() As String
    Dim blnShwCombItem As Boolean
    Dim strPayKinds As String
    Dim strCitemCode As String
    Dim strUCCode As String
    Dim strBillType As String
    blnShwCombItem = False
    
    Call CheckYM
    strCustIdStr = Empty
    strCitemCode = Empty
    strUCCode = Empty
    If Not IsDataOk Then Exit Sub
    If Trim(txtCustId) <> "" Then
        Call ParseWord(txtCustId, aryCustId)
        strCustIdStr = Join(aryCustId, ",")
    End If
    Screen.MousePointer = vbHourglass
    Me.Enabled = False
    'ReadyGoPrint
    'Call subChoose
    strPayKinds = ""
    If (gimPayType.GetQryStr <> "") And (InStr(1, gimPayType.GetQryStr, "全選") <= 0) Then
        strPayKinds = gimPayType.GetQryStr
    End If
    If gimCitemCode.GetQueryCode <> "" Then
        strCitemCode = "IN (" & gimCitemCode.GetQueryCode & ")"
    End If
    If gimUCCode.GetQueryCode <> "" Then
        strUCCode = "IN (" & gimUCCode.GetQueryCode & ")"
    End If
    If gimBillType.GetQueryCode <> "" Then
        strBillType = "IN (" & gimBillType.GetQueryCode & ")"
    End If
    
    Call Pk_CombineBill(garyGi(1), gilCompCode.GetCodeNo, strCustIdStr, _
            gdaShouldDate1.GetValue, gdaShouldDate2.GetValue, _
            gdtCreateTime1.GetValue, gdtCreateTime2.GetValue, _
             strCitemCode, strUCCode, strBillType, strPayKinds, strmsg)
             
    MsgBox strmsg, vbInformation, "警告"
    
    On Error Resume Next
    Me.Enabled = True
    Screen.MousePointer = vbDefault
    Exit Sub
chkErr:
    Me.Enabled = True
    ErrSub Me.Name, "cmdOk_Click: ErrorNo:" & intError
End Sub

Private Sub cmdPrevRpt_Click()
  On Error GoTo chkErr
    PreviousRpt GetPrinterName(5), "SO3280A.rpt", "收費單號調整 [SO3280A]"
    PreviousRpt GetPrinterName(5), "SO3280A1.rpt", "收費單號調整 [SO3280A]"
  Exit Sub
chkErr:
    Call ErrSub(Me.Name, "cmdPrevRpt_Click")
End Sub

Private Sub Form_Activate()
    On Error GoTo chkErr
        Screen.MousePointer = vbDefault
    Exit Sub
chkErr:
    ErrSub Me.Name, "Form_Activate"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo chkErr
        If Not ChkGiList(KeyCode) Then Exit Sub
        If KeyCode = vbKeyF2 Then
            If cmdOk.Enabled Then Call cmdOK_Click
            KeyCode = 0
        End If
    Exit Sub
chkErr:
    ErrSub Me.Name, "Form_KeyDown"
End Sub

Private Sub Form_Load()
  On Error GoTo ER
    subGim
    subGil
    Call CboAddItem
    DefaultValue
    Exit Sub
ER:
    If ErrHandle(, True, , "Form_Load") Then Resume 0
End Sub

Private Sub subGil()
  On Error GoTo chkErr
    SetgiList gilServiceType, "CodeNo", "Description", "CD046", , , , , 1, 12
    SetgiList gilCompCode, "CodeNo", "Description", "CD039", , , , , 3, 12, GetCompCodeFilter(gilCompCode)
    Exit Sub
chkErr:
    ErrSub Me.Name, "subGil"
End Sub

Private Sub CboAddItem()
  On Error GoTo chkErr
    If Not GetRS(rsBillHeadFmt, "Select BillHeadFmt,CitemCodeStr,MutiServiceUnion From " & GetOwner & "CD068") Then Exit Sub
    
    Do While Not rsBillHeadFmt.EOF
        cboBillHeadFmt.AddItem rsBillHeadFmt("BillHeadFmt") & ""
        rsBillHeadFmt.MoveNext
    Loop
  Exit Sub
chkErr:
  Call ErrSub(Me.Name, "CboAddItem")
End Sub

Private Sub DefaultValue()
  On Error GoTo chkErr
    Dim intPara23 As Integer
    gilCompCode.SetCodeNo garyGi(9)
    gilCompCode.Query_Description
    
    gimBillType.SetDispStr "收費單,臨時收費單"
    gimBillType.SetQueryCode "B,T"
    
    
    '處理日期
    gdaProcessDate.SetValue RightDate
            
    intPara24 = GetSystemParaItem("Para24", gilCompCode.GetCodeNo, gilServiceType.GetCodeNo, "SO043")
    intPara23 = Val(GetSystemParaItem("Para23", gilCompCode.GetCodeNo, gilServiceType.GetCodeNo, "SO043") & "")
    If intPara24 = 1 Then
        gilServiceType.Clear
        gilServiceType.Visible = False
        lblServiceType.Visible = False
        lblBillHeadFmt.Visible = True
        cboBillHeadFmt.Visible = True
        cboBillHeadFmt.ListIndex = 0
    Else
        gilServiceType.ListIndex = 1
        gilServiceType.Visible = True
        lblServiceType.Visible = True
        lblBillHeadFmt.Visible = False
        cboBillHeadFmt.Visible = False
    End If
    Exit Sub
chkErr:
    ErrSub Me.Name, "DefalutValue"
End Sub

Private Sub subGim()
  On Error GoTo chkErr
    SetgiMulti gimCitemCode, "CodeNo", "Description", "CD019", "收費項目代碼", "收費項目名稱", , True
    SetgiMulti gimUCCode, "CodeNo", "Description", "CD013", "未收費原因代碼", "未收費原因名稱", " Where Nvl(REFNO,0) NOT IN(3,7) AND NVL(PayOK,0)=0 ", True
    
    SetgiMulti gimPayType, "CodeNo", "Description", "CD112", "代碼", "名稱"
    Call SetgiMultiAddItem(gimBillType, "B,T,I,M,P", "收費單,臨時收費單,裝機單,維修單,停拆移機單", "代碼", "名稱")
    If GetPaynowFlag Then
        With gimPayType
            Select Case Val(GetRsValue("SELECT NVL(PayKindDefault,0) FROM " & GetOwner & "SO041", gcnGi) & "")
                Case 0
                    .SelectAll
                Case 1
                    .SetQueryCode "0"
                Case 2
                    .SetQueryCode "1"
            End Select
        End With
    Else
        gimPayType.Clear
        gimPayType.Enabled = False
    End If
    Exit Sub
chkErr:
    ErrSub Me.Name, "subGim"
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
    Call ReleaseCOM(Me)
End Sub
Private Sub gdaShouldDate1_GotFocus()
  On Error Resume Next
    If gdaShouldDate1.GetValue = "" Then gdaShouldDate1.SetValue (RightDate)
End Sub

Private Sub gdaShouldDate2_GotFocus()
  On Error Resume Next
    If gdaShouldDate1.GetValue <> "" Then gdaShouldDate2.SetValue gdaShouldDate1.GetValue
End Sub

Private Sub gdtCreateTime1_GotFocus()
  On Error Resume Next
    If gdtCreateTime1.GetValue = "" Then gdtCreateTime1.SetValue (RightDate)
End Sub

Private Sub gdtCreateTime2_GotFocus()
  On Error Resume Next
    If gdtCreateTime1.GetValue <> "" Then gdtCreateTime2.SetValue GetDayLastTime(gdtCreateTime1.GetValue(True))
End Sub

Private Sub gilCompCode_Change()
  On Error GoTo chkErr
    Dim rsTmp As New ADODB.Recordset
    Dim strTmpSql As String
    If Not ChgComp(gilCompCode, "SO3310", "SO3311") Then Exit Sub
    Call subGil
    Call subGim
    Call SelectServType(gilCompCode.GetCodeNo, gilServiceType)
    gilServiceType.SetCodeNo ""
    gilServiceType.Query_Description
    gilServiceType.ListIndex = 1
    strTmpSql = "Select nvl(UpdShoulddate,0) as UpdShoulddate,nvl(UpdCreatetime,0) as UpdCreatetime," & _
                "nvl(UpdCMCode,0) as UpdCMCode,nvl(UpdCreateEn,0) as UpdCreateEn from " & GetOwner & _
                "SO041 where SysID=" & gilCompCode.GetCodeNo
    If Not GetRS(rsTmp, strTmpSql) Then Exit Sub
    chkShouldDate = Val(rsTmp("UpdShoulddate") & "")
    chkCreateTime = Val(rsTmp("UpdCreatetime") & "")
    chkCmCode = Val(rsTmp("UpdCMCode") & "")
    chkUpdCreateEn = Val(rsTmp("UpdCreateEn") & "")
    On Error Resume Next
    CloseRecordset rsTmp
  Exit Sub
chkErr:
  Call ErrSub(Me.Name, "gilCompCode_Change")
End Sub

Private Sub gilServiceType_Change()
  On Error GoTo chkErr
    Call GiMultiFilter(gimCitemCode, gilServiceType.GetCodeNo)
    gimCitemCode.Filter = gimCitemCode.Filter & IIf(Trim(gimCitemCode.Filter) = "", " Where ", " And ") & " PeriodFlag = 1"
    Call GiMultiFilter(gimUCCode, gilServiceType.GetCodeNo, , , " Nvl(REFNO,0) NOT IN(3,7) AND NVL(PayOK,0)=0 And Nvl(StopFlag,0)=0 ")
  Exit Sub
chkErr:
  Call ErrSub(Me.Name, "gilServiceType_Change")
End Sub
Private Sub CheckYM()
  On Error GoTo chkErr
    If gdaShouldDate2.GetValue <> "" And gdaShouldDate1.GetValue = "" Then
        gdaShouldDate1.SetValue (gdaShouldDate2.GetValue)
    End If
    If gdaShouldDate1.GetValue <> "" And gdaShouldDate2.GetValue = "" Then
        gdaShouldDate2.SetValue (gdaShouldDate1.GetValue)
    End If
    
    If gdtCreateTime2.GetValue <> "" And gdtCreateTime1.GetValue = "" Then
        gdtCreateTime1.SetValue (gdtCreateTime2.GetValue)
    End If
    If gdtCreateTime1.GetValue <> "" And gdtCreateTime2.GetValue = "" Then
        gdtCreateTime2.SetValue (gdtCreateTime1.GetValue)
    End If
  Exit Sub
chkErr:
    ErrSub Me.Name, "CheckYM"
End Sub

Private Function IsDataOk()
  On Error GoTo chkErr
    IsDataOk = False
    If Not ChkDTok Then Exit Function
    If Not MustExist(gilCompCode, 2, "公司別") Then Exit Function
    
    If gdaShouldDate1.GetValue = "" Then MsgMustBe ("應收起始日期"): gdaShouldDate1.SetFocus: Exit Function
    If gdaShouldDate2.GetValue = "" Then MsgMustBe ("應收終止日期"): gdaShouldDate2.SetFocus: Exit Function
    
    
    If intPara24 = 0 Then If Not MustExist(gilServiceType, 2, "服務類別") Then Exit Function
    'If Not MustExist(gdaProcessDate, 1, "處理日期") Then Exit Function
    
    If gimCitemCode.GetQueryCode = "" Then MsgMustBe ("收費項目"): gimCitemCode.SetFocus: Exit Function
    
    If gdaShouldDate1.GetValue > gdaShouldDate2.GetValue Then MsgBox "應收起始日期不得大於終止日!!", vbInformation, "警告訊息": gdaShouldDate2.SetFocus: Exit Function
    If gdtCreateTime1.GetValue > gdtCreateTime2.GetValue Then MsgBox "產生起始時間不得大於終止時間!!", vbInformation, "警告訊息": gdtCreateTime2.SetFocus: Exit Function
    IsDataOk = True
    Exit Function
chkErr:
    ErrSub Me.Name, "IsDataOk"
End Function

Public Function Pk_CombineBill(ByVal p_UserId As String, ByVal p_CompCode As String, _
    ByVal p_CustStr As String, ByVal p_ShouldDate1 As String, ByVal p_ShouldDate2 As String, _
    ByVal p_CreateTime1 As String, ByVal p_CreateTime2 As String, _
    ByVal p_CitemStr As String, ByVal p_UCCode As String, _
    ByVal p_BillType As String, ByVal p_PayKind As String, ByRef p_RetMsg As String) As Long
    
  On Error GoTo chkErr
    Dim ComPk_CombineBill As New ADODB.Command
    
    With ComPk_CombineBill
        .Parameters.Append .CreateParameter("RETURN_VALUE", adVarNumeric, adParamReturnValue)
        .Parameters.Append .CreateParameter("p_UserId", adVarChar, adParamInput, 2000, p_UserId)
        .Parameters.Append .CreateParameter("p_CompCode", adVarNumeric, adParamInput, 2000, p_CompCode)
        .Parameters.Append .CreateParameter("p_CustStr", adVarChar, adParamInput, 4000, IIf(strCustIdStr = Empty, Null, strCustIdStr))
        .Parameters.Append .CreateParameter("p_ShouldDate1", adVarChar, adParamInput, 2000, IIf(p_ShouldDate1 = Empty, Null, p_ShouldDate1))
        .Parameters.Append .CreateParameter("p_ShouldDate2", adVarChar, adParamInput, 2000, IIf(p_ShouldDate2 = Empty, Null, p_ShouldDate2))
        .Parameters.Append .CreateParameter("p_CreateTime1", adVarChar, adParamInput, 2000, IIf(p_CreateTime1 = "", Null, p_CreateTime1))
        .Parameters.Append .CreateParameter("p_CreateTime2", adVarChar, adParamInput, 2000, IIf(p_CreateTime2 = "", Null, p_CreateTime2))
        .Parameters.Append .CreateParameter("p_CitemStr", adVarChar, adParamInput, 4000, p_CitemStr)
        .Parameters.Append .CreateParameter("p_UCCode", adVarChar, adParamInput, 2000, IIf(p_UCCode = "", Null, p_UCCode))
        .Parameters.Append .CreateParameter("p_BillType", adVarChar, adParamInput, 2000, IIf(p_BillType = "", Null, p_BillType))
        .Parameters.Append .CreateParameter("p_PayKind", adVarChar, adParamInput, 2000, IIf(p_PayKind = Empty, Null, p_PayKind))
        .Parameters.Append .CreateParameter("P_RETMSG", adVarChar, adParamOutput, 2000, p_RetMsg)
        
        Set .ActiveConnection = gcnGi
        .CommandType = adCmdStoredProc
        .CommandText = "Pk_CombineBill.EXECUTE" '呼叫package
        
        .Execute
        p_RetMsg = .Parameters("P_RETMSG").Value & ""
        Pk_CombineBill = Val(.Parameters("Return_value").Value & "")
    End With

  Exit Function
chkErr:
    Pk_CombineBill = -99: p_RetMsg = "程式產生:其他錯誤"
    ErrSub Me.Name, "Pk_CombineBill"
End Function

Private Sub subChoose()
  On Error Resume Next
    strChooseString = "公司別  : " & gilCompCode.GetDescription & ";" & _
                      "服務類別: " & gilServiceType.GetDescription & ";" & _
                      "收費項目: " & gimCitemCode.GetDispStr & ";" & _
                      "未收費原因: " & gimUCCode.GetDispStr & ";" & _
                      "單據類別:" & gimBillType.GetDispStr & ";" & _
                      "繳付類別:" & gimPayType.GetDispStr & ";" & _
                      "應收時間:" & gdaShouldDate1.GetValue(True) & " ~ " & gdaShouldDate2.GetValue(True) & ";" & _
                      "產生時間:" & gdtCreateTime1.GetValue(True) & " ~ " & gdtCreateTime2.GetValue(True)


End Sub
Private Sub txtCustId_KeyPress(KeyAscii As Integer)
  On Error Resume Next
    If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 44 Or KeyAscii = 8 Or KeyAscii = 45 Then
        If KeyAscii = 44 Or KeyAscii = 45 Then
            If Asc(Right(" " & txtCustId.Text, 1)) = 44 Or Asc(Right(" " & txtCustId.Text, 1)) = 45 Or txtCustId = "" Then KeyAscii = 9
        End If
        If Not ChkMaxLengthOK(txtCustId) Then
           If Not (KeyAscii = 44 Or KeyAscii = 8 Or KeyAscii = 45) Then KeyAscii = 9
        End If
    Else
        KeyAscii = 9
    End If
End Sub



Private Sub vsl1_Change()
  On Error Resume Next
    fraMulti.Top = 20 - fraMulti.Height * (vsl1.Value / 20) / 6
End Sub
