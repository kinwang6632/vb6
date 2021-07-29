VERSION 5.00
Object = "{A0C51418-555F-11D3-BC10-0050BA01A3A9}#9.8#0"; "GiMulti.ocx"
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{074771B7-6ADE-11D5-B178-00485456B028}#1.17#0"; "csMulti.ocx"
Begin VB.Form frmSo5290A 
   BorderStyle     =   1  '單線固定
   Caption         =   "每月每日收費單據數總表[SO5290A]"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7695
   Icon            =   "SO5290A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '螢幕中央
   Begin VB.Frame Frame4 
      Caption         =   "網路編號"
      Height          =   735
      Left            =   30
      TabIndex        =   30
      Top             =   4860
      Width           =   7665
      Begin VB.OptionButton optCircuitNo 
         Caption         =   "只印空白網路編號"
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   4830
         TabIndex        =   15
         Top             =   240
         Width           =   1845
      End
      Begin VB.OptionButton optAll 
         Caption         =   "全部"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   6660
         TabIndex        =   16
         Top             =   270
         Value           =   -1  'True
         Width           =   675
      End
      Begin VB.TextBox mskCircuitNo 
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   90
         TabIndex        =   14
         Top             =   240
         Width           =   4605
      End
   End
   Begin CS_Multi.CSmulti gimStrtCode 
      Height          =   405
      Left            =   0
      TabIndex        =   12
      Top             =   4080
      Width           =   7725
      _ExtentX        =   13626
      _ExtentY        =   714
      ButtonCaption   =   "街  道  範  圍"
   End
   Begin VB.Frame fraPage 
      Caption         =   "小計依據"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   735
      Left            =   30
      TabIndex        =   27
      Top             =   5640
      Width           =   7635
      Begin VB.OptionButton optClctAreaCode 
         Caption         =   "收費區"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   3104
         TabIndex        =   19
         Top             =   330
         Width           =   915
      End
      Begin VB.OptionButton optStrtCode 
         Caption         =   "街道編號"
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   4401
         TabIndex        =   20
         Top             =   330
         Width           =   1245
      End
      Begin VB.OptionButton optClctEn 
         Caption         =   "收費人員"
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   6030
         TabIndex        =   21
         Top             =   330
         Width           =   1245
      End
      Begin VB.OptionButton optAreaCode 
         Caption         =   "行政區"
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   1807
         TabIndex        =   18
         Top             =   330
         Width           =   915
      End
      Begin VB.OptionButton optServCode 
         Caption         =   "服務區"
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   510
         TabIndex        =   17
         Top             =   330
         Value           =   -1  'True
         Width           =   915
      End
   End
   Begin VB.CommandButton cmdPrevRpt 
      Caption         =   "上次統計結果"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   1710
      TabIndex        =   23
      Top             =   6690
      Width           =   1395
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   6180
      TabIndex        =   24
      Top             =   6690
      Width           =   1245
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "F5:列印"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   30
      TabIndex        =   22
      Top             =   6690
      Width           =   1245
   End
   Begin Gi_Date.GiDate gdaShouldDate2 
      Height          =   345
      Left            =   2340
      TabIndex        =   1
      Top             =   120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
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
   Begin Gi_Date.GiDate gdaShouldDate1 
      Height          =   345
      Left            =   990
      TabIndex        =   0
      Top             =   120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
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
   Begin Gi_Multi.GiMulti gimBillType 
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   1350
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   661
      ButtonCaption   =   "單  據  類  別"
      DataType        =   2
   End
   Begin Gi_Multi.GiMulti gimAreaCode 
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   3690
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   661
      ButtonCaption   =   "行     政     區"
   End
   Begin Gi_Multi.GiMulti gimServCode 
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   3300
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   661
      ButtonCaption   =   "服     務     區"
   End
   Begin Gi_Multi.GiMulti gimCMCode 
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   2910
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   661
      ButtonCaption   =   "收  費  方  式"
   End
   Begin Gi_Multi.GiMulti gimStatusCode 
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   2130
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   661
      ButtonCaption   =   "客  戶  狀  態"
   End
   Begin Gi_Multi.GiMulti gimUCCode 
      Height          =   375
      Left            =   0
      TabIndex        =   13
      Top             =   4470
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   661
      ButtonCaption   =   "未 繳 費 原 因"
   End
   Begin prjGiList.GiList gilServiceType 
      Height          =   315
      Left            =   4530
      TabIndex        =   3
      Top             =   540
      Width           =   2940
      _ExtentX        =   5186
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      F5Corresponding =   -1  'True
   End
   Begin prjGiList.GiList gilCompCode 
      Height          =   315
      Left            =   4530
      TabIndex        =   2
      Top             =   150
      Width           =   2940
      _ExtentX        =   5186
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      F5Corresponding =   -1  'True
   End
   Begin CS_Multi.CSmulti gimClctEn 
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   2520
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   661
      ButtonCaption   =   "收  費  人   員"
   End
   Begin CS_Multi.CSmulti gimCitemCode 
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   990
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   661
      ButtonCaption   =   "收  費  項  目"
   End
   Begin CS_Multi.CSmulti gimClassCode 
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   1740
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   661
      ButtonCaption   =   "客  戶  類  別"
   End
   Begin VB.Label lblCompCode 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BackStyle       =   0  '透明
      Caption         =   "公  司  別"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   3690
      TabIndex        =   29
      Top             =   210
      Width           =   765
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BackStyle       =   0  '透明
      Caption         =   "服務類別"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   3720
      TabIndex        =   28
      Top             =   600
      Width           =   780
   End
   Begin VB.Label lblShouldDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "應收日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   90
      TabIndex        =   26
      Top             =   210
      Width           =   780
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "~"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   2190
      TabIndex        =   25
      Top             =   240
      Width           =   105
   End
End
Attribute VB_Name = "frmSo5290A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Table=SO033orSO034 A,SO014 B(,SO002 C)
Option Explicit
Dim strFrom As String
'Dim strST As String
Dim strFieldName As String
Dim strTableField As String

Private Sub cmdExit_Click()
  On Error GoTo ChkErr
    Unload Me
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdExit_Click")
End Sub

Private Sub cmdPrevRpt_Click()
  On Error GoTo ChkErr
    Screen.MousePointer = vbHourglass
      Call PreviousRpt(GetPrinterName(5), RptName("SO5290"), Me.Caption)
    Screen.MousePointer = vbDefault
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdPrevRpt_Click")
End Sub

Private Sub cmdPrint_Click()
  On Error GoTo ChkErr
    Dim rsTmp As New ADODB.Recordset
    Dim strSubQry1 As String
      If Not ChkDTok Then Exit Sub
      If Not IsDataOk Then Exit Sub
        Screen.MousePointer = vbHourglass
          cmdExit.SetFocus
          Me.Enabled = False
            ReadyGoPrint
            Call subChoose
            subCreateView
            CreateTable
            
            Set rsTmp = cnn.Execute("SELECT Count(*) as intCount From SO5290A2")
            If rsTmp("intCount") = 0 Then
                MsgNoRcd
                SendSQL , , True
            Else
                strsql = "Select * From SO5290A2"
                strSubQry1 = "Select * From SO5290A3"
                
                Call PrintRpt(GetPrinterName(5), RptName("SO5290"), , Me.Caption, strsql, strChooseString, , True, "Tmp0000.MDB", , , GiPaperLandscape, , strSubQry1)
            End If
          Me.Enabled = True
        Screen.MousePointer = vbDefault
        CloseRecordset rsTmp
    Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdPrint_Click")
End Sub

Private Function IsDataOk() As Boolean
  On Error GoTo ChkErr
    If gilCompCode.GetCodeNo = "" Then gilCompCode.SetFocus: strErrFile = "公司別": GoTo 66
    If gdaShouldDate1.GetValue = "" Then gdaShouldDate1.SetFocus: strErrFile = "應收日期起始日": GoTo 66
    If gdaShouldDate2.GetValue = "" Then gdaShouldDate2.SetFocus: strErrFile = "應收日期截止日": GoTo 66
    If gdaShouldDate2.GetValue(True) >= Format(DateAdd("m", 1, gdaShouldDate1.Text), "YYYY/MM/DD") Then
        MsgBox "收費期間不能超過一個月", , "錯誤訊息"
        gdaShouldDate2_GotFocus
        Exit Function
    End If
    IsDataOk = True
    Exit Function
66:
    MsgMustBe (strErrFile)
  Exit Function
ChkErr:
  Call ErrSub(Me.Name, "IsDataOk")
End Function

Private Sub subChoose()
  On Error GoTo ChkErr
  Dim strpagetype As String
  Dim strGroupString As String
  Dim blnSO002 As Boolean
  Dim blnSO014 As Boolean
    strChoose = "A.cancelFlag=0 And Nvl(A.STCode,-1) NOT IN (Select CodeNo From CD016 Where RefNo =1)"
    strChooseString = ""
    strFrom = " A"
    blnSO002 = False
    blnSO014 = False
    
  '日期
    If gdaShouldDate1.GetValue <> "" Then Call subAnd("A.ShouldDate >= To_Date('" & gdaShouldDate1.GetValue & "','YYYYMMDD')")
    If gdaShouldDate2.GetValue <> "" Then Call subAnd("A.ShouldDate < To_Date('" & gdaShouldDate2.GetValue & "','YYYYMMDD')+1")
    
  'GiList
    If gilCompCode.GetCodeNo <> "" Then Call subAnd("A.CompCode=" & gilCompCode.GetCodeNo)
    If gilServiceType.GetCodeNo <> "" Then Call subAnd("A.ServiceType='" & gilServiceType.GetCodeNo & "'")
    'If gilServiceType.GetCodeNo <> "" Then Call subAnd("Instr(Nvl(C.ServiceType,'" & gilServiceType.GetCodeNo & "'),'" & gilServiceType.GetCodeNo & "') > 0")
    
  'GiMulti
    'If gimCompCode.GetQryStr <> "" Then Call subAnd("C.CompCode" & gimCompCode.GetQryStr)
    If gimCitemCode.GetQryStr <> "" Then Call subAnd("A.CitemCode " & gimCitemCode.GetQryStr)
    If gimBillType.GetQryStr <> "" Then Call subAnd("SubStr(A.BillNo,7,1) " & gimBillType.GetQryStr)
    If gimClassCode.GetQryStr <> "" Then Call subAnd("A.ClassCode " & gimClassCode.GetQryStr)
    If gimStatusCode.GetQryStr <> "" Then Call subAnd("C.CustStatusCode " & gimStatusCode.GetQryStr): blnSO002 = True
    If gimClctEn.GetQryStr <> "" Then Call subAnd("A.ClctEn " & gimClctEn.GetQryStr)
    If gimCMCode.GetQryStr <> "" Then Call subAnd("A.CMCode " & gimCMCode.GetQryStr)
    If gimServCode.GetQryStr <> "" Then Call subAnd("A.ServCode " & gimServCode.GetQryStr)
    If gimAreaCode.GetQryStr <> "" Then Call subAnd("A.AreaCode " & gimAreaCode.GetQryStr)
    If gimStrtCode.GetQryStr <> "" Then Call subAnd("A.StrtCode " & gimStrtCode.GetQryStr)
    If gimUCCode.GetQryStr <> "" Then Call subAnd("A.UCCode " & gimUCCode.GetQryStr)
    
  '網路編號
    If mskCircuitNo.Text <> "" Then
       Call subAnd("B.CircuitNo='" & mskCircuitNo.Text & "'")
       blnSO014 = True
    Else
      If optCircuitNo.Value Then Call subAnd("B.CircuitNo is null "): blnSO014 = True
    End If
    
  '串SO014
    If blnSO014 Then
        strFrom = strFrom & ",SO014 B"
        strChoose = "A.AddrNo=B.AddrNo And " & strChoose
    End If
  '串SO001
    If blnSO002 Then
      strFrom = strFrom & ",SO002 C"
      strChoose = "A.Custid=C.Custid And A.ServiceType=C.ServiceType And " & strChoose
    End If
      
  '小計依據
    Select Case True
           Case optServCode.Value
                strGroupName = "D.Description"
                strpagetype = "服務區"
                strFrom = strFrom & ",CD002 D"
                strChoose = "A.ServCode=D.CodeNo(+) And " & strChoose
           Case optAreaCode.Value
                strGroupName = "D.Description"
                strpagetype = "行政區"
                strFrom = strFrom & ",CD001 D"
                strChoose = "A.AreaCode=D.CodeNo(+) And " & strChoose
           Case optClctAreaCode.Value
                strGroupName = "D.Description"
                strpagetype = "收費區"
                strFrom = strFrom & ",CD040 D"
                strChoose = "A.ClctAreaCode=D.CodeNo(+) And " & strChoose
           Case optStrtCode.Value
                strGroupName = "D.Description"
                strpagetype = "街道編號"
                strFrom = strFrom & ",CD017 D"
                strChoose = "A.StrtCode=D.CodeNo(+) And " & strChoose
           Case optClctEn.Value
                strGroupName = "A.ClctName"
                strpagetype = "收費人員"
    End Select
    
    strChooseString = "應收日期: " & subSpace(gdaShouldDate1.GetValue(True)) & "~" & subSpace(gdaShouldDate2.GetValue(True)) & ";" & _
                      "公司別　: " & subSpace(gilCompCode.GetDescription) & ";" & _
                      "服務類別: " & subSpace(gilServiceType.GetDescription) & ";" & _
                      "收費項目: " & subSpace(gimCitemCode.GetDispStr) & ";" & _
                      "單據編號: " & subSpace(gimBillType.GetDispStr) & ";" & _
                      "客戶類別: " & subSpace(gimClassCode.GetDispStr) & ";" & _
                      "客戶狀態: " & subSpace(gimStatusCode.GetDispStr) & ";" & _
                      "收費人員: " & subSpace(gimClctEn.GetDispStr) & ";" & _
                      "收費方式: " & subSpace(gimCMCode.GetDispStr) & ";" & _
                      "服務區　: " & subSpace(gimServCode.GetDispStr) & ";" & _
                      "行政區　: " & subSpace(gimAreaCode.GetDispStr) & ";" & _
                      "街道範圍: " & subSpace(gimStrtCode.GetDispStr) & ";" & _
                      "未繳費原因: " & subSpace(gimUCCode.GetDispStr) & ";" & _
                      "綱路編號: " & subSpace(mskCircuitNo.Text) & ";" & _
                      "小計依據: " & strpagetype
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "subChoose")
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error GoTo ChkErr
    Call FunctionKey(KeyCode, Shift, frmSo5290A)
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "Form_KeyDown")
End Sub

Private Sub Form_Load()
  On Error GoTo ChkErr
    If Alfa2 Then Call GetGlobal
    Call subGim
    Call subGil
    'gilCompCode.ListIndex = 1
    gilCompCode.SetCodeNo garyGi(9)
    gilCompCode.Query_Description
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "Form_Load")
End Sub

Private Sub subGim()
  On Error GoTo ChkErr
    Call SetgiMulti(gimCitemCode, "CodeNo", "Description", "CD019", "收費項目代碼", "收費項目名稱")
    Call SetgiMultiAddItem(gimBillType, "B,I,P,M,T", "收費單,裝機單,停拆移機單,維修單,臨時收費單")
    Call SetgiMulti(gimClassCode, "CodeNo", "Description", "CD004", "客戶類別代碼", "客戶類別名稱")
    Call SetgiMulti(gimStatusCode, "CodeNo", "Description", "CD035", "客戶狀態代碼", "客戶狀態名稱")
    Call SetgiMulti(Me.gimClctEn, "EmpNo", "EmpName", "CM003", "收費人員代碼", "收費人員名稱")
    Call SetgiMulti(gimCMCode, "CodeNo", "Description", "CD031", "收費方式代碼", "收費方式名稱")
    Call SetgiMulti(gimServCode, "CodeNo", "Description", "CD002", "服務區代碼", "服務區名稱")
    Call SetgiMulti(gimAreaCode, "CodeNo", "Description", "CD001", "行政區代碼", "行政區名稱")
    Call SetgiMulti(gimStrtCode, "CodeNo", "Description", "CD017", "街道代碼", "街道名稱")
    Call SetgiMulti(gimUCCode, "CodeNo", "Description", "CD013", "未繳費原因代碼", "未繳費原因名稱")
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "subGim")
End Sub

Private Sub subGil()
  On Error GoTo ChkErr
    'Call SetgiList(gilCompCode, "CodeNo", "Description", "CD039")
    Call SetgiList(gilCompCode, "CodeNo", "Description", "CD039", , , , , , , GetCompCodeFilter(gilCompCode))
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "subGil")
End Sub

Private Function subCreateView() As Boolean
  Dim strView As String
    On Error Resume Next
      gcnGi.Execute "Drop View " & strViewName
    On Error GoTo ChkErr
        strViewName = GetTmpViewName
        strView = "Create View " & strViewName & " as (" & _
                      "SELECT " & strGroupName & " GroupName,A.ShouldDate,A.BillNo,A.UCCode " & _
                          "From SO033 " & strFrom & " Where " & strChoose & _
                  " Union All " & _
                      "SELECT " & strGroupName & " GroupName,A.ShouldDate,A.BillNo,A.UCCode " & _
                          "From SO034 " & strFrom & " Where " & strChoose & ")"
        gcnGi.Execute strView
        SendSQL strView
    Exit Function
ChkErr:
  Call ErrSub(Me.Name, "subCreateView")
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  On Error Resume Next
  gcnGi.Execute "Drop View " & strViewName
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
  ReleaseCOM frmSo5290A
End Sub

Private Sub gdaShouldDate1_GotFocus()
  On Error Resume Next
  If gdaShouldDate1.GetValue = "" Then gdaShouldDate1.SetValue (Format(RightDate, "YYYYMM") & "01")
End Sub

Private Sub gdaShouldDate2_GotFocus()
  On Error Resume Next
    If gdaShouldDate1.GetValue <> "" Then
        gdaShouldDate2.SetValue GetLastDate(gdaShouldDate1.GetValue(True))
    Else
        gdaShouldDate2.SetValue ""
    End If
End Sub

Private Sub gdaShouldDate2_Validate(Cancel As Boolean)
  On Error Resume Next
  Cancel = ChkDate2(gdaShouldDate1, gdaShouldDate2)
End Sub

Private Sub gilCompCode_Change()
  On Error GoTo ChkErr
    Call SelectServType(gilCompCode.GetCodeNo, gilServiceType)
    gilServiceType.ListIndex = 1
    GiMultiFilter gimServCode, , gilCompCode.GetCodeNo
    GiMultiFilter gimAreaCode, , gilCompCode.GetCodeNo
    GiMultiFilter gimStrtCode, , gilCompCode.GetCodeNo
    GiMultiFilter gimClctEn, , gilCompCode.GetCodeNo
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "gilCompCode_Change")
End Sub

Private Sub gilCompCode_Validate(Cancel As Boolean)
  On Error GoTo ChkErr
    If gilCompCode.GetCodeNo = "" Then
        MsgMustBe ("公司別")
        Cancel = True
    End If
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "gilCompCode_Validate")
End Sub

Private Sub gilServiceType_Change()
  On Error GoTo ChkErr
    GiMultiFilter gimCitemCode, gilServiceType.GetCodeNo
    GiMultiFilter gimClassCode, gilServiceType.GetCodeNo
    GiMultiFilter gimCMCode, gilServiceType.GetCodeNo
    GiMultiFilter gimUCCode, gilServiceType.GetCodeNo
    'strST = GetCode("CD016", "<>1", gilServiceType.GetCodeNo, , , False)
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "gilCompCode_Validate")
End Sub

Private Sub mskCircuitNo_Validate(Cancel As Boolean)
  On Error Resume Next
    If mskCircuitNo.Text = "" Then
        optCircuitNo.Enabled = True
        optAll.Enabled = True
    Else
        optCircuitNo.Enabled = False
        optAll.Enabled = False
    End If
End Sub

Private Sub CreateTable()
  On Error Resume Next
    cnn.Execute "Drop Table SO5290A1"
    cnn.Execute "Drop Table SO5290A2"
    cnn.Execute "Drop Table SO5290A3"
  On Error GoTo ChkErr
    FieldName
    cnn.Execute "Create Table SO5290A1 (GroupName Text(40),TypeName Text(8)" & strTableField & ")"
    '應收單數
    Call subInsertMDB("Select GroupName,To_char(ShouldDate,'DD') as ShouldDay,Count(Distinct BillNo) as CountBill From " & _
                      strViewName & " Group by GroupName,To_char(ShouldDate,'DD')", "1應收單數")
    '己收單數
    Call subInsertMDB("Select GroupName,To_char(ShouldDate,'DD') as ShouldDay,Count(Distinct BillNo) as CountBill From " & _
                      strViewName & " Where UCCode is Null Group by GroupName,To_char(ShouldDate,'DD')", "2己收單數")
    '未收單數
    Call subInsertMDB("Select GroupName,To_char(ShouldDate,'DD') as ShouldDay,Count(Distinct BillNo) as CountBill From " & _
                      strViewName & " Where UCCode not in (Select CodeNo From CD013 Where RefNo=1) Group by GroupName,To_char(ShouldDate,'DD')", "3未收單數")
    '待收單數
    Call subInsertMDB("Select GroupName,To_char(ShouldDate,'DD') as ShouldDay,Count(Distinct BillNo) as CountBill From " & _
                      strViewName & " Where UCCode in (Select CodeNo From CD013 Where RefNo=1) Group by GroupName,To_char(ShouldDate,'DD')", "4待收單數")
    InsertTypeName
    SumMDB
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "CreateTable")
End Sub

Private Sub SumMDB()
  On Error GoTo ChkErr
  Dim strSQL1 As String
    cnn.BeginTrans
    strSQL1 = "Select GroupName,TypeName" & strFieldName & _
                " Into SO5290A2 From SO5290A1 Group by " & _
                " GroupName,TypeName"
    cnn.Execute strSQL1
    SendSQL strSQL1
    strSQL1 = "Select TypeName" & strFieldName & _
              " Into SO5290A3 From SO5290A2 Group by " & _
              " TypeName"
    cnn.Execute strSQL1
    SendSQL strSQL1, True
    cnn.CommitTrans
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "SumMDB")
End Sub

Private Sub InsertTypeName()
  On Error GoTo ChkErr
    Dim rsTmp As New ADODB.Recordset
    Dim intFor As Integer
    Dim arrType As Variant
    arrType = Array("1應收單數", "2己收單數", "3未收單數", "4待收單數")
    Set rsTmp = cnn.Execute("Select GroupName From SO5290A1 Group by GroupName")
    While Not rsTmp.EOF
        For intFor = 0 To 3
            cnn.Execute "Insert into SO5290A1(GroupName,TypeName) Values(" & _
                        GetNullString(rsTmp("GroupName")) & ",'" & _
                        arrType(intFor) & "')"
        Next
        rsTmp.MoveNext
    Wend
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "InsertTypeName")
End Sub

Private Sub subInsertMDB(strsql As String, strType As String)
  On Error GoTo ChkErr
  Dim rsTmp As New ADODB.Recordset
    Set rsTmp = gcnGi.Execute(strsql)
    SendSQL strsql
    cnn.BeginTrans
    While Not rsTmp.EOF
        cnn.Execute "Insert Into SO5290A1 (GroupName,TypeName,Bill" & Format(rsTmp("ShouldDay"), "00") & _
                     ") Values(" & _
                     GetNullString(rsTmp("GroupName")) & ",'" & _
                     strType & "'," & _
                     GetNullString(rsTmp("CountBill")) & ")"
        rsTmp.MoveNext
    Wend
    cnn.CommitTrans
    CloseRecordset rsTmp
  Exit Sub
ChkErr:
  cnn.RollbackTrans
  Call ErrSub(Me.Name, "subInsertMDB")
End Sub

Private Sub FieldName()
  On Error GoTo ChkErr
  Dim intFor As Integer
  strTableField = ""
  strFieldName = ""
    For intFor = 1 To 31
        strTableField = strTableField & ",Bill" & Format(intFor, "00") & " Long"
        strFieldName = strFieldName & ",iif(isnull(sum(Bill" & Format(intFor, "00") & ")),0,sum(Bill" & Format(intFor, "00") & ")) as Bill" & Format(intFor, "00")
    Next
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "FieldName")
End Sub

