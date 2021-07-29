VERSION 5.00
Object = "{A0C51418-555F-11D3-BC10-0050BA01A3A9}#9.9#0"; "GiMulti.ocx"
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{074771B7-6ADE-11D5-B178-00485456B028}#2.1#0"; "csMulti.ocx"
Begin VB.Form frmSO5580A 
   BorderStyle     =   1  '單線固定
   Caption         =   "大樓戶增減數統計報表 [SO5580A]"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7665
   Icon            =   "SO5580A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   7665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '螢幕中央
   Begin VB.Frame Frame4 
      Caption         =   "網路編號"
      Height          =   735
      Left            =   0
      TabIndex        =   24
      Top             =   2970
      Width           =   7665
      Begin VB.TextBox mskCircuitNo 
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   90
         TabIndex        =   9
         Top             =   240
         Width           =   4605
      End
      Begin VB.OptionButton optAll 
         Caption         =   "全部"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   6660
         TabIndex        =   11
         Top             =   270
         Value           =   -1  'True
         Width           =   675
      End
      Begin VB.OptionButton optCircuitNo 
         Caption         =   "只印空白網路編號"
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   4830
         TabIndex        =   10
         Top             =   240
         Width           =   1845
      End
   End
   Begin CS_Multi.CSmulti gimStrtCode 
      Height          =   405
      Left            =   0
      TabIndex        =   8
      Top             =   2400
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   714
      ButtonCaption   =   "街  道  範  圍"
   End
   Begin VB.CommandButton cmdExcel 
      Caption         =   "匯成Excel"
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
      Left            =   3660
      TabIndex        =   14
      Top             =   4680
      Width           =   1245
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "F5:列印"
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
      Left            =   180
      TabIndex        =   12
      Top             =   4680
      Width           =   1245
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
      Left            =   6240
      TabIndex        =   15
      Top             =   4680
      Width           =   1245
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
      Left            =   1860
      TabIndex        =   13
      Top             =   4680
      Width           =   1395
   End
   Begin Gi_Date.GiDate gdaTime2 
      Height          =   345
      Left            =   2310
      TabIndex        =   1
      Top             =   90
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
      ForeColor       =   12582912
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
   Begin Gi_Date.GiDate gdaTime1 
      Height          =   345
      Left            =   960
      TabIndex        =   0
      Top             =   90
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
      ForeColor       =   12582912
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
   Begin Gi_Multi.GiMulti gimServCode 
      Height          =   405
      Left            =   0
      TabIndex        =   6
      Top             =   1680
      Width           =   7665
      _ExtentX        =   13520
      _ExtentY        =   714
      ButtonCaption   =   "服     務     區"
      FontSize        =   9
      FontName        =   "新細明體"
   End
   Begin Gi_Multi.GiMulti gimAreaCode 
      Height          =   405
      Left            =   0
      TabIndex        =   7
      Top             =   2040
      Width           =   7665
      _ExtentX        =   13520
      _ExtentY        =   714
      ButtonCaption   =   "行     政     區"
      FontSize        =   9
      FontName        =   "新細明體"
   End
   Begin prjGiList.GiList gilServiceType 
      Height          =   315
      Left            =   4590
      TabIndex        =   3
      Top             =   510
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
      F5Corresponding =   -1  'True
   End
   Begin prjGiList.GiList gilCompCode 
      Height          =   315
      Left            =   4590
      TabIndex        =   2
      Top             =   90
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
      F5Corresponding =   -1  'True
   End
   Begin CS_Multi.CSmulti gimMduId 
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   960
      Width           =   7665
      _ExtentX        =   13520
      _ExtentY        =   661
      ButtonCaption   =   "大  樓  編  號"
   End
   Begin CS_Multi.CSmulti gimClassCode1 
      Height          =   405
      Left            =   0
      TabIndex        =   5
      Top             =   1320
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   714
      ButtonCaption   =   "客  戶  類  別"
   End
   Begin VB.Label lblHelp 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "說明:"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   150
      TabIndex        =   23
      Top             =   3930
      Width           =   510
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "裝機時-->裝機時間     拆機時-->拆機時間"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   1620
      TabIndex        =   22
      Top             =   3960
      Width           =   3165
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "停機時-->停機時間     註銷時-->註銷時間"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   1620
      TabIndex        =   21
      Top             =   4170
      Width           =   3165
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "日期範圍:"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   720
      TabIndex        =   20
      Top             =   3960
      Width           =   825
   End
   Begin VB.Label lblCompCode 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BackStyle       =   0  '透明
      Caption         =   "公  司  別"
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
      Left            =   3750
      TabIndex        =   19
      Top             =   150
      Width           =   765
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BackStyle       =   0  '透明
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
      Height          =   195
      Left            =   3750
      TabIndex        =   18
      Top             =   555
      Width           =   780
   End
   Begin VB.Label lblInstTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "日期範圍"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   30
      TabIndex        =   17
      Top             =   180
      Width           =   780
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "~"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   2160
      TabIndex        =   16
      Top             =   180
      Width           =   165
   End
End
Attribute VB_Name = "frmSO5580A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Table=SO001,SO002,SO014
Option Explicit
Dim MdbChoose As String
Dim blnExcel As Boolean

Private Sub cmdExcel_Click()
  On Error GoTo ChkErr
    blnExcel = True
    cmdPrint_Click
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdExcel_Click")
End Sub

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
    Call PreviousRpt(GetPrinterName(5), RptName("SO5580"), Me.Caption)
    Screen.MousePointer = vbDefault
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "cmdPrevRpt_Click")
End Sub

Private Sub cmdPrint_Click()
  On Error GoTo ChkErr
    If Not ChkDTok Then Exit Sub
    If Not IsDataOk Then Exit Sub
    Screen.MousePointer = vbHourglass
      cmdExit.SetFocus
      Me.Enabled = False
        ReadyGoPrint
        Call subChoose
        Call SQLView
        Call subPrint
      Me.Enabled = True
    Screen.MousePointer = vbDefault
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "cmdPrint_Click")
End Sub

Private Function IsDataOk() As Boolean
  On Error GoTo ChkErr
    If gilCompCode.GetCodeNo = "" Then gilCompCode.SetFocus: strErrFile = "公司別": GoTo 66
    If gilServiceType.GetCodeNo = "" Then gilServiceType.SetFocus: strErrFile = "服務類別": GoTo 66
    IsDataOk = True
    Exit Function
66:
    MsgMustBe (strErrFile)
  Exit Function
ChkErr:
    Call ErrSub(Me.Name, "IsDataOk")
End Function

Private Sub subPrint()
  On Error GoTo ChkErr
  Dim rsTmp As New ADODB.Recordset
    Set rsTmp = gcnGi.Execute("SELECT Count(*) as intCount From " & strViewName)
      If rsTmp("intCount") = 0 Then
         MsgNoRcd
         SendSQL , , True
      Else
         strSQL = "Select * From " & strViewName & " V"
         If blnExcel Then
            Call toExcel(strSQL)
         Else
            Call PrintRpt(GetPrinterName(5), RptName("SO5580"), , Me.Caption, strSQL, strChooseString, , True, , , , GiPaperLandscape)
         End If
      End If
    CloseRecordset rsTmp
    blnExcel = False
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "subPrint")
End Sub

Private Sub toExcel(ByVal strSQL As String)
  On Error GoTo ChkErr
    Dim rsExcel As New ADODB.Recordset
      RptToTxt RptName("SO5580"), , strSQL, strChooseString, , , , , , garyGi(19) & "\SO5580"
      If Not Get_RS_From_Txt(garyGi(19), "SO5580.txt", rsExcel) Then blnExcel = False: Exit Sub
       '#7074 To Add a choice-condition parameter into excel print function by Kin 2015/09/09
      Call UseProperty(rsExcel, "大樓戶增減數統計報表", "第一頁", _
            , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , strChooseString)
      blnExcel = False
      CloseRecordset rsExcel
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "toExcel")
End Sub

Private Sub subChoose()
  On Error GoTo ChkErr
  Dim blnFlag As Boolean
    'strChoose = " From SO001,SO014 Where SO001.InstAddrNo=SO014.AddrNo And SO001.MduId is Not Null "
    strChoose = ""
    strChooseString = ""
    blnFlag = False
    
  'GiMulti
    'If gimCompCode.GetQryStr <> "" Then Call subAnd("SO001.CompCode " & gimCompCode.GetQryStr)
    '問題集 2584多串一個客戶類別
    If gimClassCode1.GetQryStr <> "" And gimClassCode1.GetQryStr <> "全選" Then
        Call subAnd("SO001.ClassCode1 " & gimClassCode1.GetQryStr)
    End If
    If gimServCode.GetQryStr <> "" Then Call subAnd("SO001.ServCode " & gimServCode.GetQryStr)
    If gimAreaCode.GetQryStr <> "" Then blnFlag = True: Call subAnd("SO014.AreaCode " & gimAreaCode.GetQryStr)
    If gimStrtCode.GetQryStr <> "" Then blnFlag = True: Call subAnd("SO014.StrtCode " & gimStrtCode.GetQryStr)
    If gimMduId.GetQryStr <> "" Then Call subAnd("SO001.MduId " & gimMduId.GetQryStr)
      
  'GiList
    If gilCompCode.GetCodeNo <> "" Then Call subAnd("SO001.CompCode=" & gilCompCode.GetCodeNo)
    If gilServiceType.GetCodeNo <> "" Then Call subAnd("instr(SO001.ServiceType,'" & gilServiceType.GetCodeNo & "')>0")
    If gilServiceType.GetCodeNo <> "" Then Call subAnd("SO002.ServiceType='" & gilServiceType.GetCodeNo & "'")
    'If gilServiceType.GetCodeNo <> "" Then Call subAnd("Instr(Nvl(SO001.ServiceType,'" & gilServiceType.GetCodeNo & "'),'" & gilServiceType.GetCodeNo & "') > 0")
    
  '網路編號
    If mskCircuitNo.Text <> "" Then
       blnFlag = True
       Call subAnd("SO014.CircuitNo='" & mskCircuitNo.Text & "'")
    Else
      blnFlag = True:
      If optCircuitNo.Value Then Call subAnd("SO014.CircuitNo is null ")
    End If
    
    If blnFlag Then
      strChoose = " From SO001,SO002,SO017,SO014 Where SO001.MDUId=SO017.MDUId And SO001.Custid=SO002.Custid And SO001.InstAddrNo=SO014.AddrNo And SO001.MduId is Not Null " & IIf(strChoose = "", "", " And " & strChoose)
    Else
      strChoose = " From SO001,SO002,SO017 Where SO001.MDUId=SO017.MDUId And SO001.Custid=SO002.Custid And SO001.MduId is Not Null " & IIf(strChoose = "", "", " And " & strChoose)
    End If

   strChooseString = "日期範圍:" & subSpace(gdaTime1.GetValue(True)) & "~" & subSpace(gdaTime2.GetValue(True)) & ";" & _
                     "公司別　:" & subSpace(gilCompCode.GetDescription) & ";" & _
                     "服務類別: " & subSpace(gilServiceType.GetDescription) & ";" & _
                     "大樓名稱:" & subSpace(gimMduId.GetDispStr) & ";" & _
                     "服務區　:" & subSpace(gimServCode.GetDispStr) & ";" & _
                     "行政區　:" & subSpace(gimAreaCode.GetDispStr) & ";" & _
                     "街道範圍:" & subSpace(gimStrtCode.GetDispStr) & ";" & _
                     "綱路編號: " & subSpace(mskCircuitNo.Text)
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "subChoose")
End Sub

Private Sub MdbAnd(strCH As String)
  On Error GoTo ChkErr
    If MdbChoose <> "" Then
       MdbChoose = MdbChoose & " And " & strCH
    Else
       MdbChoose = strCH
    End If
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "MdbAnd")
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error GoTo ChkErr
    Call FunctionKey(KeyCode, Shift, frmSO5580A)
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "Form_KeyDown")
End Sub

Private Sub Form_Load()
  On Error GoTo ChkErr
    Call subGim
    Call subGil
    gilCompCode.SetCodeNo garyGi(9)
    gilCompCode.Query_Description
    '問題集2584 客戶類別預設帶為全選
    gimClassCode1.SelectAll
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "Form_Load")
End Sub

Private Sub subGil()
  On Error GoTo ChkErr
    'Call SetgiList(gilCompCode, "CodeNo", "Description", "CD039")
    Call SetgiList(gilCompCode, "CodeNo", "Description", "CD039", , , , , , , GetCompCodeFilter(gilCompCode))
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "subGil")
End Sub

Private Sub subGim()
  On Error GoTo ChkErr
    Call SetgiMulti(gimMduId, "MduId", "Name", "SO017", "大樓代號", "大樓名稱")
    '問題集2584 新增一個客戶類別
    Call SetgiMulti(gimClassCode1, "CodeNo", "Description", "CD004", "客戶類別代碼", "客戶類別名稱")
    Call SetgiMulti(gimServCode, "CodeNo", "Description", "CD002", "服務區代碼", "服務區名稱")
    Call SetgiMulti(gimAreaCode, "CodeNo", "Description", "CD001", "行政區代碼", "行政區名稱")
    Call SetgiMulti(gimStrtCode, "CodeNo", "Description", "CD017", "街道代碼", "街道名稱")
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "subGim")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  On Error Resume Next
  gcnGi.Execute "Drop View " & strViewName
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
  ReleaseCOM frmSO5580A
End Sub

Private Sub gdaTime1_GotFocus()
  On Error Resume Next
  If gdaTime1.GetValue = "" Then gdaTime1.SetValue (RightDate)
End Sub

Private Sub gdaTime2_GotFocus()
  On Error Resume Next
  If gdaTime1.GetValue = "" Or gdaTime2.GetValue = "" Then gdaTime2.SetValue (gdaTime1.GetValue)
End Sub

Private Sub gdaTime2_Validate(Cancel As Boolean)
  On Error Resume Next
  Cancel = ChkDate2(gdaTime1, gdaTime2)
End Sub

Private Sub SQLView()
  On Error GoTo ChkErr
  Dim strSQL As String
    MdbChoose = ""
      '正常
        MdbChoose = " And SO002.CustStatusCode=1 "
        strSQL = " SELECT SO017.MDUId,SO017.Name,Count(Distinct SO001.Custid) as NormalCount,0 as InstCount,0 as StopCount," & _
                 "0 as PRCount,0 as DelCount,0 as PromCount,0 as PRingCount " & strChoose & MdbChoose & " Group by SO017.MDUId,SO017.Name"
        
      '裝機
        MdbChoose = " And SO002.CustStatusCode=1 "
        If gdaTime1.GetValue <> "" Then Call MdbAnd("SO002.InstTime >= To_Date('" & gdaTime1.GetValue & "','YYYYMMDD')")
        If gdaTime2.GetValue <> "" Then Call MdbAnd("SO002.InstTime < To_Date('" & gdaTime2.GetValue & "','YYYYMMDD')+1")
        strSQL = strSQL & " Union All SELECT SO017.MDUId,SO017.Name,0 as NormalCount,Count(Distinct SO001.Custid) as InstCount,0 as StopCount," & _
                 "0 as PRCount,0 as DelCount,0 as PromCount,0 as PRingCount " & strChoose & MdbChoose & " Group by SO017.MDUId,SO017.Name"
        
      '停機
        MdbChoose = " And SO002.CustStatusCode=2 "
        If gdaTime1.GetValue <> "" Then Call MdbAnd("SO002.StopTime >= To_Date('" & gdaTime1.GetValue & "','YYYYMMDD')")
        If gdaTime2.GetValue <> "" Then Call MdbAnd("SO002.StopTime < To_Date('" & gdaTime2.GetValue & "','YYYYMMDD')+1")
        strSQL = strSQL & " Union All SELECT SO017.MDUId,SO017.Name,0 as NormalCount,0 InstCount,Count(Distinct SO001.Custid) as StopCount," & _
                 "0 as PRCount,0 as DelCount,0 as PromCount,0 as PRingCount " & strChoose & MdbChoose & " Group by SO017.MDUId,SO017.Name"
        
      '拆機
        MdbChoose = " And SO002.CustStatusCode=3 "
        If gdaTime1.GetValue <> "" Then Call MdbAnd("SO002.PRTime >= To_Date('" & gdaTime1.GetValue & "','YYYYMMDD')")
        If gdaTime2.GetValue <> "" Then Call MdbAnd("SO002.PRTime < To_Date('" & gdaTime2.GetValue & "','YYYYMMDD')+1")
        strSQL = strSQL & " Union All SELECT SO017.MDUId,SO017.Name,0 as NormalCount,0 as InstCount,0 as StopCount," & _
                    "Count(Distinct SO001.Custid) as PRCount,0 as DelCount,0 as PromCount,0 as PRingCount " & strChoose & MdbChoose & " Group by SO017.MDUId,SO017.Name"
                    
      '註銷
        MdbChoose = " And SO002.CustStatusCode=4 "
        If gdaTime1.GetValue <> "" Then Call MdbAnd("SO002.DelDate >= To_Date('" & gdaTime1.GetValue & "','YYYYMMDD')")
        If gdaTime2.GetValue <> "" Then Call MdbAnd("SO002.DelDate < To_Date('" & gdaTime2.GetValue & "','YYYYMMDD')+1")
        strSQL = strSQL & " Union All SELECT SO017.MDUId,SO017.Name,0 as NormalCount,0 as InstCount,0 as StopCount," & _
                    "0 as PRCount,Count(Distinct SO001.Custid) as DelCount,0 as PromCount,0 as PRingCount " & strChoose & MdbChoose & " Group by SO017.MDUId,SO017.Name"
                    
      '促銷,拆機中
       ' MdbChoose = " And SO002.CustStatusCode=5 "
        strSQL = strSQL & " Union All SELECT SO017.MDUId,SO017.Name,0 as NormalCount,0 as InstCount,0 as StopCount,0 as PRCount,0 as DelCount," & _
                    "Count(Distinct DeCode(CustStatusCode,5,SO001.Custid,Null)) as PromCount,Count(Distinct DeCode(CustStatusCode,6,SO001.Custid,Null)) as PRingCount " & _
                    strChoose & " Group by SO017.MDUId,SO017.Name"
        Call subCreateView(strSQL)
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "SQLView")
End Sub

Private Sub subCreateView(strSQL As String)
  On Error Resume Next
    gcnGi.Execute "Drop View " & strViewName
  On Error GoTo ChkErr
    strViewName = GetTmpViewName
  
    strSQL = "Create View " & strViewName & " as (" & _
                  "Select A.MDUId,A.Name MDUName,Sum(NormalCount) as NormalCount,Sum(A.InstCount) as InstCount,Sum(A.StopCount) as StopCount," & _
                  "Sum(A.PRCount) as PRCount,Sum(A.DelCount) as DelCount,Sum(A.PromCount) as PromCount,Sum(A.PRingCount) as PRingCount From (" & strSQL & _
                  ") A Group by A.MDUId,A.Name) "
    gcnGi.Execute strSQL
    SendSQL strSQL, True
  Exit Sub
ChkErr:
    ErrSub Me.Name, "subCreateView"
End Sub

'Private Sub SQLMDB()
'  On Error GoTo ChkErr
'  Dim I As Integer
'    cnn.Execute "DELETE * FROM SO5580A"
'    MdbChoose = ""
'    '日期
'    blnflag = False
'    If gdaTime1.GetValue <> "" Or gdaTime2.GetValue <> "" Then
'      For I = 1 To 4
'        Select Case I
'          Case 1 '裝機
'            'MdbChoose = " SO001.CustStatusCode in (1,6) "
'            MdbChoose = ""
'            If gdaTime1.GetValue <> "" Then Call MdbAnd("SO001.InstTime >= To_Date('" & gdaTime1.GetValue & "','YYYYMMDD')")
'            If gdaTime2.GetValue <> "" Then Call MdbAnd("SO001.InstTime < To_Date('" & gdaTime2.GetValue & "','YYYYMMDD')+1")
'
'          Case 2 '停機
'            'MdbChoose = " SO001.CustStatusCode=2 "
'            MdbChoose = ""
'            If gdaTime1.GetValue <> "" Then Call MdbAnd("SO001.StopTime >= To_Date('" & gdaTime1.GetValue & "','YYYYMMDD')")
'            If gdaTime2.GetValue <> "" Then Call MdbAnd("SO001.StopTime < To_Date('" & gdaTime2.GetValue & "','YYYYMMDD')+1")
'
'          Case 3 '拆機
'            'MdbChoose = " SO001.CustStatusCode=3 "
'            MdbChoose = ""
'            If gdaTime1.GetValue <> "" Then Call MdbAnd("SO001.PRTime >= To_Date('" & gdaTime1.GetValue & "','YYYYMMDD')")
'            If gdaTime2.GetValue <> "" Then Call MdbAnd("SO001.PRTime < To_Date('" & gdaTime2.GetValue & "','YYYYMMDD')+1")
'
'          Case 4 '註銷
'            'MdbChoose = " SO001.CustStatusCode=4 "
'            MdbChoose = ""
'            If gdaTime1.GetValue <> "" Then Call MdbAnd("SO001.DelDate >= To_Date('" & gdaTime1.GetValue & "','YYYYMMDD')")
'            If gdaTime2.GetValue <> "" Then Call MdbAnd("SO001.DelDate < To_Date('" & gdaTime2.GetValue & "','YYYYMMDD')+1")
'
'        End Select
'
'
'
'        If strChoose <> "" Then
'          MdbChoose = " And " & MdbChoose & " And " & strChoose
'        Else
'          MdbChoose = " And " & MdbChoose
'        End If
'        If I = 4 Then blnflag = True
'        subInsertMDB
'
'      Next
'    Else
'      If strChoose <> "" Then
'        MdbChoose = " And " & strChoose
'      End If
'      blnflag = True
'      subInsertMDB
'    End If
'
'  Exit Sub
'ChkErr:
'    Call ErrSub(Me.Name, "SQLMDB")
'End Sub

'Private Sub subInsertMDB()
'    On Error GoTo ChkErr
'    Dim rsTmp As New adodb.Recordset
'    Dim strMDB As String
'
'      If cnn.State = adStateClosed Then Set cnn = GetTmpMdbCn
'      If rsTmp.State = 1 Then rsTmp.Close
'        strMDB = "SELECT SO001.Custid,SO001.CustStatusCode,SO014.MDUId,SO014.MDUName" & _
'                    " From SO001,SO014 Where SO001.InstAddrNo=SO014.AddrNo " & _
'                    MdbChoose & ""
'        Set rsTmp = gcnGi.Execute(strMDB)
'        If blnflag Then
'          SendSQL strMDB, True
'        Else
'          SendSQL strMDB
'        End If
'        cnn.BeginTrans
'        '(Custid,CustStatusCode,AreaCode,AreaName,ServCode,ServName)
'          While Not rsTmp.EOF
'            cnn.Execute "INSERT INTO SO5580A " & _
'                    "VALUES (" & _
'                    GetNullString(rsTmp("Custid")) & "," & _
'                    GetNullString(rsTmp("CustStatusCode")) & "," & _
'                    GetNullString(rsTmp("MDUId")) & "," & _
'                    GetNullString(rsTmp("MDUName")) & ")"
'            rsTmp.MoveNext
'            DoEvents
'
'          Wend
'        cnn.CommitTrans
'        Set rsTmp = Nothing
'
'    Exit Sub
'ChkErr:
'  cnn.RollbackTrans
'    Call ErrSub(Me.Name, "subInsertMDB")
'End Sub

Private Sub gilCompCode_Change()
  On Error GoTo ChkErr
    Call SelectServType(gilCompCode.GetCodeNo, gilServiceType)
    gilServiceType.ListIndex = 1
    GiMultiFilter gimServCode, , gilCompCode.GetCodeNo
    GiMultiFilter gimAreaCode, , gilCompCode.GetCodeNo
    GiMultiFilter gimStrtCode, , gilCompCode.GetCodeNo
    GiMultiFilter gimMduId, , gilCompCode.GetCodeNo
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

Private Sub mskCircuitNo_Validate(Cancel As Boolean)
  On Error GoTo ChkErr
  If mskCircuitNo.Text = "" Then
    optCircuitNo.Enabled = True
    optAll.Enabled = True
  Else
    optCircuitNo.Enabled = False
    optAll.Enabled = False
  End If
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "mskCircuitNo_Validate")
End Sub
