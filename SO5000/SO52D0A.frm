VERSION 5.00
Object = "{A0C51418-555F-11D3-BC10-0050BA01A3A9}#9.8#0"; "GiMulti.ocx"
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{0CE6E756-92AE-11D3-8311-0080C8453DDF}#5.0#0"; "GiYM.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{074771B7-6ADE-11D5-B178-00485456B028}#1.17#0"; "csMulti.ocx"
Begin VB.Form frmSO52D0A 
   BorderStyle     =   1  '單線固定
   Caption         =   "應收帳明細表 [SO52D0A]"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7665
   Icon            =   "SO52D0A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   7665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '螢幕中央
   Begin CS_Multi.CSmulti gimStrtCode 
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   2070
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   661
      ButtonCaption   =   "街  道  範  圍"
   End
   Begin VB.Frame Frame1 
      Caption         =   "報表型式"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   735
      Left            =   60
      TabIndex        =   21
      Top             =   3270
      Width           =   3105
      Begin VB.OptionButton optStrt 
         Caption         =   "街道"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   465
         Left            =   1800
         TabIndex        =   13
         Top             =   210
         Width           =   1245
      End
      Begin VB.OptionButton optCustId 
         Caption         =   "客戶編號"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   465
         Left            =   240
         TabIndex        =   12
         Top             =   210
         Value           =   -1  'True
         Width           =   1245
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
      Left            =   1770
      TabIndex        =   15
      Top             =   4170
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
      Left            =   6240
      TabIndex        =   16
      Top             =   4170
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
      Left            =   90
      TabIndex        =   14
      Top             =   4170
      Width           =   1245
   End
   Begin Gi_Date.GiDate gdaShouldDate2 
      Height          =   345
      Left            =   2340
      TabIndex        =   2
      Top             =   450
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
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
   Begin Gi_Date.GiDate GiCreateTime2 
      Height          =   345
      Left            =   2340
      TabIndex        =   4
      Top             =   900
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
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
   Begin Gi_Date.GiDate GiCreateTime1 
      Height          =   345
      Left            =   990
      TabIndex        =   3
      Top             =   900
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
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
   Begin Gi_Multi.GiMulti gimClctAreaCode 
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   1350
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   661
      ButtonCaption   =   "收     費     區"
   End
   Begin Gi_Multi.GiMulti gimBillType 
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   1710
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   661
      ButtonCaption   =   "單  據  類  別"
      DataType        =   2
      DIY             =   -1  'True
      Exception       =   -1  'True
   End
   Begin Gi_Date.GiDate gdaShouldDate1 
      Height          =   345
      Left            =   990
      TabIndex        =   1
      Top             =   450
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
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
   Begin prjGiList.GiList gilCompCode 
      Height          =   315
      Left            =   4530
      TabIndex        =   5
      Top             =   150
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
   Begin prjGiList.GiList gilServiceType 
      Height          =   315
      Left            =   4530
      TabIndex        =   6
      Top             =   630
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
   Begin Gi_Multi.GiMulti gimOrder 
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   2790
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   661
      ButtonCaption   =   "排  序  方  式"
      DataType        =   2
      ColumnOrder     =   -1  'True
   End
   Begin Gi_YM.GiYM gymClctYM 
      Height          =   345
      Left            =   990
      TabIndex        =   0
      Top             =   30
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
   Begin CS_Multi.CSmulti gimClctEn 
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   2430
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   661
      ButtonCaption   =   "收  費  人   員"
   End
   Begin VB.Label lblClctYM 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "收集年月"
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
      Left            =   90
      TabIndex        =   24
      Top             =   120
      Width           =   780
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
      Left            =   3690
      TabIndex        =   23
      Top             =   210
      Width           =   765
   End
   Begin VB.Label Label3 
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
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   3690
      TabIndex        =   22
      Top             =   690
      Width           =   780
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "出單日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   90
      TabIndex        =   20
      Top             =   990
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "~"
      ForeColor       =   &H00FF0000&
      Height          =   165
      Left            =   2190
      TabIndex        =   19
      Top             =   1020
      Width           =   105
   End
   Begin VB.Label lblShouldDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "應收日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   90
      TabIndex        =   18
      Top             =   540
      Width           =   780
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "~"
      ForeColor       =   &H00FF0000&
      Height          =   165
      Left            =   2190
      TabIndex        =   17
      Top             =   570
      Width           =   105
   End
End
Attribute VB_Name = "frmSO52D0A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Table:SO033,SO001,SO014
Option Explicit

Private Sub cmdPrevRpt_Click()
  On Error GoTo ChkErr
    Screen.MousePointer = vbHourglass
      Call PreviousRpt(GetPrinterName(5), RptName("SO52D0"), Me.Caption)
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
  Dim strRptType As String
  Dim arrOrder() As String
  Dim varOrder As Variant
  Dim intSort As Integer
  
  strChoose = "Where SO033.AddrNo=SO014.AddrNo And SO033.CustId=SO001.CustId " & _
              "And SO033.RealDate is null And SO033.cancelFlag=0 "
  strChooseString = ""
  strGroupName = ""
  
  '收集年月
    If gymClctYM.GetValue <> "" Then Call subAnd("SO033.ClctYM = " & gymClctYM.GetValue)
    
  '日期
    If gdaShouldDate1.GetValue <> "" Then Call subAnd("SO033.ShouldDate >= To_Date('" & gdaShouldDate1.GetValue & "','YYYYMMDD')")
    If gdaShouldDate2.GetValue <> "" Then Call subAnd("SO033.ShouldDate < To_Date('" & gdaShouldDate2.GetValue & "','YYYYMMDD')+1")
    If GiCreateTime1.GetValue <> "" Then Call subAnd("SO033.CreateTime >= To_Date('" & GiCreateTime1.GetValue & "','YYYYMMDD')")
    If GiCreateTime2.GetValue <> "" Then Call subAnd("SO033.CreateTime < To_Date('" & GiCreateTime2.GetValue & "','YYYYMMDD')+1")
    
  'GiList
    If gilCompCode.GetCodeNo <> "" Then Call subAnd("SO033.CompCode=" & gilCompCode.GetCodeNo)
    If gilServiceType.GetCodeNo <> "" Then Call subAnd("SO033.ServiceType='" & gilServiceType.GetCodeNo & "'")
        
  'GiMulti
    If gimClctAreaCode.GetQryStr <> "" Then Call subAnd("SO001.ClctAreaCode " & gimClctAreaCode.GetQryStr)
    If gimBillType.GetQryStr <> "" Then Call subAnd("SubStr(SO033.BillNo,7,1) " & gimBillType.GetQryStr)
    If gimStrtCode.GetQryStr <> "" Then Call subAnd("SO014.StrtCode " & gimStrtCode.GetQryStr)
    If gimClctEn.GetQryStr <> "" Then Call subAnd("SO033.ClctEn " & gimClctEn.GetQryStr)
        
  '報表型式
      Select Case True
             Case optCustId.Value
                  Select Case Left(gimOrder.GetColumnOrderCode, 1)
                         Case "B"
                              strGroupName = "ReportType=False;GroupName={SO033.ClctEn}"
                         Case "C"
                              strGroupName = "ReportType=False;GroupName={SO014.AddrSort}"
                         Case Else
                              strGroupName = "ReportType=False;GroupName={SO001.CustID}"
                  End Select
                  strRptType = "客戶編號"
             Case OptStrt.Value
                  strGroupName = "ReportType=True;GroupName={SO014.StrtCode}"
                  strRptType = "街道"
      End Select
      
    '排序方式
      If gimOrder.GetColumnOrderCode <> "" Then
          arrOrder = Split(gimOrder.GetColumnOrderCode, ",")
          intSort = 0
          For Each varOrder In arrOrder
              Select Case arrOrder(intSort)
                Case "A"
                  strGroupName = strGroupName & ";Sort" & intSort & "={SO001.Custid}"
                Case "B"
                  strGroupName = strGroupName & ";Sort" & intSort & "={SO033.ClctEn}"
                Case "C"
                  strGroupName = strGroupName & ";Sort" & intSort & "={SO014.AddrSort}"
              End Select
              intSort = intSort + 1
          Next
      End If
    strChooseString = "收集年月:" & subSpace(gymClctYM.GetValue(True)) & ";" & _
                      "應收日期: " & subSpace(gdaShouldDate1.GetValue(True)) & "~" & subSpace(gdaShouldDate2.GetValue(True)) & ";" & _
                      "出單日期: " & subSpace(GiCreateTime1.GetValue(True)) & "~" & subSpace(GiCreateTime2.GetValue(True)) & ";" & _
                      "公司別　: " & subSpace(gilCompCode.GetDescription) & ";" & _
                      "服務類別: " & subSpace(gilServiceType.GetDescription) & ";" & _
                      "收 費 區: " & subSpace(gimClctAreaCode.GetDispStr) & ";" & _
                      "單據類別: " & subSpace(gimBillType.GetDispStr) & ";" & _
                      "街道範圍: " & subSpace(gimStrtCode.GetDispStr) & ";" & _
                      "收費人員: " & subSpace(gimClctEn.GetDispStr) & ";" & _
                      "排序方式: " & subSpace(gimOrder.GetColumnOrderDspStr) & ";" & _
                      "報表型式: " & strRptType

   Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "subChoose")
End Sub

Private Sub subPrint()
  On Error GoTo ChkErr
  Dim rsTmp As New ADODB.Recordset
    If rsTmp.State = 1 Then rsTmp.Close
'    rsTmp.CursorLocation = adUseClient
'    rsTmp.Open "Select count(*) as intCount From SO001,SO014,SO033 " & IIf(strChoose = "", "Where RowNum=1", strChoose & " And RowNum=1"), gcnGi, adOpenForwardOnly, adLockReadOnly
    strsql = "Select count(*) as intCount From SO001,SO014,SO033 " & IIf(strChoose = "", "Where RowNum=1", strChoose & " And RowNum=1")
    If Not GetRS(rsTmp, strsql, gcnGi) Then Exit Sub
    If rsTmp("intCount") = 0 Then
        MsgNoRcd
        SendSQL , , True
    Else
' SO033.BILLNO,SO033.REALDATE, SO033.REALAMT, SO033.CLCTNAME, SO033.QUANTITY,SO014.STRTCODE, SO014.STRTNAME, SO001.CUSTID, SO001.CUSTNAME,SO001.TEL1,SO001.CHARGEADDRESS,SO001.CLCTAREANAME
      strsql = "SELECT SO033.BILLNO,SO033.SHOULDDATE, SO033.SHOULDAMT, SO033.CLCTNAME, SO033.QUANTITY" & _
               ",SO014.STRTCODE, SO014.STRTNAME, SO001.CUSTID, SO001.CUSTNAME,SO001.TEL1,SO001.CHARGEADDRESS" & _
               ",SO001.CLCTAREANAME,SO033.ClctEn,SO014.AddrSort From SO001,SO014,SO033 " & strChoose
      Call PrintRpt(GetPrinterName(5), RptName("SO52D0"), , Me.Caption, strsql, strChooseString, , True, , , strGroupName, GiPaperPortrait)
    End If
    CloseRecordset rsTmp
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "subPrint")
End Sub

Private Sub Form_Load()
  On Error GoTo ChkErr
    Call subGim
    Call subGil
    gilCompCode.SetCodeNo garyGi(9)
    gilCompCode.Query_Description
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "Form_Load")
End Sub

Private Sub cmdExit_Click()
  On Error GoTo ChkErr
    Unload Me
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdExit_Click")
End Sub

Private Sub subGim()
  On Error GoTo ChkErr
    Call SetgiMulti(gimClctAreaCode, "CodeNo", "Description", "CD040", "收費區代碼", "收費區名稱")
    Call SetgiMultiAddItem(gimBillType, "B,I,P,M,T", "收費單,裝機單,停拆移機單,維修單,臨時收費單", "單據代號", "單據名稱")
    Call SetgiMulti(gimStrtCode, "CodeNo", "Description", "CD017", "街道代碼", "街道名稱")
    Call SetgiMulti(gimClctEn, "EmpNo", "EmpName", "CM003", "收費員代號", "收費員名稱")
    Call SetgiMultiAddItem(gimOrder, "A,B,C", "客戶編號,收費人員,收費地址")
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
  
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error GoTo ChkErr
    Call FunctionKey(KeyCode, Shift, frmSO52D0A)
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "Form_KeyDown")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  On Error Resume Next
  gcnGi.Execute "Drop View " & strViewName
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
  ReleaseCOM frmSO52D0A
End Sub

Private Sub gdaShouldDate1_GotFocus()
  On Error Resume Next
  If gdaShouldDate1.GetValue = "" Then gdaShouldDate1.SetValue RightDate
End Sub

Private Sub gdaShouldDate2_GotFocus()
  On Error Resume Next
  If gdaShouldDate1.GetValue = "" Or gdaShouldDate2.GetValue = "" Then gdaShouldDate2.SetValue (gdaShouldDate1.GetValue)
End Sub

Private Sub gdaShouldDate2_Validate(Cancel As Boolean)
  On Error Resume Next
  Cancel = ChkDate2(gdaShouldDate1, gdaShouldDate2)
  gdaShouldDate2_GotFocus
End Sub

Private Sub GiCreateTime1_GotFocus()
  On Error Resume Next
  If GiCreateTime1.GetValue = "" Then GiCreateTime1.SetValue (RightDate)
End Sub

Private Sub GiCreateTime2_GotFocus()
  On Error Resume Next
  If GiCreateTime1.GetValue = "" Or GiCreateTime2.GetValue = "" Then GiCreateTime2.SetValue (GiCreateTime1.GetValue)
End Sub

Private Sub GiCreateTime2_Validate(Cancel As Boolean)
  On Error Resume Next
  Cancel = ChkDate2(GiCreateTime1, GiCreateTime2)
End Sub

Private Sub gilCompCode_Change()
  On Error GoTo ChkErr
    Call SelectServType(gilCompCode.GetCodeNo, gilServiceType)
    gilServiceType.ListIndex = 1
    GiMultiFilter gimClctEn, , gilCompCode.GetCodeNo
    GiMultiFilter gimClctAreaCode, , gilCompCode.GetCodeNo
    GiMultiFilter gimStrtCode, , gilCompCode.GetCodeNo
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
