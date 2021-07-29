VERSION 5.00
Object = "{A0C51418-555F-11D3-BC10-0050BA01A3A9}#9.9#0"; "GiMulti.ocx"
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Begin VB.Form frmSO55B0A 
   BorderStyle     =   1  '單線固定
   Caption         =   "證件繳付一覽表[SO55B0A]"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8010
   Icon            =   "SO55B0A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   8010
   StartUpPosition =   1  '所屬視窗中央
   Begin Gi_Multi.GiMulti gimServiceType 
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   840
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   661
      ButtonCaption   =   "服務類別"
      DataType        =   2
      FontSize        =   9
      FontName        =   "新細明體"
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
      Left            =   1920
      TabIndex        =   5
      Top             =   2160
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
      Left            =   6120
      TabIndex        =   7
      Top             =   2160
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
      Left            =   240
      TabIndex        =   4
      Top             =   2160
      Width           =   1245
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
      Left            =   3960
      TabIndex        =   6
      Top             =   2160
      Width           =   1245
   End
   Begin Gi_Multi.GiMulti GimInstCode 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   661
      ButtonCaption   =   "裝機派工類別"
      FontSize        =   9
      FontName        =   "新細明體"
   End
   Begin Gi_Date.GiDate gdaFinTime2 
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   240
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      ForeColor       =   0
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
   Begin Gi_Date.GiDate gdaFinTime1 
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      ForeColor       =   0
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
      Left            =   4920
      TabIndex        =   2
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
      F5Corresponding =   -1  'True
   End
   Begin VB.Label lblCompCode 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "公  司  別"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   4080
      TabIndex        =   10
      Top             =   330
      Width           =   765
   End
   Begin VB.Label Label1 
      Caption         =   "裝機完成日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "~"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2640
      TabIndex        =   8
      Top             =   360
      Width           =   105
   End
End
Attribute VB_Name = "frmSO55B0A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim blnExcel As Boolean
Private strViewName As String
Private SubRptname As String

Private Sub cmdExcel_Click()
  On Error GoTo ChkErr
    blnExcel = True
    cmdPrint_Click
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdExcel_Click")
End Sub

Private Sub cmdExit_Click() '結束
  On Error GoTo ChkErr
    Unload Me
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdExit_Click")
End Sub

Private Sub cmdPrevRpt_Click() '上次統計結果
  On Error GoTo ChkErr
    Screen.MousePointer = vbHourglass
      Call PreviousRpt(GetPrinterName(5), RptName("SO55B0"), Me.Caption)
    Screen.MousePointer = vbDefault
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdPrevRpt_Click")
End Sub

Private Sub cmdPrint_Click() '列印
  On Error GoTo ChkErr
    Dim rsTmp As New ADODB.Recordset
    Dim rsExcel As New ADODB.Recordset
    Dim strsubview As String
    Dim strsubsql As String
      If Not ChkDTok Then Exit Sub
      If Not IsDataOk Then Exit Sub
      Screen.MousePointer = vbHourglass
        cmdExit.SetFocus
        Me.Enabled = False
          ReadyGoPrint
          If Not subChoose Then Exit Sub
          If Not subCreateView Then Exit Sub
          
          Set rsTmp = gcnGi.Execute("SELECT Count(*) as intCount FROM " & strViewName & " V")
          If rsTmp("intCount") = 0 Then
              MsgNoRcd
              SendSQL , , True
          Else
              strSQL = "SELECT * FROM " & strViewName & " V " & "order by servicetype,Faciname"
               On Error Resume Next
              gcnGi.Execute "Drop View " & SubRptname
                  On Error GoTo ChkErr
              SubRptname = GetTmpViewName
              strsubview = "SELECT SERVICETYPE,IDKINDNAME1,COUNT(*) INTCOUNT FROM (Select SERVICETYPE,IDKINDNAME1 from " & strViewName & " V " & " Where IDKINDNAME1 IS NOT NULL " & _
                       "UNION ALL Select SERVICETYPE,IDKINDNAME IDKINDNAME1 from " & strViewName & " V " & " Where IDKINDNAME IS NOT NULL ) GROUP BY SERVICETYPE,IDKINDNAME1"
                
               strsubview = "Create View " & SubRptname & " As (" & strsubview & ")"
               gcnGi.Execute strsubview
               SendSQL strsubview, True
               
               strsubsql = "SELECT * FROM " & SubRptname & " S " & "order by servicetype"
                   
               
              
              If blnExcel Then
                  Call toExcel(strSQL, strsubsql)
              Else
                  Call PrintRpt(GetPrinterName(5), RptName("SO55B0"), , Me.Caption, strSQL, strChooseString, , True, , , , , , strsubsql)
             End If
          End If
          CloseRecordset rsTmp
          CloseRecordset rsExcel
        Me.Enabled = True
      Screen.MousePointer = vbDefault
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdPrint_Click")
End Sub

Private Function IsDataOk() As Boolean '檢查資料是否正確
  On Error GoTo ChkErr
    If gilCompCode.GetCodeNo = "" Then gilCompCode.SetFocus: strErrFile = "公司別": GoTo 66
    If gdaFinTime1.GetValue = "" Then gdaFinTime1.SetFocus: strErrFile = "完工起始日": GoTo 66
    If gdaFinTime2.GetValue = "" Then gdaFinTime2.SetFocus: strErrFile = "完工截止日": GoTo 66
    IsDataOk = True
  Exit Function
66:
  MsgMustBe (strErrFile)
  Exit Function
ChkErr:
  Call ErrSub(Me.Name, "IsDataOk")
End Function
Private Sub toExcel(ByVal strSQL As String, ByVal strsubsql As String)
  On Error GoTo ChkErr
    Dim rsExcel As New ADODB.Recordset
    Dim rsSubExcel1 As New ADODB.Recordset
    
    RptToTxt RptName("SO55B0", "E"), , strSQL, strChooseString, , , strGroupName, , , garyGi(19) & "\SO55B0", , , strsubsql
    If Not Get_RS_From_Txt(garyGi(19), "SO55B0.txt", rsExcel) Then blnExcel = False: Exit Sub
    Call GetRS(rsSubExcel1, strsubsql)
     '#7074 To Add a choice-condition parameter into excel print function by Kin 2015/09/09
    Call UseProperty(rsExcel, "證件繳付一覽表", "第一頁", , , rsSubExcel1, _
                    , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , strChooseString)
    
    blnExcel = False
    CloseRecordset rsExcel
    CloseRecordset rsSubExcel1
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "toExcel")
End Sub

Private Function subChoose() As Boolean '篩選條件
  On Error GoTo ChkErr
    strChoose = ""
    strChooseString = ""
     
  '日期
    If gdaFinTime1.GetValue <> "" Then Call subAnd("SO007.FinTime >= To_Date('" & gdaFinTime1.GetValue & "','YYYYMMDD')")
    If gdaFinTime2.GetValue <> "" Then Call subAnd("SO007.FinTime < To_Date('" & gdaFinTime2.GetValue & "','YYYYMMDD') +1")
  'GiMulti
    'If GimInstCode.GetQryStr <> "" Then Call subAnd("SO004.InstCode " & GimInstCode.GetQryStr)
    If GimInstCode.GetQueryCode <> "" Then Call subAnd("SO007.InstCode in (" & GimInstCode.GetQueryCode & ")")
    If gimServiceType.GetQryStr <> "" Then Call subAnd("SO004.ServiceType" & gimServiceType.GetQryStr & "")
  'GiList
    If gilCompCode.GetCodeNo <> "" Then Call subAnd("SO004.CompCode=" & gilCompCode.GetCodeNo)
    

    strChooseString = "完工日期:" & subSpace(gdaFinTime1.GetOriginalValue) & "~" & subSpace(gdaFinTime2.GetOriginalValue) & ";" & _
                      "公司別　: " & subSpace(gilCompCode.GetDescription) & ";" & _
                      "服務類別: " & subSpace(gimServiceType.GetDispStr) & ";" & _
                      "裝機類別:" & subSpace(GimInstCode.GetDispStr)
    subChoose = True
  Exit Function
ChkErr:
  Call ErrSub(Me.Name, "subChoose")
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error GoTo ChkErr
    Call FunctionKey(KeyCode, Shift, frmSO55B0A)
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
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "Form_Load")
End Sub

Private Sub subGim() 'GiMulti
  On Error GoTo ChkErr
    Call SetgiMulti(gimServiceType, "CodeNO", "Description", "CD046", "服務類別代碼", "服務類別名稱")
    Call SetgiMulti(GimInstCode, "CodeNo", "Description", "CD005", "裝機類別代碼", "裝機類別名稱")
    gimServiceType.SetQueryCode "'I','P'"

    GimInstCode.Filter = " WHERE SERVICETYPE " & gimServiceType.GetQryStr
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "subGim")
End Sub

Private Sub subGil() 'GiList
  On Error GoTo ChkErr
    Call SetgiList(gilCompCode, "CodeNo", "Description", "CD039", , , , , , , GetCompCodeFilter(gilCompCode))
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "subGil")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  On Error Resume Next
    gcnGi.Execute "Drop View " & strViewName
    gcnGi.Execute "Drop View " & SubRptname
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
    ReleaseCOM Me
End Sub

Private Sub gdaFinTime1_GotFocus()
  If gdaFinTime1.GetValue = "" Then gdaFinTime1.SetValue (RightDate)
End Sub

Private Sub gdaFinTime2_GotFocus()
  On Error GoTo ChkErr
    If gdaFinTime1.GetValue = "" Or gdaFinTime2.GetValue = "" Then gdaFinTime2.SetValue gdaFinTime1.GetValue
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "gdaFinTime2_GotFocus")
End Sub

Private Sub gdaFinTime2_Validate(Cancel As Boolean)
  On Error Resume Next
    Cancel = ChkDate2(gdaFinTime1, gdaFinTime2)
End Sub

'Private Sub gilCompCode_Change() '公司別改變時服務類別隨之改變
'  On Error GoTo ChkErr
'    Call SelectServType(gilCompCode.GetCodeNo, gimServiceType)
'   gimServiceType.Left = 1
'  Exit Sub
'ChkErr:
'  Call ErrSub(Me.Name, "gilCompCode_Change")
'End Sub

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



Private Sub gimServiceType_Change()
  On Error GoTo ChkErr
    GimInstCode.Filter = " WHERE SERVICETYPE " & gimServiceType.GetQryStr
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "gilServiceType_Change")
End Sub

Private Function subCreateView() As Boolean
  Dim strView As String
  On Error Resume Next
    gcnGi.Execute "Drop View " & strViewName
    On Error GoTo ChkErr
      strViewName = GetTmpViewName
        strView = "Select So004.CustId,so004.DeclarantName,so004.ContTel,so004.FaciName,so004.FaciSNo,so007.Address,So004.ServiceType" & _
                  ",so004.DomicileAddress,so004.InstDate,so004.ID,so004.AgentID,So004.Boss,so004.BossID" & _
                  ",so004.SNo,so004.IDKindName1,so004.IDKindName,so004.ID2,so004.Birthday" & _
                  ",so004.AgentName,so004.IDKindName2,so004.AGENTTeL,so004.AgentBirthDay,so004.DomicileAddress2" & _
                  " from so004,so007 where so004.sno(+)=so007.sno and " & strChoose & ""

  
      strView = "Create View " & strViewName & " As (" & strView & ")"

      gcnGi.Execute strView
      SendSQL strView, True
      subCreateView = True
  Exit Function
ChkErr:
  Call ErrSub(Me.Name, "subCreateView")
End Function
