VERSION 5.00
Object = "{A0C51418-555F-11D3-BC10-0050BA01A3A9}#9.8#0"; "GiMulti.ocx"
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{074771B7-6ADE-11D5-B178-00485456B028}#2.0#0"; "csMulti.ocx"
Begin VB.Form frmSO5550A 
   BorderStyle     =   1  '單線固定
   Caption         =   "新裝機客戶統計表 [SO5550A]"
   ClientHeight    =   5745
   ClientLeft      =   1950
   ClientTop       =   1530
   ClientWidth     =   8190
   BeginProperty Font 
      Name            =   "新細明體"
      Size            =   9.75
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SO5550A.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   8190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '螢幕中央
   Begin CS_Multi.CSmulti gimStrtCode 
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   2340
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   661
      ButtonCaption   =   "街  道  範  圍"
   End
   Begin VB.CommandButton cmdExcel 
      Caption         =   "匯成Excel"
      Height          =   525
      Left            =   3660
      TabIndex        =   21
      Top             =   5100
      Width           =   1245
   End
   Begin MSMask.MaskEdBox mskCircuitNo 
      Height          =   315
      Left            =   990
      TabIndex        =   2
      Top             =   630
      Width           =   2445
      _ExtentX        =   4313
      _ExtentY        =   556
      _Version        =   393216
      ForeColor       =   12582912
      PromptChar      =   "_"
   End
   Begin VB.Frame fraCustStatus 
      Caption         =   "統計對象"
      ForeColor       =   &H00C00000&
      Height          =   645
      Left            =   30
      TabIndex        =   28
      Top             =   4020
      Width           =   3585
      Begin VB.OptionButton optAll 
         Caption         =   "全部"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   330
         TabIndex        =   11
         Top             =   270
         Value           =   -1  'True
         Width           =   765
      End
      Begin VB.OptionButton optMdu 
         Caption         =   "大樓戶"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   1230
         TabIndex        =   12
         Top             =   270
         Width           =   885
      End
      Begin VB.OptionButton optNormal 
         Caption         =   "非大樓戶"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   2250
         TabIndex        =   13
         Top             =   270
         Width           =   1125
      End
   End
   Begin VB.Frame fraPage 
      Caption         =   "分頁方式"
      ForeColor       =   &H00FF0000&
      Height          =   645
      Left            =   3615
      TabIndex        =   27
      Top             =   4020
      Width           =   4425
      Begin VB.OptionButton optIntro 
         Caption         =   "介紹人"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   2970
         TabIndex        =   18
         Top             =   285
         Width           =   915
      End
      Begin VB.OptionButton optClctAreaCode 
         Caption         =   "收費區"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   2070
         TabIndex        =   17
         Top             =   270
         Width           =   915
      End
      Begin VB.OptionButton optServCode 
         Caption         =   "服務區"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   285
         Width           =   915
      End
      Begin VB.OptionButton optAreaCode 
         Caption         =   "行政區"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1155
         TabIndex        =   16
         Top             =   270
         Width           =   915
      End
      Begin VB.OptionButton optNothing 
         Caption         =   "無"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   3915
         TabIndex        =   14
         Top             =   285
         Value           =   -1  'True
         Width           =   435
      End
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "F5:列印"
      Default         =   -1  'True
      Height          =   525
      Left            =   120
      TabIndex        =   19
      Top             =   5100
      Width           =   1245
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      Height          =   525
      Left            =   6840
      TabIndex        =   22
      Top             =   5100
      Width           =   1245
   End
   Begin VB.CommandButton cmdPrevRpt 
      Caption         =   "上次統計結果"
      Height          =   525
      Left            =   1800
      TabIndex        =   20
      Top             =   5100
      Width           =   1395
   End
   Begin Gi_Date.GiDate gdaInstTime2 
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
   Begin Gi_Date.GiDate gdaInstTime1 
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
   Begin prjGiList.GiList gilCompCode 
      Height          =   315
      Left            =   5100
      TabIndex        =   3
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
      F5Corresponding =   -1  'True
   End
   Begin Gi_Multi.GiMulti gimAreaCode 
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   1950
      Width           =   8235
      _ExtentX        =   14526
      _ExtentY        =   661
      ButtonCaption   =   "行    政    區"
   End
   Begin Gi_Multi.GiMulti gimServCode 
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   1560
      Width           =   8235
      _ExtentX        =   14526
      _ExtentY        =   661
      ButtonCaption   =   "服    務    區"
   End
   Begin prjGiList.GiList gilServiceType 
      Height          =   315
      Left            =   5100
      TabIndex        =   4
      Top             =   600
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
      Height          =   405
      Left            =   0
      TabIndex        =   10
      Top             =   3540
      Width           =   8235
      _ExtentX        =   14526
      _ExtentY        =   714
      ButtonCaption   =   "排  序  方  式"
      DataType        =   2
      ColumnOrder     =   -1  'True
   End
   Begin Gi_Multi.GiMulti gimNodeNo 
      Height          =   375
      Left            =   0
      TabIndex        =   30
      Top             =   1170
      Width           =   8235
      _ExtentX        =   14526
      _ExtentY        =   661
      ButtonCaption   =   "NodeNo"
      OneColumn       =   -1  'True
   End
   Begin CS_Multi.CSmulti gimMduId 
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   3150
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   661
      ButtonCaption   =   "大  樓  編  號"
   End
   Begin CS_Multi.CSmulti gimClassCode1 
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   2730
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   661
      ButtonCaption   =   "客  戶  類  別"
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BackStyle       =   0  '透明
      Caption         =   "服務類別"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   4200
      TabIndex        =   29
      Top             =   660
      Width           =   780
   End
   Begin VB.Label lblCircuitNo 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "網路編號"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   60
      TabIndex        =   26
      Top             =   690
      Width           =   780
   End
   Begin VB.Label lblCompCode 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "公  司  別"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   4200
      TabIndex        =   25
      Top             =   210
      Width           =   765
   End
   Begin VB.Label lblInstTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "裝機日期"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   60
      TabIndex        =   24
      Top             =   210
      Width           =   780
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "~"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   2190
      TabIndex        =   23
      Top             =   210
      Width           =   165
   End
End
Attribute VB_Name = "frmSO5550A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Table:SO001,SO002,SO033,SO014
Option Explicit
Dim strCodeNo As String
Dim strChoose3 As String
Dim blnExcel As Boolean
Dim strOrder As String

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
    Call PreviousRpt(GetPrinterName(5), RptName("SO5550"), "新裝機客戶統計表 [SO5550A]")
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
    On Error Resume Next
    gcnGi.Execute "Drop View " & strViewName
    Call ErrSub(Me.Name, "cmdPrint_Click")
End Sub

Private Sub subPrint()
  On Error GoTo ChkErr
  Dim rsTmp As New ADODB.Recordset
  Dim strSubQry(3) As String
  Dim intFor As Integer
'  Dim strSubQry1 As String
'  Dim strSubQry2 As String
'  Dim strSubQry3 As String
'  Dim strSubQry4 As String

    strChoose = " From SO001,SO002,SO003,SO014 " & strChoose
    Set rsTmp = gcnGi.Execute("SELECT Count(Distinct SO001.CUSTID) as CountCust " & strChoose)
    
    If rsTmp("CountCust") = 0 Then
        MsgNoRcd
        SendSQL , , True
    Else
        strSQL = "SELECT Distinct SO001.CUSTID As CustId,SO001.CUSTNAME,SO001.TEL1,SO001.TEL3,SO002.InstTime,SO001.CLASSNAME1," & _
                 "SO002.MEDIANAME,SO002.INTRONAME,SO002.CMNAME,SO003.CITEMNAME,SO003.Amount," & _
                 "SO003.PERIOD,SO003.ClctDate,SO003.STARTDATE,'~' as Range,SO003.STOPDATE,SO001.INSTADDRESS," & _
                 "SO014.AreaCode,SO014.AreaName,SO001.ServCode,SO001.ServArea,SO002.IntroId,SO001.ClassCode1," & _
                 "SO001.InstAddrNo,SO014.MDUId,SO014.AddrSort,SO001.ClctAreaCode,SO001.ClctAreaName " & strChoose
                 
        strSubQry(0) = "SELECT SO002.MediaCode,SO002.MediaName,COUNT(DISTINCT SO001.CUSTID) COUNTCUSTID," & rsTmp("CountCust") & " as SumCust " & strChoose & " GROUP BY SO002.MediaCode,SO002.MediaName"
        strSubQry(1) = "SELECT SO014.NodeNo,COUNT(DISTINCT SO001.CUSTID) COUNTCUSTID," & rsTmp("CountCust") & " as SumCust " & strChoose & " GROUP BY SO014.NodeNo"
        strSubQry(2) = "SELECT SO001.ClassCode1,SO001.ClassName1,COUNT(DISTINCT SO001.CUSTID) COUNTCUSTID," & rsTmp("CountCust") & " as SumCust " & strChoose & " GROUP BY SO001.ClassCode1,SO001.ClassName1"
        strSubQry(3) = "SELECT SO014.AreaCode,SO014.AreaName,COUNT(DISTINCT SO001.CUSTID) COUNTCUSTID," & rsTmp("CountCust") & " as SumCust " & strChoose & " GROUP BY SO014.AreaCode,SO014.AreaName"
              
        Call subCreateView(strSQL)
        If CreateSubView2(strSubQry()) = False Then Exit Sub
        If blnExcel Then
            strSQL = "SELECT CustId,CUSTNAME,TEL1,TEL3,InstTime,CLASSNAME1," & _
                     "MEDIANAME,INTRONAME,CMNAME,CITEMNAME,Amount," & _
                     "PERIOD,ClctDate,STARTDATE,Range,STOPDATE,INSTADDRESS From " & _
                      strViewName & " V"
            strSubQry(0) = "SELECT NodeNo,COUNTCUSTID From " & strSubViewName(1) & " V"
            strSubQry(1) = "SELECT ClassName1,COUNTCUSTID From " & strSubViewName(2) & " V"
            strSubQry(2) = "SELECT MediaName,COUNTCUSTID From " & strSubViewName(0) & " V"
            strSubQry(3) = "SELECT AreaName,COUNTCUSTID From " & strSubViewName(3) & " V"
                              
            Call toExcel(strSQL, strSubQry())
        Else
            strSQL = "Select * From " & strViewName & " V"
            For intFor = 0 To UBound(strSubQry)
                strSubQry(intFor) = "Select * From " & strSubViewName(intFor) & " V"
            Next
            Call PrintRpt(GetPrinterName(5), RptName("SO5550"), , "新裝機客戶統計表 [SO5550A]", strSQL, strChooseString, , True, , , strGroupName, GiPaperLandscape, , strSubQry(0), strSubQry(1), strSubQry(2), strSubQry(3))
        End If
    End If
    CloseRecordset rsTmp
    blnExcel = False
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "subPrint")
End Sub

Private Sub toExcel(ByVal strSQL As String, strsubsql() As String)
  On Error GoTo ChkErr
    Dim rsExcel As New ADODB.Recordset
    Dim rsSubExcel(3) As New ADODB.Recordset
    Dim intFor As Integer
    '問題集2881 修改時,一併更正,如果排序沒有選擇的話,內定為CustID 2006/12/05
      strSQL = strSQL & IIf(strOrder = "", " Order by V.CustId ", " Order by " & Mid(strOrder, 2))
      If Not GetRS(rsExcel, strSQL) Then Exit Sub
      For intFor = 0 To UBound(strsubsql)
          If Not GetRS(rsSubExcel(intFor), strsubsql(intFor)) Then Exit Sub
      Next
       '#7074 To Add a choice-condition parameter into excel print function by Kin 2015/09/09
      Call UseProperty(rsExcel, "新裝機客戶統計表", "第一頁", "客戶編號,客戶名稱,電話(1),電話(3),裝機時間,客戶類別,介紹媒介,介紹人,收費方式,收費項目,金額,期數,收費日期,有效期限,,,裝機地址", _
                      , rsSubExcel(0), "NodeNo,客戶數", rsSubExcel(1), "客戶類別,客戶數", rsSubExcel(2), "介紹媒介,客戶數", rsSubExcel(3), "行政區,客戶數", _
                      , , , , , , , , , , , , , , , , , , , , , , , , , , strChooseString)
      CloseRecordset rsExcel
      For intFor = 0 To UBound(strsubsql)
          CloseRecordset rsSubExcel(intFor)
      Next
      blnExcel = False
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "subExcel")
End Sub

Private Function IsDataOk() As Boolean
  On Error GoTo ChkErr
    If gilCompCode.GetCodeNo = "" Then gilCompCode.SetFocus: strErrFile = "公司別": GoTo 66
    If gilServiceType.GetCodeNo = "" Then gilServiceType.SetFocus: strErrFile = "服務類別": GoTo 66
    If gdaInstTime1.GetValue = "" Then gdaInstTime1.SetFocus: strErrFile = "裝機日期起始日": GoTo 66
    If gdaInstTime2.GetValue = "" Then gdaInstTime2.SetFocus: strErrFile = "裝機日期截止日": GoTo 66
    
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
    Dim strPage As String
    Dim strCustStatus  As String
    Dim strGetGroupName As String
    Dim arrOrder() As String
    Dim varOrder As Variant
    Dim intSort As Integer
    
'    strChoose = " Where SO001.CustId=SO003.CustId(+) And SO001.Custid=SO002.Custid AND SO002.ServiceType=SO003.ServiceType And SO001.InstAddrNo=SO014.AddrNo"
    strChoose = " Where SO001.Custid = SO002.Custid AND SO001.CompCode=SO002.CompCode AND SO003.CustId=SO002.Custid AND SO002.ServiceType=SO003.ServiceType AND SO014.AddrNo=SO001.InstAddrNo "

    strChooseString = ""
    strGetGroupName = ""
    strOrder = ""
    
  '日期
    If gdaInstTime1.GetValue <> "" Then Call subAnd("(SO002.InstTime >= To_Date('" & gdaInstTime1.GetValue & "','YYYYMMDD')")
    If gdaInstTime2.GetValue <> "" Then Call subAnd("SO002.InstTime < To_Date('" & gdaInstTime2.GetValue & "','YYYYMMDD')+1")
    
  'GiList
    
    'If gilServiceType.GetCodeNo <> "" Then Call subAnd("Instr(Nvl(SO001.ServiceType,'" & gilServiceType.GetCodeNo & "'),'" & gilServiceType.GetCodeNo & "') > 0")
'    If gilServiceType.GetCodeNo <> "" Then Call subAnd("instr(SO001.ServiceType,'" & gilServiceType.GetCodeNo & "')>0")
    If gilServiceType.GetCodeNo <> "" Then Call subAnd("SO002.ServiceType='" & gilServiceType.GetCodeNo & "')")
    If gilCompCode.GetCodeNo <> "" Then Call subAnd("SO001.CompCode=" & gilCompCode.GetCodeNo)
  '統計對象
    Select Case True
           Case optMdu.Value
                If gimMduId.GetQryStr <> "" Then
                   subAnd "SO001.MduId " & gimMduId.GetQryStr
                Else
                   subAnd "SO001.MduId Is Not Null"
                End If
                strCustStatus = "大樓戶"
           Case optNormal.Value
                subAnd "SO001.MduId Is Null"
                strCustStatus = "非大樓戶"
           Case Else
                If gimMduId.GetQryStr <> "" Then
                   subAnd "(SO001.MduId Is Null or SO001.Mduid " & gimMduId.GetQryStr & ")"
                End If
                strCustStatus = "全部"
    End Select
'    strChoose = " Where SO001.CustId=SO003.CustId(+) And SO001.InstAddrNo=SO014.AddrNo" & _
'                 IIf(strChoose = "", "", " And ") & strChoose
  
  '網路編號
    If mskCircuitNo.Text <> "" Then Call subAnd("SO014.CircuitNo='" & mskCircuitNo.Text & "'")
    'Node No
    If gimNodeNo.GetQryStr <> "" Then Call subAnd("SO014.NodeNo " & gimNodeNo.GetQryStr)
    
  'GiMulti
    If gimServCode.GetQryStr <> "" Then Call subAnd("SO001.ServCode " & gimServCode.GetQryStr)
    If gimAreaCode.GetQryStr <> "" Then Call subAnd("SO014.AreaCode " & gimAreaCode.GetQryStr)
    If gimClassCode1.GetQryStr <> "" Then Call subAnd("SO001.ClassCode1 " & gimClassCode1.GetQryStr)
    If gimStrtCode.GetQryStr <> "" Then Call subAnd("SO014.StrtCode " & gimStrtCode.GetQryStr)
    
  '分頁方式
    Select Case True
           Case optAreaCode.Value
                strGroupName = "GroupName={V.AreaCode};GroupName1={V.AreaName};ReportType=True"
                strPage = "行政區"
                strOrder = strOrder & ",V.AreaCode"
           Case optServCode.Value
                strGroupName = "GroupName={V.ServCode};GroupName1={V.ServArea};ReportType=True"
                strPage = "服務區"
                strOrder = strOrder & ",V.ServCode"
           Case optClctAreaCode.Value
                strGroupName = "GroupName={V.ClctAreaCode};GroupName1={V.ClctAreaName};ReportType=True"
                strPage = "收費區"
                strOrder = strOrder & ",V.ClctAreaCode"
                'INTRONAME
           Case optIntro.Value
                strGroupName = "GroupName={V.INTRONAME};GroupName1={V.INTRONAME};ReportType=True"
                strPage = "介紹人"
                strOrder = strOrder & ",V.INTRONAME"
                
           Case optNothing.Value
                '問題集2881 分頁方式選擇無時，排序都會以客編做排序，將下列Mark
                'strOrder = strOrder & ",V.Custid"
                Select Case Left(gimOrder.GetColumnOrderCode, 1)
                  Case "A"
                    strGroupName = "ReportType=False;GroupName=date({V.InstTime});GroupName1={V.CustName}"
                  Case "B"
                    strGroupName = "ReportType=False;GroupName={V.IntroId};GroupName1={V.CustName}"
                  Case "C"
                    strGroupName = "ReportType=False;GroupName={V.ClassCode1};GroupName1={V.CustName}"
                  Case "D"
                    strGroupName = "ReportType=False;GroupName={V.AddrSort};GroupName1={V.CustName}"
                  Case "E"
                    strGroupName = "ReportType=False;GroupName={V.MDUId};GroupName1={V.CustName}"
                  Case "F"
                    strGroupName = "ReportType=False;GroupName={V.ServCode};GroupName1={V.CustName}"
                  Case "G"
                    strGroupName = "ReportType=False;GroupName={V.AreaCode};GroupName1={V.CustName}"
                  Case "H"
                    strGroupName = "ReportType=False;GroupName={V.CustId};GroupName1={V.CustName}"
                  Case Else
                    strGroupName = "ReportType=False;GroupName={V.CustId};GroupName1={V.CustName}"
                End Select
    End Select
    
  '排序方式
    If gimOrder.GetColumnOrderCode <> "" Then
      arrOrder = Split(gimOrder.GetColumnOrderCode, ",")
      intSort = 0
      For Each varOrder In arrOrder
        Select Case arrOrder(intSort)
          Case "A"
                strGroupName = strGroupName & ";Sort" & intSort & "=date({V.InstTime})"
                strOrder = strOrder & ",V.InstTime"
          Case "B"
                strGroupName = strGroupName & ";Sort" & intSort & "={V.IntroId}"
                strOrder = strOrder & ",V.IntroId"
          Case "C"
                strGroupName = strGroupName & ";Sort" & intSort & "={V.ClassCode1}"
                strOrder = strOrder & ",V.ClassCode1"
          Case "D"
                strGroupName = strGroupName & ";Sort" & intSort & "={V.AddrSort}"
                strOrder = strOrder & ",V.AddrSort"
          Case "E"
                strGroupName = strGroupName & ";Sort" & intSort & "={V.MDUId}"
                strOrder = strOrder & ",V.MDUId"
          Case "F"
                strGroupName = strGroupName & ";Sort" & intSort & "={V.ServCode}"
                strOrder = strOrder & ",V.ServCode"
          Case "G"
                strGroupName = strGroupName & ";Sort" & intSort & "={V.AreaCode}"
                strOrder = strOrder & ",V.AreaCode"
          Case "H"
                '問題集2881 修改時，發現排序錯誤　原本為  strOrder = strOrder & ",So001.CustId" 2006/12/05
                strGroupName = strGroupName & ";Sort" & intSort & "={V.CustId}"
                strOrder = strOrder & ",V.CustId"
        End Select
        intSort = intSort + 1
      Next
    End If

   strChooseString = "裝機日期:" & subSpace(gdaInstTime1.GetValue(True)) & "~" & subSpace(gdaInstTime2.GetValue(True)) & ";" & _
                     "公司別　:" & subSpace(gilCompCode.GetDescription) & ";" & _
                     "服務類別:" & subSpace(gilServiceType.GetDescription) & ";" & _
                     "網路編號:" & subSpace(mskCircuitNo.Text) & ";" & _
                     "NodeNo  :" & subSpace(gimNodeNo.GetDispStr) & ";" & _
                     "服務區　:" & subSpace(gimServCode.GetDispStr) & ";" & _
                     "行政區　:" & subSpace(gimAreaCode.GetDispStr) & ";" & _
                     "街道範圍:" & subSpace(gimStrtCode.GetDispStr) & ";" & _
                     "客戶類別:" & subSpace(gimClassCode1.GetDispStr) & ";" & _
                     "大樓名稱:" & subSpace(gimMduId.GetDispStr) & ";" & _
                     "排序方式:" & subSpace(gimOrder.GetColumnOrderDspStr) & ";" & _
                     "統計對象:" & subSpace(strCustStatus) & ";" & _
                     "分頁方式:" & subSpace(strPage)
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "subChoose")
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error GoTo ChkErr
    Call FunctionKey(KeyCode, Shift, frmSO5550A)
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

Private Sub subGim()
  On Error GoTo ChkErr
    Call SetgiMulti(gimNodeNo, "CompCode", "CodeNo", "CD047", "CompCode", "NodeNo")
    Call SetgiMulti(gimClassCode1, "CodeNo", "Description", "CD004", "客戶類別代碼", "客戶類別名稱")
    Call SetgiMulti(gimServCode, "CodeNo", "Description", "CD002", "服務區代碼", "服務區名稱")
    Call SetgiMulti(gimStrtCode, "CodeNo", "Description", "CD017", "街道代碼", "街道名稱")
    Call SetgiMulti(gimAreaCode, "CodeNo", "Description", "CD001", "行政區代碼", "行政區名稱")
    Call SetgiMulti(gimMduId, "MduId", "Name", "SO017", "大樓代號", "大樓名稱")
    Call SetgiMultiAddItem(gimOrder, "A,B,C,D,E,F,G,H", "裝機日期,介紹人,客戶類別,地址,大樓編號,服務區,行政區,客戶編號")
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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  On Error Resume Next
  Call DropTMPVIEW(strViewName, strSubViewName)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
  ReleaseCOM frmSO5550A
End Sub

Private Sub gdaInstTime1_GotFocus()
  On Error Resume Next
  If gdaInstTime1.GetValue = "" Then gdaInstTime1.SetValue (RightDate)
End Sub

Private Sub gdaInstTime2_GotFocus()
  On Error Resume Next
  If gdaInstTime1.GetValue = "" Or gdaInstTime2.GetValue = "" Then gdaInstTime2.SetValue (gdaInstTime1.GetValue)
End Sub

Private Sub gdaInstTime2_Validate(Cancel As Boolean)
  On Error Resume Next
  Cancel = ChkDate2(gdaInstTime1, gdaInstTime2)
End Sub

Private Sub gilCompCode_Change()
On Error GoTo ChkErr
    Call SelectServType(gilCompCode.GetCodeNo, gilServiceType)
    gilServiceType.ListIndex = 1
    Call GiMultiFilter(gimNodeNo, gilServiceType.GetCodeNo, gilCompCode.GetCodeNo)
    Call GiMultiFilter(gimServCode, , gilCompCode.GetCodeNo)
    Call GiMultiFilter(gimAreaCode, , gilCompCode.GetCodeNo)
    Call GiMultiFilter(gimStrtCode, , gilCompCode.GetCodeNo)
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
  On Error Resume Next
    Call GiMultiFilter(gimNodeNo, gilServiceType.GetCodeNo, gilCompCode.GetCodeNo)
    Call GiMultiFilter(gimClassCode1, gilServiceType.GetCodeNo)
    'strCodeNo = GetCode("CD005", " in (1,5)", gilServiceType.GetCodeNo)
    
End Sub

Private Sub gimMduId_Change()
  On Error Resume Next
    If gimMduId.GetQryStr = "" Then
       optNormal.Enabled = True
    Else
       optNormal.Enabled = False
       optMdu.Value = True
    End If
End Sub

Private Function subCreateView(ByVal viewSQL As String) As Boolean
  Dim strView As String
  On Error Resume Next
  Dim strHead As String
    gcnGi.Execute "Drop View " & strViewName
  On Error GoTo ChkErr
      strViewName = GetTmpViewName
'      strHead = "Select So003.CustId,So003.CitemName,So003.Period,So003.Amount,So003.StartDate,So003.StopDate,SO003.ClctDate "
'      'strChoose3 = " Where So001.InstSNo = So033.BillNo And CitemCode in (Select CodeNo From CD019 Where RefNo =1 ) " & IIf(strChoose3 = "", "", " And ") & strChoose3
'      strChoose3 = " Where So001.InstSNo = So033.BillNo And So033.RealPeriod > 0   " & IIf(strChoose3 = "", "", " And ") & strChoose3
' '     strChoose3 = " Where SO001.CustId=SO007.CustId And So007.SNo = So033.BillNo And So033.RealPeriod > 0   " & IIf(strChoose3 = "", "", " And ") & strChoose3
''      strView = "Create View " & strViewName & _
'                " as " & strHead & " From So001 ,So033 " & strChoose3 & _
'                " Union All " & strHead & " From So001 ,So034 So033 " & strChoose3
      strView = "Create View " & strViewName & " As (" & viewSQL & ")"

      gcnGi.Execute strView
      SendSQL strView, True
  Exit Function
ChkErr:
    ErrSub Me.Name, "subCreateView"
End Function

