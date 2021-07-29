VERSION 5.00
Object = "{A0C51418-555F-11D3-BC10-0050BA01A3A9}#9.8#0"; "GiMulti.ocx"
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{074771B7-6ADE-11D5-B178-00485456B028}#2.0#0"; "csMulti.ocx"
Begin VB.Form frmSO5570A 
   BorderStyle     =   1  '單線固定
   Caption         =   "無週期性收費項目客戶報表[SO5570A]"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8265
   BeginProperty Font 
      Name            =   "新細明體"
      Size            =   9.75
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SO5570A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form53"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   8265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '螢幕中央
   Begin VB.PictureBox pic2 
      Height          =   2745
      Left            =   30
      ScaleHeight     =   2685
      ScaleWidth      =   8115
      TabIndex        =   26
      Top             =   2250
      Width           =   8175
      Begin VB.VScrollBar vsl1 
         Height          =   2625
         LargeChange     =   100
         Left            =   7800
         Max             =   100
         SmallChange     =   100
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   30
         Width           =   285
      End
      Begin VB.Frame fraMulti 
         Height          =   3375
         Left            =   0
         TabIndex        =   27
         Top             =   0
         Width           =   7815
         Begin VB.PictureBox picMsg 
            Height          =   675
            Left            =   2520
            ScaleHeight     =   615
            ScaleWidth      =   3405
            TabIndex        =   28
            Top             =   1020
            Visible         =   0   'False
            Width           =   3465
            Begin VB.Label lblMsg 
               AutoSize        =   -1  'True
               Caption         =   "正在刪除暫存資料,請稍候!!"
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
               Left            =   390
               TabIndex        =   29
               Top             =   210
               Width           =   2685
            End
         End
         Begin CS_Multi.CSmulti gimStrtCode 
            Height          =   405
            Left            =   0
            TabIndex        =   30
            Top             =   2670
            Width           =   7815
            _ExtentX        =   13785
            _ExtentY        =   714
            ButtonCaption   =   "街  道  範  圍"
         End
         Begin Gi_Multi.GiMulti gimStatusCode 
            Height          =   375
            Left            =   0
            TabIndex        =   35
            Top             =   975
            Width           =   7815
            _ExtentX        =   13785
            _ExtentY        =   661
            ButtonCaption   =   "客  戶  狀  態"
         End
         Begin Gi_Multi.GiMulti gimServCode 
            Height          =   375
            Left            =   0
            TabIndex        =   39
            Top             =   1995
            Width           =   7815
            _ExtentX        =   13785
            _ExtentY        =   661
            ButtonCaption   =   "服     務     區"
         End
         Begin Gi_Multi.GiMulti gimAreaCode 
            Height          =   375
            Left            =   0
            TabIndex        =   40
            Top             =   2325
            Width           =   7815
            _ExtentX        =   13785
            _ExtentY        =   661
            ButtonCaption   =   "行     政     區"
         End
         Begin Gi_Multi.GiMulti gimOrder 
            Height          =   375
            Left            =   0
            TabIndex        =   31
            Top             =   3015
            Width           =   7815
            _ExtentX        =   13785
            _ExtentY        =   661
            ButtonCaption   =   "排  序  方  式"
            DataType        =   2
            ColumnOrder     =   -1  'True
         End
         Begin CS_Multi.CSmulti gimCitemCode 
            Height          =   375
            Left            =   0
            TabIndex        =   38
            Top             =   1650
            Width           =   7815
            _ExtentX        =   13785
            _ExtentY        =   661
            ButtonCaption   =   "收  費  項  目"
         End
         Begin CS_Multi.CSmulti gimClassCode 
            Height          =   375
            Left            =   0
            TabIndex        =   32
            Top             =   0
            Width           =   7815
            _ExtentX        =   13785
            _ExtentY        =   661
            ButtonCaption   =   "客  戶  類  別"
         End
         Begin Gi_Multi.GiMulti gimWipCode3 
            Height          =   375
            Left            =   0
            TabIndex        =   37
            Top             =   1305
            Width           =   7815
            _ExtentX        =   13785
            _ExtentY        =   661
            ButtonCaption   =   "非停拆移機中"
         End
         Begin CS_Multi.CSmulti gimClassCode2 
            Height          =   375
            Left            =   0
            TabIndex        =   33
            Top             =   330
            Width           =   7815
            _ExtentX        =   13785
            _ExtentY        =   661
            ButtonCaption   =   "客  戶  類  別  二"
         End
         Begin CS_Multi.CSmulti gimClassCode3 
            Height          =   375
            Left            =   0
            TabIndex        =   34
            Top             =   660
            Width           =   7815
            _ExtentX        =   13785
            _ExtentY        =   661
            ButtonCaption   =   "客  戶  類  別  三"
         End
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "網路編號"
      Height          =   735
      Left            =   30
      TabIndex        =   25
      Top             =   1320
      Width           =   8175
      Begin VB.TextBox mskCircuitNo 
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   90
         TabIndex        =   8
         Top             =   240
         Width           =   4605
      End
      Begin VB.OptionButton optAll 
         Caption         =   "全部"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   6720
         TabIndex        =   10
         Top             =   270
         Value           =   -1  'True
         Width           =   705
      End
      Begin VB.OptionButton optCircuitNo 
         Caption         =   "只印空白網路編號"
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   4830
         TabIndex        =   9
         Top             =   240
         Width           =   1875
      End
   End
   Begin VB.CommandButton cmdExcel 
      Caption         =   "匯成Excel"
      Height          =   525
      Left            =   3660
      TabIndex        =   16
      Top             =   5970
      Width           =   1245
   End
   Begin VB.CommandButton cmdPrevRpt 
      Caption         =   "上次統計結果"
      Height          =   525
      Left            =   1815
      TabIndex        =   15
      Top             =   5970
      Width           =   1395
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      Height          =   525
      Left            =   6120
      TabIndex        =   17
      Top             =   5970
      Width           =   1245
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "F5:列印"
      Default         =   -1  'True
      Height          =   525
      Left            =   120
      TabIndex        =   14
      Top             =   5970
      Width           =   1245
   End
   Begin VB.Frame Frame2 
      Caption         =   "分頁方式"
      ForeColor       =   &H00FF0000&
      Height          =   630
      Left            =   30
      TabIndex        =   22
      Top             =   570
      Width           =   3765
      Begin VB.OptionButton optAreaCode 
         Caption         =   "行政區"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1785
         TabIndex        =   3
         Top             =   270
         Width           =   1095
      End
      Begin VB.OptionButton optServCode 
         Caption         =   "服務區"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   330
         TabIndex        =   2
         Top             =   270
         Value           =   -1  'True
         Width           =   1125
      End
      Begin VB.OptionButton optNothing 
         Caption         =   "無"
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   3105
         TabIndex        =   4
         Top             =   240
         Width           =   555
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "無週期性定義"
      ForeColor       =   &H00FF0000&
      Height          =   630
      Left            =   0
      TabIndex        =   18
      Top             =   5115
      Width           =   8205
      Begin VB.OptionButton OptCitemNo 
         Caption         =   "完全無週期收費項目者"
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   5190
         TabIndex        =   13
         Top             =   270
         Width           =   2295
      End
      Begin VB.OptionButton OptCitemAll 
         Caption         =   "全部"
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   270
         TabIndex        =   11
         Top             =   270
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton OptCitemCode 
         Caption         =   "有週期收費項目,但金額為0者"
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   1740
         TabIndex        =   12
         Top             =   270
         Width           =   2865
      End
   End
   Begin Gi_Date.GiDate gdaInstDate2 
      Height          =   345
      Left            =   2640
      TabIndex        =   1
      Top             =   75
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
   Begin Gi_Date.GiDate gdaInstDate1 
      Height          =   345
      Left            =   1290
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
   Begin prjGiList.GiList gilClctEn 
      Height          =   315
      Left            =   4950
      TabIndex        =   7
      Top             =   930
      Width           =   2730
      _ExtentX        =   4815
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
      FldWidth1       =   800
      FldWidth2       =   1600
      F5Corresponding =   -1  'True
   End
   Begin prjGiList.GiList gilServiceType 
      Height          =   315
      Left            =   4950
      TabIndex        =   6
      Top             =   510
      Width           =   2730
      _ExtentX        =   4815
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
      FldWidth1       =   800
      FldWidth2       =   1600
      F5Corresponding =   -1  'True
   End
   Begin prjGiList.GiList gilCompCode 
      Height          =   315
      Left            =   4950
      TabIndex        =   5
      Top             =   90
      Width           =   2730
      _ExtentX        =   4815
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
      FldWidth1       =   800
      FldWidth2       =   1600
      F5Corresponding =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BackStyle       =   0  '透明
      Caption         =   "公司別"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   4110
      TabIndex        =   24
      Top             =   150
      Width           =   585
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BackStyle       =   0  '透明
      Caption         =   "服務類別"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   4110
      TabIndex        =   23
      Top             =   570
      Width           =   780
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "收費人員"
      ForeColor       =   &H00808000&
      Height          =   195
      Left            =   4110
      TabIndex        =   21
      Top             =   1005
      Width           =   780
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "---"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   2430
      TabIndex        =   20
      Top             =   165
      Width           =   180
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "裝機日期範圍"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   90
      TabIndex        =   19
      Top             =   165
      Width           =   1170
   End
End
Attribute VB_Name = "frmSO5570A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Table:SO014,SO001,SO002,SO075
Option Explicit
Dim lngLastBatchNo As Long
Dim blnExcel As Boolean
Dim strField As String

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
    Call PreviousRpt(GetPrinterName(5), RptName("SO5570"), "無週期性收費項目客戶報表 [SO5570A]")
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
        gcnGi.Execute "Delete From So075 Where BatchNo = " & lngLastBatchNo
          Call subChoose
          Call subPrint
        picMsg.Visible = True
        DoEvents
        gcnGi.Execute "Delete From So075 Where BatchNo = " & lngLastBatchNo
        picMsg.Visible = False
      Me.Enabled = True
    Screen.MousePointer = vbDefault
  
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "cmdPrint_Click")
End Sub

Private Sub subPrint()
  On Error GoTo ChkErr
  Dim rsTmp As New ADODB.Recordset
  Dim strSubQry(1) As String
  Dim intFor As Integer
'  Dim strSubQry1 As String
'  Dim strSubQry2 As String
    Dim S As String
    S = "SELECT Count(Distinct SO001.CUSTID) as CountCust " & strChoose
    Set rsTmp = gcnGi.Execute("SELECT Count(Distinct SO001.CUSTID) as CountCust " & strChoose)
    If rsTmp("CountCust") = 0 Then
       MsgNoRcd
       SendSQL , , True
    Else
        strsql = "Select SO001.CUSTID,SO001.CUSTNAME, SO001.TEL1,SO001.CHARGEADDRESS," & _
                 "SO001.ClASSNAME2,SO001.CLASSNAME3,SO001.CLASSCODE1,SO001.CLASSCODE2,SO001.CLASSCODE3," & _
                 "SO002.CUSTSTATUSNAME, SO001.CLASSNAME1,SO001.VIEWLEVEL,SO002.INSTCOUNT," & _
                 "SO002.INSTTIME, " & strField & ", SO014.CLCTNAME,SO014.ServCode,SO014.AreaCode,SO014.AddrSort "
               
        strSubQry(0) = "SELECT SO001.SERVCODE,SO001.SERVAREA SERVAREA,Count(Distinct SO001.CUSTID) CountCustid," & rsTmp("CountCust") & " as SumCust " & strChoose & " Group By SO001.SERVCODE,SO001.SERVAREA "
        strSubQry(1) = "SELECT SO014.AREACODE,SO014.AREANAME AREANAME,Count(Distinct SO001.CUSTID) CountCustid," & rsTmp("CountCust") & " as SumCust " & strChoose & " Group By SO014.AREACODE,SO014.AreaName "
       
        SendSQL strsql
        SendSQL strSubQry(0)
        SendSQL strSubQry(1), True
       
        If CreateSubView2(strSubQry()) = False Then Exit Sub
        For intFor = 0 To UBound(strSubQry)
            strSubQry(intFor) = "Select * From " & strSubViewName(intFor) & " V"
        Next
       If blnExcel Then
          strsql = strsql & strChoose
          Call toExcel(strsql, strSubQry())
       Else
          strsql = strsql & ",SO001.ServCode " & strChoose
          Call PrintRpt2(GetPrinterName(5), RptName("SO5570"), , "無週期性收費項目客戶報表 [SO5570A]", strsql, strChooseString, , True, , , strGroupName, GiPaperLandscape, , strSubQry(0), strSubQry(1))
       End If
    End If
    CloseRecordset rsTmp
    blnExcel = False
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "subPrint")
End Sub

Private Sub toExcel(ByVal strsql As String, strSubQry() As String)
  On Error GoTo ChkErr
    Dim rsExcel As New ADODB.Recordset
    Dim rsSubExcel(1) As New ADODB.Recordset
    Dim intFor As Integer
    '#3705 將匯出Excel的報表獨立出來 By Kin 2008/02/15
    RptToTxt RptName("SO5570", "E"), , strsql, , , , strGroupName, , , Environ("Temp") & "\SO5570", , , strSubQry(0), strSubQry(1)
    If Not Get_RS_From_Txt(Environ("Temp"), "SO5570.txt", rsExcel) Then blnExcel = False: Exit Sub
    For intFor = 0 To UBound(strSubQry)
        If Not GetRS(rsSubExcel(intFor), strSubQry(intFor)) Then Exit Sub
    Next
    Call UseProperty(rsExcel, "無週期性收費項目客戶報表", "第一頁", , , rsSubExcel(0), "服務區,統計客戶數", rsSubExcel(1), "行政區,統計客戶數")
    blnExcel = False
    CloseRecordset rsExcel
    For intFor = 0 To UBound(strSubQry)
        CloseRecordset rsSubExcel(intFor)
    Next
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "toExcel")
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
    Dim strpagetype As String
    Dim strChoose1 As String
    Dim strChoose3 As String
    Dim strChoose14 As String
    Dim StrTableName As String
    Dim rsCustid As New ADODB.Recordset '先取出有週期性收費客戶代碼的Recordset
    Dim strCustId As String '記錄取出有週期性收費客戶代碼
    Dim strGetGroupName As String
    Dim varSplit As Variant
    Dim varCollect As Variant
    Dim lngBatchNo As Long
    Dim lngOldBatchNo As Long
    Dim strChoose31 As String
    Dim strCitem As String
    Dim arrOrder() As String
    Dim varOrder As Variant
    Dim intSort As Integer
    Dim strSQLSO075 As String
    
    strChoose3 = ""
    strChoose = ""
    StrTableName = ""
    strChooseString = ""
    strGroupName = ""
    strGetGroupName = ""
    strCitem = ""
    strField = ""
    strChoose31 = ""
    
  '日期
    If gdaInstDate1.GetValue <> "" Then Call subAnd("SO002.InstTime >= To_Date('" & gdaInstDate1.GetValue & "','YYYYMMDD')")
    If gdaInstDate2.GetValue <> "" Then Call subAnd("SO002.InstTime < To_Date('" & gdaInstDate2.GetValue & "','YYYYMMDD')+1")

  'GiList
    'If gilServiceType.GetCodeNo <> "" Then Call subAnd("Instr(Nvl(SO001.ServiceType,'" & gilServiceType.GetCodeNo & "'),'" & gilServiceType.GetCodeNo & "') > 0")
    'If gilServiceType.GetCodeNo <> "" Then Call subAnd("SO001.ServiceType='" & gilServiceType.GetCodeNo & "'")
    If gilServiceType.GetCodeNo <> "" Then Call subAnd("instr(SO001.ServiceType,'" & gilServiceType.GetCodeNo & "')>0")
    If gilServiceType.GetCodeNo <> "" Then Call subAnd("SO002.ServiceType='" & gilServiceType.GetCodeNo & "'")

  'GiMulti
    If gilCompCode.GetCodeNo <> "" Then Call subAnd("SO001.CompCode = " & gilCompCode.GetCodeNo)
    'If gimCompCode.GetQryStr <> "" Then Call subAnd("SO001.CompCode " & gimCompCode.GetQryStr)
    If gimClassCode.GetQryStr <> "" Then Call subAnd("SO001.ClassCode1 " & gimClassCode.GetQryStr)
    
    '**************************************************************************************************
    '#3705 增加客戶類別2與3 By Kin 2008/02/14
    If gimClassCode2.GetQryStr <> "" Then Call subAnd("SO001.ClassCode2 " & gimClassCode2.GetQryStr)
    If gimClassCode3.GetQryStr <> "" Then Call subAnd("SO001.ClassCode3 " & gimClassCode3.GetQryStr)
    '**************************************************************************************************
    
    If gimStatusCode.GetQryStr <> "" Then Call subAnd("SO002.CustStatusCode " & gimStatusCode.GetQryStr)
    If gimWipCode3.GetQryStr <> "" Then Call subAnd("NOT SO002.WipCode3 " & gimWipCode3.GetQryStr)      '問題集2322
    If gimServCode.GetQryStr <> "" Then Call subAnd("SO001.ServCode " & gimServCode.GetQryStr)
    
    If gilClctEn.GetCodeNo <> "" Then Call subAnd("SO014.ClctEn='" & gilClctEn.GetCodeNo & "'")
    If gimAreaCode.GetQryStr <> "" Then Call subAnd("SO014.AreaCode " & gimAreaCode.GetQryStr)
    If gimStrtCode.GetQryStr <> "" Then Call subAnd("SO014.StrtCode " & gimStrtCode.GetQryStr)
    
    '網路編號
    If mskCircuitNo.Text <> "" Then
       Call subAnd("SO014.CircuitNo='" & mskCircuitNo.Text & "'")
    Else
      If optCircuitNo.Value Then Call subAnd("SO014.CircuitNo is null ")
    End If
    
    '篩選So003 很複雜!!
    lngBatchNo = Val(GetRsValue("SELECT S_SO075_BatchNo.NextVal FROM DUAL") & "")
    If gilServiceType.GetCodeNo <> "" Then Call subAnd2(strChoose3, "SO003.ServiceType = '" & gilServiceType.GetCodeNo & "'")
    
    If gimCitemCode.GetQueryCode <> "" Then
        varCollect = Split(Replace(gimCitemCode.GetQueryCode, "'", ""), ",")
        For Each varSplit In varCollect
            '無週期性收費定義
            If Trim(strChoose3) <> "" Then Call subAnd2(strChoose3, " CitemCode = " & varSplit)
        Next
    End If
    strSQLSO075 = "Insert Into SO075 (CustId ,BatchNo,PrtSNo ) " & _
                  "Select Distinct SO003.CustId," & lngBatchNo & ",Amount " & _
                  "From SO003,SO001,SO002,SO014 " & _
                  "Where SO003.CustId = SO001.CustId AND " & _
                  "SO003.CUSTID=SO002.CUSTID AND " & _
                  "SO003.SERVICETYPE=SO002.SERVICETYPE AND " & _
                  "So014.AddrNo=So001.InstAddrNo And " & strChoose & IIf(strChoose3 = "", "", " And ") & strChoose3
    gcnGi.Execute strSQLSO075
    SendSQL strSQLSO075
    
    Select Case True
        Case OptCitemAll '全部
            Call subAnd(" (TO_Number(SO075.PrtSNo) = 0 Or So075.CustId is Null)")
            strCitem = "全部"
        Case OptCitemCode '有週期收費項目,但金額為0者
            Call subAnd(" TO_Number(SO075.PrtSNo) = 0 ")
            strCitem = "有週期收費項目,但金額為0者"
        Case OptCitemNo '完全無週期收費項目者
            Call subAnd(" So075.CustId is Null ")
            strCitem = "完全無週期收費項目者"
    End Select
    
    '刪除暫存資料用
    lngLastBatchNo = lngBatchNo
    
    strChoose = " From SO014,SO001,SO002,(Select CustId ,PrtSNo From SO075 Where BatchNo = " & lngBatchNo & ") So075 " & _
                "Where SO001.InstAddrNo = SO014.AddrNo And " & _
                "SO001.CUSTID=SO002.CUSTID AND " & _
                "SO001.CustId = SO075.CustId(+) " & IIf(strChoose = "", "", " And ") & strChoose

    '分頁方式
    Select Case True
        Case optAreaCode.Value
             strGroupName = "ReportType=True;GroupName={SO014.AreaCode};GroupName1={SO014.AreaName}"
             strpagetype = "行政區"
             strField = "SO014.AreaName"
        Case optServCode.Value
             strGroupName = "ReportType=True;GroupName={SO001.ServCode};GroupName1={SO001.ServArea}"
             strpagetype = "服務區"
             strField = "SO001.ServArea"
        Case optNothing.Value
            Select Case Left(gimOrder.GetColumnOrderCode, 1)
              Case "A"
                strGroupName = "ReportType=False;GroupName={SO002.CustStatusCode};GroupName1={SO001.CustName}"
              Case "B"
                strGroupName = "ReportType=False;GroupName={SO001.ClassCode1};GroupName1={SO001.CustName}"
              Case "C"
                strGroupName = "ReportType=False;GroupName={SO001.CustID};GroupName1={SO001.CustName}"
              Case "D"
                strGroupName = "ReportType=False;GroupName={SO014.AddrSort};GroupName1={SO014.AddrSort}"
              Case "E"
                strGroupName = "ReportType=False;GroupName={SO002.InstTime};GroupName1={SO001.CustName}"
              Case Else
                strGroupName = "ReportType=False;GroupName={SO001.CustID};GroupName1={SO001.CustName}"
            End Select
            strpagetype = "無"
            strField = "SO001.ServArea"
    End Select
    
    
  '排序方式
    If gimOrder.GetColumnOrderCode <> "" Then
      arrOrder = Split(gimOrder.GetColumnOrderCode, ",")
      intSort = 0
      For Each varOrder In arrOrder
        Select Case arrOrder(intSort)
          Case "A"
            strGroupName = strGroupName & ";Sort" & intSort & "={SO002.CustStatusCode}"
          Case "B"
            strGroupName = strGroupName & ";Sort" & intSort & "={SO001.ClassCode1}"
          Case "C"
            strGroupName = strGroupName & ";Sort" & intSort & "={SO001.CustID}"
          Case "D"
            strGroupName = strGroupName & ";Sort" & intSort & "={SO014.AddrSort}"
          Case "E"
            strGroupName = strGroupName & ";Sort" & intSort & "=Date({SO002.InstTime})"
        End Select
        intSort = intSort + 1
      Next
    End If
    
    strChooseString = "裝機日期: " & subSpace(gdaInstDate1.GetValue(True)) & "~" & subSpace(gdaInstDate2.GetValue(True)) & ";" & _
                      "收費人員:" & subSpace(gilClctEn.GetDescription) & ";" & _
                      "綱路編號: " & subSpace(mskCircuitNo.Text) & ";" & _
                      "公司別　:" & subSpace(gilCompCode.GetDescription) & ";" & _
                      "服務類別: " & subSpace(gilServiceType.GetDescription) & ";" & _
                      "客戶狀態:" & subSpace(gimStatusCode.GetDispStr) & ";" & _
                      "停拆移機類狀態:" & subSpace(gimWipCode3.GetDispStr) & ";" & _
                      "客戶類別:" & subSpace(gimClassCode.GetDispStr) & ";" & _
                      "客戶類別二:" & subSpace(gimClassCode2.GetDispStr) & ";" & _
                      "客戶類別三:" & subSpace(gimClassCode3.GetDispStr) & ";" & _
                      "收費項目:" & subSpace(gimCitemCode.GetDispStr) & ";" & _
                      "服務區　:" & subSpace(gimServCode.GetDispStr) & ";" & _
                      "行政區　:" & subSpace(gimAreaCode.GetDispStr) & ";" & _
                      "街道範圍:" & subSpace(gimStrtCode.GetDispStr) & ";" & _
                      "排序方式:" & subSpace(gimOrder.GetColumnOrderDspStr) & ";" & _
                      "無週期性定義:" & strCitem & ";" & _
                      "分頁方式: " & strpagetype
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "subChoose")
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error GoTo ChkErr
    Call FunctionKey(KeyCode, Shift, frmSO5570A)
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
    Call SetgiMulti(gimStatusCode, "CodeNo", "Description", "CD035", "客戶狀態代碼", "客戶狀態名稱")
    Call SetgiMulti(gimClassCode, "CodeNo", "Description", "CD004", "客戶類別代碼", "客戶類別名稱")
    
    '********************************************************************************************************
    '#3705 增加客戶類別二與三 By Kin 2008/02/14
    Call SetgiMulti(gimClassCode2, "CodeNo", "Description", "CD004", "客戶類別代碼", "客戶類別名稱")
    Call SetgiMulti(gimClassCode3, "CodeNo", "Description", "CD004", "客戶類別代碼", "客戶類別名稱")
    '*********************************************************************************************************
    
    Call SetgiMulti(gimWipCode3, "CodeNo", "Description", "CD036", "派工狀態代碼", "派工狀態名稱")      '問題集2322
    Call SetgiMulti(gimCitemCode, "CodeNo", "Description", "CD019", "收費項目代碼", "收費項目名稱", " Where Period>0 ")
    Call SetgiMulti(gimServCode, "CodeNo", "Description", "CD002", "服務區代碼", "服務區名稱")
    Call SetgiMulti(gimAreaCode, "CodeNo", "Description", "CD001", "行政區代碼", "行政區名稱")
    Call SetgiMulti(gimStrtCode, "CodeNo", "Description", "CD017", "街道代碼", "街道名稱")
    Call SetgiMultiAddItem(gimOrder, "A,B,C,D,E", "客戶狀態,客戶類別,客戶編號,裝機地址,裝機日期")
    gimStatusCode.SetQueryCode "1,2,3"
    gimWipCode3.SetQueryCode "13"
   Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "subGim")
End Sub

Private Sub subGil()
  On Error GoTo ChkErr
    'Call SetgiList(gilCompCode, "CodeNo", "Description", "CD039")
    Call SetgiList(gilCompCode, "CodeNo", "Description", "CD039", , , , , , , GetCompCodeFilter(gilCompCode))
    Call SetgiList(gilClctEn, "EmpNo", "EmpName", "CM003")
    Call SetgiList(gilServiceType, "CodeNo", "Description", "CD046")
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "subGil")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  On Error Resume Next
    gcnGi.Execute "Delete From So075 Where BatchNo = " & lngLastBatchNo
    Call DropTMPVIEW(strViewName, strSubViewName())
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
  ReleaseCOM frmSO5570A
End Sub

Private Sub gdaInstDate1_GotFocus()
  On Error Resume Next
  If gdaInstDate1.GetValue = "" Then gdaInstDate1.SetValue (RightDate)
End Sub

Private Sub gdaInstDate2_GotFocus()
  On Error Resume Next
    If gdaInstDate1.GetValue = "" Or gdaInstDate2.GetValue = "" Then gdaInstDate2.SetValue (gdaInstDate1.GetValue)
End Sub

Private Sub gdaInstDate2_Validate(Cancel As Boolean)
  On Error Resume Next
    Cancel = ChkDate2(gdaInstDate1, gdaInstDate2)
End Sub

Private Sub gilCompCode_Change()
  On Error GoTo ChkErr
    Call SelectServType(gilCompCode.GetCodeNo, gilServiceType)
    gilServiceType.ListIndex = 1
    GiMultiFilter gimServCode, , gilCompCode.GetCodeNo
    GiMultiFilter gimAreaCode, , gilCompCode.GetCodeNo
    GiMultiFilter gimStrtCode, , gilCompCode.GetCodeNo
    GiListFilter gilClctEn, , gilCompCode.GetCodeNo
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
    gimCitemCode.Filter = gimCitemCode.Filter & IIf(gimCitemCode.Filter = "", " WHERE ", " AND ") & " PeriodFlag = 1"
    GiMultiFilter gimClassCode, gilServiceType.GetCodeNo
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "gilCompCode_Validate")
End Sub

Private Sub mskCircuitNo_Validate(Cancel As Boolean)
  If mskCircuitNo.Text = "" Then
    optCircuitNo.Enabled = True
    optAll.Enabled = True
  Else
    optCircuitNo.Enabled = False
    optAll.Enabled = False
  End If
End Sub

'Private Function GetUseIndexStr(StrTableName As String, strColumnName As String) As String
'  On Error Resume Next
'    GetUseIndexStr = UCase(" /*+ Index(" & StrTableName & " I_" & StrTableName & "_" & strColumnName & ") */ ")
'End Function

Private Sub VScroll1_Change()
End Sub

Private Sub vsl1_Change()
 On Error Resume Next
    If vsl1.Value = 0 Then
        fraMulti.Top = 20
    ElseIf vsl1.Value = 100 Then
        fraMulti.Top = -(pic2.Height) + 480
    End If

End Sub
