VERSION 5.00
Object = "{A0C51418-555F-11D3-BC10-0050BA01A3A9}#9.8#0"; "GiMulti.ocx"
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{074771B7-6ADE-11D5-B178-00485456B028}#1.16#0"; "csMulti.ocx"
Begin VB.Form frmSO51A0A 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "來電業務期間報表[SO51A0]"
   ClientHeight    =   4155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7215
   Icon            =   "SO51A0A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "分頁方式"
      ForeColor       =   &H00FF0000&
      Height          =   720
      Left            =   3150
      TabIndex        =   17
      Top             =   2220
      Width           =   2685
      Begin VB.OptionButton optNothing 
         Caption         =   "無"
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   1635
         TabIndex        =   19
         Top             =   270
         Value           =   -1  'True
         Width           =   555
      End
      Begin VB.OptionButton optAreaCode 
         Caption         =   "行政區"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   315
         TabIndex        =   18
         Top             =   300
         Width           =   1095
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
      Left            =   1920
      TabIndex        =   7
      Top             =   3480
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
      Left            =   5160
      TabIndex        =   8
      Top             =   3480
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
      Left            =   240
      TabIndex        =   6
      Top             =   3480
      Width           =   1245
   End
   Begin VB.Frame fraGroup 
      Caption         =   "統計依據"
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
      Height          =   705
      Left            =   60
      TabIndex        =   9
      Top             =   2220
      Width           =   2895
      Begin VB.OptionButton optNode 
         Caption         =   "Node No"
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
         Height          =   195
         Left            =   1350
         TabIndex        =   5
         Top             =   300
         Width           =   1125
      End
      Begin VB.OptionButton optAcceptTime 
         Caption         =   "日期"
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
         Height          =   255
         Left            =   300
         TabIndex        =   4
         Top             =   300
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin Gi_Date.GiDate gdaAcceptTime1 
      Height          =   375
      Left            =   870
      TabIndex        =   0
      Top             =   60
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
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
      Left            =   4230
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
   Begin prjGiList.GiList gilServiceType 
      Height          =   315
      Left            =   4230
      TabIndex        =   3
      Top             =   540
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
   Begin Gi_Date.GiDate gdaAcceptTime2 
      Height          =   375
      Left            =   2220
      TabIndex        =   1
      Top             =   60
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
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
   Begin Gi_Multi.GiMulti gimBulletinName 
      Height          =   375
      Left            =   0
      TabIndex        =   15
      Top             =   960
      Width           =   7245
      _ExtentX        =   12779
      _ExtentY        =   661
      ButtonCaption   =   "消   息   來   源"
   End
   Begin Gi_Multi.GiMulti gimAreaCode 
      Height          =   375
      Left            =   0
      TabIndex        =   16
      Top             =   1740
      Width           =   7245
      _ExtentX        =   12779
      _ExtentY        =   661
      ButtonCaption   =   "行     政     區"
   End
   Begin CS_Multi.CSmulti gimPromCode 
      Height          =   375
      Left            =   0
      TabIndex        =   14
      Top             =   1350
      Width           =   7245
      _ExtentX        =   12779
      _ExtentY        =   661
      ButtonCaption   =   "業   務   活   動"
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "---"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   2010
      TabIndex        =   13
      Top             =   150
      Width           =   180
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "發送日期"
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
      Left            =   60
      TabIndex        =   12
      Top             =   150
      Width           =   780
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   3390
      TabIndex        =   11
      Top             =   600
      Width           =   780
   End
   Begin VB.Label lblCompCode 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
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
      Left            =   3390
      TabIndex        =   10
      Top             =   150
      Width           =   765
   End
End
Attribute VB_Name = "frmSO51A0A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Run此報表前需先Run[SO084 其他日結]
'Table:SO096
Option Explicit
Dim strField As String

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
      Call PreviousRpt(GetPrinterName(5), RptName("SO51A0"), Me.Caption)
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
        Call subCreateView
        Call subPrint
      Me.Enabled = True
    Screen.MousePointer = vbDefault
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdPrint_Click")
End Sub

Private Sub subPrint()
  On Error GoTo ChkErr
  Dim rsTmp As New ADODB.Recordset
    Set rsTmp = gcnGi.Execute("SELECT Count(*) as intCount From " & strViewName & " Where RowNum=1")
    If rsTmp("intCount") = 0 Then
        MsgNoRcd
        SendSQL , , True
    Else
        CreateTable
        strsql = "SELECT * From SO51A0A "
        Call PrintRpt(GetPrinterName(5), RptName("SO51A0"), , Me.Caption, strsql, strChooseString, , True, "Tmp0000.MDB", , strGroupName)
    End If
    CloseRecordset rsTmp
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "subPrint")
End Sub

Private Function IsDataOk() As Boolean
  On Error GoTo ChkErr
    If Len(gilCompCode.GetCodeNo) = 0 Then gilCompCode.SetFocus: strErrFile = "公司別": GoTo 66
    If Len(gdaAcceptTime1.GetValue) = 0 Then gdaAcceptTime1_GotFocus: strErrFile = "起始日期": GoTo 66
    If Len(gdaAcceptTime2.GetValue) = 0 Then gdaAcceptTime2_GotFocus: strErrFile = "截止日期": GoTo 66
    If gdaAcceptTime2.GetValue(True) >= Format(DateAdd("m", 1, gdaAcceptTime1.Text), "YYYY/MM/DD") Then
        MsgBox "日期範圍不能超過一個月", , "錯誤訊息"
        gdaAcceptTime2_GotFocus
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
    Dim strTypeName As String
      strChoose = ""
      strChooseString = ""
      strGroupName = ""
      
    '日期
      If gdaAcceptTime1.GetValue <> "" Then Call subAnd("SO096.AcceptTime >= To_Date('" & gdaAcceptTime1.GetValue & "','YYYYMMDD')")
      If gdaAcceptTime2.GetValue <> "" Then Call subAnd("SO096.AcceptTime < To_Date('" & gdaAcceptTime2.GetValue & "','YYYYMMDD') +1")

    'GiList
      If gilCompCode.GetCodeNo <> "" Then Call subAnd("SO096.CompCode=" & gilCompCode.GetCodeNo)
      If gilServiceType.GetCodeNo <> "" Then Call subAnd("SO096.ServiceType='" & gilServiceType.GetCodeNo & "'")

    'GiMulti
      If gimBulletinName.GetQryStr <> "" Then Call subAnd("SO096.BulletinCode " & gimBulletinName.GetQryStr)
      If gimPromCode.GetQryStr <> "" Then Call subAnd("SO096.PromCode " & gimPromCode.GetQryStr)
      If gimAreaCode.GetQryStr <> "" Then Call subAnd("SO096.AreaCode " & gimAreaCode.GetQryStr)
      
    '統計依據
      Select Case True
        Case optAcceptTime.Value '日期
             strTypeName = "日期"
             strField = "to_char(SO096.AcceptTime,'MM')||'月'||to_char(SO096.AcceptTime,'DD')||'日'"
             strGroupName = "GroupTitle='Date'"
        Case optNode.Value 'Node
             strTypeName = "NodeNo"
             strGroupName = "GroupTitle='NodeNo'"
             strField = "SO096.NodeNo"
      End Select

      strChooseString = "日期範圍:" & subSpace(gdaAcceptTime1.GetOriginalValue) & "~" & subSpace(gdaAcceptTime2.GetOriginalValue) & ";" & _
                        "公司別　:" & subSpace(gilCompCode.GetDescription) & ";" & _
                        "服務類別:" & subSpace(gilServiceType.GetDescription) & ";" & _
                        "消息來源:" & subSpace(gimBulletinName.GetDispStr) & ";" & _
                        "業務活動:" & subSpace(gimPromCode.GetDispStr) & ";" & _
                        "行政區　:" & subSpace(gimAreaCode.GetDispStr) & ";" & _
                        "分頁方式:" & IIf(optAreaCode.Value = True, "行政區", "無") & ";" & _
                        "統計依據:" & strTypeName
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "subChoose")
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error GoTo ChkErr
    Call FunctionKey(KeyCode, Shift, frmSO51A0A)
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "Form_KeyDown")
End Sub

Private Sub Form_Load()
  On Error GoTo ChkErr
    If Alfa2 Then Call GetGlobal
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
    Call SetgiMulti(gimBulletinName, "CodeNo", "Description", "CD049", "消息來源代碼", "消息來源名稱")
    Call SetgiMulti(gimPromCode, "CodeNo", "Description", "CD042", "促銷活動代碼", "促銷活動名稱")
    Call SetgiMulti(gimAreaCode, "CodeNo", "Description", "CD001", "行政區代碼", "行政區名稱")
    gimBulletinName.SelectAll
    gimPromCode.SelectAll
    
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
    gcnGi.Execute "Drop View " & strViewName
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
  ReleaseCOM frmSO51A0A
End Sub

Private Sub gdaAcceptTime1_GotFocus()
  On Error Resume Next
    If gdaAcceptTime1.GetValue = "" Then gdaAcceptTime1.SetValue (Format(RightDate, "YYYYMM") & "01")
End Sub

Private Sub gdaAcceptTime2_GotFocus()
  On Error Resume Next
    If gdaAcceptTime1.GetValue <> "" Then
      gdaAcceptTime2.SetValue GetLastDate(gdaAcceptTime1.GetValue(True))
    Else
      gdaAcceptTime2.SetValue ""
    End If
End Sub

Private Sub gdaAcceptTime2_Validate(Cancel As Boolean)
  On Error Resume Next
  Cancel = ChkDate2(gdaAcceptTime1, gdaAcceptTime2)
End Sub

Private Sub gilCompCode_Change()
  On Error GoTo ChkErr
    Call SelectServType(gilCompCode.GetCodeNo, gilServiceType)
    gilServiceType.ListIndex = 1
  
    GiMultiFilter gimBulletinName, , gilCompCode.GetCodeNo
    GiMultiFilter gimPromCode, , gilCompCode.GetCodeNo
    
    gimPromCode.SelectAll
    gimBulletinName.SelectAll
  
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

Private Function subCreateView() As Boolean
  Dim strView As String
  Dim strAreaCode As String
  Dim strAreaName As String
    On Error Resume Next
      gcnGi.Execute "Drop View " & strViewName
    On Error GoTo ChkErr
        strAreaCode = IIf(optAreaCode.Value = True, "SO096.AreaCode", "999")
        strAreaName = IIf(optAreaCode.Value = True, "SO096.AreaName", "'無'")
        strViewName = GetTmpViewName
        
        strView = "Create View " & strViewName & _
                  " as (SELECT " & strField & " as GroupName," & strAreaCode & " as AreaCode," & strAreaName & " as AreaName," & _
                  "SO096.BulletinName AS Description,MediaStatus," & _
                  "'  CALL' as TypeName,sum(so096.Calls) as Quantity From SO096 Where " & strChoose & _
                  " Group By " & strField & "," & strAreaCode & "," & strAreaName & ",BulletinName,MediaStatus  UNION All " & _
                  "SELECT " & strField & " as GroupName," & strAreaCode & " as AreaCode," & strAreaName & " as AreaName," & _
                  "SO096.BulletinName AS Description,MediaStatus," & _
                  "' SALE' as TypeName,sum(so096.Sales) as Quantity From SO096 Where " & strChoose & _
                  " Group By " & strField & "," & strAreaCode & "," & strAreaName & ",BulletinName,MediaStatus  UNION All " & _
                  "SELECT " & strField & " as GroupName," & strAreaCode & " as AreaCode," & strAreaName & " as AreaName," & _
                  "SO096.BulletinName AS Description,MediaStatus," & _
                  "'INSTALL' as TypeName,sum(so096.Install) as Quantity From SO096 Where " & strChoose & _
                  " Group By " & strField & "," & strAreaCode & "," & strAreaName & ",BulletinName,MediaStatus  UNION All " & _
                  "SELECT " & strField & " as GroupName," & strAreaCode & " as AreaCode," & strAreaName & " as AreaName," & _
                  "'" & Chr(255) & "Total' AS Description ,'' as MediaStatus," & _
                  "'  CALL' as TypeName,sum(so096.Calls) as Quantity From SO096 Where " & strChoose & _
                  " Group By " & strField & "," & strAreaCode & "," & strAreaName & " UNION All " & _
                  "SELECT " & strField & " as GroupName," & strAreaCode & " as AreaCode," & strAreaName & " as AreaName," & _
                  "'" & Chr(255) & "Total' AS Description,'' as MediaStatus," & _
                  "' SALE' as TypeName,sum(so096.Sales) as Quantity From SO096 Where " & strChoose & _
                  " Group By " & strField & "," & strAreaCode & "," & strAreaName & " UNION All " & _
                  "SELECT " & strField & " as GroupName," & strAreaCode & " as AreaCode," & strAreaName & " as AreaName," & _
                  "'" & Chr(255) & "Total' AS Description,'' as MediaStatus," & _
                  "'INSTALL' as TypeName,sum(so096.Install) as Quantity From SO096 Where " & strChoose & _
                  " Group By " & strField & "," & strAreaCode & "," & strAreaName & ")"
        gcnGi.Execute strView
        SendSQL strView, True
    Exit Function
ChkErr:
  Call ErrSub(Me.Name, "subCreateView")
End Function

Private Sub CreateTable()
  On Error Resume Next
    cnn.Execute "Drop Table SO51A0A"
  On Error GoTo ChkErr
    cnn.Execute "Create Table SO51A0A (GroupName Text(40),AreaCode Text(6),AreaName Text(40),Description Text(24),MediaStatus Text(40),TypeName Text(24),Quantity Double)"
    Call subInsertMDB("Select * From " & strViewName)
    If optAcceptTime.Value = True Then Call subInsertMDB1("Select * From " & strViewName & " Order by AreaCode,GroupName")
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "CreateTable")
End Sub

Private Sub subInsertMDB(strsql As String)
  On Error GoTo ChkErr
    Dim rsTmp As New ADODB.Recordset
    Set rsTmp = gcnGi.Execute(strsql)
    cnn.BeginTrans

    While Not rsTmp.EOF
        cnn.Execute "Insert into SO51A0A (GroupName,AreaCode,AreaName,Description,MediaStatus,TypeName,Quantity) Values(" & _
                    GetNullString(rsTmp("GroupName")) & "," & _
                    GetNullString(rsTmp("AreaCode")) & "," & _
                    GetNullString(rsTmp("AreaName")) & "," & _
                    GetNullString(rsTmp("Description")) & "," & _
                    GetNullString(rsTmp("MediaStatus")) & "," & _
                    GetNullString(rsTmp("TypeName")) & "," & _
                    GetNullString(rsTmp("Quantity")) & ")"
        rsTmp.MoveNext
    Wend
    cnn.CommitTrans
  Exit Sub
ChkErr:
  cnn.RollbackTrans
  Call ErrSub(Me.Name, "subInsertMDB")
End Sub

Private Sub subInsertMDB1(strsql As String)
  On Error GoTo ChkErr
    Dim rsTmp As New ADODB.Recordset
    Dim strDate As String
    Dim datDate As Date
    Set rsTmp = gcnGi.Execute(strsql)
    datDate = gdaAcceptTime1.GetValue(True)
    cnn.BeginTrans

    While Not rsTmp.EOF
        strDate = Format(Month(datDate), "00") & "月" & Format(Day(datDate), "00") & "日"
        cnn.Execute "Insert into SO51A0A (GroupName,AreaCode,AreaName,Description,MediaStatus,TypeName,Quantity) Values('" & _
                    strDate & "'," & _
                    GetNullString(rsTmp("AreaCode")) & "," & _
                    GetNullString(rsTmp("AreaName")) & "," & _
                    GetNullString(rsTmp("Description")) & "," & _
                    GetNullString(rsTmp("MediaStatus")) & "," & _
                    GetNullString(rsTmp("TypeName")) & "," & _
                    0 & ")"
                    
        If datDate < CDate(gdaAcceptTime2.GetValue(True)) Then
            datDate = DateAdd("d", 1, datDate)
        Else
            datDate = gdaAcceptTime1.GetValue(True)
        End If
        rsTmp.MoveNext
        DoEvents
    Wend
    cnn.CommitTrans
  Exit Sub
ChkErr:
  cnn.RollbackTrans
  Call ErrSub(Me.Name, "subInsertMDB1")
End Sub

