VERSION 5.00
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Begin VB.Form frmSO5560A 
   BorderStyle     =   1  '單線固定
   Caption         =   "各服務區/行政區客戶增減數統計表 [SO5560A]"
   ClientHeight    =   4395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7695
   Icon            =   "SO5560A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton cmdExcel 
      Caption         =   "匯成Excel"
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
      Left            =   3600
      TabIndex        =   12
      Top             =   3690
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
      Left            =   1830
      TabIndex        =   11
      Top             =   3690
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
      TabIndex        =   13
      Top             =   3690
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
      Left            =   210
      TabIndex        =   10
      Top             =   3690
      Width           =   1245
   End
   Begin VB.Frame fraPage 
      Caption         =   "報表種類"
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
      Height          =   645
      Left            =   30
      TabIndex        =   15
      Top             =   1230
      Width           =   3735
      Begin VB.OptionButton optClctAreaCode 
         Caption         =   "收費區"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   2520
         TabIndex        =   6
         Top             =   270
         Width           =   915
      End
      Begin VB.OptionButton optAreaCode 
         Caption         =   "行政區"
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
         Left            =   1380
         TabIndex        =   5
         Top             =   270
         Width           =   915
      End
      Begin VB.OptionButton optServCode 
         Caption         =   "服務區"
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
         Left            =   300
         TabIndex        =   4
         Top             =   270
         Value           =   -1  'True
         Width           =   915
      End
   End
   Begin VB.Frame fraCustStatus 
      Caption         =   "統計對象"
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
      Height          =   645
      Left            =   3900
      TabIndex        =   14
      Top             =   1230
      Width           =   3705
      Begin VB.OptionButton optNormal 
         Caption         =   "非大樓戶"
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
         Left            =   2460
         TabIndex        =   9
         Top             =   270
         Width           =   1125
      End
      Begin VB.OptionButton optMdu 
         Caption         =   "大樓戶"
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
         Left            =   1230
         TabIndex        =   8
         Top             =   270
         Width           =   885
      End
      Begin VB.OptionButton optAll 
         Caption         =   "全部"
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
         Left            =   210
         TabIndex        =   7
         Top             =   270
         Value           =   -1  'True
         Width           =   765
      End
   End
   Begin Gi_Date.GiDate gdaTime2 
      Height          =   345
      Left            =   2310
      TabIndex        =   1
      Top             =   150
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
      Top             =   150
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
   Begin prjGiList.GiList gilServiceType 
      Height          =   315
      Left            =   4590
      TabIndex        =   3
      Top             =   570
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
      Left            =   540
      TabIndex        =   23
      Top             =   2400
      Width           =   825
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "停機時-->停機時間     註銷時-->註銷時間"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   1440
      TabIndex        =   22
      Top             =   2610
      Width           =   3165
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "裝機時-->裝機時間     拆機時-->拆機時間"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   1440
      TabIndex        =   21
      Top             =   2400
      Width           =   3165
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
      Left            =   120
      TabIndex        =   20
      Top             =   2040
      Width           =   510
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
      Top             =   210
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
      Left            =   3765
      TabIndex        =   18
      Top             =   645
      Width           =   780
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "~"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   2160
      TabIndex        =   17
      Top             =   240
      Width           =   165
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
      TabIndex        =   16
      Top             =   240
      Width           =   780
   End
End
Attribute VB_Name = "frmSO5560A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Table=SO001,SO014
'DataBase=MDB [SO5560A]
Option Explicit
Dim MdbChoose As String
Dim blnFlag As Boolean
Dim strField As String
Dim strField1 As String
Dim strPage As String
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
    Call PreviousRpt(GetPrinterName(5), RptName("SO5560"), Me.Caption)
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
            Call PrintRpt(GetPrinterName(5), RptName("SO5560"), , Me.Caption, strSQL, strChooseString, , True, , , strGroupName, GiPaperLandscape)
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
    
      RptToTxt RptName("SO5560"), , strSQL, strChooseString, _
               , , strGroupName, , , garyGi(19) & "\SO5560"
      
      If Not Get_RS_From_Txt(garyGi(19), "SO5560.txt", rsExcel) Then blnExcel = False: Exit Sub
      '#7074 To Add a choice-condition parameter into excel print function by Kin 2015/09/09
      Call UseProperty(rsExcel, "各" & strPage & "客戶增減數統計表 ", "第一頁", _
                        , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , strChooseString)
      blnExcel = False
      CloseRecordset rsExcel
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "toExcel")
End Sub

Private Sub subChoose()
  On Error GoTo ChkErr
    Dim strCustStatus  As String
    'strChoose = "SO001.InstAddrNo=SO014.AddrNo"
    strChoose = ""
    strChooseString = ""
    strPage = ""
    
  'GiList
    If gilCompCode.GetCodeNo <> "" Then Call subAnd("SO001.CompCode=" & gilCompCode.GetCodeNo)
    If gilServiceType.GetCodeNo <> "" Then Call subAnd("instr(SO001.ServiceType,'" & gilServiceType.GetCodeNo & "')>0")
    If gilServiceType.GetCodeNo <> "" Then Call subAnd("SO002.ServiceType='" & gilServiceType.GetCodeNo & "'")
    'If gilServiceType.GetCodeNo <> "" Then Call subAnd("SO001.ServiceType='" & gilServiceType.GetCodeNo & "'")
    'If gilServiceType.GetCodeNo <> "" Then Call subAnd("Instr(Nvl(SO001.ServiceType,'" & gilServiceType.GetCodeNo & "'),'" & gilServiceType.GetCodeNo & "') > 0")
    
  '統計對象
    Select Case True
           Case optMdu.Value
                subAnd "SO001.MduId Is Not Null"
                strCustStatus = "大樓戶"
           Case optNormal.Value
                subAnd "SO001.MduId Is Null"
                strCustStatus = "非大樓戶"
           Case Else
                strCustStatus = "全部"
    End Select
   
  '報表種類
    Select Case True
           Case optAreaCode.Value
                strGroupName = "TitleCode='行政區代碼';TitleName='行政區名稱'"
                strField = "SO014.AreaCode ServCode,SO014.AreaName ServArea"
                strField1 = "SO014.AreaCode,SO014.AreaName"
                strChoose = "From SO001,SO002,SO014 Where SO001.Custid=SO002.Custid And SO001.InstAddrNo=SO014.AddrNo AND " & IIf(strChoose = "", "", strChoose & " And ")
                strPage = "行政區"
           Case optServCode.Value
                strGroupName = "TitleCode='服務區代碼';TitleName='服務區名稱'"
                strField = "SO001.ServCode ServCode,SO001.ServArea ServArea"
                strField1 = "SO001.ServCode,SO001.ServArea "
                strChoose = "From SO001,SO002 Where SO001.Custid=SO002.Custid AND " & IIf(strChoose = "", "", strChoose & " And ")
                strPage = "服務區"
           Case optClctAreaCode.Value
                strGroupName = "TitleCode='收費區代碼';TitleName='收費區名稱'"
                strField = "SO001.ClctAreaCode ServCode,SO001.ClctAreaName ServArea"
                strField1 = "SO001.ClctAreaCode,SO001.ClctAreaName "
                strChoose = "From SO001,SO002 Where SO001.Custid=SO002.Custid AND " & IIf(strChoose = "", "", strChoose & " And ")
                strPage = "收費區"
    End Select
          
   strChooseString = "日期範圍:" & subSpace(gdaTime1.GetValue(True)) & "~" & subSpace(gdaTime2.GetValue(True)) & ";" & _
                     "服務類別: " & subSpace(gilServiceType.GetDescription) & ";" & _
                     "公司別　:" & subSpace(gilCompCode.GetDescription) & ";" & _
                     "統計對象:" & subSpace(strCustStatus) & ";" & _
                     "報表種類:" & subSpace(strPage)
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
    Call FunctionKey(KeyCode, Shift, frmSO5560A)
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "Form_KeyDown")
End Sub

Private Sub Form_Load()
  On Error GoTo ChkErr
    Call subGil
    gilCompCode.SetCodeNo garyGi(9)
    gilCompCode.Query_Description
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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  On Error Resume Next
  gcnGi.Execute "Drop View " & strViewName
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
  ReleaseCOM frmSO5560A
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
    MdbChoose = ""
    
      '裝機
        MdbChoose = " SO002.CustStatusCode=1 "
        If gdaTime1.GetValue <> "" Then Call MdbAnd("SO002.InstTime >= To_Date('" & gdaTime1.GetValue & "','YYYYMMDD')")
        If gdaTime2.GetValue <> "" Then Call MdbAnd("SO002.InstTime < To_Date('" & gdaTime2.GetValue & "','YYYYMMDD')+1")
        strSQL = " SELECT " & strField & ",Count(Distinct SO001.Custid) as InterCount,0 as StopCount,0 as PRCount,0 as DelCount " & _
                  strChoose & MdbChoose & " Group by " & strField1
                 
      '停機
        MdbChoose = " SO002.CustStatusCode=2 "
        If gdaTime1.GetValue <> "" Then Call MdbAnd("SO002.StopTime >= To_Date('" & gdaTime1.GetValue & "','YYYYMMDD')")
        If gdaTime2.GetValue <> "" Then Call MdbAnd("SO002.StopTime < To_Date('" & gdaTime2.GetValue & "','YYYYMMDD')+1")
        strSQL = strSQL & " Union All SELECT " & strField & ",0 as InterCount,Count(Distinct SO001.Custid) as StopCount,0 as PRCount,0 as DelCount " & _
                  strChoose & MdbChoose & " Group by " & strField1
                 
      '拆機
        MdbChoose = " SO002.CustStatusCode=3 "
        If gdaTime1.GetValue <> "" Then Call MdbAnd("SO002.PRTime >= To_Date('" & gdaTime1.GetValue & "','YYYYMMDD')")
        If gdaTime2.GetValue <> "" Then Call MdbAnd("SO002.PRTime < To_Date('" & gdaTime2.GetValue & "','YYYYMMDD')+1")
        strSQL = strSQL & " Union All SELECT " & strField & ",0 as InterCount,0 as StopCount,Count(Distinct SO001.Custid) as PRCount,0 as DelCount " & _
                  strChoose & MdbChoose & " Group by " & strField1

      '註銷
        MdbChoose = " SO002.CustStatusCode=4 "
        If gdaTime1.GetValue <> "" Then Call MdbAnd("SO002.DelDate >= To_Date('" & gdaTime1.GetValue & "','YYYYMMDD')")
        If gdaTime2.GetValue <> "" Then Call MdbAnd("SO002.DelDate < To_Date('" & gdaTime2.GetValue & "','YYYYMMDD')+1")
        strSQL = strSQL & " Union All SELECT " & strField & ",0 as InterCount,0 as StopCount,0 as PRCount,Count(Distinct SO001.Custid) as DelCount " & _
                  strChoose & MdbChoose & " Group by " & strField1

      subCreateView (strSQL)

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
                  "Select A.ServCode,A.ServArea,Sum(A.InterCount) as InterCount,Sum(A.StopCount) as StopCount," & _
                  "Sum(A.PRCount) as PRCount,Sum(A.DelCount) as DelCount From (" & strSQL & _
                  ") A Group by A.ServCode,A.ServArea) "
    gcnGi.Execute strSQL
    SendSQL strSQL, True
  Exit Sub
ChkErr:
    ErrSub Me.Name, "subCreateView"
End Sub

Private Sub gilCompCode_Change()
  On Error GoTo ChkErr
    Call SelectServType(gilCompCode.GetCodeNo, gilServiceType)
    gilServiceType.ListIndex = 1
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
