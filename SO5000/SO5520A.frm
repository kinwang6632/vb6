VERSION 5.00
Object = "{A0C51418-555F-11D3-BC10-0050BA01A3A9}#9.8#0"; "GiMulti.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{074771B7-6ADE-11D5-B178-00485456B028}#1.16#0"; "csMulti.ocx"
Begin VB.Form frmSO5520A 
   BorderStyle     =   1  '單線固定
   Caption         =   "各街道客戶狀態統計 [SO5520A]"
   ClientHeight    =   5610
   ClientLeft      =   1950
   ClientTop       =   1530
   ClientWidth     =   7410
   BeginProperty Font 
      Name            =   "新細明體"
      Size            =   9.75
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SO5520A.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   7410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '螢幕中央
   Begin VB.Frame fraGroupName 
      Caption         =   "分區方式"
      ForeColor       =   &H00C00000&
      Height          =   1185
      Left            =   0
      TabIndex        =   26
      Top             =   3660
      Width           =   2295
      Begin VB.OptionButton optNothing 
         Caption         =   "無"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   1230
         TabIndex        =   12
         Top             =   750
         Value           =   -1  'True
         Width           =   555
      End
      Begin VB.OptionButton optServCode 
         Caption         =   "服務區"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   150
         TabIndex        =   9
         Top             =   360
         Width           =   915
      End
      Begin VB.OptionButton optAreaCode 
         Caption         =   "行政區"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   1230
         TabIndex        =   10
         Top             =   360
         Width           =   915
      End
      Begin VB.OptionButton optClctAreaCode 
         Caption         =   "收費區"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   150
         TabIndex        =   11
         Top             =   750
         Width           =   915
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "* 統計方式如選擇街道可再挑選統計類別"
      ForeColor       =   &H000000FF&
      Height          =   1185
      Left            =   2340
      TabIndex        =   23
      Top             =   3660
      Width           =   5085
      Begin VB.Frame frmType 
         Caption         =   "統計類別"
         ForeColor       =   &H00800080&
         Height          =   735
         Left            =   2220
         TabIndex        =   25
         Top             =   300
         Width           =   2715
         Begin VB.OptionButton optAll 
            Caption         =   "全部"
            ForeColor       =   &H00800080&
            Height          =   255
            Left            =   1920
            TabIndex        =   17
            Top             =   300
            Value           =   -1  'True
            Width           =   705
         End
         Begin VB.OptionButton optMdu 
            Caption         =   "大樓戶"
            ForeColor       =   &H00800080&
            Height          =   255
            Left            =   1020
            TabIndex        =   16
            Top             =   300
            Width           =   885
         End
         Begin VB.OptionButton optGeneral 
            Caption         =   "一般戶"
            ForeColor       =   &H00800080&
            Height          =   255
            Left            =   150
            TabIndex        =   15
            Top             =   300
            Width           =   885
         End
      End
      Begin VB.Frame fraSum 
         Caption         =   "統計方式"
         ForeColor       =   &H00FF0000&
         Height          =   735
         Left            =   150
         TabIndex        =   24
         Top             =   300
         Width           =   1995
         Begin VB.OptionButton optMduId 
            Caption         =   "大樓"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   1020
            TabIndex        =   14
            Top             =   360
            Width           =   765
         End
         Begin VB.OptionButton optStrtCode 
            Caption         =   "街道"
            ForeColor       =   &H00800080&
            Height          =   195
            Left            =   210
            TabIndex        =   13
            Top             =   360
            Value           =   -1  'True
            Width           =   795
         End
      End
   End
   Begin CS_Multi.CSmulti gimStrtCode 
      Height          =   405
      Left            =   0
      TabIndex        =   8
      Top             =   3240
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   714
      ButtonCaption   =   "街  道  名  稱"
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "F5:列印"
      Default         =   -1  'True
      Height          =   525
      Left            =   90
      TabIndex        =   18
      Top             =   4980
      Width           =   1245
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      Height          =   525
      Left            =   6000
      TabIndex        =   20
      Top             =   4980
      Width           =   1245
   End
   Begin VB.CommandButton cmdPrevRpt 
      Caption         =   "上次統計結果"
      Height          =   525
      Left            =   1770
      TabIndex        =   19
      Top             =   4980
      Width           =   1395
   End
   Begin Gi_Multi.GiMulti gimCMName 
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   2085
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   661
      ButtonCaption   =   "收  費  方  式"
   End
   Begin Gi_Multi.GiMulti gimStatusCode 
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   1305
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   661
      ButtonCaption   =   "客  戶  狀  態"
   End
   Begin prjGiList.GiList gilCompCode 
      Height          =   315
      Left            =   945
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
      F5Corresponding =   -1  'True
   End
   Begin Gi_Multi.GiMulti gimAreaCode 
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   2865
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   661
      ButtonCaption   =   "行     政     區"
   End
   Begin Gi_Multi.GiMulti gimServCode 
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   2475
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   661
      ButtonCaption   =   "服     務     區"
   End
   Begin prjGiList.GiList gilServiceType 
      Height          =   315
      Left            =   945
      TabIndex        =   1
      Top             =   525
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
   Begin CS_Multi.CSmulti gimClassCode 
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   1680
      Width           =   7485
      _ExtentX        =   13203
      _ExtentY        =   661
      ButtonCaption   =   "客  戶  類  別"
   End
   Begin CS_Multi.CSmulti gimCircuitNo 
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   960
      Width           =   7485
      _ExtentX        =   13203
      _ExtentY        =   661
      ButtonCaption   =   "網  路  編  號"
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BackStyle       =   0  '透明
      Caption         =   "服務類別"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   90
      TabIndex        =   22
      Top             =   600
      Width           =   780
   End
   Begin VB.Label lblCompCode 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "公司別"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   90
      TabIndex        =   21
      Top             =   180
      Width           =   585
   End
End
Attribute VB_Name = "frmSO5520A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Table:SO014 A,SO001 B,SO002 C,SO017
Option Explicit
Dim strFormula As String
Dim strgroupcode As String

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
    Call PreviousRpt(GetPrinterName(5), RptName("SO5520"), "各街道客戶狀態統計 [SO5520A]")
    Screen.MousePointer = vbDefault
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "cmdPrevRpt_Click")
End Sub

Private Sub cmdPrint_Click()
  On Error GoTo ChkErr
  Dim rsTmp As New ADODB.Recordset
    If Not ChkDTok Then Exit Sub
    If Not IsDataOk Then Exit Sub
    Screen.MousePointer = vbHourglass
    cmdExit.SetFocus
    Me.Enabled = False
    ReadyGoPrint
    Call subChoose
    Call subPrint
    Set rsTmp = cnn.Execute("SELECT Count(*) as intCount FROM SO5520A")
    
    If rsTmp("intCount") = 0 Then
       MsgNoRcd
       SendSQL , , True
    Else
       strsql = "SELECT * FROM SO5520A"
       Call PrintRpt(GetPrinterName(5), RptName("SO5520"), , "各街道客戶狀態統計 [SO5520A]", strsql, strChooseString, , True, "TMP0000.MDB", , strFormula)
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
    Dim strGroupString As String
    Dim strSumCust As String
    Dim strType As String
    strChoose = " Where A.AddrNo=B.InstAddrNo AND B.CUSTID=C.CUSTID "
    strChooseString = ""
    strFormula = ""
    
  'GiList
    If gilCompCode.GetCodeNo <> "" Then Call subAnd("B.CompCode=" & gilCompCode.GetCodeNo)
    If gilServiceType.GetCodeNo <> "" Then Call subAnd("C.ServiceType='" & gilServiceType.GetCodeNo & "'")
    'If gilServiceType.GetCodeNo <> "" Then Call subAnd("Instr(Nvl(B.ServiceType,'" & gilServiceType.GetCodeNo & "'),'" & gilServiceType.GetCodeNo & "') > 0")
    
  'GiMulti
    If gimAreaCode.GetQryStr <> "" Then Call subAnd("A.AreaCode " & gimAreaCode.GetQryStr)
    If gimClassCode.GetQryStr <> "" Then Call subAnd("B.ClassCode1 " & gimClassCode.GetQryStr)
    If gimCMName.GetQryStr <> "" Then Call subAnd("C.CMCode " & gimCMName.GetQryStr)
    If gimServCode.GetQryStr <> "" Then Call subAnd("B.ServCode " & gimServCode.GetQryStr)
    If gimStatusCode.GetQryStr <> "" Then Call subAnd("C.CustStatusCode " & gimStatusCode.GetQryStr)
    If gimStrtCode.GetQryStr <> "" Then Call subAnd("A.StrtCode " & gimStrtCode.GetQryStr)
    
  '網路編號
    If gimCircuitNo.GetQryStr <> "" Then Call subAnd("A.CircuitNo " & gimCircuitNo.GetDescStr)

  '分區方式
    Select Case True
           Case optServCode.Value
                strgroupcode = "A.ServCode"
                strGroupName = "A.ServName "
                strGroupString = "服務區"
           Case optAreaCode.Value
                strgroupcode = "A.AreaCode"
                strGroupName = "A.AreaName "
                strGroupString = "行政區"
           Case optClctAreaCode.Value
                strgroupcode = "A.ClctAreaCode"
                strGroupName = "A.ClctAreaName "
                strGroupString = "收費區"
           Case optNothing.Value
                strgroupcode = "'999'"
                strGroupName = "'無'"
                strGroupString = "無"
    End Select

      '統計方式
    If optMduId.Value Then
       strSumCust = "大樓"
       strFormula = "TitleName='大樓'"
       strType = ""
    Else
       strSumCust = "街道代碼"
       strFormula = "TitleName='街道'"
       Select Case True
              Case optGeneral.Value
                   strType = "一般戶"
              Case optMdu.Value
                   strType = "大樓戶"
              Case optAll.Value
                   strType = "全部"
       End Select
    End If
      
    strChooseString = "公司別　:" & subSpace(gilCompCode.GetDescription) & ";" & _
                      "服務類別:" & subSpace(gilServiceType.GetDescription) & ";" & _
                      "網路編號:" & subSpace(gimCircuitNo.GetDispStr) & ";" & _
                      "客戶狀態:" & subSpace(gimStatusCode.GetDispStr) & ";" & _
                      "客戶類別:" & subSpace(gimClassCode.GetDispStr) & ";" & _
                      "收費方式:" & subSpace(gimCMName.GetDispStr) & ";" & _
                      "服務區　:" & subSpace(gimServCode.GetDispStr) & ";" & _
                      "行政區　:" & subSpace(gimAreaCode.GetDispStr) & ";" & _
                      "街道名稱:" & subSpace(gimStrtCode.GetDispStr) & ";" & _
                      "分區方式:" & subSpace(strGroupString) & ";" & _
                      "統計方式:" & subSpace(strSumCust) & ";" & _
                      "統計類別:" & subSpace(strType)
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "subChoose")
End Sub

Private Sub subPrint()
  On Error GoTo ChkErr
  Dim strCustStatusCode As String
  strCustStatusCode = "Count(Decode(C.CustStatusCode,1,B.CustId,Null)) NormalCount," & _
                      "Count(Decode(C.CustStatusCode,2,B.CustId,Null)) StopCount," & _
                      "Count(Decode(C.CustStatusCode,3,B.CustId,Null)) PRCount," & _
                      "Count(Decode(C.CustStatusCode,4,B.CustId,Null)) DelCount," & _
                      "Count(Decode(C.CustStatusCode,5,B.CustId,Null)) PromCount," & _
                      "Count(Decode(C.CustStatusCode,6,B.CustId,Null)) PRingCount"
  If optMduId.Value Then '大樓
     subInsertMDB ("SELECT " & strgroupcode & " GroupCode," & strGroupName & " GroupName,B.MduId StrtCode,so017.Name StrtName," & strCustStatusCode & " FROM SO014 A,SO001 B,SO002 C,SO017 " & strChoose & " And B.MduID=SO017.MduID And B.MduId Is Not Null Group By " & strgroupcode & "," & strGroupName & ",B.MduId,SO017.Name")
  Else '街道
      Select Case True
             Case optGeneral.Value
                  subInsertMDB ("SELECT " & strgroupcode & " GroupCode," & strGroupName & " GroupName,A.StrtCode StrtCode,A.StrtName," & strCustStatusCode & " FROM SO014 A,SO001 B,SO002 C " & strChoose & " And B.MduId Is Null Group By " & strgroupcode & "," & strGroupName & ",A.StrtCode,A.StrtName")
             Case optMdu.Value
                  subInsertMDB ("SELECT " & strgroupcode & " GroupCode," & strGroupName & " GroupName,A.StrtCode StrtCode,A.StrtName," & strCustStatusCode & " FROM SO014 A,SO001 B,SO002 C " & strChoose & " And B.MduId Is Not Null Group By " & strgroupcode & "," & strGroupName & ",A.StrtCode,A.StrtName")
             Case optAll.Value
                 subInsertMDB ("SELECT " & strgroupcode & " GroupCode," & strGroupName & " GroupName,A.StrtCode StrtCode,A.StrtName," & strCustStatusCode & " FROM SO014 A,SO001 B,SO002 C " & strChoose & " Group By " & strgroupcode & "," & strGroupName & ",A.StrtCode,A.StrtName")
      End Select
  End If

Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "subPrint")
End Sub

Private Sub subInsertMDB(strSQLQry As String)
  On Error GoTo ChkErr
  Dim rsTmp As New ADODB.Recordset
    cnn.Execute "DELETE FROM SO5520A"
    Set rsTmp = gcnGi.Execute(strSQLQry)
    SendSQL strSQLQry, True
    cnn.BeginTrans
      While Not rsTmp.EOF
         cnn.Execute "Insert Into SO5520A (groupcode,GroupName,StrtCode,StrtName,NormalCount,StopCount,PRCount,DelCount,PromCount,PRingCount) Values (" & _
                     GetNullString(rsTmp("GroupCode")) & "," & _
                     GetNullString(rsTmp("GroupName")) & "," & _
                     GetNullString(rsTmp("StrtCode")) & "," & _
                     GetNullString(rsTmp("StrtName")) & "," & _
                     GetNullString(rsTmp("NormalCount")) & "," & _
                     GetNullString(rsTmp("StopCount")) & "," & _
                     GetNullString(rsTmp("PRCount")) & "," & _
                     GetNullString(rsTmp("DelCount")) & "," & _
                     GetNullString(rsTmp("PromCount")) & "," & _
                     GetNullString(rsTmp("PRingCount")) & ")"
         rsTmp.MoveNext
         DoEvents
      Wend
    cnn.CommitTrans
  CloseRecordset rsTmp
  Exit Sub
ChkErr:
  cnn.RollbackTrans
    Call ErrSub(Me.Name, "subInsertMDB")
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error GoTo ChkErr
    Call FunctionKey(KeyCode, Shift, frmSO5520A)
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "Form_KeyDown")
End Sub

Private Sub Form_Load()
  On Error GoTo ChkErr
    Call subGim
    Call subGil
'    gilCompCode.ListIndex = 1
'    gilCompCode.Query_Description
'    gilCompCode_Change
    gilCompCode.SetCodeNo garyGi(9)
    gilCompCode.Query_Description
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "Form_Load")
End Sub

Private Sub subGim()
  On Error GoTo ChkErr
    Call SetgiMulti(gimAreaCode, "CodeNo", "Description", "CD001", "行政區代碼", "行政區名稱")
    Call SetgiMulti(gimClassCode, "CodeNo", "Description", "CD004", "客戶類別代碼", "客戶類別名稱")
    Call SetgiMulti(gimStatusCode, "CodeNo", "Description", "CD035", "客戶狀態代碼", "客戶狀態名稱")
    Call SetgiMulti(gimStrtCode, "CodeNo", "Description", "CD017", "街道代碼", "街道名稱")
    Call SetgiMulti(gimCMName, "CodeNo", "Description", "CD031", "收費方式代碼", "收費方式名稱")
    Call SetgiMulti(gimServCode, "CodeNo", "Description", "CD002", "服務區代碼", "服務區名稱")
    SetMsQry gimCircuitNo, "SELECT ROWNUM,CIRCUITNO FROM (SELECT DISTINCT CIRCUITNO FROM SO016 WHERE CIRCUITNO IS NOT NULL)", "流水編號", "網路編號"
    gimStatusCode.SetQueryCode "1,2"
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

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
  ReleaseCOM frmSO5520A
End Sub

Private Sub gilCompCode_Change()
On Error GoTo ChkErr
    Call SelectServType(gilCompCode.GetCodeNo, gilServiceType)
    gilServiceType.ListIndex = 1
    
    GiMultiFilter gimServCode, , gilCompCode.GetCodeNo
    GiMultiFilter gimAreaCode, , gilCompCode.GetCodeNo
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

Private Sub gilServiceType_Change()
  On Error GoTo ChkErr
    GiMultiFilter gimClassCode, gilServiceType.GetCodeNo
    GiMultiFilter gimCMName, gilServiceType.GetCodeNo
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "gilCompCode_Validate")
End Sub

Private Sub optMduId_Click()
  On Error Resume Next
    frmType.Enabled = False
End Sub

Private Sub optStrtCode_Click()
  On Error Resume Next
    frmType.Enabled = True
End Sub
