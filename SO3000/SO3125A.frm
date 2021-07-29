VERSION 5.00
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{A0C51418-555F-11D3-BC10-0050BA01A3A9}#9.8#0"; "GiMulti.ocx"
Object = "{074771B7-6ADE-11D5-B178-00485456B028}#1.16#0"; "csMulti.ocx"
Begin VB.Form frmSO3125A 
   Caption         =   "出單數及金額查詢 [SO3125A]"
   ClientHeight    =   3630
   ClientLeft      =   1950
   ClientTop       =   2745
   ClientWidth     =   6825
   BeginProperty Font 
      Name            =   "新細明體"
      Size            =   9.75
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SO3125A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form9"
   LockControls    =   -1  'True
   ScaleHeight     =   3630
   ScaleWidth      =   6825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton cmdPrint 
      Caption         =   "F5:列印"
      Height          =   465
      Left            =   180
      TabIndex        =   7
      Top             =   3120
      Width           =   1275
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      Height          =   465
      Left            =   5280
      TabIndex        =   6
      Top             =   3120
      Width           =   1275
   End
   Begin VB.CommandButton cmdPrevRpt 
      Caption         =   "上次統計結果"
      Height          =   465
      Left            =   2640
      TabIndex        =   5
      Top             =   3120
      Width           =   1425
   End
   Begin Gi_Multi.GiMulti gimServCode 
      Height          =   345
      Left            =   30
      TabIndex        =   3
      Top             =   1950
      Width           =   6765
      _ExtentX        =   11933
      _ExtentY        =   609
      ButtonCaption   =   "服    務    區"
   End
   Begin Gi_Multi.GiMulti gimAreaCode 
      Height          =   345
      Left            =   30
      TabIndex        =   2
      Top             =   1580
      Width           =   6765
      _ExtentX        =   11933
      _ExtentY        =   609
      ButtonCaption   =   "行    政    區"
   End
   Begin CS_Multi.CSmulti gimClctEn 
      Height          =   345
      Left            =   30
      TabIndex        =   0
      Top             =   870
      Width           =   6765
      _ExtentX        =   11933
      _ExtentY        =   609
      ButtonCaption   =   "收 費 人 員"
   End
   Begin prjGiList.GiList gilCompCode 
      Height          =   315
      Left            =   1200
      TabIndex        =   8
      Top             =   60
      Width           =   3000
      _ExtentX        =   5292
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
      F2Corresponding =   -1  'True
   End
   Begin prjGiList.GiList gilServiceType 
      Height          =   315
      Left            =   1200
      TabIndex        =   9
      Top             =   495
      Width           =   3000
      _ExtentX        =   5292
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
      F2Corresponding =   -1  'True
   End
   Begin CS_Multi.CSmulti gmdOldClctEn 
      Height          =   345
      Left            =   30
      TabIndex        =   1
      Top             =   1230
      Width           =   6765
      _ExtentX        =   11933
      _ExtentY        =   609
      ButtonCaption   =   "原收費人員"
   End
   Begin VB.Label lblServiceType 
      AutoSize        =   -1  'True
      Caption         =   "服務類別"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   570
      Width           =   780
   End
   Begin VB.Label lblCompCode 
      AutoSize        =   -1  'True
      Caption         =   "公司別"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   150
      Width           =   585
   End
   Begin VB.Label lbl1 
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
      TabIndex        =   4
      Top             =   2400
      Visible         =   0   'False
      Width           =   510
   End
End
Attribute VB_Name = "frmSO3125A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As New ADODB.Connection
Private strChoose As String
Private strChooseString As String
Private strForm As String

Private Sub cmdPrevRpt_Click()
  On Error GoTo ChkErr
   Screen.MousePointer = vbHourglass
   Call PreviousRpt(GetPrinterName(5), "SO3125A.Rpt", Me.Caption)
   Screen.MousePointer = vbDefault
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "cmdPrevRpt_Click")
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    On Error GoTo ChkErr
        Set cnn = GetTmpMdbCn
        Call subGim
        Call subGil
        gilCompCode.SetCodeNo garyGi(9)
        gilCompCode.Query_Description
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "Form_Load")
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error GoTo ChkErr
    Call FunctionKey(KeyCode, Shift)
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "FunctionKey")
End Sub

Public Sub FunctionKey(KeyCode As Integer, Shift As Integer)
  On Error GoTo ChkErr
    Select Case KeyCode
           '----------------------------------------------------
           Case vbKeyF5 '   F5: 列印 相當於按下cmdPrint
                        If cmdPrint.Enabled = True Then cmdPrint.Value = True
           Case vbKeyEscape
                        If cmdExit.Enabled = True Then cmdExit.Value = True
           '----------------------------------------------------
    End Select
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "FunctionKey")
End Sub

Private Sub cmdExit_Click()
  On Error GoTo ChkErr
   Unload Me
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "cmdExit_Click")
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo ChkErr
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, lngAffected As Long
        Screen.MousePointer = vbHourglass
        cmdExit.SetFocus
        ReadyGoPrint
        Call subChoose
        If Not subInsertMDB(lngAffected) Then Exit Sub
        If lngAffected = 0 Then
            MsgNoRcd
        Else
            strSQL = "SELECT * FROM SO3125A"
            Call PrintRpt(GetPrinterName(5), "SO3125A.rpt", "SO3125", Me.Caption, strSQL, strChooseString, , True, "Tmp0000.MDB")
        End If
        Screen.MousePointer = vbDefault
        Set rsTmp = Nothing
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "cmdPrint_Click")
End Sub

Private Function subInsertMDB(ByRef lngAffected As Long) As Boolean
    On Error GoTo ChkErr
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
        cnn.BeginTrans
        strSQL = "SELECT CitemCode,CitemName,RealPeriod,ShouldAmt,Count(*) as intCount,Sum(ShouldAmt) SumAmt " & _
              "From " & GetOwner & "SO032 " & strChoose & " Group By CitemCode,CitemName,RealPeriod,ShouldAmt"
        SendSQL strSQL, True
        If Not GetRS(rsTmp, strSQL) Then Exit Function
        If Not CreateMDBTable(rsTmp, "SO3125A", cnn) Then Exit Function
        Set rsTmp = gcnGi.Execute(strSQL)
        If Not rsTmp.EOF Then lngAffected = 1
        While Not rsTmp.EOF
            cnn.Execute GetTmpMDBExecuteStr(rsTmp, "SO3125A")
            rsTmp.MoveNext
            DoEvents
        Wend
        cnn.CommitTrans
        Call Close3Recordset(rsTmp)
        subInsertMDB = True
    Exit Function
ChkErr:
    cnn.RollbackTrans
    Call ErrSub(Me.Name, "subInsertMdb")
End Function

Private Sub subChoose()
  '串 SelectCondition
  On Error GoTo ChkErr
   strChoose = " ClctYM Is Not Null "
   '#1659  950517原"收費人員"改為"原收費人員"，另新增一"收費人員"
    If gimClctEn.GetQryStr <> "" Then Call subAnd("OldClctEn " & gimClctEn.GetQryStr)
    If gmdOldClctEn.GetQryStr <> "" Then Call subAnd("ClctEn " & gmdOldClctEn.GetQryStr)
    If gimAreaCode.GetQryStr <> "" Then Call subAnd("AreaCode " & gimAreaCode.GetQryStr)
    If gimServCode.GetQryStr <> "" Then Call subAnd("ServCode " & gimServCode.GetQryStr)
    If gilCompCode.GetCodeNo <> "" Then Call subAnd2(strChoose, "CompCode =" & gilCompCode.GetCodeNo)
    If gilServiceType.GetCodeNo <> "" Then Call subAnd2(strChoose, "ServiceType ='" & gilServiceType.GetCodeNo & "'")
    
    If Len(strChoose) > 0 Then
       strChoose = " WHERE " & strChoose
    End If
    strChooseString = "收費人員: " & subSpace(gimClctEn.GetDispStr) & ";" & _
                      "原收費人員: " & subSpace(gmdOldClctEn.GetDispStr) & ";" & _
                      "行政區　: " & subSpace(gimAreaCode.GetDispStr) & ";" & _
                      "服務區　: " & subSpace(gimServCode.GetDispStr)
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "subChoose")
End Sub

Private Sub subAnd(strCH As String)
  On Error GoTo ChkErr
    If strChoose <> "" Then
       strChoose = strChoose & " And " & strCH
    Else
       strChoose = strCH
    End If
    
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "subAnd")
End Sub

Private Sub subGil()
    On Error Resume Next
        SetgiList gilServiceType, "CodeNo", "Description", "CD046", , , , , 1, 12
        SetgiList gilCompCode, "CodeNo", "Description", "CD039", , , , , 3, 12, GetCompCodeFilter(gilCompCode)

End Sub

Private Sub subGim()
  On Error GoTo ChkErr
    Call SetgiMulti(gimClctEn, "EmpNo", "EmpName", "CM003", "收費員代碼", "收費員名稱", , True)
    Call SetgiMulti(gmdOldClctEn, "EmpNo", "EmpName", "CM003", "收費員代碼", "收費員名稱", , True)
    Call SetgiMulti(gimAreaCode, "CodeNo", "Description", "CD001", "行政區代碼", "行政區名稱", , True)
    Call SetgiMulti(gimServCode, "CodeNo", "Description", "CD002", "服務區代碼", "服務區名稱", , True)
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "subGim")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
        Set cnn = Nothing
End Sub

Public Property Get uForm() As String
    On Error GoTo ChkErr
        uForm = strForm
    Exit Property
ChkErr:
    ErrSub Me.Name, "Get uForm"
End Property

Public Property Let uForm(ByVal vForm As String)
    On Error GoTo ChkErr
        strForm = vForm
    Exit Property
ChkErr:
    ErrSub Me.Name, "Let uForm"
End Property

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
        Call ReleaseCOM(Me)
End Sub

Private Sub gilCompCode_Change()
    On Error Resume Next
        If Not ChgComp(gilCompCode, "SO3120", "SO3125") Then Exit Sub
        Call subGil
        Call subGim
        Call SelectServType(gilCompCode.GetCodeNo, gilServiceType)
        gilServiceType.SetCodeNo ""
        gilServiceType.Query_Description
        gilServiceType.ListIndex = 1
        Call GiMultiFilter(gimClctEn, , gilCompCode.GetCodeNo)
        Call GiMultiFilter(gimAreaCode, , gilCompCode.GetCodeNo)
        Call GiMultiFilter(gimServCode, , gilCompCode.GetCodeNo)
        
End Sub
