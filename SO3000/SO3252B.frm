VERSION 5.00
Object = "{3BB22E57-824A-11D3-BD70-0080C8F80BC4}#3.4#0"; "GiGridR.ocx"
Begin VB.Form frmSO3252B 
   BorderStyle     =   1  '單線固定
   Caption         =   "資料瀏覽"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11730
   Icon            =   "SO3252B.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   11730
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&X)"
      Height          =   405
      Left            =   7740
      TabIndex        =   2
      Top             =   5910
      Width           =   1425
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "F2.確定"
      Height          =   405
      Left            =   2400
      TabIndex        =   1
      Top             =   5910
      Width           =   1425
   End
   Begin prjGiGridR.GiGridR ggrSelectData 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   11475
      _ExtentX        =   20241
      _ExtentY        =   9763
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseCellForeColor=   -1  'True
   End
End
Attribute VB_Name = "frmSO3252B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'根據#3042 需求 增加此Form，按下確定按鈕後開始執行作廢 By Kin 2007/03/22
Option Explicit
Private strCompCode As String
Private strSQLWhere As String
Private strReasno As String
Private strCancelCode As String
Private strCancelName As String
Private rsUpdate As ADODB.Recordset
Private strUpdTime As String
Private Sub cmdCancel_Click()
    On Error Resume Next
    Unload Me
End Sub
Private Sub cmdOk_Click()
  On Error GoTo ChkErr
    Dim lngAffectCounts As Long
    Dim lngUpdateCount As Long
    Dim lngErrCount As Long
    gcnGi.BeginTrans
        '更新資料
        Screen.MousePointer = vbHourglass
        If ExecuteSQL("Update " & GetOwner & "So033 Set UCCode=Null,UCName=Null,CancelFlag=1,RealDate=To_Date('" & Format(RightDate, "YYYY/MM/DD") & "','YYYY/MM/DD')," & _
                      "Note = Trim(Nvl(Note,'')) || ';作廢原因:" & strReasno & "; 未收原因:' || UCName,CancelCode = " & strCancelCode & "," & _
                      "CancelName = '" & strCancelName & "',UpdEn = '" & garyGi(1) & "',UpdTime ='" & strUpdTime & "'  Where " & strSQLWhere, gcnGi, lngAffectCounts, False, False) <> giOK Then
            MsgBox "連線失敗！請重新操作！", vbExclamation, "訊息！": GoTo ChkErr
        End If
        strSQLWhere = SetAddComma(strSQLWhere)
        If ExecuteSQL("Insert into " & GetOwner & "So070 (UpdTime,UpdName,Reason,Para,RcdCount) values('" & strUpdTime & "','" & garyGi(1) & "'," & IIf(Trim(strCancelName) <> "", "'" & Trim(strCancelName) & "'", Null) & ",'" & Replace(strSQLWhere, "'", "") & "'," & IIf(lngUpdateCount < 0, 0, lngUpdateCount) & ")", gcnGi, , , False) <> giOK Then
            MsgBox "連線失敗！請重新操作！", vbExclamation, "訊息！"
        End If
    gcnGi.CommitTrans
    MsgBox "作廢完畢，筆數=" & IIf(lngAffectCounts < 0, 0, lngAffectCounts), vbInformation, "訊息！"
    Screen.MousePointer = vbDefault
    Unload Me
    Exit Sub
ChkErr:
    gcnGi.RollbackTrans
    ErrSub Me.Name, "cmdOk_Click"
End Sub

Private Sub Form_Activate()
On Error Resume Next
        Screen.MousePointer = vbDefault
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
        If Not ChkGiList(KeyCode) Then Exit Sub
        If KeyCode = vbKeyF2 And cmdOK.Enabled Then Call cmdOk_Click: Exit Sub
End Sub

Private Sub Form_Load()
  On Error GoTo ChkErr
    Call SetGrd '設定Grid的Caption
    Exit Sub
ChkErr:
    ErrSub Me.Name, "Form_Load"
End Sub

'公司別
Public Property Let uCompCode(ByVal vData As String)
    strCompCode = vData
End Property
'畫面所串出來的條件
Public Property Let uSQLWhere(ByVal vData As String)
    strSQLWhere = vData
End Property
'前端查詢出來的資料
Public Property Let uRS(ByRef vData As ADODB.Recordset)
    Set rsUpdate = vData
End Property
'備註
Public Property Let uReasno(ByVal vData As String)
    strReasno = vData
End Property
'作廢原因代碼
Public Property Let uCancelCode(ByVal vData As String)
    strCancelCode = vData
End Property
'作廢原因
Public Property Let uCancelName(ByVal vData As String)
    strCancelName = vData
End Property
'更新時間
Public Property Let uUpdTime(ByVal vData As String)
    strUpdTime = vData
End Property

Private Sub SetGrd()
   On Error GoTo ChkErr
   Dim mflds As New prjGiGridR.GiGridFlds
      With mflds
         '.Add "Chose", , , , False, "選擇", vbCenter
'         .Add "CustId", , , , False, "客戶編號", vbLeftJustify
         .Add "BillNo", , , , False, "       單據編號     ", vbLeftJustify
         .Add "CITEMNAME", , , , False, "     收費項目名稱    ", vbLeftJustify
         .Add "ShouldAmt", , , , False, "  出單金額  ", vbRightJustify
         .Add "RealPeriod", giControlTypeComboBox, , , False, " 期數 ", vbRightJustify
         .Add "ShouldDate", giControlTypeDate, , , False, "  應收日期  ", vbLeftJustify
         .Add "RealStartDate", giControlTypeDate, , , False, "   起始日   ", vbLeftJustify
         .Add "RealStopDate", giControlTypeDate, , , False, "   截止日   ", vbLeftJustify
         .Add "UCName", , , , False, "             未收原因                ", vbLeftJustify
'         .Add "GUINO", , , , False, "    發票號碼    "
      End With
      ggrSelectData.AllFields = mflds
      ggrSelectData.SetHead
      ggrSelectData.Enabled = True
      Set ggrSelectData.Recordset = rsUpdate
      ggrSelectData.Blank = False
      ggrSelectData.GotoRow 1
      ggrSelectData.Refresh
    Exit Sub
ChkErr:
   Call ErrSub(Me.Name, "subGrd4")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    CloseRecordset rsUpdate
End Sub
Private Function SetAddComma(ByVal strSQL As String) As String
    On Error GoTo ChkErr
        Dim strNewSQl As String
        Dim strTmpSql  As String
        Dim intSitus As Integer
        Do While True
            intSitus = InStr(1, strSQL, "'", vbTextCompare)
            If intSitus <= 0 Then Exit Do
            strTmpSql = strTmpSql & Left(strSQL, intSitus) & "'"
            strSQL = Mid(strSQL, intSitus + 1)
        Loop
        SetAddComma = strTmpSql
    Exit Function
ChkErr:
    ErrSub Me.Name, "SetAddComma"
End Function

