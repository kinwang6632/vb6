VERSION 5.00
Object = "{3BB22E57-824A-11D3-BD70-0080C8F80BC4}#3.4#0"; "GiGridR.ocx"
Begin VB.Form ViewForm 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "瀏覽畫面"
   ClientHeight    =   5490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9195
   ClipControls    =   0   'False
   Icon            =   "ViewForm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   9195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton cmdPrint 
      Caption         =   "F5.列印"
      Default         =   -1  'True
      Height          =   375
      Left            =   6240
      TabIndex        =   2
      Top             =   5010
      Width           =   1365
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&X)"
      Height          =   375
      Left            =   7710
      TabIndex        =   1
      Top             =   5010
      Width           =   1365
   End
   Begin prjGiGridR.GiGridR ggrView 
      Height          =   4875
      Left            =   30
      TabIndex        =   0
      ToolTipText     =   "請以Double Click選取資料！"
      Top             =   30
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   8599
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "ViewForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rsTmp As New ADODB.Recordset
Private blnIsSo3220A As Boolean
Private blnIsSo07x As Boolean
Private blnIsSo3110A As Boolean
Private blnIsSo111xB As Boolean


Private Sub cmdCancel_Click()
    On Error Resume Next
        Unload Me
End Sub

Private Sub cmdPrint_Click()
    On Error Resume Next
        ReadyGoPrint
        If blnIsSo3220A Then
            PrintRpt GetPrinterName(5), "SO3220A.rpt", , "收費單給號地址調整一覽表 [SO3220A]", "SELECT * From tmp007", , , True, , , , GiPaperLandscape
        ElseIf blnIsSo07x Then
            PrintRpt GetPrinterName(5), "SO07xLog.rpt", , "收款異常表 [SO07xLog]", "SELECT * From So07xlog", "已有其他使用者同時收款或作廢", , True, "TMP0000.MDB", , , GiPaperLandscape
        ElseIf blnIsSo3110A Then
            PrintRpt GetPrinterName(5), "SO3110A.rpt", , "收費單未產生一覽表 [SO3110A]", "SELECT * From tmp009", "未產生", , True
        End If
End Sub

Private Sub Form_Activate()
    On Error Resume Next
        Screen.MousePointer = vbDefault
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo ChkErr
        If KeyCode = vbKeyF5 Then cmdPrint.Value = True
        If KeyCode = vbKeyEscape Then Unload Me
        If KeyCode = vbKeyReturn Then Unload Me
    Exit Sub
ChkErr:
    ErrSub Me.Name, "Form_KeyDown"
End Sub

Private Sub Form_Load()
    On Error GoTo ChkErr
        If blnIsSo3220A Then
            Call SetGridData2
        ElseIf blnIsSo07x Then
            Call SetGridData3
        ElseIf blnIsSo3110A Then
            Call SetGridData4
        ElseIf blnIsSo111xB Then
            Call SetGridData5
        Else
'            Me.Width = ggrView.Width + 210
'            Me.Height = ggrView.Height + 480
'            Me.Line (ggrView.Left - 25, ggrView.Top - 25)-(ggrView.Left + ggrView.Width + 25, ggrView.Top + ggrView.Height + 25), QBColor(6), BF
'            Me.Top = (frmSO3240A.Top + frmSO3240A.Height) / 2 - (Me.Height / 2)
'            Me.Left = (frmSO3240A.Left + frmSO3240A.Width) / 2 - (Me.Width / 2)
            Me.Height = ggrView.Height + 500
            Call SetGridData
        End If
    Exit Sub
ChkErr:
    ErrSub Me.Name, "Form_Load"
End Sub

Private Sub SetGridData()
'"客戶編號CustId, 客戶名稱CustName, 電話(1)Tel, 收費地址, 客戶狀態CustStatusCode,
'CustStatusName" (註: CustStatusCode不顯示)
    On Error GoTo ChkErr
        Dim mflds As New prjGiGridR.GiGridFlds
            With mflds
                .Add "CustId", , , , False, "客戶編號", vbLeftJustify
                .Add "CustName", , , , False, "  客戶名稱  ", vbLeftJustify
                .Add "Tel1", , , , False, "   電話(1)   ", vbLeftJustify
                .Add "ChargeAddress", , , , False, "                    收費地址                  ", vbLeftJustify
                .Add "CustStatusName", , , , False, "客戶狀態", vbLeftJustify
            End With
        ggrView.AllFields = mflds
        ggrView.SetHead
        If rsTmp.EOF = False Then rsTmp.MoveFirst
        Set ggrView.Recordset = rsTmp
        ggrView.Refresh
    Exit Sub
ChkErr:
    ErrSub Me.Name, "SetGridData"
End Sub

Private Sub SetGridData2()
'"客戶編號CustId, 客戶名稱CustName, 電話(1)Tel, 收費地址, 客戶狀態CustStatusCode,
'CustStatusName" (註: CustStatusCode不顯示)
    Me.Caption = "收費單給號地址調整一覽表"
    On Error GoTo ChkErr
        Dim mflds As New prjGiGridR.GiGridFlds
            With mflds
                .Add "BillNo", , , , False, "單據編號         ", vbLeftJustify
                .Add "Item", , , , False, "項次", vbLeftJustify
                .Add "CustId", , , , False, "客戶編號", vbLeftJustify
                .Add "OldAddrNo", , , , False, "舊地址編號", vbLeftJustify
                .Add "OldAddress", , , , False, "舊地址名稱" & Space(20), vbLeftJustify
                .Add "NewAddrNo", , , , False, "新地址編號", vbLeftJustify
                .Add "NewAddress", , , , False, "新地址名稱" & Space(20), vbLeftJustify
                .Add "MduId", , , , False, "大樓編號", vbLeftJustify
            End With
        ggrView.AllFields = mflds
        ggrView.SetHead
        If rsTmp.EOF = False Then rsTmp.MoveFirst
        Set ggrView.Recordset = rsTmp
        ggrView.Refresh
    Exit Sub
ChkErr:
    ErrSub Me.Name, "SetGridData2"
End Sub

Private Sub SetGridData3()
'"客戶編號CustId, 客戶名稱CustName, 電話(1)Tel, 收費地址, 客戶狀態CustStatusCode,
'CustStatusName" (註: CustStatusCode不顯示)
    Me.Caption = "收款異常表"
    On Error GoTo ChkErr
        Dim mflds As New prjGiGridR.GiGridFlds
            With mflds
                .Add "BillNo", , , , False, "單據編號             ", vbLeftJustify
                .Add "Item", , , , False, "項次", vbLeftJustify
                .Add "CustId", , , , False, "客戶編號", vbLeftJustify
                .Add "CustName", , , , False, "客戶姓名", vbLeftJustify
                .Add "RealDate", , , , False, "實收日期", vbLeftJustify
                .Add "RealPeriod", , , , False, "實收期數", vbLeftJustify
                .Add "RealAmt", , , , False, "實收金額", vbLeftJustify
                .Add "RealStartDate", giControlTypeDate, , , False, "起始日期", vbLeftJustify
                .Add "RealStopDate", giControlTypeDate, , , False, "截止日期", vbLeftJustify
                .Add "UpdEn", , , , False, "異動人員", vbLeftJustify
                .Add "UpdTime", , , , False, "異動時間          ", vbLeftJustify
                .Add "Note", , , , False, "備註                              ", vbLeftJustify
            End With
        ggrView.AllFields = mflds
        ggrView.SetHead
        If rsTmp.EOF = False Then rsTmp.MoveFirst
        Set ggrView.Recordset = rsTmp
        ggrView.Refresh
    Exit Sub
ChkErr:
    ErrSub Me.Name, "SetGridData3"
End Sub

Private Sub SetGridData4()
'"客戶編號CustId, 客戶名稱CustName, 電話(1)Tel, 收費地址, 客戶狀態CustStatusCode,
'CustStatusName" (註: CustStatusCode不顯示)
    Me.Caption = "收費單未產生一覽表"
    On Error GoTo ChkErr
        Dim mflds As New prjGiGridR.GiGridFlds
            With mflds
                .Add "ServiceType", , , , False, "服務別", vbLeftJustify
                .Add "BillNo", , , , False, "單據編號             ", vbLeftJustify
                .Add "Item", , , , False, "項次", vbLeftJustify
                .Add "FaciSNo", , , , False, "設備序號       ", vbLeftJustify
                .Add "CitemName", , , , False, "收費項目名稱" & Space(10), vbLeftJustify
                .Add "CustId", , , , False, "客戶編號", vbLeftJustify
                .Add "ShouldDate", , , , False, "應收日期", vbLeftJustify
                .Add "ShouldAmt", , , , False, "應收金額", vbLeftJustify
                .Add "StartDate", giControlTypeDate, , , False, "起始日期", vbLeftJustify
                .Add "StopDate", giControlTypeDate, , , False, "截止日期", vbLeftJustify
            End With
        ggrView.AllFields = mflds
        ggrView.SetHead
        If rsTmp.EOF = False Then rsTmp.MoveFirst
        Set ggrView.Recordset = rsTmp
        ggrView.Refresh
    Exit Sub
ChkErr:
    ErrSub Me.Name, "SetGridData4"
End Sub

Private Sub SetGridData5()
'"客戶編號CustId, 客戶名稱CustName, 電話(1)Tel, 收費地址, 客戶狀態CustStatusCode,
'CustStatusName" (註: CustStatusCode不顯示)
    Me.Caption = "收費單分期相關資料"
    On Error GoTo ChkErr
        Dim mflds As New prjGiGridR.GiGridFlds
            With mflds
                .Add "BillNo", , , , False, "單據編號             ", vbLeftJustify
                .Add "CustId", , , , False, "客戶編號", vbLeftJustify
                .Add "ShouldDate", , , , False, "應收日期", vbLeftJustify
                .Add "ShouldAmt", , , , False, "應收金額", vbLeftJustify
                .Add "StartDate", giControlTypeDate, , , False, "起始日期", vbLeftJustify
                .Add "StopDate", giControlTypeDate, , , False, "截止日期", vbLeftJustify
            End With
        ggrView.AllFields = mflds
        ggrView.SetHead
        If rsTmp.EOF = False Then rsTmp.MoveFirst
        Set ggrView.Recordset = rsTmp
        ggrView.Refresh
    Exit Sub
ChkErr:
    ErrSub Me.Name, "SetGridData5"
End Sub

Public Property Let uRecordset(ByVal vRcdSet As ADODB.Recordset)
    On Error GoTo ChkErr
        Set rsTmp = vRcdSet
    Exit Property
ChkErr:
    ErrSub Me.Name, "Let uRecordset"
End Property

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo ChkErr
        frmSO3240A.dblBookMark = rsTmp.Bookmark
        blnIsSo3220A = False
        blnIsSo07x = False
        blnIsSo3110A = False
        rsTmp.Close
        Set rsTmp = Nothing
    Exit Sub
ChkErr:
    ErrSub Me.Name, "Form_QueryUnload"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
        Call ReleaseCOM(Me)

End Sub

Private Sub ggrView_DblClick()
    On Error GoTo ChkErr
        Unload Me
    Exit Sub
ChkErr:
    ErrSub Me.Name, "ggrView_DblClick"
End Sub


Public Property Let uIsSo3220A(ByVal vData As Boolean)
    On Error GoTo ChkErr
        blnIsSo3220A = vData
    Exit Property
ChkErr:
    ErrSub Me.Name, "Let uIsSo3220A"
End Property

Public Property Let uIsSo07x(ByVal vData As Boolean)
    On Error GoTo ChkErr
        blnIsSo07x = vData
    Exit Property
ChkErr:
    ErrSub Me.Name, "Let uIsSO07x"
End Property

Public Property Let uIsSo3110A(ByVal vData As Boolean)
    On Error GoTo ChkErr
        blnIsSo3110A = vData
    Exit Property
ChkErr:
    ErrSub Me.Name, "Let uIsSo3110A"
End Property

Public Property Let uIsSo111xB(ByVal vData As Boolean)
    On Error GoTo ChkErr
        blnIsSo111xB = vData
    Exit Property
ChkErr:
    ErrSub Me.Name, "Let uIsSo111xB"
End Property
