VERSION 5.00
Object = "{3BB22E57-824A-11D3-BD70-0080C8F80BC4}#3.4#0"; "GiGridR.ocx"
Begin VB.Form frmSO3311C 
   BorderStyle     =   1  '單線固定
   Caption         =   "收費資料已登瀏覽 [SO3311C]"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   2235
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "新細明體"
      Size            =   9.75
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SO3311C.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton cmdchkPrint 
      Caption         =   "查看異常資料"
      Height          =   345
      Left            =   7170
      TabIndex        =   5
      Top             =   5280
      Width           =   1395
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "刪除暫存資料"
      Height          =   345
      Left            =   5310
      TabIndex        =   4
      Top             =   5280
      Width           =   1395
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "F5.列印"
      Default         =   -1  'True
      Height          =   345
      Left            =   8580
      TabIndex        =   3
      Top             =   5280
      Width           =   1365
   End
   Begin prjGiGridR.GiGridR ggrData 
      Height          =   4725
      Left            =   210
      TabIndex        =   2
      Top             =   360
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   8334
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&X)"
      Height          =   345
      Left            =   10050
      TabIndex        =   1
      Top             =   5280
      Width           =   1365
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "(單據編號,客戶編號,姓名,收費項目,出單金額,應收日期,實收金額,實收日期,期數,期限起始,期限截止,作廢,登錄人員)"
      Height          =   195
      Left            =   180
      TabIndex        =   0
      Top             =   60
      Visible         =   0   'False
      Width           =   9630
   End
End
Attribute VB_Name = "frmSO3311C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rsSo3311C As New ADODB.Recordset
Dim strView As String
Dim objParentForm As Form


Private Sub cmdCancel_Click()
    On Error GoTo chkErr
        Unload Me
    Exit Sub
chkErr:
    ErrSub Me.Name, "cmdCancel_Click"
End Sub

Private Sub ConnSql()
    On Error GoTo chkErr
    Dim strSQL As String
        strSQL = "Select A.EntryNO,A.BillNO,A.PrtSno,A.MediaBillNo,A.CustId,A.CustName,B.CustStatusName,A.CitemName,A.ShouldAmt,A.ShouldDate,A.RealAmt,A.RealDate,A.RealPeriod,A.RealStartDate" & _
                    ",A.RealStopDate,A.CancelFlag,A.EntryEn,A.RowId RowNo,A.ManualNo From " & GetOwner & "So074 A, " & GetOwner & "So002 B Where A.CustId= B.CustId And A.ServiceType = B.ServiceType And EntryEn='" & garyGi(0) & "' order by  EntryNO Desc "
        If Not GetRS(rsSo3311C, strSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then MsgBox "連線失敗！請重新操作！"
        Set ggrData.Recordset = rsSo3311C
        If rsSo3311C.EOF Or rsSo3311C.BOF Then cmdPrint.Enabled = False
        ggrData.Refresh
    Exit Sub
chkErr:
    ErrSub Me.Name, "ConnSql"
End Sub

Private Sub cmdchkPrint_Click()
    On Error GoTo chkErr
    Dim clsAlterCharge As Object
    Dim strSQL As String
    Dim rsSO033 As New ADODB.Recordset
        strSQL = "Select A.*,'" & garyGi(1) & "' UpdEn,'" & GetDTString(RightNow) & "' UpdTime From " & GetOwner & "So074 A," & GetOwner & "So033  B Where A.RcdRowId=B.RowId And B.UCCode Is Null And B.RealDate Is Not Null"
        If Not GetRS(rsSO033, strSQL) Then Exit Sub
        If rsSO033.RecordCount = 0 Then
            MsgNoRcd
        Else
            Set clsAlterCharge = CreateObject("csAlterCharge4.clsAlterCharge")
            With clsAlterCharge
                Set cnn = GetTmpMdbCn
                .uOwnerName = GetOwner
                Set .ucnMDB = cnn
                If Not .InsertChkErrLog(gcnGi, rsSO033) Then Exit Sub
            End With
            Call PrintLogData
        End If
    Exit Sub
chkErr:
    ErrSub Me.Name, "cmdchkPrint_Click"
End Sub

Private Sub PrintLogData()
    On Error GoTo chkErr
    Dim rsTmp As New ADODB.Recordset
        If Not GetRS(rsTmp, "Select * From SO07xLog", cnn) Then Exit Sub
        If rsTmp.RecordCount > 0 Then
            With ViewForm
                .uIsSo07x = True
                .uRecordset = rsTmp
                .Show vbModal
            End With
        End If
    Exit Sub
chkErr:
    ErrSub Me.Name, "PrintLogData"
End Sub

Private Sub cmdDelete_Click()
    On Error GoTo chkErr
    Dim lngRealAmt As Long
        With rsSo3311C
            If .BOF Or .EOF Then Exit Sub
            lngRealAmt = .Fields("RealAmt")
            gcnGi.BeginTrans
            gcnGi.Execute "Delete From " & GetOwner & "So074 Where RowId='" & .Fields("RowNo") & "'"
            gcnGi.CommitTrans
            Call ConnSql
        End With
        With objParentForm
            .lblBillCnt = Val(.lblBillCnt) - 1
            .lblTotalAmt = .lblTotalAmt - lngRealAmt
        End With
        MsgBox "刪除成功!!", vbExclamation, gimsgPrompt
    Exit Sub
chkErr:
    ErrSub Me.Name, "cmdDelete_Click"
End Sub

Private Sub cmdPrint_Click()
    On Error Resume Next
        gcnGi.Execute "Drop View " & strView
        strView = GetTmpViewName
        gcnGi.Execute "Create View " & GetOwner & strView & " As " & rsSo3311C.Source
        Call PrintRpt(GetPrinterName(5), "SO3311C.rpt", , "收費資料已登瀏覽  [SO3311C]", "Select * From " & strView & " A Order by EntryEn", "全部", , True, , False, , GiPaperLandscape)
End Sub

Private Sub Form_Activate()
    On Error Resume Next
        Screen.MousePointer = vbDefault

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
        If Not ChkGiList(KeyCode) Then Exit Sub
        If KeyCode = vbKeyF5 Then If cmdPrint.Enabled Then cmdPrint.Value = True
        If KeyCode = vbKeyEscape Then Unload Me: Exit Sub
End Sub

Private Sub Form_Load()
    On Error Resume Next
        Call SetGrid
        Call ConnSql
        Screen.MousePointer = vbDefault
End Sub

Private Sub SetGrid()
    On Error GoTo chkErr
    Dim mFlds As New prjGiGridR.GiGridFlds
    '(單據編號,客戶編號,姓名,收費項目,出單金額,應收日期,實收金額,實收日期,期數,期限起始,期限截止,作廢,登錄人員)
        With mFlds
            .Add "EntryNO", , , , False, "單據序號", vbLeftJustify
            .Add "BillNO", , , , False, "單據編號           ", vbLeftJustify
            .Add "PrtSNo", , , , False, "印單序號     ", vbLeftJustify
            .Add "MediaBillNo", , , , False, "媒體單號    ", vbLeftJustify
            .Add "Custid", , , , False, "客戶編號", vbLeftJustify
            .Add "CustName", , , , False, "客戶姓名        ", vbLeftJustify
            '901031 Jacky
            .Add "CustStatusName", , , , False, "客戶狀態", vbLeftJustify
            '901031 Jacky
            .Add "CitemName", , , , False, "收費項目        ", vbLeftJustify
            .Add "ShouldAmt", , , , False, "出單金額", vbLeftJustify
            .Add "ShouldDate", giControlTypeDate, , , False, "   應收日期   ", vbLeftJustify
            .Add "RealAmt", , , , False, "實收金額", vbLeftJustify
            .Add "RealDate", giControlTypeDate, , , False, "   實收日期   ", vbLeftJustify
            .Add "RealPeriod", , , , False, "期數", vbLeftJustify
            .Add "RealStartDate", giControlTypeDate, , , False, "   期限起始   ", vbLeftJustify
            .Add "RealStopDate", giControlTypeDate, , , False, "   期限截止   ", vbLeftJustify
            .Add "ManualNo", , , , False, "手開單號", vbLeftJustify
            .Add "CancelFlag", , , , False, "作廢", vbLeftJustify
            .Add "EntryEn", , , , False, "登錄人員", vbLeftJustify
        End With
        ggrData.AllFields = mFlds
        ggrData.SetHead
    Exit Sub
chkErr:
    ErrSub Me.Name, "SetGrid"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
        gcnGi.Execute "Drop View " & strView

End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
        Call ReleaseCOM(Me)

End Sub

Private Sub ggrData_ShowCellData(ByVal Col As Long, ByVal Row As Long, ByVal giFld As prjGiGridR.GiGridFld, ByVal Fld As ADODB.Field, Value As Variant)
    On Error GoTo chkErr
'        If InStr(1, UCase(Fld.Name), "DATE") > 0 Then
'            Value = Format(Value, "ee/MM/DD")
'        End If
        If UCase(Fld.Name) = "CANCELFLAG" Then
            Value = IIf(Value = 0, "", "是")
        End If
    Exit Sub
chkErr:
    ErrSub Me.Name, "ggrData_ShowCellData"
End Sub

Public Property Set uParentForm(ByVal vData As Form)
    Set objParentForm = vData
End Property
