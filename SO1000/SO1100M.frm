VERSION 5.00
Object = "{3BB22E57-824A-11D3-BD70-0080C8F80BC4}#3.4#0"; "GiGridR.ocx"
Begin VB.Form frmSO1100M 
   BorderStyle     =   1  '單線固定
   Caption         =   "分期明細資料 [SO1100M]"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   1530
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
   Icon            =   "SO1100M.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   Begin prjGiGridR.GiGridR ggrData 
      Height          =   4080
      Left            =   -15
      TabIndex        =   0
      ToolTipText     =   "請按左鍵兩次,選取客戶資料"
      Top             =   0
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   7197
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
End
Attribute VB_Name = "frmSO1100M"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private rsInstallment As New ADODB.Recordset

'設備分期付款查詢機制
'SO1100B之設備頁籤加一[分期明細資料]Button, 當User點選設備Gird中之STB設備後
'再Click該Button時, 即Show該STB分期之資料明細於一New Form的Gird中,
'其Gird的格式內容如收費資料頁籤之Gird內容.但取資料部分
'需加過濾條件SO033 or SO034之 FaciSeqNo=SO004.SeqNo and
'SNO=SO004.SNO and Budget = 1(因為我們只要抓該STB的分期資料顯示)
'Default以SNO, ShouldDate做Order By

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error Resume Next
    If KeyCode = 27 Then Unload Me
End Sub

Public Sub QuerySO033SO034(ByVal FaciSeqNo As String, _
                            ByVal SNO As String)
  On Error GoTo ChkErr
    Dim strQrySO033SO034 As String
    Dim strCDT As String
    Dim strField As String
    strField = ChargeField
    
    strCDT = "CUSTID=" & gCustId & _
                " AND SERVICETYPE='" & gServiceType & "'" & _
                " AND COMPCODE=" & gCompCode & _
                " AND Budget = 1" & _
                " AND FaciSeqNo='" & FaciSeqNo & "'" & _
                " AND SNO='" & SNO & "'"

    strQrySO033SO034 = "SELECT 0 AS TYPE,ROWID," & strField & _
                        " FROM " & GetOwner & "SO033 SO033 WHERE UCCODE IS NOT NULL AND " & _
                         strCDT & _
                        " UNION  ALL " & _
                        "SELECT 1 AS TYPE,ROWID," & strField & _
                        " FROM " & GetOwner & "SO033 SO033 WHERE UCCODE IS NULL AND " & _
                         strCDT & _
                        " UNION ALL " & _
                        "SELECT 2 AS TYPE,ROWID," & strField & _
                        " FROM " & GetOwner & "SO034 SO034 WHERE " & _
                         strCDT & _
                        " ORDER BY SNO,SHOULDDATE"
    GetRS rsInstallment, strQrySO033SO034, gcnGi, adUseClient, adOpenForwardOnly, adLockPessimistic, "(SO033 UNION SO034)", "QuerySO033SO034"
    GrdFmt
  Exit Sub
ChkErr:
    ErrSub Me.Name, "QuerySO033SO034"
End Sub


Private Sub Form_Load()
  On Error GoTo ChkErr
    If Crm Then
        PolyFrmFunction Me
        Me.Move frmSO1100A.Left + ((frmSO1100A.Width - Me.Width) / 2), frmSO1100A.Top + ((frmSO1100A.Height - Me.Height) / 2)
    Else
'        If Not800600 Then Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2 - 200
        If Not800600 Then Me.Move (Screen.Width - Me.Width) / 2, ((Screen.Height - Me.Height) / 2) + 1800
    End If
  Exit Sub
ChkErr:
    ErrSub Me.Name, "Form_Load"
End Sub

Public Sub GrdFmt()
  On Error GoTo ChkErr
    Dim mFlds As New GiGridFlds
    With mFlds
        .Add "Type", , , , , " 種類  ", vbLeftJustify
        .Add "BillNo", , , , , "      單據編號    ", vbLeftJustify
        .Add "CitemName", , , , , "收費項目  ", vbLeftJustify
        .Add "ShouldAmt", , , , , "出單$", vbRightJustify
        .Add "RealAmt", , , , , "實收$", vbRightJustify
        .Add "Realperiod", , , , , " 期 ", vbRightJustify
        .Add "ShouldDate", giControlTypeDate, , , , "應收日期", vbLeftJustify
        .Add "CancelFlag", , , , , "作廢", vbLeftJustify
        .Add "RealDate", giControlTypeDate, , , , "實收日期", vbLeftJustify
        .Add "RealStartDate", giControlTypeDate, , , , "起始日", vbLeftJustify
        .Add "RealStopDate", giControlTypeDate, , , , "截止日", vbLeftJustify
        .Add "UCName", , , , , "未收原因", vbLeftJustify
        .Add "STName", , , , , " 短收原因", vbLeftJustify
        .Add "ClctName", , , , , "收費人員", vbLeftJustify
        .Add "COMPCODE", , , , , "公司別", vbLeftJustify
        .Add "CMName", , , , , "收費方式", vbLeftJustify
        .Add "CitibankATM", , , , , "虛擬帳號    " & Space(6), vbLeftJustify
        .Add "GUIno", , , , , "發票號碼  ", vbLeftJustify
        .Add "PreInvoice", , , , , "是否預開發票", vbLeftJustify
        .Add "CancelName", , , , , "作廢原因", vbLeftJustify & Space(2)
        .Add "Note", , , , , "備註" & Space(30), , , vbLeftJustify
        .Add "OldAmt", , , , , "原應收金額", vbLeftJustify
        .Add "PrtSNo", , , , , "印單序號" & Space(12), vbLeftJustify
        .Add "UpdTime", , , , , "異動時間         ", vbLeftJustify
        .Add "UpdEn", , , , , "異動人員", vbLeftJustify
    End With
    With ggrData
        .AllFields = mFlds
        .SetHead
         Set .Recordset = rsInstallment
        .Refresh
    End With
    Me.Show 1
  Exit Sub
ChkErr:
    ErrSub Me.Name, "OpenData"
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
    rsInstallment.Close
    Set rsInstallment = Nothing
    ReleaseCOM Me
End Sub

Private Sub ggrData_ShowCellData(ByVal Col As Long, ByVal Row As Long, ByVal giFld As prjGiGridR.GiGridFld, ByVal Fld As ADODB.Field, Value As Variant)
  On Error GoTo ChkErr
    If ggrData.Recordset.EOF Then Exit Sub
    With giFld
        If .FieldName = "Type" Then Value = IIf(Value = "2", "已日結", "未日結")
        If .FieldName = "CancelFlag" Then Value = IIf(Value = 0, "", "是")
            If .FieldName = "ShouldAmt" Or .FieldName = "OldAmt" Or .FieldName = "RealAmt" Then
                Value = IIf(Value <> 0, Format(Value, "###,###,###"), "")
            End If
            If .FieldName = "PreInvoice" Then
                If Value <> "" Then
                    Select Case Value
                           Case 0
                                Value = "後開"
                           Case 1
                                Value = "預開"
                           Case 2
                                Value = "現場"
                    End Select
                End If
            End If
    End With
  Exit Sub
ChkErr:
    ErrSub Me.Name, "ggrData_ShowCellData"
End Sub

