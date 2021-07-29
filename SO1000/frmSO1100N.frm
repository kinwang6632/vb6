VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSO1100N 
   BorderStyle     =   1  '單線固定
   Caption         =   "帳號相關資料更新 [SO1100N]"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   330
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
   Icon            =   "frmSO1100N.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkApplyInvDate 
      Caption         =   "申請電子計算機發票日期"
      Height          =   255
      Left            =   5370
      TabIndex        =   7
      Top             =   90
      Value           =   1  '核取
      Width           =   2475
   End
   Begin VB.CheckBox chkInvoiceKind 
      Caption         =   "發票開立種類"
      Height          =   255
      Left            =   3690
      TabIndex        =   6
      Top             =   90
      Value           =   1  '核取
      Width           =   1605
   End
   Begin VB.CheckBox chkInvTitle 
      Caption         =   "發票抬頭"
      Height          =   195
      Left            =   2400
      TabIndex        =   4
      Top             =   120
      Value           =   1  '核取
      Width           =   1185
   End
   Begin VB.CheckBox chkChargeTitle 
      Caption         =   "收件人"
      Height          =   195
      Left            =   1350
      TabIndex        =   3
      Top             =   120
      Value           =   1  '核取
      Width           =   915
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "離開(&X)"
      Height          =   345
      Left            =   10410
      TabIndex        =   1
      Top             =   60
      Width           =   1335
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "F2.選定"
      Height          =   345
      Left            =   8700
      TabIndex        =   0
      Top             =   60
      Width           =   1335
   End
   Begin MSComctlLib.ListView lstData 
      Height          =   3060
      Left            =   -30
      TabIndex        =   5
      Top             =   450
      Width           =   11940
      _ExtentX        =   21061
      _ExtentY        =   5398
      View            =   3
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      OLEDragMode     =   1
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   16711680
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDragMode     =   1
      NumItems        =   0
   End
   Begin VB.Label lblItem 
      AutoSize        =   -1  'True
      Caption         =   "更新項目 : "
      Height          =   195
      Left            =   180
      TabIndex        =   2
      Top             =   120
      Width           =   915
   End
End
Attribute VB_Name = "frmSO1100N"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Programmer : Hammer
' Start Date : 2004/09/17
' Last Modify : 2004/09/23

Private WithEvents rsSO002A As ADODB.Recordset
Attribute rsSO002A.VB_VarHelpID = -1

Private Function InitializingCtl() As Boolean
  On Error GoTo ChkErr
    Dim intLoop As Integer
    Dim strTmp As String
    With lstData
         LockWindowUpdate .hwnd
        .ZOrder 0
        .View = 3
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "V", 300
        .ColumnHeaders.Add , , "銀行名稱", 1800, 0
        .ColumnHeaders.Add , , "帳戶別", 1000, 0
        .ColumnHeaders.Add , , "帳號/卡號", 1700, 0
'       .ColumnHeaders.Add , , "檢查碼", 620, 1
        .ColumnHeaders.Add , , "信用卡別", 920, 0
        .ColumnHeaders.Add , , "卡期限", 830, 0
        .ColumnHeaders.Add , , "收費地址", 2500, 0
        .ColumnHeaders.Add , , "郵寄地址", 2500, 0
         LockWindowUpdate 0
    End With
    With rsSO002A
         CloseRS rsSO002A
         Dim strQuerySQL As String
         '#3436 改為多發票,要從SO002AD->SO138 By Kin 2007/12/06
         strQuerySQL = "Select A.BankName,A.ID,A.AccountNo,A.CardName,A.CardExpDate,B.ChargeAddress,B.MailAddress,B.CUSTID,B.COMPCODE,B.ACCOUNTNO,B.INVSEQNO " & _
                        " From " & GetOwner & "SO002A A,(Select C.INVSEQNO,C.CUSTID,C.CompCode,C.AccountNo,D.ChargeAddress,D.MailAddress " & _
                        " From " & GetOwner & "SO002AD C," & GetOwner & "SO138 D" & _
                        " Where C.CUSTID=" & gCustId & " AND C.COMPCODE=" & gCompCode & _
                        " AND C.INVSEQNO=D.INVSEQNO AND D.STOPFLAG<>1) B" & _
                        " Where A.CUSTID=" & gCustId & _
                        " And A.CompCode=" & gCompCode & " And StopFlag<>1" & _
                        " And A.CUSTID=B.CUSTID AND A.COMPCODE=B.COMPCODE" & _
                        " AND A.ACCOUNTNO=B.ACCOUNTNO"
        If Not GetRS(rsSO002A, strQuerySQL, gcnGi, adUseClient, , adLockReadOnly, "", "InitializingClt", False) Then Unload frmSO1100N
         'If Not GetRS(rsSO002A, "SELECT ROWID,SO002A.* FROM " & GetOwner & "SO002A SO002A WHERE CUSTID= " & gCustId & " AND COMPCODE=" & gCompCode & " AND STOPFLAG <> 1", gcnGi, adUseClient, , adLockReadOnly, "", "InitializingClt", False) Then Unload frmSO1100N
        If rsSO002A.EOF Or rsSO002A.BOF Then InitializingCtl = True: Exit Function
        .MoveFirst
    End With
    With lstData
          LockWindowUpdate lstData.hwnd
         .ListItems.Clear
          intLoop = 1
         
          While Not rsSO002A.EOF
                 strTmp = rsSO002A("AccountNo") & ""
                .ListItems.Add
                .ListItems(intLoop).Checked = False
                .ListItems(intLoop).ListSubItems.Add , , rsSO002A("BankName") & ""
                .ListItems(intLoop).ListSubItems.Item(1).Tag = rsSO002A("INVSEQNO")
                .ListItems(intLoop).ListSubItems.Add , , Convert2Description(Val(rsSO002A("ID") & ""))
                .ListItems(intLoop).ListSubItems.Add , , Left(strTmp, 5) & "XXX" & Mid(strTmp, 9, 2) & "XX" & Right(strTmp, 4)
                .ListItems(intLoop).ListSubItems.Item(3).Tag = strTmp
'               .ListItems(intLoop).ListSubItems.Add , , rsSO002A("CVC2") & ""
                .ListItems(intLoop).ListSubItems.Add , , rsSO002A("CardName") & ""
                .ListItems(intLoop).ListSubItems.Add , , Format(rsSO002A("CardExpDate") & "", "00/0000")
                .ListItems(intLoop).ListSubItems.Add , , rsSO002A("ChargeAddress") & ""
                .ListItems(intLoop).ListSubItems.Add , , rsSO002A("MailAddress") & ""
                 intLoop = intLoop + 1
                 rsSO002A.MoveNext
         Wend
    End With
    LockWindowUpdate 0
    Rlx intLoop
    InitializingCtl = True
  Exit Function
ChkErr:
    InitializingCtl = False
    ErrSub Me.Name, "InitializingCtl"
End Function

Private Function Convert2Description(bytValue As Byte) As String
'    0 = 轉帳帳號 '    1 = 信用卡號 '    2 = 虛擬帳號
  On Error GoTo ChkErr
    Select Case bytValue
             Case 0
                    Convert2Description = "轉帳帳號"
             Case 1
                    Convert2Description = "信用卡號"
             Case 2
                    Convert2Description = "虛擬帳號"
    End Select
  Exit Function
ChkErr:
    ErrSub Me.Name, "Convert2Description"
End Function

Private Sub cmdCancel_Click()
    ExitSO1100N
End Sub

'1101
'調整 -修改程式規格
'CSV45_SO1100帳號相關更新警告訊息_20040901_JACY_930915Modify.doc
'
'一、問題描述：
'   1.北桃園:因為常常有修改客戶名稱與發票抬頭之需要,但CSR並不知該客戶是否有帳號,
'       而漏修改帳號上之相關資料,以致拋至發票系統之收件人,發票抬頭有誤.
'
'二、新需求目標
'   1.因為各家之運作模式不同,客戶明細上之發票抬頭,不一定完全等於帳號上之收件人,發票抬頭.故建議:
'       當修改客戶名稱或客戶明細之發票抬頭時,而該客戶有帳號資料時,系統提出警告訊息,
'       ”請確認<帳號>上之相關資料是否需修正!”.
'
'2.並且將SO002A中所有非停用之帳號資料,產生另一瀏覽視窗資料.
'   (1)更新項目包括: 收件人, 發票抬頭 (可選其一或同時勾選)
'   (2)若選取 box勾選,且選擇更新項目(其一或同時),就做更新動作
'       A.選取’收件人’:即以’客戶名稱’更新至’帳號’(SO002A)之’收件人’.
'       B.選取’發票抬頭’:即以客戶明細上之發票種類,統一編號,發票抬頭,發票地址 更新至’帳號’
'       (SO002A)上之發票種類,統一編號,發票抬頭,發票地址.

' SO002A PK 為 CustId,AccountNo,CompCode

'InvoiceType 發票種類
'InvNo 發票統一編號
'InvTitle 發票抬頭
'InvAddress 發票地址

'cboInvoiceType
'mskInvNo
'txtInvTitle
'txtInvAddress

Private Sub cmdOK_Click()
  On Error GoTo ER
    Dim strTmp As String
    If (chkChargeTitle.Value = 0 And chkInvTitle.Value = 0) Then
        ExitSO1100N
        Exit Sub
    End If
    
    With frmSO1100BMDI.ChildForm
            If chkChargeTitle.Value = 1 Then strTmp = "CHARGETITLE='" & Trim(.txtInvTitle) & "'"
            If chkInvTitle.Value = 1 Then
                If strTmp <> Empty Then strTmp = strTmp & ","
                Select Case .cboInvoiceType.ListIndex
                         Case 0
                                strTmp = strTmp & "InvoiceType=0,"
                         Case 1
                                strTmp = strTmp & "InvoiceType=2,"
                         Case 2
                                strTmp = strTmp & "InvoiceType=3,"
                         Case Else
                                strTmp = strTmp & "InvoiceType=Null,"
                End Select
                strTmp = strTmp & "InvNo=" & IIf(.mskInvNo.Text = Empty, "Null", "'" & .mskInvNo.Text & "'") & ","
                strTmp = strTmp & "InvTitle=" & IIf(.txtInvTitle.Text = "", "Null", "'" & .txtInvTitle.Text & "'") & ","
                strTmp = strTmp & "InvAddress=" & IIf(.txtInvAddress.Text = "", "Null", "'" & .txtInvAddress.Text & "'")
                
            End If
            '#5885 測試不OK增加發票種類與申請電子計算發票日期 By Kin 2011/03/14
            If chkInvoiceKind.Value = 1 Then
                strTmp = strTmp & ", InvoiceKind = " & .cboDenRec.ListIndex
            End If
            If chkApplyInvDate.Value = 1 Then
                If Len(.gdaApplyInvDate.GetValue & "") = 0 Then
                    strTmp = strTmp & ",ApplyInvDate = null"
                Else
                    strTmp = strTmp & ", ApplyInvDate = To_Date('" & .gdaApplyInvDate.GetValue & "','yyyymmdd')"
                End If
            End If
    End With
   With lstData
        LockWindowUpdate .hwnd
        For intLoop = 1 To .ListItems.Count
'          Debug.Print "UPDATE " & GetOwner & "SO002A SET " & strTmp & " WHERE CUSTID=" & gCustId & " AND ACCOUNTNO='" & .ListItems(intLoop).ListSubItems.Item(3).Text & "' AND COMPCODE=" & gCompCode
'           If .ListItems(intLoop).Checked Then frmSO1100BMDI.cltSynchronizeSO002A.Add "UPDATE " & GetOwner & "SO002A SET " & strTmp & " WHERE CUSTID=" & gCustId & " AND ACCOUNTNO='" & .ListItems(intLoop).ListSubItems.Item(3).Text & "' AND COMPCODE=" & gCompCode

            '#3436 改為更新SO138 By Kin 2007/12/06
            If .ListItems(intLoop).Checked Then frmSO1100BMDI.cltSynchronizeSO002A.Add "UPDATE " & GetOwner & "SO138 SET " & strTmp & " WHERE INVSEQNO=" & .ListItems(intLoop).ListSubItems.Item(1).Tag
            
           'If .ListItems(intLoop).Checked Then frmSO1100BMDI.cltSynchronizeSO002A.Add "UPDATE " & GetOwner & "SO002A SET " & strTmp & " WHERE CUSTID=" & gCustId & " AND ACCOUNTNO='" & .ListItems(intLoop).ListSubItems.Item(3).Tag & "' AND COMPCODE=" & gCompCode
'           If .SelectedItem.Checked Then frmSO1100BMDI.cltSynchronizeSO002A.Add "UPDATE " & GetOwner & "SO002A SET " & strTmp & " WHERE CUSTID=" & gCustId & " AND ACCOUNTNO='" & .ListItems(intLoop).ListSubItems.Item(3).Text & "' AND COMPCODE=" & gCompCode
        Next
        LockWindowUpdate 0
    End With
    Rlx strTmp
    ExitSO1100N
  Exit Sub
ER:
   ErrSub Me.Name, "cmdOk_Click"
End Sub

Private Sub ExitSO1100N()
  On Error Resume Next
    frmSO1100N.Hide
    DoEvents
    Unload frmSO1100N
End Sub

Private Sub Form_Activate()
  On Error Resume Next
    FrmOnTop frmSO1100N
End Sub

Public Sub CancelGo()
  On Error GoTo ChkErr
    cmdCancel.Value = True
  Exit Sub
ChkErr:
    ErrSub Me.Name, "CancelGo"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error GoTo ER
    If KeyCode = vbKeyF2 And cmdOk.Enabled Then
        KeyCode = 0
        Call cmdOK_Click
    End If
    If KeyCode = vbKeyEscape And cmdCancel.Enabled Then
        KeyCode = 0
        Call cmdCancel_Click
        Exit Sub
    End If
   Exit Sub
ER:
   ErrSub Me.Name, "Form_KeyDown"
End Sub

Private Sub Form_Load()
  On Error GoTo ChkErr
    If Crm Then
        frmTVBasic.Move frmSO1100BMDI.Left, frmSO1100BMDI.Top + 4200
    Else
        If Not800600 Then Me.Move (Screen.Width - Me.Width) / 2, ((Screen.Height - Me.Height) / 2) + 2000
    End If
    Set ObjActiveForm = frmSO1100N
    Set rsSO002A = New ADODB.Recordset
    If Not InitializingCtl Then End
  Exit Sub
ChkErr:
   ErrSub Me.Name, "Form_Load"
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
    CloseRS rsSO002A
    Set rsSO002A = Nothing
    Set ObjActiveForm = frmSO1100BMDI
    ReleaseCOM frmSO1100N
    Set frmSO1100N = Nothing
End Sub

Private Sub lstData_DblClick()
  On Error GoTo ChkErr
    With lstData
         If .ListItems.Count = 0 Then Exit Sub
        .SelectedItem.Checked = Not .SelectedItem.Checked
    End With
  Exit Sub
ChkErr:
    ErrHandle Me.Name, "lstData_DblClick"
End Sub
