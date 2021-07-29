VERSION 5.00
Object = "{3BB22E57-824A-11D3-BD70-0080C8F80BC4}#3.4#0"; "GiGridR.ocx"
Begin VB.Form frmSO114AA 
   BorderStyle     =   1  '單線固定
   Caption         =   $"SO114AA.frx":0000
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
   Icon            =   "SO114AA.frx":001F
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
Attribute VB_Name = "frmSO114AA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'CM與CT鍵功能說明:
'當使用者Click CM Button or CT Button時, 則另開一New Form顯示so152之資料內容.如圖 :
'Gird中所顯示的內容即為so152之全部內容(不管是Click [CM] or [CT] button), 其讀取資料的語法如下 :
'select * from so152 where EmcCustID=so001.custid and EmcCompCode=so001.CompCode order by UpdTime desc
'Gird欄位中的Caption中文部分, 請參考EcSchema.mdb中之so152結構資料

Private rsSO152 As New ADODB.Recordset

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error Resume Next
    If KeyCode = 27 Then Unload Me
End Sub

Private Sub Form_Load()
  On Error GoTo ChkErr
    If Crm Then
        PolyFrmFunction Me
        Me.Move frmSO1100A.Left + ((frmSO1100A.Width - Me.Width) / 2), frmSO1100A.Top + ((frmSO1100A.Height - Me.Height) / 2)
    Else
        If Not800600 Then Me.Move (Screen.Width - Me.Width) / 2, ((Screen.Height - Me.Height) / 2) + 1800
    End If
    GetRS rsSO152, "select * from " & GetOwner & "so152 where emccustid=" & gCustId & " and emccompcode=" & gCompCode & " order by updtime desc", gcnGi, adUseClient, adOpenForwardOnly, adLockReadOnly
    GrdFmt
  Exit Sub
ChkErr:
    ErrSub Me.Name, "Form_Load"
End Sub

Public Sub GrdFmt()
  On Error GoTo ChkErr
    Dim mFlds As New GiGridFlds
    With mFlds
        .Add "EmcCompCode", , , , , "EMC公司別", vbLeftJustify
        .Add "EmcCustID", , , , , "EMC客戶編號", vbLeftJustify
        .Add "EbtContractNo", , , , , "合約編號         ", vbLeftJustify
        .Add "EbtCustID", , , , , "EBT客戶編號", vbLeftJustify
        .Add "EbtContractBDate", giControlTypeDate, , , , " 合約生效日 ", vbLeftJustify
        .Add "EbtContractEDate", giControlTypeDate, , , , " 合約截止日 ", vbLeftJustify
        .Add "EbtCustCName", , , , , "客戶中文名稱  ", vbLeftJustify
        .Add "EbtCustContactPhone", , , , , "客戶聯絡電話", vbLeftJustify
        .Add "EbtCustContactMobile", , , , , "客戶行動電話", vbLeftJustify
        .Add "EbtCompOwnerName", , , , , "公司負責人中文姓名", vbLeftJustify
        .Add "EbtContactPhone", , , , , "(公司)聯絡人聯絡電話", vbLeftJustify
        .Add "EbtItContactName", , , , , "技術聯絡人中文姓名", vbLeftJustify
        .Add "EbtItContactPhone", , , , , "技術聯絡人電話", vbLeftJustify
        .Add "EbtItContactMobile", , , , , "技術聯絡人行動電話", vbLeftJustify
        .Add "EbtItEMail", , , , , "技術聯絡人電子郵件  ", vbLeftJustify
        .Add "EbtInstAddr", , , , , "裝機地址                      ", vbLeftJustify
        .Add "EbtCustAddr", , , , , "戶籍地址                      ", vbLeftJustify
        .Add "EbtBillAddr", , , , , "帳單地址                      ", vbLeftJustify
        .Add "EbtContractStatusCode", , , , , "合約狀態代碼", vbLeftJustify
        .Add "EbtContractStatusDesc", , , , , "合約狀態說明", vbLeftJustify
        .Add "EbtFeePeriodCode", , , , , "EBT繳費週期代碼", vbLeftJustify
        .Add "EbtFeePeriodDesc", , , , , "EBT繳費週期說明", vbLeftJustify
        .Add "EbtServiceType", , , , , "服務類別", vbLeftJustify
        .Add "EbtAgentName", , , , , "代理人姓名 ", vbLeftJustify
        .Add "EbtAgentPhone", , , , , "代理人電話 ", vbLeftJustify
        .Add "EbtAgentID", , , , , "代理人身份證號", vbLeftJustify
        .Add "EbtAgentAddress", , , , , "代理人地址                      ", vbLeftJustify
        .Add "EbtIdCardId", , , , , "個人身份證號", vbLeftJustify
        .Add "EbtCompanyOwnerId", , , , , "負責人身份證號", vbLeftJustify
        .Add "EbtNonProfitCompanyId", , , , , "非營利法人證號", vbLeftJustify
        .Add "EbtCompanyId", , , , , "公司證照號碼", vbLeftJustify
        .Add "IfMoveToOtherSo", , , , , "是否已經移機", vbLeftJustify
        .Add "EbtNotes", , , , , "EBT備註" & Space(24), vbLeftJustify
        .Add "UpdTime", giControlTypeTime, , , , "異動時間            ", vbLeftJustify
        .Add "UpdEn", , , , , "異動人員   ", vbLeftJustify
'        .Add "EbtAgentZipCode", , , , , "代理人郵遞區號", vbLeftJustify
'        .Add "EbtAgentCity", , , , , "代理人縣市", vbLeftJustify
'        .Add "EbtAgentTown", , , , , "代理人鄉鎮", vbLeftJustify
    End With
    With ggrData
        .AllFields = mFlds
        .SetHead
         Set .Recordset = rsSO152
        .Refresh
    End With
    Me.Show 1
  Exit Sub
ChkErr:
    ErrSub Me.Name, "OpenData"
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
    rsSO152.Close
    Set rsSO152 = Nothing
    ReleaseCOM Me
End Sub

Private Sub ggrData_ShowCellData(ByVal Col As Long, ByVal Row As Long, ByVal giFld As prjGiGridR.GiGridFld, ByVal Fld As ADODB.Field, Value As Variant)
  On Error GoTo ChkErr
    If giFld.FieldName = "IfMoveToOtherSo" Then Value = IIf(Val(Value & "") = 0, "", "是")
  Exit Sub
ChkErr:
    ErrSub Me.Name, "ggrData_ShowCellData"
End Sub
