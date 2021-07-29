VERSION 5.00
Object = "{3BB22E57-824A-11D3-BD70-0080C8F80BC4}#3.4#0"; "GiGridR.ocx"
Begin VB.Form frmSO1133A 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "歷史收費資料瀏覽 [SO1133A]"
   ClientHeight    =   4740
   ClientLeft      =   45
   ClientTop       =   3765
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
   Icon            =   "SO1133A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   Begin prjGiGridR.GiGridR ggrData 
      Height          =   4605
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   11745
      _ExtentX        =   20717
      _ExtentY        =   8123
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseCellForeColor=   -1  'True
   End
End
Attribute VB_Name = "frmSO1133A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Program Name :     SO1133A
' Version      :     1.0.0(Alpha)
' Developer    :     Hammer
' Written Date :     2000/03/04
' Last Modify  :     2000/03/04

Option Explicit
Dim rs As New ADODB.Recordset
Attribute rs.VB_VarHelpID = -1
Dim lngCustId As Long

Public Sub UserPermissionGo()

End Sub

Public Sub CancelGo()
  On Error GoTo ChkErr
    Unload Me
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "cmdCancel_Click")
End Sub

Private Sub ChangeMode()
    On Error GoTo ChkErr
        MenuEnabled , , , , , True
    Exit Sub
ChkErr:
    ErrSub Me.Name, "ChangeMode"
End Sub

Private Sub subGrd()
    On Error GoTo ChkErr
    Dim mFlds As New GiGridFlds
    '定義GirdR各個欄位的內容
        mFlds.Add "Type", , , , , " 種類  ", vbLeftJustify
        mFlds.Add "BillNo", , , , , "      單據編號    ", vbLeftJustify
        mFlds.Add "CitemName", , , , , "收費項目  ", vbLeftJustify
        mFlds.Add "ShouldAmt", , , , , "出單$", vbRightJustify
        mFlds.Add "RealAmt", , , , , "實收$", vbRightJustify
        mFlds.Add "Realperiod", , , , , " 期 ", vbRightJustify
        mFlds.Add "ShouldDate", giControlTypeDate, , , , "  應收日期  ", vbLeftJustify
        mFlds.Add "RealDate", giControlTypeDate, , , , "  實收日期  ", vbLeftJustify
        mFlds.Add "RealStartDate", giControlTypeDate, , , , "   起始日   ", vbLeftJustify
        mFlds.Add "RealStopDate", giControlTypeDate, , , , "   截止日   ", vbLeftJustify
        mFlds.Add "UCName", , , , , "未收原因", vbLeftJustify
        mFlds.Add "STName", , , , , " 短收原因", vbLeftJustify
        mFlds.Add "ClctName", , , , , "收費人員", vbLeftJustify
        mFlds.Add "CancelFlag", , , , , "作廢", vbLeftJustify
        mFlds.Add "CancelName", , , , , "作廢原因", vbLeftJustify
        mFlds.Add "Compcode", , , , , "公司別", vbLeftJustify
        mFlds.Add "CMName", , , , , "收費方式", vbLeftJustify
        mFlds.Add "CitibankATM", , , , , "虛擬帳號          ", vbLeftJustify
        mFlds.Add "GUINo", , , , , "發票號碼", vbLeftJustify
        mFlds.Add "InvDate", giControlTypeDate, , , , "  發票日期  ", vbLeftJustify
        mFlds.Add "FaciSNo", , , , , "設備序號     ", vbLeftJustify
        
        mFlds.Add "Note", , , , , "備註" & Space(30), , , vbLeftJustify
        mFlds.Add "OldAmt", , , , , "原應收金額", vbLeftJustify
        mFlds.Add "UpdTime", , , , , "異動時間            ", vbLeftJustify
        mFlds.Add "UpdEn", , , , , "異動人員", vbLeftJustify
        '連結giGirdR
        ggrData.AllFields = mFlds
        ggrData.SetHead
    Exit Sub
ChkErr:
    ErrSub Me.Name, "subGrd"
End Sub

Private Sub Form_Activate()
    On Error Resume Next
        Set ObjActiveForm = Me
        strActFrmName = "frmSO1133A"

        Screen.MousePointer = vbDefault
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error GoTo ChkErr
    If KeyCode = vbKeyEscape Then Unload Me
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "Form_KeyDown")
End Sub

Private Sub Form_Load()
    On Error GoTo ChkErr
        If Alfa2 Then Call GetGlobal
        If Crm Then
            PolyFrmFunction Me
            Me.Move frmCrm.trv.Width + 66, frmCrm.lvwItem.Height + 3500
            frmCrm.picService.Enabled = False
        Else
            If Not800600 Then
                Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2 + 1450
            End If
        End If
        Call ChangeMode
        Call subGrd         '設定grid屬性
        Call OpenData
        Screen.MousePointer = 0
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "Form_Load")
End Sub

Private Sub OpenData()
    On Error GoTo ChkErr
    Dim strSQL As String
    
        strSQL = ChargeField
        If Not GetRS(rs, "Select * From (Select 1 as Type ," & strSQL & " From " & GetOwner & "So035 Where CustId = " & lngCustId & " And ServiceType = '" & gServiceType & "' And (STCode In (Select CodeNo From " & GetOwner & "CD016 Where Nvl(RefNo,0) <> 1 ) OR STCode Is Null) " & _
            " Union All Select 2 as Type ," & strSQL & " From " & GetOwner & "SO035 Where CustId =" & lngCustId & " And ServiceType = '" & gServiceType & "' And STCode In (Select CodeNo From " & GetOwner & "CD016 Where RefNo =1 ) " & _
            " Union All Select 3 as Type ," & strSQL & " From " & GetOwner & "SO036 Where CustId =" & lngCustId & " And ServiceType = '" & gServiceType & "' )  Order By UpdTime Desc", gcnGi, adOpenStatic, adLockReadOnly) Then Exit Sub
        Set ggrData.Recordset = rs
        ggrData.Refresh
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "OpenData")
End Sub

Public Property Get uCustId() As Long
  On Error GoTo ChkErr
    uCustId = lngCustId
  Exit Property
ChkErr:
    ErrSub Me.Name, "Get uCustID"
End Property

Public Property Let uCustId(ByVal vData As Long)
  On Error GoTo ChkErr
    lngCustId = vData
  Exit Property
ChkErr:
    ErrSub Me.Name, "Let uCustID"
End Property

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
        Set rs = Nothing
        Call FormQueryUnload2(frmSO1100BMDI)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
        ReleaseCOM frmSO1133A
End Sub

Public Sub ViewGo()
    On Error GoTo ChkErr
        '判斷Type不為1（未日結），則以訊息提示並離開
        '判斷CancelFlag=1（已作廢），以訊息提示並離開
'        If rsSO033("Type") = 2 Then MsgBox "已日結資料不能修改！", vbExclamation, "訊息": Exit Sub
'        If rsSO033("CancelFlag") = 1 Then MsgBox "已作廢資料不能修改！", vbExclamation, "訊息": Exit Sub
        '顯示修改畫面 SO1132C
        With frmSO1132C
            .strBillNo = rs("BillNo").Value
            .intItem = rs("Item").Value
            .uOpenMyself = True
            .DBFlag = giOracle
            .EditMode = giEditModeView
            .ZZ = 2
            .uCustId = lngCustId
            Set .uParentForm = Me
            MenuEnabled , , , , , True
            .Show , Me
        End With
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "ViewGo")
End Sub

Private Sub ggrData_ColorCellData(ByVal Col As Long, ByVal Row As Long, ByVal giFld As prjGiGridR.GiGridFld, ByVal Fld As ADODB.Field, Value As Variant)
    On Error Resume Next
        If giFld.FieldName = "CancelFlag" Then
            If Value <> 0 Then Value = vbRed
        End If
        If giFld.FieldName = "FaciSNo" Then
            If ggrData.Recordset("CitemCode") & "" <> "" Then
                If GetRsValue("Select ByHouse From " & GetOwner & "CD019 Where CodeNo = " & ggrData.Recordset("CitemCode")) = 1 Then
                    Value = "BY戶"
                End If
            End If
        End If
End Sub

Private Sub ggrData_DblClick()
    On Error Resume Next
    If rs.EOF Or rs.BOF Then Exit Sub
    Call ViewGo
End Sub

Private Sub ggrData_ShowCellData(ByVal Col As Long, ByVal Row As Long, ByVal giFld As prjGiGridR.GiGridFld, ByVal Fld As ADODB.Field, Value As Variant)
    On Error Resume Next
        If UCase(giFld.FieldName) = "TYPE" Then
            Select Case Value
                Case 1
                    Value = "作廢"
                Case 2
                    Value = "呆帳"
                Case 3
                    Value = "歷史資料"
            End Select
        ElseIf UCase(giFld.FieldName) = "CANCELFLAG" Then
            If Value = 1 Then
                Value = "是"
            Else
                Value = "否"
            End If
        End If
        If giFld.FieldName = "FaciSNo" Then
            If GetRsValue("Select ByHouse From " & GetOwner & "CD019 Where CodeNo = " & ggrData.Recordset("CitemCode")) = 1 Then
                Value = "BY戶"
            End If
        End If
        
End Sub
