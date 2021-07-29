VERSION 5.00
Object = "{3BB22E57-824A-11D3-BD70-0080C8F80BC4}#3.4#0"; "GiGridR.ocx"
Begin VB.Form frmSO1100C 
   BorderStyle     =   1  '單線固定
   Caption         =   "資料重複 [SO1100C]"
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
   Icon            =   "SO1100C.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   Begin prjGiGridR.GiGridR ggrData 
      Height          =   2730
      Left            =   135
      TabIndex        =   3
      Top             =   630
      Width           =   11550
      _ExtentX        =   20373
      _ExtentY        =   4815
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
      Caption         =   "離開(&X)"
      Height          =   375
      Left            =   10320
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "F2.選定"
      Height          =   375
      Left            =   8640
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "電話(1)重複: 9999999, 合計: 999位"
      Height          =   195
      Left            =   180
      TabIndex        =   2
      Top             =   210
      Width           =   2745
   End
End
Attribute VB_Name = "frmSO1100C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Programmer: Hammer
' Start Date: 1999/11/11
' Last Modify: 2001/05/17
Dim rs As New ADODB.Recordset
Public strFlag As String   ' 用來判斷是電話(1)重複還是裝機地址重複
Public Flag As Long         ' =2(電話2重覆);  =3(電話3重覆);  =其他(電話1重覆)
Public PhoneNo As String

Private Sub cmdCancel_Click()
  On Error GoTo ER
    frmSO1100BMDI.DataDup -1
'   frmSO1100B.lngCustID = -1
    Unload Me
  Exit Sub
ER:
   ErrSub Me.Name, "cmdCancel_Click"
End Sub

Private Sub cmdOK_Click()
  On Error GoTo ER
    frmSO1100BMDI.AfterDataDup rs("CustId"), strFlag
  Exit Sub
ER:
   ErrSub Me.Name, "cmdOk_Click"
End Sub

Private Sub Form_Activate()
  On Error Resume Next
    FrmOnTop Me
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
    If KeyCode = vbKeyF2 And cmdOK.Enabled Then
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
    Dim mFlds As New GiGridFlds
    If Crm Then
        Me.Move frmSO1100BMDI.Left + ((frmSO1100BMDI.Width - Me.Width) / 2), frmSO1100BMDI.Top + ((frmSO1100BMDI.Height - Me.Height) / 2)
    Else
        Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    End If
    
    Set ObjActiveForm = frmSO1100C
    Flag = IIf(Flag <> 2 And Flag <> 3, 1, Flag)
'HM  mflds.Add "InitTime", giControlTypeTime, , , , "建檔時間    "
    With mFlds
        .Add "CompName", , , , , "公司別    "
        .Add "CustID", , , , , "客戶編號", vbRightJustify
        .Add "CustName", , , , , "客戶名稱      "
        .Add "CustStatusName", , , , , "客戶狀態"
        .Add "WipName1", , , , , "派工狀態1" & Space(10)
        .Add "WipName3", , , , , "派工狀態3" & Space(10)
        .Add "InitTime", , , , , "建檔時間    "
        .Add "InstAddress", , , , , "裝機地址                                "
    End With
    
    With ggrData
        .AllFields = mFlds
        .SetHead
    End With
    
    ' MainServiceType
    rs.CursorLocation = adUseClient
    
    ' 與 Penny 討論後調整
    ' TBC 用 GServiceType
    ' EMC 用 MainServiceType
    ' 2006/11/22 PM 5:51
    
    Dim strMainSvc As String
    On Error Resume Next
    strMainSvc = RPxx(gcnGi.Execute("SELECT MAINSERVICETYPE FROM " & GetOwner & "SO041 WHERE ROWNUM=1").GetString(adClipString, 1, ",", vbCrLf, ""))
    If garyGi(17) = 1 Then strMainSvc = gServiceType
    '#6064 增加派工狀態1與3 By Kin 2011/07/25
    If strFlag = "電話重複" Then
        Select Case Flag
            Case 1
                If OpenRecordset(rs, "SELECT A.CompName,A.CustID,A.CustName,B.CustStatusName," & _
                   " A.InitTime,A.InstAddress,B.WipName1,B.WipName3 " & _
                   " FROM " & GetOwner & "SO001 A," & GetOwner & "SO002 B " & _
                   " WHERE A.Tel1='" & _
                    PhoneNo & "' AND B.CUSTID=A.CUSTID AND B.SERVICETYPE='" & strMainSvc & "'" & _
                    " AND A.CUSTID <> " & gCustId, gcnGi, adOpenForwardOnly, adLockReadOnly) < giOK Then
                    GoTo ChkErr
                    Exit Sub
                End If
                lblNote = "電話(1)重複: " & PhoneNo & " ，合計： " & rs.RecordCount & " 位"
            Case 2
                If OpenRecordset(rs, "SELECT A.CompName,A.CustID,A.CustName,B.CustStatusName," & _
                   " A.InitTime,A.InstAddress,B.WipName1,B.WipName3 " & _
                   " FROM " & GetOwner & "SO001 A," & GetOwner & "SO002 B WHERE A.Tel2 = '" & _
                    PhoneNo & "' AND B.CUSTID=A.CUSTID AND B.SERVICETYPE='" & strMainSvc & "'" & _
                    " AND A.CUSTID <> " & gCustId, gcnGi, adOpenForwardOnly, adLockReadOnly) < giOK Then
                    GoTo ChkErr
                    Exit Sub
                End If
                lblNote = "電話(2)重複: " & PhoneNo & " ，合計： " & rs.RecordCount & " 位"
            Case 3
                If OpenRecordset(rs, "SELECT A.CompName,A.CustID,A.CustName,B.CustStatusName," & _
                   " A.InitTime,A.InstAddress,B.WipName1,B.WipName3 " & _
                   " FROM " & GetOwner & "SO001 A," & GetOwner & "SO002 B WHERE A.Tel3 = '" & _
                    PhoneNo & "' AND B.CUSTID=A.CUSTID AND B.SERVICETYPE='" & strMainSvc & "'" & _
                    " AND A.CUSTID <> " & gCustId, gcnGi, adOpenForwardOnly, adLockReadOnly) < giOK Then
                    GoTo ChkErr
                    Exit Sub
                End If
                lblNote = "電話(3)重複: " & PhoneNo & " ，合計： " & rs.RecordCount & " 位"
        End Select
    ' 地址重複
    Else
        If OpenRecordset(rs, "SELECT A.CompName,A.CustID,A.CustName,B.CustStatusName," & _
           " A.InitTime,A.InstAddress,B.WipName1,B.WipName3 " & _
           " FROM " & GetOwner & "SO001 A," & GetOwner & "SO002 B WHERE A.InstAddrNo = " & _
            frmSO1100BMDI.ga1InstAddress.AddrNo & _
           " AND B.CUSTID=A.CUSTID AND B.SERVICETYPE='" & gServiceType & "'", gcnGi, adOpenForwardOnly, adLockReadOnly) < giOK Then
            GoTo ChkErr
            Exit Sub
        End If
        lblNote = "地址重複: " & frmSO1100BMDI.ga1InstAddress.AddrString & " ，合計： " & rs.RecordCount & " 位"
    End If
    With ggrData
         Set .Recordset = rs
        .Refresh
    End With
  Exit Sub
ChkErr:
   ErrSub Me.Name, "Form_Load"
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error GoTo ER
    Set ObjActiveForm = frmSO1100BMDI
    CloseRS rs
    On Error Resume Next
    ReleaseCOM Me
    'frmSO1100BMDI.PhoneObj.SetFocus
  Exit Sub
ER:
   ErrSub Me.Name, "Form_Unload"
End Sub

Private Sub ggrData_DblClick()
  On Error GoTo ER
    If cmdOK.Enabled Then cmdOK.Value = True
  Exit Sub
ER:
   ErrSub Me.Name, "ggrData_DblClick"
End Sub

Private Sub ggrdata_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error GoTo ER
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        If cmdOK.Enabled Then cmdOK.Value = True
    End If
  Exit Sub
ER:
    ErrSub Me.Name, "ggrData_KeyDown"
End Sub

Private Sub ggrData_ShowCellData(ByVal Col As Long, ByVal Row As Long, ByVal giFld As prjGiGridR.GiGridFld, ByVal Fld As ADODB.Field, Value As Variant)
'  On Error GoTo ChkErr
'    '#6064 增加派工狀態1與3 By Kin 2011/07/29
'    Select Case UCase(Fld.Name)
'        Case UCase("WipCode1")
'            If Len(Value) > 0 Then
'                Value = GetRsValue("SELECT Description FROM " & GetOwner & "CD036 WHERE CODENO = " & Value) & ""
'            Else
'                Value = ""
'            End If
'        Case UCase("WipCode3")
'            If Len(Value) > 0 Then
'                Value = GetRsValue("SELECT Description FROM " & GetOwner & "CD036 WHERE CODENO = " & Value) & ""
'            Else
'                Value = ""
'            End If
'    End Select
'  Exit Sub
'ChkErr:
'  Call ErrSub(Me.Name, "ggrData_ShowCellData")
End Sub
