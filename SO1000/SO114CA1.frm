VERSION 5.00
Object = "{3BB22E57-824A-11D3-BD70-0080C8F80BC4}#3.4#0"; "GiGridR.ocx"
Begin VB.Form frmSO114CA1 
   BorderStyle     =   1  '單線固定
   Caption         =   "認證設備序號"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4260
   Icon            =   "SO114CA1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   4260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所屬視窗中央
   Begin prjGiGridR.GiGridR ggrData 
      Height          =   4065
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   4245
      _ExtentX        =   7488
      _ExtentY        =   7170
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
   Begin VB.CommandButton cmdOk 
      Caption         =   "確 定(&Y)"
      Enabled         =   0   'False
      Height          =   405
      Left            =   30
      Picture         =   "SO114CA1.frx":0442
      TabIndex        =   0
      Top             =   4080
      Width           =   1500
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取 消(&N)"
      Height          =   405
      Left            =   2730
      TabIndex        =   1
      Top             =   4080
      Width           =   1500
   End
End
Attribute VB_Name = "frmSO114CA1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rs As New ADODB.Recordset
Attribute rs.VB_VarHelpID = -1
Private strReturnValue As String

Public Property Get ReturnValue() As String
    ReturnValue = strReturnValue
End Property

Private Sub cmdCancel_Click()
  On Error GoTo ChkErr
    Unload Me
  Exit Sub
ChkErr:
    ErrSub Me.Name, "cmdCancel_Click"
End Sub

Private Sub cmdOk_Click()
  On Error Resume Next
    strReturnValue = rs(0).Value & ""
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error GoTo ChkErr
    If KeyCode = vbKeyEscape Then Unload Me
  Exit Sub
ChkErr:
    ErrSub Me.Name, "Form_KeyDown"
End Sub

Private Sub Form_Load()
  On Error GoTo ChkErr
    InitData
    subGrd
    OpenData
  Exit Sub
ChkErr:
    ErrSub Me.Name, "Form_Load"
End Sub

Private Sub InitData()
  On Error Resume Next
    Dim lngHeight As Long
    If Crm Then
        Move frmSO1100BMDI.Left + ((frmSO1100BMDI.Width - Me.Width) / 2), frmSO1100BMDI.Top + ((frmSO1100BMDI.Height - Me.Height) / 2)
    Else
        Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    End If
    strReturnValue = ""
End Sub

Private Sub OpenData()
    On Error GoTo ChkErr
    Dim strSQL As String
    Set rs = New ADODB.Recordset
    strSQL = "SELECT FACISNO,SMARTCARDNO FROM " & GetOwner & "SO004" & _
                    " WHERE CUSTID=" & gCustId & " AND FACISNO IS NOT NULL " & _
                    " AND ( INSTDATE IS NOT NULL AND PRDATE IS NULL )" & _
                    " AND FACICODE IN ( SELECT CODENO FROM " & GetOwner & "CD022 WHERE REFNO=3) "
                    ' Law : 過濾該CustId and InstDate有值 and PRDate無值 and FaciCode對應到CD022.RefNo=3
    If Not GetRS(rs, strSQL) Then Exit Sub
    If rs.RecordCount > 0 Then cmdOk.Enabled = True
    Set ggrData.Recordset = rs
    ggrData.Refresh
  Exit Sub
ChkErr:
    ErrSub Me.Name, "OpenData"
End Sub

Private Sub subGrd()
  On Error GoTo ChkErr
    Dim mflds As New GiGridFlds
    With ggrData
            mflds.Add "FaciSNo", , , , , "物品序號                 "
            mflds.Add "SmartCardNo", , , , , "智慧卡序號            "
            .AllFields = mflds
            .SetHead
    End With
  Exit Sub
ChkErr:
    ErrSub Me.Name, "subGrd"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  On Error Resume Next
    CloseRecordset rs
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ChkErr
        ReleaseCOM frmSO114CA1
    Exit Sub
ChkErr:
    ErrSub Me.Name, "Form_Unload"
End Sub

Private Sub ggrData_DblClick()
    cmdOk.Value = True
End Sub
