VERSION 5.00
Object = "{3BB22E57-824A-11D3-BD70-0080C8F80BC4}#3.4#0"; "GiGridR.ocx"
Begin VB.Form frmSO1145A 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  '單線固定
   Caption         =   "客戶地址記錄瀏覽 [SO1145A]"
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   3870
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
   Icon            =   "SO1145A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   Begin prjGiGridR.GiGridR ggrData 
      Height          =   4455
      Left            =   105
      TabIndex        =   0
      Top             =   90
      Width           =   11670
      _ExtentX        =   20585
      _ExtentY        =   7858
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
Attribute VB_Name = "frmSO1145A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
' 記錄目前在編輯、新增或檢視模式

Private lngEditMode As giEditModeEnu
Private strParam0 As String
Private rsSO1145A As New ADODB.Recordset
Attribute rsSO1145A.VB_VarHelpID = -1
Private m_CustID As Long

'記錄目前使用者權限
Private bleUserPrivAddNew   As Boolean   '新增
Private bleUserPrivUpdate   As Boolean  '修改
Private bleUserPrivPrint As Boolean   '列印

Public Property Let uCustId(ByVal vData As Long)
  On Error GoTo ChkErr
    m_CustID = vData
  Exit Property
ChkErr:
    ErrSub Me.Name, "Let uCustID"
End Property

Public Sub CancelGo()
  On Error Resume Next
    Unload Me
End Sub

Private Sub Form_Activate()
  On Error Resume Next
    Screen.MousePointer = vbDefault
    Set ObjActiveForm = Me
    strActFrmName = "frmSO1144A"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo ChkErr
        If KeyCode = vbKeyEscape Then Unload Me
    Exit Sub
ChkErr:
    ErrSub Me.Name, "Form_KeyDown"
End Sub

Private Sub Form_Load()
  '取global變數內容,取資料內容
 On Error GoTo ChkErr
   Dim mFlds As New GiGridFlds
    If Crm Then
        PolyFrmFunction Me
        Me.Move frmSO1100BMDI.Left + ((frmSO1100BMDI.Width - Me.Width) / 2), frmSO1100BMDI.Top + ((frmSO1100BMDI.Height - Me.Height) / 2 + 1234)
    Else
        If Not800600 Then Me.Move (Screen.Width - Me.Width) / 2, ((Screen.Height - Me.Height) / 2) + 1543
    End If
   strParam0 = "SELECT INSTADDRESS,INSTADDRESSB,UPDTIME,UPDEN FROM " & GetOwner & "SO101 WHERE CUSTID=" & m_CustID & " AND SERVICETYPE='" & gServiceType & "' AND MODEID=3 AND COMPCODE=" & gCompCode & " ORDER BY UPDTIME DESC"
   'for get nextcode
   With rsSO1145A
        .CursorLocation = adUseClient
        .Open strParam0, gcnGi, adOpenForwardOnly, adLockReadOnly
   End With
   With mFlds
        .Add "INSTADDRESS", , , , , "舊地址" & Space(52)
        .Add "INSTADDRESSB", , , , , "新地址" & Space(52)
        .Add "UPDTIME", giControlTypeTime, , , , "異動時間" & Space(15)
        .Add "UPDEN", , , , , "異動人員" & Space(12)
   End With
   With ggrData
       .AllFields = mFlds
       .SetHead
        Set .Recordset = rsSO1145A
       .Refresh
   End With
   MenuEnabled , , , , , True
 Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "Form_Load")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  On Error GoTo ChkErr
    If rsSO1145A.State = adStateOpen Then rsSO1145A.Close
    Set rsSO1145A = Nothing
    Call FormQueryUnload
  Exit Sub
ChkErr:
    ErrSub Me.Name, "Form_QueryUnload"
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
    ReleaseCOM Me
End Sub
