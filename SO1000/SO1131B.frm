VERSION 5.00
Object = "{3BB22E57-824A-11D3-BD70-0080C8F80BC4}#3.4#0"; "GiGridR.ocx"
Begin VB.Form frmSO1131B 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '��u�T�w
   Caption         =   "�Φ��������� [SO1131B]"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   4470
   ClientWidth     =   11910
   BeginProperty Font 
      Name            =   "�s�ө���"
      Size            =   9.75
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SO1131B.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "����(&X)"
      Height          =   465
      Left            =   10350
      TabIndex        =   3
      Top             =   3810
      Visible         =   0   'False
      Width           =   1215
   End
   Begin prjGiGridR.GiGridR ggrData 
      Height          =   3525
      Left            =   120
      TabIndex        =   2
      Top             =   810
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   6218
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblCitemCode 
      AutoSize        =   -1  'True
      Caption         =   "�g���ʦ��O����"
      Height          =   195
      Left            =   135
      TabIndex        =   6
      Top             =   465
      Width           =   1365
   End
   Begin VB.Label lblMainCustName 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  '��u�T�w
      Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXXX"
      Height          =   285
      Left            =   3735
      TabIndex        =   1
      Top             =   75
      Width           =   3435
   End
   Begin VB.Label lblMainCustId 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  '��u�T�w
      Caption         =   "99999999"
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Top             =   75
      Width           =   870
   End
   Begin VB.Label lbl2 
      AutoSize        =   -1  'True
      Caption         =   "�W��"
      Height          =   195
      Left            =   3255
      TabIndex        =   5
      Top             =   105
      Width           =   390
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      Caption         =   "�Φ���Ȥ�s��"
      Height          =   195
      Left            =   135
      TabIndex        =   4
      Top             =   105
      Width           =   1365
   End
End
Attribute VB_Name = "frmSO1131B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Program Name :     SO1131B
' Version      :     1.0.0(Alpha)
' Developer    :     Hammer
' Written Date :     1999/11/20
' Last Modify  :     1999/12/16

Option Explicit
Private lngMainCustId As String
Private strMainCustName As String
Private rs As New ADODB.Recordset

Public Sub MainCustGo()
    On Error GoTo ChkErr
        Dim strMduId As String
        Dim rsTmp As New ADODB.Recordset
        
        strMduId = frmSO1100BMDI.rsSo001("MduId") & ""
        If Len(strMduId & "") = 0 Then MsgBox "�Ȥ�s��:[ " & gCustId & "] �D�j�Ӥ�!!", vbExclamation, "�T��..": Exit Sub
        If Not GetRS(rsTmp, "Select MainCustId From " & GetOwner & "SO017 Where MduId ='" & strMduId & "'", gcnGi) Then Exit Sub
        
        If rsTmp.RecordCount = 0 Then MsgBox "�j�Ӹ���ɤ��L���j�ӽs��!!", , "�T��..": Exit Sub
        lngMainCustId = Val(rsTmp("MainCustId") & "")
        If lngMainCustId = 0 Then MsgBox "�j�ӽs��: [" & strMduId & " ]�D�Φ���,�Ьd��!!", vbExclamation, "�T��..": Exit Sub
        
        If Not GetRS(rsTmp, "Select CustId,CustName,MduId From " & GetOwner & "SO001 Where CustId =" & lngMainCustId, gcnGi) Then Exit Sub
        If rsTmp.RecordCount = 0 Then MsgBox "�j�ӽs��: [" & strMduId & " ]���Φ���Ȥ�s��:[ " & lngMainCustId & "] ���@���~�Ƚs,�Ьd��!!", vbExclamation, "�T��..": Exit Sub
        
        strMainCustName = rsTmp("CustName") & ""
        If Not GetRS(rsTmp, "SELECT * FROM " & GetOwner & "SO003 WHERE CustId='" & lngMainCustId & "' And ServiceType = '" & gServiceType & "'", gcnGi) Then Exit Sub
        If rsTmp.RecordCount = 0 Then MsgBox "���Ȥ�s���L�g���ʦ��O����!!", vbExclamation, gimsgPrompt
        MenuEnabled , , , , , True
        frmSO1100BMDI.pic1.Enabled = False
        
        frmSO1131B.Show , frmSO1100BMDI
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "MainCustGo")
End Sub

Private Sub GrdR_FMT()
  On Error GoTo ChkErr
    Dim mFlds As New GiGridFlds
    With mFlds
        .Add "CitemCode", , , , , "�N�X ", vbRightJustify
        .Add "CitemName", , , , , "���O����" & Space(12), vbLeftJustify
        .Add "Period", , , , , "���� ", vbRightJustify
        .Add "Amount", , , , , "���B        ", vbRightJustify
        .Add "StartDate", giControlTypeDate, , , , "���Ĵ����_�l��", vbCenter
        .Add "StopDate", giControlTypeDate, , , , "���Ĵ����I���", vbCenter
        .Add "ClctDate", giControlTypeDate, , , , "  �����O��  ", vbCenter
        .Add "UpdEn", , , , , "���ʤH��" & Space(6), vbLeftJustify
        .Add "UpdTime", giControlTypeTime, , , , "���ʮɶ�" & Space(11), vbCenter
    End With
    ggrData.AllFields = mFlds
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "GrdR_FMT")
End Sub

Private Sub OpenData()
  On Error GoTo ChkErr
    lblMainCustId = lngMainCustId
    lblMainCustName = strMainCustName
    If Not GetRS(rs, "Select * From " & GetOwner & "SO003 Where CustId = " & lngMainCustId, gcnGi) Then Exit Sub
    Call GrdR_FMT
    Set ggrData.Recordset = rs
    ggrData.Refresh
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "OpenData")
End Sub

Private Sub cmdCancel_Click()
  On Error GoTo ChkErr
    Unload Me
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "cmdCancel_Click")
End Sub

Private Sub Form_Activate()
    On Error Resume Next
        frmSO1100BMDI.pic1.Enabled = False
        Set ObjActiveForm = Me
        strActFrmName = "frmSO1131B"
        MenuEnabled , , , , , True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo ChkErr
        If Not ChkGiList(KeyCode) Then Exit Sub
        If KeyCode = vbKeyEscape Then Unload Me
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "Form_KeyDown")
End Sub

Private Sub Form_Load()
    On Error GoTo ChkErr
        If Not800600 Then
            Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2 + 1800
        End If
        Call OpenData
    Exit Sub
ChkErr:
    ErrSub Me.Name, "Form_Load"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
        Call FormQueryUnload
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
        ReleaseCOM frmSO1131B

End Sub
