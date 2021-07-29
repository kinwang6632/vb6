VERSION 5.00
Object = "{3BB22E57-824A-11D3-BD70-0080C8F80BC4}#3.4#0"; "GiGridR.ocx"
Begin VB.Form frmSO1100U2 
   BorderStyle     =   1  '��u�T�w
   Caption         =   "�L��T�{ [SO1100U2]"
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11070
   Icon            =   "SO1100U2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   11070
   StartUpPosition =   2  '�ù�����
   Begin VB.CommandButton cmdCancel 
      Caption         =   "����(&X)"
      Height          =   375
      Left            =   9870
      TabIndex        =   2
      Top             =   3540
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "F2.�T�w"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   3540
      Width           =   975
   End
   Begin prjGiGridR.GiGridR ggrData 
      Height          =   3360
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   10860
      _ExtentX        =   19156
      _ExtentY        =   5927
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�s�ө���"
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
Attribute VB_Name = "frmSO1100U2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rsTmp As New ADODB.Recordset
Private strTmpFile As String
Private Sub cmdCancel_Click()
  On Error Resume Next
    frmSO1100U.uOK = False
    Unload Me
End Sub
Private Sub subGrid()
    On Error Resume Next
    Dim mFlds As New GiGridFlds
    With mFlds
        .Add "ServiceType", , , , , "�A�ȧO"
        .Add "BillNo", , , , , "��ڽs��" & Space(10)
        .Add "Item", , , , , "����"
        .Add "FaciSNO", , , , , "�]�ƧǸ�" & Space(8)
        .Add "CitemName", , , , , "���O����             "
        .Add "ShouldDate", , , , , "�������"
        .Add "ShouldAmt", , , , , "�������B" & Space(5)
        .Add "RealStartDate", , , , , "�_�l���" & Space(5)
        .Add "RealStopDate", , , , , "�I����" & Space(5)
    End With
    ggrData.AllFields = mFlds
    ggrData.SetHead
    Set ggrData.Recordset = rsTmp
    ggrData.Refresh
End Sub

Private Sub cmdOK_Click()
  On Error Resume Next
    frmSO1100U.uOK = True
    Unload Me
End Sub

Private Sub Form_Activate()
  On Error Resume Next
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error GoTo ChkErr
    If KeyCode = vbKeyF2 Then If cmdOK.Enabled Then cmdOK.Value = True
    If KeyCode = vbKeyCancel Then If cmdCancel.Enabled Then cmdCancel.Value = True
    Exit Sub
ChkErr:
  ErrSub Me.Name, "Form_KeyDown"
End Sub

Private Sub Form_Load()
  On Error GoTo ChkErr
    Call OpenData
    Call subGrid
    Exit Sub
ChkErr:
  ErrSub Me.Name, "Form_Load"
End Sub
Private Sub OpenData()
  On Error GoTo ChkErr
    Dim strQry As String
    strQry = "Select * From " & strTmpFile
    If Not GetRS(rsTmp, strQry, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
    Exit Sub
ChkErr:
  ErrSub Me.Name, "OpenData"
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  On Error Resume Next
    CloseRecordset rsTmp
End Sub
Property Let uTmpFile(ByVal vData As String)
    strTmpFile = vData
End Property
