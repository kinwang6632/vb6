VERSION 5.00
Object = "{3BB22E57-824A-11D3-BD70-0080C8F80BC4}#3.4#0"; "GiGridR.ocx"
Begin VB.Form frmSO1100L 
   BorderStyle     =   1  '��u�T�w
   Caption         =   "��Ϭd�� [SO1100L]"
   ClientHeight    =   5085
   ClientLeft      =   45
   ClientTop       =   1530
   ClientWidth     =   11880
   Icon            =   "SO1100L.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   Begin prjGiGridR.GiGridR ggrData 
      Height          =   4890
      Left            =   135
      TabIndex        =   0
      Top             =   90
      Width           =   11610
      _ExtentX        =   20479
      _ExtentY        =   8625
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
End
Attribute VB_Name = "frmSO1100L"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Programmer: Power Hammer
' Last Modify: 2002/09/28
'
'cable_jim2100: so1100L��Ϭd�ߥ\��
'cable_jim2100: �٭n�[�@�����HSI Flag
'cable_lawrence: ok
'cable_jim2100: �i�H�⤤��W�٥[��CM �i�ˡACM ���i��
Option Explicit

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
        If Not800600 Then Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2 - 200
    End If
  Exit Sub
ChkErr:
    ErrSub Me.Name, "Form_Load"
End Sub

Public Sub OpenData(ByRef rsCross As Object)
  On Error GoTo ChkErr
    Dim mflds As New GiGridFlds
    With mflds
        .Add "CompName", , , , , "���q�O "
        .Add "ServiceName", , , , , "�A�ȧO "
        .Add "CustStatusName", , , , , "�Ȥ᪬�A"
        .Add "CustID", , , , , "�Ȥ�s��", vbRightJustify
        .Add "CustName", , , , , "�Ȥ�W��    "
        .Add "Tel1", , , , , "�q��(1)   "
        .Add "CircuitNo", , , , , "�����s��"
        .Add "NodeNo", , , , , "Node No"
'       .Add "HFlag", , , , , "HSI �˾��I"
        .Add "Balance", , , "###,###,###,###", , "��O   ", vbRightJustify
        .Add "InstAddress", , , , , "�˾��a�}                                "
        .Add "MduName", , , , , "�j�ӦW��"
        .Add "WipName1", , , , , "���u���A1"
        .Add "WipName2", , , , , "���u���A2"
        .Add "WipName3", , , , , "���u���A3"
        .Add "Tel2", , , , , "�q��(2)   "
        .Add "Tel3", , , , , "�q��(3)   "
    End With
    With ggrData
        .AllFields = mflds
        .SetHead
         Set .Recordset = rsCross
        .Refresh
    End With
    Me.Show 1
  Exit Sub
ChkErr:
    ErrSub Me.Name, "OpenData"
End Sub

'Private Sub ggrData_ShowCellData(ByVal Col As Long, ByVal Row As Long, ByVal giFld As prjGiGridR.GiGridFld, ByVal Fld As ADODB.Field, Value As Variant)
'  On Error GoTo ChkErr
'    If ggrData.Recordset.EOF Then Exit Sub
'    With giFld
'         If .FieldName = "HFlag" Then Value = IIf(Value = "1", "CM �i��", "CM ���i��")
'    End With
'  Exit Sub
'ChkErr:
'    ErrSub Me.Name, "ggrData_ShowCellData"
'End Sub
