VERSION 5.00
Object = "{3BB22E57-824A-11D3-BD70-0080C8F80BC4}#3.4#0"; "GiGridR.ocx"
Begin VB.Form frmSO1131I 
   BorderStyle     =   1  '��u�T�w
   Caption         =   "�X�����{�s�� [SO1131I]"
   ClientHeight    =   2805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11625
   Icon            =   "SO1131I.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   11625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '�ù�����
   Begin VB.CommandButton cmdCancel 
      Caption         =   "���}(&X)"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   2340
      Width           =   1005
   End
   Begin prjGiGridR.GiGridR ggrData 
      Height          =   2205
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   3889
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�s�ө���"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseCellForeColor=   -1  'True
   End
End
Attribute VB_Name = "frmSO1131I"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rsData As New ADODB.Recordset
Private strServiceType As String


Private Sub cmdCancel_Click()
  On Error Resume Next
  Unload Me
End Sub

Private Sub Form_Activate()
  On Error Resume Next
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error GoTo ChkErr
    If KeyCode = vbKeyEscape Then
        cmdCancel.Value = True
    End If
    Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "Form_KeyDown")
End Sub

Private Sub Form_Load()
  On Error GoTo ChkErr
    Call subGrd
    Call OpenData
    Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "Form_Load")
End Sub
Private Sub OpenData()
  On Error GoTo ChkErr
    Dim strQry As String
'    strQry = "Select A.OrderNo,Min(A.ContStartDate) ContStartDate,Max(A.ContStopDate) ContStopDate,C.BPName,C.PromName, C.CitemName,B.BundleMon, C.PenalType,C.expiretype" & _
'             " FROM " & GetOwner & "SO003A A," & GetOwner & "CD078A B," & GetOwner & "SO003 C" & _
'             " Where A.CustID = C.CustID" & _
'             " AND A.ServiceType = C.ServiceType" & _
'             " AND A.ContNo = C.ContNo" & _
'             " AND A.CompCode=C.CompCode " & _
'             " And A.ServiceType='" & gServiceType & "'" & _
'             " And A.Custid=" & gCustId & _
'             " And A.CompCode=" & gCompCode & _
'             " And C.BpCode=B.BpCode" & _
'             " Group By A.OrderNo,C.BPName,C.PromName," & _
'             " B.BundleMon,C.CitemName,C.PenalType, C.expiretype"
    strQry = "Select A.OrderNo,Min(A.ContStartDate) ContStartDate,Max(A.ContStopDate) ContStopDate,A.PromName BPName,C.PromName,B.BundleMon" & _
             " FROM " & GetOwner & "SO003A A," & GetOwner & "CD078A B," & GetOwner & "SO003 C" & _
             " Where A.CustID = C.CustID" & _
             " AND A.ServiceType = C.ServiceType" & _
             " AND A.ContNo = C.ContNo" & _
             " AND A.CompCode=C.CompCode " & _
             " And A.ServiceType='" & strServiceType & "'" & _
             " And A.Custid=" & gCustId & _
             " And A.CompCode=" & gCompCode & _
             " And C.BpCode=B.BpCode" & _
             " Group By A.OrderNo,A.PromName,C.PromName," & _
             " B.BundleMon"
    
    If Not GetRS(rsData, strQry, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
    Set ggrData.Recordset = rsData
    ggrData.Refresh
    Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "OpenData")
  
End Sub
Private Sub subGrd()
    On Error Resume Next
    Dim mFlds As New GiGridFlds
    With mFlds
        .Add "OrderNo", , , , , "�q��s��" & Space(12)
        .Add "ContStartDate", , , , , "�X���_��"
        .Add "ContStopDate", , , , , "�X������"
        .Add "BPName", , , , , "�u�f�զX" & Space(40)
        .Add "PromName", , , , , "�P�P���" & Space(40)
        .Add "BundleMon", , , , , "�j�����"
        '.Add "PenalType", , , , , "�H���ɭp���̾�"
        '.Add "expiretype", , , , , "�u�f����p���̾�"
    End With
    ggrData.AllFields = mFlds
    ggrData.SetHead
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  On Error Resume Next
  Call Close3Recordset(rsData)
End Sub

Private Sub ggrData_ShowCellData(ByVal Col As Long, ByVal Row As Long, ByVal giFld As prjGiGridR.GiGridFld, ByVal Fld As ADODB.Field, Value As Variant)
  On Error Resume Next
    If UCase(giFld.FieldName) = "PENALTYPE" Then
          Select Case Value
              Case 1
                  Value = "�^���P��"
              Case 2
                  Value = "�u�f��"
              Case 3
                  Value = "�u�f������"
          End Select
    End If
    If UCase(giFld.FieldName) = "EXPIRETYPE" Then
        Select Case Value
            Case 1
                Value = "��_�̷s���i�P��"
            Case 2
                Value = "�H�̫�@���~���u�f"
            Case 3
                Value = "������X"
            Case 4
                Value = "���s��"
        End Select
    End If
End Sub

Public Property Get uServiceType() As String
    uServiceType = strServiceType
End Property

Public Property Let uServiceType(ByVal vData As String)
    strServiceType = vData
End Property
