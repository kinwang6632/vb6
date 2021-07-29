VERSION 5.00
Object = "{3BB22E57-824A-11D3-BD70-0080C8F80BC4}#3.4#0"; "GiGridR.ocx"
Begin VB.Form frmSO1623C 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '��u�T�w
   Caption         =   "�s���e��"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10110
   ClipControls    =   0   'False
   Icon            =   "SO1623C.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   10110
   StartUpPosition =   2  '�ù�����
   Begin VB.CommandButton cmdPrint 
      Caption         =   "F5.�C�L"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7230
      TabIndex        =   2
      Top             =   5010
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "����(&X)"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8700
      TabIndex        =   1
      Top             =   5010
      Width           =   1365
   End
   Begin prjGiGridR.GiGridR ggrView 
      Height          =   4935
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   10035
      _ExtentX        =   17701
      _ExtentY        =   8705
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
Attribute VB_Name = "frmSO1623C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rsData As New ADODB.Recordset

Private Sub cmdCancel_Click()
    On Error Resume Next
        Unload Me
End Sub

Private Sub cmdPrint_Click()
    On Error Resume Next
'        ReadyGoPrint
'        PrintRpt GetPrinterName(5), "SO3220A.rpt", , "���O�浹���a�}�վ�@���� [SO3220A]", "SELECT * From tmp007", , , True, , , , GiPaperLandscape
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
        Call InitData
        Call subGrd
        Call OpenData
    Exit Sub
ChkErr:
    ErrSub Me.Name, "Form_Load"
End Sub

Private Sub InitData()
    On Error Resume Next
        gLogInID = GetSystemParaItem("LogInID", gCompCode, "", "SO041", , True) & ""
        If gLogInID <> "" Then gLogInID = gLogInID & "."

End Sub

Private Sub OpenData()
    On Error GoTo ChkErr
        If Not GetRS(rsData, "Select SeqNo, CompCode, HIGH_LEVEL_CMD_ID, ICC_NO, STB_NO, SUBSCRIPTION_BEGIN_DATE, " & _
            "SUBSCRIPTION_END_DATE, CMD_STATUS, ProcessingDate From " & gLogInID & "SEND_NAGRA Where CompCode = " & gCompCode) Then Exit Sub
        Set ggrView.Recordset = rsData
        ggrView.Refresh
    Exit Sub
ChkErr:
    ErrSub Me.Name, "OpenData"
End Sub

Private Sub subGrd()
    On Error GoTo ChkErr
        Dim mFlds As New prjGiGridR.GiGridFlds
            With mFlds
                .Add "CompCode", , , , False, "���q�O", vbLeftJustify
                .Add "CMD_STATUS", , , , False, "�R�R���A", vbLeftJustify
                .Add "SeqNo", , , , False, "�Ǹ�   ", vbLeftJustify
                .Add "ICC_NO", , , , False, "���z�d�Ǹ�", vbLeftJustify
                .Add "STB_NO", , , , False, "STB �Ǹ�  ", vbLeftJustify
                .Add "HIGH_LEVEL_CMD_ID", , , , False, "���O�W��     ", vbLeftJustify
                .Add "SUBSCRIPTION_BEGIN_DATE", giControlTypeDate, , , False, "�����_�l��", vbLeftJustify
                .Add "SUBSCRIPTION_END_DATE", giControlTypeDate, , , False, "�����I���", vbLeftJustify
                .Add "ProcessingDate", giControlTypeTime, , , False, "�w���ɶ�        ", vbLeftJustify
            End With
        ggrView.AllFields = mFlds
        ggrView.SetHead
    Exit Sub
ChkErr:
    ErrSub Me.Name, "subGrd"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo ChkErr
    Exit Sub
ChkErr:
    ErrSub Me.Name, "Form_QueryUnload"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
        Call ReleaseCOM(Me)

End Sub

Private Sub ggrView_ShowCellData(ByVal Col As Long, ByVal Row As Long, ByVal giFld As prjGiGridR.GiGridFld, ByVal Fld As ADODB.Field, Value As Variant)
    On Error GoTo ChkErr
        Select Case UCase(Fld.Name)
            Case "HIGH_LEVEL_CMD_ID"
                Select Case Value
                    Case "A1"
                        Value = "�}��"
                    Case "A2"
                        Value = "����"
                    Case "A3"
                        Value = "����"
                    Case "A4"
                        Value = "�_��"
                    Case "A5"
                        Value = "���P"
                    Case "B1"
                        Value = "�}�W�D"
                    Case "B2"
                        Value = "���W�D(�ӧO)"
                    Case "B3"
                        Value = "���W�D(����)"
                    Case "B4"
                        Value = "�Ȱ��W�D"
                    Case "B5"
                        Value = "��_�W�D"
                    Case "B6"
                        Value = "�������i"
                    Case "B7"
                        Value = "�����v���ƻs"
                    Case "E1"
                        Value = "�]�w�ˤl�K�X"
                    Case "E2"
                        Value = "�ǰe�T��"
                End Select
            Case "CMD_STATUS"
                Select Case Value
                    Case "C"
                        Value = "����"
                    Case "W"
                        Value = "���ԳB�z"
                    Case "P"
                        Value = "�B�z��"
                    Case "E"
                        Value = "���~"
                End Select
        End Select
        
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "ggrCustRate_ShowCellDate")

End Sub
