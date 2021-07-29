VERSION 5.00
Begin VB.Form frmSO3311D 
   BorderStyle     =   1  '單線固定
   Caption         =   "登錄結束 [SO3311D]"
   ClientHeight    =   2205
   ClientLeft      =   45
   ClientTop       =   4620
   ClientWidth     =   4530
   BeginProperty Font 
      Name            =   "新細明體"
      Size            =   9.75
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SO3311D.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   4530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '螢幕中央
   Begin VB.CheckBox chkSave 
      Caption         =   "暫存本次輸入記錄"
      Height          =   195
      Left            =   930
      TabIndex        =   1
      Top             =   840
      Width           =   1995
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "返回登錄畫面"
      Default         =   -1  'True
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   1470
      Width           =   1335
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "結束(&X)"
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   1470
      Width           =   1335
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "列印本次輸入記錄"
      Height          =   375
      Left            =   930
      TabIndex        =   0
      Top             =   210
      Width           =   2535
   End
End
Attribute VB_Name = "frmSO3311D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
    On Error GoTo chkErr
        Dim clsAlterCharge As Object
        Dim clsInterface As Object
        Dim lngAffected As Long, lngProblem As Long
        If chkSave.Value = 1 Then
            Unload Me
            frmSO3311B.uCanClose = True
            Exit Sub
        End If
        Screen.MousePointer = vbHourglass
        Set cnn = GetTmpMdbCn
        Set clsInterface = CreateObject("csAlterCharge4.clsInterface")
        With clsInterface
            .uOwnerName = GetOwner
            Set .uConn = gcnGi
            Set clsAlterCharge = .objAction("csAlterCharge4.clsAlterCharge")
        End With
        'set clsAlterCharge = CreateObject("csAlterCharge.clsAlterCharge")
        On Error Resume Next
        Set clsAlterCharge.ucnMDB = cnn
        On Error GoTo chkErr
        Call clsAlterCharge.AlterFromSo3311D(garyGi(0), frmSO3311A.chkOtpion3, frmSO3311A.chkOption1, frmSO3311A.chkOption2, lngAffected, , garyGi(1), lngProblem)
        '列印問題資料
        Call PrintLogData
        If lngAffected + lngProblem > 0 Then MsgBox "共更新成功 " & lngAffected & "  筆!! 異常資料 " & lngProblem & " 筆!!", vbExclamation, gimsgPrompt
        Set clsAlterCharge = Nothing
        Set clsInterface = Nothing
        Screen.MousePointer = vbDefault
        Unload Me
        frmSO3311B.uCanClose = True
    Exit Sub
chkErr:
    Call ErrSub(Me.Name, "cmdCancel_Click")
End Sub

Private Sub PrintLogData()
    On Error GoTo chkErr
    Dim rsTmp As New ADODB.Recordset
        If Not GetRS(rsTmp, "Select * From SO07xLog", cnn) Then Exit Sub
        If rsTmp.RecordCount > 0 Then
            With ViewForm
                .uIsSo07x = True
                .uRecordset = rsTmp
                .Show vbModal
            End With
        End If
    Exit Sub
chkErr:
    ErrSub Me.Name, "PrintLogData"
End Sub

Private Sub cmdCancel_Click()
    frmSO3311B.uCanClose = False
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
        If Not ChkGiList(KeyCode) Then Exit Sub
        If KeyCode = vbKeyEscape And cmdOK.Enabled Then cmdOK.Value = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
        Call ReleaseCOM(Me)

End Sub
