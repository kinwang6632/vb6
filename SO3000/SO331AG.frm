VERSION 5.00
Begin VB.Form frmSO331AG 
   Caption         =   "派工單所附收費資料登錄結束[SO331AG]"
   ClientHeight    =   2190
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5385
   Icon            =   "SO331AG.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   5385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton cmdEnd 
      Caption         =   "放棄本次登錄"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   180
      TabIndex        =   2
      Top             =   1470
      Width           =   2205
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束並儲存本次登錄(&X)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   3060
      TabIndex        =   1
      Top             =   1470
      Width           =   2205
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "返回登錄畫面"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1530
      TabIndex        =   0
      Top             =   180
      Width           =   2205
   End
End
Attribute VB_Name = "frmSO331AG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdCancel_Click()
    frmSO331AA.uCanClose = False
    Unload Me
End Sub

Private Sub cmdEnd_Click()
On Error GoTo ChkErr
Dim intCount As Long
    'If ExecuteSQL("Delete So077 Where EntryEn='" & garyGi(0) & "'", gcnGi, intCount) <> giOK Then MsgBox "連線失敗！請重新操作！", vbExclamation, "訊息！"
    gcnGi.Execute "Delete From " & GetOwner & "So074 Where EntryEn='" & garyGi(0) & "'", intCount
    If intCount > 0 Then MsgBox "共取消 " & IIf(intCount > 0, intCount, 0) & " 筆資料登錄！", vbExclamation, "訊息！"
    Unload Me
    frmSO331AA.uCanClose = True
    Unload frmSO331AA
Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "cmdEnd_Click")
End Sub

Private Sub cmdOk_Click()
    On Error GoTo ChkErr
        Dim clsAlterCharge As Object
        Dim clsInterface As Object
        Dim lngAffected As Long, lngProblem As Long
        
        Set clsInterface = CreateObject("csAlterCharge4.clsInterface")
        Set cnn = GetTmpMdbCn
        With clsInterface
            Set .uConn = gcnGi
            .uOwnerName = GetOwner
            Set clsAlterCharge = .objAction("csAlterCharge4.clsAlterCharge")
        End With
        'set clsAlterCharge = CreateObject("csAlterCharge.clsAlterCharge")
        Screen.MousePointer = vbHourglass
        On Error Resume Next
        Set clsAlterCharge.ucnMDB = cnn
        On Error GoTo ChkErr
        
        Call clsAlterCharge.AlterFromSo331AG(garyGi(0), True, True, True, lngAffected, , garyGi(1), lngProblem)
        '列印問題資料
        Call PrintLogData
        If lngAffected + lngProblem > 0 Then MsgBox "共更新成功 " & lngAffected & "  筆!! 異常資料 " & lngProblem & " 筆!!", vbExclamation, gimsgPrompt
        Set clsAlterCharge = Nothing
        Set clsInterface = Nothing
        Screen.MousePointer = vbDefault
        Unload Me
        frmSO331AA.uCanClose = True
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "cmdCancel_Click")
    Exit Sub
End Sub

Private Sub PrintLogData()
    On Error GoTo ChkErr
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
ChkErr:
    ErrSub Me.Name, "PrintLogData"
End Sub


Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
        Call ReleaseCOM(Me)

End Sub
