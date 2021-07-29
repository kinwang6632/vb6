VERSION 5.00
Object = "{3BB22E57-824A-11D3-BD70-0080C8F80BC4}#3.4#0"; "GiGridR.ocx"
Begin VB.Form frmSO1100G1 
   BorderStyle     =   1  '單線固定
   Caption         =   "ACH授權細項查詢 [SO1100G1]"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9480
   Icon            =   "SO1100G1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   9480
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&X)"
      Height          =   435
      Left            =   8340
      TabIndex        =   6
      Top             =   4830
      Width           =   1005
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "F2.存檔"
      Enabled         =   0   'False
      Height          =   435
      Left            =   1260
      TabIndex        =   5
      Top             =   4860
      Width           =   1005
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "F11.修改"
      Enabled         =   0   'False
      Height          =   435
      Left            =   120
      TabIndex        =   4
      Top             =   4860
      Width           =   1005
   End
   Begin VB.Frame fraData 
      Height          =   4695
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   9315
      Begin VB.TextBox txtNote 
         Height          =   645
         Left            =   720
         MaxLength       =   100
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   8475
      End
      Begin prjGiGridR.GiGridR ggrData 
         Height          =   3600
         Left            =   120
         TabIndex        =   1
         Top             =   960
         Width           =   9090
         _ExtentX        =   16034
         _ExtentY        =   6350
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseCellForeColor=   -1  'True
      End
      Begin VB.Label lblNote 
         Caption         =   "備  註"
         Height          =   315
         Left            =   180
         TabIndex        =   2
         Top             =   450
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmSO1100G1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strMasterRowId As String
Private lngNowMode As giEditModeEnu
Private lngEditMode As giEditModeEnu
Private WithEvents rsSO106A As ADODB.Recordset
Attribute rsSO106A.VB_VarHelpID = -1
Private strInACHID As String
Private strInACHDesc As String
Private strFirstACHID As String
Private strFirstACHDesc As String
Private Sub cmdCancel_Click()
    On Error Resume Next
    Unload Me
End Sub
Public Property Let uMasterRowId(ByVal vData As String)
    strMasterRowId = vData
End Property
Public Property Let uNowEditMode(ByVal vData As Long)
    lngEditMode = vData
End Property
Private Sub subGrd()
    On Error GoTo ChkErr
    Dim mFlds As New prjGiGridR.GiGridFlds
        With mFlds
            .Add "RecordType", , , , , "資料狀態", vbLeftJustify
            .Add "ACHTNO", , , , , "ACH 代碼 ", vbLeftJustify
            .Add "ACHDESC", , , , , "ACH 交易說明 ", vbLeftJustify
            .Add "AuthorizeStatus", , , , , "授權狀態 ", vbLeftJustify
            .Add "CitemNameStr", , , , , "週期性項目" & Space(20), vbLeftJustify
            .Add "StopFlag", , , , , "停用", vbLeftJustify
            .Add "StopDate", giControlTypeTime, , , , "停用日期" & Space(15), vbLeftJustify
            .Add "UpdEn", , , , , "異動人員" & Space(10), vbLeftJustify
            .Add "UpdTime", , , , , "異動時間" & Space(10), vbLeftJustify
        End With
        ggrData.AllFields = mFlds
        ggrData.SetHead
        Set ggrData.Recordset = rsSO106A
        ggrData.Refresh
        If strFirstACHID <> Empty Then
            Do While Not rsSO106A.EOF
                If rsSO106A("ACHTNO") = strFirstACHID And rsSO106A("ACHDesc") = strFirstACHDesc Then Exit Do
                rsSO106A.MoveNext
            Loop
        End If
    Exit Sub
ChkErr:
    ErrSub Me.Name, "subGrd"
End Sub

Private Sub cmdEdit_Click()
  On Error GoTo ChkErr
    ChangeMode giEditModeEdit
    txtNote.SetFocus
    Exit Sub
ChkErr:
    ErrSub Me.Name, "cmdEdit_Click"
End Sub

Private Sub cmdSave_Click()
  On Error GoTo ChkErr
    Dim varBookMark As Variant
    varBookMark = rsSO106A.Bookmark
    gcnGi.BeginTrans
    If ScrToRcd Then
        gcnGi.CommitTrans
        ChangeMode giEditModeView
        rsSO106A.Requery
        Set ggrData.Recordset = rsSO106A
        ggrData.Refresh
        rsSO106A.AbsolutePosition = varBookMark
        If strInACHID <> Empty Then Unload Me
    End If
    Exit Sub
ChkErr:
    gcnGi.RollbackTrans
    ErrSub Me.Name, "cmdSave_Click"
End Sub

Private Sub Form_Activate()
  On Error Resume Next
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error GoTo ChkErr
    If KeyCode = vbKeyF11 Then
        If cmdEdit.Enabled Then
            cmdEdit.Value = True
        End If
    End If
    If KeyCode = vbKeyF2 Then
        If cmdSave.Enabled Then
            cmdSave.Value = True
        End If
    End If
    If KeyCode = vbKeyEscape Then
        If EditMode = giEditModeEdit Then
            ChangeMode giEditModeView
        ElseIf EditMode = giEditModeView Then
            cmdCancel.Value = True
        End If
    End If
  Exit Sub
ChkErr:
       Call ErrSub(Me.Name, "Form_KeyDown")
End Sub

Private Sub Form_Load()
    On Error GoTo ChkErr
    Set rsSO106A = New ADODB.Recordset
    Call OpenData
    Call subGrd
    If rsSO106A.RecordCount <= 0 Then
        cmdEdit.Enabled = False
        cmdSave.Enabled = False
        fraData.Enabled = False
        Exit Sub
    End If
    ChangeMode lngEditMode
  Exit Sub
ChkErr:
    ErrSub Me.Name, "Form_Load"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  On Error Resume Next
    If EditMode <> giEditModeView Then
        If giMsgNotSave = vbNo Then
            Cancel = True: Exit Sub
        Else
            Unload Me
        End If
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
    Call CloseRecordset(rsSO106A)
End Sub

Private Sub ggrData_ColorCellData(ByVal Col As Long, ByVal Row As Long, ByVal giFld As prjGiGridR.GiGridFld, ByVal Fld As ADODB.Field, Value As Variant)
  On Error Resume Next
    If UCase(giFld.FieldName) = "STOPFLAG" Then
        If CStr(Fld.Value) = "1" Then
            Value = vbRed
        End If
    End If

End Sub

Private Sub ggrData_ShowCellData(ByVal Col As Long, ByVal Row As Long, ByVal giFld As prjGiGridR.GiGridFld, ByVal Fld As ADODB.Field, Value As Variant)
    On Error Resume Next
    Select Case UCase(giFld.FieldName)
        Case "RECORDTYPE"
            Select Case Value
                '#3946 原本為"請授權",改為"申請授權" By Kin 2008/05/30
                Case "0"
                    Value = "申請授權"
                Case "1"
                    Value = "取消授權"
            End Select
        Case "AUTHORIZESTATUS"
            Select Case Value
                Case "1"
                    Value = "授權成功"
                Case "2"
                    Value = "已取消授權"
                Case "3"
                    Value = "授權失敗"
                '#3946 多增加狀態4 By Kin 2008/05/30
                Case "4"
                    Value = "待提出"
                Case Else
                    Value = "授權中"
            End Select
        Case "STOPFLAG"
            If Value = 1 Then
                Value = "是"
            Else
                Value = "否"
            End If
    End Select
End Sub

Private Sub rsSO106A_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  On Error Resume Next
    Call RcdToScr
End Sub
'#3946 原本以MasterRowId為PK,現在改為MasterID By Kin 2008/05/30
Private Sub OpenData()
  On Error GoTo ChkErr
    Dim strQry As String
    strQry = "Select * From " & GetOwner & "SO106A Where MasterID='" & strMasterRowId & "'" & _
            " Order by CreateTime Desc,RecordType,ACHTNO"
    If Not GetRS(rsSO106A, strQry, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
    Exit Sub
ChkErr:
    ErrSub Me.Name, "OpenData"
End Sub
Private Function ScrToRcd() As Boolean
  On Error GoTo ChkErr
    If strInACHID <> Empty Then
        Dim s As String
        gcnGi.Execute ("UPDATE " & GetOwner & "SO106A Set Notes='" & txtNote.Text & "'" & _
                        " Where ACHTNO IN(" & strInACHID & ")" & _
                        " And ACHDesc IN(" & strInACHDesc & ")" & _
                        " And MasterRowID='" & strMasterRowId & "'" & _
                        " And StopFlag=1")

    Else
        rsSO106A("NOTES") = txtNote.Text
        rsSO106A("UpdTime") = GetDTString(RightNow)
        rsSO106A("UpdEn") = garyGi(1)
        rsSO106A.Update
    End If
    ScrToRcd = True
    Exit Function
ChkErr:
    ErrSub Me.Name, "ScrToRcd"
End Function
Private Function RcdToScr() As Boolean
  On Error GoTo ChkErr
    If Not rsSO106A.EOF Then
        txtNote = rsSO106A("NOTES") & ""
    End If
    RcdToScr = True
    Exit Function
ChkErr:
    ErrSub Me.Name, "RcdToScr"
End Function
Private Sub ChangeMode(ByVal lngMode As giEditModeEnu)
  On Error GoTo ChkErr
    lngEditMode = lngMode
    cmdSave.Enabled = False
    cmdEdit.Enabled = False
    txtNote.Enabled = False
    ggrData.Enabled = True
    Select Case lngMode
        Case giEditModeView
            cmdEdit.Enabled = True
        Case giEditModeEdit
            cmdSave.Enabled = True
            txtNote.Enabled = True
            ggrData.Enabled = False
    End Select
    Exit Sub
ChkErr:
    ErrSub Me.Name, "IntialOcx"
End Sub
Public Property Get EditMode() As giEditModeEnu
  On Error GoTo ChkErr
    EditMode = lngEditMode    ' 取目前編輯模式
  Exit Property
ChkErr:
    Call ErrSub(Me.Name, "Get EditMode")
End Property

Public Property Let EditMode(ByVal vNewValue As giEditModeEnu)
  On Error GoTo ChkErr
    lngEditMode = vNewValue '設定編輯模式
  Exit Property
ChkErr:
    Call ErrSub(Me.Name, "Let EditMode")
End Property
Public Property Let uInAchId(ByVal vData As String)
  On Error GoTo ChkErr
    strInACHID = vData
    Exit Property
ChkErr:
    Call ErrSub(Me.Name, "Let uInAchId")
End Property
Public Property Let uInAchDesc(ByVal vData As String)
  On Error GoTo ChkErr
    strInACHDesc = vData
  Exit Property
ChkErr:
    Call ErrSub(Me.Name, "uInAchDesc")
End Property
Public Property Let uFirstAchId(ByVal vData As String)
  On Error GoTo ChkErr
    strFirstACHID = vData
    Exit Property
ChkErr:
    Call ErrSub(Me.Name, "uFirstAchId")
End Property
Public Property Let uFirstACHDesc(ByVal vData As String)
  On Error GoTo ChkErr
    strFirstACHDesc = vData
    Exit Property
ChkErr:
    Call ErrSub(Me.Name, "uFirstACHDesc")
End Property
