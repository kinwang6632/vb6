VERSION 5.00
Begin VB.Form frmSO3311G 
   Caption         =   "���O��/������ֳt�n���T�{[SO3311G]"
   ClientHeight    =   2220
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5340
   Icon            =   "SO3311G.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   5340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '���ݵ�������
   Begin VB.CommandButton cmdEnd 
      Caption         =   "��󥻦��n��"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   180
      TabIndex        =   2
      Top             =   1380
      Width           =   2235
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "�������x�s�����n��(&X)"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2850
      TabIndex        =   1
      Top             =   1380
      Width           =   2235
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "��^�ֳt�n���e��"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1500
      TabIndex        =   0
      Top             =   180
      Width           =   2235
   End
End
Attribute VB_Name = "frmSO3311G"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    On Error Resume Next
        frmSO3311E.uCho = False
        Unload Me
End Sub

Private Sub cmdEnd_Click()
    On Error Resume Next
    Dim intCount As Long
        frmSO3311E.uCho = True
        gcnGi.Execute "Delete From " & GetOwner & "So077 Where EntryEn='" & garyGi(0) & "'", intCount
        If intCount > 0 Then MsgBox "�@���� " & IIf(intCount > 0, intCount, 0) & " ����Ƶn���I", vbExclamation, "�T���I"
        'MsgBox "�@�R��" & IIf(intCount > 0, intCount, 0) & "���I", vbExclamation, "�T���I"
        Unload Me
        Unload frmSO331AF
End Sub

Private Sub cmdOk_Click()
    On Error GoTo ChkErr
        Dim clsAlterCharge As Object
        Dim clsInterface As Object
        Dim lngAffected As Long, lngProblem As Long
        
        Set clsInterface = CreateObject("csAlterCharge4.clsInterface")
        Set cnn = GetTmpMdbCn
        With clsInterface
            .uOwnerName = GetOwner
            Set .uConn = gcnGi
            Set clsAlterCharge = .objAction("csAlterCharge4.clsAlterCharge")
        End With
        
        'set clsAlterCharge = CreateObject("csAlterCharge.clsAlterCharge")
        On Error Resume Next
        Screen.MousePointer = vbHourglass
        Set clsAlterCharge.ucnMDB = cnn
        On Error GoTo ChkErr
        Call clsAlterCharge.AlterFromSo3311G(garyGi(0), frmSO3311A.chkOtpion3, frmSO3311A.chkOption1, frmSO3311A.chkOption2, lngAffected, , garyGi(1), lngProblem)
        '�C�L���D���
        Call PrintLogData
        If lngAffected + lngProblem > 0 Then MsgBox "�@��s���\ " & lngAffected & "  ��!! ���`��� " & lngProblem & " ��!!", vbExclamation, gimsgPrompt
        Set clsAlterCharge = Nothing
        Set clsInterface = Nothing
        frmSO3311B.uCanClose = True
        frmSO3311E.uCho = True
        Screen.MousePointer = vbDefault
        Unload Me
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

Public Function Trans_SO003(Optional lngEditMode As giEditModeEnu, Optional strRowId As String, Optional lngCustId As Long, Optional intCitemCode As Long, _
                Optional strCitemName As String, Optional strStartDate As String, Optional strStopDate As String, Optional strClctDate As String, Optional strLastDate As String, _
                Optional intPeriod As Integer, Optional lngAMOUNT As Long, Optional strUpdEn As String, Optional strUpdName As String, Optional cnn As ADODB.Connection) As Boolean
    On Error GoTo ChkErr
    Dim strmsg As String
        If lngEditMode = "" Then strmsg = "EditMode": GoTo 66
        If lngCustId = "" Then strmsg = "Custid": GoTo 66
        If intCitemCode = "" Then strmsg = "intCitemCode": GoTo 66
        If strCitemName = "" Then strmsg = "strCitemName": GoTo 66
        If lngEditMode = giEditModeEdit And strRowId = "" Then strmsg = "strRowId": GoTo 66
        If cnn Is Nothing Then
            Set cnn = gcnGi
        End If
        If lngEditMode = giEditModeEdit Then
            cnn.Execute "Update " & GetOwner & "SO003 Set " & _
                             "CustId=" & lngCustId & "," & _
                             "CitemCode=" & intCitemCode & "," & _
                             "CitemName=" & GetNullValue(strCitemName) & "," & _
                             "StartDate=To_Date('" & strStartDate & "','YYYY/MM/DD')," & _
                             "StopDate=To_Date('" & strStopDate & "','YYYY/MM/DD')," & _
                             "LastDate=To_Date('" & strLastDate & "','YYYY/MM/DD')," & _
                             "ClctDate=To_Date('" & strClctDate & "','YYYY/MM/DD')," & _
                             "Amount=" & GetNullValue(lngAMOUNT) & "," & _
                             "Period=" & GetNullValue(intPeriod) & "," & _
                             "UpdEn=" & GetNullValue(strUpdEn) & "," & _
                             "UpdName=" & GetNullValue(strUpdName) & _
                             " Where Rowid='" & strRowId & "'"
        End If
        If lngEditMode = giEditModeInsert Then
            cnn.Execute "Insert Into " & GetOwner & "SO003 " & _
                              "(CustId,CitemCode,CitemName,StartDate,StopDate," & _
                              "LastDate,ClctDate,Amount,Period,UpdEn,UpdName) Values (" & _
                              lngCustId & "," & intCitemCode & ",'" & strCitemName & _
                              "',To_Date('" & strStartDate & "','YYYY/MM/DD'), " & _
                              "To_Date('" & strStopDate & "','YYYY/MM/DD'), " & _
                              "To_Date('" & strLastDate & "','YYYY/MM/DD'), " & _
                              "To_Date('" & strClctDate & "','YYYY/MM/DD'), " & _
                              GetNullValue(lngAMOUNT) & "," & _
                              GetNullValue(intPeriod) & "," & _
                              GetNullValue(strUpdEn) & "," & _
                              GetNullValue(strUpdName) & ")"
        End If
    Exit Function
66:
    MsgBox strmsg & "�L��! [ Trans_SO003 ]", vbExclamation, "����"
    Exit Function
ChkErr:
    ErrSub Me.Name, "InsertSO003"
End Function

Private Function GetNullValue(varData As Variant) As Variant
    On Error GoTo ChkErr
        If Len(varData) = 0 Then
            GetNullValue = "Null"
        Else
            GetNullValue = "'" & varData & "'"
        End If
    Exit Function
ChkErr:
    ErrSub Me.Name, "GetNullValue"
End Function

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
        Call ReleaseCOM(Me)
End Sub
