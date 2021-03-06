VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCMControl"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim strgaryGi As String
Dim objParentForm As Object
Dim lngCustId As Long
Dim strProcessType As String
Dim rsParent As New ADODB.Recordset
Dim strProcessChStr As String
Dim blnProcessOk As Boolean
Dim strCaption As String
Dim blnTransation As Boolean
Dim strDefaultRowId As String
Dim strOrderNo As String
Dim blnProcessLock As Boolean
Dim blnAutoSend As Boolean
'Private strOrdChannelID As String
'Private strCancelChannelID As String
Private strSetCitemCode As String
Private strCancelCitemCode As String
Public Property Let uSeqNo(ByVal vData As String)
    strSeqNo = vData
End Property
Public Property Let uIsCancelCmd(ByVal vData As Boolean)
    IsCancelCmd = vData
End Property
Public Property Let uSNO(ByVal vData As String)
    str7SNO = vData
End Property
Public Property Let uMediaBillNo(ByVal vData As String)
    str3MediaBillNo = vData
End Property
Public Property Let uReasonCode(ByVal vData As String)
    strReasonCode = vData
End Property
Public Property Let uReasonName(ByVal vData As String)
    strReasonName = vData
End Property

Public Property Let uSendKey(ByVal vData As Boolean) '#3778 判斷是否要自動按下F2
    blnSendKey = vData
End Property

Public Property Let uShowFaci(ByVal vData As Boolean) '#3778 自動顯示設備記錄
    blnShowFaci = vData
End Property
'Public Property Let uOrdChannelID(ByVal vData As String)
'    strOrdChannelID = vData
'End Property
'Public Property Let uCancelChannelID(ByVal vData As String)
'    strCancelChannelID = vData
'End Property

Public Property Let uProcessingDate(ByVal vData As String)
    strResvTime = vData
End Property
Public Property Let uCitemCode(ByVal vData As String)
    strSetCitemCode = vData
End Property
Public Property Let uCancelCitemCode(ByVal vData As String)
    strCancelCitemCode = vData
End Property
Public Property Let uAutoSend(ByVal vData As Boolean)
  On Error Resume Next
    blnAutoSend = vData
    frmSO1623A.uAutoSend = vData
End Property

Public Property Let uFaciCode(ByVal vData As String)
  On Error Resume Next
    'frmSO1623A.uFaciCode = vData
    strFaciCode = vData
End Property
Public Property Set uParentForm(ByVal vData As Object)
    On Error Resume Next
        Set objParentForm = vData
End Property
Public Property Let uFaciSno(ByVal vData As String)
    On Error Resume Next
    'frmSO1623A.uFaciSno = vData
    strFaciSNo = vData
End Property

Public Property Set uConn(ByVal vData As ADODB.Connection)
    On Error Resume Next
        Set gcnGi = vData
End Property

Public Property Set uConnMDB(ByVal vData As ADODB.Connection)
    On Error Resume Next
        Set cnn = vData
End Property

Public Property Let uCompCodeFilter(ByVal vData As String)
    On Error Resume Next
        gCompFilterStr = vData
End Property

Public Property Let uCompCode(ByVal vData As String)
    On Error Resume Next
        gCompCode = vData
End Property

Public Property Let uServiceType(ByVal vData As String)
  On Error Resume Next
    gServiceType = vData
End Property

Public Property Let ugaryGi(ByVal vData As String)
    On Error Resume Next
        strgaryGi = vData
End Property

Public Property Let uDefaultComp(ByVal vData As String)
    On Error Resume Next
        gDefaultComp = vData
End Property

Public Property Let uCustId(ByVal vData As Long)
    On Error Resume Next
        lngCustId = vData
End Property

Public Property Let uErrPath(ByVal vData As String)
    On Error Resume Next
        gErrLogPath = vData
End Property

Public Property Let uProcessLock(ByVal vData As Boolean)
    On Error Resume Next
        blnProcessLock = vData
End Property

Public Property Let uProcessType(ByVal vData As String)
    On Error Resume Next
        strProcessType = vData
End Property

Public Property Get uProcessOk() As Boolean
    On Error Resume Next
        uProcessOk = blnProcessOk
End Property

Public Property Let uProcessChStr(ByVal vData As String)
    On Error Resume Next
        strProcessChStr = vData
End Property

Public Property Let uCaption(ByVal vData As String)
    On Error Resume Next
        strCaption = vData
End Property

Public Property Let uTransationMode(ByVal vData As Boolean)
    On Error Resume Next
        blnTransation = vData
End Property

Public Property Let uDefaultRowId(ByVal vData As String)
    On Error Resume Next
        strDefaultRowId = vData
End Property
'Public Property Set uUpdRs(ByRef vData As ADODB.Recordset)
'    On Error Resume Next
'    Set rsUpdSo004 = vData
'End Property
Public Property Set uParentRs(ByVal vData As ADODB.Recordset)
    On Error Resume Next
        Set rsParent = vData
        
End Property

Public Property Let uOrderNo(ByVal vData As String)
    On Error Resume Next
        strOrderNo = vData
End Property

Public Property Let uCPno(ByVal strCP As String)
    On Error Resume Next
        strCPno = strCP
End Property

Public Property Let uAccChk(ByVal blnChk As Boolean)
    On Error Resume Next
        blnAccChk = blnChk
End Property
Public Property Set NewRs(ByRef vData As ADODB.Recordset)
  On Error Resume Next
    Set rs2 = vData
End Property

Public Sub SO1623AShow(Optional Modal As Integer, Optional Owner As Object)
    On Error Resume Next
        With frmSO1623A
            Call GetGlobal(strgaryGi)
            .uOrderNo = strOrderNo
            .uProcessChStr = strProcessChStr
            .uProcessType = strProcessType
            .uProcessLock = blnProcessLock
            .uCaption = strCaption
            .uCustId = lngCustId
            .uTransationMode = blnTransation
            .uDefaultRowId = strDefaultRowId
            .uOrdChannelID = strSetCitemCode
            .uCancelChannelID = strCancelCitemCode
            Set .uParentRs = rsParent
            If Not blnAutoSend Then
                .Show 1, Owner
            Else
                .cmdSet.Value = True
            End If
            blnProcessOk = .uProcessOk
        End With
End Sub

'
'Public Sub SO1625AShow(Optional Modal As Integer, Optional Owner As Object)
'    On Error Resume Next
'        With frmSO1625A
'            Call GetGlobal(strgaryGi)
'            .Show Modal
'         End With
'End Sub
'
'Public Sub SO1623CShow(Optional Modal As Integer, Optional Owner As Object)
'    On Error Resume Next
'        With frmSO1623C
'            Call GetGlobal(strgaryGi)
'            .Show Modal
'         End With
'End Sub

Private Sub FormShow(objForm As Object, _
    Optional Modal As Integer, Optional Owner As Object)
    With objForm
        If Owner Is Nothing Then Set Owner = objParentForm
        Call GetGlobal(strgaryGi)
        .Show Modal, objParentForm
    End With
End Sub

Private Sub Class_Terminate()
    'objParentForm.SetFocus
End Sub


