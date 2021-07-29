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
Dim intProcessType As Integer
'開機-關機 11-15   開頻道 21-27  其他  31-32    (多執行 Precess )
Dim strProcessChStr As String
Dim blnProcessOk As Boolean
Dim strCaption As String
Dim blnTransation As Boolean
Dim strDefaultRowId As String
Dim strOrderNo As String
Dim blnProcessLock As Boolean

Public Property Set uParentForm(ByVal vData As Object)
    On Error Resume Next
        Set objParentForm = vData
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

Public Property Let uProcessType(ByVal vData As Integer)
    On Error Resume Next
        intProcessType = vData
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

Public Property Let uOrderNo(ByVal vData As String)
    On Error Resume Next
        strOrderNo = vData
End Property

Public Sub SO1623AShow(Optional Modal As Integer, Optional Owner As Object)
    On Error Resume Next
        With frmSO1623A
            Call GetGlobal(strgaryGi)
            .uOrderNo = strOrderNo
            .uProcessChStr = strProcessChStr
            .uProcessType = intProcessType
            .uCaption = strCaption
            .uCustId = lngCustId
            .uTransationMode = blnTransation
            .uDefaultRowId = strDefaultRowId
            .Show Modal, Owner
            blnProcessOk = .uProcessOk
        End With
End Sub

Public Sub SO1625AShow(Optional Modal As Integer, Optional Owner As Object)
    On Error Resume Next
        With frmSO1625A
            Call GetGlobal(strgaryGi)
            .Show Modal
         End With
End Sub

Public Sub SO1623CShow(Optional Modal As Integer, Optional Owner As Object)
    On Error Resume Next
        With frmSO1623C
            Call GetGlobal(strgaryGi)
            .Show Modal
         End With
End Sub

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


