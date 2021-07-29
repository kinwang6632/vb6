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
Dim strPrcType As String

Dim rsParent As New ADODB.Recordset
'Private rsUpd As New ADODB.Recordset

Dim strProcessChStr As String
Dim blnProcessOk As Boolean
Dim strCaption As String
Dim blnTransation As Boolean
Dim strDefaultRowId As String
Dim strOrderNo As String
Dim blnProcessLock As Boolean
Public Property Get uCmdSeqNo() As String
  On Error Resume Next
    uCmdSeqNo = strRetCmdSeqNo
End Property
Public Property Let uExecEntry(ByVal vData As Long)
    On Error Resume Next
        lngCustId = vData
End Property

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

Public Property Let uProcessType(ByVal vData As String)
    On Error Resume Next
        strPrcType = vData
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

Public Property Let uShowFaci(ByVal vData As Boolean) '#3778 自動顯示設備記錄
    blnShowFaci = vData
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

Public Property Let uSno(ByVal vData As String)
    strSno = vData
End Property

Public Property Let uMediaBillNo(ByVal vData As String)
    strMediaBillNo = vData
End Property

Public Property Let uResvTime(ByVal vData As String)
    strResvTime = vData
End Property
Public Property Let uNewBaudRate(ByVal vData As String)
    strNewBaudRate = vData
End Property
Public Property Let uBuadRate(ByVal vData As String)
    strBaudRate = vData
End Property
Public Property Let uNewBuadRate(ByVal vData As String)
    strNewBaudRate = vData
End Property
Public Property Let uNewModelCode(ByVal vData As String)
    strNewModelCode = vData
End Property
Public Property Let uNewFaciSno(ByVal vData As String)
    strNewFaciSno = vData
End Property
Public Property Let uFaciCode(ByVal vData As String)
    strFC = vData
End Property

Public Property Let uFaciName(ByVal vData As String)
    strFN = vData
End Property

Public Property Let uFaciSno(ByVal vData As String)
    strFCsno = vData
End Property

Public Property Let uProbStopDate(ByVal vData As String)
    strProbStopDate = vData
End Property

'Public Property Let uIPnum(ByVal vData As String)
'    strIPnum = vData
'End Property

Public Property Let uDynaIPcnt(ByVal vData As String)
    strDynaIPcnt = vData
End Property

Public Property Let uFixIPcnt(ByVal vData As String)
    strFixIPcnt = vData
End Property

Public Property Let uIP(ByVal vData As String)
    strIP = vData
End Property

Public Property Let uOldIP(ByVal vData As String)
    strOldIP = vData
End Property

Public Property Let uNewIP(ByVal vData As String)
    strNewIP = vData
End Property

Public Property Let uCPE(ByVal vData As String)
    strCPE = vData
End Property

Public Property Let uNode(ByVal vData As String)
    strNode = vData
End Property

Public Property Let uOldCPE(ByVal vData As String)
    strOldCPE = vData
End Property

Public Property Set uParentRs(ByVal vData As ADODB.Recordset)
    On Error Resume Next
        Set rsParent = vData
End Property

Public Property Set uUpdRs(ByVal vData As ADODB.Recordset)
    On Error Resume Next
        Set rsUpd = vData
End Property

Public Property Let uSendKey(ByVal vData As Boolean) '#3778 判斷是否要自動按下F2
    blnSendKey = vData
End Property

Public Sub SO1623AShow(Optional Modal As Integer, Optional Owner As Object)
    On Error Resume Next
        With frmSO1623A
            UI_Mode = True
            Call GetGlobal(strgaryGi)
            .uOrderNo = strOrderNo
            .uProcessChStr = strProcessChStr
            .uProcessType = strPrcType
            .uCaption = strCaption
            .uCustId = lngCustId
            .uTransationMode = blnTransation
            .uDefaultRowId = strDefaultRowId
            Set .uParentRs = rsParent
            .Show Modal, Owner
            blnProcessOk = .uProcessOk
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
    strNewBaudRate = Empty
    strNewFaciSno = Empty
    strNewModelCode = Empty
    strNewIP = Empty
    strDynaIPcnt = Empty
    strFixIPcnt = Empty
End Sub

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

