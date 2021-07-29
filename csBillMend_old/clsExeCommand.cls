VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsExeCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private blnShowForm As Boolean  '是否用Form展現
Private blnCloseRelate As Boolean
Private Const FormName As String = "clsExeCommand"

Public Enum ChangeDB
    giAccessDb = 0 'Access Table
    giOracle = 1    'Oracle Table
End Enum
Enum giArrow
    giLeft = 0
    giRight = 1
    giMiddle = 2
End Enum
Public Enum giOpenFileType
    giCreateTEXT = 0
    giOpenTEXT = 1
End Enum
Public Enum GiDTtpye
    GiDate = 1
    GiTime = 2
    GiYM = 3
    GiYear = 4
    GiMonth = 5
    GiDay = 6
    GiHour = 7
    GiMinute = 8
    GiSecond = 9
    GiMD = 10
End Enum
Public Enum giEditModeEnu
    giEditModeView = 0      ' 0=顯示模式
    giEditModeEdit = 1      ' 1.編輯模式(EditMode)
    giEditModeInsert = 2    ' 2=新增模式
    giEditModeDelete = 3
End Enum
Public Enum giVType
    giLongV = 0
    giStringV = 1
    giDateV = 2
End Enum
Public Property Get uExeOK() As Boolean

End Property
Public Property Let uSelectAll(ByVal vData As Boolean)
    blnSelectAll = vData
End Property
Public Property Let uCustId(ByVal vData As String)
    gCustId = Val(vData)
End Property
Public Property Get uRowIds() As String
    uRowIds = strRowIds
End Property
Public Property Get uBillNos() As String
    uBillNos = strRetBillNo
End Property
Public Property Let uCloseDate(ByVal vData As String)
    strCloseDate = vData
End Property
Public Property Let uShowForm(ByVal vData As Boolean)
    blnShowForm = vData
End Property
Public Property Get uShowForm() As Boolean
    uShowForm = blnShowForm
End Property
Public Property Let uCloseRelate(ByVal vData As Boolean)
  On Error GoTo ChkErr
  blnCloseRelate = vData
  Exit Property
ChkErr:
  ErrSub FormName, "uCloseRelate"
End Property
Public Property Let uCompCode(ByVal vData As String)
    gCompCode = vData
End Property
Public Property Set uSourceRS(ByRef vData As Recordset)
  On Error GoTo ChkErr
    If Not vData Is Nothing Then
        Set rsCopyClone = vData
        blnNoDefine = True
    End If
  Exit Property
ChkErr:
  ErrSub FormName, "uRs"
End Property
Public Property Let uTrans(ByVal vData As Boolean)
    blnTrans = vData
End Property
Public Property Set uConn(ByRef vData As ADODB.Connection)
  On Error Resume Next
    Set gcnGi = vData
End Property
Public Property Get uConn() As ADODB.Connection
    Set uConn = gcnGi
End Property
Public Property Let ugaryGi(ByVal vData As String)
  Call GetGlobal(vData)
End Property
Public Property Set uRsCopy(ByRef vData As ADODB.Recordset)
  Set rsCopyClone = vData
  blnNoDefine = True
End Property
Public Property Let uErrPath(ByVal vData As String)
    gErrLogPath = vData
End Property
Public Function Exe_UI_SF_Account(ByRef aSourceRs As ADODB.Recordset, _
    ByRef aMsg As String, ByRef aTmpSO032 As String, _
    Optional ByVal intGenOverdue As Integer = 0, _
    Optional ByVal strClctStopDate As String = Empty, _
    Optional ByVal strPeriod As String = Empty) As Boolean
    Dim aRet As Boolean
    aRet = RunCmdAccount(aSourceRs, aMsg, aTmpSO032, intGenOverdue, strClctStopDate, strPeriod)
    Exe_UI_SF_Account = aRet
End Function
Public Function EXE_DropTmpTable() As Boolean
    On Error Resume Next
    DelTmpTable
    EXE_DropTmpTable = True
End Function
Public Function Exe_UI_SF_TMPTOC1(ByRef aMsg As String, Optional ByVal aTmpSO032 As String = Empty) As Boolean
    If aTmpSO032 <> Empty Then strTmpSO032 = aTmpSO032
    Exe_UI_SF_TMPTOC1 = EXE_SF_TMPTOC1(strTmpSO032, aMsg)
End Function
Public Function GetDefineRs(Optional ByVal aCitemCodes As String = Empty) As ADODB.Recordset
    Set GetDefineRs = DefineRs(aCitemCodes)
End Function
Public Function GetShowSO1100U2(ByVal aTmpSO032 As String) As Boolean
   GetShowSO1100U2 = ShowSO1100U2(aTmpSO032)
End Function
Public Function FinallProcess() As Boolean
  FinallProcess = FinProcess()
End Function
Private Sub Class_Initialize()
    blnShowForm = True
    strRetBillNo = Empty
    strRowIds = Empty
End Sub

Public Function EXE_NOUI_COMMAND(ByRef aMsg As String, _
            Optional ByVal strCustId As String = Empty, Optional ByVal strCompId As String = Empty, _
            Optional ByVal intGenOverdue As Integer = 0, Optional ByVal aCitemCode As String = Empty, _
            Optional ByVal strPeriod As String = Empty, Optional ByVal strClctStopDate As String = Empty, _
            Optional ByVal strNewCloseDate As String = Empty, _
            Optional ByRef cnSource As Object = Nothing, Optional ByRef rsSource As Object = Nothing, _
            Optional ByRef aRetRecord As Object = Nothing) As Boolean
    blnSelectAll = True
    blnShowForm = False
    Dim aTmpSO032 As String
    EXE_NOUI_COMMAND = False
    If strCustId <> Empty Then gCustId = Val(strCustId)
    If strCompId <> Empty Then gCompCode = strCompId
    If strNewCloseDate <> Empty Then strCloseDate = strNewCloseDate
    If Not cnSource Is Nothing Then
        Set gcnGi = cnSource
    End If
    If Not rsSource Is Nothing Then
        Set rsCopyClone = rsSource
        blnNoDefine = True
    End If
    If rsSource Is Nothing And strCustId = Empty Then
        aMsg = "無任何客編資料"
        EXE_NOUI_COMMAND = False
        Exit Function
    End If
    Call DefineRs(aCitemCode)
    'If Not blnTrans Then gcnGi.BeginTrans
    If Not Exe_UI_SF_Account(rsTmp, aMsg, aTmpSO032, intGenOverdue, strClctStopDate, strPeriod) Then
        'If Not blnTrans Then gcnGi.RollbackTrans
        Exit Function
    End If
    
    
    If aRetRecord Is Nothing Then
        If ShowSO1100U2(aTmpSO032) Then
            If Not EXE_SF_TMPTOC1(aTmpSO032, aMsg) Then
                'If Not blnTrans Then gcnGi.RollbackTrans
                Exit Function
            End If
        End If
    Else
        Set aRetRecord = GetRetRecordSet(aTmpSO032)
    End If
    If Not FinallProcess Then
        If Not blnTrans Then
            'gcnGi.RollbackTrans
            'Call DelTmpTable
        End If
        Exit Function
    End If
    'DelTmpTable
    EXE_NOUI_COMMAND = True
    'gcnGi.CommitTrans
End Function

Private Sub Class_Terminate()
  On Error Resume Next
  CloseRecordset rsCopyClone
  CloseRecordset rsSO0304
  strTmpSO032 = Empty
  strRowIds = Empty
  strRetBillNo = Empty
  strCloseDate = Empty
  strAryName = Empty
  strAryName2 = Empty
  strNewClctStopDate = Empty
  strNewPeriod = Empty
  blnSelectAll = False
  blnNoDefine = False
 
End Sub
