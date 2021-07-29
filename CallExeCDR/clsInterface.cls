VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Private Const FormName = "clsInterface"

Public Enum ChangeDB
    giAccessDb = 0 'Access Table
    giOracle = 1    'Oracle Table
End Enum

Public Enum giVType
    giLongV = 0
    giStringV = 1
    giDateV = 2
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

Public Enum giEditModeEnu
    giEditModeView = 0      ' 0=顯示模式
    giEditModeEdit = 1      ' 1.編輯模式(EditMode)
    giEditModeInsert = 2    ' 2=新增模式
    giEditModeDelete = 3
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

Public Enum csErrType
    csErrCustStatus = 0
    csErrCustPR = 1
    csErrFaciPRSNo = 2
    csErrExcelField = 3
    csErrCustId = 4
    csErrFaciSno = 5
End Enum

Public Property Let uOwnerName(ByVal vErrPath As Variant)
  On Error Resume Next
    clsStoreParameter.uOwnerName = vErrPath
End Property

Public Property Get uErr() As Variant
    On Error Resume Next
        uErr = clsStoreParameter.uErr
End Property

Public Property Let uErrPath(ByVal vErrPath As Variant)
    On Error Resume Next
        clsStoreParameter.uErrPath = vErrPath
End Property

Public Property Set uConn(ByVal vConn As ADODB.Connection)
    On Error Resume Next
        Set clsStoreParameter.uConn = vConn
End Property

Public Property Let ugaryGi(ByVal vData As Variant)
    On Error Resume Next
        clsStoreParameter.ugaryGi = vData
End Property

Public Function objAction(objName As String) As Object
    On Error GoTo ChkErr
        Set objAction = CreateObject(objName)
    Exit Function
ChkErr:
    ErrSub FormName, "objAction"
End Function

Private Sub Class_Initialize()
    On Error Resume Next
    Set clsStoreParameter = New clsStoreParameter
End Sub

