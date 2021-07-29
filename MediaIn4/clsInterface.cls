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

Public Enum ChangeDB
    giAccessDb = 0 'Access Table
    giOracle = 1    'Oracle Table
End Enum

Enum giArrow
    giLeft = 0
    giRight = 1
    giMiddle = 2
End Enum

'INI檔案路徑
Public Property Let FilePath(ByVal vData As String)
    objStorePara.FilePath = vData
End Property
'入帳日期
Public Property Let RealDate(ByVal vData As String)
    objStorePara.RealDate = vData
End Property
'原始檔案路徑
Public Property Let SourcePath(ByVal vData As String)
    objStorePara.SourcePath = vData
End Property
'異動人員
Public Property Let UpdEn(ByVal vData As String)
    objStorePara.UpdEn = vData
End Property
'轉帳程式名稱
Public Property Let PrgName(ByVal vData As String)
    objStorePara.PrgName = vData
End Property
'上層的Connection
Public Property Set ugcnGi(ByVal vData As ADODB.Connection)
    Set objStorePara.ugcnGi = vData
End Property
'銀行名稱
Public Property Let BankName(ByVal vData As String)
    objStorePara.BankName = vData
End Property
'銀行代碼
Property Let BankId(ByVal vData As String)
    objStorePara.BankId = vData
End Property
'付款種類
Public Property Let uPTCode(ByVal vData As String)
    objStorePara.uPTCode = vData
End Property
'付款種類名稱
Public Property Let uPTName(ByVal vData As String)
    objStorePara.uPTName = vData
End Property
'服務種類
Public Property Let ServiceType(ByVal vData As String)
    objStorePara.ServiceType = vData
End Property

'公司別
Public Property Let CompCode(ByVal vData As String)
    objStorePara.CompCode = vData
End Property

'記錄是否轉帳成功
Public Property Get UpdState() As Boolean
     UpdState = objStorePara.UpdState
End Property

Public Function InitObject(ByVal clsName As String) As Object
    On Error GoTo ChkErr
        Dim objCls As Object
        Set objCls = CreateObject("MediaIn4." & clsName)
        Set InitObject = objCls
    Exit Function
ChkErr:
    Err.Description = "物件初始化失敗！請檢查程式名稱！"
End Function

Private Sub Class_Initialize()
    On Error GoTo ChkErr
        Set objStorePara = New clsStoreParameter
    Exit Sub
ChkErr:
    ErrSub "clsInterface", "Class_Initialize"
End Sub
'收費人員
Public Property Let uClctEn(ByVal vData As String)
    objStorePara.uClctEn = vData
End Property
'收費人員名稱
Public Property Let uClctName(ByVal vData As String)
    objStorePara.uClctName = vData
End Property

'收費方式
Public Property Let uCMCode(ByVal vData As String)
    objStorePara.uCMCode = vData
End Property
'收費方式名稱
Public Property Let uCMName(ByVal vData As String)
    objStorePara.uCMName = vData
End Property
'OwnerName
Public Property Let uGetOwner(ByVal vData As String)
    objStorePara.uGetOwner = vData
End Property

'收費項目
Public Property Let uCitemQty(ByVal vData As String)
    objStorePara.uCitemQry = vData
End Property

'流程控制
Public Property Let FlowId(ByVal vData As String)
    objStorePara.FlowId = vData
End Property

Public Property Let uRefNo(ByVal vData As Integer)
    objStorePara.uRefNo = vData
End Property
Public Property Let uUPDTime(ByVal vData As String)
    objStorePara.uUPDTime = vData
End Property
