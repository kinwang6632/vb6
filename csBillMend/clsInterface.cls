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
Public Property Let Inipath(ByVal vIniPath As String)
    objStorePara.Inipath = vIniPath
End Property

Public Property Let ErrPath(ByVal vErrPath As String)
    objStorePara.ErrPath = vErrPath
End Property

Public Property Set Connection(ByVal vData As ADODB.Connection)
    Set objStorePara.Connection = vData
End Property

Public Property Set MDBConn(ByVal vData As ADODB.Connection)
    Set objStorePara.MDBConn = vData
End Property

'使用者名稱
Public Property Let UserName(ByVal vData As String)
    strUserName = vData
End Property
'公司別
Public Property Let uCompCode(ByVal vData As String)
    objStorePara.uCompCode = vData
End Property
'服務類別
Public Property Let uServiceType(ByVal vData As String)
    objStorePara.uServiceType = vData
End Property
'執行失敗
Public Property Get uUpdate() As Boolean
    uUpdate = objStorePara.uUpdate
End Property
'執行時間
Public Property Get uTime() As Long
    uTime = objStorePara.uTime
End Property
'是否有資料
Public Property Get uNoData() As Boolean
    uNoData = objStorePara.uNoData
End Property


Public Function InitObject(ByVal clsName As String) As Object
    On Error GoTo ChkErr
        'Dim objCls As Object
        'Set objCls = CreateObject("MediaOut4." & clsName)
        'Set InitObject = objCls
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
Public Property Let uGetOwner(ByVal vData As String)
    objStorePara.uGetOwner = vData
End Property

'流程控制
Public Property Let FlowId(ByVal vData As String)
    objStorePara.FlowId = vData
End Property

'多帳戶產生依據設定
Public Property Let uBillHeadFmt(ByVal vData As String)
    objStorePara.uBillHeadFmt = vData
End Property
'#3527 新增郵局格式，以參考號判別(使用Pos2) By Kin 2007/10/04
Public Property Let uRefNo(ByVal vData As Integer)
    objStorePara.uRefNo = vData
End Property
Public Property Get uProcText() As String
    uProcText = objStorePara.uProcText
End Property
Public Property Get uProcErrText() As String
    uProcErrText = objStorePara.uProcErrText
End Property
'#4388 增加異動人員 By Kin 2009/04/29
Public Property Let uUpdEn(ByVal vData As String)
    objStorePara.uUpdEn = vData
End Property
