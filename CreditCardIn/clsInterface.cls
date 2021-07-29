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

'INI檔案路徑
Public Property Let ExportFormat(ByVal vData As Long)
    objStorePara.ExportFormat = vData
End Property
Public Property Let FilePath(ByVal vData As String)
    objStorePara.FilePath = vData
End Property
'入帳日期
Public Property Let RealDate(ByVal vData As Variant)
    objStorePara.RealDate = vData
End Property
'原始檔案路徑
Public Property Let SourcePath(ByVal vData As Variant)
    objStorePara.SourcePath = vData
End Property
'異動人員
Public Property Let UpdEn(ByVal vData As Variant)
    objStorePara.UpdEn = vData
End Property
'異動人員
Public Property Let UpdName(ByVal vData As Variant)
    objStorePara.UpdName = vData
End Property
'轉帳程式名稱
Public Property Let PrgName(ByVal vData As Variant)
    objStorePara.PrgName = vData
End Property
'上層的Connection
Public Property Set ugcnGi(ByVal vData As ADODB.Connection)
    Set objStorePara.ugcnGi = vData
End Property
'銀行名稱
Public Property Let BankName(ByVal vData As Variant)
    objStorePara.BankName = vData
End Property
'付款種類
Public Property Let PTCode(ByVal vData As Variant)
    objStorePara.PTCode = vData
End Property
'付款種類名稱
Public Property Let PTName(ByVal vData As Variant)
    objStorePara.PTName = vData
End Property
'服務種類
Public Property Let ServiceType(ByVal vData As Variant)
    objStorePara.ServiceType = vData
End Property
'公司別
Public Property Let CompCode(ByVal vData As Variant)
    objStorePara.CompCode = vData
End Property
'記錄是否轉帳成功
Public Property Get UpdState() As Boolean
     UpdState = objStorePara.UpdState
End Property
'銀行編號
Public Property Let BankId(ByVal vData As String)
      objStorePara.BankId = vData
End Property
'收費項目
Public Property Let uCitemQty(ByVal vData As String)
    objStorePara.uCitemQry = vData
End Property
'收費人員
Public Property Let uClctEn(ByVal vData As Variant)
    objStorePara.uClctEn = vData
End Property
Public Property Let uClctName(ByVal vData As Variant)
    objStorePara.uClctName = vData
End Property
Public Property Let pOwnerName(ByVal vData As String)
      frmUnion.TableOwnerName = vData
      GetOwner = vData
      strOwnerName = vData
End Property

'付費方式
Public Property Let uCMCode(ByVal vData As Variant)
    objStorePara.uCMCode = vData
End Property
Public Property Let uCMName(ByVal vData As Variant)
    objStorePara.uCMName = vData
End Property
Public Property Let uFubonIntegrate(ByVal vData As Boolean)
    objStorePara.uFubonIntegrate = vData
End Property
Public Function InitObject(ByVal clsName As String) As Object
    On Error GoTo ChkErr
        Dim objCls As Object
        Set objCls = CreateObject("CreditCardIn." & clsName)
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
