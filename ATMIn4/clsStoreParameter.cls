VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStoreParameter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private strFilePath As String       'INI檔案路徑
Private strRealDate As String       '入帳日期
Private strSourcePath As String     '原始檔案路徑
Private strUpdEn As String          '記錄異動人員
Private strPrgName As String        '轉帳程式名稱
Private strPTCode As String         '付款種類
Private strPTName As String         '付款種類名稱
Private strServiceType As String    '服務類別
Private strBankName As String       '銀行名稱
Private intBankId As Integer
Private blnUpdate As Boolean        '更新程式是否成功
Private strCompCode As String       '公司別
Private strGetOwner As String       'OwnerName
Private blnCancel As Boolean        '是否按下取消鍵

'INI檔案路徑
Public Property Let FilePath(ByVal vData As String)
    strFilePath = vData
End Property

Public Property Get FilePath() As String
     FilePath = strFilePath
End Property
'入帳日期
Public Property Let RealDate(ByVal vData As String)
    strRealDate = vData
End Property

Public Property Get RealDate() As String
     RealDate = strRealDate
End Property
'原始檔案路徑
Public Property Let SourcePath(ByVal vData As String)
    strSourcePath = vData
End Property

Public Property Get SourcePath() As String
     SourcePath = strSourcePath
End Property
'記錄異動人員
Public Property Let UpdEn(ByVal vData As String)
    strUpdEn = vData
End Property

Public Property Get UpdEn() As String
     UpdEn = strUpdEn
End Property
'轉帳程式名稱
Public Property Let PrgName(ByVal vData As String)
    strPrgName = vData
End Property

Public Property Get PrgName() As String
     PrgName = strPrgName
End Property
'上層的Connection
Public Property Set ugcnGi(ByVal vData As ADODB.Connection)
    Set gcnGi = vData
End Property
'付款種類
Public Property Let PTCode(ByVal vData As String)
    strPTCode = vData
End Property

Public Property Get PTCode() As String
     PTCode = strPTCode
End Property
'付款種類名稱
Public Property Let PTName(ByVal vData As String)
    strPTName = vData
End Property

Public Property Get PTName() As String
     PTName = strPTName
End Property
'服務種類
Public Property Let ServiceType(ByVal vData As String)
    strServiceType = vData
End Property

Public Property Get ServiceType() As String
     ServiceType = strServiceType
End Property
'銀行名稱
Public Property Let BankID(ByVal vData As Integer)
     intBankId = vData
End Property

Public Property Get BankID() As Integer
     BankID = intBankId
End Property
'銀行名稱
Public Property Let BankName(ByVal vData As String)
     strBankName = vData
End Property

Public Property Get BankName() As String
     BankName = strBankName
End Property
'記錄是否轉帳成功
Public Property Get UpdState() As Boolean
     UpdState = blnUpdate
End Property

Public Property Let UpdState(ByVal vData As Boolean)
     blnUpdate = vData
End Property
'公司別
Public Property Get CompCode() As String
    CompCode = strCompCode
End Property

Public Property Let CompCode(ByVal vData As String)
    strCompCode = vData
End Property
'取得OwnerName
Public Property Let uGetOwner(ByVal vData As String)
    strGetOwner = vData
    strOwnerName = vData
     garyGi(16) = vData
End Property

Friend Property Get uGetOwner() As String
    uGetOwner = strGetOwner
End Property


