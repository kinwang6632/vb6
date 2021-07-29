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
Private strBankId As String
Private blnUpdate As Boolean        '更新程式是否成功
Private strCompCode As String       '公司別
Private blnCancel As Boolean        '是否按下取消鍵
Private strClctEn As String         '收費人員
Private strClctName As String
Private strCMCode As String         '收費方式
Private strCMName As String
Private strGetOwner As String       'OwnerName
Private strCitemQry As String      '收費項目
Private strFlowId As String         '流程控制
Private intRefNo As Integer         '#3527 判斷CD018的參考號 By Kin 2007/10/04
Private strUPDTime As String        '#4388 增加記錄異動時間 By Kin 2009/05/04
Friend Property Let uRefNo(ByVal vData As Integer)
    intRefNo = vData
End Property
Friend Property Get uRefNo() As Integer
    uRefNo = intRefNo
End Property
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
Public Property Let uPTCode(ByVal vData As String)
    strPTCode = vData
End Property

Public Property Get uPTCode() As String
     uPTCode = strPTCode
End Property
'付款種類名稱
Public Property Let uPTName(ByVal vData As String)
    strPTName = vData
End Property

Public Property Get uPTName() As String
     uPTName = strPTName
End Property
'服務種類
Public Property Let ServiceType(ByVal vData As String)
    strServiceType = vData
End Property

Public Property Get ServiceType() As String
     ServiceType = strServiceType
End Property
'銀行名稱
Public Property Let BankName(ByVal vData As String)
     strBankName = vData
End Property
'銀行代碼
Public Property Let BankId(ByVal vData As String)
    strBankId = vData
End Property
Public Property Get BankId() As String
    BankId = strBankId
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

'收費人員
Public Property Get uClctEn() As String
    uClctEn = strClctEn
End Property

Public Property Let uClctEn(ByVal vData As String)
    strClctEn = vData
End Property
'收費人員名稱
Public Property Get uClctName() As String
    uClctName = strClctName
End Property

Public Property Let uClctName(ByVal vData As String)
    strClctName = vData
End Property

'收費方式
Public Property Get uCMCode() As String
    uCMCode = strCMCode
End Property

Public Property Let uCMCode(ByVal vData As String)
    strCMCode = vData
End Property
'收費方式名稱
Public Property Get uCMName() As String
    uCMName = strCMName
End Property

Public Property Let uCMName(ByVal vData As String)
    strCMName = vData
End Property
'取得OwnerName
Public Property Let uGetOwner(ByVal vData As String)
    strGetOwner = vData
    strOwnerName = vData
End Property

Friend Property Get uGetOwner() As String
    uGetOwner = strGetOwner
End Property

'收費項目
Public Property Get uCitemQry() As String
    uCitemQry = strCitemQry
End Property

Public Property Let uCitemQry(ByVal vData As String)
    strCitemQry = vData
End Property

'流程控制
Public Property Get FlowId() As String
     FlowId = strFlowId
End Property

Public Property Let FlowId(ByVal vData As String)
     strFlowId = vData
End Property
Public Property Let uUPDTime(ByVal vData As String)
    strUPDTime = vData
End Property
Public Property Get uUPDTime() As String
    uUPDTime = strUPDTime
End Property

