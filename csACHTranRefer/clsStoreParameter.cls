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
Private strErrPath As Variant
Private strChooseSQL As Variant     'INI檔案路徑
Private strFilePath As Variant      '檔案路徑
Private strRealDate As Variant      '入帳日期
Private strUpdEn As Variant         '記錄異動人員
Private strUpdName As Variant       '記錄異動人員
Private strPTCode As Variant        '付款種類
Private strPTName As Variant
Private strServiceType As Variant   '服務類別
Private strBankSQL As Variant       '銀行別SQL
Private strCompCode As Variant      '公司別
Private strClctEn As Variant        '收費人員
Private strClctName As Variant
Private strCMCode As Variant        '收費方式
Private strCMName As Variant
'Private strType As Variant         '性質別 1.通知 2.結果
Private blnBeOne As Boolean         '合併
Private strBankHand As String
Private strChoose33 As String
Private strGetOwner As String       'OwnerName
Private strCitemQry As String       '收費項目
Private intFlowId As Integer         '流程控制
'Private strOwnerName As String
Private blnCkZero As Integer

'INI檔案路徑
Public Property Let uChooseSQL(ByVal vData As Variant)
    strChooseSQL = vData
End Property
Public Property Get uChooseSQL() As Variant
    uChooseSQL = strChooseSQL
End Property

'INI檔案路徑
Public Property Let uChoose33(ByVal vData As Variant)
    strChoose33 = vData
End Property
Public Property Get uChoose33() As Variant
    uChoose33 = strChoose33
End Property


'檔案路徑
Public Property Let uFilePath(ByVal vData As Variant)
    strFilePath = vData
End Property
Public Property Get uFilePath() As Variant
    uFilePath = strFilePath
End Property


'INI檔案路徑
Public Property Let uErrPath(ByVal vData As Variant)
    strErrPath = vData
End Property
Public Property Get uErrPath() As Variant
    uErrPath = strErrPath
End Property



'入帳日期
Public Property Let uRealDate(ByVal vData As Variant)
    strRealDate = vData
End Property
Public Property Get uRealDate() As Variant
    uRealDate = strRealDate
End Property


'異動人員代碼
Public Property Let uUpdEn(ByVal vData As Variant)
    strUpdEn = vData
End Property
Public Property Get uUpdEn() As Variant
    uUpdEn = strUpdEn
End Property


'異動人員名稱
Public Property Let uUpdName(ByVal vData As Variant)
    strUpdName = vData
End Property
Public Property Get uUpdName() As Variant
    uUpdName = strUpdName
End Property


'上層的Connection
Public Property Set ugcnGi(ByVal vData As ADODB.Connection)
    Set gcnGi = vData
End Property

'上層的Connection
Public Property Set uCnn(ByVal vData As ADODB.Connection)
    Set cnn = vData
End Property


'銀行代碼SQL
Public Property Let uBankSQL(ByVal vData As Variant)
    strBankSQL = vData
End Property
Public Property Get uBankSQL() As Variant
    uBankSQL = strBankSQL
End Property


'付款種類
Public Property Let uPTCode(ByVal vData As Variant)
    strPTCode = vData
End Property
Public Property Get uPTCode() As Variant
    uPTCode = strPTCode
End Property
Public Property Let uPTName(ByVal vData As Variant)
    strPTName = vData
End Property
Public Property Get uPTName() As Variant
    uPTName = strPTName
End Property

'付費方式
Public Property Let uCMCode(ByVal vData As Variant)
    strCMCode = vData
End Property
Public Property Get uCMCode() As Variant
    uCMCode = strCMCode
End Property
Public Property Let uCMName(ByVal vData As Variant)
    strCMName = vData
End Property
Public Property Get uCMName() As Variant
    uCMName = strCMName
End Property



'服務種類
Public Property Let uServiceType(ByVal vData As Variant)
    strServiceType = vData
End Property
Public Property Get uServiceType() As Variant
    uServiceType = strServiceType
End Property


'公司別
Public Property Let uCompCode(ByVal vData As Variant)
    strCompCode = vData
End Property
Public Property Get uCompCode() As Variant
    uCompCode = strCompCode
End Property


'收費人員
Public Property Let uClctEn(ByVal vData As Variant)
    strClctEn = vData
End Property
Public Property Get uClctEn() As Variant
    uClctEn = strClctEn
End Property
Public Property Let uClctName(ByVal vData As Variant)
    strClctName = vData
End Property
Public Property Get uClctName() As Variant
    uClctName = strClctName
End Property

'合併檔案
Public Property Let uBeOne(ByVal vData As Boolean)
    blnBeOne = vData
End Property
Public Property Get uBeOne() As Boolean
    uBeOne = blnBeOne
End Property

'主要銀行
Public Property Let uBankHand(ByVal vData As String)
    strBankHand = vData
End Property
Public Property Get uBankHand() As String
    uBankHand = strBankHand
End Property
'OwnerName
Public Property Let uGetOwner(ByVal vData As String)
    strGetOwner = vData
    strOwnerName = vData
End Property
Public Property Get uGetOwner() As String
    uGetOwner = strGetOwner
End Property
'收費項目
Public Property Let uCitemQry(ByVal vData As String)
    strCitemQry = vData
End Property

Public Property Get uCitemQry() As String
    uCitemQry = strCitemQry
End Property

'流程控制
Public Property Let FlowId(ByVal vData As Integer)
    intFlowId = vData
End Property

Public Property Get FlowId() As Integer
    FlowId = intFlowId
End Property


'
'
''性質別 1.通知 2.結果
'Public Property Let uType(ByVal vData As Variant)
'    strType = vData
'End Property
'Public Property Get uType() As Variant
'    uType = strType
'End Property
'
'負項金額是否產生
Public Property Let uCkZero(ByVal vData As Boolean)
    blnCkZero = vData
End Property

Public Property Get uCkZero() As Boolean
    uCkZero = blnCkZero
End Property

