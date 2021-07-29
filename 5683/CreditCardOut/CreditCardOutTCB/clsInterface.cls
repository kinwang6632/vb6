VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInterface"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Property Let ChargeAddr(ByVal vCharge As Boolean)
    objStorePara.ChargeAddr = vCharge
End Property

Public Property Let ChooseStr(ByVal vChoose As String)
    objStorePara.ChooseStr = vChoose
End Property

Public Property Let BankName(ByVal vBankName As String)
    objStorePara.BankName = vBankName
End Property

Public Property Let Inipath(ByVal vIniPath As String)
    objStorePara.Inipath = vIniPath
End Property

Public Property Let ErrPath(ByVal vErrPath As Variant)
    objStorePara.ErrPath = vErrPath
End Property

Public Property Set Connection(ByVal vData As ADODB.Connection)
    Set objStorePara.Connection = vData
End Property

'�J�b�b��
Public Property Let ActNo(ByVal vData As String)
    objStorePara.ActNo = vData
End Property

'�Ʒ~���N��
Public Property Let BankId(ByVal vData As String)
    objStorePara.BankId = vData
End Property

'������N��
Public Property Let CorpID(ByVal vData As String)
    objStorePara.CorpID = vData
End Property

'�{���W��
Public Property Let PrgName(ByVal vData As String)
    objStorePara.PrgName = vData
End Property

'�ϥΪ̦W��
Public Property Let UserName(ByVal vData As String)
    strUserName = vData
End Property
'���楢��
Public Property Get uUpdate() As Boolean
    uUpdate = objStorePara.uUpdate
End Property
'����ɶ�
Public Property Get uTime() As Long
    uTime = objStorePara.uTime
End Property
'�O�_�����
Public Property Get uNoData() As Boolean
    uNoData = objStorePara.uNoData
End Property
'��������I���
Public Property Let uStopDate(ByVal vData As String)
    objStorePara.uStopDate = vData
End Property

Public Property Let uSpcNo(ByVal vData As String)
    objStorePara.uSpcNo = vData
End Property

Public Function InitObject(ByVal clsName As String) As Object
    On Error GoTo ChkErr
        Dim objCls As Object
        Set objCls = CreateObject("CreditCardOut." & clsName)
        Set InitObject = objCls
    Exit Function
ChkErr:
    Err.Description = "�����l�ƥ��ѡI���ˬd�{���W�١I"
End Function

Private Sub Class_Initialize()
    On Error GoTo ChkErr
        Set objStorePara = New clsStoreParameter
    Exit Sub
ChkErr:
    ErrSub "clsInterface", "Class_Initialize"
End Sub


Public Property Let pTableOwnerName(ByVal vNewValue As String)
    TableOwnerName = vNewValue
End Property

Public Property Get uProcText() As String
    uProcText = objStorePara.uProcText
End Property
Public Property Get uProcErrText() As String
    uProcErrText = objStorePara.uProcErrText
End Property
Public Property Get uChkType() As Integer
    uChkType = objStorePara.uChkType
End Property
Public Property Let uUPDEN(ByVal vData As String)
    objStorePara.uUPDEN = vData
End Property
Public Property Let uUPDTime(ByVal vData As String)
    objStorePara.uUPDTime = vData
End Property