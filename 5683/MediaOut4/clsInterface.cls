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
Private strProcText As String
'#5683 �W�[ú�����O
Public Property Let uPayType(ByVal uData As Integer)
    objStorePara.uPayType = Val(uData)
End Property

Public Property Let uBID(ByVal vData As String)
  objStorePara.uBID = vData
End Property
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

Public Property Let ErrPath(ByVal vErrPath As String)
    objStorePara.ErrPath = vErrPath
End Property

Public Property Set Connection(ByVal vData As ADODB.Connection)
    Set objStorePara.Connection = vData
End Property

Public Property Set MDBConn(ByVal vData As ADODB.Connection)
    Set objStorePara.MDBConn = vData
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
'���q�O
Public Property Let uCompCode(ByVal vData As String)
    objStorePara.uCompCode = vData
End Property
'�A�����O
Public Property Let uServiceType(ByVal vData As String)
    objStorePara.uServiceType = vData
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
'SO033������
Public Property Let uChoose33(ByVal vData As String)
    objStorePara.uChoose33 = vData
End Property

Public Function InitObject(ByVal clsName As String) As Object
    On Error GoTo ChkErr
        Dim objCls As Object
        Set objCls = CreateObject("MediaOut4." & clsName)
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
Public Property Let uGetOwner(ByVal vData As String)
    objStorePara.uGetOwner = vData
End Property

'�y�{����
Public Property Let FlowId(ByVal vData As String)
    objStorePara.FlowId = vData
End Property

'�h�b�Უ�ͨ̾ڳ]�w
Public Property Let uBillHeadFmt(ByVal vData As String)
    objStorePara.uBillHeadFmt = vData
End Property
'#3527 �s�W�l���榡�A�H�ѦҸ��P�O(�ϥ�Pos2) By Kin 2007/10/04
Public Property Let uRefNo(ByVal vData As Integer)
    objStorePara.uRefNo = vData
End Property
Public Property Get uProcText() As String
    uProcText = objStorePara.uProcText
End Property
Public Property Get uProcErrText() As String
    uProcErrText = objStorePara.uProcErrText
End Property
'#4388 �W�[���ʤH�� By Kin 2009/04/29
Public Property Let uUpdEn(ByVal vData As String)
    objStorePara.uUpdEn = vData
End Property