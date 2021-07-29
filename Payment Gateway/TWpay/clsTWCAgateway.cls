VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTWCAgateway"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' �ܼƦW�� : �N�q
' cavv : (optional)3D Secure��CAVV���C���׬�28�Ӧr�����r��C
' currency : ������ȥN���C���׬�3�Ӧr�����r��C(�x����901��)
' eci : ���q�l�Ӱȥ���w�����šC�@��SSL���eci�ȶ��]�w��7�C
' expDate : ���H�Υd�����Ĵ����A�榡�T�w���褸���X YYYYMM�C
' exponent : �����ȫ��ơA�s�x����0�C
' lidm : ���q�l�ӳ������ε{���ҵ�������������q��s���A��ƫ��A���̪�19�Ӧr������r��C
'          �`�N�R�q��s���r�ꪺ�r�����@���Ʀr�զX�A�ŨϥΪťդίS���r���p\/:*?"<>|%$&�����C
' merid : ��������ө���SSL HyPOS Server�����ө��s���]�w�ȡA�ѹq�l�ӳ��ۦ�]�w�C
'            �`�N�G���N�X�P�Ȧ�ұ»P��MerchantID���P�C
' order_desc : (optional)�q�满���A�̪���125�Ӧr���C
'                     �`�N�R�q��s���r�ꪺ�r�����@���Ʀr�զX�A�ŨϥΪťդίS���r���p\/:*?"<>|%$&�����C
' pan : ���H�Υd���A�d���᥽�T�X�Ȫ��[�b16�X�d����C
' purchAmt : �����O�̦���������ʶR�ӫ~�����v�`���B�A����ƫ��A�C
' server : HyPOS���A����URL�A�榡��http://ip:port�C(�phttp://127.0.0.1:2011)
' timeout : �]�w����HyPOS���A���^�Ǫ��O�ɮɶ��A��쬰���A�Ytimeout�p��ε���0�ɡA�htimeout�h�]�w���w�]��30���C

' [OUT]
' AUTH_OUT���c�R
' �ܼƦW�� : �N�q
' amount : ���v���B�C
' authcode : ������v�X�A�̤j���׬�6���r��C
' authrrpid : SSL ���v������N�X�A�̪��� 40 �Ӧ줸��r��C
' BATCHCLOSE : �p�Gbatchclose��1�A���ܧ妸�v���A���������妸��i�i�����C�p�G�妸�����A�hbatchclose��0�C
' currency : ������ȥN���C���׬�3�Ӧr�����r��C(�x����901��)
' err_code : ���~�N�X�A���׬�2���r��A�аѾ\�����G�C
' err_status : ���檬�A�A���Ʀr�κA�C(�䤤�A0 �����榨�\�A�䥦�U�����A�аѾ\�����G)
' exponent : �����ȫ��ơA�s�x����0�C
' retrref : �վ\�s���A�̪���12�Ӧ줸��r��C
' termseq : �վ\�Ǹ� , ����ƫ��A
' xid : ����ѧO�X�A�̪���40�Ӧ줸��r��C

' �禡�W�� : ����
' Auth : �i��H�Υd���v�ʧ@�C
' GetAuthcode : ���o���v������v�X�C
' GetAuthrrpid : ���o���v������N�X�C
' GetAmount : ���o���v���B�C
' GetBatchclose : �妸�������A���ܡC
' GetCurrency : ���o������ȥN���C
' GetErrStatus : ���o��������檬�A�C
' GetErrCode : ���o��������~�N�X�C
' GetExponent : ���o���ȫ��ơC
' GetRettref : ���o������վ\�s���C
' GetTermseq : ���o������վ\�Ǹ��C
' GetXID : ���o����ѧO�X�C
' SetCAVV : �]�w3D Secure��CAVV�ȡC
' SetOrderDesc : �]�w�q�满���C
' SetTimeout : �]�w����O�ɮɶ��C

' COM����һ��ɮ�
' posapi_c.dll : COM���󪺥D�n�ʺA�s����
' posapi_d.dll : pos api���D�n�ʺA�s����
' posapi_d.lib : posapi_d.dll��lib��
' posapi_d.h : posapi_d.dll���禡�w�q��
' hy_err.h : posapi_d.dll�����~�X
' libeay32.dll : OpenSSL�禡�w
' ssleay32.dll : OpenSSL�禡�w

' CLSID : C6A2E34F-B5D6-43B4-AF03-C068B5D2EE80
' ProgId:  posapi_c.Auth

Private strSvrUrl As String

Private strLIDM As String ' �q��s��
Private strMerID As String ' �ө��s��
Private strPan As String ' �H�Υd�����T�X��
Private strExpDate As String ' �H�Υd���Ĵ���
Private strCurrency As String ' ������ȥN��
Private lngPurchAmt As Long ' ������v�`���B
Private intExponent As Integer ' ���ȫ���
Private intEci As Integer ' �q�l�Ӱȥ���w������

Private Const FrmName = "clsTWCAgateway"

Private objAPI As Object ' API ����
Private lngRet As Long ' �^�ǭ�
Private api As New POSAPI_CLib.AUTH

Private strErrMsg As String
Private lngErrStatus As Long
Private strErrCode As String

' Function GetAmount() As Long
Public Property Get GetAmount() As Long
    GetAmount = objAPI.GetAmount
End Property

' Function GetAuthcode() As String
Public Property Get GetAuthcode() As String
    GetAuthcode = objAPI.GetAuthcode
End Property

' Function GetAuthrrpid() As String
Public Property Get GetAuthrrpid() As String
    GetAuthrrpid = objAPI.GetAuthrrpid
End Property

' Function GetBatchclose() As Long
Public Property Get GetBatchclose() As Long
    GetBatchclose = objAPI.GetBatchclose
End Property

' Function GetCurrency() As String
Public Property Get GetCurrency() As String
    GetCurrency = objAPI.GetCurrency
End Property

' Function GetErrCode() As String
Public Property Get GetErrCode() As String
    GetErrCode = objAPI.GetErrCode
End Property

' Function GetErrStatus() As Long
Public Property Get GetErrStatus() As Long
    GetErrStatus = objAPI.GetErrStatus
End Property

' Function GetExponent() As Long
Public Property Get GetExponent() As Long
    GetExponent = objAPI.GetExponent
End Property

' Function GetRetrref() As String
Public Property Get GetRetrref() As String
    GetRetrref = objAPI.GetRetrref
End Property

' Function GetTermseq() As Long
Public Property Get GetTermseq() As Long
    GetTermseq = objAPI.GetTermseq
End Property

' Function GetXID() As String
Public Property Get GetXID() As String
    GetXID = objAPI.GetXID
End Property

' Function SetBirthday(birthday As String) As Long
' Function SetCAVV(cavv As String) As Long
' Function SetOrderDesc(order_desc As String) As Long
' Function SetPID(pid As String) As Long
' Sub SetTimeout(timeout As Long)

' �� �v
Public Function Approve(strOrderNo As String, strCardNo As String, strDate As String, lngAmt As Long)
  On Error GoTo ChkErr
    
    strLIDM = "12345678"
    strPan = "4444444444444444"
    strExpDate = "200808"
    lngPurchAmt = 1000
    
    With objAPI

'        .SetBirthday "20040212" ' Optional Setting
'        .SetPID "P123456789" ' Optional Setting

        .SetOrderDesc "Authorize Testing" ' Optional Setting �q�满��
        
        .SetTimeout 10 ' Optional Setting ����O��
        
        ' Function AUTH(server As String, merid As Long, lidm As String, pan As String, expDate As String,
        '                        currency As String, purchAmt As Long, exponent As Long, eci As Long) As Long
        lngRet = .AUTH(strSvrUrl, strMerID, strLIDM, strPan, strExpDate, _
                                strCurrency, lngPurchAmt, intExponent, intEci)
    
        lngErrStatus = .GetErrStatus
        strErrCode = .GetErrCode()
        
        If lngRet = 0 Then
            If strErrCode <> 0 Then strErrMsg = "Failed to Auth." & lngErrStatus & "(" + strErrCode & ")"
        Else
            strErrMsg = "API Error = " & lngRet
        End If
    
    End With
  
  Exit Function
ChkErr:
    ErrHandle FrmName, "Approve"
End Function

Private Sub Class_Initialize()
  On Error GoTo ChkErr
  
    Set objAPI = CreateObject("PosApi_C.AUTH") ' �إ� API ����
    If Err Then strErrMsg = " API ( PosApi_C.AUTH ) ����إߥ��� !"
    
    'strSvrUrl = "https://demopg.hyweb.com.tw:3841" ' HyPOS���A����URL
    strSvrUrl = "http://localhost:2011" ' HyPOS���A����URL
    strCurrency = "901" ' �x��
    intExponent = 0 ' �s�x��
    intEci = 7 ' �@��SSL����]�w
    strMerID = 1 ' �ө��s��
    
  Exit Sub
ChkErr:
    ErrHandle FrmName, "Class_Initialize"
End Sub

Private Sub Class_Terminate()
  On Error Resume Next
    Set objAPI = Nothing ' ���� API ����
End Sub

'�P�¶Q���q�ĥ� TWPay �t�ΡA�����Q���q���t�Τu�{�v�ɳt�F�Ѧ����ըt�ξާ@�����A
'�ФU�� TWPay API �w�ˤΨ禡��X��U�A�Ш̷Ӧw�ˤ�U�w�� POS for Windows Platform
'�ñN��X�U�禡�ܶQ���q�������q�l�ӳ�
'
'�ϥΤ�U�U���ЦܡGhttp://service.twpay.com.tw/dlapii.htm
'
'�аȥ���X�U�C�U�禡 (�U�C�U�禡���T�W�٤ΩI�s�覡�A�аѦҤ�U����)
'
'�h Auth( ) �����v
'�h Auth_Recur( ) ���������v [�ȨѤ����I�ڻݭn�ϥ�]
'�h Auth_Rev( ) �������v
'�h Cap( ) �д�
'�h Batch_close () �妸���b
'
'�Ш̾ڥH�U������T�A�i����աC
'
'
'�@����
'�h MerchantID�G 0012370888051 �h TerminalID�G 91010001
'
'
'
'�h    PGURL�Ghttps://demopg.hyweb.com.tw:3841
'          �� �Q�S���T�w POS for windows �Ӧ��A��
'          �i�H�z�L TCP / 3841 Port �s�ܥ����q���� PGURL
'
'
'�Y�b���դW��������D�A�ܸ��w��z�^��email�� service_twpay@twca.com.tw
'���p���t�d���~�ȤH���A�ڭ̱N�ɳt�^���z�����D�A���¡C
'
' �����Ʒ~���n���� , ���y�A�Ȫ��̨ι٦�
'If you have any question, please do not hesitate to contact us !
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'TWPay Service Team E-mail�Gservice_twpay@twca.com.tw
'
'TEL: 02-2370-8886 �����C���p�U�G
'�ȪA�� - �i�Q�s 139
'�ȪA�� - �����p 135
'�~�ȳ� - ù�h�� 655
'FAX: 02-2370-0728
'Address: 100 �x�_�������n�� 85 �� 10 ��
'
'