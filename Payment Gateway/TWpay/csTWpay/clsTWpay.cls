VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTWpay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Doc : CITI-NPG-POS_API_��U_V1_06_�@��P����__960115.pdf

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

'COM API ����ҥ]�t���ɮ�
' posapi_c_npg.dll ���COM ���󪺥D�n�ʺA�s����
' posapi_d.dll pos api ���D�n�ʺA�s����
' libeay32.dll OpenSSL �禡�w
' ssleay32.dll OpenSSL �禡�w

' CLSID : C6A2E34F-B5D6-43B4-AF03-C068B5D2EE80
' ProgId:  posapi_c.Auth

Private strSvrUrl As String

Private strLIDM As String ' �q��s��
Private strMerID As String ' �ө��s��
Private strPan As String ' ���H�Υd���A�d���᥽�T�X�Ȫ��[�b16�X�d����C
Private strExpDate As String ' �H�Υd���Ĵ���
Private strCurrency As String ' ������ȥN��
Private lngPurchAmt As Long ' ������v�`���B
Private intExponent As Integer ' ���ȫ���
Private intECI As Integer ' �q�l�Ӱȥ���w������
Private intTimeOut As Integer ' Time Out

Private Const FrmName = "clsTWCAgateway"

Private objAPI As Object ' API ����
Private strRet As String ' �^�ǭ�
'Private api As New POSAPI_CLib.AUTH

Private lngErrStatus As Long
Private strErrCode As String

' ===��������ݩ�========================================================================

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

' Function GetBatchId() As String
Public Property Get GetBatchId() As String
    GetBatchId = objAPI.GetBatchId
End Property

' Function GetBatchSeq() As String
Public Property Get GetBatchSeq() As String
    GetBatchSeq = objAPI.GetBatchSeq
End Property

' Function GetCloseDateTime() As String
Public Property Get GetCloseDateTime() As String
    GetCloseDateTime = objAPI.GetCloseDateTime
End Property

' =====================================================================================

'AUTH:  �H�Υd���v�@�~
'AUTH_RECUR:  �H�Υd�����I�ڱ��v�@�~ (���A�ȵ�����Ȧ�O�_����)
'AUTH_REV:  ���� �H�Υd���v�@�~(�Y�����q��)
'CAP:  �H�Υd�дڧ@�~
'CAP_REV:   ���� �H�Υd�дڧ@�~
'Cred:   �i��h�ڰʧ@
'CredRev:   �i������h�f�ʧ@
'BATCHCLOSE:   �ө��P����Ȧ�妸���b (�������@�~�Ȧ�~�|���b�ڵ��Q�S��)

Public Function GetECI(strCardNo As String) As Integer
  On Error GoTo ChkErr
    Select Case Left(strCardNo, 1)
        Case 3, 4
            GetECI = 7
        Case 5
            GetECI = 0
        Case Else
            GetECI = 7
    End Select
  Exit Function
ChkErr:
    ErrHandle FrmName, "GetECI"
End Function

' ���U���� Jacky �g�� �ŷs��� Payment  Gateway �� Interface

' ���v
'Public Function Approve(strPostString As Variant, strRequestHeader As Variant, _
    strUserId As Variant, strPasswd As Variant, strMerchantNumber As Variant, _
    strOrderNumber As Variant, lngAmount As Variant, lngDepositFlag As Variant) As Long

' ���v
Public Function Approve(strHyPOS As String, _
                                        strCurrencyType As String, _
                                        intExponentType As Integer, _
                                        intECItype As Integer, _
                                        strMerID As String, _
                                        strOrderNo As String, _
                                        strCardNo As String, _
                                        strDate As String, _
                                        lngAmt As Long) As Long
  On Error GoTo ChkErr
  
    Dim lngAPIreturn As Long
    
    ' Function SetBirthday(birthday As String) As Long
    ' Function SetCAVV(cavv As String) As Long
    ' Function SetOrderDesc(order_desc As String) As Long
    ' Function SetPID(pid As String) As Long
    ' Sub SetTimeout(timeout As Long)

    InitializePara strHyPOS, strCurrencyType, intExponentType, intECItype, strMerID
    
    'If AUTH(strOrderNo, strCardNo, Format(strDate, "YYYYMM"), lngAmt, lngAPIreturn) Then
    If AUTH(strOrderNo, strCardNo, strDate, lngAmt, lngAPIreturn) Then
        Approve = 0
    Else
        Approve = objAPI.GetErrStatus
        'Approve = -1
    End If
    
  Exit Function
ChkErr:
    ErrHandle FrmName, "Approve"
End Function

' �H�Υd���v�@�~
Public Function AUTH(strOrderNo As String, strCardNo As String, strDate As String, lngAmt As Long, _
                                Optional ByRef lngAPIreturn As Long = -1) As Boolean
  On Error GoTo ChkErr
  
'    strLIDM = "12345678" ' �q��s��
'    strPan = "4444444444444444" ' ���H�Υd���A�d���᥽�T�X�Ȫ��[�b16�X�d����C
'    strExpDate = "200808" ' �H�Υd���Ĵ���
'    lngPurchAmt = 1000 ' ������v�`���B
    
    strLIDM = strOrderNo ' �q��s��
    strPan = strCardNo ' ���H�Υd���A�d���᥽�T�X�Ȫ��[�b16�X�d����C
    strExpDate = strDate ' �H�Υd���Ĵ���
    lngPurchAmt = lngAmt ' ������v�`���B

'    Set objAPI = CreateObject("PosApi_C.AUTH") ' �إ� API ����
    Set objAPI = CreateObject("posapi_c_npg.AUTH") ' �إ� API ����
'    If Err Then strErrMsg = " API ( PosApi_C.AUTH ) ����إߥ��� !"
    
    With objAPI

'        .SetBirthday "20040212" ' Optional Setting
'        .SetPID "P123456789" ' Optional Setting
'        .SetOrderDesc "Authorize Testing" ' Optional Setting �q�满��
        '.SetTimeout 10 ' Optional Setting ����O��
        ' Function AUTH(server As String, merid As Long, lidm As String, pan As String, expDate As String,
        '                        currency As String, purchAmt As Long, exponent As Long, eci As Long) As Long
        
        intECI = GetECI(strCardNo)
        
        strRet = .AUTH(strSvrUrl, strMerID, strLIDM, strPan, strExpDate, _
                                strCurrency, lngPurchAmt, intExponent, intECI)
        
        lngAPIreturn = strRet
        ' GetErrStatus : ���o��������檬�A�C
        ' �Ǧ^�� : ���檬�A�A0�����榨�\�A�D0���䥦���~���A�аѾ\�����G�C
        lngErrStatus = .GetErrStatus
        
        ' GetErrCode : ���o��������~�N�X�C
        ' �����檬�A����0�ɡA�i�Q��GetErrCode�禡���o������������~�N�X�C
        strErrCode = .GetErrCode
        
        If strRet = 0 Then
            'If strErrCode <> 0 Then strErrMsg = "Failed to Auth." & lngErrStatus & "(" & strErrCode & ")"
            If lngErrStatus = 0 And strErrCode = "00" Then
                AUTH = True
'                Debug.Print "���o����ѧO�X GetXID : " & .GetXID ' ���o����ѧO�X�C
'                Debug.Print "���o���v������N�X GetAuthrrpid : " & .GetAuthrrpid ' ���o���v������N�X�C
'                Debug.Print "���o������վ\�Ǹ� GetTermseq : " & .GetTermseq ' ���o������վ\�Ǹ��C
'                Debug.Print "���o������վ\�s�� GetRetrref : " & .GetRetrref ' ���o������վ\�s���C
'                Debug.Print "�妸�������A���� GetBatchclose : " & .GetBatchclose ' �妸�������A���ܡC
'                Debug.Print "���o���v������v�X GetAuthcode : " & .GetAuthcode ' ���o���v������v�X�C
'                Debug.Print "���o���v���B GetAmount : " & .GetAmount ' ���o���v���B�C
'                Debug.Print "���o������ȥN�� GetCurrency : " & .GetCurrency ' ���o������ȥN���C
'                Debug.Print "���o���ȫ��� GetExponent : " & .GetExponent ' ���o���ȫ��ơC
            Else
                strErrMsg = "Failed to AUTH." & lngErrStatus & "(" & strErrCode & ")"
            End If
        Else
            strErrMsg = "API Error = " & strRet
        End If
    
    End With
  
  Exit Function
ChkErr:
    ErrHandle FrmName, "AUTH"
End Function

'�������v
'Public Function ApproveReversal(strPostString As Variant, strRequestHeader As Variant, _
    strUserId As Variant, strPasswd As Variant, strMerchantNumber As Variant, strOrderNumber As Variant) As Long

' ���� �H�Υd���v�@�~(�Y�����q��)
Public Function AUTH_REV(strHyPOS As String, _
                                            strCurrencyType As String, _
                                            intExponentType As Integer, _
                                            intECItype As Integer, _
                                            strMerID As String, _
                                            strXID As String, _
                                            strAuthrrpid As String, _
                                            strAuthcode As String, _
                                            strTermseq As String, _
                                            lngAmt As Long) As Long
  On Error GoTo ChkErr
  
    Set objAPI = CreateObject("posapi_c_npg.AUTH_REV") ' �إ� API ����
    ' ret = pos_authrev.AuthRev(server_url, merid, xid, authrrpid, authcode, termseq, _
                                                    amt_currency, 0, ret_amount, exponent)
    ' orgamt ����v�����B
    With objAPI
        
        strRet = .AuthRev(strHyPOS, Val(strMerID), strXID, strAuthrrpid, strAuthcode, _
                                    CLng(strTermseq), strCurrencyType, 0, lngAmt, intExponentType)
        ' GetErrStatus : ���o��������檬�A�C
        ' �Ǧ^�� : ���檬�A�A0�����榨�\�A�D0���䥦���~���A�аѾ\�����G�C
        
        lngErrStatus = .GetErrStatus
        ' GetErrCode : ���o��������~�N�X�C
        ' �����檬�A����0�ɡA�i�Q��GetErrCode�禡���o������������~�N�X�C
        strErrCode = .GetErrCode
        
        If strRet = 0 Then
            If lngErrStatus = 0 And strErrCode = "00" Then
                AUTH_REV = 0
'                Debug.Print "���o������վ\�s�� GetRetrref : " & .GetRetrref ' ���o������վ\�s��
'                Debug.Print "�妸�������� GetBatchclose : " & .GetBatchclose ' �妸��������
'                Debug.Print "���o������վ\�Ǹ� GetTermseq : " & .GetTermseq ' ���o������վ\�Ǹ�
            Else
                strErrMsg = "Failed to AUTH_REV." & lngErrStatus & "(" & strErrCode & ")"
                AUTH_REV = objAPI.GetErrStatus
            End If
        Else
            strErrMsg = "API Error = " & strRet
        End If
    
    End With
  
  Exit Function
ChkErr:
    ErrHandle FrmName, "AUTH_REV"
End Function

' �H�Υd�дڧ@�~
Public Function CAP(strHyPOS As String, _
                                strCurrencyType As String, _
                                intExponentType As Integer, _
                                intECItype As Integer, _
                                strMerchantID As String, _
                                strXID As String, _
                                strAuthrrpid As String, _
                                strAuthcode As String, _
                                strTermseq As String, _
                                lngAmt As Long, _
                                Optional lngOrgAmt As Long = -1) As Long
  On Error GoTo ChkErr
  
    Set objAPI = CreateObject("posapi_c_npg.CAP") ' �إ� API ����
    ' ret = pos_cap.Cap(server_url, merid, xid, authrrpid, authcode, termseq, _
                                    amt_currency, purchAmt, purchAmt, exponent)
    With objAPI
        
        If lngOrgAmt = -1 Then lngOrgAmt = lngAmt
        
        strRet = .CAP(strHyPOS, Val(strMerchantID), strXID, strAuthrrpid, strAuthcode, _
                                CLng(strTermseq), strCurrencyType, lngAmt, lngOrgAmt, intExponentType)
        ' capamt : ��J�дڪ��B orgamt : ����v�����B
        
        ' GetErrStatus : ���o��������檬�A�C
        ' �Ǧ^�� : ���檬�A�A0�����榨�\�A�D0���䥦���~���A�аѾ\�����G�C
        lngErrStatus = .GetErrStatus
        
        ' GetErrCode : ���o��������~�N�X�C
        ' �����檬�A����0�ɡA�i�Q��GetErrCode�禡���o������������~�N�X�C
        strErrCode = .GetErrCode
        
        If strRet = 0 Then
            If lngErrStatus = 0 And strErrCode = "00" Then
                CAP = 0
'                Debug.Print "���o������վ\�s�� GetRetrref : " & .GetRetrref ' ���o������վ\�s��
'                Debug.Print "�妸�������� GetBatchclose : " & .GetBatchclose ' �妸��������
'                Debug.Print "����妸�Ǹ� GetBatchId : " & .GetBatchId ' ����妸�s��
'                Debug.Print "����妸�Ǹ� GetBatchSeq : " & GetBatchSeq ' ����妸�Ǹ�
            Else
                strErrMsg = "Failed to CAP." & lngErrStatus & "(" & strErrCode & ")"
                CAP = objAPI.GetErrStatus
            End If
        Else
            strErrMsg = "API Error = " & strRet
        End If
        
    End With
  
  Exit Function
ChkErr:
    ErrHandle FrmName, "CAP"
End Function

' ���� �H�Υd�дڧ@�~
Public Function CAP_REV(strHyPOS As String, _
                                        strCurrencyType As String, _
                                        intExponentType As Integer, _
                                        intECItype As Integer, _
                                        strMerchantID As String, _
                                        strXID As String, _
                                        strAuthrrpid As String, _
                                        strAuthcode As String, _
                                        strTermseq As String, _
                                        lngAmt As Long, _
                                        strBatchID As String, _
                                        strBatchSeq As String) As Long
  On Error GoTo ChkErr
  
    Set objAPI = CreateObject("posapi_c_npg.CAP_REV") ' �إ� API ����
    ' ret = pos_caprev.CapRev(server_url, merid, xid, authrrpid, authcode, termseq, _
                                    amt_currency, ret_amount, exponent, cap_batchid, cap_batchseq)
    
    With objAPI
        
         strRet = .CapRev(strHyPOS, Val(strMerchantID), strXID, strAuthrrpid, strAuthcode, _
                                CLng(strTermseq), strCurrencyType, lngAmt, intExponentType, _
                                CLng(strBatchID), CLng(strBatchSeq))
        
        ' GetErrStatus : ���o��������檬�A�C
        ' �Ǧ^�� : ���檬�A�A0�����榨�\�A�D0���䥦���~���A�аѾ\�����G�C
        lngErrStatus = .GetErrStatus
        
        ' GetErrCode : ���o��������~�N�X�C
        ' �����檬�A����0�ɡA�i�Q��GetErrCode�禡���o������������~�N�X�C
        strErrCode = .GetErrCode
        
        If strRet = 0 Then
            If lngErrStatus = 0 And strErrCode = "00" Then
                CAP_REV = 0
'                Debug.Print "���o������վ\�s�� GetRetrref : " & .GetRetrref ' ���o������վ\�s��
'                Debug.Print "�妸�������� GetBatchclose : " & .GetBatchclose ' �妸��������
            Else
                strErrMsg = "Failed to CAP_REV." & lngErrStatus & "(" & strErrCode & ")"
                CAP_REV = objAPI.GetErrStatus
            End If
        Else
            strErrMsg = "API Error = " & strRet
        End If
        
    End With
    
  Exit Function
ChkErr:
    ErrHandle FrmName, "CAP_REV"
End Function

' �i��h�ڰʧ@
Public Function CRED(strHyPOS As String, _
                                        strCurrencyType As String, _
                                        intExponentType As Integer, _
                                        intECItype As Integer, _
                                        strMerchantID As String, _
                                        strXID As String, _
                                        strAuthrrpid As String, _
                                        strAuthcode As String, _
                                        lngAmt As Long, _
                                        strBatchID As String, _
                                        strBatchSeq As String, _
                                        Optional lngOrgAmt As Long = -1) As Long
  On Error GoTo ChkErr
  
    Set objAPI = CreateObject("posapi_c_npg.CRED") ' �إ� API ����
    ' ret = pos_cred.CRED(server_url, merid, xid, authrrpid, authcode, amt_currency, _
                                    ret_amount, ret_amount, exponent, cap_batchid, cap_batchseq)
                                    
    ' �Y�h�ڿ��~ Status 3 , ErrCode 40 , �h�I�s�����д�
    
    With objAPI
        
        If lngOrgAmt = -1 Then lngOrgAmt = lngAmt
        ' credamt :  ���h�ڪ��B  orgamt :  ����v�����B
        
        strRet = .CRED(strHyPOS, Val(strMerchantID), strXID, strAuthrrpid, strAuthcode, _
                                strCurrencyType, lngAmt, lngOrgAmt, intExponentType, _
                                CLng(strBatchID), CLng(strBatchSeq))
                                
        ' GetErrStatus : ���o��������檬�A�C
        ' �Ǧ^�� : ���檬�A�A0�����榨�\�A�D0���䥦���~���A�аѾ\�����G�C
        lngErrStatus = .GetErrStatus
        
        ' GetErrCode : ���o��������~�N�X�C
        ' �����檬�A����0�ɡA�i�Q��GetErrCode�禡���o������������~�N�X�C
        strErrCode = .GetErrCode
        
        If strRet = 0 Then
            If lngErrStatus = 0 And strErrCode = "00" Then
                CRED = 0
'                Debug.Print "���o������վ\�s�� GetRetrref : " & .GetRetrref ' ���o������վ\�s��
'                Debug.Print "�妸�������� GetBatchclose : " & .GetBatchclose ' �妸��������
'                Debug.Print "����妸�Ǹ� GetBatchId : " & .GetBatchId ' ����妸�s��
'                Debug.Print "����妸�Ǹ� GetBatchSeq : " & GetBatchSeq ' ����妸�Ǹ�
            Else
                strErrMsg = "Failed to CRED." & lngErrStatus & "(" & strErrCode & ")"
                CRED = objAPI.GetErrStatus
            End If
        Else
            strErrMsg = "API Error = " & strRet
        End If
    
    End With
  
  Exit Function
ChkErr:
    ErrHandle FrmName, "CRED"
End Function

' �i��妸���b�ʧ@
Public Function BATCHCLOSE(strHyPOS As String, strMerchantID As String) As Long
  On Error GoTo ChkErr
    
    Set objAPI = CreateObject("posapi_c_npg.BATCHCLOSE") ' �إ� API ����
    ' ret = pos_bc.BATCHCLOSE(server_url, merid)
    
    With objAPI
        
        strRet = .BATCHCLOSE(strHyPOS, Val(strMerchantID))
        
        ' GetErrStatus : ���o��������檬�A�C
        ' �Ǧ^�� : ���檬�A�A0�����榨�\�A�D0���䥦���~���A�аѾ\�����G�C
        lngErrStatus = .GetErrStatus
        
        ' GetErrCode : ���o��������~�N�X�C
        ' �����檬�A����0�ɡA�i�Q��GetErrCode�禡���o������������~�N�X�C
        strErrCode = .GetErrCode
        
        If strRet = 0 Then
            If lngErrStatus = 0 And strErrCode = "00" Then
                BATCHCLOSE = 0
'                Debug.Print "����妸�s�� GetBatchId : " & .GetBatchId
'                Debug.Print "�妸�����ɶ� GetCloseDateTime : " & .GetCloseDateTime
            Else
                strErrMsg = "Failed to BATCHCLOSE." & lngErrStatus & "(" & strErrCode & ")"
                BATCHCLOSE = objAPI.GetErrStatus
            End If
        Else
            strErrMsg = "API Error = " & strRet
        End If
        
    End With
    
  Exit Function
ChkErr:
    ErrHandle FrmName, "BATCHCLOSE"
End Function

Public Property Get GetErrMsg(Optional GetErrDesc As Boolean = False) As String
    If GetErrDesc Then
        GetErrMsg = CvtAPImsg(strRet)
    Else
        GetErrMsg = strErrMsg
    End If
End Property

Public Property Get GetAPIreturn(Optional GetErrDesc As Boolean = False) As String
    If GetErrDesc Then
        GetAPIreturn = CvtAPImsg(strRet)
    Else
        GetAPIreturn = strRet
    End If
End Property

Public Sub InitializePara(strHyPOS As String, _
                                        strCurrencyType As String, _
                                        intExponentType As Integer, _
                                        intECItype As Integer, _
                                        strMerchantID As String)
  On Error GoTo ChkErr
  
    strSvrUrl = strHyPOS ' HyPOS���A����URL
    strCurrency = strCurrencyType ' �x��
    intExponent = intExponentType ' �s�x��
    intECI = intECItype ' �@��SSL����]�w
    strMerID = strMerchantID ' �ө��s��
    
  Exit Sub
ChkErr:
    ErrHandle FrmName, "InitializePara"
End Sub

Private Sub Class_Initialize()
  On Error GoTo ChkErr
  
    'strSvrUrl = "https://demopg.hyweb.com.tw:3841" ' HyPOS���A����URL
    ' TransAcceptPos ���Accept��m

'    If strSvrUrl = Empty Then strSvrUrl = "http://localhost:2011" ' HyPOS���A����URL
'    If strCurrency = Empty Then strCurrency = "901" ' �x��
'    If intExponent = 0 Then intExponent = 0 ' �s�x��
'    If intECI = 0 Then intECI = 7 ' �@��SSL����]�w
'    If strMerID = Empty Then strMerID = 1 ' �ө��s��

  Exit Sub
ChkErr:
    ErrHandle FrmName, "Class_Initialize"
End Sub

Public Function CvtAPImsg(Optional strAPIerr As String = "") As String
  On Error GoTo ChkErr
    If strAPIerr = Empty Then strAPIerr = strRet
    Select Case strAPIerr
            Case "12"
                CvtAPImsg = "�����s�u����"
                
            Case "14"
                CvtAPImsg = "��ƶǰe���~"
                
            Case "15"
                CvtAPImsg = "��Ʊ����O��"
                
            Case "16"
                CvtAPImsg = "�ʤ�keystore ��"
                
            Case "268435457"
                CvtAPImsg = "�q��s���榡���~"
                
            Case "268435458"
                CvtAPImsg = "�H�Υd�d���榡���~"
                
            Case "268435459"
                CvtAPImsg = "�H�Υd���Ĵ����榡���~"
                
            Case "268435460"
                CvtAPImsg = "���B�榡���~"
                
            Case "268435461"
                CvtAPImsg = "���O�榡���~"
                
            Case "268435462"
                CvtAPImsg = "CAVV�榡���~"
                
            Case "268435463"
                CvtAPImsg = "POS���A���^�Ъ���Ʀ��~"
                
            Case "268435464"
                CvtAPImsg = "����ѧO�X�榡���~"
                
            Case "268435465"
                CvtAPImsg = "���v����N�X�榡���~"
                
            Case "268435466"
                CvtAPImsg = "�q��y�z�����榡���~"
                
            Case "268435467"
                CvtAPImsg = "recur_num �榡���~"
                
            Case "268435468"
                CvtAPImsg = "birthday ���榡���~"
                
            Case "268435469"
                CvtAPImsg = "pid���榡���~"
                
            Case "268435471"
                CvtAPImsg = "������B�W�L�W��"
                
            Case "268435472"
                CvtAPImsg = "���v�X���榡���~"
                
            Case "268435473"
                CvtAPImsg = "�������B���榡���~"
                
            Case "268435474"
                CvtAPImsg = "termseq ���榡���~"
                
            Case "268435475"
                CvtAPImsg = "���O���榡���~"
                
            Case "268435476"
                CvtAPImsg = "�妸�s�����榡���~"
                
            Case "268435477"
                CvtAPImsg = "�妸�Ǹ����榡���~"
                
            Case "268435478"
                CvtAPImsg = "merid ���榡���~"
                
            Case "268435479"
                CvtAPImsg = "eci ���榡���~"
                
            Case "268435475"
                CvtAPImsg = "���O���榡���~"
                
            Case "268435480"
                CvtAPImsg = "recur_freq ��recur_end �榡���~"
            
            Case Else
                CvtAPImsg = strAPIerr
    End Select
  Exit Function
ChkErr:
    ErrHandle FrmName, "CvtAPImsg"
End Function

' CD073 [ Payment Gateway���O�N�X�� ]
' UserID �ө��b��
' UserPswd �ө��K�X
' MerchantNumber �ө��N�X
' PrgName Gateway�{���W��
' NPCPos NPC Accept�����m
' TransAcceptPos ���Accept��m
' TransAdminPos ���Admin��m
' OrderURL �ө����}
' Currency �f���O�N��
' Country ��O�N��
' PaymentType ������A

' ===����ۭq�ݩ�========================================================================

Public Property Get ServerURL() As String
    ServerURL = strSvrUrl
End Property

Public Property Let ServerURL(ByVal strServerURL As String)
    strSvrUrl = strServerURL
End Property

Public Property Get CurrencyID() As String
    CurrencyID = strCurrency
End Property

Public Property Let CurrencyID(ByVal strCurrencyID As String)
    strCurrency = strCurrencyID
End Property

Public Property Get ECI() As Integer
    ECI = intECI
End Property

Public Property Let ECI(ByVal intECIvalue As Integer)
    intECI = intECIvalue
End Property

Public Property Get MerchantID() As String
    MerchantID = strMerID
End Property

Public Property Let MerchantID(ByVal strMerchantID As String)
    strMerID = strMerchantID
End Property

Public Property Get TimeOut() As Integer
    TimeOut = intTimeOut
End Property

Public Property Let TimeOut(ByVal intTime As Integer)
    intTimeOut = intTime
End Property

' =====================================================================================

Private Sub Class_Terminate()
  On Error Resume Next
    Set objAPI = Nothing ' ���� API ����
End Sub

' �P�¶Q���q�ĥ� TWPay �t�ΡA�����Q���q���t�Τu�{�v�ɳt�F�Ѧ����ըt�ξާ@�����A
' �ФU�� TWPay API �w�ˤΨ禡��X��U�A�Ш̷Ӧw�ˤ�U�w�� POS for Windows Platform
' �ñN��X�U�禡�ܶQ���q�������q�l�ӳ�
'
' �ϥΤ�U�U���ЦܡGhttp://service.twpay.com.tw/dlapii.htm
'
' �аȥ���X�U�C�U�禡 (�U�C�U�禡���T�W�٤ΩI�s�覡�A�аѦҤ�U����)
'
' �h Auth( ) �����v
' �h Auth_Recur( ) ���������v [�ȨѤ����I�ڻݭn�ϥ�]
' �h Auth_Rev( ) �������v
' �h Cap( ) �д�
' �h Batch_close () �妸���b
'
' �Ш̾ڥH�U������T�A�i����աC
'
' �@����
' �h MerchantID�G 0012370888051 �h TerminalID�G 91010001
'
' �h    PGURL�Ghttps://demopg.hyweb.com.tw:3841
'          �� �Q�S���T�w POS for windows �Ӧ��A��
'          �i�H�z�L TCP / 3841 Port �s�ܥ����q���� PGURL
'
'
' �Y�b���դW��������D�A�ܸ��w��z�^��email�� service_twpay@twca.com.tw
' ���p���t�d���~�ȤH���A�ڭ̱N�ɳt�^���z�����D�A���¡C
'
' �����Ʒ~���n���� , ���y�A�Ȫ��̨ι٦�
' If you have any question, please do not hesitate to contact us !
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' TWPay Service Team E-mail�Gservice_twpay@twca.com.tw
'
' TEL: 02-2370-8886 �����C���p�U�G
' �ȪA�� - �i�Q�s 139
' �ȪA�� - �����p 135
' �~�ȳ� - ù�h�� 655
' FAX: 02-2370-0728
' Address: 100 �x�_�������n�� 85 �� 10 ��

' Mobile : 0968468380 Milk ��
' Tina  (03) 574-4390 Ext:619

' HyPos Server Cmd
' manage -add -m 0012370888051 -t 91010001 -p https://demopg.hyweb.com.tw:3841

' INSERT INTO CD073
' (CODENO,DESCRIPTION,MERCHANTNUMBER,PRGNAME,TRANSACCEPTPOS,CURRENCY,
' PAYMENTTYPE,EXPONENT,UPDTIME,UPDEN)
' VALUES (2,'�x�W�����{��','755842','TWPAY','HTTP://127.0.0.1:2011','901','7','0','96/01/16 13:51:21','ADMIN');

'���F�����ɮ�, �p��������D���֪����p�� �H�U���f : Kevin Liu, (02) 23956966   #576   kevin@hyweb.com.tw

'�s�� Tina 619
'03-5744390 ext 603

' 02-2395-6966 ext 579

' 5. �ˬd�ϥ�POS API ���{���A��u�H�Υd�d���vpan �ѼƬO�_����X16 �X�d����3 �XCVC2/CVV2
' 16. �O�_�w���� �Q�Τ��X���yAPI ����B�D���s�q��������t�Ϊ��ӽ�
'       �ж���uIP �}�q��v�i�������q �Q�S����X���y����~IP�A�����q�N�w�惡IP �i�樾����}�q

' 02 - 23956966 ext 576 �Ƹּ� ,  IP ����