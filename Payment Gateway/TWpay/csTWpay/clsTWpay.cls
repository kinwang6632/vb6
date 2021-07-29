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

' Doc : CITI-NPG-POS_API_手冊_V1_06_一般與分期__960115.pdf

' 變數名稱 : 意義
' cavv : (optional)3D Secure的CAVV欄位。長度為28個字元的字串。
' currency : 交易幣值代號。長度為3個字元的字串。(台幣”901”)
' eci : 為電子商務交易安全等級。一般SSL交易eci值須設定為7。
' expDate : 為信用卡之有效期限，格式固定為西元六碼 YYYYMM。
' exponent : 為幣值指數，新台幣為0。
' lidm : 為電子商場的應用程式所給予此筆交易的訂單編號，資料型態為最長19個字元的文字串。
'          注意﹕訂單編號字串的字元為一般文數字組合，勿使用空白及特殊字元如\/:*?"<>|%$&’等。
' merid : 為此交易商店於SSL HyPOS Server中的商店編號設定值，由電子商場自行設定。
'            注意：此代碼與銀行所授與的MerchantID不同。
' order_desc : (optional)訂單說明，最長為125個字元。
'                     注意﹕訂單編號字串的字元為一般文數字組合，勿使用空白及特殊字元如\/:*?"<>|%$&’等。
' pan : 為信用卡號，卡片後末三碼值附加在16碼卡號後。
' purchAmt : 為消費者此筆交易所購買商品欲授權總金額，正整數型態。
' server : HyPOS伺服器的URL，格式為http://ip:port。(如http://127.0.0.1:2011)
' timeout : 設定接收HyPOS伺服器回傳的逾時時間，單位為秒，若timeout小於或等於0時，則timeout則設定為預設值30秒。

' [OUT]
' AUTH_OUT結構﹕
' 變數名稱 : 意義
' amount : 授權金額。
' authcode : 交易授權碼，最大長度為6的字串。
' authrrpid : SSL 授權交易之代碼，最長為 40 個位元文字串。
' BATCHCLOSE : 如果batchclose為1，表示批次己滿，須關閉此批次方可進行交易。如果批次未滿，則batchclose為0。
' currency : 交易幣值代號。長度為3個字元的字串。(台幣”901”)
' err_code : 錯誤代碼，長度為2的字串，請參閱附錄二。
' err_status : 執行狀態，為數字形態。(其中，0 表執行成功，其它各類狀態請參閱附錄二)
' exponent : 為幣值指數，新台幣為0。
' retrref : 調閱編號，最長為12個位元文字串。
' termseq : 調閱序號 , 正整數型態
' xid : 交易識別碼，最長為40個位元文字串。

' 函式名稱 : 說明
' Auth : 進行信用卡授權動作。
' GetAuthcode : 取得授權交易授權碼。
' GetAuthrrpid : 取得授權交易之代碼。
' GetAmount : 取得授權金額。
' GetBatchclose : 批次關閉狀態提示。
' GetCurrency : 取得交易幣值代號。
' GetErrStatus : 取得交易的執行狀態。
' GetErrCode : 取得交易的錯誤代碼。
' GetExponent : 取得幣值指數。
' GetRettref : 取得交易之調閱編號。
' GetTermseq : 取得交易之調閱序號。
' GetXID : 取得交易識別碼。
' SetCAVV : 設定3D Secure的CAVV值。
' SetOrderDesc : 設定訂單說明。
' SetTimeout : 設定交易逾時時間。

'COM API 元件所包含之檔案
' posapi_c_npg.dll 交易COM 元件的主要動態連結檔
' posapi_d.dll pos api 的主要動態連結檔
' libeay32.dll OpenSSL 函式庫
' ssleay32.dll OpenSSL 函式庫

' CLSID : C6A2E34F-B5D6-43B4-AF03-C068B5D2EE80
' ProgId:  posapi_c.Auth

Private strSvrUrl As String

Private strLIDM As String ' 訂單編號
Private strMerID As String ' 商店編號
Private strPan As String ' 為信用卡號，卡片後末三碼值附加在16碼卡號後。
Private strExpDate As String ' 信用卡有效期限
Private strCurrency As String ' 交易幣值代號
Private lngPurchAmt As Long ' 交易授權總金額
Private intExponent As Integer ' 幣值指數
Private intECI As Integer ' 電子商務交易安全等級
Private intTimeOut As Integer ' Time Out

Private Const FrmName = "clsTWCAgateway"

Private objAPI As Object ' API 物件
Private strRet As String ' 回傳值
'Private api As New POSAPI_CLib.AUTH

Private lngErrStatus As Long
Private strErrCode As String

' ===元件對應屬性========================================================================

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

'AUTH:  信用卡授權作業
'AUTH_RECUR:  信用卡分期付款授權作業 (本服務視收單銀行是否提供)
'AUTH_REV:  取消 信用卡授權作業(即取消訂單)
'CAP:  信用卡請款作業
'CAP_REV:   取消 信用卡請款作業
'Cred:   進行退款動作
'CredRev:   進行取消退貨動作
'BATCHCLOSE:   商店與收單銀行批次結帳 (完成此作業銀行才會撥帳款給貴特店)

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

' 底下對應 Jacky 寫的 藍新科技 Payment  Gateway 的 Interface

' 授權
'Public Function Approve(strPostString As Variant, strRequestHeader As Variant, _
    strUserId As Variant, strPasswd As Variant, strMerchantNumber As Variant, _
    strOrderNumber As Variant, lngAmount As Variant, lngDepositFlag As Variant) As Long

' 授權
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

' 信用卡授權作業
Public Function AUTH(strOrderNo As String, strCardNo As String, strDate As String, lngAmt As Long, _
                                Optional ByRef lngAPIreturn As Long = -1) As Boolean
  On Error GoTo ChkErr
  
'    strLIDM = "12345678" ' 訂單編號
'    strPan = "4444444444444444" ' 為信用卡號，卡片後末三碼值附加在16碼卡號後。
'    strExpDate = "200808" ' 信用卡有效期限
'    lngPurchAmt = 1000 ' 交易授權總金額
    
    strLIDM = strOrderNo ' 訂單編號
    strPan = strCardNo ' 為信用卡號，卡片後末三碼值附加在16碼卡號後。
    strExpDate = strDate ' 信用卡有效期限
    lngPurchAmt = lngAmt ' 交易授權總金額

'    Set objAPI = CreateObject("PosApi_C.AUTH") ' 建立 API 物件
    Set objAPI = CreateObject("posapi_c_npg.AUTH") ' 建立 API 物件
'    If Err Then strErrMsg = " API ( PosApi_C.AUTH ) 物件建立失敗 !"
    
    With objAPI

'        .SetBirthday "20040212" ' Optional Setting
'        .SetPID "P123456789" ' Optional Setting
'        .SetOrderDesc "Authorize Testing" ' Optional Setting 訂單說明
        '.SetTimeout 10 ' Optional Setting 交易逾時
        ' Function AUTH(server As String, merid As Long, lidm As String, pan As String, expDate As String,
        '                        currency As String, purchAmt As Long, exponent As Long, eci As Long) As Long
        
        intECI = GetECI(strCardNo)
        
        strRet = .AUTH(strSvrUrl, strMerID, strLIDM, strPan, strExpDate, _
                                strCurrency, lngPurchAmt, intExponent, intECI)
        
        lngAPIreturn = strRet
        ' GetErrStatus : 取得交易的執行狀態。
        ' 傳回值 : 執行狀態，0表執行成功，非0為其它錯誤狀態請參閱附錄二。
        lngErrStatus = .GetErrStatus
        
        ' GetErrCode : 取得交易的錯誤代碼。
        ' 當執行狀態不為0時，可利用GetErrCode函式取得此次的交易錯誤代碼。
        strErrCode = .GetErrCode
        
        If strRet = 0 Then
            'If strErrCode <> 0 Then strErrMsg = "Failed to Auth." & lngErrStatus & "(" & strErrCode & ")"
            If lngErrStatus = 0 And strErrCode = "00" Then
                AUTH = True
'                Debug.Print "取得交易識別碼 GetXID : " & .GetXID ' 取得交易識別碼。
'                Debug.Print "取得授權交易之代碼 GetAuthrrpid : " & .GetAuthrrpid ' 取得授權交易之代碼。
'                Debug.Print "取得交易之調閱序號 GetTermseq : " & .GetTermseq ' 取得交易之調閱序號。
'                Debug.Print "取得交易之調閱編號 GetRetrref : " & .GetRetrref ' 取得交易之調閱編號。
'                Debug.Print "批次關閉狀態提示 GetBatchclose : " & .GetBatchclose ' 批次關閉狀態提示。
'                Debug.Print "取得授權交易授權碼 GetAuthcode : " & .GetAuthcode ' 取得授權交易授權碼。
'                Debug.Print "取得授權金額 GetAmount : " & .GetAmount ' 取得授權金額。
'                Debug.Print "取得交易幣值代號 GetCurrency : " & .GetCurrency ' 取得交易幣值代號。
'                Debug.Print "取得幣值指數 GetExponent : " & .GetExponent ' 取得幣值指數。
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

'取消授權
'Public Function ApproveReversal(strPostString As Variant, strRequestHeader As Variant, _
    strUserId As Variant, strPasswd As Variant, strMerchantNumber As Variant, strOrderNumber As Variant) As Long

' 取消 信用卡授權作業(即取消訂單)
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
  
    Set objAPI = CreateObject("posapi_c_npg.AUTH_REV") ' 建立 API 物件
    ' ret = pos_authrev.AuthRev(server_url, merid, xid, authrrpid, authcode, termseq, _
                                                    amt_currency, 0, ret_amount, exponent)
    ' orgamt 原授權之金額
    With objAPI
        
        strRet = .AuthRev(strHyPOS, Val(strMerID), strXID, strAuthrrpid, strAuthcode, _
                                    CLng(strTermseq), strCurrencyType, 0, lngAmt, intExponentType)
        ' GetErrStatus : 取得交易的執行狀態。
        ' 傳回值 : 執行狀態，0表執行成功，非0為其它錯誤狀態請參閱附錄二。
        
        lngErrStatus = .GetErrStatus
        ' GetErrCode : 取得交易的錯誤代碼。
        ' 當執行狀態不為0時，可利用GetErrCode函式取得此次的交易錯誤代碼。
        strErrCode = .GetErrCode
        
        If strRet = 0 Then
            If lngErrStatus = 0 And strErrCode = "00" Then
                AUTH_REV = 0
'                Debug.Print "取得交易之調閱編號 GetRetrref : " & .GetRetrref ' 取得交易之調閱編號
'                Debug.Print "批次關閉提示 GetBatchclose : " & .GetBatchclose ' 批次關閉提示
'                Debug.Print "取得交易之調閱序號 GetTermseq : " & .GetTermseq ' 取得交易之調閱序號
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

' 信用卡請款作業
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
  
    Set objAPI = CreateObject("posapi_c_npg.CAP") ' 建立 API 物件
    ' ret = pos_cap.Cap(server_url, merid, xid, authrrpid, authcode, termseq, _
                                    amt_currency, purchAmt, purchAmt, exponent)
    With objAPI
        
        If lngOrgAmt = -1 Then lngOrgAmt = lngAmt
        
        strRet = .CAP(strHyPOS, Val(strMerchantID), strXID, strAuthrrpid, strAuthcode, _
                                CLng(strTermseq), strCurrencyType, lngAmt, lngOrgAmt, intExponentType)
        ' capamt : 轉入請款金額 orgamt : 原授權之金額
        
        ' GetErrStatus : 取得交易的執行狀態。
        ' 傳回值 : 執行狀態，0表執行成功，非0為其它錯誤狀態請參閱附錄二。
        lngErrStatus = .GetErrStatus
        
        ' GetErrCode : 取得交易的錯誤代碼。
        ' 當執行狀態不為0時，可利用GetErrCode函式取得此次的交易錯誤代碼。
        strErrCode = .GetErrCode
        
        If strRet = 0 Then
            If lngErrStatus = 0 And strErrCode = "00" Then
                CAP = 0
'                Debug.Print "取得交易之調閱編號 GetRetrref : " & .GetRetrref ' 取得交易之調閱編號
'                Debug.Print "批次關閉提示 GetBatchclose : " & .GetBatchclose ' 批次關閉提示
'                Debug.Print "交易批次序號 GetBatchId : " & .GetBatchId ' 交易批次編號
'                Debug.Print "交易批次序號 GetBatchSeq : " & GetBatchSeq ' 交易批次序號
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

' 取消 信用卡請款作業
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
  
    Set objAPI = CreateObject("posapi_c_npg.CAP_REV") ' 建立 API 物件
    ' ret = pos_caprev.CapRev(server_url, merid, xid, authrrpid, authcode, termseq, _
                                    amt_currency, ret_amount, exponent, cap_batchid, cap_batchseq)
    
    With objAPI
        
         strRet = .CapRev(strHyPOS, Val(strMerchantID), strXID, strAuthrrpid, strAuthcode, _
                                CLng(strTermseq), strCurrencyType, lngAmt, intExponentType, _
                                CLng(strBatchID), CLng(strBatchSeq))
        
        ' GetErrStatus : 取得交易的執行狀態。
        ' 傳回值 : 執行狀態，0表執行成功，非0為其它錯誤狀態請參閱附錄二。
        lngErrStatus = .GetErrStatus
        
        ' GetErrCode : 取得交易的錯誤代碼。
        ' 當執行狀態不為0時，可利用GetErrCode函式取得此次的交易錯誤代碼。
        strErrCode = .GetErrCode
        
        If strRet = 0 Then
            If lngErrStatus = 0 And strErrCode = "00" Then
                CAP_REV = 0
'                Debug.Print "取得交易之調閱編號 GetRetrref : " & .GetRetrref ' 取得交易之調閱編號
'                Debug.Print "批次關閉提示 GetBatchclose : " & .GetBatchclose ' 批次關閉提示
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

' 進行退款動作
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
  
    Set objAPI = CreateObject("posapi_c_npg.CRED") ' 建立 API 物件
    ' ret = pos_cred.CRED(server_url, merid, xid, authrrpid, authcode, amt_currency, _
                                    ret_amount, ret_amount, exponent, cap_batchid, cap_batchseq)
                                    
    ' 若退款錯誤 Status 3 , ErrCode 40 , 則呼叫取消請款
    
    With objAPI
        
        If lngOrgAmt = -1 Then lngOrgAmt = lngAmt
        ' credamt :  欲退款金額  orgamt :  原授權之金額
        
        strRet = .CRED(strHyPOS, Val(strMerchantID), strXID, strAuthrrpid, strAuthcode, _
                                strCurrencyType, lngAmt, lngOrgAmt, intExponentType, _
                                CLng(strBatchID), CLng(strBatchSeq))
                                
        ' GetErrStatus : 取得交易的執行狀態。
        ' 傳回值 : 執行狀態，0表執行成功，非0為其它錯誤狀態請參閱附錄二。
        lngErrStatus = .GetErrStatus
        
        ' GetErrCode : 取得交易的錯誤代碼。
        ' 當執行狀態不為0時，可利用GetErrCode函式取得此次的交易錯誤代碼。
        strErrCode = .GetErrCode
        
        If strRet = 0 Then
            If lngErrStatus = 0 And strErrCode = "00" Then
                CRED = 0
'                Debug.Print "取得交易之調閱編號 GetRetrref : " & .GetRetrref ' 取得交易之調閱編號
'                Debug.Print "批次關閉提示 GetBatchclose : " & .GetBatchclose ' 批次關閉提示
'                Debug.Print "交易批次序號 GetBatchId : " & .GetBatchId ' 交易批次編號
'                Debug.Print "交易批次序號 GetBatchSeq : " & GetBatchSeq ' 交易批次序號
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

' 進行批次結帳動作
Public Function BATCHCLOSE(strHyPOS As String, strMerchantID As String) As Long
  On Error GoTo ChkErr
    
    Set objAPI = CreateObject("posapi_c_npg.BATCHCLOSE") ' 建立 API 物件
    ' ret = pos_bc.BATCHCLOSE(server_url, merid)
    
    With objAPI
        
        strRet = .BATCHCLOSE(strHyPOS, Val(strMerchantID))
        
        ' GetErrStatus : 取得交易的執行狀態。
        ' 傳回值 : 執行狀態，0表執行成功，非0為其它錯誤狀態請參閱附錄二。
        lngErrStatus = .GetErrStatus
        
        ' GetErrCode : 取得交易的錯誤代碼。
        ' 當執行狀態不為0時，可利用GetErrCode函式取得此次的交易錯誤代碼。
        strErrCode = .GetErrCode
        
        If strRet = 0 Then
            If lngErrStatus = 0 And strErrCode = "00" Then
                BATCHCLOSE = 0
'                Debug.Print "交易批次編號 GetBatchId : " & .GetBatchId
'                Debug.Print "批次關閉時間 GetCloseDateTime : " & .GetCloseDateTime
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
  
    strSvrUrl = strHyPOS ' HyPOS伺服器的URL
    strCurrency = strCurrencyType ' 台幣
    intExponent = intExponentType ' 新台幣
    intECI = intECItype ' 一般SSL交易設定
    strMerID = strMerchantID ' 商店編號
    
  Exit Sub
ChkErr:
    ErrHandle FrmName, "InitializePara"
End Sub

Private Sub Class_Initialize()
  On Error GoTo ChkErr
  
    'strSvrUrl = "https://demopg.hyweb.com.tw:3841" ' HyPOS伺服器的URL
    ' TransAcceptPos 交易Accept位置

'    If strSvrUrl = Empty Then strSvrUrl = "http://localhost:2011" ' HyPOS伺服器的URL
'    If strCurrency = Empty Then strCurrency = "901" ' 台幣
'    If intExponent = 0 Then intExponent = 0 ' 新台幣
'    If intECI = 0 Then intECI = 7 ' 一般SSL交易設定
'    If strMerID = Empty Then strMerID = 1 ' 商店編號

  Exit Sub
ChkErr:
    ErrHandle FrmName, "Class_Initialize"
End Sub

Public Function CvtAPImsg(Optional strAPIerr As String = "") As String
  On Error GoTo ChkErr
    If strAPIerr = Empty Then strAPIerr = strRet
    Select Case strAPIerr
            Case "12"
                CvtAPImsg = "網路連線失敗"
                
            Case "14"
                CvtAPImsg = "資料傳送錯誤"
                
            Case "15"
                CvtAPImsg = "資料接收逾時"
                
            Case "16"
                CvtAPImsg = "缺少keystore 檔"
                
            Case "268435457"
                CvtAPImsg = "訂單編號格式錯誤"
                
            Case "268435458"
                CvtAPImsg = "信用卡卡號格式錯誤"
                
            Case "268435459"
                CvtAPImsg = "信用卡有效期限格式錯誤"
                
            Case "268435460"
                CvtAPImsg = "金額格式錯誤"
                
            Case "268435461"
                CvtAPImsg = "幣別格式錯誤"
                
            Case "268435462"
                CvtAPImsg = "CAVV格式錯誤"
                
            Case "268435463"
                CvtAPImsg = "POS伺服器回覆的資料有誤"
                
            Case "268435464"
                CvtAPImsg = "交易識別碼格式錯誤"
                
            Case "268435465"
                CvtAPImsg = "授權交易代碼格式錯誤"
                
            Case "268435466"
                CvtAPImsg = "訂單描述說明格式錯誤"
                
            Case "268435467"
                CvtAPImsg = "recur_num 格式錯誤"
                
            Case "268435468"
                CvtAPImsg = "birthday 欄位格式錯誤"
                
            Case "268435469"
                CvtAPImsg = "pid欄位格式錯誤"
                
            Case "268435471"
                CvtAPImsg = "交易金額超過上限"
                
            Case "268435472"
                CvtAPImsg = "授權碼欄位格式錯誤"
                
            Case "268435473"
                CvtAPImsg = "取消金額欄位格式錯誤"
                
            Case "268435474"
                CvtAPImsg = "termseq 欄位格式錯誤"
                
            Case "268435475"
                CvtAPImsg = "幣別欄位格式錯誤"
                
            Case "268435476"
                CvtAPImsg = "批次編號欄位格式錯誤"
                
            Case "268435477"
                CvtAPImsg = "批次序號欄位格式錯誤"
                
            Case "268435478"
                CvtAPImsg = "merid 欄位格式錯誤"
                
            Case "268435479"
                CvtAPImsg = "eci 欄位格式錯誤"
                
            Case "268435475"
                CvtAPImsg = "幣別欄位格式錯誤"
                
            Case "268435480"
                CvtAPImsg = "recur_freq 或recur_end 格式錯誤"
            
            Case Else
                CvtAPImsg = strAPIerr
    End Select
  Exit Function
ChkErr:
    ErrHandle FrmName, "CvtAPImsg"
End Function

' CD073 [ Payment Gateway類別代碼檔 ]
' UserID 商店帳號
' UserPswd 商店密碼
' MerchantNumber 商店代碼
' PrgName Gateway程式名稱
' NPCPos NPC Accept交易位置
' TransAcceptPos 交易Accept位置
' TransAdminPos 交易Admin位置
' OrderURL 商店網址
' Currency 貨幣別代號
' Country 國別代號
' PaymentType 交易型態

' ===元件自訂屬性========================================================================

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
    Set objAPI = Nothing ' 釋放 API 物件
End Sub

' 感謝貴公司採用 TWPay 系統，為讓貴公司之系統工程師盡速了解此測試系統操作介面，
' 請下載 TWPay API 安裝及函式整合手冊，請依照安裝手冊安裝 POS for Windows Platform
' 並將整合各函式至貴公司的網路電子商場
'
' 使用手冊下載請至：http://service.twpay.com.tw/dlapii.htm
'
' 請務必整合下列各函式 (下列各函式正確名稱及呼叫方式，請參考手冊說明)
'
' ▇ Auth( ) 取授權
' ▇ Auth_Recur( ) 取分期授權 [僅供分期付款需要使用]
' ▇ Auth_Rev( ) 取消授權
' ▇ Cap( ) 請款
' ▇ Batch_close () 批次結帳
'
' 請依據以下相關資訊，進行測試。
'
' 一般交易
' ▇ MerchantID： 0012370888051 ▇ TerminalID： 91010001
'
' ▇    PGURL：https://demopg.hyweb.com.tw:3841
'          請 貴特店確定 POS for windows 該伺服器
'          可以透過 TCP / 3841 Port 連至本公司測試 PGURL
'
'
' 若在測試上有任何問題，竭誠歡迎您回覆email至 service_twpay@twca.com.tw
' 或聯絡負責此業務人員，我們將盡速回應您的問題，謝謝。
'
' 網路事業的好幫手 , 金流服務的最佳夥伴
' If you have any question, please do not hesitate to contact us !
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' TWPay Service Team E-mail：service_twpay@twca.com.tw
'
' TEL: 02-2370-8886 分機列表如下：
' 客服部 - 張淑貞 139
' 客服部 - 高健峰 135
' 業務部 - 羅懷玉 655
' FAX: 02-2370-0728
' Address: 100 台北市延平南路 85 號 10 樓

' Mobile : 0968468380 Milk 郭
' Tina  (03) 574-4390 Ext:619

' HyPos Server Cmd
' manage -add -m 0012370888051 -t 91010001 -p https://demopg.hyweb.com.tw:3841

' INSERT INTO CD073
' (CODENO,DESCRIPTION,MERCHANTNUMBER,PRGNAME,TRANSACCEPTPOS,CURRENCY,
' PAYMENTTYPE,EXPONENT,UPDTIME,UPDEN)
' VALUES (2,'台灣網路認證','755842','TWPAY','HTTP://127.0.0.1:2011','901','7','0','96/01/16 13:51:21','ADMIN');

'為了爭取時效, 如有任何問題儘快直接聯絡 以下窗口 : Kevin Liu, (02) 23956966   #576   kevin@hyweb.com.tw

'新竹 Tina 619
'03-5744390 ext 603

' 02-2395-6966 ext 579

' 5. 檢查使用POS API 的程式，其「信用卡卡號」pan 參數是否有整合16 碼卡號及3 碼CVC2/CVV2
' 16. 是否已完成 貴用戶整合金流API 的營運主機連通網路收單系統的申請
'       請填報「IP 開通單」告知本公司 貴特店整合金流的對外IP，本公司將針對此IP 進行防火牆開通

' 02 - 23956966 ext 576 蟻詩瑩 ,  IP 表單
