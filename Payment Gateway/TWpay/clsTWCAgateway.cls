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

' COM元件所需檔案
' posapi_c.dll : COM元件的主要動態連結檔
' posapi_d.dll : pos api的主要動態連結檔
' posapi_d.lib : posapi_d.dll的lib檔
' posapi_d.h : posapi_d.dll的函式定義檔
' hy_err.h : posapi_d.dll的錯誤碼
' libeay32.dll : OpenSSL函式庫
' ssleay32.dll : OpenSSL函式庫

' CLSID : C6A2E34F-B5D6-43B4-AF03-C068B5D2EE80
' ProgId:  posapi_c.Auth

Private strSvrUrl As String

Private strLIDM As String ' 訂單編號
Private strMerID As String ' 商店編號
Private strPan As String ' 信用卡號末三碼值
Private strExpDate As String ' 信用卡有效期限
Private strCurrency As String ' 交易幣值代號
Private lngPurchAmt As Long ' 交易授權總金額
Private intExponent As Integer ' 幣值指數
Private intEci As Integer ' 電子商務交易安全等級

Private Const FrmName = "clsTWCAgateway"

Private objAPI As Object ' API 物件
Private lngRet As Long ' 回傳值
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

' 授 權
Public Function Approve(strOrderNo As String, strCardNo As String, strDate As String, lngAmt As Long)
  On Error GoTo ChkErr
    
    strLIDM = "12345678"
    strPan = "4444444444444444"
    strExpDate = "200808"
    lngPurchAmt = 1000
    
    With objAPI

'        .SetBirthday "20040212" ' Optional Setting
'        .SetPID "P123456789" ' Optional Setting

        .SetOrderDesc "Authorize Testing" ' Optional Setting 訂單說明
        
        .SetTimeout 10 ' Optional Setting 交易逾時
        
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
  
    Set objAPI = CreateObject("PosApi_C.AUTH") ' 建立 API 物件
    If Err Then strErrMsg = " API ( PosApi_C.AUTH ) 物件建立失敗 !"
    
    'strSvrUrl = "https://demopg.hyweb.com.tw:3841" ' HyPOS伺服器的URL
    strSvrUrl = "http://localhost:2011" ' HyPOS伺服器的URL
    strCurrency = "901" ' 台幣
    intExponent = 0 ' 新台幣
    intEci = 7 ' 一般SSL交易設定
    strMerID = 1 ' 商店編號
    
  Exit Sub
ChkErr:
    ErrHandle FrmName, "Class_Initialize"
End Sub

Private Sub Class_Terminate()
  On Error Resume Next
    Set objAPI = Nothing ' 釋放 API 物件
End Sub

'感謝貴公司採用 TWPay 系統，為讓貴公司之系統工程師盡速了解此測試系統操作介面，
'請下載 TWPay API 安裝及函式整合手冊，請依照安裝手冊安裝 POS for Windows Platform
'並將整合各函式至貴公司的網路電子商場
'
'使用手冊下載請至：http://service.twpay.com.tw/dlapii.htm
'
'請務必整合下列各函式 (下列各函式正確名稱及呼叫方式，請參考手冊說明)
'
'▇ Auth( ) 取授權
'▇ Auth_Recur( ) 取分期授權 [僅供分期付款需要使用]
'▇ Auth_Rev( ) 取消授權
'▇ Cap( ) 請款
'▇ Batch_close () 批次結帳
'
'請依據以下相關資訊，進行測試。
'
'
'一般交易
'▇ MerchantID： 0012370888051 ▇ TerminalID： 91010001
'
'
'
'▇    PGURL：https://demopg.hyweb.com.tw:3841
'          請 貴特店確定 POS for windows 該伺服器
'          可以透過 TCP / 3841 Port 連至本公司測試 PGURL
'
'
'若在測試上有任何問題，竭誠歡迎您回覆email至 service_twpay@twca.com.tw
'或聯絡負責此業務人員，我們將盡速回應您的問題，謝謝。
'
' 網路事業的好幫手 , 金流服務的最佳夥伴
'If you have any question, please do not hesitate to contact us !
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'TWPay Service Team E-mail：service_twpay@twca.com.tw
'
'TEL: 02-2370-8886 分機列表如下：
'客服部 - 張淑貞 139
'客服部 - 高健峰 135
'業務部 - 羅懷玉 655
'FAX: 02-2370-0728
'Address: 100 台北市延平南路 85 號 10 樓
'
'
