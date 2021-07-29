VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  '雙線固定對話方塊
   Caption         =   " CITI Bank Payment Gateway  IN TWCA .."
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   4680
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton Command5 
      Caption         =   " 取消 信用卡請款作業 ( CAP_REV )"
      Height          =   375
      Left            =   270
      TabIndex        =   8
      Top             =   2880
      Width           =   4095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "信用卡請款作業 ( CAP )"
      Height          =   375
      Left            =   270
      TabIndex        =   7
      Top             =   2430
      Width           =   4095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "取消授權 ( AUTH_REV )"
      Height          =   375
      Left            =   270
      TabIndex        =   6
      Top             =   1980
      Width           =   4095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "授 權 ( Auth )"
      Height          =   375
      Left            =   2340
      TabIndex        =   5
      Top             =   1470
      Width           =   2025
   End
   Begin VB.CommandButton cmdShowFlow 
      Caption         =   "Function Flow Diagram"
      Height          =   375
      Left            =   1695
      TabIndex        =   4
      Top             =   840
      Width           =   2666
   End
   Begin VB.CommandButton cmdAuth 
      Caption         =   "授 權 ( Auth )"
      Height          =   375
      Left            =   270
      TabIndex        =   3
      Top             =   1470
      Width           =   2025
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Inet Obj"
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   360
      Width           =   1245
   End
   Begin VB.CommandButton cmdAPI 
      Caption         =   "API"
      Height          =   375
      Left            =   1695
      TabIndex        =   1
      Top             =   360
      Width           =   1245
   End
   Begin VB.CommandButton cmdURL 
      Caption         =   "URL"
      Height          =   375
      Left            =   270
      TabIndex        =   0
      Top             =   360
      Width           =   1245
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private strErrorMessage As String

' 授 權
'Private Function AUTH() As Boolean
'
'    Set pos_auth = server.CreateObject("posapi_c.AUTH")
'
'    server_url = "http://scorpion:2011"
'    merid = 1
'    lidm = "12345678"
'    pan = "4444444444444444"
'    expDate = "200808"
'    amt_currency = "901"
'    purchAmt = 1000
'    exponent = 0
'    eci = 7
'
'    strPara(0) = "MerchantID=0012370888064"
'    strPara(1) = "merID=61"
'    strPara(2) = "TerminalID=71010001"
'    strPara(3) = "MerchantName=PowerTV"
'    strPara(4) = "purchAmt=186"
'    strPara(5) = "AutoCap=1"
'    strPara(6) = "txType=0"
'
'    pos_auth.SetBirthday ("20040212")   ' optional setting
'    pos_auth.SetOrderDesc ("AuthTest")  ' optional setting
'    pos_auth.SetPID ("P123456789")      ' optional setting
'    pos_auth.SetTimeout (10)            ' optional setting
'
'    ret = pos_auth.AUTH(server_url, merid, _
'            lidm, pan, expDate, _
'            amt_currency, purchAmt, exponent, _
'            eci)
'
'    If ret = 0 Then
'        Status = pos_auth.GetErrStatus()
'        ErrCode = pos_auth.GetErrCode()
'        If Status <> 0 Then
'            response.write "Failed to auth." & Status & "(" + ErrCode & ")"
'            ErrExit pos_auth
'        Else
'            xid = pos_auth.GetXID()
'            authrrpid = pos_auth.GetAuthrrpid()
'            termseq = pos_auth.GetTermseq()
'            retrref = pos_auth.GetRetrref()
'            BATCHCLOSE = pos_auth.GetBatchclose()
'            authcode = pos_auth.GetAuthcode()
'            ret_amount = pos_auth.GetAmount()
'            ret_currency = pos_auth.GetCurrency()
'            ret_exponent = pos_auth.GetExponent()
'
'            response.write "xid = " + xid + "<br>"
'            response.write "authrrpid = " + authrrpid + "<br>"
'            response.write "authcoded = " + authcode + "<br>"
'            response.write "ret amount = " + ret_currency + " " + CStr(ret_amount) + " " + CStr(ret_exponent) + "<br>"
'            response.write "termseq = " + CStr(termseq) + "<br>"
'            response.write "retrref = " + retrref + "<br>"
'            response.write "batchclose = " + CStr(BATCHCLOSE) + "<br>"
'        End If
'    Else
'        response.write "API error = " + CStr(ret)
'        ErrExit pos_auth
'    End If
'
'    Set pos_auth = Nothing
'End If
'End Function

Private Sub cmdAuth_Click()
    

'    Exponent 幣值指數 Varchar2 (5)
'    Xid 交易識別碼 Varchar2(60)
'    Authrrpid SSL 授權交易之代碼 Varchar2(60)
'    AuthCode 交易授權碼 Varchar2(10)
'    TermSeq 調閱序號 Varchar2(12)
'    BatchId 交易批次編號  Varchar2(22)
'    BatchSeq 交易批次序號  Varchar2(22)
    
    
    Dim csPay As Object
    Dim lngAPIreturn As Long
    
    Set csPay = CreateObject("csPayment.clsTWpay")
    
'    ' 授 權
'    Public Function Auth(strOrderNo As String, strCardNo As String, strDate As String, lngAmt As Long) As Boolean
'       訂單編號    信用卡號    信用卡有效期限  交易授權總金額
    With csPay
        .InitializePara "http://localhost:2011", "901", 0, 7, 1
        ' HyPOS
        ' Currency
        ' Exponent
        ' ECI
        ' MerID
        If .AUTH("12345678", "4444444444444444", "200808", 1000, lngAPIreturn) Then
            Debug.Print "取得交易識別碼 GetXID : " & .GetXID ' 取得交易識別碼。
            Debug.Print "取得授權交易之代碼 GetAuthrrpid : " & .GetAuthrrpid ' 取得授權交易之代碼。
            Debug.Print "取得交易之調閱序號 GetTermseq : " & .GetTermseq ' 取得交易之調閱序號。
            Debug.Print "取得交易之調閱編號 GetRetrref : " & .GetRetrref ' 取得交易之調閱編號。
            Debug.Print "批次關閉狀態提示 GetBatchclose : " & .GetBatchclose ' 批次關閉狀態提示。
            Debug.Print "取得授權交易授權碼 GetAuthcode : " & .GetAuthcode ' 取得授權交易授權碼。
            Debug.Print "取得授權金額 GetAmount : " & .GetAmount ' 取得授權金額。
            Debug.Print "取得交易幣值代號 GetCurrency : " & .GetCurrency ' 取得交易幣值代號。
            Debug.Print "取得幣值指數 GetExponent : " & .GetExponent ' 取得幣值指數。
            MsgBox "OK !"
        Else
            ' GetErrStatus : 取得交易的執行狀態。
            ' GetErrCode : 取得交易的錯誤代碼
            Debug.Print "Err Status : " & .GetErrStatus
            Debug.Print "Err Code : " & .GetErrCode
            Debug.Print "API Return : " & .GetAPIreturn
            Debug.Print "API Return Description : " & .GetAPIreturn(True)
            Debug.Print "API Return ( Byref ) : " & lngAPIreturn
            MsgBox "No Good !"
        End If
    End With
    
    Set csPay = Nothing
    
End Sub

Private Sub cmdShowFlow_Click()
    Form2.Show 1, Me
End Sub

'測試結果
'
'1. URL 方式
'    用 XMLHTTP 物件測試 , GET 方式 OK , Post 方式 , 會參數不足
'    且網頁部份有 2 ~ 3 頁 , 且有隱藏欄位 , 所以測不成功
'    除非改網頁 , 否則目前沒辦法使用 URL 方式
'
'2. API 方式
'    應該是可以的 , 花點時間 Try , 就會有結果
'    不過有環境及安裝上的問題


Private Sub cmdURL_Click()
    
    'http://demopg.hyweb.com.tw/citipos/SSLAuthUI.cgi?MerchantID=0012370888064&merID=61&TerminalID=71010001&MerchantName=PowerTV&purchAmt=186&AutoCap=1&txType=0
    
    Dim strURL As String
    Dim strWebSite As String
    Dim strPara(6) As String
    Dim strPostData As String
    Dim strMethod As String
    Dim objXH As Object
    Dim blnOK As Boolean
    Dim strRqHdr As String
    
    strWebSite = "https://demopg.hyweb.com.tw/citipos/SSLAuthUI.cgi"
    
    strPara(0) = "MerchantID=0012370888064"
    strPara(1) = "merID=61"
    strPara(2) = "TerminalID=71010001"
    strPara(3) = "MerchantName=PowerTV"
    strPara(4) = "purchAmt=186"
    strPara(5) = "AutoCap=1"
    strPara(6) = "txType=0"
    
    'strURL = strWebSite & "?" & Join(strPara, "&")
    strURL = strWebSite
    strPostData = Join(strPara, "&")
    strMethod = "POST"
    strMethod = "GET"
    strRqHdr = "application/x-www-form-urlencoded"
'    strRqHdr = "text/html; charset=big5"

'    <META http-equiv=Content-Type content="text/html; charset=big5">
'    <META http-equiv="PRAGMA" content="NO-CACHE">
'    <META http-equiv="CACHE-CONTROL" content="NO-CACHE">
    
    If GetXHobj(objXH) Then
        
        With objXH
            If strMethod = "POST" Then
                .Open "POST", strURL, False   ' 使用 HTTP POST
            Else
                .Open "GET", strURL & "?" & strPostData, False  ' 使用 HTTP GET
            End If

'           .SetRequestHeader "Host", "IP位置"
            .setRequestHeader "Content-type", strRqHdr
'           .SetRequestHeader "Content-Length", Len(strPostData)
             
            'strPostData = "MerchantID%3D0012370888064%26merID%3D61%26TerminalID%3D71010001%26MerchantName%3DPowerTV%26purchAmt%3D186%26AutoCap%3D1%26txType%3D0"
            'strPostData = strPostData & "&" & "lidm=1111111111111111111"
             
             On Error Resume Next
            If strMethod = "POST" Then
                .send strPostData ' ?? 將參數編碼成 UTF-8
            Else
                .send
            End If
            
            If Err = 0 Then
                blnOK = (.Status = 200)
            Else
                If Err.Number = -2147024809 Then ' 發生錯誤
                    strErrorMessage = Err.Description
                Else
                    Select Case .Status ' Web site 狀態
                                Case 201
                                        strErrorMessage = "建立 ! ( Created )"
                                Case 202
                                        strErrorMessage = "接受 ! ( Accepted )"
                                Case 207
                                        strErrorMessage = "多重狀態 ! ( Multi-Status )"
                                Case 400
                                        strErrorMessage = "錯誤請求 ! ( Bad Request )"
                                Case 401
                                        strErrorMessage = "未被授權 ! ( Unauthorized )"
                                Case 403
                                        strErrorMessage = "網站禁止 ! ( Forbidden )"
                                Case 404
                                        strErrorMessage = "網址錯誤，無法找到網頁! ( Not found )"
                                Case 500
                                        strErrorMessage = "伺服器錯誤! ( Internal server error )"
                                Case 12029
                                        strErrorMessage = Err.Description '"系統找不到指定的資源。"
                                Case Else
                                        strErrorMessage = Err.Description
                    End Select
                End If
                MsgBox strErrorMessage
                Err.Clear
            End If
            
        End With
        
        If blnOK Then
        
            Dim stm
            
            Set stm = CreateObject("ADODB.Stream")
            
            With stm
                .Mode = 3
                .Type = 2
                .Open
                .Charset = "Unicode"
                .WriteText objXH.responseText
                .Position = 0
'                .Charset = "Unicode"
'                Debug.Print StrConv(.ReadText, vbFromUnicode)
                Debug.Print .ReadText
                .Close
            End With
            
            Set stm = Nothing
        
        End If
        
    End If

End Sub

Private Sub Command1_Click()

    Dim strWebSite As String
    Dim strPara(6) As String
    
    Dim objInet As Object
    Dim strURL As String
    Dim sHeader As String
    Dim strFormData As String
    
    Set objInet = CreateObject("InetCtls.Inet")
    
    strURL = "https://demopg.hyweb.com.tw/citipos/SSLAuthUI.cgi"
    
    strFormData = "MerchantID=0012370888064&merID=61&TerminalID=71010001&MerchantName=PowerTV&purchAmt=186&AutoCap=1&txType=0"
    
    sHeader = "Content-Type: application/x-www-form-urlencoded "
    
    strURL = "http://demopg.hyweb.com.tw/citipos/SSLAuthUI.cgi?MerchantID=0012370888064&merID=61&TerminalID=71010001&MerchantName=PowerTV&purchAmt=186&AutoCap=1&txType=0"
    
    ' "application/x-www-form-urlencoded"
    ' "text/html; charset=big5"
    
    With objInet
        .AccessType = 0
        .url = strURL
        .Protocol = 4
        'MsgBox .Execute(strURL, "POST", strFormData, sHeader)
        CreateObject("Scripting.FileSystemObject").CreateTextFile("C:\123.htm").write .OpenURL(strURL)
    End With
    
    While objInet.StillExecuting
        DoEvents
    Wend
    
    strWebSite = "https://demopg.hyweb.com.tw/citipos/hypage.cgi"
    
    strPara(0) = "pan_no1=1111"
    strPara(1) = "pan_no2=1111"
    strPara(2) = "pan_no3=1111"
    strPara(3) = "pan_no4=1111"
    strPara(4) = "cvc2=111"
    strPara(5) = "expire_month=12"
    strPara(6) = "expire_year=6"
    
    strURL = strWebSite & "?" & Join(strPara, "&")
    
    CreateObject("Scripting.FileSystemObject").CreateTextFile("C:\345.htm").write objInet.OpenURL(strURL)
    
    ' https://demopg.hyweb.com.tw/citipos/hypage.cgi
    ' pan_no1 pan_no2 pan_no3 pan_no4
    ' cvc2
    ' expire_month
    ' expire_year
    
End Sub

Private Sub Command2_Click()
    Dim csPay As Object
    Dim lngAPIreturn As Long
    
    Set csPay = CreateObject("csPayment.clsTWpay")
    
    With csPay
        If .Approve("http://localhost:2011", "901", 0, 7, 1, _
                            "12345678", "4444444444444444", "200808", 1000) = 0 Then
            Debug.Print "取得交易識別碼 GetXID : " & .GetXID ' 取得交易識別碼。
            Debug.Print "取得授權交易之代碼 GetAuthrrpid : " & .GetAuthrrpid ' 取得授權交易之代碼。
            Debug.Print "取得交易之調閱序號 GetTermseq : " & .GetTermseq ' 取得交易之調閱序號。
            Debug.Print "取得交易之調閱編號 GetRetrref : " & .GetRetrref ' 取得交易之調閱編號。
            Debug.Print "批次關閉狀態提示 GetBatchclose : " & .GetBatchclose ' 批次關閉狀態提示。
            Debug.Print "取得授權交易授權碼 GetAuthcode : " & .GetAuthcode ' 取得授權交易授權碼。
            Debug.Print "取得授權金額 GetAmount : " & .GetAmount ' 取得授權金額。
            Debug.Print "取得交易幣值代號 GetCurrency : " & .GetCurrency ' 取得交易幣值代號。
            Debug.Print "取得幣值指數 GetExponent : " & .GetExponent ' 取得幣值指數。
            MsgBox "OK !"
        Else
            ' GetErrStatus : 取得交易的執行狀態。
            ' GetErrCode : 取得交易的錯誤代碼
            Debug.Print "Err Status : " & .GetErrStatus
            Debug.Print "Err Code : " & .GetErrCode
            Debug.Print "API Return : " & .GetAPIreturn
            Debug.Print "API Return Description : " & .GetAPIreturn(True)
            Debug.Print "API Return ( Byref ) : " & lngAPIreturn
            MsgBox "No Good !"
        End If
    End With
    
    Set csPay = Nothing

End Sub

Private Sub Command3_Click()

    'Public Function AUTH_REV(strHyPOS As String, _
                                                strCurrencyType As String, _
                                                intExponentType As Integer, _
                                                intECItype As Integer, _
                                                strMerchantID As String, _
                                                strXID As String, _
                                                strAuthrrpid As String, _
                                                strAuthcode As String, _
                                                strTermseq As String, _
                                                lngAmt As Long) As Boolean

    Dim csPay As Object
    Dim lngAPIreturn As Long
    
    Set csPay = CreateObject("csPayment.clsTWpay")
    
    With csPay
        If .AUTH_REV("http://localhost:2011", "901", 0, 7, 1, _
                                "9B50AB4500000142080_12345678", _
                                "9B50AB4500000142080_12345678", _
                                "017994", "18", 1000) = 0 Then
                                
            Debug.Print "取得交易之調閱編號 GetRetrref : " & .GetRetrref ' 取得交易之調閱編號
            Debug.Print "批次關閉提示 GetBatchclose : " & .GetBatchclose ' 批次關閉提示
            Debug.Print "取得交易之調閱序號 GetTermseq : " & .GetTermseq ' 取得交易之調閱序號
            MsgBox "OK !"
        Else
            ' GetErrStatus : 取得交易的執行狀態。
            ' GetErrCode : 取得交易的錯誤代碼
            Debug.Print "Err Status : " & .GetErrStatus
            Debug.Print "Err Code : " & .GetErrCode
            Debug.Print "API Return : " & .GetAPIreturn
            Debug.Print "API Return Description : " & .GetAPIreturn(True)
            Debug.Print "API Return ( Byref ) : " & lngAPIreturn
            MsgBox "No Good !"
        End If
    End With
    
    Set csPay = Nothing

End Sub

Private Sub Command4_Click()

'    Public Function CAP(strHyPOS As String, _
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

    Dim csPay As Object
    Dim lngAPIreturn As Long
    
    Set csPay = CreateObject("csPayment.clsTWpay")
    
    With csPay
        If .CAP("http://localhost:2011", "901", 0, 7, 1, _
                                "9B50AB4500000142080_12345678", _
                                "9B50AB4500000142080_12345678", _
                                "005063", "18", 1000, 1000) = 0 Then
                                
            Debug.Print "取得交易之調閱編號 GetRetrref : " & .GetRetrref ' 取得交易之調閱編號
            Debug.Print "批次關閉提示 GetBatchclose : " & .GetBatchclose ' 批次關閉提示
            Debug.Print "取得交易之調閱序號 GetTermseq : " & .GetTermseq ' 取得交易之調閱序號
            Debug.Print "交易批次序號 GetBatchId : " & .GetBatchId ' 交易批次編號
            Debug.Print "交易批次序號 GetBatchSeq : " & .GetBatchSeq ' 交易批次序號
            MsgBox "OK !"
        Else
            ' GetErrStatus : 取得交易的執行狀態。
            ' GetErrCode : 取得交易的錯誤代碼
            Debug.Print "Err Status : " & .GetErrStatus
            Debug.Print "Err Code : " & .GetErrCode
            Debug.Print "API Return : " & .GetAPIreturn
            Debug.Print "API Return Description : " & .GetAPIreturn(True)
            Debug.Print "API Return ( Byref ) : " & lngAPIreturn
            MsgBox "No Good !"
        End If
    End With
    
    Set csPay = Nothing

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then End
End Sub

