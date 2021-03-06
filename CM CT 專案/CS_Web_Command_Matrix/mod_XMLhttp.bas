Attribute VB_Name = "mod_XMLhttp"
Option Explicit

Public strURL As String
Public strPostData As String
Public strCMD As String

'http://210.202.146.97:9090/CMWebServiceGWDemo/EMCWebServiceGW.asmx

Private Const strRequestHeader = "application/x-www-form-urlencoded"

'Private Const strURL = "http://localhost/ws1/Service1.asmx/GetCustInfo"
'Private Const strPostData = "strTableName=SO1100A&strCustID=12"

'    Call GetDOMdocObj(DOMobj)
'    Set DOMobj = XMLobj.ResponseXML
'    Debug.Print DOMobj.XML

'setRequestHeader('Content-Type', 'application/x-www-form-urlencoded; charset=UTF-8');
'msxmlCHT.msi
'MSXML 4.0 Service Pack 2 (Microsoft XML Core Services)
'http://www.microsoft.com/downloads/details.aspx?FamilyID=3144b72b-b4f2-46da-b4b6-c5d7485f2b42&DisplayLang=zh-tw

Public Function Call_Web_Service(strMethod As String) As Boolean
  On Error GoTo ChkErr
    Dim XMLobj As Object
    Dim DOMobj As Object
    Dim lngST As Long
    Call_Web_Service = False
    lngST = GetTickCount
    Dim strFakeRequest As String

    If GetXMLobj(XMLobj) Then
        With XMLobj
            If strMethod = "POST" Then
                .Open "POST", strURL & strCMD, False  ' 使用 HTTP POST
'               .SetRequestHeader "Host", "IP位置"
                .setRequestHeader "Content-type", strRequestHeader
'                .SetRequestHeader "Content-Length", Len(strPostData)
            Else
                .Open "GET", strURL & strCMD & "?" & strPostData, False  ' 使用 HTTP GET
                .setRequestHeader "Content-type", strRequestHeader
            End If
             On Error Resume Next
            If strMethod = "POST" Then
                .Send Encode_2_UTF8(strPostData) ' 將參數編碼成 UTF-8
            Else
                .Send
            End If
            
            If Err.Number = 0 Then
                Call_Web_Service = (.Status = 200)
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
                Err.Clear
            End If
        End With
        On Error GoTo ChkErr
        If Call_Web_Service Then
'            Debug.Print XMLobj.ResponseText
'            Debug.Print "Spend Time : " & GetTickCount - lngST ' 呼叫 Web site 所花時間
            Call_Web_Service = mod_XML.ParseXML(XMLobj.ResponseText)
'            Debug.Print "Fields count : " & rsVirtual.Fields.Count
'            Debug.Print "Record count : " & rsVirtual.RecordCount
'            Debug.Print "Data : " & vbCrLf & rsVirtual.GetString(2, , "，", vbCrLf, "")
        Else
'            MsgBox strErrorMessage, vbInformation, "訊息"
        End If
    Else
'        MsgBox strErrorMessage, vbInformation, "訊息"
    End If
    On Error Resume Next
    Set XMLobj = Nothing
  Exit Function
ChkErr:
    ErrHandle "mod_XMLhttp", "Call_Web_Service"
End Function

'XMLHTTP 物件及其方法簡介

'MSXML中提供了Microsoft.XMLHTTP物件，能夠完成從資料包到Request物件的轉換以及發送任務。

'物件建立後調用Open方法對Request物件進行初始化，語法格式?d：

'poster.Open http - method, url, async, userID, password
    
    'Open方法中包含了5個參數，前三個是必要的，後兩個是可選的(在伺服器需要進行身份驗證時提供)。參數的含義如下所示：
    'http-method： HTTP的通信方式，比如GET或是 POST
    'url： 接收XML資料的伺服器的URL位址。通常在URL中要指明 ASP或CGI程式
    'async： 一個布林標識，說明請求是否?d非同步的。如果是非同步通信方式(true)，
    '客戶機就不等待伺服器的回應；如果是同步方式(false)，
    '客戶機就要等到伺服器返回消息後才去執行其他操作
    'userID 用戶ID, 用於伺服器身份驗證
    'password 用戶密碼，用於伺服器身份驗證&nbsp;

'XMLHTTP物件的Send方法
'用Open方法對Request物件進行初始化後 , 調用Send方法發送XML資料:

    'poster.send XML - Data
    'Send方法的參數類型是Variant，可以是字串、DOM樹或任意資料流程。發送資料的方式分?d同步和非同步兩種。
    '在非同步方式下，資料包一旦發送完畢，就結束Send進程，客戶機執行其他的操作；
    '而在同步方式下，客戶機要等到伺服器返回確認消息後才結束Send進程。
    'XMLHTTP物件中的readyState屬性能夠反映出伺服器在處理請求時的進展狀況。
    '客戶機的程式可以根據這個狀態資訊設置相應的事件處理方法。屬性值及其含義如下表所示：
    '值 說明
    '0 Response物件已經建立 , 但XML文檔上載過程尚未結束
    '1 XML文檔已經裝載完畢
    '2 XML文檔已經裝載完畢 , 正在處理中
    '3 部分XML文檔已經解析
    '4 文檔已經解析完畢 , 用戶端可以接受返回消息

'客戶機處理回應資訊
'客戶機接收到返回消息後，進行簡單的處理，基本上就完成了C/S之間的一個交互周期。
'客戶機接收回應是通過XMLHTTP物件的屬性實現的：
    '● responseTxt：將返回消息作?d文本字串；
    '● responseXML：將返回消息視?dXML文檔，在伺服器回應消息中含有XML資料時使用；
    '● responseStream：將返回消息視?dStream物件。

