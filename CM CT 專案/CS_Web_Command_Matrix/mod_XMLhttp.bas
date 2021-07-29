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
                .Open "POST", strURL & strCMD, False  ' ¨Ï¥Î HTTP POST
'               .SetRequestHeader "Host", "IP¦ì¸m"
                .setRequestHeader "Content-type", strRequestHeader
'                .SetRequestHeader "Content-Length", Len(strPostData)
            Else
                .Open "GET", strURL & strCMD & "?" & strPostData, False  ' ¨Ï¥Î HTTP GET
                .setRequestHeader "Content-type", strRequestHeader
            End If
             On Error Resume Next
            If strMethod = "POST" Then
                .Send Encode_2_UTF8(strPostData) ' ±N°Ñ¼Æ½s½X¦¨ UTF-8
            Else
                .Send
            End If
            
            If Err.Number = 0 Then
                Call_Web_Service = (.Status = 200)
            Else
                If Err.Number = -2147024809 Then ' µo¥Í¿ù»~
                    strErrorMessage = Err.Description
                Else
                    Select Case .Status ' Web site ª¬ºA
                                Case 201
                                        strErrorMessage = "«Ø¥ß ! ( Created )"
                                Case 202
                                        strErrorMessage = "±µ¨ü ! ( Accepted )"
                                Case 207
                                        strErrorMessage = "¦h­«ª¬ºA ! ( Multi-Status )"
                                Case 400
                                        strErrorMessage = "¿ù»~½Ð¨D ! ( Bad Request )"
                                Case 401
                                        strErrorMessage = "¥¼³Q±ÂÅv ! ( Unauthorized )"
                                Case 403
                                        strErrorMessage = "ºô¯¸¸T¤î ! ( Forbidden )"
                                Case 404
                                        strErrorMessage = "ºô§}¿ù»~¡AµLªk§ä¨ìºô­¶! ( Not found )"
                                Case 500
                                        strErrorMessage = "¦øªA¾¹¿ù»~! ( Internal server error )"
                                Case 12029
                                        strErrorMessage = Err.Description '"¨t²Î§ä¤£¨ì«ü©wªº¸ê·½¡C"
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
'            Debug.Print "Spend Time : " & GetTickCount - lngST ' ©I¥s Web site ©Òªá®É¶¡
            Call_Web_Service = mod_XML.ParseXML(XMLobj.ResponseText)
'            Debug.Print "Fields count : " & rsVirtual.Fields.Count
'            Debug.Print "Record count : " & rsVirtual.RecordCount
'            Debug.Print "Data : " & vbCrLf & rsVirtual.GetString(2, , "¡A", vbCrLf, "")
        Else
'            MsgBox strErrorMessage, vbInformation, "°T®§"
        End If
    Else
'        MsgBox strErrorMessage, vbInformation, "°T®§"
    End If
    On Error Resume Next
    Set XMLobj = Nothing
  Exit Function
ChkErr:
    ErrHandle "mod_XMLhttp", "Call_Web_Service"
End Function

'XMLHTTP ª«¥ó¤Î¨ä¤èªkÂ²¤¶

'MSXML¤¤´£¨Ñ¤FMicrosoft.XMLHTTPª«¥ó¡A¯à°÷§¹¦¨±q¸ê®Æ¥]¨ìRequestª«¥óªºÂà´«¥H¤Îµo°e¥ô°È¡C

'ª«¥ó«Ø¥ß«á½Õ¥ÎOpen¤èªk¹ïRequestª«¥ó¶i¦æªì©l¤Æ¡A»yªk®æ¦¡úd¡G

'poster.Open http - method, url, async, userID, password
    
    'Open¤èªk¤¤¥]§t¤F5­Ó°Ñ¼Æ¡A«e¤T­Ó¬O¥²­nªº¡A«á¨â­Ó¬O¥i¿ïªº(¦b¦øªA¾¹»Ý­n¶i¦æ¨­¥÷ÅçÃÒ®É´£¨Ñ)¡C°Ñ¼Æªº§t¸q¦p¤U©Ò¥Ü¡G
    'http-method¡G HTTPªº³q«H¤è¦¡¡A¤ñ¦pGET©Î¬O POST
    'url¡G ±µ¦¬XML¸ê®Æªº¦øªA¾¹ªºURL¦ì§}¡C³q±`¦bURL¤¤­n«ü©ú ASP©ÎCGIµ{¦¡
    'async¡G ¤@­Ó¥¬ªL¼ÐÃÑ¡A»¡©ú½Ð¨D¬O§_úd«D¦P¨Bªº¡C¦pªG¬O«D¦P¨B³q«H¤è¦¡(true)¡A
    '«È¤á¾÷´N¤£µ¥«Ý¦øªA¾¹ªº¦^À³¡F¦pªG¬O¦P¨B¤è¦¡(false)¡A
    '«È¤á¾÷´N­nµ¥¨ì¦øªA¾¹ªð¦^®ø®§«á¤~¥h°õ¦æ¨ä¥L¾Þ§@
    'userID ¥Î¤áID, ¥Î©ó¦øªA¾¹¨­¥÷ÅçÃÒ
    'password ¥Î¤á±K½X¡A¥Î©ó¦øªA¾¹¨­¥÷ÅçÃÒ&nbsp;

'XMLHTTPª«¥óªºSend¤èªk
'¥ÎOpen¤èªk¹ïRequestª«¥ó¶i¦æªì©l¤Æ«á , ½Õ¥ÎSend¤èªkµo°eXML¸ê®Æ:

    'poster.send XML - Data
    'Send¤èªkªº°Ñ¼ÆÃþ«¬¬OVariant¡A¥i¥H¬O¦r¦ê¡BDOM¾ð©Î¥ô·N¸ê®Æ¬yµ{¡Cµo°e¸ê®Æªº¤è¦¡¤Àúd¦P¨B©M«D¦P¨B¨âºØ¡C
    '¦b«D¦P¨B¤è¦¡¤U¡A¸ê®Æ¥]¤@¥¹µo°e§¹²¦¡A´Nµ²§ôSend¶iµ{¡A«È¤á¾÷°õ¦æ¨ä¥Lªº¾Þ§@¡F
    '¦Ó¦b¦P¨B¤è¦¡¤U¡A«È¤á¾÷­nµ¥¨ì¦øªA¾¹ªð¦^½T»{®ø®§«á¤~µ²§ôSend¶iµ{¡C
    'XMLHTTPª«¥ó¤¤ªºreadyStateÄÝ©Ê¯à°÷¤Ï¬M¥X¦øªA¾¹¦b³B²z½Ð¨D®Éªº¶i®iª¬ªp¡C
    '«È¤á¾÷ªºµ{¦¡¥i¥H®Ú¾Ú³o­Óª¬ºA¸ê°T³]¸m¬ÛÀ³ªº¨Æ¥ó³B²z¤èªk¡CÄÝ©Ê­È¤Î¨ä§t¸q¦p¤Uªí©Ò¥Ü¡G
    '­È »¡©ú
    '0 Responseª«¥ó¤w¸g«Ø¥ß , ¦ýXML¤åÀÉ¤W¸ü¹Lµ{©|¥¼µ²§ô
    '1 XML¤åÀÉ¤w¸g¸Ë¸ü§¹²¦
    '2 XML¤åÀÉ¤w¸g¸Ë¸ü§¹²¦ , ¥¿¦b³B²z¤¤
    '3 ³¡¤ÀXML¤åÀÉ¤w¸g¸ÑªR
    '4 ¤åÀÉ¤w¸g¸ÑªR§¹²¦ , ¥Î¤áºÝ¥i¥H±µ¨üªð¦^®ø®§

'«È¤á¾÷³B²z¦^À³¸ê°T
'«È¤á¾÷±µ¦¬¨ìªð¦^®ø®§«á¡A¶i¦æÂ²³æªº³B²z¡A°ò¥»¤W´N§¹¦¨¤FC/S¤§¶¡ªº¤@­Ó¥æ¤¬©P´Á¡C
'«È¤á¾÷±µ¦¬¦^À³¬O³q¹LXMLHTTPª«¥óªºÄÝ©Ê¹ê²{ªº¡G
    '¡´ responseTxt¡G±Nªð¦^®ø®§§@úd¤å¥»¦r¦ê¡F
    '¡´ responseXML¡G±Nªð¦^®ø®§µøúdXML¤åÀÉ¡A¦b¦øªA¾¹¦^À³®ø®§¤¤§t¦³XML¸ê®Æ®É¨Ï¥Î¡F
    '¡´ responseStream¡G±Nªð¦^®ø®§µøúdStreamª«¥ó¡C

