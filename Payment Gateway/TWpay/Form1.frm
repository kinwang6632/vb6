VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  '���u�T�w��ܤ��
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
   StartUpPosition =   2  '�ù�����
   Begin VB.CommandButton Command5 
      Caption         =   " ���� �H�Υd�дڧ@�~ ( CAP_REV )"
      Height          =   375
      Left            =   270
      TabIndex        =   8
      Top             =   2880
      Width           =   4095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "�H�Υd�дڧ@�~ ( CAP )"
      Height          =   375
      Left            =   270
      TabIndex        =   7
      Top             =   2430
      Width           =   4095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "�������v ( AUTH_REV )"
      Height          =   375
      Left            =   270
      TabIndex        =   6
      Top             =   1980
      Width           =   4095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "�� �v ( Auth )"
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
      Caption         =   "�� �v ( Auth )"
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

' �� �v
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
    

'    Exponent ���ȫ��� Varchar2 (5)
'    Xid ����ѧO�X Varchar2(60)
'    Authrrpid SSL ���v������N�X Varchar2(60)
'    AuthCode ������v�X Varchar2(10)
'    TermSeq �վ\�Ǹ� Varchar2(12)
'    BatchId ����妸�s��  Varchar2(22)
'    BatchSeq ����妸�Ǹ�  Varchar2(22)
    
    
    Dim csPay As Object
    Dim lngAPIreturn As Long
    
    Set csPay = CreateObject("csPayment.clsTWpay")
    
'    ' �� �v
'    Public Function Auth(strOrderNo As String, strCardNo As String, strDate As String, lngAmt As Long) As Boolean
'       �q��s��    �H�Υd��    �H�Υd���Ĵ���  ������v�`���B
    With csPay
        .InitializePara "http://localhost:2011", "901", 0, 7, 1
        ' HyPOS
        ' Currency
        ' Exponent
        ' ECI
        ' MerID
        If .AUTH("12345678", "4444444444444444", "200808", 1000, lngAPIreturn) Then
            Debug.Print "���o����ѧO�X GetXID : " & .GetXID ' ���o����ѧO�X�C
            Debug.Print "���o���v������N�X GetAuthrrpid : " & .GetAuthrrpid ' ���o���v������N�X�C
            Debug.Print "���o������վ\�Ǹ� GetTermseq : " & .GetTermseq ' ���o������վ\�Ǹ��C
            Debug.Print "���o������վ\�s�� GetRetrref : " & .GetRetrref ' ���o������վ\�s���C
            Debug.Print "�妸�������A���� GetBatchclose : " & .GetBatchclose ' �妸�������A���ܡC
            Debug.Print "���o���v������v�X GetAuthcode : " & .GetAuthcode ' ���o���v������v�X�C
            Debug.Print "���o���v���B GetAmount : " & .GetAmount ' ���o���v���B�C
            Debug.Print "���o������ȥN�� GetCurrency : " & .GetCurrency ' ���o������ȥN���C
            Debug.Print "���o���ȫ��� GetExponent : " & .GetExponent ' ���o���ȫ��ơC
            MsgBox "OK !"
        Else
            ' GetErrStatus : ���o��������檬�A�C
            ' GetErrCode : ���o��������~�N�X
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

'���յ��G
'
'1. URL �覡
'    �� XMLHTTP ������� , GET �覡 OK , Post �覡 , �|�ѼƤ���
'    �B���������� 2 ~ 3 �� , �B��������� , �ҥH�������\
'    ���D����� , �_�h�ثe�S��k�ϥ� URL �覡
'
'2. API �覡
'    ���ӬO�i�H�� , ���I�ɶ� Try , �N�|�����G
'    ���L�����ҤΦw�ˤW�����D


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
                .Open "POST", strURL, False   ' �ϥ� HTTP POST
            Else
                .Open "GET", strURL & "?" & strPostData, False  ' �ϥ� HTTP GET
            End If

'           .SetRequestHeader "Host", "IP��m"
            .setRequestHeader "Content-type", strRqHdr
'           .SetRequestHeader "Content-Length", Len(strPostData)
             
            'strPostData = "MerchantID%3D0012370888064%26merID%3D61%26TerminalID%3D71010001%26MerchantName%3DPowerTV%26purchAmt%3D186%26AutoCap%3D1%26txType%3D0"
            'strPostData = strPostData & "&" & "lidm=1111111111111111111"
             
             On Error Resume Next
            If strMethod = "POST" Then
                .send strPostData ' ?? �N�Ѽƽs�X�� UTF-8
            Else
                .send
            End If
            
            If Err = 0 Then
                blnOK = (.Status = 200)
            Else
                If Err.Number = -2147024809 Then ' �o�Ϳ��~
                    strErrorMessage = Err.Description
                Else
                    Select Case .Status ' Web site ���A
                                Case 201
                                        strErrorMessage = "�إ� ! ( Created )"
                                Case 202
                                        strErrorMessage = "���� ! ( Accepted )"
                                Case 207
                                        strErrorMessage = "�h�����A ! ( Multi-Status )"
                                Case 400
                                        strErrorMessage = "���~�ШD ! ( Bad Request )"
                                Case 401
                                        strErrorMessage = "���Q���v ! ( Unauthorized )"
                                Case 403
                                        strErrorMessage = "�����T�� ! ( Forbidden )"
                                Case 404
                                        strErrorMessage = "���}���~�A�L�k������! ( Not found )"
                                Case 500
                                        strErrorMessage = "���A�����~! ( Internal server error )"
                                Case 12029
                                        strErrorMessage = Err.Description '"�t�Χ䤣����w���귽�C"
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
            Debug.Print "���o����ѧO�X GetXID : " & .GetXID ' ���o����ѧO�X�C
            Debug.Print "���o���v������N�X GetAuthrrpid : " & .GetAuthrrpid ' ���o���v������N�X�C
            Debug.Print "���o������վ\�Ǹ� GetTermseq : " & .GetTermseq ' ���o������վ\�Ǹ��C
            Debug.Print "���o������վ\�s�� GetRetrref : " & .GetRetrref ' ���o������վ\�s���C
            Debug.Print "�妸�������A���� GetBatchclose : " & .GetBatchclose ' �妸�������A���ܡC
            Debug.Print "���o���v������v�X GetAuthcode : " & .GetAuthcode ' ���o���v������v�X�C
            Debug.Print "���o���v���B GetAmount : " & .GetAmount ' ���o���v���B�C
            Debug.Print "���o������ȥN�� GetCurrency : " & .GetCurrency ' ���o������ȥN���C
            Debug.Print "���o���ȫ��� GetExponent : " & .GetExponent ' ���o���ȫ��ơC
            MsgBox "OK !"
        Else
            ' GetErrStatus : ���o��������檬�A�C
            ' GetErrCode : ���o��������~�N�X
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
                                
            Debug.Print "���o������վ\�s�� GetRetrref : " & .GetRetrref ' ���o������վ\�s��
            Debug.Print "�妸�������� GetBatchclose : " & .GetBatchclose ' �妸��������
            Debug.Print "���o������վ\�Ǹ� GetTermseq : " & .GetTermseq ' ���o������վ\�Ǹ�
            MsgBox "OK !"
        Else
            ' GetErrStatus : ���o��������檬�A�C
            ' GetErrCode : ���o��������~�N�X
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
                                
            Debug.Print "���o������վ\�s�� GetRetrref : " & .GetRetrref ' ���o������վ\�s��
            Debug.Print "�妸�������� GetBatchclose : " & .GetBatchclose ' �妸��������
            Debug.Print "���o������վ\�Ǹ� GetTermseq : " & .GetTermseq ' ���o������վ\�Ǹ�
            Debug.Print "����妸�Ǹ� GetBatchId : " & .GetBatchId ' ����妸�s��
            Debug.Print "����妸�Ǹ� GetBatchSeq : " & .GetBatchSeq ' ����妸�Ǹ�
            MsgBox "OK !"
        Else
            ' GetErrStatus : ���o��������檬�A�C
            ' GetErrCode : ���o��������~�N�X
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

