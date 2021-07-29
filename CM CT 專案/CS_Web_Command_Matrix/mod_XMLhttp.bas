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
                .Open "POST", strURL & strCMD, False  ' �ϥ� HTTP POST
'               .SetRequestHeader "Host", "IP��m"
                .setRequestHeader "Content-type", strRequestHeader
'                .SetRequestHeader "Content-Length", Len(strPostData)
            Else
                .Open "GET", strURL & strCMD & "?" & strPostData, False  ' �ϥ� HTTP GET
                .setRequestHeader "Content-type", strRequestHeader
            End If
             On Error Resume Next
            If strMethod = "POST" Then
                .Send Encode_2_UTF8(strPostData) ' �N�Ѽƽs�X�� UTF-8
            Else
                .Send
            End If
            
            If Err.Number = 0 Then
                Call_Web_Service = (.Status = 200)
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
                Err.Clear
            End If
        End With
        On Error GoTo ChkErr
        If Call_Web_Service Then
'            Debug.Print XMLobj.ResponseText
'            Debug.Print "Spend Time : " & GetTickCount - lngST ' �I�s Web site �Ҫ�ɶ�
            Call_Web_Service = mod_XML.ParseXML(XMLobj.ResponseText)
'            Debug.Print "Fields count : " & rsVirtual.Fields.Count
'            Debug.Print "Record count : " & rsVirtual.RecordCount
'            Debug.Print "Data : " & vbCrLf & rsVirtual.GetString(2, , "�A", vbCrLf, "")
        Else
'            MsgBox strErrorMessage, vbInformation, "�T��"
        End If
    Else
'        MsgBox strErrorMessage, vbInformation, "�T��"
    End If
    On Error Resume Next
    Set XMLobj = Nothing
  Exit Function
ChkErr:
    ErrHandle "mod_XMLhttp", "Call_Web_Service"
End Function

'XMLHTTP ����Ψ��k²��

'MSXML�����ѤFMicrosoft.XMLHTTP����A��������q��ƥ]��Request�����ഫ�H�εo�e���ȡC

'����إ߫�ե�Open��k��Request����i���l�ơA�y�k�榡�d�G

'poster.Open http - method, url, async, userID, password
    
    'Open��k���]�t�F5�ӰѼơA�e�T�ӬO���n���A���ӬO�i�諸(�b���A���ݭn�i�樭�����Үɴ���)�C�Ѽƪ��t�q�p�U�ҥܡG
    'http-method�G HTTP���q�H�覡�A��pGET�άO POST
    'url�G ����XML��ƪ����A����URL��}�C�q�`�bURL���n���� ASP��CGI�{��
    'async�G �@�ӥ��L���ѡA�����ШD�O�_�d�D�P�B���C�p�G�O�D�P�B�q�H�覡(true)�A
    '�Ȥ���N�����ݦ��A�����^���F�p�G�O�P�B�覡(false)�A
    '�Ȥ���N�n������A����^������~�h�����L�ާ@
    'userID �Τ�ID, �Ω���A����������
    'password �Τ�K�X�A�Ω���A����������&nbsp;

'XMLHTTP����Send��k
'��Open��k��Request����i���l�ƫ� , �ե�Send��k�o�eXML���:

    'poster.send XML - Data
    'Send��k���Ѽ������OVariant�A�i�H�O�r��BDOM��Υ��N��Ƭy�{�C�o�e��ƪ��覡���d�P�B�M�D�P�B��ءC
    '�b�D�P�B�覡�U�A��ƥ]�@���o�e�����A�N����Send�i�{�A�Ȥ�������L���ާ@�F
    '�Ӧb�P�B�覡�U�A�Ȥ���n������A����^�T�{������~����Send�i�{�C
    'XMLHTTP���󤤪�readyState�ݩʯ���ϬM�X���A���b�B�z�ШD�ɪ��i�i���p�C
    '�Ȥ�����{���i�H�ھڳo�Ӫ��A��T�]�m�������ƥ�B�z��k�C�ݩʭȤΨ�t�q�p�U��ҥܡG
    '�� ����
    '0 Response����w�g�إ� , ��XML���ɤW���L�{�|������
    '1 XML���ɤw�g�˸�����
    '2 XML���ɤw�g�˸����� , ���b�B�z��
    '3 ����XML���ɤw�g�ѪR
    '4 ���ɤw�g�ѪR���� , �Τ�ݥi�H������^����

'�Ȥ���B�z�^����T
'�Ȥ���������^������A�i��²�檺�B�z�A�򥻤W�N�����FC/S�������@�ӥ椬�P���C
'�Ȥ�������^���O�q�LXMLHTTP�����ݩʹ�{���G
    '�� responseTxt�G�N��^�����@�d�奻�r��F
    '�� responseXML�G�N��^�������dXML���ɡA�b���A���^���������t��XML��ƮɨϥΡF
    '�� responseStream�G�N��^�������dStream����C

