Attribute VB_Name = "mod_Func24"
Option Explicit

' Date : 2006/12/25 �t�ϸ`
' Designer : Power Hammer
' Doc : Web API����_20061225_Jacy.doc

Private cn As Object
Private strXMLresult As String
Private strCustID As String
Private strCompCode As String
Private strOwner As String
Private strDBLink As String
Private strDbg As String

Public Function JustDoIt(ByRef objWebPool As Object, _
                                        ByRef strPara As String) As String
  On Error GoTo ChkErr
    Dim varAry As Variant
    Dim varPara As Variant
    Dim varElement As Variant
    Dim objCnPool As Object
    Dim rs As Object
    Dim strQry As String
    Dim strChargeAddr As String
    Dim strMailAddr As String
    
    Set objCnPool = objWebPool
    Set cn = objCnPool.GetConnection
    
    If cn.state <= 0 Then
        Set cn = objCnPool.GetConnection
        If cn.state <= 0 Then
            strErr = "�L�k�s�u��Ʈw!"
            JustDoIt = "-99,��Ʈw�s�u����!"
            GoTo 66
        End If
    End If
    
'    ETHOME -CM������T�d��
'    �ǤJ�Ѽ�[24, �ӽФH�m�W, �������� ( 1:������;  2:�Τ@�s��), �����Ҧr�� �� �Τ@�s��,
'    �]�ƺ��� (1: MAC; 2:�b��) ,  CM MAC �� ACS �b��]

'    If RecordSet���� = 0 Then
'       �^��(-20, �d�L�ŦX���󤧸��)
'    End If
'    If RecordSet���� > 0 Then
'      ����XML��.(��select�X�Ӥ����)    (�Ѧ�: ���G)
'      �^��(0, ���\)��XML��
'    Else
'       �^��(-20, �d�L�ŦX���󤧸��)
'    End If

    varPara = Split(strPara, ",")
    
    If UBound(varPara) <> 5 Then '    If �ǤJ�Ѽ� <> 6 Then
        JustDoIt = "-99,�ѼƤ���!" '       �^��(-99, �ѼƤ���)
        GoTo 66
    End If '    End If

'    0,���\
'    <?xml version="1.0" encoding="BIG5" standalone="yes" ?>
'    <DataSet>
'      <DataTable>
'        <DataRow SYSID="9" SYSNAME="�����s" CUSTID="2" DeclarantName="���j�L" DialAccount ="a1234" FACISNO="009096D5CC2A" ID="A123456789" Instaddress="�x�_��������100��" ChargeAddress="�x�_��������104��" MailAddress="�x�_��������104��" ContTel="87910168" Contmobile="0935123456" ContTel2="87642312" AgentTel="78324764" OPENSTATUS="�}��" ACCOUNTSTATUS="�n�}" BPNAME="CM-���u���:2M�b�~ú1800��" DYNIPCOUNT="4"/>
'      </DataTable>
'    </DataSet>�

    If varPara(4) = 1 Then
    '    1.�Y�]�ƺ���=1.MAC:
    '     ����CATVN��COM��,�HMAC�dSO306�����q�O,
    '     �A�H�Ӥ��q�O������COM�Ϫ�SO507 , �d�֦p��s��SO����T
    '     �s���SO , �H�ӥӽФH�m�W, �����Ҧr�� / �Τ@�s��, MAC�d�� '���`' CM�]��
        strDbg = "SELECT COMPCODE FROM COM.SO306 WHERE HFCMAC='" & varPara(5) & "'"
        strCompCode = GetValue(cn, strDbg) & Empty
                                                
    Else
    '    2. �Y�]�ƺ���=2.�b��:
    '     ����CATVN��COM��,�HMAC�dSO305A�����q�O(�s�WTABLE: SO305A)
    '     (�`�N������PM��INSERT,����Ѥp���C���צU�t�Υx���b����,�P�ɶפ@���ܸ�TABLE),
    '     �A�H�Ӥ��q�O������COM�Ϫ�SO507 , �d�֦p��s��SO����T
    '     �s���SO , �H�ӥӽФH�m�W, �����Ҧr�� / �Τ@�s��, �b���d�� '���`'CM�]��
        strDbg = "SELECT COMPCODE FROM COM.SO305A WHERE DIALACCOUNT='" & varPara(5) & "'"
        strCompCode = GetValue(cn, strDbg) & Empty
    End If

    If strCompCode = Empty Then GoTo 99
    
    strDbg = "SELECT DESCRIPTION,LINK FROM COM.SO507 WHERE CODENO=" & strCompCode
    Set rs = cn.Execute(strDbg)
    
    If rs.EOF Then GoTo 99
    
    strOwner = Trim(rs("Description") & "")
    strDBLink = Trim(rs("Link") & "")
    
    If Len(strOwner) > 0 Then strOwner = strOwner & "."
    If Len(strDBLink) > 0 Then strDBLink = "@" & strDBLink

'    �ǤJ�Ѽ�[24, �ӽФH�m�W, �������� ( 1:������;  2:�Τ@�s��), �����Ҧr�� �� �Τ@�s��,
'    �]�ƺ��� (1: MAC; 2:�b��) ,  CM MAC �� ACS �b��]
            
'    ��=�ǤJ��"�ӽФH�m�W"
'    �Y�]�ƺ���=2(�b��) , ��=�ǤJ��"�b��"
'    �Y�]�ƺ���=1(MAC) , ��=�ǤJ��"MAC"
'    �s�W��� SO004.ContMobile

    strQry = "SELECT COMPCODE,CUSTID,DECLARANTNAME,DECLARANTNAME" & _
                    ",DIALACCOUNT,FACISNO,ID,ID2,CONTTEL,CONTMOBILE,CONTTEL2" & _
                    ",AGENTTEL,CMOPENDATE,CMCLOSEDATE,ENABLEACCOUNT" & _
                    ",DISABLEACCOUNT, BPCODE, BPNAME,IDKINDCODE1" & _
                    " FROM " & strOwner & "SO004" & strDBLink & _
                    " WHERE DECLARANTNAME='" & varPara(1) & "'" & _
                    " AND " & IIf(varPara(4) = 1, "FACISNO", "DIALACCOUNT") & "='" & varPara(5) & "'" & _
                    " AND ( PRDATE IS NULL OR INSTDATE > PRDATE)"
                    
    strDbg = strQry
    Set rs = cn.Execute(strQry)
    
    If rs.EOF Then GoTo 99
    
    'If ChkID(rs("IDKindCode1") & "", Val(varPara(2))) Then
    '    If varPara(2) & "" <> rs("ID") & "" Then GoTo 99
    'End If
    If varPara(3) & "" <> rs("ID") & "" And varPara(3) & "" <> rs("ID2") & "" Then GoTo 99
    
    While Not rs.EOF
        strCustID = rs("CustID")
        If Not GetAddress(rs("FaciSno") & "", strChargeAddr, strMailAddr) Then GoTo 99
        MakeXML strCompCode, GetCompName, strCustID, _
                        rs("DeclarantName") & "", rs("DialAccount") & "", _
                        rs("FaciSno") & "", varPara(3) & "", _
                        GetInstAddr, strChargeAddr, strMailAddr, _
                        rs("ContTel") & "", rs("ContMobile") & "", rs("ContTel2") & "", rs("AgentTel") & "", _
                        GetOpenStatus(rs("CMOpenDate") & "", rs("CMCloseDate") & ""), _
                        GetAccStatus(rs("EnableAccount") & "", rs("DisableAccount") & ""), _
                        rs("BPName") & "", GetDynIPcnt(rs("BPCode") & "")
        rs.MoveNext
    Wend
    
    JustDoIt = "0,���\" & vbCrLf & _
                        "<?xml version=""1.0"" encoding=""BIG5"" standalone=""yes"" ?>" & vbCrLf & _
                        "<DataSet>" & vbCrLf & _
                        "  <DataTable>" & vbCrLf & _
                        strXMLresult & _
                        "  </DataTable>" & vbCrLf & _
                        "</DataSet>" & vbCrLf ' �^��(0, ���\) ��XML��
    GoTo 66

'    ���G :  (API-24: ETHOME-CM������T�d��)

'   0:  ���\
'   -99 : �ѼƤ���
'   -20, �d�L�ŦX���󤧸��
    
'    PS:
'        XML (�t�Υx�N��, �t�Υx�W��, �Ƚs, �Τ�W��, ACS�����b��, CM MAC, �����Ҧr��/�νs, �˾��a�}, ���O�a�}, �l�H�a�}, �ӽФH�p���q��, �ӽФH��ʹq��, �p���H�p���q��, �N�z�H�p���q��, �}�q���A,�b�����A,�ӽФ�קO, IP��)
    
99:
    JustDoIt = "-20, �d�L�ŦX���󤧸��!" ' -20, �d�L�ŦX���󤧸��
    'JustDoIt = "-20, �d�L�ŦX���󤧸��!" & vbCrLf & strDbg ' -20, �d�L�ŦX���󤧸��
    
66:
    On Error Resume Next
    Rlx varAry
    Rlx varPara
    Rlx varElement
    Rlx strCustID
    Rlx strXMLresult
    Rlx strOwner
    Rlx strDBLink
    Rlx objCnPool
    Rlx strCompCode
    Rlx strChargeAddr
    Rlx strMailAddr
    Rlx strDbg
    Set cn = Nothing
    Set objCnPool = Nothing
  Exit Function
ChkErr:
    ErrHandle "mod_Func24", "JustDoIt"
    JustDoIt = "-99," & strErr & vbCrLf & strDbg
End Function

'' �����Ҧr�� / �νs(ID)
'Private Function ChkID(strIDKindCode As String, intRefNo As Integer) As Boolean
'  On Error GoTo ChkErr
'    ChkID = Val(GetValue(cn, "SELECT COUNT(*) FROM " & strOwner & "CD070" & strDBLink & _
'                                        " WHERE CODENO=" & strIDKindCode & _
'                                        " AND REFNO=" & intRefNo) & Empty) > 0
'
''    �HSO004.IDKindCode1 ������CD070.CODENO,
''    a.�Y�ǤJ����������=1(�����Ҧr��) CD070.REFNO = 1
''    b.�Y�ǤJ����������=2(�Τ@�s��) CD070.REFNO = 2
''    --CD070.REFNO�ݲŦX,�~�i���SO004.ID=�ǤJ��(�����Ҧr��/�Τ@�s��)
'
'  Exit Function
'ChkErr:
'    ErrHandle "mod_Func24", "ChkID"
'End Function

' �t�Υx�W�� (SYSNAME)
Private Function GetCompName() As String
  On Error GoTo ChkErr
    ' �HSO004.COMPCODE��LINK
    strDbg = "SELECT DESCRIPTION FROM " & strOwner & "CD039" & strDBLink & _
                    " WHERE CODENO=" & strCompCode
    GetCompName = GetValue(cn, strDbg) & Empty
  Exit Function
ChkErr:
    ErrHandle "mod_Func24", "GetCompName"
End Function

' �˾��a�} (Instaddress)
Private Function GetInstAddr() As String
  On Error GoTo ChkErr
    ' �HSO004.CUSTID��LINK
    strDbg = "SELECT INSTADDRESS FROM " & strOwner & "SO001" & strDBLink & _
                    " WHERE CUSTID=" & strCustID & " AND COMPCODE=" & strCompCode
    GetInstAddr = GetValue(cn, strDbg) & Empty
  Exit Function
ChkErr:
    ErrHandle "mod_Func24", "GetInstAddr"
End Function

' ���O�a�} (ChargeAddress) / �l�H�a�}(MailAddress)
Private Function GetAddress(strFaciSno As String, ByRef strChargeAddr As String, ByRef strMailAddr As String) As Boolean
  On Error GoTo ChkErr
    
    ' �HSO004.FACISNO������SO003.FACISNO
    ' �HSO003.ACCOUNTNO������SO002A.ACCOUNTNO
    
    ' ���b��: SO002A.ChargeAddress
    ' �L�b��: SO001.ChargeAddress
    
    ' ���b��: SO002A.MailAddress
    ' �L�b��: SO001.MailAddress

    Dim strAccNo As String
    Dim rs2A As Object, rs01 As Object
    
    strDbg = "SELECT ACCOUNTNO FROM " & strOwner & "SO003" & strDBLink & _
                    " WHERE CUSTID=" & strCustID & _
                    " AND COMPCODE=" & strCompCode & _
                    " AND FACISNO='" & strFaciSno & "'"
    
    strAccNo = GetValue(cn, strDbg) & Empty
    
    If strAccNo = Empty Then
        GoTo 66
    Else
        strDbg = "SELECT ACCOUNTNO,CHARGEADDRESS,MAILADDRESS FROM " & strOwner & "SO002A" & strDBLink & _
                        " WHERE CUSTID=" & strCustID & _
                        " AND COMPCODE=" & strCompCode & _
                        " AND ACCOUNTNO='" & strAccNo & "'"
        
        Set rs2A = cn.Execute(strDbg)
        
        If Not rs2A.EOF Then
            strChargeAddr = rs2A("ChargeAddress") & ""
            strMailAddr = rs2A("MailAddress") & ""
            GetAddress = True
            GoTo 99
        End If
    End If
    
66:

    strDbg = "SELECT CHARGEADDRESS,MAILADDRESS FROM " & strOwner & "SO001" & strDBLink & _
                    " WHERE CUSTID=" & strCustID & " AND COMPCODE=" & strCompCode
    
    Set rs01 = cn.Execute(strDbg)
    
    If rs01.EOF Then
        GetAddress = False
    Else
        strChargeAddr = rs01("ChargeAddress") & ""
        strMailAddr = rs01("MailAddress") & ""
        GetAddress = True
    End If

99:

    On Error Resume Next
    rs2A.Close
    rs01.Close
    Set rs2A = Nothing
    Set rs01 = Nothing
    
  Exit Function
ChkErr:
    ErrHandle "mod_Func24", "GetAddress"
End Function

' IP�� (DYNIPCOUNT)
Private Function GetDynIPcnt(strBPcode As String) As Integer
  On Error GoTo ChkErr
    ' �H SO004.BPCODE ��LINK �� CD078D.BPCODE
    '#3376 BpCode�n�h�[��޸� By Kin 2007/07/26
    strDbg = "SELECT DYNIPCOUNT FROM " & strOwner & "CD078D" & strDBLink & _
                    " WHERE BPCODE='" & strBPcode & "'"
    GetDynIPcnt = Val(GetValue(cn, strDbg) & Empty)
  Exit Function
ChkErr:
    ErrHandle "mod_Func24", "GetDynIPcnt"
End Function

' �}�q���A (OPENSTATUS)
Private Function GetOpenStatus(strCMOpenDate As String, strCMCloseDate As String) As String
  On Error GoTo ChkErr
    ' �P�_SO004�H�U�G�����:
    ' CM�}�����CMOpenDate (�}��)
    ' CM�������CMCloseDate (����)
    ' �H����̤j�����A���: �}�� �� ����
    ' �Y��̬ҵL , ��ܪ�
    strDbg = strCMOpenDate & vbTab & strCMCloseDate
    If strCMOpenDate = Empty And strCMCloseDate = Empty Then
        GetOpenStatus = ""
    Else
        If strCMOpenDate <> Empty And strCMCloseDate <> Empty Then
            GetOpenStatus = IIf(CDate(strCMOpenDate) >= CDate(strCMCloseDate), "�}��", "����")
        Else
            If strCMOpenDate <> Empty And strCMCloseDate = Empty Then
                GetOpenStatus = "�}��"
            Else
                GetOpenStatus = "����"
            End If
        End If
    End If
  Exit Function
ChkErr:
    ErrHandle "mod_Func24", "GetOpenStatus"
End Function

' �b�����A (ACCOUNTSTATUS)
Private Function GetAccStatus(strEnableAcc As String, strDisableAcc As String) As String
  On Error GoTo ChkErr
    ' �P�_SO004�H�U�G�����:
    ' �b���ҥΤ��EnableAccount (�n�})
    ' �b�����Τ��disableaccount (�n��)
    ' �H����̤j�����A��� : �n�} �� �n��
    ' �Y��̬ҵL , ��ܪť�
    strDbg = strEnableAcc & vbTab & strDisableAcc
    If strEnableAcc = Empty And strDisableAcc = Empty Then
        GetAccStatus = ""
    Else
        If strEnableAcc <> Empty And strDisableAcc <> Empty Then
            GetAccStatus = IIf(CDate(strEnableAcc) >= CDate(strDisableAcc), "�n�}", "�n��")
        Else
            If strEnableAcc <> Empty And strDisableAcc = Empty Then
                GetAccStatus = "�n�}"
            Else
                GetAccStatus = "�n��"
            End If
        End If
    End If
  Exit Function
ChkErr:
    ErrHandle "mod_Func24", "GetAccStatus"
End Function

Private Sub MakeXML(strSysID As String, strSysName As String, _
                                    strCustID As String, strDeclarantName As String, _
                                    strDialAccount As String, strFaciSno As String, _
                                    strID As String, strInstAddress As String, _
                                    strChargeAddress As String, strMailAddress As String, _
                                    strContTel As String, strContMobile As String, _
                                    strContTel2 As String, strAgentTel As String, _
                                    strOpenStatus As String, strAccountStatus As String, _
                                    strBPName As String, strDynIPCount As String)
  On Error GoTo ChkErr
'    �t�Υx�N��, �t�Υx�W��, �Ƚs, �Τ�W��, ACS�����b��, CM MAC,
'    �����Ҧr��/�νs, �˾��a�}, ���O�a�}, �l�H�a�}, �ӽФH�p���q��,
'    �ӽФH��ʹq�� , �p���H�p���q��, �N�z�H�p���q��, �}�q���A, �b�����A, �ӽФ�קO, IP��
    strXMLresult = strXMLresult & _
                            "    <DataRow SYSID=""" & strSysID & """" & _
                                " SYSNAME=""" & strSysName & """" & _
                                " CUSTID=""" & strCustID & """" & _
                                " DECLARANTNAME=""" & strDeclarantName & """" & _
                                " DIALACCOUNT=""" & strDialAccount & """" & _
                                " FACISNO=""" & strFaciSno & """" & _
                                " ID=""" & strID & """" & _
                                " INSTADDRESS=""" & strInstAddress & """" & _
                                " CHARGEADDRESS=""" & strChargeAddress & """" & _
                                " MAILADDRESS=""" & strMailAddress & """" & _
                                " CONTTEL=""" & strContTel & """" & _
                                " CONTMOBILE=""" & strContMobile & """" & _
                                " CONTTEL2=""" & strContTel2 & """" & _
                                " AGENTTEL=""" & strAgentTel & """" & _
                                " OPENSTATUS=""" & strOpenStatus & """" & _
                                " ACCOUNTSTATUS=""" & strAccountStatus & """" & _
                                " BPNAME=""" & strBPName & """" & _
                                " DYNIPCOUNT=""" & strDynIPCount & """/>" & vbCrLf
  Exit Sub
ChkErr:
    ErrHandle "mod_Func24", "MakeXML"
End Sub

