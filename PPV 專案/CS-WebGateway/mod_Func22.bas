Attribute VB_Name = "mod_Func22"
Option Explicit

Private cn As Object
Private strOwner As String

' 22�B�|���N���B�{��ID�B����(1=�d�b�ȸ�T�B2=�d�]�Ƹ�T)�B�_�l����B�I����
' �����ܮ榡 20060401

' �����אּ�h�A�ȳ]�ƻP���O��Ƭd�ߡAXML���^�� "SERVICETYPE"�A�H�ѰϤ������C
' �����p�U:
' C ��CATV&DSTB
' I ��CM
' P ��CP

' 1�D �d�ߺ���1�^�Ǥ� XML
'    (�Ƚs�B��ڽs���B�����B���O���ئW�١B��������B�������B�B�ꦬ����B�ꦬ���B�B�ꦬ����
'    �B���İ_�l����B���ĺI�����B�A�����O�B�]�ƧǸ�)

' 2�D �d�ߺ���2�^�Ǥ� XML
'    (�]�ƦW�١B�R��覡�B�w�ˤ���B�����B�]�ƫ����BCM�t�v�B�ʺAIP�ơBCM�}������BCM��������BCM�b���ҥΤ��
'    �BCM�b�����Τ���B�]�ƫO�����BCM�Ǹ��BPort�ơB��ץN�X�B��צW�١B�j�������BEBT�X���s���B�A�ȧO�B�]�ƧǸ�)

' 0: ���\
' -6 :�d�L�b�ȸ��
' -14:�d�L�w��STB�]�Ƹ��

Public Function JustDoIt(ByRef objWebPool As Object, _
                                        ByRef strPara As String) As String
  On Error GoTo ChkErr
  
    Dim strQry As String
    Dim varPara As Variant
    Dim rs As Object
    Dim objCnPool As Object
    Dim strAuthenticID As String
    Dim strStartDate As String
    Dim strStopDate As String
    Dim strCustID As String
    Dim strFaciSno As String
    Dim strBillCloseDate As String
    Set objCnPool = objWebPool
    Set cn = objCnPool.GetConnection
    
    If cn.State <= 0 Then
        Set cn = objCnPool.GetConnection
        If cn.State <= 0 Then
            strErr = "�L�k�s�u��Ʈw!"
            JustDoIt = "-99,��Ʈw�s�u����!"
            GoTo 66
        End If
    End If
    
    ' CATV&DSTB������T�d��
    ' �̻{��ID�d�ߩ��ݳ]�ƤΦ��O���
    ' 2006/04/11 ��H�{��ID�d�ߡA�ЫO�d��lAPI-22���{���X���e
    ' �ǤJ�Ѽ� [22, �|���N��, �{��ID, ����(1=�d�b�ȸ�T�B2=�d�]�Ƹ�T),�_�l���,�I����]

    varPara = Split(strPara, ",")
    '*****************************************************
    '#3018 �W�[�P�_�ǤJ���ӼƬO�_���T By Kin 2007/09/03
    If UBound(varPara) < 3 Then
        JustDoIt = "-99,�ѼƤ���!"
        GoTo 66
    End If
    '*****************************************************

    strAuthenticID = varPara(2)
    bytComp = Val(Left(strAuthenticID, 3))
    
    On Error Resume Next
    strStartDate = varPara(4)
    strStopDate = varPara(5)
    
    On Error GoTo ChkErr
    strOwner = GetOwner(cn)
    
    If Val(varPara(3)) > 2 Or Val(varPara(3)) < 1 Then
            strErr = "�ǤJ�Ѽ� [����] ���~!!"
            JustDoIt = "-99,�Ѽ� [����] ���~!"
        GoTo 66
    End If
    
    ' select Custid, FaciSNo from SO002B where MemberId=[�|���N��] and AuthenticId=[�{�ұb��] ;
    
    strQry = "SELECT CUSTID,FACISNO FROM " & strOwner & "SO002B" & Get_DB_Link(cn) & _
                    " WHERE MEMBERID='" & varPara(1) & "' AND AUTHENTICID='" & strAuthenticID & "'"

    Set rs = cn.Execute(strQry)
    
    ' If RecordSet���� = 0 Then
    '    �^��(-1,�b��{�ҿ��~)
    ' End If

    If rs.EOF Then
        JustDoIt = "-1,�b��{�ҿ��~"
    Else
        JustDoIt = Empty
        strCustID = rs!CustID & ""
        strFaciSno = rs!FaciSno & ""
'        Else If  [SO002B.FaciSNo] = null  THEN   --2006/07/12 Add
'           �^��(-23, �ӻ{��ID�d�L�{�ҳ]�ƩΪA��)
        If strFaciSno = Empty Then
            JustDoIt = "-23,�ӻ{��ID�d�L�{�ҳ]�ƩΪA��"
            GoTo 66
            Exit Function
        End If
        Select Case Val(varPara(3) & "")
                Case 1 ' CASE WHEN [����]=1 THEN
                        '#3018 ��d�� ���� ��1�ɡA�^�Ǥ� XML�A�Цb [������]] �� �W�[[�C��渹] ( MediaBillNo ) �� [ú�ںI���] ( BillCloseDate ) By Kin 2007/08/21
                        '      �^�Ǧ��O��ƪ� [�N����ú�O�I���] (BillCloseDate)�A�Y����L�ȡA�h�ǤJ [�������] (ShouldDate)�C
                        strBillCloseDate = "Decode(BillCloseDate,Null,ShouldDate,BillCloseDate) as BillCloseDate"
                        strQry = "SELECT CUSTID,BILLNO,ITEM,CITEMNAME,SHOULDDATE,SHOULDAMT,REALDATE,REALAMT" & _
                                        ",REALPERIOD,REALSTARTDATE,REALSTOPDATE,SERVICETYPE,FACISNO,UCNAME,MediaBillNo," & strBillCloseDate & _
                                        " FROM " & strOwner & "SO033" & Get_DB_Link(cn) & _
                                        " WHERE CUSTID=" & strCustID & _
                                        " AND ((FACISEQNO IN (" & strFaciSno & ")) OR AUTHENTICID='" & strAuthenticID & "')" & _
                                        " AND SHOULDDATE >= " & PrcType(strStartDate, FldDate) & _
                                        " AND SHOULDDATE < " & PrcType(strStopDate, FldDate) & "+1" & _
                                        " AND CANCELFLAG <> 1 UNION ALL " & _
                                         "SELECT CUSTID,BILLNO,ITEM,CITEMNAME,SHOULDDATE,SHOULDAMT,REALDATE,REALAMT" & _
                                         ",REALPERIOD,REALSTARTDATE,REALSTOPDATE,SERVICETYPE,FACISNO,UCNAME,MediaBillNo," & strBillCloseDate & _
                                        " FROM " & strOwner & "SO034" & Get_DB_Link(cn) & _
                                        " WHERE CUSTID=" & strCustID & _
                                        " AND ((FACISEQNO IN (" & strFaciSno & ")) OR AUTHENTICID='" & strAuthenticID & "')" & _
                                        " AND SHOULDDATE >= " & PrcType(strStartDate, FldDate) & _
                                        " AND SHOULDDATE < " & PrcType(strStopDate, FldDate) & "+1"
                                        
                        strQry = "SELECT * FROM (" & strQry & ") ORDER BY BILLNO DESC,ITEM"
                            
                        Set rs = cn.Execute(strQry)
                        
                        If rs.EOF Then ' If RecordSet���� > 0 Then
                            JustDoIt = "-6, �d�L�b�ȸ��" ' �^��(-6, �d�L�b�ȸ��)
                        Else ' Else
                            ' ����XML��.(��select�X�Ӥ����)
                            JustDoIt = "0,���\" & vbCrLf & ExpXML(rs) ' �^��(0, ���\)��XML��
                        End If ' End If
'                       PS:
'                       XML(�Ƚs�B��ڽs���B�����B���O���ئW�١B��������B�������B�B�ꦬ����B�ꦬ���B�B
'                               �ꦬ���ơB���İ_�l����B���ĺI�����B�A�����O�B�]�ƧǸ�)

                Case 2 ' CASE WHEN [����]=2 THEN
                
'                        Select FaciName, BuyName, OldInstDate,PRDate,ModelName,CMBaudRate,DynIPCount,
'                            CMOpenDate,CMCloseDate,EnableAccount, DisableAccount,Deposit,ReFaciSNo,PortNo,
'                            BPCode,BPName,ContractDate,EBTContNo,ServiceType, FaciSNO from SO004
'                            where Custid=[SO002B.Custid] and SeqNo in [SO002B.FaciSNo]
'                            Union All
'                        Select FaciName, BuyName, OldInstDate,PRDate,ModelName,CMBaudRate,DynIPCount,
'                            CMOpenDate,CMCloseDate,EnableAccount, DisableAccount,Deposit,ReFaciSNo,PortNo,
'                            BPCode,BPName,ContractDate,EBTContNo,ServiceType, FaciSNO from SO004
'                            where ServiceType='P' and Custid=[SO002B.Custid] and ReFaciSNo in
'                            (Select FaciSNo from SO004 where Custid=[SO002B.Custid] and SeqNo=[SO002B.FaciSNo] ) ;

                        strQry = "SELECT FACINAME, BUYNAME, OLDINSTDATE,PRDATE,MODELNAME,CMBAUDRATE,DYNIPCOUNT" & _
                                        ",CMOPENDATE,CMCLOSEDATE,ENABLEACCOUNT, DISABLEACCOUNT,DEPOSIT,REFACISNO,PORTNO" & _
                                        ",BPCODE,BPNAME,CONTRACTDATE,EBTCONTNO,SERVICETYPE, FACISNO" & _
                                        ",SMARTCARDNO,DECLARANTNAME,DESCRIPTION INITPLACE,SEQNO" & _
                                        " FROM " & strOwner & "SO004" & Get_DB_Link(cn) & _
                                        "," & strOwner & "CD056" & Get_DB_Link(cn) & _
                                        " WHERE CUSTID=" & strCustID & " AND SEQNO IN (" & strFaciSno & ")" & _
                                        " AND SO004.INITPLACENO=CD056.CODENO(+)"
                                        
                        strQry = strQry & " UNION ALL " & _
                                        "SELECT FACINAME, BUYNAME, OLDINSTDATE,PRDATE,MODELNAME,CMBAUDRATE,DYNIPCOUNT" & _
                                        ",CMOPENDATE,CMCLOSEDATE,ENABLEACCOUNT, DISABLEACCOUNT,DEPOSIT,REFACISNO,PORTNO" & _
                                        ",BPCODE,BPNAME,CONTRACTDATE,EBTCONTNO,SERVICETYPE, FACISNO" & _
                                        ",SMARTCARDNO,DECLARANTNAME,DESCRIPTION INITPLACE,SEQNO" & _
                                        " FROM " & strOwner & "SO004" & Get_DB_Link(cn) & _
                                        "," & strOwner & "CD056" & Get_DB_Link(cn) & _
                                        " WHERE SERVICETYPE='P' AND CUSTID=" & strCustID & _
                                        " AND REFACISNO IN (SELECT FACISNO FROM " & strOwner & "SO004" & Get_DB_Link(cn) & _
                                        " WHERE CUSTID=" & strCustID & " AND SEQNO IN (" & strFaciSno & "))" & _
                                        " AND SO004.INITPLACENO=CD056.CODENO(+)"
                                        
                        strQry = strQry & " UNION ALL " & _
                                        "SELECT 'CATV�A��', NULL BUYNAME, B.INSTTIME OLDINSTDATE,B.PRTIME PRDATE" & _
                                        ",NULL MODELNAME,NULL CMBAUDRATE,NULL DYNIPCOUNT,NULL CMOPENDATE" & _
                                        ",NULL CMCLOSEDATE,NULL ENABLEACCOUNT,NULL DISABLEACCOUNT,NULL DEPOSIT" & _
                                        ",NULL REFACISNO,NULL PORTNO,NULL BPCODE,NULL BPNAME,NULL CONTRACTDATE" & _
                                        ",NULL EBTCONTNO,B.SERVICETYPE,TO_CHAR(A.CUSTID) FACISNO,NULL SMARTCARDNO" & _
                                        ",A.CUSTNAME DECLARANTNAME,NULL INITPLACE,TO_CHAR(A.CUSTID) SEQNO" & _
                                        " FROM " & strOwner & "SO001" & Get_DB_Link(cn) & " A" & _
                                        "," & strOwner & "SO002" & Get_DB_Link(cn) & " B" & _
                                        " WHERE A.CUSTID IN (" & strFaciSno & ")" & _
                                        " AND A.CUSTID=B.CUSTID AND B.SERVICETYPE='C'"
                                        
'                        Union All
'                        Select 'CATV�A��', null BuyName, InstTime OldInstDate,PRTime PRDate,null ModelName,
'                        null CMBaudRate,null DynIPCount,null CMOpenDate,null CMCloseDate,null EnableAccount,
'                        null DisableAccount,null Deposit,null ReFaciSNo,null PortNo,null BPCode,null BPName,
'                        null ContractDate,null EBTContNo,b.ServiceType, to_char(a.custid) FaciSNO, null SmartCard,
'                        a.CustName DeclarantName, null InitPlace
'                        from SO001 a, SO002 b
'                        where a.Custid in [SO002B.FaciSNo] and a.custid=b.custid and b.servicetype='C' ;
                                        
                        Set rs = cn.Execute(strQry)
                        
                        If rs.EOF Then ' If RecordSet���� > 0 Then
                            JustDoIt = "-14, �d�L�w��STB�]�Ƹ��" ' �^��(-14, �d�L�w��STB�]�Ƹ��)
                        Else ' Else
                            ' ����XML��.(��select�X�Ӥ����)
                            JustDoIt = "0,���\" & vbCrLf & ExpXML(rs) ' �^��(0, ���\)��XML��
                        End If ' End If
'                       PS:
'                       XML(�]�ƦW�١B�R��覡�B�w�ˤ���B�����B�]�ƫ����BCM�t�v�B�ʺAIP�ơBCM�}������BCM��������B
'                               CM�b���ҥΤ���BCM�b�����Τ���B�]�ƫO�����BCM�Ǹ��BPort�ơB��ץN�X�B��צW�١B�j�������B
'                               EBT�X���s���B�A�����O�B�]�ƧǸ��B�]�Ƭy����)
        End Select
    End If

66:
    On Error Resume Next
    Rlx varPara
    Rlx strCustID
    Rlx strFaciSno
    Rlx objCnPool
    Rlx strStartDate
    Rlx strStopDate
    Rlx strAuthenticID
    Set cn = Nothing
    Set objCnPool = Nothing
  Exit Function

ChkErr:
    ErrHandle "mod_Func22", "JustDoIt"
    JustDoIt = "-99," & strErr
End Function

'Private cn As Object
'Private strOwner As String
'
'Public Function JustDoIt(ByRef objWebPool As Object, _
'                                        ByRef strPara As String) As String
'  On Error GoTo ChkErr
'
'    Dim strQry As String
'    Dim varPara As Variant
'    Dim rs As Object
'    Dim objCnPool As Object
'
'    Set objCnPool = objWebPool
'    Set cn = objCnPool.GetConnection
'
'    If cn.state <= 0 Then
'        Set cn = objCnPool.GetConnection
'        If cn.state <= 0 Then
'            strErr = "�L�k�s�u��Ʈw!"
'            JustDoIt = "-99,��Ʈw�s�u����!"
'            GoTo 66
'        End If
'    End If
'
'    ' CATV&DSTB������T�d��
'    ' �ǤJ�Ѽ� [22, ���q�O,�ӽФH, �������r��, �q��, ����(1=�d�b�ȸ�T�B2=�d�]�Ƹ�T),�_�l���,�I����]
'
'    varPara = Split(strPara, ",")
'
'    bytComp = varPara(1)
'    strOwner = GetOwner(cn)
'
'    If Val(varPara(5)) > 2 Or Val(varPara(5)) < 1 Then
'            strErr = "�ǤJ�Ѽ� [����] ���~!!"
'            JustDoIt = "-99,�Ѽ� [����] ���~!"
'        GoTo 66
'    End If
'
'    ' ---- �� �� ----
'    '   Select distinct Custid, Facisno from (
'    '       Select Custid, null FaciSNO from SO001 where Compcode=[���q�O] and CustName=[�ӽФH]
'    '           and ID=[�������r��] and Tel1=[�q��]
'    '       Union All
'    '       Select Custid, Facisno from SO004 where Compcode=[���q�O] and DeclarantName=[�ӽФH]
'    '           and ID=[�������r��] and Contel=[�q��]
'    '           and facicode in (select codeno from cd022 where refno>0) and ServiceType='C' );
'    '
'    '   If RecordSet���� = 0 Then
'    '       �^��(-20, �d�L�ӽФH���)
'    '   End If
'
'    ' ---- �s �� ----
'    '    Select distinct Custid, Facisno from (
'    '       Select A.Custid, null FaciSNO from SO001 A, SO002 B where A.Compcode=[���q�O]
'    '           and A.CUSTID=B.CUSTID AND B.SERVICETYPE='C' AND A.CustName=[�ӽФH]
'    '           and B.ID=[�������r��] and A.Tel1=[�q��]
'    '       Union All
'    '       Select Custid, Facisno from SO004 where Compcode=[���q�O] and DeclarantName=[�ӽФH]
'    '           and ID=[�������r��] and ConTtel=[�q��]
'    '           and facicode in (select codeno from cd022 where refno>0) and ServiceType='C' );
'
'
'    strQry = "SELECT DISTINCT CUSTID,FACISNO FROM (" & _
'                    " SELECT A.CUSTID,NULL FACISNO" & _
'                    " FROM " & strOwner & "SO001" & Get_DB_Link(cn) & " A," & strOwner & "SO002" & Get_DB_Link(cn) & " B" & _
'                    " WHERE A.COMPCODE=" & varPara(1) & _
'                    " AND A.CUSTID=B.CUSTID" & _
'                    " AND B.SERVICETYPE='C'" & _
'                    " AND A.CUSTNAME='" & varPara(2) & "'" & _
'                    " AND B.ID='" & varPara(3) & "'" & _
'                    " AND A.TEL1='" & varPara(4) & "'" & _
'                    " UNION ALL " & _
'                    " SELECT CUSTID,FACISNO FROM " & strOwner & "SO004" & Get_DB_Link(cn) & _
'                    " WHERE COMPCODE=" & varPara(1) & _
'                    " AND DECLARANTNAME='" & varPara(2) & "'" & _
'                    " AND ID='" & varPara(3) & "'" & _
'                    " AND CONTTEL='" & varPara(4) & "'" & _
'                    " AND FACICODE IN " & _
'                    " (SELECT CODENO FROM " & strOwner & "CD022" & Get_DB_Link(cn) & " WHERE REFNO > 0)" & _
'                    " AND SERVICETYPE='C')"
'
'' �ª�
''    strQry = "SELECT DISTINCT CUSTID,FACISNO FROM (" & _
'                    " SELECT CUSTID,NULL FACISNO FROM " & strOwner & "SO001" & Get_DB_Link(cn) & _
'                    " WHERE COMPCODE=" & varPara(1) & _
'                    " AND CUSTNAME='" & varPara(2) & "'" & _
'                    " AND ID='" & varPara(3) & "'" & _
'                    " AND TEL1='" & varPara(4) & "'" & _
'                    " UNION ALL " & _
'                    " SELECT CUSTID,FACISNO FROM " & strOwner & "SO004" & Get_DB_Link(cn) & _
'                    " WHERE COMPCODE=" & varPara(1) & _
'                    " AND DECLARANTNAME='" & varPara(2) & "'" & _
'                    " AND ID='" & varPara(3) & "'" & _
'                    " AND CONTTEL='" & varPara(4) & "'" & _
'                    " AND FACICODE IN " & _
'                    " (SELECT CODENO FROM " & strOwner & "CD022" & Get_DB_Link(cn) & " WHERE REFNO > 0)" & _
'                    " AND SERVICETYPE='C')"
'
'    Set rs = cn.Execute(strQry)
'
'    '   If RecordSet���� = 0 Then
'    '       �^��(-20, �d�L�ӽФH���)
'    '   End If
'
'    If rs.RecordCount <= 0 Then
'        JustDoIt = "-20, �d�L�ӽФH���"
'    Else
'        JustDoIt = Empty
'        With rs ' LOOP �B�z (�̥ӽФH�����X���ŦX��CustId, FaciSno)
''           [�O��select�X�Ӫ����]
'            If Val(varPara(5)) = 1 Then
'                While Not .EOF
'                    JustDoIt = JustDoIt & ExpXML(cn.Execute(GetTypeSQL(Val(varPara(5) & ""), !CustID & "", !FaciSno & "", _
'                                                                        varPara(6) & "", varPara(7) & "")), True)
'                    .MoveNext
'                Wend
'            Else
'                While Not .EOF
'                    JustDoIt = JustDoIt & ExpXML(cn.Execute(GetTypeSQL(Val(varPara(5) & ""), !CustID & "", !FaciSno & "")), True)
'                    .MoveNext
'                Wend
'            End If
'        End With ' END LOOP
'
'        If JustDoIt = Empty Then
'            If Val(varPara(5)) = 1 Then '   CASE WHEN [����]=1 THEN
'                JustDoIt = "-6, �d�L�b�ȸ��" ' �^��(-6, �d�L�b�ȸ��)
'            Else '   CASE WHEN [����]=2 THEN
'                JustDoIt = "-14, �d�L�w��STB�]�Ƹ��" ' �^��(-14, �d�L�w��STB�]�Ƹ��)
'            End If
'        Else '   If RecordSet���� > 0 Then
'            ' ����XML��.(��select�X�Ӥ����)
'            JustDoIt = "0,���\" & vbCrLf & _
'                                "<?xml version=""1.0"" encoding=""BIG5"" standalone=""yes"" ?>" & vbCrLf & _
'                                "<DataSet>" & vbCrLf & _
'                                "  <DataTable>" & vbCrLf & _
'                                JustDoIt & _
'                                "  </DataTable>" & vbCrLf & _
'                                "</DataSet>" '       �^��(0, ���\)��XML��
'        End If
'    End If
'
'66:
'    On Error Resume Next
'    Rlx varPara
'    Rlx objCnPool
'    Set cn = Nothing
'    Set objCnPool = Nothing
'  Exit Function
'
'ChkErr:
'    ErrHandle "mod_Func22", "JustDoIt"
'    JustDoIt = "-99," & strErr
'End Function
'
'Private Function GetTypeSQL(bytType As Byte, _
'                                                strCustID As String, _
'                                                ByVal strFaciSNo As String, _
'                                                Optional strStartDate As String = "", _
'                                                Optional strStopDate As String = "") As String
'  On Error GoTo ChkErr
'
'    If bytType = 1 Then
'
'        'PS:
'        'XML(�Ƚs�B��ڽs���B�����B���O���ئW�١B��������B�������B�B�ꦬ����B�ꦬ���B�B�ꦬ���ơB
'        '       ���İ_�l����B���ĺI�����B������]�B�u����]�B�A�����O�B�]�ƧǸ�)
'
'        strFaciSNo = IIf(strFaciSNo = Empty, "FACISNO IS NULL", "FACISNO='" & strFaciSNo & "'")
'
'        GetTypeSQL = "SELECT CUSTID,BILLNO,ITEM,CITEMNAME,SHOULDDATE,SHOULDAMT,REALDATE,REALAMT" & _
'                                ",REALPERIOD,REALSTARTDATE,REALSTOPDATE,UCNAME,STNAME,SERVICETYPE,FACISNO" & _
'                                " FROM " & strOwner & "SO033" & Get_DB_Link(cn) & _
'                                " WHERE CUSTID=" & strCustID & " AND " & strFaciSNo & _
'                                " AND SHOULDDATE >= " & PrcType(strStartDate, FldDate) & _
'                                " AND SHOULDDATE < " & PrcType(strStopDate, FldDate) & "+1" & _
'                                " AND SERVICETYPE = 'C'" & _
'                                " UNION ALL " & _
'                                "SELECT CUSTID,BILLNO,ITEM,CITEMNAME,SHOULDDATE,SHOULDAMT,REALDATE,REALAMT" & _
'                                ",REALPERIOD,REALSTARTDATE,REALSTOPDATE,UCNAME,STNAME,SERVICETYPE,FACISNO" & _
'                                " FROM " & strOwner & "SO034" & Get_DB_Link(cn) & _
'                                " WHERE CUSTID=" & strCustID & " AND " & strFaciSNo & _
'                                " AND SHOULDDATE >= " & PrcType(strStartDate, FldDate) & _
'                                " AND SHOULDDATE < " & PrcType(strStopDate, FldDate) & "+1" & _
'                                " AND SERVICETYPE = 'C'"
'
''        If [FaciSno] Is Null Then
''            Select Custid, BillNo, Item, CitemName, ShouldDate, ShouldAmt, RealDate, RealAmt, RealPeriod,
''                RealStartDate, RealStopDate, UCName, STName, Servicetype, FaciSno from SO033
''                Where Custid = [Custid] And FaciSno Is Null And ShouldDate >= [�_�l���] And ShouldDate < [�I����] + 1
''                and ServiceType='C'
''                Union All
''            Select Custid, BillNo, Item, CitemName, ShouldDate, ShouldAmt, RealDate, RealAmt, RealPeriod,
''                RealStartDate, RealStopDate, UCName, STName, Servicetype FaciSno from SO034
''                Where Custid = [Custid] And FaciSno Is Null And ShouldDate >= [�_�l���] And ShouldDate < [�I����] + 1
''                and ServiceType='C' ;
''        Else
''            Select Custid, BillNo, Item, CitemName, ShouldDate, ShouldAmt, RealDate, RealAmt, RealPeriod,
''                RealStartDate, RealStopDate, UCName, STName, Servicetype, FaciSno from SO033
''                Where Custid = [Custid] And FaciSno = [FaciSno] And ShouldDate >= [�_�l���] And ShouldDate < [�I����] + 1
''                and ServiceType='C'
''                Union All
''            Select Custid, BillNo, Item, CitemName, ShouldDate, ShouldAmt, RealDate, RealAmt, RealPeriod,
''                RealStartDate, RealStopDate, UCName, STName, Servicetype, FaciSno from SO034
''                Where Custid = [Custid] And FaciSno = [FaciSno] And ShouldDate >= [�_�l���] And ShouldDate < [�I����] + 1
''                and ServiceType='C' ;
''        End If
'
'    Else
'
'        'PS:
'        'XML(�Ƚs�B�]�ƦW�١B�]�ƧǸ��B�R��覡�B�w�ˤ���B�����B�]�ƫ����B�]�ƫO�����B���z�d�Ǹ��B���O���ءB���İ_�l��B���ĺI���)
'
'        GetTypeSQL = "SELECT A.CUSTID,A.FACINAME,A.FACISNO,A.BUYNAME,A.OLDINSTDATE,A.PRDATE" & _
'                                ",A.MODELNAME,A.DEPOSIT,A.SMARTCARDNO,B.CITEMNAME,B.STARTDATE,B.STOPDATE" & _
'                                " FROM " & strOwner & "SO004" & Get_DB_Link(cn) & " A," & strOwner & "SO003" & Get_DB_Link(cn) & " B" & _
'                                " WHERE A.CUSTID=" & strCustID & " AND A.FACISNO='" & strFaciSNo & "'" & _
'                                " AND A.CUSTID=B.CUSTID(+) AND A.FACISNO=B.FACISNO(+)"
'
''        If [FaciSno] Is Not Null Then
''            Select a.Custid,a.FaciName,a.FaciSNO,a.BuyName,a.OldInstDate,a.PRDate,a.ModelName,a.Deposit,
''                   a.SmartCardNo,b.citemname,b.startdate,b.stopdate from SO004 a, so003 b
''                   where a.Custid=[Custid] and a.FaciSno=[FaciSno] and a.custid=b.custid(+) and a.facisno=b.facisno(+) ;
''        End If
'    End If
'
'  Exit Function
'ChkErr:
'    ErrHandle "mod_Func22", "GetTypeSQL"
'    GetTypeSQL = "-99," & strErr
'End Function
'
