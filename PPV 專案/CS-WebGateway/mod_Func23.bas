Attribute VB_Name = "mod_Func23"
Option Explicit
    
Private cn As Object
Private strOwner As String

Public Function JustDoIt(ByRef objWebPool As Object, _
                                        ByRef strPara As String) As String
  On Error GoTo ChkErr
  
    Dim strQry As String
    Dim varPara As Variant
    Dim rs As Object
    Dim objCnPool As Object
    
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
    
    ' CMCP������T�d��
    ' �ǤJ�Ѽ� [23, ���q�O, �ӽФH, �������r��, �ͤ�, ����(1=�d�b�ȸ�T�B2=�d�]�Ƹ�T), �_�l���, �I����]
    
    varPara = Split(strPara, ",")

    bytComp = varPara(1)
    strOwner = GetOwner(cn)
    
    If Val(varPara(5)) > 2 Or Val(varPara(5)) < 1 Then
            strErr = "�ǤJ�Ѽ� [����] ���~!!"
            JustDoIt = "-99,�Ѽ� [����] ���~!"
        GoTo 66
    End If
    
    '   Select CUSTID,FACISNO From SO004
    '       Where ServiceType='I' And DeclarantName=[�ӽФH]
    '       And ID=[�������r��] And Birthday=[�ͤ�];
    
    strQry = "SELECT CUSTID,FACISNO FROM " & strOwner & "SO004" & Get_DB_Link(cn) & _
                    " WHERE SERVICETYPE='I'" & _
                    " AND DECLARANTNAME='" & varPara(2) & "'" & _
                    " AND ID='" & varPara(3) & "'" & _
                    " AND BIRTHDAY=" & PrcType(varPara(4), FldDate) & _
                    " AND COMPCODE=" & varPara(1)
    
    Set rs = cn.Execute(strQry)
    
    '   If RecordSet���� = 0 Then
    '       �^��(-20, �d�L�ӽФH���)
    '   End If
    
    If rs.RecordCount <= 0 Then
        JustDoIt = "-20, �d�L�ӽФH���"
    Else
        JustDoIt = Empty
        With rs ' LOOP �B�z (�̥ӽФH�����X���ŦX��CustId, FaciSno)
            If Val(varPara(5)) = 1 Then
                While Not .EOF
                    JustDoIt = JustDoIt & ExpXML(cn.Execute(GetTypeSQL(Val(varPara(5) & ""), !CustID & "", !FaciSno & "", _
                                                                        varPara(6) & "", varPara(7) & "")), True)
                    .MoveNext
                Wend
            Else
                While Not .EOF
                    JustDoIt = JustDoIt & ExpXML(cn.Execute(GetTypeSQL(Val(varPara(5) & ""), !CustID & "", !FaciSno) & ""), True)
                    .MoveNext
                Wend
            End If
        End With ' END LOOP
        
        If JustDoIt = Empty Then
            If Val(varPara(5)) = 1 Then '   CASE WHEN [����]=1 THEN
                JustDoIt = "-6, �d�L�b�ȸ��" ' �^��(-6, �d�L�b�ȸ��)
            Else '   CASE WHEN [����]=2 THEN
                JustDoIt = "-21, �d�L�w��CMCP�]�Ƹ��" ' �^��(-21, �d�L�w��CMCP�]�Ƹ��)
            End If
        Else
            JustDoIt = "0,���\" & vbCrLf & _
                                "<?xml version=""1.0"" encoding=""BIG5"" standalone=""yes"" ?>" & vbCrLf & _
                                "<DataSet>" & vbCrLf & _
                                "  <DataTable>" & vbCrLf & _
                                JustDoIt & _
                                "  </DataTable>" & vbCrLf & _
                                "</DataSet>" '       �^��(0, ���\)��XML��
        End If
    End If

66:
    On Error Resume Next
    Rlx varPara
    Rlx objCnPool
    Set cn = Nothing
    Set objCnPool = Nothing
  Exit Function

ChkErr:
    ErrHandle "mod_Func23", "JustDoIt"
    JustDoIt = "-99," & strErr
End Function

Private Function GetTypeSQL(bytType As Byte, _
                                                strCustID As String, _
                                                strFaciSno As String, _
                                                Optional strStartDate As String = "", _
                                                Optional strStopDate As String = "") As String
  On Error GoTo ChkErr
  
    If bytType = 1 Then
        '   PS:
        '   XML(��ڽs���B�����B���O���ئW�١B��������B�������B�B�ꦬ����B�ꦬ���B�B�ꦬ����
        '           �B���İ_�l����B���ĺI�����B������]�B�u����]�B�@�o�_�B�A�����O)
        GetTypeSQL = "SELECT BILLNO,ITEM,CITEMNAME,SHOULDDATE,SHOULDAMT" & _
                                ",REALDATE,REALAMT,REALPERIOD,REALSTARTDATE,REALSTOPDATE" & _
                                ",UCNAME,STNAME,DECODE(CANCELFLAG,1,'�O','�_') CANCELFLAG" & _
                                ",SERVICETYPE,FACISNO" & _
                                " FROM " & strOwner & "SO033" & Get_DB_Link(cn) & _
                                " WHERE CUSTID=" & strCustID & _
                                " AND FACISNO='" & strFaciSno & "'" & _
                                " AND SERVICETYPE IN ('I','P')" & _
                                " AND SHOULDDATE >= " & PrcType(strStartDate, FldDate) & _
                                " AND SHOULDDATE < " & PrcType(strStopDate, FldDate) & "+1" & _
                                " UNION ALL " & _
                                "SELECT BILLNO,ITEM,CITEMNAME,SHOULDDATE,SHOULDAMT" & _
                                ",REALDATE,REALAMT,REALPERIOD,REALSTARTDATE,REALSTOPDATE" & _
                                ",UCNAME,STNAME,DECODE(CANCELFLAG,1,'�O','�_') CANCELFLAG" & _
                                ",SERVICETYPE,FACISNO" & _
                                " FROM " & strOwner & "SO034" & Get_DB_Link(cn) & _
                                " WHERE CUSTID=" & strCustID & _
                                " AND FACISNO='" & strFaciSno & "'" & _
                                " AND SERVICETYPE IN ('I','P')" & _
                                " AND SHOULDDATE >= " & PrcType(strStartDate, FldDate) & _
                                " AND SHOULDDATE < " & PrcType(strStopDate, FldDate) & "+1"
                                
        '      Select BillNo, Item, CitemName, ShouldDate, ShouldAmt, RealDate, RealAmt, RealPeriod,
        '           RealStartDate, RealStopDate, UCName, STName, CancelFlag, ServiceType From SO033
        '           Where CustId=[CustId] And FaciSno=[FaciSno] And ServiceType In ('I','P')
        '           And ShouldDate>=[�_�l���] And ShouldDate < [�I����]+1
        '      Union All
        '          Select BillNo, Item, CitemName, ShouldDate, ShouldAmt, RealDate, RealAmt, RealPeriod,
        '           RealStartDate, RealStopDate, UCName, STName, CancelFlag, ServiceType From SO034
        '           Where CustId=[CustId] And FaciSno=[FaciSno] And ServiceType In ('I','P')
        '           And ShouldDate>=[�_�l���] And ShouldDate < [�I����]+1
    Else
        '   PS:
        '   XML(�]�ƦW�١B�R��覡�B�w�ˤ���B�����B�]�ƫ����BCM�t�v�B�ʺAIP�ơBCM�}������BCM�������
        '           �BCM�b���ҥΤ���BCM�b�����Τ���B�]�ƫO�����BCM�Ǹ��BPort�ơB��ץN�X�B��צW��
        '           �B�j�������BEBT�X���s���B�A�ȧO)
        
        GetTypeSQL = "SELECT FACINAME,BUYNAME,OLDINSTDATE,PRDATE,MODELNAME" & _
                                ",CMBAUDRATE,DYNIPCOUNT,CMOPENDATE,CMCLOSEDATE" & _
                                ",ENABLEACCOUNT,DISABLEACCOUNT,DEPOSIT,REFACISNO,PORTNO" & _
                                ",BPCODE,BPNAME,CONTRACTDATE,EBTCONTNO,SERVICETYPE,FACISNO" & _
                                " FROM " & strOwner & "SO004" & Get_DB_Link(cn) & _
                                " WHERE CUSTID=" & strCustID & _
                                " AND FACISNO='" & strFaciSno & "'" & _
                                " AND SERVICETYPE='I' " & _
                                " UNION ALL " & _
                                "SELECT FACINAME,BUYNAME,OLDINSTDATE,PRDATE,MODELNAME" & _
                                ",CMBAUDRATE,DYNIPCOUNT,CMOPENDATE,CMCLOSEDATE" & _
                                ",ENABLEACCOUNT,DISABLEACCOUNT,DEPOSIT,REFACISNO,PORTNO" & _
                                ",BPCODE,BPNAME,CONTRACTDATE,EBTCONTNO,SERVICETYPE,FACISNO" & _
                                " FROM " & strOwner & "SO004" & Get_DB_Link(cn) & _
                                " WHERE CUSTID=" & strCustID & _
                                " AND REFACISNO='" & strFaciSno & "'" & _
                                " AND SERVICETYPE='P' "

            '    Select FaciName, BuyName, OldInstDate,PRDate,ModelName,CMBaudRate,DynIPCount,CMOpenDate,
            '       CMCloseDate,EnableAccount, DisableAccount,Deposit,ReFaciSNo,PortNo,BPCode,BPName,
            '       ContractDate,EBTContNo,ServiceType, FaciSNO from SO004
            '       where ServiceType='I' and Custid=[Custid] and FaciSno=[FaciSNo]
            '    Union All
            '       Select FaciName, BuyName, OldInstDate,PRDate,ModelName,CMBaudRate,DynIPCount,CMOpenDate,
            '       CMCloseDate,EnableAccount, DisableAccount,Deposit,ReFaciSNo,PortNo,BPCode,BPName,
            '       ContractDate,EBTContNo,ServiceType, FaciSNO from SO004
            '       where ServiceType='P' and Custid=[Custid] and ReFaciSno=[FaciSNo]
                                    
    End If
      
  Exit Function
ChkErr:
    ErrHandle "mod_Func23", "GetTypeSQL"
    GetTypeSQL = "-99," & strErr
End Function

'        GetTypeSQL = "SELECT A.FACINAME,A.BUYNAME,A.OLDINSTDATE,A.PRDATE,A.MODELNAME" & _
'                                ",B.CMBAUDRATE,B.DYNIPCOUNT,A.CMOPENDATE,A.CMCLOSEDATE" & _
'                                ",A.ENABLEACCOUNT,A.DISABLEACCOUNT,A.DEPOSIT,A.REFACISNO" & _
'                                ",A.PORTNO,C.BPCODE,C.BPNAME,A.CONTRACTDATE,A.EBTCONTNO" & _
'                                ",A.SERVICETYPE,A.FACISNO" & _
'                                " FROM " & strOwner & "SO004" & Get_DB_Link(cn) & " A " & _
'                                " ," & strOwner & "CD078D" & Get_DB_Link(cn) & " B " & _
'                                " ," & strOwner & "SO003" & Get_DB_Link(cn) & " C " & _
'                                " WHERE A.CUSTID=" & strCustID & _
'                                " AND A.FACISNO='" & strFaciSno & "'" & _
'                                " AND A.SERVICETYPE='I' " & _
'                                " AND A.SERVICETYPE=C.SERVICETYPE" & _
'                                " AND A.FACISNO=C.FACISNO" & _
'                                " AND C.BPCODE=B.BPCODE" & _
'                                " UNION ALL " & _
'                                "SELECT FACINAME,BUYNAME,OLDINSTDATE,PRDATE,MODELNAME" & _
'                                ",CMBAUDRATE,DYNIPCOUNT,CMOPENDATE,CMCLOSEDATE" & _
'                                ",ENABLEACCOUNT,DISABLEACCOUNT,DEPOSIT,REFACISNO" & _
'                                ",PORTNO,BPCODE,BPNAME,CONTRACTDATE,EBTCONTNO" & _
'                                ",SERVICETYPE,FACISNO" & _
'                                " FROM " & strOwner & "SO004" & Get_DB_Link(cn) & _
'                                " WHERE CUSTID=" & strCustID & _
'                                " AND REFACISNO='" & strFaciSno & "'" & _
'                                " AND SERVICETYPE='P' "
                                    
    '       Select a.FaciName, a.BuyName, OldInstDate,PRDate,a.ModelName,b.CMBaudRate,
    '       b.DynIPCount,CMOpenDate,CMCloseDate,EnableAccount, DisableAccount,Deposit,
    '       ReFaciSNo,PortNo,c.BPCode,c.BPName,ContractDate,EBTContNo,a.ServiceType
    '       From SO004 a, CD078D b, SO003 c
    '       Where a.ServiceType='I' And a.Custid=[Custid] And a.FaciSno=[FaciSNo] And
    '       a.ServiceType=c.ServiceType And a.FaciSno=c.FaciSNo And c.bpcode=b.bpcode
    '       Union All
    '       Select FaciName, BuyName, OldInstDate,PRDate,ModelName,CMBaudRate,
    '       DynIPCount,CMOpenDate,CMCloseDate,EnableAccount, DisableAccount,Deposit,
    '       ReFaciSNo,PortNo,BPCode,BPName,ContractDate,EBTContNo,ServiceType
    '       From SO004
    '       Where ServiceType='P' And Custid=[Custid] And ReFaciSno=[FaciSNo]


'        If varPara(5) = 1 Then '   CASE WHEN [����]=1 THEN
'
'            JustDoIt = Empty
'
'            With rs
'                While Not .EOF ' LOOP �B�z (�̥ӽФH�����X���ŦX��CustId, FaciSno)
'                    JustDoIt = JustDoIt & _
'                                        ExpXML(cn.Execute(GetType1SQL(!CustID, !FaciSno, CStr(varPara(6)), CStr(varPara(7)))), True)
'                    .MoveNext
'                Wend ' END LOOP
'            End With
'
'            ' [�O��select�X�Ӫ����]
'
'            '   If RecordSet���� > 0 Then
'            '       ����XML��.(��select�X�Ӥ����)
'            '       �^��(0, ���\)��XML��
'            '   Else
'            '       �^��(-6, �d�L�b�ȸ��)
'            '   End If
'
'            If JustDoIt = Empty Then
'                JustDoIt = "-6, �d�L�b�ȸ��"
'            Else
'                JustDoIt = "0,���\" & vbCrLf & _
'                                    "<?xml version=""1.0"" encoding=""BIG5"" standalone=""yes"" ?>" & vbCrLf & _
'                                    "<DataSet>" & vbCrLf & _
'                                    "  <DataTable>" & vbCrLf & _
'                                    JustDoIt & _
'                                    "  </DataTable>" & vbCrLf & _
'                                    "</DataSet>"
'            End If
'            '   PS:
'            '   XML(��ڽs���B�����B���O���ئW�١B��������B�������B�B�ꦬ����B�ꦬ���B�B�ꦬ����
'            '           �B���İ_�l����B���ĺI�����B������]�B�u����]�B�@�o�_�B�A�����O)
'
'        ElseIf varPara(5) = 2 Then '   CASE WHEN [����]=2 THEN
'
'            JustDoIt = Empty
'            With rs
'                While Not .EOF ' LOOP �B�z (�̥ӽФH�����X���ŦX��Custid, Facisno)
'                    JustDoIt = JustDoIt & _
'                                        ExpXML(cn.Execute(GetType2SQL(!CustID, !FaciSno, "", "")), True)
'                    .MoveNext
'                Wend ' END LOOP
'            End With
'
'            ' [�O��select�X�Ӫ����]
'
'            '   If RecordSet���� > 0 Then
'            '       ����XML��.(��select�X�Ӥ����)
'            '       �^��(0, ���\)��XML��
'            '   Else
'            '       �^��(-21, �d�L�w��CMCP�]�Ƹ��)
'            '   End If
'
'            If JustDoIt = Empty Then
'                JustDoIt = "-21, �d�L�w��CMCP�]�Ƹ��"
'            Else
'                JustDoIt = "0,���\" & vbCrLf & _
'                                    "<?xml version=""1.0"" encoding=""BIG5"" standalone=""yes"" ?>" & vbCrLf & _
'                                    "<DataSet>" & vbCrLf & _
'                                    "  <DataTable>" & vbCrLf & _
'                                    JustDoIt & _
'                                    "  </DataTable>" & vbCrLf & _
'                                    "</DataSet>"
'            End If
'
'            '   PS:
'            '   XML(�]�ƦW�١B�R��覡�B�w�ˤ���B�����B�]�ƫ����BCM�t�v�B�ʺAIP�ơBCM�}������BCM�������
'            '           �BCM�b���ҥΤ���BCM�b�����Τ���B�]�ƫO�����BCM�Ǹ��BPort�ơB��ץN�X�B��צW��
'            '           �B�j�������BEBT�X���s���B�A�ȧO)
'
'        End If


'Private Function GetType1SQL(strCustID As String, _
'                                                    strFaciSno As String, _
'                                                    strStartDate As String, _
'                                                    strStopDate As String) As String
'  On Error GoTo ChkErr
'
'
'    GetType1SQL = "SELECT BILLNO,ITEM,CITEMNAME,SHOULDDATE,SHOULDAMT" & _
'                                ",REALDATE,REALAMT,REALPERIOD,REALSTARTDATE,REALSTOPDATE" & _
'                                ",UCNAME,STNAME,DECODE(CANCELFLAG,1,'�O','�_') CANCELFLAG" & _
'                                ",SERVICETYPE,FACISNO" & _
'                                " FROM " & strOwner & "SO033" & Get_DB_Link(cn) & _
'                                " WHERE CUSTID=" & strCustID & _
'                                " AND FACISNO='" & strFaciSno & "'" & _
'                                " AND SERVICETYPE IN ('I','P')" & _
'                                " AND SHOULDDATE >= " & PrcType(strStartDate, FldDate) & _
'                                " AND SHOULDDATE < " & PrcType(strStopDate, FldDate) & "+1" & _
'                                " UNION ALL " & _
'                                "SELECT BILLNO,ITEM,CITEMNAME,SHOULDDATE,SHOULDAMT" & _
'                                ",REALDATE,REALAMT,REALPERIOD,REALSTARTDATE,REALSTOPDATE" & _
'                                ",UCNAME,STNAME,DECODE(CANCELFLAG,1,'�O','�_') CANCELFLAG" & _
'                                ",SERVICETYPE,FACISNO" & _
'                                " FROM " & strOwner & "SO034" & Get_DB_Link(cn) & _
'                                " WHERE CUSTID=" & strCustID & _
'                                " AND FACISNO='" & strFaciSno & "'" & _
'                                " AND SERVICETYPE IN ('I','P')" & _
'                                " AND SHOULDDATE >= " & PrcType(strStartDate, FldDate) & _
'                                " AND SHOULDDATE < " & PrcType(strStopDate, FldDate) & "+1"
'
'    '      Select BillNo, Item, CitemName, ShouldDate, ShouldAmt, RealDate, RealAmt, RealPeriod,
'    '           RealStartDate, RealStopDate, UCName, STName, CancelFlag, ServiceType From SO033
'    '           Where CustId=[CustId] And FaciSno=[FaciSno] And ServiceType In ('I','P')
'    '           And ShouldDate>=[�_�l���] And ShouldDate < [�I����]+1
'    '      Union All
'    '          Select BillNo, Item, CitemName, ShouldDate, ShouldAmt, RealDate, RealAmt, RealPeriod,
'    '           RealStartDate, RealStopDate, UCName, STName, CancelFlag, ServiceType From SO034
'    '           Where CustId=[CustId] And FaciSno=[FaciSno] And ServiceType In ('I','P')
'    '           And ShouldDate>=[�_�l���] And ShouldDate < [�I����]+1
'
'  Exit Function
'ChkErr:
'    ErrHandle "mod_Func23", "GetType1SQL"
'    GetType1SQL = "-99," & strErr
'End Function
'
'Private Function GetType2SQL(strCustID As String, _
'                                                    strFaciSno As String, _
'                                                    strStartDate As String, _
'                                                    strStopDate As String) As String
'  On Error GoTo ChkErr
'
'
'    GetType2SQL = "SELECT A.FACINAME,A.BUYNAME,A.OLDINSTDATE,A.PRDATE,A.MODELNAME" & _
'                                ",B.CMBAUDRATE,B.DYNIPCOUNT,A.CMOPENDATE,A.CMCLOSEDATE" & _
'                                ",A.ENABLEACCOUNT,A.DISABLEACCOUNT,A.DEPOSIT,A.REFACISNO" & _
'                                ",A.PORTNO,C.BPCODE,C.BPNAME,A.CONTRACTDATE,A.EBTCONTNO" & _
'                                ",A.SERVICETYPE,A.FACISNO" & _
'                                " FROM " & strOwner & "SO004" & Get_DB_Link(cn) & " A " & _
'                                " ," & strOwner & "CD078D" & Get_DB_Link(cn) & " B " & _
'                                " ," & strOwner & "SO003" & Get_DB_Link(cn) & " C " & _
'                                " WHERE A.CUSTID=" & strCustID & _
'                                " AND A.FACISNO='" & strFaciSno & "'" & _
'                                " AND A.SERVICETYPE='I' " & _
'                                " AND A.SERVICETYPE=C.SERVICETYPE" & _
'                                " AND A.FACISNO=C.FACISNO" & _
'                                " AND C.BPCODE=B.BPCODE" & _
'                                " UNION ALL " & _
'                                "SELECT FACINAME,BUYNAME,OLDINSTDATE,PRDATE,MODELNAME" & _
'                                ",CMBAUDRATE,DYNIPCOUNT,CMOPENDATE,CMCLOSEDATE" & _
'                                ",ENABLEACCOUNT,DISABLEACCOUNT,DEPOSIT,REFACISNO" & _
'                                ",PORTNO,BPCODE,BPNAME,CONTRACTDATE,EBTCONTNO" & _
'                                ",SERVICETYPE,FACISNO" & _
'                                " FROM " & strOwner & "SO004" & Get_DB_Link(cn) & _
'                                " WHERE CUSTID=" & strCustID & _
'                                " AND REFACISNO='" & strFaciSno & "'" & _
'                                " AND SERVICETYPE='P' "
'
'    '       Select a.FaciName, a.BuyName, OldInstDate,PRDate,a.ModelName,b.CMBaudRate,
'    '       b.DynIPCount,CMOpenDate,CMCloseDate,EnableAccount, DisableAccount,Deposit,
'    '       ReFaciSNo,PortNo,c.BPCode,c.BPName,ContractDate,EBTContNo,a.ServiceType
'    '       From SO004 a, CD078D b, SO003 c
'    '       Where a.ServiceType='I' And a.Custid=[Custid] And a.FaciSno=[FaciSNo] And
'    '       a.ServiceType=c.ServiceType And a.FaciSno=c.FaciSNo And c.bpcode=b.bpcode
'    '       Union All
'    '       Select FaciName, BuyName, OldInstDate,PRDate,ModelName,CMBaudRate,
'    '       DynIPCount,CMOpenDate,CMCloseDate,EnableAccount, DisableAccount,Deposit,
'    '       ReFaciSNo,PortNo,BPCode,BPName,ContractDate,EBTContNo,ServiceType
'    '       From SO004
'    '       Where ServiceType='P' And Custid=[Custid] And ReFaciSno=[FaciSNo]
'
'  Exit Function
'ChkErr:
'    ErrHandle "mod_Func23", "GetType2SQL"
'    GetType2SQL = "-99," & strErr
'End Function
'
