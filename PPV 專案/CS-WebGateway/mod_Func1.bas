Attribute VB_Name = "mod_Func1"
Option Explicit

Private strAuthenticID As String
Private cn As Object
Private strOwner As String

'select * from so002b where facisno like '%' || (select facisno from so002b where facisno like '%''200302270069925''%') || '%'
'select * from so002b where facisno like '%' || (
'select facisno from so002b where facisno like '%''200302270069925''%' union all
'select facisno from so002b where facisno like '%''200302270069925''%' )  || '%'

' ���u�q���Ȥᨭ��SMS�b���K�X�{�ҥӽ�
' 1�B���q�O�B�Ȥ�s���B�����B�{��ID -- (�����=0��2�ɶ��ǻ{��ID) 2007/04/02 liga add (2.2����)
' C�B�Τ�m�W�B�q��(���)
' D�BDSTB�]�ƧǸ��B���z�d�d��(���)
' I�B�ӽФH�B�����ҡB�ͤ�(���)�B(�s�W) CM�X���s�� (���)
' �K�K

' �ĤG��H�U���ѼơA�L���w�{�Ҫ��A�Ⱥ������ǡA�̺����W�ҶǶi�Ӫ��A�Ȥ��P�A
' �ӶǤ��P���ѼƧY�i�A�C�W�[�@�A�Ȼ{�ҥӽЫh���U�s�W�@��C
' EX:
'  1,16,123
'  D , 119110074891#, 116010001031#
'  I , ���j��, J123456789, 19770301
'  I , �����R, H223456789, 19890607
'  C , LIGA, 3010001

' ���G�X�B���G�T���r��(���)
' �{��ID (���)
' XML (��ܳv���{�Ҧ��\ / ���Ѫ��A)

' 0:  ���\
' -22 : �ӫȽs�γ]�Ƹ�Ʃ󦹨t�Υx���s�b

Public Function JustDoIt(ByRef objWebPool As Object, _
                                        ByRef strPara As String) As String
  On Error GoTo ChkErr
    Dim rs As Object
    Dim strQry As String
    Dim strSubQry As String
    Dim strCustID As String
    Dim strItem As String
    Dim intLoop As Integer
    Dim varPara As Variant
    Dim varDetail As Variant
    Dim objCnPool As Object
    Dim strFaciSeqNo() As String
    Dim strFaciSno As String
    Dim strStatus() As String
    Dim bytLoop As Byte
    Dim bytLoop2 As Byte
    Dim bytUB As Byte
    Dim bytType As Byte
    Dim blnFlag As Boolean
    Dim strOriAuthenticID As String

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
    
    varPara = Split(strPara, vbCrLf)
    
    If UBound(varPara) = 0 Then
        JustDoIt = "-99,�ѼƸ�Ƥ����T ! �нT�{ !"
        GoTo 66
    End If
    '****************************************************************************************************************
    '#3018 �W�[�P�_�ǤJ���ӼƬO�_���T By Kin 2007/09/10
    If UBound(Split(varPara(0), ",")) < 41 Then
        JustDoIt = "-99,�ѼƤ���!"
        GoTo 66
    End If
    If UBound(varPara) < 2 Then
        JustDoIt = "-99,�ѼƤ���!"
        GoTo 66
    Else
        For intLoop = 1 To UBound(varPara)
            If varPara(intLoop) <> Empty Then
                Select Case Split(varPara(intLoop), ",")(0)
                    Case "C"
                        If UBound(Split(varPara(intLoop), ",")) < 2 Then JustDoIt = "-99,�ѼƤ���!": GoTo 66
                    Case "D"
                        If UBound(Split(varPara(intLoop), ",")) < 2 Then JustDoIt = "-99,�ѼƤ���!": GoTo 66
                    Case "I"
                        If UBound(Split(varPara(intLoop), ",")) < 3 Then JustDoIt = "-99,�ѼƤ���!": GoTo 66
                End Select
                
            End If
        Next intLoop
    End If
    '******************************************************************************************************************
'    If UBound(varPara) = 0 Then
'        varPara = Split(strPara, ",")
'        If UBound(varPara) = 0 Then
'            JustDoIt = "-99,�ѼƸ�Ƥ����T ! �нT�{ !"
'            GoTo 66
'        Else
'            bytComp = Val(varPara(1))
'            strCustID = varPara(2)
'            bytType = varPara(3)
'            On Error Resume Next
'            strOriAuthenticID = varPara(4)
'        End If
'    Else
'        varDetail = Split(varPara(0), ",")
'        bytComp = Val(varDetail(1))
'        strCustID = varDetail(2)
'        bytType = varDetail(3)
'        On Error Resume Next
'        strOriAuthenticID = varDetail(4)
'    End If
        
    varDetail = Split(varPara(0), ",")
    bytComp = Val(varDetail(1))
    strCustID = varDetail(2)
    bytType = varDetail(3)
    On Error Resume Next
    strOriAuthenticID = varDetail(4)
    
    On Error GoTo ChkErr
    strOwner = GetOwner(cn)
    
'    2006/05/04 Add
'    ��1��ѼơA�W�["����"�H�Q�ϧO�G
'    ����=1  �s�W�{�ұb��
'    ����=0  �s�Wbooking�]�ƨϥ�
    
    If bytType < 0 Or bytType > 2 Then
        JustDoIt = "-99,�ǤJ�Ѽ�[����]���~!"
        GoTo 66
    End If
    
     ' (�����=0��2�ɶ��ǻ{��ID) 2007/04/02 liga add (2.2����)
     If bytType = 0 Or bytType = 2 Then
        If Trim(strOriAuthenticID) = Empty Then
            JustDoIt = "-99,����� = 0 �� 2 �ɶ��� �{��ID!"
            GoTo 66
        End If
     End If
    
    'Select count(*) as Cnt from SO001 where Compcode=[���q�O] and Custid=[�Ƚs] ;
    strQry = "SELECT COUNT(*) CNT FROM " & strOwner & "SO001" & Get_DB_Link(cn) & _
                    "WHERE COMPCODE=" & bytComp & " AND CUSTID=" & strCustID
                
'    Set rs = cn.Execute(strQry)
'    If Not GetRS(cn, rs, strQry) Then
'        JustDoIt = "-99,�d�ߥ���! SQL:" & strQry
'        GoTo 66
'    End If
    
'    If rs.EOF Then '    If Cnt = 0 Then
    If Val(GetValue(cn, strQry) & "") = 0 Then
        JustDoIt = "-22, �ӫȽs�γ]�Ƹ�Ʃ󦹨t�Υx���s�b" '      �^��(-22, �ӫȽs�γ]�Ƹ�Ʃ󦹨t�Υx���s�b)
        GoTo 66
    End If '    End If
    
    bytUB = UBound(varPara)
    ReDim strStatus(Val(bytUB) - 1) As String
    ReDim strFaciSeqNo(Val(bytUB) - 1) As String
    
    bytLoop = 0
    bytLoop2 = 0
    
    For intLoop = 1 To bytUB '    Loop�B�z (�ĤG��Ѽƶ}�l)
        strItem = Trim(varPara(intLoop) & "")
        If strItem <> Empty Then
            varDetail = Split(varPara(intLoop), ",")
            Select Case Trim(UCase(varDetail(0)))
                Case "C" '    CASE WHEN [����]='C' THEN
                    ' C�B�Τ�m�W�B�q��(���)
                    ' Select 0 as Flag, null as FaciSEQNo, 'CATV�Τ� �Ƚs'||Custid||' �w�Q�{��(����)�ϥ�' as Status
                    ' from SO002B where Compcode=[���q�O] and Custid=[�Ƚs] and FaciSNo like [�Ƚs] ;
                    ' strQry = "SELECT 0 AS FLAG,NULL AS FACISEQNO,'CATV�Τ� �Ƚs'||CUSTID||' �w�Q�{��(����)�ϥ�' AS STATUS"
                    strQry = "SELECT 'CATV�Τ� �Ƚs'||CUSTID||' �w�Q�{��(����)�ϥ�' AS STATUS" & _
                                    " FROM " & strOwner & "SO002B" & Get_DB_Link(cn) & " WHERE COMPCODE=" & bytComp & _
                                    " AND CUSTID=" & strCustID & " AND FACISNO LIKE '%''" & strCustID & "''%'"
                    'cable_liga: ���G�X�B���G�T���r��(���)
                    '�{��ID (���)
                    'XML(Flag�B�]�ƧǸ��B�{�Ҫ��A)
                    '
                    'Hammer: <Flag>��</Flag><�]�ƧǸ�>��</�]�ƧǸ�>
                    'cable_liga: Select count as Cnt from SO001 where Compcode=[���q�O] and Custid=[�Ƚs] ;
                    'If Cnt = 0 Then
                    '  �^��(-22, �ӫȽs�γ]�Ƹ�Ʃ󦹨t�Υx���s�b)
                    'End If
'                    If GetRS(cn, rs, strQry) Then ' rs > 0
                    Set rs = cn.Execute(strQry)
                    If Err = 0 Then
                        If Not rs.EOF Then ' If Status Is Not Null Then
                            ' [�O��select�X�Ӫ�Status]
                            strStatus(bytLoop) = MakeXML("0", "", rs!Status & "")
                            bytLoop = bytLoop + 1
                            If bytType = 2 Then
                                strQry = "SELECT A.CUSTID AS FACISEQNO" & _
                                                " FROM " & strOwner & "SO001" & Get_DB_Link(cn) & " A" & _
                                                "," & strOwner & "SO002" & Get_DB_Link(cn) & " B" & _
                                                " WHERE A.CUSTID=B.CUSTID AND A.CUSTID=" & strCustID & _
                                                " AND B.CUSTSTATUSCODE=1 AND B.SERVICETYPE='C'" & _
                                                " AND A.COMPCODE=" & bytComp & " AND" & _
                                                " (A.TEL1='" & varDetail(2) & "' OR A.TEL2='" & varDetail(2) & "' OR A.TEL3='" & varDetail(2) & "')" & _
                                                " AND A.CUSTNAME='" & varDetail(1) & "'"
                                strFaciSeqNo(bytLoop2) = "'" & GetValue(cn, strQry) & "" & "'"
                                bytLoop2 = bytLoop2 + 1
                            End If
'                            strFaciSeqNo = strFaciSeqNo & ",'" & rs!FaciSeqNo & "'"
                        Else  ' Else
                            ' Select 1 as Flag, a.CustID as FaciSEQNo, 'CATV�Τ� �Ƚs'|| a.CustID||' �ӽл{�Ҧ��\' as Status
                            ' from SO001 a, SO002 b where a.Custid=b.Custid and a.Custid=[�Ƚs] and
                            ' b.custstatuscode=1 and b.ServiceType='C' and a.Compcode=[���q�O] and a.Tel1=[�q��];
                            strQry = "SELECT 1 AS FLAG,A.CUSTID AS FACISEQNO,'CATV�Τ� �Ƚs ' ||A.CUSTID|| ' �ӽл{�Ҧ��\' AS STATUS" & _
                                            " FROM " & strOwner & "SO001" & Get_DB_Link(cn) & " A" & _
                                            "," & strOwner & "SO002" & Get_DB_Link(cn) & " B" & _
                                            " WHERE A.CUSTID=B.CUSTID AND A.CUSTID=" & strCustID & _
                                            " AND B.CUSTSTATUSCODE=1 AND B.SERVICETYPE='C'" & _
                                            " AND A.COMPCODE=" & bytComp & " AND" & _
                                            " (A.TEL1='" & varDetail(2) & "' OR A.TEL2='" & varDetail(2) & "' OR A.TEL3='" & varDetail(2) & "')" & _
                                            " AND A.CUSTNAME='" & varDetail(1) & "'"
'                            If GetRS(cn, rs, strQry) Then
                            Set rs = cn.Execute(strQry)
                            If Err = 0 Then
                                If Not rs.EOF Then
                                    ' [�O��select�X�Ӫ�Status]
                                    blnFlag = True
                                    strStatus(bytLoop) = MakeXML("1", rs!FaciSeqNo & "", rs!Status & "")
                                    bytLoop = bytLoop + 1
                                    If rs!FaciSeqNo & "" <> Empty Then
                                        strFaciSeqNo(bytLoop2) = "'" & rs!FaciSeqNo & "'"
                                        bytLoop2 = bytLoop2 + 1
                                    End If
                                Else ' else    --2006/05/17 Add
                                    ' Select 2 as Flag, null as FaciSEQNo, '�d�LCATV�Τ�'||[�Τ�m�W]||'�������' as Status from dual;
                                    ' [�O��select�X�Ӫ�Status]
                                    strStatus(bytLoop) = MakeXML("2", "", "�d�LCATV�Τ�" & varDetail(1) & "�������")
                                    bytLoop = bytLoop + 1
                                End If
                            Else
                                Err.Clear
                                JustDoIt = "-99,�d�ߥ���! SQL:" & strQry
                            End If
                        End If  ' End If
                    Else
                        Err.Clear
                        JustDoIt = "-99,�d�ߥ���! SQL:" & strQry
                    End If
                Case "D" '    CASE WHEN [����]='D' THEN
                    ' D�BDSTB�]�ƧǸ��B���z�d�d��(���)
                    ' Select 0 as Flag, null as FaciSEQNo, '�Ʀ�A�� DSTB�Ǹ�'||[DSTB�]�ƧǸ�]||' �w�Q�{��(����)�ϥ�' as Status
                    ' from SO002B where Compcode=[���q�O] and Custid=[�Ƚs] and FaciSNo like
                    ' (Select SeqNo from SO004 where Compcode=[���q�O] and Custid=[�Ƚs]
                    ' and FaciSNo=[DSTB�]�ƧǸ�] and SmartCardNo=[���z�d�d��] ) ;
                    
                    ' strQry = "SELECT 0 AS FLAG,NULL AS FACISEQNO,'�Ʀ�A�� DSTB�Ǹ�'||[" & varDetail(1) & "]||' �w�Q�{��(����)�ϥ�' AS STATUS"
                    
'                    JustDoIt = JustDoIt & "�]��D�F"
                    
                    strQry = "SELECT '�Ʀ�A�� DSTB�Ǹ� " & Trim(varDetail(1)) & " �w�Q�{��(����)�ϥ�' AS STATUS" & _
                                    " FROM " & strOwner & "SO002B" & Get_DB_Link(cn) & _
                                    " WHERE COMPCODE=" & bytComp & _
                                    " AND CUSTID=" & strCustID & " AND FACISNO LIKE '%''' || " & _
                                    " (SELECT SEQNO FROM " & strOwner & "SO004" & Get_DB_Link(cn) & _
                                    " WHERE COMPCODE=" & bytComp & " AND CUSTID=" & strCustID & _
                                    " AND FACISNO='" & varDetail(1) & "' AND SMARTCARDNO='" & varDetail(2) & "'" & _
                                    " AND INSTDATE IS NOT NULL AND PRDATE IS NULL) || '''%'"
                                    ' and InstDate is not null and PRDate is null
'                                    " AND INSTDATE IS NOT NULL AND PRDATE IS NULL" & ") || '''%'"

'                    JustDoIt = JustDoIt & vbCrLf & "SQL1 : " & strQry ' Debug

'                    If GetRS(cn, rs, strQry) Then
                    Set rs = cn.Execute(strQry)
                    If Err = 0 Then
                        If Not rs.EOF Then ' If Status Is Not Null Then
                            ' [�O��select�X�Ӫ�Status]
                            strStatus(bytLoop) = MakeXML("0", "", rs!Status & "")
'                            JustDoIt = JustDoIt & vbCrLf & "SQL1 : �������"  ' Debug
                            bytLoop = bytLoop + 1
                            If bytType = 2 Then
                                strQry = "SELECT SEQNO AS FACISEQNO" & _
                                                " FROM " & strOwner & "SO004" & Get_DB_Link(cn) & _
                                                " WHERE COMPCODE=" & bytComp & " AND CUSTID=" & strCustID & _
                                                " AND FACISNO='" & varDetail(1) & "' AND SMARTCARDNO='" & varDetail(2) & "'" & _
                                                " AND INSTDATE IS NOT NULL AND PRDATE IS NULL"
                                strFaciSeqNo(bytLoop2) = "'" & GetValue(cn, strQry) & "" & "'"
                                bytLoop2 = bytLoop2 + 1
'                                JustDoIt = JustDoIt & vbCrLf & "����2" & " SeqNo : " & strFaciSeqNo(bytLoop2) ' Debug ��
                            End If
                        Else  ' Else
                            ' Select 1 as Flag, SeqNo as FaciSEQNo, '�Ʀ�A�� DSTB�Ǹ�' ||FaciSNo||' �ӽл{�Ҧ��\' as Status
                            ' from SO004 where Compcode=[���q�O] and Custid=[�Ƚs]
                            ' and FaciSNo=[DSTB�]�ƧǸ�] and SmartCardNo=[���z�d�d��] ;
                            
'                            JustDoIt = JustDoIt & vbCrLf & "SQL1 : �S�����"  ' Debug
                            
                            strQry = "SELECT 1 AS FLAG,SEQNO AS FACISEQNO,'�Ʀ�A�� DSTB�Ǹ� ' ||FACISNO||' �ӽл{�Ҧ��\' AS STATUS" & _
                                            " FROM " & strOwner & "SO004" & Get_DB_Link(cn) & _
                                            " WHERE COMPCODE=" & bytComp & " AND CUSTID=" & strCustID & _
                                            " AND FACISNO='" & varDetail(1) & "' AND SMARTCARDNO='" & varDetail(2) & "'" & _
                                            " AND INSTDATE IS NOT NULL AND PRDATE IS NULL"
'                            If GetRS(cn, rs, strQry) Then
                            
'                            JustDoIt = JustDoIt & vbCrLf & "SQL2 : " & strQry  ' Debug
                            
                            Set rs = cn.Execute(strQry)
                            If Err = 0 Then
                                If Not rs.EOF Then
                                    ' [�O��select�X�Ӫ�Status]
'                                    JustDoIt = JustDoIt & vbCrLf & "SQL2 : ������� FaciSeqNo : " & rs!FaciSeqNo & ""   ' Debug
                                    blnFlag = True
                                    strStatus(bytLoop) = MakeXML("1", rs!FaciSeqNo & "", rs!Status & "")
                                    bytLoop = bytLoop + 1
                                    If rs!FaciSeqNo & "" <> Empty Then
                                        strFaciSeqNo(bytLoop2) = "'" & rs!FaciSeqNo & "'"
                                        bytLoop2 = bytLoop2 + 1
                                    End If
                                Else ' else     --2006/05/17 Add
                                    ' Select 2 as Flag, null as FaciSEQNo, '�d�L�Ʀ�A�� DSTB�Ǹ�'||[DSTB�]�ƧǸ�]||'�������' as Status from dual;
                                    ' [�O��select�X�Ӫ�Status]
                                    
'                                    JustDoIt = JustDoIt & vbCrLf & "SQL2 : �S�����"   ' Debug
                                    
                                    strStatus(bytLoop) = MakeXML("2", "", "�d�L�Ʀ�A�� DSTB�Ǹ�" & varDetail(1) & "�������")
                                    bytLoop = bytLoop + 1
                                End If
                            Else
                                Err.Clear
                                JustDoIt = "-99,�d�ߥ���! SQL:" & strQry
                            End If
                        End If  ' End If
                    Else
                        Err.Clear
                        JustDoIt = "-99,�d�ߥ���! SQL:" & strQry
                    End If
                Case "I" '    CASE WHEN [����]='I' THEN
                    ' I�B�ӽФH�B�����ҡB�ͤ�(���)
                    ' Select 0 as Flag, null as FaciSEQNO, '�e�W�A�� '||[�ӽФH]|| '�ӽФ�CM�Ǹ��w�Q�{��(����)�ϥ�' as Status
                    ' from SO002B where Compcode=[���q�O] and Custid=[�Ƚs] and FaciSNo like
                    ' (Select SeqNo from SO004 where Compcode=[���q�O] and Custid=[�Ƚs]
                    ' and DeclarantName=[�ӽФH] and ID=[������] and Birthday=[�ͤ�]) ;
                    ' ��ӼW�[ " and EBTContNo=[CM�X���s��] "
                    ' strQry = "SELECT 0 AS FLAG,NULL AS FACISEQNO,'�e�W�A�� '||[�ӽФH]|| '�ӽФ�CM�Ǹ��w�Q�{��(����)�ϥ�' AS STATUS"
                    strQry = "SELECT '�e�W�A�� " & varDetail(1) & " �ӽФ�CM�Ǹ��w�Q�{��(����)�ϥ�' AS STATUS" & _
                                    " FROM " & strOwner & "SO002B" & Get_DB_Link(cn) & _
                                    " WHERE COMPCODE=" & bytComp & _
                                    " AND CUSTID=" & strCustID & " AND FACISNO LIKE '%''' || " & _
                                    " (SELECT SEQNO FROM " & strOwner & "SO004" & Get_DB_Link(cn) & _
                                    " WHERE COMPCODE=" & bytComp & " AND CUSTID=" & strCustID & _
                                    " AND DECLARANTNAME='" & varDetail(1) & "' AND ID='" & varDetail(2) & "'" & _
                                    " AND BIRTHDAY=" & PrcType(varDetail(3), FldDate) & _
                                    " AND EBTCONTNO='" & varDetail(4) & "'" & _
                                    " AND INSTDATE IS NOT NULL AND PRDATE IS NULL) || '''%'"
'                                    " AND INSTDATE IS NOT NULL AND PRDATE IS NULL" & ") || '''%'"
'                    If GetRS(cn, rs, strQry) Then
                    Set rs = cn.Execute(strQry)
                    If Err = 0 Then
                        If Not rs.EOF Then ' If Status Is Not Null Then
                            ' [�O��select�X�Ӫ�Status]
                            strStatus(bytLoop) = MakeXML("0", "", rs!Status & "")
                            bytLoop = bytLoop + 1
                            If bytType = 2 Then
                                strQry = "SELECT SEQNO AS FACISEQNO" & _
                                                " FROM " & strOwner & "SO004" & Get_DB_Link(cn) & _
                                                " WHERE COMPCODE=" & bytComp & " AND CUSTID=" & strCustID & _
                                                " AND DECLARANTNAME='" & varDetail(1) & "' AND ID='" & varDetail(2) & "'" & _
                                                " AND BIRTHDAY=" & PrcType(varDetail(3), FldDate) & _
                                                " AND INSTDATE IS NOT NULL AND PRDATE IS NULL" & _
                                                " AND EBTCONTNO='" & varDetail(4) & "'"
                                strFaciSeqNo(bytLoop2) = "'" & GetValue(cn, strQry) & "" & "'"
                                bytLoop2 = bytLoop2 + 1
                            End If
                        Else  ' Else
                            ' Select 1 as Flag, SeqNo as FaciSEQNo , '�e�W�A�� CM�Ǹ�'||FaciSNo||' �ӽл{�Ҧ��\' as Status
                            ' from SO004 where Compcode=[���q�O] and Custid=[�Ƚs]
                            ' and DeclarantName=[�ӽФH] and ID=[������] and Birthday=[�ͤ�] ;
                            strQry = "SELECT 1 AS FLAG,SEQNO AS FACISEQNO,'�e�W�A�� CM�Ǹ� '|| FACISNO ||' �ӽл{�Ҧ��\' AS STATUS" & _
                                            " FROM " & strOwner & "SO004" & Get_DB_Link(cn) & _
                                            " WHERE COMPCODE=" & bytComp & " AND CUSTID=" & strCustID & _
                                            " AND DECLARANTNAME='" & varDetail(1) & "' AND ID='" & varDetail(2) & "'" & _
                                            " AND BIRTHDAY=" & PrcType(varDetail(3), FldDate) & _
                                            " AND INSTDATE IS NOT NULL AND PRDATE IS NULL" & _
                                            " AND EBTCONTNO='" & varDetail(4) & "'"
'                            If GetRS(cn, rs, strQry) Then
                            Set rs = cn.Execute(strQry)
                            If Err = 0 Then
                                If Not rs.EOF Then
                                    ' [�O��select�X�Ӫ�Status]
                                    blnFlag = True
                                    strStatus(bytLoop) = MakeXML("1", rs!FaciSeqNo & "", rs!Status & "")
                                    bytLoop = bytLoop + 1
                                    If rs!FaciSeqNo & "" <> Empty Then
                                        strFaciSeqNo(bytLoop2) = "'" & rs!FaciSeqNo & "'"
                                        bytLoop2 = bytLoop2 + 1
                                    End If
                                Else ' else     --2006/05/17 Add
                                    ' Select 2 as Flag, null as FaciSEQNo, '�d�L�e�W�A�� CM�Τ�'||[�ӽФH]||'�������' as Status from dual;
                                    ' [�O��select�X�Ӫ�Status]
                                    strStatus(bytLoop) = MakeXML("2", "", "�d�L�e�W�A�� CM�Τ�" & varDetail(1) & "�������")
                                    bytLoop = bytLoop + 1
                                End If
                            Else
                                Err.Clear
                                JustDoIt = "-99,�d�ߥ���! SQL:" & strQry
                            End If
                        End If  ' End If
                    Else
                        Err.Clear
                        JustDoIt = "-99,�d�ߥ���! SQL:" & strQry
                    End If
            End Select
        End If
    Next
    
    If Len(strStatus(0)) = 0 And Len(strFaciSeqNo(0)) = 0 Then
        JustDoIt = "-22, �ӫȽs�γ]�Ƹ�Ʃ󦹨t�Υx���s�b"
        GoTo 66
'    Else
'        If UBound(strFaciSeqNo) = 0 And Len(strFaciSeqNo(0)) = 0 Then strFaciSeqNo(0) = vbCrLf
'        If UBound(strStatus) = 0 And Len(strStatus(0)) = 0 Then strStatus(0) = vbCrLf
    End If
    
    If bytLoop2 > 0 Then ReDim Preserve strFaciSeqNo(bytLoop2 - 1) As String
    If bytLoop > 0 Then ReDim Preserve strStatus(bytLoop - 1) As String
    
    If Not blnFlag Then
        If bytType = 2 Then ' Else If RecordSet(Flag=0��)>0 and [����]=2 then
            GoTo 99
        Else
            JustDoIt = "-22, �ӫȽs�γ]�Ƹ�Ʃ󦹨t�Υx���s�b" & vbCrLf & _
                                strOriAuthenticID & vbCrLf & _
                                "<?xml version=""1.0"" encoding=""BIG5"" standalone=""yes"" ?>" & vbCrLf & _
                                "<DataSet>" & vbCrLf & _
                                "  <DataTable>" & vbCrLf & _
                                Join(strStatus, vbCrLf) & vbCrLf & _
                                "  </DataTable>" & vbCrLf & _
                                "</DataSet>" & vbCrLf
            GoTo 66
        End If
    End If
    
    If bytType = 1 Then ' If Recordset(Flag=1��)>0 and [����]=1 then
'        If blnFlag Then

        strFaciSno = Join(strFaciSeqNo, ",")
        strFaciSno = Replace(strFaciSno, ",,", ",", 1)
        If Left(strFaciSno, 1) = "," Then strFaciSno = Mid(strFaciSno, 2)
        If Right(strFaciSno, 1) = "," Then strFaciSno = Left(strFaciSno, Len(strFaciSno) - 1)
        
'        JustDoIt = JustDoIt & vbCrLf & "�P�_ Flag �e , strFaciSno : " & strFaciSno ' Debug
        
        If Insert_2_SO002B(CStr(varPara(0)), strFaciSno) Then ' .[Web�B�~�ǤJ�����] ���
    '           .�Nselect�X�Ӥ�[FaciSeqNo]�զ�_�Ө�insert SO002B.FaciSNo
            JustDoIt = "0,���\" & vbCrLf & strAuthenticID & vbCrLf & _
                                "<?xml version=""1.0"" encoding=""BIG5"" standalone=""yes"" ?>" & vbCrLf & _
                                "<DataSet>" & vbCrLf & _
                                "  <DataTable>" & vbCrLf & _
                                Join(strStatus, vbCrLf) & vbCrLf & _
                                "  </DataTable>" & vbCrLf & _
                                "</DataSet>" & vbCrLf
'           .��SO041.AccountingBefore , AccountingBalance�ѼƭȦ^��SO002B.AccountingBefore, AccountingBalance���
'               �^��(0,���\,AuthenticId��)�Ψ�select�X�Ӥ���Ʋ���XML��
        Else
            If strErr = Empty Then strErr = "�s�W�Ȥ��Ʀ� [�Ȥ�{�Ҹ�T��] ����!"
            JustDoIt = "-99," & strErr
        End If
'        Else
'            JustDoIt = "0,���\" & vbCrLf & "" & vbCrLf & _
'                                "<?xml version=""1.0"" encoding=""BIG5"" standalone=""yes"" ?>" & vbCrLf & _
'                                "<DataSet>" & vbCrLf & _
'                                "  <DataTable>" & vbCrLf & _
'                                Join(strStatus, vbCrLf) & vbCrLf & _
'                                "  </DataTable>" & vbCrLf & _
'                                "</DataSet>" & vbCrLf
'        End If
    Else '    Else
    
99:
        strFaciSno = GetValue(cn, "SELECT FACISNO FROM " & strOwner & "SO002B" & Get_DB_Link(cn) & _
                                                " WHERE AUTHENTICID='" & strOriAuthenticID & "'") & ""
                                                
        If bytType = 0 Then '    If RecordSet(Flag=1��)>0 and [����]=0 then
            '     �D�Nselect�X�Ӥ�[FaciSeqNo]�զ�_�Ө�update SO002B.FaciSNo
            '        update SO002B set FaciSNo=FaciSNo || [FaciSeqNo] (append�W�h�A�ӫD�л\�����)
            '        where custid=[�Ƚs] and authenticid=[�{��ID];
'            JustDoIt = JustDoIt & vbCrLf & "����0"  ' Debug ��
            strFaciSno = IIf(strFaciSno <> Empty, strFaciSno & ",", "") & Join(strFaciSeqNo, ",")
        ElseIf bytType = 2 Then
            '    Else If RecordSet(Flag=1��)>0 and [����]=2 then
            '     �D�Nselect�X�Ӥ�[FaciSeqNo]�N��replace���ŭȡA��update�^SO002B.FaciSNo
            '      �^��(0,���\,AuthenticId��)�Ψ�select�X�Ӥ���Ʋ���XML��
'            JustDoIt = JustDoIt & vbCrLf & "99����2"  ' Debug ��
'            JustDoIt = JustDoIt & vbCrLf & "���� : " & UBound(strFaciSeqNo) ' Debug ��
'            JustDoIt = JustDoIt & vbCrLf & strFaciSno & "-" & Join(strFaciSeqNo, ",")  ' Debug ��
            strFaciSno = Minus2str(strFaciSno, Join(strFaciSeqNo, ","))
        End If
        
        strFaciSno = Replace(strFaciSno, ",,", ",", 1)
        If Left(strFaciSno, 1) = "," Then strFaciSno = Mid(strFaciSno, 2)
        If Right(strFaciSno, 1) = "," Then strFaciSno = Left(strFaciSno, Len(strFaciSno) - 1)
        strFaciSno = Replace(strFaciSno, "'", "''")
        
'        JustDoIt = JustDoIt & vbCrLf & strFaciSno ' Debug ��
        
         ' Debug ��
'        JustDoIt = JustDoIt & vbCrLf & _
                            "UPDATE " & strOwner & "SO002B" & Get_DB_Link(cn) & _
                            " SET FACISNO='" & strFaciSno & "' WHERE AUTHENTICID='" & strOriAuthenticID & "'"
        
        cn.Execute "UPDATE " & strOwner & "SO002B" & Get_DB_Link(cn) & _
                            " SET FACISNO='" & strFaciSno & "' WHERE AUTHENTICID='" & strOriAuthenticID & "'"
        
'        �^��(0,���\,AuthenticId��)�Ψ�select�X�Ӥ���Ʋ���XML��
        JustDoIt = "0,���\" & vbCrLf & strOriAuthenticID & vbCrLf & _
                            "<?xml version=""1.0"" encoding=""BIG5"" standalone=""yes"" ?>" & vbCrLf & _
                            "<DataSet>" & vbCrLf & _
                            "  <DataTable>" & vbCrLf & _
                            Join(strStatus, vbCrLf) & vbCrLf & _
                            "  </DataTable>" & vbCrLf & _
                            "</DataSet>" & vbCrLf
    End If

'    Else
'        �^��(-22, �ӫȽs�γ]�Ƹ�Ʃ󦹨t�Υx���s�b)
'    End If
    
66:
    On Error Resume Next
    Rlx rs
    Rlx strQry
    Rlx blnFlag
    Rlx strSubQry
    Rlx strItem
    Rlx strCustID
    Rlx intLoop
    Rlx varPara
    Rlx varDetail
    Rlx strOwner
    Rlx strAuthenticID
    Rlx strOriAuthenticID
    Rlx bytUB
    Rlx bytLoop
    Rlx bytType
    Rlx strFaciSno
    Erase strStatus
    Erase strFaciSeqNo
    Set cn = Nothing
    Set objCnPool = Nothing
  Exit Function
ChkErr:
    ErrHandle "mod_Func1", "JustDoIt"
    JustDoIt = "-99," & strErr
End Function

Private Function MakeXML(ByVal strFlag As String, _
                                    ByVal strFaciSno As String, _
                                    ByVal strStatus As String) As String
  On Error GoTo ChkErr
    MakeXML = "    <DataRow FLAG=""" & strFlag & """" & _
                            " FACISEQNO=""" & strFaciSno & """" & _
                            " STATUS=""" & strStatus & """/>"
  Exit Function
ChkErr:
    ErrHandle "mod_Func1", "MakeXML"
End Function

Private Function GetAccPara(strAccBefore As String, strAccBalance As String) As Boolean
  On Error GoTo ChkErr
    Dim rs041 As Object
    GetAccPara = False
'    If GetRS(cn, rs041, "SELECT ACCOUNTINGBEFORE,ACCOUNTINGBALANCE FROM " & _
                                    strOwner & "SO043" & Get_DB_Link(cn) & " WHERE COMPCODE=" & bytComp) Then
    
    Set rs041 = cn.Execute("SELECT ACCOUNTINGBEFORE,ACCOUNTINGBALANCE FROM " & _
                                    strOwner & "SO043" & Get_DB_Link(cn) & " WHERE COMPCODE=" & bytComp)
    If Err.Number = 0 Then
        With rs041
            If Not .EOF Then
                strAccBefore = !AccountingBefore & ""
                strAccBalance = !ACCOUNTINGBALANCE & ""
                GetAccPara = True
            End If
        End With
    Else
        Err.Clear
    End If
    On Error Resume Next
    Rlx rs041
  Exit Function
ChkErr:
    ErrHandle "mod_Func1", "GetAccPara"
End Function

' �s�W�Ȥ��Ʀ� [�Ȥ�{�Ҹ�T��] SO002B
Private Function Insert_2_SO002B(ByVal strPara As String, strFaciSno As String) As Boolean
  On Error GoTo ChkErr
    Dim intRet As Integer
    Dim strSQL As String
    Dim strInvTitle As String
    Dim strInvoiceType As String
    Dim strInvAddress As String
    Dim strInvNo As String
    Dim strAccBefore As String
    Dim strAccBalance As String
    Dim varP As Variant
    
    varP = Split(strPara, ",")
    
    strAuthenticID = GetID '        �s�Wso002b(�Ȥ�{�Ҹ�T��) �ò���AuthenticId��
    
    GetAccPara strAccBefore, strAccBalance
    
    On Error Resume Next
    strInvTitle = varP(38) & ""
    strInvoiceType = varP(39) & ""
    If strInvoiceType = Empty Then strInvoiceType = "NULL"
    strInvAddress = varP(40) & ""
    strInvNo = varP(41) & ""
 
    ' 1�B���q�O�B�Ȥ�s���B�����B�{��ID
    
'   �ȪA�����|���N���B�|���W�١B�����Ҧr���B�X�ͤ��-�~�B�X�ͤ��-��B�X�ͤ��-��B
'   �p���q��(�v) �B�p���q��(��)�B��ʹq�ܡB�ʧO�B�q�T�a�}-�����B�q�T�a�}-�m���ϡB�q�T�a�}-���q��ѧ˫Ѹ��ӡB
'   Email�B�ȪA�����n�J�b���BPPV�q�ʱK�X�B�ȪA�����|���K�X�B�K�X���ܰ��D�B�K�X���ܵ��סB�Ǿ��B
'   �@�N���즳�u�q���A�Ȥ��W�D�`�ص������T���B�B�ê��p�B�P��H�ơB�P�����B¾�~�O�B�ӤH�~���J�B
'   �����O�B�b�����A�B�ӽФ���ɶ��B�ҥΤ���B�̫�n�J����B�̫Ყ�ʤ���B�t�Υx�N�X�B�o�����Y�B�o������(�G�p�ΤT�p)�B�o���a�}�B�Τ@�s��(��o���������T�p�ɶ�������)
    
    On Error GoTo ChkErr
    strSQL = "INSERT INTO " & strOwner & "SO002B" & Get_DB_Link(cn) & _
                    "(MEMBERID,NAME,ID,BIRTHDAYYYYY,BIRTHDAYMM,BIRTHDAYDD," & _
                    "TEL,TEL2,MOBILE,SEXTUAL,CITY,TOWN,ADDRESS,EMAIL,LOGINID," & _
                    "PPVORDERPWD,LOGINPASSWORD,HINTNAME,HINTANSWER,EDUCATION," & _
                    "MESSAGE,MARRIED,BODYCOUNT,MEMBERAGE,JOB,RENTE,CLASSID," & _
                    "STATUS,CREATEDATE,USEDATE,LASTLOGINDATE,UPDTIME," & _
                    "COMPCODE,CUSTID,FACISNO,AUTHENTICID," & _
                    "INVTITLE,INVOICETYPE,INVADDRESS,INVNO," & _
                    "ACCOUNTINGBEFORE,ACCOUNTINGBALANCE) VALUES (" & _
                    PrcType(varP(5)) & "," & PrcType(varP(6)) & "," & PrcType(varP(7)) & "," & _
                    PrcType(varP(8)) & "," & PrcType(varP(9)) & "," & PrcType(varP(10)) & "," & _
                    PrcType(varP(11)) & "," & PrcType(varP(12)) & "," & PrcType(varP(13)) & "," & _
                    PrcType(varP(14)) & "," & PrcType(varP(15)) & "," & PrcType(varP(16)) & "," & _
                    PrcType(varP(17)) & "," & PrcType(varP(18)) & "," & PrcType(varP(19)) & "," & _
                    PrcType(varP(20)) & "," & PrcType(varP(21)) & "," & PrcType(varP(22)) & "," & _
                    PrcType(varP(23)) & "," & PrcType(varP(24)) & "," & PrcType(varP(25)) & "," & _
                    PrcType(varP(26)) & "," & PrcType(varP(27), FldNum) & "," & PrcType(varP(28)) & "," & _
                    PrcType(varP(29)) & "," & PrcType(varP(30)) & "," & PrcType(varP(31)) & "," & _
                    PrcType(varP(32)) & "," & PrcType(varP(33), FldDate) & "," & _
                    PrcType(varP(34), FldDate) & "," & PrcType(varP(35), FldDate) & "," & _
                    PrcType(varP(36)) & "," & varP(37) & "," & varP(2) & ",'" & Replace(strFaciSno, "'", "''", 1) & "','" & strAuthenticID & "'," & _
                    "'" & strInvTitle & "'," & strInvoiceType & ",'" & strInvAddress & "','" & strInvNo & "'," & _
                    "'" & strAccBefore & "','" & strAccBalance & "')"
    
    cn.Execute strSQL, intRet
    Insert_2_SO002B = (intRet > 0)
    
    On Error Resume Next
    strSQL = Empty
    Rlx intRet
    Rlx strSQL
    Rlx strInvNo
    Rlx strInvTitle
    Rlx strAccBefore
    Rlx strAccBalance
    Rlx strInvAddress
    Rlx strInvoiceType
  Exit Function
ChkErr:
    ErrHandle "mod_Func1", "Insert_2_SO002B"
    Insert_2_SO002B = False
End Function

' ���o AuthenticID
Private Function GetID() As String
  On Error GoTo ChkErr
    GetID = GetValue(cn, "SELECT " & strOwner & "S_SO002B_AUTHENTICID.NEXTVAL" & Get_DB_Link(cn) & " FROM DUAL")
'   PS : so002b Sequence Object: S_SO002B_AuthenticId
'   �{�ҽs��������!(���q�O3�X+�y����7�X) �k�a���ɹs , Ex: ���q�O=1, �y������123; �h�{�ҽs��='0010000123'
    GetID = Right("0000000" & GetID, 7)
    GetID = Right("000" & bytComp, 3) & GetID
  Exit Function
ChkErr:
    ErrHandle "mod_Func1", "GetID"
End Function



'Private strAuthenticID As String
'Private cn As Object
'Private strOwner As String
'
'' ���u�q���Ȥᨭ��SMS�b���K�X�{�ҥӽ�
'
'' 1�B���q�O�B�Ȥ�s���BSTB�Ǹ�
'
'' �ª��Ѽ� : �|���N���B�|���W�١B�����Ҧr���B�X�ͦ~���B�p���q�ܡB������X�B�ʧO�B�p���a�}�BEmail�BPPV��I�q�ʱK�X�B�n�J�K�X
'
'' �s���Ѽ� : �ȪA�����|���N���B�|���W�١B�����Ҧr���B�X�ͤ��-�~�B�X�ͤ��-��B�X�ͤ��-��B
''   �p���q��(�v) �B�p���q��(��)�B��ʹq�ܡB�ʧO�B�q�T�a�}-�����B�q�T�a�}-�m���ϡB�q�T�a�}-���q��ѧ˫Ѹ��ӡB
''   Email�B�ȪA�����n�J�b���BPPV�q�ʱK�X�B�ȪA�����|���K�X�B�K�X���ܰ��D�B�K�X���ܵ��סB�Ǿ��B
''   �@�N���즳�u�q���A�Ȥ��W�D�`�ص������T���B�B�ê��p�B�P��H�ơB�P�����B¾�~�O�B�ӤH�~���J�B
''   �����O�B�b�����A�B�ӽФ���ɶ��B�ҥΤ���B�̫�n�J����B�̫Ყ�ʤ���B�t�Υx�N�X
'
'' ���G�X�B���G�T���r��(���)
'' �{��ID
'
'' 0:  ���\
'' -1 : �b��{�ҿ��~
'' -3 : �����W���w����
'
'Public Function JustDoIt(ByRef objWebPool As Object, _
'                                        ByRef strPara As String) As String
'  On Error GoTo ChkErr
'    Dim strQry As String
'    Dim varPara As Variant
'    Dim rsSO004 As Object
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
'    varPara = Split(strPara, ",")
'    bytComp = Val(varPara(1))
'    strOwner = GetOwner(cn)
'
'    'select prdate from so004 where CompCode=[���q�O] and CustId=[�Ȥ�s��] and
'    ' FaciSNO=[STB�Ǹ�] order by Instdate desc(���Ĥ@��)
'
'    strQry = "SELECT PRDATE FROM " & strOwner & "SO004" & Get_DB_Link(cn) & _
'                    "WHERE COMPCODE=" & bytComp & _
'                    " AND CUSTID=" & varPara(2) & _
'                    " AND FACISNO='" & varPara(3) & "'" & _
'                    " ORDER BY INSTDATE DESC"
'
'    Set rsSO004 = cn.Execute(strQry)
'
'    '   If RecordCount = 0 Then
'    '       �^��(-1,�b��{�ҿ��~)
'    '   else if prdate is not null then
'    '       �^��(-3,�����W���w����)
'    '   Else
'    '       �s�Wso002b(�Ȥ�{�Ҹ�T��)[Web�B�~�ǤJ�����]��ƨò���AuthenticId��
'    '       �^��(0,���\,AuthenticId��)
'    '   End If
'
'    If rsSO004.RecordCount <= 0 Then
'        JustDoIt = "-1,�b��{�ҿ��~"
'    ElseIf Not IsNull(rsSO004("PRDATE")) Then
'        JustDoIt = "-3,�����W���w����"
'    Else
'        If Insert_2_SO002B(varPara) Then
'            JustDoIt = "0,���\" & vbCrLf & strAuthenticID
'        Else
'            If strErr = Empty Then strErr = "�s�W�Ȥ��Ʀ� [�Ȥ�{�Ҹ�T��] ����!"
'            JustDoIt = "-99," & strErr
'        End If
'    End If
'
'66:
'    On Error Resume Next
'    Rlx strQry
'    Rlx varPara
'    Rlx rsSO004
'    Rlx strOwner
'    Rlx strAuthenticID
'    Set cn = Nothing
'    Set objCnPool = Nothing
'  Exit Function
'ChkErr:
'    ErrHandle "mod_Func1", "JustDoIt"
'    JustDoIt = "-99," & strErr
'End Function
'
'' �s�W�Ȥ��Ʀ� [�Ȥ�{�Ҹ�T��] SO002B
'Private Function Insert_2_SO002B(ByRef varPara As Variant) As Boolean
'  On Error GoTo ChkErr
'    'Dim cmd As Object
'    Dim intRet As Integer
''    Set cmd = objCnPool.GetCommand
''    With cmd
''         Set .ActiveConnection = cn
''        .CommandType = 1
''        .CommandText = "INSERT INTO " & strOwner & "SO002B " & _
''                                "(AUTHENTICID,NAME,TEL,MOBILE,BIRTHDAY,ID,SEXTUAL," & _
''                                "ADDRESS,COMPCODE,CUSTID,FACISNO,EMAIL," & _
''                                "MEMBERID,PPVORDERPWD,LOGINPASSWORD)" & _
''                                " VALUES (?,?,?,?,TO_DATE(?, 'YYYY/MM/DD'),?,?,?,?,?,?,?,?,?,?)"
''         strAuthenticID = GetID
''        .Execute intRet, Array(strAuthenticID, varPara(5), varPara(8), varPara(9), varPara(7), _
''                                varPara(6), varPara(10), varPara(11), bytComp, varPara(2), _
''                                varPara(3), varPara(12), varPara(4), varPara(13), varPara(14)), 1
''        Insert_2_SO002B = (intRet > 0)
''         Set .ActiveConnection = Nothing
''    End With
'
''   �s���Ѽ� : �ȪA�����|���N���B�|���W�١B�����Ҧr���B�X�ͤ��-�~�B�X�ͤ��-��B�X�ͤ��-��B
''   �p���q��(�v) �B�p���q��(��)�B��ʹq�ܡB�ʧO�B�q�T�a�}-�����B�q�T�a�}-�m���ϡB�q�T�a�}-���q��ѧ˫Ѹ��ӡB
''   Email�B�ȪA�����n�J�b���BPPV�q�ʱK�X�B�ȪA�����|���K�X�B�K�X���ܰ��D�B�K�X���ܵ��סB�Ǿ��B
''   �@�N���즳�u�q���A�Ȥ��W�D�`�ص������T���B�B�ê��p�B�P��H�ơB�P�����B¾�~�O�B�ӤH�~���J�B
''   �����O�B�b�����A�B�ӽФ���ɶ��B�ҥΤ���B�̫�n�J����B�̫Ყ�ʤ���B�t�Υx�N�X
'
'    Dim strSQL As String
'    strAuthenticID = GetID
'
'
''   3�D����e������B(ACCOUNTINGBEFORE)�B����������B( ACCOUNTINGBALANCE)�A�{�^�ǬҬ�0�A
''       ����SO043�w���W�[�o2�ӰѼơA����N�Ŧ^�ǮɡA��ݦ�SO043Ū���o2�ӰѼƨæ^��SO002B�C
''    txtAccBefore = Val(GetSystemParaItem("AccountingBefore", gCompCode, gServiceType, "SO043", , True) & Empty)
''    txtAccBalance = Val(GetSystemParaItem("AccountingBalance", gCompCode, gServiceType, "SO043", , True) & Empty)
'
'    strSQL = "INSERT INTO " & strOwner & "SO002B" & Get_DB_Link(cn) & _
'                                "(MEMBERID,NAME,ID,BIRTHDAYYYYY,BIRTHDAYMM,BIRTHDAYDD," & _
'                                "TEL,TEL2,MOBILE,SEXTUAL,CITY,TOWN,ADDRESS,EMAIL,LOGINID," & _
'                                "PPVORDERPWD,LOGINPASSWORD,HINTNAME,HINTANSWER,EDUCATION," & _
'                                "MESSAGE,MARRIED,BODYCOUNT,MEMBERAGE,JOB,RENTE,CLASSID," & _
'                                "STATUS,CREATEDATE,USEDATE,LASTLOGINDATE,UPDTIME," & _
'                                "COMPCODE,CUSTID,FACISNO,AUTHENTICID) VALUES (" & _
'                                PrcType(varPara(4)) & "," & PrcType(varPara(5)) & "," & PrcType(varPara(6)) & "," & _
'                                PrcType(varPara(7)) & "," & PrcType(varPara(8)) & "," & PrcType(varPara(9)) & "," & _
'                                PrcType(varPara(10)) & "," & PrcType(varPara(11)) & "," & PrcType(varPara(12)) & "," & _
'                                PrcType(varPara(13)) & "," & PrcType(varPara(14)) & "," & PrcType(varPara(15)) & "," & _
'                                PrcType(varPara(16)) & "," & PrcType(varPara(17)) & "," & PrcType(varPara(18)) & "," & _
'                                PrcType(varPara(19)) & "," & PrcType(varPara(20)) & "," & PrcType(varPara(21)) & "," & _
'                                PrcType(varPara(22)) & "," & PrcType(varPara(23)) & "," & PrcType(varPara(24)) & "," & _
'                                PrcType(varPara(25)) & "," & PrcType(varPara(26), FldNum) & "," & PrcType(varPara(27)) & "," & _
'                                PrcType(varPara(28)) & "," & PrcType(varPara(29)) & "," & PrcType(varPara(30)) & "," & _
'                                PrcType(varPara(31)) & "," & PrcType(varPara(32), FldDate) & "," & _
'                                PrcType(varPara(33), FldDate) & "," & PrcType(varPara(34), FldDate) & "," & _
'                                PrcType(varPara(35)) & "," & bytComp & "," & varPara(2) & "," & _
'                                "'" & varPara(3) & "','" & strAuthenticID & "')"
'
''    strSQL = "INSERT INTO " & strOwner & "SO002B " & _
''                                "(AUTHENTICID,NAME,TEL,MOBILE,BIRTHDAY,ID,SEXTUAL," & _
''                                "ADDRESS,COMPCODE,CUSTID,FACISNO,EMAIL," & _
''                                "MEMBERID,PPVORDERPWD,LOGINPASSWORD)" & _
''                                " VALUES ('" & strAuthenticID & "','" & varPara(5) & "'," & _
''                                "'" & varPara(8) & "','" & varPara(9) & "'," & _
''                                "TO_DATE('" & varPara(7) & "', 'YYYY/MM/DD')," & _
''                                "'" & varPara(6) & "'," & varPara(10) & "," & _
''                                "'" & varPara(11) & "'," & bytComp & "," & _
''                                varPara(2) & ",'" & varPara(3) & "'," & _
''                                "'" & varPara(12) & "','" & varPara(4) & "'," & _
''                                "'" & varPara(13) & "','" & varPara(14) & "')"
'    cn.Execute strSQL, intRet
'    Insert_2_SO002B = (intRet > 0)
'    strSQL = Empty
'  Exit Function
'ChkErr:
'    ErrHandle "mod_Func1", "Insert_2_SO002B"
'    Insert_2_SO002B = False
'End Function
'
'' ���o AuthenticID
'Private Function GetID() As String
'  On Error GoTo ChkErr
''    GetID = RPxx(cn.Execute("SELECT " & strOwner & "S_SO002B_AUTHENTICID.NEXTVAL AUTHENTICID FROM DUAL").GetString(2, 1, "", "", "") & "")
'    GetID = GetValue(cn, "SELECT " & strOwner & "S_SO002B_AUTHENTICID.NEXTVAL AUTHENTICID FROM DUAL")
'
''   PS : so002b Sequence Object: S_SO002B_AuthenticId
''   �{�ҽs��������!(���q�O3�X+�y����7�X) �k�a���ɹs
''   Ex: ���q�O=1, �y������123; �h�{�ҽs��='0010000123'
'
'    GetID = Right("0000000" & GetID, 7)
'    GetID = Right("000" & bytComp, 3) & GetID
'  Exit Function
'ChkErr:
'    ErrHandle "mod_Func1", "GetID"
'End Function
'
                    
'                    strSubQry = "SELECT SEQNO FROM " & strOwner & "SO004" & Get_DB_Link(cn) & _
'                                        " WHERE COMPCODE=" & bytComp & " AND CUSTID=" & strCustID & _
'                                        " AND FACISNO='" & varDetail(1) & "' AND SMARTCARDNO='" & varDetail(2) & "'" & _
'                                        " AND INSTDATE IS NOT NULL AND PRDATE IS NULL"
'
'                    On Error Resume Next
'                    strSubQry = cn.Execute(strSubQry).GetString(2, , "", ",", "")
'                    If Err = 3021 Then
'                        JustDoIt = "-99,�d�ߥ���! SQL:" & strSubQry
'                        GoTo 66
'                    End If
'                    strSubQry = Left(strSubQry, Len(strSubQry) - 1)


