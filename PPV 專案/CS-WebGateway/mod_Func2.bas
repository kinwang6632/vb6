Attribute VB_Name = "mod_Func2"
Option Explicit

Private strAuthenticID As String
'Private objCnPool As Object
Private cn As Object

Public Function JustDoIt(ByRef objWebPool As Object, _
                                        ByRef strPara As String) As String
  On Error GoTo ChkErr
    Dim strQry As String
    Dim varPara As Variant
    Dim rs As Object
    Dim lngCustID As Long
    Dim objCnPool As Object
    Dim strOwner As String
    Dim strAccBefore As String
    Dim strAccBalance As String
    Dim strSubQry As String
    
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
    
    '  2�B�|���N���B�{��ID�BPPV��I�q�ʱK�X
    
    varPara = Split(strPara, ",")
    '******************************************************************
    '#3018 �W�[�P�_�ǤJ���ӼƬO�_���T By Kin 2007/08/30
    If UBound(varPara) <> 3 Then
        If UBound(varPara) <> 41 Then
            JustDoIt = "-99,�ѼƤ���!"
            Exit Function
        End If
    End If
    '******************************************************************
    strAuthenticID = varPara(2)
    bytComp = Val(Left(strAuthenticID, 3))
    
    strOwner = GetOwner(cn)
    
    ' �� : Select PPVOrderPwd,CompCode,CustId from so002b where MemberId=[�|���N��] and AuthenticId=[�{��ID]
    'strQry = "SELECT PPVORDERPWD,COMPCODE,CUSTID FROM " & strOwner & "SO002B" & Get_DB_Link(cn) & _
                    "WHERE MEMBERID='" & varPara(1) & "' AND AUTHENTICID='" & strAuthenticID & "'"
    
    
    ' �� : Select PPVOrderPwd,CompCode,CustId,FaciSNo,AccountingBefore, AccountingBalance from so002b
    '       where MemberId=[�|���N��] and AuthenticId=[�{��ID]
    ' Q : 2302 , Doc : Web API����_20060413_Liga.doc
    
    ' �s : Select PPVOrderPwd,CompCode,CustId,FaciSNo,AccountingBefore, AccountingBalance, AuthenticID
    '       from so002b where MemberId=[�|���N��] and AuthenticId=[�{��ID]   --2006/04/11 Modify
    ' Q : 2378 , Doc : Web API����_20060428_Liga.doc
    
    strQry = "SELECT PPVORDERPWD,COMPCODE,CUSTID,FACISNO" & _
                    ",ACCOUNTINGBEFORE,ACCOUNTINGBALANCE,AUTHENTICID" & _
                    " FROM " & strOwner & "SO002B" & Get_DB_Link(cn) & _
                    " WHERE MEMBERID='" & varPara(1) & "' AND AUTHENTICID='" & strAuthenticID & "'"

    Set rs = cn.Execute(strQry)
    
    '   If RecordCount=0 then
    '       �^��(-2,�q��PPV�{�ҿ��~)
    '   else
    '       if PPVOrderPwd<>[ PPV��I�q�ʱK�X] then
    '            �^��(-4,���ҥ�PPV�q���v)
    '   end if
    
    ' QQ : 2577
    ' Doc : Web API����_20060703_Liga.doc
    ' API-2�W�沧��
    ' ����PPV��I�q�ʱK�X�ˮ֡A�ѼƫO�d���ʡA�ȭק��ˮֻy�k�A�����Ӭq�y�k�Y�i�C
    ' else�@if�@PPVOrderPwd<>[�@PPV��I�q�ʱK�X]�@then
    '  �^��(-4,���ҥ�PPV�q���v)

    ' �վ��API������:
    ' ��API-2�W��]�p�A��FOR�N�ŭq��|���ϥΡA�Y���ˮ�PPV��I�q�ʱK�X�F
    ' �Ӥ�API���F�Ψ�LOGIN�ˮ֥~�A��i�����Ȥ�򥻸�Ƨ�s�ϥΡA�{�p�q��|���ä��w�@�}�l���|��PPV�q�ʱK�X�A
    ' �G�վ�API-2�A���@PPV�q�ʱK�X���ˮ֡C
    
    
    If rs.RecordCount <= 0 Then
        JustDoIt = "-2,�q��PPV�{�ҿ��~"
'    ElseIf StrComp(CStr(rs("PPVORDERPWD") & ""), CStr(varPara(3)), vbBinaryCompare) = 0 Then ' By Hammer 2006/07/06
    Else
        bytComp = rs("COMPCODE")
        lngCustID = rs("CUSTID")
        strAccBefore = rs("ACCOUNTINGBEFORE") & ""
        strAccBalance = rs("ACCOUNTINGBALANCE") & ""
        
        ' select count(*) as Cnt from so004 where
        '   CompCode=[so002b.CompCode] and CustId=[so002b.CustId] and
        '   InstDate is not null and prdate is null
        '   PS : ���O�W��,��O���B,�H�Υd�I�O�W�������w�q��[�J�^�ǭ�.
        
'        strQry = "SELECT COUNT(*) AS RCDCNT FROM " & strOwner & "SO004" & Get_DB_Link(cn) & _
                        "WHERE COMPCODE=" & bytComp & " AND CUSTID=" & lngCustID & _
                        " AND INSTDATE IS NOT NULL AND PRDATE IS NULL"
        
        'Doc : Web API����_20060413_Liga.doc
        'select SEQNo, FaciSNo, SmartCardNo from so004 where
        '   CompCode=[so002b.CompCode] and CustId=[so002b.CustId] and
        '   SeqNo in [so002b.FaciSNo] and InstDate is not null and prdate is null
        '   and facicode in (select codeno from cd022 where refno=3) ;
        
        strSubQry = "SELECT FACISNO FROM " & strOwner & "SO002B" & Get_DB_Link(cn) & _
                                " WHERE MEMBERID='" & varPara(1) & "' AND AUTHENTICID='" & strAuthenticID & "'" & _
                                " AND FACISNO IS NOT NULL"
                                
        On Error Resume Next
        strSubQry = RPxx(cn.Execute(strSubQry).GetString(2, , "", ",", ""))
        If Err = 3021 Then
            Err.Clear
            strSubQry = ""
        Else
            strSubQry = Left(strSubQry, Len(strSubQry) - 1)
        End If
        
        On Error GoTo ChkErr
        If Len(strSubQry) > 0 Then
            'Doc : Web API����_20060428_Liga.doc
            'select SEQNo, FaciSNo, SmartCardNo from so004 where
            '   CompCode=[so002b.CompCode] and CustId=[so002b.CustId] and
            '   InstDate is not null and prdate is null
            '   and facicode in (select codeno from cd022 where refno=3)
            '   and SeqNo in [so002b.FaciSNo] ;  --2006/04/11 Modify
            
            '2006/06/05 Modify
            '���eAPI-2�Ȧ^�Ǹӻ{��ID�i�ϥΪ�DSTB��ơA�{�վ�W��A�u�nCD022.REFNO>0���A�һݦ^��XML���N��
            'QQ:2513 , doc : Web API����_20060605_Liga.doc
            strQry = "SELECT SEQNO,FACISNO,SMARTCARDNO FROM " & strOwner & "SO004" & Get_DB_Link(cn) & _
                            " WHERE COMPCODE=" & bytComp & " AND CUSTID=" & lngCustID & _
                            " AND SEQNO IN (" & strSubQry & ")" & _
                            " AND INSTDATE IS NOT NULL AND PRDATE IS NULL AND FACICODE IN" & _
                            " (SELECT CODENO FROM " & strOwner & "CD022" & Get_DB_Link(cn) & " WHERE REFNO > 0)"
            
'            strQry = "SELECT SEQNO,FACISNO,SMARTCARDNO FROM " & strOwner & "SO004" & Get_DB_Link(cn) & _
                            " WHERE COMPCODE=" & bytComp & " AND CUSTID=" & lngCustID & _
                            " AND SEQNO IN (" & strSubQry & ")" & _
                            " AND INSTDATE IS NOT NULL AND PRDATE IS NULL AND FACICODE IN" & _
                            " (SELECT CODENO FROM " & strOwner & "CD022" & Get_DB_Link(cn) & " WHERE REFNO=3)"
            
            '   if Cnt=0 then
            '       �^��(-14, �d�L�w��STB�]�Ƹ��)
            '   else
            '       �p�G�ѼƼƶq�W�L4�Ӯ�(�Y>4),�h��ܦ��Ȥ�{�Ҹ�T [Web�B�~�ǤJ�����]
            '       �ݼW�[update to so002b�R�O.
            '       �^��(0,���\,���O�W��,��O���B,�H�Υd�I�O�W��).
            '   end if
            
            Set rs = cn.Execute(strQry)
'            If rs("RCDCNT") = 0 Then
            If rs.EOF Then
                If UBound(varPara) + 1 > 4 Then
                    If Update_2_SO002B(varPara, strOwner) Then
                        ' Delete from so002blog where MemberId=[�|���N��] and AuthenticId=[�{��ID]
                        cn.Execute "DELETE FROM " & strOwner & "SO002BLOG" & Get_DB_Link(cn) & _
                                            " WHERE MEMBERID='" & varPara(1) & "' AND" & _
                                            " AUTHENTICID='" & strAuthenticID & "'"
                        JustDoIt = "-14, �d�L�w��STB�]�Ƹ��"
                    Else
                        strErr = "-99," & "�s�W�Ȥ��Ʀ� [�Ȥ�{�Ҹ�T��] ����!" & IIf(strErr = Empty, "", strCrLf & strErr)
                    End If
                Else
                    JustDoIt = "-14, �d�L�w��STB�]�Ƹ��"
                End If
            Else
                If UBound(varPara) + 1 > 4 Then
    '                �ݼW�[update to so002b�R�O.
                    If Update_2_SO002B(varPara, strOwner) Then
                        ' Delete from so002blog where MemberId=[�|���N��] and AuthenticId=[�{��ID]
                        cn.Execute "DELETE FROM " & strOwner & "SO002BLOG" & Get_DB_Link(cn) & _
                                            " WHERE MEMBERID='" & varPara(1) & "' AND" & _
                                            " AUTHENTICID='" & strAuthenticID & "'"
                        ' �^��(0,���\)�� ���O�W��,��O���B,�H�Υd�I�O�W�� �� �i�ϥ�DSTB��XML��
                        JustDoIt = "0,���\" & vbCrLf & _
                                            strAccBefore & "," & strAccBalance & "," & _
                                            "," & GetCG1(strOwner, lngCustID) & _
                                            "," & GetCG2(strOwner) & vbCrLf & _
                                            ExpXML(rs)
                    Else
                        If strErr = Empty Then
                            strErr = "�s�W�Ȥ��Ʀ� [�Ȥ�{�Ҹ�T��] ����!"
                        Else
                            strErr = "�s�W�Ȥ��Ʀ� [�Ȥ�{�Ҹ�T��] ����!" & strCrLf & strErr
                        End If
                        JustDoIt = "-99," & strErr
                    End If
                Else
                    ' ��
                    'PS1 : ���O�W��,��O���B,�H�Υd�I�O�W���A�Ш̸ӻ{��ID��select�X�Ӫ��Ȧ^�ǡC
                    '�H�Υd�I�O�W�� (�|���w�q�ȶǪŭ�)
                    ' �s
                    '--2006/04/28 Add
                    'PS1:  ���O�W�� , ��O���B�W��, �H�Υd�I�O�W��, �ثe��O���B, ������q����B
                    '�Ш̸ӻ{��ID��select�X�Ӫ��Ȧ^�ǡC�H�Υd�I�O�W��(�|���w�q�ȶǪŭ�)
                    
                    JustDoIt = "0,���\" & vbCrLf & _
                                        strAccBefore & "," & strAccBalance & "," & _
                                        "," & GetCG1(strOwner, lngCustID) & _
                                        "," & GetCG2(strOwner) & vbCrLf & _
                                        ExpXML(rs)
                    'PS2 : �ӻ{��ID�i�ϥ�DSTB��XML(�]�Ƭy�����B�]�ƧǸ��B���z�d�d��)
                End If
            End If
        Else
            JustDoIt = "-23,�ӻ{��ID�d�L�{�ҳ]�ƩΪA��"
'            JustDoIt = "-99,[�Ȥ�{�Ҹ�T��] �d�L [�{�ҳ]�ƧǸ�] ���!"
'            JustDoIt = "-4,���ҥ�PPV�q���v"
        End If
'    Else
'        JustDoIt = "-99,[�Ȥ�{�Ҹ�T��] �d�L [�{�ҳ]�ƧǸ�] ���!"
'        JustDoIt = "-4,���ҥ�PPV�q���v"
    End If
    
    
    '   ���G�X�B���G�T���r��(���)
    '   ���O�W���B��O���B (����ݩw�q)�B�H�Υd�I�O�W��
    
    '   0:   ���\
    '   -2 : �q��PPV�{�ҿ��~
    '   -3 : �����W���w����
    '   -4 : ���ҥ�PPV�q���v
       
66:
    On Error Resume Next
    Rlx rs
    Rlx strQry
    Rlx varPara
    Rlx strOwner
    Rlx lngCustID
    Rlx strAccBefore
    Rlx strAccBalance
    Rlx strAuthenticID
    Set cn = Nothing
    Set objCnPool = Nothing
  Exit Function
ChkErr:
    ErrHandle "mod_Func2", "JustDoIt"
    JustDoIt = "-99," & strErr
End Function

Private Function GetCG1(strOwner As String, lngCustID As Long) As String
  On Error GoTo ChkErr
    '--2006/04/28 Add  count�ӻ{��ID�w���⪺�`�ضO�P�I�Ƥ�O���B
    'select sum(shouldamt) as charge1 from so033 where compcode=[CustId] and authenticid=[AuthenticID]
    'and citemcode in (select codeno from cd019 where refno in (3,17)) and uccode is not null and cancelflag!=1
    GetCG1 = GetValue(cn, "SELECT SUM(SHOULDAMT) AS CHARGE1 FROM " & strOwner & "SO033" & Get_DB_Link(cn) & _
                                            " WHERE CUSTID=" & lngCustID & " AND AUTHENTICID='" & strAuthenticID & "'" & _
                                            " AND CITEMCODE IN (SELECT CODENO FROM " & strOwner & "CD019" & Get_DB_Link(cn) & _
                                            " WHERE REFNO IN (3,17) AND UCCODE IS NOT NULL AND CANCELFLAG<>1)")
  Exit Function
ChkErr:
    ErrHandle "mod_Func2", "GetCG1"
End Function

Private Function GetCG2(strOwner As String) As String
  On Error GoTo ChkErr
    '--2006/04/28 Add count�ӻ{��ID�|�����⪺�q����B
    'select sum(b.amount) as charge2 from so105e a, so105g b Where a.AuthenticID = [AuthenticID]
    'And a.orderno = b.orderno And b.accounting! = 1 And b.canceltime Is Null
    GetCG2 = GetValue(cn, "SELECT SUM(B.AMOUNT) AS CHARGE2 FROM " & _
                                            strOwner & "SO105E" & Get_DB_Link(cn) & " A," & strOwner & "SO105G" & Get_DB_Link(cn) & " B" & _
                                            " WHERE AUTHENTICID='" & strAuthenticID & "' AND A.ORDERNO=B.ORDERNO" & _
                                            " AND B.ACCOUNTING <> 1 AND B.CANCELTIME IS NULL")
  Exit Function
ChkErr:
    ErrHandle "mod_Func2", "GetCG2"
End Function

Private Function Update_2_SO002B(ByRef varPara As Variant, ByRef strOwner As String) As Boolean
  On Error GoTo ChkErr
'    Dim cmd As Object
    Dim intRet As Integer
    '    Set cmd = objCnPool.GetCommand
    '2�B�|���N���B�{��ID�BPPV��I�q�ʱK�X
    '�|���W�١B�����Ҧr���B�X�ͦ~���B�p���q�ܡB������X�B�ʧO�B
    '�p���a�}�BEmail�BPPV��I�q�ʱK�X�B�n�J�K�X(Web���|�����ƳQ���ʮɻݶ�)
    '    With cmd
    '         Set .ActiveConnection = cn
    '        .CommandType = 1
    '        .CommandText = "UPDATE " & strOwner & "SO002B " & _
    '                                    "SET MEMBERID=?,NAME=?,ID=?," & _
    '                                    "BIRTHDAY=TO_DATE(?, 'YYYY/MM/DD')," & _
    '                                    "TEL=?,MOBILE=?,SEXTUAL=?," & _
    '                                    "ADDRESS=?,EMAIL=?," & _
    '                                    "PPVORDERPWD=?,LOGINPASSWORD=? " & _
    '                                    "WHERE ( AUTHENTICID=? )"
    '        .Execute intRet, Array(varPara(1), varPara(4), varPara(5), _
    '                                    varPara(6), varPara(7), varPara(8), varPara(9), _
    '                                    varPara(10), varPara(11), varPara(12), _
    '                                    varPara(13), strAuthenticID), 1
    '        Update_2_SO002B = (intRet > 0)
    '        Set .ActiveConnection = Nothing
    '    End With

    '   �s�Ѽ� : �ȪA�����|���N���B�|���W�١B�����Ҧr���B�X�ͤ��-�~�B�X�ͤ��-��B�X�ͤ��-��B
    '   �p���q��(�v) �B�p���q��(��)�B��ʹq�ܡB�ʧO�B�q�T�a�}-�����B�q�T�a�}-�m���ϡB�q�T�a�}-���q��ѧ˫Ѹ��ӡB
    '   Email�B�ȪA�����n�J�b���BPPV�q�ʱK�X�B�ȪA�����|���K�X�B�K�X���ܰ��D�B�K�X���ܵ��סB
    '   �Ǿ��B�@�N���즳�u�q���A�Ȥ��W�D�`�ص������T���B�B�ê��p�B�P��H�ơB�P�����B¾�~�O�B
    '   �ӤH�~���J�B�����O�B�b�����A�B�ӽФ���ɶ��B�ҥΤ���B�̫�n�J����B�̫Ყ�ʤ���B
    '   �t�Υx�N�X�B�o�����Y�B�o������(�G�p�ΤT�p)�B�o���a�}�B�Τ@�s��(��o���������T�p�ɶ�������) �B�]�Ƭy����
    '   (Web���|�����ƳQ���ʮɻݶ�)
    Dim strSQL As String
    
    varPara(41) = Replace(varPara(41), "#", ",", 1)
    varPara(41) = Replace(varPara(41), "'", "''", 1)
    strSQL = "UPDATE " & strOwner & "SO002B" & Get_DB_Link(cn) & _
                    "SET MEMBERID=" & PrcType(varPara(4)) & ",NAME=" & PrcType(varPara(5)) & _
                    ",ID=" & PrcType(varPara(6)) & "," & "BIRTHDAYYYYY=" & PrcType(varPara(7)) & _
                    ",BIRTHDAYMM=" & PrcType(varPara(8)) & ",BIRTHDAYDD=" & PrcType(varPara(9)) & _
                    ",TEL=" & PrcType(varPara(10)) & ",TEL2=" & PrcType(varPara(11)) & _
                    ",MOBILE=" & PrcType(varPara(12)) & "," & "SEXTUAL=" & PrcType(varPara(13)) & _
                    ",CITY=" & PrcType(varPara(14)) & ",TOWN=" & PrcType(varPara(15)) & "," & _
                    "ADDRESS=" & PrcType(varPara(16)) & ",EMAIL=" & PrcType(varPara(17)) & _
                    ",LOGINID=" & PrcType(varPara(18)) & "," & "PPVORDERPWD=" & PrcType(varPara(19)) & _
                    ",LOGINPASSWORD=" & PrcType(varPara(20)) & "," & "HINTNAME=" & PrcType(varPara(21)) & _
                    ",HINTANSWER=" & PrcType(varPara(22)) & "," & "EDUCATION=" & PrcType(varPara(23)) & _
                    ",MESSAGE=" & PrcType(varPara(24)) & "," & "MARRIED=" & PrcType(varPara(25)) & _
                    ",BODYCOUNT=" & PrcType(varPara(26), FldNum) & "," & "MEMBERAGE=" & PrcType(varPara(27)) & _
                    ",JOB=" & PrcType(varPara(28)) & "," & "RENTE=" & PrcType(varPara(29)) & _
                    ",CLASSID=" & PrcType(varPara(30)) & ",STATUS=" & PrcType(varPara(31)) & _
                    ",CREATEDATE=" & PrcType(varPara(32), FldDate) & "," & "USEDATE=" & PrcType(varPara(33), FldDate) & _
                    ",LASTLOGINDATE=" & PrcType(varPara(34), FldDate) & ",UPDTIME=" & PrcType(varPara(35)) & _
                    ",COMPCODE=" & PrcType(varPara(36), FldNum) & ",INVTITLE=" & PrcType(varPara(37)) & _
                    ",INVOICETYPE=" & IIf(varPara(38) = Empty, "NULL", Val(varPara(38))) & _
                    ",INVADDRESS=" & PrcType(varPara(39)) & _
                    ",INVNO=" & PrcType(varPara(40)) & ",FACISNO=" & PrcType(varPara(41)) & _
                    " WHERE ( AUTHENTICID='" & strAuthenticID & "')"

'PrcType(varPara(38), FldNum)

'    strSQL = "UPDATE " & strOwner & "SO002B " & _
'                    "SET MEMBERID='" & varPara(1) & "'," & _
'                    "NAME='" & varPara(4) & "'," & _
'                    "ID='" & varPara(5) & "'," & _
'                    "BIRTHDAY=TO_DATE('" & varPara(6) & "', 'YYYY/MM/DD')," & _
'                    "TEL='" & varPara(7) & "'," & _
'                    "MOBILE='" & varPara(8) & "'," & _
'                    "SEXTUAL=" & varPara(9) & "," & _
'                    "ADDRESS='" & varPara(10) & "'," & _
'                    "EMAIL='" & varPara(11) & "'," & _
'                    "PPVORDERPWD='" & varPara(12) & "'," & _
'                    "LOGINPASSWORD='" & varPara(13) & "' " & _
'                    "WHERE ( AUTHENTICID='" & strAuthenticID & "')"
    cn.Execute strSQL, intRet
    Update_2_SO002B = (intRet > 0)
    strSQL = Empty
  Exit Function
ChkErr:
    ErrHandle "mod_Func2", "Update_2_SO002B"
    Update_2_SO002B = False
End Function
