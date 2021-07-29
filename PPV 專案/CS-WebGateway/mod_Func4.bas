Attribute VB_Name = "mod_Func4"
Option Explicit

Private strAuthenticID As String
Private cn As Object

Public Function JustDoIt(ByRef objWebPool As Object, _
                                        ByRef strPara As String) As String
  On Error GoTo ChkErr
    Dim strQry As String
    Dim varPara As Variant
    Dim rs As Object
    Dim lngCustID As Long
    Dim strOwner As String
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
    
    varPara = Split(strPara, ",")
    '*****************************************************
    '#3018 �W�[�P�_�ǤJ���ӼƬO�_���T By Kin 2007/08/30
    If UBound(varPara) <> 4 Then
        JustDoIt = "-99,�ѼƤ���!"
        GoTo 66
    End If
    '*****************************************************

    strAuthenticID = varPara(2)
    bytComp = Val(Left(strAuthenticID, 3))
    strOwner = GetOwner(cn)
    
    'Select count(*) as Cnt from so002b where MemberId=[�|���N��] and AuthenticId=[�{��ID]
    strQry = "SELECT COUNT(*) AS CNT FROM " & strOwner & "SO002B" & Get_DB_Link(cn) & _
                    "WHERE MEMBERID='" & varPara(1) & "' AND AUTHENTICID='" & strAuthenticID & "'"
    Set rs = cn.Execute(strQry)
    
    'If Cnt = 0 Then
    '  �^��(-2,�q��PPV�{�ҿ��~)
    'End If
    
    If rs.RecordCount <= 0 Then
        JustDoIt = "-2,�q��PPV�{�ҿ��~"
        GoTo 66
    End If
    
    'XML(�q��渹�B�q�涵���B�`�ئW�١B�]�Ƭy�����BSTB�Ǹ��B���z�d�Ǹ��B���z����B���b�B
    '   �w�I�`�ءB�q�ʤ覡�B���B�B�I�ơB��������B�����H���B�H�Υd�d��)
    
    'Select a.�q��渹, a.�q�涵��, a.�`�ئW��, a.�]�Ƭy����, a.STB�Ǹ�, a.���z�d�Ǹ�, b.���z���, a.���b,
    '   a.�w�I�`��, a.�q�ʤ覡, a.���B, a.�I��, a.�������, a.�����H��, b.�H�Υd�d��
    '   from so105g a, so105e b
    '   where b.OrderNo=a.OrderNo and b.AuthenticId=[�{��ID] and
    '    b.AcceptTime >=[�q�ʰ_��] and b.AcceptTime<=[�q�ʨ���]
    '   order by b.AcceptTime Desc, a.�]�Ƭy����, a.�q��渹, a.�q�涵��
    
    ' 4�B�|���N���B�{��ID�B�q�ʰ_��B�q�ʨ���
    
    '�w�I�`��: 0����I , 1���w�I
    '�H�Υd�d��: �e�|�X�P��|�X���e���`, �����K�X�����h��XXXXXXXX.
    
    strQry = "SELECT A.ORDERNO,A.ORDERITEM ITEM,A.PRODUCTNAME,A.FACISEQNO," & _
                    "A.FACISNO,A.SMARTCARDNO,B.ACCEPTTIME," & _
                    "DECODE(A.ACCOUNTING,1,'�O','�_') ACCOUNTING," & _
                    "DECODE(A.PREPAY,1,'�w�I','��I') PREPAY,A.PPVORDERMODE," & _
                    "A.AMOUNT,A.CREDITAMT,A.CANCELTIME,A.CANCELNAME," & _
                    "DECODE(B.ACCOUNTNO,NULL,''," & _
                    "SUBSTR(B.ACCOUNTNO,1,4) || 'XXXXXXXX' || " & _
                    "SUBSTR(B.ACCOUNTNO,13)) CREDITCARD," & _
                    "C.EVENTBEGINTIME DEVENTBEGINTIME,A.PRODUCTID" & _
                    " FROM " & _
                     strOwner & "SO105G" & Get_DB_Link(cn) & "A," & _
                     strOwner & "SO105E" & Get_DB_Link(cn) & "B," & _
                     strOwner & "SO155" & Get_DB_Link(cn) & "C " & _
                    "WHERE B.ORDERNO=A.ORDERNO AND" & _
                    " B.AUTHENTICID='" & strAuthenticID & "' AND" & _
                    " B.ACCEPTTIME >= " & PrcType(varPara(3), FldTime) & " AND" & _
                    " B.ACCEPTTIME <= " & PrcType(varPara(4), FldTime) & "+1" & _
                    " AND A.PRODUCTID=C.PRODUCTID(+)" & _
                    " ORDER BY B.ACCEPTTIME DESC,A.FACISEQNO,A.ORDERNO,A.ORDERITEM"
                    
    '    SELECT A.ORDERNO,A.ORDERITEM,A.PRODUCTNAME,A.FACISEQNO,A.FACISNO,A.SMARTCARDNO,B.ACCEPTTIME,
    '        DECODE(A.ACCOUNTING,1,'�O','�_') ACCOUNTING,
    '        DECODE(A.PREPAY,1,'�w�I','��I') PREPAY,
    '        A.ORDERMODE,A.AMOUNT,A.CANCELTIME,A.CANCELNAME,
    '        SUBSTR(B.ACCOUNTNO,1,4) || 'XXXXXXXX' || SUBSTR(B.ACCOUNTNO,13) ACCOUNTNO
    '        FROM GICMIS5.SO105G A,GICMIS5.SO105E B
    '        WHERE B.ORDERNO=A.ORDERNO AND B.AUTHENTICID='0010000062' AND
    '        B.ACCEPTTIME >= TO_DATE('20050410' ,'YYYYMMDD')
    '        AND B.ACCEPTTIME <= TO_DATE('20050412' ,'YYYYMMDD')+1
    '        ORDER BY B.ACCEPTTIME DESC,A.FACISEQNO,A.ORDERNO,A.ORDERITEM
    
    '   If RecordCount = 0 Then
    '       �^��(-5,�d�L�q�ʸ��)
    '   Else
    '       ����XML��.(��select�X�Ӥ����)
    '       �^��(0,���\)��XML��
    '   End If
    
    Set rs = cn.Execute(strQry)
    
    If rs.RecordCount <= 0 Then
        JustDoIt = "-5,�d�L�q�ʸ��"
    Else
        JustDoIt = "0,���\" & vbCrLf & ExpXML(rs)
    End If
    
    'PS:  XML���X������ഫ����
    '���b: 0���_ , 1���O
    '�w�I�`��: 0����I , 1���w�I
    '�H�Υd�d��: �e�|�X�P��|�X���e���`, �����K�X�����h��XXXXXXXX.
    'Ex: 4311XXXXXXXX1234
   
66:
    On Error Resume Next
    Rlx rs
    Rlx strQry
    Rlx varPara
    Rlx strOwner
    Rlx lngCustID
    Rlx strAuthenticID
    Set cn = Nothing
    Set objCnPool = Nothing
  Exit Function
ChkErr:
    ErrHandle "mod_Func4", "JustDoIt"
    JustDoIt = "-99," & strErr
End Function

