Attribute VB_Name = "mod_Func17"
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
    Dim strFaciSno As String
    Dim objCnPool As Object
    Dim strOwner As String
    
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
    '#3018 �W�[�P�_�ǤJ���ӼƬO�_���T By Kin 2007/09/03
    If UBound(varPara) < 2 Then
        JustDoIt = "-99,�ѼƤ���!"
        GoTo 66
    End If
    '*****************************************************

    
    strAuthenticID = varPara(2)
    bytComp = Val(Left(strAuthenticID, 3))
    lngCustID = Val(Right(strAuthenticID, 7))
    strOwner = GetOwner(cn)
    
    '�Ȥ�w��STB�]�Ƹ�Ƭd��(�w�w��STB����,�s���Φ�m)
    '17�B�|���N���B�{��ID
    'XML
    '   ----------- �ª� -----------
    '(�]�Ƭy�����BSTB�Ǹ��B���z�d�Ǹ��B�����Ǹ��B�˸m�I�B�����W�١B�^�ǼҲաB�w�ˤ���BPPV�ϥ��v�BiPPV�ϥ��v�B�`�ص���)
    '   ----------- �s�� -----------
    '(�]�Ƭy�����BSTB�Ǹ��B���z�d�Ǹ��B�����Ǹ��B�˸m�I�B�����W�١B�^�ǼҲաB�w�ˤ���BPPV�ϥ��v�BiPPV�ϥ��v�B�`�ص��šB�`�ص��ŦW��)
    '0 : ���\
    '-14 : �d�L�w��STB�]�Ƹ��
    
    '   Select CustId as nCustId from so002b where MemberId=[�|���N��] and AuthenticId=[�{��ID]
    '   If nCustId Is Null Then
    '       �^��(-2,�q��PPV�{�ҿ��~)
    '   End If
    
'    strQry = "SELECT COUNT(*) AS RCDCNT FROM " & strOwner & "SO002B" & Get_DB_Link(cn) & _
                    "WHERE MEMBERID='" & varPara(1) & "' AND AUTHENTICID='" & strAuthenticID & "'"
    
'    If Val(GetValue(cn, strQry)) = 0 Then
'        JustDoIt = "-2,�q��PPV�{�ҿ��~"
'        GoTo 66
'    End If
    
'    strQry = "SELECT CUSTID FROM " & strOwner & "SO002B" & Get_DB_Link(cn) & _
'                    "WHERE MEMBERID='" & varPara(1) & "' AND AUTHENTICID='" & strAuthenticID & "'"
'
'    lngCustID = Val(GetValue(cn, strQry) & "")
'
'    If lngCustID = 0 Then
'        JustDoIt = "-2,�q��PPV�{�ҿ��~"
'        GoTo 66
'    End If
    
    strQry = "SELECT CUSTID,FACISNO FROM " & strOwner & "SO002B" & Get_DB_Link(cn) & _
                    " WHERE MEMBERID='" & varPara(1) & "' AND AUTHENTICID='" & strAuthenticID & "'"

    Set rs = cn.Execute(strQry)
    
    If rs.EOF Then
        JustDoIt = "-2,�q��PPV�{�ҿ��~"
        GoTo 66
    Else
        lngCustID = Val(rs!CustID & "")
        strFaciSno = rs!FaciSno & ""
        ' Else If  [SO002B.FaciSNo] = null  THEN   --2006/07/12 Add
        If strFaciSno = Empty Then
            JustDoIt = "-23,�ӻ{��ID�d�L�{�ҳ]�ƩΪA��"
            GoTo 66
            Exit Function
        End If
    End If
    
    '   ----------- �ª� -----------
    '   Select a.SEQNo, a.FaciSNO, a.SmartCardNo, a.ReFaciSNO, c.Description, a.ModelName, b.FuncFlag01,
    '   a.InstDate, a.PPVright, a.PPVright, a.PgNo from so004 a, cd043 b, cd056 c, cd022 d where
    '   a.CustIt=[nCustId] and a.InitPlaceNo =c.CodeNo and a.ModelCode=b.CodeNo and
    '   a.InstDate is not null and a.PRDate is null and a.FaciCode=d.CodeNo and d.RefNo=3
    '   Order By a.InstDate, a.SEQNo
    
    '   ----------- �s�� -----------
    '   Select a.SEQNo, a.FaciSNO, a.SmartCardNo, a.ReFaciSNO, c.Description, a.ModelName, b.FuncFlag01,
    '   a.InstDate, a.PPVright, a.PPVright, a.PgNo, e.Description from so004 a, cd043 b, cd056 c, cd022 d, cd029 e
    '   where a.CustIt=[nCustId] and a.InitPlaceNo =c.CodeNo and a.ModelCode=b.CodeNo and a.InstDate is not null
    '   and a.PRDate is null and a.FaciCode=d.CodeNo and d.RefNo=3 and a.PgNo=e.CodeNo
    '   Order By a.InstDate, a.SEQNo
    '   ���̪� PPVright �g�F�⦸
    
'   1913
'   Debby
'    �{������GCS_Web.dll�@2005/11/17�@�U��6:16
'    �iAPI-17�j�Ȥ�w��STB�]�Ƹ�Ƭd��(�w�w��STB����,�s���Φ�m)
'    1�D�P�N�Ŵ��ո�Ʊ������\�A���^�����ӼƤ��šG���W���^��12�@�ӡA��ڦ^��11�ӡA�䤤PPVRIGHT���аe2���C�Эק�K
'    (1)�ֵ��`�ص��ŦW�١F���̳W���^CD029���o�æ^�ǡC
'    (2)PPVRIGHT�h�e�@�ӡA���Ǥ�IPPVRIGHT�~��C
'
'    2�DAPI-17�W��O�|���p��CD056�BCD043�A��������F�~����A���]SO004�]�Ƹ���ɦ��Ǹ˸m�I�B�]�ƫ����L�ȡA
'    �H�P�Ӥᦳ�]�Ʀ��|�^���d�L�]�ơA�G���ʳW�欰�K�K
'    �u�n��Ӥ�SO004����쥿�`�ϥΤ����l�A��˸m�I�B�]�ƫ����Y�L�ȡA��ݬ������
    
'                    strOwner & "CD022" & Get_DB_Link(cn) & "D,"

'API -17
'1�D ��^�Ǧ^�Ӫ����e�A��DESCRIPTION���W�٭��СA�нվ�^�Ǧ^�����W�١K
'(1) �˸m�I�G�ثe�^�ǦW�٬�Description ' �קﵹ���O�W�� InitPlace
'(2) �`�ئW�١G�ثe�^�ǦW�٬�Description ' �קﵹ���O�W�� PgName

    strQry = "SELECT A.SEQNO,A.FACISNO,A.SMARTCARDNO,A.REFACISNO,C.DESCRIPTION INITPLACE,A.MODELNAME," & _
                    "B.FUNCFLAG01,A.INSTDATE,A.PPVRIGHT,A.IPPVRIGHT,A.PGNO,E.DESCRIPTION PGNAME " & _
                    "FROM " & _
                    strOwner & "SO004" & Get_DB_Link(cn) & "A," & _
                    strOwner & "CD043" & Get_DB_Link(cn) & "B," & _
                    strOwner & "CD056" & Get_DB_Link(cn) & "C," & _
                    strOwner & "CD029" & Get_DB_Link(cn) & "E " & _
                    " WHERE A.CUSTID=" & lngCustID & _
                    " AND A.INITPLACENO=C.CODENO(+)" & _
                    " AND A.MODELCODE=B.CODENO(+)" & _
                    " AND A.INSTDATE IS NOT NULL AND A.PRDATE IS NULL" & _
                    " AND A.PGNO=E.CODENO(+)" & _
                    " AND A.FACICODE IN (SELECT CODENO FROM " & strOwner & "CD022" & Get_DB_Link(cn) & "WHERE REFNO=3)" & _
                    " AND A.SEQNO IN (" & strFaciSno & ")" & _
                    " ORDER BY A.INSTDATE,A.SEQNO"
    
    '   ----------- �ª� -----------
'    strQry = "SELECT A.SEQNO,A.FACISNO,A.SMARTCARDNO,A.REFACISNO,C.DESCRIPTION,A.MODELNAME," & _
                    "B.FUNCFLAG01,A.INSTDATE,A.PPVRIGHT,A.PPVRIGHT,A.PGNO " & _
                    "FROM " & _
                    strOwner & "SO004" & Get_DB_Link(cn) & "A," & _
                    strOwner & "CD043" & Get_DB_Link(cn) & "B," & _
                    strOwner & "CD056" & Get_DB_Link(cn) & "C " & _
                    " WHERE A.CUSTID=" & lngCustID & _
                    " AND A.INITPLACENO=C.CODENO " & _
                    " AND A.MODELCODE=B.CODENO" & _
                    " AND A.INSTDATE IS NOT NULL AND A.PRDATE IS NULL" & _
                    " AND A.FACICODE IN " & _
                    "(SELECT CODENO FROM " & strOwner & "CD022" & Get_DB_Link(cn) & "WHERE REFNO=3)" & _
                    " ORDER BY A.INSTDATE,A.SEQNO"
    
    '   ----------- �ª� -----------
    '    strQry = "SELECT A.SEQNO,A.FACISNO,A.SMARTCARDNO,A.REFACISNO,C.DESCRIPTION,A.MODELNAME," & _
    '                    "B.FUNCFLAG01,A.INSTDATE,A.PPVRIGHT,A.PPVRIGHT,A.PGNO " & _
    '                    "FROM " & strOwner & "SO004 A," & _
    '                    strOwner & "CD043 B," & _
    '                    strOwner & "CD056 C," & _
    '                    strOwner & "CD022 D " & _
    '                    "WHERE A.CUSTID=" & lngCustId & _
    '                    " AND A.INITPLACENO=C.CODENO " & _
    '                    " AND A.MODELCODE=B.CODENO" & _
    '                    " AND A.INSTDATE IS NOT NULL AND A.PRDATE IS NULL" & _
    '                    " AND A.FACICODE=D.CODENO AND D.REFNO=3" & _
    '                    " ORDER BY A.INSTDATE,A.SEQNO"
    
    Set rs = cn.Execute(strQry)
    
    '    If RecordSet���� > 0 Then
    '      ����XML��.(��select�X�Ӥ����)
    '      �^��(0,���\)��XML��
    '    Else
    '       �^��(-14,�d�L�w��STB�]�Ƹ��)
    '    End If
    
    If rs.RecordCount <= 0 Then
        JustDoIt = "-14,�d�L�w��STB�]�Ƹ��"
    Else
        JustDoIt = "0,���\" & vbCrLf & ExpXML(rs)
    End If
    
    '   PS:
    '   ----------- �ª� -----------
    '    XML(�]�Ƭy�����BSTB�Ǹ��B���z�d�Ǹ��B�����Ǹ��B�˸m�I�B�����W�١B�^�ǼҲաB�w�ˤ���BPPV�ϥ��v�BiPPV�ϥ��v�B�`�ص���)
    '   ----------- �s�� -----------
    '    XML(�]�Ƭy�����BSTB�Ǹ��B���z�d�Ǹ��B�����Ǹ��B�˸m�I�B�����W�١B�^�ǼҲաB�w�ˤ���BPPV�ϥ��v�BiPPV�ϥ��v�B�`�ص��šB�`�ص��ŦW��)
    
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
    ErrHandle "mod_Func17", "JustDoIt"
    JustDoIt = "-99," & strErr
End Function




