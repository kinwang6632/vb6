Attribute VB_Name = "mod_Func6"
Option Explicit

Private strAuthenticID As String
Private cn As Object

Public Function JustDoIt(ByRef objWebPool As Object, _
                                        ByRef strPara As String) As String
  On Error GoTo ChkErr
    Dim strQry As String, strSubQry1 As String, strSubQry2 As String
    Dim varPara As Variant
    Dim rs As Object
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
    
    '   �Ȥᦳ�u�q�������O��b�@�~/�p���I�O�w�I���B��b�@�~(�H�Υd�p�B�I�ھ���)--�u�W�I�ڱ��v--���d��
    '   6�B�|���N���B�{��ID
    '   XML(��ڽs���B�����B���O���ئW�١B��������B�������B�B���O���ơB���İ_�l����B���ĺI�����B������]�B�A�����O);
    '   �L�o�����B�����w�I�ڤ���Ƥνu�W�I�ڽб��v�q�L����
    '   0:  ���\
    '   -6 : �d�L�b�ȸ��
    
    varPara = Split(strPara, ",")
    '*****************************************************
    '#3018 �W�[�P�_�ǤJ���ӼƬO�_���T By Kin 2007/08/30
    If UBound(varPara) <> 2 Then
        JustDoIt = "-99,�ѼƤ���!"
        GoTo 66
    End If
    '*****************************************************

    strAuthenticID = varPara(2)
    bytComp = Val(Left(strAuthenticID, 3))
    strOwner = GetOwner(cn)
    
    '    Select count(*) as Cnt from so002b where MemberId=[�|���N��] and AuthenticId=[�{��ID]
    '    Select FaciSNo from so002b where MemberId=[�|���N��] and AuthenticId=[�{��ID]
    strQry = "SELECT COUNT(*) AS CNT FROM " & strOwner & "SO002B" & Get_DB_Link(cn) & _
                    "WHERE MEMBERID='" & varPara(1) & "' AND AUTHENTICID='" & strAuthenticID & "'"
    Set rs = cn.Execute(strQry)
    
    '    If Cnt = 0 Then
    '      �^��(-2,�q��PPV�{�ҿ��~)
    '    End If
    
'    If rs.RecordCount <= 0 Then
    If rs("CNT") <= 0 Then
        JustDoIt = "-2,�q��PPV�{�ҿ��~"
        GoTo 66
    End If
    
    '    �� :
    '    Select ��ڽs���B�����B���O���ئW�١B��������B�������B�B���O���ơB���İ_�l����B���ĺI�����B������]�B�A�����O
    '    from so033 where AuthenticId=[�{��ID] and ����(���ȥB�ѦҸ�<>3)
    '    �B�����w�I�ڤ����(AccountNo is null) and realdate is null
    '    order by ShouldDate Desc, BillNo, ItemNo
    
    '   �s :
    '   Select��ڽs���B�����B���O���ئW�١B��������B�������B�B���O���ơB���İ_�l����B���ĺI�����B������]�B�A�����O
    '   from so033 where FaciSeqNo in [FaciSNo] and ����(���ȥB�ѦҸ�<>3)
    '   �B�����w�I�ڤ����(AccountNo is null) and realdate is null
    '   order by ShouldDate Desc, BillNo, ItemNo ;
    
    strSubQry1 = "SELECT FACISNO FROM " & strOwner & "SO002B" & Get_DB_Link(cn) & _
                            " WHERE MEMBERID='" & varPara(1) & "' AND AUTHENTICID='" & strAuthenticID & "'" & _
                            " AND FACISNO IS NOT NULL"
                            
    On Error Resume Next
    strSubQry1 = RPxx(cn.Execute(strSubQry1).GetString(2, , "", ",", ""))
    If Err = 3021 Then
        Err.Clear
        strSubQry1 = ""
    End If
    
    On Error GoTo ChkErr
    If Len(strSubQry1) > 0 Then
        strSubQry2 = "SELECT CODENO FROM " & strOwner & "CD013" & Get_DB_Link(cn) & " WHERE NVL(REFNO,0) <> 3"
        strQry = "SELECT BILLNO,ITEM,CITEMNAME,SHOULDDATE,SHOULDAMT,REALPERIOD,REALSTARTDATE" & _
                        ",REALSTOPDATE,UCNAME,SERVICETYPE FROM " & strOwner & "SO033" & Get_DB_Link(cn) & _
                        " WHERE FACISEQNO IN (" & Left(strSubQry1, Len(strSubQry1) - 1) & ")" & _
                        " AND UCCODE IN (" & strSubQry2 & ")" & _
                        " AND ACCOUNTNO IS NULL" & _
                        " AND REALDATE IS NULL" & _
                        " ORDER BY SHOULDDATE DESC,BILLNO,ITEM"
        Set rs = cn.Execute(strQry)
        If rs.RecordCount <= 0 Then
            JustDoIt = "-6,�d�L�b�ȸ��"
        Else
            JustDoIt = "0,���\" & vbCrLf & ExpXML(rs)
        End If
    Else
'        JustDoIt = "-99,[�Ȥ�{�Ҹ�T��] �d�L [�{�ҳ]�ƧǸ�] ���!"
        JustDoIt = "-23,�ӻ{��ID�d�L�{�ҳ]�ƩΪA��"
    End If
    
    '   If RecordCount = 0 Then
    '       �^��(-6,�d�L�b�ȸ��)
    '   Else
    '       ����XML��.(��select�X�Ӥ����)
    '       �^��(0,���\)��XML��
    '   End If
    
66:
    On Error Resume Next
    Rlx rs
    Rlx strQry
    Rlx strSubQry1
    Rlx strSubQry2
    Rlx varPara
    Rlx strOwner
    Rlx strAuthenticID
    Set cn = Nothing
    Set objCnPool = Nothing
  Exit Function
ChkErr:
    ErrHandle "mod_Func6", "JustDoIt"
    JustDoIt = "-99," & strErr
End Function





