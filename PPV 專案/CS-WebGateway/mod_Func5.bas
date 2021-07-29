Attribute VB_Name = "mod_Func5"
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
    
    '   �Ȥ�b�ȸ�Ƭd��(����϶����o�����b���ú�ڰO��)(�ӤH)
    '   5�B�|���N���B�{��ID�B����_��B�������
    
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
    
    '   Select count(*) as Cnt from so002b where MemberId=[�|���N��] and AuthenticId=[�{��ID]
    strQry = "SELECT COUNT(*) AS CNT FROM " & strOwner & "SO002B" & Get_DB_Link(cn) & _
                    "WHERE MEMBERID='" & varPara(1) & "' AND AUTHENTICID='" & strAuthenticID & "'"
    Set rs = cn.Execute(strQry)
    
    '   If Cnt = 0 Then
    '       �^��(-2,�q��PPV�{�ҿ��~)
    '   End If
    
    If rs.RecordCount <= 0 Then
        JustDoIt = "-2,�q��PPV�{�ҿ��~"
        GoTo 66
    End If
    
    '   5�B�|���N���B�{��ID�B����_��B�������

    '   Select ��ڽs���B�����B���O���ئW�١B������B�ꦬ��B�������B�B�ꦬ���B�B
    '   ���O���ơB���İ_�l����B���ĺI�����B���O�H���m�W�B���O�覡�B������]�B
    '   �u����]�B�@�o�B�@�o��]�B�A�����O
    '   from so033�Pso034 where AuthenticId=[�{��ID] and ShouldDate>=[����_��] and
    '   ShouldDate<=[�������] order by ShouldDate Desc, BillNo, ItemNo
    
    '   XML
    '   (��ڽs���B�����B���O���ئW�١B������B�ꦬ��B�������B�B�ꦬ���B�B���O���ơB���İ_�l����B
    '   ���ĺI�����B���O�H���m�W�B���O�覡�B������]�B�u����]�B�@�o�B�@�o��]�B�A�����O)
   
    '   0:  ���\
    '   -6 : �d�L�b�ȸ��
    
    strQry = "SELECT * FROM (" & _
                    "SELECT BILLNO,ITEM,CITEMNAME,SHOULDDATE,REALDATE,SHOULDAMT,REALAMT," & _
                    "REALPERIOD,REALSTARTDATE,REALSTOPDATE,CLCTNAME,CMNAME,UCNAME,STNAME," & _
                    "DECODE(CANCELFLAG,1,'�O','�_') CANCELFLAG,CANCELNAME,SERVICETYPE,GUINO" & _
                    " FROM " & strOwner & "SO033" & Get_DB_Link(cn) & _
                    "WHERE AUTHENTICID='" & strAuthenticID & "'" & _
                    " AND SHOULDDATE >= " & PrcType(varPara(3), FldDate) & _
                    " AND SHOULDDATE < " & PrcType(varPara(4), FldDate) & "+1" & _
                    " UNION ALL " & _
                    "SELECT BILLNO,ITEM,CITEMNAME,SHOULDDATE,REALDATE,SHOULDAMT,REALAMT," & _
                    "REALPERIOD,REALSTARTDATE,REALSTOPDATE,CLCTNAME,CMNAME,UCNAME,STNAME," & _
                    "DECODE(CANCELFLAG,1,'�O','�_') CANCELFLAG,CANCELNAME,SERVICETYPE,GUINO" & _
                    " FROM " & strOwner & "SO034" & Get_DB_Link(cn) & _
                    "WHERE AUTHENTICID='" & strAuthenticID & "'" & _
                    " AND SHOULDDATE >= " & PrcType(varPara(3), FldDate) & _
                    " AND SHOULDDATE < " & PrcType(varPara(4), FldDate) & "+1" & _
                    " ) ORDER BY SHOULDDATE DESC,BILLNO,ITEM"

    Set rs = cn.Execute(strQry)
    
    '   If RecordCount = 0 Then
    '       �^��(-6,�d�L�b�ȸ��)
    '   Else
    '       ����XML��.(��select�X�Ӥ����)
    '       �^��(0,���\)��XML��
    '   End If
    
    If rs.RecordCount <= 0 Then
        JustDoIt = "-6,�d�L�b�ȸ��"
    Else
        JustDoIt = "0,���\" & vbCrLf & ExpXML(rs)
    End If
    
    '   PS:  XML���X����컡��
    '   ���O�覡: CMName
    '   ������]: UCName
    '   �u����]: STName
    '   �@�o: 0���_ ,1���O
    '   �@�o��]: CancelName
   
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
    ErrHandle "mod_Func5", "JustDoIt"
    JustDoIt = "-99," & strErr
End Function



