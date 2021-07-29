Attribute VB_Name = "mod_Func14"
Option Explicit

Private cn As Object

Public Function JustDoIt(ByRef objWebPool As Object, _
                                        ByRef strPara As String) As String
  On Error GoTo ChkErr
    Dim strQry As String
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
    
    
    ' �p���I�O�`�ب�����]���
    ' 14�B�|���N���B�{��ID
    ' XML(�q�������]�N�X�B�q�������])
    
    varPara = Split(strPara, ",")
    '*****************************************************
    '#3018 �W�[�P�_�ǤJ���ӼƬO�_���T By Kin 2007/09/03
    If UBound(varPara) < 2 Then
        JustDoIt = "-99,�ѼƤ���!"
        GoTo 66
    End If
    '*****************************************************

    
    bytComp = Val(Left(varPara(2), 3))
    strOwner = GetOwner(cn)
    
    ' Old by Lawrence
    ' Select CodeNo,Description from CD015 where StopFlag=0 and Assort in (0,1) and
    '  (ServiceType='C'or ServiceType is null) Order By CodeNo
    
    ' New 2005/11/30 by Liga
    ' API-14�����o�`�حq�檺������]�A�G���O�� CD015.Assort in (0, 4)���~��
    ' Select CodeNo,Description from CD015 where StopFlag=0 and Assort in (0,4) and
    '  (ServiceType='C'or ServiceType is null) Order By CodeNo
    
    strQry = "SELECT CODENO,DESCRIPTION FROM " & strOwner & "CD015" & Get_DB_Link(cn) & _
                    "WHERE STOPFLAG=0" & _
                    " AND ASSORT IN (0,4)" & _
                    " AND SERVICETYPE='C' OR SERVICETYPE IS NULL" & _
                    " ORDER BY CODENO"

    Set rs = cn.Execute(strQry)
    
    ' If RecordSet���� > 0 Then
    '   ����XML��.(��select�X�Ӥ����)
    '   �^��(0,���\)��XML��
    ' Else
    '    �^��(-13,������]�N�XŪ������)
    ' End If
    
    If rs.RecordCount <= 0 Then
        JustDoIt = "-13,������]�N�XŪ������"
    Else
        JustDoIt = "0,���\" & vbCrLf & ExpXML(rs)
    End If
    
    ' 0 : ���\
    ' -13 : ������]�N�XŪ������
    
    ' PS:
    ' XML(�q�������]�N�X�B�q�������])
    
66:
    On Error Resume Next
    Rlx rs
    Rlx strQry
    Rlx varPara
    Rlx strOwner
    Set cn = Nothing
    Set objCnPool = Nothing
  Exit Function
ChkErr:
    ErrHandle "mod_Func14", "JustDoIt"
    JustDoIt = "-99," & strErr
End Function
