Attribute VB_Name = "mod_Func18"
Option Explicit

Private cn As Object
Private strAuthenticID As String

Public Function JustDoIt(ByRef objWebPool As Object, _
                                        ByRef strPara As String) As String
  On Error GoTo ChkErr
    Dim strQry As String
    Dim rs As Object
    Dim objCnPool As Object
    Dim strOwner As String
    Dim varPara As Variant

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
    
    strOwner = GetOwner(cn)
    
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
    
    '18�B�|���N���B�{��ID
    
    'Ū���I�Ʀ��O���
    '18
    'XML(���O���إN�X�B���O���ئW��)
    
    'Select CodeNo, Description, Amount, CreditPoint from cd019 where StopFlag=0 and
    'RefNo in (4,17) Order By CodeNo
    
    strQry = "SELECT CODENO,DESCRIPTION,AMOUNT,CREDITPOINT FROM " & strOwner & "CD019" & _
                    Get_DB_Link(cn) & " WHERE STOPFLAG <> 1 AND REFNO IN (4,17) ORDER BY CODENO"

    Set rs = cn.Execute(strQry)
    
    'If RecordSet���� > 0 Then
    '  ����XML��.(��select�X�Ӥ����)
    '  �^��(0,���\)��XML��
    'Else
    '   �^��(-15,�d�L�I�Ʀ��O���)
    'End If
    
    'PS:
    'XML(���O���إN�X�B���O���ئW�١B���B�B�I��)
    
    If rs.RecordCount <= 0 Then
        JustDoIt = "-15,�d�L�I�Ʀ��O���"
    Else
        JustDoIt = "0,���\" & vbCrLf & ExpXML(rs)
    End If
    
66:
    On Error Resume Next
    Rlx rs
    Rlx strQry
    Rlx varPara
    Rlx strOwner
    Rlx strAuthenticID
    Set cn = Nothing
    Set objCnPool = Nothing
  Exit Function
ChkErr:
    ErrHandle "mod_Func18", "JustDoIt"
    JustDoIt = "-99," & strErr
End Function
