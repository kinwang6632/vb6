Attribute VB_Name = "mod_Func20"
Option Explicit

Private cn As Object
Private strAuthenticID As String

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
    
    ' ��O�Ȥ�W��d�� ( �R�O 20 )
    '   20�B�|���N���B�{��ID
    '   0:  ���\
    '   -17 : �d�L��O���
    
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
    
    'select BillNo, Item, CitemName, ShouldDate, ShouldAmt, RealPeriod, RealStartDate, RealStopDate,
    ' UCName, ServiceType from so033 where AuthenticId=[�{��ID] and
    ' UCCode is not null and CancelFlag=0 Order by BillNo, Item, ServiceType

    strQry = "SELECT BILLNO,ITEM,CITEMNAME,SHOULDDATE,SHOULDAMT,REALPERIOD,REALSTARTDATE," & _
                    "REALSTOPDATE,UCNAME,SERVICETYPE FROM " & strOwner & "SO033" & Get_DB_Link(cn) & _
                    " WHERE AUTHENTICID='" & strAuthenticID & "' AND" & _
                    " UCCODE IS NOT NULL AND CANCELFLAG=0" & _
                    " ORDER BY BILLNO,ITEM,SERVICETYPE"
    
    Set rs = cn.Execute(strQry)
    
    'If RecordSet���� > 0 Then
    '  ����XML��.(��select�X�Ӥ����)
    '  �^��(0,���\)��XML��
    'Else
    '   �^��(-17,�d�L��O���)
    'End If
    
    If rs.RecordCount <= 0 Then
        JustDoIt = "-17,�d�L��O���"
    Else
        JustDoIt = "0,���\" & vbCrLf & ExpXML(rs)
    End If
    
    'PS:
    'XML(��ڽs���B�����B���O���ئW�١B��������B�������B�B���O���ơB
    ' ���İ_�l����B���ĺI�����B������]�B�A�����O);�������
   
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
    ErrHandle "mod_Func20", "JustDoIt"
    JustDoIt = "-99," & strErr
End Function

