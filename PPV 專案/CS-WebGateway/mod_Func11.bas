Attribute VB_Name = "mod_Func11"
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
    Dim objCnPool As Object
    Dim strOwner As String
    Dim objIPPV As Object
    Dim strPinCode As String
    
'    Set objCnPool = objWebPool
'    Set cn = objCnPool.GetConnection
'
'    If cn.State <= 0 Then
'        Set cn = objCnPool.GetConnection
'        If cn.State <= 0 Then
'            strErr = "�L�k�s�u��Ʈw!"
'            JustDoIt = "-99,��Ʈw�s�u����!"
'            GoTo 66
'        End If
'    End If
    
    varPara = Split(strPara, ",")
    '*****************************************************
    '#3018 �W�[�P�_�ǤJ���ӼƬO�_���T By Kin 2007/08/30
    If UBound(varPara) < 6 Then
        JustDoIt = "-99,�ѼƤ���!"
        GoTo 66
    End If
    '*****************************************************
    
    strAuthenticID = varPara(2)
    bytComp = Val(Left(strAuthenticID, 3))
    lngCustID = Val(Right(strAuthenticID, 7))
    Set cn = Get_General_CN2(bytComp)
    If cn.State <= 0 Then
        Set cn = Get_General_CN2(bytComp)
        If cn.State <= 0 Then
            strErr = "�L�k�s�u��Ʈw!"
            JustDoIt = "-99,��Ʈw�s�u����!"
            GoTo 66
        End If
    End If

    strOwner = GetOwner(cn)
    strPinCode = varPara(6)
'    If strPinCode = Empty Then strPinCode = "1234"
    If strPinCode = Empty Then strPinCode = ""
    
    '   Select count(*) as Cnt from so002b where MemberId=[�|���N��] and AuthenticId=[�{��ID]
    '   If Cnt = 0 Then
    '       �^��(-2,�q��PPV�{�ҿ��~)
    '   End If
    
'    strQry = "SELECT COUNT(*) AS RCDCNT FROM " & strOwner & "SO002B" & Get_DB_Link(cn) & _
'                    "WHERE MEMBERID='" & varPara(1) & "' AND AUTHENTICID='" & strAuthenticID & "'"
'
'
'    If Val(GetValue(cn, strQry)) = 0 Then
'        JustDoIt = "-2,�q��PPV�{�ҿ��~"
'        GoTo 66
'    End If
    
    strQry = "SELECT CUSTID FROM " & strOwner & "SO002B" & Get_DB_Link(cn) & _
                    "WHERE MEMBERID='" & varPara(1) & "' AND AUTHENTICID='" & strAuthenticID & "'"
    
    lngCustID = Val(GetValue(cn, strQry) & "")
    
    If lngCustID = 0 Then
        JustDoIt = "-2,�q��PPV�{�ҿ��~"
        GoTo 66
    End If
    
    Set objIPPV = CreateObject("setIPPV.clsNagraSTB")
    With objIPPV
            Set .uConn = Get_General_CN
            garyGi(1) = "WEB"
            .ugaryGi = Join(garyGi, Chr(9))
    End With
    
    '   11�B�|���N���B�{��ID�B�]�Ƭy�����BSTB�Ǹ��BICC�Ǹ��BiPPV PIN Code(4�X)
    '   0:  ���\
    '   -9 : �K�X�ܧ󥢱�
    
    '   �̶ǤJ���Ѽ�(STB�Ǹ��BICC�Ǹ��BiPPV PIN Code)�I�sCA�R�O[�ܧ�iPPV�q��PIN Code]�D��ı�ƼҲ�. �����N��E8
    '   If �R�O���\ Then
    '       �^��(0,���\)
    '   Else
    '       �^��(-9,�K�X�ܧ󥢱�)
    '   End If
    
    '   Old
    '   PS : iPPV PIN Code�|�X. �p�GiPPV PIN Code�ǤJnull��(�Y�L��),
    '   �h�I�sCA�R�O[�ܧ�iPPV�q��PIN Code]�D��ı�ƼҲծ�, PIN Code�����жǤJ'1234'
    
    '   New
    '   PS : iPPV PIN Code�|�X. �p�GiPPV PIN Code�ǤJnull��(�Y�L��),
    '   �h�I�sCA�R�O[�ܧ�iPPV�q��PIN Code]�D��ı�ƼҲծ�, PIN Code�����жǤJ'0000'
    
    With objIPPV
    '    Public Function SetResetiPPVwd(ByVal strCompCode As String, _
    '        ByVal lngCustId As Long, ByVal strSmartCard As String, _
    '        ByVal strResvTime As String, ByVal strSTB As String, _
    '        ByVal strDoCaption As String, ByVal strPinCode As String) As Boolean
    '       Format(Now, "YYYY/MM/DD HH:MM")
        If Not .SetResetiPPVwd(CStr(bytComp), _
                                                lngCustID, _
                                                CStr(varPara(5)), _
                                                "", _
                                                CStr(varPara(4)), _
                                                "", _
                                                strPinCode) Then
            strErr = .uMsgBack
            JustDoIt = "-9,�K�X�ܧ󥢱�"
        Else
            JustDoIt = "0,���\"
        End If
    End With
        
66:
    On Error Resume Next
    Rlx rs
    Rlx strQry
    Rlx varPara
    Rlx objIPPV
    Rlx strOwner
    Rlx lngCustID
    Rlx strPinCode
    Rlx strAuthenticID
    Set cn = Nothing
    Set objCnPool = Nothing
  Exit Function
ChkErr:
    ErrHandle "mod_Func11", "JustDoIt"
    JustDoIt = "-99," & strErr
End Function


