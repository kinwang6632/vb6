Attribute VB_Name = "mod_Func10"

Option Explicit

Private strAuthenticID As String
Private cn As Object

Public Function JustDoIt(ByRef objWebPool As Object, _
                                        ByRef strPara As String) As String
  On Error GoTo ChkErr
    Dim strQry As String
    Dim varPara As Variant
    Dim rs As Object
    Dim objIPPV As Object
    Dim lngCustID As Long
    Dim strOwner As String
    Dim objCnPool As Object
    
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
    
    '�Ȥ��ܧ�STB�|�X�ƦrParental PIN Code
    '10�B�|���N���B�{��ID�B�]�Ƭy�����BSTB�Ǹ��BICC�Ǹ�
    '0:  ���\
    '-9 : �K�X�ܧ󥢱�
    
    varPara = Split(strPara, ",")
    '*****************************************************
    '#3018 �W�[�P�_�ǤJ���ӼƬO�_���T By Kin 2007/08/30
    If UBound(varPara) <> 5 Then
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

'    Select count(*) as Cnt from so002b where MemberId=[�|���N��] and AuthenticId=[�{��ID]
'    If Cnt = 0 Then
'      �^��(-2,�q��PPV�{�ҿ��~)
'    End If
    
'    strQry = "SELECT COUNT(*) AS RCDCNT FROM " & strOwner & "SO002B" & Get_DB_Link(cn) & _
'                " WHERE MEMBERID='" & varPara(1) & "' AND AUTHENTICID='" & strAuthenticID & "'"
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
    
'    �̶ǤJ���Ѽ�(STB�Ǹ��BICC�Ǹ�)�I�sCA�R�O[�ܧ�Parental PIN Code]�D��ı�ƼҲ�. �����N��E1
'    If �R�O���\ Then
'       �^��(0,���\)
'    Else
'       �^��(-9,�K�X�ܧ󥢱�)
'    End If
    
    Set objIPPV = CreateObject("setIPPV.clsNagraSTB")
    With objIPPV
            Set .uConn = Get_General_CN
            garyGi(1) = "WEB"
            .ugaryGi = Join(garyGi, Chr(9))
    End With
    
    With objIPPV
        '    Public Function SetPinCode(ByVal strCompCode As String, _
        '        ByVal lngCustId As Long, ByVal strSmartCard As String, _
        '        ByVal strResvTime As String, ByVal strSTB As String, _
        '        ByVal strDoCaption As String, ByVal strPinCode As String) As Boolean
        If Not .SetPinCode(CStr(bytComp), _
                                        lngCustID, _
                                        CStr(varPara(5)), _
                                        "", _
                                        CStr(varPara(4)), _
                                        "", _
                                        "") Then
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
    Rlx strAuthenticID
    Set cn = Nothing
    Set objCnPool = Nothing
  Exit Function
ChkErr:
    ErrHandle "mod_Func10", "JustDoIt"
    JustDoIt = "-99," & strErr
End Function
