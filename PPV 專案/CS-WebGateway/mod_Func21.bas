Attribute VB_Name = "mod_Func21"
Option Explicit

Private cn As Object

Public Function JustDoIt(ByRef objWebPool As Object, _
                                        ByRef strPara As String) As String
  On Error GoTo ChkErr
    Dim varAry As Variant
    Dim varPara As Variant
    Dim varElement As Variant
    Dim strCustID As String
    Dim objCnPool As Object
    Dim strCompCode As String
    
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
    
    ' 21
    ' �g��ϴM��
    
    ' �ǤJ�Ѽ�[21�B(�t�Υx�O�ƿ�)�B�����B�m���B�����ܮH�����L�B�F�B�q�B�ѡB�ˡB���B���B��L]
    
    ' 1.  ��Ѩt�Υx�O�ƿ鷺�e
    ' 2.  �N�ǤJ�a�}�U�ѼƦ��Ȫ��ܼƦ�_��
    ' 3.  �̩�ѥX���t�Υx�O�Ӧ�Owner�ôM���Owner��SO001.InstAddress
    
    ' Loop �t�Υx�O(�h��)
    '   ��LIKE�覡���j�M�a�}��k(SO001.InstAddress)
    '   �B�H�Ӧa�}�U�Ȥᤧ�̫�˾�����������Ȥ᪬�A�̾�
    '   (�ǤJ�ѼƤ�[����]���Ȯ�, ���H����a�}�d��;
    '   �Y�䤣���, �Хh�����������b���d�ߤ@��.
    '   �Y�J�ѼƤ�[����]�L�Ȯ�,�h�������d��)
    
    '   IF DATA FOUND THEN
    '     RETURN CompName(SO001), CustStatusName(SO002)
    '     EXIT
    '   End If
    ' End Loop
    
    ' If LOOP���ӧ䤣��� Then
    '    �^�� -18�P"�d�L�g��ϸ��! �i�ण�b�F�˸g��ϩΥثe���t�u�ܸӦa�}"
    ' End If
    
    varPara = Split(strPara, ",")
    '*****************************************************
    '#3018 �W�[�P�_�ǤJ���ӼƬO�_���T By Kin 2007/09/03
    If UBound(varPara) < 2 Then
        JustDoIt = "-99,�ѼƤ���!"
        GoTo 66
    End If
    '*****************************************************

    
    strCompCode = varPara(1)
    
    ' �t�Υx�O�ƿ�榡����:
    ' �ǤJ���t�Υx�O�N�X���j�Ÿ��� [Tab]
    ' Ex:
    ' 21,1    2,�̪F��,�E�p�m,�K..
    
    ' �Ƕi�Ѽ� : 21�B(�t�Υx�O�ƿ�)�B�����B�m���B�����ܮH�����L�B�F�B�q�B�ѡB�ˡB���B���B��L
    If InStr(1, strCompCode, vbTab) Then
        varAry = Split(strCompCode, vbTab)
        For Each varElement In varAry
            If varElement <> Empty Then
                If IsNumeric(varElement) Then
                    If QueryGo(CStr(varElement), GetAddress(varPara, 2), JustDoIt) Then Exit For
                    If QueryGo(CStr(varElement), GetAddress(varPara, 3), JustDoIt) Then Exit For
                Else
                    JustDoIt = "-99,[�t�Υx] �Ѽƿ��~!"
                    Exit For
                End If
            End If
        Next
    Else
        If IsNumeric(strCompCode) Then
            If Not QueryGo(strCompCode, GetAddress(varPara, 2), JustDoIt) Then
                QueryGo strCompCode, GetAddress(varPara, 3), JustDoIt
            End If
        Else
            JustDoIt = "-99,[�t�Υx] �Ѽƿ��~!"
        End If
    End If
    
    ' 0:  ���\
    ' -18 : �d�L�g��ϸ��! �i�ण�b�F�˸g��ϩΥثe���t�u�ܸӦa�}
    
    ' ���G�X�B���G�T���r��(���)
    ' �t�Υx�W�١B�Ȥ᪬�A
    
    ' PS:
    ' a.�Ȥ᪬�A�^�ǭȬ��̫�˾�����Ȥᤧ���A
    ' b.�HHomePass��g���
    
66:
    On Error Resume Next
    Rlx varAry
    Rlx varPara
    Rlx varElement
    Rlx strCustID
    Rlx objCnPool
    Rlx strCompCode
    Set cn = Nothing
    Set objCnPool = Nothing
  Exit Function
ChkErr:
    ErrHandle "mod_Func21", "JustDoIt"
    JustDoIt = "-99," & strErr
End Function

Private Function GetAddress(varAry As Variant, intStart As Integer) As String
  On Error GoTo ChkErr
    Dim intLoop As Integer
    For intLoop = intStart To UBound(varAry)
        GetAddress = GetAddress & varAry(intLoop)
    Next
  Exit Function
ChkErr:
    ErrHandle "mod_Func21", "GetAddress"
End Function

Private Function QueryGo(strCompCode As String, _
                                            strAddress As String, _
                                            strResult As String) As Boolean
  On Error GoTo ChkErr
    Dim strOwner As String
    Dim strCustID As String
    bytComp = CByte(RPxx(strCompCode))
    strOwner = GetOwner(cn)
    strCustID = Search_Cust(strAddress, strOwner)
    QueryGo = False
    If strCustID <> Empty Then
        strResult = Get_Cust_Status(strCustID, strOwner)
        If strResult = Empty Then
            strResult = "-18,�d�L�g��ϸ��! �i�ण�b�F�˸g��ϩΥثe���t�u�ܸӦa�}"
        Else
            strResult = "0,���\" & vbCrLf & Get_Sys_Platform(CInt(strCompCode)) & "," & strResult
            QueryGo = True
        End If
    Else
        strResult = "-18,�d�L�g��ϸ��! �i�ण�b�F�˸g��ϩΥثe���t�u�ܸӦa�}"
    End If
    On Error Resume Next
    Rlx strOwner
    Rlx strCustID
  Exit Function
ChkErr:
    ErrHandle "mod_Func21", "QueryGo"
End Function

Private Function Get_Cust_Status(strCustID As String, strOwner As String) As String
  On Error GoTo ChkErr
    Get_Cust_Status = GetValue(cn, "SELECT CUSTSTATUSNAME FROM " & strOwner & "SO002" & Get_DB_Link(cn) & _
                                                        " WHERE CUSTID =" & strCustID & " AND INSTTIME=" & _
                                                        "(SELECT MAX(INSTTIME) FROM " & strOwner & "SO002" & Get_DB_Link(cn) & _
                                                        " WHERE CUSTID=" & strCustID & ")") & Empty
    '   a.�Ȥ᪬�A�^�ǭȬ��̫�˾�����Ȥᤧ���A
  
  Exit Function
ChkErr:
    ErrHandle "mod_Func21", "Get_Cust_Status"
End Function

Private Function Search_Cust(strAddress As String, strOwner As String) As String
  On Error GoTo ChkErr
 '   Search_Cust = GetValue(cn, "SELECT CUSTID FROM " & strOwner & _
                                                    "SO001 WHERE INSTADDRESS = '" & strAddress & "'") & Empty
    Search_Cust = GetValue(cn, "SELECT CUSTID FROM " & strOwner & "SO001" & Get_DB_Link(cn) & _
                                                    "WHERE INSTADDRESS LIKE '" & strAddress & "%'") & Empty
    '   b.�HHomePass��g���
  Exit Function
ChkErr:
    ErrHandle "mod_Func21", "Search_Cust"
End Function

' �t�Υx�O����:
' �[�@=1�B�̫n=2�B�n��=3�B�s�W�D=5�B�׷�=6�B���D=7�B���p=8�B�����s=9�B�s�x�_=10�B
' ���W�D=11�B�j�w��s=12�B�s��=13�B�s��=14�B�_���=16

Private Function Get_Sys_Platform(intSysID As Integer) As String
  On Error GoTo ChkErr
    Get_Sys_Platform = Choose(intSysID, "�[�@", "�̫n", "�n��", "", "�s�W�D", "�׷�", "���D", _
                                                    "���p", "�����s", "�s�x�_", "���W�D", "�j�w��s", "�s��", "�s��", "", "�_���")

  Exit Function
ChkErr:
    ErrHandle "mod_Func21", "Get_Sys_Platform"
End Function

