Attribute VB_Name = "mod_Func19"
Option Explicit

Private cn As Object

Public Function JustDoIt(ByRef objWebPool As Object, _
                                        ByRef strPara As String) As String
  On Error GoTo ChkErr
    Dim rs As Object
    Dim strQry As String
    Dim varPara As Variant
    Dim strOwner As String
    Dim objCnPool As Object
    Dim strAuthenticID As String
    
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
    strOwner = GetOwner(cn)
    
    ' �Ǧ^SMS�|����Ƶ�WEB Update
    ' 19�B�|���N���B�{��ID
    
    ' XML
    ' (�ȪA�����|���N���B�|���W�١B�����Ҧr���B�X�ͤ��-�~�B�X�ͤ��-��B�X�ͤ��-��B
    ' �p���q��(�v) �B�p���q��(��)�B��ʹq�ܡB�ʧO�B�q�T�a�}-�����B�q�T�a�}-�m���ϡB
    ' �q�T�a�}-���q��ѧ˫Ѹ��ӡBEmail�B�ȪA�����n�J�b���BPPV�q�ʱK�X�B�ȪA�����|���K�X�B
    ' �K�X���ܰ��D�B�K�X���ܵ��סB�Ǿ��B�@�N���즳�u�q���A�Ȥ��W�D�`�ص������T���B�B�ê��p�B
    ' �P��H�ơB�P�����B¾�~�O�B�ӤH�~���J�B�����O�B�b�����A�B�ӽФ���ɶ��B�ҥΤ���B
    ' �̫�n�J����B�̫Ყ�ʤ���B�t�Υx�N�X)
    
    ' 2006/04/12 Add -- �o�����Y�B�o�������B�o���a�}�B�Τ@�s��
    
    strQry = "SELECT MEMBERID,NAME,ID,BIRTHDAYYYYY,BIRTHDAYMM,BIRTHDAYDD," & _
                    "TEL,TEL2,MOBILE,SEXTUAL,CITY,TOWN,ADDRESS,EMAIL,LOGINID," & _
                    "PPVORDERPWD,LOGINPASSWORD, HINTNAME,HINTANSWER,EDUCATION," & _
                    "MESSAGE,MARRIED,BODYCOUNT, MEMBERAGE,JOB,RENTE,CLASSID,STATUS," & _
                    "CREATEDATE,USEDATE,LASTLOGINDATE,UPDTIME,COMPCODE," & _
                    "INVTITLE,INVOICETYPE,INVADDRESS,INVNO,FACISNO" & _
                    " FROM " & strOwner & "SO002BLOG" & Get_DB_Link(cn) & _
                    "WHERE MEMBERID='" & varPara(1) & "' AND AUTHENTICID='" & strAuthenticID & "'"

    ' Select * from so002blog where MemberId=[�|���N��] and AuthenticId=[�{��ID]
    Set rs = cn.Execute(strQry)
    
    ' If RecordSet���� > 0 Then
    '   ����XML��.(��select�X�Ӥ����)
    '   �^��(0,���\)��XML��
    ' Else
    '   �^��(-16,�d�L�Ǧ^SMS�|�����)
    ' End If
    
    ' 0:  ���\
    ' -16 : �d�L�Ǧ^SMS�|�����
    
    ' -16, �ӷ|����Ʃ�SMS�L����
    If rs.RecordCount <= 0 Then
        JustDoIt = "-16,�ӷ|����Ʃ�SMS�L����"
'        JustDoIt = "-16,�d�L�Ǧ^SMS�|�����"
    Else
        JustDoIt = "0,���\" & vbCrLf & ExpXML(rs)
    End If
    
66:
    On Error Resume Next
    Rlx rs
    Rlx strQry
    Rlx strOwner
    Set cn = Nothing
    Set objCnPool = Nothing
  Exit Function
ChkErr:
    ErrHandle "mod_Func19", "JustDoIt"
    JustDoIt = "-99," & strErr
End Function
