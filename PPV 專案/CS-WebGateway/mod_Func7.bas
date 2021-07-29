Attribute VB_Name = "mod_Func7"
Option Explicit

Private cn As Object
Private rs As Object
Private strOwner As String

Public Function JustDoIt(ByRef objWebPool As Object, _
                                        ByRef strPara As String) As String
  On Error GoTo ChkErr
  
    Dim objPayment As Object
    
    Dim varPara As Variant
    Dim varDetail As Variant
    
    Dim strAuthenticID As String
    Dim strAuthenticCode As String
    Dim strServiceType As String
    Dim strBillData As String
'    Dim lngCustID As Long
    
    Dim lngDepositFlag As Long
    Dim lngCardCode As Long
    Dim sglTotalAmt As Single
    Dim strCardBillNo As String
    Dim strCardExpDate As String
    Dim strAccountNo As String
    
    Dim strRealDate As String
    Dim strCMCode As String
    Dim strCMName As String
    Dim strPTCode As String
    Dim strPTName As String
    Dim strClctEn As String
    Dim strClctName As String
    Dim strUpdTime As String
    Dim strUpdEn As String
    Dim strEntryEn As String
    Dim lngReturn As Long
   
    Dim intLoop As Integer
    
    Set cn = Get_General_CN
    
    If cn.State <= 0 Then
        Set cn = Get_General_CN
        If cn.State <= 0 Then
            strErr = "�L�k�s�u��Ʈw!"
            JustDoIt = "-99,��Ʈw�s�u����!"
            GoTo 66
        End If
    End If
    
'   7�B�|���N���B�{��ID�B�d���B�H�Υd���Ĵ����B�H�Υd�I�����v�X�B�H�Υd�O(���)
'   ��ڽs���B����(���)
'   ��ڽs���B����(���)
'   �K.

    varPara = Split(strPara, vbCrLf)
    '***********************************************************************************
    '#3018 �P�_�ǤJ���ӼƬO�_���T By Kin 2007/08/31
    If UBound(varPara) < 2 Then
        JustDoIt = "-99,�ѼƸ�Ƥ����T ! �нT�{ !"
        GoTo 66
    End If
    
    varDetail = Split(varPara(0), ",")
    If UBound(varDetail) < 6 Then
        JustDoIt = "-99,�ѼƤ���!"
        GoTo 66
    End If
    For intLoop = 1 To UBound(varPara)
        If (UBound(Split(varPara(intLoop), ",")) <> 1) And (varPara(intLoop) <> Empty) Then
            JustDoIt = "-99,�ѼƤ���!"
            GoTo 66
        End If
    Next intLoop
    '***********************************************************************************
    strAuthenticID = varDetail(2)
    bytComp = Val(Left(strAuthenticID, 3)) '   ���q�O�N�X: SO033.CompCode
'    lngCustID = Val(Right(strAuthenticID, 7))
    
    strOwner = GetOwner(cn)

'   Select count(*) as Cnt from so002b where MemberId=[�|���N��] and AuthenticId=[�{��ID]
    Set rs = cn.Execute("SELECT COUNT(*) AS CNT FROM " & strOwner & "SO002B" & Get_DB_Link(cn) & _
                                    "WHERE MEMBERID='" & varDetail(1) & "' AND AUTHENTICID='" & strAuthenticID & "'")
    
    If rs("CNT") <= 0 Then '    If Cnt = 0 Then
        JustDoIt = "-2,�q��PPV�{�ҿ��~" '        �^��(-2,�q��PPV�{�ҿ��~)
        GoTo 66
    End If '    End If
    
'   strCardBillNo   �H�Υd�q��s��(��API�HSequence Object����)
    strCardBillNo = GetValue(cn, "SELECT " & strOwner & "S_SO033_CARDBILLNO.NEXTVAL FROM DUAL") & ""
    strCardBillNo = Trim(strCardBillNo)
    
    If strCardBillNo = Empty Then
        JustDoIt = "-99,[S_SO033_CARDBILLNO] Sequence �`�Ǫ��󤣦s�b !"
        GoTo 66
    End If
    
    strCardBillNo = "2" & Right("00000000" & strCardBillNo, 8)
    '   ���ͫH�Υd�q��渹(9�X). �X��h :
    '       ��1�X���ӷ��ѧO�X. [2]��WEB�I�s;
    '       ��2�X�}�l���y����. �I�s Sequence Object S_SO033_CardBillNo ���o�̷s��. 8�X�k�a���ɹs.
    '       Ex : �ӷ��ѧO�X + �y���� �� 200000001
    
    strAccountNo = varDetail(3) '   strAccountNo �H�Υd�d��(WEB�ǤJ)
    strCardExpDate = varDetail(4) & "" '   strCardExpDate �H�Υd���Ħ~��(WEB�ǤJ)
    strAuthenticCode = varDetail(5)
    lngCardCode = Val(varDetail(6) & "") '   lngCardCode �H�Υd�O(WEB�ǤJ)
'   lngDepositFlag = -1
    sglTotalAmt = 0
    
'   �̶ǨӰѼƩI�sPayment Gateway �Ҳդ������v(Approve)�ýд�(Deposit)���G�X�@API

    cn.BeginTrans
    
    For intLoop = 1 To UBound(varPara)
        
        strBillData = varPara(intLoop) & ""
        
        If strBillData <> Empty Then
            varDetail = Split(strBillData, ",")
        '    ��ڽs���B����(���)
        '    ��ڽs���B����(���)
        '    �K.
            If GetRS(cn, rs, "SELECT CARDBILLNO,SHOULDAMT,SERVICETYPE FROM " & _
                                        strOwner & "SO033" & Get_DB_Link(cn) & _
                                        " WHERE BILLNO='" & Trim(varDetail(0)) & "'" & _
                                        " AND ITEM=" & Trim(varDetail(1))) Then
                                            
                If rs.AbsolutePosition > 0 Then
                    rs!CardBillNo = strCardBillNo
                    rs.Update
                    sglTotalAmt = sglTotalAmt + Val(rs!ShouldAmt & "") '   ���v���B: ��(��ڽs���B����) Ū�� SO033 ���[�`����
                    If strServiceType = Empty Then strServiceType = rs!ServiceType & "" '   lngAmount ���v���B(��APIŪ����ƥ[�`)
                End If
                
            End If
            
        End If
        
    Next
    
    '   PS: �`�p���B�Х�Web�ݦۦ�[�`�B�z
    
    lngDepositFlag = Val(GetValue(cn, "SELECT PARA32 FROM " & strOwner & "SO043" & Get_DB_Link(cn) & _
                                                            " WHERE COMPCODE=" & bytComp & _
                                                            " AND SERVICETYPE='" & strServiceType & "'") & "")
    
    Set objPayment = CreateObject("csPaymentClass.Approve")
    
    Get_Def_CM cn, strOwner, strCMCode, strCMName '    strCMcode = "�w�]���O�覡�N�X" , strCMname = "�w�]���O�覡�W��"
    Get_Def_PT cn, strOwner, strPTCode, strPTName '    strPTcode = "�w�]�I�ں����N�X" , strPTname = "�w�]�I�ں����W��"
    strRealDate = "" ' ���ڤ��(�ťեH�t�Τ�����D) Null
    strClctEn = "WEB" '   ���O�H���N�X: 'WEB'
    strClctName = "����" '   ���O�H���W��: '����'
    strUpdTime = "" ' ���ʮɶ� Null
    strUpdEn = "" ' ���ʤH�� Null
    strEntryEn = "WEB" ' �ާ@�̥N����J 'WEB'
    '#3599 �h�W�[�@�� strAuthenticID�Ѽ�
    With objPayment '   �I�sPayment Gateway �Ҳդ������v(Approve)�ýд�(Deposit)���G�X�@API
            Set .objConn = cn
            .uOwnerName = strOwner
            lngReturn = .Approve(Val(bytComp), strCardBillNo, lngCardCode, strAccountNo, _
                                    strCardExpDate, Val(sglTotalAmt), strRealDate, strCMCode, _
                                    strCMName, strPTCode, strPTName, strClctEn, strClctName, _
                                    strUpdTime, strUpdEn, strEntryEn, lngDepositFlag, strAuthenticCode)
            
            If lngReturn = 0 Then
                cn.CommitTrans
                JustDoIt = "0,���\," & strCardBillNo '       �^��(0,���\,�H�Υd�q��s��)
            Else '    If ���v���G��ok Then
                JustDoIt = "-7,�H�α��v���~" '       �^��(-7,�H�α��v���~)
                cn.RollbackTrans
                JustDoIt = JustDoIt & " " & CStr(lngReturn) & " " & .getErrorCode & " " & .getErrorMsg
            End If
    End With
    
'    ���G�X�B���G�T���r��(���)
'    �H�Υd�q��s��
  
66:
    On Error Resume Next
    Rlx rs
    Rlx cn
    Rlx objPayment
    Rlx varPara
    Rlx varDetail
    Rlx strAuthenticID
    Rlx strAuthenticCode
    Rlx strServiceType
    Rlx strBillData
    Rlx lngDepositFlag
    Rlx lngCardCode
    Rlx sglTotalAmt
    Rlx strCardBillNo
    Rlx strCardExpDate
    Rlx strAccountNo
    Rlx strRealDate
    Rlx strCMCode
    Rlx strCMName
    Rlx strPTCode
    Rlx strPTName
    Rlx strClctEn
    Rlx strClctName
    Rlx strUpdTime
    Rlx strUpdEn
    Rlx strEntryEn
    Rlx intLoop
    Rlx strOwner
'    Rlx lngCustID
    Set cn = Nothing
  Exit Function
ChkErr:
    ErrHandle "mod_Func7", "JustDoIt"
    cn.RollbackTrans
    JustDoIt = "-99," & strErr
End Function

'csPaymentClass.Approve
'
'Public Function Approve(ByVal lngCompCode As Long, ByVal strCardBillNo As String, _
'        ByVal lngCardCode As Long, ByVal strAccountNo As String, _
'        ByVal strCardExpDate As String, ByVal lngAmount As Long, ByVal strRealDate As String, _
'        ByVal strCMCode As String, ByVal strCMName As String, _
'        ByVal strPTCode As String, ByVal strPTName As String, _
'        ByVal strClctEn As String, ByVal strClctName As String, _
'        ByVal strUpdTime As String, ByVal strUpdEn As String, ByVal strEntryEn As String, _
'        ByVal lngDepositFlag As Long) As Long
'
'   objConn
'   uOwnerName
'
'   GetErrorCode
'   GetErrorMsg

