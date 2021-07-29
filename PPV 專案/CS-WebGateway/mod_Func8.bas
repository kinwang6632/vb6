Attribute VB_Name = "mod_Func8"
Option Explicit

Private cn As Object
Private rs As Object
Private rs105E As Object
Private rs105F As Object
Private rs033 As Object
Private strOwner As String
Private strCardBillNo As String
Private strOrderNo As String
Private strAuthenticID As String
Private strAuthenticCode As String
Private strName As String
Private strCardExpDate As String
Private lngCustID As Long
Private strUpdTime As String

Public Function JustDoIt(ByRef objWebPool As Object, _
                                        ByRef strPara As String) As String
  On Error GoTo ChkErr
  
    Dim objPayment As Object
    
    Dim varPara As Variant, varDetail As Variant
    
    Dim strServiceType As String
    Dim strBillData As String
    
    Dim lngDepositFlag As Long
    Dim lngCardCode As Long
    Dim sglTotalAmt As Single
    Dim strAccountNo As String
    
    Dim strCMCode As String, strCMName As String
    Dim strPTCode As String, strPTName As String
    Dim lngReturn As Long
    
    Dim strCitemName As String
    Dim lngAmount As Long
    Dim intCreditPoint As Integer
    Dim strBillNo As String
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
    
'    8�B�|���N���B�{��ID�B�d���B�H�Υd���Ĵ����B�H�Υd�I�����v�X�B�H�Υd�O(���)

    varPara = Split(strPara, vbCrLf)
    '*************************************************************************
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
'    For intLoop = 1 To UBound(varPara)
'        If varPara(intLoop) = Empty Then
'            JustDoIt = "-99,�ѼƤ���!"
'            GoTo 66
'        End If
'    Next intLoop
    '*************************************************************************
    strAuthenticID = varDetail(2)
    bytComp = Val(Left(strAuthenticID, 3))
'    lngCustID = Val(Right(strAuthenticID, 7))
    strAccountNo = varDetail(3) '   strAccountNo �H�Υd�d��(WEB�ǤJ)
    strCardExpDate = varDetail(4) & "" '   strCardExpDate �H�Υd���Ħ~��(WEB�ǤJ)
    lngCardCode = Val(varDetail(6) & "") '   lngCardCode �H�Υd�O(WEB�ǤJ)
    strAuthenticCode = varDetail(5)
    strOwner = GetOwner(cn)
    
'   Select count(*) as Cnt from so002b where MemberId=[�|���N��] and AuthenticId=[�{��ID]
    Set rs = cn.Execute("SELECT COMPCODE,CUSTID,NAME FROM " & strOwner & "SO002B" & Get_DB_Link(cn) & _
                                    " WHERE MEMBERID='" & varDetail(1) & "' AND AUTHENTICID='" & strAuthenticID & "'")
    
    If rs.AbsolutePage <= 0 Then '    If Cnt = 0 Then
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
    
    strCardBillNo = "2" & Right("00000000" & strCardBillNo, 8) '    ���ͤ@�H�Υd�q��渹(9�X).
    '   ��h :
    '       ��1�X���ӷ��ѧO�X. [2]��WEB�I�s;
    '       ��2�X�}�l���y����. �I�s Sequence Object S_SO033_CardBillNo ���o�̷s��. 8�X�k�a���ɹs.
    '       Ex : �ӷ��ѧO�X + �y���� �� 200000001
    
    bytComp = Val(rs("COMPCODE") & "")
    lngCustID = Val(rs("CUSTID") & "")
    strName = rs("NAME") & ""
    
'   �q��渹����: �褸�~��+ Sequence Object S_SO105E_OrderNo Ex: '2005030000001'
'    strOrderNo = GetValue(cn, "SELECT " & strOwner & "S_SO105E_ORDERNO.NEXTVAL FROM DUAL") & ""
'    strOrderNo = Trim(strOrderNo)
'
'    If strOrderNo = Empty Then
'        JustDoIt = "-99,[S_SO105E_ORDERNO] Sequence �`�Ǫ��󤣦s�b !"
'        GoTo 66
'    End If
'    strOrderNo = Format(Date, "YYYYMM") & Right("0000000" & strOrderNo, 7)
    
    '   ���q��s��
    If Not GetOrderNo(cn, strOwner, strOrderNo, JustDoIt) Then GoTo 66

    sglTotalAmt = 0
    Get_Def_CM cn, strOwner, strCMCode, strCMName '    strCMcode = "�w�]���O�覡�N�X" , strCMname = "�w�]���O�覡�W��"
    Get_Def_PT cn, strOwner, strPTCode, strPTName '    strPTcode = "�w�]�I�ں����N�X" , strPTname = "�w�]�I�ں����W��"
    strUpdTime = "" ' ���ʮɶ� Null
   
    cn.BeginTrans
    
    strUpdTime = Format(RightNow(cn), "YYYY/MM/DD HH:MM:SS")
    
    If Not Ins_2_105E Then
        JustDoIt = "-99,[ PPV �q���ƥD�� ] �s�W���� !"
        GoTo RollBack
    End If
    
'    Loop�B�z (�ĤG��Ѽƶ}�l)
'        �s�W�q��樭��SO105F��(�q��渹, ���O���إN��, ���O���ئW��, ���B, iPPV/PPV�I��, �~�̪A�����O)
'        �s�W��Ʀ�SO033��.�HT��覡���ͨö�J�H�Υd�q��渹
'    End Loop

    '    ���O���إN�� (���)
    '    ���O���إN�� (���)
    '    ...
    
    '   ����ڽs��
    strBillNo = GetBillNo(cn, strOwner, JustDoIt)
    
    If strBillNo = Empty Then GoTo RollBack
    
    If GetRS(cn, rs105F, "SELECT ORDERNO,CITEMCODE,CITEMNAME,AMOUNT" & _
                                        ",CREDITPOINT,SERVICETYPE,PREPAYFLAG,BILLNO,ITEM" & _
                                        " FROM " & strOwner & "SO105F" & Get_DB_Link(cn) & _
                                        " WHERE ROWID=''") Then
                                        '" WHERE BILLNO='' AND ORDERNO=''"
    
        For intLoop = 1 To UBound(varPara)
            strBillData = Trim(varPara(intLoop) & "")
            If strBillData <> Empty Then
                If GetChargeInfo(strBillData, strCitemName, lngAmount, intCreditPoint) Then
                    With rs105F
                            .AddNew
                            .Fields("OrderNo") = strOrderNo
                            .Fields("CitemCode") = strBillData
                            .Fields("CitemName") = strCitemName
                            .Fields("Amount") = lngAmount
                            .Fields("CreditPoint") = intCreditPoint
                            .Fields("ServiceType") = "C"
                            .Fields("BillNo") = strBillNo
                            .Fields("Item") = .RecordCount
                            .Fields("PrePayFlag") = 1
                            .Update
                            If Not Ins_2_033(strBillData, strBillNo, .Fields("Item"), lngAmount, strCardBillNo, _
                                                        strPTCode, strPTName, strCMCode, strCMName) Then
                                If strErr <> Empty Then
                                    JustDoIt = "-99," & strErr
                                Else
                                    JustDoIt = "-99,[ ������������� ] �s�W���� !"
                                End If
                                GoTo RollBack
                            End If
                    End With
                Else
                    JustDoIt = "-99,[ ���O���إN�X�� ] ��������N�X [" & strBillData & "]"
                    GoTo RollBack
                End If
            End If
        Next
    
    Else
        JustDoIt = "-99,[ iPPV/PPV�I�ƭq���� ] �}�ҥ��� !"
        GoTo RollBack
    End If
    
'   �̶ǨӰѼƩI�sPayment Gateway �Ҳդ������v(Approve)�ýд�(Deposit)���G�X�@API
    
'   PS: �`�p���B�Х�Web�ݦۦ�[�`�B�z
    
    lngDepositFlag = GetValue(cn, "SELECT PARA32 FROM " & strOwner & "SO043" & Get_DB_Link(cn) & _
                                                        " WHERE COMPCODE=" & bytComp & _
                                                        " AND SERVICETYPE='C'")
    
    Set objPayment = CreateObject("csPaymentClass.Approve")
    
    With objPayment '   �I�sPayment Gateway �Ҳդ������v(Approve)�ýд�(Deposit)���G�X�@API
            Set .objConn = cn
            .uOwnerName = strOwner
            lngReturn = .Approve(Val(bytComp), strCardBillNo, lngCardCode, strAccountNo, _
                                    strCardExpDate, Val(sglTotalAmt), "", strCMCode, _
                                    strCMName, strPTCode, strPTName, "WEB", "����", _
                                    "", "", "WEB", lngDepositFlag, strAuthenticCode)
            If lngReturn = 0 Then
                cn.CommitTrans
                JustDoIt = "0,���\," & strCardBillNo '       �^��(0,���\,�H�Υd�q��s��)
                GoTo 66
            Else '    If ���v���G��ok Then
                JustDoIt = "-7,�H�α��v���~" '       �^��(-7,�H�α��v���~)
                JustDoIt = JustDoIt & " " & CStr(lngReturn) & " " & .getErrorCode & " " & .getErrorMsg
                GoTo RollBack
            End If
    End With

'    ���G�X�B���G�T���r��(���)
'    �H�Υd�q��s��

RollBack:
    cn.RollbackTrans
  
66:
    On Error Resume Next
    Rlx rs
    Rlx rs105E
    Rlx rs105F
    Rlx rs033
    Rlx cn
    Rlx strName
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
    Rlx strOrderNo
    Rlx strCardExpDate
    Rlx strAccountNo
    Rlx strCMCode
    Rlx strCMName
    Rlx strPTCode
    Rlx strPTName
    Rlx strUpdTime
    Rlx intLoop
    Rlx lngCustID
    Rlx strOwner
    Rlx strCitemName
    Rlx lngAmount
    Rlx intCreditPoint
    Rlx strBillNo
    
    Set cn = Nothing
    Set rs = Nothing
    Set rs105E = Nothing
    Set rs105F = Nothing
    Set rs033 = Nothing
  Exit Function
ChkErr:
    ErrHandle "mod_Func8", "JustDoIt"
    cn.RollbackTrans
    JustDoIt = "-99," & strErr
End Function

Private Function GetChargeInfo(strCitemCode As String, _
                                                    ByRef strCitemName As String, _
                                                    ByRef lngAmount As Long, _
                                                    ByRef intCreditPoint As Integer) As Boolean
  On Error Resume Next
    If GetRS(cn, rs, "SELECT DESCRIPTION,AMOUNT,CREDITPOINT" & _
                            " FROM " & strOwner & "CD019" & Get_DB_Link(cn) & _
                            " WHERE CODENO=" & strCitemCode) Then
        strCitemName = rs("DESCRIPTION") & ""
        lngAmount = Val(rs("AMOUNT") & "")
        intCreditPoint = Val(rs("CREDITPOINT") & "")
        GetChargeInfo = True
    Else
        strCitemName = ""
        lngAmount = ""
        intCreditPoint = ""
    End If
End Function

Private Function Ins_2_105E() As Boolean
  On Error GoTo ChkErr
'    �s�W�@���q����Y��SO105E��
'    (�q��渹, ���q�O, �Ȥ�s��, �Ȥ�W��, PPV�q�ʤ覡�N�X, PPV�q�ʤ覡, �{�ҽs��, �m�W, �d��,
'        �H�Υd���Ĵ���, ���z�ɶ�, ���z�H���N��, ���z�H���W��, �~�̪A�����O, �H�Υd�q��s��)
    Dim strCustName As String
    strCustName = RPxx(cn.Execute("SELECT CUSTNAME FROM " & strOwner & "SO001" & Get_DB_Link(cn) & _
                                                        " WHERE CUSTID=" & lngCustID & _
                                                        " AND COMPCODE=" & bytComp).GetString(2, 1, "", "", "") & "")
    
    If GetRS(cn, rs105E, "SELECT ORDERNO,COMPCODE,CUSTID,CUSTNAME," & _
                                        "PPVORDERMODENO,PPVORDERMODE,AUTHENTICID," & _
                                        "NAME,ACCOUNTNO,CARDEXPDATE,ACCEPTTIME,ACCEPTEN,ACCEPTNAME," & _
                                        "SERVICETYPE,CARDBILLNO,UPDTIME,UPDEN" & _
                                        " FROM " & strOwner & "SO105E" & Get_DB_Link(cn) & _
                                        " WHERE ROWID=''") Then
                                        '" WHERE COMPCODE=" & bytComp & " AND ORDERNO=''"
        With rs105E
                .AddNew
                .Fields("ORDERNO") = strOrderNo
                .Fields("COMPCODE") = Val(bytComp)
                .Fields("CUSTID") = lngCustID
                .Fields("CUSTNAME") = strCustName
                .Fields("PPVORDERMODENO") = 2
                .Fields("PPVORDERMODE") = "����"
                .Fields("AUTHENTICID") = strAuthenticID
                .Fields("NAME") = strName
                .Fields("CARDEXPDATE") = Val(strCardExpDate)
                .Fields("ACCEPTTIME") = strUpdTime
                .Fields("ACCEPTEN") = "WEB"
                .Fields("UPDTIME") = strUpdTime
                .Fields("UPDEN") = "WEB"
                .Fields("ACCEPTNAME") = "����"
                .Fields("SERVICETYPE") = "C"
                .Fields("CARDBILLNO") = strCardBillNo
                .Update
        End With
    End If
    Ins_2_105E = True
    On Error Resume Next
    Rlx strCustName
  Exit Function
ChkErr:
    ErrHandle "mod_Func8", "Ins_2_105E"
    Ins_2_105E = False
End Function

Private Function Ins_2_033(ByVal strCitemCode As String, ByVal strBillNo As String, ByVal lngItem As Long, _
                                            ByVal lngAmount As Long, ByVal strCardBillNo As String, _
                                            ByVal strPTCode As String, ByVal strPTName As String, _
                                            ByVal strCMCode As String, ByVal strCMName As String) As Boolean
  On Error GoTo ChkErr
    Ins_2_033 = False
    If Not GetRS(cn, rs033, "SELECT * FROM " & strOwner & "SO033" & Get_DB_Link(cn) & " WHERE ROWID =''") Then
        'WHERE BILLNO ='' AND ITEM = -1
        strErr = "-99,[ ������������� ] �}�ҥ��� ! GetRS(rs033) "
        Exit Function
    End If
    With rs033
            Set .ActiveConnection = Nothing
'            strErr = strErr & lngCustID & "," & bytComp & "," & strBillNo & "," & lngItem & strCitemCode
            If Ins_Chg_Fld(Val(lngCustID), Val(bytComp), strBillNo, lngItem, strCitemCode, Format(strUpdTime, "YYYY/MM/DD")) Then
                .Fields("RealDate") = Format(strUpdTime, "YYYY/MM/DD")
                If strCMCode & "" <> "" Then .Fields("CMCode") = strCMCode
                If strCMName & "" <> "" Then .Fields("CMName") = strCMName
                If strPTCode & "" <> "" Then .Fields("PTCode") = strPTCode
                If strPTName & "" <> "" Then .Fields("PTName") = strPTName
'                .Fields("ShouldAmt") = lngAmount
                .Fields("OrderNo") = strOrderNo
                .Fields("CardBillNo") = NoZero(strCardBillNo)
                .Fields("AuthenticId") = NoZero(strAuthenticID)
                .Fields("ServiceType") = "C"
                .Fields("CreateTime") = strUpdTime
                .Fields("UpdTime") = strUpdTime
                .Fields("UpdEn") = "WEB"
                 Ins_2_033 = InsertToOracle(cn, strOwner & "SO033", rs033, " WHERE BILLNO = '' AND ITEM = -1")
                If Not Ins_2_033 Then strErr = "-99,[ ������������� ] �}�ҥ��� ! InsertToOracle(rs033) " & strErr
            Else
                strErr = "-99,���Ʀ� [ ��ƿ�SO033 ] ( Ins_Chg_Fld ) ���� !" & strErr
            End If
            .CancelUpdate
    End With
  Exit Function
ChkErr:
    ErrHandle "mod_Func8", "Ins_2_033"
End Function

Private Function Ins_Chg_Fld(ByVal lngCustID As Long, _
                                                ByVal lngCompCode As Long, _
                                                ByVal strBillNo As String, _
                                                ByVal lngItem As Long, _
                                                ByVal strCitemCode As String, _
                                                ByVal strShouldDate As String) As Boolean
  On Error GoTo ChkErr
    Dim objINC As Object
'    Dim objINC As New csAlterCharge4.clsInsertNewCharge
    Set objINC = CreateObject("csAlterCharge4.clsInsertNewCharge")
    With objINC
        .uOwnerName = strOwner
        .uErrPath = Environ("Temp")
        Set .uConn = cn
        Ins_Chg_Fld = .InsertChargeField(rs033, lngCustID, lngCompCode, "C", _
                                                                strBillNo, lngItem, strCitemCode, strShouldDate, _
                                                                "WEB", "����", , , True, False)
    End With
    On Error Resume Next
    Set objINC = Nothing
  Exit Function
ChkErr:
    ErrHandle "mod_Func8", "Ins_Chg_Fld"
    Ins_Chg_Fld = False
End Function
'PS:
'   �q��渹����: �褸�~��+ Sequence Object S_SO105E_OrderNo Ex: '2005030000001'
'   ���v���B: �ѰѼƶǤJ���Ҧ����O���إN��Ū��cd019.Amount���[�`����
'   ���O�覡�N�X, ���O�覡�W��: select CodeNo,Description from cd031 where RefNo=4 and StopFlag=0 and rownum=1
'   �I�ں����N�X, �I�ں����W��: select CodeNo,Description from cd032 where RefNo=4 and StopFlag=0 and rownum=1
'   ���O�H���N�X: 'WEB'
'   ���O�H���W��: '����'
'   ���q�O�N�X: so002b.CompCode
'   �Ȥ�s��: so002b.CustId
'   �Ȥ�W��: so001.CustName
'   PPV�q�ʤ覡�N�X: 2
'   PPV�q�ʤ覡: 'WEB'
'   �m�W: so002b.Name
'   ���z�ɶ�: �ѵ{������
'   ���z�H���N��: 'WEB'
'   ���z�H���W��: '����'
'   �~�̪A�����O: 'C'
'   �H�Υd�q��s������: �Ѧҵ��@
'   ���O���ئW��:  CD019.Description
'   ���B:  CD019.Amount
'   iPPV/PPV�I��: CD019.CreditPoint

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
