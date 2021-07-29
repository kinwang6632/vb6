Attribute VB_Name = "mod_Func12"
Option Explicit

Private cn As Object
Private rs As Object
Private rs105G As Object
Private rs105E As Object

Private strAuthenticID As String
Private strOwner As String
Private intLoop As Integer
Private lngCustID As Long
Private strOrderNo As String
Private strOrderItem As String
Private strName As String
Private strUpdTime As String
Private blnOneOfOK As Boolean
Private strXMLresult As String

'�p���I�O�`�حq��(�ݸg�LSMS�q�ʱK�X�{��,�å]�ACA�t�έq�ʱ��v�@�~�γB�z���G)
Public Function JustDoIt(ByRef objWebPool As Object, _
                                        ByRef strPara As String) As String
  On Error GoTo ChkErr
  
    
    Dim varPara As Variant
    Dim varDetail As Variant
    
    Dim strBillData As String
    
    Dim strMsg As String
    
'    Set cn = Get_General_CN
'
'    If cn.State <= 0 Then
'        Set cn = Get_General_CN
'        If cn.State <= 0 Then
'            strErr = "�L�k�s�u��Ʈw!"
'            JustDoIt = "-99,��Ʈw�s�u����!"
'            GoTo 66
'        End If
'    End If
    
'   12�B�|���N���B�{��ID�BPPV��I�q�ʱK�X�B2(��ӷ�����) (���)

'   XML(�q��渹�B�q�涵���BSTB�Ǹ��BICC�d���B�`�ئW�١B���A)
'   0:  ���\
'   -2:  �q��PPV�{�ҿ��~
'   -4 : ���ҥ�PPV�q���v
'   -11 : �p���I�O�`�حq�ʥ���
'***************************************************************************
'#3018 �P�_�ǤJ���ӼƬO�_���T By Kin 2007/08/31
    varPara = Split(strPara, vbCrLf)
    
    If UBound(varPara) < 2 Then
        JustDoIt = "-99,�ѼƸ�Ƥ����T ! �нT�{ !"
        GoTo 66
    End If

    varDetail = Split(varPara(0), ",")
    If UBound(varDetail) < 4 Then
        JustDoIt = "-99,�ѼƤ���!"
        GoTo 66
    End If
    For intLoop = 1 To UBound(varPara)
        If (UBound(Split(varPara(intLoop), ",")) < 6) And (varPara(intLoop) <> Empty) Then
            JustDoIt = "-99,�ѼƤ���!"
            GoTo 66
        End If
    Next intLoop
'    If UBound(Split(varPara(1), ",")) < 6 Then
'        JustDoIt = "-99,�ѼƤ���!"
'        GoTo 66
'    End If
'    If UBound(Split(varPara(2), ",")) < 6 Then
'        JustDoIt = "-99,�ѼƤ���!"
'        GoTo 66
'    End If
'***************************************************************************
    strAuthenticID = varDetail(2)
    bytComp = Val(Left(strAuthenticID, 3)) '   ���q�O�N�X
'    lngCustID = Val(Right(strAuthenticID, 7))
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
    
'   Select count(*) as Cnt from so002b where MemberId=[�|���N��] and AuthenticId=[�{��ID] and PPVOrderPwd=[ PPV��I�q�ʱK�X]
    Set rs = cn.Execute("SELECT COMPCODE,CUSTID,NAME FROM " & strOwner & "SO002B" & Get_DB_Link(cn) & _
                                    " WHERE MEMBERID='" & varDetail(1) & "'" & _
                                    " AND AUTHENTICID='" & strAuthenticID & "'" & _
                                    " AND PPVORDERPWD='" & varDetail(3) & "'")
    
    If rs.AbsolutePage <= 0 Then '    If Cnt = 0 Then
        JustDoIt = "-2,�q��PPV�{�ҿ��~" '        �^��(-2,�q��PPV�{�ҿ��~)
        GoTo 66
    End If '    End If
    
'    API-12�W��վ�:�@�Y�ˮ�PPVRight<>1�ɡA�վ㬰���^���T���A��update�@so004�@ppvright=1��log SOAC0202�Y�i�C
'    If Not Chk_DSTB_OK(varPara) Then
'        JustDoIt = "-4, ���ҥ�PPV�q���v"
'        GoTo 66
'    End If
    
    bytComp = Val(rs("COMPCODE") & "")
    lngCustID = Val(rs("CUSTID") & "")
    strName = rs("NAME") & ""
    
'   �̶ǤJ���Ѽ�(ICC�Ǹ��B���~�N�X�B�`�ئW�١B�I��)�I�sCA�R�O[�q�� PPV �`��]�D��ı�ƼҲ�. �����N��P30
'   �|���N���B�{��ID�BPPV��I�q�ʱK�X�B2(��ӷ�����) (���)
'   �]�Ƭy�����BSTB�Ǹ��BICC�Ǹ��B���~�N�X�B�`�ئW�١B����(���)
'   �K�K

'   �����q��渹
    If Not GetOrderNo(cn, strOwner, strOrderNo, JustDoIt) Then GoTo 66
'   1�D �q��渹����: �褸�~��+ Sequence Object S_SO105E_OrderNo
'   Ex: '2005030000001

    If Not GetRS(cn, rs105G, "SELECT * FROM " & strOwner & "SO105G" & Get_DB_Link(cn) & _
                                            " WHERE ROWID=''") Then
                                        ' " WHERE COMPCODE=" & bytComp & " AND ORDERNO=''"
        JustDoIt = "-99,[ PPV�`�حq�ʩ����� ] �}�ҥ��� !"
        GoTo 66
    End If

    cn.BeginTrans
    strUpdTime = Format(RightNow(cn), "YYYY/MM/DD HH:MM:SS")
    blnOneOfOK = False
    
    For intLoop = 1 To UBound(varPara) '   Loop�B�z(�ĤG��Ѽƶ}�l)�F�B�z�UDSTB�ҭq�ʪ��`��
        
        strBillData = varPara(intLoop) & ""
        
        If strBillData <> Empty Then
            
            varDetail = Split(strBillData, ",")
            '   �]�Ƭy�����BSTB�Ǹ��BICC�Ǹ��B���~�N�X�B�`�ئW�١B����(���) + ���X�ɶ�(���) 2006/06/20
            '   �K�K..
            
            '   �I�sCA�R�O[�q�� PPV �`��]�D��ı�ƼҲ�. �����N��P30
            '   strNotes= "ProductId" & "~" & "ProductName" & "~" & "CreditAmt"
            If CallNagraProcess(varDetail(2), varDetail(1), varDetail(3) & "~" & varDetail(4) & "~" & varDetail(5), strMsg) Then    '       If �R�O���\ Then
                blnOneOfOK = True
                ' �s�W�q��樭��SO105G��
                ' (�q��渹,�q�涵��,���~�N�X,�v���N�X,�`�ئW��,�`�ئW��(�b��),�`�ص���,�W�D��ID,�`�ػ���,EPG�W���I��,����O,�I��,�q�ʤ覡�N�X, �q�ʤ覡,�p�p���B,�p�p�I��,�]�Ƭy����,���~�Ǹ�,���z�d�Ǹ�,���ʤH���N��,���ʤH��,�~�̪A�����O,EventId)
                If Not Ins_2_105G(varDetail(0), varDetail(1), varDetail(2), varDetail(3), varDetail(4), varDetail(5), True) Then
                    JustDoIt = "-99,[PPV�`�حq�ʩ�����] �s�W����! " & strErr
                    GoTo RollBack
                End If
                MakeXML "���\", varDetail(1), varDetail(2), varDetail(4), varDetail(6)
            Else '       Else
                ' �s�W�q��樭��SO105G��
                ' (�q��渹, �q�涵��,���~�N�X,�v���N�X,�`�ئW��,�`�ئW��(�b��),�`�ص���,�W�D��ID,�`�ػ���,EPG�W���I��,����O,�I��,�q�ʤ覡�N�X,�q�ʤ覡,�p�p���B,�p�p�I��,�]�Ƭy����,���~�Ǹ�,���z�d�Ǹ�,���ʤH���N��,���ʤH��,�~�̪A�����O, EventId,�����H���N��,�����H��,�����ɶ�,�q�������]�N�X,�q�������] )
                If Not Ins_2_105G(varDetail(0), varDetail(1), varDetail(2), varDetail(3), varDetail(4), varDetail(5), False) Then
                    JustDoIt = "-99,[PPV�`�حq�ʩ�����] �s�W����! " & strErr
                    GoTo RollBack
                End If
                MakeXML "����", varDetail(1), varDetail(2), varDetail(4), varDetail(6)
            End If '       End If
                
        End If
        
    Next '   End Loop
    
    If blnOneOfOK Then ' ���\
        ' �s�W�@���q����Y��SO105E��(�q��渹, ���q�O, �Ȥ�s��, �Ȥ�W��, PPV�q�ʤ覡�N�X, PPV�q�ʤ覡, �{�ҽs��, �m�W, ���z�ɶ�, ���z�H���N��, ���z�H���W��, �~�̪A�����O, ���ʤH��, ���ʮɶ� )
        If Ins_2_105E Then
            cn.CommitTrans
            ' XML(�q��渹�B�q�涵���BSTB�Ǹ��BICC�d���B�`�ئW�١B���A)
            JustDoIt = "0,���\" & vbCrLf & _
                                "<?xml version=""1.0"" encoding=""BIG5"" standalone=""yes"" ?>" & vbCrLf & _
                                "<DataSet>" & vbCrLf & _
                                "  <DataTable>" & vbCrLf & _
                                strXMLresult & _
                                "  </DataTable>" & vbCrLf & _
                                "</DataSet>" & vbCrLf ' �^��(0, ���\) ��XML��
            GoTo 66
        Else
            JustDoIt = "-99,[PPV�q���ƥD��] �s�W����!" & strErr
            GoTo RollBack
        End If
    Else '   Else If �B�z�UDSTB�ҭq�ʪ��`�ث��O�������� Then
        JustDoIt = "-11, �p���I�O�`�حq�ʥ���" & strErr ' �^��(-11, �p���I�O�`�حq�ʥ���)
         GoTo RollBack ' Rollback; (�Y�����q�ʸ�ƬҨ������s)
    End If '   End If
    
RollBack:
    cn.RollbackTrans
    
66:
    On Error Resume Next
    Rlx rs
    Rlx cn
    Rlx varPara
    Rlx varDetail
    Rlx strAuthenticID
    Rlx intLoop
    Rlx strOwner
    Rlx strMsg
    Rlx strBillData
    Rlx lngCustID
    Rlx strOrderNo
    Rlx strOrderItem
    Rlx strName
    Rlx strUpdTime
    Rlx blnOneOfOK
    Rlx rs105G
    Rlx rs105E
    Rlx strXMLresult
'    Rlx lngCustID
    Set cn = Nothing
  Exit Function
ChkErr:
    ErrHandle "mod_Func12", "JustDoIt"
    cn.RollbackTrans
    JustDoIt = "-99," & strErr
End Function

Private Sub MakeXML(ByVal strStatus As String, _
                                    ByVal strFaciSno As String, _
                                    ByVal strSmartCardNo As String, _
                                    ByVal strProductName As String, _
                                    ByVal strEventBeginTime As String)
  On Error GoTo ChkErr
    '   XML(�q��渹�B�q�涵���BSTB�Ǹ��BICC�d���B�`�ئW�١B���A(���\������)) + ���X�ɶ�
    strXMLresult = strXMLresult & _
                            "    <DataRow ORDERNO=""" & strOrderNo & """" & _
                                " ORDERITEM=""" & strOrderItem & """" & _
                                " FACISNO=""" & strFaciSno & """" & _
                                " SMARTCARDNO=""" & strSmartCardNo & """" & _
                                " PRODUCTNAME=""" & RPC(strProductName) & """" & _
                                " STATUS=""" & strStatus & """" & _
                                " DEVENTBEGINTIME=""" & strEventBeginTime & """/>" & vbCrLf
  Exit Sub
ChkErr:
    ErrHandle "mod_Func12", "MakeXML"
End Sub

Private Function CallNagraProcess(ByVal strSmartCardNo As String, ByVal strFaciSno As String, _
                                                        ByVal strNotes As String, ByRef strMsg As String) As Boolean
  On Error GoTo ChkErr
    Dim objNS As Object
    Dim strQry As String
    Set objNS = CreateObject("setIPPV.clsNagraSTB")
    With objNS
        strMsg = ""
        .uCompCode = bytComp
        Set .uConn = cn
        garyGi(1) = "WEB"
        .ugaryGi = Join(garyGi, Chr(9))
        .uErrPath = Environ("Temp")
        strQry = "SELECT PPVRIGHT FROM " & strOwner & "SO004" & Get_DB_Link(cn) & _
                        " WHERE FACISNO='" & strFaciSno & "'" & _
                        " AND INSTDATE IS NOT NULL AND PRDATE IS NULL"
        If Val(GetValue(cn, strQry) & "") = 0 Then .EnPPVright bytComp, Val(lngCustID), strSmartCardNo, "", strFaciSno, "�Ұ�PPV�q���v", strFaciSno, ""
        CallNagraProcess = .OrdPPVPro(bytComp, lngCustID, strSmartCardNo & "", "", strFaciSno & "", "PPV�`�س]�w", strNotes)
        strMsg = .uMsgBack

'       �ª��{���X
'        If .EnPPVright(bytComp, Val(lngCustID), strSmartCardNo, "", strFaciSno, "�Ұ�PPV�q���v", strFaciSno, "") Then
'            CallNagraProcess = .OrdPPVPro(bytComp, lngCustID, strSmartCardNo & "", "", strFaciSno & "", "PPV�`�س]�w", strNotes)
'        End If
    
    End With
    Set objNS = Nothing
  Exit Function
ChkErr:
    strMsg = "[STB �]�w���~] - " & Err.Description
    On Error Resume Next
    Set objNS = Nothing
End Function

Private Function Ins_2_105E() As Boolean
  On Error GoTo ChkErr
'   �s�W�@���q����Y��SO105E��
'   (�q��渹, ���q�O, �Ȥ�s��, �Ȥ�W��, PPV�q�ʤ覡�N�X, PPV�q�ʤ覡, �{�ҽs��, �m�W,
'       ���z�ɶ�, ���z�H���N��, ���z�H���W��, �~�̪A�����O, ���ʤH��, ���ʮɶ� )
    
    Dim strCustName As String
    strCustName = RPxx(cn.Execute("SELECT CUSTNAME FROM " & strOwner & "SO001" & Get_DB_Link(cn) & _
                                                        " WHERE CUSTID=" & lngCustID & _
                                                        " AND COMPCODE=" & bytComp).GetString(2, 1, "", "", "") & "")
    
    If GetRS(cn, rs105E, "SELECT ORDERNO,COMPCODE,CUSTID,CUSTNAME," & _
                                        "PPVORDERMODENO,PPVORDERMODE,AUTHENTICID,NAME," & _
                                        "ACCEPTTIME,ACCEPTEN,ACCEPTNAME,SERVICETYPE,UPDEN,UPDTIME" & _
                                        " FROM " & strOwner & "SO105E" & Get_DB_Link(cn) & _
                                        " WHERE ROWID=''") Then
                                        ' " WHERE COMPCODE=" & bytComp & " AND ORDERNO=''"
        With rs105E
                .AddNew
                .Fields("ORDERNO") = strOrderNo ' �q��渹
                .Fields("COMPCODE") = Val(bytComp) ' ���q�O
                .Fields("CUSTID") = lngCustID ' �Ȥ�s��
                .Fields("CUSTNAME") = strCustName ' �Ȥ�W��
                .Fields("PPVORDERMODENO") = 2 ' PPV�q�ʤ覡�N�X
                ' Liga , ���D�� : 2232 , Web API����_20060322_Liga.doc
                ' API-12:�q�椧PPV�q�ʤ覡���e�ȥѡuWEB�v�אּ�u�����v
                .Fields("PPVORDERMODE") = "����" ' PPV�q�ʤ覡
                .Fields("AUTHENTICID") = strAuthenticID ' �{�ҽs��
                .Fields("NAME") = strName ' �m�W
                .Fields("ACCEPTTIME") = strUpdTime ' ���z�ɶ�
                .Fields("ACCEPTEN") = "WEB" ' ���z�H���N��
                .Fields("ACCEPTNAME") = "����" ' ���z�H���W��
                .Fields("SERVICETYPE") = "C" ' �~�̪A�����O
                .Fields("UPDEN") = "WEB" ' ���ʤH��
                .Fields("UPDTIME") = strUpdTime ' ���ʮɶ�
                .Update
        End With
    End If
    Ins_2_105E = True
    On Error Resume Next
    Rlx strCustName
  Exit Function
ChkErr:
    ErrHandle "mod_Func12", "Ins_2_105E"
End Function

Private Function Ins_2_105G(ByVal strFaciSeqNo As String, _
                                                ByVal strFaciSno As String, _
                                                ByVal strSmartCardNo As String, _
                                                ByVal strProductID As String, _
                                                ByVal strProductName As String, _
                                                ByVal lngAmount As Long, _
                                                blnSucceed As Boolean) As Boolean
  On Error GoTo ChkErr

'   �s�W�q��樭��SO105G��

'   If �R�O���\ Then
'     (�q��渹,�q�涵��,���~�N�X,�`�ئW��,�`�ئW��(�b��),�`�ص���,�W�D��ID,�q�ʤ覡�N�X, �q�ʤ覡,
'       �p�p���B,�p�p�I��,�]�Ƭy����,���~�Ǹ�,���z�d�Ǹ�,���ʤH���N��,���ʤH��,�~�̪A�����O)
'   Else
'     (�q��渹, �q�涵��,���~�N�X,�`�ئW��,�`�ئW��(�b��),�`�ص���,�W�D��ID,�q�ʤ覡�N�X,�q�ʤ覡,
'       �p�p���B,�p�p�I��,�]�Ƭy����,���~�Ǹ�,���z�d�Ǹ�,���ʤH���N��,���ʤH��,�~�̪A�����O,
'       �����H���N��,�����H��,�����ɶ�,�q�������]�N�X,�q�������] )
'   End If
    
'   6�D SO105G��ƨӷ��G
'       (1) ��WEB�ǰѭȦ^��G���~�N�X,�`�ئW��,�p�p���B,�]�Ƭy����,���~�Ǹ�,���z�d�Ǹ��C
'       (2) �β��~�N�X(PRODUCTID)��SO155��M���ƫ�^��G�`�ئW��(�b��),�`�ص���,�W�D��ID�C
'       (3) �Ш̨�L��컡���^��G�q�ʤ覡�N�X, �q�ʤ覡,,���ʤH���N��,���ʤH��,�~�̪A�����O,
'               �����H���N�X,�����H��,�����ɶ�,������]�N�X,������]
            
    With rs105G
            .AddNew
            
            .Fields("OrderNo") = strOrderNo ' �q��渹
            strOrderItem = .RecordCount
            .Fields("OrderItem") = Val(strOrderItem) ' �q�涵��
            .Fields("ProductId") = strProductID ' ���~�N�X
            .Fields("ProductName") = strProductName ' �`�ئW��
            
            Program strProductID ' �β��~�N�X(PRODUCTID)��SO155��M���ƫ�^��G�`�ئW��(�b��),�`�ص���,�W�D��ID�C
            
            .Fields("PPVOrderModeNo") = 2 ' �q�ʤ覡�N�X
            
            ' Liga , ���D�� : 2232 , Web API����_20060322_Liga.doc
            ' API-12:�q�椧PPV�q�ʤ覡���e�ȥѡuWEB�v�אּ�u�����v
            
            .Fields("PPVOrderMode") = "����" ' �q�ʤ覡
            
            .Fields("Amount") = lngAmount ' ��������Web�Ѽ�
            
            ' Liga ������p�p�I���P
            '.Fields("CreditAmt") = .Fields("CreditAmt") ' �p�p�I�� ( EPGPrice + HandlingCredit )
            
            .Fields("FaciSeqNo") = strFaciSeqNo ' �]�Ƭy����
            .Fields("FaciSNO") = strFaciSno ' ���~�Ǹ�
            .Fields("SmartCardNo") = strSmartCardNo ' ���z�d�Ǹ�
            .Fields("ModifyEn") = "WEB" ' ���ʤH���N��
            .Fields("ModifyName") = "����" ' ���ʤH��
            .Fields("ServiceType") = "C" ' �~�̪A�����O
            
            If Not blnSucceed Then
                .Fields("CancelEn") = "WEB" ' �����H���N��
                .Fields("CancelName") = "����" ' �����H��
                .Fields("CancelTime") = strUpdTime ' �����ɶ�
                 CancelReason ' ������]
            End If
            .Update
    End With
    
    Ins_2_105G = True
  Exit Function
ChkErr:
    ErrHandle "mod_Func12", "Ins_2_105G"
End Function

Private Function Program(ByVal strProductID As String) As Boolean
  On Error GoTo ChkErr
    ' �β��~�N�X(PRODUCTID)��SO155��M���ƫ�^��G�`�ئW��(�b��),�`�ص���,�W�D��ID�C ( �W�D��ID Liga �����s )
    Program = GetRS(cn, rs, "SELECT BILLNAME,DPARENTALRATING FROM " & strOwner & "SO155" & Get_DB_Link(cn) & _
                                            "WHERE PRODUCTID='" & strProductID & "'")
    If Program Then
        With rs105G
            .Fields("BillProductName") = rs!BillName & "" ' �`�ئW��(�b��)
            .Fields("ParentalRating") = rs!dParentalRating & "" ' �`�ص���
            '.Fields("ContentProviderID") = .Fields("ContentProviderID") ' �W�D��ID Liga �����s
        End With
    End If
  Exit Function
ChkErr:
    ErrHandle "mod_Func12", "Program"
End Function

Private Function CancelReason() As Boolean ' ������]�N�X
  On Error GoTo ChkErr
    ' ������]�N�X�B������]: ��CD015.REFNO=4���C(�Y���h���h����1���^��)
    CancelReason = GetRS(cn, rs, "SELECT CODENO,DESCRIPTION FROM " & strOwner & "CD015" & Get_DB_Link(cn) & _
                                                    "WHERE STOPFLAG=0 AND ASSORT=4 AND SERVICETYPE='C' OR SERVICETYPE IS NULL" & _
                                                    " ORDER BY CODENO")
    If CancelReason Then
        With rs105G
            .Fields("ReturnCode") = rs!CodeNo & "" ' �q�������]�N�X
            .Fields("ReturnName") = rs!Description & "" ' �q�������]
        End With
    End If
  Exit Function
ChkErr:
    ErrHandle "mod_Func12", "CancelReason"
End Function

'Private Function Chk_DSTB_OK(ByRef varPara As Variant) As Boolean
'  On Error GoTo ChkErr
'    Dim strQry As String
''   Select PPVright as right from so004 where facisno=[ STB�Ǹ�] and instdate is not null and prdate is null
''   If Right! = 1 Then
''       �^��(-4, ���ҥ�PPV�q���v)
''   End If
''   �]�Ƭy�����BSTB�Ǹ��BICC�Ǹ��B���~�N�X�B�`�ئW�١B�I��(���)
'    Chk_DSTB_OK = True
'    For intLoop = 1 To UBound(varPara) '   Loop�B�z(�ĤG��Ѽƶ}�l)�F�Y�ˮ֦UDSTB�O�_�ŦX�q�ʸ��
'        strQry = "SELECT PPVRIGHT FROM " & strOwner & "SO004" & Get_DB_Link(cn) & _
'                        " WHERE FACISNO='" & Split(varPara(intLoop), ",")(1) & "'" & _
'                        " AND INSTDATE IS NOT NULL AND PRDATE IS NULL"
'        Chk_DSTB_OK = Val(GetValue(cn, strQry) & "") = 1
'        If Not Chk_DSTB_OK Then Exit For
'    Next '   End Loop
'  Exit Function
'ChkErr:
'    ErrHandle "mod_Func12", "Chk_DSTB_OK"
'End Function



' ���U XXX Block �� , ���������W��
' XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'3�D ���~�N�X,�v���N�X,�`�ئW��,�`�ئW��(�b��),�`�ص���,�W�D��ID,�`�ػ���,EPG�W���I��, EventId --- �����pSO154, SO155, SO155A���oinsert�C
' XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'4�D �q�ʤ覡�N�X,�q�ʤ覡, ����O,�I��---
'       select codeno, description, charge, credit from so156 where refno=2;
' XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'5�D �p�p���B=�`�ػ���+����O�B�p�p�I��=EPG�W���I��+�I��
' XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
    
'7�D ��L��컡���K
'       ���ʤH���N��: 'WEB'
'       ���ʤH��: "����'
'       �~�̪A�����O: 'C'
'       �����H���N��: 'WEB'
'       �����H��: "����'
'       �����ɶ�: �t�β���
'       ���q�O�N�X: so002b.CompCode
'       �Ȥ�s��: so002b.CustId
'       �Ȥ�W��: so001.CustName
'       PPV�q�ʤ覡�N�X: 2
'       PPV�q�ʤ覡: 'WEB'
'       �m�W: so002b.Name
'       ���z�ɶ�: �ѵ{������
'       ���z�H���N��: 'WEB'
'       ���z�H���W��: '����'
    

