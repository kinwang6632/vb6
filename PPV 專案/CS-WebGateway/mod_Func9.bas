Attribute VB_Name = "mod_Func9"
Option Explicit

Private cn As Object
Private strXMLresult As String
Private blnOneOfOK As Boolean

' �ˮֳB�z���� STB �� iPPV �q���v�O�_�w�}��
Private Function IPPVok(strOwner As String, strSeqNo As String, Optional blnNoDataFound As Boolean = False) As Boolean
  On Error Resume Next
    IPPVok = (Val(RPxx(cn.Execute("SELECT IPPVRIGHT FROM " & _
                                                        strOwner & "SO004" & Get_DB_Link(cn) & _
                                                        "WHERE SEQNO='" & strSeqNo & "'").GetString(2, 1, "", "", "") & "")) = 1)
    If Err.Number = 3021 Then
        Err.Clear
        blnNoDataFound = True
        IPPVok = False
    ElseIf Err.Number <> 0 Then
        ErrHandle "mod_Func9", "IPPVok"
    End If
End Function

' �P�_ ��/�L �^�Ǿ���
Private Function ReturnBack(strOwner As String, strSeqNo As String) As Boolean
  On Error GoTo ChkErr
    Dim lngModelCode As Long
    Dim strQry As String
    strQry = "SELECT IPPVCALLBACK FROM " & strOwner & "SO041" & Get_DB_Link(cn) & _
                    "WHERE ROWNUM=1"
    ReturnBack = (Val(GetValue(cn, strQry)) = 1)
    If ReturnBack Then
        strQry = "SELECT MODELCODE FROM " & strOwner & "SO004" & Get_DB_Link(cn) & _
                        "WHERE SEQNO='" & strSeqNo & "'"
        lngModelCode = Val(GetValue(cn, strQry))
        If lngModelCode > 0 Then
            strQry = "SELECT FUNCFLAG01 FROM " & strOwner & "CD043" & Get_DB_Link(cn) & _
                            "WHERE CODENO=" & lngModelCode
            ReturnBack = ReturnBack And (Val(GetValue(cn, strQry)) = 1)
        Else
            ReturnBack = False
        End If
    End If
  Exit Function
ChkErr:
    ErrHandle "mod_Func9", "ReturnBack"
End Function

Private Function Get_IPPV_Credit(strOwner As String, strSeqNo As String) As Integer
  On Error GoTo ChkErr
    Get_IPPV_Credit = Val(cn.Execute("SELECT IPPVCREDIT FROM " & strOwner & "SO004" & Get_DB_Link(cn) & _
                                                            "WHERE SEQNO='" & strSeqNo & "'").GetString(2, 1, "", "", 0))
  Exit Function
ChkErr:
    ErrHandle "mod_Func9", "Get_IPPV_Credit"
End Function

Public Function JustDoIt(ByRef objWebPool As Object, _
                                        ByRef strPara As String) As String
  On Error GoTo ChkErr
    Dim strQry As String
    Dim varPara As Variant
    Dim varData As Variant
    Dim objCnPool As Object
    Dim objIPPV As Object
    
    Dim rs As Object
    Dim lngCustID As Long
    Dim intLoop As Integer
    Dim strOwner As String
    Dim strAuthenticID As String
    Dim blnNoData As Boolean
    
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
    '************************************************************************
    '#3018 �P�_�ǤJ���ӼƬO�_���T By Kin 2007/08/31
    varPara = Split(strPara, vbCrLf)
    If UBound(varPara) < 2 Then
        JustDoIt = "-99,�ѼƤ���!"
        GoTo 66
    End If
    For intLoop = 1 To UBound(varPara)
        If (UBound(Split(varPara(intLoop), ",")) < 3) And (varPara(intLoop) <> Empty) Then
            JustDoIt = "-99,�ѼƤ���!"
            GoTo 66
        End If
    Next intLoop
    varPara = Split(varPara(0), ",")
    If UBound(varPara) <> 2 Then
        JustDoIt = "-99,�ѼƤ���!"
        GoTo 66
    End If
    '************************************************************************
    strAuthenticID = varPara(2)
    bytComp = Val(Left(strAuthenticID, 3))
    lngCustID = Val(Right(strAuthenticID, 7))
    
    varData = Split(strPara, vbCrLf)
    
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
    
    ''�̶ǨӰѼƩI�sCA�R�O�I�ƤU���Ҳ� (STBCreDold)

    '�I�ƽu�W�q��--�U���I��
    '9�B�|���N���B�{��ID(���)
    '�]�Ƭy�����BSTB�Ǹ��BICC�Ǹ��B�I��(���)
    '�]�Ƭy�����BSTB�Ǹ��BICC�Ǹ��B�I��(���)
    '�K
    '0 : ���\
    '-8 : �I�ƤU�����~
    '-19 ���ҥ�IPPV�q���v
    
    ''Loop�B�z (�ĤG��Ѽƶ}�l)
    
    blnOneOfOK = False
    strXMLresult = ""
    
    For intLoop = 1 To UBound(varData)
        If RPxx(CStr(varData(intLoop))) <> Empty Then
            varPara = Split(varData(intLoop), ",")
            '�ˮֳB�z����STB��iPPV�q���v�O�_�w�}��(IPPVright=1 ?).
            ''If IPPVright <> 1 Then
                ''  �I�sCA�R�O�D��ı�ƶ}��iPPV�q���v�Ҳ� (EnIPPVright)
            ''End If
            
            JustDoIt = ""
            
            If Not IPPVok(strOwner, CStr(varPara(0)), blnNoData) Then
                ''�Ұ�IPPV�q���v
                    'Public Function EnIPPVright(ByVal strCompCode As String, _
                    '    ByVal lngCustId As Long, ByVal strSmartCard As String, _
                    '    ByVal strResvTime As String, ByVal strSTB As String, _
                    '    ByVal strDoCaption As String, ByVal strFaciSNo As String) As Boolean
                    '   Format(Now, "YYYY/MM/DD HH:MM")
                If blnNoData Then
'                    JustDoIt = "-99,�]�Ƭy�������s�b"
                    JustDoIt = "�]�Ƭy�������s�b"
                Else
                    If Not objIPPV.EnIPPVright(CStr(bytComp), _
                                                        lngCustID, _
                                                        CStr(varPara(2)), _
                                                        "", _
                                                        CStr(varPara(1)), _
                                                        "", _
                                                        CStr(varPara(1))) Then
                        strErr = objIPPV.uMsgBack
'                        JustDoIt = "-19,���ҥ�IPPV�q���v"
                        JustDoIt = "���ҥ�IPPV�q���v " & strErr
                    End If
                End If
            End If
            
            If JustDoIt <> Empty Then
                MakeXML CStr(varPara(0)), CStr(varPara(1)), CStr(varPara(2)), CStr(varPara(3)), JustDoIt
            Else
                With objIPPV
                    If Not objIPPV.STBCreDold(CStr(bytComp), _
                                                                lngCustID, _
                                                                CStr(varPara(2)), _
                                                                "", _
                                                                CStr(varPara(1)), _
                                                                "", _
                                                                "03", _
                                                                CStr(Val(varPara(3)) + Get_IPPV_Credit(strOwner, CStr(varPara(0)))), _
                                                                Format(Now, "YYYY/MM/DD HH:MM"), _
                                                                ReturnBack(strOwner, CStr(varPara(0)))) Then
                        strErr = .uMsgBack
                        JustDoIt = "�I�ƤU�����~"
'                        JustDoIt = "-8,�I�ƤU�����~"
                        MakeXML CStr(varPara(0)), CStr(varPara(1)), CStr(varPara(2)), CStr(varPara(3)), JustDoIt & " " & strErr
                    Else
                        MakeXML CStr(varPara(0)), CStr(varPara(1)), CStr(varPara(2)), CStr(varPara(3)), .uMsgBack
                        blnOneOfOK = True
    '                    If .uMsgBack <> Empty Then
    '                        JustDoIt = "0,���\(" & .uMsgBack & ")"
    '                    Else
    '                        JustDoIt = "0,���\"
    '                    End If
                    End If
                End With
            End If
            
            ''�U��/�]�wIPPV�I��or (�L�^��)
            'Public Function STBCreDold(ByVal strCompCode As String, _
            '    ByVal lngCustId As Long, ByVal strSmartCard As String, _
            '    ByVal strResvTime As String, ByVal strSTB As String, _
            '    ByVal strDoCaption As String, ByVal strCreditMode As String, _
            '    ByVal strCredit As String, ByVal strCollectDate As String, ByVal blnReturnBk As Boolean) As Boolean
            
            '    ��i�׶}����A�]�ӹL�b�������ä�C ��:
            '    ResvTime=''
            '    ��i�׶}����A�]�ӹL�b�������ä�C ��:
            '    CreditMode='01'
            '    ��i�׶}����A�]�ӹL�b�������ä�C ��:
            '   CollectDate = sysdate
            
            ''If �I�ƤU�����G��ok Then
                ''  �^��(-8,�I�ƤU�����~)
            ''Else
                ''  �^��(0,���\)
            ''End If
            
            ''�I�sCA�R�O�D��ı���I�ƤU���Ҳ�(STBCreDold[���^�Ǿ���call]/ STBCreDoldNb[�L�^�Ǿ���call])
            ''PS:
            ''���^�Ǿ���w�q : so041. iPPVCallBack=1 and ��STB ModelCode ������ CD043. FuncFlag01=1
            ''�L�^�Ǿ���w�q : so041. iPPVCallBack=0 or (so041. iPPVCallBack=1 and ��STB ModelCode������CD043. FuncFlag01=0)
        End If
    Next
    ''End Loop
    
    JustDoIt = IIf(blnOneOfOK, "0,���\", "-8,�I�ƤU�����~")
    
    JustDoIt = JustDoIt & vbCrLf & _
                        "<?xml version=""1.0"" encoding=""BIG5"" standalone=""yes"" ?>" & vbCrLf & _
                        "<DataSet>" & vbCrLf & _
                        "  <DataTable>" & vbCrLf & _
                        strXMLresult & _
                        "  </DataTable>" & vbCrLf & _
                        "</DataSet>" & vbCrLf ' �^��(0, ���\) ��XML��
    
    '01=>ADD �W�[ Credit
    '02=>SUBTRACT ��� Credit
    '03=>SET CREDIT �����]�w Credit ���s��
    '04=>SET BALANCE �M�� Debit, �����]�w Credit ���s��
    '05=>SUB OFFSET �P�ɴ��Debit �PCredit

66:
    On Error Resume Next
    Rlx rs
    Rlx strQry
    Rlx intLoop
    Rlx varData
    Rlx varPara
    Rlx objIPPV
    Rlx lngCustID
    Rlx blnNoData
    Rlx strAuthenticID
    Rlx strOwner
    Set cn = Nothing
    Set objCnPool = Nothing
  Exit Function
ChkErr:
    ErrHandle "mod_Func9", "JustDoIt"
    JustDoIt = "-99," & strErr
End Function

' --2006/07/30 Modify
'  ����XML��.(�̰O���ǰe���O��^�Ǧ^�Ӫ��T�����)
'  �^��(0,���\)��XML��

Private Sub MakeXML(ByVal strSeqNo As String, _
                                    ByVal strFaciSno As String, _
                                    ByVal strSmartCardNo As String, _
                                    ByVal strEPGprice As String, _
                                    ByVal strStatus As String)
  On Error GoTo ChkErr
      ' �]�Ƭy�����BSTB�Ǹ��BICC�Ǹ��B�I��(���)
    strXMLresult = strXMLresult & _
                            "    <DataRow SEQNO=""" & strSeqNo & """" & _
                                " FACISNO=""" & strFaciSno & """" & _
                                " SMARTCARDNO=""" & strSmartCardNo & """" & _
                                " EPGPRICE=""" & strEPGprice & """" & _
                                " STATUS=""" & strStatus & """/>" & vbCrLf
  Exit Sub
ChkErr:
    ErrHandle "mod_Func9", "MakeXML"
End Sub


'Public Function JustDoIt(ByRef objWebPool As Object, _
'                                        ByRef strPara As String) As String
'  On Error GoTo ChkErr
'    Dim strQry As String
'    Dim varPara As Variant
'    Dim varData As Variant
'    Dim objCnPool As Object
'    Dim objIPPV As Object
'
'    Dim rs As Object
'    Dim lngCustID As Long
'    Dim intLoop As Integer
'    Dim strOwner As String
'    Dim strAuthenticID As String
'    Dim blnNoData As Boolean
'
'    Set objCnPool = objWebPool
'    Set cn = objCnPool.GetConnection
'
'    If cn.state <= 0 Then
'        Set cn = objCnPool.GetConnection
'        If cn.state <= 0 Then
'            strErr = "�L�k�s�u��Ʈw!"
'            JustDoIt = "-99,��Ʈw�s�u����!"
'            GoTo 66
'        End If
'    End If
'
'    varPara = Split(strPara, vbCrLf)
'    varPara = Split(varPara(0), ",")
'
'    strAuthenticID = varPara(2)
'    bytComp = Val(Left(strAuthenticID, 3))
'    lngCustID = Val(Right(strAuthenticID, 7))
'    strOwner = GetOwner(cn)
'
'    varData = Split(strPara, vbCrLf)
'
'
'    strQry = "SELECT CUSTID FROM " & strOwner & "SO002B" & Get_DB_Link(cn) & _
'                    "WHERE MEMBERID='" & varPara(1) & "' AND AUTHENTICID='" & strAuthenticID & "'"
'
'    lngCustID = Val(GetValue(cn, strQry) & "")
'
'    If lngCustID = 0 Then
'        JustDoIt = "-2,�q��PPV�{�ҿ��~"
'        GoTo 66
'    End If
'
'    Set objIPPV = CreateObject("setIPPV.clsNagraSTB")
'    With objIPPV
'            Set .uConn = Get_General_CN
'            garyGi(1) = "WEB"
'            .ugaryGi = Join(garyGi, Chr(9))
'    End With
'
'    ''�̶ǨӰѼƩI�sCA�R�O�I�ƤU���Ҳ� (STBCreDold)
'
'    '�I�ƽu�W�q��--�U���I��
'    '9�B�|���N���B�{��ID(���)
'    '�]�Ƭy�����BSTB�Ǹ��BICC�Ǹ��B�I��(���)
'    '�]�Ƭy�����BSTB�Ǹ��BICC�Ǹ��B�I��(���)
'    '�K
'    '0 : ���\
'    '-8 : �I�ƤU�����~
'    '-19 ���ҥ�IPPV�q���v
'
'    ''Loop�B�z (�ĤG��Ѽƶ}�l)
'    For intLoop = 1 To UBound(varData)
'        If RPxx(CStr(varData(intLoop))) <> Empty Then
'            varPara = Split(varData(intLoop), ",")
'            '�ˮֳB�z����STB��iPPV�q���v�O�_�w�}��(IPPVright=1 ?).
'            ''If IPPVright <> 1 Then
'                ''  �I�sCA�R�O�D��ı�ƶ}��iPPV�q���v�Ҳ� (EnIPPVright)
'            ''End If
'            If Not IPPVok(strOwner, CStr(varPara(0)), blnNoData) Then
'                ''�Ұ�IPPV�q���v
'                    'Public Function EnIPPVright(ByVal strCompCode As String, _
'                    '    ByVal lngCustId As Long, ByVal strSmartCard As String, _
'                    '    ByVal strResvTime As String, ByVal strSTB As String, _
'                    '    ByVal strDoCaption As String, ByVal strFaciSNo As String) As Boolean
'                    '   Format(Now, "YYYY/MM/DD HH:MM")
'                If blnNoData Then
'                    JustDoIt = "-99,�]�Ƭy�������s�b"
'                    GoTo 66
'                End If
'                If Not objIPPV.EnIPPVright(CStr(bytComp), _
'                                                    lngCustID, _
'                                                    CStr(varPara(2)), _
'                                                    "", _
'                                                    CStr(varPara(1)), _
'                                                    "", _
'                                                    CStr(varPara(1))) Then
'                    strErr = objIPPV.uMsgBack
'                    JustDoIt = "-19,���ҥ�IPPV�q���v"
'                    GoTo 66
'                End If
'            End If
'
'            ''�U��/�]�wIPPV�I��or (�L�^��)
'            'Public Function STBCreDold(ByVal strCompCode As String, _
'            '    ByVal lngCustId As Long, ByVal strSmartCard As String, _
'            '    ByVal strResvTime As String, ByVal strSTB As String, _
'            '    ByVal strDoCaption As String, ByVal strCreditMode As String, _
'            '    ByVal strCredit As String, ByVal strCollectDate As String, ByVal blnReturnBk As Boolean) As Boolean
'
'            '    ��i�׶}����A�]�ӹL�b�������ä�C ��:
'            '    ResvTime=''
'            '    ��i�׶}����A�]�ӹL�b�������ä�C ��:
'            '    CreditMode='01'
'            '    ��i�׶}����A�]�ӹL�b�������ä�C ��:
'            '   CollectDate = sysdate
'
'            With objIPPV
'                If Not objIPPV.STBCreDold(CStr(bytComp), _
'                                                            lngCustID, _
'                                                            CStr(varPara(2)), _
'                                                            "", _
'                                                            CStr(varPara(1)), _
'                                                            "", _
'                                                            "03", _
'                                                            CStr(Val(varPara(3)) + Get_IPPV_Credit(strOwner, CStr(varPara(0)))), _
'                                                            Format(Now, "YYYY/MM/DD HH:MM"), _
'                                                            ReturnBack(strOwner, CStr(varPara(0)))) Then
'                    strErr = .uMsgBack
'                    JustDoIt = "-8,�I�ƤU�����~"
'                    GoTo 66
'                Else
'                    MakeXML CStr(varPara(0)), CStr(varPara(1)), CStr(varPara(2)), CStr(varPara(3)), .uMsgBack
''                    If .uMsgBack <> Empty Then
''                        JustDoIt = "0,���\(" & .uMsgBack & ")"
''                    Else
''                        JustDoIt = "0,���\"
''                    End If
'                End If
'            End With
'
'            ''If �I�ƤU�����G��ok Then
'                ''  �^��(-8,�I�ƤU�����~)
'            ''Else
'                ''  �^��(0,���\)
'            ''End If
'
'            ''�I�sCA�R�O�D��ı���I�ƤU���Ҳ�(STBCreDold[���^�Ǿ���call]/ STBCreDoldNb[�L�^�Ǿ���call])
'            ''PS:
'            ''���^�Ǿ���w�q : so041. iPPVCallBack=1 and ��STB ModelCode ������ CD043. FuncFlag01=1
'            ''�L�^�Ǿ���w�q : so041. iPPVCallBack=0 or (so041. iPPVCallBack=1 and ��STB ModelCode������CD043. FuncFlag01=0)
'        End If
'    Next
'    ''End Loop
'    If strXMLresult <> Empty Then
'    End If
'    JustDoIt = "0,���\" & vbCrLf & _
'                        "<?xml version=""1.0"" encoding=""BIG5"" standalone=""yes"" ?>" & vbCrLf & _
'                        "<DataSet>" & vbCrLf & _
'                        "  <DataTable>" & vbCrLf & _
'                        strXMLresult & _
'                        "  </DataTable>" & vbCrLf & _
'                        "</DataSet>" & vbCrLf ' �^��(0, ���\) ��XML��
'
'    '01=>ADD �W�[ Credit
'    '02=>SUBTRACT ��� Credit
'    '03=>SET CREDIT �����]�w Credit ���s��
'    '04=>SET BALANCE �M�� Debit, �����]�w Credit ���s��
'    '05=>SUB OFFSET �P�ɴ��Debit �PCredit
'
'66:
'    On Error Resume Next
'    Rlx rs
'    Rlx strQry
'    Rlx intLoop
'    Rlx varData
'    Rlx varPara
'    Rlx objIPPV
'    Rlx lngCustID
'    Rlx blnNoData
'    Rlx strAuthenticID
'    Rlx strOwner
'    Set cn = Nothing
'    Set objCnPool = Nothing
'  Exit Function
'ChkErr:
'    ErrHandle "mod_Func9", "JustDoIt"
'    JustDoIt = "-99," & strErr
'End Function

