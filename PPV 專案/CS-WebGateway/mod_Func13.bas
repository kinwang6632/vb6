Attribute VB_Name = "mod_Func13"

Option Explicit

Private cn As Object
Private rs As Object

Private strAuthenticID As String
Private strOwner As String
Private intLoop As Integer
Private lngCustID As Long
Private strOrderNo As String
Private strOrderItem As String
Private blnOneOfOK As Boolean
Private strXMLresult As String

'�p���I�O�`�ب���(�ݸg�LSMS�q�ʱK�X�{��,�å]�ACA�t�έq�ʱ��v�@�~�γB�z���G)
Public Function JustDoIt(ByRef objWebPool As Object, _
                                        ByRef strPara As String) As String
  On Error GoTo ChkErr
  
    
    Dim varPara As Variant
    Dim varDetail As Variant
    
    Dim strBillData As String
    
    Dim strCancelTime As String
    Dim strReturnCode As String
    Dim strReturnName As String
    
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
'******************************************************************************
    '#3018 �P�_�ǤJ���ӼƬO�_���T By Kin 2007/09/03
    varPara = Split(strPara, vbCrLf)
    
    If UBound(varPara) < 2 Then
        JustDoIt = "-99,�ѼƸ�Ƥ����T ! �нT�{ !"
        GoTo 66
    End If
    varDetail = Split(varPara(0), ",")
    If UBound(varDetail) < 9 Then
        JustDoIt = "-99,�ѼƤ���!"
        GoTo 66
    End If
    For intLoop = 1 To UBound(varPara)
        If (UBound(Split(varPara(intLoop), ",")) < 1) And (varPara(intLoop) <> Empty) Then
            JustDoIt = "-99,�ѼƤ���!"
            GoTo 66
        End If
    Next intLoop
    
'*****************************************************************************
    strAuthenticID = varDetail(2)
    bytComp = Val(Left(strAuthenticID, 3)) '   ���q�O�N�X
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
    
'   Select count(*) as Cnt from so002b where MemberId=[�|���N��] and AuthenticId=[�{��ID] and PPVOrderPwd=[ PPV��I�q�ʱK�X];
    Set rs = cn.Execute("SELECT COMPCODE,CUSTID FROM " & strOwner & "SO002B" & Get_DB_Link(cn) & _
                                    " WHERE MEMBERID='" & varDetail(1) & "'" & _
                                    " AND AUTHENTICID='" & strAuthenticID & "'" & _
                                    " AND PPVORDERPWD='" & varDetail(3) & "'")
    
    If rs.AbsolutePage <= 0 Then '    If Cnt = 0 Then
        JustDoIt = "-4,�q��PPV�{�ҿ��~" '       �^��(-4,�q��PPV�{�ҿ��~)
        GoTo 66
    End If '    End If
    
    bytComp = Val(rs("COMPCODE") & "")
    lngCustID = Val(rs("CUSTID") & "")
    
    strCancelTime = varDetail(5) & ""
    strReturnCode = varDetail(8) & ""
    strReturnName = varDetail(9) & ""

    cn.BeginTrans

    '   �̶ǤJ���Ѽ�(ICC�Ǹ��B���~�N�X)
    '   �I�sCA�R�O[���� PPV �q�ʸ`��]�D��ı�ƼҲ�. �����N��P31
    '  13�B�|���N���B�{��ID�BPPV��I�q�ʱK�X�B2(��ӷ�����)�B�����ɶ�(YYYYMMDD HH24MI)�BWEB(�����H���N��)�B��������(�����H��)�B�q�������]�N�X�B�q�������](���)
    
    blnOneOfOK = False
    For intLoop = 1 To UBound(varPara) 'Loop�B�z (�ĤG��}�l)
        
        strBillData = varPara(intLoop) & ""
        '   �q��渹�B�q�涵��(���)
        
        If strBillData <> Empty Then
            
            varDetail = Split(strBillData, ",")
            
            strOrderNo = varDetail(0) & ""
            strOrderItem = varDetail(1) & ""
            
            '  Select count(*) as Cnt from so105g
            '   where Orderno=[�q��渹] and Orderitem=[�q�涵��]
            '   and Accounting=0 and canceltime is null ;
            If GetValue(cn, "SELECT COUNT(*) FROM " & strOwner & "SO105G" & Get_DB_Link(cn) & _
                                            " WHERE ORDERNO='" & strOrderNo & "' AND ORDERITEM=" & strOrderItem & _
                                            " AND ACCOUNTING=0 AND CANCELTIME IS NULL") = 0 Then '   If Cnt = 0 Then
                ' �O�������\���q��渹�B�q�涵���BSTB�Ǹ��BICC�d���B�`�ئW�١B"�w����Τw����"
                If Not UnknowStatus("�w����Τw����") Then
                    JustDoIt = "-99,[PPV�`�حq�ʩ�����]�d�L�q����!"
                    GoTo RollBack
                End If
            Else '   Else
                '       �̶ǤJ���Ѽ�(ICC�Ǹ��B���~�N�X)�I�sCA�R�O[���� PPV �q�ʸ`��]�D��ı�ƼҲ�. �����N��P31
                If CallNagraProcess(strMsg) Then  '       If �R�O���\ Then
                    ' Update so105g set CancelEn=[�����H���N��], CancelName=[�����H��], CancelTime=[�����ɶ�],
                    '   ReturnCode=[�q�������]�N�X], ReturnName=[�q�������]]
                    '   Where orderno=[�q��渹] and orderitem=[�q�涵��];
                    blnOneOfOK = True
                    cn.Execute "UPDATE " & strOwner & "SO105G" & Get_DB_Link(cn) & _
                                        " SET CANCELEN='WEB',CANCELNAME='��������'" & _
                                        ",CANCELTIME=TO_DATE('" & strCancelTime & "','YYYYMMDD HH24MI')" & _
                                        ",RETURNCODE=" & strReturnCode & _
                                        ",RETURNNAME='" & strReturnName & "'" & _
                                        " WHERE ORDERNO='" & strOrderNo & "' AND ORDERITEM=" & strOrderItem
                    
                    '   �|���N���B�{��ID�BPPV��I�q�ʱK�X�B2(��ӷ�����)�B�����ɶ�(YYYYMMDD HH24MI)�BWEB(�����H���N��)�B��������(�����H��)�B
                    ' �q�������]�N�X�B�q�������](���)
                    
                    ' select count(*) as Cnt from so105g where ordreno=[�q��渹] and canceltime is null;
                    If GetValue(cn, "SELECT COUNT(*) FROM " & strOwner & "SO105G" & Get_DB_Link(cn) & _
                                            " WHERE ORDERNO='" & strOrderNo & "' AND CANCELTIME IS NULL") = 0 Then ' If Cnt = 0 Then
                        
                        '  Update so105e set CancelEn=[�����H���N��], CancelName=[�����H��], CancelTime=[�����ɶ�],
                        '   ReturnCode=[�q�������]�N�X], ReturnName=[�q�������]] Where orderno =[�q��渹] ;
                        cn.Execute "UPDATE " & strOwner & "SO105E" & Get_DB_Link(cn) & _
                                            " SET CANCELEN='WEB',CANCELNAME='��������'" & _
                                            ",CANCELTIME=TO_DATE('" & strCancelTime & "','YYYYMMDD HH24MI')" & _
                                            ",RETURNCODE=" & strReturnCode & _
                                            ",RETURNNAME='" & strReturnName & "'" & _
                                            " WHERE ORDERNO='" & strOrderNo & "'"
                    
                    End If ' End If
                'Else
                    ' �O�������\���q��渹�B�q�涵���BSTB�Ǹ��BICC�d���B�`�ئW�١B"��������"
                    ' UnknowStatus varDetail(0) & "", varDetail(1) & "", "��������"
                End If
            
            End If '  End If
            
        End If
        
    Next '   End Loop
    
'    If strMsg <> Empty Then
'        JustDoIt = "-99, �ҥ~���~ !" & strMsg
'        GoTo RollBack '  RollBack;
'    Else
        If Not blnOneOfOK Then 'If �Ҧ����������`�ةI�sCA�R�O(P31)�������� THEN
            JustDoIt = "-12, �p���I�O�`�ب����q�ʥ���" & vbCrLf & GetXML '  �^��(-12, �p���I�O�`�ب����q�ʥ���)��XML��
            GoTo RollBack '  RollBack;
        Else 'Else
            cn.CommitTrans
            JustDoIt = "0,���\" & vbCrLf & GetXML '  �^��(0,���\)��XML��"
            GoTo 66
        End If 'End If
'    End If
    
    'Ps�K.
    'XML���e�p�U:
    '�q��渹�B�q�涵���BSTB�Ǹ��BICC�d���B�`�ئW�١B���A
    
RollBack:
    cn.RollbackTrans
    
66:
    On Error Resume Next
    Rlx strCancelTime
    Rlx strReturnCode
    Rlx strReturnName
    Rlx rs
    Rlx cn
    Rlx varPara
    Rlx varDetail
    Rlx strOwner
    Rlx strAuthenticID
    Rlx intLoop
    Rlx lngCustID
    Rlx blnOneOfOK
    Rlx strXMLresult
    Rlx strOrderNo
    Rlx strOrderItem
    Rlx strBillData
    Rlx strMsg
    Set cn = Nothing
    
  Exit Function
ChkErr:
    ErrHandle "mod_Func13", "JustDoIt"
    cn.RollbackTrans
    JustDoIt = "-99," & strErr
End Function

Private Function GetXML() As String
  On Error GoTo ChkErr
    GetXML = "<?xml version=""1.0"" encoding=""BIG5"" standalone=""yes"" ?>" & vbCrLf & _
                    "<DataSet>" & vbCrLf & _
                    "  <DataTable>" & vbCrLf & _
                    strXMLresult & _
                    "  </DataTable>" & vbCrLf & _
                    "</DataSet>" & vbCrLf
  Exit Function
ChkErr:
    ErrHandle "mod_Func13", "JustDoIt"
End Function

Private Function CallNagraProcess(ByRef strMsg As String) As Boolean
  On Error GoTo ChkErr
    Dim objNS As Object
    Dim intErrPos As Integer
    '   strNotes= "ProductId" & "~" & "ProductName" & "~" & "CreditAmt"
    intErrPos = 0
    Set rs = cn.Execute("SELECT FACISNO,SMARTCARDNO,PRODUCTID,PRODUCTNAME" & _
                                    " FROM " & strOwner & "SO105G" & Get_DB_Link(cn) & _
                                    " WHERE ORDERNO='" & strOrderNo & "' AND ORDERITEM=" & strOrderItem)
'    Set rs = cn.Execute("SELECT FACISNO,SMARTCARDNO,PRODUCTID,PRODUCTNAME,CREDITAMT" & _
                                    " FROM " & strOwner & "SO105G" & Get_DB_Link(cn) & _
                                    " WHERE ORDERNO='" & strOrderNo & "' AND ORDERITEM=" & strOrderItem)
    intErrPos = 1
    Set objNS = CreateObject("setIPPV.clsNagraSTB")
    With objNS
        intErrPos = 2
        strMsg = ""
        intErrPos = 3
        .uCompCode = bytComp
        intErrPos = 4
        Set .uConn = cn
        intErrPos = 5
        garyGi(1) = "WEB"
        intErrPos = 6
        .ugaryGi = Join(garyGi, Chr(9))
        intErrPos = 7
        .uErrPath = Environ("Temp")
        intErrPos = 8
        CallNagraProcess = .CancelPPVPro(bytComp, lngCustID, rs!SMARTCARDNO & "", "", rs!FaciSno & "", "����PPV�`�س]�w", _
                                                                rs!ProductID & "")
'        CallNagraProcess = .CancelPPVPro(bytComp, lngCustID, rs!SmartCardNo & "", "", rs!FaciSno & "", "����PPV�`�س]�w", _
                                                                rs!ProductID & "~" & rs!ProductName & "~" & rs!CreditAmt)
        intErrPos = 9
        If CallNagraProcess Then
            intErrPos = 10
            MakeXML "�������\", rs!FaciSno & "", rs!SMARTCARDNO & "", rs!ProductName & ""
        Else
            ' �O�������\���q��渹�B�q�涵���BSTB�Ǹ��BICC�d���B�`�ئW�١B"��������"
            intErrPos = 11
            MakeXML "��������", rs!FaciSno & "", rs!SMARTCARDNO & "", rs!ProductName & ""
        End If
        intErrPos = 12
        strMsg = .uMsgBack
        intErrPos = 13
    End With
    Set objNS = Nothing
  Exit Function
ChkErr:
    strMsg = "[STB �]�w���~] - " & intErrPos & Err.Description
    On Error Resume Next
    Set objNS = Nothing
End Function

Private Function UnknowStatus(strStatus As String) As Boolean
  On Error GoTo ChkErr
    Set rs = cn.Execute("SELECT FACISNO,SMARTCARDNO,PRODUCTNAME FROM " & strOwner & "SO105G" & Get_DB_Link(cn) & _
                                    " WHERE ORDERNO='" & strOrderNo & "' AND ORDERITEM=" & strOrderItem)
    ' �O�������\���q��渹�B�q�涵���BSTB�Ǹ��BICC�d���B�`�ئW�١B"�w����Τw����"
    UnknowStatus = rs.RecordCount > 0
    If UnknowStatus Then MakeXML strStatus, rs!FaciSno & "", rs!SMARTCARDNO & "", rs!ProductName & ""
  Exit Function
ChkErr:
    ErrHandle "mod_Func13", "UnknowStatus"
End Function

Private Sub MakeXML(strStatus As String, _
                                    strFaciSno As String, _
                                    strSmartCardNo As String, _
                                    strProductName As String)
  On Error GoTo ChkErr
    strXMLresult = strXMLresult & _
                            "    <DataRow ORDERNO=""" & strOrderNo & """" & _
                                " ORDERITEM=""" & strOrderItem & """" & _
                                " FACISNO=""" & strFaciSno & """" & _
                                " SMARTCARDNO=""" & strSmartCardNo & """" & _
                                " PRODUCTNAME=""" & RPC(strProductName) & """" & _
                                " STATUS=""" & strStatus & """/>" & vbCrLf
  Exit Sub
ChkErr:
    ErrHandle "mod_Func13", "MakeXML"
End Sub


