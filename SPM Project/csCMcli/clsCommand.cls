VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCommand"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Private cn As Object
Private strQry As String

'CmdDisp = Switch(Value = "C1", "�˾� (CM)", Value = "E1", "�˾� (EMTA)", _
'                                Value = "C2", "��� (CM�˾��h��)", Value = "E2", "��� (EMTA�˾��h��)", _
'                                Value = "E3", "CP �˾�", Value = "A1_P", "CP Only ���", _
'                                Value = "E4", "EMTA �[�� CP ����", Value = "E5", "EMTA �� CP ����", _
'                                Value = "A1_D", "�n��", Value = "A1_C", "�n�_", _
'                                Value = "A1_BU", "�t�v�ɭ���", Value = "A1_PIP", "�ӽаʺAIP", _
'                                Value = "A1_MIP", "�����ʺAIP", Value = "Q1", "CM reset", _
'                                Value = "Q2", "�d��CM���p", Value = "A7", "CM ���� / ���s", _
'                                Value = "A8", "EMTA ���� / ���s", Value = "A9", "����CM / ���sEMTA", _
'                                Value = "A10", "����EMTA / ���s CM", Value = "A6", "�ܧ� CPE MAC", _
'                                Value = "A5", "CM IP  CPE IP �]�w", Value = "A4", "EMTA IP Priv IP �]�w", _
'                                Value = "A3", "�����T�wIP", Value = "A2", "�ӽЩT�wIP", _
'                                Value = "C6", "CM ����", Value = "E6", "EMTA ����")

'If Not blnTransation Then gcnGi.BeginTrans
'Public strResvTime As String

Public Function ExecCmd(ByRef Gcn As Object, ByRef rs04 As Object, CmdID As String, _
                                        AryGi As String, ErrPath As String, Optional ResvTime As String = "", _
                                        Optional Transaction As Boolean = True, Optional RetMsg As String = "", _
                                        Optional FaciCode As String = "", Optional FaciName As String = "", _
                                        Optional FaciSno As String = "", Optional SNO As String = "", _
                                        Optional MediaBillNo As String = "", Optional CmdSeqNo As String = "", _
                                        Optional BaudRate As String = "", Optional ProbStopDate As String = "", _
                                        Optional ClearSNo As Boolean = False, _
                                        Optional ByRef rsNew As ADODB.Recordset = Nothing, _
                                        Optional ByVal strNewBaud As String = Empty, _
                                        Optional ByVal strNewDynaIPcnt As String = Empty, _
                                        Optional ByVal strNewFixIPcnt As String = Empty, _
                                        Optional ByVal strNewIPAddress As String = Empty, _
                                        Optional ByVal strNewModel As String = Empty, _
                                        Optional ByVal strNewFaci As String = Empty) As Boolean
  On Error GoTo ChkErr
  
    ExecCmd = False
    UI_Mode = False
    blnNoShowMsg = False
    RetMsg = ""
    strRetMsg = ""
    
    If Gcn Is Nothing Then RetMsg = "�жǤJConnection �s�u���� !": Exit Function
    If Gcn.State <= 0 Then RetMsg = "�нT�{ Connection �s�u����w�}�� !": Exit Function
    If rs04 Is Nothing Then RetMsg = "�жǤJ SO004 Recordset ��ƿ������� !": Exit Function
    If rs04.State <= 0 Then RetMsg = "�нT�{ SO004 Recordset ��ƿ����w�}�� !": Exit Function
    If rs04.EOF Then RetMsg = "�нT�{ SO004 Recordset ��ƿ���������� !": Exit Function
    If Len(AryGi) = 0 Then RetMsg = "�жǤJ GaryGi �r�� !": Exit Function
    If Len(ErrPath) = 0 Then RetMsg = "�жǤJ ErrPath ���| ! ( Error Log Path )": Exit Function
    
    Dim intComp As Integer
    Dim lngCustId As Long
    Dim SvcTp As String
    Dim rsTmp As New ADODB.Recordset
    
    'Set cn = Gcn
    Set gcnGi = Gcn
    GetGlobal AryGi
    gErrLogPath = ErrPath
    
    Set objCmdCN = CreateObject("ADODB.Connection")
    With objCmdCN
        .CursorLocation = 3
        .Open gcnGi.ConnectionString & ";Persist Security Info=True"
    End With
    '#6066 �p�GrsNew���ǭ�,�hrsUPD=rsNew By Kin 2011/07/26
    If strNewBaud <> Empty Then strNewBaudRate = strNewBaud
    If strNewDynaIPcnt <> Empty Then strDynaIPcnt = strNewDynaIPcnt
    If strNewFixIPcnt <> Empty Then strFixIPcnt = strNewFixIPcnt
    If strNewIPAddress <> Empty Then strNewIP = strNewIPAddress
    If strNewModel <> Empty Then strNewModelCode = strNewModel
    If strNewFaci <> Empty Then strNewFaciSno = strNewFaci
    
'    Set rsUpd = Nothing
'    If Not rsNew Is Nothing Then
'        If rsNew.State = adStateOpen Then
'            If rsNew.RecordCount > 0 Then
'                Set rsUpd = rsNew
'            End If
'        End If
'    End If
    
    
    intComp = Val(rs04!CompCode & "")
    lngCustId = Val(rs04!CustID & "")
    SvcTp = rs04!ServiceType & ""
    
    If ResvTime <> "" Then
        If InStr(1, ResvTime, "/") <= 0 Then ResvTime = Format(ResvTime, "####/##/## ##:##:##")
    End If
    
    If ProbStopDate <> "" Then
        If InStr(1, ProbStopDate, "/") <= 0 Then ProbStopDate = Format(ProbStopDate, "####/##/##")
    End If
    
    strResvTime = ResvTime
    strProbStopDate = ProbStopDate
    
    CmdID = UCase(CmdID)
    
    If intComp = 0 Then RetMsg = "SO004 ��Ƥ� [���q�O] �� 0 �����T !": Exit Function
    If lngCustId = 0 Then RetMsg = "SO004 ��Ƥ� [�Ƚs] �� 0 �����T !": Exit Function
    If Len(SvcTp) = 0 Then RetMsg = "SO004 ��Ƥ� [�A�ȧO] ���� �����T !": Exit Function

    gCompFilterStr = intComp
    gDefaultComp = intComp
    gCompCode = intComp
    gServiceType = SvcTp
    
    strSno = SNO
    strMediaBillNo = MediaBillNo
    strBaudRate = BaudRate
    
    int_EMTA_IP_Type = Val(GetSystemParaItem("EMTAIPTYPE", CStr(intComp), SvcTp, "SO041", , True) & "")
    
    strQry = "SELECT INSTADDRNO FROM " & GetOwner & "SO001" & _
                    " WHERE CUSTID = " & lngCustId & " AND COMPCODE = " & intComp
    
    If Not GetRS(rsTmp, strQry, gcnGi, , , , , , , True) Then
        RetMsg = "SO001 �d�� [���q�O] �� " & intComp & " , [�Ƚs] �� " & lngCustId & " �� [�a�}�s��] ��ƥ��� !"
        Exit Function
    End If
    
    If rsTmp.EOF Then
        RetMsg = "SO001 �d�L [���q�O] �� " & intComp & " , [�Ƚs] �� " & lngCustId & " �� [�a�}�s��] ��� !"
        Exit Function
    Else
        lngInstAddrNo = rsTmp!InstAddrNo & ""
    End If
    
    CloseRecordset rsTmp
    
    If lngInstAddrNo = 0 Then RetMsg = "�˾� [�a�}�s��] �� 0 �����T !": Exit Function
    
    GetSysPara
    
    strPrcType = CmdID
    
    If CmdSeqNo = "" Then
'       #3327 ��X�y�{�� , ���� IsDataOK
'        If Not IsDataOK(rs04, CmdID, FaciCode, FaciSno) Then
'            RetMsg = GetMsg(RetMsg)
'            Exit Function
'        End If
    Else
        GetRS307 , CmdSeqNo
        If rs307.State <= 0 Or rs307.EOF Then
            RetMsg = "SO307 �d�L [�Ǹ�] �� " & CmdSeqNo & " ����� !"
            Exit Function
'        Else
'            With rs307
'                ExecCmd = ParaClt(rs04, CmdID, , , , lngCustId, intComp, _
'                                                !NewModemMac & "", _
'                                                !NewModelName & "", _
'                                                !NewCMBaudRate & "", _
'                                                Val(!NewDynIPCount & ""), _
'                                                Val(!NewFixIPCount & ""), _
'                                                !NewHFCNode & "", _
'                                                !AddCPEMac & "", !DelCPEMac & "", _
'                                                !AddCPEStatIP & "", !DelCPEStatIP & "", _
'                                                !NewIPaddress & "", _
'                                                !NewMTAMAC & "", , , , CmdSeqNo)
'            End With
'            RetMsg = GetMsg(RetMsg)
'            Exit Function
        End If
    End If
    
    ' C6 �R�O�L Update SO004 �� Table , �ǩ� , �T�{�� ..
    Select Case CmdID
        
        ' �˾�,���,�n��,�n�_,CP Only �˾�,CP Only ���,CP�ˤ��u,CP����u
        Case "C1", "E1", "C2", "E2", "A1_D", "A1_C", "E3", "A1_P", "E4", "E5"
            ExecCmd = ParaClt(rs04, CmdID, , , , lngCustId, intComp, , , , , , , , , , , , , , , , CmdSeqNo)
        
        ' 6. EMTA IP ����
        ' 1.  �ѦҸ�2�ɶǰeC6 ���O�A�ª�NODE ���θ˾��a�}��T�A�sNODE ���εe����J�β��ˤu�檺�s�a�}��T
        ' 2.  �ѦҸ�5�ɶǰeE6 ���O�A�ª�NODE ���θ˾��a�}��T�A�sNODE ���εe����J�β��ˤu�檺�s�a�}��T
        ' 3.  ���Ŀ�ɫ��ӭ����ǰeA4 ���O
        Case "A4", "E6" ' EMTA IP ����
            ExecCmd = ParaClt(rs04, CmdID, , , , lngCustId, intComp, , , , , , rs307!NewHFCNode & "", , , , , rs307!NewIPaddress & "", , , , , CmdSeqNo)
        
        Case "A6" ' CPE MAC ����
            ExecCmd = ParaClt(rs04, CmdID, , , , lngCustId, intComp, , , , , , , rs307!AddCPEMac & "", , , , , , , , , CmdSeqNo)
            
        Case "A5" ' CPE IP ����
            ExecCmd = ParaClt(rs04, CmdID, , , , lngCustId, intComp, , , , , , , , , rs307!AddCPEStatIP & "", , , , , , , CmdSeqNo)
            
        Case "A2" ' �ӽЩT�wIP
            ExecCmd = ParaClt(rs04, CmdID, , , , lngCustId, intComp, , , , , Val(rs307!NewFixIPCount & ""), , rs307!AddCPEMac & "", , rs307!AddCPEStatIP & "", , , , , , , CmdSeqNo)
            
        Case "A3" ' �ӽЩT�wIP
            ExecCmd = ParaClt(rs04, CmdID, , , , lngCustId, intComp, , , , , Val(rs307!NewFixIPCount & ""), , , rs307!DelCPEMac & "", , rs307!DelCPEStatIP & "", , , , , , CmdSeqNo)
            
        Case "A1_MIP", "A1_PIP" ' �����ʺAIP , �ӽаʺAIP
            ExecCmd = ParaClt(rs04, CmdID, , , , lngCustId, intComp, , , , Val(rs307!NewDynIPCount & ""), , , , , , , , , , , , CmdSeqNo)
        
        Case "RP", "A7", "A8", "A9", "A10" ' CM / EMTA ��
            If CmdSeqNo = Empty Then
                If Len(FaciCode) > 0 And Len(FaciSno) > 0 Then
                    strFC = FaciCode & ""
                    strFN = FaciName & ""
                    strFCsno = FaciSno
                    If CmdID = "RP" Then
                        ExecCmd = ParaClt(rs04, Get_RP_Cmd(rs04, FaciCode), , , , lngCustId, intComp, FaciSno, GetModelName(FaciSno), , , , , , , , , , , , , , CmdSeqNo)
                    Else
                        ExecCmd = ParaClt(rs04, CmdID, , , , lngCustId, intComp, FaciSno, GetModelName(FaciSno), , , , , , , , , , , , , , CmdSeqNo)
                    End If
                Else
                    RetMsg = "CM / EMTA �󴫩R�O , �нT�{ [FaciSno] �� [FaciCode] �ѼƦ��� !"
                End If
            Else
                FaciSno = Replace(rs307!NewModemMac & "", ":", "", 1)
                strFCsno = FaciSno
                GetFaciName lngCustId
                ExecCmd = ParaClt(rs04, CmdID, , , , lngCustId, intComp, FaciSno, GetModelName(FaciSno), , , , , , , , , , , , , , CmdSeqNo)
            End If
        
        Case "A1_BU", "TRIAL", "STOPTRIAL"
                ' "STOPTRIAL" BaudRate = ""
                If CmdID <> "STOPTRIAL" Then
                    If BaudRate = Empty Then
                        If CmdSeqNo <> "" And rs307.State > 0 And Not rs307.EOF Then
                            ExecCmd = ParaClt(rs04, CmdID, , , , lngCustId, intComp, , , rs307!NewCMBaudRate & "", , , , , , , , , , , , , CmdSeqNo)
                        Else
                            AddRetMsg "�жǤJ [ �s�t�v ] !"
                            ExecCmd = False
                        End If
                    Else
                        ExecCmd = ParaClt(rs04, CmdID, , , , lngCustId, intComp, , , Get_BaudRate_Desc(BaudRate), , , , , , , , , , , , , CmdSeqNo)
                    End If
                Else
                    ExecCmd = ParaClt(rs04, CmdID, , , , lngCustId, intComp, , , , , , , , , , , , , , , , CmdSeqNo)
                End If
        
        Case ""
            RetMsg = "�жǤJ�R�O�N�� !"
        
        Case Else
            With rs307
                ExecCmd = ParaClt(rs04, CmdID, , , , lngCustId, intComp, _
                                                !NewModemMac & "", _
                                                !NewModelName & "", _
                                                !NewCMBaudRate & "", _
                                                Val(!NewDynIPCount & ""), _
                                                Val(!NewFixIPCount & ""), _
                                                !NewHFCNode & "", _
                                                !AddCPEMac & "", !DelCPEMac & "", _
                                                !AddCPEStatIP & "", !DelCPEStatIP & "", _
                                                !NewIPaddress & "", _
                                                !NewMTAMAC & "", , , , CmdSeqNo)
            End With
'            RetMsg = "�R�O�N�����~ !"
        
'        Case "A1_BU"
'            If Len(BuadRate) > 0 Then
'                ExecCmd = ParaClt(rs04, "A1_BU", , , , lngCustId, intComp, , , BuadRate)
'            Else
'                RetMsg = "�t�v�ɭ��ũR�O , �нT�{ [�s�t�v] �ѼƦ��� !"
'            End If
        
    
    End Select
    
    If Not ExecCmd And ClearSNo Then
        objCmdCN.Execute "UPDATE " & strCMowner & "SO307 SET SNO=NULL" & _
                                        " WHERE CMDSEQNO='" & CmdSeqNo & "'"
    End If
    
    RetMsg = GetMsg(RetMsg)
    CmdSeqNo = strRetCmdSeqNo
  Exit Function
ChkErr:
    ErrSub "clsCommand", "ExecCmd"
End Function

'Private Function Get_BaudRate_Code(ByRef BaudRateDesc As String) As String
'  On Error GoTo ChkErr
'    strQry = "SELECT CODENO FROM " & GetOwner & "CD064 WHERE DESCRIPTION='" & BaudRateDesc & "'"
'    Get_BaudRate_Code = GetRsValue(strQry, gcnGi) & ""
'  Exit Function
'ChkErr:
'    ErrSub "clsCommand", "Get_BaudRate_Code"
'End Function



Private Function GetMsg(M As String) As String
  On Error GoTo ChkErr
    If M <> Empty Then
        If strRetMsg <> Empty Then M = M & vbCrLf & strRetMsg
    Else
        M = strRetMsg
    End If
    GetMsg = Replace(M, strCrLf, vbCrLf, 1)
  Exit Function
ChkErr:
    ErrSub "clsCommand", "GetMsg"
End Function

Private Function IsDataOK(ByRef rs04 As Object, CmdID As String, _
                                            Optional FaciCode As String = "", Optional FaciSno As String = "") As Boolean
                                            
  On Error GoTo ChkErr
  
    Dim strCmd As String
    IsDataOK = False
    
    If strResvTime <> "" Then
        If CDate(strResvTime) <= RightNow Then
            AddRetMsg "[�w���ɶ�] �ݤj��{�b�ɶ�"
            Exit Function
        Else
'            AddRetMsg "�]�w�w���ɶ�, �N������ Gateway �^�� !"
        End If
    End If
    
    With rs04
    
        Select Case CmdID
        
            Case "E1", "C1" ' 1. �˾�
                ' �]�Ʃ������O�ťջP�}����� �O�ť� �~�i�H�˾����O ���\��~�i�HENABLE
                If !CMOpenDate & "" <> "" Or !PRdate & "" <> "" Then
                    AddRetMsg "�]�� [�}�����] �P [������] �L�Ȥ~�i�˾� !"
                    Exit Function
                End If
                
                Dim blnChkAddrNo As Boolean
                
                GetHFCNode !CustID, !CompCode, !ServiceType & "", blnChkAddrNo
                
                If blnChkAddrNo Then
                    AddRetMsg "�ӫȤᲾ�����A�����w�s�}�A�Ы��w�s�}��A�e�}�����O�I"
                    Exit Function
                End If
                
            Case "E2", "C2" ' 2. ��� ( �t�˾��h�� )
                ' �]�Ƹ˾�����D�ť� �������}������~�i�H�@������O ���\��~�i�HENABLE
                ' If !InstDate & "" = "" Or !CMOpenDate & "" = "" Then
                If !CMOpenDate & "" = "" Then
                    AddRetMsg "�]�� [�˾����] �P [�}�����] ���Ȥ~�i��� !"
                    Exit Function
                End If
                
            Case "A1_D" ' 3. �n��
                ' ���}���� �B�}����>(������OR �L������)�h�u�i��n��
                ' ���`�ϥΤ��B�w�ˤ�����EMTA �PCM �ӵ���ƥ����]���}����� �P ��������O�Ū��^ �Ρ]�}�����>��������^���\��~�i�HENABLE
                If !CMOpenDate & "" = "" Or !CMCloseDate & "" <> "" Then
                    If Not IsNull(!CMCloseDate) And Not IsNull(!CMOpenDate) Then
                        If CDT(!CMCloseDate) > CDT(!CMOpenDate) Then
                            AddRetMsg "�]�� [���`�ϥΤ�] �~�i�n�� !"
                            Exit Function
                        End If
                    Else
                        AddRetMsg "�]�� [���`�ϥΤ�] �~�i�n�� !"
                        Exit Function
                    End If
                End If
                
            Case "A1_C" ' 4. �n�_
                ' �n���P�n�_��U�ﶵ�O�����A�]�ƭY(������OR �L������) > �}���� �h�u�i�H��n�_
                ' ���`�ϥΤ��B�w�ˤ�����EMTA �PCM �ӵ���ƥ����]����������^ �Ρ]�}�����<��������^���\��~�i�HENABLE
                If !CMOpenDate & "" <> "" Or !CMCloseDate & "" = "" Then
                    If Not IsNull(!CMCloseDate) And Not IsNull(!CMOpenDate) Then
                        If CDT(!CMCloseDate) < CDT(!CMOpenDate) Then
                            AddRetMsg "�]�� [���`�ϥΤ�] �~�i�n�_ !"
                            Exit Function
                        End If
                    Else
                        AddRetMsg "�]�� [���`�ϥΤ�] �~�i�n�_ !"
                        Exit Function
                    End If
                End If
                
            Case "E3" ' 5. CP �˾�
                ' �]�ƥ��`�ϥΤ��B�w�ˤ�EMTA �}������j���������  �B���p��CP �A�Ȼݦ�CP ���������A�B�Ӫ������w�ˤ��Υ��`���\��~�i�HENABLE
                If Not Is_EMTA_Dev(rs04) Then
                    AddRetMsg "���]�ƫD EMTA , �L�k�i�� CP �˾� !"
                    Exit Function
                End If
                
                ' #4302,2009/01/05 By Hammer,(RA)_SO1171A_20081223_�վ�CLI����xCP�����P�_.doc
                If IsCP(rs04) Then
                    If Not IsCMCP(rs04) Then
                        AddRetMsg "�L [ �w�ˤ� ] �� [ ���` ] �� [ CP ���� ] , �L�k�i�� CP �˾� !"
                        Exit Function
                    End If
                End If
'                IsCMCP = (GetRsValue("SELECT COUNT(*) FROM " & GetOwner & "SO004" & _
                                                        " WHERE REFACISNO='" & rsData("FaciSno") & "'" & _
                                                        " AND COMPCODE=" & rsData("CompCode") & _
                                                        " AND CUSTID=" & rsData("CustID") & _
                                                        " AND SERVICETYPE='P'" & _
                                                        " AND PRDATE IS NULL") = 1)
                If Not IsNull(!CMCloseDate) And Not IsNull(!CMOpenDate) Then
                    If CDT(!CMCloseDate) > CDT(!CMOpenDate) Then
                        AddRetMsg "�L [ �w�ˤ� ] �� [ ���` ] �� [ CP ���� ] , �L�k�i�� CP �˾� !"
                        Exit Function
                    End If
                End If
                ' �Х[�j��Ʒj�M�ɱ��H�ӫȽs�ۦP����Ƭ��D�]custid = �]�w�� custid�^�A
                ' gateway ���ȮɽмW�[group by Gateway �A
                ' �p�G���X���ⵧ�]�t�^�H�W�ɻݭn���Xĵ�ܰT���]��emta �������O�ݩ󤣦P mediaGateway �L�k�]�w�^�A�������O�ǰe
                If Val(GetRsValue("SELECT COUNT(GATEWAY) FROM " & GetOwner & "SO004" & _
                                    " WHERE REFACISNO='" & !FaciSno & "'" & _
                                    " AND CUSTID=" & !CustID & _
                                    " AND COMPCODE=" & !CompCode & _
                                    " AND SERVICETYPE='P'" & _
                                    " AND PRDATE IS NULL" & _
                                    " And FaciCode In " & _
                                    " (Select CodeNo From " & GetOwner & "CD022 Where Nvl(CPImportMode,0) In (0,1))" & _
                                    " GROUP BY GATEWAY", gcnGi, "") & "") > 1 Then
                    AddRetMsg "�� EMTA �������O�ݩ󤣦P Media Gateway �L�k�]�w !"
                    Exit Function
                End If
                
            Case "A1_P" ' 6. CP ���
                ' �]�ƥ��`�ϥΤ��B�w�ˤ� EMTA �}������j��������� �B���p��P �A�� CP �������������n���O����� ���\��~�i�HENABLE
                If Not Is_EMTA_Dev(rs04) Then
                    AddRetMsg "���]�ƫD EMTA , �L�k�i�� CP ��� !"
                    Exit Function
                End If
                If IsCP(rs04) Then
                    If Not IsCMCP(rs04) Then
                        AddRetMsg "�L [ �w�ˤ� ] �� [ ���` ] �� [ CP ���� ] , �L�k�i�� CP ��� !"
                        Exit Function
                    End If
                End If
                
            Case "RP" ' CM / EMTA ��
                If Len(FaciCode) = 0 Then
                    AddRetMsg "[�~�W] �����n���,������ !!"
                    Exit Function
                End If
                If Len(Trim(FaciSno)) = 0 Then
                    AddRetMsg "[CM �Ǹ�] �����n���,������ !!"
                    Exit Function
                Else
                    ' 2006/06/21 ��J�s�Ǹ��ɻݭn�ˮָӧǸ��bSO004 ���O�_������(�S����������)�κ⭫�Ƥ��i�@�]�w
                    Dim strFaciSno As String
                    strFaciSno = FaciSno
                    strFaciSno = Replace(strFaciSno, "-", "", 1)
                    strFaciSno = Replace(strFaciSno, ":", "", 1)
                    strCmd = Get_RP_Cmd(rs04, FaciCode)
                    If strFaciSno = !FaciSno & "" Or strFaciSno = !ReFaciSno & "" Then
                        AddRetMsg "[CM �Ǹ�] �P�§Ǹ��ۦP .. �Э��s��J�s�Ǹ� !"
                        Exit Function
                    Else
                        ' Jacky ������x�A�P�_�N�n
                        ' " OR REFACISNO='" & strFaciSno & "' )"
'                        If GetRsValue("SELECT COUNT(*) FROM " & GetOwner & "SO004" & _
'                                                " WHERE FACISNO='" & strFaciSno & "' AND PRDATE IS NULL") > 0 Then
'                            AddRetMsg "���Ǹ��w�Q�ϥ� , �Э��s��J !"
'                            Exit Function
'                        End If
                        If strCmd = "A8" Or strCmd = "A9" Then
                            If GetMTAMAC(strFaciSno) = Empty Then
                                AddRetMsg "���Ǹ� [ " & strFaciSno & " ] �J�w��Ƨ䤣�� !"
                                Exit Function
                            End If
                        End If
                    End If
                    
                    ' EMTA �� CM , �n�ˮ� , �p�G ReFaciSno ������ , �h Show Msg
                    If strCmd = "A10" Then
                        If GetRsValue("SELECT COUNT(*) FROM " & GetOwner & "SO004" & _
                                                " WHERE REFACISNO='" & !FaciSno & "'" & _
                                                " AND PRDATE IS NULL AND SERVICETYPE='P'") > 0 Then
                            AddRetMsg "�� CP ���� , ���i���� CM "
                            Exit Function
                        End If
                    End If
                End If
                
            Case "E4" ' CP�ˤ��u
                ' ���`�ϥΤ��B�w�ˤ� EMTA �}������j��������� �B���p��P �A�� CP �������������n���@���]�t�H�W�^���`�A
                ' �s���@���w�ˤ��]�L�˩������^ ���\��~�i�HENABLE
                ' SNO <> "" AND InstDate is Null AND PRdate Is Null
                If Not Is_EMTA_Dev(rs04) Then
                    AddRetMsg "���]�ƫD EMTA , �L�k�i�� CP �ˤ��u !"
                    Exit Function
                End If
                If Not IsNull(!CMCloseDate) And Not IsNull(!CMOpenDate) Then
                    If CDT(!CMCloseDate) > CDT(!CMOpenDate) Then
                        AddRetMsg "�L [ �w�ˤ� ] �� [ ���` ] �� EMTA �]�� , �L�k�i�� CP �ˤ��u !"
                        Exit Function
                    End If
                End If
                If GetRsValue("SELECT COUNT(*) FROM " & GetOwner & "SO004" & _
                                " WHERE REFACISNO='" & !FaciSno & "'" & _
                                " AND COMPCODE=" & !CompCode & " AND CUSTID=" & !CustID & _
                                " AND SERVICETYPE='P' AND SNO IS NOT NULL " & _
                                " AND PRDATE IS NULL AND INSTDATE IS NULL") < 1 Then
                    AddRetMsg "�� EMTA �L���� CP ���� , �L�k�w�� CP ���u !"
                    ' MsgBox "�������L���`������ EMTA �]�� , �L�k�w�� CP ���u !", vbInformation, "�T��"
                    Exit Function
                End If
                
            Case "E5" ' CP ����u
                ' ���`�ϥΤ��B�w�ˤ� EMTA �}������j��������� �B ���p��P �A�� CP �������������n���@���]�t�H�W�^���`�A���@������]���˾�����B�L�������^ ���\��~�i�HENABLE
                ' PRSNO <> "" and PRdate =""
                If Not Is_EMTA_Dev(rs04) Then
                    AddRetMsg "���]�ƫD EMTA , �L�k�i�� CP ����u !"
                    Exit Function
                End If
                ' ���`�ϥΤ��B�w�ˤ� EMTA �}������j��������� �B���p��ӫȤ�P �A�Ȫ�CP �������������n���@���]�t�H�W�^���`�A
                ' �s���@���w�ˤ��]�L�˩������^ ���\��~�i�HENABLE
                ' EMTA �Ǹ����p��so004.servicetype='P' and prdate is null �����X p �A�Ȫ� so004.Gateway ���@���Y�i�Pcp �ۦP���ȭ�h
                ' If Not IsNull(!CMCloseDate) And Not IsNull(!CMOpenDate) Then
                '     If CDT(!CMCloseDate) > CDT(!CMOpenDate) Then
                '         ' MsgBox "�L [ �w�ˤ� ] �� [ ���` ] �� EMTA �]�� , �L�k�i�� CP ����u !", vbInformation, "�T��"
                '         MsgBox "�]�Ƥw���� , �L�k�i�� CP ����u !", vbInformation, "�T��"
                '         Exit Function
                '     End If
                ' End If
                ' PRSNO  ����, InstDate �� , PRdate �L
                ' AND SNO IS NOT NULL
                If GetRsValue("SELECT COUNT(*) FROM " & GetOwner & "SO004" & _
                                " WHERE REFACISNO='" & !FaciSno & "'" & _
                                " AND COMPCODE=" & !CompCode & " AND CUSTID=" & !CustID & _
                                " AND SERVICETYPE='P'" & _
                                " AND PRDATE IS NULL AND INSTDATE IS NOT NULL") <= 1 Then
                    AddRetMsg "�L�k����� !"
                    ' MsgBox "�L���`�� CP ���� , �L�k CP ����u !", vbInformation, "�T��"
                    ' MsgBox "�������L���`������ EMTA �]�� , �L�k�w�� CP ����u !", vbInformation, "�T��"
                    Exit Function
                End If
                ' ������渹
                If GetRsValue("SELECT COUNT(*) FROM " & GetOwner & "SO004" & _
                                " WHERE REFACISNO='" & !FaciSno & "'" & _
                                " AND COMPCODE=" & !CompCode & " AND CUSTID=" & !CustID & _
                                " AND SERVICETYPE='P' AND PRSNO IS NOT NULL" & _
                                " AND PRDATE IS NULL AND INSTDATE IS NOT NULL") > 1 Then
                    AddRetMsg "�� 2 �i (�H�W) ����� , �ФU����R�O !"
                    Exit Function
                End If
                
            Case "A1_BU", "TRIAL", "STOPTRIAL"
            
                    If strBaudRate = Empty Or Get_BaudRate_Desc(strBaudRate) & "" = !CMBaudRate & "" Then
                        AddRetMsg "�жǤJ [ �s�t�v ] !"
                        Exit Function
                    End If
                    
                    If CmdID = "TRIAL" Then
                        If strProbStopDate = Empty Then
                            AddRetMsg "�жǤJ [ �w�p�եκI��� ] ��A�i�� [ �]�w ] !"
                            Exit Function
                        End If
                        ' �w�p�եκI��� : ProbationStopDate
                        ' ����եΪ��A : ProbationStopFlag ( 0=�_, 1=�O )
                        ' �B(SO004.ProbationStopDate�L�ȩ�SO004. ProbationStopFlag=1)�~��U���R�O�C
                        If !ProbationStopDate & "" = Empty Or !ProbationStopFlag = 1 Then
                        Else
                            AddRetMsg "�L�k�U [�եγt�v] �R�O !"
                            Exit Function
                        End If
                    End If
                    
                    If CmdID = "STOPTRIAL" Then
                        If !ProbationStopDate & "" <> Empty And !ProbationStopFlag = 0 Then
                        Else
                            AddRetMsg "�L�k�U [�פ�ե�] �R�O !"
                            Exit Function
                        End If
                    End If
                    ' �w�ˤ��Υ��`���]��
                    If Not IsNull(!CMCloseDate) And Not IsNull(!CMOpenDate) Then
                        If CDT(!CMCloseDate) > CDT(!CMOpenDate) Then
                            If CmdID = "A1_BU" Then
                                AddRetMsg "�L [ �w�ˤ� ] �� [ ���` ] ���]�� , �L�k�i�� [�t�v�ɭ���] !"
                            ElseIf CmdID = "TRIAL" Then
                                AddRetMsg "�L [ �w�ˤ� ] �� [ ���` ] ���]�� , �L�k�U [�եγt�v] �R�O !"
                            Else
                                AddRetMsg "�L [ �w�ˤ� ] �� [ ���` ] ���]�� , �L�k�U [�פ�ե�] �R�O !"
                            End If
                            Exit Function
                        End If
                    End If
                    
        End Select
        
    End With
    
    IsDataOK = True
    
  Exit Function
ChkErr:
    ErrSub "clsCommand", "IsDataOK"
End Function

Private Sub GetFaciName(lngCustId As Long)
  On Error GoTo ChkErr
    Dim rsTmp As New ADODB.Recordset
    strQry = "SELECT FACICODE,FACINAME FROM " & GetOwner & "SO004" & _
                    " WHERE CUSTID=" & lngCustId & _
                    " AND COMPCODE=" & gCompCode & _
                    " AND FACISNO='" & strFCsno & "'"
    If GetRS(rsTmp, strQry, gcnGi) Then
        If rsTmp.State > 0 And Not rsTmp.EOF Then
            strFC = rsTmp!FaciCode & ""
            strFN = rsTmp!FaciName & ""
        End If
    End If
    On Error Resume Next
    CloseRecordset rsTmp
  Exit Sub
ChkErr:
    ErrSub "clsCommand", "GetFaciName"
End Sub

'�ѦҸ�2 ���sCM A7
'�ѦҸ�2 ��CM / ���s�ѦҸ�5 EMTA A9
'�ѦҸ�5 ���sEMTA A8
'�ѦҸ�5 EMTA / ���s�ѦҸ�2 CM A10
Private Function Get_RP_Cmd(ByRef rs04 As Object, FaciCode As String) As String
  On Error GoTo ChkErr
    
    Dim NewDevRefNo As Integer
    
    strQry = "SELECT REFNO FROM " & GetOwner & "CD022 WHERE CODENO=" & FaciCode
    NewDevRefNo = Val(GetRsValue(strQry, gcnGi))
    
    strQry = "SELECT REFNO FROM " & GetOwner & "CD022 WHERE CODENO=" & rs04("FaciCode")
    
    Select Case Val(GetRsValue(strQry, gcnGi))
        Case 2
            Get_RP_Cmd = IIf(NewDevRefNo = 2, "A7", "A9")
        Case 5
            Get_RP_Cmd = IIf(NewDevRefNo = 5, "A8", "A10")
    End Select
    
  Exit Function
ChkErr:
    ErrSub "clsCommand", "Get_RP_Cmd"
End Function

Private Function Is_EMTA_Dev(ByRef rs04 As Object) As Boolean
  On Error GoTo ChkErr
    If rs04.State > 0 Then
        If rs04("FaciCode") & "" = Empty Then
            AddRetMsg "�`�N [�Ȥ�]�Ƹ����].[�~�W�N�X] ����!"
        Else
            Is_EMTA_Dev = Val(GetRsValue("SELECT COUNT(*) FROM " & GetOwner & "CD022" & _
                                            " WHERE CODENO=" & rs04("FaciCode") & _
                                            " AND REFNO=5", gcnGi, "") & "") > 0
        End If
    End If
  Exit Function
ChkErr:
    ErrSub "clsCommand", "Is_EMTA_Dev"
End Function

' CD022 �~�W�s���N�X��
' CPImportMode �����פJ�Ҧ�
' 0-�x�T�B1-�ȤӡB2-�t��
' 0,1 �� CP , 2 �� VOIP

Private Function IsCMCP(ByRef rs04 As Object) As Boolean
  On Error GoTo ChkErr
        IsCMCP = GetRsValue("SELECT COUNT(*) FROM " & GetOwner & "SO004" & _
                            " WHERE REFACISNO='" & rs04("FaciSno") & "'" & _
                            " AND COMPCODE=" & rs04("CompCode") & _
                            " AND CUSTID=" & rs04("CustID") & _
                            " AND SERVICETYPE='P'" & _
                            " AND PRDATE IS NULL") > 1
  Exit Function
ChkErr:
    ErrSub "clsCommand", "IsCMCP"
End Function
