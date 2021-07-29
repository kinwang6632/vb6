Attribute VB_Name = "mduControlPanelWeb"
Option Explicit

Private Const blnDebug = False

Public gCompFilterStr As String
Public gCompCode As String
Public gServiceType As String
Public gDefaultComp As String
Public objActForm As Form
Public gLogInID As String
Public gstrUserPremission As String
Public lngSeqNo As Long
Public strCMDSeqNo As String

'Public strCN As String
'Public gVerType As Integer
'Public gTestMode As Boolean
'Public gcnGi2 As New ADODB.Connection
'Dim objCmd As New CS_Web.clsEMCcommand

' �I�s EMC CM CP Web Command ( Call Web Service )
Public Function Call_CMCP_WebCmd(strID_Code As String, _
                                                            Optional strCompCode As String = "", _
                                                            Optional strCustID As String = "", _
                                                            Optional strCM_MAC As String = "", _
                                                            Optional strEMTA_MAC As String = "", Optional strEMTA_Port As String = "", _
                                                            Optional strCP_No As String = "", _
                                                            Optional strClass_ID As String = "", _
                                                            Optional strPC_IP_NO As String = "", _
                                                            Optional strPC_IP As String = "", _
                                                            Optional strOper_Type As String = "", _
                                                            Optional strStart_Date_Time As String = "", Optional strEnd_Date_Time As String = "", _
                                                            Optional lngADDRNO As Long = 0, _
                                                            Optional ByVal strDialAccount As String = "", Optional ByVal strDialPassword As String = "", _
                                                            Optional ByVal strTel1 As String = "", Optional ByVal strTel2 As String = "", _
                                                            Optional ByVal strCustName As String, _
                                                            Optional ByVal FinTime As String, _
                                                            Optional ByRef strCMDSeqNo As String, _
                                                            Optional ByVal strCmdType As String, _
                                                            Optional ByVal strID As String, _
                                                            Optional ByRef strResult As String = "", Optional ByRef strFaultReason As String = "", _
                                                            Optional ByRef strErrMsg As String = "", _
                                                            Optional ByRef objResult As Object, _
                                                            Optional ByVal strSNO As String = "", Optional ByVal strFaciSeqNo As String = "", _
                                                            Optional ByVal strNodeNo As String = "", Optional ByVal strCircuitNo As String = "") _
                                                            As Boolean
  On Error GoTo ChkErr
  
    Dim objWebCmd As Object
    Set objWebCmd = CreateObject("CS_Web.clsEMCcommand")
    If strCompCode = 0 Then strCompCode = gCompCode
    With objWebCmd
            Call_CMCP_WebCmd = .Web_Command(gcnGi.ConnectionString, _
                                                                            strCompCode, _
                                                                            GetOwner, _
                                                                            strID_Code, _
                                                                            garyGi(0), _
                                                                            strCompCode, _
                                                                            strCustID, _
                                                                            strCM_MAC, _
                                                                            strEMTA_MAC, strEMTA_Port, _
                                                                            strCP_No, _
                                                                            strClass_ID, _
                                                                            strPC_IP_NO, strPC_IP, _
                                                                            strOper_Type, _
                                                                            strStart_Date_Time, strEnd_Date_Time, _
                                                                            lngADDRNO, _
                                                                            strDialAccount, strDialPassword, _
                                                                            strTel1, strTel2, _
                                                                            strCustName, FinTime, _
                                                                            strCMDSeqNo, strCmdType, strID, _
                                                                            strResult, strFaultReason, _
                                                                            objResult, frmSO1623A.pbr1, _
                                                                            strSNO, strFaciSeqNo, _
                                                                            strNodeNo, strCircuitNo)
            strErrMsg = .GetErrMsg
    End With
    Set objWebCmd = Nothing
  Exit Function
ChkErr:
    ErrSub "Sys_Lib", "Call_CMCP_WebCmd"
End Function

Public Function GetModeTypeValue(ByVal strValue As String) As String
    On Error Resume Next
    Dim strReturn As String
        Select Case strValue
            Case "A1"
                strReturn = "�}��"
            Case "A2"
                strReturn = "����"
            Case "A3"
                strReturn = "����"
            Case "A4"
                strReturn = "���](Reset)"
            
            Case "B1"
                strReturn = "��CM"
            Case "B2"
                strReturn = "�ܧ�CM�t�v"
            Case "B3"
                strReturn = "Ping CM"
            Case "B4"
                strReturn = "����ʺAIP"
            Case "B5"
                strReturn = "ACS�b���K�X�����\��"
            
            Case "C1"
                strReturn = "CM ���A�d��"
            Case "C2"
                strReturn = "�d�� HFC �A�����O"
            Case "C3"
                strReturn = "�M��EMM"
            Case "C4"
                strReturn = "CM �s�u�����d��"
            Case "C5"
                strReturn = "PCIP�s�u�O���d��"
        End Select
        GetModeTypeValue = strReturn
End Function

'(1) CM �} / ��/��/ ��/�ܧ�t�v
Public Function OpenCM(rsData As ADODB.Recordset, ByVal strPCMac As String, _
                                        ByVal lngInstAddrNo As Long, ByVal strClassID As String, _
                                        ByVal strDynaIP As String, _
                                        ByRef strResult As String, ByRef strFaultReason As String, _
                                        ByRef strErrMsg As String, ByRef rsResult As Object) As Boolean
  On Error GoTo ChkErr
    Dim strGroupID As String, strCTIP As String
    If Not Call_CMCP_WebCmd("02", rsData("CompCode") & "", rsData("CustId") & "", _
                                            rsData("FaciSNo") & "", IIf(frmSO1623A.IsCMCP, Get_EMTA_Mac(rsData("FaciSNo") & ""), ""), , , _
                                            strClassID, strDynaIP, , 1, , , lngInstAddrNo, _
                                            , , , , , , strCMDSeqNo, , , strResult, strFaultReason, strErrMsg, rsResult) Then
    Else
'        On Error Resume Next
'        strGroupID = rsResult("GroupID") & ""
        
        strCTIP = rsResult("CTIP") & ""
        If strCTIP <> Empty Then
            IPwarning rsData("IPAddress") & "", strCTIP
            rsData("IPAddress") = strCTIP
            rsData.Update
        End If
        OpenCM = True
    End If

    ShowMsg "(1) CM �} / ��/��/ ��/�ܧ�t�v", strResult, strFaultReason, strErrMsg

  Exit Function
ChkErr:
    ErrSub FormName, "OpenCM"
End Function

'(1) CM �} / ��/��/ ��/�ܧ�t�v
Public Function CloseCM(rsData As ADODB.Recordset, ByVal strPCMac As String, _
                                        ByVal lngInstAddrNo As Long, ByVal strClassID As String, _
                                        ByVal strDynaIP As String, _
                                        ByRef strResult As String, ByRef strFaultReason As String, _
                                        ByRef strErrMsg As String, ByRef rsResult As Object) As Boolean
  On Error GoTo ChkErr
    Dim strGroupID As String, strCTIP As String
    If Not Call_CMCP_WebCmd("02", rsData("CompCode") & "", rsData("CustId") & "", _
                                                rsData("FaciSNo") & "", Get_EMTA_Mac(rsData("FaciSNo") & ""), , , _
                                                strClassID, strDynaIP, , 2, , , lngInstAddrNo, _
                                                , , , , , , strCMDSeqNo, , , strResult, strFaultReason, strErrMsg, rsResult) Then
    Else
'        strGroupID = rsResult("GroupID") & ""
        strCTIP = rsResult("CTIP") & ""
        If strCTIP <> Empty Then
            IPwarning rsData("IPAddress") & "", strCTIP
            rsData("IPAddress") = strCTIP
            rsData.Update
        End If
        CloseCM = True
    End If
    
    ShowMsg "(1) CM �} / ��/��/ ��/�ܧ�t�v", strResult, strFaultReason, strErrMsg

  Exit Function
ChkErr:
    ErrSub FormName, "CloseCM"
End Function

'(1) CM �} / ��/��/ ��/�ܧ�t�v
Public Function StopCM(rsData As ADODB.Recordset, ByVal strPCMac As String, _
                                        ByVal lngInstAddrNo As Long, ByVal strClassID As String, _
                                        ByVal strDynaIP As String, _
                                        ByRef strResult As String, ByRef strFaultReason As String, _
                                        ByRef strErrMsg As String, ByRef rsResult As Object) As Boolean
  On Error GoTo ChkErr
    Dim strGroupID As String, strCTIP As String
    If Not Call_CMCP_WebCmd("02", rsData("CompCode") & "", rsData("CustId") & "", _
                                                rsData("FaciSNo") & "", Get_EMTA_Mac(rsData("FaciSNo") & ""), , , _
                                                strClassID, strDynaIP, , 1, , , lngInstAddrNo, _
                                                , , , , , , , , , strResult, strFaultReason, strErrMsg, rsResult) Then
    Else
'        strGroupID = rsResult("GroupID") & ""
        strCTIP = rsResult("CTIP") & ""
        If strCTIP <> Empty Then
            IPwarning rsData("IPAddress") & "", strCTIP
            rsData("IPAddress") = strCTIP
            rsData.Update
        End If
        StopCM = True
    End If
    
    ShowMsg "(1) CM �} / ��/��/ ��/�ܧ�t�v", strResult, strFaultReason, strErrMsg

  Exit Function
ChkErr:
    ErrSub FormName, "StopCM"
End Function

Private Sub IPwarning(ByVal strOriIP As String, ByVal strNewIP As String)
  On Error Resume Next
    If Len(strOriIP) > 0 Then
        If strOriIP <> strNewIP Then MsgBox "�Ҩ��o��IP [ " & strNewIP & " ] �P����oIP [ " & strOriIP & " ] ���P �A �нT�{ EMTA �O�_�i���`�B�@ !", vbInformation, "�T��"
    End If
End Sub

'(1) CM �} / ��/��/ ��/�ܧ�t�v
Public Function ChangeCM(rsData As ADODB.Recordset, ByVal strPCMac As String, _
                                            ByVal lngInstAddrNo As Long, ByVal strClassID As String, _
                                            ByVal strDynaIP As String, _
                                            ByRef strResult As String, ByRef strFaultReason As String, _
                                            ByRef strErrMsg As String, ByRef rsResult As Object) As Boolean
  On Error GoTo ChkErr
    Dim strGroupID As String, strCTIP As String
    If Not Call_CMCP_WebCmd("02", rsData("CompCode") & "", rsData("CustId") & "", _
                                                rsData("FaciSNo") & "", Get_EMTA_Mac(rsData("FaciSNo") & ""), , , _
                                                strClassID, strDynaIP, , 1, , , lngInstAddrNo, _
                                                , , , , , , , , , strResult, strFaultReason, strErrMsg, rsResult) Then
    Else
'        strGroupID = rsResult("GroupID") & ""
        strCTIP = rsResult("CTIP") & ""
        If strCTIP <> Empty Then
            IPwarning rsData("IPAddress") & "", strCTIP
            rsData("IPAddress") = strCTIP
            rsData.Update
        End If
        ChangeCM = True
    End If

    ShowMsg "(1) CM �} / ��/��/ ��/�ܧ�t�v", strResult, strFaultReason, strErrMsg

  Exit Function
ChkErr:
    ErrSub FormName, "ChangeCM"
End Function

'(1) CM �} / ��/��/ ��/�ܧ�t�v
Public Function ChangeCMRate(rsData As ADODB.Recordset, ByVal strPCMac As String, _
                                                    ByVal lngInstAddrNo As Long, ByVal strClassID As String, _
                                                    ByVal strDynaIP As String, _
                                                    ByRef strResult As String, ByRef strFaultReason As String, _
                                                    ByRef strErrMsg As String, ByRef rsResult As Object) As Boolean
  On Error GoTo ChkErr
    Dim strGroupID As String, strCTIP As String
    If Not Call_CMCP_WebCmd("02", rsData("CompCode") & "", rsData("CustId") & "", _
                                                rsData("FaciSNo") & "", Get_EMTA_Mac(rsData("FaciSNo") & ""), , , _
                                                strClassID, strDynaIP, , 1, , , lngInstAddrNo, _
                                                , , , , , , , , , strResult, strFaultReason, strErrMsg, rsResult) Then
    Else
'        strGroupID = rsResult("GroupID") & ""
        strCTIP = rsResult("CTIP") & ""
        If strCTIP <> Empty Then
            IPwarning rsData("IPAddress") & "", strCTIP
            rsData("IPAddress") = strCTIP
            rsData.Update
        End If
        ChangeCMRate = True
    End If
    
    ShowMsg "(1) CM �} / ��/��/ ��/�ܧ�t�v", strResult, strFaultReason, strErrMsg

  Exit Function
ChkErr:
    ErrSub FormName, "ChangeCMRate"
End Function

'(2) CP IP �ӽФΰh��
Public Function CPInst(rsData As ADODB.Recordset, ByVal strPCMac As String, _
                                    ByVal lngInstAddrNo As Long, ByRef strResult As String, _
                                    ByRef strFaultReason As String, ByRef strErrMsg As String, _
                                    ByRef rsResult As ADODB.Recordset) As Boolean
  On Error GoTo ChkErr
    CPInst = CPIP("1", rsData, strPCMac, lngInstAddrNo, strResult, strFaultReason, strErrMsg, rsResult)
  Exit Function
ChkErr:
    ErrSub FormName, "CPInst"
End Function

Private Function CPIP(strOperType As String, _
                                    rsData As ADODB.Recordset, ByVal strPCMac As String, _
                                    ByVal lngInstAddrNo As Long, ByRef strResult As String, _
                                    ByRef strFaultReason As String, ByRef strErrMsg As String, _
                                    ByRef rsResult As ADODB.Recordset) As Boolean
    
    Dim rsTmp As New ADODB.Recordset
    Dim strCTIP As String
    
    Dim strQry As String
    
    strQry = "SELECT PORTNO,FACISNO FROM " & GetOwner & "SO004" & _
                    " WHERE REFACISNO='" & rsData("FaciSno") & "'" & _
                    " AND COMPCODE=" & rsData("CompCode") & _
                    " AND CUSTID=" & rsData("CustID") & _
                    " AND SERVICETYPE='P'" & _
                    " AND INSTDATE IS NOT NULL" & _
                    " AND PRDATE IS NULL"
    ' Liga :
    '   2�DCP�Ӹ�[27]�G��3�ӰѼơGCMMAC, EMTAMAC, EMTAPORT�CCPNO�ǭȦ��~�A����P�A�Ȫ�FACISNO�~��A�{�O��REFACISNO
    
    CPIP = False
    
    If GetRS(rsTmp, strQry, gcnGi, adUseClient, adOpenForwardOnly, adLockReadOnly) Then
        With rsTmp
            If .RecordCount > 0 Then
                While Not .EOF
                    If Not Call_CMCP_WebCmd("27", rsData("CompCode") & "", rsData("CustId") & "", _
                                                                rsData("FaciSNo") & "", Get_EMTA_Mac(rsData("FaciSNo") & ""), _
                                                                rsTmp("PortNo") & "", rsTmp("FaciSNO") & "", , , , strOperType, , , lngInstAddrNo, _
                                                                , , , , , , , , , strResult, strFaultReason, strErrMsg, rsResult) Then
                    Else
                        strCTIP = rsResult("CTIP") & ""
                        If strCTIP <> Empty Then
                            IPwarning rsData("IPAddress") & "", strCTIP
                            rsData("IPAddress") = strCTIP
                            rsData.Update
                        End If
                        CPIP = True
                    End If
                    .MoveNext
                Wend
            Else
                ' �䤣�� P �A�ȸ��
                strFaultReason = "���]�ƵL�w�� CP !"
'                MsgBox "���]�ƵL�w�� CP !", vbInformation, "�T��"
            End If
        End With
    End If
    
    ShowMsg "(2) CP IP �ӽФΰh��", strResult, strFaultReason, strErrMsg
    
  Exit Function
ChkErr:
    ErrSub FormName, "CPIP"
End Function

'(2) CP IP �ӽФΰh��
Public Function CPPR(rsData As ADODB.Recordset, ByVal strPCMac As String, _
                                    ByVal lngInstAddrNo As Long, ByRef strResult As String, _
                                    ByRef strFaultReason As String, ByRef strErrMsg As String, _
                                    ByRef rsResult As ADODB.Recordset) As Boolean
  On Error GoTo ChkErr
    ' Liga :    3�DCP�h��[27]�GOPERTYPE���~�A���ǭ�2�~��C
    CPPR = CPIP("2", rsData, strPCMac, lngInstAddrNo, strResult, strFaultReason, strErrMsg, rsResult)
  Exit Function
ChkErr:
    ErrSub FormName, "CPPR"
End Function

'(3) Reset CM
Public Function ResetCM(rsData As ADODB.Recordset, ByRef strResult As String, _
                                        ByRef strFaultReason As String, ByRef strErrMsg As String, _
                                        ByRef rsResult As Object) As Boolean
  On Error GoTo ChkErr
    If Not Call_CMCP_WebCmd("12", rsData("CompCode") & "", rsData("CustId") & "", _
                                                rsData("FaciSNO") & "", , , , , , , , , , , _
                                                , , , , , , , , , strResult, strFaultReason, strErrMsg, rsResult) Then
    Else
        ResetCM = True
    End If
    
    ShowMsg "(3) Reset CM", strResult, strFaultReason, strErrMsg
    
  Exit Function
ChkErr:
    ErrSub FormName, "ResetCM"
End Function

'(4) Ping CM
Public Function PingCM(rsData As ADODB.Recordset, ByRef strResult As String, _
                                        ByRef strFaultReason As String, ByRef strErrMsg As String, _
                                        ByRef rsResult As Object) As Boolean
  On Error GoTo ChkErr
    If Not Call_CMCP_WebCmd("20", rsData("CompCode") & "", rsData("CustId") & "", _
                                                rsData("FaciSNO") & "", , , , , , , , , , , _
                                                , , , , , , , , , strResult, strFaultReason, strErrMsg, rsResult) Then
    Else
        PingCM = True
    End If
    
    ShowMsg "(4) Ping CM", strResult, strFaultReason, strErrMsg
    
  Exit Function
ChkErr:
    ErrSub FormName, "PingCM"
End Function

''(5) ���� �ʺAIP
'Public Function ReleaseDyIP(rsData As ADODB.Recordset, ByRef strResult As String, _
'                                                ByRef strFaultReason As String, ByRef strErrMsg As String, _
'                                                ByRef rsResult As Object) As Boolean
'  On Error GoTo ChkErr
'    If Not Call_CMCP_WebCmd("25", rsData("CompCode"), rsData("CustId"), _
'                                                rsData("FaciSNO"), , , , , , , , , , , _
'                                                strResult, strFaultReason, strErrMsg, rsResult) Then
'        Exit Function
'    End If
'
'    ReleaseDyIP = True
'  Exit Function
'ChkErr:
'    ErrSub FormName, "ReleaseDyIP"
'End Function

''(7) ACS�b���K�X�����\��
'Public Function ACSAccPWS(rsData As ADODB.Recordset, ByRef strResult As String, _
'                                            ByRef strFaultReason As String, ByRef strErrMsg As String, _
'                                            ByRef rsResult As Object) As Boolean
'  On Error GoTo ChkErr
'    If Not Call_CMCP_WebCmd("25", rsData("CompCode"), rsData("CustId"), _
'                                                rsData("FaciSNO"), , , , , , , , , , , strResult, strFaultReason, strErrMsg, rsResult) Then
'        Exit Function
'    End If
'
'    ACSAccPWS = True
'  Exit Function
'ChkErr:
'    ErrSub FormName, "ACSAccPWS"
'End Function

'(7) Clear CPE IP
Public Function ClearCPEIP(rsData As ADODB.Recordset, ByRef strResult As String, _
                                            ByRef strFaultReason As String, ByRef strErrMsg As String, _
                                            ByRef rsResult As Object) As Boolean
  On Error GoTo ChkErr
    If Not Call_CMCP_WebCmd("25", rsData("CompCode") & "", rsData("CustId") & "", _
                                                rsData("FaciSNO") & "", , , , , , , , , , , _
                                                , , , , , , , , , strResult, strFaultReason, strErrMsg, rsResult) Then
    Else
        ClearCPEIP = True
    End If
    
    ShowMsg "(5) ����ʺAIP", strResult, strFaultReason, strErrMsg

  Exit Function
ChkErr:
    ErrSub FormName, "ClearCPEIP"
End Function

'(6) CM ���A�d��
Public Function QueryCMStatus(rsData As ADODB.Recordset, ByVal lngInstAddrNo As Long, _
                                                    ByRef strResult As String, ByRef strFaultReason As String, ByRef strErrMsg As String, _
                                                    ByRef rsResult As ADODB.Recordset) As Boolean
  On Error GoTo ChkErr
    If Not Call_CMCP_WebCmd("13", rsData("CompCode") & "", rsData("CustId") & "", _
                                                rsData("FaciSNO") & "", , , , , , , , , , lngInstAddrNo, _
                                                , , , , , , , , , strResult, strFaultReason, strErrMsg, rsResult, _
                                                rsData("SNO") & "", rsData("SeqNo") & "") Then
    Else
        Dim strPCIP As String
        Dim varAry As Variant
        Dim varElement As Variant
        strPCIP = rsResult("PCIP") & ""
        If Len(strPCIP) > 0 Then
            gcnGi.Execute "DELETE FROM " & GetOwner & "SO004C" & _
                                    " WHERE CUSTID=" & rsData("CustID") & _
                                    " AND FACISNO='" & rsData("FaciSno") & "'" & _
                                    " AND SEQNO='" & rsData("SeqNo") & "'"
                                    
            If InStr(1, strPCIP, ",") Then
                varAry = Split(strPCIP, ",")
                For Each varElement In varAry
                    If varElement & "" <> Empty Then
                        gcnGi.Execute "INSERT INTO " & GetOwner & "SO004C" & _
                                                " (CUSTID,FACISNO,SEQNO,CPEMAC,IPAddress)" & _
                                                " VALUES (" & rsData("CustID") & _
                                                ",'" & rsData("FaciSno") & "'" & _
                                                ",'" & rsData("SeqNo") & "'" & _
                                                ",NULL,'" & varElement & "')"
                    End If
                Next
            Else
                gcnGi.Execute "INSERT INTO " & GetOwner & "SO004C" & _
                                        " (CUSTID,FACISNO,SEQNO,CPEMAC,IPAddress)" & _
                                        " VALUES (" & rsData("CustID") & _
                                        ",'" & rsData("FaciSno") & "'" & _
                                        ",'" & rsData("SeqNo") & "'" & _
                                        ",NULL,'" & rsResult("PCIP") & "')"
            End If
        End If
        ' �o�өR�O�� Return �@����
        ' QueryResult
        ' FaultReason
        ' GroupID
        ' CMStatus
        ' CMIP
        ' CMUpFreq
        ' CMDownFreq
        ' CMRECPW
        ' CMTransPW
        ' CMSNR
        ' CMOnlineTimes
        ' CMInOctets
        ' CMOutOctets
        ' CMInErrors
        ' CMOutErrors
        ' PCIP
        ' CMTSRECPW
        ' CMTSRXSNR
        ' CMTSUPMOD
        QueryCMStatus = True
    End If
    
    ShowMsg "(6) CM ���A�d��", strResult, strFaultReason, strErrMsg
    
  Exit Function
ChkErr:
    ErrSub FormName, "QueryCMStatus"
End Function

'(8) �d�� HFC �A�����O
Public Function QueryHFCServ(rsData As ADODB.Recordset, ByVal lngInstAddrNo As Long, _
                                                ByRef strResult As String, ByRef strFaultReason As String, _
                                                ByRef strErrMsg As String, ByRef rsResult As ADODB.Recordset) As Boolean
  On Error GoTo ChkErr
    
    Dim strDualCable As String
    Dim strNodeNo As String
    Dim lngRcdAft As Long
    Dim strCNorNN As String ' CircuitNo or NodeNo
    
    If lngInstAddrNo = 0 Then
        MsgBox "�a�}�s�� [AddrNo] �����T!", vbInformation, "�T��"
        Exit Function
    End If
    If Not Call_CMCP_WebCmd("21", rsData("CompCode") & "", , , , , , , , , , , , _
                                                lngInstAddrNo, , , , , , , , , , strResult, strFaultReason, strErrMsg, rsResult) Then
'        MsgBox strErrMsg, vbInformation, "�T��"
    Else
        ' �o�өR�O�� Return �@����
        ' OperResult
        ' FaultReason
        ' HFCType = SO014.DualCable
        ' NODEID = SO014.NodeNo
        
        If strResult = "1" Then
        
            ShowMsg "(8) �d�� HFC �A�����O", strResult, strFaultReason, strErrMsg
        
            With rsResult
                    If .EOF Then .MoveFirst
                    strDualCable = .Fields("HFCTYPE") & ""
                    strNodeNo = .Fields("NODEID") & ""
            End With
            
            ' SO041.EMTAIPTYPE ( 0=�����s�� , 1=Node )
            If GetSystemParaItem("EMTAIPTYPE", gCompCode, gServiceType, "SO041", , True) = 0 Then
                strCNorNN = "CIRCUITNO"
            Else
                strCNorNN = "NODENO"
            End If
            
'                If strDualCable <> Empty Then strDualCable = Choose(Left(strDualCable, 1), "0", "1", "2", "3")
            ' DualCable (HFCTYPE) 2 , 3 ������ Jim ��������, �ҥH���B�z
            If strDualCable <> Empty Then strDualCable = Choose(Left(strDualCable, 1), "0", "1") & ""
            
            lngRcdAft = 0
            
            If strDualCable <> "" And strNodeNo <> "" Then
                gcnGi.Execute "UPDATE " & GetOwner & "SO014" & _
                                        " SET DUALCABLE=" & strDualCable & "," & _
                                        " " & strCNorNN & "='" & strNodeNo & "'" & _
                                        " WHERE ADDRNO=" & lngInstAddrNo & _
                                        " AND COMPCODE=" & gCompCode, lngRcdAft
                                        
            ElseIf strDualCable <> "" Then
                gcnGi.Execute "UPDATE " & GetOwner & "SO014" & _
                                        " SET DUALCABLE=" & strDualCable & _
                                        " WHERE ADDRNO=" & lngInstAddrNo & _
                                        " AND COMPCODE=" & gCompCode, lngRcdAft
                                        
            ElseIf strNodeNo <> "" Then
                gcnGi.Execute "UPDATE " & GetOwner & "SO014" & _
                                        " SET " & strCNorNN & "='" & strNodeNo & "'" & _
                                        " WHERE ADDRNO=" & lngInstAddrNo & _
                                        " AND COMPCODE=" & gCompCode, lngRcdAft
            Else
                MsgBox "�^�� [HFCType] , [NODEID] ��Ƭ��� ! �a�}��ƥ���s!", vbInformation, "�T��"
            End If
                                    
            If lngRcdAft > 0 Then
                MsgBox "�a�}��� [HFCType] , [NODEID] ��s����!", vbInformation, "�T��"
            Else
                MsgBox "�d�L�a�}��� , �L�k��s!", vbInformation, "�T��"
            End If
            
'                Debug.Print rsResult.GetString

            ' QQ ? Schema �̵��O NodeNo �Ѻ����s�� (CircuitNo) �e�T�X�s�J
            
            '��Ȥ�D��[SO1100B]���e���W�[�y���u��ơz�d�ߪ�����F
            '���ϥΪ��I�惡����ɽШ̾ڡ@�������@CC&B�@�~�y�{�PSPM�@�����y�{��_Jim_20050805.XLS
            '  ���o�ӭ��ҡ@��EMC�@WEB�@Command�@matrix�@����
            '  EMC���ު��������{���ǰe��8�ӫ��O
            '  ���d�ߡ@HFC�@�A�����O�@HFCTypeQuery�@21��
            '���åB�̾ڦ��Ȥ᪺�a�}��ƨ��o�����s����NODENO��,�@�Цs�J���Ȥ�˾��a�}�����T���
            '  (SO001.INSTADDRNO=SO014.ADDRNO,�@�N���^�Ӫ��Ȧ^���SO014.Circuitno�@�Ρ@so014.Nodeno)

        Else
'            MsgBox "���浲�G : " & strResult & vbCrLf & "���ѭ�] : " & strFaultReason, vbInformation, "�T��"
            ShowMsg "(8) �d�� HFC �A�����O", strResult, strFaultReason, strErrMsg
        End If
        
        QueryHFCServ = True
    End If
    
  Exit Function
ChkErr:
    ErrSub FormName, "QueryHFCServ"
End Function

'(9) CM �s�u�����d��
Public Function QueryCMHistory(rsData As ADODB.Recordset, ByVal strStartDate As String, _
                                                    ByVal strEndDate As String, ByRef strResult As String, _
                                                    ByRef strFaultReason As String, ByRef strErrMsg As String, _
                                                    ByRef rsResult As ADODB.Recordset) As Boolean
  On Error GoTo ChkErr
    If Not Call_CMCP_WebCmd("22", rsData("CompCode") & "", rsData("CustId") & "", rsData("FaciSNo") & "", _
                                                , , , , , , , strStartDate, strEndDate, , _
                                                , , , , , , , , , strResult, strFaultReason, strErrMsg, rsResult) Then
    Else
        ' �o�өR�O�� Return �@����
        ' QueryResult
        ' FaultReason
        ' CMIP
        ' PCIP = SO004C����IP�ƻPMAC�����p��T�A�����G�@���ʩηs�W�ΧR��
        ' OffLineTime
        ' OnLineTime
        QueryCMHistory = True
    End If
    
    ShowMsg "(9) CM �s�u�����d��", strResult, strFaultReason, strErrMsg
    
  Exit Function
ChkErr:
    ErrSub FormName, "QueryCMHistory"
End Function

'(10) CM �s�u�~��d��
Public Function QueryCMQuality(rsData As ADODB.Recordset, ByVal strStartDate As String, _
                                                    ByVal strEndDate As String, ByRef strResult As String, _
                                                    ByRef strFaultReason As String, ByRef strErrMsg As String, _
                                                    ByRef rsResult As ADODB.Recordset) As Boolean
  On Error GoTo ChkErr
    If Not Call_CMCP_WebCmd("23", rsData("CompCode") & "", rsData("CustId") & "", rsData("FaciSNo") & "", _
                                                , , , , , , , strStartDate, strEndDate, , _
                                                , , , , , , , , , strResult, strFaultReason, strErrMsg, rsResult) Then
    Else
        ' �o�өR�O�� Return �@����
        ' QueryResult
        ' FaultReason
        ' CMRECPW
        ' CMTransPW
        ' CMSNR
        ' CMInOctets
        ' CMOutOctets
        ' CMInErrors
        ' CMOutErrors
        ' CMTSRECPW
        ' CMTSRXSNR
        ' DetctTime
        QueryCMQuality = True
    End If
    
    ShowMsg "(10) CM �s�u�~��d��", strResult, strFaultReason, strErrMsg
    
  Exit Function
ChkErr:
    ErrSub FormName, "QueryCMQuality"
End Function

'(11)PCIP�s�u�O���d��
Public Function QueryPCIPHistory(rsData As ADODB.Recordset, ByVal strPCIP As String, _
                                                        ByVal strStartDate As String, ByVal strEndDate As String, _
                                                        ByRef strResult As String, ByRef strFaultReason As String, _
                                                        ByRef strErrMsg As String, ByRef rsResult As ADODB.Recordset) As Boolean
  On Error GoTo ChkErr
    If Not Call_CMCP_WebCmd("26", rsData("CompCode") & "", rsData("CustId") & "", _
                                                , , , , , , strPCIP, , strStartDate, strEndDate, , _
                                                , , , , , , , , , strResult, strFaultReason, strErrMsg, rsResult) Then
    Else
        ' �o�өR�O�� Return �@����
        ' QueryResult
        ' FaultReason
        ' PConlineTime
        ' PCofflineTime
        ' CMMAC
        QueryPCIPHistory = True
    End If
    
    ShowMsg "(11) PCIP�s�u�O���d��", strResult, strFaultReason, strErrMsg
    
  Exit Function
ChkErr:
    ErrSub FormName, "QueryPCIPHistory"
End Function

' �b���ҥΰ���
Public Function Cmd_Acc_24(ByRef rsData As Object, strCMD As String, Optional strCMDSeqNo As String = "") As Boolean
  On Error GoTo ChkErr
    
    Dim strID As String
    
    'IDKindCode1 ID IDKindCode ID2
     
     With rsData
        
        If .Fields("IDKindCode1") & "" <> "" Then
            If GetRsValue("SELECT NVL(REFNO,0) FROM " & GetOwner & "CD070" & _
                                   " WHERE CODENO = " & .Fields("IDKindCode1")) = 1 Then
                strID = .Fields("ID") & ""
            End If
        End If
        
        If .Fields("IDKindCode") & "" <> "" Then
            If GetRsValue("SELECT NVL(REFNO,0) FROM " & GetOwner & "CD070" & _
                                    " WHERE CODENO = " & .Fields("IDKindCode")) = 1 Then
                strID = .Fields("ID2") & ""
            End If
        End If
     
     End With
    
    Cmd_Acc_24 = Call_CMCP_WebCmd("24", rsData("CompCode"), rsData("CustId"), _
                                                rsData("FaciSNo"), Get_EMTA_Mac(rsData("FaciSNo") & ""), _
                                                , , , , , strCMD, , , , rsData("DialAccount") & "", , , , , , strCMDSeqNo, strCMD, strID)
    
    If Cmd_Acc_24 Then MsgBox "�w�����}�q�O�� !", vbInformation, "�T��"



'    disable: �b�����v
'    Enable: �b���ҥ�
                                                    
'    �t�Υx�O Cusso
'    �Ȥ�s�� Cusid
'    CM MAC  CM MAC
'    �\����O(�}��)  OperType
'    �����b�� DialAccount
'    ���O log table �y����   CmdSeqno
  
  Exit Function
ChkErr:
    ErrSub FormName, "Cmd_Acc_24"
End Function


Private Sub ShowMsg(strCMD As String, strResult As String, strFaultReason As String, strErrMsg As String)
    On Error GoTo ChkErr
        If strResult = "1" Then
            MsgBox "[ " & strCMD & " ] ����T���w�e�X���� !! ", vbInformation, gimsgPrompt
        Else
            If strFaultReason = Empty Then
                MsgBox "[ " & strCMD & " ] ���� !! " & strCrlf & _
                                "���~��] : " & strErrMsg, vbInformation, gimsgPrompt
            Else
                MsgBox "[ " & strCMD & " ] ���� !! " & strCrlf & _
                                "���~��] : " & strFaultReason & strCrlf & _
                                strErrMsg, vbInformation, gimsgPrompt
            End If
        End If
    Exit Sub
ChkErr:
    ErrSub FormName, "subShowMessage"
End Sub



