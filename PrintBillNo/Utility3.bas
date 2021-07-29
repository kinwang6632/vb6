Attribute VB_Name = "Utility3"
Option Explicit
Private Const FormName = "Utility3"

'�O���ثe�ϥΪ��v��

Public Enum giPermissionModeEnu
    giPermissionModeAllow = True      ' True=���v��
    giPermissionModeNo = False      ' False=�S���v��
End Enum

Public Enum giFieldDataType
    giDataNumeric = 1
    giDataBoolean = 2
    giDataString = 3
    giDataMemo = 4
    giDataDate = 5
    giDataNone = 6
End Enum

Public Const ConnectErr = "�s�u���ѡI�Э��s�ާ@�I"
Public Const gmsgNotSave = "�s�W�έק良�s��,�O�_�u�n���}?"

'Public opendb As New giOpenDBObj.OpenDBObj

Public Function SF_GETDISCOUNT(ByVal P_SALECODE As Long, ByVal P_CITEMCODE As Long, _
ByVal p_StartDate As String, ByVal p_OldAmt As Long, ByVal p_CompCode As Long, _
ByVal p_SERVICETYPE As String, ByRef P_DISCOUNTAMT As Long, _
ByRef P_DISCOUNTDATE1 As String, ByRef P_DISCOUNTDATE2 As String, ByRef P_DISCOUNTCODE As Long, _
ByRef P_DISCOUNTNAME As String, ByRef P_RETMSG As String) As Integer
On Error GoTo ChkErr
        Dim ComSF_GETDISCOUNT As New ADODB.Command

        With ComSF_GETDISCOUNT
                .Parameters.Append .CreateParameter("RETURN_VALUE", adVarNumeric, adParamReturnValue)
                .Parameters.Append .CreateParameter("P_SALECODE", adVarNumeric, adParamInput, , P_SALECODE)
                .Parameters.Append .CreateParameter("P_CITEMCODE", adVarNumeric, adParamInput, , P_CITEMCODE)
                .Parameters.Append .CreateParameter("P_STARTDATE", adVarChar, adParamInput, 2000, p_StartDate)
                .Parameters.Append .CreateParameter("P_OLDAMT", adVarNumeric, adParamInput, , p_OldAmt)
                .Parameters.Append .CreateParameter("P_COMPCODE", adVarNumeric, adParamInput, , p_CompCode)
                .Parameters.Append .CreateParameter("P_SERVICETYPE", adVarChar, adParamInput, 2000, p_SERVICETYPE)
                .Parameters.Append .CreateParameter("P_DISCOUNTAMT", adVarNumeric, adParamOutput, , P_DISCOUNTAMT)
                .Parameters.Append .CreateParameter("P_DISCOUNTDATE1", adVarChar, adParamOutput, 2000, P_DISCOUNTDATE1)
                .Parameters.Append .CreateParameter("P_DISCOUNTDATE2", adVarChar, adParamOutput, 2000, P_DISCOUNTDATE2)
                .Parameters.Append .CreateParameter("P_DISCOUNTCODE", adVarNumeric, adParamOutput, , P_DISCOUNTCODE)
                .Parameters.Append .CreateParameter("P_DISCOUNTNAME", adVarChar, adParamOutput, 2000, P_DISCOUNTNAME)
                .Parameters.Append .CreateParameter("P_RETMSG", adVarChar, adParamOutput, 2000, P_RETMSG)
                Set .ActiveConnection = gcnGi
                .CommandText = "SF_GETDISCOUNT"
                .CommandType = adCmdStoredProc
                .Execute
                P_DISCOUNTAMT = Val(.Parameters("P_DISCOUNTAMT").Value & "")
                P_DISCOUNTDATE1 = .Parameters("P_DISCOUNTDATE1").Value & ""
                P_DISCOUNTDATE2 = .Parameters("P_DISCOUNTDATE2").Value & ""
                P_DISCOUNTCODE = Val(.Parameters("P_DISCOUNTCODE").Value & "")
                P_DISCOUNTNAME = .Parameters("P_DISCOUNTNAME").Value & ""
                P_RETMSG = .Parameters("P_RETMSG").Value & ""
                SF_GETDISCOUNT = Val(.Parameters("Return_value").Value & "")
        End With
        Exit Function
ChkErr:
        ErrSub "clsStoreFunction", "SF_GETDISCOUNT"
End Function

Public Function SF_GETDISCOUNT1(ByVal P_CUSTID As Long, ByVal P_CITEMCODE As Long, _
ByVal p_StopDate As String, ByRef P_DISCOUNTAMT As Long, _
ByRef P_RETMSG As String) As Integer
On Error GoTo ChkErr
        Dim ComSF_GETDISCOUNT1 As New ADODB.Command

        With ComSF_GETDISCOUNT1
                .Parameters.Append .CreateParameter("RETURN_VALUE", adVarNumeric, adParamReturnValue)
                .Parameters.Append .CreateParameter("P_CUSTID", adVarNumeric, adParamInput, , P_CUSTID)
                .Parameters.Append .CreateParameter("P_CITEMCODE", adVarNumeric, adParamInput, , P_CITEMCODE)
                .Parameters.Append .CreateParameter("P_STOPDATE", adVarChar, adParamInput, 2000, p_StopDate)
                .Parameters.Append .CreateParameter("P_DISCOUNTAMT", adVarNumeric, adParamOutput, , P_DISCOUNTAMT)
                .Parameters.Append .CreateParameter("P_RETMSG", adVarChar, adParamOutput, 2000, P_RETMSG)
                Set .ActiveConnection = gcnGi
                .CommandText = "SF_GETDISCOUNT1"
                .CommandType = adCmdStoredProc
                .Execute
                P_DISCOUNTAMT = Val(.Parameters("P_DISCOUNTAMT").Value & "")
                P_RETMSG = .Parameters("P_RETMSG").Value & ""
                SF_GETDISCOUNT1 = Val(.Parameters("Return_value").Value & "")
        End With
        Exit Function
ChkErr:
        ErrSub "clsStoreFunction", "SF_GETDISCOUNT1"
End Function

Public Function CallSF_GETAMOUNT3(ByVal P_CUSTID As Long, _
        ByVal P_CITEMCODE As Long, ByVal P_PERIOD As Long, _
        ByVal p_SERVICETYPE As String, ByVal P_SHOULDDATE As String, _
        ByVal p_CompCode As Long, ByRef p_OldAmt As Long, _
        ByRef P_RETMSG As String) As Long
  On Error GoTo ChkErr
    Dim ComSF_GETAMOUNT3 As New ADODB.Command
        With ComSF_GETAMOUNT3
                .Parameters.Append .CreateParameter("RETURN_VALUE", adVarNumeric, adParamReturnValue)
                .Parameters.Append .CreateParameter("P_CUSTID", adVarNumeric, adParamInput, , P_CUSTID)
                .Parameters.Append .CreateParameter("P_CITEMCODE", adVarNumeric, adParamInput, , P_CITEMCODE)
                .Parameters.Append .CreateParameter("P_PERIOD", adVarNumeric, adParamInput, , P_PERIOD)
                .Parameters.Append .CreateParameter("P_SERVICETYPE", adVarChar, adParamInput, 2000, p_SERVICETYPE)
                .Parameters.Append .CreateParameter("P_SHOULDDATE", adVarChar, adParamInput, 2000, P_SHOULDDATE)
                .Parameters.Append .CreateParameter("P_COMPCODE", adVarNumeric, adParamInput, , p_CompCode)
                '2007/05/09 Jacky 3216 Jim �[�@�ӰѼƨӸѨM�p�G�^�Ǫ��B�O�t�����D
                .Parameters.Append .CreateParameter("P_SHOULDAMT", adVarNumeric, adParamOutput, , p_OldAmt)
                .Parameters.Append .CreateParameter("P_RETMSG", adVarChar, adParamOutput, 2000, P_RETMSG)
                Set .ActiveConnection = gcnGi
                .CommandText = "SF_GETAMOUNT3"
                .CommandType = adCmdStoredProc
                .Execute
                p_OldAmt = Val(.Parameters("P_SHOULDAMT").Value & "")
                P_RETMSG = .Parameters("P_RETMSG").Value & ""
                CallSF_GETAMOUNT3 = Val(.Parameters("Return_value").Value & "")
        End With
  Exit Function
ChkErr:
    ErrSub FormName, "CallSF_GETAMOUNT3"
End Function

Public Function GetPermission(ByVal GaryID4 As String, ByVal strCode As String, Optional ByRef strPermission) As String
    On Error GoTo ChkErr
    Dim rs As New ADODB.Recordset
        strPermission = ""
        If GaryID4 = 0 Then Exit Function
        If Not GetRS(rs, "Select Mid From " & GetOwner & "SO029 Where Link = '" & strCode & "' And Group" & GaryID4 & "=1 Order by Mid") Then Exit Function
        If rs.RecordCount > 0 Then
            strPermission = "(" & rs.GetString(, , , ")(")
            If strPermission <> "" Then strPermission = Left(strPermission, Len(strPermission) - 1)
        End If
        
        GetPermission = GaryID4
    Exit Function
        Dim cmdUserPriv As New ADODB.Command
        Dim parPara0 As New ADODB.Parameter
        Dim parPara1 As New ADODB.Parameter
        Dim parPara2 As New ADODB.Parameter
        
        parPara0.Direction = adParamReturnValue
        parPara0.Type = adVarChar
        parPara0.Size = 2000
        cmdUserPriv.Parameters.Append parPara0
        
        parPara1.Direction = adParamInput
        parPara1.Type = adNumeric
        parPara1.Precision = 38
        parPara1.Value = GaryID4
        cmdUserPriv.Parameters.Append parPara1
        
        parPara2.Direction = adParamInput
        parPara2.Type = adVarChar
        parPara2.Size = 2000
        parPara2.Value = strCode
        cmdUserPriv.Parameters.Append parPara2
        
        Set cmdUserPriv.ActiveConnection = gcnGi
        
        cmdUserPriv.CommandType = adCmdText
        cmdUserPriv.CommandText = "{? = CALL SF_USERPRIV( ?, ?) }"
        cmdUserPriv.Execute
        strPermission = cmdUserPriv.Parameters(0).Value & ""
        GetPermission = cmdUserPriv.Parameters(1).Value & ""
    
    Exit Function
ChkErr:
    Call ErrSub("Utility3", "GetPermission")
End Function

'�]�w�U�ӥ\�઺�v���I
Public Sub SetPermission(ByVal strChkString As String _
                            , Optional ByRef blnAdd _
                            , Optional ByVal AddKey _
                            , Optional ByRef blnUpdate _
                            , Optional ByVal UpdKey _
                            , Optional ByRef blnDelete _
                            , Optional ByVal DelKey _
                            , Optional ByRef blnPrint _
                            , Optional ByVal PrintKey _
                            , Optional ByRef blnView _
                            , Optional ByVal ViewKey _
                            , Optional ByRef blnBatchUpdate _
                            , Optional ByVal BatchUpdateKey)
  On Error GoTo ChkErr
    If garyGi(4) <> 0 Then
        If Not (IsMissing(blnAdd) And IsMissing(AddKey)) Then
            blnAdd = IIf(InStr(strChkString, AddKey) > 0, True, False)
        End If
        If Not (IsMissing(blnUpdate) And IsMissing(UpdKey)) Then
            blnUpdate = IIf(InStr(strChkString, UpdKey) > 0, True, False)
        End If
        If Not (IsMissing(blnDelete) And IsMissing(DelKey)) Then
            blnDelete = IIf(InStr(strChkString, DelKey) > 0, True, False)
        End If
        If Not (IsMissing(blnPrint) And IsMissing(PrintKey)) Then
            blnPrint = IIf(InStr(strChkString, PrintKey) > 0, True, False)
        End If
        If Not (IsMissing(blnView) And IsMissing(ViewKey)) Then
            blnView = IIf(InStr(strChkString, ViewKey) > 0, True, False)
        End If
        If Not (IsMissing(blnBatchUpdate) And IsMissing(BatchUpdateKey)) Then
            blnBatchUpdate = IIf(InStr(strChkString, BatchUpdateKey) > 0, True, False)
        End If
    Else
        If Not IsMissing(blnAdd) Then blnAdd = True
        If Not IsMissing(blnUpdate) Then blnUpdate = True
        If Not IsMissing(blnDelete) Then blnDelete = True
        If Not IsMissing(blnPrint) Then blnPrint = True
        If Not IsMissing(blnView) Then blnView = True
        If Not IsMissing(blnBatchUpdate) Then blnBatchUpdate = True
        
    End If
  Exit Sub
ChkErr:
    Call ErrSub("Utility3", "ChkPermission")
End Sub

'�N����r���ഫ������榡
Public Function StrToDate(ByVal strDate As String, Optional ByVal blnAddTime As Boolean = False, _
    Optional ByVal blnAddMask As Boolean = True) As String
    On Error GoTo ChkErr
    Dim strMask As String
        If strDate = "" Then Exit Function
        If IsDate(strDate) Then
            If blnAddMask Then
                If blnAddTime Then
                    strMask = "yyyy/mm/dd hh:mm:ss"
                Else
                    strMask = "yyyy/mm/dd"
                End If
            Else
                If blnAddTime Then
                    strMask = "yyyymmddhhmmss"
                Else
                    strMask = "yyyymmdd"
                End If
            End If
        Else
            If blnAddMask Then
                If blnAddTime Then
                    strMask = "####/##/## ##:##:##"
                Else
                    strMask = "####/##/##"
                End If
            Else
                If blnAddTime Then
                    strMask = "##############"
                Else
                    strMask = "########"
                End If
            End If
        End If
        StrToDate = Format(strDate, strMask)
    Exit Function
ChkErr:
    ErrSub "Utility3", "StrToDate"
End Function

Public Function IsLastDate(ByVal strDate As String) As Boolean
    On Error Resume Next
        IsLastDate = strDate = GetLastDate(strDate)
End Function

'���o���ꪺ����榡
Public Function GetDateStr(ByVal strDate As String) As String
    On Error GoTo ChkErr
        GetDateStr = Format(strDate, "ee/MM/DD")
    Exit Function
ChkErr:
    ErrSub "Utility3", "GetDateStr"
End Function

Public Function ChkDataOk(ByVal strType As Integer, ByVal strBillNo As String, _
    ByVal strItem As Integer, strCitemName As String, _
    ByVal strCompCode As String, _
    Optional ByRef strMsg As String) As Boolean
    On Error GoTo ChkErr
    Dim intRet As Integer
        Dim ComSF_CANDEL5 As New ADODB.Command
        If strType = 0 Then
            If Val(GetRsValue("Select Count(*) From " & GetOwner & "So033 where BillNo ='" & strBillNo & "' And Item=" & strItem) & "") = 0 Then
                strType = 2
            Else
                strType = 0
            End If
        End If
        With ComSF_CANDEL5
                .Parameters.Append .CreateParameter("RETURN_VALUE", adVarNumeric, adParamReturnValue)
                .Parameters.Append .CreateParameter("P_TYPE", adVarNumeric, adParamInput, , strType)
                .Parameters.Append .CreateParameter("P_BILLNO", adVarChar, adParamInput, 2000, strBillNo)
                .Parameters.Append .CreateParameter("P_ITEM", adVarNumeric, adParamInput, , strItem)
                .Parameters.Append .CreateParameter("P_COMPCODE", adVarNumeric, adParamInput, , strCompCode)
                .Parameters.Append .CreateParameter("P_RETMSG", adVarChar, adParamOutput, 2000, strMsg)
                Set .ActiveConnection = gcnGi
                .CommandText = "SF_CANDEL5"
                .CommandType = adCmdStoredProc
                .Execute
                strMsg = .Parameters("P_RETMSG").Value & ""
                intRet = Val(.Parameters("Return_value").Value & "")
        End With
        If intRet <> 1 Then
            'strReturnMsg �Ǧ^�}
            MsgBox strMsg, vbExclamation, "�T��"
            Exit Function
        End If
        ChkDataOk = True
        Set ComSF_CANDEL5 = Nothing
    Exit Function
ChkErr:
    Call ErrSub("Utility3", "ChkDataOk")
End Function

Public Function SFGetAmount(ByVal blnChoose As Boolean, ByVal P_OPTION As Integer, _
                ByVal P_CUSTID As Long, _
                ByVal P_CITEMCODE As Long, _
                ByVal P_PERIOD As Integer, _
                ByVal P_REALSTARTDATE As String, _
                ByVal p_SERVICETYPE As String, _
                ByVal p_CompCode As Long, _
                ByRef P_REALSTOPDATE As String, _
                ByRef P_SHOULDAMT As Long, _
                ByRef P_RETMSG As String, _
                ByRef p_RealPeriod As Integer, _
                Optional ByVal P_BPCode As String, _
                Optional ByRef intPFlag1 As Integer, _
                Optional ByVal P_FaciSeqNo As String, _
                Optional ByVal p_PackageNo As String, _
                Optional ByVal p_PackageStepNo As Long, _
                Optional ByVal p_AmountType As Long, _
                Optional ByVal p_StepNo As Long, _
                Optional ByRef p_CrossStepDate As String, _
                Optional ByRef p_OrderNo As String, _
                Optional ByRef p_Expiretype As Long) As Integer
    On Error GoTo ChkErr
    
        If blnChoose Then
            SFGetAmount = SF_GetAmount2(P_OPTION, P_CUSTID, P_CITEMCODE, P_PERIOD, P_REALSTARTDATE, P_REALSTOPDATE, p_SERVICETYPE, p_CompCode, P_SHOULDAMT, P_RETMSG)
        Else
            SFGetAmount = SF_GetAmount(P_OPTION, P_CUSTID, P_CITEMCODE, P_PERIOD, P_REALSTARTDATE, p_SERVICETYPE, p_CompCode, P_REALSTOPDATE, P_SHOULDAMT, P_RETMSG, p_RealPeriod, P_BPCode, intPFlag1, P_FaciSeqNo, p_PackageNo, p_PackageStepNo, p_AmountType, p_StepNo, p_CrossStepDate, p_OrderNo, p_Expiretype)
        End If
    Exit Function
ChkErr:
    Call ErrSub("Utility3", "SFGetAmount")
End Function

Private Function SF_GetAmount(ByVal P_OPTION As Integer, _
                ByVal P_CUSTID As Long, ByVal P_CITEMCODE As Long, _
                ByVal P_PERIOD As Integer, ByVal P_REALSTARTDATE As String, _
                ByVal p_SERVICETYPE As String, ByVal p_CompCode As Long, _
                ByRef P_REALSTOPDATE As String, ByRef P_SHOULDAMT As Long, _
                ByRef P_RETMSG As String, ByRef p_RealPeriod As Integer, _
                Optional ByVal P_BPCode As String, _
                Optional ByRef P_PFlag1 As Integer, _
                Optional ByVal P_FaciSeqNo As String, Optional ByVal p_PackageNo As String, _
                Optional ByVal p_PackageStepNo As Long, Optional ByVal p_AmountType As Long, _
                Optional ByVal p_StepNo As Long, _
                Optional ByRef p_CrossStepDate As String, _
                Optional ByRef p_OrderNo As String, _
                Optional ByRef p_Expiretype As Long, _
                Optional ByRef p_BPCode1 As String, Optional ByRef p_BPName As String, _
                Optional ByRef p_PromCode As Long, Optional ByRef p_PromName As String) As Integer
                
    On Error GoTo ChkErr
    Dim comSF_GETAMOUNT As New ADODB.Command
    
    With comSF_GETAMOUNT
            .Parameters.Append .CreateParameter("RETURN_VALUE", adVarNumeric, adParamReturnValue)
            .Parameters.Append .CreateParameter("P_OPTION", adVarNumeric, adParamInput, , P_OPTION)
            .Parameters.Append .CreateParameter("P_CUSTID", adVarNumeric, adParamInput, , P_CUSTID)
            .Parameters.Append .CreateParameter("P_CITEMCODE", adVarNumeric, adParamInput, , P_CITEMCODE)
            .Parameters.Append .CreateParameter("P_PERIOD", adVarNumeric, adParamInput, , P_PERIOD)
            .Parameters.Append .CreateParameter("P_REALSTARTDATE", adVarChar, adParamInput, 2000, P_REALSTARTDATE)
            .Parameters.Append .CreateParameter("P_SERVICETYPE", adVarChar, adParamInput, 2000, p_SERVICETYPE)
            .Parameters.Append .CreateParameter("P_COMPCODE", adVarNumeric, adParamInput, , p_CompCode)
            .Parameters.Append .CreateParameter("P_BPCODE", adVarChar, adParamInput, 2000, P_BPCode)
            .Parameters.Append .CreateParameter("P_FACISEQNO", adVarChar, adParamInput, 2000, P_FaciSeqNo)
            
            .Parameters.Append .CreateParameter("P_PACKAGENO", adVarChar, adParamInput, 2000, p_PackageNo)
            .Parameters.Append .CreateParameter("P_PACKAGESTEPNO", adVarNumeric, adParamInput, , p_PackageStepNo)
            .Parameters.Append .CreateParameter("P_AMOUNTTYPE", adVarNumeric, adParamInput, , p_AmountType)
            .Parameters.Append .CreateParameter("P_STEPNO", adVarNumeric, adParamInput, , p_StepNo)
            
            .Parameters.Append .CreateParameter("P_REALSTOPDATE", adVarChar, adParamOutput, 2000, P_REALSTOPDATE)
            .Parameters.Append .CreateParameter("P_SHOULDAMT", adVarNumeric, adParamOutput, , P_SHOULDAMT)
            .Parameters.Append .CreateParameter("P_REALPERIOD", adVarNumeric, adParamOutput, , p_RealPeriod)
            .Parameters.Append .CreateParameter("P_PUNISH", adVarNumeric, adParamOutput, , 0)
            .Parameters.Append .CreateParameter("P_DISCOUNTDATE1", adVarChar, adParamOutput, 2000, "")
            .Parameters.Append .CreateParameter("P_DISCOUNTDATE2", adVarChar, adParamOutput, 2000, "")
            .Parameters.Append .CreateParameter("P_PFLAG1", adVarNumeric, adParamOutput, , P_PFlag1)
            .Parameters.Append .CreateParameter("p_CrossStepDate", adVarChar, adParamOutput, 2000, p_CrossStepDate)
            .Parameters.Append .CreateParameter("p_OrderNo", adVarChar, adParamOutput, 2000, p_OrderNo)
            .Parameters.Append .CreateParameter("p_Expiretype", adVarNumeric, adParamOutput, , p_Expiretype)
            .Parameters.Append .CreateParameter("p_BPCode1", adVarChar, adParamOutput, 2000, p_BPCode1)
            .Parameters.Append .CreateParameter("p_BPName", adVarChar, adParamOutput, 2000, p_BPName)
            .Parameters.Append .CreateParameter("p_PromCode", adVarNumeric, adParamOutput, , p_PromCode)
            .Parameters.Append .CreateParameter("p_PromName", adVarChar, adParamOutput, 2000, p_PromName)
            .Parameters.Append .CreateParameter("P_RETMSG", adVarChar, adParamOutput, 2000, P_RETMSG)
            Set .ActiveConnection = gcnGi
            .CommandText = "SF_GETAMOUNT"
            .CommandType = adCmdStoredProc
            .Execute
            P_REALSTOPDATE = .Parameters("P_REALSTOPDATE").Value & ""
            p_RealPeriod = Val(.Parameters("p_RealPeriod").Value & "")
            P_SHOULDAMT = Val(.Parameters("P_SHOULDAMT").Value & "")
            P_PFlag1 = Val(.Parameters("P_PFLAG1").Value & "")
            
            p_CrossStepDate = .Parameters("p_CrossStepDate") & ""
            
            P_RETMSG = .Parameters("P_RETMSG").Value & ""
            SF_GetAmount = Val(.Parameters("Return_value").Value & "")
    End With
    Exit Function
ChkErr:
    Call ErrSub("Utility3", "SF_GetAmount")
End Function

Public Function SF_GetAmount2(ByVal P_OPTION As Integer, _
                ByVal P_CUSTID As Long, _
                ByVal P_CITEMCODE As Long, _
                ByVal P_PERIOD As Integer, _
                ByVal P_REALSTARTDATE As String, _
                ByVal P_REALSTOPDATE As String, _
                ByVal p_SERVICETYPE As String, _
                ByVal p_CompCode As Long, _
                ByRef P_SHOULDAMT As Long, _
                ByRef P_RETMSG As String) As Integer
    On Error GoTo ChkErr
    Dim comSF_GETAMOUNT As New ADODB.Command
    
    With comSF_GETAMOUNT
            .Parameters.Append .CreateParameter("RETURN_VALUE", adVarNumeric, adParamReturnValue)
            .Parameters.Append .CreateParameter("P_OPTION", adVarNumeric, adParamInput, , P_OPTION)
            .Parameters.Append .CreateParameter("P_CUSTID", adVarNumeric, adParamInput, , P_CUSTID)
            .Parameters.Append .CreateParameter("P_CITEMCODE", adVarNumeric, adParamInput, , P_CITEMCODE)
            .Parameters.Append .CreateParameter("P_PERIOD", adVarNumeric, adParamInput, , P_PERIOD)
            .Parameters.Append .CreateParameter("P_REALSTARTDATE", adVarChar, adParamInput, 2000, P_REALSTARTDATE)
            .Parameters.Append .CreateParameter("P_REALSTOPDATE", adVarChar, adParamInput, 2000, P_REALSTOPDATE)
            .Parameters.Append .CreateParameter("P_SERVICETYPE", adVarChar, adParamInput, 2000, p_SERVICETYPE)
            .Parameters.Append .CreateParameter("P_COMPCODE", adVarNumeric, adParamInput, , p_CompCode)
            .Parameters.Append .CreateParameter("P_SHOULDAMT", adVarNumeric, adParamOutput, , P_SHOULDAMT)
            .Parameters.Append .CreateParameter("P_RETMSG", adVarChar, adParamOutput, 2000, P_RETMSG)
            Set .ActiveConnection = gcnGi
            .CommandText = "SF_GETAMOUNT2"
            .CommandType = adCmdStoredProc
            .Execute
            P_SHOULDAMT = Val(.Parameters("P_SHOULDAMT").Value & "")
            P_RETMSG = .Parameters("P_RETMSG").Value & ""
            SF_GetAmount2 = Val(.Parameters("Return_value").Value & "")
    End With
    Exit Function
ChkErr:
    Call ErrSub("Utility3", "SF_GetAmount2")
End Function

Public Function IsDateOk(ByVal strDate As String) As Boolean
  On Error GoTo ChkErr
    Dim strChkDate As String
    IsDateOk = False
    strChkDate = StrToDate(strDate)
    If Not IsDate(strChkDate) Then Exit Function
    If strChkDate = "" Then Exit Function
    If IsDate(strChkDate) And CDate(strChkDate) <= CDate("1990/01/01") Then Exit Function
    IsDateOk = True
  Exit Function
ChkErr:
    Call ErrSub("Utility3", "IsDateOk")
End Function

Public Function IstmpRsDup(ByVal rsTmp As ADODB.Recordset, ByVal FieldName As String, ByVal Value As Variant) As Boolean
  On Error GoTo ChkErr
    Dim rsdup As New ADODB.Recordset
    IstmpRsDup = False
        Set rsdup = rsTmp.Clone
        If rsdup.RecordCount > 0 Then rsdup.MoveFirst
        With rsdup
            While Not rsdup.EOF
                If .Fields(FieldName) = Value Then
                    IstmpRsDup = True
                    GoTo lblResult
                End If
                .MoveNext
            Wend
        End With
lblResult:
    'OldValue = Value
    'OldResult = IstmpRsDup
    rsdup.Close
    Set rsdup = Nothing
    Exit Function
ChkErr:
    Call ErrSub("Utility3", "IsDupData")
End Function

Public Sub ClearNoCnRs(ByRef rsData As ADODB.Recordset)
  On Error GoTo ChkErr
    If rsData.RecordCount <= 0 Or rsData.State <> adStateOpen Then Exit Sub
    With rsData
        .MoveFirst
        While Not .EOF
            .Delete
            .MoveNext
        Wend
    End With
  Exit Sub
ChkErr:
    Call ErrSub("clsModule", "ClearNoCnRs")
End Sub

'Check folder is not exist
Public Function ChkFolderExist(ByVal FolderPath As String) As Boolean
  On Error GoTo ChkErr
    Dim Folder As Object, strTmp As Variant
    strTmp = Split(FolderPath, "\")
    If Dir(FolderPath) = strTmp(UBound(strTmp)) Then
        ChkFolderExist = True
    End If
    Set Folder = CreateObject("Scripting.FileSystemObject")
    Dim intLoop As Integer, strPath As String
    For intLoop = LBound(strTmp) To UBound(strTmp) - 1
      '  Debug.Print strtmp(intLoop)
        strPath = strTmp(intLoop) & "\"
    Next
    ChkFolderExist = Folder.FolderExists(strPath)
    Set Folder = Nothing
  Exit Function
ChkErr:
    Call ErrSub("Utility3", "ChkFolderExist")
End Function

'�Ǧ^�Y�~�Y�몺�̫�@��
Public Function GetLastDay(Optional ByVal strYY As String, Optional ByVal strMM As String) As String
On Error GoTo ChkErr
Dim strDate As String
Dim intChkVal As Integer
    strDate = strYY & "/" & strMM & "/"
    For intChkVal = 31 To 28 Step -1
        If IsDate(strDate & intChkVal) Then Exit For
    Next
    GetLastDay = intChkVal
Exit Function
ChkErr:
     Call ErrSub("Utility3", "GetLastDay")
End Function

Private Function ChargeBillChoose(strForm As String, objForm As Object, _
        ByRef rs As ADODB.Recordset, ByRef lngBatchNo As Long, ByRef strSQL33 As String, _
        ByRef blnClose As Boolean) As Boolean
    On Error GoTo ChkErr
    Dim strSQL As String
    Dim strSQL1 As String
    Dim strUCSQL As String
    Dim strBillField As String
    Dim strBillNo1 As String
    Dim strBillNo2 As String
    Dim strCustId As String
        lngBatchNo = GetValueFromRs2("SELECT " & GetOwner & "S_SO075_BatchNo.NextVal FROM DUAL", gcnGi)
        Select Case UCase(strForm)
            Case "FRMSO3261A"
                With objForm
                    '2.  �]�w�ܼƪ��,
                    'SQL1 = "select distinct A.PrtSNo, A.CustId, <�C�L�帹> from SO033 A where <��ɮɶ�> = �ثe�ɶ�
                    If .cboBillType.ListIndex = 0 Then
                        strSQL1 = "SELECT distinct " & lngBatchNo & ", A.PrtSNo, A.CustId, null from " & GetOwner & "SO033 A where "
                    ElseIf .cboBillType.ListIndex = 1 Then
                        strSQL1 = "SELECT distinct " & lngBatchNo & ", null, A.CustId, A.BillNo from " & GetOwner & "SO033 A where "
                    Else
                        strSQL1 = "SELECT distinct " & lngBatchNo & ", A.MediaBillNo, A.CustId, Null from " & GetOwner & "SO033 A where "
                    End If
                    ' 4.  �Y���q�{���Ω���C�L��
                    If strForm = "FRMSO3261A" Then
                        ' �h�R����ڦC�Llog��(so060)���L����W�ٵ���<p_PrtName>�����
                        If ExecuteSQL("DELETE FROM " & GetOwner & "SO060 WHERE PRINTERNAME = '" & .cboPrinterName.Text & "'", gcnGi, , False) Then
                            Call ErrHandle("�L�k�R���Ȧs�� SO060 �̪���ơC", False, 0, "ChargeBillChoose")
                        End If
                    End If
                    strSQL33 = ""
                    ' 5.  ��_�d�߱���:
                    ' �E  �Y�����~�릳��:
                    If .gymClctYM.GetValue <> "" Then
                        'SQL = SQL + " and A.ClctYM=<p_ClctYM> "
                        strSQL = strSQL & " AND A.ClctYM = '" & .gymClctYM.GetValue & "' "
                    End If
                    '�E  �Y�C�L�ﶵ=1:
                    If .optOption2(1).Value Then
                        'SQL = SQL + " and A.PrtCount=0"
                        strSQL = strSQL & " AND A.PrtCount = 0 "
                    End If
                    '�E  �Y�����O�覡����:
                    If .gimCM.GetQryStr <> "" Then
                        'SQL=SQL+" and A.CMCode"+<p_CMSQL>
                        strSQL = strSQL & " AND A.CMCode " & .gimCM.GetQryStr & " "
                    End If
                    '�E  �Y�����O���ر���:
                    If .gimCitemCode.GetQryStr <> "" Then
                        'SQL=SQL+" and A.CitemCode"+<p_CitemSQL>
                        strSQL = strSQL & " AND A.CitemCode " & .gimCitemCode.GetQryStr & " "
                    End If
                    
                    '�E  �Y���Ȥ����O����:
                    If .gimClassCode.GetQryStr <> "" Then
                        'SQL=SQL+" and A.ClassCode"+<p_ClassSQL>
                        strSQL = strSQL & " AND A.ClassCode " & .gimClassCode.GetQryStr & " "
                    End If
                    '�E  �Y���A�Ȱϱ���:
                    If .gimServCode.GetQryStr <> "" Then
                        'SQL=SQL+" and A.ServCode"+<p_ServSQL>
                        strSQL = strSQL & " AND A.ServCode " & .gimServCode.GetQryStr & " "
                    End If
                    '�E  �Y����F�ϱ���:
                    If .gimAreaCode.GetQryStr <> "" Then
                        'SQL=SQL+" and A.AreaCode"+<p_AreaSQL>
                        strSQL = strSQL & " AND A.AreaCode " & .gimAreaCode.GetQryStr & " "
                    End If
                    '�E  �Y�����O�ϱ���:
                    If .gimClctAreaCode.GetQryStr <> "" Then
                        'SQL=SQL+" and A.ClctAreaCode"+<p_ClctAreaSQL>
                        strSQL = strSQL & " AND A.ClctAreaCode " & .gimClctAreaCode.GetQryStr & " "
                    End If
                    '�E  �Y���j�ӱ���:
                    If .gimMduId.GetQryStr <> "" Then
                        'SQL=SQL+" and A.MduId"+<p_MduSQL>
                        strSQL = strSQL & " AND A.MduID " & .gimMduId.GetQryStr & " "
                    End If
                    '�E  �Y�����O������:
                    If .gimClctEn.GetQryStr <> "" Then
                        'SQL=SQL+" and A.ClctEn"+<p_ClctEnSQL>
                        strSQL = strSQL & " AND A.ClctEn " & .gimClctEn.GetQryStr & " "
                    End If
                    '�E  �Y�����ͤH������: 901031
                    If .gimCreateEn.GetQryStr <> "" Then
                        'SQL=SQL+" and A.ClctEn"+<p_ClctEnSQL>
                        strSQL = strSQL & " AND A.CreateEn " & .gimCreateEn.GetQryStr & " "
                    End If
                    '�E  �Y�����ͮɶ�����: 901031
                    If .gdaCreateTime1.GetValue <> "" Then
                        strSQL = strSQL & " AND A.CreateTime >= " & GetNullString(.gdaCreateTime1.GetValue(True), giDateV, giOracle)
                    End If
                    If .gdaCreateTime2.GetValue <> "" Then
                        strSQL = strSQL & " AND A.CreateTime < " & GetNullString(CDate(.gdaCreateTime2.GetValue(True)) + 1, giDateV, giOracle)
                    End If
                    '�E  �Y�������ɶ�����: 901031
                    If .gdaShouldDate1.GetValue <> "" Then
                        strSQL = strSQL & " AND A.ShouldDate >= " & GetNullString(.gdaShouldDate1.GetValue(True), giDateV, giOracle)
                    End If
                    If .gdaShouldDate2.GetValue <> "" Then
                        strSQL = strSQL & " AND A.ShouldDate < " & GetNullString(CDate(.gdaShouldDate2.GetValue(True)) + 1, giDateV, giOracle)
                    End If
                    
                    '�E  �Y����D����:
                    If .gimStrtCode.GetQryStr <> "" Then
                        'SQL=SQL+" and A.StrtCode"+<p_StrtSQL>
                        strSQL = strSQL & " AND A.StrtCode " & .gimStrtCode.GetQryStr & " "
                    End If
                    
                    '�E  �Y��ڸ��X����:
                    If .cboBillType.ListIndex = 0 Then
                        strBillField = "A.PrtSNo"
                    ElseIf .cboBillType.ListIndex = 1 Then
                        strBillField = "A.BillNo"
                    Else
                        strBillField = "A.MediaBillNo"
                    End If
                    'SQL=SQL+" and A.PrtSNo>='"+<p_PrtSNo1>+" ' and A.PrtSNo<=' "+<p_PrtSNo2>+" ' "
                    If .txtPrtSNo1.Text <> "" And .txtPrtSNo2.Text <> "" Then
                        strSQL = strSQL & " AND " & strBillField & " >= '" & .txtPrtSNo1.Text & "' and " & strBillField & " <= '" & .txtPrtSNo2.Text & "'"
                    Else
                        strSQL = strSQL & " AND " & strBillField & " Is Not Null "
                    End If
                    
                    If Not (.chkIncludeUC.Value = 1 And .chkIncludeCancel.Value = 1) Then
                        '�������
                        strUCSQL = " OR (UCCode Is Not Null And CancelFlag = 0) "
                        
                        '�]�t�w��
                        If .chkIncludeUC.Value = 1 Then
                            strUCSQL = strUCSQL & " OR (UCCode Is Null And CancelFlag = 0) "
                        End If
                        '�]�t�@�o
                        If .chkIncludeCancel.Value = 1 Then
                            strUCSQL = strUCSQL & " OR CancelFlag =1 "
                        End If
                    End If
                    
                    If strUCSQL <> "" Then
                        strSQL = strSQL & " And (" & Mid(strUCSQL, 4) & ") "
                    End If
                    
                    If .gimUCCode.GetQryStr <> "" Then
                        strSQL = strSQL & " And A.UCCode " & .gimUCCode.GetQryStr
                    End If
                                    
                    '�C�L���B��0
                    If .chkPrintZero.Value = 0 Then
                        strSQL = strSQL & " And A.ShouldAmt <> 0 "
                    End If
                    
                    '���q�O
                    If .gilCompCode.GetCodeNo <> "" Then
                        strSQL = strSQL & " AND A.Compcode = " & .gilCompCode.GetCodeNo
                    End If
                    
                    '�A�����O
                    If .gilServiceType.GetCodeNo <> "" Then
                        strSQL = strSQL & " AND A.ServiceType = '" & .gilServiceType.GetCodeNo & "'"
                    End If
                    
                    '�Ȧ�O
                    If .gimBank.GetDispStr <> "" Then
                        strSQL = strSQL & " AND A.BankCode " & .gimBank.qryStr
                    End If
                    
                    strSQL = strSQL & " AND ((A.RealDate is null) OR (A.ClctEn is not null)) "
                    '�E  �h��SQL�̥��䤧" and "
                    If strSQL <> "" Then
                        strSQL33 = strSQL
                        If Left(strSQL, 4) = " AND" Then
                            strSQL = Mid(strSQL, 5)
                        End If
                    End If
                    '�E  SQL=SQL1+SQL
                    strSQL = strSQL1 & strSQL
                    '6.  �i��Ĥ@���q�d��, �ñN���G�s�J�ݦC�L�Ȥ�Ȧs�ɤ�, SQL:
                    '"insert into ????? (PrtSNo, CustId) (<SQL>)"
                    If ExecuteSQL("INSERT INTO " & GetOwner & "SO075 (BatchNO, PrtSNo, CustId, BillNo) (" & strSQL & " )", gcnGi, , False) <> giOK Then
                        Call ErrHandle("�L�k���Ʀs�J SO075 �A���{���Y�N�����C", False, 2, "ChargeBillChoose")
                        Exit Function
                    End If
                    '�_�h: SQL = "select A.PrtSNo, A.CustId from ????? A"
                    strSQL = "SELECT Distinct A.PrtSNo, A.CustId, BillNo FROM " & GetOwner & "SO075 A WHERE A.BatchNo = " & lngBatchNo & " Order By A.PrtSNo"
                End With
            Case "FRMSO3262A"
                With objForm
                    blnClose = (.chkClose = 1)
                    If .optBillNo Then
                        strBillNo1 = .mskBillNo1.Text
                        strBillNo2 = .mskBillNo2.Text
                    Else
                        strCustId = .txtCustId.Text
                    End If
                    If Not PutIntoSO075(strSQL, lngBatchNo, strCustId, strBillNo1, strBillNo2, , blnClose) Then Exit Function
                    strSQL33 = " And CancelFlag = 0 And CompCode = " & .gilCompCode.GetCodeNo & " AND A.ServiceType = '" & .gilServiceType.GetCodeNo & "'"
                    If .gimCitemCode.GetQryStr <> "" Then strSQL33 = strSQL33 & " And CitemCode " & .gimCitemCode.GetQryStr
                    If .optBillNo Then
                        strSQL33 = " AND BillNo >= '" & .mskBillNo1 & "' and BillNo <= '" & .mskBillNo2 & "'" & strSQL33
                    ElseIf .optCustId Then
                        strSQL33 = " AND ClctYM >= " & .gymClctYM1.GetValue & " and ClctYM <= " & .gymClctYM2.GetValue & strSQL33
                    Else
                        strSQL33 = " AND ClctYM >= " & .gymClctYM1.GetValue & " and ClctYM <= " & .gymClctYM2.GetValue & strSQL33
                    End If
                End With
            Case "FRMSO1100BMDI"
                With objForm.rsSO033SO034
                    If Not PutIntoSO075(strSQL, lngBatchNo, .Fields("CustId") & "", .Fields("BillNo") & "", .Fields("BillNo") & "", .Fields("Item")) Then Exit Function
                    If objForm.ChkMasterDetail And InStr(1, objForm.cboSort, "�C��") > 0 Then
                        strSQL33 = " And MediaBillNo ='" & .Fields("MediaBillNo") & "' And UCCode is Not Null  And CancelFlag = 0 "
                    Else
                        If Len(.Fields("PrtSNo") & "") = 0 Then
                            strSQL33 = " And BillNo ='" & .Fields("BillNo") & "' And UCCode is Not Null  And CancelFlag = 0 "
                        Else
                            strSQL33 = " And PrtSNo ='" & .Fields("PrtSNo") & "' And UCCode is Not Null  And CancelFlag = 0 "
                        End If
                    End If
                End With
            Case "FRMSO3273A"
                With objForm
                    lngBatchNo = .uBatchNo
                    strSQL33 = .uSQL33
                    strSQL = "SELECT * FROM " & GetOwner & "SO075 WHERE BatchNO = " & lngBatchNo & " Order By BillNo"
'
'
'                    '�E  �Y�������������: 920811
'                    If .gdaShouldDate1.GetValue <> "" Then
'                        strSQL = strSQL & " AND A.ShouldDate >= " & GetNullString(.gdaShouldDate1.GetValue(True), giDateV, giOracle)
'                    End If
'                    If .gdaShouldDate2.GetValue <> "" Then
'                        strSQL = strSQL & " AND A.ShouldDate < " & GetNullString(CDate(.gdaShouldDate2.GetValue(True)) + 1, giDateV, giOracle)
'                    End If
'
'                    '�E  �Y�����q�O����: 920811
'                    If .gilCompCode.GetCodeNo <> "" Then
'                        'SQL=SQL+" and A.CompCode"+<p_ServiceType>
'                        strSQL33 = strSQL33 & " AND A.CompCode = " & .gilCompCode.GetCodeNo & " "
'                    End If
'
'                    '�E  �Y���A�����O����: 920811
'                    If .gilServiceType.GetCodeNo <> "" Then
'                        'SQL=SQL+" and A.ServiceType"+<p_ServiceType>
'                        strSQL33 = strSQL33 & " AND A.ServiceType = '" & .gilServiceType.GetCodeNo & "' "
'                    End If
'                    '�E  �Y�����O���ر���:
'                    If .gimCitemCode.GetQryStr <> "" Then
'                        'SQL=SQL+" and A.CitemCode"+<p_CitemSQL>
'                        strSQL33 = strSQL33 & " AND A.CitemCode " & .gimCitemCode.GetQryStr & " "
'                    End If
'                    '�E  �Y�����O�覡����:
'                    If .gmdCMCode.GetQryStr <> "" Then
'                        'SQL=SQL+" and A.CMCode"+<p_CMSQL>
'                        strSQL33 = strSQL33 & " AND A.CMCode " & .gmdCMCode.GetQryStr & " "
'                    End If
'                    '�E  �Y���Ȥ����O����:
'                    If .gmdCustClass.GetQryStr <> "" Then
'                        'SQL=SQL+" and A.ClassCode"+<p_ClassSQL>
'                        strSQL33 = strSQL33 & " AND A.ClassCode " & .gmdCustClass.GetQryStr & " "
'                    End If
'                    '�E  �Y�����O������:
'                    If .gmdClctEn.GetQryStr <> "" Then
'                        'SQL=SQL+" and A.ClctEn"+<p_ClctEnSQL>
'                        strSQL33 = strSQL33 & " AND A.ClctEn " & .gmdClctEn.GetQryStr & " "
'                    End If
'                    '�E  �Y��������O����:
'                    If .gmdBillType.GetQryStr <> "" Then
'                        'SQL=SQL+" and A.ClctEn"+<p_ClctEnSQL>
'                        strSQL33 = strSQL33 & " AND SubStr(A.BillNo,7,1) " & .gmdBillType.GetQryStr & " "
'                    End If
'                    '�E  �Y�����ͤH������:
'                    If .gmdCreateEn.GetQryStr <> "" Then
'                        'SQL=SQL+" and A.ClctEn"+<p_ClctEnSQL>
'                        strSQL33 = strSQL33 & " AND CreateEn " & .gmdCreateEn.GetQryStr & " "
'                    End If
'
'                    strSQL33 = strSQL33 & " And UCCode is Not Null And CancelFlag = 0 "
'                    strSQL = "SELECT * FROM " & GetOwner & "SO075 WHERE BatchNO = " & lngBatchNo & " Order By BillNo"
                End With
            Case "FRMSO4000A"
                With objForm
                    blnClose = True
                    lngBatchNo = .uBatchNo
                    strSQL33 = .uSQL33
                    strSQL = "SELECT * FROM " & GetOwner & "SO075 WHERE BatchNO = " & lngBatchNo & " Order By BillNo"
                End With
            Case "FRMSO32A2A"
                With objForm
                    blnClose = True
                    lngBatchNo = .uBatchNo
                    strSQL33 = .uSQL33
                    strSQL = "SELECT * FROM " & GetOwner & "SO075 WHERE BatchNO = " & lngBatchNo & " Order By BillNo"
                End With
        End Select
        '8.  ����SQL, Loop�C�@��, ���ͬ�����ڸ��, �ѦҪ���y�{��
        If Not GetRS(rs, strSQL, gcnGi, adUseClient, adOpenForwardOnly, adLockReadOnly) Then Exit Function
        
        ChargeBillChoose = True
    Exit Function
ChkErr:
    ErrSub "Utility3", "ChargeBillChoose"
End Function

Private Function PutIntoSO075(ByRef strSQL As String, lngBatchNo As Long, _
        strCustIdStr As String, Optional strBillNo1 As String = "", Optional strBillNo2 As String = "", _
        Optional intItem As Integer, Optional blnClose As Boolean = False) As Boolean
    On Error GoTo ChkErr
    Dim ary()
    Dim lngX As Long
    Dim strAll As String
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL2 As String
    Dim strSQL3 As String
    Dim StrTableName As String
        If Trim(strBillNo1) <> "" Then
            If intItem > 0 Then strSQL2 = " And Item = " & intItem
            
            StrTableName = "SO033"
            strSQL3 = " And UCCode Is Not Null "
            If blnClose Then StrTableName = "SO034": strSQL3 = ""
            strSQL2 = strSQL2 & strSQL3
            gcnGi.Execute "INSERT INTO " & GetOwner & "SO075 (PrtSNO, CustID, BatchNO) Select distinct PrtSNo,CustId ," & lngBatchNo & " From " & GetOwner & StrTableName & " Where BillNo Between '" & strBillNo1 & "' And '" & strBillNo2 & "' " & strSQL2, lngX
        Else
            Call ParseWord(strCustIdStr, ary)
            On Error Resume Next
            'Debug.Print UBound(ary)
            If Err.Number <> 0 Then
                Exit Function
            End If
            For lngX = 0 To UBound(ary)
                CloseRecordset rsTmp
                If ary(lngX) = 9999 Then
                    Beep
                End If
                If Not GetRS(rsTmp, "SELECT CUSTID FROM " & GetOwner & "SO001 WHERE CustID = " & ary(lngX), gcnGi) Then Exit Function
                If rsTmp.RecordCount > 0 Then
                    If OpenDB.ExecuteSQL("INSERT INTO " & GetOwner & "SO075 (PrtSNO, CustID, BatchNO) VALUES (NULL, " & ary(lngX) & ", " & lngBatchNo & " )", gcnGi, , False) <> giOK Then
                        Call ErrHandle("�L�k���Ƽg�J SO075 �A�o�ӵ{���Y�N�����C", False, 2, "PutIntoSO075")
                        Exit Function
                    End If
                End If
            Next
        End If
        strSQL = "SELECT distinct PrtSNo, CustId, BillNo FROM " & GetOwner & "SO075 WHERE BatchNO = " & lngBatchNo & " Order By PrtSNo"
        
        Call Close3Recordset(rsTmp)
        PutIntoSO075 = True
    Exit Function
ChkErr:
    ErrSub FormName, "PutIntoSo075"
End Function

Private Function ChargeBillField(cnRptMDB As ADODB.Connection, _
        strFmtID As String, ByRef strSQLB As String, _
        ByRef strSQLC1 As String, ByRef strSQLC2 As String, ByRef strSQLC3 As String, _
        ByRef strSQLD As String, ByRef strSQLE As String, ByRef strSQLF As String, _
        ByRef strSQLI As String, ByRef strSQLJ As String, ByRef strSQLK As String, _
        ByRef strSQLL As String, ByRef strSQLL1 As String, ByRef strSQLL2 As String, _
        ByRef strSQLM As String, ByRef strSQLN1 As String, ByRef strSQLN2 As String, _
        ByRef strSQLOA As String, ByRef strSQLOB As String, _
        ByRef strFieldName As String, ByRef blnJoin2 As Boolean, _
        cnTmpMDB As ADODB.Connection, ByRef strSQLOC As String, ByRef strSQLOD As String) As Boolean
    On Error GoTo ChkErr
    Dim rsFmtFld As New ADODB.Recordset
    Dim lngX As Long
    Dim aryF() As String
    Dim strF As String
    Dim strMDBPath As String
        '1.��ȳ]�w: �ѫ���d�ߥΪ��ܼ�
        'SQLB = "select InstAddrNo, ChargeAddrNo, MduId"         (��: �o����쥲�����o)
        strSQLB = "SELECT InstAddrNo, ChargeAddrNo, MduId"
        
        'SQLC1 = SQLC2 = SQLD = SQLE = SQLF = SQLI = �Ŧr��
        'xSQLC1 = xSQLC2 = xSQLD = xSQLE = xSQLF = xSQLI = �Ŧr��
        '�}�CaryF: �ΨӦs��e���ꦬ��Ʃһݪ����W��
        ReDim Preserve aryF(0 To 0) As String
        
        '2. Loop���t�\����T��(FmtFld)�����O'A'�}�Y���C�@��:
'        If Not GetRS(rsFmtFld, "SELECT FmtFld.*, FldList.FldName FROM FmtFld, " & _
'                "FldList WHERE FmtFld.FldId = FldList.FldID AND FmtFld.FmtID = '" & _
'                strFmtID & "' AND MID(FmtFld.FldID, 1, 1) <> 'A' ORDER BY FmtFld.FldID", _
'                cnRptMDB, adUseClient, adOpenForwardOnly, adLockReadOnly) Then
'            Call errHandle("�L�k�}�� FmtFld ����ơA���{���Y�N�����C", False, 2, "PrepareA")
'            Exit Function
'        End If

        If Not GetRS(rsFmtFld, "SELECT FmtFld.FmtID,FmtFld.FldId,FmtFld.Type, FldList.FldName, FldList.TmpTable,FldList.Length,FldList.FldType FROM FmtFld, " & _
                "FldList WHERE FmtFld.FldId = FldList.FldID And FmtFld.Type = FldList.Type AND FmtFld.FmtID = '" & _
                strFmtID & "'  AND MID(FmtFld.FldID, 1, 1) <> 'A' " & _
                "Union All SELECT FmtFld.FmtID,FmtFld.FldId,FmtFld.Type, FldList.FldName, FldList.TmpTable,FldList.Length,FldList.FldType FROM FmtFld, " & _
                "FldList WHERE FmtFld.FldId = FldList.FldID And FmtFld.Type = FldList.Type AND FmtFld.FmtID = '" & _
                strFmtID & "'  AND MID(FmtFld.FldID, 1, 2)  >= 'AA' AND MID(FmtFld.FldID, 1, 2)  <= 'AZ' " & _
                "Union All SELECT '" & strFmtID & "' as FmtID, FldId,1 as Type,FldName,TmpTable,Length,FldType FROM  " & _
                "FldList WHERE MID(FldID, 1, 2) >= 'A0' And MID(FldID, 1, 2) <= 'A9' ORDER BY FldID", cnRptMDB, adUseClient, adOpenForwardOnly, adLockReadOnly) Then
            Call ErrHandle("�L�k�}�� FmtFld ����ơA���{���Y�N�����C", False, 2, "PrepareA")
            Exit Function
        End If
        strFieldName = ""
        blnJoin2 = False
        Do While Not rsFmtFld.EOF
            'Debug.Print rsFmtFld("FldID").Value
            strFieldName = strFieldName & "," & rsFmtFld("FldID")
            'Debug.Print rsFmtFld("TmpTable") & "," & rsFmtFld("FldID")
            If rsFmtFld("TmpTable") = 2 Then
                blnJoin2 = True
            End If
            
            If rsFmtFld("FldName").Value <> "" And Mid(rsFmtFld("FldID"), 1, 1) <> "A" Then
                '�E  case 'B'�}�Y, �B<���W��>����:
                If UCase(Left(rsFmtFld("FldID").Value, 1)) = "B" Then
                    'SQLB = SQLB + ", <���W��>"
                    strSQLB = strSQLB & ", " & rsFmtFld("FldName").Value
                End If
                
                '�E  case 'C01', 'C02', 'C03','C07','CAxx' �B<���W��>����:
                If UCase(rsFmtFld("FldID").Value) = "C01" Or UCase(rsFmtFld("FldID").Value) = _
                   "C02" Or UCase(rsFmtFld("FldID").Value) = "C03" Or UCase(rsFmtFld("FldID").Value) = "C07" _
                    Or UCase(Left(rsFmtFld("FldID").Value, 2)) = "CA" Then
                    'SQLC1 = SQLC1 + ", <���W��>"
                    strSQLC1 = strSQLC1 + ", " & rsFmtFld("FldName").Value
                End If
                
                '�E  case 'C04', 'C05','C06','CBxx' �B<���W��>����:
                If UCase(rsFmtFld("FldID").Value) = "C04" Or UCase(rsFmtFld("FldID").Value) = _
                   "C05" Or UCase(rsFmtFld("FldID").Value) = "C06" _
                    Or UCase(Left(rsFmtFld("FldID").Value, 2)) = "CB" Then
                    'SQLC2 = SQLC2 + ", <���W��>"
                    strSQLC2 = strSQLC2 + ", " & rsFmtFld("FldName").Value
                End If
                
                '�E  case 'CCxx' �B<���W��>����:
                If UCase(Left(rsFmtFld("FldID").Value, 2)) = "CC" Then
                    'SQLC3 = SQLC3 + ", <���W��>"
                    strSQLC3 = strSQLC3 + ", " & rsFmtFld("FldName").Value
                End If
                
                '�E  case 'D'�}�Y, �B<���W��>����:
                If UCase(Left(rsFmtFld("FldID").Value, 1)) = "D" Then
                    'SQLD = SQLD + ", <���W��>"
                    strSQLD = strSQLD & ", " & rsFmtFld("FldName").Value
                End If
                
                '�E  case 'E'�}�Y, �B<���W��>����:
                If UCase(Left(rsFmtFld("FldID").Value, 1)) = "E" Then
                    ' SQLE = SQLE + ", <���W��>"
                    strSQLE = strSQLE & ", " & rsFmtFld("FldName").Value
                End If
                
                '�E  case 'F'�}�Y, �B<���W��>����
                ' �B��<���W��>���s�b��aryF��
                If UCase(Left(rsFmtFld("FldID").Value, 1)) = "F" Then
                    If InStr(strF, Chr(9) & rsFmtFld("FldName").Value) = 0 Then
                        ReDim Preserve aryF(0 To UBound(aryF) + 1) As String
                        ' �N<���W��>�[�JaryF��
                        aryF(UBound(aryF)) = rsFmtFld("FldName").Value
                        strF = strF & Chr(9) & rsFmtFld("FldName").Value
                    End If
                End If
                
                '�E  case 'I'�}�Y, �B<���W��>����:
                If UCase(Left(rsFmtFld("FldID").Value, 1)) = "I" Then
                    'SQLI = SQLI + ", <���W��>"
                    strSQLI = strSQLI & ", " & rsFmtFld("FldName").Value
                End If
                
                '�E  case 'J'�}�Y, �B<���W��>����:
                If UCase(Left(rsFmtFld("FldID").Value, 1)) = "J" Then
                    'strSQLJ = strSQLJ + ", <���W��>"
                    If rsFmtFld("FldID").Value <> "J01" Then
                        strSQLJ = strSQLJ & ", " & rsFmtFld("FldName").Value
                    End If
                End If
                '�E  case 'K'�}�Y, �B<���W��>����:
                If UCase(Left(rsFmtFld("FldID").Value, 1)) = "K" Then
                    'strSQLK = strSQLK + ", <���W��>"
                    strSQLK = strSQLK & ", " & rsFmtFld("FldName").Value
                End If
                
                '�E  case 'L'�}�Y, �B<���W��>����:
                If UCase(Left(rsFmtFld("FldID").Value, 2)) >= "L0" And _
                    UCase(Left(rsFmtFld("FldID").Value, 2)) <= "L9" Then
                    'strSQLL = strSQLL + ", <���W��>"
                    strSQLL = strSQLL & ", " & rsFmtFld("FldName").Value
                End If
                
                '�E  case 'LA'�}�Y, �B<���W��>����:
                If UCase(Left(rsFmtFld("FldID").Value, 2)) = "LA" Then
                    'strSQLL1 = strSQLL1 + ", <���W��>"
                    strSQLL1 = strSQLL1 & ", " & rsFmtFld("FldName").Value
                End If
                
                '�E  case 'LB'�}�Y, �B<���W��>����:
                If UCase(Left(rsFmtFld("FldID").Value, 2)) = "LB" Then
                    'strSQLL2 = strSQLL2 + ", <���W��>"
                    strSQLL2 = strSQLL2 & ", " & rsFmtFld("FldName").Value
                End If
            
                '�E  case 'M'�}�Y, �B<���W��>����:
                If UCase(Left(rsFmtFld("FldID").Value, 1)) = "M" Then
                    'strSQLK = strSQLK + ", <���W��>"
                    strSQLM = strSQLM & ", " & rsFmtFld("FldName").Value
                End If
                
                '�E  case 'NA'�}�Y, �B<���W��>����:
                If UCase(Left(rsFmtFld("FldID").Value, 2)) = "NA" Then
                    'strSQLK = strSQLK + ", <���W��>"
                    strSQLN1 = strSQLN1 & ", " & rsFmtFld("FldName").Value
                End If
                
                '�E  case 'NB'�}�Y, �B<���W��>����:
                If UCase(Left(rsFmtFld("FldID").Value, 2)) = "NB" Then
                    'strSQLK = strSQLK + ", <���W��>"
                    strSQLN2 = strSQLN2 & ", " & rsFmtFld("FldName").Value
                End If
                
                '�E  case 'OA'�}�Y, �B<���W��>����:   �����Ӧ��O���إ��`�]��
                If UCase(Left(rsFmtFld("FldID").Value, 2)) = "OA" Then
                    'strSQLOA = strSQLOA + ", <���W��>"
                    If Val(Right(rsFmtFld("FldID").Value, 2)) = 1 Then
                        strSQLOA = strSQLOA & ", " & rsFmtFld("FldName").Value
                    End If
                End If
                '�E  case 'OB'�}�Y, �B<���W��>����:   �����ӫȤ᥿�`�]��
                If UCase(Left(rsFmtFld("FldID").Value, 2)) = "OB" Then
                    'strSQLOB = strSQLOB + ", <���W��>"
                    If Val(Right(rsFmtFld("FldID").Value, 2)) = 1 Then
                        strSQLOB = strSQLOB & ", " & rsFmtFld("FldName").Value
                    End If
                End If
                '�E  case 'OC'�}�Y, �B<���W��>����:   �����ӫȤ᥿�`�]��
                If UCase(Left(rsFmtFld("FldID").Value, 2)) = "OC" Then
                    'strSQLOC = strSQLOC + ", <���W��>"
                    If Val(Right(rsFmtFld("FldID").Value, 2)) = 1 Then  '�Ĥ@�����N�i�H�F
                        strSQLOC = strSQLOC & ", " & rsFmtFld("FldName").Value
                    End If
                End If
                
            End If
            rsFmtFld.MoveNext
        Loop
        
        '3. ��z�USQL���O:
        '�E  SQLB = SQLB + " from SO001 where CustId="
        '���q�O
        If InStr(1, UCase(strSQLB), "COMPCODE") = 0 Then strSQLB = strSQLB & ",CompCode"
        '�Ȥ�s��
        If InStr(1, UCase(strSQLB), "CUSTID") = 0 Then strSQLB = strSQLB & ",CustId"
        '�[�˾��a�}
        If InStr(1, UCase(strSQLB), "INSTADDRNO") = 0 Then strSQLB = strSQLB & ",InstAddrNo"
        '�[���O�a�}
        If InStr(1, UCase(strSQLB), "CHARGEADDRNO") = 0 Then strSQLB = strSQLB & ",ChargeAddrNo"
        '�[�l�H�a�}
        If InStr(1, UCase(strSQLB), "MAILADDRNO") = 0 Then strSQLB = strSQLB & ",MailAddrNo"
        
        strSQLB = strSQLB & " FROM " & GetOwner & "SO001 WHERE CustID = "
        
        '�E  �YSQLC1����: ���O�a�}
        If strSQLC1 <> "" Then
            ' �h�����䤧 ',' �ASQLC1 = "select "+SQLC1+" from SO014 where AddrNo="
            If Left(strSQLC1, 1) = "," Then
                strSQLC1 = Mid(strSQLC1, 2)
            End If
            strSQLC1 = "SELECT " & strSQLC1 & " FROM " & GetOwner & "SO014 WHERE AddrNo = "
        End If
        
        '�E  �YSQLC2����: �˾��a�}
        If strSQLC2 <> "" Then
            If Left(strSQLC2, 1) = "," Then
                strSQLC2 = Mid(strSQLC2, 2)
            End If
            strSQLC2 = "SELECT " & strSQLC2 & " FROM " & GetOwner & "SO014 WHERE AddrNo = "
        End If
        
        '�E  �YSQLC3����: �l�H�a�}
        If strSQLC3 <> "" Then
            If Left(strSQLC3, 1) = "," Then
                strSQLC3 = Mid(strSQLC3, 2)
            End If
            strSQLC3 = "SELECT " & strSQLC3 & " FROM " & GetOwner & "SO014 WHERE AddrNo = "
        End If
        
        '�E  �YSQLD����:
        If strSQLD <> "" Then
            ' �h�����䤧 ','
            If Left(strSQLD, 1) = "," Then
                strSQLD = Mid(strSQLD, 2)
            End If
            ' SQLD = "select "+SQLD+" from SO002 where CustId="
            strSQLD = "SELECT " & strSQLD & " FROM " & GetOwner & "SO002 WHERE CustId = "
        End If
        
        '�E  �YSQLE����:
        If strSQLE <> "" Then
            ' �h�����䤧 ','
            If Left(strSQLE, 1) = "," Then
                strSQLE = Mid(strSQLE, 2)
            End If
            ' SQLE = "select "+SQLE+" from SO017 where MduId="
            strSQLE = "SELECT " & strSQLE & " FROM " & GetOwner & "SO017 WHERE MduId = "
        End If
        
        '�E  �YaryF������:
        If strF <> "" Then
            '1. loop aryF�C�@����
            For lngX = 1 To UBound(aryF)
                ' SQLF = SQLF + ", <���W��>"
                strSQLF = strSQLF & "," & aryF(lngX)
            Next
            
            '2. �h�����䤧','
            If Left(strSQLF, 1) = "," Then
                strSQLF = Mid(strSQLF, 2)
            End If
            
            ' �ASQLF = "select "+SQLF+" from SO034 where PrtSNo="
            'strSQLF = "SELECT " & strSQLF & " FROM SO034 WHERE PrtSNo = "
        End If
            
        strFieldName = Mid(strFieldName, 2)
        '�E  �YSQLI����: (91/08/07 �ҥ�)
        If strSQLI <> "" Then
            ' �h�����䤧 ','
            If Left(strSQLI, 1) = "," Then
                strSQLI = Mid(strSQLI, 2)
            End If
            ' SQLI = "select "+SQLI+" from ????? where CustId="
            strSQLI = "SELECT " & strSQLI & " FROM " & GetOwner & "SO107 WHERE "
        End If
        
        '�E  �YSQLJ����:
        If strSQLJ <> "" Then
            ' �h�����䤧 ','
            If Left(strSQLJ, 1) = "," Then
                strSQLJ = Mid(strSQLJ, 2)
            End If
            ' strSQLJ = "select "+strSQLJ+" from SO304 where MduId="
            strSQLJ = "SELECT " & strSQLJ & " FROM " & GetOwner & "SO304 WHERE "
        End If
        
        '�E  �YSQLK����:
        If strSQLK <> "" Then
            ' �h�����䤧 ','
            If Left(strSQLK, 1) = "," Then
                strSQLK = Mid(strSQLK, 2)
            End If
            ' strSQLK = "select "+strSQLK+" from SO043 where MduId="
            strSQLK = "SELECT " & strSQLK & " FROM " & GetOwner & "SO043 WHERE "
        End If
        
        '�E  �YSQLL����:
        If strSQLL <> "" Then
            '�[���O�a�}
            If InStr(1, UCase(strSQLL), "CHARGEADDRNO") = 0 Then strSQLL = strSQLL & ",ChargeAddrNo"
            '�[�l�H�a�}
            If InStr(1, UCase(strSQLL), "MAILADDRNO") = 0 Then strSQLL = strSQLL & ",MailAddrNo"
            ' �h�����䤧 ','
            If Left(strSQLL, 1) = "," Then
                strSQLL = Mid(strSQLL, 2)
            End If
            ' strSQLL = "select "+strSQLL+" from SO002A where MduId="
            strSQLL = "SELECT " & strSQLL & " FROM " & GetOwner & "SO002A WHERE "
        End If
        
        '�E  �YSQLL1����: ���O�a�}
        If strSQLL1 <> "" Then
            ' �h�����䤧 ',' �ASQLL1 = "select "+strSQLL1+" from SO014 where AddrNo="
            If Left(strSQLL1, 1) = "," Then
                strSQLL1 = Mid(strSQLL1, 2)
            End If
            strSQLL1 = "SELECT " & strSQLL1 & " FROM " & GetOwner & "SO014 WHERE AddrNo = "
        End If
        
        '�E  �YSQLL2����: �l�H�a�}
        If strSQLL2 <> "" Then
            ' �h�����䤧 ',' �ASQLL2 = "select "+strSQLL2+" from SO014 where AddrNo="
            If Left(strSQLL2, 1) = "," Then
                strSQLL2 = Mid(strSQLL2, 2)
            End If
            strSQLL2 = "SELECT " & strSQLL2 & " FROM " & GetOwner & "SO014 WHERE AddrNo = "
        End If
        '�E  �YSQLM����: SO004
        If strSQLM <> "" Then
            If Left(strSQLM, 1) = "," Then
                strSQLM = Mid(strSQLM, 2)
            End If
            strSQLM = "SELECT " & strSQLM & " FROM " & GetOwner & "SO004 WHERE TranBillNO = "
        End If
        
        '�E  �YSQLN1����: SO152
        If strSQLN1 <> "" Then
            If Left(strSQLN1, 1) = "," Then
                strSQLN1 = Mid(strSQLN1, 2)
            End If
            strSQLN1 = "SELECT " & strSQLN1 & " FROM " & GetOwner & "SO152 Where EmcCustID = "
        End If
        
        '�E  �YSQLN2����: SO152
        If strSQLN2 <> "" Then
            If Left(strSQLN2, 1) = "," Then
                strSQLN2 = Mid(strSQLN2, 2)
            End If
            strSQLN2 = "SELECT " & strSQLN2 & " FROM " & GetOwner & "SO152 Where EmcCustID = "
        End If
        
        '�E  �YSQLOA����: SO004 �Ӧ��O���ت��]��
        If strSQLOA <> "" Then
            If Left(strSQLOA, 1) = "," Then
                strSQLOA = Mid(strSQLOA, 2)
            End If
            strSQLOA = "SELECT " & strSQLOA & " FROM " & GetOwner & "SO004 WHERE "
        End If
        
        '�E  �YSQLOB����: SO004 �ӫȤ᪺�]��
        If strSQLOB <> "" Then
            If Left(strSQLOB, 1) = "," Then
                strSQLOB = Mid(strSQLOB, 2)
            End If
            strSQLOB = "SELECT " & strSQLOB & " FROM " & GetOwner & "SO004,(Select RefNo,CodeNo From " & GetOwner & "CD022 ) CD022 WHERE SO004.FaciCode = CD022.CodeNo And "
        End If
        
        '�E  �YSQLP����: SO004 �ӫȤ᪺��LSO002
        If strSQLOC <> "" Then
            If Left(strSQLOC, 1) = "," Then
                strSQLOC = Mid(strSQLOC, 2)
            End If
            strSQLOC = "SELECT " & strSQLOC & " FROM " & GetOwner & "SO002 Where CustId = "
        End If
        
        '�E  �YSQLOD����: SO004 �ӫȤ᪺��LSO138
        If strSQLOD <> "" Then
            If Left(strSQLOD, 1) = "," Then
                strSQLOD = Mid(strSQLOD, 2)
            End If
            strSQLOD = "SELECT " & strSQLOD & " FROM " & GetOwner & "SO138 Where InvSeqNo = "
        End If
        
        If Not CreatePrtSNoTmp(cnTmpMDB, True, rsFmtFld) Then Exit Function
        ChargeBillField = True
    Exit Function
ChkErr:
    ErrSub FormName, "ChargeBillField"
End Function

Private Function GetChargeBillPara(strCompCode As String, strServiceType As String, _
        intBillTab As Integer, ByRef strPrinter As String, ByRef strFmtID As String) As Boolean
    On Error GoTo ChkErr
    Dim rs As New ADODB.Recordset
        If Not GetSystemPara(rs, strCompCode, strServiceType, "SO045", "BillTab" & intBillTab & " ,BillPrinter") Then Exit Function
        '���O��C���
        If rs.RecordCount = 0 Then Exit Function
        If strPrinter = "" Then strPrinter = rs("BillPrinter") & ""
        '�U�����O��榡�N��
        If strFmtID = "" Then strFmtID = rs("BillTab" & intBillTab) & ""
        Close3Recordset rs
        GetChargeBillPara = True
    Exit Function
ChkErr:
    ErrSub FormName, "GetChargeBillPara"
End Function

Private Function CreateAddrModifyTable(cnTmpMDB As ADODB.Connection) As Boolean
    On Error GoTo ChkErr
    Dim strSQL As String
    Dim rs As New ADODB.Recordset
        strSQL = "Select A.BillNo,A.Item,A.PrtSNo,A.CustId,B.CustName, " & _
            "A.CitemName,B.InstAddrNo OldAddrNo, B.InstAddress OldAddress,B.InstAddrNo NewAddrNo,B.InstAddress NewAddress From " & GetOwner & "So033 A," & GetOwner & "So001 B Where A.BillNo = '' And A.Item = -1 And A.CustId = B.CustId "
        If Not GetRS(rs, strSQL) Then Exit Function
        If Not CreateMDBTable(rs, "BillAddrModify", cnTmpMDB) Then Exit Function
        CreateAddrModifyTable = True
    Exit Function
ChkErr:
    ErrSub FormName, "CreateAddrModifyTable"
End Function

' A: �e���ǳưʧ@
Public Function PrintChargeBill(crpBill As Object, strForm As String, objForm As Object, _
        intBillTab As Integer, Optional strCompCode As String, Optional strServiceType As String, _
        Optional strPrinter As String = "", Optional strFmtID As String = "", _
        Optional blnEmailList As Boolean = False, _
        Optional strTXTFileName As String, Optional strFilePath As String, _
        Optional ByRef lngAffecteds As Long, Optional blnPrintTEXT As Boolean, _
        Optional ByVal strFormula As String, Optional blnCombineBill As Boolean, _
        Optional ByVal strRptMDBPath As String, Optional ByVal strTmpMDBPath As String, _
        Optional ByVal blnRemoteMode As Boolean) As Boolean
    On Error GoTo ChkErr
    Dim obj As Object
    Set obj = CreateObject("csPrintChargeBill.csPrintCharge")
        If Alfa Then
            If strRptMDBPath = "" Then strRptMDBPath = "c:\CSMISV40\BIN\"
            If strTmpMDBPath = "" Then strTmpMDBPath = "c:\TmpMDB\"
        Else
            strRptMDBPath = App.Path & "\"
            strTmpMDBPath = ReadGICMIS1("TmpMDBPath") & "\"
        End If
        With obj
            .ugarygi = PutGlobal
            .uErrPath = ReadGICMIS1("ErrLogPath")
            .uRptPath = strRptMDBPath
            .uMdbPath = strTmpMDBPath
            .uExportPath = ReadGICMIS1("ExportPath")
            Set .ugcnGi = gcnGi
            .uAlfa = Alfa
            .uCompCode = strCompCode
            If Not .PrintChargeBill(crpBill, strForm, objForm, intBillTab, strCompCode, strServiceType, strPrinter, strFmtID, _
                blnEmailList, strTXTFileName, strFilePath, lngAffecteds, blnPrintTEXT, strFormula, _
                blnCombineBill, strRptMDBPath, strTmpMDBPath, blnRemoteMode) Then
                MsgBox .uErrMsg, vbCritical, gimsgPrompt
                Exit Function
            End If
        End With
        Set obj = Nothing
        PrintChargeBill = True
    Exit Function
    Dim strSQLB As String, strSQLC1 As String, strSQLC2 As String, strSQLC3 As String
    Dim strSQLD As String, strSQLE As String, strSQLF As String
    Dim strSQLI As String, strSQLJ As String, strSQLK As String
    Dim strSQLL As String, strSQLL1 As String, strSQLL2 As String
    Dim strSQLM As String, strSQLN1 As String, strSQLN2 As String
    Dim strSQLOA As String, strSQLOB As String, strSQLOC As String
    Dim strSQL33 As String, strSQLOD As String
    Dim strFieldName As String
    
    Dim cnRptMDB As New ADODB.Connection, cnTmpMDB As New ADODB.Connection
    
    Dim blnToWindow As Boolean, blnJoin2 As Boolean
    Dim blnClose As Boolean
    
    Dim rsData As New ADODB.Recordset, rsPrtSNOTmp As New ADODB.Recordset
    
    Dim lngBatchNo As Long
        strForm = UCase(strForm)
        If Alfa Then
            If strRptMDBPath = "" Then strRptMDBPath = "c:\CSMISV40\BIN\"
            If strTmpMDBPath = "" Then strTmpMDBPath = "c:\TmpMDB\"
        Else
            strRptMDBPath = App.Path & "\"
            strTmpMDBPath = ReadGICMIS1("TmpMDBPath") & "\"
        End If
        cnRptMDB.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strRptMDBPath & "RptFld.mdb;Persist Security Info=False"
        cnTmpMDB.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strTmpMDBPath & "TMP1111.mdb;Persist Security Info=False"
        If Not GetChargeBillPara(strCompCode, strServiceType, intBillTab, strPrinter, strFmtID) Then Exit Function
        '�����
        If Not ChargeBillChoose(strForm, objForm, rsData, lngBatchNo, strSQL33, blnClose) Then Exit Function
        '�����
        If Not ChargeBillField(cnRptMDB, strFmtID, strSQLB, strSQLC1, strSQLC2, strSQLC3, strSQLD, strSQLE, strSQLF, strSQLI, strSQLJ, strSQLK, strSQLL, strSQLL1, strSQLL2, strSQLM, strSQLN1, strSQLN2, strSQLOA, strSQLOB, strFieldName, blnJoin2, cnTmpMDB, strSQLOC, strSQLOD) Then Exit Function
        '�إߦa�}��Ƨ�ʤ�Log File
        If Not CreateAddrModifyTable(cnTmpMDB) Then Exit Function
        '�R��TEMP���
        'If Not CreatePrtSNoTmp(cnTmpMDB, False) Then Exit Function
        '�}�lInsert Data
        'If blnJoin2 Then
        '    If Not GetRS(rsPrtSNOTmp, "Select " & strFieldName & " From PrtSNoTmp,PrtSNoTmp2 Where 1=0 ", cnTmpMDB, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
        'Else
            If Not GetRS(rsPrtSNOTmp, "Select " & strFieldName & " From PrtSNoTmp Where 1=0 ", cnTmpMDB, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
        'End If
        'Set rsPrtSNOTmp.ActiveConnection = Nothing
        On Error GoTo ErrTrans
        gcnGi.BeginTrans
        'cnTmpMDB.BeginTrans
        If Not GotoFlowProc(cnTmpMDB, cnRptMDB, strForm, objForm, strFmtID, rsData, rsPrtSNOTmp, strSQLB, strSQLC1, strSQLC2, strSQLC3, strSQLD, strSQLE, _
                    strSQLF, strSQLI, strSQLJ, strSQLK, strSQLL, strSQLL1, strSQLL2, strSQLM, strSQLN1, strSQLN2, strSQLOA, strSQLOB, strSQL33, lngAffecteds, strCompCode, blnClose, strSQLOC, strSQLOD) Then
            gcnGi.RollbackTrans
         '   cnTmpMDB.RollbackTrans
            Exit Function
        End If
        gcnGi.Execute "Delete From " & GetOwner & "SO075 Where BatchNo = " & lngBatchNo
        'cnTmpMDB.CommitTrans
        gcnGi.CommitTrans
        '�إ�TmpMDB PK
        'If Not CreatePrtSNoTmp(cnTmpMDB, True) Then Exit Function
        On Error GoTo ChkErr
        '���PConnection
        If lngAffecteds > 0 Then
            If Not blnRemoteMode Then MsgBox "�Ȧs��ƳB�z�����A�@ [" & lngAffecteds & "] ���A�{�b�}�l�C�L...�C", vbInformation, "�T��"
            If blnEmailList Then
                strTXTFileName = GetRsValue("SELECT AttachFileName FROM FmtList WHERE FmtID = '" & objForm.cboFmtID.Text & "'", cnRptMDB) & ""
                If Not EmailList(rsPrtSNOTmp, strTXTFileName, objForm.cboType.Text, strFilePath, objForm.chkIncludeFile) Then Exit Function
            Else
                If blnPrintTEXT Then
                    strTXTFileName = GetRsValue("SELECT AttachFileName FROM FmtList WHERE FmtID = '" & objForm.cboFmtID.Text & "'", cnRptMDB) & ""
                    If Not PrintBillToTEXT(rsPrtSNOTmp, strTXTFileName, strFilePath) Then Exit Function
                Else
                    Dim strRptFileName As String
                    Dim lngModiCount As Long
                    '�ˬd���O��a�}���L����
                    lngModiCount = Val(GetRsValue("Select Count(*) From BillAddrModify", cnTmpMDB) & "")
                    
                    strRptFileName = GetRsValue("Select FmtFileName From FmtList Where FmtId ='" & strFmtID & "'", cnRptMDB) & ""
                    '�D�ϥλ��ݥX�� 92/06/23
                    If strForm = "FRMSO1100BMDI" Then
                        If Not ChargeBillOutPrint(crpBill, strTmpMDBPath, strRptFileName, lngModiCount, strPrinter, , , , True, , , cnTmpMDB) Then Exit Function
                    Else
                        blnToWindow = GetSystemParaItem("PreView", gCompCode, "", "SO041")
                        If strForm = "FRMSO3261A" Then
                            If Not ChargeBillOutPrint(crpBill, strTmpMDBPath, strRptFileName, lngModiCount, strPrinter, objForm.txtNote1, objForm.txtNote2, objForm.gymClctYM.GetValue(True) & "/" & objForm.txtDueDate, blnToWindow, , , cnTmpMDB) Then Exit Function
                        ElseIf strForm = "FRMSO3262A" Then
                            If Not ChargeBillOutPrint(crpBill, strTmpMDBPath, strRptFileName, lngModiCount, strPrinter, objForm.txtNote1, objForm.txtNote2, objForm.gdaDueDate.GetValue(True), blnToWindow, , , cnTmpMDB) Then Exit Function
                        ElseIf strForm = "FRMSO3273A" Then
                            If Not ChargeBillOutPrint(crpBill, strTmpMDBPath, strRptFileName, lngModiCount, strPrinter, "", "", "", blnToWindow, strFormula, , cnTmpMDB) Then Exit Function
                        ElseIf strForm = "FRMSO4000A" Then
                            If Not ChargeBillOutPrint(crpBill, strTmpMDBPath, strRptFileName, lngModiCount, strPrinter, , , , blnToWindow, , , cnTmpMDB) Then Exit Function
                        ElseIf strForm = "FRMSO32A2A" Then
                            If Not ChargeBillOutPrint(crpBill, strTmpMDBPath, strRptFileName, lngModiCount, strPrinter, objForm.txtNote1, , , blnToWindow, , , cnTmpMDB) Then Exit Function
                        End If
                    End If
                End If
            End If
        Else
                MsgBox "�B�z�s����ơA���ݭn�C�L�C", vbInformation, "�T��"
        End If
        If strForm = "FRMSO3261A" Or strForm = "FRMSO3262A" Then
            '�E  �Y�Ƶ��榳��, �h�N���Ȫ��Ƶ��椺�e�s�^SO039������줤
            With objForm
                If .txtNote1.Text <> "" Or .txtNote2 <> "" Then
                    lngAffecteds = Val(GetRsValue("Select Count(*) From " & GetOwner & "So039 Where RowNum =1") & "")
                    If lngAffecteds > 0 Then
                        If ExecuteSQL("UPDATE " & GetOwner & "SO039 SET Memo1 = " & GetNullString(.txtNote1.Text) & ", Memo2 = " & GetNullString(.txtNote2.Text), gcnGi, , False) <> giOK Then Exit Function
                    Else
                        If ExecuteSQL("Insert Into " & GetOwner & "SO039 (Memo1,Memo2,CompCode) Values (" & GetNullString(.txtNote1.Text) & ", " & GetNullString(.txtNote2.Text) & ",1)", gcnGi, , False) <> giOK Then Exit Function
                    End If
                End If
            End With
        End If
        Close3Recordset rsData
        Close3Recordset rsPrtSNOTmp
        CloseConnection cnRptMDB
        CloseConnection cnTmpMDB
        PrintChargeBill = True
    Exit Function
ErrTrans:
    gcnGi.RollbackTrans
    cnTmpMDB.RollbackTrans
ChkErr:
    ErrSub "Utility3", "PrintChargeBill"
End Function

Public Function CreatePrtSNoTmp(cnTmpMDB As ADODB.Connection, blnCreateTable As Boolean, _
    rsField As ADODB.Recordset, Optional StrTableName As String) As Boolean
    On Error GoTo ChkErr
    Dim strExecute As String
    Dim intLength As Integer
        If StrTableName = "" Then StrTableName = "PrtSNoTmp"
        If blnCreateTable Then
            With rsField
                On Error Resume Next
                cnTmpMDB.Execute "Drop Table " & StrTableName
                On Error GoTo ChkErr
                If .RecordCount > 0 Then .MoveFirst
                Do While Not .EOF
                    intLength = Val(.Fields("Length") & "")
                    Select Case UCase(.Fields("FldType") & "")
                        Case "C"
                            If intLength = 0 Then intLength = 255
                            If intLength <= 255 Then
                                strExecute = strExecute & "," & .Fields("FldId") & " TEXT(255)"
                            Else
                                strExecute = strExecute & "," & .Fields("FldId") & " Memo"
                            End If
                        Case "D"
                            strExecute = strExecute & "," & .Fields("FldId") & " Date"
                        Case "N"
                            strExecute = strExecute & "," & .Fields("FldId") & " Double"
                    End Select
                    .MoveNext
                Loop
            End With
            strExecute = Mid(strExecute, 2)
            cnTmpMDB.Execute "Create Table " & StrTableName & " (" & strExecute & ")"
            
            'cnTmpMDB.Execute "Alter Table PrtSNoTmp Add CONSTRAINT PrimaryKey PRIMARY KEY (RowId)"
            'cnTmpMDB.Execute "Alter Table PrtSNoTmp2 Add CONSTRAINT PrimaryKey PRIMARY KEY(RowId)"
        Else
            On Error Resume Next
            cnTmpMDB.Execute "Drop Table " & StrTableName
            'If Not DeleteTmpTable(cnTmpMDB, "PrtSNoTmp") Then Exit Function
            'If Not DeleteTmpTable(cnTmpMDB, "PrtSNoTmp2") Then Exit Function
        End If
        CreatePrtSNoTmp = True
    Exit Function
ChkErr:
    ErrSub FormName, "CreatePrtSNoTmp"
End Function

Public Function ChargeBillOutPrint(crpBill As Object, _
        ByVal strTmpMDBPath As String, ByVal strRptFileName As String, _
        ByVal lngModiCount As Long, _
        Optional ByVal strPrinterName As String, _
        Optional ByVal strNote1 As String, Optional ByVal strNote2 As String, _
        Optional ByVal strDueDate As String, Optional ByVal blnToWindow As Boolean = False, _
        Optional ByVal strFormula As String = "", Optional ByVal strWhere As String, _
        Optional ByVal cnTmpMDB As ADODB.Connection) As Boolean
    On Error GoTo ChkErr
    Dim clsXPrinter As Object
    Dim strFormula2 As String
    Dim blnToPrinter As Boolean
    
        Screen.MousePointer = vbHourglass
        Set clsXPrinter = CreateObject("SetDef.Xprinter")
        Call clsXPrinter.ChangePrinter(strPrinterName)
        strRptFileName = strRptFileName & IIf(UCase(Right(strRptFileName, 4)) <> ".RPT", ".RPT", "")
        If Right(ReadFromINI2("RptPath"), 1) = "\" Then
            crpBill.ReportFileName = ReadFromINI2("RptPath") & strRptFileName
        Else
            crpBill.ReportFileName = ReadFromINI2("RptPath") & "\" & strRptFileName
        End If
        If Right(strTmpMDBPath, 1) = "\" Then
            crpBill.DataFiles(0) = strTmpMDBPath & "Tmp1111.MDB"
        Else
            crpBill.DataFiles(0) = strTmpMDBPath & "\Tmp1111.MDB"
        End If
        strFormula2 = "PrintEn = '" & garyGi(0) & "';PrintName = '" & garyGi(1) & "'"
        If strNote1 <> "" Then strFormula2 = strFormula2 & ";Memo1 = '" & Replace(strNote1, vbCrLf, "' + chr(13) + '") & "'"
        If strNote2 <> "" Then strFormula2 = strFormula2 & ";Memo2 = '" & Replace(strNote2, vbCrLf, "' + chr(13) + '") & "'"
        If strDueDate <> "" Then strFormula2 = strFormula2 & ";DueDate = '" & strDueDate & "'"
        
        If Alfa Or blnToWindow Then
            blnToPrinter = False
        Else
            blnToPrinter = True
        End If
        
        If strFormula <> "" Then strFormula = strFormula & ";"
        strFormula = strFormula & strFormula2
        InsertErrLog cnTmpMDB, 1
        PrintRpt2 strPrinterName, strRptFileName, , "���O��C�L", "Select * Form PrtSNoTmp" & strWhere, , , True, "TMP1111.mdb", , strFormula, , , , , , , , blnToPrinter, , , , True, False
        InsertErrLog cnTmpMDB, 2
        '.Action = 1
        On Error Resume Next
        Call clsXPrinter.DefaultPrinter
        InsertErrLog cnTmpMDB, 3
        '�C�L�a�}����ʪ����
        If lngModiCount > 0 Then
            InsertErrLog cnTmpMDB, 31
            MsgBox "���O��a�}��Ʀ�����,�N�C�L���ʪ�!!", vbInformation, gimsgPrompt
            PrintRpt2 GetPrinterName(5), "BillAddrModify.Rpt", , "���O��a�}���ʪ� [Utility3]", "Select * From BillAddrModify", , , True, "Tmp1111.MDB"
            InsertErrLog cnTmpMDB, 32
        End If
        InsertErrLog cnTmpMDB, 4
        
        Set clsXPrinter = Nothing
        Screen.MousePointer = vbDefault
        ChargeBillOutPrint = True
    Exit Function
ChkErr:
    ErrSub FormName, "ChargeBillOutPrint"
End Function

Public Function InsertErrLog(cnTmpMDB As ADODB.Connection, intStep As Integer) As Boolean
    On Error Resume Next
        If intStep = 1 Then
            cnTmpMDB.Execute "Drop Table TMP_ErrorLogStep"
            cnTmpMDB.Execute "Create Table TMP_ErrorLogStep (WhatNow Date,WhatName TEXT(20),WhatStep Single)"
        End If
        cnTmpMDB.Execute "Insert Into TMP_ErrorLogStep (WhatNow,WhatName,WhatStep) Values (" & _
            GetNullString(RightNow, giDateV, giAccessDb, True) & ",'" & garyGi(1) & "'," & intStep & ")"
End Function

'Private Function ChargeBillOutPrint(crpBill As Object, _
'        strTmpMDBPath As String, strRptFileName As String, _
'        lngModiCount As Long, _
'        Optional strPrinterName As String, _
'        Optional strNote1 As String, Optional strNote2 As String, _
'        Optional strDueDate As String, Optional blnToWindow As Boolean = False) As Boolean
'    On Error GoTo ChkErr
'    'Dim cls As SetDef.Xprinter
'    Dim clsXPrinter As Object
'    Dim strFormula2 As String
'    Dim blnToPrinter As Boolean
'
'        Screen.MousePointer = vbHourglass
'        Set clsXPrinter = CreateObject("SetDef.Xprinter")
'        With crpBill
'            Call clsXPrinter.ChangePrinter(strPrinterName)
'            strRptFileName = strRptFileName & IIf(UCase(Right(strRptFileName, 4)) <> ".RPT", ".RPT", "")
'            If Right(ReadFromINI2("RptPath"), 1) = "\" Then
'                crpBill.ReportFileName = ReadFromINI2("RptPath") & strRptFileName
'            Else
'                crpBill.ReportFileName = ReadFromINI2("RptPath") & "\" & strRptFileName
'            End If
'            '.PrinterName = cboPrinterName.Text
'            'Call SetPrinter(strPrinterName, crpBill)
'            If Right(strTmpMDBPath, 1) = "\" Then
'                crpBill.DataFiles(0) = strTmpMDBPath & "Tmp1111.MDB"
'            Else
'                crpBill.DataFiles(0) = strTmpMDBPath & "\Tmp1111.MDB"
'            End If
'            .Formulas(0) = "PrintEn = '" & garyGi(0) & "'"
'            '.WindowTitle = "���O��C�L"
'            Dim cc As CrystalReport
'            .WindowControlBox = False
'            If strNote1 <> "" Then .Formulas(1) = "Memo1 = '" & Replace(strNote1, vbCrLf, "' + chr(13) + '") & "'"
'            If strNote2 <> "" Then .Formulas(2) = "Memo2 = '" & Replace(strNote2, vbCrLf, "' + chr(13) + '") & "'"
'            If strDueDate <> "" Then .Formulas(3) = "DueDate = '" & strDueDate & "'"
'            strFormula2 = .Formulas(0) & ";" & .Formulas(1) & ";" & .Formulas(2) & ";" & .Formulas(3)
'
'            If Alfa Or blnToWindow Then
'                .Destination = crptToWindow
'                blnToPrinter = False
'            Else
'                .Destination = crptToPrinter
'                blnToPrinter = True
'            End If
'            .WindowControls = True
'            .WindowParentHandle = frmPrintForm.hWnd
'            .WindowState = crptMaximized
'            .WindowAllowDrillDown = True
'            .WindowShowPrintSetupBtn = True
'            .ProgressDialog = False
'
'            PrintRpt strPrinterName, strRptFileName, , "���O��C�L", "Select * Form PrtSNoTmp", , , True, "TMP1111.mdb", , strFormula2, , , , , , , , blnToPrinter, , , , True, False
'            '.Action = 1
'            On Error Resume Next
'            Call clsXPrinter.DefaultPrinter
'            'frmPrintForm.Show vbModal
'            '�C�L�a�}����ʪ����
'            'If lngModiCount > 0 Then frmPrintForm.PrintBillAddrModify
'        End With
'        Set clsXPrinter = Nothing
'        Screen.MousePointer = vbDefault
'        ChargeBillOutPrint = True
'    Exit Function
'ChkErr:
'    ErrSub FormName, "ChargeBillOutPrint"
'End Function


Private Function GotoFlowProc(cnTmpMDB As ADODB.Connection, cnRptMDB As ADODB.Connection, _
        ByRef strForm As String, ByRef objForm As Object, ByRef strFmtID As String, _
        ByRef rs As ADODB.Recordset, ByRef rsPrtSNOTmp As ADODB.Recordset, _
        ByRef strSQLB As String, ByRef strSQLC1 As String, ByRef strSQLC2 As String, ByRef strSQLC3 As String, _
        ByRef strSQLD As String, ByRef strSQLE As String, ByRef strSQLF As String, _
        ByRef strSQLI As String, ByRef strSQLJ As String, ByRef strSQLK As String, _
        ByRef strSQLL As String, ByRef strSQLL1 As String, ByRef strSQLL2 As String, _
        ByRef strSQLM As String, ByRef strSQLN1 As String, ByRef strSQLN2 As String, _
        ByRef strSQLOA As String, ByRef strSQLOB As String, _
        ByRef strSQL33 As String, ByRef lngAffecteds As Long, _
        strCompCode As String, blnClose As Boolean, ByRef strSQLOC As String, ByRef strSQLOD As String) As Boolean
    Dim lngNowCustID As Long            ' �ثe�Ƚs
    Dim strNowPrtSNO As String          ' �ثe�Ǹ�
    Dim lngCustStatusCode As Long       ' ���Ȥ᪺�Ȥ᪬�A
    Dim lShouldPrint As Boolean         ' ���Ȥ�O�_�ݭn�L��
    Dim rsA As New ADODB.Recordset
    Dim lngCounter As Long
    Dim rsB As New ADODB.Recordset
    Dim rsC1 As New ADODB.Recordset
    Dim rsC2 As New ADODB.Recordset
    Dim rsC3 As New ADODB.Recordset
    Dim rsD As New ADODB.Recordset
    Dim rsE As New ADODB.Recordset
    Dim rsF As New ADODB.Recordset
    Dim rsI As New ADODB.Recordset
    Dim rsJ As New ADODB.Recordset
    Dim rsK As New ADODB.Recordset
    Dim rsL As New ADODB.Recordset
    Dim rsL1 As New ADODB.Recordset
    Dim rsL2 As New ADODB.Recordset
    Dim rsM As New ADODB.Recordset
    Dim rsN1 As New ADODB.Recordset
    Dim rsN2 As New ADODB.Recordset
    Dim rsOA As New ADODB.Recordset
    Dim rsOB As New ADODB.Recordset
    Dim rsOC As New ADODB.Recordset
    Dim rsOD As New ADODB.Recordset
    Dim strxSQLB As String, _
         strxSQLC1 As String, strxSQLC2 As String, strxSQLC3 As String, _
         strxSQLD As String, strxSQLE As String, strxSQLF As String, _
         strxSQLI As String, strxSQLJ As String, strxSQLK As String, _
         strxSQLL As String, strxSQLL1 As String, strxSQLL2 As String, _
         strxSQLM As String, strxSQLN1 As String, strxSQLN2 As String, _
         strxSQLOA As String, strxSQLOB As String, strxSQLOC As String, strxSQLOD As String
    Dim rsFmtFld As New ADODB.Recordset
    Dim rsFmtFldTmp(20) As New ADODB.Recordset
    Dim rsSO034 As New ADODB.Recordset
    Dim lngBillcount As Long
    Dim intLoop As Integer
    Dim rsTmp As New ADODB.Recordset
    Dim strServiceType As String
    Dim rsPara As New ADODB.Recordset
    Dim strWhere As String
    Dim StrTableName As String
    Dim strField As String
    Dim lngCompCode As Long
        On Error GoTo ER
        'Set rs.ActiveConnection = Nothing
        ' Loop �C�@�����
        On Error Resume Next
        If strForm <> "FRMSO1100BMDI" Then objForm.Picture1.Visible = True
        On Error GoTo ER
        
        'If rs.RecordCount = 0 Then GotoFlowProc = True: lngAffecteds = 0: Exit Function
        If Not GetSystemPara(rsPara, strCompCode, "", "SO043", "Para19,Para20") Then Exit Function
        '�ˬdSource Table
        If blnClose Then
            StrTableName = "SO034"
        Else
            StrTableName = "SO033"
        End If
        If rs.RecordCount Then rs.MoveFirst
        Do While Not rs.EOF
            On Error Resume Next
            If strForm <> "FRMSO1100BMDI" Then objForm.pgbBar.Value = rs.AbsolutePosition / rs.RecordCount * 100
            On Error GoTo ER
            lngNowCustID = rs("CustID").Value           ' �ثe�Ƚs = �ӵ��Ȥ�s��
            If strForm = "FRMSO3261A" Then
                strNowPrtSNO = (rs("PrtSNO").Value & "")    ' �ثe�Ǹ� = �ӵ��L��Ǹ�
            End If
            lShouldPrint = True
            
'            If strForm = "FRMSO3261A" Then
'                ' �p�G���n�C�L�ݩ��
'                lngCustStatusCode = GetValueFromRs2("SELECT CustStatusCode FROM SO002 WHERE CustId = " & lngNowCustID & " And ServiceType= '", gcnGi)
'                'If objForm.chkOption1.Value = 0 Then
'                    ' ���ӫȤ᪺�Ȥ᪬�A
'                    ' �D�ݩ��?
'                    ' 3 = ���, 6 = �����
'                '    If lngCustStatusCode <> 3 And lngCustStatusCode <> 6 Then
'                '        lShouldPrint = True
'                '    End If
'                'Else
'                '    lShouldPrint = True
'                'End If
'                '�E  �Y���Ȥ����O����: 2001/11/12�s�W
'                lShouldPrint = True
'                If objForm.gimCustStatusCode.GetQryStr <> "" Then
'                    'SQL=SQL+" and A.ClassCode"+<p_ClassSQL>
'                    If InStr(1, objForm.gimCustStatusCode.GetQueryCode, lngCustStatusCode) > 0 Then
'                        lShouldPrint = True
'                    Else
'                        lShouldPrint = False
'                    End If
'                End If
'            Else
'                lShouldPrint = True
'            End If
            
            ' ���Ȥ�ݭn�L��
            If lShouldPrint Then
                'Debug.Print rs("PrtSNo")
                Select Case UCase(strForm)
                    Case "FRMSO3261A"
                        '�E  �Y���Ȥ᪬�A����: 91/09/04
                        strServiceType = ""
                        If objForm.gimCustStatusCode.GetQryStr <> "" Then
                            'SQL=SQL+" and A.CitemCode"+<p_CitemSQL>
                            If Not GetRS(rsTmp, "Select ServiceType From " & GetOwner & "SO002 Where CustId = " & rs("CustId") & " And CustStatusCode " & objForm.gimCustStatusCode.GetQryStr) Then Exit Function
                            If Not rsTmp.EOF Then
                                strServiceType = " And Instr('" & rsTmp.GetString(, , "", ",", "") & "', SERVICETYPE) >0 " & ""
                            Else
                                strServiceType = " And BillNo = '' And Item=-1 "
                            End If
                        End If
                        'SO009 ����w����� 93/08/18
                        strWhere = ""
                        If objForm.gdaResvTime1.Text <> "" Then strWhere = " And ResvTime >=" & GetNullString(objForm.gdaResvTime1.GetValue(True), giDateV)
                        If objForm.gdaResvTime2.Text <> "" Then strWhere = strWhere & " And  ResvTime <" & GetNullString(objForm.gdaResvTime2.GetValue(True), giDateV) & "+1"
                        If strWhere <> "" Then
                            If Not GetRS(rsTmp, "Select ServiceType From " & GetOwner & "SO009 Where CustId = " & rs("CustId") & " And SignDate Is Null " & strWhere) Then Exit Function
                            If Not rsTmp.EOF Then
                                strServiceType = strServiceType & " And Instr('" & rsTmp.GetString(, , "", ",", "") & "', SERVICETYPE) >0 " & ""
                            Else
                                strServiceType = " And BillNo = '' And Item=-1 "
                            End If
                        End If
                        
                        If objForm.cboBillType.ListIndex = 0 Then
                            If Not GetRS(rsA, "SELECT * FROM " & GetOwner & StrTableName & " A WHERE PrtSNO = '" & strNowPrtSNO & "' " & strSQL33 & strServiceType & " ORDER BY CitemCode,ShouldDate DESC", gcnGi) Then Exit Function
                        ElseIf objForm.cboBillType.ListIndex = 1 Then
                            If Not GetRS(rsA, "SELECT * FROM " & GetOwner & StrTableName & " A WHERE CustId =  " & rs("CustId").Value & strSQL33 & strServiceType & " ORDER BY CitemCode,ShouldDate DESC", gcnGi) Then Exit Function
                        Else
                            If Not GetRS(rsA, "SELECT * FROM " & GetOwner & StrTableName & " A WHERE MediaBillNo =  '" & rs("PrtSNO").Value & "' " & strSQL33 & strServiceType & " ORDER BY CitemCode,ShouldDate DESC", gcnGi) Then Exit Function
                        End If
                        ' rsA="select ... from SO033 where pRTsnO='<�ثe�Ǹ�>' ORDER BY ShouldDate DESC"
                    Case "FRMSO3262A"
                        If Not GetRS(rsA, "SELECT * FROM " & GetOwner & StrTableName & " A WHERE CustId = " & rs("CustId").Value & strSQL33 & " ORDER BY CitemCode,ShouldDate DESC", gcnGi) Then Exit Function
                    Case "FRMSO1100BMDI"
                        If Not GetRS(rsA, "SELECT * FROM " & GetOwner & StrTableName & " A WHERE CustId = " & rs("CustId").Value & strSQL33 & " ORDER BY CitemCode,ShouldDate DESC", gcnGi) Then Exit Function
                    Case "FRMSO3273A"
                        Select Case Len(rs("BillNo") & "")
                            Case 11
                                strWhere = " And MediaBillNO = '" & rs("BillNo") & "'"
                            Case 12
                                strWhere = " And PrtSNo = '" & rs("BillNo") & "'"
                            Case Else
                                strWhere = " And BillNO = '" & rs("BillNo") & "'"
                        End Select
                        If Not GetRS(rsA, "SELECT * FROM " & GetOwner & StrTableName & " A WHERE CustId = " & rs("CustId").Value & " " & strWhere & " " & strSQL33 & " ORDER BY CitemCode,ShouldDate DESC", gcnGi) Then Exit Function
                    Case "FRMSO4000A"
                        If Not GetRS(rsA, "SELECT * FROM " & GetOwner & StrTableName & " A WHERE CustId = " & rs("CustId").Value & strSQL33 & " ORDER BY CitemCode,ShouldDate DESC", gcnGi) Then Exit Function
                    Case "FRMSO32A2A"
                        If strField = "" Then strField = ChargeField
                        If Not GetRS(rsA, "SELECT " & strField & " FROM " & GetOwner & "SO033 A WHERE BillNo = '" & rs("BillNo").Value & "'" & strSQL33 & " Union All " & _
                                                "SELECT " & strField & " FROM " & GetOwner & "SO034 A WHERE BillNo = '" & rs("BillNo").Value & "'" & strSQL33 & " ORDER BY CitemCode,ShouldDate DESC", gcnGi) Then Exit Function
                End Select
                ' Counter = 0
                lngCounter = 0
                
                If rsA.RecordCount <> 0 Then
                    '�ˬd�a�}�O�_�����
                    If Not ChkAddressModify(rsA("BillNo"), rsA("Item"), rsA("PrtSNo") & "", rsA("CustId"), rsA("AddrNo"), rsA("CitemName"), cnTmpMDB) Then Exit Function
                    
                    On Error GoTo TransErr
                    rsPrtSNOTmp.AddNew
                    ' Loop rsA�C�@�����
                    Do While Not rsA.EOF And rsA.RecordCount > 0
                        ' �YCounter = 0

                        If lngCounter = 0 Then
                            If Not GetrsFmtFld(rsFmtFld, rsFmtFldTmp(), 0, "SELECT A.*, B.FldName FROM FmtFld A, FldList B WHERE A.FldId = B.FldID AND A.FmtID = '" & _
                               strFmtID & "' AND MID(B.FldID, 1, 2) >= 'A0' AND MID(B.FldID, 1, 2) <= 'A9' And A.Type = B.Type ORDER BY A.FldID", cnRptMDB) Then Exit Function
                            
'                            If Not GetRS(rsFmtFld, "SELECT A.*, B.FldName FROM FmtFld A, FldList B WHERE A.FldId = B.FldID AND A.FmtID = '" & _
                               strFmtID & "' AND MID(B.FldID, 1, 2) >= 'A0' AND MID(B.FldID, 1, 2) <= 'A9' And A.Type = B.Type ORDER BY A.FldID", cnRptMDB) Then Exit Function
                            Do While Not rsFmtFld.EOF
                                ' ���򥻥x������T
                                If rsFmtFld("FldId") <> "A07" And rsFmtFld("FldId") <> "A08" Then
                                    rsPrtSNOTmp(rsFmtFld("FldID").Value).Value = rsA(rsFmtFld("FldName").Value)
                                End If
                                
'                                rsPrtSNOTmp("A01").Value = rsA("BillNO").Value
'                                rsPrtSNOTmp("A02").Value = rsA("PrtSNO").Value
'                                rsPrtSNOTmp("A03").Value = rsA("CustId").Value
'                                rsPrtSNOTmp("A04").Value = rsA("ShouldDate").Value
'                                rsPrtSNOTmp("A05").Value = rsA("ClctEn").Value
'                                rsPrtSNOTmp("A06").Value = rsA("ClctName").Value
'                                rsPrtSNOTmp("A10").Value = rsA("PrtCount").Value
                                rsFmtFld.MoveNext
                            Loop
                        End If
                        lngCompCode = rsA("CompCode")
                        ' 1.Counter + 1
                        lngCounter = lngCounter + 1
                        If lngCounter <= 20 Then
                            
                            ' 2.�ھ�Counter�ȡA�]�w�������ȡA
                            '   �ҡG[AAxx],[ABxx],[ACxx],[ADxx],
                            '       [AExx],[AFxx],(xx��01~10)
    '
                            If Not GetrsFmtFld(rsFmtFld, rsFmtFldTmp(), 4, "SELECT A.*, B.FldName FROM FmtFld A, FldList B WHERE A.FldId = B.FldID AND A.FmtID = '" & _
                               strFmtID & "' AND MID(B.FldID, 1, 2) >= 'AA' AND MID(B.FldID, 1, 2) <= 'AZ' And A.Type = B.Type ORDER BY A.FldID", cnRptMDB) Then Exit Function
                            
                            On Error Resume Next
                            Do While Not rsFmtFld.EOF
                                If Mid(rsFmtFld("FldID").Value, 3, 2) = Format(lngCounter, "00") Then
                                    rsPrtSNOTmp(rsFmtFld("FldID").Value).Value = rsA(rsFmtFld("FldName").Value)
                                End If
                                rsFmtFld.MoveNext
                            Loop
'                            rsPrtSNOTmp("AA" & Format(lngCounter, "00")).Value = rsA("CitemCode").Value
'                            rsPrtSNOTmp("AB" & Format(lngCounter, "00")).Value = rsA("CitemName").Value
'                            rsPrtSNOTmp("AC" & Format(lngCounter, "00")).Value = rsA("ShouldAmt").Value
'                            rsPrtSNOTmp("AD" & Format(lngCounter, "00")).Value = rsA("RealPeriod").Value
'                            rsPrtSNOTmp("AE" & Format(lngCounter, "00")).Value = rsA("RealStartDate").Value
'                            rsPrtSNOTmp("AF" & Format(lngCounter, "00")).Value = rsA("RealStopDate").Value
'                            rsPrtSNOTmp("AH" & Format(lngCounter, "00")).Value = rsA("RealDate").Value
'                            rsPrtSNOTmp("AI" & Format(lngCounter, "00")).Value = rsA("RealAmt").Value
'                            rsPrtSNOTmp("AJ" & Format(lngCounter, "00")).Value = rsA("Budgetperiod").Value
'                            rsPrtSNOTmp("AK" & Format(lngCounter, "00")).Value = rsA("CMCode").Value
'                            rsPrtSNOTmp("AL" & Format(lngCounter, "00")).Value = rsA("CMName").Value
                            On Error GoTo TransErr
                            If strSQLOA <> "" Then
                                'OAA 95/01/02 Jacky 1944
                                If rsA("FaciSNo") & "" <> "" Then
                                    If Not GetrsFmtFld(rsFmtFld, rsFmtFldTmp(), 5, "SELECT A.*, B.FldName FROM FmtFld A, FldList B WHERE A.FldId = B.FldID AND A.FmtID = '" & _
                                       strFmtID & "' AND MID(B.FldID, 1, 3) >= 'OAA' AND MID(B.FldID, 1, 3) <= 'OAZ' And A.Type = B.Type ORDER BY A.FldID", cnRptMDB) Then Exit Function
                                    
                                    strxSQLOA = strSQLOA & " SeqNo = '" & rsA("FaciSeqNo") & "'"
                                    '����xSQLOA -> rsOA
                                    If Not GetRS(rsOA, strxSQLOA, gcnGi) Then Exit Function
                                
                                    If rsOA.RecordCount > 0 Then
                                        Do While Not rsFmtFld.EOF
                                            If Right(rsFmtFld("FldID").Value, 2) = Format(lngCounter, "00") Then
                                                rsPrtSNOTmp(rsFmtFld("FldID").Value).Value = rsOA(rsFmtFld("FldName").Value)
                                            End If
                                            rsFmtFld.MoveNext
                                        Loop
                                    End If
                                End If
                            End If
                        Else
                            Call ErrHandle("�Ȥ�s���� [" & lngNowCustID & _
                               "] ���Ȥ�A����������O��ƶW�X20���I�U���o�@�����O���ӱN" & _
                               "�|�Q�ٲ����G" & vbCrLf & "���O���إN�X: [" & rsA("CitemCode").Value & "]" & vbCrLf & _
                               "���O����: [" & rsA("CitemName").Value & "]" & vbCrLf & _
                               "�������B: [" & rsA("ShouldAmt").Value & "]", False, , "GotoFlowProc")
                        End If
                        ' 3.�֭p�����`���B[A07]
                        rsPrtSNOTmp("A07").Value = Val(rsPrtSNOTmp("A07").Value & "") + Val(rsA("ShouldAmt").Value & "")
                        On Error Resume Next
                        '6.�YSLQI����: (��ܻݨ̥ثe�Ƚs�������o�����)  (91/08/19 �ק�)
                        If rsA.AbsolutePosition = 1 Then
                            If strSQLI <> "" Then
                                'xSQLI = SQLI + "<�ثe�Ƚs>"
                                If Not GetrsFmtFld(rsFmtFld, rsFmtFldTmp(), 1, "SELECT A.*, B.FldName FROM FmtFld A, FldList B WHERE A.FldId = B.FldID AND A.FmtID = '" & _
                                   strFmtID & "' AND MID(B.FldID, 1, 1) = 'I' And A.Type = B.Type ORDER BY A.FldID", cnRptMDB) Then Exit Function
                                
                                strxSQLI = strSQLI & " BillNo = '" & rsA("BillNo") & "' And CitemName = '" & rsA("CitemName") & "'"
                                '����xSQLI -> rsI
                                If Not GetRS(rsI, strxSQLI, gcnGi) Then Exit Function
                                If rsI.RecordCount > 0 Then
                                    Do While Not rsFmtFld.EOF
                                        rsPrtSNOTmp(rsFmtFld("FldID").Value).Value = rsI(rsFmtFld("FldName").Value)
                                        rsFmtFld.MoveNext
                                    Loop
                                End If
                            End If
                            
                            '7.�YSLQK����: (��ܻݨ�CompCode,ServiceType���������O�Ѽ�)
                            If strSQLK <> "" Then
                                strxSQLK = strSQLK & " CompCode = " & rsA("CompCode").Value & " And ServiceType = '" & rsA("ServiceType") & "'"
                                If Not GetRS(rsK, strxSQLK, gcnGi) Then Exit Function
                            End If
                            
                            '8.�YSLQL����: (��ܻݨ�CustId, AccountNo���������O�Ѽ�)
                            If strSQLL <> "" Then
                                strxSQLL = strSQLL & " CustId = " & rsA("CustId").Value & " And AccountNo = '" & rsA("AccountNo") & "' And CompCode = " & rsA("CompCode")
                                If Not GetRS(rsL, strxSQLL, gcnGi) Then Exit Function
                                '9.�YSQLL1����: (��ܻݨ̦��O�a�}�������a�}���)
                                If strSQLL1 <> "" Then
                                    '93/05/13 �[�P�_�Ӹ�ƵL��
                                    If rsL.RecordCount > 0 Then
                                        strxSQLL1 = strSQLL1 & rsL("ChargeAddrNO").Value & " And CompCode = " & rsA("CompCode")
                                    Else
                                        strxSQLL1 = strSQLL1 & "0"
                                    End If
                                    If Not GetRS(rsL1, strxSQLL1, gcnGi) Then Exit Function
                                End If
                                
                                '10.�YSQLL2����: (��ܻݨ̶l�H�a�}�������a�}���)
                                If strSQLL2 <> "" Then
                                    '93/05/13 �[�P�_�Ӹ�ƵL��
                                    If rsL.RecordCount > 0 Then
                                        strxSQLL2 = strSQLL2 & rsL("MailAddrNo").Value & " And CompCode = " & rsA("CompCode")
                                    Else
                                        strxSQLL2 = strSQLL2 & "0"
                                    End If
                                    If Not GetRS(rsL2, strxSQLL2, gcnGi) Then Exit Function
                                End If
                                '11.�YSLQM����: (��ܻݨ̵���s���������]�����)
                                If strSQLM <> "" Then
                                    ' xSQLM = SQLM + " '<rsA.BillNo>' "
                                    strxSQLM = strSQLM & "'" & rsA("BillNo").Value & "'"
                                    ' ����xSQLM -> rsM
                                    If Not GetRS(rsM, strxSQLM, gcnGi) Then Exit Function
                                End If
                            End If
                            
                        
                            '4.�YSLQD����: (��ܻݨ̥ثe�Ƚs�������b�����)
                            If strSQLD <> "" Then
                                ' xSQLD = SQLD + "<�ثe�Ƚs>"
                                strxSQLD = strSQLD & rsA("CustId").Value & " And CompCode = " & rsA("CompCode") & " And ServiceType = '" & rsA("ServiceType") & "'"
                                ' ����xSQLD -> rsD
                                If Not GetRS(rsD, strxSQLD, gcnGi) Then Exit Function
                            End If
                            
                            '�Ҧ����`�]�ƪ����
                            If strSQLOB <> "" Then
                                If Not GetrsFmtFld(rsFmtFld, rsFmtFldTmp(), 6, "SELECT A.*, B.FldName FROM FmtFld A, FldList B WHERE A.FldId = B.FldID AND A.FmtID = '" & _
                                   strFmtID & "' AND MID(B.FldID, 1, 3) >= 'OBA' AND MID(B.FldID, 1, 3) <= 'OBZ' And A.Type = B.Type ORDER BY A.FldID", cnRptMDB) Then Exit Function
                                
                                strxSQLOB = strSQLOB & " CustId = " & rsA("CustId") & " And CompCode = " & rsA("CompCode") & " And PRDate is null"
                                '����xSQLOA -> rsOA
                                If Not GetRS(rsOB, strxSQLOB, gcnGi) Then Exit Function
                            
                                Do While Not rsOB.EOF
                                    rsFmtFld.MoveFirst
                                    Do While Not rsFmtFld.EOF
                                        If Right(rsFmtFld("FldID").Value, 2) = Format(rsOB.AbsolutePosition, "00") Then
                                            rsPrtSNOTmp(rsFmtFld("FldID").Value).Value = rsOB(rsFmtFld("FldName").Value)
                                        End If
                                        rsFmtFld.MoveNext
                                    Loop
                                    rsOB.MoveNext
                                Loop
                            End If
                            '�Ҧ��A�Ȫ�SO002
                            '2007/01/08 Jacky 2896 Debby
                            If strSQLOC <> "" Then
                                If Not GetrsFmtFld(rsFmtFld, rsFmtFldTmp(), 7, "SELECT A.*, B.FldName FROM FmtFld A, FldList B WHERE A.FldId = B.FldID AND A.FmtID = '" & _
                                   strFmtID & "' AND MID(B.FldID, 1, 3) >= 'OCA' AND MID(B.FldID, 1, 3) <= 'OCZ' And A.Type = B.Type ORDER BY A.FldID", cnRptMDB) Then Exit Function
                                
                                strxSQLOC = strSQLOC & " " & rsA("CustId") & " And CompCode = " & rsA("CompCode") & " Order By ServiceType"
                                If Not GetRS(rsOC, strxSQLOC, gcnGi) Then Exit Function
                                Do While Not rsOC.EOF
                                    rsFmtFld.MoveFirst
                                    Do While Not rsFmtFld.EOF
                                        If Right(rsFmtFld("FldID").Value, 2) = Format(rsOC.AbsolutePosition, "00") Then
                                            rsPrtSNOTmp(rsFmtFld("FldID").Value).Value = rsOC(rsFmtFld("FldName").Value)
                                        End If
                                        rsFmtFld.MoveNext
                                    Loop
                                    rsOC.MoveNext
                                Loop
                            End If
                            
                            If strSQLOD <> "" Then
                                If Not GetrsFmtFld(rsFmtFld, rsFmtFldTmp(), 7, "SELECT A.*, B.FldName FROM FmtFld A, FldList B WHERE A.FldId = B.FldID AND A.FmtID = '" & _
                                   strFmtID & "' AND MID(B.FldID, 1, 3) >= 'ODA' AND MID(B.FldID, 1, 3) <= 'ODZ' And A.Type = B.Type ORDER BY A.FldID", cnRptMDB) Then Exit Function
                                If rsA("InvSeqNo") & "" <> "" Then
                                    strxSQLOD = strSQLOD & " " & rsA("InvSeqNo")
                                    If Not GetRS(rsOD, strxSQLOD, gcnGi) Then Exit Function
                                    Do While Not rsOD.EOF
                                        rsFmtFld.MoveFirst
                                        Do While Not rsFmtFld.EOF
                                            If Right(rsFmtFld("FldID").Value, 2) = Format(rsOD.AbsolutePosition, "00") Then
                                                rsPrtSNOTmp(rsFmtFld("FldID").Value).Value = rsOD(rsFmtFld("FldName").Value)
                                            End If
                                            rsFmtFld.MoveNext
                                        Loop
                                        rsOD.MoveNext
                                    Loop
                                End If
                            End If
                        
                        End If
                        On Error GoTo ER
                        rsA.MoveNext
                    Loop
                
                
    '----Start �y�{��------
                    ' ����y�{B
                    'B: ���Y�Ȥ�s�W�@���b����
                    '1. �����b��U���һݸ�T:
                    '1.���ӫȤ�򥻸��:
                    'xSQLB = SQLB + "<�ثe�Ƚs>"
                    strxSQLB = strSQLB & lngNowCustID & " And CompCode = " & lngCompCode
                    
                    '����xSQLB -> rsB
                    If Not GetRS(rsB, strxSQLB, gcnGi, adUseClient) Then Exit Function
                    
                    '2.�YSQLC1����: (��ܻݨ̦��O�a�}�������a�}���)
                    If strSQLC1 <> "" Then
                        'xSQLC1 = SQLC1 + "<rsB.ChargeAddrNo>"
                        strxSQLC1 = strSQLC1 & rsB("ChargeAddrNO").Value & " And CompCode = " & lngCompCode
                        '����xSQLC1 -> rsC1
                        If Not GetRS(rsC1, strxSQLC1, gcnGi) Then Exit Function
                    End If
                    
                    '3.�YSLQC2����: (��ܻݨ̸˾��a�}�������a�}���)
                    If strSQLC2 <> "" Then
                        'xSQLC2 = SQLC2 + "<rsB.InstAddrNo>"
                        strxSQLC2 = strSQLC2 & rsB("InstAddrNo").Value & " And CompCode = " & lngCompCode
                        '����xSQLC2 -> rsC2
                        If Not GetRS(rsC2, strxSQLC2, gcnGi) Then Exit Function
                    End If
                    
                    '3.�YSLQC3����: (��ܻݨ̸˾��a�}�������a�}���)
                    If strSQLC3 <> "" Then
                        'xSQLC3 = SQLC3 + "<rsB.MailAddrNo>"
                        strxSQLC3 = strSQLC3 & rsB("MailAddrNo").Value & " And CompCode = " & lngCompCode
                        '����xSQLC2 -> rsC2
                        If Not GetRS(rsC3, strxSQLC3, gcnGi) Then Exit Function
                    End If
                    
                    '5.�YSLQE����: (��ܻݨ̤j�ӽs���������j�����)
                    If strSQLE <> "" Then
                        ' xSQLE = SQLE + " '<rsB.MduId>' "
                        strxSQLE = strSQLE & "'" & rsB("MduID").Value & "' And CompCode = " & lngCompCode
                        ' ����xSQLE -> rsE
                        If Not GetRS(rsE, strxSQLE, gcnGi) Then Exit Function
                    End If
                    
                    '6.�YSLQJ����: (��ܻݨ̦a�}�s���������a�}���)
                    If strSQLJ <> "" Then
                        ' xSQLJ = SQLJ + " '<rsB.MduId>' "
                        strxSQLJ = "Select A.*,B.Description ContractStatusName From ( " & strSQLJ & " CableAddrNo = " & rsB("InstAddrNo").Value & " Union All " & _
                            strSQLJ & " AddrNo2 = " & rsB("InstAddrNo").Value & " Union All " & _
                            strSQLJ & " AddrNo3 = " & rsB("InstAddrNo").Value & " ) A," & GetOwner & "CD062 B Where A.CATVId= B.CodeNo Order By B.Priority"
                        ' ����xSQLJ -> rsJ
                        If Not GetRS(rsJ, strxSQLJ, gcnGi, , , , , , , 1) Then Exit Function
                    End If
                    
                    '93/11/22 CM/CT�� Jacky
                    If strSQLN1 <> "" Then
                        strxSQLN1 = strSQLN1 & " " & rsB("CustId").Value & " And EbtServiceType='CM' ORDER BY UpdTime DESC"
                        ' ����xSQLN1 -> rsN1
                        If Not GetRS(rsN1, strxSQLN1, gcnGi, , , , , , , 1) Then Exit Function
                    End If
                    
                    '93/11/22 CM/CT�� Jacky
                    If strSQLN2 <> "" Then
                        strxSQLN2 = strSQLN2 & " " & rsB("CustId").Value & " And EbtServiceType='CT' ORDER BY UpdTime DESC"
                        ' xSQLN2 -> rsN2
                        If Not GetRS(rsN2, strxSQLN2, gcnGi, , , , , , , 1) Then Exit Function
                    End If
                    
                    '2. Loop���t�\����T��(FmtFld)�����O'F'�}�Y���C�@��:
                    If Not GetrsFmtFld(rsFmtFld, rsFmtFldTmp(), 2, "SELECT A.*, B.FldName FROM FmtFld A, FldList B WHERE A.FldId = B.FldID AND A.FmtID = '" & _
                           strFmtID & "' AND MID(B.FldID, 1, 1) <> 'F' And A.Type = B.Type ORDER BY A.FldID", cnRptMDB) Then Exit Function
            
                    Do While Not rsFmtFld.EOF
                        'Debug.Print rsFmtFld.Fields("FldID")
                        Select Case UCase(rsFmtFld("FldID").Value)
                            '�E  case [C01]: [C01] = rsC1.ZipCode            �l���ϸ�(���O�a�})
                            Case "C01"
                                rsPrtSNOTmp("C01").Value = rsC1("ZipCode").Value
                            '�E  case [C02]: [C02] = rsC1.AreaName           ��F�ϦW��(���O�a�})
                            Case "C02"
                                rsPrtSNOTmp("C02").Value = rsC1("AreaName").Value
                            '�E  case [C03]: [C03] = rsC1.ServName           �A�ȰϦW��(���O�a�})
                            Case "C03"
                                rsPrtSNOTmp("C03").Value = rsC1("ServName").Value
                            '�E  case [C04]: [C04] = rsC2.AreaName           ��F�ϦW��(�˾��a�})
                            Case "C04"
                                rsPrtSNOTmp("C04").Value = rsC2("AreaName").Value
                            '�E  case [C05]: [C05]= rsC2.ZipCode         �l���ϸ�(�˾��a�})
                            Case "C05"
                                rsPrtSNOTmp("C05").Value = rsC2("ZipCode").Value
                            '�E  case [C06]: [C06]= rsC2.BourgName         �����W��(�˾��a�})
                            Case "C06"
                                rsPrtSNOTmp("C06").Value = rsC2("BourgName").Value
                            '�E  case [C07]: [C07]= rsC1.BourgName         �����W��(���O�a�})
                            Case "C07"
                                rsPrtSNOTmp("C07").Value = rsC1("BourgName").Value
                            '�E  case [A08]: [A08] = [A07]�ର����j�g�r��       �����`���B(����)
                            Case "A08"
                                rsPrtSNOTmp("A08").Value = ToChineseN(rsPrtSNOTmp("A07").Value & "")
                            '�E  case [AG01]: [AG01] = [A07]�����`���B�Ӧ��(����)
                            Case "AG01"
                                rsPrtSNOTmp("AG01").Value = ToChineseNumber(rsPrtSNOTmp("A07").Value, 1)
                            '�E  case [AG02]: [AG02] = [A07]�����`���B�B���(����)
                            Case "AG02"
                                rsPrtSNOTmp("AG02").Value = ToChineseNumber(rsPrtSNOTmp("A07").Value, 2)
                            '�E  case [AG03]: [AG03] = [A07]�����`���B�զ��(����)
                            Case "AG03"
                                rsPrtSNOTmp("AG03").Value = ToChineseNumber(rsPrtSNOTmp("A07").Value, 3)
                            '�E  case [AG04]: [AG04] = [A07]�����`���B�a���(����)
                            Case "AG04"
                                rsPrtSNOTmp("AG04").Value = ToChineseNumber(rsPrtSNOTmp("A07").Value, 4)
                            '�E  case [AG05]: [AG05] = [A07]�����`���B�U���(����)
                            Case "AG05"
                                rsPrtSNOTmp("AG05").Value = ToChineseNumber(rsPrtSNOTmp("A07").Value, 5)
                            '�E  case [AG06]: [AG06] = [A07]�����`���B�B�U���(����)
                            Case "AG06"
                                rsPrtSNOTmp("AG06").Value = ToChineseNumber(rsPrtSNOTmp("A07").Value, 6)
                            '�E  ��: [A01]~[A07], [AAxx], [ABxx], [ACxx], [ADxx], [AExx], [AFxx]���C��
                            '    �]�o�Ǥw�b���e�L�B���o (xx��01~10)
                            Case Else
                                Select Case Left(UCase(rsFmtFld("FldID").Value), 2)
                                    Case "CA"       '���O�a�}
                                        If rsC1.RecordCount > 0 Then
                                            rsPrtSNOTmp(rsFmtFld("FldID").Value).Value = rsC1(rsFmtFld("FldName").Value)
                                        End If
                                    Case "CB"       '�˾��a�}
                                        If rsC2.RecordCount > 0 Then
                                            rsPrtSNOTmp(rsFmtFld("FldID").Value).Value = rsC2(rsFmtFld("FldName").Value)
                                        End If
                                    Case "CC"       '�l�H�a�}
                                        If rsC3.RecordCount > 0 Then
                                            rsPrtSNOTmp(rsFmtFld("FldID").Value).Value = rsC3(rsFmtFld("FldName").Value)
                                        End If
                                    Case "LA"       '���O�a�}
                                        If rsL1.RecordCount > 0 Then
                                            rsPrtSNOTmp(rsFmtFld("FldID").Value).Value = rsL1(rsFmtFld("FldName").Value)
                                        End If
                                    Case "LB"       '�l�H�a�}
                                        If rsL2.RecordCount > 0 Then
                                            rsPrtSNOTmp(rsFmtFld("FldID").Value).Value = rsL2(rsFmtFld("FldName").Value)
                                        End If
                                    Case "NA"
                                        If rsN1.RecordCount > 0 Then
                                            rsPrtSNOTmp(rsFmtFld("FldID").Value).Value = rsN1(rsFmtFld("FldName").Value)
                                        End If
                                    Case "NB"
                                        If rsN2.RecordCount > 0 Then
                                            rsPrtSNOTmp(rsFmtFld("FldID").Value).Value = rsN2(rsFmtFld("FldName").Value)
                                        End If
                                    Case Else
                                        Select Case Left(UCase(rsFmtFld("FldID").Value), 1)
                                            Case "B"
                                                If rsB.RecordCount > 0 Then
                                                    rsPrtSNOTmp(rsFmtFld("FldID").Value).Value = rsB(rsFmtFld("FldName").Value)
                                                End If
                                            Case "D"
                                                If rsD.RecordCount > 0 Then
                                                    rsPrtSNOTmp(rsFmtFld("FldID").Value).Value = rsD(rsFmtFld("FldName").Value)
                                                End If
                                            Case "E"
                                                If rsE.RecordCount > 0 Then
                                                    rsPrtSNOTmp(rsFmtFld("FldID").Value).Value = rsE(rsFmtFld("FldName").Value)
                                                End If
                                            Case "J"
                                                If rsJ.RecordCount > 0 Then
                                                    rsPrtSNOTmp(rsFmtFld("FldID").Value).Value = rsJ(rsFmtFld("FldName").Value)
                                                End If
                                            Case "K"
                                                If rsK.RecordCount > 0 Then
                                                    rsPrtSNOTmp(rsFmtFld("FldID").Value).Value = rsK(rsFmtFld("FldName").Value)
                                                End If
                                            Case "L"
                                                If rsL.RecordCount > 0 Then
                                                    rsPrtSNOTmp(rsFmtFld("FldID").Value).Value = rsL(rsFmtFld("FldName").Value)
                                                End If
                                            Case "M"
                                                If rsM.RecordCount > 0 Then
                                                    rsPrtSNOTmp(rsFmtFld("FldID").Value).Value = rsM(rsFmtFld("FldName").Value)
                                                End If
                                        End Select
                                End Select
                        End Select
                        rsFmtFld.MoveNext
                    Loop
                    
                    '3. �YSQLF����: (��ܻݨ��e���ꦬ���)
                    If strSQLF <> "" Then
                        On Error Resume Next
                        
                        ' ���e���L��Ǹ�
                        '"select max(PrtSNo) from SO034 where CustId=<�ثe�Ƚs>
                        'and PrtSNo < 'YYYYMM' "
                        ' ��: YYYYMM���L��~����r�� , �Y���q�Ω�浧���O��ڦC�LSO3262A��,
                        ' �hYYYYMM���L��I��~����r��
                        '90/09/07 �bNTY�ק諸.(Pierre,KC,Lawrence,Jacky )
'                        If strForm = "FRMSO3262A" Then
'                            If Not GetRS(rsSO034, "SELECT MAX(PrtSNO) AS M FROM SO034 WHERE " & "CustID = " & lngNowCustID & "AND SO034.PrtSNo < '" & objForm.gymPrtYM2.GetValue & "'") Then Exit Function
'                        ElseIf strForm = "FRMSO3261A" Then
'                            If Not GetRS(rsSO034, "SELECT MAX(PrtSNO) AS M FROM SO034 WHERE " & "CustID = " & lngNowCustID & "AND SO034.PrtSNo < '" & objForm.gymPrtYM.GetValue & "'") Then Exit Function
'                        Else
'                            If Not GetRS(rsSO034, "SELECT MAX(PrtSNO) AS M FROM SO034 WHERE " & "CustID = " & lngNowCustID) Then Exit Function
'                        End If
                        '90/09/07 �bNTY�ק諸.(Pierre,KC,Lawrence,Jacky )
                        Dim strFieldF As String
                        Dim strBillNo As String
                        strFieldF = " BillNo ,Item, " & strSQLF
                        
                        If InStr(UCase(strSQLF), "CITEMCODE") = 0 Then strFieldF = strFieldF & ", CITEMCODE"
                        If InStr(UCase(strSQLF), "REALDATE") = 0 Then strFieldF = strFieldF & ", REALDATE"
                        If InStr(UCase(strSQLF), "UCCODE") = 0 Then strFieldF = strFieldF & ", UCCODE"
                        If InStr(UCase(strSQLF), "REALAMT") = 0 Then strFieldF = strFieldF & ", REALAMT"
                        If InStr(UCase(strSQLF), "SHOULDAMT") = 0 Then strFieldF = strFieldF & ", SHOULDAMT"
                        
                        '�Ф��n�A�ç�
                        '90/10/31 Pierre & KC & Janis �Q�׫ᵲ�G RealAmt >0 And PeriodFlag =1 And BillNo = 'BT' Order By RealDate Desc, Substr(BillNo,7,1)
                        
                        '91/10/31 ���Ͱ򥻥x�Ω���
                        strWhere = ""
                        If Val(rsPara("Para19") & "") = 1 Then strWhere = " And A.CitemCode In (Select CodeNo From " & GetOwner & "CD019 Where RefNo =1) "
                        
                        If Not GetRS(rsSO034, "Select * From (Select A.BillNo,Max(RealDate) RealDate From (SELECT " & strFieldF & " FROM " & GetOwner & "SO033 A WHERE A.CustID = " & lngNowCustID & _
                            " And Nvl(CancelFlag,0)=0 Union All  SELECT " & strFieldF & " FROM " & GetOwner & "SO034 A Where  A.CustID = " & lngNowCustID & "  And Nvl(CancelFlag,0)=0  ) A," & GetOwner & "CD019 B Where A.CitemCode = B.CodeNO And RealAmt > 0 And B.PeriodFlag = 1 And A.RealDate is Not Null And A.UCCode Is Null Group By BillNo) A Order By RealDate Desc") Then Exit Function
                        
                        ' �Y���omax(PrtSNo)����, �h:
                        lngCounter = 1
                        '91/10/31 �b������ʹ���
                        Do While Not rsSO034.EOF And Val(rsPara("Para20") & "") >= lngCounter
                            ' xSQLF = SQLF + " '<max(PrtSNo) >' "
                            'strxSQLF = strSQLF & rsSO034("M").Value
                            If Not GetRS(rsF, "Select A.* From (SELECT " & strFieldF & " FROM " & GetOwner & "SO033 A WHERE A.BillNo = '" & rsSO034("BillNo") & _
                                "' And Nvl(CancelFlag,0)=0 " & strWhere & " Union All  SELECT " & strFieldF & " FROM " & GetOwner & "SO034 A Where  A.BillNo = '" & rsSO034("BillNo") & "'  And Nvl(CancelFlag,0)=0  ) A Where A.BillNo ='" & rsSO034("BillNo") & "' And A.RealDate is Not Null And A.UCCode Is Null " & strWhere & " Order By A.RealDate Desc") Then Exit Function
                            
                            ' ����xSQLF a rsF
                            'If Not GetRS(rsF, strxSQLF, gcnGi) Then Exit Function
                            Dim strFldStr As String
                            ' loop rsF �C�����:
                            Do While Not rsF.EOF And Val(rsPara("Para20") & "") >= lngCounter
                                ' [F01] = [F01] + rsF.ShouldAmt       �֭p�e�������`���B
                                rsPrtSNOTmp("F01").Value = Val(rsPrtSNOTmp("F01").Value & "") + Val(rsF("ShouldAmt").Value & "")
                                
                                ' [F02] = [F02] + rsF.RealAmt         �֭p�e���ꦬ�`���B
                                rsPrtSNOTmp("F02").Value = Val(rsPrtSNOTmp("F02").Value & "") + Val(rsF("RealAmt").Value & "")
                                ' �YrsF.CitemCode = <�򥻥x���O���إN�X>
                                If rsF("CitemCode").Value = GetValueFromRs2("SELECT " & _
                                   "CodeNo FROM " & GetOwner & "CD019 WHERE RefNo = 1", gcnGi) Then
                                    ' �B[F03]<rsF.RealDate, �h:
                                    If rsPrtSNOTmp("F03").Value < rsF("RealDate").Value Then
                                        '[F03] = rsF.RealDate                ���򥻥x�e�����O��
                                        rsPrtSNOTmp("F03").Value = rsF("RealDate").Value
                                    End If
                                End If
                                '2. loop���t�\����T�ɤ����C�@���H  'F'�}�Y�����:
                                'If strFldStr = "" Then
                                    If Not GetrsFmtFld(rsFmtFld, rsFmtFldTmp(), 3, "SELECT FmtFld.FmtID,FmtFld.FldId,FmtFld.Type, FldList.FldName FROM FmtFld, FldList WHERE FmtFld.FldId = FldList.FldID AND FmtFld.FmtID = '" & _
                                        strFmtID & "' AND MID(FmtFld.FldID, 1, 2) >= 'FA' AND MID(FmtFld.FldID, 1, 2) <= 'FZ' And FmtFld.Type = FldList.Type  ORDER BY FmtFld.FldID", cnRptMDB) Then Exit Function
                                    
                                '    If Not rsFmtFld.EOF Then strFldStr = rsFmtFld.GetString(, , ",", ",", "")
                                'End If
                                
                                Do While Not rsFmtFld.EOF
                                    If Mid(rsFmtFld("FldID").Value, 3, 2) = Format(lngCounter, "00") Then
                                        rsPrtSNOTmp(rsFmtFld("FldID").Value).Value = rsF(rsFmtFld("FldName").Value)
                                    End If
                                    rsFmtFld.MoveNext
                                Loop
                                
'                                '�E  case [FAxx]: [FAxx] = rsF.CitemCode     �e�����إN�X
'                                If InStr(1, strFldStr, "FA" & Format(lngCounter, "00")) > 0 Then
'                                    rsPrtSNOTmp("FA" & Format(lngCounter, "00")).Value = rsF("CitemCode").Value
'                                End If
'
'                                '�E  case [FBxx]: [FBxx] = rsF.CitemName     �e�����ئW��
'                                If InStr(1, strFldStr, "FB" & Format(lngCounter, "00")) > 0 Then
'                                    rsPrtSNOTmp("FB" & Format(lngCounter, "00")).Value = rsF("CitemName").Value
'                                End If
'
'                                '�E  case [FCxx]: [FCxx] = rsF.ShouldAmt     �e�������������B
'                                If InStr(1, strFldStr, "FC" & Format(lngCounter, "00")) > 0 Then
'                                    rsPrtSNOTmp("FC" & Format(lngCounter, "00")).Value = rsF("ShouldAmt").Value
'                                End If
'
'                                '�E  case [FDxx]: [FDxx] = rsF.REalPeriod     �e�����ش���
'                                If InStr(1, strFldStr, "FD" & Format(lngCounter, "00")) > 0 Then
'                                    rsPrtSNOTmp("FD" & Format(lngCounter, "00")).Value = rsF("RealPeriod").Value
'                                End If
'
'                                '�E  case [FExx]: [FExx] = rsF.RealStartDate     �e�����ش����_�l��
'                                If InStr(1, strFldStr, "FE" & Format(lngCounter, "00")) > 0 Then
'                                    rsPrtSNOTmp("FE" & Format(lngCounter, "00")).Value = rsF("RealStartDate").Value
'                                End If
'
'                                '�E  case [FFxx]: [FFxx] = rsF.RealStopDate     �e�����ش����I���
'                                If InStr(1, strFldStr, "FF" & Format(lngCounter, "00")) > 0 Then
'                                    rsPrtSNOTmp("FF" & Format(lngCounter, "00")).Value = rsF("RealStopDate").Value
'                                End If
'
'                                '�E  case [FGxx]: [FGxx] = rsF.RealAmt     �e�����ؤJ�b���B
'                                If InStr(1, strFldStr, "FG" & Format(lngCounter, "00")) > 0 Then
'                                    rsPrtSNOTmp("FG" & Format(lngCounter, "00")).Value = rsF("RealAmt").Value
'                                End If
'
'                                '�E  case [FHxx]: [FHxx] = rsF.UCCode     �e�����إ�����]�N�X
'                                If InStr(1, strFldStr, "FH" & Format(lngCounter, "00")) > 0 Then
'                                    rsPrtSNOTmp("FH" & Format(lngCounter, "00")).Value = rsF("UCCode").Value
'                                End If
'
'                                '�E  case [FIxx]: [FIxx] = rsF.STCode     �e�����صu����]�N�X
'                                If InStr(1, strFldStr, "FI" & Format(lngCounter, "00")) > 0 Then
'                                    rsPrtSNOTmp("FI" & Format(lngCounter, "00")).Value = rsF("STCode").Value
'                                End If
'
'                                '�E  case [FJxx]: [FJxx] = rsF.RealDate     �e�����ئ��O���
'                                If InStr(1, strFldStr, "FJ" & Format(lngCounter, "00")) > 0 Then
'                                    rsPrtSNOTmp("FJ" & Format(lngCounter, "00")).Value = rsF("RealDate").Value
'                                End If

                                lngCounter = lngCounter + 1
                                rsF.MoveNext
                            Loop
                            rsSO034.MoveNext
                        Loop
                        
                        On Error GoTo ER
                    End If
                    
                    '4. �s�W�@���b���Ʀܱb��Ȧs�ɤ�, �䤤:
                    '[A03]=<�ثe�Ƚs>, ��L���b�W�z�{�����w���o.
                    '�ò֭p�b���BillCount = BillCount + 1
                    rsPrtSNOTmp.Update
                    lngBillcount = lngBillcount + 1
                    '93/09/21
                    If strForm <> "FRMSO32A2A" Then
                        '5. �N�ӵ��b��ҹ������������Ӹ��(�i�঳�h��)�����C�L���ƥ[1
                        '   �ç�s�C�L�H���P�C�L�ɶ�, SQL:
                        '   "update SO033 set PrtCount=PrtCount+1, PrintEn=<p_UpdEn>,
                        '   PrintTime=<��ɮɶ�> where PrtSNo=<�����b�椧�L��Ǹ�>"
                        ' Error 90/08/20 Jacky �ק� Where ����PrsA��Where ����
                        If ExecuteSQL("UPDATE " & GetOwner & StrTableName & " A SET PrtCount = PrtCount + 1, " & _
                           "PrintEn = '" & garyGi(1) & "', PrintTime = SYSDATE " & _
                           Mid(UCase(rsA.Source), InStr(1, UCase(rsA.Source), "WHERE"), InStr(1, UCase(rsA.Source), "ORDER") - InStr(1, UCase(rsA.Source), "WHERE")), gcnGi, , False) <> giOK Then Exit Function
                    End If
                    
                    If strForm = "FRMSO3261A" Then
                        '6. �Y���q�{���Ω���C�L�A�h�s�W�@��LOG��Ʀܳ�ڦC�LLOG��(SO060), �U��쬰:
                        '  .�L����W�� PrinterName=<p_PrtName>
                        '  .���ǽs�� Ord = BillCount
                        '  .��ڽs�� BillNo = �����b�椧��ڽs��
                        '  .�L��Ǹ� PrtSNo = �����b�椧�L��Ǹ�
                        If ExecuteSQL("INSERT INTO " & GetOwner & "SO060 (PRINTERNAME, ORD, BILLNO) VALUES ('" & _
                            objForm.cboPrinterName.Text & "', " & lngBillcount & ", " & lngNowCustID & ")", gcnGi, , False) <> giOK Then Exit Function
                    End If
                                    
                    '7. �N�b���ƹ����ܼƥ��M�����
                    ' ** �]���S���s�b�ܼƸ̡A�ҥH�����M��
                    '----End �y�{��------
                    rsPrtSNOTmp.Update
                    'cnTmpMDB.Execute GetTmpMDBExecuteStr(rsPrtSNOTmp, "PrtSNoTmp")
                    'rsPrtSNOTmp.CancelUpdate
                End If
                CloseRecordset rsA
            End If
            If rs.AbsolutePosition Mod 10 = 0 Then DoEvents
            rs.MoveNext
        Loop
        'If rsPrtSNOTmp.RecordCount > 0 Then rsPrtSNOTmp.Update
        On Error Resume Next
        If strForm <> "FRMSO1100BMDI" Then objForm.Picture1.Visible = False
        On Error GoTo ER
        lngAffecteds = lngBillcount
        For intLoop = 0 To UBound(rsFmtFldTmp())
            Call CloseRecordset(rsFmtFldTmp(intLoop))
        Next
        Call CloseRecordset(rsPara)
        Call CloseRecordset(rsFmtFld)
        Call CloseRecordset(rsTmp)
        GotoFlowProc = True
    Exit Function
TransErr:
    If ErrHandle(Err.Description, True, , "GotoFlowProc") Then Resume 0
    objForm.Picture1.Visible = False
    Exit Function
ER:
    If ErrHandle(Err.Description, True, , "GotoFlowProc") Then Resume 0
    objForm.Picture1.Visible = False
End Function

Private Function GetrsFmtFld(rs As ADODB.Recordset, rsTmp() As ADODB.Recordset, _
    intIndex As Integer, strSQL As String, cnMDB As ADODB.Connection) As Boolean
    On Error GoTo ChkErr
        If rsTmp(intIndex) Is Nothing Then Set rsTmp(intIndex) = New ADODB.Recordset
        Set rs = New ADODB.Recordset
        If rsTmp(intIndex).State = adStateClosed Then
            If Not GetRS(rs, strSQL, cnMDB) Then Exit Function
            Set rsTmp(intIndex) = rs.Clone
        Else
            Set rs = rsTmp(intIndex)
            If rs.RecordCount > 0 Then rs.MoveFirst
        End If
        GetrsFmtFld = True
    Exit Function
ChkErr:
    ErrSub FormName, "GetrsFmtFld"
End Function

Private Function ChkAddressModify(strBillNo As String, _
        lngItem As Long, strPrtSNo As String, _
        lngCustId As Long, lngADDRNO As Long, _
        strCitemName As String, _
        cnTmpMDB As ADODB.Connection, _
        Optional strNewAddress As String = "") As Boolean
    On Error GoTo ChkErr
    Dim rs As New ADODB.Recordset
    Dim strOldAddress As String
        If Not GetRS(rs, "Select CustName,InstAddrNo,ChargeAddrNo,ChargeAddress From " & GetOwner & "So001 Where CustId = " & lngCustId) Then Exit Function
        If lngADDRNO <> rs("InstAddrNo") And lngADDRNO <> rs("ChargeAddrNo") Then
            strOldAddress = GetRsValue("Select Address From " & GetOwner & "So014 Where AddrNo = " & lngADDRNO) & ""
            cnTmpMDB.Execute "Insert Into BillAddrModify (BillNo,Item,PrtSNo,CustId,CustName,CitemName,OldAddrNo,OldAddress,NewAddrNo,NewAddress) values (" & _
                "'" & strBillNo & "'," & lngItem & ",'" & strPrtSNo & "'," & lngCustId & ",'" & rs("CustName") & _
                "','" & strCitemName & "'," & lngADDRNO & ",'" & strOldAddress & "'," & rs("ChargeAddrNo") & ",'" & rs("ChargeAddress") & "')"
        Else
            strNewAddress = ""
        End If
        ChkAddressModify = True
    Exit Function
ChkErr:
    ErrSub FormName, "ChkAddressModify"
End Function

Public Function ToChineseNumber(lngN As Long, Optional Pos = 0) As String
Dim lngX As Long
Dim strF0()
Dim strF1()
Dim strF2 As String
Dim strOut As String
    strF0 = Array("�s", "��", "�L", "��", "�v", "��", "��", "�m", "��", "�h", "�B")
    strF1 = Array("�ոU", "�B�U", "�U", "�a", "��", "�B", "")
    strF2 = Format(lngN, "0000000")
    If Pos <> 0 Then
        ToChineseNumber = strF0(Val(Mid(strF2, Len(strF2) - Pos + 1, 1)))
        Exit Function
    End If
    
    For lngX = 1 To 7
        strOut = strF0(Val(Mid(strF2, 8 - lngX, 1))) & strF1(7 - lngX) & strOut
    Next
    ToChineseNumber = strOut & "��"
End Function

'���u��ϥ�
Public Function GetRStartDate(strCitemCode As String, lngCustId As Long, _
        Optional objForm As Object, Optional blnUseSo003 As Boolean = True, _
        Optional strFaciSeqNo As String, Optional lngSO003BudgetPeriod As Long, _
        Optional rs003 As ADODB.Recordset) As String
  On Error GoTo ChkErr
    '�ھګȤ�s���A���O���إN�X�A�ܫȤ�g���ʦ��O������(So003)��������(Clctdate)
    '�Y���{���Ω󬣤u��U(ZZ=1)�A�h����u�{���ݭȨӸӤu�椧1�B���u��� 2�B�w�����
    Dim varClctDate As Variant
    Dim strSQL As String
    Dim lngTmpCustId As Long
    Dim intTmpItem As Long
    '���o���O���إN�X
    intTmpItem = Val(strCitemCode)
    If blnUseSo003 Then
        If rs003 Is Nothing Then Set rs003 = New ADODB.Recordset
        
        strSQL = "Select * From " & GetOwner & "So003 Where CustId=" & lngCustId & " and CItemCode=" & intTmpItem
        If strFaciSeqNo <> "" Then strSQL = strSQL & " And FaciSeqNo = '" & strFaciSeqNo & "'"
        'If Val(strSeqNo) <> 0 Then strsql = strsql & " And SeqNo = " & strSeqNo
        If Not GetRS(rs003, strSQL) Then Exit Function
        
        If Not rs003.EOF Then
            If rs003("ClctDate") & "" <> "" Then
                lngSO003BudgetPeriod = Val(rs003("BudgetPeriod") & "")
                GetRStartDate = Format(rs003("ClctDate") & "", "YYYY/MM/DD")
                Exit Function
            End If
        End If
    End If
        
    '�P�_�Ӥu��O�_�����u��
    If objForm Is Nothing Then
        GetRStartDate = Format(RightDate, "YYYY/MM/DD")
    Else
        On Error Resume Next
        With objForm
            If .gdtFinTime.GetValue <> "" Then
                GetRStartDate = .gdtFinTime.GetDate(True)
            ElseIf .gdtResvTime.GetValue <> "" Then
                    GetRStartDate = .gdtResvTime.GetDate(True)
                Else
                    GetRStartDate = Format(RightDate, "YYYY/MM/DD")
            End If
        End With
    End If
  Exit Function
ChkErr:
    ErrSub FormName, "GetRStartDate"
End Function

'���u��ϥ�
Public Function ReturnRealStartDate(strCitemCode As String, lngCustId As Long, _
    Optional strResvTime As String, Optional ByRef objForm As Object, _
    Optional strFaciSeqNo As String, Optional ByRef lngBudgetPeriod As Long) As String
  On Error GoTo ChkErr
    '�ھګȤ�s���A���O���إN�X�A�ܫȤ�g���ʦ��O������(So003)��������(Clctdate)
    '�Y���{���Ω󬣤u��U(ZZ=1)�A�h����u�{���ݭȨӸӤu�椧1�B���u��� 2�B�w�����
    Dim strClctDate As String
    Dim rsClctDate As New ADODB.Recordset
    Dim strSQL As String
    Dim lngTmpCustId As Long
    Dim intTmpItem As Long
    '���o���O���إN�X
    intTmpItem = Val(strCitemCode)
    strSQL = "Select ClctDate,BudgetPeriod From " & GetOwner & "So003 Where CustId=" & lngCustId & " and CitemCode=" & intTmpItem
    If strFaciSeqNo <> "" Then strSQL = strSQL & " And FaciSeqNo = '" & strFaciSeqNo & "'"
    
    rsClctDate.MaxRecords = 1
    If Not GetRS(rsClctDate, strSQL, gcnGi, adUseClient) Then Exit Function
    'OpenDB.OpenRecordset rsClctDate, strSQL, gcnGi, adOpenForwardOnly, adLockReadOnly, adUseClient, False
    If rsClctDate.RecordCount > 0 Then
        strClctDate = rsClctDate("ClctDate").Value & ""
        If strClctDate <> "" Then
            ReturnRealStartDate = strClctDate
            lngBudgetPeriod = Val(rsClctDate("BudgetPeriod") & "")
            Set rsClctDate = Nothing
            Exit Function
        End If
    End If
'    If Not objForm Is Nothing Then
'        ReturnRealStartDate = Format(rightnow, "YYYY/MM/DD")
'        Exit Function
'    End If
    '�P�_�Ӥu��O�_�����u��
    If UCase(Left(Screen.ActiveForm.Name, 8)) = "FRMSO111" Then
        If Screen.ActiveForm.gdtFinTime.GetDate(True) <> "" Then
            ReturnRealStartDate = Screen.ActiveForm.gdtFinTime.GetDate(True)
        ElseIf strResvTime <> "" Then
                ReturnRealStartDate = strResvTime
        Else
                ReturnRealStartDate = Format(RightNow, "YYYY/MM/DD")
        End If
    Else
        ReturnRealStartDate = Format(RightNow, "YYYY/MM/DD")
    End If
  Exit Function
ChkErr:
    ErrSub "Utility3", "ReturnRealStartDate"
End Function

Public Function InsertToOracle(ByVal strTable As String, _
            rsSox As ADODB.Recordset, _
            Optional strWhere As String, _
            Optional lngAffected As Long = 0, _
            Optional blnAddNew As Boolean = True) As Boolean
    On Error GoTo ChkErr
    Dim intLoop As Integer
    Dim rsTmp As New ADODB.Recordset
    Dim strField As String
    Dim strValue As String
    Dim strSQL As String
    Dim strFullField As String
        If rsSox.State = adStateClosed Then Exit Function
        If strWhere = "" Then strWhere = " Where 1 = 0 "
        
        If InStr(1, strTable, GetOwner) = 0 Then strTable = GetOwner & strTable
        
        If Not GetRS(rsTmp, "Select * From " & strTable & " " & strWhere, , adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
        Call ChkHaveField(strFullField, "", rsSox)
        If rsTmp.RecordCount = 0 And blnAddNew Then
            For intLoop = 0 To rsTmp.Fields.Count - 1
                '�ˬd�O���������
                If ChkHaveField(strFullField, rsTmp(intLoop).Name, rsSox) Then
                    If Len(rsSox(rsTmp(intLoop).Name) & "") > 0 Then
                        strField = strField & "," & rsTmp(intLoop).Name
                        If rsTmp(intLoop).Type = adDBTimeStamp Then
                            strValue = strValue & "," & GetNullString(rsSox(rsTmp(intLoop).Name).Value, giDateV, giOracle, True)
                        Else
                            strValue = strValue & "," & Replace(GetNullString(rsSox(rsTmp(intLoop).Name).Value, giStringV, giOracle), ",", "'||chr(44)||'")
                        End If
                    End If
                End If
            Next
            strValue = "Insert Into " & strTable & " (" & Mid(strField, 2) & ") Values (" & Mid(strValue, 2) & ")"
        Else
            For intLoop = 0 To rsTmp.Fields.Count - 1
                If ChkHaveField(strFullField, rsTmp(intLoop).Name, rsSox) Then
                    If rsTmp(intLoop).Type = adDBTimeStamp Then
                        strField = strField & "," & rsTmp(intLoop).Name & "=" & GetNullString(rsSox(rsTmp(intLoop).Name).Value, giDateV, giOracle, True)
                    Else
                        strField = strField & "," & rsTmp(intLoop).Name & "=" & Replace(GetNullString(rsSox(rsTmp(intLoop).Name).Value, giStringV, giOracle), ",", "'||chr(44)||'")
                    End If
                End If
            Next
            strValue = "Update " & strTable & " Set " & Mid(strField, 2) & " " & strWhere
        End If
        gcnGi.Execute strValue, lngAffected
        'If Not ExecuteCommand(strValue, gcnGi, lngAffected) Then Exit Function
'        If rsTmp.RecordCount = 0 Then rsTmp.AddNew
'        For intLoop = 0 To rsTmp.Fields.Count - 1
'             rsTmp(intLoop) = GetFieldValue(rsSox, rsTmp(intLoop).Name)
'        Next
'        rsTmp.Update
'        lngAffected = 1
        Call CloseRecordset(rsTmp)
        InsertToOracle = True
    Exit Function
ChkErr:
    ErrSub FormName, "InsertToOracle"
End Function

Public Function InsertToOracle2(ByVal strTable As String, _
            rsSox As ADODB.Recordset, _
            Optional strWhere As String, _
            Optional lngAffected As Long = 0, _
            Optional blnAddNew As Boolean = True) As Boolean
    On Error GoTo ChkErr
    Dim intLoop As Integer
    Dim rsTmp As New ADODB.Recordset
    Dim strField As String
    Dim strValue As String
    Dim strSQL As String
    Dim strFullField As String
    Dim varFieldX() As String
    Dim varValueX() As String
        If rsSox.State = adStateClosed Then Exit Function
        If strWhere = "" Then strWhere = " Where 1 = 0 "
        
        If InStr(1, strTable, GetOwner) = 0 Then strTable = GetOwner & strTable
        
        If Not GetRS(rsTmp, "Select * From " & strTable & " " & strWhere, , adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
        Call ChkHaveField(strFullField, "", rsSox)
        ReDim varFieldX(rsTmp.Fields.Count - 1) As String
        ReDim varValueX(rsTmp.Fields.Count - 1) As String
        If rsTmp.RecordCount = 0 And blnAddNew Then
            For intLoop = 0 To rsTmp.Fields.Count - 1
                '�ˬd�O���������
                If ChkHaveField(strFullField, rsTmp(intLoop).Name, rsSox) Then
                    If Len(rsSox(rsTmp(intLoop).Name) & "") > 0 Then
                        varFieldX(intLoop) = "," & rsTmp(intLoop).Name
                        If rsTmp(intLoop).Type = adDBTimeStamp Then
                            varValueX(intLoop) = "," & GetNullString(rsSox(rsTmp(intLoop).Name).Value, giDateV, giOracle, True)
                        Else
                            varValueX(intLoop) = "," & Replace(GetNullString(rsSox(rsTmp(intLoop).Name).Value, giStringV, giOracle), ",", "'||chr(44)||'")
                        End If
                    End If
                End If
            Next
            strField = Mid(Join(varFieldX, ""), 2)
            strValue = Mid(Join(varValueX, ""), 2)
            strValue = "Insert Into " & strTable & " (" & strField & ") Values (" & strValue & ")"
        Else
            For intLoop = 0 To rsTmp.Fields.Count - 1
                If ChkHaveField(strFullField, rsTmp(intLoop).Name, rsSox) Then
                    If rsTmp(intLoop).Type = adDBTimeStamp Then
                        varFieldX(intLoop) = "," & rsTmp(intLoop).Name & "=" & GetNullString(rsSox(rsTmp(intLoop).Name).Value, giDateV, giOracle, True)
                    Else
                        varFieldX(intLoop) = "," & rsTmp(intLoop).Name & "=" & Replace(GetNullString(rsSox(rsTmp(intLoop).Name).Value, giStringV, giOracle), ",", "'||chr(44)||'")
                    End If
                End If
            Next
            strField = Mid(Join(varFieldX, ""), 2)
            strValue = "Update " & strTable & " Set " & strField & " " & strWhere
        End If
        gcnGi.Execute strValue, lngAffected
        Call CloseRecordset(rsTmp)
        InsertToOracle2 = True
    Exit Function
ChkErr:
    ErrSub FormName, "InsertToOracle2"
End Function

Public Function ChkHaveField(ByRef strFullField As String, ByVal strField As String, _
    rsSource As ADODB.Recordset) As Boolean
    On Error GoTo ChkErr
    Dim intLoop As Integer
        If strFullField = "" Then
            For intLoop = 0 To rsSource.Fields.Count - 1
                strFullField = strFullField & ",'" & rsSource(intLoop).Name & "'"
            Next
            strFullField = Mid(strFullField, 2)
        Else
            If strField = "" Then Exit Function
            If InStr(1, strFullField, "'" & strField & "'") = 0 Then Exit Function
        End If
        ChkHaveField = True
    Exit Function
ChkErr:
    ChkHaveField = False
End Function

Public Function GetBankATM(strBillNo As String, lngCustId As Long, strBankCode As String, _
    strShouldDate As String, lngShouldAmt As Long) As String
    On Error Resume Next
    Dim rsTmp As New ADODB.Recordset
    Dim clsBankATMNo As Object
    Dim strShDate As String
        If strBankCode = "" Or strBillNo = "" Or lngCustId = 0 Then Exit Function
        If Not GetRS(rsTmp, "Select FileName2,CorpID From " & GetOwner & "CD018 Where CodeNo = " & strBankCode) Then Exit Function
        If Len(rsTmp("FileName2") & "") = 0 Or Len(rsTmp("CorpId") & "") = 0 Then MsgBox "�Ȧ�N�X�ɤ��L�w�q��b��J�ɦW�Φ�����N��", vbCritical, giMsgWarning: Exit Function
        Set clsBankATMNo = CreateObject("csBankATMNo.cls" & rsTmp("FileName2"))
        If Err.Number = 429 Then MsgBox "�Ȧ�N�X�ɤ��w�q����b��J�ɦW : [" & rsTmp("FileName2") & "] �����T,�нT�{!!", vbCritical, giMsgWarning: Exit Function
        strShDate = Format(strShouldDate, "yyyy/mm/dd")
        GetBankATM = clsBankATMNo.GetATMNo(strBillNo, lngCustId, rsTmp("CorpID") & "", strShDate, lngShouldAmt)
End Function

Public Function PrintBillToTEXT(rsPrtSNOTmp As ADODB.Recordset, ByVal strTXTFileName As String, _
    EmailListPath As String) As Boolean
    On Error GoTo ChkErr
    Dim FSO As New FileSystemObject
    Dim MainFile As TextStream      'TEXT File
    Dim DetailFile As TextStream
    Dim SampleFile As TextStream
    Dim strRptPath As String
    Dim strSample As String         'Sample �ɪ����e...
    Dim strRead As String             '�ഫ�᪺���e...
        strRptPath = ReadGICMIS1("RptPath")
        
        If strTXTFileName = "" Then MsgBox "�г]�wEmail List ��X�ɦW(xxx.TXT)", vbCritical, giMsgWarning: Exit Function
        Set DetailFile = FSO.CreateTextFile(EmailListPath & "\" & strTXTFileName)
        
        strTXTFileName = strRptPath & IIf(Right(strRptPath, 1) = "\", "", "\") & strTXTFileName
        Set SampleFile = FSO.OpenTextFile(strTXTFileName)
        
        strSample = SampleFile.ReadAll
        With rsPrtSNOTmp
            .MoveFirst
            Do While Not .EOF
                '���[�ɮ�
                strRead = strSample
                'Set DetailFile = FSO.CreateTextFile(EmailListPath & "\" & rsPrtSNOTmp("A03") & strSubject & ".Txt")
                DetailFile.Write ReplaceTXTtoField(strRead, rsPrtSNOTmp)
                .MoveNext
                DoEvents
            Loop
        End With
        On Error Resume Next
        MainFile.Close
        DetailFile.Close
        PrintBillToTEXT = True
    Exit Function
ChkErr:
    ErrSub FormName, "PrintBillToTEXT"
End Function

Public Function EmailList(rsPrtSNOTmp As ADODB.Recordset, _
        strTXTFileName As String, strSubject As String, _
        EmailListPath As String, Optional blnPrintDetail As Boolean = False) As Boolean
    On Error GoTo ChkErr
    Dim FSO As New FileSystemObject
    
    Dim SampleFile As TextStream
    Dim MainFile As TextStream      'Email List ��
    Dim DetailFile As TextStream    'ActtachFile ��....
    Dim strRptPath As String
    Dim strSample As String         'Sample �ɪ����e...
    Dim strRead As String             '�ഫ�᪺���e...
        strRptPath = ReadGICMIS1("RptPath")
        If strTXTFileName = "" Then MsgBox "�г]�wEmail List ��X�ɦW(xxx.TXT)", vbCritical, giMsgWarning: Exit Function
        strTXTFileName = strRptPath & IIf(Right(strRptPath, 1) = "\", "", "\") & strTXTFileName
        strSubject = Mid(strSubject, 3)
        'Email�榡��
        Set SampleFile = FSO.OpenTextFile(strTXTFileName)
        strSample = SampleFile.ReadAll
        SampleFile.Close
        'EmailList.txt��
        Set MainFile = FSO.CreateTextFile(EmailListPath & "\EmailList.Txt")
        EmailListPath = IIf(EmailListPath = "", "~", EmailListPath)
        '�]�w�D�ɪ�Head
        MainFile.WriteLine "ContentFileName#"
        MainFile.WriteLine "AttachmentFileName#"
        MainFile.WriteLine "ReferenceFileName#"
        MainFile.WriteLine "Subject#" & strSubject
        
        With rsPrtSNOTmp
            .MoveFirst
            Do While Not .EOF
                '���[�ɮ�
                If blnPrintDetail Then
                    strRead = strSample
                    Set DetailFile = FSO.CreateTextFile(EmailListPath & "\" & rsPrtSNOTmp("A03") & strSubject & ".Txt")
                    DetailFile.Write ReplaceTXTtoField(strRead, rsPrtSNOTmp)
                    DetailFile.Close
                End If
                MainFile.WriteLine rsPrtSNOTmp("A03") & "#" & GetRsValue("Select Email From " & GetOwner & "So002 Where CustId = " & rsPrtSNOTmp("A03")) & "#~#" & EmailListPath
                .MoveNext
                DoEvents
            Loop
        End With
        On Error Resume Next
        MainFile.Close
        DetailFile.Close
        EmailList = True
    Exit Function
ChkErr:
    ErrSub "Utility3", "EmailList"
End Function

Public Function ReplaceTXTtoField(strRead As String, ByRef rsSource As ADODB.Recordset) As String
    On Error GoTo ChkErr
    Dim lngStringS As Long
    Dim lngStringE As Long
    Dim strTmp As String
    Dim strField As String
    Dim strValue As String
        strTmp = strRead
        lngStringS = 1
        lngStringE = 1
        Do While True
            lngStringS = InStr(1, strRead, "<")
            If lngStringS = 0 Then Exit Do
            lngStringE = InStr(1, strRead, ">")
            strField = Trim(Mid(strRead, lngStringS + 1, lngStringE - lngStringS - 1))
            If rsSource(strField).Type = adDate Then
                strValue = GetString(GetDT(rsSource(strField) & "", GiDate), lngStringE - lngStringS + 1)
            Else
                strValue = GetString(rsSource(strField) & "", lngStringE - lngStringS + 1)
            End If
            
            strRead = Mid(strRead, 1, lngStringS - 1) & strValue & Mid(strRead, lngStringE + 1)
        Loop
        lngStringS = 1
        lngStringE = 1
        Do While True
            lngStringS = InStr(1, strRead, "{")
            If lngStringS = 0 Then Exit Do
            lngStringE = InStr(1, strRead, "}")
            strField = "*" & Trim(Mid(strRead, lngStringS + 1, lngStringE - lngStringS - 1)) & "*"
            '�����X�r��
            strValue = GetString(ConvertDosBarCode(strField) & "", lngStringE - lngStringS + 1)
            strRead = Mid(strRead, 1, lngStringS - 1) & strValue & Mid(strRead, lngStringE + 1)
        Loop
        ReplaceTXTtoField = strRead
    Exit Function
ChkErr:
    ErrSub "Utility3", "ReplaceTXTtoField"
End Function

Public Function ConvertDosBarCode(BarCode As String) As String
    On Error Resume Next
        ConvertDosBarCode = Replace(BarCode, "*", "��", 1, , vbTextCompare)
        ConvertDosBarCode = Replace(ConvertDosBarCode, "0", "��", 1, , vbTextCompare)
        ConvertDosBarCode = Replace(ConvertDosBarCode, "1", "��", 1, , vbTextCompare)
        ConvertDosBarCode = Replace(ConvertDosBarCode, "2", "��", 1, , vbTextCompare)
        ConvertDosBarCode = Replace(ConvertDosBarCode, "3", "��", 1, , vbTextCompare)
        ConvertDosBarCode = Replace(ConvertDosBarCode, "4", "��", 1, , vbTextCompare)
        ConvertDosBarCode = Replace(ConvertDosBarCode, "5", "��", 1, , vbTextCompare)
        ConvertDosBarCode = Replace(ConvertDosBarCode, "6", "��", 1, , vbTextCompare)
        ConvertDosBarCode = Replace(ConvertDosBarCode, "7", "��", 1, , vbTextCompare)
        ConvertDosBarCode = Replace(ConvertDosBarCode, "8", "��", 1, , vbTextCompare)
        ConvertDosBarCode = Replace(ConvertDosBarCode, "9", "��", 1, , vbTextCompare)
        ConvertDosBarCode = Replace(ConvertDosBarCode, "B", "��", 1, , vbTextCompare)
        ConvertDosBarCode = Replace(ConvertDosBarCode, "T", "��", 1, , vbTextCompare)
        ConvertDosBarCode = Replace(ConvertDosBarCode, "C", "��", 1, , vbTextCompare)
        ConvertDosBarCode = Replace(ConvertDosBarCode, "I", "��", 1, , vbTextCompare)
        ConvertDosBarCode = Replace(ConvertDosBarCode, "M", "��", 1, , vbTextCompare)
        ConvertDosBarCode = Replace(ConvertDosBarCode, "P", "��", 1, , vbTextCompare)
End Function

Public Function giMsgNotSave() As VbMsgBoxResult
    On Error GoTo ChkErr
        'giMsgNotSave = MsgBox(frmSO1100BMDI.stbEditMode.Panels.Item(2) & gmsgNotSave, vbQuestion + vbYesNo, gimsgPrompt)
        giMsgNotSave = MsgBox(gmsgNotSave, vbQuestion + vbYesNo, gimsgPrompt)
        
    Exit Function
ChkErr:
    ErrSub FormName, "giMsgNotSave"
End Function

Public Function CancelSo033(strType As String, _
                        strBillNo As String, strItem As String, _
                        strCitemName As String, _
                        strCompCode As String, _
                        Optional intCancelCode As Integer = -1, Optional strCancelName As String = "", _
                        Optional blnMsgQuestion As Boolean = True) As Boolean
    On Error GoTo ChkErr
        Dim clsAlter As Object
        Dim strMsg As String
        Dim blnChk As Boolean
        
        If Not ChkDataOk(strType, strBillNo, strItem, strCitemName, strCompCode, strMsg) Then Exit Function
        
        If blnMsgQuestion Then If vbNo = MsgBox("�T�{�n�@�o��ڽs��: " & strBillNo & " ����:" & strCitemName & " �ӵ���ƶܡH", vbQuestion + vbYesNo, gimsgPrompt) Then Exit Function
        Set clsAlter = CreateObject("csAlterCharge4.clsAlterCharge")
        clsAlter.uOwnerName = GetOwner
        blnChk = clsAlter.DeleteSo033(gcnGi, strBillNo, strItem, garyGi(1), , intCancelCode, strCancelName)
        If blnChk Then
            MsgBox "�@�o���\�I", vbExclamation, "�T��"
        Else
            MsgBox "�@�o���ѡI�Э��s�ާ@�I", vbInformation, "�T��"
        End If
        CancelSo033 = True
    Exit Function
ChkErr:
    ErrSub FormName, "CancelSo033"
End Function

Public Function GetMaxItem(rsFROM As ADODB.Recordset) As Integer
  On Error GoTo ChkErr
    Dim intRet As Integer
    Dim rsClone As New ADODB.Recordset
    intRet = 1
    Set rsClone = rsFROM.Clone
    rsClone.Filter = rsFROM.Filter
    With rsClone
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                intRet = IIf(Val(.Fields("Item").Value & "") + 1 > intRet, _
                   Val(.Fields("Item").Value & "") + 1, intRet)
                .MoveNext
            Loop
        End If
    End With
    GetMaxItem = intRet
    Call CloseRecordset(rsClone)
  Exit Function
ChkErr:
    ErrSub FormName, "GetMaxItem"
End Function

Public Function GetADString(strDate As String, _
        Optional blnAddTime As Boolean = True, Optional blnMasked As Boolean = False) As String
    On Error GoTo ChkErr
        If strDate = "" Then Exit Function
        If blnAddTime Then
            If blnMasked Then
                GetADString = Format(Trim(CStr(Val(Left(strDate, 2)) + 1911)) & Mid(strDate, 3), "####/##/## ##:##:##")
                'GetADString = Format(Trim(CStr(Val(Left(strDate, 2)) + 1911)) & Mid(strDate, 3), "yyyy/mm/dd hh:mm:ss")
            Else
                GetADString = Trim(CStr(Val(Left(strDate, 2)) + 1911)) & Mid(strDate, 3)
            End If
        Else
            If blnMasked Then
                'GetADString = Format(Trim(CStr(Val(Left(strDate, 2)) + 1911)) & Mid(strDate, 3), "yyyy/mm/dd")
                GetADString = Format(Trim(CStr(Val(Left(strDate, 2)) + 1911)) & Mid(strDate, 3), "####/##/##")
            Else
                GetADString = Trim(CStr(Val(Left(strDate, 2)) + 1911)) & Mid(strDate, 3)
            End If
        End If
    Exit Function
ChkErr:
    ErrSub FormName, "GetADString"
End Function


Public Function GetFieldValue(rs As ADODB.Recordset, _
        strFieldName As String, Optional intType As Integer = 0)
    On Error GoTo ChkErr
    Dim varField As Variant
        'intType 0:Value  1:OriginalValue
        If intType = 0 Then
            varField = rs.Fields(strFieldName).Value
        ElseIf intType = 1 Then
            varField = rs.Fields(strFieldName).OriginalValue
        End If
        
        If rs.Fields(strFieldName).Type = 131 Or rs.Fields(strFieldName).Type = 139 Then
            If intType = 0 Then
                If Len(rs.Fields(strFieldName).Value & "") > 0 Then
                    varField = Val(rs.Fields(strFieldName).Value)
                End If
            ElseIf intType = 1 Then
                If Len(rs.Fields(strFieldName).OriginalValue & "") > 0 Then
                    varField = Val(rs.Fields(strFieldName).OriginalValue)
                End If
            End If
        End If
        GetFieldValue = varField
    Exit Function
ChkErr:
    ErrSub FormName, "GetFiledValue"
End Function

Public Sub Close3Recordset(rs As ADODB.Recordset)
    On Error Resume Next
        
        If rs.State = adStateOpen Then rs.CancelUpdate: rs.Close
        Set rs = Nothing
End Sub

Public Sub CloseConnection(cn As ADODB.Connection)
    On Error Resume Next
        cn.Close
        Set cn = Nothing
End Sub

Public Function ChkSo1100BMDI() As Boolean
    On Error Resume Next
    Dim intLoop As Long
        ChkSo1100BMDI = True
        For intLoop = 0 To Forms.Count - 1
            If UCase(Forms(intLoop).Name) = "FRMSO1100BMDI" Then Exit Function
        Next
        ChkSo1100BMDI = False
End Function

Public Sub SetCombolListIndex(objCombol As ComboBox, strValue As String)
    On Error GoTo ChkErr
    Dim lngLoop As Long
        If objCombol.ListCount > 0 Then
            For lngLoop = 0 To objCombol.ListCount - 1
                objCombol.ListIndex = lngLoop
                If objCombol.Text = strValue Then Exit Sub
            Next
        End If
        objCombol.ListIndex = -1
    Exit Sub
ChkErr:
    ErrSub FormName, "SetCombolListIndex"
End Sub

Public Function GetSystemPara(ByRef rs As ADODB.Recordset, ByVal strCompCode As String, _
    ByVal strServiceType As String, ByVal strTable As String, Optional ByVal strParaField As String = "", _
    Optional blnShowMsg As Boolean = True, Optional blnServiceTypeUseLike As Boolean = False) As Boolean
    On Error GoTo ChkErr
    Dim strSQL As String
    Dim blnFlag As Boolean
        If strParaField = "" Then strParaField = "*"
        'If rs Is Nothing Then Set rs = New ADODB.Recordset
        If blnServiceTypeUseLike Then
            strSQL = " Like '%" & strServiceType & "%'"
        Else
            strSQL = " = '" & strServiceType & "'"
        End If
        '����CompCode ��ServiceType ������
        strTable = GetOwner & strTable
        
        If strCompCode <> "" And strServiceType <> "" Then
            If Not GetRS(rs, "Select " & strParaField & " From " & strTable & " Where CompCode = " & strCompCode & " And ServiceType " & strSQL) Then Exit Function
            blnFlag = Not rs.EOF
        End If
        If Not blnFlag Then
            '�A��CompCode ����
            If strCompCode <> "" Then
                If Not GetRS(rs, "Select " & strParaField & " From " & strTable & " Where CompCode = " & strCompCode) Then Exit Function
                blnFlag = Not rs.EOF
            End If
            '�̫��u�n���o�� �Ѽ��ɪ�
            If Not blnFlag Then
                If Not GetRS(rs, "Select " & strParaField & " From " & strTable) Then Exit Function
                blnFlag = Not rs.EOF
            End If
        End If
        If blnShowMsg And Not blnFlag Then MsgBox "�b�t�ΰѼ���[" & strTable & "] �䤣�� ���q�O�N���� [" & strCompCode & "] ,�B�A�����O��: [" & strServiceType & "]�����,�Ьd��!!", vbCritical, giMsgWarning
        GetSystemPara = True
    Exit Function
ChkErr:
    ErrSub FormName, "GetSystemPara"
End Function

Public Function GetSystemParaItem(strField As String, _
    strCompCode As String, strServiceType As String, strTable As String, _
    Optional blnShowMsg As Boolean = False, Optional blnServiceTypeUseLike As Boolean = False) As Variant
    On Error Resume Next
    Dim rs As New ADODB.Recordset
        If Not GetSystemPara(rs, strCompCode, strServiceType, strTable, strField, False, blnServiceTypeUseLike) Then Exit Function
        GetSystemParaItem = rs(0)
End Function

Public Function ChkHaveService(strCustId As String, strServiceType As String, _
    Optional strShowMsg As Boolean = True) As Boolean
    On Error GoTo ChkErr
    Dim strValue As String
        strValue = GetRsValue("Select ServiceType From " & GetOwner & "So002 Where CustId = " & strCustId & " And ServiceType = '" & strServiceType & "'") & ""
        If Len(strValue) = 0 And strShowMsg Then MsgBox "�ӫȤ�L�A�ȥN����[" & strServiceType & "���A��!!", vbCritical, giMsgWarning
        ChkHaveService = strValue <> ""
    Exit Function
ChkErr:
    ErrSub FormName, "ChkHaveService"
End Function

Public Function GetFloor(ByVal sngValue As Single) As Variant
    Dim var1 As Variant
        var1 = Split(sngValue, ".")
        If sngValue > var1(0) Then
            sngValue = var1(0) + 1
        End If
        GetFloor = sngValue
End Function
'
'Public Function ChkTEXTMaxLengthOk(objForm As Object) As Boolean
'    On Error Resume Next
'    Dim objControl As Object
'        For Each objControl In objForm
'            If TypeOf objControl Is TextBox Or TypeOf objControl Is MaskEdBox Then
'                If StrLen(objControl.Text) > objControl.MaxLength And objControl.Enabled And objControl.Visible And objControl.MaxLength > 0 Then
'                    MsgBox "���׹L�� !!, �������׳̪��� " & objControl.MaxLength
'                    objControl = GetString(objControl, objControl.MaxLength)
'                    objControl.SetFocus
'                    Call objSelectAll(objControl)
'                    Exit Function
'                End If
'            End If
'        Next
'        ChkTEXTMaxLengthOk = True
'End Function

Public Function GetReportFileCnn() As ADODB.Connection
    On Error GoTo ChkErr
    Dim cn As New ADODB.Connection
        cn.Open UCase("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & GetTmpReportFile & ";Persist Security Info=False")
        Set GetReportFileCnn = cn
    Exit Function
ChkErr:
    ErrSub FormName, "GetReportFileCnn"
End Function

Private Function GetTmpReportFile() As String
  On Error GoTo ChkErr
    Dim TmpMDBPath As String
    TmpMDBPath = ReadGICMIS1("TmpMDBPath")
    GetTmpReportFile = IIf(Right(TmpMDBPath, 1) = "\", TmpMDBPath, TmpMDBPath & "\") & "ReportFile.MDB"
  Exit Function
ChkErr:
    Call ErrSub("Sys_Lib", "Function GetTmpReportFile")
End Function

Public Function GetTmpReportFileFldRs(ByRef rsReportFld As ADODB.Recordset, _
    ByRef rsReportStru As ADODB.Recordset, _
    strFldName As String, strStruName As String, cnn As ADODB.Connection) As Boolean
    On Error GoTo ChkErr
        If Not GetRS(rsReportFld, "Select A.FldId, B.TableName, B.FldName, B.FldType, B.Length From " & strFldName & " A,FldList B Where A.FldId = B.FldId", cnn) Then Exit Function
        If Not CreateTmpReportFileStru(rsReportFld, strStruName) Then Exit Function
        If Not GetRS(rsReportStru, "Select * From " & strStruName, cnn, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
        GetTmpReportFileFldRs = True
    Exit Function
ChkErr:
    ErrSub FormName, "GetTmpReportFileFldRs"
End Function

Public Function CreateTmpReportFileStru(rsFld As ADODB.Recordset, strStruName As String) As Boolean
    On Error GoTo ChkErr
    Dim rsClone As New ADODB.Recordset
    Dim strSQL As String, strFieldType As String
    Dim strFieldLength As String
        Set rsClone = rsFld.Clone
        With rsClone
            Do While Not .EOF
                strFieldType = ""
                strFieldLength = ""
                Select Case UCase(.Fields("FldType") & "")
                    Case "NUMBER"
                        strFieldType = "Double"
                    Case "DATE"
                        strFieldType = "Date"
                    Case Else
                        If .Fields("Length") & "" > 255 Then
                            strFieldType = "MEMO"
                        Else
                            strFieldType = "TEXT"
                            strFieldLength = "(255)"
                        End If
                End Select
                strSQL = strSQL & "," & .Fields("FldId") & " " & strFieldType & " " & strFieldLength
                .MoveNext
            Loop
        End With
        On Error Resume Next
        rsFld.ActiveConnection.Execute "Drop Table " & strStruName
        On Error GoTo ChkErr
        rsFld.ActiveConnection.Execute UCase("Create Table " & strStruName & " (" & Mid(strSQL, 2) & ")")
        Call CloseRecordset(rsClone)
        CreateTmpReportFileStru = True
    Exit Function
ChkErr:
    ErrSub FormName, "CreateTmpReportFileStru"
End Function

Private Function GetReportFileField(ByRef strFieldName As String, _
    StrTableName As String, rsFld As ADODB.Recordset) As Boolean
    On Error GoTo ChkErr
    Dim rsClone As New ADODB.Recordset
        strFieldName = ""
        Set rsClone = rsFld.Clone
        With rsClone
            .Filter = "TableName = '" & UCase(StrTableName) & "'"
            Do While Not .EOF
                strFieldName = strFieldName & "," & .Fields("FldName")
                .MoveNext
            Loop
        End With
        If strFieldName <> "" Then strFieldName = Mid(strFieldName, 2)
        Call CloseRecordset(rsClone)
        GetReportFileField = True
    Exit Function
ChkErr:
    ErrSub FormName, "GetReportFileField"
End Function

Public Function GetRelationTableRs(ByRef rs As ADODB.Recordset, rsFld As ADODB.Recordset, _
    ByVal StrTableName As String, ByVal strChoose As String) As Boolean
    On Error GoTo ChkErr
    Dim strField As String
        If strChoose = "" Then MsgBox "�S��Join ����,�е{���]�p�d��!!", vbCritical, giMsgWarning: Exit Function
        If InStr(1, UCase(Trim(Left(strChoose, InStr(1, Trim(strChoose), " ")))), "WHERE") = 0 Then
            strChoose = " Where " & strChoose
        End If
        If Not GetReportFileField(strField, StrTableName, rsFld) Then Exit Function
        If Not GetRS(rs, "Select " & strField & " From " & GetOwner & StrTableName & strChoose) Then Exit Function
        GetRelationTableRs = True
    Exit Function
ChkErr:
    ErrSub FormName, "GetRelationTableRs"
End Function

Public Function ChkDupData(ByVal strWhere As String, ByVal StrTableName As String) As Boolean
    On Error GoTo ChkErr
    Dim lngCount As Long
        If strWhere = "" Then Exit Function
        If InStr(1, UCase(strWhere), "WHERE") = 0 Then strWhere = " Where " & strWhere
        StrTableName = GetOwner & StrTableName
        lngCount = GetRsValue("Select Count(*) From " & StrTableName & " " & strWhere)
        If lngCount > 0 Then ChkDupData = True
    Exit Function
ChkErr:
    ErrSub FormName, "ChkDupData"
End Function

Public Function ShowFullTimeStr(ByVal lngTime As Long) As String
    On Error Resume Next
    Dim strTimeStr As String
        If lngTime > 3600 Then
            strTimeStr = GetString(lngTime \ 3600, 2, giRight) & "�p��"
            lngTime = lngTime Mod 3600
        End If
        If lngTime > 60 Then
            strTimeStr = strTimeStr & GetString((lngTime \ 60), 2, giRight) & "��"
            lngTime = lngTime Mod 60
        End If
        strTimeStr = strTimeStr & GetString(lngTime, 2, giRight) & "��"
        ShowFullTimeStr = strTimeStr
End Function

Public Function MackCABReport(strSourceFile As String, strTarget As String) As Boolean
    On Error GoTo ChkErr
    Dim MCobj As Object
        Set MCobj = CreateObject("CSmc.MCclass")
        '���Y�ɮ�
        If Not MCobj.MakeCab(strSourceFile, strTarget) Then Exit Function
        MackCABReport = True
    Exit Function
ChkErr:
    ErrSub FormName, "MackCABReport"
End Function

Public Function FTPDownLoadReport(strIP As String, strUserName As String, _
    strPassword As String, strFTPFileName As String, strPutFileName As String, _
    Optional blnShowMessage As Boolean = False) As Boolean
    Dim FTPobj As Object
    Dim MCobj As Object
        'FTP �U�����
        Set FTPobj = CreateObject("CSftp.FtpClass")
        If Not FTPobj.FtpGet(strIP, strUserName, strPassword, strFTPFileName, strPutFileName, blnShowMessage) Then Exit Function
        FTPDownLoadReport = True
    Exit Function
ChkErr:
    ErrSub FormName, "FTPDownLoadReport"
End Function

Public Function ExpnCABReport(strSourceFile As String, strTarget As String) As Boolean
    On Error GoTo ChkErr
    Dim MCobj As Object
        '�����Y�ɮ�
        Set MCobj = CreateObject("CSmc.MCclass")
        If Not MCobj.ExpnCAB(strSourceFile, strTarget) Then Exit Function
        ExpnCABReport = True
    Exit Function
ChkErr:
    ErrSub FormName, "ExpnCABReport"
End Function

Public Function GetAutoCreateRptRS(ByRef rs As ADODB.Recordset) As Boolean
    On Error Resume Next
        If rs.State = adStateOpen Then rs.CancelUpdate
    On Error GoTo ChkErr
        With rs
            If .State = adStateOpen Then .Close
            .CursorLocation = adUseClient
            .CursorType = adOpenKeyset
            .LockType = adLockOptimistic
            
            .Fields.Append "TableName", adVarChar, 30
            .Fields.Append "SQLQuery", adVarChar, 2000
            .Open
        End With
        GetAutoCreateRptRS = True
    Exit Function
ChkErr:
    ErrSub FormName, "GetAutoCreateRptRS"
End Function

Public Function AutoCreateReport(cnReport As ADODB.Connection, _
    ByVal strFmtID As String, rsQuery As ADODB.Recordset, cnOutPutReport As ADODB.Connection, _
    Optional ByVal strMainTable As String = "", Optional ByVal blnCreateStru As Boolean = True, _
    Optional ByRef lngAffected As Long, Optional ByVal blnCreateView As Boolean, _
    Optional ByRef strTmpViewName As String) As Boolean
    On Error GoTo ChkErr
    Dim rsFmtlist As New ADODB.Recordset
    Dim rsJoin As New ADODB.Recordset
    Dim rsFld As New ADODB.Recordset
    Dim rsFld2 As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim rsData2 As New ADODB.Recordset
    Dim intLoop As Integer
    Dim strJoin As String
    Dim strTable As String
    Dim strField As String
    Dim strWhere As String
    Dim strExecuteStr As String
    Dim strInsertField As String
    Dim strInsertValue As String
    Dim strQueryTable As String
    Dim strSQL As String
    
        If Not GetRS(rsFmtlist, "Select FmtId,Type,MainTable,FmtFileName,JoinName,ReportName From FmtList Where FmtId= '" & strFmtID & "'", cnReport) Then Exit Function
        If rsFmtlist.RecordCount = 0 Then MsgBox "�L�Ӫ��N��,�Ьd��!!", vbCritical, gimsgPrompt: Exit Function
        If Not GetRS(rsJoin, "Select * From " & rsFmtlist("JoinName") & " Where FmtID = '" & strFmtID & "'", cnReport) Then Exit Function
        
        If Not GetRS(rsFld, "Select 1 as Type , A.* From " & rsFmtlist("FmtFileName") & " A Where FmtID = '" & strFmtID & "' And TableName = '" & rsFmtlist("MainTable") & "'" & _
                " Union All Select * From (Select 2 as Type , A.* From " & rsFmtlist("FmtFileName") & " A Where FmtID = '" & strFmtID & "' And TableName <> '" & rsFmtlist("MainTable") & "' Order By TableName)", cnReport) Then Exit Function
        If Not GetReportFieldStr(rsFld, strField, strTable) Then Exit Function
        
        strField = Mid(strField, 2)
        If blnCreateStru Then
            If Not GetRS(rsTmp, "Select " & strField & " From " & Mid(strTable, 2) & " Where 1=0 ") Then Exit Function
            If Not CreateMDBTable(rsTmp, rsFmtlist("ReportName"), cnOutPutReport) Then Exit Function
        End If
        '����Join
        If strMainTable = "" Then strMainTable = rsJoin("MainTable") & ""
        strMainTable = UCase(strMainTable)
        
        If rsQuery.RecordCount > 0 Then rsQuery.MoveFirst
        Do While Not rsQuery.EOF
            ' Join ����һ�Table
            If InStr(1, strTable, UCase(rsQuery("TableName"))) = 0 Then strTable = strTable & "," & GetOwner & UCase(rsQuery("TableName")) & " " & UCase(rsQuery("TableName"))
            ' Join �������
            strJoin = strJoin & " And " & rsQuery("SQLQuery")
            strQueryTable = strQueryTable & "," & GetOwner & UCase(rsQuery("TableName")) & " " & UCase(rsQuery("TableName"))
            rsQuery.MoveNext
        Loop
        
        If rsFmtlist.Fields("Type") = 1 Then
            Do While Not rsJoin.EOF
                ' Join ������Table
                If InStr(1, strTable, UCase(strMainTable)) = 0 Then strTable = strTable & "," & GetOwner & UCase(strMainTable) & " " & UCase(rsJoin("MainTable") & "")
                If InStr(1, strTable, UCase(rsJoin("SecondTable"))) = 0 Then strTable = strTable & "," & GetOwner & UCase(rsJoin("SecondTable")) & " " & UCase(rsJoin("SecondTable"))
                For intLoop = 1 To 5
                    If rsJoin("MainJoinFld" & intLoop) & "" <> "" Then
                        ' Join ���һݱ���
                        strJoin = strJoin & " And " & strMainTable & "." & rsJoin("MainJoinFld" & intLoop) & " = " & rsJoin("SecondTable") & "." & rsJoin("SecondJoinFld" & intLoop)
                    End If
                Next
                rsJoin.MoveNext
            Loop
            If strJoin <> "" Then strWhere = " Where " & Mid(strJoin, 5)
            strSQL = "Select " & strField & " From " & Mid(strTable, 2) & strWhere
            If blnCreateView Then
                strTmpViewName = GetTmpViewName
                gcnGi.Execute "Create View " & GetOwner & strTmpViewName & " AS (" & strSQL & ")"
            Else
                If Not GetRS(rsData, strSQL) Then Exit Function
                With rsData
                    Do While Not .EOF
                        strExecuteStr = GetTmpMDBExecuteStr(rsData, rsFmtlist("ReportName"))
                        strExecuteStr = Replace(Left(UCase(strExecuteStr), InStr(1, UCase(strExecuteStr), " VALUES ")), ".", "") & Mid(strExecuteStr, InStr(1, UCase(strExecuteStr), " VALUES "))
                        cnOutPutReport.Execute strExecuteStr
                        DoEvents
                        .MoveNext
                    Loop
                    lngAffected = .RecordCount
                End With
            End If
        Else
            
            rsFld.Filter = "Type = 1"
            strField = ""
            strTable = ""
            If Not GetReportFieldStr(rsFld, strField) Then Exit Function
            If Not GetRS(rsFld2, "Select A.TableName From " & rsFmtlist("FmtFileName") & " A Where FmtID = '" & strFmtID & "' And instr(1,'" & Replace(strQueryTable, ",", "") & "' ,TableName) > 0  Group By TableName", cnReport) Then Exit Function
            With rsFld2
                If .RecordCount > 0 Then .MoveFirst
                Do While Not .EOF
                    rsFld.Filter = "TableName = '" & rsFld2("TableName") & "'"
                    rsJoin.Filter = "SecondTable = '" & rsFld2("TableName") & "'"
                    If rsJoin.RecordCount > 0 Then
                        For intLoop = 1 To 5
                            If rsJoin("MainJoinFld" & intLoop) & "" <> "" Then
                                ' Join ���һݱ���
                                strJoin = strJoin & " And " & strMainTable & "." & rsJoin("MainJoinFld" & intLoop) & " = " & rsJoin("SecondTable") & "." & rsJoin("SecondJoinFld" & intLoop)
                            End If
                        Next
                    End If
                    Do While Not rsFld.EOF
                        strField = strField & "," & rsFld("TableName") & "." & rsFld("FldName") & " " & rsFld("TableName") & rsFld("FldName")
                        rsFld.MoveNext
                    Loop
                    .MoveNext
                Loop
            End With
            
            If strJoin <> "" Then strWhere = " Where " & Mid(strJoin, 5)
            strTable = GetOwner & strMainTable & " " & strMainTable & strQueryTable
            If Not GetRS(rsData, "Select " & Mid(strField, 2) & " From " & GetOwner & strTable & strWhere) Then Exit Function
            
            If Not GetRS(rsFld2, "Select A.TableName From " & rsFmtlist("FmtFileName") & " A Where FmtID = '" & strFmtID & "' And instr(1,'" & Replace(strQueryTable, ",", "") & "' ,TableName) = 0 And TableName <> '" & strMainTable & "'  Group By TableName", cnReport) Then Exit Function
            
            Do While Not rsData.EOF
                strInsertField = ""
                strInsertValue = ""
                For intLoop = 0 To rsData.Fields.Count - 1
                    strInsertField = strInsertField & "," & rsData.Fields(intLoop).Name
                    strInsertValue = strInsertValue & "," & GetValueExecuteStr(rsData, rsData.Fields(intLoop).Name)
                Next
                With rsFld2
                    If .RecordCount > 0 Then .MoveFirst
                    Do While Not .EOF
                        rsFld.Filter = "TableName = '" & .Fields("TableName") & "'"
                        strField = ""
                        Do While Not rsFld.EOF
                            strField = strField & "," & rsFld("TableName") & "." & rsFld("FldName")
                            rsFld.MoveNext
                        Loop
                        rsJoin.Filter = "SecondTable = '" & .Fields("TableName") & "' And MainTable = '" & strMainTable & "'"
                        strJoin = ""
                        If rsJoin.RecordCount > 0 Then
                            For intLoop = 1 To 5
                                If rsJoin("MainJoinFld" & intLoop) & "" <> "" Then
                                    ' Join ���һݱ���
                                    strJoin = strJoin & " And " & rsJoin("SecondTable") & "." & rsJoin("SecondJoinFld" & intLoop) & " = " & GetValueExecuteStr(rsData, rsData(rsJoin("MainTable") & rsJoin("MainJoinFld" & intLoop)).Name)
                                End If
                            Next
                        End If
                        
                        strWhere = ""
                        If strJoin <> "" Then strWhere = " Where " & Mid(strJoin, 5)
                        If Not GetRS(rsData2, "Select " & Mid(strField, 2) & " From " & GetOwner & rsFld2("TableName") & " " & rsFld2("TableName") & strWhere) Then Exit Function
                        If rsData2.RecordCount > 0 Then
                            For intLoop = 0 To rsData2.Fields.Count - 1
                                strInsertValue = strInsertValue & "," & GetValueExecuteStr(rsData2, rsData2.Fields(intLoop).Name)
                            Next
                            strInsertField = strInsertField & Replace(strField, ".", "")
                        End If
                        
                        
                        .MoveNext
                    Loop
                End With
                strExecuteStr = "Insert Into " & rsFmtlist("ReportName") & " (" & Mid(strInsertField, 2) & ") Values (" & Mid(strInsertValue, 2) & ")"
                cnOutPutReport.Execute strExecuteStr
                rsData.MoveNext
            Loop
            'if not getrs (rsdata,"Select " & strfield & " From " & strmaintable & " Where " & rsquery
            
        End If
        AutoCreateReport = True
    Exit Function
ChkErr:
    ErrSub FormName, "AutoCreateReport"
End Function

Private Function GetReportFieldStr(rsFld As ADODB.Recordset, ByRef strField As String, _
    Optional ByRef strTable As String) As Boolean
    On Error GoTo ChkErr
        If rsFld.RecordCount > 0 Then rsFld.MoveFirst
        Do While Not rsFld.EOF
            '���һ�Table
            If InStr(1, strTable, UCase(rsFld("TableName"))) = 0 Then strTable = strTable & "," & GetOwner & UCase(rsFld("TableName")) & " " & UCase(rsFld("TableName"))
            
            strField = strField & "," & rsFld("TableName") & "." & rsFld("FldName") & " " & rsFld("TableName") & rsFld("FldName")
            rsFld.MoveNext
        Loop
        GetReportFieldStr = True
    Exit Function
ChkErr:
    ErrSub FormName, "GetReportFieldStr"
End Function

Private Function GetValueExecuteStr(rs As ADODB.Recordset, strFieldName As String) As String
    On Error GoTo ChkErr
        Select Case chkFieldDataType(rs, strFieldName)
            Case giDataNumeric, giDataBoolean
                GetValueExecuteStr = GetNullString(rs.Fields(strFieldName).Value, giLongV, giAccessDb)
            Case giDataString
                GetValueExecuteStr = GetNullString(rs.Fields(strFieldName).Value, giStringV, giAccessDb)
            Case giDataMemo
                GetValueExecuteStr = GetNullString(rs.Fields(strFieldName).Value, giStringV, giAccessDb)
            Case giDataDate
                GetValueExecuteStr = GetNullString(rs.Fields(strFieldName).Value, giStringV, giAccessDb, True)
            Case giDataNone
                GetValueExecuteStr = GetNullString(rs.Fields(strFieldName).Value, giStringV, giAccessDb)
        End Select
    Exit Function
ChkErr:
    ErrSub FormName, "GetValueExecuteStr"
End Function

Public Function chkFieldDataType(rs As ADODB.Recordset, strFieldName As String) As giFieldDataType
    On Error GoTo ChkErr
        Select Case rs.Fields(strFieldName).Type
               Case adBigInt, adVarNumeric, adNumeric, adDecimal, adSingle, adDouble, adInteger '"Double"
                    chkFieldDataType = giDataNumeric
               Case adBoolean '"Logical"
                    chkFieldDataType = giDataBoolean
               Case adBSTR, adChar, adChapter, adVarChar, adVariant, adVarWChar, adLongVarChar '"Text("
                    If rs.Fields(strFieldName).DefinedSize > 512 Then
                        chkFieldDataType = giDataMemo
                    ElseIf rs.Fields(strFieldName).DefinedSize * 2 > 255 Then
                        chkFieldDataType = giDataMemo
                    Else
                        chkFieldDataType = giDataString
                    End If
               Case adDate, adDBDate, adDBTimeStamp, adDBTime '"Date"
                    chkFieldDataType = giDataDate
               Case Else
                    chkFieldDataType = giDataNone
        End Select
    Exit Function
ChkErr:
    ErrSub FormName, "chkFieldDataType"
End Function

Public Function chkFieldDataType2(rs As ADODB.Recordset, strFieldName As String) As giVType
    On Error GoTo ChkErr
        
        Select Case rs.Fields(strFieldName).Type
               Case adBigInt, adVarNumeric, adNumeric, adDecimal, adSingle, adDouble, adInteger '"Double"
                    chkFieldDataType2 = giLongV
               Case adBoolean '"Logical"
                    chkFieldDataType2 = giLongV
               Case adBSTR, adChar, adChapter, adVarChar, adVariant, adVarWChar, adLongVarChar '"Text("
                    chkFieldDataType2 = giStringV
               Case adDate, adDBDate, adDBTimeStamp, adDBTime '"Date"
                    chkFieldDataType2 = giDateV
               Case Else
                    chkFieldDataType2 = giStringV
        End Select
    Exit Function
ChkErr:
    ErrSub FormName, "chkFieldDataType2"
End Function

Public Function SetFmtID(cnRptMDB As ADODB.Connection, strCompCode As String, rsRpt As ADODB.Recordset, _
    cboFmtID As Object) As Boolean
    On Error GoTo ChkErr
    Dim strWhere As String
    Dim strSOCompCode As String
        If cnRptMDB Is Nothing Then Exit Function
        If cnRptMDB.State = adStateClosed Then Exit Function
        strSOCompCode = GetRsValue("Select SOCompCode From " & GetOwner & "CD039 Where CodeNo = " & strCompCode) & ""
        If strSOCompCode <> "" Then strWhere = " Where CompCode = '" & strSOCompCode & "'"
        
        'If strCompCode <> "" Then strWhere = " Where CompCode = " & strCompCode
        If Not GetRS(rsRpt, "SELECT * FROM FmtList" & strWhere, cnRptMDB, adUseClient, adOpenKeyset, adLockOptimistic) Then
            ErrHandle "�b�}�� [���@����](FmtList) ����Ʈɵo�Ϳ��~: " & Err.Description & " �o�ӵ{���Y�N�����C", False, 2, "Form_Load"
            
            Exit Function
        End If
        cboFmtID.Clear
        If rsRpt.RecordCount > 0 Then
            rsRpt.MoveFirst
            Do While Not rsRpt.EOF
                If (rsRpt("FmtID").Value & "") <> "" And rsRpt("Type").Value = 1 Then
                    cboFmtID.AddItem rsRpt("FmtID").Value
                End If
                rsRpt.MoveNext
            Loop
        End If
        Call CloseRecordset(rsRpt)
        SetFmtID = True
    Exit Function
ChkErr:
    ErrSub FormName, "SetFmtID"
End Function

Public Function SetDefaultCbo(strCompCode As String, cboType As Object, _
    cboFmtID As Object, cboPrinterName As Object, Optional strServiceType As String = "") As Boolean
    On Error GoTo ChkErr
    Dim rsTmp As New ADODB.Recordset
    Dim intIndex As Integer
        ' �ھڳ�O�ܿ�X�Ѽ���(SO045)����������X�榡�N���ΦC����W��, �]�������ѼƤ����
        If cboType.ListIndex < 0 Then Exit Function
        intIndex = cboType.ListIndex + 1
        If intIndex > 5 Then intIndex = 1
        If Not GetSystemPara(rsTmp, strCompCode, strServiceType, "SO045", "BillTab" & intIndex & ",BillPrinter") Then
            Exit Function
        End If
        If rsTmp.EOF Then
           MsgBox "�Х��]�w[��X�Ѽ���(SO045)", vbCritical, "�T��"
           Exit Function
        Else
           If SetCboText(cboFmtID, rsTmp("BillTab" & intIndex).Value & "", "�䤣����w�����O��榡��") = False Then Exit Function
           If SetCboText(cboPrinterName, GetPrinterName(4, strServiceType), "�䤣����w���L���") = False Then Exit Function
        End If
        SetDefaultCbo = True
    Exit Function
ChkErr:
    ErrSub FormName, "SetDefaultCbo"
End Function

Private Function SetCboText(cbo As ComboBox, strValue As String, strMsg As String) As Boolean
    On Error GoTo ChkErr
    Dim lngX As Long
        SetCboText = True
        For lngX = 0 To cbo.ListCount - 1
            If UCase(cbo.List(lngX)) = UCase(strValue) Then
                cbo.ListIndex = lngX
                Exit Function
            End If
        Next
        MsgBox strMsg & " [" & strValue & "]�I", vbExclamation, "ĵ�i"
        SetCboText = True
        If cbo.ListCount > 0 Then
            cbo.ListIndex = 0
        Else
            'Call errHandle("�ϥΪ��t�ΰѼƦ��~�A���{���Y�N�����I", False, 2, "SetCboText")
            SetCboText = False
            Exit Function
        End If
    Exit Function
ChkErr:
    ErrSub FormName, "SetCboText"
End Function

Public Function GetMaskAccountNo(strAccountNo As String) As String
    On Error Resume Next
    Dim strTmpNo As String
    Dim intLength As Integer
        If strAccountNo = "" Then Exit Function
        intLength = Len(strAccountNo)
        If intLength >= 5 Then strTmpNo = Left(strAccountNo, 5)
        If intLength >= 6 Then strTmpNo = strTmpNo & "X"
        If intLength >= 7 Then strTmpNo = strTmpNo & "X"
        If intLength >= 8 Then strTmpNo = strTmpNo & "X"
        If intLength >= 9 Then strTmpNo = strTmpNo & Mid(strAccountNo, 9, 1)
        If intLength >= 10 Then strTmpNo = strTmpNo & Mid(strAccountNo, 10, 1)
        If intLength >= 11 Then strTmpNo = strTmpNo & "X"
        If intLength >= 12 Then strTmpNo = strTmpNo & "X"
        If intLength >= 13 Then strTmpNo = strTmpNo & Mid(strAccountNo, 13)
        GetMaskAccountNo = strTmpNo
End Function

Public Function GetMaskPersonID(strID As String) As String
    On Error Resume Next
        If strID = "" Then Exit Function
        GetMaskPersonID = Left(strID, 3) & "XXXX" & Right(strID, 3)
End Function





Public Function GetTmpMdbConnection(Optional ByVal strMDBName As String = "TMP0000.MDB") As Connection
    On Error GoTo ChkErr
    Dim strTMDBFileName As String
    Dim cn As New ADODB.Connection
        strTMDBFileName = ReadGICMIS1("TmpMDBPath")
        strTMDBFileName = IIf(Right(strTMDBFileName, 1) = "\", strTMDBFileName, strTMDBFileName & "\") & strMDBName
        
        cn.CursorLocation = adUseClient
        cn.Open UCase("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strTMDBFileName & ";Persist Security Info=False")
        Set GetTmpMdbConnection = cn
    Exit Function
ChkErr:
    Call ErrSub(FormName, "GetTmpMdbConnection")
End Function
'
'Public Function ExporttoExcel(Optional ByVal strSQL As String, Optional rs As ADODB.Recordset, _
'    Optional strFileName As String, Optional strSheetName As String, Optional blnAddHead As Boolean, _
'    Optional cnn As ADODB.Connection) As Boolean
'    On Error GoTo ChkErr
'    Dim objXls As Object
'    Dim rsXls As New ADODB.Recordset
'    Dim lngPos As Long
'    Dim intLoop As Integer
'        If cnn Is Nothing Then Set cnn = gcnGi
'        Set objXls = CreateObject("Excel.Application")
'        If strSQL <> "" Then
'            If Not GetRS(rsXls, strSQL, cnn) Then Exit Function
'        Else
'            Set rsXls = rs.Clone
'        End If
'        With objXls
'            .Workbooks.Add
'            .DisplayAlerts = False
'            Do While .Sheets.Count > 1
'                .Sheets(2).Delete
'            Loop
'            .ActiveSheet.Cells.Select
'            .Selection.NumberFormatLocal = "@"
'            .Selection.Font.Size = 9
'            .Rows.Item(1).RowHeight = 15
'            .Rows.RowHeight = 14
'            rsXls.MoveFirst
'            While Not rsXls.EOF
'                 lngPos = rsXls.AbsolutePosition + 1
'                 For intLoop = 0 To rsXls.Fields.Count - 1
'                    If blnAddHead And rsXls.AbsolutePosition = 1 Then .ActiveSheet.Cells(1, intLoop + 1).Value = Trim(rsXls(intLoop).Name & "")
'                    .ActiveSheet.Cells(lngPos, intLoop + 1).Value = Trim(rsXls(intLoop) & "")
'                Next
'                 rsXls.MoveNext
'             Wend
'            .ActiveSheet.Range(.Cells(1, 1), .Cells(rsXls.RecordCount + 1, 4)).Select
'            .Selection.EntireColumn.AutoFit
'            If strSheetName <> "" Then .Worksheets(1).Name = strSheetName
'            If strFileName <> "" Then
'                On Error Resume Next
'                Kill strFileName
'                .ActiveWorkbook.SaveAs strFileName
'            End If
'            .Visible = True
'        End With
'        Set objXls = Nothing
'        rsXls.Close
'        Set rsXls = Nothing
'        ExporttoExcel = True
'    Exit Function
'ChkErr:
'    ErrSub FormName, "ExporttoExcel"
'    Set objXls = Nothing
'End Function

Public Function ChargeField(Optional strAlias As String) As String
    On Error GoTo ChkErr
    Dim rs As New ADODB.Recordset
    Dim strField As String
    Dim intX As Integer
        If Not GetRS(rs, "Select * From " & GetOwner & "SO033 SO033 Where SO033.BillNo = '' And SO033.Item = -1", gcnGi) Then Exit Function
        For intX = 0 To rs.Fields.Count - 1
            If strAlias <> "" Then
                strField = strField & "," & strAlias & "." & rs.Fields(intX).Name
            Else
                strField = strField & "," & rs.Fields(intX).Name
            End If
        Next
        ChargeField = Mid(strField, 2)
        Call CloseRecordset(rs)
    Exit Function
ChkErr:
    ErrSub FormName, "ChargeField"
End Function

Public Function ChkCloseDate(strCloseDate As String, _
    strCompCode As String, strServiceType As String) As Boolean
    On Error GoTo ChkErr
    Dim intDayCut As Integer
    Dim strTranDate As String
    Dim intPara6 As Integer
    Dim strWhere As String
        intDayCut = Val(GetSystemParaItem("DayCut", strCompCode, "", "SO041") & "")
        '2007/10/12 �令�]�[�A�ȧO�P�_
        If strServiceType <> "" Then strWhere = " And ServiceType = '" & strServiceType & "'"
        strTranDate = GetRsValue("Select tranDate from " & GetOwner & "So062 Where CompCode= " & strCompCode & strWhere & " And Type=1 Order By TranDate Desc") & ""
        
        '���O�Ѽ��ɡ]SO043)�A�����O��n������<Para6>�A�ѫ���ϥΡI
        intPara6 = Val(GetSystemParaItem("Para6", strCompCode, strServiceType, "SO043") & "")
    
        If strCloseDate = "" Then
            If vbNo = MsgBox("�����O�_���ŭȡI", vbExclamation + vbYesNo, "�T���I") Then Exit Function
        Else
            If Not IsDate(strCloseDate) Then
                MsgBox "������X�k�I", vbExclamation, "�T���I"
                Exit Function
            End If
            
            If CDate(strCloseDate) > RightDate Then MsgBox "������W�L���Ѥ���I", vbExclamation, "�T���I": Exit Function
            
            If (RightDate - CDate(strCloseDate)) > intPara6 Then
                MsgBox "������w�W�L�t�γ]�w���w�������I", vbExclamation, "�T���I":  Exit Function
            End If
            If CDate(strCloseDate) <= IIf(strTranDate <> "", CDate(strTranDate), 0) Then
                If intDayCut = 1 And CDate(strCloseDate) = IIf(strTranDate <> "", CDate(strTranDate), 0) Then ChkCloseDate = True: Exit Function
                MsgBox strTranDate & "���e�w���L�鵲�A���i�ϥΤ��e����J�b", vbExclamation, "�T���I"
                Exit Function
            End If
        End If
        ChkCloseDate = True
    Exit Function
ChkErr:
    ErrSub FormName, "ChkCloseDate"
End Function

Public Function GetFaciSNo(strSeqNo As String) As String
    On Error Resume Next
        If strSeqNo = "" Then Exit Function
        GetFaciSNo = GetRsValue("Select FaciSNo From " & GetOwner & "SO004 Where SeqNo = '" & strSeqNo & "'") & ""
End Function

Public Function GetByHouse(strCitemCode As String) As Integer
    On Error Resume Next
        If strCitemCode = "" Then Exit Function
        GetByHouse = GetRsValue("Select ByHouse From " & GetOwner & "CD019 Where CodeNo = " & strCitemCode)
End Function

'�Ĥ@�ӶǱ��󻡩�,�ĤG�ӶǱ��󤺮e,�ĤT�ӶǬO�_�[����(;) 93/11/24
'Ex: GetSelectConditionStr("�ꦬ���","93/11/24~93/11/31",False)
'���G��: "�ꦬ���:93/11/24~93/11/31;"
Public Function GetSelectConditionStr(ByVal strCaption As String, ByVal strSELECT As String, _
    Optional ByVal blnSemicolon As Boolean = True) As String
    Dim strSelectConditionStr As String
    If strSELECT <> "" Then
        strSelectConditionStr = strCaption & ":" & strSELECT
        If blnSemicolon Then strSelectConditionStr = strSelectConditionStr & ";"
    End If
    GetSelectConditionStr = strSelectConditionStr
End Function

Public Function CreateTmpView(ByVal strSQL As String, ByRef strTmpViewName As String, _
    Optional blnTable As Boolean = False) As Boolean
    On Error GoTo ChkErr
        strTmpViewName = GetTmpViewName
        If blnTable Then
            gcnGi.Execute "Create Table " & GetOwner & strTmpViewName & " as " & strSQL
        Else
            gcnGi.Execute "Create View " & GetOwner & strTmpViewName & " as " & strSQL
        End If
        CreateTmpView = True
    Exit Function
ChkErr:
    ErrSub FormName, "CreateTmpView"
End Function

Public Function GetDefaultFaciSNo(ByVal str004Filter As String, ByVal strRefNo As String, _
    ByRef strFaciSeqNo As String, ByRef strFaciSNo As String, _
    Optional blnIsbyRsCopy As Boolean, Optional rsFaci As ADODB.Recordset = Nothing, _
    Optional strServiceType As String = "") As Boolean
    On Error GoTo ChkErr
    Dim rsData As New ADODB.Recordset
    Dim strWhere As String
    Dim strFaciWhere As String
    Dim strSQL As String
    Dim intCount As Integer
    Dim strField As String
    Dim strSeqField As String
    Dim strRefFaciCode As String
        If strServiceType = "" Then strServiceType = gServiceType
        strFaciWhere = " PRDate is Null and A.FaciCode In (Select CodeNo From " & GetOwner & "CD022 Where "
        
        Select Case strServiceType
            Case "I"
                strWhere = " And " & strFaciWhere & " RefNo In (2,5,7,8" & IIf(strRefNo = "", "", "," & strRefNo)
            Case "D"
                strWhere = " And " & strFaciWhere & " RefNo In (3" & IIf(strRefNo = "", "", "," & strRefNo)
            Case "P"
                strWhere = " And " & strFaciWhere & " RefNo In (6" & IIf(strRefNo = "", "", "," & strRefNo)
            Case Else
                If strRefNo <> "" Then strWhere = " And " & strFaciWhere & " RefNo In (" & strRefNo
        End Select
        
        If strWhere <> "" Then strWhere = strWhere & "))"
        strSQL = "Select A.SeqNo,a.FaciSNo From " & GetOwner & "SO004 A Where A.CustId=" & gCustId & _
            " And A.ServiceType='" & strServiceType & "' and a.CompCode=" & gCompCode & strWhere
            
        If str004Filter <> "" Then strSQL = strSQL & " And " & str004Filter
        
        If Not GetRS(rsData, strSQL, , adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
        intCount = rsData.RecordCount
        If intCount > 0 Then
            strFaciSeqNo = rsData("SeqNo") & ""
            strFaciSNo = rsData("FaciSNo") & ""
        End If
        Call CloseRecordset(rsData)
        If intCount > 1 Then GoTo lExit
        
        If strRefNo <> "" Then strRefFaciCode = "(" & gcnGi.Execute("Select CodeNo From " & GetOwner & "CD022 Where RefNo In (" & strRefNo & ")").GetString(, , "", "),(")
        If Not rsFaci Is Nothing Then
            Set rsData = rsFaci.Clone
            If blnIsbyRsCopy Then
                strField = "FaciCode"
                strSeqField = "SeqNo"
            Else
                strField = "FaciNo"
                strSeqField = "FaciSeqNo"
            End If
            If rsData.RecordCount > 0 Then rsData.MoveFirst
            Do While Not rsData.EOF
                If InStr(1, strRefFaciCode, "(" & rsData(strField) & ")") > 0 Or strRefNo = "" Then
                    strFaciSeqNo = rsData(strSeqField) & ""
                    If blnIsbyRsCopy Then strFaciSNo = rsData("FaciSNo") & ""
                    intCount = intCount + 1
                End If
                rsData.MoveNext
            Loop
        End If
        If intCount > 1 Then GoTo lExit
        GetDefaultFaciSNo = True
        Exit Function
lExit:
        strFaciSeqNo = ""
        strFaciSNo = ""
        Call CloseRecordset(rsData)
    Exit Function
ChkErr:
    ErrSub FormName, "GetDefaultFaciSNo"
End Function

'�ˮָ�ƪ��@�P��,
Public Function ChkDataCons(rsOld As ADODB.Recordset, ByVal strSQL As String, _
    Optional blnShowMsg As Boolean = True, Optional ByRef strErrorMsg As String = "") As Boolean
    Dim rsDB As New ADODB.Recordset
    Dim intLoop As Integer
    Dim lngBR As Long
    Dim cn As New ADODB.Connection
        cn.Open gcnGi.ConnectionString
        On Error Resume Next
        cn.BeginTrans
        cn.Execute strSQL & " For Update Nowait"
        If Err.Number <> 0 Then
            cn.RollbackTrans
            Err.Number = 0
            gcnGi.Execute strSQL & " For Update Nowait"
            If Err.Number <> 0 Then
                MsgBox "�ӵ���ƥ��Q��L�H���ʤ�,�Ьd��!!", vbCritical, giMsgWarning
                cn.RollbackTrans
                GoTo lClose
            End If
        End If
        cn.RollbackTrans
        On Error GoTo ChkErr
        '��l��Ƭ�EOF�L�k���,�����P�_
        If Not rsOld.EOF Then
            If Not GetRS(rsDB, strSQL) Then Exit Function
            '��Ʈw���s�b,�����P�_
            If rsDB.RecordCount > 0 Then
                For intLoop = 0 To rsDB.Fields.Count - 1
                    '�P�_��줺�e���P�h��ܸ�Ʈw�w�g�Q����
                    If rsDB(intLoop) & "" <> rsOld(rsDB(intLoop).Name).OriginalValue & "" Then
                        strErrorMsg = "�ӵ���Ƥw�g�Q��L�H���ʹL,�Ьd��!!"
                        If blnShowMsg Then MsgBox strErrorMsg, vbCritical, giMsgWarning
                        Exit Function
                    End If
                Next
            End If
        End If
        ChkDataCons = True
        On Error Resume Next
lClose:
        cn.Close
        Set cn = Nothing
    Exit Function
ChkErr:
    ErrSub FormName, "ChkDataCons"
End Function

Public Function GetFaciStatus(ByVal rs004 As ADODB.Recordset, Optional ByVal Value) As String
    On Error Resume Next
    Dim FaciWipStatus As String
    Dim rs As New ADODB.Recordset
    Dim strPRField As String
        With rs004
            If .State = adStateClosed Then Exit Function
            If .EOF Or .BOF Then Exit Function
'            1.���`=>InstDate is not null and PRDate is null
'            2.����=>InstDate is not null and PRDate is null And FaciStatusCode = 3
'            3.���/�����^=>PRDate is not null and GetDate is null
'            4.���/���^=>PRDate is not null and GetDate is not null
'            5.�L:InstDate is null and PrDate is null
            If GetSystemParaItem("FaciAuthority", gCompCode, gServiceType, "SO042") = 2 Then
                strPRField = "GetDate"
            Else
                strPRField = "PRDate"
            End If
            If .Fields("InstDate") & "" <> "" And .Fields(strPRField) & "" = "" Then
                If .Fields("FaciStatusCode") & "" = 3 Then
                    FaciWipStatus = "����"
                Else
                    If StrToDate(.Fields("StopTime") & "", True) > StrToDate(.Fields("ReInstTime") & "", True) Then
                        FaciWipStatus = "����"
                    Else
                        FaciWipStatus = "���`"
                    End If
                End If
            ElseIf .Fields(strPRField) & "" <> "" Then
                FaciWipStatus = "���"
                If .Fields("GetDate") & "" = "" Then
                    FaciWipStatus = FaciWipStatus & "/�����^"
                End If
            Else
                FaciWipStatus = "�L"
            End If
            If .Fields("GetDate") & "" <> "" Then
                FaciWipStatus = FaciWipStatus & "/���^"
            End If
        End With
        GetFaciStatus = FaciWipStatus
End Function

Public Function GetFaciCommandStatus(ByVal rs004 As ADODB.Recordset) As String
    On Error Resume Next
    Dim FaciWipStatus As String
    Dim rs As New ADODB.Recordset
        With rs004
            If .State = adStateClosed Then Exit Function
            If .EOF Or .BOF Then Exit Function
'            �}�q���A
'            1.���}��=>CMOpenDate is null and CMCloseDate is null
'            2.�}��=>CMOpenDate > CMCloseDate or CMOpendate is not null and CMCloseDate is null
'            3.����=>CMOpenDate < CMCloseDate
'            4.�}��/���v=>(CMOpenDate > CMCloseDate or CMOpendate is not null and CMCloseDate is null) and
'            (disableaccount> enableAccount or (disableaccount is not null and enableAccount is null)
'            5.����/���v=>(CMOpenDate < CMCloseDate) and
'            (disableaccount> enableAccount or (disableaccount is not null and enableAccount is null)
            
            If StrToDate(.Fields("CMOpenDate") & "", True, True) > StrToDate(.Fields("CMCloseDate") & "", True, True) Then
                FaciWipStatus = "�}��"
            ElseIf StrToDate(.Fields("CMOpenDate") & "", True, True) < StrToDate(.Fields("CMCloseDate") & "", True, True) Then
                FaciWipStatus = "����"
            ElseIf .Fields("CMOpenDate") & "" = "" And .Fields("CMCloseDate") & "" = "" Then
                FaciWipStatus = "���}��"
            Else
                FaciWipStatus = "�L"
            End If
            If StrToDate(.Fields("disableaccount") & "", True, True) > StrToDate(.Fields("enableAccount") & "", True, True) Then
                FaciWipStatus = FaciWipStatus & "/���v"
            End If
        End With
        GetFaciCommandStatus = FaciWipStatus
End Function

Public Function GetFaciWipStatus(ByVal rs004 As ADODB.Recordset) As String
    On Error Resume Next
    Dim strStatus As String
        strStatus = gcnGi.Execute("Select Kind From " & GetOwner & "SO004D Where CustId = " & rs004("CustId") & " And SeqNo = '" & rs004("SeqNo") & "' And SignDate is Null Order by UpdTime Desc").GetString(, , , "/") & ""
        GetFaciWipStatus = Mid(strStatus, 1, Len(strStatus) - 1)
End Function

Public Function InsertChargeField(rsSO033 As ADODB.Recordset, _
    ByVal lngCustId As Long, ByVal lngCompCode As Long, ByVal strServiceType As String, _
        ByVal strBillNo As String, ByVal lngItem As Long, _
        ByVal strCitemCode As String, ByVal strShouldDate As String, _
        Optional ByVal strClctEn As String, Optional ByVal strClctName As String, _
        Optional ByVal strFaciSeqNo As String = "", Optional ByVal strFaciSNo As String = "", _
        Optional ByVal blnAddNew As Boolean = True, Optional strCMCode As String = "") As Boolean
    On Error GoTo ChkErr
    Dim obj As Object
        If rsSO033 Is Nothing Then
            If Not GetRS(rsSO033, "Select * From " & GetOwner & "SO033 Where BillNo ='' And Item = -1", , adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
            Set rsSO033.ActiveConnection = Nothing
        End If
        Set obj = CreateObject("csAlterCharge4.clsInsertNewCharge")
        With obj
            .uOwnerName = GetOwner
            .uErrPath = ReadGICMIS1("ErrLogPath")
            Set .uconn = gcnGi
            If Not .InsertChargeField(rsSO033, lngCustId, lngCompCode, strServiceType, strBillNo, lngItem, strCitemCode, strShouldDate, strClctEn, strClctName, _
                strFaciSeqNo, strFaciSNo, blnAddNew) Then Exit Function
            If strCMCode <> "" Then
                rsSO033("CMCode") = strCMCode
                rsSO033("CMName") = GetRsValue("Select Description From " & GetOwner & "CD031 Where CodeNo = " & strCMCode) & ""
                rsSO033("CreateEn") = garyGi(0)
                rsSO033.Update
            End If
        End With
        InsertChargeField = True
    Exit Function
ChkErr:
    ErrSub FormName, "InsertChargeField"
End Function

Public Function CreditCardApprove(ByVal lngCompCode As Long, ByVal strCardBillNo As String, _
        ByVal lngCardCode As Long, ByVal strAccountNo As String, _
        ByVal strCardExpDate As String, ByVal lngAMOUNT As Long, ByVal strRealDate As String, _
        ByVal strCMCode As String, ByVal strCMName As String, _
        ByVal strPTCode As String, ByVal strPTName As String, _
        ByVal strClctEn As String, ByVal strClctName As String, _
        ByVal strUpdTime As String, ByVal strUpdEn As String, ByVal strEntryEn As String, _
        ByVal strAuthorizeCode As String, ByRef strErrorCode As String, ByRef strErrorMsg As String) As Boolean
    On Error GoTo ChkErr
    Dim objPayment As Object
    Dim lngDepositFlag As Long
        strErrorCode = ""
        strErrorMsg = ""
        lngDepositFlag = GetSystemParaItem("Para32", gCompCode, gServiceType, "SO043")
        Set objPayment = CreateObject("csPaymentClass.Approve")
        With objPayment
            Set .objConn = gcnGi
            .uOwnerName = garyGi(16)
            CreditCardApprove = .Approve(lngCompCode, strCardBillNo, lngCardCode, strAccountNo, strCardExpDate, lngAMOUNT, strRealDate, _
                strCMCode, strCMName, strPTCode, strPTName, strClctEn, strClctName, strUpdTime, strUpdEn, strEntryEn, lngDepositFlag, strAuthorizeCode) = 0
            strErrorCode = .getErrorCode
            strErrorMsg = .getErrorMsg
        End With
        Set objPayment = Nothing
    Exit Function
ChkErr:
    ErrSub FormName, "CreditCardApprove"
End Function

Public Function CreditCardApproveReversal(lngCompCode As Long, strCardBillNo As String, _
        strUpdTime As String, strUpdEn As String, _
        ByRef strErrorCode As String, ByRef strErrorMsg As String) As Boolean
    On Error GoTo ChkErr
    Dim objPayment As Object
        strErrorCode = ""
        strErrorMsg = ""
        Set objPayment = CreateObject("csPaymentClass.ApproveReversal")
        With objPayment
            Set .objConn = gcnGi
            .uOwnerName = garyGi(16)
            CreditCardApproveReversal = .ApproveReversal(lngCompCode, strCardBillNo, strUpdTime, strUpdEn) = 0
            strErrorCode = .getErrorCode
            strErrorMsg = .getErrorMsg
        End With
        Set objPayment = Nothing
    Exit Function
ChkErr:
    ErrSub FormName, "CreditCardApproveReversal"
End Function

Public Function CreditCardDeposit(lngCompCode As Long, strCardBillNo As String, _
    strRealDate As String, strUpdTime As String, strUpdEn As String, _
        ByRef strErrorCode As String, ByRef strErrorMsg As String) As Boolean
    On Error GoTo ChkErr
    Dim objPayment As Object
        strErrorCode = ""
        strErrorMsg = ""
        Set objPayment = CreateObject("csPaymentClass.Deposit")
        With objPayment
            Set .objConn = gcnGi
            .uOwnerName = garyGi(16)
            CreditCardDeposit = .Deposit(lngCompCode, strCardBillNo, strRealDate, strUpdTime, strUpdEn) = 0
            strErrorCode = .getErrorCode
            strErrorMsg = .getErrorMsg
        End With
        Set objPayment = Nothing
    Exit Function
ChkErr:
    ErrSub FormName, "CreditCardDeposit"
End Function

Public Function CreditCardDepositReversal(lngCompCode As Long, strCardBillNo As String, _
        strUpdTime As String, strUpdEn As String, ByVal strEntryEn As String, _
        ByRef strErrorCode As String, ByRef strErrorMsg As String) As Boolean
    On Error GoTo ChkErr
    Dim objPayment As Object
    Dim lngDepositFlag As Long
        strErrorCode = ""
        strErrorMsg = ""
        Set objPayment = CreateObject("csPaymentClass.DepositReversal")
        lngDepositFlag = GetSystemParaItem("Para32", gCompCode, gServiceType, "SO043")
        With objPayment
            Set .objConn = gcnGi
            .uOwnerName = garyGi(16)
            CreditCardDepositReversal = .DepositReversal(lngCompCode, strCardBillNo, strUpdTime, strUpdEn, strEntryEn, lngDepositFlag) = 0
            strErrorCode = .getErrorCode
            strErrorMsg = .getErrorMsg
        End With
        Set objPayment = Nothing
    Exit Function
ChkErr:
    ErrSub FormName, "CreditCardDepositReversal"
End Function

Public Function GetCombolIndex(objCombol As Object, strValue As String) As Integer
    Dim intCount As Integer
        With objCombol
            For intCount = 0 To objCombol.ListCount - 1
                If objCombol.List(intCount) = strValue Then objCombol.ListIndex = intCount: Exit For
            Next
        End With
End Function

Public Function CopyRecordset(rsSource As ADODB.Recordset, _
    rsTarget As ADODB.Recordset, Optional ByVal blnCompareField As Boolean, _
    Optional ByVal blnAll As Boolean = True, Optional ByVal blnAddNew As Boolean = True) As Boolean
    On Error GoTo ChkErr
    Dim intLoop As Integer
    Dim lngBR As Long
        If rsSource Is Nothing Or rsTarget Is Nothing Then CopyRecordset = True: Exit Function
        If rsSource.State = adStateClosed Or rsTarget.State = adStateClosed Then CopyRecordset = True: Exit Function

        With rsSource
            If blnAll Then
                lngBR = .AbsolutePosition
                If .RecordCount > 0 Then
                    .MoveFirst
                End If
            End If
            Do While Not .EOF
                If blnAddNew Then rsTarget.AddNew
                For intLoop = 0 To .Fields.Count - 1
                    If Not blnCompareField Then On Error Resume Next
                    rsTarget.Fields(.Fields(intLoop).Name) = .Fields(intLoop)
                Next
                rsTarget.Update
                If Not blnAll Then Exit Do
                .MoveNext
            Loop
            If lngBR > 0 Then .AbsolutePosition = lngBR
        End With
        CopyRecordset = True
    Exit Function
ChkErr:
    ErrSub FormName, "CopyRecordset"
End Function

Public Function DeleteRSRecord(rs As ADODB.Recordset) As Boolean
    On Error GoTo ChkErr
        With rs
            If .EOF Or .BOF Then Exit Function
            If MsgBox("�T�w�n�R���o����ơH", vbQuestion + vbYesNo, "�T�{") = vbYes Then
                .Delete
                .MoveNext
                If .RecordCount > 0 Then
                    .MoveLast
                End If
            End If
        End With
        DeleteRSRecord = True
    Exit Function
ChkErr:
    ErrSub FormName, "DeleteRecordset"
End Function

Public Function FrameCanCutData(blnFlag As Boolean, fraData As Object, _
    objForm As Object, Optional ByVal blnIncLabel As Boolean = False)
    On Error Resume Next
    Dim obj As Object
        For Each obj In objForm
            If Not blnFlag Then
                If Not TypeOf obj Is Label Or blnIncLabel Then
                    If obj.Container.Name = fraData.Name Then
                        On Error Resume Next
                        If TypeOf obj Is GiNumber Then obj.Enabled = False
                        If TypeOf obj Is GiList Then obj.Enabled = False
    
                        obj.Locked = True
                        If Err.Number <> 0 Then obj.Enabled = False
                        
                    End If
                End If
            End If
        Next
        fraData.Enabled = True
End Function

Public Function ChkRealDateOK(ByVal gdaRealDate As Object, _
    ByRef Cancel As Boolean, ByVal strCompCode As String, ByVal strServiceType As String, _
    Optional ByVal XstrTranDate As String = "", _
    Optional ByVal intDayCut As Integer = -1, _
    Optional ByVal XintPara6 As Integer = -1) As Boolean
    On Error GoTo ChkErr
        Cancel = False
        If intDayCut = -1 Then intDayCut = GetSystemParaItem("DayCut", strCompCode, "", "SO041", , True)
        If XintPara6 = -1 Then XintPara6 = GetSystemParaItem("Para6", strCompCode, strServiceType, "SO043", , True)
        If XstrTranDate = "" Then XstrTranDate = GetRsValue("Select tranDate from " & GetOwner & "So062 Where CompCode= " & gCompCode & " And Type=1") & ""

        If gdaRealDate.Text = "" Then Exit Function
        '�Y�j�󤵤�A�hĵ�i�G"������W�L���Ѥ��"
        If Not IsDate(Format(gdaRealDate.GetValue, "####/##/##")) Then
            MsgBox "������~�I", , "�T���I"
            Cancel = True
            Exit Function
        End If
        If CDate(gdaRealDate.Text) > RightDate Then MsgBox "������W�L���Ѥ��": Cancel = True: Exit Function
        
        '�Y�]����-������- >�@��Para6>)�A�hĵ�i�G"������w�W�L�t�γ]���w������"
        If DateDiff("d", CDate(gdaRealDate.Text), RightDate) > XintPara6 Then
            MsgBox "������w�W�L�t�γ]�w���w������", vbExclamation, "�T���I"
            Cancel = True
            Exit Function
        End If
        
        '�Y�p��<TransDate>�A�h���~�G"<TransDate>���e�w���L�鵲�A���i�ϥΤ��e����J�b",Focus���d�b�����I
        If DateDiff("d", CDate(XstrTranDate), CDate(gdaRealDate.Text)) <= 0 Then
            If intDayCut = 1 And DateDiff("d", CDate(XstrTranDate), CDate(gdaRealDate.Text)) = 0 Then Exit Function
            MsgBox XstrTranDate & "���e�w���L�鵲�A���i�ϥΤ��e����J�b", vbExclamation, "�T���I"
            Cancel = True
            Exit Function
        End If
    Exit Function
ChkErr:
    ErrSub FormName, "ChkRealDateOK"
End Function

Public Function AutoDefineRs(rs As ADODB.Recordset, rsSource As ADODB.Recordset, ByVal strField As String, _
    Optional ByVal blnChoose As Boolean, Optional ByVal blnHave As Boolean, Optional colAddField As Collection = Nothing) As Boolean
    On Error GoTo ChkErr
    Dim varX As Variant
    Dim varX2 As Variant
    Dim varX3 As Variant
    Dim intLoop As Integer
    
        Set rs = New ADODB.Recordset
        If strField = "*" Then
            strField = ""
            For intLoop = 0 To rsSource.Fields.Count - 1
                strField = strField & "," & rsSource.Fields.Item(intLoop).Name
            Next
            strField = Mid(strField, 2)
        End If
        varX = Split(Replace(strField, " ", ""), ",")
        
        With rs
            If .State = adStateOpen Then .Close
            .CursorLocation = adUseClient
            .CursorType = adOpenKeyset
            .LockType = adLockOptimistic
            If blnChoose Then
                .Fields.Append "CHOOSE", adBSTR, 1, adFldMayBeNull
            End If
            For intLoop = 0 To UBound(varX)
                Select Case rsSource(varX(intLoop)).Type
                       Case adBigInt, adVarNumeric, adNumeric, adDecimal, adSingle, adDouble, adInteger '"Double"
                            .Fields.Append UCase(varX(intLoop)), adDouble, , rsSource(varX(intLoop)).Attributes
                       Case Else
                            .Fields.Append UCase(varX(intLoop)), rsSource(varX(intLoop)).Type, IIf(rsSource(varX(intLoop)).DefinedSize <= 0, 1, rsSource(varX(intLoop)).DefinedSize), rsSource(varX(intLoop)).Attributes
                End Select
            Next
            If Not colAddField Is Nothing Then
                varX = Split(colAddField(1), ",")
                varX2 = Split(colAddField(2), ",")
                varX3 = Split(colAddField(3), ",")
                For intLoop = 0 To UBound(varX)
                    .Fields.Append UCase(varX(intLoop) & ""), varX2(intLoop), varX3(intLoop)
                Next
            End If
            If blnHave Then
                .Fields.Append "HAVECHOOSE", adBSTR, 1, adFldMayBeNull
            End If
            .Open
        End With
        AutoDefineRs = True
    Exit Function
ChkErr:
    ErrSub FormName, "AutoDefineRs"
End Function

Public Function GetTableField(StrTableName As String, _
    Optional ByVal strOwner As String, Optional ByVal strAlias As String) As String
    On Error GoTo ChkErr
    Dim rs As New ADODB.Recordset
    Dim strField As String
    Dim intX As Integer
        If strOwner = "" Then strOwner = GetOwner
        If Not GetRS(rs, "Select * From " & strOwner & StrTableName & " " & strAlias & " Where RowId = ''", gcnGi) Then Exit Function
        For intX = 0 To rs.Fields.Count - 1
            If strAlias <> "" Then
                strField = strField & "," & strAlias & "." & rs.Fields(intX).Name
            Else
                strField = strField & "," & rs.Fields(intX).Name
            End If
        Next
        GetTableField = Mid(strField, 2)
        Call CloseRecordset(rs)
    Exit Function
ChkErr:
    ErrSub FormName, "GetTableField"
End Function

Private Sub GetAgencyPosition()
    Dim rs As New ADODB.Recordset
    Dim intLoop As Integer
    Dim intLength As Integer
    Dim strFieldName As String
    Dim intDefinesize As Long
    GetRS rs, "Select * from TmpSo033", GetTmpMdbCn
    intLength = 1
    For intLoop = 0 To rs.Fields.Count - 1
'        If Right(rs.Fields(intLoop).Name, 2) <> Right(strFieldName, 2) And InStr(rs.Fields(intLoop).Name, "_") Then
'            Debug.Print ""
'        End If
        Select Case rs.Fields(intLoop).Name
            Case "Message_1_01", "Message_2_01"
                intDefinesize = 300
            Case "Message_02", "Message_03"
                intDefinesize = 750
            Case Else
            intDefinesize = rs.Fields(intLoop).DefinedSize
        End Select
        'Debug.Print GetString(rs.Fields(intLoop).Name, 30) & GetString(intDefinesize, 5) & " " & GetString(intLength, 5) & "-" & GetString(intLength + intDefinesize, 5) - 1
        intLength = intLength + intDefinesize
        
        strFieldName = rs.Fields(intLoop).Name
        
    Next
End Sub

Public Function GetDefaultSTType(Optional strServiceType As String = "") As String
    On Error Resume Next
        If strServiceType = "" Then strServiceType = gServiceType
        GetDefaultSTType = Val(GetRsValue("Select AdjustFlag from " & GetOwner & "cd046 where CodeNo = '" & strServiceType & "'") & "")

End Function

Private Function GetMIDvalue(ByVal strReportName As String, ByRef MIDvalue As String) As Boolean
    On Error GoTo ChkErr
    Dim strMidValue As String
        If MIDvalue <> "" Then GetMIDvalue = True: Exit Function
        strMidValue = GetRsValue("Select Mid From " & GetOwner & "SO029 Where Mid = '" & Left(strReportName, 6) & "'") & ""
        MIDvalue = strMidValue
        GetMIDvalue = True
    Exit Function
ChkErr:
    ErrSub "PrintModule", "GetMIDvalue"
End Function

Public Function GetSo3311XTmpMDB(cnn As ADODB.Connection) As Boolean
    On Error GoTo ChkErr
    Dim rsTmp As New ADODB.Recordset
        If Not GetRS(rsTmp, "Select RowId,A.* From " & GetOwner & "SO074 A Where BillNo = '' And Item = -1") Then Exit Function
        If Not CreateMDBTable(rsTmp, "SO3311", cnn) Then Exit Function
        
        Call CloseRecordset(rsTmp)
        GetSo3311XTmpMDB = True
    Exit Function
ChkErr:
    ErrSub FormName, "GetSo3311XTmpMDB"
End Function

Public Function GetFalseSNo(strSNOType As String) As String
    On Error GoTo ChkErr
        GetFalseSNo = strSNOType & Left(garyGi(0), 7) & Format(Time, "HHMMSS") & Right(Format(Timer, "0.0"), 1)
    Exit Function
ChkErr:
    ErrSub FormName, "GetFalseSNo"
End Function

Public Function GetSO003Value(ByVal strField As String, ByVal lngCustId As Long, _
    ByVal lngCompCode As Long, ByVal strServiceType As String, _
    Optional ByVal strFaciSeqNo As String, Optional ByVal strCitemCode As String) As String
    On Error GoTo ChkErr
    Dim strWhere As String
        strWhere = "CustId = " & lngCustId
        strWhere = strWhere & " And CompCode = " & lngCompCode
        strWhere = strWhere & " And ServiceType = '" & strServiceType & "'"
        If strFaciSeqNo <> "" Then subAnd2 strWhere, "FaciSeqNo = '" & strFaciSeqNo & "'"
        If strCitemCode <> "" Then subAnd2 strWhere, "CitemCode = " & strCitemCode
        GetSO003Value = GetRsValue("Select " & strField & " From " & GetOwner & "SO003 Where " & strWhere) & ""
    Exit Function
ChkErr:
    ErrSub FormName, "GetSO003Value"
End Function

Public Function GetSO003Rs(ByRef rs As ADODB.Recordset, ByVal strField As String, ByVal lngCustId As Long, _
    ByVal lngCompCode As Long, ByVal strServiceType As String, _
    Optional ByVal strFaciSeqNo As String, Optional ByVal strCitemCode As String) As Boolean
    On Error GoTo ChkErr
    Dim strWhere As String
        strWhere = "CustId = " & lngCustId
        strWhere = strWhere & " And CompCode = " & lngCompCode
        strWhere = strWhere & " And ServiceType = '" & strServiceType & "'"
        If strFaciSeqNo <> "" Then subAnd2 strWhere, "FaciSeqNo = '" & strFaciSeqNo & "'"
        If strCitemCode <> "" Then subAnd2 strWhere, "CitemCode = " & strCitemCode
        If Not GetRS(rs, "Select " & strField & " From " & GetOwner & "SO003 Where " & strWhere) Then Exit Function
        GetSO003Rs = True
    Exit Function
ChkErr:
    ErrSub FormName, "GetSO003Rs"
End Function

Public Function GetAutoStopDate(lngPeriod As Long, strStartDate As String, _
    ByRef strStopDate As String, ByRef strClctDate As String) As Boolean
    On Error GoTo ChkErr
    Dim rsSO043 As New ADODB.Recordset
        If strStartDate <> "" Then
            If Not GetSystemPara(rsSO043, gCompCode, gServiceType, "SO043", "Para3,Para5") Then Exit Function
            
            If rsSO043.Fields("para3") = 1 Or lngPeriod = 0 Then
                strStopDate = CStr(DateAdd("m", lngPeriod, CDate(strStartDate)))
            Else
                strStopDate = CStr(DateAdd("m", lngPeriod, CDate(strStartDate)) - 1)
            End If
        
            strClctDate = CStr(CDate(strStopDate) + Val(rsSO043.Fields("para5")))
        End If
        Call CloseRecordset(rsSO043)
        GetAutoStopDate = True
    Exit Function
ChkErr:
    Call ErrSub(FormName, "GetAutoStopDate")
End Function

Public Function GetKindSO004D(strKind As String, rsSO004D As ADODB.Recordset, _
    rsSource As ADODB.Recordset) As Boolean
    On Error GoTo ChkErr
        If Not GetRS(rsSO004D, "Select RowId,SO004D.* From " & GetOwner & "SO004D Where SeqNo = ''", , adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
        Set rsSO004D.ActiveConnection = Nothing
        If Not CopyRecordset(rsSource, rsSO004D, , False) Then Exit Function
        rsSO004D("Kind") = strKind
        rsSO004D.Update
        GetKindSO004D = True
    Exit Function
ChkErr:
    ErrSub FormName, "GetKindSO004D"
End Function



Public Function InsertToMemRS(rs As ADODB.Recordset, rsSource As ADODB.Recordset, _
    Optional ByVal strFilter As String = "", Optional blnAutoDefineRS As Boolean, _
    Optional ByVal strField As String, _
    Optional ByVal blnChoose As Boolean, Optional ByVal blnHave As Boolean, _
    Optional ByVal strSort As String, Optional colAddField As Collection = Nothing) As Boolean
    On Error GoTo ChkErr
    Dim rsClone As New ADODB.Recordset
    Dim intLoop As Integer
        If rsSource Is Nothing Then InsertToMemRS = True: Exit Function
        If rsSource.State = adStateClosed Then InsertToMemRS = True: Exit Function
        If blnAutoDefineRS Then If Not AutoDefineRs(rs, rsSource, strField, blnChoose, blnHave, colAddField) Then Exit Function
        Set rsClone = rsSource.Clone
        With rsClone
            .Filter = strFilter
            .Sort = strSort
            If .RecordCount > 0 Then .MoveFirst
            Do While Not .EOF
                rs.AddNew
                For intLoop = 0 To rs.Fields.Count - 1
'                    If rs(intloop).Name <> "CHOOSE" And rs(intloop).Name <> "HAVECHOOSE" Then
                        On Error Resume Next
                        rs(intLoop) = .Fields(rs(intLoop).Name)
'                    End If
                Next
                rs.Update
                .MoveNext
            Loop
        End With
        Call CloseRecordset(rsClone)
        InsertToMemRS = True
    Exit Function
ChkErr:
    ErrSub FormName, "InsertToMemRS"
End Function

Public Function DefineAddFieldCollection(ByVal strFieldString As String, _
    ByVal strFieldTypeString As String, ByVal strFieldLengthString As String) As Collection
    Dim tmpCol As New Collection
        tmpCol.Add strFieldString
        tmpCol.Add strFieldTypeString
        tmpCol.Add strFieldLengthString
        Set DefineAddFieldCollection = tmpCol
End Function

Public Function GetServiceRefNo(strServiceType As String) As String
    On Error Resume Next
    Select Case strServiceType
        Case "C"
            GetServiceRefNo = "1,3"
        Case "I"
            GetServiceRefNo = "2,5,7,8"
        Case "D"
            GetServiceRefNo = "3"
        Case "P"
            GetServiceRefNo = "6"
    End Select
End Function

'2007/12/20 Jacky 3163 Liga
Public Function ChkManualNoOk(ByVal strManualNo As String, Optional ByVal strBillNo As String, _
    Optional strServiceType As String, Optional strCompCode As String) As Boolean
    On Error GoTo ChkErr
    Dim YUseManual As Integer
    Dim strBillSQL As String
    Dim rsTmp As New ADODB.Recordset
        If strServiceType = "" Then strServiceType = gServiceType
        If strCompCode = "" Then strCompCode = gCompCode
        
        YUseManual = Val(GetSystemParaItem("UseManual", strCompCode, strServiceType, "SO043") & "")
        If strManualNo <> "" And YUseManual = 1 Then
            If Not GetRS(rsTmp, "Select Status From " & GetOwner & "SO127 Where PaperNum = '" & strManualNo & "' And CompCode = " & strCompCode) Then Exit Function
            If rsTmp.EOF Then
                MsgBox "�L����}�渹,�Ьd��!!", vbExclamation, gimsgPrompt
                Exit Function
            Else
                If rsTmp("Status") = 0 Then
                    MsgBox "����}�渹�w�@�o���i�ϥ�,�Ьd��!!", vbExclamation, gimsgPrompt
                    Exit Function
                End If
            End If
            
            If strBillNo = "" Then
                strBillSQL = " And BillNo is not null"
            Else
                strBillSQL = " And BillNo <> '" & strBillNo & "'"
            End If
            
            If GetRsValue("Select Count(*) From " & GetOwner & "SO127 Where PaperNum = '" & strManualNo & "' And CompCode = " & strCompCode & strBillSQL) > 0 Then
                MsgBox "�Ӥ�}��w�����O���,�Ьd��!!", vbExclamation, gimsgPrompt
                Exit Function
            End If
        End If
        Call CloseRecordset(rsTmp)
        ChkManualNoOk = True
    Exit Function
ChkErr:
    ErrSub FormName, "ChkManualNoOk"
End Function

Public Function DateAddCS(Interval As String, Number, strDate)
    DateAddCS = DateAdd(Interval, Number, CDate(strDate) + 1) - 1
End Function

Public Function ChkRowIsLock(ByVal strSQL As String, Optional cnn As ADODB.Connection) As Boolean
    On Error Resume Next
        If cnn Is Nothing Then
            Set cnn = New ADODB.Connection
            cnn.Open gcnGi.ConnectionString
        End If
        cnn.Execute strSQL & " For Update Nowait"
        If Err.Number <> 0 Then ChkRowIsLock = True: Exit Function
        
End Function

Public Function RoundUP(ByVal vData As Double) As Double
    Dim Value As Double
        Value = Int(vData)
        If Value <> vData Then Value = CLng(vData) + 1
        RoundUP = Value
End Function

Public Function GetRealStopDate(ByVal rs As ADODB.Recordset, ByVal strStartDate As String, _
    ByVal lngPeriod As Long, ByVal strCompCode As String, ByVal strServiceType As String, _
    ByRef intAddDay As Integer) As Date
    Dim intPFlag1 As Integer
    Dim intPara3 As Integer
    Dim strStopDate As Date
        
        intPFlag1 = GetPFlag1(strStartDate, strCompCode, strServiceType, Val(rs("PFlag1") & ""), rs("PFlagDate") & "", intPara3)
        intAddDay = IIf(intPara3 = 1, 0, 1)
        strStopDate = DateAddCS("m", lngPeriod, strStartDate)
        If intPFlag1 > 0 Then
            If Format(strStartDate, "DD") <> "01" Then
                If intPFlag1 = 1 Then    '�p�����
                    strStopDate = GetLastDate(DateAddCS("m", lngPeriod - 1, strStartDate))
                Else
                    strStopDate = GetLastDate(DateAddCS("m", lngPeriod, strStartDate))
                End If
            Else
                strStopDate = GetLastDate(DateAddCS("m", lngPeriod - 1, strStartDate))
            End If
        End If
        If intPFlag1 = 0 Then strStopDate = strStopDate - intAddDay
        GetRealStopDate = strStopDate
        
End Function

Private Function GetPFlag1(ByVal strStartDate As String, _
    ByVal strCompCode As String, ByVal strServiceType As String, _
    ByVal intPFlag1 As Integer, ByVal strPFlagDate As String, _
    ByRef intPara3 As Integer) As Integer
        If intPFlag1 = 1 Or (Format(strStartDate, "DD") <= Format(strPFlagDate, "00") And intPFlag1 = 3) Then
            GetPFlag1 = 1
        ElseIf intPFlag1 = 2 Or (Format(strStartDate, "DD") > Format(strPFlagDate, "00") And intPFlag1 = 3) Then
            GetPFlag1 = 2
        End If
        intPara3 = GetSystemParaItem("Para3", strCompCode, strServiceType, "SO043")
End Function

Public Function GetRealPeriod(ByVal strStartDate As String, ByVal strStopDate As String, _
    Optional ByRef blnRemain As Boolean) As Long
    On Error Resume Next
    Dim lngPeriod As Long
        blnRemain = False
        
        lngPeriod = DateDiff("m", strStartDate, strStopDate)
        If DateAddCS("m", lngPeriod, strStartDate) - 1 <> CDate(strStopDate) Then
            blnRemain = True
        End If
        If DateAddCS("m", lngPeriod, strStartDate) < CDate(strStopDate) Then
            lngPeriod = lngPeriod + 1
        End If
        GetRealPeriod = lngPeriod
End Function

Public Function ChkHave2Discount(ByVal strCitemCode As String) As Boolean
    On Error Resume Next
    Dim rs019 As New ADODB.Recordset
        If Val(GetSystemParaItem("SecondDiscount", gCompCode, "", "SO041") & "") = 0 Then Exit Function
        If Not GetRS(rs019, "Select PayCHDependCode,PayCHDependName,SecendDiscount,RefNo From " & GetOwner & "CD019 Where CodeNo = " & strCitemCode) Then Exit Function
        '����2�h�u�f���W�D���O���� �� ��������2�h�u�f
        If Val(rs019("SecendDiscount") & "") > 0 Then
            ChkHave2Discount = True
        End If
End Function

Public Function Chk2Discount(ByVal strCitemCode As String) As Boolean
    Dim rs019 As New ADODB.Recordset
        If Val(GetSystemParaItem("SecondDiscount", gCompCode, "", "SO041") & "") = 0 Then Exit Function
        If Not GetRS(rs019, "Select RefNo From " & GetOwner & "CD019 Where CodeNo = " & strCitemCode) Then Exit Function
        '����2�h�u�f
        If Val(rs019("RefNo") & "") = 19 Then
            Chk2Discount = True
            Exit Function
        End If
        If Not GetRS(rs019, "Select RefNo From " & GetOwner & "CD019 Where ReturnCode = " & strCitemCode) Then Exit Function
        '���ɦ���2�h�u�f
        If Not rs019.EOF Then
            If Val(rs019("RefNo") & "") = 19 Then
                Chk2Discount = True
                Exit Function
            End If
        End If
End Function

Public Function Chk3Discount(ByVal strCitemCode As String) As Boolean
    On Error Resume Next
    Dim rs019 As New ADODB.Recordset
        If Val(GetSystemParaItem("ThirdDiscount", gCompCode, "", "SO041") & "") = 0 Then Exit Function
        If Not GetRS(rs019, "Select RefNo From " & GetOwner & "CD019 Where CodeNo = " & strCitemCode) Then Exit Function
        '��������3�h�u�f
        If Val(rs019("RefNo") & "") = 20 Then
            Chk3Discount = True
        End If

End Function

Public Function GetRsFieldString(rs As ADODB.Recordset, strField As String) As String
    On Error Resume Next
    Dim rsTmp As New ADODB.Recordset
    Dim strFString As String
        Set rsTmp = rs.Clone
        Do While Not rsTmp.EOF
            strFString = strFString & ",'" & rsTmp(strField) & "'"
            rsTmp.MoveNext
        Loop
        strFString = Mid(strFString, 2)
        Call CloseRecordset(rsTmp)
        GetRsFieldString = strFString
End Function

Public Function CallGenCharge(YYYYMM As String, strDAY1 As String, strDAY2 As String, _
    strCompCode As String, strCitemCode As String, strServiceType As String, _
    strTmpFiles As String, ByRef strRetMsg As String) As Boolean
    On Error GoTo ChkErr
    Dim lngRetCode As Integer
        lngRetCode = SF_ACCOUNTING(gcnGi, Format(YYYYMM, "yyyymm"), StrToDate(strDAY1, , False), StrToDate(strDAY2, , False), _
                                 " In (" & gCompCode & " )", strCitemCode, "", "", "", "", "", "", 1 _
                                 , 0, 0, 0, str(gCustId), garyGi(0), garyGi(1), _
                                 strServiceType, "", "", 0, strTmpFiles, strRetMsg)
        CallGenCharge = lngRetCode >= 0
    Exit Function
ChkErr:
    ErrSub FormName, "CallGenCharge"
End Function

Public Function GetGenChargeTmpFiles(strTmpSO032 As String) As String
    On Error GoTo ChkErr
    Dim strAryName As String
    Dim i As Long
    Dim strTmp As String
        strAryName = Empty
        strTmp = Empty
        strTmpSO032 = Empty
        For i = 0 To 8
            strTmp = GetTmpViewName
            If i = 5 Then strTmpSO032 = strTmp
            strAryName = IIf(strAryName = Empty, strTmp, strAryName & "," & strTmp)
        Next i
        GetGenChargeTmpFiles = strAryName
    Exit Function
ChkErr:
    ErrSub FormName, "GetTmpFiles"
End Function


'�I�s�X�b��ݵ{�� By Kin 2008/08/06
Private Function SF_ACCOUNTING(ByVal cnConn As ADODB.Connection, ByVal p_YM As String, ByVal P_DAY1 As Long, _
ByVal P_DAY2 As Long, ByVal p_CompSQL As String, ByVal p_CitemSQL As String, _
ByVal p_ClassSQL As String, ByVal P_AREASQL As String, ByVal p_ServSQL As String, _
ByVal p_ClctAreaSQL As String, ByVal P_MDUSQL As String, ByVal p_StrtSQL As String, _
ByVal P_ADDRTYPE As Long, ByVal p_GENOVERDUE As Long, ByVal p_GENPRCUST As Long, _
ByVal P_TOENDDATE As Long, ByVal P_CUSTIDLIST As String, ByVal P_UPDEN As String, _
ByVal P_UPDNAME As String, ByVal p_SERVICETYPE As String, ByVal P_BANKCODE As String, _
ByVal p_CMSQL As String, ByVal p_StopPRCust As Long, ByVal p_FileName As String, ByRef P_RETMSG As String) As Integer
On Error GoTo ChkErr
        Dim ComSF_ACCOUNTING As New ADODB.Command
        If cnConn Is Nothing Then
                'MsgBox vmsgSendConn, vbCritical, vmsgPrompt
                Exit Function
        End If

        With ComSF_ACCOUNTING
                .Parameters.Append .CreateParameter("RETURN_VALUE", adVarNumeric, adParamReturnValue)
                .Parameters.Append .CreateParameter("P_YM", adVarChar, adParamInput, 2000, p_YM)
                .Parameters.Append .CreateParameter("P_DAY1", adVarNumeric, adParamInput, , P_DAY1)
                .Parameters.Append .CreateParameter("P_DAY2", adVarNumeric, adParamInput, , P_DAY2)
                .Parameters.Append .CreateParameter("P_COMPSQL", adVarChar, adParamInput, 2000, p_CompSQL)
                .Parameters.Append .CreateParameter("P_CITEMSQL", adVarChar, adParamInput, 2000, p_CitemSQL)
                .Parameters.Append .CreateParameter("P_CLASSSQL", adVarChar, adParamInput, 2000, p_ClassSQL)
                .Parameters.Append .CreateParameter("P_AREASQL", adVarChar, adParamInput, 2000, P_AREASQL)
                .Parameters.Append .CreateParameter("P_SERVSQL", adVarChar, adParamInput, 2000, p_ServSQL)
                .Parameters.Append .CreateParameter("P_CLCTAREASQL", adVarChar, adParamInput, 2000, p_ClctAreaSQL)
                .Parameters.Append .CreateParameter("P_MDUSQL", adVarChar, adParamInput, 2000, P_MDUSQL)
                .Parameters.Append .CreateParameter("P_STRTSQL", adVarChar, adParamInput, 2000, p_StrtSQL)
                .Parameters.Append .CreateParameter("P_ADDRTYPE", adVarNumeric, adParamInput, , P_ADDRTYPE)
                .Parameters.Append .CreateParameter("P_GENOVERDUE", adVarNumeric, adParamInput, , p_GENOVERDUE)
                .Parameters.Append .CreateParameter("P_GENPRCUST", adVarNumeric, adParamInput, , p_GENPRCUST)
                .Parameters.Append .CreateParameter("P_TOENDDATE", adVarNumeric, adParamInput, , P_TOENDDATE)
                .Parameters.Append .CreateParameter("P_CUSTIDLIST", adVarChar, adParamInput, 2000, P_CUSTIDLIST)
                .Parameters.Append .CreateParameter("P_UPDEN", adVarChar, adParamInput, 2000, P_UPDEN)
                .Parameters.Append .CreateParameter("P_UPDNAME", adVarChar, adParamInput, 2000, P_UPDNAME)
                .Parameters.Append .CreateParameter("P_SERVICETYPE", adVarChar, adParamInput, 2000, p_SERVICETYPE)
                .Parameters.Append .CreateParameter("P_BANKCODE", adVarChar, adParamInput, 2000, P_BANKCODE)
                .Parameters.Append .CreateParameter("P_CMSQL", adVarChar, adParamInput, 2000, p_CMSQL)
                .Parameters.Append .CreateParameter("P_STOPPRCUST", adNumeric, adParamInput, , p_StopPRCust)
                .Parameters.Append .CreateParameter("p_FileName", adVarChar, adParamInput, 2000, p_FileName)
                .Parameters.Append .CreateParameter("P_RETMSG", adVarChar, adParamOutput, 2000, P_RETMSG)
                Set .ActiveConnection = cnConn
                .CommandText = "SF_ACCOUNTING"
                .CommandType = adCmdStoredProc
                .Execute
                P_RETMSG = .Parameters("P_RETMSG").Value & ""
                SF_ACCOUNTING = Val(.Parameters("Return_value").Value & "")
        End With
        Exit Function
ChkErr:
    ErrSub FormName, "SF_ACCOUNTING"
End Function


