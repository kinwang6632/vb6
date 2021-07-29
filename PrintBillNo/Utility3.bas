Attribute VB_Name = "Utility3"
Option Explicit
Private Const FormName = "Utility3"

'記錄目前使用者權限

Public Enum giPermissionModeEnu
    giPermissionModeAllow = True      ' True=有權限
    giPermissionModeNo = False      ' False=沒有權限
End Enum

Public Enum giFieldDataType
    giDataNumeric = 1
    giDataBoolean = 2
    giDataString = 3
    giDataMemo = 4
    giDataDate = 5
    giDataNone = 6
End Enum

Public Const ConnectErr = "連線失敗！請重新操作！"
Public Const gmsgNotSave = "新增或修改未存檔,是否真要離開?"

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
                '2007/05/09 Jacky 3216 Jim 加一個參數來解決如果回傳金額是負的問題
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

'設定各個功能的權限！
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

'將日期字串轉換成日期格式
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

'取得民國的日期格式
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
            'strReturnMsg 傳回址
            MsgBox strMsg, vbExclamation, "訊息"
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

'傳回某年某月的最後一天
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
                    '2.  設定變數初值,
                    'SQL1 = "select distinct A.PrtSNo, A.CustId, <列印批號> from SO033 A where <當時時間> = 目前時間
                    If .cboBillType.ListIndex = 0 Then
                        strSQL1 = "SELECT distinct " & lngBatchNo & ", A.PrtSNo, A.CustId, null from " & GetOwner & "SO033 A where "
                    ElseIf .cboBillType.ListIndex = 1 Then
                        strSQL1 = "SELECT distinct " & lngBatchNo & ", null, A.CustId, A.BillNo from " & GetOwner & "SO033 A where "
                    Else
                        strSQL1 = "SELECT distinct " & lngBatchNo & ", A.MediaBillNo, A.CustId, Null from " & GetOwner & "SO033 A where "
                    End If
                    ' 4.  若本段程式用於整批列印時
                    If strForm = "FRMSO3261A" Then
                        ' 則刪除單據列印log檔(so060)中印表機名稱等於<p_PrtName>的資料
                        If ExecuteSQL("DELETE FROM " & GetOwner & "SO060 WHERE PRINTERNAME = '" & .cboPrinterName.Text & "'", gcnGi, , False) Then
                            Call ErrHandle("無法刪除暫存檔 SO060 裡的資料。", False, 0, "ChargeBillChoose")
                        End If
                    End If
                    strSQL33 = ""
                    ' 5.  串起查詢條件:
                    ' ‧  若收集年月有值:
                    If .gymClctYM.GetValue <> "" Then
                        'SQL = SQL + " and A.ClctYM=<p_ClctYM> "
                        strSQL = strSQL & " AND A.ClctYM = '" & .gymClctYM.GetValue & "' "
                    End If
                    '‧  若列印選項=1:
                    If .optOption2(1).Value Then
                        'SQL = SQL + " and A.PrtCount=0"
                        strSQL = strSQL & " AND A.PrtCount = 0 "
                    End If
                    '‧  若有收費方式條件:
                    If .gimCM.GetQryStr <> "" Then
                        'SQL=SQL+" and A.CMCode"+<p_CMSQL>
                        strSQL = strSQL & " AND A.CMCode " & .gimCM.GetQryStr & " "
                    End If
                    '‧  若有收費項目條件:
                    If .gimCitemCode.GetQryStr <> "" Then
                        'SQL=SQL+" and A.CitemCode"+<p_CitemSQL>
                        strSQL = strSQL & " AND A.CitemCode " & .gimCitemCode.GetQryStr & " "
                    End If
                    
                    '‧  若有客戶類別條件:
                    If .gimClassCode.GetQryStr <> "" Then
                        'SQL=SQL+" and A.ClassCode"+<p_ClassSQL>
                        strSQL = strSQL & " AND A.ClassCode " & .gimClassCode.GetQryStr & " "
                    End If
                    '‧  若有服務區條件:
                    If .gimServCode.GetQryStr <> "" Then
                        'SQL=SQL+" and A.ServCode"+<p_ServSQL>
                        strSQL = strSQL & " AND A.ServCode " & .gimServCode.GetQryStr & " "
                    End If
                    '‧  若有行政區條件:
                    If .gimAreaCode.GetQryStr <> "" Then
                        'SQL=SQL+" and A.AreaCode"+<p_AreaSQL>
                        strSQL = strSQL & " AND A.AreaCode " & .gimAreaCode.GetQryStr & " "
                    End If
                    '‧  若有收費區條件:
                    If .gimClctAreaCode.GetQryStr <> "" Then
                        'SQL=SQL+" and A.ClctAreaCode"+<p_ClctAreaSQL>
                        strSQL = strSQL & " AND A.ClctAreaCode " & .gimClctAreaCode.GetQryStr & " "
                    End If
                    '‧  若有大樓條件:
                    If .gimMduId.GetQryStr <> "" Then
                        'SQL=SQL+" and A.MduId"+<p_MduSQL>
                        strSQL = strSQL & " AND A.MduID " & .gimMduId.GetQryStr & " "
                    End If
                    '‧  若有收費員條件:
                    If .gimClctEn.GetQryStr <> "" Then
                        'SQL=SQL+" and A.ClctEn"+<p_ClctEnSQL>
                        strSQL = strSQL & " AND A.ClctEn " & .gimClctEn.GetQryStr & " "
                    End If
                    '‧  若有產生人員條件: 901031
                    If .gimCreateEn.GetQryStr <> "" Then
                        'SQL=SQL+" and A.ClctEn"+<p_ClctEnSQL>
                        strSQL = strSQL & " AND A.CreateEn " & .gimCreateEn.GetQryStr & " "
                    End If
                    '‧  若有產生時間條件: 901031
                    If .gdaCreateTime1.GetValue <> "" Then
                        strSQL = strSQL & " AND A.CreateTime >= " & GetNullString(.gdaCreateTime1.GetValue(True), giDateV, giOracle)
                    End If
                    If .gdaCreateTime2.GetValue <> "" Then
                        strSQL = strSQL & " AND A.CreateTime < " & GetNullString(CDate(.gdaCreateTime2.GetValue(True)) + 1, giDateV, giOracle)
                    End If
                    '‧  若有應收時間條件: 901031
                    If .gdaShouldDate1.GetValue <> "" Then
                        strSQL = strSQL & " AND A.ShouldDate >= " & GetNullString(.gdaShouldDate1.GetValue(True), giDateV, giOracle)
                    End If
                    If .gdaShouldDate2.GetValue <> "" Then
                        strSQL = strSQL & " AND A.ShouldDate < " & GetNullString(CDate(.gdaShouldDate2.GetValue(True)) + 1, giDateV, giOracle)
                    End If
                    
                    '‧  若有街道條件:
                    If .gimStrtCode.GetQryStr <> "" Then
                        'SQL=SQL+" and A.StrtCode"+<p_StrtSQL>
                        strSQL = strSQL & " AND A.StrtCode " & .gimStrtCode.GetQryStr & " "
                    End If
                    
                    '‧  若單據號碼有值:
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
                        '未收資料
                        strUCSQL = " OR (UCCode Is Not Null And CancelFlag = 0) "
                        
                        '包含已收
                        If .chkIncludeUC.Value = 1 Then
                            strUCSQL = strUCSQL & " OR (UCCode Is Null And CancelFlag = 0) "
                        End If
                        '包含作廢
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
                                    
                    '列印金額為0
                    If .chkPrintZero.Value = 0 Then
                        strSQL = strSQL & " And A.ShouldAmt <> 0 "
                    End If
                    
                    '公司別
                    If .gilCompCode.GetCodeNo <> "" Then
                        strSQL = strSQL & " AND A.Compcode = " & .gilCompCode.GetCodeNo
                    End If
                    
                    '服務類別
                    If .gilServiceType.GetCodeNo <> "" Then
                        strSQL = strSQL & " AND A.ServiceType = '" & .gilServiceType.GetCodeNo & "'"
                    End If
                    
                    '銀行別
                    If .gimBank.GetDispStr <> "" Then
                        strSQL = strSQL & " AND A.BankCode " & .gimBank.qryStr
                    End If
                    
                    strSQL = strSQL & " AND ((A.RealDate is null) OR (A.ClctEn is not null)) "
                    '‧  去掉SQL最左邊之" and "
                    If strSQL <> "" Then
                        strSQL33 = strSQL
                        If Left(strSQL, 4) = " AND" Then
                            strSQL = Mid(strSQL, 5)
                        End If
                    End If
                    '‧  SQL=SQL1+SQL
                    strSQL = strSQL1 & strSQL
                    '6.  進行第一階段查詢, 並將結果存入待列印客戶暫存檔中, SQL:
                    '"insert into ????? (PrtSNo, CustId) (<SQL>)"
                    If ExecuteSQL("INSERT INTO " & GetOwner & "SO075 (BatchNO, PrtSNo, CustId, BillNo) (" & strSQL & " )", gcnGi, , False) <> giOK Then
                        Call ErrHandle("無法把資料存入 SO075 ，此程式即將關閉。", False, 2, "ChargeBillChoose")
                        Exit Function
                    End If
                    '否則: SQL = "select A.PrtSNo, A.CustId from ????? A"
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
                    If objForm.ChkMasterDetail And InStr(1, objForm.cboSort, "媒體") > 0 Then
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
'                    '‧  若有應收日期條件: 920811
'                    If .gdaShouldDate1.GetValue <> "" Then
'                        strSQL = strSQL & " AND A.ShouldDate >= " & GetNullString(.gdaShouldDate1.GetValue(True), giDateV, giOracle)
'                    End If
'                    If .gdaShouldDate2.GetValue <> "" Then
'                        strSQL = strSQL & " AND A.ShouldDate < " & GetNullString(CDate(.gdaShouldDate2.GetValue(True)) + 1, giDateV, giOracle)
'                    End If
'
'                    '‧  若有公司別條件: 920811
'                    If .gilCompCode.GetCodeNo <> "" Then
'                        'SQL=SQL+" and A.CompCode"+<p_ServiceType>
'                        strSQL33 = strSQL33 & " AND A.CompCode = " & .gilCompCode.GetCodeNo & " "
'                    End If
'
'                    '‧  若有服務類別條件: 920811
'                    If .gilServiceType.GetCodeNo <> "" Then
'                        'SQL=SQL+" and A.ServiceType"+<p_ServiceType>
'                        strSQL33 = strSQL33 & " AND A.ServiceType = '" & .gilServiceType.GetCodeNo & "' "
'                    End If
'                    '‧  若有收費項目條件:
'                    If .gimCitemCode.GetQryStr <> "" Then
'                        'SQL=SQL+" and A.CitemCode"+<p_CitemSQL>
'                        strSQL33 = strSQL33 & " AND A.CitemCode " & .gimCitemCode.GetQryStr & " "
'                    End If
'                    '‧  若有收費方式條件:
'                    If .gmdCMCode.GetQryStr <> "" Then
'                        'SQL=SQL+" and A.CMCode"+<p_CMSQL>
'                        strSQL33 = strSQL33 & " AND A.CMCode " & .gmdCMCode.GetQryStr & " "
'                    End If
'                    '‧  若有客戶類別條件:
'                    If .gmdCustClass.GetQryStr <> "" Then
'                        'SQL=SQL+" and A.ClassCode"+<p_ClassSQL>
'                        strSQL33 = strSQL33 & " AND A.ClassCode " & .gmdCustClass.GetQryStr & " "
'                    End If
'                    '‧  若有收費員條件:
'                    If .gmdClctEn.GetQryStr <> "" Then
'                        'SQL=SQL+" and A.ClctEn"+<p_ClctEnSQL>
'                        strSQL33 = strSQL33 & " AND A.ClctEn " & .gmdClctEn.GetQryStr & " "
'                    End If
'                    '‧  若有單據類別條件:
'                    If .gmdBillType.GetQryStr <> "" Then
'                        'SQL=SQL+" and A.ClctEn"+<p_ClctEnSQL>
'                        strSQL33 = strSQL33 & " AND SubStr(A.BillNo,7,1) " & .gmdBillType.GetQryStr & " "
'                    End If
'                    '‧  若有產生人員條件:
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
        '8.  執行SQL, Loop每一筆, 產生相關單據資料, 參考附件流程圖
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
                        Call ErrHandle("無法把資料寫入 SO075 ，這個程式即將關閉。", False, 2, "PutIntoSO075")
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
        '1.初值設定: 供後續查詢用的變數
        'SQLB = "select InstAddrNo, ChargeAddrNo, MduId"         (註: 這些欄位必須取得)
        strSQLB = "SELECT InstAddrNo, ChargeAddrNo, MduId"
        
        'SQLC1 = SQLC2 = SQLD = SQLE = SQLF = SQLI = 空字串
        'xSQLC1 = xSQLC2 = xSQLD = xSQLE = xSQLF = xSQLI = 空字串
        '陣列aryF: 用來存放前期實收資料所需的欄位名稱
        ReDim Preserve aryF(0 To 0) As String
        
        '2. Loop表單含蓋欄位資訊檔(FmtFld)中不是'A'開頭的每一筆:
'        If Not GetRS(rsFmtFld, "SELECT FmtFld.*, FldList.FldName FROM FmtFld, " & _
'                "FldList WHERE FmtFld.FldId = FldList.FldID AND FmtFld.FmtID = '" & _
'                strFmtID & "' AND MID(FmtFld.FldID, 1, 1) <> 'A' ORDER BY FmtFld.FldID", _
'                cnRptMDB, adUseClient, adOpenForwardOnly, adLockReadOnly) Then
'            Call errHandle("無法開啟 FmtFld 的資料，此程式即將關閉。", False, 2, "PrepareA")
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
            Call ErrHandle("無法開啟 FmtFld 的資料，此程式即將關閉。", False, 2, "PrepareA")
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
                '‧  case 'B'開頭, 且<欄位名稱>有值:
                If UCase(Left(rsFmtFld("FldID").Value, 1)) = "B" Then
                    'SQLB = SQLB + ", <欄位名稱>"
                    strSQLB = strSQLB & ", " & rsFmtFld("FldName").Value
                End If
                
                '‧  case 'C01', 'C02', 'C03','C07','CAxx' 且<欄位名稱>有值:
                If UCase(rsFmtFld("FldID").Value) = "C01" Or UCase(rsFmtFld("FldID").Value) = _
                   "C02" Or UCase(rsFmtFld("FldID").Value) = "C03" Or UCase(rsFmtFld("FldID").Value) = "C07" _
                    Or UCase(Left(rsFmtFld("FldID").Value, 2)) = "CA" Then
                    'SQLC1 = SQLC1 + ", <欄位名稱>"
                    strSQLC1 = strSQLC1 + ", " & rsFmtFld("FldName").Value
                End If
                
                '‧  case 'C04', 'C05','C06','CBxx' 且<欄位名稱>有值:
                If UCase(rsFmtFld("FldID").Value) = "C04" Or UCase(rsFmtFld("FldID").Value) = _
                   "C05" Or UCase(rsFmtFld("FldID").Value) = "C06" _
                    Or UCase(Left(rsFmtFld("FldID").Value, 2)) = "CB" Then
                    'SQLC2 = SQLC2 + ", <欄位名稱>"
                    strSQLC2 = strSQLC2 + ", " & rsFmtFld("FldName").Value
                End If
                
                '‧  case 'CCxx' 且<欄位名稱>有值:
                If UCase(Left(rsFmtFld("FldID").Value, 2)) = "CC" Then
                    'SQLC3 = SQLC3 + ", <欄位名稱>"
                    strSQLC3 = strSQLC3 + ", " & rsFmtFld("FldName").Value
                End If
                
                '‧  case 'D'開頭, 且<欄位名稱>有值:
                If UCase(Left(rsFmtFld("FldID").Value, 1)) = "D" Then
                    'SQLD = SQLD + ", <欄位名稱>"
                    strSQLD = strSQLD & ", " & rsFmtFld("FldName").Value
                End If
                
                '‧  case 'E'開頭, 且<欄位名稱>有值:
                If UCase(Left(rsFmtFld("FldID").Value, 1)) = "E" Then
                    ' SQLE = SQLE + ", <欄位名稱>"
                    strSQLE = strSQLE & ", " & rsFmtFld("FldName").Value
                End If
                
                '‧  case 'F'開頭, 且<欄位名稱>有值
                ' 且該<欄位名稱>不存在於aryF中
                If UCase(Left(rsFmtFld("FldID").Value, 1)) = "F" Then
                    If InStr(strF, Chr(9) & rsFmtFld("FldName").Value) = 0 Then
                        ReDim Preserve aryF(0 To UBound(aryF) + 1) As String
                        ' 將<欄位名稱>加入aryF中
                        aryF(UBound(aryF)) = rsFmtFld("FldName").Value
                        strF = strF & Chr(9) & rsFmtFld("FldName").Value
                    End If
                End If
                
                '‧  case 'I'開頭, 且<欄位名稱>有值:
                If UCase(Left(rsFmtFld("FldID").Value, 1)) = "I" Then
                    'SQLI = SQLI + ", <欄位名稱>"
                    strSQLI = strSQLI & ", " & rsFmtFld("FldName").Value
                End If
                
                '‧  case 'J'開頭, 且<欄位名稱>有值:
                If UCase(Left(rsFmtFld("FldID").Value, 1)) = "J" Then
                    'strSQLJ = strSQLJ + ", <欄位名稱>"
                    If rsFmtFld("FldID").Value <> "J01" Then
                        strSQLJ = strSQLJ & ", " & rsFmtFld("FldName").Value
                    End If
                End If
                '‧  case 'K'開頭, 且<欄位名稱>有值:
                If UCase(Left(rsFmtFld("FldID").Value, 1)) = "K" Then
                    'strSQLK = strSQLK + ", <欄位名稱>"
                    strSQLK = strSQLK & ", " & rsFmtFld("FldName").Value
                End If
                
                '‧  case 'L'開頭, 且<欄位名稱>有值:
                If UCase(Left(rsFmtFld("FldID").Value, 2)) >= "L0" And _
                    UCase(Left(rsFmtFld("FldID").Value, 2)) <= "L9" Then
                    'strSQLL = strSQLL + ", <欄位名稱>"
                    strSQLL = strSQLL & ", " & rsFmtFld("FldName").Value
                End If
                
                '‧  case 'LA'開頭, 且<欄位名稱>有值:
                If UCase(Left(rsFmtFld("FldID").Value, 2)) = "LA" Then
                    'strSQLL1 = strSQLL1 + ", <欄位名稱>"
                    strSQLL1 = strSQLL1 & ", " & rsFmtFld("FldName").Value
                End If
                
                '‧  case 'LB'開頭, 且<欄位名稱>有值:
                If UCase(Left(rsFmtFld("FldID").Value, 2)) = "LB" Then
                    'strSQLL2 = strSQLL2 + ", <欄位名稱>"
                    strSQLL2 = strSQLL2 & ", " & rsFmtFld("FldName").Value
                End If
            
                '‧  case 'M'開頭, 且<欄位名稱>有值:
                If UCase(Left(rsFmtFld("FldID").Value, 1)) = "M" Then
                    'strSQLK = strSQLK + ", <欄位名稱>"
                    strSQLM = strSQLM & ", " & rsFmtFld("FldName").Value
                End If
                
                '‧  case 'NA'開頭, 且<欄位名稱>有值:
                If UCase(Left(rsFmtFld("FldID").Value, 2)) = "NA" Then
                    'strSQLK = strSQLK + ", <欄位名稱>"
                    strSQLN1 = strSQLN1 & ", " & rsFmtFld("FldName").Value
                End If
                
                '‧  case 'NB'開頭, 且<欄位名稱>有值:
                If UCase(Left(rsFmtFld("FldID").Value, 2)) = "NB" Then
                    'strSQLK = strSQLK + ", <欄位名稱>"
                    strSQLN2 = strSQLN2 & ", " & rsFmtFld("FldName").Value
                End If
                
                '‧  case 'OA'開頭, 且<欄位名稱>有值:   對應該收費項目正常設備
                If UCase(Left(rsFmtFld("FldID").Value, 2)) = "OA" Then
                    'strSQLOA = strSQLOA + ", <欄位名稱>"
                    If Val(Right(rsFmtFld("FldID").Value, 2)) = 1 Then
                        strSQLOA = strSQLOA & ", " & rsFmtFld("FldName").Value
                    End If
                End If
                '‧  case 'OB'開頭, 且<欄位名稱>有值:   對應該客戶正常設備
                If UCase(Left(rsFmtFld("FldID").Value, 2)) = "OB" Then
                    'strSQLOB = strSQLOB + ", <欄位名稱>"
                    If Val(Right(rsFmtFld("FldID").Value, 2)) = 1 Then
                        strSQLOB = strSQLOB & ", " & rsFmtFld("FldName").Value
                    End If
                End If
                '‧  case 'OC'開頭, 且<欄位名稱>有值:   對應該客戶正常設備
                If UCase(Left(rsFmtFld("FldID").Value, 2)) = "OC" Then
                    'strSQLOC = strSQLOC + ", <欄位名稱>"
                    If Val(Right(rsFmtFld("FldID").Value, 2)) = 1 Then  '第一筆取就可以了
                        strSQLOC = strSQLOC & ", " & rsFmtFld("FldName").Value
                    End If
                End If
                
            End If
            rsFmtFld.MoveNext
        Loop
        
        '3. 整理各SQL指令:
        '‧  SQLB = SQLB + " from SO001 where CustId="
        '公司別
        If InStr(1, UCase(strSQLB), "COMPCODE") = 0 Then strSQLB = strSQLB & ",CompCode"
        '客戶編號
        If InStr(1, UCase(strSQLB), "CUSTID") = 0 Then strSQLB = strSQLB & ",CustId"
        '加裝機地址
        If InStr(1, UCase(strSQLB), "INSTADDRNO") = 0 Then strSQLB = strSQLB & ",InstAddrNo"
        '加收費地址
        If InStr(1, UCase(strSQLB), "CHARGEADDRNO") = 0 Then strSQLB = strSQLB & ",ChargeAddrNo"
        '加郵寄地址
        If InStr(1, UCase(strSQLB), "MAILADDRNO") = 0 Then strSQLB = strSQLB & ",MailAddrNo"
        
        strSQLB = strSQLB & " FROM " & GetOwner & "SO001 WHERE CustID = "
        
        '‧  若SQLC1有值: 收費地址
        If strSQLC1 <> "" Then
            ' 去掉左邊之 ',' 再SQLC1 = "select "+SQLC1+" from SO014 where AddrNo="
            If Left(strSQLC1, 1) = "," Then
                strSQLC1 = Mid(strSQLC1, 2)
            End If
            strSQLC1 = "SELECT " & strSQLC1 & " FROM " & GetOwner & "SO014 WHERE AddrNo = "
        End If
        
        '‧  若SQLC2有值: 裝機地址
        If strSQLC2 <> "" Then
            If Left(strSQLC2, 1) = "," Then
                strSQLC2 = Mid(strSQLC2, 2)
            End If
            strSQLC2 = "SELECT " & strSQLC2 & " FROM " & GetOwner & "SO014 WHERE AddrNo = "
        End If
        
        '‧  若SQLC3有值: 郵寄地址
        If strSQLC3 <> "" Then
            If Left(strSQLC3, 1) = "," Then
                strSQLC3 = Mid(strSQLC3, 2)
            End If
            strSQLC3 = "SELECT " & strSQLC3 & " FROM " & GetOwner & "SO014 WHERE AddrNo = "
        End If
        
        '‧  若SQLD有值:
        If strSQLD <> "" Then
            ' 去掉左邊之 ','
            If Left(strSQLD, 1) = "," Then
                strSQLD = Mid(strSQLD, 2)
            End If
            ' SQLD = "select "+SQLD+" from SO002 where CustId="
            strSQLD = "SELECT " & strSQLD & " FROM " & GetOwner & "SO002 WHERE CustId = "
        End If
        
        '‧  若SQLE有值:
        If strSQLE <> "" Then
            ' 去掉左邊之 ','
            If Left(strSQLE, 1) = "," Then
                strSQLE = Mid(strSQLE, 2)
            End If
            ' SQLE = "select "+SQLE+" from SO017 where MduId="
            strSQLE = "SELECT " & strSQLE & " FROM " & GetOwner & "SO017 WHERE MduId = "
        End If
        
        '‧  若aryF不為空:
        If strF <> "" Then
            '1. loop aryF每一元素
            For lngX = 1 To UBound(aryF)
                ' SQLF = SQLF + ", <欄位名稱>"
                strSQLF = strSQLF & "," & aryF(lngX)
            Next
            
            '2. 去掉左邊之','
            If Left(strSQLF, 1) = "," Then
                strSQLF = Mid(strSQLF, 2)
            End If
            
            ' 再SQLF = "select "+SQLF+" from SO034 where PrtSNo="
            'strSQLF = "SELECT " & strSQLF & " FROM SO034 WHERE PrtSNo = "
        End If
            
        strFieldName = Mid(strFieldName, 2)
        '‧  若SQLI有值: (91/08/07 啟用)
        If strSQLI <> "" Then
            ' 去掉左邊之 ','
            If Left(strSQLI, 1) = "," Then
                strSQLI = Mid(strSQLI, 2)
            End If
            ' SQLI = "select "+SQLI+" from ????? where CustId="
            strSQLI = "SELECT " & strSQLI & " FROM " & GetOwner & "SO107 WHERE "
        End If
        
        '‧  若SQLJ有值:
        If strSQLJ <> "" Then
            ' 去掉左邊之 ','
            If Left(strSQLJ, 1) = "," Then
                strSQLJ = Mid(strSQLJ, 2)
            End If
            ' strSQLJ = "select "+strSQLJ+" from SO304 where MduId="
            strSQLJ = "SELECT " & strSQLJ & " FROM " & GetOwner & "SO304 WHERE "
        End If
        
        '‧  若SQLK有值:
        If strSQLK <> "" Then
            ' 去掉左邊之 ','
            If Left(strSQLK, 1) = "," Then
                strSQLK = Mid(strSQLK, 2)
            End If
            ' strSQLK = "select "+strSQLK+" from SO043 where MduId="
            strSQLK = "SELECT " & strSQLK & " FROM " & GetOwner & "SO043 WHERE "
        End If
        
        '‧  若SQLL有值:
        If strSQLL <> "" Then
            '加收費地址
            If InStr(1, UCase(strSQLL), "CHARGEADDRNO") = 0 Then strSQLL = strSQLL & ",ChargeAddrNo"
            '加郵寄地址
            If InStr(1, UCase(strSQLL), "MAILADDRNO") = 0 Then strSQLL = strSQLL & ",MailAddrNo"
            ' 去掉左邊之 ','
            If Left(strSQLL, 1) = "," Then
                strSQLL = Mid(strSQLL, 2)
            End If
            ' strSQLL = "select "+strSQLL+" from SO002A where MduId="
            strSQLL = "SELECT " & strSQLL & " FROM " & GetOwner & "SO002A WHERE "
        End If
        
        '‧  若SQLL1有值: 收費地址
        If strSQLL1 <> "" Then
            ' 去掉左邊之 ',' 再SQLL1 = "select "+strSQLL1+" from SO014 where AddrNo="
            If Left(strSQLL1, 1) = "," Then
                strSQLL1 = Mid(strSQLL1, 2)
            End If
            strSQLL1 = "SELECT " & strSQLL1 & " FROM " & GetOwner & "SO014 WHERE AddrNo = "
        End If
        
        '‧  若SQLL2有值: 郵寄地址
        If strSQLL2 <> "" Then
            ' 去掉左邊之 ',' 再SQLL2 = "select "+strSQLL2+" from SO014 where AddrNo="
            If Left(strSQLL2, 1) = "," Then
                strSQLL2 = Mid(strSQLL2, 2)
            End If
            strSQLL2 = "SELECT " & strSQLL2 & " FROM " & GetOwner & "SO014 WHERE AddrNo = "
        End If
        '‧  若SQLM有值: SO004
        If strSQLM <> "" Then
            If Left(strSQLM, 1) = "," Then
                strSQLM = Mid(strSQLM, 2)
            End If
            strSQLM = "SELECT " & strSQLM & " FROM " & GetOwner & "SO004 WHERE TranBillNO = "
        End If
        
        '‧  若SQLN1有值: SO152
        If strSQLN1 <> "" Then
            If Left(strSQLN1, 1) = "," Then
                strSQLN1 = Mid(strSQLN1, 2)
            End If
            strSQLN1 = "SELECT " & strSQLN1 & " FROM " & GetOwner & "SO152 Where EmcCustID = "
        End If
        
        '‧  若SQLN2有值: SO152
        If strSQLN2 <> "" Then
            If Left(strSQLN2, 1) = "," Then
                strSQLN2 = Mid(strSQLN2, 2)
            End If
            strSQLN2 = "SELECT " & strSQLN2 & " FROM " & GetOwner & "SO152 Where EmcCustID = "
        End If
        
        '‧  若SQLOA有值: SO004 該收費項目的設備
        If strSQLOA <> "" Then
            If Left(strSQLOA, 1) = "," Then
                strSQLOA = Mid(strSQLOA, 2)
            End If
            strSQLOA = "SELECT " & strSQLOA & " FROM " & GetOwner & "SO004 WHERE "
        End If
        
        '‧  若SQLOB有值: SO004 該客戶的設備
        If strSQLOB <> "" Then
            If Left(strSQLOB, 1) = "," Then
                strSQLOB = Mid(strSQLOB, 2)
            End If
            strSQLOB = "SELECT " & strSQLOB & " FROM " & GetOwner & "SO004,(Select RefNo,CodeNo From " & GetOwner & "CD022 ) CD022 WHERE SO004.FaciCode = CD022.CodeNo And "
        End If
        
        '‧  若SQLP有值: SO004 該客戶的其他SO002
        If strSQLOC <> "" Then
            If Left(strSQLOC, 1) = "," Then
                strSQLOC = Mid(strSQLOC, 2)
            End If
            strSQLOC = "SELECT " & strSQLOC & " FROM " & GetOwner & "SO002 Where CustId = "
        End If
        
        '‧  若SQLOD有值: SO004 該客戶的其他SO138
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
        '收費單列表機
        If rs.RecordCount = 0 Then Exit Function
        If strPrinter = "" Then strPrinter = rs("BillPrinter") & ""
        '各類收費單格式代號
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

' A: 前期準備動作
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
        '取資料
        If Not ChargeBillChoose(strForm, objForm, rsData, lngBatchNo, strSQL33, blnClose) Then Exit Function
        '取欄位
        If Not ChargeBillField(cnRptMDB, strFmtID, strSQLB, strSQLC1, strSQLC2, strSQLC3, strSQLD, strSQLE, strSQLF, strSQLI, strSQLJ, strSQLK, strSQLL, strSQLL1, strSQLL2, strSQLM, strSQLN1, strSQLN2, strSQLOA, strSQLOB, strFieldName, blnJoin2, cnTmpMDB, strSQLOC, strSQLOD) Then Exit Function
        '建立地址資料更動之Log File
        If Not CreateAddrModifyTable(cnTmpMDB) Then Exit Function
        '刪除TEMP資料
        'If Not CreatePrtSNoTmp(cnTmpMDB, False) Then Exit Function
        '開始Insert Data
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
        '建立TmpMDB PK
        'If Not CreatePrtSNoTmp(cnTmpMDB, True) Then Exit Function
        On Error GoTo ChkErr
        '關闢Connection
        If lngAffecteds > 0 Then
            If Not blnRemoteMode Then MsgBox "暫存資料處理完畢，共 [" & lngAffecteds & "] 筆，現在開始列印...。", vbInformation, "訊息"
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
                    '檢查收費單地址有無異動
                    lngModiCount = Val(GetRsValue("Select Count(*) From BillAddrModify", cnTmpMDB) & "")
                    
                    strRptFileName = GetRsValue("Select FmtFileName From FmtList Where FmtId ='" & strFmtID & "'", cnRptMDB) & ""
                    '非使用遠端出單 92/06/23
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
                MsgBox "處理零筆資料，不需要列印。", vbInformation, "訊息"
        End If
        If strForm = "FRMSO3261A" Or strForm = "FRMSO3262A" Then
            '‧  若備註欄有值, 則將有值的備註欄內容存回SO039對應欄位中
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
        PrintRpt2 strPrinterName, strRptFileName, , "收費單列印", "Select * Form PrtSNoTmp" & strWhere, , , True, "TMP1111.mdb", , strFormula, , , , , , , , blnToPrinter, , , , True, False
        InsertErrLog cnTmpMDB, 2
        '.Action = 1
        On Error Resume Next
        Call clsXPrinter.DefaultPrinter
        InsertErrLog cnTmpMDB, 3
        '列印地址有更動的資料
        If lngModiCount > 0 Then
            InsertErrLog cnTmpMDB, 31
            MsgBox "收費單地址資料有異動,將列印異動表!!", vbInformation, gimsgPrompt
            PrintRpt2 GetPrinterName(5), "BillAddrModify.Rpt", , "收費單地址異動表 [Utility3]", "Select * From BillAddrModify", , , True, "Tmp1111.MDB"
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
'            '.WindowTitle = "收費單列印"
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
'            PrintRpt strPrinterName, strRptFileName, , "收費單列印", "Select * Form PrtSNoTmp", , , True, "TMP1111.mdb", , strFormula2, , , , , , , , blnToPrinter, , , , True, False
'            '.Action = 1
'            On Error Resume Next
'            Call clsXPrinter.DefaultPrinter
'            'frmPrintForm.Show vbModal
'            '列印地址有更動的資料
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
    Dim lngNowCustID As Long            ' 目前客編
    Dim strNowPrtSNO As String          ' 目前序號
    Dim lngCustStatusCode As Long       ' 此客戶的客戶狀態
    Dim lShouldPrint As Boolean         ' 此客戶是否需要印單
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
        ' Loop 每一筆資料
        On Error Resume Next
        If strForm <> "FRMSO1100BMDI" Then objForm.Picture1.Visible = True
        On Error GoTo ER
        
        'If rs.RecordCount = 0 Then GotoFlowProc = True: lngAffecteds = 0: Exit Function
        If Not GetSystemPara(rsPara, strCompCode, "", "SO043", "Para19,Para20") Then Exit Function
        '檢查Source Table
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
            lngNowCustID = rs("CustID").Value           ' 目前客編 = 該筆客戶編號
            If strForm = "FRMSO3261A" Then
                strNowPrtSNO = (rs("PrtSNO").Value & "")    ' 目前序號 = 該筆印單序號
            End If
            lShouldPrint = True
            
'            If strForm = "FRMSO3261A" Then
'                ' 如果不要列印待拆戶
'                lngCustStatusCode = GetValueFromRs2("SELECT CustStatusCode FROM SO002 WHERE CustId = " & lngNowCustID & " And ServiceType= '", gcnGi)
'                'If objForm.chkOption1.Value = 0 Then
'                    ' 取該客戶的客戶狀態
'                    ' 非待拆戶?
'                    ' 3 = 拆機, 6 = 拆機中
'                '    If lngCustStatusCode <> 3 And lngCustStatusCode <> 6 Then
'                '        lShouldPrint = True
'                '    End If
'                'Else
'                '    lShouldPrint = True
'                'End If
'                '‧  若有客戶類別條件: 2001/11/12新增
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
            
            ' 此客戶需要印單
            If lShouldPrint Then
                'Debug.Print rs("PrtSNo")
                Select Case UCase(strForm)
                    Case "FRMSO3261A"
                        '‧  若有客戶狀態條件: 91/09/04
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
                        'SO009 拆機預約日期 93/08/18
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
                        ' rsA="select ... from SO033 where pRTsnO='<目前序號>' ORDER BY ShouldDate DESC"
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
                    '檢查地址是否有更改
                    If Not ChkAddressModify(rsA("BillNo"), rsA("Item"), rsA("PrtSNo") & "", rsA("CustId"), rsA("AddrNo"), rsA("CitemName"), cnTmpMDB) Then Exit Function
                    
                    On Error GoTo TransErr
                    rsPrtSNOTmp.AddNew
                    ' Loop rsA每一筆資料
                    Do While Not rsA.EOF And rsA.RecordCount > 0
                        ' 若Counter = 0

                        If lngCounter = 0 Then
                            If Not GetrsFmtFld(rsFmtFld, rsFmtFldTmp(), 0, "SELECT A.*, B.FldName FROM FmtFld A, FldList B WHERE A.FldId = B.FldID AND A.FmtID = '" & _
                               strFmtID & "' AND MID(B.FldID, 1, 2) >= 'A0' AND MID(B.FldID, 1, 2) <= 'A9' And A.Type = B.Type ORDER BY A.FldID", cnRptMDB) Then Exit Function
                            
'                            If Not GetRS(rsFmtFld, "SELECT A.*, B.FldName FROM FmtFld A, FldList B WHERE A.FldId = B.FldID AND A.FmtID = '" & _
                               strFmtID & "' AND MID(B.FldID, 1, 2) >= 'A0' AND MID(B.FldID, 1, 2) <= 'A9' And A.Type = B.Type ORDER BY A.FldID", cnRptMDB) Then Exit Function
                            Do While Not rsFmtFld.EOF
                                ' 取基本台相關資訊
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
                            
                            ' 2.根據Counter值，設定對應欄位值，
                            '   例：[AAxx],[ABxx],[ACxx],[ADxx],
                            '       [AExx],[AFxx],(xx表01~10)
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
                                    '執行xSQLOA -> rsOA
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
                            Call ErrHandle("客戶編號為 [" & lngNowCustID & _
                               "] 的客戶，此月份的收費資料超出20項！下面這一項收費明細將" & _
                               "會被省略掉：" & vbCrLf & "收費項目代碼: [" & rsA("CitemCode").Value & "]" & vbCrLf & _
                               "收費項目: [" & rsA("CitemName").Value & "]" & vbCrLf & _
                               "應收金額: [" & rsA("ShouldAmt").Value & "]", False, , "GotoFlowProc")
                        End If
                        ' 3.累計應收總金額[A07]
                        rsPrtSNOTmp("A07").Value = Val(rsPrtSNOTmp("A07").Value & "") + Val(rsA("ShouldAmt").Value & "")
                        On Error Resume Next
                        '6.若SLQI有值: (表示需依目前客編取相關發票欄位)  (91/08/19 修改)
                        If rsA.AbsolutePosition = 1 Then
                            If strSQLI <> "" Then
                                'xSQLI = SQLI + "<目前客編>"
                                If Not GetrsFmtFld(rsFmtFld, rsFmtFldTmp(), 1, "SELECT A.*, B.FldName FROM FmtFld A, FldList B WHERE A.FldId = B.FldID AND A.FmtID = '" & _
                                   strFmtID & "' AND MID(B.FldID, 1, 1) = 'I' And A.Type = B.Type ORDER BY A.FldID", cnRptMDB) Then Exit Function
                                
                                strxSQLI = strSQLI & " BillNo = '" & rsA("BillNo") & "' And CitemName = '" & rsA("CitemName") & "'"
                                '執行xSQLI -> rsI
                                If Not GetRS(rsI, strxSQLI, gcnGi) Then Exit Function
                                If rsI.RecordCount > 0 Then
                                    Do While Not rsFmtFld.EOF
                                        rsPrtSNOTmp(rsFmtFld("FldID").Value).Value = rsI(rsFmtFld("FldName").Value)
                                        rsFmtFld.MoveNext
                                    Loop
                                End If
                            End If
                            
                            '7.若SLQK有值: (表示需依CompCode,ServiceType取相關收費參數)
                            If strSQLK <> "" Then
                                strxSQLK = strSQLK & " CompCode = " & rsA("CompCode").Value & " And ServiceType = '" & rsA("ServiceType") & "'"
                                If Not GetRS(rsK, strxSQLK, gcnGi) Then Exit Function
                            End If
                            
                            '8.若SLQL有值: (表示需依CustId, AccountNo取相關收費參數)
                            If strSQLL <> "" Then
                                strxSQLL = strSQLL & " CustId = " & rsA("CustId").Value & " And AccountNo = '" & rsA("AccountNo") & "' And CompCode = " & rsA("CompCode")
                                If Not GetRS(rsL, strxSQLL, gcnGi) Then Exit Function
                                '9.若SQLL1有值: (表示需依收費地址取相關地址欄位)
                                If strSQLL1 <> "" Then
                                    '93/05/13 加判斷該資料無值
                                    If rsL.RecordCount > 0 Then
                                        strxSQLL1 = strSQLL1 & rsL("ChargeAddrNO").Value & " And CompCode = " & rsA("CompCode")
                                    Else
                                        strxSQLL1 = strSQLL1 & "0"
                                    End If
                                    If Not GetRS(rsL1, strxSQLL1, gcnGi) Then Exit Function
                                End If
                                
                                '10.若SQLL2有值: (表示需依郵寄地址取相關地址欄位)
                                If strSQLL2 <> "" Then
                                    '93/05/13 加判斷該資料無值
                                    If rsL.RecordCount > 0 Then
                                        strxSQLL2 = strSQLL2 & rsL("MailAddrNo").Value & " And CompCode = " & rsA("CompCode")
                                    Else
                                        strxSQLL2 = strSQLL2 & "0"
                                    End If
                                    If Not GetRS(rsL2, strxSQLL2, gcnGi) Then Exit Function
                                End If
                                '11.若SLQM有值: (表示需依結轉編號取相關設備欄位)
                                If strSQLM <> "" Then
                                    ' xSQLM = SQLM + " '<rsA.BillNo>' "
                                    strxSQLM = strSQLM & "'" & rsA("BillNo").Value & "'"
                                    ' 執行xSQLM -> rsM
                                    If Not GetRS(rsM, strxSQLM, gcnGi) Then Exit Function
                                End If
                            End If
                            
                        
                            '4.若SLQD有值: (表示需依目前客編取相關帳戶欄位)
                            If strSQLD <> "" Then
                                ' xSQLD = SQLD + "<目前客編>"
                                strxSQLD = strSQLD & rsA("CustId").Value & " And CompCode = " & rsA("CompCode") & " And ServiceType = '" & rsA("ServiceType") & "'"
                                ' 執行xSQLD -> rsD
                                If Not GetRS(rsD, strxSQLD, gcnGi) Then Exit Function
                            End If
                            
                            '所有正常設備的資料
                            If strSQLOB <> "" Then
                                If Not GetrsFmtFld(rsFmtFld, rsFmtFldTmp(), 6, "SELECT A.*, B.FldName FROM FmtFld A, FldList B WHERE A.FldId = B.FldID AND A.FmtID = '" & _
                                   strFmtID & "' AND MID(B.FldID, 1, 3) >= 'OBA' AND MID(B.FldID, 1, 3) <= 'OBZ' And A.Type = B.Type ORDER BY A.FldID", cnRptMDB) Then Exit Function
                                
                                strxSQLOB = strSQLOB & " CustId = " & rsA("CustId") & " And CompCode = " & rsA("CompCode") & " And PRDate is null"
                                '執行xSQLOA -> rsOA
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
                            '所有服務的SO002
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
                
                
    '----Start 流程Ｂ------
                    ' 執行流程B
                    'B: 為某客戶新增一筆帳單資料
                    '1. 取此帳單各欄位所需資訊:
                    '1.取該客戶基本資料:
                    'xSQLB = SQLB + "<目前客編>"
                    strxSQLB = strSQLB & lngNowCustID & " And CompCode = " & lngCompCode
                    
                    '執行xSQLB -> rsB
                    If Not GetRS(rsB, strxSQLB, gcnGi, adUseClient) Then Exit Function
                    
                    '2.若SQLC1有值: (表示需依收費地址取相關地址欄位)
                    If strSQLC1 <> "" Then
                        'xSQLC1 = SQLC1 + "<rsB.ChargeAddrNo>"
                        strxSQLC1 = strSQLC1 & rsB("ChargeAddrNO").Value & " And CompCode = " & lngCompCode
                        '執行xSQLC1 -> rsC1
                        If Not GetRS(rsC1, strxSQLC1, gcnGi) Then Exit Function
                    End If
                    
                    '3.若SLQC2有值: (表示需依裝機地址取相關地址欄位)
                    If strSQLC2 <> "" Then
                        'xSQLC2 = SQLC2 + "<rsB.InstAddrNo>"
                        strxSQLC2 = strSQLC2 & rsB("InstAddrNo").Value & " And CompCode = " & lngCompCode
                        '執行xSQLC2 -> rsC2
                        If Not GetRS(rsC2, strxSQLC2, gcnGi) Then Exit Function
                    End If
                    
                    '3.若SLQC3有值: (表示需依裝機地址取相關地址欄位)
                    If strSQLC3 <> "" Then
                        'xSQLC3 = SQLC3 + "<rsB.MailAddrNo>"
                        strxSQLC3 = strSQLC3 & rsB("MailAddrNo").Value & " And CompCode = " & lngCompCode
                        '執行xSQLC2 -> rsC2
                        If Not GetRS(rsC3, strxSQLC3, gcnGi) Then Exit Function
                    End If
                    
                    '5.若SLQE有值: (表示需依大樓編號取相關大樓欄位)
                    If strSQLE <> "" Then
                        ' xSQLE = SQLE + " '<rsB.MduId>' "
                        strxSQLE = strSQLE & "'" & rsB("MduID").Value & "' And CompCode = " & lngCompCode
                        ' 執行xSQLE -> rsE
                        If Not GetRS(rsE, strxSQLE, gcnGi) Then Exit Function
                    End If
                    
                    '6.若SLQJ有值: (表示需依地址編號取相關地址資料)
                    If strSQLJ <> "" Then
                        ' xSQLJ = SQLJ + " '<rsB.MduId>' "
                        strxSQLJ = "Select A.*,B.Description ContractStatusName From ( " & strSQLJ & " CableAddrNo = " & rsB("InstAddrNo").Value & " Union All " & _
                            strSQLJ & " AddrNo2 = " & rsB("InstAddrNo").Value & " Union All " & _
                            strSQLJ & " AddrNo3 = " & rsB("InstAddrNo").Value & " ) A," & GetOwner & "CD062 B Where A.CATVId= B.CodeNo Order By B.Priority"
                        ' 執行xSQLJ -> rsJ
                        If Not GetRS(rsJ, strxSQLJ, gcnGi, , , , , , , 1) Then Exit Function
                    End If
                    
                    '93/11/22 CM/CT用 Jacky
                    If strSQLN1 <> "" Then
                        strxSQLN1 = strSQLN1 & " " & rsB("CustId").Value & " And EbtServiceType='CM' ORDER BY UpdTime DESC"
                        ' 執行xSQLN1 -> rsN1
                        If Not GetRS(rsN1, strxSQLN1, gcnGi, , , , , , , 1) Then Exit Function
                    End If
                    
                    '93/11/22 CM/CT用 Jacky
                    If strSQLN2 <> "" Then
                        strxSQLN2 = strSQLN2 & " " & rsB("CustId").Value & " And EbtServiceType='CT' ORDER BY UpdTime DESC"
                        ' xSQLN2 -> rsN2
                        If Not GetRS(rsN2, strxSQLN2, gcnGi, , , , , , , 1) Then Exit Function
                    End If
                    
                    '2. Loop表單含蓋欄位資訊檔(FmtFld)中不是'F'開頭的每一筆:
                    If Not GetrsFmtFld(rsFmtFld, rsFmtFldTmp(), 2, "SELECT A.*, B.FldName FROM FmtFld A, FldList B WHERE A.FldId = B.FldID AND A.FmtID = '" & _
                           strFmtID & "' AND MID(B.FldID, 1, 1) <> 'F' And A.Type = B.Type ORDER BY A.FldID", cnRptMDB) Then Exit Function
            
                    Do While Not rsFmtFld.EOF
                        'Debug.Print rsFmtFld.Fields("FldID")
                        Select Case UCase(rsFmtFld("FldID").Value)
                            '‧  case [C01]: [C01] = rsC1.ZipCode            郵遞區號(收費地址)
                            Case "C01"
                                rsPrtSNOTmp("C01").Value = rsC1("ZipCode").Value
                            '‧  case [C02]: [C02] = rsC1.AreaName           行政區名稱(收費地址)
                            Case "C02"
                                rsPrtSNOTmp("C02").Value = rsC1("AreaName").Value
                            '‧  case [C03]: [C03] = rsC1.ServName           服務區名稱(收費地址)
                            Case "C03"
                                rsPrtSNOTmp("C03").Value = rsC1("ServName").Value
                            '‧  case [C04]: [C04] = rsC2.AreaName           行政區名稱(裝機地址)
                            Case "C04"
                                rsPrtSNOTmp("C04").Value = rsC2("AreaName").Value
                            '‧  case [C05]: [C05]= rsC2.ZipCode         郵遞區號(裝機地址)
                            Case "C05"
                                rsPrtSNOTmp("C05").Value = rsC2("ZipCode").Value
                            '‧  case [C06]: [C06]= rsC2.BourgName         村里名稱(裝機地址)
                            Case "C06"
                                rsPrtSNOTmp("C06").Value = rsC2("BourgName").Value
                            '‧  case [C07]: [C07]= rsC1.BourgName         村里名稱(收費地址)
                            Case "C07"
                                rsPrtSNOTmp("C07").Value = rsC1("BourgName").Value
                            '‧  case [A08]: [A08] = [A07]轉為中文大寫字串       應收總金額(中文)
                            Case "A08"
                                rsPrtSNOTmp("A08").Value = ToChineseN(rsPrtSNOTmp("A07").Value & "")
                            '‧  case [AG01]: [AG01] = [A07]應收總金額個位數(中文)
                            Case "AG01"
                                rsPrtSNOTmp("AG01").Value = ToChineseNumber(rsPrtSNOTmp("A07").Value, 1)
                            '‧  case [AG02]: [AG02] = [A07]應收總金額拾位數(中文)
                            Case "AG02"
                                rsPrtSNOTmp("AG02").Value = ToChineseNumber(rsPrtSNOTmp("A07").Value, 2)
                            '‧  case [AG03]: [AG03] = [A07]應收總金額佰位數(中文)
                            Case "AG03"
                                rsPrtSNOTmp("AG03").Value = ToChineseNumber(rsPrtSNOTmp("A07").Value, 3)
                            '‧  case [AG04]: [AG04] = [A07]應收總金額仟位數(中文)
                            Case "AG04"
                                rsPrtSNOTmp("AG04").Value = ToChineseNumber(rsPrtSNOTmp("A07").Value, 4)
                            '‧  case [AG05]: [AG05] = [A07]應收總金額萬位數(中文)
                            Case "AG05"
                                rsPrtSNOTmp("AG05").Value = ToChineseNumber(rsPrtSNOTmp("A07").Value, 5)
                            '‧  case [AG06]: [AG06] = [A07]應收總金額拾萬位數(中文)
                            Case "AG06"
                                rsPrtSNOTmp("AG06").Value = ToChineseNumber(rsPrtSNOTmp("A07").Value, 6)
                            '‧  註: [A01]~[A07], [AAxx], [ABxx], [ACxx], [ADxx], [AExx], [AFxx]未列於此
                            '    因這些已在先前他處取得 (xx表01~10)
                            Case Else
                                Select Case Left(UCase(rsFmtFld("FldID").Value), 2)
                                    Case "CA"       '收費地址
                                        If rsC1.RecordCount > 0 Then
                                            rsPrtSNOTmp(rsFmtFld("FldID").Value).Value = rsC1(rsFmtFld("FldName").Value)
                                        End If
                                    Case "CB"       '裝機地址
                                        If rsC2.RecordCount > 0 Then
                                            rsPrtSNOTmp(rsFmtFld("FldID").Value).Value = rsC2(rsFmtFld("FldName").Value)
                                        End If
                                    Case "CC"       '郵寄地址
                                        If rsC3.RecordCount > 0 Then
                                            rsPrtSNOTmp(rsFmtFld("FldID").Value).Value = rsC3(rsFmtFld("FldName").Value)
                                        End If
                                    Case "LA"       '收費地址
                                        If rsL1.RecordCount > 0 Then
                                            rsPrtSNOTmp(rsFmtFld("FldID").Value).Value = rsL1(rsFmtFld("FldName").Value)
                                        End If
                                    Case "LB"       '郵寄地址
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
                    
                    '3. 若SQLF有值: (表示需取前次實收資料)
                    If strSQLF <> "" Then
                        On Error Resume Next
                        
                        ' 取前期印單序號
                        '"select max(PrtSNo) from SO034 where CustId=<目前客編>
                        'and PrtSNo < 'YYYYMM' "
                        ' 註: YYYYMM為印單年月西曆字串 , 若本段用於單筆收費單據列印SO3262A中,
                        ' 則YYYYMM為印單截止年月西曆字串
                        '90/09/07 在NTY修改的.(Pierre,KC,Lawrence,Jacky )
'                        If strForm = "FRMSO3262A" Then
'                            If Not GetRS(rsSO034, "SELECT MAX(PrtSNO) AS M FROM SO034 WHERE " & "CustID = " & lngNowCustID & "AND SO034.PrtSNo < '" & objForm.gymPrtYM2.GetValue & "'") Then Exit Function
'                        ElseIf strForm = "FRMSO3261A" Then
'                            If Not GetRS(rsSO034, "SELECT MAX(PrtSNO) AS M FROM SO034 WHERE " & "CustID = " & lngNowCustID & "AND SO034.PrtSNo < '" & objForm.gymPrtYM.GetValue & "'") Then Exit Function
'                        Else
'                            If Not GetRS(rsSO034, "SELECT MAX(PrtSNO) AS M FROM SO034 WHERE " & "CustID = " & lngNowCustID) Then Exit Function
'                        End If
                        '90/09/07 在NTY修改的.(Pierre,KC,Lawrence,Jacky )
                        Dim strFieldF As String
                        Dim strBillNo As String
                        strFieldF = " BillNo ,Item, " & strSQLF
                        
                        If InStr(UCase(strSQLF), "CITEMCODE") = 0 Then strFieldF = strFieldF & ", CITEMCODE"
                        If InStr(UCase(strSQLF), "REALDATE") = 0 Then strFieldF = strFieldF & ", REALDATE"
                        If InStr(UCase(strSQLF), "UCCODE") = 0 Then strFieldF = strFieldF & ", UCCODE"
                        If InStr(UCase(strSQLF), "REALAMT") = 0 Then strFieldF = strFieldF & ", REALAMT"
                        If InStr(UCase(strSQLF), "SHOULDAMT") = 0 Then strFieldF = strFieldF & ", SHOULDAMT"
                        
                        '請不要再亂改
                        '90/10/31 Pierre & KC & Janis 討論後結果 RealAmt >0 And PeriodFlag =1 And BillNo = 'BT' Order By RealDate Desc, Substr(BillNo,7,1)
                        
                        '91/10/31 產生基本台或明細
                        strWhere = ""
                        If Val(rsPara("Para19") & "") = 1 Then strWhere = " And A.CitemCode In (Select CodeNo From " & GetOwner & "CD019 Where RefNo =1) "
                        
                        If Not GetRS(rsSO034, "Select * From (Select A.BillNo,Max(RealDate) RealDate From (SELECT " & strFieldF & " FROM " & GetOwner & "SO033 A WHERE A.CustID = " & lngNowCustID & _
                            " And Nvl(CancelFlag,0)=0 Union All  SELECT " & strFieldF & " FROM " & GetOwner & "SO034 A Where  A.CustID = " & lngNowCustID & "  And Nvl(CancelFlag,0)=0  ) A," & GetOwner & "CD019 B Where A.CitemCode = B.CodeNO And RealAmt > 0 And B.PeriodFlag = 1 And A.RealDate is Not Null And A.UCCode Is Null Group By BillNo) A Order By RealDate Desc") Then Exit Function
                        
                        ' 若取得max(PrtSNo)有值, 則:
                        lngCounter = 1
                        '91/10/31 在選取產生期數
                        Do While Not rsSO034.EOF And Val(rsPara("Para20") & "") >= lngCounter
                            ' xSQLF = SQLF + " '<max(PrtSNo) >' "
                            'strxSQLF = strSQLF & rsSO034("M").Value
                            If Not GetRS(rsF, "Select A.* From (SELECT " & strFieldF & " FROM " & GetOwner & "SO033 A WHERE A.BillNo = '" & rsSO034("BillNo") & _
                                "' And Nvl(CancelFlag,0)=0 " & strWhere & " Union All  SELECT " & strFieldF & " FROM " & GetOwner & "SO034 A Where  A.BillNo = '" & rsSO034("BillNo") & "'  And Nvl(CancelFlag,0)=0  ) A Where A.BillNo ='" & rsSO034("BillNo") & "' And A.RealDate is Not Null And A.UCCode Is Null " & strWhere & " Order By A.RealDate Desc") Then Exit Function
                            
                            ' 執行xSQLF a rsF
                            'If Not GetRS(rsF, strxSQLF, gcnGi) Then Exit Function
                            Dim strFldStr As String
                            ' loop rsF 每筆資料:
                            Do While Not rsF.EOF And Val(rsPara("Para20") & "") >= lngCounter
                                ' [F01] = [F01] + rsF.ShouldAmt       累計前次應收總金額
                                rsPrtSNOTmp("F01").Value = Val(rsPrtSNOTmp("F01").Value & "") + Val(rsF("ShouldAmt").Value & "")
                                
                                ' [F02] = [F02] + rsF.RealAmt         累計前次實收總金額
                                rsPrtSNOTmp("F02").Value = Val(rsPrtSNOTmp("F02").Value & "") + Val(rsF("RealAmt").Value & "")
                                ' 若rsF.CitemCode = <基本台收費項目代碼>
                                If rsF("CitemCode").Value = GetValueFromRs2("SELECT " & _
                                   "CodeNo FROM " & GetOwner & "CD019 WHERE RefNo = 1", gcnGi) Then
                                    ' 且[F03]<rsF.RealDate, 則:
                                    If rsPrtSNOTmp("F03").Value < rsF("RealDate").Value Then
                                        '[F03] = rsF.RealDate                取基本台前次收費日
                                        rsPrtSNOTmp("F03").Value = rsF("RealDate").Value
                                    End If
                                End If
                                '2. loop表單含蓋欄位資訊檔中的每一筆以  'F'開頭的資料:
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
                                
'                                '‧  case [FAxx]: [FAxx] = rsF.CitemCode     前次項目代碼
'                                If InStr(1, strFldStr, "FA" & Format(lngCounter, "00")) > 0 Then
'                                    rsPrtSNOTmp("FA" & Format(lngCounter, "00")).Value = rsF("CitemCode").Value
'                                End If
'
'                                '‧  case [FBxx]: [FBxx] = rsF.CitemName     前次項目名稱
'                                If InStr(1, strFldStr, "FB" & Format(lngCounter, "00")) > 0 Then
'                                    rsPrtSNOTmp("FB" & Format(lngCounter, "00")).Value = rsF("CitemName").Value
'                                End If
'
'                                '‧  case [FCxx]: [FCxx] = rsF.ShouldAmt     前次項目應收金額
'                                If InStr(1, strFldStr, "FC" & Format(lngCounter, "00")) > 0 Then
'                                    rsPrtSNOTmp("FC" & Format(lngCounter, "00")).Value = rsF("ShouldAmt").Value
'                                End If
'
'                                '‧  case [FDxx]: [FDxx] = rsF.REalPeriod     前次項目期數
'                                If InStr(1, strFldStr, "FD" & Format(lngCounter, "00")) > 0 Then
'                                    rsPrtSNOTmp("FD" & Format(lngCounter, "00")).Value = rsF("RealPeriod").Value
'                                End If
'
'                                '‧  case [FExx]: [FExx] = rsF.RealStartDate     前次項目期限起始日
'                                If InStr(1, strFldStr, "FE" & Format(lngCounter, "00")) > 0 Then
'                                    rsPrtSNOTmp("FE" & Format(lngCounter, "00")).Value = rsF("RealStartDate").Value
'                                End If
'
'                                '‧  case [FFxx]: [FFxx] = rsF.RealStopDate     前次項目期限截止日
'                                If InStr(1, strFldStr, "FF" & Format(lngCounter, "00")) > 0 Then
'                                    rsPrtSNOTmp("FF" & Format(lngCounter, "00")).Value = rsF("RealStopDate").Value
'                                End If
'
'                                '‧  case [FGxx]: [FGxx] = rsF.RealAmt     前次項目入帳金額
'                                If InStr(1, strFldStr, "FG" & Format(lngCounter, "00")) > 0 Then
'                                    rsPrtSNOTmp("FG" & Format(lngCounter, "00")).Value = rsF("RealAmt").Value
'                                End If
'
'                                '‧  case [FHxx]: [FHxx] = rsF.UCCode     前次項目未收原因代碼
'                                If InStr(1, strFldStr, "FH" & Format(lngCounter, "00")) > 0 Then
'                                    rsPrtSNOTmp("FH" & Format(lngCounter, "00")).Value = rsF("UCCode").Value
'                                End If
'
'                                '‧  case [FIxx]: [FIxx] = rsF.STCode     前次項目短收原因代碼
'                                If InStr(1, strFldStr, "FI" & Format(lngCounter, "00")) > 0 Then
'                                    rsPrtSNOTmp("FI" & Format(lngCounter, "00")).Value = rsF("STCode").Value
'                                End If
'
'                                '‧  case [FJxx]: [FJxx] = rsF.RealDate     前次項目收費日期
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
                    
                    '4. 新增一筆帳單資料至帳單暫存檔中, 其中:
                    '[A03]=<目前客編>, 其他欄位在上述程式中已取得.
                    '並累計帳單數BillCount = BillCount + 1
                    rsPrtSNOTmp.Update
                    lngBillcount = lngBillcount + 1
                    '93/09/21
                    If strForm <> "FRMSO32A2A" Then
                        '5. 將該筆帳單所對應的應收明細資料(可能有多筆)中的列印次數加1
                        '   並更新列印人員與列印時間, SQL:
                        '   "update SO033 set PrtCount=PrtCount+1, PrintEn=<p_UpdEn>,
                        '   PrintTime=<當時時間> where PrtSNo=<本筆帳單之印單序號>"
                        ' Error 90/08/20 Jacky 修改 Where 條件同rsA的Where 條件
                        If ExecuteSQL("UPDATE " & GetOwner & StrTableName & " A SET PrtCount = PrtCount + 1, " & _
                           "PrintEn = '" & garyGi(1) & "', PrintTime = SYSDATE " & _
                           Mid(UCase(rsA.Source), InStr(1, UCase(rsA.Source), "WHERE"), InStr(1, UCase(rsA.Source), "ORDER") - InStr(1, UCase(rsA.Source), "WHERE")), gcnGi, , False) <> giOK Then Exit Function
                    End If
                    
                    If strForm = "FRMSO3261A" Then
                        '6. 若本段程式用於整批列印，則新增一筆LOG資料至單據列印LOG檔(SO060), 各欄位為:
                        '  .印表機名稱 PrinterName=<p_PrtName>
                        '  .順序編號 Ord = BillCount
                        '  .單據編號 BillNo = 本筆帳單之單據編號
                        '  .印單序號 PrtSNo = 本筆帳單之印單序號
                        If ExecuteSQL("INSERT INTO " & GetOwner & "SO060 (PRINTERNAME, ORD, BILLNO) VALUES ('" & _
                            objForm.cboPrinterName.Text & "', " & lngBillcount & ", " & lngNowCustID & ")", gcnGi, , False) <> giOK Then Exit Function
                    End If
                                    
                    '7. 將帳單資料對應變數全清為初值
                    ' ** 因為沒有存在變數裡，所以不須清空
                    '----End 流程Ｂ------
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
    strF0 = Array("零", "壹", "貳", "參", "肆", "伍", "陸", "柒", "捌", "玖", "拾")
    strF1 = Array("佰萬", "拾萬", "萬", "仟", "佰", "拾", "")
    strF2 = Format(lngN, "0000000")
    If Pos <> 0 Then
        ToChineseNumber = strF0(Val(Mid(strF2, Len(strF2) - Pos + 1, 1)))
        Exit Function
    End If
    
    For lngX = 1 To 7
        strOut = strF0(Val(Mid(strF2, 8 - lngX, 1))) & strF1(7 - lngX) & strOut
    Next
    ToChineseNumber = strOut & "圓"
End Function

'派工單使用
Public Function GetRStartDate(strCitemCode As String, lngCustId As Long, _
        Optional objForm As Object, Optional blnUseSo003 As Boolean = True, _
        Optional strFaciSeqNo As String, Optional lngSO003BudgetPeriod As Long, _
        Optional rs003 As ADODB.Recordset) As String
  On Error GoTo ChkErr
    '根據客戶編號，收費項目代碼，至客戶週期性收費項目檔(So003)取次收日(Clctdate)
    '若本程式用於派工單下(ZZ=1)，則派單工程式需值來該工單之1、完工日期 2、預約日期
    Dim varClctDate As Variant
    Dim strSQL As String
    Dim lngTmpCustId As Long
    Dim intTmpItem As Long
    '取得收費項目代碼
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
        
    '判斷該工單是否有完工日
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

'派工單使用
Public Function ReturnRealStartDate(strCitemCode As String, lngCustId As Long, _
    Optional strResvTime As String, Optional ByRef objForm As Object, _
    Optional strFaciSeqNo As String, Optional ByRef lngBudgetPeriod As Long) As String
  On Error GoTo ChkErr
    '根據客戶編號，收費項目代碼，至客戶週期性收費項目檔(So003)取次收日(Clctdate)
    '若本程式用於派工單下(ZZ=1)，則派單工程式需值來該工單之1、完工日期 2、預約日期
    Dim strClctDate As String
    Dim rsClctDate As New ADODB.Recordset
    Dim strSQL As String
    Dim lngTmpCustId As Long
    Dim intTmpItem As Long
    '取得收費項目代碼
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
    '判斷該工單是否有完工日
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
                '檢查是不有該欄位
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
                '檢查是不有該欄位
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
        If Len(rsTmp("FileName2") & "") = 0 Or Len(rsTmp("CorpId") & "") = 0 Then MsgBox "銀行代碼檔中無定義轉帳輸入檔名或收件單位代號", vbCritical, giMsgWarning: Exit Function
        Set clsBankATMNo = CreateObject("csBankATMNo.cls" & rsTmp("FileName2"))
        If Err.Number = 429 Then MsgBox "銀行代碼檔中定義之轉帳輸入檔名 : [" & rsTmp("FileName2") & "] 不正確,請確認!!", vbCritical, giMsgWarning: Exit Function
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
    Dim strSample As String         'Sample 檔的內容...
    Dim strRead As String             '轉換後的內容...
        strRptPath = ReadGICMIS1("RptPath")
        
        If strTXTFileName = "" Then MsgBox "請設定Email List 輸出檔名(xxx.TXT)", vbCritical, giMsgWarning: Exit Function
        Set DetailFile = FSO.CreateTextFile(EmailListPath & "\" & strTXTFileName)
        
        strTXTFileName = strRptPath & IIf(Right(strRptPath, 1) = "\", "", "\") & strTXTFileName
        Set SampleFile = FSO.OpenTextFile(strTXTFileName)
        
        strSample = SampleFile.ReadAll
        With rsPrtSNOTmp
            .MoveFirst
            Do While Not .EOF
                '附加檔案
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
    Dim MainFile As TextStream      'Email List 檔
    Dim DetailFile As TextStream    'ActtachFile 檔....
    Dim strRptPath As String
    Dim strSample As String         'Sample 檔的內容...
    Dim strRead As String             '轉換後的內容...
        strRptPath = ReadGICMIS1("RptPath")
        If strTXTFileName = "" Then MsgBox "請設定Email List 輸出檔名(xxx.TXT)", vbCritical, giMsgWarning: Exit Function
        strTXTFileName = strRptPath & IIf(Right(strRptPath, 1) = "\", "", "\") & strTXTFileName
        strSubject = Mid(strSubject, 3)
        'Email格式檔
        Set SampleFile = FSO.OpenTextFile(strTXTFileName)
        strSample = SampleFile.ReadAll
        SampleFile.Close
        'EmailList.txt檔
        Set MainFile = FSO.CreateTextFile(EmailListPath & "\EmailList.Txt")
        EmailListPath = IIf(EmailListPath = "", "~", EmailListPath)
        '設定主檔的Head
        MainFile.WriteLine "ContentFileName#"
        MainFile.WriteLine "AttachmentFileName#"
        MainFile.WriteLine "ReferenceFileName#"
        MainFile.WriteLine "Subject#" & strSubject
        
        With rsPrtSNOTmp
            .MoveFirst
            Do While Not .EOF
                '附加檔案
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
            '取條碼字型
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
        
        If blnMsgQuestion Then If vbNo = MsgBox("確認要作廢單據編號: " & strBillNo & " 項目:" & strCitemName & " 該筆資料嗎？", vbQuestion + vbYesNo, gimsgPrompt) Then Exit Function
        Set clsAlter = CreateObject("csAlterCharge4.clsAlterCharge")
        clsAlter.uOwnerName = GetOwner
        blnChk = clsAlter.DeleteSo033(gcnGi, strBillNo, strItem, garyGi(1), , intCancelCode, strCancelName)
        If blnChk Then
            MsgBox "作廢成功！", vbExclamation, "訊息"
        Else
            MsgBox "作廢失敗！請重新操作！", vbInformation, "訊息"
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
        '先找CompCode 及ServiceType 都有的
        strTable = GetOwner & strTable
        
        If strCompCode <> "" And strServiceType <> "" Then
            If Not GetRS(rs, "Select " & strParaField & " From " & strTable & " Where CompCode = " & strCompCode & " And ServiceType " & strSQL) Then Exit Function
            blnFlag = Not rs.EOF
        End If
        If Not blnFlag Then
            '再找CompCode 有的
            If strCompCode <> "" Then
                If Not GetRS(rs, "Select " & strParaField & " From " & strTable & " Where CompCode = " & strCompCode) Then Exit Function
                blnFlag = Not rs.EOF
            End If
            '最後找只要有這個 參數檔的
            If Not blnFlag Then
                If Not GetRS(rs, "Select " & strParaField & " From " & strTable) Then Exit Function
                blnFlag = Not rs.EOF
            End If
        End If
        If blnShowMsg And Not blnFlag Then MsgBox "在系統參數檔[" & strTable & "] 找不到 公司別代號為 [" & strCompCode & "] ,且服務類別為: [" & strServiceType & "]的資料,請查證!!", vbCritical, giMsgWarning
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
        If Len(strValue) = 0 And strShowMsg Then MsgBox "該客戶無服務代號為[" & strServiceType & "之服務!!", vbCritical, giMsgWarning
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
'                    MsgBox "長度過長 !!, 該欄位長度最長為 " & objControl.MaxLength
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
        If strChoose = "" Then MsgBox "沒有Join 條件,請程式設計查證!!", vbCritical, giMsgWarning: Exit Function
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
            strTimeStr = GetString(lngTime \ 3600, 2, giRight) & "小時"
            lngTime = lngTime Mod 3600
        End If
        If lngTime > 60 Then
            strTimeStr = strTimeStr & GetString((lngTime \ 60), 2, giRight) & "分"
            lngTime = lngTime Mod 60
        End If
        strTimeStr = strTimeStr & GetString(lngTime, 2, giRight) & "秒"
        ShowFullTimeStr = strTimeStr
End Function

Public Function MackCABReport(strSourceFile As String, strTarget As String) As Boolean
    On Error GoTo ChkErr
    Dim MCobj As Object
        Set MCobj = CreateObject("CSmc.MCclass")
        '壓縮檔案
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
        'FTP 下載資料
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
        '解壓縮檔案
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
        If rsFmtlist.RecordCount = 0 Then MsgBox "無該表單代號,請查證!!", vbCritical, gimsgPrompt: Exit Function
        If Not GetRS(rsJoin, "Select * From " & rsFmtlist("JoinName") & " Where FmtID = '" & strFmtID & "'", cnReport) Then Exit Function
        
        If Not GetRS(rsFld, "Select 1 as Type , A.* From " & rsFmtlist("FmtFileName") & " A Where FmtID = '" & strFmtID & "' And TableName = '" & rsFmtlist("MainTable") & "'" & _
                " Union All Select * From (Select 2 as Type , A.* From " & rsFmtlist("FmtFileName") & " A Where FmtID = '" & strFmtID & "' And TableName <> '" & rsFmtlist("MainTable") & "' Order By TableName)", cnReport) Then Exit Function
        If Not GetReportFieldStr(rsFld, strField, strTable) Then Exit Function
        
        strField = Mid(strField, 2)
        If blnCreateStru Then
            If Not GetRS(rsTmp, "Select " & strField & " From " & Mid(strTable, 2) & " Where 1=0 ") Then Exit Function
            If Not CreateMDBTable(rsTmp, rsFmtlist("ReportName"), cnOutPutReport) Then Exit Function
        End If
        '全部Join
        If strMainTable = "" Then strMainTable = rsJoin("MainTable") & ""
        strMainTable = UCase(strMainTable)
        
        If rsQuery.RecordCount > 0 Then rsQuery.MoveFirst
        Do While Not rsQuery.EOF
            ' Join 條件所需Table
            If InStr(1, strTable, UCase(rsQuery("TableName"))) = 0 Then strTable = strTable & "," & GetOwner & UCase(rsQuery("TableName")) & " " & UCase(rsQuery("TableName"))
            ' Join 選取條件
            strJoin = strJoin & " And " & rsQuery("SQLQuery")
            strQueryTable = strQueryTable & "," & GetOwner & UCase(rsQuery("TableName")) & " " & UCase(rsQuery("TableName"))
            rsQuery.MoveNext
        Loop
        
        If rsFmtlist.Fields("Type") = 1 Then
            Do While Not rsJoin.EOF
                ' Join 欄位先需Table
                If InStr(1, strTable, UCase(strMainTable)) = 0 Then strTable = strTable & "," & GetOwner & UCase(strMainTable) & " " & UCase(rsJoin("MainTable") & "")
                If InStr(1, strTable, UCase(rsJoin("SecondTable"))) = 0 Then strTable = strTable & "," & GetOwner & UCase(rsJoin("SecondTable")) & " " & UCase(rsJoin("SecondTable"))
                For intLoop = 1 To 5
                    If rsJoin("MainJoinFld" & intLoop) & "" <> "" Then
                        ' Join 欄位所需條件
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
                                ' Join 欄位所需條件
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
                                    ' Join 欄位所需條件
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
            '欄位所需Table
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
            ErrHandle "在開啟 [表單一覽表](FmtList) 的資料時發生錯誤: " & Err.Description & " 這個程式即將關閉。", False, 2, "Form_Load"
            
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
        ' 根據單別至輸出參數檔(SO045)取對應之輸出格式代號及列表機名稱, 設為對應參數之初值
        If cboType.ListIndex < 0 Then Exit Function
        intIndex = cboType.ListIndex + 1
        If intIndex > 5 Then intIndex = 1
        If Not GetSystemPara(rsTmp, strCompCode, strServiceType, "SO045", "BillTab" & intIndex & ",BillPrinter") Then
            Exit Function
        End If
        If rsTmp.EOF Then
           MsgBox "請先設定[輸出參數檔(SO045)", vbCritical, "訊息"
           Exit Function
        Else
           If SetCboText(cboFmtID, rsTmp("BillTab" & intIndex).Value & "", "找不到指定的收費單格式檔") = False Then Exit Function
           If SetCboText(cboPrinterName, GetPrinterName(4, strServiceType), "找不到指定的印表機") = False Then Exit Function
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
        MsgBox strMsg & " [" & strValue & "]！", vbExclamation, "警告"
        SetCboText = True
        If cbo.ListCount > 0 Then
            cbo.ListIndex = 0
        Else
            'Call errHandle("使用的系統參數有誤，此程式即將關閉！", False, 2, "SetCboText")
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
        '2007/10/12 改成也加服務別判斷
        If strServiceType <> "" Then strWhere = " And ServiceType = '" & strServiceType & "'"
        strTranDate = GetRsValue("Select tranDate from " & GetOwner & "So062 Where CompCode= " & strCompCode & strWhere & " And Type=1 Order By TranDate Desc") & ""
        
        '收費參數檔（SO043)，取收費日登錄期限<Para6>，供後續使用！
        intPara6 = Val(GetSystemParaItem("Para6", strCompCode, strServiceType, "SO043") & "")
    
        If strCloseDate = "" Then
            If vbNo = MsgBox("本欄位是否為空值！", vbExclamation + vbYesNo, "訊息！") Then Exit Function
        Else
            If Not IsDate(strCloseDate) Then
                MsgBox "日期不合法！", vbExclamation, "訊息！"
                Exit Function
            End If
            
            If CDate(strCloseDate) > RightDate Then MsgBox "此日期超過今天日期！", vbExclamation, "訊息！": Exit Function
            
            If (RightDate - CDate(strCloseDate)) > intPara6 Then
                MsgBox "此日期已超過系統設定的安全期限！", vbExclamation, "訊息！":  Exit Function
            End If
            If CDate(strCloseDate) <= IIf(strTranDate <> "", CDate(strTranDate), 0) Then
                If intDayCut = 1 And CDate(strCloseDate) = IIf(strTranDate <> "", CDate(strTranDate), 0) Then ChkCloseDate = True: Exit Function
                MsgBox strTranDate & "之前已做過日結，不可使用之前日期入帳", vbExclamation, "訊息！"
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

'第一個傳條件說明,第二個傳條件內容,第三個傳是否加分號(;) 93/11/24
'Ex: GetSelectConditionStr("實收日期","93/11/24~93/11/31",False)
'結果為: "實收日期:93/11/24~93/11/31;"
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

'檢核資料的一致性,
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
                MsgBox "該筆資料正被其他人異動中,請查證!!", vbCritical, giMsgWarning
                cn.RollbackTrans
                GoTo lClose
            End If
        End If
        cn.RollbackTrans
        On Error GoTo ChkErr
        '原始資料為EOF無法比對,不做判斷
        If Not rsOld.EOF Then
            If Not GetRS(rsDB, strSQL) Then Exit Function
            '資料庫不存在,不做判斷
            If rsDB.RecordCount > 0 Then
                For intLoop = 0 To rsDB.Fields.Count - 1
                    '判斷欄位內容不同則表示資料庫已經被異動
                    If rsDB(intLoop) & "" <> rsOld(rsDB(intLoop).Name).OriginalValue & "" Then
                        strErrorMsg = "該筆資料已經被其他人異動過,請查證!!"
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
'            1.正常=>InstDate is not null and PRDate is null
'            2.停機=>InstDate is not null and PRDate is null And FaciStatusCode = 3
'            3.拆機/未取回=>PRDate is not null and GetDate is null
'            4.拆機/取回=>PRDate is not null and GetDate is not null
'            5.無:InstDate is null and PrDate is null
            If GetSystemParaItem("FaciAuthority", gCompCode, gServiceType, "SO042") = 2 Then
                strPRField = "GetDate"
            Else
                strPRField = "PRDate"
            End If
            If .Fields("InstDate") & "" <> "" And .Fields(strPRField) & "" = "" Then
                If .Fields("FaciStatusCode") & "" = 3 Then
                    FaciWipStatus = "停機"
                Else
                    If StrToDate(.Fields("StopTime") & "", True) > StrToDate(.Fields("ReInstTime") & "", True) Then
                        FaciWipStatus = "停機"
                    Else
                        FaciWipStatus = "正常"
                    End If
                End If
            ElseIf .Fields(strPRField) & "" <> "" Then
                FaciWipStatus = "拆機"
                If .Fields("GetDate") & "" = "" Then
                    FaciWipStatus = FaciWipStatus & "/未取回"
                End If
            Else
                FaciWipStatus = "無"
            End If
            If .Fields("GetDate") & "" <> "" Then
                FaciWipStatus = FaciWipStatus & "/取回"
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
'            開通狀態
'            1.未開機=>CMOpenDate is null and CMCloseDate is null
'            2.開機=>CMOpenDate > CMCloseDate or CMOpendate is not null and CMCloseDate is null
'            3.關機=>CMOpenDate < CMCloseDate
'            4.開機/停權=>(CMOpenDate > CMCloseDate or CMOpendate is not null and CMCloseDate is null) and
'            (disableaccount> enableAccount or (disableaccount is not null and enableAccount is null)
'            5.關機/停權=>(CMOpenDate < CMCloseDate) and
'            (disableaccount> enableAccount or (disableaccount is not null and enableAccount is null)
            
            If StrToDate(.Fields("CMOpenDate") & "", True, True) > StrToDate(.Fields("CMCloseDate") & "", True, True) Then
                FaciWipStatus = "開機"
            ElseIf StrToDate(.Fields("CMOpenDate") & "", True, True) < StrToDate(.Fields("CMCloseDate") & "", True, True) Then
                FaciWipStatus = "關機"
            ElseIf .Fields("CMOpenDate") & "" = "" And .Fields("CMCloseDate") & "" = "" Then
                FaciWipStatus = "未開機"
            Else
                FaciWipStatus = "無"
            End If
            If StrToDate(.Fields("disableaccount") & "", True, True) > StrToDate(.Fields("enableAccount") & "", True, True) Then
                FaciWipStatus = FaciWipStatus & "/停權"
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
            If MsgBox("確定要刪除這筆資料？", vbQuestion + vbYesNo, "確認") = vbYes Then
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
        '若大於今日，則警告："此日期超過今天日期"
        If Not IsDate(Format(gdaRealDate.GetValue, "####/##/##")) Then
            MsgBox "日期錯誤！", , "訊息！"
            Cancel = True
            Exit Function
        End If
        If CDate(gdaRealDate.Text) > RightDate Then MsgBox "此日期超過今天日期": Cancel = True: Exit Function
        
        '若（今日-本欄位值- >　＜Para6>)，則警告："此日期已超過系統設的安全期限"
        If DateDiff("d", CDate(gdaRealDate.Text), RightDate) > XintPara6 Then
            MsgBox "此日期已超過系統設定的安全期限", vbExclamation, "訊息！"
            Cancel = True
            Exit Function
        End If
        
        '若小於<TransDate>，則錯誤："<TransDate>之前已做過日結，不可使用之前日期入帳",Focus停留在本欄位！
        If DateDiff("d", CDate(XstrTranDate), CDate(gdaRealDate.Text)) <= 0 Then
            If intDayCut = 1 And DateDiff("d", CDate(XstrTranDate), CDate(gdaRealDate.Text)) = 0 Then Exit Function
            MsgBox XstrTranDate & "之前已做過日結，不可使用之前日期入帳", vbExclamation, "訊息！"
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
                MsgBox "無此手開單號,請查證!!", vbExclamation, gimsgPrompt
                Exit Function
            Else
                If rsTmp("Status") = 0 Then
                    MsgBox "此手開單號已作廢不可使用,請查證!!", vbExclamation, gimsgPrompt
                    Exit Function
                End If
            End If
            
            If strBillNo = "" Then
                strBillSQL = " And BillNo is not null"
            Else
                strBillSQL = " And BillNo <> '" & strBillNo & "'"
            End If
            
            If GetRsValue("Select Count(*) From " & GetOwner & "SO127 Where PaperNum = '" & strManualNo & "' And CompCode = " & strCompCode & strBillSQL) > 0 Then
                MsgBox "該手開單已有收費資料,請查證!!", vbExclamation, gimsgPrompt
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
                If intPFlag1 = 1 Then    '小於期數
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
        '為第2層優惠的頻道收費項目 或 本身為第2層優惠
        If Val(rs019("SecendDiscount") & "") > 0 Then
            ChkHave2Discount = True
        End If
End Function

Public Function Chk2Discount(ByVal strCitemCode As String) As Boolean
    Dim rs019 As New ADODB.Recordset
        If Val(GetSystemParaItem("SecondDiscount", gCompCode, "", "SO041") & "") = 0 Then Exit Function
        If Not GetRS(rs019, "Select RefNo From " & GetOwner & "CD019 Where CodeNo = " & strCitemCode) Then Exit Function
        '為第2層優惠
        If Val(rs019("RefNo") & "") = 19 Then
            Chk2Discount = True
            Exit Function
        End If
        If Not GetRS(rs019, "Select RefNo From " & GetOwner & "CD019 Where ReturnCode = " & strCitemCode) Then Exit Function
        '為補收第2層優惠
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
        '本身為第3層優惠
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


'呼叫出帳後端程式 By Kin 2008/08/06
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


