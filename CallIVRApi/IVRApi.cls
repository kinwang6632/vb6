VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IVRApi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public gcnGi As New ADODB.Connection
'Private objWebPool As Object ' Connection Pool 物件
Private cmd As New ADODB.Connection
Private FSO As New FileSystemObject
Private strmTextFile1 As TextStream
Dim strConn As String

Public Function Xdecrypt(EncryptString As String, Optional EncryptKey As String = "")
    Xdecrypt = CreateObject("Encryption.Password").Decrypt(EncryptString, EncryptKey)
End Function

Public Function Executing(ByVal strPara As String) As String
Dim strTextData As String

'    Set objWebPool = CreateObject("WebPool.PoolManager") ' 建立 Connection Pool 物件
'    Set gcnGi = objWebPool.GetConnection
    Dim arrString() As String
    Dim intFor As Integer
    Dim strPara1
    Dim strOwner As String
    Dim ErrNum As String, ErrDesc As String, ErrSrc As String
    On Error GoTo ChkErr
    strTextData = FSO.OpenTextFile("C:\WINDOWS\system32\DBInfo.ini").ReadAll
    strConn = strTextData
    strConn = Xdecrypt(strConn, "CS")

    With cmd
         If .State <= 0 Then
            .CursorLocation = adUseServer
            .Open strConn & ";Persist Security Info=True"
        End If
    End With
    arrString = Split(Mid(strPara, 1, Len(strPara) - 1), ",")
    strOwner = GetOwner(arrString(1))
    If Not UBound(arrString) Then
    Select Case arrString(0)
            Case 1
                strPara = CallSP_IVR01_CheckID(strOwner, arrString(1), arrString(2))
            Case 2
                strPara = CallSP_IVR02_CheckCustID(strOwner, arrString(1), arrString(2))
            Case 3
                strPara = CallSP_IVR03_CheckPhone(strOwner, arrString(1), arrString(2), arrString(3))
            Case 4
                strPara = CallSP_IVR04_CheckAlarm(strOwner, arrString(1), arrString(2))
            Case 5
                strPara = CallSP_IVR05_GenMaintenanceWo(strOwner, arrString(1), arrString(2), arrString(3), arrString(4))
            Case 6
                strPara = CallSP_IVR06_QueryCheckID(strOwner, arrString(1), arrString(2))
            Case 7
                strPara = CallSP_IVR07_QueryCheckCustID(strOwner, arrString(1), arrString(2))
            Case 8
                strPara = CallSP_IVR08_QueryCheckPhone(strOwner, arrString(1), arrString(2), arrString(3))
            Case 9
                strPara = CallSP_IVR09_CancelWo(strOwner, arrString(1), arrString(2))
            Case 10
                strPara = CallSP_IVR10_WoTime(strOwner, arrString(1), arrString(2))
        End Select
    End If
    gcnGi.Cancel
    Set gcnGi = Nothing

    cmd.Cancel
    Set cmd = Nothing
 Executing = strPara    '"AA;" & strConn
 Exit Function
ChkErr:
    ErrNum = Err.Number
    ErrDesc = Err.Description
    ErrSrc = Err.Source
    Executing = ErrNum & ";" & ErrDesc & ";" & ErrSrc
End Function

Public Function RPxx(strValue As String) As String
  On Error Resume Next
    strValue = Trim(Replace(strValue, Chr(13), "", 1, , vbTextCompare))
    strValue = Trim(Replace(strValue, Chr(10), "", 1, , vbTextCompare))
    RPxx = Trim(Replace(strValue, Chr(9), "", 1, , vbTextCompare))
End Function

Public Function GetOwner(intComCode As String)
    Dim strOwner As String, strOwnerName As String
'    Set objWebPool = CreateObject("WebPool.PoolManager") ' 建立 Connection Pool 物件
'    Set gcnGi = objWebPool.GetConnection
    strOwner = RPxx(cmd.Execute("SELECT TABLEOWNER FROM CD039 Where CODENO=" & intComCode).GetString(2, 1, "", "", "") & "")
    If strOwnerName = "" Then strOwnerName = strOwner
    If strOwnerName <> "" Then GetOwner = strOwnerName & "."

End Function

'01
Public Function CallSP_IVR01_CheckID(strOwner As String, intCompCode As String, strCustomerID As String) As String
    Dim strCCBOP1 As String, strCCBOP2 As String, strCCBOP3 As String, strCustomerNo As String
    Dim strIVR As String
    Dim comIVR As New ADODB.Command
    Dim lngRet As Long
    With comIVR
        .Parameters.Append .CreateParameter("CompCode", adVarNumeric, adParamInput, 2000, intCompCode)
        .Parameters.Append .CreateParameter("CustomerID", adVarChar, adParamInput, 100, strCustomerID)
        .Parameters.Append .CreateParameter("CCBOP1", adVarChar, adParamOutput, 100)
        .Parameters.Append .CreateParameter("CCBOP2", adVarChar, adParamOutput, 600)
        .Parameters.Append .CreateParameter("CCBOP3", adVarChar, adParamOutput, 100)
        .Parameters.Append .CreateParameter("CustomerNo", adVarChar, adParamOutput, 100)

    Set .ActiveConnection = cmd
        .CommandText = strOwner & "SP_IVR01_CheckID"
        .CommandType = adCmdStoredProc
        .Execute
    strCCBOP1 = .Parameters("CCBOP1").Value & ","
    strCCBOP2 = .Parameters("CCBOP2").Value & ","
    strCCBOP3 = .Parameters("CCBOP3").Value & ","
    strCustomerNo = .Parameters("CustomerNo").Value & ""
    End With
'    cmd.Execute ("select EMCNTP.SP_IVR01_CheckID(10,'A100018816',strCCBOP1,strCCBOP2,strCCBOP3,strCustomerNo) from Dual")
'    With comIVR
'        Set .ActiveConnection = cmd
'        .CommandType = 1
'        .CommandText = "SELECT SP_IVR01_CheckID(10,'A121179690',strCCBOP1,strCCBOP2,strCCBOP3,strCustomerNo) FROM DUAL"
'        .Execute lngRet, Array(intCompCode, strCustomerID, strCCBOP1, strCCBOP2, strCCBOP3, strCustomerNo), 1
'    End With
    strIVR = strCCBOP1 & strCCBOP2 & strCCBOP3 & strCustomerNo
    CallSP_IVR01_CheckID = strIVR & ";"
  Exit Function
'ER:
'    ErrSub "Utility1", "CallSP_IVR01_CheckID"
End Function

'02
Public Function CallSP_IVR02_CheckCustID(strOwner As String, intCompCode As String, strCustomerID As String) As String
    Dim strCCBOP1 As String, strCCBOP2 As String, strCCBOP3 As String, strCustomerNo As String
    Dim strIVR As String
    Dim comIVR As New ADODB.Command
    
    With comIVR
        .Parameters.Append .CreateParameter("CompCode", adVarNumeric, adParamInput, 2000, intCompCode)
        .Parameters.Append .CreateParameter("CustomerID", adVarChar, adParamInput, 100, strCustomerID)
        .Parameters.Append .CreateParameter("CCBOP1", adVarChar, adParamOutput, 100)
        .Parameters.Append .CreateParameter("CCBOP2", adVarChar, adParamOutput, 600)
        .Parameters.Append .CreateParameter("CCBOP3", adVarChar, adParamOutput, 100)

    Set .ActiveConnection = cmd
        .CommandText = strOwner & "SP_IVR02_CheckCustID"
        .CommandType = adCmdStoredProc
        .Execute
    strCCBOP1 = .Parameters("CCBOP1").Value & ","
    strCCBOP2 = .Parameters("CCBOP2").Value & ","
    strCCBOP3 = .Parameters("CCBOP3").Value & ""
    End With
    strIVR = strCCBOP1 & strCCBOP2 & strCCBOP3
    CallSP_IVR02_CheckCustID = strIVR & ";"
  Exit Function
'ER:
'    ErrSub "Utility1", "CallSP_IVR01_CheckID"
End Function

'03
Public Function CallSP_IVR03_CheckPhone(strOwner As String, intCompCode As String, strPhoneNo As String, strFloor As String) As String
    Dim strCCBOP1 As String, strCCBOP2 As String, strCCBOP3 As String
    Dim strCustomerNo As String, strCustomerName As String
    Dim strIVR As String
    Dim comIVR As New ADODB.Command
    
    With comIVR
        .Parameters.Append .CreateParameter("CompCode", adVarNumeric, adParamInput, 2000, intCompCode)
        .Parameters.Append .CreateParameter("PhoneNo", adVarChar, adParamInput, 100, strPhoneNo)
        .Parameters.Append .CreateParameter("Floor", adVarChar, adParamInput, 100, strFloor)
        .Parameters.Append .CreateParameter("CCBOP1", adVarChar, adParamOutput, 100)
        .Parameters.Append .CreateParameter("CCBOP2", adVarChar, adParamOutput, 600)
        .Parameters.Append .CreateParameter("CCBOP3", adVarChar, adParamOutput, 100)
        .Parameters.Append .CreateParameter("CustomerNo", adVarChar, adParamOutput, 100)
        .Parameters.Append .CreateParameter("CustomerName", adVarChar, adParamOutput, 100)
    Set .ActiveConnection = cmd
        .CommandText = strOwner & "SP_IVR03_CheckPhone"
        .CommandType = adCmdStoredProc
        .Execute
    strCCBOP1 = .Parameters("CCBOP1").Value & ","
    strCCBOP2 = .Parameters("CCBOP2").Value & ","
    strCCBOP3 = .Parameters("CCBOP3").Value & ","
    strCustomerNo = .Parameters("CustomerNo").Value & ","
    strCustomerName = .Parameters("CustomerName").Value & ""
    End With
    strIVR = strCCBOP1 & strCCBOP2 & strCCBOP3 & strCustomerNo & strCustomerName
    CallSP_IVR03_CheckPhone = strIVR & ";"
  Exit Function
'ER:
'    ErrSub "Utility1", "CallSP_IVR01_CheckID"
End Function

'04
Public Function CallSP_IVR04_CheckAlarm(strOwner As String, intCompCode As String, strCustomerID As String) As String
    Dim strCCBOP1 As String, strCCBOP2 As String
    Dim strCustomerNo As String, strCustomerName As String
    Dim strIVR As String
    Dim comIVR As New ADODB.Command
    
    With comIVR
        .Parameters.Append .CreateParameter("CompCode", adVarNumeric, adParamInput, 2000, intCompCode)
        .Parameters.Append .CreateParameter("CustomerID", adVarChar, adParamInput, 100, strCustomerID)
        .Parameters.Append .CreateParameter("CCBOP1", adVarChar, adParamOutput, 100)
        .Parameters.Append .CreateParameter("CCBOP2", adVarChar, adParamOutput, 600)
    Set .ActiveConnection = cmd
        .CommandText = strOwner & "SP_IVR04_CheckAlarm"
        .CommandType = adCmdStoredProc
        .Execute
    strCCBOP1 = .Parameters("CCBOP1").Value & ","
    strCCBOP2 = .Parameters("CCBOP2").Value & ""
    End With
    strIVR = strCCBOP1 & strCCBOP2
    CallSP_IVR04_CheckAlarm = strIVR & ";"
  Exit Function
'ER:
'    ErrSub "Utility1", "CallSP_IVR01_CheckID"
End Function

'05
Public Function CallSP_IVR05_GenMaintenanceWo(strOwner As String, intCompCode As String, strCustomerNo As String, strMaintainDate As String, strMaintainTime As String) As String
    Dim strCCBOP1 As String
    Dim strIVR As String
    Dim comIVR As New ADODB.Command
    
    With comIVR
        .Parameters.Append .CreateParameter("CompCode", adVarNumeric, adParamInput, 2000, intCompCode)
        .Parameters.Append .CreateParameter("CustomerNo", adVarChar, adParamInput, 100, strCustomerNo)
        .Parameters.Append .CreateParameter("MaintainDate", adVarChar, adParamInput, 100, strMaintainDate)
        .Parameters.Append .CreateParameter("MaintainTime", adVarChar, adParamInput, 100, strMaintainTime)
        .Parameters.Append .CreateParameter("CCBOP1", adVarChar, adParamOutput, 100)
    Set .ActiveConnection = cmd
        .CommandText = strOwner & "SP_IVR05_GenMaintenanceWo"
        .CommandType = adCmdStoredProc
        .Execute
    strCCBOP1 = .Parameters("CCBOP1").Value & ""
    
    End With
    strIVR = strCCBOP1
    CallSP_IVR05_GenMaintenanceWo = strIVR & ";"
  Exit Function
'ER:
'    ErrSub "Utility1", "CallSP_IVR01_CheckID"
End Function

'06
Public Function CallSP_IVR06_QueryCheckID(strOwner As String, intCompCode As String, strCustomerID As String) As String
    Dim strCCBOP1 As String, strCCBOP2 As String
    Dim strCustomerNo As String
    Dim strMaintainDate As String, strMaintainTime As String
    Dim strIVR As String
    Dim comIVR As New ADODB.Command
    
    With comIVR
        .Parameters.Append .CreateParameter("CompCode", adVarNumeric, adParamInput, 2000, intCompCode)
        .Parameters.Append .CreateParameter("CustomerID", adVarChar, adParamInput, 100, strCustomerID)
        .Parameters.Append .CreateParameter("CCBOP1", adVarChar, adParamOutput, 100)
        .Parameters.Append .CreateParameter("CCBOP2", adVarChar, adParamOutput, 600)
        .Parameters.Append .CreateParameter("CustomerNo", adVarChar, adParamOutput, 100)
        .Parameters.Append .CreateParameter("MaintainDate", adVarChar, adParamOutput, 100)
        .Parameters.Append .CreateParameter("MaintainTime", adVarChar, adParamOutput, 100)
        
    Set .ActiveConnection = cmd
        .CommandText = strOwner & "SP_IVR06_QueryCheckID"
        .CommandType = adCmdStoredProc
        .Execute
    strCCBOP1 = .Parameters("CCBOP1").Value & ","
    strCCBOP2 = .Parameters("CCBOP2").Value & ","
    strCustomerNo = .Parameters("CustomerNo").Value & ","
    strMaintainDate = .Parameters("MaintainDate").Value & ","
    strMaintainTime = .Parameters("MaintainTime").Value & ""
'    strMsg = .Parameters("p_RetMsg").Value & ""
    End With
    strIVR = strCCBOP1 & strCCBOP2 & strCustomerNo & strMaintainDate & strMaintainTime
    CallSP_IVR06_QueryCheckID = strIVR & ";"
  Exit Function
'ER:
'    ErrSub "Utility1", "CallSP_IVR01_CheckID"
End Function

'07
Public Function CallSP_IVR07_QueryCheckCustID(strOwner As String, intCompCode As String, strCustomerNo As String) As String
    Dim strCCBOP1 As String, strCCBOP2 As String, strCCBOP3 As String
    Dim strMaintainDate As String, strMaintainTime As String
    Dim strIVR As String
    Dim comIVR As New ADODB.Command
    
    With comIVR
        .Parameters.Append .CreateParameter("CompCode", adVarNumeric, adParamInput, 2000, intCompCode)
        .Parameters.Append .CreateParameter("CustomerNo", adVarChar, adParamInput, 100, strCustomerNo)
        .Parameters.Append .CreateParameter("CCBOP1", adVarChar, adParamOutput, 100)
        .Parameters.Append .CreateParameter("CCBOP2", adVarChar, adParamOutput, 600)
        .Parameters.Append .CreateParameter("CCBOP3", adVarChar, adParamOutput, 100)
        .Parameters.Append .CreateParameter("MaintainDate", adVarChar, adParamOutput, 100)
        .Parameters.Append .CreateParameter("MaintainTime", adVarChar, adParamOutput, 100)
        
    Set .ActiveConnection = cmd
        .CommandText = strOwner & "SP_IVR07_QueryCheckCustID"
        .CommandType = adCmdStoredProc
        .Execute
    strCCBOP1 = .Parameters("CCBOP1").Value & ","
    strCCBOP2 = .Parameters("CCBOP2").Value & ","
    strCCBOP3 = .Parameters("CCBOP3").Value & ","
    strMaintainDate = .Parameters("MaintainDate").Value & ","
    strMaintainTime = .Parameters("MaintainTime").Value & ""
    
    End With
    strIVR = strCCBOP1 & strCCBOP2 & strCCBOP3 & strMaintainDate & strMaintainTime
    CallSP_IVR07_QueryCheckCustID = strIVR & ";"
  Exit Function
'ER:
'    ErrSub "Utility1", "CallSP_IVR01_CheckID"
End Function

'08
Public Function CallSP_IVR08_QueryCheckPhone(strOwner As String, intCompCode As String, strPhoneNo As String, strFloor As String) As String
    Dim strCCBOP1 As String, strCCBOP2 As String
    Dim strCustomerNo As String
    Dim strMaintainDate As String, strMaintainTime As String, strCustomerName As String
    Dim strIVR As String
    Dim comIVR As New ADODB.Command
    
    With comIVR
        .Parameters.Append .CreateParameter("CompCode", adVarNumeric, adParamInput, 2000, intCompCode)
        .Parameters.Append .CreateParameter("PhoneNo", adVarChar, adParamInput, 100, strPhoneNo)
        .Parameters.Append .CreateParameter("Floor", adVarChar, adParamInput, 100, strFloor)
        .Parameters.Append .CreateParameter("CCBOP1", adVarChar, adParamOutput, 100)
        .Parameters.Append .CreateParameter("CCBOP2", adVarChar, adParamOutput, 600)
        .Parameters.Append .CreateParameter("CustomerNo", adVarChar, adParamOutput, 100)
        .Parameters.Append .CreateParameter("CustomerName", adVarChar, adParamOutput, 100)
        .Parameters.Append .CreateParameter("MaintainDate", adVarChar, adParamOutput, 100)
        .Parameters.Append .CreateParameter("MaintainTime", adVarChar, adParamOutput, 100)
        
    Set .ActiveConnection = cmd
        .CommandText = strOwner & "SP_IVR08_QueryCheckPhone"
        .CommandType = adCmdStoredProc
        .Execute
    strCCBOP1 = .Parameters("CCBOP1").Value & ","
    strCCBOP2 = .Parameters("CCBOP2").Value & ","
    strCustomerNo = .Parameters("CustomerNo").Value & ","
    strCustomerName = .Parameters("CustomerName").Value & ","
    strMaintainDate = .Parameters("MaintainDate").Value & ","
    strMaintainTime = .Parameters("MaintainTime").Value & ""
    
    End With
    strIVR = strCCBOP1 & strCCBOP2 & strCustomerNo & strCustomerName & strMaintainDate & strMaintainTime
    CallSP_IVR08_QueryCheckPhone = strIVR & ";"
  Exit Function
'ER:
'    ErrSub "Utility1", "CallSP_IVR01_CheckID"
End Function

'09
Public Function CallSP_IVR09_CancelWo(strOwner As String, intCompCode As String, strCustomerNo As String) As String
    Dim strCCBOP1 As String
    Dim strIVR As String
    Dim comIVR As New ADODB.Command
    
    With comIVR
        .Parameters.Append .CreateParameter("CompCode", adVarNumeric, adParamInput, 2000, intCompCode)
        .Parameters.Append .CreateParameter("CustomerNo", adVarChar, adParamInput, 100, strCustomerNo)
        .Parameters.Append .CreateParameter("CCBOP1", adVarChar, adParamOutput, 100)
        
    Set .ActiveConnection = cmd
        .CommandText = strOwner & "SP_IVR09_CancelWo"
        .CommandType = adCmdStoredProc
        .Execute
    strCCBOP1 = .Parameters("CCBOP1").Value & ""
    
    End With
    strIVR = strCCBOP1
    CallSP_IVR09_CancelWo = strIVR & ";"
  Exit Function
'ER:
'    ErrSub "Utility1", "CallSP_IVR01_CheckID"
End Function

'10
Public Function CallSP_IVR10_WoTime(strOwner As String, intCompCode As String, strCustomerNo As String) As String
    Dim strCCBOP1 As String
    Dim strCCBOP2 As String
    Dim strIVR As String
    Dim comIVR As New ADODB.Command
    
    With comIVR
        .Parameters.Append .CreateParameter("CompCode", adVarNumeric, adParamInput, 2000, intCompCode)
        .Parameters.Append .CreateParameter("CustomerNo", adVarChar, adParamInput, 100, strCustomerNo)
        .Parameters.Append .CreateParameter("CCBOP1", adVarChar, adParamOutput, 100)
        .Parameters.Append .CreateParameter("CCBOP2", adVarChar, adParamOutput, 60)
        
    Set .ActiveConnection = cmd
        .CommandText = strOwner & "SP_IVR10_WoTime"
        .CommandType = adCmdStoredProc
        .Execute
    strCCBOP1 = .Parameters("CCBOP1").Value & ","
    strCCBOP2 = .Parameters("CCBOP2").Value & ""
    
    End With
    strIVR = strCCBOP1 & Trim(strCCBOP2)
    CallSP_IVR10_WoTime = strIVR & ";"
  Exit Function
'ER:
'    ErrSub "Utility1", "CallSP_IVR01_CheckID"
End Function

