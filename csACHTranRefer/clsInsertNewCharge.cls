VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInsertNewCharge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const FormName = "clsInsertNewCharge"
Dim strErrorCode As String
Dim strErrorMsg As String

Public Function InsertChargeField(ByRef rsSo033 As ADODB.Recordset, _
        ByVal lngCustId As Long, ByVal lngCompCode As Long, ByVal strServiceType As String, _
        ByVal strBillNo As String, ByVal lngItem As Long, _
        ByVal strCitemCode As String, ByVal strShouldDate As String, _
        Optional ByVal strClctEn As String, Optional ByVal strClctName As String, _
        Optional ByVal strFaciSeqNo As String = "", Optional ByVal strFaciSNo As String = "", _
        Optional ByVal blnAddNew As Boolean = True, Optional ByVal blnShowMsg As Boolean = True) As Boolean

    On Error GoTo ChkErr
    Dim rsCD013 As New ADODB.Recordset
    Dim rsSo001 As New ADODB.Recordset
    Dim rsSO014 As New ADODB.Recordset
    Dim rsX As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim lngValueC As Long
    Dim lngPeriod As Long
    Dim pRealStopDate As String
    Dim pShouldAmt As Long
    Dim pRetMsg As String
    Dim Ret As Integer
    Dim lngAmtOfCD019 As Long
    Dim intRealPeriod As Integer
    Dim strPro As String
        strPro = "取收費資料"
        If rsSo033.State = adStateClosed Then
            If Not GetRS(rsSo033, "Select * From " & GetOwner & "SO033 Where RowId = ''", , adUseClient, adOpenKeyset, adLockOptimistic) Then GoTo lExit
            Set rsSo033.ActiveConnection = Nothing
        End If
        strPro = "收費項目代碼"
        If Not GetRS(rsX, "Select CodeNo,Description,PeriodFlag,Amount,Period,Sign From " & GetOwner & "CD019 Where CodeNo = " & strCitemCode, gcnGi) Then GoTo lExit
        If rsX.RecordCount = 0 Then InsertChargeField = True: Exit Function
        strPro = "未繳費代碼"
        If Not GetRS(rsCD013, "Select CodeNo,Description From " & GetOwner & "CD013 Where RefNo = 1 ", gcnGi) Then InsertChargeField = True: GoTo lExit
        lngPeriod = Val(rsX("PeriodFlag").Value & "")
        lngAmtOfCD019 = Val(rsX("Amount").Value & "")
        
        strPro = "客戶主檔"
        If Not GetRS(rsSo001, "Select InstAddrNO,MduId,ServCode,ClctAreaCode,ClassCode1 From " & GetOwner & "SO001 Where CustId = " & lngCustId & " And CompCode = " & lngCompCode) Then GoTo lExit
        strPro = "地址資料檔"
        If Not GetRS(rsSO014, "Select StrtCode,AreaCode,ClctEn,ClctName From " & GetOwner & "SO014 Where AddrNo = " & rsSo001("InstAddrNo") & " And CompCode = " & lngCompCode) Then GoTo lExit
        strPro = "回填收費資料"
        With rsSo033
            If blnAddNew Then
                .AddNew
                ' BY戶BY機 93/11/11
                If strFaciSeqNo <> "" Then
                    .Fields("FaciSeqNo") = NoZero(strFaciSeqNo)
                    .Fields("FaciSNo") = NoZero(strFaciSNo)
                Else
                    .Fields("FaciSeqNo") = lngCustId
                    .Fields("FaciSNo") = Null
                End If
            End If
            .Fields("CustID") = lngCustId
            .Fields("BillNo") = strBillNo
            .Fields("Item") = lngItem
            .Fields("CitemCode") = NoZero(rsX("CodeNo"))
            .Fields("CitemName") = NoZero(rsX("Description"))
            .Fields("ShouldDate") = strShouldDate
            .Fields("RealAmt") = 0
            '.Fields("RealPeriod") = 0
            .Fields("CMCode") = 1
            .Fields("CMName") = GetRsValue("Select Description From " & GetOwner & "CD031 Where CodeNo = 1") & ""
            .Fields("PTCode") = 1
            .Fields("PTName") = GetRsValue("Select Description From " & GetOwner & "CD032 Where CodeNo = 1") & ""
            .Fields("CreateTime") = Format(RightNow, "YYYY/MM/DD HH:MM:SS")
            .Fields("CreateEn") = garyGi(0)
            .Fields("CompCode") = lngCompCode
            .Fields("AddrNo") = NoZero(rsSo001("InstAddrNO"))
            .Fields("StrtCode") = NoZero(rsSO014("StrtCode"))
            .Fields("MduId") = NoZero(rsSo001("MduId"))
            .Fields("ServCode") = NoZero(rsSo001("ServCode"))
            .Fields("ClctAreaCode") = NoZero(rsSo001("ClctAreaCode"))
            .Fields("AreaCode") = NoZero(rsSO014("AreaCode"))
            .Fields("OldClctEn") = NoZero(strClctEn)
            .Fields("OldClctName") = NoZero(strClctName)
            If strClctEn = "" Then
                .Fields("ClctEn") = NoZero(rsSO014("ClctEn"))
                .Fields("ClctName") = NoZero(rsSO014("ClctName"))
            Else
                .Fields("ClctEn") = NoZero(strClctEn)
                .Fields("ClctName") = NoZero(strClctName)
            End If
            .Fields("CancelFlag") = 0
            .Fields("ClassCode") = NoZero(rsSo001("ClassCode1"))
            .Fields("UPDTIME") = GetDTString(RightNow)
            .Fields("UPDEN") = garyGi(1)
            .Fields("UCCode") = NoZero(rsCD013("CodeNo"))  '未收原因
            .Fields("UCName") = NoZero(rsCD013("Description"))
            .Fields("ServiceType") = strServiceType
            .Fields("Quantity") = 0
            '.Fields("Type") = 0
            '判斷是否為週期性費用
            If lngPeriod > 0 Then
                If Val(.Fields("RealPeriod") & "") = 0 Then
                    lngPeriod = Val(rsX("Period") & "")
                Else
                    lngPeriod = Val(.Fields("RealPeriod") & "")
                End If
                lngValueC = SF_GETAMOUNT3(lngCustId, rsX("CodeNo"), lngPeriod, strServiceType, Format(.Fields("ShouldDate").Value, "YYYYMMDD"), lngCompCode)
                If lngValueC < 0 Then GoTo lExit
                .Fields("OldPeriod").Value = lngPeriod
                .Fields("OldAmt").Value = lngValueC
                .Fields("ShouldAmt").Value = lngValueC
                .Fields("OldStartDate").Value = .Fields("ShouldDate").Value
                
                Ret = SFGetAmount(False, 0, lngCustId, rsX("CodeNo").Value, lngPeriod, _
                    Format(rsSo033.Fields("OldStartDate").Value, "YYYYMMDD"), strServiceType, lngCompCode, pRealStopDate, pShouldAmt, pRetMsg, intRealPeriod)
                If Ret <> 0 Then
                    If blnShowMsg Then MsgBox "執行 SF_GETAMOUNT 時發生錯誤: " & pRetMsg & " 這個程式即將關閉。", vbCritical, gimsgWarning
                    .CancelUpdate
                    strPro = "取後端費率資料"
                    GoTo lExit
                End If
                
                If .Fields("OldAmt") = 0 Then .Fields("OldAmt") = pShouldAmt
                If pShouldAmt > 0 Then
                    .Fields("ShouldAmt").Value = pShouldAmt
                    .Fields("Amt").Value = pShouldAmt
                End If
                .Fields("OldStopDate").Value = Left(pRealStopDate, 4) & "/" & Mid(pRealStopDate, 5, 2) & "/" & Right(pRealStopDate, 2)
                .Fields("RealStartDate").Value = .Fields("OldStartDate").Value
                .Fields("RealStopDate").Value = Left(pRealStopDate, 4) & "/" & Mid(pRealStopDate, 5, 2) & "/" & Right(pRealStopDate, 2)
                .Fields("RealPeriod").Value = intRealPeriod
            Else
                .Fields("OldAmt") = Val(rsX("Sign") & Abs(lngAmtOfCD019))
                .Fields("OldPeriod") = 0
                .Fields("RealPeriod") = 0
                .Fields("Amt").Value = .Fields("OldAmt")
                .Fields("ShouldAmt") = .Fields("OldAmt")
            End If
            .Update
        End With
        On Error Resume Next
        Call CloseRecordset(rsCD013)
        Call CloseRecordset(rsX)
        Call CloseRecordset(rsTmp)
        
        InsertChargeField = True
lExit:
        strErrorCode = -1
        strErrorMsg = strPro
        
    Exit Function
ChkErr:
    strErrorCode = Err.Number
    strErrorMsg = strPro & Err.Description
    If blnShowMsg Then ErrSub FormName, "InsertChargeField"
    Err.Number = 0
End Function

Public Function CreditInsertTo033(ByVal strCitemCode As String, _
    ByVal strBillNo As String, ByVal lngItem As Long, ByVal lngAmount As Long, _
    ByVal strUpdTime As String, ByVal strCardBillNo As String, ByVal lngCustId As Long, _
    ByVal lngCompCode As Long, ByVal strServiceType As String, ByVal strAuthenticID As String, _
    ByVal strOrderNo As String) As Boolean
    On Error GoTo ChkErr
    Dim rs As New ADODB.Recordset
        If Not GetRS(rs, "Select * From " & GetOwner & "SO033 Where BillNo ='' And Item = -1", , adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
        Set rs.ActiveConnection = Nothing
        With rs
            If Not InsertChargeField(rs, lngCustId, lngCompCode, strServiceType, strBillNo, _
                lngItem, strCitemCode, RightDate) Then Exit Function
            rs("ShouldAmt") = lngAmount
            rs("OrderNo") = strOrderNo
            rs("CardBillNo") = NoZero(strCardBillNo)
            rs("AuthenticId") = NoZero(strAuthenticID)
            rs("ServiceType") = strServiceType
            rs("UpdTime") = GetDTString(strUpdTime)
            rs("UpdEn") = garyGi(1)
            If Not InsertToOracle("SO033", rs, "Where BillNo = '' And Item = -1") Then Exit Function
            rs.CancelUpdate
            .MoveNext
        End With
        CreditInsertTo033 = True
    Exit Function
ChkErr:
    ErrSub FormName, "InsertTo033"
End Function

Public Property Let uOwnerName(ByVal vErrPath As Variant)
    On Error Resume Next
        gOwnerName = vErrPath
        garyGi(16) = vErrPath
End Property

Public Property Let uErrPath(ByVal vErrPath As Variant)
    On Error Resume Next
        gErrLogPath = vErrPath
End Property

Public Property Set uConn(ByVal vConn As ADODB.Connection)
    On Error Resume Next
        Set gcnGi = vConn
End Property

Public Property Get uGetErrorCode() As String
    uGetErrorCode = strErrorCode
End Property

Public Property Get uGetErrorMsg() As String
    uGetErrorMsg = strErrorMsg

End Property
