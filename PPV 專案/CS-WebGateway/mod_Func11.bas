Attribute VB_Name = "mod_Func11"
Option Explicit

Private strAuthenticID As String
Private cn As Object

Public Function JustDoIt(ByRef objWebPool As Object, _
                                       ByRef strPara As String) As String
  On Error GoTo ChkErr
    Dim strQry As String
    Dim varPara As Variant
    Dim rs As Object
    Dim lngCustID As Long
    Dim objCnPool As Object
    Dim strOwner As String
    Dim objIPPV As Object
    Dim strPinCode As String
    
'    Set objCnPool = objWebPool
'    Set cn = objCnPool.GetConnection
'
'    If cn.State <= 0 Then
'        Set cn = objCnPool.GetConnection
'        If cn.State <= 0 Then
'            strErr = "無法連線資料庫!"
'            JustDoIt = "-99,資料庫連線失敗!"
'            GoTo 66
'        End If
'    End If
    
    varPara = Split(strPara, ",")
    '*****************************************************
    '#3018 增加判斷傳入的個數是否正確 By Kin 2007/08/30
    If UBound(varPara) < 6 Then
        JustDoIt = "-99,參數不足!"
        GoTo 66
    End If
    '*****************************************************
    
    strAuthenticID = varPara(2)
    bytComp = Val(Left(strAuthenticID, 3))
    lngCustID = Val(Right(strAuthenticID, 7))
    Set cn = Get_General_CN2(bytComp)
    If cn.State <= 0 Then
        Set cn = Get_General_CN2(bytComp)
        If cn.State <= 0 Then
            strErr = "無法連線資料庫!"
            JustDoIt = "-99,資料庫連線失敗!"
            GoTo 66
        End If
    End If

    strOwner = GetOwner(cn)
    strPinCode = varPara(6)
'    If strPinCode = Empty Then strPinCode = "1234"
    If strPinCode = Empty Then strPinCode = ""
    
    '   Select count(*) as Cnt from so002b where MemberId=[會員代號] and AuthenticId=[認證ID]
    '   If Cnt = 0 Then
    '       回傳(-2,訂購PPV認證錯誤)
    '   End If
    
'    strQry = "SELECT COUNT(*) AS RCDCNT FROM " & strOwner & "SO002B" & Get_DB_Link(cn) & _
'                    "WHERE MEMBERID='" & varPara(1) & "' AND AUTHENTICID='" & strAuthenticID & "'"
'
'
'    If Val(GetValue(cn, strQry)) = 0 Then
'        JustDoIt = "-2,訂購PPV認證錯誤"
'        GoTo 66
'    End If
    
    strQry = "SELECT CUSTID FROM " & strOwner & "SO002B" & Get_DB_Link(cn) & _
                    "WHERE MEMBERID='" & varPara(1) & "' AND AUTHENTICID='" & strAuthenticID & "'"
    
    lngCustID = Val(GetValue(cn, strQry) & "")
    
    If lngCustID = 0 Then
        JustDoIt = "-2,訂購PPV認證錯誤"
        GoTo 66
    End If
    
    Set objIPPV = CreateObject("setIPPV.clsNagraSTB")
    With objIPPV
            Set .uConn = Get_General_CN
            garyGi(1) = "WEB"
            .ugaryGi = Join(garyGi, Chr(9))
    End With
    
    '   11、會員代號、認證ID、設備流水號、STB序號、ICC序號、iPPV PIN Code(4碼)
    '   0:  成功
    '   -9 : 密碼變更失敗
    
    '   依傳入之參數(STB序號、ICC序號、iPPV PIN Code)呼叫CA命令[變更iPPV訂購PIN Code]非視覺化模組. 高階代號E8
    '   If 命令成功 Then
    '       回傳(0,成功)
    '   Else
    '       回傳(-9,密碼變更失敗)
    '   End If
    
    '   Old
    '   PS : iPPV PIN Code四碼. 如果iPPV PIN Code傳入null時(即無值),
    '   則呼叫CA命令[變更iPPV訂購PIN Code]非視覺化模組時, PIN Code部分請傳入'1234'
    
    '   New
    '   PS : iPPV PIN Code四碼. 如果iPPV PIN Code傳入null時(即無值),
    '   則呼叫CA命令[變更iPPV訂購PIN Code]非視覺化模組時, PIN Code部分請傳入'0000'
    
    With objIPPV
    '    Public Function SetResetiPPVwd(ByVal strCompCode As String, _
    '        ByVal lngCustId As Long, ByVal strSmartCard As String, _
    '        ByVal strResvTime As String, ByVal strSTB As String, _
    '        ByVal strDoCaption As String, ByVal strPinCode As String) As Boolean
    '       Format(Now, "YYYY/MM/DD HH:MM")
        If Not .SetResetiPPVwd(CStr(bytComp), _
                                                lngCustID, _
                                                CStr(varPara(5)), _
                                                "", _
                                                CStr(varPara(4)), _
                                                "", _
                                                strPinCode) Then
            strErr = .uMsgBack
            JustDoIt = "-9,密碼變更失敗"
        Else
            JustDoIt = "0,成功"
        End If
    End With
        
66:
    On Error Resume Next
    Rlx rs
    Rlx strQry
    Rlx varPara
    Rlx objIPPV
    Rlx strOwner
    Rlx lngCustID
    Rlx strPinCode
    Rlx strAuthenticID
    Set cn = Nothing
    Set objCnPool = Nothing
  Exit Function
ChkErr:
    ErrHandle "mod_Func11", "JustDoIt"
    JustDoIt = "-99," & strErr
End Function


