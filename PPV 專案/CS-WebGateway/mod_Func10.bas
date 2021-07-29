Attribute VB_Name = "mod_Func10"

Option Explicit

Private strAuthenticID As String
Private cn As Object

Public Function JustDoIt(ByRef objWebPool As Object, _
                                        ByRef strPara As String) As String
  On Error GoTo ChkErr
    Dim strQry As String
    Dim varPara As Variant
    Dim rs As Object
    Dim objIPPV As Object
    Dim lngCustID As Long
    Dim strOwner As String
    Dim objCnPool As Object
    
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
    
    '客戶變更STB四碼數字Parental PIN Code
    '10、會員代號、認證ID、設備流水號、STB序號、ICC序號
    '0:  成功
    '-9 : 密碼變更失敗
    
    varPara = Split(strPara, ",")
    '*****************************************************
    '#3018 增加判斷傳入的個數是否正確 By Kin 2007/08/30
    If UBound(varPara) <> 5 Then
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

'    Select count(*) as Cnt from so002b where MemberId=[會員代號] and AuthenticId=[認證ID]
'    If Cnt = 0 Then
'      回傳(-2,訂購PPV認證錯誤)
'    End If
    
'    strQry = "SELECT COUNT(*) AS RCDCNT FROM " & strOwner & "SO002B" & Get_DB_Link(cn) & _
'                " WHERE MEMBERID='" & varPara(1) & "' AND AUTHENTICID='" & strAuthenticID & "'"
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
    
'    依傳入之參數(STB序號、ICC序號)呼叫CA命令[變更Parental PIN Code]非視覺化模組. 高階代號E1
'    If 命令成功 Then
'       回傳(0,成功)
'    Else
'       回傳(-9,密碼變更失敗)
'    End If
    
    Set objIPPV = CreateObject("setIPPV.clsNagraSTB")
    With objIPPV
            Set .uConn = Get_General_CN
            garyGi(1) = "WEB"
            .ugaryGi = Join(garyGi, Chr(9))
    End With
    
    With objIPPV
        '    Public Function SetPinCode(ByVal strCompCode As String, _
        '        ByVal lngCustId As Long, ByVal strSmartCard As String, _
        '        ByVal strResvTime As String, ByVal strSTB As String, _
        '        ByVal strDoCaption As String, ByVal strPinCode As String) As Boolean
        If Not .SetPinCode(CStr(bytComp), _
                                        lngCustID, _
                                        CStr(varPara(5)), _
                                        "", _
                                        CStr(varPara(4)), _
                                        "", _
                                        "") Then
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
    Rlx strAuthenticID
    Set cn = Nothing
    Set objCnPool = Nothing
  Exit Function
ChkErr:
    ErrHandle "mod_Func10", "JustDoIt"
    JustDoIt = "-99," & strErr
End Function
