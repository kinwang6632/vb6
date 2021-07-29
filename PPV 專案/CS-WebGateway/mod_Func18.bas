Attribute VB_Name = "mod_Func18"
Option Explicit

Private cn As Object
Private strAuthenticID As String

Public Function JustDoIt(ByRef objWebPool As Object, _
                                        ByRef strPara As String) As String
  On Error GoTo ChkErr
    Dim strQry As String
    Dim rs As Object
    Dim objCnPool As Object
    Dim strOwner As String
    Dim varPara As Variant

    Set objCnPool = objWebPool
    Set cn = objCnPool.GetConnection
    
    If cn.State <= 0 Then
        Set cn = objCnPool.GetConnection
        If cn.State <= 0 Then
            strErr = "無法連線資料庫!"
            JustDoIt = "-99,資料庫連線失敗!"
            GoTo 66
        End If
    End If
    
    strOwner = GetOwner(cn)
    
    varPara = Split(strPara, ",")
    '*****************************************************
    '#3018 增加判斷傳入的個數是否正確 By Kin 2007/09/03
    If UBound(varPara) < 2 Then
        JustDoIt = "-99,參數不足!"
        GoTo 66
    End If
    '*****************************************************

    strAuthenticID = varPara(2)
    bytComp = Val(Left(strAuthenticID, 3))
    
    '18、會員代號、認證ID
    
    '讀取點數收費資料
    '18
    'XML(收費項目代碼、收費項目名稱)
    
    'Select CodeNo, Description, Amount, CreditPoint from cd019 where StopFlag=0 and
    'RefNo in (4,17) Order By CodeNo
    
    strQry = "SELECT CODENO,DESCRIPTION,AMOUNT,CREDITPOINT FROM " & strOwner & "CD019" & _
                    Get_DB_Link(cn) & " WHERE STOPFLAG <> 1 AND REFNO IN (4,17) ORDER BY CODENO"

    Set rs = cn.Execute(strQry)
    
    'If RecordSet筆數 > 0 Then
    '  產生XML檔.(依select出來之資料)
    '  回傳(0,成功)及XML檔
    'Else
    '   回傳(-15,查無點數收費資料)
    'End If
    
    'PS:
    'XML(收費項目代碼、收費項目名稱、金額、點數)
    
    If rs.RecordCount <= 0 Then
        JustDoIt = "-15,查無點數收費資料"
    Else
        JustDoIt = "0,成功" & vbCrLf & ExpXML(rs)
    End If
    
66:
    On Error Resume Next
    Rlx rs
    Rlx strQry
    Rlx varPara
    Rlx strOwner
    Rlx strAuthenticID
    Set cn = Nothing
    Set objCnPool = Nothing
  Exit Function
ChkErr:
    ErrHandle "mod_Func18", "JustDoIt"
    JustDoIt = "-99," & strErr
End Function
