Attribute VB_Name = "mod_Func14"
Option Explicit

Private cn As Object

Public Function JustDoIt(ByRef objWebPool As Object, _
                                        ByRef strPara As String) As String
  On Error GoTo ChkErr
    Dim strQry As String
    Dim varPara As Variant
    Dim rs As Object
    Dim strOwner As String
    Dim objCnPool As Object
    
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
    
    
    ' 計次付費節目取消原因資料
    ' 14、會員代號、認證ID
    ' XML(訂單取消原因代碼、訂單取消原因)
    
    varPara = Split(strPara, ",")
    '*****************************************************
    '#3018 增加判斷傳入的個數是否正確 By Kin 2007/09/03
    If UBound(varPara) < 2 Then
        JustDoIt = "-99,參數不足!"
        GoTo 66
    End If
    '*****************************************************

    
    bytComp = Val(Left(varPara(2), 3))
    strOwner = GetOwner(cn)
    
    ' Old by Lawrence
    ' Select CodeNo,Description from CD015 where StopFlag=0 and Assort in (0,1) and
    '  (ServiceType='C'or ServiceType is null) Order By CodeNo
    
    ' New 2005/11/30 by Liga
    ' API-14為取得節目訂單的取消原因，故應是取 CD015.Assort in (0, 4)的才對
    ' Select CodeNo,Description from CD015 where StopFlag=0 and Assort in (0,4) and
    '  (ServiceType='C'or ServiceType is null) Order By CodeNo
    
    strQry = "SELECT CODENO,DESCRIPTION FROM " & strOwner & "CD015" & Get_DB_Link(cn) & _
                    "WHERE STOPFLAG=0" & _
                    " AND ASSORT IN (0,4)" & _
                    " AND SERVICETYPE='C' OR SERVICETYPE IS NULL" & _
                    " ORDER BY CODENO"

    Set rs = cn.Execute(strQry)
    
    ' If RecordSet筆數 > 0 Then
    '   產生XML檔.(依select出來之資料)
    '   回傳(0,成功)及XML檔
    ' Else
    '    回傳(-13,取消原因代碼讀取失敗)
    ' End If
    
    If rs.RecordCount <= 0 Then
        JustDoIt = "-13,取消原因代碼讀取失敗"
    Else
        JustDoIt = "0,成功" & vbCrLf & ExpXML(rs)
    End If
    
    ' 0 : 成功
    ' -13 : 取消原因代碼讀取失敗
    
    ' PS:
    ' XML(訂單取消原因代碼、訂單取消原因)
    
66:
    On Error Resume Next
    Rlx rs
    Rlx strQry
    Rlx varPara
    Rlx strOwner
    Set cn = Nothing
    Set objCnPool = Nothing
  Exit Function
ChkErr:
    ErrHandle "mod_Func14", "JustDoIt"
    JustDoIt = "-99," & strErr
End Function
