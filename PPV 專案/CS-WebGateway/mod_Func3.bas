Attribute VB_Name = "mod_Func3"
Option Explicit

' �W���� ...
Private cn As Object

Public Function JustDoIt(ByRef objWebPool As Object, _
                                        ByRef strPara As String) As String
  On Error GoTo ChkErr
    Dim strQry As String
    Dim varPara As Variant
    Dim rs As Object
    Dim lngCustID As Long
    Dim objCnPool As Object
    Dim strAuthenticID As String
    
    Set objCnPool = objWebPool
    Set cn = objCnPool.GetConnection
    
    If cn.state <= 0 Then
        Set cn = objCnPool.GetConnection
        If cn.state <= 0 Then
            strErr = "�L�k�s�u��Ʈw!"
            JustDoIt = "-99,��Ʈw�s�u����!"
            GoTo 66
        End If
    End If
    
    JustDoIt = "�W����.."
    
    varPara = Split(strPara, ",")
    
    strAuthenticID = varPara(2)
    bytComp = Val(Left(strAuthenticID, 3))
    
66:
    On Error Resume Next
    Rlx rs
    Rlx strQry
    Rlx varPara
    Rlx lngCustID
    Rlx strAuthenticID
    Set cn = Nothing
    Set objCnPool = Nothing
  Exit Function
ChkErr:
    ErrHandle "mod_Func3", "JustDoIt"
    JustDoIt = "-99," & strErr
End Function

