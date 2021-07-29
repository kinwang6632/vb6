Imports System
Imports System.Windows
Imports System.Windows.Forms
Module mdlException
    Public Const ErrTxtFileName As String = "GateWayErr.Log"
    Public Const SysEventLogName As String = "GateWay_SysEvent.Log"
    Public Const CmdErrLogName As String = "GateWay_CmdErr.Log"
    Public Const SOStatusLogName As String = "GateWay_SOStatus.Log"
    Public Const CableSoftErrMsg As String = "錯誤時間: {0} " & vbCrLf & _
                            "錯誤Function: {1} " & vbCrLf & _
                            "錯誤描述 : {2} " & vbCrLf & _
                            "錯誤行號 : {3} "

    Public Function GetAppPath() As String
        Dim strRet As String = Application.StartupPath
        If Not strRet.EndsWith("\") Then
            strRet = strRet & "\"
        End If
        Return strRet
    End Function
End Module
