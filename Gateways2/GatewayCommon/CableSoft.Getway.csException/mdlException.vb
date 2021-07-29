Imports System
Imports System.Windows
Imports System.Windows.Forms
Module mdlException
    Public Const fErrTxtFileName As String = "GatewayErr.Log"
    Public Const fSysEventLogFileName As String = "Gateway_SysEvent.Log"
    Public Const fCmdErrLogFileName As String = "Gateway_CmdErr.Log"
    Public Const fSOStatusLogFileName As String = "Gateway_SOStatus.Log"
    Public Const fTVMailErrLogFileName As String = "Gateway_TVMailError.Log"
    Public Const fEmailLogFileName As String = "Gateway_SendMail.Log"
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
