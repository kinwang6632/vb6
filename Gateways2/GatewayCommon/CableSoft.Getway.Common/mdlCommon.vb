Module mdlCommon
    Public Structure MailInfo
        Public SMTP_Host As String
        Public SMTP_PORT As String
        Public MailTo As List(Of String)
        Public MailCC As List(Of String)
        Public UserId As String
        Public Password As String
        Public EnabledSSL As Boolean
        Public Sender As String
        Public SenderDisplayName As String
        Public MailSubject As String
        Public MailBody As String
    End Structure
    Public aryMailLst As Dictionary(Of String, MailInfo) = Nothing
End Module
