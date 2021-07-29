Imports System.Windows.Forms

Public Class MDBMailPara
    '判斷是否使用警告Email設定
    Private Const fMDBFileName As String = "WarningEMail.SET"
    Private Shared objLock As New Object
    Public Shared Function GetDefaultMDBName() As String
        Try
            Return csCommon.GetAppPath & fMDBFileName
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function
    Public Overloads Shared Function IsUseSndMail(ByVal aMDBFile As String) As Boolean
        Try
            Dim aryLst As ArrayList = csCommon.GetFindFiles(aMDBFile)
            If aryLst Is Nothing Then
                Return False
            End If
            If aryLst.Count <> 1 Then
                Throw New Exception("檔案數量不為1")
                Return False
            End If
            Return True
        Catch ex As Exception
            Throw New Exception(ex.Message.ToString)
            Return False
        End Try
    End Function
    Public Overloads Shared Function IsUseSndMail() As Boolean
        Try
            Dim aFile As String = csCommon.GetAppPath & fMDBFileName
            Return IsUseSndMail(aFile)
        Catch ex As Exception
            Throw New Exception(ex.Message.ToString)
            Return False
        End Try
    End Function
    ''' <summary>
    ''' 讀取MDB EMAIL設定檔
    ''' </summary>
    ''' <param name="aMDBName">MDB 檔名</param>
    ''' <param name="aMDBPassword">MDB 密碼</param>
    ''' <param name="aGuid">給予設定唯一的值</param>
    ''' <param name="aMailInfoTable">EMail一般設定的 Table </param>
    ''' <param name="aMailToTable">Email 收件者的 Table</param>
    ''' <param name="aMailCCTable">Email 副本的 Table </param>
    ''' <returns>讀取成功回傳True,否則回傳False</returns>
    ''' <remarks></remarks>
    Public Overloads Shared Function ReadEMailPara(ByVal aMDBName As String, _
                                                ByVal aMDBPassword As String, _
                                                ByVal aGuid As String, _
                                                ByVal aMailInfoTable As String, _
                                                ByVal aMailToTable As String, _
                                                ByVal aMailCCTable As String) As Boolean
        Try
            SyncLock objLock


                '如果已經設定過不要再讀取了
                If aryMailLst IsNot Nothing Then
                    If aryMailLst.ContainsKey(aGuid) Then
                        Return True
                    End If
                End If

                Dim tblEmailInfo As DataTable = MDBCommon.ReadMDBDataTable(aMDBName, aMDBPassword, aMailInfoTable)
                Dim tblEmailTo As DataTable = MDBCommon.ReadMDBDataTable(aMDBName, aMDBPassword, aMailToTable)
                Dim tblEmailCC As DataTable = MDBCommon.ReadMDBDataTable(aMDBName, aMDBPassword, aMailCCTable)
                Dim lstMailTo As List(Of String) = Nothing
                Dim lstMailCC As List(Of String) = Nothing
                If (tblEmailInfo Is Nothing) OrElse (tblEmailInfo.Rows.Count <= 0) Then
                    Throw New Exception("郵件資訊未設定")
                    Return False
                End If
                If (tblEmailTo Is Nothing) OrElse (tblEmailTo.Rows.Count <= 0) Then
                    Throw New Exception("收件人資訊未設定")
                    Return False
                End If
                If lstMailTo Is Nothing Then
                    lstMailTo = New List(Of String)
                End If
                For i As Int32 = 0 To tblEmailTo.Rows.Count - 1
                    lstMailTo.Add(tblEmailTo.Rows(i).Item("MailTo"))
                Next
                If lstMailCC Is Nothing Then
                    lstMailCC = New List(Of String)
                End If
                For i As Int32 = 0 To tblEmailCC.Rows.Count - 1
                    lstMailCC.Add(tblEmailCC.Rows(i).Item("MailCC"))
                Next

                Dim StruEmail As New MailInfo()
                StruEmail.EnabledSSL = False
                StruEmail.SMTP_PORT = 25
                StruEmail.MailTo = lstMailTo
                StruEmail.MailCC = lstMailCC
                With tblEmailInfo.Rows(0)
                    StruEmail.SMTP_Host = .Item("SmtpHost")
                    StruEmail.EnabledSSL = Convert.ToInt32(.Item("EnabledSSL") & "") = 1
                    StruEmail.Password = .Item("SmtpPassword")
                    StruEmail.UserId = .Item("SmtpUserId")
                    StruEmail.Sender = .Item("MailSender")
                    StruEmail.SenderDisplayName = .Item("MailDisplayNam")
                    StruEmail.SMTP_PORT = Convert.ToInt32(.Item("SmtpPort") & "")
                    StruEmail.MailSubject = .Item("MailSubJect")
                    StruEmail.MailBody = .Item("MailBody")
                End With
                If String.IsNullOrEmpty(StruEmail.Sender) OrElse _
                    String.IsNullOrEmpty(StruEmail.SMTP_Host) OrElse _
                    String.IsNullOrEmpty(StruEmail.UserId) OrElse _
                    String.IsNullOrEmpty(StruEmail.Password) OrElse _
                    String.IsNullOrEmpty(StruEmail.MailSubject) Then
                    Throw New Exception("郵件資訊不完整!")
                    Return False
                End If
                If aryMailLst Is Nothing Then
                    aryMailLst = New Dictionary(Of String, MailInfo)
                End If
                aryMailLst.Add(aGuid, StruEmail)
                Return True
            End SyncLock
        Catch ex As Exception
            Throw New Exception(ex.Message)
            Return False
        End Try
    End Function
    Public Shared Function DefaultPSPara(ByVal aMDBFile As String, ByVal aMDBPassword As String, _
                                          ByVal aMDBTable As String, _
                                            Optional ByVal aDescrField As Array = Nothing) As Boolean
        Try
            Dim aDataTable As DataTable = MDBCommon.ReadMDBDataTable(aMDBFile, aMDBPassword, aMDBTable, Nothing)
            Dim rw As DataRow = aDataTable.NewRow
            rw("ReadWebErrCnt") = "100"
            aDataTable.Rows.Add(rw)
            CableSoft.Gateway.Common.MDBCommon.SaveMDBDataTable(aMDBFile, aMDBPassword, aMDBTable, aDataTable)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function


    Public Shared Function DefaultPRMPara(ByVal aMDBFile As String, ByVal aMDBPassword As String, _
                                          ByVal aMDBTable As String, _
                                            Optional ByVal aDescrField As Array = Nothing) As Boolean
        Try
            Dim aDataTable As DataTable = MDBCommon.ReadMDBDataTable(aMDBFile, aMDBPassword, aMDBTable, Nothing)
            Dim rw As DataRow = aDataTable.NewRow
            rw("SourceIdErrCnt") = "3"
            aDataTable.Rows.Add(rw)
            Cablesoft.Gateway.Common.MDBCommon.SaveMDBDataTable(aMDBFile, aMDBPassword, aMDBTable, aDataTable)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Shared Function DefaultMailInfo(ByVal aMDBFile As String, ByVal aMDBPassword As String, _
                                            ByVal aMDBTable As String, _
                                            Optional ByVal aDescrField As Array = Nothing) As Boolean
        Try
            Dim aDataTable As DataTable = MDBCommon.ReadMDBDataTable(aMDBFile, aMDBPassword, aMDBTable, Nothing)
            Dim rw As DataRow = aDataTable.NewRow
            rw("SmtpHost") = "mail.cablesoft.com.tw"
            rw("SmtplUserId") = "kinwang"
            rw("SmtpPassword") = "chunya36"
            rw("EnabledSSL") = "0"
            rw("MailDisplayNam") = "自由"
            rw("MailSender") = "kinwang@cablesoft.com.tw"
            rw("MailSubject") = "測試主旨"
            rw("MailBody") = "測試內容"
            rw("SmtpPort") = "25"
            aDataTable.Rows.Add(rw)
            CableSoft.Gateway.Common.MDBCommon.SaveMDBDataTable(aMDBFile, aMDBPassword, aMDBTable, aDataTable)
            Return True
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function
    Public Shared Function DefaultMailTo(ByVal aMDBFile As String, ByVal aMDBPassword As String, _
                                            ByVal aMDBTable As String, _
                                            Optional ByVal aDescrField As Array = Nothing) As Boolean
        Try
            Dim aDataTable As DataTable = MDBCommon.ReadMDBDataTable(aMDBFile, aMDBPassword, aMDBTable, Nothing)
            Dim rw As DataRow = aDataTable.NewRow
            rw("MailTo") = "kinwang@cablesoft.com.tw"
            aDataTable.Rows.Add(rw)
            CableSoft.Gateway.Common.MDBCommon.SaveMDBDataTable(aMDBFile, aMDBPassword, aMDBTable, aDataTable)
            Return True
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function
End Class
