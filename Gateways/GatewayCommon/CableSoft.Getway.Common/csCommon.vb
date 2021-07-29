Imports System
Imports System.Windows
Imports System.Windows.Forms
Imports System.Diagnostics
Imports System.IO
Imports System.Net
Imports System.Net.Mail
Imports System.Threading
Imports System.Net.NetworkInformation

Public Class csCommon
    Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Int32
    Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Int32, ByVal lpFileName As String) As Int32
    Private Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Int32
    Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Int32
    Private Shared objLock As New Object
    Private Shared fIsSyncThread As Boolean = False
    Private Shared fWriteLog As Boolean = True
    ''' <summary>
    ''' 取得程式目錄
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetAppPath() As String
        Dim strRet As String = Application.StartupPath
        If Not strRet.EndsWith("\") Then
            strRet = strRet & "\"
        End If
        Return strRet
    End Function
    ''' <summary>
    ''' 取出指定路徑的所有檔案
    ''' </summary>
    ''' <param name="AppPath">路徑</param>
    ''' <param name="aFileName">尋找的檔案</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overloads Shared Function GetFindFiles(ByVal AppPath As String, ByVal aFileName As String) As ArrayList
        Dim strFiles() As String
        Dim aPath As String = String.Empty
        aPath = AppPath
        If Not AppPath.EndsWith("\") Then
            aPath = aPath & "\"
        End If

        Try
            strFiles = Directory.GetFiles(aPath, aFileName).ToArray
            If strFiles.Length = 0 Then
                Return Nothing
            Else
                Dim retAry As New ArrayList
                retAry.AddRange(strFiles)
                Return retAry
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
            Return Nothing
        End Try

    End Function
    Public Overloads Shared Function GetFindFiles(ByVal aFileName As String) As ArrayList
        Return GetFindFiles(GetAppPath(), aFileName)
    End Function
    Public Overloads Shared Sub WriteLog(ByVal aMsg As String, ByVal aFileName As String)
        SyncLock objLock
            Dim aFile As String = String.Empty
            Dim aDirectory As String = String.Empty
            aDirectory = GetAppPath() & "Log\" & Format(Now, "yyyyMMdd")
            If Not String.IsNullOrEmpty(aFileName) Then
                aFile = aDirectory & "\" & aFileName
            Else
                aFile = aDirectory & "\Gateway_Log.Txt"
            End If

            Try
                If Not Directory.Exists(aDirectory) Then
                    Directory.CreateDirectory(aDirectory)
                End If
                Using f As New StreamWriter(aFile, True)
                    f.WriteLine(String.Format("{0} > {1}", Format(Now, "yyyy/MM/dd HH:mm:ss"), aMsg))
                    f.Flush()
                    f.Close()
                End Using
            Finally
                csException.WriteErrTxtLog.DeleteLog(GetAppPath() & "Log", 1)
            End Try
        End SyncLock
    End Sub
    Public Overloads Shared Sub WriteLog(ByVal aMsg As String)
        WriteLog(aMsg, String.Empty)

    End Sub
    Public Shared Function ReadINI(ByVal aFilename As String, _
                                       ByVal aSection As String, _
                                       ByVal aKey As String, _
                                       Optional ByVal DefaultValue As String = "") As String
        Try
            Dim aRet As String = New String(Chr(0), 255)
            GetPrivateProfileString(aSection, aKey, String.Empty, aRet, 255, aFilename)
            aRet = aRet.TrimEnd(Chr(0))
            If String.IsNullOrEmpty(aRet) Then
                aRet = DefaultValue
            End If
            Return aRet
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function
    Public Shared Sub WriteINI(ByVal aFilename As String, ByVal aSection As String, ByVal aKey As String, ByVal aValue As String)
        Try
            WritePrivateProfileString(aSection, aKey, aValue, aFilename)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub
    ''' <summary>
    ''' 寄發Mail
    ''' </summary>
    ''' <param name="aMDBName">Mail MDB設定檔</param>
    ''' <param name="aMDBPassword">MDB 密碼</param>
    ''' <param name="aGuid">設定唯一值，讀取MDB過程式如果已有讀取過，不在重複讀取</param>
    ''' <param name="aMailInfoTable">讀取Mail基本設定TABLE</param>
    ''' <param name="aMailToTable">讀取收件人設定TABLE</param>
    ''' <param name="aMailCCTable">讀取寄件人設定TABLE</param>
    ''' <param name="aMailSubject">主旨設定(空值代表讀取MDB設定)</param>
    ''' <param name="aMailBody">內容設定(空值代表讀取MDB設定)</param>
    ''' <param name="IsSyncThread">是否同步執行(設定為False,會另開Thread執行)</param>
    ''' <param name="aWriteLog">是否寫入LOG檔(IsSyncThread=False才會生效)</param>
    ''' <remarks></remarks>

    Public Overloads Shared Function SendMail(ByVal aMDBName As String, _
                                                ByVal aMDBPassword As String, _
                                                ByVal aGuid As String, _
                                                ByVal aMailInfoTable As String, _
                                                ByVal aMailToTable As String, _
                                                ByVal aMailCCTable As String,
                                                ByVal aMailSubject As String, _
                                                ByVal aMailBody As String, _
                                                ByVal IsSyncThread As Boolean, _
                                                ByVal aWriteLog As Boolean) As Boolean

        Try


            If MDBMailPara.ReadEMailPara(aMDBName, aMDBPassword, _
                                            aGuid, aMailInfoTable, aMailToTable, aMailCCTable) Then
                If String.IsNullOrEmpty(aMailSubject) Then
                    aMailSubject = aryMailLst.Item(aGuid).MailSubject
                End If
                If String.IsNullOrEmpty(aMailBody) Then
                    aMailBody = aryMailLst.Item(aGuid).MailBody
                End If


                Return SendMail(aryMailLst.Item(aGuid).SMTP_Host, _
                                aryMailLst.Item(aGuid).SMTP_PORT, _
                                aryMailLst.Item(aGuid).UserId, _
                                aryMailLst.Item(aGuid).Password, _
                                aryMailLst.Item(aGuid).EnabledSSL, _
                                aryMailLst.Item(aGuid).MailTo, _
                                aryMailLst.Item(aGuid).MailCC, _
                                aryMailLst.Item(aGuid).Sender, _
                                aryMailLst.Item(aGuid).SenderDisplayName, _
                                aMailSubject, aMailBody, IsSyncThread, aWriteLog)
            Else
                Return False
            End If


        Catch ex As Exception
            If IsSyncThread Then
                Throw New Exception(ex.Message)
            Else
                If aWriteLog Then
                    csException.WriteErrTxtLog.WriteSndEmailLog(ex.Message)
                End If

            End If

            Return False
        End Try
    End Function
    Private Shared Function BeginSendMail(ByVal aMailInfo As MailInfo) As Boolean
        Try
            Using objSmtp As New SmtpClient(aMailInfo.SMTP_Host, aMailInfo.SMTP_PORT)
                objSmtp.Credentials = New NetworkCredential(aMailInfo.UserId, aMailInfo.Password)
                Using objMailMsg As New MailMessage()
                    objMailMsg.From = New MailAddress(aMailInfo.Sender, aMailInfo.SenderDisplayName)
                    For i As Int32 = 0 To aMailInfo.MailTo.Count - 1
                        objMailMsg.To.Add(aMailInfo.MailTo.Item(i))
                    Next
                    For i As Int32 = 0 To aMailInfo.MailCC.Count - 1
                        objMailMsg.CC.Add(aMailInfo.MailCC.Item(i))
                    Next
                    objMailMsg.Subject = aMailInfo.MailSubject
                    objMailMsg.Body = aMailInfo.MailBody
                    objSmtp.Send(objMailMsg)
                End Using
            End Using
            If fWriteLog Then
                csException.WriteErrTxtLog.WriteSndEmailLog("寄送Mail成功")
            End If
            Return True
        Catch ex As SmtpFailedRecipientsException

            Dim exString As String = "傳送郵件失敗" & Environment.NewLine
            For i As Int32 = 0 To ex.InnerExceptions.Length - 1
                Dim status As SmtpStatusCode = ex.InnerExceptions(i).StatusCode
                exString = exString & "收件者：" & ex.InnerExceptions(i).FailedRecipient & "--->" & status & _
                            Environment.NewLine
            Next
            If fIsSyncThread Then
                Throw New Exception(exString)
            Else
                If fWriteLog Then
                    csException.WriteErrTxtLog.WriteSndEmailLog(exString)
                End If

            End If

        Catch ex As Exception
            If fIsSyncThread Then
                Throw New Exception(ex.Message)
            Else
                If fWriteLog Then
                    csException.WriteErrTxtLog.WriteSndEmailLog(ex.Message)
                End If

            End If
            Return False
        End Try
    End Function

    Public Overloads Shared Function SendMail(ByVal aSMTPHost As String,
                                         ByVal aSMTPPort As String, _
                                         ByVal aLoginId As String, _
                                         ByVal aLoginPwd As String, _
                                        ByVal aEnabledSSL As Boolean, _
                                        ByVal aMailTo As List(Of String), _
                                        ByVal aMailCC As List(Of String), _
                                        ByVal aSender As String, _
                                        ByVal aMailDisplayName As String, _
                                        ByVal aMailSubject As String, _
                                        ByVal aMailBody As String, _
                                         ByVal IsSyncThread As Boolean, ByVal aWriteLog As Boolean) As Boolean
        Try
            If String.IsNullOrEmpty(aSMTPHost) OrElse _
                String.IsNullOrEmpty(aSMTPPort) OrElse _
                String.IsNullOrEmpty(aLoginId) OrElse _
                String.IsNullOrEmpty(aLoginPwd) Then
                Exit Function
            End If
            Dim objMailInfo As New MailInfo()
            With objMailInfo
                .EnabledSSL = aEnabledSSL
                .MailBody = aMailBody
                .MailCC = aMailCC
                .MailSubject = aMailSubject
                .MailTo = aMailTo
                .Password = aLoginPwd
                .Sender = aSender
                .SenderDisplayName = aMailDisplayName
                .SMTP_Host = aSMTPHost
                .SMTP_PORT = aSMTPPort
                .UserId = aLoginId
            End With
            fIsSyncThread = IsSyncThread
            fWriteLog = aWriteLog
            If IsSyncThread Then
                Return BeginSendMail(objMailInfo)
            Else
                Dim thr As New Thread(AddressOf BeginSendMail)
                thr.IsBackground = True
                thr.Priority = ThreadPriority.Highest
                thr.Start(objMailInfo)

            End If
            Return True
        Catch ex As Exception
            If IsSyncThread Then
                Throw New Exception(ex.Message)
            Else
                If aWriteLog Then
                    csException.WriteErrTxtLog.WriteSndEmailLog(ex.Message)
                End If
            End If
            Return False
        End Try
    End Function

 



End Class
