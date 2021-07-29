Imports System
Imports System.IO
Imports System.Text
Imports System.Data
Imports System.Data.OracleClient
Imports System.Threading
Imports System.Threading.Timeout
Imports System.Threading.Tasks
Imports System.Net.Mime.MediaTypeNames
Imports CableSoft.DynCallWebService

Public Class SO
    Implements IDisposable    
    Public Property CompCode As Integer
    Public Property CompName As String
    Public Property prgName As String
    Public Property Sid As String
    Public Property DBUser As String
    Public Property DBPassword As String
    Public Property DBMinPool As Integer
    Public Property DBMaxPool As Integer
    Public Property DBLifetime As Integer
    Public Property MonitorType As String
    Public Property Runtimer As Integer
    Public Property Timeout As Integer
    Public Property SeqSql As String
    Public Property InsSql As String
    Public Property GetSql As String
    Public Property WebServiceName As String
    Public Property WebSrvMethod As String
    Public Property WebServicePara As String
    Public Property url As String
    Public Property WebGetPara As String
    Public Property WebPostPara As String
    Public Property logReserve As Integer = 6
    Public Property sequence As Integer
    Public Event StatusChange(ByVal StatusOK As StatusOfProcess,
                              ByVal sequence As Integer,
                              ByVal comId As Integer,
                              ByVal compName As String,
                              ByVal prgName As String,
                              ByVal msg As String)
    Private isRunning As Boolean = False
    Private isStopAll As Boolean = False
    Private WaitProcessCount As Integer
    Private conBuilder As OracleConnectionStringBuilder    
    Private connectString As String = Nothing
    Private lstMonitortype As List(Of OverseeType)
    Private cnDB As OracleConnection = Nothing
    'Private cnDB As IDbConnection = Nothing
    Private Const sequenceKey As String = "Sequence"
    Private Const getSQLKey As String = "getSQL"
    Private Const insSQLKey As String = "insSQL"
    Private Const LogFileName As String = "Log.txt"
    Private Const errLogFileName As String = "Errlog.txt"
    Private logPathName As String = Nothing
    Private historicalName As String = Nothing
    Private errFileName As String = Nothing
    Private startRunTime As Date = Date.Today
    Private firstRun As Boolean = True
    Private CallWebService As WebServiceObject = Nothing
    Private lckObject As New Object
    Private tmr As Threading.Timer = Nothing

    Public Enum OverseeType
        WebService = 0
        WebPost = 1
        WebGet = 2
        DBAccess = 3
    End Enum
    Public Enum StatusOfProcess
        Processing = 0
        Success = 1
        Failure = 2
    End Enum
    Public Sub New()        
    End Sub
    Public Sub Execute()
        If lstMonitortype Is Nothing Then
            lstMonitortype = New List(Of OverseeType)
            For Each Str As String In MonitorType.Split(",")
                lstMonitortype.Add(Integer.Parse(Str))
            Next
        End If
        logPathName = String.Format("{0}\{1}\{2}",
                                     Path.GetDirectoryName(System.AppDomain.CurrentDomain.BaseDirectory),
                                     Me.CompCode, Me.prgName)

        Try
            If lstMonitortype.Contains(OverseeType.DBAccess) Then
                If conBuilder Is Nothing Then
                    conBuilder = New OracleConnectionStringBuilder
                End If
                conBuilder.UserID = Me.DBUser
                conBuilder.Password = Me.DBPassword
                conBuilder.DataSource = Me.Sid
                conBuilder.Pooling = True
                conBuilder.MaxPoolSize = Me.DBMaxPool
                conBuilder.MinPoolSize = Me.DBMinPool
                conBuilder.LoadBalanceTimeout = Me.DBLifetime
                conBuilder.PersistSecurityInfo = True
                connectString = conBuilder.ToString
                If cnDB Is Nothing Then
                    cnDB = New OracleConnection(connectString)
                    cnDB.Open()
                End If
            End If

            'mtxWait = New Mutex()

            'evn = New AutoResetEvent(False)
            'ti = ThreadPool.RegisterWaitForSingleObject(evn, _
            '                                                   New WaitOrTimerCallback(AddressOf WaitProc), _
            '                                                   Nothing, TimeSpan.FromSeconds(Me.Runtimer), False)
            tmr = New Threading.Timer(AddressOf WaitProc, Nothing, 0, Me.Runtimer * 1000)

            'evn.Set()

        Catch ex As Exception
            WriteMsg(ex, Nothing)
        End Try

    End Sub
    Private Sub chkDBStatus(ByVal stateInfo As Object)
        Try
            Dim sqlList As Dictionary(Of String, String) = CType(stateInfo, Dictionary(Of String, String))

        Catch ex As Exception

        End Try
    End Sub
    Private Function analyzeInsSql(ByVal sequence As String) As String
        Return InsSql.ToUpper.Replace("<SeqNo>".ToUpper, sequence)
    End Function
    Private Function analyzeGetSql(ByVal sequence As String) As String
        Return GetSql.ToUpper.Replace("<SeqNo>".ToUpper, sequence)
    End Function
    Private Function getSequence() As String
        Dim result As String = Nothing
        Try

            If cnDB Is Nothing Then
                cnDB = New OracleConnection(connectString)
            End If
            If cnDB.State <> ConnectionState.Open Then
                cnDB.Open()
            End If
            Using cnCommand As OracleCommand = New OracleCommand(Me.SeqSql, cnDB)
                result = cnCommand.ExecuteScalar
                cnCommand.Dispose()
                cnDB.Close()
            End Using
        Catch ex As Exception
            WriteMsg(ex, Nothing)
        Finally
            If cnDB IsNot Nothing Then
                If cnDB.State = ConnectionState.Open Then
                    cnDB.Close()
                End If
            End If
        End Try
        Return result
    End Function
    Private Function DeleteLog() As Boolean

        Try

            Parallel.ForEach(New DirectoryInfo(logPathName).GetDirectories,
                                                                        Sub(currenPath)
                                                                            If (IsNumeric(currenPath.Name)) AndAlso (currenPath.Name.Length = 8) Then
                                                                                If Integer.Parse(Format(Date.Now, "yyyyMM")) -
                                                                                                Integer.Parse(Left(currenPath.Name, 6)) >= Me.logReserve Then
                                                                                    Directory.Delete(currenPath.FullName, True)
                                                                                End If
                                                                            End If
                                                                        End Sub)
            'For Each s As DirectoryInfo In New DirectoryInfo(logPathName).GetDirectories
            '    If (IsNumeric(s.Name)) AndAlso (s.Name.Length = 8) Then
            '        If Int32.Parse(Format(Date.Now, "yyyyMM")) - Int32.Parse(Left(s.Name, 6)) >= Me.logReserve Then
            '            Directory.Delete(s.FullName, True)
            '        End If
            '    End If
            'Next
        Catch ex As Exception

        End Try
        Return True
    End Function
    Private Overloads Shared Sub WriteLog(ByVal FileName As String, ByVal msg As String)
        Dim dirRoot As New ThreadLocal(Of String)
        dirRoot.Value = Path.GetDirectoryName(FileName)
        Try

            If Not Directory.Exists(dirRoot.Value) Then
                Directory.CreateDirectory(dirRoot.Value)
            End If
            Using WriteFile As New StreamWriter(FileName, True)
                SyncLock WriteFile
                    WriteFile.WriteLine(msg)
                    WriteFile.Flush()
                    WriteFile.Close()
                End SyncLock
            End Using
        Catch ex As Exception
            Debug.Print(ex.ToString)
        Finally
            If dirRoot IsNot Nothing Then
                dirRoot.Value = Nothing
                dirRoot.Dispose()
                dirRoot = Nothing
            End If
        End Try
    End Sub
    Private Sub WriteMsg(ByVal ex As Exception, ByVal msg As String)
        Dim Nowtime As New ThreadLocal(Of Date)
        Try

            Nowtime.Value = Date.Now
            If ex IsNot Nothing Then

                WriteLog(String.Format("{0}\{1}\{2}", logPathName,
                                                             String.Format(Nowtime.Value.ToString("yyyyMMdd")),
                                                             errLogFileName),
                                                         String.Format("{0} => 發生錯誤：{1}", Nowtime.Value.ToString("HH:mm:ss"),
                                                                       ex.ToString))
                WriteLog(String.Format("{0}\{1}\{2}", logPathName,
                                                            String.Format(Nowtime.Value.ToString("yyyyMMdd")),
                                                            LogFileName),
                                                        String.Format("{0} => 發生錯誤：{1}", Nowtime.Value.ToString("HH:mm:ss"),
                                                                      ex.ToString))
                raiseStatusEven(StatusOfProcess.Failure, Me.sequence, Me.CompCode, Me.CompName, Me.prgName,
                                      String.Format("{0} => 發生錯誤：{1}", Nowtime.Value.ToString("yyyy-MM-dd HH:mm:ss"),
                                                                      ex.Message))
            End If
            If Not String.IsNullOrEmpty(msg) Then
                WriteLog(String.Format("{0}\{1}\{2}", logPathName,
                                                      String.Format(Nowtime.Value.ToString("yyyyMMdd")),
                                                      LogFileName),
                                                    String.Format("{0} => {1}", Nowtime.Value.ToString("HH:mm:ss"), msg))
            End If

        Catch exp As Exception
        Finally
            If Nowtime IsNot Nothing Then
                Nowtime.Value = Nothing
                Nowtime.Dispose()
                Nowtime = Nothing
            End If
        End Try
    End Sub
    Private Function ProcWebPost() As Boolean
        Try
            If ExecuteWeb(System.Net.WebRequestMethods.Http.Post) Then
                Interlocked.Decrement(WaitProcessCount)
                WriteMsg(Nothing, String.Format("測試成功：{0}", url))
                raiseStatusEven(StatusOfProcess.Success,
                                    Me.sequence, Me.CompCode,
                                    Me.CompName, Me.prgName,
                                    String.Format("{0} => {1}", Date.Now.ToString("yyyy-MM-dd HH:mm:ss"), "測試成功！"))

            Else
                Interlocked.Exchange(WaitProcessCount, 0)
            End If
        Catch ex As Exception
            WriteMsg(ex, Nothing)
            Interlocked.Exchange(WaitProcessCount, 0)
        End Try
        Return True
    End Function
    Private Function ProcWebGet() As Boolean
        Try
            If ExecuteWeb(System.Net.WebRequestMethods.Http.Get) Then
                Interlocked.Decrement(WaitProcessCount)
                WriteMsg(Nothing, String.Format("測試成功：{0}", url))
                'If Not isStopAll Then                
                raiseStatusEven(StatusOfProcess.Success,
                                    Me.sequence, Me.CompCode,
                                    Me.CompName, Me.prgName,
                                    String.Format("{0} => {1}", Date.Now.ToString("yyyy-MM-dd HH:mm:ss"), "測試成功！"))
                'End If
            Else
                'WriteMsg(Nothing, "Processing WebGet  In Failure.")

                Interlocked.Exchange(WaitProcessCount, 0)
            End If
        Catch ex As Exception
            WriteMsg(ex, Nothing)
            Interlocked.Exchange(WaitProcessCount, 0)
        End Try
        Return True
    End Function
    Private Function ProcWebService() As Boolean
        Dim webServicePara As New ThreadLocal(Of ArrayList)
        Try
            If CallWebService Is Nothing Then
                CallWebService = New WebServiceObject(Me.url, Me.Timeout * 1000)
            End If
            webServicePara.Value = New ArrayList
            If Not String.IsNullOrEmpty(Me.WebServicePara) Then
                For Each ob As Object In Me.WebServicePara
                    webServicePara.Value.Add(ob)
                Next
            End If
            If CallWebService.Execute(Me.WebSrvMethod, webServicePara.Value.ToArray) Then
                Interlocked.Decrement(WaitProcessCount)
                If WaitProcessCount = 0 Then
                    WriteMsg(Nothing, String.Format("測試成功：{0}",
                             WebSrvMethod))
                    raiseStatusEven(StatusOfProcess.Success,
                                    Me.sequence, Me.CompCode,
                                    Me.CompName, Me.prgName,
                                    String.Format("{0} => {1}", Date.Now.ToString("yyyy-MM-dd HH:mm:ss"), "測試成功！"))

                End If
            Else
                WriteMsg(Nothing, String.Format("Processing webservice of {0}  was unsuccessful.", WebSrvMethod))
                raiseStatusEven(StatusOfProcess.Failure,
                                  Me.sequence, Me.CompCode,
                                  Me.CompName, Me.prgName,
                                  String.Format("{0} => {1}", Date.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                                                "Processing webservice of {0}  was unsuccessful."))
                Interlocked.Exchange(WaitProcessCount, 0)
            End If
            webServicePara.Value.Clear()
        Catch ex As Exception
            WriteMsg(ex, Nothing)
            Interlocked.Exchange(WaitProcessCount, 0)
            If CallWebService IsNot Nothing Then
                CallWebService.Dispose()
                CallWebService = Nothing
            End If
        Finally
            If webServicePara IsNot Nothing Then
                If webServicePara.Value IsNot Nothing Then
                    webServicePara.Value.Clear()
                End If
                webServicePara.Dispose()
                webServicePara = Nothing
            End If
        End Try
        Return True
    End Function
    Private Function ProcDBAccess() As Boolean
        Dim corInsSql As New ThreadLocal(Of String)
        Dim corGetSql As New ThreadLocal(Of String)
        Dim Sequence As New ThreadLocal(Of String)
        Dim sw As New ThreadLocal(Of System.Diagnostics.Stopwatch)
        Try
            Sequence.Value = getSequence()
            corInsSql.Value = analyzeInsSql(Sequence.Value)
            corGetSql.Value = analyzeGetSql(Sequence.Value)

            Using cnDB
                If cnDB Is Nothing Then
                    cnDB = New OracleConnection(connectString)
                End If
                If cnDB.State = ConnectionState.Closed Then
                    cnDB.Open()
                End If
                Using cmd As OracleCommand = cnDB.CreateCommand
                    cmd.CommandText = corInsSql.Value
                    cmd.ExecuteNonQuery()
                    WriteMsg(Nothing, String.Format("寫入資料庫成功：{0}",
                                                    corInsSql.Value))
                    cmd.CommandText = corGetSql.Value
                    sw.Value = New Stopwatch
                    sw.Value.Start()
                    Do
                        Select Case cmd.ExecuteScalar                          
                            Case 0
                                If sw.Value.Elapsed.TotalSeconds >= Me.Timeout Then
                                    WriteMsg(New Exception(String.Format("回應失敗（timeout）：{0}",
                                                          corGetSql.Value)), Nothing)
                                    sw.Value.Stop()
                                    Exit Do
                                End If
                            Case 1
                                WriteMsg(Nothing, String.Format("回應成功：{0}",
                                                 corGetSql.Value))
                                raiseStatusEven(StatusOfProcess.Success,
                                                       Me.sequence, Me.CompCode,
                                                       Me.CompName, Me.prgName,
                                                       String.Format("{0} => {1}", Date.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                                                                           "回應成功：" & corGetSql.Value))
                                sw.Value.Stop()
                                Exit Do
                            Case 2
                                WriteMsg(New Exception(String.Format("回應失敗（指令失敗）：{0}", corGetSql.Value)), Nothing)
                                sw.Value.Stop()
                                Exit Do
                            Case Else
                                WriteMsg(New Exception(String.Format("回應未知：{0}", corGetSql.Value)), Nothing)
                                sw.Value.Stop()
                                Exit Do
                        End Select
                        Thread.Sleep(100)
                    Loop
                    cmd.Dispose()
                End Using
            End Using
            Interlocked.Decrement(WaitProcessCount)
        Catch ex As Exception
            WriteMsg(ex, Nothing)
            Interlocked.Exchange(WaitProcessCount, 0)
        Finally
            If cnDB IsNot Nothing Then
                cnDB.Close()
                cnDB.Dispose()
                cnDB = Nothing
            End If
            If Sequence IsNot Nothing Then
                Sequence.Dispose()
                Sequence = Nothing
            End If
            If corGetSql IsNot Nothing Then
                corGetSql.Dispose()
                corGetSql = Nothing
            End If
            If corInsSql IsNot Nothing Then
                corInsSql.Dispose()
                corInsSql = Nothing
            End If
            If sw IsNot Nothing Then
                sw.Value = Nothing
                sw.Dispose()
                sw = Nothing
            End If

        End Try
        Return True
    End Function
    Private Sub raiseStatusEven(Status As StatusOfProcess,
                              ByVal sequence As Integer,
                              ByVal comId As Integer,
                              ByVal compName As String,
                              ByVal prgName As String,
                              ByVal msg As String)
        RaiseEvent StatusChange(Status, sequence, comId, compName, prgName, msg)
    End Sub
    Private Sub WaitProc(ByVal state As Object)
        GC.Collect()
        'SyncLock lckObject
        If isRunning OrElse isStopAll Then Exit Sub
        tmr.Change(System.Threading.Timeout.Infinite, System.Threading.Timeout.Infinite)

        isRunning = True
        WriteMsg(Nothing, "測試中！")
        raiseStatusEven(StatusOfProcess.Processing,
                                                    Me.sequence, Me.CompCode,
                                                    Me.CompName, Me.prgName,
                                                    String.Format("{0} => {1}", Date.Now.ToString("yyyy-MM-dd HH:mm:ss"), "測試中！"))

        Try
            If (firstRun) OrElse
                 (Integer.Parse(Date.Now.ToString("yyyyMMdd")) > Integer.Parse(startRunTime.ToString("yyyyMMdd"))) Then
                startRunTime = Date.Now
                DeleteLog()
            End If
            firstRun = False
            Try
                For Each oType As OverseeType In lstMonitortype
                    Interlocked.Increment(WaitProcessCount)
                    Select Case oType
                        Case OverseeType.DBAccess
                            ProcDBAccess()
                        Case OverseeType.WebService
                            ProcWebService()
                        Case OverseeType.WebGet
                            ProcWebGet()
                        Case OverseeType.WebPost
                            ProcWebPost()
                    End Select
                Next
            Catch ex As Exception
                WriteMsg(ex, Nothing)
                Interlocked.Exchange(WaitProcessCount, 0)
            End Try
        Catch ex As Exception
            WriteMsg(ex, Nothing)
        Finally
            isRunning = False
            If Not isStopAll Then
                isRunning = False
                tmr.Change(Me.Runtimer * 1000, Me.Runtimer * 1000)
            End If
        End Try
        'End SyncLock
    End Sub
    Private Function ExecuteWeb(ByVal requestMethod As String) As Object
        Dim req As New Threading.ThreadLocal(Of System.Net.HttpWebRequest)
        Dim response As New Threading.ThreadLocal(Of System.Net.WebResponse)
        Dim result As New ThreadLocal(Of Object)
        Try
            Try
                Select Case requestMethod.ToUpper
                    Case System.Net.WebRequestMethods.Http.Post
                        req.Value = System.Net.HttpWebRequest.Create(Me.url)
                        req.Value.Method = System.Net.WebRequestMethods.Http.Post
                        req.Value.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"
                        req.Value.ContentType = "application/x-www-form-urlencoded"
                        req.Value.ContentLength = System.Text.Encoding.UTF8.GetByteCount(Me.WebPostPara)
                        req.Value.KeepAlive = False
                        If Me.Timeout > 0 Then req.Value.Timeout = Me.Timeout * 1000
                        If Me.Timeout > 0 Then req.Value.ReadWriteTimeout = Me.Timeout * 1000
                        Using requestWriter As System.IO.StreamWriter =
                                New System.IO.StreamWriter(req.Value.GetRequestStream)
                            Try
                                requestWriter.Write(Me.WebPostPara, 0,
                                                    System.Text.Encoding.UTF8.GetByteCount(Me.WebPostPara))
                                requestWriter.Flush()
                                requestWriter.Close()
                            Catch exWriter As Exception
                                Throw exWriter
                            End Try
                        End Using
                    Case System.Net.WebRequestMethods.Http.Get
                        If String.IsNullOrEmpty(Me.WebGetPara) Then
                            req.Value = System.Net.HttpWebRequest.Create(url)
                        Else
                            req.Value = System.Net.HttpWebRequest.Create(String.Format("{0}?{1}", url, Me.WebGetPara))
                        End If
                        req.Value.Method = System.Net.WebRequestMethods.Http.Get
                        req.Value.ContentType = "application/x-www-form-urlencoded"
                        If Me.Timeout > 0 Then req.Value.Timeout = Me.Timeout * 1000
                        'req.Value.UserAgent = "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.2; .NET CLR 1.0.3705;)"
                        req.Value.KeepAlive = False
                End Select


                response.Value = req.Value.GetResponse
                If CType(response.Value, System.Net.HttpWebResponse).StatusCode <> Net.HttpStatusCode.OK Then
                    Throw New Exception("Response is unsuccessful.")
                Else
                    Using responseReader As System.IO.StreamReader =
                                        New System.IO.StreamReader(response.Value.GetResponseStream, New UTF8Encoding)
                        result.Value = responseReader.ReadToEnd
                        result.Value = result.Value.ToString.StartsWith("True", StringComparison.OrdinalIgnoreCase)
                        responseReader.Close()
                        responseReader.Dispose()
                        Return result.Value
                    End Using
                End If

            Catch webEx As System.Net.WebException
                WriteMsg(webEx, Nothing)

                Interlocked.Exchange(WaitProcessCount, 0)
            Catch ex As Exception
                WriteMsg(ex, Nothing)
                Interlocked.Exchange(WaitProcessCount, 0)
                Return False
            Finally

            End Try

        Catch ex As Exception
        Finally
            If req IsNot Nothing Then
                req.Value = Nothing
                req.Dispose()
                req = Nothing
            End If
            If response IsNot Nothing Then
                response.Value = Nothing
                response.Dispose()
                response = Nothing
            End If
            If result IsNot Nothing Then
                result.Value = Nothing
                result.Dispose()
                result = Nothing
            End If
        End Try
    End Function
    Public Sub StopRun()

        isStopAll = True
        If tmr IsNot Nothing Then
            tmr.Change(Threading.Timeout.Infinite, Threading.Timeout.Infinite)
        End If

        Try
            Do
                If (Not isRunning) AndAlso (WaitProcessCount <= 0) Then
                    If tmr IsNot Nothing Then
                        tmr.Change(Threading.Timeout.Infinite, Threading.Timeout.Infinite)
                        tmr.Dispose()
                        tmr = Nothing
                    End If
                    Exit Do
                End If
                Thread.Sleep(1000)
            Loop
        Catch ex As Exception

        End Try

    End Sub

#Region "IDisposable Support"
    Private disposedValue As Boolean ' 偵測多餘的呼叫

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: 處置 Managed 狀態 (Managed 物件)。
                StopRun()
                'isStopAll = True
                'Do
                '    If (Not isRunning) AndAlso (WaitProcessCount = 0) Then
                '        If tmr IsNot Nothing Then
                '            tmr.Change(Threading.Timeout.Infinite, Threading.Timeout.Infinite)
                '            tmr.Dispose()
                '            tmr = Nothing
                '        End If
                '        Exit Do
                '    End If
                '    Thread.Sleep(100)
                'Loop
                OracleConnection.ClearAllPools()
                If CallWebService IsNot Nothing Then
                    CallWebService.Dispose()
                    CallWebService = Nothing
                End If
                conBuilder = Nothing
                If lstMonitortype IsNot Nothing Then
                    lstMonitortype.Clear()
                    lstMonitortype = Nothing
                End If
                If cnDB IsNot Nothing Then                    
                    cnDB.Close()
                    cnDB.Dispose()
                    cnDB = Nothing
                End If

                If lckObject IsNot Nothing Then
                    lckObject = Nothing
                End If
                RemoveHandler Me.StatusChange, AddressOf raiseStatusEven
            End If

            ' TODO: 釋放 Unmanaged 資源 (Unmanaged 物件) 並覆寫下面的 Finalize()。
            ' TODO: 將大型欄位設定為 null。
        End If
        Me.disposedValue = True
    End Sub

    ' TODO: 只有當上面的 Dispose(ByVal disposing As Boolean) 有可釋放 Unmanaged 資源的程式碼時，才覆寫 Finalize()。
    'Protected Overrides Sub Finalize()
    '    ' 請勿變更此程式碼。在上面的 Dispose(ByVal disposing As Boolean) 中輸入清除程式碼。
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' 由 Visual Basic 新增此程式碼以正確實作可處置的模式。
    Public Sub Dispose() Implements IDisposable.Dispose
        ' 請勿變更此程式碼。在以上的 Dispose 置入清除程式碼 (ByVal 視為布林值處置)。
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class
