Imports System.Data.OracleClient
Imports CableSoft.Nstv.Log
Imports System.Threading
Imports CableSoft.Nagra.BuilderLowerCmd
Imports CableSoft.Nstv.Encry

Public Class SO
    Implements IDisposable
  
    Public Property StopProcess As Boolean = False
    Private cn As OracleConnection = Nothing
    Private ti As RegisteredWaitHandle = Nothing
    Private evn As AutoResetEvent = Nothing
    Private WaitProcessCount As Integer
    Private ReadTime As Integer = 30
    Private ProcCmdCount As Integer = 100
    Private QuerySQL As String = Nothing

    Private sckClient() As CableSoft.CardLess.SocketClient.Client
    Private mtxWait As Mutex = Nothing

    Private mtxCount As Mutex = Nothing
    Private CompCode As Integer
    Private IsProcessing As Boolean = False
    Private SocketIndex As Integer = 0
    Private tbSO As DataTable = Nothing
    Private tbSystem As DataTable = Nothing
    Private tbHighCmd As DataTable = Nothing
    Private tbLowCmd As DataTable = Nothing
    Private LowerCmd As LowCmd = Nothing
    Private tbNagra As DataTable = Nothing
    Private tbSourceId As DataTable = Nothing
    Private RegEndDate As Date = Now
    Private HaveWriteRegisterError As Boolean = False
    Public Property ConnectString As String
    Private ReadDBErrorCount As Integer = 0
    Private Enum calculateMode
        Add = 0
        minus = 1
    End Enum


    Public Sub New(ByVal CompCode As Integer, ByVal RegEndDate As Date,
                   ByVal xmlFile As String, ByVal xmlIsEncrypt As Boolean,
                   ByVal socketArray() As CableSoft.CardLess.SocketClient.Client)
        Try
            Using xmlIO As New CableSoft.CardLess.XMLFileIO.XmlFileIO(xmlFile)
                tbSO = xmlIO.ReadSO(xmlIsEncrypt)
                tbSystem = xmlIO.ReadSystem(xmlIsEncrypt)
                'tbHighCmd = xmlIO.ReadHightCmd(xmlIsEncrypt)
                'tbLowCmd = xmlIO.ReadLowCmd(xmlIsEncrypt)
                tbSourceId = xmlIO.ReadServer(xmlIsEncrypt, "Server")
                tbNagra = xmlIO.ReadNagra(xmlIsEncrypt)
            End Using
            Me.CompCode = CompCode
            Me.RegEndDate = RegEndDate
            Me.ReadTime = tbSystem.Rows(0).Item("ReadTime")
            Me.ProcCmdCount = tbSystem.Rows(0).Item("ProcCmdCount")
            Dim bduConnectString As New OracleConnectionStringBuilder
            bduConnectString.Pooling = True
            bduConnectString.PersistSecurityInfo = True
            bduConnectString.DataSource = tbSO.Select("IsSelect='1' AND COMPCODE='" & CompCode & "'").ToList.Item(0).Item("SID")
            bduConnectString.UserID = tbSO.Select("IsSelect='1' AND COMPCODE='" & CompCode & "'").ToList.Item(0).Item("User")
            bduConnectString.Password = tbSO.Select("IsSelect='1' AND COMPCODE='" & CompCode & "'").ToList.Item(0).Item("Password")
            bduConnectString.MaxPoolSize = tbSystem.Rows(0).Item("DBMaxPool")
            bduConnectString.MinPoolSize = tbSystem.Rows(0).Item("DBMinPool")
            bduConnectString.LoadBalanceTimeout = tbSystem.Rows(0).Item("DBPoolLiveTime")
            cn = New OracleConnection(bduConnectString.ConnectionString)
            Me.ConnectString = bduConnectString.ConnectionString
            Select Case Integer.Parse(tbNagra.Rows(0).Item("ProcKind"))

                Case 0
                    '全部指令


                    QuerySQL = String.Format("SELECT * FROM ({0}) WHERE ROWNUM<={1}",
                                             String.Format("SELECT * FROM SEND_NAGRA  WHERE  COMPCODE={0}   " & _
                                                         " AND CMD_STATUS='W' AND GATEWAYTYPE = 1 " & _
                                                         " AND NVL(PROCESSINGDATE,SYSDATE-1) <= SYSDATE " & _
                                                         " ORDER BY  NVL(PROCESSINGDATE,SYSDATE-1), SEQNO",
                                            Me.CompCode), ProcCmdCount)


                Case 1
                    '即時指令

                    QuerySQL = String.Format("SELECT * FROM ( {0} ) WHERE ROWNUM<={1} ",
                                            String.Format("SELECT * FROM SEND_NAGRA  WHERE COMPCODE={0}   " & _
                                                         " AND CMD_STATUS='W' " & _
                                                         " AND PROCESSINGDATE IS NULL " & _
                                                         " ORDER BY  SEQNO",
                                     Me.CompCode), ProcCmdCount)
                Case 2
                    '預約指令

                    QuerySQL = String.Format("SELECT * FROM ( {0} ) WHERE ROWNUM<={1} ",
                                            String.Format("SELECT * FROM SEND_NAGRA  WHERE  COMPCODE={0}   " & _
                                                      " AND CMD_STATUS='W' " & _
                                                      " AND PROCESSINGDATE <= SYSDATE " & _
                                                      " AND PROCESSINGDATE Is Not Null " & _
                                                      " ORDER BY  PROCESSINGDATE, SEQNO",
                                 Me.CompCode), ProcCmdCount)
            End Select



            Me.sckClient = socketArray

            mtxWait = New Mutex()
            mtxCount = New Mutex()

        Catch ex As Exception
            NstvLog.WriteErrorLog(ex, "SO系統台建立失敗")
        End Try

    End Sub
  
    Public Sub Run()
        If evn IsNot Nothing Then
            evn.Dispose()
            evn = Nothing
        End If
        HaveWriteRegisterError = False
        ReadDBErrorCount = 0
        evn = New AutoResetEvent(False)
        ti = ThreadPool.RegisterWaitForSingleObject(evn, _
                                                           New WaitOrTimerCallback(AddressOf WaitProc), _
                                                           Nothing, TimeSpan.FromSeconds(ReadTime), False)
        evn.Set()
    End Sub
    Private Sub WaitProc(ByVal state As Object, ByVal timedOut As Boolean)
        mtxWait.WaitOne()
        If (StopProcess) OrElse (WaitProcessCount > 0) OrElse
                (IsProcessing) Then
            mtxWait.ReleaseMutex()
            Exit Sub
        End If
        If (DateTime.Today > RegEndDate) Then
            If Not HaveWriteRegisterError Then
                NstvLog.WriteErrorLog(New Exception("使用日期已過期！"),
                                  String.Format("使用日期:{0}", RegEndDate.ToShortDateString))
            End If
            HaveWriteRegisterError = True
            mtxWait.ReleaseMutex()
            Exit Sub
        End If
        IsProcessing = True
        Try
            If (cn Is Nothing) AndAlso (Not String.IsNullOrEmpty(Me.ConnectString)) Then
                cn = New OracleConnection(Me.ConnectString)
            End If
            If cn.State <> ConnectionState.Open Then
                cn.Open()            
            End If
            '下面這段註解可以稍為省掉一點時間
            '--------------------------------------------------------------------------------------------------------------------------------
            'Using cmd As OracleCommand = cn.CreateCommand
            '    cmd.CommandText = QuerySQL
            '    Using dr As OracleDataReader = cmd.ExecuteReader(CommandBehavior.SequentialAccess)
            '        Try
            '            Using cmdProceing As OracleCommand = cn.CreateCommand
            '                cmdProceing.CommandText = String.Format("UPDATE SEND_NSTV  SET CMD_STATUS='P' " & _
            '                                                " WHERE SEQNO IN ({0})", String.Format("SELECT SEQNO FROM ({0})", QuerySQL))
            '                WaitProcessCount = cmdProceing.ExecuteNonQuery
            '            End Using
            '        Catch ex As Exception
            '            NstvLog.WriteErrorLog(ex, Nothing)
            '            WaitProcessCount = 0
            '            Exit Sub
            '        End Try
            '        Try
            '            While dr.Read
            '                Dim aryDataRow As New Dictionary(Of String, Object)(System.StringComparer.OrdinalIgnoreCase)
            '                aryDataRow.Clear()
            '                For i As Int32 = 0 To dr.FieldCount - 1
            '                    aryDataRow.Add(dr.GetName(i).ToUpper, dr.GetValue(i))
            '                Next
            '                aryDataRow.Add("SendLog".ToUpper, Nothing)
            '                aryDataRow.Add("RecvLog".ToUpper, Nothing)
            '                ThreadPool.QueueUserWorkItem(AddressOf ProcessSocket, aryDataRow)
            '            End While
            '        Catch ex As Exception
            '            NstvLog.WriteErrorLog(ex, Nothing)
            '            WaitProcessCount = 0
            '            Exit Sub
            '        End Try
            '    End Using      
            'End Using
            'Exit Sub
            '--------------------------------------------------------------------------------------------------------------------------------
            Using ds As New DataSet
                'Debug.Print(QuerySQL)
                Using dpr As New OracleDataAdapter(QuerySQL, cn)
                    dpr.Fill(ds)
                    If ds.Tables(0).Rows.Count > 0 Then
                        WaitProcessCount = ds.Tables(0).Rows.Count
                        'Dim tra As OracleTransaction = cn.BeginTransaction
                        Try
                            For Each rw As DataRow In ds.Tables(0).Rows
                                Using cmd As OracleCommand = cn.CreateCommand
                                    '            cmd.Transaction = tra
                                    cmd.CommandText = String.Format("UPDATE SEND_NAGRA SET CMD_STATUS='P' " & _
                                                                    " WHERE SEQNO={0}", rw.Item("SEQNO"))
                                    cmd.ExecuteNonQuery()
                                End Using
                                Dim aryDataRow As New Dictionary(Of String, Object)(System.StringComparer.OrdinalIgnoreCase)
                                aryDataRow.Clear()
                                For Each colName As DataColumn In ds.Tables(0).Columns
                                    aryDataRow.Add(colName.ColumnName.ToUpper, rw.Item(colName.ColumnName))
                                Next
                                'Add with Send and Receive information to nagracmdlog table By Kin 2015/09/11
                                aryDataRow.Add("RetryCount", 0)
                                aryDataRow.Add("CATRANNUM", "")
                                aryDataRow.Add("GWTRANNUM", "")
                                aryDataRow.Add("SendCmdText", "")
                                aryDataRow.Add("RecvCmdText", "")
                                aryDataRow.Add("ACK", "")
                                aryDataRow.Add("SendTime", Date.Now)
                                aryDataRow.Add("RecvTime", Date.Now)
                                aryDataRow.Add("LowCmd", "126")
                                aryDataRow.Add("HighCmd", rw.Item("HIGH_LEVEL_CMD_ID"))
                                'aryDataRow.Add("SendLog".ToUpper, Nothing)
                                'aryDataRow.Add("RecvLog".ToUpper, Nothing)
                                ThreadPool.QueueUserWorkItem(AddressOf ProcessSocket, aryDataRow)
                            Next
                            'tra.Commit()
                        Catch ex As Exception
                            NstvLog.WriteErrorLog(ex, Nothing)
                            'tra.Rollback()
                            WaitProcessCount = 0
                            Exit Sub
                        Finally
                            'tra.Dispose()
                            'tra = Nothing
                        End Try

                        'For Each rw As DataRow In ds.Tables(0).Rows
                        '    Dim aryDataRow As New Dictionary(Of String, Object)
                        '    aryDataRow.Clear()
                        '    For Each colName As DataColumn In ds.Tables(0).Columns
                        '        aryDataRow.Add(colName.ColumnName.ToUpper, rw.Item(colName.ColumnName))
                        '    Next
                        '    aryDataRow.Add("SendLog".ToUpper, Nothing)
                        '    aryDataRow.Add("RecvLog".ToUpper, Nothing)
                        '    ThreadPool.QueueUserWorkItem(AddressOf ProcessSocket, aryDataRow)
                        'Next
                    End If
                End Using
            End Using
            Interlocked.Exchange(ReadDBErrorCount, 0)
        Catch ex As Exception

            If ReadDBErrorCount = 100 Then
                Interlocked.Increment(ReadDBErrorCount)
                NstvLog.WriteErrorLog(ex,
                                      "讀取資料庫已有一段時間失敗，系統不再寫入錯誤訊息，請檢查資料庫連線")
            Else
                If ReadDBErrorCount <= 100 Then
                    NstvLog.WriteErrorLog(ex, Nothing)
                    Interlocked.Increment(ReadDBErrorCount)
                End If

            End If

        Finally
            If cn IsNot Nothing Then
                cn.Close()
            End If
            IsProcessing = False
            mtxWait.ReleaseMutex()
        End Try
    End Sub
    Private Sub calculateWaitCount(ByVal Mode As calculateMode)
        mtxCount.WaitOne()
        Try
            If Mode = calculateMode.Add Then
                Interlocked.Increment(WaitProcessCount)
            Else
                Interlocked.Decrement(WaitProcessCount)
            End If
        Finally
            mtxCount.ReleaseMutex()
        End Try
    End Sub
    Private Sub ProcessCmd(ByVal state As Object)

        Dim retRW As New ThreadLocal(Of Dictionary(Of String, Object))
        Dim localprocInfo As New ThreadLocal(Of procInfo)

        localprocInfo.Value = CType(state, procInfo)
        retRW.Value = localprocInfo.Value.highCmdData
        Try
            '傳入的Socket如果是Nothing,代表全部都斷線了
            If localprocInfo.Value.SocketClient Is Nothing Then
                retRW.Value.Item("CMD_STATUS".ToUpper) = "E"
                retRW.Value.Item("ERR_CODE".ToUpper) = Int32.Parse(CardLess.SocketClient.Client.CableSoftError.SocketAllDisconnect)
                retRW.Value.Item("ERR_MSG".ToUpper) = [Enum].GetName(GetType(CardLess.SocketClient.Client.CableSoftError),
                                                             CardLess.SocketClient.Client.CableSoftError.SocketAllDisconnect)

            Else
                '送出高階命令
                retRW.Value = localprocInfo.Value.SocketClient.SendHighCmd(retRW.Value)
                '如果送出高階命令，回來的Cmd_Status="P"代表資料未送出，再給ProcessSocket重新處理
                If (retRW.Value IsNot Nothing) AndAlso ((retRW.Value.Item("CMD_STATUS".ToUpper) = "W") OrElse
                                                        (retRW.Value.Item("CMD_STATUS".ToUpper) = "P")) Then
                    'retRW.Value.Item("SendLog".ToUpper) = Nothing
                    'retRW.Value.Item("RecvLog".ToUpper) = Nothing
                    ThreadPool.QueueUserWorkItem(AddressOf ProcessSocket, retRW.Value)
                    Interlocked.Increment(WaitProcessCount)
                End If
            End If
        Catch ex As Exception

            retRW.Value.Item("CMD_STATUS") = "E"
            retRW.Value.Item("ERR_CODE") = Int32.Parse(CardLess.SocketClient.Client.CableSoftError.NstvFormatError)
            retRW.Value.Item("ERR_MSG") = [Enum].GetName(GetType(CardLess.SocketClient.Client.CableSoftError),
                                                         CardLess.SocketClient.Client.CableSoftError.NstvFormatError)
            NstvLog.WriteErrorLog(ex, String.Format("公司別 = {0}, SEQNO={1}",
                                                retRW.Value.Item("CompCode".ToUpper),
                                                retRW.Value.Item("SEQNO")))

        Finally
            '更新高階命令
            UpdateHighCmd(retRW.Value)
            retRW.Dispose()
            localprocInfo.Value.Dispose()
            localprocInfo.Dispose()
            Interlocked.Decrement(WaitProcessCount)
        End Try


    End Sub
    Private Sub UpdateHighCmd(ByVal rw As Dictionary(Of String, Object))
        Dim localCn As New ThreadLocal(Of OracleConnection)
        Dim pCmd_Status As New ThreadLocal(Of OracleParameter)
        Dim pUPDTIME As New ThreadLocal(Of OracleParameter)
        Dim pERR_CODE As New ThreadLocal(Of OracleParameter)
        Dim pERR_MSG As New ThreadLocal(Of OracleParameter)
        Dim pRETURNDATA As New ThreadLocal(Of OracleParameter)
        Dim pSeqNo As New ThreadLocal(Of OracleParameter)
        Try
            localCn.Value = New OracleConnection(Me.ConnectString)
            If (rw.Item("Cmd_Status".ToUpper).ToString.ToUpper = "W".ToUpper) OrElse
                (rw.Item("Cmd_Status".ToUpper).ToString.ToUpper = "P".ToUpper) Then
                Exit Sub
            End If
            If localCn.Value.State = ConnectionState.Closed Then
                localCn.Value.Open()
            End If

            Dim aSQL As String = "UPDATE SEND_NAGRA SET CMD_STATUS = :CMD_STATUS," & _
                            "UPDTIME =:UPDTIME,ERR_CODE = :ERR_CODE,ERR_MSG = :ERR_MSG," & _
                            "RETURNDATA = :RETURNDATA WHERE SEQNO = :SEQNO"

            pCmd_Status.Value = New OracleParameter("CMD_STATUS", OracleType.VarChar)
            pUPDTIME.Value = New OracleParameter("UPDTIME", OracleType.DateTime)
            pERR_CODE.Value = New OracleParameter("ERR_CODE", OracleType.VarChar)
            pERR_MSG.Value = New OracleParameter("ERR_MSG", OracleType.VarChar)
            pRETURNDATA.Value = New OracleParameter("RETURNDATA", OracleType.VarChar)
            pSeqNo.Value = New OracleParameter("SEQNO", OracleType.Number)

            pCmd_Status.Value.Value = rw.Item("CMD_STATUS".ToUpper)
            pUPDTIME.Value.Value = Date.Parse(rw.Item("UPDTIME".ToUpper))
            pERR_CODE.Value.Value = rw.Item("ERR_CODE".ToUpper)
            pERR_MSG.Value.Value = rw.Item("ERR_MSG".ToUpper)
            pRETURNDATA.Value.Value = rw.Item("RETURNDATA".ToUpper)
            pSeqNo.Value.Value = Integer.Parse(rw.Item("SEQNO".ToUpper))

            Using cmd As OracleCommand = localCn.Value.CreateCommand
                cmd.Parameters.Clear()
                cmd.CommandText = aSQL
                cmd.Parameters.Add(pCmd_Status.Value)
                cmd.Parameters.Add(pUPDTIME.Value)
                cmd.Parameters.Add(pERR_CODE.Value)
                cmd.Parameters.Add(pERR_MSG.Value)
                cmd.Parameters.Add(pRETURNDATA.Value)
                cmd.Parameters.Add(pSeqNo.Value)
                cmd.ExecuteNonQuery()
                cmd.Parameters.Clear()
            End Using
            aSQL = "Insert Into nagracmdlog (CompCode,CaTranNum,GWTranNum,HighCmd,LowCmd, " & _
                                                    "ICC,STB,SendCmdText,RecvCmdText,Ack,Operator,updtime,sendtime,recvtime ) " & _
                                                    "Values (:Compcode,:CaTranNum,:GWTranNum,:HighCmd,:LowCmd, " & _
                                                    ":ICC,:STB,:SendCmdText,:RecvCmdText,:Ack,'cardless',sysdate,:sendtime,:recvtime)"

            Try
                Using cmd As OracleCommand = localCn.Value.CreateCommand
                    cmd.CommandText = aSQL
                    cmd.Parameters.Add("Compcode", OracleType.Number).Value = rw.Item("CompCode")
                    cmd.Parameters.Add("CaTranNum", OracleType.VarChar).Value = rw.Item("CaTranNum")
                    cmd.Parameters.Add("GWTranNum", OracleType.VarChar).Value = rw.Item("GWTranNum")
                    cmd.Parameters.Add("HighCmd", OracleType.VarChar).Value = rw.Item("HighCmd")
                    cmd.Parameters.Add("LowCmd", OracleType.VarChar).Value = rw.Item("LowCmd")
                    cmd.Parameters.Add("ICC", OracleType.VarChar).Value = rw.Item("ICC_NO")
                    cmd.Parameters.Add("STB", OracleType.VarChar).Value = rw.Item("STB_NO")
                    cmd.Parameters.Add("SendCmdText", OracleType.VarChar).Value = rw.Item("SendCmdText")
                    cmd.Parameters.Add("RecvCmdText", OracleType.VarChar).Value = rw.Item("RecvCmdText")
                    cmd.Parameters.Add("Ack", OracleType.VarChar).Value = rw.Item("Ack")
                    cmd.Parameters.Add("sendtime", OracleType.DateTime).Value = rw.Item("sendtime")
                    cmd.Parameters.Add("recvtime", OracleType.DateTime).Value = rw.Item("recvtime")
                    cmd.ExecuteNonQuery()
                    cmd.Parameters.Clear()
                End Using
            Catch ex As Exception
                NstvLog.WriteErrorLog(ex, String.Format("記錄nagracmdlog失敗，但不影嚮命令，" &
                                                        "公司別 = {0}, SEQNO={1}",
                                                   rw.Item("CompCode".ToUpper),
                                                   rw.Item("SEQNO".ToUpper)))
            End Try


            'cmd.CommandText = String.Format("UPDATE SEND_NSTV SET CMD_STATUS='{0}', " & _
            '                    "UPDTIME=SYSDATE,ERR_CODE='{1}',ERR_MSG='{2}',ReSentTimes='{3}' " & _
            '                    " WHERE SEQNO={4}", rw.Item("CMD_STATUS".ToUpper),
            '                    rw.Item("ERR_CODE".ToUpper),
            '                    rw.Item("ERR_MSG".ToUpper),
            '                    IIf(DBNull.Value.Equals(rw.Item("ReSentTimes".ToUpper).ToString), "",
            '                        rw.Item("ReSentTimes".ToUpper).ToString),
            '                    rw.Item("SEQNO".ToUpper))
            'cmd.ExecuteNonQuery()

            'If (rw.Item("SendLog".ToUpper) IsNot Nothing) AndAlso
            '    (Not DBNull.Value.Equals(rw.Item("SendLog".ToUpper))) AndAlso
            '    (Not String.IsNullOrEmpty(rw.Item("SendLog".ToUpper).ToString)) Then
            '    For i As Int32 = 0 To rw.Item("SendLog".ToUpper).ToString.Split(",").Count - 1
            '        cmd.CommandText = String.Format("INSERT INTO SEND_NSTVSENDLOG " & _
            '                                        "(SeqNo, SocketSndMsg,SocketRcvMsg,CMDSeqNo ) VALUES " & _
            '                                        "(S_SEND_NSTVSENDLOG_SeqNo.Nextval,'{0}' " & _
            '                                        ",'{1}',{2} )",
            '                                        rw.Item("SendLog".ToUpper).ToString.Split(",")(i),
            '                                        rw.Item("RecvLog".ToUpper).ToString.Split(",")(i),
            '                                        rw.Item("SEQNO".ToUpper))
            '        cmd.ExecuteNonQuery()

            '    Next
            'End If



        Catch ex As Exception
            NstvLog.WriteErrorLog(ex, String.Format("公司別 = {0}, SEQNO={1}",
                                                    rw.Item("CompCode".ToUpper),
                                                    rw.Item("SEQNO".ToUpper)))
        Finally
            localCn.Value.Close()
            localCn.Value.Dispose()
            localCn.Dispose()
            pCmd_Status.Dispose()
            pUPDTIME.Dispose()
            pERR_CODE.Dispose()
            pERR_MSG.Dispose()
            pRETURNDATA.Dispose()
            pSeqNo.Dispose()
        End Try
    End Sub
    '處理Socket
    Private Sub ProcessSocket(ByVal state As Object)

        Dim rw As New ThreadLocal(Of Dictionary(Of String, Object))
        Dim IsAllDisconnect As New ThreadLocal(Of Boolean)
        IsAllDisconnect.Value = True
        mtxCount.WaitOne()
        rw.Value = CType(state, Dictionary(Of String, Object))
        Try
            Interlocked.Increment(SocketIndex)
            If SocketIndex > sckClient.Length - 1 Then
                SocketIndex = 0
            End If
            If (sckClient(SocketIndex) IsNot Nothing) AndAlso
                (sckClient(SocketIndex).Connected) Then
                IsAllDisconnect.Value = False
            Else
                For i As Integer = 0 To sckClient.Length - 1
                    If (sckClient(i) IsNot Nothing) AndAlso (sckClient(i).Connected) Then
                        Interlocked.Exchange(SocketIndex, i)
                        IsAllDisconnect.Value = False
                        Exit For
                    Else
                        sckClient(i).RetryConnect()
                    End If
                Next
            End If
            If IsAllDisconnect.Value Then
                ThreadPool.QueueUserWorkItem(AddressOf ProcessCmd, New procInfo(rw.Value, Nothing))
            Else
                ThreadPool.QueueUserWorkItem(AddressOf ProcessCmd, New procInfo(state, sckClient(SocketIndex)))
            End If

        Catch ex As Exception
            NstvLog.WriteErrorLog(ex, Nothing)
        Finally
            IsAllDisconnect.Value = True
            IsAllDisconnect.Dispose()
            rw.Dispose()
            mtxCount.ReleaseMutex()
        End Try
    End Sub

    Private Function IsConnect() As Boolean
        Try
            If Me.cn.State <> ConnectionState.Open Then
                cn.Open()
            End If
        Catch ex As Exception
            Return False
        End Try
        Return True
    End Function

#Region "IDisposable Support"
    Private disposedValue As Boolean ' 偵測多餘的呼叫

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                Do
                    If (Not IsProcessing) AndAlso (WaitProcessCount = 0) Then
                        If ti IsNot Nothing Then
                            ti.Unregister(Nothing)
                            ti = Nothing
                        End If
                        Exit Do
                    End If
                    Thread.Sleep(100)
                Loop
                OracleConnection.ClearAllPools()

                If cn IsNot Nothing Then
                    cn.Close()
                    cn.Dispose()
                    cn = Nothing
                End If
                If evn IsNot Nothing Then
                    evn.Set()
                    evn.Dispose()
                    evn = Nothing
                End If
                If mtxCount IsNot Nothing Then
                    mtxCount.Dispose()
                    mtxCount = Nothing
                End If

                If mtxWait IsNot Nothing Then
                    mtxWait.Dispose()
                    mtxWait = Nothing
                End If

                If tbHighCmd IsNot Nothing Then
                    tbHighCmd.Dispose()
                    tbHighCmd = Nothing
                End If
                If tbLowCmd IsNot Nothing Then
                    tbLowCmd.Dispose()
                    tbLowCmd = Nothing
                End If
                If tbNagra IsNot Nothing Then
                    tbNagra.Dispose()
                    tbNagra = Nothing
                End If
                If tbSourceId IsNot Nothing Then
                    tbSourceId.Dispose()
                    tbSourceId = Nothing
                End If
                If tbSO IsNot Nothing Then
                    tbSO.Dispose()
                    tbSO = Nothing
                End If
                If tbSystem IsNot Nothing Then
                    tbSystem.Dispose()
                    tbSystem = Nothing
                End If
                ' TODO: 處置 Managed 狀態 (Managed 物件)。
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
Friend Class procInfo
    Implements IDisposable
    Property highCmdData As Dictionary(Of String, Object)
    Property SocketClient As CableSoft.CardLess.SocketClient.Client
    Property SocketIndex As Int32 = 0
    Public Sub New(ByVal highcmdData As Dictionary(Of String, Object),
                   ByVal socketClient As CableSoft.CardLess.SocketClient.Client)
        Me.highCmdData = highcmdData
        Me.SocketClient = socketClient
    End Sub


#Region "IDisposable Support"
    Private disposedValue As Boolean ' 偵測多餘的呼叫

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: 處置 Managed 狀態 (Managed 物件)。
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
