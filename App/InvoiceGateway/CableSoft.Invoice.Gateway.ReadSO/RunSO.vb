Imports System.Threading
Imports System.Data.OracleClient

Public Class RunSO
    Implements IDisposable

    Private SO As CableSoft.Invoice.Gateway.SOInfo.SO
    Private cn As OracleConnection = Nothing
    Private ti As RegisteredWaitHandle = Nothing
    Private evn As AutoResetEvent = Nothing
    Private mtxWait As Mutex = Nothing
    Private IsProcessing As Boolean = False
    Private StopProcess As Boolean = False
    Public Sub New(ByVal SO As CableSoft.Invoice.Gateway.SOInfo.SO)
        SO = Me.SO
        mtxWait = New Mutex()
    End Sub
    Public Sub Run()
        If evn IsNot Nothing Then
            evn.Dispose()
            evn = Nothing
        End If
        evn = New AutoResetEvent(False)
        ti = ThreadPool.RegisterWaitForSingleObject(evn, _
                                                           New WaitOrTimerCallback(AddressOf WaitProc), _
                                                           Nothing, TimeSpan.FromMinutes(SO.MonitorTime), False)
        evn.Set()
    End Sub
    Private Sub WaitProc(ByVal state As Object, ByVal timedOut As Boolean)
        mtxWait.WaitOne()
        If (StopProcess) OrElse (IsProcessing) Then
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
                                    cmd.CommandText = String.Format("UPDATE SEND_NSTV SET CMD_STATUS='P' " & _
                                                                    " WHERE SEQNO={0}", rw.Item("SEQNO"))
                                    cmd.ExecuteNonQuery()
                                End Using
                                Dim aryDataRow As New Dictionary(Of String, Object)(System.StringComparer.OrdinalIgnoreCase)
                                aryDataRow.Clear()
                                For Each colName As DataColumn In ds.Tables(0).Columns
                                    aryDataRow.Add(colName.ColumnName.ToUpper, rw.Item(colName.ColumnName))
                                Next
                                aryDataRow.Add("SendLog".ToUpper, Nothing)
                                aryDataRow.Add("RecvLog".ToUpper, Nothing)
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
    Public Sub StopGateway()
        If SO IsNot Nothing Then
            SO.Dispose()
            SO = Nothing
        End If
        If ti IsNot Nothing Then
            ti.Unregister(evn)

        End If
    End Sub

#Region "IDisposable Support"
    Private disposedValue As Boolean ' 偵測多餘的呼叫

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: 處置 Managed 狀態 (Managed 物件)。
                If SO IsNot Nothing Then
                    SO.Dispose()
                    SO = Nothing
                End If
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
