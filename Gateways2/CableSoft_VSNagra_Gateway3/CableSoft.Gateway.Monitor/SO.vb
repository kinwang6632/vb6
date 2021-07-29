Imports System.Data.OracleClient
Imports System.Threading

Public Class SO
    Implements IDisposable
    Public Property ProcessSeqNo As UInteger = UInteger.MaxValue
    Public Property WaitUpdCmd As Dictionary(Of String, String)
    Public Property CompCode As String
    Public Property CompName As String
    Public Property WaitProcessCount As Int32 = 0
    Public Property ProcessingCount As Int32 = 0
    Public Property ConnectString As String
    Public Property IsStop As Boolean = False
    Public Property ShowUICount As Integer
    Public Property cn As OracleConnection = Nothing
    Private ti As RegisteredWaitHandle = Nothing
    Private evn As AutoResetEvent = Nothing
    Private ReadTime As Integer = 30
    Private ProcCmdCount As Integer = 100
    Private mtxWait As Mutex = Nothing
    Private QueryCountSQL As String = Nothing
    Private QuerySQL As String = Nothing
    Private guidKey As String = Nothing
    Public Property UpdateUICount As Int32 = 0
    Private Const MaxValue As Integer = 99999999
    Public Sub New(ByVal CompCode As String,
                   ByVal CompName As String,
                   ByVal ConnectString As String,
                   ByVal tbSystem As DataTable)

        Try
            Me.CompCode = CompCode
            Me.CompName = CompName
            Me.ConnectString = ConnectString
            ProcCmdCount = Integer.Parse("0" & tbSystem.Rows(0).Item("ProcCmdCount"))
            ReadTime = Integer.Parse("0" & tbSystem.Rows(0).Item("ReadTime"))
            cn = New OracleConnection(Me.ConnectString)
            mtxWait = New System.Threading.Mutex
            WaitUpdCmd = New Dictionary(Of String, String)
            Me.ShowUICount = Integer.Parse(tbSystem.Rows(0).Item("ShowUIRecord"))
            'Me.ShowUICount = Me.ProcCmdCount
            QueryCountSQL = String.Format("SELECT COUNT(CMD_STATUS) cnt ,0 FLAG  " & _
                                                                   " FROM SEND_NAGRA WHERE CMD_STATUS='W' " & _
                                                                   " AND COMPCODE = {0} " & _
                                                                   " AND GATEWAYTYPE = 1 " & _
                                                                  " GROUP BY CMD_STATUS " & _
                                                                  " UNION ALL " & _
                                                                  " SELECT COUNT(CMD_STATUS) cnt ,1 FLAG " & _
                                                                  " FROM SEND_NAGRA  " & _
                                                                  " WHERE CMD_STATUS='P' " & _
                                                                  " AND COMPCODE = {1} " & _
                                                                  " AND GATEWAYTYPE = 1 " & _
                                                                  " GROUP BY CMD_STATUS", Me.CompCode, Me.CompCode)
        Catch ex As Exception
            Throw
        End Try
    End Sub
    Private Sub WaitProc(ByVal state As Object, ByVal timedOut As Boolean)
        'If Me.UpdateUICount > 0 Then
        '    Exit Sub
        'End If
        mtxWait.WaitOne()
        guidKey = Guid.NewGuid.ToString
        Try
            If Me.IsStop = True Then
                ti.Unregister(Nothing)
                Exit Sub
            End If
            If (cn Is Nothing) AndAlso (Not String.IsNullOrEmpty(Me.ConnectString)) Then
                cn = New OracleConnection(Me.ConnectString)
            End If
            Common.UpdateSODatabase(Me, Common.SOStatus.NotUpdateImg, guidKey)

            If cn.State <> ConnectionState.Open Then
                cn.Open()
            End If

            Using cmd As OracleCommand = cn.CreateCommand
                cmd.CommandText = QueryCountSQL
                Me.WaitProcessCount = 0
                Me.ProcessingCount = 0
                Using rdr As OracleDataReader = cmd.ExecuteReader(CommandBehavior.CloseConnection)
                    While rdr.Read
                        Select Case Integer.Parse(rdr.Item("Flag"))
                            Case 0
                                Me.WaitProcessCount = Integer.Parse(rdr.Item("cnt"))
                            Case 1
                                Me.ProcessingCount = Integer.Parse(rdr.Item("cnt"))
                        End Select
                    End While
                End Using
            End Using
            Common.UpdateShowSO(Me, Common.SOStatus.Yes)
            If Me.ProcessSeqNo = UInteger.MaxValue Then
                Me.QuerySQL = String.Format("SELECT * FROM SEND_NAGRA " & _
                " WHERE ( CMD_STATUS = 'W' OR CMD_STATUS = 'P' ) " & _
                " AND COMPCODE={1} " & _
                " AND GATEWAYTYPE = 1 " & _
                " ORDER BY CMD_STATUS,SEQNO ", Me.ProcessSeqNo, Me.CompCode)
            Else
                Me.QuerySQL = String.Format("SELECT * FROM SEND_NAGRA " & _
                " WHERE  SEQNO > {0} " & _
                " AND COMPCODE={1} " & _
                " AND GATEWAYTYPE = 1 " & _
                " ORDER BY CMD_STATUS,SEQNO ", Me.ProcessSeqNo, Me.CompCode)
            End If


            Me.QuerySQL = String.Format("SELECT * FROM ({0}) WHERE ROWNUM<={1} AND GATEWAYTYPE  = 1 " & _
                                       "ORDER BY SEQNO ",
                                       Me.QuerySQL, ProcCmdCount)

            If WaitUpdCmd IsNot Nothing AndAlso WaitUpdCmd.Count > 0 Then

                Dim Data As String = Nothing
                For Each Str As String In WaitUpdCmd.Keys
                    If String.IsNullOrEmpty(Data) Then
                        Data = Str
                    Else
                        Data = Data & "," & Str
                    End If
                Next
                Me.QuerySQL = String.Format("SELECT * FROM ( {0} )  UNION ALL " & _
                                            " SELECT * FROM SEND_NAGRA WHERE SEQNO IN ({1})",
                                          Me.QuerySQL, Data)
                Me.QuerySQL = String.Format("SELECT DISTINCT * FROM ({0}) ORDER BY SEQNO ", Me.QuerySQL)
            End If


            Using cmd As OracleCommand = cn.CreateCommand
                cmd.CommandText = Me.QuerySQL
                If cn.State <> ConnectionState.Open Then
                    cn.Open()
                End If
                Using rdr As OracleDataReader = cmd.ExecuteReader(CommandBehavior.CloseConnection)
                    If Not rdr.HasRows Then
                        Using cmd2 As OracleCommand = cn.CreateCommand
                            cmd.CommandText = String.Format("SELECT NVL(MAX(SEQNO),0) MAX " & _
                                " FROM SEND_NAGRA " & _
                                "WHERE COMPCODE={0} AND GATEWAYTYPE = 1", Me.CompCode)
                            Me.ProcessSeqNo = cmd.ExecuteScalar
                            If ProcessSeqNo >= MaxValue Then
                                Me.ProcessSeqNo = 0
                            End If
                        End Using
                    End If

                    While rdr.Read
                        Interlocked.Increment(Me.UpdateUICount)
                        If (Me.ProcessSeqNo < UInteger.Parse(rdr.Item("SEQNO"))) OrElse
                            (Me.ProcessSeqNo = UInteger.MaxValue) Then
                            Me.ProcessSeqNo = UInteger.Parse(rdr.Item("SEQNO"))
                        End If

                        Dim aryRw As New Dictionary(Of String, Object)
                        For i As Int32 = 0 To rdr.FieldCount - 1
                            aryRw.Add(rdr.GetName(i).ToUpper, rdr.Item(i))
                        Next
                        Select Case rdr.Item("CMD_STATUS").ToString.ToUpper
                            Case "C", "T", "E"
                                If WaitUpdCmd.ContainsKey(rdr.Item("SEQNO").ToString.ToUpper) Then
                                    WaitUpdCmd.Remove(rdr.Item("SEQNO").ToString.ToUpper)
                                End If
                            Case Else
                                If Not WaitUpdCmd.ContainsKey(rdr.Item("SEQNO").ToString.ToUpper) Then
                                    WaitUpdCmd.Add(rdr.Item("SEQNO").ToString.ToUpper,
                                                   rdr.Item("SEQNO").ToString.ToUpper)


                                End If
                        End Select
                        Common.UpdateHighCmd(aryRw, Me)
                    End While
                End Using
            End Using
            Common.UpdateSODatabase(Me, Common.SOStatus.Yes, guidKey)
        Catch ex As Exception            
            If (cn IsNot Nothing) AndAlso (cn.State = ConnectionState.Open) Then
                Try
                    cn.Close()
                    cn.Open()
                Catch ex2 As Exception
                    Common.UpdateShowSO(Me, Common.SOStatus.NO)
                    Common.UpdateSODatabase(Me, Common.SOStatus.NO, guidKey)
                End Try

            Else
                Common.UpdateShowSO(Me, Common.SOStatus.NO)
                Common.UpdateSODatabase(Me, Common.SOStatus.NO, guidKey)
            End If
        Finally


            If cn IsNot Nothing Then
                cn.Close()
            End If
            mtxWait.ReleaseMutex()
        End Try
    End Sub

    Public Sub Run()
        Me.IsStop = False
        If evn IsNot Nothing Then
            evn.Dispose()
            evn = Nothing
        End If

        evn = New AutoResetEvent(False)
        ti = ThreadPool.RegisterWaitForSingleObject(evn, _
                                                           New WaitOrTimerCallback(AddressOf WaitProc), _
                                                           Nothing, TimeSpan.FromSeconds(ReadTime), False)
        evn.Set()
    End Sub

#Region "IDisposable Support"
    Private disposedValue As Boolean ' 偵測多餘的呼叫

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                If cn IsNot Nothing Then
                    cn.Close()
                    cn.Dispose()
                    cn = Nothing
                End If
                If evn IsNot Nothing Then
                    evn.Close()
                    evn.Dispose()
                    evn = Nothing
                End If
                If mtxWait IsNot Nothing Then
                    mtxWait.Close()
                    mtxWait.Dispose()
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
