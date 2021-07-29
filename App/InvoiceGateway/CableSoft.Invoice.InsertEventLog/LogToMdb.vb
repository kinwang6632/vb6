Imports System.IO

Public Class LogToMdb
    Implements IDisposable
    Private CompCode As String
    Private CompName As String    
    Private Const AccessFileName As String = "ProcessInvoiceLog.accdb"
    Private Const AccessProvider As String = "Microsoft.ACE.OLEDB.12.0"
    Private AccessConnectString As String = Nothing
    Public Property CanInsertAccessEvent As Boolean = False
    Public Property CanInsertSQLToLog As Boolean = False
    Private cn As OleDb.OleDbConnection = Nothing
    Dim updMutex As New Threading.Mutex()
    
    Public Sub New(ByVal CompCode As String,
                   ByVal CompName As String)
        Me.CompCode = CompCode
        Me.CompName = CompName
        Dim AccessFile As String =
                        String.Format("{0}\{1}",
                                                         Path.GetDirectoryName(System.AppDomain.CurrentDomain.BaseDirectory),
                                                        AccessFileName)
        If Not File.Exists(AccessFile) Then
            CanInsertAccessEvent = False
            Exit Sub
        End If

        If Not CanInsertAccessEvent Then
            CanInsertAccessEvent = LogToMdb.chkCanUseFile()
        End If

        If CanInsertAccessEvent Then
            Dim bduConnectString As New OleDb.OleDbConnectionStringBuilder
            bduConnectString.Provider = AccessProvider
            bduConnectString.DataSource = AccessFile
            bduConnectString.PersistSecurityInfo = False
            AccessConnectString = bduConnectString.ToString
            cn = New OleDb.OleDbConnection(AccessConnectString)
            cn.Open()
        End If
    End Sub
    Public Sub UpdateAllComplete()

        updMutex.WaitOne()
        Try
            If Not CanInsertAccessEvent Then
                Exit Sub
            End If
            Dim aSQL As String = String.Format("update ProcessCommand set C0401=0,  " & _
                                               "C0501=0,D0401=0,D0501=0," & _
                                               "C0701=0,X0401=0,B08=0 " & _
                                                 " where CompCode = '{0}'", CompCode)

            If cn.State = ConnectionState.Closed Then
                cn.Open()
            End If
            Using cmd As OleDb.OleDbCommand = cn.CreateCommand
                cmd.CommandText = aSQL
                cmd.ExecuteNonQuery()
                cmd.Dispose()
            End Using
        Catch ex As Exception
            CanInsertAccessEvent = False
            CableSoft.NSTV.Log.NstvLog.WriteErrorLog(ex, Nothing)
        Finally
            updMutex.ReleaseMutex()
        End Try

    End Sub
    Public Sub InsertSQLTextLog(ByVal InvRunType As CableSoft.B07.InvType.InvTypeEnum.InvFileType,
                                ByVal SQLText As String, ProcessTimeCount As Double, ProcessDate As DateTime)

        If (Not CanInsertAccessEvent) OrElse (Not CanInsertSQLToLog) Then
            Exit Sub
        End If
        updMutex.WaitOne()
        Try
            Dim aCountSQL As String = String.Format("select count(1) as cnt from ProcessSQL " & _
                                            " where compcode = {0}", CompCode)

            Dim aSQL As String = "Insert into ProcessSQL (CompCode,SQLText,ProcessType,ProcessTimeCount,ProcessDate) " & _
                            " Values (@CompCode,@SQLText,@ProcessType,@ProcessTimeCount,@ProcessDate)"
            If cn.State <> ConnectionState.Open Then
                cn.Open()
            End If
            Using cmd As OleDb.OleDbCommand = cn.CreateCommand
                cmd.CommandText = aCountSQL
                If Integer.Parse(cmd.ExecuteScalar) > 300 Then
                    Using cmd2 As OleDb.OleDbCommand = cn.CreateCommand
                        cmd2.CommandText = "delete from ProcessSQL"
                    End Using
                End If
            End Using
            Using cmd As OleDb.OleDbCommand = cn.CreateCommand
                cmd.Parameters.Add("CompCode", OleDb.OleDbType.Char).Value = CompCode
                cmd.Parameters.Add("SQLText", OleDb.OleDbType.Char).Value = SQLText
                cmd.Parameters.Add("ProcessType", OleDb.OleDbType.Char).Value =
                        [Enum].GetName(GetType(CableSoft.B07.InvType.InvTypeEnum.InvFileType), InvRunType)
                cmd.Parameters.Add("ProcessTimeCount", OleDb.OleDbType.Double).Value = ProcessTimeCount
                cmd.Parameters.Add("ProcessDate", OleDb.OleDbType.Date).Value = ProcessDate
                cmd.CommandText = aSQL
                cmd.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            CanInsertAccessEvent = False
            CableSoft.NSTV.Log.NstvLog.WriteErrorLog(ex, Nothing)
        Finally
            updMutex.ReleaseMutex()
        End Try
    End Sub
    Public Sub UpdateProcessStatus(ByVal InvRunType As CableSoft.B07.InvType.InvTypeEnum.InvFileType,
                                   ByVal Status As Status)

        If Not CanInsertAccessEvent Then
            Exit Sub
        End If
        updMutex.WaitOne()
        Try
          
            Dim aSQL As String = String.Format("update ProcessCommand set {0}={1}  " & _
                                               " where CompCode = '{2}'",
                                               [Enum].GetName(GetType(CableSoft.B07.InvType.InvTypeEnum.InvFileType), InvRunType),
                                               Integer.Parse(Status), CompCode)

            If cn.State = ConnectionState.Closed Then
                cn.Open()
            End If
            Using cmd As OleDb.OleDbCommand = cn.CreateCommand
                cmd.CommandText = aSQL
                cmd.ExecuteNonQuery()
                cmd.Dispose()
            End Using

        Catch ex As Exception
            CanInsertAccessEvent = False
            CableSoft.NSTV.Log.NstvLog.WriteErrorLog(ex, Nothing)
        Finally
            updMutex.ReleaseMutex()
        End Try
    End Sub
    Public Sub InsertSOInformat()
        If Not CanInsertAccessEvent Then
            Exit Sub
        End If
        updMutex.WaitOne()
        Try

            Dim aSQL As String = "Insert into ProcessCommand (CompCode,CompName) " & _
                            " Values (@CompCode,@CompName)"
            If cn.State = ConnectionState.Closed Then
                cn.Open()
            End If

            Using cmd As OleDb.OleDbCommand = cn.CreateCommand
                cmd.Parameters.Add("CompCode", OleDb.OleDbType.Char).Value = CompCode
                cmd.Parameters.Add("CompName", OleDb.OleDbType.Char).Value = CompName
                cmd.CommandText = aSQL
                cmd.ExecuteNonQuery()
            End Using

        Catch ex As Exception
            CanInsertAccessEvent = False
            CableSoft.NSTV.Log.NstvLog.WriteErrorLog(ex, Nothing)
        Finally
            updMutex.ReleaseMutex()
        End Try
    End Sub
    Public Shared Function chkCanUseFile() As Boolean
        Dim AccessFile As String =
                      String.Format("{0}\{1}",
                                                       Path.GetDirectoryName(System.AppDomain.CurrentDomain.BaseDirectory),
                                                      AccessFileName)
        Dim bduConnectString As New OleDb.OleDbConnectionStringBuilder
        bduConnectString.Provider = AccessProvider
        bduConnectString.DataSource = AccessFile
        bduConnectString.PersistSecurityInfo = False

        Try            
            Using cn As New OleDb.OleDbConnection(bduConnectString.ToString)
                cn.Open()
                Using cmd As OleDb.OleDbCommand = cn.CreateCommand
                    cmd.CommandText = "delete from ProcessCommand"
                    cmd.ExecuteNonQuery()
                    cmd.Dispose()
                End Using
                Using cmd As OleDb.OleDbCommand = cn.CreateCommand
                    cmd.CommandText = "delete from ProcessSQL "
                    cmd.ExecuteNonQuery()
                    cmd.Dispose()
                End Using

                cn.Close()
                cn.Dispose()

            End Using
        Catch ex As Exception
            Throw
        Finally

        End Try
        Return True
    End Function
#Region "IDisposable Support"
    Private disposedValue As Boolean ' 偵測多餘的呼叫

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: 處置 Managed 狀態 (Managed 物件)。
                If cn IsNot Nothing Then
                    cn.Close()
                    cn.Dispose()
                    cn = Nothing
                End If
                If updMutex IsNot Nothing Then
                    updMutex.Close()
                    updMutex.Dispose()
                    updMutex = Nothing
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
