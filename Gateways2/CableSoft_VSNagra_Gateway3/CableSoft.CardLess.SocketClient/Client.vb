Imports System.Net
Imports System.Text
Imports System.Threading
Imports System.Net.Sockets
Imports CableSoft.Nagra.BuilderLowerCmd
Imports CableSoft.NSTV.Log
Imports CableSoft.NSTV.BuilderHighCmd
Public Class Client
    'Inherits Socket
    Implements IDisposable
    Private Shared _ProcessSendCount As Integer = 0
    'Private Shared sendDone As New AutoResetEvent(False)
    'Private Shared recvDone As New AutoResetEvent(False)
    Private _chkStateTime As Integer = 30
    Private tmrNagraStatus As Threading.Timer = Nothing
    Private LowerCmd As CableSoft.Nagra.BuilderLowerCmd.LowCmd = Nothing
    'Private HighCmd As HighCmd = Nothing
    Private _Host As String
    Private _Port As Integer
    Private Shared IsChkStatus As Boolean = False
    Private mtxSendData As New Mutex()
    Private mtxChkStatus As New Mutex()
    'Private _AddressFamily As AddressFamily
    'Private _SocketType As SocketType
    'Private _ProtocolType As ProtocolType
    Private FirstTime As Boolean = False
    'Private tbSetNstv As DataTable = Nothing
    'Private tbSetHighCmd As DataTable = Nothing
    'Private tbSetLowCmd As DataTable = Nothing
    
    Private sckClient As Socket = Nothing
    Public Property SourceId As String
    Public Property DestId As String
    Public Property MoppId As String
    Public Property SendTimeout As Integer = 5
    Public Property ReceiveTimeout As Integer = 5
    Public Property DisconnectRetryTime As Integer = 3  '斷線後每幾秒再重連一次
    Public Property RetryProcCount As Integer           '傳送失敗可嘗試幾次
    'Public Property SendDelayTime As Double            '每傳一筆停幾秒
    Public Property ForceStop As Boolean = False
    Public Property IsEncrypt As Boolean = True
    Public Property tbError As DataTable = Nothing
    Public Property tbNagra As DataTable = Nothing
    Private ConnectErrCount As Int32 = 0
    Public ReadOnly Property Host
        Get
            Return _Host
        End Get
    End Property
    Public Enum CableSoftError
        SocketSendOrRecvError = -9999
        GetLowOrHighCmdError = -9998
        NstvFormatError = -9997
        SocketAllDisconnect = -9996
        GetZeroLowCmd = -9995
    End Enum

    Public ReadOnly Property Connected As Boolean
        Get
            If sckClient IsNot Nothing Then
                Return sckClient.Connected
            Else
                Return False
            End If
        End Get
    End Property



    Public WriteOnly Property CheckStateTime As Int32
        Set(value As Int32)
            _chkStateTime = value

        End Set

    End Property
    Public Function SendHighCmd(ByVal rw As Dictionary(Of String, Object)) As Dictionary(Of String, Object)
        Dim SendByte As New ThreadLocal(Of Byte())
        Dim RecvByte As New ThreadLocal(Of Byte())
        Dim result As New ThreadLocal(Of Dictionary(Of String, Object))
        Dim commandType As New ThreadLocal(Of LowCmd.CmdType)
        Dim NagraResult As New ThreadLocal(Of NagraCommandResult)
        Do
            If Not IsChkStatus Then
                Exit Do
            Else
                Thread.Sleep(100)
            End If
        Loop
        If Not Monitor.TryEnter(sckClient, 10000) Then
            Return rw
        End If
        IsChkStatus = True
        tmrNagraStatus.Change(Timeout.Infinite, Timeout.Infinite)
        Array.Resize(RecvByte.Value, 1024)
        Try
            If (sckClient Is Nothing) OrElse (Not sckClient.Connected) Then
                Return rw
            End If
            result.Value = rw
            result.Value.Item("UPDTIME") = DateTime.Now
            'result.Value.Item("UPDTIME") = DateTime.Now
            'Add with Send and Receive information to nagracmdlog table By Kin 2015/09/11
            'aryDataRow.Add("RetryCount", 0)
            'aryDataRow.Add("CATRANNUM", "")
            'aryDataRow.Add("GWTRANNUM", "")
            'aryDataRow.Add("SendCmdText", "")
            'aryDataRow.Add("RecvCmdText", "")
            'aryDataRow.Add("ACK", "")
            'aryDataRow.Add("SendTime", Date.Now)
            'aryDataRow.Add("RecvTime", Date.Now)
            Select Case rw.Item("HIGH_LEVEL_CMD_ID").ToString.ToUpper
                Case "B1".ToUpper
                    SendByte.Value = LowerCmd.GetLowCmd(LowCmd.CmdType.Cmd0851, rw)
                    commandType.Value = LowCmd.CmdType.Cmd0851
                    result.Value.Item("LowCmd".ToUpper) = "851"
                Case "D1".ToUpper
                    SendByte.Value = LowerCmd.GetLowCmd(LowCmd.CmdType.Cmd0126, rw)
                    commandType.Value = LowCmd.CmdType.Cmd0126
                    result.Value.Item("LowCmd".ToUpper) = "126"
                Case Else
                    commandType.Value = LowCmd.CmdType.OtherCMD
            End Select
            result.Value("SendCmdText") = System.Text.Encoding.ASCII.GetString(SendByte.Value).Substring(2)
            result.Value("SendTime") = Date.Now
            If System.Text.Encoding.ASCII.GetString(SendByte.Value).Substring(2).Length > 9 Then
                result.Value("GWTRANNUM") = System.Text.Encoding.ASCII.GetString(SendByte.Value).Substring(2).Substring(0, 9)
            End If

            'Try

            sckClient.Send(SendByte.Value)            
            Thread.Sleep(100)
            sckClient.Receive(RecvByte.Value)
            'Catch ex As Exception
            ' CableSoft.NSTV.Log.NstvLog.WriteErrorLog(ex, Nothing)
            'Return result.Value
            'End Try

            NagraResult.Value = LowerCmd.AnalyseSendLowCmd(RecvByte.Value, commandType.Value)
            result.Value.Item("CATRANNUM") = NagraResult.Value.GwtRanNum
            result.Value.Item("ACK") = NagraResult.Value.Ack
            result.Value.Item("RecvCmdText") = NagraResult.Value.RecvString
            Select Case commandType.Value
                Case LowCmd.CmdType.Cmd0126
                    Select Case NagraResult.Value.Status
                        Case LowCmd.AckCode.Ack
                            result.Value.Item("CMD_STATUS") = "C"
                            result.Value.Item("RETURNDATA") = NagraResult.Value.VUA
                            result.Value.Item("ERR_CODE") = String.Empty
                            result.Value.Item("ERR_MSG") = String.Empty
                        Case LowCmd.AckCode.Nack, LowCmd.AckCode.Other
                            result.Value.Item("CMD_STATUS") = "E"
                            result.Value.Item("RETURNDATA") = String.Empty
                            result.Value.Item("ERR_CODE") = NagraResult.Value.ErrorCode
                            result.Value.Item("ERR_MSG") = NagraResult.Value.ErrorName
                    End Select
                Case Else
                    result.Value.Item("CMD_STATUS") = "E"
                    result.Value.Item("RETURNDATA") = String.Empty
                    result.Value.Item("ERR_CODE") = LowCmd.CableSoftCustomerError.CableSoft_UnKnowCommand
                    result.Value.Item("ERR_MSG") = [Enum].GetName(GetType(LowCmd.CableSoftCustomerError),
                                                                  LowCmd.CableSoftCustomerError.CableSoft_UnKnowCommand)
            End Select
            result.Value.Item("RecvTime") = Date.Now
            Return result.Value
        Catch exSocket As SocketException
            CableSoft.NSTV.Log.NstvLog.WriteErrorLog(exSocket, Nothing)
            rw.Item("ERR_CODE") = exSocket.SocketErrorCode
            rw.Item("ERR_MSG") = exSocket.Message
            Try
                If sckClient.Connected Then
                    sckClient.Disconnect(False)
                End If
            Finally
                sckClient.Close()
                sckClient.Dispose()
            End Try
            Return rw
        Catch ex As Exception
            CableSoft.NSTV.Log.NstvLog.WriteErrorLog(ex, Nothing)
            rw.Item("CMD_STATUS") = "E"
            rw.Item("ERR_CODE") = LowCmd.CableSoftCustomerError.CableSoft_OtherError
            rw.Item("ERR_MSG") = ex.Message
            Return result.Value
        Finally
            SendByte.Dispose()
            result.Dispose()
            RecvByte.Dispose()
            commandType.Dispose()
            IsChkStatus = False
            tmrNagraStatus.Change(_chkStateTime * 1000, _chkStateTime * 1000)
            Monitor.Exit(sckClient)
        End Try
    End Function



    
    Private Sub WriteSendLog(ByVal CmdSeqNo As String, ByVal Data As String)

        Try

        Catch ex As Exception
        Finally

        End Try
    End Sub


    
    Public Sub RetryConnect()
        Try
            If tmrNagraStatus IsNot Nothing Then
                tmrNagraStatus.Change(Me.DisconnectRetryTime * 1000, Me.DisconnectRetryTime * 1000)
            End If
        Finally

        End Try
    End Sub
   
    Public Sub New(
                  ByVal SourceId As String, DestId As String, ByVal MoppId As String)
        Try
            sckClient = New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)
            
            FirstTime = True
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Public Sub New(ByVal tbNagra As DataTable, ByVal tbError As DataTable,
                   ByVal SourceId As String, ByVal DestId As String,
                   ByVal MoppId As String)
        Try
            sckClient = New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)
            sckClient.SendTimeout = Integer.Parse(tbNagra.Rows(0).Item("SendTimeOut"))
            sckClient.ReceiveTimeout = Integer.Parse(tbNagra.Rows(0).Item("ReceiveTimeout"))
            Me._chkStateTime = tbNagra.Rows(0).Item("CheckTime")
            'Me.SendDelayTime = tbSetNstv.Rows(0).Item("SendDelayTime")
            Me.RetryProcCount = tbNagra.Rows(0).Item("RetryProcCount")
            Me.DisconnectRetryTime = tbNagra.Rows(0).Item("DisconnectRetryTime")
            Me._Host = tbNagra.Rows(0).Item("IPAddress").ToString
            Me._Port = Integer.Parse(tbNagra.Rows(0).Item("Port"))
            Me.SourceId = SourceId
            Me.DestId = DestId
            Me.MoppId = MoppId
            FirstTime = True
        Catch ex As Exception
            Throw
        End Try

    End Sub




    Public Sub Connect(ByVal host As String, ByVal Port As Integer)
        Try
            sckClient.Connect(host, Port)
        Catch ex As Exception
            Throw
        End Try
    End Sub
    Public Sub ConnectNagra()
        Try
            ConnectErrCount = 0
            ConnectNagra(_Host, _Port)
        Catch ex As Exception
            Throw
        End Try
    End Sub
    Public Sub ConnectNagra(ByVal host As String, ByVal Port As Int32)
        Try
            _Host = host
            _Port = Port
            ConnectErrCount = 0
            If LowerCmd Is Nothing Then
                LowerCmd = New CableSoft.Nagra.BuilderLowerCmd.LowCmd(SourceId, DestId, MoppId, tbError)
            End If

         
            If tmrNagraStatus Is Nothing Then
                tmrNagraStatus = New Threading.Timer(AddressOf chkStatus, Nothing, 0, _chkStateTime * 1000)
            Else
                tmrNagraStatus.Change(0, _chkStateTime * 1000)
            End If
        Catch ex As Exception
            Throw ex

        End Try


    End Sub
    Public Overloads Function SendSynchronize(ByVal SendData() As Byte) As Byte()
        mtxSendData.WaitOne()
        If tmrNagraStatus IsNot Nothing Then
            tmrNagraStatus.Change(Timeout.Infinite, Timeout.Infinite)
        End If
        Dim ByteRecv As New ThreadLocal(Of Byte())
        Dim IsSendOK As Boolean = False
        Array.Resize(ByteRecv.Value, 2048)
      
        Try
            If IsChkStatus Then
                Do
                    Thread.Sleep(100)
                    If Not IsChkStatus Then
                        Exit Do
                    End If
                Loop
            End If
            sckClient.Send(SendData)
            IsSendOK = True
            sckClient.Receive(ByteRecv.Value)

            Return ByteRecv.Value
        Catch ex As Exception

            Try
                If sckClient.Connected Then
                    sckClient.Disconnect(False)
                End If
            Finally
                sckClient.Close()
                sckClient.Dispose()
                FirstTime = True
                tmrNagraStatus.Change(0, Me.DisconnectRetryTime * 1000)
            End Try

            Throw ex
        Finally
            ByteRecv.Dispose()
            If tmrNagraStatus IsNot Nothing Then
                tmrNagraStatus.Change(_chkStateTime * 1000, _chkStateTime * 1000)
            End If

            mtxSendData.ReleaseMutex()
        End Try


    End Function
    

    Private Sub chkStatus(ByVal status As Object)
        If IsChkStatus Then
            Exit Sub
        End If
        mtxChkStatus.WaitOne()
        Dim recvData As New ThreadLocal(Of Byte())
        If ForceStop Then
            IsChkStatus = False
            mtxChkStatus.ReleaseMutex()
            Exit Sub
        End If
        IsChkStatus = True
        If tmrNagraStatus IsNot Nothing Then
            tmrNagraStatus.Change(Timeout.Infinite, Timeout.Infinite)
        End If
        Dim RecvStatus As New ThreadLocal(Of CheckStatusResult)
        Dim RecvNagraResult As New ThreadLocal(Of NagraCommandResult)
        Try
            If (sckClient Is Nothing) OrElse (Not sckClient.Connected) Then               
                FirstTime = True
                sckClient = New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)
                sckClient.SendTimeout = Me.SendTimeout * 1000
                sckClient.ReceiveTimeout = Me.ReceiveTimeout * 1000
                sckClient.Connect(_Host, _Port)
                Array.Resize(recvData.Value, 512)
                sckClient.Send(LowerCmd.GetLowCmd(LowCmd.CmdType.ConfirmConnect, Nothing))
                Thread.Sleep(300)
                sckClient.Receive(recvData.Value)
                RecvStatus.Value = LowerCmd.ChkeckState(recvData.Value)
                If (RecvStatus.Value.Answer_Code <> 0) OrElse (RecvStatus.Value.Success <> 6) Then
                    CableSoft.NSTV.Log.NstvLog.WriteErrorLog(New Exception("主機連接失敗"),
                                        String.Format("SourceId：{0}" & Environment.NewLine & _
                                                      "Connect Status：{1}" & Environment.NewLine & _
                                                      "Answer_Code：{2}", SourceId, RecvStatus.Value.Success,
                                                      RecvStatus.Value.Answer_Code))
                    Try
                        FirstTime = True
                        If sckClient.Connected Then
                            sckClient.Disconnect(False)
                        End If
                    Finally
                        sckClient.Close()
                        sckClient.Dispose()
                        sckClient = Nothing
                    End Try
                End If
            End If

            If (sckClient IsNot Nothing) AndAlso (sckClient.Connected) Then

                Array.Resize(recvData.Value, 512)
                sckClient.Send(LowerCmd.GetLowCmd(LowCmd.CmdType.Cmd1002, Nothing))
                Thread.Sleep(200)
                sckClient.Receive(recvData.Value)
            End If
            RecvNagraResult.Value = LowerCmd.AnalyseSendLowCmd(recvData.Value, LowCmd.CmdType.Cmd1002)
            If RecvNagraResult.Value.Status = LowCmd.AckCode.Ack Then
                tmrNagraStatus.Change(_chkStateTime * 1000, _chkStateTime * 1000)
                Interlocked.Exchange(ConnectErrCount, 0)
                FirstTime = False
            Else
                CableSoft.NSTV.Log.NstvLog.WriteErrorLog(New Exception("主機連接失敗"),
                                         String.Format("SourceId：{0}" & Environment.NewLine & _
                                                       "ErrorCode：{1}" & Environment.NewLine & _
                                                       "ErrorName：{2}", SourceId, RecvNagraResult.Value.ErrorCode,
                                                       RecvNagraResult.Value.ErrorName))
                Try
                    FirstTime = True
                    If sckClient.Connected Then
                        sckClient.Disconnect(False)
                    End If
                Finally
                    sckClient.Close()
                    sckClient.Dispose()

                End Try

            End If



        Catch ex As Exception
            If ConnectErrCount = 300 Then
                Interlocked.Increment(ConnectErrCount)
                CableSoft.NSTV.Log.NstvLog.WriteErrorLog(ex,
                                                         "TCP已連接多次失敗，系統不再寫入錯誤，請檢查網路連線")
            Else
                If ConnectErrCount <= 300 Then
                    Interlocked.Increment(ConnectErrCount)
                    CableSoft.NSTV.Log.NstvLog.WriteErrorLog(ex, Nothing)
                End If

            End If

            Try
                FirstTime = True
                If sckClient.Connected Then
                    sckClient.Disconnect(False)
                End If
            Finally
                sckClient.Close()
                sckClient.Dispose()

            End Try
        Finally
            RecvStatus.Dispose()
            recvData.Dispose()
            RecvNagraResult.Dispose()
            If sckClient IsNot Nothing AndAlso sckClient.Connected Then
                tmrNagraStatus.Change(Me._chkStateTime * 1000, Me._chkStateTime * 1000)
            Else
                tmrNagraStatus.Change(Me.DisconnectRetryTime * 1000, Me.DisconnectRetryTime * 1000)
            End If
            IsChkStatus = False
            mtxChkStatus.ReleaseMutex()
        End Try
    End Sub





#Region "IDisposable Support"
    Private disposedValue As Boolean ' 偵測多餘的呼叫

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                If tmrNagraStatus IsNot Nothing Then
                    tmrNagraStatus.Change(Timeout.Infinite, Timeout.Infinite)
                    tmrNagraStatus.Dispose()
                    tmrNagraStatus = Nothing
                End If
                If sckClient IsNot Nothing Then
                    Try
                        If sckClient.Connected Then
                            sckClient.Disconnect(False)
                        End If

                    Finally
                        sckClient.Close()
                        sckClient.Dispose()
                        sckClient = Nothing
                    End Try
                End If
                If mtxChkStatus IsNot Nothing Then
                    mtxChkStatus.Dispose()
                    mtxChkStatus = Nothing
                End If
                If mtxSendData IsNot Nothing Then
                    mtxSendData.Close()
                    mtxSendData.Dispose()
                    mtxSendData = Nothing
                End If
                If mtxChkStatus IsNot Nothing Then
                    mtxChkStatus.Close()
                    mtxSendData.Dispose()
                    mtxSendData = Nothing
                End If
                'If tbSetHighCmd IsNot Nothing Then
                '    tbSetHighCmd.Dispose()
                '    tbSetHighCmd = Nothing
                'End If
                'If tbSetLowCmd IsNot Nothing Then
                '    tbSetLowCmd.Dispose()
                '    tbSetLowCmd = Nothing
                'End If
                'If tbSetNstv IsNot Nothing Then
                '    tbSetNstv.Dispose()
                '    tbSetNstv = Nothing
                'End If
                If tbError IsNot Nothing Then
                    tbError.Dispose()
                End If
                'If HighCmd IsNot Nothing Then
                '    HighCmd.Dispose()
                '    HighCmd = Nothing
                'End If
                If LowerCmd IsNot Nothing Then
                    LowerCmd.Dispose()
                    LowerCmd = Nothing
                End If
                ' TODO: 處置 Managed 狀態 (Managed 物件)。
            End If

            ' TODO: 釋放 Unmanaged 資源 (Unmanaged 物件) 並覆寫下面的 Finalize()。
            ' TODO: 將大型欄位設定為 null。
        End If
        Me.disposedValue = True
    End Sub



    ' 由 Visual Basic 新增此程式碼以正確實作可處置的模式。
    Public Sub Dispose() Implements IDisposable.Dispose
        ' 請勿變更此程式碼。在以上的 Dispose 置入清除程式碼 (ByVal 視為布林值處置)。
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class
Friend Class StateObject
    ' Client socket.
    Public workSocket As Socket = Nothing

    Public BufferSize As Integer = 4095
    ' Receive buffer.
    Public buffer(BufferSize) As Byte


End Class 'StateObject
