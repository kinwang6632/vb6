Imports System.Net
Imports System.Text
Imports System.Threading
Imports System.Net.Sockets
Imports CableSoft.CardLess.BuilderLowerCmd
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
    Private LowerCmd As CableSoft.CardLess.BuilderLowerCmd.LowCmd = Nothing
    'Private HighCmd As HighCmd = Nothing
    Private _Host As String
    Private _Port As Integer
    Private Shared IsChkStatus As Boolean = False
    'Private mtxSendData As New Mutex()
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
    'Private Sub ReadxmlFileData(ByVal xmlFile As String, ByVal xmlIsEncrypt As Boolean)
    '    Try
    '        Using xmlIO As New CableSoft.CardLess.XMLFileIO.XmlFileIO(xmlFile)
    '            tbSetHighCmd = xmlIO.ReadHightCmd(xmlIsEncrypt)
    '            tbSetLowCmd = xmlIO.ReadLowCmd(xmlIsEncrypt)
    '            'tbSMS = xmlIO.ReadSMS(xmlIsEncrypt)
    '            tbSetNstv = xmlIO.ReadNagra(xmlIsEncrypt)
    '            tbError = xmlIO.ReadErrorList(xmlIsEncrypt)
    '        End Using
    '        Me.SMS_ID = tbSetNstv.Rows(0).Item("SMS_ID")
    '        Me.OPE_ID = tbSetNstv.Rows(0).Item("OPE_ID")
    '        Me.Proto_Ver = tbSetNstv.Rows(0).Item("ProtoVer")
    '        Me.Crypt_Ver = tbSetNstv.Rows(0).Item("CryptVer")
    '        Me.Key_Type = tbSetNstv.Rows(0).Item("KeyType")
    '        Me._chkStateTime = tbSetNstv.Rows(0).Item("CheckTime")
    '        Me.SendDelayTime = tbSetNstv.Rows(0).Item("SendDelayTime")
    '        Me.RetryProcCount = tbSetNstv.Rows(0).Item("RetryProcCount")
    '        Me.DisconnectRetryCount = tbSetNstv.Rows(0).Item("DisconnectRetry")
    '    Catch ex As Exception
    '        Throw
    '    End Try

    'End Sub



    'Public Function SendHighCmd(ByVal rw As Dictionary(Of String, Object)) As Dictionary(Of String, Object)
    '    Dim SendLowCmdList As New ThreadLocal(Of List(Of SendLowCmd))
    '    Dim IsSendData As New ThreadLocal(Of Boolean)
    '    Dim recvData As New ThreadLocal(Of Byte())
    '    Dim response As New ThreadLocal(Of ResponseLowCmd)
    '    Dim rwReturn As New ThreadLocal(Of Dictionary(Of String, Object))
    '    SendLowCmdList.Value = New List(Of SendLowCmd)
    '    Array.Resize(recvData.Value, 1024)
    '    'Monitor.Enter(sckClient)        
    '    Try
    '        rwReturn.Value = rw
    '        rwReturn.Value.Item("ReSentTimes".ToUpper) = 0
    '        'If Me.ForceStop Then
    '        '    Return rwReturn.Value
    '        'End If
    '        If IsChkStatus Then
    '            Do
    '                If (Not IsChkStatus) Then
    '                    Exit Do
    '                End If
    '                Thread.Sleep(100)
    '            Loop
    '        End If

    '        Try
    '            If tbSetHighCmd Is Nothing OrElse tbSetHighCmd.Rows.Count = 0 Then
    '                Throw (New Exception("無任何高階命令可分析！"))
    '            End If
    '            If tbSetLowCmd Is Nothing OrElse tbSetLowCmd.Rows.Count = 0 Then
    '                Throw (New Exception("無任何低階命令可分析！"))
    '            End If
    '            SendLowCmdList.Value = HighCmd.GetSendDataList(rwReturn.Value)
    '            SyncLock sckClient
    '                'SendLowCmdList.Value = HighCmd.GetSendDataList(rwReturn.Value)
    '                If SendLowCmdList.Value.Count > 0 Then
    '                    For i As Int32 = 0 To SendLowCmdList.Value.Count - 1
    '                        tmrNagraStatus.Change(Timeout.Infinite, Timeout.Infinite)
    '                        Try
    '                            If sckClient.Connected Then
    '                                If i = 0 Then
    '                                    rwReturn.Value.Item("SendLog".ToUpper) = SendLowCmdList.Value(i).SendDataString
    '                                Else
    '                                    rwReturn.Value.Item("SendLog".ToUpper) = rwReturn.Value.Item("SendLog".ToUpper) & "," & _
    '                                                        SendLowCmdList.Value(i).SendDataString
    '                                End If

    '                                sckClient.Send(SendLowCmdList.Value.Item(i).SendData)
    '                                IsSendData.Value = True
    '                                sckClient.Receive(recvData.Value)
    '                                response.Value = LowerCmd.AnalyseLowCmd(recvData.Value)
    '                                If Not response.Value.Success Then
    '                                    rwReturn.Value.Item("CMD_STATUS".ToUpper) = "E"
    '                                    rwReturn.Value.Item("ERR_CODE".ToUpper) = IIf(String.IsNullOrEmpty(response.Value.ErrorCode.ToString),
    '                                                                                 CableSoftError.SocketSendOrRecvError, response.Value.ErrorCode)
    '                                    rwReturn.Value.Item("ERR_MSG".ToUpper) = IIf(String.IsNullOrEmpty(response.Value.ErrorName),
    '                                                                                [Enum].GetName(GetType(CableSoftError), CableSoftError.SocketSendOrRecvError),
    '                                                                                response.Value.ErrorName)
    '                                    Return rwReturn.Value
    '                                    Exit For
    '                                End If
    '                                If response.Value.ErrorCode = 0 Then
    '                                    rwReturn.Value.Item("CMD_STATUS".ToUpper) = "C"
    '                                Else
    '                                    rwReturn.Value.Item("CMD_STATUS".ToUpper) = "E"
    '                                    rwReturn.Value.Item("ERR_CODE".ToUpper) = "0x" & Right("0000" & Hex(response.Value.ErrorCode), 4)
    '                                    rwReturn.Value.Item("ERR_MSG".ToUpper) = response.Value.ErrorName
    '                                    Return rwReturn.Value
    '                                    Exit For
    '                                End If
    '                            Else
    '                                rwReturn.Value.Item("CMD_STATUS".ToUpper) = "E"
    '                                rwReturn.Value.Item("ERR_CODE".ToUpper) = Int32.Parse(CableSoftError.SocketSendOrRecvError)
    '                                rwReturn.Value.Item("ERR_MSG".ToUpper) = [Enum].GetName(GetType(CableSoftError), CableSoftError.SocketSendOrRecvError)
    '                                If Integer.Parse(rwReturn.Value.Item("ReSentTimes".ToUpper)) <= Me.RetryProcCount Then
    '                                    rwReturn.Value.Item("CMD_STATUS".ToUpper) = "P"
    '                                    rwReturn.Value.Item("ReSentTimes".ToUpper) = rwReturn.Value.Item("ReSentTimes".ToUpper) + 1
    '                                    rwReturn.Value.Item("ERR_CODE".ToUpper) = DBNull.Value
    '                                    rwReturn.Value.Item("ERR_MSG".ToUpper) = DBNull.Value
    '                                End If
    '                                Return rwReturn.Value
    '                                Exit For
    '                            End If
    '                        Catch exSock As SocketException
    '                            NstvLog.WriteErrorLog(exSock, String.Format("公司別 = {0}, SEQNO={1}, Host={2}",
    '                                                    rwReturn.Value.Item("CompCode".ToUpper),
    '                                                    rwReturn.Value.Item("SEQNO".ToUpper),
    '                                                    _Host))
    '                            Try
    '                                rwReturn.Value.Item("CMD_STATUS") = "E"
    '                                rwReturn.Value.Item("ERR_CODE".ToUpper) = Int32.Parse(exSock.SocketErrorCode)
    '                                rwReturn.Value.Item("ERR_MSG".ToUpper) = [Enum].GetName(GetType(SocketError),
    '                                                                                        exSock.SocketErrorCode)
    '                                If exSock.SocketErrorCode = SocketError.TimedOut Then
    '                                    rwReturn.Value.Item("CMD_STATUS") = "T"
    '                                End If
    '                                If Not IsSendData.Value Then
    '                                    If Integer.Parse(rwReturn.Value.Item("ReSentTimes".ToUpper)) <= Me.RetryProcCount Then
    '                                        rwReturn.Value.Item("CMD_STATUS".ToUpper) = "P"
    '                                        rwReturn.Value.Item("ReSentTimes".ToUpper) = rwReturn.Value.Item("ReSentTimes".ToUpper) + 1
    '                                        rwReturn.Value.Item("ERR_CODE".ToUpper) = DBNull.Value
    '                                        rwReturn.Value.Item("ERR_MSG".ToUpper) = DBNull.Value
    '                                    End If
    '                                End If
    '                                If sckClient.Connected Then
    '                                    sckClient.Disconnect(False)
    '                                End If
    '                                Return rwReturn.Value
    '                                Exit For
    '                            Finally
    '                                sckClient.Close()
    '                                sckClient.Dispose()
    '                            End Try

    '                        Catch ex As Exception
    '                            NstvLog.WriteErrorLog(ex, String.Format("公司別 = {0}, SEQNO={1}, Host={2}",
    '                                                    rwReturn.Value.Item("CompCode".ToUpper),
    '                                                    rwReturn.Value.Item("SEQNO".ToUpper),
    '                                                    _Host))

    '                            Try
    '                                rwReturn.Value.Item("CMD_STATUS".ToUpper) = "E"
    '                                rwReturn.Value.Item("ERR_CODE".ToUpper) = Int32.Parse(CableSoftError.SocketSendOrRecvError)
    '                                rwReturn.Value.Item("ERR_MSG".ToUpper) = [Enum].GetName(GetType(CableSoftError), CableSoftError.SocketSendOrRecvError)
    '                                If IsSendData.Value Then
    '                                    rwReturn.Value.Item("CMD_STATUS".ToUpper) = "E"
    '                                    rwReturn.Value.Item("ERR_CODE".ToUpper) = Int32.Parse(CableSoftError.SocketSendOrRecvError)
    '                                    rwReturn.Value.Item("ERR_MSG".ToUpper) = [Enum].GetName(GetType(CableSoftError), CableSoftError.SocketSendOrRecvError)
    '                                    Return rwReturn.Value
    '                                Else
    '                                    If Integer.Parse(rwReturn.Value.Item("ReSentTimes".ToUpper)) <= Me.RetryProcCount Then
    '                                        rwReturn.Value.Item("CMD_STATUS".ToUpper) = "P"
    '                                        rwReturn.Value.Item("ReSentTimes".ToUpper) = rwReturn.Value.Item("ReSentTimes".ToUpper) + 1
    '                                        rwReturn.Value.Item("ERR_CODE".ToUpper) = DBNull.Value
    '                                        rwReturn.Value.Item("ERR_MSG".ToUpper) = DBNull.Value
    '                                    End If
    '                                End If
    '                                If sckClient.Connected Then
    '                                    sckClient.Disconnect(False)
    '                                End If
    '                                Return rwReturn.Value
    '                                Exit For
    '                            Finally
    '                                'FirstTime = True
    '                                sckClient.Close()
    '                                sckClient.Dispose()
    '                            End Try
    '                        Finally
    '                            If String.IsNullOrEmpty(rwReturn.Value.Item("RecvLog".ToUpper)) Then
    '                                rwReturn.Value.Item("RecvLog".ToUpper) = IIf(String.IsNullOrEmpty(response.Value.RecvDataString),
    '                                                                             "0", response.Value.RecvDataString)
    '                            Else
    '                                rwReturn.Value.Item("RecvLog".ToUpper) = rwReturn.Value.Item("RecvLog".ToUpper) & "," & _
    '                                     IIf(String.IsNullOrEmpty(response.Value.RecvDataString),
    '                                                                             "0", response.Value.RecvDataString)
    '                            End If
    '                            tmrNagraStatus.Change(Me._chkStateTime * 1000, Me._chkStateTime * 1000)
    '                        End Try
    '                    Next
    '                Else
    '                    rwReturn.Value.Item("CMD_STATUS".ToUpper) = "E"
    '                    rwReturn.Value.Item("ERR_CODE".ToUpper) = Int32.Parse(CableSoftError.GetZeroLowCmd)
    '                    rwReturn.Value.Item("ERR_MSG".ToUpper) = [Enum].GetName(GetType(CableSoftError), CableSoftError.GetZeroLowCmd)
    '                    Return rwReturn.Value
    '                End If
    '            End SyncLock
    '        Catch ex As Exception
    '            NstvLog.WriteErrorLog(ex, String.Format("序號={0}, Host={1}",
    '                                                    rwReturn.Value.Item("SEQNO".ToUpper),
    '                                                    _Host))
    '            rwReturn.Value.Item("CMD_STATUS".ToUpper) = "E"
    '            rwReturn.Value.Item("ERR_CODE".ToUpper) = Int32.Parse(CableSoftError.GetLowOrHighCmdError)
    '            rwReturn.Value.Item("ERR_MSG".ToUpper) = [Enum].GetName(GetType(CableSoftError), CableSoftError.GetLowOrHighCmdError)
    '            Return rwReturn.Value
    '        Finally
    '            If tmrNagraStatus IsNot Nothing Then
    '                If (sckClient Is Nothing) OrElse (Not sckClient.Connected) Then
    '                    tmrNagraStatus.Change(Me.DisconnectRetryCount * 1000, Me.DisconnectRetryCount * 1000)
    '                Else

    '                End If
    '            End If
    '            SendLowCmdList.Dispose()
    '            IsSendData.Dispose()
    '            recvData.Dispose()
    '            response.Dispose()
    '        End Try
    '        Return rwReturn.Value
    '    Finally

    '        rwReturn.Dispose()
    '    End Try

    'End Function
    Private Sub WriteSendLog(ByVal CmdSeqNo As String, ByVal Data As String)

        Try

        Catch ex As Exception
        Finally

        End Try
    End Sub


    'Private Function GetLowCmd(ByVal rw As DataRow, ByVal MsgId As String) As SendLowCmd
    '    Dim result As New ThreadLocal(Of SendLowCmd)

    '    result.Value = Nothing
    '    Try
    '        Select Case Convert.ToInt32(MsgId, 16)
    '            Case 1
    '                result.Value = LowerCmd.Builder_SMS_CA_CREATE_SESSION_REQUEST()
    '            Case &H201
    '                result.Value = LowerCmd.Builder_SMS_CA_OPEN_ACCOUNT_REQUEST(
    '                    (rw.Item("ICC_NO") & New String("0", 16)).Substring(0, 16))
    '            Case &H206
    '                result.Value = LowerCmd.Builder_SMS_CA_REPAIR_REQUEST((rw.Item("ICC_NO") & New String("0", 16)).Substring(0, 16), rw.Item("STB_NO"))
    '            Case &H20D
    '                result.Value = LowerCmd.Builder_SMS_CA_ENTITLE_REQUEST((rw.Item("ICC_NO") & New String("0", 16)).Substring(0, 16), rw.Item("Notes"))
    '            Case &H202
    '                result.Value = LowerCmd.Builder_SMS_CA_STOP_ACCOUNT_REQUEST((rw.Item("ICC_NO") & New String("0", 16)).Substring(0, 16))
    '            Case &H203
    '                result.Value = LowerCmd.Builder_SMS_CA_SET_LOCK_REQUEST((rw.Item("ICC_NO") & New String("0", 16)).Substring(0, 16))
    '            Case &H204
    '                result.Value = LowerCmd.Builder_SMS_CA_SET_UNLOCK_REQUEST((rw.Item("ICC_NO") & New String("0", 16)).Substring(0, 16))
    '            Case &H402
    '                result.Value = LowerCmd.Builder_SMS_CA_SET_CHILD_REQUEST((rw.Item("ICC_NO") & New String("0", 16)).Substring(0, 16),
    '                                                                         rw.Item("STB_NO"), "3")
    '            Case &H403
    '                result.Value = LowerCmd.Builder_SMS_CA_CANCEL_CHILD_REQUEST((rw.Item("ICC_NO") & New String("0", 16)).Substring(0, 16))
    '            Case &H207
    '                result.Value = LowerCmd.Builder_SMS_CA_RESETCARDPIN_REQUEST((rw.Item("ICC_NO") & New String("0", 16)).Substring(0, 16))

    '            Case &H20C
    '                result.Value = LowerCmd.Builder_SMS_CA_SET_CHARACTER_REQUEST((rw.Item("ICC_NO") & New String("0", 16)).Substring(0, 16), "9", "AAAA")
    '            Case Else
    '                Throw New Exception(String.Format("此低階命令{0}尚未完成!", MsgId))
    '                '    MessageBox.Show("命令尚未完成!", "訊息", MessageBoxButtons.OK)
    '                '    Exit Function

    '        End Select
    '        Return result.Value
    '    Catch ex As Exception
    '        Throw New Exception(ex.ToString & Environment.NewLine & "MsgId=" & MsgId)

    '    Finally
    '        result.Dispose()
    '    End Try

    'End Function
    Public Sub RetryConnect()
        Try
            If tmrNagraStatus IsNot Nothing Then
                tmrNagraStatus.Change(Me.DisconnectRetryTime * 1000, Me.DisconnectRetryTime * 1000)
            End If
        Finally

        End Try
    End Sub
   
    Public Sub New(ByVal IPAddress As String, ByVal Port As Integer,
                  ByVal SourceId As String, DestId As String, ByVal MoppId As String)
        Try
            sckClient = New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)
            Me._Host = IPAddress
            Me._Port = Port
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
            Me.DisconnectRetryTime = tbNagra.Rows(0).Item("DisconnectRetryCount")
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
                LowerCmd = New CableSoft.CardLess.BuilderLowerCmd.LowCmd(SourceId, DestId, MoppId)
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
        'If String.IsNullOrEmpty(TalkKey) Then
        '    sckClient.Send(LowerCmd.Builder_SMS_CA_CREATE_SESSION_REQUEST().SendData)
        '    sckClient.Receive(ByteRecv.Value)
        '    If (LowerCmd.AnalyseLowCmd(ByteRecv.Value).ErrorCode = 0) AndAlso
        '        (LowerCmd.AnalyseLowCmd(ByteRecv.Value).Success) Then
        '        TalkKey = LowerCmd.AnalyseLowCmd(ByteRecv.Value).TalkKey
        '    Else
        '        Throw New Exception("會話命令失敗!")
        '    End If
        'End If
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
    'Public Overloads Function SendSynchronize(ByVal SendDataList As List(Of Byte())) As List(Of Byte())
    '    'Dim state As New StateObject
    '    'state.workSocket = Me
    '    'If _RecvDataList IsNot Nothing Then _RecvDataList.Clear()
    '    'Interlocked.Add(_ProcessSendCount, SendDataList.Count)
    '    '_ProcessOK = False

    '    'Try
    '    '    For Each byteData As Byte() In SendDataList
    '    '        Me.BeginSend(byteData, 0, byteData.Length, 0, New AsyncCallback(AddressOf SendCallback), Me)
    '    '        sendDone.WaitOne()
    '    '        '                Array.Clear(state.buffer, 0, state.BufferSize)
    '    '        Me.BeginReceive(state.buffer, 0, state.BufferSize, SocketFlags.None, AddressOf ReceiveCallback, state)
    '    '        recvDone.WaitOne()
    '    '    Next


    '    'Catch ex As Exception
    '    '    Throw ex

    '    'End Try
    '    '_ProcessOK = True
    '    'Return _RecvDataList

    'End Function

    Private Shared Sub SendCallback(ByVal ar As IAsyncResult)
        Try
            Dim client As Socket = CType(ar.AsyncState, Socket)
            Dim bytesSent As Integer = client.EndSend(ar)
            Interlocked.Decrement(_ProcessSendCount)
            sendDone.Set()

        Catch ex As Exception
            Throw
        End Try


        'sendDone.Set()
    End Sub 'SendCallback
    Private Sub ReceiveCallback(ByVal ar As IAsyncResult)
        Dim state As StateObject = CType(ar.AsyncState, StateObject)
        Dim client As Socket = state.workSocket
        Try
            ' Read data from the remote device.
            Dim bytesRead As Integer = client.EndReceive(ar)
            If bytesRead > 0 Then
                recvDone.Set()
            End If

        Catch ex As Exception
            Throw

        End Try

    End Sub 'ReceiveCallback

    Private Sub chkStatus(ByVal status As Object)
        If IsChkStatus Then
            Exit Sub
        End If
        mtxChkStatus.WaitOne()
        Dim recvData As New ThreadLocal(Of Byte())
        Array.Resize(recvData.Value, 1024)
        If ForceStop Then
            IsChkStatus = False
            mtxChkStatus.ReleaseMutex()
            Exit Sub
        End If
        IsChkStatus = True
        If tmrNagraStatus IsNot Nothing Then
            tmrNagraStatus.Change(Timeout.Infinite, Timeout.Infinite)
        End If

        Try
            If (sckClient Is Nothing) OrElse (Not sckClient.Connected) Then
                FirstTime = True
                sckClient = New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)
                sckClient.SendTimeout = Me.SendTimeout * 1000
                sckClient.ReceiveTimeout = Me.ReceiveTimeout * 1000
                sckClient.Connect(_Host, _Port)
            End If

            sckClient.Send(LowerCmd.BuilderCheckState)
            sckClient.Receive(recvData.Value)
            If Not LowerCmd.ChkeckState(recvData.Value) Then
                Try
                    FirstTime = True
                    If sckClient.Connected Then
                        sckClient.Disconnect(False)
                    End If
                Finally
                    sckClient.Close()
                    sckClient.Dispose()
                End Try
            Else
                If FirstTime Then
                    sckClient.Send(LowerCmd.Builder_SMS_CA_CREATE_SESSION_REQUEST().SendData)
                    sckClient.Receive(recvData.Value)
                    Dim result As CableSoft.CardLess.BuilderLowerCmd.ResponseLowCmd = LowerCmd.AnalyseLowCmd(recvData.Value)
                    If result.Success AndAlso result.ErrorCode = 0 Then
                        tmrNagraStatus.Change(_chkStateTime * 1000, _chkStateTime * 1000)
                        Interlocked.Exchange(ConnectErrCount, 0)

                        FirstTime = False
                    Else
                        Try
                            sckClient.Close()
                            sckClient.Dispose()
                        Finally
                            FirstTime = True
                        End Try
                    End If
                End If
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
                sckClient = Nothing               
            End Try
        Finally
            If sckClient IsNot Nothing AndAlso sckClient.Connected Then
                tmrNagraStatus.Change(Me._chkStateTime * 1000, Me._chkStateTime * 1000)
            Else
                tmrNagraStatus.Change(Me.DisconnectRetryTime * 1000, Me.DisconnectRetryTime * 1000)
            End If
            IsChkStatus = False
            recvData.Dispose()
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
                If tbSetHighCmd IsNot Nothing Then
                    tbSetHighCmd.Dispose()
                    tbSetHighCmd = Nothing
                End If
                If tbSetLowCmd IsNot Nothing Then
                    tbSetLowCmd.Dispose()
                    tbSetLowCmd = Nothing
                End If
                If tbSetNstv IsNot Nothing Then
                    tbSetNstv.Dispose()
                    tbSetNstv = Nothing
                End If
                If tbError IsNot Nothing Then
                    tbError.Dispose()
                End If
                If HighCmd IsNot Nothing Then
                    HighCmd.Dispose()
                    HighCmd = Nothing
                End If
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
