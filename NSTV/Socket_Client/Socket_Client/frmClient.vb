Imports System.Net
Imports System.Text
Imports System.Threading
Imports System.Net.Sockets
Public Class frmClient
    Private Client As CableSoft.CAS.SocketClient.Client = Nothing
    Private Event DataRecv(ByVal DataCount As Int32)
    Private Shared iCount As Integer = 0
    Private Shared blnStopSend As Boolean = False
    Private Shared evn As New System.Threading.AutoResetEvent(False)

    Private Shared mut As New Mutex()
    Private sendwait As New System.Threading.AutoResetEvent(False)
    Private recvwait As New System.Threading.AutoResetEvent(False)
    Private isRecv As Boolean = False
    Private Sub btnConnect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnConnect.Click
        If Client IsNot Nothing Then Exit Sub
        'Client = New Sockets.Socket(Sockets.AddressFamily.InterNetwork, Sockets.SocketType.Stream, Sockets.ProtocolType.Tcp)
        'Client.Connect("127.0.0.1", 8888)
        Client = New CableSoft.CAS.SocketClient.Client(Sockets.AddressFamily.InterNetwork, Sockets.SocketType.Stream, Sockets.ProtocolType.Tcp)
        Client.Connect("127.0.0.1", 7364)
        'Dim thr As New Thread(AddressOf Lookdata)
        'thr.IsBackground = True
        'thr.Start()

    End Sub

    Private Sub btnSendData_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSendData.Click
        ''blnStopSend = False
        ''Dim thr As New Thread(AddressOf SendData)
        ''thr.IsBackground = True
        ''thr.Start()
        Try
            '    Dim btyData(7) As Byte
            '    btyData(0) = &H0
            '    btyData(1) = &H4
            '    btyData(2) = &HFF
            '    btyData(3) = &HFF
            '    btyData(4) = &HFF
            '    btyData(5) = &HFF
            '    btyData(6) = &H0
            '    btyData(7) = &H0


            '    Dim bty As New List(Of Byte())
            '    bty.Add(btyData)
            '    'Client.SendDataList = bty
            '    Dim bytelst As List(Of Byte()) = Client.SendSynchronize(bty)

            '    For Each btyResult As Byte() In bytelst
            '        Debug.Print(Encoding.ASCII.GetString(btyResult.ToArray))
            '    Next
            Dim LowerCmd As New CableSoft.CAS.BuilderLowerCmd.LowCmd(1, 1, 0, 1, 1, "ABCDEFGH12345678")

            Dim lstByte As New List(Of Byte())
            Dim lstResult As New List(Of Byte())
            Dim byteData() As Byte
            'Dim byteData() As Byte = obj.BuilderCheckState
            'lstByte.Add(byteData)
            'lstResult = Client.SendSynchronize(lstByte)
            'For Each bty() As Byte In lstResult
            '    If obj.ChkeckState(bty) Then
            '        Debug.Print("yes")
            '    End If
            '    'Debug.Print(Encoding.ASCII.GetString(bty.ToArray))
            'Next
            lstByte.Clear()
            lstResult.Clear()

            'byteData = LowerCmd.BuilderLowCmd(1)
            lstByte.Add(byteData)
            lstResult = Client.SendSynchronize(lstByte)

            LowerCmd.AnalyseLowCmd(lstResult.Item(0))
        Catch exSocket As SocketException
            MessageBox.Show(exSocket.ToString)
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub
    Private Sub SendData()

        If Client IsNot Nothing Then
            While Not blnStopSend
                iCount += 1
                Dim bty() As Byte = Encoding.ASCII.GetBytes("This is a test" & iCount & "$")
                Client.Send(bty)
                Thread.Sleep(10)


            End While




        End If

    End Sub

    Private Sub Lookdata()
        While True
            If Client.Available > 0 Then
                isRecv = True

                RaiseEvent DataRecv(Client.Available)
            Else
                If Not isRecv Then
                    sendwait.Set()
                End If

                End If
        End While
    End Sub
    Private Sub btnRecv_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRecv.Click
        For i As Int32 = 0 To 10
            Debug.Print(i.ToString)
        Next


    End Sub

    Private Sub frmClient_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        AddHandler DataRecv, AddressOf RecvData

    End Sub
    Private Sub RecvData(ByVal DataCount As Int32)
        'Dim bty(1024) As Byte
        'Dim lstBty As New List(Of Byte)
        'Do While lstBty.Count < DataCount
        '    If DataCount - lstBty.Count < bty.Count Then
        '        Erase bty
        '        ReDim Preserve bty(DataCount - lstBty.Count - 1)
        '    End If
        '    Client.Receive(bty)
        '    lstBty.AddRange(bty)
        '    For i As Int32 = 0 To bty.Count - 1
        '        bty(i) = 0
        '    Next
        '    'Debug.Print(Encoding.ASCII.GetString(bty))
        'Loop
        'Debug.Print(Encoding.ASCII.GetString(lstBty.ToArray))
        'Erase bty
        ''Debug.Print(DataCount)
        'evn.Set()
        'evn.WaitOne()
        Dim state As New StateObject
        state.workSocket = Client


        ' Begin receiving the data from the remote device.
        'isRecv = False
        Client.BeginReceive(state.buffer, 0, StateObject.BufferSize, 0, New AsyncCallback(AddressOf ReceiveCallback), state)
        'recvwait.WaitOne()
    End Sub

    Private Sub btnStopSendData_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStopSendData.Click
        blnStopSend = True
    End Sub
    Private Sub BeginSendThread()
        Do While True
            If blnStopSend = True Then
                Exit Do
            End If
            If Not isRecv AndAlso Client.Available <= 0 Then

                Send(Client, "TEST11111$")
                Thread.Sleep(1)
                'sendwait.WaitOne()
            End If
            'evn.WaitOne()
        Loop
    End Sub
    Private Sub btnBeginSend_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBeginSend.Click
        Dim thr As New Thread(AddressOf BeginSendThread)
        thr.Start()
        
    End Sub
    Private Shared Sub SendCallback(ByVal ar As IAsyncResult)
        ' Retrieve the socket from the state object.

        Dim client As Socket = CType(ar.AsyncState, Socket)

        ' Complete sending the data to the remote device.
        Dim bytesSent As Integer = client.EndSend(ar)

        
        'Console.WriteLine("Sent {0} bytes to server.", bytesSent)

        ' Signal that all bytes have been sent.
        'sendDone.Set()
    End Sub 'SendCallback
    Private Sub Send(ByVal client As Socket, ByVal data As String)
        ' Convert the string data to byte data using ASCII encoding.
        Dim byteData As Byte() = Encoding.ASCII.GetBytes(data)
        client.BeginSend(byteData, 0, byteData.Length, 0, New AsyncCallback(AddressOf SendCallback), client)
        ' Begin sending the data to the remote device.
        sendwait.WaitOne()
    End Sub 'Send
    Private Sub Receive(ByVal client As Socket)

        ' Create the state object.
        Dim state As New StateObject
        state.workSocket = client

        ' Begin receiving the data from the remote device.

        client.BeginReceive(state.buffer, 0, StateObject.BufferSize, 0, New AsyncCallback(AddressOf ReceiveCallback), state)
        recvwait.WaitOne()
    End Sub 'Receive


    Private Sub ReceiveCallback(ByVal ar As IAsyncResult)

        ' Retrieve the state object and the client socket 
        ' from the asynchronous state object.


        Dim state As StateObject = CType(ar.AsyncState, StateObject)
        Dim client As Socket = state.workSocket
        'RaiseEvent DataRecv(client.EndReceive(ar))


        ' Read data from the remote device.
        Dim bytesRead As Integer = client.EndReceive(ar)

        If bytesRead > 0 Then
            ' There might be more data, so store the data received so far.
            state.sb.Append(Encoding.ASCII.GetString(state.buffer, 0, bytesRead))

            Debug.Print(state.sb.ToString)
            'client.BeginReceive(state.buffer, 0, StateObject.BufferSize, 0, New AsyncCallback(AddressOf ReceiveCallback), state)
            'recvwait.Set()
            isRecv = False
            ' Get the rest of the data.
            'client.BeginReceive(state.buffer, 0, StateObject.BufferSize, 0, New AsyncCallback(AddressOf ReceiveCallback), state)
            'evn.Set()

        Else
            If state.sb.Length > 0 Then
                Debug.Print(state.sb.ToString)
            End If
            ' All the data has arrived; put it in response.
            'If state.sb.Length > 1 Then
            '    response = state.sb.ToString()
            'End If
            '' Signal that all bytes have been received.
            'receiveDone.Set()
        End If
    End Sub 'ReceiveCallback
End Class
Public Class StateObject
    ' Client socket.
    Public workSocket As Socket = Nothing
    ' Size of receive buffer.
    Public Const BufferSize As Integer = 1024
    ' Receive buffer.
    Public buffer(BufferSize) As Byte
    ' Received data string.
    Public sb As New StringBuilder
End Class 'StateObject