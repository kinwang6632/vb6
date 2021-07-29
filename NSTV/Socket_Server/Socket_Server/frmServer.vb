Imports System.Net.Sockets
Imports System.Text
Imports System.Threading

Public Class frmServer

    Dim socket As System.Net.Sockets.Socket
    Private Sub frmServer_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        
    End Sub

    Private Sub frmServer_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        If socket IsNot Nothing Then
            socket.Close()

        End If
    End Sub
    Private Sub StarSocket()
        Dim requestCount As Integer
        Dim serverSocket As New TcpListener(8888)
        serverSocket.Start()
        Dim clientSocket As TcpClient
        clientSocket = serverSocket.AcceptTcpClient()
        'Dim obj As New SocketAsyncEventArgs()


        requestCount = 0
        Dim serverResponse As String = String.Empty
        While (True)
            Try
                requestCount = requestCount + 1
                Dim networkStream As NetworkStream = _
                        clientSocket.GetStream()
                Dim bytesFrom(10024) As Byte
                networkStream.Read(bytesFrom, 0, CInt(clientSocket.ReceiveBufferSize))
                Dim dataFromClient As String = _
                        System.Text.Encoding.ASCII.GetString(bytesFrom)
                dataFromClient = _
            dataFromClient.Substring(0, dataFromClient.IndexOf("$"))

                serverResponse = serverResponse & "---" & "Server response " + Convert.ToString(requestCount)
                If requestCount Mod 1 = 0 Then
                    Dim sendBytes As [Byte]() = _
                        Encoding.ASCII.GetBytes(serverResponse)
                    networkStream.Write(sendBytes, 0, sendBytes.Length)
                    networkStream.Flush()
                    serverResponse = String.Empty
                End If
            Catch ex As Exception
                'MsgBox(ex.ToString)
                serverSocket.Stop()
                StarSocket()
                Exit While
            End Try
        End While


        clientSocket.Close()
        serverSocket.Stop()

        'Console.ReadLine()

    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
       
        Dim thr As New Thread(AddressOf StarSocket)
        thr.IsBackground = True
        thr.Start()
        
        'requestCount = 0

        'While (True)
        '    Try
        '        requestCount = requestCount + 1
        '        Dim networkStream As NetworkStream = _
        '                clientSocket.GetStream()
        '        Dim bytesFrom(10024) As Byte
        '        networkStream.Read(bytesFrom, 0, CInt(clientSocket.ReceiveBufferSize))
        '        Dim dataFromClient As String = _
        '                System.Text.Encoding.ASCII.GetString(bytesFrom)
        '        dataFromClient = _
        '    dataFromClient.Substring(0, dataFromClient.IndexOf("$"))

        '        Dim serverResponse As String = _
        '            "Server response " + Convert.ToString(requestCount)
        '        Dim sendBytes As [Byte]() = _
        '            Encoding.ASCII.GetBytes(serverResponse)
        '        networkStream.Write(sendBytes, 0, sendBytes.Length)
        '        networkStream.Flush()

        '    Catch ex As Exception
        '        MsgBox(ex.ToString)
        '    End Try
        'End While


        'clientSocket.Close()
        'serverSocket.Stop()

        'Console.ReadLine()
    End Sub
End Class
