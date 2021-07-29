Imports System.IO
Imports System.Net
Imports CableSoft.Nstv.Log
Imports CableSoft.Nstv.Process

Public Class frmTest_NstvProcess

    Private client As New CableSoft.NSTV.SocketClient.Client(Sockets.AddressFamily.InterNetwork,
                                                           Sockets.SocketType.Stream, Sockets.ProtocolType.Tcp)
    Private objProcess As New CableSoft.Nstv.Process.RunProcess()
    Private Sub btnTestClient_Click(sender As System.Object, e As System.EventArgs)
        client.CheckStateTime = 3
        client.ConnectNstv("127.0.0.1", 7364)
    End Sub

    Private Sub frmTest_NstvProcess_FormClosing(sender As Object, e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If client IsNot Nothing Then
                client.Dispose()
            End If

            objProcess.Dispose()
        Catch ex As Exception

        End Try
        
    End Sub

    Private Sub btnNstvProcess_Click(sender As System.Object, e As System.EventArgs) Handles btnNstvProcess.Click

        'obj.RunProcess()
        objProcess.NeedRegister = False
        objProcess.RunProcess()
        btnNstvProcess.Enabled = False
    End Sub

    Private Sub frmTest_NstvProcess_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class
