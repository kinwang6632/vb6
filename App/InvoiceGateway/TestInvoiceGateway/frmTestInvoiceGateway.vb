Imports CableSoft.Invoice.Gateway.Execute
Public Class frmTestInvoiceGateway
    Dim objExecute As CableSoft.Invoice.Gateway.Execute.RunGateway = Nothing
    Private Sub btnExecute_Click(sender As System.Object, e As System.EventArgs) Handles btnExecute.Click

        If objExecute Is Nothing Then
            objExecute = New CableSoft.Invoice.Gateway.Execute.RunGateway
        End If
        objExecute.Execute()
        btnExecute.Enabled = False
        btnStop.Enabled = True
        'Using obj As New CableSoft.Invoice.Gateway.Execute.RunGateway
        '    obj.Execute()
        'End Using
    End Sub

    Private Sub btnStop_Click(sender As System.Object, e As System.EventArgs) Handles btnStop.Click
        objExecute.Dispose()
        objExecute = Nothing
        btnExecute.Enabled = True
        btnStop.Enabled = False
    End Sub


End Class
