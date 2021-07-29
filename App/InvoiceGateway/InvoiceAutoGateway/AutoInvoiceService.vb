Public Class AutoInvoiceService
    Private InvGateway As CableSoft.Invoice.Gateway.Execute.RunGateway = Nothing
    Protected Overrides Sub OnStart(ByVal args() As String)
        ' 在此加入啟動服務的程式碼。這個方法必須設定已啟動的
        ' 事項，否則可能導致服務無法工作。
        InvGateway = New CableSoft.Invoice.Gateway.Execute.RunGateway()
        InvGateway.Execute()
    End Sub

    Protected Overrides Sub OnStop()
        ' 在此加入停止服務所需執行的終止程式碼。
        InvGateway.StopGateway()
        InvGateway.Dispose()
        InvGateway = Nothing
    End Sub

End Class
