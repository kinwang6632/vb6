Public Class srvCardLess
    Private GatewayProcess As CableSoft.Cardless.Process.RunProcess = Nothing
    Protected Overrides Sub OnStart(ByVal args() As String)
        ' 在此加入啟動服務的程式碼。這個方法必須設定已啟動的
        ' 事項，否則可能導致服務無法工作。
        If GatewayProcess Is Nothing Then
            GatewayProcess = New CableSoft.Cardless.Process.RunProcess()
        End If
        GatewayProcess.NeedRegister = True
        GatewayProcess.RunProcess()

    End Sub

    Protected Overrides Sub OnStop()
        ' 在此加入停止服務所需執行的終止程式碼。
        GatewayProcess.Dispose()
        GatewayProcess = Nothing
    End Sub

End Class
