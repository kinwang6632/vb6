Public Class svcNstv
    Private runGateway As CableSoft.NSTV.Process.RunProcess
    Protected Overrides Sub OnStart(ByVal args() As String)
        ' 在此加入啟動服務的程式碼。這個方法必須設定已啟動的
        ' 事項，否則可能導致服務無法工作。

        Try
            runGateway = New CableSoft.NSTV.Process.RunProcess()
            runGateway.RunProcess()
        Catch ex As Exception            
            Me.OnStop()
        End Try
        
    End Sub

    Protected Overrides Sub OnStop()
        ' 在此加入停止服務所需執行的終止程式碼。
        Try
            If runGateway IsNot Nothing Then
                runGateway.Dispose()
                runGateway = Nothing
            End If
        Catch ex As Exception
        Finally
            runGateway = Nothing
        End Try
        
        
    End Sub

End Class
