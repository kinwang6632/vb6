Public Class RunGateway
    Implements IDisposable
    Private Const xmlFileName As String = "InvoiceSetup.Set"
    Private SetupFile As String = Nothing
    Private lstSOInfo As List(Of CableSoft.Invoice.Gateway.SOInfo.SO)
    Private lstSORun As New List(Of CableSoft.Invoice.Gateway.SO.RunSO)
    Public Sub New()
        Me.New(String.Format("{0}{1}",
                             System.AppDomain.CurrentDomain.BaseDirectory,
                             xmlFileName))
    End Sub
    Public Sub New(ByVal SetupFile As String)
        Me.SetupFile = SetupFile
    End Sub


    Public Sub Execute()
        Try
            'Dim SOInfo As CableSoft.Invoice.Gateway.SOInfo.SO() =
            '        CableSoft.Invoice.Gateway.ReadSO.ReadSOInfo.CreateSOParameters(SetupFile)
           
            lstSOInfo = CableSoft.Invoice.Gateway.SO.ReadSOInfo.CreateSOParameters(SetupFile).ToList
            For Each SO As CableSoft.Invoice.Gateway.SOInfo.SO In lstSOInfo
                lstSORun.Add(New CableSoft.Invoice.Gateway.SO.RunSO(SO))
            Next
            For Each SORun As CableSoft.Invoice.Gateway.SO.RunSO In lstSORun
                SORun.Run()
            Next
        Catch ex As Exception
            CableSoft.NSTV.Log.NstvLog.WriteErrorLog(ex, Nothing)
        End Try
    End Sub
    Public Sub StopGateway()
        For Each SORun As CableSoft.Invoice.Gateway.SO.RunSO In lstSORun
            If SORun IsNot Nothing Then
                SORun.StopGateway()
                SORun.Dispose()
                SORun = Nothing
            End If
        Next

        lstSORun.Clear()
    End Sub

#Region "IDisposable Support"
    Private disposedValue As Boolean ' 偵測多餘的呼叫

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                StopGateway()
                ' TODO: 處置 Managed 狀態 (Managed 物件)。
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



