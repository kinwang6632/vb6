Imports System.ServiceProcess

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class AutoInvoiceService
    Inherits System.ServiceProcess.ServiceBase

    'UserService 覆寫 Dispose 以清除元件清單。
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    ' 處理序的主要進入點
    <MTAThread()> _
    <System.Diagnostics.DebuggerNonUserCode()> _
    Shared Sub Main()
        Dim ServicesToRun() As System.ServiceProcess.ServiceBase

        ' 在同一個處理序中可以執行多個 NT 服務。若要在這個處理序中
        ' 加入另一項服務，請修改下行程式碼，
        ' 以建立第二個服務物件。例如，
        '
        '   ServicesToRun = New System.ServiceProcess.ServiceBase () {New Service1, New MySecondUserService}
        '
        ServicesToRun = New System.ServiceProcess.ServiceBase() {New AutoInvoiceService}

        System.ServiceProcess.ServiceBase.Run(ServicesToRun)
    End Sub

    '為元件設計工具的必要項
    Private components As System.ComponentModel.IContainer

    ' 注意: 以下為元件設計工具的所需的程序
    ' 您可以使用元件設計工具進行修改。
    ' 請不要使用程式碼編輯器進行修改。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        '
        'AutoInvoiceService
        '
        Me.ServiceName = "srvInvoiceGateway"

    End Sub

End Class
