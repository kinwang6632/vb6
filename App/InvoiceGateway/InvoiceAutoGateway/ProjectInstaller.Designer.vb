<System.ComponentModel.RunInstaller(True)> Partial Class ProjectInstaller
    Inherits System.Configuration.Install.Installer

    'Installer 覆寫 Dispose 以清除元件清單。
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

    '為元件設計工具的必要項
    Private components As System.ComponentModel.IContainer

    '注意: 以下為元件設計工具所需的程序
    '您可以使用元件設計工具進行修改。
    '請不要使用程式碼編輯器進行修改。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.SvcProInstaller = New System.ServiceProcess.ServiceProcessInstaller()
        Me.svcInstaller = New System.ServiceProcess.ServiceInstaller()
        '
        'SvcProInstaller
        '
        Me.SvcProInstaller.Account = System.ServiceProcess.ServiceAccount.LocalSystem
        Me.SvcProInstaller.Password = Nothing
        Me.SvcProInstaller.Username = Nothing
        '
        'svcInstaller
        '
        Me.svcInstaller.Description = "開博自動產生發票系統電子檔"
        Me.svcInstaller.DisplayName = "CableSoft Invoice Gateway"
        Me.svcInstaller.ServiceName = "InvoiceGateway"
        '
        'ProjectInstaller
        '
        Me.Installers.AddRange(New System.Configuration.Install.Installer() {Me.SvcProInstaller, Me.svcInstaller})

    End Sub
    Friend WithEvents SvcProInstaller As System.ServiceProcess.ServiceProcessInstaller
    Friend WithEvents svcInstaller As System.ServiceProcess.ServiceInstaller

End Class
