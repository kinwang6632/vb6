<System.ComponentModel.RunInstaller(True)> Partial Class svcNstInstaller
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
        Me.svcprocInstall = New System.ServiceProcess.ServiceProcessInstaller()
        Me.svcInstall = New System.ServiceProcess.ServiceInstaller()
        '
        'svcprocInstall
        '
        Me.svcprocInstall.Account = System.ServiceProcess.ServiceAccount.LocalSystem
        Me.svcprocInstall.Password = Nothing
        Me.svcprocInstall.Username = Nothing
        '
        'svcInstall
        '
        Me.svcInstall.Description = "執行NSTV Gateway 監控命令程式"
        Me.svcInstall.DisplayName = "CableSoft NSTV Gateway"
        Me.svcInstall.ServiceName = "Service1"
        Me.svcInstall.StartType = System.ServiceProcess.ServiceStartMode.Automatic
        '
        'svcNstInstaller
        '
        Me.Installers.AddRange(New System.Configuration.Install.Installer() {Me.svcprocInstall, Me.svcInstall})

    End Sub
    Friend WithEvents svcprocInstall As System.ServiceProcess.ServiceProcessInstaller
    Friend WithEvents svcInstall As System.ServiceProcess.ServiceInstaller

End Class
