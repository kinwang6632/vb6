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
        Me.srvproInstall = New System.ServiceProcess.ServiceProcessInstaller()
        Me.srvInstall = New System.ServiceProcess.ServiceInstaller()
        '
        'srvproInstall
        '
        Me.srvproInstall.Account = System.ServiceProcess.ServiceAccount.LocalSystem
        Me.srvproInstall.Password = Nothing
        Me.srvproInstall.Username = Nothing
        '
        'srvInstall
        '
        Me.srvInstall.Description = "CableSoft CardLess Gateway Service"
        Me.srvInstall.DisplayName = "CableSoft.CardLess.Gateway"
        Me.srvInstall.ServiceName = "CableSoft.CardLess.Gateway"
        '
        'ProjectInstaller
        '
        Me.Installers.AddRange(New System.Configuration.Install.Installer() {Me.srvproInstall, Me.srvInstall})

    End Sub
    Friend WithEvents srvproInstall As System.ServiceProcess.ServiceProcessInstaller
    Friend WithEvents srvInstall As System.ServiceProcess.ServiceInstaller

End Class
