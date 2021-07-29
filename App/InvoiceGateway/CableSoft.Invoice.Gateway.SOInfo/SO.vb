
Public Class SO
    Implements IDisposable

    Public Sub New()

    End Sub
    Public Property CompCode As Integer
    Public Property CompName As String
    Public Property InvoiceDBSid As String
    Public Property InvoiceDBUserId As String
    Public Property InvoiceDBPassword As String
    Public Property MisDBSid As String
    Public Property MisDBUserId As String
    Public Property MisDBPassword As String
    Public Property MisOwner As String
    Public Property ComDBSid As String
    Public Property ComDBUserId As String
    Public Property ComDBPassword As String
    Public Property MinDBPool As Integer
    Public Property MaxDBPool As Integer
    Public Property DBPollLifeTime As Integer
    Public Property MonitorTime As Integer
    Public Property CreateInvoiceType As List(Of CableSoft.B07.InvType.InvTypeEnum.CreateInvoiceType)
    Public Property RunCommands As List(Of CableSoft.B07.InvType.GatewayRunCommand)
    Public Property DBLink As String
    Public Property LoginUserId As String
    Public Property LoginUserName As String
    Public Property DestroyReason As String = "Gateway上傳註銷"
    Public Property StartAllUpload As Boolean
    Public Property LimtBeforeUpload As Integer
    Public Property InvoiceBackupPath As String
    Public Property IsEmailNotify As Boolean
    Public Property IsLogSQL As Boolean
    Public Property LogReserveTime As Integer
    Public Property B08Data As B08Info
#Region "IDisposable Support"
    Private disposedValue As Boolean ' 偵測多餘的呼叫

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
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
Public Structure B08Info
    Public Property CompId As Integer
    Public Property IsEmailNotify As Boolean
    Public Property IsSmsNotify As Boolean
    Public Property IsTvMailNotify As Boolean
    Public Property IsCmNotify As Boolean
    Public Property SendOrd As String
    Public Property AddSend As String
    Public Property DonateMark As Integer
    Public Property ExportPath As String
    Public Property AutoDo As Boolean
    Public Property StartEmailNotify As Integer
    Public Property StartSMSNotify As Integer
    Public Property StartTVMailNotify As Integer
    Public Property StartCMNotify As Integer
    Public Property EndCMNotify As Integer
    Public Property CompType As Integer
    Public Property MaskInvNo As Integer
    Public Property StartNotifyPrize As Integer
End Structure

'rwNew.Item("IsEmailNotify") = data.Item("IsEmailNotify")
'                   rwNew.Item("IsSmsNotify") = data.Item("IsSmsNotify")
'                   rwNew.Item("IsTvMailNotify") = data.Item("IsTvMailNotify")
'                   rwNew.Item("IsCmNotify") = data.Item("IsCmNotify")
'                   rwNew.Item("SendOrd") = data.Item("SendOrd")
'                   rwNew.Item("AddSend") = data.Item("AddSend")
'                   rwNew.Item("DonateMark") = data.Item("DonateMark")
'                   rwNew.Item("ExportPath") = data.Item("ExportPath")
'                   rwNew.Item("AutoDo") = data.Item("AutoDo")
'                   rwNew.Item("StartEmailNotify") = data.Item("StartEmailNotify")
'                   rwNew.Item("StartSMSNotify") = data.Item("StartSMSNotify")
'                   rwNew.Item("StartTVMailNotify") = data.Item("StartTVMailNotify")
'                   rwNew.Item("StartCMNotify") = data.Item("StartCMNotify")
'                   rwNew.Item("EndCMNotify") = data.Item("EndCMNotify")
'                   rwNew.Item("CompType") = data.Item("CompType")
'                   rwNew.Item("MaskInvNo") = data.Item("MaskInvNo")
'                   rwNew.Item("StartNotifyPrize") = data.Item("StartNotifyPrize")