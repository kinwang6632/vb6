Public Class ExecuteSO
    Implements IDisposable
    Private tbSODB As DataTable
    Private tbCommand As DataTable
    Private overseeSetting As New OverseeXmlSetting.XmlSetting
    Private lstSODB As New List(Of Cablesoft.Oversee.Gateway.SODB.SO)
    Public SODBArray As List(Of Cablesoft.Oversee.Gateway.SODB.SO)
    Public Event StatusChange(Status As Cablesoft.Oversee.Gateway.SODB.SO.StatusOfProcess,
                              ByVal sequence As Integer,
                              ByVal comId As Integer,
                              ByVal compName As String,
                              ByVal prgName As String,
                              ByVal msg As String)
    Public Sub New()
        Try
            tbSODB = overseeSetting.CreateSODB
            tbCommand = overseeSetting.CreateCommon
            If Not overseeSetting.ReadSetting(tbSODB, tbCommand) Then
                Throw New Exception("讀取設定檔錯誤！")
            End If
            Threading.ThreadPool.SetMaxThreads(Integer.Parse(tbCommand.Rows(0).Item("MaxThread")),
                                               Integer.Parse(tbCommand.Rows(0).Item("MaxThread")))
            For Each rw As DataRow In tbSODB.Rows
                If (Not DBNull.Value.Equals(rw.Item("SelectID"))) AndAlso (Integer.Parse(rw.Item("SelectID")) = 1) Then
                    Try
                        Dim SO As New Cablesoft.Oversee.Gateway.SODB.SO()
                        SO.CompCode = rw.Item("CompCode")
                        SO.CompName = rw.Item("CompName")
                        SO.prgName = rw.Item("prgName")
                        SO.Sid = rw.Item("Sid")
                        SO.DBUser = rw.Item("DBUser")
                        SO.DBPassword = rw.Item("DBPassword")
                        SO.DBMinPool = rw.Item("DBMinPool")
                        SO.DBMaxPool = rw.Item("DBMaxPool")
                        SO.DBLifetime = Integer.Parse(rw.Item("DBLifetime"))
                        SO.MonitorType = rw.Item("MonitorType").ToString.Replace(" ", "")
                        SO.Runtimer = Integer.Parse(rw.Item("Runtimer"))
                        SO.Timeout = Integer.Parse(rw.Item("Timeout"))
                        SO.sequence = rw.Item("sequence")
                        SO.logReserve = Integer.Parse(tbCommand.Rows(0).Item("ErrHistorical"))
                        If Not DBNull.Value.Equals(rw.Item("SeqSql")) Then
                            SO.SeqSql = rw.Item("SeqSql")
                        End If
                        If Not DBNull.Value.Equals(rw.Item("InsSql")) Then
                            SO.InsSql = rw.Item("InsSql")
                        End If
                        If Not DBNull.Value.Equals(rw.Item("GetSql")) Then
                            SO.GetSql = rw.Item("GetSql")
                        End If
                        If Not DBNull.Value.Equals(rw.Item("WebServiceName")) Then
                            SO.WebServiceName = rw.Item("WebServiceName")
                        End If
                        If Not DBNull.Value.Equals(rw.Item("WebSrvMethod")) Then
                            SO.WebSrvMethod = rw.Item("WebSrvMethod")
                        End If
                        If Not DBNull.Value.Equals(rw.Item("WebServicePara")) Then
                            SO.WebServicePara = rw.Item("WebServicePara").ToString.Replace(" ", "")
                        End If
                        If Not DBNull.Value.Equals(rw.Item("url")) Then
                            SO.url = rw.Item("url")
                        End If
                        If Not DBNull.Value.Equals(rw.Item("WebGetPara")) Then
                            SO.WebGetPara = rw.Item("WebGetPara")
                        End If
                        If Not DBNull.Value.Equals(rw.Item("WebPostPara")) Then
                            SO.WebPostPara = rw.Item("WebPostPara")
                        End If
                        lstSODB.Add(SO)
                        Threading.Thread.Sleep(100)
                    Catch ex As Exception
                        lstSODB.Clear()
                        Throw New Exception(ex.ToString)
                    End Try
                End If
            Next
            Me.SODBArray = lstSODB
        Catch ex As Exception
            Throw New Exception(ex.ToString)
        End Try
    End Sub
    Public Sub StopAll()
        Try
            If lstSODB IsNot Nothing Then
                For Each SO As Cablesoft.Oversee.Gateway.SODB.SO In lstSODB
                    If SO IsNot Nothing Then                        
                        SO.StopRun()
                        SO.Dispose()
                        'RemoveHandler SO.StatusChange, AddressOf prcStatus
                        SO = Nothing
                    End If
                Next
            End If
            lstSODB.Clear()
        Catch ex As Exception
            Throw ex
        End Try
    
    End Sub
    Public Sub ExecuteAll()
        Try
            If (lstSODB IsNot Nothing) AndAlso (lstSODB.Count > 0) Then
                For Each SO As Cablesoft.Oversee.Gateway.SODB.SO In lstSODB
                    SO.Execute()
                    AddHandler SO.StatusChange, AddressOf prcStatus
                Next
            Else
                Throw New Exception("無任何SO可執行")
            End If
        Catch ex As Exception
            Throw ex
        End Try
      
    End Sub
    Private Sub prcStatus(Status As Cablesoft.Oversee.Gateway.SODB.SO.StatusOfProcess,
                              ByVal sequence As Integer,
                              ByVal comId As Integer,
                              ByVal compName As String,
                              ByVal prgName As String,
                              ByVal msg As String)
        RaiseEvent StatusChange(Status, sequence, comId, compName, prgName, msg)
    End Sub
#Region "IDisposable Support"
    Private disposedValue As Boolean ' 偵測多餘的呼叫

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: 處置 Managed 狀態 (Managed 物件)。
                If overseeSetting IsNot Nothing Then
                    overseeSetting.Dispose()
                    overseeSetting = Nothing
                End If
                If lstSODB IsNot Nothing Then
                    For Each SO As Cablesoft.Oversee.Gateway.SODB.SO In lstSODB
                        If SO IsNot Nothing Then
                            SO.Dispose()
                            SO = Nothing
                        End If
                    Next
                End If
                lstSODB.Clear()

                RemoveHandler Me.StatusChange, AddressOf prcStatus
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
