Public Class ReadSOInfo
    Implements IDisposable
    Private Shared xmlSetupFile As String
    Private Const SetupFileName As String = "InvoiceSetup.Set"
    Private Const IsEncryFile As Boolean = True
    
    Private Shared Function GetCmdName(ByVal Data As String) As String
        Dim Ret As String = Nothing
        Try
            If (Not Data.Contains("(")) OrElse ((Not Data.Contains(")"))) Then
                Ret = Ret.Replace(" ", "")
            Else
                Dim StartIndex As Integer = Data.IndexOf("(")
                Dim EndIndex As Integer = Data.IndexOf(")")
                Ret = Data.Substring(StartIndex + 1, (EndIndex - StartIndex) - 1).Replace(" ", "")
            End If
        Catch ex As Exception
            Throw
        End Try
        Return Ret
    End Function
    Private Shared Function ReadRunListCommand(ByVal Data As String, ByVal tbCommand As DataTable,
                                               ByVal tbRunCommand As DataTable) As List(Of CableSoft.B07.InvType.GatewayRunCommand)

        Dim lstRet As New List(Of CableSoft.B07.InvType.GatewayRunCommand)

        Try
            For Each Str As String In Data.Split(",")
                Dim lstCommand As List(Of DataRow) = tbCommand.Select("CodeNo = '" & Str & "'").ToList
                If (lstCommand IsNot Nothing) AndAlso (lstCommand.Count > 0) Then
                    Dim lstRunCommand As List(Of DataRow) = tbRunCommand.Select("CodeNo='" & Str & "'").ToList
                    If (lstRunCommand IsNot Nothing) AndAlso (lstRunCommand.Count > 0) Then
                        Dim CommandStructure As New CableSoft.B07.InvType.GatewayRunCommand
                        CommandStructure.RunFrequency = Integer.Parse(lstRunCommand.Item(0).Item("RunFrequency"))
                        Select Case GetCmdName(lstCommand.Item(0).Item("CmdName").ToString)
                            Case [Enum].GetName(GetType(CableSoft.B07.InvType.InvTypeEnum.InvFileType),
                                                CableSoft.B07.InvType.InvTypeEnum.InvFileType.A0401).ToString.ToUpper
                                CommandStructure.Command = B07.InvType.InvTypeEnum.InvFileType.A0401
                            Case [Enum].GetName(GetType(CableSoft.B07.InvType.InvTypeEnum.InvFileType),
                                                CableSoft.B07.InvType.InvTypeEnum.InvFileType.A0501).ToString.ToUpper
                                CommandStructure.Command = B07.InvType.InvTypeEnum.InvFileType.A0501
                            Case [Enum].GetName(GetType(CableSoft.B07.InvType.InvTypeEnum.InvFileType),
                                               CableSoft.B07.InvType.InvTypeEnum.InvFileType.B0301).ToString.ToUpper
                                CommandStructure.Command = B07.InvType.InvTypeEnum.InvFileType.B0301
                            Case [Enum].GetName(GetType(CableSoft.B07.InvType.InvTypeEnum.InvFileType),
                                               CableSoft.B07.InvType.InvTypeEnum.InvFileType.B0401).ToString.ToUpper
                                CommandStructure.Command = B07.InvType.InvTypeEnum.InvFileType.B0401
                            Case [Enum].GetName(GetType(CableSoft.B07.InvType.InvTypeEnum.InvFileType),
                                               CableSoft.B07.InvType.InvTypeEnum.InvFileType.B08).ToString.ToUpper
                                CommandStructure.Command = B07.InvType.InvTypeEnum.InvFileType.B08
                            Case [Enum].GetName(GetType(CableSoft.B07.InvType.InvTypeEnum.InvFileType),
                                          CableSoft.B07.InvType.InvTypeEnum.InvFileType.C0401).ToString.ToUpper
                                CommandStructure.Command = B07.InvType.InvTypeEnum.InvFileType.C0401
                            Case [Enum].GetName(GetType(CableSoft.B07.InvType.InvTypeEnum.InvFileType),
                                          CableSoft.B07.InvType.InvTypeEnum.InvFileType.C0501).ToString.ToUpper
                                CommandStructure.Command = B07.InvType.InvTypeEnum.InvFileType.C0501
                            Case [Enum].GetName(GetType(CableSoft.B07.InvType.InvTypeEnum.InvFileType),
                                              CableSoft.B07.InvType.InvTypeEnum.InvFileType.C0701).ToString.ToUpper
                                CommandStructure.Command = B07.InvType.InvTypeEnum.InvFileType.C0701
                            Case [Enum].GetName(GetType(CableSoft.B07.InvType.InvTypeEnum.InvFileType),
                                              CableSoft.B07.InvType.InvTypeEnum.InvFileType.D0401).ToString.ToUpper
                                CommandStructure.Command = B07.InvType.InvTypeEnum.InvFileType.D0401
                            Case [Enum].GetName(GetType(CableSoft.B07.InvType.InvTypeEnum.InvFileType),
                                          CableSoft.B07.InvType.InvTypeEnum.InvFileType.D0501).ToString.ToUpper
                                CommandStructure.Command = B07.InvType.InvTypeEnum.InvFileType.D0501
                            Case [Enum].GetName(GetType(CableSoft.B07.InvType.InvTypeEnum.InvFileType),
                                          CableSoft.B07.InvType.InvTypeEnum.InvFileType.E0402).ToString.ToUpper
                                CommandStructure.Command = B07.InvType.InvTypeEnum.InvFileType.E0402
                            Case [Enum].GetName(GetType(CableSoft.B07.InvType.InvTypeEnum.InvFileType),
                                          CableSoft.B07.InvType.InvTypeEnum.InvFileType.X0401.ToString.ToUpper)
                                CommandStructure.Command = B07.InvType.InvTypeEnum.InvFileType.X0401
                            Case Else
                                CommandStructure.Command = B07.InvType.InvTypeEnum.InvFileType.UNKnow
                        End Select
                        lstRet.Add(CommandStructure)
                    End If
                End If
            Next
        Catch ex As Exception
            Throw
        End Try
        Return lstRet
    End Function
    Private Shared Function ReadCreateInvoiceType(ByVal Data As String) As List(Of CableSoft.B07.InvType.InvTypeEnum.CreateInvoiceType)
        Dim lstRet As New List(Of CableSoft.B07.InvType.InvTypeEnum.CreateInvoiceType)
        Try
            For Each Str As String In Data.Split(",")
                If Not lstRet.Contains(Integer.Parse(Str.Trim.Substring(0, 1))) Then
                    lstRet.Add(Integer.Parse(Str.Trim.Substring(0, 1)))
                End If
            Next
            If (lstRet.Contains(B07.InvType.InvTypeEnum.CreateInvoiceType.icAfter)) AndAlso
                (lstRet.Contains(B07.InvType.InvTypeEnum.CreateInvoiceType.icLocale) AndAlso
                    (lstRet.Contains(B07.InvType.InvTypeEnum.CreateInvoiceType.icNormal)) AndAlso
                    (lstRet.Contains(B07.InvType.InvTypeEnum.CreateInvoiceType.icPrev))) Then
                lstRet.Clear()
                lstRet.Add(B07.InvType.InvTypeEnum.CreateInvoiceType.icAll)
            End If
        Catch ex As Exception
            Throw
        End Try
        Return lstRet
    End Function

    Public Shared Function CreateSOParameters(ByVal SetupFile As String) As CableSoft.Invoice.Gateway.SOInfo.SO()
        xmlSetupFile = SetupFile
        Dim SetupIO As New CableSoft.Invoice.Gateway.SetupFileIO.InvoiceSetupIO(SetupFile)
        'Dim tbSystem As DataTable = SetupIO.ReadSystem(IsEncryFile)
        Dim tbCommand As DataTable = SetupIO.ReadCommand(IsEncryFile)
        Dim tbRunCommand As DataTable = SetupIO.ReadRunCommand(IsEncryFile)
        Dim tbSO As DataTable = SetupIO.ReadSO(IsEncryFile)
        Dim lstSO As New List(Of CableSoft.Invoice.Gateway.SOInfo.SO)
        Try
            For Each rw As DataRow In tbSO.Rows
                If (Not DBNull.Value.Equals(rw.Item("SelectID"))) AndAlso
                    (Integer.Parse(rw.Item("SelectID")) = 1) Then
                    Dim SingleSO As New CableSoft.Invoice.Gateway.SOInfo.SO()
                    SingleSO.ComDBPassword = rw.Item("ComDBPassword")
                    SingleSO.ComDBSid = rw.Item("ComDBSid")
                    SingleSO.ComDBUserId = rw.Item("ComDbUser")
                    SingleSO.CompCode = rw.Item("CompID")
                    SingleSO.CompName = rw.Item("CompName")
                    SingleSO.ComDBPassword = rw.Item("ComDBPassword")
                    SingleSO.CreateInvoiceType = ReadCreateInvoiceType(rw.Item("InvoiceType"))
                    SingleSO.DBPollLifeTime = Integer.Parse(rw.Item("DBPoolLifetime"))
                    SingleSO.InvoiceDBPassword = rw.Item("DBPassword")
                    SingleSO.InvoiceDBSid = rw.Item("DBSid")
                    SingleSO.InvoiceDBUserId = rw.Item("DBUser")
                    SingleSO.MaxDBPool = Integer.Parse(rw.Item("MaxDBPool"))
                    SingleSO.MinDBPool = Integer.Parse(rw.Item("MinDBPool"))
                    SingleSO.MonitorTime = Integer.Parse(rw.Item("RunCmdTimer"))
                    SingleSO.RunCommands = ReadRunListCommand(rw.Item("RunCommands"),
                                                            tbCommand, tbRunCommand)
                    lstSO.Add(SingleSO)
                End If
            Next
        Catch ex As Exception
            Throw
        Finally
            'tbSystem.Dispose()           
            tbSO.Dispose()
            tbSO = Nothing
            tbCommand.Dispose()
            tbCommand = Nothing
            tbRunCommand.Dispose()
            tbRunCommand = Nothing
            SetupIO.Dispose()
            SetupIO = Nothing
        End Try
        Return lstSO.ToArray
    End Function
    Public Shared Function CreateSOParameters() As CableSoft.Invoice.Gateway.SOInfo.SO()
        xmlSetupFile = String.Format("{0}\{1}",
                                     Environment.CurrentDirectory,
                                     SetupFileName)
        Return CreateSOParameters(xmlSetupFile)
    End Function
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
