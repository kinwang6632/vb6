Imports DevExpress.XtraTreeList.Nodes
Imports Microsoft.Win32
Imports System.Security.AccessControl
Imports System.Reflection

Public Class Common
    Public Shared rfrmMain As rfrmMain
    Private Shared xmlRedear As CableSoft.NSTV.XMLFileIO.XmlFileIO = Nothing
    Private Const SetingFileName As String = "Monitor.Set"
    Public Shared FLstSysMsg As List(Of String)
    Public Shared tbSystem As DataTable = Nothing
    Public Shared tbSO As DataTable = Nothing
    Private Const ExeName As String = "CableSoft.NSTV.Process.dll"
    Private Shared GatewayProcess As Object = Nothing

    Public Enum SOStatus
        Initial = 0
        Yes = 1
        NO = 2
        Key = 3
        NotUpdateImg = 4
    End Enum
    Public Enum MsgStyle
        [None] = -1
        [OK] = 0
        [Error] = 1
        [Info] = 2
        [Warning] = 3
        [Question] = 4
        [Select] = 5
        [Cancel] = 6
        [Process] = 7
        [OKLa] = 9
        [ErrorLa] = 10
        [InfoLa] = 11
        [db] = 13
        [Select2] = 14
        [Select3] = 15
        [Wait] = 16
    End Enum
    Public Shared Function ReadSystem(ByVal IsEncryFile As Boolean) As DataTable
        If xmlRedear Is Nothing Then
            xmlRedear = New CableSoft.NSTV.XMLFileIO.XmlFileIO(Application.StartupPath & "\" & SetingFileName)
        End If

        Dim tbSystem As New DataTable
        Try
            tbSystem.Columns.Add(New DataColumn("Title", GetType(String)))
            tbSystem.Columns.Add(New DataColumn("ReadTime", GetType(Int32)))
            tbSystem.Columns.Add(New DataColumn("ProcCmdCount", GetType(Int32)))
            tbSystem.Columns.Add(New DataColumn("MaxThread", GetType(Int32)))
            tbSystem.Columns.Add(New DataColumn("DBMinPool", GetType(Int32)))
            tbSystem.Columns.Add(New DataColumn("DBMaxPool", GetType(Int32)))
            tbSystem.Columns.Add(New DataColumn("DBPoolLiveTime", GetType(Int32)))
            tbSystem.Columns.Add(New DataColumn("AutoRun", GetType(Int32)))
            tbSystem.Columns.Add(New DataColumn("ShowUIRecord", GetType(Int32)))
            tbSystem.Columns.Add(New DataColumn("StartGateway", GetType(Boolean)))
            Try
                Dim rwNew As DataRow = tbSystem.NewRow
                rwNew.Item("Title") = xmlRedear.ReadXmlNodeValue("CableSoft/System", "Title", IsEncryFile)
                rwNew.Item("ReadTime") = Decimal.Parse(xmlRedear.ReadXmlNodeValue("CableSoft/System", "ReadTime", IsEncryFile))
                rwNew.Item("ProcCmdCount") = Decimal.Parse(xmlRedear.ReadXmlNodeValue("CableSoft/System", "ProcCmdCount", IsEncryFile))
                rwNew.Item("MaxThread") = Decimal.Parse(xmlRedear.ReadXmlNodeValue("CableSoft/System", "MaxThread", IsEncryFile))
                rwNew.Item("DBMinPool") = Decimal.Parse(xmlRedear.ReadXmlNodeValue("CableSoft/System", "DBMinPool", IsEncryFile))
                rwNew.Item("DBMaxPool") = Decimal.Parse(xmlRedear.ReadXmlNodeValue("CableSoft/System", "DBMaxPool", IsEncryFile))
                rwNew.Item("DBPoolLiveTime") = Decimal.Parse(xmlRedear.ReadXmlNodeValue("CableSoft/System", "DBPoolLiveTime", IsEncryFile))
                rwNew.Item("AutoRun") = Decimal.Parse(xmlRedear.ReadXmlNodeValue("CableSoft/System", "AutoRun", IsEncryFile))
                rwNew.Item("ShowUIRecord") = Decimal.Parse(xmlRedear.ReadXmlNodeValue("CableSoft/System", "ShowUIRecord", IsEncryFile))
                rwNew.Item("StartGateway") = Decimal.Parse("0" & xmlRedear.ReadXmlNodeValue("CableSoft/System", "StartGateway", IsEncryFile)) = 1
                tbSystem.Rows.Add(rwNew)
                tbSystem.AcceptChanges()
            Catch ex As Exception
                Throw
            End Try
        Catch ex As Exception
            Throw
        Finally
            xmlRedear.Dispose()
            xmlRedear = Nothing
        End Try
        Return tbSystem
    End Function
    Public Shared Function ReadSO(ByVal IsEncryFile As Boolean) As DataTable
        If xmlRedear Is Nothing Then
            xmlRedear = New CableSoft.NSTV.XMLFileIO.XmlFileIO(Application.StartupPath & "\" & SetingFileName)
        End If
        Dim tbSO As New DataTable("SO")
        tbSO.Columns.Add(New DataColumn("IsSelect", GetType(String)))
        tbSO.Columns.Add(New DataColumn("CompCode", GetType(String)))
        tbSO.Columns.Add(New DataColumn("CompName", GetType(String)))
        tbSO.Columns.Add(New DataColumn("SID", GetType(String)))
        tbSO.Columns.Add(New DataColumn("User", GetType(String)))
        tbSO.Columns.Add(New DataColumn("Password", GetType(String)))

        Try
            Dim lstDBValue As List(Of Object) = xmlRedear.ReadXmlNodeAttributes("CableSoft/SODB", "DB", "Value", IsEncryFile)
            Dim lstCompCode As List(Of Object) = xmlRedear.ReadXmlNodeAttributes("CableSoft/SODB", "DB", "CompCode", IsEncryFile)
            Dim lstCompName As List(Of Object) = xmlRedear.ReadXmlNodeValues("CableSoft/SODB", "DB", IsEncryFile)
            Dim lstSID As List(Of Object) = xmlRedear.ReadXmlNodeAttributes("CableSoft/SODB", "DB", "SID", IsEncryFile)
            Dim lstUSER As List(Of Object) = xmlRedear.ReadXmlNodeAttributes("CableSoft/SODB", "DB", "USER", IsEncryFile)
            Dim lstPassword As List(Of Object) = xmlRedear.ReadXmlNodeAttributes("CableSoft/SODB", "DB", "PASSWORD", IsEncryFile)
            If lstDBValue IsNot Nothing Then
                For i As Int32 = 0 To lstDBValue.Count - 1
                    Dim data(5) As Object
                    data(0) = lstDBValue.Item(i)
                    data(1) = lstCompCode.Item(i)
                    data(2) = lstCompName.Item(i)
                    data(3) = lstSID.Item(i)
                    data(4) = lstUSER.Item(i)
                    data(5) = lstPassword.Item(i)
                    Dim rwNew As DataRow = tbSO.NewRow()
                    rwNew.Item("IsSelect") = Int32.Parse(data(0))
                    rwNew.Item("CompCode") = data(1)
                    rwNew.Item("CompName") = data(2)
                    rwNew.Item("SID") = data(3)
                    rwNew.Item("User") = data(4)
                    rwNew.Item("Password") = data(5)
                    tbSO.Rows.Add(rwNew)
                    tbSO.AcceptChanges()

                Next
            End If

        Catch ex As Exception
            Throw ex
        Finally
            xmlRedear.Dispose()
            xmlRedear = Nothing
        End Try
        Return tbSO
    End Function

    Public Shared Sub UpdateHighCmd(ByVal aryRw As Dictionary(Of String, Object), ByVal SO As SO)
        Try
            If rfrmMain.InvokeRequired Then
                Dim act As New Action(Of Dictionary(Of String, Object), SO)(AddressOf UpdateHighCmd)
                rfrmMain.Invoke(act, aryRw, SO)
            Else
                Dim ParentNote As New Threading.ThreadLocal(Of TreeListNode)
                If SO.IsStop Then
                    Exit Sub
                End If
                With rfrmMain
                    .TreLstInstruct.BeginUpdate()
                    '清除所有Notes資料

                    Try
                        ParentNote.Value = .TreLstInstruct.FindNodeByFieldValue("GUID",
                                                                                      String.Format("{0}-{1}",
                                                                                                    aryRw.Item("CompCode".ToUpper),
                                                                                                    aryRw.Item("SEQNO".ToUpper)))
                        If ParentNote.Value Is Nothing Then
                            If .TreLstInstruct.Nodes.Count >= SO.ShowUICount Then
                                If SO.WaitUpdCmd.ContainsKey(Int32.Parse(.TreLstInstruct.Nodes(0).GetValue("SeqNo")).ToString.ToUpper) Then
                                    SO.WaitUpdCmd.Remove(.TreLstInstruct.Nodes(0).GetValue("SeqNo").ToString.ToUpper)
                                End If
                                .TreLstInstruct.Nodes.RemoveAt(0)
                            End If
                            ParentNote.Value = .TreLstInstruct.AppendNode(Nothing, -1)
                        Else
                            If aryRw.Item("Cmd_Status".ToUpper) = ParentNote.Value.GetValue("Cmd_Status") Then
                                Exit Sub
                            End If
                            Select Case aryRw.Item("Cmd_Status".ToUpper).ToString.ToUpper
                                Case "W"
                                    ParentNote.Value.SetValue("Cmd_Status", "W")
                                    ParentNote.Value.StateImageIndex = MsgStyle.Wait
                                    Exit Sub
                                Case "P"
                                    ParentNote.Value.SetValue("Cmd_Status", "P")
                                    ParentNote.Value.StateImageIndex = MsgStyle.Process
                                    Exit Sub
                            End Select

                        End If
                        For i As Int32 = 0 To .TreLstInstruct.Columns.Count - 1
                            If .TreLstInstruct.Columns(i).FieldName.ToUpper = "GUID".ToUpper Then
                                ParentNote.Value.SetValue("GUID".ToUpper,
                                                    String.Format("{0}-{1}",
                                                    aryRw.Item("CompCode".ToUpper),
                                                    aryRw.Item("SEQNO".ToUpper)))
                            Else
                                ParentNote.Value.SetValue(.TreLstInstruct.Columns(i).FieldName,
                                                    aryRw.Item(.TreLstInstruct.Columns(i).FieldName.ToUpper))
                            End If
                            Select Case aryRw.Item("Cmd_Status".ToUpper)
                                Case "E", "T"
                                    ParentNote.Value.StateImageIndex = MsgStyle.ErrorLa
                                Case "P"
                                    ParentNote.Value.StateImageIndex = MsgStyle.Process
                                Case "W"
                                    ParentNote.Value.StateImageIndex = MsgStyle.Wait
                                Case Else
                                    ParentNote.Value.StateImageIndex = MsgStyle.OK
                            End Select
                        Next
                        .TreLstInstruct.FocusedNode = ParentNote.Value
                    Finally
                        .TreLstInstruct.EndUpdate()
                        System.Threading.Interlocked.Decrement(SO.UpdateUICount)
                        ParentNote.Dispose()
                    End Try
                End With

            End If
        Catch ex As Exception
        Finally
        End Try
    End Sub
    Public Shared Sub UpdateSysMsgUI()

        Try

            Dim trlParent As TreeListNode = Nothing
            Dim trlChild As TreeListNode = Nothing


            If FLstSysMsg Is Nothing Then
                Exit Sub
            End If
            If rfrmMain.InvokeRequired Then
                Dim act As New Action(AddressOf UpdateSysMsgUI)
                rfrmMain.Invoke(act)
            Else
                Try

                    rfrmMain.TreLstSysMsg.ClearNodes()
                    rfrmMain.TreLstSysMsg.BeginUpdate()

                    For i As Int32 = 0 To FLstSysMsg.Count - 1
                        If FLstSysMsg.Item(i).ToString.ToUpper.StartsWith("PARENT") Then
                            trlParent = rfrmMain.TreLstSysMsg.AppendNode(Nothing, i)

                            trlParent.SetValue("SysInfo", FLstSysMsg.Item(i).Replace("Parent", ""))
                            trlParent.StateImageIndex = MsgStyle.Info
                            '記錄事件

                        Else
                            trlParent = rfrmMain.TreLstSysMsg.Nodes.LastNode
                            trlChild = rfrmMain.TreLstSysMsg.AppendNode(Nothing, trlParent)
                            trlChild.SetValue("SysInfo", FLstSysMsg.Item(i))
                            trlChild.StateImageIndex = MsgStyle.InfoLa

                        End If
                    Next

                Finally
                    rfrmMain.TreLstSysMsg.EndUpdate()
                    rfrmMain.TreLstSysMsg.ExpandAll()
                    rfrmMain.TreLstSysMsg.MoveFirst()
                End Try

            End If

        Catch ex As Exception
            ' WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
            'Throw New CableSoft.Gateway.csException.CablesoftException(ex, True)
        End Try

    End Sub
    Public Shared Sub UpdateSODatabase(ByVal So As SO, ByVal Status As SOStatus, ByVal guid As String)
        Try
            If rfrmMain.InvokeRequired Then
                Dim act As New Action(Of SO, SOStatus, String)(AddressOf UpdateSODatabase)
                rfrmMain.Invoke(act, So, Status, guid)
            Else
                Dim trlParent As New Threading.ThreadLocal(Of TreeListNode)
                Dim trlChild As New Threading.ThreadLocal(Of TreeListNode)
                Try
                    rfrmMain.TreLstConnStatus.BeginUpdate()
                    If ((rfrmMain.TreLstConnStatus.Nodes.Count) >= So.ShowUICount) Then
                        rfrmMain.TreLstConnStatus.Nodes.Clear()
                    End If
                    If Status = SOStatus.NotUpdateImg Then
                        trlParent.Value = rfrmMain.TreLstConnStatus.AppendNode(Nothing, -1)
                        trlParent.Value.SetValue("SOInfo", String.Format(Format(Now, "yyyy/MM/dd HH:mm:ss") & _
                                           " > [ {0} ] 資料庫連接中..", So.CompName))
                        trlParent.Value.SetValue("uID", guid)
                        trlParent.Value.StateImageIndex = MsgStyle.db
                        Exit Sub
                    End If

                    trlParent.Value = rfrmMain.TreLstConnStatus.FindNodeByFieldValue("uID", guid)
                    If trlParent.Value IsNot Nothing Then
                        trlChild.Value = rfrmMain.TreLstConnStatus.AppendNode(Nothing, trlParent.Value)
                        Select Case Status
                            Case SOStatus.Yes
                                trlChild.Value.SetValue("SOInfo", Format(Now, "yyyy/MM/dd HH:mm:ss") & _
                                            " > [ " & So.CompName & " ] 資料庫連接成功")
                                trlChild.Value.StateImageIndex = MsgStyle.OKLa
                            Case Else
                                trlChild.Value.SetValue("SOInfo", Format(Now, "yyyy/MM/dd HH:mm:ss") & _
                                      " > [ " & So.CompName & " ] 資料庫連接失敗")
                                trlChild.Value.StateImageIndex = MsgStyle.ErrorLa

                        End Select
                    End If


                Finally
                    rfrmMain.TreLstConnStatus.ExpandAll()
                    rfrmMain.TreLstConnStatus.MoveLast()
                    rfrmMain.TreLstConnStatus.EndUpdate()
                    trlChild.Dispose()
                    trlParent.Dispose()
                End Try
            End If
        Catch ex As Exception
            'WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
        Finally
            'aObj.Dispose()
            'aObj = Nothing
        End Try
    End Sub
    Public Shared Sub SetSysMsgString()
        Try
            If FLstSysMsg Is Nothing Then
                FLstSysMsg = New List(Of String)
            Else
                FLstSysMsg.Clear()
            End If
            FLstSysMsg.Add("Parent讀取系統參數設定..")
            FLstSysMsg.Add(String.Format("Windows 開機後自動執行 Gateway : {0}", IIf(Integer.Parse("0" & tbSystem.Rows(0).Item("AutoRun")) = 1,
                                                                              "是", "否")))


            FLstSysMsg.Add(String.Format("每 {0} 秒讀系統台資料", tbSystem.Rows(0).Item("ReadTime")))
            FLstSysMsg.Add(String.Format("每次處理 {0} 筆高階指令", tbSystem.Rows(0).Item("ProcCmdCount")))
            FLstSysMsg.Add(String.Format("畫面資料顯示 {0} 筆訊息", tbSystem.Rows(0).Item("ShowUIRecord")))
            FLstSysMsg.Add(String.Format("最大命令處理線程 {0} ", tbSystem.Rows(0).Item("MaxThread")))

            FLstSysMsg.Add(String.Format("DB最小集區 : {0} , DB最大集區 : {1} , 集區Lifetime : {2}",
                                        tbSystem.Rows(0).Item("DBMinPool"),
                                        tbSystem.Rows(0).Item("DBMaxPool"),
                                        tbSystem.Rows(0).Item("DBPoolLiveTime")))
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub
    Public Shared Sub UpdateSysErrUI(ByVal exSource As Exception, _
                          ByVal aMsg As String)

        Try
            If String.IsNullOrEmpty(aMsg) AndAlso (exSource Is Nothing) Then
                Exit Sub
            End If
            If rfrmMain.InvokeRequired Then
                Dim act As New Action(Of Exception, String)(AddressOf UpdateSysErrUI)
                rfrmMain.Invoke(act, exSource, aMsg)
            Else
                With rfrmMain
                    Dim trlParent As New Threading.ThreadLocal(Of TreeListNode)
                    Dim trlChild As New Threading.ThreadLocal(Of TreeListNode)
                    Try

                        .TreLstSysErr.BeginUpdate()

                        If (.TreLstSysErr.Nodes.Count >= tbSystem.Rows(0).Item("ShowUIRecord")) Then
                            .TreLstSysErr.Nodes.RemoveAt(0)
                        End If
                        trlParent.Value = .TreLstSysErr.AppendNode(Nothing, -1)
                        trlParent.Value.SetValue("ErrMsg", Format(Now, "yyyy/MM/dd HH:mm:ss") & " > " & _
                               aMsg)
                        trlParent.Value.StateImageIndex = MsgStyle.Error
                        If Not String.IsNullOrEmpty(aMsg) Then
                            trlChild.Value = .TreLstSysErr.AppendNode(Nothing, trlParent.Value)
                            trlChild.Value.SetValue("ErrMsg", exSource.Message.Trim(Environment.NewLine).ToString().Trim(vbLf))
                            trlChild.Value.StateImageIndex = MsgStyle.ErrorLa
                        End If
                    Finally

                        .TreLstSysErr.EndUpdate()
                        .TreLstSysErr.ExpandAll()
                        trlParent.Dispose()
                        trlChild.Dispose()
                    End Try
                End With
            End If
        Catch ex As Exception
            'WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)

        End Try

    End Sub
    Public Shared Sub UpdateShowSO(ByVal UpdSO As SO, ByVal Status As SOStatus)
        Try
            If rfrmMain.InvokeRequired Then
                Dim act As New Action(Of SO, SOStatus)(AddressOf UpdateShowSO)
                rfrmMain.Invoke(act, UpdSO, Status)
            Else
                With rfrmMain
                    Dim Note As New Threading.ThreadLocal(Of TreeListNode)
                    Try
                        .treLstSO.BeginUpdate()
                        Select Case Status

                            Case SOStatus.Initial, SOStatus.Key
                                '.treLstSO.ClearNodes()

                                Note.Value = .treLstSO.AppendNode(Nothing, UpdSO.CompCode)
                                Note.Value.SetValue("CompId", UpdSO.CompCode)
                                Note.Value.SetValue("CompName", UpdSO.CompName)
                                Note.Value.SetValue("WaitProc", "0")
                                Note.Value.SetValue("Procing", "0")
                                Note.Value.StateImageIndex = Status
                            Case Else
                                Note.Value = .treLstSO.FindNodeByFieldValue("CompId", UpdSO.CompCode)
                                If Note.Value IsNot Nothing Then
                                    If Status = SOStatus.NO Then
                                        Note.Value.SetValue("WaitProc", "0")
                                        Note.Value.SetValue("Procing", "0")
                                    Else
                                        Note.Value.SetValue("WaitProc", UpdSO.WaitProcessCount)
                                        Note.Value.SetValue("Procing", UpdSO.ProcessingCount)
                                    End If
                                    Note.Value.StateImageIndex = Status
                                End If
                        End Select
                    Finally
                        'aTreeList.ExpandAll()
                        .treLstSO.EndUpdate()
                        Note.Dispose()
                    End Try
                End With
            End If

        Catch ex As Exception
            'WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
        Finally
            'aSO.Dispose()

        End Try
    End Sub
    Public Shared Function IsRegister() As Boolean
        Dim mAssembly As Assembly

        Dim CableSoftRegister As Object

        mAssembly = Assembly.LoadFile(Application.StartupPath & "\CableSoft_KeyPro.dll")
        CableSoftRegister = mAssembly.CreateInstance("CableSoft_KeyPro.GetSystemInfo", True,
                                              BindingFlags.CreateInstance, Nothing,
                                              Nothing, Nothing, Nothing)

        Return CableSoftRegister.IsRegister(Application.StartupPath & "\KeyPro.lic",
                                            ExeName, 0)
    End Function
    Public Shared Sub CloseGateway()
        If GatewayProcess IsNot Nothing Then
            GatewayProcess.Dispose()
            GatewayProcess = Nothing
        End If
    End Sub
    Public Shared Sub StartGateway()
        'Dim obj As New CableSoft.NSTV.Process.RunProcess()
        'obj.NeedRegister = True
        'obj.ExeName = ExeName
        'obj.RunProcess()
        If GatewayProcess IsNot Nothing Then
            Exit Sub
        End If
        Dim mAssembly As Assembly
        mAssembly = Assembly.LoadFile(Application.StartupPath & "\CableSoft.NSTV.Process.dll")
        GatewayProcess = mAssembly.CreateInstance("CableSoft.Nstv.Process.RunProcess", True,
                                              BindingFlags.CreateInstance, Nothing,
                                              Nothing, Nothing, Nothing)

        GatewayProcess.NeedRegister = True
        GatewayProcess.ExeName = ExeName
        GatewayProcess.RunProcess()
       
    End Sub


    Public Shared Sub ProcAutoRun(ByVal AutoExec As Boolean)
        Dim RegRoot As RegistryKey = Nothing
        Dim RegKey As RegistryKey = Nothing

        Try
            RegRoot = Registry.LocalMachine

            RegKey = RegRoot.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Run", _
                                        RegistryKeyPermissionCheck.ReadWriteSubTree, _
                                        RegistryRights.SetValue)


            Dim App As String = String.Format("{0} /AutoExec", Application.ExecutablePath)
            If AutoExec Then
                RegKey.SetValue("Cablesoft NSTV Gateway Monitor", App, RegistryValueKind.String)
            Else
                RegKey.DeleteValue("Cablesoft NSTV Gateway Monitor", False)
            End If
        Catch ex As Exception

        Finally
            RegKey.Close()
            RegRoot.Close()
        End Try
    End Sub
End Class
