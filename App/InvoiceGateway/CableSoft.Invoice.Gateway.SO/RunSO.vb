Imports System.Threading
Imports System.Data.OracleClient
Imports CableSoft.NSTV.Log
Imports System
Imports System.IO
Public Class RunSO
    Implements IDisposable
    Private SO As CableSoft.Invoice.Gateway.SOInfo.SO
    Private cn As OracleConnection = Nothing
    Private ti As RegisteredWaitHandle = Nothing
    Private evn As AutoResetEvent = Nothing
    Private mtxWait As Mutex = Nothing
    Private IsProcessing As Boolean = False
    Private StopProcess As Boolean = False
    Private invArgument As CableSoft.B07.Parameters.Argument = Nothing
    Private invIO As CableSoft.B07.SystemIO.Invoice = Nothing
    Private InvoiceConnectionString As String = Nothing
    Private EInvConnectionString As String = Nothing
    Private CCBConnectionSTring As String = Nothing
    Private fB08UploadTime As String = "X"
    Private dsB08 As DataSet = Nothing
    Private dtB08 As DataTable = Nothing
    Private objB08 As CableSoft.Invoice.BLL.B08 = Nothing
    Private Const B08TableName As String = "Para"
    Private Const DebugFileName As String = "CableSoftDebug.txt"
    Private IsDebugMode As Boolean = False
    Private InsertToLog As CableSoft.Invoice.InsertEventLog.LogToTxt = Nothing
    Dim TimeWatch As New Stopwatch
    Private Enum InvDate
        StartDay = 0
        EndDay = 1
    End Enum
    Public Sub New(ByVal SO As CableSoft.Invoice.Gateway.SOInfo.SO)
        Me.SO = SO
        mtxWait = New Mutex()
    End Sub

    Private Function CreateB08Table() As Boolean
        If dtB08 IsNot Nothing Then
            Return True
        End If
        dtB08 = New DataTable("Para")
        Try
            With dtB08.Columns
                .Add("CompID", GetType(Decimal))
                .Add("InvDate1", GetType(DateTime))
                .Add("InvDate2", GetType(DateTime))
                .Add("InvID1", GetType(String))
                .Add("InvID2", GetType(String))
                .Add("ImportFile", GetType(String))
                .Add("Email", GetType(Boolean))
                .Add("MobilePhone", GetType(Boolean))
                .Add("TVMail", GetType(Boolean))
                .Add("CMSend", GetType(Boolean))
                .Add("SendOrder", GetType(String))
                .Add("AddSend", GetType(String))
                .Add("DonateMark", GetType(Decimal))
                .Add("ExportPath", GetType(String))
                .Add("UpdId", GetType(String))
                .Add("UpdName", GetType(String))
                .Add("UpdTime", GetType(DateTime))
                .Add("TmpTableName", GetType(String))
                .Add("AutoDo", GetType(Boolean))
                .Add("EmailNotify", GetType(Decimal))
                .Add("SmsNotify", GetType(Decimal))
                .Add("TVMAILNotify", GetType(Decimal))
                .Add("CMNotify", GetType(Decimal))
                .Add("EndCMNotify", GetType(Decimal))
                .Add("CompType", GetType(Decimal))
                .Add("MaskInvNo", GetType(Decimal))
                .Add("StartNotifyPrize", GetType(Decimal))
                .Add("DBLinkStr", GetType(String))
                .Add("UploadTime", GetType(String))
                .Add("OwnerName", GetType(String))
            End With

        Catch ex As Exception
            Throw
        End Try
        Return True
    End Function
    Private Function AddDefaultB08Data() As Boolean
        dtB08.Rows.Clear()
        Try
            Dim rwNew As DataRow = dtB08.NewRow
            rwNew.Item("CompID") = SO.CompCode
            'rwNew.Item("InvID1") = "0"
            'rwNew.Item("InvID2") = "0"
            rwNew.Item("Email") = SO.B08Data.IsEmailNotify
            rwNew.Item("MobilePhone") = SO.B08Data.IsSmsNotify
            rwNew.Item("TVMail") = SO.B08Data.IsTvMailNotify
            rwNew.Item("CMSend") = SO.B08Data.IsCmNotify
            rwNew.Item("SendOrder") = SO.B08Data.SendOrd
            rwNew.Item("AddSend") = SO.B08Data.AddSend
            rwNew.Item("DonateMark") = SO.B08Data.DonateMark
            rwNew.Item("UpdId") = SO.LoginUserId
            rwNew.Item("UpdName") = SO.LoginUserName
            rwNew.Item("AutoDo") = SO.B08Data.AutoDo
            rwNew.Item("EmailNotify") = SO.B08Data.StartEmailNotify
            rwNew.Item("SmsNotify") = SO.B08Data.StartSMSNotify
            rwNew.Item("TVMAILNotify") = SO.B08Data.StartTVMailNotify
            rwNew.Item("CMNotify") = SO.B08Data.StartCMNotify
            rwNew.Item("EndCMNotify") = SO.B08Data.EndCMNotify
            rwNew.Item("CompType") = SO.B08Data.CompType
            rwNew.Item("MaskInvNo") = SO.B08Data.MaskInvNo
            rwNew.Item("StartNotifyPrize") = SO.B08Data.StartNotifyPrize
            If Not String.IsNullOrEmpty(SO.DBLink) Then
                rwNew.Item("DBLinkStr") = SO.DBLink
            Else
                rwNew.Item("DBLinkStr") = "X"
            End If
            If Not String.IsNullOrEmpty(SO.MisOwner) Then
                rwNew.Item("OwnerName") = SO.MisOwner
            Else
                rwNew.Item("OwnerName") = "X"
            End If
            If Not String.IsNullOrEmpty(SO.B08Data.ExportPath) Then
                rwNew.Item("ExportPath") = SO.B08Data.ExportPath
            End If
            dtB08.Rows.Add(rwNew)
            dtB08.AcceptChanges()
        Catch ex As Exception
            Throw
        End Try
        Return True
    End Function
    Public Sub Run()
        Dim connectionBuilder As New OracleConnectionStringBuilder

        Try
            If evn IsNot Nothing Then
                evn.Dispose()
                evn = Nothing
            End If
            evn = New AutoResetEvent(False)
            connectionBuilder.Unicode = True
            connectionBuilder.PersistSecurityInfo = True
            connectionBuilder.DataSource = SO.InvoiceDBSid
            connectionBuilder.UserID = SO.InvoiceDBUserId
            connectionBuilder.Password = SO.InvoiceDBPassword
            connectionBuilder.MaxPoolSize = SO.MaxDBPool
            connectionBuilder.MinPoolSize = SO.MinDBPool
            connectionBuilder.LoadBalanceTimeout = SO.DBPollLifeTime
            InvoiceConnectionString = connectionBuilder.ConnectionString
            EInvConnectionString = connectionBuilder.ConnectionString
            Dim ComDBArray As String = String.Format("{0} {1} {2}",
                                                                 SO.InvoiceDBSid, SO.InvoiceDBUserId, SO.InvoiceDBPassword)
            Dim CCBDBArray As String = String.Format("{0} {1} {2}", "X", "X", "X")
            If Not String.IsNullOrEmpty(SO.ComDBSid) Then
                connectionBuilder.DataSource = SO.ComDBSid
                connectionBuilder.Password = SO.ComDBPassword
                connectionBuilder.UserID = SO.ComDBUserId
                EInvConnectionString = connectionBuilder.ConnectionString
                ComDBArray = String.Format("{0} {1} {2}",
                                                                  SO.ComDBSid, SO.ComDBUserId, SO.ComDBPassword)

            End If
            CCBConnectionSTring = String.Empty
            If Not String.IsNullOrEmpty(SO.MisDBSid) Then
                connectionBuilder.DataSource = SO.MisDBSid
                connectionBuilder.Password = SO.MisDBPassword
                connectionBuilder.UserID = SO.MisDBUserId
                CCBConnectionSTring = connectionBuilder.ConnectionString
                CCBDBArray = String.Format("{0} {1} {2}",
                                                                 SO.MisDBSid, SO.MisDBUserId, SO.MisDBPassword)
            End If
            connectionBuilder.Clear()


            objB08 = New CableSoft.Invoice.BLL.B08(String.Format("{0} {1} {2}",
                                                                 SO.InvoiceDBSid, SO.InvoiceDBUserId, SO.InvoiceDBPassword),
                                                                ComDBArray,
                                                             CCBDBArray)

            If dsB08 Is Nothing Then
                dsB08 = New DataSet()
            End If

            If Not CreateB08Table() Then
                NstvLog.WriteErrorLog(New Exception("建立B08 Table失敗"),
                                      SO.CompName & " 啟動失敗，無法執行")
                Exit Sub
            End If
            If Not AddDefaultB08Data() Then
                NstvLog.WriteErrorLog(New Exception("建立B08 預設值失敗"),
                                      SO.CompName & " 啟動失敗，無法執行")
                Exit Sub
            End If
            dsB08.Tables.Clear()
            dsB08.Tables.Add(dtB08)
            InsertToLog = New CableSoft.Invoice.InsertEventLog.LogToTxt(SO.CompCode, SO.IsLogSQL, SO.LogReserveTime)
            'InsertToLog = New CableSoft.Invoice.InsertEventLog.LogToMdb(SO.CompCode, SO.CompName)
            'InsertToLog.CanInsertAccessEvent = Me.CanLogToAccess
            'InsertToLog.InsertSOInformat()
#If DEBUG Then
            MsgBox("Debug模式，請連絡RD", MsgBoxStyle.Critical, "警告")
            ti = ThreadPool.RegisterWaitForSingleObject(evn, _
                                                             New WaitOrTimerCallback(AddressOf WaitProc), _
                                                             Nothing, TimeSpan.FromSeconds(3), False)

#Else
            If File.Exists(String.Format("{0}\{1}",
                                         Path.GetDirectoryName(System.AppDomain.CurrentDomain.BaseDirectory),
                                         DebugFileName)) Then

                IsDebugMode = True
            Else
                IsDebugMode = False
            End If

            If IsDebugMode Then
                ti = ThreadPool.RegisterWaitForSingleObject(evn, _
                                                           New WaitOrTimerCallback(AddressOf WaitProc), _
                                                           Nothing, TimeSpan.FromSeconds(3), False)
            Else
                ti = ThreadPool.RegisterWaitForSingleObject(evn, _
                                                         New WaitOrTimerCallback(AddressOf WaitProc), _
                                                         Nothing, TimeSpan.FromMinutes(SO.MonitorTime), False)
            End If
#End If


            evn.Set()
        Catch ex As Exception
            NstvLog.WriteErrorLog(ex, SO.CompName & " 啟動失敗，無法執行")
        End Try

    End Sub
    Private Sub WaitProc(ByVal state As Object, ByVal timedOut As Boolean)
        mtxWait.WaitOne()
        If (StopProcess) OrElse (IsProcessing) Then
            mtxWait.ReleaseMutex()
            Exit Sub
        End If
        IsProcessing = True
        TimeWatch.Reset()
        Try
            'Dim aUploadTime As DateTime = DateTime.Now
            If invArgument Is Nothing Then
                Dim aryArgument() As String = GetInvArgument()
                invArgument = New CableSoft.B07.Parameters.Argument(aryArgument)
                invArgument.MaxDBPool = SO.MaxDBPool
                invArgument.MinDBPool = SO.MinDBPool
                invArgument.LifePoolTime = SO.DBPollLifeTime
            End If
            If invIO Is Nothing Then
                invIO = New CableSoft.B07.SystemIO.Invoice(invArgument)
            End If

            invIO.IsAllCompCode = False
            invIO.UploadSource = B07.InvType.InvTypeEnum.UploadSource.Gateway
            invIO.CreateInvType = SO.CreateInvoiceType
            invIO.DestroyReason = SO.DestroyReason
            fB08UploadTime = "X"
            Dim RunFrequencyTime As String = Nothing
            Dim aNow As DateTime = invArgument.GetSysDate
            Dim aStartDate As DateTime = GetInvDate(aNow, InvDate.StartDay)
            Dim aEndDate As DateTime = GetInvDate(aNow, InvDate.EndDay)
            For Each RunCmd As CableSoft.B07.InvType.GatewayRunCommand In SO.RunCommands
                If StopProcess Then
                    IsProcessing = False
                    Exit Sub
                End If
                invIO.InvDate1 = ""
                invIO.InvDate2 = ""
                invIO.YearMonth1 = ""
                invIO.YearMonth2 = ""
                invIO.DiscountDate1 = ""
                invIO.DiscountDate2 = ""
                invIO.RunFrequencyTime = ""
                InsertToLog.UpdateProcessStatus(RunCmd.Command, InsertEventLog.Status.Processing)

#If DEBUG Then
                RunFrequencyTime = aNow.AddSeconds(0 - RunCmd.RunFrequency).ToString("yyyyMMddHHmmss")
#Else
                If IsDebugMode Then
                    RunFrequencyTime = aNow.AddSeconds(0 - RunCmd.RunFrequency).ToString("yyyyMMddHHmmss")
                Else
                    RunFrequencyTime = aNow.AddMinutes(0 - RunCmd.RunFrequency).ToString("yyyyMMddHHmmss")
                End If
#End If
                invIO.RunFrequencyTime = RunFrequencyTime
                Select Case RunCmd.Command
                    Case B07.InvType.InvTypeEnum.InvFileType.C0401
                        invIO.RunInvType = B07.InvType.InvTypeEnum.INVTYPE.OnlyNormalInv
                        invIO.InvDate1 = aStartDate.ToString("yyyy/MM/dd")
                        invIO.InvDate2 = aEndDate.ToString("yyyy/MM/dd")
                    Case B07.InvType.InvTypeEnum.InvFileType.C0501
                        invIO.RunInvType = B07.InvType.InvTypeEnum.INVTYPE.NormalInvOV
                        invIO.InvDate1 = aStartDate.ToString("yyyy/MM/dd")
                        invIO.InvDate2 = aEndDate.ToString("yyyy/MM/dd")
                    Case B07.InvType.InvTypeEnum.InvFileType.C0701
                        invIO.RunInvType = B07.InvType.InvTypeEnum.INVTYPE.DestroyInv
                        invIO.InvDate1 = aStartDate.ToString("yyyyMM")
                        invIO.InvDate1 = aEndDate.ToString("yyyyMM")
                    Case B07.InvType.InvTypeEnum.InvFileType.D0401
                        invIO.RunInvType = B07.InvType.InvTypeEnum.INVTYPE.OnlyDiscount
                        invIO.YearMonth1 = aStartDate.ToString("yyyyMM")
                        invIO.YearMonth2 = aEndDate.ToString("yyyyMM")
                    Case B07.InvType.InvTypeEnum.InvFileType.D0501
                        invIO.RunInvType = B07.InvType.InvTypeEnum.INVTYPE.OnlyDiscountOV
                        invIO.YearMonth1 = aStartDate.ToString("yyyyMM")
                        invIO.YearMonth2 = aEndDate.ToString("yyyyMM")
                    Case B07.InvType.InvTypeEnum.InvFileType.X0401
                        invIO.RunInvType = B07.InvType.InvTypeEnum.INVTYPE.DestroyReLoadInv
                        invIO.InvDate1 = aStartDate.ToString("yyyy/MM/dd")
                        invIO.InvDate2 = aEndDate.ToString("yyyy/MM/dd")
                    Case B07.InvType.InvTypeEnum.InvFileType.B08
                        dsB08.Tables(B08TableName).Rows(0).Item("InvDate1") = aStartDate
                        dsB08.Tables(B08TableName).Rows(0).Item("InvDate2") = aEndDate
                        dsB08.Tables(B08TableName).Rows(0).Item("UploadTime") = fB08UploadTime
                        dsB08.Tables(B08TableName).Rows(0).Item("UpdTime") = aNow
                        TimeWatch.Restart()
#If DEBUG Then
                        Try

                            dsB08.WriteXml(String.Format("{0}\{1}",
                                         Path.GetDirectoryName(System.AppDomain.CurrentDomain.BaseDirectory),
                                         "B08DataSet.xml"), XmlWriteMode.WriteSchema)
                        Catch ex As Exception
                        Finally

                        End Try

#Else
                        Try
                            If IsDebugMode Then
                                dsB08.WriteXml(String.Format("{0}\{1}",
                                        Path.GetDirectoryName(System.AppDomain.CurrentDomain.BaseDirectory),
                                        "B08DataSet.xml"), XmlWriteMode.WriteSchema)
                            End If
                        Catch ex As Exception
                        Finally

                        End Try

#End If

                        Try
                            If Not objB08.Submit(dsB08) Then
                                NstvLog.WriteErrorLog(New Exception("B08執行失敗"), Nothing)
                            End If
                        Catch ex As Exception
                            NSTV.Log.NstvLog.WriteErrorLog(ex, "B08執行有誤！")
                        Finally
                            InsertToLog.UpdateProcessStatus(RunCmd.Command, InsertEventLog.Status.Complete)
                            TimeWatch.Stop()
                            InsertToLog.InsertSQLTextLog(RunCmd.Command, "呼叫CableSoft.Invoice.BLL.B08.dll元件",
                                                   Math.Round(TimeWatch.Elapsed.TotalSeconds, 2), aNow)
                        End Try                                                                 
                        Continue For
                End Select

                TimeWatch.Restart()
                Using dsInv As DataSet = invIO.GetAllDataToDataSet
                    TimeWatch.Stop()
                    InsertToLog.InsertSQLTextLog(RunCmd.Command, invIO.SelectSQLText,
                                                 Math.Round(TimeWatch.Elapsed.TotalSeconds, 2), aNow)
                    'InsertToLog.InsertSQLTextLog(RunCmd.Command, invIO.SelectSQLText,
                    '                             Math.Round(TimeWatch.Elapsed.TotalSeconds, 2), aNow)
                    If 1 = 1 OrElse Integer.Parse(dsInv.Tables("INV001").Rows(0).Item("UploadType")) = 1 Then
                        If (dsInv.Tables("GiveErrData") IsNot Nothing) AndAlso
                            (dsInv.Tables("GiveErrData").Rows.Count > 0) Then
                            Dim aFile As String = "電子發票上傳異常資料記錄_" & _
                                aNow.ToString("yyyyMMdd_HHmmss") & ".txt"

                            If Not String.IsNullOrEmpty(SO.InvoiceBackupPath) Then
                                If Not SO.InvoiceBackupPath.EndsWith("\") Then
                                    aFile = SO.InvoiceBackupPath & "\" & aFile
                                Else
                                    aFile = SO.InvoiceBackupPath & aFile
                                End If
                                B07.SystemIO.Invoice.ShowGiveErrData(dsInv.Tables("GiveErrData"),
                                                                     SO.InvoiceBackupPath, aNow)
                                NstvLog.WriteErrorLog(New Exception("發票系統有異常資料，無法產上上傳檔案"),
                                                                String.Format("已產生檔案至：{0}",
                                                                              aFile))
                            Else
                                NstvLog.WriteErrorLog(New Exception("發票系統有異常資料，無法產上上傳檔案"), Nothing)
                            End If
                            Continue For
                        End If
                    End If
                    Using DataToFile As New CableSoft.B07.InvDataToFile.ToInvFile(invArgument, dsInv,
                                                                                  aNow, B07.InvType.InvTypeEnum.UploadSource.Gateway)
                        DataToFile.DestroyReason = SO.DestroyReason
                        DataToFile.UploadSource = B07.InvType.InvTypeEnum.UploadSource.Gateway
                        DataToFile.ExportElectronPath = RunCmd.ExportElectronPath
                        If DataToFile.WriteINVFile(invIO.RunInvType) Then
                            fB08UploadTime = aNow.ToString("yyyyMMddHHmmss")
                            If DataToFile.TotalCount = 0 Then
                                If Not String.IsNullOrEmpty(SO.InvoiceBackupPath) Then
                                    invIO.BackupScreenMsg("無任何資料", aNow, SO.InvoiceBackupPath,
                                                     [Enum].GetName(GetType(B07.InvType.InvTypeEnum.InvFileType), RunCmd.Command))
                                End If
                               
                            Else
                                If 1 = 1 OrElse Integer.Parse(dsInv.Tables("Inv001").Rows(0).Item("UploadType").ToString) = 1 Then
                                    fB08UploadTime = aNow.ToString("yyyyMMddHHmmss")
                                Else
                                    If invIO.UpdateINV(aNow.ToString, CableSoft.B07.SystemIO.UpdateInvKin.ALL) Then
                                        fB08UploadTime = aNow.ToString("yyyyMMddHHmmss")
                                    End If
                                End If
                                If Not String.IsNullOrEmpty(SO.InvoiceBackupPath) Then
                                    Dim aMsg As String = String.Empty
                                    aMsg = DataToFile.GetCompleteMsg(invIO.RunInvType)
                                    invIO.BackupScreenMsg(aMsg, aNow, SO.InvoiceBackupPath,
                                                          [Enum].GetName(GetType(B07.InvType.InvTypeEnum.InvFileType), RunCmd.Command))
                                End If
                            End If
                        End If
                    End Using

                End Using
                InsertToLog.UpdateProcessStatus(RunCmd.Command, InsertEventLog.Status.Complete)
                'InsertToLog.UpdateProcessStatus(RunCmd.Command, InsertEventLog.LogToMdb.Status.Complete)
            Next

        Catch ex As Exception
            NstvLog.WriteErrorLog(ex, Nothing)
        Finally
            InsertToLog.DeleteAllStatus()
            'InsertToLog.UpdateAllComplete()
            IsProcessing = False
            mtxWait.ReleaseMutex()
        End Try
    End Sub
    Private Function GetInvDate(ByVal aNow As DateTime, ByVal Flag As InvDate) As Date
        Dim RetDate As DateTime = aNow
        Dim aMonth As Integer = RetDate.Month
        Dim aDay As Integer = RetDate.Day
        Try
            '偶數月
            If aMonth Mod 2 = 0 Then
                Select Case Flag
                    Case InvDate.StartDay
                        RetDate = New DateTime(aNow.AddMonths(-1).Year, aNow.AddMonths(-1).Month, 1)
                    Case InvDate.EndDay
                        RetDate = New DateTime(RetDate.Year, RetDate.Month, RetDate.Day)
                End Select
                'If aDay >= SO.LimtBeforeUpload Then
                '    Select Case Flag
                '        Case InvDate.StartDay
                '            RetDate = New DateTime(RetDate.Year, RetDate.Month, 1)
                '        Case InvDate.EndDay
                '            RetDate = New DateTime(RetDate.Year, RetDate.Month, RetDate.Day)
                '    End Select
                'Else
                '    Select Case Flag
                '        Case InvDate.StartDay
                '            Dim BeforeDate As DateTime = RetDate
                '            BeforeDate = BeforeDate.AddMonths(-1)
                '            RetDate = New DateTime(BeforeDate.Year, BeforeDate.Month, 1)
                '        Case InvDate.EndDay
                '            Return RetDate
                '    End Select
                'End If
            Else
                '單數月
                If aDay >= SO.LimtBeforeUpload Then
                    Select Case Flag
                        Case InvDate.StartDay
                            RetDate = New DateTime(RetDate.Year, RetDate.Month, 1)
                        Case InvDate.EndDay
                            RetDate = New DateTime(RetDate.Year, RetDate.Month, RetDate.Day)
                    End Select
                Else

                    Select Case Flag
                        Case InvDate.StartDay
                            Dim BeforeDate As DateTime = RetDate
                            BeforeDate = BeforeDate.AddMonths(-2)
                            RetDate = New DateTime(BeforeDate.Year, BeforeDate.Month, 1)
                        Case InvDate.EndDay
                            Return RetDate
                    End Select
                End If
            End If

        Catch ex As Exception
            NstvLog.WriteErrorLog(ex, Nothing)
        End Try
        Return RetDate
    End Function
    Private Function GetInvArgument() As String()
        Dim lstRet As New List(Of String)
        Try
            With lstRet
                .Add(SO.InvoiceDBSid)               '0
                .Add(SO.InvoiceDBUserId)        '1
                .Add(SO.InvoiceDBPassword)  '2
                .Add(SO.CompCode)               '3
                .Add(SO.InvoiceDBUserId)            '4
                .Add("1") '是否啟用B08,由Gateway設定,這個參數隨便設就好     '5
                If String.IsNullOrEmpty(SO.ComDBSid) Then           '6
                    .Add(SO.InvoiceDBSid)
                Else
                    .Add(SO.ComDBSid)
                End If
                If String.IsNullOrEmpty(SO.ComDBUserId) Then            '7
                    .Add(SO.InvoiceDBUserId)
                Else
                    .Add(SO.ComDBUserId)
                End If
                If String.IsNullOrEmpty(SO.ComDBPassword) Then          '8
                    .Add(SO.InvoiceDBPassword)
                Else
                    .Add(SO.ComDBPassword)
                End If
                .Add(SO.LoginUserId) '第9個參數 LoginUserId
                .Add(SO.LoginUserName)      '10
                If String.IsNullOrEmpty(SO.DBLink) Then     '11
                    .Add("X")
                Else
                    .Add(SO.DBLink)
                End If

                If SO.StartAllUpload Then           '12
                    .Add(1)
                Else
                    .Add(0)
                End If

                If String.IsNullOrEmpty(SO.MisDBSid) Then

                End If
                If String.IsNullOrEmpty(SO.MisDBSid) Then   '13
                    .Add("X")
                Else
                    .Add(SO.MisDBSid)
                End If
                If String.IsNullOrEmpty(SO.MisDBUserId) Then   '14
                    .Add("X")
                Else
                    .Add(SO.MisDBUserId)
                End If
                If String.IsNullOrEmpty(SO.MisDBPassword) Then   '15
                    .Add("X")
                Else
                    .Add(SO.MisDBPassword)
                End If
                If String.IsNullOrEmpty(SO.MisOwner) Then   '16
                    .Add("X")
                Else
                    .Add(SO.MisOwner)
                End If
            End With
        Catch ex As Exception
            Throw
        End Try
        Return lstRet.ToArray
    End Function
    Public Sub StopGateway()
        StopProcess = True
        Dim StopTime As New Stopwatch()
        StopTime.Start()
        Do
            If (Not IsProcessing) OrElse (StopTime.Elapsed.Seconds >= 60) Then
                If ti IsNot Nothing Then
                    ti.Unregister(Nothing)
                    ti = Nothing
                End If
                Exit Do
            End If
            Thread.Sleep(100)
        Loop
        StopTime.Stop()
        If cn IsNot Nothing Then
            cn.Close()
            cn.Dispose()
            cn = Nothing
        End If
        If mtxWait IsNot Nothing Then
            mtxWait.Close()
            mtxWait.Dispose()
        End If
        If SO IsNot Nothing Then
            SO.Dispose()
            SO = Nothing
        End If
        If invArgument IsNot Nothing Then
            invArgument.Dispose()
            invArgument = Nothing
        End If
        If invIO IsNot Nothing Then
            invIO.Dispose()
            invIO = Nothing
        End If
        If dtB08 IsNot Nothing Then
            dtB08.Dispose()
            dtB08 = Nothing
        End If
        If dsB08 IsNot Nothing Then
            dsB08.Dispose()
            dsB08 = Nothing
        End If
        If objB08 IsNot Nothing Then
            objB08 = Nothing
        End If
        If InsertToLog IsNot Nothing Then
            InsertToLog.Dispose()
            InsertToLog = Nothing
        End If
    End Sub

#Region "IDisposable Support"
    Private disposedValue As Boolean ' 偵測多餘的呼叫

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: 處置 Managed 狀態 (Managed 物件)。

                StopGateway()
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
