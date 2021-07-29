Imports System
Imports System.Windows
Imports System.Windows.Forms
Imports System.IO
Imports System.Text
Imports System.Data
Imports System.Data.OracleClient
Imports System.Threading
Imports CableSoft.Gateway.Common
Imports CableSoft.Gateway.csException
Imports CableSoft.Gateway.SystemIO
Imports DevExpress.XtraEditors
Imports CableSoft_KeyPro

Module RunThread
    Private ti As RegisteredWaitHandle = Nothing
    Public evn As New AutoResetEvent(False)
    Private FMainForm As rfrmMain
    Private FThreadsTotal As Int32 = 0
    Public FSOIndex As New System.Collections.Generic.Dictionary(Of String, SOInfoClass)
    Private FExeCmdSQL As String
    Private aSQLCmd As String = "SELECT * FROM SO555 WHERE ( CMDSTATUS = 'W' OR CMDSTATUS ='P' )  " & _
                        " AND ROWNUM<={0} "
    Public rwl As New ReaderWriterLockSlim(LockRecursionPolicy.NoRecursion)
    Public FNagraStatusOK As Boolean = False
    Public FNagraConnecting As Boolean = False
    Public tmrNagraStatus As Threading.Timer = Nothing
    Public Class objNagra
        Public Shared NagraStatusOK As Boolean
    End Class
    Public Sub RunGateway(ByVal aParent As Form)
        Try

            FMainForm = aParent
            Nagra_Socket.FMainFrm = aParent
            SendCmdInfo.GetInfoLst.Clear()
            
            If ThreadInitial() Then
                Try
                    If tmrNagraStatus Is Nothing Then
                        tmrNagraStatus = New Threading.Timer(AddressOf chkStatus, Nothing, 0, FCheckStatusTime * 1000)
                    Else
                        tmrNagraStatus.Change(0, 3000)
                    End If

                    ti = ThreadPool.UnsafeRegisterWaitForSingleObject(evn, _
                                                                New WaitOrTimerCallback(AddressOf ThreadProcess), _
                                                                FSOInfoList, TimeSpan.FromSeconds(FReadDataTime), False)
                    evn.Set()

                Catch ex As Exception
                    WriteErrTxtLog.WriteTxtError(ex, Nothing)

                End Try
            End If
        Catch ex As Exception
            If FMainForm Is Nothing Then

                XtraMessageBox.Show("主畫面初始化失敗！無法執行本系統", "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                XtraMessageBox.Show("初始化系統失敗！", "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If

        End Try

    End Sub
    Private Sub chkStatus(ByVal status As Object)
        If FNagraConnecting OrElse FThreadsTotal > 0 Then
            Exit Sub
        End If
        Try
            Interlocked.Increment(Nagra_Socket.FUseSocket)

            If Nagra_Socket.FUseSocket <= Nagra_Socket.FMaxSocket Then

                FNagraConnecting = True
                FNagraStatusOK = Nagra_Socket.ConnectSocket
            End If


        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
            UpdateSysErrUI(FMainForm, FMainForm.TreLstSysErr, ex, "Socket主機連線失敗")
        Finally
            Interlocked.Decrement(Nagra_Socket.FUseSocket)
            FNagraConnecting = False

        End Try
    End Sub
    Private Function ThreadInitial() As Boolean
        Try
            If FMaxThread > (Environment.ProcessorCount * 25) Then
                FMaxThread = (Environment.ProcessorCount * 25)
            End If
            If FReadDataTime <= 0 Then
                FReadDataTime = 5
            End If

            ThreadPool.SetMaxThreads(FMaxThread, FMaxThread)
            ThreadPool.SetMinThreads(2, 2)
            If FReadDataTime <= 0 Then
                FReadDataTime = 30
            End If
            aSQLCmd = String.Format(aSQLCmd, FProcessNumber)
            FSOIndex.Clear()
            For i As Int32 = 0 To FSOInfoList.Count - 1
                FSOIndex.Add(FSOInfoList(i).CompID, FSOInfoList.Item(i))
            Next
            Return True
        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
        End Try


    End Function
    Private Sub ThreadProcess(ByVal obj As List(Of SOInfoClass), ByVal timedOut As Boolean)

        Try

            FRegOK = GetSystemInfo.IsRegister(ShowForm.No)
            

            If (Not FThreadStop) AndAlso (FRegOK) Then
                

                'If Nagra_Socket.ConnectSocket() Then


                For i As Int32 = 0 To obj.Count - 1
                    ThreadPool.QueueUserWorkItem(AddressOf ThreadCmd, obj.Item(i))
                Next
                'End If
            Else
            If FThreadsTotal <= 0 Then
                Dim blnAllLoad As Boolean = True
                If FRegOK Then
                    For i As Int32 = 0 To FSOInfoList.Count - 1
                        If FSOInfoList(i).FirstLoad Then
                            blnAllLoad = False
                            Exit For
                        End If
                    Next
                Else
                    blnAllLoad = True
                End If

                If (evn IsNot Nothing) AndAlso (blnAllLoad) Then
                    evn.Reset()
                    ti.Unregister(evn)
                    For i As Int32 = 0 To FSOInfoList.Count - 1
                        FSOInfoList(i).ConnectionOK = False
                        FSOInfoList(i).ProcessingNumber = 0
                        FSOInfoList(i).WaitProcessNumber = 0
                        FSOInfoList(i).FirstLoad = True
                        If FSOInfoList(i).OracleCn IsNot Nothing _
                            AndAlso FSOInfoList(i).OracleCn.State = ConnectionState.Open Then
                            FSOInfoList(i).OracleCn.Close()
                        End If
                    Next
                    DisableControl(FMainForm, RunStatus.StopCmd)
                    If FRegOK Then
                        UpdateSOStatusUI(FMainForm, FMainForm.treLstSO, Nothing, SOStatus.Initial)
                    Else
                        UpdateSOStatusUI(FMainForm, FMainForm.treLstSO, Nothing, SOStatus.Key)
                    End If

                End If
            End If
            End If
        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
            XtraMessageBox.Show("停止失敗！" & Environment.NewLine & ex.Message, "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub
    Private Sub ThreadCmd(ByVal obj As SOInfoClass)
        Try

            If (FThreadStop) OrElse (obj.ProcessingNumber > 0) Then
                Exit Sub
            End If

            SyncLock obj
                Dim aSytle As SOStatus
                Dim aId As String = Guid.NewGuid.ToString
                UpdateSODatabase(FMainForm, aId, obj, SOStatus.NotUpdateImg)
                '判斷是否連線
                If Not IsConnectOK(obj.OracleCn) Then
                    aSytle = SOStatus.NO
                    If (Not obj.FirstLoad) AndAlso (obj.ConnectionOK) Then
                        WriteErrTxtLog.WriteSOStatus("公司 : " & obj.CompName & " ---> 連線後已斷線")
                    End If
                    obj.ConnectionOK = False

                Else
                    aSytle = SOStatus.Yes
                    If (Not obj.FirstLoad) AndAlso (Not obj.ConnectionOK) Then
                        WriteErrTxtLog.WriteSOStatus("公司 : " & obj.CompName & " ---> 斷線後已連線")
                    End If
                    obj.ConnectionOK = True
                End If
                If obj.FirstLoad Then
                    WriteErrTxtLog.WriteSOStatus("公司 : " & obj.CompName & _
                                                 IIf(obj.ConnectionOK, " ---> 首次連線成功", " ---> 首次連線失敗"))
                End If
                obj.FirstLoad = False
                UpdateSODatabase(FMainForm, aId, obj, aSytle)
                UpdateSOStatusUI(FMainForm, FMainForm.treLstSO, obj, aSytle)
                If (obj.WaitProcessNumber <= 0) AndAlso (obj.ConnectionOK) AndAlso (FNagraStatusOK) Then
                    ProcessCMD(obj)
                End If
            End SyncLock

        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, obj.CompID, Nothing, Nothing, Nothing, Nothing)
        Finally
            'obj = Nothing
        End Try
    End Sub
    '開始處理資料
    Private Sub ProcessCMD(ByVal obj As SOInfoClass)

        Try
            SyncLock obj
                Interlocked.Increment(FThreadsTotal)
                If (obj.OracleCn.State = ConnectionState.Closed) Then
                    Try
                        obj.OracleCn.Open()
                    Catch ex As Exception
                        UpdateSOStatusUI(FMainForm, FMainForm.treLstSO, obj, SOStatus.NO)
                        obj.ConnectionOK = False
                        Exit Sub
                    End Try

                    Using cmd As OracleCommand = obj.OracleCn.CreateCommand
                        cmd.CommandText = aSQLCmd
                        Using r As OracleDataReader = cmd.ExecuteReader(CommandBehavior.CloseConnection)
                            While r.Read

                                SendCmdInfo.ADD(obj.CompID, obj.CompName, r("SEQNO"), r("CMD"), obj.OraConnectString)
                                Try
                                    SyncLock obj.WaitProcessNumber.GetType
                                        obj.WaitProcessNumber += 1
                                    End SyncLock
                                    SyncLock obj.ProcessingNumber.GetType
                                        obj.ProcessingNumber += 1
                                    End SyncLock
                                    
                                    UpdateSOStatusUI(FMainForm, FMainForm.treLstSO, obj, SOStatus.NotUpdateImg)
                                    Nagra_Socket.SendCMD()
                                Catch ex As Exception
                                    SyncLock obj.WaitProcessNumber.GetType
                                        obj.WaitProcessNumber -= 1
                                    End SyncLock
                                    SyncLock obj.ProcessingNumber.GetType
                                        obj.ProcessingNumber -= 1
                                    End SyncLock
                                End Try

                                'ThreadPool.QueueUserWorkItem(AddressOf UpdateProcessCmd, New String() {r.Item("SeqNo"), IIf(IsDBNull(r.Item("CMD")), Nothing, r.Item("Cmd")), obj.CompID})

                            End While

                        End Using
                    End Using
                End If
            End SyncLock
        Catch ex As Exception
            If Not IsConnectOK(obj.OracleCn) Then
                If obj.ConnectionOK Then
                    UpdateSOStatusUI(FMainForm, FMainForm.treLstSO, obj, SOStatus.NO)
                    'obj.ConnectionOK = False
                End If
            End If
            UpdateSysErrUI(FMainForm, FMainForm.TreLstSysErr, ex, "公司別 : " & obj.CompID.ToString)
            WriteErrTxtLog.WriteTxtError(ex, obj.CompID, Nothing, Nothing, "SQL : " & aSQLCmd, Nothing)
        Finally
            If obj.OracleCn.State = ConnectionState.Open Then
                obj.OracleCn.Close()
            End If
            Interlocked.Decrement(FThreadsTotal)
            obj = Nothing
        End Try



    End Sub
    Private Sub UpdateProcessCmd(ByVal obj As Array)

        Dim aSO As SOInfoClass = FSOIndex(obj.GetValue(2).ToString)
        Dim aSEQNo As String = obj.GetValue(0)
        Dim aCMD As String = obj.GetValue(1)
        Try
            Interlocked.Increment(FThreadsTotal)
            SyncLock aSO.WaitProcessNumber.GetType
                aSO.WaitProcessNumber += 1
            End SyncLock
            SyncLock aSO.ProcessingNumber.GetType
                aSO.ProcessingNumber += 1
            End SyncLock
            Using cn As New OracleConnection(aSO.OraConnectString)
                Try
                    cn.Open()
                Catch ex As Exception
                    UpdateSOStatusUI(FMainForm, FMainForm.treLstSO, aSO, SOStatus.NO)
                    Exit Sub
                End Try
                Using cmd As OracleCommand = cn.CreateCommand
                    cmd.CommandText = String.Format("UPDATE SO555 SET CMDSTATUS='G',DefaultUser=Nvl(DefaultUser,0)+1 WHERE SEQNO={0}", aSEQNo)
                    cmd.ExecuteNonQuery()

                    If cn IsNot Nothing AndAlso cn.State = ConnectionState.Open Then
                        cn.Close()
                    End If
                End Using

            End Using

        Catch ex As Exception
            If Not IsConnectOK(aSO.OracleCn) Then
                If aSO.ConnectionOK Then
                    UpdateSOStatusUI(FMainForm, FMainForm.treLstSO, aSO, SOStatus.NO)
                End If
            Else
                UpdateSOStatusUI(FMainForm, FMainForm.treLstSO, aSO, SOStatus.Yes)
            End If
            UpdateSysErrUI(FMainForm, FMainForm.TreLstSysErr, ex, "TEST")
            WriteErrTxtLog.WriteTxtError(ex, aSO.CompID, aSEQNo, Nothing, Nothing, Nothing)
        Finally
            SyncLock aSO.ProcessingNumber.GetType
                aSO.ProcessingNumber -= 1
            End SyncLock
            SyncLock aSO.WaitProcessNumber.GetType
                aSO.WaitProcessNumber -= 1
            End SyncLock
            Interlocked.Decrement(FThreadsTotal)
            UpdateSOStatusUI(FMainForm, FMainForm.treLstSO, aSO, SOStatus.NotUpdateImg)
            Erase obj
            obj = Nothing
            aSO = Nothing
            aSEQNo = Nothing
            aCMD = Nothing
        End Try
    End Sub



    'Private Sub UpdateProcessCmd(ByVal obj As String)

    '    Dim aSO As SOInfoClass = FSOIndex(obj.Split(Environment.NewLine).GetValue(1).ToString.Trim())
    '    Dim aSEQNo As String = obj.Split(Environment.NewLine).GetValue(0).ToString.Trim()
    '    'If Not FSOIndex.ContainsKey(aSO.CompID.ToString()) Then
    '    '    Exit Sub
    '    'End If

    '    Try
    '        Interlocked.Increment(FThreadsTotal)
    '        Using cn As New OracleConnection(aSO.OraConnectString)
    '            Try
    '                cn.Open()
    '                If Not aSO.ConnectionOK Then
    '                    UpdateSOStatusUI(FMainForm, FMainForm.treLstSO, aSO, SOStatus.Yes)
    '                End If
    '            Catch ex As Exception
    '                If aSO.ConnectionOK Then
    '                    UpdateSOStatusUI(FMainForm, FMainForm.treLstSO, aSO, SOStatus.NO)
    '                End If
    '                Exit Sub
    '            End Try
    '            Using cmd As OracleCommand = cn.CreateCommand
    '                cmd.CommandText = String.Format("UPDATE SO555 SET CMDSTATUS='G',DefaultUser=Nvl(DefaultUser,0)+1 WHERE SEQNO={0}", aSEQNo)
    '                cmd.ExecuteNonQuery()
    '                SyncLock aSO.ProcessingNumber.GetType
    '                    aSO.ProcessingNumber -= 1
    '                End SyncLock
    '                SyncLock aSO.WaitProcessNumber.GetType
    '                    aSO.WaitProcessNumber -= 1
    '                End SyncLock
    '                UpdateSOStatusUI(FMainForm, FMainForm.treLstSO, aSO, SOStatus.NotUpdateImg)
    '                If cn IsNot Nothing AndAlso cn.State = ConnectionState.Open Then
    '                    cn.Close()
    '                End If
    '            End Using

    '        End Using

    '    Catch ex As Exception


    '        If Not IsConnectOK(aSO.OracleCn) Then
    '            If aSO.ConnectionOK Then
    '                UpdateSOStatusUI(FMainForm, FMainForm.treLstSO, aSO, SOStatus.NO)
    '            End If
    '        Else
    '            UpdateSOStatusUI(FMainForm, FMainForm.treLstSO, aSO, SOStatus.Yes)
    '        End If
    '        UpdateSysErrUI(FMainForm, FMainForm.TreLstSysErr, ex, "TEST")
    '        WriteErrTxtLog.WriteTxtError(ex, aSO.CompID, aSEQNo, Nothing, Nothing, Nothing)
    '    Finally
    '        SyncLock aSO.ProcessingNumber.GetType
    '            aSO.ProcessingNumber -= 1
    '        End SyncLock
    '        SyncLock aSO.WaitProcessNumber.GetType
    '            aSO.WaitProcessNumber -= 1
    '        End SyncLock
    '        Interlocked.Decrement(FThreadsTotal)
    '        aSO = Nothing
    '        aSEQNo = Nothing
    '    End Try
    'End Sub
End Module
