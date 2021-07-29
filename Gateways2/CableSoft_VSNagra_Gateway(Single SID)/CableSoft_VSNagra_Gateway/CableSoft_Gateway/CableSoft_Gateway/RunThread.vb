Imports System
Imports System.Windows
Imports System.Windows.Forms
Imports System.IO
Imports System.Text
Imports System.Data
Imports System.Data.OracleClient
Imports System.Threading
Imports CableSoft.GateWay.Common
Imports CableSoft.GateWay.csException
Imports CableSoft.GateWay.SystemIO
Imports DevExpress.XtraEditors
Imports CableSoft_KeyPro
Imports System.Net.Sockets
Imports System.Collections
Imports System.Collections.Generic
Imports CableSoft_Nagra_BuilderCmd
Imports Raize.CodeSiteLogging
Imports System.Globalization
Imports System.Net
Imports System.Net.Mail
Module RunThread
    Private ti As RegisteredWaitHandle = Nothing
    Public evn As New AutoResetEvent(False)
    'Private FMainForm As rfrmMain
    Private FThreadsTotal As Int32 = 0
    Public FSOIndex As New System.Collections.Generic.Dictionary(Of String, SOInfoClass)
    Public FUseCompIndex As New Dictionary(Of String, UseCompanyClass)
    'Private Const FSQLCMD As String = "SELECT * FROM ( SELECT * FROM SO560 ORDER BY NVL(BATCH,0), SEQNO ) A WHERE ( A.CMDSTATUS = 'W' )  " & _
    '                    " AND ROWNUM<={0}  AND A.COMPCODE IN ( {1} )  " & _
    '                    " AND DECODE(A.RESVTIME,NULL,TO_DATE('10000101000000','YYYYMMDDHH24MISS'),A.RESVTIME) <= SYSDATE " & _
    '                    " ORDER BY Nvl(A.BATCH,0), A.SEQNO "
    Private Const FSQLCMD As String = "SELECT * FROM SEND_NAGRA A WHERE A.CMD_STATUS='W' " & _
        " AND ROWNUM <= {0} AND  A.GATEWAYTYPE=1 AND COMPCODE IN ({1}) " & _
        " ORDER BY A.SEQNO "



    Private FExeSqlCMD As String = String.Empty
    Public rwl As New ReaderWriterLockSlim(LockRecursionPolicy.NoRecursion)
    Public FNagraStatusOK As Boolean = False
    'Public FNagraConnecting As Boolean = False
    Public tmrNagraStatus As Threading.Timer = Nothing
    Public tmrSendCmd As Threading.Timer = Nothing
    'Public sendDone As New ManualResetEvent(False)
    'Public receiveDone As New ManualResetEvent(False)
    Public fSocket() As Socket
    Public fNagraCommand() As NagraCommand

    Public fFirstChkNagraStatus As Boolean
    Public fNagraSnd1002 As Boolean
    'Public fHaveCount As Int32
    Private Const SplitReturnChar As Char = ","c
    Private Const SplitNoteChar As Char = ","c
    'Private fRunTime As Int32 = 0
    Private FWatch As New Stopwatch
    Private Const fGCTime As Int32 = 30

    Public Sub RunGateway()
        Try

            'FMainForm = aParent
            'Nagra_Socket.FMainFrm = aParent
            'SendCmdInfo.GetInfoLst.Clear()

            If ThreadInitial() Then
                Try
                    FWatch.Start()
                    If tmrNagraStatus Is Nothing Then
                        tmrNagraStatus = New Threading.Timer(AddressOf chkStatus, Nothing, 0, FCheckStatusTime * 1000)
                    Else
                        tmrNagraStatus.Change(0, FCheckStatusTime * 1000)
                    End If

                    'ti = ThreadPool.RegisterWaitForSingleObject(evn, _
                    '                                            New WaitOrTimerCallback(AddressOf ThreadProcess), _
                    '                                            FSOInfoList, TimeSpan.FromSeconds(FReadDataTime), False)
                    'evn.Set()
                    If tmrSendCmd Is Nothing Then
                        tmrSendCmd = New Threading.Timer(AddressOf ThreadProcess, FSOInfoList, 0, FReadDataTime * 1000)
                    Else
                        tmrSendCmd.Change(0, FReadDataTime * 1000)

                    End If

                Catch ex As Exception
                    WriteErrTxtLog.WriteTxtError(ex, Nothing)
                End Try
            End If
        Catch ex As Exception
            XtraMessageBox.Show("初始化系統失敗！", "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
            'If FMainForm Is Nothing Then

            '    XtraMessageBox.Show("主畫面初始化失敗！無法執行本系統", "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
            'Else
            '    XtraMessageBox.Show("初始化系統失敗！", "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
            'End If

        End Try

    End Sub
    '連接Socket 並送出 Cmd1002命令
    Private Function ConnectSocket(ByRef aSocket As Socket, ByVal aSocketNum As Int32) As Boolean

        Dim aMsgStyle As New ThreadLocal(Of MsgStyle)
        aMsgStyle.Value = MsgStyle.ErrorLa
        Dim aDisconnect As New ThreadLocal(Of Boolean)
        Dim aSend As New ThreadLocal(Of List(Of Byte()))
        aDisconnect.Value = False
        Try
            If aSocket Is Nothing Then
                aDisconnect.Value = True
                aSocket = New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)
                aSocket.SendTimeout = FSndTimeout * 1000
                aSocket.ReceiveTimeout = FRcvTimeout * 1000
            End If

            If Not aSocket.Connected Then
                aDisconnect.Value = True
                Try
                    aSocket.Shutdown(SocketShutdown.Both)
                    aSocket.Disconnect(False)
                    aSocket = Nothing
                    aSocket = New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)
                    aSocket.SendTimeout = FSndTimeout
                    aSocket.ReceiveTimeout = FRcvTimeout
                Catch ex As Exception

                End Try
                FLstSysMsg.Clear()
                FLstSysMsg.Add(String.Format("Parent{0} > Nagra主機連接中..", Format(Now, "yyyy/MM/dd HH:mm:ss")))
                aSocket.Connect(FNagraIPAddress, FNagraPort)


                '取得連入Nagra後要與Nagra Talk的資訊
                aSend.Value = fNagraCommand(aSocketNum).BuilderCmd(CmdType.ConfirmConnect, Nothing)
                SyncLock aSocket    '將斷線的Socket Lock起來,避免送連線時又送命令
                    If aSend IsNot Nothing AndAlso aSend.Value.Count = 1 Then
                        Dim aBty As New ThreadLocal(Of Byte())
                        aBty.Value = aSend.Value.Item(0)
                        aSocket.Send(aBty.Value)
                        Erase aBty.Value
                        aBty.Dispose()
                    Else
                        FLstSysMsg.Add("主機連接失敗 -- 產生資料失敗")
                        UpdateSysErrUI(New Exception("Socket 連接失敗"), "產生資料失敗")

                        aSocket.Shutdown(SocketShutdown.Both)
                        aSocket.Disconnect(False)
                        aSend.Dispose()
                        Exit Function
                    End If

                    Dim aRcv(6) As Byte
                    '一定要Sleep 100秒以上，不然第一次都會出現失敗,要第二次才會出現成功
                    Thread.Sleep(300)
                    If aSocket.Receive(aRcv) > 0 Then

                        Dim aStatus As New ThreadLocal(Of NagraStatus_Type)
                        aStatus.Value = fNagraCommand(aSocketNum).NagraStatus(aRcv, CmdType.ConfirmConnect)
                        If (aStatus.Value.answer_code <> 0) OrElse (aStatus.Value.SUCCESS <> 6) Then

                            Try
                                FLstSysMsg.Add("主機連接失敗 > SourceId : " & fNagraCommand(aSocketNum).SourceId)
                                UpdateSysErrUI(New Exception("Socket 連接失敗> SourceId:" & fNagraCommand(aSocketNum).SourceId), " Connect Status : " & _
                                               aStatus.Value.SUCCESS & " Answer_Code : " & aStatus.Value.answer_code)
                                aSocket.Shutdown(SocketShutdown.Both)
                                aSocket.Disconnect(False)
                            Catch ex As Exception
                            Finally
                                aSocket = Nothing
                                aStatus.Dispose()
                            End Try
                            Return False
                            'aSend = String.Empty
                            'aSend = NagraCommand.BuilderCmd(CmdType.Cmd1002)
                            'aBty = Encoding.ASCII.GetBytes(aSend)
                            'FSocket.Send(aBty)
                            ''Y00000000205123420005678920101008100120000000110003000003220000000105200012345678920101108
                            ''E000000003051234200056789201010081000200000001000000000000000000000000
                            'Dim aRcv2(1023) As Byte
                            'If FSocket.Receive(aRcv2) > 0 Then
                            '    Dim aRcvString = Encoding.ASCII.GetString(aRcv2, 1, aRcv2.Length - 1).TrimEnd(Chr(0))
                            'End If
                        Else
                            '       FLstSysMsg.Add("主機連接成功 > SourceId : " & fNagraCommand(aSocketNum).SourceId)
                        End If

                    End If
                End SyncLock
                aSend.Dispose()
                'UpdateNagraMsgUI(FMainForm, FMainForm.TreLstSysMsg, MsgStyle.OKLa)
            End If
            If aSocket IsNot Nothing AndAlso aSocket.Connected Then
                SyncLock aSocket    '將Socket Lock起來,避免有命令在送
                    If SendCmd1002(aSocket, aSocketNum) Then
                        aMsgStyle.Value = MsgStyle.OKLa
                        If aDisconnect.Value Then
                            FLstSysMsg.Add("主機連接成功 > SourceId : " & fNagraCommand(aSocketNum).SourceId)
                        Else
                            FLstSysMsg.Clear()
                        End If

                        Return True
                    Else
                        'FLstSysMsg.Add("主機連接失敗 > SourceId : " & fNagraCommand(aSocketNum).SourceId)
                        aMsgStyle.Value = MsgStyle.ErrorLa
                        If aDisconnect.Value Then
                            FLstSysMsg.Add("主機連接失敗 > SourceId : " & fNagraCommand(aSocketNum).SourceId)
                        Else
                            FLstSysMsg.Clear()
                            'FLstSysMsg.Add(String.Format("Parent{0} > 送出1002命令..", Format(Now, "yyyy/MM/dd hh:mm:ss")))
                            'FLstSysMsg.Add("1002 命令失敗 > SourceId : " & fNagraCommand(aSocketNum).SourceId)
                        End If
                        Return False
                    End If
                End SyncLock
            End If

        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
            FLstSysMsg.Add("Nagra主機連接失敗 > SourceId : " & fNagraCommand(aSocketNum).SourceId)
            UpdateSysErrUI(ex, " 連接 Nagra Socket 失敗 > SourceId : " & fNagraCommand(aSocketNum).SourceId)


        Finally

            UpdateNagraMsgUI(aMsgStyle.Value)
            If fIsUseMail Then
                If (aSocket Is Nothing) OrElse (Not aSocket.Connected) Then
                    If fNagraCommand(aSocketNum).ConnectErrCnt < Int32.Parse(fSourceIdErrCnt) Then
                        fNagraCommand(aSocketNum).ConnectErrCnt += 1
                    End If
                    If fNagraCommand(aSocketNum).ConnectErrCnt >= Int32.Parse(fSourceIdErrCnt) Then
                        If Not fNagraCommand(aSocketNum).SndConnectErrMail Then
                            fNagraCommand(aSocketNum).SndConnectErrMail = True
                            csCommon.SendMail(MDBMailPara.GetDefaultMDBName, _
                                  fMDBMailPassword, "SourceIDErr", fMailInfoTableName, _
                            fMailToTableName, fMailCCTableName, String.Empty, _
                            String.Empty, False, True)


                        End If
                    End If
                Else
                    fNagraCommand(aSocketNum).ConnectErrCnt = 0
                    fNagraCommand(aSocketNum).SndConnectErrMail = False
                End If
            End If
            aMsgStyle.Dispose()
            aDisconnect.Dispose()
            aSend.Dispose()
        End Try

    End Function
    Public Sub SndCmdErrMail(ByVal aSubject As String, ByVal aBody As String)

    End Sub
    Public Sub SndSourceIdMail(ByVal aSocketNum As Int32, _
                       ByVal aSubject As String, _
                       ByVal aBody As String)

    End Sub
    'Public Sub SndMailThread()
    '    Try

    '        csCommon.SendMail(MDBMailPara.GetDefaultMDBName, _
    '                              fMDBMailPassword, "SourceIDErr", fMailInfoTableName, _
    '                              fMailToTableName, fMailCCTableName, String.Empty, String.Empty)
    '        WriteErrTxtLog.WriteSndEmailLog("成功寄出Mail")
    '        'Catch ex As SmtpException
    '        '    WriteErrTxtLog.WriteSndEmailLog("錯誤代碼：" & ex.StatusCode.ToString & "，錯誤內容：" & ex.Message)
    '    Catch ex As Exception
    '        WriteErrTxtLog.WriteSndEmailLog(ex.Message)
    '    End Try

    'End Sub

    '送出CMD1002 訊息
    Public Function SendCmd1002(ByRef aSocket As Socket, ByVal aSocketNum As Int32) As Boolean

        Dim objRet As New ThreadLocal(Of NagraCmdStatus_Type)
        Dim aSndLst As New ThreadLocal(Of List(Of Byte()))
        Try


            aSndLst.Value = fNagraCommand(aSocketNum).BuilderCmd(CmdType.Cmd1002, Nothing)
            If (aSndLst.Value Is Nothing) OrElse (aSndLst.Value.Count <> 1) Then
                UpdateSysErrUI(New Exception("傳送資料產生有誤"), " 1002命令 失敗")
                WriteErrTxtLog.WriteTxtError(New Exception("傳送資料產生有誤"), Nothing, Nothing)
            End If
            Dim aSndbty() As Byte = aSndLst.Value.Item(0)
            '要送之前再判斷一次是否有命令在送了,如果有不要送1002
            If FThreadsTotal > 0 Then
                Return True
            End If
            aSocket.Send(aSndbty)

            Dim aRcvBty(511) As Byte
            Thread.Sleep(200)
            If aSocket.Receive(aRcvBty) > 0 Then
                'CodeSite.SendCollection("Recv", aRcvBty)


                objRet.Value = fNagraCommand(aSocketNum).NagraStatus(aRcvBty, CmdType.Cmd1002)
                If objRet.Value.Status = AckCode.Ack Then
                    Return True
                Else
                    UpdateSysErrUI(New Exception("1002 命令有錯 > SourceId : " & fNagraCommand(aSocketNum).SourceId), _
                                   "ErrorCode : " & objRet.Value.ErrorCode & _
                                   " ErrorName : " & objRet.Value.ErrorName)
                    Return False
                End If
            End If
            Erase aSndbty
            Erase aRcvBty
            Return True

        Catch ex As Exception
            UpdateSysErrUI(ex, " 1002命令 失敗")
            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
        Finally
            '記錄所有CMD1002
            'If objRet.Value IsNot Nothing Then
            '    Try
            '        CableSoft.GateWay.Common.csCommon.WriteLog(String.Format(" Source Id > {0} [ " & _
            '                                                                 "1002 命令 回應 > {1} , " & _
            '                                                                "錯誤代碼 > {2} , 錯誤名稱  > {3} ]", _
            '                                                                fNagraCommand(aSocketNum).SourceId, _
            '                                                                 objRet.Value.Status, _
            '                                                                objRet.Value.ErrorCode, _
            '                                                                objRet.Value.ErrorName))
            '    Finally

            '    End Try
            'End If
            objRet.Dispose()
            aSndLst.Value.Clear()
            aSndLst.Dispose()
            'Erase aSndbty
            'Erase aRcvBty
            'FSocket.Shutdown(SocketShutdown.Receive)
        End Try

    End Function
    '檢查連線並測試CMD1002
    Private Sub chkStatus(ByVal status As Object)

        Try
            If fNagraSnd1002 OrElse FThreadStop Then
                Exit Sub
            End If
            fNagraSnd1002 = True
            If FNagraHaveProcing > 0 Then
                Exit Sub
            End If
            For i As Int32 = 0 To fSocket.Length - 1
                'CodeSite.Send("ChkStatus")
                Thread.Sleep(100)
                '如果Socket不在連線狀態或是還未建立則直接連線
                If Not fFirstChkNagraStatus Then
                    Try
                        fSocket(i).Shutdown(SocketShutdown.Both)
                    Catch ex As Exception
                    Finally
                        If fSocket(i) IsNot Nothing Then
                            fSocket(i).Dispose()
                        End If

                        fSocket(i) = Nothing
                    End Try
                End If
                If (fSocket(i) Is Nothing) OrElse (Not fSocket(i).Connected) Then
                    If Not ConnectSocket(fSocket(i), i) Then
                        Try
                            fSocket(i).Shutdown(SocketShutdown.Both)
                            fSocket(i).Disconnect(False)
                        Catch ex As Exception
                        Finally
                            If fSocket(i) IsNot Nothing Then
                                fSocket(i).Dispose()
                            End If

                            fSocket(i) = Nothing
                        End Try
                    End If
                Else
                    Dim aCanCheck As Boolean = True
                    '檢查是否有命令還沒完成
                    For aIndex As Int32 = 0 To FSOInfoList.Count - 1
                        If FSOInfoList.Item(aIndex).WaitProcessNumber > 0 Then
                            aCanCheck = False
                        End If
                    Next
                    '如果沒有任何命令正在執行,則直接連線
                    If aCanCheck AndAlso FThreadsTotal <= 0 Then
                        If (Not ConnectSocket(fSocket(i), i)) Then
                            Try
                                fSocket(i).Shutdown(SocketShutdown.Both)
                                fSocket(i).Disconnect(False)
                            Catch ex As Exception
                            Finally
                                If fSocket(i) IsNot Nothing Then
                                    fSocket(i).Dispose()
                                End If
                                fSocket(i) = Nothing
                            End Try
                        End If
                    End If
                End If
            Next
        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
            UpdateSysErrUI(ex, "Socket主機連線失敗")
        Finally
            fFirstChkNagraStatus = True
            fNagraSnd1002 = False
        End Try
    End Sub
    Private Function ThreadInitial() As Boolean
        Try
            FThreadsTotal = 0
            'fRunTime = 0
            If FMaxThread > (Environment.ProcessorCount * 250) Then
                FMaxThread = (Environment.ProcessorCount * 250)
            End If
            If FReadDataTime <= 0 Then
                FReadDataTime = 5
            End If

            ThreadPool.SetMaxThreads(FMaxThread, FMaxThread)
            ThreadPool.SetMinThreads(10, 10)
            If FNagraMaxProcing < 1000 Then
                FNagraMaxProcing = 1000
            End If

            If FReadDataTime <= 0 Then
                FReadDataTime = 30
            End If

            FExeSqlCMD = String.Format(FSQLCMD, FProcessNumber, FUseCompWhere)
            FSOIndex.Clear()
            For i As Int32 = 0 To FSOInfoList.Count - 1
                FSOIndex.Add(FSOInfoList(i).CompID, FSOInfoList.Item(i))
            Next

            FUseCompIndex.Clear()
            For i As Int32 = 0 To FUseCompList.Count - 1
                FUseCompIndex.Add(FUseCompList(i).CompID, FUseCompList(i))
            Next


            ReDim fSocket(FSndSourceId.Length - 1)
            ReDim fNagraCommand(FSndSourceId.Length - 1)
            NagraCommand.IniPath = CableSoft.GateWay.Common.csCommon.GetAppPath
            NagraCommand.ErrCodeLst = FErrCodeLst
            '將Soruce Id 分配給不同的 Socket
            For i As Int32 = 0 To fNagraCommand.Length - 1
                fNagraCommand(i) = New NagraCommand(Right("0000" & FSndSourceId(i), 4), _
                                                            Right("0000" & FSndDestId(i), 4), _
                                                            Right("00000" & FSndMopPPId(i), 5))
                fNagraCommand(i).SndConnectErrMail = False
                fNagraCommand(i).ConnectErrCnt = 0
            Next
            fFirstChkNagraStatus = False
            SendCmdInfo.GetInfoLst.Clear()
            Return True
        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
        End Try


    End Function
    Private Sub ThreadProcess(ByVal obj As Object)
        'ByVal obj As List(Of SOInfoClass), ByVal timedOut As Boolean
        Dim aCanExit As New ThreadLocal(Of Boolean)
        Try

            'FRegOK = GetSystemInfo.IsRegister(ShowForm.No)
            FRegOK = True            
            If Date.Now.Date > Date.Parse(FRegDate, CultureInfo.CreateSpecificCulture("zh-TW")).Date Then
                FRegOK = False
            End If

            If (Not FThreadStop) AndAlso (FRegOK) Then

                For i As Int32 = 0 To obj.Count - 1
                    ThreadPool.QueueUserWorkItem(AddressOf ThreadCmd, obj.Item(i))
                Next

            Else

                If FThreadsTotal <= 0 Then

                    aCanExit.Value = True
                    If FThreadsTotal > 0 Then
                        aCanExit.Value = False
                    End If

                    If (evn IsNot Nothing) AndAlso (aCanExit.Value) Then
                        'evn.Reset()
                        'ti.Unregister(evn)

                        tmrSendCmd.Change(Timeout.Infinite, Timeout.Infinite)
                        tmrSendCmd.Dispose()
                        tmrSendCmd = Nothing

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

                        If FRegOK Then

                            UpdateSOStatusUI(Nothing, SOStatus.Initial)
                        Else
                            UpdateSOStatusUI(Nothing, SOStatus.Key)
                        End If
                        DisableControl(RunStatus.StopCmd)
                        '將所有的Socket 釋放掉
                        If fSocket IsNot Nothing Then
                            For i As Int32 = 0 To fSocket.Length - 1
                                If fSocket(i) IsNot Nothing Then
                                    Try
                                        fSocket(i).Shutdown(SocketShutdown.Both)
                                        fSocket(i).Disconnect(False)
                                    Finally
                                        If fSocket(i) IsNot Nothing Then
                                            fSocket(i).Dispose()
                                        End If

                                        fSocket(i) = Nothing
                                    End Try
                                End If
                            Next
                        End If

                    Else
                        If aCanExit.Value Then
                            If fSocket IsNot Nothing Then
                                For i As Int32 = 0 To fSocket.Length - 1
                                    If fSocket(i) IsNot Nothing Then
                                        Try
                                            fSocket(i).Shutdown(SocketShutdown.Both)
                                            fSocket(i).Disconnect(False)
                                        Finally
                                            If fSocket(i) IsNot Nothing Then
                                                fSocket(i).Dispose()
                                            End If

                                            fSocket(i) = Nothing
                                        End Try
                                    End If

                                Next
                            End If
                            DisableControl(RunStatus.StopCmd)
                        End If
                    End If
                    If aCanExit IsNot Nothing Then
                        aCanExit.Dispose()
                    End If

                End If
            End If
        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
        Finally
            If aCanExit IsNot Nothing Then
                aCanExit.Dispose()
            End If
            obj = Nothing
            'obj.Clear()

            'XtraMessageBox.Show("停止失敗！" & Environment.NewLine & ex.Message, "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub
    Private Sub ThreadCmd(ByVal obj As SOInfoClass)
        Dim aSytle As New ThreadLocal(Of SOStatus)
        Dim aId As New ThreadLocal(Of String)
        Try
            
            If (FThreadStop) OrElse (obj.ProcessingNumber > 0) OrElse (Not fFirstChkNagraStatus) Then
                Exit Sub
            Else
                If (FWatch.Elapsed.TotalMinutes >= fGCTime) AndAlso (FThreadsTotal <= 0) Then
                    Interlocked.Increment(FThreadsTotal)
                    Try
                        If FThreadsTotal = 1 Then
                            FWatch.Reset()
                            GC.Collect()
                            GC.WaitForPendingFinalizers()
                        Else
                            Exit Sub
                        End If
                    Finally
                        If FThreadsTotal > 0 Then
                            Interlocked.Decrement(FThreadsTotal)
                        End If
                        If FThreadsTotal <= 0 Then
                            FWatch.Restart()
                        End If

                    End Try
                End If

            End If

            SyncLock obj

                'Date.Today.ToLongDateString()
                'aId.Value = Guid.NewGuid.ToString
                aId.Value = Format(Date.Now, "yyyyMMddHHmmssff")
                UpdateSODatabase(aId.Value, obj, SOStatus.NotUpdateImg)

                '判斷是否連線
                If Not IsConnectOK(obj.OracleCn) Then
                    aSytle.Value = SOStatus.NO
                    If (Not obj.FirstLoad) AndAlso (obj.ConnectionOK) Then
                        WriteErrTxtLog.WriteSOStatus(obj.CompName & " ---> 連線後已斷線")
                    End If
                    obj.ConnectionOK = False

                Else
                    aSytle.Value = SOStatus.Yes
                    If (Not obj.FirstLoad) AndAlso (Not obj.ConnectionOK) Then
                        WriteErrTxtLog.WriteSOStatus(obj.CompName & " ---> 斷線後已連線")
                    End If
                    obj.ConnectionOK = True
                End If
                If obj.FirstLoad Then
                    WriteErrTxtLog.WriteSOStatus(obj.CompName & _
                                                 IIf(obj.ConnectionOK, " ---> 首次連線成功", " ---> 首次連線失敗"))
                End If
                obj.FirstLoad = False
                UpdateSODatabase(aId.Value, obj, aSytle.Value)
                UpdateSOStatusUI(obj, aSytle.Value)
                If (obj.WaitProcessNumber <= 0) AndAlso _
                    (obj.ConnectionOK) Then
                    ProcessCMD(obj)
                End If
            End SyncLock

        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, obj.CompID, Nothing, Nothing, Nothing, Nothing)
        Finally
            aSytle.Dispose()
            aId.Dispose()
            'obj = Nothing
           
        End Try
    End Sub
    Private Sub ReadOK()
        Dim aSocketNum As Int32 = 0
        Dim aThreadPoolOK As Boolean = False


        Try

            'evn.Set()
            Do While Not (SendCmdInfo.GetInfoLst.Count = 0)

                Interlocked.Increment(FThreadsTotal)    '使用的Thread加1,UPDUIANDDB會扣回來
                If (FNagraHaveProcing < FNagraMaxProcing) Then
                    '判斷Socket有連線就送資料
                    If (fSocket(aSocketNum) IsNot Nothing) AndAlso (fSocket(aSocketNum).Connected) Then
                        aThreadPoolOK = True    '其中一個成功就設成功,後面的命令可以繼續送
                        Dim objCollect As New List(Of Object)
                        objCollect.Add(aSocketNum)
                        objCollect.Add(SendCmdInfo.GetInfoLst(0))
                        '將處理的資料丟至ThreadPool
                        ThreadPool.QueueUserWorkItem(AddressOf SendNagraCmd, objCollect)
                        SendCmdInfo.GetInfoLst.RemoveAt(0)

                        'evn.Set()
                    End If

                    aSocketNum += 1
                    If aSocketNum >= fSocket.Length Then
                        aSocketNum = 0
                        '沒有任何Socket可以連線,就將資料設成失敗
                        'KC說Socket 全部斷掉不要往下送 By Kin 2011/03/23
                        If Not aThreadPoolOK Then
                            Try
                                SendCmdInfo.GetInfoLst(0).AllFieldsValue("Cmd_Status".ToUpper) = "E"
                                SendCmdInfo.GetInfoLst(0).AllFieldsValue("Err_Code".ToUpper) = "-9997"
                                SendCmdInfo.GetInfoLst(0).AllFieldsValue("Err_Msg".ToUpper) = "CableSoft_Socket_All_Disconnect"
                                ThreadPool.QueueUserWorkItem(AddressOf UpdUiAndDB, _
                                                                    SendCmdInfo.GetInfoLst(0))
                                'evn.Set()
                            Finally
                                SendCmdInfo.GetInfoLst.RemoveAt(0)
                            End Try
                        End If
                    End If
                End If

                If FThreadStop Then
                    SendCmdInfo.GetInfoLst.Clear()
                    Exit Do
                End If

            Loop

        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, "準備送出資料失敗 ! ")
            UpdateSysErrUI(ex, "準備送出資料失敗 ! ")
        Finally
            '清空所有資料
            SendCmdInfo.GetInfoLst.Clear()
            'objCollect.Value.Clear()
            'objCollect.Value = Nothing
            'objCollect.Dispose()
        End Try

    End Sub

    '開始處理資料
    Private Sub ProcessCMD(ByRef obj As SOInfoClass)

        Try

            If FThreadStop OrElse SendCmdInfo.GetInfoLst.Count > 0 Then
                Exit Sub
            End If
            If (obj.OracleCn.State = ConnectionState.Closed) Then
                Try
                    obj.OracleCn.Open()
                Catch ex As Exception
                    UpdateSOStatusUI(obj, SOStatus.NO)
                    obj.ConnectionOK = False
                    Exit Sub
                End Try


                Using cmd As OracleCommand = obj.OracleCn.CreateCommand
                    cmd.CommandText = FExeSqlCMD
                    Using r As OracleDataReader = cmd.ExecuteReader(CommandBehavior.CloseConnection)
                        SendCmdInfo.GetInfoLst.Clear()

                        While r.Read
                            Try
                                SyncLock obj.WaitProcessNumber.GetType
                                    obj.WaitProcessNumber += 1
                                    obj.UseCompId = Convert.ToInt32(r("CompCode"))
                                    FUseCompIndex.Item(r("CompCode")).WaitProcessNumber += 1
                                    FUseCompIndex.Item(r("CompCode")).ProcessingNumber += 1
                                End SyncLock
                                'UpdateSOStatusUI(FMainForm, FMainForm.treLstSO, obj, SOStatus.NotUpdateImg)
                                SendCmdInfo.ADD(obj.CompID, obj.CompName, r("SeqNo"), r("HIGH_LEVEL_CMD_ID"), obj.OraConnectString, r)
                            Catch ex As Exception
                                SyncLock obj.WaitProcessNumber.GetType
                                    If obj.WaitProcessNumber > 0 Then
                                        obj.WaitProcessNumber -= 1
                                    End If
                                End SyncLock

                            End Try
                        End While
                        'Socket 計數器
                        'Dim aSocketNum As Int32 = 0

                        UpdateSOStatusUI(obj, SOStatus.NotUpdateImg)

                        '準備送出資料給Nagra
                        If SendCmdInfo.GetInfoLst.Count > 0 Then
                            'ThreadPool.QueueUserWorkItem(AddressOf ReadOK)
                            ReadOK()
                            evn.Set()
                            'evn.WaitOne(TimeSpan.FromSeconds(1))
                        End If

                    End Using
                End Using

            End If

        Catch ex As Exception
            If Not IsConnectOK(obj.OracleCn) Then
                If obj.ConnectionOK Then
                    UpdateSOStatusUI(obj, SOStatus.NO)
                    'obj.ConnectionOK = False
                End If
            End If

            WriteErrTxtLog.WriteTxtError(ex, obj.CompID, Nothing, Nothing, "SQL : " & FExeSqlCMD, Nothing)
            UpdateSysErrUI(ex, "公司別 : " & obj.CompID.ToString)
        Finally
            If obj.OracleCn.State = ConnectionState.Open Then
                obj.OracleCn.Close()
                obj.OracleCn.Dispose()
            End If
            obj = Nothing


            'Interlocked.Decrement(FThreadsTotal)
            'System.GC.Collect()

        End Try



    End Sub
    Private Sub SendNagraCmd(ByVal aInfo As List(Of Object))
        Dim aSocketNum As New ThreadLocal(Of Int32)
        Dim aSoCmdInfo As New ThreadLocal(Of SOCMDInfo)
        Dim aTypeCmd As New ThreadLocal(Of CmdType)
        Try
            aSocketNum.Value = Int32.Parse(aInfo.Item(0))
            aSoCmdInfo.Value = CType(aInfo.Item(1), SOCMDInfo)
            '如果是處理斷線狀態，將資料恢復成W(待處理狀態,從新再傳)
            'If fSocket(aSocketNum.Value) Is Nothing Then
            '    aSoCmdInfo.Value.AllFieldsValue("Cmd_Status".ToUpper) = "W"
            '    aSoCmdInfo.Value.AllFieldsValue("ErrCode".ToUpper) = String.Empty
            '    aSoCmdInfo.Value.AllFieldsValue("ErrMsg".ToUpper) = String.Empty
            '    aSoCmdInfo.Value.Returnlicenses.Clear()
            '    aSoCmdInfo.Value.NagraSendAndRecvMsg.Clear()
            '    Exit Sub
            'End If
            SyncLock fSocket(aSocketNum.Value)
                Try
                    '進入到裡面將Send命令加1(有可能進不來,因為Socket有可能是Nothing,所以要加Try)
                    Interlocked.Increment(FNagraHaveProcing)

                    Select Case aSoCmdInfo.Value.CMD.ToUpper
                        Case "B1"
                            aTypeCmd.Value = CmdType.Cmd0851
                        Case "D1"
                            aTypeCmd.Value = CmdType.cmd0126
                        Case Else
                            aTypeCmd.Value = CmdType.OtherCMD
                            Exit Sub
                    End Select
                    '取出要送給Nagra的資訊
                    Dim aryMsg As New ThreadLocal(Of List(Of Byte()))
                    aSoCmdInfo.Value.AllFieldsValue("UPDTIME".ToUpper) = Date.Now
                    Try
                        aryMsg.Value = New List(Of Byte())
                        aryMsg.Value = fNagraCommand(aSocketNum.Value).BuilderCmd(aTypeCmd.Value, aSoCmdInfo.Value.AllFieldsValue)
                        If (aryMsg.Value Is Nothing) OrElse (aryMsg.Value.Count = 0) Then
                            aSoCmdInfo.Value.AllFieldsValue("Cmd_Status".ToUpper) = "E"
                            aSoCmdInfo.Value.AllFieldsValue("Err_Code".ToUpper) = Int32.Parse(CustomerNagraError.CableSoft_0851String_Error).ToString
                            aSoCmdInfo.Value.AllFieldsValue("Err_Msg".ToUpper) = [Enum].GetName(GetType(CustomerNagraError), CustomerNagraError.CableSoft_CmdString_Error)
                            aSoCmdInfo.Value.Returnlicenses.Clear()
                            aSoCmdInfo.Value.NagraSendAndRecvMsg.Clear()
                            UpdateSysErrUI(New Exception("命令產生資料失敗"), " [" & aSoCmdInfo.Value.CompName & "] 命令序號 [ " & aSoCmdInfo.Value.SEQNO & " ]")
                            WriteErrTxtLog.WriteTxtError(New Exception(aTypeCmd.Value & "產生資料失敗"), " [" & aSoCmdInfo.Value.CompName & "] 命令序號 [ " & aSoCmdInfo.Value.SEQNO & " ]", Nothing)
                            'aryMsg.Dispose()
                            Exit Sub
                        End If

                        '開始送出資料
                        For i As Int32 = 0 To aryMsg.Value.Count - 1
                            '送出資料
                            Try
                                If SendMsg(aSocketNum.Value, aSoCmdInfo.Value, aTypeCmd.Value, aryMsg.Value(i)) Then
                                    If Not RecvMsg(aSocketNum.Value, aSoCmdInfo.Value, aTypeCmd.Value, aSoCmdInfo.Value.NagraCommandType, aryMsg.Value(i)) Then
                                        Exit For
                                    Else
                                        '儲存取出的Licenses Key
                                        'If Not String.IsNullOrEmpty(aSoCmdInfo.Value.NagraCommandType.VUA) Then


                                        'End If

                                        'If Not String.IsNullOrEmpty(aSoCmdInfo.Value.NagraCommandType.Entity_data) Then
                                        '    aSoCmdInfo.Value.Returnlicenses.Add(i, aSoCmdInfo.Value.NagraCommandType.Entity_data)
                                        'End If
                                    End If
                                Else

                                    '如果傳送失敗,要將計數減一
                                    'Interlocked.Decrement(FNagraHaveProcing)
                                    aTypeCmd.Dispose()
                                    Exit For
                                End If
                            Finally

                                aSoCmdInfo.Value.NagraCommandType.CatRanNum = fNagraCommand(aSocketNum.Value).GetReturnString(aryMsg.Value(i), True).Substring(0, 9)

                            End Try
                            
                        Next
                    Finally
                        aryMsg.Value.Clear()
                        aryMsg.Dispose()
                    End Try
                Finally
                    '要將送出的命令減一個回來
                    If FNagraHaveProcing > 0 Then
                        Interlocked.Decrement(FNagraHaveProcing)
                    End If

                End Try
            End SyncLock
        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, "[命令 : " & aSoCmdInfo.Value.SEQNO & "失敗 ]")
        Finally
            ThreadPool.QueueUserWorkItem(AddressOf UpdUiAndDB, aSoCmdInfo.Value)

            aSocketNum.Dispose()
            aSoCmdInfo.Dispose()
            aInfo.Clear()
            aInfo = Nothing
            aTypeCmd.Dispose()

        End Try
    End Sub


    Public Function UpdSO561B(ByRef aCMDInFo As SOCMDInfo, ByRef objCn As OracleConnection, ByRef objTra As OracleTransaction) As Boolean

        Try
            If aCMDInFo.Returnlicenses Is Nothing OrElse aCMDInFo.Returnlicenses.Count <= 0 Then
                Return True
                Exit Function
            End If
            If objCn.State = ConnectionState.Closed Then
                objCn.Open()
            End If

            'Dim tra As New ThreadLocal(Of OracleTransaction)
            'tra.Value = objTra

            Dim strAryNotes As New ThreadLocal(Of String())
            Try
                strAryNotes.Value = aCMDInFo.AllFieldsValue("NOTES").Split(SplitNoteChar)
                Using cmd As OracleClient.OracleCommand = objCn.CreateCommand
                    cmd.Transaction = objTra
                    For i As Int32 = 0 To aCMDInFo.Returnlicenses.Count - 1
                        Dim aGetSeq As New ThreadLocal(Of Int32)
                        aGetSeq.Value = Convert.ToInt32(aCMDInFo.Returnlicenses.Keys.ToList.Item(i))
                        Dim strAry As New ThreadLocal(Of String())

                        strAry.Value = strAryNotes.Value.GetValue(aGetSeq.Value).ToString.Split("~")
                        Dim aProduct As New ThreadLocal(Of String)
                        aProduct.Value = strAry.Value.GetValue(0)
                        Dim aAsset As New ThreadLocal(Of String)
                        aAsset.Value = strAry.Value.GetValue(1)
                        Dim aSQL As String = String.Format("UPDATE SO561B SET LicensesKey='{0}'" & _
                                                            " WHERE SEQNO = {1} AND ProductCASID = '{2}'" & _
                                                            " AND MediaCASID = '{3}' ", aCMDInFo.Returnlicenses.Item(i.ToString) _
                                                            , aCMDInFo.SEQNO, aProduct.Value, aAsset.Value)
                        cmd.CommandText = aSQL
                        cmd.ExecuteNonQuery()
                        aGetSeq.Dispose()
                        strAry.Dispose()
                        aProduct.Dispose()
                        aAsset.Dispose()
                    Next
                End Using
            Finally
                strAryNotes.Dispose()
            End Try
           
            Return True
        Catch ex As Exception
            UpdateSysErrUI(ex, " [" & aCMDInFo.CompName & "] 命令序號 [ " & aCMDInFo.SEQNO & " ]")
            WriteErrTxtLog.WriteTxtError(ex, " [" & aCMDInFo.CompName & "] 命令序號 [ " & aCMDInFo.SEQNO & " ]", Nothing)
        Finally

        End Try
    End Function

    Private Function InsNagraCmdLog(ByRef obj As SOCMDInfo, ByRef objCn As OracleConnection, ByRef objTra As OracleTransaction) As Boolean
        Dim pCompCode As New ThreadLocal(Of OracleParameter)
        Dim pCatRanNum As New ThreadLocal(Of OracleParameter)
        Dim pGwtRanNum As New ThreadLocal(Of OracleParameter)
        Dim pHighCmd As New ThreadLocal(Of OracleParameter)
        Dim pLowCmd As New ThreadLocal(Of OracleParameter)
        Dim pICC As New ThreadLocal(Of OracleParameter)
        Dim pSTB As New ThreadLocal(Of OracleParameter)
        Dim pSendCmdText As New ThreadLocal(Of OracleParameter)
        Dim pRecvCmdText As New ThreadLocal(Of OracleParameter)
        Dim pACK As New ThreadLocal(Of OracleParameter)
        Dim pOperator As New ThreadLocal(Of OracleParameter)
        Dim pUpdTime As New ThreadLocal(Of OracleParameter)
        Dim pSendTime As New ThreadLocal(Of OracleParameter)
        Dim pRecvTime As New ThreadLocal(Of OracleParameter)
        Try
            If obj.NagraSendAndRecvMsg.Count > 0 Then
                If objCn.State = ConnectionState.Closed Then
                    objCn.Open()
                End If

                For i As Int32 = 0 To obj.NagraSendAndRecvMsg.Count - 1



                    pCompCode.Value = New OracleParameter("COMPCODE", Int32.Parse(obj.AllFieldsValue("CompCode".ToUpper)))
                    'Int32.Parse(obj.AllFieldsValue("CompCode".ToUpper))

                    pCatRanNum.Value = New OracleParameter("CATRANNUM", obj.NagraCommandType.CatRanNum)


                    pGwtRanNum.Value = New OracleParameter("GWTRANNUM", obj.NagraCommandType.GwtRanNum)


                    pHighCmd.Value = New OracleParameter("HIGHCMD", obj.AllFieldsValue.Item("HIGH_LEVEL_CMD_ID".ToUpper))


                    pLowCmd.Value = New OracleParameter("LowCmd".ToUpper, obj.NagraCommandType.LowCmd)

                    pICC.Value = New OracleParameter("ICC".ToUpper, obj.AllFieldsValue.Item("ICC_NO".ToUpper))


                    pSTB.Value = New OracleParameter("STB".ToUpper, obj.AllFieldsValue.Item("STB_NO".ToUpper))


                    pSendCmdText.Value = New OracleParameter("SENDCMDTEXT".ToUpper, obj.NagraSendAndRecvMsg.Keys(i))


                    pRecvCmdText.Value = New OracleParameter("RECVCMDTEXT".ToUpper, obj.NagraSendAndRecvMsg.Values(i))


                    pACK.Value = New OracleParameter("ACK".ToUpper, obj.NagraCommandType.Ack)


                    pOperator.Value = New OracleParameter("OPERATOR".ToUpper, "VSNagra")


                    pUpdTime.Value = New OracleParameter("UPDTIME".ToUpper, Date.Parse(Now.ToString, CultureInfo.CreateSpecificCulture("zh-TW")))


                    pSendTime.Value = New OracleParameter("SENDTIME".ToUpper, obj.NagraCommandType.SendTime)


                    pRecvTime.Value = New OracleParameter("RECVTIME".ToUpper, obj.NagraCommandType.RecvTime)


                    Using cmd As OracleClient.OracleCommand = objCn.CreateCommand
                        cmd.Transaction = objTra
                        cmd.CommandText = "INSERT INTO NAGRACMDLOG ( COMPCODE, " & _
                                "CATRANNUM, GWTRANNUM," & _
                                "HIGHCMD, LOWCMD, ICC, STB," & _
                                "SENDCMDTEXT, RECVCMDTEXT, ACK," & _
                                "OPERATOR, UPDTIME, SENDTIME, RECVTIME  )" & _
                                 "VALUES ( :COMPCODE, :CATRANNUM, :GWTRANNUM, :HIGHCMD" & _
                                 ", :LOWCMD, :ICC, :STB," & _
                                  ":SENDCMDTEXT, :RECVCMDTEXT, :ACK, :OPERATOR, " & _
                                  ":UPDTIME, :SENDTIME, :RECVTIME )"


                        cmd.Parameters.Clear()
                        cmd.Parameters.Add(pCompCode.Value)
                        cmd.Parameters.Add(pCatRanNum.Value)
                        cmd.Parameters.Add(pGwtRanNum.Value)

                        cmd.Parameters.Add(pHighCmd.Value)
                        cmd.Parameters.Add(pLowCmd.Value)
                        cmd.Parameters.Add(pICC.Value)
                        cmd.Parameters.Add(pSTB.Value)
                        cmd.Parameters.Add(pSendCmdText.Value)
                        cmd.Parameters.Add(pRecvCmdText.Value)

                        cmd.Parameters.Add(pACK.Value)
                        cmd.Parameters.Add(pOperator.Value)
                        cmd.Parameters.Add(pUpdTime.Value)
                        cmd.Parameters.Add(pSendTime.Value)
                        cmd.Parameters.Add(pRecvTime.Value)

                        cmd.ExecuteNonQuery()
                        cmd.Parameters.Clear()
                    End Using
                   

                Next

            End If
            Return True
        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, "[" & obj.CompName & "] 命令序號 [ " & obj.SEQNO & " ]", Nothing)
            UpdateSysErrUI(ex, " [" & obj.CompName & "] 命令序號 [ " & obj.SEQNO & " ]")
            Return False
        Finally
            pCompCode.Dispose()
            pCatRanNum.Dispose()
            pGwtRanNum.Dispose()
            pHighCmd.Dispose()
            pLowCmd.Dispose()
            pICC.Dispose()
            pSTB.Dispose()
            pSendCmdText.Dispose()
            pRecvCmdText.Dispose()
            pACK.Dispose()
            pOperator.Dispose()
            pUpdTime.Dispose()
            pSendTime.Dispose()
            pRecvTime.Dispose()
        End Try
    End Function


    Private Function UpdSO561(ByRef obj As SOCMDInfo, ByRef objCn As OracleConnection, ByRef objTra As OracleTransaction) As Boolean
        Try
            If obj.NagraSendAndRecvMsg.Count > 0 Then

                'Using cn As New OracleClient.OracleConnection(obj.OraConnectString)
                'Using cn As OracleClient.OracleConnection = objCn
                If objCn.State = ConnectionState.Closed Then
                    objCn.Open()
                End If
                For i As Int32 = 0 To obj.NagraSendAndRecvMsg.Count - 1
                    Dim pSocketSndMsg As New ThreadLocal(Of OracleParameter)
                    Dim pSocketRcvMsg As New ThreadLocal(Of OracleParameter)
                    Dim pSeqNo As New ThreadLocal(Of OracleParameter)
                    Try
                        Using cmd As OracleClient.OracleCommand = objCn.CreateCommand
                            cmd.Transaction = objTra

                            pSocketSndMsg.Value = New OracleParameter("SOCKETSNDMSG", OracleType.VarChar)

                            pSocketRcvMsg.Value = New OracleParameter("SOCKETRCVMSG", OracleType.VarChar)

                            pSeqNo.Value = New OracleParameter("SEQNO", OracleType.Number)
                            'pSocketSndMsg.Value = obj.AllFieldsValue.Item("SOCKETSNDMSG")
                            'pSocketRcvMsg.Value = obj.AllFieldsValue.Item("SOCKETRCVMSG")
                            pSocketSndMsg.Value.Value = obj.NagraSendAndRecvMsg.Keys(i)
                            pSocketRcvMsg.Value.Value = obj.NagraSendAndRecvMsg.Values(i)
                            pSeqNo.Value.Value = Convert.ToInt32(obj.SEQNO)
                            Dim aSQL As String = "INSERT INTO SO561 (SEQNO,SOCKETSNDMSG,SOCKETRCVMSG) " & _
                                                " VALUES ( :SEQNO,:SOCKETSNDMSG,:SOCKETRCVMSG ) "

                            cmd.CommandText = aSQL
                            cmd.Parameters.Add(pSeqNo.Value)
                            cmd.Parameters.Add(pSocketRcvMsg.Value)
                            cmd.Parameters.Add(pSocketSndMsg.Value)

                            cmd.ExecuteNonQuery()
                            cmd.Parameters.Clear()
                            pSocketSndMsg.Dispose()
                            pSocketRcvMsg.Dispose()
                            pSeqNo.Dispose()

                            'End Using
                            'If cn.State = ConnectionState.Open Then
                            '    cn.Close()
                            'End If
                        End Using
                    Finally

                        pSocketSndMsg.Dispose()
                        pSocketRcvMsg.Dispose()
                        pSeqNo.Dispose()
                    End Try

                Next
            End If
            Return True
        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, "[" & obj.CompName & "] 命令序號 [ " & obj.SEQNO & " ]", Nothing)
            UpdateSysErrUI(ex, " [" & obj.CompName & "] 命令序號 [ " & obj.SEQNO & " ]")
        Finally

        End Try
    End Function
    Private Function UpdSend_Nagra(ByRef objSOCMDInfo As SOCMDInfo) As Boolean

        Dim cn As New ThreadLocal(Of OracleConnection)
        Dim tra As New ThreadLocal(Of OracleTransaction)
        Dim pCmd_Status As New ThreadLocal(Of OracleParameter)
        Dim pUPDTIME As New ThreadLocal(Of OracleParameter)
        Dim pERR_CODE As New ThreadLocal(Of OracleParameter)
        Dim pERR_MSG As New ThreadLocal(Of OracleParameter)
        Dim pRETURNDATA As New ThreadLocal(Of OracleParameter)
        Dim pSeqNo As New ThreadLocal(Of OracleParameter)
        
        cn.Value = New OracleConnection(objSOCMDInfo.OraConnectString)
        Try
            If cn.Value.State = ConnectionState.Closed Then
                cn.Value.Open()
            End If
            tra.Value = cn.Value.BeginTransaction
            Dim aSQL As String = "UPDATE SEND_NAGRA SET CMD_STATUS = :CMD_STATUS," & _
                "UPDTIME =:UPDTIME,ERR_CODE = :ERR_CODE,ERR_MSG = :ERR_MSG," & _
                "RETURNDATA = :RETURNDATA WHERE SEQNO = :SEQNO"

            pCmd_Status.Value = New OracleParameter("CMD_STATUS", OracleType.VarChar)
            pUPDTIME.Value = New OracleParameter("UPDTIME", OracleType.DateTime)
            pERR_CODE.Value = New OracleParameter("ERR_CODE", OracleType.VarChar)
            pERR_MSG.Value = New OracleParameter("ERR_MSG", OracleType.VarChar)
            pRETURNDATA.Value = New OracleParameter("RETURNDATA", OracleType.VarChar)
            pSeqNo.Value = New OracleParameter("SEQNO", OracleType.Number)
            pCmd_Status.Value.Value = objSOCMDInfo.AllFieldsValue.Item("CMD_STATUS".ToUpper)
            pUPDTIME.Value.Value = Date.Parse(objSOCMDInfo.AllFieldsValue.Item("UPDTIME".ToUpper))
            pERR_CODE.Value.Value = objSOCMDInfo.AllFieldsValue.Item("ERR_CODE".ToUpper)
            pERR_MSG.Value.Value = objSOCMDInfo.AllFieldsValue.Item("ERR_MSG".ToUpper)
            pRETURNDATA.Value.Value = objSOCMDInfo.AllFieldsValue.Item("RETURNDATA".ToUpper)
            pSeqNo.Value.Value = Int32.Parse(objSOCMDInfo.AllFieldsValue.Item("SEQNO".ToUpper))
            Using cmd As OracleCommand = cn.Value.CreateCommand
                cmd.Parameters.Clear()
                cmd.CommandText = aSQL
                cmd.Transaction = tra.Value
                cmd.Parameters.Add(pCmd_Status.Value)
                cmd.Parameters.Add(pUPDTIME.Value)
                cmd.Parameters.Add(pERR_CODE.Value)
                cmd.Parameters.Add(pERR_MSG.Value)
                cmd.Parameters.Add(pRETURNDATA.Value)
                cmd.Parameters.Add(pSeqNo.Value)
                cmd.ExecuteNonQuery()
                cmd.Parameters.Clear()
            End Using

            If InsNagraCmdLog(objSOCMDInfo, cn.Value, tra.Value) Then
                tra.Value.Commit()
            Else
                tra.Value.Rollback()
            End If
        Catch ex As Exception
            Try
                tra.Value.Rollback()
            Finally

            End Try
            WriteErrTxtLog.WriteTxtError(ex, "[" & objSOCMDInfo.CompName & "] 命令 [ " & objSOCMDInfo.SEQNO & " ]", Nothing)
            UpdateSysErrUI(ex, " [" & objSOCMDInfo.CompName & "] 命令 [ " & objSOCMDInfo.SEQNO & " ]")
           
        Finally
            pCmd_Status.Dispose()
            pUPDTIME.Dispose()
            pERR_CODE.Dispose()
            pERR_MSG.Dispose()
            pRETURNDATA.Dispose()
            pSeqNo.Dispose()

            If tra.Value IsNot Nothing Then
                tra.Value.Dispose()
            End If
            tra.Dispose()
            If cn.Value IsNot Nothing Then
                cn.Value.Close()
                cn.Value.Dispose()
            End If
            
            cn.Dispose()

            SyncLock FSOInfoList.Item(0)
                If FSOInfoList.Item(0).WaitProcessNumber > 0 Then
                    FSOInfoList.Item(0).WaitProcessNumber -= 1
                End If
                If FUseCompIndex.Item(objSOCMDInfo.AllFieldsValue.Item("CompCode".ToUpper)).WaitProcessNumber > 0 Then
                    FUseCompIndex.Item(objSOCMDInfo.AllFieldsValue.Item("CompCode".ToUpper)).WaitProcessNumber -= 1
                End If
                If FUseCompIndex.Item(objSOCMDInfo.AllFieldsValue.Item("CompCode".ToUpper)).ProcessingNumber > 0 Then
                    FUseCompIndex.Item(objSOCMDInfo.AllFieldsValue.Item("CompCode".ToUpper)).ProcessingNumber -= 1
                End If
            End SyncLock
        End Try

      

        Return True
        'If cn.Value.State = ConnectionState.Open Then
        '    cn.Value.Close()
        '    cn.Value.Dispose()
        'End If
        'cn.Dispose()
        'If tra IsNot Nothing Then

        '    tra.Value.Dispose()
        '    tra.Dispose()
        'End If
    End Function


    Private Function UpdSO560(ByRef objSOCMDInfo As SOCMDInfo) As Boolean
        Dim cn As New ThreadLocal(Of OracleConnection)
        Dim tra As New ThreadLocal(Of OracleTransaction)
        cn.Value = New OracleConnection(objSOCMDInfo.OraConnectString)
        Try

            If cn.Value.State = ConnectionState.Closed Then
                cn.Value.Open()
            End If


            tra.Value = cn.Value.BeginTransaction
            Using cmd As OracleClient.OracleCommand = cn.Value.CreateCommand
                cmd.Transaction = tra.Value
                Dim pCmdStatus As New ThreadLocal(Of OracleParameter)
                pCmdStatus.Value = New OracleParameter("CMDSTATUS", OracleType.VarChar)
                Dim pErrCode As New ThreadLocal(Of OracleParameter)
                pErrCode.Value = New OracleParameter("ERRCODE", OracleType.VarChar)
                Dim pErrMsg As New ThreadLocal(Of OracleParameter)
                pErrMsg.Value = New OracleParameter("ERRMSG", OracleType.VarChar)
                Dim pRequestTime As New ThreadLocal(Of OracleParameter)
                pRequestTime.Value = New OracleParameter("REQUESTTIME", OracleType.DateTime)
                Dim pResponseTime As New ThreadLocal(Of OracleParameter)
                pResponseTime.Value = New OracleParameter("RESPONSETIME", OracleType.DateTime)

                Dim pSeq As New ThreadLocal(Of OracleParameter)
                Dim pPrmErrCnt As New ThreadLocal(Of OracleParameter)
                pPrmErrCnt.Value = New OracleParameter("PrmErrCnt".ToUpper, OracleType.Number)
                pSeq.Value = New OracleParameter("SEQNO", OracleType.Number)
                pCmdStatus.Value.Value = objSOCMDInfo.AllFieldsValue.Item("CMDSTATUS")
                pErrCode.Value.Value = objSOCMDInfo.AllFieldsValue.Item("ERRCODE")
                pErrMsg.Value.Value = objSOCMDInfo.AllFieldsValue.Item("ERRMSG")
                If String.IsNullOrEmpty(objSOCMDInfo.AllFieldsValue.Item("REQUESTTIME")) Then
                    pRequestTime.Value.Value = DBNull.Value
                Else
                    pRequestTime.Value.Value = Date.Parse(objSOCMDInfo.AllFieldsValue.Item("REQUESTTIME"), _
                                                            CultureInfo.CreateSpecificCulture("zh-TW"))

                End If
                If String.IsNullOrEmpty(objSOCMDInfo.AllFieldsValue.Item("RESPONSETIME")) Then
                    pResponseTime.Value.Value = DBNull.Value
                Else
                    pResponseTime.Value.Value = Date.Parse(objSOCMDInfo.AllFieldsValue.Item("RESPONSETIME"), _
                                                           CultureInfo.CreateSpecificCulture("zh-TW"))


                End If
                pSeq.Value.Value = Convert.ToInt32(objSOCMDInfo.SEQNO)
                pPrmErrCnt.Value.Value = 0



                '-----------------------------------判斷是否要還原命令--------------------------------------------------
                If (objSOCMDInfo.AllFieldsValue.Item("CMDSTATUS".ToUpper) <> "S1".ToUpper) AndAlso _
                   (Not String.IsNullOrEmpty(objSOCMDInfo.AllFieldsValue.Item("ERRCODE".ToUpper))) Then

                    If fErrResumeLst.Count > 0 Then
                        If fErrResumeLst.Contains(objSOCMDInfo.AllFieldsValue.Item("ERRCODE".ToUpper)) Then
                            If Int32.Parse("0" & objSOCMDInfo.AllFieldsValue.Item("PRMERRCNT")) >= Int32.Parse("0" & fResumeCnt) Then
                                objSOCMDInfo.AllFieldsValue.Item("PRMERRCNT") = Int32.Parse("0" & objSOCMDInfo.AllFieldsValue.Item("PRMERRCNT")) + 1
                                objSOCMDInfo.AllFieldsValue.Item("CmdStatus".ToUpper) = "E"

                            Else
                                objSOCMDInfo.AllFieldsValue.Item("PRMERRCNT".ToUpper) = Int32.Parse("0" & objSOCMDInfo.AllFieldsValue.Item("PRMERRCNT")) + 1
                                objSOCMDInfo.AllFieldsValue.Item("CmdStatus".ToUpper) = "W"
                            End If
                            pCmdStatus.Value.Value = objSOCMDInfo.AllFieldsValue.Item("CmdStatus".ToUpper)
                            pPrmErrCnt.Value.Value = Int32.Parse(objSOCMDInfo.AllFieldsValue.Item("PRMERRCNT"))
                        End If

                    End If
                End If
                '--------------------------------------------------------------------------------------------------------
                'If objSOCMDInfo.AllFieldsValue.Item("PrmErrCnt".ToUpper) Then

                'End If


                Dim aSQL As String = "UPDATE SO560 SET CMDSTATUS = :CMDSTATUS," & _
                                    "ERRCODE = :ERRCODE,ERRMSG = :ERRMSG, " & _
                                    "REQUESTTIME = :REQUESTTIME,RESPONSETIME = :RESPONSETIME ," & _
                                    "PRMERRCNT = :PRMERRCNT " & _
                                    " WHERE SEQNO = :SEQNO "
                cmd.CommandText = aSQL
                cmd.Parameters.Add(pCmdStatus.Value)
                cmd.Parameters.Add(pErrCode.Value)
                cmd.Parameters.Add(pErrMsg.Value)
                cmd.Parameters.Add(pRequestTime.Value)
                cmd.Parameters.Add(pResponseTime.Value)
                cmd.Parameters.Add(pSeq.Value)
                cmd.Parameters.Add(pPrmErrCnt.Value)
                cmd.ExecuteNonQuery()
                cmd.Parameters.Clear()
                pSeq.Dispose()
                pCmdStatus.Dispose()
                pErrCode.Dispose()
                pErrMsg.Dispose()
                pRequestTime.Dispose()
                pResponseTime.Dispose()
                pPrmErrCnt.Dispose()
            End Using
            If UpdSO561(objSOCMDInfo, cn.Value, tra.Value) Then
                If UpdSO561B(objSOCMDInfo, cn.Value, tra.Value) Then
                    tra.Value.Commit()
                Else
                    tra.Value.Rollback()
                End If

            Else
                tra.Value.Rollback()
            End If

            If cn.Value.State = ConnectionState.Open Then
                cn.Value.Close()
            End If


            Return True
        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, "[" & objSOCMDInfo.CompName & "] 命令 [ " & objSOCMDInfo.SEQNO & " ]", Nothing)
            UpdateSysErrUI(ex, " [" & objSOCMDInfo.CompName & "] 命令 [ " & objSOCMDInfo.SEQNO & " ]")
        Finally

            SyncLock FSOInfoList.Item(0)
                If FSOInfoList.Item(0).WaitProcessNumber > 0 Then
                    FSOInfoList.Item(0).WaitProcessNumber -= 1
                End If
                If FUseCompIndex.Item(objSOCMDInfo.AllFieldsValue.Item("CompCode".ToUpper)).WaitProcessNumber > 0 Then
                    FUseCompIndex.Item(objSOCMDInfo.AllFieldsValue.Item("CompCode".ToUpper)).WaitProcessNumber -= 1
                End If
                If FUseCompIndex.Item(objSOCMDInfo.AllFieldsValue.Item("CompCode".ToUpper)).ProcessingNumber > 0 Then
                    FUseCompIndex.Item(objSOCMDInfo.AllFieldsValue.Item("CompCode".ToUpper)).ProcessingNumber -= 1
                End If
            End SyncLock
            If cn.Value.State = ConnectionState.Open Then
                cn.Value.Close()
                cn.Value.Dispose()
            End If
            cn.Dispose()
            If tra IsNot Nothing Then

                tra.Value.Dispose()
                tra.Dispose()
            End If
        End Try
    End Function

    '更新畫面與寫入資料庫
    Public Sub UpdUiAndDB(ByVal aSOCmdInfo As SOCMDInfo)
        Try
            '更新資料庫,如果恢復成W,則低階命令與高階命令視窗不要顯示出來
            'If aSOCmdInfo.AllFieldsValue("CmdStatus".ToUpper) = "W" Then
            '    '將處理的資料減一個回來
            '    SyncLock FSOInfoList.Item(0)
            '        If FSOInfoList.Item(0).WaitProcessNumber > 0 Then
            '            FSOInfoList.Item(0).WaitProcessNumber -= 1
            '        End If
            '        If FUseCompIndex.Item(aSOCmdInfo.AllFieldsValue.Item("CompCode".ToUpper)).WaitProcessNumber > 0 Then
            '            FUseCompIndex.Item(aSOCmdInfo.AllFieldsValue.Item("CompCode".ToUpper)).WaitProcessNumber -= 1
            '        End If
            '        If FUseCompIndex.Item(aSOCmdInfo.AllFieldsValue.Item("CompCode".ToUpper)).ProcessingNumber > 0 Then
            '            FUseCompIndex.Item(aSOCmdInfo.AllFieldsValue.Item("CompCode".ToUpper)).ProcessingNumber -= 1
            '        End If
            '    End SyncLock
            '    UpdateSOStatusUI(FSOInfoList.Item(0), SOStatus.NotUpdateImg)
            '    Exit Sub
            'End If

            'UpdSO560(aSOCmdInfo)
            UpdSend_Nagra(aSOCmdInfo)
            If (aSOCmdInfo.AllFieldsValue("Cmd_Status".ToUpper).ToUpper = "E".ToUpper) AndAlso _
                    (fErrSndMail) AndAlso (fIsUseMail) Then
                csCommon.SendMail(MDBMailPara.GetDefaultMDBName, _
                                  fMDBMailPassword, "CMD_ERR", fMailInfoTableName, _
                            fMailToTableName, fMailCCTableName, "VSNagra_GATEWAY 命令有錯", _
                            String.Format("命令序號 {0} 有錯誤", aSOCmdInfo.AllFieldsValue.Item("SEQNO".ToUpper)), _
                             False, True)
            End If


            '更新畫面,如果是還原資料不要更新高低階命令
            If aSOCmdInfo.AllFieldsValue("Cmd_Status".ToUpper) <> "W" Then
                UpdateHightCmdUI(aSOCmdInfo, UpdCmdMode.Add)
                UpdateLowCmdUI(aSOCmdInfo)
            End If

            UpdateSOStatusUI(FSOInfoList.Item(0), SOStatus.NotUpdateImg)
        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, "更新畫面與回填資料庫失敗")
        Finally
            '將有使用到的Thread減回來
            If FThreadsTotal > 0 Then
                Interlocked.Decrement(FThreadsTotal)
            End If
            If FThreadsTotal <= 0 Then
                evn.Set()
            End If


        End Try
    End Sub

    Private Function RecvMsg(ByVal aSocketNum As Int32, _
                                    ByRef aCmdInfo As SOCMDInfo, _
                                    ByVal aCmdType As CmdType, _
                                    ByRef aNagraStatus As NagraCmdStatus_Type, _
                                    ByVal aSndMsg() As Byte) As Boolean


        Dim aRetCmdString As New ThreadLocal(Of String)
        Dim pos As New ThreadLocal(Of Int32)
        Dim aTimeStop As New ThreadLocal(Of Stopwatch)
        Dim aRcvBty As New ThreadLocal(Of Byte())
        Try
            For i As Int32 = 1 To FRcvErrStopCnt
                Try

                    pos.Value = 0

                    aTimeStop.Value = New Stopwatch
                    aTimeStop.Value.Start()
                    Do
                        If (aTimeStop.Value.Elapsed.Seconds) >= (FRcvTimeout) Then
                            aCmdInfo.AllFieldsValue.Item("CMD_STATUS".ToUpper) = "T"
                            aCmdInfo.AllFieldsValue.Item("ERR_CODE".ToUpper) = String.Empty
                            aCmdInfo.AllFieldsValue.Item("ERR_MSG".ToUpper) = String.Empty
                            aTimeStop.Value.Stop()
                            aTimeStop.Dispose()
                            pos.Dispose()
                            Exit Do
                        End If
                        If (fSocket(aSocketNum) Is Nothing) OrElse (Not fSocket(aSocketNum).Connected) Then
                            aCmdInfo.AllFieldsValue.Item("CMD_STATUS".ToUpper) = "E"
                            aCmdInfo.AllFieldsValue.Item("ERR_CODE".ToUpper) = CustomerNagraError.CableSoft_Recv_Socket_Disconnect
                            aCmdInfo.AllFieldsValue.Item("ERR_MSG".ToUpper) = [Enum].GetName(GetType(CustomerNagraError), CustomerNagraError.CableSoft_Recv_Socket_Disconnect)
                            aTimeStop.Dispose()
                            pos.Dispose()
                            Return False
                        End If



                        ReDim aRcvBty.Value(1023)

                        If fSocket(aSocketNum).Receive(aRcvBty.Value) > 0 Then

                            '取出此次接收的長度
                            If pos.Value = 0 Then
                                pos.Value = fNagraCommand(aSocketNum).GetRetunLength(aRcvBty.Value)
                            End If
                            If pos.Value > 0 Then
                                If String.IsNullOrEmpty(aRetCmdString.Value) Then
                                    aRetCmdString.Value = fNagraCommand(aSocketNum).GetReturnString(aRcvBty.Value, True)
                                Else
                                    aRetCmdString.Value = aRetCmdString.Value & fNagraCommand(aSocketNum).GetReturnString(aRcvBty.Value, False)
                                End If
                                '比對接收的資料與長度是否符合
                                '成功回傳的資料 000095437050000040000259201101202005000021959001036020a102000063fc0100010000ffffffffdeec762fa2cef41d9ab58f6a05126e2b4c2024db9557468ab992bef4a501b820ffc18cb1bc086a6ea906fcb43528904c8b3cce7822117181bd039cb8fce04e7a4eba1dff4b9ae8a87f1215dd77baf6770419b33326a3128eba340bc3a29d99dbdcbb6a732c8534530a245bee49f78a679f47ab0c0af92c88b8cfcde5b8781241ec59cd0c13ec060869039d8b11e285e4f0bac9a0e4dec31ea1f8fb781131bcb8c932201a6bca935925a102000034ff01caa70624739ebb055579218ebf337fc9606cc2b303f68f1d0abb23ef818f4b6d03d33fd9f3b6529dcbaf04658f3b57cffa38ee113a822267541f6501fbbc0bd09fc142b6c6cba6856c4f643aac3aec0b92421a58e45f479ba90f46a936f8ae6447ef7eb39a707f5596f73b1794b3ff804586499a89c5f892b0215fdd02782b7eb2903c7b0c80018ea54b76d73dac2c7753432305f630dcb69bd3f535d05afd9acec13daf4c42a66d
                                'Nack 回傳的資料 000020199050000040000259201101282006001753480105060000
                                If pos.Value <= aRetCmdString.Value.Length Then
                                    'aRetCmdString.Value = "000095437050000040000259201101202005000021959001036020a102000063fc0100010000ffffffffdeec762fa2cef41d9ab58f6a05126e2b4c2024db9557468ab992bef4a501b820ffc18cb1bc086a6ea906fcb43528904c8b3cce7822117181bd039cb8fce04e7a4eba1dff4b9ae8a87f1215dd77baf6770419b33326a3128eba340bc3a29d99dbdcbb6a732c8534530a245bee49f78a679f47ab0c0af92c88b8cfcde5b8781241ec59cd0c13ec060869039d8b11e285e4f0bac9a0e4dec31ea1f8fb781131bcb8c932201a6bca935925a102000034ff01caa70624739ebb055579218ebf337fc9606cc2b303f68f1d0abb23ef818f4b6d03d33fd9f3b6529dcbaf04658f3b57cffa38ee113a822267541f6501fbbc0bd09fc142b6c6cba6856c4f643aac3aec0b92421a58e45f479ba90f46a936f8ae6447ef7eb39a707f5596f73b1794b3ff804586499a89c5f892b0215fdd02782b7eb2903c7b0c80018ea54b76d73dac2c7753432305f630dcb69bd3f535d05afd9acec13daf4c42a66d"
                                    'aRetCmdString.Value = "000020199050000040000259201101282006001753480105060000"
                                    'pos.Value = aRetCmdString.Value.Length

                                    
                                    'aCmdInfo.AllFieldsValue.Item("SOCKETRCVMSG") = aRetCmdString.Value
                                    '儲存接收資料
                                    If aCmdInfo.NagraSendAndRecvMsg.ContainsKey(fNagraCommand(aSocketNum).GetReturnString(aSndMsg, True)) Then
                                        aCmdInfo.NagraSendAndRecvMsg.Item(fNagraCommand(aSocketNum).GetReturnString(aSndMsg, True)) = _
                                                                      aRetCmdString.Value
                                    End If
                                    Dim aSendTime As Date = aNagraStatus.SendTime
                                    aNagraStatus = fNagraCommand(aSocketNum).GetCMDStatus(aRetCmdString.Value, pos.Value, aCmdType)
                                    aNagraStatus.SendTime = aSendTime
                                    aNagraStatus.RecvTime = Now
                                    aNagraStatus.GwtRanNum = aRetCmdString.Value.Substring(0, 9)
                                    aCmdInfo.AllFieldsValue.Item("UpdTime".ToUpper) = Date.Now                                    
                                    '取得Nagra回覆的訊息
                                    Select Case aNagraStatus.Status
                                        Case AckCode.Ack
                                            aCmdInfo.AllFieldsValue.Item("CMD_STATUS".ToUpper) = "C"
                                            aCmdInfo.AllFieldsValue.Item("ERR_CODE".ToUpper) = String.Empty
                                            aCmdInfo.AllFieldsValue.Item("ERR_MSG".ToUpper) = String.Empty
                                            Select Case aCmdType
                                                Case CmdType.cmd0126
                                                    aCmdInfo.AllFieldsValue.Item("RETURNDATA".ToUpper) =
                                                        aNagraStatus.VUA

                                            End Select
                                            pos.Dispose()
                                            'aRcvBty.Dispose()
                                            Return True     '如果是ACK 回傳True

                                        Case AckCode.Nack
                                            aCmdInfo.AllFieldsValue.Item("CMD_STATUS".ToUpper) = "E"
                                            aCmdInfo.AllFieldsValue.Item("ERR_CODE".ToUpper) = aNagraStatus.ErrorCode
                                            aCmdInfo.AllFieldsValue.Item("ERR_MSG".ToUpper) = aNagraStatus.ErrorName
                                            If aNagraStatus.Nack_Status = "1" Then
                                                Try
                                                    'fSocket(aSocketNum).Shutdown(SocketShutdown.Both)
                                                    'fSocket(aSocketNum).Disconnect(False)
                                                Catch ex As Exception

                                                End Try
                                            End If
                                        Case AckCode.Other
                                            aCmdInfo.AllFieldsValue.Item("CMD_STATUS".ToUpper) = "E"
                                            aCmdInfo.AllFieldsValue.Item("ERR_CODE".ToUpper) = aNagraStatus.ErrorCode
                                            aCmdInfo.AllFieldsValue.Item("ERR_MSG".ToUpper) = "CableSoft_Recv_Othre_Error"
                                    End Select

                                    Return False
                                End If
                            End If
                        End If
                    Loop
                    Return False
                Catch ex As SocketException
                    Try
                        fSocket(aSocketNum).Shutdown(SocketShutdown.Both)
                        fSocket(aSocketNum).Disconnect(False)
                    Finally
                        If fSocket(aSocketNum) IsNot Nothing Then
                            fSocket(aSocketNum).Dispose()
                        End If
                        fSocket(aSocketNum) = Nothing
                    End Try
                    WriteErrTxtLog.WriteTxtError(ex, " [" & aCmdInfo.CompName & "] 命令序號 [ " & aCmdInfo.SEQNO & " ] 第 " & i & " 次接收", Nothing)
                    UpdateSysErrUI(ex, " [" & aCmdInfo.CompName & "] 命令序號 [ " & aCmdInfo.SEQNO & " ] 第 " & i & " 次接收")
                    aCmdInfo.AllFieldsValue.Item("Cmd_Status".ToUpper) = "E"
                    aCmdInfo.AllFieldsValue.Item("Err_Code".ToUpper) = Int32.Parse(ex.SocketErrorCode)
                    aCmdInfo.AllFieldsValue.Item("Err_Msg".ToUpper) = ex.SocketErrorCode
                    Return False

                Catch ex As Exception
                    UpdateSysErrUI(ex, " [" & aCmdInfo.CompName & "] 命令序號 [ " & aCmdInfo.SEQNO & " ] 第 " & i & " 次接收")
                    WriteErrTxtLog.WriteTxtError(ex, " [" & aCmdInfo.CompName & "] 命令序號 [ " & aCmdInfo.SEQNO & " ] 第 " & i & " 次接收", Nothing)
                    If i = FRcvErrStopCnt Then
                        aCmdInfo.AllFieldsValue.Item("Cmd_Status".ToUpper) = "E"
                        aCmdInfo.AllFieldsValue.Item("Err_Code".ToUpper) = CustomerNagraError.CableSoft_Socket_Recv_Error
                        aCmdInfo.AllFieldsValue.Item("Err_Msg".ToUpper) = [Enum].GetName(GetType(CustomerNagraError), CustomerNagraError.CableSoft_Socket_Recv_Error)
                        Return False
                    End If
                End Try
            Next
        Catch ex As Exception
            'FNagraStatusOK = False
            WriteErrTxtLog.WriteTxtError(ex, " [" & aCmdInfo.CompName & "] 命令序號 [ " & aCmdInfo.SEQNO & " ]", Nothing)
            UpdateSysErrUI(ex, " [" & aCmdInfo.CompName & "] 命令序號 [ " & aCmdInfo.SEQNO & " ]")
            Return False
        Finally
            aRetCmdString.Dispose()
            Erase aSndMsg
            pos.Dispose()
            aTimeStop.Dispose()
            If aRcvBty.Value IsNot Nothing Then
                Erase aRcvBty.Value
            End If
            If aRcvBty IsNot Nothing Then
                aRcvBty.Dispose()
            End If
            'Interlocked.Decrement(FNagraHaveProcing)
        End Try

    End Function




    Private Function SendMsg(ByVal aSocketNum As Int32, _
                                    ByRef aCMDInfo As SOCMDInfo, _
                                    ByVal aCmdType As CmdType, _
                                    ByVal aMsgBty() As Byte) As Boolean



        Try
            For i As Int32 = 1 To FSndErrStopCnt
                Try
                    If (fSocket(aSocketNum) Is Nothing) OrElse (Not fSocket(aSocketNum).Connected) Then
                        aCMDInfo.AllFieldsValue.Item("Cmd_Status".ToUpper) = "E"
                        aCMDInfo.AllFieldsValue.Item("Err_Code".ToUpper) = Int32.Parse(CustomerNagraError.CableSoft_Socket_Disconnect).ToString
                        aCMDInfo.AllFieldsValue.Item("Err_Msg".ToUpper) = [Enum].GetName(GetType(CustomerNagraError), CustomerNagraError.CableSoft_Socket_Disconnect)
                        Return False
                    End If

                    'aCMDInfo.AllFieldsValue.Item("SOCKETSNDMSG") = fNagraCommand(aSocketNum).GetReturnString(aMsgBty, True)
                    'aCMDInfo.AllFieldsValue.Item("REQUESTTIME") = Now
                    aCMDInfo.NagraCommandType.SendTime = Now
                    aCMDInfo.NagraSendAndRecvMsg.Add(fNagraCommand(aSocketNum).GetReturnString(aMsgBty, True), String.Empty)

                    fSocket(aSocketNum).Send(aMsgBty)
                    Return True
                Catch ex As SocketException
                    Try
                        fSocket(aSocketNum).Shutdown(SocketShutdown.Both)
                        fSocket(aSocketNum).Disconnect(False)
                    Finally
                        If fSocket(aSocketNum) IsNot Nothing Then
                            fSocket(aSocketNum).Dispose()
                        End If
                        fSocket(aSocketNum) = Nothing
                    End Try
                    WriteErrTxtLog.WriteTxtError(ex, " [" & aCMDInfo.CompName & "] 命令序號 [ " & aCMDInfo.SEQNO & " ] 第 " & i & " 次傳送", Nothing)
                    UpdateSysErrUI(ex, " [" & aCMDInfo.CompName & "] 命令序號 [ " & aCMDInfo.SEQNO & " ] 第 " & i & " 次傳送")
                    aCMDInfo.AllFieldsValue.Item("Cmd_Status".ToUpper) = "E"
                    aCMDInfo.AllFieldsValue.Item("Err_Code".ToUpper) = ex.SocketErrorCode.ToString
                    aCMDInfo.AllFieldsValue.Item("Err_Msg".ToUpper) = [Enum].GetName(GetType(SocketError), ex.SocketErrorCode)

                    Return False
                Catch ex As Exception

                    UpdateSysErrUI(ex, " [" & aCMDInfo.CompName & "] 命令序號 [ " & aCMDInfo.SEQNO & " ] 第 " & i & " 次傳送")
                    WriteErrTxtLog.WriteTxtError(ex, " [" & aCMDInfo.CompName & "] 命令序號 [ " & aCMDInfo.SEQNO & " ] 第 " & i & " 次傳送", Nothing)
                    If i = FSndErrStopCnt Then
                        aCMDInfo.NagraCommandType.Status = AckCode.Nack
                        aCMDInfo.NagraCommandType.ErrorCode = CustomerNagraError.CableSoft_Socket_SndOver
                        aCMDInfo.NagraCommandType.ErrorName = [Enum].GetName(GetType(CustomerNagraError), CustomerNagraError.CableSoft_Socket_SndOver)
                        aCMDInfo.AllFieldsValue.Item("Cmd_Status".ToUpper) = "E"
                        aCMDInfo.AllFieldsValue.Item("Err_Code".ToUpper) = CustomerNagraError.CableSoft_Socket_SndOver
                        aCMDInfo.AllFieldsValue.Item("Err_Msg".ToUpper) = [Enum].GetName(GetType(CustomerNagraError), CustomerNagraError.CableSoft_Socket_SndOver)
                        Return False
                    End If
                End Try

            Next
        Catch ex As Exception
            'FNagraStatusOK = False
            WriteErrTxtLog.WriteTxtError(ex, " [" & aCMDInfo.CompName & "] 命令序號 [ " & aCMDInfo.SEQNO & " ]", Nothing)
            UpdateSysErrUI(ex, " [" & aCMDInfo.CompName & "] 命令序號 [ " & aCMDInfo.SEQNO & " ]")

        Finally
            Erase aMsgBty
            'Interlocked.Increment(FNagraHaveProcing)
        End Try
    End Function


    'Private Sub SendNagraCmd(ByVal aSocket As Socket, ByVal aSocketNum As Int32, _
    '                         ByVal aSoCmdInfo As SOCMDInfo)
    '    Try
    '        SyncLock aSocket
    '            Dim aTypeCmd As New ThreadLocal(Of CmdType)

    '            Select Case aSoCmdInfo.CMD.ToUpper
    '                Case "B1"
    '                    aTypeCmd.Value = CmdType.Cmd0851
    '                Case Else
    '                    aTypeCmd.Value = CmdType.OtherCMD
    '                    Exit Sub
    '            End Select
    '            Dim aryMsg As New ThreadLocal(Of List(Of Byte()))
    '            aryMsg.Value = New List(Of Byte())
    '            aryMsg.Value = fNagraCommand(aSocketNum).BuilderCmd(aTypeCmd.Value, aSoCmdInfo.AllFieldsValue)

    '        End SyncLock

    '    Catch ex As Exception

    '    Finally

    '    End Try
    'End Sub

    Private Sub UpdateProcessCmd(ByVal obj As Array)

        Dim aSO As SOInfoClass = FSOIndex(obj.GetValue(2).ToString)
        Dim aSEQNo As String = obj.GetValue(0)
        Dim aCMD As String = obj.GetValue(1)
        Try
            Interlocked.Increment(FThreadsTotal)
            SyncLock aSO.WaitProcessNumber.GetType
                aSO.WaitProcessNumber += 1
                aSO.ProcessingNumber += 1
            End SyncLock
            SyncLock aSO.ProcessingNumber.GetType

            End SyncLock
            Using cn As New OracleConnection(aSO.OraConnectString)
                Try
                    cn.Open()
                Catch ex As Exception
                    UpdateSOStatusUI(aSO, SOStatus.NO)
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
                    UpdateSOStatusUI(aSO, SOStatus.NO)
                End If
            Else
                UpdateSOStatusUI(aSO, SOStatus.Yes)
            End If
            UpdateSysErrUI(ex, "TEST")
            WriteErrTxtLog.WriteTxtError(ex, aSO.CompID, aSEQNo, Nothing, Nothing, Nothing)
        Finally
            SyncLock aSO.ProcessingNumber.GetType
                aSO.ProcessingNumber -= 1
            End SyncLock
            SyncLock aSO.WaitProcessNumber.GetType
                aSO.WaitProcessNumber -= 1
            End SyncLock
            Interlocked.Decrement(FThreadsTotal)
            UpdateSOStatusUI(aSO, SOStatus.NotUpdateImg)
            Erase obj




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

    'Private Sub SendNagraSocket(ByVal objSoCmdInfo As SOCMDInfo)
    '    Dim lst As List(Of Byte()) = NagraCommand.BuilderCmd(CmdType.Cmd0851, objSoCmdInfo.AllFieldsValue)

    '    For i As Int32 = 0 To lst.Count - 1
    '        Send(Nagra_Socket.FSocket, lst.Item(i))
    '        sendDone.WaitOne(TimeSpan.FromSeconds(FSndTimeout))
    '        Receive(Nagra_Socket.FSocket)
    '        receiveDone.WaitOne(TimeSpan.FromSeconds(FRcvTimeout))
    '        'sendDone.WaitOne(TimeSpan.FromSeconds(FSndTimeout))
    '    Next

    'End Sub
    Private Sub Send(ByVal client As Socket, ByVal data As Byte())
        client.BeginSend(data, 0, data.Length, SocketFlags.None, New AsyncCallback(AddressOf SendCallback), client)
    End Sub
    Private Sub SendCallback(ByVal ar As IAsyncResult)

        Dim client As Socket = CType(ar.AsyncState, Socket)

        Dim bytesSent As Integer = client.EndSend(ar)

        'sendDone.Set()


    End Sub
    'Private Sub Receive(ByVal client As Socket)
    '    Dim state As New StateSocket
    '    state.workSocket = client

    '    client.BeginReceive(state.buffer, 0, StateSocket.BufferSize, 0, New AsyncCallback(AddressOf ReceiveCallback), state)
    'End Sub

    'Private Sub ReceiveCallback(ByVal ar As IAsyncResult)

    '    Dim state As StateSocket = CType(ar.AsyncState, StateSocket)
    '    Dim client As Socket = state.workSocket
    '    Dim bytesRead As Integer = client.EndReceive(ar)
    '    If bytesRead > 0 Then
    '        'Dim bty() As Byte = Encoding.ASCII.GetBytes(Encoding.ASCII.GetChars(state.buffer, 0, bytesRead))
    '        state.bty = Encoding.ASCII.GetBytes(New String(Chr(0), state.bty.Length - 1))
    '        Array.Copy(state.buffer, 0, state.bty, 0, bytesRead)
    '        CodeSite.SendCollection("Recv,Len=" & bytesRead, state.bty)

    '        If state.fPos = 0 Then
    '            state.fPos = NagraCommand.GetRetunLength(state.bty)
    '            state.sb.Append(NagraCommand.GetReturnString(state.bty, True))
    '        Else
    '            state.sb.Append(NagraCommand.GetReturnString(state.bty, False))
    '        End If
    '        If state.fPos = state.sb.ToString.Length Then
    '            CodeSite.Send("完成,Handle=" & state.workSocket.Handle.ToString & "," & state.sb.ToString)

    '            'receiveDone.Set()
    '            'Exit Sub
    '        End If
    '        If state.workSocket.Available > 0 Then


    '            client.BeginReceive(state.buffer, 0, StateSocket.BufferSize, 0, New AsyncCallback(AddressOf ReceiveCallback), state)


    '        Else
    '            If state.fPos <> state.sb.ToString.Length Then
    '                CodeSite.Send("長度=" & state.fPos & "," & state.sb.ToString)
    '            End If
    '            'client.BeginReceive(state.buffer, 0, 0, 0, New AsyncCallback(AddressOf ReceiveCallback), state)
    '            'receiveDone.Set()
    '        End If

    '        'receiveDone.Set()
    '    Else
    '        receiveDone.Set()
    '    End If
    'End Sub






    'Private Sub Send(ByVal client As Socket, ByVal data As String)
    '    ' Convert the string data to byte data using ASCII encoding.
    '    Dim byteData As Byte() = Encoding.ASCII.GetBytes(data)

    '    ' Begin sending the data to the remote device.
    '    client.BeginSend(byteData, 0, byteData.Length, 0, New AsyncCallback(AddressOf SendCallback), client)
    'End Sub 'Sen

    Private Function FMainForm() As Exception
        Throw New NotImplementedException
    End Function





End Module

Public Class StateSocket
    <ThreadStatic()> _
    Public workSocket As Socket = Nothing
    <ThreadStatic()> _
    Public Const BufferSize As Integer = 100
    <ThreadStatic()> _
    Public buffer(BufferSize - 1) As Byte
    <ThreadStatic()> _
    Public sb As New StringBuilder
    <ThreadStatic()> _
    Public fPos As Int32 = 0
    <ThreadStatic()> _
    Public bty(1024) As Byte
End Class

Public Class SOCMDInfo


    Protected Friend Sub New()

    End Sub
    Private _CompId As String
    Private _SEQNO As String
    Private _CMD As String
    Private _OraConnectString As String
    Private _NeedSocket As Int32 = 1
    Private _CompName As String
    Private _CmdPriority As Int32
    Private _AllFieldsValue As Dictionary(Of String, String)
    Public Property NagraCommandType As NagraCmdStatus_Type
    Public Property Returnlicenses As Dictionary(Of String, String)
    Public Property NagraSendAndRecvMsg As Dictionary(Of String, String)


    Public Property CompId() As String
        Get
            Return _CompId

        End Get
        Set(ByVal value As String)
            _CompId = value
        End Set
    End Property

    Public Property SEQNO() As String
        Get
            Return _SEQNO
        End Get
        Set(ByVal value As String)
            _SEQNO = value
        End Set
    End Property
    Public Property CMD() As String
        Get
            Return _CMD
        End Get
        Set(ByVal value As String)
            _CMD = value
        End Set
    End Property
    Public Property OraConnectString()
        Get
            Return _OraConnectString
        End Get
        Set(ByVal value)
            _OraConnectString = value
        End Set
    End Property
    Public Property NeedSocket() As Int32
        Get
            Return _NeedSocket
        End Get
        Set(ByVal value As Int32)
            _NeedSocket = CMD
        End Set
    End Property
    Public Property CompName() As String
        Get
            Return _CompName
        End Get
        Set(ByVal value As String)
            _CompName = value
        End Set
    End Property
    Public Property CmdPriority() As Integer
        Get
            Return _CmdPriority
        End Get
        Set(ByVal value As Integer)
            _CmdPriority = value
        End Set
    End Property

    Public Property AllFieldsValue() As Dictionary(Of String, String)
        Get
            Return _AllFieldsValue
        End Get
        Set(ByVal value As Dictionary(Of String, String))
            _AllFieldsValue = value
        End Set
    End Property





End Class


Public Class SendCmdInfo

    Private Shared SOCMDInfoLst As New List(Of SOCMDInfo)

    Public Shared Function ADD(ByVal aCompID As String, ByVal aCompName As String, _
                            ByVal aSEQNO As String, _
                            ByVal aCMDName As String, _
                            ByVal aConnectString As String, _
                            ByVal aDataReader As OracleClient.OracleDataReader) As SOCMDInfo


        Try
            rwl.EnterWriteLock()
            SyncLock aDataReader
                Dim obj As New SOCMDInfo()
                Dim objList As New Dictionary(Of String, String)
                For i As Int32 = 0 To aDataReader.FieldCount - 1
                    If aDataReader.IsDBNull(i) Then
                        objList.Add(aDataReader.GetName(i).ToUpper, String.Empty)
                    Else
                        objList.Add(aDataReader.GetName(i).ToUpper, aDataReader.Item(i))
                    End If
                Next
                objList.Add("SOCKETSNDMSG", String.Empty)
                objList.Add("SOCKETRCVMSG", String.Empty)
                objList.Add("GUIDNO", String.Empty)
                objList.Add("TranNum", String.Empty)
                If Not objList.ContainsKey("Returnlicenses".ToUpper) Then
                    objList.Add("Returnlicenses".ToUpper, String.Empty)
                End If
                objList.Item("Cmd_Status".ToUpper) = "P"
                obj.NagraSendAndRecvMsg = New Dictionary(Of String, String)
                obj.CompId = aCompID
                obj.SEQNO = aSEQNO
                obj.CMD = aCMDName
                obj.CompName = aCompName
                obj.OraConnectString = aConnectString
                obj.CmdPriority = 9
                obj.AllFieldsValue = objList
                obj.Returnlicenses = New Dictionary(Of String, String)
                obj.NagraCommandType = New NagraCmdStatus_Type
                obj.NagraCommandType.Status = AckCode.Other
                obj.NagraCommandType.ErrorCode = CustomerNagraError.CableSoft_Initial
                obj.NagraCommandType.ErrorName = [Enum].GetName(GetType(CustomerNagraError), CustomerNagraError.CableSoft_Initial)
                SOCMDInfoLst.Add(obj)
                Return obj
            End SyncLock
        Catch ex As Exception

            CableSoft.GateWay.csException.WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
            Return Nothing
        Finally
            rwl.ExitWriteLock()
        End Try
    End Function

    Public Shared Function GetInfoLst() As List(Of SOCMDInfo)
        Return SOCMDInfoLst
    End Function

    Private Shared Function CompareCmdPriority(ByVal p1 As SOCMDInfo, ByVal p2 As SOCMDInfo) As Int32
        Try
            Return Comparer(Of Int32).Default.Compare(p1.SEQNO, p2.SEQNO)
        Catch ex As Exception
            Return 0
        End Try
    End Function


End Class