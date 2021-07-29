'Imports System
'Imports System.IO
'Imports System.Text
'Imports System.Net
'Imports System.Net.Sockets
'Imports CableSoft.Gateway.csException
'Imports System.Threading
'Imports CableSoft_Nagra_BuilderCmd
'Imports System.Data
'Imports System.Data.OracleClient
'Imports Raize.CodeSiteLogging
'Public Class Nagra_Socket
'    Public Shared FMaxSocket As Int32 = 30
'    Public Shared FUseSocket As Int32 = 0
'    Public Shared FSocket As Socket = Nothing
'    Public Shared FMainFrm As rfrmMain
'    Public Shared tmrCount As New Stopwatch
'    Private Const SplitReturnChar As Char = ","c
'    Private Const SplitNoteChar As Char = ","c

'    'Public Shared Sub SendCMD()
'    '    FNagraConnecting = True

'    '    If FThreadStop Then
'    '        SendCmdInfo.GetInfoLst.Clear()
'    '        tmrCount.Stop()
'    '        tmrCount.Reset()
'    '        Exit Sub
'    '    End If
'    '    If (Not tmrCount.IsRunning) Then
'    '        tmrCount.Start()
'    '    End If
'    '    If tmrCount.Elapsed.Minutes > 60 Then
'    '        tmrCount.Reset()
'    '    End If
'    '    Try
'    '        'SyncLock SendCmdInfo.GetInfoLst
'    '        Do While Not SendCmdInfo.GetInfoLst.Count <= 0
'    '            'Application.DoEvents()



'    '            If FThreadStop Then
'    '                SendCmdInfo.GetInfoLst.Clear()
'    '                tmrCount.Stop()
'    '                tmrCount.Reset()
'    '                Exit Do
'    '            End If

'    '            'Thread.Sleep(100)
'    '            If tmrCount.IsRunning Then

'    '                If tmrCount.Elapsed.Seconds >= FCheckStatusTime Then
'    '                    Try
'    '                        Interlocked.Increment(FUseSocket)
'    '                        If FUseSocket <= FMaxSocket Then
'    '                            Try
'    '                                FNagraStatusOK = ConnectSocket()
'    '                            Finally

'    '                                tmrCount.Stop()
'    '                                tmrCount.Reset()
'    '                                tmrCount.Start()
'    '                            End Try
'    '                        End If
'    '                    Finally
'    '                        Interlocked.Decrement(FUseSocket)
'    '                    End Try

'    '                Else
'    '                    If (FUseSocket + SendCmdInfo.GetInfoLst(0).NeedSocket) <= (FMaxSocket) Then
'    '                        CodeSite.Clear()

'    '                        Try

'    '                            If UpdProcing(SendCmdInfo.GetInfoLst(0)) Then

'    '                                If FNagraStatusOK Then

'    '                                    If FSocket.Connected Then
'    '                                        Interlocked.Add(FUseSocket, SendCmdInfo.GetInfoLst(0).NeedSocket)

'    '                                        If FUseSocket <= FMaxSocket Then
'    '                                            '一次只能送一個Socket
'    '                                            RunSocketThread(SendCmdInfo.GetInfoLst(0))
'    '                                            'FNagraConnecting = True
'    '                                            'Dim thr As New Thread(AddressOf SendSocket)
'    '                                            'thr.IsBackground = True
'    '                                            'thr.Start(SendCmdInfo.GetInfoLst(0))
'    '                                        Else
'    '                                            Interlocked.Decrement(FUseSocket)
'    '                                        End If
'    '                                    Else

'    '                                        FNagraStatusOK = False
'    '                                        'Interlocked.Add(FUseSocket, SendCmdInfo.GetInfoLst(0).NeedSocket)
'    '                                        Try
'    '                                            Interlocked.Increment(FUseSocket)
'    '                                            If FUseSocket <= FMaxSocket Then
'    '                                                FNagraStatusOK = ConnectSocket()
'    '                                            End If
'    '                                        Finally
'    '                                            Interlocked.Decrement(FUseSocket)
'    '                                        End Try



'    '                                        If Not FNagraStatusOK Then
'    '                                            SendCmdInfo.GetInfoLst(0).AllFieldsValue.Item("ERRCODE") = "-9999"
'    '                                            SendCmdInfo.GetInfoLst(0).AllFieldsValue.Item("ERRMSG") = "CableSoft_Socket_1002CMD_Disconnect"
'    '                                            'Interlocked.Add(FUseSocket, (SendCmdInfo.GetInfoLst(0).NeedSocket * -1))
'    '                                            UpdFailData(SendCmdInfo.GetInfoLst(0), Nothing)

'    '                                        Else
'    '                                            Interlocked.Add(FUseSocket, SendCmdInfo.GetInfoLst(0).NeedSocket)
'    '                                            RunSocketThread(SendCmdInfo.GetInfoLst(0))
'    '                                        End If

'    '                                    End If
'    '                                Else

'    '                                    FNagraStatusOK = False
'    '                                    Try
'    '                                        Interlocked.Increment(FUseSocket)
'    '                                        If FUseSocket <= FMaxSocket Then
'    '                                            FNagraStatusOK = ConnectSocket()
'    '                                        End If
'    '                                    Finally
'    '                                        Interlocked.Decrement(FUseSocket)
'    '                                    End Try

'    '                                    If Not FNagraStatusOK Then
'    '                                        SendCmdInfo.GetInfoLst(0).AllFieldsValue.Item("ERRCODE") = "-9999"
'    '                                        SendCmdInfo.GetInfoLst(0).AllFieldsValue.Item("ERRMSG") = "CableSoft_Socket_Disconnect"
'    '                                        UpdFailData(SendCmdInfo.GetInfoLst(0), Nothing)
'    '                                        'Interlocked.Add(FUseSocket, SendCmdInfo.GetInfoLst(0).NeedSocket * -1)
'    '                                    Else
'    '                                        Interlocked.Add(FUseSocket, SendCmdInfo.GetInfoLst(0).NeedSocket)
'    '                                        RunSocketThread(SendCmdInfo.GetInfoLst(0))
'    '                                    End If
'    '                                    'FNagraStatusOK = False
'    '                                End If
'    '                            Else
'    '                                SendCmdInfo.GetInfoLst(0).AllFieldsValue.Item("ERRCODE") = "-9998"
'    '                                SendCmdInfo.GetInfoLst(0).AllFieldsValue.Item("ERRMSG") = "CableSoft_Update_P_Error"
'    '                                UpdFailData(SendCmdInfo.GetInfoLst(0), Nothing)
'    '                            End If
'    '                        Finally
'    '                            SyncLock FSOIndex.Item(SendCmdInfo.GetInfoLst(0).CompId).ProcessingNumber.GetType()
'    '                                FSOIndex.Item(SendCmdInfo.GetInfoLst(0).CompId).ProcessingNumber -= 1
'    '                                FUseCompIndex.Item(SendCmdInfo.GetInfoLst(0).AllFieldsValue.Item("COMPCODE")).ProcessingNumber -= 1
'    '                            End SyncLock
'    '                            SyncLock FSOIndex.Item(SendCmdInfo.GetInfoLst(0).CompId).WaitProcessNumber.GetType()
'    '                                FSOIndex.Item(SendCmdInfo.GetInfoLst(0).CompId).WaitProcessNumber -= 1
'    '                                FUseCompIndex.Item(SendCmdInfo.GetInfoLst(0).AllFieldsValue.Item("COMPCODE")).WaitProcessNumber -= 1
'    '                            End SyncLock
'    '                            UpdateSOStatusUI(FMainFrm, FMainFrm.treLstSO, FSOIndex.Item(SendCmdInfo.GetInfoLst(0).CompId), SOStatus.NotUpdateImg)
'    '                            SendCmdInfo.GetInfoLst.RemoveAt(0)
'    '                        End Try
'    '                    Else
'    '                        CodeSite.Send("SEQNO = " & SendCmdInfo.GetInfoLst(0).SEQNO & ",FUseSocket = " & FUseSocket.ToString)
'    '                    End If
'    '                End If
'    '            Else
'    '                tmrCount.Start()
'    '            End If
'    '        Loop

'    '        'End SyncLock


'    '    Catch ex As Exception
'    '        WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
'    '    Finally

'    '        FNagraConnecting = False
'    '        'tmrCount.Reset()
'    '        'tmrCount.Stop()
'    '    End Try

'    'End Sub
'    Private Shared Sub RunSocketThread(ByVal obj As SOCMDInfo)
'        Try
'            ThreadPool.QueueUserWorkItem(AddressOf SendSocket, obj)
'            'evn.Set()
'            'Dim thr As New Thread(AddressOf SendSocket)
'            'thr.IsBackground = True
'            'thr.Start(obj)
'            FNagraConnecting = True

'        Catch ex As Exception
'            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
'        End Try

'    End Sub
'    '送出1002命令，確保指令可以安全送出
'    'Public Shared Function SendCmd1002() As Boolean


'    '    Try

'    '        Dim aSndLst As List(Of Byte()) = NagraCommand.BuilderCmd(CmdType.Cmd1002, Nothing)
'    '        If (aSndLst Is Nothing) OrElse (aSndLst.Count <> 1) Then
'    '            UpdateSysErrUI(FMainFrm, FMainFrm.TreLstSysErr, New Exception("傳送資料產生有誤"), " 1002命令 失敗")
'    '            WriteErrTxtLog.WriteTxtError(New Exception("傳送資料產生有誤"), Nothing, Nothing)
'    '        End If
'    '        Dim aSndbty() As Byte = aSndLst.Item(0)
'    '        Dim aRcvBty(1023) As Byte

'    '        FSocket.Send(aSndbty)

'    '        'Thread.Sleep(100)

'    '        If FSocket.Receive(aRcvBty) > 0 Then

'    '            Dim objRet As NagraCmdStatus_Type = NagraCommand.NagraStatus(aRcvBty, CmdType.Cmd1002)
'    '            If objRet.Status = AckCode.Ack Then
'    '                Return True
'    '            Else
'    '                UpdateSysErrUI(FMainFrm, FMainFrm.TreLstSysErr, _
'    '                               New Exception("1002 命令有錯"), _
'    '                               "ErrorCode : " & objRet.ErrorCode & " ErrorName : " & objRet.ErrorName)
'    '                Return False
'    '            End If
'    '        End If
'    '        Return True

'    '    Catch ex As Exception
'    '        UpdateSysErrUI(FMainFrm, FMainFrm.TreLstSysErr, ex, " 1002命令 失敗")
'    '        WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
'    '    Finally
'    '        'Erase aSndbty
'    '        'Erase aRcvBty
'    '        'FSocket.Shutdown(SocketShutdown.Receive)
'    '    End Try

'    'End Function
'    '建立Socket 連線
'    'Public Shared Function ConnectSocket() As Boolean

'    '    Try
'    '        If FSocket Is Nothing Then
'    '            FSocket = New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)
'    '            FSocket.SendTimeout = FSndTimeout * 1000
'    '            FSocket.ReceiveTimeout = FRcvTimeout * 1000

'    '        End If

'    '        If Not FSocket.Connected Then

'    '            Try
'    '                FSocket.Shutdown(SocketShutdown.Both)
'    '                FSocket.Disconnect(False)
'    '                FSocket = Nothing
'    '                FSocket = New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)
'    '                FSocket.SendTimeout = FSndTimeout
'    '                FSocket.ReceiveTimeout = FRcvTimeout
'    '            Catch ex As Exception

'    '            End Try
'    '            FLstSysMsg.Clear()
'    '            FLstSysMsg.Add(String.Format("Parent{0} > Nagra主機連接中..", Format(Now, "yyyy/MM/dd hh:mm:ss")))
'    '            FSocket.Connect(FNagraIPAddress, FNagraPort)
'    '            Dim aSend As List(Of Byte()) = NagraCommand.BuilderCmd(CmdType.ConfirmConnect, Nothing)
'    '            If aSend IsNot Nothing AndAlso aSend.Count = 1 Then
'    '                Dim aBty() As Byte = aSend.Item(0)
'    '                FSocket.Send(aBty)
'    '            Else
'    '                FLstSysMsg.Add("主機連接失敗 -- 產生資料失敗")
'    '                UpdateSysErrUI(FMainFrm, FMainFrm.TreLstSysErr, New Exception("Socket 連接失敗"), "產生資料失敗")

'    '                FSocket.Shutdown(SocketShutdown.Both)
'    '                FSocket.Disconnect(False)
'    '                Exit Function
'    '            End If

'    '            Dim aRcv(6) As Byte
'    '            Thread.Sleep(500)
'    '            If FSocket.Receive(aRcv) > 0 Then
'    '                Dim aStatus As NagraStatus_Type = NagraCommand.NagraStatus(aRcv, CmdType.ConfirmConnect)
'    '                If (aStatus.answer_code <> 0) OrElse (aStatus.SUCCESS <> 6) Then

'    '                    Try
'    '                        FLstSysMsg.Add("主機連接失敗")
'    '                        UpdateSysErrUI(FMainFrm, FMainFrm.TreLstSysErr, New Exception("Socket 連接失敗"), " Connect Status : " & aStatus.SUCCESS & " Answer_Code : " & aStatus.answer_code)
'    '                        FSocket.Shutdown(SocketShutdown.Both)
'    '                        FSocket.Disconnect(False)
'    '                    Catch ex As Exception

'    '                    End Try
'    '                    Return False
'    '                    'aSend = String.Empty
'    '                    'aSend = NagraCommand.BuilderCmd(CmdType.Cmd1002)
'    '                    'aBty = Encoding.ASCII.GetBytes(aSend)
'    '                    'FSocket.Send(aBty)
'    '                    ''Y00000000205123420005678920101008100120000000110003000003220000000105200012345678920101108
'    '                    ''E000000003051234200056789201010081000200000001000000000000000000000000
'    '                    'Dim aRcv2(1023) As Byte
'    '                    'If FSocket.Receive(aRcv2) > 0 Then
'    '                    '    Dim aRcvString = Encoding.ASCII.GetString(aRcv2, 1, aRcv2.Length - 1).TrimEnd(Chr(0))
'    '                    'End If
'    '                Else
'    '                    FLstSysMsg.Add("主機連接成功")
'    '                End If

'    '            End If
'    '            UpdateNagraMsgUI(FMainFrm, FMainFrm.TreLstSysMsg, MsgStyle.OKLa)
'    '        End If
'    '        If SendCmd1002() Then
'    '            Return True
'    '        Else
'    '            Return False
'    '        End If
'    '    Catch ex As Exception
'    '        FLstSysMsg.Add("Nagra主機連接失敗")

'    '        UpdateNagraMsgUI(FMainFrm, FMainFrm.TreLstSysMsg, MsgStyle.ErrorLa)
'    '        UpdateSysErrUI(FMainFrm, FMainFrm.TreLstSysErr, ex, " 連接 Nagra Socket 失敗")
'    '        WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
'    '        Try
'    '            FSocket.Shutdown(SocketShutdown.Both)
'    '            FSocket.Disconnect(False)
'    '            FSocket = Nothing
'    '        Catch exClose As Exception
'    '        End Try
'    '    End Try
'    'End Function
'    Private Shared Sub UpdDataOK(ByVal obj As SOCMDInfo, ByVal aLicensesKeyLst As Dictionary(Of String, String))
'        Try
'            SyncLock obj
'                Using cn As New OracleClient.OracleConnection(obj.OraConnectString)
'                    cn.Open()
'                    Using cmd As OracleClient.OracleCommand = cn.CreateCommand
'                        obj.AllFieldsValue.Item("CMDSTATUS") = "S1"
'                        Dim pCmdStatus As New OracleParameter("CMDSTATUS", OracleType.VarChar)
'                        Dim pErrCode As New OracleParameter("ERRCODE", OracleType.VarChar)
'                        Dim pErrMsg As New OracleParameter("ERRMSG", OracleType.VarChar)
'                        Dim pRequestTime As New OracleParameter("REQUESTTIME", OracleType.DateTime)
'                        Dim pResponseTime As New OracleParameter("RESPONSETIME", OracleType.DateTime)
'                        Dim pSeq As New OracleParameter("SEQNO", OracleType.Number)
'                        'Dim pReturnlicenses As New OracleParameter("RETURNLICENSES", OracleType.VarChar)
'                        pCmdStatus.Value = obj.AllFieldsValue.Item("CMDSTATUS")
'                        pErrCode.Value = obj.AllFieldsValue.Item("ERRCODE")
'                        pErrMsg.Value = obj.AllFieldsValue.Item("ERRMSG")
'                        If String.IsNullOrEmpty(obj.AllFieldsValue.Item("REQUESTTIME")) Then
'                            pRequestTime.Value = DBNull.Value
'                        Else
'                            pRequestTime.Value = Date.Parse(obj.AllFieldsValue.Item("REQUESTTIME"))
'                        End If
'                        If String.IsNullOrEmpty(obj.AllFieldsValue.Item("RESPONSETIME")) Then
'                            pResponseTime.Value = DBNull.Value
'                        Else
'                            pResponseTime.Value = Date.Parse(obj.AllFieldsValue.Item("RESPONSETIME"))
'                        End If
'                        pSeq.Value = Convert.ToInt32(obj.SEQNO)
'                        'pReturnlicenses.Value = aLicenses
'                        Dim aSQL As String = "UPDATE SO560 SET CMDSTATUS = :CMDSTATUS," & _
'                                            "ERRCODE = :ERRCODE,ERRMSG = :ERRMSG, " & _
'                                            "REQUESTTIME = :REQUESTTIME,RESPONSETIME = :RESPONSETIME  " & _
'                                            " WHERE SEQNO = :SEQNO "

'                        cmd.CommandText = aSQL
'                        cmd.Parameters.Add(pCmdStatus)
'                        cmd.Parameters.Add(pErrCode)
'                        cmd.Parameters.Add(pErrMsg)
'                        cmd.Parameters.Add(pRequestTime)
'                        cmd.Parameters.Add(pResponseTime)
'                        cmd.Parameters.Add(pSeq)

'                        cmd.ExecuteNonQuery()
'                        UpdSO561B(obj, aLicensesKeyLst)
'                    End Using
'                End Using
'            End SyncLock

'        Catch ex As Exception
'            WriteErrTxtLog.WriteTxtError(ex, "[" & obj.CompName & "] 命令序號 [ " & obj.SEQNO & " ]", Nothing)
'            UpdateSysErrUI(FMainFrm, FMainFrm.TreLstSysErr, ex, " [" & obj.CompName & "] 命令序號 [ " & obj.SEQNO & " ]")
'        Finally
'            UpdateHightCmdUI(FMainFrm, obj, UpdCmdMode.Upd)
'        End Try
'    End Sub
'    Private Shared Function UpdFailData(ByVal obj As SOCMDInfo, ByVal aLicensesKeyLst As Dictionary(Of String, String)) As Boolean
'        Try
'            SyncLock obj
'                Using cn As New OracleClient.OracleConnection(obj.OraConnectString)
'                    cn.Open()

'                    Using cmd As OracleClient.OracleCommand = cn.CreateCommand
'                        If obj.AllFieldsValue.Item("CMDSTATUS") <> "T" Then
'                            obj.AllFieldsValue.Item("CMDSTATUS") = "E"
'                        End If

'                        Dim pCmdStatus As New OracleParameter("CMDSTATUS", OracleType.VarChar)
'                        Dim pErrCode As New OracleParameter("ERRCODE", OracleType.VarChar)
'                        Dim pErrMsg As New OracleParameter("ERRMSG", OracleType.VarChar)
'                        Dim pRequestTime As New OracleParameter("REQUESTTIME", OracleType.DateTime)
'                        Dim pResponseTime As New OracleParameter("RESPONSETIME", OracleType.DateTime)

'                        Dim pSeq As New OracleParameter("SEQNO", OracleType.Number)
'                        pCmdStatus.Value = obj.AllFieldsValue.Item("CMDSTATUS")
'                        pErrCode.Value = obj.AllFieldsValue.Item("ERRCODE")
'                        pErrMsg.Value = obj.AllFieldsValue.Item("ERRMSG")
'                        If String.IsNullOrEmpty(obj.AllFieldsValue.Item("REQUESTTIME")) Then
'                            pRequestTime.Value = DBNull.Value
'                        Else
'                            pRequestTime.Value = Date.Parse(obj.AllFieldsValue.Item("REQUESTTIME"))
'                        End If
'                        If String.IsNullOrEmpty(obj.AllFieldsValue.Item("RESPONSETIME")) Then
'                            pResponseTime.Value = DBNull.Value
'                        Else
'                            pResponseTime.Value = Date.Parse(obj.AllFieldsValue.Item("RESPONSETIME"))
'                        End If
'                        pSeq.Value = Convert.ToInt32(obj.SEQNO)

'                        Dim aSQL As String = "UPDATE SO560 SET CMDSTATUS = :CMDSTATUS," & _
'                                            "ERRCODE = :ERRCODE,ERRMSG = :ERRMSG, " & _
'                                            "REQUESTTIME = :REQUESTTIME,RESPONSETIME = :RESPONSETIME " & _
'                                            " WHERE SEQNO = :SEQNO "
'                        cmd.CommandText = aSQL
'                        cmd.Parameters.Add(pCmdStatus)
'                        cmd.Parameters.Add(pErrCode)
'                        cmd.Parameters.Add(pErrMsg)
'                        cmd.Parameters.Add(pRequestTime)
'                        cmd.Parameters.Add(pResponseTime)
'                        cmd.Parameters.Add(pSeq)
'                        cmd.ExecuteNonQuery()
'                        cmd.Parameters.Clear()
'                        UpdSO561B(obj, aLicensesKeyLst)
'                    End Using
'                End Using
'            End SyncLock

'        Catch ex As Exception
'            WriteErrTxtLog.WriteTxtError(ex, "[" & obj.CompName & "] 命令 [ " & obj.SEQNO & " ]", Nothing)
'            UpdateSysErrUI(FMainFrm, FMainFrm.TreLstSysErr, ex, " [" & obj.CompName & "] 命令 [ " & obj.SEQNO & " ]")
'        Finally
'            'obj.AllFieldsValue.Item("ERRCODE") = aErrCode
'            'obj.AllFieldsValue.Item("ERRMSG") = aErrName
'            'obj.AllFieldsValue.Item("CMDSTATUS") = "E"
'            UpdateHightCmdUI(FMainFrm, obj, UpdCmdMode.Upd)
'        End Try
'    End Function
'    Private Shared Function UpdProced(ByVal obj As SOCMDInfo) As Boolean
'        Try
'            SyncLock obj
'                Using cn As New OracleClient.OracleConnection(obj.OraConnectString)
'                    cn.Open()
'                    Using cmd As OracleClient.OracleCommand = cn.CreateCommand
'                        cmd.CommandText = String.Format("UPDATE SO560 SET RequestTime=to_date('{0}','YYYYMMDDHH24MISS') " & _
'                                                        " WHERE SEQNO ={1} ", _
'                                                        Format(Date.Parse(obj.AllFieldsValue.Item("REQUESTTIME")), "yyyyMMddhhmmss"), _
'                                                        obj.SEQNO)
'                        cmd.ExecuteNonQuery()

'                    End Using
'                End Using
'            End SyncLock
'            Return True
'        Catch ex As Exception
'            WriteErrTxtLog.WriteTxtError(ex, "[" & obj.CompName & "] 命令序號 [ " & obj.SEQNO & " ]", Nothing)
'            UpdateSysErrUI(FMainFrm, FMainFrm.TreLstSysErr, ex, " [" & obj.CompName & "] 命令序號 [ " & obj.SEQNO & " ]")
'        End Try
'    End Function
'    Private Shared Function UpdProcing(ByVal obj As SOCMDInfo) As Boolean
'        Try
'            SyncLock obj
'                SyncLock FSOIndex.Item(obj.CompId).ProcessingNumber.GetType
'                    FSOIndex.Item(obj.CompId).ProcessingNumber += 1
'                    FUseCompIndex.Item(obj.AllFieldsValue.Item("COMPCODE")).ProcessingNumber += 1
'                End SyncLock
'                Using cn As New OracleClient.OracleConnection(obj.OraConnectString)
'                    cn.Open()
'                    Using cmd As OracleClient.OracleCommand = cn.CreateCommand
'                        cmd.CommandText = String.Format("UPDATE SO560 SET CMDSTATUS='P1' " & _
'                                                        " WHERE SEQNO ={0} ", obj.SEQNO)
'                        cmd.ExecuteNonQuery()
'                    End Using
'                End Using
'                obj.AllFieldsValue.Item("CMDSTATUS") = "P1"

'                Return True
'            End SyncLock
'        Catch ex As Exception
'            WriteErrTxtLog.WriteTxtError(ex, "[" & obj.CompName & "] 命令序號 [ " & obj.SEQNO & " ]", Nothing)
'            UpdateSysErrUI(FMainFrm, FMainFrm.TreLstSysErr, ex, " [" & obj.CompName & "] 命令序號 [ " & obj.SEQNO & " ]")
'            Return False
'        Finally
'            UpdateSOStatusUI(FMainFrm, FMainFrm.treLstSO, FSOIndex.Item(obj.CompId), SOStatus.NotUpdateImg)
'            UpdateHightCmdUI(FMainFrm, obj, UpdCmdMode.Add)
'        End Try
'    End Function
'    Private Shared Sub UpdMsgLog(ByVal obj As SOCMDInfo, ByVal blnOK As Boolean)
'        Try
'            If (Not String.IsNullOrEmpty(obj.AllFieldsValue.Item("SOCKETSNDMSG"))) OrElse _
'               (Not String.IsNullOrEmpty(obj.AllFieldsValue.Item("SOCKETRCVMSG"))) Then
'                Try
'                    Using cn As New OracleClient.OracleConnection(obj.OraConnectString)
'                        cn.Open()
'                        Using cmd As OracleClient.OracleCommand = cn.CreateCommand
'                            Dim pSocketSndMsg As New OracleParameter("SOCKETSNDMSG", OracleType.VarChar)
'                            Dim pSocketRcvMsg As New OracleParameter("SOCKETRCVMSG", OracleType.VarChar)
'                            Dim pSeqNo As New OracleParameter("SEQNO", OracleType.Number)
'                            pSocketSndMsg.Value = obj.AllFieldsValue.Item("SOCKETSNDMSG")
'                            pSocketRcvMsg.Value = obj.AllFieldsValue.Item("SOCKETRCVMSG")
'                            pSeqNo.Value = Convert.ToInt32(obj.SEQNO)
'                            Dim aSQL As String = "INSERT INTO SO561 (SEQNO,SOCKETSNDMSG,SOCKETRCVMSG) " & _
'                                                " VALUES ( :SEQNO,:SOCKETSNDMSG,:SOCKETRCVMSG ) "

'                            cmd.CommandText = aSQL
'                            cmd.Parameters.Add(pSeqNo)
'                            cmd.Parameters.Add(pSocketRcvMsg)
'                            cmd.Parameters.Add(pSocketSndMsg)
'                            cmd.ExecuteNonQuery()
'                        End Using
'                    End Using
'                    If blnOK Then
'                        obj.AllFieldsValue.Item("CMDSTATUS") = "S1"
'                    End If
'                Finally
'                    UpdateLowCmdUI(FMainFrm, obj)
'                End Try

'            End If

'        Catch ex As Exception
'            WriteErrTxtLog.WriteTxtError(ex, "[" & obj.CompName & "] 命令序號 [ " & obj.SEQNO & " ]", Nothing)
'            UpdateSysErrUI(FMainFrm, FMainFrm.TreLstSysErr, ex, " [" & obj.CompName & "] 命令序號 [ " & obj.SEQNO & " ]")
'        End Try
'    End Sub

'    Private Shared Function RecvMsg2(ByVal aSocket As Socket, _
'                                    ByVal aCmdInfo As SOCMDInfo, _
'                                    ByVal aCmdType As CmdType, _
'                                    ByRef aNagraStatus As NagraCmdStatus_Type) As Boolean





'        'Dim aRetCmdString As String = String.Empty

'        SyncLock aSocket
'            Try
'                For i As Int32 = 1 To FRcvErrStopCnt
'                    Try
'                        'Dim pos As Int32 = 0
'                        Dim aTimeStop As New Stopwatch()
'                        aTimeStop.Start()
'                        Do
'                            If (aTimeStop.Elapsed.Seconds) >= (FRcvTimeout) Then
'                                Exit Do
'                            End If



'                            'Thread.Sleep(5000)
'                            If SocketCommon.intRcvPos = 0 Then
'                                SocketCommon.RcvBty = Encoding.ASCII.GetBytes(New String(Chr(0), 99))
'                            Else
'                                SocketCommon.RcvBty = Encoding.ASCII.GetBytes(New String(Chr(0), SocketCommon.intRcvPos - 97))
'                            End If

'                            If aSocket.Receive(SocketCommon.RcvBty) > 0 Then
'                                If SocketCommon.intRcvPos = 0 Then
'                                    SocketCommon.intRcvPos = NagraCommand.GetRetunLength(SocketCommon.RcvBty)
'                                End If
'                                If SocketCommon.intRcvPos > 0 Then
'                                    If String.IsNullOrEmpty(SocketCommon.RetCmdString) Then
'                                        SocketCommon.RetCmdString = NagraCommand.GetReturnString(SocketCommon.RcvBty, True)
'                                    Else
'                                        SocketCommon.RetCmdString = SocketCommon.RetCmdString & NagraCommand.GetReturnString(SocketCommon.RcvBty, False)
'                                    End If
'                                    If SocketCommon.intRcvPos <= SocketCommon.RetCmdString.Length Then

'                                        'If Debugger.IsAttached Then
'                                        '    aRetCmdString = aCmdInfo.SEQNO
'                                        '    pos = aRetCmdString.Length
'                                        'End If

'                                        aCmdInfo.AllFieldsValue.Item("RESPONSETIME") = Date.Now
'                                        aCmdInfo.AllFieldsValue.Item("SOCKETRCVMSG") = SocketCommon.RetCmdString
'                                        aNagraStatus = NagraCommand.GetCMDStatus(SocketCommon.RetCmdString, SocketCommon.intRcvPos, aCmdType)
'                                        SocketCommon.intRcvPos = 0
'                                        Return True
'                                    End If
'                                End If
'                            End If
'                        Loop
'                        SocketCommon.intRcvPos = 0
'                        Return False

'                    Catch ex As Exception
'                        FNagraStatusOK = False
'                        UpdateSysErrUI(FMainFrm, FMainFrm.TreLstSysErr, ex, " [" & aCmdInfo.CompName & "] 命令序號 [ " & aCmdInfo.SEQNO & " ] 第 " & i & " 次接收")
'                        WriteErrTxtLog.WriteTxtError(ex, " [" & aCmdInfo.CompName & "] 命令序號 [ " & aCmdInfo.SEQNO & " ] 第 " & i & " 次接收", Nothing)
'                        If i = FRcvErrStopCnt Then
'                            Return False
'                        End If
'                    End Try
'                Next
'            Catch ex As Exception
'                FNagraStatusOK = False
'                UpdateSysErrUI(FMainFrm, FMainFrm.TreLstSysErr, ex, " [" & aCmdInfo.CompName & "] 命令序號 [ " & aCmdInfo.SEQNO & " ]")
'                WriteErrTxtLog.WriteTxtError(ex, " [" & aCmdInfo.CompName & "] 命令序號 [ " & aCmdInfo.SEQNO & " ]", Nothing)
'                Return False
'            End Try
'        End SyncLock
'    End Function


'    Private Shared Function RecvMsg(ByVal aSocket As Socket, _
'                                    ByVal aCmdInfo As SOCMDInfo, _
'                                    ByVal aCmdType As CmdType, _
'                                    ByRef aNagraStatus As NagraCmdStatus_Type) As Boolean





'        Dim aRetCmdString As String = String.Empty

'        Try
'            For i As Int32 = 1 To FRcvErrStopCnt
'                Try
'                    Dim pos As Int32 = 0
'                    Dim aTimeStop As New Stopwatch()
'                    aTimeStop.Start()
'                    Do
'                        If (aTimeStop.Elapsed.Seconds) >= (FRcvTimeout) Then
'                            Exit Do
'                        End If

'                        Dim aRcvBty(1023) As Byte

'                        'Thread.Sleep(3000)

'                        'If pos = 0 Then
'                        '    ReDim aRcvBty(99)
'                        'Else
'                        '    ReDim aRcvBty(pos - 100)
'                        'End If



'                        If aSocket.Receive(aRcvBty) > 0 Then

'                            If pos = 0 Then
'                                pos = NagraCommand.GetRetunLength(aRcvBty)
'                            End If
'                            If pos > 0 Then
'                                If String.IsNullOrEmpty(aRetCmdString) Then
'                                    aRetCmdString = NagraCommand.GetReturnString(aRcvBty, True)
'                                Else
'                                    aRetCmdString = aRetCmdString & NagraCommand.GetReturnString(aRcvBty, False)
'                                End If
'                                If pos <= aRetCmdString.Length Then

'                                    'If Debugger.IsAttached Then
'                                    '    aRetCmdString = aCmdInfo.SEQNO
'                                    '    pos = aRetCmdString.Length
'                                    'End If

'                                    aCmdInfo.AllFieldsValue.Item("RESPONSETIME") = Date.Now
'                                    aCmdInfo.AllFieldsValue.Item("SOCKETRCVMSG") = aRetCmdString
'                                    aNagraStatus = NagraCommand.GetCMDStatus(aRetCmdString, pos, aCmdType)
'                                    Return True
'                                End If
'                            End If
'                        End If
'                    Loop
'                    Return False


'                Catch ex As Exception
'                    FNagraStatusOK = False
'                    UpdateSysErrUI(FMainFrm, FMainFrm.TreLstSysErr, ex, " [" & aCmdInfo.CompName & "] 命令序號 [ " & aCmdInfo.SEQNO & " ] 第 " & i & " 次接收")
'                    WriteErrTxtLog.WriteTxtError(ex, " [" & aCmdInfo.CompName & "] 命令序號 [ " & aCmdInfo.SEQNO & " ] 第 " & i & " 次接收", Nothing)
'                    If i = FRcvErrStopCnt Then
'                        Return False
'                    End If
'                End Try
'            Next
'        Catch ex As Exception
'            FNagraStatusOK = False
'            UpdateSysErrUI(FMainFrm, FMainFrm.TreLstSysErr, ex, " [" & aCmdInfo.CompName & "] 命令序號 [ " & aCmdInfo.SEQNO & " ]")
'            WriteErrTxtLog.WriteTxtError(ex, " [" & aCmdInfo.CompName & "] 命令序號 [ " & aCmdInfo.SEQNO & " ]", Nothing)
'            Return False
'        End Try

'    End Function
'    Private Shared Function SendMsg(ByVal aSocket As Socket, _
'                                    ByVal aCMDInfo As SOCMDInfo, _
'                                    ByVal aCmdType As CmdType, _
'                                    ByVal aMsgBty() As Byte) As Boolean



'        Try
'            For i As Int32 = 1 To FSndErrStopCnt
'                Try
'                    aCMDInfo.AllFieldsValue.Item("SOCKETSNDMSG") = NagraCommand.GetReturnString(aMsgBty, True)
'                    aCMDInfo.AllFieldsValue.Item("REQUESTTIME") = Now
'                    UpdProced(aCMDInfo)
'                    aSocket.Send(aMsgBty)

'                    Return True
'                Catch ex As Exception
'                    UpdateSysErrUI(FMainFrm, FMainFrm.TreLstSysErr, ex, " [" & aCMDInfo.CompName & "] 命令序號 [ " & aCMDInfo.SEQNO & " ] 第 " & i & " 次傳送")
'                    WriteErrTxtLog.WriteTxtError(ex, " [" & aCMDInfo.CompName & "] 命令序號 [ " & aCMDInfo.SEQNO & " ] 第 " & i & " 次傳送", Nothing)
'                    If i = FSndErrStopCnt Then
'                        Return False
'                    End If
'                End Try

'            Next
'        Catch ex As Exception
'            FNagraStatusOK = False
'            UpdateSysErrUI(FMainFrm, FMainFrm.TreLstSysErr, ex, " [" & aCMDInfo.CompName & "] 命令序號 [ " & aCMDInfo.SEQNO & " ]")
'            WriteErrTxtLog.WriteTxtError(ex, " [" & aCMDInfo.CompName & "] 命令序號 [ " & aCMDInfo.SEQNO & " ]", Nothing)
'        End Try
'    End Function
'    Private Shared Sub UpdSO561B(ByVal aCMDInFo As SOCMDInfo, ByVal aLicensesKeyLst As Dictionary(Of String, String))

'        Try
'            If aLicensesKeyLst Is Nothing OrElse aLicensesKeyLst.Count <= 0 Then
'                Exit Sub
'            End If
'            SyncLock aCMDInFo

'                Using cn As New OracleClient.OracleConnection(aCMDInFo.OraConnectString)
'                    cn.Open()
'                    Dim tra As OracleTransaction = cn.BeginTransaction
'                    Dim strAryNotes() As String = aCMDInFo.AllFieldsValue("NOTES").Split(SplitNoteChar)
'                    Using cmd As OracleClient.OracleCommand = cn.CreateCommand
'                        cmd.Transaction = tra
'                        For i As Int32 = 0 To aLicensesKeyLst.Count - 1


'                            Dim aGetSeq As Int32 = Convert.ToInt32(aLicensesKeyLst.Keys.ToList.Item(i))
'                            Dim strAry() As String = strAryNotes.GetValue(aGetSeq).ToString.Split("~")
'                            Dim aProduct As String = strAry.GetValue(0)
'                            Dim aAsset As String = strAry.GetValue(1)
'                            Dim aSQL As String = String.Format("UPDATE SO561B SET LicensesKey='{0}'" & _
'                                                                " WHERE SEQNO = {1} AND ProductID = '{2}'" & _
'                                                                " AND VODItem = '{3}' ", aLicensesKeyLst.Item(i.ToString) _
'                                                                , aCMDInFo.SEQNO, aProduct, aAsset)
'                            cmd.CommandText = aSQL
'                            cmd.ExecuteNonQuery()
'                        Next
'                    End Using
'                    tra.Commit()
'                End Using
'                'obj.AllFieldsValue.Item("CMDSTATUS") = "P1"
'            End SyncLock

'        Catch ex As Exception
'            UpdateSysErrUI(FMainFrm, FMainFrm.TreLstSysErr, ex, " [" & aCMDInFo.CompName & "] 命令序號 [ " & aCMDInFo.SEQNO & " ]")
'            WriteErrTxtLog.WriteTxtError(ex, " [" & aCMDInFo.CompName & "] 命令序號 [ " & aCMDInFo.SEQNO & " ]", Nothing)
'        End Try
'    End Sub
'    'Public Shared Sub Read_Callback(ByVal ar As IAsyncResult)
'    '    Dim so As StateObject = CType(ar.AsyncState, StateObject)
'    '    Dim s As Socket = so.workSocket

'    '    Dim read As Integer = s.EndReceive(ar)

'    '    If read > 0 Then
'    '        so.sb.Append(Encoding.ASCII.GetString(so.buffer, 0, read))
'    '        s.BeginReceive(so.buffer, 0, StateObject.BUFFER_SIZE, 0, New AsyncCallback(AddressOf Read_Callback), so)
'    '    Else
'    '        If so.sb.Length > 1 Then
'    '            'All the data has been read, so displays it to the console
'    '            Dim strContent As String
'    '            strContent = so.sb.ToString()
'    '            'Console.WriteLine([String].Format("Read {0} byte from socket" + "data = {1} ", strContent.Length, strContent))
'    '        End If
'    '        's.Close()
'    '    End If
'    'End Sub 'Read_Callback


'    Private Shared Sub SendSocket(ByVal obj As SOCMDInfo)


'        Try
'            Dim aClientSocket As Socket = FSocket
'            Dim aType As CmdType
'            'Dim pos As Int32 = 0
'            Select Case obj.CMD.ToUpper
'                Case "B1"
'                    aType = CmdType.Cmd0851
'                Case Else
'                    aType = CmdType.OtherCMD
'            End Select
'            If aType = CmdType.OtherCMD Then
'                Exit Sub
'            End If

'            '取出要送出的低階命令
'            Dim aryMsg As List(Of Byte()) = NagraCommand.BuilderCmd(aType, obj.AllFieldsValue)
'            If (aryMsg Is Nothing) OrElse (aryMsg.Count = 0) Then
'                FNagraStatusOK = False
'                UpdateSysErrUI(FMainFrm, FMainFrm.TreLstSysErr, New Exception("0851產生資料失敗"), " [" & obj.CompName & "] 命令序號 [ " & obj.SEQNO & " ]")
'                WriteErrTxtLog.WriteTxtError(New Exception(aType & "產生資料失敗"), " [" & obj.CompName & "] 命令序號 [ " & obj.SEQNO & " ]", Nothing)
'                Exit Sub
'            End If
'            Dim aAllOK As Boolean = False
'            Dim objNagraStatus As New NagraCmdStatus_Type
'            Dim aEntities As String = String.Empty
'            Dim aEntity_data As String = String.Empty

'            '初始化回應的訊息
'            objNagraStatus.Status = AckCode.Nack
'            objNagraStatus.ErrorCode = "-9999"
'            objNagraStatus.ErrorName = "CableSoft_Initial"
'            objNagraStatus.Entities = String.Empty
'            objNagraStatus.Entity_data = String.Empty
'            aAllOK = True
'            Dim aLicensesKeyLst As New Dictionary(Of String, String)
'            '只要其中一個低階命令失敗，就算失敗

'            For i As Int32 = 0 To aryMsg.Count - 1

'                Try

'                    If Not SendMsg(aClientSocket, obj, aType, aryMsg.Item(i)) Then
'                        objNagraStatus.ErrorCode = "-9995"
'                        objNagraStatus.ErrorName = "CableSoft_SendCMD_Error"
'                        UpdFailData(obj, Nothing)
'                        aAllOK = False
'                        Exit For
'                    End If

'                    'Dim so2 As New StateObject()
'                    'so2.workSocket = FSocket
'                    'FSocket.BeginReceive(so2.buffer, 0, StateObject.BUFFER_SIZE, 0, New AsyncCallback(AddressOf Read_Callback), so2)
'                    'Exit Sub

'                    If aAllOK Then
'                        'FSocket.BeginReceive


'                        If Not RecvMsg(aClientSocket, obj, aType, objNagraStatus) Then
'                            objNagraStatus.ErrorCode = "-9994"
'                            objNagraStatus.ErrorName = "CableSoft_RecvCMD_Error"
'                            obj.AllFieldsValue.Item("CMDSTATUS") = "T"
'                            UpdFailData(obj, Nothing)
'                            aAllOK = False
'                        Else

'                            If (objNagraStatus.Status = AckCode.Nack) OrElse (objNagraStatus.Status = AckCode.Other) Then
'                                aAllOK = False
'                            Else
'                                If Not String.IsNullOrEmpty(objNagraStatus.Entities) Then
'                                    If String.IsNullOrEmpty(aEntities) Then
'                                        aEntities = objNagraStatus.Entities
'                                    Else
'                                        aEntities = aEntities & "," & objNagraStatus.Entities
'                                    End If
'                                End If
'                                If Not String.IsNullOrEmpty(objNagraStatus.Entity_data) Then
'                                    If String.IsNullOrEmpty(aEntity_data) Then
'                                        aEntity_data = objNagraStatus.Entity_data
'                                    Else
'                                        aEntity_data = aEntity_data & "," & objNagraStatus.Entity_data
'                                    End If
'                                    aLicensesKeyLst.Add(i.ToString, objNagraStatus.Entity_data)
'                                End If
'                                obj.AllFieldsValue.Item("Returnlicenses".ToUpper) = aEntity_data

'                            End If
'                        End If
'                        If Not aAllOK Then
'                            Exit For
'                        End If
'                    Else
'                        Exit For
'                    End If
'                Finally
'                    UpdMsgLog(obj, aAllOK)
'                End Try
'            Next

'            If (objNagraStatus.Status = AckCode.Nack) OrElse (objNagraStatus.Status = AckCode.Other) Then
'                obj.AllFieldsValue.Item("ERRCODE") = objNagraStatus.ErrorCode
'                obj.AllFieldsValue.Item("ERRMSG") = objNagraStatus.ErrorName

'                UpdFailData(obj, aLicensesKeyLst)
'            Else
'                If (aEntity_data.Split(SplitReturnChar).Length) <> _
'                     (obj.AllFieldsValue.Item("NOTES").Split(SplitNoteChar).Length) Then
'                    objNagraStatus.Status = AckCode.Other
'                    objNagraStatus.ErrorCode = "-9991"
'                    objNagraStatus.ErrorName = "CableSoft_RetLengeth_Err"

'                End If
'                objNagraStatus.ErrorCode = String.Empty
'                objNagraStatus.ErrorName = String.Empty
'                obj.AllFieldsValue.Item("ERRCODE") = String.Empty
'                obj.AllFieldsValue.Item("ERRMSG") = String.Empty
'                UpdDataOK(obj, aLicensesKeyLst)
'            End If

'        Catch ex As Exception
'            FNagraStatusOK = False
'            UpdateSysErrUI(FMainFrm, FMainFrm.TreLstSysErr, ex, " [" & obj.CompName & "] 命令序號 [ " & obj.SEQNO & " ]")
'            WriteErrTxtLog.WriteTxtError(ex, " [" & obj.CompName & "] 命令序號 [ " & obj.SEQNO & " ]", Nothing)

'        Finally
'            UpdateHightCmdUI(FMainFrm, obj, UpdCmdMode.Upd)
'            Interlocked.Add(FUseSocket, (obj.NeedSocket * -1))

'            Thread.Sleep(TimeSpan.FromSeconds(FSndDelayTime))


'        End Try


'    End Sub
'End Class
'Public Class SendCmdInfo

'    Private Shared SOCMDInfoLst As New List(Of SOCMDInfo)

'    Public Shared Function ADD(ByVal aCompID As String, ByVal aCompName As String, _
'                            ByVal aSEQNO As String, _
'                            ByVal aCMDName As String, _
'                            ByVal aConnectString As String, _
'                            ByVal aDataReader As OracleClient.OracleDataReader) As SOCMDInfo


'        Try
'            rwl.EnterWriteLock()
'            SyncLock aDataReader
'                Dim obj As New SOCMDInfo()
'                Dim objList As New Dictionary(Of String, String)
'                For i As Int32 = 0 To aDataReader.FieldCount - 1
'                    If aDataReader.IsDBNull(i) Then
'                        objList.Add(aDataReader.GetName(i).ToUpper, String.Empty)
'                    Else
'                        objList.Add(aDataReader.GetName(i).ToUpper, aDataReader.Item(i))
'                    End If
'                Next
'                objList.Add("SOCKETSNDMSG", String.Empty)
'                objList.Add("SOCKETRCVMSG", String.Empty)
'                objList.Add("GUIDNO", String.Empty)
'                objList.Add("TranNum", String.Empty)
'                If Not objList.ContainsKey("Returnlicenses".ToUpper) Then
'                    objList.Add("Returnlicenses".ToUpper, String.Empty)
'                End If
'                obj.CompId = aCompID
'                obj.SEQNO = aSEQNO
'                obj.CMD = aCMDName
'                obj.CompName = aCompName
'                obj.OraConnectString = aConnectString
'                obj.CmdPriority = 9
'                obj.AllFieldsValue = objList
'                SOCMDInfoLst.Add(obj)
'                Return obj
'            End SyncLock
'        Catch ex As Exception

'            Cablesoft.GateWay.csException.WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
'            Return Nothing
'        Finally
'            rwl.ExitWriteLock()
'        End Try
'    End Function

'    Public Shared Function GetInfoLst() As List(Of SOCMDInfo)
'        Return SOCMDInfoLst
'    End Function

'    Private Shared Function CompareCmdPriority(ByVal p1 As SOCMDInfo, ByVal p2 As SOCMDInfo) As Int32
'        Try
'            Return Comparer(Of Int32).Default.Compare(p1.SEQNO, p2.SEQNO)
'        Catch ex As Exception
'            Return 0
'        End Try
'    End Function


'End Class
'Public Class SocketCommon
'    Public Sub New()

'    End Sub
'    Private Shared _intRcvPos As Int32
'    Private Shared _RetCmdString As String
'    Private Shared _RcvBty() As Byte
'    Public Shared Property intRcvPos() As Int32
'        Get
'            Return _intRcvPos
'        End Get
'        Set(ByVal value As Int32)
'            _intRcvPos = value
'        End Set
'    End Property
'    Public Shared Property RetCmdString() As String
'        Get
'            Return _RetCmdString
'        End Get
'        Set(ByVal value As String)
'            _RetCmdString = value
'        End Set
'    End Property
'    Public Shared Property RcvBty() As Byte()
'        Get
'            Return _RcvBty
'        End Get
'        Set(ByVal value As Byte())
'            _RcvBty = value
'        End Set
'    End Property
'End Class

'Public Class SOCMDInfo

'    Protected Friend Sub New()

'    End Sub
'    Private _CompId As String
'    Private _SEQNO As String
'    Private _CMD As String
'    Private _OraConnectString As String
'    Private _NeedSocket As Int32 = 1
'    Private _CompName As String
'    Private _CmdPriority As Int32

'    Private _AllFieldsValue As Dictionary(Of String, String)
'    Public Property CompId() As String
'        Get
'            Return _CompId

'        End Get
'        Set(ByVal value As String)
'            _CompId = value
'        End Set
'    End Property

'    Public Property SEQNO() As String
'        Get
'            Return _SEQNO
'        End Get
'        Set(ByVal value As String)
'            _SEQNO = value
'        End Set
'    End Property
'    Public Property CMD() As String
'        Get
'            Return _CMD
'        End Get
'        Set(ByVal value As String)
'            _CMD = value
'        End Set
'    End Property
'    Public Property OraConnectString()
'        Get
'            Return _OraConnectString
'        End Get
'        Set(ByVal value)
'            _OraConnectString = value
'        End Set
'    End Property
'    Public Property NeedSocket() As Int32
'        Get
'            Return _NeedSocket
'        End Get
'        Set(ByVal value As Int32)
'            _NeedSocket = CMD
'        End Set
'    End Property
'    Public Property CompName() As String
'        Get
'            Return _CompName
'        End Get
'        Set(ByVal value As String)
'            _CompName = value
'        End Set
'    End Property
'    Public Property CmdPriority() As Integer
'        Get
'            Return _CmdPriority
'        End Get
'        Set(ByVal value As Integer)
'            _CmdPriority = value
'        End Set
'    End Property

'    Public Property AllFieldsValue() As Dictionary(Of String, String)
'        Get
'            Return _AllFieldsValue
'        End Get
'        Set(ByVal value As Dictionary(Of String, String))
'            _AllFieldsValue = value
'        End Set
'    End Property
'End Class