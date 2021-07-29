Imports System
Imports System.IO
Imports System.Text
Imports System.Net
Imports System.Net.Sockets
Imports CableSoft.Gateway.csException
Imports System.Threading
Imports CableSoft_Nagra_BuilderCmd
Public Class Nagra_Socket
    Public Shared FMaxSocket As Int32 = 1
    Public Shared FUseSocket As Int32 = 0
    Private Shared FSocket As Socket = Nothing
    Public Shared FMainFrm As rfrmMain
    Private Const FSocketName As String = "SMS_GWY"
    Public Shared tmrCount As New Stopwatch


    Public Shared Sub SendCMD()
        FNagraConnecting = True
        If FThreadStop Then
            SendCmdInfo.GetInfoLst.Clear()
            Exit Sub
        End If
        If Not tmrCount.IsRunning Then
            tmrCount.Start()
        End If

        Try

            SyncLock SendCmdInfo.GetInfoLst

                Do While Not SendCmdInfo.GetInfoLst.Count <= 0
                    If FThreadStop Then
                        SendCmdInfo.GetInfoLst.Clear()
                        Exit Do
                    End If
                    Thread.Sleep(TimeSpan.FromSeconds(FSndDelayTime))
                    If tmrCount.IsRunning Then

                        If tmrCount.Elapsed.Seconds >= FCheckStatusTime AndAlso 1 = 1 Then
                            Try
                                Interlocked.Increment(FUseSocket)
                                If FUseSocket <= FMaxSocket Then
                                    FNagraStatusOK = ConnectSocket()
                                End If
                            Finally
                                Interlocked.Decrement(FUseSocket)
                                tmrCount.Reset()
                                tmrCount.Start()
                            End Try
                            If FNagraStatusOK Then
                                If FSocket.Connected Then
                                    Interlocked.Add(FUseSocket, SendCmdInfo.GetInfoLst(0).NeedSocket)

                                    If FUseSocket <= FMaxSocket Then
                                        '一次只能送一個Socket
                                        If FNagraConnecting Then

                                            Dim thr As New Thread(AddressOf SendSocket)
                                            thr.IsBackground = True
                                            thr.Start(SendCmdInfo.GetInfoLst(0))
                                            SendCmdInfo.GetInfoLst.RemoveAt(0)
                                        Else
                                            FNagraConnecting = True
                                            Interlocked.Add(FUseSocket, (SendCmdInfo.GetInfoLst(0).NeedSocket * -1))

                                        End If
                                    Else
                                        Interlocked.Add(FUseSocket, (SendCmdInfo.GetInfoLst(0).NeedSocket * -1))

                                    End If

                                End If
                            Else
                                FNagraStatusOK = False
                            End If



                        End If
                    Else
                        tmrCount.Start()
                    End If







                    'For i As Int32 = 0 To SendCmdInfo.GetInfoLst.Count - 1
                    '    If (SendCmdInfo.GetInfoLst(i).NeedSocket + FUseSocket) <= (FMaxSocket) Then

                    '        SyncLock FUseSocket.GetType
                    '            FUseSocket += SendCmdInfo.GetInfoLst(i).NeedSocket
                    '        End SyncLock

                    '        Dim thr As New Thread(AddressOf SendSocket)
                    '        thr.IsBackground = True
                    '        'SendSocket("公司別:" & SendCmdInfo.GetInfoLst(i).CompId & _
                    '        '          " 序號 : " & SendCmdInfo.GetInfoLst(i).SEQNO & _
                    '        '          " 命令 : " & SendCmdInfo.GetInfoLst(i).CMD)
                    '        '一次只能送一個Socket
                    '        SyncLock FSocketCanSend.GetType
                    '            If FSocketCanSend Then
                    '                FSocketCanSend = False
                    '                thr.Start(SendCmdInfo.GetInfoLst(i))
                    '                SendCmdInfo.GetInfoLst.RemoveAt(i)
                    '            End If
                    '        End SyncLock
                    '        'thr.Start("公司別:" & SendCmdInfo.GetInfoLst(i).CompId & _
                    '        '          " 序號 : " & SendCmdInfo.GetInfoLst(i).SEQNO & _
                    '        '          " 命令 : " & SendCmdInfo.GetInfoLst(i).CMD)

                    '    End If
                    'Next

                Loop

            End SyncLock


        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
        Finally

            FNagraConnecting = False
          
        End Try

    End Sub
    '送出1002命令，確保指令可以安全送出
    Public Shared Function SendCmd1002() As Boolean

        Dim aSndbty() As Byte = Encoding.ASCII.GetBytes(NagraCommand.BuilderCmd(CmdType.Cmd1002))
        Dim aRcvBty(255) As Byte
        Try

            FSocket.Send(aSndbty)
            Thread.Sleep(100)
            If FSocket.Receive(aRcvBty) > 0 Then
                Dim objRet As NagraCmdStatus_Type = NagraCommand.NagraStatus(aRcvBty, CmdType.Cmd1002)
                If objRet.Status = AckCode.Ack Then
                    Return True
                Else
                    UpdateSysErrUI(FMainFrm, FMainFrm.TreLstSysErr, _
                                   New Exception("1002 命令有錯"), _
                                   "ErrorCode : " & objRet.ErrorCode & " ErrorName : " & objRet.ErrorName)
                    Return False
                End If
            End If
            Return True

        Catch ex As Exception
            UpdateSysErrUI(FMainFrm, FMainFrm.TreLstSysErr, ex, " 1002命令 失敗")
            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
        Finally
            Erase aSndbty
            Erase aRcvBty
            'FSocket.Shutdown(SocketShutdown.Receive)
        End Try

    End Function
    '建立Socket 連線
    Public Shared Function ConnectSocket() As Boolean


        Try
            If FSocket Is Nothing Then
                FSocket = New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)
                FSocket.SendTimeout = 1000
                FSocket.ReceiveTimeout = 1000
            End If

            If Not FSocket.Connected Then

                Try
                    FSocket.Shutdown(SocketShutdown.Both)
                    FSocket.Disconnect(False)
                    FSocket = Nothing
                    FSocket = New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)
                    FSocket.SendTimeout = 3000
                    FSocket.ReceiveTimeout = 3000
                Catch ex As Exception

                End Try
                FSocket.Connect(FNagraIPAddress, FNagraPort)
                Dim aSend As String = NagraCommand.BuilderCmd(CmdType.ConfirmConnect)
                Dim aBty() As Byte = Encoding.ASCII.GetBytes(aSend)
                FSocket.Send(aBty)

                Dim aRcv(255) As Byte
                Thread.Sleep(100)
                If FSocket.Receive(aRcv) > 0 Then
                    Dim aStatus As NagraStatus_Type = NagraCommand.NagraStatus(aRcv, CmdType.ConfirmConnect)
                    If aStatus.answer_code = 0 And aStatus.SUCCESS = 6 Then


                    Else
                        Try
                            UpdateSysErrUI(FMainFrm, FMainFrm.TreLstSysErr, New Exception("Socket 連接失敗"), " Connect Status : " & aStatus.SUCCESS & " Answer_Code : " & aStatus.answer_code)
                            FSocket.Shutdown(SocketShutdown.Both)
                            FSocket.Disconnect(False)
                        Catch ex As Exception

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
                    End If

                End If



            End If

            If SendCmd1002() Then
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            UpdateSysErrUI(FMainFrm, FMainFrm.TreLstSysErr, ex, " 連接 Nagra Socket 失敗")
            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
            Try
                FSocket.Shutdown(SocketShutdown.Both)
                FSocket.Disconnect(False)
                FSocket = Nothing
            Catch exClose As Exception

            End Try

        End Try

    End Function


    Private Shared Sub SendSocket(ByVal obj As SOCMDInfo)
        'SyncLock FSOIndex.Item(obj.CompId).WaitProcessNumber.GetType
        '    FSOIndex.Item(obj.CompId).WaitProcessNumber -= 1
        'End SyncLock
        'SyncLock FSOIndex.Item(obj.CompId).ProcessingNumber.GetType
        '    FSOIndex.Item(obj.CompId).ProcessingNumber -= 1
        'End SyncLock


        'UpdateSOStatusUI(FMainFrm, FMainFrm.treLstSO, FSOIndex.Item(obj.CompId), SOStatus.NotUpdateImg)
        'Exit Sub
        SyncLock obj


            Dim aMsg As String = "公司別 : " & obj.CompId & _
                            " 序號 : " & obj.SEQNO & _
                            " 命令 : " & obj.CMD
            Try
                Dim btySend() As Byte = Encoding.Default.GetBytes(aMsg)




                FSocket.Send(btySend)

                Dim aRcv(1024) As Byte
                If FSocket.Receive(aRcv) > 0 Then

                End If
                'Thread.Sleep(3000)
            

                Using cn As New OracleClient.OracleConnection(obj.OraConnectString)
                    cn.Open()
                    Using cmd As OracleClient.OracleCommand = cn.CreateCommand
                        cmd.CommandText = String.Format("UPDATE SO555 SET CMDSTATUS='P' " & _
                                                        " WHERE SEQNO ={0} ", obj.SEQNO)
                        cmd.ExecuteNonQuery()
                    End Using
                End Using
            
            Catch ex As Exception
                UpdateSysErrUI(FMainFrm, FMainFrm.TreLstSysErr, ex, " [" & obj.CompName & "] 命令 [ " & obj.SEQNO & " ]")
                WriteErrTxtLog.WriteTxtError(ex, " [" & obj.CompName & "] 命令 [ " & obj.SEQNO & " ]", Nothing)

            Finally



                SyncLock FSOIndex.Item(obj.CompId).WaitProcessNumber.GetType
                    FSOIndex.Item(obj.CompId).WaitProcessNumber -= 1
                End SyncLock
                SyncLock FSOIndex.Item(obj.CompId).ProcessingNumber.GetType
                    FSOIndex.Item(obj.CompId).ProcessingNumber -= 1
                End SyncLock


                UpdateSOStatusUI(FMainFrm, FMainFrm.treLstSO, FSOIndex.Item(obj.CompId), SOStatus.NotUpdateImg)
                

                Interlocked.Add(FUseSocket, (obj.NeedSocket * -1))
            End Try

        End SyncLock
    End Sub
End Class
Public Class SendCmdInfo

    Private Shared SOCMDInfoLst As New List(Of SOCMDInfo)
    Public Shared Sub ADD(ByVal aCompID As String, ByVal aCompName As String, _
                            ByVal aSEQNO As String, _
                            ByVal aCMDName As String, _
                            ByVal aConnectString As String)

        Try
            rwl.EnterWriteLock()
            Dim obj As New SOCMDInfo()
            obj.CompId = aCompID
            obj.SEQNO = aSEQNO
            obj.CMD = aCMDName
            obj.CompName = aCompName
            obj.OraConnectString = aConnectString
            obj.CmdPriority = 9

            Dim o As New List(Of String)
            SOCMDInfoLst.Add(obj)
            SOCMDInfoLst.Sort(AddressOf CompareCmdPriority)

        Catch ex As Exception
            CableSoft.GateWay.csException.WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
        Finally
            rwl.ExitWriteLock()
        End Try
    End Sub

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
End Class