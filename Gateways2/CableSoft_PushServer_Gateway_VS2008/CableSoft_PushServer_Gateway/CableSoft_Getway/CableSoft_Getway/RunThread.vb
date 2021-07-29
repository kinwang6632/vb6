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
Imports System.Net
Imports Newtonsoft.Json
Module RunThread
    Private ti As RegisteredWaitHandle = Nothing
    Public evn As New AutoResetEvent(False)
    Private FMainForm As rfrmMain
    Public FThreadsTotal As Int32 = 0
    
    Public FSOIndex As New System.Collections.Generic.Dictionary(Of String, SOInfoClass)
    Public FUseCompIndex As New Dictionary(Of String, UseCompanyClass)
    Private FExeCmdSQL As String
    Private aSQLCmd As String = "SELECT * FROM SO560 WHERE ( CMDSTATUS = 'S1' )  " & _
                        " AND ROWNUM<={0}  AND COMPCODE IN ( {1} ) " & _
                        " AND AUTHTIME IS NULL ORDER BY SEQNO "

    
    Private Const SplitReturnlicenses As Char = ","c
    Private Const SplitNoteChar As Char = ","c
    Public Sub RunGateway(ByVal aParent As Form)
        Try

            FMainForm = aParent
            

            If ThreadInitial() Then
                Try
                   

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
            aSQLCmd = String.Format(aSQLCmd, FProcessNumber, FUseCompWhere)
            FSOIndex.Clear()
            For i As Int32 = 0 To FSOInfoList.Count - 1
                FSOIndex.Add(FSOInfoList(i).CompID, FSOInfoList.Item(i))
            Next
            FUseCompIndex.Clear()
            For i As Int32 = 0 To FUseCompList.Count - 1
                FUseCompIndex.Add(FUseCompList(i).CompID, FUseCompList(i))
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
                
                For i As Int32 = 0 To obj.Count - 1
                    ThreadPool.QueueUserWorkItem(AddressOf ThreadCmd, obj.Item(i))
                Next

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

            If (FThreadStop) OrElse (FThreadsTotal > 0) Then
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
                        WriteErrTxtLog.WriteSOStatus(obj.CompName & " ---> 連線後已斷線")
                    End If
                    obj.ConnectionOK = False

                Else
                    aSytle = SOStatus.Yes
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
                UpdateSODatabase(FMainForm, aId, obj, aSytle)
                UpdateSOStatusUI(FMainForm, FMainForm.treLstSO, obj, aSytle)
                If (obj.WaitProcessNumber <= 0) AndAlso _
                    (obj.ConnectionOK) Then
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

                'Interlocked.Increment(FThreadsTotal)
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
                                Dim aryFieldLst As New Dictionary(Of String, String)
                                aryFieldLst.Clear()
                                For i As Int32 = 0 To r.FieldCount - 1
                                    If r.IsDBNull(i) Then
                                        aryFieldLst.Add(r.GetName(i).ToUpper, String.Empty)
                                    Else
                                        aryFieldLst.Add(r.GetName(i).ToUpper, r.Item(i).ToString)
                                    End If
                                Next
                                aryFieldLst.Add("GUIDNO", String.Empty)
                                aryFieldLst.Add("PushServerUrl".ToUpper, String.Empty)
                                aryFieldLst.Add("PushServerCode".ToUpper, String.Empty)
                                aryFieldLst.Add("PushServerMsg".ToUpper, String.Empty)
                                aryFieldLst.Add("PushServerHeadMsg".ToUpper, String.Empty)
                                If Not aryFieldLst.ContainsKey("Returnlicenses".ToUpper) Then
                                    aryFieldLst.Add("Returnlicenses".ToUpper, String.Empty)
                                End If
                                Dim aSQLLicenses As String = String.Format("SELECT LicensesKey FROM SO561B WHERE SEQNO = {0}", _
                                                                    r.Item("SEQNO"))

                                Using cmd2 As New OracleCommand(aSQLLicenses, obj.OracleCn)
                                    Using r2 As OracleDataReader = cmd2.ExecuteReader()
                                        While r2.Read
                                            If String.IsNullOrEmpty(aryFieldLst.Item("Returnlicenses".ToUpper)) Then
                                                aryFieldLst.Item("Returnlicenses".ToUpper) = r2("LicensesKey")
                                            Else
                                                aryFieldLst.Item("Returnlicenses".ToUpper) = _
                                                    aryFieldLst.Item("Returnlicenses".ToUpper) & SplitReturnlicenses & r2("LicensesKey")

                                            End If

                                        End While
                                    End Using
                                End Using

                                Try

                                    Interlocked.Increment(FThreadsTotal)
                                    obj.UseCompId = Convert.ToInt32(r("CompCode"))
                                    SyncLock FUseCompIndex.Item(r("CompCode")).WaitProcessNumber.GetType
                                        FUseCompIndex.Item(r("CompCode")).WaitProcessNumber += 1
                                    End SyncLock

                                    UpdateSOStatusUI(FMainForm, FMainForm.treLstSO, obj, SOStatus.NotUpdateImg)

                                    ThreadPool.QueueUserWorkItem(AddressOf UpdateProcessCmd, aryFieldLst)
                                    evn.Set()
                                Catch ex As Exception
                                    Interlocked.Decrement(FThreadsTotal)
                                End Try
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
            'Interlocked.Decrement(FThreadsTotal)

            obj = Nothing
        End Try



    End Sub
    '更新成功狀態
    Private Sub UpdDataOK(ByVal obj As Dictionary(Of String, String))
        Try
            SyncLock obj
                Using cn As New OracleClient.OracleConnection(FSOIndex.Item("-99").OraConnectString)
                    cn.Open()
                    Using cmd As OracleClient.OracleCommand = cn.CreateCommand

                        obj.Item("ERRCODE") = String.Empty
                        obj.Item("ERRMSG") = String.Empty
                        obj.Item("AuthTime".ToUpper) = Now
                        Dim pCmdStatus As New OracleParameter("CMDSTATUS", OracleType.VarChar)
                        Dim pErrCode As New OracleParameter("ERRCODE", OracleType.VarChar)
                        Dim pErrMsg As New OracleParameter("ERRMSG", OracleType.VarChar)
                        Dim pAuthTime As New OracleParameter("AUTHTIME", OracleType.DateTime)
                        Dim pSeq As New OracleParameter("SEQNO", OracleType.Number)
                        pCmdStatus.Value = "S"
                        
                        pAuthTime.Value = Date.Parse(obj.Item("AuthTime".ToUpper))
                        
                        pSeq.Value = Convert.ToInt32(obj.Item("SEQNO"))
                        pErrCode.Value = DBNull.Value
                        pErrMsg.Value = DBNull.Value

                        Dim aSQL As String = "UPDATE SO560 SET CMDSTATUS = :CMDSTATUS," & _
                                            " AUTHTIME = :AUTHTIME,ERRCODE = :ERRCODE, " & _
                                            " ERRMSG = : ERRMSG " & _
                                            " WHERE SEQNO = :SEQNO "

                        cmd.CommandText = aSQL
                        cmd.Parameters.Add(pCmdStatus)
                        cmd.Parameters.Add(pAuthTime)
                        cmd.Parameters.Add(pSeq)
                        cmd.Parameters.Add(pErrCode)
                        cmd.Parameters.Add(pErrMsg)
                        cmd.ExecuteNonQuery()
                    End Using
                End Using
            End SyncLock
            obj.Item("CMDSTATUS") = "S"
            
        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, " 命令序號 [ " & obj.Item("SEQNO") & " ]", Nothing)
            UpdateSysErrUI(FMainForm, FMainForm.TreLstSysErr, ex, " 命令序號 [ " & obj.Item("SEQNO") & " ]")
        Finally
            UpdateHightCmdUI(FMainForm, obj, UpdCmdMode.Upd)
        End Try
    End Sub


    '更新失敗狀態
    Private Function UpdFailData(ByRef obj As Dictionary(Of String, String)) As Boolean
        Try
            SyncLock obj
                obj.Item("CMDSTATUS") = "E"
                obj.Item("AuthTime".ToUpper) = Now
                If String.IsNullOrEmpty(obj.Item("ERRCODE")) Then
                    obj.Item("ERRORCODE") = "-7999"
                    obj.Item("ERRMSG") = "CableSoft_UnKnow_Err"
                End If
                Using cn As New OracleConnection(FSOIndex.Item("-99").OraConnectString)
                    cn.Open()
                    Using cmd As OracleClient.OracleCommand = cn.CreateCommand
                        Dim pCmdStatus As New OracleParameter("CMDSTATUS", OracleType.VarChar)
                        Dim pErrCode As New OracleParameter("ERRCODE", OracleType.VarChar)
                        Dim pErrMsg As New OracleParameter("ERRMSG", OracleType.VarChar)
                        Dim pSeq As New OracleParameter("SEQNO", OracleType.Number)
                        Dim pAuthTime As New OracleParameter("AUTHTIME", OracleType.DateTime)
                        Dim aSQL As String = "UPDATE SO560 SET CMDSTATUS = :CMDSTATUS," & _
                                                " AUTHTIME = :AUTHTIME,ERRCODE = :ERRCODE, " & _
                                                " ERRMSG = :ERRMSG " & _
                                                " WHERE SEQNO = :SEQNO "

                        cmd.CommandText = aSQL
                        pCmdStatus.Value = obj.Item("CMDSTATUS")
                        pAuthTime.Value = DateTime.Parse(obj.Item("AuthTime".ToUpper))
                        pErrCode.Value = obj.Item("ERRCODE")
                        pErrMsg.Value = obj.Item("ERRMSG")
                        pSeq.Value = Convert.ToInt32(obj.Item("SEQNO"))
                        cmd.Parameters.Add(pCmdStatus)
                        cmd.Parameters.Add(pAuthTime)
                        cmd.Parameters.Add(pSeq)
                        cmd.Parameters.Add(pErrCode)
                        cmd.Parameters.Add(pErrMsg)
                        cmd.ExecuteNonQuery()
                    End Using
                End Using

                Return True
            End SyncLock
        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, " 命令序號 [ " & obj.Item("SEQNO") & " ]", Nothing)
            UpdateSysErrUI(FMainForm, FMainForm.TreLstSysErr, ex, " 命令序號 [ " & obj.Item("SEQNO") & " ]")
            'Finally
            '    UpdateHightCmdUI(FMainForm, obj, UpdCmdMode.Upd)
        End Try
    End Function
    '更新異動狀態
    Private Function UpdProcing(ByRef obj As Dictionary(Of String, String)) As Boolean
        Try
            SyncLock obj
                Using cn As New OracleClient.OracleConnection(FSOIndex.Item("-99").OraConnectString)
                    cn.Open()
                    Using cmd As OracleClient.OracleCommand = cn.CreateCommand
                        cmd.CommandText = String.Format("UPDATE SO560 SET CMDSTATUS='P2' " & _
                                                        " WHERE SEQNO ={0} ", obj.Item("SEQNO"))

                        cmd.ExecuteNonQuery()

                    End Using
                End Using
                obj.Item("CMDSTATUS") = "P2"
                Return True
            End SyncLock
        Catch ex As Exception
            obj.Item("CMDSTATUS") = "E"
            obj.Item("ERRCODE") = "-8888"
            obj.Item("ERRMSG") = "CableSoft_UPDATE_P2_ERR"
            WriteErrTxtLog.WriteTxtError(ex, "[" & FSOIndex.Item("-99").CompName & "] 命令序號 [ " & obj.Item("SEQNO") & " ]", Nothing)
            UpdateSysErrUI(FMainForm, FMainForm.TreLstSysErr, ex, " 命令序號 [ " & obj.Item("SEQNO") & " ]")
            Return False
        Finally
            UpdateSOStatusUI(FMainForm, FMainForm.treLstSO, FSOIndex.Item("-99"), SOStatus.NotUpdateImg)
            UpdateHightCmdUI(FMainForm, obj, UpdCmdMode.Add)
        End Try
    End Function

    Private Function ReadWebSite(ByRef obj As Dictionary(Of String, String)) As Boolean
        SyncLock obj
            Dim aRequest As HttpWebRequest = Nothing
            Dim aResponse As HttpWebResponse = Nothing
            Dim aRet As Boolean = True
            Dim aUrl As String = FPushUrl.TrimEnd("?")
            'aUrl = aUrl & "?target=" & obj.Item("STBSNO") & _
            '        "&duration=" & FPushDuration & _
            '        "&repeat=" & FPushRepeat
            Dim aryLicenses As Array = obj.Item("RETURNLICENSES").Split(SplitReturnlicenses)
            Dim aNoteCnt As Int32 = obj.Item("NOTES").Split(SplitNoteChar).Length
            If aryLicenses.Length <> aNoteCnt Then
                obj.Item("CMDSTATUS") = "E"
                obj.Item("ERRCODE") = "-7999"
                obj.Item("ERRMSG") = "CableSoft_ReturnLength_Err"
                Return False
            End If

            For i As Int32 = 0 To aryLicenses.Length - 1
                Try
                    Thread.Sleep(100)

                    'aUrl = aUrl & "&licenses=" & _
                    '    Convert.ToBase64String(Encoding.UTF8.GetBytes(aryLicenses.GetValue(i)))

                    Dim aObjPara As New objPara()
                    aObjPara.target = Int32.Parse(obj.Item("STBSNO"))
                    aObjPara.duration = Int32.Parse(FPushDuration)
                    aObjPara.repeat = Int32.Parse(FPushRepeat)
                    aObjPara.licenses = New String() {aryLicenses.GetValue(i)}

                    Dim bty() As Byte = Encoding.ASCII.GetBytes(JsonConvert.SerializeObject(aObjPara))
                    aRequest = HttpWebRequest.Create(aUrl)
                    aRequest.Method = FPushUrlMethod
                    aRequest.ContentType = "text/plain"

                    aRequest.ContentLength = bty.Length
                    aRequest.Timeout = Convert.ToInt32(FPushUrlTimeout) * 1000
                    aRequest.Credentials = CredentialCache.DefaultCredentials
                    obj.Item("PushServerUrl".ToUpper) = Encoding.ASCII.GetString(bty)
                    Using stm As Stream = aRequest.GetRequestStream
                        stm.Write(bty, 0, bty.Length)
                    End Using
                    aResponse = aRequest.GetResponse
                    'aRequest = HttpWebRequest.Create(aUrl)
                    ''aRequest.KeepAlive = True
                    'aRequest.Method = FUrlMetho
                    ''aRequest.ContentType = "application/x-www-form-urlencoded"
                    'aRequest.ContentType = "text/plain"

                    'aRequest.ContentLength = bty.Length
                    'aRequest.Timeout = Convert.ToInt32(FUrlTimeOut) * 1000


                    'aRequest.Credentials = CredentialCache.DefaultCredentials



                    obj.Item("PushServerCode".ToUpper) = aResponse.StatusCode
                    obj.Item("PushServerMsg".ToUpper) = aResponse.StatusDescription.ToString
                    obj.Item("PushServerHeadMsg".ToUpper) = aResponse.Headers.ToString
                    If aResponse.StatusCode <> HttpStatusCode.OK Then
                        obj.Item("CMDSTATUS") = "E"
                        obj.Item("ERRCODE") = obj.Item("PushServerCode".ToUpper)
                        obj.Item("ERRMSG") = obj.Item("PushServerHeadMsg".ToUpper)
                        aRet = False

                    End If
                    If FErrCodeLst.ContainsKey(obj.Item("PushServerCode".ToUpper)) Then
                        obj.Item("PushServerMsg".ToUpper) = FErrCodeLst.Item("PushServerCode".ToUpper)
                    End If

                    obj.Item("ERRCODE") = String.Empty
                    obj.Item("ERRMSG") = String.Empty


                Catch ex As WebException
                    aRet = False
                    obj.Item("CMDSTATUS") = "E"
                    obj.Item("PushServerCode".ToUpper) = ex.Status
                    obj.Item("PushServerMsg".ToUpper) = ex.Status.ToString
                    obj.Item("PushServerHeadMsg".ToUpper) = ex.Message
                    If FErrCodeLst.ContainsKey(obj.Item("PushServerCode".ToUpper)) Then
                        obj.Item("PushServerMsg".ToUpper) = FErrCodeLst.Item("PushServerCode".ToUpper)
                    End If
                    obj.Item("ERRCODE") = obj.Item("PushServerCode".ToUpper)
                    obj.Item("ERRMSG") = obj.Item("PushServerMsg".ToUpper)
                    WriteErrTxtLog.WriteTxtError(ex, obj.Item("COMPCODE"), obj.Item("SEQNO"), Nothing, Nothing, Nothing)
                    Exit For
                Catch ex As Exception
                    aRet = False

                    obj.Item("CMDSTATUS") = "E"
                    obj.Item("PushServerCode".ToUpper) = "-7775"
                    obj.Item("PushServerMsg".ToUpper) = "CableSoft_ReadWebSite_ERR"
                    obj.Item("PushServerHeadMsg".ToUpper) = ex.Message
                    obj.Item("ERRCODE") = obj.Item("PushServerCode".ToUpper)
                    obj.Item("ERRMSG") = obj.Item("PushServerMsg".ToUpper)
                    WriteErrTxtLog.WriteTxtError(ex, obj.Item("COMPCODE"), obj.Item("SEQNO"), Nothing, Nothing, Nothing)
                    Exit For
                Finally
                    If aResponse IsNot Nothing Then
                        aResponse.Close()
                    End If
                    If aRequest IsNot Nothing Then
                        aRequest.Abort()
                    End If

                    UpdateLowCmdUI(FMainForm, obj)
                End Try
            Next
            Return aRet
        End SyncLock

    End Function
    Private Sub UpdateProcessCmd(ByVal obj As Dictionary(Of String, String))

        SyncLock obj
            
            Dim aSO As SOInfoClass = FSOIndex("-99")

            Try

                SyncLock FUseCompIndex(obj.Item("COMPCODE")).ProcessingNumber.GetType
                    FUseCompIndex(obj.Item("COMPCODE")).ProcessingNumber += 1
                    
                End SyncLock
                aSO.UseCompId = obj.Item("COMPCODE")
                UpdateSOStatusUI(FMainForm, FMainForm.treLstSO, aSO, SOStatus.NotUpdateImg)
                'Thread.Sleep(5000)
                If FThreadStop Then
                    Exit Sub
                End If

                If UpdProcing(obj) Then
                    If ReadWebSite(obj) Then
                        UpdDataOK(obj)
                    Else
                        UpdFailData(obj)
                    End If
                Else
                    UpdFailData(obj)
                End If



            Catch ex As Exception
                If Not IsConnectOK(aSO.OracleCn) Then
                    If aSO.ConnectionOK Then
                        UpdateSOStatusUI(FMainForm, FMainForm.treLstSO, aSO, SOStatus.NO)
                    End If
                Else
                    UpdateSOStatusUI(FMainForm, FMainForm.treLstSO, aSO, SOStatus.Yes)
                End If

                WriteErrTxtLog.WriteTxtError(ex, obj.Item("COMPCODE"), obj.Item("SEQNO"), Nothing, Nothing, Nothing)
            Finally
                SyncLock aSO
                    aSO.UseCompId = obj.Item("COMPCODE")
                    SyncLock FUseCompIndex(obj.Item("COMPCODE")).ProcessingNumber.GetType
                        FUseCompIndex(obj.Item("COMPCODE")).ProcessingNumber -= 1
                    End SyncLock
                    SyncLock FUseCompIndex(aSO.UseCompId).WaitProcessNumber.GetType
                        FUseCompIndex(obj.Item("COMPCODE")).WaitProcessNumber -= 1
                    End SyncLock


                    UpdateSOStatusUI(FMainForm, FMainForm.treLstSO, aSO, SOStatus.NotUpdateImg)
                    UpdateHightCmdUI(FMainForm, obj, UpdCmdMode.Upd)
                    Interlocked.Decrement(FThreadsTotal)
                    'UpdateLowCmdUI(FMainForm, obj)
                    obj.Clear()
                End SyncLock
            End Try

        End SyncLock
    End Sub
    Private Function TransformJson(ByVal aTarget As String, _
                                   ByVal aDuration As String, _
                                   ByVal aRepeat As String, _
                                   ByVal aLicenses As String()) As String
        Try

        Catch ex As Exception
            'WriteErrTxtLog.WriteTxtError(ex, "[" & FSOIndex.Item("-99").CompName & "] 命令序號 [ " & obj.Item("SEQNO") & " ]", Nothing)
            'UpdateSysErrUI(FMainForm, FMainForm.TreLstSysErr, ex, " 命令序號 [ " & obj.Item("SEQNO") & " ]")
        End Try

    End Function
    Public Class objPara
        Private _target As Int32
        Private _duration As Int32
        Private _repeat As Int32
        Private _licenses As String()
        Public Property target() As Int32
            Get
                Return _target
            End Get
            Set(ByVal value As Int32)
                _target = value
            End Set
        End Property
        Public Property duration() As Int32
            Get
                Return _duration
            End Get
            Set(ByVal value As Int32)
                _duration = value
            End Set
        End Property
        Public Property repeat() As Int32
            Get
                Return _repeat
            End Get
            Set(ByVal value As Int32)
                _repeat = value
            End Set
        End Property
        Public Property licenses() As String()
            Get
                Return _licenses
            End Get
            Set(ByVal value As String())
                _licenses = value
            End Set
        End Property

        'aUrl = aUrl & "?target=" & obj.Item("STBSNO") & _
        '            "&duration=" & FPushDuration & _
        '            "&repeat=" & FPushRepeat
    End Class


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
