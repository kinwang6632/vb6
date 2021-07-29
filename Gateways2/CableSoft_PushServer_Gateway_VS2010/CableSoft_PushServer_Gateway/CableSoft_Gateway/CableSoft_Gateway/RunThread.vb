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
Imports System.Net
Imports Newtonsoft.Json
Imports Raize.CodeSiteLogging
Imports System.Globalization
Imports System.Text.RegularExpressions
Module RunThread
    Private ti As RegisteredWaitHandle = Nothing
    Public evn As New AutoResetEvent(False)

    Public FThreadsTotal As Int32 = 0


    Public fWebSiteErrTotal As Int32 = 0                 'WEB 錯誤幾次了


    Public fSOIndex As New System.Collections.Generic.Dictionary(Of String, SOInfoClass)
    Public fUseCompIndex As New Dictionary(Of String, UseCompanyClass)
    Private fExeCmdSQL As String
    Private Const aSelectSQL As String = "SELECT * FROM SO560 WHERE ( CMDSTATUS = 'S1' )  " & _
                        " AND ROWNUM<={0}  AND COMPCODE IN ( {1} ) " & _
                        " AND Nvl(BATCH,0)=0 ORDER BY SEQNO "
    Private aSQLCmd As String


    Private Const SplitReturnlicenses As Char = ","c
    Private Const SplitNoteChar As Char = ","c
    Private fUpdTotal As Int32 = 0
    Private Const fSO561AField As String = " SEQNO,CMD,REQUESTTIME,RESPONSETIME,EXCHTIME," & _
                        "ENCRYPTED_ASSET_FLAG,VALIDITY_FLAG,NB_OF_R_VOD_ENT,CHIPSET_TYPE_FLAG, " & _
                        "ADDITIONAL_ADDRESS_FLAG,DMM_CREATION_FATE_FLAG,DMM_CREATION_DATE," & _
                        "COMPCODE,NOTES,UPDEN,UPDTIME,ERRCODE,ERRMSG,RETURNLICENSES,STBSNO," & _
                        "ICCNO,RESVTIME,BATCH,MODETYPE,SNO,AUTHTIME,AUTHREQUESTTIME,CMDSTATUS"

    'Private Const fInsSendNagra As String = "Insert Into {0}Send_Nagra (HIGH_LEVEL_CMD_ID," &
    '                    "ICC_NO,STB_NO," &
    '                    "NOTES,CMD_STATUS,OPERATOR,UPDTIME,MIS_IRD_CMD_DATA," &
    '                    "SEQNO,COMPCODE,RESENTTIMES,STB_FLAG," &
    '                    "RETURNWORKORDER,MODETYPE) Values ({1},{2},{3},{4},{5},{6},{7},{8}," &
    '                    "{9},{10},{11},{12},{13},{14})"
    Private Const fInsSendNagra As String = "Insert Into {0}Send_Nagra (HIGH_LEVEL_CMD_ID," &
                    "ICC_NO,STB_NO," &
                    "NOTES,CMD_STATUS,OPERATOR,UPDTIME,MIS_IRD_CMD_DATA," &
                    "SEQNO,COMPCODE,RESENTTIMES,STB_FLAG," &
                    "RETURNWORKORDER,MODETYPE) Values (:HIGH_LEVEL_CMD_ID,:ICC_NO," &
                    ":STB_NO,:NOTES,:CMD_STATUS,:OPERATOR,:UPDTIME,:MIS_IRD_CMD_DATA," &
                    ":SEQNO,:COMPCODE,:RESENTTIMES,:STB_FLAG,:RETURNWORKORDER,:MODETYPE)"
    Private FWatch As New Stopwatch
    Private Const fGCTime As Int32 = 120

    Public Sub RunGateway()
        Try
            'FMainForm = aParent
            If ThreadInitial() Then
                Try

                    FWatch.Start()

                    ti = ThreadPool.RegisterWaitForSingleObject(evn, _
                                                                New WaitOrTimerCallback(AddressOf ThreadProcess), _
                                                                fSOInfoList, TimeSpan.FromSeconds(fReadDataTime), False)

                    evn.Set()

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

    Private Function ThreadInitial() As Boolean
        Try
            'If FMaxThread > (Environment.ProcessorCount * 25) Then
            '    FMaxThread = (Environment.ProcessorCount * 25)
            'End If
            If fReadDataTime <= 0 Then
                fReadDataTime = 5
            End If

            ThreadPool.SetMaxThreads(fMaxThread, fMaxThread)
            ThreadPool.SetMinThreads(2, 2)
            If fReadDataTime <= 0 Then
                fReadDataTime = 30
            End If
            aSQLCmd = String.Format(aSelectSQL, fProcessNumber, fUseCompWhere)
            fSOIndex.Clear()
            For i As Int32 = 0 To fSOInfoList.Count - 1
                fSOIndex.Add(fSOInfoList(i).CompID, fSOInfoList.Item(i))
            Next
            fUseCompIndex.Clear()
            For i As Int32 = 0 To fUseCompList.Count - 1
                fUseCompIndex.Add(fUseCompList(i).CompID, fUseCompList(i))
            Next
            Return True
        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
        End Try


    End Function
    Private Sub ThreadProcess(ByVal objSOInfoLst As List(Of SOInfoClass), ByVal timedOut As Boolean)

        Try
            'CodeSite.SendMemoryStatus()
            'fRegOK = GetSystemInfo.IsRegister(ShowForm.No)
            fRegOK = True
            'Date.Parse(aRegInfo(6), CultureInfo.CreateSpecificCulture("zh-TW")).Date()
            If Date.Now.Date > Date.Parse(fRegDate, CultureInfo.CreateSpecificCulture("zh-TW")).Date Then
                fRegOK = False
            End If

            If (Not fThreadStop) AndAlso (fRegOK) Then

                For i As Int32 = 0 To objSOInfoLst.Count - 1
                    ThreadPool.QueueUserWorkItem(AddressOf ThreadCmd, objSOInfoLst.Item(i))
                Next
            Else
                If FThreadsTotal <= 0 Then
                    Dim blnAllLoad As Boolean = True
                    If fRegOK Then
                        For i As Int32 = 0 To fSOInfoList.Count - 1
                            If fSOInfoList(i).FirstLoad Then
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
                        For i As Int32 = 0 To fSOInfoList.Count - 1
                            fSOInfoList(i).ConnectionOK = False
                            fSOInfoList(i).ProcessingNumber = 0
                            fSOInfoList(i).WaitProcessNumber = 0
                            fSOInfoList(i).FirstLoad = True
                            If fSOInfoList(i).OracleCn IsNot Nothing _
                                AndAlso fSOInfoList(i).OracleCn.State = ConnectionState.Open Then
                                fSOInfoList(i).OracleCn.Close()
                            End If
                        Next
                        DisableControl(RunStatus.StopCmd)
                        If fRegOK Then
                            UpdateSOStatusUI(Nothing, SOStatus.Initial)
                        Else
                            UpdateSOStatusUI(Nothing, SOStatus.Key)
                        End If

                    End If
                End If
            End If
        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
            XtraMessageBox.Show("停止失敗！" & Environment.NewLine & ex.Message, "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally

            'objSOInfoLst = Nothing
        End Try

    End Sub
    Private Sub ThreadCmd(ByVal objSOInfoClass As SOInfoClass)

        Try

            If (fThreadStop) OrElse (FThreadsTotal > 0) Then
                Exit Sub
            Else
                If FWatch.Elapsed.TotalMinutes >= fGCTime Then
                    Interlocked.Increment(FThreadsTotal)
                    Try
                        If FThreadsTotal = 1 Then
                            FWatch.Reset()
                            GC.Collect()
                            GC.WaitForFullGCComplete()
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

            SyncLock objSOInfoClass


                'Dim aSytle As SOStatus = SOStatus.NO
                'Dim aId As String = Date.Now.ToString("yyyyMMddHHmmssff")
                Dim aSytle As New ThreadLocal(Of SOStatus)
                Dim aId As New ThreadLocal(Of String)
                aId.Value = Date.Now.ToString("yyyyMMddHHmmssff")
                aSytle.Value = SOStatus.NO
                Try
                    UpdateSODatabase(aId.Value, objSOInfoClass, SOStatus.NotUpdateImg)
                    '判斷是否連線
                    If Not IsConnectOK(objSOInfoClass.OracleCn) Then
                        aSytle.Value = SOStatus.NO
                        If (Not objSOInfoClass.FirstLoad) AndAlso (objSOInfoClass.ConnectionOK) Then
                            WriteErrTxtLog.WriteSOStatus(objSOInfoClass.CompName & " ---> 連線後已斷線")
                        End If
                        objSOInfoClass.ConnectionOK = False

                    Else
                        aSytle.Value = SOStatus.Yes
                        If (Not objSOInfoClass.FirstLoad) AndAlso (Not objSOInfoClass.ConnectionOK) Then
                            WriteErrTxtLog.WriteSOStatus(objSOInfoClass.CompName & " ---> 斷線後已連線")
                        End If
                        objSOInfoClass.ConnectionOK = True
                    End If
                    If objSOInfoClass.FirstLoad Then
                        WriteErrTxtLog.WriteSOStatus(objSOInfoClass.CompName & _
                                                     IIf(objSOInfoClass.ConnectionOK, " ---> 首次連線成功", " ---> 首次連線失敗"))
                    End If
                    objSOInfoClass.FirstLoad = False
                    UpdateSODatabase(aId.Value, objSOInfoClass, aSytle.Value)
                    UpdateSOStatusUI(objSOInfoClass, aSytle.Value)
                    If (objSOInfoClass.WaitProcessNumber <= 0) AndAlso _
                        (objSOInfoClass.ConnectionOK) Then
                        ProcessCMD(objSOInfoClass)
                    End If
                Finally
                    aId.Value = Nothing
                    aSytle.Dispose()
                    aId.Dispose()
                    aSytle = Nothing
                    aId = Nothing
                End Try
            End SyncLock
        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, objSOInfoClass.CompID, Nothing, Nothing, Nothing, Nothing)
        Finally
            'Thread.Sleep(500)
            'If objSOInfoClass IsNot Nothing Then
            '    If objSOInfoClass.OracleCn IsNot Nothing Then
            '        objSOInfoClass.OracleCn.Close()
            '        objSOInfoClass.OracleCn.Dispose()
            '    End If

            '    objSOInfoClass = Nothing
            'End If

        End Try
    End Sub
    '開始處理資料
    Private Sub ProcessCMD(ByRef objSOInfoClass As SOInfoClass)
        If fThreadStop OrElse FThreadsTotal > 0 Then
            Exit Sub
        Else
            Interlocked.Exchange(fWebSiteErrTotal, 0)       '讀取WEB錯誤重新計數
        End If

        Try

            If (objSOInfoClass.OracleCn.State = ConnectionState.Closed) Then
                Try
                    objSOInfoClass.OracleCn.Open()
                Catch ex As Exception
                    UpdateSOStatusUI(objSOInfoClass, SOStatus.NO)
                    objSOInfoClass.ConnectionOK = False
                    Exit Sub
                End Try


                Using cmd As OracleCommand = objSOInfoClass.OracleCn.CreateCommand
                    cmd.CommandText = aSQLCmd

                    Using r As OracleDataReader = cmd.ExecuteReader(CommandBehavior.CloseConnection)
                        While r.Read
                            '讀出所有SO560資料,並將欄位資訊記錄至aryFieldLst
                            Dim aryFieldLst As New Dictionary(Of String, String)
                            aryFieldLst.Clear()
                            For i As Int32 = 0 To r.FieldCount - 1
                                If r.IsDBNull(i) Then
                                    aryFieldLst.Add(r.GetName(i).ToUpper, String.Empty)
                                Else
                                    aryFieldLst.Add(r.GetName(i).ToUpper, r.Item(i).ToString)
                                End If
                            Next
                            aryFieldLst.Add("RightCompId".ToUpper, objSOInfoClass.CompID)
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

                            '讀出SO561B將Liceness資訊儲存至aryFieldLst
                            Using cmd2 As New OracleCommand(aSQLLicenses, objSOInfoClass.OracleCn)
                                Using r2 As OracleDataReader = cmd2.ExecuteReader()
                                    While r2.Read
                                        If Not IsDBNull(r2("LicensesKey")) Then
                                            If String.IsNullOrEmpty(aryFieldLst.Item("Returnlicenses".ToUpper)) Then
                                                aryFieldLst.Item("Returnlicenses".ToUpper) = r2("LicensesKey")
                                            Else
                                                aryFieldLst.Item("Returnlicenses".ToUpper) = _
                                                    aryFieldLst.Item("Returnlicenses".ToUpper) & SplitReturnlicenses & r2("LicensesKey")
                                            End If
                                        End If
                                    End While
                                End Using
                            End Using
                            aryFieldLst.Item("CmdStatus".ToUpper) = "P2"
                            Try
                                objSOInfoClass.UseCompId = Convert.ToInt32(r("CompCode"))
                                SyncLock objSOInfoClass.WaitProcessNumber.GetType
                                    objSOInfoClass.WaitProcessNumber += 1
                                    fUseCompIndex.Item(r("CompCode")).WaitProcessNumber += 1
                                    objSOInfoClass.ProcessingNumber += 1
                                    fUseCompIndex.Item(r("CompCode")).ProcessingNumber += 1
                                End SyncLock

                                'SyncLock aSO.Value.ProcessingNumber.GetType
                                '    aSO.Value.ProcessingNumber += 1
                                '    FUseCompIndex(objFieldLst.Item("COMPCODE")).ProcessingNumber += 1
                                'End SyncLock

                                'UpdateSOStatusUI(FMainForm, FMainForm.treLstSO, objSOInfoClass, SOStatus.NotUpdateImg)
                                ThreadPool.QueueUserWorkItem(AddressOf UpdateProcessCmd, aryFieldLst)
                                'evn.Set()
                            Catch ex As Exception

                                SyncLock objSOInfoClass.WaitProcessNumber.GetType
                                    If objSOInfoClass.WaitProcessNumber > 0 Then
                                        objSOInfoClass.WaitProcessNumber -= 1
                                    End If
                                    If objSOInfoClass.ProcessingNumber > 0 Then
                                        objSOInfoClass.ProcessingNumber -= 1
                                    End If
                                End SyncLock
                                WriteErrTxtLog.WriteTxtError(ex, objSOInfoClass.CompID, Nothing, Nothing, "SQL : " & aSQLCmd, Nothing)
                            End Try
                        End While

                    End Using
                End Using
            End If

        Catch ex As Exception
            If Not IsConnectOK(objSOInfoClass.OracleCn) Then
                If objSOInfoClass.ConnectionOK Then
                    UpdateSOStatusUI(objSOInfoClass, SOStatus.NO)
                    'obj.ConnectionOK = False
                End If
            End If
            UpdateSysErrUI(ex, "公司別 : " & objSOInfoClass.CompID.ToString)
            WriteErrTxtLog.WriteTxtError(ex, objSOInfoClass.CompID, Nothing, Nothing, "SQL : " & aSQLCmd, Nothing)
        Finally

            'If objSOInfoClass.OracleCn.State = ConnectionState.Open Then
            '    objSOInfoClass.OracleCn.Close()
            '    objSOInfoClass.OracleCn.Dispose()
            'End If
            'objSOInfoClass = Nothing
            'System.GC.Collect()
            'obj = Nothing
        End Try



    End Sub
    '更新成功狀態
    Private Sub UpdDataOK(ByVal obj As Dictionary(Of String, String))
        Try
            SyncLock obj
                Using cn As New OracleClient.OracleConnection(fSOIndex.Item("-99").OraConnectString)
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
                        Dim pAuthRequestTime As New OracleParameter("AuthRequestTime".ToUpper, OracleType.DateTime)
                        pCmdStatus.Value = "S"

                        pAuthTime.Value = Date.Parse(obj.Item("AuthTime".ToUpper), _
                                                     CultureInfo.CreateSpecificCulture("zh-TW"))

                        pSeq.Value = Convert.ToInt32(obj.Item("SEQNO"))
                        pErrCode.Value = DBNull.Value
                        pErrMsg.Value = DBNull.Value

                        Dim aSQL As String = "UPDATE SO560 SET CMDSTATUS = :CMDSTATUS," & _
                                            " AUTHTIME = :AUTHTIME,ERRCODE = :ERRCODE, " & _
                                            " ERRMSG = : ERRMSG, AUTHREQUESTTIME = :AUTHREQUESTTIME " & _
                                            " WHERE SEQNO = :SEQNO "

                        cmd.CommandText = aSQL
                        cmd.Parameters.Add(pCmdStatus)
                        cmd.Parameters.Add(pAuthTime)
                        cmd.Parameters.Add(pSeq)
                        cmd.Parameters.Add(pErrCode)
                        cmd.Parameters.Add(pErrMsg)
                        cmd.Parameters.Add(pAuthTime.Value)
                        cmd.ExecuteNonQuery()
                    End Using
                End Using
            End SyncLock
            obj.Item("CMDSTATUS") = "S"

        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, " 命令序號 [ " & obj.Item("SEQNO") & " ]", Nothing)
            UpdateSysErrUI(ex, " 命令序號 [ " & obj.Item("SEQNO") & " ]")
        Finally
            UpdateHightCmdUI(obj, UpdCmdMode.Upd)

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
                Using cn As New OracleConnection(fSOIndex.Item("-99").OraConnectString)
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
            UpdateSysErrUI(ex, " 命令序號 [ " & obj.Item("SEQNO") & " ]")
            'Finally
            '    UpdateHightCmdUI(FMainForm, obj, UpdCmdMode.Upd)
        End Try
    End Function
    '更新異動狀態
    Private Function UpdProcing(ByRef obj As Dictionary(Of String, String)) As Boolean
        Try
            SyncLock obj
                Using cn As New OracleClient.OracleConnection(fSOIndex.Item("-99").OraConnectString)
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
            WriteErrTxtLog.WriteTxtError(ex, "[" & fSOIndex.Item("-99").CompName & "] 命令序號 [ " & obj.Item("SEQNO") & " ]", Nothing)
            UpdateSysErrUI(ex, " 命令序號 [ " & obj.Item("SEQNO") & " ]")
            Return False
        Finally
            UpdateSOStatusUI(fSOIndex.Item("-99"), SOStatus.NotUpdateImg)
            UpdateHightCmdUI(obj, UpdCmdMode.Add)
        End Try
    End Function

    Private Sub ReadWebSite(ByRef objFields As Dictionary(Of String, String))

        Dim aRequest As New ThreadLocal(Of HttpWebRequest)
        aRequest.Value = Nothing
        Dim aResponse As New ThreadLocal(Of HttpWebResponse)
        aResponse.Value = Nothing
        Dim aUrl As New ThreadLocal(Of String)
        Dim aRet As New ThreadLocal(Of Boolean)
        Dim aryLicenses As New ThreadLocal(Of Array)
        Dim aNoteCnt As New ThreadLocal(Of Int32)
        Dim aChangUrlNum As New ThreadLocal(Of Byte)        '改變Url的計數
        Dim aSendUrl As New ThreadLocal(Of Boolean)
        Dim aJson As New ThreadLocal(Of JsonClass)
        Dim bytData As New ThreadLocal(Of Byte())
        aNoteCnt.Value = objFields.Item("NOTES").Split(SplitNoteChar).Length
        aryLicenses.Value = objFields.Item("RETURNLICENSES").Split(SplitReturnlicenses)
        aRet.Value = True

        If aryLicenses.Value.Length <> aNoteCnt.Value Then
            objFields.Item("CMDSTATUS") = "E"
            objFields.Item("ERRCODE") = CustomerUrlError.ReturnLicenses_Length_Error
            objFields.Item("ERRMSG") = [Enum].GetName(GetType(CustomerUrlError), CustomerUrlError.ReturnLicenses_Length_Error)
            objFields.Item("PushServerMsg") = [Enum].GetName(GetType(CustomerUrlError), CustomerUrlError.ReturnLicenses_Length_Error)
            UpdUiAndData(objFields)
            Exit Sub
        End If


        aChangUrlNum.Value = 0

        aSendUrl.Value = True
        objFields.Item("AuthRequestTime".ToUpper) = Now
        Interlocked.Increment(FThreadsTotal)
        Try
            '有幾個Licenses就要送幾次,其中一個失敗就要算失敗
            For i = 0 To aryLicenses.Value.Length - 1
                If Not aSendUrl.Value Then
                    Exit For
                End If

                'Dim aJson As New ThreadLocal(Of JsonClass)
                'Dim bytData As New ThreadLocal(Of Byte())

                Try
                    aJson.Value = New JsonClass()
                    aJson.Value.target = Int32.Parse(objFields.Item("STBSNO".ToUpper))
                    aJson.Value.duration = Int32.Parse(fPushDuration)
                    aJson.Value.repeat = Int32.Parse(fPushRepeat)
                    aJson.Value.licenses = New String() {aryLicenses.Value.GetValue(i)}
                    bytData.Value = Encoding.ASCII.GetBytes(JsonConvert.SerializeObject(aJson.Value))
                Catch ex As Exception
                    objFields.Item("CmdStatus".ToUpper) = "E"
                    objFields.Item("ErrCode".ToUpper) = CustomerUrlError.Data_Format_Error
                    objFields.Item("ErrMsg".ToUpper) = [Enum].GetName(GetType(CustomerUrlError), CustomerUrlError.Data_Format_Error)
                    objFields.Item("PushServerCode".ToUpper) = CustomerUrlError.Data_Format_Error
                    objFields.Item("PushServerMsg".ToUpper) = [Enum].GetName(GetType(CustomerUrlError), CustomerUrlError.Data_Format_Error)
                    objFields.Item("PushServerUrl".ToUpper) = "Json格式錯誤!!!"
                    UpdateLowCmdUI(objFields)
                    Exit For
                End Try

                '記錄送出Url的時間
                'objFields.Item("AuthRequestTime".ToUpper) = Now
                '連結失敗重新連結,直到成功或Retry次數滿為止
                For aIndex As Int32 = 0 To fPushServerRetryNum - 1
                    aSendUrl.Value = False
                    Dim stm As New ThreadLocal(Of Stream)
                    Try
                        'CodeSite.Send("第 " & aIndex + 1 & " 次連線," & objFields.Item("SEQNO"))
                        aUrl.Value = fPushServerUrls.Item(aChangUrlNum.Value)
                        objFields.Item("PushServerUrl".ToUpper) = Encoding.ASCII.GetString(bytData.Value)
                        aRequest.Value = HttpWebRequest.Create(aUrl.Value)

                        aRequest.Value.Method = fPushUrlMethod
                        aRequest.Value.ContentType = "text/plain"
                        aRequest.Value.ContentLength = bytData.Value.Length
                        aRequest.Value.Timeout = Convert.ToInt32(fPushUrlTimeout) * 1000
                        aRequest.Value.Credentials = CredentialCache.DefaultCredentials
                        aRequest.Value.ProtocolVersion = HttpVersion.Version11
                        aRequest.Value.ReadWriteTimeout = Convert.ToInt32(fPushUrlTimeout) * 1000

                        stm.Value = aRequest.Value.GetRequestStream

                        stm.Value.WriteTimeout = Convert.ToInt32(fPushUrlTimeout) * 1000
                        stm.Value.Write(bytData.Value, 0, bytData.Value.Length)
                        stm.Value.Flush()

                        aResponse.Value = aRequest.Value.GetResponse

                        If aResponse.Value.StatusCode = HttpStatusCode.OK Then

                            objFields.Item("CmdStatus".ToUpper) = "S"
                            objFields.Item("AuthTime".ToUpper) = Now
                            objFields.Item("ErrCode".ToUpper) = String.Empty
                            objFields.Item("ErrMsg".ToUpper) = String.Empty
                            objFields.Item("PushServerCode".ToUpper) = aResponse.Value.StatusCode
                            objFields.Item("PushServerMsg".ToUpper) = aResponse.Value.StatusDescription
                            aSendUrl.Value = True
                            aRequest.Value.Abort()
                            aRequest.Value = Nothing
                            aResponse.Value.Close()
                            aResponse.Value = Nothing
                            Exit For
                        Else
                            objFields.Item("CmdStatus".ToUpper) = "E"
                            objFields.Item("AuthTime".ToUpper) = Now
                            objFields.Item("ErrCode".ToUpper) = aResponse.Value.StatusCode
                            objFields.Item("ErrMsg".ToUpper) = aResponse.Value.StatusDescription
                            objFields.Item("PushServerCode".ToUpper) = aResponse.Value.StatusCode
                            objFields.Item("PushServerMsg".ToUpper) = aResponse.Value.StatusDescription
                        End If

                    Catch ex As WebException
                        WriteErrTxtLog.WriteTxtError(ex, objFields.Item("COMPCODE"), _
                                                     objFields.Item("SEQNO"), _
                                                     objFields.Item("Cmd".ToUpper), _
                                                     "Url = " & aUrl.Value & " , 第 " & aIndex + 1 & " 次連線 ", _
                                                     Nothing)
                        objFields.Item("CmdStatus".ToUpper) = "E"
                        If ex.Status = WebExceptionStatus.ProtocolError Then
                            objFields.Item("PushServerCode".ToUpper) = CType(ex.Response, HttpWebResponse).StatusCode
                            objFields.Item("PushServerMsg".ToUpper) = CType(ex.Response, HttpWebResponse).StatusDescription
                            objFields.Item("ErrCode".ToUpper) = CType(ex.Response, HttpWebResponse).StatusCode
                            objFields.Item("ErrMsg".ToUpper) = CType(ex.Response, HttpWebResponse).StatusDescription
                            objFields.Item("AuthTime".ToUpper) = Now
                        Else

                            objFields.Item("PushServerCode".ToUpper) = ex.Status
                            objFields.Item("PushServerMsg".ToUpper) = [Enum].GetName(GetType(WebExceptionStatus), ex.Status)
                            objFields.Item("ErrCode".ToUpper) = ex.Status
                            objFields.Item("ErrMsg".ToUpper) = [Enum].GetName(GetType(WebExceptionStatus), ex.Status)
                            objFields.Item("AuthTime".ToUpper) = Now
                        End If
                        UpdateSysErrUI(ex, _
                                        "命令序號 > " & objFields.Item("SEQNO".ToUpper) & _
                                        " 第 " & aIndex + 1 & " 次連線 , Url = " & aUrl.Value)

                    Catch ex As Exception
                        WriteErrTxtLog.WriteTxtError(ex, objFields.Item("COMPCODE"), _
                                                     objFields.Item("SEQNO"), _
                                                     objFields.Item("Cmd".ToUpper), _
                                                     "Url = " & aUrl.Value & " , 第 " & aIndex + 1 & " 次連線 ", _
                                                     Nothing)
                        objFields.Item("CmdStatus".ToUpper) = "E"
                        objFields.Item("PushServerCode".ToUpper) = CustomerUrlError.ConnectUrl_Error
                        objFields.Item("PushServerMsg".ToUpper) = [Enum].GetName(GetType(CustomerUrlError), CustomerUrlError.ConnectUrl_Error)
                        objFields.Item("ErrCode".ToUpper) = CustomerUrlError.ConnectUrl_Error
                        objFields.Item("ErrMsg".ToUpper) = [Enum].GetName(GetType(CustomerUrlError), CustomerUrlError.ConnectUrl_Error)
                        objFields.Item("AuthTime".ToUpper) = Now
                        UpdateSysErrUI(ex, _
                                       "命令序號 > " & objFields.Item("SEQNO".ToUpper) & _
                                       " 第 " & aIndex + 1 & " 次連線 , Url = " & aUrl.Value)
                    Finally
                        Try
                            stm.Value.Close()
                            stm.Value.Dispose()
                            stm.Dispose()
                            'aJson.Value.Dispose()
                            'aJson.Dispose()
                            aRequest.Value.Abort()
                            aRequest.Value = Nothing
                            aResponse.Value.Close()
                            aResponse.Value = Nothing

                        Catch ex As Exception

                        End Try

                    End Try

                    If aChangUrlNum.Value >= fPushServerUrls.Count - 1 Then
                        aChangUrlNum.Value = 0
                    Else
                        aChangUrlNum.Value += 1
                    End If
                    '找尋錯誤是否要重試
                    If Not String.IsNullOrEmpty(objFields.Item("ErrCode".ToUpper)) Then
                        If Not fErrRetryLst.Contains(objFields.Item("ErrCode".ToUpper)) Then Exit For
                    End If

                Next
                UpdateLowCmdUI(objFields)
                Try
                    aRequest.Value = Nothing
                    aResponse.Value = Nothing

                Finally

                End Try
            Next
        Catch ex As Exception
            objFields.Item("CmdStatus".ToUpper) = "E"
            objFields.Item("ErrCode".ToUpper) = CustomerUrlError.Other_Error
            objFields.Item("ErrMsg".ToUpper) = [Enum].GetName(GetType(CustomerUrlError), CustomerUrlError.Other_Error)
            WriteErrTxtLog.WriteTxtError(ex, objFields.Item("COMPCODE"), _
                                                     objFields.Item("SEQNO"), _
                                                     objFields.Item("Cmd".ToUpper), "Url = " & aUrl.Value, _
                                                     Nothing)
        Finally

            aRequest.Dispose()
            aRequest.Dispose()
            aUrl.Dispose()
            aRet.Dispose()
            aryLicenses.Dispose()
            aNoteCnt.Dispose()
            aChangUrlNum.Dispose()
            aSendUrl.Dispose()
            aJson.Value.Dispose()
            Erase bytData.Value
            aJson.Dispose()
            bytData.Dispose()
            UpdUiAndData(objFields)
            If objFields.Item("CmdStatus".ToUpper) = "S" Then
                If fSendTVMail Then
                    SendTVMail(objFields)
                End If
            End If
            If FThreadsTotal > 0 Then
                Interlocked.Decrement(FThreadsTotal)
            End If

            If objFields.Item("CmdStatus".ToUpper) = "E" Then
                If fIsUseMail Then
                    Interlocked.Increment(fWebSiteErrTotal)
                    If fWebSiteErrTotal >= fReadWebErrCnt Then
                        Interlocked.Exchange(fWebSiteErrTotal, 0)
                        csCommon.SendMail(MDBMailPara.GetDefaultMDBName, _
                                 fMDBMailPassword, "PSErr", fMailInfoTableName, _
                                 fMailToTableName, fMailCCTableName, _
                                 String.Empty, _
                                 String.Empty, False, True)
                    End If
                End If

            End If

            If FThreadsTotal = 0 Then
                'evn.Set()
            End If

        End Try

    End Sub
    Public Function BuilderTVMailMsg(ByVal aCmdInfo As Dictionary(Of String, String),
                                 ByVal aSOInfo As SOInfoClass, ByVal aProductId As String) As String
        Dim aMsg As New ThreadLocal(Of String)
        Dim adr As New ThreadLocal(Of OracleDataReader)
        'FSOIndex(objFieldLst.Item("RightCompId".ToUpper))
        Try
            aMsg.Value = fTVMailMDBMsgTxt
            adr.Value = Nothing
            If (fLstCustomerVar IsNot Nothing) AndAlso (fLstCustomerVar.Count > 0) Then
                Using cn As New OracleConnection(aSOInfo.OraConnectString)
                    cn.Open()
                    Using cmd As New OracleCommand("SELECT * FROM SO563 WHERE " &
                                                                        " TYPE = 1 AND NVL(LOCALLANGUAGE,0) = 0 " &
                                                                        " AND ID = '" & aProductId & "' ", cn)

                        adr.Value = cmd.ExecuteReader(CommandBehavior.CloseConnection)
                        If (adr.Value IsNot Nothing) AndAlso (adr.Value.HasRows) Then
                            adr.Value.Read()
                            For i As Integer = 0 To fLstCustomerVar.Count - 1
                                aMsg.Value = Regex.Replace(aMsg.Value,
                                                         fSplitTVMailMsgChar(0) & fLstCustomerVar.Item(i) & fSplitTVMailMsgChar(1),
                                                         adr.Value.Item(fLstCustomerVar.Item(i)).ToString, RegexOptions.IgnoreCase)

                            Next
                        Else
                            aMsg.Value = Chr(0)
                            WriteErrTxtLog.WriteTVMailErrLog(Nothing, "找不到Epg內容說明資料！命令流水號 = " & aCmdInfo.Item("SEQNO"))
                        End If
                    End Using
                End Using

            End If

            Return aMsg.Value
        Catch ex As Exception
            WriteErrTxtLog.WriteTVMailErrLog(ex, "組合TVMailMessage錯誤！命令流水號 = " & aCmdInfo.Item("SEQNO"))
            Return String.Empty
        Finally
            If (adr.Value IsNot Nothing) AndAlso (Not adr.Value.IsClosed) Then
                adr.Value.Close()
            End If

            aMsg.Dispose()
            adr.Dispose()
        End Try
    End Function
    Private Function GetLogInID(ByVal aCompId As String) As String
        Try
            'TABLEOWNER
            Using cn As New OracleConnection(fUseCompIndex(aCompId).ConnectionString)
                cn.Open()
                Using cmd As New OracleCommand("SELECT LOGINID FROM SO041 " &
                                                                        " WHERE SYSID = " & aCompId, cn)
                    Using dr As OracleDataReader = cmd.ExecuteReader(CommandBehavior.CloseConnection)
                        If (dr.HasRows) Then
                            dr.Read()
                            If IsDBNull(dr.Item("LOGINID")) Then
                                Return String.Empty
                            Else
                                Return dr.Item("LOGINID").ToString & "."
                            End If

                        Else
                            Return String.Empty
                        End If
                        dr.Close()
                    End Using
                End Using
            End Using
        Catch ex As Exception
            WriteErrTxtLog.WriteTVMailErrLog(ex, "取出Owner失敗")
            Return String.Empty


        End Try
    End Function
    Private Sub SendTVMail(ByVal aCmdInfo As Dictionary(Of String, String))
        Try
            Dim thr As New Thread(AddressOf BeginSendTVMail)
            thr.IsBackground = True
            thr.Start(aCmdInfo)
        Catch ex As Exception
            WriteErrTxtLog.WriteTVMailErrLog(ex, "啟動TVMail Thread 失敗")
        End Try


    End Sub
    '#6115 送資料至TVMail By Kin 2011/09/21
    Private Sub BeginSendTVMail(ByVal aCmdInfo As Dictionary(Of String, String))
        Try
            Dim adr As New ThreadLocal(Of OracleDataReader)
            Dim aMsg As New ThreadLocal(Of String)
            Dim aInsSQL As New ThreadLocal(Of String)
            Dim aOraPara As New ThreadLocal(Of OracleParameterCollection)
            adr.Value = Nothing
            aMsg.Value = String.Empty
            aInsSQL.Value = String.Empty
            aOraPara.Value = Nothing
            Try
                Using cn As New OracleConnection(fSOIndex(aCmdInfo.Item("RightCompId".ToUpper)).OraConnectString)
                    cn.Open()
                    Using cmd As New OracleCommand("SELECT Distinct ProductId FROM SO561B " &
                                                                        " WHERE SEQNO = " & aCmdInfo.Item("SEQNO"), cn)
                        adr.Value = cmd.ExecuteReader(CommandBehavior.CloseConnection)

                        While adr.Value.Read
                            If Not IsDBNull(adr.Value.Item("Productid")) Then
                                aMsg.Value = BuilderTVMailMsg(aCmdInfo,
                                                          fSOIndex(aCmdInfo.Item("RightCompId".ToUpper)),
                                                          adr.Value.Item("Productid"))
                                If String.IsNullOrEmpty(
                                    fUseCompIndex(aCmdInfo.Item("CompCode".ToUpper)).LoginID) Then

                                    fUseCompIndex(aCmdInfo.Item("CompCode".ToUpper)).LoginID =
                                            GetLogInID(aCmdInfo.Item("CompCode".ToUpper))

                                End If
                                If aMsg.Value <> Chr(0) Then
                                    If String.IsNullOrEmpty(aMsg.Value) Then
                                        WriteErrTxtLog.WriteTVMailErrLog(Nothing, "未設定傳送訊息！命令序號 = " &
                                                                        aCmdInfo.Item("SEQNO").ToString)
                                    Else
                                        'aInsSQL.Value = String.Format(fInsSendNagra,
                                        '                              fUseCompIndex(aCmdInfo("CompCode".ToUpper)).LoginID,
                                        '                                "'E2'", "'" & aCmdInfo("ICCNo".ToUpper) & "'", "'" & aCmdInfo("STBSNo".ToUpper) & "'",
                                        '                                aMsg.Value, "W", aCmdInfo("UpdEn".ToUpper), "sysdate",
                                        '                                "1", "S_SendNagra_SeqNo.NextVal")
                                        aInsSQL.Value = String.Format(fInsSendNagra, fUseCompIndex(aCmdInfo("CompCode".ToUpper)).LoginID)
                                        Using cnSo As New OracleConnection(fUseCompIndex(aCmdInfo("CompCode".ToUpper)).ConnectionString)
                                            cnSo.Open()
                                            Using cmdSo As OracleCommand = cnSo.CreateCommand()
                                                cmdSo.Parameters.Clear()
                                                cmdSo.CommandText = aInsSQL.Value
                                                cmdSo.Parameters.Add("HIGH_LEVEL_CMD_ID", OracleType.VarChar).Value = "E2"
                                                cmdSo.Parameters.Add("ICC_NO", OracleType.VarChar).Value = aCmdInfo("ICCNO".ToUpper)
                                                cmdSo.Parameters.Add("STB_NO", OracleType.VarChar).Value = aCmdInfo("STBSNO".ToUpper)
                                                cmdSo.Parameters.Add("NOTES", OracleType.VarChar).Value = aMsg.Value
                                                cmdSo.Parameters.Add("CMD_STATUS", OracleType.VarChar).Value = "W"
                                                cmdSo.Parameters.Add("OPERATOR", OracleType.VarChar).Value = aCmdInfo("UpdEn".ToUpper)
                                                cmdSo.Parameters.Add("UPDTIME", OracleType.DateTime).Value = Now

                                                cmdSo.Parameters.Add("MIS_IRD_CMD_DATA", OracleType.VarChar).Value = "1"
                                                cmdSo.Parameters.Add("SEQNO", OracleType.Int32).Value =
                                                        Int32.Parse(GetSendNagra_SEQ(fUseCompIndex(aCmdInfo("CompCode".ToUpper)).ConnectionString))
                                                cmdSo.Parameters.Add("COMPCODE", OracleType.Int32).Value = Int32.Parse(aCmdInfo("CompCode".ToUpper))
                                                cmdSo.Parameters.Add("RESENTTIMES", OracleType.Int32).Value = 1
                                                cmdSo.Parameters.Add("STB_FLAG", OracleType.Int32).Value = 0
                                                cmdSo.Parameters.Add("RETURNWORKORDER", OracleType.Int32).Value = 0
                                                cmdSo.Parameters.Add("MODETYPE", OracleType.VarChar).Value = "E2"
                                                cmdSo.ExecuteNonQuery()
                                            End Using
                                            cnSo.Close()
                                        End Using

                                    End If
                                End If

                            Else
                                WriteErrTxtLog.WriteTVMailErrLog(Nothing, "ProductId = Nothing 序號 = " &
                                                                 aCmdInfo.Item("SEQNO".ToUpper))

                            End If


                        End While
                        adr.Value.Close()
                    End Using
                End Using

            Catch ex As Exception
                WriteErrTxtLog.WriteTVMailErrLog(ex, "傳送TVMail失敗 ！序號 = " &
                                                                 aCmdInfo.Item("SEQNO".ToUpper))

            Finally
                If (adr IsNot Nothing) AndAlso (Not adr.Value.IsClosed) Then
                    adr.Value.Close()
                End If
                adr.Dispose()
                aMsg.Dispose()
                aInsSQL.Dispose()
                aOraPara.Dispose()
            End Try

        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, "傳送TVMail失敗")
        End Try
    End Sub
    Private Function GetOracleSysDate(ByVal aConnectString As String) As String
        Try
            Using cn As New OracleConnection(aConnectString)
                cn.Open()
                Using cmd As OracleCommand = cn.CreateCommand
                    cmd.CommandText = "SELECT Sysdate FROM DUAL"
                    Return cmd.ExecuteScalar.ToString
                End Using
            End Using
        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, "取出系統時間失敗")
            Return Now
        End Try
    End Function
    Private Function GetSendNagra_SEQ(ByVal aConnectString As String) As String
        Try
            Using cn As New OracleConnection(aConnectString)
                cn.Open()
                Using cmd As OracleCommand = cn.CreateCommand
                    cmd.CommandText = "SELECT S_SendNagra_SeqNo.NEXTVAL FROM DUAL"
                    Return cmd.ExecuteScalar.ToString
                End Using
            End Using
        Catch ex As Exception
            WriteErrTxtLog.WriteTVMailErrLog(ex, "取出SendNagra_seqno失敗")
        End Try
    End Function
    'Friend Function GetTVMailCommand(ByVal COMOwner As String) As String
    '    Dim strSQL As String
    '    strSQL = String.Format("Insert Into {1}Send_Nagra (HIGH_LEVEL_CMD_ID,ICC_NO,STB_NO," & _
    '                     "NOTES,CMD_STATUS,OPERATOR,UPDTIME,MIS_IRD_CMD_DATA,SEQNO,COMPCODE,RESENTTIMES,STB_FLAG," & _
    '                    "RETURNWORKORDER,MODETYPE) Values ({0}0,{0}1,{0}2,{0}3,{0}4,{0}5,{0}6,{0}7,{0}8,{0}9,{0}10,{0}11,{0}12,{0}13)",
    '            Sign, COMOwner)
    '    Return strSQL

    '    DAO.ExecNqry(_DAL.GetTVMailCommand(COMOwner), New Object() {"E2", FacilityRow.Item("SmartCardNo"), FacilityRow.Item("FaciSNo"), _
    '                                                                strMessage, "W", LoginInfo.EntryName, UpdTime, 1, Integer.Parse(util.GetSequenceNo("S_SendNagra_SeqNo", 8)), _
    '                                                                 FacilityRow.Item("CompCode"), 1, 0, 0, "E2"})

    'End Function

    'Public Sub SndMailThread()
    '    Try

    '        csCommon.SendMail(MDBMailPara.GetDefaultMDBName, _
    '                              fMDBMailPassword, "PSErr", fMailInfoTableName, _
    '                              fMailToTableName, fMailCCTableName, String.Empty, String.Empty)
    '        WriteErrTxtLog.WriteSndEmailLog("成功寄出Mail")

    '    Catch ex As Exception
    '        WriteErrTxtLog.WriteSndEmailLog(ex.Message)
    '    End Try

    'End Sub

    Private Sub InsSO561AData(ByRef objFieldLst As Dictionary(Of String, String), ByRef Cn As OracleConnection, _
                              ByRef tra As OracleTransaction)
        Dim aInsSQL As New ThreadLocal(Of String)

        aInsSQL.Value = "INSERT  INTO SO561A ( " & fSO561AField & " )   " & _
                        "  SELECT " & fSO561AField & "  FROM SO560 WHERE SEQNO = " & _
                        objFieldLst.Item("SEQNO".ToUpper)
        Try
            If (objFieldLst.Item("CMDSTATUS".ToUpper) = "S") OrElse (tra IsNot Nothing) Then
                Using cmd As OracleCommand = Cn.CreateCommand
                    cmd.Transaction = tra
                    cmd.CommandText = aInsSQL.Value
                    cmd.ExecuteNonQuery()
                    cmd.CommandText = "DELETE SO560 WHERE SEQNO = " & objFieldLst.Item("SEQNO".ToUpper)
                    cmd.ExecuteNonQuery()

                End Using

            End If

        Catch ex As Exception

            WriteErrTxtLog.WriteTxtError(ex, "[ 公司別 >" & objFieldLst.Item("CompCode".ToUpper) & "] 命令 [ " & objFieldLst.Item("SEQNO".ToUpper) & " ]", Nothing)
            UpdateSysErrUI(ex, " [ 公司別 > " & objFieldLst.Item("CompCode".ToUpper) & "] 命令 [ " & objFieldLst.Item("SEQNO".ToUpper) & " ]")
        Finally
            aInsSQL.Dispose()
        End Try
    End Sub
    Private Sub UPDSO560Data(ByRef objFieldLst As Dictionary(Of String, String))
        Dim aSO As New ThreadLocal(Of SOInfoClass)
        aSO.Value = fSOIndex(objFieldLst.Item("RightCompId".ToUpper))
        aSO.Value.UseCompId = objFieldLst.Item("CompCode".ToUpper)
        Dim tra As New ThreadLocal(Of OracleTransaction)
        tra.Value = Nothing
        Try
            If objFieldLst.Item("CmdStatus".ToUpper).ToUpper = "W" OrElse _
                objFieldLst.Item("CmdStatus".ToUpper).ToUpper = "S1" Then
                Exit Sub
            End If
            If objFieldLst.Item("CMDSTATUS".ToUpper) = "E".ToUpper Then
                objFieldLst.Item("PSERRCNT".ToUpper) = (Int32.Parse("0" & objFieldLst.Item("PSERRCNT")) + 1).ToString
                If fErrResumeLst.Count > 0 Then
                    If fErrResumeLst.Contains(objFieldLst.Item("ErrCode".ToUpper)) AndAlso _
                        Int32.Parse(objFieldLst.Item("PSERRCNT".ToUpper)) < fNoneResumeCnt Then
                        objFieldLst.Item("CmdStatus".ToUpper) = "S1"
                    End If
                End If
                If fIsUseMail AndAlso fCmdErrSndMail Then
                    csCommon.SendMail(MDBMailPara.GetDefaultMDBName, _
                                 fMDBMailPassword, "PS_Cmd_Err", fMailInfoTableName, _
                                 fMailToTableName, fMailCCTableName, _
                                 "Push Server Gateway 命令有錯誤", _
                                 String.Format("命令序號 {0} 有錯誤 ", objFieldLst.Item("SEQNO".ToUpper)), _
                                 False, True)
                End If
            End If
            Using cn As New OracleConnection(aSO.Value.OraConnectString)
                cn.Open()
                tra.Value = cn.BeginTransaction
                Using cmd As OracleCommand = cn.CreateCommand

                    Dim aSQL As New ThreadLocal(Of String)
                    aSQL.Value = "UPDATE SO560 SET CMDSTATUS=:CMDSTATUS , " & _
                                    "AUTHTIME=:AUTHTIME," & _
                                    " ERRCODE=:ERRCODE,ERRMSG=:ERRMSG , " & _
                                    " AUTHREQUESTTIME=:AUTHREQUESTTIME," & _
                                    " PSERRCNT = :PSERRCNT  " & _
                                    " WHERE SEQNO=:SEQNO"
                    Dim pSeqNo As New ThreadLocal(Of OracleParameter)
                    pSeqNo.Value = New OracleParameter("SEQNO", OracleType.Number)
                    Dim pCmdStatus As New ThreadLocal(Of OracleParameter)
                    pCmdStatus.Value = New OracleParameter("CMDSTATUS".ToUpper, OracleType.VarChar)
                    Dim pErrCode As New ThreadLocal(Of OracleParameter)
                    pErrCode.Value = New OracleParameter("ErrCode".ToUpper, OracleType.VarChar)
                    Dim pErrMsg As New ThreadLocal(Of OracleParameter)
                    pErrMsg.Value = New OracleParameter("ErrMsg".ToUpper, OracleType.VarChar)
                    Dim pAuthTime As New ThreadLocal(Of OracleParameter)
                    pAuthTime.Value = New OracleParameter("AuthTime".ToUpper, OracleType.DateTime)
                    Dim pAuthRequestTime As New ThreadLocal(Of OracleParameter)
                    pAuthRequestTime.Value = New OracleParameter("AuthRequestTime", OracleType.DateTime)
                    Dim pPsErrCnt As New ThreadLocal(Of OracleParameter)
                    pPsErrCnt.Value = New OracleParameter("PSERRCNT", OracleType.Number)
                    pPsErrCnt.Value.Value = 0
                    If String.IsNullOrEmpty(objFieldLst.Item("AuthTime".ToUpper)) Then
                        pAuthTime.Value.Value = DBNull.Value
                    Else
                        pAuthTime.Value.Value = Date.Parse(objFieldLst.Item("AuthTime".ToUpper), _
                                                            CultureInfo.CreateSpecificCulture("zh-TW"))
                    End If
                    If String.IsNullOrEmpty(objFieldLst.Item("AuthRequestTime".ToUpper)) Then
                        pAuthRequestTime.Value.Value = DBNull.Value
                    Else
                        pAuthRequestTime.Value.Value = Date.Parse(objFieldLst.Item("AuthRequestTime".ToUpper), _
                                                            CultureInfo.CreateSpecificCulture("zh-TW"))
                    End If
                    pSeqNo.Value.Value = Int32.Parse(objFieldLst.Item("SEQNO".ToUpper))
                    pCmdStatus.Value.Value = objFieldLst.Item("CmdStatus".ToUpper)
                    If Not String.IsNullOrEmpty(objFieldLst.Item("PSERRCNT".ToUpper)) Then
                        pPsErrCnt.Value.Value = Int32.Parse(objFieldLst.Item("PSERRCNT".ToUpper))
                    End If

                    If String.IsNullOrEmpty(objFieldLst.Item("ErrCode".ToUpper)) Then
                        pErrCode.Value.Value = DBNull.Value
                    Else
                        pErrCode.Value.Value = objFieldLst.Item("ErrCode".ToUpper)
                    End If
                    If String.IsNullOrEmpty(objFieldLst.Item("ErrMsg".ToUpper)) Then
                        pErrMsg.Value.Value = DBNull.Value
                    Else
                        If fErrCodeLst.ContainsKey(objFieldLst.Item("ErrCode".ToUpper)) Then
                            objFieldLst.Item("ErrMsg".ToUpper) = fErrCodeLst.Item(objFieldLst.Item("ErrCode".ToUpper))
                        End If
                        pErrMsg.Value.Value = objFieldLst.Item("ErrMsg".ToUpper)
                    End If
                    cmd.Transaction = tra.Value
                    cmd.CommandText = aSQL.Value
                    cmd.Parameters.Add(pSeqNo.Value)
                    cmd.Parameters.Add(pCmdStatus.Value)
                    cmd.Parameters.Add(pErrCode.Value)
                    cmd.Parameters.Add(pErrMsg.Value)
                    cmd.Parameters.Add(pAuthTime.Value)
                    cmd.Parameters.Add(pAuthRequestTime.Value)
                    cmd.Parameters.Add(pPsErrCnt.Value)
                    cmd.ExecuteNonQuery()
                    cmd.Parameters.Clear()
                    If objFieldLst.Item("CMDSTATUS".ToUpper) = "S".ToUpper Then
                        InsSO561AData(objFieldLst, cn, tra.Value)
                    End If

                    tra.Value.Commit()
                    aSQL.Dispose()
                    pCmdStatus.Dispose()
                    pSeqNo.Dispose()
                    pErrCode.Dispose()
                    pErrMsg.Dispose()
                    pAuthTime.Dispose()
                    pAuthRequestTime.Dispose()
                    pPsErrCnt.Dispose()
                End Using
            End Using
            'cn.Value = New OracleConnection(objSOCMDInfo.OraConnectString)
        Catch ex As Exception
            If tra IsNot Nothing Then
                tra.Value.Rollback()
            End If

            objFieldLst.Item("CmdStatus".ToUpper) = "E"
            objFieldLst.Item("ErrCode".ToUpper) = CustomerUrlError.UpdateSO560_Error
            objFieldLst.Item("ErrMsg".ToUpper) = [Enum].GetName(GetType(CustomerUrlError), CustomerUrlError.UpdateSO560_Error)
            WriteErrTxtLog.WriteTxtError(ex, "[ 公司別 >" & objFieldLst.Item("CompCode".ToUpper) & "] 命令 [ " & objFieldLst.Item("SEQNO".ToUpper) & " ]", Nothing)
            UpdateSysErrUI(ex, " [ 公司別 > " & objFieldLst.Item("CompCode".ToUpper) & "] 命令 [ " & objFieldLst.Item("SEQNO".ToUpper) & " ]")
        Finally
            SyncLock aSO.Value.WaitProcessNumber.GetType
                If aSO.Value.WaitProcessNumber > 0 Then
                    aSO.Value.WaitProcessNumber -= 1
                End If
                If fUseCompIndex(aSO.Value.UseCompId).WaitProcessNumber > 0 Then
                    fUseCompIndex.Item(objFieldLst.Item("CompCode".ToUpper)).WaitProcessNumber -= 1
                End If
                If fUseCompIndex(objFieldLst.Item("CompCode".ToUpper)).ProcessingNumber > 0 Then
                    fUseCompIndex(objFieldLst.Item("CompCode".ToUpper)).ProcessingNumber -= 1
                End If
            End SyncLock
            tra.Value.Dispose()
            tra.Dispose()
            aSO.Dispose()
        End Try
    End Sub
    Private Sub UpdUiAndData(ByVal objFieldLst As Dictionary(Of String, String))
        Dim aSO As New ThreadLocal(Of SOInfoClass)
        aSO.Value = fSOIndex(objFieldLst.Item("RightCompId".ToUpper))
        aSO.Value.UseCompId = objFieldLst.Item("CompCode".ToUpper)
        Try
            UPDSO560Data(objFieldLst)
            UpdateSOStatusUI(aSO.Value, SOStatus.NotUpdateImg)
            If objFieldLst.Item("CMDSTATUS".ToUpper) <> "S1".ToUpper Then
                UpdateHightCmdUI(objFieldLst, UpdCmdMode.Add)
            End If

            'UpdateLowCmdUI(FMainForm, objFieldLst)
            Interlocked.Increment(fUpdTotal)
            If fUpdTotal >= fProcessNumber Then
                Interlocked.Exchange(fUpdTotal, 0)
                'evn.Set()
                Application.DoEvents()
            End If

            '
            'UpdateSOStatusUI(FMainForm, FMainForm.treLstSO, aSO.Value, SOStatus.NotUpdateImg)
        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, "[ 公司別 >" & objFieldLst.Item("CompCode".ToUpper) & "] 命令 [ " & objFieldLst.Item("SEQNO".ToUpper) & " ]", Nothing)
            UpdateSysErrUI(ex, " [ 公司別 > " & objFieldLst.Item("CompCode".ToUpper) & "] 命令 [ " & objFieldLst.Item("SEQNO".ToUpper) & " ]")
        Finally
            aSO.Dispose()

        End Try
    End Sub
    Private Sub UpdateProcessCmd(ByVal objFieldLst As Dictionary(Of String, String))
        Dim aSO As New ThreadLocal(Of SOInfoClass)
        aSO.Value = fSOIndex(objFieldLst.Item("RightCompId".ToUpper))
        aSO.Value.UseCompId = objFieldLst.Item("CompCode".ToUpper)

        If fThreadStop Then
            Exit Sub
        End If

        Try
            aSO.Value.UseCompId = objFieldLst.Item("COMPCODE".ToUpper)
            'UpdateSOStatusUI(FMainForm, FMainForm.treLstSO, aSO.Value, SOStatus.NotUpdateImg)
            '連結網站
            ReadWebSite(objFieldLst)

        Catch ex As Exception
            If Not IsConnectOK(aSO.Value.OracleCn) Then
                If aSO.Value.ConnectionOK Then
                    UpdateSOStatusUI(aSO.Value, SOStatus.NO)
                End If
            Else
                UpdateSOStatusUI(aSO.Value, SOStatus.Yes)
            End If
            WriteErrTxtLog.WriteTxtError(ex, objFieldLst.Item("COMPCODE"), objFieldLst.Item("SEQNO"), Nothing, Nothing, Nothing)
        Finally

            'objFieldLst.Clear()
            'objFieldLst = Nothing
            aSO.Dispose()

        End Try


    End Sub


    Public Class JsonClass
        Implements IDisposable


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




#Region "IDisposable Support"
        Private disposedValue As Boolean ' 偵測多餘的呼叫

        ' IDisposable
        Protected Overridable Sub Dispose(ByVal disposing As Boolean)
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
    Public Enum CustomerUrlError
        ReturnLicenses_Length_Error = -7777
        ConnectUrl_Error = -7776
        UpdateSO560_Error = -7775
        Data_Format_Error = -7774
        Other_Error = -7773
    End Enum

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
