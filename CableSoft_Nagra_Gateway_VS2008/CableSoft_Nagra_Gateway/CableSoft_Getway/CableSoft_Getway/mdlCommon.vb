Imports System
Imports Microsoft.Win32
Imports System.Security
Imports DevExpress.Utils
Imports DevExpress.XtraEditors
Imports DevExpress.XtraTreeList
Imports System.Data
Imports System.IO
Imports System.Text
Imports System.Data.OracleClient
Imports DevExpress.XtraTreeList.Nodes
Imports CableSoft.Gateway.csException
Imports CableSoft.Gateway.Common
Imports CableSoft.Gateway.SystemIO
Imports System.Net
Module mdlCommon
    Public Enum SOStatus
        Initial = 0
        Yes = 1
        NO = 2
        Key = 3
        NotUpdateImg = 4

    End Enum
    Public Enum RunStatus
        Initial = 0
        Execute = 1
        StopCmd = 2
        LoadComplete = 3
    End Enum
    Public FMDBFile As String = String.Empty
    Public FMDBPassword As String = String.Empty
    Public FSysTable As DataTable = Nothing
    Public FSOTable As DataTable = Nothing
    Public FNagraParaTable As DataTable = Nothing
    Public FComparedErrorTable As DataTable = Nothing
    Public FSOInfoList As List(Of SOInfoClass) = Nothing
    Public Const SystemTableName As String = "SysOpt"
    Public Const SOTableName As String = "SO"
    Public Const NagraParaName As String = "NagraParaSet"
    Public Const ComparedErrorName As String = "ComparedError"
    Public FThreadStop As Boolean = False
    '---------------以下為系統公用屬性-------------------------------------
    Public FAppTitle As String = String.Empty
    Public FAutoRunGW As Boolean = False
    Public FUseTray As Boolean = False
    Public FShowResource As Boolean = False
    Public FUseHotKey As Boolean = False
    Public FAutoConnectTime As Int32
    Public FReadDataTime As Integer = 100
    Public FProcessNumber As Int32 = 100
    Public FShowDataCount As Int32 = 200
    Public FClearDataCount As Int32 = 20
    Public FMaxThread As Int32 = 25
    Public FDBPoolMinNumber As Int32 = 0
    Public FDBPoolMaxNumber As Int32 = 100
    Public FDBPoolLiveTime As Int32 = 0
    Public FTestNagraSock As Int32 = 30
    Public FNagraIPAddress As IPAddress = IPAddress.Parse("127.0.0.1")
    Public FNagraServer As String = String.Empty
    Public FNagraPort As Integer = 0
    Public FSndSourceId As String = String.Empty
    Public FSndDestId As String = String.Empty
    Public FSndMopPPId As String = String.Empty
    Public FRcvSourceId As String = String.Empty
    Public FRcvDestId As String = String.Empty
    Public FRcvMopPPId As String = String.Empty
    Public FSenErrStopCnt As Int32 = 3
    Public FRcvErrStopCnt As Int32 = 3
    Public FCheckStatusTime As Int32 = 3
    Public FSndDelayTime As Single = 0.1
    Public FSndTimeout As Int32 = 0
    Public FRcvTimeout As Int32 = 0
    Public FErrCodeLst As New Dictionary(Of String, String)
    '-------------------------------------------------------------------
    Public FLstSysMsg As List(Of String)
    Public FInsDate As String = String.Empty
    Public FRegDate As String = String.Empty
    Public FRegOK As Boolean = False
    


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
    End Enum
    '----------------------------------------------------------------------
    '-----------------------------------SO的連線資訊-----------------------

    Public Class SOInfoClass
        Private _OraConnectString As String
        Private _WaitProcessNumber As Int32
        Private _ProcessingNumber As Int32
        Private _CompID As Int32
        Private _CompName As String
        Private _ConnectionOK As Boolean
        Private _OracleCn As OracleConnection = Nothing
        Private _FirstLoad As Boolean = True
        Public Sub New()

        End Sub
        Public Property OraConnectString() As String
            Get
                Return _OraConnectString
            End Get
            Set(ByVal value As String)
                _OraConnectString = value
                _OracleCn = New OracleConnection(_OraConnectString)
            End Set
        End Property
        Public Property WaitProcessNumber() As Int32
            Get
                Return _WaitProcessNumber
            End Get
            Set(ByVal value As Int32)
                _WaitProcessNumber = value
            End Set
        End Property
        Public Property ProcessingNumber() As Int32
            Get
                Return _ProcessingNumber
            End Get
            Set(ByVal value As Int32)
                _ProcessingNumber = value
            End Set
        End Property
        Public Property CompID() As Int32
            Get
                Return _CompID
            End Get
            Set(ByVal value As Int32)
                _CompID = value
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
        Public Property ConnectionOK() As Boolean
            Get
                Return _ConnectionOK
            End Get
            Set(ByVal value As Boolean)
                _ConnectionOK = value
            End Set
        End Property
        Public Property OracleCn() As OracleConnection
            Get
                Return _OracleCn
            End Get
            Set(ByVal value As OracleConnection)
                _OracleCn = value
            End Set
        End Property

        Public Property FirstLoad() As Boolean
            Get
                Return _FirstLoad
            End Get
            Set(ByVal value As Boolean)
                _FirstLoad = value
            End Set
        End Property
    End Class
    '----------------------------------------------------------------------
    '註刪或刪除機碼

    Public Sub ProcAutoRun(ByVal AutoExec As Boolean)
        Dim RegRoot As RegistryKey = Nothing
        Dim RegKey As RegistryKey = Nothing

        Try
            RegRoot = Registry.LocalMachine

            RegKey = RegRoot.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Run", _
                                        RegistryKeyPermissionCheck.ReadWriteSubTree, _
                                        AccessControl.RegistryRights.SetValue)


            Dim App As String = String.Format("{0} /AutoExec", Application.ExecutablePath)
            If AutoExec Then
                RegKey.SetValue("Cablesoft VOD Gateway", App, RegistryValueKind.String)
            Else
                RegKey.DeleteValue("Cablesoft VOD Gateway", False)
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
            'Msg.Show(ex.ToString(), "訊息", MsgBtn.OK, MsgIco.Error)
            'Bs.Write2MdbLog(Nothing, Nothing, Nothing, ex)
        Finally
            RegKey.Close()
            RegRoot.Close()
        End Try
    End Sub
    Public Sub UpdMainFormUI(ByVal aForm As rfrmMain)
        Try
            If aForm.InvokeRequired Then
                Dim act As New Action(Of Form)(AddressOf UpdMainFormUI)
                aForm.Invoke(act, aForm)

            Else
                aForm.Text = FAppTitle


            End If
        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
        End Try
    End Sub
    '設定錯誤對照表
    Public Function SetErrCode() As Boolean
        Try
            FComparedErrorTable = CableSoft.GateWay.SystemIO.GetwaySys.ReadErrorCode(FMDBFile, FMDBPassword)
            FErrCodeLst.Clear()
            For i As Int32 = 0 To FComparedErrorTable.Rows.Count - 1
                FErrCodeLst.Add(FComparedErrorTable.Rows(i)("ErrorCode"), _
                                FComparedErrorTable.Rows(i)("ErrorName"))
            Next
            Return True
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function
    '設定Nagra參數值
    Public Function SetNagraPara() As Boolean
        Try
            FNagraParaTable = CableSoft.Gateway.SystemIO.GetwaySys.ReadNagraParaSet(FMDBFile, FMDBPassword)
            If (FNagraParaTable IsNot Nothing) AndAlso (FNagraParaTable.Rows.Count > 0) Then
                With FNagraParaTable
                    FNagraServer = GetFieldValue(.Rows(0), "NagraIPAddress")
                    FNagraPort = Convert.ToInt32(GetFieldValue(.Rows(0), "NagraPort"))
                    FSndSourceId = GetFieldValue(.Rows(0), "SndSourceId")
                    FSndDestId = GetFieldValue(.Rows(0), "SndDestId")
                    FSndMopPPId = GetFieldValue(.Rows(0), "SndMopPPId")
                    FRcvSourceId = GetFieldValue(.Rows(0), "RcvSourceId")
                    FRcvDestId = GetFieldValue(.Rows(0), "RcvDestId")
                    FRcvMopPPId = GetFieldValue(.Rows(0), "RcvMopPPId")
                    FSenErrStopCnt = Convert.ToInt32(GetFieldValue(.Rows(0), "SenErrStopCnt"))
                    FRcvErrStopCnt = Convert.ToInt32(GetFieldValue(.Rows(0), "RcvErrStopCnt"))
                    FCheckStatusTime = Convert.ToInt32(GetFieldValue(.Rows(0), "CheckStatusTime"))
                    FSndDelayTime = Convert.ToSingle(GetFieldValue(.Rows(0), "SndDelayTime"))
                    FSndTimeout = Convert.ToInt32(GetFieldValue(.Rows(0), "SndTimeout"))
                    FRcvTimeout = Convert.ToInt32(GetFieldValue(.Rows(0), "RcvTimeout"))
                    Try
                        FNagraIPAddress = IPAddress.Parse(FNagraServer)
                    Catch ex As Exception
                        FNagraIPAddress = IPAddress.Parse("127.0.0.1")
                    End Try

                End With
            Else
                Cablesoft.Gateway.SystemIO.GetwaySys.DefaultNagra(FMDBFile, FMDBPassword, NagraParaName)
                SetNagraPara()
            End If
            Return True
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function
    '設定公用系統值
    Public Function SetSystem() As Boolean
        Try
            FSysTable = CableSoft.GateWay.SystemIO.GetwaySys.ReadSystemIO(FMDBFile, FMDBPassword)
            If (FSysTable IsNot Nothing) AndAlso (FSysTable.Rows.Count > 0) Then
                With FSysTable
                    FAppTitle = GetFieldValue(.Rows(0), "AppCaption")
                    FAutoRunGW = GetFieldValue(.Rows(0), "AutoRunGw") = "1"
                    FUseTray = GetFieldValue(.Rows(0), "UseTray") = 1
                    FShowResource = GetFieldValue(.Rows(0), "ShowResource") = 1
                    FUseHotKey = GetFieldValue(.Rows(0), "UseHotKey") = 1
                    FAutoConnectTime = GetFieldValue(.Rows(0), "AutoConnectTime")
                    FReadDataTime = GetFieldValue(.Rows(0), "ReadDataTime")
                    FProcessNumber = GetFieldValue(.Rows(0), "ProcessNumber")
                    FShowDataCount = GetFieldValue(.Rows(0), "ShowDataCount")
                    FClearDataCount = GetFieldValue(.Rows(0), "ClearDataCount")
                    FMaxThread = GetFieldValue(.Rows(0), "MaxThread")

                    FDBPoolMaxNumber = GetFieldValue(.Rows(0), "DBPoolMaxNumber")
                    FDBPoolMinNumber = GetFieldValue(.Rows(0), "DBPoolMinNumber")
                    FDBPoolLiveTime = GetFieldValue(.Rows(0), "DBPoolLiveTime")
                    FTestNagraSock = GetFieldValue(.Rows(0), "TestSocketTime")
                    'Try
                    '    FNagraIPAddress = IPAddress.Parse(GetFieldValue(.Rows(0), "NagraIPAddress"))
                    'Catch ex As Exception
                    '    FNagraIPAddress = IPAddress.Parse("127.0.0.1")
                    'End Try
                    'FNagraPort = GetFieldValue(.Rows(0), "NagraPort")

                    '設定顯示在畫面的訊息
                    SetSysMsgString()

                End With
            Else
                CableSoft.GateWay.SystemIO.GetwaySys.DefaultSystem(FMDBFile, FMDBPassword, SystemTableName)
                SetSystem()
            End If
            Return True
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Function
    Private Sub SetSysMsgString()
        Try
            If FLstSysMsg Is Nothing Then
                FLstSysMsg = New List(Of String)
            Else
                FLstSysMsg.Clear()
            End If
            FLstSysMsg.Add("Parent讀取系統參數設定..")
            FLstSysMsg.Add(String.Format("Windows 開機後自動執行 Gateway : {0}", IIf(FAutoRunGW, "是", "否")))
            FLstSysMsg.Add(String.Format("最小化後常駐在系統工作列上 : {0}", IIf(FUseTray, "是", "否")))
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub
    '設定SO
    Public Function SetSO() As Boolean
        Try


            Dim _ConnectString As String = "Data Source={0};Persist Security Info=True;" & _
                "User ID={1};Password={2};Min Pool Size={3};" & _
                "Max Pool Size={4};Unicode=True;Load Balance Timeout={5}"
            If FSOInfoList IsNot Nothing Then
                FSOInfoList.Clear()
            End If

            FSOTable = CableSoft.Gateway.SystemIO.GetwaySys.ReadSO(FMDBFile, FMDBPassword)
            Dim row() As DataRow = FSOTable.Select("SelectID=1")
            If (FSOTable IsNot Nothing) AndAlso (row.Length > 0) Then
                FSOInfoList = New List(Of SOInfoClass)
                For i As Int32 = 0 To row.Length - 1
                    Dim obj As New SOInfoClass()
                    obj.OraConnectString = String.Format(_ConnectString, _
                        row(i)("SID"), row(i)("UserId"), _
                        row(i)("UserPassword"), "0", FDBPoolMaxNumber.ToString, FDBPoolLiveTime)
                    obj.ProcessingNumber = 0
                    obj.WaitProcessNumber = 0
                    obj.CompID = Int32.Parse(row(i)("CompID"))
                    obj.CompName = row(i)("CompName")
                    obj.ConnectionOK = False

                    FSOInfoList.Add(obj)
                Next

            Else
                Throw New Exception("無任何SO資料！")
            End If
            Return True
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function
    Public Sub DisableControl(ByVal aParentForm As rfrmMain, ByVal aStatus As RunStatus)
        On Error Resume Next
        If aParentForm.InvokeRequired Then
            Dim act As New Action(Of Form, RunStatus)(AddressOf DisableControl)
            aParentForm.Invoke(act, aParentForm, aStatus)
        Else
            Select Case aStatus
                Case RunStatus.Initial
                    aParentForm.bbiExec.Enabled = False
                    aParentForm.bbiOpt.Enabled = False
                    aParentForm.bbiStop.Enabled = False
                    aParentForm.bsiView.Enabled = False
                    aParentForm.bbiAbout.Enabled = False
                    aParentForm.bbiQuit.Enabled = False
                    aParentForm.biSets.Enabled = False
                Case RunStatus.Execute
                    aParentForm.bbiExec.Enabled = False
                    aParentForm.bbiStop.Enabled = True
                    aParentForm.bbiOpt.Enabled = False
                    aParentForm.bsiView.Enabled = True
                    aParentForm.bbiAbout.Enabled = False
                    aParentForm.bbiQuit.Enabled = False
                    aParentForm.biSets.Enabled = False
                Case RunStatus.StopCmd, RunStatus.LoadComplete
                    aParentForm.bbiExec.Enabled = True
                    aParentForm.bbiStop.Enabled = False
                    aParentForm.bbiOpt.Enabled = True
                    aParentForm.bsiView.Enabled = True
                    aParentForm.bbiAbout.Enabled = True
                    aParentForm.bbiQuit.Enabled = True
                    aParentForm.biSets.Enabled = True
            End Select
            aParentForm.mnuExec.Enabled = aParentForm.bbiExec.Enabled
            aParentForm.mnuStop.Enabled = aParentForm.bbiStop.Enabled
            aParentForm.mnuOpt.Enabled = aParentForm.bbiOpt.Enabled
            aParentForm.mnuAbout.Enabled = aParentForm.bbiAbout.Enabled
            aParentForm.mnuQuit.Enabled = aParentForm.bbiQuit.Enabled
            aParentForm.mnuAbout.Enabled = aParentForm.bbiAbout.Enabled
        End If
    End Sub
    '測試是否可連線
    Public Function IsConnectOK(ByVal obj As OracleConnection, Optional ByRef aErrMsg As Object = Nothing) As Boolean
        Dim aRet As Boolean = False
        Try


            If obj Is Nothing Then
                Return False
            End If
            If obj.State = ConnectionState.Closed Then
                Try
                    obj.Open()
                    obj.Close()
                    aRet = True
                Catch ex As Exception
                    If aErrMsg IsNot Nothing Then
                        aErrMsg = ex.Message
                    End If
                    aRet = False
                End Try
            Else
                aRet = True
            End If
            Return aRet

        Catch ex As Exception

            Return False
        End Try

    End Function


    Private Function GetFieldValue(ByVal row As DataRow, ByVal FieldName As String) As String
        Try
            If Not DBNull.Value.Equals(row.Item(FieldName)) Then
                Return System.Convert.ToString(row.Item(FieldName))
            Else
                Return Nothing
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Function
    '-----------------------------以下為更新畫面的Function--------------------------------
    '更新系統台連線訊息畫面
    Public Sub UpdateSODatabase(ByVal afrm As rfrmMain, _
                                ByVal aId As String, _
                                ByVal aObj As SOInfoClass, _
                                ByVal aSOStatus As SOStatus)
        Try

            If afrm.InvokeRequired Then

                Dim act As New Action(Of Form, String, SOInfoClass, SOStatus)(AddressOf UpdateSODatabase)
                afrm.Invoke(act, afrm, aId, aObj, aSOStatus)

            Else
                Try
                    afrm.TreLstConnStatus.BeginUpdate()

                    afrm.TreLstConnStatus.BeginUnboundLoad()

                    Dim trlParent As TreeListNode = Nothing
                    Dim trlChild As TreeListNode = Nothing


                    If ((afrm.TreLstConnStatus.Nodes.Count - FShowDataCount) > FClearDataCount) Then
                        afrm.TreLstConnStatus.Nodes.RemoveAt(afrm.TreLstConnStatus.Nodes.Count - 1)
                    End If
                    If aSOStatus = SOStatus.NotUpdateImg Then
                        trlParent = afrm.TreLstConnStatus.AppendNode(Nothing, -1)
                        trlParent.SetValue("SOInfo", Format(Now, "yyyy/MM/dd hh:mm:ss") & _
                                           " > 系統台 : [ " & aObj.CompName & " ] 資料庫連接中..")
                        trlParent.SetValue("uID", aId)

                        trlParent.StateImageIndex = MsgStyle.db
                    Else
                        trlParent = afrm.TreLstConnStatus.FindNodeByFieldValue("uID", aId)

                        If trlParent IsNot Nothing Then

                            trlChild = afrm.TreLstConnStatus.AppendNode(Nothing, trlParent)
                            Select Case aSOStatus
                                Case SOStatus.Yes
                                    trlChild.SetValue("SOInfo", Format(Now, "yyyy/MM/dd hh:mm:ss") & _
                                                      " > [ " & aObj.CompName & " ] 資料庫連接成功")
                                    trlChild.StateImageIndex = MsgStyle.OKLa
                                Case SOStatus.NO
                                    trlChild.SetValue("SOInfo", Format(Now, "yyyy/MM/dd hh:mm:ss") & _
                                                      " > [ " & aObj.CompName & " ] 資料庫連接失敗")
                                    trlChild.StateImageIndex = MsgStyle.ErrorLa
                            End Select

                        End If

                    End If

                    

                Finally
                    afrm.TreLstConnStatus.EndUnboundLoad()
                    afrm.TreLstConnStatus.EndUpdate()
                    afrm.TreLstConnStatus.ExpandAll()
                    afrm.TreLstConnStatus.MoveFirst()
                End Try
            End If
        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
        End Try
    End Sub

    '更新系統訊息畫面
    Public Sub UpdateSysMsgUI(ByVal aFrm As Form, _
                              ByVal aTreeList As TreeList)


        Try

            Dim trlParent As TreeListNode = Nothing
            Dim trlChild As TreeListNode = Nothing


            If FLstSysMsg Is Nothing Then
                Exit Sub
            End If
            If aFrm.InvokeRequired Then
                Dim act As New Action(Of Form, TreeList)(AddressOf UpdateSysMsgUI)
                aFrm.Invoke(act, aFrm, aTreeList)
            Else
                Try
                    aTreeList.ClearNodes()
                    aTreeList.BeginUpdate()
                    aTreeList.BeginUnboundLoad()
                    For i As Int32 = 0 To FLstSysMsg.Count - 1
                        If FLstSysMsg.Item(i).ToString.ToUpper.StartsWith("PARENT") Then
                            trlParent = aTreeList.AppendNode(Nothing, i)

                            trlParent.SetValue("SysInfo", FLstSysMsg.Item(i).Replace("Parent", ""))
                            trlParent.StateImageIndex = MsgStyle.Info
                            '記錄事件
                            WriteErrTxtLog.WriteSysEventLog(FLstSysMsg.Item(i).Replace("Parent", ""))
                        Else
                            trlParent = aTreeList.Nodes.LastNode
                            trlChild = aTreeList.AppendNode(Nothing, trlParent)
                            trlChild.SetValue("SysInfo", FLstSysMsg.Item(i))
                            trlChild.StateImageIndex = MsgStyle.InfoLa
                            '記錄事件
                            WriteErrTxtLog.WriteSysEventLog(FLstSysMsg.Item(i))
                        End If
                    Next

                Finally

                    aTreeList.EndUnboundLoad()
                    aTreeList.EndUpdate()
                    aTreeList.ExpandAll()
                    aTreeList.MoveFirst()
                End Try

            End If

        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
            'Throw New CableSoft.Gateway.csException.CablesoftException(ex, True)
        End Try

    End Sub

    '更新系統系統錯誤系統畫面
    Public Sub UpdateSysErrUI(ByVal aFrm As Form, _
                              ByVal aTreeList As TreeList, _
                              ByVal exSource As Exception, _
                              ByVal aMsg As String)

        Try
            If String.IsNullOrEmpty(aMsg) AndAlso (exSource Is Nothing) Then
                Exit Sub
            End If
            If aFrm.InvokeRequired Then
                Dim act As New Action(Of Form, TreeList, Exception, String)(AddressOf UpdateSysErrUI)
                aFrm.Invoke(act, aFrm, aTreeList, exSource, aMsg)
            Else
                Try
                    SyncLock FShowDataCount.GetType
                        aTreeList.BeginUpdate()
                        aTreeList.BeginUnboundLoad()
                        Dim trlParent As TreeListNode = Nothing
                        Dim trlChild As TreeListNode = Nothing
                        If ((aTreeList.Nodes.Count - FShowDataCount) > FClearDataCount) Then
                            aTreeList.Nodes.RemoveAt(aTreeList.Nodes.Count - 1)
                        End If
                        trlParent = aTreeList.AppendNode(Nothing, -1)
                        trlParent.SetValue("ErrMsg", Format(Now, "yyyy/MM/dd hh:mm:ss") & " > " & _
                               aMsg)
                        trlParent.StateImageIndex = MsgStyle.Error
                        If Not String.IsNullOrEmpty(aMsg) Then
                            trlChild = aTreeList.AppendNode(Nothing, trlParent)
                            trlChild.SetValue("ErrMsg", exSource.Message.Trim(Environment.NewLine).ToString().Trim(vbLf))
                            trlChild.StateImageIndex = MsgStyle.ErrorLa
                        End If

                        'aTreeList.MoveFirst()
                    End SyncLock
                Finally
                    aTreeList.EndUnboundLoad()
                    aTreeList.EndUpdate()
                    aTreeList.ExpandAll()
                    aTreeList.MoveFirst()
                End Try
            End If
        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
        
        End Try

    End Sub


    '更新SO Information的畫面
    Public Sub UpdateSOStatusUI(ByVal aFrm As Form, _
                                ByVal aTreeList As TreeList, _
                                ByVal aSO As SOInfoClass, _
                                ByVal aStatus As SOStatus)
        Try
            If aFrm.InvokeRequired Then
                Dim act As New Action(Of Form, TreeList, SOInfoClass, SOStatus)(AddressOf UpdateSOStatusUI)
                aFrm.Invoke(act, aFrm, aTreeList, aSO, aStatus)

            Else
                Try
                    aTreeList.BeginUpdate()
                    aTreeList.BeginUnboundLoad()
                    Select Case aStatus
                        Case SOStatus.Initial, SOStatus.Key
                            aTreeList.ClearNodes()
                            For i As Int32 = 0 To FSOInfoList.Count - 1
                                aTreeList.AppendNode(Nothing, FSOInfoList.Item(i).CompID)
                                aTreeList.Nodes(i).SetValue("CompId", FSOInfoList.Item(i).CompID)
                                aTreeList.Nodes(i).SetValue("CompName", FSOInfoList.Item(i).CompName)
                                aTreeList.Nodes(i).SetValue("WaitProc", "0")
                                aTreeList.Nodes(i).SetValue("Procing", "0")
                                aTreeList.Nodes(i).StateImageIndex = aStatus
                                SyncLock FSOInfoList(i).ProcessingNumber.GetType
                                    FSOInfoList(i).ProcessingNumber = 0
                                End SyncLock
                                SyncLock FSOInfoList(i).WaitProcessNumber.GetType
                                    FSOInfoList(i).WaitProcessNumber = 0
                                End SyncLock

                            Next
                        Case Else
                            Dim aNote As TreeListNode = aTreeList.FindNodeByFieldValue("CompId", aSO.CompID)
                            If aNote IsNot Nothing Then
                                If (FThreadStop) OrElse (Not FRegOK) Then
                                    aNote.SetValue("WaitProc", "0")
                                    aNote.SetValue("Procing", "0")
                                Else
                                    aNote.SetValue("WaitProc", aSO.WaitProcessNumber)
                                    aNote.SetValue("Procing", aSO.ProcessingNumber)
                                End If

                                If (aStatus <> SOStatus.NotUpdateImg) AndAlso (Not FThreadStop) Then
                                    If FRegOK Then
                                        aNote.StateImageIndex = aStatus
                                    End If

                                End If

                            End If

                    End Select
                Finally
                    aTreeList.ExpandAll()
                    aTreeList.EndUnboundLoad()
                    aTreeList.EndUpdate()

                End Try
            End If
        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
        End Try
    End Sub



End Module
