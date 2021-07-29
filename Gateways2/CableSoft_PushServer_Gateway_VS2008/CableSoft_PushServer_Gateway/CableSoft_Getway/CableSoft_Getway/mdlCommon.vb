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
Imports Cablesoft.CAS.CryptUtil
Module mdlCommon
    Public Enum UpdCmdMode
        Add = 0
        Upd = 1
    End Enum
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
    Public Const FMDBPassword As String = "chunya36"
    '-------------------公用Table----------------------------------
    Public FSysTable As DataTable = Nothing     '系統參數Table
    Public FSOTable As DataTable = Nothing      'COM區 Table
    Public FPushServerParaTable As DataTable = Nothing   'Nagra參數 Table
    Public FComparedErrorTable As DataTable = Nothing   '錯誤對照表 Table
    Public FUseCompanyTable As DataTable = Nothing      '使用的公司 Table

    '-------------------------------------------------------------------
    Public FSOInfoList As List(Of SOInfoClass) = Nothing
    Public FUseCompList As List(Of UseCompanyClass) = Nothing


    Public Const SystemTableName As String = "SysOpt"
    Public Const SOTableName As String = "SO"
    Public Const PushServerParaTableName As String = "PushServerParaSet"
    Public Const UseCompanyTableName As String = "UseCompany"
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

    Public FPushUrl As String = String.Empty
    Public FPushRepeat As String = String.Empty
    Public FPushDuration As String = String.Empty
    Public FPushUrlMethod As String = "POST"
    Public FPushUrlTimeout As Int32 = 30

    Public FUseCompWhere As String = String.Empty
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
    Public Class UseCompanyClass
        Private _CompID As Int32
        Private _CompName As String = String.Empty
        Private _WaitProcessNumber As Int32 = 0
        Private _ProcessingNumber As Int32 = 0
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
    End Class
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
        Private _UseCompId As Int32 = 0

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
        Public Property UseCompId() As Int32
            Get
                Return _UseCompId
            End Get
            Set(ByVal value As Int32)
                _UseCompId = value
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
                RegKey.SetValue("Cablesoft PVOD Nagra Gateway", App, RegistryValueKind.String)
            Else
                RegKey.DeleteValue("Cablesoft PVOD Nagra Gateway", False)
            End If
        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
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
                FErrCodeLst.Add(GetFieldValue(FComparedErrorTable.Rows(i), "ErrorCode"), _
                                GetFieldValue(FComparedErrorTable.Rows(i), "ErrorName"))
            Next
            Return True
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function
    '設定Push Server參數值
    Public Function SetPushServerPara() As Boolean
        Try
            FPushServerParaTable = CableSoft.Gateway.SystemIO.GetwaySys.ReadPushServerParaSet(FMDBFile, FMDBPassword)
            If (FPushServerParaTable IsNot Nothing) AndAlso (FPushServerParaTable.Rows.Count > 0) Then
                With FPushServerParaTable
                    FPushDuration = GetFieldValue(.Rows(0), "PushServerDuration")
                    FPushRepeat = GetFieldValue(.Rows(0), "PushServerRepeat")
                    FPushUrl = GetFieldValue(.Rows(0), "PushServerUrl")
                    FPushUrlMethod = GetFieldValue(.Rows(0), "PushServerUrlMethod")
                    FPushUrlTimeout = Convert.ToInt32("0" & GetFieldValue(.Rows(0), "PushServerUrlTimeOut"))
                End With
            Else
                CableSoft.Gateway.SystemIO.GetwaySys.DefaultPushServerPara(FMDBFile, FMDBPassword, PushServerParaTableName)
                SetPushServerPara()
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
                    FUseTray = GetFieldValue(.Rows(0), "UseTray") = "1"
                    FShowResource = Convert.ToInt32("0" & GetFieldValue(.Rows(0), "ShowResource")) = 1
                    FUseHotKey = Convert.ToInt32("0" & GetFieldValue(.Rows(0), "UseHotKey")) = 1
                    FAutoConnectTime = Convert.ToInt32("0" & GetFieldValue(.Rows(0), "AutoConnectTime"))
                    FReadDataTime = Convert.ToInt32("0" & GetFieldValue(.Rows(0), "ReadDataTime"))
                    FProcessNumber = Convert.ToInt32("0" & GetFieldValue(.Rows(0), "ProcessNumber"))
                    FShowDataCount = Convert.ToInt32("0" & GetFieldValue(.Rows(0), "ShowDataCount"))
                    FClearDataCount = Convert.ToInt32("0" & GetFieldValue(.Rows(0), "ClearDataCount"))
                    FMaxThread = Convert.ToInt32("0" & GetFieldValue(.Rows(0), "MaxThread"))

                    FDBPoolMaxNumber = Convert.ToInt32("0" & GetFieldValue(.Rows(0), "DBPoolMaxNumber"))
                    FDBPoolMinNumber = Convert.ToInt32("0" & GetFieldValue(.Rows(0), "DBPoolMinNumber"))
                    FDBPoolLiveTime = Convert.ToInt32("0" & GetFieldValue(.Rows(0), "DBPoolLiveTime"))
                    'FTestNagraSock = GetFieldValue(.Rows(0), "TestSocketTime")


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
            FLstSysMsg.Add(String.Format("顯示系統資源於狀態列 : {0}", IIf(FShowResource, "是", "否")))
            FLstSysMsg.Add(String.Format("每 {0} 秒讀系統台資料", FReadDataTime))
            FLstSysMsg.Add(String.Format("每次處理 {0} 筆高階指令", FProcessNumber))
            FLstSysMsg.Add(String.Format("畫面資料顯示 {0} 筆訊息 ,  每超過 {1} 筆訊息 ,  則開始清除多餘訊息", FShowDataCount, FClearDataCount))
            FLstSysMsg.Add(String.Format("最大命令處理線程 {0} ", FMaxThread))
            FLstSysMsg.Add(String.Format("DB最小集區 : {0} , DB最大集區 : {1} , 集區Lifetime : {2}", FDBPoolMinNumber, FDBPoolMaxNumber, FDBPoolLiveTime))
            'FLstSysMsg.Add("Parent 讀取Nagra參數設定..")
            'FLstSysMsg.Add(String.Format("傳送指令 Source Id : {0} , Dest Id : {1} , Mop PPId : {2} ", _
            '                             FSndSourceId, FSndDestId, FSndMopPPId))
            'FLstSysMsg.Add(String.Format("回傳指令 Source Id : {0} , Dest Id : {1} , Mop PPId : {2} ", _
            '                             FRcvSourceId, FRcvDestId, FRcvMopPPId))
            'FLstSysMsg.Add(String.Format("傳送指令, 發生 {0} 次錯誤即停止執行", FSndErrStopCnt))
            'FLstSysMsg.Add(String.Format("接收指令, 發生 {0} 次錯誤即停止執行", FRcvErrStopCnt))
            'FLstSysMsg.Add(String.Format("每 {0} 秒傳送一次 Commnucation Check", FCheckStatusTime))
            'FLstSysMsg.Add(String.Format("每傳送一筆低階指令後延遲 {0} 秒", FSndDelayTime))
            'FLstSysMsg.Add(String.Format("連線主機 : {0} , 通訊埠 : {1} ", FNagraServer, FNagraPort))
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub
    '設定使用公司別
    Public Function SetUseComp() As Boolean
        Try
            FUseCompWhere = String.Empty
            FUseCompanyTable = CableSoft.GateWay.SystemIO.GetwaySys.ReadUseComp(FMDBFile, FMDBPassword)
            If FUseCompanyTable IsNot Nothing AndAlso FUseCompanyTable.Rows.Count > 0 Then
                'If FUseCompanyTable.Select("SelectID=1").Length = 0 Then
                '    Throw New Exception("無任何執行SO資料")
                'End If
                If FUseCompList Is Nothing Then
                    FUseCompList = New List(Of UseCompanyClass)
                End If
                FUseCompList.Clear()
                For i As Int32 = 0 To FUseCompanyTable.Rows.Count - 1
                    If Convert.ToUInt32(GetFieldValue(FUseCompanyTable.Rows(i), "SelectID")) = 1 Then
                        Dim obj As New UseCompanyClass
                        obj.CompID = Convert.ToUInt32(GetFieldValue(FUseCompanyTable.Rows(i), "CompID"))
                        obj.CompName = GetFieldValue(FUseCompanyTable.Rows(i), "CompName")
                        FUseCompList.Add(obj)
                        If String.IsNullOrEmpty(FUseCompWhere) Then
                            FUseCompWhere = obj.CompID
                        Else
                            FUseCompWhere = FUseCompWhere & "," & obj.CompID
                        End If
                    End If

                Next
                If FUseCompList Is Nothing OrElse FUseCompList.Count <= 0 Then
                    Throw New Exception("無任何執行SO資料")
                End If
            Else
                Throw New Exception("無任何執行SO資料")
            End If
            Return True
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function
    '設定SO
    Public Function SetSO() As Boolean
        Try


            Dim _ConnectString As String = "Data Source={0};Persist Security Info=True;" & _
                "User ID={1};Password={2};Min Pool Size={3};" & _
                "Max Pool Size={4};Unicode=True;Load Balance Timeout={5}"
            If FSOInfoList IsNot Nothing Then
                FSOInfoList.Clear()
            End If

            FSOTable = CableSoft.GateWay.SystemIO.GetwaySys.ReadSO(FMDBFile, FMDBPassword)

            If (FSOTable IsNot Nothing) AndAlso (FSOTable.Rows.Count > 0) Then
                If FSOInfoList Is Nothing Then
                    FSOInfoList = New List(Of SOInfoClass)
                End If

                FSOInfoList.Clear()
                For i As Int32 = 0 To FSOTable.Rows.Count - 1

                    Dim obj As New SOInfoClass()

                    obj.OraConnectString = String.Format(_ConnectString, _
                         GetFieldValue(FSOTable.Rows(i), "SID"), _
                         GetFieldValue(FSOTable.Rows(i), "UserId"), _
                         GetFieldValue(FSOTable.Rows(i), "UserPassword"), _
                         "0", _
                         FDBPoolMaxNumber.ToString, _
                         FDBPoolLiveTime)
                    obj.ProcessingNumber = 0
                    obj.WaitProcessNumber = 0
                    'obj.CompID = Int32.Parse(FSOTable.Rows(i)("CompID"))
                    'obj.CompName = FSOTable.Rows(i)("CompName")
                    obj.CompID = -99
                    obj.CompName = GetFieldValue(FSOTable.Rows(i), "SID")
                    obj.ConnectionOK = False
                    FSOInfoList.Add(obj)
                Next
            Else
                Throw New Exception("COM區未設定連線資料")
            End If

            'Dim row() As DataRow = FSOTable.Select("SelectID=1")



            'If (FSOTable IsNot Nothing) AndAlso (row.Length > 0) Then
            '    FSOInfoList = New List(Of SOInfoClass)
            '    For i As Int32 = 0 To row.Length - 1
            '        Dim obj As New SOInfoClass()
            '        obj.OraConnectString = String.Format(_ConnectString, _
            '            row(i)("SID"), row(i)("UserId"), _
            '            row(i)("UserPassword"), "0", FDBPoolMaxNumber.ToString, FDBPoolLiveTime)
            '        obj.ProcessingNumber = 0
            '        obj.WaitProcessNumber = 0
            '        obj.CompID = Int32.Parse(row(i)("CompID"))
            '        obj.CompName = row(i)("CompName")
            '        obj.ConnectionOK = False
            '        FSOInfoList.Add(obj)
            '    Next

            'Else
            '    Throw New Exception("無任何SO資料！")
            'End If
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


    Public Function GetFieldValue(ByVal row As DataRow, ByVal FieldName As String, _
                                  Optional ByVal blnDecrypt As Boolean = True) As String
        Try
            If Not DBNull.Value.Equals(row.Item(FieldName)) Then
                If blnDecrypt Then
                    Return CryptUtil.Decrypt(System.Convert.ToString(row.Item(FieldName)))
                Else
                    Return System.Convert.ToString(row.Item(FieldName))
                End If

            Else
                Return String.Empty
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
                                           " > 資料庫連接中..")
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
    Public Sub UpdateNagraMsgUI(ByVal aFrm As Form, _
                                ByVal aTreeList As TreeList, ByVal aMsgStyle As MsgStyle)
        Try
            Dim trlParent As TreeListNode = Nothing
            Dim trlChild As TreeListNode = Nothing
            If aFrm.InvokeRequired Then
                Dim act As New Action(Of Form, TreeList, MsgBoxStyle)(AddressOf UpdateNagraMsgUI)
                aFrm.Invoke(act, aFrm, aTreeList, aMsgStyle)
            Else
                Try


                    aTreeList.BeginUpdate()
                    aTreeList.BeginUnboundLoad()
                    If ((aTreeList.Nodes.Count - FShowDataCount) > FClearDataCount) Then
                        aTreeList.Nodes.RemoveAt(aTreeList.Nodes.Count - 1)
                    End If
                    For i As Int32 = 0 To FLstSysMsg.Count - 1
                        If FLstSysMsg.Item(i).ToString.ToUpper.StartsWith("PARENT") Then
                            trlParent = aTreeList.AppendNode(Nothing, -1)

                            trlParent.SetValue("SysInfo", FLstSysMsg.Item(i).Replace("Parent", ""))
                            trlParent.StateImageIndex = MsgStyle.Info
                            '記錄事件
                            WriteErrTxtLog.WriteSysEventLog(FLstSysMsg.Item(i).Replace("Parent", ""))
                        Else
                            trlParent = aTreeList.Nodes.LastNode
                            trlChild = aTreeList.AppendNode(Nothing, trlParent)
                            trlChild.SetValue("SysInfo", FLstSysMsg.Item(i))
                            trlChild.StateImageIndex = aMsgStyle
                            '記錄事件
                            WriteErrTxtLog.WriteSysEventLog(FLstSysMsg.Item(i))
                        End If
                    Next
                Finally
                    aTreeList.EndUnboundLoad()
                    aTreeList.EndUpdate()
                    aTreeList.ExpandAll()
                    aTreeList.MoveLast()
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
    '更新低階命令視窗
    Public Sub UpdateLowCmdUI(ByVal aFrm As rfrmMain, _
                              ByVal aCMDInfo As Dictionary(Of String, String))

        Try
            If aFrm.InvokeRequired Then
                Dim act As New Action(Of rfrmMain, Dictionary(Of String, String))(AddressOf UpdateLowCmdUI)
                aFrm.Invoke(act, aFrm, aCMDInfo)
            Else
                Try
                    aFrm.TreLstLowCmd.BeginUpdate()
                    aFrm.TreLstLowCmd.BeginUnboundLoad()
                    If ((aFrm.TreLstLowCmd.Nodes.Count - FShowDataCount) > FClearDataCount) Then
                        aFrm.TreLstLowCmd.Nodes.RemoveAt(0)
                    End If
                    Dim trlParent As TreeListNode = aFrm.TreLstLowCmd.AppendNode(Nothing, -1)
                    For i As Int32 = 0 To aFrm.TreLstLowCmd.Columns.Count - 1
                        Dim aFieldName As String = aFrm.TreLstLowCmd.Columns(i).FieldName
                        trlParent.SetValue(aFieldName, aCMDInfo.Item(aFieldName.ToUpper))
                    Next
                    Select Case aCMDInfo.Item("PushServerCode".ToUpper)
                        Case "200"
                            trlParent.StateImageIndex = MsgStyle.OKLa
                        Case Else
                            trlParent.StateImageIndex = MsgStyle.ErrorLa
                    End Select
                Finally

                    aFrm.TreLstLowCmd.EndUnboundLoad()
                    aFrm.TreLstLowCmd.EndUpdate()
                    aFrm.TreLstLowCmd.MoveLast()
                End Try


            End If
        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
        End Try

    End Sub
    '更新高階命令視窗
    Public Sub UpdateHightCmdUI(ByVal aFrm As rfrmMain, _
                        ByVal aCmdInfo As Dictionary(Of String, String), _
                        ByVal aType As UpdCmdMode)
        Try
            If aFrm.InvokeRequired Then
                Dim act As New Action(Of rfrmMain, Dictionary(Of String, String), UpdCmdMode)(AddressOf UpdateHightCmdUI)
                aFrm.Invoke(act, aFrm, aCmdInfo, aType)
            Else
                Try
                    Dim trlParent As TreeListNode = Nothing
                    aFrm.TreLstInstruct.BeginUpdate()
                    aFrm.TreLstInstruct.BeginUnboundLoad()
                    If aType = UpdCmdMode.Add Then
                        aCmdInfo.Item("GUIDNO") = Guid.NewGuid.ToString
                        If ((aFrm.TreLstInstruct.Nodes.Count - FShowDataCount) > FClearDataCount) Then
                            aFrm.TreLstInstruct.Nodes.RemoveAt(0)
                        End If
                        trlParent = aFrm.TreLstInstruct.AppendNode(Nothing, -1)
                    End If
                    If aType = UpdCmdMode.Upd Then
                        trlParent = aFrm.TreLstInstruct.FindNodeByFieldValue("GuidNo", aCmdInfo.Item("GUIDNO"))
                    End If
                    If trlParent IsNot Nothing Then
                        For i As Int32 = 0 To aFrm.TreLstInstruct.Columns.Count - 1
                            Dim aFieldName As String = aFrm.TreLstInstruct.Columns(i).FieldName
                            trlParent.SetValue(aFieldName, aCmdInfo.Item(aFieldName.ToUpper))
                        Next
                        Select Case aCmdInfo.Item("CMDSTATUS")
                            Case "S"
                                trlParent.StateImageIndex = MsgStyle.OKLa
                            Case "P2", "W"
                                trlParent.StateImageIndex = MsgStyle.Process
                            Case Else
                                trlParent.StateImageIndex = MsgStyle.ErrorLa
                        End Select
                    End If

                Finally
                    'aFrm.TreLstInstruct.ExpandAll()
                    aFrm.TreLstInstruct.EndUnboundLoad()
                    aFrm.TreLstInstruct.EndUpdate()
                    aFrm.TreLstInstruct.MoveLast()
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
                            If (FSOInfoList.Count = 1) AndAlso (FSOInfoList.Item(0).CompID = -99) Then

                                For i As Int32 = 0 To FUseCompList.Count - 1
                                    aTreeList.AppendNode(Nothing, FUseCompList.Item(i).CompID)
                                    aTreeList.Nodes(i).SetValue("CompId", FUseCompList.Item(i).CompID)
                                    aTreeList.Nodes(i).SetValue("CompName", FUseCompList.Item(i).CompName)
                                    aTreeList.Nodes(i).SetValue("WaitProc", "0")
                                    aTreeList.Nodes(i).SetValue("Procing", "0")
                                    aTreeList.Nodes(i).StateImageIndex = aStatus
                                    SyncLock FUseCompList(i).ProcessingNumber.GetType
                                        FUseCompList(i).ProcessingNumber = 0
                                    End SyncLock
                                    SyncLock FUseCompList(i).WaitProcessNumber.GetType
                                        FUseCompList(i).WaitProcessNumber = 0
                                    End SyncLock
                                Next

                            Else

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
                            End If

                        Case Else

                            If aSO.CompID = -99 Then
                                If aStatus <> SOStatus.NotUpdateImg Then

                                    For i As Int32 = 0 To aTreeList.Nodes.Count - 1
                                        If (FThreadStop) OrElse (Not FRegOK) Then
                                            aTreeList.Nodes(i).SetValue("WaitProc", "0")
                                            aTreeList.Nodes(i).SetValue("Procing", "0")

                                        End If
                                        If (aStatus <> SOStatus.NotUpdateImg) AndAlso (Not FThreadStop) Then
                                            If FRegOK Then
                                                aTreeList.Nodes(i).StateImageIndex = aStatus
                                            End If

                                        End If
                                    Next
                                Else
                                    Dim aNote As TreeListNode = aTreeList.FindNodeByFieldValue("CompId", aSO.UseCompId)

                                    If aNote IsNot Nothing Then
                                        If (FThreadStop) OrElse (Not FRegOK) OrElse (FThreadsTotal <= 0) Then
                                            aNote.SetValue("WaitProc", "0")
                                            aNote.SetValue("Procing", "0")
                                        Else
                                            aNote.SetValue("WaitProc", FUseCompIndex.Item(aSO.UseCompId).WaitProcessNumber)
                                            aNote.SetValue("Procing", FUseCompIndex.Item(aSO.UseCompId).ProcessingNumber)
                                        End If
                                    End If

                                End If
                            Else

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
