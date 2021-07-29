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
    Public FNagraParaTable As DataTable = Nothing   'Nagra參數 Table
    Public FComparedErrorTable As DataTable = Nothing   '錯誤對照表 Table
    Public FUseCompanyTable As DataTable = Nothing      '使用的公司 Table
    Public fMailInfoTable As DataTable = Nothing
    Public fMailToTable As DataTable = Nothing
    Public fMailCCTable As DataTable = Nothing
    Public fMailPRMParaTable As DataTable = Nothing
    '-------------------------------------------------------------------
    Public FSOInfoList As List(Of SOInfoClass) = Nothing
    Public FUseCompList As List(Of UseCompanyClass) = Nothing


    Public Const SystemTableName As String = "SysOpt"
    Public Const SOTableName As String = "SO"
    Public Const NagraParaName As String = "NagraParaSet"
    Public Const UseCompanyTableName As String = "UseCompany"
    Public Const ComparedErrorName As String = "ComparedError"
    Public Const fSourceIdTableName As String = "SocketSourceId"
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
    Public FSndSourceId() As String
    Public FSndDestId() As String
    Public FSndMopPPId() As String
    Public FRcvSourceId As String = String.Empty
    Public FRcvDestId As String = String.Empty
    Public FRcvMopPPId As String = String.Empty
    Public FNagraMaxProcing As Int32 = 1000
    Public FNagraHaveProcing As Int32 = 0
    Public FNagraUpdEn As String = String.Empty
    Public fResumeCnt As String = "1"
    Public fErrSndMail As Boolean = False
    Public FSndErrStopCnt As Int32 = 3
    Public FRcvErrStopCnt As Int32 = 3
    Public FCheckStatusTime As Int32 = 3
    Public FSndDelayTime As Single = 0.1
    Public FSndTimeout As Int32 = 0
    Public FRcvTimeout As Int32 = 0
    Public FUseCompWhere As String = String.Empty
    Public FErrCodeLst As New Dictionary(Of String, String)
    Public fErrResumeLst As New List(Of String)
    '-------------------------------------------------------------------
    Public FLstSysMsg As List(Of String)
    Public FInsDate As String = String.Empty
    Public FRegDate As String = String.Empty
    Public FRegOK As Boolean = False

    '---------------------SourceID有問題發送Mail的資訊--------------------------------
    Public Const fMDBMailFileName As String = "WarningEMail.Set"
    Public Const fMDBMailPassword As String = Nothing
    Public Const fMailInfoTableName As String = "CsMailInfo"
    Public Const fMailToTableName As String = "CsMailTo"
    Public Const fMailCCTableName As String = "CsMailCC"
    Public Const fMailPRMParaTableName As String = "CsPRMPara"
    Public fSmtpHost As String = String.Empty
    Public fSmtplUserId As String = String.Empty
    Public fSmtpPassword As String = String.Empty
    Public fEnabledSSL As Boolean = False
    Public fMailDisplayName As String = String.Empty
    Public fMailSender As String = String.Empty
    Public fSmtpPort As String = "25"
    Public fMailSubject As String = String.Empty
    Public fMailBody As String = String.Empty
    Public fSourceIdErrCnt As String = "3"
    Public fIsUseMail As Boolean = False
    '-------------------------------------------------------------------------------------


    '每跑幾個高階命令畫面更新就做DoEvent
    Public fDoEventCount As Int32
    Public Const fDoEventNum As Int32 = 30
    Public fFrom As rfrmMain = Nothing
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
        Implements IDisposable
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



#Region "IDisposable Support"
        Private disposedValue As Boolean ' 偵測多餘的呼叫

        ' IDisposable
        Protected Overridable Sub Dispose(disposing As Boolean)
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
            FComparedErrorTable = CableSoft.Gateway.SystemIO.GetwaySys.ReadErrorCode(FMDBFile, FMDBPassword)
            FErrCodeLst.Clear()
            fErrResumeLst.Clear()
            For i As Int32 = 0 To FComparedErrorTable.Rows.Count - 1
                FErrCodeLst.Add(GetFieldValue(FComparedErrorTable.Rows(i), "ErrorCode"), _
                                GetFieldValue(FComparedErrorTable.Rows(i), "ErrorName"))
                If GetFieldValue(FComparedErrorTable.Rows(i), "Choice") = "1" Then
                    fErrResumeLst.Add(GetFieldValue(FComparedErrorTable.Rows(i), "ErrorCode"))
                End If
            Next
            Return True
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function
    '讀取Mail設值
    Public Function ReadMailPara() As Boolean
        Try
            Dim aFile As String = MDBMailPara.GetDefaultMDBName
            fIsUseMail = MDBMailPara.IsUseSndMail(fMDBMailFileName)

            If Not fIsUseMail Then
                WriteErrTxtLog.WriteSndEmailLog("找不到Email設定檔！")
                Return True
            End If

            fMailInfoTable = MDBCommon.ReadMDBDataTable(aFile, fMDBMailPassword, fMailInfoTableName, Nothing)
            fMailToTable = MDBCommon.ReadMDBDataTable(aFile, fMDBMailPassword, fMailToTableName, Nothing)
            fMailCCTable = MDBCommon.ReadMDBDataTable(aFile, fMDBMailPassword, fMailCCTableName, Nothing)
            fMailPRMParaTable = MDBCommon.ReadMDBDataTable(aFile, fMDBMailPassword, fMailPRMParaTableName, Nothing)

            If fMailInfoTable Is Nothing OrElse fMailInfoTable.Rows.Count <= 0 Then
                MDBMailPara.DefaultMailInfo(aFile, fMDBMailPassword, fMailInfoTableName)
                ReadMailPara()
            End If
            If fMailPRMParaTable Is Nothing OrElse fMailPRMParaTable.Rows.Count <= 0 Then
                MDBMailPara.DefaultPRMPara(aFile, fMDBMailPassword, fMailPRMParaTableName)
                ReadMailPara()
            End If

            With fMailInfoTable
                fSmtpHost = GetFieldValue(.Rows(0), "SmtpHost")
                fSmtplUserId = GetFieldValue(.Rows(0), "SmtpUserId")
                fSmtpPassword = GetFieldValue(.Rows(0), "SmtpPassword")
                fSmtpPort = GetFieldValue(.Rows(0), "SmtpPort")
                fEnabledSSL = GetFieldValue(.Rows(0), "EnabledSSL") = "1"
                fMailDisplayName = GetFieldValue(.Rows(0), "MailDisplayNam")
                fMailSender = GetFieldValue(.Rows(0), "MailSender")
                fMailSubject = GetFieldValue(.Rows(0), "MailSubject")
                fMailBody = GetFieldValue(.Rows(0), "MailBody")
            End With
            With fMailPRMParaTable
                fSourceIdErrCnt = GetFieldValue(.Rows(0), "SourceIdErrCnt")
            End With


           
        Catch ex As Exception
            'Throw New Exception(ex.Message)
            WriteErrTxtLog.WriteSndEmailLog(ex.Message)
        End Try
    End Function


    '設定Nagra參數值
    Public Function ReadNagraPara() As Boolean
        Try
            FNagraParaTable = CableSoft.Gateway.SystemIO.GetwaySys.ReadNagraParaSet(FMDBFile, FMDBPassword)



            If (FNagraParaTable IsNot Nothing) AndAlso (FNagraParaTable.Rows.Count > 0) Then
                With FNagraParaTable
                    FNagraServer = GetFieldValue(.Rows(0), "NagraIPAddress")
                    FNagraPort = Convert.ToInt32("0" & GetFieldValue(.Rows(0), "NagraPort"))
                    'FSndSourceId(0) = GetFieldValue(.Rows(0), "SndSourceId")
                    'FSndDestId(0) = GetFieldValue(.Rows(0), "SndDestId")
                    'FSndMopPPId(0) = GetFieldValue(.Rows(0), "SndMopPPId")
                    FRcvSourceId = GetFieldValue(.Rows(0), "RcvSourceId")
                    FRcvDestId = GetFieldValue(.Rows(0), "RcvDestId")
                    FRcvMopPPId = GetFieldValue(.Rows(0), "RcvMopPPId")
                    FSndErrStopCnt = Convert.ToInt32("0" & GetFieldValue(.Rows(0), "SenErrStopCnt"))
                    FRcvErrStopCnt = Convert.ToInt32("0" & GetFieldValue(.Rows(0), "RcvErrStopCnt"))
                    FCheckStatusTime = Convert.ToInt32("0" & GetFieldValue(.Rows(0), "CheckStatusTime"))
                    FSndDelayTime = Convert.ToSingle(GetFieldValue(.Rows(0), "SndDelayTime"))
                    FSndTimeout = Convert.ToInt32("0" & GetFieldValue(.Rows(0), "SndTimeout"))
                    FRcvTimeout = Convert.ToInt32("0" & GetFieldValue(.Rows(0), "RcvTimeout"))
                    FNagraMaxProcing = Convert.ToInt32("0" & GetFieldValue(.Rows(0), "MaxProcing"))
                    fResumeCnt = Convert.ToInt32("0" & GetFieldValue(.Rows(0), "ResumeCmdCnt")).ToString
                    fErrSndMail = GetFieldValue(.Rows(0), "CmdErrSndMail") = "1"
                    FNagraUpdEn = GetFieldValue(.Rows(0), "UPDEN")
                    Try
                        FNagraIPAddress = IPAddress.Parse(FNagraServer)
                    Catch ex As Exception
                        FNagraIPAddress = IPAddress.Parse("127.0.0.1")
                    End Try

                End With
            Else
                CableSoft.Gateway.SystemIO.GetwaySys.DefaultNagra(FMDBFile, FMDBPassword, NagraParaName)
                ReadNagraPara()
            End If
            '讀取Nagra 所有的SourceId
            Dim tblSourceId As DataTable = CableSoft.Gateway.SystemIO.GetwaySys.ReadSocketSourceId(FMDBFile, FMDBPassword)
            If tblSourceId.Rows.Count = 0 Then
                Throw New Exception("Nagra SourceId 未設定 !")
                Return False
            End If
            ReDim FSndSourceId(tblSourceId.Rows.Count - 1)
            ReDim FSndDestId(tblSourceId.Rows.Count - 1)
            ReDim FSndMopPPId(tblSourceId.Rows.Count - 1)
            For i As Int32 = 0 To tblSourceId.Rows.Count - 1
                FSndSourceId(i) = GetFieldValue(tblSourceId.Rows(i), "SourceId")
                FSndDestId(i) = GetFieldValue(tblSourceId.Rows(i), "DestId")
                FSndMopPPId(i) = GetFieldValue(tblSourceId.Rows(i), "MoppId")
                If String.IsNullOrEmpty(FSndSourceId(i)) Then
                    Throw New Exception("SourceId 設定有誤!")
                    Return False
                End If

                If String.IsNullOrEmpty(FSndDestId(i)) Then
                    Throw New Exception("Dest Id 設定有誤!")
                    Return False
                End If
                If String.IsNullOrEmpty(FSndMopPPId(i)) Then
                    Throw New Exception("Mopp Id 設定有誤!")
                    Return False
                End If

            Next

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
                    'SetSysMsgString()

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
    Public Sub SetSysMsgString()
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
            FLstSysMsg.Add("Parent 讀取 PRM 參數設定..")
            'FLstSysMsg.Add(String.Format("傳送指令 Source Id : {0} , Dest Id : {1} , Mop PPId : {2} ", _
            '                             FSndSourceId, FSndDestId, FSndMopPPId))
            'FLstSysMsg.Add(String.Format("回傳指令 Source Id : {0} , Dest Id : {1} , Mop PPId : {2} ", _
            '                             FRcvSourceId, FRcvDestId, FRcvMopPPId))
            FLstSysMsg.Add(String.Format("傳送指令, 發生 {0} 次錯誤即停止執行", FSndErrStopCnt))
            FLstSysMsg.Add(String.Format("接收指令, 發生 {0} 次錯誤即停止執行", FRcvErrStopCnt))
            FLstSysMsg.Add(String.Format("每 {0} 秒傳送一次 Commnucation Check", FCheckStatusTime))
            'FLstSysMsg.Add(String.Format("每傳送一筆低階指令後延遲 {0} 秒", FSndDelayTime))
            FLstSysMsg.Add(String.Format("連線主機 : {0} , 通訊埠 : {1} ", FNagraServer, FNagraPort))
            FLstSysMsg.Add(String.Format("送出命令數量達 {0} 筆, 暫時停止傳送", FNagraMaxProcing))
            FLstSysMsg.Add(String.Format("錯誤命令還原到達 {0} 次,不再還原", fResumeCnt))
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
                    '
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
    Public Sub DisableControl(ByVal aStatus As RunStatus)
        On Error Resume Next
        If fFrom.InvokeRequired Then
            Dim act As New Action(Of RunStatus)(AddressOf DisableControl)
            fFrom.Invoke(act, aStatus)
        Else
            Select Case aStatus
                Case RunStatus.Initial
                    fFrom.bbiExec.Enabled = False
                    fFrom.bbiOpt.Enabled = False
                    fFrom.bbiStop.Enabled = False
                    fFrom.bsiView.Enabled = False
                    fFrom.bbiAbout.Enabled = False
                    fFrom.bbiQuit.Enabled = False
                    fFrom.biSets.Enabled = False
                Case RunStatus.Execute
                    fFrom.bbiExec.Enabled = False
                    fFrom.bbiStop.Enabled = True
                    fFrom.bbiOpt.Enabled = False
                    fFrom.bsiView.Enabled = True
                    fFrom.bbiAbout.Enabled = False
                    fFrom.bbiQuit.Enabled = False
                    fFrom.biSets.Enabled = False
                Case RunStatus.StopCmd, RunStatus.LoadComplete
                    fFrom.bbiExec.Enabled = True
                    fFrom.bbiStop.Enabled = False
                    fFrom.bbiOpt.Enabled = True
                    fFrom.bsiView.Enabled = True
                    fFrom.bbiAbout.Enabled = True
                    fFrom.bbiQuit.Enabled = True
                    fFrom.biSets.Enabled = True
            End Select
            fFrom.mnuExec.Enabled = fFrom.bbiExec.Enabled
            fFrom.mnuStop.Enabled = fFrom.bbiStop.Enabled
            fFrom.mnuOpt.Enabled = fFrom.bbiOpt.Enabled
            fFrom.mnuAbout.Enabled = fFrom.bbiAbout.Enabled
            fFrom.mnuQuit.Enabled = fFrom.bbiQuit.Enabled
            fFrom.mnuAbout.Enabled = fFrom.bbiAbout.Enabled
        End If
    End Sub
    '測試是否可連線
    Public Function IsConnectOK(ByRef obj As OracleConnection, Optional ByRef aErrMsg As Object = Nothing) As Boolean
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
    Public Sub UpdateSODatabase(ByVal aId As String, _
                                ByVal aObj As SOInfoClass, _
                                ByVal aSOStatus As SOStatus)

        Try

            If fFrom.InvokeRequired Then

                Dim act As New Action(Of String, SOInfoClass, SOStatus)(AddressOf UpdateSODatabase)
                fFrom.Invoke(act, aId, aObj, aSOStatus)

            Else
                Try
                    fFrom.TreLstConnStatus.BeginUpdate()

                    fFrom.TreLstConnStatus.BeginUnboundLoad()

                    Dim trlParent As TreeListNode = Nothing
                    Dim trlChild As TreeListNode = Nothing


                    If ((fFrom.TreLstConnStatus.Nodes.Count - FShowDataCount) > FClearDataCount) Then
                        fFrom.TreLstConnStatus.Nodes.RemoveAt(fFrom.TreLstConnStatus.Nodes.Count - 1)
                    End If
                    If aSOStatus = SOStatus.NotUpdateImg Then
                        trlParent = fFrom.TreLstConnStatus.AppendNode(Nothing, -1)
                        trlParent.SetValue("SOInfo", Format(Now, "yyyy/MM/dd HH:mm:ss") & _
                                           " > 資料庫連接中..")
                        trlParent.SetValue("uID", aId)

                        trlParent.StateImageIndex = MsgStyle.db
                    Else
                        trlParent = fFrom.TreLstConnStatus.FindNodeByFieldValue("uID", aId)

                        If trlParent IsNot Nothing Then

                            trlChild = fFrom.TreLstConnStatus.AppendNode(Nothing, trlParent)
                            Select Case aSOStatus
                                Case SOStatus.Yes
                                    trlChild.SetValue("SOInfo", Format(Now, "yyyy/MM/dd HH:mm:ss") & _
                                                      " > [ " & aObj.CompName & " ] 資料庫連接成功")
                                    trlChild.StateImageIndex = MsgStyle.OKLa
                                Case SOStatus.NO
                                    trlChild.SetValue("SOInfo", Format(Now, "yyyy/MM/dd HH:mm:ss") & _
                                                      " > [ " & aObj.CompName & " ] 資料庫連接失敗")
                                    trlChild.StateImageIndex = MsgStyle.ErrorLa
                            End Select

                        End If

                    End If


                Finally
                    fFrom.TreLstConnStatus.EndUnboundLoad()
                    fFrom.TreLstConnStatus.EndUpdate()
                    fFrom.TreLstConnStatus.ExpandAll()
                    fFrom.TreLstConnStatus.MoveFirst()

                End Try
            End If
        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
        Finally
            aObj.Dispose()
        End Try
    End Sub
    Public Sub UpdateNagraMsgUI( ByVal aMsgStyle As MsgStyle)
        Try
            Dim trlParent As TreeListNode = Nothing
            Dim trlChild As TreeListNode = Nothing
            If fFrom.InvokeRequired Then
                Dim act As New Action(Of MsgBoxStyle)(AddressOf UpdateNagraMsgUI)
                fFrom.Invoke(act, aMsgStyle)

            Else
                With fFrom
                    Try

                        .TreLstSysMsg.BeginUpdate()
                        .TreLstSysMsg.BeginUnboundLoad()
                        If ((.TreLstSysMsg.Nodes.Count - FShowDataCount) > FClearDataCount) Then
                            .TreLstSysMsg.Nodes.RemoveAt(.TreLstSysMsg.Nodes.Count - 1)
                        End If
                        For i As Int32 = 0 To FLstSysMsg.Count - 1
                            If FLstSysMsg.Item(i).ToString.ToUpper.StartsWith("PARENT") Then
                                trlParent = .TreLstSysMsg.AppendNode(Nothing, -1)

                                trlParent.SetValue("SysInfo", FLstSysMsg.Item(i).Replace("Parent", ""))
                                trlParent.StateImageIndex = MsgStyle.Info
                                '記錄事件
                                WriteErrTxtLog.WriteSysEventLog(FLstSysMsg.Item(i).Replace("Parent", ""))
                            Else
                                trlParent = .TreLstSysMsg.Nodes.LastNode
                                trlChild = .TreLstSysMsg.AppendNode(Nothing, trlParent)
                                trlChild.SetValue("SysInfo", FLstSysMsg.Item(i))
                                trlChild.StateImageIndex = aMsgStyle
                                '記錄事件
                                WriteErrTxtLog.WriteSysEventLog(FLstSysMsg.Item(i))
                            End If
                        Next

                    Finally
                        .TreLstSysMsg.EndUnboundLoad()
                        .TreLstSysMsg.EndUpdate()
                        .TreLstSysMsg.ExpandAll()
                        .TreLstSysMsg.MoveLast()

                    End Try
                End With
            End If
        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
        End Try
    End Sub

    '更新系統訊息畫面
    Public Sub UpdateSysMsgUI()


        Try

            Dim trlParent As TreeListNode = Nothing
            Dim trlChild As TreeListNode = Nothing


            If FLstSysMsg Is Nothing Then
                Exit Sub
            End If
            If fFrom.InvokeRequired Then
                Dim act As New Action(AddressOf UpdateSysMsgUI)
                fFrom.Invoke(act)
            Else
                Try

                    fFrom.TreLstSysMsg.ClearNodes()
                    fFrom.TreLstSysMsg.BeginUpdate()
                    fFrom.TreLstSysMsg.BeginUnboundLoad()
                    For i As Int32 = 0 To FLstSysMsg.Count - 1
                        If FLstSysMsg.Item(i).ToString.ToUpper.StartsWith("PARENT") Then
                            trlParent = fFrom.TreLstSysMsg.AppendNode(Nothing, i)

                            trlParent.SetValue("SysInfo", FLstSysMsg.Item(i).Replace("Parent", ""))
                            trlParent.StateImageIndex = MsgStyle.Info
                            '記錄事件
                            WriteErrTxtLog.WriteSysEventLog(FLstSysMsg.Item(i).Replace("Parent", ""))
                        Else
                            trlParent = fFrom.TreLstSysMsg.Nodes.LastNode
                            trlChild = fFrom.TreLstSysMsg.AppendNode(Nothing, trlParent)
                            trlChild.SetValue("SysInfo", FLstSysMsg.Item(i))
                            trlChild.StateImageIndex = MsgStyle.InfoLa
                            '記錄事件
                            WriteErrTxtLog.WriteSysEventLog(FLstSysMsg.Item(i))
                        End If
                    Next

                Finally

                    fFrom.TreLstSysMsg.EndUnboundLoad()
                    fFrom.TreLstSysMsg.EndUpdate()
                    fFrom.TreLstSysMsg.ExpandAll()
                    fFrom.TreLstSysMsg.MoveFirst()

                End Try

            End If

        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
            'Throw New CableSoft.Gateway.csException.CablesoftException(ex, True)
        End Try

    End Sub

    '更新系統系統錯誤系統畫面
    Public Sub UpdateSysErrUI(ByVal exSource As Exception, _
                              ByVal aMsg As String)

        Try
            If String.IsNullOrEmpty(aMsg) AndAlso (exSource Is Nothing) Then
                Exit Sub
            End If
            If fFrom.InvokeRequired Then
                Dim act As New Action(Of Exception, String)(AddressOf UpdateSysErrUI)
                fFrom.Invoke(act, exSource, aMsg)
            Else
                With fFrom
                    Try
                        SyncLock FShowDataCount.GetType
                            .TreLstSysErr.BeginUpdate()
                            .TreLstSysErr.BeginUnboundLoad()
                            Dim trlParent As TreeListNode = Nothing
                            Dim trlChild As TreeListNode = Nothing
                            If ((.TreLstSysErr.Nodes.Count - FShowDataCount) > FClearDataCount) Then
                                .TreLstSysErr.Nodes.RemoveAt(.TreLstSysErr.Nodes.Count - 1)
                            End If
                            trlParent = .TreLstSysErr.AppendNode(Nothing, -1)
                            trlParent.SetValue("ErrMsg", Format(Now, "yyyy/MM/dd HH:mm:ss") & " > " & _
                                   aMsg)
                            trlParent.StateImageIndex = MsgStyle.Error
                            If Not String.IsNullOrEmpty(aMsg) Then
                                trlChild = .TreLstSysErr.AppendNode(Nothing, trlParent)
                                trlChild.SetValue("ErrMsg", exSource.Message.Trim(Environment.NewLine).ToString().Trim(vbLf))
                                trlChild.StateImageIndex = MsgStyle.ErrorLa
                            End If

                            'aTreeList.MoveFirst()
                        End SyncLock
                    Finally
                        .TreLstSysErr.EndUnboundLoad()
                        .TreLstSysErr.EndUpdate()
                        .TreLstSysErr.ExpandAll()
                        .TreLstSysErr.MoveFirst()
                    End Try
                End With
            End If
        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)

        End Try

    End Sub
    '更新低階命令視窗
    Public Sub UpdateLowCmdUI(ByVal aCMDInfo As SOCMDInfo)

        Try
            If fFrom.InvokeRequired Then


                Dim act As New Action(Of SOCMDInfo)(AddressOf UpdateLowCmdUI)

                'aFrm.Invoke(act, aFrm, aCMDInfo)
                fFrom.Invoke(act, aCMDInfo)

            Else

                Try

                    fFrom.TreLstLowCmd.BeginUpdate()
                    fFrom.TreLstLowCmd.BeginUnboundLoad()
                    For aIndex As Int32 = 0 To aCMDInfo.NagraSendAndRecvMsg.Count - 1
                        If ((fFrom.TreLstLowCmd.Nodes.Count - FShowDataCount) > FClearDataCount) Then
                            fFrom.TreLstLowCmd.Nodes.RemoveAt(0)
                        End If
                        Dim trlParent As TreeListNode = fFrom.TreLstLowCmd.AppendNode(Nothing, -1)
                        'For i As Int32 = 0 To aFrm.TreLstLowCmd.Columns.Count - 1
                        '    Dim aFieldName As String = aFrm.TreLstLowCmd.Columns(i).FieldName
                        '    trlParent.SetValue(aFieldName, aCMDInfo.AllFieldsValue.Item(aFieldName.ToUpper))
                        'Next
                        trlParent.SetValue("SeqNo", aCMDInfo.AllFieldsValue.Item("SeqNo".ToUpper))
                        trlParent.SetValue("SOCKETSNDMSG", aCMDInfo.NagraSendAndRecvMsg.Keys(aIndex))
                        trlParent.SetValue("SOCKETRCVMSG", aCMDInfo.NagraSendAndRecvMsg.Values(aIndex))
                        Select Case aCMDInfo.AllFieldsValue.Item("CMDSTATUS")
                            Case "S1"
                                trlParent.StateImageIndex = MsgStyle.OKLa
                            Case Else
                                trlParent.StateImageIndex = MsgStyle.ErrorLa
                        End Select
                    Next
                Finally
                    'aFrm.TreLstLowCmd.ExpandAll()
                    fFrom.TreLstLowCmd.EndUnboundLoad()
                    fFrom.TreLstLowCmd.EndUpdate()
                    fFrom.TreLstLowCmd.MoveLast()

                End Try


            End If
        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
        Finally
            aCMDInfo.Dispose()
        End Try

    End Sub
    '更新高階命令視窗
    Public Sub UpdateHightCmdUI(ByVal aCmdInfo As SOCMDInfo, _
                        ByVal aType As UpdCmdMode)


        Try
            If fFrom.InvokeRequired Then
                Dim act As New Action(Of SOCMDInfo, UpdCmdMode)(AddressOf UpdateHightCmdUI)
                'aFrm.BeginInvoke(act, aFrm, aCmdInfo, aType)
                fFrom.Invoke(act, aCmdInfo, aType)

            Else

                'Application.DoEvents()


                Threading.Interlocked.Increment(fDoEventCount)
                If fDoEventCount >= fDoEventNum Then
                    Application.DoEvents()
                    Threading.Interlocked.Exchange(fDoEventCount, 1)
                End If


                Try

                    Dim trlParent As TreeListNode = Nothing
                    fFrom.TreLstInstruct.BeginUpdate()
                    fFrom.TreLstInstruct.BeginUnboundLoad()
                    If aType = UpdCmdMode.Add Then
                        aCmdInfo.AllFieldsValue.Item("GUIDNO") = Guid.NewGuid.ToString
                        If ((fFrom.TreLstInstruct.Nodes.Count - FShowDataCount) > FClearDataCount) Then
                            fFrom.TreLstInstruct.Nodes.RemoveAt(0)
                        End If
                        trlParent = fFrom.TreLstInstruct.AppendNode(Nothing, -1)
                    End If
                    If aType = UpdCmdMode.Upd Then
                        trlParent = fFrom.TreLstInstruct.FindNodeByFieldValue("GuidNo", aCmdInfo.AllFieldsValue.Item("GUIDNO"))
                    End If
                    If trlParent IsNot Nothing Then
                        For i As Int32 = 0 To fFrom.TreLstInstruct.Columns.Count - 1
                            Dim aFieldName As String = fFrom.TreLstInstruct.Columns(i).FieldName
                            'VALIDITY_FLAG
                            If aFieldName.ToUpper = "VALIDITY_FLAG" Then
                                If aCmdInfo.AllFieldsValue("VALIDITY_FLAG") = "1" Then
                                    'absolute=0,relevant=1
                                    trlParent.SetValue(aFieldName, "Relevant")
                                Else
                                    trlParent.SetValue(aFieldName, "Absolute")

                                End If
                            Else
                                trlParent.SetValue(aFieldName, aCmdInfo.AllFieldsValue.Item(aFieldName.ToUpper))
                            End If

                        Next
                        Select Case aCmdInfo.AllFieldsValue.Item("CMDSTATUS")
                            Case "S1", "S"
                                trlParent.StateImageIndex = MsgStyle.OKLa
                            Case "P1", "W"
                                trlParent.StateImageIndex = MsgStyle.Process
                            Case Else
                                trlParent.StateImageIndex = MsgStyle.ErrorLa
                        End Select
                    End If

                Finally
                    ' aFrm.TreLstInstruct.ExpandAll()

                    fFrom.TreLstInstruct.EndUnboundLoad()
                    fFrom.TreLstInstruct.EndUpdate()
                    fFrom.TreLstInstruct.MoveLast()
                    aCmdInfo.Dispose()
                    'aFrm.TreLstInstruct.Refresh()
                    'Application.DoEvents()
                    'aFrm.Refresh()
                End Try
            End If


        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
        End Try

    End Sub


    '更新SO Information的畫面
    Public Sub UpdateSOStatusUI(ByVal aSO As SOInfoClass, _
                                ByVal aStatus As SOStatus)
        Try
            If fFrom.InvokeRequired Then
                Dim act As New Action(Of SOInfoClass, SOStatus)(AddressOf UpdateSOStatusUI)
                'aFrm.Invoke(act, aFrm, aTreeList, aSO, aStatus)
                fFrom.Invoke(act, aSO, aStatus)
            Else
                With fFrom
                    Try
                        .treLstSO.BeginUpdate()
                        .treLstSO.BeginUnboundLoad()
                        Select Case aStatus
                            Case SOStatus.Initial, SOStatus.Key
                                .treLstSO.ClearNodes()
                                If (FSOInfoList.Count = 1) AndAlso (FSOInfoList.Item(0).CompID = -99) Then

                                    For i As Int32 = 0 To FUseCompList.Count - 1
                                        .treLstSO.AppendNode(Nothing, FUseCompList.Item(i).CompID)
                                        .treLstSO.Nodes(i).SetValue("CompId", FUseCompList.Item(i).CompID)
                                        .treLstSO.Nodes(i).SetValue("CompName", FUseCompList.Item(i).CompName)
                                        .treLstSO.Nodes(i).SetValue("WaitProc", "0")
                                        .treLstSO.Nodes(i).SetValue("Procing", "0")
                                        .treLstSO.Nodes(i).StateImageIndex = aStatus
                                        SyncLock FUseCompList(i).ProcessingNumber.GetType
                                            FUseCompList(i).ProcessingNumber = 0
                                        End SyncLock
                                        SyncLock FUseCompList(i).WaitProcessNumber.GetType
                                            FUseCompList(i).WaitProcessNumber = 0
                                        End SyncLock
                                    Next

                                Else

                                    For i As Int32 = 0 To FSOInfoList.Count - 1
                                        .treLstSO.AppendNode(Nothing, FSOInfoList.Item(i).CompID)
                                        .treLstSO.Nodes(i).SetValue("CompId", FSOInfoList.Item(i).CompID)
                                        .treLstSO.Nodes(i).SetValue("CompName", FSOInfoList.Item(i).CompName)
                                        .treLstSO.Nodes(i).SetValue("WaitProc", "0")
                                        .treLstSO.Nodes(i).SetValue("Procing", "0")
                                        .treLstSO.Nodes(i).StateImageIndex = aStatus
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

                                        For i As Int32 = 0 To .treLstSO.Nodes.Count - 1
                                            If (FThreadStop) OrElse (Not FRegOK) Then
                                                .treLstSO.Nodes(i).SetValue("WaitProc", "0")
                                                .treLstSO.Nodes(i).SetValue("Procing", "0")


                                            End If
                                            If (aStatus <> SOStatus.NotUpdateImg) AndAlso (Not FThreadStop) Then
                                                If FRegOK Then
                                                    .treLstSO.Nodes(i).StateImageIndex = aStatus
                                                End If

                                            End If
                                        Next
                                    Else
                                        Dim aNote As TreeListNode = .treLstSO.FindNodeByFieldValue("CompId", aSO.UseCompId)

                                        If aNote IsNot Nothing Then
                                            If (FThreadStop) OrElse (Not FRegOK) Then
                                                aNote.SetValue("WaitProc", "0")
                                                aNote.SetValue("Procing", "0")
                                            Else

                                                aNote.SetValue("WaitProc", FUseCompIndex.Item(aSO.UseCompId).WaitProcessNumber)
                                                aNote.SetValue("Procing", FUseCompIndex.Item(aSO.UseCompId).ProcessingNumber)
                                            End If
                                        End If
                                    End If
                                Else

                                    Dim aNote As TreeListNode = .treLstSO.FindNodeByFieldValue("CompId", aSO.CompID)
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
                        'aTreeList.ExpandAll()
                        .treLstSO.EndUnboundLoad()
                        .treLstSO.EndUpdate()
                        aSO.Dispose()
                    End Try
                End With
            End If

        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
        End Try
    End Sub



End Module
