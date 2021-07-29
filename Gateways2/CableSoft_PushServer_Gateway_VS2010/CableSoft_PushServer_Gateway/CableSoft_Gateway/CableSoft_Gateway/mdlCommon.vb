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
Imports CableSoft.GateWay.csException
Imports CableSoft.GateWay.Common
Imports CableSoft.GateWay.SystemIO
Imports System.Net
Imports CableSoft.CAS.CryptUtil


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
    Public fMDBFile As String = String.Empty
    Public Const fMDBPassword As String = "chunya36"
    '-------------------公用Table----------------------------------
    Public fSysTable As DataTable = Nothing     '系統參數Table
    Public fSOTable As DataTable = Nothing      'COM區 Table
    Public fPushServerParaTable As DataTable = Nothing      'Nagra參數 Table
    Public fComparedErrorTable As DataTable = Nothing       '錯誤對照表 Table
    Public fUseCompanyTable As DataTable = Nothing          '使用的公司 Table
    Public fTVmailSetTable As DataTable = Nothing               'TVMail訊息設定 Table
    Public fCustomerFieldTable As DataTable = Nothing        '客戶自訂變數Table
    Public fMailInfoTable As DataTable = Nothing                  'Mail基本設定
    Public fMailToTable As DataTable = Nothing                      '收件人
    Public fMailCCTable As DataTable = Nothing                  '副本
    Public fMailPSParaTable As DataTable = Nothing           'PS裡的mail設定檔
    '-------------------------------------------------------------------
    Public fSOInfoList As List(Of SOInfoClass) = Nothing
    Public fUseCompList As List(Of UseCompanyClass) = Nothing


    Public Const fSystemTableName As String = "SysOpt"
    Public Const fSOTableName As String = "SO"
    Public Const fPushServerParaTableName As String = "PushServerParaSet"
    Public Const fUseCompanyTableName As String = "UseCompany"
    Public Const fComparedErrorTableName As String = "ComparedError"
    Public Const fPushServerUrlTableName As String = "PushServerUrls"
    Public Const fTVMailSetTableName As String = "TVMailSet"
    Public Const fCustomerFieldTableName As String = "CustomerField"
    Public fThreadStop As Boolean = False
    '---------------以下為系統公用屬性-------------------------------------
    Public fAppTitle As String = String.Empty
    Public fAutoRunGW As Boolean = False
    Public fUseTray As Boolean = False
    Public fShowResource As Boolean = False
    Public fUseHotKey As Boolean = False
    Public fAutoConnectTime As Int32
    Public fReadDataTime As Integer = 100
    Public fProcessNumber As Int32 = 100
    Public fShowDataCount As Int32 = 200
    Public fClearDataCount As Int32 = 20
    Public fMaxThread As Int32 = 25
    Public fDBPoolMinNumber As Int32 = 0
    Public fDBPoolMaxNumber As Int32 = 100
    Public fDBPoolLiveTime As Int32 = 0
    'Public FTestNagraSock As Int32 = 30
    '---------------Push Server 參數---------------------------
    Public fPushUrl As String = String.Empty
    Public fPushRepeat As String = String.Empty
    Public fPushDuration As String = String.Empty
    Public fPushUrlMethod As String = "POST"
    Public fPushUrlTimeout As Int32 = 30
    Public fPushServerRetryNum As Int32 = 3
    Public fPushServerSlowTime As Int32 = 1
    Public fPushServerUrls As New List(Of String)
    Public fCmdErrSndMail As Boolean
    Public fSendTVMail As Boolean
    Public fNoneResumeCnt As Int32 = 0
    '----------------------------------------------------------------

    '---------------------WEB有問題發送Mail的資訊--------------------------------
    Public Const fMDBMailFileName As String = "WarningEMail.SET"
    Public Const fMDBMailPassword As String = Nothing
    Public Const fMailInfoTableName As String = "CsMailInfo"
    Public Const fMailToTableName As String = "CsMailTo"
    Public Const fMailCCTableName As String = "CsMailCC"
    Public Const fMailPSParaTableName As String = "CsPSPara"
    Public fSmtpHost As String = String.Empty
    Public fSmtplUserId As String = String.Empty
    Public fSmtpPassword As String = String.Empty
    Public fEnabledSSL As Boolean = False
    Public fMailDisplayName As String = String.Empty
    Public fMailSender As String = String.Empty
    Public fSmtpPort As String = "25"
    Public fMailSubject As String = String.Empty
    Public fMailBody As String = String.Empty
    Public fReadWebErrCnt As String = "100"
    Public fIsUseMail As Boolean = False
    '-------------------------------------------------------------------------------------
    '-------------------------------- TVMail 設定檔 -------------------------------------
    Public fTVMailMDBMsgTxt As String = String.Empty
    Public fTVMailMsgString As String = String.Empty
    'Public fTVMailVarLst As ArrayList = Nothing
    Public Const fSplitTVMailMsgChar As String = "<>"
    Public fLstCustomerVar As List(Of String) = Nothing
    
    '------------------------------------------------------------------------------------


    Public fUseCompWhere As String = String.Empty
    Public fErrCodeLst As New Dictionary(Of String, String) '錯誤對照表
    Public fErrRetryLst As New List(Of String)          '錯誤重試清單
    Public fErrResumeLst As New List(Of String)      '錯誤還原清單
    '-------------------------------------------------------------------
    Public fLstSysMsg As List(Of String)
    Public fInsDate As String = String.Empty
    Public fRegDate As String = String.Empty
    Public fRegOK As Boolean = False
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
        Public Property SidDBName As String
        Public Property SidUserId As String
        Public Property SidPassword As String
        Public Property ConnectionString As String
        Public Property LoginID As String
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
        Dim RegKeyName As String = "Cablesoft Push Server Gateway"
        Try
            RegRoot = Registry.LocalMachine

            RegKey = RegRoot.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Run", _
                                        RegistryKeyPermissionCheck.ReadWriteSubTree, _
                                        AccessControl.RegistryRights.SetValue)


            Dim App As String = String.Format("{0} /AutoExec", Application.ExecutablePath)
            If AutoExec Then
                RegKey.SetValue(RegKeyName, App, RegistryValueKind.String)
            Else
                RegKey.DeleteValue(RegKeyName, False)
            End If
        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
            Throw New Exception(ex.Message)

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
                aForm.Text = fAppTitle


            End If
        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
        End Try
    End Sub

    '設定錯誤對照表
    Public Function SetErrCode() As Boolean
        Try
            fComparedErrorTable = CableSoft.GateWay.SystemIO.GetwaySys.ReadErrorCode(fMDBFile, fMDBPassword)
            fErrCodeLst.Clear()
            fErrRetryLst.Clear()
            fErrResumeLst.Clear()
            For i As Int32 = 0 To fComparedErrorTable.Rows.Count - 1
                fErrCodeLst.Add(GetFieldValue(fComparedErrorTable.Rows(i), "ErrorCode"), _
                                GetFieldValue(fComparedErrorTable.Rows(i), "ErrorName"))
                If GetFieldValue(fComparedErrorTable.Rows(i), "ErrorRetry") = "1" Then
                    fErrRetryLst.Add(GetFieldValue(fComparedErrorTable.Rows(i), "ErrorCode"))
                End If
                If GetFieldValue(fComparedErrorTable.Rows(i), "ErrorResume") = "1" Then
                    fErrResumeLst.Add(GetFieldValue(fComparedErrorTable.Rows(i), "ErrorCode"))
                End If
            Next
            Return True
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function
    '設定Push Server參數值
    Public Function SetPushServerPara() As Boolean
        Try
            fPushServerParaTable = CableSoft.GateWay.SystemIO.GetwaySys.ReadPushServerParaSet(fMDBFile, fMDBPassword)
            If (fPushServerParaTable IsNot Nothing) AndAlso (fPushServerParaTable.Rows.Count > 0) Then
                With fPushServerParaTable
                    fPushDuration = GetFieldValue(.Rows(0), "PushServerDuration")
                    fPushRepeat = GetFieldValue(.Rows(0), "PushServerRepeat")
                    fPushUrl = GetFieldValue(.Rows(0), "PushServerUrl")
                    fPushUrlMethod = GetFieldValue(.Rows(0), "PushServerUrlMethod")
                    fPushUrlTimeout = Convert.ToInt32("0" & GetFieldValue(.Rows(0), "PushServerUrlTimeOut"))
                    fPushServerSlowTime = Convert.ToInt32("0" & GetFieldValue(.Rows(0), "PushServerSlowTime"))
                    fPushServerRetryNum = Convert.ToInt32("0" & GetFieldValue(.Rows(0), "PushServerRetryNum"))
                    fCmdErrSndMail = GetFieldValue(.Rows(0), "CmdErrSndMail") = "1"
                    fNoneResumeCnt = Convert.ToInt32("0" & GetFieldValue(.Rows(0), "NoneResumeCnt"))
                    fSendTVMail = Convert.ToInt32("0" & GetFieldValue(.Rows(0), "SendTVMail"))
                End With
            Else
                CableSoft.GateWay.SystemIO.GetwaySys.DefaultPushServerPara(fMDBFile, fMDBPassword, fPushServerParaTableName)
                SetPushServerPara()
            End If
            fPushServerUrls.Clear()
            Dim tblUrl As DataTable = CableSoft.GateWay.SystemIO.GetwaySys.ReadPushServerUrl(fMDBFile, fMDBPassword)
            If tblUrl.Rows.Count = 0 Then
                Throw New Exception("Push Server Url 未設定 ! ")
                Return False
            Else
                For i As Int32 = 0 To tblUrl.Rows.Count - 1
                    fPushServerUrls.Add(GetFieldValue(tblUrl.Rows(i), "PushServerUrl"))
                Next
            End If
            Return True
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function
    Public Function SetMailPara() As Boolean
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
            fMailPSParaTable = MDBCommon.ReadMDBDataTable(aFile, fMDBMailPassword, fMailPSParaTableName, Nothing)

            If fMailInfoTable Is Nothing OrElse fMailInfoTable.Rows.Count <= 0 Then
                MDBMailPara.DefaultMailInfo(aFile, fMDBMailPassword, fMailInfoTableName)
                SetMailPara()
            End If
            If fMailPSParaTable Is Nothing OrElse fMailPSParaTable.Rows.Count <= 0 Then
                MDBMailPara.DefaultPSPara(aFile, fMDBMailPassword, fMailPSParaTableName)
                SetMailPara()
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
            With fMailPSParaTable
                fReadWebErrCnt = GetFieldValue(.Rows(0), "ReadWebErrCnt")
            End With



        Catch ex As Exception
            'Throw New Exception(ex.Message)
            WriteErrTxtLog.WriteSndEmailLog(ex.Message)
        End Try
    End Function


    '設定公用系統值
    Public Function SetSystem() As Boolean
        Try
            fSysTable = CableSoft.GateWay.SystemIO.GetwaySys.ReadSystemIO(fMDBFile, fMDBPassword)
            If (fSysTable IsNot Nothing) AndAlso (fSysTable.Rows.Count > 0) Then
                With fSysTable
                    fAppTitle = GetFieldValue(.Rows(0), "AppCaption")
                    fAutoRunGW = GetFieldValue(.Rows(0), "AutoRunGw") = "1"
                    fUseTray = GetFieldValue(.Rows(0), "UseTray") = "1"
                    fShowResource = Convert.ToInt32("0" & GetFieldValue(.Rows(0), "ShowResource")) = 1
                    fUseHotKey = Convert.ToInt32("0" & GetFieldValue(.Rows(0), "UseHotKey")) = 1
                    fAutoConnectTime = Convert.ToInt32("0" & GetFieldValue(.Rows(0), "AutoConnectTime"))
                    fReadDataTime = Convert.ToInt32("0" & GetFieldValue(.Rows(0), "ReadDataTime"))
                    fProcessNumber = Convert.ToInt32("0" & GetFieldValue(.Rows(0), "ProcessNumber"))
                    fShowDataCount = Convert.ToInt32("0" & GetFieldValue(.Rows(0), "ShowDataCount"))
                    fClearDataCount = Convert.ToInt32("0" & GetFieldValue(.Rows(0), "ClearDataCount"))
                    fMaxThread = Convert.ToInt32("0" & GetFieldValue(.Rows(0), "MaxThread"))

                    fDBPoolMaxNumber = Convert.ToInt32("0" & GetFieldValue(.Rows(0), "DBPoolMaxNumber"))
                    fDBPoolMinNumber = Convert.ToInt32("0" & GetFieldValue(.Rows(0), "DBPoolMinNumber"))
                    fDBPoolLiveTime = Convert.ToInt32("0" & GetFieldValue(.Rows(0), "DBPoolLiveTime"))
                    'FTestNagraSock = GetFieldValue(.Rows(0), "TestSocketTime")


                    '設定顯示在畫面的訊息
                    SetSysMsgString()

                End With
            Else
                CableSoft.GateWay.SystemIO.GetwaySys.DefaultSystem(fMDBFile, fMDBPassword, fSystemTableName)
                SetSystem()
            End If
            Return True
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Function
    Private Sub SetSysMsgString()
        Try
            If fLstSysMsg Is Nothing Then
                fLstSysMsg = New List(Of String)
            Else
                fLstSysMsg.Clear()
            End If
            fLstSysMsg.Add("Parent讀取系統參數設定..")
            fLstSysMsg.Add(String.Format("Windows 開機後自動執行 Gateway : {0}", IIf(fAutoRunGW, "是", "否")))
            fLstSysMsg.Add(String.Format("最小化後常駐在系統工作列上 : {0}", IIf(fUseTray, "是", "否")))
            fLstSysMsg.Add(String.Format("顯示系統資源於狀態列 : {0}", IIf(fShowResource, "是", "否")))
            fLstSysMsg.Add(String.Format("每 {0} 秒讀系統台資料", fReadDataTime))
            fLstSysMsg.Add(String.Format("每次處理 {0} 筆高階指令", fProcessNumber))
            fLstSysMsg.Add(String.Format("畫面資料顯示 {0} 筆訊息 ,  每超過 {1} 筆訊息 ,  則開始清除多餘訊息", fShowDataCount, fClearDataCount))
            fLstSysMsg.Add(String.Format("最大命令處理線程 {0} ", fMaxThread))
            fLstSysMsg.Add(String.Format("DB最小集區 : {0} , DB最大集區 : {1} , 集區Lifetime : {2}", fDBPoolMinNumber, fDBPoolMaxNumber, fDBPoolLiveTime))

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub
    '設定使用公司別
    Public Function SetUseComp() As Boolean
        Try
            Dim _ConnectString As String = "Data Source={0};Persist Security Info=True;" &
                                                            "User ID={1};Password={2};Min Pool Size={3};" &
                                                            "Max Pool Size={4};Unicode=True;Load Balance Timeout={5}"

            fUseCompWhere = String.Empty
            fUseCompanyTable = CableSoft.GateWay.SystemIO.GetwaySys.ReadUseComp(fMDBFile, fMDBPassword)
            If fUseCompanyTable IsNot Nothing AndAlso fUseCompanyTable.Rows.Count > 0 Then

                If fUseCompList Is Nothing Then
                    fUseCompList = New List(Of UseCompanyClass)
                End If
                fUseCompList.Clear()
                For i As Int32 = 0 To fUseCompanyTable.Rows.Count - 1
                    If Convert.ToUInt32(GetFieldValue(fUseCompanyTable.Rows(i), "SelectID")) = 1 Then
                        Dim obj As New UseCompanyClass
                        obj.CompID = Convert.ToUInt32(GetFieldValue(fUseCompanyTable.Rows(i), "CompID"))
                        obj.CompName = GetFieldValue(fUseCompanyTable.Rows(i), "CompName")
                        obj.SidDBName = GetFieldValue(fUseCompanyTable.Rows(i), "SidDBName")
                        obj.SidUserId = GetFieldValue(fUseCompanyTable.Rows(i), "SidUserId")
                        obj.SidPassword = GetFieldValue(fUseCompanyTable.Rows(i), "SidPassword")
                        obj.ConnectionString = String.Empty
                        If (Not String.IsNullOrEmpty(obj.SidDBName)) OrElse
                            (Not String.IsNullOrEmpty(obj.SidUserId)) OrElse
                             (Not String.IsNullOrEmpty(obj.SidPassword)) Then
                            obj.ConnectionString = String.Format(_ConnectString, obj.SidDBName,
                                                                                obj.SidUserId, obj.SidPassword,
                                                                                 "0",
                                                                                fDBPoolMaxNumber.ToString,
                                                                                fDBPoolLiveTime)

                        End If

                        fUseCompList.Add(obj)
                        If String.IsNullOrEmpty(fUseCompWhere) Then
                            fUseCompWhere = obj.CompID
                        Else
                            fUseCompWhere = fUseCompWhere & "," & obj.CompID
                        End If
                    End If

                Next
                If fUseCompList Is Nothing OrElse fUseCompList.Count <= 0 Then
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
    '讀取TVMail訊息設定
    Public Function SetTVMailMsg() As Boolean
        Try

            fTVmailSetTable = CableSoft.GateWay.SystemIO.GetwaySys.ReadTable(fMDBFile, fMDBPassword, fTVMailSetTableName)
            fTVMailMDBMsgTxt = String.Empty
            If (fTVmailSetTable IsNot Nothing) AndAlso (fTVmailSetTable.Rows.Count > 0) Then
                fTVMailMDBMsgTxt = GetFieldValue(fTVmailSetTable.Rows(0), "TVMailMsg")
            End If
            Return True
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function
    '讀取TVMailMsg可設定變數內容 
    Public Function SetCustomerField() As Boolean
        Try
            Dim aFieldString As String = String.Empty
            fCustomerFieldTable = CableSoft.GateWay.SystemIO.GetwaySys.ReadTable(fMDBFile, fMDBPassword, fCustomerFieldTableName)

            If fCustomerFieldTable.Rows.Count > 0 Then
                aFieldString = GetFieldValue(fCustomerFieldTable.Rows(0), "VarField1")
            Else
                Dim aRow As DataRow = fCustomerFieldTable.NewRow()
                Dim aMsg As String = String.Empty
                Try
                    Using cn As New OracleConnection(fSOInfoList.Item(0).OraConnectString)
                        cn.Open()
                        Using cmd As New OracleCommand("select * from so563 where 1=0", cn)
                            Dim rd As OracleDataReader = cmd.ExecuteReader(CommandBehavior.CloseConnection)
                            For i As Int32 = 0 To rd.FieldCount - 1
                                If String.IsNullOrEmpty(aMsg) Then
                                    aMsg = rd.GetName(i)
                                Else
                                    aMsg = aMsg & "," & rd.GetName(i)
                                End If
                            Next
                        End Using
                    End Using
                    aRow.Item("VarField1") = aMsg
                    fCustomerFieldTable.Rows.Add(aRow)
                    CableSoft.GateWay.Common.MDBCommon.SaveMDBDataTable(fMDBFile, fMDBPassword,
                                                                        fCustomerFieldTableName, fCustomerFieldTable)
                    SetCustomerField()
                Catch ex As Exception

                End Try
                
            End If
            If Not String.IsNullOrEmpty(aFieldString) Then
                fLstCustomerVar = New List(Of String)
                fLstCustomerVar.AddRange(aFieldString.Split(","))
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
            If fSOInfoList IsNot Nothing Then
                fSOInfoList.Clear()
            End If

            fSOTable = CableSoft.GateWay.SystemIO.GetwaySys.ReadSO(fMDBFile, fMDBPassword)

            If (fSOTable IsNot Nothing) AndAlso (fSOTable.Rows.Count > 0) Then
                If fSOInfoList Is Nothing Then
                    fSOInfoList = New List(Of SOInfoClass)
                End If

                fSOInfoList.Clear()
                For i As Int32 = 0 To fSOTable.Rows.Count - 1

                    Dim obj As New SOInfoClass()

                    obj.OraConnectString = String.Format(_ConnectString, _
                         GetFieldValue(fSOTable.Rows(i), "SID"), _
                         GetFieldValue(fSOTable.Rows(i), "UserId"), _
                         GetFieldValue(fSOTable.Rows(i), "UserPassword"), _
                         "0", _
                         fDBPoolMaxNumber.ToString, _
                         fDBPoolLiveTime)
                    obj.ProcessingNumber = 0
                    obj.WaitProcessNumber = 0

                    obj.CompID = -99
                    obj.CompName = GetFieldValue(fSOTable.Rows(i), "SID")
                    obj.ConnectionOK = False
                    fSOInfoList.Add(obj)
                Next
            Else
                Throw New Exception("COM區未設定連線資料")
            End If


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
        'Dim aRet As Boolean = False
        Try
            If obj Is Nothing Then
                Return False
            End If
            If obj.State = ConnectionState.Closed Then
                Try
                    obj.Open()
                    obj.Close()
                    'aRet = True
                    Return True
                Catch ex As Exception
                    If aErrMsg IsNot Nothing Then
                        aErrMsg = ex.Message
                    End If
                    'aRet = False
                    Return False
                End Try

                'aRet = True
            End If
            Return True

        Catch ex As Exception
            Return False
        End Try

    End Function


    Public Function GetFieldValue(ByVal row As DataRow, ByVal FieldName As String, _
                                  Optional ByVal blnDecrypt As Boolean = True) As String
        Try
            If (Not DBNull.Value.Equals(row.Item(FieldName))) Then
                If String.IsNullOrEmpty(row.Item(FieldName)) Then
                    Return String.Empty
                End If
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

                    Dim trlParent As TreeListNode = Nothing
                    Dim trlChild As TreeListNode = Nothing


                    If ((fFrom.TreLstConnStatus.Nodes.Count - fShowDataCount) > fClearDataCount) Then
                        'fFrom.TreLstConnStatus.Nodes.RemoveAt(fFrom.TreLstConnStatus.Nodes.Count - 1)
                        fFrom.TreLstConnStatus.Nodes.Clear()
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
                    fFrom.TreLstConnStatus.ExpandAll()
                    fFrom.TreLstConnStatus.MoveFirst()
                    fFrom.TreLstConnStatus.EndUpdate()


                 
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

                    If ((aTreeList.Nodes.Count - fShowDataCount) > fClearDataCount) Then
                        aTreeList.Nodes.RemoveAt(aTreeList.Nodes.Count - 1)
                    End If
                    For i As Int32 = 0 To fLstSysMsg.Count - 1
                        If fLstSysMsg.Item(i).ToString.ToUpper.StartsWith("PARENT") Then
                            trlParent = aTreeList.AppendNode(Nothing, -1)

                            trlParent.SetValue("SysInfo", fLstSysMsg.Item(i).Replace("Parent", ""))
                            trlParent.StateImageIndex = MsgStyle.Info
                            '記錄事件
                            WriteErrTxtLog.WriteSysEventLog(fLstSysMsg.Item(i).Replace("Parent", ""))
                        Else
                            trlParent = aTreeList.Nodes.LastNode
                            trlChild = aTreeList.AppendNode(Nothing, trlParent)
                            trlChild.SetValue("SysInfo", fLstSysMsg.Item(i))
                            trlChild.StateImageIndex = aMsgStyle
                            '記錄事件
                            WriteErrTxtLog.WriteSysEventLog(fLstSysMsg.Item(i))
                        End If
                    Next
                Finally

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
    Public Sub UpdateSysMsgUI()


        Try

            Dim trlParent As TreeListNode = Nothing
            Dim trlChild As TreeListNode = Nothing


            If fLstSysMsg Is Nothing Then
                Exit Sub
            End If
            If fFrom.InvokeRequired Then
                Dim act As New Action(AddressOf UpdateSysMsgUI)
                fFrom.Invoke(act)

            Else
                With fFrom
                    Try
                        .TreLstSysMsg.ClearNodes()
                        .TreLstSysMsg.BeginUpdate()

                        For i As Int32 = 0 To fLstSysMsg.Count - 1
                            If fLstSysMsg.Item(i).ToString.ToUpper.StartsWith("PARENT") Then
                                trlParent = .TreLstSysMsg.AppendNode(Nothing, i)

                                trlParent.SetValue("SysInfo", fLstSysMsg.Item(i).Replace("Parent", ""))
                                trlParent.StateImageIndex = MsgStyle.Info
                                '記錄事件
                                WriteErrTxtLog.WriteSysEventLog(fLstSysMsg.Item(i).Replace("Parent", ""))
                            Else
                                trlParent = .TreLstSysMsg.Nodes.LastNode
                                trlChild = .TreLstSysMsg.AppendNode(Nothing, trlParent)
                                trlChild.SetValue("SysInfo", fLstSysMsg.Item(i))
                                trlChild.StateImageIndex = MsgStyle.InfoLa
                                '記錄事件
                                WriteErrTxtLog.WriteSysEventLog(fLstSysMsg.Item(i))
                            End If
                        Next

                    Finally


                        .TreLstSysMsg.EndUpdate()
                        .TreLstSysMsg.ExpandAll()
                        .TreLstSysMsg.MoveFirst()
                    End Try
                End With
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
                        SyncLock fShowDataCount.GetType
                            .TreLstSysErr.BeginUpdate()

                            Dim trlParent As TreeListNode = Nothing
                            Dim trlChild As TreeListNode = Nothing
                            If ((.TreLstSysErr.Nodes.Count - fShowDataCount) > fClearDataCount) Then
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


                        End SyncLock
                    Finally

                        .TreLstSysErr.EndUpdate()
                        .TreLstSysErr.ExpandAll()
                        .TreLstSysErr.MoveFirst()
                    End Try
                End With
            End If
        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
        Finally

        End Try

    End Sub
    '更新低階命令視窗
    Public Sub UpdateLowCmdUI(ByVal aCMDInfo As Dictionary(Of String, String))

        Try
            If fFrom.InvokeRequired Then
                Dim act As New Action(Of Dictionary(Of String, String))(AddressOf UpdateLowCmdUI)
                fFrom.Invoke(act, aCMDInfo)
            Else
                Try
                    fFrom.TreLstLowCmd.BeginUpdate()

                    If ((fFrom.TreLstLowCmd.Nodes.Count - fShowDataCount) > fClearDataCount) Then
                        fFrom.TreLstLowCmd.Nodes.RemoveAt(0)
                    End If
                    Dim trlParent As TreeListNode = fFrom.TreLstLowCmd.AppendNode(Nothing, -1)
                    For i As Int32 = 0 To fFrom.TreLstLowCmd.Columns.Count - 1
                        Dim aFieldName As String = fFrom.TreLstLowCmd.Columns(i).FieldName
                        trlParent.SetValue(aFieldName, aCMDInfo.Item(aFieldName.ToUpper))
                    Next
                    Select Case aCMDInfo.Item("CmdStatus".ToUpper)
                        Case "S"
                            trlParent.StateImageIndex = MsgStyle.OKLa
                        Case Else
                            trlParent.StateImageIndex = MsgStyle.ErrorLa
                    End Select
                Finally


                    fFrom.TreLstLowCmd.EndUpdate()
                    fFrom.TreLstLowCmd.MoveLast()
                End Try


            End If
        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
        Finally
            'aCMDInfo.Clear()
            'aCMDInfo = Nothing
        End Try

    End Sub
    '更新高階命令視窗
    Public Sub UpdateHightCmdUI(ByVal aCmdInfo As Dictionary(Of String, String), _
                        ByVal aType As UpdCmdMode)
        Try

            If fFrom.InvokeRequired Then
                Dim act As New Action(Of Dictionary(Of String, String), UpdCmdMode)(AddressOf UpdateHightCmdUI)
                fFrom.Invoke(act, aCmdInfo, aType)
            Else
                
                Try
                    'Application.DoEvents()
                    Dim trlParent As TreeListNode = Nothing
                    fFrom.TreLstInstruct.BeginUpdate()

                    If aType = UpdCmdMode.Add Then
                        aCmdInfo.Item("GUIDNO") = Guid.NewGuid.ToString
                        If ((fFrom.TreLstInstruct.Nodes.Count - fShowDataCount) > fClearDataCount) Then
                            fFrom.TreLstInstruct.Nodes.RemoveAt(0)
                        End If
                        trlParent = fFrom.TreLstInstruct.AppendNode(Nothing, -1)
                    End If
                    If aType = UpdCmdMode.Upd Then
                        trlParent = fFrom.TreLstInstruct.FindNodeByFieldValue("GuidNo", aCmdInfo.Item("GUIDNO"))
                    End If
                    If trlParent IsNot Nothing Then
                        For i As Int32 = 0 To fFrom.TreLstInstruct.Columns.Count - 1
                            Dim aFieldName As String = fFrom.TreLstInstruct.Columns(i).FieldName
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

                    fFrom.TreLstInstruct.EndUpdate()
                    fFrom.TreLstInstruct.MoveLast()

                End Try
            End If


        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
        Finally
            'aCmdInfo.Clear()
            'aCmdInfo = Nothing
        End Try

    End Sub


    '更新SO Information的畫面
    Public Sub UpdateSOStatusUI(ByVal aSO As SOInfoClass, _
                                ByVal aStatus As SOStatus)
        Try
            If fFrom.InvokeRequired Then
                Dim act As New Action(Of SOInfoClass, SOStatus)(AddressOf UpdateSOStatusUI)
                fFrom.Invoke(act, aSO, aStatus)

            Else
                With fFrom
                    Try
                        .treLstSO.BeginUpdate()

                        Select Case aStatus
                            Case SOStatus.Initial, SOStatus.Key
                                .treLstSO.ClearNodes()
                                If (fSOInfoList.Count = 1) AndAlso (fSOInfoList.Item(0).CompID = -99) Then

                                    For i As Int32 = 0 To fUseCompList.Count - 1
                                        .treLstSO.AppendNode(Nothing, fUseCompList.Item(i).CompID)
                                        .treLstSO.Nodes(i).SetValue("CompId", fUseCompList.Item(i).CompID)
                                        .treLstSO.Nodes(i).SetValue("CompName", fUseCompList.Item(i).CompName)
                                        .treLstSO.Nodes(i).SetValue("WaitProc", "0")
                                        .treLstSO.Nodes(i).SetValue("Procing", "0")
                                        .treLstSO.Nodes(i).StateImageIndex = aStatus
                                        SyncLock fUseCompList(i).ProcessingNumber.GetType
                                            fUseCompList(i).ProcessingNumber = 0
                                        End SyncLock
                                        SyncLock fUseCompList(i).WaitProcessNumber.GetType
                                            fUseCompList(i).WaitProcessNumber = 0
                                        End SyncLock
                                    Next

                                Else

                                    For i As Int32 = 0 To fSOInfoList.Count - 1
                                        .treLstSO.AppendNode(Nothing, fSOInfoList.Item(i).CompID)
                                        .treLstSO.Nodes(i).SetValue("CompId", fSOInfoList.Item(i).CompID)
                                        .treLstSO.Nodes(i).SetValue("CompName", fSOInfoList.Item(i).CompName)
                                        .treLstSO.Nodes(i).SetValue("WaitProc", "0")
                                        .treLstSO.Nodes(i).SetValue("Procing", "0")
                                        .treLstSO.Nodes(i).StateImageIndex = aStatus
                                        SyncLock fSOInfoList(i).ProcessingNumber.GetType
                                            fSOInfoList(i).ProcessingNumber = 0
                                        End SyncLock
                                        SyncLock fSOInfoList(i).WaitProcessNumber.GetType
                                            fSOInfoList(i).WaitProcessNumber = 0
                                        End SyncLock

                                    Next
                                End If

                            Case Else

                                If aSO.CompID = -99 Then
                                    If aStatus <> SOStatus.NotUpdateImg Then

                                        For i As Int32 = 0 To .treLstSO.Nodes.Count - 1
                                            If (fThreadStop) OrElse (Not fRegOK) Then
                                                .treLstSO.Nodes(i).SetValue("WaitProc", "0")
                                                .treLstSO.Nodes(i).SetValue("Procing", "0")

                                            End If
                                            If (aStatus <> SOStatus.NotUpdateImg) AndAlso (Not fThreadStop) Then
                                                If fRegOK Then
                                                    .treLstSO.Nodes(i).StateImageIndex = aStatus
                                                End If

                                            End If
                                        Next
                                    Else
                                        Dim aNote As TreeListNode = .treLstSO.FindNodeByFieldValue("CompId", aSO.UseCompId)

                                        If aNote IsNot Nothing Then
                                            If (fThreadStop) OrElse (Not fRegOK) OrElse (FThreadsTotal <= 0) Then
                                                aNote.SetValue("WaitProc", "0")
                                                aNote.SetValue("Procing", "0")
                                            Else
                                                aNote.SetValue("WaitProc", fUseCompIndex.Item(aSO.UseCompId).WaitProcessNumber)
                                                aNote.SetValue("Procing", fUseCompIndex.Item(aSO.UseCompId).ProcessingNumber)
                                            End If
                                        End If

                                    End If
                                Else

                                    Dim aNote As TreeListNode = .treLstSO.FindNodeByFieldValue("CompId", aSO.CompID)
                                    If aNote IsNot Nothing Then
                                        If (fThreadStop) OrElse (Not fRegOK) Then
                                            aNote.SetValue("WaitProc", "0")
                                            aNote.SetValue("Procing", "0")
                                        Else
                                            aNote.SetValue("WaitProc", aSO.WaitProcessNumber)
                                            aNote.SetValue("Procing", aSO.ProcessingNumber)
                                        End If

                                        If (aStatus <> SOStatus.NotUpdateImg) AndAlso (Not fThreadStop) Then
                                            If fRegOK Then
                                                aNote.StateImageIndex = aStatus
                                            End If

                                        End If

                                    End If
                                End If
                        End Select
                    Finally
                        'aTreeList.ExpandAll()

                        .treLstSO.EndUpdate()

                    End Try
                End With
            End If
        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
        End Try
    End Sub



End Module
