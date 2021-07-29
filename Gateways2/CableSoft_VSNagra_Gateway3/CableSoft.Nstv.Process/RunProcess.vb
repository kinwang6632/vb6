Imports System.Data
Imports System.Net
Imports System.Threading
Imports System.IO
Imports System
Imports System.Text
Imports System.Windows.Forms
Imports CableSoft.NSTV.Log
Imports CableSoft.NSTV.BuilderLowerCmd
Imports CableSoft.NSTV.Encry
Imports System.Data.OracleClient
Imports System.Globalization
Imports CableSoft.CardLess
Imports CableSoft.CardLess.SocketClient
Imports CableSoft.CardLess.SODB

Public Class RunProcess
    Implements IDisposable

    Private xmlFileName As String = "CardLess.Set"
    Private Const IsEncrypt As Boolean = True
    Private tbSystem As DataTable = Nothing
    Private tbSO As DataTable = Nothing
    Private tbErrorList As DataTable = Nothing
    'Private tbLowCmd As DataTable = Nothing
    'Private tbHighCmd As DataTable = Nothing
    Private tbNagra As DataTable = Nothing
    Private tbServer As DataTable = Nothing
    Private xmlIO As CableSoft.CardLess.XMLFileIO.XmlFileIO = Nothing
    Private SODB() As SODB.SO = Nothing
    Private SckClient() As Client = Nothing
    Private Const RegisterErrorFileName As String = "RegError.txt"
    Private Const ErrorLogDirectory As String = "ErrorLog"
    Public Property NeedRegister As Boolean = True
    Public Property ExeName As String = Nothing
    Public Sub New()
        xmlFileName = Application.StartupPath & "\" & xmlFileName
    End Sub
    Public Sub New(ByVal xmlFile As String)
        xmlFileName = xmlFile
    End Sub

    Public Sub RunProcess()
        Try
            Dim RegEndDate As Date
            If Me.NeedRegister Then
                Dim IsRegisterOK As Boolean
                If String.IsNullOrEmpty(ExeName) Then
                    IsRegisterOK = CableSoft_KeyPro.GetSystemInfo.IsRegister(CableSoft_KeyPro.ShowForm.No)
                Else
                    IsRegisterOK = CableSoft_KeyPro.GetSystemInfo.IsRegister(Application.StartupPath & "\KeyPro.Lic",
                                                                            ExeName, CableSoft_KeyPro.ShowForm.No)
                End If
                If Not IsRegisterOK Then
                    CableSoft.NSTV.Log.NstvLog.WriteErrorLog(New Exception("註冊不合法"), "RunProcess")
                    Exit Sub
                End If
                Dim RegInfo As String() = CableSoft_KeyPro.GetSystemInfo.DecryKeyPro.Split(Environment.NewLine)
                Dim InstallDate As Date = Date.Parse(RegInfo(5), CultureInfo.CreateSpecificCulture("zh-TW")).Date
                RegEndDate = Date.Parse(RegInfo(6), CultureInfo.CreateSpecificCulture("zh-TW")).Date
            Else
                RegEndDate = New DateTime(9999, 12, 31)
            End If

            xmlIO = New XMLFileIO.XmlFileIO(xmlFileName)
            tbSystem = xmlIO.ReadSystem(IsEncrypt)
            tbSO = xmlIO.ReadSO(IsEncrypt)
            tbErrorList = xmlIO.ReadErrorList(IsEncrypt)           
            tbNagra = xmlIO.ReadNagra(IsEncrypt)
            tbServer = xmlIO.ReadServer(IsEncrypt, "Server")
            If tbSO.Select("IsSelect = '1'").Length = 0 Then Exit Sub
            Array.Resize(SODB, tbSO.Select("IsSelect = '1'").Length)
            Array.Resize(SckClient, tbServer.Rows.Count)
            Dim rwArray() As DataRow = tbSO.Select("IsSelect = '1'")

            For i As Integer = 0 To tbServer.Rows.Count - 1
                SckClient(i) = New Client(tbNagra.Copy, tbServer.Rows(i).Item("SourceId"),
                                          tbServer.Rows(i).Item("DestId").ToString,
                                          tbServer.Rows(i).Item("MoppId").ToString, IsEncrypt)
                SckClient(i).ConnectNagra()
            Next
            Thread.Sleep(1000)
            RegEndDate = New DateTime(2099, 12, 31)
            For i As Integer = 0 To rwArray.Length - 1
                SODB(i) = New SO(rwArray(i).Item("CompCode"), RegEndDate, xmlFileName, IsEncrypt, SckClient)
                SODB(i).Run()
            Next


        Catch ex As Exception
            NstvLog.WriteErrorLog(ex, "讀取參數失敗!")
        End Try
    End Sub


#Region "IDisposable Support"
    Private disposedValue As Boolean ' 偵測多餘的呼叫

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: 處置 Managed 狀態 (Managed 物件)。
                If tbSystem IsNot Nothing Then
                    tbSystem.Dispose()
                    tbSystem = Nothing
                End If
                If tbErrorList IsNot Nothing Then
                    tbErrorList.Dispose()
                    tbErrorList = Nothing
                End If
                If tbHighCmd IsNot Nothing Then
                    tbHighCmd.Dispose()
                    tbHighCmd = Nothing
                End If
                If tbLowCmd IsNot Nothing Then
                    tbLowCmd.Dispose()
                    tbLowCmd = Nothing
                End If
                If tbNagra IsNot Nothing Then
                    tbNagra.Dispose()
                    tbNagra = Nothing
                End If
                If tbServer IsNot Nothing Then
                    tbServer.Dispose()
                    tbServer = Nothing
                End If
                If tbSO IsNot Nothing Then
                    tbSO.Dispose()
                    tbSO = Nothing
                End If

                If SODB IsNot Nothing AndAlso SODB.Length > 0 Then
                    For Each so As SO In SODB
                        Try
                            If so IsNot Nothing Then
                                so.StopProcess = True
                                so.Dispose()
                            End If
                            'so = Nothing
                        Catch ex As Exception
                        Finally
                            so = Nothing
                        End Try
                    Next
                    Erase SODB
                End If
                If SckClient IsNot Nothing AndAlso SckClient.Length > 0 Then
                    For Each sck As Client In SckClient
                        Try
                            If sck IsNot Nothing Then
                                sck.Dispose()
                            End If
                        Finally
                            sck = Nothing
                        End Try
                    Next
                    Erase SckClient
                End If
                If xmlIO IsNot Nothing Then
                    xmlIO.Dispose()
                    xmlIO = Nothing
                End If
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
'Public Class SOInfo
'    Implements IDisposable
'    Public Property ProcessingCount As Int32
'    Public Property WaitProcessCount As Int32
'    Public Property CompCode As Int32
'    Public Property cn As OracleConnection
'    Public Property identify As String
'    Public Sub New(ByVal CompCode As Int32,
'                   ByVal SID As String,
'                   ByVal UserId As String,
'                   ByVal Password As String,
'                   ByVal MaxPool As Integer,
'                   ByVal MinPool As Integer,
'                   ByVal LiveTime As Integer
'                   )
'        Try
'            Dim bduConnectString As New OracleConnectionStringBuilder
'            bduConnectString.Pooling = True
'            bduConnectString.PersistSecurityInfo = True
'            bduConnectString.DataSource = SID
'            bduConnectString.UserID = UserId
'            bduConnectString.Password = Password
'            bduConnectString.MaxPoolSize = MaxPool
'            bduConnectString.MinPoolSize = MinPool
'            bduConnectString.LoadBalanceTimeout = LiveTime
'            cn = New OracleConnection(bduConnectString.ConnectionString)
'            Me.identify = Guid.NewGuid.ToString()
'        Catch ex As Exception
'            NstvLog.WriteErrorLog(ex, "SO系統台建立失敗")
'        End Try

'    End Sub
'    Public Function IsConnect() As Boolean
'        Try
'            If Me.cn.State <> ConnectionState.Open Then
'                cn.Open()
'            End If
'        Catch ex As Exception
'            Return False
'        End Try
'        Return True
'    End Function

'#Region "IDisposable Support"
'    Private disposedValue As Boolean ' 偵測多餘的呼叫

'    ' IDisposable
'    Protected Overridable Sub Dispose(disposing As Boolean)
'        If Not Me.disposedValue Then
'            If disposing Then
'                OracleConnection.ClearAllPools()
'                If cn IsNot Nothing Then
'                    cn.Close()
'                    cn.Dispose()
'                    cn = Nothing

'                End If
'                ' TODO: 處置 Managed 狀態 (Managed 物件)。
'            End If

'            ' TODO: 釋放 Unmanaged 資源 (Unmanaged 物件) 並覆寫下面的 Finalize()。
'            ' TODO: 將大型欄位設定為 null。
'        End If
'        Me.disposedValue = True
'    End Sub

'    ' TODO: 只有當上面的 Dispose(ByVal disposing As Boolean) 有可釋放 Unmanaged 資源的程式碼時，才覆寫 Finalize()。
'    'Protected Overrides Sub Finalize()
'    '    ' 請勿變更此程式碼。在上面的 Dispose(ByVal disposing As Boolean) 中輸入清除程式碼。
'    '    Dispose(False)
'    '    MyBase.Finalize()
'    'End Sub

'    ' 由 Visual Basic 新增此程式碼以正確實作可處置的模式。
'    Public Sub Dispose() Implements IDisposable.Dispose
'        ' 請勿變更此程式碼。在以上的 Dispose 置入清除程式碼 (ByVal 視為布林值處置)。
'        Dispose(True)
'        GC.SuppressFinalize(Me)
'    End Sub
'#End Region

'End Class