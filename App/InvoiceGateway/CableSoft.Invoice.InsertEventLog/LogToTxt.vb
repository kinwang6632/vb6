Imports System
Imports System.IO
Imports System.Threading
Public Class LogToTxt
    Implements IDisposable
    Private CompCode As String
    Private IsLog As Boolean
    Private LogReserveTime As Integer
    Private Const Status As String = "Status"
    Private Const LogSQLDirectory As String = "SQLText"    
    Private StatusDirectory As String
    Private LogDirectory As String
    Private mtx As Threading.Mutex = Nothing
    Private mtxFile As Mutex = Nothing
    Private IsCanLogSQL As Boolean
    Private MonitorPath As String
    Private Const C0401FileName As String = "C0401.Status"
    Private Const C0501FileName As String = "C0501.Status"
    Private Const D0401FileName As String = "D0401.Status"
    Private Const D0501FileName As String = "D0501.Status"
    Private Const X0401FileName As String = "X0401.Status"
    Private Const C0701FileName As String = "C0701.Status"
    Private Const B08Filename As String = "B08.Status"
    Public Function IsRunC0401() As Boolean
        Return File.Exists(String.Format("{0}\{1}",
                                          String.Format("{0}\{1}\{2}",
                                                                MonitorPath,
                                                                CompCode,
                                                                Status),
                                          C0401FileName))
    End Function
    Public Function IsRunC0501() As Boolean
        Return File.Exists(String.Format("{0}\{1}",
                                          String.Format("{0}\{1}\{2}",
                                                                MonitorPath,
                                                                CompCode,
                                                                Status),
                                          C0501FileName))
    End Function
    Public Function IsRunD0401() As Boolean
        Return File.Exists(String.Format("{0}\{1}",
                                          String.Format("{0}\{1}\{2}",
                                                                MonitorPath,
                                                                CompCode,
                                                                Status),
                                          D0401FileName))
    End Function
    Public Function IsRunD0501() As Boolean
        Return File.Exists(String.Format("{0}\{1}",
                                          String.Format("{0}\{1}\{2}",
                                                                MonitorPath,
                                                                CompCode,
                                                                Status),
                                          D0501FileName))
    End Function
    Public Function IsRunC0701() As Boolean
        Return File.Exists(String.Format("{0}\{1}",
                                          String.Format("{0}\{1}\{2}",
                                                                MonitorPath,
                                                                CompCode,
                                                                Status),
                                          C0701FileName))
    End Function
    Public Function IsRunX0401() As Boolean
        Return File.Exists(String.Format("{0}\{1}",
                                          String.Format("{0}\{1}\{2}",
                                                               MonitorPath,
                                                                CompCode,
                                                                Status),
                                          X0401FileName))
    End Function
    Public Function IsRunB08() As Boolean
        Return File.Exists(String.Format("{0}\{1}",
                                          String.Format("{0}\{1}\{2}",
                                                                MonitorPath,
                                                                CompCode,
                                                                Status),
                                          B08Filename))
    End Function
    Public Sub New(ByVal CompCode As String,
                   ByVal IsLog As Boolean,
                   ByVal LogReserveTime As Integer,
                   ByVal MonitorPath As String)
        Try
            Me.MonitorPath = MonitorPath
            IsCanLogSQL = False
            If (IsLog) AndAlso (LogReserveTime > 0) Then
                IsCanLogSQL = True
            End If
            mtx = New Threading.Mutex()
            mtxFile = New Mutex
            Me.CompCode = CompCode
            Me.IsLog = IsLog
            Me.LogReserveTime = LogReserveTime
            StatusDirectory = String.Format("{0}\{1}\{2}",
                             MonitorPath,
                             CompCode,
                             Status)
            LogDirectory = String.Format("{0}\{1}\{2}",
                             MonitorPath,
                             CompCode,
                             LogSQLDirectory)
            If Not Directory.Exists(StatusDirectory) Then
                Directory.CreateDirectory(StatusDirectory)
            End If

            If Not Directory.Exists(LogDirectory) Then
                Directory.CreateDirectory(LogDirectory)
            End If

            DeleteAllStatus()
        Catch ex As Exception
            IsCanLogSQL = False
            NSTV.Log.NstvLog.WriteErrorLog(ex, Nothing)
        End Try
    End Sub
    Public Sub New(ByVal CompCode As String,
                   ByVal IsLog As Boolean,
                   ByVal LogReserveTime As Integer)
        
        Me.New(CompCode, IsLog, LogReserveTime,
               Path.GetDirectoryName(AppDomain.CurrentDomain.BaseDirectory))
    End Sub
    Private Sub DeleteLog(ByVal aPath As String, ByVal reserve As Integer, ByVal aNow As DateTime)
        Try
            For Each s As DirectoryInfo In New DirectoryInfo(aPath).GetDirectories
                If (IsNumeric(s.Name)) AndAlso (s.Name.Length = 8) Then
                    If Integer.Parse(aNow.ToString("yyyyMM")) - Integer.Parse(Left(s.Name, 6)) >= reserve Then
                        Directory.Delete(s.FullName, True)
                    End If
                End If
            Next
        Catch ex As Exception

        End Try
    End Sub
    Private Sub LogSQLFile(ByVal InvRunType As CableSoft.B07.InvType.InvTypeEnum.InvFileType,
                           ByVal SQLText As String, ProcessTimeCount As Double,
                           ProcessDate As DateTime)
        mtxFile.WaitOne()
        Try
            DeleteLog(LogDirectory, Me.LogReserveTime, ProcessDate)
            If Not Directory.Exists(String.Format("{0}\{1}", LogDirectory,
                                              ProcessDate.ToString("yyyyMMdd"))) Then

                Directory.CreateDirectory(String.Format("{0}\{1}", LogDirectory,
                                              ProcessDate.ToString("yyyyMMdd")))

            End If
            If Not Directory.Exists(String.Format("{0}\{1}\{2}", LogDirectory,
                                             ProcessDate.ToString("yyyyMMdd"),
                                             ProcessDate.ToString("yyyyMMddHHmmss"))) Then

                Directory.CreateDirectory(String.Format("{0}\{1}\{2}", LogDirectory,
                                              ProcessDate.ToString("yyyyMMdd"), ProcessDate.ToString("yyyyMMddHHmmss")))

            End If



            Using WriteFile As New StreamWriter(String.Format("{0}\{1}\{2}\{3}", LogDirectory,
                                              ProcessDate.ToString("yyyyMMdd"),
                                              ProcessDate.ToString("yyyyMMddHHmmss"), "RunSQL.Log"), True)

                Dim strbdu As New System.Text.StringBuilder

                strbdu.AppendLine([Enum].GetName(GetType(CableSoft.B07.InvType.InvTypeEnum.InvFileType),
                                                 InvRunType))
                strbdu.AppendLine(SQLText)
                strbdu.AppendLine("執行語法時間：" & ProcessTimeCount.ToString & " 秒")
                strbdu.AppendLine(New String("-", 150))
                WriteFile.Write(strbdu.ToString)
                WriteFile.Close()
                WriteFile.Dispose()
            End Using

        Catch ex As Exception
            NSTV.Log.NstvLog.WriteErrorLog(ex, Nothing)
        Finally
            mtxFile.ReleaseMutex()
        End Try
    End Sub
    Private Sub WriteStatus(ByVal FileName As String, ByVal Status As Status)
        Try

            If Status = InsertEventLog.Status.Complete Then
                File.Delete(String.Format("{0}\{1}",
                                          StatusDirectory,
                                          FileName))
            Else
                File.Create(String.Format("{0}\{1}",
                                          StatusDirectory,
                                          FileName)).Close()
            End If
        Catch ex As Exception
            NSTV.Log.NstvLog.WriteErrorLog(ex, Nothing)
        End Try
    End Sub
    Public Sub UpdateProcessStatus(ByVal InvRunType As CableSoft.B07.InvType.InvTypeEnum.InvFileType,
                                   ByVal Status As Status)
        mtx.WaitOne()
        Try

            Select Case InvRunType
                Case B07.InvType.InvTypeEnum.InvFileType.A0401, B07.InvType.InvTypeEnum.InvFileType.C0401
                    WriteStatus(C0401FileName, Status)
                Case B07.InvType.InvTypeEnum.InvFileType.C0501, B07.InvType.InvTypeEnum.InvFileType.A0501
                    WriteStatus(C0501FileName, Status)
                Case B07.InvType.InvTypeEnum.InvFileType.C0701
                    WriteStatus(C0701FileName, Status)
                Case B07.InvType.InvTypeEnum.InvFileType.D0401
                    WriteStatus(D0401FileName, Status)
                Case B07.InvType.InvTypeEnum.InvFileType.D0501
                    WriteStatus(D0501FileName, Status)
                Case B07.InvType.InvTypeEnum.InvFileType.X0401
                    WriteStatus(X0401FileName, Status)
                Case B07.InvType.InvTypeEnum.InvFileType.B08
                    WriteStatus(B08Filename, Status)
            End Select

        Catch ex As Exception
            NSTV.Log.NstvLog.WriteErrorLog(ex, Nothing)
        Finally
            mtx.ReleaseMutex()
        End Try
    End Sub
    Public Sub InsertSQLTextLog(ByVal InvRunType As CableSoft.B07.InvType.InvTypeEnum.InvFileType,
                                    ByVal SQLText As String, ProcessTimeCount As Double, ProcessDate As DateTime)
        mtx.WaitOne()
        Try
            If IsCanLogSQL Then
                Dim act As New Action(Of CableSoft.B07.InvType.InvTypeEnum.InvFileType,
                                  String, Double, DateTime)(AddressOf LogSQLFile)
                act.BeginInvoke(InvRunType, SQLText, ProcessTimeCount, ProcessDate, Nothing, Nothing)

            End If

        Catch ex As Exception
            NSTV.Log.NstvLog.WriteErrorLog(ex, "記錄SQL語法")
        Finally
            mtx.ReleaseMutex()
        End Try

    End Sub
    Public Sub DeleteAllStatus()
        mtx.WaitOne()
        Try
            Dim files() As String = Directory.GetFiles(StatusDirectory, "*.Status")
            For Each Str As String In files

                File.Delete(Str)
            Next
        Catch ex As Exception
            IsCanLogSQL = False
            NSTV.Log.NstvLog.WriteErrorLog(ex, Nothing)
        Finally
            mtx.ReleaseMutex()
        End Try
    End Sub
#Region "IDisposable Support"
    Private disposedValue As Boolean ' 偵測多餘的呼叫

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: 處置 Managed 狀態 (Managed 物件)。
                If mtx IsNot Nothing Then
                    mtx.Close()
                    mtx.Dispose()
                    mtx = Nothing
                End If
                If mtxFile IsNot Nothing Then
                    mtxFile.Close()
                    mtxFile.Dispose()
                    mtxFile = Nothing
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
