Imports System
Imports System.IO
Imports System.Threading
Imports CableSoft.Invoice.InsertEventLog
Imports CableSoft.Invoice.Gateway.SOInfo
Imports DevExpress.XtraTreeList.Nodes
Imports CableSoft.NSTV.Log.NstvLog
Public Class MonitorFile
    Implements IDisposable

    Private InsertLog As CableSoft.Invoice.InsertEventLog.LogToTxt
    Private evn As AutoResetEvent = Nothing
    Private ti As RegisteredWaitHandle = Nothing
    Private mtxWait As Mutex = Nothing
    Private IsProcessing As Boolean
    Private IsStop As Boolean
    Private CompCode As String
    Private xfrmMain As frmMonitor    
    Public Sub New(ByVal CompCode As String, ByVal MonitorPath As String, ByVal xfrm As frmMonitor)
        Me.CompCode = CompCode
        InsertLog = New LogToTxt(CompCode, False, 0, MonitorPath)
        xfrmMain = xfrm
    End Sub
    Public Sub New(ByVal CompCode As String, ByVal xfrm As frmMonitor)
        'Me.CompCode = CompCode
        'InsertLog = New LogToTxt(CompCode, False, 0)
        Me.New(CompCode,
               Path.GetDirectoryName(AppDomain.CurrentDomain.BaseDirectory), xfrm)
    End Sub
    Public Sub StartMonitor()
        Try
            If evn IsNot Nothing Then
                evn.Dispose()
                evn = Nothing
            End If
            mtxWait = New Mutex()
            evn = New AutoResetEvent(False)
            Dim ReadTime As Double = 1
            If xfrmMain.barSpinReadTime.EditValue IsNot Nothing Then
                ReadTime = Double.Parse(xfrmMain.barSpinReadTime.EditValue)
            End If
            If ReadTime <= 0 Then
                ReadTime = 1
            End If
            ti = ThreadPool.RegisterWaitForSingleObject(evn, _
                                                                New WaitOrTimerCallback(AddressOf WaitProc), _
                                                                Nothing,
                                                                TimeSpan.FromSeconds(ReadTime),
                                                                False)
        Catch ex As Exception
            WriteErrorLog(ex, "無法啟動監控")
        End Try
       
    End Sub
    Public Sub StopMonitor()
        IsStop = True
        Dim StopTime As New Stopwatch()
        StopTime.Start()
        'If ti IsNot Nothing Then
        '    ti.Unregister(Nothing)
        '    ti = Nothing
        'End If
        Do
            If (Not IsProcessing) OrElse (StopTime.Elapsed.Seconds >= 60) Then
                If ti IsNot Nothing Then
                    ti.Unregister(Nothing)
                    ti = Nothing
                End If
                Exit Do
            End If
            Thread.Sleep(100)
        Loop

        If evn IsNot Nothing Then
            evn.Close()
            evn.Dispose()
            evn = Nothing
        End If
        If mtxWait IsNot Nothing Then
            mtxWait.Close()
            mtxWait.Dispose()
            mtxWait = Nothing
        End If
    End Sub

    Private Sub WaitProc(ByVal state As Object, ByVal timedOut As Boolean)
        mtxWait.WaitOne()
        If (IsProcessing) OrElse (IsStop) Then
            mtxWait.ReleaseMutex()
            Exit Sub
        End If
        IsProcessing = True
        Try
            Dim nod As TreeListNode = xfrmMain.trelstRunSO.FindNodeByFieldValue("CompCode", CompCode)
            UpdateNode(nod)
        Catch ex As Exception
            WriteErrorLog(ex, Nothing)
        Finally
            IsProcessing = False
            mtxWait.ReleaseMutex()
        End Try

    End Sub
    Private Sub UpdateNode(ByVal nod As TreeListNode)
        If xfrmMain.InvokeRequired Then
            Dim act As New Action(Of TreeListNode)(AddressOf UpdateNode)
            xfrmMain.Invoke(act, nod)
        Else
            nod.Item(frmMonitor.colC0401) = Convert.ToInt16(InsertLog.IsRunC0401).ToString
            nod.Item(frmMonitor.colC0501) = Convert.ToInt16(InsertLog.IsRunC0501).ToString
            nod.Item(frmMonitor.colD0401) = Convert.ToInt16(InsertLog.IsRunD0401).ToString
            nod.Item(frmMonitor.colD0501) = Convert.ToInt16(InsertLog.IsRunD0501).ToString
            nod.Item(frmMonitor.colC0701) = Convert.ToInt16(InsertLog.IsRunC0701).ToString
            nod.Item(frmMonitor.colX0401) = Convert.ToInt16(InsertLog.IsRunX0401).ToString
            nod.Item(frmMonitor.colB08) = Convert.ToInt16(InsertLog.IsRunB08).ToString
        End If
    End Sub
#Region "IDisposable Support"
    Private disposedValue As Boolean ' 偵測多餘的呼叫

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: 處置 Managed 狀態 (Managed 物件)。
                If InsertLog IsNot Nothing Then
                    InsertLog.Dispose()
                    InsertLog = Nothing
                End If
                StopMonitor()
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
