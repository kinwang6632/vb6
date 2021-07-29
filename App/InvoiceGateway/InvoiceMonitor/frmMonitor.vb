Imports System
Imports System.Threading
Imports System.IO
Imports System.Text
Imports DevExpress.XtraTreeList.Nodes
Imports DevExpress.XtraEditors

Public Class frmMonitor
    Private MonitorPath As String
    Private Const MonitorSetting As String = "MonitorPath"
    Private Const ReadTimeName As String = "ReadTime"
    Private Const ShowPathText = "監控路徑：{0}"
    Private Const GatewaySetting As String = "InvoiceSetup.Set"
    Private lstSO As List(Of CableSoft.Invoice.Gateway.SOInfo.SO)
    Private lstMonitor As List(Of MonitorFile)
    Private Enum MonitorStatus
        Initial = 0
        PlayGateway = 1
        StopGateway = 2
        AllDisabled = 3
        NoneSetting = 4
    End Enum
    Private Sub ChangeStatus(ByVal Status As MonitorStatus)
        barbtnPlay.Enabled = False
        barbtnStop.Enabled = False
        barbtnExit.Enabled = False
        barbtnSetting.Enabled = False
        barSpinReadTime.Enabled = False
        Select Case Status
            Case MonitorStatus.Initial, MonitorStatus.StopGateway
                barbtnPlay.Enabled = True
                barbtnSetting.Enabled = True
                barbtnExit.Enabled = True
                barSpinReadTime.Enabled = True
            Case MonitorStatus.PlayGateway
                barbtnStop.Enabled = True
            Case MonitorStatus.NoneSetting
                barbtnSetting.Enabled = True
                barbtnExit.Enabled = True
        End Select
        If trelstRunSO.Nodes.Count = 0 Then
            barbtnPlay.Enabled = False
            barbtnStop.Enabled = False
            barbtnExit.Enabled = False
            barbtnExit.Enabled = True
            barbtnSetting.Enabled = True
        End If
    End Sub

    Private Sub frmMonitor_FormClosing(sender As Object, e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If (lstMonitor IsNot Nothing) AndAlso (lstMonitor.Count > 0) Then
                For Each objMonitor As MonitorFile In lstMonitor
                    objMonitor.Dispose()
                    objMonitor = Nothing
                Next
                lstMonitor.Clear()
            End If
            If (lstSO IsNot Nothing) AndAlso (lstSO.Count > 0) Then
                For Each so As CableSoft.Invoice.Gateway.SOInfo.SO In lstSO
                    so.Dispose()
                    so = Nothing
                Next
                lstSO.Clear()
            End If
            My.Settings.Item(ReadTimeName) = Double.Parse(barSpinReadTime.EditValue)
            My.Settings.Save()
        Catch ex As Exception

        End Try
    End Sub
    Private Sub frmMonitor_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        MonitorPath = Path.GetDirectoryName(AppDomain.CurrentDomain.BaseDirectory)
        If String.IsNullOrEmpty(My.Settings.Item(MonitorSetting)) Then
            My.Settings.Item(MonitorSetting) = MonitorPath
        Else
            MonitorPath = My.Settings.Item(MonitorSetting).ToString
        End If
        barSpinReadTime.EditValue = Double.Parse(My.Settings(ReadTimeName))
        If Double.Parse(barSpinReadTime.EditValue) = 0 Then
            barSpinReadTime.EditValue = 1
        End If
        My.Settings.Save()
        ChangeStatus(MonitorStatus.AllDisabled)
        ShowMonitorPathText()
        BackgroundWorker1.RunWorkerAsync()
    End Sub
    Private Sub ShowMonitorPathText()
        barstcPath.Caption = String.Format(ShowPathText, MonitorPath)
    End Sub
    Private Sub StopMonitor()
        
        If lstMonitor IsNot Nothing AndAlso lstMonitor.Count > 0 Then
            For Each run As MonitorFile In lstMonitor
                run.StopMonitor()
                run.Dispose()
            Next
            lstMonitor.Clear()
        End If
        For Each nod As TreeListNode In trelstRunSO.Nodes
            nod.Item(colB08) = ""
            nod.Item(colC0401) = ""
            nod.Item(colC0501) = ""
            nod.Item(colC0701) = ""
            nod.Item(colD0401) = ""
            nod.Item(colD0501) = ""
            nod.Item(colX0401) = ""
        Next
    End Sub
    Private Sub StartMonitor()
        Try
            For Each nod As TreeListNode In trelstRunSO.Nodes
                nod.Item(colB08) = "0"
                nod.Item(colC0401) = "0"
                nod.Item(colC0501) = "0"
                nod.Item(colC0701) = "0"
                nod.Item(colD0401) = "0"
                nod.Item(colD0501) = "0"
                nod.Item(colX0401) = "0"
            Next
            If lstMonitor IsNot Nothing AndAlso lstMonitor.Count > 0 Then
                lstMonitor.Clear()
            End If
            If lstMonitor Is Nothing Then
                lstMonitor = New List(Of MonitorFile)
            End If
            For Each so As CableSoft.Invoice.Gateway.SOInfo.SO In lstSO
                lstMonitor.Add(New MonitorFile(so.CompCode, MonitorPath, Me))
            Next
            For Each starMonitor As MonitorFile In lstMonitor
                starMonitor.StartMonitor()
            Next
           
        Catch ex As Exception
            CableSoft.NSTV.Log.NstvLog.WriteErrorLog(ex, Nothing)
        End Try
    End Sub
    Private Sub barbtnPlay_ItemClick(sender As System.Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles barbtnPlay.ItemClick
        ChangeStatus(MonitorStatus.PlayGateway)
        StartMonitor()
    End Sub

    Private Sub barbtnStop_ItemClick(sender As System.Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles barbtnStop.ItemClick
        ChangeStatus(MonitorStatus.StopGateway)
        StopMonitor()
    End Sub

    Private Sub barbtnSetting_ItemClick(sender As System.Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles barbtnSetting.ItemClick
        Using folderSelect As New FolderBrowserDialog()
            folderSelect.Description = "選擇要監控的路徑"
            If folderSelect.ShowDialog = Windows.Forms.DialogResult.OK Then
                MonitorPath = folderSelect.SelectedPath
                My.Settings.Item(MonitorSetting) = MonitorPath
                My.Settings.Save()
                ShowMonitorPathText()
                BackgroundWorker1.RunWorkerAsync()
            End If
        End Using
    End Sub

    Private Sub barbtnExit_ItemClick(sender As System.Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles barbtnExit.ItemClick
        If lstSO IsNot Nothing AndAlso lstSO.Count Then
            For Each so As CableSoft.Invoice.Gateway.SOInfo.SO In lstSO
                so.Dispose()
                so = Nothing
            Next
        End If
        Application.Exit()
    End Sub
    Private Sub ClearNotes()
        If Me.InvokeRequired Then
            Dim act As New Action(AddressOf ClearNotes)
            Me.Invoke(act)
        Else
            trelstRunSO.Nodes.Clear()
        End If

    End Sub

    Private Sub BackgroundWorker1_DoWork(sender As System.Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Thread.Sleep(100)
        ClearNotes()
        If Not File.Exists(String.Format("{0}\{1}",
                            MonitorPath, GatewaySetting)) Then
            XtraMessageBox.Show("找不到任何設定檔，請重新設定路徑",
                                "警告", MessageBoxButtons.OK,
                                MessageBoxIcon.Warning)
            Exit Sub
        End If
        Try
            lstSO = CableSoft.Invoice.Gateway.SO.ReadSOInfo.CreateSOParameters(String.Format("{0}\{1}",
                                                                                        MonitorPath, GatewaySetting)).ToList
            If (lstSO IsNot Nothing) AndAlso (lstSO.Count > 0) Then
                For Each so As CableSoft.Invoice.Gateway.SOInfo.SO In lstSO
                    InitialForm(New String() {so.CompCode, so.CompName}, trelstRunSO.Nodes.Count)
                Next
            End If
        Catch ex As Exception
            XtraMessageBox.Show(ex.ToString, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted       
        ChangeStatus(MonitorStatus.Initial)
    End Sub
    Private Sub InitialForm(ByVal obj As Object, ByVal index As Integer)
        If Me.InvokeRequired Then
            Dim act As New Action(Of Object, Integer)(AddressOf InitialForm)
            Me.Invoke(act, obj, index)
        Else
            trelstRunSO.AppendNode(obj, index)
        End If
    End Sub
End Class
