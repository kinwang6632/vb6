Imports DevExpress.XtraEditors
Imports System.IO
Imports DevExpress.XtraTreeList.Nodes

Public Class fmOversee
    Private Const OverseeSettingFileName As String = "OverseeSetting.exe"
    Private readXmlSetting As OverseeXmlSetting.XmlSetting
    Private tbSO As DataTable = Nothing
    Private tbCommon As DataTable = Nothing
    Private monitorSO As Cablesoft.Oversee.Gateway.ExecuteSO.ExecuteSO
    Private act As Action(Of Cablesoft.Oversee.Gateway.SODB.SO.StatusOfProcess,
                              Integer, Integer, String, String, String)
    Private Enum StatusImg
        OK = 0
        Failure = 1
        Processing = 2
        StopRun = 3
    End Enum
    Private Enum checkImg
        StopALL = 0
        OK = 1
        Errormsg = 2
    End Enum
    Private Enum ButtonMode
        StopAll = 0
        Play = 1
        Initial = 2
    End Enum
    Private Sub btnPlay_ItemClick(sender As System.Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnPlay.ItemClick
        Me.Cursor = Cursors.WaitCursor
        ChangeButton(ButtonMode.Play)
        Try
            monitorSO = New Cablesoft.Oversee.Gateway.ExecuteSO.ExecuteSO
            AddHandler monitorSO.StatusChange, AddressOf procStatus
            act = AddressOf updUi
            monitorSO.ExecuteAll()
        Catch ex As Exception
            XtraMessageBox.Show(ex.ToString, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Me.Cursor = Cursors.Default
            writeLog(String.Format("{0}：啟動Gateway", Now.ToString("yyyy/MM/dd HH:mm:ss")))
        End Try


    End Sub
    Private Sub writeLog(msg As String)
        Dim fileName As String = Nothing
        fileName = Path.GetDirectoryName(System.AppDomain.CurrentDomain.BaseDirectory)
        fileName = String.Format("{0}\oversee.Log", fileName)
        Try
            Using stm As StreamWriter = New StreamWriter(fileName, True, System.Text.Encoding.UTF8)
                stm.WriteLine(msg)
                stm.Flush()
                stm.Close()
            End Using
        Catch ex As Exception

        End Try
    End Sub
    Private Sub procStatus(StatusKind As Cablesoft.Oversee.Gateway.SODB.SO.StatusOfProcess,
                              ByVal sequence As Integer,
                              ByVal comId As Integer,
                              ByVal compName As String,
                              ByVal prgName As String,
                              ByVal msg As String)

        If Me.InvokeRequired Then
            Dim objary As New List(Of Object)
            objary.Add(StatusKind)
            objary.Add(sequence)
            objary.Add(comId)
            objary.Add(compName)
            objary.Add(prgName)
            objary.Add(msg)
            Try
                Me.Invoke(act, objary.ToArray)
            Catch ex As Exception
            Finally
                objary.Clear()
                objary = Nothing
            End Try
        End If
    End Sub
    Private Sub btnStop_ItemClick(sender As System.Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnStop.ItemClick
        Me.Cursor = Cursors.WaitCursor
        ChangeButton(ButtonMode.Initial)
        writeLog(String.Format("{0}：停止Gateway", Now.ToString("yyyy/MM/dd HH:mm:ss")))
        RemoveHandler monitorSO.StatusChange, AddressOf procStatus

        Try
            Threading.Thread.Sleep(500)            
            monitorSO.StopAll()
            monitorSO.Dispose()
            monitorSO = Nothing
            For Each nd As TreeListNode In treSO.Nodes
                nd.Item(colStatus) = Integer.Parse(StatusImg.StopRun)
            Next

        Catch ex As Exception
            XtraMessageBox.Show(ex.ToString, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Me.Cursor = Cursors.Default
            ChangeButton(ButtonMode.StopAll)
            'monitorSO = Nothing
            GC.Collect()
        End Try

    End Sub

    Private Sub fmOversee_FormClosed(sender As Object, e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Try
            If readXmlSetting IsNot Nothing Then
                readXmlSetting.Dispose()
                readXmlSetting = Nothing
            End If
            If tbCommon IsNot Nothing Then
                tbCommon.Dispose()
                tbCommon = Nothing
            End If
            If tbSO IsNot Nothing Then
                tbSO.Dispose()
                tbSO = Nothing
            End If
        Catch ex As Exception
        Finally
            writeLog(String.Format("{0}：關閉Gateway", Now.ToString("yyyy/MM/dd HH:mm:ss")))
        End Try

    End Sub

    Private Sub fmOversee_FormClosing(sender As Object, e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If Not btnPlay.Enabled Then
            XtraMessageBox.Show("程式執行中！,請先停止監控", "警告視窗", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            e.Cancel = True
        End If
    End Sub

    Private Sub btnExit_ItemClick(sender As System.Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnExit.ItemClick
        'treSO.FocusedNode.Item(colErrmsg) = treld.RegisterTreeListLookUpEdit
        Me.Close()        
    End Sub

    Private Sub btnSetting_ItemClick(sender As System.Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnSetting.ItemClick

        Dim fileName As String = String.Format("{0}\{1}",
                                     Path.GetDirectoryName(System.AppDomain.CurrentDomain.BaseDirectory),
                                     OverseeSettingFileName)
        Dim isTopMost As Boolean = Me.TopMost
        Me.Cursor = Cursors.WaitCursor
        ChangeButton(ButtonMode.Initial)
        If Me.TopMost Then Me.TopMost = False
        Try
            If File.Exists(fileName) Then
                Dim p As Process = Process.Start(fileName)                
                p.WaitForExit()
                treSO.Nodes.Clear()
                Threading.Thread.Sleep(400)
                bkWork.RunWorkerAsync()
                'Me.Cursor = Cursors.Default
                'p.EnableRaisingEvents = True
                'AddHandler p.Exited, Sub(sd As Object, e1 As EventArgs)
                '                         Me.Invoke(New Action(Sub()
                '                                                  Me.Enabled = True
                '                                                  Me.Cursor = Cursors.Default
                '                                              End Sub))

                '                     End Sub

            Else
                XtraMessageBox.Show("設定執行檔不存在！", "警告", MessageBoxButtons.OK, MessageBoxIcon.Error)
                'Me.Enabled = True
                Me.Cursor = Cursors.Default
            End If
        Catch ex As Exception
            XtraMessageBox.Show(ex.ToString, "警告", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Me.Enabled = True
            Me.Cursor = Cursors.Default
        Finally
            Me.TopMost = isTopMost
            ChangeButton(ButtonMode.StopAll)
            writeLog(String.Format("{0}：設定Gateway", Now.ToString("yyyy/MM/dd HH:mm:ss")))
            'Me.Enabled = True
            'Me.Cursor = Cursors.Default

        End Try

    End Sub

    Private Sub bkWork_DoWork(sender As System.Object, e As System.ComponentModel.DoWorkEventArgs) Handles bkWork.DoWork
        System.Threading.Thread.Sleep(TimeSpan.FromSeconds(0.2))
        Me.Invoke(New Action(Sub()
                                 Me.Cursor = Cursors.WaitCursor
                                 treSO.Nodes.Clear()

                             End Sub))
        Try
            If tbSO Is Nothing Then
                tbSO = readXmlSetting.CreateSODB
            End If
            If tbCommon Is Nothing Then
                tbCommon = readXmlSetting.CreateCommon
            End If
            readXmlSetting.ReadSetting(tbSO, tbCommon)    
            If Me.InvokeRequired Then
                Me.Invoke(New Action(Sub()
                                         Try
                                             For Each rw As DataRow In tbSO.Rows
                                                 If Integer.Parse(rw.Item("SelectId")) = 1 Then
                                                     Dim nd As TreeListNode = treSO.AppendNode(Nothing, -1, -1, -1, 0)
                                                     'treSO.BeginUpdate()
                                                     Try                                      
                                                         nd.Item(colSequence) = rw.Item("sequence")
                                                         nd.Item(colCompId) = rw.Item("CompCode")
                                                         nd.Item(colCompName) = rw.Item("CompName")
                                                         nd.Item(colprgName) = rw.Item("prgName")
                                                         nd.Item(colStatus) = Integer.Parse(StatusImg.StopRun)
                                                         'Dim obj As New Repository.RepositoryItem
                                                         Dim tbTest As New DataTable()
                                                         tbTest.Columns.Add("MsgTime", GetType(String))
                                                         tbTest.Columns.Add("msg", GetType(String))
                                                         Dim rwtest As DataRow = tbTest.NewRow
                                                         rwtest.Item(0) = Now.ToString("yyyy/MM/dd HH:mm:ss")
                                                         rwtest.Item(1) = "TEST"
                                                         tbTest.Rows.Add(rwtest)
                                                         treld.AutoExpandAllNodes = True
                                                         treSO.FocusedNode.StateImageIndex = Integer.Parse(checkImg.StopALL)
                                                     
                                                         treSO.FocusedNode.StateImageIndex = 0
                                                         treSO.Refresh()
                                                     Catch ex As Exception
                                                         XtraMessageBox.Show(ex.ToString, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                                     Finally
                                                         'treSO.EndUpdate()
                                                     End Try
                                                 End If
                                                 Threading.Thread.Sleep(TimeSpan.FromSeconds(0.2))
                                             Next
                                         Finally
                                             ChangeButton(ButtonMode.StopAll)
                                             Me.Cursor = Cursors.Default
                                         End Try

                                     End Sub))
            End If

        Catch ex As Exception
            Me.Cursor = Cursors.Default
            XtraMessageBox.Show(ex.ToString, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub fmOversee_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        'TODO: 這行程式碼會將資料載入 'DataSet1.CD001' 資料表。您可以視需要進行移動或移除。

        ChangeButton(ButtonMode.Initial)
        If readXmlSetting Is Nothing Then
            readXmlSetting = New OverseeXmlSetting.XmlSetting
        End If
    End Sub
    Private Sub updUi()
        'For Each rw As DataRow In tbSO.Rows
        '    Dim nd As TreeListNode = treSO.AppendNode(Nothing, -1, -1, -1, 0)
        '    nd.StateImageIndex = 0
        '    nd.Item(colSeq) = rw.Item("sequence")
        'Next
    End Sub
    Private Sub ChangeButton(ByVal Mode As ButtonMode)
        bmMain.BeginUpdate()
        Try
            Select Case Mode
                Case ButtonMode.Initial
                    btnPlay.Enabled = False
                    btnExit.Enabled = False
                    btnRefresh.Enabled = False
                    btnSetting.Enabled = False
                    btnStop.Enabled = False
                    btnTopMost.Enabled = False
                Case ButtonMode.StopAll
                    btnStop.Enabled = False
                    btnPlay.Enabled = True
                    btnExit.Enabled = True
                    btnSetting.Enabled = True
                    btnRefresh.Enabled = True
                    btnTopMost.Enabled = True
                Case ButtonMode.Play
                    btnPlay.Enabled = False
                    btnExit.Enabled = False
                    btnStop.Enabled = True
                    btnSetting.Enabled = False
                    btnRefresh.Enabled = False
                    btnTopMost.Enabled = True
            End Select
            For Each nd As TreeListNode In treSO.Nodes
                If nd.StateImageIndex < Integer.Parse(checkImg.Errormsg) Then
                    nd.StateImageIndex = Integer.Parse(checkImg.StopALL)
                End If
            Next
        Finally
            bmMain.EndUpdate()
        End Try

    End Sub

    Private Sub fmOversee_Shown(sender As Object, e As System.EventArgs) Handles Me.Shown
        Me.Cursor = Cursors.WaitCursor
        bkWork.RunWorkerAsync()
    End Sub

    Private Sub btnOpenlog_Click(sender As System.Object, e As System.EventArgs) Handles btnOpenlog.Click
        Try
            Dim pathName As String = Path.GetDirectoryName(System.AppDomain.CurrentDomain.BaseDirectory)
            pathName = String.Format("{0}\{1}\{2}", pathName,
                                                        treSO.FocusedNode.Item(colCompId),
                                                        treSO.FocusedNode.Item(colprgName))
            If Directory.Exists(pathName) Then
                System.Diagnostics.Process.Start("explorer.exe", pathName)
            Else
                XtraMessageBox.Show(String.Format("{0} 資料夾不存在！", pathName),
                                            "訊息", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
            If btnStop.Enabled Then
                Me.Invoke(New Action(Sub()
                                         treSO.FocusedNode.StateImageIndex = Integer.Parse(checkImg.OK)
                                     End Sub))
            Else
                Me.Invoke(New Action(Sub()
                                         treSO.FocusedNode.StateImageIndex = Integer.Parse(checkImg.StopALL)
                                     End Sub))
            End If
            'Debug.Print(pathName)
        Catch ex As Exception
            XtraMessageBox.Show(ex.ToString, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally

        End Try
        'treSO.FocusedNode.Item(colCompId)
    End Sub
    Private Sub updUi(StatusKind As Cablesoft.Oversee.Gateway.SODB.SO.StatusOfProcess,
                              ByVal sequence As Integer,
                              ByVal comId As Integer,
                              ByVal compName As String,
                              ByVal prgName As String,
                              ByVal msg As String)
        Try
            For Each nd As TreeListNode In treSO.Nodes
                If Integer.Parse(nd.Item("Sequence")) = sequence Then
                    Select Case StatusKind
                        Case Cablesoft.Oversee.Gateway.SODB.SO.StatusOfProcess.Failure
                            nd.Item(colStatus) = Integer.Parse(StatusImg.Failure)
                            nd.StateImageIndex = Integer.Parse(checkImg.Errormsg)
                        Case Cablesoft.Oversee.Gateway.SODB.SO.StatusOfProcess.Success
                            nd.Item(colStatus) = Integer.Parse(StatusImg.OK)
                            If nd.StateImageIndex <> Integer.Parse(checkImg.Errormsg) Then
                                nd.StateImageIndex = Integer.Parse(checkImg.OK)
                            End If
                        Case Cablesoft.Oversee.Gateway.SODB.SO.StatusOfProcess.Processing
                            nd.Item(colStatus) = Integer.Parse(StatusImg.Processing)
                            If nd.StateImageIndex <> Integer.Parse(checkImg.Errormsg) Then
                                nd.StateImageIndex = Integer.Parse(checkImg.OK)
                            End If
                    End Select
                    nd.Item(colErrmsg) = msg
                End If
            Next
        Catch ex As Exception
        Finally
            compName = Nothing
            prgName = Nothing
            msg = Nothing
        End Try
    End Sub

    Private Sub btnTopMost_ItemClick(sender As System.Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnTopMost.ItemClick
        Me.TopMost = Not Me.TopMost

        If Me.TopMost Then
            btnTopMost.SuperTip.Items.Clear()
            btnTopMost.SuperTip.Items.AddTitle("取消最上層顯示")
            btnTopMost.SuperTip.Items.Add("取消一直顯示在最上層")
            btnTopMost.AllowGlyphSkinning = DevExpress.Utils.DefaultBoolean.True
        Else
            btnTopMost.SuperTip.Items.Clear()
            btnTopMost.SuperTip.Items.AddTitle("最上層顯示")
            btnTopMost.SuperTip.Items.Add("一直顯示在最上層")
            btnTopMost.AllowGlyphSkinning = DevExpress.Utils.DefaultBoolean.False
        End If
        
    End Sub
End Class
