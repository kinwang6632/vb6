Imports System
Imports System.IO
Imports System.Text
Imports DevExpress.XtraEditors
Imports DevExpress.XtraTreeList.Nodes
Imports DevExpress.XtraTreeList

Public Class xfrmSetting
    Private Const WM_NCLBUTTONDOWN As Int32 = &HA1S
    Private Const HTCAPTION As Int32 = 2
    Private Const IsEncryFile As Boolean = True
    Private Const SetingFileName As String = "CardLessMonitor.Set"
    Dim xmlRedear As CableSoft.CardLess.XMLFileIO.XmlFileIO = Nothing

    Private Sub grpCtl_MouseDown(sender As System.Object, e As System.Windows.Forms.MouseEventArgs) Handles grpCtl.MouseDown
        Try
            If e.Button = MouseButtons.Left Then
                Dim c As Control = DirectCast(sender, Control)
                c.Capture = False
                DefWndProc(Message.Create(ActiveForm.Handle, WM_NCLBUTTONDOWN, New IntPtr(HTCAPTION), IntPtr.Zero))
            End If
        Catch ex As Exception
            'CableSoft.Gateway.csException.WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
        End Try
    End Sub
    Private Sub btnCancel_Click_1(sender As System.Object, e As System.EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub
    Private Function ReadSO() As DataTable
        Dim tbSO As New DataTable("SO")
        tbSO.Columns.Add(New DataColumn("IsSelect", GetType(String)))
        tbSO.Columns.Add(New DataColumn("CompCode", GetType(String)))
        tbSO.Columns.Add(New DataColumn("CompName", GetType(String)))
        tbSO.Columns.Add(New DataColumn("SID", GetType(String)))
        tbSO.Columns.Add(New DataColumn("User", GetType(String)))
        tbSO.Columns.Add(New DataColumn("Password", GetType(String)))

        Try
            Dim lstDBValue As List(Of Object) = xmlRedear.ReadXmlNodeAttributes("CableSoft/SODB", "DB", "Value", IsEncryFile)
            Dim lstCompCode As List(Of Object) = xmlRedear.ReadXmlNodeAttributes("CableSoft/SODB", "DB", "CompCode", IsEncryFile)
            Dim lstCompName As List(Of Object) = xmlRedear.ReadXmlNodeValues("CableSoft/SODB", "DB", IsEncryFile)
            Dim lstSID As List(Of Object) = xmlRedear.ReadXmlNodeAttributes("CableSoft/SODB", "DB", "SID", IsEncryFile)
            Dim lstUSER As List(Of Object) = xmlRedear.ReadXmlNodeAttributes("CableSoft/SODB", "DB", "USER", IsEncryFile)
            Dim lstPassword As List(Of Object) = xmlRedear.ReadXmlNodeAttributes("CableSoft/SODB", "DB", "PASSWORD", IsEncryFile)
            For i As Int32 = 0 To lstDBValue.Count - 1
                Dim data(5) As Object
                data(0) = lstDBValue.Item(i)
                data(1) = lstCompCode.Item(i)
                data(2) = lstCompName.Item(i)
                data(3) = lstSID.Item(i)
                data(4) = lstUSER.Item(i)
                data(5) = lstPassword.Item(i)
                Dim rwNew As DataRow = tbSO.NewRow()
                rwNew.Item("IsSelect") = Int32.Parse(data(0))
                rwNew.Item("CompCode") = data(1)
                rwNew.Item("CompName") = data(2)
                rwNew.Item("SID") = data(3)
                rwNew.Item("User") = data(4)
                rwNew.Item("Password") = data(5)
                tbSO.Rows.Add(rwNew)
                tbSO.AcceptChanges()
            Next
        Catch ex As Exception
            Throw ex
        End Try
        Return tbSO
    End Function
    Private Function DataToTreList(ByVal aTree As TreeList,
                               ByVal Data() As Object, ByVal Index As Int32) As Boolean
        Try
            aTree.AppendNode(Data, Index)
        Catch ex As Exception
            Throw
        End Try
        Return True
    End Function
    Private Function ReadSystem() As DataTable
        Dim tbSystem As New DataTable
        Try
            tbSystem.Columns.Add(New DataColumn("Title", GetType(String)))
            tbSystem.Columns.Add(New DataColumn("ReadTime", GetType(Int32)))
            tbSystem.Columns.Add(New DataColumn("ProcCmdCount", GetType(Int32)))
            tbSystem.Columns.Add(New DataColumn("MaxThread", GetType(Int32)))
            tbSystem.Columns.Add(New DataColumn("DBMinPool", GetType(Int32)))
            tbSystem.Columns.Add(New DataColumn("DBMaxPool", GetType(Int32)))
            tbSystem.Columns.Add(New DataColumn("DBPoolLiveTime", GetType(Int32)))
            tbSystem.Columns.Add(New DataColumn("AutoRun", GetType(Int32)))
            tbSystem.Columns.Add(New DataColumn("ShowUIRecord", GetType(Int32)))
            Try
                Dim rwNew As DataRow = tbSystem.NewRow
                rwNew.Item("Title") = xmlRedear.ReadXmlNodeValue("CableSoft/System", "Title", IsEncryFile)
                rwNew.Item("ReadTime") = Decimal.Parse(xmlRedear.ReadXmlNodeValue("CableSoft/System", "ReadTime", IsEncryFile))
                rwNew.Item("ProcCmdCount") = Decimal.Parse(xmlRedear.ReadXmlNodeValue("CableSoft/System", "ProcCmdCount", IsEncryFile))
                rwNew.Item("MaxThread") = Decimal.Parse(xmlRedear.ReadXmlNodeValue("CableSoft/System", "MaxThread", IsEncryFile))
                rwNew.Item("DBMinPool") = Decimal.Parse(xmlRedear.ReadXmlNodeValue("CableSoft/System", "DBMinPool", IsEncryFile))
                rwNew.Item("DBMaxPool") = Decimal.Parse(xmlRedear.ReadXmlNodeValue("CableSoft/System", "DBMaxPool", IsEncryFile))
                rwNew.Item("DBPoolLiveTime") = Decimal.Parse(xmlRedear.ReadXmlNodeValue("CableSoft/System", "DBPoolLiveTime", IsEncryFile))
                rwNew.Item("DBPoolLiveTime") = Decimal.Parse(xmlRedear.ReadXmlNodeValue("CableSoft/System", "AutoRun", IsEncryFile))
                rwNew.Item("DBPoolLiveTime") = Decimal.Parse(xmlRedear.ReadXmlNodeValue("CableSoft/System", "ShowUIRecord", IsEncryFile))
                tbSystem.Rows.Add(rwNew)
                tbSystem.AcceptChanges()
            Catch ex As Exception
                Throw
            End Try
        Catch ex As Exception
            XtraMessageBox.Show(ex.ToString, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        Return tbSystem
    End Function

    Private Sub ReadData()
        Dim tbSystem As DataTable = Common.ReadSystem(IsEncryFile)
        Dim tbSO As DataTable = Common.ReadSO(IsEncryFile)
        Try
            txtTitle.Text = tbSystem.Rows(0).Item("Title")
            spinFreq.Value = Decimal.Parse(tbSystem.Rows(0).Item("ReadTime"))
            spinProc.Value = Decimal.Parse(tbSystem.Rows(0).Item("ProcCmdCount"))
            spinRcd.Value = Decimal.Parse("0" & tbSystem.Rows(0).Item("ShowUIRecord"))
            chkAutoRun.Checked = False
            If Decimal.Parse(" 0" & tbSystem.Rows(0).Item("AutoRun")) = 1 Then
                chkAutoRun.Checked = True
            End If
            spinMaxThreads.Value = Decimal.Parse(tbSystem.Rows(0).Item("MaxThread"))
            spinMinP.Value = Decimal.Parse(tbSystem.Rows(0).Item("DBMinPool"))
            spinMaxP.Value = Decimal.Parse(tbSystem.Rows(0).Item("DBMaxPool"))
            spinLifeP.Value = Decimal.Parse(tbSystem.Rows(0).Item("DBPoolLiveTime"))
            chkStartGateway.Checked = Boolean.Parse(tbSystem.Rows(0).Item("StartGateway"))
            TreLstDB.ClearNodes()
            If tbSO IsNot Nothing AndAlso tbSO.Rows.Count > 0 Then
                For i As Int32 = 0 To tbSO.Rows.Count - 1
                    DataToTreList(TreLstDB, tbSO.Rows(i).ItemArray, i)
                Next
            End If            
        Catch ex As Exception
            XtraMessageBox.Show(ex.ToString, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            tbSO.Dispose()
            tbSystem.Dispose()
        End Try
    End Sub

    Private Sub xfrmSetting_FormClosed(sender As Object, e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        If xmlRedear IsNot Nothing Then
            xmlRedear.Dispose()
            xmlRedear = Nothing
        End If
    End Sub
    Private Sub xfrmSetting_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        Try
            lblCore.Text = String.Format("( {0} 核心 )", Environment.ProcessorCount)
            xmlRedear = New CableSoft.CardLess.XMLFileIO.XmlFileIO(Application.StartupPath & "\" & SetingFileName)
            If Not File.Exists(Application.StartupPath & "\" & SetingFileName) Then
                txtTitle.Text = "Gateway Monitor"
                spinFreq.Value = 3
                spinProc.Value = 100
                spinRcd.Value = 200
                spinMaxThreads.Value = 100
                spinMinP.Value = 0
                spinMaxP.Value = 50
                spinLifeP.Value = 0
            Else
                ReadData()
            End If
            XtabCtl.SelectedTabPageIndex = 0
        Catch ex As Exception
            XtraMessageBox.Show(ex.ToString, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub
    Private Function WriteData() As Boolean
        Try
            Using mem As New MemoryStream
                Using XmlFile As New Xml.XmlTextWriter(mem, New UTF8Encoding(False))
                    XmlFile.Formatting = Xml.Formatting.Indented
                    XmlFile.Indentation = 4
                    XmlFile.WriteStartDocument()
                    XmlFile.WriteStartElement("CableSoft")  'CableSoft
                    XmlFile.WriteStartElement("System")     'System
                    XmlFile.WriteStartElement("ReadTime")
                    XmlFile.WriteValue(spinFreq.Value)
                    XmlFile.WriteEndElement() 'ReadTime
                    XmlFile.WriteStartElement("ProcCmdCount")
                    XmlFile.WriteValue(spinProc.Value)
                    XmlFile.WriteEndElement() 'ProcCmdCount
                    XmlFile.WriteStartElement("MaxThread")
                    XmlFile.WriteValue(spinMaxThreads.Value)
                    XmlFile.WriteEndElement() 'MaxThread
                    XmlFile.WriteStartElement("DBMinPool")
                    XmlFile.WriteValue(spinMinP.Value)
                    XmlFile.WriteEndElement() 'DBMinPool
                    XmlFile.WriteStartElement("DBMaxPool")
                    XmlFile.WriteValue(spinMaxP.Value)
                    XmlFile.WriteEndElement() 'DBMaxPool
                    XmlFile.WriteStartElement("DBPoolLiveTime")
                    XmlFile.WriteValue(spinLifeP.Value)
                    XmlFile.WriteEndElement() 'DBPoolLiveTime
                    XmlFile.WriteStartElement("Title")   'Title
                    XmlFile.WriteValue(txtTitle.Text)
                    XmlFile.WriteEndElement()     'Title
                    XmlFile.WriteStartElement("AutoRun")    'AutoRun
                    If chkAutoRun.Checked Then
                        XmlFile.WriteValue(1)
                    Else
                        XmlFile.WriteValue(0)
                    End If
                    XmlFile.WriteEndElement()                         'AutoRun
                    XmlFile.WriteStartElement("ShowUIRecord")                           'ShowUIRecord
                    XmlFile.WriteValue(spinRcd.Value)
                    XmlFile.WriteEndElement()                     'ShowUIRecord

                    XmlFile.WriteStartElement("StartGateway")       'StartGateway
                    If chkStartGateway.Checked Then
                        XmlFile.WriteValue(1)
                    Else
                        XmlFile.WriteValue(0)
                    End If
                    XmlFile.WriteEndElement()   'StartGateway
                    XmlFile.WriteEndElement()   'System
                    '---------------------------------------------------------------------------------------------------------------------------
                    XmlFile.WriteStartElement("SODB")   'SODB
                    For i As Int32 = 0 To TreLstDB.Nodes.Count - 1
                        Dim Node As TreeListNode = TreLstDB.Nodes(i)
                        XmlFile.WriteStartElement("DB")     'DB
                        XmlFile.WriteAttributeString("Value", Node.Item("SelectID"))
                        XmlFile.WriteAttributeString("CompCode", Node.Item("CompID"))
                        XmlFile.WriteAttributeString("SID", Node.Item("DBSid"))
                        XmlFile.WriteAttributeString("USER", Node.Item("DBUser"))
                        XmlFile.WriteAttributeString("PASSWORD", Node.Item("DBPassword"))
                        XmlFile.WriteValue(Node.Item("CompName"))
                        XmlFile.WriteEndElement()            'DB
                    Next
                    XmlFile.WriteEndElement()                 'SODB
                    '------------------------------------------------------------------------------------------------------------------------------

                    XmlFile.WriteEndElement() 'CableSoft
                    XmlFile.WriteEndDocument()
                    XmlFile.Flush()
                    XmlFile.Close()

                    xmlRedear.WriteXmlFile(Application.StartupPath & "\" & SetingFileName, mem, IsEncryFile)
                End Using
            End Using
            Return True
        Catch ex As Exception
            Throw
        End Try
        Return True
    End Function

    Private Sub btnOK_Click(sender As System.Object, e As System.EventArgs) Handles btnOK.Click

        If WriteData() Then
            Common.ProcAutoRun(chkAutoRun.Checked)
            Me.Close()
        End If
    End Sub

    Private Sub btnAddConn_Click(sender As System.Object, e As System.EventArgs) Handles btnAddConn.Click
        Try
            TreLstDB.AppendNode(Nothing, -1, -1, -1, 0).StateImageIndex = 0

        Catch ex As Exception
            XtraMessageBox.Show(ex.ToString(), "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub

    Private Sub btnRemoveConn_Click(sender As System.Object, e As System.EventArgs) Handles btnRemoveConn.Click
        Try
            TreLstDB.DeleteNode(TreLstDB.FocusedNode)
        Catch ex As Exception
            XtraMessageBox.Show(ex.ToString)
        End Try
    End Sub

    Private Sub btnTestConn_Click(sender As System.Object, e As System.EventArgs) Handles btnTestConn.Click
        Dim bduConnect As New OracleClient.OracleConnectionStringBuilder()
        bduConnect.Password = TreLstDB.FocusedNode.Item("DBPassword")
        bduConnect.UserID = TreLstDB.FocusedNode.Item("DBUser")
        bduConnect.DataSource = TreLstDB.FocusedNode.Item("DBSid")
        bduConnect.PersistSecurityInfo = True
        Using cn As New OracleClient.OracleConnection(bduConnect.ConnectionString)
            Try
                cn.Open()
                cn.Close()
                XtraMessageBox.Show("連接成功！", "訊息", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Catch ex As Exception
                XtraMessageBox.Show("連接失敗！", "警告", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Using
    End Sub

    Private Sub chkRunGateway_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkStartGateway.CheckedChanged
        If chkStartGateway.Checked Then
            If Not File.Exists(".\GatewaySetting.Set") Then
                XtraMessageBox.Show("Gateway 設定檔不存在，請檢查設定檔！", "訊息", MessageBoxButtons.OK, MessageBoxIcon.Information)
                chkStartGateway.Checked = False
            End If
        End If
    End Sub
End Class