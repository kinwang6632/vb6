Imports DevExpress.XtraTreeList.Nodes
Imports DevExpress.XtraEditors
Imports System.Threading
Imports System.Text

Public Class frmOverseeSetting
    Private tbSO As DataTable
    Private tbCommon As DataTable
    Private OverseeSetting As New OverseeXmlSetting.XmlSetting
    Private Sub btnAdd_Click(sender As System.Object, e As System.EventArgs) Handles btnAdd.Click
        Try
            Dim nod As TreeListNode = TreLstDB.AppendNode(Nothing, -1, -1, -1, 0)
            nod.StateImageIndex = 0
            nod.Item(colSelSO) = 0
            nod.Item(colSid) = String.Empty

        Catch ex As Exception
            XtraMessageBox.Show(ex.ToString(), "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnSave_Click(sender As System.Object, e As System.EventArgs) Handles btnSave.Click
        tbCommon.Rows.Clear()
        tbSO.Rows.Clear()
        Try
            For Each nod As DevExpress.XtraTreeList.Nodes.TreeListNode In TreLstDB.Nodes
                Dim rwNew As DataRow = tbSO.NewRow
                rwNew.Item("SelectID") = nod.Item(colSelSO)
                rwNew.Item("CompCode") = nod.Item(colCompCode)
                rwNew.Item("CompName") = nod.Item(colCompName)
                rwNew.Item("prgName") = nod.Item(colPrgName)
                rwNew.Item("Sid") = nod.Item(colSid)
                rwNew.Item("DBUser") = nod.Item(colDBUserName)
                rwNew.Item("DBPassword") = nod.Item(colDBPassword)
                If nod.Item(colDBMinPool) IsNot Nothing Then
                    rwNew.Item("DBMinPool") = nod.Item(colDBMinPool)
                End If
                If nod.Item(colDBMaxPool) IsNot Nothing Then
                    rwNew.Item("DBMaxPool") = nod.Item(colDBMaxPool)
                End If
                If nod.Item(colDBPoolLifetime) IsNot Nothing Then
                    rwNew.Item("DBLifetime") = nod.Item(colDBPoolLifetime)
                End If
                Select Case nod.Item(colMonitorType).ToString.Replace(" ", "").ToUpper
                    Case "WebService".ToUpper
                        rwNew.Item("MonitorType") = 0
                    Case "WebPost".ToUpper
                        rwNew.Item("MonitorType") = 1
                    Case "WebGet".ToUpper
                        rwNew.Item("MonitorType") = 2
                    Case Else
                        rwNew.Item("MonitorType") = 3
                End Select
                'rwNew.Item("MonitorType")
                'TreLstDB.RepositoryItems("cmbMonitorType")
                If nod.Item(colRunTimer) IsNot Nothing Then
                    rwNew.Item("Runtimer") = nod.Item(colRunTimer)
                End If
                If nod.Item(colTimeout) IsNot Nothing Then
                    rwNew.Item("Timeout") = nod.Item(colTimeout)
                End If
                rwNew.Item("SeqSql") = nod.Item(colSeqSQL)
                rwNew.Item("InsSql") = nod.Item(ColInsSql)
                rwNew.Item("GetSql") = nod.Item(colGetSql)
                rwNew.Item("WebServiceName") = nod.Item(colWebServiceName)
                rwNew.Item("WebSrvMethod") = nod.Item(colWebSrvMethod)
                rwNew.Item("WebServicePara") = nod.Item(colWebServicePara)
                rwNew.Item("url") = nod.Item(colUrl)
                rwNew.Item("WebGetPara") = nod.Item(colWebGetPara)
                rwNew.Item("WebPostPara") = nod.Item(colWebPostPara)
                tbSO.Rows.Add(rwNew)
                tbSO.AcceptChanges()
            Next
            Dim rw As DataRow = tbCommon.NewRow
            rw("MaxThread") = spinMaxThreads.Value
            rw("ErrHistorical") = spinLogHistorical.Value
            tbCommon.Rows.Add(rw)
            tbCommon.AcceptChanges()
            OverseeSetting.WriteSetting(tbSO, tbCommon)
            OverseeSetting.ReadSetting(tbSO, tbCommon)
        Catch ex As Exception
            XtraMessageBox.Show(ex.ToString)
        End Try
        XtraMessageBox.Show("儲檔成功！", "訊息", MessageBoxButtons.OK, MessageBoxIcon.Information)

    End Sub
    Private Function CreateSODBCol() As Boolean
        tbSO.Columns.Add("SelectID", GetType(Integer))
        tbSO.Columns.Add("CompCode", GetType(String))
        tbSO.Columns.Add("CompName", GetType(String))
        tbSO.Columns.Add("prgName", GetType(String))
        tbSO.Columns.Add("Sid", GetType(String))
        tbSO.Columns.Add("DBUser", GetType(String))
        tbSO.Columns.Add("DBPassword", GetType(String))
        tbSO.Columns.Add("DBMinPool", GetType(Integer))
        tbSO.Columns.Add("DBMaxPool", GetType(Integer))
        tbSO.Columns.Add("DBLifetime", GetType(Integer))
        tbSO.Columns.Add("MonitorType", GetType(String)) '0.Web Service  1.Post  2.Get  3.DB
        tbSO.Columns.Add("Runtimer", GetType(Integer))
        tbSO.Columns.Add("Timeout", GetType(Integer))
        tbSO.Columns.Add("SeqSql", GetType(String))
        tbSO.Columns.Add("InsSql", GetType(String))
        tbSO.Columns.Add("GetSql", GetType(String))
        tbSO.Columns.Add("WebServiceName", GetType(String))
        tbSO.Columns.Add("WebSrvMethod", GetType(String))
        tbSO.Columns.Add("WebServicePara", GetType(String))
        tbSO.Columns.Add("url", GetType(String))
        tbSO.Columns.Add("WebGetPara", GetType(String))
        tbSO.Columns.Add("WebPostPara", GetType(String))
        tbCommon.Columns.Add("MaxThread", GetType(Integer))
        tbCommon.Columns.Add("ErrHistorical", GetType(Integer))
        Return True
    End Function
    Private Sub DataToScr()
        Try
            TreLstDB.Nodes.Clear()
            For Each rw As DataRow In tbSO.Rows
                Dim nod As TreeListNode = TreLstDB.AppendNode(Nothing, -1, -1, -1, 0)
                nod.Item(colSelSO) = rw.Item("SelectID")
                nod.Item(colCompCode) = rw.Item("CompCode")
                nod.Item(colCompName) = rw.Item("CompName")
                nod.Item(colPrgName) = rw.Item("prgName")
                nod.Item(colSid) = rw.Item("Sid")
                nod.Item(colDBUserName) = rw.Item("DBUser")
                nod.Item(colDBPassword) = rw.Item("DBPassword")
                nod.Item(colDBMinPool) = rw.Item("DBMinPool")
                nod.Item(colDBMaxPool) = rw.Item("DBMaxPool")
                nod.Item(colDBPoolLifetime) = rw.Item("DBLifetime")
                Select Case Integer.Parse(rw.Item("MonitorType") & "")
                    Case 0
                        nod.Item(colMonitorType) = "Web Service"
                    Case 1
                        nod.Item(colMonitorType) = "Web Post"
                    Case 2
                        nod.Item(colMonitorType) = "Web Get"
                    Case Else
                        nod.Item(colMonitorType) = "DB Access"
                End Select
                'nod.Item(colMonitorType) = rw.Item("MonitorType")
                nod.Item(colRunTimer) = rw.Item("Runtimer")
                nod.Item(colTimeout) = rw.Item("Timeout")
                nod.Item(colSeqSQL) = rw.Item("SeqSql")
                nod.Item(ColInsSql) = rw.Item("InsSql")
                nod.Item(colGetSql) = rw.Item("GetSql")
                nod.Item(colWebServiceName) = rw.Item("WebServiceName")
                nod.Item(colWebSrvMethod) = rw.Item("WebSrvMethod")
                nod.Item(colWebServicePara) = rw.Item("WebServicePara")
                nod.Item(colUrl) = rw.Item("url")
                nod.Item(colWebGetPara) = rw.Item("WebGetPara")
                nod.Item(colWebPostPara) = rw.Item("WebPostPara")
            Next
            spinMaxThreads.Value = Decimal.Parse(tbCommon.Rows(0).Item("MaxThread"))
            spinLogHistorical.Value = Decimal.Parse(tbCommon.Rows(0).Item("ErrHistorical"))
        Catch ex As Exception
            XtraMessageBox.Show(ex.ToString, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub frmOverseeSetting_FormClosed(sender As Object, e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Try
            If OverseeSetting IsNot Nothing Then
                OverseeSetting.Dispose()
                OverseeSetting = Nothing
            End If
            If tbCommon IsNot Nothing Then
                tbCommon.Dispose()
                tbCommon = Nothing
            End If
            If tbSO IsNot Nothing Then
                tbSO.Dispose()
                tbCommon = Nothing
            End If
        Catch ex As Exception

        End Try
    End Sub
    Private Sub frmOverseeSetting_Load(sender As Object, e As System.EventArgs) Handles Me.Load

        tbSO = OverseeSetting.CreateSODB
        tbCommon = OverseeSetting.CreateCommon
        If Not OverseeSetting.ReadSetting(tbSO, tbCommon) Then Exit Sub
        DataToScr()
    End Sub

    Private Sub btnDel_Click(sender As System.Object, e As System.EventArgs) Handles btnDel.Click
        Try
            TreLstDB.DeleteNode(TreLstDB.FocusedNode)
        Catch ex As Exception
            XtraMessageBox.Show(ex.ToString, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub TreLstDB_FocusedNodeChanged(sender As System.Object, e As DevExpress.XtraTreeList.FocusedNodeChangedEventArgs)

    End Sub
    Private Function ExecuteWeb(ByVal requestMethod As String, url As String, webPara As String) As Object
        Dim req As New Threading.ThreadLocal(Of System.Net.HttpWebRequest)
        Dim response As New Threading.ThreadLocal(Of System.Net.WebResponse)
        Dim result As New ThreadLocal(Of Object)
        Try
            Try
                Select Case requestMethod.ToUpper
                    Case System.Net.WebRequestMethods.Http.Post
                        req.Value = System.Net.HttpWebRequest.Create(url)
                        req.Value.Method = System.Net.WebRequestMethods.Http.Post
                        req.Value.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"
                        req.Value.ContentType = "application/x-www-form-urlencoded"
                        req.Value.ContentLength = System.Text.Encoding.UTF8.GetByteCount(webPara)
                        req.Value.KeepAlive = False

                        Using requestWriter As System.IO.StreamWriter =
                                New System.IO.StreamWriter(req.Value.GetRequestStream)
                            Try
                                requestWriter.Write(webPara, 0,
                                                    System.Text.Encoding.UTF8.GetByteCount(webPara))
                                requestWriter.Flush()
                                requestWriter.Close()
                            Catch exWriter As Exception
                                Throw exWriter
                            End Try
                        End Using
                    Case System.Net.WebRequestMethods.Http.Get
                        If String.IsNullOrEmpty(webPara) Then
                            req.Value = System.Net.HttpWebRequest.Create(url)
                        Else
                            req.Value = System.Net.HttpWebRequest.Create(String.Format("{0}?{1}", url, webPara))
                        End If
                        req.Value.Method = System.Net.WebRequestMethods.Http.Get
                        req.Value.ContentType = "application/x-www-form-urlencoded"

                        'req.Value.UserAgent = "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.2; .NET CLR 1.0.3705;)"
                        req.Value.KeepAlive = False
                End Select


                response.Value = req.Value.GetResponse
                If CType(response.Value, System.Net.HttpWebResponse).StatusCode <> Net.HttpStatusCode.OK Then
                    Throw New Exception("Response is unsuccessful.")
                Else
                    Using responseReader As System.IO.StreamReader =
                                        New System.IO.StreamReader(response.Value.GetResponseStream, New UTF8Encoding)
                        result.Value = responseReader.ReadToEnd
                        responseReader.Close()
                        responseReader.Dispose()
                        Return result.Value
                    End Using
                End If

            Catch webEx As System.Net.WebException
                XtraMessageBox.Show(webEx.ToString, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Catch ex As Exception
                XtraMessageBox.Show(ex.ToString, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Finally

            End Try

        Catch ex As Exception
            XtraMessageBox.Show(ex.ToString, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If req IsNot Nothing Then
                req.Value = Nothing
                req.Dispose()
                req = Nothing
            End If
            If response IsNot Nothing Then
                response.Value = Nothing
                response.Dispose()
                response = Nothing
            End If
            If result IsNot Nothing Then
                result.Value = Nothing
                result.Dispose()
                result = Nothing
            End If
        End Try

    End Function
    Private Sub btnTest_Click(sender As System.Object, e As System.EventArgs) Handles btnTest.Click
        Me.Cursor = Cursors.WaitCursor
        Try
            If (TreLstDB.FocusedNode.Item(colMonitorType) IsNot Nothing) AndAlso
          (Not String.IsNullOrEmpty(TreLstDB.FocusedNode.Item(colMonitorType))) Then
                Select Case TreLstDB.FocusedNode.Item(colMonitorType).ToString.Replace(" ", "").ToUpper
                    Case "DBACCESS".ToUpper
                        Dim cnbuilder As New OracleClient.OracleConnectionStringBuilder
                        cnbuilder.DataSource = TreLstDB.FocusedNode.Item(colSid)
                        cnbuilder.Password = TreLstDB.FocusedNode.Item(colDBPassword)
                        cnbuilder.UserID = TreLstDB.FocusedNode.Item(colDBUserName)
                        Try
                            Using cn As New OracleClient.OracleConnection(cnbuilder.ToString)
                                cn.Open()
                            End Using
                            XtraMessageBox.Show("連接成功", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        Catch ex As Exception
                            XtraMessageBox.Show(ex.ToString, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End Try
                    Case "WEBService".ToUpper
                        Dim url As String = TreLstDB.FocusedNode(colUrl)
                        Dim timeout As Integer = 0
                        Dim webServiceName As String = Nothing
                        Dim servicePara As New List(Of Object)
                        Try
                            If TreLstDB.FocusedNode(colTimeout) IsNot Nothing Then
                                timeout = Integer.Parse(TreLstDB.FocusedNode(colTimeout)) * 1000
                            End If
                            If TreLstDB.FocusedNode(colWebServiceName) IsNot Nothing Then
                                webServiceName = TreLstDB.FocusedNode(colWebServiceName)
                            End If
                            If TreLstDB.FocusedNode(colWebServicePara) IsNot Nothing Then
                                For Each ob As Object In TreLstDB.FocusedNode(colWebServicePara).ToString.Split(",")
                                    servicePara.Add(ob)
                                Next
                            End If
                            Using callWebService As New CableSoft.DynCallWebService.WebServiceObject(
                                             url, timeout, webServiceName)
                                Dim result As Object = callWebService.Execute(TreLstDB.FocusedNode(colWebSrvMethod),
                                                      servicePara.ToArray)
                                If result = True Then
                                    XtraMessageBox.Show("Web Service 測試成功", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information)
                                Else
                                    XtraMessageBox.Show("Web Service 回傳不正確", "失敗", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                                End If
                                callWebService.Dispose()
                            End Using
                        Catch ex As Exception
                            XtraMessageBox.Show(ex.ToString, "失敗", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End Try
                    Case "WebPost".ToUpper
                        Try
                            Dim result = ExecuteWeb(System.Net.WebRequestMethods.Http.Post)
                            If result = True Then
                                XtraMessageBox.Show("Web Post 測試成功", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            Else
                                XtraMessageBox.Show("Web Post 回傳不正確", "失敗", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                            End If
                        Catch ex As Exception
                            XtraMessageBox.Show(ex.ToString, "失敗", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End Try
                    Case "WebGet".ToUpper
                        Try
                            Dim result = ExecuteWeb(System.Net.WebRequestMethods.Http.Get)
                            If result = True Then
                                XtraMessageBox.Show("Web Get 測試成功", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            Else
                                XtraMessageBox.Show("Web Get 回傳不正確", "失敗", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                            End If
                        Catch ex As Exception
                            XtraMessageBox.Show(ex.ToString, "失敗", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End Try
                End Select
            Else
                XtraMessageBox.Show("未設定監控方式！", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
        Finally
            Me.Cursor = Cursors.Default

        End Try

    End Sub

    Private Function ExecuteWeb(ByVal requestMethod As String) As Object
        Dim req As System.Net.HttpWebRequest = Nothing
        Dim response As System.Net.WebResponse = Nothing
        Dim result As Object = Nothing
        Dim timeout As Integer = 0        
        Try
            Try
                If TreLstDB.FocusedNode(colTimeout) IsNot Nothing Then
                    timeout = Integer.Parse(TreLstDB.FocusedNode(colTimeout))
                End If
                Select Case requestMethod.ToUpper
                    Case System.Net.WebRequestMethods.Http.Post
                        req = System.Net.HttpWebRequest.Create(TreLstDB.FocusedNode(colUrl))
                        req.Method = System.Net.WebRequestMethods.Http.Post
                        req.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"
                        req.ContentType = "application/x-www-form-urlencoded"
                        req.ContentLength = System.Text.Encoding.UTF8.GetByteCount(TreLstDB.FocusedNode(colWebPostPara))
                        req.KeepAlive = False
                        If timeout > 0 Then req.Timeout = timeout * 1000
                        If timeout > 0 Then req.ReadWriteTimeout = timeout * 1000
                        Using requestWriter As System.IO.StreamWriter =
                                New System.IO.StreamWriter(req.GetRequestStream)
                            Try
                                requestWriter.Write(TreLstDB.FocusedNode(colWebPostPara), 0,
                                                    System.Text.Encoding.UTF8.GetByteCount(TreLstDB.FocusedNode(colWebPostPara)))
                                requestWriter.Flush()
                                requestWriter.Close()
                            Catch exWriter As Exception
                                Throw exWriter
                            End Try
                        End Using
                    Case System.Net.WebRequestMethods.Http.Get
                        If (TreLstDB.FocusedNode(colWebGetPara) Is Nothing) OrElse
                            (String.IsNullOrEmpty(TreLstDB.FocusedNode(colWebGetPara))) Then
                            req = System.Net.HttpWebRequest.Create(TreLstDB.FocusedNode(colUrl))
                        Else
                            req = System.Net.HttpWebRequest.Create(String.Format("{0}?{1}", TreLstDB.FocusedNode(colUrl),
                                                                                 TreLstDB.FocusedNode(colWebGetPara)))
                        End If
                        req.Method = System.Net.WebRequestMethods.Http.Get
                        req.ContentType = "application/x-www-form-urlencoded"
                        If timeout > 0 Then req.Timeout = timeout * 1000

                        req.KeepAlive = False
                End Select
                response = req.GetResponse
                If CType(response, System.Net.HttpWebResponse).StatusCode <> Net.HttpStatusCode.OK Then
                    Throw New Exception("Response is unsuccessful.")
                Else
                    Using responseReader As System.IO.StreamReader =
                                        New System.IO.StreamReader(response.GetResponseStream, New UTF8Encoding)
                        result = responseReader.ReadToEnd
                        result = result.ToString.StartsWith("True", StringComparison.OrdinalIgnoreCase)
                        responseReader.Close()
                        responseReader.Dispose()
                        Return result
                    End Using
                End If

            Catch webEx As System.Net.WebException

                Throw New Exception(webEx.ToString)
                Return False
            Catch ex As Exception
                Throw New Exception(ex.ToString)
                Return False
            Finally
            End Try

        Catch ex As Exception
        Finally
            If req IsNot Nothing Then                
                req = Nothing
            End If
            If response IsNot Nothing Then              
                response = Nothing
            End If
            If result IsNot Nothing Then       
                result = Nothing
            End If
        End Try
    End Function
  
    Private Sub btnExit_Click(sender As System.Object, e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

 
End Class