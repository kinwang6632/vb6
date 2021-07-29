Imports System.IO
Imports System
Imports CableSoft.NSTV
Imports CableSoft.NSTV.SocketClient
Imports System.Net
Imports System.Net.Sockets
Public Class frmLowCmd
    Private Const Sys_Setting As String = "Sys_Setting.xml"
    Private Const ParentNodeName As String = "SMS"
    Private MsgIdlst As New Dictionary(Of String, String)
    Private Client As CableSoft.NSTV.SocketClient.Client = Nothing
    Private LowerCmd As CableSoft.NSTV.BuilderLowerCmd.LowCmd = Nothing
    Private xmlFile As CableSoft.NSTV.XMLFileIO.XmlFileIO = Nothing
    Private IsTalk As Boolean = False
    Private Sub frmLowCmd_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        If Not File.Exists(Application.StartupPath & "\" & Sys_Setting) Then
            MessageBox.Show(Sys_Setting & "檔案不存在!!", "錯誤", MessageBoxButtons.OK)
            Application.Exit()
        End If        
        Try
            xmlFile = New CableSoft.NSTV.XMLFileIO.XmlFileIO(Application.StartupPath & "\" & Sys_Setting)
            txtProto_Ver.Text = xmlFile.ReadXmlNodeValue(ParentNodeName, "Proto_Ver", False)
            txtCrypt_Ver.Text = xmlFile.ReadXmlNodeValue(ParentNodeName, "Crypt_Ver", False)
            txtKey_Type.Text = xmlFile.ReadXmlNodeValue(ParentNodeName, "Key_Type", False)
            txtOPE_ID.Text = xmlFile.ReadXmlNodeValue(ParentNodeName, "OPE_ID", False)
            txtRootKey.Text = xmlFile.ReadXmlNodeValue(ParentNodeName, "Root_Key", False)
            txtSMS_ID.Text = xmlFile.ReadXmlNodeValue(ParentNodeName, "SMS_ID", False)
            txtIPAddress.Text = xmlFile.ReadXmlNodeValue(ParentNodeName, "IPAddress", False)
            txtPort.Text = xmlFile.ReadXmlNodeValue(ParentNodeName, "Port", False)
            txtCard_SN.Text = xmlFile.ReadXmlNodeValue(ParentNodeName, "CARD_SN", False)
            txtSTBID.Text = xmlFile.ReadXmlNodeValue(ParentNodeName, "STBID", False)
            txtENTITLEEXT.Text = xmlFile.ReadXmlNodeValue(ParentNodeName, "ENTITLE", False)
            txtParent_Card_SN.Text = xmlFile.ReadXmlNodeValue(ParentNodeName, "PARENT_CARD_SN", False)
            txtFeed_Interval_Hour.Text = xmlFile.ReadXmlNodeValue(ParentNodeName, "FEED_INTERVAL_HOUR", False)
            txtNo.Text = xmlFile.ReadXmlNodeValue(ParentNodeName, "NO", False)
            txtCha.Text = xmlFile.ReadXmlNodeValue(ParentNodeName, "CHA", False)
            Dim lstAttribute As List(Of Object) = xmlFile.ReadXmlNodeAttributes(ParentNodeName & "/CmdKind", "MsgId", "Value", False)
            Dim lstNodeValue As List(Of Object) = xmlFile.ReadXmlNodeValues(ParentNodeName, "CmdKind", False)
            
            For i As Int32 = 0 To lstNodeValue.Count - 1
                MsgIdlst.Add(lstNodeValue.Item(i), lstAttribute.Item(i))
                cboMsgId.Items.Add(lstNodeValue.Item(i))
            Next
            If cboMsgId.Items.Count > 0 Then
                cboMsgId.SelectedIndex = 0
            End If
            
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
        

    End Sub

    Private Sub btnConnect_Click(sender As System.Object, e As System.EventArgs) Handles btnConnect.Click
        Try
            If (Client IsNot Nothing) AndAlso (Client.Connected) Then Exit Sub
            IsTalk = False
            Client = New CableSoft.NSTV.SocketClient.Client(Sockets.AddressFamily.InterNetwork,
                                                           Sockets.SocketType.Stream, Sockets.ProtocolType.Tcp)
            Client.Connect(txtIPAddress.Text, txtPort.Text)
            LowerCmd = New CableSoft.NSTV.BuilderLowerCmd.LowCmd(txtProto_Ver.Text,
                                                                txtCrypt_Ver.Text,
                                                                Int32.Parse(txtKey_Type.Text),
                                                                txtOPE_ID.Text,
                                                                txtSMS_ID.Text, txtRootKey.Text, Nothing, Nothing)
        Catch ex As Exception
            Try
                'Client.Disconnect(False)
                Client.Dispose()
                Client = Nothing
            Finally

            End Try
            MessageBox.Show(ex.ToString)
        End Try


    End Sub
    Private Function IsDataOK(ByVal MsgId As Int32) As Boolean
        
        Try
            If Int32.Parse(txtKey_Type.Text) = 2 Then
                IsTalk = True
            End If
            If (Not IsTalk) AndAlso (MsgId <> 1) Then
                MessageBox.Show("請先執行會話命令", "警告", MessageBoxButtons.OK)
                Return False
            End If
            Select Case MsgId
                Case 1

                Case &H201
                    If String.IsNullOrEmpty(txtCard_SN.Text) Then
                        MessageBox.Show("CARD_SN需有值!", "訊息", MessageBoxButtons.OK)
                        Return False
                    End If

                Case Else
                    Return True
            End Select
        Catch ex As Exception
            Throw
        End Try
        Return True
    End Function
    Private Sub SendCommand(ByVal MsgIdName As String)
        If (Client Is Nothing) OrElse (Not Client.Connected) Then

            MessageBox.Show("請先連線!", "警告")
            Exit Sub
        End If
        Try
            LowerCmd.Proto_Ver = txtProto_Ver.Text
            LowerCmd.Crypt_Ver = txtCrypt_Ver.Text
            LowerCmd.Key_Type = Byte.Parse(txtKey_Type.Text)
            LowerCmd.OPE_ID = Int32.Parse(txtOPE_ID.Text)
            LowerCmd.SMS_ID = Int32.Parse(txtSMS_ID.Text)
            LowerCmd.Root_Key = txtRootKey.Text

            Dim SendStruct As CableSoft.NSTV.BuilderLowerCmd.SendLowCmd = Nothing
            Dim RecvStruct As CableSoft.NSTV.BuilderLowerCmd.ResponseLowCmd = Nothing
            If Not MsgIdlst.Keys.Contains(MsgIdName) Then
                MessageBox.Show("找不到命令種類!")
            End If
            Dim value As Int32 = Int32.Parse(Convert.ToInt32(MsgIdlst.Item(MsgIdName), 16))
            If Not IsDataOK(value) Then
                Exit Sub
            End If


            Select Case value
                Case 1
                    SendStruct = LowerCmd.Builder_SMS_CA_CREATE_SESSION_REQUEST()
                    
                Case &H201
                    SendStruct = LowerCmd.Builder_SMS_CA_OPEN_ACCOUNT_REQUEST((txtCard_SN.Text & New String("0", 16)).Substring(0, 16))
                Case &H206
                   
                    SendStruct = LowerCmd.Builder_SMS_CA_REPAIR_REQUEST((txtCard_SN.Text & New String("0", 16)).Substring(0, 16), txtSTBID.Text)
                Case &H20D
                    SendStruct = LowerCmd.Builder_SMS_CA_ENTITLE_REQUEST((txtCard_SN.Text & New String("0", 16)).Substring(0, 16), txtENTITLEEXT.Text.Trim(Environment.NewLine))
                Case &H202
                    SendStruct = LowerCmd.Builder_SMS_CA_STOP_ACCOUNT_REQUEST((txtCard_SN.Text & New String("0", 16)).Substring(0, 16))
                Case &H203
                    SendStruct = LowerCmd.Builder_SMS_CA_SET_LOCK_REQUEST((txtCard_SN.Text & New String("0", 16)).Substring(0, 16))
                Case &H204
                    SendStruct = LowerCmd.Builder_SMS_CA_SET_UNLOCK_REQUEST((txtCard_SN.Text & New String("0", 16)).Substring(0, 16))                    
                Case &H402
                    SendStruct = LowerCmd.Builder_SMS_CA_SET_CHILD_REQUEST((txtCard_SN.Text & New String("0", 16)).Substring(0, 16),
                                                                           txtParent_Card_SN.Text, txtFeed_Interval_Hour.Text)
                Case &H403
                    SendStruct = LowerCmd.Builder_SMS_CA_CANCEL_CHILD_REQUEST((txtCard_SN.Text & New String("0", 16)).Substring(0, 16))
                Case &H207
                    SendStruct = LowerCmd.Builder_SMS_CA_RESETCARDPIN_REQUEST((txtCard_SN.Text & New String("0", 16)).Substring(0, 16))

                Case &H20C
                    SendStruct = LowerCmd.Builder_SMS_CA_SET_CHARACTER_REQUEST((txtCard_SN.Text & New String("0", 16)).Substring(0, 16), txtNo.Text, txtCha.Text)
                Case Else
                    MessageBox.Show("命令尚未完成!", "訊息", MessageBoxButtons.OK)
                    Exit Sub
            End Select
            Dim byteRecv() As Byte = Client.SendSynchronize(SendStruct.SendData)
            txtSendData.Text = SendStruct.SendDataString
            RecvStruct = LowerCmd.AnalyseLowCmd(byteRecv)
            txtRecvData.Text = RecvStruct.RecvDataString
            If (RecvStruct.ErrorCode) = 0 AndAlso (RecvStruct.Success) AndAlso (RecvStruct.MSG_ID = &H8001) Then
                IsTalk = True
            End If
            If Not RecvStruct.MacOK Then
                txtMsg.Text = "MAC檢驗有誤!"
                Exit Sub
            End If
            If Not RecvStruct.Success Then
                txtMsg.Text = "拆解命令失敗!"
                Exit Sub
            End If
            If RecvStruct.Success AndAlso RecvStruct.ErrorCode = 0 Then
                txtMsg.Text = "命令成功!"

            Else
                txtMsg.Text = "錯誤代碼：" & RecvStruct.ErrorCode
            End If
        Catch ex As Exception
            txtMsg.Text = ex.Message
            MessageBox.Show(ex.ToString)
            Client.Dispose()
            Client = Nothing
        End Try

    End Sub
    Private Sub btnSend_Click(sender As System.Object, e As System.EventArgs) Handles btnSend.Click
        SendCommand(cboMsgId.Text)
        Exit Sub

       
    End Sub

    Private Sub btnCheck_Click(sender As System.Object, e As System.EventArgs) Handles btnCheck.Click
        If (LowerCmd Is Nothing) OrElse (Client Is Nothing) OrElse (Not Client.Connected) Then
            MessageBox.Show("請先連線!", "訊息")
            Exit Sub
        End If
        Dim byteData() As Byte = LowerCmd.BuilderCheckState
        Dim lstData As New List(Of Byte())
        lstData.Add(byteData)
        Try

            Dim RecvData() As Byte = Client.SendSynchronize(byteData)

            If LowerCmd.ChkeckState(RecvData) Then
                txtMsg.Text = "命令成功！"
            Else
                txtMsg.Text = "命令失敗！"
            End If
        Catch ex As Exception
            txtMsg.Text = ex.ToString
            MessageBox.Show(ex.ToString)
        End Try

    End Sub

    
    


    Private Sub btnSaveSet_Click(sender As System.Object, e As System.EventArgs) Handles btnSaveSet.Click
        xmlFile.WriteXmlNodeValue(ParentNodeName, "Proto_Ver", txtProto_Ver.Text, False)
        xmlFile.WriteXmlNodeValue(ParentNodeName, "Crypt_Ver", txtCrypt_Ver.Text, False)
        xmlFile.WriteXmlNodeValue(ParentNodeName, "Key_Type", txtKey_Type.Text, False)
        xmlFile.WriteXmlNodeValue(ParentNodeName, "OPE_ID", txtOPE_ID.Text, False)
        xmlFile.WriteXmlNodeValue(ParentNodeName, "Root_Key", txtRootKey.Text, False)
        xmlFile.WriteXmlNodeValue(ParentNodeName, "SMS_ID", txtSMS_ID.Text, False)
        xmlFile.WriteXmlNodeValue(ParentNodeName, "IPAddress", txtIPAddress.Text, False)
        xmlFile.WriteXmlNodeValue(ParentNodeName, "Port", txtPort.Text, False)

        'XMLFileIO.XmlFileIO.WriteXmlNodeValue(ParentNodeName, "Proto_Ver", txtProto_Ver.Text, False)
        'XMLFileIO.XmlFileIO.WriteXmlNodeValue(ParentNodeName, "Crypt_Ver", txtCrypt_Ver.Text, False)
        'XMLFileIO.XmlFileIO.WriteXmlNodeValue(ParentNodeName, "Key_Type", txtKey_Type.Text, False)
        'XMLFileIO.XmlFileIO.WriteXmlNodeValue(ParentNodeName, "OPE_ID", txtOPE_ID.Text, False)
        'XMLFileIO.XmlFileIO.WriteXmlNodeValue(ParentNodeName, "Root_Key", txtRootKey.Text, False)
        'XMLFileIO.XmlFileIO.WriteXmlNodeValue(ParentNodeName, "SMS_ID", txtSMS_ID.Text, False)
        'XMLFileIO.XmlFileIO.WriteXmlNodeValue(ParentNodeName, "IPAddress", txtIPAddress.Text, False)
        'XMLFileIO.XmlFileIO.WriteXmlNodeValue(ParentNodeName, "Port", txtPort.Text, False)


        MessageBox.Show("儲存成功!", "訊息", MessageBoxButtons.OK)
      
    End Sub

    Private Sub btnSavePara_Click(sender As System.Object, e As System.EventArgs) Handles btnSavePara.Click
        xmlFile.WriteXmlNodeValue(ParentNodeName, "CARD_SN", txtCard_SN.Text, False)
        xmlFile.WriteXmlNodeValue(ParentNodeName, "STBID", txtSTBID.Text, False)
        xmlFile.WriteXmlNodeValue(ParentNodeName, "ENTITLE", txtENTITLEEXT.Text, False)
        xmlFile.WriteXmlNodeValue(ParentNodeName, "PARENT_CARD_SN", txtParent_Card_SN.Text, False)
        xmlFile.WriteXmlNodeValue(ParentNodeName, "FEED_INTERVAL_HOUR", txtFeed_Interval_Hour.Text, False)
        xmlFile.WriteXmlNodeValue(ParentNodeName, "NO", txtNo.Text, False)
        xmlFile.WriteXmlNodeValue(ParentNodeName, "CHA", txtCha.Text, False)


        'XMLFileIO.XmlFileIO.WriteXmlNodeValue(ParentNodeName, "CARD_SN", txtCard_SN.Text, False)
        'XMLFileIO.XmlFileIO.WriteXmlNodeValue(ParentNodeName, "STBID", txtSTBID.Text, False)
        'XMLFileIO.XmlFileIO.WriteXmlNodeValue(ParentNodeName, "ENTITLE", txtENTITLEEXT.Text, False)
        'XMLFileIO.XmlFileIO.WriteXmlNodeValue(ParentNodeName, "PARENT_CARD_SN", txtParent_Card_SN.Text, False)
        'XMLFileIO.XmlFileIO.WriteXmlNodeValue(ParentNodeName, "FEED_INTERVAL_HOUR", txtFeed_Interval_Hour.Text, False)
        'XMLFileIO.XmlFileIO.WriteXmlNodeValue(ParentNodeName, "NO", txtNo.Text, False)
        'XMLFileIO.XmlFileIO.WriteXmlNodeValue(ParentNodeName, "CHA", txtCha.Text, False)
        MessageBox.Show("儲存成功!", "訊息", MessageBoxButtons.OK)
    End Sub

    Private Sub frmLowCmd_FormClosing(sender As System.Object, e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        Try
            If Client IsNot Nothing Then
                Client.Dispose()
                Client = Nothing
            End If
        Finally

        End Try
    End Sub

    Private Sub btnDisconnect_Click(sender As System.Object, e As System.EventArgs) Handles btnDisconnect.Click
        If Client IsNot Nothing Then
            Client.Dispose()
            Client = Nothing
        End If
    End Sub
End Class
