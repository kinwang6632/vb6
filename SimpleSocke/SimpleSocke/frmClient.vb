Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.IO
Imports System.Text
Imports System.Net
Imports System.Net.Sockets
Imports Cablesoft.CAS.CryptUtil
Imports Cablesoft.Gateway.Common
Imports Cablesoft.Gateway.csException
Imports CableSoft_KeyPro
Imports CableSoft_Nagra_BuilderCmd
Public Class frmClient
    Private intRcvTimeOut As Int32 = 30
    Private intSndTimeOut As Int32 = 30

    '驗證是否為IP格式
    Private Function IsValidIP(ByVal addr As String) As Boolean
        Try
            Dim strPattern As String = "^([1-9]|[1-9][0-9]|1[0-9][0-9]|2[0-4][0-9]|25[0-5])(\.([0-9]|[1-9][0-9]|1[0-9][0-9]|2[0-4][0-9]|25[0-5])){3}$"
            Dim chkRegex As Regex = New Regex(strPattern)
            Dim blnValid As Boolean = False
            If String.IsNullOrEmpty(addr.Trim()) Then
                blnValid = False

            Else
                blnValid = chkRegex.IsMatch(addr, 0)
            End If
            Return blnValid
        Catch ex As Exception
            Return False
        End Try

    End Function


    'Private Function IsDataOK() As Boolean
    '    Try
    '        If numPort.Value <= 0 Then
    '            ErrorProvider1.SetError(numPort, "Port 設定不正確")
    '            Return False
    '        End If
    '        If String.IsNullOrEmpty(txtSndPara.Text) Then
    '            ErrorProvider1.SetError(txtSndPara, "傳送參數未設定")
    '            Return False
    '        End If
    '        Return True
    '    Catch ex As Exception
    '        Return False
    '    End Try
    'End Function
    'Private Sub frmClient_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    '    On Error Resume Next
    '    If Not String.IsNullOrEmpty(My.Settings.IPAddress) Then
    '        txtIP.Text = My.Settings.IPAddress
    '    End If
    '    If Not String.IsNullOrEmpty(My.Settings.Port) Then
    '        numPort.Value = Convert.ToDecimal(My.Settings.Port)
    '    End If
    '    If Not String.IsNullOrEmpty(My.Settings.SendTimeOut) Then
    '        numSendTimeOut.Value = Convert.ToInt32(My.Settings.SendTimeOut)
    '    End If
    '    If Not String.IsNullOrEmpty(My.Settings.RcvTimeOut) Then
    '        numRcvTimeOut.Value = Convert.ToInt32(My.Settings.RcvTimeOut)
    '    End If
    'End Sub

    'Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    Try
    '        ErrorProvider1.Clear()
    '        If Not IsValidIP(txtIP.Text) Then
    '            ErrorProvider1.SetError(txtIP, "IP不正確")
    '            Exit Sub
    '        End If
    '        If Not IsDataOK() Then
    '            Exit Sub
    '        End If
    '        My.Settings.IPAddress = txtIP.Text
    '        My.Settings.Port = numPort.Value.ToString
    '        My.Settings.RcvTimeOut = numRcvTimeOut.Value.ToString
    '        My.Settings.SendTimeOut = numSendTimeOut.Value.ToString
    '        btnOK.Enabled = False
    '        txtResult.Text = String.Empty
    '        Dim thr As New Thread(AddressOf SendSocket)
    '        thr.IsBackground = True
    '        thr.Start()
    '    Catch ex As Exception
    '        MessageBox.Show(ex.Message, "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '    End Try



    'End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        On Error Resume Next
        ErrorProvider1.Clear()
    End Sub
    'Private Sub SendSocket()
    '    Dim oClient As New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)
    '    Try

    '        'oClient.SendTimeout = intTimeOut * 1000
    '        'oClient.ReceiveTimeout = intTimeOut * 1000
    '        oClient.SendTimeout = intSndTimeOut * 1000
    '        oClient.ReceiveTimeout = intRcvTimeOut * 1000
    '        Dim aSendMsg() As Byte = Encoding.Default.GetBytes(txtSndPara.Text)
    '        oClient.Connect(IPAddress.Parse(txtIP.Text), Convert.ToInt32(numPort.Value))
    '        oClient.Send(aSendMsg)

    '        Dim aMsg As String = String.Empty
    '        Dim strMsg As String = String.Empty
    '        Dim iTotal As Int32 = 0

    '        'Dim stm As New NetworkStream(oClient)
    '        'Do While True
    '        '    Dim aRcvMsg(1023) As Byte
    '        '    stm.Read(aRcvMsg, 0, aRcvMsg.Length)
    '        '    aMsg = aMsg.TrimEnd(Chr(0)) & Encoding.Default.GetChars(aRcvMsg)
    '        '    'lstMsg.AddRange(aRcvMsg)

    '        '    If aMsg.IndexOf("#") > 0 Then
    '        '        iTotal = Int32.Parse(aMsg.Substring(0, aMsg.IndexOf("#")))
    '        '    End If
    '        '    If Encoding.Default.GetByteCount(aMsg) >= iTotal Then
    '        '        Exit Do
    '        '    End If
    '        'Loop
    '        'strMsg = aMsg
    '        'strMsg = Encoding.Default.GetString(lstMsg.ToArray)
    '        Dim aRcvMsg(4095) As Byte
    '        If oClient.Receive(aRcvMsg) > 1 Then
    '            aMsg = Encoding.Default.GetString(aRcvMsg).TrimEnd(Chr(0))


    '        End If

    '        'Do While True
    '        '    Dim aRcvMsg(1023) As Byte

    '        '    If oClient.Receive(aRcvMsg) > 1 Then
    '        '        'aMsg = Encoding.Default.GetString(aRcvMsg)
    '        '        aMsg = aMsg.TrimEnd(Chr(0)) & Encoding.Default.GetString(aRcvMsg)
    '        '        If aMsg.IndexOf("#") > 0 Then
    '        '            iTotal = aMsg.Substring(0, aMsg.IndexOf("#"))
    '        '        End If
    '        '        If iTotal > 0 Then
    '        '            If Encoding.Default.GetByteCount(aMsg) >= iTotal Then
    '        '                Exit Do
    '        '            End If
    '        '        End If


    '        '    End If

    '        'Loop
    '        strMsg = aMsg
    '        If Not String.IsNullOrEmpty(strMsg) Then
    '            UpdParaUI(strMsg)
    '        Else
    '            UpdParaUI("無任何資料回傳！")
    '        End If

    '    Catch ex As Exception
    '        UpdParaUI(ex.Message)
    '    Finally
    '        Try
    '            oClient.Shutdown(SocketShutdown.Both)
    '            oClient.Disconnect(False)
    '        Catch ex As Exception

    '        End Try
    '    End Try
    'End Sub
    Private Sub UpdParaUI(ByVal aMsg As String)
        Try
            If Me.InvokeRequired Then
                Dim act As New Action(Of String)(AddressOf UpdParaUI)
                Me.Invoke(act, aMsg)
            Else
                Me.txtResult.Text = aMsg

            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            UpdBtnUI(True)
        End Try
    End Sub
    Private Sub UpdBtnUI(ByVal aEnable As Boolean)
        Try
            If Me.InvokeRequired Then
                Dim act As New Action(Of Boolean)(AddressOf UpdBtnUI)
                Me.Invoke(act, aEnable)
            Else
                Me.btnConnect.Enabled = aEnable
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Function IsDataOK() As Boolean
        Try

            If String.IsNullOrEmpty(txtIPAddress.Text) Then
                MessageBox.Show("IP Address不允許空白", "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return False
            End If
            If Not IsValidIP(txtIPAddress.Text) Then
                MessageBox.Show("IP 不合法", "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return False
            End If
            If String.IsNullOrEmpty(txtSource_id.Text) Then
                MessageBox.Show("Source Id 不允許空白", "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return False
            End If
            If String.IsNullOrEmpty(txtdest_id.Text) Then
                MessageBox.Show("Dest_Id 不允許空白", "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return False
            End If

            If String.IsNullOrEmpty(txtMOP_PPID.Text) Then
                MessageBox.Show("MOP_PPID 不允許空白", "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return False
            End If
            If String.IsNullOrEmpty(txtSTBNo.Text) Then
                MessageBox.Show("STBNo 不允許空白", "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return False
            End If
            If String.IsNullOrEmpty(txtProductId.Text) Then
                MessageBox.Show("Product Id 不允許空白", "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return False
            End If
            If String.IsNullOrEmpty(txtcontent_id.Text) Then
                MessageBox.Show("Content Id 不允許空白", "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return False
            End If
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Private Sub btnSetOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSetOK.Click, btnSaveWebSet.Click
        Try
            If Not IsDataOK() Then
                Exit Sub
            End If
            My.Settings.additional_address_flag = txtadditional_address_flag.Text
            My.Settings.chipset_type_flag = txtchipset_type_flag.Text
            My.Settings.dest_id = txtdest_id.Text
            My.Settings.DMM_creation_date = dtpDMM_creation_date.Text
            My.Settings.DMM_creation_fate_flag = txtDMM_creation_fate_flag.Text
            My.Settings.Encrypted_asset_flag = txtEncrypted_asset_flag.Text
            My.Settings.IPAddress = txtIPAddress.Text
            My.Settings.MOP_PPID = txtMOP_PPID.Text
            My.Settings.nb_of_R_VOD_ENT = txtnb_of_R_VOD_ENT.Text
            My.Settings.OrderDate = dtpOrderDate.Text
            My.Settings.Port = numPort.Value
            My.Settings.RcvTimeOut = numRcvTimeOut.Value
            My.Settings.SendTimeOut = numSndTimeOut.Value
            My.Settings.source_id = txtSource_id.Text
            My.Settings.STBSNo = txtSTBNo.Text
            My.Settings.Validity_Flag = txtValidity_Flag.Text
            My.Settings.LowCmd = txtLowCmd.Text
            My.Settings.viewing_duration = txtviewing_duration.Text
            My.Settings.ProductId = txtProductId.Text
            My.Settings.Asset = txtcontent_id.Text
            My.Settings.expiration_date = dtpexpiration_date.Text
            My.Settings.Url = txtUrl.Text
            My.Settings.UrlDuration = numUrlDuration.Value.ToString
            My.Settings.UrlMetho = cboUrlMetho.Text
            My.Settings.UrlRepeat = numUrlRepeat.Value.ToString
            My.Settings.urlTimeOut = numUrlTimeout.Value.ToString
            InitialData()
            TabControl1.SelectTab(0)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try


    End Sub



    Private Sub frmClient_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If Not CableSoft_KeyPro.GetSystemInfo.IsRegister(ShowForm.Yes) Then
                MessageBox.Show("註冊不合法！", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Application.Exit()
            End If
            InitialData()
            DataToScreen()
            CtlEnableStatus(CmdStatus.Initial)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Application.Exit()
        End Try
        
    End Sub
    Private Sub ClearTextBox()
        Try
            txtSendMsg.Clear()
            txtResult.Clear()
            txtStatus.Clear()
            txtErrCode.Clear()
            txtErrName.Clear()
            txtReturnLines.Clear()
            txtUrlCode.Clear()
            txtUrlDesc.Clear()
        Catch ex As Exception

        End Try
    End Sub
    Private Sub btnConnect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnConnect.Click
        Try
            If Not IsDataOK() Then
                TabControl1.SelectTab(1)
                Exit Sub
            End If
            ClearTextBox()
            InitialSocket()
            FSocket.Connect(IPAddress.Parse(FIPAddress), Convert.ToInt32(FPort))
            FSocket.SendTimeout = FSendTimeOut * 1000
            FSocket.ReceiveTimeout = FRcvTimeOut * 1000
            Dim bty() As Byte = NagraCommand.BuilderCmd(CmdType.ConfirmConnect, Nothing).Item(0)
            txtSendMsg.Text = NagraCommand.GetReturnString(bty, True).TrimStart(Chr(0))
            FSocket.Send(bty)

            Dim Rcv(6) As Byte
            Thread.Sleep(300)
            If FSocket.Receive(Rcv) > 0 Then
                Dim obj As New NagraStatus_Type
                obj = NagraCommand.NagraStatus(Rcv, CmdType.ConfirmConnect)
                If (obj.SUCCESS = 6) AndAlso (obj.answer_code = 0) Then

                    txtResult.Text = "連線成功"
                    CtlEnableStatus(CmdStatus.ConnectOK)
                Else

                    txtResult.Text = "連線失敗" & vbCrLf & "Success = " & obj.SUCCESS & vbCrLf & _
                            "Answer : " & obj.answer_code
                    DisconnectSock()
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
            DisconnectSock()
        End Try
    End Sub
    Private Function RecvMsg(ByVal aType As CableSoft_Nagra_BuilderCmd.CmdType, _
                             ByRef aRetString As String) As Boolean
        Dim Rcv(1023) As Byte
        Dim pos As Int32 = 0
        Dim aMsg As String = String.Empty
        Dim aStopCount As New Stopwatch()
        aStopCount.Start()
        Try

            If FSocket.Receive(Rcv) > 0 Then
                Do
                    If pos = 0 Then
                        pos = NagraCommand.GetRetunLength(Rcv)
                        aMsg = NagraCommand.GetReturnString(Rcv, True)
                    Else
                        aMsg = aMsg & NagraCommand.GetReturnString(Rcv, False)
                    End If

                    If pos <= aMsg.Length Then
                        aRetString = aMsg
                        Return True
                    End If
                    If aStopCount.ElapsedMilliseconds * 1000 > FRcvTimeOut Then
                        aStopCount.Stop()
                        aStopCount.Reset()
                        Return False
                    End If
                Loop
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try




    End Function
    Private Sub btnCMD1002_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCMD1002.Click
        ClearTextBox()
        Try
            Dim objNagraCmdStatus As New NagraCmdStatus_Type
            Dim aString As String = String.Empty
            objNagraCmdStatus.Status = AckCode.Nack
            objNagraCmdStatus.ErrorCode = "-9999"
            objNagraCmdStatus.ErrorName = "CableSoft_Initial"
            aString = SendCmd1002(objNagraCmdStatus)

            If Not String.IsNullOrEmpty(txtResult.Text) Then
                aString = txtResult.Text & vbCrLf & aString
            End If
            txtResult.Text = aString
            StatusToScreen(objNagraCmdStatus)
            CtlEnableStatus(CmdStatus.CMD1002)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub DataToScreen()
        Try
            txtIPAddress.Text = FIPAddress
            numPort.Value = Convert.ToInt32(FPort)
            numSndTimeOut.Value = Convert.ToInt32(FSendTimeOut)
            numRcvTimeOut.Value = Convert.ToInt32(FRcvTimeOut)
            txtLowCmd.Text = FLowCmd
            txtSource_id.Text = Fsource_id
            txtSTBNo.Text = FSTBSNo
            txtValidity_Flag.Text = FValidity_Flag
            txtviewing_duration.Text = Fviewing_duration
            txtnb_of_R_VOD_ENT.Text = Fnb_of_R_VOD_ENT
            txtMOP_PPID.Text = FMOP_PPID
            txtEncrypted_asset_flag.Text = FEncrypted_asset_flag
            txtDMM_creation_fate_flag.Text = FDMM_creation_fate_flag
            txtdest_id.Text = Fdest_id
            txtchipset_type_flag.Text = Fchipset_type_flag
            txtadditional_address_flag.Text = Fadditional_address_flag
            dtpDMM_creation_date.Text = FDMM_creation_date
            dtpOrderDate.Value = FOrderDate
            txtProductId.Text = FProductId
            txtcontent_id.Text = FAsset
            dtpexpiration_date.Text = Fexpiration_date
            txtUrl.Text = FUrl
            numUrlDuration.Value = Convert.ToInt32(FUrlDuration)
            numUrlRepeat.Value = Convert.ToInt32(FUrlRepeat)
            numUrlTimeout.Value = Convert.ToInt32(FUrlTimeOut)
            cboUrlMetho.Text = FUrlMetho
        Catch ex As Exception

            MessageBox.Show("參數設定有誤！請先設定參數", "訊息", MessageBoxButtons.OK, MessageBoxIcon.Information)
            TabControl1.SelectTab(1)
            'MessageBox.Show(ex.Message, "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub StatusToScreen(ByVal obj As NagraCmdStatus_Type)
        Try
            txtStatus.Text = obj.Status.ToString
            If obj.Status = AckCode.Ack Then
                obj.ErrorCode = String.Empty
                obj.ErrorName = String.Empty
            End If
            txtErrCode.Text = obj.ErrorCode
            txtErrName.Text = obj.ErrorName
            txtReturnLines.Text = obj.Entity_data
        Catch ex As Exception
            MessageBox.Show(ex.Message, "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Function SendCmd1002(ByRef obj As NagraCmdStatus_Type) As String
        Try

            Dim aRetString As String = String.Empty
            Dim aSndbty() As Byte = NagraCommand.BuilderCmd(CmdType.Cmd1002, Nothing).Item(0)
            txtSendMsg.Text = NagraCommand.GetReturnString(aSndbty, True)
            Dim aRcvBty(1023) As Byte
            FSocket.Send(aSndbty)
            Thread.Sleep(100)
            If FSocket.Receive(aRcvBty) > 0 Then
                obj = NagraCommand.NagraStatus(aRcvBty, CmdType.Cmd1002)
            End If
            aRetString = NagraCommand.GetReturnString(aRcvBty, True)

            Return aRetString
        Catch ex As Exception
            MessageBox.Show(ex.Message, "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtResult.Text = ex.Message
            Return String.Empty
        End Try
    End Function


    Private Sub btnCMD0851_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCMD0851.Click
        ClearTextBox()
       
        Dim objAllFieldValue As New Dictionary(Of String, String)

        My.Settings.SEQNO = Convert.ToInt32(My.Settings.SEQNO) + 1
        Try
            objAllFieldValue.Add("SEQNO", My.Settings.SEQNO.ToString)
            objAllFieldValue.Add("STBSNO", FSTBSNo)
            objAllFieldValue.Add("VALIDITY_FLAG", FValidity_Flag)
            objAllFieldValue.Add("NB_OF_R_VOD_ENT", Fnb_of_R_VOD_ENT)
            objAllFieldValue.Add("CHIPSET_TYPE_FLAG", Fchipset_type_flag)
            objAllFieldValue.Add("ADDITIONAL_ADDRESS_FLAG", Fadditional_address_flag)
            objAllFieldValue.Add("DMM_CREATION_FATE_FLAG", FDMM_creation_fate_flag)
            objAllFieldValue.Add("DMM_CREATION_DATE", FDMM_creation_date)
            objAllFieldValue.Add("ENCRYPTED_ASSET_FLAG", FEncrypted_asset_flag)
            Dim aNotes As String = String.Empty
            aNotes = FProductId & "~" & FAsset & "~" & Format(Date.Parse(FOrderDate), "yyyyMMddhhmmss") & _
                    "~" & Format(Date.Parse(Fexpiration_date), "yyyyMMddhhmmss") & "~" & Fviewing_duration
            objAllFieldValue.Add("NOTES", aNotes)
            Dim lstBty As List(Of Byte()) = NagraCommand.BuilderCmd(CmdType.Cmd0851, objAllFieldValue)
            If lstBty Is Nothing OrElse lstBty.Count <= 0 Then
                txtResult.Text = "CMD0851組成有誤"
                Exit Sub
            End If
            '4C59313535383232323B3230313031313236303534373536
            Dim bty() As Byte = lstBty.Item(0)
            txtSendMsg.Text = NagraCommand.GetReturnString(bty, True)
            FSocket.Send(bty)
            Thread.Sleep(100)
            Dim aRetString As String = String.Empty
            If RecvMsg(CmdType.Cmd0851, aRetString) Then
                txtResult.Text = aRetString
                Dim objNagraStatus As New NagraCmdStatus_Type
                objNagraStatus = NagraCommand.GetCMDStatus(aRetString, aRetString.Length, CmdType.Cmd0851)
                StatusToScreen(objNagraStatus)
            Else
                txtResult.Text = "接收訊息錯誤" & vbCrLf & aRetString
            End If
            'txtReturnLines.Text = "333333"
        Catch ex As Exception
            MessageBox.Show(ex.Message, "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtResult.Text = ex.Message
        End Try

    End Sub



    Private Sub btnDisconnect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDisconnect.Click
        ClearTextBox()
        DisconnectSock()
        CtlEnableStatus(CmdStatus.Initial)
    End Sub
    Private Sub CtlEnableStatus(ByVal aType As CmdStatus)
        Try
            btnCMD0851.Enabled = False
            btnCMD1002.Enabled = False
            btnConnect.Enabled = False
            btnDisconnect.Enabled = False
            btnGetWebStatus.Enabled = False
            Select Case aType
                Case CmdStatus.Initial
                    btnConnect.Enabled = True
                Case CmdStatus.ConnectOK
                    btnCMD1002.Enabled = True
                    btnDisconnect.Enabled = True
                Case CmdStatus.CMD1002
                    btnCMD1002.Enabled = True
                    btnCMD0851.Enabled = True
                    btnGetWebStatus.Enabled = True
                    btnDisconnect.Enabled = True

            End Select
        Catch ex As Exception

        End Try
    End Sub

    Private Sub btnGetWebStatus_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGetWebStatus.Click
        Try
            Dim alicenses As String = txtReturnLines.Text
            ClearTextBox()
            If String.IsNullOrEmpty(FUrl) Then
                MessageBox.Show("網址不允許空值！", "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
                TabControl1.SelectTab(2)
                Exit Sub
            End If
            If String.IsNullOrEmpty(alicenses) Then
                MessageBox.Show("尚未取得授權碼", "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
            Dim aUrl As String = FUrl.TrimEnd("?")
            Dim aTurnBase64 As String = String.Empty
            aTurnBase64 = Convert.ToBase64String(Encoding.UTF8.GetBytes(alicenses))
            aUrl = aUrl & "?" & "target=" & FSTBSNo & _
                "&duration=" & FUrlDuration & _
                "&repeat=" & FUrlRepeat & _
                "&licenses=" & aTurnBase64
            ReadWeb(aUrl)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub ReadWeb(ByVal aUrl As String)
        Dim aRequest As HttpWebRequest = Nothing
        Dim aResponse As HttpWebResponse = Nothing
        Try
            txtSendMsg.Text = aUrl
            aRequest = HttpWebRequest.Create(aUrl)
            aRequest.Method = FUrlMetho
            aRequest.Timeout = Convert.ToInt32(FUrlTimeOut) * 1000

            aRequest.Credentials = CredentialCache.DefaultCredentials

            aResponse = aRequest.GetResponse

            txtUrlCode.Text = aResponse.StatusCode
            txtUrlDesc.Text = aResponse.StatusDescription
            txtResult.Text = aResponse.Headers.ToString
            'Using stm As Stream = aResponse.GetResponseStream
            '    Using reader As New StreamReader(stm)
            '        txtResult.Text = reader.ReadToEnd
            '        reader.Close()
            '    End Using
            '    stm.Close()
            'End Using
        Catch ex As WebException
            txtUrlCode.Text = ex.Status
            txtUrlDesc.Text = ex.Message
            Try
                Dim stc As New StackTrace(ex)
                txtResult.Text = stc.GetFrame(0).GetMethod.Name
            Catch ex2 As Exception

            End Try

        Catch ex As Exception
            'txtResult.Text = aResponse.Headers.ToString
            txtResult.Text = ex.Message
        End Try
    End Sub





End Class
Public Enum CmdStatus
    Initial = 0
    ConnectOK = 1
    CMD1002 = 3
    Other = 4
End Enum