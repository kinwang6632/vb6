Imports System
Imports System.Data
Imports System.Data.OracleClient

Public Class frmMain
    Private FcnBdu As OracleConnectionStringBuilder = Nothing
    Private FNotes As String = String.Empty

    Private Function GetDate(ByVal aStr As String) As DateTime

        Dim aStrDate As String = Strings.Left(aStr, 4) & "/" & _
                aStr.Substring(4, 2) & "/" & aStr.Substring(6, 2) & " " & _
                aStr.Substring(8, 2) & ":" & aStr.Substring(10, 2) & ":" & _
                Strings.Right(aStr, 2)
        Return DateTime.Parse(aStrDate)
        'Return Format(TimeZoneInfo.ConvertTimeToUtc(DateTime.Parse(aStrDate)), "yyyyMMddhhmmss")
    End Function
    Private Sub btnInsBatch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnInsBatch.Click
        If Not IsDataOK() Then
            Exit Sub
        End If
        My.Settings.FSetAsset = txtAsset.Text
        My.Settings.FSetPassword = txtPassword.Text
        My.Settings.FSetProductId = txtProductId.Text
        My.Settings.FSetSid = txtSid.Text
        My.Settings.FSetSTBNo = txtSTBNo.Text
        My.Settings.FSetUserId = txtUserId.Text
        My.Settings.FSetViewing_duration = numViewing_duration.Value
        My.Settings.FSetCompCode = numCompCode.Value
        If FcnBdu Is Nothing Then
            FcnBdu = New OracleConnectionStringBuilder()
        End If
        FcnBdu.Clear()
        FcnBdu.UserID = txtUserId.Text
        FcnBdu.Password = txtPassword.Text
        FcnBdu.DataSource = txtSid.Text
        FcnBdu.Unicode = True
        Dim aNotes As String = txtProductId.Text & "~" & txtAsset.Text & "~" & _
                Format(dtOrderDate.Value, "yyyyMMddHHmmss") & "~" & _
                Format(dtexpiration_date.Value, "yyyyMMddHHmmss") & "~" & _
                numViewing_duration.Value.ToString
        InsertData(numBatch.Value, aNotes)
        
    End Sub
    Private Function IsDataOK() As Boolean
        If String.IsNullOrEmpty(txtUserId.Text) Then
            MessageBox.Show("UserId無值", "訊息", MessageBoxButtons.OK)
            Return False
        End If
        If String.IsNullOrEmpty(txtPassword.Text) Then
            MessageBox.Show("Password無值", "訊息", MessageBoxButtons.OK)
            Return False
        End If
        If String.IsNullOrEmpty(txtSid.Text) Then
            MessageBox.Show("Password無值", "訊息", MessageBoxButtons.OK)
            Return False
        End If
        If String.IsNullOrEmpty(txtSTBNo.Text) Then
            MessageBox.Show("STBNo無值", "訊息", MessageBoxButtons.OK)
            Return False
        End If
        If String.IsNullOrEmpty(txtProductId.Text) Then
            MessageBox.Show("ProductId無值", "訊息", MessageBoxButtons.OK)
            Return False
        End If
        If String.IsNullOrEmpty(txtAsset.Text) Then
            MessageBox.Show("Asset無值", "訊息", MessageBoxButtons.OK)
            Return False
        End If
        Return True
    End Function
    Private Sub InsertData(ByVal aRcd As Int32, ByVal aNotes As String)
        Try
            Using cn As New OracleConnection(FcnBdu.ConnectionString)
                cn.Open()
                Dim tra As OracleTransaction = cn.BeginTransaction
                Using cmd As OracleCommand = cn.CreateCommand
                    cmd.Transaction = tra
                    For i As Int32 = 1 To aRcd
                        cmd.Parameters.Clear()
                        Dim aSQLSeqNo As String = "select S_SO560_SEQNO.nextval from dual "
                        cmd.CommandText = aSQLSeqNo
                        Dim aSEQNo As Int32 = Convert.ToInt32(cmd.ExecuteScalar)
                        Dim aInsSQL As String = "Insert into SO560 (SeqNo,Cmd,CmdStatus, " & _
                                "ExchTime,Encrypted_asset_flag,Validity_Flag,nb_of_R_VOD_ENT," & _
                                "chipset_type_flag,additional_address_flag,DMM_creation_fate_flag," & _
                                "DMM_creation_date,CompCode,Notes,STBSNo) Values( " & _
                                ":SeqNo,:Cmd,:CmdStatus,:ExchTime,:Encrypted_asset_flag," & _
                                ":Validity_Flag,:nb_of_R_VOD_ENT,:chipset_type_flag,:additional_address_flag," & _
                                ":DMM_creation_fate_flag,:DMM_creation_date,:CompCode,:Notes,:STBSNo)"

                        Dim pSeqNo As New OracleClient.OracleParameter("SeqNo", OracleClient.OracleType.Number)
                        Dim pCmd As New OracleParameter("Cmd", OracleType.VarChar)
                        Dim pCmdStatus As New OracleParameter("CmdStatus", OracleType.VarChar)
                        Dim pExchTime As New OracleParameter("ExchTime", OracleType.DateTime)
                        Dim pEncrypted_asset_flag As New OracleParameter("Encrypted_asset_flag", OracleType.Number)
                        Dim pValidity_Flag As New OracleParameter("Validity_Flag", OracleType.Number)
                        Dim pnb_of_R_VOD_ENT As New OracleParameter("nb_of_R_VOD_ENT", OracleType.Number)
                        Dim pchipset_type_flag As New OracleParameter("chipset_type_flag", OracleType.Number)
                        Dim padditional_address_flag As New OracleParameter("additional_address_flag", OracleType.Number)
                        Dim pDMM_creation_fate_flag As New OracleParameter("DMM_creation_fate_flag", OracleType.Number)
                        Dim pDMM_creation_date As New OracleParameter("DMM_creation_date", OracleType.DateTime)
                        Dim pCompCode As New OracleParameter("CompCode", OracleType.Number)
                        Dim pNotes As New OracleParameter("Notes", OracleType.VarChar)
                        Dim pSTBSNo As New OracleParameter("STBSNo", OracleType.VarChar)
                        pSeqNo.Value = aSEQNo
                        pCmd.Value = "B1"
                        pCmdStatus.Value = "W"
                        pExchTime.Value = Date.Now
                        pEncrypted_asset_flag.Value = 1
                        pValidity_Flag.Value = numValidity_Flag.Value
                        pnb_of_R_VOD_ENT.Value = 0
                        pchipset_type_flag.Value = 0
                        padditional_address_flag.Value = 0
                        pDMM_creation_fate_flag.Value = numDMM_creation_fate_flag.Value
                        If numDMM_creation_fate_flag.Value = 1 Then
                            pDMM_creation_date.Value = dtDMM_creation_date.Value
                        Else
                            pDMM_creation_date.Value = DBNull.Value
                        End If

                        pCompCode.Value = numCompCode.Value
                        pNotes.Value = aNotes
                        pSTBSNo.Value = txtSTBNo.Text
                        cmd.Parameters.Add(pSeqNo)
                        cmd.Parameters.Add(pCmd)
                        cmd.Parameters.Add(pCmdStatus)
                        cmd.Parameters.Add(pExchTime)
                        cmd.Parameters.Add(pEncrypted_asset_flag)
                        cmd.Parameters.Add(pValidity_Flag)
                        cmd.Parameters.Add(pnb_of_R_VOD_ENT)
                        cmd.Parameters.Add(pchipset_type_flag)
                        cmd.Parameters.Add(padditional_address_flag)
                        cmd.Parameters.Add(pDMM_creation_fate_flag)
                        cmd.Parameters.Add(pDMM_creation_date)
                        cmd.Parameters.Add(pCompCode)
                        cmd.Parameters.Add(pNotes)
                        cmd.Parameters.Add(pSTBSNo)
                        cmd.CommandText = aInsSQL
                        cmd.ExecuteNonQuery()


                        Dim strAry() As String = aNotes.Split(","c)
                        For j As Int32 = 0 To strAry.Length - 1
                            cmd.Parameters.Clear()
                            Dim pSeq2 As New OracleParameter("SeqNo", OracleType.Number)
                            Dim pProductID As New OracleParameter("ProductID", OracleType.VarChar)
                            Dim pVODItem As New OracleParameter("VODItem", OracleType.VarChar)
                            Dim pOrderDate As New OracleParameter("OrderDate", OracleType.DateTime)
                            Dim pexpiration_date As New OracleParameter("expiration_date", OracleType.DateTime)
                            Dim pviewing_duration As New OracleParameter("viewing_duration", OracleType.Number)
                            Dim aValue() As String = strAry.GetValue(j).ToString.Split("~"c)
                            pSeq2.Value = aSEQNo
                            pProductID.Value = aValue.GetValue(0)
                            pVODItem.Value = aValue(1)
                            pOrderDate.Value = GetDate(aValue(2))
                            pexpiration_date.Value = GetDate(aValue(3))
                            pviewing_duration.Value = Int32.Parse(aValue(4))

                            Dim aSO561BSQL As String = "Insert Into SO561B (SeqNo,ProductCASID," & _
                                                    " MediaCASID,OrderDate,expiration_date,viewing_duration)" & _
                                                    " Values(:SeqNo,:ProductID,:VODItem,:OrderDate,:expiration_date," & _
                                                    " :viewing_duration)"
                            cmd.CommandText = aSO561BSQL
                            cmd.Parameters.Add(pSeq2)
                            cmd.Parameters.Add(pProductID)
                            cmd.Parameters.Add(pVODItem)
                            cmd.Parameters.Add(pOrderDate)
                            cmd.Parameters.Add(pexpiration_date)
                            cmd.Parameters.Add(pviewing_duration)
                            cmd.ExecuteNonQuery()
                        Next


                    Next
                    tra.Commit()
                    MessageBox.Show("完成", "訊息", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End Using
            End Using
        Catch ex As Exception

            MessageBox.Show(ex.Message, "訊息", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Try
    End Sub
    Private Sub frmMain_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        txtUserId.Text = My.Settings.FSetUserId
        txtSTBNo.Text = My.Settings.FSetSTBNo
        txtSid.Text = My.Settings.FSetSid
        txtProductId.Text = My.Settings.FSetProductId
        txtPassword.Text = My.Settings.FSetPassword
        txtAsset.Text = My.Settings.FSetAsset
        numViewing_duration.Value = My.Settings.FSetViewing_duration
        numCompCode.Value = My.Settings.FSetCompCode

    End Sub

    Private Sub btnBatchSingle_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBatchSingle.Click
        If Not IsDataOK() Then Exit Sub
        My.Settings.FSetAsset = txtAsset.Text
        My.Settings.FSetPassword = txtPassword.Text
        My.Settings.FSetProductId = txtProductId.Text
        My.Settings.FSetSid = txtSid.Text
        My.Settings.FSetSTBNo = txtSTBNo.Text
        My.Settings.FSetUserId = txtUserId.Text
        My.Settings.FSetViewing_duration = numViewing_duration.Value
        My.Settings.FSetCompCode = numCompCode.Value
        FNotes = String.Empty
        btnIns.Enabled = True
        btnBatchSingle.Enabled = False
    End Sub

    Private Sub btnIns_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnIns.Click
        btnOK.Enabled = True
        Dim aNotes As String = txtProductId.Text & "~" & txtAsset.Text & "~" & _
                Format(dtOrderDate.Value, "yyyyMMddhhmmss") & "~" & _
                Format(dtexpiration_date.Value, "yyyyMMddhhmmss") & "~" & _
                numViewing_duration.Value.ToString
        If String.IsNullOrEmpty(FNotes) Then
            FNotes = aNotes
        Else
            FNotes = FNotes & "," & aNotes
        End If

    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Try
            If FcnBdu Is Nothing Then
                FcnBdu = New OracleConnectionStringBuilder()
            End If
            FcnBdu.Clear()
            FcnBdu.UserID = txtUserId.Text
            FcnBdu.Password = txtPassword.Text
            FcnBdu.DataSource = txtSid.Text
            FcnBdu.Unicode = True
            InsertData(1, FNotes)
        Finally
            btnOK.Enabled = False
            btnIns.Enabled = False
            btnBatchSingle.Enabled = True
        End Try
        
    End Sub
End Class
