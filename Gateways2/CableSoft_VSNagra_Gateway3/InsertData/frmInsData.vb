Imports System
Imports System.Data

Public Class frmInsData
    Private connectString As String = Nothing
    Private Sub btnOK_Click(sender As System.Object, e As System.EventArgs) Handles btnOK.Click
        Dim bducnstring As New OracleClient.OracleConnectionStringBuilder()
        bducnstring.DataSource = txtSid.Text
        bducnstring.UserID = txtUserId.Text
        bducnstring.Password = txtPassword.Text
        bducnstring.PersistSecurityInfo = True
        connectString = bducnstring.ConnectionString
        Dim tra As OracleClient.OracleTransaction = Nothing
        If String.IsNullOrEmpty(txtHIGH_LEVEL_CMD_ID.Text) Then
            MessageBox.Show("HIGH_LEVEL_CMD_ID 無值!")
            Exit Sub
        End If
        If String.IsNullOrEmpty(txtCompCode.Text) Then
            MessageBox.Show("CompCode 無值!")
            Exit Sub
        End If
        Try
            Using cn As New OracleClient.OracleConnection(connectString)
                cn.Open()
                Dim strSeq As String = Nothing
                tra = cn.BeginTransaction
                For i As Int32 = 1 To Int32.Parse(txtRcd.Text)
                    'Using cmdSeq As New OracleClient.OracleCommand("SELECT S_SendNSTV_SeqNo.Nextval FROM dual")
                    '    cmdSeq.Transaction = tra
                    '    cmdSeq.Connection = cn
                    '    strSeq = cmdSeq.ExecuteScalar
                    'End Using
                    Using cmd As New OracleClient.OracleCommand
                        cmd.Connection = cn
                        cmd.Transaction = tra
                        cmd.Parameters.Clear()
                        cmd.Parameters.Add(New OracleClient.OracleParameter("HIGH_LEVEL_CMD_ID", txtHIGH_LEVEL_CMD_ID.Text))
                        cmd.Parameters.Add(New OracleClient.OracleParameter("NOTES", IIf(String.IsNullOrEmpty(txtNOTES.Text), DBNull.Value, txtNOTES.Text)))
                        cmd.Parameters.Add(New OracleClient.OracleParameter("ICC_NO", txtICC_NO.Text))
                        cmd.Parameters.Add(New OracleClient.OracleParameter("STB_NO", txtSTB_NO.Text))
                        cmd.Parameters.Add(New OracleClient.OracleParameter("COMPCODE", Int32.Parse(txtCompCode.Text)))
                        'cmd.Parameters.Add(New OracleClient.OracleParameter("SEQNO", Int32.Parse(strSeq)))
                        cmd.Parameters.Add(New OracleClient.OracleParameter("CMD_STATUS", "W"))
                        cmd.Parameters.Add(New OracleClient.OracleParameter("ZIP_CODE", txt_ZipCode.Text))
                        cmd.Parameters.Add(New OracleClient.OracleParameter("subscription_begin_date", Date.Parse(dtpSubscription_begin_date.Value.ToShortDateString)))
                        cmd.Parameters.Add(New OracleClient.OracleParameter("subscription_end_date", Date.Parse(dtpSubscription_end_date.Value.ToShortDateString)))
                        cmd.Parameters.Add(New OracleClient.OracleParameter("MIS_IRD_CMD_DATA", txtMIS_IRD_CMD_DATA.Text))
                        cmd.Parameters.Add(New OracleClient.OracleParameter("GateWayType", txtGatewayType.Text))
                        If Not String.IsNullOrEmpty(txtProcessingDate.Text) Then
                            cmd.Parameters.Add(New OracleClient.OracleParameter("ProcessingDate", New DateTime(
                                                                                Integer.Parse(txtProcessingDate.Text.Substring(0, 4)),
                                                                                Integer.Parse(txtProcessingDate.Text.Substring(4, 2)),
                                                                                Integer.Parse(txtProcessingDate.Text.Substring(6, 2)),
                                                                                Integer.Parse(txtProcessingDate.Text.Substring(8, 2)),
                                                                                Integer.Parse(txtProcessingDate.Text.Substring(10, 2)),
                                                                                Integer.Parse(txtProcessingDate.Text.Substring(12, 2)))))
                        Else
                            cmd.Parameters.Add(New OracleClient.OracleParameter("ProcessingDate", DBNull.Value))
                        End If

                        cmd.CommandText = "INSERT INTO SEND_Nagra (HIGH_LEVEL_CMD_ID, " & _
                                                      "SEQNO,CMD_STATUS,NOTES,ICC_NO,STB_NO,COMPCODE,ZIP_CODE " & _
                                                      ",subscription_begin_date,subscription_end_date" & _
                                                      ",MIS_IRD_CMD_DATA,ProcessingDate,GateWayType) VALUES ( " & _
                                                      ":HIGH_LEVEL_CMD_ID,S_SendNSTV_SeqNo.Nextval," & _
                                                      ":CMD_STATUS,:NOTES,:ICC_NO" & _
                                                      ",:STB_NO,:COMPCODE,:ZIP_CODE, " & _
                                                      ":subscription_begin_date,:subscription_end_date," & _
                                                      ":MIS_IRD_CMD_DATA,:ProcessingDate,:GateWayType)"

                        cmd.ExecuteNonQuery()

                        'cmd.CommandText = "INSERT INTO SEND_NAGRA (HIGH_LEVEL_CMD_ID, " & _
                        '                                "NOTES,ICC_NO,STB_NO,SEQNO,COMPCODE,CMD_STATUS) VALUES ( " & _
                        '                                "'" & txtHIGH_LEVEL_CMD_ID.Text & "','" & txtNOTES.Text & "','" & txtICC_NO.Text & "'" & _
                        '                                ",'" & txtSTB_NO.Text & "'," & _
                        '                                "SELECT S_SendNagra_SeqNo.Nextval FROM dual," & txtCompCode.Text & ",'W' )"
                        'cmd.ExecuteNonQuery()
                    End Using
                Next
             
                tra.Commit()
                MessageBox.Show("新增完成", "訊息", MessageBoxButtons.OK)
            End Using

        Catch ex As Exception
            If tra IsNot Nothing Then
                tra.Rollback()
            End If
            MessageBox.Show(ex.ToString)
        End Try
       


    End Sub
End Class
