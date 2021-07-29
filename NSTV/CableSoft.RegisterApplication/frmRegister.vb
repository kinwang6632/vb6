Public Class frmRegister

    Private Sub btnOK_Click(sender As System.Object, e As System.EventArgs) Handles btnOK.Click
        If CableSoft_KeyPro.GetSystemInfo.RegisterApplication(txtFileName.Text) Then
            MessageBox.Show("註冊成功！", "訊息", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            MessageBox.Show("註冊失敗！", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    Private Sub btnSelect_Click(sender As System.Object, e As System.EventArgs) Handles btnSelect.Click
        If OpenFileDialog1.ShowDialog Then
            txtFileName.Text = OpenFileDialog1.SafeFileName
        End If
    End Sub
End Class
