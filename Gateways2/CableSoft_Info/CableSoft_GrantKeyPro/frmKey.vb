Imports System
Imports System.Windows
Imports System.Windows.Forms
Imports System.IO
Imports System.Text
Imports CableSoft_KeyPro
Imports System.Globalization

Public Class frmKey
    Private FSoftInfo As String = Nothing
    Private FInfoLst As System.Collections.Generic.Dictionary(Of String, String)
    Private Sub txtSoftSN_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSoftSN.TextChanged
        FSoftInfo = Nothing
        btnOK.Enabled = False
        FInfoLst.Clear()
        ClearTextBox()
        If String.IsNullOrEmpty(txtSoftSN.Text) Then
            txtRegSN.Text = String.Empty
        Else
            Dim bty() As Byte = Encoding.Default.GetBytes(txtSoftSN.Text)
            Using stm As New MemoryStream(bty)
                FSoftInfo = CableSoft_KeyPro.GetSystemInfo.DecryKeyPro(stm)
                If FSoftInfo.Split(Chr(0)).Length = 2 Then
                    FSoftInfo = FSoftInfo.Split(Chr(0)).GetValue(1)
                    If (FSoftInfo.Split(Environment.NewLine).Length >= 5) AndAlso _
                        (FSoftInfo.Split(Environment.NewLine).Length <= 7) Then
                        FInfoLst.Add("CPU", FSoftInfo.Split(Environment.NewLine).GetValue(0))
                        FInfoLst.Add("HD", FSoftInfo.Split(Environment.NewLine).GetValue(1))
                        FInfoLst.Add("MB", FSoftInfo.Split(Environment.NewLine).GetValue(2))
                        FInfoLst.Add("MAC", FSoftInfo.Split(Environment.NewLine).GetValue(3))
                        FInfoLst.Add("APP", FSoftInfo.Split(Environment.NewLine).GetValue(4))
                        txtCpu.Text = FInfoLst("CPU")
                        txtHD.Text = FInfoLst("HD")
                        txtMB.Text = FInfoLst("MB")
                        txtMac.Text = FInfoLst("MAC")
                        txtApp.Text = FInfoLst("APP")
                        If FSoftInfo.Split(Environment.NewLine).Length > 5 Then

                        End If
                        btnOK.Enabled = True
                    Else
                        MessageBox.Show("軟體序號解析有誤!", "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End If
                Else
                    MessageBox.Show("軟體序號不正確 !", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)

                End If
            End Using
        End If
    End Sub

    Private Sub ClearTextBox()
        Try
            txtCpu.Text = String.Empty
            txtApp.Text = String.Empty
            txtHD.Text = String.Empty
            txtMac.Text = String.Empty
            txtMB.Text = String.Empty
            txtRegSN.Text = String.Empty
        Catch ex As Exception

        End Try
    End Sub

    Private Sub frmKey_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If FInfoLst Is Nothing Then
            FInfoLst = New System.Collections.Generic.Dictionary(Of String, String)

        End If

        btnOK.Enabled = False

    End Sub
    Private Function IsDataOK() As Boolean
        Try
            If dtpUseDate.Value.Date <= Date.Now.Date Then
                MessageBox.Show("使用到期日必需大於今天日期！", "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return False
            End If
            If dtpUseDate.Value < dtpInsDate.Value Then
                MessageBox.Show("使用到期日必需大於安裝日期！", "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return False
            End If
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function
    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Try
            If Not IsDataOK() Then
                Exit Sub
            End If
            Dim aRet As String = Nothing
            aRet = txtCpu.Text & Environment.NewLine & _
                txtHD.Text & Environment.NewLine & _
                txtMB.Text & Environment.NewLine & _
                txtMac.Text & Environment.NewLine & _
                txtApp.Text & Environment.NewLine & _
                dtpInsDate.Value.Date.ToString(CultureInfo.CreateSpecificCulture("zh-TW")) & Environment.NewLine & _
                dtpUseDate.Value.Date.ToString(CultureInfo.CreateSpecificCulture("zh-TW"))
            aRet = GetSystemInfo.EncryKeyPro(aRet)

            txtRegSN.Text = aRet

        Catch ex As Exception

        End Try


    End Sub

    Private Sub btnCopy_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCopy.Click
        Clipboard.SetText(txtRegSN.Text)
    End Sub

    Private Sub txtRegSN_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRegSN.TextChanged
        If String.IsNullOrEmpty(txtRegSN.Text) Then
            btnCopy.Enabled = False
        Else
            btnCopy.Enabled = True
        End If
    End Sub
End Class
