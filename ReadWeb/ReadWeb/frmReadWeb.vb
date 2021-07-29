Imports System
Imports System.Text
Imports System.IO
Imports System.Net
Imports CableSoft_KeyPro

Public Class frmReadWeb
    Private strUrl As String = String.Empty
    Private Const FTimeout As Int32 = 10
    Private Function IsDataOK() As Boolean
        Try
            If String.IsNullOrEmpty(txtUrl.Text) Then
                MessageBox.Show("請輸入網址", "訊息", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtUrl.Focus()
                Return False
            End If
            If String.IsNullOrEmpty(txtSTBNo.Text) Then
                MessageBox.Show("請輸入STBNo", "訊息", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtSTBNo.Focus()
                Return False
            End If
            If txtSTBNo.Text.Length <> 10 Then
                MessageBox.Show("STBNo 碼數不正確", "訊息", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtSTBNo.Focus()
                Return False
            End If
            If String.IsNullOrEmpty(txtLicenese.Text) Then
                MessageBox.Show("請輸入授權碼", "訊息", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtLicenese.Focus()
                Return False
            End If
            Return True
        Catch ex As Exception
            MessageBox.Show(ex.Message, "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function
    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        If Not IsDataOK() Then
            Exit Sub
        End If
        Try
            btnOK.Enabled = False
            txtResult.Clear()
            txtCode.Clear()
            txtDesc.Clear()

            strUrl = txtUrl.Text & txtPara.Text
            My.Settings.Url = txtUrl.Text
            My.Settings.Para = txtPara.Text
            My.Settings.Metho = cboMetho.Text
            My.Settings.TimeOut = numTimeout.Value.ToString
            My.Settings.Repeat = numRepeat.Value.ToString
            My.Settings.Duration = numDuration.Value.ToString
            My.Settings.STBNo = txtSTBNo.Text
            My.Settings.LiceneseNo = txtLicenese.Text
            My.Settings.Save()
            Threading.Thread.Sleep(100)
            ReadWeb(strUrl)
        Catch ex As Exception
            txtResult.Text = ex.Message
        Finally
            btnOK.Enabled = True
        End Try

    End Sub
    Private Sub ReadWeb(ByVal aUrl As String)
        Dim aRequest As HttpWebRequest = Nothing
        Dim aResponse As HttpWebResponse = Nothing
        Try
            aRequest = HttpWebRequest.Create(aUrl)
            aRequest.Method = cboMetho.Text
            aRequest.Timeout = Convert.ToInt32((numTimeout.Value * 1000))

            aRequest.Credentials = CredentialCache.DefaultCredentials

            aResponse = aRequest.GetResponse
            'MsgBox(aResponse.Headers.Count)
            txtCode.Text = aResponse.StatusCode
            txtDesc.Text = aResponse.StatusDescription
            txtResult.Text = aResponse.Headers.ToString
            'Using stm As Stream = aResponse.GetResponseStream
            '    Using reader As New StreamReader(stm)
            '        txtResult.Text = reader.ReadToEnd
            '        reader.Close()
            '    End Using
            '    stm.Close()
            'End Using
        Catch ex As WebException
            txtCode.Text = ex.Status
            txtDesc.Text = ex.Message
            Try
                Dim stc As New StackTrace(ex)
                'txtResult.Text = aResponse.Headers.ToString
                txtResult.Text = stc.GetFrame(0).GetMethod.Name
            Catch ex2 As Exception

            End Try

        Catch ex As Exception
            'txtResult.Text = aResponse.Headers.ToString
            txtResult.Text = ex.Message
        End Try
    End Sub

    Private Sub frmReadWeb_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If Not GetSystemInfo.IsRegister(ShowForm.Yes) Then
                MessageBox.Show("註冊不合法！", "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Application.Exit()
            End If
            txtUrl.Text = My.Settings.Url
            txtPara.Text = My.Settings.Para
            cboMetho.Text = My.Settings.Metho
            If Not String.IsNullOrEmpty(My.Settings.TimeOut) Then
                numTimeout.Value = Convert.ToInt32(My.Settings.TimeOut)
            End If
            txtSTBNo.Text = My.Settings.STBNo
            txtLicenese.Text = My.Settings.LiceneseNo
            If Not String.IsNullOrEmpty(My.Settings.Duration) Then
                numDuration.Value = Convert.ToInt32(My.Settings.Duration)
            End If
            If Not String.IsNullOrEmpty(My.Settings.Repeat) Then
                numRepeat.Value = Convert.ToInt32(My.Settings.Repeat)
            End If
            If String.IsNullOrEmpty(cboMetho.Text) Then
                cboMetho.Text = "POST"
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Application.Exit()
        End Try
        
    End Sub


End Class
