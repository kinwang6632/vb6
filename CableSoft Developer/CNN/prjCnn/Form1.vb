Imports System.IO
Imports System.Text
Imports System.Collections
Imports Microsoft.VisualBasic.FileIO
Imports System.Diagnostics
Imports System.Security.Cryptography

Public Class frmConnPara
    Private strFile As String = Nothing

    Private Sub btnOpenFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOpenFile.Click
        Dim i As Integer = 0
        Dim strPara As String = String.Empty
        Using aFile As New OpenFileDialog()
            aFile.Filter = "連線參數|ConPara.Dat"
            aFile.Multiselect = False
            aFile.ShowDialog()
            strFile = aFile.FileName
        End Using
        If strFile <> Nothing Then
            Try
                Using strReader As New StreamReader(strFile, Encoding.Default)
                    dtgConn.Rows.Clear()
                    Do While Not strReader.EndOfStream
                        strPara = CryptUtil.Decrypt(strReader.ReadLine)
                        If strPara.Split(New String() {",@"}, StringSplitOptions.None).Length <> 2 Then
                            MessageBox.Show("檔案格式錯誤！", "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            Exit Do
                        End If
                        dtgConn.Rows.Add()
                        i += 1
                        dtgConn.Rows(i - 1).Cells(0).Value = strPara.Split(New String() {",@"}, StringSplitOptions.None).GetValue(0)
                        dtgConn.Rows(i - 1).Cells(1).Value = strPara.Split(New String() {",@"}, StringSplitOptions.None).GetValue(1)
                    Loop
                    strReader.Close()
                    'strReader = Nothing
                    'dtgConn.Rows(0).Visible = True
                    dtgConn.Update()
                    btnSave.Enabled = True
                End Using
            Catch exp As Exception
                MessageBox.Show("檔案開啟錯誤", "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        Else
            MessageBox.Show("沒有選擇檔案", "訊息", MessageBoxButtons.OK, MessageBoxIcon.Information)
            btnSave.Enabled = False
        End If
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Dim strContext As String = String.Empty
        Try
            Dim stmSave As New StreamWriter(strFile, False, Encoding.Default)
            If dtgConn.Rows.Count >= 1 Then
                For i As Integer = 0 To dtgConn.Rows.Count - 1
                    If dtgConn.Rows(i).Cells(0).Value <> String.Empty And dtgConn.Rows(i).Cells(1).Value <> String.Empty Then
                        strContext = dtgConn.Rows(i).Cells(0).Value & ",@" & dtgConn.Rows(i).Cells(1).Value
                        stmSave.WriteLine(CryptUtil.Encrypt(strContext))
                    End If
                Next
            End If
            stmSave.Flush()
            stmSave.Close()
            stmSave = Nothing
            MessageBox.Show("儲存成功！", "訊息", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show("儲存失敗！", "警告訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

End Class
Public Class CryptUtil

    Private Shared KEY_64() As Byte = {42, 16, 93, 156, 78, 4, 218, 32}
    Private Shared IV_64() As Byte = {55, 103, 246, 79, 36, 99, 167, 3}

    Private Shared KEY_192() As Byte = {42, 16, 93, 156, 78, 4, 218, 32, 15, 167, 44, 80, 26, 250, 155, 112, 2, 94, 11, 204, 119, 35, 184, 197}
    Private Shared IV_192() As Byte = {55, 103, 246, 79, 36, 99, 167, 3, 42, 5, 62, 83, 184, 7, 209, 13, 145, 23, 200, 58, 173, 10, 121, 222}

    Public Shared Function Encrypt(ByVal value As String) As String
        Dim cryptoProvider As New DESCryptoServiceProvider
        Dim ms As New MemoryStream()
        Dim cs As New CryptoStream(ms, cryptoProvider.CreateEncryptor(KEY_64, IV_64), CryptoStreamMode.Write)
        Dim sw As New StreamWriter(cs)
        sw.Write(value)
        sw.Flush()
        cs.FlushFinalBlock()
        ms.Flush()
        Return Convert.ToBase64String(ms.GetBuffer(), 0, ms.Length)
    End Function

    Public Shared Function Decrypt(ByVal value As String) As String
        Dim cryptoProvider As New DESCryptoServiceProvider()
        Dim buffer As Byte() = Convert.FromBase64String(value)
        Dim ms As MemoryStream = New MemoryStream(buffer)
        Dim cs As New CryptoStream(ms, cryptoProvider.CreateDecryptor(KEY_64, IV_64), CryptoStreamMode.Read)
        Dim sr As StreamReader = New StreamReader(cs)
        Return sr.ReadToEnd()
    End Function

    Public Shared Function EncryptTripleDES(ByVal value As String) As String
        Dim cryptoProvider As New TripleDESCryptoServiceProvider()
        Dim ms As New MemoryStream()
        Dim cs As New CryptoStream(ms, cryptoProvider.CreateEncryptor(KEY_192, IV_192), CryptoStreamMode.Write)
        Dim sw As New StreamWriter(cs)
        sw.Write(value)
        sw.Flush()
        cs.FlushFinalBlock()
        ms.Flush()
        Return Convert.ToBase64String(ms.GetBuffer(), 0, ms.Length)
    End Function

    Public Shared Function DecryptTripleDES(ByVal value As String) As String
        Dim cryptoProvider As New TripleDESCryptoServiceProvider()
        Dim buffer As Byte() = Convert.FromBase64String(value)
        Dim ms As New MemoryStream(buffer)
        Dim cs As New CryptoStream(ms, cryptoProvider.CreateDecryptor(KEY_192, IV_192), CryptoStreamMode.Read)
        Dim sr As New StreamReader(cs)
        Return sr.ReadToEnd()
    End Function

End Class