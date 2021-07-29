Imports CableSoft.NSTV
Imports CableSoft.NSTV.Encry
Imports CableSoft.NSTV.XMLFileIO
Imports System
Imports System.Data
Imports System.IO
Imports System.Text
Public Class frmEncry
    Private FileName As String = Nothing
    Private Const Key As String = "CABLESOFT1234567"
    Private xmlFile As CableSoft.NSTV.XMLFileIO.XmlFileIO
    Private Sub btnSelect_Click(sender As System.Object, e As System.EventArgs) Handles btnSelect.Click
        If OpenFileDialog1.ShowDialog Then
            xmlFile = New CableSoft.NSTV.XMLFileIO.XmlFileIO(OpenFileDialog1.FileName)
            txtFileName.Text = OpenFileDialog1.FileName
            FileName = txtFileName.Text
            btnDescry.Enabled = True
            btnEncry.Enabled = True
        End If
    End Sub

    Private Sub btnEncry_Click(sender As System.Object, e As System.EventArgs) Handles btnEncry.Click

        Try
            Dim strContext As String = Nothing
            Using stm As New StreamReader(FileName, New UTF8Encoding(False))
                Dim encryString As String = stm.ReadToEnd
                strContext = CableSoft.NSTV.Encry.Encryption.Encrypt3DES(encryString, Key)
                stm.Close()
            End Using
            Using stm As New StreamWriter(FileName)
                stm.WriteLine(strContext)
                stm.Flush()
                stm.Close()
            End Using
            MessageBox.Show("加密完成!", "訊息", MessageBoxButtons.OK)
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
        
    End Sub


    Private Sub btnDescry_Click(sender As System.Object, e As System.EventArgs) Handles btnDescry.Click
        Try
            Dim strContext As String = Nothing
            Using stm As New StreamReader(txtFileName.Text)
                strContext = CableSoft.NSTV.Encry.Encryption.Decrypt3DES(stm.ReadToEnd, Key)
                stm.Close()
            End Using
            Using stm As New StreamWriter(FileName)
                stm.Write(strContext)
                stm.Flush()
                stm.Close()
            End Using
            MessageBox.Show("解密完成!", "訊息", MessageBoxButtons.OK)
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub
    Private Sub frmEncry_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        btnEncry.Enabled = False
        btnDescry.Enabled = False
    End Sub
End Class
