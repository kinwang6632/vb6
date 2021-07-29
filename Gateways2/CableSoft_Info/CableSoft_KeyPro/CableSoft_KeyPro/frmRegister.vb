Imports System
Imports System.Windows
Imports System.Windows.Forms
Imports System.IO
Imports System.Text
Public Class frmRegister
    Private _RegisterOK As Boolean = False
    Public Property ExeName As String = Nothing
    Public ReadOnly Property RegisterOK() As Boolean
        Get
            Return _RegisterOK
        End Get
        
    End Property
    Private Sub frmRegister_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        txtSoftSN.Text = CableSoft_KeyPro.GetSystemInfo.GetAllInfo(True, ExeName)
    End Sub

    Private Sub btnCopy_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCopy.Click
        Clipboard.SetText(txtSoftSN.Text)
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim bty() As Byte = Encoding.Default.GetBytes(txtRegSN.Text)
        Dim aFile As String = String.Empty
        aFile = Application.StartupPath
        If Not aFile.EndsWith("\") Then
            aFile = aFile & "\"
        End If
        aFile = aFile & GetSystemInfo.FKeyProName
        Using stm As New MemoryStream(bty)
            If GetSystemInfo.IsRegister(stm, ExeName) Then
                Using obj As New StreamWriter(aFile, False)
                    obj.WriteLine(txtRegSN.Text)
                    obj.Flush()
                    obj.Close()
                End Using
                _RegisterOK = True
                Me.Close()
            Else

                _RegisterOK = False
                Me.Close()

            End If
        End Using
    End Sub

    Private Sub txtRegSN_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRegSN.TextChanged
        If String.IsNullOrEmpty(txtRegSN.Text) Then
            btnOK.Enabled = False
        Else
            btnOK.Enabled = True
        End If
    End Sub

    Private Sub btnNO_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNO.Click
        Me.Close()
    End Sub
End Class