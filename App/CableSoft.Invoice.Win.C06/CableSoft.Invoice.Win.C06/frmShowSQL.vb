Public Class frmShowSQL 
    Private _SQLText As String
    Public Property SQLText()
        Get
            Return _SQLText
        End Get
        Set(ByVal value)
            _SQLText = value
        End Set
    End Property

    Private Sub frmShowSQL_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        edtShowSQL.Text = _SQLText
    End Sub
End Class