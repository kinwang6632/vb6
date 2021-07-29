Public Class XfrmAbout 

    Private Sub grpCtl_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles grpCtl.MouseDown
        Me.Close()
    End Sub

    Private Sub XfrmAbout_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        lblCopyright.Text = lblCopyright.Text & Environment.NewLine & Environment.NewLine & _
                    "                   Validity Date " & FInsDate & " - " & FRegDate
    End Sub
End Class