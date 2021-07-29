Imports System.Data.OleDb
Public Class Form1

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'Dim cnn As New OleDbConnection(My.Settings.Setting)
        'Dim datApd As New OleDbDataAdapter()
        'Dim dst As New DataSet

        'datApd.SelectCommand = New OleDbCommand("Select * From SO001", cnn)
        'cnn.Open()
        'datApd.Fill(dst, "SO001")
        'dst.Tables("SO001").Rows(0).Item("CustName") = "王淳平"
        'datApd.Update(dst)
        'MsgBox("完成")
        Using connection As New OleDbConnection(My.Settings.Setting)
            Dim adapter As New OleDbDataAdapter()
            adapter.SelectCommand = New OleDbCommand("Select * From SO001", connection)
            Dim builder As OleDbCommandBuilder = New OleDbCommandBuilder(adapter)

            connection.Open()

            Dim customers As DataSet = New DataSet
            adapter.Fill(customers)

            customers.Tables(0).Rows(0).Item("CustName") = "王淳平"

            adapter.Update(customers)
            MsgBox("完成")

        End Using




    End Sub
End Class
