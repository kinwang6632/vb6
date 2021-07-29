Imports System
Imports System.IO
Imports System.Data
Imports CableSoft.Gateway.Common
Imports System.Data.OleDb
Public Class frmMoveSet




    Private flstOldTableName As List(Of String) = Nothing
    Private flstNewTableName As List(Of String) = Nothing

    Private fOldConnectString As String = String.Empty
    Private fNewConnectString As String = String.Empty
    Private fOldCn As OleDbConnection = Nothing
    Private fNewCn As OleDbConnection = Nothing
    Private fNewDs As DataSet = Nothing
    Private fTraNew As OleDbTransaction = Nothing
    Private Sub btnOldPath_Click(sender As System.Object, e As System.EventArgs) Handles btnOldPath.Click
        Try
            OpenFileDialog1.DefaultExt = "CFG"
            OpenFileDialog1.Filter = "設定檔(新)|*.CFG|所有檔案 (*.*)|*.*"
            OpenFileDialog1.Title = "選擇設定檔"
            If OpenFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
                txtOldFileName.Text = OpenFileDialog1.FileName
            End If


        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            OpenFileDialog1.Dispose()
        End Try
    End Sub
    Private Function IsDataOK() As Boolean
        Try
            If String.IsNullOrEmpty(txtOldFileName.Text) Then
                MessageBox.Show("請選擇 [ 舊設定檔案 ]", "警告", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return False
            End If
            If String.IsNullOrEmpty(txtNewFileName.Text) Then
                MessageBox.Show("請選擇 [ 新設定檔案 ]", "警告", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return False
            End If
            If txtNewFileName.Text = txtOldFileName.Text Then
                MessageBox.Show("設定檔不能同一個，請重新選擇", "警告", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return False
            End If
            If Not File.Exists(txtOldFileName.Text) Then
                MessageBox.Show("舊定檔不存在，請重新選擇", "警告", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return False
            End If
            If Not File.Exists(txtNewFileName.Text) Then
                MessageBox.Show("新定檔不存在，請重新選擇", "警告", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return False
            End If
            Return True
        Catch ex As Exception

            MessageBox.Show(ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function
    Private Sub btnNewPath_Click(sender As System.Object, e As System.EventArgs) Handles btnNewPath.Click
        Try
            OpenFileDialog1.DefaultExt = "CFG"
            OpenFileDialog1.Filter = "設定檔(新)|*.CFG|所有檔案 (*.*)|*.*"
            OpenFileDialog1.Title = "選擇設定檔"
            If OpenFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
                txtNewFileName.Text = OpenFileDialog1.FileName
            End If


        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            OpenFileDialog1.Dispose()
        End Try
    End Sub
    Private Function ClearTable(ByRef aCn As OleDbConnection, ByVal aTableName As String) As Boolean
        Try
            If aCn.State = ConnectionState.Closed Then
                aCn.Open()
            End If
            Using cmd As OleDbCommand = aCn.CreateCommand
                cmd.Transaction = fTraNew
                cmd.CommandText = "delete Table from " & aTableName
                cmd.ExecuteNonQuery()
            End Using
            Return True
        Catch ex As Exception
            MessageBox.Show(ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function
    Private Function CopyTableData(ByRef aNewCn As OleDbConnection, ByRef aOldCn As OleDbConnection,
                                    ByVal aTableName As String)
        Dim aOldTbl As DataTable = New DataTable(aTableName)
        Dim aNewTbl As DataTable = New DataTable(aTableName)
        If aNewCn.State = ConnectionState.Closed Then
            aNewCn.Open()
        End If
        If aOldCn.State = ConnectionState.Closed Then
            aOldCn.Open()
        End If
        Try
            Using aOldDp As New OleDbDataAdapter("select * from " & aTableName, aOldCn)
                aOldDp.Fill(aOldTbl)
            End Using
            Using aNewDp As New OleDbDataAdapter("select * from  " & aTableName, aNewCn)
                Using aNewBduSQL As New OleDbCommandBuilder(aNewDp)
                    aNewDp.SelectCommand.Transaction = fTraNew
                    aNewDp.InsertCommand = aNewBduSQL.GetInsertCommand()
                    aNewDp.UpdateCommand = aNewBduSQL.GetUpdateCommand()

                    aNewDp.InsertCommand.Transaction = fTraNew
                    aNewDp.UpdateCommand.Transaction = fTraNew

                    aNewDp.Fill(fNewDs, aTableName)
                    For i As Int32 = 0 To aOldTbl.Rows.Count - 1
                        Dim rw As DataRow = fNewDs.Tables(aTableName).NewRow
                        For j As Int32 = 0 To aOldTbl.Columns.Count - 1
                            rw(aOldTbl.Columns(j).ColumnName) = aOldTbl.Rows(i).Item(j)
                        Next
                        fNewDs.Tables(aTableName).Rows.Add(rw)
                    Next
                    aNewDp.Update(fNewDs, aTableName)
                End Using
                
            End Using

            Return True
        Catch ex As Exception
            MessageBox.Show(ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        Finally
            aOldTbl.Dispose()
            aNewTbl.Dispose()
        End Try
    End Function


    Private Sub btnOK_Click(sender As System.Object, e As System.EventArgs) Handles btnOK.Click
        Try
            Me.Cursor = Cursors.WaitCursor
            If Not IsDataOK() Then
                Exit Sub
            End If
            fOldConnectString = String.Format(fMDBConnectString, txtOldFileName.Text, fMdbPassword)
            fNewConnectString = String.Format(fMDBConnectString, txtNewFileName.Text, fMdbPassword)
            flstOldTableName = MDBCommon.GetAllTableName(txtOldFileName.Text, fMdbPassword)
            flstNewTableName = MDBCommon.GetAllTableName(txtNewFileName.Text, fMdbPassword)
            fOldCn = New OleDbConnection(fOldConnectString)
            fNewCn = New OleDbConnection(fNewConnectString)
            fOldCn.Open()
            fNewCn.Open()
            fTraNew = fNewCn.BeginTransaction
            If fNewDs Is Nothing Then
                fNewDs = New DataSet()
            End If
            Try
                For Each aTableName As String In flstOldTableName
                    If flstNewTableName.Contains(aTableName, StringComparer.OrdinalIgnoreCase) Then
                        If Not ClearTable(fNewCn, aTableName) Then
                            fTraNew.Rollback()
                            Exit Sub
                        End If
                        If Not CopyTableData(fNewCn, fOldCn, aTableName) Then
                            fTraNew.Rollback()
                            Exit Sub
                        End If
                    End If
                Next
                fTraNew.Commit()
                MessageBox.Show("複制完成", "訊息", MessageBoxButtons.OK)
            Catch ex As Exception
                If fTraNew IsNot Nothing Then
                    fTraNew.Rollback()
                End If

            End Try


        Catch ex As Exception
            MessageBox.Show(ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        Finally
            If fOldCn IsNot Nothing Then
                fOldCn.Close()
                fOldCn.Dispose()
            End If
            If fNewCn IsNot Nothing Then
                fNewCn.Close()
                fNewCn.Dispose()
            End If
            If fNewDs IsNot Nothing Then
                fNewDs.Dispose()
            End If
            Me.Cursor = Cursors.Default
        End Try
    End Sub






    
End Class