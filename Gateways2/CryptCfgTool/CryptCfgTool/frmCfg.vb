Imports CableSoft.Gateway.Common

Imports CableSoft.Gateway.csException
Imports CableSoft.CAS.CryptUtil
Imports System.Data.OleDb
Imports System.IO
Public Class frmCryptCfg
    Private fCfgName As String = String.Empty

    Private flstTableName As List(Of String) = Nothing
    Private fCn As OleDbConnection = Nothing
    Private fDs As DataSet = Nothing
    Private fTra As OleDbTransaction = Nothing
    Private Sub btnPath_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPath.Click
        Try
            OpenFileDialog1.DefaultExt = "CFG"
            OpenFileDialog1.Filter = "設定檔|*.CFG"
            OpenFileDialog1.Title = "選擇設定檔"
            If OpenFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
                fCfgName = OpenFileDialog1.FileName
            End If
            txtFileName.Text = fCfgName
            flstTableName = MDBCommon.GetAllTableName(fCfgName, fMdbPassword)
            For i As Int32 = 0 To flstTableName.Count - 1
                lsvTable.Items.Add(flstTableName.Item(i)).Checked = True
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            OpenFileDialog1.Dispose()
        End Try

    End Sub


    Private Sub btnDecry_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDecry.Click        
        Try
            If fCn Is Nothing Then
                fCn = New OleDbConnection(String.Format(fMDBConnectString, txtFileName.Text, fMdbPassword))
            End If
            If fDs Is Nothing Then
                fDs = New DataSet()
            End If
            fCn.Open()
            fTra = fCn.BeginTransaction
            For i As Int32 = 0 To flstTableName.Count - 1
                Using aDp As New OleDbDataAdapter("select * from " & flstTableName.Item(i), fCn)
                    Using aBdu As New OleDbCommandBuilder(aDp)
                        aDp.SelectCommand.Transaction = fTra
                        aDp.DeleteCommand = aBdu.GetDeleteCommand
                        aDp.InsertCommand = aBdu.GetInsertCommand
                        aDp.UpdateCommand = aBdu.GetUpdateCommand
                        aDp.DeleteCommand.Transaction = fTra
                        aDp.UpdateCommand.Transaction = fTra
                        aDp.InsertCommand.Transaction = fTra
                    End Using
                    aDp.Fill(fDs, flstTableName.Item(i))
                    For j As Int32 = 0 To fDs.Tables(flstTableName.Item(i)).Rows.Count - 1
                        With fDs.Tables(flstTableName.Item(i)).Rows(j)
                            .BeginEdit()
                            For k As Integer = 0 To fDs.Tables(flstTableName.Item(i)).Columns.Count - 1
                                If fDs.Tables(flstTableName.Item(i)).Columns(k).AutoIncrement Then

                                End If
                            Next
                            .EndEdit()
                        End With
                    Next
                    aDp.Update(fDs)
                End Using
                fTra.Commit()

                'Dim aTb As DataTable = MDBCommon.ReadMDBDataTable(fCfgName, fMdbPassword, flstTableName.Item(i))


                'MDBCommon.SaveMDBDataTable()
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            If fTra IsNot Nothing Then
                fTra.Dispose()
            End If
            If fCn IsNot Nothing Then
                fCn.Close()
                fCn.Dispose()
            End If
            If fDs IsNot Nothing Then
                fDs.Dispose()
            End If
        End Try
    End Sub
End Class
