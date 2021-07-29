Imports System
Imports System.Data
Imports System.Data.OleDb
Imports DevExpress.Utils
Imports DevExpress.XtraTreeList
Imports DevExpress.XtraEditors
Imports CableSoft.CAS.CryptUtil
Imports CableSoft.Gateway.Common
Imports CableSoft.Gateway.csException

Public Class MDBCommon
    Implements IDisposable

    Private Shared _MDBConnectString1 As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Persist Security Info=True;" & _
                                "Jet OLEDB:Database Password={1}"
    Private Shared _MDBConnectString2 As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Persist Security Info=True;" & _
                                "Jet OLEDB:Database"
    Private Shared _MDBConnectString As String = String.Empty
    Private _MDBPassword As String
    Private _MDBFile As String
    Private Shared _Cn As OleDb.OleDbConnection = Nothing
    Public Sub New()

    End Sub
    Public Sub New(ByVal MDBFile As String, ByVal MDBPassword As String)
        _MDBFile = MDBFile
        _MDBPassword = MDBPassword
        SetConnectSring(_MDBFile, _MDBPassword)
    End Sub
    Public Sub New(ByVal MDBFile As String)
        _MDBFile = MDBFile
        SetConnectSring(_MDBFile, String.Empty)
    End Sub
    Public Function GetConnectString() As String
        Return _MDBConnectString
    End Function
    Public Overloads Function GetMDBCn() As OleDb.OleDbConnection
        Try
            If _Cn IsNot Nothing Then
                _Cn.Close()
                _Cn = Nothing
            End If
            _Cn = New OleDbConnection(_MDBConnectString)
            Return _Cn
        Catch ex As System.Exception
            Throw New System.Exception(ex.Message.ToString)
        End Try
    End Function
    Public Overloads Function GetMDBCn(ByVal aMDBFile As String, ByVal aMDBPassword As String) As OleDb.OleDbConnection
        SetConnectSring(aMDBFile, aMDBPassword)
        GetMDBCn()
        Return _Cn
    End Function
    Public Shared Function ReadMDBDataTable(ByVal aMdbFile As String, _
                                            ByVal aMdbPassword As String, _
                                            ByVal aMDBTable As String, _
                                            Optional ByVal DecryField As Array = Nothing) As DataTable
        Dim retTable As DataTable = Nothing
        Try
            SetConnectSring(aMdbFile, aMdbPassword)

            Using cn As New OleDbConnection(_MDBConnectString)

                Using dap As New OleDbDataAdapter("SELECT * FROM " & aMDBTable, cn)

                    Using ds As New DataSet
                        dap.Fill(ds)
                        retTable = ds.Tables(0)

                        If DecryField IsNot Nothing Then
                            For i As Int32 = 0 To DecryField.Length - 1
                                For iRows As Int32 = 0 To retTable.Rows.Count - 1
                                    retTable.Rows(iRows).BeginEdit()
                                    retTable.Rows(iRows).Item(DecryField.GetValue(i)) = _
                                        CableSoft.CAS.CryptUtil.CryptUtil.Decrypt(retTable.Rows(iRows).Item(DecryField.GetValue(i)))
                                    retTable.Rows(iRows).EndEdit()

                                Next
                            Next
                        End If
                        retTable.AcceptChanges()
                    End Using
                End Using

            End Using
            Return retTable
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Function
    Public Shared Function SaveTreeListToMDB(ByVal aMDBFile As String, _
                                            ByVal aMDBPassword As String, _
                                            ByVal aMDBTable As String, _
                                            ByVal aTreeList As TreeList, _
                                            Optional ByVal EncryField As Array = Nothing) As Boolean
        SetConnectSring(aMDBFile, aMDBPassword)
        Try
            Dim aSQL As String = "SELECT * FROM " & aMDBTable & " WHERE 1=0 "
            Dim aDataRow As DataRow = Nothing
            Dim aValue As String = String.Empty
            Dim lst As List(Of String) = Nothing
            If EncryField IsNot Nothing Then

                lst = New List(Of String)
                For i As Int32 = 0 To EncryField.Length - 1
                    lst.Add(EncryField.GetValue(i).ToString.ToUpper)
                Next
            End If
            Using cn As New OleDbConnection(_MDBConnectString)
                Using dap As New OleDbDataAdapter(aSQL, cn)
                    Using bdu As New OleDbCommandBuilder(dap)
                        dap.InsertCommand = bdu.GetInsertCommand
                        dap.UpdateCommand = bdu.GetUpdateCommand
                        dap.DeleteCommand = bdu.GetDeleteCommand

                        Using ds As New DataSet()
                            dap.Fill(ds)
                            For iNotes As Int32 = 0 To aTreeList.Nodes.Count - 1
                                aDataRow = Nothing

                                aDataRow = ds.Tables(0).NewRow()

                                For cols As Int32 = 0 To aTreeList.Columns.Count - 1
                                    aValue = String.Empty
                                    If ds.Tables(0).Columns.Contains(aTreeList.Columns(cols).FieldName.ToUpper()) Then

                                        If Not IsDBNull(aTreeList.Nodes(iNotes).GetValue(cols)) Then
                                            aValue = aTreeList.Nodes(iNotes).GetValue(cols)
                                        End If

                                        If lst IsNot Nothing Then
                                            If lst.Contains(aTreeList.Columns(cols).FieldName.ToUpper) Then
                                                aValue = CryptUtil.Encrypt(aValue)
                                            End If
                                        End If
                                        If Not String.IsNullOrEmpty(aValue) Then
                                            aDataRow.Item(aTreeList.Columns(cols).FieldName) = aValue

                                        End If

                                    End If
                                Next
                                ds.Tables(0).Rows.Add(aDataRow)
                            Next
                            If cn.State = ConnectionState.Closed Then
                                cn.Open()
                            End If

                            Using tra As OleDbTransaction = cn.BeginTransaction
                                Dim cmd As OleDbCommand = cn.CreateCommand
                                cmd.CommandText = "delete from " & aMDBTable
                                cmd.Transaction = tra
                                dap.UpdateCommand.Transaction = tra
                                dap.InsertCommand.Transaction = tra
                                dap.DeleteCommand.Transaction = tra
                                cmd.ExecuteNonQuery()

                                dap.Update(ds.Tables(0))
                                tra.Commit()
                            End Using
                        End Using
                    End Using
                End Using

            End Using
            Return True
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Shared Function ReadMDBDataToTreeList(ByVal aMDBFile As String, _
                                                 ByVal aMDBPassword As String, _
                                                 ByVal aMDBTable As String, _
                                                 ByVal aTreeList As TreeList, _
                                                 Optional ByVal DecryField As Array = Nothing) As DataTable
        Dim retTable As DataTable = Nothing
        Try
            Dim aFiles As String = "*"
            If aTreeList IsNot Nothing Then
                aFiles = String.Empty
                For i As Int32 = 0 To aTreeList.Columns.Count - 1
                    If String.IsNullOrEmpty(aFiles) Then
                        aFiles = aTreeList.Columns(i).FieldName
                    Else
                        aFiles = aFiles & "," & aTreeList.Columns(i).FieldName
                    End If
                Next
            End If
            Dim aSQL As String = "SELECT " & aFiles & " FROM " & aMDBTable
            SetConnectSring(aMDBFile, aMDBPassword)
            If aTreeList IsNot Nothing Then
                aTreeList.ClearNodes()
            End If

            Using cn As New OleDbConnection(_MDBConnectString)
                Using dap As New OleDbDataAdapter(aSQL, cn)
                    Using ds As New DataSet()
                        dap.Fill(ds)
                        For iRows As Int32 = 0 To ds.Tables(0).Rows.Count - 1
                            If DecryField IsNot Nothing Then
                                For i As Int32 = 0 To DecryField.Length - 1
                                    ds.Tables(0).Rows(iRows).BeginEdit()
                                    ds.Tables(0).Rows(iRows).Item(DecryField.GetValue(i)) = _
                                        CryptUtil.Decrypt(ds.Tables(0).Rows(iRows).Item(DecryField.GetValue(i)))
                                    ds.Tables(0).Rows(iRows).EndEdit()
                                    ds.Tables(0).Rows(iRows).AcceptChanges()
                                Next
                                If aTreeList IsNot Nothing Then
                                    aTreeList.AppendNode(ds.Tables(0).Rows(iRows), iRows)
                                End If
                            Else
                                If aTreeList IsNot Nothing Then
                                    aTreeList.AppendNode(ds.Tables(0).Rows(iRows), iRows)
                                End If
                            End If
                        Next
                        retTable = ds.Tables(0).Copy
                    End Using
                End Using
            End Using
            Return retTable
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function
    


    Public Shared Function SaveMDBDataTable(ByVal aMdbFile As String, _
                                            ByVal aMdbPassword As String, _
                                            ByVal aMDBTable As String, _
                                            ByVal aSrcTable As DataTable, _
                                            Optional ByVal EncryField As Array = Nothing) As Boolean
        Try
            SetConnectSring(aMdbFile, aMdbPassword)
            Using cn As New OleDbConnection(_MDBConnectString)

                Using dap As New OleDbDataAdapter("SELECT * FROM " & aMDBTable, cn)
                    Using bdu As New OleDbCommandBuilder(dap)
                        Using ds As New DataSet()
                            Dim aTabCopy As DataTable = aSrcTable.Copy
                            aTabCopy.TableName = "UPD"
                            aTabCopy.Rows.Clear()
                            Dim rowColect As New List(Of DataRow)
                            If aSrcTable.GetChanges(DataRowState.Added) IsNot Nothing Then
                                For Each rc As DataRow In aSrcTable.GetChanges(DataRowState.Added).Rows
                                    If EncryField IsNot Nothing Then
                                        For i As Int32 = 0 To EncryField.Length - 1
                                            rc.BeginEdit()
                                            rc.Item(EncryField.GetValue(i)) = _
                                                CableSoft.CAS.CryptUtil.CryptUtil.Encrypt(rc.Item(EncryField.GetValue(i)))
                                            rc.EndEdit()
                                        Next
                                    End If
                                    rowColect.Add(rc)
                                Next
                            End If
                            If aSrcTable.GetChanges(DataRowState.Modified) IsNot Nothing Then


                                For Each rc As DataRow In aSrcTable.GetChanges(DataRowState.Modified).Rows
                                    If EncryField IsNot Nothing Then
                                        For i As Int32 = 0 To EncryField.Length - 1
                                            rc.BeginEdit()
                                            rc.Item(EncryField.GetValue(i)) = _
                                                CableSoft.CAS.CryptUtil.CryptUtil.Encrypt(rc.Item(EncryField.GetValue(i)))
                                            rc.EndEdit()
                                        Next
                                    End If
                                    rowColect.Add(rc)
                                Next
                            End If
                            For i As Int32 = 0 To rowColect.Count - 1
                                aTabCopy.ImportRow(rowColect.Item(i))
                            Next

                            ds.Tables.Add(aTabCopy)
                            dap.Fill(ds)
                            dap.Update(ds.Tables("UPD"))
                        End Using
                    End Using
                End Using
            End Using
            Return True
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function
    Public Function CloseMDBCn() As Boolean
        Try
            If _Cn IsNot Nothing Then
                _Cn.Close()
            End If
            _Cn = Nothing
            Return True
        Catch ex As System.Exception
            Throw New System.Exception(ex.Message.ToString)
        End Try
    End Function
    Public Overloads Shared Function ReadMDBData(ByVal aMDBFile As String, ByVal aMDBPassword As String, ByVal aSQL As String) As DataTable
        Try
            SetConnectSring(aMDBFile, aMDBPassword)
            Dim ds As New DataSet()
            Using MdbCn As New OleDbConnection(_MDBConnectString)
                Dim dap As New OleDbDataAdapter(aSQL, MdbCn)
                dap.Fill(ds)
                If MdbCn IsNot Nothing Then
                    MdbCn.Close()
                End If
            End Using
            Return ds.Tables(0).Copy
        Catch ex As Exception
            Throw New Exception(ex.Message.ToString)
        End Try
    End Function

    Public Shared Function SaveMDBData(ByVal aMDBFile As String, ByVal aMDBPassword As String, _
                                       ByVal aScrTable As DataTable, _
                                       ByVal aTableName As String) As Boolean
        SetConnectSring(aMDBFile, aMDBPassword)
        Try
            Using MDBCn As New OleDbConnection(_MDBConnectString)
                Using ds As New DataSet()
                    Using dap As New OleDbDataAdapter("SELECT * FROM " & aTableName, MDBCn)
                        Using bdu As New OleDbCommandBuilder(dap)
                            dap.Fill(ds)
                            Dim aSaveTable As DataTable = aScrTable.Copy
                            aSaveTable.TableName = "UPD"
                            ds.Tables.Add(aSaveTable)
                            dap.Update(aSaveTable)
                        End Using
                    End Using
                End Using
            End Using
            Return True

        Catch ex As Exception
            Throw New Exception(ex.Message.ToString)
        End Try
    End Function

    Private Shared Sub SetConnectSring(ByVal aMDBFile As String, ByVal aMDBPassword As String)
        If String.IsNullOrEmpty(aMDBPassword) Then
            _MDBConnectString = String.Format(_MDBConnectString2, aMDBFile)
        Else
            _MDBConnectString = String.Format(_MDBConnectString1, aMDBFile, aMDBPassword)
        End If
    End Sub


    Private disposedValue As Boolean = False        ' 偵測多餘的呼叫

    ' IDisposable
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: 釋放其他狀態 (Managed 物件)。
                CloseMDBCn()
            End If

            ' TODO: 釋放您自己的狀態 (Unmanaged 物件)
            ' TODO: 將大型欄位設定為 null。
        End If
        Me.disposedValue = True
    End Sub

#Region " IDisposable Support "
    ' 由 Visual Basic 新增此程式碼以正確實作可處置的模式。
    Public Sub Dispose() Implements IDisposable.Dispose
        ' 請勿變更此程式碼。在以上的 Dispose 置入清除程式碼 (ByVal 視為布林值處置)。
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class
Public Enum DATATYPE
    SQL = 0
    TABLE = 1
End Enum