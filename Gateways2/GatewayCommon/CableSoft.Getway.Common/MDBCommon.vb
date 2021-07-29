Imports System
Imports System.Data
Imports System.Data.OleDb
Imports DevExpress.Utils
Imports DevExpress.XtraTreeList
Imports DevExpress.XtraEditors
Imports CableSoft.CAS.CryptUtil
Imports CableSoft.Gateway.Common
Imports CableSoft.Gateway.csException
Imports System.Net

Public Class MDBCommon
    Implements IDisposable
    'Microsoft.ACE.OLEDB.12.0
    'Private Shared _MDBConnectString1 As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Persist Security Info=True;" & _
    '                            "Jet OLEDB:Database Password={1}"
    'Private Shared _MDBConnectString2 As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Persist Security Info=True" 
    Private Shared _MDBConnectString1 As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Persist Security Info=True;" & _
                                "Jet OLEDB:Database Password={1}"
    Private Shared _MDBConnectString2 As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Persist Security Info=True"
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
    '取得所有Schema欄位
    Private Shared Function GetSchemaField(ByVal aMdbFile As String, _
                                           ByVal aMdbPassword As String, _
                                           ByVal aMDBTable As String) As Array
        Try
            Dim aryField() As String = Nothing
            Dim aSchemaTable As DataTable = Nothing
            SetConnectSring(aMdbFile, aMdbPassword)

            Using cn As New OleDbConnection(_MDBConnectString)
                cn.Open()
                Using dap As New OleDbDataAdapter("SELECT * FROM " & aMDBTable & " WHERE 1=0 ", cn)
                    Using ds As New DataSet()
                        dap.Fill(ds)
                        aSchemaTable = ds.Tables(0)
                    End Using
                End Using

                If aSchemaTable.Columns.Count > 0 Then
                    Array.Resize(Of String)(aryField, aSchemaTable.Columns.Count)
                    For i As Int32 = 0 To aSchemaTable.Columns.Count - 1
                        aryField(i) = aSchemaTable.Columns(i).ColumnName
                    Next
                End If

            End Using
            Return aryField
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function


    '將MDB的資料解密
    Public Overloads Shared Function ReadMDBDataTable(ByVal aMdbFile As String, _
                                            ByVal aMdbPassword As String, _
                                            ByVal aMDBTable As String) As DataTable


        Try

            Dim aryField() As String = GetSchemaField(aMdbFile, aMdbPassword, aMDBTable)


            Return ReadMDBDataTable(aMdbFile, aMdbPassword, aMDBTable, aryField)



        Catch ex As Exception
            Throw New Exception(ex.Message)


        End Try
    End Function


    Public Overloads Shared Function ReadMDBDataTable(ByVal aMdbFile As String, _
                                            ByVal aMdbPassword As String, _
                                            ByVal aMDBTable As String, _
                                            ByVal DecryField As Array) As DataTable
        Dim retTable As DataTable = Nothing
        Try
            SetConnectSring(aMdbFile, aMdbPassword)

            Using cn As New OleDbConnection(_MDBConnectString)

                Using dap As New OleDbDataAdapter("SELECT * FROM " & aMDBTable, cn)

                    Using ds As New DataSet
                        dap.Fill(ds)
                        retTable = ds.Tables(0)
                        dap.FillSchema(retTable, SchemaType.Mapped)
                        
                        If DecryField IsNot Nothing Then
                            For i As Int32 = 0 To DecryField.Length - 1
                                If (Not retTable.Columns(DecryField.GetValue(i).ToString).AutoIncrement) OrElse _
                                    (Not retTable.Columns(DecryField.GetValue(i).ToString).Unique) Then
                                    For iRows As Int32 = 0 To retTable.Rows.Count - 1
                                        If Not String.IsNullOrEmpty(retTable.Rows(iRows).Item(DecryField.GetValue(i)).ToString) Then
                                            retTable.Rows(iRows).BeginEdit()
                                            retTable.Rows(iRows).Item(DecryField.GetValue(i)) = _
                                                CableSoft.CAS.CryptUtil.CryptUtil.Decrypt(retTable.Rows(iRows).Item(DecryField.GetValue(i)))
                                            retTable.Rows(iRows).EndEdit()
                                        End If
                                    Next

                                End If

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
    Public Overloads Shared Function SaveTreeListToMDB(ByVal aMDBFile As String, _
                                            ByVal aMDBPassword As String, _
                                            ByVal aMDBTable As String, _
                                            ByVal aTreeList As TreeList) As Boolean
        Try
            Dim aryField() As String = GetSchemaField(aMDBFile, aMDBPassword, aMDBTable)
            Return SaveTreeListToMDB(aMDBFile, aMDBPassword, aMDBTable, aTreeList, aryField)

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Function

    Public Overloads Shared Function SaveTreeListToMDB(ByVal aMDBFile As String, _
                                            ByVal aMDBPassword As String, _
                                            ByVal aMDBTable As String, _
                                            ByVal aTreeList As TreeList, _
                                            ByVal EncryField As Array) As Boolean
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
                            'dap.FillSchema(ds.Tables(0), SchemaType.Mapped)
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
    Public Overloads Shared Function ReadMDBDataToTreeList(ByVal aMDBFile As String, _
                                                     ByVal aMDBPassword As String, _
                                                     ByVal aMDBTable As String, _
                                                     ByVal aTreeList As TreeList) As DataTable
        Try
            Dim aryField() As String = GetSchemaField(aMDBFile, aMDBPassword, aMDBTable)
            Return ReadMDBDataToTreeList(aMDBFile, aMDBPassword, aMDBTable, aTreeList, aryField)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function


    Public Overloads Shared Function ReadMDBDataToTreeList(ByVal aMDBFile As String, _
                                                 ByVal aMDBPassword As String, _
                                                 ByVal aMDBTable As String, _
                                                 ByVal aTreeList As TreeList, _
                                                 ByVal DecryField As Array) As DataTable
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
                                    If ds.Tables(0).Columns.Contains(DecryField.GetValue(i)) Then
                                        If Not IsDBNull(ds.Tables(0).Rows(iRows).Item(DecryField.GetValue(i))) Then
                                            ds.Tables(0).Rows(iRows).BeginEdit()
                                            ds.Tables(0).Rows(iRows).Item(DecryField.GetValue(i)) = _
                                                CryptUtil.Decrypt(ds.Tables(0).Rows(iRows).Item(DecryField.GetValue(i)))
                                            ds.Tables(0).Rows(iRows).EndEdit()
                                            ds.Tables(0).Rows(iRows).AcceptChanges()
                                        End If
                                        
                                    End If
                                    
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

    Public Shared Function GetAllTableName(ByVal aMDBFile As String, ByVal aMdbPassword As String) As List(Of String)
        Try
            Dim aLst As New List(Of String)
            SetConnectSring(aMDBFile, aMdbPassword)
            Using cn As New OleDbConnection(_MDBConnectString)
                cn.Open()
                Dim dtTableName As DataTable = cn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, _
                                                                    New Object() {Nothing, Nothing, Nothing, "TABLE"})
                For i As Int32 = 0 To dtTableName.Rows.Count - 1
                    aLst.Add(dtTableName.Rows(i)("TABLE_NAME"))
                Next
                If cn.State = ConnectionState.Open Then
                    cn.Close()
                End If
            End Using
            Return aLst
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function
    '將所有要儲存的資料加密
    Public Overloads Shared Function SaveMDBDataTable(ByVal aMdbFile As String, _
                                            ByVal aMdbPassword As String, _
                                            ByVal aMDBTable As String, _
                                            ByVal aSrcTable As DataTable) As Boolean
        Try
            Dim aryField As Array = GetSchemaField(aMdbFile, aMdbPassword, aMDBTable)
            Return SaveMDBDataTable(aMdbFile, aMdbPassword, aMDBTable, aSrcTable, aryField)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function


    Public Overloads Shared Function SaveMDBDataTable(ByVal aMdbFile As String, _
                                            ByVal aMdbPassword As String, _
                                            ByVal aMDBTable As String, _
                                            ByVal aSrcTable As DataTable, _
                                             ByVal EncryField As Array) As Boolean
        Try
            SetConnectSring(aMdbFile, aMdbPassword)
            Using cn As New OleDbConnection(_MDBConnectString)

                Using dap As New OleDbDataAdapter("SELECT * FROM " & aMDBTable, cn)
                    AddHandler dap.RowUpdated, New OleDbRowUpdatedEventHandler(AddressOf OnRowUpdated)

                    Using bdu As New OleDbCommandBuilder(dap)
                        Using ds As New DataSet()
                            Dim aTabCopy As DataTable = aSrcTable.Copy
                            aTabCopy.TableName = "UPD"
                            aTabCopy.Rows.Clear()
                            dap.FillSchema(aTabCopy, SchemaType.Mapped)
                            Dim rowColect As New List(Of DataRow)
                            If aSrcTable.GetChanges(DataRowState.Added) IsNot Nothing Then
                                For Each rc As DataRow In aSrcTable.GetChanges(DataRowState.Added).Rows
                                    If EncryField IsNot Nothing Then
                                        For i As Int32 = 0 To EncryField.Length - 1
                                            If (Not aTabCopy.Columns(EncryField.GetValue(i).ToString).AutoIncrement) OrElse _
                                                (Not aTabCopy.Columns(EncryField.GetValue(i).ToString).Unique) Then

                                                If Not IsDBNull(rc.Item(i)) Then
                                                    rc.BeginEdit()
                                                    rc.Item(EncryField.GetValue(i)) = _
                                                        CableSoft.CAS.CryptUtil.CryptUtil.Encrypt(rc.Item(EncryField.GetValue(i)))
                                                    rc.EndEdit()
                                                End If



                                            End If
                                            

                                        Next
                                    End If
                                    rowColect.Add(rc)
                                Next
                            End If
                            If aSrcTable.GetChanges(DataRowState.Modified) IsNot Nothing Then


                                For Each rc As DataRow In aSrcTable.GetChanges(DataRowState.Modified).Rows
                                    If EncryField IsNot Nothing Then
                                        For i As Int32 = 0 To EncryField.Length - 1
                                            If (Not aTabCopy.Columns(EncryField.GetValue(i).ToString).AutoIncrement) OrElse _
                                                (Not aTabCopy.Columns(EncryField.GetValue(i).ToString).Unique) Then
                                                If Not IsDBNull(rc.Item(EncryField.GetValue(i))) Then
                                                    rc.BeginEdit()
                                                    rc.Item(EncryField.GetValue(i)) = _
                                                        CableSoft.CAS.CryptUtil.CryptUtil.Encrypt(rc.Item(EncryField.GetValue(i)))
                                                    rc.EndEdit()
                                                End If
                                                
                                            End If
                                           
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
    Private Shared Sub OnRowUpdated(ByVal sender As Object, ByVal args As OleDbRowUpdatedEventArgs)
        If args.RecordsAffected = 0 Then
            args.Row.RowError = "Optimistic Concurrency Violation!"
            args.Status = UpdateStatus.SkipCurrentRow
            WriteErrTxtLog.WriteTxtError(args.Errors, "儲存資料失敗")
        End If
    End Sub

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
                            dap.Update(ds.Tables("UPD"))
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