Imports System
Imports System.Data
Imports System.Data.OracleClient
Imports System.Collections.Generic
Imports System.IO
Imports System.Text
Imports System.Windows
Imports System.Windows.Forms
Public NotInheritable Class GetwaySys

    Private _AppTitle As String

    Private _ConnectString As String = "Data Source={0};Persist Security Info=True;" & _
                "User ID={1};Password={2};Min Pool Size={3};" & _
                "Max Pool Size={4};Unicode=True"
    Private _MDBFile As String
    Private _MDBPassword As String
    Private _AutoRunGW As Boolean
    Private _UseTray As Boolean
    Private _DBPool As Int32 = 100
    Private Const FSysTable As String = "SysOpt"
    Private Const FDatabase As String = "SO"
    Private Const FNagraParaSetName As String = "NagraParaSet"
    Private Const FComparedErrorName As String = "ComparedError"
    Public Sub New(ByVal aMDBFile As String, ByVal aMdbPassword As String)
        _MDBFile = aMDBFile
        _MDBPassword = aMdbPassword

    End Sub
    Public Property AppTitle() As String
        Get
            Return _AppTitle
        End Get
        Set(ByVal value As String)
            _AppTitle = value
        End Set
    End Property
    Public Property AutoRunGW() As Boolean
        Get
            Return _AutoRunGW
        End Get
        Set(ByVal value As Boolean)
            _AutoRunGW = value
        End Set
    End Property
    Public Property UseTray() As Boolean
        Get
            Return _UseTray
        End Get
        Set(ByVal value As Boolean)
            _UseTray = value
        End Set
    End Property
    Public Property DBPool() As Int32
        Get
            Return _DBPool
        End Get
        Set(ByVal value As Int32)
            _DBPool = value
        End Set
    End Property
    Public Shared Function ReadErrorCode(ByVal aMDBFile As String, ByVal aMDBPassword As String)
        Try
            Return CableSoft.GateWay.Common.MDBCommon.ReadMDBDataTable(aMDBFile, aMDBPassword, FComparedErrorName)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function
    Public Shared Function ReadSystemIO(ByVal aMDBFile As String, ByVal aMDBPassword As String) As DataTable
        Try
            Return CableSoft.GateWay.Common.MDBCommon.ReadMDBDataTable(aMDBFile, aMDBPassword, FSysTable)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function
    Public Shared Function ReadSO(ByVal aMDBFile As String, ByVal aMDBPassword As String) As DataTable
        Try
            Return CableSoft.GateWay.Common.MDBCommon.ReadMDBDataTable(aMDBFile, aMDBPassword, FDatabase, New String() {"UserPassword"})
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Function
    Public Shared Function ReadNagraParaSet(ByVal aMDBFile As String, ByVal aMDBPassword As String) As DataTable
        Try
            Return CableSoft.GateWay.Common.MDBCommon.ReadMDBDataTable(aMDBFile, aMDBPassword, FNagraParaSetName)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function
    Public Shared Function ReadComparedError(ByVal aMDBFile As String, ByVal aMDBPassword As String) As DataTable
        Try
            Return CableSoft.GateWay.Common.MDBCommon.ReadMDBDataTable(aMDBFile, aMDBPassword, FComparedErrorName)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function
    Private Function ReadDataCrypt(ByVal aContext As String, ByVal blnEncrypt As Boolean) As String
        If blnEncrypt Then
            Return CableSoft.CAS.CryptUtil.CryptUtil.Decrypt(aContext)
        Else
            Return aContext

        End If
    End Function
    Public Shared Function DefaultNagra(ByVal aMDBFile As String, ByVal aMDBPassword As String, _
                                        ByVal aMDBTable As String, _
                                        Optional ByVal aDescrField As Array = Nothing) As Boolean
        Try
            Dim aDataTable As DataTable = CableSoft.GateWay.Common.MDBCommon.ReadMDBDataTable(aMDBFile, aMDBPassword, aMDBTable, Nothing)
            Dim rw As DataRow = aDataTable.NewRow
            rw("NagraIPAddress") = "172.17.64.1"
            rw("NagraPort") = 60002
            rw("SndSourceId") = "100"
            rw("SndDestId") = "2"
            rw("SndMopPPId") = "00545"
            rw("RcvSourceId") = "100"
            rw("RcvDestId") = "2"
            rw("RcvMopPPId") = "00545"
            rw("SenErrStopCnt") = 3
            rw("RcvErrStopCnt") = 3
            rw("CheckStatusTime") = 3
            rw("SndDelayTime") = 0.1
            rw("SndTimeout") = 30
            rw("RcvTimeout") = 30
            aDataTable.Rows.Add(rw)
            CableSoft.GateWay.Common.MDBCommon.SaveMDBDataTable(aMDBFile, aMDBPassword, aMDBTable, aDataTable)
            Return True
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function
    Public Shared Function DefaultSystem(ByVal aMDBFile As String, ByVal aMDBPassword As String, _
                                  ByVal aMDBTable As String, _
                                  Optional ByVal aDecryField As Array = Nothing) As Boolean
        Try
            Dim aDataTable As DataTable = CableSoft.Gateway.Common.MDBCommon.ReadMDBDataTable(aMDBFile, aMDBPassword, aMDBTable, Nothing)
            Dim rw As DataRow = aDataTable.NewRow
            rw("Appcaption") = "CableSoft GateWaay"
            rw("AutoRunGw") = 0
            rw("UseTray") = 0
            rw("ShowResource") = 0
            rw("UseHotKey") = 0
            rw("AutoConnect") = 100
            rw("ReadData") = 100
            rw("ProcessNumber") = 100
            rw("ShowDataCount") = 200
            rw("ClearDataCount") = 20
            rw("MaxThread") = 25
            rw("DBPool") = 100
            rw("TestSocketTime") = 30
            rw("NagraIPAddress") = "127.0.0.1"
            rw("NagraPort") = 8080
            aDataTable.Rows.Add(rw)
            CableSoft.Gateway.Common.MDBCommon.SaveMDBDataTable(aMDBFile, aMDBPassword, aMDBTable, aDataTable)
            Return True
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try



    End Function
    Public Shared Function EncryptData(ByVal aContext As String) As String
        Return CableSoft.CAS.CryptUtil.CryptUtil.Encrypt(aContext)
    End Function
    Public Function GetOracleCns(ByVal aMDBFile As String, ByVal aMDBPassword As String) As List(Of OracleConnection)

        Return Nothing
    End Function
End Class
