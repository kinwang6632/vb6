Imports System
Imports System.Net
Imports System.Text
Imports System.IO
Imports System.Net.Sockets
Imports CableSoft.Gateway.csException
Imports System.Runtime.InteropServices.Marshal
Imports System.Configuration
Imports System.Collections
Imports System.Collections.Generic
Public Class NagraCommand
    
    Private Const Fob_Name As String = "SMS_PVOD"
    Private Shared _SourceId As String = "0000"
    Private Shared _Dest_Id As String = "0000"
    Private Shared _MOP_PPID As String = "00000"
    Private Shared _Transaction_number As Int32
    Private Shared _ExeConfigName As String = String.Empty
    Private Shared _ErrCodeLst As Dictionary(Of String, String)
    Private Shared lckObject As New Object
    Public Shared Property SourceId() As String
        Get
            Return _SourceId
        End Get
        Set(ByVal value As String)
            _SourceId = value
            _SourceId = "0000" & _SourceId
            _SourceId = Right(_SourceId, 4)
        End Set
    End Property
    Public Shared Property Dest_Id() As String
        Get
            Return _Dest_Id
        End Get
        Set(ByVal value As String)
            _Dest_Id = value
            _Dest_Id = "00000" & _Dest_Id
            _Dest_Id = Right(_Dest_Id, 4)
        End Set
    End Property
    Public Shared Property MOP_PPID() As String
        Get
            Return _MOP_PPID
        End Get
        Set(ByVal value As String)
            _MOP_PPID = value
            _MOP_PPID = "00000" & _MOP_PPID
            _MOP_PPID = Right(_MOP_PPID, 5)
        End Set
    End Property
    Public Shared Property Transaction_number() As Int32
        Get
            Return _Transaction_number
        End Get
        Set(ByVal value As Int32)
            _Transaction_number = value
        End Set
    End Property
    Public Shared Property ExeConfigName() As String
        Get
            Return _ExeConfigName
        End Get
        Set(ByVal value As String)
            _ExeConfigName = value
        End Set
    End Property
    Public Shared Property ErrCodeLst() As Dictionary(Of String, String)
        Get
            Return _ErrCodeLst
        End Get
        Set(ByVal value As Dictionary(Of String, String))
            _ErrCodeLst = value
        End Set
    End Property
    Public Shared Function BuilderCmd(ByVal aCmdType As CmdType) As String
        Dim aRet As String = String.Empty
        Try
            Select Case aCmdType
                Case CmdType.ConfirmConnect
                    aRet = ConfirmConnect()
                Case CmdType.Cmd1002
                    Return GetCMD1002()
            End Select
            Return aRet
        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
            Return aRet
        End Try
    End Function

    Public Shared Function NagraStatus(ByVal aByte() As Byte, ByVal aType As CmdType) As Object
        Try

            Select Case aType
                Case CmdType.ConfirmConnect
                    Return GetConfirmtStatus(aByte)
                Case CmdType.Cmd1002
                    Return GetCmd1002Status(aByte)
            End Select
        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)

        End Try
    End Function
    '取得連接狀態
    Private Shared Function GetConfirmtStatus(ByVal aByte() As Byte) As NagraStatus_Type
        Try
            Dim obj As New NagraStatus_Type
            obj.CmdType = CmdType.ConfirmConnect
            obj.SUCCESS = aByte(2)
            obj.answer_code = aByte(5)
            Return obj
        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
        End Try
    End Function
    Private Shared Function GetCmd1002Status(ByVal aByte() As Byte) As NagraCmdStatus_Type
        Dim objRet As New NagraCmdStatus_Type
        objRet.Status = AckCode.Nack
        objRet.ErrorCode = -9998
        objRet.ErrorName = "初始化狀態"

        Try


            '0E000000003051234200056789201010081000200000001000000000000000000000000
            '前面2個為總長度
            Dim aLen As Int32 = Convert.ToInt32(aByte(0).ToString & aByte(1).ToString)
            Dim aString As String = Encoding.ASCII.GetString(aByte, 2, aByte.Length - 2).TrimEnd(Chr(0))
            If aString.Length <> aLen Then
                objRet.Status = AckCode.Nack

                objRet.ErrorCode = -9999
                objRet.ErrorName = "回傳長度比對錯誤"
            Else
                If aString.Substring(32, 4) = "1000" Then
                    objRet.Status = AckCode.Ack
                Else
                    objRet.Status = AckCode.Nack
                    objRet.ErrorCode = aString.Substring(46, 4)
                    objRet.ErrorName = aString.Substring(50, 4)
                End If
            End If
            Return objRet
        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
            Return objRet
        End Try
    End Function
    Private Shared Function ConfirmConnect() As String
        Try
            Dim ob_name_len As Byte = Encoding.ASCII.GetByteCount(Fob_Name)
            Dim op_mode As Byte = 0
            Dim Len1 As Byte = 0
            Dim Len2 As Byte = SizeOf(op_mode) + SizeOf(ob_name_len) + Convert.ToInt32(ob_name_len)
            Dim aRet As String = Chr(Len1) & Chr(Len2) & Chr(op_mode) & Chr(ob_name_len) & Fob_Name
            Return aRet
        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
            Return String.Empty
        End Try
    End Function
    Private Shared Function GetCMD1002() As String
        Try
            SyncLock lckObject
                Dim aRet As String = GetHeadString(CmdType.Cmd1002) & GetSectionString(CmdType.Cmd1002) & GetBodyString(CmdType.Cmd1002)
                Dim aTotalLen As String = Right("00" & aRet.Length, 4)
                Dim aLen1 As Int32 = Convert.ToInt32(Left(aTotalLen, 2))
                Dim aLen2 As Int32 = Convert.ToInt32(Right(aTotalLen, 2))
                aRet = Chr(aLen1) & Chr(aLen2) & aRet
                Return aRet
            End SyncLock
        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
            Return String.Empty
        End Try
    End Function
    Private Shared Function GetHeadString(ByVal aType As CmdType) As String
        Try

            Dim transaction_number As String = String.Empty
            Dim command_type As String = "05"
            If aType = CmdType.Cmd0851 Then
                command_type = "08"
            End If

            Dim source_id As String = _SourceId
            Dim dest_id As String = _Dest_Id
            Dim MOP_PPID As String = _MOP_PPID
            Dim creation_date As String = Format(Now, "yyyyMMdd")
            transaction_number = Right("000000000" & Next_transaction_number.ToString, 9)
            Return transaction_number & command_type & source_id & _
                dest_id & MOP_PPID & creation_date

        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
            Return String.Empty
        End Try
    End Function
    Private Shared Function GetSectionString(ByVal aType As CmdType)

        Return String.Empty

    End Function
    Private Shared Function GetBodyString(ByVal aType As CmdType)
        Try

            If aType = CmdType.Cmd1002 Then
                Return "1002"
            Else
                Return String.Empty
            End If

        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
            Return String.Empty
        End Try
    End Function
    Private Shared Function Next_transaction_number() As UInt32
        Try

            If _Transaction_number = 999999999 Then
                Threading.Interlocked.Exchange(_Transaction_number, 1)
            Else
                Threading.Interlocked.Increment(_Transaction_number)
            End If

            Return _Transaction_number
        Catch ex As Exception
            Return "0"
        End Try
    End Function

End Class

Public Enum CmdType
    ConfirmConnect = 0
    Cmd1002 = 1
    Cmd0851 = 2
End Enum
Public Enum AckCode
    Ack = 0
    Nack = 1
End Enum
Public Structure NagraCmdStatus_Type
    Private _ErrCode As String
    Private _ErrName As String
    Private _Status As AckCode
    Public Property ErrorCode() As String
        Get
            Return _ErrCode
        End Get
        Set(ByVal value As String)
            _ErrCode = value
        End Set
    End Property
    Public Property ErrorName() As String
        Get
            Return _ErrName
        End Get
        Set(ByVal value As String)
            _ErrName = value
        End Set
    End Property
    Public Property Status() As AckCode
        Get
            Return _Status
        End Get
        Set(ByVal value As AckCode)
            _Status = value
        End Set
    End Property
End Structure

Public Structure NagraStatus_Type
    Private _SUCCESS As Byte
    Private _answer_code As Byte
    Private _CmdType As CmdType
    Private _ErrorCode As Int32
    Private _ErrorName As Int32
    Public Property SUCCESS() As Byte
        Get
            Return _SUCCESS
        End Get
        Set(ByVal value As Byte)
            _SUCCESS = value
        End Set
    End Property
    Public Property answer_code() As Byte
        Get
            Return _answer_code
        End Get
        Set(ByVal value As Byte)
            _answer_code = value
        End Set
    End Property
    Public Property CmdType() As CmdType
        Get
            Return _CmdType
        End Get
        Set(ByVal value As CmdType)
            _CmdType = value
        End Set
    End Property
End Structure