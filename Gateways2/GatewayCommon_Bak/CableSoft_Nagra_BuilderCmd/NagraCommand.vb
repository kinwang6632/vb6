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
Imports System.Threading
Public Class NagraCommand
    Private Const FixCmd1002Len As Int32 = 32
    Private Const FixCmd As Int32 = 32
    Private Const FIniFile As String = "TranNum.Ini"
    Private Const Fob_Name As String = "SMS_PVOD"
    Private _SourceId As String = "0000"
    Private _Dest_Id As String = "0000"
    Private _MOP_PPID As String = "00000"
    Public Property SourceId As String
        Get
            Return _SourceId
        End Get
        Set(ByVal value As String)
            _SourceId = Right("0000" & value, 4)
        End Set
    End Property
    Public Property Dest_Id As String
        Get
            Return _Dest_Id
        End Get
        Set(ByVal value As String)
            _Dest_Id = Right("0000" & value, 4)
        End Set
    End Property
    Public Property MOP_PPID As String
        Get
            Return _MOP_PPID
        End Get
        Set(ByVal value As String)
            _MOP_PPID = Right("00000" & value, 5)
        End Set
    End Property
    Public Shared Property ErrCodeLst As Dictionary(Of String, String)
        Get
            Return _ErrCodeLst
        End Get
        Set(ByVal value As Dictionary(Of String, String))
            _ErrCodeLst = value
        End Set
    End Property

    Public Shared Property _Transaction_number As Int32
    Public Shared Property IniPath As String = String.Empty
    Private Shared _ErrCodeLst As Dictionary(Of String, String)
    Public Sub New()

    End Sub
    Public Sub New(ByVal aSourceId As String, ByVal aDest_Id As String, ByVal aMOP_PPID As String)
        _SourceId = aSourceId
        _Dest_Id = aDest_Id
        _MOP_PPID = aMOP_PPID
    End Sub
    Public Function BuilderCmd(ByVal aCmdType As CmdType, _
                                      ByVal aValueLst As Dictionary(Of String, String)) As List(Of Byte())
        'Dim aRet As String = String.Empty
        Try
            Select Case aCmdType
                Case CmdType.ConfirmConnect
                    Return ConfirmConnect()
                Case CmdType.Cmd1002
                    Return GetCMD1002()
                Case CmdType.Cmd0851
                    Return GetCMD0851(aValueLst)
                Case Else
                    Return Nothing
            End Select
            'Return aRet
        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
            Return Nothing
        End Try
    End Function
    Public Overloads Function NagraStatus(ByVal aByte() As Byte, ByVal aType As CmdType) As Object
        Try

            Select Case aType
                Case CmdType.ConfirmConnect
                    Return GetConfirmtStatus(aByte)
                Case CmdType.Cmd1002
                    Return GetCmd1002Status(aByte)
                Case CmdType.Cmd0851
                    Dim aStr As String = Encoding.ASCII.GetString(aByte, 2, aByte.Length - 2).TrimEnd(Chr(0))
                    Return GetCMDStatus(aStr, aStr.Length, aType)
                Case Else
                    Return Nothing
            End Select
        Catch ex As Exception

            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
            Return Nothing
        End Try
    End Function
    Public Overloads Function NagraStatus(ByVal aString As String, ByVal aLen As Int32, ByVal aType As CmdType) As Object
        Try
            Select Case aType
                Case CmdType.ConfirmConnect
                    Return GetConfirmtStatus(Encoding.ASCII.GetBytes(aString))
                Case CmdType.Cmd1002
                    Return GetCmd1002Status(Encoding.ASCII.GetBytes(aString))
                Case CmdType.OtherCMD
                    Return Nothing
                Case Else
                    Return GetCMDStatus(aString, aLen, aType)

            End Select

        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
            Return Nothing
        End Try
    End Function
    Public Overloads Function GetCMDStatus(ByVal aString As String, ByVal aLen As Int32, ByVal aType As CmdType) As NagraCmdStatus_Type
        Dim obj As New ThreadLocal(Of NagraCmdStatus_Type)
        obj.Value = New NagraCmdStatus_Type()


        obj.Value.Entities = String.Empty
        obj.Value.ErrorCode = "-9999"
        obj.Value.ErrorName = "CableSoft_Initial_Type"
        obj.Value.Status = AckCode.Nack
        Try

            If aString.Length <> aLen Then
                obj.Value.ErrorName = "CableSoft_Length_Error"
            Else
                If aType = CmdType.Cmd0851 Then
                    If aString.Substring(FixCmd, 4) = "2005" Then
                        obj.Value.ErrorCode = String.Empty
                        obj.Value.ErrorName = String.Empty
                        obj.Value.Status = AckCode.Ack
                        Dim aObjNagraResult As New ThreadLocal(Of Dictionary(Of String, String))
                        aObjNagraResult.Value = New Dictionary(Of String, String)
                        aObjNagraResult.Value = GetNagraResult(aString, aType)
                        If aObjNagraResult.Value IsNot Nothing AndAlso aObjNagraResult.Value.Count > 0 Then
                            obj.Value.Entities = aObjNagraResult.Value.Item("ENTITIES")
                            obj.Value.Entity_data = aObjNagraResult.Value.Item("ENTITY_DATA")
                        End If

                    Else
                        If aString.Substring(FixCmd, 4) <> "2006" Then
                            obj.Value.Status = AckCode.Other
                            obj.Value.ErrorCode = aString.Substring(FixCmd + 4 + 9 + 1, 4)
                            obj.Value.ErrorName = "CableSoft_Not_Know_Error"
                        Else
                            obj.Value.ErrorCode = aString.Substring(FixCmd + 4 + 9 + 1, 4)
                            obj.Value.ErrorName = aString.Substring(FixCmd + 4 + 9 + 1 + 4, 4)
                        End If
                        If _ErrCodeLst IsNot Nothing AndAlso _ErrCodeLst.Count > 0 Then
                            If _ErrCodeLst.ContainsKey(obj.Value.ErrorName) Then
                                obj.Value.ErrorName = _ErrCodeLst.Item(obj.Value.ErrorName)
                            End If
                        End If

                    End If
                End If

            End If
            Return obj.Value
        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
            Return obj.Value
        End Try
    End Function
    Private Function GetNagraResult(ByVal aString As String, _
                                           ByVal aType As CmdType) As Dictionary(Of String, String)
        Try
            Dim obj As New ThreadLocal(Of Dictionary(Of String, String))
            obj.Value = New Dictionary(Of String, String)
            obj.Value.Add("ENTITIES", String.Empty)
            obj.Value.Add("ENTITY_DATA", String.Empty)
            Select Case aType
                Case CmdType.Cmd0851
                    If aString.Length > FixCmd + 4 + 9 + 3 + 4 Then
                        obj.Value.Item("ENTITIES") = aString.Substring(FixCmd + 4 + 9, 3)
                        Dim aEntity_Len As Int32 = Convert.ToInt32(aString.Substring(FixCmd + 4 + 9 + 3, 4)) * 2
                        If (aEntity_Len > 0) Then
                            If aString.Length >= (FixCmd + 4 + 9 + 3 + 4 + aEntity_Len) Then
                                'Dim aEntity_Data As String = aString.Substring(FixCmd + 4 + 9 + 3 + 4, aEntity_Len)
                                obj.Value.Item("ENTITY_DATA") = aString.Substring(FixCmd + 4 + 9 + 3 + 4, aEntity_Len)
                                'HexToAscii(aEntity_Data)
                            End If
                        End If
                    End If
            End Select
            Return obj.Value
        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
            Return Nothing
        End Try
    End Function
    '取得連接狀態
    Private Function GetConfirmtStatus(ByVal aByte() As Byte) As NagraStatus_Type
        Try
            Dim obj As New NagraStatus_Type
            obj.CmdType = CmdType.ConfirmConnect
            obj.SUCCESS = aByte(2)
            obj.answer_code = aByte(5)
            Return obj
        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
            Return Nothing
        End Try
    End Function
    'Private Shared Function GetCMDStatus(ByVal aByte() As Byte, ByVal aType As CmdType) As NagraCmdStatus_Type
    '    Dim objRet As New NagraCmdStatus_Type
    '    objRet.Status = AckCode.Nack
    '    objRet.ErrorCode = -9998
    '    objRet.ErrorName = "初始化狀態"
    '    Try

    '    Catch ex As Exception
    '        WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
    '        Return objRet
    '    End Try

    'End Function
    '取得回傳資料總長度
    Public Function GetRetunLength(ByVal aByte() As Byte) As Int32
        Try
            If aByte.Length < 2 Then
                Return 0
            Else

                Dim bty(1) As Byte
                bty(0) = aByte(0)
                bty(1) = aByte(1)
                Array.Reverse(bty)
                Return BitConverter.ToInt16(bty, 0)
            End If
        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
            Return 0
        End Try
    End Function
    '取出接收的字串,aContainLen=True 陣列包含長度
    Public Function GetReturnString(ByVal aBty() As Byte, ByVal aContainLen As Boolean) As String
        Try
            Dim aRet As New ThreadLocal(Of String)
            If Not aContainLen Then
                aRet.Value = Encoding.ASCII.GetString(aBty).TrimEnd(Chr(0))
                'Return Encoding.ASCII.GetString(aBty).TrimEnd(Chr(0))
            Else
                aRet.Value = Encoding.ASCII.GetString(aBty, 2, aBty.Length - 2).TrimEnd(Chr(0))
                'Return Encoding.ASCII.GetString(aBty, 2, aBty.Length - 2).TrimEnd(Chr(0))

            End If
            Return aRet.Value
        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
            Return String.Empty
        End Try
    End Function
    Private Function GetCmd1002Status(ByVal aByte() As Byte) As NagraCmdStatus_Type
        Dim objRet As New NagraCmdStatus_Type
        objRet.Status = AckCode.Nack
        objRet.ErrorCode = -9998
        objRet.ErrorName = "初始化狀態"

        Try

            'Y 00000000205123420005678920101008100120000000110003000003220000000105200012345678920101108
            '0E000000003051234200056789201010081000200000001000000000000000000000000
            '前面2個為總長度
            'Dim aLen As Int32 = Convert.ToInt32(aByte(0).ToString & aByte(1).ToString)
            Dim aLenbty(1) As Byte
            aLenbty(0) = aByte(1)
            aLenbty(1) = aByte(0)
            Dim aLen As Short = BitConverter.ToInt16(aLenbty, 0)
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

                    If _ErrCodeLst.ContainsKey(objRet.ErrorName) Then
                        objRet.ErrorName = _ErrCodeLst.Item(objRet.ErrorName)
                    End If
                End If
            End If
            Return objRet
        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
            Return objRet
        End Try
    End Function
    Private Function ConfirmConnect() As List(Of Byte())

        Try
            Dim ob_name_len As Byte = Encoding.ASCII.GetByteCount(Fob_Name)
            Dim op_mode As Byte = 0
            Dim Len1 As Byte = 0
            Dim Len2 As Byte = SizeOf(op_mode) + SizeOf(ob_name_len) + Convert.ToInt32(ob_name_len)

            Dim aRet As String = Chr(Len1) & Chr(Len2) & Chr(op_mode) & Chr(ob_name_len) & Fob_Name
            Dim retLst As New List(Of Byte())
            retLst.Add(Encoding.ASCII.GetBytes(aRet))
            Return retLst
        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
            Return Nothing
        End Try
    End Function
    Private Function GetCMD0851(ByVal aValueLst As Dictionary(Of String, String)) As List(Of Byte())
        Try



            Dim aRet As New ThreadLocal(Of String)
            Dim RetLst As New ThreadLocal(Of List(Of Byte()))
            Dim aBodyValue As New ThreadLocal(Of List(Of String))
            RetLst.Value = New List(Of Byte())
            aBodyValue.Value = New List(Of String)
            aBodyValue.Value = GetBodyCMD0851(aValueLst)


            For i As Int32 = 0 To aBodyValue.Value.Count - 1
                Dim aHeadValue As New ThreadLocal(Of String)
                aHeadValue.Value = GetHeadString(CmdType.Cmd0851) & GetSectionString(CmdType.Cmd0851, aValueLst)
                Dim aValue As New ThreadLocal(Of String)
                aValue.Value = aHeadValue.Value & aBodyValue.Value.Item(i)
                Dim btyLst As New ThreadLocal(Of List(Of Byte))
                btyLst.Value = New List(Of Byte)
                btyLst.Value.Clear()
                btyLst.Value.AddRange(DataLength(aValue.Value.Length))
                btyLst.Value.AddRange(Encoding.ASCII.GetBytes(aValue.Value))

                RetLst.Value.Add(btyLst.Value.ToArray)
            Next

            Return RetLst.Value


        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
            Return Nothing
        End Try
    End Function
    Private Function GetCMD1002() As List(Of Byte())
        Try

            Dim aRet As New ThreadLocal(Of String)
            aRet.Value = GetHeadString(CmdType.Cmd1002) & GetSectionString(CmdType.Cmd1002) & GetBodyString(CmdType.Cmd1002)
            'Dim retBtyLst As New List(Of Byte())
            Dim retBtyLst As New ThreadLocal(Of List(Of Byte()))

            'Dim strBty() As Byte = Encoding.ASCII.GetBytes(aRet)
            Dim strBty As New ThreadLocal(Of Byte())
            strBty.Value = Encoding.ASCII.GetBytes(aRet.Value)

            Dim btyLst As New ThreadLocal(Of List(Of Byte))
            btyLst.Value = New List(Of Byte)
            btyLst.Value.Clear()
            btyLst.Value.AddRange(DataLength(aRet.Value.Length))
            btyLst.Value.AddRange(strBty.Value)

            retBtyLst.Value = New List(Of Byte())
            retBtyLst.Value.Add(btyLst.Value.ToArray)

            Return retBtyLst.Value

        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
            Return Nothing
        End Try
    End Function
    Private Shared Function DataLength(ByVal aLen As Short) As Byte()
        Try
            Dim aLenBty As New ThreadLocal(Of Byte())
            aLenBty.Value = BitConverter.GetBytes(aLen)
            Array.Reverse(aLenBty.Value)
            Return aLenBty.Value
            'Dim aLenBty() As Byte = BitConverter.GetBytes(aLen)
            'Array.Reverse(aLenBty)
            'Return aLenBty

        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
            Return Nothing
        End Try
    End Function
    Public Shared Function GetSndTransNum(ByVal aData As String) As String
        Return aData.Substring(0, 9)
    End Function
    Public Shared Function GetRcvTransNum(ByVal aData As String, ByVal aCmdType As CmdType) As String
        Select Case aCmdType
            Case CmdType.Cmd0851
                Return String.Empty
            Case CmdType.Cmd1002
                If aData.Substring(32, 4) = "1000" Then
                    Return aData.Substring(37, 9)
                Else
                    Return aData.Substring(57, 9)
                End If


            Case Else
                Return String.Empty
        End Select
    End Function
    Private Function GetHeadString(ByVal aType As CmdType) As String
        Try

            Dim transaction_number As New ThreadLocal(Of String)
            Dim command_type As New ThreadLocal(Of String)
            command_type.Value = "05"
            If aType = CmdType.Cmd0851 Then
                command_type.Value = "08"
            End If

            'Dim source_id As String = _SourceId
            'Dim dest_id As String = _Dest_Id
            'Dim MOP_PPID As String = _MOP_PPID

            Dim creation_date As New ThreadLocal(Of String)
            creation_date.Value = Format(Date.UtcNow, "yyyyMMdd")
            transaction_number.Value = Right("000000000" & Next_transaction_number.ToString, 9)

            Return transaction_number.Value & command_type.Value & _SourceId & _
                _Dest_Id & _MOP_PPID & creation_date.Value

        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
            Return String.Empty
        End Try
    End Function
    Private Shared Function GetSectionString(ByVal aType As CmdType, Optional ByVal aValueLst As Dictionary(Of String, String) = Nothing)
        Try
            If aType = CmdType.Cmd1002 Then
                Return String.Empty
            Else
                Select Case aType
                    Case CmdType.Cmd0851
                        Return "E" & "U" & "S" & Right("0000000000" & aValueLst.Item("STBSNO"), 10) & "C"
                    Case Else
                        Return String.Empty
                End Select
            End If

        Catch ex As Exception
            Return String.Empty
        End Try



    End Function
    Private Shared Function StringToUTCDate(ByVal aSourceDate As String) As String
        Try
            If aSourceDate.Length <> 14 Then
                Return String.Empty
            End If
            Dim aStrDate As New ThreadLocal(Of String)
            aStrDate.Value = Left(aSourceDate, 4) & "/" & _
                aSourceDate.Substring(4, 2) & "/" & aSourceDate.Substring(6, 2) & " " & _
                aSourceDate.Substring(8, 2) & ":" & aSourceDate.Substring(10, 2) & ":" & _
                Right(aSourceDate, 2)
            Return Format(TimeZoneInfo.ConvertTimeToUtc(DateTime.Parse(aStrDate.Value)), "yyyyMMddhhmmss")
        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
            Return String.Empty
        End Try
    End Function
    Private Shared Function GetBodyCMD0851(ByVal aValueLst As Dictionary(Of String, String)) As List(Of String)
        'Dim strbdu As New StringBuilder()
        Try

            Dim aValue As New ThreadLocal(Of String)

            Dim aryStr As New ThreadLocal(Of String())
            aryStr.Value = aValueLst.Item("NOTES").Split(","c)
            Dim strLst As New ThreadLocal(Of List(Of String))
            strLst.Value = New List(Of String)
            For i As Int32 = 0 To aryStr.Value.Length - 1
                Dim aryItem As New ThreadLocal(Of String())
                aryItem.Value = aryStr.Value(i).Split("~"c)
                Dim metadata_length As New ThreadLocal(Of String)
                Dim metadata As New ThreadLocal(Of String)
                aValue.Value = String.Empty
                metadata.Value = aryItem.Value(0) & ";" & StringToUTCDate(aryItem.Value(2))
                metadata_length.Value = Right("000" & metadata.Value.Length.ToString, 3)

                'command_id (r_num->4)
                'VOD_ENT_ID (num->10)   流水號
                'content_id (num->10)   Asset

                aValue.Value = "0851" & _
                        Right("0000000000" & aryItem.Value(0), 10) & _
                        Right("0000000000" & aryItem.Value(1), 10) & _
                        aValueLst.Item("ENCRYPTED_ASSET_FLAG") & _
                        Left(StringToUTCDate(aryItem.Value(3)), 8) & _
                        Right(StringToUTCDate(aryItem.Value(3)), 6) & _
                        Right("00000000" & aryItem.Value(4), 8) & _
                        aValueLst.Item("VALIDITY_FLAG") & _
                        Right("000" & aValueLst.Item("NB_OF_R_VOD_ENT"), 3) & _
                        metadata_length.Value & _
                        ASCiiToHex(metadata.Value) & _
                        aValueLst.Item("CHIPSET_TYPE_FLAG") & _
                        aValueLst.Item("ADDITIONAL_ADDRESS_FLAG") & _
                        aValueLst.Item("DMM_CREATION_FATE_FLAG")

                If Convert.ToInt32(aValueLst.Item("DMM_CREATION_FATE_FLAG")) = 1 Then
                    aValue.Value = aValue.Value & Format(TimeZoneInfo.ConvertTimeToUtc(Date.Parse(aValueLst.Item("DMM_CREATION_DATE"))), "yyyyMMdd")
                End If
                strLst.Value.Add(aValue.Value)


            Next
            Return strLst.Value
            'Return aRet
        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
            Return Nothing
        End Try
    End Function
    Private Shared Function HexToAscii(ByVal aString As String) As String
        Try
            Dim aRet As New ThreadLocal(Of String)
            Dim aValue As New ThreadLocal(Of String)
            For i As Int32 = 0 To aString.Length - 1 Step 2
                aValue.Value = String.Empty
                aValue.Value = aString.Substring(i, 2)
                aValue.Value = Convert.ToInt32("0X" & aValue.Value, 16).ToString
                If String.IsNullOrEmpty(aRet.Value) Then
                    aRet.Value = Convert.ToString(Convert.ToInt32(aValue.Value))
                Else
                    aRet.Value = aRet.Value & Convert.ToString(Convert.ToInt32(aValue.Value))
                End If
            Next
            Return aRet.Value
        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
            Return String.Empty
        End Try
    End Function
    Private Shared Function ASCiiToHex(ByVal aStr As String) As String

        Try
            Dim aRet As New ThreadLocal(Of String)
            For i As Int32 = 0 To aStr.Length - 1
                Dim aValue As New ThreadLocal(Of String)
                aValue.Value = Right("00" & Hex(Asc(aStr.Substring(i))).ToString, 2)
                If String.IsNullOrEmpty(aRet.Value) Then
                    aRet.Value = aValue.Value
                Else
                    aRet.Value = aRet.Value & aValue.Value
                End If

            Next
            Return aRet.Value
        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
            Return String.Empty
        End Try
    End Function
    Private Shared Function GetBodyString(ByVal aType As CmdType)
        Try

            If aType = CmdType.Cmd1002 Then
                Return "1002"
            Else
                Select Case aType
                    Case CmdType.Cmd0851
                        Return String.Empty
                    Case Else
                        Return String.Empty
                End Select
            End If

        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
            Return String.Empty
        End Try
    End Function
    Private Shared Function BuilderCMD851() As String
        Try
            Return String.Empty

        Catch ex As Exception
            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
            Return String.Empty
        End Try
    End Function
    Private Function Next_transaction_number() As UInt32
        Try
            _Transaction_number = Convert.ToInt32(CableSoft.Gateway.Common.csCommon.ReadINI(_IniPath & FIniFile, "TranNum", "Num", "0"))
            If _Transaction_number > 999999999 Then
                Threading.Interlocked.Exchange(_Transaction_number, 1)
            Else
                Threading.Interlocked.Increment(_Transaction_number)
            End If


            CableSoft.Gateway.Common.csCommon.WriteINI(_IniPath & FIniFile, "TranNum", "Num", _Transaction_number)
            Return _Transaction_number

        Catch ex As Exception
            Return "0"
        End Try
    End Function

End Class















'Public Class NagraCommand
'    Private Const FixCmd1002Len As Int32 = 32
'    Private Const FixCmd As Int32 = 32
'    Private Const Fob_Name As String = "SMS_PVOD"
'    Private Shared _SourceId As String = "0000"
'    Private Shared _Dest_Id As String = "0000"
'    Private Shared _MOP_PPID As String = "00000"
'    Private Shared _Transaction_number As Int32
'    Private Shared _IniPath As String = String.Empty
'    Private Shared _ErrCodeLst As Dictionary(Of String, String)
'    Private Shared lckObject As New Object

'    Private Const FIniFile As String = "TranNum.Ini"
'    Public Shared Property SourceId() As String
'        Get
'            Return _SourceId
'        End Get
'        Set(ByVal value As String)
'            _SourceId = value
'            _SourceId = "0000" & _SourceId
'            _SourceId = Right(_SourceId, 4)
'        End Set
'    End Property
'    Public Shared Property Dest_Id() As String
'        Get
'            Return _Dest_Id
'        End Get
'        Set(ByVal value As String)
'            _Dest_Id = value
'            _Dest_Id = "00000" & _Dest_Id
'            _Dest_Id = Right(_Dest_Id, 4)
'        End Set
'    End Property
'    Public Shared Property MOP_PPID() As String
'        Get
'            Return _MOP_PPID
'        End Get
'        Set(ByVal value As String)
'            _MOP_PPID = value
'            _MOP_PPID = "00000" & _MOP_PPID
'            _MOP_PPID = Right(_MOP_PPID, 5)
'        End Set
'    End Property
'    Public Shared Property Transaction_number() As Int32
'        Get
'            Return _Transaction_number
'        End Get
'        Set(ByVal value As Int32)
'            _Transaction_number = value
'        End Set
'    End Property
'    Public Shared Property IniPath() As String
'        Get
'            Return _IniPath
'        End Get
'        Set(ByVal value As String)
'            _IniPath = value
'        End Set
'    End Property
'    Public Shared Property ErrCodeLst() As Dictionary(Of String, String)
'        Get
'            Return _ErrCodeLst
'        End Get
'        Set(ByVal value As Dictionary(Of String, String))
'            _ErrCodeLst = value
'        End Set
'    End Property

'    Public Shared Function BuilderCmd(ByVal aCmdType As CmdType, _
'                                      ByVal aValueLst As Dictionary(Of String, String)) As List(Of Byte())
'        'Dim aRet As String = String.Empty
'        Try
'            Select Case aCmdType
'                Case CmdType.ConfirmConnect
'                    Return ConfirmConnect()
'                Case CmdType.Cmd1002
'                    Return GetCMD1002()
'                Case CmdType.Cmd0851
'                    Return GetCMD0851(aValueLst)
'                Case Else
'                    Return Nothing
'            End Select
'            'Return aRet
'        Catch ex As Exception
'            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
'            Return Nothing
'        End Try
'    End Function
'    Public Overloads Shared Function NagraStatus(ByVal aString As String, ByVal aLen As Int32, ByVal aType As CmdType) As Object
'        Try
'            Select Case aType
'                Case CmdType.ConfirmConnect
'                    Return GetConfirmtStatus(Encoding.ASCII.GetBytes(aString))
'                Case CmdType.Cmd1002
'                    Return GetCmd1002Status(Encoding.ASCII.GetBytes(aString))
'                Case CmdType.OtherCMD
'                    Return Nothing
'                Case Else
'                    Return GetCMDStatus(aString, aLen, aType)

'            End Select

'        Catch ex As Exception
'            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
'            Return Nothing
'        End Try
'    End Function
'    Public Overloads Shared Function NagraStatus(ByVal aByte() As Byte, ByVal aType As CmdType) As Object
'        Try

'            Select Case aType
'                Case CmdType.ConfirmConnect
'                    Return GetConfirmtStatus(aByte)
'                Case CmdType.Cmd1002
'                    Return GetCmd1002Status(aByte)
'                Case CmdType.Cmd0851
'                    Dim aStr As String = Encoding.ASCII.GetString(aByte, 2, aByte.Length - 2).TrimEnd(Chr(0))
'                    Return GetCMDStatus(aStr, aStr.Length, aType)
'                Case Else
'                    Return Nothing
'            End Select
'        Catch ex As Exception

'            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
'            Return Nothing
'        End Try
'    End Function
'    Public Overloads Shared Function GetCMDStatus(ByVal aString As String, ByVal aLen As Int32, ByVal aType As CmdType) As NagraCmdStatus_Type
'        Dim obj As New NagraCmdStatus_Type()


'        obj.Entities = String.Empty
'        obj.ErrorCode = "-9999"
'        obj.ErrorName = "CableSoft_Initial_Type"
'        obj.Status = AckCode.Nack
'        Try

'            If aString.Length <> aLen Then
'                obj.ErrorName = "CableSoft_Length_Error"
'            Else
'                If aType = CmdType.Cmd0851 Then
'                    If aString.Substring(FixCmd, 4) = "2005" Then
'                        obj.ErrorCode = String.Empty
'                        obj.ErrorName = String.Empty
'                        obj.Status = AckCode.Ack
'                        Dim aObjNagraResult As Dictionary(Of String, String) = GetNagraResult(aString, aType)
'                        If aObjNagraResult IsNot Nothing AndAlso aObjNagraResult.Count > 0 Then
'                            obj.Entities = aObjNagraResult.Item("ENTITIES")
'                            obj.Entity_data = aObjNagraResult.Item("ENTITY_DATA")
'                        End If

'                    Else
'                        If aString.Substring(FixCmd, 4) <> "2006" Then
'                            obj.Status = AckCode.Other
'                            obj.ErrorCode = aString.Substring(FixCmd + 4 + 9 + 1, 4)
'                            obj.ErrorName = "CableSoft_Not_Know_Error"
'                        Else
'                            obj.ErrorCode = aString.Substring(FixCmd + 4 + 9 + 1, 4)
'                            obj.ErrorName = aString.Substring(FixCmd + 4 + 9 + 1 + 4, 4)
'                        End If
'                        If _ErrCodeLst IsNot Nothing AndAlso _ErrCodeLst.Count > 0 Then
'                            If _ErrCodeLst.ContainsKey(obj.ErrorName) Then
'                                obj.ErrorName = _ErrCodeLst.Item(obj.ErrorName)
'                            End If
'                        End If

'                    End If
'                End If

'            End If




'            Return obj
'        Catch ex As Exception
'            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
'            Return obj
'        End Try
'    End Function
'    Private Shared Function GetNagraResult(ByVal aString As String, _
'                                           ByVal aType As CmdType) As Dictionary(Of String, String)
'        Try
'            Dim obj As New ThreadLocal(Of Dictionary(Of String, String))
'            obj.Value.Add("ENTITIES", String.Empty)
'            obj.Value.Add("ENTITY_DATA", String.Empty)
'            Select Case aType
'                Case CmdType.Cmd0851
'                    If aString.Length > FixCmd + 4 + 9 + 3 + 4 Then
'                        obj.Value.Item("ENTITIES") = aString.Substring(FixCmd + 4 + 9, 3)
'                        Dim aEntity_Len As Int32 = Convert.ToInt32(aString.Substring(FixCmd + 4 + 9 + 3, 4)) * 2
'                        If (aEntity_Len > 0) Then
'                            If aString.Length >= (FixCmd + 4 + 9 + 3 + 4 + aEntity_Len) Then
'                                'Dim aEntity_Data As String = aString.Substring(FixCmd + 4 + 9 + 3 + 4, aEntity_Len)
'                                obj.Value.Item("ENTITY_DATA") = aString.Substring(FixCmd + 4 + 9 + 3 + 4, aEntity_Len)
'                                'HexToAscii(aEntity_Data)
'                            End If
'                        End If
'                    End If
'            End Select
'            Return obj.Value
'        Catch ex As Exception
'            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
'            Return Nothing
'        End Try
'    End Function
'    '取得連接狀態
'    Private Shared Function GetConfirmtStatus(ByVal aByte() As Byte) As NagraStatus_Type
'        Try
'            Dim obj As New NagraStatus_Type
'            obj.CmdType = CmdType.ConfirmConnect
'            obj.SUCCESS = aByte(2)
'            obj.answer_code = aByte(5)
'            Return obj
'        Catch ex As Exception
'            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
'            Return Nothing
'        End Try
'    End Function
'    'Private Shared Function GetCMDStatus(ByVal aByte() As Byte, ByVal aType As CmdType) As NagraCmdStatus_Type
'    '    Dim objRet As New NagraCmdStatus_Type
'    '    objRet.Status = AckCode.Nack
'    '    objRet.ErrorCode = -9998
'    '    objRet.ErrorName = "初始化狀態"
'    '    Try

'    '    Catch ex As Exception
'    '        WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
'    '        Return objRet
'    '    End Try

'    'End Function
'    '取得回傳資料總長度
'    Public Shared Function GetRetunLength(ByVal aByte() As Byte) As Int32
'        Try
'            If aByte.Length < 2 Then
'                Return 0
'            Else

'                Dim bty(1) As Byte
'                bty(0) = aByte(0)
'                bty(1) = aByte(1)
'                Array.Reverse(bty)
'                Return BitConverter.ToInt16(bty, 0)
'            End If
'        Catch ex As Exception
'            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
'            Return 0
'        End Try
'    End Function
'    '取出接收的字串,aContainLen=True 陣列包含長度
'    Public Shared Function GetReturnString(ByVal aBty() As Byte, ByVal aContainLen As Boolean) As String
'        Try
'            Dim aRet As New ThreadLocal(Of String)
'            If Not aContainLen Then
'                aRet.Value = Encoding.ASCII.GetString(aBty).TrimEnd(Chr(0))
'                'Return Encoding.ASCII.GetString(aBty).TrimEnd(Chr(0))
'            Else
'                aRet.Value = Encoding.ASCII.GetString(aBty, 2, aBty.Length - 2).TrimEnd(Chr(0))
'                'Return Encoding.ASCII.GetString(aBty, 2, aBty.Length - 2).TrimEnd(Chr(0))

'            End If
'            Return aRet.Value
'        Catch ex As Exception
'            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
'            Return String.Empty
'        End Try
'    End Function
'    Private Shared Function GetCmd1002Status(ByVal aByte() As Byte) As NagraCmdStatus_Type
'        Dim objRet As New NagraCmdStatus_Type
'        objRet.Status = AckCode.Nack
'        objRet.ErrorCode = -9998
'        objRet.ErrorName = "初始化狀態"

'        Try

'            'Y 00000000205123420005678920101008100120000000110003000003220000000105200012345678920101108
'            '0E000000003051234200056789201010081000200000001000000000000000000000000
'            '前面2個為總長度
'            'Dim aLen As Int32 = Convert.ToInt32(aByte(0).ToString & aByte(1).ToString)
'            Dim aLenbty(1) As Byte
'            aLenbty(0) = aByte(1)
'            aLenbty(1) = aByte(0)
'            Dim aLen As Short = BitConverter.ToInt16(aLenbty, 0)
'            Dim aString As String = Encoding.ASCII.GetString(aByte, 2, aByte.Length - 2).TrimEnd(Chr(0))
'            If aString.Length <> aLen Then
'                objRet.Status = AckCode.Nack

'                objRet.ErrorCode = -9999
'                objRet.ErrorName = "回傳長度比對錯誤"
'            Else
'                If aString.Substring(32, 4) = "1000" Then
'                    objRet.Status = AckCode.Ack
'                Else
'                    objRet.Status = AckCode.Nack
'                    objRet.ErrorCode = aString.Substring(46, 4)
'                    objRet.ErrorName = aString.Substring(50, 4)

'                    If _ErrCodeLst.ContainsKey(objRet.ErrorName) Then
'                        objRet.ErrorName = _ErrCodeLst.Item(objRet.ErrorName)
'                    End If
'                End If
'            End If
'            Return objRet
'        Catch ex As Exception
'            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
'            Return objRet
'        End Try
'    End Function
'    Private Shared Function ConfirmConnect() As List(Of Byte())

'        Try
'            Dim ob_name_len As Byte = Encoding.ASCII.GetByteCount(Fob_Name)
'            Dim op_mode As Byte = 0
'            Dim Len1 As Byte = 0
'            Dim Len2 As Byte = SizeOf(op_mode) + SizeOf(ob_name_len) + Convert.ToInt32(ob_name_len)

'            Dim aRet As String = Chr(Len1) & Chr(Len2) & Chr(op_mode) & Chr(ob_name_len) & Fob_Name
'            Dim retLst As New List(Of Byte())
'            retLst.Add(Encoding.ASCII.GetBytes(aRet))
'            Return retLst
'        Catch ex As Exception
'            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
'            Return Nothing
'        End Try
'    End Function
'    Private Shared Function GetCMD0851(ByVal aValueLst As Dictionary(Of String, String)) As List(Of Byte())
'        Try


'            SyncLock lckObject
'                Dim aRet As String = String.Empty
'                Dim RetLst As New List(Of Byte())

'                Dim aBodyValue As List(Of String) = Nothing
'                aBodyValue = GetBodyCMD0851(aValueLst)


'                For i As Int32 = 0 To aBodyValue.Count - 1
'                    Dim aHeadValue As String = String.Empty
'                    aHeadValue = GetHeadString(CmdType.Cmd0851) & GetSectionString(CmdType.Cmd0851, aValueLst)
'                    Dim aValue As String = aHeadValue & aBodyValue.Item(i)

'                    Dim btyLst As New List(Of Byte)
'                    btyLst.Clear()
'                    btyLst.AddRange(DataLength(aValue.Length))
'                    btyLst.AddRange(Encoding.ASCII.GetBytes(aValue))

'                    RetLst.Add(btyLst.ToArray)
'                Next

'                Return RetLst
'            End SyncLock

'        Catch ex As Exception
'            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
'            Return Nothing
'        End Try
'    End Function
'    Private Shared Function GetCMD1002() As List(Of Byte())
'        Try
'            SyncLock lckObject
'                Dim aRet As String = GetHeadString(CmdType.Cmd1002) & GetSectionString(CmdType.Cmd1002) & GetBodyString(CmdType.Cmd1002)
'                Dim retBtyLst As New List(Of Byte())
'                Dim strBty() As Byte = Encoding.ASCII.GetBytes(aRet)

'                Dim btyLst As New List(Of Byte)
'                btyLst.Clear()
'                btyLst.AddRange(DataLength(aRet.Length))
'                btyLst.AddRange(strBty)


'                retBtyLst.Add(btyLst.ToArray)

'                Return retBtyLst
'            End SyncLock
'        Catch ex As Exception
'            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
'            Return Nothing
'        End Try
'    End Function
'    Private Shared Function DataLength(ByVal aLen As Short) As Byte()
'        Try
'            Dim aLenBty() As Byte = BitConverter.GetBytes(aLen)
'            Array.Reverse(aLenBty)
'            Return aLenBty

'        Catch ex As Exception
'            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
'            Return Nothing
'        End Try
'    End Function
'    Public Shared Function GetSndTransNum(ByVal aData As String) As String
'        Return aData.Substring(0, 9)
'    End Function
'    Public Shared Function GetRcvTransNum(ByVal aData As String, ByVal aCmdType As CmdType) As String
'        Select Case aCmdType
'            Case CmdType.Cmd0851
'                Return String.Empty
'            Case CmdType.Cmd1002
'                If aData.Substring(32, 4) = "1000" Then
'                    Return aData.Substring(37, 9)
'                Else
'                    Return aData.Substring(57, 9)
'                End If


'            Case Else
'                Return String.Empty
'        End Select
'    End Function
'    Private Shared Function GetHeadString(ByVal aType As CmdType) As String
'        Try

'            Dim transaction_number As String = String.Empty
'            Dim command_type As String = "05"
'            If aType = CmdType.Cmd0851 Then
'                command_type = "08"
'            End If

'            Dim source_id As String = _SourceId
'            Dim dest_id As String = _Dest_Id
'            Dim MOP_PPID As String = _MOP_PPID

'            Dim creation_date As String = Format(Date.UtcNow, "yyyyMMdd")
'            transaction_number = Right("000000000" & Next_transaction_number.ToString, 9)

'            Return transaction_number & command_type & source_id & _
'                dest_id & MOP_PPID & creation_date

'        Catch ex As Exception
'            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
'            Return String.Empty
'        End Try
'    End Function
'    Private Shared Function GetSectionString(ByVal aType As CmdType, Optional ByVal aValueLst As Dictionary(Of String, String) = Nothing)
'        Try
'            If aType = CmdType.Cmd1002 Then
'                Return String.Empty
'            Else
'                Select Case aType
'                    Case CmdType.Cmd0851
'                        Return "E" & "U" & "S" & Right("0000000000" & aValueLst.Item("STBSNO"), 10) & "C"
'                    Case Else
'                        Return String.Empty
'                End Select
'            End If

'        Catch ex As Exception
'            Return String.Empty
'        End Try



'    End Function
'    Private Shared Function StringToUTCDate(ByVal aSourceDate As String) As String
'        Try
'            If aSourceDate.Length <> 14 Then
'                Return String.Empty
'            End If
'            Dim aStrDate As String = Left(aSourceDate, 4) & "/" & _
'                aSourceDate.Substring(4, 2) & "/" & aSourceDate.Substring(6, 2) & " " & _
'                aSourceDate.Substring(8, 2) & ":" & aSourceDate.Substring(10, 2) & ":" & _
'                Right(aSourceDate, 2)
'            Return Format(TimeZoneInfo.ConvertTimeToUtc(DateTime.Parse(aStrDate)), "yyyyMMddhhmmss")
'        Catch ex As Exception
'            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
'            Return String.Empty
'        End Try
'    End Function
'    Private Shared Function GetBodyCMD0851(ByVal aValueLst As Dictionary(Of String, String)) As List(Of String)
'        'Dim strbdu As New StringBuilder()
'        Try

'            Dim aValue As String = String.Empty

'            Dim aryStr() As String = aValueLst.Item("NOTES").Split(","c)
'            Dim strLst As New List(Of String)
'            For i As Int32 = 0 To aryStr.Length - 1
'                Dim aryItem As String() = aryStr(i).Split("~"c)
'                Dim metadata_length As String = String.Empty
'                Dim metadata As String = String.Empty
'                aValue = String.Empty
'                metadata = aryItem(0) & ";" & StringToUTCDate(aryItem(2))
'                metadata_length = Right("000" & metadata.Length.ToString, 3)

'                'command_id (r_num->4)
'                'VOD_ENT_ID (num->10)   流水號
'                'content_id (num->10)   Asset

'                aValue = "0851" & _
'                        Right("0000000000" & aryItem(0), 10) & _
'                        Right("0000000000" & aryItem(1), 10) & _
'                        aValueLst.Item("ENCRYPTED_ASSET_FLAG") & _
'                        Left(StringToUTCDate(aryItem(3)), 8) & _
'                        Right(StringToUTCDate(aryItem(3)), 6) & _
'                        Right("00000000" & aryItem(4), 8) & _
'                        aValueLst.Item("VALIDITY_FLAG") & _
'                        Right("000" & aValueLst.Item("NB_OF_R_VOD_ENT"), 3) & _
'                        metadata_length & _
'                        ASCiiToHex(metadata) & _
'                        aValueLst.Item("CHIPSET_TYPE_FLAG") & _
'                        aValueLst.Item("ADDITIONAL_ADDRESS_FLAG") & _
'                        aValueLst.Item("DMM_CREATION_FATE_FLAG")

'                If Convert.ToInt32(aValueLst.Item("DMM_CREATION_FATE_FLAG")) = 1 Then
'                    aValue = aValue & Format(TimeZoneInfo.ConvertTimeToUtc(Date.Parse(aValueLst.Item("DMM_CREATION_DATE"))), "yyyyMMdd")
'                End If
'                strLst.Add(aValue)


'            Next
'            Return strLst
'            'Return aRet
'        Catch ex As Exception
'            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
'            Return Nothing
'        End Try
'    End Function
'    Private Shared Function HexToAscii(ByVal aString As String) As String
'        Try
'            Dim aRet As String = String.Empty
'            Dim aValue As String = String.Empty
'            For i As Int32 = 0 To aString.Length - 1 Step 2
'                aValue = String.Empty
'                aValue = aString.Substring(i, 2)
'                aValue = Convert.ToInt32("0X" & aValue, 16).ToString
'                If String.IsNullOrEmpty(aRet) Then
'                    aRet = Convert.ToString(Convert.ToInt32(aValue))
'                Else
'                    aRet = aRet & Convert.ToString(Convert.ToInt32(aValue))
'                End If
'            Next
'            Return aRet
'        Catch ex As Exception
'            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
'            Return String.Empty
'        End Try
'    End Function
'    Private Shared Function ASCiiToHex(ByVal aStr As String) As String

'        Try
'            Dim aRet As String = String.Empty
'            For i As Int32 = 0 To aStr.Length - 1
'                Dim aValue As String = Right("00" & Hex(Asc(aStr.Substring(i))).ToString, 2)
'                If String.IsNullOrEmpty(aRet) Then
'                    aRet = aValue
'                Else
'                    aRet = aRet & aValue
'                End If

'            Next
'            Return aRet
'        Catch ex As Exception
'            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
'            Return String.Empty
'        End Try
'    End Function
'    Private Shared Function GetBodyString(ByVal aType As CmdType)
'        Try

'            If aType = CmdType.Cmd1002 Then
'                Return "1002"
'            Else
'                Select Case aType
'                    Case CmdType.Cmd0851
'                        Return String.Empty
'                    Case Else
'                        Return String.Empty
'                End Select
'            End If

'        Catch ex As Exception
'            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
'            Return String.Empty
'        End Try
'    End Function
'    Private Shared Function BuilderCMD851() As String
'        Try
'            Return String.Empty

'        Catch ex As Exception
'            WriteErrTxtLog.WriteTxtError(ex, Nothing, Nothing)
'            Return String.Empty
'        End Try
'    End Function
'    Private Shared Function Next_transaction_number() As UInt32
'        Try
'            _Transaction_number = Convert.ToInt32(CableSoft.Gateway.Common.csCommon.ReadINI(_IniPath & FIniFile, "TranNum", "Num", "0"))
'            If _Transaction_number > 999999999 Then
'                Threading.Interlocked.Exchange(_Transaction_number, 1)
'            Else
'                Threading.Interlocked.Increment(_Transaction_number)
'            End If
'            CableSoft.Gateway.Common.csCommon.WriteINI(_IniPath & FIniFile, "TranNum", "Num", _Transaction_number)
'            Return _Transaction_number

'        Catch ex As Exception
'            Return "0"
'        End Try
'    End Function

'End Class

Public Enum CmdType
    ConfirmConnect = 0
    Cmd1002 = 1
    Cmd0851 = 2
    OtherCMD = 99
End Enum
Public Enum AckCode
    Ack = 0
    Nack = 1
    Other = 2
End Enum
Public Class NagraCmdStatus_Type

    Private _ErrCode As String

    Private _ErrName As String

    Private _Status As AckCode

    Private _Entities As String

    Private _Entity_Data As String
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
    Public Property Entities() As String
        Get
            Return _Entities
        End Get
        Set(ByVal value As String)
            _Entities = value
        End Set
    End Property
    Public Property Entity_data() As String
        Get
            Return _Entity_Data
        End Get
        Set(ByVal value As String)
            _Entity_Data = value
        End Set
    End Property
End Class

Public Class NagraStatus_Type
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
End Class
'Dim bty() As Byte = System.BitConverter.GetBytes(256)   '整數轉成Byte陣列(整數佔4個Byte)
'BitConverter.ToInt32(Btye(3), 0)   Byte陣列轉整數(整數佔4個byte,所以要傳入4個byte陣列)

'MessageBox.Show(Convert.ToInt32("0X30", 16))   '16進制30轉成10進制(48)
'MessageBox.Show(Convert.ToString(48, 16))      '10進制48轉成16進制(30)