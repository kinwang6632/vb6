Imports System.Threading
Imports System.Windows.Forms
Imports System.Text
Imports CableSoft.CardLess
Imports CableSoft.CardLess.XMLFileIO
Imports System.Runtime.InteropServices.Marshal
Imports System.Globalization

Public Class LowCmd
    Implements IDisposable
    Private SourceId As String
    Private DestId As String
    Private MoppId As String
    Private Shared Transaction_Number As Integer
    Private Const FixCmd1002Len As Byte = 32
    Private Const FixCmd As Byte = 32
    Private MaxTransaction_Number As UInteger = 999999999
    Private xmlFile As CableSoft.CardLess.XMLFileIO.XmlFileIO
    Private Const Fob_Name As String = "SMS_PVOD"
    Public Property UseUTC As Boolean = True
    Public Property EncodingType As String = "UTF-8"
    Public Property mtxWait As Mutex = Nothing
    Public Property ErrorDataTable As DataTable = Nothing
    Private ErrorList As Dictionary(Of String, String)
    Private Const Transaction_Number_ParentNote As String = "Nagra"
    Private Const Transaction_Number_ChildNote As String = "Transaction_Number"
    Public Enum AckCode
        Ack = 0
        Nack = 1
        Other = 2
    End Enum
    Public Enum CmdType
        ConfirmConnect = 0
        Cmd1002 = 1
        Cmd0851 = 2
        Cmd0126 = 3
        OtherCMD = 99
    End Enum
    Public Enum CableSoftCustomerError
        CableSoft_Initial = -9999
        CableSoft_Socket_Disconnect = -9998
        CableSoft_Socket_All_Disconnect = -9997
        CableSoft_Socket_Recv_Error = -9996
        CableSoft_Socket_Send_Error = -9995
        CableSoft_Socket_Length_Error = -9994
        CableSoft_Socket_Not_Return2006 = -9993
        CableSoft_Socket_RecvOver = -9992
        CableSoft_Socket_SndOver = -9991
        CableSoft_Recv_Socket_Disconnect = -9990
        CableSoft_0851String_Error = -5999
        CableSoft_CmdString_Error = -5998
        CableSoft_OtherError = -9000
        CableSoft_UnKnowCommand = -9001
    End Enum
    Public Sub New(ByVal SourceId As String,
                   ByVal DestId As String,
                   ByVal MoppId As String, ByVal tbError As DataTable)
        Try
            Me.SourceId = SourceId
            Me.DestId = DestId
            Me.MoppId = MoppId
            If tbError IsNot Nothing Then
                ErrorDataTable = tbError
                If ErrorList Is Nothing Then
                    ErrorList = New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
                End If
                For Each rw As DataRow In tbError.Rows
                    ErrorList.Add(rw.Item("ErrorCode"), rw.Item("ErrorName"))
                Next
            End If
            Dim maxTemp As Object = 999999999
            Try
                Transaction_Number = XmlFileIO.Read_Transaction_Number()
                'If xmlFile Is Nothing Then
                '    xmlFile = New CableSoft.CardLess.XMLFileIO.XmlFileIO()
                'End If
                'maxTemp = xmlFile.ReadXmlAttributeValue(Transaction_Number_ParentNote,
                '                                        Transaction_Number_ChildNote, "Max", False)
            Catch ex As Exception
                Transaction_Number = 0
            End Try

            If maxTemp IsNot Nothing Then
                MaxTransaction_Number = UInteger.Parse(maxTemp)
            End If
        Catch ex As Exception
            Throw
        End Try

    End Sub
    Public Sub New(ByVal SourceId As String, ByVal DestId As String, ByVal MoppId As String)
        Me.New(SourceId, DestId, MoppId, Nothing)
    End Sub
    Public Function GetLowCmd(ByVal CmdType As CmdType, ByVal rw As Dictionary(Of String, Object)) As Byte()
        Select Case CmdType
            Case LowCmd.CmdType.ConfirmConnect
                Return BuilderCheckState()
            Case LowCmd.CmdType.Cmd1002
                Return BuilderCMD1002()
            Case LowCmd.CmdType.Cmd0851
                Return BuilderCMD0851(rw)
            Case LowCmd.CmdType.Cmd0126
                Return BuilderCMD0126(rw)
            Case Else
                Return Nothing
        End Select
    End Function
    Private Function Next_transaction_number() As Integer
        Try
            If Transaction_Number > 999999999 Then
                Threading.Interlocked.Exchange(Transaction_Number, 1)
            Else
                Threading.Interlocked.Increment(Transaction_Number)
            End If
            Return Transaction_Number
        Catch ex As Exception
            Return 0
        End Try
    End Function
    Private Function GetHeadString(ByVal CommandType As CmdType) As String
        Dim transaction_number As New ThreadLocal(Of String)
        Dim command_type As New ThreadLocal(Of String)
        Dim creation_date As New ThreadLocal(Of String)
        Try
            Select Case CommandType
                Case CmdType.Cmd0851
                    command_type.Value = "08"
                Case CmdType.Cmd0126
                    command_type.Value = "02"
                Case Else
                    command_type.Value = "05"
            End Select
            If UseUTC Then
                creation_date.Value = Format(Date.UtcNow, "yyyyMMdd")
            Else
                creation_date.Value = Format(Date.Now, "yyyyMMdd")
            End If

            transaction_number.Value = Right("000000000" & Next_transaction_number.ToString, 9)

            Return transaction_number.Value & command_type.Value & Right("0000" & SourceId, 4) & _
              Right("0000" & DestId, 4) & Right("00000" & MoppId, 5) & creation_date.Value

        Catch ex As Exception
            CableSoft.NSTV.Log.NstvLog.WriteErrorLog(ex, Nothing)
            Return String.Empty
        Finally
            transaction_number.Dispose()
            command_type.Dispose()
            creation_date.Dispose()
        End Try
    End Function
    Private Function GetSectionString(ByVal CommandType As CmdType,
                                             ByVal FieldValue As Dictionary(Of String, Object))
        Try
            If CommandType = CmdType.Cmd1002 Then
                Return String.Empty
            Else
                Select Case CommandType
                    Case CmdType.Cmd0851
                        Return "E" & "U" & "S" & Right("0000000000" & FieldValue.Item("STBSNO"), 10) & "C"
                    Case CmdType.Cmd0126
                        Return "N" & Format(Date.Now.AddDays(-1).ToUniversalTime, "yyyyMMdd") & _
                            Format(Date.Now.AddDays(1).ToUniversalTime, "yyyyMMdd") & "G"
                    Case Else
                        Return String.Empty
                End Select
            End If
        Catch ex As Exception
            CableSoft.NSTV.Log.NstvLog.WriteErrorLog(ex, Nothing)
            Return String.Empty
        End Try
    End Function
    Private Function GetBodyString(ByVal aType As CmdType)
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
            CableSoft.NSTV.Log.NstvLog.WriteErrorLog(ex, Nothing)
            Return String.Empty
        End Try
    End Function
    Private Function DataLength(ByVal aLen As Short) As Byte()
        Dim aLenBty As New ThreadLocal(Of Byte())
        Try

            aLenBty.Value = BitConverter.GetBytes(aLen)
            Array.Reverse(aLenBty.Value)
            Return aLenBty.Value

        Catch ex As Exception
            CableSoft.NSTV.Log.NstvLog.WriteErrorLog(ex, Nothing)
            Return Nothing
        Finally
            aLenBty.Dispose()
        End Try
    End Function
    Private Function GetBodyCMD0126(ByVal rw As Dictionary(Of String, Object)) As String

        Using aValue As New ThreadLocal(Of String)
            Try
                aValue.Value = "0126" & Right("0000000000" & rw.Item("NOTES"), 10)
            Catch ex As Exception
                CableSoft.NSTV.Log.NstvLog.WriteErrorLog(ex, Nothing)
                Throw
            End Try
            Return aValue.Value
        End Using

    End Function
    Public Function BuilderCMD0126(ByVal rw As Dictionary(Of String, Object)) As Byte()
        'Dim aRet As New ThreadLocal(Of String)
        'Dim RetLst As New ThreadLocal(Of List(Of Byte()))
        'Dim aBodyValue As New ThreadLocal(Of String)
        'Dim aHeadValue As New ThreadLocal(Of String)
        Dim aValue As New ThreadLocal(Of String)
        Dim btyLst As New ThreadLocal(Of List(Of Byte))
        Try
          
            aValue.Value = GetHeadString(CmdType.Cmd0126) & GetSectionString(CmdType.Cmd0126, rw) & _
                                        GetBodyCMD0126(rw)
            btyLst.Value = New List(Of Byte)
            btyLst.Value.Clear()
            btyLst.Value.AddRange(DataLength(aValue.Value.Length))
            btyLst.Value.AddRange(Encoding.ASCII.GetBytes(aValue.Value))
            'RetLst.Value.Add(btyLst.Value.ToArray)
            Return btyLst.Value.ToArray
        Catch ex As Exception
            CableSoft.NSTV.Log.NstvLog.WriteErrorLog(ex, Nothing)
            Throw
        Finally
            'aRet.Dispose()
            'RetLst.Dispose()
            'aBodyValue.Dispose()
            'aHeadValue.Dispose()
            aValue.Dispose()
            btyLst.Dispose()
            'aValueLst.Clear()
        End Try
    End Function
    Public Function BuilderCMD0851(ByVal rw As Dictionary(Of String, Object)) As Byte()
        Dim aRet As New ThreadLocal(Of String)
        'Dim RetLst As New ThreadLocal(Of List(Of Byte()))
        Dim aBodyValue As New ThreadLocal(Of List(Of String))
        Dim aHeadValue As New ThreadLocal(Of String)
        Dim aValue As New ThreadLocal(Of String)
        Dim btyLst As New ThreadLocal(Of List(Of Byte))
        Try
            'RetLst.Value = New List(Of Byte())
            aBodyValue.Value = New List(Of String)
            aBodyValue.Value = GetBodyCMD0851(rw)


            For i As Int32 = 0 To aBodyValue.Value.Count - 1
                btyLst.Value = Nothing
                aHeadValue.Value = GetHeadString(CmdType.Cmd0851) & GetSectionString(CmdType.Cmd0851, rw)

                aValue.Value = aHeadValue.Value & aBodyValue.Value.Item(i)

                btyLst.Value = New List(Of Byte)
                btyLst.Value.Clear()
                btyLst.Value.AddRange(DataLength(aValue.Value.Length))
                btyLst.Value.AddRange(Encoding.ASCII.GetBytes(aValue.Value))

                'RetLst.Value.Add(btyLst.Value.ToArray)
            Next

            Return btyLst.Value.ToArray


        Catch ex As Exception
            Throw
        Finally
            aRet.Dispose()
            'RetLst.Dispose()
            aBodyValue.Dispose()
            aHeadValue.Dispose()
            aValue.Dispose()
            btyLst.Dispose()

        End Try
    End Function
    Private Function GetBodyCMD0851(ByVal aValueLst As Dictionary(Of String, Object)) As List(Of String)

        Dim aValue As New ThreadLocal(Of String)

        Dim aryStr As New ThreadLocal(Of String())
        Dim strLst As New ThreadLocal(Of List(Of String))
        Dim metadata_length As New ThreadLocal(Of String)
        Dim metadata As New ThreadLocal(Of String)
        Dim aryItem As New ThreadLocal(Of String())
        Try

            aryStr.Value = aValueLst.Item("NOTES").Split(","c)

            strLst.Value = New List(Of String)
            For i As Int32 = 0 To aryStr.Value.Length - 1

                aryItem.Value = aryStr.Value(i).Split("~"c)

                aValue.Value = String.Empty
                'Alticast那間白痴廠商要求Metadata的資料送到分就好 By Kin 2011/05/03
                metadata.Value = aryItem.Value(5) & ";" & Right("00000" & aryItem.Value(0), 5) & ";" & Left(StringToUTCDate(aryItem.Value(2)), 12)
                metadata_length.Value = Right("000" & metadata.Value.Length.ToString, 3)

                'command_id (r_num->4)
                'VOD_ENT_ID (num->10)   流水號
                'content_id (num->10)   Asset

                aValue.Value = "0851" & _
                        Right("0000000000" & aryItem.Value(1), 10) & _
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
                    aValue.Value = aValue.Value & Format(TimeZoneInfo.ConvertTimeToUtc( _
                                                         Date.Parse(aValueLst.Item("DMM_CREATION_DATE"), CultureInfo.CreateSpecificCulture("zh-TW"))), "yyyyMMdd")
                End If
                strLst.Value.Add(aValue.Value)


            Next
            Return strLst.Value
            'Return aRet
        Catch ex As Exception
            CableSoft.NSTV.Log.NstvLog.WriteErrorLog(ex, Nothing)
            Throw
        Finally
            aValue.Dispose()
            aryStr.Dispose()
            strLst.Dispose()
            metadata_length.Dispose()
            metadata.Dispose()
            aryItem.Dispose()
        End Try
    End Function
    Private Function ASCiiToHex(ByVal aStr As String) As String
        Dim aRet As New ThreadLocal(Of String)
        Dim aValue As New ThreadLocal(Of String)
        Try

            For i As Int32 = 0 To aStr.Length - 1

                aValue.Value = Right("00" & Hex(Asc(aStr.Substring(i))).ToString, 2)
                If String.IsNullOrEmpty(aRet.Value) Then
                    aRet.Value = aValue.Value
                Else
                    aRet.Value = aRet.Value & aValue.Value
                End If

            Next
            Return aRet.Value
        Catch ex As Exception
            CableSoft.NSTV.Log.NstvLog.WriteErrorLog(ex, Nothing)
            Throw
        Finally
            aRet.Dispose()
            aValue.Dispose()
        End Try
    End Function
    Private Function StringToUTCDate(ByVal SourceDate As String) As String
        Dim aStrDate As New ThreadLocal(Of String)
        Try
            SourceDate = SourceDate.Replace("/", "").Replace(":", "").Trim()
            If SourceDate.Length <> 14 Then
                Return String.Empty
            End If

            aStrDate.Value = Left(SourceDate, 4) & "/" & _
                SourceDate.Substring(4, 2) & "/" & SourceDate.Substring(6, 2) & " " & _
                SourceDate.Substring(8, 2) & ":" & SourceDate.Substring(10, 2) & ":" & _
                Right(SourceDate, 2)

            Return Format(TimeZoneInfo.ConvertTimeToUtc( _
                          DateTime.Parse(aStrDate.Value, CultureInfo.CreateSpecificCulture("zh-TW"))), "yyyyMMddHHmmss")
        Catch ex As Exception
            CableSoft.NSTV.Log.NstvLog.WriteErrorLog(ex, Nothing)
            Throw
        Finally
            aStrDate.Dispose()
        End Try
    End Function
    Public Function BuilderCMD1002() As Byte()

        Dim aRet As New ThreadLocal(Of String)

        Dim strBty As New ThreadLocal(Of Byte())
        Dim btyLst As New ThreadLocal(Of List(Of Byte))
        Try
            aRet.Value = GetHeadString(CmdType.Cmd1002) & _
                        GetSectionString(CmdType.Cmd1002, Nothing) & _
                        GetBodyString(CmdType.Cmd1002)

            strBty.Value = Encoding.ASCII.GetBytes(aRet.Value)

            btyLst.Value = New List(Of Byte)
            btyLst.Value.Clear()
            btyLst.Value.AddRange(DataLength(aRet.Value.Length))
            btyLst.Value.AddRange(strBty.Value)
            Return btyLst.Value.ToArray

        Catch ex As Exception
            Throw
        Finally

            aRet.Dispose()
            strBty.Dispose()
            btyLst.Dispose()
        End Try
    End Function
    Public Function BuilderCheckState() As Byte()
        Dim ob_name_len As New ThreadLocal(Of Byte)
        Dim op_mode As New ThreadLocal(Of Byte)
        Dim Len1 As New ThreadLocal(Of Byte)
        Dim Len2 As New ThreadLocal(Of Byte)
        Dim aRet As New ThreadLocal(Of String)
        Try
            ob_name_len.Value = Encoding.ASCII.GetByteCount(Fob_Name)
            op_mode.Value = 0
            Len1.Value = 0
            Len2.Value = SizeOf(op_mode.Value) + SizeOf(ob_name_len.Value) + Convert.ToInt32(ob_name_len.Value)
            aRet.Value = Chr(Len1.Value) & _
                            Chr(Len2.Value) & _
                            Chr(op_mode.Value) & _
                            Chr(ob_name_len.Value) & _
                            Fob_Name
            Return System.Text.Encoding.ASCII.GetBytes(aRet.Value)
        Catch ex As Exception
            Throw
        Finally
            ob_name_len.Dispose()
            op_mode.Dispose()
            Len1.Dispose()
            Len2.Dispose()
            aRet.Dispose()
        End Try

    End Function
    Public Function ChkeckState(ByVal RecvData As Byte()) As CheckStatusResult
        Dim Result As New ThreadLocal(Of CheckStatusResult)
        Result.Value = New CheckStatusResult
        Try
            Result.Value.Success = RecvData(2)
            Result.Value.Answer_Code = RecvData(5)
            Return Result.Value
        Catch ex As Exception
            Throw
        Finally
            Result.Dispose()
        End Try

    End Function
    Private Function GetNagraResultValue(ByVal RecvString As String, _
                                       ByVal CommandType As CmdType) As Dictionary(Of String, String)
        Dim Result As New ThreadLocal(Of Dictionary(Of String, String))
        Try

            Result.Value = New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

            Select Case CommandType
                Case CmdType.Cmd0851
                    Result.Value.Add("ENTITIES", String.Empty)
                    Result.Value.Add("ENTITY_DATA", String.Empty)
                    If RecvString.Length > FixCmd + 4 + 9 + 3 + 4 Then
                        Result.Value.Item("ENTITIES") = RecvString.Substring(FixCmd + 4 + 9, 3)
                        Dim aEntity_Len As Int32 = Convert.ToInt32(RecvString.Substring(FixCmd + 4 + 9 + 3, 4)) * 2
                        If (aEntity_Len > 0) Then
                            If RecvString.Length >= (FixCmd + 4 + 9 + 3 + 4 + aEntity_Len) Then

                                Result.Value.Item("ENTITY_DATA") = RecvString.Substring(FixCmd + 4 + 9 + 3 + 4, aEntity_Len)

                            End If
                        End If
                    End If
                Case CmdType.Cmd0126
                    Result.Value.Add("NUID", String.Empty)
                    Result.Value.Add("VUA", String.Empty)
                    If RecvString.Length >= FixCmd + 4 + 9 + 10 + 10 Then
                        Result.Value.Item("VUA") = RecvString.Substring(FixCmd + 4 + 9, 10)
                        Result.Value.Item("NUID") = RecvString.Substring(FixCmd + 4 + 9 + 10, 10)
                    End If
            End Select
            Return Result.Value
        Catch ex As Exception
            CableSoft.NSTV.Log.NstvLog.WriteErrorLog(ex, Nothing)
            Throw
        Finally
            Result.Dispose()
        End Try
    End Function


    Public Function AnalyseSendLowCmd(ByVal bytRecvData() As Byte,
                                      ByVal CommandType As CmdType) As NagraCommandResult

        Dim Result As New ThreadLocal(Of NagraCommandResult)
        Dim RecvString As New ThreadLocal(Of String)
        Dim NagraResultValue As New ThreadLocal(Of Dictionary(Of String, String))
        RecvString.Value = Encoding.ASCII.GetString(bytRecvData, 2, bytRecvData.Length - 2).TrimEnd(Chr(0))
        Result.Value = New NagraCommandResult
        Result.Value.RecvString = RecvString.Value
        Result.Value.Entities = String.Empty
        Result.Value.ErrorCode = CableSoftCustomerError.CableSoft_Initial
        Result.Value.ErrorName = [Enum].GetName(GetType(CableSoftCustomerError), CableSoftCustomerError.CableSoft_Initial)
        Result.Value.Status = AckCode.Nack
        Result.Value.Ack = Integer.Parse(CableSoftCustomerError.CableSoft_Initial).ToString
        Result.Value.GwtRanNum = RecvString.Value.Substring(0, 9)
        Try
            Result.Value.Ack = RecvString.Value.Substring(FixCmd, 4)
            Select Case CommandType
                Case CmdType.Cmd1002
                    If RecvString.Value.Substring(32, 4) = "1000" Then
                        Result.Value.Status = AckCode.Ack
                    Else
                        Result.Value.Status = AckCode.Nack
                        Result.Value.ErrorCode = RecvString.Value.Substring(46, 4)
                        Result.Value.ErrorName = RecvString.Value.Substring(50, 4)
                        If (ErrorList IsNot Nothing) AndAlso (ErrorList.Count > 0) Then
                            If ErrorList.ContainsKey(Result.Value.ErrorCode) Then
                                Result.Value.ErrorName = ErrorList.Item(Result.Value.ErrorCode)
                            End If
                        End If
                    End If
                Case CmdType.Cmd0851
                    Result.Value.LowCmd = "0851"
                    If RecvString.Value.Substring(FixCmd, 4) = "2005" Then
                        Result.Value.ErrorCode = String.Empty
                        Result.Value.ErrorName = String.Empty
                        Result.Value.Status = AckCode.Ack
                        NagraResultValue.Value = GetNagraResultValue(RecvString.Value, CommandType)
                        If NagraResultValue.Value IsNot Nothing AndAlso NagraResultValue.Value.Count > 0 Then
                            Result.Value.Entities = NagraResultValue.Value.Item("ENTITIES")
                            Result.Value.Entity_data = NagraResultValue.Value.Item("ENTITY_DATA")
                        End If
                        Result.Value.RcvTranNum = RecvString.Value.Substring(FixCmd + 4, 9)    '送給Nagra的TranNum
                    Else
                        If RecvString.Value.Substring(FixCmd, 4) <> "2006" Then
                            Result.Value.Status = AckCode.Other
                            Result.Value.ErrorCode = RecvString.Value.Substring(FixCmd + 4 + 9 + 1, 4)
                            Result.Value.ErrorName = "CableSoft_Not_Know_Error"
                        Else
                            Result.Value.ErrorCode = RecvString.Value.Substring(FixCmd + 4 + 9 + 1, 4)
                            Result.Value.ErrorName = RecvString.Value.Substring(FixCmd + 4 + 9 + 1 + 4, 4)
                            Result.Value.Nack_Status = RecvString.Value.Substring(FixCmd + 4 + 9, 1)
                            Result.Value.RcvTranNum = RecvString.Value.Substring(FixCmd + 4, 9) '送給Nagra的TranNum
                        End If
                        If (ErrorList IsNot Nothing) AndAlso (ErrorList.Count > 0) Then
                            If ErrorList.ContainsKey(Result.Value.ErrorCode) Then
                                Result.Value.ErrorName = ErrorList.Item(Result.Value.ErrorCode)
                            End If
                        End If
                    End If
                Case CmdType.Cmd0126
                    Result.Value.LowCmd = "0126"
                    If RecvString.Value.Substring(FixCmd, 4) = "1003" Then
                        Result.Value.ErrorCode = String.Empty
                        Result.Value.ErrorName = String.Empty
                        Result.Value.Status = AckCode.Ack

                        NagraResultValue.Value = GetNagraResultValue(RecvString.Value, CommandType)
                        If (NagraResultValue.Value IsNot Nothing) AndAlso (NagraResultValue.Value.Count > 0) Then
                            Result.Value.NUID = NagraResultValue.Value.Item("NUID")
                            Result.Value.VUA = NagraResultValue.Value.Item("VUA")
                        End If
                        Result.Value.RcvTranNum = RecvString.Value.Substring(FixCmd + 4, 9)    '送給Nagra的TranNum
                    Else
                        If RecvString.Value.Substring(FixCmd, 4) <> "1003" Then
                            Result.Value.Status = AckCode.Other
                            Result.Value.ErrorCode = RecvString.Value.Substring(FixCmd + 4 + 9 + 1, 4)
                            Result.Value.ErrorName = "CableSoft_Not_Know_Error"
                        Else
                            Result.Value.ErrorCode = RecvString.Value.Substring(FixCmd + 4 + 9 + 1, 4)
                            Result.Value.ErrorName = RecvString.Value.Substring(FixCmd + 4 + 9 + 1 + 4, 4)
                            Result.Value.Nack_Status = RecvString.Value.Substring(FixCmd + 4 + 9, 1)
                            Result.Value.RcvTranNum = RecvString.Value.Substring(FixCmd + 4, 9) '送給Nagra的TranNum
                        End If
                        If (ErrorList IsNot Nothing) AndAlso (ErrorList.Count > 0) Then
                            If ErrorList.ContainsKey(Result.Value.ErrorCode) Then
                                Result.Value.ErrorName = ErrorList.Item(Result.Value.ErrorCode)
                            End If
                        End If

                    End If
            End Select
            Return Result.Value
        Catch ex As Exception
            Throw
        Finally
            Result.Dispose()
            RecvString.Dispose()
            NagraResultValue.Dispose()
        End Try
    End Function




#Region "IDisposable Support"
    Private disposedValue As Boolean ' 偵測多餘的呼叫

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                XmlFileIO.Write_Transaction_Number(Transaction_Number)
                If xmlFile IsNot Nothing Then
                    xmlFile.Dispose()
                    xmlFile = Nothing
                End If
                If mtxWait IsNot Nothing Then
                    mtxWait.Dispose()
                    mtxWait = Nothing
                End If
                ' TODO: 處置 Managed 狀態 (Managed 物件)。
            End If

            ' TODO: 釋放 Unmanaged 資源 (Unmanaged 物件) 並覆寫下面的 Finalize()。
            ' TODO: 將大型欄位設定為 null。
        End If
        Me.disposedValue = True
    End Sub

    ' TODO: 只有當上面的 Dispose(ByVal disposing As Boolean) 有可釋放 Unmanaged 資源的程式碼時，才覆寫 Finalize()。
    'Protected Overrides Sub Finalize()
    '    ' 請勿變更此程式碼。在上面的 Dispose(ByVal disposing As Boolean) 中輸入清除程式碼。
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' 由 Visual Basic 新增此程式碼以正確實作可處置的模式。
    Public Sub Dispose() Implements IDisposable.Dispose
        ' 請勿變更此程式碼。在以上的 Dispose 置入清除程式碼 (ByVal 視為布林值處置)。
        Dispose(True)
        GC.SuppressFinalize(Me)



    End Sub
#End Region

End Class








Public Class CheckStatusResult
    Implements IDisposable

    Property Success As Byte
    Property Answer_Code As Byte
    Public Sub New()

    End Sub




#Region "IDisposable Support"
    Private disposedValue As Boolean ' 偵測多餘的呼叫

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: 處置 Managed 狀態 (Managed 物件)。
            End If

            ' TODO: 釋放 Unmanaged 資源 (Unmanaged 物件) 並覆寫下面的 Finalize()。
            ' TODO: 將大型欄位設定為 null。
        End If
        Me.disposedValue = True
    End Sub

    ' TODO: 只有當上面的 Dispose(ByVal disposing As Boolean) 有可釋放 Unmanaged 資源的程式碼時，才覆寫 Finalize()。
    'Protected Overrides Sub Finalize()
    '    ' 請勿變更此程式碼。在上面的 Dispose(ByVal disposing As Boolean) 中輸入清除程式碼。
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' 由 Visual Basic 新增此程式碼以正確實作可處置的模式。
    Public Sub Dispose() Implements IDisposable.Dispose
        ' 請勿變更此程式碼。在以上的 Dispose 置入清除程式碼 (ByVal 視為布林值處置)。
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class
Public Class NagraCommandResult
    Implements IDisposable

    Public Property Nack_Status As String
    Public Property RcvTranNum As String
    Public Property Ack As String
    Public Property LowCmd As String
    Public Property CatRanNum As String
    Public Property GwtRanNum As String
    Public Property SendTime As Date
    Public Property RecvTime As Date
    Public Property ErrorCode() As String
    Public Property ErrorName() As String
    Public Property Status As BuilderLowerCmd.LowCmd.AckCode
    Public Property Entities() As String
    Public Property Entity_data() As String
    Public Property VUA As String
    Public Property NUID As String
    Public Property RecvString As String



#Region "IDisposable Support"
    Private disposedValue As Boolean ' 偵測多餘的呼叫

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: 處置 Managed 狀態 (Managed 物件)。
            End If

            ' TODO: 釋放 Unmanaged 資源 (Unmanaged 物件) 並覆寫下面的 Finalize()。
            ' TODO: 將大型欄位設定為 null。
        End If
        Me.disposedValue = True
    End Sub

    ' TODO: 只有當上面的 Dispose(ByVal disposing As Boolean) 有可釋放 Unmanaged 資源的程式碼時，才覆寫 Finalize()。
    'Protected Overrides Sub Finalize()
    '    ' 請勿變更此程式碼。在上面的 Dispose(ByVal disposing As Boolean) 中輸入清除程式碼。
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' 由 Visual Basic 新增此程式碼以正確實作可處置的模式。
    Public Sub Dispose() Implements IDisposable.Dispose
        ' 請勿變更此程式碼。在以上的 Dispose 置入清除程式碼 (ByVal 視為布林值處置)。
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class