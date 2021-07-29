Imports System.Threading
Imports System.Windows.Forms
Imports System.Text
Public Class LowCmd
    Implements IDisposable
    Private _Proto_Ver As Byte
    Private _Crypt_Ver As Byte
    Private _Key_Type As Byte
    Private _OPE_ID As Integer
    Private _SMS_ID As Integer
    Private _Root_Key As String
    Private _TalkKey As String
    Private Shared DBId As Integer
    Private MaxDBId As Integer = 65535
    Private xmlFile As CableSoft.CAS.XMLFileIO.XmlFileIO
    Private Const xmlFileName As String = "DBId.xml"
    Private Const SmsNodeName As String = "SMS"
    Private Const DBIdNodeName As String = "DBId"
    
    Public Property Proto_Ver As Byte
        Get
            Return _Proto_Ver
        End Get
        Set(value As Byte)
            _Proto_Ver = value
        End Set
    End Property
    Public Property Crypt_Ver As Byte
        Get
            Return _Crypt_Ver
        End Get
        Set(value As Byte)
            _Crypt_Ver = value + 3
        End Set
    End Property
    Public Property Key_Type As Byte
        Get
            Return _Key_Type
        End Get
        Set(value As Byte)
            _Key_Type = value
        End Set
    End Property
    Public Property OPE_ID As Integer
        Get
            Return _OPE_ID
        End Get
        Set(value As Integer)
            _OPE_ID = value
        End Set
    End Property
    Public Property SMS_ID As Integer
        Get
            Return _SMS_ID
        End Get
        Set(value As Integer)
            _SMS_ID = value
        End Set
    End Property
    Public Property Root_Key As String
        Get
            Return _Root_Key
        End Get
        Set(value As String)
            _Root_Key = value
        End Set
    End Property
    
    Public Sub New(ByVal Proto_Ver As Byte, ByVal Crypt_Ver As Byte,
                   ByVal Key_Type As Key_TypeEnum, ByVal OPE_ID As Integer,
                    ByVal SMS_ID As Integer, ByVal Root_Key As String)
        _Proto_Ver = Proto_Ver
        _Crypt_Ver = Crypt_Ver + 3
        _Key_Type = Key_Type

        _OPE_ID = OPE_ID
        _SMS_ID = SMS_ID
        _Root_Key = Root_Key
        Try
            If Not System.IO.File.Exists(Application.StartupPath & "\" & xmlFileName) Then
                Throw New Exception("找不到" & xmlFileName & "檔案!!")
            End If
            xmlFile = New CableSoft.CAS.XMLFileIO.XmlFileIO(Application.StartupPath & "\" & xmlFileName)
            DBId = xmlFile.ReadXmlNodeValue(SmsNodeName, DBIdNodeName, False)
            Dim maxTemp As Object = xmlFile.ReadXmlAttributeValue(SmsNodeName, DBIdNodeName, "Max", False)
            If maxTemp IsNot Nothing Then
                MaxDBId = Integer.Parse(maxTemp)
            End If
        Catch ex As Exception
            Throw ex
        End Try

    End Sub
    Public Function BuilderCheckState() As Byte()
        Dim byteData(7) As Byte
        byteData(0) = &H0
        byteData(1) = &H4
        byteData(2) = &HFF
        byteData(3) = &HFF
        byteData(4) = &HFF
        byteData(5) = &HFF
        byteData(6) = &H0
        byteData(7) = &H0
        Return byteData
    End Function
    Public Function ChkeckState(ByVal byteData() As Byte) As Boolean
        Try
            If (byteData Is Nothing) OrElse (byteData.Count < 8) Then
                Throw New Exception("傳入的資料不正確！")
            End If
            If byteData(0) <> &H1 Then
                Return False
            End If
            If byteData(1) <> &H6 Then
                Return False
            End If
            If byteData(2) <> &HFF Then
                Return False
            End If
            If byteData(3) <> &HFF Then
                Return False
            End If
            If byteData(4) <> &HFF Then
                Return False
            End If
            If byteData(5) <> &HFF Then
                Return False
            End If
            If byteData(6) <> &H0 Then
                Return False
            End If
            If byteData(7) <> &H0 Then
                Return False
            End If
        Catch ex As Exception
            Throw
        End Try
        Return True
    End Function
    'Public Function BuilderLowCmd(ByVal MsgId As Integer) As Byte()
    '    Return Create_Data_Section(MsgId)
    'End Function
    Private Function CreateFixData(ByVal byteData() As Byte) As SendLowCmd
        Dim strBuilder As New ThreadLocal(Of System.Text.StringBuilder)
        strBuilder.Value = New System.Text.StringBuilder
        Dim SendCmdStruct As SendLowCmd
        Try
            SendCmdStruct.SendData = byteData
            For i As Int32 = 0 To SendCmdStruct.SendData.Length - 1
                strBuilder.Value.Append("0X" & Right("0" & Hex(SendCmdStruct.SendData(i)), 2))
            Next
            SendCmdStruct.SendData = byteData
            SendCmdStruct.SendDataString = strBuilder.Value.ToString
            SendCmdStruct.Proto_Ver = _Proto_Ver
            SendCmdStruct.Crypt_Ver = _Crypt_Ver
            SendCmdStruct.Key_Type = SendCmdStruct.SendData(1) - 4
            SendCmdStruct.OPE_ID = _OPE_ID
            SendCmdStruct.SMS_ID = _SMS_ID            
            SendCmdStruct.Data_Len = BitConverter.ToInt16(New Byte() {SendCmdStruct.SendData(7), SendCmdStruct.SendData(6)}, 0)
            SendCmdStruct.DB_ID = BitConverter.ToInt16(New Byte() {SendCmdStruct.SendData(9), SendCmdStruct.SendData(8)}, 0)
            SendCmdStruct.MSG_ID = BitConverter.ToInt16(New Byte() {SendCmdStruct.SendData(11), SendCmdStruct.SendData(10)}, 0)
            SendCmdStruct.DB_Len = BitConverter.ToInt16(New Byte() {SendCmdStruct.SendData(13), SendCmdStruct.SendData(12)}, 0)
            SendCmdStruct.Data_Count = SendCmdStruct.SendData.Skip(14).ToArray.Take(
                        SendCmdStruct.SendData.Skip(14).ToArray.Length - 16).ToArray          
            SendCmdStruct.MAC = SendCmdStruct.SendData.Skip(14 + SendCmdStruct.Data_Count.Length).ToArray
            Return SendCmdStruct
        Catch ex As Exception
            Throw
        Finally
            strBuilder.Value.Clear()
            strBuilder.Dispose()
        End Try
    End Function
    Public Function Builder_SMS_CA_SET_CHARACTER_REQUEST(ByVal Card_SN As String, ByVal No As String, ByVal Cha As String) As SendLowCmd
        Try
            Return CreateFixData(Create_Data_Section(&H20C,
                                                    Create_Data_Body(&H20C,
                                                                     CreateCharacter_Data_Count(Card_SN, No, Cha))))
        Catch ex As Exception
            Throw
        End Try
    End Function
    Public Function Builder_SMS_CA_RESETCARDPIN_REQUEST(ByVal Card_SN As String) As SendLowCmd
        Try
            Return CreateFixData(Create_Data_Section(&H207,
                                                     Create_Data_Body(&H207,
                                                                      CreateOpenAccount_Data_Count(Card_SN))))


        Catch ex As Exception
            Throw
        End Try
    End Function
    Public Function Builder_SMS_CA_CANCEL_CHILD_REQUEST(ByVal Card_SN As String) As SendLowCmd
        Try
            Return CreateFixData(Create_Data_Section(&H403,
                                                     Create_Data_Body(&H403,
                                                                      CreateOpenAccount_Data_Count(Card_SN))))


        Catch ex As Exception
            Throw
        End Try
    End Function
    Public Function Builder_SMS_CA_SET_CHILD_REQUEST(ByVal Card_SN As String,
                                                     ByVal Parent_Card_SN As String,
                                                     ByVal Feed_Interval_Hour As String) As SendLowCmd
        Try
            Return CreateFixData(Create_Data_Section(&H402,
                                                     Create_Data_Body(&H402,
                                                                      CreateSetChild_Data_Count(Card_SN, Parent_Card_SN, Feed_Interval_Hour))))


        Catch ex As Exception
            Throw
        End Try
    End Function
    Public Function Builder_SMS_CA_SET_UNLOCK_REQUEST(ByVal Card_SN As String) As SendLowCmd
        Try
            Return CreateFixData(Create_Data_Section(&H204,
                                                     Create_Data_Body(&H204,
                                                                      CreateOpenAccount_Data_Count(Card_SN))))


        Catch ex As Exception
            Throw
        End Try
    End Function
    Public Function Builder_SMS_CA_SET_LOCK_REQUEST(ByVal Card_SN As String) As SendLowCmd
        Try
            Return CreateFixData(Create_Data_Section(&H203,
                                                     Create_Data_Body(&H203,
                                                                      CreateOpenAccount_Data_Count(Card_SN))))


        Catch ex As Exception
            Throw
        End Try
    End Function
    Public Function Builder_SMS_CA_STOP_ACCOUNT_REQUEST(ByVal Card_SN As String) As SendLowCmd
        Try
            Return CreateFixData(Create_Data_Section(&H202,
                                                     Create_Data_Body(&H202,
                                                                      CreateOpenAccount_Data_Count(Card_SN))))


        Catch ex As Exception
            Throw
        End Try
    End Function
    Public Function Builder_SMS_CA_ENTITLE_REQUEST(ByVal Card_SN As String, ByVal Product As String) As SendLowCmd
        Try
            Return CreateFixData(Create_Data_Section(&H20D,
                                                     Create_Data_Body(&H20D,
                                                                      CreateENTITLE_Data_Count(Card_SN, Product))))

            'Return CreateFixData(Create_Data_Body(&H201, CreateOpenAccount_Data_Count(Card_SN)))
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function Builder_SMS_CA_REPAIR_REQUEST(ByVal Card_SN As String, ByVal STBID As String) As SendLowCmd
        Try
            Return CreateFixData(Create_Data_Section(&H206,
                                                     Create_Data_Body(&H206,
                                                                      CreateRepair_Data_Count(Card_SN, STBID))))

            'Return CreateFixData(Create_Data_Body(&H201, CreateOpenAccount_Data_Count(Card_SN)))
        Catch ex As Exception
            Throw
        End Try
    End Function
    Public Function Builder_SMS_CA_OPEN_ACCOUNT_REQUEST(ByVal Card_SN As String) As SendLowCmd

        Try
            Return CreateFixData(Create_Data_Section(&H201,
                                                     Create_Data_Body(&H201,
                                                                      CreateOpenAccount_Data_Count(Card_SN))))

            'Return CreateFixData(Create_Data_Body(&H201, CreateOpenAccount_Data_Count(Card_SN)))
        Catch ex As Exception
            Throw
        End Try
    End Function
    Public Function Builder_SMS_CA_CREATE_SESSION_REQUEST() As SendLowCmd

        Try
            'Return CreateFixData(Create_Data_Section(1, Nothing))
            Return CreateFixData(Create_Data_Section(1, Create_Data_Body(1, CreateSession_Data_Count)))

        Catch ex As Exception
            Throw
        Finally

        End Try

    End Function
    Public Function AnalyseLowCmd(ByVal byteData() As Byte) As ResponseLowCmd
        Dim DataBody As New ThreadLocal(Of Byte())
        Dim strBuilder As New ThreadLocal(Of System.Text.StringBuilder)
        strBuilder.Value = New System.Text.StringBuilder
        Dim Response As ResponseLowCmd = Nothing
        If byteData.Count < 8 Then
            Throw New Exception("解析指令的長度不足8!")
        End If
        Response.Success = True
        Try

            Response.Proto_Ver = byteData(0)
            Response.Crypt_Ver = _Crypt_Ver
            Response.Key_Type = byteData(1) - Response.Crypt_Ver
            Response.OPE_ID = Convert.ToInt32(String.Format("{0}{1}", Right("0" & Hex(byteData(2)), 2), Right("0" & Hex(byteData(3)), 2)), 16)

            Response.SMS_ID = Convert.ToInt32(String.Format("{0}{1}", Right("0" & Hex(byteData(4)), 2), Right("0" & Hex(byteData(5)), 2)), 16)

            Response.DB_Len = Convert.ToInt32(String.Format("{0}{1}", Right("0" & Hex(byteData(6)), 2), Right("0" & Hex(byteData(7)), 2)), 16)

            If Response.DB_Len = 0 Then
                If Response.OPE_ID = &HFFFF AndAlso Response.SMS_ID = &HFFFF Then
                    Response.Success = False
                    Response.RecvData = byteData.Take(8).ToArray
                Else
                    Throw New Exception("無任何資料解析")
                End If

            Else
                Response.RecvData = byteData.Take(8 + Response.DB_Len).ToArray
            End If
            For i As Int32 = 0 To Response.RecvData.Length - 1
                strBuilder.Value.Append("0X" & Right("0" & Hex(Response.RecvData(i)), 2))
            Next
            Response.RecvDataString = strBuilder.ToString
            If Not Response.Success Then
                Return Response
            End If

            Select Case Response.Key_Type
                Case 0
                    DataBody.Value = Encry.Encryption.Decrypt3DES(Response.RecvData.Skip(8).ToArray, _Root_Key)
                Case 1
                    DataBody.Value = Encry.Encryption.DescryptDES(Response.RecvData.Skip(8).ToArray, _TalkKey)
                Case 2
                    DataBody.Value = Response.RecvData.Skip(8).ToArray

            End Select
            Dim str As String = Nothing
            For Each bty1 As Byte In DataBody.Value
                str = str & "," & Right("0" & Hex(bty1), 2)
            Next
            If DataBody.Value.Length < 20 Then
                Throw New Exception("回傳DATA_BODY有誤!")
            End If
            Response.DB_ID = Convert.ToInt32(String.Format("{0}{1}", Right("0" & Hex(DataBody.Value(0)), 2), Right("0" & Hex(DataBody.Value(1)), 2)), 16)
            Response.MSG_ID = Convert.ToInt32(String.Format("{0}{1}", Right("0" & Hex(DataBody.Value(2)), 2), Right("0" & Hex(DataBody.Value(3)), 2)), 16)
            Response.Data_Len = Convert.ToInt32(String.Format("{0}{1}", Right("0" & Hex(DataBody.Value(4)), 2), Right("0" & Hex(DataBody.Value(5)), 2)), 16)
            If Response.Data_Len = 0 Then
                Throw New Exception("回傳資料長度錯誤!")
            End If
            Response.ErrorCode = Convert.ToInt32(String.Format("{0}{1}{2}{3}",
                                                                Right("0" & Hex(DataBody.Value(6)), 2),
                                                                Right("0" & Hex(DataBody.Value(7)), 2),
                                                                Right("0" & Hex(DataBody.Value(8)), 2),
                                                                Right("0" & Hex(DataBody.Value(9)), 2)
                                                                ), 16)
            If Response.ErrorCode = 0 Then
                Select Case Response.MSG_ID
                    Case MsgIdEnum.CA_SMS_CREATE_SESSION_RESPONSE



                        Response.KeyLen = Convert.ToInt32(String.Format("{0}{1}",
                                                                        Right("0" & Hex(DataBody.Value(10)), 2),
                                                                        Right("0" & Hex(DataBody.Value(11)), 2)
                                                                        ), 16)
                        Response.KeyChar = Encoding.ASCII.GetString(DataBody.Value.Skip(12).ToArray.Take(Response.KeyLen).ToArray)
                        _TalkKey = Response.KeyChar
                End Select
            End If
            Response.MacOK = True
            Response.MAC = DataBody.Value.Skip(DataBody.Value.Length - 16).ToArray
            Dim bty() As Byte = Encry.Encryption.EncryMD5(DataBody.Value.Take(DataBody.Value.Length - 16).ToArray)
            If Not Encry.Encryption.EncryMD5(DataBody.Value.Take(DataBody.Value.Length - 16).ToArray).SequenceEqual(Response.MAC) Then
                Response.MacOK = False
            End If

            Return Response
            '不加密 (00=>0 根密加密,01=> 1 會話密碼加密, 10=>2 不加密, 11=>3 保留)
        Catch ex As Exception
            Throw
        Finally
            DataBody.Dispose()
            strBuilder.Value.Clear()
            strBuilder.Dispose()
            'Response.Dispose()
        End Try
    End Function
    Private Function Create_Data_Section(ByVal MsgId As Integer, ByVal DataBody() As Byte) As Byte()
        Dim lstByteData As New ThreadLocal(Of List(Of Byte))
        Dim byteDataBody As New ThreadLocal(Of Byte())
        lstByteData.Value = New List(Of Byte)
        Try
            '先取出DataBody資料
            byteDataBody.Value = DataBody
            '1Bytes
            lstByteData.Value.Add(_Proto_Ver)
            '1Bytes
            If MsgId = MsgIdEnum.SMS_CA_CREATE_SESSION_REQUEST Then
                lstByteData.Value.Add(_Crypt_Ver + 2) '不加密 (00=>0 根密加密,01=> 1 會話密碼加密, 10=>2 不加密, 11=>3 保留)
            Else
                lstByteData.Value.Add(_Crypt_Ver + _Key_Type)
            End If
            Dim bty() As Byte
            If MsgId <> 1 Then
                Select Case _Key_Type
                    Case 0
                        byteDataBody.Value = Encry.Encryption.Encrypt3DES(byteDataBody.Value, _Root_Key)

                    Case 1
                        byteDataBody.Value = Encry.Encryption.EncryptDES(byteDataBody.Value, _TalkKey)
                End Select
            End If


            '2Bytes
            lstByteData.Value.AddRange(BitConverter.GetBytes(_OPE_ID).ToArray.Reverse.Skip(2).ToArray)  '16bit
            '2Bytes
            lstByteData.Value.AddRange(BitConverter.GetBytes(_SMS_ID).ToArray.Reverse.Skip(2).ToArray)  '16bit
            'Data_len 2Bytes
            lstByteData.Value.AddRange(BitConverter.GetBytes(byteDataBody.Value.Length).ToArray.Reverse.Skip(2).ToArray)
            lstByteData.Value.AddRange(byteDataBody.Value)
            Return lstByteData.Value.ToArray
        Catch ex As Exception
            Throw
        Finally
            lstByteData.Value.Clear()
            lstByteData.Dispose()
            byteDataBody.Dispose()
        End Try

    End Function

    'Private Function Create_Data_Section(ByVal MsgId As Integer, ByVal lstPara As Dictionary(Of String, Object)) As Byte()
    '    Dim lstByteData As New ThreadLocal(Of List(Of Byte))
    '    Dim byteDataBody As New ThreadLocal(Of Byte())
    '    lstByteData.Value = New List(Of Byte)
    '    Try
    '        '先取出DataBody資料
    '        byteDataBody.Value = Create_Data_Body(MsgId, lstPara)
    '        '1Bytes
    '        lstByteData.Value.Add(_Proto_Ver)
    '        '1Bytes
    '        If MsgId = MsgIdEnum.SMS_CA_CREATE_SESSION_REQUEST Then
    '            lstByteData.Value.Add(_Crypt_Ver + 2) '不加密 (00=>0 根密加密,01=> 1 會話密碼加密, 10=>2 不加密, 11=>3 保留)
    '        Else
    '            lstByteData.Value.Add(_Crypt_Ver + _Key_Type)
    '        End If
    '        Dim bty() As Byte
    '        If MsgId <> 1 Then
    '            Select Case _Key_Type
    '                Case 0
    '                    byteDataBody.Value = Encry.Encryption.Encrypt3DES(byteDataBody.Value, _Root_Key)

    '                Case 1
    '                    byteDataBody.Value = Encry.Encryption.EncryptDES(byteDataBody.Value, _TalkKey)
    '            End Select
    '        End If


    '        '2Bytes
    '        lstByteData.Value.AddRange(BitConverter.GetBytes(_OPE_ID).ToArray.Reverse.Skip(2).ToArray)  '16bit
    '        '2Bytes
    '        lstByteData.Value.AddRange(BitConverter.GetBytes(_SMS_ID).ToArray.Reverse.Skip(2).ToArray)  '16bit
    '        'Data_len 2Bytes
    '        lstByteData.Value.AddRange(BitConverter.GetBytes(byteDataBody.Value.Length).ToArray.Reverse.Skip(2).ToArray)
    '        lstByteData.Value.AddRange(byteDataBody.Value)
    '        Return lstByteData.Value.ToArray
    '    Catch ex As Exception
    '        Throw
    '    Finally
    '        lstByteData.Value.Clear()
    '        lstByteData.Dispose()
    '        byteDataBody.Dispose()
    '    End Try

    'End Function

    Private Function Create_Data_Body(ByVal MsgId As Integer, ByVal Data_Count() As Byte) As Byte()
        Dim Result As New ThreadLocal(Of List(Of Byte))
        'Dim byteDataCount As New ThreadLocal(Of Byte())
        Dim byteMD5 As New ThreadLocal(Of Byte())
        Result.Value = New List(Of Byte)
        Try
            'DB_ID
            SyncLock xmlFile
                If DBId = MaxDBId Then
                    Interlocked.Exchange(DBId, 0)
                End If
                Result.Value.AddRange(BitConverter.GetBytes(DBId).ToArray.Reverse.Skip(2).ToArray)
                Interlocked.Increment(DBId)
                xmlFile.WriteXmlNodeValue(SmsNodeName, DBIdNodeName, DBId, False)
            End SyncLock

            Result.Value.AddRange(BitConverter.GetBytes(MsgId).ToArray.Reverse.Skip(2).ToArray)

            If MsgId = MsgIdEnum.SMS_CA_CREATE_SESSION_REQUEST Then
                Result.Value.Add(0)
                Result.Value.Add(0)
            Else
                Result.Value.AddRange(BitConverter.GetBytes(Data_Count.Length).ToArray.Reverse.Skip(2).ToArray)

            End If
            'Data_Count
            Result.Value.AddRange(Data_Count)
            Dim int As Int32 = 8 - ((Result.Value.Count + 16) Mod 8)
            If int <> 8 Then
                For i As Int32 = 1 To int
                    Result.Value.Add(&HF0)
                Next
            End If
            '將整個DataBody用MD5加密
            byteMD5.Value = Encry.Encryption.EncryMD5(Result.Value.ToArray)
            '加入MAC欄位
            Result.Value.AddRange(byteMD5.Value)
            Return Result.Value.ToArray
        Catch ex As Exception
            Throw
        Finally
            Result.Value.Clear()
            Result.Dispose()
            'byteDataCount.Dispose()
            byteMD5.Dispose()
        End Try
    End Function

    'Private Function Create_Data_Body(ByVal MsgId As Integer, ByVal lstPara As Dictionary(Of String, Object)) As Byte()
    '    Dim Result As New ThreadLocal(Of List(Of Byte))
    '    Dim byteDataCount As New ThreadLocal(Of Byte())
    '    Dim byteMD5 As New ThreadLocal(Of Byte())
    '    Result.Value = New List(Of Byte)
    '    Try
    '        'DB_ID
    '        SyncLock xmlFile
    '            If DBId = MaxDBId Then
    '                Interlocked.Exchange(DBId, 0)
    '            End If
    '            Result.Value.AddRange(BitConverter.GetBytes(DBId).ToArray.Reverse.Skip(2).ToArray)
    '            Interlocked.Increment(DBId)
    '            xmlFile.WriteXmlNodeValue(SmsNodeName, DBIdNodeName, DBId, False)
    '        End SyncLock
    '        'MSG_ID
    '        Result.Value.AddRange(BitConverter.GetBytes(MsgId).ToArray.Reverse.Skip(2).ToArray)
    '        byteDataCount.Value = Create_Data_Count(MsgId, lstPara) '取出Data_count資料
    '        'Data_Len
    '        If MsgId = MsgIdEnum.SMS_CA_CREATE_SESSION_REQUEST Then
    '            Result.Value.Add(0)
    '            Result.Value.Add(0)
    '        Else
    '            Result.Value.AddRange(BitConverter.GetBytes(byteDataCount.Value.Length).ToArray.Reverse.Skip(2).ToArray)

    '        End If
    '        'Data_Count
    '        Result.Value.AddRange(byteDataCount.Value)
    '        Dim int As Int32 = 8 - ((Result.Value.Count + 16) Mod 8)
    '        If int <> 8 Then
    '            For i As Int32 = 1 To int
    '                Result.Value.Add(&HF0)
    '            Next
    '        End If



    '        '將整個DataBody用MD5加密
    '        byteMD5.Value = Encry.Encryption.EncryMD5(Result.Value.ToArray)
    '        '加入MAC欄位
    '        Result.Value.AddRange(byteMD5.Value)


    '        Return Result.Value.ToArray
    '    Catch ex As Exception
    '        Throw
    '    Finally
    '        Result.Value.Clear()
    '        Result.Dispose()
    '        byteDataCount.Dispose()
    '        byteMD5.Dispose()
    '    End Try

    'End Function
    'Private Function Create_Common_Data_Count(ByVal Card_SN As String) As Byte()
    '    Dim byteData As New ThreadLocal(Of List(Of Byte))
    '    byteData.Value = New List(Of Byte)
    '    Try
    '        For Each Str As Char In Card_SN
    '            byteData.Value.AddRange(Encoding.ASCII.GetBytes(Str))
    '        Next
    '        Return byteData.Value.ToArray()
    '    Catch ex As Exception
    '        Throw
    '    Finally
    '        byteData.Dispose()
    '    End Try
    'End Function
    Private Function CreateSession_Data_Count() As Byte()
        Return New Byte() {&HF0, &HF0}
    End Function
    Private Function CreateCharacter_Data_Count(ByVal Card_SN As String, ByVal No As String, ByVal Cha As String)
        Dim byteData As New ThreadLocal(Of List(Of Byte))
        byteData.Value = New List(Of Byte)
        Try

            'For Each Str As Char In Card_SN
            '    byteData.Value.AddRange(Encoding.ASCII.GetBytes(Str))
            'Next
            byteData.Value.AddRange(CreateOpenAccount_Data_Count(Card_SN))
            If (String.IsNullOrEmpty(No)) OrElse (Not IsNumeric(No) OrElse (Int32.Parse(No) > 9)) Then
                Throw New Exception("特徵序號有誤!")
            End If
            byteData.Value.Add(Byte.Parse(No))
            Cha = Right("0000" & Cha, 4)
            For Each str As Char In Cha
                byteData.Value.AddRange(Encoding.ASCII.GetBytes(str))
            Next
            Return byteData.Value.ToArray()
        Catch ex As Exception
            Throw
        Finally
            byteData.Dispose()
        End Try
    End Function
    Private Function CreateOpenAccount_Data_Count(ByVal CARD_SN As String) As Byte()
        Dim byteData As New ThreadLocal(Of List(Of Byte))
        byteData.Value = New List(Of Byte)
        Try

            For Each Str As Char In CARD_SN
                byteData.Value.AddRange(Encoding.ASCII.GetBytes(Str))
            Next
            Return byteData.Value.ToArray()
        Catch ex As Exception
            Throw
        Finally
            byteData.Dispose()
        End Try
        ' Return Create_Common_Data_Count(lstPara.Item("CARD_SN".ToUpper))
    End Function
    Private Function DateToTime_T(ByVal ChangeDate As Date, Optional ChangeUTC As Boolean = True) As Byte()
        Dim time_t As New ThreadLocal(Of TimeSpan)
        Dim byteData As New ThreadLocal(Of List(Of Byte))
        byteData.Value = New List(Of Byte)
        Try

            Dim currtime As Date = New Date(1970, 1, 1, 0, 0, 0)
            If ChangeUTC Then
                ChangeDate = ChangeDate.ToUniversalTime
            End If
            time_t.Value = (ChangeDate - currtime)

            For i As Int32 = 0 To (Hex(time_t.Value.TotalSeconds).Length / 2) - 1
                byteData.Value.AddRange(BitConverter.GetBytes(
                                        Convert.ToInt32(Hex(time_t.Value.TotalSeconds).Substring(i * 2, 2), 16)).Take(1).ToArray)
            Next
            Return byteData.Value.ToArray.Clone
        Catch ex As Exception
            Throw
        Finally
            time_t.Dispose()
            byteData.Dispose()
        End Try
    End Function
    Private Function CreateSetChild_Data_Count(ByVal Card_SN As String,
                                               ByVal Parent_Card_SN As String,
                                               ByVal Feed_Interval_Hour As String) As Byte()
        Dim byteData As New ThreadLocal(Of List(Of Byte))
        byteData.Value = New List(Of Byte)
        Try
            byteData.Value.AddRange(CreateOpenAccount_Data_Count(Card_SN))
            If Parent_Card_SN.Length <> 16 Then
                Parent_Card_SN = Right("0000000000000000" & Parent_Card_SN, 16)
            End If
            byteData.Value.AddRange(CreateOpenAccount_Data_Count(Parent_Card_SN))
            If (String.IsNullOrEmpty(Feed_Interval_Hour)) OrElse Not (IsNumeric(Feed_Interval_Hour)) Then

                Throw New Exception("Feed_Interval_Hour 參數錯誤!")
            End If
            byteData.Value.AddRange(BitConverter.GetBytes(
                                    Int32.Parse(Feed_Interval_Hour)).Take(2).Reverse.ToArray())
            Return byteData.Value.ToArray
        Catch ex As Exception
            Throw
        Finally
            byteData.Value.Clear()
            byteData.Dispose()
        End Try
    End Function
    Private Function CreateENTITLE_Data_Count(ByVal Card_SN As String, ByVal ProductId As String) As Byte()
        Dim byteData As New ThreadLocal(Of List(Of Byte))
        byteData.Value = New List(Of Byte)

        Try
            If String.IsNullOrEmpty(Card_SN) Then
                Throw New Exception("CARD_SN無資料!")
            End If
            If String.IsNullOrEmpty(ProductId) Then
                Throw New Exception("ProductId無資料!")
            End If
            byteData.Value.AddRange(CreateOpenAccount_Data_Count(Card_SN))
            byteData.Value.Add(ProductId.Split(";").Count)
            For Each Str As String In ProductId.Split(";")


                If Str.Split(",").Count <> 5 Then
                    Throw New Exception("ProductId資料有誤!")
                Else
                    If Str.Split(",").ToList.Item(0).ToString.Length <> 1 Then
                        Throw New Exception("Enti_Type長度錯誤!")
                    End If
                    If Not IsNumeric(Str.Split(",").ToList.Item(0)) Then
                        Throw New Exception("Enti_Type格式錯誤!")
                    End If
                    If Not IsNumeric(Str.Split(",").ToList.Item(1)) Then
                        Throw New Exception("Prod_ID格式錯誤!")
                    End If
                    If Int32.Parse(Str.Split(",").ToList.Item(1)) > 1023 Then
                        Throw New Exception("Prod_ID資料錯誤!")
                    End If
                    If Not IsNumeric(Str.Split(",").ToList.Item(2)) Then
                        Throw New Exception("Tape_Ctrl格式錯誤!")
                    End If
                    If Int32.Parse(Str.Split(",").ToList.Item(2)) > 1 Then
                        Throw New Exception("Tape_Ctrl資料錯誤!")
                    End If
                    If Not IsDate(Str.Split(",").ToList.Item(3)) Then
                        Throw New Exception("Start_Time格式錯誤!")
                    End If
                    If Not IsDate(Str.Split(",").ToList.Item(4)) Then
                        Throw New Exception("End_Time格式錯誤!")
                    End If
                    byteData.Value.AddRange(BitConverter.GetBytes(Convert.ToInt32(Str.Split(",")(0) & "000000000000000", 2) +
                        Convert.ToInt32("000000000000000", 2) +
                        Int32.Parse(Str.Split(",").ToList.Item(1))).Take(2).ToArray.Reverse.ToArray)

                    byteData.Value.AddRange(New Byte() {Byte.Parse(Str.Split(",").ToList.Item(2))})

                    byteData.Value.AddRange(DateToTime_T(Date.Parse(Str.Split(",").ToList.Item(3))))
                    byteData.Value.AddRange(DateToTime_T(Date.Parse(Str.Split(",").ToList.Item(4))))

                End If
            Next

            Return byteData.Value.ToArray
        Catch ex As Exception
            Throw
        Finally
            byteData.Dispose()

        End Try
    End Function

    Private Function CreateRepair_Data_Count(ByVal Card_SN As String, ByVal STBID As String) As Byte()
        Dim byteData As New ThreadLocal(Of List(Of Byte))
        byteData.Value = New List(Of Byte)
        Try
            For Each Str As Char In Card_SN
                byteData.Value.AddRange(Encoding.ASCII.GetBytes(Str))
            Next
            If Not String.IsNullOrEmpty(STBID) Then
                STBID = (STBID.Replace(",", "").Replace(";", "") & New String("0", 30)).Substring(0, 30)
                For Each Str As Char In STBID
                    byteData.Value.AddRange(Encoding.ASCII.GetBytes(Str))
                Next
            End If

            Return byteData.Value.ToArray()
        Catch ex As Exception
            Throw
        Finally
            byteData.Dispose()
        End Try
    End Function
    'Private Function Create_Data_Count(ByVal MsgId As Integer, ByVal lstPara As Dictionary(Of String, Object)) As Byte()
    '    Dim lstByte As New ThreadLocal(Of List(Of Byte))
    '    lstByte.Value = New List(Of Byte)
    '    Try
    '        Select Case MsgId
    '            Case 1
    '                lstByte.Value.Add(&HF0)
    '                lstByte.Value.Add(&HF0)
    '            Case &H201
    '                If (lstPara.Count = 0) OrElse (Not lstPara.Keys.Contains("CARD_SN".ToUpper)) Then
    '                    Throw New Exception("無CARD_SN資料!")
    '                End If
    '                If lstPara.Item("Card_SN".ToUpper).ToString.Length <> 16 Then
    '                    Throw New Exception("CARD_SN長度錯誤！")
    '                End If
    '                Return Create_Common_Data_Count(lstPara.Item("CARD_SN".ToUpper))
    '            Case Else
    '                Throw New Exception(Hex(MsgId) & "不是有效的命令!")
    '        End Select
    '        Return lstByte.Value.ToArray
    '    Catch ex As Exception
    '        Throw ex
    '    Finally
    '        lstByte.Value.Clear()
    '        lstByte.Dispose()
    '    End Try
    'End Function

    '檢查數據內容是否為8Bit倍數
    Private Function CheckDataCount(ByVal byteData() As Byte) As Byte()
        Try
            Dim FillCount As Byte = byteData.Length Mod 8
            If FillCount > 0 Then
                Dim byteFill(FillCount) As Byte
                For i As Byte = 0 To byteFill.Length - 1
                    byteFill(i) = &HF0
                Next
                byteData.ToList.AddRange(byteFill)
            End If

        Catch ex As Exception
            Throw ex
        End Try
        Return byteData
    End Function
#Region "IDisposable Support"
    Private disposedValue As Boolean ' 偵測多餘的呼叫

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                If xmlFile IsNot Nothing Then
                    xmlFile.Dispose()
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
        Dim obj As ResponseLowCmd


    End Sub
#End Region

End Class


Public Structure ResponseLowCmd
    Public RecvData() As Byte
    Public RecvDataString As String
    Public Proto_Ver As Byte '1Byte
    Public Crypt_Ver As Byte
    Public Key_Type As Byte
    Public OPE_ID As Int32  '2Byes
    Public SMS_ID As Int32  '2Bytes
    Public DB_Len As Int32   '2Bytes
    Public DB_ID As Int32
    Public MSG_ID As Int32
    Public Data_Len As Int32
    Public MAC() As Byte
    Public ErrorCode As Byte
    Public KeyLen As Byte
    Public KeyChar As String
    Public MacOK As Boolean
    Public Success As Boolean
End Structure

Public Structure SendLowCmd
    Public SendData() As Byte
    Public SendDataString As String
    Public Proto_Ver As Byte '1Byte
    Public Crypt_Ver As Byte
    Public Key_Type As Byte
    Public OPE_ID As Int32  '2Byes
    Public SMS_ID As Int32  '2Bytes
    Public DB_Len As Int32   '2Bytes
    Public DB_ID As Int32
    Public MSG_ID As Int32
    Public Data_Len As Int32
    Public Data_Count() As Byte
    Public MAC() As Byte
End Structure
Public Enum MsgIdEnum As Integer
    SMS_CA_CREATE_SESSION_REQUEST = &H1
    CA_SMS_CREATE_SESSION_RESPONSE = &H8001
    SMS_CA_OPEN_ACCOUNT_REQUEST = &H201
    CA_SMS_OPEN_ACCOUNT_RESPONSE = &H8201
End Enum

Public Enum Key_TypeEnum
    RootKey = 0         '00
    SectionKey = 1      '01
    None = 2             '10
    reserve = 3          '11
End Enum