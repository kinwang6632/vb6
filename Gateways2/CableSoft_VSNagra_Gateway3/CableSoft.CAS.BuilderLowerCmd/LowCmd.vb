Imports System.Threading
Imports System.Windows.Forms
Imports System.Text
Imports CableSoft.CardLess
Imports CableSoft.CardLess.XMLFileIO

Public Class LowCmd
    Implements IDisposable
    Private SourceId As String
    Private DestId As String
    Private MoppId As String
    Private Shared Transaction_Number As UInteger
    Private MaxTransaction_Number As UInteger = 999999999
    Private xmlFile As CableSoft.CardLess.XMLFileIO.XmlFileIO
    Public Property UseUTC As Boolean = True
    Public Property Encoding As String = "UTF-8"
    Public Property mtxWait As Mutex = Nothing
    Public Property ErrorList As DataTable = Nothing

    Public Sub New(ByVal SourceId As String, ByVal DestId As String, ByVal MoppId As String)
        Me.SourceId = SourceId
        Me.DestId = DestId
        Me.MoppId = MoppId

      
        Me.ErrorList = ErrorList
        If tbSetNstv IsNot Nothing AndAlso tbSetNstv.Rows.Count > 0 Then
            Me.CanRecord = tbSetNstv.Rows(0).Item("Record") = 1
            Me.UseUTC = tbSetNstv.Rows(0).Item("UTC") = 1
            Me.Encoding = tbSetNstv.Rows(0).Item("Encoding")
        End If

        Try
            If Not System.IO.File.Exists(Application.StartupPath & "\" & xmlFileName) Then
                Throw New Exception("找不到" & xmlFileName & "檔案!!")
            End If
            xmlFile = New CableSoft.CardLess.XMLFileIO.XmlFileIO(Application.StartupPath & "\" & xmlFileName)
            'DBId = xmlFile.ReadXmlNodeValue(SmsNodeName, DBIdNodeName, False)
            Dim maxTemp As Object = 65535
            Try

                Transaction_Number = XMLFileIO.XmlFileIO.ReadTransaction_Number(Application.StartupPath & "\" & xmlFileName, SmsNodeName, DBIdNodeName)
                maxTemp = xmlFile.ReadXmlAttributeValue(SmsNodeName, DBIdNodeName, "Max", False)
            Catch ex As Exception
                Transaction_Number = 0
            End Try


            If maxTemp IsNot Nothing Then
                MaxTransaction_Number = Integer.Parse(maxTemp)
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

    Private Function CreateFixData(ByVal byteData() As Byte) As SendLowCmd
        Dim strBuilder As New ThreadLocal(Of System.Text.StringBuilder)
        strBuilder.Value = New System.Text.StringBuilder
        Dim SendCmdStruct As SendLowCmd
        Try
            SendCmdStruct.SendData = byteData
            For i As Int32 = 0 To SendCmdStruct.SendData.Length - 1
                If i = 0 Then
                    strBuilder.Value.Append(Right("0" & Hex(SendCmdStruct.SendData(i)), 2))
                Else
                    strBuilder.Value.Append(" " & Right("0" & Hex(SendCmdStruct.SendData(i)), 2))
                End If
                'strBuilder.Value.Append("0X" & Right("0" & Hex(SendCmdStruct.SendData(i)), 2))
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
    Public Overloads Function Builder_SMS_CA_ADV_CONTROL_PREVIEW_REQUEST(ByVal Card_SN As String,
                                                                         ByVal rwHigh As Dictionary(Of String, Object)) As SendLowCmd
        Try
            If (DBNull.Value.Equals(rwHigh.Item("Notes".ToUpper))) OrElse
             ((rwHigh.Item("Notes".ToUpper).ToString.Split(",").Count <> 4)) Then
                Throw New Exception("Notes 資料不正確!")
            End If
            Return Builder_SMS_CA_ADV_CONTROL_PREVIEW_REQUEST(Card_SN,
                                                              rwHigh.Item("Notes".ToUpper).ToString.Split(",")(0),
                                                                rwHigh.Item("Notes".ToUpper).ToString.Split(",")(1),
                                                                rwHigh.Item("Notes".ToUpper).ToString.Split(",")(2),
                                                                rwHigh.Item("Notes".ToUpper).ToString.Split(",")(3))
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Overloads Function Builder_SMS_CA_ADV_CONTROL_PREVIEW_REQUEST(ByVal Card_SN As String,
                                                                        ByVal PreviewDuration As Byte,
                                                                        ByVal WatchTime As Byte,
                                                                        ByVal TotalCount As Byte,
                                                                        ByVal TotalTime As Integer) As SendLowCmd
        Try
            Return CreateFixData(Create_Data_Section(&H20F,
                                                    Create_Data_Body(&H20F,
                                                                      CreatePreview_Data_Count(Card_SN, PreviewDuration,
                                                                                               WatchTime, TotalCount, TotalTime))))
        Catch ex As Exception
            Throw
        End Try
    End Function
    Public Overloads Function Builder_SMS_CA_SEND_SUPEROSD_REQUEST(ByVal Exp As String,
                                                                   ByVal rwHigh As Dictionary(Of String, Object)) As SendLowCmd
        Try

            If (DBNull.Value.Equals(rwHigh.Item("Notes".ToUpper))) OrElse
                ((rwHigh.Item("Notes".ToUpper).ToString.Split(",").Count <> 8)) Then
                Throw New Exception("Notes 資料不正確!")
            End If
            Return Builder_SMS_CA_SEND_SUPEROSD_REQUEST(Exp,
                                                    rwHigh.Item("Notes".ToUpper).ToString.Split(",")(0),
                                                    rwHigh.Item("Notes".ToUpper).ToString.Split(",")(1),
                                                    rwHigh.Item("Notes".ToUpper).ToString.Split(",")(2),
                                                    rwHigh.Item("Notes".ToUpper).ToString.Split(",")(3),
                                                    rwHigh.Item("Notes".ToUpper).ToString.Split(",")(4),
                                                    rwHigh.Item("Notes".ToUpper).ToString.Split(",")(5),
                                                    rwHigh.Item("Notes".ToUpper).ToString.Split(",")(6),
                                                    rwHigh.Item("Notes".ToUpper).ToString.Split(",")(7))


        Catch ex As Exception
            Throw
        End Try
    End Function
    Public Overloads Function Builder_SMS_CA_SEND_SUPEROSD_REQUEST(ByVal Exp As String,
                                                         ByVal Context As String, ByVal Style As Byte,
                                               ByVal Duration As Integer, ByVal ForcedDisplay As Byte,
                                               ByVal FontSize As Byte, ByVal FontColor As Byte,
                                               ByVal BackgroundColor As Byte, ByVal BackgroundArea As Byte) As SendLowCmd
        Try
            Return CreateFixData(Create_Data_Section(&H308,
                                                    Create_Data_Body(&H308,
                                                                      CreateSuperOSD_Data_Count(Exp, Context, Style, Duration,
                                                                                                ForcedDisplay, FontSize, FontColor,
                                                                                                BackgroundColor, BackgroundArea))))
        Catch ex As Exception
            Throw
        End Try
    End Function
    Public Function Builder_SMS_CA_UNLOCK_SERVICE_REQUEST(ByVal Exp As String) As SendLowCmd
        Try
            Return CreateFixData(Create_Data_Section(&H307,
                                                    Create_Data_Body(&H307,
                                                                      CreateUnLockService_Data_Count(Exp))))
        Catch ex As Exception
            Throw
        End Try
    End Function
    Public Overloads Function Builder_SMS_CA_LOCK_SERVICE_REQUEST(ByVal Exp As String,
                                                                 rwHight As Dictionary(Of String, Object)) As SendLowCmd
        Dim Para As New ThreadLocal(Of String)
        Dim ComPonentNum As New ThreadLocal(Of String)

        Try
            Para.Value = rwHight.Item("Notes".ToUpper).ToString.Split("@")(0)
            ComPonentNum.Value = rwHight.Item("Notes".ToUpper).ToString.Split("@")(1)
            Return Builder_SMS_CA_LOCK_SERVICE_REQUEST(Exp,
                                                       Para.Value.Split(",")(0),
                                                       Para.Value.Split(",")(1),
                                                        Para.Value.Split(",")(2),
                                                       Para.Value.Split(",")(3),
                                                      Para.Value.Split(",")(4),
                                                        Para.Value.Split(",")(5),
                                                       Para.Value.Split(",")(6),
                                                        ComPonentNum.Value)
        Catch ex As Exception
            Throw
        Finally
            Para.Dispose()
            ComPonentNum.Dispose()
        End Try
    End Function
    'ByVal Exp As String, ByVal LockFlag As Integer,
    '                                          ByVal Frequency As String,
    '                                          ByVal Fec_outer As Byte,
    '                                          ByVal Modulation As Byte,
    '                                          ByVal symbolrate As String, Fec_inner As Byte, ByVal PRC_PID As Integer,
    '                                          ByVal ComponentNum As Integer, ByVal CompType As Integer,
    '                                          ByVal CompPID As Integer, ByVal ECMPID As Integer

    Public Overloads Function Builder_SMS_CA_LOCK_SERVICE_REQUEST(ByVal Exp As String,
                                            ByVal LockFlag As Integer,
                                              ByVal Frequency As String,
                                              ByVal Fec_outer As Byte,
                                              ByVal Modulation As Byte,
                                              ByVal symbolrate As String, Fec_inner As Byte, ByVal PRC_PID As Integer,
                                              ByVal ComponentData As String) As SendLowCmd
        Try
            Return CreateFixData(Create_Data_Section(&H306,
                                                    Create_Data_Body(&H306,
                                                                      CreateLockService_Data_Count(Exp, LockFlag, Frequency,
                                                                                                    Fec_outer, Modulation, symbolrate,
                                                                                                    Fec_inner, PRC_PID, ComponentData))))

        Catch ex As Exception
            Throw
        End Try


    End Function

    'Public Overloads Function Builder_SMS_CA_LOCK_SERVICE_REQUEST(ByVal Exp As String) As SendLowCmd
    '    Try
    '        Return CreateFixData(Create_Data_Section(&H306,
    '                                                Create_Data_Body(&H306,
    '                                                                  CreateLockService_Data_Count(Exp, 0, "438.59",
    '                                                                                               12, 1, "999.99", 11, 1023, 2, 200, 1024, 1025))))

    '    Catch ex As Exception
    '        Throw
    '    End Try
    'End Function

    Public Function Builder_SMS_CA_SETSLOTMONEY(ByVal Card_SN As String, ByVal Slot_Id As Integer,
                                                ByVal Cred_Lim As Integer) As SendLowCmd
        Try
            Return CreateFixData(Create_Data_Section(&H208,
                                                    Create_Data_Body(&H208,
                                                                     CreateSetSlotMoney_Data_Count(Card_SN, Slot_Id, Cred_Lim))))



        Catch ex As Exception
            Throw
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
    Public Function Builder_SMS_CA_SEND_OSD_REQUEST(ByVal Exp As String,
                                                   ByVal Notes As String) As SendLowCmd
        Try
            Return CreateFixData(Create_Data_Section(&H304,
                                 Create_Data_Body(&H304,
                                         CreateSendOSD_Data_Count(Exp, Notes))))
        Catch ex As Exception
            Throw
        End Try
    End Function
    Public Function Builder_SMS_CA_SEND_OSD_REQUEST(ByVal Exp As String,
                                                    ByVal ConText As String,
                                                    ByVal Style As Integer, ByVal Duration As Integer) As SendLowCmd
        Try
            Return CreateFixData(Create_Data_Section(&H304,
                                 Create_Data_Body(&H304,
                                         CreateSendOSD_Data_Count(Exp, ConText, Style, Duration))))
        Catch ex As Exception
            Throw
        End Try
    End Function
    Public Function Builder_SMS_CA_SEND_EMAIL_REQUEST(ByVal Exp As String,
                                                      ByVal MailTitle As String,
                                                      ByVal MailContext As String,
                                                      ByVal Import As Integer) As SendLowCmd
        Try
            Return CreateFixData(Create_Data_Section(&H301,
                                    Create_Data_Body(&H301,
                                            CreateSendEMail_Data_Count(Exp, MailTitle, MailContext, Import))))
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
    'Public Function Builder_SMS_CA_ENTITLEEXT_REQUEST(ByVal ICC_NO As String,
    '                                                  ByVal rwHigh As Dictionary(Of String, Object)) As SendLowCmd
    '    Try
    '        Return CreateFixData(Create_Data_Section(&H20E,
    '                                                 Create_Data_Body(&H20E,
    '                                                                  CreateENTITLE_Data_Count(&H20E, ICC_NO, rwHigh))))
    '    Catch ex As Exception
    '        Throw
    '    End Try
    'End Function
    Public Function Builder_SMS_CA_CARD_REFRESH_REQUEST(ByVal Card_SN As String) As SendLowCmd
        Return CreateFixData(Create_Data_Section(&H401,
                                                     Create_Data_Body(&H401,
                                                                      CreateCardRefresh_Data_Count(Card_SN))))
    End Function

    Public Function Builder_SMS_CA_ENTITLE_REQUEST(ByVal MsgId As Integer, ByVal ICC_NO As String,
                                                   ByVal rwHigh As Dictionary(Of String, Object)) As SendLowCmd
        Try

            Return CreateFixData(Create_Data_Section(MsgId,
                                                     Create_Data_Body(MsgId,
                                                                      CreateENTITLE_Data_Count(MsgId, ICC_NO, rwHigh))))

            'Return CreateFixData(Create_Data_Body(&H201, CreateOpenAccount_Data_Count(Card_SN)))
        Catch ex As Exception
            Throw
        End Try
    End Function
    Public Function Builder_SMS_CA_ENTITLE_REQUEST(ByVal ICC_NO As String,
                                                   ByVal rwHigh As Dictionary(Of String, Object)) As SendLowCmd
        Try

            Return Builder_SMS_CA_ENTITLE_REQUEST(&H20D, ICC_NO, rwHigh)
            'Return CreateFixData(Create_Data_Section(&H20D,
            '                                         Create_Data_Body(&H20D,
            '                                                          CreateENTITLE_Data_Count(ICC_NO, rwHigh))))

            'Return CreateFixData(Create_Data_Body(&H201, CreateOpenAccount_Data_Count(Card_SN)))
        Catch ex As Exception
            Throw
        End Try
    End Function
    Public Function Builder_SMS_CA_ENTITLE_REQUEST(ByVal Card_SN As String,
                                                   ByVal Product As String) As SendLowCmd
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
    Public Function AnalyseSendLowCmd(ByVal byteData() As Byte) As SendLowCmd


        Return Nothing
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
                    Response.ErrorCode = -1
                    Response.ErrorName = "SendFormatError"
                    Response.RecvData = byteData.Take(8).ToArray
                Else
                    Response.Success = False
                    Response.ErrorCode = -2
                    Response.ErrorName = "NoDataAnalyse"
                    Response.RecvData = byteData.Take(8).ToArray
                    'Throw New Exception("無任何資料解析")
                End If
            Else
                Response.RecvData = byteData.Take(8 + Response.DB_Len).ToArray
            End If
            For i As Int32 = 0 To Response.RecvData.Length - 1
                If i = 0 Then
                    strBuilder.Value.Append(Right("0" & Hex(Response.RecvData(i)), 2))
                Else
                    strBuilder.Value.Append(" " & Right("0" & Hex(Response.RecvData(i)), 2))
                End If
                'strBuilder.Value.Append("0X" & Right("0" & Hex(Response.RecvData(i)), 2))
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
            If (Me.ErrorList IsNot Nothing) AndAlso (ErrorList.Rows.Count > 0) Then
                Dim rw() As DataRow = Me.ErrorList.Select("ErrorCode='" & Response.ErrorCode & "'")
                Response.ErrorName = Response.ErrorCode
                If rw IsNot Nothing AndAlso rw.Length = 1 Then
                    Response.ErrorName = rw(0).Item("ErrorName")
                End If
            Else
                Response.ErrorName = Response.ErrorCode
            End If
            If Response.ErrorCode = 0 Then
                Select Case Response.MSG_ID
                    Case MsgIdEnum.CA_SMS_CREATE_SESSION_RESPONSE
                        Response.KeyLen = Convert.ToInt32(String.Format("{0}{1}",
                                                                        Right("0" & Hex(DataBody.Value(10)), 2),
                                                                        Right("0" & Hex(DataBody.Value(11)), 2)
                                                                        ), 16)
                        For Each bytTalk As Byte In DataBody.Value.Skip(12).ToArray.Take(Response.KeyLen).ToArray
                            If String.IsNullOrEmpty(Response.TalkKey) Then
                                Response.TalkKey = Right("0" & Hex(bytTalk), 2)
                            Else
                                Response.TalkKey = Response.TalkKey & " " & Right("0" & Hex(bytTalk), 2)
                            End If
                        Next

                        'Response.TalkKey = Encoding.ASCII.GetString(DataBody.Value.Skip(12).ToArray.Take(Response.KeyLen).ToArray)
                        '_TalkKey = Response.TalkKey
                        _TalkKey = DataBody.Value.Skip(12).ToArray.Take(Response.KeyLen).ToArray
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
            Response.Success = False
            Response.ErrorCode = -3
            Response.ErrorName = "AnalyseOtherError"
            Response.RecvData = byteData.Take(8).ToArray
            Return Response
            'Throw
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
            'Dim bty() As Byte
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
            Return lstByteData.Value.ToArray.Clone
        Catch ex As Exception
            Throw
        Finally
            'lstByteData.Value.Clear()
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
        Dim int As New ThreadLocal(Of Int32)
        Result.Value = New List(Of Byte)
        Try
            'DB_ID
            'mtxWait.WaitOne()
            Try
                If Transaction_Number = MaxTransaction_Number Then
                    Interlocked.Exchange(Transaction_Number, 0)
                End If
                Result.Value.AddRange(BitConverter.GetBytes(Transaction_Number).ToArray.Reverse.Skip(2).ToArray)
                Interlocked.Increment(Transaction_Number)
                'xmlFile.WriteXmlNodeValue(SmsNodeName, DBIdNodeName, DBId, False)
                'XMLFileIO.XmlFileIO.WriteDBId(Application.StartupPath & "\" & xmlFileName,
                '                               DBId,
                '                              SmsNodeName,
                '                              DBIdNodeName)
            Finally
                'mtxWait.ReleaseMutex()
            End Try


            Result.Value.AddRange(BitConverter.GetBytes(MsgId).ToArray.Reverse.Skip(2).ToArray)

            If MsgId = MsgIdEnum.SMS_CA_CREATE_SESSION_REQUEST Then
                Result.Value.Add(0)
                Result.Value.Add(0)
            Else
                Result.Value.AddRange(BitConverter.GetBytes(Data_Count.Length).ToArray.Reverse.Skip(2).ToArray)

            End If
            'Data_Count
            Result.Value.AddRange(Data_Count)
            int.Value = 8 - ((Result.Value.Count + 16) Mod 8)
            If int.Value <> 8 Then
                For i As Int32 = 1 To int.Value
                    Result.Value.Add(&HF0)
                Next
            End If
            '將整個DataBody用MD5加密
            byteMD5.Value = Encry.Encryption.EncryMD5(Result.Value.ToArray)
            '加入MAC欄位
            Result.Value.AddRange(byteMD5.Value)
            Return Result.Value.ToArray.Clone
        Catch ex As Exception
            Throw
        Finally
            'Result.Value.Clear()
            Result.Dispose()
            'byteDataCount.Dispose()
            byteMD5.Dispose()
            int.Dispose()
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
    Private Function GetExpData(ByVal Exp As String) As Byte()
        Dim byteData As New ThreadLocal(Of List(Of Byte))
        byteData.Value = New List(Of Byte)
        Try
            '表達示長度
            byteData.Value.AddRange(BitConverter.GetBytes(
                                    System.Text.Encoding.ASCII.GetByteCount(Exp)).Take(1).ToArray)
            '表達示
            For Each chrExp As Char In Exp
                byteData.Value.AddRange(System.Text.Encoding.ASCII.GetBytes(chrExp))
            Next
            Return byteData.Value.ToArray.Clone
        Catch ex As Exception
            Throw
        Finally
            byteData.Value.Clear()
            byteData.Dispose()
        End Try
    End Function
    Private Function CreateSuperOSD_Data_Count(ByVal Exp As String,
                                               ByVal ConText As String, ByVal Style As Byte,
                                               ByVal Duration As Integer, ByVal ForcedDisplay As Byte,
                                               ByVal FontSize As Byte, ByVal FontColor As Byte,
                                               ByVal BackgroundColor As Byte, ByVal BackgroundArea As Byte) As Byte()
        Dim byteData As New ThreadLocal(Of List(Of Byte))
        byteData.Value = New List(Of Byte)
        Try
            byteData.Value.AddRange(GetExpData(Exp))
            byteData.Value.AddRange(GetConText(ConText, &H308))
            byteData.Value.Add(Style)
            byteData.Value.AddRange(
                BitConverter.GetBytes(Duration).Take(4).Reverse.ToArray())
            byteData.Value.Add(ForcedDisplay)
            byteData.Value.Add(FontSize)
            byteData.Value.Add(FontColor)
            byteData.Value.Add(BackgroundColor)
            byteData.Value.Add(BackgroundArea)
            Return byteData.Value.ToArray.Clone
        Catch ex As Exception
            Throw
        Finally
            byteData.Dispose()
        End Try

    End Function



    Private Function CreatePreview_Data_Count(ByVal Card_SN As String,
                                                                        ByVal PreviewDuration As Byte,
                                                                        ByVal WatchTime As Byte,
                                                                        ByVal TotalCount As Byte,
                                                                        ByVal TotalTime As Integer) As Byte()
        Dim byteData As New ThreadLocal(Of List(Of Byte))
        byteData.Value = New List(Of Byte)
        Try
            byteData.Value.AddRange(CreateOpenAccount_Data_Count(Card_SN))
            byteData.Value.Add(PreviewDuration)
            byteData.Value.Add(WatchTime)
            byteData.Value.Add(TotalCount)
            byteData.Value.AddRange(BitConverter.GetBytes(TotalTime).Take(2).Reverse.ToArray)
            Return byteData.Value.ToArray.Clone
        Catch ex As Exception
            Throw
        Finally
            byteData.Dispose()
        End Try

    End Function

    Private Function CreateUnLockService_Data_Count(ByVal Exp As String) As Byte()
        Dim byteData As New ThreadLocal(Of List(Of Byte))
        byteData.Value = New List(Of Byte)
        Try

            byteData.Value.AddRange(GetExpData(Exp))
            Return byteData.Value.ToArray.Clone
        Catch ex As Exception
            Throw
        Finally
            byteData.Dispose()
        End Try
    End Function
    Private Function CreateLockService_Data_Count(ByVal Exp As String, ByVal LockFlag As Integer,
                                                  ByVal Frequency As String,
                                                  ByVal Fec_outer As Byte,
                                                  ByVal Modulation As Byte,
                                                  ByVal symbolrate As String, Fec_inner As Byte, ByVal PRC_PID As Integer,
                                                  ByVal ComponentData As String) As Byte()
        'ByVal ComponentNum As Integer, ByVal CompType As Integer,
        'ByVal CompPID As Integer, ByVal ECMPID As Integer) As Byte()
        Dim byteData As New ThreadLocal(Of List(Of Byte))
        byteData.Value = New List(Of Byte)
        Try
            '表達示
            byteData.Value.AddRange(GetExpData(Exp))
            '切換方式
            byteData.Value.Add(LockFlag)
            '保留字段(2個 Byte)
            byteData.Value.Add(0)
            byteData.Value.Add(0)
            'CreateLockService_Data_Count(Exp, 0,  "4358.59",
            '                                                                                       12, 1, "999.99", 11, 1023, 2, 200, 1024, 1025)))

            '頻率
            byteData.Value.AddRange(
                    BitConverter.GetBytes(
                       Convert.ToInt32(Integer.Parse(
                            FormatNumber(Double.Parse(Frequency), 4).Replace(".", "").Replace(",", "")), 16)).Take(4).Reverse.ToArray())
            '保留字段(12個Bit)+前項糾錯外碼(4個Bit)
            byteData.Value.AddRange(BitConverter.GetBytes(0 + Fec_outer).Take(2).Reverse.ToArray)
            '調制方式
            byteData.Value.Add(Byte.Parse(Modulation))
            '符號率(28個Bit)+前項糾錯內碼(4個Bit)
            byteData.Value.AddRange(
                    BitConverter.GetBytes(
                        (Convert.ToInt32(Integer.Parse(
                         FormatNumber(Double.Parse(symbolrate), 4).Replace(".", "").Replace(",", "")), 16) + Fec_inner)).Take(4).Reverse.ToArray())
            '時鐘同步PID碼
            byteData.Value.AddRange(BitConverter.GetBytes(PRC_PID).Take(2).Reverse.ToArray)
            ''節目組數個數
            'byteData.Value.Add(Byte.Parse(ComponentNum))
            'For i As Integer = 0 To ComponentNum - 1
            '    '基楚流類型
            '    byteData.Value.Add(Byte.Parse(CompType))
            '    '基楚流PID
            '    byteData.Value.AddRange(BitConverter.GetBytes(CompPID).Take(2).Reverse.ToArray)
            '    'ECMPID
            '    byteData.Value.AddRange(BitConverter.GetBytes(ECMPID).Take(2).Reverse.ToArray)
            'Next

            '節目組數個數
            byteData.Value.Add(Byte.Parse(ComponentData.Split(";").Count))
            For Each Str As String In ComponentData.Split(";")
                byteData.Value.Add(Byte.Parse(Str.Split(",")(0)))
                byteData.Value.AddRange(BitConverter.GetBytes(Integer.Parse(Str.Split(",")(1))).Take(2).Reverse.ToArray)
                byteData.Value.AddRange(BitConverter.GetBytes(Integer.Parse(Str.Split(",")(2))).Take(2).Reverse.ToArray)
            Next
            '保留字段(固定為0x00)
            byteData.Value.AddRange(New Byte() {0, 0, 0, 0, 0, 0})
            Return byteData.Value.ToArray.Clone
        Catch ex As Exception
            Throw
        Finally
            byteData.Dispose()
        End Try
    End Function
    Private Function GetConText(ByVal Context As String, ByVal MsgId As Integer) As Byte()
        Dim byteData As New ThreadLocal(Of List(Of Byte))
        Dim EncodingType As New ThreadLocal(Of Encoding)
        byteData.Value = New List(Of Byte)
        Try
            EncodingType.Value = System.Text.Encoding.GetEncoding(Me.Encoding)
            '正文長度
            Select Case MsgId
                Case &H304
                    byteData.Value.AddRange(
                    BitConverter.GetBytes(
                    EncodingType.Value.GetByteCount(Context)).Take(1).ToArray())
                Case Else
                    byteData.Value.AddRange(
                    BitConverter.GetBytes(
                    EncodingType.Value.GetByteCount(Context)).Take(2).ToArray.Reverse.ToArray)
            End Select

            '正文
            For Each chrText As Char In Context
                byteData.Value.AddRange(EncodingType.Value.GetBytes(chrText))
            Next
            Return byteData.Value.ToArray.Clone
        Catch ex As Exception
            Throw
        Finally
            byteData.Dispose()
            EncodingType.Dispose()
        End Try
    End Function
    Private Function CreateSendOSD_Data_Count(ByVal Exp As String, ByVal Notes As String) As Byte()
        If String.IsNullOrEmpty(Notes) OrElse
            Notes.Split(",").Count <> 3 Then
            Throw New Exception("Notes 參數有誤")
        End If
        Return CreateSendOSD_Data_Count(Exp,
                                        Notes.Split(",")(0),
                                        Notes.Split(",")(1),
                                        Notes.Split(",")(2))
    End Function
    Private Function CreateSendOSD_Data_Count(ByVal Exp As String,
                                              ByVal Context As String,
                                                ByVal Style As Integer,
                                                ByVal Duration As Integer) As Byte()
        Dim byteData As New ThreadLocal(Of List(Of Byte))
        Dim EncodingType As New ThreadLocal(Of Encoding)
        byteData.Value = New List(Of Byte)

        Try
            '表達示
            byteData.Value.AddRange(GetExpData(Exp))
            '先取出編碼方式
            EncodingType.Value = System.Text.Encoding.GetEncoding(Me.Encoding)
            '正文長度
            'byteData.Value.AddRange(
            '    BitConverter.GetBytes(
            '        EncodingType.Value.GetByteCount(Context)).Take(1).ToArray())
            ''正文
            'For Each chrText As Char In Context
            '    byteData.Value.AddRange(EncodingType.Value.GetBytes(chrText))
            'Next
            byteData.Value.AddRange(GetConText(Context, &H304))
            'Style
            byteData.Value.Add(Byte.Parse(Style))
            'Duration
            byteData.Value.AddRange(
                    BitConverter.GetBytes(Duration).Take(4).ToArray.Reverse.ToArray())
            Return byteData.Value.ToArray.Clone
        Catch ex As Exception
            Throw
        Finally
            byteData.Dispose()
            EncodingType.Dispose()
        End Try

    End Function

    Private Function CreateSendEMail_Data_Count(ByVal Exp As String,
                                                ByVal EmailTitle As String,
                                                ByVal EMailContext As String, ByVal Import As Integer) As Byte()
        Dim byteData As New ThreadLocal(Of List(Of Byte))
        Dim EncodingType As New ThreadLocal(Of Encoding)
        byteData.Value = New List(Of Byte)

        Try
            '表達示長度
            'byteData.Value.AddRange(BitConverter.GetBytes(
            '                        System.Text.Encoding.ASCII.GetByteCount(Exp)).Take(1).ToArray)
            ''表達示
            'For Each chrExp As Char In Exp
            '    byteData.Value.AddRange(System.Text.Encoding.ASCII.GetBytes(chrExp))
            'Next

            byteData.Value.AddRange(GetExpData(Exp))
            '先取出編碼方式
            EncodingType.Value = System.Text.Encoding.GetEncoding(Me.Encoding)
            '標題長度
            byteData.Value.AddRange(
                BitConverter.GetBytes(
                    EncodingType.Value.GetByteCount(EmailTitle)).Take(1).ToArray())
            '標題
            For Each chrTitle As Char In EmailTitle
                byteData.Value.AddRange(EncodingType.Value.GetBytes(chrTitle))
            Next
            '正文長度
            byteData.Value.AddRange(
                BitConverter.GetBytes(EncodingType.Value.GetByteCount(EMailContext)).Take(1).ToArray())
            '正文內容
            For Each chrConText As Char In EMailContext
                byteData.Value.AddRange(EncodingType.Value.GetBytes(chrConText))
            Next
            '重要性
            byteData.Value.Add(Byte.Parse(Import))
            'Select Case Me.Encoding.ToUpper
            '    Case "UTF-8".ToUpper
            '        EncodingType.Value = System.Text.Encoding.GetEncoding("UTF-8")
            '        ''標題長度
            '        'byteData.Value.AddRange(BitConverter.GetBytes(
            '        '                        System.Text.Encoding.UTF8.GetBytes(EmailTitle).Length).Reverse.ToArray.Take(1).ToArray)
            '        ''標題
            '        'For Each chrEmail As Char In EmailTitle
            '        '    byteData.Value.AddRange(System.Text.Encoding.UTF8.GetBytes(chrEmail))
            '        'Next
            '        ''正文長度
            '        'byteData.Value.AddRange(BitConverter.GetBytes(
            '        '                        System.Text.Encoding.UTF8.GetBytes(EmailTitle).Length).Reverse.ToArray.Take(1).ToArray)
            '    Case "UTF-32".ToUpper
            '        '標題長度
            '        'byteData.Value.AddRange(BitConverter.GetBytes(
            '        '                        System.Text.Encoding.UTF8.GetBytes(EmailTitle).Length).Reverse.ToArray.Take(1).ToArray)
            '        ''標題
            '        'For Each chrExp As Char In Exp
            '        '    byteData.Value.AddRange(System.Text.Encoding.UTF8.GetBytes(chrExp))
            '        'Next
            'End Select


            Return byteData.Value.ToArray.Clone
        Catch ex As Exception
            Throw
        Finally
            byteData.Dispose()
            EncodingType.Dispose()
        End Try
    End Function
    Private Function CreateCardRefresh_Data_Count(ByVal Card_SN As String) As Byte()
        Dim byteData As New ThreadLocal(Of List(Of Byte))
        byteData.Value = New List(Of Byte)
        Try
            byteData.Value.AddRange(CreateOpenAccount_Data_Count(Card_SN))
            byteData.Value.Add(1)
            Return byteData.Value.ToArray.Clone
        Catch ex As Exception
            Throw
        Finally
            byteData.Dispose()
        End Try
    End Function
    Private Function CreateSetSlotMoney_Data_Count(ByVal Card_SN As String,
                                                   ByVal Slot_ID As Integer,
                                                   ByVal Cred_Lim As Integer) As Byte()
        Dim byteData As New ThreadLocal(Of List(Of Byte))
        byteData.Value = New List(Of Byte)
        Try
            byteData.Value.AddRange(CreateOpenAccount_Data_Count(Card_SN))
            byteData.Value.Add(Byte.Parse(Slot_ID))
            byteData.Value.AddRange(BitConverter.GetBytes(Cred_Lim).Take(2).Reverse.ToArray())
            Return byteData.Value.ToArray.Clone
        Catch ex As Exception
            Throw
        Finally
            byteData.Dispose()
        End Try
    End Function
    Private Function CreateCharacter_Data_Count(ByVal Card_SN As String, ByVal No As String, ByVal Cha As String) As Byte()
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
            If Not IsNumeric(Cha) Then
                Throw New Exception("特徵值格式有誤!")
            End If

            byteData.Value.Add(Byte.Parse(No))
            byteData.Value.AddRange(BitConverter.GetBytes(UInt32.Parse(Cha)).Take(4).ToArray.Reverse.ToArray)

            'byteData.Value.AddRange(BitConverter.GetBytes(UInt32.Parse(Convert.ToUInt64(Cha, 16))).Take(4).ToArray.Reverse.ToArray)


            'For i As Int32 = 0 To 
            '    If i = 0 ThenD:\NSTV\CableSoft.CAS.Encry\CableSoft.CAS.Encry\Encryption.vb
            '        byteData.Value.AddRange(BitConverter.GetBytes(UInt64.Parse(Convert.ToUInt64(STBID.Substring(0, 12), 16))).Take(6).ToArray.Reverse.ToArray)
            '    Else
            '        byteData.Value.AddRange(BitConverter.GetBytes(UInt64.Parse(Convert.ToUInt64(STBID.Substring(i + 12, 12), 16))).Take(6).ToArray.Reverse.ToArray)
            '    End If

            'Next
            'For Each str As Char In Cha
            '    byteData.Value.AddRange(Encoding.ASCII.GetBytes(str))
            'Next
            Return byteData.Value.ToArray().Clone
        Catch ex As Exception
            Throw
        Finally
            'byteData.Value.Clear()
            byteData.Dispose()
        End Try
    End Function
    Private Function CreateOpenAccount_Data_Count(ByVal CARD_SN As String) As Byte()
        Dim byteData As New ThreadLocal(Of List(Of Byte))
        byteData.Value = New List(Of Byte)
        Try

            For Each Str As Char In CARD_SN
                byteData.Value.AddRange(System.Text.Encoding.ASCII.GetBytes(Str))
            Next
            Return byteData.Value.ToArray().Clone
        Catch ex As Exception
            Throw
        Finally
            'byteData.Value.Clear()
            byteData.Dispose()
        End Try
        ' Return Create_Common_Data_Count(lstPara.Item("CARD_SN".ToUpper))
    End Function
    Private Function DateToTime_T(ByVal ChangeDate As Date, Optional ChangeUTC As Boolean = False) As Byte()
        Dim time_t As New ThreadLocal(Of TimeSpan)
        Dim byteData As New ThreadLocal(Of List(Of Byte))
        byteData.Value = New List(Of Byte)
        Try

            Dim currtime As Date = New Date(1970, 1, 1, 0, 0, 0)
            If ChangeUTC Then
                ChangeDate = ChangeDate.ToUniversalTime
            End If
            time_t.Value = (ChangeDate - currtime)

            'For i As Int32 = 0 To (Hex(time_t.Value.TotalSeconds).Length / 2) - 1
            '    byteData.Value.AddRange(BitConverter.GetBytes(
            '                            Convert.ToInt32(Hex(time_t.Value.TotalSeconds).Substring(i * 2, 2), 16)).Take(1).ToArray)
            'Next

            Return BitConverter.GetBytes(Int32.Parse(Math.Round(time_t.Value.TotalSeconds, 0))).Reverse.ToArray.Clone
            'Return byteData.Value.ToArray.Clone
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
            Return byteData.Value.ToArray.Clone
        Catch ex As Exception
            Throw
        Finally
            'byteData.Value.Clear()
            byteData.Dispose()
        End Try
    End Function
    Private Overloads Function CreateENTITLE_Data_Count(ByVal MsgId As Integer,
                                                        ByVal ICC_NO As String,
                                          ByVal rwHigh As Dictionary(Of String, Object)) As Byte()
        Dim byteData As New ThreadLocal(Of List(Of Byte))
        Dim subscription_begin_date As New ThreadLocal(Of DateTime)
        Dim Notes As New ThreadLocal(Of String)
        Dim subscription_end_date As New ThreadLocal(Of DateTime)

        byteData.Value = New List(Of Byte)

        Try
            If DBNull.Value.Equals(rwHigh.Item("ICC_NO".ToUpper)) Then
                Throw New Exception("ICC_NO無資料!")
            End If
            If DBNull.Value.Equals(rwHigh.Item("NOTES".ToUpper)) Then
                Throw New Exception("NOTES無資料!")
            End If
            If rwHigh.Item("NOTES".ToUpper).ToString.Split(";").Count > 1 Then
                Notes.Value = rwHigh.Item("NOTES".ToUpper).ToString.Split(";").ToList(1)
            Else
                Notes.Value = rwHigh.Item("NOTES".ToUpper).ToString.Split(";").ToList(0)
            End If

            byteData.Value.AddRange(CreateOpenAccount_Data_Count(ICC_NO))
            byteData.Value.Add(Notes.Value.Split(",").Count)
            For Each Str As String In Notes.Value.Split(",")
                subscription_begin_date.Value = Now
                subscription_end_date.Value = New DateTime(2020, 12, 31, 23, 59, 59)
                '大於1代表授權,0代表反授權
                If Str.IndexOf("~") >= 0 Then
                    If MsgId = &H20D Then
                        byteData.Value.AddRange(BitConverter.GetBytes(Convert.ToInt32("1000000000000000", 2) +
                            Convert.ToInt32("000000000000000", 2) +
                                Int32.Parse(Str.Split("~").ToList.Item(1))).Take(2).ToArray.Reverse.ToArray)
                    Else
                        '擴展授權
                        byteData.Value.AddRange(BitConverter.GetBytes(Convert.ToInt32("100000000000000000000000", 2) +
                                                    Convert.ToInt32("00000000000000000000000", 2) +
                                                        Int32.Parse(Str.Split("~").ToList.Item(1))).Take(3).ToArray.Reverse.ToArray)
                    End If

                    Select Case Str.Split("~").ToList.Item(0)
                        Case "A"
                            If Not DBNull.Value.Equals(rwHigh.Item("subscription_begin_date".ToUpper)) Then
                                subscription_begin_date.Value = Date.Parse(rwHigh.Item("subscription_begin_date".ToUpper))
                            End If
                            If Not DBNull.Value.Equals(rwHigh.Item("subscription_end_date".ToUpper)) Then
                                subscription_end_date.Value = Date.Parse(rwHigh.Item("subscription_end_date".ToUpper))
                            End If
                            If (subscription_begin_date.Value.Hour = 0) AndAlso
                                    (subscription_begin_date.Value.Minute = 0) AndAlso
                                    (subscription_begin_date.Value.Second = 0) Then
                                subscription_begin_date.Value.AddHours(0)
                                subscription_begin_date.Value.AddMinutes(0)
                                subscription_begin_date.Value.AddSeconds(0)
                            End If
                            If (subscription_end_date.Value.Hour = 0) AndAlso
                                   (subscription_end_date.Value.Minute = 0) AndAlso
                                   (subscription_end_date.Value.Second = 0) Then
                                subscription_end_date.Value.AddHours(23)
                                subscription_end_date.Value.AddMinutes(59)
                                subscription_end_date.Value.AddSeconds(59)
                            End If
                        Case "B"

                            subscription_begin_date.Value = New DateTime(
                                                                                    Int32.Parse(Str.Split("~").ToList.Item(2).Substring(0, 4)),
                                                                                    Int32.Parse(Str.Split("~").ToList.Item(2).Substring(4, 2)),
                                                                                    Int32.Parse(Str.Split("~").ToList.Item(2).Substring(6, 2)),
                                                                                    0, 0, 0)

                            subscription_end_date.Value = New DateTime(
                                                                                    Int32.Parse(Str.Split("~").ToList.Item(3).Substring(0, 4)),
                                                                                    Int32.Parse(Str.Split("~").ToList.Item(3).Substring(4, 2)),
                                                                                    Int32.Parse(Str.Split("~").ToList.Item(3).Substring(6, 2)),
                                                                                    23, 59, 59)
                        Case Else
                            Throw New Exception("NOTES格式有誤!")
                    End Select

                Else
                    If MsgId = &H20D Then
                        byteData.Value.AddRange(BitConverter.GetBytes(Convert.ToInt32("0000000000000000", 2) +
                           Convert.ToInt32("000000000000000", 2) +
                               Int32.Parse(Str)).Take(2).ToArray.Reverse.ToArray)
                    Else
                        byteData.Value.AddRange(BitConverter.GetBytes(Convert.ToInt32("000000000000000000000000", 2) +
                        Convert.ToInt32("00000000000000000000000", 2) +
                            Int32.Parse(Str)).Take(3).ToArray.Reverse.ToArray)
                    End If


                End If


                'byteData.Value.AddRange(BitConverter.GetBytes(Convert.ToInt32(Str.Split(",")(0) & "000000000000000", 2) +
                '    Convert.ToInt32("000000000000000", 2) +
                '    Int32.Parse(Str.Split(",").ToList.Item(1))).Take(2).ToArray.Reverse.ToArray)

                'byteData.Value.AddRange(New Byte() {Byte.Parse(Str.Split(",").ToList.Item(2))})
                '錄像控制
                byteData.Value.AddRange(New Byte() {IIf(Me.CanRecord, 1, 0)})

                'byteData.Value.AddRange(DateToTime_T(Date.Parse(Str.Split(",").ToList.Item(3)), Me.UseUTC))
                'byteData.Value.AddRange(DateToTime_T(Date.Parse(Str.Split(",").ToList.Item(4)), Me.UseUTC))
                '開始時間與過期時間
                byteData.Value.AddRange(DateToTime_T(subscription_begin_date.Value, Me.UseUTC))
                byteData.Value.AddRange(DateToTime_T(subscription_end_date.Value, Me.UseUTC))

            Next
            Return byteData.Value.ToArray.Clone
        Catch ex As Exception
            Throw
        Finally
            'byteData.Value.Clear()
            byteData.Dispose()
            Notes.Dispose()
            subscription_begin_date.Dispose()
            subscription_end_date.Dispose()
        End Try
    End Function



    Private Overloads Function CreateENTITLE_Data_Count(ByVal ICC_NO As String,
                                              ByVal rwHigh As Dictionary(Of String, Object)) As Byte()
        Dim byteData As New ThreadLocal(Of List(Of Byte))
        Dim subscription_begin_date As New ThreadLocal(Of DateTime)
        Dim Notes As New ThreadLocal(Of String)
        Dim subscription_end_date As New ThreadLocal(Of DateTime)

        byteData.Value = New List(Of Byte)

        Try
            If DBNull.Value.Equals(rwHigh.Item("ICC_NO".ToUpper)) Then
                Throw New Exception("ICC_NO無資料!")
            End If
            If DBNull.Value.Equals(rwHigh.Item("NOTES".ToUpper)) Then
                Throw New Exception("NOTES無資料!")
            End If
            If rwHigh.Item("NOTES".ToUpper).ToString.Split(";").Count > 1 Then
                Notes.Value = rwHigh.Item("NOTES".ToUpper).ToString.Split(";").ToList(1)
            Else
                Notes.Value = rwHigh.Item("NOTES".ToUpper).ToString.Split(";").ToList(0)
            End If

            byteData.Value.AddRange(CreateOpenAccount_Data_Count(ICC_NO))
            byteData.Value.Add(Notes.Value.Split(",").Count)
            For Each Str As String In Notes.Value.Split(",")
                subscription_begin_date.Value = Now
                subscription_end_date.Value = New DateTime(2020, 12, 31, 23, 59, 59)
                '大於1代表授權,0代表反授權
                If Str.IndexOf("~") >= 0 Then
                    byteData.Value.AddRange(BitConverter.GetBytes(Convert.ToInt32("1000000000000000", 2) +
                            Convert.ToInt32("000000000000000", 2) +
                                Int32.Parse(Str.Split("~").ToList.Item(1))).Take(2).ToArray.Reverse.ToArray)
                    Select Case Str.Split("~").ToList.Item(0)
                        Case "A"
                            If Not DBNull.Value.Equals(rwHigh.Item("subscription_begin_date".ToUpper)) Then
                                subscription_begin_date.Value = Date.Parse(rwHigh.Item("subscription_begin_date".ToUpper))
                            End If
                            If Not DBNull.Value.Equals(rwHigh.Item("subscription_end_date".ToUpper)) Then
                                subscription_end_date.Value = Date.Parse(rwHigh.Item("subscription_end_date".ToUpper))
                            End If
                            If (subscription_begin_date.Value.Hour = 0) AndAlso
                                    (subscription_begin_date.Value.Minute = 0) AndAlso
                                    (subscription_begin_date.Value.Second = 0) Then
                                subscription_begin_date.Value.AddHours(0)
                                subscription_begin_date.Value.AddMinutes(0)
                                subscription_begin_date.Value.AddSeconds(0)
                            End If
                            If (subscription_end_date.Value.Hour = 0) AndAlso
                                   (subscription_end_date.Value.Minute = 0) AndAlso
                                   (subscription_end_date.Value.Second = 0) Then
                                subscription_end_date.Value.AddHours(23)
                                subscription_end_date.Value.AddMinutes(59)
                                subscription_end_date.Value.AddSeconds(59)
                            End If
                        Case "B"

                            subscription_begin_date.Value = New DateTime(
                                                                                    Int32.Parse(Str.Split("~").ToList.Item(2).Substring(0, 4)),
                                                                                    Int32.Parse(Str.Split("~").ToList.Item(2).Substring(4, 2)),
                                                                                    Int32.Parse(Str.Split("~").ToList.Item(2).Substring(6, 2)),
                                                                                    0, 0, 0)

                            subscription_end_date.Value = New DateTime(
                                                                                    Int32.Parse(Str.Split("~").ToList.Item(3).Substring(0, 4)),
                                                                                    Int32.Parse(Str.Split("~").ToList.Item(3).Substring(4, 2)),
                                                                                    Int32.Parse(Str.Split("~").ToList.Item(3).Substring(6, 2)),
                                                                                    23, 59, 59)
                        Case Else
                            Throw New Exception("NOTES格式有誤!")
                    End Select

                Else
                    byteData.Value.AddRange(BitConverter.GetBytes(Convert.ToInt32("0000000000000000", 2) +
                            Convert.ToInt32("000000000000000", 2) +
                                Int32.Parse(Str)).Take(2).ToArray.Reverse.ToArray)

                End If


                'byteData.Value.AddRange(BitConverter.GetBytes(Convert.ToInt32(Str.Split(",")(0) & "000000000000000", 2) +
                '    Convert.ToInt32("000000000000000", 2) +
                '    Int32.Parse(Str.Split(",").ToList.Item(1))).Take(2).ToArray.Reverse.ToArray)

                'byteData.Value.AddRange(New Byte() {Byte.Parse(Str.Split(",").ToList.Item(2))})
                '錄像控制
                byteData.Value.AddRange(New Byte() {IIf(Me.CanRecord, 1, 0)})

                'byteData.Value.AddRange(DateToTime_T(Date.Parse(Str.Split(",").ToList.Item(3)), Me.UseUTC))
                'byteData.Value.AddRange(DateToTime_T(Date.Parse(Str.Split(",").ToList.Item(4)), Me.UseUTC))
                '開始時間與過期時間
                byteData.Value.AddRange(DateToTime_T(subscription_begin_date.Value, Me.UseUTC))
                byteData.Value.AddRange(DateToTime_T(subscription_end_date.Value, Me.UseUTC))

            Next
            Return byteData.Value.ToArray.Clone
        Catch ex As Exception
            Throw
        Finally
            'byteData.Value.Clear()
            byteData.Dispose()
            Notes.Dispose()
            subscription_begin_date.Dispose()
            subscription_end_date.Dispose()
        End Try
    End Function


    Private Overloads Function CreateENTITLE_Data_Count(ByVal Card_SN As String, ByVal ProductId As String) As Byte()
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

                    byteData.Value.AddRange(DateToTime_T(Date.Parse(Str.Split(",").ToList.Item(3)), Me.UseUTC))
                    byteData.Value.AddRange(DateToTime_T(Date.Parse(Str.Split(",").ToList.Item(4)), Me.UseUTC))

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
        Dim arySTB As New ThreadLocal(Of List(Of UInt64))
        arySTB.Value = New List(Of UInt64)

        Try
            For Each Str As Char In Card_SN
                byteData.Value.AddRange(System.Text.Encoding.ASCII.GetBytes(Str))
            Next

            arySTB.Value.AddRange(New UInt64() {0, 0, 0, 0, 0})
            For i As Int32 = 0 To STBID.Split(",").Count - 1
                If i > 4 Then
                    Exit For
                End If
                '廠商格式改成,"12B000000000(固定)+STBNO後7碼" By Kin 2013/07/26
                'arySTB.Value.Item(i) = UInt64.Parse(STBID.Split(",")(i))
                arySTB.Value.Item(i) = Convert.ToUInt64("12B000000000", 16) +
                                                        UInt64.Parse(Right("0000000" & STBID.Split(",")(i), 7))

            Next
            For Each intStb As UInt64 In arySTB.Value
                byteData.Value.AddRange(BitConverter.GetBytes(intStb).Take(6).ToArray.Reverse.ToArray)
            Next

            'If Not String.IsNullOrEmpty(STBID) Then
            '    STBID = (STBID.Replace(",", "").Replace(";", "") & New String("0", 60)).Substring(0, 60)

            'End If
            'For i As Int32 = 0 To 4
            '    If i = 0 Then
            '        byteData.Value.AddRange(BitConverter.GetBytes(UInt64.Parse(Convert.ToUInt64(STBID.Substring(0, 12), 16))).Take(6).ToArray.Reverse.ToArray)
            '    Else
            '        byteData.Value.AddRange(BitConverter.GetBytes(UInt64.Parse(Convert.ToUInt64(STBID.Substring(i + 12, 12), 16))).Take(6).ToArray.Reverse.ToArray)
            '    End If

            'Next

            Return byteData.Value.ToArray.Clone
        Catch ex As Exception
            Throw
        Finally
            'byteData.Value.Clear()
            byteData.Dispose()
            arySTB.Dispose()
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
        Finally

        End Try
        Return byteData.Clone
    End Function
#Region "IDisposable Support"
    Private disposedValue As Boolean ' 偵測多餘的呼叫

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                XMLFileIO.Write_Transaction_Number(Application.StartupPath & "\" & xmlFileName,
                                               Transaction_Number,
                                              SmsNodeName,
                                              DBIdNodeName)
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
    Public ErrorCode As Integer
    Public ErrorName As String
    Public KeyLen As Byte
    Public TalkKey As String
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
    SMS_CA_STOP_ACCOUNT_REQUEST = &H202
    CA_SMS_STOP_ACCOUNT_RESPONSE = &H8202
    SMS_CA_SET_LOCK_REQUEST = &H203
    CA_SMS_SET_LOCK_RESPONSE = &H8203
    SMS_CA_SET_UNLOCK_REQUEST = &H204
    CA_SMS_SET_UNLOCK_RESPONSE = &H8204
    SMS_CA_REPAIR_REQUEST = &H206
    CA_SMS_REPAIR_RESPONSE = &H8206
    SMS_CA_RESETCARDPIN_REQUEST = &H207
    CA_SMS_RESETCARDPIN_RESPONSE = &H8207
    SMS_CA_SETSLOTMONEY_REQUEST = &H208
    CA_SMS_SETSLOTMONEY_RESPONSE = &H8208
    SMS_CA_SET_CHARACTER_REQUEST = &H20C
    CA_SMS_SET_CHARACTER_RESPONSE = &H820C
    SMS_CA_ENTITLE_REQUEST = &H20D
    CA_SMS_ENTITLE_RESPONSE = &H820D
    SMS_CA_ENTITLEEXT_REQUEST = &H20E
    CA_SMS_ENTITLEEXT_RESPONSE = &H820E
    SMS_CA_SET_ADV_CONTROL_PREVIEW_REQUEST = &H20F
    CA_SMS_SET_ADV_CONTROL_PREVIEW_RESPONSE = &H820F
    SMS_CA_SEND_EMAIL_REQUEST = &H301
    CA_SMS_SEND_EMAIL_RESPONSE = &H8301
    SMS_CA_SEND_OSD_REQUEST = &H304
    CA_SMS_SEND_OSD_RESPONSE = &H8304
    SMS_CA_LOCK_SERVICE_REQUEST = &H306
    CA_SMS_LOCK_SERVICE_RESPONSE = &H8306
    SMS_CA_UNLOCK_SERVICE_REQUEST = &H307
    CA_SMS_UNLOCK_SERVICE_RESPONSE = &H8307
    SMS_CA_SEND_SUPEROSD_REQUEST = &H308
    CA_SMS_SEND_SUPEROSD_RESPONSE = &H8308
    SMS_CA_CARD_REFRESH_DATA_REQUEST = &H401
    CA_SMS_CARD_REFRESH_DATA_RESPONSE = &H8401
    SMS_CA_SET_CHILD_REQUEST = &H402
    CA_SMS_SET_CHILD_RESPONSE = &H8402
    SMS_CA_CANCEL_CHILD_REQUEST = &H403
    CA_SMS_CANCEL_CHILD_RESPONSE = &H8403

End Enum

Public Enum Key_TypeEnum
    RootKey = 0         '00
    SectionKey = 1      '01
    None = 2             '10
    reserve = 3          '11

End Enum