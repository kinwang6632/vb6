Imports System.Data
Imports CableSoft.Nagra.BuilderLowerCmd
Imports System.Threading

Public Class HighCmd
    Implements IDisposable

    Private tbSetHighCmd As DataTable = Nothing
    Private tbSetLowCmd As DataTable = Nothing
    Private LowerCmd As LowCmd = Nothing
    Private tbSetNSTV As DataTable = Nothing
    'Public Sub New(ByVal tbSetHighCmd As DataTable,
    '               ByVal tbSetLowCmd As DataTable,
    '               ByVal LowCmd As LowCmd)
    '    Me.tbSetHighCmd = tbSetHighCmd
    '    Me.tbSetLowCmd = tbSetLowCmd
    '    Me.LowerCmd = LowCmd
    'End Sub
    Public Sub New(ByVal tbSetHighCmd As DataTable,
                   ByVal tbSetLowCmd As DataTable,
                   ByVal tbSetNSTV As DataTable,
                   ByVal LowCmd As LowCmd)
        Me.tbSetHighCmd = tbSetHighCmd
        Me.tbSetLowCmd = tbSetLowCmd
        Me.LowerCmd = LowCmd
        Me.tbSetNSTV = tbSetNSTV
    End Sub
    Public Function GetSendDataList(ByVal rwDictionary As Dictionary(Of String, Object)) As List(Of SendLowCmd)
        Dim result As New ThreadLocal(Of List(Of SendLowCmd))
        Dim rwHigh As New ThreadLocal(Of DataRow())
        Dim rwLow As New ThreadLocal(Of DataRow())
        Dim aryLow As New ThreadLocal(Of Dictionary(Of String, Object))
        aryLow.Value = New Dictionary(Of String, Object)
        'Dim SourceRw As New ThreadLocal(Of DataRow)

        Try
            result.Value = New List(Of SendLowCmd)
            'SourceRw.Value = rw.Table.NewRow
            'SourceRw.Value.ItemArray = rw.ItemArray
            rwHigh.Value = tbSetHighCmd.Select(String.Format("CodeNo='{0}'", rwDictionary.Item("HIGH_LEVEL_CMD_ID".ToUpper)))
            If (rwHigh.Value Is Nothing) OrElse (rwHigh.Value.Length = 0) Then
                Throw New Exception(String.Format("公司別={0}，找不到SEQNO={1}高階命令{2}!",
                                                  rwDictionary.Item("CompCode".ToUpper),
                                                  rwDictionary.Item("SEQNO".ToUpper),
                                                  rwDictionary.Item("HIGH_LEVEL_CMD_ID".ToUpper)))
            End If
            For Each Str As String In rwHigh.Value(0).Item("RunLowCmd").ToString.Split(",")
                rwLow.Value = tbSetLowCmd.Select(String.Format("CodeNo='{0}'", Str.Trim))
                If (rwLow.Value Is Nothing) OrElse (rwLow.Value.Length = 0) Then
                    Throw New Exception(String.Format("公司別={0}，高階命令[ {1} ] 找不到低階命令[ {2} ]",
                                                      rwDictionary.Item("CompCode".ToUpper),
                                                      rwHigh.Value("CodeNo"),
                                                      Str.Trim))


                Else
                    aryLow.Value.Clear()
                    For Each colName As DataColumn In tbSetLowCmd.Columns
                        aryLow.Value.Add(colName.ColumnName.ToUpper, rwLow.Value(0).Item(colName))
                    Next
                    result.Value.Add(GetLowCmd(rwDictionary, aryLow.Value, 1))
                End If
            Next
            If (Not DBNull.Value.Equals(rwHigh.Value(0).Item("RunLowCmd2"))) AndAlso
                (Not String.IsNullOrEmpty(rwHigh.Value(0).Item("RunLowCmd2"))) Then
                For Each Str As String In rwHigh.Value(0).Item("RunLowCmd2").ToString.Split(",")
                    rwLow.Value = tbSetLowCmd.Select(String.Format("CodeNo='{0}'", Str.Trim))
                    If (rwLow.Value Is Nothing) OrElse (rwLow.Value.Length = 0) Then
                        Throw New Exception(String.Format("公司別={0}，高階命令[ {1} ] 找不到低階命令[ {2} ]",
                                                          rwDictionary.Item("CompCode".ToUpper),
                                                          rwHigh.Value("CodeNo"),
                                                          Str.Trim))


                    Else
                        aryLow.Value.Clear()
                        For Each colName As DataColumn In tbSetLowCmd.Columns
                            aryLow.Value.Add(colName.ColumnName.ToUpper, rwLow.Value(0).Item(colName))
                        Next
                        result.Value.Add(GetLowCmd(rwDictionary, aryLow.Value, 2))
                    End If
                Next
            End If

            Return result.Value
        Catch ex As Exception
            Throw
        Finally
            
            rwHigh.Dispose()
            rwLow.Dispose()

            result.Dispose()
            'aryLow.Value.Clear()
            aryLow.Dispose()
        End Try

    End Function
    Private Function GetLowCmd(ByVal rwHigh As Dictionary(Of String, Object),
                               ByVal rwSetLow As Dictionary(Of String, Object),
                               ByVal SetingCount As Int32) As SendLowCmd
        Dim result As New ThreadLocal(Of SendLowCmd)
        Dim ICC_NO As New ThreadLocal(Of String)
        Dim STB_NO As New ThreadLocal(Of String)
        result.Value = Nothing
       
        If Not DBNull.Value.Equals(rwHigh.Item("ICC_NO".ToUpper)) Then
            ICC_NO.Value = rwHigh.Item("ICC_NO".ToUpper)
        End If
        If Not DBNull.Value.Equals(rwHigh.Item("STB_NO".ToUpper)) Then
            STB_NO.Value = rwHigh.Item("STB_NO".ToUpper)
        End If
        Select Case rwHigh.Item("HIGH_LEVEL_CMD_ID".ToUpper)
            Case "A6", "A8", "A9"
                '判斷要抓取舊卡還是新卡
                Select Case SetingCount
                    Case 1                        
                        ICC_NO.Value = rwHigh.Item("NOTES".ToUpper).ToString.Split(";").ToList(0).Split("@").ToList(0).Split("/").ToList(1)
                        STB_NO.Value = rwHigh.Item("NOTES".ToUpper).ToString.Split(";").ToList(0).Split("@").ToList(0).Split("/").ToList(0)
                    Case 2                        
                        ICC_NO.Value = rwHigh.Item("NOTES".ToUpper).ToString.Split(";").ToList(0).Split("@").ToList(1).Split("/").ToList(1)
                        STB_NO.Value = rwHigh.Item("NOTES".ToUpper).ToString.Split(";").ToList(0).Split("@").ToList(1).Split("/").ToList(0)
                End Select
        End Select
        Try
            Select Case Convert.ToInt32(rwSetLow.Item("MsgId".ToUpper), 16)
                Case 1
                    result.Value = LowerCmd.Builder_SMS_CA_CREATE_SESSION_REQUEST()
                Case &H201
                    result.Value = LowerCmd.Builder_SMS_CA_OPEN_ACCOUNT_REQUEST(ICC_NO.Value)
                Case &H206
                    result.Value = LowerCmd.Builder_SMS_CA_REPAIR_REQUEST(
                                                ICC_NO.Value,
                                                STB_NO.Value)
                Case &H208
                    If (DBNull.Value.Equals(rwHigh.Item("Notes".ToUpper))) OrElse
                        (rwHigh.Item("Notes".ToUpper).ToString.Split(",").Count <> 2) Then
                        Throw New Exception("Notes 參數有誤!")
                    End If
                    result.Value = LowerCmd.Builder_SMS_CA_SETSLOTMONEY(ICC_NO.Value,
                                                                        rwHigh.Item("Notes".ToUpper).ToString.Split(",")(0),
                                                                          rwHigh.Item("Notes".ToUpper).ToString.Split(",")(1))
                Case &H20D
                    result.Value = LowerCmd.Builder_SMS_CA_ENTITLE_REQUEST(ICC_NO.Value, rwHigh)
                Case &H20E
                    result.Value = LowerCmd.Builder_SMS_CA_ENTITLE_REQUEST(&H20E, ICC_NO.Value, rwHigh)
                Case &H401
                    result.Value = LowerCmd.Builder_SMS_CA_CARD_REFRESH_REQUEST(ICC_NO.Value)
                Case &H202
                    result.Value = LowerCmd.Builder_SMS_CA_STOP_ACCOUNT_REQUEST(
                        ICC_NO.Value)
                Case &H203
                    result.Value = LowerCmd.Builder_SMS_CA_SET_LOCK_REQUEST(ICC_NO.Value)
                Case &H204
                    result.Value = LowerCmd.Builder_SMS_CA_SET_UNLOCK_REQUEST(ICC_NO.Value)
                Case &H402
                    '母卡抓取MIS_IRD_CMD_DATA Split(~)最後一個 (多少天驗證一次 ~ 彈性天數 ~ 多少時間(小時))
                    result.Value = LowerCmd.Builder_SMS_CA_SET_CHILD_REQUEST(ICC_NO.Value,
                                                                             rwHigh.Item("NOTES".ToUpper),
                                                                             rwHigh.Item("MIS_IRD_CMD_DATA".ToUpper). _
                                                                             ToString.Split("~"). _
                                                                             ToArray.Reverse.ToArray.Take(1).ToList(0)
                                                                             )
                Case &H403
                    result.Value = LowerCmd.Builder_SMS_CA_CANCEL_CHILD_REQUEST(ICC_NO.Value)
                Case &H207
                    result.Value = LowerCmd.Builder_SMS_CA_RESETCARDPIN_REQUEST(ICC_NO.Value)

                Case &H20C
                    If DBNull.Value.Equals(rwSetLow.Item("RealField".ToUpper)) OrElse (String.IsNullOrEmpty(rwSetLow.Item("RealField".ToUpper))) Then
                        Throw New Exception(String.Format("公司別:{0}", rwHigh.Item("CompCode".ToUpper)) & Environment.NewLine & _
                                String.Format("SEQNO={0}", rwHigh.Item("SEQNO".ToUpper)) & Environment.NewLine & _
                        String.Format("此低階命令{0}設定值有誤", rwSetLow.Item("CodeNo")))
                    End If
                    result.Value = LowerCmd.Builder_SMS_CA_SET_CHARACTER_REQUEST(ICC_NO.Value,
                                                                                 IIf(DBNull.Value.Equals(rwSetLow.Item("DefaultValue".ToUpper)),
                                                                                     "0", rwSetLow.Item("DefaultValue".ToUpper)),
                                                                                     rwHigh.Item(rwSetLow.Item("RealField".ToUpper).ToString.ToUpper))
                Case &H20F
                    'result.Value = LowerCmd.Builder_SMS_CA_ADV_CONTROL_PREVIEW_REQUEST(ICC_NO.Value,
                    '                                                     200, 100, 60, 1000)
                    result.Value = LowerCmd.Builder_SMS_CA_ADV_CONTROL_PREVIEW_REQUEST(ICC_NO.Value,
                                                                        rwHigh)
                Case &H301
                    result.Value = LowerCmd.Builder_SMS_CA_SEND_EMAIL_REQUEST("card=" & ICC_NO.Value,
                                                                              tbSetNSTV.Rows(0).Item("EmailTitle"),
                                                               rwHigh.Item("Notes".ToUpper),
                                                               Integer.Parse("0" & rwHigh.Item("MIS_IRD_CMD_DATA")))
                Case &H304
                    'result.Value = LowerCmd.Builder_SMS_CA_SEND_OSD_REQUEST("card=" & ICC_NO.Value,
                    '                                          rwHigh.Item("Notes".ToUpper),
                    '                                           1, 900)

                    result.Value = LowerCmd.Builder_SMS_CA_SEND_OSD_REQUEST("card=" & ICC_NO.Value,
                                                              rwHigh.Item("Notes".ToUpper))

                Case &H306
                    result.Value = LowerCmd.Builder_SMS_CA_LOCK_SERVICE_REQUEST("card=" & ICC_NO.Value, rwHigh)
                Case &H307
                    result.Value = LowerCmd.Builder_SMS_CA_UNLOCK_SERVICE_REQUEST("card=" & ICC_NO.Value)
                Case &H308
                    'result.Value = LowerCmd.Builder_SMS_CA_SEND_SUPEROSD_REQUEST("card=" & ICC_NO.Value,
                    '                                                             rwHigh.Item("Notes".ToUpper), 1, 900, 1,
                    '                                                             0, 255, 255, 50)
                    result.Value = LowerCmd.Builder_SMS_CA_SEND_SUPEROSD_REQUEST("card=" & ICC_NO.Value,
                                                                                rwHigh)
                Case Else
                    Throw New Exception(String.Format("公司別:{0}", rwHigh.Item("CompCode".ToUpper)) & Environment.NewLine & _
                                String.Format("SEQNO={0}", rwHigh.Item("SEQNO".ToUpper)) & Environment.NewLine & _
                        String.Format("此低階命令{0}尚未完成!", rwSetLow.Item("MsgId".ToUpper)))

            End Select
            Return result.Value
        Catch ex As Exception
            Throw New Exception(ex.ToString & _
                                String.Format("公司別:{0}", rwHigh.Item("CompCode".ToUpper)) & Environment.NewLine & _
                                String.Format("SEQNO={0}", rwHigh.Item("SEQNO".ToUpper)) & Environment.NewLine & _
                                Environment.NewLine & "MsgId=" & rwSetLow.Item("MsgId".ToUpper))
        Finally        
            result.Dispose()
            ICC_NO.Dispose()
            STB_NO.Dispose()
            'rwSource.Dispose()


        End Try

    End Function
#Region "IDisposable Support"
    Private disposedValue As Boolean ' 偵測多餘的呼叫

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                If tbSetHighCmd IsNot Nothing Then
                    tbSetHighCmd.Dispose()
                    tbSetHighCmd = Nothing
                End If
                If tbSetLowCmd IsNot Nothing Then
                    tbSetLowCmd.Dispose()
                    tbSetLowCmd = Nothing
                End If
                If tbSetNSTV IsNot Nothing Then
                    tbSetNSTV.Dispose()
                    tbSetNSTV = Nothing
                End If
                If LowerCmd IsNot Nothing Then
                    LowerCmd.Dispose()
                    LowerCmd = Nothing
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
