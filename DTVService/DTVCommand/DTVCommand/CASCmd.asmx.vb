Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel
Imports System.Text
Imports System.IO
Imports System.Data
Imports System.Xml
Imports System.Data.OracleClient
' 若要允許使用 ASP.NET AJAX 從指令碼呼叫此 Web 服務，請取消註解下一行。
' <System.Web.Script.Services.ScriptService()> _
<System.Web.Services.WebService(Namespace:="http://www.cablesoft.com.tw/")> _
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<ToolboxItem(False)> _
Public Class CASCmd
    Inherits System.Web.Services.WebService
    Private strIPAddress As String = String.Empty           '呼叫端IP
    Private strPassword As String = String.Empty            '密碼
    Private xmlSource As XmlDocument = Nothing
    Private strNowTime As DateTime                          '呼叫Fun的時間
    Private strCnString As String = String.Empty            'ConnectString
    Private oracleCn As OracleConnection = Nothing          'Connect
    Private OwnerName As String = String.Empty              'OwnerName
    Private strXMLSource As String = String.Empty           '來源XML
    Private strXMLOut As String = String.Empty              '回傳的XML
    Private strComp As String = String.Empty                '公司別
    Private strCMD As String = String.Empty                 '命令
    Private strCmdStatus As String = String.Empty            '命令狀態
    Private strProcessType As String = String.Empty          '程序別
    Private strICCNo As String = String.Empty                  'ICCNo
    Private strRetValue As String = String.Empty
    Private strRetCode As String = String.Empty
    Private strRetMsg As String = String.Empty
    Private aryCmd As ArrayList = Nothing
    Private strAccountID As String = String.Empty
    Private strSTBNo As String = String.Empty
    Private strCmdString As String = String.Empty
    Private intMaxCmd As Int32 = 0
    Private intHoldTime As Int32 = 0
    Private intNagraSEQ As Int32 = 0
    Private tra As OracleTransaction = Nothing
    Private str_CAS007SEQ As String = Nothing
    Private Enum RetXMLCoode
        Complete = 0        '指令傳送結果完成
        TimeOut = -1        '作業逾時
        VerifyErr = -2      '認證錯誤
        CommandErr = -3     '不合法的指令
        IDErr = -4          '不合法的授權ID
        CmdOrFmtErr = -5    '命令錯誤或格式有誤
        IPErr = -6          '不合法的IP
        NagraRet = -7       'Nagra回應的訊息
        No_ICCNo = -8       'ICC序號不存在
        No_STBNo = -9       'STB序號不存在
        No_HD = -10         '硬碟序號不存在
        Over_Cmd = -11      '超過可用命令數
        No_Dvr = -12        '無設定DVR頻道
        Other = -99          '其它錯誤
    End Enum
    Private Enum BlackIP
        UpdateTime = 0
        UpdateErr = 1
        AddData = 2
        NotRet = 3
    End Enum
    ' 0,@指令傳送結果完成
    '-1,J01@作業逾時
    '-2,A01@認證錯誤
    '-3,A02@不合法的指令
    '-4,A03@不合法的授權ID
    '-5,F01@命令錯誤或格式有誤
    '-6,C01@不合法的IP
    '-7,@

    
    <WebMethod()> _
    Public Function DTVCMD(ByVal strXML As String) As String

        strNowTime = Now        '記錄呼叫Function的時間
        strXMLSource = strXML   '記錄來源完整的XML
        strIPAddress = FormatIP(GetIP())   '取得呼叫端IP
        AddCmd()                '將命令定義至ArrayList
        strCmdStatus = "P"

        'Dim s As String = System.Web.Configuration.WebConfigurationManager.AppSettings("SODB")
        Try
            If Not OpenConnect() Then
                WriteErrToTxt(strXML, "資料庫連線建立不成功")
                strXMLOut = RetXML(RetXMLCoode.Other, "連線建立不成功")
                Return strXMLOut
            Else
                If oracleCn.State = ConnectionState.Closed Then
                    oracleCn.Open()
                End If
                tra = oracleCn.BeginTransaction
                strCmdStatus = "E"
                '檢查是否符合XML格式與xml必要的參數
                If Not ChkXMLFmt(strXML) OrElse Not ChkNode(xmlSource) Then
                    strXMLOut = RetXML(RetXMLCoode.CmdOrFmtErr)
                    If Not WriteLogToDB(strXMLSource) Then
                        tra.Rollback()
                    Else
                        tra.Commit()
                    End If
                    Return strXMLOut
                End If
                '如果是不合法IP直接回傳回去
                If Not ChkIP() Then
                    strXMLOut = RetXML(RetXMLCoode.IPErr)
                    If VerifyIP() = BlackIP.NotRet Then
                        tra.Commit()
                    Else
                        If Not WriteLogToDB(strXMLSource) Then
                            tra.Rollback()
                        Else
                            tra.Commit()
                        End If
                    End If
                    Return strXMLOut
                End If
                '檢查帳號密碼是否正確
                If Not ChkPassWord() Then
                    strXMLOut = RetXML(RetXMLCoode.VerifyErr)
                    If Not WriteLogToDB(strXMLSource) Then
                        tra.Rollback()
                    Else
                        tra.Commit()
                    End If
                    Return strXMLOut
                End If
                '取出限制命令數與TimeOut時間
                If Not GetMaxCmd() OrElse Not GetHoldTime() Then
                    strXMLOut = RetXML(RetXMLCoode.Other, "取得參數錯誤")
                    If Not WriteLogToDB(strXMLSource) Then
                        tra.Rollback()
                    Else
                        tra.Commit()
                    End If
                    Return strXMLOut
                End If
            End If

            If (tra IsNot Nothing) AndAlso (tra.Connection IsNot Nothing) Then
                tra.Commit()
            End If

            Try
                strCmdStatus = "P"
                WriteLogToDB(strXMLSource)
            Catch ex As Exception
                strXMLOut = RetXML(RetXMLCoode.Other, ex.Message)
                WriteErrToTxt(strXMLSource, ex.Message)
                Return False
            End Try

            Dim aHoldRet As String = String.Empty
            If TryCmd(aHoldRet, 90) Then
                If ExeNagra() Then

                End If
            Else
                If String.IsNullOrEmpty(aHoldRet) Then
                    strXMLOut = RetXML(RetXMLCoode.Other, aHoldRet)
                    WriteLogToDB(strXMLSource, False)
                    Return strXMLOut
                Else
                    If aHoldRet = "1" Then
                        strXMLOut = RetXML(RetXMLCoode.Over_Cmd)
                        WriteLogToDB(strXMLSource, False)
                    End If
                End If
            End If

            Try
                If oracleCn IsNot Nothing And oracleCn.State = ConnectionState.Open Then
                    oracleCn.Close()
                End If
            Finally
                oracleCn.Dispose()
                oracleCn = Nothing
            End Try
        Catch ex As Exception
            WriteErrToTxt(strXMLSource, "呼叫命令失敗! 原因 :" & ex.Message)
        Finally
            If oracleCn IsNot Nothing AndAlso oracleCn.State = ConnectionState.Open Then
                oracleCn.Close()
            End If
        End Try
        Return strXMLOut
    End Function

    '檢查XML的格式正不正確
    Private Function ChkXMLFmt(ByVal strXML As String) As Boolean
        
        Try
            xmlSource = New XmlDocument()
            xmlSource.LoadXml(strXML)
            'ChkNode(xmlSource)
        Catch ex As Exception
            Return False
        End Try

        Return True
    End Function
    '檢查所有必要的值,並檢查指令是否存在
    Private Function ChkMustValue(ByVal cmd As String) As Boolean
        If Not aryCmd.Contains(cmd) Then
            Return False
        End If
        '檢查第2個參數
        If cmd <> "B7" OrElse cmd <> "C1" Then
            If String.IsNullOrEmpty(strSTBNo) Then
                Return False
            End If
        End If
        '檢查第3個參數
        Select Case cmd
            Case "A2,A3,A4"
            Case "B1", "B2", "B3", "B4", "B5", "B6", "B7"
            Case "C1", "C4"
            Case Else
                If String.IsNullOrEmpty(strICCNo) Then
                    Return False
                End If
        End Select

        '檢查第4個參數
        Select Case cmd
            Case "A2,A3,A4,A7"
            Case "B3"
            Case "C4"
            Case "E9", "E15", "E16"
            Case Else
                If String.IsNullOrEmpty(strCmdString) Then
                    Return False
                End If
        End Select


        Return True

    End Function
    '將所有的命令存入陣列
    Private Sub AddCmd()
        aryCmd = New ArrayList()
        For i As Int32 = 1 To 9
            If i <> 5 Then
                aryCmd.Add("A" & i)
            End If
        Next
        For i As Int32 = 1 To 10
            If i <> 8 Then
                aryCmd.Add("B" & i)
            End If
        Next
        aryCmd.Add("C1")
        aryCmd.Add("C4")
        aryCmd.Add("E1")
        aryCmd.Add("E2-M")
        aryCmd.Add("E2-P")
        aryCmd.Add("E3")
        aryCmd.Add("E9")
        For i As Int32 = 15 To 18
            aryCmd.Add("E" & i)
        Next
    End Sub
    '檢核XML的Node是否有符合格式
    Private Function ChkNode(ByVal xmlDoc As XmlDocument) As Boolean
        Dim axmlNodeValue As XmlNode = Nothing
        If xmlDoc Is Nothing Then
            Return False
        End If
        If xmlDoc.SelectSingleNode("DataTable") Is Nothing Then
            Return False
        End If
        '公司別
        If xmlDoc.SelectSingleNode("DataTable/CompanyCode") Is Nothing Then

            Return False
        End If
        If xmlDoc.SelectSingleNode("DataTable/AccountID") Is Nothing Then
            Return False
        End If
        If xmlDoc.SelectSingleNode("DataTable/Password") Is Nothing Then
            Return False
        End If
        If xmlDoc.SelectSingleNode("DataTable/ProcessType") Is Nothing Then
            Return False
        End If

        'If xmlDoc.SelectSingleNode("DataTable/ICC") Is Nothing Then
        '    Return False
        'End If
        'If xmlDoc.SelectSingleNode("DataTable/STBNo") Is Nothing Then
        '    Return False
        'End If
        'If xmlDoc.SelectSingleNode("DataTable/CMDString") Is Nothing Then
        '    Return False
        'End If
        axmlNodeValue = xmlDoc.SelectSingleNode("DataTable/CompanyCode")
        strComp = axmlNodeValue.InnerText
        axmlNodeValue = xmlDoc.SelectSingleNode("DataTable/AccountID")
        strAccountID = axmlNodeValue.InnerText
        axmlNodeValue = xmlDoc.SelectSingleNode("DataTable/Password")
        strPassword = axmlNodeValue.InnerText
        axmlNodeValue = xmlDoc.SelectSingleNode("DataTable/ProcessType")
        strProcessType = axmlNodeValue.InnerText
        If String.IsNullOrEmpty(strComp) OrElse String.IsNullOrEmpty(strAccountID) _
            OrElse String.IsNullOrEmpty(strPassword) Then
            Return False
        End If
        axmlNodeValue = xmlDoc.SelectSingleNode("DataTable/ICC")
        If axmlNodeValue IsNot Nothing Then
            If Not String.IsNullOrEmpty(axmlNodeValue.InnerText) Then
                strICCNo = axmlNodeValue.InnerText
            End If
        End If

        axmlNodeValue = xmlDoc.SelectSingleNode("DataTable/STBNo")
        If axmlNodeValue IsNot Nothing Then
            If Not String.IsNullOrEmpty(axmlNodeValue.InnerText) Then
                strSTBNo = axmlNodeValue.InnerText
            End If
        End If

        axmlNodeValue = xmlDoc.SelectSingleNode("DataTable/CMDString")
        If axmlNodeValue IsNot Nothing Then
            If Not String.IsNullOrEmpty(axmlNodeValue.InnerText) Then
                strCmdString = axmlNodeValue.InnerText
            End If
        End If
        '檢查必要的參數
        If Not ChkMustValue(strProcessType) Then
            Return False
        End If
        Return True
    End Function
    Private Function GetIP() As String
        Return Context.Request.ServerVariables("REMOTE_HOST")
    End Function
    Private Function VerifyIP() As BlackIP
        Dim aSQL As String = String.Empty
        Dim aBLMinte As Int32 = 0
        Dim aBLCount As Int32 = 0
        Dim aTime As Date = Nothing
        Dim aErrCnd As Int32 = 0

        Try
            If oracleCn.State = ConnectionState.Closed Then
                oracleCn.Open()
            End If

            Dim cmd As OracleCommand = oracleCn.CreateCommand()
            cmd.Transaction = tra
            '取出參數幾分鐘內與錯誤次數
            aSQL = "SELECT Nvl(BLMinute,0) BLMinute,Nvl(BLCount,0) BLCount From  " & OwnerName & "CAS006"
            cmd.CommandText = aSQL
            Dim rd As OracleDataReader = cmd.ExecuteReader()
            If rd.Read() Then
                aBLMinte = rd.Item("BLMinute")
                aBLCount = rd.Item("BLCount")
            End If
            rd.Close()

            '取得目前的黑名單
            aSQL = "SELECT Nvl(ErrCount,0),CreateTime FROM " & OwnerName & "CAS010A " & _
                    " WHERE IPAddress='" & strIPAddress & "'"

            cmd.CommandText = aSQL

            rd = cmd.ExecuteReader()

            '有找到黑名單
            If rd.Read() Then
                aErrCnd = rd.Item(0)
                aTime = rd.Item(1)
                '判斷設定的分鐘數裡錯誤次數是否已達設定的錯誤次數
                If Math.Abs(DateDiff(DateInterval.Second, aTime, strNowTime)) <= (aBLMinte * 60) Then
                    '錯誤次還在範圍內,更新錯誤次數
                    If (aErrCnd > 0) AndAlso (aErrCnd + 1 < aBLCount) Then
                        UpdBlackIP(BlackIP.UpdateErr, False)
                        Return BlackIP.UpdateErr
                    Else
                        '錯誤次數已經大於最大次數,直接回傳回去
                        If aErrCnd + 1 > aBLCount Then
                            Return BlackIP.NotRet
                        Else
                            If aErrCnd + 1 = aBLCount Then
                                UpdBlackIP(BlackIP.UpdateErr, True)
                                Return BlackIP.NotRet
                            Else
                                '錯誤次數沒到達最大,更新錯誤時間
                                UpdBlackIP(BlackIP.UpdateTime, False)
                                Return BlackIP.UpdateTime
                            End If
                        End If

                    End If
                Else
                    UpdBlackIP(BlackIP.UpdateTime, False)
                    Return BlackIP.UpdateTime
                End If
            Else
                UpdBlackIP(BlackIP.AddData, False)
                Return BlackIP.AddData
            End If
            rd.Close()

        Catch ex As Exception
            WriteErrToTxt(strXMLSource, " Function : VerifyIP 原因 : " & ex.ToString())
            Return False

        End Try
    End Function
    Private Function UpdBlackIP(ByVal aUpd As BlackIP, ByVal Addrcd As Boolean) As Boolean
        Dim aSQL As String = String.Empty
        Dim aSQLAdd As String = String.Empty

        If aUpd = BlackIP.ADDDATA Then
            aSQL = "Insert Into " & OwnerName & "CAS010A (" & _
            "IPAddress,CreateTime,ErrCount) Values(" & _
            "'{0}', TO_DATE('{1}','yyyymmddhh24miss'),{2})"
            aSQL = String.Format(aSQL, strIPAddress, _
                        Format(strNowTime, "yyyyMMddHHmmss"), "1")
        Else
            If aUpd = BlackIP.UPDATEERR Then
                aSQL = "UPDATE " & OwnerName & "CAS010A SET Errcount=ErrCount+1" & _
                    " Where IPAddress='" & strIPAddress & "'"
            Else
                aSQL = "UPDATE " & OwnerName & "CAS010A SET CreateTime=" & _
                    "TO_DATE('{0}','yyyymmddhh24miss') " & _
                    " Where IPAddress='" & strIPAddress & "'"
                aSQL = String.Format(aSQL, Format(strNowTime, "yyyyMMddHHmmss"))
            End If
            
            aSQLAdd = "Insert Into " & OwnerName & "CAS010 (" & _
                "IPAddress,CreateTime) Values(" & _
                "'{0}',TO_DATE('{1}','yyyymmddhh24miss'))"

            aSQLAdd = String.Format(aSQLAdd, strIPAddress, Format(strNowTime, "yyyyMMddHHmmss"))

            

        End If

        Try
            If oracleCn.State = ConnectionState.Closed Then
                oracleCn.Open()
            End If
            Dim cmd As OracleCommand = oracleCn.CreateCommand()
            cmd.Transaction = tra
            cmd.CommandText = aSQL
            cmd.ExecuteNonQuery()
            If Addrcd Then
                Try
                    If oracleCn.State = ConnectionState.Closed Then
                        oracleCn.Open()
                    End If

                    cmd.CommandText = "Select Count(*) From " & OwnerName & "CAS010 " & _
                            " WHERE IPAddress='" & strIPAddress & "'"
                    If cmd.ExecuteScalar = 0 Then
                        cmd.CommandText = aSQLAdd
                        cmd.ExecuteNonQuery()
                    End If

                Catch ex As Exception
                    WriteErrToTxt(strXMLSource, "Function : UpdBlackIP(Addrcd)  原因 :" & ex.Message)
                    'tra.Rollback()
                    oracleCn.Close()
                    Return False
                End Try
            
            End If
            'tra.Commit()
        Catch ex As Exception
            'tra.Rollback()
            WriteErrToTxt(strXMLSource, "Function : UpdBlackIP(Addrcd)  原因 :" & ex.Message)
        
        End Try
        Return True
    End Function
    ''' <summary>
    ''' 回應的XML
    ''' </summary>
    ''' <param name="aFmt">RetValue</param>
    ''' <param name="aMsg">其它的錯誤訊息,非必要值</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function RetXML(ByVal aFmt As RetXMLCoode, _
                            Optional ByVal aMsg As String = Nothing) As String
        Dim aBdr As New StringBuilder()
        Dim aXMLString As String = String.Empty
        
        Select Case aFmt
            Case RetXMLCoode.Complete
                strRetValue = RetXMLCoode.Complete
                strRetMsg = "指令傳送結果完成"
            Case RetXMLCoode.TimeOut
                strRetValue = RetXMLCoode.TimeOut
                strRetCode = "J01"
                strRetMsg = "作業逾時"
            Case RetXMLCoode.VerifyErr
                strRetValue = RetXMLCoode.VerifyErr
                strRetCode = "A01"
                strRetMsg = "認證錯誤"
            Case RetXMLCoode.CommandErr
                strRetValue = RetXMLCoode.CommandErr
                strRetValue = "A02"
                strRetMsg = "不合法的指令"
            Case RetXMLCoode.IDErr
                strRetValue = RetXMLCoode.IDErr
                strRetCode = "A03"
                strRetMsg = "不合法的授權ID"
            Case RetXMLCoode.CmdOrFmtErr
                strRetValue = RetXMLCoode.CmdOrFmtErr
                strRetCode = "F01"
                strRetMsg = "命令錯誤或格式有誤"
            Case RetXMLCoode.IPErr
                strRetValue = RetXMLCoode.IPErr
                strRetCode = "C01"
                strRetMsg = "不合法的IP"
            Case RetXMLCoode.NagraRet
                strRetValue = RetXMLCoode.NagraRet
                'Nagra回應的訊息
            Case RetXMLCoode.No_ICCNo
                strRetValue = RetXMLCoode.No_ICCNo
                strRetCode = "A04"
                strRetMsg = "ICC序號不存在"
            Case RetXMLCoode.No_STBNo
                strRetValue = RetXMLCoode.No_STBNo
                strRetCode = "A05"
                strRetMsg = "STB序號不存在"
            Case RetXMLCoode.No_HD
                strRetValue = RetXMLCoode.No_HD
                strRetCode = "A06"
                strRetMsg = "硬碟序號不存在"
            Case RetXMLCoode.Over_Cmd
                strRetValue = RetXMLCoode.Over_Cmd
                strRetCode = "A07"
                strRetMsg = "超過可用命令數"
            Case RetXMLCoode.No_Dvr
                strRetValue = RetXMLCoode.No_Dvr
                strRetCode = "A08"
                strRetMsg = "無設定DVR頻道"

            Case Else
                If String.IsNullOrEmpty(aMsg) Then
                    aMsg = "其它錯誤"
                End If
                strRetValue = "-99"
                strRetCode = "XXX"
                strRetMsg = aMsg
        End Select

        aBdr.AppendLine("<?xml version=""1.0"" encoding=""UTF-8""?>")
        aBdr.AppendLine("<DataTable>")
        aBdr.AppendLine(String.Format("<ResultValue>{0}</ResultValue>", strRetValue))
        aBdr.AppendLine(String.Format("<ErrorCode>{0}</ErrorCode>", strRetCode))
        aBdr.AppendLine(String.Format("<ErrorString>{0}</ErrorString>", strRetMsg))
        aBdr.AppendLine("</DataTable>")
        Return aBdr.ToString()



    End Function
    Private Function WriteLogToDB(ByVal strSourceXML As String, Optional ByVal AddData As Boolean = True) As Boolean
        Dim aSQL As String = String.Empty
        Dim aSeq As String = String.Empty
        Try
            
            Try
                If AddData Then
                    aSeq = GetSeq("S_CAS007_SeqNo")
                    str_CAS007SEQ = aSeq
                    If String.IsNullOrEmpty(aSeq) Then
                        Return False
                    End If
                End If
                
            Catch ex As Exception
                WriteErrToTxt(strSourceXML, "原因:" & ex.Message.ToString() & _
                        " Function: GetSeq")
                Return False
            End Try

            If AddData Then
                aSQL = "INSERT INTO " & OwnerName & "CAS007 ( " & _
                    "SEQNO,CMD_STATUS,XMLINTIME,XMLIN,COMPCODE,ProcessType," & _
                    "ICCNo,XMLOutTime,XMLOut,RetCode,RetMsg ) VALUES (" & _
                    "{0},'{1}',TO_DATE('{2}','yyyymmddhh24miss')," & _
                    "'{3}',{4},'{5}','{6}',TO_DATE('{7}','yyyymmddhh24miss' )," & _
                    "'{8}','{9}','{10}')"

                aSQL = String.Format(aSQL, aSeq, strCmdStatus, _
                                     Format(strNowTime, "yyyyMMddHHmmss"), strXMLSource, _
                                     IIf(String.IsNullOrEmpty(strComp), "Null", strComp), _
                                     strProcessType, strICCNo, _
                                     Format(Now, "yyyyMMddHHmmss"), strXMLOut, _
                                     strRetCode, strRetMsg)
            Else
                aSQL = "UPDATE " & OwnerName & "CAS007 SET " & _
                    " Cmd_STATUS='{0}',XMLOut='{1}'," & _
                    "XMLOutTime=To_DATE('{2}','yyyymmddhh24miss'),RetCode='{3}'," & _
                    "RetMsg='{4}' WHERE SEQNO={5}"

                aSQL = String.Format(aSQL, strCmdStatus, strXMLOut, _
                                     Format(Now, "yyyyMMddhhmmss"), strRetCode, _
                                     strRetMsg, str_CAS007SEQ)
            End If
            
            If oracleCn.State = ConnectionState.Closed Then
                oracleCn.Open()
            End If
            Dim cmd As OracleCommand = oracleCn.CreateCommand
            If tra.Connection IsNot Nothing Then
                cmd.Transaction = tra
            End If

            cmd.CommandText = aSQL
            
            cmd.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            WriteErrToTxt(strSourceXML, "原因:" & ex.Message.ToString() & " SQL:" & aSQL & _
                          " Function : WriteLogToDB")
            Return False
        End Try

    End Function
    '寫入錯誤訊息
    Private Function WriteErrToTxt(ByVal strXML As String, ByVal strErrMsg As String) As Boolean
        Dim stm As StreamWriter = Nothing
        Dim aMsg As New StringBuilder
        aMsg.AppendLine("***************************************************")
        aMsg.AppendLine("傳入時間：" & strNowTime.ToString())
        aMsg.AppendLine("傳入資料：" & strXML)
        aMsg.AppendLine("錯誤原因：" & strErrMsg)
        aMsg.AppendLine("***************************************************")
        Try
            'If Not Directory.Exists(Server.MapPath(".\") & Format(Now, "CAS_yyyyMM")) Then
            '    Try
            '        Directory.CreateDirectory(Server.MapPath(".\") & Format(Now, "CAS_yyyyMM"))
            '    Catch
            '        strXMLOut = "建立錯誤"
            '        Return False
            '    End Try
            'End If
            Dim strFileName As String = Server.MapPath(".\") & Format(Now, "CAS_yyyyMM") & "\CASErrLog_" & Format(Now, "yyyyMMdd") & ".txt"
            strFileName = Server.MapPath(".\") & Format(Now, "yyyyMMdd") & ".txt"
            
            stm = File.AppendText(strFileName)
            stm.WriteLine(aMsg.ToString)
            stm.Flush()
            stm.Close()
            stm.Dispose()
            Return True
        Catch ex As Exception

            Return False
        Finally
            If stm IsNot Nothing Then
                stm.Close()
                stm.Dispose()
            End If
        End Try

    End Function
    Private Function GetSeq(ByVal seqName As String) As String
        Dim strRet As String = String.Empty
        Try
            Dim aSQL As String = String.Empty
            aSQL = "SELECT " & OwnerName & seqName & ".NEXTVAL FROM DUAL "
            If oracleCn.State = ConnectionState.Closed Then
                oracleCn.Open()
            End If
            Dim cmd As OracleCommand = oracleCn.CreateCommand()
            If tra.Connection IsNot Nothing Then
                cmd.Transaction = tra
            End If
            cmd.CommandText = aSQL
            strRet = cmd.ExecuteScalar
            Return strRet
        Catch ex As Exception
            WriteErrToTxt(strXMLSource, "Function : GetSeq 原因:" & ex.Message)
            Return String.Empty
        End Try
    End Function
    
    Private Overloads Function OpenConnect() As Boolean
        Try
            strCnString = System.Web.Configuration.WebConfigurationManager.AppSettings("SODB")
            strCnString = CryptUtil.Decrypt(strCnString)
            Try
                oracleCn = New OracleConnection(strCnString)
                oracleCn.Open()
            Catch ex As Exception
                Return False
            End Try
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function
    '開啟連接
    Private Overloads Function OpenConnect(ByVal strFile As String, Optional ByVal intComp As Int32 = 0) As Boolean
        Try
            Using stmRead As New StreamReader(strFile, Encoding.Default)
                Dim strConText As String = String.Empty
                Dim strKey As String = String.Empty
                Dim strCnn As String = String.Empty
                Dim aHtb As New Hashtable()
                Do While Not stmRead.EndOfStream
                    strConText = CryptUtil.Decrypt(stmRead.ReadLine)
                    strKey = strConText.Split(New String() {",@"}, StringSplitOptions.None).GetValue(0)
                    strCnn = strConText.Split(New String() {",@"}, StringSplitOptions.None).GetValue(1)
                    aHtb.Add(Convert.ToInt32(strKey), strCnn)
                Loop
                If intComp = 0 Then
                    If aHtb.Count > 0 Then
                        strCnString = aHtb.Values(0)
                    Else
                        strCnString = String.Empty
                    End If
                Else
                    If aHtb.Contains(intComp) Then
                        strCnString = aHtb.Item(intComp)
                    Else
                        strCnString = String.Empty
                    End If
                End If
                stmRead.Close()
            End Using
            If String.IsNullOrEmpty(strCnString) Then
                Return False
            End If
            Try
                oracleCn = New OracleConnection(strCnString)
                oracleCn.Open()
            Catch ex As Exception
                Return False
            End Try
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function
    Private Function ChkPassWord() As Boolean
        Dim aPassWord As String = String.Empty
        Dim aSQL As String = String.Empty
        Try
            aSQL = "SELECT COUNT(*) FROM " & OwnerName & "CAS002 " & _
                " WHERE UserId='" & CryptUtil.Encrypt(strAccountID) & "'" & _
                " AND Pws='" & CryptUtil.Encrypt(strPassword) & "'" & _
                " AND NVL(STOPFLAG,0)=0 AND CODENO=" & strComp
            If oracleCn.State = ConnectionState.Closed Then
                oracleCn.Open()
            End If
            Dim cmd As OracleCommand = oracleCn.CreateCommand
            cmd.Transaction = tra
            cmd.CommandText = aSQL

            Dim obj As Object = cmd.ExecuteScalar()
            If obj IsNot Nothing AndAlso obj > 0 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Return False
        
        End Try
    End Function
    Private Function ChkIP() As Boolean
        Dim aIPAddress As String = String.Empty
        Dim aSQL As String = String.Empty
        Try
            aSQL = "SELECT COUNT(*) FROM " & OwnerName & "CAS002 " & _
                " WHERE ClientIp='" & strIPAddress & "'" & _
                " AND NVL(STOPFLAG,0)=0 AND CODENO=" & strComp
            If oracleCn.State = ConnectionState.Closed Then
                oracleCn.Open()
            End If

            Dim cmd As OracleCommand = oracleCn.CreateCommand
            cmd.Transaction = tra
            cmd.CommandText = aSQL
            Dim r As OracleDataReader=cmd.ExecuteReader( CommandBehavior.CloseConnection)
            Dim obj As Object = cmd.ExecuteScalar()
            If obj IsNot Nothing AndAlso obj > 0 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Return False
        
        End Try
    End Function
    Private Function FormatIP(ByVal aIP As String) As String
        Try
            Dim aTmp As String = String.Empty
            Dim aAry() As String = aIP.Split(".")
            For i As Int32 = LBound(aAry) To UBound(aAry)
                If String.IsNullOrEmpty(aTmp) Then
                    aTmp = Right("000" & aAry(i), 3)
                Else
                    aTmp = aTmp & "." & Right("000" & aAry(i), 3)
                End If
            Next
            Return aTmp
        Catch ex As Exception
            Return ""
        End Try

    End Function
    
    Private Function TryCmd(ByRef aRetMsg As String, Optional ByVal aPercent As Int32 = 0) As Boolean
        Dim aSQL As String = String.Empty
        Dim aCmdCntSQL As String = String.Empty
        Dim aCmdCnt As Int32 = 0
        Dim aRet As Boolean = False
        Dim stw As New Stopwatch()

        aSQL = "SELECT COUNT(*) CNT FROM " & OwnerName & "SEND_NAGRA " & _
            " WHERE HIGH_LEVEL_CMD_ID='" & strProcessType & "'" & _
            " AND (CMD_STATUS='P' OR CMD_STATUS='W') " & _
            " AND RealCompCode = " & strComp

        aCmdCntSQL = "SELECT CmdCount FROM " & OwnerName & "CAS008 " & _
            " WHERE CODENO='" & strProcessType & "'"

        Try
            If oracleCn.State = ConnectionState.Closed Then
                oracleCn.Open()
            End If

            Dim cmd As OracleCommand = oracleCn.CreateCommand()
            If tra.Connection IsNot Nothing Then
                cmd.Transaction = tra
            End If
            cmd.CommandText = aCmdCntSQL
            aCmdCnt = cmd.ExecuteScalar '取出CAS008最大命令數
            cmd.CommandText = aSQL
            stw.Reset()
            stw.Start()

            'If intHoldTime > 0 Then
            '    If (stw.ElapsedMilliseconds / 1000) >= intHoldTime Then

            '    End If
            'End If
            If aPercent > 0 Then
                Dim a As Int32 = (cmd.ExecuteScalar + aCmdCnt) * (aPercent / 100)
                If ((cmd.ExecuteScalar + aCmdCnt) * (aPercent / 100)) <= (intMaxCmd * (aPercent / 100)) Then
                    aRet = True

                Else
                    aRetMsg = "1"
                    aRet = False
                End If
            Else
                If cmd.ExecuteScalar + aCmdCnt <= intMaxCmd Then
                    aRet = True

                Else
                    aRetMsg = "1"
                    aRet = False
                End If
            End If
            stw.Stop()
            Return aRet
        Catch ex As Exception
            aRetMsg = "Function : HoldCmd 原因:" & ex.ToString
            Return False
        Finally
            If stw.IsRunning Then
                stw.Stop()
            End If

        End Try
    End Function
    Private Function GetMaxCmd() As Boolean
        Dim aMaxSQL As String = String.Empty
        aMaxSQL = "SELECT CASMaxCount FROM " & OwnerName & "CAS002 " & _
            " WHERE CODENO=" & strComp & " AND NVL(STOPFLAG,0)=0 "
        Try
            If oracleCn.State = ConnectionState.Closed Then
                oracleCn.Open()
            End If
            Dim cmd As OracleCommand = oracleCn.CreateCommand()
            cmd.Transaction = tra
            cmd.CommandText = aMaxSQL
            intMaxCmd = cmd.ExecuteScalar()
            Return True
        Catch ex As Exception
            WriteErrToTxt(strXMLSource, "取出CASMaxCount錯誤 原因:" & ex.ToString)
            Return False
        End Try
    End Function
    Private Function GetHoldTime() As Boolean
        Dim aTimeSQL As String = String.Empty

        aTimeSQL = "SELECT AccTimeOut FROM " & OwnerName & "CAS006 " & _
            " WHERE SYSID=" & strComp
        Try
            If oracleCn.State = ConnectionState.Closed Then
                oracleCn.Open()
            End If
            Dim cmd As OracleCommand = oracleCn.CreateCommand()
            cmd.Transaction = tra
            cmd.CommandText = aTimeSQL
            intHoldTime = cmd.ExecuteScalar()
            Return True
        Catch ex As Exception
            WriteErrToTxt(strXMLSource, "取出AccTimeOut錯誤 原因:" & ex.ToString)
            Return False
        End Try
    End Function
    Private Function ExeNagra() As Boolean

        Dim objNagra As New clsNagra(strProcessType, oracleCn, tra)
        With objNagra
            .Company = strComp
            .ICCNo = strICCNo
            .STBNo = strSTBNo
            .CmdString = strCmdString
            .TimeOut = intHoldTime
            .SEQ = GetSeq("S_SendNagra_SeqNo")
            .ExeTime = strNowTime
            .ExeCmd()
            If .RetResult = NagraRet.Complete Then
                strXMLOut = RetXML(RetXMLCoode.Complete)
                If tra.Connection Is Nothing Then
                    tra = oracleCn.BeginTransaction
                    .Transcation = tra
                End If
                If .UpdData Then
                    If WriteLogToDB(strXMLSource, False) Then
                        tra.Commit()
                        Return True
                    Else
                        strXMLOut = RetXML(RetXMLCoode.Other, objNagra.NagraMsg)
                        WriteErrToTxt(strXMLSource, objNagra.NagraMsg)
                        tra.Rollback()
                        Return False
                    End If
                Else
                    tra.Rollback()
                    strXMLOut = RetXML(RetXMLCoode.Other, objNagra.NagraMsg)
                    WriteLogToDB(strXMLSource, False)
                    WriteErrToTxt(strXMLSource, objNagra.NagraMsg)
                    Return False
                End If

                Return True
            Else
                Select Case .RetResult
                    Case NagraRet.Nagra_RetErr
                        strRetCode = .NagraCode
                        strRetMsg = .NagraMsg
                        strXMLOut = RetXML(RetXMLCoode.NagraRet)
                    Case NagraRet.No_ICCNo
                        strXMLOut = RetXML(RetXMLCoode.No_ICCNo)
                    Case NagraRet.NO_STBNo
                        strXMLOut = RetXML(RetXMLCoode.No_STBNo)
                    Case NagraRet.NO_HDNo
                        strXMLOut = RetXML(RetXMLCoode.No_HD)
                    Case NagraRet.TimeOut
                        strXMLOut = RetXML(RetXMLCoode.TimeOut)
                    Case NagraRet.NO_Dvr
                        strXMLOut = RetXML(RetXMLCoode.No_Dvr)
                    Case NagraRet.SysError
                        strXMLOut = RetXML(RetXMLCoode.Other, .NagraMsg)
                        WriteErrToTxt(strXMLSource, .NagraMsg)

                End Select
                If tra.Connection Is Nothing Then
                    tra = oracleCn.BeginTransaction
                End If

                If WriteLogToDB(strXMLSource) Then
                    tra.Commit()
                Else
                    tra.Rollback()
                End If
                

                Return False
            End If
        End With

        
    End Function
End Class

