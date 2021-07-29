Imports System.IO
Imports System.Xml
Imports Oracle.DataAccess.Client
Public Class IVRCommand
    Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
    Private sIVRCMD As String = String.Empty
    Private sTel As String = String.Empty
    Private sOwnerName As String = String.Empty
    Private sFloor As String = String.Empty
    Private sIniFile As String = String.Empty
    Private sCnn As String = String.Empty
    Private iCompCode As Integer = 0
    Private sCustId As String = String.Empty
    Private sIVRFileCode As String = String.Empty
    Private iAddrNo As Integer
    Private Oracn As OracleConnection = New OracleConnection()
    Private strRunTime As String = String.Empty
    Private sNodeNo As String = String.Empty
    Private sTel1 As String = String.Empty
    Private sTel2 As String = String.Empty
    Private sTel3 As String = String.Empty
    Private sCustName As String = String.Empty
    Private sCircuitNo As String = String.Empty
    Private sbDebug As StringBuilder
    Private sIVRDeclarantName As String = String.Empty
    Private sIVRInstaddress As String = String.Empty
    Private sCMCode As String = String.Empty
    Private iCloseDate As Integer = 0
    Private iStatus As Integer = 0
    Private iOweTotal As Integer = 0
    Private iTotal As Integer = 0
    'Private sbDebug As StreamWriter

    'Private OraCn As New OracleConnection()
    ''' <summary>
    ''' 執行命令的建構式
    ''' </summary>
    ''' <param name="strTel">XML的Tel</param>
    ''' <param name="strFloor">XML的Floor</param>
    ''' <param name="intComp">XML的Company</param>
    ''' <param name="strOwner">OwnerName</param>
    ''' <param name="strIniFile">虛擬目錄的Ini檔案</param>
    ''' <param name="strCnn">連線字串</param>
    Sub New(ByVal strTel As String, _
            ByVal strFloor As String, _
            ByVal intComp As Integer, _
            ByVal strOwner As String, _
            ByVal strIniFile As String, _
            ByVal strCnn As String)
        strRunTime = Date.Now.ToString
        sTel = strTel
        sFloor = strFloor
        iCompCode = intComp
        sIniFile = strIniFile
        'sOwnerName = strOwner
        sCnn = strCnn
        Oracn.ConnectionString = sCnn
        Try
            Oracn.Open()
            GetOwner()
        Catch ex As Exception

        End Try

    End Sub
    Sub New(ByVal strTel As String, _
            ByVal strFloor As String, _
            ByVal intComp As Integer, _
            ByVal strOwner As String, _
            ByVal strIniFile As String, _
            ByVal strCnn As String, _
            ByRef sb As StringBuilder)
        sbDebug = sb
        strRunTime = Date.Now.ToString
        sTel = strTel
        sFloor = strFloor
        iCompCode = intComp
        sIniFile = strIniFile
        'sOwnerName = strOwner
        sCnn = strCnn
        Oracn.ConnectionString = sCnn
        Try
            Oracn.Open()
            GetOwner()
        Catch ex As Exception
            If Not sbDebug Is Nothing Then
                sbDebug.AppendLine("[New] Oracle開啟失敗" & Environment.NewLine)
            End If
        End Try

    End Sub
    Public Function N21() As Boolean

        If SearchSO0102("'C'") Then
            sIVRFileCode = SF_GETMFMSG(iAddrNo, iCompCode, "C")
            Return True
        Else
            Return False
        End If

    End Function
    Public Function N22() As Boolean

        If SearchSO004("'D'", "3") Then
            sIVRFileCode = SF_GETMFMSG(iAddrNo, iCompCode, "D")
            Return True
        Else
            Return False
        End If
    End Function
    Public Function N23() As Boolean

        If SearchSO004("'I','P'", "2,5") Then
            sIVRFileCode = SF_GETMFMSG(iAddrNo, iCompCode, "I")
            Return True
        Else
            Return False
        End If
    End Function
    Public Function N24() As Boolean

        If SearchSO0102("'C'") Then
            sIVRFileCode = SF_GETMFMSG(iAddrNo, iCompCode, "C")
            Return True
        Else
            Return False
        End If
    End Function
    Public Function N25() As Boolean

        If SearchSO004("'D'", "3") Then
            sIVRFileCode = SF_GETMFMSG(iAddrNo, iCompCode, "D")
            Return True
        Else
            Return False
        End If
    End Function

    Public Function N26() As Boolean

        If SearchSO004("'I','P'", "2,5") Then
            sIVRFileCode = SF_GETMFMSG(iAddrNo, iCompCode, "I")
            Return True
        Else
            Return False
        End If
    End Function
    Public Function N31() As Boolean
        If Not SearchSO0102() Then
            If Not SearchSO004() Then
                iStatus = 2
                Return False
            End If
        End If
        If Not GetSO033AMT("'C','D'") Then
            iStatus = 2
            Return False
        End If
        If iTotal = 0 AndAlso iOweTotal = 0 Then
            iStatus = 0
        Else
            If iTotal > 0 Then
                iStatus = 1
            Else
                iStatus = -1
            End If
        End If
        Return True
    End Function
    Public Function N32() As Boolean
        If Not SearchSO0102() Then
            If Not SearchSO004() Then
                iStatus = 2
                Return False
            End If
        End If
        If Not GetSO033AMT("'C','D'") Then
            iStatus = 2
            Return False
        End If
        If iTotal = 0 AndAlso iOweTotal = 0 Then
            iStatus = 0
        Else
            If iTotal > 0 Then
                iStatus = 1
            Else
                iStatus = -1
            End If
        End If
        Return True
    End Function
    Public Function N33() As Boolean
        If Not SearchSO0102() Then
            If Not SearchSO004() Then
                iStatus = 2
                Return False
            End If
        End If
        If Not GetSO033AMT("'I','P'") Then
            iStatus = 2
            Return False
        End If
        If iTotal = 0 AndAlso iOweTotal = 0 Then
            iStatus = 0
        Else
            If iTotal > 0 Then
                iStatus = 1
            Else
                iStatus = -1
            End If
        End If
        Return True
    End Function
    Public Function N34() As Boolean
        If Not SearchSO0102() Then
            If Not SearchSO004() Then
                iStatus = -1
                Return False
            End If
        End If
        If Not GetSO033AMT("'C','D'") Then
            iStatus = -1
            Return False
        End If
        If iTotal = 0 AndAlso iOweTotal = 0 Then
            iStatus = 0
        Else
            If iOweTotal > 0 Then
                iStatus = 2
            Else
                If iTotal > 0 Then
                    iStatus = 1
                End If
            End If
        End If
        Return True
    End Function
    Public Function N35() As Boolean
        If Not SearchSO0102() Then
            If Not SearchSO004() Then
                iStatus = -1
                Return False
            End If
        End If
        If Not GetSO033AMT("'C','D'") Then
            iStatus = -1
            Return False
        End If
        If iTotal = 0 AndAlso iOweTotal = 0 Then
            iStatus = 0
        Else
            If iOweTotal > 0 Then
                iStatus = 2
            Else
                If iTotal > 0 Then
                    iStatus = 1
                End If
            End If
        End If
        Return True
    End Function
    Public Function N36() As Boolean
        If Not SearchSO0102() Then
            If Not SearchSO004() Then
                iStatus = -1
                Return False
            End If
        End If
        If Not GetSO033AMT("'I','P'") Then
            iStatus = -1
            Return False
        End If
        If iTotal = 0 AndAlso iOweTotal = 0 Then
            iStatus = 0
        Else
            If iOweTotal > 0 Then
                iStatus = 2
            Else
                If iTotal > 0 Then
                    iStatus = 1
                End If
            End If
        End If
        Return True
    End Function
    Private Function SearchSO004() As Boolean
        Dim blnReturn As Boolean = False
        Dim strQry As String = "Select CustId,PRDate,DeclarantName From " & sOwnerName & "SO004" & _
                                " Where (ContTel='" & sTel & "' Or ContTel2='" & sTel & "' Or AgentTel='" & sTel & "')" & _
                                " And CompCode=" & iCompCode
        Try
            Using adapter As New OracleDataAdapter()
                adapter.SelectCommand = New OracleCommand(strQry, Oracn)
                Using builder As OracleCommandBuilder = New OracleCommandBuilder(adapter)
                    Using dst As New Data.DataSet()
                        adapter.Fill(dst, "SO004")
                        '找不到SO004的資料,則尋找SO001與SO002的資料,如果有找到則以客編反查SO004
                        If dst.Tables("SO004").Rows.Count <= 0 Then
                            Return False
                        Else
                            For i As Integer = 0 To dst.Tables("SO004").Rows.Count - 1
                                sCustId = String.Empty
                                sCustId = dst.Tables("SO004").Rows(i).Item("CustId") & ""
                                sIVRDeclarantName = String.Empty
                                sIVRDeclarantName = dst.Tables("SO004").Rows(i).Item("DeclarantName") & ""
                                If SearchSO0102() Then
                                    blnReturn = True
                                End If
                            Next i
                        End If
                    End Using
                End Using
            End Using
            If Not sbDebug Is Nothing Then
                sbDebug.AppendLine("[SearchSO004] " & IIf(blnReturn, "成功", "失敗") & " stQry:" & strQry & Environment.NewLine)
            End If
            Return blnReturn
        Catch ex As Exception
            If Not sbDebug Is Nothing Then
                sbDebug.AppendLine("[SearchSO004] 載入失敗" & Environment.NewLine)
            End If
            Return False
        End Try
    End Function
    ''' <summary>
    ''' 尋找SO004是否有符合的資料
    ''' </summary>
    ''' <param name="sServiceType">服務別(例'C' Or 'I','P')</param>
    ''' <param name="sRefNo">設備參考號(CD022.RefNo)</param>
    ''' <returns>布林值</returns>
    Private Function SearchSO004(ByVal sServiceType As String, ByVal sRefNo As String) As Boolean
        Dim blnReturn As Boolean = False
        Dim strQry As String = "Select CustId,PRDate,DeclarantName From " & sOwnerName & "SO004" & _
                                " Where (ContTel='" & sTel & "' Or ContTel2='" & sTel & "' Or AgentTel='" & sTel & "')" & _
                                " And CompCode=" & iCompCode & _
                                " And ServiceType In(" & sServiceType & ")" & _
                                " And FaciCode In ( Select CodeNo From CD022 Where RefNo In(" & sRefNo & "))"

        Try
            Using adapter As New OracleDataAdapter()
                adapter.SelectCommand = New OracleCommand(strQry, Oracn)
                Using builder As OracleCommandBuilder = New OracleCommandBuilder(adapter)
                    Using dst As New Data.DataSet()
                        adapter.Fill(dst, "SO004")
                        '找不到SO004的資料,則尋找SO001與SO002的資料,如果有找到則以客編反查SO004
                        If dst.Tables("SO004").Rows.Count <= 0 Then
                            If SearchSO0102(sServiceType) Then
                                If ReverSO004(sServiceType, sRefNo) Then
                                    If Not sbDebug Is Nothing Then
                                        sbDebug.AppendLine("[SearchSO004] 成功 SO004找不到,但以SearchSO0102有找到  strQry:" & strQry & Environment.NewLine)
                                    End If
                                    Return True
                                Else
                                    If Not sbDebug Is Nothing Then
                                        sbDebug.AppendLine("[SearchSO004] 失敗 SO004找不到;以SearchSO0102也反查不到SO004 strQry:" & strQry & Environment.NewLine)
                                    End If
                                    Return False
                                End If
                            Else
                                If Not sbDebug Is Nothing Then
                                    sbDebug.AppendLine("[SearchSO004] 失敗 SO004找不到;以SearchSO0102也找不到符合資料 strQry:" & strQry & Environment.NewLine)
                                End If
                                Return False
                            End If
                        Else
                            For i As Integer = 0 To dst.Tables("SO004").Rows.Count - 1
                                sCustId = String.Empty
                                '如果拆機日期為空,則直接尋找SO001與SO002的資料,如果有找到符合的資料,代表正常
                                If dst.Tables("SO004").Rows(i).IsNull("PRDate") Then
                                    sIVRDeclarantName = dst.Tables("SO004").Rows(i).Item("DeclarantName") & ""
                                    If SearchSO0102(sServiceType) Then
                                        blnReturn = True
                                        Exit For
                                    Else
                                        blnReturn = False
                                    End If
                                    '如果拆機日期不為空值,則以客編反查SO001與SO002的資料,如果找到符合的資料,再以客編反查SO004看資料是否符合
                                Else
                                    sCustId = dst.Tables("SO004").Rows(i).Item("CustId") & ""
                                    sIVRDeclarantName = String.Empty
                                    sIVRDeclarantName = dst.Tables("SO004").Rows(i).Item("DeclarantName") & ""
                                    If SearchSO0102(sServiceType) Then
                                        If ReverSO004(sServiceType, sRefNo) Then
                                            blnReturn = True
                                        Else
                                            blnReturn = False
                                        End If
                                    End If
                                End If
                            Next i
                        End If
                    End Using
                End Using
            End Using
            If Not sbDebug Is Nothing Then
                sbDebug.AppendLine("[SearchSO004] " & IIf(blnReturn, "成功", "失敗") & " stQry:" & strQry & Environment.NewLine)
            End If
            Return blnReturn
        Catch ex As Exception
            If Not sbDebug Is Nothing Then
                sbDebug.AppendLine("[SearchSO004] 載入失敗" & Environment.NewLine)
            End If
            Return False
        End Try
    End Function
    Private Function ReverSO004() As Boolean
        Dim strQry As String = "Select * From " & sOwnerName & "SO004" & _
                             " Where CompCode=" & iCompCode & _
                             " And CustId=" & sCustId & _
                             " And PRDate is Null"
        Try
            Using adapter As New OracleDataAdapter()
                adapter.SelectCommand = New OracleCommand(strQry, Oracn)
                Using builder As OracleCommandBuilder = New OracleCommandBuilder(adapter)
                    Using dst As New Data.DataSet()
                        adapter.Fill(dst, "SO004")
                        '找不到SO004的資料,則尋找SO001與SO002的資料,如果有找到則以客編反查SO004
                        If dst.Tables("SO004").Rows.Count <= 0 Then
                            If sbDebug Is Nothing Then
                                sbDebug.AppendLine("[ReverSO004] 反查客編失敗 strQry:" & strQry & Environment.NewLine)
                            End If
                            Return False
                        Else
                            If Not sbDebug Is Nothing Then
                                sbDebug.AppendLine("[ReverSO004] 反查客編成功 strQry:" & strQry & Environment.NewLine)
                            End If
                            If Not dst.Tables("SO004").Rows(0).IsNull("DeclarantName") Then
                                sIVRDeclarantName = dst.Tables("SO004").Rows(0).Item("DeclarantName") & ""
                            End If
                            Return True
                        End If
                    End Using
                End Using
            End Using
        Catch ex As Exception
            If Not sbDebug Is Nothing Then
                sbDebug.AppendLine("[ReverSO004] 載入失敗 strQry:" & strQry & Environment.NewLine)
            End If
            Return False
        End Try
    End Function
    ''' <summary>
    ''' 以客編反查SO004的資料
    ''' </summary>
    ''' <param name="sServiceType">服務別('C' Or 'C','P')</param>
    ''' <param name="sRefNo">設備參考號CD022(3 Or 2,5)</param>
    ''' <returns>布林值</returns>
    Private Function ReverSO004(ByVal sServiceType As String, ByVal sRefNo As String) As Boolean
        Dim strQry As String = "Select * From " & sOwnerName & "SO004" & _
                             " Where CompCode=" & iCompCode & _
                             IIf(sServiceType <> String.Empty, " ServiceType In(" & sServiceType & ")", "") & _
                             " And CustId=" & sCustId & _
                             " And PRDate is Null" & _
                             IIf(sRefNo <> String.Empty, " And FaciCode In ( Select CodeNo From CD022 Where RefNo In(" & sRefNo & "))", "")
        Try
            Using adapter As New OracleDataAdapter()
                adapter.SelectCommand = New OracleCommand(strQry, Oracn)
                Using builder As OracleCommandBuilder = New OracleCommandBuilder(adapter)
                    Using dst As New Data.DataSet()
                        adapter.Fill(dst, "SO004")
                        '找不到SO004的資料,則尋找SO001與SO002的資料,如果有找到則以客編反查SO004
                        If dst.Tables("SO004").Rows.Count <= 0 Then
                            If sbDebug Is Nothing Then
                                sbDebug.AppendLine("[ReverSO004] 反查客編失敗 strQry:" & strQry & Environment.NewLine)
                            End If
                            Return False
                        Else
                            If Not sbDebug Is Nothing Then
                                sbDebug.AppendLine("[ReverSO004] 反查客編成功 strQry:" & strQry & Environment.NewLine)
                            End If
                            If Not dst.Tables("SO004").Rows(0).IsNull("DeclarantName") Then
                                sIVRDeclarantName = dst.Tables("SO004").Rows(0).Item("DeclarantName") & ""
                            End If
                            Return True
                        End If
                    End Using
                End Using
            End Using


            
        Catch ex As Exception
            If Not sbDebug Is Nothing Then
                sbDebug.AppendLine("[ReverSO004] 載入失敗 strQry:" & strQry & Environment.NewLine)
            End If
            Return False
        End Try
    End Function
    ''' <summary>
    ''' 新增資料至SO552
    ''' </summary>
    ''' <param name="strCmd">命令代號</param>
    ''' <param name="sServiceType">服務別</param>
    ''' <returns>布林值</returns>
    Public Function InsertSO552(ByVal strCmd As String, ByVal sServiceType As String)
        Dim strInsert As String = "Insert Into " & sOwnerName & "SO552" & _
                                " (CmdID,CompCode,ServiceType,CustID,ErrorTime,CircuitNo) Values " & _
                                "(:CmdID,:CompCode,:ServiceType,:CustID,:ErrorTime,:CircuitNo)"
        Try
            Using cmd As New OracleCommand(strInsert, Oracn)
                cmd.Parameters.Add("CmdID", strCmd)
                cmd.Parameters.Add("CompCode", iCompCode)
                cmd.Parameters.Add("ServiceType", sServiceType)
                cmd.Parameters.Add("CustID", sCustId)
                cmd.Parameters.Add("ErrorTime", Date.Parse(strRunTime))
                cmd.Parameters.Add("CircuitNo", sCircuitNo)
                cmd.ExecuteNonQuery()
            End Using
            If Not sbDebug Is Nothing Then
                sbDebug.AppendLine("[InsertSO552] 成功 CmdID:" & strCmd & " CompCode:" & iCompCode & " SerViceType:" & sServiceType & _
                                 " CustID:" & sCustId & " ErrorTime:" & Date.Parse(strRunTime).ToString & " CircuitNo:" & sCircuitNo & Environment.NewLine)
            End If
            Return True
        Catch ex As Exception
            If Not sbDebug Is Nothing Then
                sbDebug.AppendLine("[InsertSO552] 載入入敗 CmdID:" & strCmd & " CompCode:" & iCompCode & " SerViceType:" & sServiceType & _
                                 " CustID:" & sCustId & " ErrorTime:" & Date.Parse(strRunTime).ToString & " CircuitNo:" & sCircuitNo & Environment.NewLine)
            End If
            Return False
        End Try
    End Function
    ''' <summary>
    ''' 新增資料至SO551
    ''' </summary>
    ''' <param name="sFromData">對方的參數資料</param>
    ''' <param name="sXml">將對方組合成的XML字串</param>
    ''' <param name="strCmd">命令代號</param>
    ''' <returns>布林值</returns>
    Public Function InsertSO551(ByVal sFromData As String, ByVal sXml As String, ByVal strCmd As String) As Boolean
        Dim strInsert As String = "Insert Into " & sOwnerName & "SO551" & _
                                " (ExchTime,CmdID,CompCode,CustID,XML_In,XML_Out) Values " & _
                                "(To_Date('" & Format(Date.Parse(strRunTime), "yyyyMMddHHmmss") & "','yyyymmddhh24miss')," & _
                                IIf(strCmd = String.Empty, "Null", "'" & strCmd & "'") & "," & _
                                IIf(iCompCode = 0, "Null", iCompCode) & "," & _
                                IIf(sCustId = String.Empty, "Null", sCustId) & "," & _
                                "'" & sFromData & "'," & _
                                "'" & sXml & "')"

        Try
            Using cmd As New OracleCommand(strInsert, Oracn)
                Try
                    cmd.ExecuteNonQuery()
                    If Not sbDebug Is Nothing Then
                        sbDebug.AppendLine("[InsertSO551] 成功 strInsert:" & strInsert & Environment.NewLine)
                    End If
                    Return True
                Catch ex As Exception
                    If Not sbDebug Is Nothing Then
                        sbDebug.AppendLine("[InsertSO551] 錯誤 strInsert:" & strInsert & Environment.NewLine)
                    End If
                    Return False
                End Try
            End Using
        Catch ex As Exception
            If Not sbDebug Is Nothing Then
                sbDebug.AppendLine("[InsertSO551] 載入失敗 strInsert:" & strInsert & Environment.NewLine)
            End If
            Return False
        End Try
    End Function
    ''' <summary>
    ''' 尋找SO001的CustName與Tel1的資料
    ''' </summary>
    ''' <returns>布林值</returns>
    ''' <remarks>新增SO006時會使用到</remarks>
    Private Function SearchCustName() As Boolean
        Dim strQry As String = "Select CustName,Tel1 From " & sOwnerName & "SO001" & _
                                " Where Custid=" & sCustId & _
                                " And CompCode=" & iCompCode
        Try
            Using cmd As New OracleCommand(strQry, Oracn)
                Dim rdr As OracleDataReader = cmd.ExecuteReader
                While rdr.Read()
                    sCustName = rdr.GetString(0) & ""
                    If rdr.IsDBNull(1) Then
                        sTel1 = String.Empty
                    Else
                        sTel1 = rdr.GetString(1) & ""
                    End If
                End While
                Return True
            End Using
        Catch ex As Exception
            Return False
        End Try
    End Function
    ''' <summary>
    ''' 新增資料至SO006
    ''' </summary>
    ''' <param name="sServiceType">服務別(例 C Or D)</param>
    ''' <param name="strCmd">命令代號</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function InsertSO006(ByVal sServiceType As String, ByVal strCmd As String) As Boolean
        
        Dim strSeqNo As String = GetInvoiceNo2()
        Dim strInsert As String = "Insert Into " & sOwnerName & "SO006" & _
                                " (CompCode,ServiceType,SeqNo,CustID,CustName,Tel1,AcceptTime,ServiceCode,ServiceName," & _
                                "NodeNo,ServDescCode,ServDescName,AcceptEn,AcceptName,UpdEn,UpdTime) Values (" & _
                                ":CompCode,:ServiceType,:SeqNo,:CustId,:CustName,:Tel1,:AcceptTime,:ServiceCode,:ServiceCode," & _
                                ":NodeNo,:ServDescCode,:ServDescName,:AcceptEn,:AcceptName,:UpdEn,:UpdTime)"
        Using cmd As New OracleCommand(strInsert, Oracn)
            Try
                cmd.CommandType = Data.CommandType.Text
                cmd.Parameters.Add("CompCode", iCompCode)
                cmd.Parameters.Add("ServiceType", sServiceType)
                cmd.Parameters.Add("SeqNo", strSeqNo)
                cmd.Parameters.Add("CustID", CType(sCustId, Integer))
                cmd.Parameters.Add("CustName", sCustName)
                cmd.Parameters.Add("Tel1", sTel1)
                cmd.Parameters.Add("AcceptTime", Date.Parse(strRunTime))
                cmd.Parameters.Add("ServiceCode", ReadIni(strCmd, "ServiceCode"))
                cmd.Parameters.Add("ServiceName", ReadIni(strCmd, "ServiceName"))
                cmd.Parameters.Add("NodeNo", sNodeNo)
                cmd.Parameters.Add("ServDescCode", ReadIni(strCmd, "ServDescCode"))
                cmd.Parameters.Add("ServDescName", ReadIni(strCmd, "ServDescName"))
                cmd.Parameters.Add("AcceptEn", ReadIni(strCmd, "AcceptEn"))
                cmd.Parameters.Add("AcceptName", ReadIni(strCmd, "AcceptName"))
                cmd.Parameters.Add("UpdEn", ReadIni(strCmd, "UpdEn"))
                cmd.Parameters.Add("UpdTime", (Integer.Parse(strRunTime.Substring(0, 4)) - 1911).ToString & Format(Date.Parse(strRunTime), "/MM/dd HH:mm:ss"))
                cmd.ExecuteNonQuery()
                If Not sbDebug Is Nothing Then
                    sbDebug.AppendLine("[InsertSO006] 成功 " & Environment.NewLine)
                End If
                Return True
            Catch ex As Exception
                If Not sbDebug Is Nothing Then
                    sbDebug.AppendLine("[InsertSO006] 載入失敗 " & Environment.NewLine)
                End If
                Return False
            End Try
        End Using
    End Function
    ''' <summary>
    ''' 呼叫後端
    ''' </summary>
    ''' <param name="iAddrNo">SO014.AddrNo</param>
    ''' <param name="iCompCode">公司別</param>
    ''' <param name="sServiceType">服務別(例 C Or I)</param>
    ''' <returns>後端的Output(p_IVRFileCode)</returns>
    Private Function SF_GETMFMSG(ByVal iAddrNo As Integer, ByVal iCompCode As Integer, ByVal sServiceType As String) As String
        Try
            Using Cmd As New OracleCommand("SF_GETMFMSG", Oracn)
                Cmd.CommandType = Data.CommandType.StoredProcedure
                Cmd.Parameters.Clear()
                With Cmd.Parameters
                    .Add("RETURN_VALUE", OracleDbType.Varchar2, Data.ParameterDirection.ReturnValue)
                    .Add("p_AddrNo", OracleDbType.Int32, 2000, iAddrNo, Data.ParameterDirection.Input)
                    .Add("p_CompCode", OracleDbType.Int32, 2000, iCompCode, Data.ParameterDirection.Input)
                    .Add("p_ServiceType", OracleDbType.Varchar2, 2000, sServiceType, Data.ParameterDirection.Input)
                    .Add("p_Msg", OracleDbType.Varchar2, Data.ParameterDirection.Output)
                    .Add("p_IVRFileCode", OracleDbType.Int32, Data.ParameterDirection.Output)
                End With
                Cmd.Parameters("RETURN_VALUE").Size = 4000
                Cmd.Parameters("p_Msg").Size = 4000
                Cmd.Parameters("p_IVRFileCode").Size = 4000
                Cmd.ExecuteNonQuery()
                If Not sbDebug Is Nothing Then
                    sbDebug.AppendLine("[SF_GETMFMSG] 成功 p_addrNo:" & iAddrNo.ToString & " p_CompCode:" & iCompCode.ToString & " p_ServiceType:" & sServiceType & Environment.NewLine)
                End If
                Return Cmd.Parameters("p_IVRFileCode").Value.ToString()
            End Using
        Catch ex As Exception
            If Not sbDebug Is Nothing Then
                sbDebug.AppendLine("[SF_GETMFMSG] 載入失敗 p_addrNo:" & iAddrNo.ToString & " p_CompCode:" & iCompCode.ToString & " p_ServiceType:" & sServiceType & Environment.NewLine)

            End If
            Return String.Empty
        End Try
    End Function
    Private Function GetSO041Para() As Boolean
        Dim strQry As String = String.Empty
        strQry = "Select VBCMCode,VBClosingDate From " & sOwnerName & "SO041"
        Try
            Using cmd As New OracleCommand(strQry, Oracn)
                Dim rdr As OracleDataReader = cmd.ExecuteReader
                While rdr.Read()
                    sCMCode = rdr.GetString(0) & ""
                    If Not rdr.IsDBNull(1) Then
                        iCloseDate = CType(rdr.GetValue(1), Integer)
                    End If
                End While
                Return True
            End Using
        Catch ex As Exception
            Return False
        End Try

    End Function
    Private Function GetNowDate() As DateTime
        Dim strQry As String = String.Empty
        strQry = "SELECT TO_CHAR(SYSDATE,'YYYY/MM/DD') FROM DUAL"
        Try
            Using cmd As New OracleCommand(strQry, Oracn)
                Dim rdr As OracleDataReader = cmd.ExecuteReader
                While rdr.Read
                    If Not rdr.IsDBNull(0) Then
                        Return CType(rdr.GetValue(0), DateTime)
                    End If
                End While
            End Using
        Catch ex As Exception
            If Not sbDebug Is Nothing Then
                sbDebug.AppendLine("取得系統時間失敗" & Environment.NewLine)
            End If
            Return Nothing
        End Try
    End Function
    Private Function GetSO033AMT(ByVal sServiceType As String) As Boolean
        Dim strQry As String = String.Empty
        Dim dTmpDate As DateTime = Nothing
        Dim dNow As DateTime = GetNowDate()
        Dim blnReturn As Boolean = False
        If Not GetSO041Para() Then
            If Not sbDebug Is Nothing Then
                sbDebug.AppendLine("取得系統參數失敗" & Environment.NewLine)
            End If
            Return False
        End If
        Try
            strQry = "Select A.ShouldAmt,A.BarcodeCloseDate From " & sOwnerName & "SO033 A," & sOwnerName & "CD013 B," & sOwnerName & "SO041 C" & _
                " Where A.ServiceType In (" & sServiceType & ")" & _
                " And A.UCCode is Not Null" & _
                " And A.CancelFlag<>1" & _
                " And A.RealDate is Null" & _
                " And A.UCCode=B.CodeNo" & _
                " And B.RefNo<>3" & _
                " And A.CMCode In(" & sCMCode & ")" & _
                " AND A.CUSTID=" & sCustId & _
                " AND A.BarcodeCloseDate IS NOT NULL"
            Using adapter As New OracleDataAdapter()
                adapter.SelectCommand = New OracleCommand(strQry, Oracn)
                Using builder As OracleCommandBuilder = New OracleCommandBuilder(adapter)
                    Using dst As New Data.DataSet()
                        adapter.Fill(dst, "SO033")
                        If dst.Tables("SO033").Rows.Count <= 0 Then
                            iStatus = 0
                        Else
                            For i As Integer = 0 To dst.Tables("SO033").Rows.Count - 1
                                dTmpDate = CType(dst.Tables("SO033").Rows(i).Item("BarcodeCloseDate"), Date)
                                If dNow.AddDays(iCloseDate) <= dTmpDate Then
                                    iTotal += dst.Tables("SO033").Rows(i).Item("ShouldAmt")
                                Else
                                    iOweTotal += dst.Tables("SO033").Rows(i).Item("ShouldAmt")
                                End If
                            Next
                        End If
                    End Using
                End Using
            End Using
            blnReturn = True
            If Not sbDebug Is Nothing Then
                sbDebug.AppendLine("[GetSO033AMT] " & IIf(blnReturn, "成功", "失敗") & "  strQry:" & strQry & Environment.NewLine)
            End If
            Return blnReturn

        Catch ex As Exception
            If Not sbDebug Is Nothing Then
                sbDebug.AppendLine("[GetSO033AMT] 載入失敗 strQry:" & strQry & Environment.NewLine)
            End If
            Return False
        End Try
        
    End Function
    Private Function SearchSO0102() As Boolean
        Dim strQry As String
        Dim blnReturn As Boolean = False
        strQry = "Select Distinct A.CustId,A.CustName,A.Tel1,A.Tel2,A.Tel3,B.CustStatusCode,A.InstAddrNo,A.InstAddress From " & sOwnerName & "SO001 A," & sOwnerName & "SO002 B" & _
                " Where (A.Tel1='" & sTel & "' Or A.Tel2='" & sTel & "' Or A.Tel3='" & sTel & "'" & _
                        " Or B.ContTel='" & sTel & "')" & _
                        " And A.CompCode=" & iCompCode & _
                        " And A.CompCode=B.CompCode" & _
                        " And A.CustId=B.CustId" & _
                        IIf(sCustId <> String.Empty, " And A.CustId=" & sCustId, "")
        Try
            Using adapter As New OracleDataAdapter()
                adapter.SelectCommand = New OracleCommand(strQry, Oracn)
                Using builder As OracleCommandBuilder = New OracleCommandBuilder(adapter)
                    'Oracn.Open()
                    Using dst As New Data.DataSet()
                        adapter.Fill(dst, "SO0102")
                        If dst.Tables("SO0102").Rows.Count <= 0 Then
                            If Not sbDebug Is Nothing Then
                                sbDebug.AppendLine("[SearchSO0102] 找不到資料 strQry:" & strQry & Environment.NewLine)
                            End If
                            Return False
                        Else
                            For i As Integer = 0 To dst.Tables("SO0102").Rows.Count - 1
                                If SearchSO014(dst.Tables("SO0102").Rows(i).Item("InstAddrNo")) Then
                                    sCustId = CType(dst.Tables("SO0102").Rows(i).Item("CustId"), String)
                                    iAddrNo = CType(dst.Tables("SO0102").Rows(i).Item("InstAddrNo"), Integer)
                                    sCustName = dst.Tables("SO0102").Rows(i).Item("CustName") & ""
                                    '************************************************************************************
                                    '加回傳Para1:客戶姓名,Para2:裝機地址

                                    sIVRDeclarantName = sCustName

                                    sIVRInstaddress = CType(dst.Tables("SO0102").Rows(i).Item("InstAddress"), String)
                                    '************************************************************************************
                                    If dst.Tables("SO0102").Rows(i).IsNull("Tel1") Then
                                        sTel1 = String.Empty
                                    Else
                                        sTel1 = dst.Tables("SO0102").Rows(i).Item("Tel1") & ""
                                    End If
                                    If dst.Tables("SO0102").Rows(i).IsNull("Tel2") Then
                                        sTel2 = String.Empty
                                    Else
                                        sTel2 = dst.Tables("SO0102").Rows(i).Item("Tel2") & ""
                                    End If
                                    If dst.Tables("SO0102").Rows(i).IsNull("Tel3") Then
                                        sTel3 = String.Empty
                                    Else
                                        sTel3 = dst.Tables("SO0102").Rows(i).Item("Tel3") & ""
                                    End If
                                    blnReturn = True
                                    Exit For
                                End If
                            Next i
                        End If
                    End Using
                End Using
            End Using
            If Not sbDebug Is Nothing Then
                sbDebug.AppendLine("[SearchSO001] " & IIf(blnReturn, "成功", "失敗") & "  strQry:" & strQry & Environment.NewLine)
            End If
            Return blnReturn
        Catch ex As Exception
            If Not sbDebug Is Nothing Then
                sbDebug.AppendLine("[SearchSO001] 載入失敗 strQry:" & strQry & Environment.NewLine)
            End If
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 尋找SO001與SO002的資料
    ''' </summary>
    ''' <param name="sServiceType">服務別(例'C' Or 'P','I')</param>
    ''' <returns>布林值</returns>
    ''' <remarks>尋找是否有符合的電話,並判斷SO002的客戶狀態</remarks>
    Private Function SearchSO0102(ByVal sServiceType As String) As Boolean
        Dim strQry As String
        Dim blnReturn As Boolean = False
        strQry = "Select Distinct A.CustId,A.CustName,A.Tel1,A.Tel2,A.Tel3,B.CustStatusCode,A.InstAddrNo,A.InstAddress From " & sOwnerName & "SO001 A," & sOwnerName & "SO002 B" & _
                " Where (A.Tel1='" & sTel & "' Or A.Tel2='" & sTel & "' Or A.Tel3='" & sTel & "'" & _
                        " Or B.ContTel='" & sTel & "')" & _
                        IIf(sServiceType <> String.Empty, " And B.ServiceType IN(" & sServiceType & ")", "") & _
                        " And A.CompCode=" & iCompCode & _
                        " And A.CompCode=B.CompCode" & _
                        " And A.CustId=B.CustId" & _
                        IIf(sCustId <> String.Empty, " And A.CustId=" & sCustId, "")
        Try
            Using adapter As New OracleDataAdapter()
                adapter.SelectCommand = New OracleCommand(strQry, Oracn)
                Using builder As OracleCommandBuilder = New OracleCommandBuilder(adapter)
                    'Oracn.Open()
                    Using dst As New Data.DataSet()
                        adapter.Fill(dst, "SO0102")
                        If dst.Tables("SO0102").Rows.Count <= 0 Then
                            If Not sbDebug Is Nothing Then
                                sbDebug.AppendLine("[SearchSO0102] 找不到資料 strQry:" & strQry & Environment.NewLine)
                            End If
                            Return False
                        Else
                            For i As Integer = 0 To dst.Tables("SO0102").Rows.Count - 1

                                Select Case dst.Tables("SO0102").Rows(i).Item("CustStatusCode")
                                    Case "3", "4", "5"
                                        sCustId = CType(dst.Tables("SO0102").Rows(i).Item("CustId"), String)
                                        iAddrNo = CType(dst.Tables("SO0102").Rows(i).Item("InstAddrNo"), Integer)
                                        If Not sbDebug Is Nothing Then
                                            sbDebug.AppendLine("[SearchSO0102] 找不正常戶 strQry:" & strQry & Environment.NewLine)
                                        End If
                                        blnReturn = False
                                    Case Else
                                        If SearchSO014(dst.Tables("SO0102").Rows(i).Item("InstAddrNo")) Then
                                            sCustId = CType(dst.Tables("SO0102").Rows(i).Item("CustId"), String)
                                            iAddrNo = CType(dst.Tables("SO0102").Rows(i).Item("InstAddrNo"), Integer)
                                            sCustName = dst.Tables("SO0102").Rows(i).Item("CustName") & ""
                                            '************************************************************************************
                                            '加回傳Para1:客戶姓名,Para2:裝機地址
                                            If sIVRDeclarantName = String.Empty Then
                                                sIVRDeclarantName = sCustName
                                            End If
                                            sIVRInstaddress = CType(dst.Tables("SO0102").Rows(i).Item("InstAddress"), String)
                                            '************************************************************************************
                                            If dst.Tables("SO0102").Rows(i).IsNull("Tel1") Then
                                                sTel1 = String.Empty
                                            Else
                                                sTel1 = dst.Tables("SO0102").Rows(i).Item("Tel1") & ""
                                            End If
                                            If dst.Tables("SO0102").Rows(i).IsNull("Tel2") Then
                                                sTel2 = String.Empty
                                            Else
                                                sTel2 = dst.Tables("SO0102").Rows(i).Item("Tel2") & ""
                                            End If
                                            If dst.Tables("SO0102").Rows(i).IsNull("Tel3") Then
                                                sTel3 = String.Empty
                                            Else
                                                sTel3 = dst.Tables("SO0102").Rows(i).Item("Tel3") & ""
                                            End If
                                            blnReturn = True
                                            Exit For
                                        End If
                                End Select
                            Next i
                        End If
                    End Using
                End Using
            End Using
            If Not sbDebug Is Nothing Then
                sbDebug.AppendLine("[SearchSO0102] " & IIf(blnReturn, "成功", "失敗") & "  strQry:" & strQry & Environment.NewLine)
            End If
            Return blnReturn
        Catch ex As Exception
            If Not sbDebug Is Nothing Then
                sbDebug.AppendLine("[SearchSO0102] 載入失敗 strQry:" & strQry & Environment.NewLine)
            End If
            Return False
        End Try

    End Function
    ''' <summary>
    ''' 取得SO006.Sequence
    ''' </summary>
    ''' <returns>String</returns>
    Private Function GetInvoiceNo2() As String
        Dim strSeq = "S_SO006"
        Dim strQry As String = "Select '" & Format(Date.Now, "yyyyMMdd") & _
                             "' ||Ltrim(To_Char(" & sOwnerName & strSeq & ".NextVal,'0999999')) From Dual"
        Try
            Dim cmd As New OracleCommand(strQry, Oracn)
            cmd.CommandType = Data.CommandType.Text
            Return cmd.ExecuteScalar
        Catch ex As Exception
            If Not sbDebug Is Nothing Then
                sbDebug.AppendLine("[GetInvoiceNo2] 載入失敗 strQry:" & strQry & Environment.NewLine)
            End If
            Return String.Empty
        End Try

    End Function
    ''' <summary>
    ''' 讀取虛擬目錄下的Ini檔
    ''' </summary>
    ''' <param name="sSection">Section</param>
    ''' <param name="sKey">Key</param>
    ''' <returns>String</returns>
    Private Function ReadIni(ByVal sSection As String, ByVal sKey As String) As String
        Dim sRet As String = Space(1024)
        GetPrivateProfileString(sSection, sKey, String.Empty, sRet, Len(sRet), sIniFile)
        sRet = sRet.Trim(New Char() {Chr(32), Chr(10), Chr(13), Chr(0)})
        If sRet = "" Then
            Return String.Empty
        Else
            Return sRet
        End If

    End Function
    ''' <summary>
    ''' 以尋找到SO001.InstAddrNo反查在SO014與XML.Floor的資料是否相符
    ''' </summary>
    ''' <param name="strAddrNo">SO001.InstAddrNo</param>
    ''' <returns>布林值</returns>
    ''' <remarks>如果符合，會指定NodeNo與CircuitNo</remarks>
    Private Function SearchSO014(ByVal strAddrNo As String) As Boolean
        Dim strQry As String
        Dim blnRet As Boolean = False
        strQry = "Select Flour,NodeNo,CircuitNo From " & sOwnerName & "SO014" & _
               " Where CompCode=" & iCompCode & _
               " And AddrNo=" & strAddrNo

        Try
            Using adapter As New OracleDataAdapter()
                adapter.SelectCommand = New OracleCommand(strQry, Oracn)
                Using builder As OracleCommandBuilder = New OracleCommandBuilder(adapter)
                    Using dst As New Data.DataSet()
                        adapter.Fill(dst, "SO014")
                        If dst.Tables("SO014").Rows.Count <= 0 Then
                            If Not sbDebug Is Nothing Then
                                sbDebug.AppendLine("[SearchSO014] 找不到資料 strQry:" & strQry & Environment.NewLine)
                            End If
                            Return False
                        Else
                            For i As Integer = 0 To dst.Tables("SO014").Rows.Count - 1
                                If dst.Tables("SO014").Rows(i).Item("Flour") & "" = sFloor Then
                                    sNodeNo = dst.Tables("SO014").Rows(i).Item("NodeNo") & ""
                                    sCircuitNo = dst.Tables("SO014").Rows(i).Item("CircuitNo") & ""
                                    blnRet = True
                                    Exit For
                                End If
                            Next i
                        End If
                    End Using
                End Using
            End Using
            If Not sbDebug Is Nothing Then
                sbDebug.AppendLine("[SearchSO014] " & IIf(blnRet, "成功", "失敗") & " strQry:" & strQry & Environment.NewLine)
            End If
            Return blnRet
        Catch ex As Exception
            If Not sbDebug Is Nothing Then
                sbDebug.AppendLine("[SearchSO014] 載入失敗 strQry:" & strQry & Environment.NewLine)
            End If
            Return False
        End Try

    End Function
    Private Sub GetOwner()
        Dim strOwner As String = String.Empty
        Try
            Using cmd As New OracleCommand("Select TableOwner From CD039 Where CodeNo=" & iCompCode, Oracn)
                cmd.CommandType = Data.CommandType.Text
                strOwner = CType(cmd.ExecuteScalar, String).ToString()
                If strOwner <> String.Empty AndAlso Not strOwner Is Nothing Then
                    If Not sbDebug Is Nothing Then
                        sbDebug.AppendLine("[GetOwner] 成功 OwnerName:" & strOwner & " strQry:Select TableOwner From CD039 Where CodeNo=" & iCompCode & Environment.NewLine)
                    End If
                    sOwnerName = strOwner & "."

                Else
                    If Not sbDebug Is Nothing Then
                        sbDebug.AppendLine("[GetOwner] 成功，但為空值 strQry:Select TableOwner From CD039 Where CodeNo=" & iCompCode & Environment.NewLine)
                    End If
                    sOwnerName = String.Empty
                End If
            End Using

        Catch ex As Exception
            If Not sbDebug Is Nothing Then
                sbDebug.AppendLine("[GetOwner] 載入失敗" & Environment.NewLine)
            End If

        End Try
    End Sub
    Public ReadOnly Property uTel1() As String
        Get
            Return sTel1
        End Get
    End Property
    Public ReadOnly Property uTel2() As String
        Get
            Return sTel2
        End Get
    End Property
    Public ReadOnly Property uTel3() As String
        Get
            Return sTel3
        End Get
    End Property
    Public ReadOnly Property uNodeNo() As String
        Get
            Return sNodeNo
        End Get
    End Property
    Public ReadOnly Property uCircuitNo() As String
        Get
            Return sCircuitNo
        End Get
    End Property
    Public ReadOnly Property uCnn() As OracleConnection
        Get
            Return Oracn
        End Get
    End Property
    Public ReadOnly Property SumTotal() As String
        Get
            Return iTotal.ToString
        End Get
    End Property
    Public ReadOnly Property OweTotal() As String
        Get
            Return iOweTotal.ToString
        End Get
    End Property
    Public ReadOnly Property IVRStatus() As String
        Get
            Return iStatus.ToString
        End Get
    End Property
    Public ReadOnly Property IVRDeclarantName() As String
        Get
            Return sIVRDeclarantName
        End Get
    End Property
    Public ReadOnly Property IVRInstaddress() As String
        Get
            Return sIVRInstaddress
        End Get
    End Property
    Public ReadOnly Property IVRFileCode() As String
        Get
            Return sIVRFileCode
        End Get
    End Property
    Public ReadOnly Property CustId() As Integer
        Get
            Return CType(sCustId, Integer)
        End Get
    End Property
    Protected Overrides Sub Finalize()
        'OraCn = Nothing
        MyBase.Finalize()
    End Sub
End Class
