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
    Private sCustName As String = String.Empty
    Private sCircuitNo As String = String.Empty
    Private stmDug As StreamWriter
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
        sOwnerName = strOwner
        sCnn = strCnn
        Oracn.ConnectionString = sCnn
        Try
            Oracn.Open()
        Catch ex As Exception

        End Try

    End Sub
    Sub New(ByVal strTel As String, _
            ByVal strFloor As String, _
            ByVal intComp As Integer, _
            ByVal strOwner As String, _
            ByVal strIniFile As String, _
            ByVal strCnn As String, _
            ByRef stm As StreamWriter)
        stmDug = stm
        strRunTime = Date.Now.ToString
        sTel = strTel
        sFloor = strFloor
        iCompCode = intComp
        sIniFile = strIniFile
        sOwnerName = strOwner
        sCnn = strCnn
        Oracn.ConnectionString = sCnn
        Try
            Oracn.Open()
        Catch ex As Exception
            If Not stmDug Is Nothing Then
                stmDug.WriteLine("[New] Oracle開啟失敗")
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
    ''' <summary>
    ''' 尋找SO004是否有符合的資料
    ''' </summary>
    ''' <param name="sServiceType">服務別(例'C' Or 'I','P')</param>
    ''' <param name="sRefNo">設備參考號(CD022.RefNo)</param>
    ''' <returns>布林值</returns>
    Private Function SearchSO004(ByVal sServiceType As String, ByVal sRefNo As String) As Boolean
        Dim blnReturn As Boolean = False
        Dim strQry As String = "Select CustId,PRDate From " & sOwnerName & "SO004" & _
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
                                    Return True
                                Else
                                    If Not stmDug Is Nothing Then
                                        stmDug.WriteLine("[SearchSO004] 失敗 SO004找不到;以SearchSO0102也反查不到SO004 strQry:" & strQry)
                                    End If
                                    Return False
                                End If
                            Else
                                If Not stmDug Is Nothing Then
                                    stmDug.WriteLine("[SearchSO004] 失敗 SO004找不到;以SearchSO0102也找不到符合資料 strQry:" & strQry)
                                End If
                                Return False
                            End If
                        Else
                            For i As Integer = 0 To dst.Tables("SO004").Rows.Count - 1
                                sCustId = String.Empty
                                '如果拆機日期為空,則直接尋找SO001與SO002的資料,如果有找到符合的資料,代表正常
                                If dst.Tables("SO004").Rows(i).IsNull("PRDate") Then
                                    If SearchSO0102(sServiceType) Then
                                        blnReturn = True
                                        Exit For
                                    Else
                                        blnReturn = False
                                    End If
                                    '如果拆機日期不為空值,則以客編反查SO001與SO002的資料,如果找到符合的資料,再以客編反查SO004看資料是否符合
                                Else
                                    sCustId = dst.Tables("SO004").Rows(i).Item("CustId") & ""
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
            If Not stmDug Is Nothing Then
                stmDug.WriteLine("[SearchSO004] Return:" & blnReturn.ToString & " stQry:" & strQry)
            End If
            Return blnReturn
        Catch ex As Exception
            If Not stmDug Is Nothing Then
                stmDug.WriteLine("[SearchSO004] 載入失敗")
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
        Dim strQry As String = "Select Count(*) From " & sOwnerName & "SO004" & _
                             " Where ServiceType In(" & sServiceType & ")" & _
                             " And CompCode=" & iCompCode & _
                             " And CustId=" & sCustId & _
                             " And PRDate is Null" & _
                             " And FaciCode In ( Select CodeNo From CD022 Where RefNo In(" & sRefNo & "))"
        Try
            Using cmd As New OracleCommand(strQry, Oracn)
                If CType(cmd.ExecuteScalar, Integer) = 0 Then
                    Return False
                Else
                    Return True
                End If
            End Using
            stmDug.WriteLine("[ReverSO004] strQry:" & strQry)
        Catch ex As Exception
            If Not stmDug Is Nothing Then
                stmDug.WriteLine("[ReverSO004] 失敗 strQry:" & strQry)
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
            If Not stmDug Is Nothing Then
                stmDug.WriteLine("[InsertSO552] 成功 CmdID:" & strCmd & " CompCode:" & iCompCode & " SerViceType:" & sServiceType & _
                                 " CustID:" & sCustId & " ErrorTime:" & Date.Parse(strRunTime).ToString & " CircuitNo:" & sCircuitNo)
            End If
            Return True
        Catch ex As Exception
            If Not stmDug Is Nothing Then
                stmDug.WriteLine("[InsertSO552] 錯誤 CmdID:" & strCmd & " CompCode:" & iCompCode & " SerViceType:" & sServiceType & _
                                 " CustID:" & sCustId & " ErrorTime:" & Date.Parse(strRunTime).ToString & " CircuitNo:" & sCircuitNo)
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
                    If Not stmDug Is Nothing Then
                        stmDug.WriteLine("[InsertSO551] 成功 strInsert:" & strInsert)
                    End If
                    Return True
                Catch ex As Exception
                    If Not stmDug Is Nothing Then
                        stmDug.WriteLine("[InsertSO551] 錯誤 strInsert:" & strInsert)
                    End If
                    Return False
                End Try
            End Using
        Catch ex As Exception
            If Not stmDug Is Nothing Then
                stmDug.WriteLine("[InsertSO551] 失敗 strInsert:" & strInsert)
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
                If Not stmDug Is Nothing Then
                    stmDug.WriteLine("[InsertSO006] 成功 ")
                End If
                Return True
            Catch ex As Exception
                If Not stmDug Is Nothing Then
                    stmDug.WriteLine("[InsertSO006] 失敗 ")
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
                If Not stmDug Is Nothing Then
                    stmDug.WriteLine("[SF_GETMFMSG] 成功 p_addrNo:" & iAddrNo.ToString & " p_CompCode:" & iCompCode.ToString & " p_ServiceType:" & sServiceType)
                End If
                Return Cmd.Parameters("p_IVRFileCode").Value.ToString()
            End Using
        Catch ex As Exception
            If Not stmDug Is Nothing Then
                stmDug.WriteLine("[SF_GETMFMSG] 失敗 p_addrNo:" & iAddrNo.ToString & " p_CompCode:" & iCompCode.ToString & " p_ServiceType:" & sServiceType)

            End If
            Return String.Empty
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
        strQry = "Select Distinct A.CustId,A.CustName,A.Tel1,B.CustStatusCode,A.InstAddrNo From " & sOwnerName & "SO001 A," & sOwnerName & "SO002 B" & _
                " Where (A.Tel1='" & sTel & "' Or A.Tel2='" & sTel & "' Or A.Tel3='" & sTel & "'" & _
                        " Or B.ContTel='" & sTel & "')" & _
                        " And B.ServiceType IN(" & sServiceType & ")" & _
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
                            If Not stmDug Is Nothing Then
                                stmDug.WriteLine("[SearchSO0102] 找不到資料!")
                                stmDug.WriteLine("[SearchSO0102] strQry=" & strQry)
                            End If
                            Return False
                        Else
                            For i As Integer = 0 To dst.Tables("SO0102").Rows.Count - 1
                                Select Case dst.Tables("SO0102").Rows(i).Item("CustStatusCode")
                                    Case "3", "4", "5"
                                        sCustId = CType(dst.Tables("SO0102").Rows(i).Item("CustId"), String)
                                        iAddrNo = CType(dst.Tables("SO0102").Rows(i).Item("InstAddrNo"), Integer)
                                        If Not stmDug Is Nothing Then
                                            stmDug.WriteLine("[SearchSO0102] 找不正常戶!")
                                            stmDug.WriteLine("[SearchSO0102] strQry=" & strQry)
                                        End If
                                        blnReturn = False
                                    Case Else
                                        If SearchSO014(dst.Tables("SO0102").Rows(i).Item("InstAddrNo")) Then
                                            sCustId = CType(dst.Tables("SO0102").Rows(i).Item("CustId"), String)
                                            iAddrNo = CType(dst.Tables("SO0102").Rows(i).Item("InstAddrNo"), Integer)
                                            sCustName = dst.Tables("SO0102").Rows(i).Item("CustName") & ""
                                            If dst.Tables("SO0102").Rows(i).IsNull("Tel1") Then
                                                sTel1 = String.Empty
                                            Else
                                                sTel1 = dst.Tables("SO0102").Rows(i).Item("Tel1") & ""
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
            If Not stmDug Is Nothing Then
                stmDug.WriteLine("[SearchSO0102] Return:" & blnReturn & "  strQry:" & strQry)
            End If
            Return blnReturn
        Catch ex As Exception
            If Not stmDug Is Nothing Then
                stmDug.WriteLine("[SearchSO0102] 失敗")
                stmDug.WriteLine("[SearchSO0102] strQry=" & strQry)
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
            If Not stmDug Is Nothing Then
                stmDug.WriteLine("[GetInvoiceNo2] 失敗 strQry:" & strQry)
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
                            stmDug.WriteLine("[SearchSO014] 找不到資料")
                            stmDug.WriteLine("[SearchSO014] strQry:" & strQry)
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
            If Not stmDug Is Nothing Then
                stmDug.WriteLine("[SearchSO014] Return:" & blnRet.ToString & " strQry:" & strQry)
            End If
            Return blnRet
        Catch ex As Exception
            If Not stmDug Is Nothing Then
                stmDug.WriteLine("[SearchSO014] 失敗")
                stmDug.WriteLine("[SearchSO014] strQry:" & strQry)
            End If
            Return False
        End Try

    End Function
    Public ReadOnly Property uCnn() As OracleConnection
        Get
            Return Oracn
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
