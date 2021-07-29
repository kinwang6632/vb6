Imports System
Imports System.Data
Imports System.Data.OracleClient
Public Class clsNagra
    Private _STBNo As String
    Private _ICCNo As String
    Private _ProcessType As String
    Private _cn As OracleConnection
    Private _tra As OracleTransaction
    Private _cnString As String
    Private _OwnerName As String
    
    Private _Company As String
    Private _Seq As String
    Private strZIP_CODE As String = String.Empty
    Private strBAT_ID As String = String.Empty
    Private _CmdString As String = String.Empty

    Private strUserId As String = String.Empty

    Private strNewFaci As String = String.Empty
    Private strNewProduct As String = String.Empty
    Private strNewPinCode As String = String.Empty
    Private _TimeOut As Int32 = 30
    Private FRetResult As NagraRet
    Private FNagraCode As String = String.Empty
    Private FNagraMsg As String = String.Empty
    Private strHDSpace As String = String.Empty
    Private strHDNo As String = String.Empty
    Private _ExeTime As Date = Nothing
    Private FNewFaci As ArrayList = Nothing
    Private FNewChCode As ArrayList = Nothing
    Private FChCode As String = String.Empty
    Public Property Transcation()
        Get
            Return _tra
        End Get
        Set(ByVal value)
            _tra = value
        End Set
    End Property
    Public Property OracleConnect()
        Get
            Return _cn
        End Get
        Set(ByVal value)
            _cn = value
        End Set
    End Property
    Public Property STBNo()
        Get
            Return _STBNo
        End Get
        Set(ByVal value)
            _STBNo = value
        End Set
    End Property
    Public Property ICCNo()
        Get
            Return _ICCNo
        End Get
        Set(ByVal value)
            _ICCNo = value
        End Set
    End Property
    Public Property ConnectString()
        Get
            Return _cnString
        End Get
        Set(ByVal value)
            _cnString = value
        End Set
    End Property
    Public Property ProcessType()
        Get
            Return _ProcessType
        End Get
        Set(ByVal value)
            _ProcessType = value
        End Set
    End Property
    Public Property Company()
        Get
            Return _Company
        End Get
        Set(ByVal value)
            _Company = value
        End Set
    End Property
    Public Property SEQ() As String
        Get
            Return _Seq
        End Get
        Set(ByVal value As String)
            _Seq = value
        End Set
    End Property
    Public Property CmdString()
        Get
            Return _CmdString
        End Get
        Set(ByVal value)
            _CmdString = value
        End Set
    End Property
    Public Property TimeOut()
        Get
            Return _TimeOut
        End Get
        Set(ByVal value)
            _TimeOut = value
        End Set
    End Property
    Public Property ExeTime() As Date
        Get
            Return _ExeTime
        End Get
        Set(ByVal value As Date)
            _ExeTime = value
        End Set
    End Property
    Public Sub New()

    End Sub
    Public Sub New(ByVal aCmdCode As String, ByVal aCn As OracleConnection, ByVal aTra As OracleTransaction)
        _ProcessType = aCmdCode
        _tra = aTra
        _cn = aCn
    End Sub
    Public ReadOnly Property RetResult() As NagraRet
        Get
            Return FRetResult
        End Get
    End Property
    Public ReadOnly Property NagraCode() As String
        Get
            Return FNagraCode
        End Get
    End Property
    Public ReadOnly Property NagraMsg() As String
        Get
            Return FNagraMsg
        End Get
    End Property

    Public Function HoldCmd() As Boolean
        Dim aSQL As String = String.Empty
        Dim aStp As New Stopwatch
        Dim r As OracleDataReader
        Dim cmd As OracleCommand = Nothing
        aSQL = "SELECT CMD_STATUS,ERR_CODE,ERR_MSG FROM " & _OwnerName & "SEND_NAGRA " & _
            " WHERE SEQNO=" & _Seq
        Try
            If _cn.State = ConnectionState.Closed Then
                _cn.Open()
            End If
            cmd = _cn.CreateCommand
            If _tra.Connection IsNot Nothing Then
                cmd.Transaction = _tra
            End If
            cmd.CommandText = aSQL
        Catch ex As Exception
            Return False
            Throw New Exception("Function : ClsNagra.HoldCmd 原因 : " & ex.Message)
        End Try

        Try
            aStp.Start()

            Do

                If aStp.Elapsed.TotalSeconds >= _TimeOut Then
                    FRetResult = NagraRet.TimeOut
                    Exit Do
                End If
                r = cmd.ExecuteReader()
                If r.Read Then
                    If r.Item("CMD_STATUS") <> "W" Then
                        If r.Item("CMD_STATUS") = "S" OrElse r.Item("CMD_STATUS") = "C" Then
                            FRetResult = NagraRet.Complete
                            r.Close()
                            Return True
                            Exit Do
                        Else
                            If Not IsDBNull(r.Item("ERR_CODE")) Then
                                FRetResult = NagraRet.Nagra_RetErr
                                FNagraCode = r.Item("ERR_CODE")
                                FNagraMsg = r.Item("ERR_MSG")
                                r.Close()
                                Return False
                                Exit Do
                            End If
                        End If
                    End If
                End If
                r.Close()
            Loop
        Catch ex As Exception
            Throw New Exception("Function : ClsNagra.HoldCmd 原因 : " & ex.Message)
        Finally
            If aStp.IsRunning Then
                aStp.Stop()
            End If
            
        End Try

    End Function
    Public Function UpdData() As Boolean
        Try
            If _cn.State = ConnectionState.Closed Then
                _cn.Open()
            End If
            If _tra.Connection Is Nothing Then
                _tra = _cn.BeginTransaction
            End If
            Select Case _ProcessType
                Case "A2"
                    If Not UpdA2Cmd() Then Return False
                Case "A6"
                    If Not UpdA6Cmd() Then Return False
                Case "A7"
                    If Not UpdA7Cmd() Then Return False
                Case "A8", "A9"
                    If Not UpdA8Cmd() Then Return False
                Case "B1"
                    If Not UpdB1Cmd() Then Return False
                Case "B2", "B3", "B6"
                    If Not UpdB2Cmd() Then Return False
                Case "B10", "E17"
                    If Not UpdB10Cmd() Then Return False

            End Select
            Return True
        Catch ex As Exception
            FRetResult = NagraRet.SysError
            FNagraMsg = ex.Message
            Return False
        End Try
    End Function
    Public Function ExeCmd() As Boolean
        
        If _tra.Connection Is Nothing Then
            _tra = _cn.BeginTransaction
        End If
        Dim dt As New DataSet()
        Dim dap As New OracleDataAdapter("SELECT * FROM " & _OwnerName & "SEND_NAGRA WHERE 1=0", _cn)
        Dim bdu As New OracleCommandBuilder(dap)
        dap.SelectCommand.Transaction = _tra

        Try
            dap.Fill(dt, "NAGRA")
            If Not ChkICCNo() Then

                FRetResult = NagraRet.No_ICCNo
                Return False
            End If
            If Not ChkSTBNo() Then
                FRetResult = NagraRet.NO_STBNo
                Return False
            End If
            If Not ChkHDNo() Then
                If _ProcessType = "B10" Then
                    FRetResult = NagraRet.NO_HDNo
                Else
                    FRetResult = NagraRet.NO_Dvr
                End If
            End If
            GetPara()
            Select Case _ProcessType
                Case "A1"
                    Dim row As DataRow = A1Cmd(dt.Tables(0))
                    dt.Tables(0).Rows.Add(row)
                Case "A2", "A3", "A4"
                    Dim row As DataRow = A2Cmd(dt.Tables(0))
                    dt.Tables(0).Rows.Add(row)
                Case "A6"
                    Dim row As DataRow = A6Cmd(dt.Tables(0))
                    dt.Tables(0).Rows.Add(row)
                Case "A7"
                    Dim row As DataRow = A7Cmd(dt.Tables(0))
                    dt.Tables(0).Rows.Add(row)
                Case "A8"
                    Dim row As DataRow = A8Cmd(dt.Tables(0))
                    dt.Tables(0).Rows.Add(row)
                Case "A9"
                    Dim row As DataRow = A9Cmd(dt.Tables(0))
                    dt.Tables(0).Rows.Add(row)
                Case "B1", "B2", "B3", "B4", "B5", "B6", "B7"
                    Dim row As DataRow = B1Cmd(dt.Tables(0))
                    dt.Tables(0).Rows.Add(row)
                Case "B9"
                    Dim row As DataRow = B9Cmd(dt.Tables(0))
                    dt.Tables(0).Rows.Add(row)
                Case "B10"
                    Dim row As DataRow = B10Cmd(dt.Tables(0))
                    dt.Tables(0).Rows.Add(row)
                Case "C1"
                    Dim row As DataRow = C1Cmd(dt.Tables(0))
                    dt.Tables(0).Rows.Add(row)
                Case "C4"
                    Dim row As DataRow = C4Cmd(dt.Tables(0))
                    dt.Tables(0).Rows.Add(row)
                Case "E1"
                    Dim row As DataRow = E1Cmd(dt.Tables(0))
                    dt.Tables(0).Rows.Add(row)
                Case "E2-M", "E2-P"
                    Dim row As DataRow = E2MCmd(dt.Tables(0))
                    dt.Tables(0).Rows.Add(row)
                Case "E3", "E9", "E18"
                    Dim row As DataRow = E3Cmd(dt.Tables(0))
                    dt.Tables(0).Rows.Add(row)
                Case "E15", "E16"
                    Dim row As DataRow = E15Cmd(dt.Tables(0))
                    dt.Tables(0).Rows.Add(row)
                Case "E17"
                    Dim row As DataRow = E17Cmd(dt.Tables(0))
                    dt.Tables(0).Rows.Add(row)

                Case Else
                    FRetResult = NagraRet.SysError
                    FNagraMsg = "不正確的命令"
                    Return False
            End Select
            dap.Update(dt.Tables(0))
            _tra.Commit()
            If HoldCmd() Then

                Return True
            End If

        Catch ex As Exception
            If _tra.Connection IsNot Nothing Then
                _tra.Rollback()
            End If

            FRetResult = NagraRet.SysError
            FNagraMsg = ex.Message
            Return False
        End Try


    End Function


    'A1(開機)
    Private Function A1Cmd(ByVal aDT As DataTable) As DataRow
        Try
            Dim row As DataRow = aDT.NewRow()
            row.BeginEdit()
            row.Item("ICC_NO") = _ICCNo
            row.Item("STB_NO") = _STBNo
            row.Item("ZIP_CODE") = strZIP_CODE
            row.Item("MIS_IRD_CMD_Data") = strBAT_ID
            row.Item("PinCode") = _CmdString
            UpdCommon(row)
            row.EndEdit()
            Return row
        Catch ex As Exception
            Throw New Exception("Function : ClsNagra.A1Cmd 原因 : " & ex.Message)
        End Try
    End Function
    'A2、A3、A4(拆機、停機、復機)
    Private Function A2Cmd(ByVal aDt As DataTable) As DataRow
        Try
            Dim row As DataRow = aDt.NewRow()
            row.BeginEdit()
            row.Item("ICC_NO") = _ICCNo

            UpdCommon(row)
            row.EndEdit()
            Return row
        Catch ex As Exception
            Throw New Exception("Function : ClsNagra.A2Cmd 原因 : " & ex.Message)
        End Try
    End Function
    'A6 維修換拆(換整組)
    Private Function A6Cmd(ByVal aDt As DataTable) As DataRow
        Try
            Dim row As DataRow = aDt.NewRow()
            row.BeginEdit()
            row.Item("NOTES") = SpliteA6Cmd(";", _CmdString)
            row.Item("ZIP_CODE") = strZIP_CODE
            row.Item("MIS_IRD_CMD_Data") = strBAT_ID
            row.Item("PinCode") = strNewPinCode
            UpdCommon(row)
            row.EndEdit()
            Return row
        Catch ex As Exception
            Throw New Exception("Function : ClsNagra.A6Cmd 原因 : " & ex.Message)
        End Try
    End Function
    Private Function A7Cmd(ByVal aDt As DataTable) As DataRow
        Try
            Dim row As DataRow = aDt.NewRow()
            row.BeginEdit()
            row.Item("ICC_NO") = _ICCNo
            row.Item("STB_NO") = _STBNo

            UpdCommon(row)
            row.EndEdit()
            Return row
        Catch ex As Exception
            Throw New Exception("Function : ClsNagra.A7Cmd 原因 : " & ex.Message)
        End Try
    End Function
    Private Function A8Cmd(ByVal aDt As DataTable) As DataRow
        Try
            Dim row As DataRow = aDt.NewRow()
            row.BeginEdit()
            row.Item("NOTES") = _STBNo & "/" & _ICCNo & "@" & _
                        _CmdString.Split(";")(0) & "/" & _ICCNo
            row.Item("MIS_IRD_CMD_Data") = strBAT_ID
            If _CmdString.Contains(";") Then
                row.Item("PinCode") = _CmdString.Split(";").GetValue(1)
            End If
            UpdCommon(row)
            row.EndEdit()
            Return row
        Catch ex As Exception
            Throw New Exception("Function : ClsNagra.A8Cmd 原因 : " & ex.Message)
        End Try
    End Function

    Private Function A9Cmd(ByVal aDt As DataTable) As DataRow
        Try
            Dim row As DataRow = aDt.NewRow()
            row.BeginEdit()
            row.Item("NOTES") = SpliteA6Cmd(";", _CmdString)
            row.Item("ZIP_CODE") = strZIP_CODE
            UpdCommon(row)
            row.EndEdit()
            Return row
        Catch ex As Exception
            Throw New Exception("Function : ClsNagra.A9Cmd 原因 : " & ex.Message)
        End Try
    End Function

    Private Function B1Cmd(ByVal aDt As DataTable) As DataRow
        Try
            Dim row As DataRow = aDt.NewRow()
            row.BeginEdit()
            row.Item("ICC_NO") = _ICCNo
            Select Case _ProcessType
                Case "B1", "B5", "B6"
                    row.Item("NOTES") = SpliteA6Cmd("", _CmdString)
                Case "B2", "B4"
                    row.Item("NOTES") = _CmdString.Replace("~", ",")
                Case "B7"
                    row.Item("NOTES") = _CmdString
            End Select
            

            UpdCommon(row)
            row.EndEdit()
            Return row
        Catch ex As Exception
            Throw New Exception("Function : ClsNagra.B1Cmd 原因 : " & ex.Message)
        End Try
    End Function
    Private Function B9Cmd(ByVal aDt As DataTable) As DataRow
        Try
            Dim row As DataRow = aDt.NewRow()
            row.BeginEdit()
            row.Item("ICC_NO") = _ICCNo
            row.Item("STB_NO") = _STBNo
            row.Item("NOTES") = SpliteA6Cmd(";", _CmdString)
            row.Item("ZIP_CODE") = strZIP_CODE
            row.Item("MIS_IRD_CMD_Data") = strBAT_ID
            row.Item("PinCode") = strNewPinCode
            UpdCommon(row)
            row.EndEdit()
            Return row
        Catch ex As Exception
            Throw New Exception("Function : ClsNagra.B9Cmd 原因 : " & ex.Message)
        End Try
    End Function
    Private Function B10Cmd(ByVal aDt As DataTable) As DataRow
        Try
            Dim row As DataRow = aDt.NewRow()
            row.BeginEdit()
            row.Item("ICC_NO") = _ICCNo
            row.Item("STB_NO") = _STBNo
            row.Item("NOTES") = SplitB10Cmd()
            row.Item("MIS_IRD_CMD_Data") = strHDNo
            row.Item("DVRQuota") = strHDSpace
            UpdCommon(row)
            row.EndEdit()
            Return row
        Catch ex As Exception
            Throw New Exception("Function : ClsNagra.B10Cmd 原因 : " & ex.Message)
        End Try
    End Function
    Private Function C1Cmd(ByVal aDt As DataTable) As DataRow
        Try
            Dim row As DataRow = aDt.NewRow()
            row.BeginEdit()
            row.Item("NOTES") = _CmdString
            UpdCommon(row)
            row.EndEdit()
            Return row
        Catch ex As Exception
            Throw New Exception("Function : ClsNagra.C1Cmd 原因 : " & ex.Message)
        End Try
    End Function
    Private Function C4Cmd(ByVal aDt As DataTable) As DataRow
        Try
            Dim row As DataRow = aDt.NewRow()
            row.BeginEdit()
            row.Item("ICC_NO") = _ICCNo
            row.Item("zip_code") = strZIP_CODE
            UpdCommon(row)
            row.EndEdit()
            Return row
        Catch ex As Exception
            Throw New Exception("Function : ClsNagra.C4Cmd 原因 : " & ex.Message)
        End Try
    End Function
    Private Function E1Cmd(ByVal aDt As DataTable) As DataRow
        Try
            Dim row As DataRow = aDt.NewRow()
            row.BeginEdit()
            row.Item("ICC_NO") = _ICCNo
            row.Item("STB_NO") = _STBNo
            row.Item("PinCode") = _CmdString
            UpdCommon(row)
            row.EndEdit()
            Return row
        Catch ex As Exception
            Throw New Exception("Function : ClsNagra.E1Cmd 原因 : " & ex.Message)
        End Try
    End Function
    Private Function E2MCmd(ByVal aDt As DataTable) As DataRow
        Try
            Dim row As DataRow = aDt.NewRow()
            row.BeginEdit()
            row.Item("ICC_NO") = _ICCNo
            row.Item("STB_NO") = _STBNo
            row.Item("Notes") = StrConv(_CmdString, VbStrConv.Wide)
            If _ProcessType = "E2-M" Then
                row.Item("MIS_IRD_CMD_Data") = 0
            Else
                row.Item("MIS_IRD_CMD_Data") = 1
            End If
            UpdCommon(row)
            row.EndEdit()
            Return row
        Catch ex As Exception
            Throw New Exception("Function : ClsNagra.E2MCmd 原因 : " & ex.Message)
        End Try
    End Function
    Private Function E3Cmd(ByVal aDt As DataTable) As DataRow
        Try
            Dim row As DataRow = aDt.NewRow()
            row.BeginEdit()
            row.Item("ICC_NO") = _ICCNo
            row.Item("STB_NO") = _STBNo
            If (_ProcessType = "E3") OrElse (_ProcessType = "E18") Then
                row.Item("MIS_IRD_CMD_Data") = _CmdString
            Else
                row.Item("MIS_IRD_CMD_Data") = strBAT_ID
            End If
            UpdCommon(row)
            row.EndEdit()
            Return row
        Catch ex As Exception
            Throw New Exception("Function : ClsNagra.E3Cmd 原因 : " & ex.Message)
        End Try
    End Function
    Private Function E15Cmd(ByVal aDt As DataTable) As DataRow
        Try
            Dim row As DataRow = aDt.NewRow()
            row.BeginEdit()
            row.Item("ICC_NO") = _ICCNo
            row.Item("STB_NO") = _STBNo
            
            UpdCommon(row)
            row.EndEdit()
            Return row
        Catch ex As Exception
            Throw New Exception("Function : ClsNagra.E15Cmd 原因 : " & ex.Message)
        End Try
    End Function
    Private Function E17Cmd(ByVal aDt As DataTable) As DataRow
        Try
            Dim row As DataRow = aDt.NewRow()
            row.BeginEdit()
            row.Item("ICC_NO") = _ICCNo
            row.Item("STB_NO") = _STBNo
            row.Item("DVRQuota") = _CmdString
            UpdCommon(row)
            row.EndEdit()
            Return row
        Catch ex As Exception
            Throw New Exception("Function : ClsNagra.E17Cmd 原因 : " & ex.Message)
        End Try
    End Function
    
    Private Function SplitB10Cmd() As String
        Dim strAry() As String = _CmdString.Split("~")

        Try
            For i As Int32 = LBound(strAry) To UBound(strAry)
                If strAry(i).Contains("@") Then
                    strNewProduct = SpliteA6Cmd("", strAry(i))
                Else
                    If i >= 1 AndAlso i < UBound(strAry) Then
                        strHDNo = strAry(i)
                    Else
                        strHDSpace = strAry(i)
                    End If
                End If
            Next
            Return strNewProduct
        Catch ex As Exception
            Throw New Exception("Function : ClsNagra.SplitB10Cmd 原因 : " & ex.Message)
        End Try
    End Function
    '所有命令都需要的欄位
    Private Sub UpdCommon(ByVal arow As DataRow)
        Try
            arow.Item("CMD_STATUS") = "W"
            arow.Item("OPERATOR") = strUserId
            arow.Item("UPDTIME") = Now
            arow.Item("ProcessingDate") = Now
            arow.Item("CompCode") = 99
            arow.Item("RealCompCode") = _Company
            arow.Item("SEQNO") = _Seq
            arow.Item("HIGH_LEVEL_CMD_ID") = _ProcessType
        Catch ex As Exception
            Throw New Exception("Function : ClsNagra.UpdCommon 原因 : " & ex.Message)
        End Try
    End Sub
    '取得參數的值
    Private Function GetPara() As Boolean
        Dim aSQL As String = String.Empty
        aSQL = "SELECT * From " & _OwnerName & "CAS002 " & _
            " Where CodeNo=" & _Company & " And Nvl(StopFlag,0)=0"
        Try
            Dim cmd As OracleCommand = _cn.CreateCommand()
            cmd.Transaction = _tra
            cmd.CommandText = aSQL
            Dim r As OracleDataReader = cmd.ExecuteReader
            If r.Read Then
                strBAT_ID = r.Item("BATID")
                strZIP_CODE = r.Item("ZIPCode")
                strUserId = r.Item("UserId")
            End If
            r.Close()
        Catch ex As Exception
            Throw New Exception("Function : ClsNagra.GetPara 原因 : " & ex.Message)
        End Try
    End Function
    '檢查ICCNO是否存在
    Private Function ChkICCNo() As Boolean
        Dim aSQL As String = String.Empty
        Select Case _ProcessType
            Case "B7"
                Return True
            Case "C1"
                Return True
        End Select
        Try
            aSQL = "SELECT COUNT(*) FROM " & _OwnerName & "CAS009B" & _
                " WHERE CompCode={0} AND FaciSNo='{1}' " & _
                " AND InTime > decode(outtime,null,to_date('00010101000000','yyyymmddhh24miss')) "
            aSQL = String.Format(aSQL, _Company, _ICCNo)
            Dim cmd As OracleCommand = _cn.CreateCommand()
            cmd.Transaction = _tra
            cmd.CommandText = aSQL
            If cmd.ExecuteScalar > 0 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Throw New Exception("Function : ClsNagra.ChkICCNo 原因 : " & ex.Message)
        End Try

    End Function
    '檢查STBNO是否存在
    Private Function ChkSTBNo() As Boolean
        Dim aSQL As String = String.Empty
        Select Case _ProcessType
            Case "A2", "A3", "A4"
                Return True
            Case "B1", "B2", "B3", "B4", "B5", "B6"
                Return True
            Case "C1", "C4"
                Return True
        End Select
        Try
            aSQL = "SELECT COUNT(*) FROM " & _OwnerName & "CAS009A" & _
                " WHERE CompCode={0} AND FACISNO='{1}' " & _
                " AND InTime > decode(outtime,null,to_date('00010101000000','yyyymmddhh24miss')) "
            aSQL = String.Format(aSQL, _Company, _STBNo)
            Dim cmd As OracleCommand = _cn.CreateCommand()
            If _tra.Connection IsNot Nothing Then
                cmd.Transaction = _tra
            End If

            cmd.CommandText = aSQL
            If cmd.ExecuteScalar > 0 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Throw New Exception("Function : ClsNagra.ChkSTBNo 原因 : " & ex.Message)
        End Try

    End Function
    Private Function ChkHDNo() As Boolean
        Dim aSQL As String = String.Empty
        Try
            SplitB10Cmd()
            If (_ProcessType = "B10") Then
                aSQL = "Select Count(*) From " & _OwnerName & "CAS009D " & _
                    " Where FaciSNo='" & strHDNo & "' " & _
                    " AND CompCode=" & _Company
                Dim cmd As OracleCommand = _cn.CreateCommand()
                If _tra.Connection IsNot Nothing Then
                    cmd.Transaction = _tra
                End If
                cmd.CommandText = aSQL
                If cmd.ExecuteScalar > 0 Then
                    Return True
                Else
                    Return False
                End If
            End If
            If _ProcessType = "E17" Then
                aSQL = "SELECT CodeNo FROM " & _OwnerName & "CAS001 " & _
                  " WHERE IsDVR=1"
                Dim cmd As OracleCommand = _cn.CreateCommand()
                If _tra.Connection IsNot Nothing Then
                    cmd.Transaction = _tra
                End If
                cmd.CommandText = aSQL
                Dim r As OracleDataReader = cmd.ExecuteReader
                If r.Read() Then
                    FChCode = r.Item("CodeNo")
                    Return True
                Else
                    Return False
                End If
                
            End If
            Return True
        Catch ex As Exception
            Throw New Exception("Function : ClsNagra.ChkHDNo 原因 : " & ex.Message)
        End Try
    End Function
    Private Function SpliteA6Cmd(ByVal aSplitChr As String, ByVal aSplitString As String) As String
        Dim aTmp As String = String.Empty
        strNewFaci = String.Empty
        strNewProduct = String.Empty
        strNewPinCode = String.Empty
        Dim strAry() As String
        If FNewFaci IsNot Nothing Then
            FNewFaci.Clear()
        End If
        If FNewChCode IsNot Nothing Then
            FNewChCode.Clear()
        End If
        Try
            strAry = aSplitString.Split(aSplitChr)
            For i As Int32 = LBound(strAry) To UBound(strAry)
                If strAry(i).Contains("/") Then
                    strNewFaci = strAry(i)
                    If FNewFaci Is Nothing Then
                        FNewFaci = New ArrayList()
                    End If

                    FNewFaci.Add(strAry(i))
                Else
                    If i = 0 AndAlso _ProcessType = "A9" Then
                        strNewFaci = strAry(0) & "/" & _ICCNo
                    End If
                End If
                If strAry(i).Contains("~") OrElse strAry(i).Contains("@") Then
                    Dim strPrdAry() As String = strAry(i).Split("~")
                    For j As Int32 = LBound(strPrdAry) To UBound(strPrdAry)
                        If FNewChCode Is Nothing Then
                            FNewChCode = New ArrayList()
                        End If
                        FNewChCode.Add(strPrdAry(j))
                        aTmp = "B~" & strPrdAry(j).Replace("@", "~")
                        If Not String.IsNullOrEmpty(strNewProduct) Then
                            aTmp = "," & aTmp
                        End If
                        strNewProduct = strNewProduct & aTmp
                    Next
                End If
                If _ProcessType = "A6" AndAlso i >= 2 Then
                    strNewPinCode = strAry(i)
                End If
                If _ProcessType = "B9" AndAlso i >= 1 Then
                    strNewPinCode = strAry(i)
                End If
            Next

            If aSplitChr <> "" AndAlso _ProcessType <> "B9" Then
                Return _STBNo & "/" & _ICCNo & "@" & strNewFaci & _
                    IIf(String.IsNullOrEmpty(strNewProduct), "", ";" & strNewProduct)
            Else
                Return strNewProduct
            End If

        Catch ex As Exception
            Throw New Exception("Function : ClsNagra.SpliteA1Cmd 原因 : " & ex.Message)
        End Try
    End Function
    Private Function UpdA7Cmd() As Boolean
        Dim aSQL As String = String.Empty
        aSQL = "Update " & _OwnerName & "CAS009B SET STBNo='" & _STBNo & "'," & _
            " UpdTime=To_Date('" & Format(Now, "yyyyMMddHHmmss") & "','yyyymmddhh24miss')," & _
            " UpdEn='" & strUserId & "'" & _
            " Where FaciSNo='" & _ICCNo & "' And CompCode=" & _Company
        Try
            Dim cmd As OracleCommand = _cn.CreateCommand
            If _tra.Connection IsNot Nothing Then
                cmd.Transaction = _tra
            End If
            cmd.CommandText = aSQL
            cmd.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            Throw New Exception("Function : ClsNagra.UpdA7Cmd 原因 : " & ex.Message)
        End Try
    End Function
    Private Function UpdA8Cmd() As Boolean
        Dim aSQL As String = String.Empty
        Dim aWhere As String = String.Empty
        Dim aNewFaci As String = String.Empty
        Try
            If _ProcessType = "A8" Then
                aWhere = " STBNo='" & _STBNo & "'"
            Else
                aWhere = " SmartCardNo='" & _ICCNo & "'"
            End If
            aNewFaci = _CmdString.Split(";")(0)
            aSQL = "Update " & _OwnerName & "CAS003 Set STBNo='" & aNewFaci & "'," & _
                " UpdTime=To_Date('" & Format(Now, "yyyyMMddHHmmss") & "','yyyymmddhh24miss')," & _
                " UpdEn='" & strUserId & "'" & _
                " Where CompCode=" & _Company & " And " & aWhere
            Dim cmd As OracleCommand = _cn.CreateCommand
            If _tra.Connection IsNot Nothing Then
                cmd.Transaction = _tra
            End If
            cmd.CommandText = aSQL
            cmd.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            Throw New Exception("Function : ClsNagra.UpdA8Cmd 原因 : " & ex.Message)
        End Try

    End Function
    Private Function UpdB1Cmd() As Boolean
        Dim aSQL As String = String.Empty
        Dim aNow As String = Format(Now, "yyyyMMddHHmmss")
        Dim aInTime As String = Format(_ExeTime, "yyyyMMddHHmmss")
        
        If FNewChCode Is Nothing OrElse FNewChCode.Count = 0 Then
            Throw New Exception("Function : clsNagra.UpdB1Cmd 原因 : 無任何傳送資訊")
            Return False
        End If
        Try
            Dim cmd As OracleCommand = _cn.CreateCommand
            If _tra.Connection IsNot Nothing Then
                cmd.Transaction = _tra
            End If
            For i As Int32 = 0 To FNewChCode.Count - 1
                Dim aTmp() As String
                Dim aChCode As String = String.Empty
                Dim aStartDate As String = String.Empty
                Dim aStopDate As String = String.Empty
                Dim aSEQ As String = String.Empty
                aTmp = FNewChCode.Item(i).ToString.Split("@")
                If aTmp.Length <> 3 Then
                    Throw New Exception("Function : clsNagra.UpdB1Cmd 原因 : 傳入資訊不正確")
                    Return False
                End If
                aChCode = aTmp(0)
                aStartDate = aTmp(1)
                aStopDate = aTmp(2)
                aSEQ = GetSeq("S_CAS003_SeqNo")
                aSQL = "Insert Into " & _OwnerName & "CAS003 (" & _
                        "CompCode,STBSNo,SmartCardNo,ChCode,SetTime,StartDate," & _
                        "StopDate,UpdTime,UpdEn,SEQNO) Values(" & _
                        "{0},{1},'{2}','{3}',To_Date('{4}','yyyymmddhh24miss')," & _
                        "To_Date('{5}','yyyymmdd'),To_Date('{6}','yyyymmdd')," & _
                        "To_Date('{7}','yyyymmddhh24miss'),'{8}',{9})"
                aSQL = String.Format(aSQL, _Company, "Null", _ICCNo, aChCode, _
                                     aInTime, aStartDate, aStopDate, _
                                     aInTime, strUserId, aSEQ)
                cmd.CommandText = aSQL
                cmd.ExecuteNonQuery()
            Next

            
            Return True
        Catch ex As Exception
            Throw New Exception("Function : ClsNagra.UpdB1Cmd 原因 : " & ex.Message)
        End Try
    End Function
    Private Function UpdB2Cmd() As Boolean
        Dim aSQL As String = String.Empty
        Dim aWhere As String = String.Empty
        Dim aNow As String = Format(Now, "yyyyMMddHHmmss")
        aSQL = "Update " & _OwnerName & "CAS003 SET CloseDate=To_Date('" & aNow & "','yyyymmddhh24miss') ," & _
            " UpdTime=To_Date('" & Format(_ExeTime, "yyyyMMddHHmmss") & "','yyyymmddhh24miss')," & _
            " UpdEn='" & strUserId & "' WHERE SmartCardNo ='" & _ICCNo & "'" & _
            " AND CompCode=" & _Company
        If (_ProcessType = "B2" OrElse _ProcessType = "B6") _
                AndAlso (String.IsNullOrEmpty(_CmdString)) Then
            Throw New Exception("Function : clsNagra.UpdB2Cmd 原因 : 傳入資訊不正確")
        End If
        Try
            Dim cmd As OracleCommand = _cn.CreateCommand
            If _tra.Connection IsNot Nothing Then
                cmd.Transaction = _tra
            End If
            If _ProcessType = "B2" OrElse _ProcessType = "B6" Then
                Dim aAryStr() As String = _CmdString.Split("~")
                Dim aTmp As String = String.Empty

                For i As Int32 = LBound(aAryStr) To UBound(aAryStr)
                    If _ProcessType = "B2" Then
                        If String.IsNullOrEmpty(aTmp) Then
                            aTmp = "'" & aAryStr(i) & "'"
                        Else
                            aTmp = aTmp & ",'" & aAryStr(i) & "'"
                        End If
                    Else
                        Dim aAryTmp() As String = aAryStr(i).Split("@")
                        If String.IsNullOrEmpty(aTmp) Then
                            aTmp = "'" & aAryTmp(0) & "'"
                        Else
                            aTmp = aTmp & ",'" & aAryTmp(0) & "'"
                        End If
                    End If
                Next




                If Not String.IsNullOrEmpty(aTmp) Then
                    aWhere = " AND ChCode IN (" & aTmp & ")"
                End If
            End If
            aSQL = aSQL & aWhere
            cmd.CommandText = aSQL
            cmd.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            Throw New Exception("Function : ClsNagra.UpdB2Cmd 原因 : " & ex.Message)
        End Try
    End Function
    Private Function UpdB10Cmd() As Boolean
        Dim aUpdSQL As String = String.Empty
        Dim aInsSQL As String = String.Empty
        Dim aFindSQL As String = String.Empty
        Dim aWhere As String = String.Empty
        Dim aNow As String = Format(Now, "yyyyMMddHHmmss")
        Dim aInTime As String = Format(_ExeTime, "yyyyMMddHHmmss")
        aUpdSQL = "UPDATE " & _OwnerName & "CAS009B SET DVRNo='" & strHDNo & "'" & _
                " WHERE FaciSNO='" & _ICCNo & "'"
        If FNewChCode Is Nothing OrElse FNewChCode.Count = 0 Then
            Throw New Exception("Function : clsNagra.UpdB10Cmd 原因 : 傳入資訊不正確")
            Return False
        End If

        Dim cmd As OracleCommand = _cn.CreateCommand
        If _tra.Connection IsNot Nothing Then
            cmd.Transaction = _tra
        End If
        Try
            cmd.CommandText = aUpdSQL
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            Throw New Exception("Function : ClsNagra.UpdB10Cmd(UPDATE CAS009D) 原因 : " & ex.Message)
            Return False
        End Try

        Try
            Dim aChCode As String = String.Empty
            Dim aStartDate As String = String.Empty
            Dim aStopDate As String = String.Empty
            Dim aUpdField As String = String.Empty
            If _ProcessType = "B10" Then
                Dim aTmpStr() As String = FNewChCode.Item(0).ToString.Split("@")
                If aTmpStr.Length <> 3 Then
                    Throw New Exception("Function : ClsNagra.UpdB10Cmd 原因 : 拆解資訊不正確")
                    Return False
                End If
                aChCode = aTmpStr(0)
                aStartDate = aTmpStr(1)
                aStopDate = aTmpStr(2)
                aUpdField = "StartDate=To_Date('{4}','yyyymmdd')," & _
                            "StopDate=To_Date('{5}','yyyymmdd'),DVRNo='{6}',"
                aWhere = " And DVRNo='" & strHDNo & "'"
            Else
                aChCode = FChCode
                aStartDate = "Null"
                aStopDate = "Null"
                strHDNo = "Null"
                aUpdField = "StartDate={4}," & _
                            "StopDate={5},DVRNo={6},"
            End If

            aFindSQL = "Select Count(*) From " & _OwnerName & "CAS003 " & _
                    " WHERE STBSNo='" & _STBNo & "' AND SmartCardNo='" & _ICCNo & "'" & _
                    " AND CompCode=" & _Company & " AND ChCode='" & aChCode & "'" & aWhere

            



            aUpdSQL = "UPDATE " & _OwnerName & "CAS003 SET STBSNo='{0}'," & _
                "SmartCardNo='{1}',ChCode='{2}',SetTime=To_Date('{3}','yyyymmddhh24miss')," & _
                aUpdField & _
                "UpdTime=To_Date('{7}','yyyymmddhh24miss')," & _
                "UpdEn='{8}',DVRSize={9}  Where " & _
                " SmartCardNo={10} AND CompCode={11} And ChCode='{12}'" & _
                " And STBSNo='{13}'" & aWhere
            aUpdSQL = String.Format(aUpdSQL, _STBNo, _ICCNo, aChCode, aInTime, _
                                  aStartDate, _
                                  aStopDate, _
                                   strHDNo, _
                                  aInTime, _
                                  strUserId, strHDSpace, _
                                  _ICCNo, _Company, aChCode, _STBNo)
            


            aInsSQL = "Insert Into " & _OwnerName & "CAS003 ( " & _
                "SEQNO,UpdTime,UpdEn,CompCode,SmartCardNo,ChCode, " & _
                "SetTime,StartDate,StopDate,DVRSize,DVRNo,STBSNo) Values(" & _
                "{0},To_Date('{1}','yyyymmddhh24miss'),'{2}',{3},'{4}'," & _
                "'{5}',To_Date('{6}','yyyymmddhh24miss'), " & _
                "To_Date('{7}','yyyymmddhh24miss')," & _
                "To_Date('{8}','yyyymmddhh24miss')," & _
                "{9},'{10}','{11}')"
            aInsSQL = String.Format(aInsSQL, GetSeq("S_CAS003_SeqNo"), aInTime, _
                                    strUserId, _Company, _ICCNo, aChCode, _
                                    aInTime, aStartDate, aStopDate, _
                                    strHDSpace, strHDNo, _STBNo)

            If _ProcessType = "B10" Then
                cmd.CommandText = aFindSQL
                If cmd.ExecuteScalar <= 0 Then
                    cmd.CommandText = aInsSQL
                Else
                    cmd.CommandText = aUpdSQL
                End If
            Else
                cmd.CommandText = aUpdSQL
            End If
            cmd.ExecuteNonQuery()


            Return True
        Catch ex As Exception
            Throw New Exception("Function : ClsNagra.UpdB10Cmd 原因 : " & ex.Message)
        End Try

    End Function
    Private Function GetSeq(ByVal seqName As String) As String
        Dim strRet As String = String.Empty
        Try
            Dim aSQL As String = String.Empty
            aSQL = "SELECT " & _OwnerName & seqName & ".NEXTVAL FROM DUAL "
            If _cn.State = ConnectionState.Closed Then
                _cn.Open()
            End If
            Dim cmd As OracleCommand = _cn.CreateCommand()
            If _tra.Connection IsNot Nothing Then
                cmd.Transaction = _tra
            End If
            cmd.CommandText = aSQL
            strRet = cmd.ExecuteScalar
            Return strRet
        Catch ex As Exception
            Throw New Exception("Function : ClsNagra.GetSeq 原因 : " & ex.Message)
        End Try

    End Function
    Private Function UpdA6Cmd() As Boolean
        Dim aSQL As String = String.Empty
        Dim aStrAry() As String
        If FNewFaci Is Nothing Then
            Return True
        End If

        Try
            Dim cmd As OracleCommand = _cn.CreateCommand
            If _tra.Connection IsNot Nothing Then
                cmd.Transaction = _tra
            End If
            For i As Int32 = 0 To FNewFaci.Count - 1
                aStrAry = FNewFaci.Item(i).ToString.Split("/")
                aSQL = "UPDATE " & _OwnerName & "CAS003 SET " & _
                    " STBSNo='" & aStrAry(0) & "',SmartCardNo='" & aStrAry(1) & "'," & _
                    " UpdTime=To_Date('" & Format(Now, "yyyyMMddHHmmss") & "','yyyymmddhh24miss')," & _
                    " UpdEn='" & strUserId & "' Where " & _
                    " CompCode=" & _Company & " AND STBSNO='" & _STBNo & "'" & _
                    " AND SmartCardNo='" & _ICCNo & "'"

                cmd.CommandText = aSQL
                cmd.ExecuteNonQuery()
            Next
            Return True
        Catch ex As Exception
            Throw New Exception("Function : ClsNagra.UpdA6Cmd 原因 : " & ex.Message)
        End Try
    End Function
    Private Function UpdA2Cmd() As Boolean
        Dim aSQL As String = String.Empty
        aSQL = "UPDATE " & _OwnerName & "CAS003 Set " & _
            " CloseDate=To_DATE('{0}','yyyymmdd')," & _
            " UpdTime=To_Date('{1}','yyyymmddhh24miss')," & _
            " UPDEN='{2}' Where CompCode={3} AND SmartCardNo='{4}'"
        aSQL = String.Format(aSQL, Format(Now, "yyyyMMdd"), Format(_ExeTime, "yyyyMMddhhmmss"), _
                    strUserId, _Company, _ICCNo)

        Try
            Dim cmd As OracleCommand = _cn.CreateCommand
            If _tra.Connection IsNot Nothing Then
                cmd.Transaction = _tra
            End If

            cmd.CommandText = aSQL
            cmd.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            Throw New Exception("Function : ClsNagra.UpdA2Cmd 原因 : " & ex.Message)
        End Try
    End Function

    'Private Function SplitCmd(ByVal aSplitChr As String, ByVal aFindChr As String, ByVal aCombinChr As String) As String
    '    Dim aTmp As String = String.Empty
    '    If Not String.IsNullOrEmpty(aFindChr) Then
    '        aTmp = aTmp.Substring(0, aTmp.IndexOf(aFindChr))
    '    End If
    '    Try
    '        Dim aryString() As String = _CmdString.Split(aTmp)
    '        For i As Int32 = LBound(aryString) To UBound(aryString)

    '        Next

    '    Catch ex As Exception
    '        Throw New Exception("Function : ClsNagra.SplitCmd 原因 : " & ex.Message)
    '    End Try
    'End Function
End Class
Public Enum NagraRet
    Complete = 0
    TimeOut = 1
    Nagra_RetErr = 2
    SysError = 3
    No_ICCNo = 4
    NO_STBNo = 5
    NO_HDNo = 6
    NO_Dvr = 7
End Enum