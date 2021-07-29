Imports System.Data.OracleClient
Imports CableSoft.B07.InvType

Public Class Invoice
    Implements IDisposable

    Private InvArgument As CableSoft.B07.Parameters.Argument
    Private GiveErrData As DataTable = Nothing
    Private cn As OracleClient.OracleConnection = Nothing
    Private dap As OracleDataAdapter = Nothing
    'Private _InvId1 As String = Nothing
    'Private _InvId2 As String = Nothing
    'Private _InvDate1 As String = Nothing
    'Private _InvDate2 As String = Nothing
    'Private _InvNumDate1 As String = Nothing
    'Private _InvNumDate2 As String = Nothing
    'Private _DiscountDate1 As String = Nothing
    'Private _DiscountDate2 As String = Nothing
    'Private _DiscountOv As Boolean = False
    Private FWhere As String = " A.IDENTIFYID1='1' AND A.IDENTIFYID2=0 "
    Private FOwner As String = String.Empty

    Public Property IsAllCompCode As Boolean = False
    Public Property RunInvType As InvTypeEnum.INVTYPE = InvTypeEnum.INVTYPE.NormalInvAll
    Public Property InvId1 As String
    Public Property InvId2 As String
    Public Property InvDate1 As String
    Public Property InvDate2 As String
    Public Property InvNumDate1 As String
    Public Property InvNumDate2 As String
    Public Property DiscountDate1 As String
    Public Property DiscountDate2 As String
    Public Property DiscountOv As Boolean
    Private _dtGiveErrData As DataTable

    Public Function GetAllDataToDataSet() As DataSet
        Dim dsInv As New DataSet("InvData")
        Try
            If (_dtGiveErrData IsNot Nothing) AndAlso (_dtGiveErrData.Rows.Count > 0) Then
                _dtGiveErrData.Rows.Clear()
            End If
            dsInv.Tables.Add(GetINV007)
            dsInv.Tables.Add(GetINV001)
            dsInv.Tables.Add(GetINV014)
            dsInv.Tables.Add(GetInv003)
            dsInv.Tables.Add(GetINV099)
            dsInv.Tables.Add(GetGiveErrDataTable)
        Catch ex As Exception
            Throw
        End Try
        Return dsInv
    End Function

    Public Property dtGiveErrData As DataTable
        Get
            'GetGiveErrDataTable(GetInv007SQL).Copy()
            Return GetGiveErrDataTable().Copy()
        End Get
        Set(value As DataTable)
            _dtGiveErrData = value
        End Set
    End Property
    Public Sub New(ByVal objParameters As CableSoft.B07.Parameters.Argument)
        InvArgument = objParameters
        FOwner = InvArgument.GetOwner
    End Sub
    Public Sub New(ByVal strArgument As String)
        Me.New(New CableSoft.B07.Parameters.Argument(strArgument))
        'InvArgument = New CableSoft.B07.Parameters.Argument(strArgument)
    End Sub
    Public Function GetInv003() As DataTable
        Dim aSQL As String = String.Empty
        Try
            aSQL = "SELECT * FROM " & FOwner & "INV003 A WHERE " & FWhere
            Using tbl As DataTable = OpenData(aSQL)
                tbl.TableName = "INV003"
                Return tbl.Copy()
            End Using
        Catch ex As Exception
            Throw New Exception("Function: GetINV003 訊息: " & ex.ToString)
        End Try
    End Function
    Public Function GetINV001() As DataTable
        Dim aSQL As String = String.Empty
        Try
            aSQL = "SELECT * FROM " & FOwner & "INV001 A WHERE " & FWhere
            Using tbl As DataTable = OpenData(aSQL)
                tbl.TableName = "INV001"
                Return tbl.Copy
            End Using

        Catch ex As Exception
            Throw New Exception("Function: GetINV001 訊息: " & ex.ToString)
        End Try
    End Function
    Public Function TranInvType(ByVal aInvId As List(Of String)) As Boolean

        If aInvId Is Nothing Then
            Return True
        Else
            If cn Is Nothing Then
                cn = New OracleConnection(InvArgument.GetConnect)
            End If
            If cn.State = ConnectionState.Closed Then
                cn.Open()
            End If
            Dim tra As OracleTransaction = cn.BeginTransaction
            Try
                Dim aSQL As String = "UPDATE " & FOwner & "INV007 " & _
                                " SET INVOICEKIND = 0 " & _
                                " WHERE INVID = '{0}'"
                Using cmd As OracleCommand = cn.CreateCommand
                    cmd.Transaction = tra
                    For i As Integer = 0 To aInvId.Count - 1
                        cmd.CommandText = String.Format(aSQL, aInvId.Item(i))
                        cmd.ExecuteNonQuery()
                    Next
                End Using
                tra.Commit()
                Return True
            Catch ex As Exception
                tra.Rollback()
                Throw New Exception(ex.ToString)
                Return False
            Finally
                tra.Dispose()
                If cn.State = ConnectionState.Open Then
                    cn.Close()
                End If
            End Try

        End If
    End Function
    Public Overridable Overloads Function UpdateINV() As String
        Return UpdateINV(String.Empty, UpdateInvKin.ALL)
    End Function
    Private Function GetInv007SQL() As String
        Dim aRet As String = Nothing
        Dim aSQL As String = String.Empty
        Dim aSQL2 As String = String.Empty
        Dim aWhere As String = String.Empty
        Dim aInvDateWhere As String = "1=0 "
        Try
            If InvArgument.GetStartAllUpload Then
                aWhere = " exists (select * from inv099 " & _
               " WHERE(inv099.UPLOADFLAG = 1 And substr(a.invid, 0, 2) = inv099.Prefix) " & _
               " AND TO_NUMBER(SUBSTR(A.INVID,3,8)) >= TO_NUMBER( inv099.startnum) " & _
               " AND TO_NUMBER(substr(a.invid,3,8)) <= to_number( inv099.endnum)) "
            Else
                aWhere = " A.INVOICEKIND = 1 AND A.BUSINESSID IS NULL  "
            End If
            If (Me.RunInvType = InvTypeEnum.INVTYPE.NormalInvAll) OrElse (Me.RunInvType = InvTypeEnum.INVTYPE.OnlyInvNum) OrElse
                (Me.RunInvType = InvTypeEnum.INVTYPE.NormalInvOV) Then
                If String.IsNullOrEmpty(Me.InvDate1) OrElse String.IsNullOrEmpty(Me.InvDate2) Then
                    aInvDateWhere = " 1 = 1 "
                Else
                    If String.IsNullOrEmpty(Me.InvDate1) Then
                        Me.InvDate1 = "99991231"
                    End If
                    If String.IsNullOrEmpty(Me.InvDate2) Then
                        Me.InvDate2 = "10000101"
                    End If
                    aInvDateWhere = String.Format(" INVDATE>=TO_DATE('{0}','yyyymmdd')" & _
                            " AND INVDATE<=TO_DATE('{1}','yyyymmdd') ",
                            Me.InvDate1.Replace("/", "").Replace(" ", ""),
                            Me.InvDate2.Replace("/", "").Replace(" ", ""))
                End If
            End If

            '----------------------------------------------------------------------------------------------------------------------------------------
            '正常發票
            aSQL = "SELECT A.*,B.DESCRIPTION,B.QUANTITY,B.TOTALAMOUNT,B.SEQ,B.UNITPRICE " & _
                    " FROM " & FOwner & "INV007 A, " & FOwner & "INV008 B " & _
                    " WHERE 1=1 AND " & aInvDateWhere & _
                    " AND A.INVID=B.INVID " & _
                    " AND NVL(A.UPLOADFLAG,0) = 0 " & _
                    " AND " & aWhere & _
                    " AND A.ISOBSOLETE = 'N' " & _
                    " AND A.INVFORMAT = 1 " & _
                    IIf(Me.IsAllCompCode, "", " AND COMPID='" & InvArgument.GetCompCode & "'") & " AND " & FWhere
            '#5966 增加過濾發票號碼 By Kin 2011/03/30
            If Not String.IsNullOrEmpty(Me.InvId1) Then
                aSQL = aSQL & " AND A.INVID >='" & Me.InvId1 & "' "
            End If
            If Not String.IsNullOrEmpty(Me.InvId2) Then
                aSQL = aSQL & " AND A.INVID <='" & Me.InvId2 & "' "
            End If
            '-------------------------------------------------------------------------------------------------------------------------------------------
            '-------------------------------------------------------------------------------------------------------------------------------------------
            '作廢發票
            '#6262 增加過濾OBUPLOADFLAG=0
            aSQL2 = "SELECT A.*,B.DESCRIPTION,B.QUANTITY,B.TOTALAMOUNT,B.SEQ,B.UNITPRICE " & _
                    " FROM " & FOwner & "INV007 A," & FOwner & "INV008 B " & _
                    " WHERE 1 =1 AND " & aInvDateWhere & _
                    " AND A.INVID=B.INVID " & _
                    " AND A.ISOBSOLETE = 'Y' " & _
                    " AND A.INVFORMAT = 1 " & _
                    " AND NVL(A.OBUPLOADFLAG,0) = 0 " & _
                    " AND " & aWhere & _
                    IIf(Me.IsAllCompCode, "", " AND COMPID='" & InvArgument.GetCompCode & "'") & " AND " & FWhere
            '#5966 增加過濾發票號碼 By Kin 2011/03/30
            If Not String.IsNullOrEmpty(Me.InvId1) Then
                aSQL2 = aSQL2 & " AND A.INVID >='" & Me.InvId1 & "' "
            End If
            If Not String.IsNullOrEmpty(Me.InvId2) Then
                aSQL2 = aSQL2 & " AND A.INVID <='" & Me.InvId2 & "' "
            End If
            '-------------------------------------------------------------------------------------------------------------------------------------------
            '如果前端選擇只有作廢發票，第一個SQL不要找資料(正常的發票)

            If Me.RunInvType = InvTypeEnum.INVTYPE.NormalInvOV Then
                aSQL = aSQL & " AND 1=0 "
            End If
            aSQL = aSQL & " UNION ALL " & aSQL2
            aSQL = "SELECT Distinct A.*,Nvl(B.REFNO,0) REFNO,C.BUSINESSID BUSINESSID2 " & _
                " FROM (" & aSQL & ") A," & FOwner & "INV028 B," & FOwner & "INV041 C " & _
                " WHERE A.INVUSEID=B.ITEMID(+) " & IIf(Me.IsAllCompCode, "", " AND B.COMPID(+)='" & InvArgument.GetCompCode & "'") & _
                " AND A.GIVEUNITID = C.ITEMID(+) AND A.COMPID = C.COMPID(+)  " & _
                " Order By A.INVID,A.SEQ "
            aRet = aSQL
            'aRet = "SELECT DISTINCT INVID FROM (" & aSQL & ") "
        Catch ex As Exception
            Throw

        End Try
        Return aRet
    End Function
    Private Function GetGiveINV007DataSQL() As String
        Dim aRet As String = Nothing
        Dim Inv007SQL As String = GetInv007SQL()
        Inv007SQL = "SELECT INVID FROM (" & Inv007SQL & ")"
        Try
            aRet = " SELECT INV007.INVID,NULL GIVEUNITDESC FROM INV007,INV028 " & _
                                " WHERE INV007.INVUSEID IS NOT NULL " & _
                                " AND INV007.INVUSEID = INV028.ITEMID " & _
                                " AND INV028.REFNO = 1 " & _
                                " AND INV007.GIVEUNITID IS NULL " & _
                                " AND INV007.LOVENUM IS NULL " & _
                                " AND INV007.INVID IN (" & Inv007SQL & ")" & _
                                " UNION " & _
                                 "SELECT INV007.INVID,INV041.DESCRIPTION GIVEUNITDESC FROM INV007,INV028,INV041 " & _
                                " WHERE INV007.INVUSEID IS NOT NULL " & _
                                " AND INV007.INVUSEID = INV028.ITEMID " & _
                                " AND INV028.REFNO = 1 " & _
                                " AND INV007.GIVEUNITID IS NOT  NULL " & _
                                " AND INV041.ITEMID = INV007.GIVEUNITID " & _
                                " AND INV041.BUSINESSID IS NULL " & _
                                " AND INV007.INVID IN (" & Inv007SQL & ")"
        Catch ex As Exception
            Throw
        End Try
        Return aRet
    End Function
    Public Function GetGiveErrDataTable() As DataTable
        Dim aSQL As String = Nothing
        'Dim tbReturn As DataTable = Nothing
        If _dtGiveErrData IsNot Nothing AndAlso _dtGiveErrData.Rows.Count > 0 Then
            _dtGiveErrData.Rows.Clear()
        End If
        For i As Integer = 0 To 1
            If i = 0 Then
                aSQL = GetGiveINV007DataSQL()
            Else
                aSQL = GetGiveINV014DataSQL()
            End If
            If _dtGiveErrData Is Nothing Then
                _dtGiveErrData = New DataTable("GiveErrData")
                _dtGiveErrData.Columns.Add("InvId", GetType(String))
                _dtGiveErrData.Columns.Add("Msg", GetType(String))
            End If
            Try
                Using tbExec As DataTable = OpenData(aSQL).Copy
                    If (tbExec IsNot Nothing) AndAlso
                        (tbExec.Rows.Count > 0) Then
                        If _dtGiveErrData Is Nothing Then
                            _dtGiveErrData = New DataTable("GiveErrData")
                            _dtGiveErrData.Columns.Add("InvId", GetType(String))
                            _dtGiveErrData.Columns.Add("Msg", GetType(String))
                        End If
                        For Each rw As DataRow In tbExec.Rows
                            Dim rwNew As DataRow = _dtGiveErrData.NewRow
                            rwNew.Item("InvId") = rw.Item("InvId")
                            If DBNull.Value.Equals(rw.Item("GIVEUNITDESC")) Then
                                rwNew.Item("Msg") = String.Format("發票號碼：[ {0} ] 捐贈單位無設定！", rw.Item("InvId"))
                            Else
                                rwNew.Item("Msg") = String.Format("發票號碼：[ {0} ] 捐贈單位：[ {1} ] 無設定統一編號！",
                                                                  rw.Item("InvId"), rw.Item("GIVEUNITDESC"))
                            End If
                            _dtGiveErrData.Rows.Add(rwNew)
                        Next
                        _dtGiveErrData.AcceptChanges()
                    End If
                End Using

            Catch ex As Exception
                Throw
            Finally
                'If tbReturn IsNot Nothing Then
                '    tbReturn.Dispose()
                'End If
            End Try
        Next       
        Return _dtGiveErrData.Copy
    End Function
    Public Function GetINV007() As DataTable
        Try
            Dim aSQL As String = GetInv007SQL()
            Using tbl As DataTable = OpenData(aSQL)
                tbl.TableName = "INV007"
                Return tbl.Copy
            End Using
        Catch ex As Exception
            Throw
        End Try
    End Function
    Private Function GetINV099SQL() As String
        Dim aSQL As String = String.Empty
        Dim aWhere As String = String.Empty        
        Try
            aSQL = "SELECT * FROM " & FOwner & "INV099 "
            If RunInvType <> InvTypeEnum.INVTYPE.OnlyNotUseInvNum Then
                aSQL = aSQL & " WHERE 1 =0 "
                Return aSQL
            End If

            If (String.IsNullOrEmpty(Me.InvNumDate1)) AndAlso (String.IsNullOrEmpty(Me.InvNumDate2)) Then
                aWhere = " WHERE 1= 0 "
            Else
                aWhere = " WHERE NVL(UPLOADFLAG,0)=1 AND CURNUM<=ENDNUM "
                If Not String.IsNullOrEmpty(Me.InvNumDate1) Then
                    aWhere = String.Format(" {0} AND YEARMONTH>={1}", aWhere, Me.InvNumDate1.Replace("/", "").Replace(" ", ""))
                End If
                If Not String.IsNullOrEmpty(Me.InvNumDate2) Then
                    aWhere = String.Format(" {0} AND YEARMONTH<={1}", aWhere, Me.InvNumDate2.Replace("/", "").Replace(" ", ""))
                End If
                If Not Me.IsAllCompCode Then
                    aWhere = String.Format(" {0} AND COMPID = '{1}' ", aWhere, InvArgument.GetCompCode)
                End If
            End If
            aSQL = aSQL & aWhere
        Catch ex As Exception
            Throw
        End Try
        Return aSQL
    End Function
    Public Function GetINV099() As DataTable
        Try
            Dim aSQL As String = GetINV099SQL()
            Using tbl As DataTable = OpenData(aSQL)
                tbl.TableName = "INV099"
                Return tbl.Copy
            End Using
        Catch ex As Exception
            Throw
        End Try        
    End Function
    Private Function GetGiveINV014DataSQL() As String
        Dim aINV014SQL As String = GetINV014SQL()
        aINV014SQL = "SELECT ALLOWANCENO FROM (" & aINV014SQL & ") "
        '        GetGiveINV007DataSQL()
        Dim aGiveSQL As String = "SELECT INV007.INVID,NULL GIVEUNITDESC FROM INV007,INV028 " & _
                             " WHERE INV007.INVUSEID IS NOT NULL " & _
                             " AND INV007.INVUSEID = INV028.ITEMID " & _
                             " AND INV028.REFNO = 1 " & _
                             " AND INV007.GIVEUNITID IS NULL " & _
                             " AND INV007.INVID IN (SELECT INVID FROM INV014 " & _
                                          " WHERE INVID IS NOT NULL AND ALLOWANCENO IN (" & aINV014SQL & "))" & _
                             " UNION ALL " & _
                              "SELECT INV007.INVID,INV041.DESCRIPTION GIVEUNITDESC FROM INV007,INV028,INV041 " & _
                             " WHERE INV007.INVUSEID IS NOT NULL " & _
                             " AND INV007.INVUSEID = INV028.ITEMID " & _
                             " AND INV028.REFNO = 1 " & _
                             " AND INV007.GIVEUNITID IS NOT  NULL " & _
                             " AND INV041.ITEMID = INV028.ITEMID " & _
                             " AND INV041.BUSINESSID IS NULL " & _
                             " AND INV007.INVID IN (SELECT INVID FROM INV014 " & _
                                          " WHERE INVID IS NOT NULL AND ALLOWANCENO IN (" & aINV014SQL & "))"
        Return aGiveSQL
    End Function
    Private Function GetINV014SQL() As String
        '折讓改成永遠不產生
        '#5881 將產生折讓恢復 By Kin 2010/12/14
        '#5922 增加INVFORMAT =1 條件 By Kin 2010/02/14
        Dim aSQL As String = String.Empty
        Dim aField As String = "A.PaperNo,A.PaperDate,A.InvDate, " & _
                        "A.InvID,A.SaleAmount,A.TaxAmount, " & _
                        "A.TaxType,A.COMPID,A.AllowanceNo,A.ObsoleteReason,B.BusinessId  BusinessId2,B.InvTitle "
        Dim aGroupField As String = "A.PaperNo,A.PaperDate,A.InvDate, " & _
                        "A.InvID,A.SaleAmount,A.TaxAmount, " & _
                        "A.TaxType,A.COMPID,A.AllowanceNo,A.ObsoleteReason,A.BusinessId2,A.InvTitle"
        Dim aAllUplaodWhere As String = String.Empty
       
        Try
            If InvArgument.GetStartAllUpload Then
                aAllUplaodWhere = " 1 = 1"
            Else
                aAllUplaodWhere = " B.INVOICEKIND = 1 AND A.BUSINESSID IS NULL  "
            End If
            If String.IsNullOrEmpty(Me.DiscountDate1) Then
                Me.DiscountDate1 = "99991231"
            End If
            If String.IsNullOrEmpty(_DiscountDate2) Then
                Me.DiscountDate2 = "10000101"
            End If
            aSQL = "SELECT " & aField & _
                    " FROM " & FOwner & "INV014 A,INV007 B " & _
                    " WHERE A.PaperDate>=TO_DATE('" & Me.DiscountDate1.Replace("/", "").Replace(" ", "") & "','yyyymmdd')" & _
                    " AND A.PaperDate<=TO_DATE('" & Me.DiscountDate2.Replace("/", "").Replace(" ", "") & "','yyyymmdd')" & _
                    IIf(Me.IsAllCompCode, "", " AND A.COMPID='" & InvArgument.GetCompCode & "'") & _
                    " AND A.INVID = B.INVID " & _
                    " AND B.UPLOADFLAG =1 AND B.INVFORMAT = 1 " & _
                    " AND " & aAllUplaodWhere
            '#5966 增加過濾發票號碼 By Kin 2011/03/30
            If Not String.IsNullOrEmpty(Me.InvId1) Then
                aSQL = aSQL & " AND B.INVID >='" & Me.InvId1 & "' "
            End If
            If Not String.IsNullOrEmpty(Me.InvId2) Then
                aSQL = aSQL & " AND B.INVID <='" & Me.InvId2 & "' "
            End If
            '#6001 增加折讓是否作廢，如果選擇作廢不要撈取資料 By Kin 2011/07/13
            'If Me.DiscountOv Then
            '    aSQL = aSQL & " AND  A.ISOBSOLETE = 'Y' AND Nvl(A.UPLOADFLAG,0)=1 "
            'Else
            '    aSQL = aSQL & " AND A.ISOBSOLETE = 'N' AND Nvl(A.UPLOADFLAG,0)=0  "
            'End If
            If Me.RunInvType = InvTypeEnum.INVTYPE.OnlyDiscountOV Then
                aSQL = aSQL & " AND  A.ISOBSOLETE = 'Y' AND Nvl(A.UPLOADFLAG,0)=1 "
            ElseIf Me.RunInvType = InvTypeEnum.INVTYPE.OnlyDiscount Then
                aSQL = aSQL & " AND A.ISOBSOLETE = 'N' AND Nvl(A.UPLOADFLAG,0)=0  "
            Else
                aSQL = aSQL & " AND 1 = 0  "
            End If
            
            '#6001 撈取資料增加 UPTTIME For B0401 By Kin 2011/07/13
            aSQL = "SELECT Distinct A.*,Min(B.SEQ) SEQ,MAX(C.UPTTIME) UPTTIME " & _
                " FROM (" & aSQL & ") A,INV008 B,INV014 C " & _
                " WHERE A.INVID=B.INVID AND B.UNITPRICE > 0 " & _
                " AND A.ALLOWANCENO = C.ALLOWANCENO AND A.COMPID = C.COMPID " & _
                " GROUP BY " & aGroupField & _
                " ORDER BY A.INVID "

            '#6207 不同折讓單同一發票只要抓其中一筆明細就好 By Kin 2012/01/11
            aSQL = "SELECT DISTINCT A.*,B.DESCRIPTION," & _
                    " B.QUANTITY FROM (" & aSQL & ") A,INV008 B " & _
                    " WHERE A.INVID=B.INVID AND A.SEQ=B.SEQ AND B.IDENTIFYID1='1' " & _
                    " AND B.IDENTIFYID2=0 "

            'aSQL = "SELECT ALLOWANCENO FROM (" & aSQL & ") "
        Catch ex As Exception
            Throw
        End Try
        Return aSQL
    End Function
    Public Function GetINV014() As DataTable
        Try
            Dim aSQL As String = GetINV014SQL()
            Using tbl As DataTable = OpenData(aSQL)
                tbl.TableName = "INV014"
                Return tbl.Copy
            End Using            
        Catch ex As Exception
            Throw
        End Try        
    End Function
    'Public Function GetINV014() As DataTable
    '    '折讓改成永遠不產生
    '    '#5881 將產生折讓恢復 By Kin 2010/12/14
    '    '#5922 增加INVFORMAT =1 條件 By Kin 2010/02/14
    '    Dim aSQL As String = String.Empty
    '    Dim aField As String = "A.PaperNo,A.PaperDate,A.InvDate, " & _
    '                    "A.InvID,A.SaleAmount,A.TaxAmount, " & _
    '                    "A.TaxType,A.COMPID,A.AllowanceNo,A.ObsoleteReason,B.BusinessId  BusinessId2,B.InvTitle "
    '    Dim aGroupField As String = "A.PaperNo,A.PaperDate,A.InvDate, " & _
    '                    "A.InvID,A.SaleAmount,A.TaxAmount, " & _
    '                    "A.TaxType,A.COMPID,A.AllowanceNo,A.ObsoleteReason,A.BusinessId2,A.InvTitle"
    '    Dim aAllUplaodWhere As String = String.Empty
    '    '#6600 測試不ok, 若有啟動(StartAllUpload=1), 若針對折讓資料異動, 需"取消"INVOICEKIND=1及BUSINESSID IS NULL的條件 By Kin 2013/11/11
    '    If StartAllUpload Then
    '        'aAllUplaodWhere = " exists (select * from inv099 " & _
    '        '    " WHERE(inv099.UPLOADFLAG = 1 And substr(a.invid, 0, 2) = inv099.Prefix) " & _
    '        '    " AND TO_NUMBER(SUBSTR(A.INVID,3,8)) >= TO_NUMBER( inv099.startnum) " & _
    '        '    " AND TO_NUMBER(substr(a.invid,3,8)) <= to_number( inv099.endnum)) "
    '        aAllUplaodWhere = " 1 = 1"
    '    Else
    '        aAllUplaodWhere = " B.INVOICEKIND = 1 AND A.BUSINESSID IS NULL  "
    '    End If
    '    Try
    '        aSQL = "SELECT " & aField & _
    '                " FROM " & FOwner & "INV014 A,INV007 B " & _
    '                " WHERE A.PaperDate>=TO_DATE('" & _DiscountDate1 & "','yyyymmdd')" & _
    '                " AND A.PaperDate<=TO_DATE('" & _DiscountDate2 & "','yyyymmdd')" & _
    '                IIf(FUseAllComp, "", " AND A.COMPID='" & FCompCode & "'") & _
    '                " AND A.INVID = B.INVID " & _
    '                " AND B.UPLOADFLAG =1 AND B.INVFORMAT = 1 " & _
    '                " AND " & aAllUplaodWhere
    '        '#5966 增加過濾發票號碼 By Kin 2011/03/30
    '        If Not String.IsNullOrEmpty(_InvId1) Then
    '            aSQL = aSQL & " AND B.INVID >='" & _InvId1 & "' "
    '        End If
    '        If Not String.IsNullOrEmpty(_InvId2) Then
    '            aSQL = aSQL & " AND B.INVID <='" & _InvId2 & "' "
    '        End If
    '        '#6001 增加折讓是否作廢，如果選擇作廢不要撈取資料 By Kin 2011/07/13
    '        If _DiscountOv Then
    '            aSQL = aSQL & " AND  A.ISOBSOLETE = 'Y' AND Nvl(A.UPLOADFLAG,0)=1 "
    '        Else
    '            aSQL = aSQL & " AND A.ISOBSOLETE = 'N' AND Nvl(A.UPLOADFLAG,0)=0  "
    '        End If

    '        '#6001 撈取資料增加 UPTTIME For B0401 By Kin 2011/07/13
    '        aSQL = "SELECT Distinct A.*,Min(B.SEQ) SEQ,MAX(C.UPTTIME) UPTTIME " & _
    '            " FROM (" & aSQL & ") A,INV008 B,INV014 C " & _
    '            " WHERE A.INVID=B.INVID AND B.UNITPRICE > 0 " & _
    '            " AND A.ALLOWANCENO = C.ALLOWANCENO AND A.COMPID = C.COMPID " & _
    '            " GROUP BY " & aGroupField & _
    '            " ORDER BY A.INVID "

    '        '#6207 不同折讓單同一發票只要抓其中一筆明細就好 By Kin 2012/01/11
    '        aSQL = "SELECT DISTINCT A.*,B.DESCRIPTION," & _
    '                " B.QUANTITY FROM (" & aSQL & ") A,INV008 B " & _
    '                " WHERE A.INVID=B.INVID AND A.SEQ=B.SEQ AND B.IDENTIFYID1='1' " & _
    '                " AND B.IDENTIFYID2=0 "

    '        fInv014IdSQL = "SELECT ALLOWANCENO FROM (" & aSQL & ") "

    '        Dim aGiveSQL As String = "SELECT INV007.INVID,NULL GIVEUNITDESC FROM INV007,INV028 " & _
    '                           " WHERE INV007.INVUSEID IS NOT NULL " & _
    '                           " AND INV007.INVUSEID = INV028.ITEMID " & _
    '                           " AND INV028.REFNO = 1 " & _
    '                           " AND INV007.GIVEUNITID IS NULL " & _
    '                           " AND INV007.INVID IN (SELECT INVID FROM INV014 " & _
    '                                        " WHERE INVID IS NOT NULL AND ALLOWANCENO IN (" & fInv014IdSQL & "))" & _
    '                           " UNION ALL " & _
    '                            "SELECT INV007.INVID,INV041.DESCRIPTION GIVEUNITDESC FROM INV007,INV028,INV041 " & _
    '                           " WHERE INV007.INVUSEID IS NOT NULL " & _
    '                           " AND INV007.INVUSEID = INV028.ITEMID " & _
    '                           " AND INV028.REFNO = 1 " & _
    '                           " AND INV007.GIVEUNITID IS NOT  NULL " & _
    '                           " AND INV041.ITEMID = INV028.ITEMID " & _
    '                           " AND INV041.BUSINESSID IS NULL " & _
    '                           " AND INV007.INVID IN (SELECT INVID FROM INV014 " & _
    '                                        " WHERE INVID IS NOT NULL AND ALLOWANCENO IN (" & fInv014IdSQL & "))"

    '        ChkGiveData(aGiveSQL)
    '        Dim tbl As DataTable = OpenData(aSQL).Copy
    '        tbl.TableName = "INV014"
    '        Return tbl
    '    Catch ex As Exception
    '        Throw New Exception("Function: GetINV014 訊息: " & ex.ToString)
    '    End Try
    'End Function


    'Public Function GetINV099() As DataTable
    '    Dim aSQL As String = String.Empty
    '    Dim aWhere As String = String.Empty
    '    Try
    '        aSQL = "SELECT * FROM " & FOwner & "INV099 "

    '        If (String.IsNullOrEmpty(_InvNumDate1)) AndAlso (String.IsNullOrEmpty(_InvNumDate2)) Then
    '            aWhere = " WHERE 1= 0 "
    '        Else
    '            aWhere = " WHERE NVL(UPLOADFLAG,0)=1 AND CURNUM<=ENDNUM "
    '            If Not String.IsNullOrEmpty(_InvNumDate1) Then
    '                aWhere = String.Format(" {0} AND YEARMONTH>={1}", aWhere, _InvNumDate1)
    '            End If
    '            If Not String.IsNullOrEmpty(_InvNumDate2) Then
    '                aWhere = String.Format(" {0} AND YEARMONTH<={1}", aWhere, _InvNumDate2)
    '            End If
    '            If Not FUseAllComp Then
    '                aWhere = String.Format(" {0} AND COMPID = '{1}' ", aWhere, FCompCode)
    '            End If
    '        End If
    '        aSQL = aSQL & aWhere
    '        Dim tbl As DataTable = OpenData(aSQL).Copy
    '        tbl.TableName = "INV099"
    '        Return tbl
    '    Catch ex As Exception
    '        Throw New Exception("Function: GetINV099 訊息: " & ex.ToString)
    '    End Try
    'End Function


    'Public Function GetINV007() As DataTable
    '    Dim aSQL As String = String.Empty
    '    Dim aSQL2 As String = String.Empty
    '    Dim aWhere As String = String.Empty
    '    '#5922 增加Invformat = 1 條件 By Kin 2010/02/14
    '    '#6370 增加判斷StartAllUpload By Kin 2012/11/23
    '    '#6370 StartAllUpload = true 取消INV007.INVOICEKIND=1，及取消INV007.BUSINESSID IS NULL的條件 By Kin 2012/12/05

    '    If StartAllUpload Then
    '        aWhere = " exists (select * from inv099 " & _
    '            " WHERE(inv099.UPLOADFLAG = 1 And substr(a.invid, 0, 2) = inv099.Prefix) " & _
    '            " AND TO_NUMBER(SUBSTR(A.INVID,3,8)) >= TO_NUMBER( inv099.startnum) " & _
    '            " AND TO_NUMBER(substr(a.invid,3,8)) <= to_number( inv099.endnum)) "
    '    Else
    '        aWhere = " A.INVOICEKIND = 1 AND A.BUSINESSID IS NULL  "
    '    End If
    '    Try
    '        '正常發票
    '        aSQL = "SELECT A.*,B.DESCRIPTION,B.QUANTITY,B.TOTALAMOUNT,B.SEQ,B.UNITPRICE " & _
    '                " FROM " & FOwner & "INV007 A, " & FOwner & "INV008 B " & _
    '                " WHERE INVDATE>=TO_DATE('" & _InvDate1 & "','yyyymmdd')" & _
    '                " AND INVDATE<=TO_DATE('" & _InvDate2 & "','yyyymmdd') " & _
    '                " AND A.INVID=B.INVID " & _
    '                " AND NVL(A.UPLOADFLAG,0) = 0 " & _
    '                " AND " & aWhere & _
    '                " AND A.ISOBSOLETE = 'N' " & _
    '                " AND A.INVFORMAT = 1 " & _
    '                IIf(FUseAllComp, "", " AND COMPID='" & FCompCode & "'") & " AND " & FWhere
    '        '#5966 增加過濾發票號碼 By Kin 2011/03/30
    '        If Not String.IsNullOrEmpty(_InvId1) Then
    '            aSQL = aSQL & " AND A.INVID >='" & _InvId1 & "' "
    '        End If
    '        If Not String.IsNullOrEmpty(_InvId2) Then
    '            aSQL = aSQL & " AND A.INVID <='" & _InvId2 & "' "
    '        End If

    '        '作廢發票
    '        '#6262 增加過濾OBUPLOADFLAG=0
    '        aSQL2 = "SELECT A.*,B.DESCRIPTION,B.QUANTITY,B.TOTALAMOUNT,B.SEQ,B.UNITPRICE " & _
    '                " FROM " & FOwner & "INV007 A," & FOwner & "INV008 B " & _
    '                " WHERE INVDATE>=TO_DATE('" & _InvDate1 & "','yyyymmdd')" & _
    '                " AND INVDATE<=TO_DATE('" & _InvDate2 & "','yyyymmdd') " & _
    '                " AND A.INVID=B.INVID " & _
    '                " AND A.ISOBSOLETE = 'Y' " & _
    '                " AND A.INVFORMAT = 1 " & _
    '                " AND NVL(A.OBUPLOADFLAG,0) = 0 " & _
    '                " AND " & aWhere & _
    '                IIf(FUseAllComp, "", " AND COMPID='" & FCompCode & "'") & " AND " & FWhere

    '        '#5966 增加過濾發票號碼 By Kin 2011/03/30
    '        If Not String.IsNullOrEmpty(_InvId1) Then
    '            aSQL2 = aSQL2 & " AND A.INVID >='" & _InvId1 & "' "
    '        End If
    '        If Not String.IsNullOrEmpty(_InvId2) Then
    '            aSQL2 = aSQL2 & " AND A.INVID <='" & _InvId2 & "' "
    '        End If

    '        '如果前端選擇只有作廢發票，第一個SQL不要找資料(正常的發票)
    '        If FINVType = INVTYPE.OV Then
    '            aSQL = aSQL & " AND 1=0 "
    '        End If
    '        aSQL = aSQL & " UNION ALL " & aSQL2
    '        aSQL = "SELECT Distinct A.*,Nvl(B.REFNO,0) REFNO,C.BUSINESSID BUSINESSID2 " & _
    '            " FROM (" & aSQL & ") A," & FOwner & "INV028 B," & FOwner & "INV041 C " & _
    '            " WHERE A.INVUSEID=B.ITEMID(+) " & IIf(FUseAllComp, "", " AND B.COMPID(+)='" & FCompCode & "'") & _
    '            " AND A.GIVEUNITID = C.ITEMID(+) AND A.COMPID = C.COMPID(+)  " & _
    '            " Order By A.INVID,A.SEQ "
    '        fInv007InvidSQL = "SELECT DISTINCT INVID FROM (" & aSQL & ") "
    '        '#6600 GIVEUNITID 與 LOVENUM 都是NULL才算錯誤 By Kin 2013/10/02
    '        Dim aGiveSQL As String = "SELECT INV007.INVID,NULL GIVEUNITDESC FROM INV007,INV028 " & _
    '                            " WHERE INV007.INVUSEID IS NOT NULL " & _
    '                            " AND INV007.INVUSEID = INV028.ITEMID " & _
    '                            " AND INV028.REFNO = 1 " & _
    '                            " AND INV007.GIVEUNITID IS NULL " & _
    '                            " AND INV007.LOVENUM IS NULL " & _
    '                            " AND INV007.INVID IN (" & fInv007InvidSQL & ")" & _
    '                            " UNION " & _
    '                             "SELECT INV007.INVID,INV041.DESCRIPTION GIVEUNITDESC FROM INV007,INV028,INV041 " & _
    '                            " WHERE INV007.INVUSEID IS NOT NULL " & _
    '                            " AND INV007.INVUSEID = INV028.ITEMID " & _
    '                            " AND INV028.REFNO = 1 " & _
    '                            " AND INV007.GIVEUNITID IS NOT  NULL " & _
    '                            " AND INV041.ITEMID = INV007.GIVEUNITID " & _
    '                            " AND INV041.BUSINESSID IS NULL " & _
    '                            " AND INV007.INVID IN (" & fInv007InvidSQL & ")"

    '        ChkGiveData(aGiveSQL)
    '        Dim tbl As DataTable = OpenData(aSQL).Copy
    '        tbl.TableName = "INV007"
    '        Return tbl
    '    Catch ex As Exception
    '        Throw New Exception("Function: GetINV007 訊息: " & ex.ToString)
    '    End Try
    'End Function
    Public Overloads Function UpdateINV(ByVal Kind As UpdateInvKin) As String
        Return UpdateINV(String.Empty, Kind)
    End Function

    Public Overridable Overloads Function UpdateINV(ByVal aDate As String, ByVal Kind As UpdateInvKin) As String
        Dim aSQL As String = String.Empty
        Dim aUpdTime As String = String.Empty
        If (String.IsNullOrEmpty(aDate)) OrElse (Not IsDate(aDate)) Then
            aUpdTime = Format(Now, "yyyyMMddHHmmss")
        Else
            aUpdTime = Format(Date.Parse(aDate), "yyyyMMddHHmmss")
        End If

        If cn Is Nothing Then
            cn = New OracleConnection(InvArgument.GetConnect)
        End If
        If cn.State = ConnectionState.Closed Then
            cn.Open()
        End If

        '#5945 增加上傳檔案時間 By Kin 2011/03/17
        Dim tra As OracleTransaction = cn.BeginTransaction

        Try
            Dim aInv007InvId As String = String.Format("SELECT DISTINCT INVID FROM ({0})", GetInv007SQL)
            Dim aInv014Invid As String = String.Format("SELECT DISTINCT INVID FROM ({0})", GetINV014SQL)
            Using cmd As OracleCommand = cn.CreateCommand
                '直接從撈取INV007的SQL語法 UPD ，避免UPD到別筆資料
                '#6262 增加作廢上傳註記，所以正常發票與作廢發票要分開做 By Kin 2012/06/21
                aSQL = "UPDATE " & FOwner & "INV007 A SET A.UPLOADFLAG = 1, " & _
                        " A.UPLOADTIME = TO_DATE('" & aUpdTime & "','YYYYMMDDHH24MISS') " & _
                        " WHERE IsObsolete <> 'Y' AND INVID IN (" & aInv007InvId & ") "
                If Kind = UpdateInvKin.ALL OrElse Kind = UpdateInvKin.INV007 Then
                    cmd.Transaction = tra
                    cmd.CommandText = aSQL
                    cmd.ExecuteNonQuery()
                    aSQL = "UPDATE " & FOwner & "INV007 A SET A.OBUPLOADFLAG = 1, " & _
                      " A.OBUPLOADTIME = TO_DATE('" & aUpdTime & "','YYYYMMDDHH24MISS') " & _
                      " WHERE IsObsolete = 'Y' AND INVID IN (" & aInv007InvId & ") "
                    cmd.CommandText = aSQL
                    cmd.ExecuteNonQuery()
                End If
             
                '直接從撈取INV014的SQL語法 UPD ，避免UPD到別筆資料
                aSQL = "UPDATE " & FOwner & "INV014 SET UPLOADFLAG = 1, " & _
                    " UPLOADTIME = TO_DATE('" & aUpdTime & "','YYYYMMDDHH24MISS') " & _
                    " WHERE ALLOWANCENO IN (" & aInv014Invid & ") " & _
                    IIf(Me.IsAllCompCode, "", " AND COMPID='" & InvArgument.GetCompCode & "'")
                If Kind = UpdateInvKin.ALL OrElse Kind = UpdateInvKin.INV014 Then
                    cmd.CommandText = aSQL
                    cmd.ExecuteNonQuery()
                End If
               
                tra.Commit()
            End Using
        Catch ex As Exception
            tra.Rollback()
            Throw New Exception(ex.ToString)
        Finally
            tra.Dispose()
            If cn.State = ConnectionState.Open Then
                cn.Close()
                cn = Nothing
            End If
        End Try
        Return aUpdTime
    End Function
    Private Function OpenData(ByVal aSQL As String) As DataTable
        Try
            If cn Is Nothing Then
                cn = New OracleConnection(InvArgument.GetConnect)
            End If
            If dap Is Nothing Then
                dap = New OracleDataAdapter(aSQL, cn)
            End If
            Using ds As New DataSet
                dap.Fill(ds)
                Return ds.Tables(0)
            End Using
            
        Catch ex As Exception
            Throw New Exception("Function: OpenData 訊息: " & ex.ToString)
        Finally
            If cn.State = ConnectionState.Open Then
                cn.Close()
            End If
            cn.Dispose()
            cn = Nothing
            If dap IsNot Nothing Then
                dap.Dispose()
            End If
            dap = Nothing
        End Try

    End Function
    Public Shared Function GetInvParam(ByVal aXML As String, ByVal aAttrName As String) As String
        Dim aXmlDoc As New System.Xml.XmlDocument()
        Dim aRet As String = String.Empty
        aXmlDoc.LoadXml(aXML)
        Dim aXmlAttr As Xml.XmlAttribute = aXmlDoc.DocumentElement.Attributes(aAttrName)
        If aXmlAttr Is Nothing Then
            Return String.Empty
        End If
        aRet = aXmlAttr.Value
        Return aRet
    End Function

    Public Shared Function GetInvParamXML(ByVal ConnectString As String, ByVal aCompCode As String) As String
        Dim aSQL As String = String.Empty
        Dim aXML As String = String.Empty
        aSQL = "SELECT INVPARAM FROM INV003 WHERE COMPID='" & aCompCode & "'"
        Dim cn2 As New OracleConnection(ConnectString)
        Dim cmd As New OracleCommand(aSQL, cn2)
        Try
            If cn2.State = ConnectionState.Closed Then
                cn2.Open()
            End If
            Return cmd.ExecuteOracleScalar().ToString
            'If String.IsNullOrEmpty(aXML) Then Return String.Empty
            'Dim aXmlDoc As New System.Xml.XmlDocument()

        Catch ex As Exception
            Throw ex
        Finally
            If cmd IsNot Nothing Then
                cmd.Dispose()
                cmd = Nothing
            End If
            If cn2 IsNot Nothing Then
                cn2.Close()
                cn2.Dispose()
                cn2 = Nothing
            End If
        End Try
       


    End Function
    Public Shared Function GetINVPathName(ByVal ConnectString As String, ByVal aCompCode As String,
                                          ByVal aAttrName As String) As String
        Dim aSQL As String = String.Empty
        Dim aXML As String = String.Empty
        aSQL = "SELECT INVPARAM FROM INV003 WHERE COMPID='" & aCompCode & "'"
        Dim cn2 As New OracleConnection(ConnectString)
        Dim cmd As New OracleCommand(aSQL, cn2)
        Try
            If cn2.State = ConnectionState.Closed Then
                cn2.Open()
            End If
            aXML = cmd.ExecuteOracleScalar().ToString
            If String.IsNullOrEmpty(aXML) Then Return String.Empty
            Dim aXmlDoc As New System.Xml.XmlDocument()
            aXmlDoc.LoadXml(aXML)
            Dim aXmlAttr As Xml.XmlAttribute = aXmlDoc.DocumentElement.Attributes(aAttrName)
            If aXmlAttr Is Nothing Then
                Return String.Empty
            End If
            Dim aRet As String = aXmlAttr.Value
            Return aRet
        Catch ex As Exception
            Return String.Empty
        Finally
            If cmd IsNot Nothing Then
                cmd.Dispose()
                cmd = Nothing
            End If
            If cn2 IsNot Nothing Then
                cn2.Close()
                cn2.Dispose()
                cn2 = Nothing
            End If

        End Try
    End Function
    Public Function GetINVPath() As String
        Try
            Dim aSQL As String = String.Empty
            Dim aXML As String = String.Empty
            aSQL = "SELECT INVPARAM FROM " & FOwner & "INV003 WHERE COMPID='" & InvArgument.GetCompCode & "'"
            Using tbl As DataTable = OpenData(aSQL)
                If tbl.Rows.Count <= 0 Then
                    Throw New Exception("Function: GetINVPath 訊息: 找不到參數檔")
                End If
                aXML = tbl.Rows(0)("INVPARAM")
            End Using           
            Dim aXmlDoc As New System.Xml.XmlDocument()
            aXmlDoc.LoadXml(aXML)
            Dim aXmlAttr As Xml.XmlAttribute = aXmlDoc.DocumentElement.Attributes("EInvoiceFilePath")
            If aXmlAttr Is Nothing Then
                Throw New Exception("Function: GetINVPath 訊息: 參數設定錯誤")
            End If
            Dim aRet As String = aXmlAttr.Value
            Return aRet
        Catch ex As Exception
            Throw New Exception("Function: GetINVPath 訊息: " & ex.ToString)
        End Try
    End Function
    Private Overloads Function ExeSQL(ByVal aSQL As String, Optional ByVal aConnectStrong As String = "") As Boolean
        Try
            If cn Is Nothing Then
                If String.IsNullOrEmpty(aConnectStrong) Then
                    cn = New OracleConnection(InvArgument.GetConnect)
                Else
                    cn = New OracleConnection(aConnectStrong)
                End If
            End If

            If cn.State = ConnectionState.Closed Then
                cn.Open()
            End If

            Using cmd As New OracleCommand(aSQL, cn)
                cmd.ExecuteNonQuery()
            End Using

        Catch ex As Exception

            Throw New Exception("Function: ExeSQL 訊息: " & ex.ToString)
        Finally
            If cn.State = ConnectionState.Open Then
                cn.Close()
            End If
            cn = Nothing
        End Try
        Return True
    End Function
    Private Sub ClearDataTableAndInfo()
        If Me.dtGiveErrData IsNot Nothing Then
            Me.dtGiveErrData.Rows.Clear()
            Me.dtGiveErrData.Dispose()
        End If
        _DiscountDate1 = Nothing
        _DiscountDate2 = Nothing
        _DiscountOv = False
        _InvDate1 = Nothing
        _InvDate2 = Nothing
        _InvId1 = Nothing
        _InvId2 = Nothing
        _InvNumDate1 = Nothing
        _InvNumDate2 = Nothing


    End Sub
#Region "IDisposable Support"
    Private disposedValue As Boolean ' 偵測多餘的呼叫

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: 處置 Managed 狀態 (Managed 物件)。
                If cn IsNot Nothing Then
                    cn.Close()
                    cn.Dispose()
                    cn = Nothing
                End If
                If dap IsNot Nothing Then
                    dap.Dispose()
                    dap = Nothing
                End If
                'If InvArgument IsNot Nothing Then
                '    InvArgument.Dispose()
                '    InvArgument = Nothing
                'End If
                'If Me.dtGiveErrData IsNot Nothing Then
                '    Me.dtGiveErrData.Rows.Clear()
                '    Me.dtGiveErrData.Dispose()
                'End If

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
Public Enum UpdateInvKin
    INV007 = 0
    INV014 = 1
    ALL = 2
End Enum