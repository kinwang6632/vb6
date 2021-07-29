Imports System
Imports System.Data
Imports System.Data.OracleClient
Public Class clsData
    Implements IDisposable

    Private strConnectString As String
    Private FINVType As INVTYPE
    Private FOwner As String = String.Empty
    Private FCompCode As String = String.Empty
    Private _InvDate1 As String
    Private _InvDate2 As String
    Private _InvId1 As String = String.Empty
    Private _InvId2 As String = String.Empty
    Private _DiscountDate1 As String
    Private _DiscountDate2 As String
    Private _InvNumDate1 As String
    Private _InvNumDate2 As String
    Private _DiscountOv As Boolean
    Private cn As OracleConnection = Nothing
    Private dap As OracleDataAdapter = Nothing
    Private FWhere As String = " A.IDENTIFYID1='1' AND A.IDENTIFYID2=0 "
    Private FUseAllComp As Boolean = False
    Private fInv007InvidSQL As String = String.Empty
    Private fInv014IdSQL As String = String.Empty
    Private GiveErrData As DataTable = Nothing

    Public Sub New()

    End Sub
    Public Sub New(ByVal aConnectString As String, ByVal aINVType As INVTYPE, _
                   ByVal aUseAllComp As Boolean)

        FINVType = aINVType
        strConnectString = aConnectString
        FOwner = clsCommon.GetOwner
        FCompCode = clsCommon.GetCompCode
        FUseAllComp = aUseAllComp
    End Sub
    Public Property StartAllUpload As Boolean = False
    Public ReadOnly Property dtGiveErrData As DataTable
        Get
            Return GiveErrData
        End Get
    End Property
    Public WriteOnly Property InvId1() As String
        Set(ByVal value As String)
            _InvId1 = value
        End Set
    End Property
    Public WriteOnly Property InvId2() As String
        Set(ByVal value As String)
            _InvId2 = value
        End Set
    End Property
    Public WriteOnly Property InvDate1() As String

        Set(ByVal value As String)
            _InvDate1 = value
        End Set
    End Property
    Public WriteOnly Property InvDate2() As String
        Set(ByVal value As String)
            _InvDate2 = value
        End Set
    End Property
    Public WriteOnly Property InvNumDate1 As String
        Set(value As String)
            _InvNumDate1 = value
        End Set
    End Property
    Public WriteOnly Property InvNumDate2 As String
        Set(value As String)
            _InvNumDate2 = value
        End Set
    End Property
    Public WriteOnly Property DiscountDate1() As String
        Set(ByVal value As String)
            _DiscountDate1 = value
        End Set
    End Property
    Public WriteOnly Property DiscountDate2() As String
        Set(ByVal value As String)
            _DiscountDate2 = value
        End Set
    End Property
    Public WriteOnly Property DiscountOv() As Boolean
        Set(ByVal value As Boolean)
            _DiscountOv = value
        End Set
    End Property
    Public Function GetInv003() As DataTable
        Dim aSQL As String = String.Empty
        Try
            'aSQL = "SELECT * FROM " & FOwner & "INV001 A WHERE A.COMPID='" & FCompCode & "' " & _
            '    " AND " & FWhere
            aSQL = "SELECT * FROM " & FOwner & "INV003 A WHERE " & FWhere
            Dim tbl As DataTable = OpenData(aSQL).Copy
            tbl.TableName = "INV003"
            Return tbl
        Catch ex As Exception
            Throw New Exception("Function: GetINV003 訊息: " & ex.ToString)
        End Try
    End Function
    Public Function GetINV001() As DataTable
        Dim aSQL As String = String.Empty
        Try
            'aSQL = "SELECT * FROM " & FOwner & "INV001 A WHERE A.COMPID='" & FCompCode & "' " & _
            '    " AND " & FWhere
            aSQL = "SELECT * FROM " & FOwner & "INV001 A WHERE " & FWhere
            Dim tbl As DataTable = OpenData(aSQL).Copy
            tbl.TableName = "INV001"
            Return tbl
        Catch ex As Exception
            Throw New Exception("Function: GetINV001 訊息: " & ex.ToString)
        End Try
    End Function
    '#5945 如果有負項資料，將發票種類改為電子計算機發票 By Kin 2011/03/17
    Public Function TranInvType(ByVal aInvId As List(Of String)) As Boolean

        If aInvId Is Nothing Then
            Return True
        Else
            If cn Is Nothing Then
                cn = New OracleConnection(strConnectString)
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
        Return UpdateINV(String.Empty)

    End Function
    Public Overridable Overloads Function UpdateINV(ByVal aDate As String) As String
        Dim aSQL As String = String.Empty
        Dim aUpdTime As String = String.Empty
        If (String.IsNullOrEmpty(aDate)) OrElse (Not IsDate(aDate)) Then
            aUpdTime = Format(Now, "yyyyMMddHHmmss")
        Else
            aUpdTime = Format(Date.Parse(aDate), "yyyyMMddHHmmss")
        End If

        If cn Is Nothing Then
            cn = New OracleConnection(strConnectString)
        End If
        If cn.State = ConnectionState.Closed Then
            cn.Open()
        End If

        '#5945 增加上傳檔案時間 By Kin 2011/03/17
        Dim tra As OracleTransaction = cn.BeginTransaction
        Try
            Using cmd As OracleCommand = cn.CreateCommand
                'aSQL = "UPDATE " & FOwner & "INV007 A SET A.UPLOADFLAG = 1, " & _
                '        " UPLOADTIME = TO_DATE('" & aUpdTime & "','YYYYMMDDHH24MISS') " & _
                '        " WHERE A.INVDATE>=TO_DATE('" & _InvDate1 & "','yyyymmdd')" & _
                '        " AND A.INVDATE<=TO_DATE('" & _InvDate2 & "','yyyymmdd') " & _
                '        " AND A.INVOICEKIND = 1 " & _
                '        " AND A.INVFORMAT = 1 " & _
                '        " AND A.BUSINESSID IS NULL " & _
                '        IIf(FUseAllComp, "", " AND A.COMPID='" & FCompCode & "'") & " AND " & FWhere
                'If FINVType = INVTYPE.OV Then
                '    aSQL = aSQL & " AND ISOBSOLETE = 'Y'"
                'End If

                '直接從撈取INV007的SQL語法 UPD ，避免UPD到別筆資料
                '#6262 增加作廢上傳註記，所以正常發票與作廢發票要分開做 By Kin 2012/06/21
                aSQL = "UPDATE " & FOwner & "INV007 A SET A.UPLOADFLAG = 1, " & _
                        " A.UPLOADTIME = TO_DATE('" & aUpdTime & "','YYYYMMDDHH24MISS') " & _
                        " WHERE IsObsolete <> 'Y' AND INVID IN (" & fInv007InvidSQL & ") "

                cmd.Transaction = tra
                cmd.CommandText = aSQL
                cmd.ExecuteNonQuery()

                aSQL = "UPDATE " & FOwner & "INV007 A SET A.OBUPLOADFLAG = 1, " & _
                        " A.OBUPLOADTIME = TO_DATE('" & aUpdTime & "','YYYYMMDDHH24MISS') " & _
                        " WHERE IsObsolete = 'Y' AND INVID IN (" & fInv007InvidSQL & ") "
                cmd.CommandText = aSQL
                cmd.ExecuteNonQuery()


                'aSQL = "UPDATE " & FOwner & "INV014 A SET UPLOADFLAG = 1 " & _
                '    " WHERE A.PaperDate >= TO_DATE('" & _DiscountDate1 & "','yyyymmdd')" & _
                '    " AND A.PaperDate <= TO_DATE('" & _DiscountDate2 & "','yyyymmdd')" & _
                '    " AND NVL(A.UPLOADFLAG,0) = 0 " & _
                '    " AND A.InvID IN (SELECT A.INVID FROM " & FOwner & " INV014 A," & FOwner & "INV007 B " & _
                '    " WHERE A.INVID = B.INVID AND B.UPLOADFLAG = 1  AND A.INVFORMAT = 1  ) " & _
                '    " AND " & FWhere
                'IIf(FUseAllComp, "", " AND A.COMPID='" & FCompCode & "'")

                '直接從撈取INV014的SQL語法 UPD ，避免UPD到別筆資料
                aSQL = "UPDATE " & FOwner & "INV014 SET UPLOADFLAG = 1, " & _
                    " UPLOADTIME = TO_DATE('" & aUpdTime & "','YYYYMMDDHH24MISS') " & _
                    " WHERE ALLOWANCENO IN (" & fInv014IdSQL & ") " & _
                    IIf(FUseAllComp, "", " AND COMPID='" & FCompCode & "'")

                cmd.CommandText = aSQL
                cmd.ExecuteNonQuery()

                tra.Commit()

            End Using
        Catch ex As Exception
            tra.Rollback()
            Throw New Exception(ex.ToString)

        Finally
            tra.Dispose()
            If cn.State = ConnectionState.Open Then
                cn.Close()
            End If
        End Try
        Return aUpdTime
    End Function
    Public Function GetINV007() As DataTable
        Dim aSQL As String = String.Empty
        Dim aSQL2 As String = String.Empty
        Dim aWhere As String = String.Empty
        '#5922 增加Invformat = 1 條件 By Kin 2010/02/14
        '#6370 增加判斷StartAllUpload By Kin 2012/11/23
        '#6370 StartAllUpload = true 取消INV007.INVOICEKIND=1，及取消INV007.BUSINESSID IS NULL的條件 By Kin 2012/12/05

        If StartAllUpload Then
            aWhere = " exists (select * from inv099 " & _
                " WHERE(inv099.UPLOADFLAG = 1 And substr(a.invid, 0, 2) = inv099.Prefix) " & _
                " AND TO_NUMBER(SUBSTR(A.INVID,3,8)) >= TO_NUMBER( inv099.startnum) " & _
                " AND TO_NUMBER(substr(a.invid,3,8)) <= to_number( inv099.endnum)) "
        Else
            aWhere = " A.INVOICEKIND = 1 AND A.BUSINESSID IS NULL  "
        End If
        Try
            '正常發票
            aSQL = "SELECT A.*,B.DESCRIPTION,B.QUANTITY,B.TOTALAMOUNT,B.SEQ,B.UNITPRICE " & _
                    " FROM " & FOwner & "INV007 A, " & FOwner & "INV008 B " & _
                    " WHERE INVDATE>=TO_DATE('" & _InvDate1 & "','yyyymmdd')" & _
                    " AND INVDATE<=TO_DATE('" & _InvDate2 & "','yyyymmdd') " & _
                    " AND A.INVID=B.INVID " & _
                    " AND NVL(A.UPLOADFLAG,0) = 0 " & _
                    " AND " & aWhere & _
                    " AND A.ISOBSOLETE = 'N' " & _
                    " AND A.INVFORMAT = 1 " & _
                    IIf(FUseAllComp, "", " AND COMPID='" & FCompCode & "'") & " AND " & FWhere
            '#5966 增加過濾發票號碼 By Kin 2011/03/30
            If Not String.IsNullOrEmpty(_InvId1) Then
                aSQL = aSQL & " AND A.INVID >='" & _InvId1 & "' "
            End If
            If Not String.IsNullOrEmpty(_InvId2) Then
                aSQL = aSQL & " AND A.INVID <='" & _InvId2 & "' "
            End If

            '作廢發票
            '#6262 增加過濾OBUPLOADFLAG=0
            aSQL2 = "SELECT A.*,B.DESCRIPTION,B.QUANTITY,B.TOTALAMOUNT,B.SEQ,B.UNITPRICE " & _
                    " FROM " & FOwner & "INV007 A," & FOwner & "INV008 B " & _
                    " WHERE INVDATE>=TO_DATE('" & _InvDate1 & "','yyyymmdd')" & _
                    " AND INVDATE<=TO_DATE('" & _InvDate2 & "','yyyymmdd') " & _
                    " AND A.INVID=B.INVID " & _
                    " AND A.ISOBSOLETE = 'Y' " & _
                    " AND A.INVFORMAT = 1 " & _
                    " AND NVL(A.OBUPLOADFLAG,0) = 0 " & _
                    " AND " & aWhere & _
                    IIf(FUseAllComp, "", " AND COMPID='" & FCompCode & "'") & " AND " & FWhere

            '#5966 增加過濾發票號碼 By Kin 2011/03/30
            If Not String.IsNullOrEmpty(_InvId1) Then
                aSQL2 = aSQL2 & " AND A.INVID >='" & _InvId1 & "' "
            End If
            If Not String.IsNullOrEmpty(_InvId2) Then
                aSQL2 = aSQL2 & " AND A.INVID <='" & _InvId2 & "' "
            End If
            
            '如果前端選擇只有作廢發票，第一個SQL不要找資料(正常的發票)
            If FINVType = INVTYPE.OV Then
                aSQL = aSQL & " AND 1=0 "
            End If
            aSQL = aSQL & " UNION ALL " & aSQL2
            aSQL = "SELECT Distinct A.*,Nvl(B.REFNO,0) REFNO,C.BUSINESSID BUSINESSID2 " & _
                " FROM (" & aSQL & ") A," & FOwner & "INV028 B," & FOwner & "INV041 C " & _
                " WHERE A.INVUSEID=B.ITEMID(+) " & IIf(FUseAllComp, "", " AND B.COMPID(+)='" & FCompCode & "'") & _
                " AND A.GIVEUNITID = C.ITEMID(+) AND A.COMPID = C.COMPID(+)  " & _
                " Order By A.INVID,A.SEQ "
            fInv007InvidSQL = "SELECT DISTINCT INVID FROM (" & aSQL & ") "
            '#6600 GIVEUNITID 與 LOVENUM 都是NULL才算錯誤 By Kin 2013/10/02
            Dim aGiveSQL As String = "SELECT INV007.INVID,NULL GIVEUNITDESC FROM INV007,INV028 " & _
                                " WHERE INV007.INVUSEID IS NOT NULL " & _
                                " AND INV007.INVUSEID = INV028.ITEMID " & _
                                " AND INV028.REFNO = 1 " & _
                                " AND INV007.GIVEUNITID IS NULL " & _
                                " AND INV007.LOVENUM IS NULL " & _
                                " AND INV007.INVID IN (" & fInv007InvidSQL & ")" & _
                                " UNION " & _
                                 "SELECT INV007.INVID,INV041.DESCRIPTION GIVEUNITDESC FROM INV007,INV028,INV041 " & _
                                " WHERE INV007.INVUSEID IS NOT NULL " & _
                                " AND INV007.INVUSEID = INV028.ITEMID " & _
                                " AND INV028.REFNO = 1 " & _
                                " AND INV007.GIVEUNITID IS NOT  NULL " & _
                                " AND INV041.ITEMID = INV007.GIVEUNITID " & _
                                " AND INV041.BUSINESSID IS NULL " & _
                                " AND INV007.INVID IN (" & fInv007InvidSQL & ")"

            ChkGiveData(aGiveSQL)
            Dim tbl As DataTable = OpenData(aSQL).Copy
            tbl.TableName = "INV007"
            Return tbl
        Catch ex As Exception
            Throw New Exception("Function: GetINV007 訊息: " & ex.ToString)
        End Try
    End Function
    Public Sub ChkGiveData(ByVal aSQL As String)
        Using tbReturn As DataTable = OpenData(aSQL).Copy
            If (tbReturn IsNot Nothing) AndAlso
                (tbReturn.Rows.Count > 0) Then
                If GiveErrData Is Nothing Then
                    GiveErrData = New DataTable("GiveErrData")
                    GiveErrData.Columns.Add("InvId", GetType(String))
                    GiveErrData.Columns.Add("Msg", GetType(String))
                End If
                For Each rw As DataRow In tbReturn.Rows
                    Dim rwNew As DataRow = GiveErrData.NewRow
                    rwNew.Item("InvId") = rw.Item("InvId")
                    If DBNull.Value.Equals(rw.Item("GIVEUNITDESC")) Then
                        rwNew.Item("Msg") = String.Format("發票號碼：[ {0} ] 捐贈單位無設定！", rw.Item("InvId"))
                    Else
                        rwNew.Item("Msg") = String.Format("發票號碼：[ {0} ] 捐贈單位：[ {1} ] 無設定統一編號！",
                                                          rw.Item("InvId"), rw.Item("GIVEUNITDESC"))
                    End If
                    GiveErrData.Rows.Add(rwNew)
                Next
                GiveErrData.AcceptChanges()
            End If
        End Using
    End Sub
    Public Function GetINV099() As DataTable
        Dim aSQL As String = String.Empty
        Dim aWhere As String = String.Empty
        Try
            aSQL = "SELECT * FROM " & FOwner & "INV099 "

            If (String.IsNullOrEmpty(_InvNumDate1)) AndAlso (String.IsNullOrEmpty(_InvNumDate2)) Then
                aWhere = " WHERE 1= 0 "
            Else
                aWhere = " WHERE NVL(UPLOADFLAG,0)=1 AND CURNUM<=ENDNUM "
                If Not String.IsNullOrEmpty(_InvNumDate1) Then
                    aWhere = String.Format(" {0} AND YEARMONTH>={1}", aWhere, _InvNumDate1)
                End If
                If Not String.IsNullOrEmpty(_InvNumDate2) Then
                    aWhere = String.Format(" {0} AND YEARMONTH<={1}", aWhere, _InvNumDate2)
                End If                
                If Not FUseAllComp Then
                    aWhere = String.Format(" {0} AND COMPID = '{1}' ", aWhere, FCompCode)
                End If
            End If
            aSQL = aSQL & aWhere
            Dim tbl As DataTable = OpenData(aSQL).Copy
            tbl.TableName = "INV099"
            Return tbl
        Catch ex As Exception
            Throw New Exception("Function: GetINV099 訊息: " & ex.ToString)
        End Try
    End Function
    Public Function GetINV014() As DataTable
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
        '#6600 測試不ok, 若有啟動(StartAllUpload=1), 若針對折讓資料異動, 需"取消"INVOICEKIND=1及BUSINESSID IS NULL的條件 By Kin 2013/11/11
        If StartAllUpload Then
            'aAllUplaodWhere = " exists (select * from inv099 " & _
            '    " WHERE(inv099.UPLOADFLAG = 1 And substr(a.invid, 0, 2) = inv099.Prefix) " & _
            '    " AND TO_NUMBER(SUBSTR(A.INVID,3,8)) >= TO_NUMBER( inv099.startnum) " & _
            '    " AND TO_NUMBER(substr(a.invid,3,8)) <= to_number( inv099.endnum)) "
            aAllUplaodWhere = " 1 = 1"
        Else
            aAllUplaodWhere = " B.INVOICEKIND = 1 AND A.BUSINESSID IS NULL  "
        End If
        Try
            aSQL = "SELECT " & aField & _
                    " FROM " & FOwner & "INV014 A,INV007 B " & _
                    " WHERE A.PaperDate>=TO_DATE('" & _DiscountDate1 & "','yyyymmdd')" & _
                    " AND A.PaperDate<=TO_DATE('" & _DiscountDate2 & "','yyyymmdd')" & _
                    IIf(FUseAllComp, "", " AND A.COMPID='" & FCompCode & "'") & _
                    " AND A.INVID = B.INVID " & _
                    " AND B.UPLOADFLAG =1 AND B.INVFORMAT = 1 " & _
                    " AND " & aAllUplaodWhere
            '#5966 增加過濾發票號碼 By Kin 2011/03/30
            If Not String.IsNullOrEmpty(_InvId1) Then
                aSQL = aSQL & " AND B.INVID >='" & _InvId1 & "' "
            End If
            If Not String.IsNullOrEmpty(_InvId2) Then
                aSQL = aSQL & " AND B.INVID <='" & _InvId2 & "' "
            End If
            '#6001 增加折讓是否作廢，如果選擇作廢不要撈取資料 By Kin 2011/07/13
            If _DiscountOv Then
                aSQL = aSQL & " AND  A.ISOBSOLETE = 'Y' AND Nvl(A.UPLOADFLAG,0)=1 "
            Else
                aSQL = aSQL & " AND A.ISOBSOLETE = 'N' AND Nvl(A.UPLOADFLAG,0)=0  "
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

            fInv014IdSQL = "SELECT ALLOWANCENO FROM (" & aSQL & ") "

            Dim aGiveSQL As String = "SELECT INV007.INVID,NULL GIVEUNITDESC FROM INV007,INV028 " & _
                               " WHERE INV007.INVUSEID IS NOT NULL " & _
                               " AND INV007.INVUSEID = INV028.ITEMID " & _
                               " AND INV028.REFNO = 1 " & _
                               " AND INV007.GIVEUNITID IS NULL " & _
                               " AND INV007.INVID IN (SELECT INVID FROM INV014 " & _
                                            " WHERE INVID IS NOT NULL AND ALLOWANCENO IN (" & fInv014IdSQL & "))" & _
                               " UNION ALL " & _
                                "SELECT INV007.INVID,INV041.DESCRIPTION GIVEUNITDESC FROM INV007,INV028,INV041 " & _
                               " WHERE INV007.INVUSEID IS NOT NULL " & _
                               " AND INV007.INVUSEID = INV028.ITEMID " & _
                               " AND INV028.REFNO = 1 " & _
                               " AND INV007.GIVEUNITID IS NOT  NULL " & _
                               " AND INV041.ITEMID = INV028.ITEMID " & _
                               " AND INV041.BUSINESSID IS NULL " & _
                               " AND INV007.INVID IN (SELECT INVID FROM INV014 " & _
                                            " WHERE INVID IS NOT NULL AND ALLOWANCENO IN (" & fInv014IdSQL & "))"

            ChkGiveData(aGiveSQL)
            Dim tbl As DataTable = OpenData(aSQL).Copy
            tbl.TableName = "INV014"
            Return tbl
        Catch ex As Exception
            Throw New Exception("Function: GetINV014 訊息: " & ex.ToString)
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
            End If
            If cn2 IsNot Nothing Then
                cn2.Close()
                cn2.Dispose()
                cn2 = Nothing
            End If
        End Try
        If cn2.State = ConnectionState.Closed Then
            cn2.Open()
        End If


    End Function
    Public Shared Function GetINVPathName(ByVal ConnectString As String, ByVal aCompCode As String, ByVal aAttrName As String) As String
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
            aSQL = "SELECT INVPARAM FROM " & FOwner & "INV003 WHERE COMPID='" & FCompCode & "'"
            Dim tbl As DataTable = OpenData(aSQL)
            If tbl.Rows.Count <= 0 Then
                Throw New Exception("Function: GetINVPath 訊息: 找不到參數檔")
            End If
            aXML = tbl.Rows(0)("INVPARAM")
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
                    cn = New OracleConnection(strConnectString)
                Else
                    cn = New OracleConnection(aConnectStrong)
                End If
            End If

            If cn.State = ConnectionState.Closed Then
                cn.Open()
            End If

            Dim cmd As OracleCommand

            cmd = New OracleCommand(aSQL, cn)
            cmd.ExecuteNonQuery()
        Catch ex As Exception

            Throw New Exception("Function: ExeSQL 訊息: " & ex.ToString)
        Finally
            If cn.State = ConnectionState.Open Then
                cn.Close()
            End If
            cn = Nothing
        End Try
    End Function
    Private Function OpenData(ByVal aSQL As String) As DataTable
        Try
            If cn Is Nothing Then
                cn = New OracleConnection(strConnectString)
            End If
            If dap Is Nothing Then
                dap = New OracleDataAdapter(aSQL, cn)
            End If
            Dim ds As New DataSet
            dap.Fill(ds)
            Return ds.Tables(0)
        Catch ex As Exception
            Throw New Exception("Function: OpenData 訊息: " & ex.ToString)
        Finally
            If cn.State = ConnectionState.Open Then
                cn.Close()
            End If
            cn = Nothing
            If dap IsNot Nothing Then
                dap.Dispose()
            End If
            dap = Nothing
        End Try

    End Function

    Protected Overrides Sub Finalize()
        If cn IsNot Nothing Then
            If cn.State = ConnectionState.Open Then
                cn.Close()
            End If
        End If
        cn = Nothing
        If dap IsNot Nothing Then
            dap.Dispose()
        End If
        dap = Nothing
        MyBase.Finalize()

    End Sub

#Region "IDisposable Support"
    Private disposedValue As Boolean ' 偵測多餘的呼叫

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                If cn IsNot Nothing Then
                    If cn.State = ConnectionState.Open Then
                        cn.Close()
                    End If
                End If
                cn = Nothing
                If dap IsNot Nothing Then
                    dap.Dispose()
                End If
                dap = Nothing
                If GiveErrData IsNot Nothing Then
                    GiveErrData.Dispose()
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
Public Enum INVTYPE
    All = 0
    OV = 1
    OnlyDiscount = 2
    OnlyInvNum = 3
End Enum