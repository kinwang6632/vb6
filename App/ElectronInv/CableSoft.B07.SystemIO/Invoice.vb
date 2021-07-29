Imports System.Data.OracleClient
Imports CableSoft.B07.InvType
Imports System.IO

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
    Public Property RunFrequencyTime As String = Nothing
    Public Property InvNumDate1 As String
    Public Property InvNumDate2 As String
    Public Property DiscountDate1 As String
    Public Property DiscountDate2 As String
    Public Property DiscountOv As Boolean
    Public Property DestroyReason As String = String.Empty
    Public Property UploadSource As CableSoft.B07.InvType.InvTypeEnum.UploadSource = InvTypeEnum.UploadSource.B07
    Public Property CreateInvType As List(Of CableSoft.B07.InvType.InvTypeEnum.CreateInvoiceType)
    Public Property YearMonth1 As String = Nothing
    Public Property YearMonth2 As String = Nothing
    Public Property C0701FilterPrintTime As Boolean = False

    Public Property SelectSQLText As String

        Get
            Dim Result As String
            Select Case RunInvType
                Case InvTypeEnum.INVTYPE.DestroyInv, InvTypeEnum.INVTYPE.DestroyReLoadInv,
                    InvTypeEnum.INVTYPE.NormalInvAll, InvTypeEnum.INVTYPE.NormalInvOV,
                    InvTypeEnum.INVTYPE.OnlyNormalInv
                    Result = GetInv007SQL()
                Case Else
                    Result = GetINV014SQL()
            End Select
            Return Result
        End Get
        Set(value As String)

        End Set
    End Property

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
        'DefaultCreateInvType()
    End Sub
    Public Sub New(ByVal strArgument As String)
        Me.New(New CableSoft.B07.Parameters.Argument(strArgument))
        'DefaultCreateInvType()
    End Sub
    Private Sub DefaultCreateInvType()
        Me.CreateInvType = New List(Of CableSoft.B07.InvType.InvTypeEnum.CreateInvoiceType)
        Me.CreateInvType.Add(InvTypeEnum.CreateInvoiceType.icAll)
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
    Public Overloads Function GetScreenWhere() As String
        Return GetScreenWhere(Me.RunInvType)
    End Function
    Public Overloads Function GetScreenWhere(ByVal InvType As CableSoft.B07.InvType.InvTypeEnum.INVTYPE) As String
        Try
            Dim aRetString As String = Nothing
            Select Case InvType
                Case InvTypeEnum.INVTYPE.DestroyInv, InvTypeEnum.INVTYPE.DestroyReLoadInv,
                        InvTypeEnum.INVTYPE.NormalInvAll, InvTypeEnum.INVTYPE.NormalInvOV
                    Dim Flag As String = String.Empty
                    Select Case InvType
                        Case InvTypeEnum.INVTYPE.DestroyInv
                            Flag = "註銷(產生註銷發票)"
                        Case InvTypeEnum.INVTYPE.DestroyReLoadInv
                            Flag = "註銷(註銷後重新上傳)"
                        Case InvTypeEnum.INVTYPE.NormalInvAll
                            Flag = "含作廢發票"
                        Case InvTypeEnum.INVTYPE.NormalInvOV
                            Flag = "僅作廢發票"
                    End Select
                    aRetString = String.Format("一般發票頁籤" & Environment.NewLine & _
                                              "發票起迄日 : {0} - {1} " & Environment.NewLine & _
                                              "發票號碼起迄 : {2} - {3} " & Environment.NewLine & _
                                              "上傳種類 : {4} " & Environment.NewLine & _
                                              "註銷原因 : {5} ",
                                              InvDate1, InvDate2, InvId1, InvId2, Flag, DestroyReason)
                Case InvTypeEnum.INVTYPE.OnlyDiscount, InvTypeEnum.INVTYPE.OnlyDiscountOV
                    Dim Flag As String = String.Empty
                    Select Case InvType
                        Case InvTypeEnum.INVTYPE.OnlyDiscount
                            Flag = "折讓發票"
                        Case InvTypeEnum.INVTYPE.OnlyDiscountOV
                            Flag = "折讓作廢發票"
                    End Select
                    aRetString = String.Format("折讓發票頁籤" & Environment.NewLine & _
                                          "折讓單據起迄日 : {0} - {1} " & Environment.NewLine & _
                                           "作廢註記 : {2} ",
                                          Me.DiscountDate1, Me.DiscountDate2, Flag)
                Case InvTypeEnum.INVTYPE.OnlyNotUseInvNum
                    aRetString = String.Format("空白未使用發票" & Environment.NewLine & _
                                          "空白未使用發票起迄年月 : {0} - {1} ",
                                          Me.InvNumDate1, Me.InvNumDate2)
            End Select

            Return aRetString
        Catch ex As Exception
            CableSoft.B07.Parameters.Argument.WriteErr(ex.ToString)
            Return String.Empty
        End Try
    End Function

    'Public Function chkNegativeData(ByVal dt As DataTable, ByVal TranInvType As Boolean,
    '                                ByRef WriteFileName As String, ByRef outMsg As String) As Boolean
    '    Dim aMsg As String = String.Empty
    '    Dim aChgMsg As String = String.Empty
    '    Try
    '        Dim str As String = GetScreenWhere()

    '        Dim aChkInvid = From aInvids In dt _
    '                      Order By aInvids.Item("INVID") Ascending _
    '                      Where Int32.Parse(aInvids.Item("Totalamount".ToUpper).ToString) < 0 _
    '                      Select aInvids.Item("INVID")
    '        Dim aINVLst As New List(Of String)
    '        For Each aInvid In aChkInvid
    '            aMsg = aMsg & Environment.NewLine & aInvid & " 中有負項收費資料"
    '            If Not aINVLst.Contains(aInvid) Then
    '                aChgMsg = aChgMsg & Environment.NewLine & aInvid
    '                aINVLst.Add(aInvid)
    '            End If
    '        Next
    '        If Not String.IsNullOrEmpty(aMsg) Then


    '        End If

    '        'If Not String.IsNullOrEmpty(aMsg) Then
    '        '    'XtraMessageBox.Show("發票資料含有負項資料！" & vbCrLf & "請修改資料再做上傳動作", "訊息", MessageBoxButtons.OK, MessageBoxIcon.Warning)
    '        '    If XtraMessageBox.Show("發票資料含有負項資料！" & vbCrLf & "是否要轉成 0.電子計算機發票 ？ ", _
    '        '                           "詢問", MessageBoxButtons.YesNo, _
    '        '                           MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then

    '        '        If objInvIO.TranInvType(aINVLst) Then
    '        '            XtraMessageBox.Show(String.Format("  已調整 {0} 筆資料 ", aINVLst.Count.ToString) & _
    '        '                                 Environment.NewLine & " 請重新產生上傳檔 ！", "訊息", MessageBoxButtons.OK, MessageBoxIcon.Information)
    '        '            aMsg = "已調整下列發票種類" & Environment.NewLine & New String("-"c, 40) & aChgMsg
    '        '            'clsCommon.WriteLog(FLogPath & aFile, aMsg)
    '        '            CableSoft.B07.Parameters.Argument.WriteLog(FLogPath & aFile, aMsg)
    '        '            Shell("notepad " & FLogPath & aFile, AppWinStyle.MaximizedFocus)
    '        '        End If
    '        '    Else
    '        '        aMsg = GetScreenWhere() & Environment.NewLine & "異常資料 : " & aMsg
    '        '        'clsCommon.WriteLog(FLogPath & aFile, aMsg)
    '        '        CableSoft.B07.Parameters.Argument.WriteLog(FLogPath & aFile, aMsg)
    '        '        Shell("notepad " & FLogPath & aFile, AppWinStyle.MaximizedFocus)
    '        '    End If
    '        '    Return False
    '        'End If
    '        Return True
    '    Catch ex As Exception
    '        Throw
    '    Finally
    '        outMsg = aMsg
    '    End Try

    'End Function
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
    Public Overridable Overloads Function UpdateINV() As Boolean
        Return UpdateINV(String.Empty, UpdateInvKin.ALL)
    End Function
    Private Function GetHowToCreateWhere() As String
        Dim aHowToCreateWhere As String = " 1=1 "
        If (Me.CreateInvType IsNot Nothing) Then
            If (Me.CreateInvType.Count > 0) AndAlso
                (Not Me.CreateInvType.Contains(InvTypeEnum.CreateInvoiceType.icAll)) Then
                aHowToCreateWhere = String.Empty
                If Me.CreateInvType.Count = 1 Then
                    aHowToCreateWhere = " HowToCreate = " & Integer.Parse(Me.CreateInvType.Item(0)).ToString
                Else
                    For Each HowtoCreate As CableSoft.B07.InvType.InvTypeEnum.CreateInvoiceType In Me.CreateInvType
                        If String.IsNullOrEmpty(aHowToCreateWhere) Then
                            aHowToCreateWhere = HowtoCreate
                        Else
                            aHowToCreateWhere = String.Format("{0},{1}",
                                                              aHowToCreateWhere, Integer.Parse(HowtoCreate))
                        End If
                    Next
                    aHowToCreateWhere = String.Format(" {0} IN ({1}) ",
                                                      "HowToCreate", aHowToCreateWhere)
                End If
            End If
        End If
        Return aHowToCreateWhere
    End Function
    Private Function GetInv007SQL() As String
        Dim aRet As String = Nothing
        Dim aSQL As String = String.Empty
        Dim aSQL2 As String = String.Empty
        Dim aWhere As String = String.Empty
        Dim aInvDateWhere As String = "1=0 "
        Dim aHowToCreateWhere As String = String.Empty
        aHowToCreateWhere = GetHowToCreateWhere()

        Try
            If InvArgument.GetStartAllUpload Then
                aWhere = " exists (select * from inv099 " & _
               " WHERE (inv099.UPLOADFLAG = 1 And substr(a.invid, 0, 2) = inv099.Prefix) " & _
               " AND TO_NUMBER(SUBSTR(A.INVID,3,8)) >= TO_NUMBER( inv099.startnum) " & _
               " AND TO_NUMBER(substr(a.invid,3,8)) <= to_number( inv099.endnum)) "
            Else
                aWhere = " A.INVOICEKIND = 1 AND A.BUSINESSID IS NULL  "
            End If
            If (Me.RunInvType = InvTypeEnum.INVTYPE.NormalInvAll) OrElse (Me.RunInvType = InvTypeEnum.INVTYPE.OnlyInvNum) OrElse
                (Me.RunInvType = InvTypeEnum.INVTYPE.NormalInvOV) OrElse (Me.RunInvType = InvTypeEnum.INVTYPE.DestroyInv) OrElse
                (Me.RunInvType = InvTypeEnum.INVTYPE.DestroyReLoadInv) OrElse (Me.RunInvType = InvTypeEnum.INVTYPE.OnlyNormalInv) Then
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

           
            If (Not String.IsNullOrEmpty(Me.RunFrequencyTime)) Then
                If IsDate(Me.RunFrequencyTime) Then
                    Me.RunFrequencyTime = DateTime.Parse(Me.RunFrequencyTime).ToString("yyyyMMddHHmmss")
                End If
                Select Case Me.RunInvType
                    Case InvTypeEnum.INVTYPE.NormalInvAll, InvTypeEnum.INVTYPE.NormalInvOV, InvTypeEnum.INVTYPE.OnlyNormalInv
                        aInvDateWhere = String.Format(" {0} AND A.CREATEINVDATE <= TO_DATE('{1}','YYYYMMDDHH24MISS') ",
                                               aInvDateWhere, Me.RunFrequencyTime.Replace(" ", "").Replace("/", "").ToString)
                    Case InvTypeEnum.INVTYPE.DestroyReLoadInv
                        aInvDateWhere = String.Format(" {0} AND A.DESTROYUPLOADTIME <= TO_DATE('{1}','YYYYMMDDHH24MISS') ",
                                               aInvDateWhere, Me.RunFrequencyTime.Replace(" ", "").Replace("/", "").ToString)
                    Case InvTypeEnum.INVTYPE.DestroyInv
                        aInvDateWhere = String.Format(" {0} AND A.PRINTTIME <= TO_DATE('{1}','YYYYMMDDHH24MISS') ",
                                              aInvDateWhere, Me.RunFrequencyTime.Replace(" ", "").Replace("/", "").ToString)
                    Case Else
                        aInvDateWhere = String.Format(" {0} AND  1= 0 ",
                                              aInvDateWhere)

                End Select

            End If
            '#6843 SYSDATE　>=　註銷時間+註銷後重新上傳排程，才可以把符合的資料撈出來並產生上傳 By Kin 2014/08/5
            If (UploadSource = InvTypeEnum.UploadSource.B07) AndAlso
                (Me.RunInvType = InvTypeEnum.INVTYPE.DestroyReLoadInv) Then
                Dim X0401Notify As String = GetX0401Notify(InvArgument.GetConnect, InvArgument.GetCompCode)
                If String.IsNullOrEmpty(X0401Notify) Then
                    X0401Notify = "0"
                End If
                aInvDateWhere = aInvDateWhere & String.Format("  AND SYSDATE >= A.DESTROYUPLOADTIME + ({0}/1440) ",
                                                              X0401Notify)

            End If
            'If (Not String.IsNullOrEmpty(Me.PrintTime)) Then
            '    If IsDate(Me.PrintTime) Then
            '        Me.PrintTime = DateTime.Parse(Me.PrintTime).ToString("yyyyMMddHHmmss")
            '    End If
            '    aInvDateWhere = String.Format(" {0} AND A.PRINTTIME <= TO_DATE('{1}','YYYYMMDDHH24MISS') ",
            '                                  aInvDateWhere, Me.PrintTime.Replace(" ", "").Replace("/", "").ToString)
            'End If

            aInvDateWhere = aInvDateWhere & " AND " & aHowToCreateWhere
            '----------------------------------------------------------------------------------------------------------------------------------------
            '正常發票
            '#7076 If creditcard4d is empty then add accountno field to sql,or add creditcard4d field By Kin 2015/09/14
            aSQL = "SELECT A.*,B.DESCRIPTION,B.QUANTITY,B.TOTALAMOUNT,B.SEQ,B.UNITPRICE, " & _
                        " Decode(substr(B.Creditcard4d,-4,999),null,substr(B.AccountNo,-4,999),substr( B.Creditcard4d,-4,999))  Creditcard4d " & _
                    " FROM " & FOwner & "INV007 A, " & FOwner & "INV008 B " & _
                    " WHERE 1=1 AND " & aInvDateWhere & _
                    " AND A.INVID=B.INVID " & _
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
            '6691 增加判斷是不是註銷發票 By Kin 2013/12/12
            If Me.RunInvType = InvTypeEnum.INVTYPE.DestroyInv Then
                If (UploadSource = InvTypeEnum.UploadSource.B07) AndAlso
                    (Not Me.C0701FilterPrintTime) Then
                    aSQL = String.Format("{0} AND A.UPLOADTIME IS NOT NULL " & _
                                     " AND ( A.DESTROYUPLOADTIME IS NULL OR DESTROYUPLOADTIME < A.UPLOADTIME ) ",
                                     aSQL)
                Else
                    '#6711 增加畫面條件過濾是否有
                    'Gateway  要用的語法與畫面條件
                    '#6794 原本A.PRINTTIME>A.DESTROYUPLOADTIME改為A.UPLOADTIME > A.DESTROYUPLOADTIME By Kin 2014/05/23
                    aSQL = String.Format("{0} AND A.UPLOADTIME IS NOT NULL " & _
                                    " AND A.PRINTTIME IS NOT NULL " & _
                                    " AND A.PRINTTIME > A.UPLOADTIME " & _
                                    " AND  ( A.DESTROYUPLOADTIME IS NULL  OR  A.UPLOADTIME > A.DESTROYUPLOADTIME ) ",
                                    aSQL)
                End If

            End If

            '-------------------------------------------------------------------------------------------------------------------------------------------
            '6691 增加判斷是否為註銷重新上傳 By Kin 2013/12/12
            If Me.RunInvType = InvTypeEnum.INVTYPE.DestroyReLoadInv Then
                aSQL = String.Format("{0} AND A.UPLOADTIME IS NOT NULL " & _
                                   " AND A.DESTROYUPLOADTIME > A.UPLOADTIME ", aSQL)
            End If

            '            aSQL = aSQL & " AND NVL(A.UPLOADFLAG,0) = 0 "
            '-------------------------------------------------------------------------------------------------------------------------------------------
            '作廢發票
            '#6262 增加過濾OBUPLOADFLAG=0
            '#7076 If creditcard4d is empty then add accountno field to sql,or add creditcard4d field By Kin 2015/09/14
            aSQL2 = "SELECT A.*,B.DESCRIPTION,B.QUANTITY,B.TOTALAMOUNT,B.SEQ,B.UNITPRICE, " & _
                    " Decode(substr(B.Creditcard4d,-4,999),null,substr(B.AccountNo,-4,999),substr( B.Creditcard4d,-4,999))  Creditcard4d " & _
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
            '#6691 如果是註銷發票則不抓取作廢的資料 By Kin 2013/12/12
            If (Me.RunInvType = InvTypeEnum.INVTYPE.DestroyInv) OrElse
                (Me.RunInvType = InvTypeEnum.INVTYPE.DestroyReLoadInv) OrElse
                (Me.RunInvType = InvTypeEnum.INVTYPE.OnlyNormalInv) Then
                aSQL2 = String.Empty
                If (Me.RunInvType = InvTypeEnum.INVTYPE.OnlyNormalInv) Then
                    aSQL = aSQL & " AND NVL(A.UPLOADFLAG,0) = 0 "
                End If
            Else
                aSQL = aSQL & " AND NVL(A.UPLOADFLAG,0) = 0 "
            End If
            If Not String.IsNullOrEmpty(aSQL2) Then
                aSQL = aSQL & " UNION ALL " & aSQL2
            End If

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
        '#6791 增加載具與愛心碼檢查 By Kin 2014/06/04
        '#6823 增加隨機碼檢核 By Kin 2014/07/08
        Try         
            aRet = " SELECT INV007.INVID,NULL GIVEUNITDESC ,1 FLAG" & _
                                " FROM INV007,INV028 " & _
                                " WHERE INV007.INVUSEID IS NOT NULL " & _
                                " AND INV007.INVUSEID = INV028.ITEMID " & _
                                " AND INV028.REFNO = 1 " & _
                                " AND INV007.GIVEUNITID IS NULL " & _
                                " AND INV007.COMPID = INV028.COMPID " & _
                                " AND INV007.LOVENUM IS NULL " & _
                                " AND INV007.INVID IN (" & Inv007SQL & ")" & _
                                " UNION ALL " & _
                                 "SELECT INV007.INVID,INV041.DESCRIPTION GIVEUNITDESC,2 FLAG " & _
                                 " FROM INV007,INV028,INV041 " & _
                                " WHERE INV007.INVUSEID IS NOT NULL " & _
                                " AND INV007.INVUSEID = INV028.ITEMID " & _
                                " AND INV028.REFNO = 1 " & _
                                " AND INV007.GIVEUNITID IS NOT  NULL " & _
                                " AND INV041.ITEMID = INV007.GIVEUNITID " & _
                                " AND INV007.COMPID = INV028.COMPID " & _
                                " AND INV007.COMPID = INV041.COMPID " & _
                                " AND INV041.BUSINESSID IS NULL " & _
                                " AND INV007.INVID IN (" & Inv007SQL & ")" & _
                                " UNION ALL " & _
                                 "SELECT INV007.INVID,NULL GIVEUNITDESC ,3 FLAG" & _
                                 " FROM INV007  " & _
                                " WHERE  INV007.CARRIERID1 IS  NULL " & _
                                " AND INV007.PRINTTIME IS NULL " & _
                                " AND INV007.INVID IN (" & Inv007SQL & ")" & _
                                " UNION ALL " & _
                                " SELECT INV007.INVID,NULL GIVEUNITDESC ,4 FLAG" & _
                                " FROM INV007,INV028 " & _
                                " WHERE INV007.LOVENUM IS NOT NULL " & _
                                " AND INV007.INVUSEID = INV028.ITEMID " & _
                                " AND INV028.REFNO = 1 " & _
                                " AND INV007.COMPID = INV028.COMPID " & _
                                " AND (LENGTH(INV007.LOVENUM)<3 OR LENGTH(INV007.LOVENUM)>10) " & _
                                " AND INV007.INVID IN (" & Inv007SQL & ")  " & _
                                  " UNION ALL " & _
                                " SELECT INV007.INVID,RANDOMNUM GIVEUNITDESC ,5 FLAG" & _
                                " FROM INV007 " & _
                                " WHERE INV007.INVID IN (" & Inv007SQL & ")  " & _
                                " AND LENGTH(NVL(RANDOMNUM,'X')) < 4 "                                
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
        For i As Integer = 0 To 0
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
                Using tbError As DataTable = OpenData(aSQL).Copy
                    If (tbError IsNot Nothing) AndAlso
                        (tbError.Rows.Count > 0) Then
                        If _dtGiveErrData Is Nothing Then
                            _dtGiveErrData = New DataTable("GiveErrData")
                            _dtGiveErrData.Columns.Add("InvId", GetType(String))
                            _dtGiveErrData.Columns.Add("Msg", GetType(String))
                        End If
                        '#6791 增加判斷載具資訊與愛心碼檢查 By Kin 2014/05/28
                        For Each rw As DataRow In tbError.Rows
                            Dim rwNew As DataRow = _dtGiveErrData.NewRow
                            rwNew.Item("InvId") = rw.Item("InvId")
                            Select Case Integer.Parse(rw.Item("Flag"))
                                Case 1
                                    rwNew.Item("Msg") = String.Format("發票號碼：[ {0} ] 捐贈單位無設定！", rw.Item("InvId"))
                                Case 2
                                    rwNew.Item("Msg") = String.Format("發票號碼：[ {0} ] 捐贈單位：[ {1} ] 無設定統一編號！",
                                                                 rw.Item("InvId"), rw.Item("GIVEUNITDESC"))
                                Case 3
                                    rwNew.Item("Msg") = String.Format("發票號碼：[ {0} ] 發票資料未列印，且無載具資訊！", rw.Item("InvId"))
                                Case 4
                                    rwNew.Item("Msg") = String.Format("發票號碼：[ {0} ] 發票資料愛心碼的欄位長度不符(應為3~10碼)！", rw.Item("InvId"))
                                Case 5
                                    rwNew.Item("Msg") = String.Format("發票號碼：[ {0} ] 發票資料無防偽隨機碼或長度不正確(不可少於4碼)！", rw.Item("InvId"))
                            End Select
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
        '#7001 判斷Inv001.InvoiceType是否有設定 By Kin 2015/04/08
        If _dtGiveErrData.Rows.Count = 0 Then
            aSQL = String.Format("Select * From {0}INV001 A Where 1 = 1 And {1} " & _
                           IIf(Me.IsAllCompCode, "", " AND COMPID='" & InvArgument.GetCompCode & "'"), FOwner, FWhere)
            Using dtInv001 As DataTable = OpenData(aSQL)
                For Each rw As DataRow In dtInv001.Rows
                    If DBNull.Value.Equals(rw.Item("InvoiceType")) Then
                        Dim rwNew As DataRow = _dtGiveErrData.NewRow
                        rwNew.Item("Msg") = String.Format("公司別 : [ {0} >> {1} ]，D01公司資料維護未設定電子發票上傳發票類別(InvoiceType) ！",
                                                          rw.Item("CompId").ToString,
                                                          rw.Item("CompsName").ToString)
                        _dtGiveErrData.Rows.Add(rwNew)
                        _dtGiveErrData.AcceptChanges()
                    End If
                Next
                dtInv001.Dispose()
            End Using
        End If
      
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
        Dim aGiveSQL As String = "SELECT INV007.INVID,NULL GIVEUNITDESC,1 Flag " & _
                            " FROM INV007,INV028 " & _
                             " WHERE INV007.INVUSEID IS NOT NULL " & _
                             " AND INV007.INVUSEID = INV028.ITEMID " & _
                             " AND INV028.REFNO = 1 " & _
                             " AND INV007.GIVEUNITID IS NULL " & _
                             " AND INV007.COMPID = INV028.COMPID " & _
                             " AND INV007.INVID IN (SELECT INVID FROM INV014 " & _
                                          " WHERE INVID IS NOT NULL AND ALLOWANCENO IN (" & aINV014SQL & "))" & _
                             " UNION ALL " & _
                              "SELECT INV007.INVID,INV041.DESCRIPTION GIVEUNITDESC,2 Flag " & _
                              " FROM INV007,INV028,INV041 " & _
                             " WHERE INV007.INVUSEID IS NOT NULL " & _
                             " AND INV007.INVUSEID = INV028.ITEMID " & _
                             " AND INV028.REFNO = 1 " & _
                             " AND INV007.GIVEUNITID IS NOT  NULL " & _
                             " AND INV041.ITEMID = INV007.GIVEUNITID " & _
                             " AND INV041.BUSINESSID IS NULL " & _
                             " AND INV007.COMPID = INV028.COMPID " & _
                             " AND INV007.COMPID = INV041.COMPID " & _
                             " AND INV007.INVID IN (SELECT INVID FROM INV014 " & _
                                          " WHERE INVID IS NOT NULL AND ALLOWANCENO IN (" & aINV014SQL & "))"
        Return aGiveSQL
    End Function
    Private Function GetINV014SQL() As String
        '折讓改成永遠不產生
        '#5881 將產生折讓恢復 By Kin 2010/12/14
        '#5922 增加INVFORMAT =1 條件 By Kin 2010/02/14
        '#6815 增加Inv014.BusinessId By Kin 2014/06/25
        Dim aSQL As String = String.Empty
        Dim aField As String = "A.PaperNo,A.PaperDate,A.InvDate, " & _
                        "A.InvID,A.SaleAmount,A.TaxAmount,A.BusinessId, " & _
                        "A.TaxType,A.COMPID,A.AllowanceNo,A.ObsoleteReason,B.BusinessId  BusinessId2,B.InvTitle "
        Dim aGroupField As String = "A.PaperNo,A.PaperDate,A.InvDate, " & _
                        "A.InvID,A.SaleAmount,A.TaxAmount,A.BusinessId, " & _
                        "A.TaxType,A.COMPID,A.AllowanceNo,A.ObsoleteReason,A.BusinessId2,A.InvTitle"
        Dim aAllUplaodWhere As String = String.Empty
        Dim aDiscountDateWhere As String = String.Empty
        Dim aUptTimeWhere As String = " 1 = 1 "

        Try
            If InvArgument.GetStartAllUpload Then
                aAllUplaodWhere = " 1 = 1"
            Else
                aAllUplaodWhere = " B.INVOICEKIND = 1 AND A.BUSINESSID IS NULL  "
            End If

            If (Me.RunInvType = InvTypeEnum.INVTYPE.OnlyDiscount) OrElse
                (Me.RunInvType = InvTypeEnum.INVTYPE.OnlyDiscountOV) Then
                aDiscountDateWhere = " 1 = 1 "
                If Not String.IsNullOrEmpty(Me.DiscountDate1) Then
                    aDiscountDateWhere = String.Format(" {0} AND  A.PAPERDATE  >= TO_DATE('{1}','YYYYMMDD')",
                                                        aDiscountDateWhere, Me.DiscountDate1.Replace("/", "").Replace(" ", "").ToString)
                End If
                If Not String.IsNullOrEmpty(_DiscountDate2) Then
                    If Not String.IsNullOrEmpty(aDiscountDateWhere) Then
                        aDiscountDateWhere = String.Format("{0} AND  A.PAPERDATE  <= TO_DATE('{1}','YYYYMMDD')",
                                                         aDiscountDateWhere, Me.DiscountDate2.Replace("/", "").Replace(" ", "").ToString)
                    Else
                        aDiscountDateWhere = String.Format("  A.PAPERDATE  <= TO_DATE('{0}','YYYYMMDD')",
                                                         Me.DiscountDate2.Replace("/", "").Replace(" ", "").ToString)
                    End If
                End If
            Else
                aDiscountDateWhere = " 1 = 0 "
            End If
         
            If (Not String.IsNullOrEmpty(Me.RunFrequencyTime)) Then
                aUptTimeWhere = String.Format(" {0} AND A.UPTTIME <= TO_DATE('{1}','YYYYMMDDHH24MISS') ",
                                             aUptTimeWhere, Me.RunFrequencyTime.Replace(" ", "").Replace("/", "").ToString)
            End If
            aSQL = "SELECT " & aField & _
                    " FROM " & FOwner & "INV014 A,INV007 B " & _
                    " WHERE " & aDiscountDateWhere & _
                    IIf(Me.IsAllCompCode, "", " AND A.COMPID='" & InvArgument.GetCompCode & "'") & _
                    " AND A.INVID = B.INVID " & _
                    " AND B.UPLOADFLAG =1 AND B.INVFORMAT = 1 " & _
                    " AND " & aAllUplaodWhere & _
                    " AND " & aUptTimeWhere
            '#5966 增加過濾發票號碼 By Kin 2011/03/30
            If Not String.IsNullOrEmpty(Me.InvId1) Then
                aSQL = aSQL & " AND B.INVID >='" & Me.InvId1 & "' "
            End If
            If Not String.IsNullOrEmpty(Me.InvId2) Then
                aSQL = aSQL & " AND B.INVID <='" & Me.InvId2 & "' "
            End If

            If Not String.IsNullOrEmpty(Me.YearMonth1) Then
                aSQL = String.Format("{0} AND TO_NUMBER( A.YEARMONTH ) >= {1} ",
                                     aSQL, Me.YearMonth1.Replace("/", "").Replace(" ", "").ToString)
            End If

            If Not String.IsNullOrEmpty(Me.YearMonth2) Then
                aSQL = String.Format("{0} AND TO_NUMBER( A.YEARMONTH ) <= {1} ",
                                     aSQL, Me.YearMonth2.Replace("/", "").Replace(" ", "").ToString)
            End If

            '#6001 增加折讓是否作廢，如果選擇作廢不要撈取資料 By Kin 2011/07/13
            'If Me.DiscountOv Then
            '    aSQL = aSQL & " AND  A.ISOBSOLETE = 'Y' AND Nvl(A.UPLOADFLAG,0)=1 "
            'Else
            '    aSQL = aSQL & " AND A.ISOBSOLETE = 'N' AND Nvl(A.UPLOADFLAG,0)=0  "
            'End If
            '#6615 增加判斷作廢上傳條件(OBUPLOADFLAG=0) By Kin 2014/01/06
            If Me.RunInvType = InvTypeEnum.INVTYPE.OnlyDiscountOV Then
                aSQL = aSQL & " AND  A.ISOBSOLETE = 'Y' AND Nvl(A.UPLOADFLAG,0)=1  AND NVL(A.OBUPLOADFLAG,0) = 0 "
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
    Public Overloads Function UpdateINV(ByVal Kind As UpdateInvKin) As Boolean
        Return UpdateINV(String.Empty, Kind)
    End Function
    Public Overloads Function BackupScreenMsg(ByVal aMsg As String, ByVal aTime As DateTime,
                                    ByVal BackupPath As String, ByVal FileType As String) As Boolean
        Try
            Dim aWriteMsg As String = Me.GetScreenWhere()
            aWriteMsg = aMsg & Environment.NewLine & aWriteMsg
            CableSoft.B07.Parameters.Argument.WriteProdcutLog(BackupPath, aTime, aWriteMsg, FileType)
        Catch ex As Exception
            'clsCommon.WriteErr(ex.ToString)
            CableSoft.B07.Parameters.Argument.WriteErr(ex.ToString)
        End Try
        Return True
    End Function
    Public Overloads Function BackupScreenMsg(ByVal aMsg As String, ByVal aTime As DateTime,
                                    ByVal BackupPath As String) As Boolean
        Return BackupScreenMsg(aMsg, aTime, BackupPath, String.Empty)
        'Try
        '    Dim aWriteMsg As String = Me.GetScreenWhere()
        '    aWriteMsg = aMsg & Environment.NewLine & aWriteMsg
        '    CableSoft.B07.Parameters.Argument.WriteProdcutLog(BackupPath, aTime, aWriteMsg)
        'Catch ex As Exception
        '    'clsCommon.WriteErr(ex.ToString)
        '    CableSoft.B07.Parameters.Argument.WriteErr(ex.ToString)
        'End Try
        'Return True
    End Function
    Public Overloads Shared Function ShowGiveErrData(ByVal dtError As DataTable, ByVal SaveFilePath As String,
                                                     ByVal aTime As DateTime) As Boolean
        If Not SaveFilePath.EndsWith("\") Then
            SaveFilePath = SaveFilePath & "\"
        End If
        Return ShowGiveErrData(dtError, String.Format("{0}\電子發票上傳異常資料記錄_{1}.Txt",
                                                     Path.GetDirectoryName(SaveFilePath), aTime.ToString("yyyyMMdd_HHmmss")))


    End Function
    Public Overloads Shared Function ShowGiveErrData(ByVal dtError As DataTable, ByVal LogFileName As String) As Boolean
        Try

            Dim aMsg As New System.Text.StringBuilder()
            'Dim aFile As String = "電子發票上傳異常資料記錄_" & _
            '            Format(Now, "yyyyMMdd_HHmmss") & ".Txt"
            For Each rw As DataRow In dtError.Rows
                aMsg.AppendLine(rw.Item("Msg").ToString)
            Next
            CableSoft.B07.Parameters.Argument.WriteLog(LogFileName, aMsg.ToString)
            'clsCommon.WriteLog(FLogPath & aFile, aMsg.ToString)
            'Shell("notepad " & FLogPath & aFile, AppWinStyle.MaximizedFocus)
        Catch ex As Exception
            Throw
        End Try
        Return True
    End Function
    Public Overridable Overloads Function UpdateINV(ByVal aDate As String, ByVal Kind As UpdateInvKin) As Boolean
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
            Dim aInv014Invid As String = String.Format("SELECT DISTINCT ALLOWANCENO FROM ({0})", GetINV014SQL)
            Using cmd As OracleCommand = cn.CreateCommand
                '直接從撈取INV007的SQL語法 UPD ，避免UPD到別筆資料
                '#6262 增加作廢上傳註記，所以正常發票與作廢發票要分開做 By Kin 2012/06/21
                '6691 增加判斷是否為註銷發票 By Kin 2013/12/12
                cmd.Parameters.Clear()
                Select Case Me.RunInvType
                    Case InvTypeEnum.INVTYPE.DestroyInv
                        cmd.Transaction = tra
                        cmd.Parameters.Add("DESTROYFLAG", OracleType.Number).Value = 1
                        cmd.Parameters.Add("DESTROYUPLOADTIME", OracleType.DateTime).Value = Date.Parse(aDate)
                        cmd.Parameters.Add("DESTROYREASON", OracleType.VarChar).Value = Me.DestroyReason
                        aSQL = String.Format("UPDATE INV007  SET  DESTROYFLAG = {0}, " & _
                                           "DESTROYUPLOADTIME = {1},DESTROYREASON = {2} " & _
                                           " WHERE ISOBSOLETE <> 'Y' AND UPLOADTIME IS NOT NULL " & _
                                           " AND INVID IN ( {3} )",
                                            ":DESTROYFLAG", ":DESTROYUPLOADTIME", ":DESTROYREASON",
                                            aInv007InvId)
                        'aSQL = "UPDATE " & FOwner & "INV007 A SET A.DESTROYFLAG = 1, " & _
                        '          " A.DESTROYUPLOADTIME = TO_DATE('" & aUpdTime & "','YYYYMMDDHH24MISS') " & _
                        '        " WHERE IsObsolete <> 'Y' AND INVID IN (" & aInv007InvId & ") "
                        cmd.Transaction = tra
                        cmd.CommandText = aSQL
                        cmd.ExecuteNonQuery()
                        cmd.Parameters.Clear()
                    Case InvTypeEnum.INVTYPE.NormalInvAll,
                        InvTypeEnum.INVTYPE.NormalInvOV, InvTypeEnum.INVTYPE.DestroyReLoadInv
                        cmd.Parameters.Clear()
                        cmd.Parameters.Add("UPLOADFLAG", OracleType.Number).Value = 1
                        cmd.Parameters.Add("UPLOADTIME", OracleType.DateTime).Value = DateTime.Parse(aDate)
                        aSQL = String.Format("UPDATE INV007 SET UPLOADFLAG = {0}, " & _
                                    "UPLOADTIME={1} " & _
                                    " WHERE ISOBSOLETE <> 'Y' AND INVID IN ({2}) ",
                                    ":UPLOADFLAG", ":UPLOADTIME", aInv007InvId)

                        'aSQL = "UPDATE " & FOwner & "INV007 A SET A.UPLOADFLAG = 1, " & _
                        '          " A.UPLOADTIME = TO_DATE('" & aUpdTime & "','YYYYMMDDHH24MISS') " & _
                        '        " WHERE IsObsolete <> 'Y' AND INVID IN (" & aInv007InvId & ") "
                        If Kind = UpdateInvKin.ALL OrElse Kind = UpdateInvKin.INV007 Then
                            cmd.Transaction = tra
                            cmd.CommandText = aSQL
                            cmd.ExecuteNonQuery()
                            cmd.Parameters.Clear()
                            If RunInvType <> InvTypeEnum.INVTYPE.DestroyReLoadInv Then
                                cmd.Parameters.Add("OBUPLOADFLAG", OracleType.Number).Value = 1
                                cmd.Parameters.Add("OBUPLOADTIME", OracleType.DateTime).Value = DateTime.Parse(aDate)
                                aSQL = String.Format("UPDATE INV007 SET OBUPLOADFLAG = {0}," & _
                                                   " OBUPLOADTIME = {1}  " & _
                                                   " WHERE ISOBSOLETE = 'Y' AND INVID IN ({2})",
                                                    ":OBUPLOADFLAG", ":OBUPLOADTIME", aInv007InvId)

                                '  aSQL = "UPDATE " & FOwner & "INV007 A SET A.OBUPLOADFLAG = 1, " & _
                                '" A.OBUPLOADTIME = TO_DATE('" & aUpdTime & "','YYYYMMDDHH24MISS') " & _
                                '" WHERE IsObsolete = 'Y' AND INVID IN (" & aInv007InvId & ") "
                                cmd.CommandText = aSQL
                                cmd.ExecuteNonQuery()
                            End If
                        End If
                    Case InvTypeEnum.INVTYPE.OnlyDiscount, InvTypeEnum.INVTYPE.OnlyDiscountOV
                        cmd.Transaction = tra
                        cmd.Parameters.Clear()
                        If Me.RunInvType = InvTypeEnum.INVTYPE.OnlyDiscount Then
                            cmd.Parameters.Add("UPLOADFLAG", OracleType.Number).Value = 1
                            cmd.Parameters.Add("UPLOADTIME", OracleType.DateTime).Value = DateTime.Parse(aDate)
                            If Not Me.IsAllCompCode Then
                                cmd.Parameters.Add("COMPID", OracleType.VarChar).Value = InvArgument.GetCompCode
                            End If
                            aSQL = String.Format("UPDATE INV014 SET UPLOADFLAG = {0}, " & _
                                "UPLOADTIME = {1} " & _
                                " WHERE ALLOWANCENO IN ( {2} ) " & _
                                " AND " & IIf(Me.IsAllCompCode, "{3}", "COMPID = {3}"),
                                              ":UPLOADFLAG", ":UPLOADTIME", aInv014Invid,
                                              IIf(Me.IsAllCompCode, "1=1", ":COMPID"))
                        Else
                            cmd.Parameters.Add("OBUPLOADFLAG", OracleType.Number).Value = 1
                            cmd.Parameters.Add("OBUPLOADTIME", OracleType.DateTime).Value = DateTime.Parse(aDate)
                            If Not Me.IsAllCompCode Then
                                cmd.Parameters.Add("COMPID", OracleType.VarChar).Value = InvArgument.GetCompCode
                            End If
                            aSQL = String.Format("UPDATE INV014 SET OBUPLOADFLAG = {0}, " & _
                                "OBUPLOADTIME = {1} " & _
                                " WHERE ALLOWANCENO IN ( {2} ) " & _
                                " AND " & IIf(Me.IsAllCompCode, "{3}", "COMPID = {3}"),
                                              ":OBUPLOADFLAG", ":OBUPLOADTIME", aInv014Invid,
                                              IIf(Me.IsAllCompCode, "1=1", ":COMPID"))
                        End If
                      
                      
                        If Kind = UpdateInvKin.ALL OrElse Kind = UpdateInvKin.INV014 Then
                            cmd.CommandText = aSQL
                            cmd.ExecuteNonQuery()
                        End If

                End Select

                tra.Commit()
            End Using
        Catch ex As Exception
            tra.Rollback()
            Throw
        Finally
            tra.Dispose()
            If cn.State = ConnectionState.Open Then
                cn.Close()
                cn.Dispose()
                cn = Nothing
            End If
        End Try
        Return True
        'Return aUpdTime
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
            Throw New Exception("Function: OpenData 訊息: " & ex.ToString & Environment.NewLine & _
                                "SQL語法：" & aSQL)
        Finally
            If cn IsNot Nothing Then
                cn.Close()
                cn.Dispose()
                cn = Nothing
            End If            
            If dap IsNot Nothing Then
                dap.Dispose()
                dap = Nothing
            End If
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
    Public Shared Function GetInvLogPath(ByVal ConnectString As String,
                                         ByVal CompCode As String) As String
        Dim LogName As String = GetINVPathName(ConnectString, CompCode, "MediaFilePath")

        LogName = Path.GetFullPath(LogName)
        If Not LogName.EndsWith("\") Then
            LogName = LogName & "\"
        End If
        Return LogName
    End Function
    Public Shared Function GetX0401Notify(ByVal ConnectString As String, ByVal CompCode As String) As String
        Return GetINVPathName(ConnectString, CompCode, "X0401notify")
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
            Throw
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
            If cn IsNot Nothing Then
                If cn.State = ConnectionState.Open Then
                    cn.Close()
                End If
                cn.Dispose()
                cn = Nothing
            End If
           
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