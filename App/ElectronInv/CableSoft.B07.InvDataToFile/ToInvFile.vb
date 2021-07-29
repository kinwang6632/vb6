Imports CableSoft.B07.InvType
Imports System.IO
Imports System.Text
Imports System.Data.OracleClient

Public Class ToInvFile
    Implements IDisposable
    Private FFilePath As String
    Private FInvDs As DataSet
    Private strTmp As String = String.Empty
    Private FNow As Date = Nothing
    Private _TotalCount As Integer = 0
    Private _A0401RcdCount As Integer = 0
    Private _A0501RcdCount As Integer = 0
    Private _B0301RcdCount As Integer = 0
    Private _B0401RcdCount As Integer = 0
    Private _E0402RcdCount As Integer = 0
    Private _C0701RcdCount As Integer = 0
    Private _BackPath As String = String.Empty
    Private Const fCutCount As Int32 = 4000
    Private Const fA0401FilePath As String = "A0401FilePath"
    Private Const fA0501FilePath As String = "A0501FilePath"
    Private Const fB0301FilePath As String = "B0301FilePath"
    Private Const fB0401FilePath As String = "B0401FilePath"
    Private Const fE0402FilePath As String = "E0402FilePath"
    Private Const fC0701FilePath As String = "C0701FilePath"
    Private Const fLogFilePathName As String = "MediaFilePath"
    Private InvArgument As CableSoft.B07.Parameters.Argument
    Private LogData As Dictionary(Of String, Object) = Nothing
    Public Property ExportElectronPath As String = Nothing
    Public ReadOnly Property TotalCount() As Int32
        Get
            Return _A0401RcdCount + A0501RcdCount + B0301RcdCount +
                 B0401RcdCount + E0402RcdCount + _C0701RcdCount
        End Get
    End Property
    Public ReadOnly Property A0401RcdCount() As Int32
        Get
            Return _A0401RcdCount
        End Get
    End Property
    Public ReadOnly Property A0501RcdCount() As Int32
        Get
            Return _A0501RcdCount
        End Get
    End Property
    Public ReadOnly Property B0301RcdCount() As Int32
        Get
            Return _B0301RcdCount
        End Get
    End Property
    Public ReadOnly Property B0401RcdCount() As Int32
        Get
            Return _B0401RcdCount
        End Get
    End Property
    Public ReadOnly Property E0402RcdCount As Integer
        Get
            Return _E0402RcdCount
        End Get
    End Property
    Public ReadOnly Property C0701RcdCount As Integer
        Get
            Return _C0701RcdCount
        End Get
    End Property
    Public Property UploadSource As CableSoft.B07.InvType.InvTypeEnum.UploadSource = InvTypeEnum.UploadSource.B07
    Public Property DestroyReason As String = String.Empty
    Public Sub New(ByVal InvArgument As CableSoft.B07.Parameters.Argument, ByVal aDataSet As DataSet)      
        Me.New(InvArgument, aDataSet, Now)
    End Sub

    Public Sub New(ByVal InvArgument As CableSoft.B07.Parameters.Argument,
                   ByVal aDataSet As DataSet, ByVal WrietTime As DateTime)
        'FNow = WrietTime
        'FInvDs = aDataSet
        'Me.InvArgument = InvArgument
        'Me.LogData = New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
        Me.New(InvArgument, aDataSet, WrietTime, InvTypeEnum.UploadSource.B07)
    End Sub
    Public Sub New(ByVal InvArgument As CableSoft.B07.Parameters.Argument,
                   ByVal aDataSet As DataSet, ByVal WrietTime As DateTime,
                   ByVal UploadSource As CableSoft.B07.InvType.InvTypeEnum.UploadSource)
        FNow = WrietTime
        FInvDs = aDataSet
        Me.InvArgument = InvArgument
        Me.LogData = New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
        Me.UploadSource = UploadSource
        If UploadSource = InvTypeEnum.UploadSource.Gateway Then
            Me.DestroyReason = "上傳後才列印"
        End If
    End Sub

    'Public Sub New(ByVal aFilePath As String, ByVal aDataSet As DataSet,
    '               ByVal InvArgument As CableSoft.B07.Parameters.Argument)
    '    FFilePath = aFilePath
    '    FInvDs = aDataSet
    '    FNow = Now
    '    Me.InvArgument = InvArgument
    'End Sub
    
    Public Overloads Function GetCompleteMsg(ByVal InvType As CableSoft.B07.InvType.InvTypeEnum.INVTYPE) As String
        Dim aMsg As String = Nothing
        Dim aA0401Name As String = "A0401"
        Dim aA0501Name As String = "A0501"
        Dim aB0301Name As String = "B0301"
        Dim aB0401Name As String = "B0401"
        Dim aE0402Name As String = "E0402"
        Dim aC0701Name As String = "C0701"
        If Int32.Parse(FInvDs.Tables("Inv001").Rows(0).Item("UploadType")) = 1 Then
            aA0401Name = "C0401"
            aA0501Name = "C0501"
            aB0301Name = "D0401"
            aB0401Name = "D0501"
        End If
        Try
            Select Case InvType
                Case CableSoft.B07.InvType.InvTypeEnum.INVTYPE.NormalInvAll
                    aMsg = String.Format("產生完成" & ControlChars.CrLf & _
                         "{0} (一般發票) : {1} 筆" & ControlChars.CrLf & _
                         "{2} (作廢發票) : {3} 筆", aA0401Name, _
                         Me.A0401RcdCount, _
                         aA0501Name, _
                         Me.A0501RcdCount)                
                Case CableSoft.B07.InvType.InvTypeEnum.INVTYPE.NormalInvOV
                    aMsg = String.Format("產生完成" & Environment.NewLine & _
                         "{0} (作廢發票) : {1} 筆",
                         aA0501Name, _
                         Me.A0501RcdCount)
                Case CableSoft.B07.InvType.InvTypeEnum.INVTYPE.OnlyDiscount
                    aMsg = String.Format("產生完成" & ControlChars.CrLf & _
                        "{0} (折讓發票) : {1} 筆", aB0301Name, Me.B0301RcdCount)
                Case CableSoft.B07.InvType.InvTypeEnum.INVTYPE.OnlyDiscountOV
                    aMsg = String.Format("產生完成" & ControlChars.CrLf & _
                      "{0} (折讓作廢發票) : {1} 筆", aB0401Name, Me.B0401RcdCount)
                Case CableSoft.B07.InvType.InvTypeEnum.INVTYPE.OnlyNotUseInvNum
                    aMsg = String.Format("產生完成" & ControlChars.CrLf & _
                                   "{0} (未空白使用字軌) : {1} 筆", aE0402Name,
                                   Me.E0402RcdCount)
                Case CableSoft.B07.InvType.InvTypeEnum.INVTYPE.DestroyInv
                    aMsg = String.Format("產生完成" & Environment.NewLine & _
                                       "{0} (註銷發票) : {1} 筆", aC0701Name,
                                       Me.C0701RcdCount)
                Case CableSoft.B07.InvType.InvTypeEnum.INVTYPE.DestroyReLoadInv,
                        InvTypeEnum.INVTYPE.OnlyInvNum,
                        InvTypeEnum.INVTYPE.OnlyNormalInv
                    aMsg = String.Format("產生完成" & ControlChars.CrLf & _
                         "{0} (一般發票) : {1} 筆",
                          aA0401Name, _
                         Me.A0401RcdCount)

            End Select
        Catch ex As Exception
            Throw
        End Try
        Return aMsg
    End Function

    Public Overloads Function WriteINVFile(ByVal aInvType As CableSoft.B07.InvType.InvTypeEnum.INVTYPE) As Boolean
        _TotalCount = 0
        'FNow = Now
        Try
            '_BackPath = CableSoft.B07.SystemIO.Invoice.GetINVPathName(InvArgument.GetConnect,
            '                                                    InvArgument.GetCompCode, fLogFilePathName)

            Dim aUploadType As Integer = 0
            If Not DBNull.Value.Equals(FInvDs.Tables("Inv001").Rows(0).Item("UploadType")) Then
                aUploadType = Integer.Parse(FInvDs.Tables("Inv001").Rows(0).Item("UploadType").ToString)
            End If
            If UploadSource = InvTypeEnum.UploadSource.Gateway Then
                aUploadType = 1
            End If
            Select Case aInvType
                Case InvTypeEnum.INVTYPE.OnlyInvNum, InvTypeEnum.INVTYPE.OnlyNormalInv
                    WriteA0401(aUploadType)
                Case InvType.InvTypeEnum.INVTYPE.NormalInvAll
                    WriteA0401(aUploadType)
                    WriteA0501(aUploadType)
                Case InvTypeEnum.INVTYPE.NormalInvOV
                    WriteA0501(aUploadType)
                Case InvTypeEnum.INVTYPE.OnlyDiscount
                    WriteB0301(aUploadType)
                Case InvTypeEnum.INVTYPE.OnlyDiscountOV
                    WriteB0401(aUploadType)
                Case InvTypeEnum.INVTYPE.OnlyNotUseInvNum
                    WriteE0402()
                Case InvTypeEnum.INVTYPE.DestroyInv
                    WriteC0701()
                Case InvTypeEnum.INVTYPE.DestroyReLoadInv
                    WriteA0401(aUploadType)


            End Select
            Return True
        Catch ex As Exception
            Throw
        End Try

        'clsData.GetINVPathName(clsCommon.GetConnect, clsCommon.GetCompCode, "MediaFilePath")
    End Function

    Public Overloads Function WriteINVFile(ByVal aInvType As CableSoft.B07.InvType.InvTypeEnum.INVTYPE, _
                             ByVal aUseDiscount As Boolean, _
                             ByVal aUseDisCountOv As Boolean, _
                             ByVal aUseSpaceInv As Boolean, _
                             ByVal aBackPath As String,
                             ByVal UploadType As Integer) As Boolean
        Try
            _TotalCount = 0
            'FNow = Now
            _BackPath = aBackPath
            Select Case aInvType
                Case InvType.InvTypeEnum.INVTYPE.NormalInvAll
                    WriteA0401(UploadType)
                    WriteA0501(UploadType)
                Case InvTypeEnum.INVTYPE.NormalInvOV
                    WriteA0501(UploadType)
                Case InvTypeEnum.INVTYPE.OnlyDiscount

            End Select
            If aUseDiscount Then
                If Not aUseDisCountOv Then
                    WriteB0301(UploadType)
                Else
                    WriteB0401(UploadType)
                End If
            End If
            If aUseSpaceInv Then
                WriteE0402()
            End If
            Return True
        Catch ex As Exception
            Throw New Exception(ex.ToString)
        End Try
    End Function
    Public Function WriteA0401(ByVal aUploadType As Integer) As Boolean
        Dim bduString As StringBuilder = Nothing
        'Dim objXml As Xml.XmlTextWriter = Nothing
        Dim dtXml As DataTable = Nothing
        Dim sInvId As String = String.Empty
        Dim aRowInv007s As DataRow() = _
            FInvDs.Tables("INV007").Select("IsObsolete='N'")
        Dim aFilePath As String = Nothing
        Dim aSeq As Integer = 1
        Dim strEnd As String = String.Empty
        Dim iCount As Int32 = 0
        Dim rowInv001 As DataRow = Nothing
        _A0401RcdCount = 0
        If aRowInv007s.Length <= 0 Then
            Return True
        End If

        If Not String.IsNullOrEmpty(ExportElectronPath) Then
            aFilePath = ExportElectronPath
            If Not aFilePath.EndsWith("\") Then
                aFilePath = aFilePath & "\"
            End If
        Else
            aFilePath = CableSoft.B07.SystemIO.Invoice.GetINVPathName(InvArgument.GetConnect,
                                                                               InvArgument.GetCompCode, fA0401FilePath)
        End If
       
        If String.IsNullOrEmpty(aFilePath) Then
            Throw New Exception("A0401 檔案路徑設定有誤！")
        End If
        If aUploadType = 0 Then
            bduString = New StringBuilder()
        Else
            dtXml = New DataTable
            For Each col As DataColumn In FInvDs.Tables("INV007").Columns
                dtXml.Columns.Add(col.ColumnName, col.DataType)
            Next
        End If


        Try
            For i As Integer = 0 To aRowInv007s.Length - 1

                '#5881 有可能是全部公司,所以INV001要Select By Kin 2010/12/14
                rowInv001 = FInvDs.Tables("INV001").Select("COMPID='" & aRowInv007s(i)("COMPID") & "'").GetValue(0)
                If sInvId <> aRowInv007s(i)("INVID") Then
                    iCount += 1
                    _TotalCount += 1
                    If aUploadType = 0 Then
                        If (iCount >= fCutCount) AndAlso (iCount Mod fCutCount = 0) Then
                            bduString.AppendLine(strEnd)
                            strEnd = WriteA0401End(aRowInv007s(i))
                            bduString.AppendLine(WriteA0401Head(aRowInv007s(i), _
                                                                rowInv001))
                            bduString.AppendLine(WriteA0401Detail(aRowInv007s(i), _
                                                                  rowInv001))
                            bduString.AppendLine(strEnd)
                            '#6296 第二個檔案會存到舊的路徑 By Kin 2012/08/16
                            'writeFile(InvFileType.A0401, bduString.ToString, aSeq)
                            writeFile(aFilePath, InvTypeEnum.InvFileType.A0401, bduString.ToString, aSeq)
                            aSeq += 1
                            strEnd = String.Empty
                            bduString = New StringBuilder
                        Else

                            If (Not String.IsNullOrEmpty(sInvId)) AndAlso _
                                (Not String.IsNullOrEmpty(sInvId)) AndAlso _
                                (Not String.IsNullOrEmpty(strEnd)) Then
                                bduString.AppendLine(strEnd)
                            End If

                            strEnd = WriteA0401End(aRowInv007s(i))
                            bduString.AppendLine(WriteA0401Head(aRowInv007s(i), _
                                                                rowInv001))
                            bduString.AppendLine(WriteA0401Detail(aRowInv007s(i), _
                                                                  rowInv001))
                        End If
                    Else
                        '#6371 將資料轉成XML檔 By Kin 2013/02/19                       
                        If dtXml.Rows.Count > 0 Then
                            WriteToXml(aFilePath, dtXml, rowInv001, InvTypeEnum.InvFileType.C0401)
                        End If
                        dtXml.Rows.Clear()
                        Dim rwNew As DataRow = dtXml.NewRow
                        rwNew.ItemArray = aRowInv007s(i).ItemArray
                        dtXml.Rows.Add(rwNew)
                    End If


                Else
                    If aUploadType = 0 Then
                        bduString.AppendLine(WriteA0401Detail(aRowInv007s(i), _
                                                         rowInv001))
                    Else
                        '#6371 將資料轉成XML檔 By Kin 2013/02/19
                        Dim rwNew As DataRow = dtXml.NewRow
                        rwNew.ItemArray = aRowInv007s(i).ItemArray
                        dtXml.Rows.Add(rwNew)
                    End If
                End If
                sInvId = aRowInv007s(i)("INVID").ToString
            Next
            If aUploadType = 0 Then
                bduString.AppendLine(WriteA0401End(aRowInv007s(aRowInv007s.Length - 1)))
                writeFile(aFilePath, InvTypeEnum.InvFileType.A0401, bduString.ToString, aSeq)
            Else
                WriteToXml(aFilePath, dtXml, rowInv001, InvTypeEnum.InvFileType.C0401)
            End If
            _A0401RcdCount = iCount
            Return True
        Catch ex As Exception
            Throw New Exception("Function: WriteA0401 訊息: " & ex.ToString)
        Finally
            If dtXml IsNot Nothing Then
                dtXml.Dispose()
            End If
        End Try
    End Function
    Private Function WriteA0401Head(ByVal aRowINV007 As DataRow, _
                                ByVal aRowINV001 As DataRow) As String
        Try
            '#6262 增加判斷INV003.ImportNote By Kin 2012/06/22
            Dim aImportNote As String = Left(Common.GetString(
                FInvDs.Tables("INV003").Select(String.Format("CompId='{0}'", aRowINV007.Item("CompId").ToString))(0).Item("ImportNote"), 1) & " ", 1)
            If aRowINV007.Item("TaxType").ToString <> "2" Then
                aImportNote = " "
            End If
            strTmp = String.Empty
            strTmp = "M" & aRowINV007("INVID") & aRowINV007("INVDATE") & _
                Space(8) & Common.GetString(aRowINV001("BusinessId"), 10, AlignType.Left) & _
                Common.GetString(aRowINV001("CompName"), 60, AlignType.Left) & _
                Space(228) & New String("0"c, 10) & _
                Common.GetString(Right(aRowINV007("INVID").ToString, 4), 60) & Space(439) & aImportNote & _
                Common.GetString(aRowINV001("TaxDivName"), 40) & Common.GetString(aRowINV001("TaxGetDate"), 10) & _
                Common.GetString(aRowINV001("TaxNum1"), 40) & Common.GetString(aRowINV001("TaxNum2"), 20) & _
                Space(22) & "05" & Space(1) & Common.GetString(aRowINV007("REFNO"), 1)
            Return strTmp

        Catch ex As Exception
            Throw New Exception("Function: WriteA0401Head 訊息: " & ex.ToString)
        End Try
    End Function
    Private Function WriteA0401Detail(ByVal aRowINV007 As DataRow, _
                                  ByVal aRowINV001 As DataRow) As String
        Try
            strTmp = String.Empty
            Dim aQuantity As String = String.Empty
            Dim aPrice As String = String.Empty
            If IsDBNull(aRowINV007("Quantity")) Then
                aQuantity = "000000000000.0000"
            Else
                aQuantity = New String("0", 12) & Format(System.Convert.ToInt32(aRowINV007("Quantity")), "#.0000")
                aQuantity = Right(aQuantity, 17)
            End If
            '#5910 Quantity 固定填1 By Kin 2011/01/12
            'aQuantity = "000000000001.0000"
            If IsDBNull(aRowINV007("TotalAmount")) OrElse (IsDBNull(aRowINV007("Quantity"))) Then
                aPrice = "000000000000.0000"
            Else
                aPrice = Format(System.Convert.ToDouble(Math.Round(System.Convert.ToDouble(aRowINV007("TotalAmount")) / _
                        System.Convert.ToDouble(aRowINV007("Quantity")), 2)), "#.0000").ToString
                aPrice = Right("000000000000" & aPrice, 17)
            End If
            '#5910 UnitPrice 固定填Inv007.InvAmount By Kin 2011/01/12
            'If IsDBNull(aRowINV007("InvAmount")) Then
            '    aPrice = "000000000000.0000"
            'Else
            '    aPrice = Right("000000000000" & Format(System.Convert.ToDouble(aRowINV007("InvAmount")), "#.0000"), 17)
            'End If
            strTmp = "D" & Common.GetString(aRowINV007("Description"), 256) & _
                    aQuantity & Space(6) & aPrice & _
                    Right("000000000000" & Format(aRowINV007("TotalAmount"), "#.0000"), 17) & _
                    Common.GetString(aRowINV007("SEQ"), 2, AlignType.Right, True) & Space(66)

            Return strTmp

        Catch ex As Exception
            Throw New Exception("Function: WriteA0401Detail 訊息: " & ex.ToString)
        End Try
    End Function
    Private Function WriteA0401End(ByVal aRowInv007 As DataRow) As String
        Try
            strTmp = String.Empty
            strTmp = "A" & Common.GetString( _
                Math.Round(System.Convert.ToDouble _
                           (aRowInv007("InvAmount"))).ToString, 12, _
                           AlignType.Right, True) & aRowInv007("TaxType") & _
                           "0.0" & aRowInv007("TAXRATE") & "00" & _
                           "000000000000" & Common.GetString(aRowInv007("INVAmount"), 12, AlignType.Right, True) & _
                            Space(45)

            '"000000000000" & "000000000000.0000" & "00000000.0000" & Space(3)

            Return strTmp
        Catch ex As Exception
            Throw New Exception("Function: WriteA0401End 訊息: " & ex.ToString)
        End Try

    End Function
    Public Function WriteA0501(ByVal UploadType As Integer) As Boolean
        Dim bduString As StringBuilder = Nothing
        Dim aRowINV007 As DataRow() = _
            FInvDs.Tables("INV007").Select("IsObsolete='Y'")
        Dim dtXML As DataTable = Nothing
        Dim aSeq As Integer = 1
        Dim iCount As Integer = 0
        Dim sINvId As String = String.Empty
        Dim rowInv001 As DataRow = Nothing
        Dim aFilePath As String = Nothing
        _A0501RcdCount = 0
        If aRowINV007.Length <= 0 Then
            Return True
        End If
        If Not String.IsNullOrEmpty(ExportElectronPath) Then
            aFilePath = ExportElectronPath
            If Not aFilePath.EndsWith("\") Then
                aFilePath = aFilePath & "\"
            End If
        Else
            aFilePath = CableSoft.B07.SystemIO.Invoice.GetINVPathName(InvArgument.GetConnect,
                                                                                InvArgument.GetCompCode, fA0501FilePath)
        End If

        If String.IsNullOrEmpty(aFilePath) Then
            Throw New Exception("A0501 檔案路徑設定有誤！")
        End If
        If UploadType = 0 Then
            bduString = New StringBuilder
        Else
            dtXML = New DataTable
            For Each col As DataColumn In FInvDs.Tables("INV007").Columns
                dtXML.Columns.Add(col.ColumnName, col.DataType)
            Next
        End If
        strTmp = String.Empty
        Try
            For i As Integer = 0 To aRowINV007.Length - 1
                rowInv001 = FInvDs.Tables("INV001").Select("COMPID='" & aRowINV007(i)("COMPID") & "'").GetValue(0)
                If sINvId <> aRowINV007(i).Item("INVID") Then
                    iCount += 1
                    _TotalCount += 1
                    If UploadType = 0 Then
                        strTmp = String.Empty
                        strTmp = "C" & aRowINV007(i).Item("INVID") & aRowINV007(i).Item("INVDATE") & _
                            New String("0"c, 10) & Common.GetString(rowInv001("BusinessId"), 10) & _
                            Format(aRowINV007(i)("UptTime"), "d") & Format(aRowINV007(i)("UptTime"), "HH:mm:ss") & _
                            Space(60) & Common.GetString(aRowINV007(i)("ObsoleteReason"), 200)
                        bduString.AppendLine(strTmp)
                        If (iCount >= fCutCount) AndAlso (iCount Mod fCutCount = 0) Then
                            '#6296 第二個檔案會存到舊的路徑 By Kin 2012/08/16
                            'writeFile(InvFileType.A0501, bduString.ToString, aSeq)
                            writeFile(aFilePath, InvTypeEnum.InvFileType.A0501, bduString.ToString, aSeq)
                            aSeq += 1
                            bduString = New StringBuilder
                        End If
                    Else
                        '#6371 將資料轉成XML檔 By Kin 2013/02/19                       
                        If dtXML.Rows.Count > 0 Then
                            WriteToXml(aFilePath, dtXML, rowInv001, InvTypeEnum.InvFileType.C0501)
                        End If
                        dtXML.Rows.Clear()
                        Dim rwNew As DataRow = dtXML.NewRow
                        rwNew.ItemArray = aRowINV007(i).ItemArray
                        dtXML.Rows.Add(rwNew)
                    End If
                End If
                sINvId = aRowINV007(i)("INVID").ToString
            Next
            If UploadType = 0 Then
                writeFile(aFilePath, InvTypeEnum.InvFileType.A0501, bduString.ToString, aSeq)
            Else
                WriteToXml(aFilePath, dtXML, rowInv001, InvTypeEnum.InvFileType.C0501)
            End If

            _A0501RcdCount = iCount
            Return True
        Catch ex As Exception
            Throw New Exception("Function: WriteA0501 訊息: " & ex.ToString)
        Finally
            If dtXML IsNot Nothing Then
                dtXML.Dispose()
            End If
        End Try
    End Function
    Private Sub WriteToXml(ByVal FilePath As String, ByVal dtXml As DataTable, ByVal rwInv001 As DataRow,
                           FileType As CableSoft.B07.InvType.InvTypeEnum.InvFileType)
        Dim XmlFileName As String = Nothing
        Try
            Select Case FileType
                Case InvTypeEnum.InvFileType.E0402

                    XmlFileName = GetXMLFile(FilePath, String.Format("{0}-{1}",
                                                                 dtXml.Rows(0).Item("YearMonth"),
                                                                 dtXml.Rows(0).Item("CurNum")), FileType)
                 
                  
                 
                Case InvTypeEnum.InvFileType.D0401, InvTypeEnum.InvFileType.D0501
                    XmlFileName = GetXMLFile(FilePath, dtXml.Rows(0).Item("AllowanceNo").ToString, FileType)
                Case Else
                    XmlFileName = GetXMLFile(FilePath, dtXml.Rows(0).Item("InvId"), FileType)
            End Select
            'If FileType <> InvTypeEnum.InvFileType.E0402 Then

            'Else

            'End If
            Dim XmlFile As Xml.XmlTextWriter
            'XmlFile = New Xml.XmlTextWriter(XmlFileName, System.Text.Encoding.UTF8)
            XmlFile = New Xml.XmlTextWriter(XmlFileName, New UTF8Encoding(False))
            Select Case FileType
                Case InvTypeEnum.InvFileType.C0401
                    WriteC0401(dtXml, rwInv001, XmlFile)
                Case InvTypeEnum.InvFileType.C0501
                    WriteC0501(dtXml, rwInv001, XmlFile)
                Case InvTypeEnum.InvFileType.D0401
                    WriteD0401(dtXml, rwInv001, XmlFile)
                Case InvTypeEnum.InvFileType.D0501
                    WriteD0501(dtXml, rwInv001, XmlFile)
                Case InvTypeEnum.InvFileType.E0402
                    WriteE0402(dtXml, rwInv001, XmlFile)
                Case InvTypeEnum.InvFileType.C0701
                    WriteC0701(dtXml, rwInv001, XmlFile)

            End Select

            XmlFile.WriteEndDocument()
            XmlFile.Flush()
            XmlFile.Close()

            '將備份取消(_BackPath設為Nothing) By Kin 2013/12/19
            _BackPath = Nothing
            Dim aBackFileName As String = Path.GetFileName(XmlFileName)
            If Not String.IsNullOrEmpty(_BackPath) Then
                If Not _BackPath.EndsWith("\") Then
                    aBackFileName = _BackPath & "\" & aBackFileName
                Else
                    aBackFileName = _BackPath & aBackFileName
                End If


            Else
                aBackFileName = String.Empty
            End If
            If (Not String.IsNullOrEmpty(aBackFileName)) AndAlso (aBackFileName <> XmlFileName) Then
                Dim byt() As Byte = System.IO.File.ReadAllBytes(XmlFileName)
                Using stm As New FileStream(aBackFileName, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.ReadWrite)
                    stm.Write(byt, 0, byt.Length)
                    stm.Flush()
                End Using
            End If
        Catch ex As Exception
            Throw
        End Try
        
    End Sub
    Private Function GetXMLFile(ByVal aFilePath As String, ByVal aInvId As String, ByVal aInvFileType As CableSoft.B07.InvType.InvTypeEnum.InvFileType) As String
        Dim aReturn As String = Nothing
        Try
            aReturn = [Enum].GetName(GetType(CableSoft.B07.InvType.InvTypeEnum.InvFileType), aInvFileType) & "-" & aInvId & "-" & _
          FNow.ToString("yyyyMMdd") & "-" & FNow.ToString("HHmmss") & ".xml"
            If aFilePath.EndsWith("\") Then
                aReturn = aFilePath & aReturn
            Else
                aReturn = aFilePath & "\" & aReturn
            End If
            Return aReturn
        Catch ex As Exception
            Throw
        End Try
      
    End Function
    Private Function AddDataToLogData(ByVal dtData As DataTable, ByVal CompId As String,
                                 ByVal InvFileType As CableSoft.B07.InvType.InvTypeEnum.InvFileType) As Boolean
        Try
            Dim UpLoadTypeName As String = "UploadType"

            LogData.Clear()
            LogData.Add("CompId", CompId)
            LogData.Add("UploadTime", FNow)
            LogData.Add("UploadSource", Integer.Parse(Me.UploadSource))
            LogData.Add(UpLoadTypeName, Integer.Parse(InvFileType) - 3)
            If InvFileType = InvTypeEnum.InvFileType.E0402 Then
                Dim lstRw As List(Of DataRow) = dtData.AsEnumerable.OrderBy(Function(rw As DataRow)
                                                                                Return rw.Item("YearMonth")
                                                                            End Function).ToList()
                LogData.Add("InvId", String.Format("{0}-{1}",
                                                   lstRw.Item(0).Item("YearMonth"),
                                                   lstRw.Item(lstRw.Count - 1).Item("YearMonth")))

            Else
                Select Case InvFileType
                    Case InvTypeEnum.InvFileType.C0401,
                            InvTypeEnum.InvFileType.C0501,
                            InvTypeEnum.InvFileType.C0701
                        LogData.Add("InvId", dtData.Rows(0).Item("InvId").ToString)
                    Case Else
                        LogData.Add("InvId", dtData.Rows(0).Item("AllowanceNo").ToString)
                End Select
            End If
        Catch ex As Exception
            LogData.Clear()
            Return False
        End Try
        Return True        
    End Function
    Private Function WriteC0401(ByVal dtInv007 As DataTable, ByVal rwInv001 As DataRow, ByVal XmlFile As Xml.XmlWriter) As Boolean
        Try
            Dim aCreditcard4d As String = Nothing
            CreateXMLFile(XmlFile)
            XmlFile.WriteStartElement("Invoice")
            XmlFile.WriteAttributeString("xmlns", "urn:GEINV:eInvoiceMessage:C0401:3.1")
            XmlFile.WriteAttributeString("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance")
            XmlFile.WriteAttributeString("xsi:schemaLocation", "urn:GEINV:eInvoiceMessage:C0401: 3.1 C0401.xsd")
            XmlFile.WriteStartElement("Main")
            '-------------------------------InvoiceNumbe----------------------------------------
            XmlFile.WriteStartElement("InvoiceNumber")
            XmlFile.WriteValue(dtInv007.Rows(0).Item("InvId").ToString)

            XmlFile.WriteEndElement() 'InvoiceNumber
            '------------------------------InvoiceDate---------------------------------------------------------------------
            XmlFile.WriteStartElement("InvoiceDate")

            XmlFile.WriteValue(DateTime.Parse(dtInv007.Rows(0).Item("InvDate").ToString).ToString("yyyyMMdd"))
            XmlFile.WriteEndElement() 'InvoiceDate
            '-----------------------------------------------------------------------------
            '------------------------------InvoiceTime---------------------------------------------------------------------
            '#6600 測試不OK,少了InvoiceTime By Kin 2013/11/12
            XmlFile.WriteStartElement("InvoiceTime")
            If Not DBNull.Value.Equals(rwInv001.Item("DefaultInvTime")) Then
                XmlFile.WriteValue(rwInv001.Item("DefaultInvTime"))
            Else
                XmlFile.WriteValue(DateTime.Parse(dtInv007.Rows(0).Item("CreateInvDate").ToString).ToString("HH:mm:ss"))
            End If

            XmlFile.WriteEndElement() 'InvoiceTime
            '-----------------------------------------------------------------------------
            '----------------------------Seller & Buyer----------------------------------------
            For i As Int32 = 0 To 1
                If i = 0 Then
                    XmlFile.WriteStartElement("Seller")
                Else
                    XmlFile.WriteStartElement("Buyer")
                End If
                '----------------------------------Identifier----------------------------------------------------------
                If i = 0 Then
                    'Seller
                    XmlFile.WriteStartElement("Identifier")
                    If Not DBNull.Value.Equals(rwInv001.Item("BusinessId")) Then
                        XmlFile.WriteValue(rwInv001.Item("BusinessId").ToString)
                    End If
                    XmlFile.WriteEndElement() 'Identifier
                Else
                    'Buyer
                    XmlFile.WriteStartElement("Identifier")
                    If DBNull.Value.Equals(dtInv007.Rows(0).Item("BusinessId")) Then
                        XmlFile.WriteValue("0000000000")
                    Else
                        XmlFile.WriteValue(dtInv007.Rows(0).Item("BusinessId").ToString)
                    End If
                    XmlFile.WriteEndElement() 'Identifier
                End If
                '---------------------------------------------------------------------------------------------------------
                '---------------------------------Name------------------------------------------------------------------
                XmlFile.WriteStartElement("Name")
                If i = 0 Then
                    'Seller
                    XmlFile.WriteValue(rwInv001.Item("CompName").ToString)
                Else
                    'Buyer
                    If DBNull.Value.Equals(dtInv007.Rows(0).Item("BusinessId")) Then
                        XmlFile.WriteValue(Right(dtInv007.Rows(0).Item("InvId").ToString, 4))
                    Else
                        If Not DBNull.Value.Equals(dtInv007.Rows(0).Item("InvTitle").ToString) Then
                            XmlFile.WriteValue(dtInv007.Rows(0).Item("InvTitle").ToString)
                        End If
                    End If
                End If
                XmlFile.WriteEndElement() 'Name
                '-------------------------------------------------------------------------------------------------------------
                '-------------------------------Address-----------------------------------------------------------------------
                'Seller need to  add  Address element  For Jacy 
                If i = 0 Then 'Seller 才要 For Jacy
                    XmlFile.WriteStartElement("Address")
                    XmlFile.WriteValue(rwInv001.Item("CompAddr").ToString)
                    XmlFile.WriteEndElement() 'Address
                End If
                '-------------------------------------------------------------------------------------------------------------
                XmlFile.WriteEndElement() 'Seller & Buyer
            Next
            '----------------------------------------------------------------------------
            '-----------------------------------------------------------------------------
            If (Not DBNull.Value.Equals(dtInv007.Rows(0).Item("PrintTime"))) Then
                XmlFile.WriteStartElement("CheckNumber")
                XmlFile.WriteValue(dtInv007.Rows(0).Item("CheckNo").ToString)
                XmlFile.WriteEndElement() 'CheckNumber
            End If
            If (Not DBNull.Value.Equals(dtInv007.Rows(0).Item("TaxType"))) AndAlso
                (Integer.Parse(dtInv007.Rows(0).Item("TaxType")) = 2) Then
                Dim drInv003() As DataRow = FInvDs.Tables("INV003").Select("COMPID='" & dtInv007.Rows(0).Item("COMPID") & "'")
                If drInv003 IsNot Nothing AndAlso drInv003.Count > 0 Then
                    XmlFile.WriteStartElement("CustomsClearanceMark")
                    XmlFile.WriteValue(drInv003(0).Item("ImportNote").ToString)
                    XmlFile.WriteEndElement() 'CustomsClearanceMark
                End If

            End If
            '#7076 增加 MainRemark By Kin 2015/08/31
            '----------------------------MainRemark------------------------------------
            aCreditcard4d = Nothing
            For Each rw As DataRow In dtInv007.Rows
                If Not DBNull.Value.Equals(rw.Item("Creditcard4d")) Then
                    aCreditcard4d = "信用卡末四碼:" & rw.Item("Creditcard4d")
                    Exit For
                End If
            Next
            If Not String.IsNullOrEmpty(aCreditcard4d) Then
                XmlFile.WriteStartElement("MainRemark")
                XmlFile.WriteValue(aCreditcard4d)
                XmlFile.WriteEndElement() 'MainRemark
            End If

            '--------------------------------------------------------------------------------
            '----------------------------PermitDate--------------------------------------
            '#6600 取消此元素 By Kin 2013/10/02
            'XmlFile.WriteStartElement("PermitDate")
            'If Not DBNull.Value.Equals(rwInv001.Item("TaxGetDate")) Then
            '    XmlFile.WriteValue(Date.Parse(rwInv001.Item("TaxGetDate").ToString).ToString("yyyyMMdd"))
            'End If
            'XmlFile.WriteEndElement() 'PermitDate
            '---------------------------------------------------------------------------------
            '----------------------------------PermitWord--------------------------------
            '#6600 取消此元素 By Kin 2013/10/02
            'XmlFile.WriteStartElement("PermitWord")
            'If Not DBNull.Value.Equals(rwInv001.Item("taxnum1")) Then
            '    XmlFile.WriteValue(rwInv001.Item("taxnum1").ToString)
            'End If
            'XmlFile.WriteEndElement() 'PermitWord
            '---------------------------------------------------------------------------------
            '--------------------------------PermitNumber------------------------------
            '#6600 取消此元素 By Kin 2013/10/02
            'XmlFile.WriteStartElement("PermitNumber")
            'If Not DBNull.Value.Equals(rwInv001.Item("taxnum2")) Then
            '    XmlFile.WriteValue(rwInv001.Item("taxnum2").ToString)
            'End If
            'XmlFile.WriteEndElement() 'PermitNumber
            '---------------------------------------------------------------------------------
            '-----------------------------InvoiceType------------------------------------
            XmlFile.WriteStartElement("InvoiceType")
            Select Case Int32.Parse(dtInv007.Rows(0).Item("InvFormat"))
                Case 1
                    '#7001 用Inv001.InvoiceType By Kin 2015/04/02
                    'XmlFile.WriteValue("05")
                    XmlFile.WriteValue(rwInv001.Item("InvoiceType").ToString)
                Case 2
                    XmlFile.WriteValue("02")
                Case 3
                    XmlFile.WriteValue("01")
            End Select
            XmlFile.WriteEndElement() 'InvoiceType
            '---------------------------------------------------------------------------------

            '-----------------------------------Donate Mark--------------------------------
            '#6600 如果PrintTime有值,則捐贈為 0 By Kin 2013/10/07
            XmlFile.WriteStartElement("DonateMark")
            If (Int32.Parse(dtInv007.Rows(0).Item("RefNo")) = 1) AndAlso
                 (DBNull.Value.Equals(dtInv007.Rows(0).Item("PrintTime"))) Then
                XmlFile.WriteValue("1")
            Else
                XmlFile.WriteValue("0")
            End If
            XmlFile.WriteEndElement() 'DonateMark
            '---------------------------------------------------------------------------------
            '----------------------------------CarrierType--------------------------------
            '#6600 如果是PrintTime is Null才可寫入 By Kin 2013/10/07
            '#6600 測試不OK,捐贈也不需要載具資料 By Kin 2013/11/11
            If (Not DBNull.Value.Equals(dtInv007.Rows(0).Item("CarrierType"))) AndAlso
                (DBNull.Value.Equals(dtInv007.Rows(0).Item("PrintTime"))) Then
                '#6600 若INV007. PrintTime有值，此欄NULL By Kin 2013/10/02
                XmlFile.WriteStartElement("CarrierType")
                XmlFile.WriteValue(dtInv007.Rows(0).Item("CarrierType").ToString)
                XmlFile.WriteEndElement() 'CarrierType
            End If
            '-------------------------------------------------------------------------------------
            '----------------------------------CarrierId1--------------------------------------
            '#6600 如果是PrintTime is Null才可寫入 By Kin 2013/10/07
            '#6600 測試不OK,捐贈也不需要載具資料 By Kin 2013/11/11
            '捐贈也要將載具資料留住,所以將判斷捐贈拿掉 By Kin 2013/12/10 For Jacy
            If (Not DBNull.Value.Equals(dtInv007.Rows(0).Item("CarrierId1"))) AndAlso
                 (DBNull.Value.Equals(dtInv007.Rows(0).Item("PrintTime"))) Then
                XmlFile.WriteStartElement("CarrierId1")
                XmlFile.WriteValue(dtInv007.Rows(0).Item("CarrierId1").ToString)
                XmlFile.WriteEndElement() 'CarrierId1
            End If
            '-------------------------------------------------------------------------------------
            '----------------------------------CarrierId2--------------------------------------
            '#6600 測試不OK,捐贈也不需要載具資料 By Kin 2013/11/11
            If (DBNull.Value.Equals(dtInv007.Rows(0).Item("PrintTime"))) Then
                '#6600 如果是PrintTime is Null才可寫入 By Kin 2013/10/07
                If Not DBNull.Value.Equals(dtInv007.Rows(0).Item("CarrierId2")) Then
                    XmlFile.WriteStartElement("CarrierId2")
                    XmlFile.WriteValue(dtInv007.Rows(0).Item("CarrierId2").ToString)
                    XmlFile.WriteEndElement() 'CarrierId2
                Else
                    If Not DBNull.Value.Equals(dtInv007.Rows(0).Item("CarrierId1")) Then
                        XmlFile.WriteStartElement("CarrierId2")
                        XmlFile.WriteValue(dtInv007.Rows(0).Item("CarrierId1").ToString)
                        XmlFile.WriteEndElement() 'CarrierId2
                    End If
                End If
            End If

            '-------------------------------------------------------------------------------------
            '----------------------------------PrintMark-------------------------------------------
            '#6600 抓取INV007. PrintTime有值就=’Y’，反之則=’N’ By Kin 2013/10/07
            XmlFile.WriteStartElement("PrintMark")
            If DBNull.Value.Equals(dtInv007.Rows(0).Item("PrintTime")) Then
                XmlFile.WriteValue("N")
            Else
                XmlFile.WriteValue("Y")
            End If
            'If (Int32.Parse(dtInv007.Rows(0).Item("RefNo")) = 1) OrElse
            '    (Not DBNull.Value.Equals(dtInv007.Rows(0).Item("CarrierId1"))) Then
            '    XmlFile.WriteValue("N")
            'Else
            '    If DBNull.Value.Equals(dtInv007.Rows(0).Item("PrintTime")) Then
            '        XmlFile.WriteValue("N")
            '    Else
            '        XmlFile.WriteValue("Y")
            '    End If
            'End If
            XmlFile.WriteEndElement() 'PrintMark
            '-----------------------------------------------------------------------------------------
            '----------------------------------NPOBAN--------------------------------------------
            '#6600 若INV007屬捐贈，(1)則填INV007.愛心碼(LoveNum)或(2)用INV007. GiveUnitId對應至捐贈單位統一編號(INV041. BUSINESSID)
            '#7215 Don't write element of NPOBAN when printtime has value by kin 2015/03/18
            If DBNull.Value.Equals(dtInv007.Rows(0).Item("PrintTime")) Then
                If (Int32.Parse(dtInv007.Rows(0).Item("RefNo")) = 1) Then
                    If (Not DBNull.Value.Equals(dtInv007.Rows(0).Item("LoveNum"))) Then
                        XmlFile.WriteStartElement("NPOBAN")
                        XmlFile.WriteValue(dtInv007.Rows(0).Item("LoveNum").ToString)
                        XmlFile.WriteEndElement()
                    Else
                        If (Not DBNull.Value.Equals(dtInv007.Rows(0).Item("BUSINESSID2"))) Then
                            XmlFile.WriteStartElement("NPOBAN")
                            XmlFile.WriteValue(dtInv007.Rows(0).Item("BUSINESSID2").ToString)
                            XmlFile.WriteEndElement()
                        End If
                    End If
                End If
            End If
            'If (Int32.Parse(dtInv007.Rows(0).Item("RefNo")) = 1) AndAlso
            '    (Not DBNull.Value.Equals(dtInv007.Rows(0).Item("BUSINESSID2"))) Then
            '    XmlFile.WriteStartElement("NPOBAN")
            '    XmlFile.WriteValue(dtInv007.Rows(0).Item("BUSINESSID2").ToString)
            '    XmlFile.WriteEndElement()
            'End If
            '-------------------------------------------------------------------------------------------
            '---------------------------------RandomNumber-------------------------------------
            '#6600 如果隨機碼有值,則填入隨機碼,否則填入AAAA By Kin 2013/10/07
            XmlFile.WriteStartElement("RandomNumber")
            If DBNull.Value.Equals(dtInv007.Rows(0).Item("RandomNum")) Then
                XmlFile.WriteValue("AAAA")
            Else
                XmlFile.WriteValue(dtInv007.Rows(0).Item("RandomNum").ToString)
            End If

            XmlFile.WriteEndElement() 'RandomNumber
            '--------------------------------------------------------------------------------------------
            XmlFile.WriteEndElement() 'Main

            '---------------------------------------------------------------------------------------------
            '-----------------------------------'Details-------------------------------------------------
            XmlFile.WriteStartElement("Details")
            For Each rw As DataRow In dtInv007.Rows
                '-------------------------ProductItem-------------------------------------------------
                XmlFile.WriteStartElement("ProductItem")
                '-----------------------------------------------------------------------------------------
                '---------------------------------Description------------------------------------------
                XmlFile.WriteStartElement("Description")
                XmlFile.WriteValue(rw.Item("Description").ToString)
                XmlFile.WriteEndElement() 'Description
                '-----------------------------------------------------------------------------------------
                '--------------------------------Quantity----------------------------------------------
                XmlFile.WriteStartElement("Quantity")
                XmlFile.WriteValue(rw.Item("Quantity").ToString)
                XmlFile.WriteEndElement() 'Quantity
                '-------------------------------------------------------------------------------------------                
                '---------------------------------UnitPrice-----------------------------------------------
                XmlFile.WriteStartElement("UnitPrice")
                XmlFile.WriteValue(Math.Round((Double.Parse(rw.Item("TotalAmount")) / Double.Parse(rw.Item("Quantity"))), 0, MidpointRounding.AwayFromZero))
                'XmlFile.WriteValue(rw.Item("UnitPrice").ToString)
                XmlFile.WriteEndElement() 'UnitPrice
                '-------------------------------------------------------------------------------------------
                '---------------------------------Amount------------------------------------------------
                XmlFile.WriteStartElement("Amount")
                XmlFile.WriteValue(rw.Item("TotalAmount"))
                XmlFile.WriteEndElement() 'Amount
                '---------------------------------------------------------------------------------------------
                '---------------------------------SequenceNumber-------------------------------------
                XmlFile.WriteStartElement("SequenceNumber")
                XmlFile.WriteValue(rw.Item("Seq").ToString)
                XmlFile.WriteEndElement() 'SequenceNumber
                '---------------------------------------------------------------------------------------------                              
                XmlFile.WriteEndElement() 'ProductItem
                '------------------------------------------------------------------------------------------
            Next
            XmlFile.WriteEndElement() 'Details
            '---------------------------------------------------------------------------------------------
            '-----------------------------------End Amount-----------------------------------------------
            XmlFile.WriteStartElement("Amount")
            '-------------------------------------SalesAmount----------------------------------------------
            XmlFile.WriteStartElement("SalesAmount")
            '取小數四捨五入至整數
            '#7139 It has to consider BusinessId value which is none value  then write InvAmount     By Kin 2015/11/10
            If DBNull.Value.Equals(dtInv007.Rows(0).Item("BusinessId")) Then
                XmlFile.WriteValue(Math.Round(Double.Parse(dtInv007.Rows(0).Item("InvAmount").ToString), 0))
            Else
                XmlFile.WriteValue(Math.Round(Double.Parse(dtInv007.Rows(0).Item("SaleAmount").ToString), 0))
            End If

            XmlFile.WriteEndElement() 'SalesAmount
            '----------------------------------------------------------------------------------------------------
            '------------------------------------FreeTaxSalesAmount-----------------------------------
            XmlFile.WriteStartElement("FreeTaxSalesAmount")
            XmlFile.WriteValue("0")
            XmlFile.WriteEndElement() 'FreeTaxSalesAmount
            '---------------------------------------------------------------------------------------------------
            '------------------------------------ZeroTaxSalesAmount-----------------------------------
            XmlFile.WriteStartElement("ZeroTaxSalesAmount")
            XmlFile.WriteValue("0")
            XmlFile.WriteEndElement() 'ZeroTaxSalesAmount
            '---------------------------------------------------------------------------------------------------
            '-----------------------------------TaxType----------------------------------------------------
            XmlFile.WriteStartElement("TaxType")
            XmlFile.WriteValue(dtInv007.Rows(0).Item("TaxType").ToString)
            XmlFile.WriteEndElement() 'TaxType
            '---------------------------------------------------------------------------------------------------
            '-----------------------------------TaxRate-----------------------------------------------------
            XmlFile.WriteStartElement("TaxRate")
            If Double.Parse(dtInv007.Rows(0).Item("TaxRate")) > 0 Then
                XmlFile.WriteValue(Double.Parse(dtInv007.Rows(0).Item("TaxRate")) / 100)
            Else
                XmlFile.WriteValue("0")
            End If
            XmlFile.WriteEndElement() 'TaxRate
            '---------------------------------------------------------------------------------------------------
            '----------------------------------TaxAmount--------------------------------------------------
            XmlFile.WriteStartElement("TaxAmount")
            If DBNull.Value.Equals(dtInv007.Rows(0).Item("BusinessId")) Then
                XmlFile.WriteValue("0")
            Else
                XmlFile.WriteValue(Math.Round(Double.Parse(dtInv007.Rows(0).Item("TaxAmount")), 0))
            End If
            XmlFile.WriteEndElement() 'TaxAmount
            '----------------------------------------------------------------------------------------------------
            '----------------------------------TotalAmount----------------------------------------------------
            XmlFile.WriteStartElement("TotalAmount")
            XmlFile.WriteValue(dtInv007.Rows(0).Item("InvAmount").ToString)
            XmlFile.WriteEndElement() 'TotalAmount
            '-------------------------------------------------------------------------------------------------------         
            XmlFile.WriteEndElement() 'End Amount
            '---------------------------------------------------------------------------------------------
            XmlFile.WriteEndElement() 'Invoice
            If AddDataToLogData(dtInv007,
                             rwInv001.Item("CompId"), InvTypeEnum.InvFileType.C0401) Then
                WriteUploadLog(LogData, InvTypeEnum.InvFileType.C0401)
            End If
        Catch ex As Exception
            Throw New Exception("Function: WriteC0401 訊息: " & ex.ToString)

        End Try
        Return True
    End Function
    Private Function WriteC0501(ByVal dtInv007 As DataTable, ByVal rwInv001 As DataRow, ByVal XmlFile As Xml.XmlWriter) As Boolean
        Try
            CreateXMLFile(XmlFile)
            XmlFile.WriteStartElement("CancelInvoice")
            XmlFile.WriteAttributeString("xmlns", "urn:GEINV:eInvoiceMessage:C0501:3.1")
            XmlFile.WriteAttributeString("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance")
            XmlFile.WriteAttributeString("xsi:schemaLocation", "urn:GEINV:eInvoiceMessage:C0501: 3.1 C0501.xsd")
            '--------------------------------------------------------------------------
            '---------------------------CancelInvoiceNumber-------------------
            XmlFile.WriteStartElement("CancelInvoiceNumber")
            XmlFile.WriteValue(dtInv007.Rows(0).Item("InvId").ToString)
            XmlFile.WriteEndElement() 'CancelInvoiceNumber
            '---------------------------------------------------------------------------
            '---------------------------InvoiceDate--------------------------------
            XmlFile.WriteStartElement("InvoiceDate")
            XmlFile.WriteValue(Date.Parse(dtInv007.Rows(0).Item("InvDate").ToString).ToString("yyyyMMdd"))
            XmlFile.WriteEndElement() 'InvoiceDate
            '----------------------------------------------------------------------------
            '---------------------------BuyerId-----------------------------------------
            XmlFile.WriteStartElement("BuyerId")
            XmlFile.WriteValue("0000000000")
            XmlFile.WriteEndElement() 'BuyerId
            '-------------------------------------------------------------------------------
            '----------------------------SellerId------------------------------------------
            XmlFile.WriteStartElement("SellerId")
            XmlFile.WriteValue(rwInv001.Item("BusinessId").ToString)
            XmlFile.WriteEndElement() 'SellerId
            '---------------------------------------------------------------------------------
            '-------------------------------CancelDate-----------------------------------
            XmlFile.WriteStartElement("CancelDate")
            XmlFile.WriteValue(Date.Parse(dtInv007.Rows(0).Item("UptTime").ToString).ToString("yyyyMMdd"))
            XmlFile.WriteEndElement() 'CancelDate
            '--------------------------------------------------------------------------------
            '------------------------------------CancelTime------------------------------
            XmlFile.WriteStartElement("CancelTime")
            XmlFile.WriteValue(Date.Parse(dtInv007.Rows(0).Item("UptTime").ToString).ToString("HH:mm:ss"))
            XmlFile.WriteEndElement() 'CancelTime
            '----------------------------------------------------------------------------------
            '-----------------------------------CancelReason----------------------------------
            XmlFile.WriteStartElement("CancelReason")
            XmlFile.WriteValue(dtInv007.Rows(0).Item("ObsoleteReason").ToString)
            XmlFile.WriteEndElement()
            '---------------------------------------------------------------------------------------
            XmlFile.WriteEndElement() 'CancelInvoice
            If AddDataToLogData(dtInv007,
                             rwInv001.Item("CompId"), InvTypeEnum.InvFileType.C0501) Then
                WriteUploadLog(LogData, InvTypeEnum.InvFileType.C0501)
            End If
        Catch ex As Exception
            Throw New Exception("Function: WriteC0401 訊息: " & ex.ToString)
        End Try
        Return True
    End Function
    Private Function WriteD0401(ByVal dtInv014 As DataTable, ByVal rwInv001 As DataRow, ByVal XmlFile As Xml.XmlWriter) As Boolean
        Try
            CreateXMLFile(XmlFile)
            XmlFile.WriteStartElement("Allowance")
            XmlFile.WriteAttributeString("xmlns", "urn:GEINV:eInvoiceMessage:D0401:3.1")
            XmlFile.WriteAttributeString("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance")
            XmlFile.WriteAttributeString("xsi:schemaLocation", "urn:GEINV:eInvoiceMessage:D0401: 3.1 D0401.xsd")
            '-------------------------Main--------------------------------------------------------------
            XmlFile.WriteStartElement("Main")
            '-------------------------AllowanceNumber---------------------------------------------
            XmlFile.WriteStartElement("AllowanceNumber")
            XmlFile.WriteValue(dtInv014.Rows(0).Item("PaperNo").ToString)
            XmlFile.WriteEndElement()
            '--------------------------------------------------------------------------------------------
            '-------------------------AllowanceDate-------------------------------------------------
            XmlFile.WriteStartElement("AllowanceDate")
            XmlFile.WriteValue(Date.Parse(dtInv014.Rows(0).Item("PaperDate").ToString).ToString("yyyyMMdd"))
            XmlFile.WriteEndElement() 'AllowanceDate
            '---------------------------------------------------------------------------------------------
            '----------------------------Seller & Buyer----------------------------------------
            For i As Int32 = 0 To 1
                If i = 0 Then
                    XmlFile.WriteStartElement("Seller")
                Else
                    XmlFile.WriteStartElement("Buyer")
                End If
                '----------------------------------Identifer----------------------------------------------------------
                If i = 0 Then
                    'Seller
                    XmlFile.WriteStartElement("Identifier")
                    If Not DBNull.Value.Equals(rwInv001.Item("BusinessId")) Then
                        XmlFile.WriteValue(rwInv001.Item("BusinessId").ToString)
                    End If
                    XmlFile.WriteEndElement() 'Identifier
                Else
                    'Buyer
                    XmlFile.WriteStartElement("Identifier")
                    If DBNull.Value.Equals(dtInv014.Rows(0).Item("BusinessId2")) Then
                        XmlFile.WriteValue("0000000000")
                    Else
                        XmlFile.WriteValue(dtInv014.Rows(0).Item("BusinessId2").ToString)
                    End If
                    XmlFile.WriteEndElement() 'Identifier
                End If
                '---------------------------------------------------------------------------------------------------------
                '---------------------------------Name------------------------------------------------------------------
                XmlFile.WriteStartElement("Name")
                If i = 0 Then
                    'Seller
                    XmlFile.WriteValue(rwInv001.Item("CompName").ToString)
                Else
                    'Buyer
                    If DBNull.Value.Equals(dtInv014.Rows(0).Item("BusinessId2")) Then
                        XmlFile.WriteValue(Right(dtInv014.Rows(0).Item("InvId").ToString, 4))
                    Else
                        If Not DBNull.Value.Equals(dtInv014.Rows(0).Item("InvTitle").ToString) Then
                            XmlFile.WriteValue(dtInv014.Rows(0).Item("InvTitle").ToString)
                        End If
                    End If
                End If
                XmlFile.WriteEndElement() 'Name
                '-------------------------------------------------------------------------------------------------------------
                '-------------------------------Address-----------------------------------------------------------------------
                'Seller need to  add  Address element  For Jacy 
                If i = 0 Then 'Seller 才要 For Jacy
                    XmlFile.WriteStartElement("Address")
                    XmlFile.WriteValue(rwInv001.Item("CompAddr").ToString)
                    XmlFile.WriteEndElement() 'Address
                End If
                '-------------------------------------------------------------------------------------------------------------
                XmlFile.WriteEndElement() 'Seller & Buyer
            Next
            '-----------------------------------------------------------------------------------------------------------------
            '-------------------------------------AllowanceType---------------------------------------------------------
            XmlFile.WriteStartElement("AllowanceType")
            XmlFile.WriteValue("2")
            XmlFile.WriteEndElement() 'AllowanceType
            '------------------------------------------------------------------------------------------------------------------
            XmlFile.WriteEndElement() 'Main
            '-------------------------------------Details----------------------------------------------------------------------
            XmlFile.WriteStartElement("Details")
            For Each rw As DataRow In dtInv014.Rows
                '-------------------------------------ProductItem--------------------------------------------------------------
                XmlFile.WriteStartElement("ProductItem")
                '------------------------------------OriginalInvoiceDate-------------------------------------------------------
                XmlFile.WriteStartElement("OriginalInvoiceDate")
                XmlFile.WriteValue(Date.Parse(rw.Item("InvDate").ToString).ToString("yyyyMMdd"))
                XmlFile.WriteEndElement() 'OriginalInvoiceDate
                '--------------------------------------------------------------------------------------------------------------------
                '-----------------------------------OriginalInvoiceNumber---------------------------------------------------
                XmlFile.WriteStartElement("OriginalInvoiceNumber")
                XmlFile.WriteValue(rw.Item("InvId").ToString)
                XmlFile.WriteEndElement() 'OriginalInvoiceNumber
                '-------------------------------------------------------------------------------------------------------------------
                '-----------------------------------OriginalDescription---------------------------------------------------------
                XmlFile.WriteStartElement("OriginalDescription")
                XmlFile.WriteValue(rw.Item("Description").ToString)
                XmlFile.WriteEndElement() 'OriginalDescription
                '-------------------------------------------------------------------------------------------------------------------
                '-----------------------------------Quantity--------------------------------------------------------------------
                XmlFile.WriteStartElement("Quantity")
                XmlFile.WriteValue(Int32.Parse(rw.Item("Quantity").ToString))
                XmlFile.WriteEndElement() 'Quantity
                '------------------------------------------------------------------------------------------------------------------
                '-----------------------------------UnitPrice--------------------------------------------------------------------
                XmlFile.WriteStartElement("UnitPrice")
                XmlFile.WriteValue(Math.Round(
                                   Double.Parse(rw.Item("SaleAmount").ToString) / Double.Parse(rw.Item("Quantity").ToString), 0))
                XmlFile.WriteEndElement() 'UnitPrice
                '------------------------------------------------------------------------------------------------------------------
                '----------------------------------Amount-------------------------------------------------------------------------
                XmlFile.WriteStartElement("Amount")
                XmlFile.WriteValue(Decimal.Parse(rw.Item("SaleAmount").ToString))
                XmlFile.WriteEndElement() 'Amount
                '-----------------------------------------------------------------------------------------------------------------------
                '---------------------------------Tax---------------------------------------------------------------------------------
                XmlFile.WriteStartElement("Tax")
                XmlFile.WriteValue(Math.Round(Decimal.Parse(rw.Item("TaxAmount")), 0))
                XmlFile.WriteEndElement() 'Tax
                '------------------------------------------------------------------------------------------------------------------------
                '---------------------------------AllowanceSequenceNumber---------------------------------------------------
                XmlFile.WriteStartElement("AllowanceSequenceNumber")
                XmlFile.WriteValue(1)
                XmlFile.WriteEndElement() 'AllowanceSequenceNumber
                '------------------------------------------------------------------------------------------------------------------------
                '---------------------------------TaxType--------------------------------------------------------------------------
                XmlFile.WriteStartElement("TaxType")
                XmlFile.WriteValue(Int32.Parse(rw.Item("TaxType").ToString))
                XmlFile.WriteEndElement() 'TaxType
                '------------------------------------------------------------------------------------------------------------------------
                XmlFile.WriteEndElement() 'ProductItem
                '--------------------------------------------------------------------------------------------------------------------
            Next
            XmlFile.WriteEndElement() 'Details
            '------------------------------------End Amount---------------------------------------------------------------------------
            XmlFile.WriteStartElement("Amount")
            '-----------------------------------TaxAmount---------------------------------------------------------------------
            XmlFile.WriteStartElement("TaxAmount")
            XmlFile.WriteValue(dtInv014.Rows(0).Item("TaxAmount").ToString)
            XmlFile.WriteEndElement() 'TaxAmount
            '-------------------------------------------------------------------------------------------------------------------------
            '-----------------------------------TotalAmount--------------------------------------------------------------------
            XmlFile.WriteStartElement("TotalAmount")
            XmlFile.WriteValue(dtInv014.Rows(0).Item("SaleAmount").ToString)
            XmlFile.WriteEndElement() 'TotalAmount
            '-------------------------------------------------------------------------------------------------------------------------
            XmlFile.WriteEndElement() 'End Amount
            '-------------------------------------------------------------------------------------------------------------------
            XmlFile.WriteEndElement() 'Allowance
            If AddDataToLogData(dtInv014,
                             rwInv001.Item("CompId"), InvTypeEnum.InvFileType.D0401) Then
                WriteUploadLog(LogData, InvTypeEnum.InvFileType.D0401)
            End If
        Catch ex As Exception
            Throw New Exception("Function: WriteD0401 訊息: " & ex.ToString)
        End Try
        Return True
    End Function
    Private Function WriteD0501(ByVal dtInv014 As DataTable, ByVal rwInv001 As DataRow, ByVal XmlFile As Xml.XmlWriter) As Boolean
        Try
            CreateXMLFile(XmlFile)
            XmlFile.WriteStartElement("CancelAllowance")
            XmlFile.WriteAttributeString("xmlns", "urn:GEINV:eInvoiceMessage:D0501:3.1")
            XmlFile.WriteAttributeString("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance")
            XmlFile.WriteAttributeString("xsi:schemaLocation", "urn:GEINV:eInvoiceMessage:D0501:3.1 ../XSD/D0501.xsd")
            '------------------------------------------CancelAllowanceNumber-----------------------------------------
            XmlFile.WriteStartElement("CancelAllowanceNumber")
            XmlFile.WriteValue(dtInv014.Rows(0).Item("PaperNo").ToString)
            XmlFile.WriteEndElement()
            '-------------------------------------------------------------------------------------------------------------------
            '------------------------------------------AllowanceDate-------------------------------------------------------
            XmlFile.WriteStartElement("AllowanceDate")
            XmlFile.WriteValue(Date.Parse(dtInv014.Rows(0).Item("PaperDate").ToString).ToString("yyyyMMdd"))
            XmlFile.WriteEndElement() 'AllowanceDate
            '-------------------------------------------------------------------------------------------------------------------
            '-----------------------------------------BuyerId----------------------------------------------------------------
            XmlFile.WriteStartElement("BuyerId")
            '#6815 如果INV014.BusinessId有值,則要填INV014的統編 By Kin 2014/06/25
            If Not DBNull.Value.Equals(dtInv014.Rows(0).Item("BusinessId")) Then
                XmlFile.WriteValue(dtInv014.Rows(0).Item("BusinessId").ToString)
            Else
                XmlFile.WriteValue("0000000000")
            End If

            XmlFile.WriteEndElement() 'BuyerId
            '-------------------------------------------------------------------------------------------------------------------
            '-----------------------------------------SellerId-----------------------------------------------------------------
            XmlFile.WriteStartElement("SellerId")
            XmlFile.WriteValue(rwInv001.Item("BusinessId").ToString)
            XmlFile.WriteEndElement() 'SellerId
            '-------------------------------------------------------------------------------------------------------------------
            '-----------------------------------------CancelDate-----------------------------------------------------------
            XmlFile.WriteStartElement("CancelDate")
            XmlFile.WriteValue(Date.Parse(dtInv014.Rows(0).Item("UptTime").ToString).ToString("yyyyMMdd"))
            XmlFile.WriteEndElement() 'CancelDate
            '-------------------------------------------------------------------------------------------------------------------
            '-----------------------------------------CancelTime-----------------------------------------------------------
            XmlFile.WriteStartElement("CancelTime")
            XmlFile.WriteValue(Date.Parse(dtInv014.Rows(0).Item("UptTime").ToString).ToString("HH:mm:ss"))
            XmlFile.WriteEndElement() 'CancelTime
            '-------------------------------------------------------------------------------------------------------------------
            '-----------------------------------------CancelTime-----------------------------------------------------------
            XmlFile.WriteStartElement("CancelReason")
            XmlFile.WriteValue(dtInv014.Rows(0).Item("ObsoleteReason").ToString)
            XmlFile.WriteEndElement() 'CancelReason
            '-------------------------------------------------------------------------------------------------------------------
            XmlFile.WriteEndElement() 'CancelAllowance
            If AddDataToLogData(dtInv014,
                            rwInv001.Item("CompId"), InvTypeEnum.InvFileType.D0501) Then
                WriteUploadLog(LogData, InvTypeEnum.InvFileType.D0501)
            End If
        Catch ex As Exception
            Throw New Exception("Function: WriteD0501 訊息: " & ex.ToString)
        End Try
        Return True
    End Function
    Private Overloads Function WriteE0402(ByVal dt099 As DataTable, ByVal rwInv001 As DataRow, ByVal XmlFile As Xml.XmlWriter) As Boolean
        Try

            CreateXMLFile(XmlFile)

            XmlFile.WriteStartElement("BranchTrackBlank")           'BranchTrackBlank
            XmlFile.WriteAttributeString("xmlns", "urn:GEINV:eInvoiceMessage:E0402:3.1")
            XmlFile.WriteAttributeString("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance")
            XmlFile.WriteAttributeString("xsi:schemaLocation", "urn:GEINV:eInvoiceMessage:E0402: 3.1 E0402.xsd")
            '-------------------------Main--------------------------------------------------------------
            XmlFile.WriteStartElement("Main")
            '-------------------------HeadBan----------------------------------------------------------
            XmlFile.WriteStartElement("HeadBan")
            '#6791 總公司統編有值就用總公司統編 By Kin 2014/06/04
            If Not DBNull.Value.Equals(rwInv001.Item("MainBusinessId")) Then
                XmlFile.WriteValue(rwInv001.Item("MainBusinessId").ToString)
            Else
                XmlFile.WriteValue(rwInv001.Item("BusinessId").ToString)
            End If
            XmlFile.WriteEndElement()       'HeadBan
            '---------------------------------------------------------------------------------------------
            '-------------------------BranchBan----------------------------------------------------------
            XmlFile.WriteStartElement("BranchBan")
            XmlFile.WriteValue(rwInv001.Item("BusinessId").ToString)
            XmlFile.WriteEndElement()       'BranchBan
            '---------------------------------------------------------------------------------------------
            '-------------------------InvoiceType----------------------------------------------------------
            XmlFile.WriteStartElement("InvoiceType")
            '#7001 改用Inv001.InvoiceType By Kin 2015/04/02
            'XmlFile.WriteValue("05")
            XmlFile.WriteValue(rwInv001("InvoiceType").ToString)
            XmlFile.WriteEndElement()       'InvoiceType
            '---------------------------------------------------------------------------------------------
            '-------------------------YearMonth----------------------------------------------------------
            XmlFile.WriteStartElement("YearMonth")
            XmlFile.WriteValue(((Integer.Parse(dt099.Rows(0).Item("YearMonth")) - 191100) + 1).ToString)
            XmlFile.WriteEndElement()       'YearMonth
            '---------------------------------------------------------------------------------------------
            '-------------------------InvoiceTrack----------------------------------------------------------
            XmlFile.WriteStartElement("InvoiceTrack")
            XmlFile.WriteValue(dt099.Rows(0).Item("Prefix").ToString.ToUpper)
            XmlFile.WriteEndElement()       'InvoiceTrack
            '---------------------------------------------------------------------------------------------
            '--------------------------------------------------------------------------------------------
            XmlFile.WriteEndElement() 'Main
            '-------------------------Details--------------------------------------------------------------
            XmlFile.WriteStartElement("Details")
            '------------------------BranchTrackBlankItem-------------------------------------------
            For Each rw099 As DataRow In dt099.Rows
                XmlFile.WriteStartElement("BranchTrackBlankItem")
                '----------------------------InvoiceBeginNo----------------------------------------------
                XmlFile.WriteStartElement("InvoiceBeginNo")
                XmlFile.WriteValue(rw099.Item("CurNum").ToString)
                XmlFile.WriteEndElement() 'InvoiceBeginNo
                '----------------------------InvoiceEndNo------------------------------------------------
                XmlFile.WriteStartElement("InvoiceEndNo")
                XmlFile.WriteValue(rw099.Item("EndNum").ToString)
                XmlFile.WriteEndElement() 'InvoiceEndNo
                '--------------------------------------------------------------------------------------------
                XmlFile.WriteEndElement() 'BranchTrackBlankItem
            Next
            XmlFile.WriteEndElement() 'Details
            XmlFile.WriteEndElement()   'BranchTrackBlank

            'If AddDataToLogData(dt099,
            '                rwInv001.Item("CompId"), InvTypeEnum.InvFileType.E0402) Then
            '    WriteUploadLog(LogData, InvTypeEnum.InvFileType.E0402)
            'End If

            If AddDataToLogData(dt099,
                            rwInv001.Item("Companies"), InvTypeEnum.InvFileType.E0402) Then
                WriteUploadLog(LogData, InvTypeEnum.InvFileType.E0402)
            End If
        Catch ex As Exception
            Throw New Exception("Function: WriteE0402 訊息: " & ex.ToString)
        End Try
        Return True
    End Function

    Private Sub WriteUploadLog(ByVal WriteData As Dictionary(Of String, Object),
                               ByVal InvFileType As CableSoft.B07.InvType.InvTypeEnum.InvFileType)
        Dim tra As OracleTransaction = Nothing
        Try
            Using cnLog As New OracleConnection(InvArgument.GetConnect)
                cnLog.Open()
                tra = cnLog.BeginTransaction
                Using cmd As OracleCommand = cnLog.CreateCommand
                    cmd.Transaction = tra
                    cmd.Parameters.Add("InvId", OracleType.VarChar).Value = WriteData.Item("InvId")
                    cmd.Parameters.Add("UploadType", OracleType.VarChar).Value = WriteData.Item("UploadType")
                    cmd.Parameters.Add("UploadTime", OracleType.DateTime).Value = WriteData.Item("UploadTime")
                    cmd.Parameters.Add("UploadSource", OracleType.Number).Value = WriteData.Item("UploadSource")
                    cmd.Parameters.Add("CompId", OracleType.VarChar).Value = WriteData.Item("CompId")
                    cmd.CommandText = String.Format("Insert Into Inv048 ({0},{1},{2},{3},{4}) " & _
                                                  " Values ({5},{6},{7},{8},{9})",
                                                  "InvId", "UploadType", "UploadTime", "UploadSource", "CompId",
                                                  ":InvId", ":UploadType", ":UploadTime", ":UploadSource", ":CompId")
                    cmd.ExecuteNonQuery()
                    cmd.Parameters.Clear()
                    Select Case InvFileType
                        Case InvTypeEnum.InvFileType.C0701
                            cmd.Parameters.Clear()
                            cmd.Parameters.Add("DESTROYFLAG", OracleType.Number).Value = 1
                            cmd.Parameters.Add("DESTROYUPLOADTIME", OracleType.DateTime).Value = WriteData.Item("UploadTime")
                            cmd.Parameters.Add("DESTROYREASON", OracleType.VarChar).Value = Me.DestroyReason
                            cmd.Parameters.Add("INVID", OracleType.VarChar).Value = WriteData.Item("INVID").ToString
                            cmd.Parameters.Add("COMPID", OracleType.VarChar).Value = WriteData.Item("COMPID").ToString
                            cmd.Parameters.Add("AUTOUPLOADFLAG", OracleType.Number).Value = WriteData.Item("UploadSource")
                            cmd.CommandText = String.Format("UPDATE INV007  SET  DESTROYFLAG = {0}, " & _
                                               "DESTROYUPLOADTIME = {1},DESTROYREASON = {2},AUTOUPLOADFLAG = {3} " & _
                                               " WHERE ISOBSOLETE <> 'Y' AND UPLOADTIME IS NOT NULL " & _
                                               " AND INVID = {4} AND COMPID = {5}",
                                                ":DESTROYFLAG", ":DESTROYUPLOADTIME",
                                                ":DESTROYREASON", ":AUTOUPLOADFLAG",
                                                ":INVID", ":COMPID")
                            cmd.ExecuteNonQuery()
                            cmd.Parameters.Clear()
                        Case InvTypeEnum.InvFileType.A0401, InvTypeEnum.InvFileType.C0401
                            cmd.Parameters.Clear()
                            cmd.Parameters.Add("UPLOADFLAG", OracleType.Number).Value = 1
                            cmd.Parameters.Add("UPLOADTIME", OracleType.DateTime).Value = WriteData.Item("UPLOADTIME")
                            cmd.Parameters.Add("INVID", OracleType.VarChar).Value = WriteData.Item("INVID").ToString
                            cmd.Parameters.Add("COMPID", OracleType.VarChar).Value = WriteData.Item("COMPID").ToString
                            cmd.Parameters.Add("AUTOUPLOADFLAG", OracleType.Number).Value = WriteData.Item("UploadSource")
                            cmd.CommandText = String.Format("UPDATE INV007 SET UPLOADFLAG = {0}, " & _
                                            "UPLOADTIME = {1},AUTOUPLOADFLAG = {2}  " & _
                                            " WHERE INVID = {3} AND COMPID = {4}",
                                            ":UPLOADFLAG", ":UPLOADTIME",
                                            ":AUTOUPLOADFLAG", ":INVID", ":COMPID")
                            cmd.ExecuteNonQuery()
                            cmd.Parameters.Clear()
                        Case InvTypeEnum.InvFileType.A0501, InvTypeEnum.InvFileType.C0501
                            cmd.Parameters.Clear()
                            cmd.Parameters.Add("OBUPLOADFLAG", OracleType.Number).Value = 1
                            cmd.Parameters.Add("OBUPLOADTIME", OracleType.DateTime).Value = WriteData.Item("UPLOADTIME")
                            cmd.Parameters.Add("INVID", OracleType.VarChar).Value = WriteData.Item("INVID").ToString
                            cmd.Parameters.Add("COMPID", OracleType.VarChar).Value = WriteData.Item("COMPID").ToString
                            cmd.Parameters.Add("AUTOUPLOADFLAG", OracleType.Number).Value = WriteData.Item("UploadSource")
                            cmd.CommandText = String.Format("UPDATE INV007 SET OBUPLOADFLAG = {0}, " & _
                                            "OBUPLOADTIME = {1},AUTOUPLOADFLAG = {2} " & _
                                            " WHERE INVID = {3} AND COMPID = {4}",
                                            ":OBUPLOADFLAG",
                                            ":OBUPLOADTIME",
                                            ":AUTOUPLOADFLAG",
                                            ":INVID",
                                            ":COMPID")
                            cmd.ExecuteNonQuery()
                            cmd.Parameters.Clear()
                        Case InvTypeEnum.InvFileType.B0301, InvTypeEnum.InvFileType.D0401
                            cmd.Parameters.Clear()
                            cmd.Parameters.Add("UPLOADFLAG", OracleType.Number).Value = 1
                            cmd.Parameters.Add("UPLOADTIME", OracleType.DateTime).Value = WriteData.Item("UPLOADTIME")
                            cmd.Parameters.Add("ALLOWANCENO", OracleType.VarChar).Value = WriteData.Item("INVID").ToString
                            cmd.Parameters.Add("COMPID", OracleType.VarChar).Value = WriteData.Item("COMPID").ToString
                            cmd.Parameters.Add("AUTOUPLOADFLAG", OracleType.Number).Value = WriteData.Item("UploadSource")
                            cmd.CommandText = String.Format("UPDATE INV014 SET UPLOADFLAG = {0} , " & _
                                                          "UPLOADTIME = {1},AUTOUPLOADFLAG = {2} " & _
                                                          " WHERE ALLOWANCENO = {3} AND COMPID = {4} ",
                                                          ":UPLOADFLAG",
                                                          ":UPLOADTIME",
                                                          ":AUTOUPLOADFLAG",
                                                          ":ALLOWANCENO",
                                                          ":COMPID")
                            cmd.ExecuteNonQuery()
                            cmd.Parameters.Clear()
                        Case InvTypeEnum.InvFileType.B0401, InvTypeEnum.InvFileType.D0501
                            cmd.Parameters.Clear()
                            cmd.Parameters.Add("OBUPLOADFLAG", OracleType.Number).Value = 1
                            cmd.Parameters.Add("OBUPLOADTIME", OracleType.DateTime).Value = WriteData.Item("UPLOADTIME")
                            cmd.Parameters.Add("ALLOWANCENO", OracleType.VarChar).Value = WriteData.Item("INVID").ToString
                            cmd.Parameters.Add("COMPID", OracleType.VarChar).Value = WriteData.Item("COMPID").ToString
                            cmd.Parameters.Add("AUTOUPLOADFLAG", OracleType.Number).Value = WriteData.Item("UploadSource")
                            cmd.CommandText = String.Format("UPDATE INV014 SET OBUPLOADFLAG = {0} , " & _
                                                          "OBUPLOADTIME = {1},AUTOUPLOADFLAG = {2} " & _
                                                          " WHERE ALLOWANCENO = {3} AND COMPID = {4} ",
                                                          ":OBUPLOADFLAG",
                                                          ":OBUPLOADTIME",
                                                          ":AUTOUPLOADFLAG",
                                                          ":ALLOWANCENO",
                                                          ":COMPID")
                            cmd.ExecuteNonQuery()
                            cmd.Parameters.Clear()
                    End Select
                End Using
                tra.Commit()

            End Using
        Catch ex As Exception
            Throw
        Finally
            If tra IsNot Nothing Then
                tra.Dispose()
                tra = Nothing
            End If
        End Try
    End Sub
    Private Overloads Function WriteC0701(ByVal dt007 As DataTable,
                                          ByVal rwInv001 As DataRow, ByVal XmlFile As Xml.XmlWriter) As Boolean
        Try
            CreateXMLFile(XmlFile)
            XmlFile.WriteStartElement("VoidInvoice")           'BranchTrackBlank
            XmlFile.WriteAttributeString("xmlns", "urn:GEINV:eInvoiceMessage:C0701:3.1")
            XmlFile.WriteAttributeString("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance")
            XmlFile.WriteAttributeString("xsi:schemaLocation", "urn:GEINV:eInvoiceMessage:C0701: 3.1 C0701.xsd")
            
            '-------------------------VoidInvoiceNumber----------------------------------------------------------
            XmlFile.WriteStartElement("VoidInvoiceNumber")
            XmlFile.WriteValue(dt007.Rows(0).Item("InvId").ToString)
            XmlFile.WriteEndElement()       'VoidInvoiceNumber
            '---------------------------------------------------------------------------------------------
            '-------------------------InvoiceDate----------------------------------------------------------
            XmlFile.WriteStartElement("InvoiceDate")
            XmlFile.WriteValue(DateTime.Parse(dt007.Rows(0).Item("InvDate").ToString).ToString("yyyyMMdd"))
            XmlFile.WriteEndElement()       'InvoiceDate
            '---------------------------------------------------------------------------------------------

            '-------------------------BuyerId----------------------------------------------------------
            XmlFile.WriteStartElement("BuyerId")
            If Not DBNull.Value.Equals(dt007.Rows(0).Item("BusinessId")) Then
                XmlFile.WriteValue(dt007.Rows(0).Item("BusinessId").ToString)
            Else
                XmlFile.WriteValue(New String("0", 10))
            End If
            XmlFile.WriteEndElement()       'BuyerId
            '------------------------------------------------------------------------------------------------

            '-------------------------SellerId----------------------------------------------------------
            XmlFile.WriteStartElement("SellerId")
            XmlFile.WriteValue(rwInv001.Item("BusinessId").ToString)
            XmlFile.WriteEndElement()   'SellerId
            '------------------------------------------------------------------------------------------------

            '-------------------------VoidDate----------------------------------------------------------
            XmlFile.WriteStartElement("VoidDate")
            XmlFile.WriteValue(FNow.ToString("yyyyMMdd"))
            XmlFile.WriteEndElement()   'VoidDate
            '------------------------------------------------------------------------------------------------

            '-------------------------VoidTime----------------------------------------------------------
            XmlFile.WriteStartElement("VoidTime")
            XmlFile.WriteValue(FNow.ToString("HH:mm:ss"))
            XmlFile.WriteEndElement()   'VoidTime
            '------------------------------------------------------------------------------------------------

            '-------------------------VoidReason----------------------------------------------------------
            XmlFile.WriteStartElement("VoidReason")
            XmlFile.WriteValue(Me.DestroyReason)
            XmlFile.WriteEndElement()   'VoidReason
            '------------------------------------------------------------------------------------------------

            XmlFile.WriteEndElement()   'VoidInvoice
            If AddDataToLogData(dt007,
                            rwInv001.Item("CompId"), InvTypeEnum.InvFileType.C0701) Then
                WriteUploadLog(LogData, InvTypeEnum.InvFileType.C0701)
            End If
        Catch ex As Exception
            Throw New Exception("Function: WriteC0701 訊息: " & ex.ToString)
        End Try
        Return True
        Return True
    End Function
    Private Overloads Function WriteC0701() As Boolean
        Dim dtXml As DataTable = Nothing
        Dim aTblInv007 As DataTable = FInvDs.Tables("INV007").Copy
        Dim rowInv001 As DataRow = Nothing
        Dim aFilePath As String = Nothing
        If aTblInv007.Rows.Count = 0 Then
            Return True
        Else
            dtXml = New DataTable            
            For Each col As DataColumn In FInvDs.Tables("INV007").Columns
                dtXml.Columns.Add(col.ColumnName, col.DataType)
            Next
        End If
        If Not String.IsNullOrEmpty(ExportElectronPath) Then           
            aFilePath = ExportElectronPath
            If Not aFilePath.EndsWith("\") Then
                aFilePath = aFilePath & "\"
            End If
        Else
            aFilePath = CableSoft.B07.SystemIO.Invoice.GetINVPathName(InvArgument.GetConnect,
                                                                                InvArgument.GetCompCode, fC0701FilePath)
        End If
         
        If String.IsNullOrEmpty(aFilePath) Then
            Throw New Exception("C0701 檔案路徑設定有誤！")
        End If
        Try
            Dim lstRw As New List(Of DataRow)
            Dim strInvNo As String = String.Empty
            For Each rw As DataRow In aTblInv007.Rows
                If rw.Item("InvId").ToString <> strInvNo Then
                    lstRw.Add(rw)
                End If
                strInvNo = rw.Item("InvId").ToString
            Next
            For Each rw As DataRow In lstRw
                dtXml.Rows.Clear()
                rowInv001 = FInvDs.Tables("INV001").Select("COMPID='" & rw("COMPID") & "'").GetValue(0)
                Dim rwNew As DataRow = dtXml.NewRow
                rwNew.ItemArray = rw.ItemArray
                dtXml.Rows.Add(rwNew)
                dtXml.AcceptChanges()
                WriteToXml(aFilePath, dtXml, rowInv001, InvTypeEnum.InvFileType.C0701)
            Next
            _C0701RcdCount = lstRw.Count
        Catch ex As Exception
            Throw New Exception("Function: WriteC0701 訊息: " & ex.ToString)
        End Try
        Return True
    End Function

    Private Overloads Function WriteE0402() As Boolean

        Dim dtXml As DataTable = Nothing
        Dim aTblINV099 As DataTable = FInvDs.Tables("INV099")
        Dim rowInv001 As DataRow = Nothing
        Dim aFilePath As String = Nothing
        If aTblINV099.Rows.Count <= 0 Then
            Return True
        Else
            dtXml = New DataTable
            For Each col As DataColumn In FInvDs.Tables("INV099").Columns
                dtXml.Columns.Add(col.ColumnName, col.DataType)
            Next
        End If
        If Not String.IsNullOrEmpty(ExportElectronPath) Then           
            aFilePath = ExportElectronPath
            If Not aFilePath.EndsWith("\") Then
                aFilePath = aFilePath & "\"
            End If
        Else
            aFilePath = CableSoft.B07.SystemIO.Invoice.GetINVPathName(InvArgument.GetConnect,
                                                                                InvArgument.GetCompCode, fE0402FilePath)
        End If
              


        If String.IsNullOrEmpty(aFilePath) Then
            Throw New Exception("E0402 檔案路徑設定有誤！")
        End If
        Dim FilteredInv099 As DataTable = GetSameMainSection(aTblINV099, FInvDs.Tables("INV001"))
        Try
            'E0402不用一個檔案一筆資料,可將資料合併成一個XML就好
            For Each rw001 As DataRow In FInvDs.Tables("INV001").Rows

                'Dim lstRw As List(Of DataRow) = aTblINV099.AsEnumerable.Where(Function(rw As DataRow)
                '                                                                  Return rw.Item("CompId") = rw001.Item("CompId")
                '                                                              End Function).ToList
                Dim lstRw As List(Of DataRow) = FilteredInv099.AsEnumerable.Where(Function(rw As DataRow)
                                                                                      Return rw.Item("CompId") = rw001.Item("CompId")
                                                                                  End Function).ToList
                dtXml.Rows.Clear()
                If lstRw IsNot Nothing AndAlso lstRw.Count > 0 Then
                    For Each rw As DataRow In lstRw
                        Dim rwNew As DataRow = dtXml.NewRow
                        rwNew.ItemArray = rw.ItemArray
                        dtXml.Rows.Add(rwNew)
                        dtXml.AcceptChanges()
                    Next
                    WriteToXml(aFilePath, dtXml, rw001, InvTypeEnum.InvFileType.E0402)
                End If
            Next
            _E0402RcdCount = aTblINV099.Rows.Count
            Return True
        Catch ex As Exception
            Throw New Exception("Function: WriteE0402 訊息: " & ex.ToString)
        Finally
            If FilteredInv099 IsNot Nothing Then
                FilteredInv099.Dispose()
                FilteredInv099 = Nothing
            End If
        End Try
    End Function
    Private Function GetSameMainSection(ByVal dtInv099 As DataTable, ByVal dtInv001 As DataTable) As DataTable        
        Dim return099 As DataTable = dtInv099.Copy
        'dtInv001Copy.Rows.Clear()
        'dtInv099Copy.Rows.Clear()
        Dim businessId As String = Nothing
        Dim YearMonthPrefix As String = Nothing
        dtInv001.Columns.Add("Companies", GetType(String))
        For i As Integer = 0 To dtInv001.Rows.Count - 1
            dtInv001.Rows(i).Item("Companies") = dtInv001.Rows(i).Item("CompID")
            dtInv001.AcceptChanges()
        Next

        For Each rwInv001 As DataRow In dtInv001.Rows
            businessId = Nothing
            YearMonthPrefix = Nothing
            If Not DBNull.Value.Equals(rwInv001("MainBusinessId")) Then
                businessId = rwInv001("MainBusinessId")
            End If
            If Not DBNull.Value.Equals(rwInv001("BusinessId")) AndAlso String.IsNullOrEmpty(businessId) Then
                businessId = businessId & rwInv001("BusinessId")
            End If
            Dim sameCompany As List(Of DataRow) = dtInv001.AsEnumerable.Where(Function(rw001 As DataRow)
                                                                                  If rw001("CompID") <> rwInv001("CompID") Then
                                                                                      If Not DBNull.Value.Equals(rw001("MainBusinessId")) Then
                                                                                          If businessId = rw001("MainBusinessId") Then Return True
                                                                                      Else
                                                                                          If Not DBNull.Value.Equals(rw001("BusinessId")) Then
                                                                                              If businessId = rw001("BusinessId") Then Return True
                                                                                          End If
                                                                                      End If
                                                                                  End If
                                                                                  Return False
                                                                              End Function).ToList
            Dim o As List(Of DataRow) = return099.AsEnumerable.Where(Function(rw099 As DataRow)
                                                                         If rw099("CompID") = rwInv001("CompID") Then
                                                                             Return True
                                                                         Else
                                                                             Return False
                                                                         End If
                                                                     End Function).ToList
            For Each specific099 As DataRow In o
                YearMonthPrefix = specific099.Item("YearMonth") & "" & specific099.Item("Prefix") & ""
                For Each same001 As DataRow In sameCompany
                    Dim same009 As List(Of DataRow) = return099.AsEnumerable.Where(Function(rw099 As DataRow)
                                                                                       If rw099("CompID") <> rwInv001("CompID") Then
                                                                                           Return rw099.Item("YearMonth") & "" & rw099.Item("Prefix") & "" = YearMonthPrefix
                                                                                       End If
                                                                                       Return False
                                                                                   End Function).ToList
                    If same009.Count > 0 Then
                        For Each delrow As DataRow In same009
                            Dim rwAdd As DataRow = return099.NewRow
                            rwAdd.ItemArray = delrow.ItemArray
                            rwAdd.Item("CompId") = rwInv001("CompID")
                            rwInv001.Item("Companies") = rwInv001.Item("Companies") & "," & delrow.Item("CompID")
                            return099.Rows.Remove(delrow)
                            return099.Rows.Add(rwAdd)
                            return099.AcceptChanges()
                        Next
                    End If
                Next
            Next
        Next
        Return return099

        'For Each rw As DataRow In dtInv001.Rows
        '    businessId = Nothing
        '    YearMonthPrefix = Nothing
        '    If Not DBNull.Value.Equals(rw("MainBusinessId")) Then
        '        businessId = rw("MainBusinessId")
        '    End If
        '    If Not DBNull.Value.Equals(rw("BusinessId")) Then
        '        businessId = businessId & rw("BusinessId")
        '    End If
        '    Dim o As DataRow = dtInv099.AsEnumerable.Single(Function(rw099 As DataRow)
        '                                                        Return rw099("YearMonth") & "" & rw099("Prefix") & ""
        '                                                    End Function)
        '    YearMonthPrefix = o.Item("YearMonth") & "" & o.Item("Prefix") & ""
        '    Dim sameRw As List(Of DataRow) = dtInv099.AsEnumerable.Where(Function(rw099 As DataRow)

        '                                                                     If (rw099("CompId") <> rw("CompId")) Then
        '                                                                         Return rw099("YearMonth") & "" & rw099("Prefix") & "" = YearMonthPrefix
        '                                                                     Else
        '                                                                         Return False
        '                                                                     End If
        '                                                                 End Function).ToList

        'Next


    End Function
    Private Function WriteB0301(ByVal UploadType As Integer) As Boolean
        Dim bduString As StringBuilder = Nothing
        Dim sInvId As String = String.Empty
        Dim aSeq As Integer = 1
        Dim strEnd As String = String.Empty
        Dim iCount As Integer = 0
        Dim dtXml As DataTable = Nothing
        Dim aTblINV014 As DataTable = FInvDs.Tables("INV014")
        Dim rowInv001 As DataRow = Nothing
        Dim aFilePath As String = Nothing
        If aTblINV014.Rows.Count <= 0 Then
            Return True
        Else
            If UploadType = 0 Then
                bduString = New StringBuilder()
            Else
                dtXml = New DataTable
                For Each col As DataColumn In FInvDs.Tables("INV014").Columns
                    dtXml.Columns.Add(col.ColumnName, col.DataType)
                Next
            End If
        End If
        If Not String.IsNullOrEmpty(ExportElectronPath) Then
            aFilePath = ExportElectronPath
            If Not aFilePath.EndsWith("\") Then
                aFilePath = aFilePath & "\"
            End If
        Else
            aFilePath = CableSoft.B07.SystemIO.Invoice.GetINVPathName(InvArgument.GetConnect,
                                                                                InvArgument.GetCompCode, fB0301FilePath)
        End If

        If String.IsNullOrEmpty(aFilePath) Then
            Throw New Exception("B0301 檔案路徑設定有誤！")
        End If
        Try
            For i As Int32 = 0 To aTblINV014.Rows.Count - 1
                '#5881 有可能是全部公司,所以INV001要Select By Kin 2010/12/14
                rowInv001 = FInvDs.Tables("INV001").Select("COMPID='" & aTblINV014.Rows(i)("COMPID") & "'").GetValue(0)
                '#6014 測試不OK，不要判斷是否為同一張發票，因為折讓單不同但發票號碼可能一樣 By Kin 2011/06/01
                'If sInvId <> aTblINV014.Rows(i)("INVID") Then
                iCount += 1
                _TotalCount += 1
                If UploadType = 0 Then
                    If (iCount >= fCutCount) AndAlso (iCount Mod fCutCount = 0) Then
                        bduString.AppendLine(strEnd)
                        strEnd = WriteB0301End(aTblINV014.Rows(i))
                        bduString.AppendLine(WriteB0301Header(aTblINV014.Rows(i), _
                                                            rowInv001))
                        bduString.AppendLine(WriteB0301Detail(aTblINV014.Rows(i)))
                        bduString.AppendLine(strEnd)
                        '#6296 第二個檔案會存到舊的路徑 By Kin 2012/08/16
                        'writeFile(InvFileType.B0301, bduString.ToString, aSeq)
                        writeFile(aFilePath, InvTypeEnum.InvFileType.B0301, bduString.ToString, aSeq)
                        aSeq += 1
                        strEnd = String.Empty
                        bduString = New StringBuilder
                    Else
                        If (Not String.IsNullOrEmpty(sInvId)) AndAlso _
                            (Not String.IsNullOrEmpty(sInvId)) AndAlso _
                            (Not String.IsNullOrEmpty(strEnd)) Then
                            bduString.AppendLine(strEnd)
                        End If
                        strEnd = WriteB0301End(aTblINV014.Rows(i))
                        bduString.AppendLine(WriteB0301Header(aTblINV014.Rows(i), _
                                                            rowInv001))
                        bduString.AppendLine(WriteB0301Detail(aTblINV014.Rows(i)))
                    End If
                Else
                    '#6371 將資料轉成XML檔 By Kin 2013/02/19                       
                    If dtXml.Rows.Count > 0 Then
                        WriteToXml(aFilePath, dtXml, rowInv001, InvTypeEnum.InvFileType.D0401)
                    End If
                    dtXml.Rows.Clear()
                    Dim rwNew As DataRow = dtXml.NewRow
                    rwNew.ItemArray = aTblINV014.Rows(i).ItemArray
                    dtXml.Rows.Add(rwNew)
                End If

                'End If
                sInvId = aTblINV014.Rows(i)("INVID").ToString
            Next
            If UploadType = 0 Then
                bduString.AppendLine(WriteB0301End(aTblINV014.Rows(aTblINV014.Rows.Count - 1)))
                writeFile(aFilePath, InvTypeEnum.InvFileType.B0301, bduString.ToString, aSeq)
            Else
                '#6371 將資料轉成XML檔 By Kin 2013/02/19                       
                WriteToXml(aFilePath, dtXml, rowInv001, InvTypeEnum.InvFileType.D0401)
            End If

            _B0301RcdCount = iCount
            Return True
        Catch ex As Exception
            Throw New Exception("Function: WriteB0301 訊息: " & ex.ToString)
        End Try
    End Function
    Private Function WriteB0301Header(ByVal aRowINV014 As DataRow, _
                                 ByVal aRowINV001 As DataRow) As String
        Try
            '#5945 原本是抓PAPERNO欄位,改抓AllowanceNo By Kin 2011/03/16
            strTmp = String.Empty
            strTmp = "M" & Common.GetString(aRowINV014("ALLOWANCENO"), 16) & _
                Common.GetString(aRowINV014("PAPERDATE"), 10) & _
                Common.GetString(aRowINV001("BUSINESSID"), 10) & _
                Common.GetString(aRowINV001("CompName"), 60) & Space(228) & _
                New String("0", 10) & Common.GetString(Right(aRowINV014("INVID"), 4), 60) & _
                Space(228) & "2"
            Return strTmp
        Catch ex As Exception
            Throw New Exception("Function: WriteB0301Header 訊息: " & ex.ToString)
        End Try

    End Function
    Private Function WriteB0301Detail(ByVal aRowINV014 As DataRow) As String
        Try
            Dim aQty As Integer = 0
            Dim aUnitPric As Integer = 0
            If Not IsDBNull(aRowINV014("QUANTITY")) Then
                aQty = System.Convert.ToDouble(aRowINV014("QUANTITY"))
            End If
            If (Not IsDBNull(aRowINV014("SALEAMOUNT"))) AndAlso (aQty > 0) Then
                aUnitPric = Math.Round(Convert.ToDouble(aRowINV014("SALEAMOUNT")) / aQty, 4)
            End If
            strTmp = String.Empty
            strTmp = "D" & Common.GetString(aRowINV014("INVDATE"), 10) & _
                Common.GetString(aRowINV014("INVID"), 10) & Space(8) & _
                Common.GetString(aRowINV014("DESCRIPTION"), 256) & _
                Right(New String("0"c, 12) & Format(aQty, "#.0000"), 17) & Space(6) & _
                Right(New String("0"c, 12) & Format(aUnitPric, "#.0000"), 17) & _
                Right(New String("0"c, 12) & Format(Convert.ToDouble(aRowINV014("SALEAMOUNT")), "#.0000"), 17) & _
                Common.GetString(Math.Round(Convert.ToDouble(aRowINV014("TAXAMOUNT"))), 12, AlignType.Right, True) & "1 " & _
                aRowINV014("TaxType")
            Return strTmp
        Catch ex As Exception
            Throw New Exception("Function: WriteB0301Detail 訊息: " & ex.ToString)
        End Try
    End Function
    Private Function WriteB0301End(ByVal aRowInv014 As DataRow) As String
        Try
            strTmp = String.Empty
            strTmp = "A" & Common.GetString( _
                Math.Round(Convert.ToDouble(aRowInv014("TaxAmount"))).ToString, 12, AlignType.Right, True) & _
                Common.GetString( _
                Math.Round(Convert.ToDouble(aRowInv014("SALEAMOUNT"))).ToString, 12, AlignType.Right, True)
            Return strTmp
        Catch ex As Exception
            Throw New Exception("Function: WritB0301End 訊息: " & ex.ToString)
        End Try
    End Function
    Private Function WriteB0401(ByVal UploadType As Integer) As Boolean
        Dim bduString As StringBuilder = Nothing
        'Dim sInvId As String = String.Empty
        Dim aSeq As Integer = 1
        Dim dtXML As DataTable = Nothing
        'Dim strEnd As String = String.Empty
        Dim iCount As Integer = 0
        Dim aTblINV014 As DataTable = FInvDs.Tables("INV014")
        Dim rowInv001 As DataRow = Nothing
        Dim aFilePath As String = Nothing
        If aTblINV014.Rows.Count <= 0 Then
            Return True
        Else
            If UploadType = 0 Then
                bduString = New StringBuilder
            Else
                dtXML = New DataTable
                For Each col As DataColumn In FInvDs.Tables("INV014").Columns
                    dtXML.Columns.Add(col.ColumnName, col.DataType)
                Next
            End If
        End If
        If Not String.IsNullOrEmpty(ExportElectronPath) Then
            aFilePath = ExportElectronPath
            If Not aFilePath.EndsWith("\") Then
                aFilePath = aFilePath & "\"
            End If
        Else
            aFilePath = CableSoft.B07.SystemIO.Invoice.GetINVPathName(InvArgument.GetConnect,
                                                                               InvArgument.GetCompCode, fB0401FilePath)
        End If

        If String.IsNullOrEmpty(aFilePath) Then
            Throw New Exception("B0401 檔案路徑設定有誤！")
        End If
        Try
            For i As Int32 = 0 To aTblINV014.Rows.Count - 1
                rowInv001 = FInvDs.Tables("INV001").Select("COMPID='" & aTblINV014.Rows(i)("COMPID") & "'").GetValue(0)
                iCount += 1
                _TotalCount += 1

                If UploadType = 0 Then
                    bduString.AppendLine(WriteB0401Header(aTblINV014.Rows(i), _
                                                      rowInv001))
                    If (iCount >= fCutCount) AndAlso (iCount Mod fCutCount = 0) Then

                        '#6296 第二個檔案會存到舊的路徑 By Kin 2012/08/16
                        'writeFile(InvFileType.B0401, bduString.ToString, aSeq)
                        writeFile(aFilePath, InvTypeEnum.InvFileType.B0401, bduString.ToString, aSeq)
                        aSeq += 1
                        'strEnd = String.Empty
                        bduString = New StringBuilder
                    Else

                    End If
                Else
                    '#6371 將資料轉成XML檔 By Kin 2013/02/19                       
                    If dtXML.Rows.Count > 0 Then
                        WriteToXml(aFilePath, dtXML, rowInv001, InvTypeEnum.InvFileType.D0501)
                    End If
                    dtXML.Rows.Clear()
                    Dim rwNew As DataRow = dtXML.NewRow
                    rwNew.ItemArray = aTblINV014.Rows(i).ItemArray
                    dtXML.Rows.Add(rwNew)
                End If

                'sInvId = aTblINV014.Rows(i)("INVID").ToString
            Next
            'bduString.AppendLine(WriteB0301End(aTblINV014.Rows(aTblINV014.Rows.Count - 1)))

            If UploadType = 0 Then
                writeFile(aFilePath, InvTypeEnum.InvFileType.A0501, bduString.ToString, aSeq)
            Else
                WriteToXml(aFilePath, dtXML, rowInv001, InvTypeEnum.InvFileType.D0501)
            End If
            _B0401RcdCount = iCount
            Return True

        Catch ex As Exception
            Throw New Exception("Function: WriteB0401 訊息: " & ex.ToString)
        Finally
            If dtXML IsNot Nothing Then
                dtXML.Dispose()
            End If
        End Try

    End Function
    Private Function WriteB0401Header(ByVal aRowInv014 As DataRow, _
                                      ByVal aRowInv001 As DataRow) As String
        Try
            Dim aDate As DateTime = DateTime.Parse(aRowInv014("UPTTIME").ToString)
            strTmp = String.Empty
            strTmp = "C" & Common.GetString(aRowInv014("ALLOWANCENO"), 16) & _
                New String("0", 10) & Common.GetString(aRowInv001("BUSINESSID"), 10) & _
                Common.GetString(aRowInv014("PAPERDATE"), 10) & _
                Common.GetString(Format(aDate, "yyyy/MM/dd"), 10) & _
                Common.GetString(Format(aDate, "hh:mm:ss"), 8) & _
                Common.GetString(aRowInv014("ObsoleteReason"), 200, AlignType.Left)
            Return strTmp
        Catch ex As Exception
            Throw New Exception("Function: WriteB0401Header 訊息: " & ex.ToString)
        End Try
    End Function
    Private Function CreateXMLFile(ByVal XmlFile As Xml.XmlTextWriter) As Boolean
        XmlFile.Formatting = Xml.Formatting.Indented
        XmlFile.Indentation = 4
        XmlFile.WriteStartDocument()
        Return True
    End Function
    Private Overloads Sub writeFile(ByVal WritePath As String,
                            ByVal aFileType As InvTypeEnum.InvFileType, ByVal aContText As String, _
                  ByVal aSeq As Integer)
        Dim aFileName As String = Format(FNow, "yyyyMMdd") & "-" & Format(Now, "HHmmss")
        aFileName = [Enum].GetName(GetType(InvTypeEnum.InvFileType), aFileType) & _
                "-" & aFileName & "-" & Format(aSeq, "0000") & ".TXT"
        Dim aBackFileName As String = aFileName
        If Not String.IsNullOrEmpty(_BackPath) Then
            aBackFileName = _BackPath & aFileName
        Else
            aBackFileName = String.Empty
        End If
        If WritePath.EndsWith("\") Then
            aFileName = WritePath & aFileName
        Else
            aFileName = WritePath & "\" & aFileName
        End If

        Try
            Using objWrite As New StreamWriter(aFileName, True, Encoding.GetEncoding(950))
                objWrite.Write(aContText)
                objWrite.Flush()
                objWrite.Close()
            End Using
            '#5922 備份一分至媒體申報檔輸出路徑 By Kin 2011/02/14
            If (Not String.IsNullOrEmpty(aBackFileName)) AndAlso (aBackFileName <> aFileName) Then
                Using objBackWrite As New StreamWriter(aBackFileName, False, Encoding.GetEncoding(950))
                    objBackWrite.Write(aContText)
                    objBackWrite.Flush()
                    objBackWrite.Close()
                End Using
            End If

        Catch ex As Exception
            Throw New Exception(ex.ToString)
        End Try

    End Sub
    Private Overloads Sub writeFile(ByVal aFileType As InvTypeEnum.InvFileType, ByVal aContText As String, _
                          ByVal aSeq As Integer)
        writeFile(FFilePath, aFileType, aContText, aSeq)
    End Sub

#Region "IDisposable Support"
    Private disposedValue As Boolean ' 偵測多餘的呼叫

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: 處置 Managed 狀態 (Managed 物件)。
                If InvArgument IsNot Nothing Then
                    InvArgument.Dispose()
                    InvArgument = Nothing
                End If
                If Me.FInvDs IsNot Nothing Then
                    Me.FInvDs.Dispose()
                    Me.FInvDs = Nothing
                End If
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
