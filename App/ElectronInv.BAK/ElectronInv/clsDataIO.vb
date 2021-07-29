Imports System
Imports System.IO
Imports System.Data
Imports System.Text

Public Class clsDataIO
    Private FFilePath As String
    Private FDs As DataSet
    Private strTmp As String = String.Empty
    Private FNow As Date = Nothing
    Private _TotalCount As Int32 = 0
    Private _A0401RcdCount As Int32 = 0
    Private _A0501RcdCount As Int32 = 0
    Private _B0301RcdCount As Int32 = 0
    Private _B0401RcdCount As Int32 = 0
    Private _E0402RcdCount As Int32 = 0
    Private _BackPath As String = String.Empty
    Private Const fCutCount As Int32 = 4000
    Private Const fA0401FilePath As String = "A0401FilePath"
    Private Const fA0501FilePath As String = "A0501FilePath"
    Private Const fB0301FilePath As String = "B0301FilePath"
    Private Const fB0401FilePath As String = "B0401FilePath"
    Private Const fE0402FilePath As String = "E0402FilePath"

    Public Sub New(ByVal aFilePath As String, ByVal aDataSet As DataSet)
        FFilePath = aFilePath
        FDs = aDataSet
        FNow = Now
    End Sub
    Public Sub New()
        FNow = Now
    End Sub
    Public Property A0401FilePath As String
    Public Property A0501FilePath As String
    Public Property B0301FilePath As String
    Public Property B0401FilePath As String
    Public Property E0402FilePath As String
    'Public Property tbInv003 As DataTable
    Public ReadOnly Property WriteFileTime() As Date
        Get
            Return FNow
        End Get
    End Property
    Public ReadOnly Property TotalCount() As Int32
        Get
            Return _TotalCount
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
    Public Property BackPath() As String
        Get
            Return _BackPath
        End Get
        Set(ByVal value As String)
            _BackPath = value
        End Set
    End Property
    Public Function WriteINVFile(ByVal aInvType As INVTYPE, _
                                 ByVal aUseDiscount As Boolean, _
                                 ByVal aUseDisCountOv As Boolean, _
                                 ByVal aUseSpaceInv As Boolean, _
                                 ByVal aBackPath As String,
                                 ByVal UploadType As Integer) As Boolean
        Try
            _TotalCount = 0
            FNow = Now
            _BackPath = aBackPath
            Select Case aInvType
                Case INVTYPE.All
                    WriteA0401(UploadType)
                    WriteA0501(UploadType)
                Case INVTYPE.OV
                    WriteA0501(UploadType)
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

    Public Function WriteA0501(ByVal UploadType As Int32) As Boolean
        Dim bduString As StringBuilder = Nothing
        Dim aRowINV007 As DataRow() = _
            FDs.Tables("INV007").Select("IsObsolete='Y'")
        Dim dtXML As DataTable = Nothing
        Dim aSeq As Integer = 1
        Dim iCount As Integer = 0
        Dim sINvId As String = String.Empty
        Dim rowInv001 As DataRow = Nothing
        _A0501RcdCount = 0
        If aRowINV007.Length <= 0 Then
            Return True
        End If
        Dim aFilePath As String = clsData.GetINVPathName(clsCommon.GetConnect, clsCommon.GetCompCode, fA0501FilePath)
        If String.IsNullOrEmpty(aFilePath) Then
            Throw New Exception("A0501 檔案路徑設定有誤！")
        End If
        If UploadType = 0 Then
            bduString = New StringBuilder
        Else
            dtXML = New DataTable
            For Each col As DataColumn In FDs.Tables("INV007").Columns
                dtXML.Columns.Add(col.ColumnName, col.DataType)
            Next
        End If
        strTmp = String.Empty
        Try
            For i As Integer = 0 To aRowINV007.Length - 1
                rowInv001 = FDs.Tables("INV001").Select("COMPID='" & aRowINV007(i)("COMPID") & "'").GetValue(0)
                If sINvId <> aRowINV007(i).Item("INVID") Then
                    iCount += 1
                    _TotalCount += 1
                    If UploadType = 0 Then
                        strTmp = String.Empty
                        strTmp = "C" & aRowINV007(i).Item("INVID") & aRowINV007(i).Item("INVDATE") & _
                            New String("0"c, 10) & clsCommon.GetString(rowInv001("BusinessId"), 10) & _
                            Format(aRowINV007(i)("UptTime"), "d") & Format(aRowINV007(i)("UptTime"), "HH:mm:ss") & _
                            Space(60) & clsCommon.GetString(aRowINV007(i)("ObsoleteReason"), 200)
                        bduString.AppendLine(strTmp)
                        If (iCount >= fCutCount) AndAlso (iCount Mod fCutCount = 0) Then
                            '#6296 第二個檔案會存到舊的路徑 By Kin 2012/08/16
                            'writeFile(InvFileType.A0501, bduString.ToString, aSeq)
                            writeFile(aFilePath, InvFileType.A0501, bduString.ToString, aSeq)
                            aSeq += 1
                            bduString = New StringBuilder
                        End If
                    Else
                        '#6371 將資料轉成XML檔 By Kin 2013/02/19                       
                        If dtXML.Rows.Count > 0 Then
                            WriteToXml(aFilePath, dtXML, rowInv001, InvFileType.C0501)
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
                writeFile(aFilePath, InvFileType.A0501, bduString.ToString, aSeq)
            Else
                WriteToXml(aFilePath, dtXML, rowInv001, InvFileType.C0501)
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
    Private Function GetXMLFile(ByVal aFilePath As String, ByVal aInvId As String, ByVal aInvFileType As InvFileType) As String
        Dim aReturn As String = Nothing
     
            aReturn = [Enum].GetName(GetType(InvFileType), aInvFileType) & "-" & aInvId & "-" & _
                FNow.ToString("yyyyMMdd") & "-" & FNow.ToString("HHmmss") & ".xml"


        If aFilePath.EndsWith("\") Then
            aReturn = aFilePath & aReturn
        Else
            aReturn = aFilePath & "\" & aReturn
        End If
        Return aReturn
    End Function

    Private Function WriteD0401(ByVal dtInv014 As DataTable, ByVal rwInv001 As DataRow, ByVal XmlFile As Xml.XmlWriter) As Boolean
        Try
            CreateXMLFile(XmlFile)
            XmlFile.WriteStartElement("Allowance")
            XmlFile.WriteAttributeString("xsi:schemaLocation", "urn:GEINV:eInvoiceMessage:D0401:3.0 D0401.xsd")
            XmlFile.WriteAttributeString("xmlns", "urn:GEINV:eInvoiceMessage:D0401:3.0")
            XmlFile.WriteAttributeString("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance")
            '-------------------------Main--------------------------------------------------------------
            XmlFile.WriteStartElement("Main")
            '-------------------------AllowanceNumber---------------------------------------------
            XmlFile.WriteStartElement("AllowanceNumber")
            XmlFile.WriteValue(dtInv014.Rows(0).Item("PaperNo").ToString)
            XmlFile.WriteEndElement()
            '--------------------------------------------------------------------------------------------
            '-------------------------AllowanceDate-------------------------------------------------
            XmlFile.WriteStartElement("AllowanceDate")
            XmlFile.WriteValue(Date.Parse(dtInv014.Rows(0).Item("PaperDate").ToString).ToString("yyyy-MM-dd"))
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
                    XmlFile.WriteValue(rwInv001.Item("Manager").ToString)
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
                XmlFile.WriteValue(Date.Parse(rw.Item("InvDate").ToString).ToString("yyyy-MM-dd"))
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

        Catch ex As Exception
            Throw New Exception("Function: WriteD0401 訊息: " & ex.ToString)
        End Try
    End Function
    Private Overloads Function WriteE0402() As Boolean

        Dim dtXml As DataTable = Nothing
        Dim aTblINV099 As DataTable = FDs.Tables("INV099")
        Dim rowInv001 As DataRow = Nothing
        If aTblINV099.Rows.Count <= 0 Then
            Return True
        Else
            dtXml = New DataTable
            For Each col As DataColumn In FDs.Tables("INV099").Columns
                dtXml.Columns.Add(col.ColumnName, col.DataType)
            Next
        End If

        Dim aFilePath As String = clsData.GetINVPathName(clsCommon.GetConnect, clsCommon.GetCompCode, fE0402FilePath)
        If String.IsNullOrEmpty(aFilePath) Then
            Throw New Exception("E0402 檔案路徑設定有誤！")
        End If
        Try
            For i As Int32 = 0 To aTblINV099.Rows.Count - 1
                rowInv001 = FDs.Tables("INV001").Select("COMPID='" & aTblINV099.Rows(i)("COMPID") & "'").GetValue(0)
                _TotalCount += 1
                dtXml.Rows.Clear()
                Dim rwNew As DataRow = dtXml.NewRow
                rwNew.ItemArray = aTblINV099.Rows(i).ItemArray
                dtXml.Rows.Add(rwNew)
                WriteToXml(aFilePath, dtXml, rowInv001, InvFileType.E0402)
            Next



            _E0402RcdCount = aTblINV099.Rows.Count
            Return True
        Catch ex As Exception
            Throw New Exception("Function: WriteB0301 訊息: " & ex.ToString)
        End Try
    End Function

    Private Overloads Function WriteE0402(ByVal dt099 As DataTable, ByVal rwInv001 As DataRow, ByVal XmlFile As Xml.XmlWriter) As Boolean
        Try
            CreateXMLFile(XmlFile)
            XmlFile.WriteStartElement("BranchTrack")
            XmlFile.WriteAttributeString("xsi:schemaLocation", "urn:GEINV:eInvoiceMessage:E0402:3.0 E0402.xsd")
            XmlFile.WriteAttributeString("xmlns", "urn:GEINV:eInvoiceMessage:E0402:3.0")
            XmlFile.WriteAttributeString("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance")
            '-------------------------Main--------------------------------------------------------------
            XmlFile.WriteStartElement("Main")
            '-------------------------HeadBan----------------------------------------------------------
            XmlFile.WriteStartElement("HeadBan")
            XmlFile.WriteValue(rwInv001.Item("BusinessId").ToString)
            XmlFile.WriteEndElement()       'HeadBan
            '---------------------------------------------------------------------------------------------
            '-------------------------BranchBan----------------------------------------------------------
            XmlFile.WriteStartElement("BranchBan")
            XmlFile.WriteValue(rwInv001.Item("BusinessId").ToString)
            XmlFile.WriteEndElement()       'BranchBan
            '---------------------------------------------------------------------------------------------
            '-------------------------InvoiceType----------------------------------------------------------
            XmlFile.WriteStartElement("InvoiceType")
            XmlFile.WriteValue("05")
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
            XmlFile.WriteStartElement("BranchTrackBlankItem")
            '----------------------------InvoiceBeginNo----------------------------------------------
            XmlFile.WriteStartElement("InvoiceBeginNo")
            XmlFile.WriteValue(dt099.Rows(0).Item("CurNum").ToString)
            XmlFile.WriteEndElement() 'InvoiceBeginNo
            '----------------------------InvoiceEndNo------------------------------------------------
            XmlFile.WriteStartElement("InvoiceEndNo")
            XmlFile.WriteValue(dt099.Rows(0).Item("EndNum").ToString)
            XmlFile.WriteEndElement() 'InvoiceEndNo
            '--------------------------------------------------------------------------------------------
            XmlFile.WriteEndElement() 'BranchTrackBlankItem
            XmlFile.WriteEndElement() 'Details
        Catch ex As Exception
            Throw New Exception("Function: WriteE0402 訊息: " & ex.ToString)
        End Try
    End Function
    Private Function WriteD0501(ByVal dtInv014 As DataTable, ByVal rwInv001 As DataRow, ByVal XmlFile As Xml.XmlWriter) As Boolean
        Try
            CreateXMLFile(XmlFile)
            XmlFile.WriteStartElement("CancelAllowance")
            XmlFile.WriteAttributeString("xsi:schemaLocation", "urn:GEINV:eInvoiceMessage:D0501:3.0 D0501.xsd")
            XmlFile.WriteAttributeString("xmlns", "urn:GEINV:eInvoiceMessage:D0501:3.0")
            XmlFile.WriteAttributeString("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance")
            '------------------------------------------CancelAllowanceNumber-----------------------------------------
            XmlFile.WriteStartElement("CancelAllowanceNumber")
            XmlFile.WriteValue(dtInv014.Rows(0).Item("AllowanceNo").ToString)
            XmlFile.WriteEndElement()
            '-------------------------------------------------------------------------------------------------------------------
            '------------------------------------------AllowanceDate-------------------------------------------------------
            XmlFile.WriteStartElement("AllowanceDate")
            XmlFile.WriteValue(Date.Parse(dtInv014.Rows(0).Item("PaperDate").ToString).ToString("yyyy-MM-dd"))
            XmlFile.WriteEndElement() 'AllowanceDate
            '-------------------------------------------------------------------------------------------------------------------
            '-----------------------------------------BuyerId----------------------------------------------------------------
            XmlFile.WriteStartElement("BuyerId")
            XmlFile.WriteValue("0000000000")
            XmlFile.WriteEndElement() 'BuyerId
            '-------------------------------------------------------------------------------------------------------------------
            '-----------------------------------------SellerId-----------------------------------------------------------------
            XmlFile.WriteStartElement("SellerId")
            XmlFile.WriteValue(rwInv001.Item("BusinessId").ToString)
            XmlFile.WriteEndElement() 'SellerId
            '-------------------------------------------------------------------------------------------------------------------
            '-----------------------------------------CancelDate-----------------------------------------------------------
            XmlFile.WriteStartElement("CancelDate")
            XmlFile.WriteValue(Date.Parse(dtInv014.Rows(0).Item("UptTime").ToString).ToString("yyyy-MM-dd"))
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
        Catch ex As Exception
            Throw New Exception("Function: WriteD0501 訊息: " & ex.ToString)
        End Try
    End Function
    Private Function WriteC0501(ByVal dtInv007 As DataTable, ByVal rwInv001 As DataRow, ByVal XmlFile As Xml.XmlWriter) As Boolean
        Try
            CreateXMLFile(XmlFile)
            XmlFile.WriteStartElement("CancelInvoice")
            XmlFile.WriteAttributeString("xsi:schemaLocation", "urn:GEINV:eInvoiceMessage:C0501:3.0 C0501.xsd")
            XmlFile.WriteAttributeString("xmlns", "urn:GEINV:eInvoiceMessage:C0501:3.0")
            XmlFile.WriteAttributeString("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance")
            '--------------------------------------------------------------------------
            '---------------------------CancelInvoiceNumber-------------------
            XmlFile.WriteStartElement("CancelInvoiceNumber")
            XmlFile.WriteValue(dtInv007.Rows(0).Item("InvId").ToString)
            XmlFile.WriteEndElement() 'CancelInvoiceNumber
            '---------------------------------------------------------------------------
            '---------------------------InvoiceDate--------------------------------
            XmlFile.WriteStartElement("InvoiceDate")
            XmlFile.WriteValue(Date.Parse(dtInv007.Rows(0).Item("InvDate").ToString).ToString("yyyy-MM-dd"))
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
            XmlFile.WriteValue(Date.Parse(dtInv007.Rows(0).Item("UptTime").ToString).ToString("yyyy-MM-dd"))
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

        Catch ex As Exception
            Throw New Exception("Function: WriteC0401 訊息: " & ex.ToString)
        End Try
        Return True
    End Function


    Private Function WriteC0401(ByVal dtInv007 As DataTable, ByVal rwInv001 As DataRow, ByVal XmlFile As Xml.XmlWriter) As Boolean
        Try
            CreateXMLFile(XmlFile)
            XmlFile.WriteStartElement("Invoice")
            XmlFile.WriteAttributeString("xsi:schemaLocation", "urn:GEINV:eInvoiceMessage:C0401:3.0 C0401.xsd")
            XmlFile.WriteAttributeString("xmlns", "urn:GEINV:eInvoiceMessage:C0401:3.0")
            XmlFile.WriteAttributeString("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance")
            XmlFile.WriteStartElement("Main")
            '-------------------------------InvoiceNumbe----------------------------------------
            XmlFile.WriteStartElement("InvoiceNumber")
            XmlFile.WriteValue(dtInv007.Rows(0).Item("InvId").ToString)

            XmlFile.WriteEndElement() 'InvoiceNumber
            '------------------------------InvoiceDate---------------------------------------------------------------------
            XmlFile.WriteStartElement("InvoiceDate")
            XmlFile.WriteValue(DateTime.Parse(dtInv007.Rows(0).Item("InvDate").ToString).ToString("yyyy-MM-dd"))
            XmlFile.WriteEndElement() 'InvoiceDate
            '-----------------------------------------------------------------------------
            '------------------------------InvoiceTime---------------------------------------------------------------------
            '#6600 測試不OK,少了InvoiceTime By Kin 2013/11/12
            XmlFile.WriteStartElement("InvoiceTime")
            XmlFile.WriteValue(DateTime.Parse(dtInv007.Rows(0).Item("CreateInvDate").ToString).ToString("HH:mm:ss"))
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
                    XmlFile.WriteValue(rwInv001.Item("Manager").ToString)
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
                XmlFile.WriteEndElement() 'Seller & Buyer
            Next
            '----------------------------------------------------------------------------

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
                    XmlFile.WriteValue("05")
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
                (DBNull.Value.Equals(dtInv007.Rows(0).Item("PrintTime"))) AndAlso
                (Int32.Parse(dtInv007.Rows(0).Item("RefNo")) <> 1) Then
                '#6600 若INV007. PrintTime有值，此欄NULL By Kin 2013/10/02
                XmlFile.WriteStartElement("CarrierType")
                XmlFile.WriteValue(dtInv007.Rows(0).Item("CarrierType").ToString)
                XmlFile.WriteEndElement() 'CarrierType
            End If
            '-------------------------------------------------------------------------------------
            '----------------------------------CarrierId1--------------------------------------
            '#6600 如果是PrintTime is Null才可寫入 By Kin 2013/10/07
            '#6600 測試不OK,捐贈也不需要載具資料 By Kin 2013/11/11
            If (Not DBNull.Value.Equals(dtInv007.Rows(0).Item("CarrierId1"))) AndAlso
                 (DBNull.Value.Equals(dtInv007.Rows(0).Item("PrintTime"))) AndAlso
                 (Int32.Parse(dtInv007.Rows(0).Item("RefNo")) <> 1) Then
                XmlFile.WriteStartElement("CarrierId1")
                XmlFile.WriteValue(dtInv007.Rows(0).Item("CarrierId1").ToString)
                XmlFile.WriteEndElement() 'CarrierId1
            End If
            '-------------------------------------------------------------------------------------
            '----------------------------------CarrierId2--------------------------------------
            '#6600 測試不OK,捐贈也不需要載具資料 By Kin 2013/11/11
            If (DBNull.Value.Equals(dtInv007.Rows(0).Item("PrintTime"))) AndAlso
                (Int32.Parse(dtInv007.Rows(0).Item("RefNo")) <> 1) Then
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
                'XmlFile.WriteValue(Math.Round((Double.Parse(rw.Item("TotalAmount")) / Double.Parse(rw.Item("Quantity"))), 2))
                XmlFile.WriteValue(rw.Item("UnitPrice").ToString)
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
            XmlFile.WriteValue(Math.Round(Double.Parse(dtInv007.Rows(0).Item("InvAmount").ToString), 0))
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
            XmlFile.WriteValue(Math.Round(Double.Parse(dtInv007.Rows(0).Item("TaxAmount")), 0))
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
        Catch ex As Exception
            Throw New Exception("Function: WriteC0401 訊息: " & ex.ToString)
        End Try
        Return True
    End Function
    Public Function WriteA0401(ByVal aUploadType As Integer) As Boolean
        Dim bduString As StringBuilder = Nothing
        'Dim objXml As Xml.XmlTextWriter = Nothing
        Dim dtXml As DataTable = Nothing
        Dim sInvId As String = String.Empty
        Dim aRowInv007 As DataRow() = _
            FDs.Tables("INV007").Select("IsObsolete='N'")

        Dim aSeq As Integer = 1
        Dim strEnd As String = String.Empty
        Dim iCount As Int32 = 0
        Dim rowInv001 As DataRow = Nothing
        _A0401RcdCount = 0
        If aRowInv007.Length <= 0 Then
            Return True
        End If
        Dim aFilePath As String = clsData.GetINVPathName(clsCommon.GetConnect, clsCommon.GetCompCode, fA0401FilePath)
        If String.IsNullOrEmpty(aFilePath) Then
            Throw New Exception("A0401 檔案路徑設定有誤！")
        End If
        If aUploadType = 0 Then
            bduString = New StringBuilder()
        Else
            dtXml = New DataTable
            For Each col As DataColumn In FDs.Tables("INV007").Columns
                dtXml.Columns.Add(col.ColumnName, col.DataType)
            Next
        End If


        Try
            For i As Integer = 0 To aRowInv007.Length - 1

                '#5881 有可能是全部公司,所以INV001要Select By Kin 2010/12/14
                rowInv001 = FDs.Tables("INV001").Select("COMPID='" & aRowInv007(i)("COMPID") & "'").GetValue(0)
                If sInvId <> aRowInv007(i)("INVID") Then
                    iCount += 1
                    _TotalCount += 1
                    If aUploadType = 0 Then
                        If (iCount >= fCutCount) AndAlso (iCount Mod fCutCount = 0) Then
                            bduString.AppendLine(strEnd)
                            strEnd = WriteA0401End(aRowInv007(i))
                            bduString.AppendLine(WriteA0401Head(aRowInv007(i), _
                                                                rowInv001))
                            bduString.AppendLine(WriteA0401Detail(aRowInv007(i), _
                                                                  rowInv001))
                            bduString.AppendLine(strEnd)
                            '#6296 第二個檔案會存到舊的路徑 By Kin 2012/08/16
                            'writeFile(InvFileType.A0401, bduString.ToString, aSeq)
                            writeFile(aFilePath, InvFileType.A0401, bduString.ToString, aSeq)
                            aSeq += 1
                            strEnd = String.Empty
                            bduString = New StringBuilder
                        Else

                            If (Not String.IsNullOrEmpty(sInvId)) AndAlso _
                                (Not String.IsNullOrEmpty(sInvId)) AndAlso _
                                (Not String.IsNullOrEmpty(strEnd)) Then
                                bduString.AppendLine(strEnd)
                            End If

                            strEnd = WriteA0401End(aRowInv007(i))
                            bduString.AppendLine(WriteA0401Head(aRowInv007(i), _
                                                                rowInv001))
                            bduString.AppendLine(WriteA0401Detail(aRowInv007(i), _
                                                                  rowInv001))
                        End If
                    Else
                        '#6371 將資料轉成XML檔 By Kin 2013/02/19                       
                        If dtXml.Rows.Count > 0 Then
                            WriteToXml(aFilePath, dtXml, rowInv001, InvFileType.C0401)
                        End If
                        dtXml.Rows.Clear()
                        Dim rwNew As DataRow = dtXml.NewRow
                        rwNew.ItemArray = aRowInv007(i).ItemArray
                        dtXml.Rows.Add(rwNew)
                    End If


                Else
                    If aUploadType = 0 Then
                        bduString.AppendLine(WriteA0401Detail(aRowInv007(i), _
                                                         rowInv001))
                    Else
                        '#6371 將資料轉成XML檔 By Kin 2013/02/19
                        Dim rwNew As DataRow = dtXml.NewRow
                        rwNew.ItemArray = aRowInv007(i).ItemArray
                        dtXml.Rows.Add(rwNew)
                    End If
                End If
                sInvId = aRowInv007(i)("INVID").ToString
            Next
            If aUploadType = 0 Then
                bduString.AppendLine(WriteA0401End(aRowInv007(aRowInv007.Length - 1)))
                writeFile(aFilePath, InvFileType.A0401, bduString.ToString, aSeq)
            Else
                WriteToXml(aFilePath, dtXml, rowInv001, InvFileType.C0401)
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
    Private Sub WriteToXml(ByVal FilePath As String, ByVal dtXml As DataTable, ByVal rwInv001 As DataRow, FileType As InvFileType)
        Dim XmlFileName As String = Nothing
        If FileType <> InvFileType.E0402 Then
            XmlFileName = GetXMLFile(FilePath, dtXml.Rows(0).Item("InvId"), FileType)
        Else
            XmlFileName = GetXMLFile(FilePath, String.Format("{0}-{1}",
                                                             dtXml.Rows(0).Item("YearMonth"),
                                                             dtXml.Rows(0).Item("CurNum")), FileType)
        End If
        Dim XmlFile As Xml.XmlTextWriter
        XmlFile = New Xml.XmlTextWriter(XmlFileName, System.Text.Encoding.UTF8)
        Select Case FileType
            Case InvFileType.C0401
                WriteC0401(dtXml, rwInv001, XmlFile)
            Case InvFileType.C0501
                WriteC0501(dtXml, rwInv001, XmlFile)
            Case InvFileType.D0401
                WriteD0401(dtXml, rwInv001, XmlFile)
            Case InvFileType.D0501
                WriteD0501(dtXml, rwInv001, XmlFile)
            Case InvFileType.E0402
                WriteE0402(dtXml, rwInv001, XmlFile)
        End Select

        XmlFile.WriteEndDocument()
        XmlFile.Flush()
        XmlFile.Close()

        Dim aBackFileName As String = Path.GetFileName(XmlFileName)
        If Not String.IsNullOrEmpty(_BackPath) Then

            aBackFileName = _BackPath & aBackFileName

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
    End Sub
    Private Function WriteXmlEndElement(ByVal XmlFile As Xml.XmlTextWriter, ByRef StartElementCount As Integer) As Boolean
        XmlFile.WriteEndElement()
        StartElementCount -= 1
    End Function
    Private Function WriteXmlElement(ByVal XmlFile As Xml.XmlTextWriter, ByVal ElementName As String, ByRef StartElementCount As Integer) As Boolean
        Try
            StartElementCount += 1
            XmlFile.WriteStartElement(ElementName)

        Catch ex As Exception
            Throw New Exception("Function:WriteXmlElement 訊息:" & ex.ToString)

        End Try
        Return True
    End Function
    Private Function CreateXMLFile(ByVal XmlFile As Xml.XmlTextWriter) As Boolean
        XmlFile.Formatting = Xml.Formatting.Indented
        XmlFile.Indentation = 4
        XmlFile.WriteStartDocument()
        Return True
    End Function


    '#6001 增加B0401檔案 By Kin 2011/07/14
    Private Function WriteB0401(ByVal UploadType As Int32) As Boolean
        Dim bduString As StringBuilder = Nothing
        'Dim sInvId As String = String.Empty
        Dim aSeq As Integer = 1
        Dim dtXML As DataTable = Nothing
        'Dim strEnd As String = String.Empty
        Dim iCount As Int32 = 0
        Dim aTblINV014 As DataTable = FDs.Tables("INV014")
        Dim rowInv001 As DataRow = Nothing
        If aTblINV014.Rows.Count <= 0 Then
            Return True
        Else
            If UploadType = 0 Then
                bduString = New StringBuilder
            Else
                dtXML = New DataTable
                For Each col As DataColumn In FDs.Tables("INV014").Columns
                    dtXML.Columns.Add(col.ColumnName, col.DataType)
                Next
            End If
        End If
        Dim aFilePath As String = clsData.GetINVPathName(clsCommon.GetConnect, clsCommon.GetCompCode, fB0401FilePath)
        If String.IsNullOrEmpty(aFilePath) Then
            Throw New Exception("B0401 檔案路徑設定有誤！")
        End If
        Try
            For i As Int32 = 0 To aTblINV014.Rows.Count - 1
                rowInv001 = FDs.Tables("INV001").Select("COMPID='" & aTblINV014.Rows(i)("COMPID") & "'").GetValue(0)
                iCount += 1
                _TotalCount += 1

                If UploadType = 0 Then
                    bduString.AppendLine(WriteB0401Header(aTblINV014.Rows(i), _
                                                      rowInv001))
                    If (iCount >= fCutCount) AndAlso (iCount Mod fCutCount = 0) Then

                        '#6296 第二個檔案會存到舊的路徑 By Kin 2012/08/16
                        'writeFile(InvFileType.B0401, bduString.ToString, aSeq)
                        writeFile(aFilePath, InvFileType.B0401, bduString.ToString, aSeq)
                        aSeq += 1
                        'strEnd = String.Empty
                        bduString = New StringBuilder
                    Else

                    End If
                Else
                    '#6371 將資料轉成XML檔 By Kin 2013/02/19                       
                    If dtXML.Rows.Count > 0 Then
                        WriteToXml(aFilePath, dtXML, rowInv001, InvFileType.D0501)
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
                writeFile(aFilePath, InvFileType.A0501, bduString.ToString, aSeq)
            Else
                WriteToXml(aFilePath, dtXML, rowInv001, InvFileType.D0501)
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

    Private Function WriteB0301(ByVal UploadType As Int32) As Boolean
        Dim bduString As StringBuilder = Nothing
        Dim sInvId As String = String.Empty
        Dim aSeq As Integer = 1
        Dim strEnd As String = String.Empty
        Dim iCount As Int32 = 0
        Dim dtXml As DataTable = Nothing
        Dim aTblINV014 As DataTable = FDs.Tables("INV014")
        Dim rowInv001 As DataRow = Nothing
        If aTblINV014.Rows.Count <= 0 Then
            Return True
        Else
            If UploadType = 0 Then
                bduString = New StringBuilder()
            Else
                dtXml = New DataTable
                For Each col As DataColumn In FDs.Tables("INV014").Columns
                    dtXml.Columns.Add(col.ColumnName, col.DataType)
                Next
            End If
        End If
        Dim aFilePath As String = clsData.GetINVPathName(clsCommon.GetConnect, clsCommon.GetCompCode, fB0301FilePath)
        If String.IsNullOrEmpty(aFilePath) Then
            Throw New Exception("B0301 檔案路徑設定有誤！")
        End If
        Try
            For i As Int32 = 0 To aTblINV014.Rows.Count - 1
                '#5881 有可能是全部公司,所以INV001要Select By Kin 2010/12/14
                rowInv001 = FDs.Tables("INV001").Select("COMPID='" & aTblINV014.Rows(i)("COMPID") & "'").GetValue(0)
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
                        writeFile(aFilePath, InvFileType.B0301, bduString.ToString, aSeq)
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
                        WriteToXml(aFilePath, dtXml, rowInv001, InvFileType.D0401)
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
                writeFile(aFilePath, InvFileType.B0301, bduString.ToString, aSeq)
            Else
                '#6371 將資料轉成XML檔 By Kin 2013/02/19                       
                WriteToXml(aFilePath, dtXml, rowInv001, InvFileType.D0401)
            End If

            _B0301RcdCount = iCount
            Return True
        Catch ex As Exception
            Throw New Exception("Function: WriteB0301 訊息: " & ex.ToString)
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
            strTmp = "D" & clsCommon.GetString(aRowINV007("Description"), 256) & _
                    aQuantity & Space(6) & aPrice & _
                    Right("000000000000" & Format(aRowINV007("TotalAmount"), "#.0000"), 17) & _
                    clsCommon.GetString(aRowINV007("SEQ"), 2, AlignType.Right, True) & Space(66)


            '#5910 第6欄原本填Inv007.TotalAmount,現在改填Inv007.InvAmount By Kin 2011/01/12
            ''#5910 第7欄原本填Inv008.SEQ 現在固定填1 By Kin 2011/01/12
            'strTmp = "D" & clsCommon.GetString(aRowINV007("Description"), 256) & _
            '        aQuantity & Space(6) & aPrice & _
            '        aPrice & _
            '        clsCommon.GetString("1", 2, AlignType.right, True) & Space(66)

            Return strTmp

        Catch ex As Exception
            Throw New Exception("Function: WriteA0401Detail 訊息: " & ex.ToString)
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
            strTmp = "D" & clsCommon.GetString(aRowINV014("INVDATE"), 10) & _
                clsCommon.GetString(aRowINV014("INVID"), 10) & Space(8) & _
                clsCommon.GetString(aRowINV014("DESCRIPTION"), 256) & _
                Right(New String("0"c, 12) & Format(aQty, "#.0000"), 17) & Space(6) & _
                Right(New String("0"c, 12) & Format(aUnitPric, "#.0000"), 17) & _
                Right(New String("0"c, 12) & Format(Convert.ToDouble(aRowINV014("SALEAMOUNT")), "#.0000"), 17) & _
                clsCommon.GetString(Math.Round(Convert.ToDouble(aRowINV014("TAXAMOUNT"))), 12, AlignType.Right, True) & "1 " & _
                aRowINV014("TaxType")
            Return strTmp
        Catch ex As Exception
            Throw New Exception("Function: WriteB0301Detail 訊息: " & ex.ToString)
        End Try
    End Function

    Private Function WriteA0401End(ByVal aRowInv007 As DataRow) As String
        Try
            strTmp = String.Empty
            strTmp = "A" & clsCommon.GetString( _
                Math.Round(System.Convert.ToDouble _
                           (aRowInv007("InvAmount"))).ToString, 12, _
                           AlignType.Right, True) & aRowInv007("TaxType") & _
                           "0.0" & aRowInv007("TAXRATE") & "00" & _
                           "000000000000" & clsCommon.GetString(aRowInv007("INVAmount"), 12, AlignType.Right, True) & _
                            Space(45)

            '"000000000000" & "000000000000.0000" & "00000000.0000" & Space(3)

            Return strTmp
        Catch ex As Exception
            Throw New Exception("Function: WriteA0401End 訊息: " & ex.ToString)
        End Try

    End Function
    Private Function WriteB0301End(ByVal aRowInv014 As DataRow) As String
        Try
            strTmp = String.Empty
            strTmp = "A" & clsCommon.GetString( _
                Math.Round(Convert.ToDouble(aRowInv014("TaxAmount"))).ToString, 12, AlignType.Right, True) & _
                clsCommon.GetString( _
                Math.Round(Convert.ToDouble(aRowInv014("SALEAMOUNT"))).ToString, 12, AlignType.Right, True)
            Return strTmp
        Catch ex As Exception
            Throw New Exception("Function: WritB0301End 訊息: " & ex.ToString)
        End Try
    End Function

    Private Function WriteA0401Head(ByVal aRowINV007 As DataRow, _
                                    ByVal aRowINV001 As DataRow) As String
        Try
            '#6262 增加判斷INV003.ImportNote By Kin 2012/06/22
            Dim aImportNote As String = Left(clsCommon.GetString(
                FDs.Tables("INV003").Select(String.Format("CompId='{0}'", aRowINV007.Item("CompId").ToString))(0).Item("ImportNote"), 1) & " ", 1)
            If aRowINV007.Item("TaxType").ToString <> "2" Then
                aImportNote = " "
            End If
            strTmp = String.Empty
            strTmp = "M" & aRowINV007("INVID") & aRowINV007("INVDATE") & _
                Space(8) & clsCommon.GetString(aRowINV001("BusinessId"), 10, AlignType.Left) & _
                clsCommon.GetString(aRowINV001("CompName"), 60, AlignType.Left) & _
                Space(228) & New String("0"c, 10) & _
                clsCommon.GetString(Right(aRowINV007("INVID").ToString, 4), 60) & Space(439) & aImportNote & _
                clsCommon.GetString(aRowINV001("TaxDivName"), 40) & clsCommon.GetString(aRowINV001("TaxGetDate"), 10) & _
                clsCommon.GetString(aRowINV001("TaxNum1"), 40) & clsCommon.GetString(aRowINV001("TaxNum2"), 20) & _
                Space(22) & "05" & Space(1) & clsCommon.GetString(aRowINV007("REFNO"), 1)
            Return strTmp

        Catch ex As Exception
            Throw New Exception("Function: WriteA0401Head 訊息: " & ex.ToString)
        End Try
    End Function
    Private Function WriteB0301Header(ByVal aRowINV014 As DataRow, _
                                     ByVal aRowINV001 As DataRow) As String
        Try
            '#5945 原本是抓PAPERNO欄位,改抓AllowanceNo By Kin 2011/03/16
            strTmp = String.Empty
            strTmp = "M" & clsCommon.GetString(aRowINV014("ALLOWANCENO"), 16) & _
                clsCommon.GetString(aRowINV014("PAPERDATE"), 10) & _
                clsCommon.GetString(aRowINV001("BUSINESSID"), 10) & _
                clsCommon.GetString(aRowINV001("CompName"), 60) & Space(228) & _
                New String("0", 10) & clsCommon.GetString(Right(aRowINV014("INVID"), 4), 60) & _
                Space(228) & "2"
            Return strTmp
        Catch ex As Exception
            Throw New Exception("Function: WriteB0301Header 訊息: " & ex.ToString)
        End Try

    End Function
    Private Function WriteB0401Header(ByVal aRowInv014 As DataRow, _
                                      ByVal aRowInv001 As DataRow) As String
        Try
            Dim aDate As DateTime = DateTime.Parse(aRowInv014("UPTTIME").ToString)
            strTmp = String.Empty
            strTmp = "C" & clsCommon.GetString(aRowInv014("ALLOWANCENO"), 16) & _
                New String("0", 10) & clsCommon.GetString(aRowInv001("BUSINESSID"), 10) & _
                clsCommon.GetString(aRowInv014("PAPERDATE"), 10) & _
                clsCommon.GetString(Format(aDate, "yyyy/MM/dd"), 10) & _
                clsCommon.GetString(Format(aDate, "hh:mm:ss"), 8) & _
                clsCommon.GetString(aRowInv014("ObsoleteReason"), 200, AlignType.Left)
            Return strTmp
        Catch ex As Exception
            Throw New Exception("Function: WriteB0401Header 訊息: " & ex.ToString)
        End Try
    End Function
    Private Overloads Sub writeFile(ByVal WritePath As String,
                                    ByVal aFileType As InvFileType, ByVal aContText As String, _
                          ByVal aSeq As Integer)
        Dim aFileName As String = Format(FNow, "yyyyMMdd") & "-" & Format(Now, "HHmmss")
        aFileName = [Enum].GetName(GetType(InvFileType), aFileType) & _
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

    Private Overloads Sub writeFile(ByVal aFileType As InvFileType, ByVal aContText As String, _
                          ByVal aSeq As Integer)

        writeFile(FFilePath, aFileType, aContText, aSeq)
    End Sub
End Class
Public Enum InvFileType
    A0401 = 0
    A0501 = 1
    B0301 = 2
    B0401 = 3
    C0401 = 4
    C0501 = 5
    D0401 = 6
    D0501 = 7
    E0402 = 8
End Enum