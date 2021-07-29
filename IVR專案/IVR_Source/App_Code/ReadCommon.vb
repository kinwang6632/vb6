Imports Microsoft.VisualBasic
Imports System.IO
Imports System.Xml
Imports Oracle.DataAccess.Client
#Region "讀取系統對應的資料"
Public Class ReadCommon
    Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
    Private intCompCode As Integer = 0
    Private strOwnerName As String = String.Empty
    Private strCn As String = String.Empty
    Private strFloor As String = String.Empty
    Private strServiceType As String = String.Empty
    Private intCustId As String = String.Empty
    Private strCmd As String = String.Empty
    Private blnIsXML As Boolean = False
    Private blnTestXML As Boolean = False
    Private strXMlRead As New XmlDocument()
    Private strSysPara As String = String.Empty
    Private strIniFile As String = String.Empty
    Private strRetXml As String = String.Empty
    Private strTel As String = String.Empty
    'Private sbDebug As StreamWriter
    Private sbDebug As StringBuilder
    ''' <summary>
    ''' 物件初始化
    ''' </summary>
    ''' <param name="strXml">接收到的XML字串</param>
    ''' <param name="sParaFile">虛擬目錄下的加密連線檔案(需包含路徑)</param>
    ''' <param name="sIniFile">虛擬目錄下的Ini檔案(需包含路徑)</param>
    Sub New(ByVal strXml As String, ByVal sParaFile As String, ByVal sIniFile As String)
        strSysPara = sParaFile
        strIniFile = sIniFile
        LoadXmlFile(strXml)
        GetXmlPara()


    End Sub
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="strXml">接收到的XML字串</param>
    ''' <param name="sParaFile">虛擬目錄下的加密連線檔案(需包含路徑)</param>
    ''' <param name="sIniFile">虛擬目錄下的Ini檔案(需包含路徑)</param>
    ''' <param name="sb">除錯模式(StreamWriter)</param>
    ''' <remarks>會將錯誤訊息寫入文字檔</remarks>
    Sub New(ByVal strXml As String, ByVal sParaFile As String, ByVal sIniFile As String, ByRef sb As StringBuilder)
        sbDebug = sb
        strSysPara = sParaFile
        strIniFile = sIniFile
        LoadXmlFile(strXml)
        GetXmlPara()


    End Sub
    '取得XML必要參數
    Private Function GetXmlPara() As Boolean
        Try
            strCmd = ReadXml(strXMlRead, "CMD")
            strFloor = ReadXml(strXMlRead, "Floor")
            strTel = ReadXml(strXMlRead, "Tel")

            If strCmd = String.Empty OrElse strTel = String.Empty OrElse intCompCode = 0 Then
                blnTestXML = False
                If Not sbDebug Is Nothing Then
                    sbDebug.AppendLine("[GetXmlPara] 失敗")
                    sbDebug.AppendLine("[GetXmlPara] strCmd:" & strCmd & "  strTel:" & strTel & "  intCompCode:" & intCompCode.ToString & Environment.NewLine)
                End If
                Return False
                Exit Function
            End If
            If strCn = String.Empty Then
                If Not sbDebug Is Nothing Then
                    sbDebug.AppendLine("[GetXmlPara] 失敗 連接字串為空白" & Environment.NewLine)
                End If
                blnTestXML = False
                Return False
            Else
                Return True
            End If
        Catch ex As Exception
            blnTestXML = False
            If Not sbDebug Is Nothing Then
                sbDebug.AppendLine("[GetXmlPara] 載入失敗" & Environment.NewLine)
            End If
            Return False
        End Try
    End Function

    ''' <summary>
    '''測試參數檔的連線字串和公司別是否存在
    ''' </summary>
    ''' <param name="strFile">虛擬目錄的參數檔案</param>
    ''' <param name="intComp">XML字串的公司別</param>
    ''' <returns>布林值</returns>
    Private Function OpenConnect(ByVal strFile As String, ByVal intComp As Integer) As Boolean
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
                If aHtb.Contains(intComp) Then
                    intCompCode = intComp
                    strCn = aHtb.Item(intComp)
                    Return True
                Else
                    intCompCode = 0
                    strCn = String.Empty
                    If Not sbDebug Is Nothing Then
                        sbDebug.AppendLine("[OpenConnect]  失敗")
                        sbDebug.AppendLine("[OpenConnect]  strFile:" & strFile & "  intComp:" & intComp)
                        sbDebug.AppendLine("[OpenConnect]  strCn:" & strCn & "  intComp:" & intCompCode & Environment.NewLine)
                    End If
                    Return False
                End If
            End Using
        Catch ex As Exception
            strCn = String.Empty
            intCompCode = 0
            If Not sbDebug Is Nothing Then
                sbDebug.AppendLine("[OpenConnect] 載入失敗" & Environment.NewLine)
            End If
            Return False
        End Try
    End Function
    Private Sub LoadXmlFile(ByVal strXml As String)
        If (Not strXml Is Nothing) AndAlso (strXml <> String.Empty) Then
            Try
                strXMlRead.LoadXml(strXml)
                blnIsXML = True
                intCompCode = Int32.Parse(ReadXml(strXMlRead, "Company"))
                '測試XML的CompCode與參數檔案是否吻合,並檢查連線參數
                If Not OpenConnect(strSysPara, intCompCode) Then
                    blnTestXML = False

                Else
                    '檢查XML檔是否有包含應有的參數
                    If GetXmlPara() Then
                        blnTestXML = True
                    Else
                        blnTestXML = False
                    End If
                End If
                If Not blnTestXML Then
                    WriteXml(strXMlRead, "Status", "0")
                    strRetXml = strXMlRead.InnerText
                    Dim mstm As New MemoryStream()
                    strXMlRead.Save(mstm)
                    mstm.Flush()
                    mstm.Position = 0
                    Dim stm As New StreamReader(mstm)
                    strRetXml = stm.ReadToEnd
                End If
            Catch ex As Exception
                strRetXml = strXml
                blnIsXML = False
                blnTestXML = False
                If Not sbDebug Is Nothing Then
                    sbDebug.AppendLine("LoadXmlFile 載入失敗" & Environment.NewLine)
                End If
            End Try
        End If
    End Sub
    Public Shared Sub WriteXml(ByVal oXmlDocument As XmlDocument, ByVal strKey As String, ByVal strValue As String)
        On Error Resume Next
        Dim xmlNodeValue As XmlNode
        xmlNodeValue = oXmlDocument.SelectSingleNode("CableSoft/" & strKey)
        If Not xmlNodeValue Is Nothing Then
            xmlNodeValue.InnerText = strValue
        End If
    End Sub
    Private Function ReadXml(ByVal oXmlDocument As XmlDocument, ByVal strKey As String) As String
        Dim xmlNodeValue As XmlNode

        Try
            xmlNodeValue = oXmlDocument.SelectSingleNode("CableSoft/" & strKey)
            If Not xmlNodeValue Is Nothing Then
                If (xmlNodeValue.InnerText <> String.Empty) AndAlso (Not xmlNodeValue.InnerText Is Nothing) Then
                    Return xmlNodeValue.InnerText
                Else
                    Return String.Empty
                End If
            Else
                Return String.Empty
            End If
        Catch ex As Exception
            If Not sbDebug Is Nothing Then
                sbDebug.AppendLine("[ReadXml] 載入失敗" & Environment.NewLine)
            End If
            Return String.Empty
        End Try
    End Function

    '''<summary>取得OwnerName</summary>
    Private Function GetOwner() As String
        Dim strOwner As String = String.Empty
        Try
            If strCn <> String.Empty Then
                Try
                    Using cn As New OracleConnection(strCn)
                        cn.Open()
                        Using cmd As New OracleCommand("Select TableOwner From CD039 Where CodeNo=" & intCompCode, cn)
                            cmd.CommandType = Data.CommandType.Text
                            strOwner = CType(cmd.ExecuteScalar, String).ToString()
                            If strOwner <> String.Empty AndAlso Not strOwner Is Nothing Then
                                Return strOwner & "."
                            Else
                                Return String.Empty
                            End If
                        End Using
                        cn.Close()
                    End Using
                Catch ex As Exception
                    If Not sbDebug Is Nothing Then
                        sbDebug.AppendLine("[GetOwner] 開啟Oracle失敗" & Environment.NewLine)
                    End If
                    Return String.Empty
                End Try
            Else
                If Not sbDebug Is Nothing Then
                    sbDebug.AppendLine("[GetOwner] 錯誤 連接串字為空值" & Environment.NewLine)
                End If
                Return String.Empty

            End If
        Catch ex As Exception
            If Not sbDebug Is Nothing Then
                sbDebug.AppendLine("[GetOwner] 載入失敗" & Environment.NewLine)
            End If
            Return String.Empty
        End Try
    End Function
    Public ReadOnly Property OwnerName() As String
        Get
            'Return GetOwner()
            Return String.Empty
        End Get
    End Property

    ''' <summary>
    ''' XML的Tel
    ''' </summary>
    Public ReadOnly Property Tel() As String
        Get
            Return strTel
        End Get
    End Property


    ''' <summary>
    ''' XML的Cmd
    ''' </summary>
    Public ReadOnly Property Cmd() As String
        Get
            Return strCmd
        End Get
    End Property

    ''' <summary>
    ''' XML的Floor
    ''' </summary>
    Public ReadOnly Property Floor() As String
        Get
            '#3966 如果為0要當做是Empty處理 By Kin 2008/06/09
            If strFloor = "0" OrElse strFloor = String.Empty Then
                Return String.Empty
            Else
                Return strFloor
            End If

        End Get
    End Property

    ''' <summary>
    ''' XML的CustId
    ''' </summary>
    Public ReadOnly Property CustId() As Integer
        Get
            Return intCustId
        End Get
    End Property

    ''' <summary>
    ''' 錯誤的XML
    ''' </summary>
    Public ReadOnly Property RetunXml() As String
        Get
            Return strRetXml
        End Get
    End Property
    ''' <summary>
    ''' Oracle連線字串
    ''' </summary>
    Public ReadOnly Property OraConnstring() As String
        Get
            Return strCn
        End Get
    End Property
    ''' <summary>
    ''' 公司別
    ''' </summary>
    Public ReadOnly Property CompCode() As Integer
        Get
            Return intCompCode
        End Get
    End Property
    ''' <summary>
    ''' 判斷接收的資料是否符合XML格式
    ''' </summary>
    Public ReadOnly Property IsXML() As Boolean
        Get
            Return blnIsXML
        End Get
    End Property
    ''' <summary>
    ''' 接數的XML資料流內容是否都有包含符合的參數
    ''' </summary>
    Public ReadOnly Property TestXML() As Boolean
        Get
            Return blnTestXML
        End Get
    End Property
End Class
#End Region
