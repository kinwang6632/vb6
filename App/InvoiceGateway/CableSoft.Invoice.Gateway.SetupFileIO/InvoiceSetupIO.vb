Imports System.IO
Imports System.Text

Public Class InvoiceSetupIO
    Implements IDisposable
    Private Shared XmlFileName As String = "Setting.bin"
    Private Const Key As String = "CABLESOFT1234567"
    Private xmlDoc As Xml.XmlDocument = Nothing
    Private Shared xmlDBIdDoc As Xml.XmlDocument = Nothing
    Private Shared mtxWait As New System.Threading.Mutex
    Private Const SODBParentName As String = "CableSoft/SODB"
    Private Const CommandParentName As String = "CableSoft/CommandType"
    Private Const RunCommandParentName As String = "CableSoft/RunCommand"
    Private Const SystemParentName As String = "CableSoft/System"
    Private Const SODBNoteName As String = "DB"
    Private Const CommandNoteName As String = "Command"
    Private Const RunCommandNoteName As String = "RunType"
    
    Sub New()
        XmlFileName = String.Format("{0}\{1}",
                                    Environment.CurrentDirectory, XmlFileName)
      
    End Sub
        Sub New(ByVal XmlFile As String)
            XmlFileName = XmlFile
        End Sub
        Public Function ReadRunCommand(ByVal IsEncryFile As Boolean) As DataTable
            Dim tbRunCommand As New DataTable("RunCommand")
            With tbRunCommand
                .Columns.Add(New DataColumn("CompID", GetType(String)))
                .Columns.Add(New DataColumn("CompName", GetType(String)))
                .Columns.Add(New DataColumn("CodeNo", GetType(String)))
                .Columns.Add(New DataColumn("CmdName", GetType(String)))
            .Columns.Add(New DataColumn("Runfrequency", GetType(String)))
            .Columns.Add(New DataColumn("ExportElectronPath", GetType(String)))
            End With
            Try
                Dim data As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
                Dim lstCompId As IList(Of Object) = ReadXmlNodeAttributes(RunCommandParentName, RunCommandNoteName, "CompID", IsEncryFile)
                Dim lstCompName As IList(Of Object) = ReadXmlNodeAttributes(RunCommandParentName, RunCommandNoteName, "CompDescript", IsEncryFile)
                Dim lstCodeNo As List(Of Object) = ReadXmlNodeAttributes(RunCommandParentName, RunCommandNoteName, "CmdType", IsEncryFile)
            Dim lstFrequency As List(Of Object) = ReadXmlNodeAttributes(RunCommandParentName, RunCommandNoteName, "Runfrequency", IsEncryFile)
            Dim lstExportElectronPath As List(Of Object) = ReadXmlNodeAttributes(RunCommandParentName, RunCommandNoteName, "ExportElectronPath", IsEncryFile)
            Dim lstCmdName As List(Of Object) = ReadXmlNodeValues(RunCommandParentName, RunCommandNoteName, IsEncryFile)

                If lstCompId IsNot Nothing Then
                    For i As Int32 = 0 To lstCompId.Count - 1
                        data.Clear()
                        data.Add("CompID", lstCompId.Item(i))
                        data.Add("CompName", lstCompName.Item(i))
                        data.Add("CodeNo", lstCodeNo.Item(i))
                        data.Add("CmdName", lstCmdName.Item(i))
                    data.Add("Runfrequency", lstFrequency.Item(i))
                    data.Add("ExportElectronPath", lstExportElectronPath.Item(i))
                        Dim rw As DataRow = tbRunCommand.NewRow
                        With rw
                            .Item("CompID") = data.Item("CompID")
                            .Item("CompName") = data.Item("CompName")
                            .Item("CodeNo") = data.Item("CodeNo")
                            .Item("CmdName") = data.Item("CmdName")
                        .Item("Runfrequency") = data.Item("Runfrequency")
                        .Item("ExportElectronPath") = data.Item("ExportElectronPath")
                        End With
                        tbRunCommand.Rows.Add(rw)
                    Next

                    tbRunCommand.AcceptChanges()
                End If
            Catch ex As Exception
                Throw
            End Try
            Return tbRunCommand
        End Function
    Public Function ReadSystem(ByVal IsEncryFile As Boolean) As DataTable
        Dim tbSystem As New DataTable("System")
        Try
            With tbSystem
                .Columns.Add(New DataColumn("MaxThreads", GetType(Integer)))
                .Columns.Add(New DataColumn("LoginUserId", GetType(String)))
                .Columns.Add(New DataColumn("LoginUserName", GetType(String)))
                .Columns.Add(New DataColumn("DestroyReason", GetType(String)))
                .Columns.Add(New DataColumn("IsLogSQL", GetType(Integer)))
                .Columns.Add(New DataColumn("LogReserveTime", GetType(Integer)))
            End With
            Dim MaxThreads As Object = ReadXmlNodeValue(SystemParentName, "MaxThread", IsEncryFile)
            Dim LoginUserId As Object = ReadXmlNodeValue(SystemParentName, "LoginUserId", IsEncryFile)
            Dim LoginName As Object = ReadXmlNodeValue(SystemParentName, "LoginUserName", IsEncryFile)
            Dim DestroyReason As Object = ReadXmlNodeValue(SystemParentName, "DestroyReason", IsEncryFile)
            Dim IsLogSQL As Object = ReadXmlNodeValue(SystemParentName, "IsLogSQL", IsEncryFile)
            Dim LogReserveTime As Object = ReadXmlNodeValue(SystemParentName, "LogReserveTime", IsEncryFile)
            Dim rw As DataRow = tbSystem.NewRow
            If MaxThreads IsNot Nothing Then
                rw.Item("MaxThreads") = Integer.Parse(MaxThreads.ToString)
            Else
                rw.Item("MaxThreads") = 20
            End If
            If (LoginUserId Is Nothing) OrElse (String.IsNullOrEmpty(LoginUserId.ToString)) Then
                LoginUserId = "Gateway"
            End If
            If (LoginName Is Nothing) OrElse (String.IsNullOrEmpty(LoginName.ToString)) Then
                LoginName = "Gateway"
            End If
            If IsLogSQL IsNot Nothing Then
                rw.Item("IsLogSQL") = Integer.Parse("0" & IsLogSQL.ToString)
            Else
                rw.Item("IsLogSQL") = 0
            End If
            If LogReserveTime IsNot Nothing Then
                rw.Item("LogReserveTime") = Integer.Parse("0" & LogReserveTime.ToString)
            Else
                rw.Item("LogReserveTime") = 1
            End If
            rw.Item("LoginUserId") = LoginUserId
            rw.Item("LoginUserName") = LoginName
            rw.Item("DestroyReason") = DestroyReason
            tbSystem.Rows.Add(rw)
            tbSystem.AcceptChanges()
        Catch ex As Exception
            Throw
        End Try
        Return tbSystem
    End Function
        Public Function ReadCommand(ByVal IsEncryFile As Boolean) As DataTable
            Dim tbCommand As New DataTable("Command")
            With tbCommand
                .Columns.Add(New DataColumn("CodeNo", GetType(String)))
                .Columns.Add(New DataColumn("CmdName", GetType(String)))
            End With
            Try
                Dim data As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
                Dim lstCodeNo As List(Of Object) = ReadXmlNodeAttributes(CommandParentName, CommandNoteName, "CodeNo", IsEncryFile)
                Dim lstCmdName As List(Of Object) = ReadXmlNodeValues(CommandParentName, CommandNoteName, IsEncryFile)
                If lstCodeNo IsNot Nothing Then
                    For i As Integer = 0 To lstCodeNo.Count - 1
                        data.Clear()
                        data.Add("CodeNo", lstCodeNo.Item(i))
                        data.Add("CmdName", lstCmdName.Item(i))
                        Dim rw As DataRow = tbCommand.NewRow
                        rw.Item("CodeNo") = data.Item("CodeNo")
                        rw.Item("CmdName") = data.Item("CmdName")
                        tbCommand.Rows.Add(rw)
                        tbCommand.AcceptChanges()
                    Next
                End If
            Catch ex As Exception
                Throw
            End Try
            Return tbCommand
        End Function

    Public Function ReadSO(ByVal IsEncryFile As Boolean) As DataTable
        Dim tbSO As New DataTable("SO")
        tbSO.Columns.Add(New DataColumn("SelectID", GetType(String)))
        tbSO.Columns.Add(New DataColumn("CompID", GetType(String)))
        tbSO.Columns.Add(New DataColumn("CompName", GetType(String)))
        tbSO.Columns.Add(New DataColumn("InvBackupPath", GetType(String)))
        tbSO.Columns.Add(New DataColumn("DBSid", GetType(String)))
        tbSO.Columns.Add(New DataColumn("DBUser", GetType(String)))
        tbSO.Columns.Add(New DataColumn("DBPassword", GetType(String)))
        tbSO.Columns.Add(New DataColumn("ComDBSid", GetType(String)))
        tbSO.Columns.Add(New DataColumn("ComDbUser", GetType(String)))
        tbSO.Columns.Add(New DataColumn("ComDBPassword", GetType(String)))
        tbSO.Columns.Add(New DataColumn("MisDBSid", GetType(String)))
        tbSO.Columns.Add(New DataColumn("MisDBUser", GetType(String)))
        tbSO.Columns.Add(New DataColumn("MisDBPassword", GetType(String)))
        tbSO.Columns.Add(New DataColumn("MisOwner", GetType(String)))
        tbSO.Columns.Add(New DataColumn("DBLink", GetType(String)))
        tbSO.Columns.Add(New DataColumn("MinDBPool", GetType(String)))
        tbSO.Columns.Add(New DataColumn("MaxDBPool", GetType(String)))
        tbSO.Columns.Add(New DataColumn("DBPoolLifetime", GetType(String)))
        tbSO.Columns.Add(New DataColumn("RunCmdTimer", GetType(Integer)))
        tbSO.Columns.Add(New DataColumn("LimtBeforeUpload", GetType(Integer)))
        tbSO.Columns.Add(New DataColumn("InvoiceType", GetType(String)))
        tbSO.Columns.Add(New DataColumn("RunCommands", GetType(String)))
        tbSO.Columns.Add(New DataColumn("IsEmailNotify", GetType(String)))
        tbSO.Columns.Add(New DataColumn("IsSmsNotify", GetType(String)))
        tbSO.Columns.Add(New DataColumn("IsTvMailNotify", GetType(String)))
        tbSO.Columns.Add(New DataColumn("IsCmNotify", GetType(String)))
        tbSO.Columns.Add(New DataColumn("SendOrd", GetType(String)))
        tbSO.Columns.Add(New DataColumn("AddSend", GetType(String)))
        tbSO.Columns.Add(New DataColumn("DonateMark", GetType(String)))
        tbSO.Columns.Add(New DataColumn("ExportPath", GetType(String)))
        tbSO.Columns.Add(New DataColumn("AutoDo", GetType(String)))
        tbSO.Columns.Add(New DataColumn("StartEmailNotify", GetType(String)))
        tbSO.Columns.Add(New DataColumn("StartSMSNotify", GetType(String)))
        tbSO.Columns.Add(New DataColumn("StartTVMailNotify", GetType(String)))
        tbSO.Columns.Add(New DataColumn("StartCMNotify", GetType(String)))
        tbSO.Columns.Add(New DataColumn("EndCMNotify", GetType(String)))
        tbSO.Columns.Add(New DataColumn("CompType", GetType(String)))
        tbSO.Columns.Add(New DataColumn("MaskInvNo", GetType(String)))
        tbSO.Columns.Add(New DataColumn("StartNotifyPrize", GetType(String)))
        tbSO.Columns.Add(New DataColumn("MaxThread", GetType(Integer)))

        
        Try
            Dim data As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
            Dim lstDBValue As List(Of Object) = ReadXmlNodeAttributes(SODBParentName, SODBNoteName, "SelectID", IsEncryFile)
            Dim lstCompCode As List(Of Object) = ReadXmlNodeAttributes(SODBParentName, SODBNoteName, "CompID", IsEncryFile)
            Dim lstCompName As List(Of Object) = ReadXmlNodeValues(SODBParentName, SODBNoteName, IsEncryFile)
            Dim lstSID As List(Of Object) = ReadXmlNodeAttributes(SODBParentName, SODBNoteName, "DBSid", IsEncryFile)
            Dim lstUSER As List(Of Object) = ReadXmlNodeAttributes(SODBParentName, SODBNoteName, "DBUser", IsEncryFile)
            Dim lstPassword As List(Of Object) = ReadXmlNodeAttributes(SODBParentName, SODBNoteName, "DBPassword", IsEncryFile)
            Dim lstComDBSid As List(Of Object) = ReadXmlNodeAttributes(SODBParentName, SODBNoteName, "ComDBSid", IsEncryFile)
            Dim lstComDBUser As List(Of Object) = ReadXmlNodeAttributes(SODBParentName, SODBNoteName, "ComDbUser", IsEncryFile)
            Dim lstComDBPassword As List(Of Object) = ReadXmlNodeAttributes(SODBParentName, SODBNoteName, "ComDBPassword", IsEncryFile)
            Dim lstMisDBSid As List(Of Object) = ReadXmlNodeAttributes(SODBParentName, SODBNoteName, "MisDBSid", IsEncryFile)
            Dim lstMisDBUserId As List(Of Object) = ReadXmlNodeAttributes(SODBParentName, SODBNoteName, "MisDBUser", IsEncryFile)
            Dim lstMisDBPassword As List(Of Object) = ReadXmlNodeAttributes(SODBParentName, SODBNoteName, "MisDBPassword", IsEncryFile)
            Dim lstMisOwner As List(Of Object) = ReadXmlNodeAttributes(SODBParentName, SODBNoteName, "MisOwner", IsEncryFile)
            Dim lstDBLink As List(Of Object) = ReadXmlNodeAttributes(SODBParentName, SODBNoteName, "DBLink", IsEncryFile)
            Dim lstRunTime As List(Of Object) = ReadXmlNodeAttributes(SODBParentName, SODBNoteName, "RunCmdTimer", IsEncryFile)
            Dim lstInvoiceType As List(Of Object) = ReadXmlNodeAttributes(SODBParentName, SODBNoteName, "InvoiceType", IsEncryFile)
            Dim lstMinDBPool As List(Of Object) = ReadXmlNodeAttributes(SODBParentName, SODBNoteName, "MinDBPool", IsEncryFile)
            Dim lstMaxDBPool As List(Of Object) = ReadXmlNodeAttributes(SODBParentName, SODBNoteName, "MaxDBPool", IsEncryFile)
            Dim lstDBPoolLifetime As List(Of Object) = ReadXmlNodeAttributes(SODBParentName, SODBNoteName, "DBPoolLifetime", IsEncryFile)
            Dim lstRunCommands As List(Of Object) = ReadXmlNodeAttributes(SODBParentName, SODBNoteName, "RunCommands", IsEncryFile)
            Dim lstLimtBeforeUpload As List(Of Object) = ReadXmlNodeAttributes(SODBParentName, SODBNoteName, "LimtBeforeUpload", IsEncryFile)
            Dim lstInvBackupPath As List(Of Object) = ReadXmlNodeAttributes(SODBParentName, SODBNoteName, "InvBackupPath", IsEncryFile)
            Dim lstIsEmailNotify As List(Of Object) = ReadXmlNodeAttributes(SODBParentName, SODBNoteName, "IsEmailNotify", IsEncryFile)
            Dim lstIsSmsNotify As List(Of Object) = ReadXmlNodeAttributes(SODBParentName, SODBNoteName, "IsSmsNotify", IsEncryFile)
            Dim lstIsTvMailNotify As List(Of Object) = ReadXmlNodeAttributes(SODBParentName, SODBNoteName, "IsTvMailNotify", IsEncryFile)
            Dim lstIsCmNotify As List(Of Object) = ReadXmlNodeAttributes(SODBParentName, SODBNoteName, "IsCmNotify", IsEncryFile)
            Dim lstSendOrd As List(Of Object) = ReadXmlNodeAttributes(SODBParentName, SODBNoteName, "SendOrd", IsEncryFile)
            Dim lstAddSend As List(Of Object) = ReadXmlNodeAttributes(SODBParentName, SODBNoteName, "AddSend", IsEncryFile)
            Dim lstDonateMark As List(Of Object) = ReadXmlNodeAttributes(SODBParentName, SODBNoteName, "DonateMark", IsEncryFile)
            Dim lstExportPath As List(Of Object) = ReadXmlNodeAttributes(SODBParentName, SODBNoteName, "ExportPath", IsEncryFile)
            Dim lstAutoDo As List(Of Object) = ReadXmlNodeAttributes(SODBParentName, SODBNoteName, "AutoDo", IsEncryFile)
            Dim lstStartEmailNotify As List(Of Object) = ReadXmlNodeAttributes(SODBParentName, SODBNoteName, "StartEmailNotify", IsEncryFile)
            Dim lstStartSMSNotify As List(Of Object) = ReadXmlNodeAttributes(SODBParentName, SODBNoteName, "StartSMSNotify", IsEncryFile)
            Dim lstStartTVMailNotify As List(Of Object) = ReadXmlNodeAttributes(SODBParentName, SODBNoteName, "StartTVMailNotify", IsEncryFile)
            Dim lstStartCMNotify As List(Of Object) = ReadXmlNodeAttributes(SODBParentName, SODBNoteName, "StartCMNotify", IsEncryFile)
            Dim lstEndCMNotify As List(Of Object) = ReadXmlNodeAttributes(SODBParentName, SODBNoteName, "EndCMNotify", IsEncryFile)
            Dim lstCompType As List(Of Object) = ReadXmlNodeAttributes(SODBParentName, SODBNoteName, "CompType", IsEncryFile)
            Dim lstMaskInvNo As List(Of Object) = ReadXmlNodeAttributes(SODBParentName, SODBNoteName, "MaskInvNo", IsEncryFile)
            Dim lstStartNotifyPrize As List(Of Object) = ReadXmlNodeAttributes(SODBParentName, SODBNoteName, "StartNotifyPrize", IsEncryFile)

           
            If lstDBValue IsNot Nothing Then
                For i As Int32 = 0 To lstDBValue.Count - 1
                    data.Clear()
                    data.Add("SelectID", lstDBValue.Item(i))
                    data.Add("CompID", lstCompCode.Item(i))
                    data.Add("CompName", lstCompName.Item(i))
                    data.Add("DBSid", lstSID.Item(i))
                    data.Add("DBUser", lstUSER.Item(i))
                    data.Add("DBPassword", lstPassword.Item(i))
                    data.Add("ComDBSid", lstComDBSid.Item(i))
                    data.Add("ComDbUser", lstComDBUser.Item(i))
                    data.Add("ComDBPassword", lstComDBPassword.Item(i))
                    data.Add("RunCmdTimer", lstRunTime.Item(i))
                    data.Add("InvoiceType", lstInvoiceType.Item(i))
                    data.Add("MinDBPool", lstMinDBPool.Item(i))
                    data.Add("MaxDBPool", lstMaxDBPool.Item(i))
                    data.Add("DBPoolLifetime", lstDBPoolLifetime.Item(i))
                    data.Add("RunCommands", lstRunCommands.Item(i))
                    data.Add("LimtBeforeUpload", lstLimtBeforeUpload.Item(i))
                    data.Add("MisDBSid", lstMisDBSid.Item(i))
                    data.Add("MisDBUser", lstMisDBUserId.Item(i))
                    data.Add("MisDBPassword", lstMisDBPassword.Item(i))
                    data.Add("MisOwner", lstMisOwner.Item(i))
                    data.Add("DBLink", lstDBLink.Item(i))
                    data.Add("InvBackupPath", lstInvBackupPath.Item(i))
                    data.Add("IsEmailNotify", lstIsEmailNotify.Item(i))
                    data.Add("IsSmsNotify", lstIsSmsNotify.Item(i))
                    data.Add("IsTvMailNotify", lstIsTvMailNotify.Item(i))
                    data.Add("IsCmNotify", lstIsCmNotify.Item(i))
                    data.Add("SendOrd", lstSendOrd.Item(i))
                    data.Add("AddSend", lstAddSend.Item(i))
                    data.Add("DonateMark", lstDonateMark.Item(i))
                    data.Add("ExportPath", lstExportPath.Item(i))
                    data.Add("AutoDo", lstAutoDo.Item(i))
                    data.Add("StartEmailNotify", lstStartEmailNotify.Item(i))
                    data.Add("StartSMSNotify", lstStartSMSNotify.Item(i))
                    data.Add("StartTVMailNotify", lstStartTVMailNotify.Item(i))
                    data.Add("StartCMNotify", lstStartCMNotify.Item(i))
                    data.Add("EndCMNotify", lstEndCMNotify.Item(i))
                    data.Add("CompType", lstCompType.Item(i))
                    data.Add("MaskInvNo", lstMaskInvNo.Item(i))
                    data.Add("StartNotifyPrize", lstStartNotifyPrize.Item(i))
                 

                    Dim rwNew As DataRow = tbSO.NewRow()
                    rwNew.Item("SelectID") = data.Item("SelectID")
                    rwNew.Item("CompID") = data.Item("CompID")
                    rwNew.Item("CompName") = data.Item("CompName")
                    rwNew.Item("DBSid") = data.Item("DBSid")
                    rwNew.Item("DBUser") = data.Item("DBUser")
                    rwNew.Item("DBPassword") = data.Item("DBPassword")
                    rwNew.Item("ComDBSid") = data.Item("ComDBSid")
                    rwNew.Item("ComDbUser") = data.Item("ComDbUser")
                    rwNew.Item("ComDBPassword") = data.Item("ComDBPassword")
                    rwNew.Item("RunCmdTimer") = data.Item("RunCmdTimer")
                    rwNew.Item("MinDBPool") = data.Item("MinDBPool")
                    rwNew.Item("MaxDBPool") = data.Item("MaxDBPool")
                    rwNew.Item("DBPoolLifetime") = data.Item("DBPoolLifetime")
                    rwNew.Item("RunCommands") = data.Item("RunCommands")
                    rwNew.Item("MisDBSid") = data.Item("MisDBSid")
                    rwNew.Item("MisDBUser") = data.Item("MisDBUser")
                    rwNew.Item("MisDBPassword") = data.Item("MisDBPassword")
                    rwNew.Item("MisOwner") = data.Item("MisOwner")
                    rwNew.Item("DBLink") = data.Item("DBLink")
                    rwNew.Item("InvBackupPath") = data.Item("InvBackupPath")
                    rwNew.Item("IsEmailNotify") = data.Item("IsEmailNotify")
                    rwNew.Item("IsSmsNotify") = data.Item("IsSmsNotify")
                    rwNew.Item("IsTvMailNotify") = data.Item("IsTvMailNotify")
                    rwNew.Item("IsCmNotify") = data.Item("IsCmNotify")
                    rwNew.Item("SendOrd") = data.Item("SendOrd")
                    rwNew.Item("AddSend") = data.Item("AddSend")
                    rwNew.Item("DonateMark") = data.Item("DonateMark")
                    rwNew.Item("ExportPath") = data.Item("ExportPath")
                    rwNew.Item("AutoDo") = data.Item("AutoDo")
                    rwNew.Item("StartEmailNotify") = data.Item("StartEmailNotify")
                    rwNew.Item("StartSMSNotify") = data.Item("StartSMSNotify")
                    rwNew.Item("StartTVMailNotify") = data.Item("StartTVMailNotify")
                    rwNew.Item("StartCMNotify") = data.Item("StartCMNotify")
                    rwNew.Item("EndCMNotify") = data.Item("EndCMNotify")
                    rwNew.Item("CompType") = data.Item("CompType")
                    rwNew.Item("MaskInvNo") = data.Item("MaskInvNo")
                    rwNew.Item("StartNotifyPrize") = data.Item("StartNotifyPrize")

                    If (DBNull.Value.Equals(rwNew.Item("CompType"))) OrElse
                        (String.IsNullOrEmpty(rwNew.Item("CompType"))) Then
                        rwNew.Item("CompType") = "0.個別公司"
                    End If

                    If (DBNull.Value.Equals(rwNew.Item("DonateMark"))) OrElse
                        (String.IsNullOrEmpty(rwNew.Item("DonateMark"))) Then
                        rwNew.Item("DonateMark") = "2.全部"
                    End If
                    If (DBNull.Value.Equals(rwNew.Item("AutoDo"))) OrElse
                        (String.IsNullOrEmpty(rwNew.Item("AutoDo"))) Then
                        rwNew.Item("AutoDo") = "1.B07進入自動執行"
                    End If

                    If DBNull.Value.Equals(rwNew("RunCmdTimer")) OrElse String.IsNullOrEmpty(rwNew.Item("RunCmdTimer").ToString) Then
                        rwNew.Item("RunCmdTimer") = "1"
                    End If
                    rwNew.Item("InvoiceType") = data.Item("InvoiceType")
                    If (data.Item("LimtBeforeUpload") Is Nothing) OrElse (String.IsNullOrEmpty(data.Item("LimtBeforeUpload").ToString)) Then
                        data.Item("LimtBeforeUpload") = 15
                    End If
                    rwNew.Item("LimtBeforeUpload") = data.Item("LimtBeforeUpload")
                    tbSO.Rows.Add(rwNew)
                    tbSO.AcceptChanges()
                Next
            End If
        Catch ex As Exception
            Throw ex
        End Try
        Return tbSO
    End Function
        Public Shared Function WriteXmlFile(ByVal FileName As String,
                                  ByVal stmWrite As MemoryStream, ByVal IsEncryFile As Boolean) As Boolean
            Try
                Using stm As StreamWriter = New StreamWriter(FileName, False, Encoding.UTF8)
                    If IsEncryFile Then
                        'Dim strWrite As String = CableSoft.CAS.Encry.Encryption.Encrypt3DES(
                        '    Convert.ToBase64String(Encoding.UTF8.GetBytes(Encoding.UTF8.GetString(stmWrite.ToArray))), Key)
                        Dim strWrite As String = CableSoft.NSTV.Encry.Encryption.Encrypt3DES(
                           Encoding.UTF8.GetString(stmWrite.ToArray), Key)
                        stm.WriteLine(strWrite)
                        stm.Flush()
                        stm.Close()
                    Else
                        stm.WriteLine(Encoding.UTF8.GetString(stmWrite.ToArray))
                        stm.Flush()
                        stm.Close()
                    End If
            End Using

            Catch ex As Exception
                Throw
            End Try
            Return True
        End Function
        Public Function ReadXmlNodeAttributes(ByVal ParentNode As String, ByVal NodeName As String,
                                  ByVal AttributeName As String, ByVal IsDecry As Boolean) As List(Of Object)
            Dim aResult As New List(Of Object)
            Try
                If xmlDoc Is Nothing Then
                    xmlDoc = New Xml.XmlDocument()
                    xmlDoc.XmlResolver = Nothing
                    If Not File.Exists(XmlFileName) Then
                        Return Nothing
                    End If

                    Using stm As Stream = New FileStream(XmlFileName,
                                                         FileMode.OpenOrCreate, FileAccess.ReadWrite)


                        Using rder As New System.IO.StreamReader(stm, New UTF8Encoding(False))
                            If IsDecry Then
                                'Dim xmlStr As String = UTF8Encoding.UTF8.GetString(Convert.FromBase64String(CableSoft.NSTV.Encry.Encryption.Decrypt3DES(rder.ReadToEnd(), Key)))
                                Dim redString As String = rder.ReadToEnd()
                                Dim xmlStr As String = CableSoft.NSTV.Encry.Encryption.Decrypt3DES(redString, Key)
                                xmlDoc.LoadXml(xmlStr)
                            Else
                                xmlDoc.LoadXml(rder.ReadToEnd())
                            End If
                        End Using
                    End Using
                End If

                Dim Node As Xml.XmlNode = xmlDoc.SelectSingleNode(ParentNode)
                If Node Is Nothing Then
                    Return Nothing
                End If
                Dim lstChildNode As Xml.XmlNodeList = Node.SelectNodes(NodeName)

                If lstChildNode Is Nothing OrElse lstChildNode.Count = 0 Then
                    Return Nothing
                End If
                For Each aNode As Xml.XmlNode In lstChildNode
                    For i As Int32 = 0 To aNode.ChildNodes.Count - 1
                        Dim aAttribute As Xml.XmlAttribute = aNode.Attributes(AttributeName)
                        If aAttribute IsNot Nothing Then
                            aResult.Add(aAttribute.Value)
                        Else
                            aResult.Add(String.Empty)
                        End If
                    Next
                Next
            Catch ex As Exception
                Throw ex

            End Try
            Return aResult
        End Function
        Public Function ReadXmlNodeValues(ByVal ParentNode As String, ByVal NodeName As String,
                              ByVal IsDecry As Boolean) As List(Of Object)
            Dim aResult As New List(Of Object)
            Try
                If xmlDoc Is Nothing Then
                    xmlDoc = New Xml.XmlDocument()
                    xmlDoc.XmlResolver = Nothing


                    Using stm As Stream = New FileStream(XmlFileName,
                                                         FileMode.OpenOrCreate, FileAccess.ReadWrite)


                        Using rder As New System.IO.StreamReader(stm, New UTF8Encoding(False))
                            If IsDecry Then
                                'Dim xmlStr As String = UTF8Encoding.UTF8.GetString(Convert.FromBase64String(CableSoft.NSTV.Encry.Encryption.Decrypt3DES(rder.ReadToEnd(), Key)))
                                Dim redString As String = rder.ReadToEnd()
                                Dim xmlStr As String = CableSoft.NSTV.Encry.Encryption.Decrypt3DES(redString, Key)
                                xmlDoc.LoadXml(xmlStr)
                            Else
                                xmlDoc.LoadXml(rder.ReadToEnd())
                            End If
                        End Using
                    End Using
                End If

                Dim Node As Xml.XmlNode = xmlDoc.SelectSingleNode(ParentNode)
                If Node Is Nothing Then
                    Return Nothing
                End If
                Dim lstChildNode As Xml.XmlNodeList = Node.SelectNodes(NodeName)


                If lstChildNode Is Nothing OrElse lstChildNode.Count = 0 Then
                    Return Nothing
                End If
                For Each aNode As Xml.XmlNode In lstChildNode
                    For i As Int32 = 0 To aNode.ChildNodes.Count - 1
                        aResult.Add(aNode.ChildNodes(i).InnerText)
                    Next

                Next
                'aResult = ChildNode.InnerText
            Catch ex As Exception
                Throw
                'MessageBox.Show(ex.ToString, "訊息", MessageBoxButtons.OK)
            End Try
            Return aResult
        End Function
        Public Function ReadXmlNodeValue(ByVal ParentNode As String, ByVal NodeName As String,
                                  ByVal IsDecry As Boolean) As Object
            Dim aResult As String = Nothing
            Try
                If xmlDoc Is Nothing Then
                    xmlDoc = New Xml.XmlDocument()
                    xmlDoc.XmlResolver = Nothing


                    Using stm As Stream = New FileStream(XmlFileName,
                                                         FileMode.OpenOrCreate, FileAccess.ReadWrite)


                        Using rder As New System.IO.StreamReader(stm, New UTF8Encoding(False))
                            If IsDecry Then

                                'Dim xmlStr As String = CableSoft.CAS.CryptUtil.CryptUtil.Decrypt(rder.ReadToEnd())
                                'Dim xmlStr As String = UTF8Encoding.UTF8.GetString(Convert.FromBase64String(CableSoft.NSTV.Encry.Encryption.Decrypt3DES(rder.ReadToEnd(), Key)))
                                Dim redString As String = rder.ReadToEnd()
                                Dim xmlStr As String = CableSoft.NSTV.Encry.Encryption.Decrypt3DES(redString, Key)
                                xmlDoc.LoadXml(xmlStr)
                            Else
                                xmlDoc.LoadXml(rder.ReadToEnd())
                                '照他的嘗試，他只要先將第一個字元強迫刪掉，
                                '也就是寫成 xmldoc.LoadXml(s.Substring(1)); ，
                                '問題雖然解決，但確不知道為什麼，
                                '而且這樣寫很low，我們討論了一會兒，
                                '猜測是否UTF8文件的前置表頭(BOM)搞鬼，
                                '沒多久他就找到了另一個解決方法，
                                '將 settings.Encoding = Encoding.UTF8;
                                ' 改成 settings.Encoding = new UTF8Encoding(false);，
                                '雖然目的都是要避開因為utf8表頭造成的解析錯誤，
                                '但至少這樣的寫法是明確且易懂的。
                                '我個人會希望 XmlDocument 要再聰明一點。
                            End If

                        End Using
                    End Using
                End If

                'Debug.Print(ReadNodeName(xmlDoc.ChildNodes, NodeName))
                'ParentNode = ReadNodeName(xmlDoc.ChildNodes, ParentNode)
                'NodeName = ReadNodeName(xmlDoc.ChildNodes, NodeName)
                Dim Node As Xml.XmlNode = xmlDoc.SelectSingleNode(ParentNode)
                If Node Is Nothing Then
                    Return Nothing
                End If
                Dim ChildNode As Xml.XmlNode = Node.SelectSingleNode(NodeName)

                If ChildNode Is Nothing Then
                    Return Nothing
                End If
                aResult = ChildNode.InnerText
            Catch ex As Exception
            Throw New Exception(ex.ToString & Environment.NewLine & "File = " & XmlFileName)
                'MessageBox.Show(ex.ToString, "訊息", MessageBoxButtons.OK)
            End Try
            Return aResult
        End Function

#Region "IDisposable Support"
        Private disposedValue As Boolean ' 偵測多餘的呼叫

        ' IDisposable
        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not Me.disposedValue Then
                If disposing Then
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



