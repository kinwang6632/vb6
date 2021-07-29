Imports System.Xml
Imports System.IO
Imports Oracle.DataAccess.Client

Partial Class _Default
    Inherits System.Web.UI.Page
    Private sSysPara As String = Server.MapPath("\\IVR\ConPara.dat")
    Private sIniFile As String = Server.MapPath("\\IVR\apply.Ini")
    Private sDebugFile As String = Server.MapPath("\\IVR\Debug.Txt")
    Private strCompCode As String = String.Empty
    Private strLanguage As String = String.Empty
    Private strFunc As String = String.Empty
    Private strTel As String = String.Empty
    Private strHistory As String = String.Empty
    Private strFloor As String = String.Empty
    Private strCMD As String = String.Empty
    Private strXML As String = String.Empty
    Private blnDebug As Boolean = False
    Private strDebug As String = String.Empty
    Const sQryPara As String = "ID"
    'Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Long
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim sRet As String = Space(1024)
        Dim strFromData As String = String.Empty

        '制訂XML格式
        strXML = "<?xml version=" & """1.0""" & " encoding=" & """UTF-8"" ?" & ">" & vbCrLf & _
                    "<CableSoft>" & vbCrLf & _
                    "<AGENT_ID></AGENT_ID>" & vbCrLf & _
                    "<EXT></EXT>" & vbCrLf & _
                    "<PORT></PORT>" & vbCrLf & _
                    "<RecURL></RecURL>" & vbCrLf & _
                    "<InTel></InTel>" & vbCrLf & _
                    "<ObTel></ObTel>" & vbCrLf & _
                    "<CustID></CustID>" & vbCrLf & _
                    "<CMD></CMD>" & vbCrLf & _
                    "<CMD_Value></CMD_Value>" & vbCrLf & _
                    "<S-EXT></S-EXT>" & vbCrLf & _
                    "<D-EXT></D-EXT>" & vbCrLf & _
                    "<CSID></CSID>" & vbCrLf & _
                    "<SERVICETYPE></SERVICETYPE>" & vbCrLf & _
                    "<SEQNO></SEQNO>" & vbCrLf & _
                    "<EXTLOG></EXTLOG>" & vbCrLf & _
                    "<Para1></Para1>" & vbCrLf & _
                    "<Para2></Para2>" & vbCrLf & _
                    "<Para3></Para3>" & vbCrLf & _
                    "<Company></Company>" & vbCrLf & _
                    "<Language></Language>" & vbCrLf & _
                    "<Func></Func>" & vbCrLf & _
                    "<History></History>" & vbCrLf & _
                    "<Tel></Tel>" & vbCrLf & _
                    "<Floor></Floor>" & vbCrLf & _
                    "<Status></Status>" & vbCrLf & _
                    "<Audio></Audio>" & vbCrLf & _
                    "<Payment></Payment>" & vbCrLf & _
                    "</CableSoft>"

        '對方傳進來的資料
        strFromData = Request.Url.Query
        Dim stmDug As StreamWriter = Nothing
        If strFromData <> String.Empty Then
            If ReadPara() Then
                Dim asys As ReadCommon
                If blnDebug Then
                    stmDug = New StreamWriter(sDebugFile, False, Encoding.Default)
                    asys = New ReadCommon(strXML, sSysPara, sIniFile, stmDug)
                Else
                    asys = New ReadCommon(strXML, sSysPara, sIniFile)
                End If

                'Dim asys As New ReadCommon(strXML, sSysPara, sIniFile, stmDug)
                If Not asys.TestXML Then
                    Response.ClearHeaders()
                    Response.Clear()
                    Response.Write(asys.RetunXml)
                    Response.Flush()
                    Response.Close()
                    If Not stmDug Is Nothing Then stmDug.Close()
                    Response.End()
                Else
                    Dim xmlDoc As New XmlDocument
                    xmlDoc.LoadXml(strXML)
                    Try
                        Dim IvrCmd As IVRCommand
                        If blnDebug Then
                            IvrCmd = New IVRCommand(asys.Tel, asys.Floor, asys.CompCode, asys.OwnerName, sIniFile, asys.OraConnstring, stmDug)
                        Else
                            IvrCmd = New IVRCommand(asys.Tel, asys.Floor, asys.CompCode, asys.OwnerName, sIniFile, asys.OraConnstring)
                        End If
                        'Dim IvrCmd As New IVRCommand(asys.Tel, asys.Floor, asys.CompCode, asys.OwnerName, sIniFile, asys.OraConnstring, stmDug)
                        Select Case asys.Cmd
                            Case "N21"
                                If IvrCmd.N21() Then
                                    ReadCommon.WriteXml(xmlDoc, "Status", "1")
                                    If IvrCmd.IVRFileCode <> String.Empty Then
                                        ReadCommon.WriteXml(xmlDoc, "Audio", IvrCmd.IVRFileCode)
                                    End If
                                    XMLToString(xmlDoc)
                                Else
                                    ReadCommon.WriteXml(xmlDoc, "Status", "0")
                                    XMLToString(xmlDoc)
                                End If
                            Case "N22"
                                If IvrCmd.N22 Then
                                    If IvrCmd.IVRFileCode <> String.Empty Then
                                        ReadCommon.WriteXml(xmlDoc, "Audio", IvrCmd.IVRFileCode)
                                    End If
                                    ReadCommon.WriteXml(xmlDoc, "Status", "1")
                                    XMLToString(xmlDoc)
                                Else
                                    ReadCommon.WriteXml(xmlDoc, "Status", "0")
                                    XMLToString(xmlDoc)
                                End If
                            Case "N23"
                                If IvrCmd.N23() Then
                                    If IvrCmd.IVRFileCode <> String.Empty Then
                                        ReadCommon.WriteXml(xmlDoc, "Audio", IvrCmd.IVRFileCode)
                                    End If
                                    ReadCommon.WriteXml(xmlDoc, "Status", "1")
                                    XMLToString(xmlDoc)
                                Else
                                    ReadCommon.WriteXml(xmlDoc, "Status", "0")
                                    XMLToString(xmlDoc)
                                End If
                            Case "N24"
                                If IvrCmd.N24() Then
                                    If IvrCmd.IVRFileCode <> String.Empty Then
                                        ReadCommon.WriteXml(xmlDoc, "Audio", IvrCmd.IVRFileCode)
                                    End If
                                    Using aTrancation As OracleTransaction = IvrCmd.uCnn.BeginTransaction
                                        If IvrCmd.InsertSO006("C", "N24") Then
                                            If IvrCmd.InsertSO552("N24", "C") Then
                                                aTrancation.Commit()
                                            Else
                                                aTrancation.Rollback()
                                            End If
                                        Else
                                            aTrancation.Rollback()
                                        End If
                                    End Using

                                    ReadCommon.WriteXml(xmlDoc, "Status", "1")
                                    XMLToString(xmlDoc)
                                Else
                                    ReadCommon.WriteXml(xmlDoc, "Status", "0")
                                    XMLToString(xmlDoc)
                                End If

                            Case "N25"
                                If IvrCmd.N25 Then
                                    If IvrCmd.IVRFileCode <> String.Empty Then
                                        ReadCommon.WriteXml(xmlDoc, "Audio", IvrCmd.IVRFileCode)
                                    End If
                                    Using aTrancation As OracleTransaction = IvrCmd.uCnn.BeginTransaction
                                        If IvrCmd.InsertSO006("D", "N25") Then
                                            If IvrCmd.InsertSO552("N25", "D") Then
                                                aTrancation.Commit()
                                            Else
                                                aTrancation.Rollback()
                                            End If
                                        Else
                                            aTrancation.Rollback()
                                        End If
                                    End Using
                                    ReadCommon.WriteXml(xmlDoc, "Status", "1")
                                    XMLToString(xmlDoc)
                                Else
                                    ReadCommon.WriteXml(xmlDoc, "Status", "0")
                                    XMLToString(xmlDoc)
                                End If
                            Case "N26"
                                If IvrCmd.N26 Then
                                    If IvrCmd.IVRFileCode <> String.Empty Then
                                        ReadCommon.WriteXml(xmlDoc, "Audio", IvrCmd.IVRFileCode)
                                    End If
                                    Using aTrancation As OracleTransaction = IvrCmd.uCnn.BeginTransaction
                                        If IvrCmd.InsertSO006("I", "N26") Then
                                            If IvrCmd.InsertSO552("N26", "I") Then
                                                aTrancation.Commit()
                                            Else
                                                aTrancation.Rollback()
                                            End If
                                        Else
                                            aTrancation.Rollback()
                                        End If
                                    End Using
                                    ReadCommon.WriteXml(xmlDoc, "Status", "1")
                                    XMLToString(xmlDoc)
                                Else
                                    ReadCommon.WriteXml(xmlDoc, "Status", "0")
                                    XMLToString(xmlDoc)
                                End If
                            Case Else
                                ReadCommon.WriteXml(xmlDoc, "Status", "0")
                                XMLToString(xmlDoc)
                        End Select

                        '填入SO551
                        IvrCmd.InsertSO551(strFromData, strXML, asys.Cmd)
                    Catch ex As Exception
                        If Not stmDug Is Nothing Then
                            stmDug.WriteLine("[IVRCommand]載入失敗")
                        End If

                        ReadCommon.WriteXml(xmlDoc, "Status", "0")
                        XMLToString(xmlDoc)
                    Finally
                        If Not stmDug Is Nothing Then stmDug.Close()
                        Response.ClearHeaders()
                        Response.Clear()
                        Response.Write(strXML)
                        Response.Flush()
                        Response.Close()
                        Response.End()

                    End Try
                End If
            Else
                Response.ClearHeaders()
                Response.Clear()
                Response.Write(strXML)
                Response.Flush()
                Response.Close()
                Response.End()
            End If
        End If
    End Sub
    Private Sub XMLToString(ByVal xmlDoc As XmlDocument)
        Try
            Dim mstm As New MemoryStream()
            xmlDoc.Save(mstm)
            mstm.Flush()
            mstm.Position = 0
            Dim stm As New StreamReader(mstm)
            strXML = stm.ReadToEnd
            stm.Close()
            mstm.Close()
            stm = Nothing
            mstm = Nothing
        Catch ex As Exception

        End Try
    End Sub
    Private Function ReadPara() As Boolean
        Try
            Dim xmlDoc As New XmlDocument
            Dim xmlNodeValue As XmlNode
            strCMD = Request.QueryString("CMD")
            strCompCode = Request.QueryString("Company")
            strLanguage = Request.QueryString("Language")
            strTel = Request.QueryString("Tel")
            strFunc = Request.QueryString("Func")
            strFloor = Request.QueryString("Floor")
            strHistory = Request.QueryString("History")
            strDebug = Request.QueryString("IVRDebug")
            xmlDoc.LoadXml(strxml)
            If strCMD <> String.Empty Then
                xmlNodeValue = xmlDoc.SelectSingleNode("CableSoft/CMD")
                xmlNodeValue.InnerText = strCMD
            End If
            If strCompCode <> String.Empty Then
                xmlNodeValue = xmlDoc.SelectSingleNode("CableSoft/Company")
                xmlNodeValue.InnerText = strCompCode
            End If
            If strLanguage <> String.Empty Then
                xmlNodeValue = xmlDoc.SelectSingleNode("CableSoft/Language")
                xmlNodeValue.InnerText = strLanguage
            End If
            If strTel <> String.Empty Then
                xmlNodeValue = xmlDoc.SelectSingleNode("CableSoft/Tel")
                xmlNodeValue.InnerText = strTel
            End If
            If strFunc <> String.Empty Then
                xmlNodeValue = xmlDoc.SelectSingleNode("CableSoft/Func")
                xmlNodeValue.InnerText = strFunc
            End If
            If strHistory <> String.Empty Then
                xmlNodeValue = xmlDoc.SelectSingleNode("CableSoft/History")
                xmlNodeValue.InnerText = strHistory
            End If
            If strFloor <> String.Empty Then
                xmlNodeValue = xmlDoc.SelectSingleNode("CableSoft/Floor")
                xmlNodeValue.InnerText = strFloor
            End If
            If strDebug <> String.Empty AndAlso strDebug = "6632" Then
                blnDebug = True
            End If
            Dim mstm As New MemoryStream()
            xmlDoc.Save(mstm)
            mstm.Flush()
            mstm.Position = 0
            Dim stm As New StreamReader(mstm)
            strXML = stm.ReadToEnd
            stm.Close()
            mstm.Close()
            stm = Nothing
            mstm = Nothing
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

End Class



