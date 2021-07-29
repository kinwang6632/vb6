Imports System
Imports System.IO
Imports System.Text
Imports System.Data
Imports System.Data.OracleClient
Imports DevExpress
Imports DevExpress.Utils
Imports DevExpress.XtraEditors
Imports DevExpress.XtraEditors.DXErrorProvider
Imports System.Xml
Public Class frmElectron
    'Private strConnectString As String
    'Private cn As OracleClient.OracleConnection
    'Private ds As DataSet
    'Private dap As OracleDataAdapter
    Private Const B08ExeName As String = "CableSoft.Invoice.Win.B08.exe"
    Private FLogPath As String = String.Empty
    Private Const LogPathName As String = "MediaFilePath"
    Private FUseAllComp As Boolean = False
    Private FUseDiscount As Boolean = False
    Private FUseSpaceInv As Boolean = False
    Private fUploadTime As String = String.Empty
    Private invArgument As CableSoft.B07.Parameters.Argument = Nothing
    Dim invIO As CableSoft.B07.SystemIO.Invoice = Nothing
    Private Sub RunExec()
        fUploadTime = "X"
        If invIO Is Nothing Then
            invIO = New CableSoft.B07.SystemIO.Invoice(invArgument)
        End If
        Try
            Dim InvType As CableSoft.B07.InvType.InvTypeEnum.INVTYPE

            invIO.IsAllCompCode = chkAllComp.Checked
            If (chkAll.Checked) AndAlso (Not chkDiscount.Checked) AndAlso (Not chkNotUseInv.Checked) Then
                InvType = CableSoft.B07.InvType.InvTypeEnum.INVTYPE.NormalInvAll
                ' invIO.RunInvType = CableSoft.B07.InvType.InvTypeEnum.INVTYPE.NormalInvAll
            ElseIf (chkVo.Checked) AndAlso (Not chkDiscount.Checked) AndAlso (Not chkNotUseInv.Checked) Then
                InvType = CableSoft.B07.InvType.InvTypeEnum.INVTYPE.NormalInvOV
                'invIO.RunInvType = CableSoft.B07.InvType.InvTypeEnum.INVTYPE.NormalInvOV
            ElseIf (chkDisInv.Checked) AndAlso (chkDiscount.Checked) Then
                InvType = CableSoft.B07.InvType.InvTypeEnum.INVTYPE.OnlyDiscount
                'invIO.RunInvType = CableSoft.B07.InvType.InvTypeEnum.INVTYPE.OnlyDiscount
            ElseIf (chkDisVoInv.Checked) AndAlso (chkDiscount.Checked) Then
                InvType = CableSoft.B07.InvType.InvTypeEnum.INVTYPE.OnlyDiscountOV
            ElseIf chkNotUseInv.Checked Then
                InvType = CableSoft.B07.InvType.InvTypeEnum.INVTYPE.OnlyNotUseInvNum
                'invIO.RunInvType = CableSoft.B07.InvType.InvTypeEnum.INVTYPE.OnlyNotUseInvNum
            End If
            invIO.RunInvType = InvType
            invIO.InvDate1 = dtInv1.Text
            invIO.InvDate2 = dtInv2.Text
            invIO.DiscountDate1 = dtDiscount1.Text
            invIO.DiscountDate2 = dtDiscount2.Text
            invIO.InvNumDate1 = dtInvNum1.Text
            invIO.InvNumDate2 = dtInvNum2.Text
            Using dsInv As DataSet = invIO.GetAllDataToDataSet
                '#5922 判斷是否有負項資料 By Kin 2011/02/16
                '#6371 XML格式不用檢查負項資料 By Kin 2013/03/08
                If Int32.Parse("0" & dsInv.Tables("INV001").Rows(0).Item("UploadType")) = 0 Then
                    If Not chkNegativeData(dsInv.Tables("INV007"), invIO) Then
                        Exit Sub
                    End If
                End If
                '#6371 判斷發票是否為捐贈，如為捐贈則捐贈單位與捐贈單位統編要有值 By Kin 2013/02/19
                If Int32.Parse("0" & dsInv.Tables("INV001").Rows(0).Item("UploadType")) = 1 Then
                    If (dsInv.Tables("GiveErrData") IsNot Nothing) AndAlso
                             (dsInv.Tables("GiveErrData").Rows.Count > 0) Then
                        ShowGiveErrData(dsInv.Tables("GiveErrData"))
                        Exit Sub
                    End If
                End If
                Using DataToFile As New CableSoft.B07.InvDataToFile.ToInvFile(invArgument, dsInv)
                    If DataToFile.WriteINVFile(InvType) Then
                        If DataToFile.TotalCount = 0 Then
                            Me.BeginInvoke(New Action(Of String)(AddressOf ShowMsgAndWriteMsg), "無任何資料")
                        Else
                            fUploadTime = invIO.UpdateINV '  objData.UpdateINV()
                            Dim aMsg As String = String.Empty
                            Dim aA0401Name As String = "A0401"
                            Dim aA0501Name As String = "A0501"
                            Dim aB0301Name As String = "B0301"
                            Dim aB0401Name As String = "B0401"
                            Dim aE0402Name As String = "E0402"
                            If Int32.Parse(dsInv.Tables("Inv001").Rows(0).Item("UploadType")) = 1 Then
                                aA0401Name = "C0401"
                                aA0501Name = "C0501"
                                aB0301Name = "D0401"
                                aB0401Name = "D0501"
                            End If
                            Select Case InvType
                                Case CableSoft.B07.InvType.InvTypeEnum.INVTYPE.NormalInvAll
                                    aMsg = String.Format("產生完成" & ControlChars.CrLf & _
                                         "{0} (一般發票) : {1} 筆" & ControlChars.CrLf & _
                                         "{2} (作廢發票) : {3} 筆", aA0401Name, _
                                         DataToFile.A0401RcdCount, _
                                         aA0501Name, _
                                         DataToFile.A0501RcdCount)
                                Case CableSoft.B07.InvType.InvTypeEnum.INVTYPE.NormalInvOV
                                    aMsg = String.Format("產生完成" & Environment.NewLine & _
                                         "{0} (作廢發票) : {1} 筆",
                                         aA0501Name, _
                                         DataToFile.A0501RcdCount)
                                Case CableSoft.B07.InvType.InvTypeEnum.INVTYPE.OnlyDiscount
                                    aMsg = String.Format(aMsg & ControlChars.CrLf & _
                                        "{0} (折讓發票) : {1} 筆", aB0301Name, DataToFile.B0301RcdCount)
                                Case CableSoft.B07.InvType.InvTypeEnum.INVTYPE.OnlyDiscountOV
                                    aMsg = String.Format(aMsg & ControlChars.CrLf & _
                                      "{0} (折讓作廢發票) : {1} 筆", aB0401Name, DataToFile.B0401RcdCount)
                                Case CableSoft.B07.InvType.InvTypeEnum.INVTYPE.OnlyNotUseInvNum
                                    aMsg = String.Format(aMsg & ControlChars.CrLf & _
                                                   "{0} (未空白使用字軌) : {1} 筆", aE0402Name,
                                                   DataToFile.E0402RcdCount)
                            End Select

                            'If FUseDiscount Then
                            '    '#6001 增加折讓作廢發票 By Kin 2011/07/14
                            '    If chkDisInv.Checked Then
                            '        aMsg = String.Format(aMsg & ControlChars.CrLf & _
                            '            "{0} (折讓發票) : {1} 筆", aB0301Name, DataToFile.B0301RcdCount)
                            '    Else
                            '        aMsg = String.Format(aMsg & ControlChars.CrLf & _
                            '            "{0} (折讓作廢發票) : {1} 筆", aB0401Name, DataToFile.B0401RcdCount)
                            '    End If
                            'End If
                            'If FUseSpaceInv Then
                            '    aMsg = String.Format(aMsg & ControlChars.CrLf & _
                            '                       "{0} (未空白使用字軌) : {1} 筆", aE0402Name,
                            '                       DataToFile.E0402RcdCount)
                            'End If
                            Me.BeginInvoke(New Action(Of String)(AddressOf ShowMsgAndWriteMsg), aMsg)

                            '#5966 原本以參數為主，改為以使用者自行勾選為主 By Kin 2011/03/29
                            '#6600 測試不OK,如果有勾選空白未使用,不能啟動B08 By Kin 2013/11/11
                            If chkAutoNotify.Checked Then
                                If (chkAll.Checked) AndAlso (Not Convert.ToBoolean(chkDiscount.EditValue)) AndAlso
                                    (Not chkNotUseInv.Checked) Then
                                    StartB08EXE()
                                End If

                            End If
                            'XtraMessageBox.Show(aMsg, "訊息", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        End If
                    End If
                End Using
            End Using

        Catch ex As Exception
        Finally
            'If invArgument IsNot Nothing Then
            '    invArgument.Dispose()
            'End If
            'If invIO IsNot Nothing Then
            '    invIO.Dispose()
            'End If
            'invIO = Nothing
        End Try
    End Sub
    Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click

        fUploadTime = "X"
        If Not IsDataOK2() Then
            Exit Sub
        End If
        RunExec()
        Exit Sub
        File.WriteAllLines(Application.StartupPath & "\B07Para.Set", Environment.GetCommandLineArgs.Skip(1).ToArray)
        FUseAllComp = Convert.ToBoolean(chkAllComp.EditValue)
        FUseDiscount = Convert.ToBoolean(chkDiscount.EditValue)
        FUseSpaceInv = Convert.ToBoolean(chkNotUseInv.EditValue)
        Dim tbl007 As DataTable = Nothing
        Dim tbl001 As DataTable = Nothing
        Dim tbl014 As DataTable = Nothing
        Dim tbl003 As DataTable = Nothing
        Dim tbl099 As DataTable = Nothing
        Dim objData As clsData = Nothing
        Dim dsInv As New DataSet()
        Try
            Me.Cursor = Cursors.WaitCursor

            Dim objIO As clsDataIO = Nothing
            Dim aInvPath As String = String.Empty

            If chkAll.EditValue Then
                objData = New clsData(clsCommon.GetConnect, INVTYPE.All, FUseAllComp)
            Else
                objData = New clsData(clsCommon.GetConnect, INVTYPE.OV, FUseAllComp)
            End If
            If Not String.IsNullOrEmpty(dtInv1.Text) Then
                objData.InvDate1 = Format(dtInv1.EditValue, "yyyyMMdd")
            End If
            If Not String.IsNullOrEmpty(dtInv2.Text) Then
                objData.InvDate2 = Format(dtInv2.EditValue, "yyyyMMdd")
            End If

            '#5966 增加發票號碼條件 By Kin 2011/03/30
            objData.InvId1 = edtInvId1.Text
            objData.InvId2 = edtInvId2.Text
            objData.StartAllUpload = clsCommon.GetStartAllUpload
            If Convert.ToBoolean(chkDiscount.EditValue) Then
                objData.DiscountDate1 = Format(dtDiscount1.EditValue, "yyyyMMdd")
                objData.DiscountDate2 = Format(dtDiscount2.EditValue, "yyyyMMdd")
                objData.InvDate1 = "00010101"
                objData.InvDate2 = "00010101"
                objData.DiscountOv = chkDisVoInv.Checked
            End If
            '#6600 增加E0402 空白未使用的發票 By Kin 2013/10/03
            If grpNotUseInv.Enabled Then
                objData.InvNumDate1 = Date.Parse(dtInvNum1.EditValue).ToString("yyyyMM")
                objData.InvNumDate2 = Date.Parse(dtInvNum2.EditValue).ToString("yyyyMM")
            End If
            tbl007 = objData.GetINV007()
            tbl001 = objData.GetINV001()
            tbl014 = objData.GetINV014()
            tbl099 = objData.GetINV099()
            '#6262 增加上傳檔案路徑參數 By Kin 2012/07/11
            tbl003 = objData.GetInv003()
            '#5922 判斷是否有負項資料 By Kin 2011/02/16
            '#6371 XML格式不用檢查負項資料 By Kin 2013/03/08
            If Int32.Parse("0" & tbl001.Rows(0).Item("UploadType")) = 0 Then
                If Not chkNegativeData(tbl007, objData) Then
                    Exit Sub
                End If
            End If
            '#6371 判斷發票是否為捐贈，如為捐贈則捐贈單位與捐贈單位統編要有值 By Kin 2013/02/19
            If Int32.Parse("0" & tbl001.Rows(0).Item("UploadType")) = 1 Then
                If (objData.dtGiveErrData IsNot Nothing) AndAlso
                         (objData.dtGiveErrData.Rows.Count > 0) Then
                    ShowGiveErrData(objData.dtGiveErrData)
                    Exit Sub
                End If
            End If
            dsInv.Tables.Add(tbl007)
            dsInv.Tables.Add(tbl001)
            dsInv.Tables.Add(tbl014)
            dsInv.Tables.Add(tbl003)
            dsInv.Tables.Add(tbl099)
            objIO = New clsDataIO(aInvPath, dsInv)

            'objIO.tbInv003 = tbl003
            '#5922 增加備份路徑 By Kin 2011/02/14
            If chkAll.EditValue Then
                objIO.WriteINVFile(INVTYPE.All, FUseDiscount, chkDisVoInv.Checked, chkNotUseInv.Checked, _
                                   FLogPath, Int32.Parse("0" & tbl001.Rows(0).Item("UploadType")))
            Else
                objIO.WriteINVFile(INVTYPE.OV, FUseDiscount, chkDisVoInv.Checked, chkNotUseInv.Checked, _
                                    FLogPath, Int32.Parse("0" & tbl001.Rows(0).Item("UploadType")))
            End If
            If objIO.TotalCount = 0 Then
                Me.BeginInvoke(New Action(Of String)(AddressOf ShowMsgAndWriteMsg), "無任何資料")
            Else
                fUploadTime = objData.UpdateINV()
                Dim aMsg As String = String.Empty
                Dim aA0401Name As String = "A0401"
                Dim aA0501Name As String = "A0501"
                Dim aB0301Name As String = "B0301"
                Dim aB0401Name As String = "B0401"
                Dim aE0402Name As String = "E0402"
                If Int32.Parse(dsInv.Tables("Inv001").Rows(0).Item("UploadType")) = 1 Then
                    aA0401Name = "C0401"
                    aA0501Name = "C0501"
                    aB0301Name = "D0401"
                    aB0401Name = "D0501"
                End If
                aMsg = String.Format("產生完成" & ControlChars.CrLf & _
                              "{0} (一般發票) : {1} 筆" & ControlChars.CrLf & _
                              "{2} (作廢發票) : {3} 筆", aA0401Name, _
                              objIO.A0401RcdCount, _
                              aA0501Name, _
                              objIO.A0501RcdCount)
                If FUseDiscount Then
                    '#6001 增加折讓作廢發票 By Kin 2011/07/14
                    If chkDisInv.Checked Then
                        aMsg = String.Format(aMsg & ControlChars.CrLf & _
                            "{0} (折讓發票) : {1} 筆", aB0301Name, objIO.B0301RcdCount)
                    Else
                        aMsg = String.Format(aMsg & ControlChars.CrLf & _
                            "{0} (折讓作廢發票) : {1} 筆", aB0401Name, objIO.B0401RcdCount)
                    End If
                End If
                If FUseSpaceInv Then
                    aMsg = String.Format(aMsg & ControlChars.CrLf & _
                                       "{0} (未空白使用字軌) : {1} 筆", aE0402Name,
                                       objIO.E0402RcdCount)
                End If
                Me.BeginInvoke(New Action(Of String)(AddressOf ShowMsgAndWriteMsg), aMsg)

                '#5966 原本以參數為主，改為以使用者自行勾選為主 By Kin 2011/03/29
                '#6600 測試不OK,如果有勾選空白未使用,不能啟動B08 By Kin 2013/11/11
                If chkAutoNotify.Checked Then
                    If (chkAll.Checked) AndAlso (Not Convert.ToBoolean(chkDiscount.EditValue)) AndAlso
                        (Not chkNotUseInv.Checked) Then
                        StartB08EXE()
                    End If

                End If
                'XtraMessageBox.Show(aMsg, "訊息", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If

        Catch ex As Exception
            XtraMessageBox.Show(ex.ToString, "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
            clsCommon.WriteErr(ex.ToString)
        Finally
            If tbl007 IsNot Nothing Then
                tbl007.Dispose()
            End If
            If tbl001 IsNot Nothing Then
                tbl001.Dispose()
            End If
            If tbl014 IsNot Nothing Then
                tbl014.Dispose()
            End If
            If tbl003 IsNot Nothing Then
                tbl003.Dispose()
            End If
            If dsInv IsNot Nothing Then
                dsInv.Dispose()
            End If
            If objData IsNot Nothing Then
                objData.Dispose()
            End If
            Me.Cursor = Cursors.Default
        End Try
    End Sub
    Private Function ShowGiveErrData(ByVal dt As DataTable) As Boolean
        Try
            Dim aMsg As New System.Text.StringBuilder()
            Dim aFile As String = "電子發票上傳異常資料記錄_" & _
                        Format(Now, "yyyyMMdd_HHmmss") & ".Txt"
            For Each rw As DataRow In dt.Rows
                aMsg.AppendLine(rw.Item("Msg").ToString)
            Next
            CableSoft.B07.Parameters.Argument.WriteLog(FLogPath & aFile, aMsg.ToString)            
            'clsCommon.WriteLog(FLogPath & aFile, aMsg.ToString)
            Shell("notepad " & FLogPath & aFile, AppWinStyle.MaximizedFocus)
        Catch ex As Exception
            XtraMessageBox.Show(ex.ToString)
        End Try
        
    End Function
    Private Overloads Function chkNegativeData(ByVal dt As DataTable, ByVal objInvIO As CableSoft.B07.SystemIO.Invoice) As Boolean
        Try

            Dim aMsg As String = String.Empty
            Dim aChgMsg As String = String.Empty
            Dim aFile As String = "電子發票上傳異常資料記錄_" & _
                    Format(Now, "yyyyMMdd_HHmmss") & ".Txt"

            '過濾重複資料
            'Dim aChkInvid = (From aInvids In dt _
            '              Order By aInvids.Item("INVID") Ascending _
            '              Where Int32.Parse("0" & aInvids.Item("Totalamount".ToUpper).ToString) <= 0 _
            '              Select aInvids.Item("INVID")).Distinct
            Dim aChkInvid = From aInvids In dt _
                          Order By aInvids.Item("INVID") Ascending _
                          Where Int32.Parse(aInvids.Item("Totalamount".ToUpper).ToString) < 0 _
                          Select aInvids.Item("INVID")
            Dim aINVLst As New List(Of String)
            For Each aInvid In aChkInvid
                aMsg = aMsg & Environment.NewLine & aInvid & " 中有負項收費資料"
                If Not aINVLst.Contains(aInvid) Then
                    aChgMsg = aChgMsg & Environment.NewLine & aInvid
                    aINVLst.Add(aInvid)
                End If
            Next
            If Not String.IsNullOrEmpty(aMsg) Then
                'XtraMessageBox.Show("發票資料含有負項資料！" & vbCrLf & "請修改資料再做上傳動作", "訊息", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                If XtraMessageBox.Show("發票資料含有負項資料！" & vbCrLf & "是否要轉成 0.電子計算機發票 ？ ", _
                                       "詢問", MessageBoxButtons.YesNo, _
                                       MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then

                    If objInvIO.TranInvType(aINVLst) Then
                        XtraMessageBox.Show(String.Format("  已調整 {0} 筆資料 ", aINVLst.Count.ToString) & _
                                             Environment.NewLine & " 請重新產生上傳檔 ！", "訊息", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        aMsg = "已調整下列發票種類" & Environment.NewLine & New String("-"c, 40) & aChgMsg
                        'clsCommon.WriteLog(FLogPath & aFile, aMsg)
                        CableSoft.B07.Parameters.Argument.WriteLog(FLogPath & aFile, aMsg)
                        Shell("notepad " & FLogPath & aFile, AppWinStyle.MaximizedFocus)
                    End If
                Else
                    aMsg = GetScreenWhere() & Environment.NewLine & "異常資料 : " & aMsg
                    'clsCommon.WriteLog(FLogPath & aFile, aMsg)
                    CableSoft.B07.Parameters.Argument.WriteLog(FLogPath & aFile, aMsg)
                    Shell("notepad " & FLogPath & aFile, AppWinStyle.MaximizedFocus)
                End If
                Return False
            End If
            Return True
        Catch ex As Exception
            XtraMessageBox.Show(ex.ToString)
        End Try
    End Function
    '檢查是否有負項資料
    Private Overloads Function chkNegativeData(ByVal dt As DataTable, ByVal objUpd As clsData) As Boolean
        Try

            Dim aMsg As String = String.Empty
            Dim aChgMsg As String = String.Empty
            Dim aFile As String = "電子發票上傳異常資料記錄_" & _
                    Format(Now, "yyyyMMdd_HHmmss") & ".Txt"

            '過濾重複資料
            'Dim aChkInvid = (From aInvids In dt _
            '              Order By aInvids.Item("INVID") Ascending _
            '              Where Int32.Parse("0" & aInvids.Item("Totalamount".ToUpper).ToString) <= 0 _
            '              Select aInvids.Item("INVID")).Distinct
            Dim aChkInvid = From aInvids In dt _
                          Order By aInvids.Item("INVID") Ascending _
                          Where Int32.Parse(aInvids.Item("Totalamount".ToUpper).ToString) < 0 _
                          Select aInvids.Item("INVID")
            Dim aINVLst As New List(Of String)
            For Each aInvid In aChkInvid
                aMsg = aMsg & Environment.NewLine & aInvid & " 中有負項收費資料"
                If Not aINVLst.Contains(aInvid) Then
                    aChgMsg = aChgMsg & Environment.NewLine & aInvid
                    aINVLst.Add(aInvid)
                End If
            Next
            If Not String.IsNullOrEmpty(aMsg) Then
                'XtraMessageBox.Show("發票資料含有負項資料！" & vbCrLf & "請修改資料再做上傳動作", "訊息", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                If XtraMessageBox.Show("發票資料含有負項資料！" & vbCrLf & "是否要轉成 0.電子計算機發票 ？ ", _
                                       "詢問", MessageBoxButtons.YesNo, _
                                       MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
                    If objUpd.TranInvType(aINVLst) Then
                        XtraMessageBox.Show(String.Format("  已調整 {0} 筆資料 ", aINVLst.Count.ToString) & _
                                             Environment.NewLine & " 請重新產生上傳檔 ！", "訊息", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        aMsg = "已調整下列發票種類" & Environment.NewLine & New String("-"c, 40) & aChgMsg
                        'clsCommon.WriteLog(FLogPath & aFile, aMsg)
                        CableSoft.B07.Parameters.Argument.WriteLog(FLogPath & aFile, aMsg)
                        Shell("notepad " & FLogPath & aFile, AppWinStyle.MaximizedFocus)
                    End If
                Else
                    aMsg = GetScreenWhere() & Environment.NewLine & "異常資料 : " & aMsg
                    'clsCommon.WriteLog(FLogPath & aFile, aMsg)
                    CableSoft.B07.Parameters.Argument.WriteLog(FLogPath & aFile, aMsg)
                    Shell("notepad " & FLogPath & aFile, AppWinStyle.MaximizedFocus)
                End If
                Return False
            End If
            Return True
        Catch ex As Exception
            XtraMessageBox.Show(ex.ToString)
        End Try
    End Function
    Private Sub StartB08EXE()
        Try
            'Dim strArgument As String = clsCommon.GetB08Arguments
            Dim strArgument As String = invArgument.GetB08Arguments
            Dim aArgInv1, aArgInv2 As String
            aArgInv1 = "0"
            aArgInv2 = "0"
            If Not String.IsNullOrEmpty(edtInvId1.Text) Then
                aArgInv1 = edtInvId1.Text
            End If
            If Not String.IsNullOrEmpty(edtInvId2.Text) Then
                aArgInv2 = edtInvId2.Text
            End If
            strArgument = strArgument & " 1 " & dtInv1.Text & " " & dtInv2.Text
            If Convert.ToBoolean(chkSingleComp.EditValue) Then
                strArgument = strArgument & " 0 "
            Else
                strArgument = strArgument & " 1 "
            End If
            strArgument = strArgument & aArgInv1 & " " & aArgInv2
            '#6080 測試不OK 增加 DBLink
            'strArgument = strArgument & " " & clsCommon.GetDbLink
            strArgument = strArgument & " " & invArgument.GetDbLink
            '#6396 增加UploadTime參數 By Kin 2012/12/18
            strArgument = strArgument & " " & fUploadTime
            '#增加MisDbSid、MisDbUserId、MisDbPassword、MisDbOwner For Miggie By Kin 2013/01/28
            'strArgument = strArgument & " " & clsCommon.GetMisDbSid & " " & clsCommon.GetMisDbUserId & " " & _
            '                clsCommon.GetMisDbPassword & " " & clsCommon.GetMisDbOwner
            strArgument = strArgument & " " & invArgument.GetMisDbSid & " " & invArgument.GetMisDbUserId & " " & _
                            invArgument.GetMisDbPassword & " " & invArgument.GetMisDbOwner


            'Process.Start(clsCommon.GetAppPath & B08ExeName, strArgument)
            Process.Start(CableSoft.B07.Parameters.Argument.GetAppPath & B08ExeName, strArgument)

        Catch ex As Exception
            XtraMessageBox.Show(ex.ToString, "訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
            'clsCommon.WriteErr(ex.ToString)
            CableSoft.B07.Parameters.Argument.WriteErr(ex.ToString)
        End Try

    End Sub
    Private Sub ShowMsgAndWriteMsg(ByVal aMsg As String)
        Try
            Dim aWriteMsg As String = GetScreenWhere()
            'aWriteMsg = "發票起迄日 : " & dtInv1.Text & " - " & _
            'dtInv2.Text & Environment.NewLine & "作廢註記 : " & IIf(chkAll.Checked, chkAll.Text, chkVo.Text) & _
            'Environment.NewLine & "是否勾選折讓註記 : " & IIf(Convert.ToBoolean(chkDiscount.EditValue), "是", "否") & _
            'Environment.NewLine & "折讓單據起訖日 : " & dtDiscount1.Text & " - " & dtDiscount2.Text & _
            'Environment.NewLine & "公司別選項 : " & IIf(Convert.ToBoolean(chkSingleComp.EditValue), chkSingleComp.Text, chkAllComp.Text)
            aWriteMsg = aMsg & Environment.NewLine & aWriteMsg
            'clsCommon.WriteProdcutLog(FLogPath, aWriteMsg)
            CableSoft.B07.Parameters.Argument.WriteProdcutLog(FLogPath, aWriteMsg)
            XtraMessageBox.Show(aMsg, "訊息", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            'clsCommon.WriteErr(ex.ToString)
            CableSoft.B07.Parameters.Argument.WriteErr(ex.ToString)
        End Try
    End Sub
    Private Function GetScreenWhere() As String
        Try
            Dim aRetString As String = String.Empty
            aRetString = "發票起迄日 : " & dtInv1.Text & " - " & _
            dtInv2.Text & Environment.NewLine & "作廢註記 : " & IIf(chkAll.Checked, chkAll.Text, chkVo.Text) & _
            Environment.NewLine & "是否勾選折讓註記 : " & IIf(Convert.ToBoolean(chkDiscount.EditValue), "是", "否") & _
            Environment.NewLine & "折讓單據起訖日 : " & dtDiscount1.Text & " - " & dtDiscount2.Text & _
            Environment.NewLine & "折讓作廢註記：" & IIf(chkDiscount.Checked, "是", "否") & _
            Environment.NewLine & "公司別選項 : " & IIf(Convert.ToBoolean(chkSingleComp.EditValue), chkSingleComp.Text, chkAllComp.Text) & _
            Environment.NewLine & "是否勾選空白未使用字軌檔：" & IIf(chkNotUseInv.Checked, "是", "否") & _
            Environment.NewLine & "空白未使用發票起迄年月：" & dtInvNum1.Text & " - " & dtInvNum2.Text
            Return aRetString
        Catch ex As Exception
            'clsCommon.WriteErr(ex.ToString)
            CableSoft.B07.Parameters.Argument.WriteErr(ex.ToString)
            Return String.Empty
        End Try
    End Function
    Private Function IsDataOK2() As Boolean
        On Error Resume Next
        If Not Convert.ToBoolean(chkDiscount.EditValue) Then
            If (String.IsNullOrEmpty(dtInvNum1.Text)) AndAlso (String.IsNullOrEmpty(dtInvNum2.Text)) Then
                If String.IsNullOrEmpty(dtInv1.Text) Then
                    XtraMessageBox.Show("請輸入發票起始日期", "訊息", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    dtInv1.Focus()
                    Return False
                End If
                If String.IsNullOrEmpty(dtInv2.Text) Then
                    XtraMessageBox.Show("請輸入發票終止日期", "訊息", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    dtInv2.Focus()
                    Return False
                End If
            End If
        End If
        If Convert.ToBoolean(chkDiscount.EditValue) Then
            If String.IsNullOrEmpty(dtDiscount1.Text) Then
                XtraMessageBox.Show("請輸入折讓單據起始日期", "訊息", MessageBoxButtons.OK, MessageBoxIcon.Information)
                dtDiscount1.Focus()
                Return False
            End If
        End If
        If Convert.ToBoolean(chkDiscount.EditValue) Then
            If String.IsNullOrEmpty(dtDiscount2.Text) Then
                XtraMessageBox.Show("請輸入折讓單據終止日期", "訊息", MessageBoxButtons.OK, MessageBoxIcon.Information)
                dtDiscount2.Focus()
                Return False
            End If
        End If
        If chkNotUseInv.Checked Then
            If String.IsNullOrEmpty(dtInvNum1.Text) Then
                XtraMessageBox.Show("請輸入未使用發票起始日期", "訊息", MessageBoxButtons.OK, MessageBoxIcon.Information)
                dtInvNum1.Focus()
                Return False
            End If
            If String.IsNullOrEmpty(dtInvNum2.Text) Then
                XtraMessageBox.Show("請輸入未使用發票終止日期", "訊息", MessageBoxButtons.OK, MessageBoxIcon.Information)
                dtInvNum2.Focus()
                Return False
            End If
        End If
        Return True
    End Function
    Private Function IsDataOK() As Boolean
        On Error Resume Next
        Dim notEmptyValidationRule As New ConditionValidationRule()
        Dim Ret As Boolean = False
        notEmptyValidationRule.ConditionOperator = ConditionOperator.IsNotBlank
        If Not Convert.ToBoolean(chkDiscount.EditValue) Then
            notEmptyValidationRule.ErrorText = "請輸發票起日"
            DxValidationProvider1.SetValidationRule(dtInv1, notEmptyValidationRule)
            Ret = DxValidationProvider1.Validate()
            If Not Ret Then
                dtInv1.Focus()
                Exit Function
            End If
            notEmptyValidationRule.ErrorText = "請輸發票迄日"
            DxValidationProvider1.SetValidationRule(dtInv2, notEmptyValidationRule)
            Ret = DxValidationProvider1.Validate()
            If Not Ret Then
                dtInv2.Focus()
                Exit Function
            End If
        End If        
        Ret = True
        If Convert.ToBoolean(chkDiscount.EditValue) Then
            notEmptyValidationRule.ErrorText = "請輸入折讓單據起日"
            DxValidationProvider1.SetValidationRule(dtDiscount1, notEmptyValidationRule)
            Ret = DxValidationProvider1.Validate()
            DxValidationProvider1.RemoveControlError(dtInv1)
            If Not Ret Then
                dtDiscount1.Focus()
                Exit Function
            End If
            notEmptyValidationRule.ErrorText = "請輸入折讓單據迄日"
            DxValidationProvider1.SetValidationRule(dtDiscount2, notEmptyValidationRule)
            Ret = DxValidationProvider1.Validate()
            If Not Ret Then
                dtDiscount2.Focus()
                Exit Function
            End If
        End If
        Return Ret
    End Function

    Private Sub dtInv1_EditValueChanging(ByVal sender As Object, ByVal e As DevExpress.XtraEditors.Controls.ChangingEventArgs) Handles dtInv1.EditValueChanging, dtDiscount2.EditValueChanging
        On Error Resume Next
        DxValidationProvider1.RemoveControlError(sender)
    End Sub

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Application.Exit()
    End Sub

    Private Sub frmElectron_Disposed(sender As Object, e As System.EventArgs) Handles Me.Disposed
        Try
            If invArgument IsNot Nothing Then
                invArgument.Dispose()
            End If
            If invIO IsNot Nothing Then
                invIO.Dispose()
            End If
            invArgument = Nothing
            invIO = Nothing
        Catch ex As Exception

        End Try
    End Sub

    Private Sub frmElectron_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.F2 Then
            cmdOK.PerformClick()
        End If
        If e.KeyCode = Keys.Escape Then
            cmdExit.PerformClick()
        End If
        If e.Control AndAlso e.Alt AndAlso e.KeyCode = Keys.H Then
            Dim strArg() As String = Environment.GetCommandLineArgs
            Dim bduArg As New System.Text.StringBuilder()
            Dim cmdArg As String = String.Empty
            For i As Int32 = 0 To strArg.Length - 1
                If i <> 0 Then
                    If i = 1 Then
                        cmdArg = strArg(i)
                    Else
                        cmdArg = cmdArg & " " & strArg(i)
                    End If
                End If
            Next
            XtraMessageBox.Show(cmdArg, "傳入參數", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

    Private Sub frmElectron_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        'Dim aFilePath As String = clsData.GetINVPathName(clsCommon.GetConnect, clsCommon.GetCompCode, "EInvoiceFilePath")
        If invArgument Is Nothing Then
            invArgument = New CableSoft.B07.Parameters.Argument(Environment.GetCommandLineArgs.Skip(1).ToArray)
        End If
        'If String.IsNullOrEmpty(aFilePath) Then
        '    Exit Sub
        'End If
        'If Not aFilePath.EndsWith("\") Then
        '    aFilePath = aFilePath & "\"
        'End If
        'txtInvPath.Text = aFilePath
        If invArgument.GetStarNotify Then
            chkAutoNotify.Checked = True
        End If
        FLogPath = CableSoft.B07.SystemIO.Invoice.GetINVPathName(invArgument.GetConnect,
                                                      invArgument.GetCompCode,
                                                      LogPathName)


        'FLogPath = clsData.GetINVPathName(clsCommon.GetConnect, clsCommon.GetCompCode, "MediaFilePath")
        FLogPath = Path.GetFullPath(FLogPath)
        If Not FLogPath.EndsWith("\") Then
            FLogPath = FLogPath & "\"
        End If
    End Sub

    Private Sub dtInv2_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtInv2.Enter
        On Error Resume Next
        If String.IsNullOrEmpty(dtInv2.Text) OrElse (dtInv2.DateTime < dtInv1.DateTime) Then
            If Not String.IsNullOrEmpty(dtInv1.Text) Then
                dtInv2.Text = dtInv1.Text
            End If
        End If
    End Sub

    Public Sub New()

        ' 此為 Windows Form 設計工具所需的呼叫。
        DevExpress.Skins.SkinManager.EnableFormSkins()
        Dim objZhch As New Editzhcn()
        DevExpress.XtraEditors.Controls.Localizer.Active = objZhch
        InitializeComponent()

        ' 在 InitializeComponent() 呼叫之後加入任何初始設定。

    End Sub



    Private Sub chkDiscount_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkDiscount.CheckedChanged

        grpDiscount.Enabled = Convert.ToBoolean(chkDiscount.EditValue)
        If grpDiscount.Enabled Then
            dtInv1.Text = String.Empty
            dtInv2.Text = String.Empty
            dtInv1.Enabled = False
            dtInv2.Enabled = False
            chkNotUseInv.Checked = False
            dtInvNum1.Text = String.Empty
            dtInvNum2.Text = String.Empty
        Else
            dtDiscount1.Text = String.Empty
            dtDiscount2.Text = String.Empty
            dtInv1.Enabled = True
            dtInv2.Enabled = True
        End If
        edtInvId1.Enabled = dtInv1.Enabled
        edtInvId2.Enabled = dtInv2.Enabled
        grpNotUseInv.Enabled = Not grpDiscount.Enabled
    End Sub

    Private Sub edtInvId1_InvalidValue(ByVal sender As System.Object, ByVal e As DevExpress.XtraEditors.Controls.InvalidValueExceptionEventArgs) Handles edtInvId1.InvalidValue, edtInvId2.InvalidValue
        e.ErrorText = "發票號碼不正確"
    End Sub

    Private Sub chkNotUseInv_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkNotUseInv.CheckedChanged
        grpNotUseInv.Enabled = Convert.ToBoolean(chkNotUseInv.EditValue)
        If grpNotUseInv.Enabled Then           
            dtInvNum1.Enabled = True
            dtInvNum2.Enabled = True
            dtInv1.Text = String.Empty
            dtInv2.Text = String.Empty
            chkDiscount.Checked = False
            dtDiscount1.Text = String.Empty
            dtDiscount2.Text = String.Empty
            grpDiscount.Enabled = False
        Else
            dtInvNum1.EditValue = Nothing
            dtInvNum2.EditValue = Nothing
            dtInvNum1.Enabled = False
            dtInvNum1.Enabled = False
        End If
        grpDiscount.Enabled = Not grpNotUseInv.Enabled
        'dtInv1.Enabled = (Not grpNotUseInv.Enabled) AndAlso (Not grpDiscount.Enabled)
        'dtInv2.Enabled = (Not grpNotUseInv.Enabled) AndAlso (Not grpDiscount.Enabled)
        'edtInvId1.Enabled = (Not grpNotUseInv.Enabled) AndAlso (Not grpDiscount.Enabled)
        'edtInvId2.Enabled = (Not grpNotUseInv.Enabled) AndAlso (Not grpDiscount.Enabled)
        dtInv1.Enabled = (Not chkNotUseInv.Checked) AndAlso (Not chkDiscount.Checked)
        dtInv2.Enabled = (Not chkNotUseInv.Checked) AndAlso (Not chkDiscount.Checked)
        edtInvId1.Enabled = (Not chkNotUseInv.Checked) AndAlso (Not chkDiscount.Checked)
        edtInvId2.Enabled = (Not chkNotUseInv.Checked) AndAlso (Not chkDiscount.Checked)
    End Sub

    Private Sub dtInvNum1_Enter(sender As System.Object, e As System.EventArgs) Handles dtInvNum1.Enter

    End Sub
End Class
Public Class Editzhcn
    Inherits DevExpress.XtraEditors.Controls.Localizer
    Public Overrides Function GetLocalizedString(ByVal id As DevExpress.XtraEditors.Controls.StringId) As String
        'Return MyBase.GetLocalizedString(id)
        Select Case id
            Case XtraEditors.Controls.StringId.DateEditClear
                Return "清除"
            Case XtraEditors.Controls.StringId.DateEditToday
                Return "今天"
            Case Else
                Return MyBase.GetLocalizedString(id)
        End Select
    End Function

End Class
