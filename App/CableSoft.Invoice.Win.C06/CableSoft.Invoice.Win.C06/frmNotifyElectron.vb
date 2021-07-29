Imports DevExpress.XtraExport
Imports DevExpress.XtraReports.Design
Imports DevExpress.XtraReports.UI
Imports DevExpress.XtraEditors
Imports System
Imports System.IO
Imports System.Text
Imports System.Data
Imports System.Data.OracleClient

Public Class frmNotifyElectron
    Private Const fRptName_Detail As String = "XtraRptInvC06_1.repx"
    Private Const fRptName_Total As String = "XtraRptInvC06_2.repx"
    Private fRptPath As String = String.Empty
    Private fWhereString As String
    Private fDBName As String = String.Empty
    Private fDBUserId As String = String.Empty
    Private fDBPassword As String = String.Empty
    Private fCompCode As String = String.Empty
    Private fLoginName As String = String.Empty
    Private fSQLText As String = String.Empty


    Private Sub frmNotifyElectron_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        Static aDeveloper As String = String.Empty
        If e.KeyCode = Keys.Escape Then
            btnExit.PerformClick()
        End If
        If e.KeyCode = Keys.F2 Then
            btnOK.PerformClick()
        End If
        If e.Control Then
            If e.KeyCode <> Keys.ControlKey Then
                aDeveloper = aDeveloper & Convert.ToChar(e.KeyValue)
            Else
                aDeveloper = String.Empty
            End If

            If aDeveloper.ToUpper = "DEVELOPER" Then
                btnDesign.Visible = Not btnDesign.Visible
                btnShowSQL.Visible = Not btnShowSQL.Visible
            End If
        Else
            aDeveloper = String.Empty
        End If
        If (e.Control) AndAlso (e.KeyCode = Keys.H) Then
            QrySQL()
        End If

    End Sub


    Private Sub frmNotifyElectron_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            'fRptPath = Application.StartupPath
            fRptPath = Environment.GetCommandLineArgs(6)
            If Not fRptPath.EndsWith("\") Then
                fRptPath = fRptPath & "\"
            End If

            If Not File.Exists(fRptPath & fRptName_Detail) Then
                InfoMsg(fRptPath & fRptName_Detail & Environment.NewLine & "檔案不存在", MessageBoxIcon.Error)
                Application.Exit()
                Exit Sub
            End If
            If Not File.Exists(fRptPath & fRptName_Total) Then
                InfoMsg(fRptPath & fRptName_Total & Environment.NewLine & "檔案不存在", MessageBoxIcon.Error)
                Application.Exit()
                Exit Sub
            End If

            If Environment.GetCommandLineArgs.Length <> 7 Then
                InfoMsg("傳入參數不正確 ! ", MessageBoxIcon.Error)
                Application.Exit()
                Exit Sub
            End If
            fDBName = Environment.GetCommandLineArgs(1)
            fDBUserId = Environment.GetCommandLineArgs(2)
            fDBPassword = Environment.GetCommandLineArgs(3)
            fCompCode = Environment.GetCommandLineArgs(4)
            fLoginName = Environment.GetCommandLineArgs(5)

        Catch ex As Exception
            XtraMessageBox.Show(ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Application.Exit()
        End Try
    End Sub


    Private Sub InfoMsg(ByVal aMsg As String, ByVal aMsgStyle As MessageBoxIcon)
        XtraMessageBox.Show(aMsg, "訊息", MessageBoxButtons.OK, aMsgStyle)
    End Sub
    Private Sub DesignRpt()
        
        Using objRpt As New XtraReport()
            objRpt.ShowDesignerDialog()
        End Using
    End Sub
    Private Sub PrintRpt(ByVal dt As DataTable, ByVal aRptType As RptType)
        Dim aRptFile As String = String.Empty
        If aRptType = RptType.Detail Then
            aRptFile = fRptPath & fRptName_Detail

        Else
            aRptFile = fRptPath & fRptName_Total
        End If
        Try

            Using objRpt As XtraReport = XtraReport.FromFile(aRptFile, True)
                objRpt.DataSource = dt
                objRpt.DataMember = dt.TableName
                objRpt.FillDataSource()
                objRpt.CreateDocument(True)

                objRpt.PrintingSystem.SetCommandVisibility(DevExpress.XtraPrinting.PrintingSystemCommand.Parameters, DevExpress.XtraPrinting.CommandVisibility.None)

                objRpt.PrintingSystem.SetCommandVisibility(DevExpress.XtraPrinting.PrintingSystemCommand.SendFile, DevExpress.XtraPrinting.CommandVisibility.None)

                If aRptType = RptType.Detail Then
                    objRpt.ExportOptions.PrintPreview.DefaultFileName = "Notify_Details"
                Else
                    objRpt.ExportOptions.PrintPreview.DefaultFileName = "Notify_Total"
                End If
                objRpt.ExportOptions.PrintPreview.DefaultDirectory = Application.StartupPath

                objRpt.PrintingSystem.ExportDefault = DevExpress.XtraPrinting.PrintingSystemCommand.ExportXls
                



                'Activator.CreateInstance

                'System.Runtime.InteropServices.Marshal.ReleaseComObject()
                



                objRpt.ShowPreviewDialog()
            End Using

        Catch ex As Exception
            InfoMsg(ex.Message, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub rptC06_Details_BeforePrint(ByVal sender As Object, ByVal e As System.Drawing.Printing.PrintEventArgs)
        If e.PrintAction = Printing.PrintAction.PrintToFile Then

        End If
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim aSQL As String = String.Empty

        Dim bdrCn As New OracleConnectionStringBuilder()
        Dim ds As New DataSet()
        Dim dt As New DataTable()

        bdrCn.DataSource = fDBName
        bdrCn.UserID = fDBUserId
        bdrCn.Password = fDBPassword
        bdrCn.Unicode = True
        Dim aRptType As RptType = RptType.Detail
        If chkNotifyCnt.Checked Then
            aRptType = RptType.Count
        End If

        Try
            Me.Cursor = Cursors.WaitCursor

            If Not IsDataOK() Then
                Exit Sub
            End If
            GetWhereString()
            aSQL = GetSelectString(aRptType)

            aSQL = aSQL & fWhereString
            If chkNotifyDetail.Checked Then
                aSQL = aSQL & " ORDER BY A.INVDATE, A.INVID "
            Else
                aSQL = aSQL & " GROUP BY A.INVDATE,A.COMPID ORDER BY A.INVDATE "
            End If
            fSQLText = aSQL
            Using cn As New OracleConnection(bdrCn.ConnectionString)
                cn.Open()
                Using cmd As OracleCommand = cn.CreateCommand
                    cmd.CommandText = aSQL
                    Using dap As New OracleDataAdapter(cmd)
                        dap.Fill(ds)
                        dt = ds.Tables(0)
                        If dt.Rows.Count = 0 Then
                            InfoMsg("無任何資料!", MessageBoxIcon.Information)
                        Else
                            PrintRpt(dt, aRptType)
                        End If
                    End Using
                End Using

            End Using

        Catch ex As Exception
            InfoMsg(ex.Message, MessageBoxIcon.Error)
        Finally
            ds.Dispose()
            dt.Dispose()
            Me.Cursor = Cursors.Default

        End Try
    End Sub
    Private Function IsDataOK() As Boolean
        Try
            If String.IsNullOrEmpty(dtInv1.Text) Then
                InfoMsg("請輸入發票起始日期!", MessageBoxIcon.Warning)
                dtInv1.Focus()
                Return False
            End If
            If String.IsNullOrEmpty(dtInv2.Text) Then
                InfoMsg("請輸入發票結束日期!", MessageBoxIcon.Warning)
                dtInv2.Focus()
                Return False
            End If
            Return True
        Catch ex As Exception
            InfoMsg(ex.Message, MessageBoxIcon.Error)
        End Try
    End Function

    Private Sub GetWhereString()
        fWhereString = String.Empty
        Try
            If String.IsNullOrEmpty(fWhereString) Then
                fWhereString = " WHERE INVOICEKIND = 1 "
            End If
            If Not String.IsNullOrEmpty(dtInv1.Text) Then
                fWhereString = fWhereString & " AND A.INVDATE >= TO_DATE('" & dtInv1.Text & "','yyyy/mm/dd')"
            End If
            If Not String.IsNullOrEmpty(dtInv2.Text) Then
                fWhereString = fWhereString & " AND A.INVDATE <= TO_DATE('" & dtInv2.Text & "','yyyy/mm/dd')"
            End If
            If chkSingleComp.Checked Then
                fWhereString = fWhereString & " AND A.COMPID='" & fCompCode & "'"
            End If
            If Not String.IsNullOrEmpty(edtInvId1.Text) Then
                fWhereString = fWhereString & " AND A.INVID >= '" & edtInvId1.Text & "'"
            End If
            If Not String.IsNullOrEmpty(edtInvId2.Text) Then
                fWhereString = fWhereString & " AND A.INVID <= '" & edtInvId2.Text & "'"
            End If
        Catch ex As Exception
            InfoMsg(ex.Message, MessageBoxIcon.Error)
            Application.Exit()
        End Try
    End Sub
    Private Function GetSelectString(ByVal aRptType As RptType) As String
        Dim aRetString As String = String.Empty
        Try
            Select Case aRptType
                Case RptType.Detail
                    aRetString = "SELECT A.* , " & _
                            " inv_single_check2(A.invid, A.identifyid1, A.identifyid2, A.compid, A.invuseid, A.isobsolete, 1) 作廢註記, " & _
                            " inv_single_check2(A.invid, A.identifyid1, A.identifyid2, A.compid, A.invuseid, A.isobsolete, 3) 捐贈註記, " & _
                            " inv_single_check2(A.invid, A.identifyid1, A.identifyid2, A.compid, A.invuseid, A.isobsolete, 4) 通知最終狀態 ," & _
                            " inv_single_check1(A.invid, A.identifyid1, A.identifyid2, A.compid, A.invuseid, A.isobsolete, 1, 1) Email發送 , " & _
                            " inv_single_check1(A.invid, A.identifyid1, A.identifyid2, A.compid, A.invuseid, A.isobsolete, 1, 2) Email回查, " & _
                            " inv_single_check1(A.invid, A.identifyid1, A.identifyid2, A.compid, A.invuseid, A.isobsolete, 1, 3) Email失敗, " & _
                            " inv_single_check1(A.invid, A.identifyid1, A.identifyid2, A.compid, A.invuseid, A.isobsolete, 2, 1) 簡訊發送, " & _
                            " inv_single_check1(A.invid, A.identifyid1, A.identifyid2, A.compid, A.invuseid, A.isobsolete, 2, 2) 簡訊回查, " & _
                            " inv_single_check1(A.invid, A.identifyid1, A.identifyid2, A.compid, A.invuseid, A.isobsolete, 2, 3) 簡訊失敗, " & _
                            " inv_single_check1(A.invid, A.identifyid1, A.identifyid2, A.compid, A.invuseid, A.isobsolete, 3, 1) TVmail發送, " & _
                            " inv_single_check1(A.invid, A.identifyid1, A.identifyid2, A.compid, A.invuseid, A.isobsolete, 3, 3) TVmail失敗, " & _
                            " inv_single_check1(A.invid, A.identifyid1, A.identifyid2, A.compid, A.invuseid, A.isobsolete, 4, 1) CM導流發送, " & _
                            " inv_single_check1(A.invid, A.identifyid1, A.identifyid2, A.compid, A.invuseid, A.isobsolete, 4, 3)  CM導流失敗" & _
                            " FROM inv007 A "
                Case RptType.Count
                    aRetString = "SELECT A.invdate,A.compid, " & _
                                " inv_total_check2(A.invdate,A.compid, 1) 總發票筆數 , " & _
                                " inv_total_check2(A.invdate, A.compid,2) 非電子發票筆數, " & _
                                " inv_total_check2(A.invdate,A.compid, 3) 電子發票筆數, " & _
                                " inv_total_check1(A.invdate, A.compid,6, 2) 非電子發票筆數_非作廢, " & _
                                " inv_total_check1(A.invdate,A.compid, 5, 2) 電子發票筆數_非作廢, " & _
                                " inv_total_check1(A.invdate,A.compid, 6, 1) 非電子發票筆數_作廢, " & _
                                " inv_total_check1(A.invdate,A.compid, 5, 1) 電子發票筆數_作廢, " & _
                                " inv_total_check2(A.invdate,A.compid, 4) 總作廢筆數, " & _
                                " inv_total_check1(A.invdate,A.compid, 6, 3) 非電子發票筆數_捐贈, " & _
                                " inv_total_check1(A.invdate,A.compid, 5, 3) 電子發票筆數_捐贈, " & _
                                " inv_total_check1(A.invdate,A.compid, 5, 4) 通知狀態成功筆數, " & _
                                " inv_total_check1(A.invdate,A.compid, 5, 5) 通知狀態失敗筆數, " & _
                                " inv_total_check1(A.invdate,A.compid, 1, 1) Email發送成功筆數, " & _
                                " inv_total_check1(A.invdate,A.compid, 1, 2) Email發送回查筆數, " & _
                                " inv_total_check1(A.invdate,A.compid, 1, 3) Email發送失敗筆數, " & _
                                " inv_total_check1(A.invdate,A.compid, 2, 1) 簡訊發送成功筆數, " & _
                                " inv_total_check1(A.invdate,A.compid, 2, 2) 簡訊回查成功筆數, " & _
                                " inv_total_check1(A.invdate,A.compid, 2, 3) 簡訊發送失敗筆數, " & _
                                " inv_total_check1(A.invdate,A.compid, 3, 1) TVmail發送成功筆數, " & _
                                " inv_total_check1(A.invdate, A.compid,3, 3) TVmail發送失敗筆數, " & _
                                " inv_total_check1(A.invdate, A.compid,4, 1) CM導流發送成功筆數, " & _
                                " inv_total_check1(A.invdate, A.compid,4, 3) CM導流發送失敗筆數 " & _
                                " FROM inv007 A "
            End Select
            Return aRetString
        Catch ex As Exception
            InfoMsg(ex.Message, MessageBoxIcon.Error)
            Return String.Empty
        End Try
    End Function

    Private Sub btnDesign_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDesign.Click
        Try
            Me.Cursor = Cursors.WaitCursor
            DesignRpt()
        Catch ex As Exception
            InfoMsg(ex.Message, MessageBoxIcon.Error)
        Finally
            Me.Cursor = Cursors.Default
        End Try



    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Application.Exit()
    End Sub

    Public Sub New()

        ' 此為 Windows Form 設計工具所需的呼叫。

        DevExpress.Skins.SkinManager.EnableFormSkins()
        'Dim o As New Editzhcn()
        'DevExpress.XtraEditors.Controls.Localizer.Active = o
        InitializeComponent()

        ' 在 InitializeComponent() 呼叫之後加入任何初始設定。

    End Sub

    Private Sub edtInvId1_InvalidValue(ByVal sender As System.Object, ByVal e As DevExpress.XtraEditors.Controls.InvalidValueExceptionEventArgs) Handles edtInvId1.InvalidValue, edtInvId2.InvalidValue
        e.ErrorText = "發票號碼不正確"
    End Sub

    Private Sub btnShowSQL_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnShowSQL.Click
        QrySQL()
        
    End Sub
    Private Sub QrySQL()
        If String.IsNullOrEmpty(fSQLText) Then
            XtraMessageBox.Show("尚未執行任何查詢！請先執行查詢動作", "訊息", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        Else
            frmShowSQL.SQLText = fSQLText
            frmShowSQL.ShowDialog()
        End If
    End Sub
End Class
Public Enum RptType
    Detail = 0
    Count = 1
End Enum

Public Class Editzhcn
    Inherits DevExpress.XtraEditors.Controls.Localizer
    Public Overrides Function GetLocalizedString(ByVal id As DevExpress.XtraEditors.Controls.StringId) As String
        'Return MyBase.GetLocalizedString(id)
        Select Case id
            Case Controls.StringId.DateEditToday
                Return "今天"

            Case Else
                Return MyBase.GetLocalizedString(id)
        End Select
    End Function
End Class
