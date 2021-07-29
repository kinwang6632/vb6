Imports Oracle.DataAccess.Client
Imports System.Text
Imports System.IO

Public Class Form1
    Dim strCn As String = "Data Source=RDKNET;Persist Security Info=True;User ID=EMCNCC;Password=EMCNCC;Min Pool Size=1;Max Pool Size=1000;Connection Lifetime=6000"
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Using oraCn As New OracleConnection(strCn)
            oraCn.Open()
            Using oraCmd As New OracleCommand("sf_GetDiscount1", oraCn)
                Dim strDis As String = Nothing
                Dim strRetMsg As String = Nothing
                oraCmd.CommandType = CommandType.StoredProcedure
                oraCmd.Parameters.Clear()
                With oraCmd.Parameters
                    .Add("RETURN_VALUE", OracleDbType.Int32, ParameterDirection.ReturnValue)
                    .Add("p_CustId", OracleDbType.Int32, 2000, 1, ParameterDirection.Input)
                    .Add("p_CitemCode", OracleDbType.Int32, 2000, 1, ParameterDirection.Input)
                    .Add("p_StopDate", OracleDbType.Varchar2, 2000, "20080401", ParameterDirection.Input)
                    .Add("p_CompCode", OracleDbType.Int32, 2000, 1, ParameterDirection.Input)
                    .Add("p_ServiceType", OracleDbType.Varchar2, 2000, "C", ParameterDirection.Input)
                    .Add("p_DiscountAmt", OracleDbType.Int32, ParameterDirection.Output)
                    .Add("p_RetMsg", OracleDbType.Varchar2, ParameterDirection.Output)
                End With
                oraCmd.Parameters("RETURN_VALUE").Size = 4000
                oraCmd.Parameters("p_DiscountAmt").Size = 2000
                oraCmd.Parameters("p_RetMsg").Size = 4000
                oraCmd.ExecuteNonQuery()
                MsgBox(oraCmd.Parameters("p_RetMsg").Value)
                MsgBox(oraCmd.Parameters("RETURN_VALUE").Value)
            End Using
        End Using
        MsgBox("完成")
    End Sub
    '使用ADO.NET直接抓取Oracle Store Procedure 的參數集合
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Using cnn As New System.Data.OracleClient.OracleConnection(strCn)
            cnn.Open()
            Using cmd As New System.Data.OracleClient.OracleCommand("sf_GetDiscount1", cnn)
                cmd.CommandType = CommandType.StoredProcedure
                System.Data.OracleClient.OracleCommandBuilder.DeriveParameters(cmd)
                For i As Integer = 0 To cmd.Parameters.Count - 1
                    cmd.Parameters(i).Size = 4000
                    Select Case UCase(cmd.Parameters(i).ParameterName)
                        Case "P_CUSTID"
                            cmd.Parameters(i).Value = 1
                        Case "P_CITEMCODE"
                            cmd.Parameters(i).Value = 1
                        Case "P_STOPDATE"
                            cmd.Parameters(i).Value = "20080401"
                        Case "P_COMPCODE"
                            cmd.Parameters(i).Value = 5
                        Case "P_SERVICETYPE"
                            cmd.Parameters(i).Value = "C"
                        Case Else
                    End Select
                Next
                cmd.ExecuteNonQuery()
                MsgBox(cmd.Parameters("p_DiscountAmt").Value.ToString)
                MsgBox(cmd.Parameters("RETURN_VALUE").Value.ToString)
            End Using
        End Using

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

    End Sub
End Class
