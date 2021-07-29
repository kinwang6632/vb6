Public Class _Default
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim Name As String = Nothing
        Response.Clear()

        If Request.HttpMethod = "GET" Then
            If Request.QueryString("alive") = "heartbeat" Then
                Response.AddHeader("Content-Length", "True".Length)
                Response.Write("True")
            Else
                Response.AddHeader("Content-Length", "Get Error".Length)
                Response.Write("Get Error")

            End If
        ElseIf Request.HttpMethod = "POST" Then

            If Request.Form("alive") = "heartbeat" Then
                Response.AddHeader("Content-Length", "True".Length)
                Response.Write("True")
            Else
                If String.IsNullOrEmpty(Request.Form("alive")) Then
                    Response.AddHeader("Content-Length", "alive nothing".Length)
                    Response.Write("alive nothing")
                Else
                    Response.AddHeader("Content-Length", Request.Form(0).ToString.Length)
                    Response.Write(Request.Form(0).ToString.Length)
                End If
             
            End If
        Else

            Response.Write("Other Method")
        End If
    End Sub

End Class