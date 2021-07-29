
Partial Class _Default
    Inherits System.Web.UI.Page

    Protected Sub form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles form1.Load
        Dim x As String = String.Empty
        Dim A As NameValueCollection
        Response.Write("Result" & Request.QueryString("1231"))


    End Sub
End Class
