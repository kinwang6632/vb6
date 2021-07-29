Public Class frmPosAndGet
    Private Sub btnPost_Click(sender As Object, e As EventArgs) Handles btnPost.Click
        Dim Name As String = "POST"
        Dim password As String = "Your Passord"
        'Dim postString As String = String.Format("Name={0}&Password={1}", Name, password)
        Dim postString As String = String.Format("Name={0}", Name)
        Dim req As System.Net.HttpWebRequest = System.Net.HttpWebRequest.Create("http://localhost/default.aspx")
        Dim contentType As String = "application/x-www-form-urlencoded"
        req.Method = System.Net.WebRequestMethods.Http.Post
        req.ContentType = contentType
        req.ContentLength = System.Text.Encoding.UTF8.GetByteCount(postString)
        req.Timeout = 30000
        req.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"
        Dim requestWriter As System.IO.StreamWriter = New System.IO.StreamWriter(req.GetRequestStream)

        Try
            requestWriter.Write(postString)
            requestWriter.Flush()
            requestWriter.Close()
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
        Dim response As System.Net.WebResponse = req.GetResponse



        If CType(response, System.Net.HttpWebResponse).StatusCode = Net.HttpStatusCode.OK Then
            Dim responseReader As System.IO.StreamReader = New System.IO.StreamReader(response.GetResponseStream)

            Dim responseData As String = responseReader.ReadToEnd

            responseReader.Close()
            req.GetResponse.Close()
        Else

        End If
    End Sub
    Private Sub btnGet_Click(sender As Object, e As EventArgs) Handles btnGet.Click
        Dim Name As String = "GET1"
        Dim password As String = "Your Passord"
        'Dim postString As String = String.Format("Name={0}&Password={1}", Name, password)
        Dim postString As String = String.Format("Name={0}", Name)
        Dim url As String = "http://localhost/default.aspx"
        If Not String.IsNullOrEmpty(postString) Then
            url = url & "?" & postString
        End If
        Dim req As System.Net.HttpWebRequest = System.Net.HttpWebRequest.Create(url)

        Dim contentType As String = "application/x-www-form-urlencoded"
        req.Method = System.Net.WebRequestMethods.Http.Get

        req.ContentType = contentType

        req.Timeout = req.Timeout
        'req.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"

        Dim response As System.Net.WebResponse = req.GetResponse



        If CType(response, System.Net.HttpWebResponse).StatusCode = Net.HttpStatusCode.OK Then
            Dim responseReader As System.IO.StreamReader = New System.IO.StreamReader(response.GetResponseStream)

            Dim responseData As String = responseReader.ReadToEnd

            responseReader.Close()
            req.GetResponse.Close()
        Else

        End If
    End Sub
End Class
