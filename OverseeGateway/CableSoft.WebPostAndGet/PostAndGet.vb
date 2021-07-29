Public Class PostAndGet
    Implements IDisposable
    Private url As String
    Private para As String
    Private timeout As Integer = 0
    Private Const contentType As String = "application/x-www-form-urlencoded"
    Private requestMethod As String
    Public Sub New(ByVal url As String, ByVal para As String, ByVal timeout As Integer, ByVal requestMethod As String)
        Me.url = url
        Me.para = para
        Me.timeout = timeout
        Me.requestMethod = requestMethod
        'If Not Create() Then
        '    Throw New Exception("HttpWebRequest has been created in failure.")
        'End If
    End Sub
    Public Sub New(ByVal url As String, ByVal para As String, ByVal requestMethod As String)
        'Me.url = url
        'Me.para = para
        'Me.timeout = timeout
        Me.New(url, para, 0, requestMethod)
    End Sub
    Public Function Execute(ByVal requestMethod As String) As Object
        Dim req As Threading.ThreadLocal(Of System.Net.HttpWebRequest)
        Dim response As Threading.ThreadLocal(Of System.Net.WebResponse)
        Try
            Select Case requestMethod.ToUpper
                Case System.Net.WebRequestMethods.Http.Post
                    req.Value = System.Net.HttpWebRequest.Create(url)
                    req.Value.Method = System.Net.WebRequestMethods.Http.Post
                    req.Value.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"
                    req.Value.ContentLength = System.Text.Encoding.UTF8.GetByteCount(para)
                    req.Value.ReadWriteTimeout = Me.timeout
                    Using requestWriter As System.IO.StreamWriter =
                            New System.IO.StreamWriter(req.Value.GetRequestStream)
                        Try
                            requestWriter.Write(Me.para)
                            requestWriter.Flush()
                            requestWriter.Close()
                        Catch exWriter As Exception
                            Throw exWriter
                        End Try
                    End Using
                Case System.Net.WebRequestMethods.Http.Get
                    req.Value = System.Net.HttpWebRequest.Create(String.Format("{0}?{1}", url, para))
                    req.Value.Method = System.Net.WebRequestMethods.Http.Get
            End Select
            req.Value.ContentType = contentType
            If timeout > 0 Then req.Value.Timeout = Me.timeout
            response.Value = req.Value.GetResponse
            If CType(response.Value, System.Net.HttpWebResponse).StatusCode <> Net.HttpStatusCode.OK Then
                Throw New Exception("Response fail.")
            Else
                Using responseReader As System.IO.StreamReader =
                                    New System.IO.StreamReader(response.Value.GetResponseStream)
                    Return responseReader.ReadToEnd
                    responseReader.Close()
                End Using
            End If
        Catch ex As Exception
            Throw
        Finally
            If req IsNot Nothing Then
                req.Value = Nothing
                req.Dispose()
                req = Nothing
            End If
            If response IsNot Nothing Then
                response.Value = Nothing
                response.Dispose()
                response = Nothing
            End If
        End Try
        
    End Function
   
    'Public Sub Post()
    '    Dim Name As String = "POST"
    '    Dim password As String = "Your Passord"
    '    'Dim postString As String = String.Format("Name={0}&Password={1}", Name, password)
    '    Dim postString As String = String.Format("Name={0}", Name)
    '    Dim req As System.Net.HttpWebRequest = System.Net.HttpWebRequest.Create("http://localhost/default.aspx")
    '    Dim contentType As String = "application/x-www-form-urlencoded"
    '    req.Method = System.Net.WebRequestMethods.Http.Post
    '    req.ContentType = contentType
    '    req.ContentLength = System.Text.Encoding.UTF8.GetByteCount(postString)
    '    req.Timeout = 30000


    '    req.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"
    '    Dim requestWriter As System.IO.StreamWriter = New System.IO.StreamWriter(req.GetRequestStream)

    '    Try
    '        requestWriter.Write(postString)
    '        requestWriter.Flush()
    '        requestWriter.Close()
    '    Catch ex As Exception
    '        MessageBox.Show(ex.ToString)
    '    End Try
    '    Dim response As System.Net.WebResponse = req.GetResponse



    '    If CType(response, System.Net.HttpWebResponse).StatusCode = Net.HttpStatusCode.OK Then
    '        Dim responseReader As System.IO.StreamReader = New System.IO.StreamReader(response.GetResponseStream)

    '        Dim responseData As String = responseReader.ReadToEnd

    '        responseReader.Close()
    '        req.GetResponse.Close()
    '    Else

    '    End If
    'End Sub
    'Private Sub Get(sender As Object, e As EventArgs) 
    '    Dim Name As String = "GET1"
    '    Dim password As String = "Your Passord"
    '    'Dim postString As String = String.Format("Name={0}&Password={1}", Name, password)
    '    Dim postString As String = String.Format("Name={0}", Name)
    '    Dim url As String = "http://localhost/default.aspx"
    '    If Not String.IsNullOrEmpty(postString) Then
    '        url = url & "?" & postString
    '    End If
    '    Dim req As System.Net.HttpWebRequest = System.Net.HttpWebRequest.Create(url)

    '    Dim contentType As String = "application/x-www-form-urlencoded"
    '    req.Method = System.Net.WebRequestMethods.Http.Get

    '    req.ContentType = contentType

    '    req.Timeout = req.Timeout
    '    'req.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"

    '    Dim response As System.Net.WebResponse = req.GetResponse



    '    If CType(response, System.Net.HttpWebResponse).StatusCode = Net.HttpStatusCode.OK Then
    '        Dim responseReader As System.IO.StreamReader = New System.IO.StreamReader(response.GetResponseStream)

    '        Dim responseData As String = responseReader.ReadToEnd

    '        responseReader.Close()
    '        req.GetResponse.Close()
    '    Else

    '    End If
    'End Sub

#Region "IDisposable Support"
    Private disposedValue As Boolean ' 偵測多餘的呼叫

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: 處置 Managed 狀態 (Managed 物件)。
                requestMethod = Nothing
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
