Imports System.Text
Imports System.IO

Public Class Argument
    Implements IDisposable
    Private lstParameters As List(Of String)
    Public Enum AlignType As Integer
        Left = 0
        Right = 1
    End Enum
    Public Sub New(ByVal strParameters As String)
        lstParameters = strParameters.Split(" ").ToList
    End Sub
    Public Sub New(ByVal aryParameters As String())
        lstParameters = aryParameters.ToList
    End Sub

    Public Function GetB08Arguments() As String
        Dim aRet As String = String.Empty

        aRet = GetCompCode() & " " & lstParameters.Item(0) & " " & _
            lstParameters.Item(1) & " " & lstParameters.Item(2) & " " & _
            GetB08_04EDbSId() & " " & GetB08_05EUserId() & " " & GetB08_06EPassword() & " " & _
            GetB08_07ELoginUser() & " " & GetB08_08ELoginName()
        Return aRet

    End Function
    '#增加第13、14、15、16參數MisDbSid、MisDbUserId、MisDbPassword、MisDBOwner
    Public Function GetMisDbSid() As String
        Return lstParameters.Item(13)
    End Function
    Public Function GetMisDbUserId() As String
        Return lstParameters.Item(14)
    End Function
    Public Function GetMisDbPassword() As String
        Return lstParameters.Item(15)
    End Function
    Public Function GetMisDbOwner() As String
        Return lstParameters.Item(16)
    End Function

    '#6337 增加第12個參數 判斷INV001.StartAllUpload
    Public Function GetStartAllUpload() As Boolean
        If IsNumeric(lstParameters.Item(12)) Then
            Return Int32.Parse(lstParameters.Item(12)) = 1
        Else
            Return False
        End If
    End Function
    '#6080 增加第11個參數 B08 第9個參數 DBLink 
    Public Function GetDbLink() As String
        Dim aRet As String = String.Empty
        Try
            aRet = lstParameters.Item(11)
            If String.IsNullOrEmpty(aRet) Then
                Throw New Exception("傳入B08_8參數錯誤！")
            Else
                Return aRet
            End If
        Catch ex As Exception
            Throw New Exception("Function: GetDbLink 訊息: " & ex.ToString)
        End Try
    End Function

    '第10個參數 B08 第8個參數
    Public Function GetB08_08ELoginName() As String
        Dim aRet As String = String.Empty
        Try
            aRet = lstParameters.Item(10)
            If String.IsNullOrEmpty(aRet) Then
                Throw New Exception("傳入B08_8參數錯誤！")
            Else
                Return aRet
            End If
        Catch ex As Exception
            Throw New Exception("Function: GetB08_08ELoginName 訊息: " & ex.ToString)
        End Try
    End Function

    '第9個參數 B08 第7個參數
    Public Function GetB08_07ELoginUser() As String
        Dim aRet As String = String.Empty
        Try
            aRet = lstParameters.Item(9)
            If String.IsNullOrEmpty(aRet) Then
                Throw New Exception("傳入B08_7參數錯誤！")
            Else
                Return aRet
            End If
        Catch ex As Exception
            Throw New Exception("Function: GetB08_07ELoginUser 訊息: " & ex.ToString)
        End Try
    End Function

    '第8個參數 B08 第6個參數
    Public Function GetB08_06EPassword() As String
        Dim aRet As String = String.Empty
        Try
            aRet = lstParameters.Item(8)
            If String.IsNullOrEmpty(aRet) Then
                Throw New Exception("傳入B08_6參數錯誤！")
            Else
                Return aRet
            End If
        Catch ex As Exception
            Throw New Exception("Function: GetB08_06EPassword 訊息: " & ex.ToString)
        End Try
    End Function
    '第7個參數 B08 第5個參數
    Public Function GetB08_05EUserId() As String
        Dim aRet As String = String.Empty
        Try
            aRet = lstParameters.Item(7)
            If String.IsNullOrEmpty(aRet) Then
                Throw New Exception("傳入B08_5參數錯誤！")
            Else
                Return aRet
            End If
        Catch ex As Exception
            Throw New Exception("Function: GetB08_05EUserId 訊息: " & ex.ToString)
        End Try
    End Function
    '第6個參數 B08 第4個參數
    Public Function GetB08_04EDbSId() As String
        Dim aRet As String = String.Empty
        Try
            aRet = lstParameters.Item(6)
            If String.IsNullOrEmpty(aRet) Then
                Throw New Exception("傳入B08_4參數錯誤！")
            Else
                Return aRet
            End If
        Catch ex As Exception
            Throw New Exception("Function: GetB08_04EDbSId 訊息: " & ex.ToString)
        End Try
    End Function
    ''' <summary>
    ''' Owner 參數 第4個參數
    ''' </summary>
    ''' <remarks></remarks>
    Public Function GetOwner() As String
        Dim aRet As String = String.Empty
        Return aRet
        Try
            aRet = lstParameters.Item(4)

            If String.IsNullOrEmpty(aRet) Then
                Throw New Exception("傳入OWNER參數錯誤！")
            Else
                If aRet.EndsWith(".") Then
                    Return aRet
                Else
                    Return aRet & "."
                End If
            End If

        Catch ex As Exception
            Throw New Exception("Function: GetOwner 訊息: " & ex.ToString)
            'Return String.Empty
        End Try
    End Function
    ''' <summary>
    ''' CompCode 第3個參數
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetCompCode() As String
        Dim aRet As String = String.Empty

        Try
            aRet = lstParameters.Item(3) 'My.Application.CommandLineArgs(3)
            If String.IsNullOrEmpty(aRet) Then
                Throw New Exception("傳入公司別參數錯誤！")
            Else
                Return aRet
            End If
        Catch ex As Exception
            Throw New Exception("Function: GetCompCode 訊息: " & ex.ToString)
        End Try
    End Function
    ''' <summary>
    ''' 是否啟用通知 第5個參數
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetStarNotify() As Boolean
        Dim aRet As String = String.Empty
        Try
            aRet = lstParameters.Item(5) '  My.Application.CommandLineArgs(5)
            If String.IsNullOrEmpty(aRet) OrElse aRet = "0" Then
                Return False
            Else
                Return True
            End If
        Catch ex As Exception
            Throw New Exception("Function: GetStarNotify 訊息: " & ex.ToString)
        End Try


    End Function
    ''' <summary>
    ''' Connect資訊 第0~2參數
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetConnect() As String

        Dim aRet As String = String.Empty
        aRet = "Data Source={0};Persist Security Info=True;User ID={1};Password={2};Unicode=True"
        Try

            aRet = String.Format(aRet, lstParameters.Item(0), _
                                lstParameters.Item(1), _
                                lstParameters.Item(2))
            If String.IsNullOrEmpty(aRet) Then
                Throw New Exception("傳入連線字串參數錯誤！")
            Else
                Return aRet
            End If
        Catch ex As Exception
            Throw New Exception("Function: GetConnect 訊息: " & ex.ToString)
        End Try
    End Function
    Public Shared Sub WriteErr(ByVal aErrMsg As String)
        Dim sbErr As New StringBuilder()
        Dim aWriteFile As String = GetAppPath() & "ElectronErr.Txt"
        sbErr.AppendLine("錯誤日期：" & Now & "  錯誤原因：" & aErrMsg)
        Try

            Using oWriter As New StreamWriter(aWriteFile, True)
                oWriter.WriteLine(sbErr.ToString())
                oWriter.Flush()
                oWriter.Close()
            End Using
        Catch ex As Exception
            Throw New Exception("Function: WriteErr 訊息: " & ex.ToString)
        End Try
    End Sub
    Public Shared Sub WriteLog(ByVal aFile As String, ByVal aMsg As String)
        Try
            Using oWriter As New StreamWriter(aFile, False)
                oWriter.WriteLine(aMsg)
                oWriter.Flush()
                oWriter.Close()

            End Using
        Catch ex As Exception
            Throw New Exception("Function: WriteLog 訊息: " & ex.ToString)
        End Try
    End Sub

    Public Shared Sub WriteProdcutLog(ByVal aPath As String, ByVal aContext As String)
        Dim aFile As String = aPath & "電子發票上傳記錄_" & Format(Now, "yyyyMMddHHddss") & ".TXT"
        Try
            Using oWriter As New StreamWriter(aFile, False)
                oWriter.WriteLine(aContext)
                oWriter.Flush()
                oWriter.Close()
            End Using
        Catch ex As Exception
            Throw New Exception("Function: WriteLog 訊息: " & ex.ToString)
        End Try
    End Sub
    Public Shared Function GetAppPath() As String
        Dim aRet As String = String.Empty
        aRet = System.IO.Directory.GetCurrentDirectory  'Application.StartupPath
        If aRet.EndsWith("\") Then
            Return aRet
        Else
            Return aRet & "\"
        End If
    End Function
    Public Shared Function GetString(ByVal aString As Object, ByVal aLength As Integer, _
                                    Optional ByVal aAllign As AlignType = AlignType.Left, _
                                    Optional ByVal aFillZero As Boolean = False) As String
        Dim strFill As String = String.Empty
        Dim intLength As Integer = 0
        Dim aRet As String = String.Empty
        Try

            If aString Is Nothing OrElse IsDBNull(aString) Then
                aString = ""
            End If



            Dim bts As Byte() = Encoding.GetEncoding(950).GetBytes(aString)

            intLength = Encoding.GetEncoding(950).GetByteCount(aString)

            'Encoding.ASCII.GetBytes(aString)

            'If aAllign = AlignType.Left Then
            '    'aRet = Encoding.GetEncoding(950).GetChars(bts, 0, aLength)
            '    Return Left(Encoding.GetEncoding(950).GetChars(bts, 0, intLength), aLength)

            'Else

            '    Return Right(Encoding.GetEncoding(950).GetChars(bts, 0, intLength), aLength)
            'End If
            If intLength > aLength Then
                aRet = Encoding.GetEncoding(950).GetChars(bts, 0, aLength)
            Else
                aRet = Encoding.GetEncoding(950).GetChars(bts, 0, intLength)
            End If


            If aLength - intLength > 0 Then
                If aFillZero Then
                    strFill = New String("0"c, aLength - intLength)
                Else
                    strFill = New String(" ", aLength - intLength)
                End If
            Else
                strFill = ""
            End If
            If aAllign = AlignType.Left Then
                aRet = aRet & strFill
            Else
                aRet = strFill & aRet
            End If
            Return aRet
        Catch ex As Exception
            Throw New Exception("Function: GetString 訊息: " & ex.ToString)
        End Try


    End Function

#Region "IDisposable Support"
    Private disposedValue As Boolean ' 偵測多餘的呼叫

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: 處置 Managed 狀態 (Managed 物件)。
                'If lstParameters IsNot Nothing Then
                '    lstParameters.Clear()
                '    lstParameters = Nothing
                'End If
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
