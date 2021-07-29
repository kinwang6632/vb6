Imports System
Imports System.IO
Imports System.Text

Public Class clsCommon
    Public Shared Function GetB08Arguments() As String
        Dim aRet As String = String.Empty
        aRet = GetCompCode() & " " & My.Application.CommandLineArgs(0) & " " & _
            My.Application.CommandLineArgs(1) & " " & My.Application.CommandLineArgs(2) & " " & _
            GetB08_04EDbSId() & " " & GetB08_05EUserId() & " " & GetB08_06EPassword() & " " & _
            GetB08_07ELoginUser() & " " & GetB08_08ELoginName()
        Return aRet

    End Function
    '#增加第13、14、15、16參數MisDbSid、MisDbUserId、MisDbPassword、MisDBOwner
    Public Shared Function GetMisDbSid() As String
        Return My.Application.CommandLineArgs(13)
    End Function
    Public Shared Function GetMisDbUserId() As String
        Return My.Application.CommandLineArgs(14)
    End Function
    Public Shared Function GetMisDbPassword() As String
        Return My.Application.CommandLineArgs(15)
    End Function
    Public Shared Function GetMisDbOwner() As String
        Return My.Application.CommandLineArgs(16)
    End Function

    '#6337 增加第12個參數 判斷INV001.StartAllUpload
    Public Shared Function GetStartAllUpload() As Boolean
        Return Int32.Parse(My.Application.CommandLineArgs(12)) = 1
    End Function
    '#6080 增加第11個參數 B08 第9個參數 DBLink 
    Public Shared Function GetDbLink() As String
        Dim aRet As String = String.Empty
        Try
            aRet = My.Application.CommandLineArgs(11)
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
    Public Shared Function GetB08_08ELoginName() As String
        Dim aRet As String = String.Empty
        Try
            aRet = My.Application.CommandLineArgs(10)
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
    Public Shared Function GetB08_07ELoginUser() As String
        Dim aRet As String = String.Empty
        Try
            aRet = My.Application.CommandLineArgs(9)
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
    Public Shared Function GetB08_06EPassword() As String
        Dim aRet As String = String.Empty
        Try
            aRet = My.Application.CommandLineArgs(8)
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
    Public Shared Function GetB08_05EUserId() As String
        Dim aRet As String = String.Empty
        Try
            aRet = My.Application.CommandLineArgs(7)
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
    Public Shared Function GetB08_04EDbSId() As String
        Dim aRet As String = String.Empty
        Try
            aRet = My.Application.CommandLineArgs(6)
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
    Public Shared Function GetOwner() As String
        Dim aRet As String = String.Empty
        Return aRet

        Try
            aRet = My.Application.CommandLineArgs(4)

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
            Return String.Empty
        End Try
    End Function
    ''' <summary>
    ''' CompCode 第3個參數
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetCompCode() As String
        Dim aRet As String = String.Empty

        Try
            aRet = My.Application.CommandLineArgs(3)
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
    Public Shared Function GetStarNotify() As Boolean

        Dim aRet As String = String.Empty
        Try
            aRet = My.Application.CommandLineArgs(5)
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
    Public Shared Function GetConnect() As String

        Dim aRet As String = String.Empty
        aRet = "Data Source={0};Persist Security Info=True;User ID={1};Password={2};Unicode=True"
        Try

            aRet = String.Format(aRet, My.Application.CommandLineArgs(0), _
                                My.Application.CommandLineArgs(1), _
                                My.Application.CommandLineArgs(2))
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
        aRet = Application.StartupPath
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

End Class
Public Enum AlignType As Integer
    Left = 0
    Right = 1
End Enum