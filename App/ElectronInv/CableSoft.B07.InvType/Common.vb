Imports System.IO
Imports System.Text

Public Class Common
    Public Shared Function GetAppPath() As String
        Dim aRet As String = String.Empty
        aRet = Environment.CurrentDirectory ' Application.StartupPath
        If aRet.EndsWith("\") Then
            Return aRet
        Else
            Return aRet & "\"
        End If
    End Function
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