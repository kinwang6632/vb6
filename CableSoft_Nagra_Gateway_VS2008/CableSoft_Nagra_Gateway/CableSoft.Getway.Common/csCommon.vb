Imports System
Imports System.Windows
Imports System.Windows.Forms
Imports System.Diagnostics
Imports System.IO
Public Class csCommon
    ''' <summary>
    ''' 取得程式目錄
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetAppPath() As String
        Dim strRet As String = Application.StartupPath
        If Not strRet.EndsWith("\") Then
            strRet = strRet & "\"
        End If
        Return strRet
    End Function
    ''' <summary>
    ''' 取出指定路徑的所有檔案
    ''' </summary>
    ''' <param name="AppPath">路徑</param>
    ''' <param name="aFileName">尋找的檔案</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overloads Shared Function GetFindFiles(ByVal AppPath As String, ByVal aFileName As String) As ArrayList
        Dim strFiles() As String
        Dim aPath As String = String.Empty
        aPath = AppPath
        If Not AppPath.EndsWith("\") Then
            aPath = aPath & "\"
        End If
        
        Try
            strFiles = Directory.GetFiles(aPath, aFileName).ToArray
            If strFiles.Length = 0 Then
                Return Nothing
            Else
                Dim retAry As New ArrayList
                retAry.AddRange(strFiles)
                Return retAry
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
            Return Nothing
        End Try

    End Function
    Public Overloads Shared Function GetFindFiles(ByVal aFileName As String) As ArrayList
        Return GetFindFiles(GetAppPath(), aFileName)
    End Function
    Public Shared Sub WriteLog(ByVal aMsg As String)
        Dim aFile As String = String.Empty
        Dim aDirectory As String = String.Empty
        aDirectory = GetAppPath() & "Log"
        aFile = aDirectory & "\GateWay_" & Format(Now, "yyyyMMdd")
        Try
            If Not Directory.Exists(aDirectory) Then
                Directory.CreateDirectory(aDirectory)
            End If
            Using f As New StreamWriter(aFile, True)
                f.WriteLine(String.Format("{0}->{1}", Format(Now, "yyyy:MM:dd hh:mm:ss"), aMsg))
            End Using
        Finally

        End Try
    End Sub
End Class
