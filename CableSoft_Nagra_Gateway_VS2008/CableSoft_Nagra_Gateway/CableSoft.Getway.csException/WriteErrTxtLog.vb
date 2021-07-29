Imports System
Imports System.Windows
Imports System.Windows.Forms
Imports System.Diagnostics
Imports System.IO
Public Class WriteErrTxtLog
    Private Shared _ErrFile As String = String.Empty
    Private Shared intlock As Int32 = 0 'Lock用的,避免Thread同時存取同一個檔案
    Private Shared intSysEvent As Int32 = 0
    Private Shared intCmdErr As Int32 = 0
    Private Shared intSOStatus As Int32 = 0
    Public Overloads Shared Sub WriteTxtError(ByVal exSource As Exception, ByVal OtherMsg As String)

        SyncLock intlock.GetType
            Try
                Dim stk As New StackTrace(exSource, True)
                Dim aMethod As String = String.Empty
                Dim aErrLineNumber As Int32 = 0
                aMethod = stk.GetFrames(stk.FrameCount - 1).GetMethod.Name
                aErrLineNumber = stk.GetFrames(stk.FrameCount - 1).GetFileLineNumber
                If aErrLineNumber <= 0 Then
                    aMethod = stk.GetFrame(0).GetMethod.Name
                    aErrLineNumber = stk.GetFrame(0).GetFileLineNumber
                End If


                Dim aMsg As String = String.Format(CableSoftErrMsg, Now.ToString, _
                                        aMethod, exSource.Message, aErrLineNumber.ToString())

                If Not String.IsNullOrEmpty(OtherMsg) Then
                    aMsg = aMsg & Environment.NewLine & OtherMsg
                End If
                aMsg = aMsg & Environment.NewLine & _
                    New String("-"c, 300)


                Dim aDirectory As String = GetAppPath() & "ErrTxtLog\" & Format(Now, "yyyyMMdd")
                If Not Directory.Exists(aDirectory) Then
                    Directory.CreateDirectory(aDirectory)
                End If
                If Not aDirectory.EndsWith("\") Then
                    aDirectory = aDirectory & "\"
                End If
                If String.IsNullOrEmpty(_ErrFile) Then
                    _ErrFile = aDirectory & ErrTxtFileName
                End If

                Using oFile As New StreamWriter(_ErrFile, True)
                    oFile.WriteLine(aMsg)
                    oFile.Flush()
                    oFile.Close()
                End Using

            Catch ex As Exception

            End Try
        End SyncLock
    End Sub
    Public Overloads Shared Sub WriteTxtError(ByVal exSource As Exception, ByVal OtherMsg As String, _
                                            ByVal aFile As String)
        Try
            _ErrFile = aFile
            WriteTxtError(exSource, OtherMsg)
        Catch ex As Exception

        End Try
        
    End Sub
    Public Overloads Shared Sub WriteTxtError(ByVal exSource As Exception, ByVal aCompId As String, _
                                              ByVal aSeqNo As String, _
                                              ByVal aCmd As String, _
                                              ByVal OtherMsg As String, _
                                              ByVal aFile As String)
        Try

            Dim aResultMsg As String = String.Empty
            If Not String.IsNullOrEmpty(aCompId) Then
                aResultMsg = "公司別 : " & aCompId & Environment.NewLine
            End If
            If Not String.IsNullOrEmpty(aSeqNo) Then
                aResultMsg = aResultMsg & "命令序號 : " & aSeqNo & Environment.NewLine
            End If

            If Not String.IsNullOrEmpty(aCmd) Then
                aResultMsg = aResultMsg & "執行命令 : " & aCmd & Environment.NewLine
            End If
            aResultMsg = aResultMsg & OtherMsg
            WriteTxtError(exSource, aResultMsg, aFile)
        Catch ex As Exception

        End Try

    End Sub
    Public Overloads Shared Sub WriteSOStatus(ByVal aMsg As String)
        SyncLock intSOStatus.GetType
            Try
                Dim aFile As String = String.Empty
                Dim aDirectory As String = GetAppPath() & "SOStatusLog\" & Format(Now, "yyyyMMdd")
                If Not Directory.Exists(aDirectory) Then
                    Directory.CreateDirectory(aDirectory)
                End If
                If Not aDirectory.EndsWith("\") Then
                    aDirectory = aDirectory & "\"
                End If
                aFile = aDirectory & SOStatusLogName
                Using oFile As New StreamWriter(aFile, True)
                    oFile.WriteLine(Format(Now, "yyyy/MM/dd hh:mm:ss") & " > " & aMsg)
                    oFile.Flush()
                    oFile.Close()
                End Using

            Catch ex As Exception

            End Try
        End SyncLock
    End Sub

    Public Overloads Shared Sub WriteCmdErrLog(ByVal aMsg As String)
        SyncLock intCmdErr.GetType
            Try
                Dim aFile As String = String.Empty
                Dim aDirectory As String = GetAppPath() & "CmdErrLog\" & Format(Now, "yyyyMMdd")
                If Not Directory.Exists(aDirectory) Then
                    Directory.CreateDirectory(aDirectory)
                End If
                If Not aDirectory.EndsWith("\") Then
                    aDirectory = aDirectory & "\"
                End If
                aFile = aDirectory & CmdErrLogName
                Using oFile As New StreamWriter(aFile, True)
                    oFile.WriteLine(Format(Now, "yyyy/MM/dd hh:mm:ss") & " > " & aMsg)
                    oFile.Flush()
                    oFile.Close()
                End Using

            Catch ex As Exception

            End Try
        End SyncLock
    End Sub

    '寫入系統訊息Log
    Public Overloads Shared Sub WriteSysEventLog(ByVal aMsg As String)
        SyncLock intSysEvent.GetType
            Try
                Dim aFile As String = String.Empty
                Dim aDirectory As String = GetAppPath() & "SysEventLog\" & Format(Now, "yyyyMMdd")
                If Not Directory.Exists(aDirectory) Then
                    Directory.CreateDirectory(aDirectory)
                End If
                If Not aDirectory.EndsWith("\") Then
                    aDirectory = aDirectory & "\"
                End If
                aFile = aDirectory & SysEventLogName

                Using oFile As New StreamWriter(aFile, True)
                    oFile.WriteLine(Format(Now, "yyyy/MM/dd hh:mm:ss") & " > " & aMsg)
                    oFile.Flush()
                    oFile.Close()
                End Using

                
            Catch ex As Exception

            End Try
        End SyncLock
    End Sub


End Class
