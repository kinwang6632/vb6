Imports System
Imports System.Windows
Imports System.Windows.Forms
Imports System.Runtime.Serialization

Public Class CablesoftException
    Inherits System.Exception
    Implements ISerializable

    Private _ex As System.Exception = Nothing
    Public Sub New()
        MyBase.New()
    End Sub

    Public Sub New(ByVal exSource As System.Exception, ByVal blnWriteErrTxt As Boolean)
        Try
            _ex = exSource
            If _ex.GetType Is Me.GetType Then
                Dim stk As New StackTrace(_ex)
           
                Exit Sub
            End If
            If blnWriteErrTxt Then

                Exit Sub
                'WriteError()
            End If

        Catch ex As Exception

        End Try
        
    End Sub

    Public Sub New(ByVal ex As System.Exception, ByVal blnWriteErrTxt As Boolean, ByVal strErrTxtFileName As String)
        _ex = ex
        If _ex.GetType Is Me.GetType Then
            Exit Sub
        End If
        If blnWriteErrTxt Then
            WriteError(strErrTxtFileName)
        End If
    End Sub
    Public Overloads Sub WriteError()
        WriteErrTxtLog.WriteTxtError(_ex, Nothing)
    End Sub
    Public Overloads Sub WriteError(ByVal ErrFile As String)
        WriteErrTxtLog.WriteTxtError(_ex, Nothing, ErrFile)
    End Sub

End Class
