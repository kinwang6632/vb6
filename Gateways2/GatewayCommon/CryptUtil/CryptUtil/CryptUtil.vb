Imports Microsoft.VisualBasic
Imports System.Security.Cryptography
Imports System.IO
#Region "連線參數加解密"
Public Class CryptUtil
    Private Shared KEY_64() As Byte = {42, 16, 93, 156, 78, 4, 218, 32}
    Private Shared IV_64() As Byte = {55, 103, 246, 79, 36, 99, 167, 3}

    Private Shared KEY_192() As Byte = {42, 16, 93, 156, 78, 4, 218, 32, 15, 167, 44, 80, 26, 250, 155, 112, 2, 94, 11, 204, 119, 35, 184, 197}
    Private Shared IV_192() As Byte = {55, 103, 246, 79, 36, 99, 167, 3, 42, 5, 62, 83, 184, 7, 209, 13, 145, 23, 200, 58, 173, 10, 121, 222}

    Public Shared Function Encrypt(ByVal value As String) As String
        Dim cryptoProvider As New DESCryptoServiceProvider
        Dim ms As New MemoryStream()
        Dim cs As New CryptoStream(ms, cryptoProvider.CreateEncryptor(KEY_64, IV_64), CryptoStreamMode.Write)
        Dim sw As New StreamWriter(cs)
        sw.Write(value)
        sw.Flush()
        cs.FlushFinalBlock()
        ms.Flush()
        Return Convert.ToBase64String(ms.GetBuffer(), 0, ms.Length)
    End Function

    Public Shared Function Decrypt(ByVal value As String) As String
        Dim cryptoProvider As New DESCryptoServiceProvider()
        Dim buffer As Byte() = Convert.FromBase64String(value)
        Dim ms As MemoryStream = New MemoryStream(buffer)
        Dim cs As New CryptoStream(ms, cryptoProvider.CreateDecryptor(KEY_64, IV_64), CryptoStreamMode.Read)
        Dim sr As StreamReader = New StreamReader(cs)
        Return sr.ReadToEnd()
    End Function

    Public Shared Function EncryptTripleDES(ByVal value As String) As String
        Dim cryptoProvider As New TripleDESCryptoServiceProvider()
        Dim ms As New MemoryStream()
        Dim cs As New CryptoStream(ms, cryptoProvider.CreateEncryptor(KEY_192, IV_192), CryptoStreamMode.Write)
        Dim sw As New StreamWriter(cs)
        sw.Write(value)
        sw.Flush()
        cs.FlushFinalBlock()
        ms.Flush()
        Return Convert.ToBase64String(ms.GetBuffer(), 0, ms.Length)
    End Function

    Public Shared Function DecryptTripleDES(ByVal value As String) As String
        Dim cryptoProvider As New TripleDESCryptoServiceProvider()
        Dim buffer As Byte() = Convert.FromBase64String(value)
        Dim ms As New MemoryStream(buffer)
        Dim cs As New CryptoStream(ms, cryptoProvider.CreateDecryptor(KEY_192, IV_192), CryptoStreamMode.Read)
        Dim sr As New StreamReader(cs)
        Return sr.ReadToEnd()
    End Function
End Class
#End Region