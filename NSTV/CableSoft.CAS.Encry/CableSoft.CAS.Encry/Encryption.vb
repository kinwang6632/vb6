Imports System.Security.Cryptography
Imports System.Text
Imports System.Reflection

Public Class Encryption
    Public Shared Function DescryptDES(ByVal byteData() As Byte, ByVal btyKey() As Byte) As Byte()
        'Dim provider As DESCryptoServiceProvider = New DESCryptoServiceProvider()
        'Dim transform As ICryptoTransform = Nothing
        'Dim buffer() As Byte = byteData
        'Dim bty() As Byte
        'Try
        '    provider.Key = Encoding.ASCII.GetBytes(strKey)
        '    provider.Mode = CipherMode.ECB
        '    provider.Padding = PaddingMode.Zeros
        '    transform = provider.CreateDecryptor
        '    bty = transform.TransformFinalBlock(buffer, 0, buffer.Length)
        'Catch ex As Exception
        '    Throw ex
        'Finally
        '    provider.Dispose()
        '    If transform IsNot Nothing Then
        '        transform.Dispose()
        '    End If

        'End Try

        'Return bty
        Dim provider As DESCryptoServiceProvider = New DESCryptoServiceProvider()

        Try


            Dim transform As ICryptoTransform = Nothing

            Dim bty() As Byte

            Try
                'Dim Proverider As TripleDESCryptoServiceProvider = New TripleDESCryptoServiceProvider()
                provider.Padding = PaddingMode.None
                Dim t As Type = Type.GetType("System.Security.Cryptography.CryptoAPITransformMode")
                Dim objDecrypt As Object    'Encrypt 
                objDecrypt = t.GetField("Decrypt", BindingFlags.Instance Or BindingFlags.Static Or BindingFlags.Public Or _
                                   BindingFlags.NonPublic Or BindingFlags.DeclaredOnly).GetValue(t)
                'Dim param() As Object = {Encoding.ASCII.GetBytes(strKey), CipherMode.ECB, Nothing, provider.FeedbackSize, objDecrypt}
                Dim param() As Object = {btyKey, CipherMode.ECB, Nothing, provider.FeedbackSize, objDecrypt}
                Dim mi As MethodInfo
                mi = provider.GetType().GetMethod("_NewEncryptor", BindingFlags.Instance Or BindingFlags.NonPublic)
                Dim desCrypt As ICryptoTransform = mi.Invoke(provider, param)
                bty = desCrypt.TransformFinalBlock(byteData, 0, byteData.Length)
                Return bty

            Catch ex As Exception
                Throw ex
            Finally
                provider.Dispose()
                If transform IsNot Nothing Then
                    transform.Dispose()
                End If

            End Try


        Catch ex As Exception
            Throw ex
        Finally
            provider.Dispose()
        End Try

    End Function
    Public Shared Function EncryptDES(ByVal Plain() As Byte, ByVal Key() As Byte) As Byte()
        Dim provider As DESCryptoServiceProvider = New DESCryptoServiceProvider()

        Try


            Dim transform As ICryptoTransform = Nothing

            Dim bty() As Byte

            Try
                'Dim Proverider As TripleDESCryptoServiceProvider = New TripleDESCryptoServiceProvider()
                provider.Padding = PaddingMode.None
                Dim t As Type = Type.GetType("System.Security.Cryptography.CryptoAPITransformMode")
                Dim objDecrypt As Object    'Encrypt 
                objDecrypt = t.GetField("Encrypt", BindingFlags.Instance Or BindingFlags.Static Or BindingFlags.Public Or _
                                   BindingFlags.NonPublic Or BindingFlags.DeclaredOnly).GetValue(t)
                'Dim param() As Object = {Encoding.ASCII.GetBytes(Key), CipherMode.ECB, Nothing, provider.FeedbackSize, objDecrypt}
                Dim param() As Object = {Key, CipherMode.ECB, Nothing, provider.FeedbackSize, objDecrypt}
                Dim mi As MethodInfo
                mi = provider.GetType().GetMethod("_NewEncryptor", BindingFlags.Instance Or BindingFlags.NonPublic)
                Dim desCrypt As ICryptoTransform = mi.Invoke(provider, param)
                bty = desCrypt.TransformFinalBlock(Plain, 0, Plain.Length)
                Return bty

            Catch ex As Exception
                Throw ex
            Finally
                provider.Dispose()
                If transform IsNot Nothing Then
                    transform.Dispose()
                End If

            End Try


        Catch ex As Exception
            Throw ex
        Finally
            provider.Dispose()
        End Try

    End Function
    Public Shared Function Decrypt3DES(ByVal byteData() As Byte, ByVal strKey As String) As Byte()
        Dim provider As TripleDESCryptoServiceProvider = New TripleDESCryptoServiceProvider()
        Dim transform As ICryptoTransform = Nothing
        Dim buffer() As Byte = byteData
        Dim bty() As Byte

        Try
            'Dim Proverider As TripleDESCryptoServiceProvider = New TripleDESCryptoServiceProvider()
            provider.Padding = PaddingMode.None
            Dim t As Type = Type.GetType("System.Security.Cryptography.CryptoAPITransformMode")
            Dim objDecrypt As Object    'Encrypt 
            objDecrypt = t.GetField("Decrypt", BindingFlags.Instance Or BindingFlags.Static Or BindingFlags.Public Or _
                               BindingFlags.NonPublic Or BindingFlags.DeclaredOnly).GetValue(t)
            Dim param() As Object = {Encoding.ASCII.GetBytes(strKey), CipherMode.ECB, Nothing, provider.FeedbackSize, objDecrypt}
            'Dim param() As Object = {strKey, CipherMode.ECB, Nothing, provider.FeedbackSize, objDecrypt}
            Dim mi As MethodInfo
            mi = provider.GetType().GetMethod("_NewEncryptor", BindingFlags.Instance Or BindingFlags.NonPublic)
            Dim desCrypt As ICryptoTransform = mi.Invoke(provider, param)
            bty = desCrypt.TransformFinalBlock(buffer, 0, buffer.Length)
            Return bty

        Catch ex As Exception
            Throw ex
        Finally
            provider.Dispose()
            If transform IsNot Nothing Then
                transform.Dispose()
            End If

        End Try

        Return bty
    End Function
    Public Shared Function Decrypt3DES(ByVal byteData As String, ByVal strKey As String) As String
        'Dim bty() As Byte = Decrypt3DES(Convert.FromBase64String(byteData), strKey)
        Dim bty() As Byte = Decrypt3DES(Convert.FromBase64String(byteData), strKey)
        'Return Encoding.UTF8.GetString(Convert.FromBase64String(Encoding.UTF8.GetString(bty)))
        'bty = bty.TakeWhile(Function(chkData As Byte)
        '                        Return chkData <> 0
        '                    End Function).ToArray
        Dim SkipCount As Integer = _
            bty.Reverse.Take(8).Count(Function(byteCheck As Byte)
                                          Return byteCheck = 0
                                      End Function)
       



       Return  Encoding.UTF8.GetString(bty.Reverse.ToArray.Skip(SkipCount).Reverse.ToArray())




        'Return Encoding.UTF8.GetString(bty)

        'Return Encoding.UTF8.GetString(bty.TakeWhile(Function(chkData As Byte)
        '                                                 Return chkData <> 0
        '                                             End Function).ToArray)

    'Dim bty() As Byte = Encrypt3DES(Encoding.ASCII.GetBytes(Convert.ToBase64String(Encoding.UTF8.GetBytes(Plain))),
    '                              Encoding.ASCII.GetBytes(Key))
    'Return Convert.ToBase64String(bty)
    End Function
    Public Shared Function Encrypt3DES(ByVal Plain() As Byte, ByVal Key() As Byte) As Byte()
        Dim provider As TripleDESCryptoServiceProvider = New TripleDESCryptoServiceProvider()

        Try
            Dim transform As ICryptoTransform = Nothing
            Dim buffer() As Byte = Plain
            Dim bty() As Byte
            Try
                'Dim Proverider As TripleDESCryptoServiceProvider = New TripleDESCryptoServiceProvider()
                provider.Padding = PaddingMode.None
                Dim t As Type = Type.GetType("System.Security.Cryptography.CryptoAPITransformMode")
                Dim objDecrypt As Object    'Encrypt 
                objDecrypt = t.GetField("Encrypt", BindingFlags.Instance Or BindingFlags.Static Or BindingFlags.Public Or _
                                   BindingFlags.NonPublic Or BindingFlags.DeclaredOnly).GetValue(t)
                Dim param() As Object = {Key, CipherMode.ECB, Nothing, provider.FeedbackSize, objDecrypt}
                Dim mi As MethodInfo
                mi = provider.GetType().GetMethod("_NewEncryptor", BindingFlags.Instance Or BindingFlags.NonPublic)
                Dim desCrypt As ICryptoTransform = mi.Invoke(provider, param)
                bty = desCrypt.TransformFinalBlock(buffer, 0, buffer.Length)
                Return bty

            Catch ex As Exception
                Throw ex
            Finally
                provider.Dispose()
                If transform IsNot Nothing Then
                    transform.Dispose()
                End If

            End Try

            Return bty
        Catch ex As Exception
            Throw ex
        Finally
            provider.Dispose()
        End Try

    End Function
    Public Shared Function Encrypt3DES(ByVal Plain() As Byte, ByVal Key As String) As Byte()
        Return Encrypt3DES(Plain, Encoding.ASCII.GetBytes(Key))
    End Function
    Public Shared Function Encrypt3DES(ByVal Plain As String, ByVal Key As String) As String
        Dim btyPlain() As Byte = Encoding.UTF8.GetBytes(Plain)
        Dim FillCount As Byte = 8 - (btyPlain.Length Mod 8)
        If FillCount <> 8 Then
            Dim lstBty As New List(Of Byte)
            lstBty = btyPlain.ToList
            For i As Int32 = 1 To FillCount
                lstBty.Add(0)
            Next
            btyPlain = lstBty.ToArray        
        End If
        'Dim bty() As Byte = Encrypt3DES(Encoding.ASCII.GetBytes(Convert.ToBase64String(Encoding.UTF8.GetBytes(Plain))),
        '                                Encoding.ASCII.GetBytes(Key))
        Dim bty() As Byte = Encrypt3DES(btyPlain,
                                        Encoding.ASCII.GetBytes(Key))
        Return Convert.ToBase64String(bty)


        'Dim FillCount As Byte = byteData.Length Mod 8
        'If FillCount > 0 Then
        '    Dim byteFill(FillCount) As Byte
        '    For i As Byte = 0 To byteFill.Length - 1
        '        byteFill(i) = &HF0
        '    Next
        '    byteData.ToList.AddRange(byteFill)
        'End If

    End Function
    Public Shared Function EncryMD5(ByVal byteData() As Byte) As Byte()
        Try
            Dim submd5 As New Security.Cryptography.MD5CryptoServiceProvider
            Dim byteResult() As Byte = submd5.ComputeHash(byteData)
            Return byteResult
        Catch ex As Exception
            Throw
        End Try

    End Function
    Public Shared Function EncryMD5(ByVal byteData As String) As String
        Dim bty() As Byte = EncryMD5(Convert.FromBase64String(byteData))
        Return Convert.ToBase64String(bty)
    End Function
End Class
