Imports Newtonsoft.Json
Imports System
Imports System.CodeDom
Imports System.CodeDom.Compiler
Imports System.Collections
Imports System.Collections.Generic
Imports System.Linq
Imports System.Reflection
Imports System.Text
Imports System.Threading.Tasks
Imports System.Web.Services.Description
Public Class WebServiceObject
    Implements IDisposable

    Private webServiceAsmxUrl As String
    Private timeout As Integer = 0
    Private serviceName As String = Nothing
    Private wsvcClass As Object
    Private ProtocolName As String = "Soap12"
    Public Sub New(ByVal url As String, ByVal timeout As Integer)
        webServiceAsmxUrl = url
        Me.timeout = timeout
        If Not Create() Then
            Throw New Exception("物件建立失敗")
        End If
    End Sub
    Public Sub New(ByVal url As String)
        webServiceAsmxUrl = url
        If Not Create() Then
            Throw New Exception("物件建立失敗")
        End If
    End Sub
    Public Sub New(ByVal url As String, ByVal timeout As Integer, ByVal serviceName As String)
        webServiceAsmxUrl = url
        Me.timeout = timeout
        Me.serviceName = serviceName
        If Not Create() Then
            Throw New Exception("物件建立失敗")
        End If
    End Sub
    Public Sub New(ByVal url As String, ByVal serviceName As String)
        webServiceAsmxUrl = url
        Me.serviceName = serviceName
        If Not Create() Then
            Throw New Exception("物件建立失敗")
        End If
    End Sub
    Private Function Create() As Boolean

        Dim client As System.Net.WebClient = New System.Net.WebClient()
        'Connect To the web service
        Dim stream As System.IO.Stream = client.OpenRead(webServiceAsmxUrl & "?wsdl")
        'Read the WSDL file describing a service.
        Dim description As ServiceDescription = ServiceDescription.Read(stream)

        If String.IsNullOrEmpty(serviceName) Then
            If description.Services.Count > 0 Then
                serviceName = description.Services(0).Name
            End If
        End If
        'Load the DOM
        '--Initialize a service description importer.
        Dim importer As ServiceDescriptionImporter = New ServiceDescriptionImporter()
        'Use SOAP 1.2. (有些Java的WebService不Support 1.2，請設成空字串
        If String.IsNullOrEmpty(Me.ProtocolName) Then
            importer.ProtocolName = Me.ProtocolName
        End If
        importer.AddServiceDescription(description, Nothing, Nothing)

        ' --Generate a proxy client. 
        importer.Style = ServiceDescriptionImportStyle.Client
        'Generate properties to represent primitive values.
        importer.CodeGenerationOptions = System.Xml.Serialization.CodeGenerationOptions.GenerateProperties
        ' Initialize a Code-DOM tree into which we will import the service.
        Dim codenamespace As CodeNamespace = New CodeNamespace()

        Dim codeunit As CodeCompileUnit = New CodeCompileUnit()

        codeunit.Namespaces.Add(codenamespace)

        'Import the service into the Code-DOM tree. 
        'This creates proxy code that uses the service.
        Dim warning As ServiceDescriptionImportWarnings = importer.Import(codenamespace, codeunit)
        If (warning = 0) Then

            'Generate the proxy code
            'CodeDomProvider provider = CodeDomProvider.CreateProvider("CSharp");
            Dim provider As CodeDomProvider = CodeDomProvider.CreateProvider("VisualBasic")

            'Compile the assembly proxy with the 
            ' appropriate references
            Dim assemblyReferences() As String = {"System.dll",
                                   "System.Web.Services.dll",
                                   "System.Web.dll",
                                   "System.Xml.dll",
                                   "System.Data.dll"}
            'Add parameters
            Dim parms As CompilerParameters = New CompilerParameters(assemblyReferences)
            parms.GenerateInMemory = True
            Dim results As CompilerResults = provider.CompileAssemblyFromDom(parms, codeunit)
            If results.Errors.Count > 0 Then
                For Each oops As CompilerError In results.Errors
                    System.Diagnostics.Debug.WriteLine("========Compiler error============")
                    System.Diagnostics.Debug.WriteLine(oops.ErrorText)
                Next
                Throw New Exception("Compile Error Occured calling WebService.")
                Return Nothing
            End If
            'Finally, Invoke the web service method
            wsvcClass = results.CompiledAssembly.CreateInstance(serviceName)
            If timeout > 0 Then
                wsvcClass.timeout = timeout
            End If
        Else
            Return Nothing
        End If

        Return True

    End Function
    Public Function Execute(ByVal methodName As String) As Object
        Return Me.Execute(methodName, Nothing)
    End Function
    Public Function Execute(ByVal methodName As String, args() As Object) As Object
        Dim mi As MethodInfo = wsvcClass.GetType().GetMethod(methodName)
        If mi Is Nothing Then
            Throw New Exception(String.Format("There doesn't have the {0} method in WebService", methodName))
        End If

        '這裡取出 WebService 的參數
        '判斷如果是物件或是Array的話， 就透過JSON.NET來轉成WS要的物件

        Try
            Dim pInfos() As ParameterInfo = mi.GetParameters()

            Dim newArgs As ArrayList = New ArrayList()
            Dim i = 0
            For Each p As ParameterInfo In pInfos
                Dim pType As Type = p.ParameterType
                If (pType.IsPrimitive) OrElse (pType = GetType(String)) Then
                    newArgs.Add(args(i))
                Else
                    Dim argObj = JsonConvert.DeserializeObject(args(i).ToString, pType)
                    newArgs.Add(argObj)
                End If
            Next
            Return mi.Invoke(wsvcClass, newArgs.ToArray())
        Catch ex As Exception
            Throw New Exception(ex.ToString)
        End Try


    End Function
    
#Region "IDisposable Support"
    Private disposedValue As Boolean ' 偵測多餘的呼叫

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: 處置 Managed 狀態 (Managed 物件)。
                webServiceAsmxUrl = Nothing
                timeout = 0
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
