VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_Url"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Host As String                        'www01.test.schmidks.de
Private m_Domain As String                      'schmidks.de
Private m_Subdomain As String                   'www01.test
Private m_Path As String                        '/mpad/m-pad.asp
Private m_Page As String                        'm-pad.asp
Private m_PagePath As String                    'm-pad.asp/mpad.asp
Private m_Port As String                        '1080
Private m_Scheme As String                      'http:
Private m_Fragment As String                    '#Fragment
Private m_Params As String                      ';par1=xxx
Private m_Login As String                       'proxy login user/pw
Private m_QueryString As String                 'f1=v1&f2=v2
Private m_FieldNames As Collection
Private m_FieldValues As Collection

Public Property Get Href() As String
    Dim s$, s1$
    
    'assemble
    s = PathUrl
    s = s & Page
    If Len(m_Params) > 0 Then
       s = s & ";" & m_Params
    End If
    s1 = QueryString
    If Len(s1) > 0 Then
       s = s & "?" & s1
    End If
    If Len(m_Fragment) > 0 Then
       s = s & "#" & m_Fragment
    End If
    
    'return the result
    Href = s
End Property

Public Property Let Href(ByVal s As String)
Dim p%, I%
Dim oRegExp As New RegExp
Dim oMatchCol As MatchCollection
Dim oMatch As Match
Dim strSubMatch As String
Dim strUrlPattern$
Dim strPagePath As String

On Error GoTo Trap

'init
m_Domain = ""
m_Path = ""
m_Subdomain = ""
m_Scheme = ""
m_Login = ""
m_Params = ""
m_Page = ""
m_PagePath = ""
m_Scheme = ""
m_Port = ""
m_Fragment = ""
Set m_FieldNames = New Collection
Set m_FieldValues = New Collection

'build the regular expression splitting the url into its parts
strUrlPattern = "(.*?:)//" 'scheme
strUrlPattern = strUrlPattern & "(?:([^.?]+)@)?"  'login
strUrlPattern = strUrlPattern & "([^:?#/]*)"  'host
strUrlPattern = strUrlPattern & "(?:\:(\d*))?"  'port
strUrlPattern = strUrlPattern & "(/[^;?#]*)?"  'page path
strUrlPattern = strUrlPattern & "(?:;([^?#]*))?"  'parameter
strUrlPattern = strUrlPattern & "(?:\?([^#]*))?"  'query string
strUrlPattern = strUrlPattern & "(?:\#(.*))?"  'fragment

'split
With oRegExp
'   .Global = True
   .IgnoreCase = True
   .MultiLine = False
   .Pattern = strUrlPattern
   Set oMatchCol = .Execute(s)
End With

'assign parts
Set oMatch = oMatchCol(0)
For I = 0 To oMatch.SubMatches.Count - 1
   strSubMatch = oMatch.SubMatches(I)
   If Len(strSubMatch) > 0 Then
      Select Case I
         Case 0:  m_Scheme = strSubMatch
         Case 1:  m_Login = strSubMatch
         Case 2:  m_Host = strSubMatch
         Case 3:  m_Port = strSubMatch
         Case 4:  m_PagePath = strSubMatch
         Case 5:  m_Params = strSubMatch
         Case 6:  Query = strSubMatch
         Case 7:  m_Fragment = strSubMatch
      End Select
   End If
Next

'split the hostname
If m_Host Like "#*#" Then
   'host is an IP address: nothing to do
ElseIf LCase(m_Host) = "localhost" Then
   'nothing to do
Else
   'split into subdomain and domain
   oRegExp.Pattern = "((?:[^:?#/]*\.)*)([^.]+\.[^.]+)"
   Set oMatchCol = oRegExp.Execute(m_Host)
   'assign parts
   Set oMatch = oMatchCol(0)
   m_Subdomain = oMatch.SubMatches(0)
   If m_Subdomain Like "*." Then
      m_Subdomain = Left$(m_Subdomain, Len(m_Subdomain) - 1)
   End If
   m_Domain = oMatch.SubMatches(1)
End If

'Didn't find a regexp that also extracts the page
'there I need to do it in two steps
If Len(m_PagePath) > 0 Then
   p = InStrRev(m_PagePath, "/")
   If InStr(p, m_PagePath, ".") > 0 Then
      m_Page = Mid$(m_PagePath, p + 1)
      m_Path = Left(m_PagePath, p)
   Else
      m_Path = m_PagePath
      m_Page = ""
   End If
End If
Exit Property

Trap:
Err.Description = "Invalid URL:" & vbCrLf & s
Err.Raise 5 'ERR_INVALIDURL
End Property

Public Property Get PathUrl() As String
'the URL from beginnung up to the path
Dim s$

s = HostUrl
s = s & Path
PathUrl = s
End Property

Public Property Get PageUrl() As String
'the URL from beginnung up to the page
Dim s$

s = s & PathUrl
PageUrl = s & Page
End Property

Public Property Get AfterPage() As String
'all the stuff following the page
Dim s$, s1$

'assemble
If Len(m_Params) > 0 Then
   s = s & ";" & m_Params
End If
s1 = QueryString
If Len(s1) > 0 Then
   s = s & "?" & s1
End If
If Len(m_Fragment) > 0 Then
   s = s & "#" & m_Fragment
End If

'return the result
AfterPage = s
End Property

Public Property Get Parameters() As String
Parameters = m_Params
End Property

Public Property Get Port() As String
Port = m_Port
End Property

Public Property Get Host() As String
Dim s$
s = m_Subdomain
If Len(s) > 0 Then s = s & "."
s = s & m_Domain
If Len(s) > 0 Then
   Host = s
Else
   Host = m_Host
End If
End Property

Public Property Get HostUrl() As String
Dim s$

s = m_Scheme
If Len(s) > 0 Then s = s & "//"
If Len(m_Login) > 0 Then
   s = s & m_Login & "@"
End If
s = s & Host
If Len(m_Port) > 0 Then
   s = s & ":" & m_Port
End If
HostUrl = s
End Property

Public Property Get Domain() As String
Domain = m_Domain
End Property

Public Property Let Domain(NewValue As String)
m_Domain = NewValue
End Property

Public Property Get Subdomain() As String
Subdomain = m_Subdomain
End Property

Public Property Let Subdomain(NewValue As String)
m_Subdomain = NewValue
End Property

Public Property Get Path() As String
Path = m_Path
End Property

Public Property Get Scheme() As String
Scheme = m_Scheme
End Property

Public Property Let Scheme(ByVal s As String)
m_Scheme = s
End Property

Public Property Get QueryString() As String
Dim s$, I%

If m_FieldNames Is Nothing Then Exit Property

For I = 1 To m_FieldNames.Count
   If I > 1 Then s = s & "&"
   s = s & m_FieldNames(I)
   If IsEmpty(m_FieldValues(I)) Then
      'don't append a "="
   Else
      'reduce space to empty string
      s = s & "=" & Trim(m_FieldValues(I))
   End If
Next
'return
QueryString = s
End Property

Public Property Let Query(ByVal s As String)
Dim aPairs As Variant
Dim vPair                     'name/value pair as string
Dim aNameValue As Variant     'name/value pair as array
Dim strName$                  'field name
Dim strValue$                 'field value

If m_FieldNames Is Nothing Then Exit Property

aPairs = Split(s, "&")
'split name/value pairs
For Each vPair In aPairs
   If Len(vPair) > 0 Then
      aNameValue = Split(vPair, "=")
      m_FieldNames.Add aNameValue(0)
      If UBound(aNameValue) > 0 Then
         m_FieldValues.Add aNameValue(1)
      Else
         If InStr(vPair, "=") > 0 Then
            m_FieldValues.Add " "
         Else
            m_FieldValues.Add Empty
         End If
      End If
   End If
Next

End Property

Public Property Get ParamNames() As Variant
Dim aParams() As String
Dim I%

If m_FieldNames Is Nothing Then
   ParamNames = Empty
   Exit Property
End If
If m_FieldNames.Count <= 0 Then
   ParamNames = Empty
   Exit Property
End If

ReDim aParams(m_FieldNames.Count - 1)
For I = 1 To m_FieldNames.Count
   aParams(I - 1) = m_FieldNames(I)
Next
ParamNames = aParams
End Property

Public Property Get ParamValues() As Variant
Dim aParams() As String
Dim I%

If m_FieldNames Is Nothing Then
   ParamValues = Empty
   Exit Property
End If
If m_FieldNames.Count <= 0 Then
   ParamValues = Empty
   Exit Property
End If

ReDim aParams(m_FieldValues.Count - 1)
For I = 1 To m_FieldValues.Count
   aParams(I - 1) = m_FieldValues(I)
Next
ParamValues = aParams
End Property

Public Function PutParams(Names As Variant, ByVal Values As Variant)
Dim I%

'check
If Not IsArray(Names) Or Not IsArray(Values) Then
   Err.Description = "You must pass arrays as parameters"
   Err.Raise 438
End If
If UBound(Names) <> UBound(Values) Or LBound(Names) <> LBound(Values) Then
   Err.Description = "Array must have the same lbound and ubound"
   Err.Raise 438
End If

Set m_FieldNames = New Collection
Set m_FieldValues = New Collection

'copy
For I = LBound(Names) To UBound(Names)
   m_FieldNames.Add Names(I)
   m_FieldValues.Add Values(I)
Next

End Function

Public Function GetParam(ByVal Name As String) As Variant
Dim Cnt%, I%, intFound%
Dim aParam As Variant
Dim aResult()

If m_FieldNames Is Nothing Then
   GetParam = ""
   Exit Function
End If

'count the number of params with this name
For I = 1 To m_FieldNames.Count
   If m_FieldNames(I) = Name Then
      Cnt = Cnt + 1
      intFound = I
   End If
Next
If Cnt > 1 Then
   'build the array
   ReDim aResult(Cnt - 1)
   Cnt = 0
   For Each aParam In m_FieldNames
      If aParam = Name Then
         aResult(Cnt) = m_FieldValues(Cnt + 1)
         Cnt = Cnt + 1
      End If
   Next
   'return an array
   GetParam = aResult
Else
   'return the string
   GetParam = m_FieldValues(intFound)
End If
End Function

Public Function PutParam(ByVal Name As String, ByVal NewValue As Variant, Optional ByVal Replace As Boolean = True)
'NewValue may be a string or an array
Dim vValue As Variant
Dim I%

If m_FieldNames Is Nothing Then Exit Function

If Replace Then
   'delete all entries for this name
   For I = m_FieldNames.Count To 1 Step -1
      If m_FieldNames(I) = Name Then
         m_FieldNames.Remove (I)
      End If
   Next
End If
'add new entries
If IsArray(NewValue) Then
   For Each vValue In NewValue
      m_FieldNames.Add Name
      m_FieldValues.Add vValue
   Next
Else
   m_FieldNames.Add Name
   m_FieldValues.Add NewValue
End If
End Function

Public Property Get Login() As String
Login = m_Login
End Property

Public Property Let Login(ByVal s As String)
m_Login = s
End Property

Public Property Get Page() As String
Page = m_Page
End Property

Public Property Get PagePath() As String
PagePath = Path & m_Page
End Property

Public Property Let Page(ByVal s As String)
m_Page = s
End Property

Public Property Get Fragment() As String
Fragment = m_Fragment
End Property

Public Property Let Fragment(ByVal s As String)
m_Fragment = s
End Property

