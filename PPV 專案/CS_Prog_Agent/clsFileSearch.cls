VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFileSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' PR : Hammer
' Start date : 前兩年吧 ! 忘了
' Last Modify : 誰曉得 ! 天知道

Option Explicit

Option Compare Text

Private Const MAX_PATH = 512
Private Const PATH_SEP = "\"

Private Const INVALID_HANDLE_VALUE = -1

Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_COMPRESSED = &H800
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type

Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
    
Private Declare Function FileTimeToDosDateTime Lib "kernel32" (lpFileTime As FILETIME, ByVal lpFatDate As Long, ByVal lpFatTime As Long) As Long

Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
    
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long

Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long

Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
  
Private Declare Sub GetLocalTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)

Public FilesFound       As Long
Public PathsFound       As Long

Private Files()         As String
Private Paths()         As String
Private FilePatterns()  As String
Private PathPatterns()  As String
Private uboundFilePatterns As Long
Private uboundPathPatterns As Long

Private StrName         As String

Public Sub Class_Initialize()
    FilesFound = 0
    PathsFound = 0
    uboundFilePatterns = -1
    uboundPathPatterns = -1
End Sub

Public Sub Class_Terminate()
    ResetSearch True
End Sub

Property Get File(index As Long) As String
    File = Files(index)
End Property

Property Get Path(index As Long) As String
    Path = Paths(index)
End Property

Property Get FileSize(index As Long) As Long
    On Error Resume Next
    FileSize = FileLen(Files(index))
    If Err.Number Then FileSize = 0: Err.Clear
End Property

Public Function CopyFiles() As Variant
    CopyFiles = Files()
End Function

Public Function CopyPaths() As Variant
    CopyPaths = Paths()
End Function

Private Sub StripNulls(s As String)
    If (InStr(s, vbNullChar)) Then
        StrName = Left$(s, InStr(s, vbNullChar) - 1)
    Else
        StrName = s
    End If
End Sub

Public Function FindFiles(ByVal sPath As String, searchStr As String, Recurse As Boolean)

    Dim WFD       As WIN32_FIND_DATA
    Dim fDirs()   As String
    Dim hSearch   As Long
    Dim cont      As Integer
    Dim i%, j     As Integer

    If Right$(sPath, 1) <> PATH_SEP Then sPath = sPath & PATH_SEP

    hSearch = FindFirstFile(sPath & "*", WFD)

    If (hSearch <> INVALID_HANDLE_VALUE) Then

        cont = 1
        While cont

            StripNulls WFD.cFileName

            If Len(StrName) Then
            If ((StrName <> ".") And (StrName <> "..")) Then

                If ((GetFileAttributes(sPath & StrName) And _
                    FILE_ATTRIBUTE_DIRECTORY) <> FILE_ATTRIBUTE_DIRECTORY) Then

                    If (StrName Like searchStr) Then

                        ReDim Preserve Files(FilesFound)
                        Files(FilesFound) = sPath & StrName
                        FilesFound = FilesFound + 1

                    End If

                ElseIf Recurse Then

                    i = i + 1
                    ReDim Preserve fDirs(i)
                    fDirs(i) = StrName

                End If

            End If
            End If

        cont = FindNextFile(hSearch, WFD) 'naeste fil
        Wend

    cont = FindClose(hSearch)
    End If

    If Recurse Then
        For j = 1 To i
            FindFiles sPath & fDirs(j), searchStr, Recurse
        Next
    End If

    On Error Resume Next
    Erase fDirs
End Function

Public Sub FindFilesWithPatterns(ByVal sPath As String, Recurse As Boolean)

    Dim WFD             As WIN32_FIND_DATA
    Dim fDirs()         As String
    Dim hSearch         As Long
    Dim cont            As Integer
    Dim i%, j%, count   As Integer

    If Right$(sPath, 1) <> PATH_SEP Then sPath = sPath & PATH_SEP

    hSearch = FindFirstFile(sPath & "*", WFD)

    If (hSearch <> INVALID_HANDLE_VALUE) Then

        cont = 1
        While cont

            StripNulls WFD.cFileName

            If Len(StrName) Then
            If ((StrName <> ".") And (StrName <> "..")) Then

                If ((GetFileAttributes(sPath & StrName) And _
                    FILE_ATTRIBUTE_DIRECTORY) <> FILE_ATTRIBUTE_DIRECTORY) Then

                    For j = 0 To uboundFilePatterns Step 1
                    If (StrName Like FilePatterns(j)) Then

                        ReDim Preserve Files(FilesFound)
                        Files(FilesFound) = sPath & StrName
                        FilesFound = FilesFound + 1
                        Exit For

                    End If
                    Next

                ElseIf Recurse Then

                    count = count + 1
                    ReDim Preserve fDirs(count)
                    fDirs(count) = StrName

                End If

            End If
            End If

        cont = FindNextFile(hSearch, WFD) 'naeste fil
        Wend

    cont = FindClose(hSearch)
    End If

    If Recurse Then
        For i = 1 To count
            FindFilesWithPatterns sPath & fDirs(i), True
        Next
    End If
    
    On Error Resume Next
    Erase fDirs

End Sub

Public Sub FindPaths(ByVal sPath As String, Recurse As Boolean)

    Dim WFD             As WIN32_FIND_DATA
    Dim fDirs()         As String
    Dim hSearch         As Long
    Dim i%, count       As Integer
    Dim cont            As Integer

    If Right$(sPath, 1) <> PATH_SEP Then sPath = sPath & PATH_SEP

    hSearch = FindFirstFile(sPath & "*", WFD)

    If (hSearch <> INVALID_HANDLE_VALUE) Then

        cont = 1
        While cont

            StripNulls WFD.cFileName

            If (StrName <> "." And StrName <> "..") Then
            If ((GetFileAttributes(sPath & StrName) And _
                 FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY) Then

                count = count + 1
                ReDim Preserve fDirs(count)
                fDirs(count) = StrName

            End If
            End If

        cont = FindNextFile(hSearch, WFD)
        Wend

    cont = FindClose(hSearch)
    End If

    For i = 1 To count
    
        ReDim Preserve Paths(PathsFound)
        Paths(PathsFound) = sPath & fDirs(i)
        PathsFound = PathsFound + 1

        If Recurse Then
            FindPaths sPath & fDirs(i), True
        End If

    Next
    
    On Error Resume Next
    Erase fDirs
    
End Sub

Public Sub FindPathsWithPatterns(ByVal sPath As String, Recurse As Boolean)

    Dim WFD             As WIN32_FIND_DATA
    Dim fDirs()         As String
    Dim temp            As String
    Dim hSearch         As Long
    Dim i%, j%, count   As Integer
    Dim cont            As Integer

    If Right$(sPath, 1) <> PATH_SEP Then sPath = sPath & PATH_SEP

    hSearch = FindFirstFile(sPath & "*", WFD)

    If (hSearch <> INVALID_HANDLE_VALUE) Then

        cont = 1
        While cont

        StripNulls WFD.cFileName

        If (StrName <> ".") And (StrName <> "..") Then
        If ((GetFileAttributes(sPath & StrName) And FILE_ATTRIBUTE_DIRECTORY) = _
            FILE_ATTRIBUTE_DIRECTORY) Then
        
            count = count + 1
            ReDim Preserve fDirs(count)
            fDirs(count) = StrName
        
        End If
        End If

        cont = FindNextFile(hSearch, WFD)
        Wend

    cont = FindClose(hSearch)
    End If

    For i = 1 To count

        temp = sPath & fDirs(i)

        For j = 0 To uboundPathPatterns Step 1
        If temp Like PathPatterns(j) Then

            ReDim Preserve Paths(FilesFound)
            Paths(FilesFound) = temp
            PathsFound = PathsFound + 1

        End If
        Next j

        If Recurse Then
           FindPathsWithPatterns temp, Recurse
        End If

    Next i

    On Error Resume Next
    Erase fDirs

End Sub

Public Sub SetFilePatterns(s As String)
    FilePatterns = Split(s, ";")
    uboundFilePatterns = UBound(FilePatterns)
End Sub

Public Function GetFilePatterns(s As String) As String
    GetFilePatterns = Join(FilePatterns(), ";")
End Function

Public Sub SetPathPatterns(s As String)
    PathPatterns = Split(s, ";")
    uboundPathPatterns = UBound(PathPatterns)
End Sub

Public Function GetPathPatterns(s As String) As String
    GetPathPatterns = Join(PathPatterns(), ";")
End Function

Public Sub ResetFileSearch(Optional erasePatterns As Boolean = False)
    FilesFound = 0
    Erase Files
    If erasePatterns Then
        Erase FilePatterns
        uboundFilePatterns = -1
    End If
End Sub

Public Sub ResetPathSearch(Optional erasePatterns As Boolean = False)
    PathsFound = 0
    Erase Paths
    If erasePatterns Then
        Erase PathPatterns
        uboundPathPatterns = -1
    End If
End Sub

Public Sub ResetSearch(Optional erasePatterns As Boolean = False)
    ResetFileSearch erasePatterns
    ResetPathSearch erasePatterns
End Sub

