Attribute VB_Name = "mod_ComDlg"
' PR : Hammer
' Start date : 粉久以前
' Last Modify : 不可考

Option Explicit

'//
'// Structures
'//

Private Type OPENFILENAME
    lStructSize As Long
    hwnd As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Type COLORSTRUC
    lStructSize As Long
    hwnd As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As String
    Flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Const LF_FACESIZE = 32

Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(LF_FACESIZE) As Byte
End Type

Private Type FONTSTRUC
    lStructSize As Long
    hwnd As Long
    hDC As Long
    lpLogFont As Long
    iPointSize As Long
    Flags As Long
    rgbColors As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
    hInstance As Long
    lpszStyle As String
    nFontType As Integer
    MISSING_ALIGNMENT As Integer
    nSizeMin As Long
    nSizeMax As Long
End Type

Private Type DEVMODE
    dmDeviceName As String * 32
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * 32
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFreq As Long
End Type

Private Type PRINTDLGSTRUC
    lStructSize As Long
    hwnd As Long
    hDevMode As Long
    hDevNames As Long
    hDC As Long
    Flags As Long
    nFromPage As Integer
    nToPage As Integer
    nMinPage As Integer
    nMaxPage As Integer
    nCopies As Integer
    hInstance As Long
    lCustData As Long
    lpfnPrintHook As Long
    lpfnSetupHook As Long
    lpPrintTemplateName As String
    lpSetupTemplateName As String
    hPrintTemplate As Long
    hSetupTemplate As Long
End Type

Public Type PRINTPROPS
    Cancel As Boolean
    Device As String
    Copies As Integer
    FromPage As Integer
    ToPage As Integer
    ToFile As Boolean
    Range As Integer
End Type

Private Type SHITEMID
    cb As Long
    abID As Byte
End Type

Private Type ITEMIDLIST
    mkid As SHITEMID
End Type

Private Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

'//
'// Win32s
'//

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function PrintDlg Lib "comdlg32.dll" Alias "PrintDlgA" (pPrintdlg As PRINTDLGSTRUC) As Long
Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As COLORSTRUC) As Long
Private Declare Function ChooseFont Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As FONTSTRUC) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function ConnectToPrinterDlg Lib "winspool.drv" (ByVal hwnd As Long, ByVal Flags As Long) As Long
Private Declare Function SHGetPathFromIDList Lib "SHELL32.DLL" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "SHELL32.DLL" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Declare Function SHBrowseForFolder Lib "SHELL32.DLL" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long 'ITEMIDLIST
Private Declare Function WriteProfileString Lib "kernel32" Alias "WriteProfileStringA" (ByVal lpszSection As String, ByVal lpszKeyName As String, ByVal lpszString As String) As Long
Private Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
Private Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Declare Function WinHelp Lib "user32" Alias "WinHelpA" (ByVal hwnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Long) As Long


'//
'// Constants (Public for Print Properties Structure)
'//

Public Const ppRangeAll = 0
Public Const ppRangePages = 1
Public Const ppRangeSelection = 2

'//
'// Constants (Public for Print Dialog Box)
'//

Public Const PD_NOSELECTION = &H4
Public Const PD_DISABLEPRINTTOFILE = &H80000
Public Const PD_PRINTTOFILE = &H20
Public Const PD_RETURNDC = &H100
Public Const PD_RETURNDEFAULT = &H400
Public Const PD_RETURNIC = &H200
Public Const PD_SELECTION = &H1
Public Const PD_SHOWHELP = &H800
Public Const PD_NOPAGENUMS = &H8
Public Const PD_PAGENUMS = &H2

'//
'// Constants (Public for WinHelp)
'//

Public Const HELP_COMMAND = &H102&
Public Const HELP_CONTENTS = &H3&
Public Const HELP_CONTEXT = &H1
Public Const HELP_CONTEXTPOPUP = &H8&
Public Const HELP_FORCEFILE = &H9&
Public Const HELP_HELPONHELP = &H4
Public Const HELP_INDEX = &H3
Public Const HELP_KEY = &H101
Public Const HELP_MULTIKEY = &H201&
Public Const HELP_PARTIALKEY = &H105&
Public Const HELP_QUIT = &H2
Public Const HELP_SETCONTENTS = &H5&
Public Const HELP_SETINDEX = &H5
Public Const HELP_SETWINPOS = &H203&


'//
'// Constants (Private)
'//

Private Const FW_BOLD = 700
Private Const GMEM_MOVEABLE = &H2
Private Const GMEM_ZEROINIT = &H40
Private Const GHND = (GMEM_MOVEABLE Or GMEM_ZEROINIT)
Private Const OFN_ALLOWMULTISELECT = &H200
Private Const OFN_CREATEPROMPT = &H2000
Private Const OFN_ENABLEHOOK = &H20
Private Const OFN_ENABLETEMPLATE = &H40
Private Const OFN_ENABLETEMPLATEHANDLE = &H80
Private Const OFN_EXPLORER = &H80000
Private Const OFN_EXTENSIONDIFFERENT = &H400
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_LONGNAMES = &H200000
Private Const OFN_NOCHANGEDIR = &H8
Private Const OFN_NODEREFERENCELINKS = &H100000
Private Const OFN_NOLONGNAMES = &H40000
Private Const OFN_NONETWORKBUTTON = &H20000
Private Const OFN_NOREADONLYRETURN = &H8000
Private Const OFN_NOTESTFILECREATE = &H10000
Private Const OFN_NOVALIDATE = &H100
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_PATHMUSTEXIST = &H800
Private Const OFN_READONLY = &H1
Private Const OFN_SHAREAWARE = &H4000
Private Const OFN_SHAREFALLTHROUGH = 2
Private Const OFN_SHARENOWARN = 1
Private Const OFN_SHAREWARN = 0
Private Const OFN_SHOWHELP = &H10
Private Const PD_ALLPAGES = &H0
Private Const PD_COLLATE = &H10
Private Const PD_ENABLEPRINTHOOK = &H1000
Private Const PD_ENABLEPRINTTEMPLATE = &H4000
Private Const PD_ENABLEPRINTTEMPLATEHANDLE = &H10000
Private Const PD_ENABLESETUPHOOK = &H2000
Private Const PD_ENABLESETUPTEMPLATE = &H8000
Private Const PD_ENABLESETUPTEMPLATEHANDLE = &H20000
Private Const PD_HIDEPRINTTOFILE = &H100000
Private Const PD_NONETWORKBUTTON = &H200000
Private Const PD_PRINTSETUP = &H40
Private Const PD_USEDEVMODECOPIES = &H40000
Private Const PD_USEDEVMODECOPIESANDCOLLATE = &H40000
Private Const PD_NOWARNING = &H80
Private Const CF_ANSIONLY = &H400&
Private Const CF_APPLY = &H200&
Private Const CF_BITMAP = 2
Private Const CF_PRINTERFONTS = &H2
Private Const CF_PRIVATEFIRST = &H200
Private Const CF_PRIVATELAST = &H2FF
Private Const CF_RIFF = 11
Private Const CF_SCALABLEONLY = &H20000
Private Const CF_SCREENFONTS = &H1
Private Const CF_BOTH = (CF_SCREENFONTS Or CF_PRINTERFONTS)
Private Const CF_DIB = 8
Private Const CF_DIF = 5
Private Const CF_DSPBITMAP = &H82
Private Const CF_DSPENHMETAFILE = &H8E
Private Const CF_DSPMETAFILEPICT = &H83
Private Const CF_DSPTEXT = &H81
Private Const CF_EFFECTS = &H100&
Private Const CF_ENABLEHOOK = &H8&
Private Const CF_ENABLETEMPLATE = &H10&
Private Const CF_ENABLETEMPLATEHANDLE = &H20&
Private Const CF_ENHMETAFILE = 14
Private Const CF_FIXEDPITCHONLY = &H4000&
Private Const CF_FORCEFONTEXIST = &H10000
Private Const CF_GDIOBJFIRST = &H300
Private Const CF_GDIOBJLAST = &H3FF
Private Const CF_INITTOLOGFONTSTRUCT = &H40&
Private Const CF_LIMITSIZE = &H2000&
Private Const CF_METAFILEPICT = 3
Private Const CF_NOFACESEL = &H80000
Private Const CF_NOVERTFONTS = &H1000000
Private Const CF_NOVECTORFONTS = &H800&
Private Const CF_NOOEMFONTS = CF_NOVECTORFONTS
Private Const CF_NOSCRIPTSEL = &H800000
Private Const CF_NOSIMULATIONS = &H1000&
Private Const CF_NOSIZESEL = &H200000
Private Const CF_NOSTYLESEL = &H100000
Private Const CF_OEMTEXT = 7
Private Const CF_OWNERDISPLAY = &H80
Private Const CF_PALETTE = 9
Private Const CF_PENDATA = 10
Private Const CF_SCRIPTSONLY = CF_ANSIONLY
Private Const CF_SELECTSCRIPT = &H400000
Private Const CF_SHOWHELP = &H4&
Private Const CF_SYLK = 4
Private Const CF_TEXT = 1
Private Const CF_TIFF = 6
Private Const CF_TTONLY = &H40000
Private Const CF_UNICODETEXT = 13
Private Const CF_USESTYLE = &H80&
Private Const CF_WAVE = 12
Private Const CF_WYSIWYG = &H8000
Private Const CFERR_CHOOSEFONTCODES = &H2000
Private Const CFERR_MAXLESSTHANMIN = &H2002
Private Const CFERR_NOFONTS = &H2001
Private Const CC_ANYCOLOR = &H100
Private Const CC_CHORD = 4
Private Const CC_CIRCLES = 1
Private Const CC_ELLIPSES = 8
Private Const CC_ENABLEHOOK = &H10
Private Const CC_ENABLETEMPLATE = &H20
Private Const CC_ENABLETEMPLATEHANDLE = &H40
Private Const CC_FULLOPEN = &H2
Private Const CC_INTERIORS = 128
Private Const CC_NONE = 0
Private Const CC_PIE = 2
Private Const CC_PREVENTFULLOPEN = &H4
Private Const CC_RGBINIT = &H1
Private Const CC_ROUNDRECT = 256 '
Private Const CC_SHOWHELP = &H8
Private Const CC_SOLIDCOLOR = &H80
Private Const CC_STYLED = 32
Private Const CC_WIDE = 16
Private Const CC_WIDESTYLED = 64
Private Const CCERR_CHOOSECOLORCODES = &H5000
Private Const LOGPIXELSY = 90
Private Const CCHDEVICENAME = 32
Private Const CCHFORMNAME = 32
Private Const SIMULATED_FONTTYPE = &H8000
Private Const PRINTER_FONTTYPE = &H4000
Private Const SCREEN_FONTTYPE = &H2000
Private Const BOLD_FONTTYPE = &H100
Private Const ITALIC_FONTTYPE = &H200
Private Const REGULAR_FONTTYPE = &H400
Private Const WM_CHOOSEFONT_GETLOGFONT = (&H400 + 1)
Private Const LBSELCHSTRING = "commdlg_LBSelChangedNotify"
Private Const SHAREVISTRING = "commdlg_ShareViolation"
Private Const FILEOKSTRING = "commdlg_FileNameOK"
Private Const COLOROKSTRING = "commdlg_ColorOK"
Private Const SETRGBSTRING = "commdlg_SetRGBColor"
Private Const FINDMSGSTRING = "commdlg_FindReplace"
Private Const HELPMSGSTRING = "commdlg_help"
Private Const CD_LBSELNOITEMS = -1
Private Const CD_LBSELCHANGE = 0
Private Const CD_LBSELSUB = 1
Private Const CD_LBSELADD = 2
Private Const NOERROR = 0
Private Const CSIDL_DESKTOP = &H0
Private Const CSIDL_PROGRAMS = &H2
Private Const CSIDL_CONTROLS = &H3
Private Const CSIDL_PRINTERS = &H4
Private Const CSIDL_PERSONAL = &H5
Private Const CSIDL_FAVORITES = &H6
Private Const CSIDL_STARTUP = &H7
Private Const CSIDL_RECENT = &H8
Private Const CSIDL_SENDTO = &H9
Private Const CSIDL_BITBUCKET = &HA
Private Const CSIDL_STARTMENU = &HB
Private Const CSIDL_DESKTOPDIRECTORY = &H10
Private Const CSIDL_DRIVES = &H11
Private Const CSIDL_NETWORK = &H12
Private Const CSIDL_NETHOOD = &H13
Private Const CSIDL_FONTS = &H14
Private Const CSIDL_TEMPLATES = &H15
Private Const BIF_RETURNONLYFSDIRS = &H1
Private Const BIF_DONTGOBELOWDOMAIN = &H2
Private Const BIF_STATUSTEXT = &H4
Private Const BIF_RETURNFSANCESTORS = &H8
Private Const BIF_BROWSEFORCOMPUTER = &H1000
Private Const BIF_BROWSEFORPRINTER = &H2000
Private Const HWND_BROADCAST = &HFFFF&
Private Const WM_WININICHANGE = &H1A

' Common Dialog 瀏覽資料夾
Public Function BrowseFolder(szDialogTitle As String) As String
  On Error GoTo ChkErr
    
    Dim BI As BROWSEINFO
    Dim dwIList As Long
    Dim szPath As String
    Dim wPos As Integer
    
    With BI
            .hOwner = 0&
            .pidlRoot = 0&
            .lpszTitle = szDialogTitle
            .ulFlags = BIF_RETURNONLYFSDIRS
    End With
    
    dwIList = SHBrowseForFolder(BI)
    szPath = Space(512)
    
    If (SHGetPathFromIDList(ByVal dwIList, ByVal szPath)) Then
        wPos = InStr(szPath, Chr(0))
        BrowseFolder = Left(szPath, wPos - 1)
    Else
        BrowseFolder = ""
    End If
    
  Exit Function
ChkErr:
    ErrHandle "CommDlgs", "BrowseFolder"
End Function

' Common Dialog 瀏覽檔案
Public Function DialogFile(szDialogTitle As String, _
                           szFilter As String, _
                           Optional szDefDir As String, _
                           Optional FilterIndex As Integer, _
                           Optional szFilename As String, _
                           Optional szDefExt As String, _
                           Optional wMode As Integer) As String
  On Error GoTo ChkErr
    
    Dim lngRet As Long
    Dim OFN As OPENFILENAME
    Dim szFile As String
    Dim szFileTitle As String
    
    With OFN
            .lStructSize = Len(OFN)
            .hwnd = 0&
            .lpstrTitle = szDialogTitle
            .lpstrFile = szFilename & String$(250 - Len(szFilename), 0)
            .nMaxFile = 255
            .lpstrFileTitle = String$(255, 0)
            .nMaxFileTitle = 255
            .lpstrFilter = Replace(szFilter, "|", Chr(0))
            .nFilterIndex = FilterIndex
            If Len(Trim(szDefDir)) > 0 Then .lpstrInitialDir = szDefDir
        
            If wMode = 1 Then
                .Flags = OFN_HIDEREADONLY Or OFN_PATHMUSTEXIST Or OFN_FILEMUSTEXIST
                lngRet = GetOpenFileName(OFN)
            Else
                .lpstrDefExt = szDefExt
                .Flags = OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT Or OFN_PATHMUSTEXIST
                lngRet = GetSaveFileName(OFN)
            End If
            
            If lngRet <> 0 Then
                If InStr(.lpstrFile, Chr$(0)) > 0 Then
                    szFile = Left$(.lpstrFile, InStr(.lpstrFile, Chr$(0)) - 1)
                End If
                DialogFile = szFile
            Else
                DialogFile = ""
            End If
    End With
    
  Exit Function
ChkErr:
    ErrHandle "mod_ComDlg", "DialogFile"
End Function

