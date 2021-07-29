Attribute VB_Name = "mod_Const"
' PR : Hammer
' Start date : 2004/12/24
' Last Modify : 2004/12/25

Option Explicit

Public Const TwoPower16 = 2 ^ 16

Public Const SM_CYCAPTION = 4
Public Const SM_CXDLGFRAME = 7 'Width of dialog box borders
Public Const SM_CYDLGFRAME = 8 'Height of dialog box borders

'GLOBAL DRAW STAGE VALUES
'Public Const CDDS_ITEM           As Long = &H10000
'Public Const CDDS_PREPAINT       As Long = &H1        'Code Not Ready ;)
'Public Const CDRF_NOTIFYITEMDRAW As Long = &H20
'Public Const CDDS_ITEMPREPAINT   As Long = (CDDS_ITEM Or CDDS_PREPAINT)
'Public Const CDRF_DODEFAULT As Long = &H0
'Public Const CDDS_POSTPAINT As Long = &H2
'Public Const CDDS_ITEMPOSTPAINT As Long = (CDDS_ITEM Or CDDS_POSTPAINT)

'COMBOBOX CONST
Public Const CBN_CLOSEUP        As Long = 8
Public Const CB_GETDROPPEDSTATE As Long = &H157
 
'PROGRESS BAR CONST
 Public Const WM_USER         As Long = &H400
 Public Const CCM_FIRST       As Long = &H2000&
 Public Const CCM_SETBKCOLOR  As Long = (CCM_FIRST + 1)
 Public Const PBM_SETBARCOLOR As Long = (WM_USER + 9)

'FORMAT OF TEXT
Public Const DT_CENTER     As Long = &H1
Public Const DT_VCENTER    As Long = &H4
Public Const DT_SINGLELINE As Long = &H20
Public Const DT_CALCRECT   As Long = &H400
Public Const DT_WORDBREAK  As Long = &H10
Public Const WM_SETFONT    As Long = &H30
Public Const WM_GETFONT    As Long = &H31

'SPECIFIES GRADIENT FILL MODE
Public Const GRADIENT_HORIZONTAL As Long = &H0
Public Const GRADIENT_VERTICAL   As Long = &H1

''BUTTONS STATES
Public Const BM_GETCHECK As Long = &HF0
Public Const BS_LEFTTEXT As Long = &H20&
Public Const BM_GETSTATE As Long = &HF2&
Public Const BST_CHECKED As Long = &H1&
Public Const BST_PUSHED  As Long = &H4&

'WINDOWS SUBCLASS MESSAGES
Public Const WM_ACTIVATE     As Long = &H6
Public Const WM_COMMAND      As Long = &H111
Public Const WM_DRAWITEM     As Long = &H2B
Public Const WM_ENABLE       As Long = &HA
Public Const WM_KEYDOWN      As Long = &H100
Public Const WM_KEYUP        As Long = &H101
Public Const WM_KILLFOCUS    As Long = &H8
Public Const WM_LBUTTONDOWN  As Long = &H201        '//--- Subclass Messages 'iMsg'
Public Const WM_LBUTTONUP    As Long = &H202
Public Const WM_MOUSEMOVE    As Long = &H200
Public Const WM_PAINT        As Long = &HF
Public Const WM_SETFOCUS     As Long = &H7
Public Const WM_TIMER        As Long = &H113
Public Const WM_PRINTCLIENT  As Long = &H318
Public Const WM_NOTIFY       As Long = &H4E
Public Const WM_ENTERIDLE    As Long = &H121
Public Const WM_GETICON      As Long = &H7F
Public Const WM_INITDIALOG   As Long = &H110

'WINDOWS STYLES
Public Const GWL_WNDPROC          As Long = (-4)
Public Const GWL_STYLE            As Long = (-16)
Public Const GWL_EXSTYLE          As Long = (-20)
Public Const GW_CHILD             As Long = &H5
Public Const WH_CALLWNDPROC       As Long = 4
Public Const WS_BORDER            As Long = &H800000
Public Const WS_EX_CLIENTEDGE     As Long = &H200
Public Const WS_EX_STATICEDGE     As Long = &H20000
Public Const SWP_NOMOVE           As Long = &H2
Public Const SWP_NOSIZE           As Long = &H1
Public Const SWP_FRAMECHANGED     As Long = &H20
Public Const SWP_NOACTIVATE       As Long = &H10
Public Const SWP_NOZORDER         As Long = &H4
Public Const SWP_DRAWFRAME        As Long = SWP_FRAMECHANGED
Public Const SWP_FLAGS            As Long = SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_DRAWFRAME

'COLOR SCHEMES CONST
Public Const XPBlue_ButtonFace = &HD8E9EC
Public Const XPBlue_Highlight = &HC56A31
Public Const XPGreen_Highlight = &H70A093
Public Const XPSilver_Highlight = &HB99D7F
Public Const XPBlue_GrayText = &H99A8AC
Public Const XPBlue_ProgressBar = &H2BD228
Public Const XPGreen_ProgressBar = &H4A86E4
Public Const XPSilver_ProgressBar = &H76AE83

'AREAS CONSTS
 Public Const RGN_DIFF  As Long = 4




