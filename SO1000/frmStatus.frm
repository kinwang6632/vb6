VERSION 5.00
Object = "{3BB22E57-824A-11D3-BD70-0080C8F80BC4}#3.4#0"; "GiGridR.ocx"
Begin VB.Form frmStatus 
   BorderStyle     =   1  '單線固定
   Caption         =   "客戶狀態"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11790
   Icon            =   "frmStatus.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   11790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '系統預設值
   Begin prjGiGridR.GiGridR ggrData 
      Height          =   3405
      Left            =   0
      TabIndex        =   0
      Tag             =   "3495"
      Top             =   0
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   6006
      QuickMode       =   0   'False
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseCellForeColor=   -1  'True
   End
   Begin VB.Label Label6 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "BISP拆機時間"
      Height          =   180
      Left            =   9225
      TabIndex        =   6
      Top             =   3480
      Width           =   1080
   End
   Begin VB.Label lblBISPInstTime 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      BorderStyle     =   1  '單線固定
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   5790
      TabIndex        =   5
      Top             =   3450
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "BISP裝機時間"
      Height          =   180
      Left            =   4575
      TabIndex        =   4
      Top             =   3480
      Width           =   1080
   End
   Begin VB.Label lblBISPStatus 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      BorderStyle     =   1  '單線固定
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   975
      TabIndex        =   3
      Top             =   3450
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "BISP狀態"
      Height          =   180
      Left            =   165
      TabIndex        =   2
      Top             =   3480
      Width           =   720
   End
   Begin VB.Label lblBISPPRTime 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      BorderStyle     =   1  '單線固定
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   10395
      TabIndex        =   1
      Top             =   3450
      Width           =   1215
   End
End
Attribute VB_Name = "frmStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
'  On Error Resume Next
'    Dim intLoop As Integer
'    If Len(Environ("OS")) > 0 Then
'        For intLoop = 0 To 255 Step 17
'            MakeTransparent Me.hwnd, intLoop
'            DoEvents
'        Next
'    End If
  On Error GoTo ChkErr
    Dim mFlds As New GiGridFlds
    Dim rsStatus As New ADOR.Recordset
    With rsStatus
        .CursorLocation = adUseClient
'       .Open "SELECT COMPCODE,SERVICETYPE,CUSTSTATUSNAME,WIPNAME1,WIPNAME2,WIPNAME3,AccountOwner FROM " & GetOwner & "SO002 WHERE CUSTID=" & gCustId & " AND COMPCODE=" & gCompCode, gcnGi, adOpenForwardOnly, adLockReadOnly
'       .Open "SELECT A.COMPCODE,B.DESCRIPTION,A.CUSTSTATUSNAME,A.WIPNAME1,A.WIPNAME2,A.WIPNAME3,A.ACCOUNTOWNER,C.INSTCOUNT,C.BALANCE,A.INSTTIME,A.PRTIME,D.DisConnectDate FROM " & GetOwner & "SO002 A," & GetOwner & "CD046 B," & GetOwner & "SO001 C,SO031 D WHERE C.CUSTID=" & gCustId & " AND A.CUSTID=C.CUSTID AND A.COMPCODE=" & gCompCode & " AND B.CODENO=A.SERVICETYPE AND D.CUSTID(+)=A.CUSTID", gcnGi, adOpenForwardOnly, adLockReadOnly
        '#3387 增加風管類別欄位 By Kin 2007/09/07
        .Open "SELECT A.COMPCODE,B.DESCRIPTION,A.CUSTSTATUSNAME,C.RiskName,A.WIPNAME1,A.WIPNAME2,A.WIPNAME3,A.ACCOUNTOWNER,A.INSTCOUNT,C.BALANCE,A.BALANCE,A.INSTTIME,A.PRTIME FROM " & GetOwner & "SO002 A," & GetOwner & "CD046 B," & GetOwner & "SO001 C WHERE C.CUSTID=" & gCustId & " AND A.CUSTID=C.CUSTID AND A.COMPCODE=" & gCompCode & " AND B.CODENO=A.SERVICETYPE", gcnGi, adOpenForwardOnly, adLockReadOnly
    End With
    With mFlds
        .Add "COMPCODE", , , , , "公司別  "
        .Add "DESCRIPTION", , , , , "服務別  "
        .Add "CUSTSTATUSNAME", , , , , "客戶狀態   "
        '#3387 增加風管類別欄位 By Kin 2007/09/07
        .Add "RiskName", , , , , "風管類別" & Space(12)
        .Add "WIPNAME1", , , , , "裝機類狀態"
        .Add "WIPNAME2", , , , , "維修類狀態"
        .Add "WIPNAME3", , , , , "停拆移類狀態"
'        .Add "BALANCE", , , "###,###,###,###", , "欠費  ", vbRightJustify
        .Add "BALANCE", , , "###,###,###,###", , "欠費     ", vbRightJustify
        .Add "INSTCOUNT", , , , , "裝機台數", vbCenter
        .Add "INSTTIME", giControlTypeTime, , , , "裝機日期         "
        .Add "PRTIME", giControlTypeTime, , , , "拆機日期         "
    End With
    With ggrData
        .AllFields = mFlds
        .SetHead
         Set .Recordset = rsStatus
        .Refresh
    End With
    If gcnGi.Execute("SELECT SHOWBISP FROM " & GetOwner & "SO041 WHERE COMPCODE=" & gCompCode & " AND SERVICETYPE LIKE '%" & gServiceType & "%'").GetString(adClipString, , "", "", 0) = 1 Then
        ggrData.Height = 3400
        With frmSO1100BMDI
             lblBISPstatus = .rsSo001!BISPStatus & ""
'             lblBISPInstTime = Format(.rsSo001!BISPInstTime & "", "EE/MM/DD")
'             lblBISPPRTime = Format(.rsSo001!BISPPRTime & "", "EE/MM/DD")
             lblBISPInstTime = GetDTString(.rsSo001!BISPInstTime & "")
             lblBISPPRTime = GetDTString(.rsSo001!BISPPRTime & "")
             
        End With
    Else
        ggrData.Height = 3750
    End If
  Exit Sub
ChkErr:
    ErrSub Me.Name, "Form_Activate"
End Sub

'FOR v3.00
'Schema增加部分:
'   SO041 : ShowBISP Number(1) Default 0
'   SO001 : BISPStatus Varchar2(20) , BISPInstTime Date , BISPPRTime Date
'畫面處理機制:
'   1. SO7300Bp : 加一CheckBox元件處理[顯示BISP資料]
'   2. 於SO1100B中[狀態]Button之客戶狀態Form中加顯示BISP內容. 當SO041.ShowBISP=1時,則在客戶狀態Form中加Show BISP狀態資料(BISPStatus , BISPInstTime , BISPPRTime);
'       當SO041.ShowBISP=0時,則回覆為目前客戶狀態Form內容
'   3.請於拆機單列印機制中加串BISPStatus , BISPInstTime , BISPPRTime三個欄位,讓UESER可於列印拆機單時印出.

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error GoTo ChkErr
    If KeyCode = 27 Then Unload Me
  Exit Sub
ChkErr:
    ErrSub Me.Name, "Form_KeyDown"
End Sub

Private Sub Form_Load()
  On Error GoTo ChkErr
    With Me
        .Top = ((Screen.Height - .Height) / 2) - 800
        .Left = (Screen.Width - .Width) / 2
'         If .Picture <> 0 Then Call SetAutoRgn(Me)
'         If Len(Environ("OS")) > 0 Then MakeTransparent Me.hwnd, 0
    End With
  Exit Sub
ChkErr:
    ErrSub Me.Name, "Form_Load"
End Sub

'Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'  On Error GoTo ChkErr
'    If Button = vbLeftButton Then
'        Call SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
'        Call ReleaseCapture
'    End If
'  Exit Sub
'ChkErr:
'    errHandle Me.Name, "Form_MouseMove"
'End Sub

'Private Sub Form_Unload(Cancel As Integer)
''  On Error Resume Next
''    If Len(Environ("OS")) > 0 Then
''        Dim intLoop As Integer
''        For intLoop = 255 To 0 Step -17
''            MakeTransparent Me.hwnd, intLoop
''            DoEvents
''        Next
''    End If
''    Me.Hide
''    DeleteObject 0
''    ReleaseCapture
''    SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
'End Sub

'Private Sub ggrData_ShowCellData(ByVal Col As Long, ByVal Row As Long, ByVal giFld As prjGiGridR.GiGridFld, ByVal Fld As ADODB.Field, Value As Variant)
'  On Error GoTo ChkErr
'    If giFld.FieldName = "SERVICETYPE" Then
'        Select Case Value
'               Case "C"
'                    Value = "CATV"
'               Case "I"
'                    Value = "CAISP"
'        End Select
'    End If
'  Exit Sub
'ChkErr:
'    ErrSub Me.Name, "ggrData_ShowCellData"
'End Sub
