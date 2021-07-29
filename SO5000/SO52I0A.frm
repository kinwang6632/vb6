VERSION 5.00
Object = "{A0C51418-555F-11D3-BC10-0050BA01A3A9}#9.8#0"; "GiMulti.ocx"
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{074771B7-6ADE-11D5-B178-00485456B028}#1.16#0"; "csMulti.ocx"
Begin VB.Form frmSO52I0A 
   BorderStyle     =   1  '單線固定
   Caption         =   "產生未收費 Call Out電話 [SO52I0A]"
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7815
   Icon            =   "SO52I0A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   7815
   StartUpPosition =   2  '螢幕中央
   Begin VB.TextBox txtOut 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   2040
      TabIndex        =   19
      Top             =   4785
      Width           =   1890
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "F5.開始"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1965
      TabIndex        =   25
      Top             =   6045
      Width           =   1275
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&X)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4560
      TabIndex        =   26
      Top             =   6045
      Width           =   1275
   End
   Begin VB.TextBox txtDataPath 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2040
      TabIndex        =   23
      ToolTipText     =   "請輸入字檔之路徑及檔名！"
      Top             =   5610
      Width           =   4755
   End
   Begin VB.CommandButton cmdDataPath 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6810
      TabIndex        =   24
      Top             =   5610
      Width           =   435
   End
   Begin VB.TextBox txtNotOut 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   2040
      TabIndex        =   20
      Top             =   5205
      Width           =   1905
   End
   Begin VB.Frame fraPage 
      Caption         =   "產出格式"
      ForeColor       =   &H00000000&
      Height          =   630
      Left            =   4095
      TabIndex        =   36
      Top             =   4785
      Width           =   3615
      Begin VB.OptionButton optOrigi 
         Caption         =   "原Call Out格式"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   255
         TabIndex        =   21
         Top             =   285
         Value           =   -1  'True
         Width           =   1440
      End
      Begin VB.OptionButton optChungHwa 
         Caption         =   "中華電信格式"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1935
         TabIndex        =   22
         Top             =   270
         Width           =   1515
      End
   End
   Begin VB.TextBox txtPeriod 
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   5055
      MaxLength       =   2
      TabIndex        =   7
      Top             =   945
      Width           =   405
   End
   Begin VB.TextBox txtAmoun 
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   6525
      MaxLength       =   6
      TabIndex        =   9
      Top             =   945
      Width           =   1125
   End
   Begin VB.ComboBox cmbPeriod 
      Height          =   300
      ItemData        =   "SO52I0A.frx":0442
      Left            =   4455
      List            =   "SO52I0A.frx":0455
      Style           =   2  '單純下拉式
      TabIndex        =   6
      Top             =   975
      Width           =   585
   End
   Begin VB.ComboBox cmbAmoun 
      Height          =   300
      ItemData        =   "SO52I0A.frx":046A
      Left            =   5925
      List            =   "SO52I0A.frx":047D
      Style           =   2  '單純下拉式
      TabIndex        =   8
      Top             =   975
      Width           =   585
   End
   Begin VB.TextBox txtVioNo 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1305
      MaxLength       =   1
      TabIndex        =   3
      Top             =   1005
      Width           =   405
   End
   Begin Gi_Date.GiDate gdaClctDate2 
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   150
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      ForeColor       =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Gi_Date.GiDate gdaClctDate1 
      Height          =   375
      Left            =   1305
      TabIndex        =   0
      Top             =   150
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      ForeColor       =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin CS_Multi.CSmulti gimCitemCode 
      Height          =   375
      Left            =   15
      TabIndex        =   10
      Top             =   1365
      Width           =   7755
      _ExtentX        =   13679
      _ExtentY        =   661
      ButtonCaption   =   "收  費  項  目"
   End
   Begin Gi_Multi.GiMulti gimCMCode 
      Height          =   375
      Left            =   15
      TabIndex        =   11
      Top             =   1725
      Width           =   7755
      _ExtentX        =   13679
      _ExtentY        =   661
      ButtonCaption   =   "收  費  方  式"
   End
   Begin Gi_Multi.GiMulti gimStatusCode 
      Height          =   375
      Left            =   15
      TabIndex        =   13
      Top             =   2475
      Width           =   7755
      _ExtentX        =   13679
      _ExtentY        =   661
      ButtonCaption   =   "客  戶  狀  態"
   End
   Begin CS_Multi.CSmulti gimClassCode 
      Height          =   375
      Left            =   15
      TabIndex        =   12
      Top             =   2085
      Width           =   7755
      _ExtentX        =   13679
      _ExtentY        =   661
      ButtonCaption   =   "客  戶  類  別"
   End
   Begin Gi_Multi.GiMulti gimAreaCode 
      Height          =   375
      Left            =   30
      TabIndex        =   15
      Top             =   3255
      Width           =   7755
      _ExtentX        =   13679
      _ExtentY        =   661
      ButtonCaption   =   "行     政     區"
   End
   Begin Gi_Multi.GiMulti gimServCode 
      Height          =   375
      Left            =   30
      TabIndex        =   14
      Top             =   2865
      Width           =   7755
      _ExtentX        =   13679
      _ExtentY        =   661
      ButtonCaption   =   "服     務     區"
   End
   Begin CS_Multi.CSmulti gimStrtCode 
      Height          =   375
      Left            =   15
      TabIndex        =   16
      Top             =   3630
      Width           =   7755
      _ExtentX        =   13679
      _ExtentY        =   661
      ButtonCaption   =   "街  道  範  圍"
   End
   Begin CS_Multi.CSmulti gimMduId 
      Height          =   375
      Left            =   15
      TabIndex        =   17
      Top             =   4005
      Width           =   7755
      _ExtentX        =   13679
      _ExtentY        =   661
      ButtonCaption   =   "大  樓  編  號"
   End
   Begin CS_Multi.CSmulti gimClctEn 
      Height          =   345
      Left            =   0
      TabIndex        =   18
      Top             =   4380
      Width           =   7755
      _ExtentX        =   13679
      _ExtentY        =   609
      ButtonCaption   =   "收  費  人   員"
   End
   Begin Gi_Date.GiDate gdaCallEnd 
      Height          =   375
      Left            =   1305
      TabIndex        =   2
      Top             =   585
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   661
      ForeColor       =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin prjGiList.GiList gilCompCode 
      Height          =   315
      Left            =   4740
      TabIndex        =   4
      Top             =   105
      Width           =   2940
      _ExtentX        =   5186
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      F5Corresponding =   -1  'True
   End
   Begin prjGiList.GiList gilServiceType 
      Height          =   315
      Left            =   4740
      TabIndex        =   5
      Top             =   540
      Width           =   2940
      _ExtentX        =   5186
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      F5Corresponding =   -1  'True
   End
   Begin MSComDlg.CommonDialog cnd1 
      Left            =   7260
      Top             =   5505
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "不輸出的電話號碼"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   390
      TabIndex        =   39
      Top             =   5250
      Width           =   1560
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "資料檔位置(路徑+名稱)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   15
      TabIndex        =   38
      Top             =   5685
      Width           =   1995
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "要輸出的電話號碼開頭"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   30
      TabIndex        =   37
      Top             =   4860
      Width           =   1950
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "期數"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   4005
      TabIndex        =   35
      Top             =   1035
      Width           =   390
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "金額"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   5505
      TabIndex        =   34
      Top             =   1035
      Width           =   390
   End
   Begin VB.Label lblVioNo 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "語音檔編號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   105
      TabIndex        =   33
      Top             =   1020
      Width           =   975
   End
   Begin VB.Label lblCallEnd 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "撥號結束日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   270
      Left            =   105
      TabIndex        =   32
      Top             =   675
      Width           =   1185
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BackStyle       =   0  '透明
      Caption         =   "服務類別"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   3900
      TabIndex        =   31
      Top             =   630
      Width           =   810
   End
   Begin VB.Label lblCompCode 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BackStyle       =   0  '透明
      Caption         =   "公  司  別"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   3900
      TabIndex        =   30
      Top             =   195
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "---"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   2445
      TabIndex        =   29
      Top             =   240
      Width           =   150
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "日期範圍"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   90
      TabIndex        =   28
      Top             =   360
      Width           =   780
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "下次收費"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   90
      TabIndex        =   27
      Top             =   150
      Width           =   780
   End
End
Attribute VB_Name = "frmSO52I0A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private strPath As String
Private strErrPath As String
Private rsTmp As New ADODB.Recordset
Private FSO As New FileSystemObject
Private Sub cmdOK_Click()
    Dim S As String
    If Not IsDataOk Then Exit Sub
    S = Left(txtDataPath, InStrRev(txtDataPath, "\"))
    If S = "C:\" Or S = "c:\" Then
    Else
        If Not ChkDirExist(S) Then MsgBox "路徑 " & S & " 不存在!", vbExclamation: Exit Sub
    End If
    Call ScrToRcd
    Call subChoose
    If Not BeginTran Then
        Screen.MousePointer = vbDefault
        cmdOK.Enabled = True: cmdCancel.Enabled = True
'        Unload Me
        Exit Sub
    End If
End Sub

Private Function OriFmt(ByRef rsObj As ADODB.Recordset) As String
    With rsObj
        OriFmt = " " & GetString(.Fields("Tel1"), 7) & _
                   "," & GetString(.Fields("Tel3"), 18) & _
                   "," & GetString(gdaCallEnd.GetValue, 8) & _
                   "," & GetString(txtVioNo, 1) & _
                   "," & GetString(.Fields("Custid"), 8, giRight)
    End With
End Function
Private Function ChuHw(ByRef rsObj As ADODB.Recordset) As String
    With rsObj
        ChuHw = GetString(rsTmp("TelArea"), 2) & _
                 GetString(rsTmp("Tel1"), 7)
    End With
End Function

Private Function WriteLine(rsTmp As ADODB.Recordset, ByRef strTxt As String) As Boolean
    On Error GoTo ChkErr
    WriteLine = False
    '原格式
    Dim bytFileNo  As Byte
    bytFileNo = GetFileNo(txtDataPath)
    If optOrigi Then
        Print #bytFileNo, " "  '首筆
        With rsTmp                     ' 內容
            While Not .EOF
                 Print #bytFileNo, OriFmt(rsTmp)
                .MoveNext
            Wend
        End With
    Else
        With rsTmp                     ' 內容
            While Not .EOF
                 Print #bytFileNo, ChuHw(rsTmp)
                .MoveNext
            Wend
        End With
    End If
''    '原格式
'    If optOrigi Then
''        首筆
'        strTxt = GetString("", 1) & vbCrLf
''        內容
'        Do Until rsTmp.EOF
'            strTxt = strTxt & GetString("", 1) & GetString(rsTmp("Tel1"), 7) & "," & GetString(rsTmp("Tel3"), 18) & _
'                     "," & GetString(gdaCallEnd.GetValue, 8) & "," & GetString(txtVioNo, 1) & "," & GetString(rsTmp("Custid"), 8, giRight) & vbCrLf
'            rsTmp.MoveNext
'        Loop
'    Else
'    '中華電信格式
'        Do Until rsTmp.EOF
'            strTxt = strTxt & GetString(rsTmp("TelArea"), 2) & GetString(rsTmp("Tel1"), 7) & vbCrLf
'            rsTmp.MoveNext
'             DoEvents
'        Loop
'    End If
'    Call CreateObject("Scripting.FileSystemObject").CreateTextFile(txtDataPath, True).Write(strTxt)

    Close #bytFileNo

    WriteLine = True
    Exit Function
ChkErr:
    WriteLine = False
    ErrSub Me.Name, "WriteLine"
End Function

Private Function GetRsTmp(ByRef rs As ADODB.Recordset) As Boolean
    On Error GoTo ChkErr
    Dim strFrom  As String, strWhere As String
    Dim strsql As String, strField As String
    strFrom = "From SO003,SO001 "
    strField = "Select SO001.CUSTID,SO001.CUSTNAME,SO001.TEL1,SO001.TEL3,SO001.CLASSNAME1 "
    strWhere = " Where SO003.CUSTID=SO001.CUSTID "
    If InStr(1, strChoose, "SO002.") > 0 Then
        strFrom = strFrom & ",SO002"
        strWhere = strWhere & "AND SO003.CUSTID=SO002.CUSTID AND SO003.ServiceType=SO002.ServiceType AND SO003.CompCode=SO002.CompCode "
    End If
    If InStr(1, strChoose, "SO014.") > 0 Then
        strFrom = strFrom & ",SO014"
        strWhere = strWhere & "AND SO001.MailAddrNo=SO014.AddrNo "
    End If
    If optChungHwa Then
        strField = strField & ",SO041.TelArea "
        strFrom = strFrom & ",SO041"
        strWhere = strWhere & "AND SO003.CompCode=SO041.CompCode "
    End If
    strsql = strField & strFrom & strWhere & IIf(strChoose = "", "", " AND ") & strChoose
    If Not GetRS(rs, strsql, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic, , , True) Then Exit Function
    GetRsTmp = True
    Exit Function
ChkErr:
    ErrSub Me.Name, "GetRsTmp"
End Function

Private Function BeginTran() As Boolean
    On Error GoTo ChkErr
    Dim strBankId As String, strsql As String
    Dim strOldBillNo As String
    Dim lngCountCustId As Long, lngErrCount As Long
    Dim lngPara24 As Long
    Dim lngTime As Long
    Dim rsUpd As New ADODB.Recordset
    Dim strTxt As String
        lngTime = Timer
'        取資料
        Screen.MousePointer = vbHourglass
        cmdOK.Enabled = False: cmdCancel.Enabled = False
        
        If Not GetRsTmp(rsTmp) Then Exit Function

        If rsTmp.RecordCount = 0 Then
            MsgBox "無資料！請重新操作！", vbInformation, "訊息！"
            Exit Function
        Else
            If Not WriteLine(rsTmp, strTxt) Then Exit Function
        
'            If Not InsertHead(rsTmp) Then GoTo KillFile
'            If Not InsertDetail(rsTmp) Then GoTo KillFile
'            If Not InsertFinal(rsTmp) Then GoTo KillFile
            MsgBox "已完成資料筆數共" & rsTmp.RecordCount & "筆," & vbCrLf & vbCrLf & _
           "問題筆數共" & lngErrCount & "筆," & vbCrLf & vbCrLf & _
           "共花費:" & (Timer - lngTime) \ 1 & "秒"
'            msgResult rsTmp.RecordCount, lngErrCount, lngTime       '顯示執行結果
        End If
        Screen.MousePointer = vbDefault
        cmdOK.Enabled = True: cmdCancel.Enabled = True
        BeginTran = True
        On Error Resume Next
        rsTmp.Close
        Set rsTmp = Nothing
    Exit Function

ChkErr:
    Call ErrSub(Me.Name, "BeginTran")
End Function

Private Sub Form_Activate()
    On Error Resume Next
        Screen.MousePointer = vbDefault
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Call FunctionKey(KeyCode, Shift, frmSO52I0A)
End Sub

Private Sub Form_Load()
  On Error GoTo ChkErr
    Call subGim
    Call subGil
    
    Call RcdToScr
    gilCompCode.SetCodeNo garyGi(9)
    gilCompCode.Query_Description
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "Form_Load")
End Sub

Private Sub ScrToRcd()
    On Error GoTo ChkErr
    Dim LogFile As TextStream
        
        Set LogFile = FSO.CreateTextFile(strErrPath & "\" & Me.Name & ".log", True)
         With LogFile
                .WriteLine (txtOut.Text)        '要輸出的電話號碼開頭
                .WriteLine (txtNotOut.Text)     '不輸出的電話號碼
                .WriteLine (txtDataPath.Text)   '資料檔位
        End With
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "ScrToRcd")
End Sub

Private Sub RcdToScr()
    On Error GoTo ChkErr
    Dim LogFile As TextStream
    Dim rsInv As New ADODB.Recordset
    Dim strsql As String
        Set rsInv = Nothing
        strErrPath = ReadGICMIS1("ErrLogPath")
        If Dir(strErrPath & "\" & Me.Name & ".log") = "" Then Exit Sub
            Set LogFile = FSO.OpenTextFile(strErrPath & "\" & Me.Name & ".log", ForReading)
            With LogFile
                    If Not .AtEndOfStream Then txtOut.Text = .ReadLine & ""
                        '要輸出的電話號碼開頭
                    If Not .AtEndOfStream Then txtNotOut.Text = .ReadLine & ""
                        '不輸出的電話號碼
                    If Not .AtEndOfStream Then txtDataPath.Text = .ReadLine & ""
                        '資料檔位
            End With
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "RcdToScr")
End Sub

Private Sub subGim()
  On Error GoTo ChkErr
    Call SetgiMulti(gimCitemCode, "CodeNo", "Description", "CD019", "收費項目代碼", "收費項目名稱")
    Call SetgiMulti(gimCMCode, "CodeNo", "Description", "CD031", "收費方式代碼", "收費方式名稱")
    Call SetgiMulti(gimClassCode, "CodeNo", "Description", "CD004", "客戶類別代碼", "客戶類別名稱")
    Call SetgiMulti(gimStatusCode, "CodeNo", "Description", "CD035", "客戶狀態代碼", "客戶狀態名稱")
    Call SetgiMulti(gimServCode, "CodeNo", "Description", "CD002", "服務區代碼", "服務區名稱")
    Call SetgiMulti(gimAreaCode, "CodeNo", "Description", "CD001", "行政區代碼", "行政區名稱")
    Call SetgiMulti(gimStrtCode, "CodeNo", "Description", "CD017", "街道代碼", "街道名稱")
    Call SetgiMulti(gimMduId, "MduId", "Name", "SO017", "大樓編號", "大樓名稱")
    Call SetgiMulti(Me.gimClctEn, "EmpNo", "EmpName", "CM003", "收費人員代碼", "收費人員名稱")
    gimStatusCode.SetQueryCode "1"
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "subGim")
End Sub
Private Sub subGil()
  On Error GoTo ChkErr
    'Call SetgiList(gilClctEn, "EmpNo", "EmpName", "CM003")
    'Call SetgiList(gilCompCode, "CodeNo", "Description", "CD039")
    Call SetgiList(gilCompCode, "CodeNo", "Description", "CD039", , , , , , , GetCompCodeFilter(gilCompCode))
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "subGil")
End Sub
Public Function FolderDialog(Title As String) As String
  On Error GoTo ChkErr
    Dim ComDlg As Object
    Dim Result As String
    Set ComDlg = CreateObject("Common.Dialog")
    Result = ComDlg.FolderDialog(Title)
    Set ComDlg = Nothing
    FolderDialog = Result
  Exit Function
ChkErr:
    ErrSub Me.Name, "FolderDialog"
End Function

Private Sub cmdCancel_Click()
  On Error GoTo ChkErr
    Unload Me
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "GetBankData")
End Sub

Private Sub cmdDataPath_Click()
    On Error GoTo ChkErr
        With cnd1
                .FileName = txtDataPath.Text
                .Filter = "文字檔|*.txt"
                .InitDir = strPath
                .Action = 1
                txtDataPath = .FileName
        End With
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "cmdDataPath_Click")
End Sub


Private Sub gdaClctDate1_GotFocus()
  On Error Resume Next
    If gdaClctDate1.Text = "" Then gdaClctDate1.SetValue (RightDate)
End Sub

Private Sub gdaClctDate2_GotFocus()
  On Error Resume Next
    If gdaClctDate1.GetValue = "" Or gdaClctDate2.GetValue = "" Then gdaClctDate2.SetValue (gdaClctDate1.GetValue)
End Sub

Private Sub gdaClctDate2_Validate(Cancel As Boolean)
  On Error Resume Next
    Cancel = ChkDate2(gdaClctDate1, gdaClctDate2)
End Sub

Private Function IsDataOk() As Boolean
  On Error GoTo ChkErr
    If Len(gilCompCode.GetCodeNo) = 0 Then gilCompCode.SetFocus: strErrFile = "公司別": GoTo 66
'    If optAcceptTime Then
    If gdaClctDate1.GetValue = "" Then gdaClctDate1.SetFocus: strErrFile = "下次收費起始日": GoTo 66
    If gdaClctDate2.GetValue = "" Then gdaClctDate2.SetFocus: strErrFile = "下次收費截止日": GoTo 66
    If txtDataPath.Text = "" Then strErrFile = "資料檔位置": GoTo 66
    If optOrigi Then
        If gdaCallEnd.Text = "" Then strErrFile = "撥號結束日期": GoTo 66
        If txtVioNo.Text = "" Then strErrFile = "語音檔編號": GoTo 66
    End If
'    End If

    IsDataOk = True
    Exit Function
66:
    MsgMustBe (strErrFile)
  Exit Function
ChkErr:
  Call ErrSub(Me.Name, "IsDataOk")
End Function
Private Sub subChoose()

  On Error GoTo ChkErr
  Dim intSort As Integer
    strChoose = ""

    If gdaClctDate1.GetValue <> "" And gdaClctDate2.GetValue <> "" Then
        Call subAnd("SO003.ClctDate Between To_Date('" & gdaClctDate1.GetValue & "','YYYYMMDD') And To_Date('" & gdaClctDate2.GetValue & "','YYYYMMDD')")
    Else
        If gdaClctDate1.GetValue <> "" Then Call subAnd("SO003.ClctDate >= To_Date('" & gdaClctDate1.GetValue & "','YYYYMMDD')")
        If gdaClctDate2.GetValue <> "" Then Call subAnd("SO003.ClctDate < To_Date('" & gdaClctDate2.GetValue & "','YYYYMMDD')+1")
    End If
'    If gdaClctDate1.GetValue <> "" Or gdaClctDate2.GetValue <> "" Then strUseIndex = strUseIndex & " " & GetUseIndexStr("SO003", "ClctDate")
    
  '期數
    If txtPeriod <> "" Then Call subAnd("SO003.Period " & cmbPeriod & txtPeriod)
  '金額
    If txtAmoun <> "" Then Call subAnd("SO003.Amount" & cmbAmoun & txtAmoun)
    
  'GiMulti
    If gimCitemCode.GetQueryCode <> "" Then Call subAnd("SO003.CitemCode IN (" & gimCitemCode.GetQueryCode & ")")
    If gimCMCode.GetQryStr <> "" Then Call subAnd("SO003.CMCode " & gimCMCode.GetQryStr)
    If gimClassCode.GetQryStr <> "" Then Call subAnd("SO001.ClassCode1 " & gimClassCode.GetQryStr)
    If gimStatusCode.GetQryStr <> "" Then Call subAnd("SO002.CustStatusCode " & gimStatusCode.GetQryStr)
    If gimServCode.GetQryStr <> "" Then Call subAnd("SO001.ServCode " & gimServCode.GetQryStr)
    If gimAreaCode.GetQryStr <> "" Then Call subAnd("SO014.AreaCode " & gimAreaCode.GetQryStr)
    If gimStrtCode.GetQryStr <> "" Then Call subAnd("SO014.StrtCode " & gimStrtCode.GetQryStr)
    If gimClctEn.GetQryStr <> "" Then Call subAnd("SO014.ClctEn " & gimClctEn.GetQryStr)
    
    
  '大樓編號
    If gimMduId.GetQryStr <> "" Then
        Call subAnd("SO014.MduId " & gimMduId.GetQryStr)
'        blnSO014 = True
    Else
        If gimMduId.GetDispStr <> "" Then subAnd ("SO014.MduId is Not Null")
    End If
    
  'GiList
    If gilCompCode.GetCodeNo <> "" Then Call subAnd("SO003.CompCode=" & gilCompCode.GetCodeNo)
    If gilServiceType.GetCodeNo <> "" Then Call subAnd("SO003.ServiceType='" & gilServiceType.GetCodeNo & "'")
    If txtOut.Text <> "" Then Call subAnd("substr(SO001.TEL1,1,1) In ('" & Replace(txtOut.Text, ",", "','") & "')")
    If txtNotOut.Text <> "" Then Call subAnd("SO001.TEL1 Not In ('" & Replace(txtNotOut.Text, ",", "','") & "')")
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "subChoose")
End Sub
    
Private Sub gilCompCode_Change()
  On Error GoTo ChkErr
    Call SelectServType(gilCompCode.GetCodeNo, gilServiceType)
    gilServiceType.ListIndex = 1
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "gilCompCode_Change")
End Sub

Private Sub optChungHwa_Click()
    If optChungHwa Then
        gdaCallEnd.Enabled = False
        txtVioNo.Enabled = False
        lblCallEnd.Enabled = False
        lblVioNo.Enabled = False
    End If
End Sub

Private Sub optOrigi_Click()
    If optOrigi Then
        gdaCallEnd.Enabled = True
        txtVioNo.Enabled = True
        lblCallEnd.Enabled = True
        lblVioNo.Enabled = True
    End If
End Sub

