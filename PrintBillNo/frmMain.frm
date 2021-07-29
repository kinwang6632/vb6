VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  '單線固定
   Caption         =   "查詢帳務"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12825
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   12825
   StartUpPosition =   2  '螢幕中央
   Begin VB.Frame fraFind 
      Height          =   4755
      Left            =   120
      TabIndex        =   5
      Top             =   90
      Width           =   12615
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   705
         Left            =   11370
         TabIndex        =   8
         Top             =   4020
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "查 詢"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   36
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   300
         TabIndex        =   3
         Top             =   3900
         Width           =   2715
      End
      Begin VB.TextBox txtID 
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   36
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   840
         IMEMode         =   3  '暫止
         Left            =   4890
         MaxLength       =   9
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   2940
         Width           =   3945
      End
      Begin VB.TextBox txtTel 
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   36
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   840
         Left            =   3990
         MaxLength       =   12
         TabIndex        =   0
         Top             =   1050
         Width           =   4815
      End
      Begin VB.ComboBox cboIDWord 
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   36
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   840
         ItemData        =   "frmMain.frx":0442
         Left            =   3990
         List            =   "frmMain.frx":0494
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   2940
         Width           =   945
      End
      Begin VB.Label Label1 
         Caption         =   "(一次只能輸入一種查詢資料)"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   18
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   3150
         TabIndex        =   9
         Top             =   4050
         Width           =   4785
      End
      Begin VB.Label lblID 
         Caption         =   "身份證字號(寬頻查詢專用)"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   36
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   855
         Left            =   2190
         TabIndex        =   7
         Top             =   2010
         Width           =   9075
      End
      Begin VB.Label lblTel 
         Caption         =   "電話號碼"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   36
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   705
         Left            =   4560
         TabIndex        =   6
         Top             =   210
         Width           =   3105
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdData 
      Height          =   2685
      Left            =   30
      TabIndex        =   4
      ToolTipText     =   "滑鼠左鍵點2下即可選擇"
      Top             =   4860
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   4736
      _Version        =   393216
      Cols            =   5
      HighLight       =   2
      SelectionMode   =   1
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "標楷體"
         Size            =   18
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rs As New ADODB.Recordset
Private rsTmp As New ADODB.Recordset
Private gstrStationCount As String
Private blnNoDefine As Boolean
Private blnOK As Boolean

Private Sub cboIDWord_Change()
  On Error Resume Next
    If cboIDWord.Text <> "" Then
        txtTel.Text = ""
    End If
End Sub

Private Sub cboIDWord_Click()
  On Error Resume Next
    If cboIDWord.Text <> "" Then
        txtTel.Text = ""
    End If
End Sub

Private Sub cboIDWord_KeyPress(keyAscii As Integer)
    If keyAscii = 8 Then Exit Sub
    If cboIDWord.SelText = "" Then
        If Len(cboIDWord.Text) > 0 Then keyAscii = 0: Exit Sub
    End If
    If keyAscii >= 97 And keyAscii <= 122 Then
        keyAscii = keyAscii - 32
    End If
    If keyAscii >= 65 And keyAscii <= 90 Then Exit Sub
    keyAscii = 0
End Sub

Private Sub cmdFind_Click()
  On Error GoTo ChkErr
    Dim strQry As String
    Dim lngBillcount As Long
    strQry = Empty
    If Not IsDataOK Then
        GoTo 88
    End If

    Screen.MousePointer = vbHourglass
    blnNoDefine = False
    If Not DefineRs Then Exit Sub
    Set grdData.Recordset = Nothing
    'Set grdData.Recordset = rsTmp
    SetData
    If rs.RecordCount <= 0 Then
        ClearCtl
        grdData.Rows = 2
        grdData.Col = 1
        grdData.Row = 1
        txtTel.SetFocus
        CreateObject("WScript.Shell").Popup "查無此電話號碼/身分證！", 1, "警告", 0
        GoTo 88
    Else
        rs.MoveFirst
        rsTmp.MoveFirst
        If rsTmp.RecordCount = 1 Then
            lngBillcount = GetRsValue("Select Count(*) From " & GetOwner & "SO033 Where " & _
                                " UCCode is Not Null And CancelFlag=0 And CustId=" & rsTmp("CustId"), gcnGi)
            If lngBillcount = 0 Then
                MsgBox "無任何欠費資料！", vbInformation, "警告"
                ClearCtl
                txtTel.SetFocus
                blnNoDefine = True
                DefineRs
                SetData
                GoTo 88
            Else
                frmPrint.uCustId = rsTmp("CustId") & ""
                frmPrint.Show 1
                ClearCtl
                txtTel.SetFocus
                blnNoDefine = True
                DefineRs
                SetData
            End If
        End If
    End If
88:
    On Error Resume Next
    Screen.MousePointer = vbDefault
    'grdData.SetFocus
    Exit Sub
ChkErr:
  ErrSub Me.Name, "CmdFind_Click"
End Sub
Private Function IsDataOK() As Boolean
  On Error GoTo ChkErr
    If cboIDWord.Text <> "" Or txtID.Text <> "" Then
        If cboIDWord.Text = "" Then
            MsgBox "請輸入身份證第一碼英文字母", vbInformation, "警告"
            cboIDWord.SetFocus
            IsDataOK = False
            Exit Function
        End If
        If Trim(txtID.Text) = "" Then
            MsgBox "請輸入身份證完整數字", vbInformation, "警告"
            txtID.SetFocus
            IsDataOK = False
            Exit Function
        End If
    End If
    If Trim(txtTel.Text) = "" And Trim(txtID.Text) = "" Then
        MsgBox "請至少輸入一種查詢條件", vbInformation, "警告"
        txtTel.SetFocus
        IsDataOK = False
        Exit Function
    End If
    IsDataOK = True
    Exit Function
ChkErr:
  Call ErrSub(Me.Name, "IsDataOK")
End Function
Private Function subChoose() As String
  On Error GoTo ChkErr
    Dim strWhere As String
    Dim strTel As String
    Dim strQry1, strQry2 As String
    Dim strTable As String
    strTel = Empty
    strWhere = Empty
    'strqry = " Select * From " & GetOwner & "SO001 A"
    If txtTel.Text <> "" Then
        strTel = Replace(txtTel.Text, "-", "")
        strTel = Trim(strTel)
'        strWhere = "(A.Tel1='" & strTel & "' Or A.Tel2='" & strTel & "'" & _
'                     " Or A.Tel3='" & strTel & "'" & _
'                     " Or A.ContTel='" & strTel & "'" & _
'                     " Or A.TelH1 || A.Tel1='" & strTel & "'" & _
'                     " Or A.TelH2 || A.Tel2='" & strTel & "'" & _
'                     " Or A.TelH3 || A.Tel3='" & strTel & "'" & _
'                     " Or B.ContTel='" & strTel & "' Or B.Contmobile='" & strTel & "')"
        strTable = "Select A.CustId,A.CustName,Null DeclarantName," & _
                "A.Tel1,A.Tel2,A.Tel3,Null ContTel,Null ContMobile " & _
                " From " & GetOwner & "SO001 A Where A.Tel1='" & strTel & "'" & _
                " Or A.Tel2='" & strTel & "' Or A.Tel3='" & strTel & "'" & _
                " Union " & _
                "Select A.CustId,Null CustName,A.DeclarantName,Null Tel1,Null Tel2,Null Tel3," & _
                "A.ContTel,A.ContMobile From " & GetOwner & "SO004 A" & _
                " Where ContTel='" & strTel & "' Or ContMobile='" & strTel & "'" & _
                " Union " & _
                "Select A.CustId,Null CustName,Null DeclarantName,ContTel Tel1,Null Tel2,Null Tel3," & _
                "Null ContTel,Null ContMobile From " & GetOwner & "SO002 A" & _
                " Where ContTel='" & strTel & "'"
    End If

    If txtID.Text <> "" Then
        If strTable = Empty Then
            strTable = strTable & IIf(strTable = Empty, "", " Union ") & _
                    " Select A.CustId,Null CustName,A.DeclarantName,Null Tel1,Null Tel2,Null Tel3," & _
                        "A.ContTel,A.ContMobile From " & GetOwner & "SO004 A" & _
                        " Where A.ID='" & cboIDWord.Text & txtID.Text & "' Or A.ID2='" & cboIDWord & txtID.Text & "'"
        Else
            strTable = "Select A.* From (" & strTable & ") A Where Exists(Select * From SO004 Where " & _
                    "SO004.CUSTID=A.CUSTID AND (SO004.ID='" & cboIDWord.Text & txtID.Text & "' OR SO004.ID2='" & _
                    cboIDWord.Text & txtID.Text & "'))"
        End If
        'strWhere = " And A.ID='" & cboIDWord.Text & txtID.Text & "'"
        'strTable = "," & GetOwner & "SO004 B"
        'strWhere = IIf(strWhere = "", "And A.ID='" & cboIDWord.Text & txtID & "'", strWhere & " And A.ID='" & txtID.Text & "'")
    End If
    strTable = ",(" & strTable & ") B"
    subChoose = "Select distinct A.RowId,A.CustId,A.CustName,B.DeclarantName," & _
                "A.Tel1,A.Tel2,A.Tel3,A.InstAddress,B.ContTel,B.Contmobile,A.InstAddress " & _
                " From " & GetOwner & "SO001 A" & strTable & _
                " Where A.CustId=B.CustId " & strWhere & " Order By A.CustId"
        
   Exit Function
ChkErr:
  ErrSub Me.Name, "SubChoose"
End Function

Private Sub Command1_Click()
    Dim o As Printer
    Dim strComputer As String
    Dim objWMIService As Object
    Dim colItems As Object
    Dim objItem As Object
    Dim strDefault As String
    strComputer = "."
    Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
    Set colItems = objWMIService.ExecQuery( _
        "SELECT * FROM Win32_Printer", , 48)
    For Each objItem In colItems
        If objItem.Default Then
            strDefault = objItem.Caption
        End If
        
    Next
    For Each o In Printers
        If o.DeviceName = strDefault Then
            MsgBox strDefault
        End If
    Next
End Sub

Private Sub Form_Initialize()
  On Error GoTo ChkErr
    SetLocaleInfo GetSystemDefaultLCID, &H1F, "yyyy/MM/dd"
    gstrCurDir = CurPath
    
    If OpenDB Is Nothing Then Set OpenDB = CreateObject("giOpenDBObj.OpenDBObj")
    Call ReadINI ' 由 GICMIS2.INI 讀取資訊
    Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "Form_Initialize")
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error Resume Next
    If KeyCode = vbKeyEscape Then
        End
    End If
    If KeyCode = vbKeyF2 Then
        cmdFind.Value = True
    End If
End Sub

Private Sub Form_Load()
  On Error GoTo ChkErr
    
    'Set rs = New ADODB.Recordset
    If Not Alfa Then
        gcnGi.ConnectionString = garyGi(3)
        gcnGi.CursorLocation = adUseClient
        gcnGi.Open
        garyGi(9) = Command
        gCompCode = garyGi(9)
    Else
        gCompCode = garyGi(9)
    End If
    garyGi(16) = GetRsValue("Select TableOwner From CD039 Where CodeNo=" & gCompCode, gcnGi) & ""
    blnMedia = (Val(GetRsValue("Select Para24 From " & GetOwner & "SO043", gcnGi) & "") = 1)
    blnCombineBill = (GetSystemParaItem("CombineBill", gCompCode, "", "SO041", , True) = 1)
    strBillTab6 = GetRsValue("Select BillTab6 From " & GetOwner & "SO045 Where ServiceType='C'", gcnGi) & ""
    strBillPrinter2 = GetRsValue("Select BillPrinter2 From " & GetOwner & "SO045 Where ServiceType='C'", gcnGi) & ""
    blnNoDefine = True
    DefineRs
    'Set grdData.DataSource = rsTmp
    'grdData.Rows = 2
    'grdData.Row = 1
    SetData
    grdData.Refresh
    If strBillTab6 = "" Then
        MsgBox "[ 客戶補印收費單代號 ]設定錯誤！", vbInformation, "警告"
        End
    End If
    If strBillPrinter2 = "" Then
        MsgBox "[ 客戶補印收費單列表機 ]設定錯誤！", vbInformation, "警告"
        End
    End If
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "Form_Load")
End Sub
Private Function ReadINI() ' 由 GICMIS2.INI 讀取設定資料
  On Error GoTo ChkErr
    Dim SysPath As String
    Dim strTmp As String
    Dim lngLen As Long
    Dim pstrDataSource As String
    Dim pstrUserId As String
    Dim pstrMDBPath As String
    Dim pstrPassword As String
    Dim cpCode As String

    SysPath = GetIniPath2
    
    If Dir(SysPath) = "" Then ' 如果還是找不到 GICMIS.INI ，則脫離系統
        MsgBox "系統環境參數設定錯誤，無法開啟" & SysPath & " !", vbCritical, "錯誤"
        Unload Me
        End
    End If
    
    If Command <> Empty Then
        cpCode = Command()
        cpCode = Replace(cpCode, "CRM", "", 1, , vbTextCompare)
        If cpCode <> Empty Then
            cpCode = Val(Trim(cpCode))
            If cpCode = 0 Then cpCode = ReadFROMINI(SysPath, "DEFAULT", "DEFAULTDB", False) ' 由 GICMIS2.INI 讀取 設定值
        Else
            cpCode = ReadFROMINI(SysPath, "DEFAULT", "DEFAULTDB", False) ' 由 GICMIS2.INI 讀取 設定值
        End If
    Else
        cpCode = ReadFROMINI(SysPath, "DEFAULT", "DEFAULTDB", False) ' 由 GICMIS2.INI 讀取 設定值
    End If
    
    garyGi(14) = cpCode
    pstrDataSource = ReadFROMINI(SysPath, cpCode, "1", True)
    pstrUserId = ReadFROMINI(SysPath, cpCode, "2", True)
    pstrPassword = ReadFROMINI(SysPath, cpCode, "3", True)
    gstrStationCount = ReadFROMINI(SysPath, cpCode, "4", True) ' 允許多少工作站登入
    
    If Len(pstrUserId) = 0 Or Len(pstrPassword) = 0 Or Len(gstrStationCount) = 0 Then
        MsgBox "系 統 環 境 參 數 設 定 錯 誤 !!", , "錯誤訊息..!!"
        Unload Me ' 如果 資料來源 或 使用者代碼 空白，
        End
    End If
   
    If Alfa Then
        GetGlobal
    Else
        garyGi(3) = "Provider=MSDAORA.1;" & _
                     IIf(pstrDataSource = "", "", "Data Source=" + pstrDataSource + ";") & _
                    "Password=" & pstrPassword + ";" & _
                    "User ID=" & pstrUserId + ";" & _
                    "Persist Security Info=True"
    End If
    Exit Function
ChkErr:
    If ErrHandle(, True, , "ReadINI") Then Resume 0
    Unload Me
    End
End Function
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    Close3Recordset rs
    Close3Recordset rsTmp
End Sub
Private Function DefineRs() As Boolean
  On Error GoTo ChkErr
    Dim strQry As String
    Dim strDouble As String
    Dim i As Long
    strQry = Empty
    strQry = "Select distinct A.CustId,A.CustName,B.DeclarantName," & _
             " A.Tel1,A.Tel2,A.Tel3,B.ContTel,B.Contmobile,A.InstAddress,A.RowId " & _
             " From " & GetOwner & "SO001 A," & GetOwner & "SO004 B Where 1=0"
    If Not GetRS(rs, strQry, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
'    Set grdData.Recordset = Nothing
'    grdData.Clear
'    grdData.ClearStructure
'    grdData.Rows = 2
'    grdData.Row = 1
    With rsTmp
        If .State = adStateOpen Then .Close
        .CursorLocation = adUseClient
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        For i = 0 To rs.Fields.Count - 1
            .Fields.Append rs.Fields(i).Name, adBSTR, rs.Fields(i).DefinedSize, adFldIsNullable
            If i = 2 Then
                .Fields.Append "CustNameA", adBSTR, 50, adFldIsNullable
            End If
            If i = 7 Then
                .Fields.Append "TelA", adBSTR, 20, adFldIsNullable
            End If
        Next i
    End With
    rsTmp.Open
    If Not blnNoDefine Then
        strQry = subChoose
        If Not GetRS(rs, strQry, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
    Else
        blnNoDefine = False
        DefineRs = True
        Exit Function
    End If
    If rs.RecordCount > 0 Then
        rs.MoveFirst
    End If
    Do While Not rs.EOF
        If strDouble <> rs("CUSTID") & "" & IIf(IsNull(rs("DeclarantName")), rs("CUSTNAME") & "", rs("DeclarantName") & "") & _
                    IIf(txtTel.Text <> "", txtTel.Text, rs("Tel1")) & "" Then
            rsTmp.AddNew
            For i = 0 To rs.Fields.Count - 1
                If Not IsNull(rs.Fields(i).Value) Then
                    rsTmp.Fields(rs.Fields(i).Name).Value = rs.Fields(i).Value & ""
                End If
            Next i
            If Not IsNull(rs("DeclarantName")) Then
                rsTmp("CustNameA") = rs("DeclarantName") & ""
            Else
                rsTmp("CustNameA") = rs("CustName") & ""
            End If
            If txtTel.Text <> "" Then
                rsTmp("TelA") = txtTel.Text
            Else
                rsTmp("TelA") = rs("Tel1")
            End If
            rsTmp.Update
        End If
        strDouble = rs("CUSTID") & "" & IIf(IsNull(rs("DeclarantName")), rs("CUSTNAME") & "", rs("DeclarantName") & "") & _
                    IIf(txtTel.Text <> "", txtTel.Text, rs("Tel1")) & ""
        rs.MoveNext
        
    Loop
    If rsTmp.RecordCount > 0 Then
        rsTmp.MoveFirst
    End If
    DefineRs = True
    On Error Resume Next
    blnNoDefine = False
    Exit Function
ChkErr:
  ErrSub Me.Name, "DefineRs"
End Function

Private Sub SetData()
  On Error GoTo ChkErr
    Dim i As Long
'    grdData.ClearStructure
'    grdData.Rows = 2
'    grdData.Row = 1
'    Set grdData.Recordset = rsTmp
    grdData.Clear
    grdData.ColWidth(0, 0) = 400
    grdData.Cols = rsTmp.Fields.Count
    grdData.Rows = 2
    For i = 0 To rsTmp.Fields.Count - 1
        Select Case UCase(rsTmp.Fields(i).Name)
            Case "CUSTID"
                grdData.ColWidth(i + 1) = 1700
                grdData.TextMatrix(0, i + 1) = "客戶編號"
            Case "CUSTNAMEA"
                grdData.ColWidth(i + 1) = 2500
                grdData.TextMatrix(0, i + 1) = "申請人姓名"
            Case "TELA"
                grdData.ColWidth(i + 1) = 2000
                grdData.TextMatrix(0, i + 1) = "電話號碼"
            Case "INSTADDRESS"
                grdData.ColWidth(i + 1) = 6000
                grdData.TextMatrix(0, i + 1) = "裝機地址"
                
            Case Else
                grdData.ColWidth(i + 1) = 0
        End Select
    Next i
    If rsTmp.RecordCount = 0 Then
        grdData.Rows = 2
        grdData.Row = 1
    Else
        rsTmp.MoveFirst
        Do While Not rsTmp.EOF
            grdData.Rows = rsTmp.RecordCount + 1
            grdData.TextMatrix(rsTmp.AbsolutePosition, 1) = rsTmp("CustId") & ""
            grdData.Col = 1
            grdData.Row = rsTmp.AbsolutePosition
            grdData.CellAlignment = vbLeftJustify
            grdData.TextMatrix(rsTmp.AbsolutePosition, 4) = rsTmp("CUSTNAMEA") & ""
            grdData.TextMatrix(rsTmp.AbsolutePosition, 10) = rsTmp("TELA") & ""
            grdData.Col = 10
            grdData.Row = rsTmp.AbsolutePosition
            grdData.CellAlignment = vbLeftJustify
            grdData.TextMatrix(rsTmp.AbsolutePosition, 11) = rsTmp("InstAddress") & ""
            rsTmp.MoveNext
        Loop
    End If
    grdData.Refresh

    Exit Sub
ChkErr:
  
  Call ErrSub(Me.Name, "SetData")
End Sub

Private Sub grdData_DblClick()
  On Error GoTo ChkErr
    Dim lngBillcount As Long
    If rs.RecordCount > 1 Then
        rsTmp.AbsolutePosition = grdData.Row
        lngBillcount = GetRsValue("Select Count(*) From " & GetOwner & "SO033" & _
                                " Where UCCode is Not Null And CancelFlag=0 " & _
                                " And CustId=" & rsTmp("CustId"), gcnGi)
        If lngBillcount > 0 Then
            frmPrint.uCustId = rsTmp("CustId") & ""
            frmPrint.Show 1
            blnNoDefine = True
            DefineRs
            SetData
            ClearCtl
            txtTel.SetFocus
        Else
            MsgBox "選取的客戶資料無任何欠費資料！", vbInformation, "警告"
            ClearCtl
            blnNoDefine = True
            SetData
        End If
    End If
  Exit Sub
ChkErr:
  ErrSub Me.Name, "grdData_DlClick"
End Sub



Private Sub txtID_Change()
  On Error Resume Next
    If txtID.Text <> "" Then
        txtTel.Text = ""
    End If
End Sub

Private Sub txtID_KeyPress(keyAscii As Integer)
  On Error Resume Next
    If keyAscii = 8 Then
        Exit Sub
    End If
    If keyAscii >= 48 And keyAscii <= 57 Then
        
        Exit Sub
    End If
    keyAscii = 0
End Sub

Private Sub txtTel_Change()
  On Error Resume Next
    If txtTel.Text <> "" Then
        cboIDWord.Text = ""
        txtID.Text = ""
    End If
End Sub

Private Sub txtTel_KeyPress(keyAscii As Integer)
    On Error Resume Next
    If keyAscii = 8 Then Exit Sub
    If keyAscii >= 48 And keyAscii <= 57 Then Exit Sub
    keyAscii = 0
End Sub
Property Let uOK(ByVal vData As Boolean)
    blnOK = vData
End Property
Private Sub ClearCtl()
  On Error Resume Next
  txtID.Text = ""
  txtTel.Text = ""
  cboIDWord.Text = ""
End Sub
