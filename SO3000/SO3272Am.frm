VERSION 5.00
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.3#0"; "GiList.ocx"
Object = "{A0C51418-555F-11D3-BC10-0050BA01A3A9}#9.4#0"; "GiMulti.ocx"
Object = "{074771B7-6ADE-11D5-B178-00485456B028}#1.4#0"; "csMulti.ocx"
Begin VB.Form frmSO3272A 
   BorderStyle     =   1  '單線固定
   Caption         =   "信用卡扣款資料產生作業 [SO3272A]"
   ClientHeight    =   7110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10995
   Icon            =   "SO3272A.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   10995
   StartUpPosition =   1  '所屬視窗中央
   Begin VB.Frame Frame1 
      Caption         =   "查詢條件"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   6465
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   10695
      Begin VB.Frame Frame2 
         Caption         =   " 大樓依據"
         Height          =   1005
         Left            =   210
         TabIndex        =   22
         Top             =   5310
         Width           =   10185
         Begin VB.OptionButton optMduYes 
            Caption         =   "&A.要產生"
            Height          =   255
            Left            =   450
            TabIndex        =   17
            Top             =   690
            Value           =   -1  'True
            Width           =   1065
         End
         Begin VB.OptionButton optMduNo 
            Caption         =   "&B.不產生"
            Height          =   285
            Left            =   1830
            TabIndex        =   18
            Top             =   660
            Width           =   1125
         End
         Begin CS_Multi.CSmulti gmdMduid 
            Height          =   405
            Left            =   150
            TabIndex        =   16
            Top             =   240
            Width           =   9855
            _ExtentX        =   17383
            _ExtentY        =   714
            ButtonCaption   =   "大 樓 名 稱"
         End
      End
      Begin VB.CheckBox chkOther 
         Caption         =   "包含一般戶"
         Height          =   225
         Left            =   6375
         TabIndex        =   7
         Top             =   1155
         Value           =   1  '核取
         Width           =   2145
      End
      Begin Gi_Multi.GiMulti gmdBillType 
         Height          =   405
         Left            =   210
         TabIndex        =   14
         Top             =   4365
         Width           =   10035
         _ExtentX        =   17701
         _ExtentY        =   714
         ButtonCaption   =   "單 據 類 別"
      End
      Begin Gi_Multi.GiMulti gmdClctEn 
         Height          =   405
         Left            =   210
         TabIndex        =   13
         Top             =   3960
         Width           =   10035
         _ExtentX        =   17701
         _ExtentY        =   714
         ButtonCaption   =   "收    費    員"
      End
      Begin Gi_Multi.GiMulti gmdCustClass 
         Height          =   405
         Left            =   210
         TabIndex        =   12
         Top             =   3540
         Width           =   10035
         _ExtentX        =   17701
         _ExtentY        =   714
         ButtonCaption   =   "客 戶 類 別"
      End
      Begin Gi_Multi.GiMulti gmdServCode 
         Height          =   405
         Left            =   210
         TabIndex        =   10
         Top             =   2700
         Width           =   10035
         _ExtentX        =   17701
         _ExtentY        =   714
         ButtonCaption   =   "服    務    區"
      End
      Begin Gi_Multi.GiMulti gmdAreaCode 
         Height          =   405
         Left            =   210
         TabIndex        =   9
         Top             =   2280
         Width           =   10035
         _ExtentX        =   17701
         _ExtentY        =   714
         ButtonCaption   =   "行    政    區"
      End
      Begin Gi_Multi.GiMulti gmdCMCode 
         Height          =   405
         Left            =   210
         TabIndex        =   8
         Top             =   1890
         Width           =   10035
         _ExtentX        =   17701
         _ExtentY        =   714
         ButtonCaption   =   "收 費 方 式"
      End
      Begin Gi_Multi.GiMulti gimCustStatus 
         Height          =   405
         Left            =   210
         TabIndex        =   11
         Top             =   3120
         Width           =   10035
         _ExtentX        =   17701
         _ExtentY        =   714
         ButtonCaption   =   "客  戶 狀 態"
      End
      Begin Gi_Multi.GiMulti gimCreateEn 
         Height          =   405
         Left            =   210
         TabIndex        =   15
         Top             =   4785
         Width           =   10035
         _ExtentX        =   17701
         _ExtentY        =   714
         ButtonCaption   =   "產 生 人 員"
      End
      Begin prjGiList.GiList gilCompCode 
         Height          =   315
         Left            =   1200
         TabIndex        =   0
         Top             =   270
         Width           =   2940
         _ExtentX        =   5186
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         F2Corresponding =   -1  'True
      End
      Begin prjGiList.GiList gilServiceType 
         Height          =   315
         Left            =   6315
         TabIndex        =   1
         Top             =   270
         Width           =   2940
         _ExtentX        =   5186
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         F2Corresponding =   -1  'True
      End
      Begin Gi_Date.GiDate gdaCreateTime2 
         Height          =   345
         Left            =   2790
         TabIndex        =   3
         Top             =   675
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   609
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
      Begin Gi_Date.GiDate gdaCreateTime1 
         Height          =   345
         Left            =   1200
         TabIndex        =   2
         Top             =   675
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   609
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
      Begin Gi_Date.GiDate gdaShouldDate2 
         Height          =   345
         Left            =   7920
         TabIndex        =   6
         Top             =   660
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   609
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
      Begin Gi_Date.GiDate gdaShouldDate1 
         Height          =   345
         Left            =   6330
         TabIndex        =   5
         Top             =   660
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   609
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
      Begin prjGiList.GiList gilBank 
         Height          =   315
         Left            =   3255
         TabIndex        =   4
         Top             =   1080
         Width           =   2940
         _ExtentX        =   5186
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         F2Corresponding =   -1  'True
      End
      Begin Gi_Multi.GiMulti gmBank 
         Height          =   405
         Left            =   225
         TabIndex        =   30
         Top             =   1485
         Width           =   10035
         _ExtentX        =   17701
         _ExtentY        =   714
         ButtonCaption   =   "承  辦 銀 行"
      End
      Begin VB.Label lbluntil 
         AutoSize        =   -1  'True
         Caption         =   "至"
         Height          =   195
         Left            =   7590
         TabIndex        =   29
         Top             =   705
         Width           =   195
      End
      Begin VB.Label lblReadAmt 
         AutoSize        =   -1  'True
         Caption         =   "應收日期範圍"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   5070
         TabIndex        =   28
         Top             =   705
         Width           =   1170
      End
      Begin VB.Label lblServiceType 
         AutoSize        =   -1  'True
         Caption         =   "服務類別"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   5385
         TabIndex        =   27
         Top             =   315
         Width           =   780
      End
      Begin VB.Label lblBank 
         AutoSize        =   -1  'True
         Caption         =   "承辦銀行"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   2790
         TabIndex        =   26
         Top             =   1155
         Width           =   780
      End
      Begin VB.Label lblCompCode 
         AutoSize        =   -1  'True
         Caption         =   "公司別"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   270
         TabIndex        =   25
         Top             =   345
         Width           =   585
      End
      Begin VB.Label lblunti2 
         AutoSize        =   -1  'True
         Caption         =   "至"
         Height          =   195
         Left            =   2460
         TabIndex        =   24
         Top             =   750
         Width           =   195
      End
      Begin VB.Label lblCreateTime 
         AutoSize        =   -1  'True
         Caption         =   "產生時間"
         Height          =   195
         Left            =   270
         TabIndex        =   23
         Top             =   750
         Width           =   780
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "F2.確定"
      Default         =   -1  'True
      Height          =   405
      Left            =   3015
      TabIndex        =   19
      Top             =   6570
      Width           =   1275
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&X)"
      Height          =   405
      Left            =   6450
      TabIndex        =   20
      Top             =   6570
      Width           =   1275
   End
End
Attribute VB_Name = "frmSO3272A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private objAgency As Object
'Private objAgency As New clsInterface
Private objAction As Object
Private rsBankData As New ADODB.Recordset
Private rsDefVal As New ADODB.Recordset
Private strChoose As String
Private blnUnload As Boolean
Private strINIpath1 As String
Private strErrPath As String
Private strChooseString As String
Private Sub cmdCancel_Click()
  On Error GoTo ChkErr
    Unload Me
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "GetBankData")
End Sub

Private Function GetBankData() As Boolean
On Error GoTo ChkErr
    Dim strSQL  As String
       '' strSQL = "Select CodeNO,Description,PrgName,ActNo,BankId,CorpID,Spcno From CD018 where CodeNo=" & gilBank.GetCodeNo
        strSQL = "Select CodeNO,Description,PrgName,ActNo,BankId,CorpID,Spcno From CD018 where CodeNo  " & gmBank.GetQueryCode
        Set rsBankData = gcnGi.Execute(strSQL)
        If rsBankData.EOF Or rsBankData.BOF Then MsgBox "銀行資料有誤！請檢查銀行代碼檔！", vbExclamation, "訊息！": Exit Function
        GetBankData = True
    Exit Function
ChkErr:
    Call ErrSub(Me.Name, "GetBankData")
End Function

Private Sub cmdOk_Click()
  On Error GoTo ChkErr
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
      
      If Not ChkDTok Then Exit Sub
      If Not IsDataOk Then Exit Sub
      blnUnload = False
      If Not subChoose Then GoTo lExit
      Set objAgency = CreateObject("CreditCardOut.clsInterface")
      If gmdCustClass.GetQryStr <> "" Then
         strSQL = " From SO033 A,SO001 B Where " & strChoose
      Else
         strSQL = " From SO033 A Where " & strChoose
      End If
  ''    If Not GetBankData Then Exit Sub
      With objAgency
            .errPath = strErrPath
            .iniPath = strINIpath1
            .ChooseStr = strChoose
            '.ActNo = rsBankData("ActNo") & ""
            '.BankId = rsBankData("CodeNO") & ""
            .BankName = Me.Caption & " " & rsBankData("Description") & ""
            '.CorpID = rsBankData("CorpId") & ""
            .PrgName = rsBankData("PrgName") & ""
            .uSpcNO = rsBankData("SpcNO") & ""
            Set .Connection = gcnGi
            If Len(rsBankData("PrgName") & "") = 0 Then
                MsgBox "設定程式名稱無設或無使用權限！！", vbExclamation, "提示"
            Else
                'MsgBox ReadGICMIS1("ErrLogPath"), , "frmSO3273A"
                Set objAction = .InitObject(rsBankData("Prgname") & "")
                objAction.Action
                Set objAction = Nothing
            End If
      End With
      '判斷是否有資料
      If objAgency.unodata Then
        Set objAgency = Nothing
        Screen.MousePointer = vbDefault
        blnUnload = True
        On Error Resume Next
        gilCompCode.SetFocus
        Exit Sub
      Else
      '判斷是否執行成功
        If objAgency.uupdate Then
           Call subInsert
           blnUnload = True
        End If
      End If
lExit:
      blnUnload = True
      Set objAgency = Nothing
      DoEvents
      Screen.MousePointer = vbDefault
     
  Exit Sub
ChkErr:
    blnUnload = True
    If Err.Number = 20 Then Resume Next
    If Err.Number = 91 Then
       MsgBox "轉帳程式名稱錯誤或者該銀行沒有轉帳資料產生功能!", vbExclamation, "警告..."
       Exit Sub
    End If
    Call ErrSub(Me.Name, "cmdOk_Click")
End Sub

Private Function subChoose() As Boolean
  On Error GoTo ChkErr
  Dim strAddr As String

   subChoose = False
   strChoose = ""
       
    If gilCompCode.GetCodeNo <> "" Then subAnd ("A.CompCode=" & gilCompCode.GetCodeNo)
    If gilServiceType.GetCodeNo <> "" Then subAnd ("A.ServiceType='" & gilServiceType.GetCodeNo & "'")
    If gdaShouldDate1.GetValue <> "" Then subAnd ("A.ShouldDate >= To_Date(" & gdaShouldDate1.GetValue & ",'YYYYMMDD')")
    If gdaShouldDate2.GetValue <> "" Then subAnd ("A.ShouldDate < To_Date(" & gdaShouldDate2.GetValue & ",'YYYYMMDD') +1")
    If gdaCreateTime1.GetValue <> "" Then subAnd ("A.CreateTime >= To_Date(" & gdaCreateTime1.GetValue & ",'YYYYMMDD')")
    If gdaCreateTime2.GetValue <> "" Then subAnd ("A.CreateTime < To_Date(" & gdaCreateTime2.GetValue & ",'YYYYMMDD') +1")
    If gmdCMCode.GetQryStr <> "" Then subAnd ("A.CMCode " & gmdCMCode.GetQryStr)
    If gmdAreaCode.GetQryStr <> "" Then subAnd ("A.AreaCode " & gmdAreaCode.GetQryStr)
    If gmdServCode.GetQryStr <> "" Then subAnd ("A.ServCode " & gmdServCode.GetQryStr)
    If gimCustStatus.GetQryStr <> "" Then subAnd ("B.CustStatusCode " & gimCustStatus.GetQryStr)
    If gmdCustClass.GetQryStr <> "" Then subAnd ("B.ClassCode1 " & gmdCustClass.GetQryStr)
    If gmdClctEn.GetQryStr <> "" Then subAnd ("A.OldClctEn " & gmdClctEn.GetQryStr)
    If gmdBillType.GetQryStr <> "" Then
       subAnd ("SubStr(A.BillNo,7,1) In (" & gmdBillType.GetQueryCode & ")")
    Else
       subAnd ("SubStr(A.BillNo,7,1) In ('B','T')")
    End If
    If gimCreateEn.GetQryStr <> "" Then subAnd ("A.CreateEn " & gimCreateEn.GetQryStr)
    
    '產生大樓
    If optMduYes Then
       '產生一般戶
       If chkOther.Value = 1 Then
          If gmdMduid.GetQryStr <> "" Then subAnd ("(A.MduId " & gmdMduid.GetQryStr & " Or A.MduId Is Null) ")
       Else
       '不產生一般戶
          If gmdMduid.GetQryStr <> "" Then
             subAnd ("A.MduId " & gmdMduid.GetQryStr)
          Else
             subAnd ("A.MduId Is Not Null")
          End If
       End If
    '不產生大樓
    Else
       '產生一般戶
       If chkOther.Value = 1 Then
          If gmdMduid.GetQryStr <> "" Then
             subAnd ("(Not A.MduId " & gmdMduid.GetQryStr & " Or A.MduId Is Null) ")
          Else
             subAnd ("A.MduId Is Null")
          End If
       Else
       '不產生一般戶
          If gmdMduid.GetQryStr <> "" Then
             subAnd ("Not A.MduId " & gmdMduid.GetQryStr)
          Else
             MsgBox "此條件查無資料!!", vbExclamation, "訊息"
             Exit Function
          End If
       End If
    End If
    
'    If optChargeAddr.Value Then
'       strAddr = "收費地址"
'    Else
'       strAddr = "裝機地址"
'    End If
    subChoose = True
    strChooseString = "公司別　    :" & subSpace(gilCompCode.GetDescription) & ";" & _
                      "服務類別    :" & subSpace(gilServiceType.GetDescription) & ";" & _
                      "承辦銀行    :" & subSpace(gmBank.GetDispStr) & ";" & _
                      "收費日期　　:" & subSpace(gdaShouldDate1.GetOriginalValue) & "~" & subSpace(gdaShouldDate2.GetOriginalValue) & ";" & _
                      "收費方式    :" & subSpace(gmdCMCode.GetDispStr) & ";" & _
                      "行政區      :" & subSpace(gmdAreaCode.GetDispStr) & ";" & _
                      "服務區      :" & subSpace(gmdServCode.GetDispStr) & ";" & _
                      "客戶狀態    :" & subSpace(gimCustStatus.GetDispStr) & ";" & _
                      "客戶類別　　:" & subSpace(gmdCustClass.GetDispStr) & ";" & _
                      "收費員      :" & subSpace(gmdClctEn.GetDispStr) & ";" & _
                      "單據類別    :" & subSpace(gmdBillType.GetDispStr) & ";" & _
                      "產生人員    :" & subSpace(gimCreateEn.GetDispStr) & ";" & _
                      "大樓名稱    :" & subSpace(gmdMduid.GetDispStr)
  
  Exit Function
ChkErr:
    Call ErrSub(Me.Name, "subChoose")
End Function

Private Sub subAnd(strCh As String)
  On Error GoTo ChkErr
    '是空白加And
    If strChoose <> "" Then
       strChoose = strChoose & " And " & strCh
    Else
       strChoose = strCh
    End If
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "subAnd")
End Sub

Private Function AddChar(ByVal vStr As String, ByVal vChar As String, Optional OverWrite As Boolean = True) As String
On Error GoTo ChkErr
    Static strChar As String
        If OverWrite Then strChar = ""
        If InStr(1, vStr, vChar) > 0 Then
            strChar = strChar & IIf(strChar <> "", ",'", "'") & vChar & "'"
        End If
        AddChar = strChar
Exit Function
ChkErr:
    Call ErrSub(Me.Name, "ADDChar")
End Function

Private Sub Form_Activate()
On Error GoTo ChkErr
  Dim strSQL As String
  Dim rsTmp As New ADODB.Recordset
  
    strSQL = "Select PrgName From CD018 Where UPPER(SUBSTR(PrgName,1,10)) = '" & "CREDITCARD" & "'"
    If rsTmp.State = adStateOpen Then rsTmp.Close
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSQL, gcnGi, adOpenForwardOnly, adLockReadOnly
    If rsTmp.EOF Then
       MsgBox "請設定銀行代碼中轉帳程式名稱!", vbExclamation, "警告"
       Unload Me
    End If
    Screen.MousePointer = vbDefault
  Exit Sub
ChkErr:
  ErrSub Me.Name, "Form_Activate"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If Not ChkGiList(KeyCode) Then Exit Sub
    If KeyCode = vbKeyF2 Then cmdOK.Value = True: KeyCode = 0
    If KeyCode = vbKeyEscape Then Unload Me: KeyCode = 0
End Sub

Private Sub Form_Load()
    On Error GoTo ChkErr
        If Alfa2 Then
            Call GetGlobal
        End If
        blnUnload = True
        strINIpath1 = GetIniPath1
        strErrPath = ReadGICMIS1("ErrLogPath")
        Call subGil
        Call subGim
        Call getDefault
   Exit Sub
ChkErr:
   Call ErrSub(Me.Name, "Form_Load")
End Sub

Private Sub subGil()
  On Error GoTo ChkErr
  
        SetgiList gilCompCode, "CodeNo", "Description", "CD039", , , , , 3, 12
        SetgiList gilServiceType, "CodeNo", "Description", "CD046", , , , , 1, 12
       '' SetgiList gilBank, "CodeNo", "Description", "CD018", , , , , 3, 24, " Where UPPER(SUBSTR(PrgName,1,10)) = '" & "CREDITCARD" & "'"
   Exit Sub
ChkErr:
   ErrSub Me.Name, "subGil"
End Sub

Private Sub subGim()
  On Error GoTo ChkErr
  
    Call SetgiMulti(gmBank, "CodeNo", "Description", "CD018", "代碼", "名稱")
    Call SetgiMulti(gmdCMCode, "CodeNo", "Description", "CD031", "代碼", "名稱")
    Call SetgiMulti(gmdAreaCode, "CodeNo", "Description", "CD001", "代碼", "名稱")
    Call SetgiMulti(gmdServCode, "CodeNo", "Description", "CD002", "代碼", "名稱")
    Call SetgiMulti(gimCustStatus, "CodeNo", "Description", "CD035", "代碼", "名稱")
    Call SetgiMulti(gmdCustClass, "CodeNo", "Description", "CD004", "代碼", "名稱")
    Call SetgiMulti(gmdClctEn, "EmpNo", "EmpName", "CM003", "代碼", "名稱")
    Call SetgiMultiAddItem(gmdBillType, "B,T", "收費單,臨時收費單", "代碼", "名稱")
    Call SetgiMulti(gimCreateEn, "EmpNo", "EmpName", "CM003", "代碼", "名稱")
    Call SetgiMulti(gmdMduid, "Mduid", "Name", "SO017", "代碼", "名稱")
    gimCustStatus.SetQueryCode ("1")
    gmdBillType.SelectAll
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "subGim")
End Sub

Private Sub getDefault()
  On Error GoTo ChkErr
     gilCompCode.SetCodeNo garyGi(9)
     gilCompCode.Query_Description
  Exit Sub
ChkErr:
  ErrSub Me.Name, "getDefault"
End Sub

Private Function IsDataOk() As Boolean
On Error GoTo ChkErr
Dim strErrMsg As String

    IsDataOk = False
        If Len(gilCompCode.GetCodeNo & "") = 0 Then strErrMsg = "公司別": gilCompCode.GetCodeNo: GoTo Warning
        If Len(gilServiceType.GetCodeNo & "") = 0 Then strErrMsg = "服務類別": gilServiceType.GetCodeNo: GoTo Warning
       '' If gilBank.GetDescription = "" Then strErrMsg = "承辦銀行": gilBank.SetFocus: GoTo Warning
        If Len(gmBank.GetDispStr) > 0 Then strErrMsg = "承辦銀行": gmBank.SetFocus: GoTo Warning
        If gdaShouldDate1.GetValue = "" Then strErrMsg = "收費日期起始日": gdaShouldDate1.SetFocus: GoTo Warning
        If gdaShouldDate2.GetValue = "" Then strErrMsg = "收費日期截止日": gdaShouldDate2.SetFocus: GoTo Warning
        If DateDiff("d", gdaShouldDate1.Text, gdaShouldDate2.Text) < 0 Then MsgBox "收費截止日不得小於收費起始日！", vbExclamation, "訊息！": Exit Function
        
    IsDataOk = True
    Exit Function
Warning:
    Call MsgMustBe(strErrMsg)
    Exit Function
ChkErr:
    Call ErrSub(Me.Name, "IsDataOk")
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo ChkErr
    If blnUnload Then
       Set objAgency = Nothing
    Else
       Cancel = True
    End If
    Screen.MousePointer = vbDefault
    
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "Form_QueryUnload")
End Sub

Private Sub Frame4_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub gdaCreateTime1_GotFocus()
  On Error Resume Next
    If gdaCreateTime1.GetValue <> "" Then Exit Sub
    gdaCreateTime1.SetValue Date
End Sub

Private Sub gdaCreateTime1_Validate(Cancel As Boolean)
On Error GoTo ChkErr
     
   If Not IsDate(gdaCreateTime1.Text) Then Exit Sub
   If gdaCreateTime1.GetValue <> "" Then
      If gdaCreateTime2.GetValue = "" Then gdaCreateTime2.SetValue GetLastDate(gdaCreateTime1.GetValue(True))
   End If
Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "gdaCreateTime1_Validate")
End Sub

Private Sub gdaCreateTime2_GotFocus()
  On Error Resume Next
    If gdaCreateTime2.GetValue <> "" Then Exit Sub
    gdaCreateTime2.SetValue GetLastDate(Date)
End Sub

Private Sub gdaCreateTime2_Validate(Cancel As Boolean)
   On Error Resume Next
     If gdaCreateTime1.GetValue = "" Or gdaCreateTime2.GetValue = "" Then Exit Sub
     If Not IsDate(gdaCreateTime2.Text) Then Exit Sub
     If DateDiff("d", gdaCreateTime1.GetValue(True), gdaCreateTime2.GetValue(True)) < 0 Then
        MsgBox "產生時間截止日不可小於產生時間起始日!", vbExclamation, "警告"
        gdaCreateTime2.SetValue gdaCreateTime1.GetValue
        Cancel = True
     End If
End Sub

Private Sub gdaShouldDate1_GotFocus()
  On Error Resume Next
    If gdaShouldDate1.GetValue <> "" Then Exit Sub
    gdaShouldDate1.SetValue Date
End Sub

Private Sub gdaShouldDate1_Validate(Cancel As Boolean)
On Error GoTo ChkErr
     
   If Not IsDate(gdaShouldDate1.Text) Then Exit Sub
   If gdaShouldDate1.GetValue <> "" Then
      If gdaShouldDate2.GetValue = "" Then gdaShouldDate2.SetValue GetLastDate(gdaShouldDate1.GetValue(True))
   End If
Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "gdaShouldDate1_Validate")
End Sub

Private Sub gdaShouldDate2_GotFocus()
  On Error Resume Next
    If gdaShouldDate2.GetValue <> "" Then Exit Sub
    gdaShouldDate2.SetValue GetLastDate(Date)
End Sub

Private Sub gdaShouldDate2_Validate(Cancel As Boolean)
   On Error Resume Next
     If gdaShouldDate1.GetValue = "" Or gdaShouldDate2.GetValue = "" Then Exit Sub
     If Not IsDate(gdaShouldDate2.Text) Then Exit Sub
     If DateDiff("d", gdaShouldDate1.GetValue(True), gdaShouldDate2.GetValue(True)) < 0 Then
        MsgBox "應收截止日不可小於應收起始日!", vbExclamation, "警告"
        gdaShouldDate2.SetValue gdaShouldDate1.GetValue
        Cancel = True
     End If
End Sub

Private Sub subInsert()
On Error GoTo ChkErr
  Dim strSQL As String
  
    strSQL = "INSERT INTO SO097 (EMPNO,PRGNAME,RUNDATETIME,SECOND,SELECTION) " & _
                   "VALUES ('" & garyGi(0) & "','SO3273A',sysdate,'" & _
                   objAgency.utime & "秒','" & Replace(strChooseString, "'", "") & "')"
  
   gcnGi.Execute (strSQL)
   Exit Sub
ChkErr:
   ErrSub Me.Name, "subInsert"
End Sub

Private Sub gilCompCode_Change()
    On Error Resume Next
        Call SelectServType(gilCompCode.GetCodeNo, gilServiceType)
        gilServiceType.ListIndex = 1
        
       ''' GiListFilter gilBank, , gilCompCode.GetCodeNo
       '' gilBank.Filter = gilBank.Filter & IIf(gilBank.Filter = "", " Where ", " And ") & " UPPER(SUBSTR(PrgName,1,10)) = '" & "CREDITCARD" & "'"
        
        Call GiMultiFilter(gmBank, , gilCompCode.GetCodeNo)
        gmBank.Filter = gmBank.Filter & IIf(gmBank.Filter = "", " Where ", " And ") & " UPPER(SUBSTR(PrgName,1,10)) = '" & "CREDITCARD" & "'"
        Call GiMultiFilter(gmdAreaCode, , gilCompCode.GetCodeNo)
        Call GiMultiFilter(gmdMduid, , gilCompCode.GetCodeNo)
        Call GiMultiFilter(gmdServCode, , gilCompCode.GetCodeNo)

End Sub

Private Sub gilServiceType_Change()
    On Error Resume Next
        Call GiMultiFilter(gmdCMCode, gilServiceType.GetCodeNo)
        Call GiMultiFilter(gmdCustClass, gilServiceType.GetCodeNo)
End Sub


