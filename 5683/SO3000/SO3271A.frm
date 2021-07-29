VERSION 5.00
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{A0C51418-555F-11D3-BC10-0050BA01A3A9}#9.8#0"; "GiMulti.ocx"
Object = "{074771B7-6ADE-11D5-B178-00485456B028}#2.0#0"; "csMulti.ocx"
Begin VB.Form frmSO3271A 
   AutoRedraw      =   -1  'True
   Caption         =   "媒體轉帳資料產生作業 [SO3271A]"
   ClientHeight    =   7965
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10980
   BeginProperty Font 
      Name            =   "新細明體"
      Size            =   9.75
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SO3271A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7965
   ScaleWidth      =   10980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&X)"
      Height          =   405
      Left            =   6690
      TabIndex        =   19
      Top             =   7440
      Width           =   1275
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "F2.確定"
      Default         =   -1  'True
      Height          =   405
      Left            =   3270
      TabIndex        =   18
      Top             =   7455
      Width           =   1275
   End
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
      Height          =   7185
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   10695
      Begin VB.TextBox txtCustId 
         Height          =   360
         Left            =   1740
         TabIndex        =   14
         Top             =   5310
         Width           =   5925
      End
      Begin VB.CheckBox chkErr 
         Caption         =   "檢核產出電子檔"
         Height          =   315
         Left            =   270
         TabIndex        =   32
         Top             =   6750
         Width           =   1695
      End
      Begin VB.TextBox txtBID 
         Alignment       =   1  '靠右對齊
         Height          =   285
         Left            =   6000
         MaxLength       =   2
         TabIndex        =   31
         Top             =   420
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.ComboBox cboBillHeadFmt 
         Height          =   315
         Left            =   1950
         Style           =   2  '單純下拉式
         TabIndex        =   28
         Top             =   780
         Visible         =   0   'False
         Width           =   1725
      End
      Begin CS_Multi.CSmulti gmdClctEn 
         Height          =   345
         Left            =   210
         TabIndex        =   10
         Top             =   3420
         Width           =   10035
         _ExtentX        =   17701
         _ExtentY        =   609
         ButtonCaption   =   "收  費  人  員"
      End
      Begin CS_Multi.CSmulti gmdCustClass 
         Height          =   345
         Left            =   210
         TabIndex        =   9
         Top             =   3060
         Width           =   10035
         _ExtentX        =   17701
         _ExtentY        =   609
         ButtonCaption   =   "客  戶  類  別"
      End
      Begin Gi_Multi.GiMulti gmdBillType 
         Height          =   390
         Left            =   210
         TabIndex        =   12
         Top             =   4155
         Width           =   10035
         _ExtentX        =   17701
         _ExtentY        =   688
         ButtonCaption   =   "單  據  類  別"
      End
      Begin Gi_Multi.GiMulti gmdServCode 
         Height          =   375
         Left            =   210
         TabIndex        =   7
         Top             =   2340
         Width           =   10035
         _ExtentX        =   17701
         _ExtentY        =   661
         ButtonCaption   =   "服    務    區"
      End
      Begin Gi_Multi.GiMulti gmdAreaCode 
         Height          =   345
         Left            =   210
         TabIndex        =   6
         Top             =   1960
         Width           =   10035
         _ExtentX        =   17701
         _ExtentY        =   609
         ButtonCaption   =   "行    政    區"
      End
      Begin Gi_Multi.GiMulti gmdCMCode 
         Height          =   345
         Left            =   210
         TabIndex        =   5
         Top             =   1600
         Width           =   10035
         _ExtentX        =   17701
         _ExtentY        =   609
         ButtonCaption   =   "收 費 方 式"
      End
      Begin Gi_Date.GiDate gdaShouldDate2 
         Height          =   345
         Left            =   9210
         TabIndex        =   4
         Top             =   795
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
         Left            =   7650
         TabIndex        =   3
         Top             =   795
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
         Left            =   7650
         TabIndex        =   2
         Top             =   345
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
      Begin prjGiList.GiList gilCompCode 
         Height          =   315
         Left            =   1920
         TabIndex        =   0
         Top             =   390
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
         Left            =   1920
         TabIndex        =   1
         Top             =   780
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
      Begin Gi_Multi.GiMulti gimCustStatus 
         Height          =   465
         Left            =   210
         TabIndex        =   8
         Top             =   2700
         Width           =   10035
         _ExtentX        =   17701
         _ExtentY        =   820
         ButtonCaption   =   "客  戶  狀  態"
      End
      Begin VB.Frame Frame2 
         Caption         =   " 大樓依據"
         Height          =   1005
         Left            =   240
         TabIndex        =   21
         Top             =   5700
         Width           =   10185
         Begin VB.OptionButton optMduNo 
            Caption         =   "&B.不產生"
            Height          =   285
            Left            =   1770
            TabIndex        =   17
            Top             =   660
            Width           =   1125
         End
         Begin VB.OptionButton optMduYes 
            Caption         =   "&A.要產生"
            Height          =   255
            Left            =   480
            TabIndex        =   16
            Top             =   690
            Value           =   -1  'True
            Width           =   1065
         End
         Begin CS_Multi.CSmulti gmdMduid 
            Height          =   405
            Left            =   210
            TabIndex        =   15
            Top             =   240
            Width           =   9855
            _ExtentX        =   17383
            _ExtentY        =   714
            ButtonCaption   =   "大 樓 名 稱"
         End
      End
      Begin Gi_Multi.GiMulti gimCitemCode 
         Height          =   465
         Left            =   210
         TabIndex        =   27
         Top             =   1260
         Width           =   10035
         _ExtentX        =   17701
         _ExtentY        =   820
         ButtonCaption   =   "收 費 項 目"
      End
      Begin CS_Multi.CSmulti gmdOldClctEn 
         Height          =   345
         Left            =   210
         TabIndex        =   11
         Top             =   3780
         Width           =   10035
         _ExtentX        =   17701
         _ExtentY        =   609
         ButtonCaption   =   "原 收 費 人 員"
      End
      Begin Gi_Multi.GiMulti gimUCCode 
         Height          =   405
         Left            =   210
         TabIndex        =   13
         Top             =   4530
         Width           =   10035
         _ExtentX        =   17701
         _ExtentY        =   714
         ButtonCaption   =   "未收原因"
      End
      Begin Gi_Multi.GiMulti gmdPayType 
         Height          =   390
         Left            =   210
         TabIndex        =   35
         Top             =   4920
         Width           =   10035
         _ExtentX        =   17701
         _ExtentY        =   688
         ButtonCaption   =   "繳  付  類  別"
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "請用「,」分開 Exp 67,82,199"
         Height          =   195
         Left            =   7710
         TabIndex        =   34
         Top             =   5370
         Width           =   2535
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "產生客編"
         Height          =   195
         Left            =   900
         TabIndex        =   33
         Top             =   5400
         Width           =   855
      End
      Begin VB.Label lblBID 
         AutoSize        =   -1  'True
         Caption         =   "企業識別碼"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   4950
         TabIndex        =   30
         Top             =   460
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lblBillHeadFmt 
         AutoSize        =   -1  'True
         Caption         =   "多帳戶產生依據設定"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   840
         Width           =   1755
      End
      Begin VB.Label lblServiceType 
         AutoSize        =   -1  'True
         Caption         =   "服務類別"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   1080
         TabIndex        =   26
         Top             =   855
         Width           =   780
      End
      Begin VB.Label lblCompCode 
         AutoSize        =   -1  'True
         Caption         =   "公司別"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   1270
         TabIndex        =   25
         Top             =   460
         Width           =   585
      End
      Begin VB.Label lblBank 
         AutoSize        =   -1  'True
         Caption         =   "承辦銀行"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   6750
         TabIndex        =   24
         Top             =   460
         Width           =   780
      End
      Begin VB.Label lblReadAmt 
         AutoSize        =   -1  'True
         Caption         =   "應收日期範圍"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   6360
         TabIndex        =   23
         Top             =   855
         Width           =   1170
      End
      Begin VB.Label lbluntil 
         AutoSize        =   -1  'True
         Caption         =   "至"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   8850
         TabIndex        =   22
         Top             =   855
         Width           =   180
      End
   End
End
Attribute VB_Name = "frmSO3271A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private objMedia As Object
Private objAction As Object
Private rsBankData As New ADODB.Recordset
Private rsDefVal As New ADODB.Recordset
Private blnUnload As Boolean
Private strINIpath1 As String
Private strErrPath As String
Private strChooseString As String
Private strSaveChooseString As String
Private strChoose33 As String
Dim rsBillHeadFmt As New ADODB.Recordset
Dim intPara24 As Integer
Private strIDFile As String
Private txtStm As TextStream
Private FSO As New FileSystemObject
Private blnPaynowFlag As Boolean
Private Sub cboBillHeadFmt_Click()
    On Error Resume Next
    Dim strCitemCode As String
    Dim rsCD019  As New ADODB.Recordset
        If rsBillHeadFmt.RecordCount = 0 Then Exit Sub
        rsBillHeadFmt.AbsolutePosition = cboBillHeadFmt.ListIndex + 1
        If Not GetRS(rsCD019, "Select CodeNo From " & GetOwner & "CD019 Where CodeNo In (" & rsBillHeadFmt("CitemCodeStr") & ") And StopFlag=0") Then Exit Sub
        If rsCD019.EOF Then Exit Sub
        strCitemCode = rsCD019.GetString(, , , ",")
        strCitemCode = Left(strCitemCode, Len(strCitemCode) - 1)
        gimCitemCode.Filter = gimCitemCode.Filter
        gimCitemCode.SetQueryCode strCitemCode
        'gimCitemCode.ShowMulti
End Sub

Private Sub cmdCancel_Click()
  On Error GoTo chkErr
    Unload Me
  Exit Sub
chkErr:
    Call ErrSub(Me.Name, "GetBankData")
End Sub

Private Function GetBankData() As Boolean
On Error GoTo chkErr
    Dim strSQL  As String
        '#3527 多增加參考號 By Kin 2007/10/04
        strSQL = "Select CodeNO,Description,PrgName,ActNo,BankId,CorpID,Nvl(RefNo,0) RefNo From " & GetOwner & "CD018 where CodeNo=" & gilBank.GetCodeNo & " And StopFlag=0"
        Set rsBankData = gcnGi.Execute(strSQL)
        If rsBankData.EOF Or rsBankData.BOF Then MsgBox "銀行資料有誤！請檢查銀行代碼檔！", vbExclamation, "訊息！": Exit Function
        GetBankData = True
    Exit Function
chkErr:
    Call ErrSub(Me.Name, "GetBankData")
End Function

Private Sub cmdOK_Click()
  On Error GoTo chkErr
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim objChkErr As Object
    Dim blnErrNum As Boolean
    '************************************************
    '#3968 將企業識別碼儲存 By Kin 2008/07/10
    If txtBID.Visible Then
        Set txtStm = FSO.CreateTextFile(strIDFile, True)
        txtStm.WriteLine txtBID.Text
        txtStm.Close
    End If
    '************************************************
    If Not ChkDTok Then Exit Sub
    If Not IsDataOk Then Exit Sub
    blnUnload = False
    Call subChoose          '串Where 指令
    Set objMedia = CreateObject("MediaOut4.clsInterface")
    If garyGi(17) = 1 Then
        If gmdCustClass.GetQryStr <> "" Then
            strSQL = " From " & GetOwner & "SO033 A," & GetOwner & "SO001 B Where " & strChoose
        Else
            strSQL = " From " & GetOwner & "SO033 A Where " & strChoose
        End If
    End If
    If Not GetBankData Then Exit Sub
    With objMedia
        .errPath = strErrPath
        .iniPath = strINIpath1
        .ChooseStr = strChoose
        .ActNo = rsBankData("ActNo") & ""
        .BankId = rsBankData("CodeNo") & ""
        .BankName = Me.Caption & " " & rsBankData("Description") & ""
        .CorpID = rsBankData("CorpId") & ""
        .PrgName = rsBankData("PrgName") & ""
        .uCompCode = gilCompCode.GetCodeNo
        .uServiceType = gilServiceType.GetCodeNo
        .uStopDate = gdaShouldDate2.GetOriginalValue & ""
        .uChoose33 = strChoose33
        .uGetOwner = GetOwner
        .FlowId = garyGi(17)
        .uUPDEN = garyGi(1) '#4388 增加異動人員 By Kin 2009/04/29
        If txtBID.Visible Then
            .uBID = txtBID.Text
        Else
            .uBID = ""
        End If
        .uBillHeadFmt = Trim(cboBillHeadFmt.Text)
        
        '#3527 新增郵局格式，以參考號判別(使用Pos2) By Kin 2007/10/04
        .uRefNo = rsBankData("RefNo")
        Set .Connection = gcnGi
        Set .MDBConn = cnn
        If Len(rsBankData("PrgName") & "") = 0 Then
            MsgBox "設定程式名稱無設或無使用權限！！", vbExclamation, "提示"
        Else
            Set objAction = .InitObject(rsBankData("Prgname").Value & "")
            objAction.Action
            Set objAction = Nothing
        End If
    End With
    '判斷是否有資料
    If objMedia.unodata Then
        
        Set objMedia = Nothing
        Screen.MousePointer = vbDefault
        gilCompCode.SetFocus
        blnUnload = True
        Screen.MousePointer = vbDefault
        Exit Sub
    Else
    '判斷是否執行成功
        If objMedia.uupdate Then
            Call subInsert
            Call subPrint
            If chkErr.Value And (Not objMedia.unodata) Then
                Set objChkErr = CreateObject("CheckMediaTEXT.clsExechkMediaText")
                 With objChkErr
                     .ugiOpenFileType = 1 'Error文字檔開啟方式  0=Create  1=Append
                     .uClassName = rsBankData("Prgname").Value & "" 'Class名稱
                     .uPathFile = objMedia.uProcText '資料路徑檔 路徑+檔名
                     .uErrPathFile = objMedia.uProcErrText '錯誤檔案 路徑+檔名
                     '.uSetPathFile = Text3.Text '設定檔案 路徑+名稱
                     .uBankCode = gilBank.GetCodeNo
                     .uOwner = GetOwner
                     .uCompCode = gilCompCode.GetCodeNo
                     .ugcnGi = gcnGi
                     .CheckMediaTextShow
                     'blnReturn = .uReturnOK '回傳是否已經結束 TRUE=執行完畢沒有問題  FALSE=執行中或有問題(需要驗證)
                     blnErrNum = .ublnError '回傳是否有ERROR的資料 TRUE=有錯誤資料  FALSE=無錯誤資料
                 End With
                 If blnErrNum Then
                    Shell "Notepad " & objMedia.uProcErrText, vbNormalFocus
                 End If
            End If
        End If
    End If
    On Error Resume Next
    Set objMedia = Nothing
    Set objChkErr = Nothing
    Screen.MousePointer = vbDefault
    blnUnload = True
    'Unload Me
  Exit Sub
chkErr:
    blnUnload = True
    If Err.Number = 20 Then Resume Next
    If Err.Number = 91 Then
       MsgBox "轉帳程式名稱錯誤或者該銀行沒有轉帳資料產生功能!", vbExclamation, "警告..."
       Exit Sub
    End If
    
    Call ErrSub(Me.Name, "cmdOk_Click")
    Resume 0
End Sub

Private Sub subChoose()
  On Error GoTo chkErr
  Dim strAddr As String
  Dim rsLength As New ADODB.Recordset
     strChoose = ""
    If gilCompCode.GetCodeNo <> "" Then subAnd ("A.CompCode=" & gilCompCode.GetCodeNo)
    If gilServiceType.GetCodeNo <> "" Then subAnd ("A.ServiceType='" & gilServiceType.GetCodeNo & "'")
    If gilBank.GetCodeNo <> "" Then subAnd ("A.BankCode=" & gilBank.GetCodeNo)
    If gdaShouldDate1.GetValue <> "" Then subAnd ("A.ShouldDate >= To_Date(" & gdaShouldDate1.GetValue & ",'YYYYMMDD')")
    If gdaShouldDate2.GetValue <> "" Then subAnd ("A.ShouldDate < To_Date(" & gdaShouldDate2.GetValue & ",'YYYYMMDD') +1")
    If gmdCMCode.GetQryStr <> "" Then subAnd ("A.CMCode " & gmdCMCode.GetQryStr)
    '92/05/20 加收費項目
    If gimCitemCode.GetQryStr <> "" Then subAnd ("A.CitemCode " & gimCitemCode.GetQryStr)
    If gmdAreaCode.GetQryStr <> "" Then subAnd ("A.AreaCode " & gmdAreaCode.GetQryStr)
    If gmdServCode.GetQryStr <> "" Then subAnd ("A.ServCode " & gmdServCode.GetQryStr)
    '#1659  950517原"收費人員"改為"原收費人員"，另新增一"收費人員"
    If gmdClctEn.GetQryStr <> "" Then subAnd ("A.ClctEn " & gmdClctEn.GetQryStr)
    If gmdOldClctEn.GetQryStr <> "" Then subAnd ("A.OldClctEn " & gmdOldClctEn.GetQryStr)
    If gmdBillType.GetDispStr <> "" Then subAnd ("SubStr(A.BillNo,7,1) In (" & gmdBillType.GetQueryCode & ")")
    '#4388 增加未收原因 By Kin 2009/04/29
    If gimUCCode.GetQryStr <> "" Then subAnd ("A.UCCODE " & gimUCCode.GetQryStr)
    If optMduYes Then
       If gmdMduid.GetQryStr <> "" Then subAnd ("A.MduId " & gmdMduid.GetQryStr)
    Else
       If gmdMduid.GetQryStr <> "" Then subAnd ("Not A.MduId " & gmdMduid.GetQryStr)
    End If
    '#5683 增加繳款類別 By Kin 2010/08/03
    If gmdPayType.GetQryStr <> "" Then
        subAnd (" A.PayKind " & gmdPayType.GetQryStr)
    End If
    '#4388 增加客編條件 By Kin 2009/04/29
    If txtCustId.Text <> "" Then
        Call TimetxtCustId(txtCustId)
    End If
    strChoose33 = strChoose
    
    If gimCustStatus.GetQryStr <> "" Then subAnd ("C.CustStatusCode " & gimCustStatus.GetQryStr)
    If gmdCustClass.GetQryStr <> "" Then subAnd ("B.ClassCode1 " & gmdCustClass.GetQryStr)
    '條件太多會導致存入SO054時出錯,判斷SO054 Selection長度,最多只取Selection長度 For Karen 2007/01/03
    rsLength.Open "Select Selection from " & GetOwner & "SO054 Where Rownum=1 and 1<>0 ", gcnGi, adOpenForwardOnly, adLockReadOnly
    strChooseString = "公司別　    :" & subSpace(gilCompCode.GetDescription) & ";" & _
                      "承辦銀行    :" & subSpace(gilBank.GetDescription) & ";" & _
                      "服務類別    :" & subSpace(gilServiceType.GetDescription) & ";" & _
                      "收費日期　　:" & subSpace(gdaShouldDate1.GetValue(True)) & "~" & subSpace(gdaShouldDate2.GetValue(True)) & ";" & _
                      "收費方式    :" & subSpace(gmdCMCode.GetDispStr) & ";" & _
                      "行政區      :" & subSpace(gmdAreaCode.GetDispStr) & ";" & _
                      "服務區      :" & subSpace(gmdServCode.GetDispStr) & ";" & _
                      "客戶狀態    :" & subSpace(gimCustStatus.GetDispStr) & ";" & _
                      "客戶類別　　:" & subSpace(gmdCustClass.GetDispStr) & ";" & _
                      "收費員      :" & subSpace(gmdClctEn.GetDispStr) & ";" & _
                      "原收費員    :" & subSpace(gmdOldClctEn.GetDispStr) & ";" & _
                      "單據類別    :" & subSpace(gmdBillType.GetDispStr) & ";" & _
                      "大樓名稱    :" & subSpace(gmdMduid.GetDispStr) & ";" & _
                      "收費項目    :" & subSpace(gimCitemCode.GetDispStr) & ";"
    strChooseString = MidMbcs(strChooseString, 1, rsLength(0).DefinedSize)
    CloseRecordset rsLength
   '收費項目因為條件太多,導致存入SO097時會溢位,將存入的資料改為代碼 For Karen 2007/01/03
    rsLength.Open "Select Selection from " & GetOwner & "SO097 Where Rownum=1 and 1<>0 ", gcnGi, adOpenForwardOnly, adLockReadOnly
    strSaveChooseString = "公司別　    :" & subSpace(gilCompCode.GetDescription) & ";" & _
                          "承辦銀行    :" & subSpace(gilBank.GetDescription) & ";" & _
                          "服務類別    :" & subSpace(gilServiceType.GetDescription) & ";" & _
                          "收費日期　　:" & subSpace(gdaShouldDate1.GetValue(True)) & "~" & subSpace(gdaShouldDate2.GetValue(True)) & ";" & _
                          "收費方式    :" & subSpace(gmdCMCode.GetDispStr) & ";" & _
                          "行政區      :" & subSpace(gmdAreaCode.GetDispStr) & ";" & _
                          "服務區      :" & subSpace(gmdServCode.GetDispStr) & ";" & _
                          "客戶狀態    :" & subSpace(gimCustStatus.GetDispStr) & ";" & _
                          "客戶類別　　:" & subSpace(gmdCustClass.GetDispStr) & ";" & _
                          "收費員      :" & subSpace(gmdClctEn.GetDispStr) & ";" & _
                          "原收費員    :" & subSpace(gmdOldClctEn.GetDispStr) & ";" & _
                          "單據類別    :" & subSpace(gmdBillType.GetDispStr) & ";" & _
                          "大樓名稱    :" & subSpace(gmdMduid.GetDispStr) & ";" & _
                          "收費項目    :" & subSpace(gimCitemCode.GetQryStr) & ";"
    strSaveChooseString = MidMbcs(strSaveChooseString, 1, rsLength(0).DefinedSize)
    CloseRecordset rsLength
  Exit Sub
chkErr:
    CloseRecordset rsLength
    Call ErrSub(Me.Name, "subChoose")
End Sub

Private Sub subAnd(strCH As String)
  On Error GoTo chkErr
    '是空白加And
    If strChoose <> "" Then
       strChoose = strChoose & " And " & strCH
    Else
       strChoose = strCH
    End If
  Exit Sub
chkErr:
    Call ErrSub(Me.Name, "subAnd")
End Sub

Private Function AddChar(ByVal vStr As String, ByVal vChar As String, Optional OverWrite As Boolean = True) As String
On Error GoTo chkErr
    Static strChar As String
        If OverWrite Then strChar = ""
        If InStr(1, vStr, vChar) > 0 Then
            strChar = strChar & IIf(strChar <> "", ",'", "'") & vChar & "'"
        End If
        AddChar = strChar
Exit Function
chkErr:
    Call ErrSub(Me.Name, "ADDChar")
End Function

Private Sub Form_Activate()
On Error GoTo chkErr
  Dim strSQL As String
  Dim rsTmp As New ADODB.Recordset
    strSQL = "Select PrgName From " & GetOwner & "CD018 Where UPPER(SUBSTR(PrgName,1,5)) = '" & "MEDIA" & "'"
    If rsTmp.State = adStateOpen Then rsTmp.Close
'    rsTmp.CursorLocation = adUseClient
'    rsTmp.Open strSQL, gcnGi, adOpenForwardOnly, adLockReadOnly
    If Not GetRS(rsTmp, strSQL, gcnGi, adUseClient, adOpenForwardOnly, adLockReadOnly) Then Exit Sub
    If rsTmp.EOF Then
       MsgBox "請設定銀行代碼中轉帳程式名稱!", vbExclamation, "警告"
       Unload Me
    End If
    Screen.MousePointer = vbDefault

  Exit Sub
chkErr:
  ErrSub Me.Name, "Form_Activate"
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If Not ChkGiList(KeyCode) Then Exit Sub
    If KeyCode = vbKeyF2 Then cmdOK.Value = True: KeyCode = 0
    If KeyCode = vbKeyEscape Then Unload Me: KeyCode = 0
End Sub

Private Sub Form_Load()
    On Error GoTo chkErr
        If Alfa2 Then
            Call GetGlobal
        End If
        blnUnload = True
        strINIpath1 = GetIniPath1
        strErrPath = ReadGICMIS1("ErrLogPath")
        blnPaynowFlag = GetPaynowFlag
        Call subGil
        Call subGim
        Call getDefault
        Set cnn = GetTmpMdbCn
        '************************************************************************************************
        '#3968 將儲存的企業識別碼載入 By Kin 2008/07/10
        strIDFile = IIf(Right(strErrPath, 1) = "\", strErrPath, strErrPath & "\") & "SO3271ID.Txt"
        If FSO.FileExists(strIDFile) Then
            Set txtStm = FSO.OpenTextFile(strIDFile, ForReading)
            If Not txtStm.AtEndOfStream Then
                txtBID.Text = txtStm.ReadAll
            End If
            txtStm.Close
            Set txtStm = Nothing
        Else
            txtBID.Text = ""
        End If
'        Set txtStm = FSO.CreateTextFile(strIDFile, True)
        
        '************************************************************************************************
        CboAddItem
        If Not CreateTable() Then
            Exit Sub
        End If
        chkErr.Enabled = gcnGi.Execute("Select CheckText From " & GetOwner & "SO041  WHERE CompCode = " & garyGi(9))(0)
        If chkErr.Enabled Then chkErr.Value = 1 Else chkErr.Value = 0
Exit Sub
chkErr:
    Call ErrSub(Me.Name, "Form_Load")
End Sub
Private Sub CboAddItem()
    On Error Resume Next
        'If cboBillHeadFmt.ListCount > 0 Then Exit Sub
        If Not GetRS(rsBillHeadFmt, "Select BillHeadFmt,CitemCodeStr From " & GetOwner & "CD068") Then Exit Sub
        Do While Not rsBillHeadFmt.EOF
            cboBillHeadFmt.AddItem rsBillHeadFmt("BillHeadFmt") & ""
            rsBillHeadFmt.MoveNext
        Loop
End Sub
Private Sub subGil()
  On Error GoTo chkErr
        SetgiList gilCompCode, "CodeNo", "Description", "CD039", , , , , 3, 12, GetCompCodeFilter(gilCompCode)
        SetgiList gilServiceType, "CodeNo", "Description", "CD046", , , , , 1, 12
        SetgiList gilBank, "CodeNo", "Description", "CD018", , , , , 3, 24, " Where UPPER(SUBSTR(PrgName,1,5)) = '" & "MEDIA" & "'"
  Exit Sub
chkErr:
  ErrSub Me.Name, "subGil"
End Sub
Private Sub subGim()
  On Error GoTo chkErr
    Call SetgiMulti(gmdCMCode, "CodeNo", "Description", "CD031", "代碼", "名稱")
    Call SetgiMulti(gmdAreaCode, "CodeNo", "Description", "CD001", "代碼", "名稱")
    Call SetgiMulti(gmdServCode, "CodeNo", "Description", "CD002", "代碼", "名稱")
    Call SetgiMulti(gimCustStatus, "CodeNo", "Description", "CD035", "代碼", "名稱")
    Call SetgiMulti(gmdCustClass, "CodeNo", "Description", "CD004", "代碼", "名稱")
    Call SetgiMulti(gmdClctEn, "EmpNo", "EmpName", "CM003", "代碼", "名稱")
    Call SetgiMulti(gmdOldClctEn, "EmpNo", "EmpName", "CM003", "代碼", "名稱")
    '************************************************************************************************************
    '#4198 增加I、M、P單 By Kin 2008/11/03
    'Call SetgiMultiAddItem(gmdBillType, "B,T", "收費單,臨時收費單", "代碼", "名稱")
    Call SetgiMultiAddItem(gmdBillType, "B,T,I,M,P", "收費單,臨時收費單,裝機單,維修單,停拆移機單", "代碼", "名稱")
    '#5683 增加繳費類別 By Kin 2010/08/03
    Call SetgiMulti(gmdPayType, "CodeNo", "Description", "CD112", "代碼", "名稱")
    'Call SetgiMultiAddItem(gmdPayType, "0,1", "預付制,現付制", "代碼", "名稱")
    '************************************************************************************************************
    
    Call SetgiMulti(gmdMduid, "Mduid", "Name", "SO017", "代碼", "名稱")
    '收費項目
    SetgiMulti gimCitemCode, "CodeNo", "Description", "CD019", "代碼", "名稱", , True
    gimCustStatus.SetQueryCode ("1")
    
    '#4388 增加未收原因，但櫃台已收不列入 By Kin 2009/04/29
    '#5218 要用Nvl的方式 By Kin 2009/08/05
    SetgiMulti gimUCCode, "CodeNo", "Description", "CD013", "代碼", "名稱", " Where Nvl(REFNO,0) NOT IN(3,7) AND NVL(PayOK,0)=0 ", True
    
    '**********************************************
    '#4198 前端畫面預設B、T單 By Kin 2008/11/03
    With gmdBillType
        .SetDispStr "收費單,臨時收費單"
        .SetQueryCode "B,T"
    End With
    '#5683 增加繳費類別 By Kin 2010/08/03
    If blnPaynowFlag Then
        With gmdPayType
            '.SetDispStr "預付制"
            .SetQueryCode "0"
        End With
    Else
        gmdPayType.Clear
        gmdPayType.Enabled = False
    End If
    '**********************************************
Exit Sub
chkErr:
    Call ErrSub(Me.Name, "subGim")
End Sub

Private Sub getDefault()
  On Error GoTo chkErr
    Dim intPara23 As Integer
     gilCompCode.SetCodeNo garyGi(9)
     gilCompCode.Query_Description
        intPara24 = GetSystemParaItem("Para24", gilCompCode.GetCodeNo, gilServiceType.GetCodeNo, "SO043")
        intPara23 = Val(GetSystemParaItem("Para23", gilCompCode.GetCodeNo, gilServiceType.GetCodeNo, "SO043") & "")
        If intPara24 = 1 Then
            lblBillHeadFmt.Visible = True
            cboBillHeadFmt.Visible = True
            '  用媒體入帳才Lock
            If intPara23 = 2 Then gimCitemCode.Enabled = False
            gilServiceType.SetCodeNo ""
            gilServiceType.Query_Description
            gilServiceType.Enabled = False
            gilServiceType.Visible = False
            lblServicetype.Visible = False
        Else
            lblBillHeadFmt.Visible = False
        End If
  Exit Sub
chkErr:
  ErrSub Me.Name, "getDefault"
End Sub

Private Function IsDataOk() As Boolean
On Error GoTo chkErr
Dim strErrMsg As String

        IsDataOk = False
        If Len(gilCompCode.GetCodeNo & "") = 0 Then strErrMsg = "公司別": gilCompCode.SetFocus: GoTo Warning
        If intPara24 = 0 And Len(gilServiceType.GetCodeNo & "") = 0 Then strErrMsg = "服務類別": gilServiceType.SetFocus: GoTo Warning
        If Len(gilBank.GetCodeNo & "") = 0 Then strErrMsg = "承辦銀行": gilBank.SetFocus: GoTo Warning
        If Len(gdaShouldDate1.GetValue) = 0 Then strErrMsg = "收費日期起始日": gdaShouldDate1.SetFocus: GoTo Warning
        If Len(gdaShouldDate2.GetValue) = 0 Then strErrMsg = "收費日期截止日": gdaShouldDate2.SetFocus: GoTo Warning
        If DateDiff("d", gdaShouldDate1.Text, gdaShouldDate2.Text) < 0 Then MsgBox "收費截止日不得小於收費起始日！", vbExclamation, "訊息！": Exit Function
        
    IsDataOk = True
    Exit Function
Warning:
    Call MsgMustBe(strErrMsg)
    Exit Function
chkErr:
    Call ErrSub(Me.Name, "IsDataOk")
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo chkErr
    If blnUnload Then
       Set objMedia = Nothing
    Else
       Cancel = True
    End If
    On Error Resume Next
    cnn.Close
    Set cnn = Nothing
    'txtStm.Close
    Set txtStm = Nothing
    Set FSO = Nothing
    cnn.Close
    Set cnn = Nothing
    Exit Sub
chkErr:
    Call ErrSub(Me.Name, "Form_QueryUnload")
End Sub

Private Sub gdaShouldDate1_GotFocus()
On Error Resume Next
   
   If gdaShouldDate1.GetValue <> "" Then Exit Sub
   gdaShouldDate1.SetValue Date

End Sub

Private Sub gdaShouldDate1_Validate(Cancel As Boolean)
On Error GoTo chkErr
  
  If Not IsDate(gdaShouldDate1.Text) Then Exit Sub
  If gdaShouldDate1.GetValue <> "" Then
     If gdaShouldDate2.GetValue = "" Then gdaShouldDate2.SetValue GetLastDate(gdaShouldDate1.GetValue(True))
  End If
Exit Sub
chkErr:
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

Private Sub gilBank_Change()
  On Error GoTo chkErr
    Dim strPrgName As String
    lblBID.Visible = False
    txtBID.Visible = False
    '*************************************************************************************************************
    '#3968 判斷轉帳程式名稱為MEDIACHUNAN則出現企業識別碼供輸入 Kin 2008/07/10
    strPrgName = ""
    strPrgName = GetRsValue("Select PRGNAME From " & GetOwner & "CD018 Where CodeNo=" & gilBank.GetCodeNo, gcnGi)
    If UCase(strPrgName) = "MEDIACHUNAN" Then
        'txtBID = txtStm.ReadLine
        lblBID.Visible = True
        txtBID.Visible = True
    End If
    '*************************************************************************************************************
    Exit Sub
chkErr:
  Call ErrSub(Me.Name, "gilBank_Change")
End Sub

Private Sub gilCompCode_Change()
    On Error Resume Next
        garyGi(16) = gilCompCode.GetOwner
        Call subGil
        Call subGim
        Call SelectServType(gilCompCode.GetCodeNo, gilServiceType)
        gilServiceType.ListIndex = 1
        
'        GiListFilter gilBank, , gilCompCode.GetCodeNo
'        gilBank.Filter = gilBank.Filter & IIf(gilBank.Filter = "", " Where ", " And ") & " UPPER(SUBSTR(PrgName,1,5)) = '" & "MEDIA" & "' And ServiceType='" & gilServiceType.GetCodeNo & "' or ServiceType is null"
        
        Call GiMultiFilter(gmdAreaCode, , gilCompCode.GetCodeNo)
        Call GiMultiFilter(gmdMduid, , gilCompCode.GetCodeNo)
        Call GiMultiFilter(gmdServCode, , gilCompCode.GetCodeNo)
End Sub

Private Sub gilServiceType_Change()
    On Error Resume Next
        Call GiMultiFilter(gmdCMCode, gilServiceType.GetCodeNo)
        Call GiMultiFilter(gmdCustClass, gilServiceType.GetCodeNo)
        gmdCMCode.SetQueryCode ("4")

End Sub

Private Sub subInsert()
On Error GoTo chkErr
  Dim strSQL As String
    strSQL = "INSERT INTO " & GetOwner & "SO097 (EMPNO,PRGNAME,RUNDATETIME,SECOND,SELECTION) " & _
                   "VALUES ('" & garyGi(0) & "','SO3271A',sysdate,'" & _
                   objMedia.utime & "秒','" & Replace(strSaveChooseString, "'", "") & "')"
    gcnGi.Execute (strSQL)
   Exit Sub
chkErr:
   ErrSub Me.Name, "subInsert"
End Sub

Private Sub subPrint()
  On Error GoTo chkErr
  Dim rsTmp As New ADODB.Recordset
  Dim strFormula As String
  
    Set rsTmp = cnn.Execute("SELECT Count(*) as intCount From SO3271A")
    If rsTmp("intCount") = 0 Then
       MsgNoRcd
       SendSQL , , True
    Else
       ReadyGoPrint
       strSQL = "Select * From SO3271A"
       Call PrintRpt(GetPrinterName(5), "SO3271A.RPT", , Me.Caption, strSQL, strChooseString, , True, "Tmp0000.mdb", , , GiPaperPortrait)
    End If
    
    Set rsTmp = Nothing
    Screen.MousePointer = vbDefault
    
  Exit Sub
chkErr:
  Call ErrSub(Me.Name, "subPrint")
End Sub

Private Function CreateTable() As Boolean
  On Error Resume Next
    cnn.Execute "Drop Table SO3271A"
  On Error GoTo chkErr
    cnn.Execute "Create Table SO3271A (AccountNo text(20),TransDate Date,ShouldAmt Long Default 0,BillNo_Old Text(20),BillNo_New Text(20))"
    CreateTable = True
    Exit Function
chkErr:
    Call ErrSub(Me.Name, "CreateTable")
End Function

Private Sub txtCustId_KeyPress(KeyAscii As Integer)
  On Error Resume Next
    If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 44 Or KeyAscii = 8 Or KeyAscii = 45 Then
        If KeyAscii = 44 Or KeyAscii = 45 Then
            If Asc(Right(" " & txtCustId.Text, 1)) = 44 Or Asc(Right(" " & txtCustId.Text, 1)) = 45 Or txtCustId = "" Then KeyAscii = 9
        End If
        If Not ChkMaxLengthOK(txtCustId) Then
           If Not (KeyAscii = 44 Or KeyAscii = 8 Or KeyAscii = 45) Then KeyAscii = 9
        End If
    Else
        KeyAscii = 9
    End If
End Sub
