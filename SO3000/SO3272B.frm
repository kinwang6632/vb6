VERSION 5.00
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{A0C51418-555F-11D3-BC10-0050BA01A3A9}#9.9#0"; "GiMulti.ocx"
Object = "{6DDA0569-6454-11D3-9612-0048544B09DF}#4.1#0"; "GiTime.ocx"
Object = "{074771B7-6ADE-11D5-B178-00485456B028}#2.1#0"; "csMulti.ocx"
Begin VB.Form frmSO3272B 
   BorderStyle     =   1  '單線固定
   Caption         =   "信用卡扣款資料產生作業 [SO3272B]"
   ClientHeight    =   8160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10740
   Icon            =   "SO3272B.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   10740
   StartUpPosition =   1  '所屬視窗中央
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&X)"
      Height          =   405
      Left            =   6390
      TabIndex        =   23
      Top             =   7695
      Width           =   1275
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "F2.確定"
      Default         =   -1  'True
      Height          =   405
      Left            =   2940
      TabIndex        =   22
      Top             =   7695
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
      Height          =   7500
      Left            =   -210
      TabIndex        =   24
      Top             =   105
      Width           =   10695
      Begin CS_Multi.CSmulti gimCitemCode 
         Height          =   315
         Left            =   210
         TabIndex        =   40
         Top             =   5430
         Width           =   10035
         _ExtentX        =   17701
         _ExtentY        =   556
         ButtonCaption   =   "收費項目"
      End
      Begin VB.CheckBox chkZero 
         Caption         =   "單張帳單合計總額<=0 是否要產生"
         Height          =   195
         Left            =   1200
         TabIndex        =   39
         Top             =   1440
         Width           =   3135
      End
      Begin VB.TextBox txtCustId 
         Height          =   285
         Left            =   2160
         TabIndex        =   35
         Top             =   6180
         Width           =   5925
      End
      Begin VB.CheckBox chkErr 
         Caption         =   "檢核產出電子檔"
         Height          =   315
         Left            =   2430
         TabIndex        =   34
         Top             =   1035
         Width           =   1605
      End
      Begin VB.ComboBox cboBillHeadFmt 
         Height          =   300
         ItemData        =   "SO3272B.frx":0442
         Left            =   5970
         List            =   "SO3272B.frx":0444
         Style           =   2  '單純下拉式
         TabIndex        =   4
         Top             =   600
         Width           =   4275
      End
      Begin VB.CheckBox chkOther 
         Caption         =   "包含一般戶"
         Height          =   225
         Left            =   1200
         TabIndex        =   9
         Top             =   1065
         Value           =   1  '核取
         Width           =   2145
      End
      Begin VB.Frame Frame2 
         Caption         =   " 大樓依據"
         Height          =   1005
         Left            =   210
         TabIndex        =   25
         Top             =   6420
         Width           =   10185
         Begin VB.OptionButton optMduNo 
            Caption         =   "&B.不產生"
            Height          =   285
            Left            =   1830
            TabIndex        =   21
            Top             =   660
            Width           =   1125
         End
         Begin VB.OptionButton optMduYes 
            Caption         =   "&A.要產生"
            Height          =   255
            Left            =   450
            TabIndex        =   20
            Top             =   690
            Value           =   -1  'True
            Width           =   1065
         End
         Begin CS_Multi.CSmulti gmdMduid 
            Height          =   405
            Left            =   150
            TabIndex        =   19
            Top             =   240
            Width           =   9855
            _ExtentX        =   17383
            _ExtentY        =   714
            ButtonCaption   =   "大 樓 名 稱"
         End
      End
      Begin Gi_Multi.GiMulti gmdBillType 
         Height          =   405
         Left            =   210
         TabIndex        =   17
         Top             =   4665
         Width           =   10035
         _ExtentX        =   17701
         _ExtentY        =   714
         ButtonCaption   =   "單 據 類 別"
         FontSize        =   9
         FontName        =   "新細明體"
      End
      Begin Gi_Multi.GiMulti gmdClctEn 
         Height          =   405
         Left            =   210
         TabIndex        =   16
         Top             =   3945
         Width           =   10035
         _ExtentX        =   17701
         _ExtentY        =   714
         ButtonCaption   =   "收    費    員"
         FontSize        =   9
         FontName        =   "新細明體"
      End
      Begin Gi_Multi.GiMulti gmdCustClass 
         Height          =   405
         Left            =   210
         TabIndex        =   15
         Top             =   3585
         Width           =   10035
         _ExtentX        =   17701
         _ExtentY        =   714
         ButtonCaption   =   "客 戶 類 別"
         FontSize        =   9
         FontName        =   "新細明體"
      End
      Begin Gi_Multi.GiMulti gmdServCode 
         Height          =   405
         Left            =   210
         TabIndex        =   13
         Top             =   2850
         Width           =   10035
         _ExtentX        =   17701
         _ExtentY        =   714
         ButtonCaption   =   "服    務    區"
         FontSize        =   9
         FontName        =   "新細明體"
      End
      Begin Gi_Multi.GiMulti gmdAreaCode 
         Height          =   405
         Left            =   210
         TabIndex        =   12
         Top             =   2490
         Width           =   10035
         _ExtentX        =   17701
         _ExtentY        =   714
         ButtonCaption   =   "行    政    區"
         FontSize        =   9
         FontName        =   "新細明體"
      End
      Begin Gi_Multi.GiMulti gmdCMCode 
         Height          =   405
         Left            =   210
         TabIndex        =   11
         Top             =   2145
         Width           =   10035
         _ExtentX        =   17701
         _ExtentY        =   714
         ButtonCaption   =   "收 費 方 式"
         FontSize        =   9
         FontName        =   "新細明體"
      End
      Begin Gi_Multi.GiMulti gimCustStatus 
         Height          =   405
         Left            =   210
         TabIndex        =   14
         Top             =   3225
         Width           =   10035
         _ExtentX        =   17701
         _ExtentY        =   714
         ButtonCaption   =   "客  戶 狀 態"
         FontSize        =   9
         FontName        =   "新細明體"
      End
      Begin Gi_Multi.GiMulti gimCreateEn 
         Height          =   405
         Left            =   210
         TabIndex        =   18
         Top             =   5040
         Width           =   10035
         _ExtentX        =   17701
         _ExtentY        =   714
         ButtonCaption   =   "產 生 人 員"
         FontSize        =   9
         FontName        =   "新細明體"
      End
      Begin prjGiList.GiList gilCompCode 
         Height          =   315
         Left            =   1200
         TabIndex        =   0
         Top             =   255
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
         Left            =   1200
         TabIndex        =   1
         Top             =   600
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
         Left            =   7155
         TabIndex        =   8
         Top             =   -165
         Visible         =   0   'False
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
         Left            =   5565
         TabIndex        =   7
         Top             =   -165
         Visible         =   0   'False
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
         Left            =   7545
         TabIndex        =   6
         Top             =   945
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
         Left            =   5985
         TabIndex        =   5
         Top             =   945
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
      Begin Gi_Multi.GiMulti gmBank 
         Height          =   405
         Left            =   210
         TabIndex        =   10
         Top             =   1800
         Width           =   10035
         _ExtentX        =   17701
         _ExtentY        =   714
         ButtonCaption   =   "承 辦 銀 行"
         FontSize        =   9
         FontName        =   "新細明體"
      End
      Begin Gi_Multi.GiMulti gmdOldClctEn 
         Height          =   360
         Left            =   210
         TabIndex        =   33
         Top             =   4320
         Width           =   10035
         _ExtentX        =   17701
         _ExtentY        =   635
         ButtonCaption   =   "原 收 費  員"
         FontSize        =   9
         FontName        =   "新細明體"
      End
      Begin Gi_Time.GiTime gitCreateTime1 
         Height          =   345
         Left            =   5970
         TabIndex        =   2
         Top             =   240
         Width           =   1935
         _ExtentX        =   3413
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
      Begin Gi_Time.GiTime gitCreateTime2 
         Height          =   345
         Left            =   8280
         TabIndex        =   3
         Top             =   240
         Width           =   1935
         _ExtentX        =   3413
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
      Begin Gi_Multi.GiMulti gmdPayType 
         Height          =   390
         Left            =   210
         TabIndex        =   38
         Top             =   5790
         Width           =   10035
         _ExtentX        =   17701
         _ExtentY        =   688
         ButtonCaption   =   "繳  付  類  別"
         FontSize        =   9
         FontName        =   "新細明體"
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "產生客編"
         Height          =   195
         Left            =   1260
         TabIndex        =   37
         Top             =   6240
         Width           =   855
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "請用「,」分開 Exp 67,82,199"
         Height          =   195
         Left            =   8130
         TabIndex        =   36
         Top             =   6240
         Width           =   2535
      End
      Begin VB.Label lblBillHeadFmt 
         AutoSize        =   -1  'True
         Caption         =   "多帳戶產生依據設定"
         Height          =   180
         Left            =   4290
         TabIndex        =   32
         Top             =   690
         Width           =   1620
      End
      Begin VB.Label lblCreateTime 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "產生時間"
         Height          =   195
         Left            =   5130
         TabIndex        =   31
         Top             =   315
         Width           =   780
      End
      Begin VB.Label lblunti2 
         AutoSize        =   -1  'True
         Caption         =   "至"
         Height          =   195
         Left            =   7980
         TabIndex        =   30
         Top             =   300
         Width           =   195
      End
      Begin VB.Label lblCompCode 
         AutoSize        =   -1  'True
         Caption         =   "公司別"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   585
         TabIndex        =   29
         Top             =   345
         Width           =   585
      End
      Begin VB.Label lblServiceType 
         AutoSize        =   -1  'True
         Caption         =   "服務類別"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   390
         TabIndex        =   28
         Top             =   645
         Width           =   780
      End
      Begin VB.Label lblReadAmt 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "應收日期範圍"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   4770
         TabIndex        =   27
         Top             =   1035
         Width           =   1170
      End
      Begin VB.Label lbluntil 
         AutoSize        =   -1  'True
         Caption         =   "至"
         Height          =   195
         Left            =   7215
         TabIndex        =   26
         Top             =   990
         Width           =   195
      End
   End
End
Attribute VB_Name = "frmSO3272B"
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
Private blnUnload As Boolean
Private strINIpath1 As String
Private strErrPath As String
Private strChooseString As String

Dim pChooseMultiAcc As String
Dim intPara24 As Integer  '' 判斷是否使用  多媒體   2003/08/05
Dim intPara23 As Integer
Dim rsBillHeadFmt As New ADODB.Recordset
Private blnPaynowFlag As Boolean
Private Sub cboBillHeadFmt_Click()
    On Error Resume Next
    Dim strCitemCode As String
    Dim rsCD019  As New ADODB.Recordset
    Dim strCD019SQL As String
        If rsBillHeadFmt.RecordCount = 0 Then Exit Sub
        rsBillHeadFmt.AbsolutePosition = cboBillHeadFmt.ListIndex + 1
'        If Not GetRS(rsCD019, "Select CodeNo From " & GetOwner & "CD019 Where CodeNo In (" & rsBillHeadFmt("CitemCodeStr") & ") And StopFlag=0") Then Exit Sub
        '#7049 改用CD068.CitemCode By Kin 2015/07/13
        strCD019SQL = "Select CodeNo From " & GetOwner & "CD019 " & _
                            " Where CodeNo In ( Select CitemCode From " & GetOwner & "CD068A " & _
                            "         Where Exists (Select * From " & GetOwner & "CD068 Where CD068.BillHeadFmt = CD068A.BillHeadFmt " & _
                                                    " And CD068.BillHeadFmt = '" & rsBillHeadFmt("BillHeadFmt") & "' )) And StopFlag <>1"
         If Not GetRS(rsCD019, strCD019SQL, gcnGi, adUseClient, adOpenKeyset) Then Exit Sub
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
Dim strWhere As String
 
   Dim strSQL  As String
   
   If Len(gmBank.GetQryStr) <> 0 Then
      strWhere = " AND CodeNo " & gmBank.GetQryStr
   End If
   
   strSQL = "Select CodeNO,Description,PrgName,ActNo,BankId,CorpID,Spcno From CD018 where " & " UPPER(SUBSTR(PrgName,1,10)) = '" & "CREDITCARD" & "'" & strWhere
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
    Dim strBank As String
    Dim LastTime As Single
    If Not ChkDTok Then Exit Sub
    If Not IsDataOk Then Exit Sub
    ReadyGoPrint
    blnUnload = False
    If Not subChoose Then GoTo lExit
    Set objAgency = CreateObject("CreditCardOut.clsInterface")
    If gmdCustClass.GetQryStr <> "" Then
        strSQL = " From SO033 A,SO001 B Where " & strChoose
    Else
        strSQL = " From SO033 A Where " & strChoose
    End If
    If InStr(1, gmBank.GetQueryCode, ",") > 0 Then
        strBank = Mid(gmBank.GetQueryCode, 1, InStr(1, gmBank.GetQueryCode, ",") - 1)
    Else
        strBank = gmBank.GetQueryCode
    End If

    If Not GetBankData Then Exit Sub
        With objAgency
            .errPath = strErrPath
            .iniPath = strINIpath1
            .ChooseStr = strChoose
            .uUPDEN = garyGi(1)
            .uUpdTime = GetDTString(RightNow)
            '.ActNo = rsBankData("ActNo") & ""
           ' .BankId = rsBankData("CodeNO") & ""
           '#5267 改用SO106,所以銀行別改用A.,不然語法很難串 By Kin 2010/04/21
            If Len(gmBank.GetQryStr) <> 0 Then
                .BankId = " A.BANKCODE " & gmBank.GetQryStr
            Else
                .BankId = " A.BANKCODE IN (" & gmBank.GetQueryCode & ")"
            End If
            .BankName = Me.Caption & " 花旗格式 "
            '.CorpID = rsBankData("CorpId") & ""
            
            .pTableOwnerName = GetOwner
            .prgName = rsBankData("PrgName") & ""
            .uSpcNO = rsBankData("SpcNO") & ""
            '#6441 張帳單合計總額<=0是否產生 By Kin 2013/05/13
            .uOutZero = chkZero.Value = 1
            Set .Connection = gcnGi
            If Len(rsBankData("PrgName") & "") = 0 Then
                MsgBox "設定程式名稱無設或無使用權限！！", vbExclamation, "提示"
            Else
                'MsgBox ReadGICMIS1("ErrLogPath"), , "frmSO3273A"
                Set objAction = .InitObject(rsBankData("Prgname") & "")
                
                Dim intPara24 As Integer
                
                 intPara24 = CInt(RPxx(gcnGi.Execute( _
                                                          "SELECT Para24 FROM " & GetOwner & "SO043 " & _
                                                          "WHERE " & _
                                                          "CompCode =" & gilCompCode.GetCodeNo & " AND ServiceType ='" & gilServiceType.GetCodeNo & "'" _
                                                          ).GetString))
               Call objAction.Action(pChooseMultiAcc, intPara24)
               
                Set objAction = Nothing
            End If
      End With
      '判斷是否有資料
      If objAgency.unodata Then
        Set objAgency = Nothing
         '7207 Add Data into SO054 Table By Kin 2016/03/08
        LastTime = timeGetTime / 1000
        LogToSO054 Replace(UCase(Screen.ActiveForm.Name), "FRM", ""), _
                        garyGi(0), Format(NowTime, "YYYYMMDD HHMMSS"), SecToHMS(LastTime - startTime), strChooseString, RightNow
        Screen.MousePointer = vbDefault
        blnUnload = True
        On Error Resume Next
        gilCompCode.SetFocus
        Exit Sub
      Else
      '判斷是否執行成功
        If objAgency.uupdate Then
            Call subInsert
            If chkErr.Value And (Not objAgency.unodata) Then
                Set objChkErr = CreateObject("CheckMediaTEXT.clsExechkMediaText")
                With objChkErr
                    .ugiOpenFileType = 1 'Error文字檔開啟方式  0=Create  1=Append
                    .uClassName = "CreditCardCitiBank_" & objAgency.uChkType 'Class名稱
                    .uPathFile = objAgency.uProcText '資料路徑檔 路徑+檔名
                    .uErrPathFile = objAgency.uProcErrText '錯誤檔案 路徑+檔名
                         '.uSetPathFile = Text3.Text '設定檔案 路徑+名稱
                    .uBankCode = strBank
                    .uOwner = GetOwner
                    .uCompCode = gilCompCode.GetCodeNo
                    .ugcnGi = gcnGi
                    .CheckMediaTextShow
                         'blnReturn = .uReturnOK '回傳是否已經結束 TRUE=執行完畢沒有問題  FALSE=執行中或有問題(需要驗證)
                    blnErrNum = .ublnError '回傳是否有ERROR的資料 TRUE=有錯誤資料  FALSE=無錯誤資料
                End With
                If blnErrNum Then
                    Shell "Notepad " & objAgency.uProcErrText, vbNormalFocus
                End If
            End If

           blnUnload = True
        End If
      End If
       '7207 Add Data into SO054 Table By Kin 2016/03/08
        LastTime = timeGetTime / 1000
        LogToSO054 Replace(UCase(Screen.ActiveForm.Name), "FRM", ""), _
                        garyGi(0), Format(NowTime, "YYYYMMDD HHMMSS"), SecToHMS(LastTime - startTime), strChooseString, RightNow
lExit:
      On Error Resume Next
      blnUnload = True
      Set objAgency = Nothing
      Set objChkErr = Nothing
      DoEvents
      Screen.MousePointer = vbDefault
     
  Exit Sub
chkErr:
    blnUnload = True
    If Err.Number = 20 Then Resume Next
    If Err.Number = 91 Then
       MsgBox "轉帳程式名稱錯誤或者該銀行沒有轉帳資料產生功能!", vbExclamation, "警告..."
       Exit Sub
    End If
    Call ErrSub(Me.Name, "cmdOk_Click")
End Sub
Private Sub LogToSO054(prgName As String, _
                      EmpNo As String, _
                      RunDateTime As String, _
                      SpendTime As String, _
                      Selection As String, _
                      Optional ProcessTime As String)
  On Error GoTo chkErr
  Dim strSQL As String
    strSQL = "INSERT INTO " & GetOwner & "SO054 " & _
                    "(PRGNAME,EMPNO,RUNDATETIME,SECOND,SELECTION,StopDateTime,ReportType) VALUES (" & _
                    GetNullString(prgName) & "," & GetNullString(EmpNo) & _
                    "," & GetNullString(RunDateTime, giDateV, giOracle, True) & "," & _
                    GetNullString(SpendTime) & "," & GetNullString(Selection) & "," & _
                    GetNullString(ProcessTime, giDateV, giOracle, True) & "," & _
                    "0)"
    strSQL = Replace(strSQL, Chr(0), "")
    gcnGi.Execute strSQL
  Exit Sub
chkErr:
    If Err.Number = -2147217865 Then
        MsgBox "SO054 Table 不存在 !! 無法寫入 ExportLog", vbInformation, "訊息"
    Else
        Call ErrSub(Me.Name, "LogToSO054")
    End If
End Sub
Private Function subChoose() As Boolean
  On Error GoTo chkErr
  Dim strAddr As String

   subChoose = False
   strChoose = ""
       
    If gilCompCode.GetCodeNo <> "" Then subAnd ("A.CompCode=" & gilCompCode.GetCodeNo)
    
    If intPara24 = 0 Then
           If gilServiceType.GetCodeNo <> "" Then subAnd ("A.ServiceType='" & gilServiceType.GetCodeNo & "'")
    End If
    
    If gdaShouldDate1.GetValue <> "" Then subAnd ("A.ShouldDate >= To_Date(" & gdaShouldDate1.GetValue & ",'YYYYMMDD')")
    If gdaShouldDate2.GetValue <> "" Then subAnd ("A.ShouldDate < To_Date(" & gdaShouldDate2.GetValue & ",'YYYYMMDD') +1")
    '*******************************************************************************************************************************
    '#2672 產生時間改為到時分秒 By Kin 2008/05/12
    'If gdaCreateTime1.GetValue <> "" Then subAnd ("A.CreateTime >= To_Date(" & gdaCreateTime1.GetValue & ",'YYYYMMDD')")
    'If gdaCreateTime2.GetValue <> "" Then subAnd ("A.CreateTime < To_Date(" & gdaCreateTime2.GetValue & ",'YYYYMMDD') +1")
    If gitCreateTime1.GetValue <> "" Then subAnd ("A.CreateTime>=To_Date(" & gitCreateTime1.GetValue & ",'YYYYMMDDHH24MISS')")
    If gitCreateTime2.GetValue <> "" Then subAnd ("A.CreateTime<To_Date(" & gitCreateTime2.GetValue & ",'YYYYMMDDHH24MISS')+1")
    '*******************************************************************************************************************************
    If gmdCMCode.GetQryStr <> "" Then subAnd ("A.CMCode " & gmdCMCode.GetQryStr)
    If gmdAreaCode.GetQryStr <> "" Then subAnd ("A.AreaCode " & gmdAreaCode.GetQryStr)
    If gmdServCode.GetQryStr <> "" Then subAnd ("A.ServCode " & gmdServCode.GetQryStr)
    '#1659  950517原"收費人員"改為"原收費人員"，另新增一"收費人員"

    If gmdClctEn.GetQryStr <> "" Then subAnd ("A.ClctEn " & gmdClctEn.GetQryStr)
    If gmdOldClctEn.GetQryStr <> "" Then subAnd ("A.OldClctEn " & gmdOldClctEn.GetQryStr)
    '''收費項目
    '#7049 改抓CD068A.CitemCode By Kin 2015/07/13
    'If gimCitemCode.GetQryStr <> "" Then subAnd ("A.CitemCode " & gimCitemCode.GetQryStr)
    If Len(gimCitemCode.GetQryStr & "") > 0 Then
        subAnd ("A.CitemCode In (Select CitemCode From " & GetOwner & "CD068A " & _
                                                                    " Where BillHeadFmt = '" & rsBillHeadFmt("BillHeadFmt") & "')")
    End If
    
    '#4388 增加客編條件 By Kin 2009/04/30
    If txtCustId.Text <> "" Then
        Call TimetxtCustId(txtCustId)
    End If
    '#5683 增加繳付類別 By Kin 2010/08/06
    If gmdPayType.GetQryStr <> "" Then subAnd ("A.PayKind " & gmdPayType.GetQryStr)
    If gmdBillType.GetQryStr <> "" Then
       subAnd ("SubStr(A.BillNo,7,1) In (" & gmdBillType.GetQueryCode & ")")
    Else
       subAnd ("SubStr(A.BillNo,7,1) In ('B','T','I','M','P')")
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
    '#6441 增加是否要判斷金額小於0 By Kin 2013/03/19
    'If chkZero.Value = 0 Then subAnd "A.ShouldAmt > 0"
     pChooseMultiAcc = strChoose
     
     ''  2003/08/15 將   CustStatusCode  改成SO002 的 D
     
    If gimCustStatus.GetQryStr <> "" Then subAnd ("D.CustStatusCode " & gimCustStatus.GetQryStr)
    If gmdCustClass.GetQryStr <> "" Then subAnd ("B.ClassCode1 " & gmdCustClass.GetQryStr)
    
    
'    If optChargeAddr.Value Then
'       strAddr = "收費地址"
'    Else
'       strAddr = "裝機地址"
'    End If
    subChoose = True
    strChooseString = "公司別　    :" & subSpace(gilCompCode.GetDescription) & ";" & _
                      "服務類別    :" & subSpace(gilServiceType.GetDescription) & ";" & _
                      "承辦銀行    :" & subSpace(gmBank.GetDispStr) & ";" & _
                      "收費日期　　:" & subSpace(gdaShouldDate1.GetValue(True)) & "~" & subSpace(gdaShouldDate2.GetValue(True)) & ";" & _
                      "收費方式    :" & subSpace(gmdCMCode.GetDispStr) & ";" & _
                      "行政區      :" & subSpace(gmdAreaCode.GetDispStr) & ";" & _
                      "服務區      :" & subSpace(gmdServCode.GetDispStr) & ";" & _
                      "客戶狀態    :" & subSpace(gimCustStatus.GetDispStr) & ";" & _
                      "客戶類別　　:" & subSpace(gmdCustClass.GetDispStr) & ";" & _
                      "收費員      :" & subSpace(gmdClctEn.GetDispStr) & ";" & _
                      "原收費員    :" & subSpace(gmdOldClctEn.GetDispStr) & ";" & _
                      "單據類別    :" & subSpace(gmdBillType.GetDispStr) & ";" & _
                      "產生人員    :" & subSpace(gimCreateEn.GetDispStr) & ";" & _
                      "大樓名稱    :" & subSpace(gmdMduid.GetDispStr)
  
  Exit Function
chkErr:
    Call ErrSub(Me.Name, "subChoose")
End Function

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

Private Sub CSmulti1_Click()

End Sub

Private Sub Form_Activate()
On Error GoTo chkErr
  Dim strSQL As String
  Dim rsTmp As New ADODB.Recordset
  
    strSQL = "Select PrgName From CD018 Where UPPER(SUBSTR(PrgName,1,10)) = '" & "CREDITCARD" & "'"
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
        '******************************************************
        '#2672 產生時間改為到時分秒 By Kin 2008/05/12
        gitCreateTime1.ShowSecond
        gitCreateTime2.ShowSecond
        '*****************************************************
        blnPaynowFlag = GetPaynowFlag
        Call subGil
        Call subGim
        Call getDefault
       intPara23 = Val(GetSystemParaItem("Para23", gilCompCode.GetCodeNo, gilServiceType.GetCodeNo, "SO043") & "")
       intPara24 = GetSystemParaItem("Para24", gilCompCode.GetCodeNo, gilServiceType.GetCodeNo, "SO043")
      If intPara24 = 1 Then
             cboBillHeadFmt.Enabled = True
            gilServiceType.Enabled = False
            lblServiceType.ForeColor = &H808080
            CboAddItem
            If intPara23 = 2 Then gimCitemCode.IsReadOnly = True Else gimCitemCode.IsReadOnly = False
      Else
           cboBillHeadFmt.Enabled = False
            lblServiceType.ForeColor = &HC0&
            gilServiceType.Enabled = True
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
  
        SetgiList gilCompCode, "CodeNo", "Description", "CD039", , , , , 3, 12
        SetgiList gilServiceType, "CodeNo", "Description", "CD046", , , , , 1, 12
        'SetgiList gilBank, "CodeNo", "Description", "CD018", , , , , 3, 24, " Where UPPER(SUBSTR(PrgName,1,10)) = '" & "CREDITCARD" & "'"
   Exit Sub
chkErr:
   ErrSub Me.Name, "subGil"
End Sub

Private Sub subGim()
  On Error GoTo chkErr
    Call SetgiMulti(gmBank, "CodeNo", "Description", "CD018", "代碼", "名稱", " Where UPPER(SUBSTR(PrgName,1,10)) = '" & "CREDITCARD" & "'")
    Call SetgiMulti(gmdCMCode, "CodeNo", "Description", "CD031", "代碼", "名稱")
    Call SetgiMulti(gmdAreaCode, "CodeNo", "Description", "CD001", "代碼", "名稱")
    Call SetgiMulti(gmdServCode, "CodeNo", "Description", "CD002", "代碼", "名稱")
    Call SetgiMulti(gimCustStatus, "CodeNo", "Description", "CD035", "代碼", "名稱")
    Call SetgiMulti(gmdCustClass, "CodeNo", "Description", "CD004", "代碼", "名稱")
    Call SetgiMulti(gmdClctEn, "EmpNo", "EmpName", "CM003", "代碼", "名稱")
    Call SetgiMulti(gmdOldClctEn, "EmpNo", "EmpName", "CM003", "代碼", "名稱")
'    Call SetgiMultiAddItem(gmdBillType, "B,T", "收費單,臨時收費單", "代碼", "名稱")
    '2627 單據類別開放可以使用所有單據類別---for jim 2006/07/27 Edit by Crystal
    Call SetgiMultiAddItem(gmdBillType, "B,T,I,M,P", "收費單,臨時收費單,裝機單,維修單,停拆移機單", "代碼", "名稱")
    Call SetgiMulti(gimCreateEn, "EmpNo", "EmpName", "CM003", "代碼", "名稱")
    '#5683 增加繳付類別 By Kin 2010/08/06
    Call SetgiMulti(gmdPayType, "CodeNo", "Description", "CD112", "代碼", "名稱")
    'Call SetgiMultiAddItem(gmdPayType, "0,1", "預付制,現付制", "代碼", "名稱")
     '收費項目
    Call SetgiMulti(gimCitemCode, "CodeNo", "Description", "CD019", "代碼", "名稱")
    
    Call SetgiMulti(gmdMduid, "Mduid", "Name", "SO017", "代碼", "名稱")
    gimCustStatus.SetQueryCode ("1")
    gmdBillType.SetQueryCode ("B,T")
    '#5925 繳費類別增加預設值 By Kin 2011/04/12
    If blnPaynowFlag Then
        With gmdPayType
             Select Case Val(GetRsValue("SELECT NVL(PayKindDefault,0) FROM " & GetOwner & "SO041", gcnGi) & "")
                 Case 0
                     .SelectAll
                 Case 1
                     .SetQueryCode "0"
                 Case 2
                     .SetQueryCode "1"
             End Select
         End With
    Else
        gmdPayType.Clear
        gmdPayType.Enabled = False
    End If
    Exit Sub
chkErr:
    Call ErrSub(Me.Name, "subGim")
End Sub

Private Sub getDefault()
  On Error GoTo chkErr
     gilCompCode.SetCodeNo garyGi(9)
     gilCompCode.Query_Description
  Exit Sub
chkErr:
  ErrSub Me.Name, "getDefault"
End Sub

Private Function IsDataOk() As Boolean
On Error GoTo chkErr
Dim strErrMsg As String

    IsDataOk = False
        If Len(gilCompCode.GetCodeNo & "") = 0 Then strErrMsg = "公司別": gilCompCode.GetCodeNo: GoTo Warning
        If Len(gilServiceType.GetCodeNo & "") = 0 Then strErrMsg = "服務類別": gilServiceType.GetCodeNo: GoTo Warning
     ''   If gilBank.GetDescription = "" Then strErrMsg = "承辦銀行": gilBank.SetFocus: GoTo Warning
        If gmBank.GetQueryCode = "" Then strErrMsg = "承辦銀行": gmBank.SetFocus: GoTo Warning
        If gdaShouldDate1.GetValue = "" Then strErrMsg = "收費日期起始日": gdaShouldDate1.SetFocus: GoTo Warning
        If gdaShouldDate2.GetValue = "" Then strErrMsg = "收費日期截止日": gdaShouldDate2.SetFocus: GoTo Warning
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
       Set objAgency = Nothing
    Else
       Cancel = True
    End If
    Screen.MousePointer = vbDefault
    
    Exit Sub
chkErr:
    Call ErrSub(Me.Name, "Form_QueryUnload")
End Sub

Private Sub Frame4_DragDrop(Source As Control, x As Single, y As Single)

End Sub

Private Sub gdaCreateTime1_GotFocus()
  On Error Resume Next
    If gdaCreateTime1.GetValue <> "" Then Exit Sub
    gdaCreateTime1.SetValue Date
End Sub

Private Sub gdaCreateTime1_Validate(Cancel As Boolean)
On Error GoTo chkErr
     
   If Not IsDate(gdaCreateTime1.Text) Then Exit Sub
   If gdaCreateTime1.GetValue <> "" Then
      If gdaCreateTime2.GetValue = "" Then gdaCreateTime2.SetValue GetLastDate(gdaCreateTime1.GetValue(True))
   End If
Exit Sub
chkErr:
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

Private Sub subInsert()
On Error GoTo chkErr
  Dim strSQL As String
  
    strSQL = "INSERT INTO SO097 (EMPNO,PRGNAME,RUNDATETIME,SECOND,SELECTION) " & _
                   "VALUES ('" & garyGi(0) & "','SO3273A',sysdate,'" & _
                   objAgency.utime & "秒','" & Replace(strChooseString, "'", "") & "')"
    gcnGi.Execute (strSQL)
   Exit Sub
chkErr:
   ErrSub Me.Name, "subInsert"
End Sub

Private Sub gilCompCode_Change()
    On Error Resume Next
        If ChgComp(gilCompCode, "SO3270", "SO3272") = False Then Exit Sub
        Call SelectServType(gilCompCode.GetCodeNo, gilServiceType)
        gilServiceType.ListIndex = 1
        
       '' GiListFilter gilBank, , gilCompCode.GetCodeNo
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

Private Sub gimCitemCodeX_Click()

End Sub

'#2672 產生時間改為到時分秒 By Kin 2008/05/12
Private Sub gitCreateTime1_GotFocus()
  On Error Resume Next
    If gitCreateTime1.GetValue <> "" Then Exit Sub
    gitCreateTime1.SetValue Now
End Sub
'#2672 產生時間改為到時分秒 By Kin 2008/05/12
Private Sub gitCreateTime1_Validate(Cancel As Boolean)
  On Error GoTo chkErr
   If Not IsDate(gitCreateTime1.Text) Then Exit Sub
   If gitCreateTime1.GetValue <> "" Then
      If gitCreateTime2.GetValue = "" Then gitCreateTime2.SetValue GetLastDate(gitCreateTime1.GetValue(True)) & "235959"
   End If
Exit Sub
chkErr:
    Call ErrSub(Me.Name, "gitCreateTime1_Validate")
End Sub

'#2672 產生時間改為到時分秒 By Kin 2008/05/12
Private Sub gitCreateTime2_GotFocus()
    If gitCreateTime2.GetValue <> "" Then Exit Sub
    gitCreateTime2.SetValue GetLastDate(Date) & "235959"
End Sub
'#2672 產生時間改為到時分秒 By Kin 2008/05/12
Private Sub gitCreateTime2_Validate(Cancel As Boolean)
  On Error Resume Next
    If gitCreateTime1.GetValue = "" Or gitCreateTime2.GetValue = "" Then Exit Sub
    If Not IsDate(gitCreateTime2.Text) Then Exit Sub
    If DateDiff("s", gitCreateTime1.GetValue(True), gitCreateTime2.GetValue(True)) < 0 Then
        MsgBox "產生時間截止日不可小於產生時間起始日!", vbExclamation, "警告"
        Cancel = True
    End If
End Sub

Private Sub txtCustId_KeyPress(keyAscii As Integer)
  On Error Resume Next
    If keyAscii >= 48 And keyAscii <= 57 Or keyAscii = 44 Or keyAscii = 8 Or keyAscii = 45 Then
        If keyAscii = 44 Or keyAscii = 45 Then
            If Asc(Right(" " & txtCustId.Text, 1)) = 44 Or Asc(Right(" " & txtCustId.Text, 1)) = 45 Or txtCustId = "" Then keyAscii = 9
        End If
        If Not ChkMaxLengthOK(txtCustId) Then
           If Not (keyAscii = 44 Or keyAscii = 8 Or keyAscii = 45) Then keyAscii = 9
        End If
    Else
        keyAscii = 9
    End If

End Sub
