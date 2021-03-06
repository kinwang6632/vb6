VERSION 5.00
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmACHTranRefer 
   Caption         =   "ACH轉帳扣款資料產生[ACHTranRefer]"
   ClientHeight    =   3765
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   8970
   Icon            =   "ACHTranRefer.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3765
   ScaleWidth      =   8970
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&X)"
      Height          =   375
      Left            =   7440
      TabIndex        =   11
      Top             =   3270
      Width           =   1275
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "F2.開始"
      Height          =   375
      Left            =   165
      TabIndex        =   10
      Top             =   3270
      Width           =   1275
   End
   Begin VB.Frame Frame1 
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
      Height          =   3060
      Left            =   165
      TabIndex        =   12
      Top             =   105
      Width           =   8595
      Begin VB.TextBox txtPutAccount 
         Height          =   300
         Left            =   5670
         MaxLength       =   14
         TabIndex        =   5
         Top             =   1140
         Width           =   1830
      End
      Begin VB.TextBox txtGotSecId 
         Height          =   300
         Left            =   2025
         MaxLength       =   7
         TabIndex        =   1
         Top             =   780
         Width           =   1830
      End
      Begin VB.Frame frmData 
         Caption         =   "資料檔位置"
         Height          =   1275
         HelpContextID   =   2
         Left            =   210
         TabIndex        =   13
         Top             =   1650
         Width           =   8160
         Begin VB.TextBox txtErrPath 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3015
            TabIndex        =   8
            Top             =   765
            Width           =   4095
         End
         Begin VB.CommandButton cmdDataPath 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   14.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   7320
            TabIndex        =   7
            Top             =   315
            Width           =   375
         End
         Begin VB.CommandButton cmdErrPath 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   14.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   7335
            TabIndex        =   9
            Top             =   780
            Width           =   375
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
            Left            =   3015
            TabIndex        =   6
            ToolTipText     =   "請輸入字檔之路徑及檔名！"
            Top             =   300
            Width           =   4095
         End
         Begin VB.Label lblErrpath 
            AutoSize        =   -1  'True
            Caption         =   "問題參考檔位置（路徑 + 名稱）"
            Height          =   195
            Left            =   510
            TabIndex        =   15
            Top             =   810
            Width           =   2460
         End
         Begin VB.Label lblDatapath 
            AutoSize        =   -1  'True
            Caption         =   "資料檔位置（路徑 +名稱)"
            Height          =   180
            Left            =   930
            TabIndex        =   14
            Top             =   375
            Width           =   1995
         End
      End
      Begin VB.TextBox txtInvoiceId 
         Height          =   345
         Left            =   2025
         MaxLength       =   10
         TabIndex        =   2
         Top             =   1140
         Width           =   1830
      End
      Begin VB.TextBox txtPutBankId 
         Height          =   300
         Left            =   5670
         MaxLength       =   7
         TabIndex        =   4
         Top             =   780
         Width           =   1830
      End
      Begin VB.TextBox txtSendSpcId 
         Height          =   300
         Left            =   2025
         MaxLength       =   7
         TabIndex        =   0
         Top             =   435
         Width           =   1830
      End
      Begin Gi_Date.GiDate gdaHandleDate 
         Height          =   345
         Left            =   5670
         TabIndex        =   3
         Top             =   435
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
      Begin VB.Label Label4 
         Caption         =   "注意：此為扣帳的前一個工作天"
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   6915
         TabIndex        =   22
         Top             =   375
         Width           =   1500
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         Caption         =   "接收單位代號"
         Height          =   285
         Left            =   510
         TabIndex        =   21
         Top             =   870
         Width           =   1425
      End
      Begin VB.Label Label2 
         Alignment       =   1  '靠右對齊
         Caption         =   "發動者帳號 (SO)"
         Height          =   210
         Left            =   4095
         TabIndex        =   20
         Top             =   1170
         Width           =   1335
      End
      Begin VB.Label Label3 
         Alignment       =   1  '靠右對齊
         Caption         =   "發動者統一編號 (SO)"
         Height          =   210
         Left            =   90
         TabIndex        =   19
         Top             =   1245
         Width           =   1845
      End
      Begin VB.Label Label8 
         Alignment       =   1  '靠右對齊
         Caption         =   "處理日期"
         Height          =   285
         Left            =   4365
         TabIndex        =   18
         Top             =   435
         Width           =   1050
      End
      Begin VB.Label Label5 
         Alignment       =   1  '靠右對齊
         Caption         =   "提出行代號"
         Height          =   285
         Left            =   4170
         TabIndex        =   17
         Top             =   780
         Width           =   1260
      End
      Begin VB.Label Label7 
         Alignment       =   1  '靠右對齊
         Caption         =   "發送單位代號"
         Height          =   285
         Left            =   210
         TabIndex        =   16
         Top             =   510
         Width           =   1725
      End
   End
   Begin MSComDlg.CommonDialog cnd1 
      Left            =   7755
      Top             =   570
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmACHTranRefer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private strPath As String
Private strChoose As String
Private strCompCode As String
Private strACHTNo As String
Private strBankHand As String
Private strAccNId As String
Private intPara24 As String
Private strPrgName As String
Private strBankId2 As String
Private strErrPath As String
Private strUpdEn As String
Private strUpdName As String
Private rsTmp As New ADODB.Recordset
Dim intAmount As Double
Private DataPath As TextStream
Private ErrPath As TextStream
Private FSO As New FileSystemObject
Private strPostUnit As String
Dim strUpdFields As String
Dim strUpdOldWhere As String
Dim strUpdWhere As String
Dim blnUpdUCCode As Boolean
Dim strUCCode As String
Dim strUCName As String
Private lngCount As Long
Private lngErrCount As Long
Private blnChkProc As Boolean
Private strBank As String
Private Const strMinRealStopDate As String = " MIN(DECODE(NVL(A.PAYKIND,0),1, " & _
                "DECODE(RealStopDate,NULL,TO_DATE('10000101','YYYYMMDD'),REALSTOPDATE)," & _
                "TO_DATE('10000101','YYYYMMDD'))) REALSTOPDATE "
Private Const strMinPayKind As String = " MAX(Nvl(A.PAYKIND,0)) PayKind "
Private strHavingOutZero As String
Private blnChkPayKind As Boolean
Public Enum ChkPayKind
    giSingle = 0
    giAll = 1
End Enum
Private sqlTmpViewName As String
Private isCrossCustCombine  As Boolean
Private intCrossCustCombine As Integer
Private strChooseCondition As String
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Private Sub cmdCancel_Click()
  On Error GoTo chkErr
    Unload Me
    
  Exit Sub
chkErr:
    Call ErrSub(Me.Name, "cmdCancel")
End Sub

Private Sub cmdDataPath_Click()
        On Error GoTo chkErr
        With cnd1
                .FileName = txtDataPath.Text
                .Filter = "文字檔 (*.txt;*.dat)|*.txt;*.dat|所有檔案 (*.*)|*.*"
                .InitDir = ReadGICMIS1("ErrLogPath")
                .Action = 1
                txtDataPath = .FileName
        End With
    Exit Sub
    
chkErr:
    Call ErrSub(Me.Name, "cmdDataPath_Click")
End Sub


Private Function IsDataOk() As Boolean
On Error GoTo chkErr
Dim strErrMsg As String

    IsDataOk = False
    If txtSendSpcId = "" Then strErrMsg = "發送單位代號": txtSendSpcId.SetFocus: GoTo Warning
    If txtGotSecId = "" Then strErrMsg = "接收單位代號": txtGotSecId.SetFocus: GoTo Warning
    If txtInvoiceId = "" Then strErrMsg = "發動者統一編號": txtInvoiceId.SetFocus: GoTo Warning
    If txtPutBankId = "" Then strErrMsg = "提出行代號": txtPutBankId.SetFocus: GoTo Warning
    If txtPutAccount = "" Then strErrMsg = "發動者帳號": txtPutAccount.SetFocus: GoTo Warning
    If gdaHandleDate.GetValue = "" Then strErrMsg = "處理日期": gdaHandleDate.SetFocus: GoTo Warning
    If txtDataPath.Text = "" Then strErrMsg = "資料檔位置": txtDataPath.SetFocus: GoTo Warning
    If txtErrPath.Text = "" Then strErrMsg = "資料檔位置": txtErrPath.SetFocus: GoTo Warning
    IsDataOk = True
    Exit Function
Warning:
    Call MsgMustBe(strErrMsg)
    Exit Function
chkErr:
    Call ErrSub(Me.Name, "IsDataOk")
End Function


Private Sub cmdErrPath_Click()
        On Error GoTo chkErr
        With cnd1
                .FileName = txtDataPath.Text
                .Filter = "文字檔 (*.txt;*.dat)|*.txt;*.dat|所有檔案 (*.*)|*.*"
                .InitDir = ReadGICMIS1("ErrLogPath")
                .Action = 1
                txtErrPath = .FileName
        End With
    Exit Sub
    
chkErr:
    Call ErrSub(Me.Name, "cmdErrPath_Click")
End Sub

Private Sub cmdOk_Click()
    On Error GoTo chkErr
    Dim S As String
    Dim Z As String
    Dim objChkErr As Object
    Dim blnErrNum As Boolean
    Dim LastTime As Single
    If Not IsDataOk Then Exit Sub
    S = Left(txtDataPath, InStrRev(txtDataPath, "\"))
    If S = "C:\" Or S = "c:\" Then
    Else
        If Not ChkDirExist(S) Then MsgBox "路徑 " & S & " 不存在!", vbExclamation: Exit Sub
    End If
    Z = Left(txtErrPath, InStrRev(txtErrPath, "\"))
    If Z = "C:\" Or Z = "c:\" Then
    Else
        If Not ChkDirExist(Z) Then MsgBox "路徑 " & Z & " 不存在!", vbExclamation: Exit Sub
    End If
    If OpenFile(DataPath, txtDataPath, True) = False Then Exit Sub
    If OpenFile(ErrPath, txtErrPath, True) = False Then Exit Sub
'    If OpenFile(txtDataPath, True) = False Then Exit Sub

    Call ScrToRcd
    '#7207
    NowTime = RightNow
    StartTime = timeGetTime / 1000
    If Not BeginTran Then
        CloseFS
        Screen.MousePointer = vbDefault
        cmdOK.Enabled = True: cmdCancel.Enabled = True
        Unload Me
        Exit Sub
    Else
        If lngErrCount > 0 Then
            Shell "Notepad " & txtErrPath.Text, vbNormalFocus
        End If
        '#7207
        If lngCount > 0 Then
            LastTime = timeGetTime / 1000
                LogToSO054 "SO3277A", _
                        garyGi(0), Format(NowTime, "YYYYMMDD HHMMSS"), SecToHMS(LastTime - StartTime), strChooseCondition, RightNow

        End If
        If blnChkProc And lngCount > 0 Then
            Set objChkErr = CreateObject("CheckMediaTEXT.clsExechkMediaText")
             With objChkErr
                 .ugiOpenFileType = 1 'Error文字檔開啟方式  0=Create  1=Append
                 .uClassName = "clsACHTranReferShow"  'Class名稱
                 .uPathFile = txtDataPath.Text  '資料路徑檔 路徑+檔名
                 .uErrPathFile = txtErrPath.Text  '錯誤檔案 路徑+檔名
                 '.uSetPathFile = Text3.Text '設定檔案 路徑+名稱
                 .uBankCode = strBank
                 .ugcnGi = gcnGi
                 .uOwner = GetOwner
                 .uCompCode = strCompCode
                 .CheckMediaTextShow
                 'blnReturn = .uReturnOK '回傳是否已經結束 TRUE=執行完畢沒有問題  FALSE=執行中或有問題(需要驗證)
                 blnErrNum = .ublnError '回傳是否有ERROR的資料 TRUE=有錯誤資料  FALSE=無錯誤資料
             End With
             If blnErrNum Then
                Shell "Notepad " & txtErrPath.Text, vbNormalFocus
             End If
        End If
    End If
    On Error Resume Next
    Set objChkErr = Nothing
    Exit Sub
chkErr:
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
        Call ErrSub("SO3277", "LogToSO054")
    End If
End Sub
Private Function GetRsTmp(ByRef rs As ADODB.Recordset) As Boolean
    On Error GoTo chkErr
    Dim strFrom  As String, strWhere As String
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim strCD013 As String
    strUpdFields = Empty
    strCD013 = "," & GetOwner & "CD013"
       'EMC
'        If Not GetRS(rsTmp, "Select A.BankCode as BankCode,A.BankName as BankName,A.RealDate " & _
'          " as TranDate,A.AccountNo,A.ShouldAmt as ShouldAmt,A.BillNo as BillNo_Old,A.BillNo as BillNo_New From " & strOwnerName & "So033 A Where RowId=''") Then Exit Function
            strFrom = " " & GetOwner & "SO033 A "
                If strChoose <> "" Then
                If InStr(1, strChoose, "C.") Then
                    strFrom = strFrom & "," & GetOwner & "So014 C "
                    
                    '#3417 要更新的TableName By Kin 2007/12/06
                    strUpdFields = GetOwner & "SO014 C"
                    
                    strWhere = strWhere & IIf(strWhere = "", "", " And ") & " A.AddrNo=C.AddrNo "
                End If
                If InStr(1, strChoose, "E.") Then
                    strFrom = strFrom & "," & GetOwner & "So002 E "
                    
                    '#3417 要更新的TableName By Kin 2007/12/06
                    strUpdFields = IIf(Len(strUpdFields) = 0, "", strUpdFields & ",") & GetOwner & "SO002 E"
                    
                    strWhere = strWhere & IIf(strWhere = "", "", " And ") & " A.ServiceType = E.ServiceType And A.CustId=E.CustId And A.CompCode=E.CompCode "
                End If
            End If
            'strWhere = strWhere & IIf(strWhere = "", "", " And ") & strChoose & " And A.CancelFlag=0 And A.UCCode is Not Null And A.ShouldAmt > 0"
            strWhere = strWhere & IIf(strWhere = "", "", " And ") & strChoose & " And A.CancelFlag=0 And A.UCCode is Not Null"
            '*********************************************************************************************************************************************************************
            '#3878 不以AddrNo Group By了 By Kin 2008/05/09
'            strSQL = "SELECT " & GetUseIndexStr("So033", "ShouldDate") & " A.CustId,A.MediaBillNo BillNo,A.AccountNo,A.BankCode,A.AddrNo, Sum(ShouldAmt) ShouldAmt  From " & _
'                        strFrom & " Where " & _
'                        strWhere & " Group By A.CustId,A.MediaBillNo,A.AccountNo,A.BankCode,A.AddrNo "
'            strUpdOldWhere = strWhere
            '#5056 如果CD013.Refno=3(櫃台已收)則不再出帳 By Kin 2009/04/30
            '#5218 櫃台已收要用Nvl的方式 By Kin 2009/08/05
            '#5564 增加參考號7,8與CD013.PAYOK=1都代表已收 By Kin 2010/05/17
            '#5683 增加RealStopDate 與 PayKind By Kin 2010/08/11
            '#6441 單張帳單合計總額<=0是否產生 By Kin 2013/05/14
            strSQL = "SELECT " & GetUseIndexStr("So033", "ShouldDate") & " A.CustId,A.MediaBillNo BillNo," & _
                        " A.AccountNo,A.BankCode," & strMinPayKind & " , " & strMinRealStopDate & " , " & _
                        " Sum(ShouldAmt) ShouldAmt  From " & _
                        strFrom & strCD013 & " Where " & _
                        strWhere & " And A.UCCode=CD013.CodeNo AND Nvl(CD013.REFNO,0) Not In(3,7) " & _
                        " AND NVL(CD013.PAYOK,0)=0 " & _
                        strHavingOutZero & _
                        " Group By A.CustId,A.MediaBillNo,A.AccountNo,A.BankCode "
             '#7179***********************************************************************************************************************************************
            If isCrossCustCombine Then
              
                intCrossCustCombine = 1
                sqlTmpViewName = GetTmpViewName
                '將套房與套房主客編找出來,如果無套房則主客編用custid代入
                strSQL = "Select A.*," & _
                            " (Case " & intCrossCustCombine & " When 1 then Nvl(SO001.AMduId,Null)   else Null End ) AMduId, " & _
                            " (Case " & intCrossCustCombine & " When 1 Then " & _
                                           " ( Case  When AMDUID Is Null Then A.CustId Else Nvl((Select MainCustId From " & GetOwner & "SO202 Where SO001.AMduId = SO202.MduId  ),-1)  End ) " & _
                            "  Else A.CustId  End ) MainCustId " & _
                            " From (" & strSQL & ") A," & GetOwner & "SO001 " & _
                    " Where A.CustId=SO001.CustId"
                strSQL = "Create View " & GetOwner & sqlTmpViewName & " As (" & strSQL & ")"
                gcnGi.Execute strSQL
                'If Not ExecuteCommand(strsql, cn, lngCount) Then Exit Function
                '判斷套房主客編是否有出現在SO033.CUSTID
'                strSQL = " select A.*,(select (case " & _
'                                      "                when amduid is null then 1 Else " & _
'                                                              " ( select  count(1) from " & GetOwner & sqlTmpViewName & _
'                                                              " Where A.MainCustId = Custid And A.AmdUid = AmduId And Nvl(A.BILLNO,'X')= Nvl(BILLNO,'X') " & _
'                                                               " And A.AccountNo = AccountNo ) " & _
'                                                      " End)  from dual) CustIdExistFlag  from " & GetOwner & sqlTmpViewName & " A"
                strSQL = " select A.*,(select (case " & _
                                      "                when amduid is null then 1 Else " & _
                                                              " ( select  count(1) from " & GetOwner & sqlTmpViewName & _
                                                              " Where A.MainCustId = Custid  ) " & _
                                                      " End)  from dual) CustIdExistFlag  from " & GetOwner & sqlTmpViewName & " A"
                strSQL = "Select Decode(CustIdExistFlag,0,Custid,MainCustId) CustId,Nvl(BILLNO,'X') BILLNO,AccountNo,Max(PayKind) PayKind,Min(RealStopDate) RealStopDate," & _
                                " Sum(ShouldAmt) ShouldAmt,AMDUID,CustIdExistFlag " & _
                        " From ( " & strSQL & " ) A Group By Decode(CustIdExistFlag,0,Custid,MainCustId),BILLNO,AccountNo,AMDUID,CustIdExistFlag "
               '將BankCode 找出來
                strSQL = "Select distinct A.*, " & _
                             " (select BankCode from " & GetOwner & sqlTmpViewName & _
                                    " where a.BillNo=BillNo And  A.AccountNo = AccountNo And rownum<=1) BankCode " & _
                            " From (" & strSQL & ") A"
                           
            End If
            '***********************************************************************************************************************************************

            strUpdOldWhere = strWhere
            '*********************************************************************************************************************************************************************
        If Not GetRS(rs, strSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic, , , True) Then Exit Function
        
    GetRsTmp = True
    Exit Function
chkErr:
     ErrSub Me.Name, "GetRsTmp"
End Function

Private Function GetRealDateTran(strRealDate As String) As String
    On Error Resume Next
        GetRealDateTran = Val(Replace(strRealDate, "/", "")) - 19110000
End Function

Private Function InsertHead(rs As ADODB.Recordset) As Boolean
    On Error GoTo chkErr
    Dim intReserve As Integer
    Dim strData As Variant
        '區別碼(錄別) , 資料代號, 處理日期, 處理時間, 發送單位代號, 接收單位代號, 備用
        strData = ""
        strData = strData & "BOF" & "ACHP01"
        '首錄別(1-3)資料代號(4-9)
        strData = strData & GetString(GetRealDateTran(gdaHandleDate.Text), 8, giRight, True)
        '處理日期(10-17)
        strData = strData & GetString(Format(Time, "HHMMSS"), 6, giLeft)
        '處理時間(18-23)
        strData = strData & GetString(txtSendSpcId, 7, giLeft)
        '發送單位代號(24-30)
        strData = strData & GetString(txtGotSecId, 7, giLeft)
        '接收單位代號(31-37)
        strData = strData & GetString("", 100, giLeft)
        
        '94/08/10 Jacky 改Jim 提 1715
        '備用(38-137)
        strData = strData & GetString(strPostUnit, 3, giLeft)
        '部門代碼(138-140)    CD068. PostUnit
        strData = strData & GetString("", 20, giLeft)
        '附註(141-160)   發動者自訂，如不用為空白
        WriteTextLine DataPath, strData
        InsertHead = True
    Exit Function
chkErr:
    ErrSub Me.Name, "InsertHead"

End Function

Private Function InsertFinal(rs As ADODB.Recordset) As Boolean
    On Error GoTo chkErr
    Dim intReserve As Integer
    Dim strData As Variant
        strData = "EOF" & "ACHP01"
        '尾錄別(1-3) , 資料代號(4-9)
        strData = strData & GetString(GetRealDateTran(gdaHandleDate.Text), 8, giRight, True)
        '處理日期(9-17)
        strData = strData & GetString(txtSendSpcId, 7, giLeft)
        '發送單位代號(18-24)
        strData = strData & GetString(txtGotSecId, 7, giLeft)
        '接收單位代號(25-31)
        strData = strData & GetString(rs.RecordCount, 8, giRight, True)
        '總筆數(32-39)
        strData = strData & GetString(intAmount, 16, giRight, True)
        '總金額(40-55)
        strData = strData & GetString("", 105, giLeft)
        '備用(56-160)
        
        WriteTextLine DataPath, strData
        InsertFinal = True
    Exit Function
chkErr:
    ErrSub Me.Name, "InsertFinal"

End Function
Private Function GetChkPayKind() As Boolean
    On Error Resume Next
    Dim aRet As Boolean
    aRet = False
    aRet = Val(gcnGi.Execute("SELECT NVL(PayNowChkStopDate,0) FROM " & GetOwner & "SO041")(0)) = 1
    GetChkPayKind = aRet
    
End Function
Private Function InsertDetail(rs As ADODB.Recordset) As Boolean
    On Error GoTo chkErr
    Dim strData As Variant
    Dim rsSo106 As New ADODB.Recordset
    Dim strSQL As String
    Dim intSeq As Integer
    Dim strACHCid As String
    Dim strErrMsg As String
    Dim strUpdTime As String
    Dim rsCustId As New ADODB.Recordset
    strUpdTime = GetDTString(RightNow)
    intAmount = 0: intSeq = 0
    If rs.RecordCount > 0 Then rs.MoveFirst
    Dim strPindKindName As String
    strPindKindName = GetRsValue("SELECT Description FROM " & GetOwner & "CD112 WHERE CODENO=1", gcnGi)
    
    If strPindKindName & "" = "" Then
        strPindKindName = "現付制"
    End If
    
   
    '#5843 增加檢核參數 By Kin 2010/11/24
    blnChkPayKind = GetChkPayKind
    Do While Not rs.EOF
         '#7179
        If isCrossCustCombine Then
            If Val(rs("CustId") & "") = -1 Then
                lngErrCount = lngErrCount + 1
                strErrMsg = "單據編號： " & rs("BillNO") & "套房找不到統收戶客戶編號，請設定統收戶客戶編號"
                WriteTextLine ErrPath, strErrMsg
                GoTo lNext
            End If
'            If Val(rs("CustIdExistFlag") & "") = 0 Then
'                lngErrCount = lngErrCount + 1
'                strErrMsg = "單據編號：[ " & rs("BillNO") & " ]　出帳資料沒有包含統收戶客戶編號：" & rs("CustId")
'                WriteTextLine ErrPath, strErrMsg
'                GoTo lNext
'            End If
            If rs("BillNo") = "X" Then
                lngErrCount = lngErrCount + 1
                strErrMsg = "客戶編號: " & rs("CustId") & " 媒體單號為空值！"
                WriteTextLine ErrPath, strErrMsg
                GoTo lNext
            End If
        End If
        '7179 客編要從新判斷，如果有多客編代表是統收戶，如果不是,則以查出的為主 By Kin 2016/03/30
        Dim aCustid As String
        aCustid = rs("CustId")
        If isCrossCustCombine Then
           If Not GetRS(rsCustId, "Select Distinct Custid From " & GetOwner & "SO033 Where MediaBillNo='" & rs("BillNO") & "'" & _
                                    " And AccountNo = '" & rs("AccountNo") & "' ", _
                        gcnGi, adUseClient, adOpenKeyset) Then
                        Exit Function
            Else
                If rsCustId.RecordCount = 1 Then
                    aCustid = rsCustId(0)
                End If
            End If
        End If
        strErrMsg = ""
        '#5683 如果收視截止日小於畫面條件不要出帳 By Kin 2010/08/12
        '#5843 如果設定不檢核收視截止日,則強迫出帳 By Kin 2010/11/24
        If (Not IsPayKindOK(rs, gdaHandleDate.GetValue, giAll)) _
            And (blnChkPayKind) Then
            strErrMsg = "收視截止日不正確 --> 收視截止日大於扣款處理日  客編 : " & aCustid & " 單據編號：" & rs("BillNo") & _
                    " 收視截止日：" & rs("RealStopDate") & " 繳付類別：" & strPindKindName
            WriteTextLine ErrPath, strErrMsg
            lngErrCount = lngErrCount + 1
            GoTo lNext
        End If
        strData = "N" & "SD"
        '交易型態(1-1) 交易類別(2-3)
        strData = strData & GetString(strACHTNo, 3, giLeft)
        strErrMsg = "交易代號"
        '交易代號(4-6)
        intSeq = intSeq + 1
        strData = strData & GetString(intSeq, 6, giRight, True)
        strErrMsg = "交易序號"
        '交易序號(7-12)
        strData = strData & GetString(txtPutBankId, 7, giLeft)
        strErrMsg = "提出行代號"
        '提出行代號(13-19)
        strData = strData & GetString(txtPutAccount, 14, giRight, True)
        strErrMsg = "發動者帳號"
        '發動者帳號(20-33)
'        strData = strData & GetString(Abs(rs("RealAmt")), 12, giright, True)
'        strData = strData & GetString("", 2, , True)
        strBankId2 = GetRsValue("Select BankId2 From " & GetOwner & "CD018 Where CodeNo='" & rs("BankCode") & "" & "'") & ""
        strData = strData & GetString(strBankId2, 7, giLeft)
        strErrMsg = "提回行代號"
        '提回行代號(34-40)
        strData = strData & GetString(rs("AccountNo") & "", 14, giRight, True)
        strErrMsg = "收受者帳號"
        '收受者帳號(41-54)
        strData = strData & GetString(rs("ShouldAmt") & "", 10, giRight, True)
        strErrMsg = "金額"
        '金額(55-64)
        strData = strData & GetString("00", 2, giLeft)
        strErrMsg = "退件理由代號"
        '退件理由代號 (65-66)
        strData = strData & "B" & GetString(txtInvoiceId, 10, giLeft)
        strErrMsg = "提示交換次序"
        '提示交換次序(67-67)發動者統一編號(68-77)
        strACHCid = strACHTNo & GetString(rs("Custid") & "", 8, giRight, True)
'        strsql = "SELECT AccountNameID,ACHCustId FROM " & GetOwner & "SO106 WHERE AccountID='" & rs("AccountNo") & _
                 "' And Bankcode='" & rs("BankCode") & "' And ACHCustId='" & strACHCid & "' and AuthorizeStatus=1"
        '題號2105 Jim提出 原ACHCustid 改由 achtno 欄位抓取 ..............by Crystal
        '題號2215 加上客編條件 ............by Crystal for Lydia
        strAccNId = "": strACHCid = ""
        strSQL = "SELECT AccountNameID,ACHCustId FROM " & GetOwner & "SO106 WHERE AccountID='" & rs("AccountNo") & _
                 "' And Bankcode='" & rs("BankCode") & "' And ACHTNO like '%" & strACHTNo & "%' And CustId='" & aCustid & "'" & _
                 " And AuthorizeStatus=1 and StopFlag=0 GROUP BY AccountNameID,ACHCustId"
              
        If Not GetRS(rsSo106, strSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic, , , True) Then Exit Function
        strErrMsg = "rsSo106"
        If Not rsSo106.EOF Then
            strAccNId = rsSo106("AccountNameID") & ""
            strACHCid = rsSo106("ACHCustid") & ""
        End If
        strData = strData & GetString(strAccNId, 10, giLeft)
        strErrMsg = "收受者統一編號"
        '收受者統一編號(78-87)
        strData = strData & GetString("", 6, giLeft)
        '上市上櫃公司代號(88-93)
        strData = strData & GetString("", 14, giLeft, True)
        '原提示交易日期,序號(94-107)
        strData = strData & GetString("", 1, giLeft)
        '退件必要欄位(107-108)出單不必印出
        strData = strData & GetString(strACHCid, 20, giLeft)
        strErrMsg = "SO106ACHCustid"
        'SO106ACHCustid(109-128)
        strData = strData & GetString(rs("BillNo") & "", 24, giLeft)
        strErrMsg = "媒體單號"
        '媒體單號(128-147)
        strData = strData & GetString("", 8, giLeft)
        strErrMsg = "備用"
        '備用(153-160)
        WriteTextLine DataPath, strData

        intAmount = intAmount + Abs(rs("ShouldAmt") & "")
        lngCount = lngCount + 1
        '*************************************************************************************************************************
        '#3417 成功匯出資料，要更新UCCode和UCName By Kin 2007/12/07
        If blnUpdUCCode Then
            '#3878 不以AddrNo為依據了 By Kin 2008/05/09
'            strUpdWhere = "Update " & GetOwner & "SO033 A Set UCCode=" & strUCCode & ",UCName='" & strUCName & "'" & _
'                        " Where Exists(Select A.* From " & strUpdFields & " Where " & strUpdOldWhere & _
'                                        " And A.MediaBillNo='" & rs("BillNo") & "'" & _
'                                        " And A.AccountNo='" & rs("AccountNo") & "'" & _
'                                        " And A.CUSTID=" & rs("CustId") & _
'                                        " And A.BankCode=" & rs("BankCode") & _
'                                        " And A.AddrNo=" & rs("AddrNo") & ")"

            '#3990 速度變慢，調整SQL語法 By Kin 2008/07/03

'            strUpdWhere = "Update " & GetOwner & "SO033 A Set UCCode=" & strUCCode & ",UCName='" & strUCName & "'" & _
'                        " Where Exists(Select A.* From " & strUpdFields & " Where " & strUpdOldWhere & _
'                                        " And A.MediaBillNo='" & rs("BillNo") & "'" & _
'                                        " And A.AccountNo='" & rs("AccountNo") & "'" & _
'                                        " And A.CUSTID=" & rs("CustId") & _
'                                        " And A.BankCode=" & rs("BankCode") & ")"
            '#4388 增加更新異動人員與異動時間 By Kin 2009/04/30
            '#5056 CD013.REFNO=3的不更新 By Kin 2009/04/30
            '#5218 櫃台已收要用Nvl的方式 By Kin 2009/08/05
            '#5564 增加參考號7,8與CD013.PAYOK=1都代表已收 By Kin 2010/05/17
'            strUpdWhere = "Update " & GetOwner & "SO033  Set UCCode=" & strUCCode & _
'                                    ",UCName='" & strUCName & "'" & _
'                                    ",UpdEN='" & garyGi(1) & "'" & _
'                                    ",UpdTime='" & strUpdTime & "' " & _
'                        " Where RowId In (Select A.RowId From " & GetOwner & "SO033 A" & IIf(strUpdFields <> Empty, "," & strUpdFields, "") & _
'                                        "," & GetOwner & "CD013 " & _
'                                        " Where " & strUpdOldWhere & _
'                                        " And A.MediaBillNo='" & rs("BillNo") & "'" & _
'                                        " And A.AccountNo='" & rs("AccountNo") & "'" & _
'                                        " And A.CUSTID=" & aCustid & _
'                                        " And A.BankCode=" & rs("BankCode") & _
'                                        " And A.UCCode=CD013.CodeNo And Nvl(CD013.REFNO,0) NOT IN(3,7) " & _
'                                        " AND NVL(CD013.PAYOK,0) =  0 )"
              strUpdWhere = "Update " & GetOwner & "SO033  Set UCCode=" & strUCCode & _
                                    ",UCName='" & strUCName & "'" & _
                                    ",UpdEN='" & garyGi(1) & "'" & _
                                    ",UpdTime='" & strUpdTime & "' " & _
                                     " Where RowId In (Select A.RowId From " & GetOwner & "SO033 A" & IIf(strUpdFields <> Empty, "," & strUpdFields, "") & _
                                        "," & GetOwner & "CD013 " & _
                                        " Where " & strUpdOldWhere & _
                                        " And A.MediaBillNo='" & rs("BillNo") & "'" & _
                                        " And A.AccountNo='" & rs("AccountNo") & "'" & _
                                        " And A.BankCode=" & rs("BankCode") & _
                                        " And A.UCCode=CD013.CodeNo And Nvl(CD013.REFNO,0) NOT IN(3,7) " & _
                                        " AND NVL(CD013.PAYOK,0) =  0 )"
            gcnGi.Execute strUpdWhere
        
        End If
lNext:
        '***************************************************************************************************************************
        rs.MoveNext
        DoEvents
    Loop
    
     InsertDetail = True
     On Error Resume Next
     CloseRecordset rsCustId
  Exit Function
chkErr:
'    cnn.RollbackTrans
    ErrSub Me.Name, "InsertDetail" & strErrMsg
End Function

Private Function BeginTran() As Boolean
    On Error GoTo chkErr
'    Dim rsTmp As New ADODB.Recordset
    Dim strBankId As String, strSQL As String
    Dim strOldBillNo As String
    Dim lngCountCustId As Long
    Dim lngPara24 As Long
    Dim lngTime As Long
    Dim rsUpd As New ADODB.Recordset
    Dim rsCD013 As New ADODB.Recordset
        lngTime = Timer
        lngCount = 0
        lngErrCount = 0
        '********************************************************************************************************************************************************
        '#3417 電子檔匯出時,要填入未收原因(RefNo=4) By Kin 2007/12/07
        If Not GetRS(rsCD013, "Select * From " & GetOwner & "CD013 Where RefNo=4 And StopFlag<>1 Order By CodeNo Desc", gcnGi, adUseClient, adOpenKeyset) Then Exit Function
        If rsCD013.EOF Then
            blnUpdUCCode = False
        Else
            blnUpdUCCode = True
            strUCCode = rsCD013("CodeNo") & ""
            strUCName = rsCD013("Description") & ""
        End If
        '*********************************************************************************************************************************************************
         '#7179 ********************************************************************************************************************
        isCrossCustCombine = Val(GetRsValue("Select Nvl(CrossCustCombine,0) From " & GetOwner & "SO041", gcnGi) & "")
        isCrossCustCombine = False
        If Len(sqlTmpViewName) > 0 Then
            gcnGi.Execute "Drop View " & GetOwner & sqlTmpViewName
            sqlTmpViewName = Empty
        End If
        '*****************************************************************************************************************************
        
'        取資料
        Screen.MousePointer = vbHourglass
        cmdOK.Enabled = False: cmdCancel.Enabled = False
        If Not GetRsTmp(rsTmp) Then Exit Function

        If rsTmp.RecordCount = 0 Then
            MsgBox "無資料！請重新操作！", vbInformation, "訊息！"
            GoTo KillFile
        Else
            If Not InsertHead(rsTmp) Then GoTo KillFile
            If Not InsertDetail(rsTmp) Then GoTo KillFile
            If Not InsertFinal(rsTmp) Then GoTo KillFile
            MsgBox "已完成資料筆數共" & lngCount & "筆," & vbCrLf & vbCrLf & _
           "問題筆數共" & lngErrCount & "筆," & vbCrLf & vbCrLf & _
           "共花費:" & (Timer - lngTime) \ 1 & "秒"
'            msgResult rsTmp.RecordCount, lngErrCount, lngTime       '顯示執行結果
        End If
        Screen.MousePointer = vbDefault
        cmdOK.Enabled = True: cmdCancel.Enabled = True
        Call CloseFS
       
        BeginTran = True
        On Error Resume Next
        '#5683 如果沒有產生資料把文字檔刪除
        If lngCount = 0 Then
            Kill txtDataPath.Text
        End If
        rsTmp.Close
        Set rsTmp = Nothing
        CloseRecordset rsCD013
    Exit Function
KillFile:
'    刪除檔案..
    CloseFS
    Exit Function
chkErr:
    Call ErrSub(Me.Name, "BeginTran")
End Function
Private Function IsPayKindOK(ByRef rsSource As ADODB.Recordset, _
    ByVal strRealStopDate As String, ByVal aChkPayKind As ChkPayKind, _
    Optional ByVal aMedia As Integer = 1) As Boolean
  On Error GoTo chkErr
    Dim aQry As String
    Dim aPayKind As Integer
    Dim aRsTmp As New ADODB.Recordset
    Dim aBillField As String
    Select Case aMedia
        Case 0
            aBillField = " BILLNO "
        Case 1
            aBillField = " MEDIABILLNO "
        Case Else
            aBillField = " CitibankATM "
    End Select
    If Not GetPaynowFlag Then
        IsPayKindOK = True
        GoTo lEnd
    End If
    Select Case aChkPayKind
        Case giSingle
            aQry = "SELECT MAX(NVL(A.PAYKIND,0)) PAYKIND, " & _
                        strMinRealStopDate & _
                        " FROM " & GetOwner & "SO033 A WHERE " & _
                        aBillField & "='" & rsSource("BILLNO") & "'"
            Set aRsTmp = gcnGi.Execute(aQry)
            aPayKind = Val(aRsTmp("PAYKIND") & "")
            If aPayKind = 0 Then
                IsPayKindOK = True
                GoTo lEnd
            Else
                If Val(Format(aRsTmp("REALSTOPDATE") & "", "YYYYMMDD")) > Val(strRealStopDate & "") Then
                    IsPayKindOK = False
                Else
                    IsPayKindOK = True
                End If
            End If
        Case giAll
            If Val(rsSource("PayKind") & "") = 0 Then
                IsPayKindOK = True
                GoTo lEnd
            End If
            If Val(Format(rsSource("REALSTOPDATE") & "", "YYYYMMDD")) > Val(strRealStopDate & "") Then
                IsPayKindOK = False
            Else
                IsPayKindOK = True
            End If
    End Select
lEnd:
    On Error Resume Next
    Call CloseRecordset(aRsTmp)
    Exit Function
chkErr:
  Call ErrSub(Me.Name, "IsPayKindOK")
End Function
Public Sub CloseFS()
    On Error Resume Next
    DataPath.Close
    ErrPath.Close
    Set FSO = Nothing
End Sub

Private Sub ScrToRcd()
    On Error GoTo chkErr
    Dim LogFile As TextStream
        Set LogFile = FSO.CreateTextFile(strErrPath & "\" & strPrgName & "Out.log", True)
        With LogFile
                .WriteLine (txtSendSpcId.Text)         '發送單位代號
                .WriteLine (txtGotSecId.Text)          '接收單位代號
                .WriteLine (txtInvoiceId.Text)         '發動者統一編號
'                .WriteLine (gdaHandleDate.Text)        '處理日期
                .WriteLine (txtPutBankId.Text)         '提出行代號
                .WriteLine (txtPutAccount.Text)        '發動者帳號
                .WriteLine (txtDataPath.Text)          '資料檔位置
                .WriteLine (txtErrPath.Text)           '問題檔位置
        End With
    Exit Sub
chkErr:
    Call ErrSub(Me.Name, "ScrToRcd")
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = vbKeyF2 Then
        If cmdOK.Enabled Then
            Call cmdOk_Click: KeyCode = 0
        End If
    End If
    If KeyCode = vbKeyEscape Then Unload Me: KeyCode = 0
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Screen.MousePointer = vbDefault
      Call InitData
      RcdToScr
End Sub

Private Sub RcdToScr()
    On Error GoTo chkErr
    Dim LogFile As TextStream
    If Dir(strErrPath & "\" & strPrgName & "Out.log") = "" Then Exit Sub
            On Error Resume Next
            Set LogFile = FSO.OpenTextFile(strErrPath & "\" & strPrgName & "Out.log", ForReading)
            With LogFile
                    If Not .AtEndOfStream Then txtSendSpcId.Text = .ReadLine & ""
                       '發送單位代號
                    If Not .AtEndOfStream Then txtGotSecId.Text = .ReadLine & ""
                       '接收單位代號
                    If Not .AtEndOfStream Then txtInvoiceId.Text = .ReadLine & ""
                       '發動者統一編號
'                    If Not .AtEndOfStream Then gdaHandleDate.Text = .ReadLine & ""
                       '處理日期
                    If Not .AtEndOfStream Then txtPutBankId.Text = .ReadLine & ""
                       '提出行代號
                    If Not .AtEndOfStream Then txtPutAccount.Text = .ReadLine & ""
                        '發動者帳號
                    If Not .AtEndOfStream Then txtDataPath.Text = .ReadLine & ""
                       '資料檔位置
                    If Not .AtEndOfStream Then txtErrPath.Text = .ReadLine & ""
                       '問題檔位置
            End With
    Exit Sub
chkErr:
    Call ErrSub(Me.Name, "RcdToScr")
End Sub

Private Function InitData() As Boolean
    On Error GoTo chkErr
        With frmSO3277A
             strChoose = .strChoose
             strCompCode = .gilCompCode.GetCodeNo
             strACHTNo = .txtACHTNo
'             strBankHand = .gilBankHand.GetCodeNo
'             intPara24 = .intPara24
             strPrgName = .strPrgName
             strHavingOutZero = ""
             If Not .uOutZero Then
                strHavingOutZero = " Having Sum(A.ShouldAmt)>0 "
             End If
        End With
        gdaHandleDate.SetValue ""
        strUpdEn = garyGi(0)
        strUpdName = garyGi(1)
        strErrPath = ReadGICMIS1("ErrLogPath")
        txtInvoiceId = GetRsValue("Select InvoiceId From " & GetOwner & "So041") & ""
        InitData = True
    Exit Function
chkErr:
    Call ErrSub(Me.Name, "InitData")
End Function
Public Property Let uChkProc(ByVal vData As Boolean)
    blnChkProc = vData
End Property
Public Property Let uPostUnit(ByVal vData As String)
    strPostUnit = vData
End Property

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
        DataPath.Close
        ErrPath.Close
        Set FSO = Nothing
        Call CloseRecordset(rsTmp)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
      '7179
    If Len(sqlTmpViewName) > 0 Then
        gcnGi.Execute "Drop View " & GetOwner & sqlTmpViewName
        sqlTmpViewName = Empty
    End If
    ReleaseCOM frmACHTranRefer

End Sub
Public Property Let uBank(ByVal vData As String)
  On Error Resume Next
  strBank = vData
End Property
Public Property Let uChooseString(ByVal vData As String)
    On Error Resume Next
    strChooseCondition = vData
End Property
Private Function GetPaynowFlag() As Boolean
  On Error Resume Next
     GetPaynowFlag = Val(GetRsValue("SELECT PaynowFlag FROM " & GetOwner & "SO041") & "") = 1
End Function
