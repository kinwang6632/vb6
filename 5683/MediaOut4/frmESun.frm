VERSION 5.00
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmESun 
   ClientHeight    =   6810
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11190
   Icon            =   "frmESun.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6810
   ScaleWidth      =   11190
   StartUpPosition =   1  '所屬視窗中央
   Begin MSComDlg.CommonDialog cnd1 
      Left            =   10530
      Top             =   6240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "輸出值"
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
      Height          =   5745
      Left            =   240
      TabIndex        =   11
      Top             =   180
      Width           =   10695
      Begin VB.TextBox txtBankId 
         Height          =   375
         Left            =   6660
         MaxLength       =   4
         TabIndex        =   2
         Top             =   480
         Width           =   1125
      End
      Begin VB.Frame frmData 
         Caption         =   "資料檔位置"
         Height          =   1875
         Left            =   450
         TabIndex        =   12
         Top             =   1680
         Width           =   9075
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
            Height          =   375
            Left            =   7530
            TabIndex        =   24
            Top             =   1320
            Width           =   375
         End
         Begin VB.CommandButton cmdMemoPath 
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
            Height          =   375
            Left            =   7530
            TabIndex        =   23
            Top             =   810
            Width           =   375
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
            Height          =   375
            Left            =   7530
            TabIndex        =   22
            Top             =   270
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
            Height          =   375
            Left            =   3240
            TabIndex        =   4
            ToolTipText     =   "請輸入字檔之路徑及檔名！"
            Top             =   270
            Width           =   4095
         End
         Begin VB.TextBox txtMemoPath 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3240
            TabIndex        =   5
            Top             =   795
            Width           =   4095
         End
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
            Height          =   375
            Left            =   3240
            TabIndex        =   6
            Top             =   1320
            Width           =   4095
         End
         Begin VB.Label lblDatapath 
            AutoSize        =   -1  'True
            Caption         =   "資料檔位置（路徑 + 名稱）"
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
            Left            =   330
            TabIndex        =   15
            Top             =   390
            Width           =   2340
         End
         Begin VB.Label lblMemopath 
            AutoSize        =   -1  'True
            Caption         =   "備註檔位置（路徑 + 名稱）"
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
            Left            =   330
            TabIndex        =   14
            Top             =   885
            Width           =   2340
         End
         Begin VB.Label lblErrpath 
            AutoSize        =   -1  'True
            Caption         =   "問題參考檔位置（路徑 + 名稱）"
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
            Left            =   330
            TabIndex        =   13
            Top             =   1380
            Width           =   2730
         End
      End
      Begin VB.TextBox txtMemo1 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1755
         Left            =   1800
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   3690
         Width           =   3255
      End
      Begin VB.TextBox txtMemo2 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1755
         Left            =   6510
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   3720
         Width           =   3015
      End
      Begin VB.TextBox txtCorpId 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6660
         MaxLength       =   10
         TabIndex        =   3
         Top             =   1050
         Width           =   1125
      End
      Begin Gi_Date.GiDate gdaTransfer 
         Height          =   375
         Left            =   2280
         TabIndex        =   1
         Top             =   1050
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
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
      Begin Gi_Date.GiDate gdaAgency 
         Height          =   375
         Left            =   2280
         TabIndex        =   0
         Top             =   480
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
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
      Begin VB.Label lblagency 
         AutoSize        =   -1  'True
         Caption         =   "臨櫃代收繳費期限"
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
         Left            =   540
         TabIndex        =   21
         Top             =   600
         Width           =   1560
      End
      Begin VB.Label lblBankId 
         AutoSize        =   -1  'True
         Caption         =   "主辦行"
         Height          =   180
         Left            =   5640
         TabIndex        =   20
         Top             =   600
         Width           =   540
      End
      Begin VB.Label lblTransfer 
         AutoSize        =   -1  'True
         Caption         =   "自動轉帳扣款日期"
         Height          =   180
         Left            =   540
         TabIndex        =   19
         Top             =   1170
         Width           =   1440
      End
      Begin VB.Label lblMemo1 
         AutoSize        =   -1  'True
         Caption         =   "收費單備註一："
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
         Left            =   480
         TabIndex        =   18
         Top             =   3750
         Width           =   1365
      End
      Begin VB.Label lblMemo2 
         AutoSize        =   -1  'True
         Caption         =   "收費單備註二："
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
         Left            =   5160
         TabIndex        =   17
         Top             =   3750
         Width           =   1365
      End
      Begin VB.Label lblCorpId 
         AutoSize        =   -1  'True
         Caption         =   "營利編號"
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
         Left            =   5640
         TabIndex        =   16
         Top             =   1170
         Width           =   780
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "F2.開始"
      Height          =   405
      Left            =   3570
      TabIndex        =   9
      Top             =   6180
      Width           =   1275
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&X)"
      Height          =   405
      Left            =   5940
      TabIndex        =   10
      Top             =   6180
      Width           =   1275
   End
End
Attribute VB_Name = "frmESun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As New ADODB.Connection
Private rsDefTmp As New ADODB.Recordset

'入帳帳號
Private strActNo As String
'事業單位代號
Private strBankId As String
'事業單位名稱
Private strBankName As String
'收件單位代號
Private strCorpID As String
'Where 條件
Private strChoose As String
'程式名稱
Private strPrgName As String
'公司別
Private strCompCode As String
'服務類別
Private strServiceType As String

'以下是WriteLine 在用的

'GICMIS1.INI路徑
Private strPath As String
'ErrLog 路徑
Private strErrPath As String
'檔案名稱
Private strDataName As String
Private strMemoName As String
Private strErrName As String
Private strTrType As String
'轉帳組別
Private strStopDate As String  '應收日期截止日
'媒體多帳戶處理 (0=否 , 1=是)
Private strMedia As String
Private strChoose33 As String
Private strGetOwner As String   'OwnerName
Private strFlowId As String     '流程控制


Private Sub InitData()
  On Error GoTo ChkErr
    With objStorePara
        strActNo = .ActNo
        Set gcnGi = .Connection
        Set cnn = .MDBConn
        strBankId = .BankId
        strBankName = .BankName
        strCorpID = .CorpID
        strChoose = .ChooseStr
        strPath = .Inipath
        strErrPath = .ErrPath
        strCompCode = .uCompCode
        strServiceType = .uServiceType
        strStopDate = .uStopDate
        strChoose33 = .uChoose33
        strGetOwner = .uGetOwner
        strFlowId = .FlowId
    End With
            
    '定義記憶體recordset
    Call DefineRs

    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "InitData")
End Sub

Private Function BeginTran() As Boolean
  Dim strSQL As String
  Dim strOldBillNo As String
  Dim lngTime As Long
  Dim lngErrCount As Long
  Dim lngCount As Long
  Dim lngShouldAmt  As Long
  Dim lngTotAmt As Long
  Dim strBillNo As String, strBillNo_Old As String
  Dim rsTmp As New ADODB.Recordset
  Dim rsPara As New ADODB.Recordset
  Dim rsUpd As New ADODB.Recordset
  Dim rsCD013 As New ADODB.Recordset
  Dim StrTableName As String
  Dim strWhere As String, strSQLA As String, strID As String
  Dim strSQL106 As String, strSQL002 As String
  Dim strUcCode As String
  Dim strUcName As String
  Dim blnUpdUcCode As Boolean
  Dim strUPDUCCode As String
  Dim strUpdTime As String
  On Error Resume Next
     cnn.BeginTrans
     cnn.Execute ("Delete From SO3271A")
  On Error GoTo ChkErr

    BeginTran = False
    lngTime = Timer
    strUpdTime = Format(RightNow, "EE/MM/DD HH:MM:SS")
    '********************************************************************************************************************************************************
    '#3417 電子檔匯出時,要填入未收原因(RefNo=4) By Kin 2007/12/04
    If Not GetRS(rsCD013, "Select * From " & strGetOwner & "CD013 Where RefNo=4 And StopFlag<>1 Order By CodeNo Desc", gcnGi, adUseClient, adOpenKeyset) Then Exit Function
    If rsCD013.EOF Then
        blnUpdUcCode = False
    Else
        blnUpdUcCode = True
        strUcCode = rsCD013("CodeNo") & ""
        strUcName = rsCD013("Description") & ""
    End If
    '*********************************************************************************************************************************************************
    
        strSQL = "Select Para24 From " & strGetOwner & "SO043 Where Rownum=1"
        If Not GetRS(rsPara, strSQL, gcnGi) Then Exit Function
        strMedia = GetFieldValue(rsPara, "Para24") & ""

        ' 2003/3/26 Crystal Edit
        StrTableName = "," & strGetOwner & "SO002 C"
        strWhere = strWhere & " And A.CustId = C.CustId And A.Servicetype = C.ServiceType "
        StrTableName = StrTableName & "," & strGetOwner & "SO001 B"
        strWhere = strWhere & " And A.CustId = B.CustId "
        '************************************************************************
        '#5055 如果CD013.REFNO=3(櫃台已收)不再出帳 By Kin 2009/04/28
        '#5218 櫃台已收要用Nvl的方式 By Kin 2009/08/05
        '#5564 增加參考7,8代表已收,PayOK,也代表是已收 By Kin 2010/05/17
        StrTableName = StrTableName & "," & strGetOwner & "CD013"
        strWhere = strWhere & " And A.CancelFlag<> 1 And A.UCCode Is Not Null And " & _
                    " A.UCCode=CD013.CodeNo And Nvl(CD013.REFNO,0) NOT IN(3,7) " & _
                    " AND NVL(CD013.PAYOK,0)=0 "
        '************************************************************************
        If strWhere <> "" Then strWhere = Mid(strWhere, 5)
        '#5683 增加取出RealStopDate與PayKind By Kin 2010/08/04
        If strMedia = 0 Then
            strSQL = "SELECT " & GetUseIndexStr("SO033", "ShouldDate") & " A.BillNo," & _
                    " Sum(A.ShouldAmt) ShouldAmt,A.AccountNo," & _
                    " A.CustId,A.MediaBillNo,C.Id,Max(Nvl(PayKind,0)) PayKind, " & _
                    strMinRealStopDate & _
                     " From " & strGetOwner & "SO033 A " & StrTableName & _
                     " Where " & strWhere & _
                     IIf(strWhere = Empty, "", " And ") & strChoose & _
                     " Group By A.BillNo,A.AccountNo,A.CustId,A.MediaBillNo,C.Id"
        
            If Not GetRS(rsTmp, strSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
            If Not BatchUpdateMediaBillNo(rsTmp, strChoose33, gcnGi) Then Exit Function
        Else
            strSQLA = "SELECT " & GetUseIndexStr("SO033", "ShouldDate") & "A.BillNo ,A.MediaBillNo,A.AccountNo,A.CustId " & _
                     "From " & strGetOwner & "SO033 A " & " Where " & strChoose33 & " And A.MediaBillNo Is Null"
    
            strSQL = "SELECT " & GetUseIndexStr("SO033", "ShouldDate") & _
                    " A.MediaBillNo BillNo,Sum(A.ShouldAmt) ShouldAmt,A.AccountNo, " & _
                    " Max(Nvl(PayKind,0)) PayKind," & strMinRealStopDate & _
                     " From " & strGetOwner & "SO033 A " & StrTableName & _
                     " Where " & strWhere & _
                      IIf(strWhere = Empty, "", " And ") & strChoose & _
                      " Group By A.AccountNo,A.MediaBillNo"
                      
                                            
            If Not GetRS(rsUpd, strSQLA, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
            '批次更新MediaBillNo欄位
            If Not BatchUpdateMediaBillNo(rsUpd, strChoose33, gcnGi) Then Exit Function
            If Not GetRS(rsTmp, strSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
        End If
        
        '**************************************************************************************************
        '#3417 更新UCCode與UCName的語法 By Kin 2007/12/05
        '#4388 增加更新異動人員與異動時間 By Kin 2009/04/29
'        strUPDUCCode = "Select A.UCCode,A.UCName,A.UPDEN,A.UPDTime From " & _
'                        strGetOwner & "SO033 A" & StrTableName & _
'                        " Where " & strWhere & IIf(strWhere = Empty, "", " And ") & strChoose
        strUPDUCCode = GetUpdateSQL(Val(strMedia), StrTableName, strWhere, strChoose)
        '**************************************************************************************************
        
        Set rsTmp.ActiveConnection = Nothing
        '批次更新MediaBillNo欄位
        
        If rsTmp.RecordCount > 0 Then
           rsTmp.MoveFirst
        End If
    
    If rsTmp.BOF Or rsTmp.EOF Then
       MsgBox "無資料！請重新操作！", vbInformation, "訊息！"
       blnNodata = True
       Exit Function
    Else
        blnNodata = False
        strOldBillNo = "-1"
        
        While Not rsTmp.EOF
            '#5683 如果RealStopDate<畫面條件,不要出帳 By Kin 2010/08/04
            If Not IsPayKindOK(rsTmp, gdaTransfer.GetValue & "", giAll, Val(strMedia)) Then
                Call WriteSO3271Err(GetPayKindCustId(rsTmp, Val(strMedia)), rsTmp("BillNo"), rsTmp("REALSTOPDATE"))
                lngErrCount = lngErrCount + 1
                GoTo lNextRcd
            End If
           '如有多筆同BillNo 的資料
           lngCount = lngCount + 1
           If strMedia = 0 Then
                lngShouldAmt = Val(GetRsValue("Select Sum(ShouldAmt) From " & strGetOwner & "So033 Where BillNo = '" & rsTmp("BillNo") & "' And CancelFlag = 0 And UCCode Is Not Null ", gcnGi) & "")
           Else
                lngShouldAmt = Val(GetRsValue("Select Sum(ShouldAmt) From " & strGetOwner & "So033 Where MediaBillNo = '" & rsTmp("BillNo") & "' And AccountNo='" & GetFieldValue(rsTmp, "AccountNo") & "' And CancelFlag = 0 And UCCode Is Not Null ", gcnGi) & "")
           End If
           
'           lngShouldAmt = Val(GetRsValue("Select Sum(ShouldAmt) From " & strGetOwner & "So033 Where BillNo = '" & rsTmp("BillNo") & "' And CancelFlag = 0 And UCCode Is Not Null ", gcnGi) & "")
           If rsTmp("BillNo") = strOldBillNo & "" Then
              'lngShouldAmt = lngShouldAmt + GetFieldValue(rsTmp, "ShouldAmt") + 0
              GoTo lNextRcd
           End If
           
           If strMedia = 1 Then
              strBillNo = GetFieldValue(rsTmp, "BillNo") & ""
              strBillNo_Old = GetRsValue("Select BillNo From " & strGetOwner & "SO033 Where MediaBillNO='" & GetFieldValue(rsTmp, "BillNo") & "'", gcnGi) & ""
           Else
              strBillNo = GetBillNo_New(GetFieldValue(rsTmp, "BillNo") & "")
              strBillNo_Old = GetBillNo_Old(rsTmp("BillNo") & "")
           End If
           
           '計算扣帳總金額
           lngTotAmt = lngTotAmt + lngShouldAmt
           With rsDefTmp
                .AddNew
                .Fields("TransDate") = Replace(gdaTransfer.GetOriginalValue, "/", "")
                '轉帳日期 1-6
                .Fields("BankId") = txtBankId.Text
                '主辦行 7-10
                .Fields("CorpId") = txtCorpId.Text
                '營利編號 11-20
                .Fields("AccountNo") = Mid(GetFieldValue(rsTmp, "AccountNo") & "", 1, 13)
                '扣帳帳號 21-33
                strSQL106 = "Select AccountNameID From " & strGetOwner & "So106 Where AccountID='" & rsTmp("AccountNo") & _
                         "' And BankCode='" & strBankId & "' And SnactionDate Is Not Null And StopDate Is Null"
                strSQL002 = "Select Id From " & strGetOwner & "So002 Where AccountNo='" & rsTmp("AccountNo") & "'"
'               不論是否啟動多媒體都要先取SO106的值----for Lydia
'                If strMedia = 0 Then
'                    strID = GetFieldValue(rsTmp, "ID") & ""
'                Else    '啟動多媒體
                    strID = GetRsValue(strSQL106, gcnGi) & ""
                    If strID = "" Then strID = GetRsValue(strSQL002, gcnGi) & ""
'                End If
                .Fields("Id") = strID
                '身份證編號 34-43
                '#4082 將009改成010 By Kin 2008/09/16
                .Fields("TransType") = "010"
                '轉帳別 44-46
                .Fields("Amt") = lngShouldAmt
                '扣帳金額 47-59
                .Fields("BillNo") = strBillNo
                '保留欄位(11碼舊單據編號) 60-80
           End With
                'Insert Into MDB(SO3271A)
                cnn.Execute Replace("Insert Into SO3271A (AccountNo, TransDate, ShouldAmt, BillNo_Old, BillNo_New) Values (" & _
                     GetNullString(rsTmp("AccountNo")) & "," & GetNullString(gdaTransfer.GetValue(True), giDateV, giAccessDb) & "," & _
                     GetNullString(lngShouldAmt) & "," & GetNullString(strBillNo_Old) & "," & _
                     GetNullString(strBillNo) & ")", Chr(0), "")
'                     GetNullString(lngShouldAmt) & "," & IIf(strMedia = "0", GetNullString(GetBillNo_Old(rsTmp("BillNo"))), GetNullString(rsTmp("BillNo"))) & "," & _


lNextRcd:
           strOldBillNo = rsTmp("BillNo")
           rsTmp.MoveNext
           DoEvents
        Wend
        cnn.CommitTrans
        '**************************************************************************************************************************
        '#3417 更新UCCode與UCName By Kin 2007/12/04
        '#4388 增加更新異動人員與異動時間 By Kin 2009/04/29
        If blnUpdUcCode Then
            Dim rsUpdUCCode As New ADODB.Recordset
            If Not GetRS(rsUpdUCCode, strUPDUCCode, gcnGi, adUseClient, adOpenKeyset, adLockPessimistic) Then Exit Function
            Do While Not rsUpdUCCode.EOF
                '#5683 如果RealStopDate<畫面條件則不要UPD By Kin 2010/08/04
                If IsPayKindOK(rsUpdUCCode, gdaTransfer.GetValue, giSingle, Val(strMedia)) Then
                    rsUpdUCCode("UPDEN") = objStorePara.uUpdEn
                    rsUpdUCCode("UPDTime") = strUpdTime
                    rsUpdUCCode("UCcode") = strUcCode
                    rsUpdUCCode("UCName") = strUcName
                    rsUpdUCCode.Update
                End If
                rsUpdUCCode.MoveNext
            Loop
            CloseRecordset rsUpdUCCode
        End If
        '**************************************************************************************************************************
        
        rsTmp.Close
        Set rsTmp = Nothing
        If strFlowId = 0 Then
            rsPara.Close
            Set rsPara = Nothing
        End If
        CloseRecordset rsCD013
        If Not subWriteLine() Then
            msgResult lngCount, lngErrCount, lngTime      '顯示執行結果
            CloseFS
            Exit Function      '客戶內容
        End If
        WriteLineData txtMemo1 & vbCrLf & txtMemo2, 1 '備註檔
        msgResult lngCount, lngErrCount, lngTime      '顯示執行結果
        BeginTran = True
        CloseFS
     End If
Exit Function
ChkErr:
    CloseFS
    Call ErrSub(Me.Name, "BeginTran")
End Function

Private Function subWriteLine() As Boolean
  On Error GoTo ChkErr
    Dim varData As Variant
    Dim strData As String
    Dim lngloop As Long
    
    subWriteLine = False
    With rsDefTmp
         If .BOF Or .EOF Then Exit Function
         .MoveFirst
         varData = .GetRows
         '內容
         For lngloop = 0 To .RecordCount - 1
            WriteLineData GetString(varData(0, lngloop), 6) & GetString(varData(1, lngloop), 4) & _
                GetString(varData(2, lngloop), 10) & GetString(varData(3, lngloop), 13) & _
                GetString(varData(4, lngloop), 10) & GetString(varData(5, lngloop), 3) & _
                GetString(varData(6, lngloop), 11, GIRIGHT, True) & "00" & GetString(varData(7, lngloop), 11), 0
             DoEvents
         Next
    End With
    subWriteLine = True
    On Error Resume Next
    rsDefTmp.Close
    Set rsDefTmp = Nothing
    Exit Function
ChkErr:
    subWriteLine = False
    Call ErrSub(Me.Name, "subWriteLine")
End Function

'定義記憶體Recordset
Private Sub DefineRs()
    On Error GoTo ChkErr
    With rsDefTmp
        If .State = adStateOpen Then .Close
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        
        .Fields.Append ("TransDate"), adBSTR, 6, adFldIsNullable    '轉帳日期
        .Fields.Append ("BankId"), adBSTR, 4, adFldIsNullable       '主辦行
        .Fields.Append ("CorpId"), adBSTR, 10, adFldIsNullable      '營利編號
        .Fields.Append ("AccountNo"), adBSTR, 13, adFldIsNullable   '開戶人帳號
        .Fields.Append ("Id"), adBSTR, 10, adFldIsNullable          '身份證編號
        .Fields.Append ("TransType"), adBSTR, 3, adFldIsNullable    '轉帳別
        .Fields.Append ("Amt"), adBSTR, 13, adFldIsNullable         '轉帳金額(後2位是小數)
        .Fields.Append ("BillNo"), adBSTR, 21, adFldIsNullable      '保留欄位(11碼新單據編號+3個空白)
        .Open
    End With
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "DefineRs")
End Sub

Private Sub cmdCancel_Click()
    On Error Resume Next
    Unload Me
End Sub

Private Sub cmdDataPath_Click()
    On Error GoTo ChkErr
        With cnd1
                .FileName = txtDataPath.Text
                .Filter = "文字檔|*.txt"
                .InitDir = strErrPath
                .Action = 1
                txtDataPath = .FileName
        End With
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "cmdDataPath_Click")
End Sub

Private Sub cmdErrPath_Click()
    On Error GoTo ChkErr
        With cnd1
                .FileName = txtErrPath.Text
                .Filter = "文字檔|*.txt"
                .InitDir = strErrPath
                .Action = 1
                txtErrPath = .FileName
        End With
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "cmdErrPath_Click")
End Sub

Private Sub cmdMemoPath_Click()
    On Error GoTo ChkErr
        With cnd1
                .FileName = txtMemoPath.Text
                .Filter = "文字檔|*.txt"
                .InitDir = strErrPath
                .Action = 1
                txtMemoPath = .FileName
        End With
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "cmdMemoPath_Click")
End Sub

Private Sub cmdOK_Click()
    On Error GoTo ChkErr
    
    If Not ChkDTok Then Exit Sub
    If Not IsDataOk Then Exit Sub
    If OpenFile(txtDataPath, 0, True) = False Then txtDataPath.SetFocus: Exit Sub
    If OpenFile(txtMemoPath, 1, True) = False Then txtMemoPath.SetFocus: Exit Sub
    If OpenFile(txtErrPath, 2, True) = False Then txtErrPath.SetFocus: Exit Sub
    Screen.MousePointer = vbHourglass
    objStorePara.uProcText = txtDataPath.Text
    objStorePara.uProcErrText = txtErrPath.Text
    Call ScrToRcd         '將畫面輸入條件寫到Log檔
    
    If Not BeginTran Then
       If Not blnNodata Then
            objStorePara.uUpdate = False
            MsgBox "轉帳資料輸出失敗!", vbExclamation, "警告..."
            CloseFS
            Screen.MousePointer = vbDefault
            Unload Me
            Exit Sub
       Else
            objStorePara.uNoData = True
            CloseFS
            Screen.MousePointer = vbDefault
            Unload Me
            Exit Sub
       End If
    End If

    objStorePara.uUpdate = True
    Screen.MousePointer = vbDefault
    Unload Me
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "cmdOK_Click")
End Sub

Private Function IsDataOk() As Boolean
On Error GoTo ChkErr
Dim strErrMsg As String
        
        If gdaTransfer.GetValue = "" Then strErrMsg = "自動轉帳扣款日期": gdaTransfer.SetFocus:  GoTo Warning
        If txtDataPath.Text = "" Then strErrMsg = "資料檔位置": txtDataPath.SetFocus: GoTo Warning
        If txtMemoPath.Text = "" Then strErrMsg = "備註檔位置": txtMemoPath.SetFocus: GoTo Warning
        If txtErrPath.Text = "" Then strErrMsg = "問題參考檔位置": txtErrPath.SetFocus: GoTo Warning
        
        strDataName = Mid(txtDataPath.Text, InStrRev(txtDataPath.Text, "\") + 1)
        strMemoName = Mid(txtMemoPath.Text, InStrRev(txtMemoPath.Text, "\") + 1)
        strErrName = Mid(txtErrPath.Text, InStrRev(txtErrPath.Text, "\") + 1)
        If (strDataName = strMemoName) Or (strMemoName = strErrName) Or (strErrName = strDataName) Then
           MsgBox "檔案名稱不可重複!", vbExclamation, "警告"
           Exit Function
        End If
        
    IsDataOk = True
    Exit Function
Warning:
    MsgBox strErrMsg & "  " & gMsgIsDataOK, vbExclamation, "訊息！"
    Exit Function
ChkErr:
    Call ErrSub(Me.Name, "IsDataOk")
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   On Error GoTo ChkErr
    Select Case KeyCode
             Case vbKeyEscape
                     Unload Me
             Case vbKeyF2
                     If cmdOK.Enabled = True Then cmdOK.Value = True
    End Select
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "Form_KeyDown")
End Sub

Private Sub Form_Load()
    On Error Resume Next
      Me.Caption = objStorePara.BankName & ""
      Call InitData
      RcdToScr
End Sub

Private Sub RcdToScr()
    On Error GoTo ChkErr
    Dim LogFile As TextStream
    Dim rsBank As New ADODB.Recordset
    Dim strSQL As String
    
        strSQL = "Select BankId,CorpId From " & strGetOwner & "CD018 Where CompCode =" & strCompCode & " And CodeNo = '" & strBankId & "'"
        If Not GetRS(rsBank, strSQL, gcnGi) Then Exit Sub
        txtBankId.Text = GetFieldValue(rsBank, "BankId") & ""
        txtCorpId.Text = GetFieldValue(rsBank, "CorpId") & ""
        
        rsBank.Close
        Set rsBank = Nothing

        If Dir(strErrPath & "\" & strPrgName & "Out.log") = "" Then Exit Sub
            Set LogFile = fso.OpenTextFile(strErrPath & "\" & strPrgName & "Out.log", ForReading)
            With LogFile
                    If Not .AtEndOfStream Then gdaAgency.Text = .ReadLine & ""
                        '臨櫃代收繳費期限
                    If Not .AtEndOfStream Then gdaTransfer.Text = .ReadLine & ""
                        '自動轉帳繳費期限
                    If Not .AtEndOfStream Then txtBankId.Text = .ReadLine & ""
                        '主辦行
                    If Not .AtEndOfStream Then txtCorpId.Text = .ReadLine & ""
                        '營利編號
                    If Not .AtEndOfStream Then txtDataPath = .ReadLine & ""
                        '資料檔
                    If Not .AtEndOfStream Then txtErrPath = .ReadLine & ""
                        '問題檔
                    If Not .AtEndOfStream Then txtMemoPath = .ReadLine & ""
                        '備註檔
                    If Not .AtEndOfStream Then txtMemo1 = .ReadLine & ""
                        '備註1
                    If Not .AtEndOfStream Then txtMemo2 = .ReadLine & ""
                        '備註2
            End With
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "RcdToScr")
End Sub

Private Sub ScrToRcd()
    On Error GoTo ChkErr
    Dim LogFile As TextStream
    
        Set LogFile = fso.CreateTextFile(strErrPath & "\" & strPrgName & "Out.log", True)
         With LogFile
           Call .WriteLine(gdaAgency.Text)           '臨櫃代收繳費期限
                .WriteLine (gdaTransfer.Text)        '自動轉帳繳費期限
                .WriteLine (txtBankId.Text)          '主辦行
                .WriteLine (txtCorpId.Text)          '營利編號
                .WriteLine (txtDataPath)             '資料檔
                .WriteLine (txtErrPath)              '問題檔
                .WriteLine (txtMemoPath)             '備註檔
                .WriteLine (txtMemo1)                '備註1
                .WriteLine (txtMemo2)                '備註2
        End With
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "ScrToRcd")
End Sub

Public Property Get PrgName() As String
    PrgName = strPrgName
End Property

Public Property Let PrgName(ByVal vData As String)
    strPrgName = vData
End Property
