VERSION 5.00
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPost3 
   ClientHeight    =   6810
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11190
   Icon            =   "frmPost3.frx":0000
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
      TabIndex        =   12
      Top             =   180
      Width           =   10695
      Begin VB.Frame frmData 
         Caption         =   "資料檔位置"
         Height          =   1875
         Left            =   450
         TabIndex        =   13
         Top             =   1680
         Width           =   9075
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
            TabIndex        =   2
            ToolTipText     =   "請輸入字檔之路徑及檔名！"
            Top             =   360
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
            Height          =   375
            Left            =   7440
            TabIndex        =   3
            Top             =   360
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
            Height          =   375
            Left            =   7440
            TabIndex        =   7
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
            Left            =   7440
            TabIndex        =   5
            Top             =   810
            Width           =   375
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
            TabIndex        =   4
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
            Top             =   1200
            Width           =   4095
         End
         Begin VB.Label lblDatapath 
            AutoSize        =   -1  'True
            Caption         =   "資料檔位置（路徑）"
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
            Left            =   360
            TabIndex        =   19
            Top             =   480
            Width           =   1755
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
            Left            =   360
            TabIndex        =   15
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
            Left            =   360
            TabIndex        =   14
            Top             =   1320
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
         TabIndex        =   8
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
         TabIndex        =   9
         Top             =   3720
         Width           =   3015
      End
      Begin Gi_Date.GiDate gdaAgency 
         Height          =   375
         Left            =   2160
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
      Begin Gi_Date.GiDate gdaTransfer 
         Height          =   375
         Left            =   2160
         TabIndex        =   1
         Top             =   1020
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
      Begin VB.Label lblTransfer 
         AutoSize        =   -1  'True
         Caption         =   "自動轉帳扣款日期"
         Height          =   180
         Left            =   540
         TabIndex        =   20
         Top             =   1110
         Width           =   1440
      End
      Begin VB.Label lblagency 
         AutoSize        =   -1  'True
         Caption         =   "繳費期限"
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
         TabIndex        =   18
         Top             =   570
         Width           =   780
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
         TabIndex        =   17
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
         TabIndex        =   16
         Top             =   3750
         Width           =   1365
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "F2.開始"
      Height          =   405
      Left            =   3570
      TabIndex        =   10
      Top             =   6180
      Width           =   1275
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&X)"
      Height          =   405
      Left            =   5940
      TabIndex        =   11
      Top             =   6180
      Width           =   1275
   End
End
Attribute VB_Name = "frmPost3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As New ADODB.Connection
Private rsDefTmp As New ADODB.Recordset
Private rsBank As New ADODB.Recordset

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
Private strGetOwner As String       'OwnerName
Private strFlowId As String         '流程控制
Private strBillHeadFmt As String    '多帳戶產生依據設定

Dim strPostUnit As String           'PostUnit 郵局發件單位
Dim intRefNo As Integer
Private strHavingOutZero As String
Private blnHavingOutZero As Boolean
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
        strBillHeadFmt = .uBillHeadFmt '多帳戶產生依據設定 PostUnit 郵局發件單位
        intRefNo = .uRefNo
        '#6441 單張帳單合計總額<=0是否產生 By Kin 2013/05/13
        strHavingOutZero = ""
        blnHavingOutZero = True
        If Not .uOutZero Then
            strHavingOutZero = " Having Sum(A.ShouldAmt)>0 "
            strHavingOutZero = ""
            blnHavingOutZero = False
        End If
        '定義記憶體recordset
        Call DefineRs
    End With
    
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "InitData")
End Sub

Private Function BeginTran() As Boolean
    Dim strSQL As String
    Dim strAMduSQL As String
    Dim strOldBillNo As String
    Dim strOldAccountId As String
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
    Dim strWhere As String
    Dim strSQLA As String
    Dim strAccNameID As String
    Dim strCitibankATM As String
    Dim strUcCode As String
    Dim strUcName As String
    Dim blnUpdUcCode As Boolean
    Dim strUPDUCCode As String
    Dim strUpdTime As String
    Dim intCrossCustCombine As Integer
    Dim rsUpdUCCode As New ADODB.Recordset
    Dim rsCustId As New ADODB.Recordset
    Dim aBillNoType As String
  On Error Resume Next
     cnn.BeginTrans
     cnn.Execute ("Delete From SO3271A")
  On Error GoTo ChkErr

    BeginTran = False
    If Len(sqlTmpViewName) > 0 Then
        gcnGi.Execute "Drop View " & strGetOwner & sqlTmpViewName
        sqlTmpViewName = Empty
    End If
    
    lngTime = Timer
'    strUpdTime = Format(RightNow, "EE/MM/DD HH:MM:SS")
    strUpdTime = GetDTString(RightNow)
    '#5843 增加檢核參數 By Kin 2010/11/24
    blnChkPayKind = GetChkPayKind
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
    '#7179
    isCrossCustCombine = Val(GetRsValue("Select Nvl(CrossCustCombine,0) From " & strGetOwner & "SO041", gcnGi) & "")
    If Not GetRS(rsPara, strSQL, gcnGi) Then Exit Function
    strMedia = GetFieldValue(rsPara, "Para24") & ""
    
    ' 2003/3/26 Crystal Edit
    If InStr(1, strChoose, "C.") Then
        StrTableName = "," & strGetOwner & "SO002 C"
        strWhere = strWhere & " And A.CustId = C.CustId And A.Servicetype = C.ServiceType "
    End If
    If InStr(1, strChoose, "B.") Then
        StrTableName = StrTableName & "," & strGetOwner & "SO001 B"
        strWhere = strWhere & " And A.CustId = B.CustId"
    End If
    
    '************************************************************************
    '#5055 如果CD013.REFNO=3(櫃台已收)不再出帳 By Kin 2009/04/29
    '#5218 櫃台已收要用Nvl的方式 By Kin 2009/08/05
    '#5564 增加參考7,8代表已收,PayOK,也代表是已收 By Kin 2010/05/17
    StrTableName = StrTableName & "," & strGetOwner & "CD013"
    strWhere = strWhere & " And A.CancelFlag<> 1 And UCCode Is Not Null " & _
                " And A.UCCode=CD013.CodeNo And Nvl(CD013.REFNO,0) NOT IN(3,7) " & _
                " AND NVL(CD013.PAYOK,0) = 0 "
    '************************************************************************
    If strWhere <> "" Then strWhere = Mid(strWhere, 5)
    '************************************************************************************************************************************************************************************************************
    '#3527 如果參考號為2，則使用新的POS規格 By Kin 2007/10/04
    '#5683 增加挑選RealStopDate 與 PayKind By Kin 2010/08/05
    If intRefNo <> 2 Then
        If strMedia = 0 Then
             '#6441 單張帳單合計總額<=0是否產生 By Kin 2013/05/13
             aBillNoType = "BillNo"
            strSQL = "SELECT " & GetUseIndexStr("SO033", "ShouldDate") & _
                    " A.BillNo,Sum(A.ShouldAmt) ShouldAmt,A.AccountNo," & _
                    " A.CustId,A.MediaBillNo ,TO_CHAR(A.ShouldDate,'YYYYMM') as ShouldDate, " & _
                    " Max(Nvl(A.PayKind,0)) PayKind," & strMinRealStopDate & _
                    " From " & strGetOwner & "SO033 A " & StrTableName & _
                    " Where " & strWhere & _
                     IIf(Len(strWhere) = 0, "", " And ") & strChoose & strHavingOutZero & _
                     " Group By A.BillNo,A.AccountNo,A.CustId,A.MediaBillNo,A.ShouldDate"
            strSQL = "Select * From ( " & strSQL & " ) "
            If Not blnHavingOutZero Then
                strSQL = strSQL & " Where ShouldAmt > 0 "
            End If
            If Not GetRS(rsTmp, strSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
            If Not BatchUpdateMediaBillNo(rsTmp, strChoose33, gcnGi) Then Exit Function
        Else
            aBillNoType = "MediabillNo"
            strSQLA = "SELECT " & GetUseIndexStr("SO033", "ShouldDate") & "A.BillNo ," & _
                        "A.MediaBillNo,A.AccountNo,A.CustId " & _
                     "From " & strGetOwner & "SO033 A " & " Where " & strChoose33 & " And A.MediaBillNo Is Null"
            '#6320 增加客編 By Kin 2012/09/18
             '#6441 單張帳單合計總額<=0是否產生 By Kin 2013/05/13
            strSQL = "SELECT " & GetUseIndexStr("SO033", "ShouldDate") & _
                    " A.MediaBillNo BillNo,Sum(A.ShouldAmt) ShouldAmt," & _
                    " A.AccountNo,TO_CHAR(A.ShouldDate,'YYYYMM') as ShouldDate,A.CustId, " & _
                    " Max(Nvl(A.PayKind,0)) PayKind," & strMinRealStopDate & _
                    " From " & strGetOwner & "SO033 A " & StrTableName & _
                    " Where " & strWhere & _
                     IIf(Len(strWhere) = 0, "", " And ") & strChoose & strHavingOutZero & _
                     " Group By A.AccountNo,A.MediaBillNo,A.ShouldDate,A.CustId"
            strSQL = "Select * From (" & strSQL & ") "
            If Not GetRS(rsUpd, strSQLA, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
            '批次更新MediaBillNo欄位
            If Not BatchUpdateMediaBillNo(rsUpd, strChoose33, gcnGi) Then Exit Function
            If isCrossCustCombine Then
                '#7179
                intCrossCustCombine = 1
                sqlTmpViewName = GetTmpViewName(gcnGi)
                '將套房與套房主客編找出來,如果無套房則主客編用custid代入
                strSQL = "Select A.*," & _
                            " (Case " & intCrossCustCombine & " When 1 then Nvl(SO001.AMduId,Null)   else Null End ) AMduId, " & _
                            " (Case " & intCrossCustCombine & " When 1 Then " & _
                                           " ( Case  When AMDUID Is Null Then A.CustId Else Nvl((Select MainCustId From " & strGetOwner & "SO202 Where SO001.AMduId = SO202.MduId  ),-1)  End ) " & _
                            "  Else A.CustId  End ) MainCustId " & _
                            " From (" & strSQL & ") A," & strGetOwner & "SO001 " & _
                    " Where A.CustId=SO001.CustId"
                strSQL = "Create View " & strGetOwner & sqlTmpViewName & " As (" & strSQL & ")"
                If Not ExecuteCommand(strSQL, gcnGi, lngCount) Then Exit Function
                '判斷套房主客編是否有出現在SO033.CUSTID
'                strSQL = " select A.*,(select (case " & _
'                                      "                when amduid is null then 1 Else " & _
'                                                              " ( select  count(1) from " & strGetOwner & sqlTmpViewName & _
'                                                              " Where A.MainCustId = Custid And A.AmdUid = AmduId And Nvl(A.BillNo,'X')= Nvl(BillNo,'X') " & _
'                                                               " And A.AccountNo = AccountNo ) " & _
'                                                      " End)  from dual) CustIdExistFlag  from " & strGetOwner & sqlTmpViewName & " A"
                 strSQL = " select A.*,(select (case " & _
                                      "                when amduid is null then 1 Else " & _
                                                              " ( select  count(1) from " & strGetOwner & sqlTmpViewName & _
                                                              " Where A.MainCustId = Custid  ) " & _
                                                      " End)  from dual) CustIdExistFlag  from " & strGetOwner & sqlTmpViewName & " A"
                strSQL = "Select Decode(CustIdExistFlag,0,Custid,MainCustId)  CustId,Nvl(BillNo,'X') BillNo,AccountNo,Max(PayKind) PayKind,Min(RealStopDate) RealStopDate," & _
                                " Sum(ShouldAmt) ShouldAmt,AMDUID,CustIdExistFlag " & _
                        " From ( " & strSQL & " ) A Group By Decode(CustIdExistFlag,0,Custid,MainCustId) ,BillNo,AccountNo,AMDUID,CustIdExistFlag "
                '將應收日找出來
                strSQL = "Select distinct A.*, " & _
                            " (select ShouldDate from " & strGetOwner & sqlTmpViewName & _
                                    " where a.billno=billno and custid=maincustid and  A.AccountNo = AccountNo And rownum<=1) ShouldDate " & _
                            " From (" & strSQL & ") A"
                           
            End If
            If Not blnHavingOutZero Then
                strSQL = strSQL & " Where ShouldAmt > 0 "
            End If
            If Not GetRS(rsTmp, strSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
        End If
    Else
        If strMedia = 1 Then
            strSQLA = "SELECT " & GetUseIndexStr("SO033", "ShouldDate") & _
                    " A.BillNo ,A.MediaBillNo,A.AccountNo,A.CustId " & _
                    " From " & strGetOwner & "SO033 A " & " Where " & strChoose33 & " And A.MediaBillNo Is Null"
            If Not GetRS(rsUpd, strSQLA, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
            '批次更新MediaBillNo欄位
            If Not BatchUpdateMediaBillNo(rsUpd, strChoose33, gcnGi) Then Exit Function
        End If
        aBillNoType = "CitibankATM"
        '#6320 增加客編 By Kin 2012/09/18
         '#6441 單張帳單合計總額<=0是否產生 By Kin 2013/05/13
        strSQL = "SELECT " & GetUseIndexStr("SO033", "ShouldDate") & _
                " A.CitibankATM BillNo,Sum(A.ShouldAmt) ShouldAmt,A.AccountNo,A.CustId," & _
                " TO_CHAR(A.ShouldDate,'YYYYMM') as ShouldDate,A.BillNo OldBillNo, " & _
                " Max(Nvl(A.PayKind,0)) PayKind," & strMinRealStopDate & _
               "From " & strGetOwner & "SO033 A " & StrTableName & _
               " Where " & strWhere & _
               IIf(Len(strWhere) = 0, "", " And ") & strChoose & strHavingOutZero & _
               " Group By A.AccountNo,A.CitibankATM,A.ShouldDate,BillNo,A.CustId Order by BillNo"
         strSQL = "Select * From ( " & strSQL & " ) "
        If isCrossCustCombine Then
            '#7179
            intCrossCustCombine = 1
            sqlTmpViewName = GetTmpViewName(gcnGi)
            '將套房與套房主客編找出來,如果無套房則主客編用custid代入
            strSQL = "Select A.*," & _
                        " (Case " & intCrossCustCombine & " When 1 then Nvl(SO001.AMduId,Null)   else Null End ) AMduId, " & _
                        " (Case " & intCrossCustCombine & " When 1 Then " & _
                                       " ( Case  When AMDUID Is Null Then A.CustId Else Nvl((Select MainCustId From " & strGetOwner & "SO202 Where SO001.AMduId = SO202.MduId  ),-1)  End ) " & _
                        "  Else A.CustId  End ) MainCustId " & _
                        " From (" & strSQL & ") A," & strGetOwner & "SO001 " & _
                " Where A.CustId=SO001.CustId"
            strSQL = "Create View " & strOwnerName & sqlTmpViewName & " As (" & strSQL & ")"
            If Not ExecuteCommand(strSQL, gcnGi, lngCount) Then Exit Function
            '判斷套房主客編是否有出現在SO033.CUSTID
'            strSQL = " select A.*,(select (case " & _
'                                  "                when amduid is null then 1 Else " & _
'                                                          " ( select  count(1) from " & strGetOwner & sqlTmpViewName & _
'                                                          " Where A.MainCustId = Custid And A.AmdUid = AmduId And Nvl(A.BillNo,'X')= Nvl(BillNo,'X') " & _
'                                                           " And A.AccountNo = AccountNo ) " & _
'                                                  " End)  from dual) CustIdExistFlag  from " & strGetOwner & sqlTmpViewName & " A"
'             select A.*,(select (case                 when amduid is null then 1 Else  ( select  count(1)
' From CNSCL.TMP_0016026
' Where A.MainCustId = CustId
' AND A.AmdUid = AmduId
' AND Nvl(A.BillNo,'X')= Nvl(BillNo,'X')
' AND A.AccountNo = AccountNo )  End)
' FROM dual) CustIdExistFlag ,
' (Select max(Nvl(CustAllot,0)) from so003,so033 where so003.custid=a.custid  and so033.custid=so003.custid and a.billno= so033.citibankatm and so033.faciseqno = so003.faciseqno)
' FROM CNSCL.TMP_0016026 A
            strSQL = " select A.*,(select (case " & _
                                  "                when amduid is null then 1 Else " & _
                                                          " ( select  count(1) from " & strGetOwner & sqlTmpViewName & _
                                                          " Where A.MainCustId = Custid  ) " & _
                                                  " End)  from dual) CustIdExistFlag  from " & strGetOwner & sqlTmpViewName & " A"
            strSQL = "Select Decode(CustIdExistFlag,0,Custid,MainCustId)  CustId,Nvl(BillNo,'X') BillNo,AccountNo,Max(PayKind) PayKind,Min(RealStopDate) RealStopDate," & _
                            " Sum(ShouldAmt) ShouldAmt,AMDUID,CustIdExistFlag " & _
                    " From ( " & strSQL & " ) A Group By Decode(CustIdExistFlag,0,Custid,MainCustId) ,BillNo,AccountNo,AMDUID,CustIdExistFlag "
            '將舊單號與應收日找出來
            strSQL = "Select distinct A.*, " & _
                        " (select OldBillNo from " & strGetOwner & sqlTmpViewName & _
                                " where a.billno=billno and a.custid=maincustid and A.AccountNo = AccountNo And rownum<=1) OldBillNo, " & _
                        " (select ShouldDate from " & strGetOwner & sqlTmpViewName & _
                                " where a.billno=billno and  A.AccountNo = AccountNo And rownum<=1) ShouldDate " & _
                        " From (" & strSQL & ") A"
                       
        End If
        '#3583 批次更新虛擬帳號(CitiBankATM) By Kin 2007/10/26
        If Not BatchUpdateCitiBankATM(strWhere & IIf(Len(strWhere) = 0, "", " And ") & strChoose, StrTableName, gcnGi) Then Exit Function
        If Not blnHavingOutZero Then
            strSQL = strSQL & " Where ShouldAmt > 0 "
        End If
        If Not GetRS(rsTmp, strSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function

    End If
    '************************************************************************************************************************************************************************************************************
    
    '**************************************************************************************************
    '#3417 更新UCCode與UCName的語法 By Kin 2007/12/05
    '#4388 增加更新異動人員與異動時間 By Kin 2009/04/29
'    strUPDUCCode = "Select A.UCCode,A.UCName,A.UPDTime,A.UPDEN From " & _
'                strGetOwner & "SO033 A" & StrTableName & _
'              " Where " & strWhere & _
'              IIf(Len(strWhere) = 0, "", " And ") & strChoose
    '#6441 單張帳單合計總額<=0是否產生 By Kin 2013/05/13
    '#7179 傳入SQL語法，不然會Update錯誤
    strUPDUCCode = GetUpdateSQL(IIf(intRefNo = 2, 3, Val(strMedia)), StrTableName, strWhere, _
                     strChoose, IIf(Len(strHavingOutZero) = 0, True, False), IIf(isCrossCustCombine, strSQL, ""))
    '**************************************************************************************************
    Set rsTmp.ActiveConnection = Nothing
    
'    If Not GetRS(rsTmp, strSQL, gcnGi) Then Exit Function

    strSQL = "Select * From " & strGetOwner & "CD018 Where CodeNo ='" & strBankId & "'"
    If Not GetRS(rsBank, strSQL, gcnGi) Then Exit Function
    
    strPostUnit = GetRsValue("Select PostUnit From " & strGetOwner & "CD068 Where BillHeadFmt='" & strBillHeadFmt & "'", gcnGi) & ""
    
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
        strOldAccountId = "-1"
        While Not rsTmp.EOF
        
            '#5683 如果RealStopDate<畫面條件,不要出帳 By Kin 2010/08/05
            '#5843 如果設定不檢核則強迫出帳 By Kin 2010/11/24
            If (Not IsPayKindOK(rsTmp, gdaTransfer.GetValue & "", giAll, IIf(intRefNo = 2, 3, Val(strMedia)))) And (blnChkPayKind) Then
                Call WriteSO3271Err(GetPayKindCustId(rsTmp, IIf(intRefNo = 2, 3, Val(strMedia))), rsTmp("BillNo"), rsTmp("REALSTOPDATE"))
                lngErrCount = lngErrCount + 1
                GoTo lNextRcd
            End If
           '如有多筆同BillNo 的資料
           'lngCount = lngCount + 1
           '************************************************************************************************************************************************************************************************************************************************************************************************
           '#3527　參考號為2，使用新的POS流程 By Kin 2007/10/04
            If intRefNo <> 2 Then
                If strMedia = 0 Then
                     lngShouldAmt = Val(GetRsValue("Select Sum(A.ShouldAmt) From " & strGetOwner & "So033 A Where A.BillNo = '" & rsTmp("BillNo") & "' And A.CancelFlag = 0 And A.UCCode Is Not Null AND " & strChoose33, gcnGi) & "")
                Else
                     lngShouldAmt = Val(GetRsValue("Select Sum(A.ShouldAmt) From " & strGetOwner & "So033 A Where A.MediaBillNo = '" & rsTmp("BillNo") & "' And A.AccountNo='" & GetFieldValue(rsTmp, "AccountNo") & "' And A.CancelFlag = 0 And A.UCCode Is Not Null AND " & strChoose33, gcnGi) & "")
                End If
                '7179
                If isCrossCustCombine Then
                    lngShouldAmt = rsTmp("ShouldAmt") & ""
                End If
            Else
                lngShouldAmt = Val(GetRsValue("Select Sum(A.ShouldAmt) From " & strGetOwner & "So033 A Where A.CitibankATM = '" & rsTmp("BillNo") & "' And A.AccountNo='" & GetFieldValue(rsTmp, "AccountNo") & "' And A.CancelFlag = 0 And A.UCCode Is Not Null AND " & strChoose33, gcnGi) & "")
                '7179
                If isCrossCustCombine Then
                    lngShouldAmt = rsTmp("ShouldAmt") & ""
                End If
            End If
            '***********************************************************************************************************************************************************************************************************************************************************************************************
            
            If (rsTmp("BillNo") = strOldBillNo & "") And (rsTmp("AccountNo") = strOldAccountId) Then
               'lngShouldAmt = lngShouldAmt + GetFieldValue(rsTmp, "ShouldAmt") + 0
               GoTo lNextRcd
            End If
            
           '********************************************************************************************************************************************************
           '#3527　參考號為2，使用新的POS流程 By Kin 2007/10/04
            If intRefNo <> 2 Then
                If strMedia = 1 Then
                   strBillNo = GetFieldValue(rsTmp, "BillNo") & ""
                    strBillNo_Old = GetRsValue("Select BillNo From SO033 Where " & _
                                    " 1=1 " & IIf(isCrossCustCombine, " And Custid=" & rsTmp("CustId"), "") & _
                                    "  And MediaBillNO='" & GetFieldValue(rsTmp, "BillNo") & "'", gcnGi) & ""
                Else
                   strBillNo = GetBillNo_New(GetFieldValue(rsTmp, "BillNo") & "")
                   strBillNo_Old = GetBillNo_Old(rsTmp("BillNo") & "")
                End If
            Else
                strBillNo = GetFieldValue(rsTmp, "BillNo") & ""
                If strBillNo <> "" Then
                    strBillNo_Old = GetRsValue("Select BillNo From " & strGetOwner & "SO033 Where CitibankATM='" & GetFieldValue(rsTmp, "BillNo") & "'", gcnGi) & ""
'                    strBillNo_Old = GetRsValue("Select BillNo From SO033 Where " & _
'                                    " 1=1 " & IIf(isCrossCustCombine, " And Custid=" & rsTmp("CustId"), "") & _
'                                    "  And CitibankATM ='" & GetFieldValue(rsTmp, "BillNo") & "'", gcnGi) & ""
                Else
                    WriteLineData "單據編號： " & rsTmp("OldBillNo") & "" & " 虛擬帳號資料不正確！！", 2
                    lngErrCount = lngErrCount + 1
                    GoTo lNextRcd
                End If
            End If
            '********************************************************************************************************************************************************
            '#7179
            If isCrossCustCombine Then
                If Val(rsTmp("CustId") & "") = -1 Then
                        WriteLineData "單據編號： " & rsTmp("BillNo") & "套房找不到統收戶客戶編號，請設定統收戶客戶編號", 2
                        lngErrCount = lngErrCount + 1
                        GoTo lNextRcd
                End If
'                If Val(rsTmp("CustIdExistFlag") & "") = 0 Then
'                    WriteLineData "單據編號：[ " & rsTmp("BillNo") & " ]　出帳資料沒有包含統收戶客戶編號：" & rsTmp("CustId"), 2
'                    lngErrCount = lngErrCount + 1
'                    GoTo lNextRcd
'                End If
            End If
            lngCount = lngCount + 1
            '計算扣帳總金額
             lngTotAmt = lngTotAmt + lngShouldAmt
              '問題集2386   Lydia提 Edit by Crystal 2006/06/1
              With rsDefTmp
                    '7179 客編需重新判斷,如果客編有多筆代表有母子戶，如果只有一筆代表就是子戶一筆 By Kin 2016/03/29
                    Dim aCustid As String
                    aCustid = rsTmp("CustId")
                    If isCrossCustCombine Then
                       
                       If Not GetRS(rsCustId, "Select Distinct Custid From " & strGetOwner & "SO033 Where " & aBillNoType & "='" & rsTmp("BillNO") & "'", _
                                    gcnGi, adUseClient, adOpenKeyset) Then
                                    Exit Function
                        Else
                            If rsCustId.RecordCount = 1 Then
                                aCustid = rsCustId(0)
                            End If
                        End If
                    End If
                   .AddNew
                   .Fields("DataNo") = "1"
                   '資料別
                   .Fields("Case") = IIf(Left(GetFieldValue(rsTmp, "AccountNo") & "", 6) = "000000", "G", "P")
                   '存款別 2-2
                   .Fields("Comid") = GetFieldValue(rsBank, "BankId") & ""
                   '事業單位代號 3-5
                   .Fields("Citem") = GetString(strPostUnit & "", 4, GIRIGHT)
                   '區處站所代號 6-9
                   '2010/07/15
                   .Fields("Acctno") = Mid(GetFieldValue(rsTmp, "AccountNo") & "", 1, 14)
                   '繳費帳號 20-33
                   '#6320 增加客編條件 By Kin 2012/09/18
                   strAccNameID = GetRsValue("Select AccountNameID From " & strGetOwner & "SO106 " & _
                                                " Where AccountID='" & Mid(GetFieldValue(rsTmp, "AccountNo") & "", 1, 14) & _
                                                "' And StopFlag<>1 AND CUSTID = " & aCustid, gcnGi) & ""
                   .Fields("Idno") = strAccNameID
                   '身份證字號 34-43
                   .Fields("Amount") = lngShouldAmt   ' GetFieldValue(rsTmp, "ShouldAMT") & ""
                   '繳費金額 44-54
                   
                   '*********************************************************************************************************
                   '#3527 參考號為2，使用新的POS規格 By Kin 2007/10/04
                    If intRefNo <> 2 Then
                        .Fields("Custno") = GetString(strAccNameID & "", 10) & GetFieldValue(rsTmp, "BillNo") & ""
                    Else
                        '#3527 測試不OK，55-70個字元要抓16碼，不是15碼 By Kin 2007/10/17
                        strCitibankATM = GetString(Trim(GetFieldValue(rsTmp, "BillNo") & ""), 16)
                        .Fields("CustNo") = strCitibankATM & Space(3) & "N"
                    End If
                   '*********************************************************************************************************
                   
                   '問題集2386   Lydia提 身份證字號(10)+用戶資料(11位),所以原先20位改21位,保留欄少一位 Edit by Crystal 2006/06/1
                   '事業單位用戶資料 55-75
                   .Fields("Space1") = ""
                   '保留欄 76-78
                   .Fields("Flag") = ""
                   '代繳結果 79-80
                   '2010/07/15 Jacky 5680 Jacy
                    If rsBank("RefNo2") = 1 Then
                        '.Fields("ReciDay") = Replace(gdaAgency.GetOriginalValue, "/", "")

                        .Fields("ReciDay") = Format(Trim(Val(gdaAgency.GetValue) - 19110000), "0000000")
                        '繳費日期 10-16
                        .Fields("Recimnth") = ""
                        '保留欄 17-19
                       
                       .Fields("NewRecimnth") = Format(Trim(Val(GetFieldValue(rsTmp, "Shoulddate") & "") - 191100), "00000")
                       '繳費月份    81-85 民國年(3位)YYYMM
                       .Fields("Space2") = ""
                       '保留欄      86-90 空白
                    Else
                        .Fields("ReciDay") = Trim(Val(gdaAgency.GetValue) - 19110000)
                        '繳費日期 10-15
                        .Fields("Recimnth") = Trim(Val(GetFieldValue(rsTmp, "Shoulddate") & "") - 191100)
                        '繳費月份 16-19
                    End If
                   'Insert Into MDB(SO3271A)
              End With
                cnn.Execute Replace("Insert Into SO3271A (AccountNo, TransDate, ShouldAmt, BillNo_Old, BillNo_New) Values (" & _
                     GetNullString(rsTmp("AccountNo")) & "," & GetNullString(gdaTransfer.GetValue(True), giDateV, giAccessDb) & "," & _
                     GetNullString(lngShouldAmt) & "," & GetNullString(strBillNo_Old) & "," & _
                     GetNullString(strBillNo) & ")", Chr(0), "")
                strUPDUCCode = GetLiteraUpdateSQL(IIf(intRefNo = 2, 3, Val(strMedia)), StrTableName, strWhere, _
                                 strChoose, IIf(Len(strHavingOutZero) = 0, True, False), rsTmp("BillNo"))
                
                If blnUpdUcCode Then
                    
                    If Not GetRS(rsUpdUCCode, strUPDUCCode, gcnGi, adUseClient, adOpenKeyset, adLockPessimistic) Then Exit Function
                     Do While Not rsUpdUCCode.EOF
           
                        If (IsPayKindOK(rsUpdUCCode, gdaTransfer.GetValue & "", giSingle, IIf(intRefNo = 2, 3, Val(strMedia)))) _
                            Or (Not blnChkPayKind) Then
                            rsUpdUCCode("UPDEN") = objStorePara.uUpdEn
                            rsUpdUCCode("UPDTIME") = strUpdTime
                            rsUpdUCCode("UCcode") = strUcCode
                            rsUpdUCCode("UCName") = strUcName
                            rsUpdUCCode.Update
                        End If
                        rsUpdUCCode.MoveNext
                    Loop
                        CloseRecordset rsUpdUCCode
                 End If
                
                
lNextRcd:
           strOldBillNo = rsTmp("BillNo") & ""
           strOldAccountId = rsTmp("AccountNo")
           rsTmp.MoveNext
           DoEvents
        Wend
        cnn.CommitTrans
        
        '**************************************************************************************************************************
        '#3417 更新UCCode與UCName By Kin 2007/12/04
        '#4388 增加更新異動人員與異動時間 By Kin 2009/04/29
        If blnUpdUcCode And 1 = 0 Then
            
            If Not GetRS(rsUpdUCCode, strUPDUCCode, gcnGi, adUseClient, adOpenKeyset, adLockPessimistic) Then Exit Function
            'Set rsUpdUCCode.ActiveConnection = Nothing
            Do While Not rsUpdUCCode.EOF
                '#5683 如果收視截止日小於畫面條件不要UPD By Kin 2010/08/05
                '#5843 如果設定不檢核,則強迫出帳 By Kin 2010/11/24
                If (IsPayKindOK(rsUpdUCCode, gdaTransfer.GetValue & "", giSingle, IIf(intRefNo = 2, 3, Val(strMedia)))) _
                    Or (Not blnChkPayKind) Then
'                    Dim aBillField As String
'                    Select Case IIf(intRefNo = 2, 3, Val(strMedia))
'                        Case 0
'                            aBillField = " BILLNO "
'                        Case 1
'                            aBillField = " MEDIABILLNO "
'                        Case Else
'                            aBillField = " CitibankATM "
'                    End Select
'                    Dim aSQL As String
'                    aSQL = "UPDATE " & strGetOwner & "SO033 " & _
'                        " SET UPDEN = '" & objStorePara.uUpdEn & "'," & _
'                        " UPDTIME = '" & strUpdTime & "'," & _
'                        " UCCODE = " & strUcCode & "," & _
'                        " UCNAME = '" & strUcName & "' " & _
'                        " WHERE " & aBillField & "= '" & rsUpdUCCode("BILLNO") & "'"
'                    gcnGi.Execute aSQL
                    rsUpdUCCode("UPDEN") = objStorePara.uUpdEn
                    rsUpdUCCode("UPDTIME") = strUpdTime
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
        CloseRecordset rsUpdUCCode
        If strFlowId = 0 Then
            rsPara.Close
            Set rsPara = Nothing
        End If
        CloseRecordset rsCD013
        CloseRecordset rsCustId
        If Not subWriteLine(lngCount, lngTotAmt) Then
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

Private Function subWriteLine(ByVal lngCount As Long, ByVal lngTotAmt As Long) As Boolean
  On Error GoTo ChkErr
    Dim varData As Variant
    Dim strData As String
    Dim lngloop As Long
    Dim strYMdata As String
    Dim strLastData As String
    subWriteLine = False
    With rsDefTmp
         If .BOF Or .EOF Then Exit Function
         .MoveFirst
         varData = .GetRows
         '內容
         For lngloop = 0 To .RecordCount - 1
            '新舊格式取的欄位不同 2010/07/15 5680 Jacky
            If Val(rsBank("RefNo2") & "") = 0 Then
                strYMdata = GetString(varData(4, lngloop), 6) & _
                            GetString(varData(5, lngloop), 4)
                strLastData = ""
            Else
                strYMdata = GetString(varData(4, lngloop), 7) & _
                            GetString(varData(5, lngloop), 3)
                strLastData = GetString(varData(12, lngloop), 5) & GetString(varData(13, lngloop), 5)
            End If

            WriteLineData GetString(varData(0, lngloop), 1) & _
                        GetString(varData(1, lngloop), 1) & _
                        GetString(varData(2, lngloop), 3) & _
                        GetString(varData(3, lngloop), 4) & _
                        strYMdata & _
                        GetString(varData(6, lngloop), 14) & _
                        GetString(varData(7, lngloop), 10, giLeft) & _
                        GetString(varData(8, lngloop), 9, GIRIGHT, True) & "00" & _
                        GetString(varData(9, lngloop), 21, giLeft) & _
                        GetString(varData(10, lngloop), 3) & _
                        GetString(varData(11, lngloop), 2) & _
                        strLastData, 0
             DoEvents
         Next
'         '末筆
        '#3583 測試未OK,參考號為2時,最未筆要保留5個空白 By Kin 2007/10/30
        
        '新舊格式取的欄位不同 2010/07/15 5680 Jacky
        If Val(rsBank("RefNo2") & "") = 0 Then
            strYMdata = Left(strYMdata, 6) & "0000"
            strLastData = Space(5)
        Else
            strYMdata = strYMdata
            strLastData = Space(15)
        End If
        
        WriteLineData "2" & GetString("", 1) & _
                   GetString(rsBank("BankId") & "", 3) & _
                   GetString(strPostUnit, 4, GIRIGHT) & _
                   strYMdata & _
                   GetString(lngCount, 7, GIRIGHT, True) & _
                   GetString(lngTotAmt, 11, GIRIGHT, True) & "00" & _
                   GetString(rsBank("ActNO") & "", 8) & _
                   GetString(rsBank("ActNO") & "", 8) & _
                   GetString("", 7, giLeft, True) & _
                   GetString("", 13, giLeft, True) & _
                   strLastData, 0
                    
    End With
    subWriteLine = True
    On Error Resume Next
    rsDefTmp.Close
    Set rsDefTmp = Nothing
    rsBank.Close
    Set rsBank = Nothing
    Exit Function
ChkErr:
    subWriteLine = False
    Call ErrSub(Me.Name, "subWriteLine")
    Resume 0
End Function

'定義記憶體Recordset
Private Sub DefineRs()
    On Error GoTo ChkErr
    With rsDefTmp
        If .State = adStateOpen Then .Close
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        
        .Fields.Append ("DataNo"), adBSTR, 1, adFldIsNullable       '資料別
        .Fields.Append ("Case"), adBSTR, 1, adFldIsNullable         '存款別
        .Fields.Append ("Comid"), adBSTR, 3, adFldIsNullable        '事業單位代號
        .Fields.Append ("Citem"), adBSTR, 4, adFldIsNullable        '區處站所代號
        .Fields.Append ("Reciday"), adBSTR, 7, adFldIsNullable      '繳費日期
        .Fields.Append ("Recimnth"), adBSTR, 4, adFldIsNullable     '繳費月份
        .Fields.Append ("Acctno"), adBSTR, 14, adFldIsNullable      '繳費帳號
        .Fields.Append ("Idno"), adBSTR, 10, adFldIsNullable        '身份證字號
        .Fields.Append ("Amount"), adBSTR, 11, adFldIsNullable      '繳費金額
        .Fields.Append ("Custno"), adBSTR, 21, adFldIsNullable      '事業單位用戶資料
        .Fields.Append ("Space1"), adBSTR, 3, adFldIsNullable       '保留欄
        .Fields.Append ("Flag"), adBSTR, 2, adFldIsNullable         '代繳結果
        
        .Fields.Append ("NewRecimnth"), adBSTR, 5, adFldIsNullable    '新繳費月份
        .Fields.Append ("Space2"), adBSTR, 5, adFldIsNullable         '保留
        
        
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
        txtDataPath = FolderDialog("開始")
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "cmdDataPath_Click")
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
    Dim strFileName1 As String
    Dim lngFilePath As Long
    Dim strFileName As String
    If Not ChkDTok Then Exit Sub
    If Not IsDataOk Then Exit Sub
    '找出匯出的檔案名稱
    strFileName1 = GetRsValue("Select FileName1 From " & strGetOwner & "CD018 Where CodeNo=" & strBankId, gcnGi) & ""
    '修改儲存在畫面的記錄,只能儲存資料的路徑,不能連檔名都儲,不然會與strFileName1衝突
    If InStr(txtDataPath, strFileName1) = 0 Then
      lngFilePath = InStrRev(strFileName1, "\")
      If lngFilePath <> 0 Then strFileName1 = Mid(strFileName1, lngFilePath + 1, 9999)
        strFileName = txtDataPath & IIf(Right(txtDataPath, 1) = "\", "", "\") & strFileName1
    End If
    If strFileName1 = "" Then MsgBox "無設定資料檔名稱!!", vbExclamation, gimsgPrompt: txtDataPath.SetFocus: Exit Sub
    '#5596 測試不OK 要更改檔名 By Kin 2010/04/30
    If gdaTransfer.GetValue <> "" Then
        strFileName = Replace(strFileName, "YYYYMMDD", gdaTransfer.GetValue)
    End If
'    If OpenFile(txtDataPath, 0, True) = False Then txtDataPath.SetFocus: CloseFS: Exit Sub
    If OpenFile(strFileName, 0, True) = False Then txtDataPath.SetFocus: CloseFS: Exit Sub
    If OpenFile(txtMemoPath, 1, True) = False Then txtMemoPath.SetFocus: CloseFS: Exit Sub
    If OpenFile(txtErrPath, 2, True) = False Then txtErrPath.SetFocus: CloseFS: Exit Sub
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

        If gdaTransfer.GetValue = "" Then strErrMsg = "自動轉帳扣款日期": gdaTransfer.SetFocus: GoTo Warning
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
    Dim rsInv As New ADODB.Recordset
    Dim strSQL As String
    
        strSQL = "Select InvoiceId From " & strGetOwner & "SO041 Where CompCode =" & strCompCode
        If Not GetRS(rsInv, strSQL, gcnGi) Then Exit Sub
'        txtInVoice.Text = GetFieldValue(rsInv, "InvoiceId") & ""
        
        rsInv.Close
        Set rsInv = Nothing

        If Dir(strErrPath & "\" & strPrgName & "Out.log") = "" Then Exit Sub
            Set LogFile = fso.OpenTextFile(strErrPath & "\" & strPrgName & "Out.log", ForReading)
            With LogFile
                    If Not .AtEndOfStream Then gdaAgency.Text = .ReadLine & ""
                       '臨櫃代收繳費期限
                    If Not .AtEndOfStream Then gdaTransfer.Text = .ReadLine & ""
                        '自動轉帳繳費期限
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
'    Dim strFilePath As String
'    strFilePath = txtDataPath.Text
'    strFilePath = Left(strFilePath, InStrRev(strFilePath, "\")) - 1
    
        Set LogFile = fso.CreateTextFile(strErrPath & "\" & strPrgName & "Out.log", True)
         With LogFile
               Call .WriteLine(gdaAgency.Text)   '臨櫃代收繳費期限
                .WriteLine (gdaTransfer.Text)    '自動轉帳繳費期限
                .WriteLine (txtDataPath)     '資料檔
                .WriteLine (txtErrPath)          '問題檔
                .WriteLine (txtMemoPath)         '備註檔
                .WriteLine (txtMemo1)            '備註1
                .WriteLine (txtMemo2)            '備註2
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

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
    If Len(sqlTmpViewName) > 0 Then
        gcnGi.Execute "Drop View " & strGetOwner & sqlTmpViewName
        sqlTmpViewName = Empty
    End If
End Sub
