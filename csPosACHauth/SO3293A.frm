VERSION 5.00
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{A0C51418-555F-11D3-BC10-0050BA01A3A9}#9.9#0"; "GiMulti.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{074771B7-6ADE-11D5-B178-00485456B028}#2.1#0"; "csMulti.ocx"
Begin VB.Form frmSO3293A 
   BorderStyle     =   1  '單線固定
   Caption         =   "郵局核印作業 [SO3293A]"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7875
   Icon            =   "SO3293A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   7875
   StartUpPosition =   2  '螢幕中央
   Begin VB.CheckBox chkStop 
      Caption         =   "是否將授權失敗停用"
      Height          =   180
      Left            =   3240
      TabIndex        =   30
      Top             =   4800
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox txtSQL 
      Height          =   525
      Left            =   3330
      MultiLine       =   -1  'True
      ScrollBars      =   2  '垂直捲軸
      TabIndex        =   29
      Top             =   2850
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Frame Frame3 
      Height          =   1605
      Left            =   165
      TabIndex        =   21
      Top             =   90
      Width           =   7605
      Begin VB.ComboBox CmbApplyType 
         Height          =   300
         ItemData        =   "SO3293A.frx":0442
         Left            =   1035
         List            =   "SO3293A.frx":0444
         TabIndex        =   0
         Top             =   210
         Width           =   2640
      End
      Begin Gi_Date.GiDate gdaPropDate2 
         Height          =   345
         Left            =   2580
         TabIndex        =   2
         Top             =   1110
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
      Begin Gi_Date.GiDate gdaPropDate1 
         Height          =   345
         Left            =   1035
         TabIndex        =   1
         Top             =   1110
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
      Begin Gi_Date.GiDate gdaStopDate2 
         Height          =   345
         Left            =   6150
         TabIndex        =   4
         Top             =   1110
         Width           =   1095
         _ExtentX        =   1931
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
      Begin Gi_Date.GiDate gdaStopDate1 
         Height          =   345
         Left            =   4605
         TabIndex        =   3
         Top             =   1110
         Width           =   1095
         _ExtentX        =   1931
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
      Begin prjGiList.GiList gilCompCode 
         Height          =   315
         Left            =   4605
         TabIndex        =   5
         Top             =   210
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
      Begin Gi_Date.GiDate gdaSendDate 
         Height          =   345
         Left            =   1020
         TabIndex        =   33
         Top             =   630
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "送件日期"
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
         Left            =   180
         TabIndex        =   32
         Top             =   720
         Width           =   780
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "公司別"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   3810
         TabIndex        =   28
         Top             =   270
         Width           =   585
      End
      Begin VB.Label lblStopDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "終止日期"
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
         Left            =   3795
         TabIndex        =   26
         Top             =   1185
         Width           =   750
      End
      Begin VB.Label lblPropDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "申請日期"
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
         Left            =   165
         TabIndex        =   25
         Top             =   1200
         Width           =   780
      End
      Begin VB.Label lblStop 
         AutoSize        =   -1  'True
         Caption         =   "至"
         Height          =   195
         Left            =   5805
         TabIndex        =   24
         Top             =   1185
         Width           =   165
      End
      Begin VB.Label lbluntil 
         AutoSize        =   -1  'True
         Caption         =   "至"
         Height          =   195
         Left            =   2250
         TabIndex        =   23
         Top             =   1200
         Width           =   195
      End
      Begin VB.Label lblComp 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "申請代號"
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
         Left            =   180
         TabIndex        =   22
         Top             =   240
         Width           =   780
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1140
      Left            =   165
      TabIndex        =   18
      Top             =   3390
      Width           =   7590
      Begin VB.TextBox txtDataPath 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1800
         TabIndex        =   9
         ToolTipText     =   "請輸入字檔之路徑及檔名！"
         Top             =   180
         Width           =   5145
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
         Left            =   7005
         TabIndex        =   10
         Top             =   195
         Width           =   450
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
         Height          =   285
         Left            =   7005
         TabIndex        =   12
         Top             =   615
         Width           =   450
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
         Height          =   315
         Left            =   1800
         TabIndex        =   11
         Top             =   585
         Width           =   5145
      End
      Begin VB.Label lblDatapath 
         AutoSize        =   -1  'True
         Caption         =   "資料檔路徑及檔名"
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
         Left            =   180
         TabIndex        =   20
         Top             =   240
         Width           =   1560
      End
      Begin VB.Label lblErrpath 
         AutoSize        =   -1  'True
         Caption         =   "問題檔路徑及檔名"
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
         Left            =   180
         TabIndex        =   19
         Top             =   615
         Width           =   1560
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   165
      TabIndex        =   16
      Top             =   1740
      Width           =   7605
      Begin prjGiList.GiList gilBank 
         Height          =   315
         Left            =   1170
         TabIndex        =   35
         Top             =   210
         Width           =   3735
         _ExtentX        =   6588
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
         FldWidth1       =   900
         FldWidth2       =   2500
      End
      Begin CS_Multi.CSmulti gimCitemCode 
         Height          =   405
         Left            =   120
         TabIndex        =   34
         Top             =   990
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   714
         ButtonCaption   =   "收 費 項 目"
      End
      Begin VB.ComboBox cboACHbankType 
         Height          =   300
         ItemData        =   "SO3293A.frx":0446
         Left            =   4800
         List            =   "SO3293A.frx":0448
         Style           =   2  '單純下拉式
         TabIndex        =   6
         Top             =   240
         Visible         =   0   'False
         Width           =   2730
      End
      Begin VB.TextBox txtACHTNo 
         Height          =   345
         Left            =   5430
         TabIndex        =   8
         Top             =   1230
         Visible         =   0   'False
         Width           =   1830
      End
      Begin Gi_Multi.GiMulti gimBankId 
         Height          =   375
         Left            =   270
         TabIndex        =   7
         Top             =   1200
         Visible         =   0   'False
         Width           =   7260
         _ExtentX        =   12806
         _ExtentY        =   661
         ButtonCaption   =   "轉帳銀行名稱"
         FontSize        =   9
         FontName        =   "新細明體"
      End
      Begin Gi_Multi.GiMulti gimBillHeadFmt 
         Height          =   375
         Left            =   120
         TabIndex        =   31
         Top             =   600
         Width           =   7275
         _ExtentX        =   12832
         _ExtentY        =   661
         ButtonCaption   =   "多帳戶產生依據"
         DataType        =   2
         FontSize        =   9
         FontName        =   "新細明體"
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "交易代碼"
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
         Left            =   4455
         TabIndex        =   27
         Top             =   1320
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "承 辦 銀 行"
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
         Left            =   165
         TabIndex        =   17
         Top             =   270
         Width           =   915
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "授權扣款提出"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   0
      Left            =   180
      TabIndex        =   13
      Top             =   4665
      Width           =   1410
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
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
      Height          =   465
      Left            =   6255
      TabIndex        =   15
      Top             =   4695
      Width           =   1410
   End
   Begin MSComDlg.CommonDialog cnd1 
      Left            =   5235
      Top             =   4665
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "授權扣款提回"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   1
      Left            =   1665
      TabIndex        =   14
      Top             =   4665
      Width           =   1410
   End
End
Attribute VB_Name = "frmSO3293A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FSO As New FileSystemObject
Private strPath As String
Private strPrgName As String
Private strChoose As String
Private strWhere As String
Private strUpdEn As String
Private strUpdName As String
Private strUpdTime As String
Private strNowTime As String
Private DataPath As TextStream
Private ErrPath As TextStream
Private lngErrCount As Long
Private intSeq As Integer
Private strAchCid As String
Private rsTmp As New ADODB.Recordset
Dim rsBillHeadFmt As New ADODB.Recordset
Dim rsSO106 As New ADODB.Recordset
Dim strReturnSQL As String
Dim strDate As String
Dim strFrom  As String
Dim strErrAuth As String
Dim strINCD008Where As String
Dim blnSnactionNotNull As Boolean
Dim strACHTNO As String
Dim strACHDesc As String
Dim strBillHeadFmt As String
Dim strInACHTNO As String
Dim strInBillHeadFmt As String
Dim strInACHDesc As String
Dim strCancelWhere As String
Private blnHasUpd As Boolean
Private blnStopAll As Boolean
Private strNewRowId As String
Private strMasterId As String
Private blnCopy As Boolean
Private strUpdRowIds As String
Private strSO106AMasterid As String
Private Function applyAuthFail(ByVal strData As String) As Boolean
  On Error GoTo ChkErr
    Dim strCustId As String
    Dim blnBeginTrans As Boolean
    Dim rsCustid As New ADODB.Recordset
    Dim rsSO106A As New ADODB.Recordset
    Dim strSQL As String
    Dim strErrDescription As String
    
    blnBeginTrans = False
    strSQL = getApplySO106SQL(strData)
   
    If Not GetRS(rsCustid, strSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic, , , True) Then Exit Function
     If rsCustid.RecordCount = 0 Then
        strCustId = ""
        strErrDescription = "此筆ACHCustid 未存在於資料庫 "
        Call InsertToErr(strData, strErrDescription, strCustId, Trim(Mid(strData, 28, 14)))
        Exit Function
    End If
    If Not GetRS(rsSO106A, GetSO106A("1", rsCustid("MasterId"), strData), gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
    If rsSO106A.RecordCount = 0 Then
        strErrDescription = "找不到SO106A對應資料,該筆資料可能已授權或取消授權 "
        Call InsertToErr(strData, strErrDescription, rsCustid("CustId"), GetAccID(strData))
        Exit Function
    End If
    gcnGi.BeginTrans
    blnBeginTrans = True
    Call UpdateSO106A(rsSO106A, "E")
    
    If AlterSO106(gcnGi, , "E", strData, strErrAuth) = False Then lngErrCount = lngErrCount
        
    If Not UpdateSO106(rsCustid("RowId") & "", rsSO106A) Then
        strErrDescription = "取消SO106 ACH項目錯誤！"
        Call InsertToErr(strData, strErrDescription, strCustId, GetAccID(strData))
    End If
    intSeq = intSeq + 1
    gcnGi.CommitTrans
    On Error Resume Next
    CloseRecordset rsCustid
    CloseRecordset rsSO106A
    Exit Function
ChkErr:
    If blnBeginTrans Then
        gcnGi.RollbackTrans
    End If
    Call ErrSub(Me.Name, "applyAuthFail")
End Function
Private Function getApplySO106SQL(ByVal strData As String) As String
    Dim applyNumber As Integer
    Dim aWhere As String
    aWhere = ""
    applyNumber = Val(Mid(strData, 26, 1) & "")
    Select Case applyNumber
        Case 1
            aWhere = " And StopFlag <> 1 "
        Case 4
            aWhere = " And StopFlag = 1 "
    End Select
    getApplySO106SQL = "Select RowId,CUSTID,MasterId From " & GetOwner & "SO106 A Where ACHCUSTID='" & GetACHCustID(strData) & _
                             "' And LPAD(AccountID,14,'0')='" & GetString(Mid(strData, 28, 14), 14, giRight, True) & "'" & _
                             " And RPAD(nvl(AccountNameID,'0'),10,'0')='" & GetString(Mid(strData, 62, 10), 10, giLeft, True) & "'" & aWhere

End Function
Private Function StopAuth(ByVal strData As String) As Boolean
  On Error GoTo ChkErr
    Dim strSQL As String
    Dim strCustId As String
    Dim strErrDescription As String
    Dim rsCustid As New ADODB.Recordset
    Dim rsSO106A As New ADODB.Recordset
    Dim blnBeginTrans As Boolean
    Dim applyNumber As Integer
    applyNumber = Val(Mid(strData, 26, 1) & "")
    blnBeginTrans = False
    strSQL = getApplySO106SQL(strData)
    If Not GetRS(rsCustid, strSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic, , , True) Then Exit Function
    If rsCustid.RecordCount = 0 Then
        strCustId = ""
        strErrDescription = "此筆ACHCustid 未存在於資料庫 "
        Call InsertToErr(strData, strErrDescription, strCustId, Trim(Mid(strData, 28, 14)))
        Exit Function
    End If
    strCustId = rsCustid("CustId")
    If Not GetRS(rsSO106A, GetSO106A(applyNumber, rsCustid("MasterId"), strData), gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
    If rsSO106A.RecordCount = 0 Then
        strErrDescription = "找不到SO106A對應資料,該筆資料可能已授權或取消授權 "
        Call InsertToErr(strData, strErrDescription, rsCustid("CustId"), GetAccID(strData))
        Exit Function
    End If
    gcnGi.BeginTrans
    blnBeginTrans = True
     If UpdateSO106A(rsSO106A, CStr(applyNumber)) Then
            If AlterSO106(gcnGi, rsCustid, CStr(applyNumber), strData) Then
                If StopSO003(rsCustid("RowId") & "", rsSO106A) Then
                    intSeq = intSeq + 1
                Else
                    strErrDescription = "停用SO003資料錯誤！"
                    Call InsertToErr(strData, strErrDescription, strCustId, GetAccID(strData))
                End If
            Else
                strErrDescription = "未成功更新SO106！請確認用戶識別碼、帳戶、核准日期或送出日期是否錯誤！"
                Call InsertToErr(strData, strErrDescription, strCustId, GetAccID(strData))
            End If
    Else
        strErrDescription = "未成功更新SO106A！"
        Call InsertToErr(strData, strErrDescription, strCustId, GetAccID(strData))
    End If
    gcnGi.CommitTrans
                                        
    On Error Resume Next
    CloseRecordset rsCustid
    CloseRecordset rsSO106A
    Exit Function
ChkErr:
  If blnBeginTrans Then
    gcnGi.RollbackTrans
  End If
  Call ErrSub(Me.Name, "StopAuth")
End Function
Private Function ReCoverSO106(rsSO106A As ADODB.Recordset, ByVal strCustId As String) As Boolean
   On Error GoTo ChkErr
    Dim sql As String
    Dim rsSO106Upd As New Recordset
    Dim rsSO003 As New ADODB.Recordset
    Dim citemCode As String
    sql = "Select * From " & GetOwner & "SO106 A Where Masterid = " & rsSO106A("MasterId") & " And CompCode = " & gilCompCode.GetCodeNo
    If Not GetRS(rsSO106Upd, sql, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
    If InStr(1, rsSO106Upd("ACHTNo"), "'" & rsSO106A("ACHTNo") & "'") <= 0 Then
        rsSO106Upd("ACHTNo") = IIf(Len(rsSO106Upd("ACHTNo") & "") = 0, "'" & rsSO106A("ACHTNo") & "'", rsSO106Upd("ACHTNo") & ",'" & rsSO106A("ACHTNo") & "'")
        rsSO106Upd("ACHTDESC") = IIf(Len(rsSO106Upd("ACHTDESC") & "") = 0, "'" & rsSO106A("ACHDesc") & "'", rsSO106Upd("ACHDesc") & ",'" & rsSO106A("ACHDesc") & "'")
        rsSO106Upd.Update
    End If
    citemCode = rsSO106A("CitemCodeStr") & ""
    If Len(citemCode) > 0 Then
        sql = "Select SeqNo From " & GetOwner & "SO003 " & _
            " Where Custid = " & strCustId & " And CitemCode In ( " & citemCode & " ) " & _
            " And CompCode = " & gilCompCode.GetCodeNo
        If Not GetRS(rsSO003, sql, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
        If rsSO003.RecordCount > 0 Then
            rsSO003.MoveFirst
            rsSO106Upd.MoveFirst
            Do While Not rsSO003.EOF
                If InStr(1, rsSO106Upd("CitemStr"), "'" & rsSO003("SeqNo") & "'") <= 0 Then
                    rsSO106Upd("CitemStr") = IIf(Len(rsSO106Upd("CitemStr") & "") = 0, "'" & rsSO003("SeqNo") & "'", rsSO106Upd("CitemStr") & ",'" & rsSO003("SeqNo") & "'")
                    rsSO106Upd.Update
                End If
                rsSO003.MoveNext
            Loop
        End If
    End If
    ReCoverSO106 = True
    On Error Resume Next
    CloseRecordset rsSO106Upd
    CloseRecordset rsSO003
    Exit Function
ChkErr:
  Call ErrSub(Me.Name, "ReCoverSO106")
End Function

Private Function applyAuthOK(ByVal strData As String) As Boolean
    On Error GoTo ChkErr
    Dim strSQL As String
    Dim strCustId As String
    Dim strErrDescription As String
    Dim applyNumber As Integer
    Dim rsCustid As New ADODB.Recordset
    Dim rsSO106A As New ADODB.Recordset
    Dim blnBeginTrans As Boolean
    Dim retState As String
    blnBeginTrans = False
    strSQL = getApplySO106SQL(strData)
    If Not GetRS(rsCustid, strSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic, , , True) Then Exit Function
    If rsCustid.RecordCount = 0 Then
        strCustId = ""
        strErrDescription = "此筆ACHCustid 未存在於資料庫 "
        Call InsertToErr(strData, strErrDescription, strCustId, Trim(Mid(strData, 28, 14)))
        Exit Function
    End If
    applyNumber = Val(Mid(strData, 26, 1) & "")
    If Not GetRS(rsSO106A, GetSO106A(CStr(applyNumber), rsCustid("MasterId"), strData), gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
    If rsSO106A.RecordCount = 0 Then
        strErrDescription = "找不到SO106A對應資料,該筆資料可能已授權或取消授權 "
        Call InsertToErr(strData, strErrDescription, rsCustid("CustId"), GetAccID(strData))
        Exit Function
    End If
    gcnGi.BeginTrans
    blnBeginTrans = True
    strCustId = rsCustid("CustId")
    retState = "RETURNOK"
    If applyNumber = 4 Then retState = CStr(applyNumber)
    
    If AlterSO106(gcnGi, rsCustid, retState, strData) Then
        If InsSO002A(gcnGi, , strData, strCustId) = False Then
            strErrDescription = "未成功新增SO002A "
            Call InsertToErr(strData, strErrDescription, strCustId, GetAccID(strData))
            gcnGi.RollbackTrans
            Exit Function
        End If
        If applyNumber = 4 Then
            If Not ReCoverSO106(rsSO106A, rsCustid("CustId") & "") Then
                strErrDescription = "回復SO106失敗 "
                Call InsertToErr(strData, strErrDescription, strCustId, GetAccID(strData))
                gcnGi.RollbackTrans
                Exit Function
            End If
        End If
        If AlterSO0032(gcnGi, rsSO106A, , strData, strCustId) = False Then
            strErrDescription = "未成功更新SO003 "
            Call InsertToErr(strData, strErrDescription, strCustId, GetAccID(strData))
            gcnGi.RollbackTrans
            Exit Function
        Else
            '#更新SO106A,如果連SO106A也更新成功時則將成功筆數+1
            If UpdateSO106A(rsSO106A, "RETURNOK") Then
                intSeq = intSeq + 1
            Else
                strErrDescription = "未成功更新SO106A "
                Call InsertToErr(strData, strErrDescription, strCustId, GetAccID(strData))
                gcnGi.RollbackTrans
            Exit Function
            End If
        End If
    Else
        strErrDescription = "未成功更新SO106! 可能為該筆資料已有核准日期..請確認用戶識別碼、帳戶、核准日期或送出日期是否錯誤！"
        Call InsertToErr(strData, strErrDescription, strCustId, GetAccID(strData))
        gcnGi.RollbackTrans
        Exit Function
    End If
    gcnGi.CommitTrans
    On Error Resume Next
    CloseRecordset rsCustid
    CloseRecordset rsSO106A
    Exit Function
ChkErr:
    If blnBeginTrans Then
        gcnGi.RollbackTrans
    End If
    Call ErrSub(Me.Name, "applyAuthOK")
End Function
Private Sub FileReturn()
  On Error GoTo ChkErr
        If Not OpenData(1) Then Exit Sub
        Dim O As String
        Dim aryContext As Variant
        Dim i  As Integer
        Dim blnFileOK As Boolean
        Dim applyNumber As String
        Dim strSQL As String
        Dim lngTime As Long
        blnFileOK = False
        O = DataPath.ReadAll
        aryContext = Split(O, vbCrLf)
        lngTime = Timer
        For i = UBound(aryContext) - 1 To LBound(aryContext) Step -1
            If Len(aryContext(i) & "") > 0 Then
                If Mid(aryContext(i), 1, 1) = 2 Then
                    strDate = Trim(Mid(aryContext(i), 27, 8) & "")
                    If (Len(strDate & "") = 0) Or (Len(strDate) <> 8) Or (Not IsNumeric(strDate)) Then
                        strDate = Format(RightNow, "yyyymmdd")
                    End If
                    
                    If Val(Mid(aryContext(i), 35, 6) & "") > 0 Then
                        blnFileOK = True
                        Exit For
                    End If
                    If Val(Mid(aryContext(i), 41, 6) & "") > 0 Then
                        blnFileOK = True
                        Exit For
                    End If
                    
                End If
            End If
        Next
        If Not blnFileOK Then
            MsgBox "該檔案不符合提回格式！", vbInformation, "訊息"
            Exit Sub
        End If
        
        For i = LBound(aryContext) To UBound(aryContext) - 1
            If Mid(aryContext(i), 1, 1) = "1" Then
                applyNumber = Val(Mid(aryContext(i), 26, 1) & "")
                Select Case applyNumber
                    Case 1
                         If Val(Trim(Mid(aryContext(i), 72, 2) & "")) = 0 And Val(Trim(Mid(aryContext(i), 74, 1) & "")) = 0 Then
                            Call applyAuthOK(aryContext(i))
                         Else
                            strErrAuth = getApplyErrMsg(aryContext(i))
                            Call applyAuthFail(aryContext(i))
                         End If
                    Case 2, 3
                        Call StopAuth(aryContext(i))
                    Case 4
                        Call applyAuthOK(aryContext(i))
                End Select
                    
            End If
          
        Next i
        MsgBox "已完成資料筆數共" & intSeq & "筆," & vbCrLf & vbCrLf & _
          "問題筆數共" & lngErrCount & "筆," & vbCrLf & vbCrLf & _
          "共花費:" & (Timer - lngTime) \ 1 & "秒"
    Exit Sub
       
ChkErr:
    Call ErrSub(Me.Name, "FileReturn")
  
End Sub
Private Sub cboACHbankType_Click()
  Select Case cboACHbankType.ItemData(cboACHbankType.ListIndex)
       Case 0
'           e_CreditCardType = CardType_default
'           chkChinaBank.Visible = True
'           chkChinaBank.Value = 0
'           fraTCB.Visible = True
            
           Call SetgiMulti(gimBankId, "CodeNo", "Description", "CD018", "代碼", "名稱", _
                                   " Where UPPER(PrgName) = 'ACHTRANREFER'  AND  COMPCODE =" & _
                                   gilCompCode.GetCodeNo & " And StopFlag=0")
    End Select
End Sub

Private Sub cboBillHeadFmt_Click()
    On Error Resume Next
    Dim strCitemCode As String
    Dim rsCD019  As New ADODB.Recordset
        If rsBillHeadFmt.RecordCount = 0 Then Exit Sub
        'rsBillHeadFmt.AbsolutePosition = cboBillHeadFmt.ListIndex + 1
        If Not GetRS(rsCD019, "Select CodeNo From " & GetOwner & "CD019 Where CodeNo In (" & rsBillHeadFmt("CitemCodeStr") & ") And StopFlag=0") Then Exit Sub
        If rsCD019.EOF Then Exit Sub
        'txtACHTNo = GetRsValue("Select ACHTNO From " & GetOwner & "CD068 Where BillHeadFmt='" & cboBillHeadFmt.Text & "'") & ""
        strCitemCode = rsCD019.GetString(, , , ",")
        strCitemCode = Left(strCitemCode, Len(strCitemCode) - 1)
        gimCitemCode.Filter = gimCitemCode.Filter
        gimCitemCode.SetQueryCode strCitemCode
        gimCitemCode.IsReadOnly = True
        'gimCitemCode.ShowMulti
End Sub

Private Sub CmbAutoType_Click()
    

End Sub

Private Sub CmbApplyType_Click()
On Error Resume Next
    gdaPropDate1.Text = "": gdaPropDate2.Text = ""
    gdaStopDate1.Text = "": gdaStopDate2.Text = ""
    Select Case CmbApplyType.ListIndex
    Case 0
        gdaStopDate1.Enabled = False
        gdaStopDate2.Enabled = False
        'lblstoptit.Enabled = False
        lblStop.Enabled = False
        gdaPropDate1.Enabled = True
        gdaPropDate2.Enabled = True
        'lblPropDate.ForeColor = vbRed
        'lblStopDate.ForeColor = vbBlack
        lbluntil.Enabled = True
'        lbluntiltit.Enabled = True
    Case 1
        gdaStopDate1.Enabled = True
        gdaStopDate2.Enabled = True
'        lblstoptit.Enabled = True
        lblStop.Enabled = True
        gdaPropDate1.Enabled = False
        gdaPropDate2.Enabled = False
        lbluntil.Enabled = False
        'lblPropDate.ForeColor = vbBlack
        'lblStopDate.ForeColor = vbRed
'        lbluntiltit.Enabled = False
    
    End Select
End Sub

Private Sub cmdCancel_Click()
  On Error GoTo ChkErr
    Unload Me
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "cmdCancel_Click")
End Sub

Public Sub CloseFS()
    On Error Resume Next
    DataPath.Close
    ErrPath.Close
    Set FSO = Nothing
End Sub

Private Function IsDataOk(intCmd As Integer) As Boolean
On Error GoTo ChkErr
    Dim strErrMsg As String
    Dim S As String
    Dim Z As String
    IsDataOk = False
    If Not MustExist(gilCompCode, 2, "公司別") Then Exit Function
    Select Case intCmd
        Case 0
            If CmbApplyType = "" Then strErrMsg = "申請代號": GoTo Warning
            If Len(gilBank.GetCodeNo & "") = 0 Then strErrMsg = "承辦銀行": gilBank.SetFocus:  GoTo Warning
            If gdaSendDate.GetValue = "" Then strErrMsg = "送件日期": gdaSendDate.SetFocus: GoTo Warning
            Select Case CmbApplyType.ListIndex
                Case 0
                    If gdaPropDate1.GetValue = "" Then strErrMsg = "申請日期": gdaPropDate1.SetFocus: GoTo Warning
                    If gdaPropDate2.GetValue = "" Then strErrMsg = "申請日期": gdaPropDate2.SetFocus: GoTo Warning
                Case 1
                    If gdaStopDate1.GetValue = "" Then strErrMsg = "停用日期": gdaStopDate1.SetFocus: GoTo Warning
                    If gdaStopDate2.GetValue = "" Then strErrMsg = "停用日期": gdaStopDate2.SetFocus: GoTo Warning
            End Select
        Case 1
       
    End Select
    If txtDataPath.Text = "" Then strErrMsg = "資料檔位置": txtDataPath.SetFocus: GoTo Warning
    If txtErrPath.Text = "" Then strErrMsg = "問題檔路徑位置": txtErrPath.SetFocus: GoTo Warning
    S = Left(txtDataPath, InStrRev(txtDataPath, "\"))
    If S = "C:\" Or S = "c:\" Then
    Else
        If Not ChkDirExist(S) Then MsgBox "路徑 " & S & " 不存在!", vbExclamation: Exit Function
    End If
    Z = Left(txtErrPath, InStrRev(txtErrPath, "\"))
    If Z = "C:\" Or Z = "c:\" Then
    Else
        If Not ChkDirExist(Z) Then MsgBox "路徑 " & Z & " 不存在!", vbExclamation: Exit Function
    End If
    IsDataOk = True
    Exit Function
Warning:
    Call MsgMustBe(strErrMsg)
    Exit Function
ChkErr:
    Call ErrSub(Me.Name, "IsDataOk")
End Function

Private Sub cmdDataPath_Click()
        On Error GoTo ChkErr
        With cnd1
                .FileName = txtDataPath.Text
                .Filter = "文字檔 (*.txt;*.dat)|*.txt;*.dat|所有檔案 (*.*)|*.*"
                .InitDir = ReadGICMIS1("ErrLogPath")
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
                .FileName = txtDataPath.Text
                .Filter = "文字檔 (*.txt;*.dat)|*.txt;*.dat|所有檔案 (*.*)|*.*"
                .InitDir = ReadGICMIS1("ErrLogPath")
                .Action = 1
                txtErrPath = .FileName
        End With
    Exit Sub
    
ChkErr:
    Call ErrSub(Me.Name, "cmdErrPath_Click")
End Sub
Private Function OpenData(intCmd As Integer) As Boolean
    On Error GoTo ChkErr
    If intCmd = 0 Then
        If OpenFile(DataPath, txtDataPath, True) = False Then Exit Function
    Else
        If OpenFile(DataPath, txtDataPath, True, giOpenTEXT) = False Then Exit Function
    End If
        If OpenFile(ErrPath, txtErrPath, True) = False Then Exit Function
        OpenData = True
    Exit Function
ChkErr:
    Call ErrSub(Me.Name, "OpenData")
End Function

Private Sub cmdOk_Click(intCmd As Integer)
    On Error GoTo ChkErr
    Dim rsCD068 As New ADODB.Recordset
    Dim i As Long
    Dim aryACHTNO() As String
    strInACHDesc = Empty
    strInACHTNO = Empty
    strInBillHeadFmt = Empty
    intSeq = 0
    lngErrCount = 0
    If Not IsDataOk(intCmd) Then Exit Sub
   
'    If Not OpenData(intCmd) Then Exit Sub
    Screen.MousePointer = vbHourglass
    '*************************************************************************************************
    '#3728 ACH改由多筆提出與提回 By Kin 2008/03/10
    If gimBillHeadFmt.GetDescStr = Empty Or gimBillHeadFmt.GetDescStr = "(全選)" Then
        If txtACHTNo.Text <> Empty Then
            aryACHTNO = Split(txtACHTNo.Text, ",")
            For i = LBound(aryACHTNO) To UBound(aryACHTNO)
                strInACHTNO = strInACHTNO & "'" & aryACHTNO(i) & "',"
            Next i
            
        '#3946 CD068要過濾 ACHTNO Is Not Null的條件 By Kin 2008/05/30
        Else
            '#4094 過濾停用要取消 By Kin 2008/09/17
            '#4106 增加判斷ACHTYPE參數 BY Kin 2008/09/22
            strInACHTNO = gcnGi.Execute("Select ''''||ACHTNO||'''' From " & GetOwner & "CD068 Where ACHTNO is Not Null And ACHTDesc Is Not Null And ACHType=2").GetString(, , , ",")
            'strInACHTNO = gcnGi.Execute("Select ''''||ACHTNO||'''' From " & GetOwner & "CD068 Where StopFlag<>1 And ACHTNO is Not Null And ACHTDesc Is Not Null").GetString(, , , ",")
        End If
        '#4094 過濾停用要取消 By Kin 2008/09/17
        'strInACHDesc = gcnGi.Execute("Select ''''||ACHTDESC||'''' From " & GetOwner & "CD068 Where StopFlag<>1 And ACHTNO is Not Null").GetString(, , , ",")
        'strInBillHeadFmt = gcnGi.Execute("Select ''''||BillHeadFmt||'''' From " & GetOwner & "CD068 Where StopFlag<>1 And ACHTNO is Not Null And ACHTDesc is Not Null").GetString(, , , ",")
        '#4106 增加判斷ACHTYPE參數 BY Kin 2008/09/22
        strInACHDesc = gcnGi.Execute("Select ''''||ACHTDESC||'''' From " & GetOwner & "CD068 Where  ACHTNO is Not Null And ACHType=2").GetString(, , , ",")
        strInBillHeadFmt = gcnGi.Execute("Select ''''||BillHeadFmt||'''' From " & GetOwner & "CD068 Where ACHTNO is Not Null And ACHTDesc is Not Null And ACHType=2").GetString(, , , ",")
    Else
        If txtACHTNo.Text <> Empty Then
            aryACHTNO = Split(txtACHTNo.Text, ",")
            For i = LBound(aryACHTNO) To UBound(aryACHTNO)
                strInACHTNO = strInACHTNO & "'" & aryACHTNO(i) & "',"
            Next i
        Else
            '#4094 過濾停用要取消 By Kin 2008/09/17
            'strInACHTNO = gcnGi.Execute("Select ''''||ACHTNO||'''' From " & GetOwner & "CD068 Where StopFlag<>1 And BillHeadFmt In(" & gimBillHeadFmt.GetQueryCode & ")").GetString(, , , ",")
            '#4106 增加判斷ACHTYPE參數 BY Kin 2008/09/22
            strInACHTNO = gcnGi.Execute("Select ''''||ACHTNO||'''' From " & GetOwner & "CD068 Where  BillHeadFmt In(" & gimBillHeadFmt.GetQueryCode & ") And ACHType=2").GetString(, , , ",")
        End If
        '#4094 過濾停用要取消 By Kin 2008/09/17
'        strInACHDesc = gcnGi.Execute("Select ''''||ACHTDESC||'''' From " & GetOwner & "CD068 Where StopFlag<>1 And BillHeadFmt In(" & gimBillHeadFmt.GetQueryCode & ")").GetString(, , , ",")
'        strInBillHeadFmt = gcnGi.Execute("Select ''''||BillHeadFmt||'''' From " & GetOwner & "CD068 Where StopFlag<>1 And BillHeadFmt In(" & gimBillHeadFmt.GetQueryCode & ")").GetString(, , , ",")
        '#4094 原本為BillHeadFmt 現在改為ACHTDESC By Kin 2008/09/17
        '#4106 增加判斷ACHTYPE參數 BY Kin 2008/09/22
        strInACHDesc = gcnGi.Execute("Select ''''||ACHTDESC||'''' From " & GetOwner & "CD068 Where  ACHTDESC In(" & gimBillHeadFmt.GetQueryCode & ") And ACHType=2").GetString(, , , ",")
        strInBillHeadFmt = gcnGi.Execute("Select ''''||BillHeadFmt||'''' From " & GetOwner & "CD068 Where  ACHTDESC In(" & gimBillHeadFmt.GetQueryCode & ") And ACHType=2").GetString(, , , ",")
    End If
    strInACHTNO = Left(strInACHTNO, Len(strInACHTNO) - 1)
    strInACHDesc = Left(strInACHDesc, Len(strInACHDesc) - 1)
    strInBillHeadFmt = Left(strInBillHeadFmt, Len(strInBillHeadFmt) - 1)
    '*************************************************************************************************
    Call ScrToRcd
    strACHTNO = Empty
    strACHDesc = Empty
    strBillHeadFmt = Empty
    Call subChoose          '串Where 指令
    If Not BeginTran(intCmd) Then
            CloseFS
            Screen.MousePointer = vbDefault
            cmdOK(0).Enabled = True: cmdOK(1).Enabled = True: cmdCancel.Enabled = True
            Screen.MousePointer = vbDefault
'            Exit Sub
    End If
    Screen.MousePointer = vbDefault
    On Error Resume Next
    Call CloseRecordset(rsCD068)
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "cmdOk_Click")
End Sub

Private Function AlterSO106(cn As ADODB.Connection, Optional rs As ADODB.Recordset, _
                            Optional strState As String, Optional strData As String, _
                            Optional strErrDes As String, Optional str106MasterId As String = Empty) As Boolean
        On Error GoTo ChkErr
    Dim strSQL As String
    Dim lngAffected As Long
    Dim strCitemStr As String
    Dim rsTmp As New ADODB.Recordset
    Dim rsNote As New ADODB.Recordset
    Dim strRowId As String
    Dim strReStatus As String
    Dim rsNew As New ADODB.Recordset
    '94/08/10 Jacky改 Jim 提1715 問題
    strReStatus = GetReStatus(strData)
    strCitemStr = Replace(gimCitemCode.GetQryStr, ",", "','")
    strCitemStr = Replace(strCitemStr, "(", "('")
    strCitemStr = Replace(strCitemStr, ")", "')")
    strWhere = strWhere ' & " And CitemStr " & strCitemStr
    strUpdTime = Empty
    strNowTime = Empty
    AlterSO106 = False
    Select Case UCase(strState)
        Case "1" '提出
            If rs("ACHCUSTID") & "" <> "" Then
                strSQL = "Update " & GetOwner & "SO106 A Set " & _
                        "SendDate = To_Date('" & gdaSendDate.GetValue & "','YYYYMMDD')," & _
                        "NewUpdTime = sysdate,UpdEn='" & strUpdName & "',UpdTime='" & GetDTString(Now) & "' Where" & _
                        " AccountId='" & rs("AccountId") & "'" & _
                        IIf(str106MasterId <> Empty, " And A.MasterId=" & str106MasterId, "") & _
                        " And CustId=" & rs("Custid") & " And " & strWhere

             Else
                strSQL = "Update " & GetOwner & "SO106 A Set ACHCustid='" & strAchCid & _
                         "' , NewUpdTime = sysdate,SendDate = To_Date('" & gdaSendDate.GetValue & "','YYYYMMDD')," & _
                         "UpdEn='" & strUpdName & "',UpdTime='" & GetDTString(Now) & "' Where" & _
                         " AccountId='" & rs("AccountId") & "'" & _
                         IIf(str106MasterId <> Empty, " And A.MasterId=" & str106MasterId, "") & _
                         " And CustId=" & rs("Custid") & " And " & strWhere
            End If
        
        Case "E" '提回錯誤
            If blnStopAll Then
                strSQL = "Update " & GetOwner & "SO106 A Set SendDate='',ACHCustId=ACHCustId,UpdEn='" & strUpdName & "',UpdTime='" & GetDTString(Now) & "'," & _
                             "StopFlag=1,NewUpdTime = sysdate,StopDate=To_Date('" & Format(Now, "YYYYMMDD") & "','YYYYMMDD'),ReAuthorizeStatus='" & strErrDes & "' Where" & _
                             " LPAD(AccountId,14,'0')='" & GetString(Mid(strData, 28, 14), 14, giRight, True) & "'" & " And ACHCustId='" & GetACHCustID(strData) & "' And SnactionDate is Null And nvl(StopFlag,0) = 0 "
            End If
        Case "2", "3" '郵局通知
            If blnStopAll Then
                Select Case Val(Mid(strData, 26, 1) & "")
                    Case 2
                        strReStatus = "已提回終止扣款!"
                    Case 3
                        strReStatus = "郵局終止扣款!"
                End Select
                
                If Not GetRS(rsNote, "Select Note From SO106 Where Masterid = " & rs("Masterid"), gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
                rsNote("Note") = rsNote("Note") & IIf(Len(rsNote("Note") & "") = 0, "", vbCrLf) & "日期:" & Format(Now, "yyyy/mm/dd") & " " & strReStatus
                rsNote.Update
                
                strSQL = "Update " & GetOwner & "SO106 A Set AuthorizeStatus=2, " & _
                                " SendDate='',ACHCustId=ACHCustId,UpdEn='" & strUpdName & "',UpdTime='" & GetDTString(Now) & "'," & _
                             " StopFlag=1,NewUpdTime = sysdate,StopDate=To_Date('" & Format(Now, "YYYYMMDD") & "','YYYYMMDD'), " & _
                             " ReAuthorizeStatus='" & strReStatus & "' Where" & _
                             " LPAD(AccountId,14,'0')='" & GetString(Mid(strData, 28, 14), 14, giRight, True) & "'" & _
                             " And ACHCustId='" & GetACHCustID(strData) & "' " & _
                             " And MasterId = " & rs("MasterId")
            End If
        Case "4"
            strReStatus = "郵局誤終止扣款，已回復為申請!"
            If Not GetRS(rsNote, "Select Note From SO106 Where Masterid = " & rs("Masterid"), gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
            rsNote("Note") = rsNote("Note") & IIf(Len(rsNote("Note") & "") = 0, "", vbCrLf) & "日期:" & Format(Now, "yyyy/mm/dd") & " " & strReStatus
            rsNote.Update
        
            strSQL = "Update " & GetOwner & "SO106 A Set AuthorizeStatus=1" & _
                  ",SnactionDate=To_Date('" & strDate & "','YYYYMMDD'),ReAuthorizeStatus='" & _
                  strReStatus & "',UpdEn='" & strUpdName & "',NewUpdTime = sysdate,UpdTime='" & GetDTString(RightNow) & "'," & _
                  " SendDate=To_Date('" & getApplyDate(strData) & "','YYYYMMDD'), " & _
                  "StopFlag = 0,StopDate = null " & _
                  " Where" & _
                  " ACHCustId='" & GetACHCustID(strData) & "'" & _
                  " And LPAD(AccountID,14,'0')='" & GetString(Mid(strData, 28, 14), 14, giRight, True) & "'" & _
                  " And SnactionDate is Null  " & _
                  " And StopFlag = 1 And Masterid =" & rs("Masterid")
        Case "RETURNOK" '提回回覆OK
            strSQL = "Update " & GetOwner & "SO106 A Set AuthorizeStatus=1" & _
                  ",SnactionDate=To_Date('" & strDate & "','YYYYMMDD'),ReAuthorizeStatus='" & _
                  strReStatus & "',UpdEn='" & strUpdName & "', NewUpdTime = sysdate,UpdTime='" & GetDTString(RightNow) & _
              "' Where" & _
                  " ACHCustId='" & GetACHCustID(strData) & "'" & _
                  " And SendDate=To_Date('" & getApplyDate(strData) & "','YYYYMMDD')" & _
                  " And LPAD(AccountID,14,'0')='" & GetString(Mid(strData, 28, 14), 14, giRight, True) & "'" & _
                  " And nvl(StopFlag,0) = 0 And SnactionDate is Null And " & strINCD008Where
    End Select
    Select Case UCase(strState)
        Case "E", "2", "3"
            If blnStopAll Then
                If Not ExecuteCommand(strSQL, cn, lngAffected) Then GoTo Finall
            Else
                AlterSO106 = True
                GoTo Finall
            End If
        Case Else
            If Not ExecuteCommand(strSQL, cn, lngAffected) Then GoTo Finall
    End Select

    If lngAffected <> 0 Then AlterSO106 = True
        
    
   
Finall:
    On Error Resume Next
    Call CloseRecordset(rsNew)
    Call CloseRecordset(rsTmp)
    Call CloseRecordset(rsNote)
    Exit Function
ChkErr:
    AlterSO106 = False
End Function

Private Function GetReStatus(strData) As String
    On Error Resume Next
        Select Case Trim(Mid(strData, 72, 2))
            Case "P"
                GetReStatus = "P:已發送授權書及授權扣款檔"
            Case "R"
                GetReStatus = "R:先收到回覆訊息但未收到授權書"
            Case "Y"
                GetReStatus = "Y:先收到回覆訊息後收到授權書"
            Case "M"
                GetReStatus = "M:先收到授權書但未收到回覆訊息"
            Case "S"
                GetReStatus = "S:先收到授權書後收到回覆訊息"
            Case "C"
                GetReStatus = "C:已收到舊件轉檔回覆訊息"
            Case "D"
                GetReStatus = "D:已收到取消授權扣款回覆訊息"
            Case ""
                GetReStatus = "系統日期郵局核印成功!"
        End Select
End Function

Private Function InsSO002A(cn As ADODB.Connection, Optional rs As ADODB.Recordset, Optional strData As String, Optional strCustId As String) As Boolean
    On Error GoTo ChkErr
    Dim strSQL As String
    Dim lngAffected As Long
    Dim bytCnt As Integer
    
    strWhere = strWhere & " And CitemStr " & gimCitemCode.GetQryStr
    
    '*********************************************************************************************************
    '#3676 ACH提回,帳號改由取文字檔完整的14碼在與SO106.ACCOUNTID左補0比較 By Kin 2007/12/12
    bytCnt = Val(RPxx(gcnGi.Execute("select count(*) from " & GetOwner & "so002a where " & _
                     "LPAD(AccountNo,14,'0')='" & GetString(Mid(strData, 28, 14), 14, giRight, True) & "'" & _
                     " And CustId=" & strCustId & _
                     " And CompCode=" & gilCompCode.GetCodeNo).GetString))
                       
    strSQL = "Select DISTINCT A.CUSTID,A.BANKCODE,A.BANKNAME,A.AccountID," & _
             "A.AccountName,B.ChargeAddrNo,B.ChargeAddress,B.MailAddrNo," & _
             "B.MailAddress,C.InvNo,C.InvTitle,C.InvAddress,C.InvoiceType,A.InvSeqNo " & _
             " From " & GetOwner & "SO106 A," & GetOwner & "So001 B," & GetOwner & "So002 C" & _
             " Where A.CustId=B.Custid And A.Custid=C.Custid And A.Custid=" & strCustId & _
             " And LPAD(A.AccountID,14,'0')='" & GetString(Mid(strData, 28, 14), 14, giRight, True) & "'" & _
             " And A.Compcode=" & gilCompCode.GetCodeNo
    '***********************************************************************************************************
    
    If Not GetRS(rsSO106, strSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic, , , True) Then Exit Function
    If Not rsSO106.EOF Then
       If bytCnt = 0 Then
         strSQL = "INSERT INTO " & GetOwner & "SO002A " & _
                  "(CUSTID,COMPCODE,BANKCODE,BANKNAME,ID,ACCOUNTNO," & _
                  "CHARGEADDRNO,CHARGEADDRESS," & _
                  "MAILADDRNO,MAILADDRESS,CHARGETITLE," & _
                  "INVNO,INVTITLE,INVADDRESS,INVOICETYPE)" & _
                  " VALUES (" & _
                  strCustId & "," & _
                  gilCompCode.GetCodeNo & "," & _
                  rsSO106("BankCode") & ",'" & _
                  rsSO106("BankName") & "'," & _
                  0 & ",'" & rsSO106("AccountID") & "','" & _
                  rsSO106("ChargeAddrNo") & "','" & _
                  rsSO106("ChargeAddress") & "'," & _
                  rsSO106("MailAddrNo") & ",'" & _
                  rsSO106("MailAddress") & "','" & _
                  rsSO106("AccountName") & "','" & _
                  rsSO106("InvNo") & "','" & _
                  rsSO106("InvTitle") & "','" & _
                  rsSO106("InvAddress") & "'," & _
                  rsSO106("InvoiceType") & ")"
          
        
       ElseIf bytCnt = 1 Then
         strSQL = "UPDATE " & GetOwner & "SO002A SET " & _
                 "BANKCODE=" & rsSO106("BankCode") & _
                  ",BANKNAME='" & rsSO106("BankName") & _
                  "',ID=0,ACCOUNTNO='" & rsSO106("AccountID") & _
                  "',CHARGEADDRNO=" & rsSO106("ChargeAddrNo") & _
                  ",CHARGEADDRESS='" & rsSO106("ChargeAddress") & _
                  "',MAILADDRNO=" & rsSO106("MailAddrNo") & _
                  ",MAILADDRESS='" & rsSO106("MailAddress") & _
                  "',STOPFLAG=0,STOPDATE=NULL " & _
                  " WHERE CUSTID=" & strCustId & _
                  " AND LPAD(AccountNo,14,'0')='" & GetString(Mid(strData, 28, 14), 14, giRight, True) & _
                  "' AND COMPCODE=" & gilCompCode.GetCodeNo
        End If
        If Not ExecuteCommand(strSQL, cn, lngAffected) Then Exit Function
        '#3728發現的問題,新增SO002A時,如果SO002AD沒有資料也要新增一筆 By Kin 2008/03/14
        bytCnt = gcnGi.Execute("Select Count(*) From " & GetOwner & "SO002AD" & _
                                " Where CustId=" & strCustId & _
                                " And AccountNo='" & rsSO106("AccountID") & "'" & _
                                " And COMPCODE=" & gilCompCode.GetCodeNo & _
                                " AND INVSEQNO=" & rsSO106("INVSEQNO"))(0)
        If bytCnt = 0 Then
            strSQL = "Insert Into " & GetOwner & "SO002AD " & _
                     "(AccountNo,CompCode,CustId,InvSeqNo)" & _
                     " Values(" & _
                     "'" & rsSO106("AccountId") & "'," & _
                     gilCompCode.GetCodeNo & "," & _
                     strCustId & "," & rsSO106("InvSeqNo") & ")"
            gcnGi.Execute strSQL
        End If
                                
                                
    End If
'    rsSO106.Close
    
'    End If
    InsSO002A = True
    Exit Function
ChkErr:
    InsSO002A = False
End Function

Private Function AlterSO003(cn As ADODB.Connection, Optional rs As ADODB.Recordset, Optional strData As String, Optional strCustId As String) As Boolean
    On Error GoTo ChkErr
    Dim strSQL As String
    Dim lngAffected As Long
    Dim rsSO106 As New ADODB.Recordset
    Dim strCitem As String, i As Integer
    Dim aryCitem() As String
    
    '*****************************************************************************************************************
'    strSQL = "SELECT Citemstr,AccountID,BankCode,BankName,PTCode,PTName,CMCode,CMName FROM " & GetOwner & _
        "SO106 WHERE CUSTID=" & GetCustID(strData) & " AND AccountID='" & GetAccID(strData) & "'"
    '94/11/18 Jacky 改
    '#3481 取SO106收費項目時,多增加ACHCustId條件,以防止更新錯誤 By Kin 2007/08/27
    '#3481 Penny要多加條件Nvl(StopFlag,0)=0 And SendDate By Kin 2007/08/28
    '#3676 ACH提回,帳號改由取文字檔完整的14碼在與SO106.ACCOUNTID左補0比較 By Kin 2007/12/12
    strSQL = "SELECT Citemstr,AccountID,BankCode,BankName,PTCode,PTName,CMCode,CMName " & _
            "FROM " & GetOwner & "SO106 WHERE CUSTID=" & strCustId & _
            " AND LPAD(AccountId,14,'0')='" & GetString(Mid(strData, 27, 14), 14, giRight, True) & "'" & _
            " And ACHCustId='" & GetACHCustID(strData) & "'" & _
            " AND NVL(StopFlag,0)=0 AND SendDate=To_Date('" & GetTxtDate(strData) & "','YYYYMMDD')"
    '*****************************************************************************************************************
    
    If txtACHTNo <> "" Then strSQL = strSQL & " And instr(ACHTNo,chr(39)||'" & txtACHTNo & "'||chr(39))>0"
    
    If Not GetRS(rsSO106, strSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic, , , True) Then Exit Function
    If rsSO106.EOF Then Exit Function
    If rsSO106("Citemstr") & "" <> "" Then
        aryCitem = Split(rsSO106("Citemstr") & "", ",")
        For i = 0 To UBound(aryCitem)
            '94/11/17 修改Bug?? SeqNo => CitemCode
            '#3481 PENNY提說要多串CeaseDate By Kin 2007/08/27
            '#3481 Penny說CeaseDate條件又要拿掉了 2007/08/28
            strSQL = "Update " & GetOwner & "SO003 SET BANKCODE='" & _
                     rsSO106("BankCode") & _
                     "',BANKNAME='" & rsSO106("BankName") & _
                     "',ACCOUNTNO='" & rsSO106("AccountID") & _
                     "',PTCode='" & rsSO106("PTCode") & _
                     "',PTName='" & rsSO106("PTName") & _
                     "',CMCode='" & rsSO106("CMCode") & _
                     "',CMName='" & rsSO106("CMName") & _
                     "' WHERE CUSTID=" & strCustId & _
                     " AND SeqNO =" & aryCitem(i)
            If Not ExecuteCommand(strSQL, cn, lngAffected) Then Exit Function
        Next
    End If
    AlterSO003 = True
    Exit Function
ChkErr:
    Call ErrSub(Me.Name, "AlterSO003")
End Function
Private Function AlterSO0032(cn As ADODB.Connection, rsSO106A As ADODB.Recordset, Optional rs As ADODB.Recordset, Optional strData As String, Optional strCustId As String) As Boolean
    On Error GoTo ChkErr
    Dim strSQL As String
    Dim lngAffected As Long
    Dim rsSO106 As New ADODB.Recordset
    Dim strCitem As String, i As Integer
    Dim aryCitem() As String
    strSQL = "SELECT Citemstr,AccountID,BankCode,BankName,PTCode,PTName,CMCode,CMName,InvSeqNo " & _
            "FROM " & GetOwner & "SO106 WHERE CUSTID=" & strCustId & _
            " AND LPAD(AccountId,14,'0')='" & GetString(Mid(strData, 28, 14), 14, giRight, True) & "'" & _
            " And ACHCustId='" & GetACHCustID(strData) & "'" & _
            " AND NVL(StopFlag,0)=0 AND SendDate=To_Date('" & getApplyDate(strData) & "','YYYYMMDD')"
    '*****************************************************************************************************************
    
    strSQL = strSQL & " And instr(ACHTNo,chr(39)||'" & rsSO106A("ACHTNO") & "'||chr(39))>0"
    
    If Not GetRS(rsSO106, strSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic, , , True) Then Exit Function
    If rsSO106.EOF Then Exit Function
    strSQL = "Update " & GetOwner & "SO003 Set BankCode='" & rsSO106("BankCode") & "'," & _
            "BANKNAME='" & rsSO106("BANKNAME") & "'," & _
            "ACCOUNTNO='" & rsSO106("ACCOUNTID") & "'," & _
            "PTCode='" & rsSO106("PTCode") & "'," & _
            "PTName='" & rsSO106("PTName") & "'," & _
            "CMCode='" & rsSO106("CMCode") & "'," & _
            "CMName='" & rsSO106("CMName") & "'," & _
            "InvSeqNo=" & rsSO106("InvSeqNo") & _
            " WHERE CUSTID=" & strCustId & _
            " And CitemCode In(" & rsSO106A("CitemCodeStr") & ")"
     If Not ExecuteCommand(strSQL, cn, lngAffected) Then Exit Function
    AlterSO0032 = True
    Exit Function
ChkErr:
    Call ErrSub(Me.Name, "AlterSO003")
End Function
Private Function BeginTran(intCmd As Integer) As Boolean
    On Error GoTo ChkErr
'    Dim rsTmp As New ADODB.Recordset
    Dim strBankId As String, strSQL As String
    Dim strOldBillNo As String
    Dim lngCountCustId As Long
    Dim lngTime As Long
    Dim strData As String
    Dim blnBigError As Boolean, blnTotalLog As Boolean
    Dim rsUpd As New ADODB.Recordset
    Dim rsCustid As New ADODB.Recordset
    Dim strErrDescription As String
    Dim strCustId As String
    Dim rsSO106A As New ADODB.Recordset
    Dim rsClone As New ADODB.Recordset
    Dim strQrySO106A As String
    Dim strSO106RowId As String
    Dim strSO106ARowId As String
    Dim blnUpdTypeO As Boolean
    Dim strUpdTypeO As String
    
    lngTime = Timer
    blnUpdTypeO = False
    strUpdTypeO = " And SnactionDate Is not null and Propdate is not null" & _
                  " And OldACH=0 And nvl(StopFlag,0) = 0"
    
    If intCmd = 0 Then
'        授權扣款提出
        Screen.MousePointer = vbHourglass
        cmdOK(0).Enabled = False: cmdOK(1).Enabled = False: cmdCancel.Enabled = False
        If Not GetRsTmp(rsTmp) Then Exit Function

        If rsTmp.RecordCount = 0 Then
            MsgBox "無資料！請重新操作！", vbInformation, "訊息！"
            Exit Function
            'GoTo KillFile
        Else
            If Not OpenData(intCmd) Then Exit Function
            
            'If Not InsertHead(rsTmp) Then GoTo KillFile
            If Not InsertDetail(rsTmp) Then Exit Function
            If Not InsertFinal(rsTmp) Then Exit Function
            MsgBox "需處理筆數共" & intSeq + lngErrCount & "筆," & vbCrLf & vbCrLf & _
            "已完成資料筆數共" & intSeq & "筆," & vbCrLf & vbCrLf & _
            "問題筆數共" & lngErrCount & "筆," & vbCrLf & vbCrLf & _
            "共花費:" & (Timer - lngTime) \ 1 & "秒", vbInformation, "訊息"
'            msgResult rsTmp.RecordCount, lngErrCount, lngTime       '顯示執行結果
        End If
        Screen.MousePointer = vbDefault
        cmdOK(0).Enabled = True: cmdOK(1).Enabled = True: cmdCancel.Enabled = True
        Call CloseFS
        BeginTran = True
        On Error Resume Next
        rsTmp.Close
        Set rsTmp = Nothing
    Else
    '   授權扣款提回
        Call FileReturn
  End If
  Exit Function
ChkErr:
    Call ErrSub(Me.Name, "BeginTran")
End Function
Private Function UpdateOType(ByVal aRowid As String) As Boolean
  On Error GoTo ChkErr
    Dim aRet As Boolean
    Dim aSQL As String
    
    If aRowid = Empty Then
        aRet = True
    Else
        aSQL = "Update " & GetOwner & "SO106 A Set StopFlag=1," & _
                     "Stopdate=A.SendDate,OldACH =1 Where RowId In(" & aRowid & ")"
        gcnGi.Execute aSQL
        aRet = True
    End If
    UpdateOType = aRet
    Exit Function
ChkErr:
  Call ErrSub(Me.Name, "UpdateOType")
End Function
Private Function getApplyErrMsg(ByVal strReadLine As String) As String
    Dim retResult As String
    Select Case UCase(Mid(strReadLine, 72, 2))
        Case "03"
            retResult = "已終止代繳"
        Case "06"
            retResult = "凍結戶或警示戶"
        Case "07"
            retResult = "業務支票專戶"
        Case "08"
            retResult = "帳號錯誤"
        Case "09"
            retResult = "終止戶"
        Case "10"
            retResult = "身分證號不符"
        Case "11"
            retResult = "轉出戶"
        Case "12"
            retResult = "拒絕往來戶"
        Case "13"
            retResult = "無此用戶編號"
        Case "14"
            retResult = "用戶編號已存在"
        Case "16"
            retResult = "管制帳戶"
        Case "17"
            retResult = "掛失戶"
        Case "18"
            retResult = "異常交易帳戶"
        Case "19"
            retResult = "用戶編號非英數字元"
         Case "91"
            retResult = "規定期限內未有扣款"
        Case "98"
            retResult = "其他"
    End Select
    If Len(retResult) = 0 Then
        Select Case UCase(Mid(strReadLine, 74, 1))
            Case "1"
                retResult = "局帳號不符"
            Case "2"
                retResult = "戶名不符"
            Case "3"
                retResult = "身分證號不符"
            Case "4"
                retResult = "印鑑不符"
            Case "9"
                retResult = "印鑑不符"
        End Select
    End If
    If Len(retResult) = 0 Then
        retResult = "找不到錯誤代碼"
    End If
    getApplyErrMsg = retResult
End Function


Private Function AuthInOk(strReadLine As String, _
        Optional BigError As Boolean = False, Optional TotalLog As Boolean = True, _
        Optional strBankName As String, Optional strCustId As String, _
        Optional rs As ADODB.Recordset = Nothing) As Boolean
    On Error GoTo ChkErr
    Dim rsTmp As New ADODB.Recordset
    Dim strBillNo As String
    Dim strAccID As String
    Dim strUpdSQL As String
    strAccID = GetAccID(strReadLine)
    strErrAuth = Empty
    Select Case UCase(Mid(strReadLine, 72, 2))
        Case "03"
            strErrAuth = "已終止代繳"
        Case "06"
            strErrAuth = "凍結戶或警示戶"
        Case "07"
            strErrAuth = "業務支票專戶"
        Case "08"
            strErrAuth = "帳號錯誤"
        Case "09"
            strErrAuth = "終止戶"
        Case "10"
            strErrAuth = "身分證號不符"
        Case "11"
            strErrAuth = "轉出戶"
        Case "12"
            strErrAuth = "拒絕往來戶"
        Case "13"
            strErrAuth = "無此用戶編號"
        Case "14"
            strErrAuth = "用戶編號已存在"
        Case "16"
            strErrAuth = "管制帳戶"
        Case "17"
            strErrAuth = "管制帳戶"
    End Select
     
     
     Select Case UCase(Mid(strReadLine, 108, 1))
     Case "0"
          strErrAuth = ""
     Case "1" '                1 =印鑑不符
          strErrAuth = " 印鑑不符!! "
     Case "2" '                2 = 無此帳號
          strErrAuth = "無此帳號 !! "
     Case "3" '                3 = 委繳戶統編不存在
          strErrAuth = "委繳戶統編不存在 !! "
     Case "4" '                4 = 資料重覆
          strErrAuth = "資料重覆 !! "
     Case "5" '                5 = 原交易不存在
          strErrAuth = "原交易不存在 !! "
     Case "6" '                6 = 電子資料與授書內容不符
          strErrAuth = "電子資料與授權書內容不符 !! "
     Case "7" '                7 = 帳戶已結清
          strErrAuth = "帳戶已結清 !! "
     Case "8" '                8 = 印鑑不清
          strErrAuth = "印鑑不清 !! "
     Case "A"
          strErrAuth = "未收到授權書 !!"
     Case "B"
          strErrAuth = "用戶號碼錯誤 !!"
     Case "C"
          strErrAuth = "靜止戶 !!"
     Case "D"
          strErrAuth = "未收到聲明書 !!"
     Case Else '                9 = 其他
          strErrAuth = "其他不成功原因 !! "
     End Select
        '銀行之錯誤
     If strErrAuth <> "" Then
        '****************************************************************************************************
        '#3739 授權失敗停用有選擇時，要將錯誤原因填入SO106.NOTE By Kin 2008/01/22
'        If chkStop.Value = 1 Then
'            rs.MoveFirst
'            Do While Not rs.EOF
'                strUpdSQL = "Update SO106 Set NOTE='" & strErrAuth & "' Where RowId='" & rs("RowId") & "'"
'                gcnGi.Execute strUpdSQL
'                rs.MoveNext
'            Loop
'        End If
        '****************************************************************************************************
        Call InsertToErr(strReadLine, strErrAuth, strCustId, strAccID)
        Exit Function
     End If

        AuthInOk = True
    Exit Function
ChkErr:
    ErrSub Me.Name, "AuthInOk"
End Function

Private Function InsertToErr(strReadLine As String, strErrDescription As String, strCustId As String, strAccID As String) As Boolean
    On Error Resume Next
        WriteTextLine ErrPath, " 客編: " & strCustId & "; 帳號:" & strAccID & "; 失敗原因:" & strErrDescription & ""
        lngErrCount = lngErrCount + 1
        InsertToErr = True
End Function


Private Function GetRealDateTran(strRealDate As String) As String
    On Error Resume Next
        GetRealDateTran = Val(Replace(strRealDate, "/", ""))
End Function

Private Function InsertDetail(rs As ADODB.Recordset) As Boolean
  On Error GoTo ChkErr
    Dim strData As Variant
    Dim rsCD018 As New ADODB.Recordset
    Dim rsAchCid As New ADODB.Recordset
    Dim strSQL As String
    Dim strErrlog As String
    Dim strBankId As String
    Dim strSqlCid As String
    Dim GetACHCustID As Integer
    Dim blnNoAchReceptor As Boolean
    Dim rsSO106A As New ADODB.Recordset
    Dim strQrySO106A As String
    Dim strUpdSO106A As String
    intSeq = 0: lngErrCount = 0
    
    blnNoAchReceptor = False
    If rs.RecordCount > 0 Then
        rs.MoveFirst
    End If

    '*********************************************************************************************************************************************************
    '#3253 判斷87-106位元要填空白或跑舊流程 By Kin 2007/12/19
'    If cboBillHeadFmt.Text <> "" Then
'        blnNoAchReceptor = gcnGi.Execute("Select Nvl(NoAchReceptor,0) From " & strOwnerName & "CD068 Where BillHeadFmt='" & strBillHeadFmt & "'")(0) = 1
'    End If
    '*********************************************************************************************************************************************************
    
    Do While Not rs.EOF
        '**********************************************************************************************************
        '#3728 將SO106的資料對應到SO106A,以SO106A資料做提出的動作 By Kin 2008/03/11
        strQrySO106A = Empty
        Select Case CmbApplyType.ListIndex
            Case 0
               
                strQrySO106A = "Select RowID,A.* From " & GetOwner & "SO106A A " & _
                               "Where MasterId=" & rs("MasterId") & _
                               " And ACHTNO In(" & strInACHTNO & ")" & _
                               " And ACHDesc In(" & strInACHDesc & ")" & _
                               " And RecordType=0 And AuthorizeStatus=4 " & _
                               " And StopFlag<>1 And StopDate is Null Order By ACHTNO"
               
            Case 1
                strQrySO106A = "Select ROWID,A.* From " & GetOwner & "SO106A A " & _
                               "Where MasterId=" & rs("MasterId") & _
                               " And ACHTNO In(" & strInACHTNO & ")" & _
                               " And ACHDesc In(" & strInACHDesc & ")" & _
                               " And RecordType=1 And AuthorizeStatus=4 "
                               
          
        End Select
        If Not GetRS(rsSO106A, strQrySO106A, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
        If rsSO106A.RecordCount > 0 Then
            rsSO106A.MoveFirst
        End If
        '**********************************************************************************************************
        
        
        
        blnHasUpd = False
        '#3728 原本讀取SO106,現在改為讀取符合SO106A條件的資料 By Kin 2008/03/11
        Do While Not rsSO106A.EOF
            'blnNoAchReceptor = gcnGi.Execute("Select Nvl(NoAchReceptor,0) From " & strOwnerName & "CD068 Where ACHTNO='" & rsSO106A("ACHTNO") & "' And ACHTDESC='" & rsSO106A("ACHDesc") & "'")(0) = 1
            'If rs("Citemstr") & "" <> "" And ChkInCitem(gimCitemCode.GetQueryCode, rs("Citemstr") & "") Or Asc(CmbAutoType) = 79 Then
            If rsSO106A("CitemCodeStr") & "" <> "" And ChkInSO106ACitem(gimCitemCode.GetQueryCode, rsSO106A("CitemCodeStr") & "", IIf(gimCitemCode.GetQueryCode = Empty, True, False)) Then
                '資料別
                strData = "1"
                '委託機構代號
                'strBankId = GetRsValue("Select BankId From " & strOwnerName & "CD018 Where CodeNo = " & rs("BankCode"), gcnGi) & ""
                strData = strData & GetString(rs("BankID") & "", 3, giLeft, False) & ""
                '保留欄
                strData = strData & Space(4)
                '媒體產生日期
                strData = strData & GetString(GetRealDateTran(gdaSendDate.GetValue), 8, giRight, True)
                '批號
                strData = strData & "001"
                '流水號
                intSeq = intSeq + 1
                strData = strData & GetString(intSeq, 6, giRight, True)
                '申請代號
                strData = strData & (CmbApplyType.ListIndex + 1) & ""
                '帳戶別
                If Len(rs("AccountID") & "") = 8 Then
                    strData = strData & "G"
                Else
                    strData = strData & "P"
                End If
                '儲金帳號
                strData = strData & GetString(rs("AccountID") & "", 14, giRight, True)
                '用戶編號
                If Len(rs("ACHCUSTID") & "") = 0 Then
                    intSeq = intSeq - 1
                    lngErrCount = lngErrCount + 1
                    strErrlog = " 客編 : " & rs("Custid") & " 帳號 : " & rs("AccountID") & " 沒有ACH用戶號碼"
                    WriteLineData strErrlog, ErrPath
                    GoTo GotoLoop
                End If
                strData = strData & GetString(rs("ACHCustId") & "", 20, giRight, False)
                '身分證統一編號或統一證號
                strData = strData & GetString(rs("AccountNameID") & "", 10, giLeft, False)
                '狀況代號
                strData = strData & Space(2)
                '核對註記
                strData = strData & Space(1)
                '保留欄
                strData = strData & Space(26)
                WriteLineData strData, DataPath
                Select Case CmbApplyType.ListIndex
                    Case 0, 2
                        If Not blnHasUpd Then
                            '#3946 改用MasterId為PK By Kin 2008/05/30
                            If AlterSO106(gcnGi, rsTmp, "1", , , rs("MasterId")) = False Then
                                Exit Function
                            Else
                                blnHasUpd = True    '代表已更新過,不用在更新一次SO106
                            End If
                        End If
                End Select
                '更新SO106A 狀態
                If rsSO106A("AuthorizeStatus") & "" = "4" Then
                    rsSO106A("AuthorizeStatus") = Null
                    rsSO106A.Update
                End If
            Else
                '將符合條件但沒有指定扣款週期收費項目之轉帳資料出log以告知USER
                lngErrCount = lngErrCount + 1
                strErrlog = " 客編 : " & rs("Custid") & " 帳號 : " & rs("AccountID") & " 沒有指定扣款週期收費項目"
                WriteLineData strErrlog, ErrPath
            End If
            rsSO106A.MoveNext
        Loop
GotoLoop:
        blnHasUpd = False
        rs.MoveNext
        DoEvents
    Loop
     InsertDetail = True
     blnHasUpd = False
     On Error Resume Next
     Call CloseRecordset(rsSO106A)
     Call CloseRecordset(rsCD018)
     Call CloseRecordset(rsAchCid)
     
  Exit Function
ChkErr:
'    cnn.RollbackTrans
    
    ErrSub Me.Name, "InsertDetail"
End Function

Private Function GetPostUnit(ByVal strACHTNO As String, ByVal strACHTNOstr As String) As String
    On Error Resume Next
        If InStr(1, strACHTNOstr, "'" & strACHTNO & "'") > 0 Then
            GetPostUnit = GetRsValue("Select PostUnit From " & GetOwner & "CD068 Where AchtNo = '" & strACHTNO & "'")
        End If
End Function

Private Function WriteLineData(ByVal vData As String, ByVal objFile As Object)
    On Error GoTo ChkErr
        Call objFile.WriteLine(vData)
    Exit Function
ChkErr:
    Call ErrSub("clsModule", "WriteLineData")
End Function

Private Function ChkInCitem(strCode As String, strCitemStr As String) As Boolean
    On Error GoTo ChkErr
    Dim aryCitem() As String
    Dim i As Integer
    Dim rsCit As New ADODB.Recordset
    ChkInCitem = False
    If strCitemStr <> "" Then
        If Not GetRS(rsCit, "Select CitemCode From " & GetOwner & "SO003 WHERE SeqNo IN(" & strCitemStr & ")") Then Exit Function
        Do While Not rsCit.EOF
             If InStr(1, strCode, rsCit("CitemCode") & "") <> 0 Then
                 ChkInCitem = True
             End If
             rsCit.MoveNext
        Loop
    End If
'    aryCitem = Split(strCitemStr, ",")
'    For i = 0 To UBound(aryCitem)
'        If InStr(1, strCode, aryCitem(i)) <> 0 Then
'            ChkInCitem = True
'        End If
'    Next
    
    Exit Function
ChkErr:
    ErrSub Me.Name, "ChkInCitem"
End Function
'#3728 檢查SO106A的收費項目是否有符合CD068裡的收費項目 By Kin 2008/03/11
Private Function ChkInSO106ACitem(strCode As String, strCitemStr As String, ByVal blnAll As Boolean) As Boolean
    On Error GoTo ChkErr
    Dim i As Integer
    Dim strCD068 As String
    Dim aryCitemStr() As String
    Dim rsCD068Temp As New ADODB.Recordset
    'blnAll=Ture代表全選或沒有選擇多帳戶依據,則要挑選CD068裡的收費項目
    If blnAll Then
        'strCD068 = gcnGi.Execute("Select CitemCodeStr From " & strOwnerName & "CD068 Where BillHeadFmt In (" & strInBillHeadFmt & ") And StopFlag<>1").GetString(, , , ",")
        '#4094 過濾停用要取消 By Kin 2008/09/17
        '#4106 增加判斷ACHTYPE參數 BY Kin 2008/09/22
'        strCD068 = gcnGi.Execute("Select CitemCodeStr From " & strOwnerName & "CD068 Where BillHeadFmt In (" & strInBillHeadFmt & ") And ACHType=1").GetString(, , , ",")
'        If Len(strCD068) > 0 Then
'            strCD068 = Left(strCD068, Len(strCD068) - 1)
'            strCode = strCD068
'        End If
        '#7049 改用CD068A.CitemCode By Kin 2015/07/08
        If Not GetRS(rsCD068Temp, "Select distinct CitemCode From " & strOwnerName & "CD068A " & _
                        " Where Exists (Select * From " & strOwnerName & "CD068 Where CD068.BillHeadFmt = CD068A.BillHeadFmt " & _
                                            " And CD068.ACHType = 2 And BillHeadFmt In (" & strInBillHeadFmt & "))", gcnGi, adUseClient, adOpenKeyset) Then Exit Function
        strCD068 = rsCD068Temp.GetString(, , , ",")
        strCD068 = Left(strCD068, Len(strCD068) - 1)
        strCode = strCD068
    End If
    ChkInSO106ACitem = False
    If strCitemStr <> "" Then
        aryCitemStr = Split(strCitemStr, ",")
        For i = LBound(aryCitemStr) To UBound(aryCitemStr)
            If InStr(1, strCode, aryCitemStr(i) & "") <> 0 Then
                ChkInSO106ACitem = True
            Else
                ChkInSO106ACitem = False
                Exit For
            End If
        Next i
    End If
    On Error Resume Next
    CloseRecordset rsCD068Temp
    Exit Function
ChkErr:
    ErrSub Me.Name, "ChkInSO106ACitem"
End Function
Private Function WriteHead() As String

End Function
Private Function InsertHead(rs As ADODB.Recordset) As Boolean
    On Error GoTo ChkErr
    Dim intReserve As Integer
    
    
    
    Dim strData As Variant
        '區別碼(錄別) , 資料代號, 交易日期, 發送單位代號, 備用
        
        strData = ""
        strData = strData & "BOF" & "ACHP02"
        '首錄別(1-3)資料代號(4-9)
        strData = strData & GetString(GetRealDateTran(gdaSendDate.Text), 8, giRight, True)
        '處理日期(10-17)
        'strData = strData & GetString(txtSendSpcId, 7, giLeft)
        '發送單位代號(18-24)
        strData = strData & GetString("", 96, giLeft)
        '備用(25-120)
        WriteLineData strData, DataPath
        InsertHead = True
    Exit Function
ChkErr:
    ErrSub Me.Name, "InsertHead"

End Function

Private Function GetRsTmp(ByRef rs As ADODB.Recordset) As Boolean
    On Error GoTo ChkErr
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset

       'EMC
    strFrom = " " & GetOwner & "SO106 A "
'    if instr(CmbAutoType.Text,"O")
            
            strWhere = strWhere & IIf(strWhere = "", "", " And ") & strChoose
            
            '#3585 多挑選ACHCUSTID欄位,判斷ACHCUSTID是否有值，如果有值不進行Update By Kin 2007/10/26
            '#3946 多增加MasterId 欄位 By Kin 2008/05/30
            strSQL = "SELECT A.RowID,A.Custid,A.BankCode,AccountID,A.AccountNameId,A.CitemStr,A.ACHTNo,A.ACHSN,A.ACHCUSTID,A.MasterId," & _
                        " (Select BankId From " & strOwnerName & "CD018 Where CD018.CodeNo = A.BankCode And RowNum = 1 ) BankID " & _
                        " From " & _
                        strFrom & " Where " & strWhere & strCancelWhere
         
        If Not GetRS(rs, strSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic, , , True) Then Exit Function
        strReturnSQL = rs.Source
    GetRsTmp = True
    Exit Function
ChkErr:
     ErrSub Me.Name, "GetRsTmp"
End Function

Private Sub Form_Activate()
    On Error Resume Next
        Screen.MousePointer = vbDefault
End Sub
Private Sub DefaultValue()
    On Error Resume Next
'    Dim intPara23 As Integer
        strPrgName = "SO3293A"
        gilCompCode.SetCodeNo garyGi(9)
        gilCompCode.Query_Description
        'gdaSendDate.SetValue (RightNow)
        strUpdEn = garyGi(0)
        strUpdName = garyGi(1)
        strErrPath = ReadGICMIS1("ErrLogPath")
       ' txtInvoiceId = GetRsValue("Select InvoiceId From " & GetOwner & "So041") & ""
       CmbApplyType.ListIndex = 0
        If CmbApplyType.ListIndex = 0 Then
            gdaStopDate1.Enabled = False
            gdaStopDate2.Enabled = False
            'lblPropDate.ForeColor = vbRed
            
            'lblstoptit.Enabled = False
            'lblStop.Enabled = False
        End If
Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "DefaultValue")
End Sub
Private Sub Form_Load()
    On Error GoTo ChkErr
        Call subGim
        Call subGil
        CboAddItem
        Call DefaultValue
        RcdToScr
Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "Form_Load")
End Sub

Private Sub RcdToScr()
    On Error GoTo ChkErr
    Dim LogFile As TextStream
    If Dir(strErrPath & "\" & strPrgName & "Out.log") = "" Then Exit Sub
            On Error Resume Next
            Set LogFile = FSO.OpenTextFile(strErrPath & "\" & strPrgName & "Out.log", ForReading)
            With LogFile
                    'If Not .AtEndOfStream Then txtSendSpcId.Text = .ReadLine & ""
                       '發送單位代號
                                        
                    'If Not .AtEndOfStream Then gdaSendDate.Text = .ReadLine & ""
                       '送出日期
                    'If Not .AtEndOfStream Then txtPutId.Text = .ReadLine & ""
                       '發動者代號
                    'If Not .AtEndOfStream Then txtInvoiceId.Text = .ReadLine & ""
                       '發動者統一編號
                    If Not .AtEndOfStream Then txtDataPath.Text = .ReadLine & ""
                       '資料檔位置
                    If Not .AtEndOfStream Then txtErrPath.Text = .ReadLine & ""
                       '問題檔位置
                    If Not .AtEndOfStream Then chkStop.Value = .ReadLine & ""
                       '是否將授權失敗停用
            End With
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "RcdToScr")
End Sub
Private Sub ScrToRcd()
    On Error GoTo ChkErr
    Dim LogFile As TextStream
        Set LogFile = FSO.CreateTextFile(strErrPath & "\" & strPrgName & "Out.log", True)
        With LogFile
                '.WriteLine (txtSendSpcId.Text)         '發送單位代號
                '.WriteLine (gdaSendDate.Text)          '送出日期
                '.WriteLine (txtPutId.Text)             '發動者代號
                '.WriteLine (txtInvoiceId.Text)         '發動者統一編號
                .WriteLine (txtDataPath.Text)          '資料檔位置
                .WriteLine (txtErrPath.Text)           '問題檔位置
                .WriteLine (chkStop.Value)             '是否將授權失敗停用
        End With
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "ScrToRcd")
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If Shift = 2 Then
        If KeyCode = vbKeyF Then
            KeyCode = 0
            txtSQL.Visible = True
            txtSQL = strReturnSQL
            txtSQL.Move 0, 0, Me.Width, Me.Height / 2
        ElseIf KeyCode = vbKeyX Then
            txtSQL.Visible = False
        End If
    End If
'    If KeyCode = vbKeyF2 Then Call cmdOk_Click: KeyCode = 0
    If KeyCode = vbKeyEscape Then Unload Me: KeyCode = 0
End Sub

Private Sub subGim()
    On Error GoTo ChkErr
        SetgiMulti gimBankId, "CodeNo", "Description", "CD018", "代碼", "名稱", , True
        SetgiMulti gimCitemCode, "CodeNo", "Description", "CD019", "代碼", "名稱", , True
        '#3946 要過濾ACHTNO 為Null 不顯示 By Kin 2008/05/30
        'SetgiMulti gimBillHeadFmt, "BillHeadFmt", "ACHTNO", "CD068", "單據格式", "交易代碼", " Where ACHTNO Is Not Null And ACHTDesc Is not Null"
        '#4016 增加一個ACHType參數 By Kin 2008/09/22
        SetgiMulti gimBillHeadFmt, "ACHTDESC", "ACHTNO", "CD068", "ACH交易說明", "ACH交易代碼", " Where ACHTNO Is Not Null And ACHTDesc Is not Null And ACHType=2"
        gimCitemCode.IsReadOnly = True
Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "subGim")
End Sub

Private Sub subGil()
  On Error GoTo ChkErr
    Call SetgiList(gilCompCode, "CodeNo", "Description", "CD039", , , , , , , GetCompCodeFilter(gilCompCode))
    SetgiList gilBank, "CodeNo", "Description", "CD018", , , , , 3, 24, " Where CD018.PRGNAME like '%POST%' AND STOPFLAG=0 "
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "subGil")
End Sub

Private Sub CboAddItem()
    On Error Resume Next
    With CmbApplyType
        .AddItem "1：申請", 0
        .AddItem "2：終止", 1
    End With
    If Not GetRS(rsBillHeadFmt, "Select BillHeadFmt,CitemCodeStr From " & GetOwner & "CD068 WHERE ACHTNo IS NOT NULL And ACHTDesc is Not Null") Then Exit Sub
'#3728 將此原件改為gimBillHeadFmt By Kin 2008/03/10
'        If Not GetRS(rsBillHeadFmt, "Select BillHeadFmt,CitemCodeStr From " & GetOwner & "CD068 WHERE ACHTNo IS NOT NULL") Then Exit Sub
'        Do While Not rsBillHeadFmt.EOF
'            cboBillHeadFmt.AddItem rsBillHeadFmt("BillHeadFmt") & ""
'            rsBillHeadFmt.MoveNext
'        Loop
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    Call CloseRecordset(rsBillHeadFmt)
    Call CloseRecordset(rsSO106)
    Call CloseRecordset(rsTmp)
End Sub

Private Sub gdaPropDate1_GotFocus()
    On Error GoTo ChkErr
    gdaPropDate1.SetValue (RightNow)
Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "gdaPropDate1_GotFocus")
End Sub

Private Sub gdaPropDate1_Validate(Cancel As Boolean)
    On Error GoTo ChkErr
        If Not IsDate(gdaPropDate1.Text) Then Exit Sub
        If gdaPropDate1.GetValue <> "" Then
            If gdaPropDate2.GetValue = "" Then gdaPropDate2.SetValue GetLastDate(gdaPropDate1.GetValue(True))
        End If
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "gdaPropDate1_Validate")
End Sub

Private Sub gdaPropDate2_GotFocus()
    On Error Resume Next
        If gdaPropDate2.GetValue <> "" Then Exit Sub
        gdaPropDate2.SetValue GetLastDate(Date)
End Sub

Private Sub gdaPropDate2_Validate(Cancel As Boolean)
    On Error Resume Next
        If gdaPropDate1.GetValue = "" Or gdaPropDate2.GetValue = "" Then Exit Sub
        If Not IsDate(gdaPropDate2.Text) Then Exit Sub
        If DateDiff("d", gdaPropDate1.Text, gdaPropDate2.Text) < 0 Then
            MsgBox "收費截止日不得小於收費起始日！", vbExclamation, "訊息！"
            gdaPropDate2.SetValue gdaPropDate1.GetValue
            Cancel = True
        End If
End Sub

Private Sub gdaSendDate_GotFocus()
    On Error GoTo ChkErr
    If gdaSendDate.Text = "" Then gdaSendDate.SetValue (RightNow)
Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "gdaSendDate_GotFocus")
End Sub

Private Sub gdaStopDate1_GotFocus()
    On Error GoTo ChkErr
    gdaStopDate1.SetValue (RightNow)
Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "gdaStopDate1_GotFocus")
End Sub

Private Sub gdaStopDate1_Validate(Cancel As Boolean)
    On Error GoTo ChkErr
        If Not IsDate(gdaStopDate1.Text) Then Exit Sub
        If gdaStopDate1.GetValue <> "" Then
            If gdaStopDate2.GetValue = "" Then gdaStopDate2.SetValue GetLastDate(gdaStopDate1.GetValue(True))
        End If
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "gdaStopDate1_Validate")
End Sub

Private Sub gdaStopDate2_GotFocus()
    On Error Resume Next
        If gdaStopDate2.GetValue <> "" Then Exit Sub
        gdaStopDate2.SetValue GetLastDate(Date)
End Sub

Private Sub gdaStopDate2_Validate(Cancel As Boolean)
    On Error Resume Next
        If gdaStopDate1.GetValue = "" Or gdaStopDate2.GetValue = "" Then Exit Sub
        If Not IsDate(gdaStopDate2.Text) Then Exit Sub
        If DateDiff("d", gdaStopDate1.Text, gdaStopDate2.Text) < 0 Then
            MsgBox "收費截止日不得小於收費起始日！", vbExclamation, "訊息！"
            gdaStopDate2.SetValue gdaStopDate1.GetValue
            Cancel = True
        End If
End Sub

Private Sub gilCompCode_Change()
    On Error Resume Next
        garyGi(16) = gilCompCode.GetOwner
        Call subGim
        Call subGil
        Call DefaultValue
       'If SetACHbankTypeCbo = False Then cmdOK(0).Visible = False: cmdOK(1).Visible = False

End Sub

Private Function SetACHbankTypeCbo() As Boolean

  Dim strSQL As String
  Dim rsTmp As New ADODB.Recordset
  Dim lbn
  Dim lngIndex As Long
  On Error GoTo ChkErr
  
  For lngIndex = 0 To cboACHbankType.ListCount - 1
       cboACHbankType.RemoveItem (0)
  Next
    lngIndex = 0
    strSQL = "SELECT DISTINCT PrgName FROM " & GetOwner & "CD018 WHERE UPPER(PrgName) like '%" & "ACH" & "%'  AND COMPCODE =" & gilCompCode.GetCodeNo & "  AND STOPFLAG = 0 "
    Call GetRS(rsTmp, strSQL, gcnGi, adUseClient, adOpenDynamic)
    Dim blnAdd As Boolean   ''  判斷如果有加入ACH項目清單的話，則將其設為flase ，這樣底下的lngIndex才可以加 1  不然下筆資料的迴圈會出現 索引超出陣列範圍
    blnAdd = False
  
    Do While Not (rsTmp.EOF Or rsTmp.BOF)
        Select Case UCase(rsTmp("PrgName"))
            Case "ACHTRANREFER"
                 cboACHbankType.AddItem ("ACH產出格式 ")
                 cboACHbankType.ItemData(lngIndex) = 0
                 blnAdd = True
          End Select
          If blnAdd = True Then lngIndex = lngIndex + 1
          blnAdd = False
         rsTmp.MoveNext
     Loop
    rsTmp.Close
    Set rsTmp = Nothing
    If lngIndex = 0 Then
         MsgBox "銀行別資料未設，請設定相關銀行別代碼之後，重新執行  !!"
         SetACHbankTypeCbo = False
    Else
         cboACHbankType.ListIndex = 0
          SetACHbankTypeCbo = True
    End If
    Screen.MousePointer = vbDefault
   
Exit Function
ChkErr:
  ErrSub Me.Name, "SetACHbankTypeCbo"
End Function

Private Sub subChoose()
    On Error GoTo ChkErr
    Dim strAddr As String
    Dim strMduChoose As String
    Dim aryACHID() As String
    Dim strWhereACHID As String
    Dim i As Long
    strChoose = ""
    strWhere = ""
    strCancelWhere = ""
    If gdaPropDate1.GetValue <> "" Then Call subAnd("PropDate >= To_date('" & gdaPropDate1.GetValue & "','YYYYMMDD')")
    If gdaPropDate2.GetValue <> "" Then Call subAnd("PropDate < To_date('" & gdaPropDate2.GetValue & "','YYYYMMDD') +1")
    '************************************************************************************************************************
    Select Case CmbApplyType.ListIndex
        Case 0
        Case 1
            If gdaStopDate1.GetValue <> "" Then Call subAnd("A.StopDate >= To_date('" & gdaStopDate1.GetValue & "','YYYYMMDD')")
            If gdaStopDate2.GetValue <> "" Then Call subAnd("A.StopDate < To_date('" & gdaStopDate2.GetValue & "','YYYYMMDD') +1")
    End Select
'
'    '#3728 如果是提出授權取消則停用日期改串在SO106A By Kin 2008/03/11
'    Select Case CmbApplyType.ListIndex
'        Case 0, 2
'            If gdaStopDate1.GetValue <> "" Then Call subAnd("StopDate >= To_date('" & gdaStopDate1.GetValue & "','YYYYMMDD')")
'            If gdaStopDate2.GetValue <> "" Then Call subAnd("StopDate < To_date('" & gdaStopDate2.GetValue & "','YYYYMMDD') +1")
'        Case 1
'            If gdaStopDate1.GetValue <> "" Then strCancelWhere = IIf(strCancelWhere = Empty, "StopDate >= To_date('" & gdaStopDate1.GetValue & "','YYYYMMDD')", strCancelWhere & " And StopDate >= To_date('" & gdaStopDate1.GetValue & "','YYYYMMDD')")
'            If gdaStopDate2.GetValue <> "" Then strCancelWhere = IIf(strCancelWhere = Empty, "StopDate < To_date('" & gdaStopDate2.GetValue & "','YYYYMMDD')+1", strCancelWhere & " And StopDate < To_date('" & gdaStopDate2.GetValue & "','YYYYMMDD')+1")
'    End Select
    '***************************************************************************************************************************
    If Len(gilBank.GetCodeNo & "") > 0 Then
        subAnd ("A.BankCode = " & gilBank.GetCodeNo)
    End If
    
    
'        If gimCitemCode.GetQryStr <> "" Then
'            subAnd ("(ACHTNo is not null And CitemStr " & gimCitemCode.GetQryStr & " or ACHTNo is null) ")
'        End If
        '94/11/18 Jacky 改
        '*************************************************************
        '#3728 ACHTNO改為多選 By Kin 2008/03/11
        aryACHID = Split(txtACHTNo, ",")
        For i = LBound(aryACHID) To UBound(aryACHID)
            If aryACHID(i) <> Empty Then strWhereACHID = IIf(strWhereACHID <> Empty, strWhereACHID & " or Instr(ACHTNo,chr(39)||'" & aryACHID(i) & "'||chr(39)) > 0", "Instr(ACHTNo,chr(39)||'" & aryACHID(i) & "'||chr(39)) > 0")
        Next i
        '*************************************************************
        
        '*********************************************************************
        '#3728 如果是取消授權不用尋找ACHTNO By Kin 2008/03/11
        If CmbApplyType.ListIndex = 0 Or CmbApplyType.ListIndex = 2 Then
            If strWhereACHID <> Empty Then subAnd "(" & strWhereACHID & ")"
        End If
        '*********************************************************************
'        If txtACHTNo <> "" Then subAnd "Instr(ACHTNo,chr(39)||'" & strInACHTNO & "'||chr(39)) > 0"
        
        '#2699 測試不OK,Where條件SO106.CitemStr必需要符合CD008裡的CitmCodeStr By Kin 2007/10/24
        '#4106 增加判斷ACHType參數 By Kin 2008/09/22
'        strINCD008Where = "Exists(Select CitemCode From " & GetOwner & "SO003 B Where " & _
'                  "B.Custid = A.Custid " & _
'                  "And B.CompCode=A.CompCode " & _
'                  "And instr(','||A.Citemstr||',',','||Chr(39)||B.Seqno||Chr(39)||',')>0 " & _
'                  "And Exists(Select * From " & GetOwner & "CD068 C Where " & _
'                  "instr(','||C.Citemcodestr||',',','||B.CitemCode||',')>0 " & _
'                  " And C.BillHeadFmt In(" & strInBillHeadFmt & ") And C.ACHType=1 ))"
        '#7049 改用CD068A.CitemCode By Kin 2015/07/08
        strINCD008Where = "Exists(Select CitemCode From " & GetOwner & "SO003 B Where " & _
                " B.Custid = A.Custid " & _
                " And B.CompCode=A.CompCode " & _
                " And instr(','||A.Citemstr||',',','||Chr(39)||B.Seqno||Chr(39)||',')>0 " & _
                " And B.CitemCode In (Select CitemCode From " & GetOwner & "CD068A " & _
                " Where Exists (Select * From CD068 Where CD068.BillHeadFmt=CD068A.BillHeadFmt " & _
                                    " And CD068.BillHeadFmt In (" & strInBillHeadFmt & ") And  CD068.ACHType = 2 )))"
                
        '依授權類別串不同的篩選條件
        Select Case CmbApplyType.ListIndex
            '#3946 改用MasterId做尋找 By Kin 2008/05/30
            Case 0
                'A 新增授權
                strWhere = " A.SnactionDate Is Null And A.SendDate Is Null And nvl(A.StopFlag,0) = 0 And " & strINCD008Where & _
                            " And Exists(Select * From " & strOwnerName & " SO106A B" & _
                            " Where A.MasterId=B.MasterId And B.StopFlag<>1 And B.StopDate is Null And ACHTNO IN(" & strInACHTNO & "))"
                            
                            
            Case 1
                
                '問題集2835 增加判斷如果有下停用日期的話,StopFlag=1，如果沒下的話，StopFlga=0
                'D 取消授權
                If gdaStopDate1.GetValue = "" And gdaStopDate2.GetValue = "" Then
                    'strWhere = " DeAuthorize=1 And nvl(StopFlag,0) = 0 "
                    strCancelWhere = IIf(strCancelWhere = Empty, "Nvl(A.StopFlag,0)=0", strCancelWhere & " And Nvl(A.StopFlag,0)=0")
                Else
                    'strWhere = " DeAuthorize=1 And StopFlag=1 "
                    strCancelWhere = IIf(strCancelWhere = Empty, "Nvl(A.StopFlag,0)=1", strCancelWhere & " And Nvl(A.StopFlag,0)=1")
                End If
                '#3946 改用MasterId做尋找 By Kin 2008/05/30
                If strCancelWhere <> Empty Then strCancelWhere = " And Exists(Select * From " & strOwnerName & "SO106A B " & _
                                                                " Where A.MasterId=B.MasterId And ACHTNO IN(" & strInACHTNO & ") And " & strCancelWhere & ")"
            
        End Select
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "subChoose")
End Sub

Public Sub subAnd(strCH As String)
  On Error GoTo ChkErr
    If strChoose <> "" Then
        strChoose = strChoose & " And " & strCH
    Else
        strChoose = strCH
    End If
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "subAnd")
End Sub

Private Function InsertFinal(rs As ADODB.Recordset) As Boolean
  On Error GoTo ChkErr
    
    Dim intReserve As Integer
    Dim strData As Variant
    Dim strBankId As String
    strBankId = GetRsValue("Select BankID From " & strOwnerName & "CD018 Where CodeNo = " & gilBank.GetCodeNo, gcnGi) & ""
    '資料別
    strData = "2"
    '委託機構代號
    strData = strData & GetString(strBankId, 3, giLeft, False)
    '保留欄
    strData = strData & Space(4)
    '媒體產生日期
    strData = strData & GetString(GetRealDateTran(gdaSendDate.GetValue), 8, giRight, True)
    '批號
    strData = strData & "001"
    '建檔記號
    strData = strData & "B"
    '總筆數
    strData = strData & GetString(intSeq, 6, giRight, True)
    '資料建檔日期
    strData = strData & Space(8)
    '錯誤筆數
    strData = strData & "000000"
    '成功筆數
    strData = strData & "000000"
    '保留欄
    strData = strData & Space(54)
    WriteLineData strData, DataPath
    InsertFinal = True
    Exit Function
ChkErr:
    ErrSub Me.Name, "InsertFinal"
End Function

Private Function GetCustID(strReadLine As String) As Long
    On Error Resume Next
    
        GetCustID = Trim(Val(Mid(strReadLine, 54, 20)))
End Function

Private Function GetACHCustID(strReadLine As String) As String
    On Error Resume Next
        GetACHCustID = Trim(Mid(strReadLine, 42, 20))
End Function

Private Function GetAccID(strReadLine As String) As String
    On Error Resume Next
    Dim intLen As Integer
       ' intLen = Val(GetRsValue("Select ActLength From " & GetOwner & "CD018" & " Where BankId2 = '" & GetBankCode(strReadLine) & "'") & "")
        'GetAccID = Trim(Right(Mid(strReadLine, 27, 14), intLen))
        GetAccID = Trim(Mid(strReadLine, 28, 14))
End Function

Private Function GetBankCode(strReadLine As String) As String
    On Error Resume Next
        GetBankCode = Trim(Mid(strReadLine, 20, 7))
End Function
Private Function getCreateDate(strReadLine As String) As String
    On Error Resume Next
    getCreateDate = Mid(strReadLine, 27, 8)
End Function
Private Function getApplyDate(strReadLine As String) As String
    On Error Resume Next
    Dim Ret As String
    Ret = Trim(Mid(strReadLine, 9, 8))
    If (Len(Ret & "") = 0) Or (Len(Ret) <> 8) Or (Not IsNumeric(Ret)) Then
        Ret = Format(RightNow, "yyyymmdd")
    End If
    getApplyDate = Ret
End Function
Private Function GetTxtDate(strReadLine As String) As String
    On Error Resume Next
        GetTxtDate = CStr(Val(Mid(strReadLine, 72, 8)) + 19110000)
End Function
'#3728 提回時找出相對應的SO106A By Kin 2008/03/12
'參數 1.strType:授權類別 2.strRowId:SO106.RowId  3.strData:提回電子檔目前讀取的Record
Private Function GetSO106A(ByVal strType As String, ByVal strRowId As String, ByVal strData As String) As String
  On Error GoTo ChkErr
    Select Case strType
        Case "1"
            '#3946 改用MasterId做PK By Kin 2008/05/30
            GetSO106A = "Select A.RowId,A.* From " & GetOwner & "SO106A A " & _
                        "Where A.MasterId='" & strRowId & "'" & _
                        " And A.ACHTNO In(" & strInACHTNO & ")" & _
                        " And A.RecordType=0" & _
                        " And A.AuthorizeStatus is null" & _
                        " And A.StopFlag<>1"
        Case "2"
            GetSO106A = "Select A.RowId,A.* From " & GetOwner & "SO106A A " & _
                        "Where A.MasterId='" & strRowId & "'" & _
                        " And A.ACHTNO In(" & strInACHTNO & ")" & _
                        " And A.RecordType=1" & _
                        " And A.AuthorizeStatus is null" & _
                        " And A.StopFlag=1" & _
                        IIf(gdaStopDate1.GetValue <> Empty, " And A.StopDate>=To_Date('" & gdaStopDate1.GetValue & "','YYYYMMDD')", "") & _
                        IIf(gdaStopDate1.GetValue <> Empty, " And A.StopDate<To_Date('" & gdaStopDate1.GetValue & "','YYYYMMDD')+1", "")
        Case "3"
            GetSO106A = "Select A.RowId,A.* From " & GetOwner & "SO106A A " & _
                        "Where A.MasterId='" & strRowId & "'" & _
                        " And A.ACHTNO In(" & strInACHTNO & ")" & _
                        " And A.AuthorizeStatus = 1 " & _
                        " And StopFlag <> 1 "
                                              
        Case "4"
            GetSO106A = "Select A.RowId,A.* From " & GetOwner & "SO106A A" & _
                        " Where A.MasterId='" & strRowId & "'" & _
                        " And A.ACHTNO In(" & strInACHTNO & ")" & _
                        " And A.RecordType=0" & _
                        " And A.AuthorizeStatus=2"

    End Select
  
    Exit Function
ChkErr:
    ErrSub Me.Name, "GetSO106A"
End Function
'#3728 取消授權成功時要連SO003一起停用
Private Function StopSO003(ByVal strRowId As String, ByRef rsSO106A As ADODB.Recordset) As Boolean
  On Error GoTo ChkErr
    Dim strSQL As String
    Dim rsQrySO106 As New ADODB.Recordset
    If strRowId = Empty Then Exit Function
    If rsSO106A.State = adStateClosed Then Exit Function
    If rsSO106A.RecordCount = 0 Then Exit Function
    strSQL = "Select * From " & strOwnerName & "SO106 Where RowId='" & strRowId & "'"
    If Not GetRS(rsQrySO106, strSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
    If Not rsQrySO106.EOF Then
        strSQL = "Update " & strOwnerName & "SO003 Set " & _
                 "BANKCODE=NULL" & _
                 ",BANKNAME=NULL" & _
                 ",ACCOUNTNO=NULL" & _
                 ",InvSeqNO=NULL" & _
                 " Where CompCode=" & gilCompCode.GetCodeNo & _
                 " And CustId=" & rsQrySO106("CustId") & _
                 " And CitemCode In(" & rsSO106A("CitemCodeStr") & ")" & _
                 " And ACCOUNTNO='" & rsQrySO106("ACCOUNTID") & "'"
        gcnGi.Execute strSQL
        StopSO003 = True
    End If
    On Error Resume Next
    Call CloseRecordset(rsQrySO106)
    Exit Function
ChkErr:
    ErrSub Me.Name, "StopSO003"
End Function

'#3278 授權失敗或取消授權時,要將對應的ACH、收費項目給拿掉 By Kin 2008/03/18
Private Function UpdateSO106(ByVal strRowId As String, ByRef rsSO106A As ADODB.Recordset) As Boolean
  On Error GoTo ChkErr
    Dim strQry As String
    Dim rsUpdSO106 As New ADODB.Recordset
    Dim rsSO003 As New ADODB.Recordset
    Dim strOldACHTNO As String
    Dim strOldACHTDesc As String
    Dim strOldCitemStr As String
    Dim strQrySeqNo As String
    If strRowId = Empty Then Exit Function
    If rsSO106A.State = adStateClosed Then Exit Function
    If rsSO106A.RecordCount = 0 Then Exit Function
    strQry = "Select A.RowId,A.* From " & strOwnerName & "SO106 A Where RowId='" & strRowId & "'"
    If Not GetRS(rsUpdSO106, strQry, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
    If Not rsUpdSO106.EOF Then
        With rsUpdSO106
            If Not IsNull(.Fields("ACHTNO")) Then strOldACHTNO = .Fields("ACHTNO")
            If Not IsNull(.Fields("ACHTDesc")) Then strOldACHTDesc = .Fields("ACHTDesc")
            If Not IsNull(.Fields("CitemStr")) Then strOldCitemStr = .Fields("CitemStr")
            strQry = "Select * From " & strOwnerName & "SO003" & _
                    " Where CustId=" & .Fields("CustId") & _
                    " And CompCode=" & gilCompCode.GetCodeNo & _
                    " And CitemCode IN(" & rsSO106A("CitemCodeStr") & ")"
            If Not GetRS(rsSO003, strQry, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
            If Not rsSO003.EOF Then
                rsSO003.MoveFirst
                Do While Not rsSO003.EOF
                    strQrySeqNo = "'" & rsSO003("SeqNo") & "'"
                    strOldCitemStr = Replace(Replace(strOldCitemStr, "," & strQrySeqNo, ""), strQrySeqNo, "")
                    rsSO003.MoveNext
                Loop
            End If
            If Left(strOldCitemStr, 1) = "," Then
                strOldCitemStr = Mid(strOldCitemStr, 2)
            End If
            If Right(strOldCitemStr, 1) = "," Then
                strOldCitemStr = Mid(strOldCitemStr, 1, Len(strOldCitemStr) - 1)
            End If
            
            strOldACHTNO = Replace(Replace(strOldACHTNO, ",'" & rsSO106A("ACHTNO") & "'", ""), "'" & rsSO106A("ACHTNO") & "'", "")
            strOldACHTDesc = Replace(Replace(strOldACHTDesc, ",'" & rsSO106A("ACHDesc") & "'", ""), "'" & rsSO106A("ACHDesc") & "'", "")
            If Left(strOldACHTNO, 1) = "," Then
                strOldACHTNO = Mid(strOldACHTNO, 2)
            End If
            If Right(strOldACHTNO, 1) = "," Then
                strOldACHTNO = Mid(strOldACHTNO, 1, Len(strOldACHTNO) - 1)
            End If
            If Left(strOldACHTDesc, 1) = "," Then
                strOldACHTDesc = Mid(strOldACHTDesc, 2)
            End If
            If Right(strOldACHTDesc, 1) = "," Then
                strOldACHTDesc = Mid(strOldACHTDesc, 1, Len(strOldACHTDesc) - 1)
            End If


            rsUpdSO106("ACHTNO") = IIf(strOldACHTNO = "", Null, strOldACHTNO)
            rsUpdSO106("ACHTDesc") = IIf(strOldACHTDesc = "", Null, strOldACHTDesc)
            rsUpdSO106("CitemStr") = IIf(strOldCitemStr = "", Null, strOldCitemStr)
            rsUpdSO106.Update
        End With
    End If
    UpdateSO106 = True
    On Error Resume Next
    Call CloseRecordset(rsUpdSO106)
    Call CloseRecordset(rsSO003)
  Exit Function
ChkErr:
    ErrSub Me.Name, "UpdateSO106"
End Function

Private Function UpdateSO106A(ByVal rsSO106A As ADODB.Recordset, ByVal strType As String) As Boolean
  On Error GoTo ChkErr
    Dim strChkSQL As String
    Dim strSO106Err As String
    Dim strTime As String
    Dim rsErr As New ADODB.Recordset
    blnStopAll = False
    If rsSO106A.RecordCount > 0 Then
        Select Case strType
            Case "RETURNOK"
                rsSO106A("AuthorizeStatus") = 1
                rsSO106A.Update
            Case "2"    '取消授權
                rsSO106A("AuthorizeStatus") = 2
                'rsSO106A("Notes") = "已提回終止扣款!  失敗日期:" & Format(Now, "yyyy/mm/dd")
                rsSO106A.Update
                '#3946 改用MasterId
                strChkSQL = "Select Count(Decode(AuthorizeStatus,Null,Null)) A," & _
                                    "Count(Decode(AuthorizeStatus,1,1)) B," & _
                                    "Count(Decode(AuthorizeStatus,2,2)) C," & _
                                    "Count(Decode(AuthorizeStatus,3,3)) D" & _
                            " From " & GetOwner & "SO106A" & _
                            " Where MasterId='" & rsSO106A("MasterId") & "'"
                            
                Set rsErr = gcnGi.Execute(strChkSQL)
                If rsErr("A") = 0 Then
                    If rsErr("C") = rsErr("B") Then
                        blnStopAll = True
                    Else
                        blnStopAll = False
                    End If
                Else
                    blnStopAll = False
                End If
            Case "3"
                rsSO106A("AuthorizeStatus") = 2
                'rsSO106A("Notes") = "郵局終止扣款!  失敗日期:" & Format(Now, "yyyy/mm/dd")
                rsSO106A.Update
                strChkSQL = "Select Count(Decode(AuthorizeStatus,Null,Null)) A," & _
                                    "Count(Decode(AuthorizeStatus,1,1)) B," & _
                                    "Count(Decode(AuthorizeStatus,2,2)) C," & _
                                    "Count(Decode(AuthorizeStatus,3,3)) D" & _
                            " From " & GetOwner & "SO106A" & _
                            " Where MasterId='" & rsSO106A("MasterId") & "'"
                Set rsErr = gcnGi.Execute(strChkSQL)
                 If rsErr("A") = 0 And rsErr("D") = 0 Then
                    blnStopAll = True
                Else
                    blnStopAll = False
                End If
            Case "E"        '授權失敗,檢查授權的
                rsSO106A("AuthorizeStatus") = 3
                rsSO106A("UpdTime") = GetDTString(RightNow)
                rsSO106A("UpdEn") = garyGi(1)
                '#3946 失敗時備註多增加系統日期
                rsSO106A("Notes") = strErrAuth & "   失敗日期:" & Format(Now, "yyyy/mm/dd")
                rsSO106A.Update
                '#3946 改用MasterId為PK By Kin 2008/05/30
                strChkSQL = "Select Count(*) From " & GetOwner & "SO106A" & _
                            " Where AchtNO<>'" & rsSO106A("ACHTNO") & "'" & _
                            " And ACHDesc<>'" & rsSO106A("ACHDesc") & "'" & _
                            " And MasterId=" & rsSO106A("MasterId") & _
                            " And (AuthorizeStatus=1 or AuthorizeStatus is Null)"
                If gcnGi.Execute(strChkSQL)(0) = 0 Then
                    blnStopAll = True
                    '#3946 改用MasterId為PK By Kin 2008/05/30
                    strChkSQL = "Select * From " & GetOwner & "SO106A " & _
                                " Where MasterId=" & rsSO106A("MasterId") & _
                                " And AuthorizeStatus=3"
                    If Not GetRS(rsErr, strChkSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
                    strSO106Err = Empty
                    Do While Not rsErr.EOF
                        strSO106Err = IIf(strSO106Err = Empty, rsErr("ACHDesc") & ":" & rsErr("Notes"), _
                                            strSO106Err & vbCrLf & rsErr("ACHDesc") & ":" & rsErr("Notes"))
                    
                        rsErr.MoveNext
                    Loop
                    If strSO106Err <> Empty Then strSO106Err = Left(strSO106Err, 125)
                    '#3946 改用MasterId為PK By Kin 2008/05/30
                    'Call CloseRecordset(rsErr)
                    '***************************************************************************************************************************************************************************
                    '#4126 原本將備註直接Update掉,此需求要求要串之前的Note內容 By Kin 2008/10/09
                    If Not GetRS(rsErr, "Select Note From " & GetOwner & "SO106 Where Masterid=" & rsSO106A("MasterId"), gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
                    'Set rsErr = gcnGi.Execute("Select Note From " & GetOwner & "SO106 Where Masterid=" & rsSO106A("MasterId"))
                    If Not rsErr.EOF Then
                        strSO106Err = MidMbcs(IIf(rsErr("Note") & "" <> "", rsErr("Note") & vbCrLf & strSO106Err, strSO106Err), 1, rsErr.Fields(0).DefinedSize)
                        If strSO106Err <> Empty Then rsErr("Note").Value = strSO106Err: rsErr.Update
                    End If
                    '***************************************************************************************************************************************************************************
                    'gcnGi.Execute ("Update " & GetOwner & "SO106 Set Note=Note||Chr(13)||'" & strSO106Err & "' Where MasterId=" & rsSO106A("MasterId"))
                Else
                    blnStopAll = False
                End If
                
            
            
            
        End Select
    End If
    UpdateSO106A = True
    On Error Resume Next
    Call CloseRecordset(rsErr)
    Exit Function
ChkErr:
    ErrSub Me.Name, "UpdateSO106A"
End Function

'#3728 多帳戶依據設定改為複選 By Kin 2008/03/10
Private Sub gimBillHeadFmt_Change()
    On Error Resume Next
    Dim strCitemCode As String
    Dim strCodeNo As String
    Dim rsCD019  As New ADODB.Recordset
    Dim strQry As String
    Dim strACHQry As String
    Screen.MousePointer = vbHourglass
    Me.Enabled = False
    If gimBillHeadFmt.GetDispStr <> "(全選)" Then
        gimBillHeadFmt.SetDispStr Replace(gimBillHeadFmt.GetQueryCode, "'", "")

    End If
    If gimBillHeadFmt.GetDispStr = Empty Then
        gimCitemCode.Clear
        txtACHTNo.Text = Empty
    End If
    If gimBillHeadFmt.GetDispStr = "(全選)" Then
        '#3946 要過濾 ACHTNO Is Not Null By Kin 2008/05/30
        '#4094 過濾停用要取消 By Kin 2008/09/17
        '#4106 增加判斷ACHTYPE參數 BY Kin 2008/09/22
        strQry = "Select CitemCodeStr From " & GetOwner & "CD068 " & _
                 "Where ACHTNO Is Not Null And ACHTDesc Is Not Null And ACHType=2"
        '#7049 改用CD068A.CitemCode By Kin 2015/07/08
        strQry = "Select CitemCode From " & GetOwner & "CD068A " & _
                " Where  Exists (Select * From " & GetOwner & "CD068 " & _
                        "Where ACHTNO Is Not Null And ACHTDesc Is Not Null And ACHType=2 And CD068.BillHeadFmt = CD068A.BillHeadFmt)"
    Else
        '#4094 過濾停用要取消 By Kin 2008/09/17
        '#4094 原本是代BillHeadFmt 現在改為ACHTDESC By Kin 2008/09/17
        '#4106 增加判斷ACHTYPE參數 BY Kin 2008/09/22
        strQry = "Select CitemCodeStr From " & GetOwner & "CD068 " & _
                 "Where ACHTDESC In(" & gimBillHeadFmt.GetQueryCode & ") And ACHType=2 "
         '#7049 改用CD068A.CitemCode By Kin 2015/07/08
        strQry = "Select CitemCode From " & GetOwner & "CD068A " & _
                " Where  Exists (Select * From " & GetOwner & "CD068 " & _
                        "Where ACHTDESC In(" & gimBillHeadFmt.GetQueryCode & ")   And ACHType=2 And CD068.BillHeadFmt = CD068A.BillHeadFmt)"
    End If
    strCodeNo = gcnGi.Execute(strQry).GetString(, , , ",")
    If Len(strCodeNo) > 0 Then
       strCodeNo = Mid(strCodeNo, 1, Len(strCodeNo) - 1)
    Else
        Me.Enabled = True
        Screen.MousePointer = vbDefault
       Exit Sub
    End If
'    If Not GetRS(rsCD019, "Select CodeNo From " & GetOwner & "CD019 Where CodeNo In (" & strCodeNo & ") And StopFlag=0") Then Exit Sub
    '#7219
    If Not GetRS(rsCD019, "Select CodeNo From " & GetOwner & "CD019 Where CodeNo In (" & strQry & ") And StopFlag=0") Then Exit Sub
    If rsCD019.EOF Then Exit Sub
    If gimBillHeadFmt.GetDispStr = "(全選)" Then
        '#3946 要過濾 ACHTNO Is Not Null By Kin 2008/05/30
        '#4094 過濾停用要取消 By Kin 2008/09/17
        '#4106 增加判斷ACHTYPE參數 BY Kin 2008/09/22
        strACHQry = "Select ACHTNO From " & GetOwner & "CD068 Where  ACHTNO Is Not NULL And ACHTDesc is Not Null And ACHType=2"
    Else
        '#4094 原本是代BillHeadFmt 現在改為ACHTDESC By Kin 2008/09/17
        '#4106 增加判斷ACHTYPE參數 BY Kin 2008/09/22
        strACHQry = "Select ACHTNO From " & GetOwner & "CD068 Where ACHTDESC IN(" & gimBillHeadFmt.GetQueryCode & ") And ACHType=2 "
    End If
    
    strACHQry = gcnGi.Execute(strACHQry).GetString(, , , ",")
    txtACHTNo.Text = Left(strACHQry, Len(strACHQry) - 1)
    strCitemCode = rsCD019.GetString(, , , ",")
    strCitemCode = Left(strCitemCode, Len(strCitemCode) - 1)
    'gimCitemCode.Filter = gimCitemCode.Filter
    
    
    DoEvents
    gimCitemCode.SetQueryCode strCitemCode
    Me.Enabled = True
    Screen.MousePointer = vbDefault

End Sub
'#3946 取得SO106.MasterId的Sequence By Kin 2008/05/30
Private Function Get_MasterID_Seq() As String
  On Error GoTo ChkErr
    Get_MasterID_Seq = RPxx(gcnGi.Execute("SELECT " & GetOwner & "S_SO106_MasterId.NEXTVAL FROM DUAL").GetString & "")
    'Get_Cmd_Seq_No = Format(Get_Cmd_Seq_No, "00000000")
  Exit Function
ChkErr:
    ErrSub "mod_SysLib", "Get_MasterID_Seq"
End Function


