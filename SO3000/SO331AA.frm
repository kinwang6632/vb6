VERSION 5.00
Object = "{3BB22E57-824A-11D3-BD70-0080C8F80BC4}#3.4#0"; "GiGridR.ocx"
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Begin VB.Form frmSO331AA 
   BorderStyle     =   3  '雙線固定對話方塊
   Caption         =   "派工單所附收費資料登錄 [SO331AA]"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   BeginProperty Font 
      Name            =   "新細明體"
      Size            =   9.75
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SO331AA.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '螢幕中央
   Begin VB.Frame fraMain 
      Height          =   2025
      Left            =   60
      TabIndex        =   14
      Top             =   510
      Width           =   11715
      Begin VB.CheckBox chkNineCode 
         Caption         =   "使用9碼模式登錄"
         Height          =   345
         Left            =   2730
         TabIndex        =   34
         Top             =   1470
         Width           =   2115
      End
      Begin VB.TextBox txtBillNo 
         Height          =   375
         Left            =   990
         MaxLength       =   15
         TabIndex        =   1
         Top             =   240
         Width           =   1485
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "F2. 存檔"
         Height          =   375
         Left            =   9120
         TabIndex        =   4
         Top             =   1410
         Width           =   1185
      End
      Begin VB.CheckBox chkDate 
         Caption         =   "入帳日以簽收日為主"
         Height          =   345
         Left            =   270
         TabIndex        =   3
         Top             =   1470
         Value           =   1  '核取
         Width           =   2115
      End
      Begin Gi_Date.GiDate gdaRealDate 
         Height          =   375
         Left            =   990
         TabIndex        =   2
         Top             =   870
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
         Enabled         =   0   'False
      End
      Begin VB.Label lblCustId 
         AutoSize        =   -1  'True
         Caption         =   "99999999"
         Height          =   195
         Left            =   6840
         TabIndex        =   28
         Top             =   330
         Width           =   720
      End
      Begin VB.Label lbl3 
         AutoSize        =   -1  'True
         Caption         =   "姓名:"
         Height          =   195
         Left            =   2760
         TabIndex        =   27
         Top             =   330
         Width           =   435
      End
      Begin VB.Label lblBillNo 
         AutoSize        =   -1  'True
         Caption         =   "派工單號"
         Height          =   195
         Left            =   150
         TabIndex        =   26
         Top             =   330
         Width           =   780
      End
      Begin VB.Label lbl4 
         AutoSize        =   -1  'True
         Caption         =   "客戶編號:"
         Height          =   195
         Left            =   5940
         TabIndex        =   25
         Top             =   330
         Width           =   825
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Caption         =   "XXXXXXXXXXXXXXXXXX"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   3300
         TabIndex        =   24
         Top             =   330
         Width           =   2430
      End
      Begin VB.Label lblRealDate 
         AutoSize        =   -1  'True
         Caption         =   "入帳日期"
         Height          =   195
         Left            =   150
         TabIndex        =   23
         Top             =   990
         Width           =   780
      End
      Begin VB.Label lblSubTotal1 
         AutoSize        =   -1  'True
         Caption         =   "應收小計:"
         Height          =   195
         Left            =   7920
         TabIndex        =   22
         Top             =   330
         Width           =   825
      End
      Begin VB.Label lblSubTotal 
         Caption         =   "XXXXXXXXXXXXXXXX"
         Height          =   255
         Left            =   8910
         TabIndex        =   21
         Top             =   330
         Width           =   2115
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "派工類別:"
         Height          =   195
         Left            =   2730
         TabIndex        =   20
         Top             =   960
         Width           =   825
      End
      Begin VB.Label lblClass 
         AutoSize        =   -1  'True
         Caption         =   "XXXXXXXXXXX"
         Height          =   195
         Left            =   3660
         TabIndex        =   19
         Top             =   960
         Width           =   1485
      End
      Begin VB.Label lblTime2 
         AutoSize        =   -1  'True
         Caption         =   "完工時間:"
         Height          =   195
         Left            =   5610
         TabIndex        =   18
         Top             =   960
         Width           =   825
      End
      Begin VB.Label lblTime 
         AutoSize        =   -1  'True
         Caption         =   "YYYY/MM/DD HH:MM"
         Height          =   195
         Left            =   6540
         TabIndex        =   17
         Top             =   960
         Width           =   1920
      End
      Begin VB.Label lblSignDate 
         AutoSize        =   -1  'True
         Caption         =   "YYYY/MM/DD"
         Height          =   195
         Left            =   9630
         TabIndex        =   16
         Top             =   960
         Width           =   1230
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "簽收日期:"
         Height          =   195
         Left            =   8700
         TabIndex        =   15
         Top             =   960
         Width           =   825
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      Height          =   375
      Left            =   10200
      TabIndex        =   10
      Top             =   90
      Width           =   1215
   End
   Begin VB.Frame fraData 
      Height          =   4815
      Left            =   60
      TabIndex        =   11
      Top             =   2580
      Width           =   11715
      Begin VB.CommandButton cmdDetDelete 
         Caption         =   "作廢(&3)"
         Height          =   375
         Left            =   3570
         TabIndex        =   7
         Top             =   180
         Width           =   930
      End
      Begin VB.CommandButton cmdDetEdit 
         Caption         =   "修改(&2)"
         Height          =   375
         Left            =   2430
         TabIndex        =   6
         Top             =   180
         Width           =   930
      End
      Begin VB.CommandButton cmdDetAdd 
         Caption         =   "新增(&1)"
         Height          =   375
         Left            =   1320
         TabIndex        =   5
         Top             =   180
         Width           =   930
      End
      Begin prjGiGridR.GiGridR ggrData 
         Height          =   3795
         Left            =   180
         TabIndex        =   8
         Top             =   900
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   6694
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "(收費項目、出單金額、期數、實收金額、期限起始、期限截止、作廢、備註)"
         Height          =   315
         Left            =   180
         TabIndex        =   13
         Top             =   630
         Visible         =   0   'False
         Width           =   8715
      End
      Begin VB.Label lblDetail 
         AutoSize        =   -1  'True
         Caption         =   "單據內容"
         Height          =   195
         Left            =   300
         TabIndex        =   12
         Top             =   300
         Width           =   780
      End
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "F3. 已登查詢"
      Height          =   375
      Left            =   8370
      TabIndex        =   9
      Top             =   90
      Width           =   1665
   End
   Begin prjGiList.GiList gilCompCode 
      Height          =   285
      Left            =   1020
      TabIndex        =   0
      Top             =   90
      Width           =   2430
      _ExtentX        =   4286
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FldWidth1       =   600
      FldWidth2       =   1500
      F5Corresponding =   -1  'True
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      Caption         =   "登錄單據數:"
      Height          =   195
      Left            =   3690
      TabIndex        =   33
      Top             =   150
      Width           =   1020
   End
   Begin VB.Label lblBillCnt 
      AutoSize        =   -1  'True
      Caption         =   "999999"
      Height          =   195
      Left            =   4950
      TabIndex        =   32
      Top             =   150
      Width           =   540
   End
   Begin VB.Label lbl2 
      AutoSize        =   -1  'True
      Caption         =   "登錄總金額:"
      Height          =   195
      Left            =   6000
      TabIndex        =   31
      Top             =   150
      Width           =   1020
   End
   Begin VB.Label lblTotalAmt 
      AutoSize        =   -1  'True
      Caption         =   "999,999,999"
      Height          =   195
      Left            =   7260
      TabIndex        =   30
      Top             =   150
      Width           =   900
   End
   Begin VB.Label lblComp 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "公  司  別"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   150
      TabIndex        =   29
      Top             =   150
      Width           =   765
   End
End
Attribute VB_Name = "frmSO331AA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private XstrTranDate As String

Private XintPara5 As Integer '延後幾日為次收日
Private XintPara6 As Integer '收費日登錄期限
Private XintPara8 As Integer '隨收費資料改變1
Private XintPara9 As Integer '隨收費資料改變2


'----------------------------------------------------------------------------------------------
Private rsSo3311B As New ADODB.Recordset
Private rsArray As New ADODB.Recordset
Private rsArrayNo As New ADODB.Recordset

Private intDayCut As Integer
Private intPara6 As Integer
Private strSQL As String

Private intCustid As Long '記錄客戶編號(供So3311F使用）
Private strCustName As String '記錄客戶姓名(供So3311F使用）

Private strConnString As String
Private strRowId As String '暫存RowId以傳至So3311F 使用
Private cnMDB As New ADODB.Connection
Private blnCanClose As Boolean '是否可結束本表單 True=Yes ,False =No
Dim intCount As Long
Dim lngSeqNo As Long
Dim lngSO003BudgetPeriod As Long
Dim lngSO003Period As Long

Private Sub chkDate_Click()
    On Error Resume Next
        If chkDate.Value = 1 Then
            gdaRealDate.Text = ""
            gdaRealDate.Enabled = False
        Else
            gdaRealDate.Enabled = True
        End If
End Sub

Private Sub chkDate_Validate(Cancel As Boolean)
    On Error Resume Next
        If chkDate.Value = 1 Then
            gdaRealDate.Text = ""
            gdaRealDate.Enabled = False
        Else
            gdaRealDate.Enabled = True
        End If
End Sub

Private Sub cmdCancel_Click()
    On Error Resume Next
        Unload Me
End Sub

Private Sub cmdDetAdd_Click()
    On Error Resume Next
        With frmSO331AF
            Set .uParentForm = Me
            Set .Conn = cnMDB
            .uDBType = giAccessDb
            .EditMode = giEditModeInsert
            .uBillNo = rsArray("BillNO").Value & ""
            .uItem = GetMaxItem(rsArray)
            .uPrtSNo = rsArray("PrtSno").Value & ""
            .uClctEn = rsArray("ClctEn").Value & ""
            .uClctName = rsArray("ClctName").Value & ""
            .uPTCode = IIf(IsNull(rsArray("PTCode").Value), 0, rsArray("PTCode").Value)
            .uPTName = rsArray("PTName").Value & ""
            .uRealDate = rsArray("RealDate").Value & ""
            .uCustId = rsArray("CustID").Value
            .uCustName = rsArray("CustName").Value
            gServiceType = rsArray("ServiceType").Value
            .Show vbModal
        End With
        Call RefreshData
End Sub

Private Sub cmdDetDelete_Click()
        '檢查是否可以作廢，若已作廢，則錯誤：顯示"已作發資料不可再作廢 "，無以下動作！
    On Error Resume Next
        
        If rsArray("CancelFlag").Value = 1 Then MsgBox "已作發資料不可再作廢", vbExclamation, "訊息！": Exit Sub
        '#XXXX 判斷是否有第二層優惠 By Kin 2009/04/01
        If ChkHave2Discount(rsArray("CitemCode")) Then MsgBox "有第2層優惠不能作廢!!", vbCritical, gimsgPrompt:  Exit Sub
        If Chk2Discount(rsArray("CitemCode")) Then MsgBox "第2層優惠不能作廢!!", vbCritical, gimsgPrompt:  Exit Sub
    On Error GoTo chkErr
        cnMDB.BeginTrans
        If Len(rsArray("RowId") & "") < 10 Then
            cnMDB.Execute "Delete From So3311 Where RowId='" & rsArray("RowId") & "'"
            MsgBox "刪除成功！", vbInformation, "訊息！"
        Else
            If Not rsArray.EOF Then
                With frmSO1110E
                    Set .uParentForm = Me
                    .uType = 1  '收費資料
                    Set .uRS = rsArray
                    .uBillNo = rsArray("BillNO")
                    .uItem = rsArray("Item")
                    .Show vbModal
                End With
            End If
        End If
        cnMDB.CommitTrans
    On Error Resume Next
        Call RefreshData
    Exit Sub
chkErr:
    cnMDB.RollbackTrans
    Call ErrSub(Me.Name, "cmdDetDelete")
End Sub

Private Sub cmdDetEdit_Click()
    On Error Resume Next
        If rsArray("CancelFlag").Value = 1 Then MsgBox "已作廢資料不能修改！", vbExclamation, "訊息！": Exit Sub
        With frmSO331AF
            Set .uParentForm = Me
            Set .Conn = cnMDB
            .uDBType = giAccessDb
            .EditMode = giEditModeEdit
            .uBillNo = rsArray("BillNO").Value
            .uPrtSNo = rsArray("PrtSno").Value & ""
            .uItem = rsArray("Item").Value
            .uRowId = rsArray("RowId").Value
            .uClctEn = rsArray("ClctEn").Value & ""
            .uClctName = rsArray("ClctName").Value & ""
            .uPTCode = IIf(IsNull(rsArray("PTCode").Value), 0, rsArray("PTCode").Value)
            .uPTName = rsArray("PTName").Value & ""
            .uRealDate = rsArray("RealDate").Value & ""
            .uCustId = rsArray("CustID").Value
            .uCustName = rsArray("CustName").Value
            gServiceType = rsArray("ServiceType").Value
            .Show vbModal
        End With
        Call RefreshData
End Sub

Private Sub cmdQuery_Click()
    On Error Resume Next
        With frmSO3311C
            Set .uParentForm = Me
            .Show vbModal
        End With
End Sub

Private Function IsDataOk() As Boolean
    On Error GoTo chkErr
        If Not ChkDTok Then Exit Function
        If rsArray.State = adStateClosed Then MsgBox "請先輸入資料！", vbExclamation, "訊息！": Exit Function
        If rsArray.EOF Or (rsArray.RecordCount <= 0) Then MsgBox "無資料，請先輸入資料！", vbExclamation, "訊息！": Exit Function
        ' lblBillCnt
        '必要欄位檢查，若不過不可繼續！
        '<派工單號＞
        '若<入帳日期以簽收日為主>not checked , 則<入帳日期>為必要！
        If chkDate.Value = 0 And Len(gdaRealDate.Text) = 0 Then
            MsgBox "若入帳日期不以簽收日為主,則入帳日期為必要", vbExclamation, "訊息！"
            gdaRealDate.SetFocus
            Exit Function
        End If
        If rsArray.RecordCount > 0 Then
            rsArray.MoveFirst
            Do While Not rsArray.EOF
                If rsArray("RealAmt") & "" <> rsArray("ShouldAmt") & "" And rsArray("CancelFlag") = 0 And Len(rsArray("STCode") & "") = 0 Then
                    MsgBox "實收金額與出單金額不相等,須有短收原因", vbCritical, giMsgWarning
                    If cmdDetEdit.Enabled Then cmdDetEdit.Value = True
                    Exit Function
                End If
                gServiceType = rsArray("ServiceType")
                gCustId = rsArray("CustId")
                If rsArray("RealPeriod") > 0 Then
                    '******************************************************************************************************************************************************
                    '#3465 測試不OK，新增一筆資料時會出現錯誤，因為新增時沒有RowId導致，儲存時改用BillNo+Item By Kin 2007/08/24
                    'lngSeqNo = Val(GetRsValue("Select SeqNo From " & GetOwner & "SO033 Where RowId = '" & rsArray("RowId") & "'") & "")
                    lngSeqNo = Val(GetRsValue("Select SeqNo From " & GetOwner & "SO033 Where BillNo = '" & rsArray("BillNo") & "' And Item=" & rsArray("Item")) & "")
                    If lngSeqNo = 0 Then If Not CallForm1132D(rsArray("CitemCode"), Me, giEditModeEdit, rsArray("BillNo"), lngSeqNo, lngSO003BudgetPeriod, lngSO003Period) Then Exit Function
                    'gcnGi.Execute "Update " & GetOwner & "SO033 Set SeqNo = " & lngSeqNo & " Where RowId = '" & rsArray("RowId") & "'"
                    '******************************************************************************************************************************************************
                    gcnGi.Execute "Update " & GetOwner & "SO033 Set SeqNo = " & lngSeqNo & " Where BillNo = '" & rsArray("BillNO") & "' And Item=" & rsArray("Item")
                End If
                
                rsArray.MoveNext
            Loop
        End If
        IsDataOk = True
    Exit Function
chkErr:
    ErrSub Me.Name, "IsDataOk"
End Function

Private Sub cmdSave_Click()
    On Error GoTo chkErr
    Dim lngBillCnt As Long
    Dim rsSaveSo074 As New ADODB.Recordset
    Dim strSQL As String
        
        If Not IsDataOk Then Exit Sub
        strSQL = "Select  * From " & GetOwner & "So074 Where 0=1"
        If Not GetRS(rsSaveSo074, strSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then MsgBox "連接失敗！請重新操作！", vbExclamation, "訊息！": Exit Sub
        intCount = intCount + 1
        rsArray.MoveFirst
        With rsSaveSo074
            While Not rsArray.EOF
                .AddNew
                .Fields("BillNo").Value = GetFieldValue(rsArray, "Billno")
                .Fields("Item").Value = GetFieldValue(rsArray, "Item")
                .Fields("Custid").Value = GetFieldValue(rsArray, "Custid")
                .Fields("CustName").Value = lblName.Caption
                .Fields("CitemCode").Value = GetFieldValue(rsArray, "CitemCode")
                .Fields("CitemName").Value = GetFieldValue(rsArray, "CitemName")
                .Fields("ShouldAmt").Value = GetFieldValue(rsArray, "ShouldAmt")
                .Fields("ShouldDate").Value = GetFieldValue(rsArray, "ShouldDate")
                .Fields("RealAmt").Value = NoZero(rsArray("RealAmt").Value, True)
                .Fields("RealDate").Value = GetFieldValue(rsArray, "RealDate")
                .Fields("RealPeriod").Value = NoZero(rsArray("RealPeriod").Value, True)
                .Fields("RealStartDate").Value = GetFieldValue(rsArray, "RealStartDate")
                .Fields("RealStopDate").Value = GetFieldValue(rsArray, "RealStopDate")
                .Fields("CancelFlag").Value = NoZero(rsArray("CancelFlag").Value, True)
                .Fields("EntryEn").Value = GetFieldValue(rsArray, "EntryEn")
                .Fields("Note").Value = GetFieldValue(rsArray, "Note")
                .Fields("CMCode").Value = GetFieldValue(rsArray, "CMCode")
                .Fields("CMName").Value = GetFieldValue(rsArray, "CMName")
                .Fields("ClctEn").Value = GetFieldValue(rsArray, "ClctEn")
                .Fields("ClctName").Value = GetFieldValue(rsArray, "ClctName")
                .Fields("PTCode").Value = GetFieldValue(rsArray, "PTCode")
                .Fields("PTName").Value = GetFieldValue(rsArray, "PTName")
                .Fields("RcdRowId").Value = GetFieldValue(rsArray, "RowId")
                .Fields("EntryNO").Value = intCount
                .Fields("StCode").Value = GetFieldValue(rsArray, "STCode")
                .Fields("STName").Value = GetFieldValue(rsArray, "STName")
                .Fields("CompCode").Value = GetFieldValue(rsArray, "CompCode")
                .Fields("ServiceType").Value = GetFieldValue(rsArray, "ServiceType")
                .Fields("ManualNo").Value = NoZero(rsArray("ManualNo").Value)
                .Fields("CancelFlag").Value = NoZero(rsArray("CancelFlag").Value, True)
                .Fields("CancelCode") = NoZero(rsArray("CancelCode"))
                .Fields("CancelName") = NoZero(rsArray("CancelName"))
                .Fields("BankCode") = NoZero(rsArray("BankCode"))
                .Fields("BankName") = NoZero(rsArray("BankName"))
                .Fields("AccountNo") = NoZero(rsArray("AccountNo"))
                .Fields("AuthorizeNo") = NoZero(rsArray("AuthorizeNo"))
                .Fields("AdjustFlag") = Val(rsArray("AdjustFlag") & "")
                .Fields("NextPeriod") = Val(rsArray("NextPeriod") & "")
                .Fields("NextAmt") = Val(rsArray("NextAmt") & "")
                '***************************************************************************
                '#3465　將設備流水號寫入SO074 By Kin 2007/08/28
                .Fields("FaciSeqNo") = rsArray("FaciSeqNo")
                '***************************************************************************
                
                '*****************************************************************
                '#3436 儲存使用發票的資訊 By Kin 2007/12/17
                If Not IsNull(rsArray("InvSeqNo")) Then
                    .Fields("InvSeqNo") = rsArray("InvSeqNo") & ""
                End If
                '*****************************************************************
                .Update
                rsArray.MoveNext
            Wend
        End With
        Call ClearRcd
        Call InitData
        txtBillNo.SetFocus
    Exit Sub
chkErr:
    ErrSub Me.Name, "cmdSave_Click"
End Sub

Private Sub Form_Activate()
    On Error Resume Next
        Screen.MousePointer = vbDefault
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
     If cmdSave.Enabled And KeyCode = vbKeyF2 Then Call cmdSave_Click: Exit Sub
     If KeyCode = vbKeyF3 Then Call cmdQuery_Click
End Sub

Private Sub Form_Load()
    On Error GoTo chkErr
        '初始設定
        Dim strPath As String
        strPath = ReadGICMIS1("TmpMDBPath")
        
        'cnMDB.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & IIf(Right(strPath, 1) = "\", strPath, strPath & "\") & "Tmp1111.MDB" & ";Persist Security Info=False"
        cnMDB.Open "Provider=" & GetOleDbProvider & ";Data Source=" & IIf(Right(strPath, 1) = "\", strPath, strPath & "\") & "Tmp1111.MDB" & ";Persist Security Info=False"
        
        If Not GetSo3311XTmpMDB(cnMDB) Then Exit Sub
        
        If Not GetRS(rsArray, "Select * from So3311", cnMDB, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
        
        If Alfa2 = False Then
            'Call GetGlobal
        End If
        Call subGrd
        Call subGil
        rsArrayNo.Open "Select * from so3311 Where 0=1", cnMDB, adOpenKeyset, adLockOptimistic
        Call InitData
        Call GetInitData
        Call ClearRcd
        
    Exit Sub
chkErr:
    ErrSub Me.Name, "Form_Load"
End Sub

Private Sub subGil()
    On Error Resume Next
        '公司別
        SetgiList gilCompCode, "CodeNo", "Description", "CD039", , , , , , , GetCompCodeFilter(gilCompCode)

End Sub

Private Sub InitData()
    On Error GoTo chkErr
    '至收費登錄暫存檔取前次登錄且存檔的資訊, 以便顯示<登錄單據數>, <登錄總金額>的內容: (註: <登錄總金額>格式為99, 999, 999)
    '‧  <登錄單據數> = “Select max(EntryNo) from SO074 where EntryEn=’<操作人員代號>’ ”
    '‧  若<登錄單據數>大於0, 則:
    '<登錄總金額> = “select sum(RealAmt) from SO074 where EntryEn=’<操作人員代號>’ and CancelFalg=0”
    '否則: <登錄總金額> = 0
    Dim rsTmp As New ADODB.Recordset
        gilCompCode.SetCodeNo garyGi(9)
        gilCompCode.Query_Description
        rsTmp.MaxRecords = 1
        If OpenRecordset(rsTmp, "Select Max(EntryNo) as BillCount From " & GetOwner & "So074 where EntryEn='" & garyGi(0) & "'", gcnGi, adOpenForwardOnly, adLockReadOnly, adUseClient, False, False) <> giOK Then MsgBox "連線失敗！請重新操作！", vbExclamation, "訊息！"
        If Not rsTmp.EOF Then
            If rsTmp("BillCount").Value > 0 Then
                lblBillCnt.Caption = rsTmp("BillCount").Value
                intCount = rsTmp("BillCount").Value
                If OpenRecordset(rsTmp, "Select Sum(RealAmt) as AmtCount From " & GetOwner & "So074 where EntryEn='" & garyGi(0) & "'", gcnGi, adOpenForwardOnly, adLockReadOnly, adUseClient, False, False) <> giOK Then MsgBox "連線失敗！請重新操作！", vbExclamation, "訊息！"
                lblTotalAmt.Caption = rsTmp("AmtCount").Value & ""
                GoTo CloseRS
            End If
        End If
        intDayCut = Val(GetSystemParaItem("DayCut", gilCompCode.GetCodeNo, "", "SO041") & "")
        lblBillCnt.Caption = 0
        lblTotalAmt.Caption = 0
CloseRS:
        rsTmp.Close
        Set rsTmp = Nothing
    Exit Sub
chkErr:
     Call ErrSub(Me.Name, "InitData")
End Sub

Private Function ChkDupData() As Boolean
    On Error GoTo chkErr
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    ChkDupData = True
        
        strSQL = "Select BillNo From " & GetOwner & "So074 Where " & IIf(Len(txtBillNo.Text) <> 12, "Billno='", "PrtSno='") & txtBillNo.Text & "'"
        If OpenRecordset(rsTmp, strSQL, gcnGi, adOpenForwardOnly, adLockReadOnly, adUseClient, False, False) <> giOK Then
                MsgBox "連線失敗！請重新操作！", vbExclamation, "訊息！"
                GoTo CloseRS
        End If
        If rsTmp.EOF Then
            ChkDupData = False
        End If
CloseRS:
    rsTmp.Close
    Set rsTmp = Nothing
    Exit Function
chkErr:
    Call ErrSub(Me.Name, "ChkDupData")
End Function

Private Sub subGrd()
    On Error GoTo chkErr
    Dim mFlds As New prjGiGridR.GiGridFlds
        strSQL = "Select RowId, So033.* "
        With mFlds
            .Add "Citemname", , , , False, "    收費項目    ", vbLeftJustify
            .Add "ShouldAmt", , , , False, "出單金額", vbRightJustify
            .Add "RealPeriod", , , , False, "期數", vbLeftJustify
            .Add "RealAmt", , , , False, "實收金額", vbRightJustify
            .Add "RealStartDate", giControlTypeDate, , , False, "期限起始", vbLeftJustify
            .Add "RealStopDate", giControlTypeDate, , , False, "期限截止", vbLeftJustify
            .Add "CancelFlag", , , , False, "作廢", vbLeftJustify
            .Add "Note", , , , False, Space(20) & "         備註            ", vbLeftJustify
        End With
            
        ggrData.AllFields = mFlds
        ggrData.SetHead
            
    Exit Sub
chkErr:
    Call ErrSub(Me.Name, "subGrd")
End Sub

Private Sub ClearRcd(Optional blnOnlyClearRS As Boolean)
    On Error GoTo chkErr
        If Not blnOnlyClearRS Then
            txtBillNo.Text = ""
        End If
        lblName.Caption = ""    '姓名
        lblCustId.Caption = ""  '客戶編號
        lblTime.Caption = ""    '派工類別
        lblClass.Caption = ""   '完工時間
        lblSubTotal.Caption = ""    '應收小計
        lblSignDate.Caption = ""
        '#3440 存檔時不要將入帳日期取消 By Kin 2007/08/17
        'gdaRealDate.Text = ""
        
        cnMDB.Execute "Delete * from So3311"
        If rsArray.State = adStateOpen Then rsArray.Close
        Set ggrData.Recordset = rsArrayNo
        ggrData.Refresh
        cmdDetAdd.Enabled = False
        cmdDetEdit.Enabled = False
        cmdDetDelete.Enabled = False
        cmdSave.Enabled = False

    Exit Sub
chkErr:
    ErrSub Me.Name, "ClearRcd"
'    cmdSave.Enabled = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
        frmSO331AG.Show vbModal
        If blnCanClose = False Then Cancel = True: Exit Sub
        intCount = 0
        If cnMDB.State = adStateOpen Then cnMDB.Close: Set cnMDB = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
        Call ReleaseCOM(Me)

End Sub

Private Sub gdaRealDate_Validate(Cancel As Boolean)
    On Error GoTo chkErr
        '#3411 如果未輸入則不檢查 By Kin 2008/06/02
        If gdaRealDate.GetValue = "" Then Exit Sub
        Cancel = Not ChkCloseDate(gdaRealDate.GetValue(True), gilCompCode.GetCodeNo, "")
        Exit Sub
'        'gdaRealDate.Text
'        '可空白！
'        If gdaRealDate.Text = "" Then Exit Sub
'        '若大於今日，則警告："此日期超過今天日期"
'        If Not IsDate(Format(gdaRealDate.GetValue, "####/##/##")) Then
'            MsgBox "日期錯誤！", , "訊息！"
'            Cancel = True
'            Exit Sub
'        End If
'        If CDate(gdaRealDate.Text) > RightDate Then MsgBox "此日期超過今天日期": Cancel = True: Exit Sub
'
'        '若（今日-本欄位值- >　＜Para6>)，則警告："此日期已超過系統設的安全期限"
'        If DateDiff("d", CDate(gdaRealDate.Text), RightDate) > XintPara6 Then
'            MsgBox "此日期已超過系統設定的安全期限", vbExclamation, "訊息！"
'            Cancel = True
'            Exit Sub
'        End If
'
'        '若小於<TransDate>，則錯誤："<TransDate>之前已做過日結，不可使用之前日期入帳",Focus停留在本欄位！
'        If DateDiff("d", CDate(XstrTranDate), CDate(gdaRealDate.Text)) <= 0 Then
'            If intDayCut = 1 And DateDiff("d", CDate(XstrTranDate), CDate(gdaRealDate.Text)) = 0 Then Exit Sub
'            MsgBox XstrTranDate & "之前已做過日結，不可使用之前日期入帳", vbExclamation, "訊息！"
'            Cancel = True
'            Exit Sub
'        End If
    Exit Sub
chkErr:
    ErrSub Me.Name, "gdaRealDate_Validate"
End Sub

'‧  若輸入值非B單格式(而是臨時單或印單序號格式), 則根據取出之收費資料第一筆的客戶編號, 至客戶基本資料檔取<客戶姓名>與<客戶狀態代碼/名稱>, 將<客戶姓名>與<客戶狀態名稱>顯示在畫面上
Private Function GetCustname(ByVal intCustid As Long, ByRef RetCustName, ByRef REtCustStatus)
    On Error GoTo chkErr
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
        strSQL = "Select CustName,CustStatusCode From " & GetOwner & "So001 Where Custid=" & intCustid
        If OpenRecordset(rsTmp, strSQL, gcnGi, adOpenForwardOnly, adLockReadOnly, adUseServer, False, True) <> giOK Then MsgBox "連線失敗！請重新操作！", vbExclamation, "訊息！": Exit Function
        If Not rsTmp.EOF Then
            RetCustName = rsTmp("CustName").Value
            REtCustStatus = rsTmp("CustStatusCode").Value
        End If
        rsTmp.Close
        Set rsTmp = Nothing
    Exit Function
chkErr:
    Call ErrSub(Me.Name, "GetCustId")
End Function

Private Sub ggrData_DblClick()
    On Error Resume Next
        If rsArray.State = adStateClosed Then Exit Sub
        If (rsArray.EOF Or rsArray.BOF) Then Exit Sub
        With frmSO331AF
            Set .uParentForm = Me
            Set .Conn = cnMDB
            .uDBType = giAccessDb
            .EditMode = giEditModeView
            .uBillNo = rsArray("BillNO").Value
            .uPrtSNo = rsArray("PrtSno").Value & ""
            .uItem = rsArray("Item").Value
            .uRowId = rsArray("RowId").Value
            .uClctEn = rsArray("ClctEn").Value & ""
            .uClctName = rsArray("ClctName").Value & ""
            .uPTCode = IIf(IsNull(rsArray("PTCode").Value), 0, rsArray("PTCode").Value)
            .uPTName = rsArray("PTName").Value & ""
            .uRealDate = rsArray("RealDate").Value & ""
            .uCustId = rsArray("CustID").Value
            .uCustName = rsArray("CustName").Value
            gServiceType = rsArray("ServiceType").Value
            .Show vbModal
        End With
End Sub

Private Sub ggrData_ShowCellData(ByVal Col As Long, ByVal Row As Long, ByVal giFld As prjGiGridR.GiGridFld, ByVal Fld As ADODB.Field, Value As Variant)
    On Error Resume Next
        Select Case UCase(Fld.Name)
            Case "CANCELFLAG"
                    Value = IIf(Value = 0, "否", "是")
            '#3465 測試不OK Grid時間起訖改為設定Grid的內定值(西元年)，不在此做轉換(自行發現) By Kin 2007/09/11
            'Case "REALSTARTDATE", "REALSTOPDATE"
            '        Value = Format(Value, "yyyy/mm/dd")
            '#3465 測試不OK　金額顯示原本為"###,###,###",會導致變成空值，改為###,###,###0，並設定常數 By Kin 2007/09/11
            Case "SHOULDAMT", "REALAMT"
                Value = Format(Value, strMoney)
        End Select
        
End Sub

Private Sub gilCompCode_Change()
    On Error Resume Next
        If Not ChgComp(gilCompCode, "SO3310", "SO331A") Then Exit Sub

End Sub

Private Sub txtBillNo_Change()
    On Error Resume Next
        '#3349 9碼模式輸入時，直接進入驗証模式 By Kin 2007/07/12
        If Len(txtBillNo) = 9 And chkNineCode Then txtBillNo.Tag = "": txtBillNo_Validate (False): Exit Sub
        If Len(txtBillNo) = 15 Then txtBillNo.Tag = "": txtBillNo_Validate (False)
End Sub

Private Sub txtBillNo_GotFocus()
    On Error Resume Next
        txtBillNo.SelStart = 0
        txtBillNo.SelLength = Len(txtBillNo.Text)
        txtBillNo.Tag = ""
End Sub

Private Sub txtBillNo_KeyPress(keyAscii As Integer)
    On Error Resume Next
        keyAscii = Asc(UCase(Chr(keyAscii)))
'        If Len(txtBillNo) > 15 Then KeyAscii = 0: Exit Sub
'        If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 73 Or _
'                    KeyAscii = 105 Or KeyAscii = 77 Or KeyAscii = 109 Or KeyAscii = 80 Or KeyAscii = 112 Then
'                    KeyAscii = Asc(UCase(Chr(KeyAscii)))
'        Else
'                    KeyAscii = 9
'        End If
End Sub

Private Sub txtBillNo_Validate(Cancel As Boolean)
    On Error GoTo chkErr
    Dim lngBillNo As Long
    Dim strQuerySql As String
    Dim intCustStatus As Integer
    Dim intCustid As Long
    Dim strCustName As String
    Dim strClassName As String
    Dim strSNoType As String
    Dim lngStrStar As Long
    Dim lngStrEnd As Long
    Dim strChkRef3 As String
        If txtBillNo.Tag <> "" Then Exit Sub
        txtBillNo.Tag = "Do"
        If txtBillNo.Text = "" Then Exit Sub
        '#3349 取SNO長度，以便給SQL語法使用 By Kin 2007/07/12
        If chkNineCode Then
            lngStrStar = 7
            lngStrEnd = 9
        Else
            lngStrStar = 1
            lngStrEnd = 15
        End If
        '#3349 增加9碼模式，所以要判斷是9碼或15碼 By Kin 2007/07/12
        If chkNineCode Then
            If Len(txtBillNo) <> 9 Or InStr(1, "IMP", Left(txtBillNo, 1)) = 0 Then MsgBox "派工單號錯誤！,請重新輸入", vbExclamation, "訊息！": Cancel = True: Exit Sub
        Else
            If Len(txtBillNo.Text) <> 15 Or InStr(1, "IMP", Mid(txtBillNo, 7, 1)) = 0 Then MsgBox "派工單號錯誤！,請重新輸入", vbExclamation, "訊息！": Cancel = True: Exit Sub
        End If
        '#3349 指定SNO型態，以往是以Mid(txtBillNo, 7, 1)取值，增加9碼模式改以變數替代 By Kin 2007/07/12
        If chkNineCode Then
            strSNoType = Left(txtBillNo, 1)
        Else
            strSNoType = Mid(txtBillNo, 7, 1)
        End If
    '***********************************************************************************************************************************************************************************************************************************
    '#3122 為了要判斷簽收日期是否小於日結日期,必需將ServiceType挑選出來
    '#3349 增加9碼輸入模式，所以SQL語法全部增加Substr函數，值以lngStrStar與lngStrEnd來取 By Kin 2007/07/12
    '#3349 查詢SO007、SO008、SO009如果使用SubStr會導致效降低，所以使用原來15碼查詢的語法先保持不動(SO033使用Substr效率無影嚮) By Kin 2007/07/13
        Select Case strSNoType
            Case "I"
                If chkNineCode Then
                    strQuerySql = "Select Custid,CustName,InstName,FinTime,WorkerEn1,WorkerName1,SignDate,ServiceType From " & GetOwner & "So007 Where Substr(Sno," & lngStrStar & "," & lngStrEnd & ") ='" & txtBillNo.Text & "'"
                Else
                    strQuerySql = "Select Custid,CustName,InstName,FinTime,WorkerEn1,WorkerName1,SignDate,ServiceType From " & GetOwner & "So007 Where Sno ='" & txtBillNo.Text & "'"
                End If
            Case "M"
                If chkNineCode Then
                    strQuerySql = "Select Custid,CustName,ServiceName,FinTime,WorkerEn1,WorkerName1,SignDate,ServiceType From " & GetOwner & "So008 Where Substr(Sno," & lngStrStar & "," & lngStrEnd & ") ='" & txtBillNo.Text & "'"
                Else
                    strQuerySql = "Select Custid,CustName,ServiceName,FinTime,WorkerEn1,WorkerName1,SignDate,ServiceType From " & GetOwner & "So008 Where Sno ='" & txtBillNo.Text & "'"
                End If
            Case "P"
                If chkNineCode Then
                    strQuerySql = "Select Custid,CustName,PRName,FinTime,WorkerEn1,WorkerName1,SignDate,ServiceType From " & GetOwner & "So009 Where Substr(Sno," & lngStrStar & "," & lngStrEnd & ") ='" & txtBillNo.Text & "'"
                Else
                    strQuerySql = "Select Custid,CustName,PRName,FinTime,WorkerEn1,WorkerName1,SignDate,ServiceType From " & GetOwner & "So009 Where Sno='" & txtBillNo.Text & "'"
                End If
            Case Else
                MsgBox "派工單號錯誤！,請重新輸入", vbExclamation, "訊息！"
                txtBillNo.Text = ""
                Cancel = True
                Call ClearRcd(True)
                Exit Sub
        End Select
    '***********************************************************************************************************************************************************************************************************************************
        Dim rsTmp As New ADODB.Recordset
        '查檢是否已過入暫存檔so074
        Screen.MousePointer = vbHourglass
        If Not GetRS(rsTmp, strQuerySql, gcnGi) Then MsgBox "連線失敗！請重新操作！", vbExclamation, "訊息！": Screen.MousePointer = vbDefault: Exit Sub
        If rsTmp.EOF Then
                MsgBox "無此派工單號", vbExclamation, "訊息！"
                rsTmp.Close
                Cancel = True
                Call ClearRcd(True)
                Screen.MousePointer = vbDefault
                Exit Sub
        Else
            '‧  若'入帳日以簽收日為主'為checked, 且若取出之<簽收日期>無值,
            '則錯誤: "此派工單無簽收日期, 入帳日無法以簽收日為主", 並無以下動作  2000/04/19
    '        If chkDate.Value = 1 Then
    '            If (rsTmp("SignDate").Value & "") = "" Then MsgBox "此派工單無簽收日期, 入帳日無法以簽收日為主", vbExclamation, "訊息！": Exit Sub
    '        End If
            Dim rstmp1 As New ADODB.Recordset
            '#3349 9碼模式,SQL語法要改用SubStr By Kin 2007/07/12
            If Not GetRS(rstmp1, "Select Billno From " & GetOwner & "So074 Where SubStr(BillNo," & lngStrStar & "," & lngStrEnd & ") ='" & txtBillNo & "'", gcnGi, adUseClient, adOpenForwardOnly, adLockReadOnly) Then MsgBox "連線失敗！請重新操作！", vbExclamation, "訊息！": Cancel = True: Screen.MousePointer = vbDefault: Exit Sub
            If rstmp1.RecordCount > 0 Then
                MsgBox "本收費單已過入！請重新輸入！", vbExclamation, "訊息！"
                Cancel = True
                Call ClearRcd(True)
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            rstmp1.Close
            Set rstmp1 = Nothing
        End If
        If chkDate.Value = 1 Then
            If Len(rsTmp("SignDate").Value & "") = 0 Then
                chkDate.Value = False
                gdaRealDate.Enabled = True
                gdaRealDate.SetFocus
                MsgBox "此派工單無簽收日期，入帳日期無法以簽收日為主", vbExclamation, "訊息！"
                If gdaRealDate.Enabled Then gdaRealDate.SetFocus
                Call ClearRcd(True)
                Screen.MousePointer = vbDefault
                Exit Sub
            '#3122 如果以簽收日期為準，需再判斷簽收日期是否小於日結日期
            Else
                If Not ChkCloseDate(rsTmp("SignDate"), gilCompCode.GetCodeNo, rsTmp("ServiceType")) Then
                    chkDate.Value = False
                    gdaRealDate.Enabled = True
                    gdaRealDate.SetFocus
                    Call ClearRcd(True)
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
            End If
        Else
            If Len(gdaRealDate.GetValue) = 0 Then
                MsgBox "請輸入帳日期!!", vbExclamation, "訊息！"
                If gdaRealDate.Enabled Then gdaRealDate.SetFocus
                Call ClearRcd(True)
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        End If
        '#3349 以strSnoType 代替 Mid(txtBillNo, 7, 1)
        If strSNoType = "I" Then
            strClassName = rsTmp("INstName").Value
        ElseIf strSNoType = "M" Then
                strClassName = rsTmp("ServiceName").Value
            Else
                strClassName = rsTmp("PRName").Value
        End If
        
        intCustid = rsTmp("CustId").Value
        strCustName = rsTmp("CustName").Value
        '#3349 9碼模式輸入 SQL語法多增加SubStr By Kin 2007/07/12
        If Not GetRS(rsSo3311B, strSQL & "  From " & GetOwner & "So033 SO033 Where SubStr(BillNo," & lngStrStar & "," & lngStrEnd & ")='" & txtBillNo.Text & "' And RealDate is Null And UCCode is Not Null", gcnGi) Then MsgBox "連線失敗！請重新操作！", vbExclamation, "訊息！": Screen.MousePointer = vbDefault: Exit Sub
        
        If rsSo3311B.EOF Then
            MsgBox "無此單號或此工單已登錄過！", vbExclamation, "訊息！"
            rsSo3311B.Close
            Cancel = True
            Call ClearRcd
            Screen.MousePointer = vbDefault
            Exit Sub
        'Else
          '  cmdDetAdd.Enabled = False
         '   cmdDetDelete.Enabled = False
          '  cmdDetEdit.Enabled = False
        End If
        '*********************************************************************************************************
        '#4010 判斷UCCode的參考號是否為3 By Kin 2008/08/06
        strChkRef3 = "Select Count(*) Cnt From " & GetOwner & "SO033 A," & GetOwner & "CD013 B" & _
                    " Where SubStr(A.BillNo," & lngStrStar & "," & lngStrEnd & ")='" & txtBillNo.Text & "'" & _
                    " And A.RealDate is Null And A.UCCode is Not Null" & _
                    " And A.UCCode=B.CodeNo And B.RefNo=3 And A.CompCode=" & gCompCode
        If gcnGi.Execute(strChkRef3)(0) > 0 Then
            MsgBox "此單號或此工單櫃臺已收！", vbExclamation, "訊息！"
            Cancel = True
            Call ClearRcd
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        '*********************************************************************************************************
            cmdDetAdd.Enabled = Not rsTmp.EOF
            cmdDetDelete.Enabled = Not rsTmp.EOF
            cmdDetEdit.Enabled = Not rsTmp.EOF
            
       
        '檢查單據是否已登錄！
    '    If ChkDupData Then
    '        MsgBox "此單據已登錄過！請重新輸入！", vbExclamation, "訊息！"
    '        txtBillNo.Text = ""
    '        Cancel = True
    '        Exit Sub
    '    End If
    
        
    '將客戶編號、名稱、 示在螢幕上
        lblCustId.Caption = intCustid
        lblName.Caption = strCustName
        lblClass.Caption = strClassName
        '**************************************************************************
        '#3465 測試不OK，將Lable民國年改為西元年 By Kin 2007/09/11
        'lblTime.Caption = GetDTString(rsTmp("FinTime").Value & "")
        'lblSignDate.Caption = GetDT(rsTmp("SignDate").Value & "", GiDate)
        'lblTime.Caption = GetADdate(rsTmp("FinTime").Value & "")
        lblTime.Caption = Format(rsTmp("FinTime").Value & "", "YYYY/MM/DD HH:MM:SS")
        lblSignDate.Caption = GetADdate(rsTmp("SignDate").Value & "")
        '**************************************************************************
        If Not rsSo3311B.EOF Then
            rsSo3311B.MoveFirst
            'cnMdb.BeginTrans
            If rsArray.State = adStateOpen Then rsArray.Close
            cnMDB.Execute "Delete * from So3311"
            rsArray.CursorLocation = adUseClient
            rsArray.Open "Select * from So3311", cnMDB, adOpenKeyset, adLockOptimistic
            'lblSubTotal 應收小計
            On Error GoTo rsCancel
           While Not rsSo3311B.EOF
                With rsArray
                    .AddNew
                    .Fields("BillNo").Value = GetFieldValue(rsSo3311B, "BillNO")
                    .Fields("PrtsNo").Value = GetFieldValue(rsSo3311B, "PrtSNO")
                    .Fields("Item").Value = GetFieldValue(rsSo3311B, "Item")
                    .Fields("CustName").Value = strCustName
                    .Fields("CustId").Value = GetFieldValue(rsSo3311B, "Custid")
                    .Fields("CItemCode").Value = GetFieldValue(rsSo3311B, "CitemCode")
                    .Fields("CItemName").Value = GetFieldValue(rsSo3311B, "CitemName")
                    
                    .Fields("ShouldAmt").Value = NoZero(rsSo3311B("ShouldAmt").Value, True)
                    .Fields("RealAmt").Value = NoZero(rsSo3311B("ShouldAmt").Value, True)
                    .Fields("RealPeriod").Value = IIf(Len(rsSo3311B("RealPeriod").Value & "") = 0, NoZero(rsSo3311B("OldPeriod"), True), NoZero(rsSo3311B("RealPeriod"), True))
                    .Fields("CancelFlag").Value = NoZero(rsSo3311B("CancelFlag").Value, True)
                    
                    lblSubTotal.Caption = IIf(lblSubTotal.Caption = "", Val(lblSubTotal.Caption), lblSubTotal.Caption) + IIf(IsNull(rsSo3311B("ShouldAmt").Value), 0, rsSo3311B("ShouldAmt").Value)
                    .Fields("ShouldDate").Value = GetFieldValue(rsSo3311B, "ShouldDate")
                    
                    If Not (varType(rsSo3311B("RealStartDate")) = vbNull And varType(rsSo3311B("OldStartDate")) = vbNull) Then
                        .Fields("RealStartDate").Value = IIf(Len(rsSo3311B("RealStartDate").Value & "") = 0, Format(rsSo3311B("OldStartDate"), "YYYY/MM/DD"), Format(rsSo3311B("RealStartDate"), "YYYY/MM/DD"))
                        .Fields("RealStopDate").Value = IIf(Len(rsSo3311B("RealStopDate").Value & "") = 0, Format(rsSo3311B("OldStopDate"), "YYYY/MM/DD"), Format(rsSo3311B("RealStopDate"), "YYYY/MM/DD"))
                    End If
                    '若<入帳日以簽收日為主>為checked, 則實收日期為<簽收日期>, 否則為<入帳日期>
                    .Fields("RealDate").Value = IIf(chkDate.Value = 0, IIf(gdaRealDate.Text = "", Null, gdaRealDate.Text), rsTmp("SignDate").Value)
                    '實收金額 = 該筆之出單金額
                    
                    '.Fields("EntryNO").Value = intCount
                    .Fields("Note").Value = GetFieldValue(rsSo3311B, "Note")
                '*********************************************************************************
                    '#3349 收費人員直接代入工單WorkerEn1與WorkerName1 By Kin 2007/07/12
'                    If Len(rsSo3311B("ClctName").Value & "") > 0 Then
'                        .Fields("ClctEn").Value = GetFieldValue(rsSo3311B, "ClctEN")
'                        .Fields("ClctName").Value = GetFieldValue(rsSo3311B, "ClctName")
'                    Else
                        .Fields("ClctEn").Value = GetFieldValue(rsTmp, "WorkerEn1")
                        .Fields("ClctName").Value = GetFieldValue(rsTmp, "WorkerName1")
'                    End If
                '*********************************************************************************
                    .Fields("CmCode").Value = GetFieldValue(rsSo3311B, "CMCode")
                    .Fields("CMName").Value = GetFieldValue(rsSo3311B, "CMName")
                    .Fields("PTCode").Value = GetFieldValue(rsSo3311B, "PTCode")
                    .Fields("PTName").Value = GetFieldValue(rsSo3311B, "PTName")
                    .Fields("RowID").Value = GetFieldValue(rsSo3311B, "RowID")
                    strRowId = rsSo3311B("RowId").Value
                    .Fields("STCode").Value = GetFieldValue(rsSo3311B, "STCode")
                    .Fields("STName").Value = GetFieldValue(rsSo3311B, "STName")
                    .Fields("CompCode").Value = GetFieldValue(rsSo3311B, "CompCode")
                    .Fields("ManualNo").Value = GetFieldValue(rsSo3311B, "ManualNo")
                    .Fields("ServiceType").Value = GetFieldValue(rsSo3311B, "ServiceType")
                    .Fields("EntryEn").Value = garyGi(0)
                '***********************************************************************************
                    '#3349 收費資料之相關帳號資訊勿清除(原本是Mark) By Kin 2007/07/12
                    .Fields("BankCode") = rsSo3311B("BankCode")
                    .Fields("BankName") = rsSo3311B("BankName")
                    .Fields("AccountNo") = rsSo3311B("AccountNo")
'                    .Fields("AuthorizeNo") = rsSo3311B("AuthorizeNo")
'                    .Fields("AdjustFlag") = rsSo3311B("AdjustFlag")
                '***********************************************************************************
                '***********************************************************************************
                '#3465 將設備流水號寫入至Mdb By Kin 2007/08/28
                    .Fields("FaciSeqNo") = rsSo3311B("FaciSeqNo")
                    .Update
                '***********************************************************************************
                '***********************************************************************
                    '#3636 儲存所使用的發票資訊 By Kin 2007/12/17
                    .Fields("InvSeqNo") = rsSo3311B("InvSeqNO")
                '**********************************************************************
                End With
                
                    
                    
                rsSo3311B.MoveNext
           Wend
           On Error GoTo chkErr
           Call RefreshData
           cmdSave.Enabled = True
           Screen.MousePointer = vbDefault
        End If
    Exit Sub
rsCancel:
    rsSo3311B.CancelUpdate
    Screen.MousePointer = vbDefault
chkErr:
    Call ErrSub(Me.Name, "txtBillNo_Validate")
    Screen.MousePointer = vbDefault
    Exit Sub
lblRollback:
'    Resume 0
'    cnMDB.RollbackTrans
    Screen.MousePointer = vbDefault
    MsgBox "異動失敗！請重新操作！", vbExclamation, "訊息！"
End Sub

Private Sub TotalAmt()
    On Error Resume Next
    Dim rsTmp As New ADODB.Recordset
        If rsArray.RecordCount > 0 Then
            Set rsTmp = rsArray.Clone
            rsTmp.MoveFirst
            Dim lngSubTotal As Long
            While Not rsTmp.EOF
                lngSubTotal = lngSubTotal + IIf((rsTmp("RealAmt").Value & "") <> "", rsTmp("RealAmt").Value, 0)
                rsTmp.MoveNext
            Wend
            lblSubTotal.Caption = lngSubTotal
        Else
            lblSubTotal.Caption = 0
        End If
        rsTmp.Close
        Set rsTmp = Nothing
End Sub

Private Sub RefreshData()
    On Error GoTo chkErr
        If Not GetRS(rsArray, "Select * from So3311", cnMDB, adUseClient, adOpenForwardOnly, adLockOptimistic) Then Exit Sub
        Set ggrData.Recordset = rsArray
        ggrData.Refresh
        Call TotalAmt
    Exit Sub
chkErr:
    Call ErrSub(Me.Name, "RefreshData")
End Sub

'客戶編號
Public Property Get uCustId() As Long
    On Error Resume Next
        uCustId = intCustid
End Property

'客戶姓名
Public Property Get uCustName() As String
    On Error Resume Next
        uCustName = strCustName
End Property

'--------------------------------------------------------------------------------------------------------------------------------------------
'取得初始值！
Private Sub GetInitData()
    On Error GoTo chkErr
    Dim rsTmp As New ADODB.Recordset
        XstrTranDate = GetRsValue("Select tranDate from " & GetOwner & "So062 Where CompCode= " & gCompCode & " And Type=1") & ""
        If XstrTranDate = "" Then XstrTranDate = "1990/01/01"
        
        If Not GetSystemPara(rsTmp, gCompCode, gServiceType, "SO043", "Para5,Para6,Para8,Para9") Then Exit Sub
        If rsTmp.EOF = False Then
                XintPara5 = rsTmp("Para5").Value '延後幾日為次收日
                XintPara6 = rsTmp("Para6").Value '收費日登錄期限
                XintPara8 = rsTmp("Para8").Value '隨收費資料改變1
                XintPara9 = rsTmp("Para9").Value '隨收費資料改變2
        End If
        Call CloseRecordset(rsTmp)
    Exit Sub
chkErr:
    Call ErrSub(Me.Name, "GetInitData")
End Sub

Public Property Get uPara8() As Integer
    On Error Resume Next
        uPara8 = XintPara6
End Property

Public Property Let uCanClose(ByVal vData As Integer)
    On Error Resume Next
        blnCanClose = vData
End Property
'#3122 檢查簽收日期是否小於日結日期
Private Function ChkOverCloseDate(ByVal strCloseDate As String, _
    ByVal strCompCode As String, ByVal strServiceType As String) As Boolean
    On Error GoTo chkErr
    Dim intDayCut As Integer
    Dim strTranDate As String
    Dim intPara6 As Integer
        intDayCut = Val(GetSystemParaItem("DayCut", strCompCode, "", "SO041") & "")
        strTranDate = GetRsValue("Select tranDate from " & GetOwner & "So062 Where CompCode= " & strCompCode & " And Type=1 Order By TranDate Desc") & ""
        '收費參數檔（SO043)，取收費日登錄期限<Para6>，供後續使用！
        intPara6 = Val(GetSystemParaItem("Para6", strCompCode, strServiceType, "SO043") & "")
        If CDate(strCloseDate) > RightDate Then MsgBox "簽收期日期超過今天日期！", vbExclamation, "訊息！": Exit Function
        
        If (RightDate - CDate(strCloseDate)) > intPara6 Then
            MsgBox "簽收日期已超過系統設定的安全期限！", vbExclamation, "訊息！":  Exit Function
        End If
        If CDate(strCloseDate) <= IIf(strTranDate <> "", CDate(strTranDate), 0) Then
            If intDayCut = 1 And CDate(strCloseDate) = IIf(strTranDate <> "", CDate(strTranDate), 0) Then Exit Function
            MsgBox "簽收日期 " & strTranDate & " 之前已做過日結，不可使用之前日期入帳", vbExclamation, "訊息！"
            Exit Function
        End If
        ChkOverCloseDate = True
    Exit Function
chkErr:
    ErrSub Me.Name, "ChkOverCloseDate"
End Function


