VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSO5000A 
   BackColor       =   &H00FFFFFF&
   Caption         =   "報表管理[SO5000]"
   ClientHeight    =   8025
   ClientLeft      =   60
   ClientTop       =   1230
   ClientWidth     =   11880
   Icon            =   "SO5000A.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   8025
   ScaleWidth      =   11880
   StartUpPosition =   1  '所屬視窗中央
   Visible         =   0   'False
   WindowState     =   2  '最大化
   Begin MSComctlLib.StatusBar sbr1 
      Align           =   2  '對齊表單下方
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   7710
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
            Text            =   "報表管理[SO5000]"
            TextSave        =   "報表管理[SO5000]"
         EndProperty
      EndProperty
   End
   Begin VB.Image imgPic 
      Height          =   4665
      Left            =   0
      Top             =   0
      Width           =   6270
   End
   Begin VB.Menu mnuSO5100 
      Caption         =   "&1.人員績效"
      Begin VB.Menu mnuSO5110 
         Caption         =   "&1.電話業務績效統計"
      End
      Begin VB.Menu mnuSO5120 
         Caption         =   "&2.技術人員績效統計"
      End
      Begin VB.Menu mnuSO5130 
         Caption         =   "&3.服務區派工績效統計"
      End
      Begin VB.Menu mnuSO5140 
         Caption         =   "&4.收費人員狀況分析"
         Begin VB.Menu mnuSO5141 
            Caption         =   "&1.收費員出/回單查詢"
         End
         Begin VB.Menu mnuSO5142 
            Caption         =   "&2.收費員未收查詢"
         End
         Begin VB.Menu mnuSO5143 
            Caption         =   "&3.收費員未收費單查詢"
         End
         Begin VB.Menu mnuSO5144 
            Caption         =   "&4.收費人員績效統計"
         End
         Begin VB.Menu mnuSO5145 
            Caption         =   "&5.收費員回收率統計"
         End
      End
      Begin VB.Menu mnuSO5150 
         Caption         =   "&5.銷售點績效統計"
      End
      Begin VB.Menu mnuSO5160 
         Caption         =   "&6.業務人員績效統計"
      End
      Begin VB.Menu mnuSO5170 
         Caption         =   "&7.客戶介紹績效統計"
      End
      Begin VB.Menu mnuSO5180 
         Caption         =   "&8.Node 促銷業務報表"
      End
      Begin VB.Menu mnuSO5190 
         Caption         =   "&9.來電業務日報表"
      End
      Begin VB.Menu mnuSO51A0 
         Caption         =   "&A.來電業務期間報表"
      End
   End
   Begin VB.Menu mnuSO5200 
      Caption         =   "&2.收費資料"
      Begin VB.Menu mnuSO5210 
         Caption         =   "&1.每月未繳費客戶明細表"
      End
      Begin VB.Menu mnuSO5220 
         Caption         =   "&2.每月各項收費加總統計"
      End
      Begin VB.Menu mnuSO5230 
         Caption         =   "&3.每月應收與實收異常表"
      End
      Begin VB.Menu mnuSO5240 
         Caption         =   "&4.收費金額及期數異常表"
      End
      Begin VB.Menu mnuSO5250 
         Caption         =   "&5.收費報表查詢"
      End
      Begin VB.Menu mnuSO5260 
         Caption         =   "&6.預估未來一年收費客戶及金額統計"
      End
      Begin VB.Menu mnuSO5270 
         Caption         =   "&7.每月收費單完成率統計"
      End
      Begin VB.Menu mnuSO5280 
         Caption         =   "&8.收費到期客戶數統計表"
      End
      Begin VB.Menu mnuSO5290 
         Caption         =   "&9.每月每日收費單據數總表 "
      End
      Begin VB.Menu mnuSO52A0 
         Caption         =   "&A.每月每日應收單據金額統計"
      End
      Begin VB.Menu mnuSO52B0 
         Caption         =   "&B.收費金額資料統計"
      End
      Begin VB.Menu mnuSO52C0 
         Caption         =   "&C.每月每日實收月費分佈表"
      End
      Begin VB.Menu mnuSO52D0 
         Caption         =   "&D.應收帳明細表"
      End
      Begin VB.Menu mnuSO52E0 
         Caption         =   "&E.應收報表查詢"
      End
      Begin VB.Menu mnuSO52F0 
         Caption         =   "&F.收費方式日報表"
      End
      Begin VB.Menu mnuSO52G0 
         Caption         =   "&G.客戶繳款習性分析報表"
      End
      Begin VB.Menu mnuSO52H0 
         Caption         =   "&H.客戶分期收費繳款表"
      End
      Begin VB.Menu mnuSO52I0 
         Caption         =   "&I.產生未收費Call Out 電話"
      End
      Begin VB.Menu mnuSO52J0 
         Caption         =   "&J.IVR回單稽核報表"
      End
      Begin VB.Menu mnuSO52K0 
         Caption         =   "&K.預付制攤提報表"
      End
   End
   Begin VB.Menu mnuSO5300 
      Caption         =   "&3.收視頻道"
      Begin VB.Menu mnuSO5310 
         Caption         =   "&1.定址系統設定紀錄一覽表"
      End
      Begin VB.Menu mnuSO5320 
         Caption         =   "&2.鎖碼頻道戶數統計"
      End
      Begin VB.Menu mnuSO5330 
         Caption         =   "&3.鎖碼頻道客戶一覽表"
      End
      Begin VB.Menu mnuSO5340 
         Caption         =   "&4.解碼器使用紀錄一覽表"
      End
      Begin VB.Menu mnuSO5350 
         Caption         =   "&5.STB/ICC 設定記錄一覽表"
      End
   End
   Begin VB.Menu mnuSO5400 
      Caption         =   "&4.客戶設備"
      Begin VB.Menu mnuSO5410 
         Caption         =   "&1.客戶設備加裝一覽表"
      End
      Begin VB.Menu mnuSO5420 
         Caption         =   "&2.租借到期客戶一覽表"
      End
      Begin VB.Menu mnuSO5430 
         Caption         =   "&3.設備拆除客戶一覽表"
      End
      Begin VB.Menu mnuSO5440 
         Caption         =   "&4.客戶贈品兌換一覽表"
      End
      Begin VB.Menu mnuSO5450 
         Caption         =   "&5.設備異動表"
      End
      Begin VB.Menu mnuSO5460 
         Caption         =   "&6.統計超過特定台數設備"
      End
      Begin VB.Menu mnuSO5470 
         Caption         =   "&7.設備合約到期報表"
      End
      Begin VB.Menu mnuSO5480 
         Caption         =   "&8.設備取回稽核表"
      End
   End
   Begin VB.Menu mnuSO5500 
      Caption         =   "&5.客戶資料"
      Begin VB.Menu mnuSO5510 
         Caption         =   "&1.現有各類客戶數統計"
      End
      Begin VB.Menu mnuSO5520 
         Caption         =   "&2.各街道客戶狀態統計"
      End
      Begin VB.Menu mnuSO5530 
         Caption         =   "&3.客戶基本資料一覽表"
      End
      Begin VB.Menu mnuSO5540 
         Caption         =   "&4.註銷拆機客戶資料一覽表"
      End
      Begin VB.Menu mnuSO5550 
         Caption         =   "&5.新裝機客戶統計表"
      End
      Begin VB.Menu mnuSO5560 
         Caption         =   "&6.各服務區/行政區客戶增減數統計"
      End
      Begin VB.Menu mnuSO5570 
         Caption         =   "&7.無週期性收費項目客戶報表"
      End
      Begin VB.Menu mnuSO5580 
         Caption         =   "&8.大樓戶增減數統計"
      End
      Begin VB.Menu mnuSO5590 
         Caption         =   "&9.客戶信用卡資料統計表"
      End
      Begin VB.Menu mnuSO55A0 
         Caption         =   "&A.申請轉帳記錄報表"
      End
      Begin VB.Menu mnuSO55B0 
         Caption         =   "&B.證件繳付一覽表"
      End
      Begin VB.Menu mnuSO55C0 
         Caption         =   "&C.數位服務訂戶數報表"
      End
      Begin VB.Menu mnuSO55D0 
         Caption         =   "&D.全數位化稽核報表"
      End
   End
   Begin VB.Menu mnuSO5600 
      Caption         =   "&6.派工單"
      Begin VB.Menu mnuSO5610 
         Caption         =   "&1.裝機統計報表"
      End
      Begin VB.Menu mnuSO5620 
         Caption         =   "&2.維修統計報表"
      End
      Begin VB.Menu mnuSO5630 
         Caption         =   "&3.停拆移機統計報表"
      End
      Begin VB.Menu mnuSO5640 
         Caption         =   "&4.裝機未完成及退單統計"
      End
      Begin VB.Menu mnuSO5650 
         Caption         =   "&5.維修未完成及退單統計"
      End
      Begin VB.Menu mnuSO5660 
         Caption         =   "&6.停拆移機未完成及退單統計"
      End
      Begin VB.Menu mnuSO5670 
         Caption         =   "&7.各類派工單完工總表"
      End
      Begin VB.Menu mnuSO5680 
         Caption         =   "&8.每日派工單件數總表"
      End
      Begin VB.Menu mnuSO5690 
         Caption         =   "&9.裝機維修日報表"
      End
      Begin VB.Menu mnuSO56A0 
         Caption         =   "&A.派工單完成績效統計"
      End
      Begin VB.Menu mnuSO56B0 
         Caption         =   "&B.各行政區裝機類完工統計表"
      End
      Begin VB.Menu mnuSO56C0 
         Caption         =   "&C.停拆機復機比例分析"
      End
      Begin VB.Menu mnuSO56D0 
         Caption         =   "&D.取消拆機挽回報表"
      End
      Begin VB.Menu mnuSO56E0 
         Caption         =   "&E.各類派工單受理完工表"
      End
      Begin VB.Menu mnuSO56F0 
         Caption         =   "&F.裝機outbound報表"
      End
      Begin VB.Menu mnuSO56G0 
         Caption         =   "&G.維修outbound報表"
      End
      Begin VB.Menu mnuSO56H0 
         Caption         =   "&H.訂單統計報表"
      End
      Begin VB.Menu mnuSO56I0 
         Caption         =   "&I.停拆移機統計表(二)"
      End
      Begin VB.Menu mnuSO56J0 
         Caption         =   "&J.拆復機狀況抽查明細表"
      End
   End
   Begin VB.Menu mnuSO5700 
      Caption         =   "&7.各類原因"
      Begin VB.Menu mnuSO5710 
         Caption         =   "&1.裝機退單原因統計"
      End
      Begin VB.Menu mnuSO5720 
         Caption         =   "&2.維修退單原因統計"
      End
      Begin VB.Menu mnuSO5730 
         Caption         =   "&3.停拆移機退單原因統計"
      End
      Begin VB.Menu mnuSO5740 
         Caption         =   "&4.維修申告及故障原因統計"
      End
      Begin VB.Menu mnuSO5750 
         Caption         =   "&5.停拆機原因統計"
      End
      Begin VB.Menu mnuSO5760 
         Caption         =   "&6.未繳費原因統計"
      End
      Begin VB.Menu mnuSO5770 
         Caption         =   "&7.媒體介紹類別統計"
      End
      Begin VB.Menu mnuSO5780 
         Caption         =   "&8.客戶電話來電分類統計"
      End
      Begin VB.Menu mnuSO5790 
         Caption         =   "&9.客戶收費短收原因統計"
      End
      Begin VB.Menu mnuSO57A0 
         Caption         =   "&A.裝機退單追蹤表"
      End
      Begin VB.Menu mnuSO57B0 
         Caption         =   "&B.收費作廢原因統計表 "
      End
   End
   Begin VB.Menu mnuSO5800 
      Caption         =   "&8.客戶服務"
      Begin VB.Menu mnuSO5810 
         Caption         =   "&1.客戶派工紀錄統計"
      End
      Begin VB.Menu mnuSO5820 
         Caption         =   "&2.客戶維修紀錄統計"
      End
      Begin VB.Menu mnuSO5830 
         Caption         =   "&3.客戶裝機紀錄統計"
      End
      Begin VB.Menu mnuSO5840 
         Caption         =   "&4.客戶停拆移機紀錄統計"
      End
   End
   Begin VB.Menu mnuSO5900 
      Caption         =   "&9.郵遞標籤"
      Begin VB.Menu mnuSO5910 
         Caption         =   "&1.未繳費客戶標籤列印"
      End
      Begin VB.Menu mnuSO5920 
         Caption         =   "&2.現有客戶標籤列印"
      End
      Begin VB.Menu mnuSO5930 
         Caption         =   "&3.單一客戶標籤列印"
      End
      Begin VB.Menu mnuSO5940 
         Caption         =   "&4.已收費客戶標籤列印"
      End
      Begin VB.Menu mnuSO5950 
         Caption         =   "&5.銷售點標籤列印"
      End
      Begin VB.Menu mnuSO5960 
         Caption         =   "&6.客戶介紹客戶標籤列印"
      End
      Begin VB.Menu mnuSO5970 
         Caption         =   "&7.客戶促銷標籤列印"
      End
      Begin VB.Menu mnuSO5980 
         Caption         =   "&8.停拆移機標籤列印 "
      End
      Begin VB.Menu mnuSO5990 
         Caption         =   "&9.信用卡到期客戶列印"
      End
   End
   Begin VB.Menu mnuSO5A00 
      Caption         =   "&A.輔助分析"
      Begin VB.Menu mnuSO5A10 
         Caption         =   "&1.預收收入報表"
         Begin VB.Menu mnuSO5A11 
            Caption         =   "&1.預收收入報表(一般)"
         End
         Begin VB.Menu mnuSO5A12 
            Caption         =   "&2.預收收入報表(已日結)"
         End
      End
      Begin VB.Menu mnuSO5A20 
         Caption         =   "&2.每月裝機完工收費金額統計"
      End
      Begin VB.Menu mnuSO5A30 
         Caption         =   "&3.每月未繳費金額統計"
      End
      Begin VB.Menu mnuSO5A40 
         Caption         =   "&4.每月收費金額統計"
      End
      Begin VB.Menu mnuSO5A50 
         Caption         =   "&5.每月收費金額應收歸屬統計"
      End
      Begin VB.Menu mnuSO5A60 
         Caption         =   "&6.年度收入金額統計報表"
      End
      Begin VB.Menu mnuSO5A70 
         Caption         =   "&7.應收收入統計報表"
      End
      Begin VB.Menu mnuSO5A80 
         Caption         =   "&8.客戶匯率表"
      End
      Begin VB.Menu mnuSO5A90 
         Caption         =   "&9.已開發票未收明細表"
      End
      Begin VB.Menu mnuSO5AA0 
         Caption         =   "&A.訂單未派工明細表"
      End
      Begin VB.Menu mnuSO5AB0 
         Caption         =   "&B.版權成本分攤表"
      End
      Begin VB.Menu mnuSO5AC0 
         Caption         =   "&C.VOD 點數攤提報表"
      End
      Begin VB.Menu mnuSO5AD0 
         Caption         =   "&D.分攤調整數報表"
      End
   End
   Begin VB.Menu mnuSO5B00 
      Caption         =   "&B.資料檢核"
      Begin VB.Menu mnuSO5B10 
         Caption         =   "&1.異常代碼客戶明細表"
      End
      Begin VB.Menu mnuSO5B20 
         Caption         =   "&2.地址/電話重覆客戶明細表"
      End
      Begin VB.Menu mnuSO5B30 
         Caption         =   "&3.裝/拆機未開派工單客戶明細表"
      End
      Begin VB.Menu mnuSO5B40 
         Caption         =   "&4.收費期限(週期收費、收費資料)核對"
      End
      Begin VB.Menu mnuSO5B50 
         Caption         =   "&5.首次登錄時間異常報表 "
      End
      Begin VB.Menu mnuSO5B60 
         Caption         =   "&6.保證金未退費客戶一覽表"
      End
      Begin VB.Menu mnuSO5B70 
         Caption         =   "&7.收費資料異常報表"
      End
      Begin VB.Menu mnuSO5B80 
         Caption         =   "&8.週期性收費設定資料異動報表"
      End
      Begin VB.Menu mnuSO5B90 
         Caption         =   "&9.客戶資料異動報表"
      End
      Begin VB.Menu mnuSO5BA0 
         Caption         =   "&A.報表使用記錄資料檔"
      End
      Begin VB.Menu mnuSO5BB0 
         Caption         =   "&B.客戶統編異常資料報表 "
      End
      Begin VB.Menu mnuSO5BC0 
         Caption         =   "&C.收費資料異動一覽表"
      End
      Begin VB.Menu mnuSO5BD0 
         Caption         =   "&D.異動人員統計表"
      End
      Begin VB.Menu mnuSO5BE0 
         Caption         =   "&E.客戶狀態異常表"
      End
      Begin VB.Menu mnuSO5BF0 
         Caption         =   "&F.收費與頻道異常報表"
      End
      Begin VB.Menu mnuSO5BG0 
         Caption         =   "&G.客戶收費方式異常表"
      End
      Begin VB.Menu mnuSO5BH0 
         Caption         =   "&H.週期與未收費資料期限核對報表"
      End
      Begin VB.Menu mnuSO5BI0 
         Caption         =   "&I.已繳費未開頻道檢核報表"
      End
   End
   Begin VB.Menu mnuSO5C00 
      Caption         =   "&C.管理性報表"
      Begin VB.Menu mnuSO5C10 
         Caption         =   "&1.記點月報表"
      End
      Begin VB.Menu mnuSO5C20 
         Caption         =   "&2.客戶數統計月報表"
      End
      Begin VB.Menu mnuSO5C30 
         Caption         =   "&3.約當戶報表"
      End
      Begin VB.Menu mnuSO5C40 
         Caption         =   "&4.客戶所在Node分佈報表"
      End
      Begin VB.Menu mnuSO5C50 
         Caption         =   "&5.營運日報表"
      End
      Begin VB.Menu mnuSO5C60 
         Caption         =   "&6.工程日/週/月報表"
      End
      Begin VB.Menu mnuSO5C70 
         Caption         =   "&7.來電分類報表"
      End
      Begin VB.Menu mnuSO5C80 
         Caption         =   "&8.應收日分類報表"
      End
      Begin VB.Menu mnuSO5C90 
         Caption         =   "&9.預約派工彙總表"
      End
      Begin VB.Menu mnuSO5CA0 
         Caption         =   "&A.營運日報表"
      End
      Begin VB.Menu mnuSO5CB0 
         Caption         =   "&B.取消裝機挽回數統計報表"
      End
      Begin VB.Menu mnuSO5CC0 
         Caption         =   "&C.派工完工收費統計表 "
      End
      Begin VB.Menu mnuSO5CD0 
         Caption         =   "&D.取消拆機挽回報表"
      End
   End
   Begin VB.Menu mnuSO5D00 
      Caption         =   "&D.營運分析報表"
   End
   Begin VB.Menu mnuSO5E00 
      Caption         =   "&E.頻道商管理"
   End
   Begin VB.Menu mnuSO5F00 
      Caption         =   "&F.PPV相關報表 "
      Begin VB.Menu mnuSO5F10 
         Caption         =   "&1.IPPV/PPV點數訂購收費一覽表"
      End
      Begin VB.Menu mnuSO5F20 
         Caption         =   "&2.PPV節目訂購紀錄統計表 "
      End
      Begin VB.Menu mnuSO5F30 
         Caption         =   "&3.PPV節目訂購排行榜 "
      End
   End
   Begin VB.Menu mnuSO5G00 
      Caption         =   "&G.切換資料庫"
   End
   Begin VB.Menu mnuExit 
      Caption         =   "&0.結束"
   End
End
Attribute VB_Name = "frmSO5000A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim GB As Object
'
'Private Sub Form_Activate()
'  On Error Resume Next
'    Me.Move 0, 0
'    Debug.Print Err.Number
'    Me.WindowState = 2
'    Debug.Print Err.Number
'    If Not Me.Visible Then Me.Visible = True
'End Sub

Dim oldCompStr As String

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error GoTo ChkErr
   Call FunctionKey(KeyCode, Shift)
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "Form_KeyDown")
End Sub

Private Sub FunctionKey(KeyCode As Integer, Shift As Integer)
  On Error GoTo ChkErr
    Select Case KeyCode
           Case vbKeyEscape '   Esc
                If mnuExit.Enabled = True Then mnuExit_Click
    End Select
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "FunctionKey")
End Sub

Private Sub Form_Load()
  On Error Resume Next
    Dim lpdwResult As Long
    Me.Visible = False
'    Do While DoEvents()
'        If Not SendMessageTimeout(Me.hWnd, WM_NULL, 0&, 0&, SMTO_ABORTIFHUNG And SMTO_BLOCK, 15000, lpdwResult) Then
'            Kill5000
'        End If
'    Loop
'    Kill5000
'   Me.Visible = True
'    If App.PrevInstance Then
'        MsgBox "SO5000.EXE 存在於記憶體(行程)中 , 無法重覆執行 , 請用 Ctrl+Alt+Del 手動結束它 !!"
'        End
'    End If

    Set GB = CreateObject("GICMIS_BOJ.OBJ")
    Set GB.Form5000 = Me
    Me.Refresh
    Me.Show
    GB.FormGICMIS.Hide
    
  On Error GoTo ChkErr
  
    Correlate.CreateDicObj
    
    Call GetGlobal
    oldCompStr = garyGi(15)
    garyGi(15) = garyGi(9)
    Me.Caption = Me.Caption & garyGi(13)
    Call subSetPerv
    Set cnn = GetTmpMdbCn
'   If garyGi(4) = 0 Then Exit Sub
'    On Error Resume Next
'    Me.Move 0, 0
    If Me.WindowState = 1 Then Me.WindowState = 2
    If Not Me.Visible Then Me.Visible = True
    lopd = True '2013.05.10 by Corey 法律保障個人資料 功能啟動
    lopd = Val(GetRsValue("SELECT NVL(MustreLogin,0) FROM " & GetOwner & "SO041 Order By SysId", gcnGi) & "") = 1
    LoadImage imgPic
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "Form_Load")
End Sub

'Private Sub Kill5000()
'  On Error Resume Next
'    If App.PrevInstance Then K5000: DoEvents
'    If App.PrevInstance Then K5000: DoEvents
'    If App.PrevInstance Then K5000: DoEvents
'    If App.PrevInstance Then
'        MsgBox "SO5000.EXE 存在於記憶體(行程)中 , 無法重覆執行 , 請用 Ctrl+Alt+Del 手動結束它 !!"
'        End
'    End If
'End Sub

'Private Sub Form_Unload(Cancel As Integer)
'  On Error Resume Next
'   Set GB.Form5000 = Nothing
'   GB.FormGicmis.Show
'   Set cnn = Nothing
'   End
'End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
    Dim hwnd As Long
    hwnd = FindWindow(vbNullString, "SO5D00")
    If hwnd <> 0 Then
        SetForegroundWindow hwnd
        PostMessage hwnd, &H10, 0&, 0&
        Delay 0.24
        hwnd = FindWindow(vbNullString, "SO5D00")
        If hwnd <> 0 Then
            MessageBox hwnd, "請先關閉操作視窗再離開系統!!", "提示", vbInformation
            Cancel = 1
            Exit Sub
        End If
    End If
    Set GB.Form5000 = Nothing
    DoEvents
    GB.FormGICMIS.Show
    If cnn.State <> 0 Then
      cnn.Close
      Set cnn = Nothing
    End If
    'ReleaseCOM frmSO5000A
    'DoEvents
    End
End Sub

Private Sub Form_Resize()
  On Error Resume Next
   imgPic.Move (Me.Width - imgPic.Width) / 2, (Me.Height - imgPic.Height) / 2
End Sub

Private Sub subSetPerv()
  On Error GoTo ChkErr
    'False  為未做之報表
    mnuSO5100.Enabled = GetUserPriv("SO5000", ("SO5100"))
    mnuSO5110.Enabled = GetUserPriv("SO5100", ("SO5110"))
    mnuSO5120.Enabled = GetUserPriv("SO5100", ("SO5120"))
    mnuSO5130.Enabled = GetUserPriv("SO5100", ("SO5130"))
    mnuSO5140.Enabled = GetUserPriv("SO5100", ("SO5140"))
    mnuSO5141.Enabled = GetUserPriv("SO5140", ("SO5141"))
    mnuSO5142.Enabled = GetUserPriv("SO5140", ("SO5142"))
    mnuSO5143.Enabled = GetUserPriv("SO5140", ("SO5143"))
    mnuSO5144.Enabled = GetUserPriv("SO5140", ("SO5144"))
    mnuSO5150.Visible = GetUserPriv("SO5100", ("SO5150"))
    mnuSO5160.Enabled = GetUserPriv("SO5100", ("SO5160"))
    mnuSO5170.Enabled = GetUserPriv("SO5100", ("SO5170"))
    mnuSO5180.Enabled = GetUserPriv("SO5100", ("SO5180"))
    mnuSO5190.Enabled = GetUserPriv("SO5100", ("SO5190"))
    mnuSO51A0.Enabled = GetUserPriv("SO5100", ("SO51A0"))
    
    mnuSO5200.Enabled = GetUserPriv("SO5000", ("SO5200"))
    mnuSO5210.Enabled = GetUserPriv("SO5200", ("SO5210"))
    mnuSO5220.Enabled = GetUserPriv("SO5200", ("SO5220"))
    mnuSO5230.Enabled = GetUserPriv("SO5200", ("SO5230"))
    mnuSO5240.Enabled = GetUserPriv("SO5200", ("SO5240"))
    mnuSO5250.Enabled = GetUserPriv("SO5200", ("SO5250"))
    mnuSO5260.Enabled = GetUserPriv("SO5200", ("SO5260"))
    mnuSO5270.Enabled = GetUserPriv("SO5200", ("SO5270"))
    mnuSO5280.Enabled = GetUserPriv("SO5200", ("SO5280"))
    mnuSO5290.Enabled = GetUserPriv("SO5200", ("SO5290"))
    mnuSO52A0.Enabled = GetUserPriv("SO5200", ("SO52A0"))
    mnuSO52B0.Enabled = GetUserPriv("SO5200", ("SO52B0"))
    mnuSO52C0.Enabled = GetUserPriv("SO5200", ("SO52C0"))
    mnuSO52D0.Enabled = GetUserPriv("SO5200", ("SO52D0"))
    mnuSO52E0.Enabled = GetUserPriv("SO5200", ("SO52E0"))
    mnuSO52F0.Enabled = GetUserPriv("SO5200", ("SO52F0"))
    mnuSO52G0.Enabled = GetUserPriv("SO5200", ("SO52G0"))
    mnuSO52H0.Enabled = GetUserPriv("SO5200", ("SO52H0"))
    mnuSO52I0.Enabled = GetUserPriv("SO5200", ("SO52I0"))
    mnuSO52J0.Enabled = GetUserPriv("SO5200", ("SO52J0"))
    mnuSO52K0.Enabled = GetUserPriv("SO5200", ("SO52K0"))
    
    mnuSO5300.Enabled = GetUserPriv("SO5000", ("SO5300"))
    mnuSO5310.Enabled = GetUserPriv("SO5300", ("SO5310"))
    mnuSO5320.Enabled = GetUserPriv("SO5300", ("SO5320"))
    mnuSO5330.Enabled = GetUserPriv("SO5300", ("SO5330"))
    mnuSO5340.Enabled = GetUserPriv("SO5300", ("SO5340"))
    mnuSO5350.Enabled = GetUserPriv("SO5300", ("SO5350"))
'    mnuSO5360.Enabled = GetUserPriv("SO5300", ("SO5360"))
'    mnuSO5370.Enabled = GetUserPriv("SO5300", ("SO5370"))
'    mnuSO5380.Enabled = GetUserPriv("SO5300", ("SO5380"))
'    mnuSO5390.Enabled = GetUserPriv("SO5300", ("SO5390"))
    
    mnuSO5400.Enabled = GetUserPriv("SO5000", ("SO5400"))
    mnuSO5410.Enabled = GetUserPriv("SO5400", ("SO5410"))
    mnuSO5420.Enabled = GetUserPriv("SO5400", ("SO5420"))
    mnuSO5430.Enabled = GetUserPriv("SO5400", ("SO5430"))
    mnuSO5440.Enabled = GetUserPriv("SO5400", ("SO5440"))
    mnuSO5450.Enabled = GetUserPriv("SO5400", ("SO5450"))
    mnuSO5460.Enabled = GetUserPriv("SO5400", ("SO5460"))
    mnuSO5470.Enabled = GetUserPriv("SO5400", ("SO5470"))
'    mnuSO5480.Enabled = GetUserPriv("SO5400", ("SO5480"))
'    mnuSO5490.Enabled = GetUserPriv("SO5400", ("SO5490"))
    
    mnuSO5500.Enabled = GetUserPriv("SO5000", ("SO5500"))
    mnuSO5510.Enabled = GetUserPriv("SO5500", ("SO5510"))
    mnuSO5520.Enabled = GetUserPriv("SO5500", ("SO5520"))
    mnuSO5530.Enabled = GetUserPriv("SO5500", ("SO5530"))
    mnuSO5540.Enabled = GetUserPriv("SO5500", ("SO5540"))
    mnuSO5550.Enabled = GetUserPriv("SO5500", ("SO5550"))
    mnuSO5560.Enabled = GetUserPriv("SO5500", ("SO5560"))
    mnuSO5570.Enabled = GetUserPriv("SO5500", ("SO5570"))
    mnuSO5580.Enabled = GetUserPriv("SO5500", ("SO5580"))
    mnuSO5590.Visible = GetUserPriv("SO5500", ("SO5590"))
    mnuSO55A0.Visible = GetUserPriv("SO5500", ("SO55A0"))
    mnuSO55B0.Visible = GetUserPriv("SO5500", ("SO55B0"))
    mnuSO55C0.Visible = GetUserPriv("SO5500", ("SO55C0"))
    
    mnuSO5600.Enabled = GetUserPriv("SO5000", ("SO5600"))
    mnuSO5610.Enabled = GetUserPriv("SO5600", ("SO5610"))
    mnuSO5620.Enabled = GetUserPriv("SO5600", ("SO5620"))
    mnuSO5630.Enabled = GetUserPriv("SO5600", ("SO5630"))
    mnuSO5640.Enabled = GetUserPriv("SO5600", ("SO5640"))
    mnuSO5650.Enabled = GetUserPriv("SO5600", ("SO5650"))
    mnuSO5660.Enabled = GetUserPriv("SO5600", ("SO5660"))
    mnuSO5670.Enabled = GetUserPriv("SO5600", ("SO5670"))
    mnuSO5680.Enabled = GetUserPriv("SO5600", ("SO5680"))
    mnuSO5690.Enabled = GetUserPriv("SO5600", ("SO5690"))
    mnuSO56A0.Enabled = GetUserPriv("SO5600", ("SO56A0"))
    mnuSO56B0.Enabled = GetUserPriv("SO5600", ("SO56B0"))
    mnuSO56C0.Enabled = GetUserPriv("SO5600", ("SO56C0"))
    mnuSO56D0.Enabled = GetUserPriv("SO5600", ("SO56D0"))
    mnuSO56E0.Enabled = GetUserPriv("SO5600", ("SO56E0"))
    mnuSO56F0.Enabled = GetUserPriv("SO5600", ("SO56F0"))
    mnuSO56G0.Enabled = GetUserPriv("SO5600", ("SO56G0"))
    mnuSO56H0.Enabled = GetUserPriv("SO5600", ("SO56H0"))
    mnuSO56I0.Enabled = GetUserPriv("SO5600", ("SO56I0"))
    mnuSO56J0.Enabled = GetUserPriv("SO5600", ("SO56J0"))
    
    mnuSO5700.Enabled = GetUserPriv("SO5000", ("SO5700"))
    mnuSO5710.Enabled = GetUserPriv("SO5700", ("SO5710"))
    mnuSO5720.Enabled = GetUserPriv("SO5700", ("SO5720"))
    mnuSO5730.Enabled = GetUserPriv("SO5700", ("SO5730"))
    mnuSO5740.Enabled = GetUserPriv("SO5700", ("SO5740"))
    mnuSO5750.Enabled = GetUserPriv("SO5700", ("SO5750"))
    mnuSO5760.Enabled = GetUserPriv("SO5700", ("SO5760"))
    mnuSO5770.Enabled = GetUserPriv("SO5700", ("SO5770"))
    mnuSO5780.Enabled = GetUserPriv("SO5700", ("SO5780"))
    mnuSO5790.Enabled = GetUserPriv("SO5700", ("SO5790"))
    mnuSO57A0.Enabled = GetUserPriv("SO5700", ("SO57A0"))
    mnuSO57B0.Enabled = GetUserPriv("SO5700", ("SO57B0"))
    
    mnuSO5800.Enabled = GetUserPriv("SO5000", ("SO5800"))
    mnuSO5810.Enabled = GetUserPriv("SO5800", ("SO5810"))
    mnuSO5820.Enabled = GetUserPriv("SO5800", ("SO5820"))
    mnuSO5830.Enabled = GetUserPriv("SO5800", ("SO5830"))
    mnuSO5840.Enabled = GetUserPriv("SO5800", ("SO5840"))
'    mnuSO5850.Enabled = GetUserPriv("SO5800", ("SO5850"))
'    mnuSO5860.Enabled = GetUserPriv("SO5800", ("SO5860"))
'    mnuSO5870.Enabled = GetUserPriv("SO5800", ("SO5870"))
'    mnuSO5880.Enabled = GetUserPriv("SO5800", ("SO5880"))
'    mnuSO5890.Enabled = GetUserPriv("SO5800", ("SO5890"))
    
    mnuSO5900.Enabled = GetUserPriv("SO5000", ("SO5900"))
    mnuSO5910.Enabled = GetUserPriv("SO5900", ("SO5910"))
    mnuSO5920.Enabled = GetUserPriv("SO5900", ("SO5920"))
    mnuSO5930.Enabled = GetUserPriv("SO5900", ("SO5930"))
    mnuSO5940.Enabled = GetUserPriv("SO5900", ("SO5940"))
    mnuSO5950.Enabled = GetUserPriv("SO5900", ("SO5950"))
    mnuSO5960.Enabled = GetUserPriv("SO5900", ("SO5960"))
    mnuSO5970.Enabled = GetUserPriv("SO5900", ("SO5970"))
    mnuSO5980.Enabled = GetUserPriv("SO5900", ("SO5980"))
    mnuSO5990.Enabled = GetUserPriv("SO5900", ("SO5990"))
    
    mnuSO5A00.Enabled = GetUserPriv("SO5000", ("SO5A00"))
    mnuSO5A10.Enabled = GetUserPriv("SO5A00", ("SO5A10"))
    mnuSO5A11.Enabled = GetUserPriv("SO5A00", ("SO5A11"))
    
    '#7256 don't need to check CloseProAmortize field of so041 by Kin 2016/07/19
    mnuSO5A12.Enabled = GetUserPriv("SO5A00", ("SO5A12"))
    mnuSO5A20.Enabled = GetUserPriv("SO5A00", ("SO5A20"))
    mnuSO5A30.Enabled = GetUserPriv("SO5A00", ("SO5A30"))
    mnuSO5A40.Enabled = GetUserPriv("SO5A00", ("SO5A40"))
    mnuSO5A50.Enabled = GetUserPriv("SO5A00", ("SO5A50"))
    mnuSO5A60.Enabled = GetUserPriv("SO5A00", ("SO5A60"))
    mnuSO5A70.Visible = GetUserPriv("SO5A00", ("SO5A70")) And False
    mnuSO5A80.Enabled = GetUserPriv("SO5A00", ("SO5A80"))
    mnuSO5A90.Visible = GetUserPriv("SO5A00", ("SO5A90"))
    mnuSO5AA0.Visible = GetUserPriv("SO5A00", ("SO5AA0"))
    mnuSO5AB0.Visible = GetUserPriv("SO5A00", ("SO5AB0"))
    '#5448 增加點數攤提報表
    mnuSO5AC0.Visible = GetUserPriv("SO5A00", ("SO5AC0"))
    
    mnuSO5B00.Enabled = GetUserPriv("SO5000", ("SO5B00"))
    mnuSO5B10.Visible = GetUserPriv("SO5B00", ("SO5B10"))
    mnuSO5B20.Enabled = GetUserPriv("SO5B00", ("SO5B20"))
    mnuSO5B30.Enabled = GetUserPriv("SO5B00", ("SO5B30"))
    mnuSO5B40.Visible = GetUserPriv("SO5B00", ("SO5B40"))
    mnuSO5B50.Enabled = GetUserPriv("SO5B00", ("SO5B50"))
    mnuSO5B60.Enabled = GetUserPriv("SO5B00", ("SO5B60"))
    mnuSO5B70.Enabled = GetUserPriv("SO5B00", ("SO5B70"))
    mnuSO5B80.Enabled = GetUserPriv("SO5B00", ("SO5B80"))
    mnuSO5B90.Enabled = GetUserPriv("SO5B00", ("SO5B90"))
    mnuSO5BA0.Enabled = GetUserPriv("SO5B00", ("SO5BA0"))
    mnuSO5BB0.Enabled = GetUserPriv("SO5B00", ("SO5BB0"))
    mnuSO5BC0.Enabled = GetUserPriv("SO5B00", ("SO5BC0"))
    mnuSO5BD0.Enabled = GetUserPriv("SO5B00", ("SO5BD0"))
    mnuSO5BE0.Enabled = GetUserPriv("SO5B00", ("SO5BE0"))
    mnuSO5BF0.Enabled = GetUserPriv("SO5B00", ("SO5BF0"))
    mnuSO5BG0.Enabled = GetUserPriv("SO5B00", ("SO5BG0"))
    mnuSO5BH0.Enabled = GetUserPriv("SO5B00", ("SO5BH0"))
    mnuSO5BI0.Enabled = GetUserPriv("SO5B00", ("SO5BI0"))
    
    mnuSO5C00.Enabled = GetUserPriv("SO5000", ("SO5C00"))
    mnuSO5C10.Enabled = GetUserPriv("SO5C00", ("SO5C10"))
    mnuSO5C20.Enabled = GetUserPriv("SO5C00", ("SO5C20"))
    mnuSO5C30.Enabled = GetUserPriv("SO5C00", ("SO5C30"))
    mnuSO5C40.Enabled = GetUserPriv("SO5C00", ("SO5C40"))
    mnuSO5C50.Enabled = GetUserPriv("SO5C00", ("SO5C50"))
    mnuSO5C60.Enabled = GetUserPriv("SO5C00", ("SO5C60"))
    mnuSO5C70.Enabled = GetUserPriv("SO5C00", ("SO5C70"))
    mnuSO5C80.Enabled = GetUserPriv("SO5C00", ("SO5C80"))
    mnuSO5C90.Enabled = GetUserPriv("SO5C00", ("SO5C90"))
    mnuSO5CA0.Enabled = GetUserPriv("SO5C00", ("SO5CA0"))
    mnuSO5CB0.Enabled = GetUserPriv("SO5C00", ("SO5CB0"))
    mnuSO5CC0.Enabled = GetUserPriv("SO5C00", ("SO5CC0"))
    mnuSO5CD0.Enabled = GetUserPriv("SO5C00", ("SO5CD0"))
    
    mnuSO5D00.Enabled = GetUserPriv("SO5000", ("SO5D00"))
    mnuSO5E00.Enabled = GetUserPriv("SO5000", ("SO5E00"))
    mnuSO5G00.Enabled = GetUserPriv("SO8000", ("SO8200"))
    mnuSO5F00.Enabled = GetUserPriv("SO5000", ("SO5F00"))
    mnuSO5F10.Enabled = GetUserPriv("SO5F00", ("SO5F10"))
    mnuSO5F20.Enabled = GetUserPriv("SO5F00", ("SO5F20"))
    mnuSO5F30.Enabled = GetUserPriv("SO5F00", ("SO5F30"))
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "subSetPerv")
End Sub

Public Sub ReLoad()
  On Error GoTo ChkErr
    Form_Load
    garyGi(9) = gCompCode
    garyGi(14) = gCompCode
    garyGi(3) = gcnGi.ConnectionString
    lblObj = GetRsValue("SELECT DESCRIPTION FROM " & GetOwner & "CD052 WHERE CODENO=" & gCompCode, gcnGi)
    lblObj.Left = (Screen.ActiveForm.Width - lblObj.Width) / 2
    lblObj2 = lblObj
    lblObj2.Left = (Screen.ActiveForm.Width - lblObj.Width) / 2 - 20
    If App.Title <> "CSMIS" Then RefreshMenuForm
  Exit Sub
ChkErr:
    ErrSub Me.Name, "ReLoad"
End Sub

Private Sub mnuExit_Click()
  On Error Resume Next
   Unload Me
End Sub

Private Sub mnuSO5110_Click()
  On Error Resume Next
    frmSO5110A.Show vbModal
End Sub

Private Sub mnuSO5120_Click()
  On Error Resume Next
    frmSO5120A.Show vbModal
End Sub

Private Sub mnuSO5130_Click()
  On Error Resume Next
    frmSO5130A.Show vbModal
End Sub

Private Sub mnuSO5141_Click()
  On Error Resume Next
    frmSO5141A.Show vbModal
End Sub

Private Sub mnuSO5142_Click()
  On Error Resume Next
    frmSO5142A.Show vbModal
End Sub

Private Sub mnuSO5143_Click()
  On Error Resume Next
    frmSO5143A.Show vbModal
End Sub

Private Sub mnuSO5144_Click()
  On Error Resume Next
   frmSO5144A.Show vbModal
End Sub

Private Sub mnuSO5145_Click()
  On Error Resume Next
   frmSO5145A.Show vbModal
End Sub

Private Sub mnuSO5150_Click()
  On Error Resume Next
   frmSO5150A.Show vbModal
End Sub

Private Sub mnuSO5160_Click()
   On Error Resume Next
   frmSO5160A.Show vbModal
End Sub

Private Sub mnuSO5170_Click()
  On Error Resume Next
   frmSO5170A.Show vbModal
End Sub

Private Sub mnuSO5180_Click()
  On Error Resume Next
   frmSO5180A.Show vbModal
End Sub

Private Sub mnuSO5190_Click()
  On Error Resume Next
   frmSO5190A.Show vbModal
End Sub

Private Sub mnuSO51A0_Click()
  On Error Resume Next
   frmSO51A0A.Show vbModal
End Sub

Private Sub mnuSO5210_Click()
  On Error Resume Next
   frmSO5210A.Show vbModal
End Sub

Private Sub mnuSO5220_Click()
  On Error Resume Next
   frmSO5220A.Show vbModal
End Sub

Private Sub mnuSO5230_Click()
  On Error Resume Next
   frmSO5230A.Show vbModal
End Sub

Private Sub mnuSO5240_Click()
  On Error Resume Next
   frmSO5240A.Show vbModal
End Sub

Private Sub mnuSO5250_Click()
  On Error GoTo ChkErr
    With frmSO3320A
      .uForm = "SO5250A"
      .Show vbModal
    End With
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "cmdExit_Click")
End Sub

Private Sub mnuSO5260_Click()
  On Error Resume Next
   frmSO5260A.Show vbModal
End Sub

Private Sub mnuSO5270_Click()
  On Error Resume Next
   frmSO5270A.Show vbModal
End Sub

Private Sub mnuSO5280_Click()
  On Error Resume Next
   frmSO5280A.Show vbModal
End Sub

Private Sub mnuSO5290_Click()
  On Error Resume Next
   frmSo5290A.Show vbModal
End Sub

Private Sub mnuSO52A0_Click()
  On Error Resume Next
   frmSO52A0A.Show vbModal
End Sub

Private Sub mnuSO52B0_Click()
  On Error Resume Next
   frmSO52B0A.Show vbModal
End Sub

Private Sub mnuSO52C0_Click()
  On Error Resume Next
   frmSO52C0A.Show vbModal
End Sub

Private Sub mnuSO52D0_Click()
  On Error Resume Next
   frmSO52D0A.Show vbModal
End Sub

Private Sub mnuSO52E0_Click()
  On Error Resume Next
   frmSO52E0A.Show vbModal
End Sub

Private Sub mnuSO52F0_Click()
  On Error Resume Next
   frmSO52F0A.Show vbModal
End Sub

Private Sub mnuSO52G0_Click()
  On Error Resume Next
   frmSO52G0A.Show vbModal
End Sub

Private Sub mnuSO52H0_Click()
  On Error Resume Next
   frmSO52H0A.Show vbModal
End Sub

Private Sub mnuSO52I0_Click()
  On Error Resume Next
   frmSO52I0A.Show vbModal
End Sub

Private Sub mnuSO52J0_Click()
  On Error Resume Next
    frmSO52J0A.Show vbModal
End Sub

Private Sub mnuSO52K0_Click()
  frmSO52K0A.Show vbModal
End Sub

Private Sub mnuSO5310_Click()
  On Error Resume Next
   frmSO5310A.Show vbModal
End Sub

Private Sub mnuSO5320_Click()
  On Error Resume Next
   frmSO5320A.Show vbModal
End Sub

Private Sub mnuSO5330_Click()
  On Error Resume Next
   frmSO5330A.Show vbModal
End Sub

Private Sub mnuSO5340_Click()
  On Error Resume Next
   frmSO5340A.Show vbModal
End Sub

Private Sub mnuSO5350_Click()
  On Error Resume Next
   frmSO5350A.Show vbModal
End Sub

Private Sub mnuSO5410_Click()
  On Error Resume Next
   frmSO5410A.Show vbModal
End Sub

Private Sub mnuSO5420_Click()
  On Error Resume Next
   frmSO5420A.Show vbModal
End Sub

Private Sub mnuSO5430_Click()
  On Error Resume Next
   frmSO5430A.Show vbModal
End Sub

Private Sub mnuSO5440_Click()
  On Error Resume Next
   frmSO5440A.Show vbModal
End Sub

Private Sub mnuSO5450_Click()
  On Error Resume Next
   frmSO5450A.Show vbModal
End Sub
Private Sub mnuSO5460_Click()
  On Error Resume Next
  frmSO5460A.Show vbModal
End Sub
Private Sub mnuSO5470_Click()
  On Error Resume Next
  frmSO5470A.Show vbModal
End Sub

Private Sub mnuSO5480_Click()
  On Error Resume Next
    frmSO5480A.Show vbModal
End Sub

Private Sub mnuSO5510_Click()
  On Error Resume Next
   frmSO5510A.Show vbModal
End Sub

Private Sub mnuSO5520_Click()
  On Error Resume Next
   frmSO5520A.Show vbModal
End Sub

Private Sub mnuSO5530_Click()
  On Error Resume Next
   frmSO5530A.Show vbModal
End Sub

Private Sub mnuSO5540_Click()
  On Error Resume Next
   frmSO5540A.Show vbModal
End Sub

Private Sub mnuSO5550_Click()
  On Error Resume Next
   frmSO5550A.Show vbModal
End Sub

Private Sub mnuSO5560_Click()
  On Error Resume Next
   frmSO5560A.Show vbModal
End Sub

Private Sub mnuSO5570_Click()
  On Error Resume Next
   frmSO5570A.Show vbModal
End Sub

Private Sub mnuSO5580_Click()
  On Error Resume Next
   frmSO5580A.Show vbModal
End Sub

Private Sub mnuSO5590_Click()
  On Error Resume Next
   frmSO5590A.Show vbModal

End Sub

Private Sub mnuSO55A0_Click()
  On Error Resume Next
   frmSO55A0A.Show vbModal
End Sub

Private Sub mnuSO55B0_Click()
  On Error Resume Next
   frmSO55B0A.Show vbModal

End Sub

Private Sub mnuSO55C0_Click()
   frmSO55C0A.Show vbModal
End Sub

Private Sub mnuSO55D0_Click()
  On Error Resume Next
    frmSO55D0A.Show vbModal
End Sub

Private Sub mnuSO5610_Click()
  On Error Resume Next
   frmSO5610A.Show vbModal
End Sub

Private Sub mnuSO5620_Click()
  On Error Resume Next
   frmSO5620A.Show vbModal
End Sub

Private Sub mnuSO5630_Click()
  On Error Resume Next
   frmSO5630A.Show vbModal
End Sub

Private Sub mnuSO5640_Click()
  On Error Resume Next
   frmSO5640A.Show vbModal
End Sub

Private Sub mnuSO5650_Click()
  On Error Resume Next
   frmSO5650A.Show vbModal
End Sub

Private Sub mnuSO5660_Click()
  On Error Resume Next
   frmSO5660A.Show vbModal
End Sub

Private Sub mnuSO5670_Click()
  On Error Resume Next
   frmSO5670A.Show vbModal
End Sub

Private Sub mnuSO5680_Click()
  On Error Resume Next
   frmSO5680A.Show vbModal
End Sub

Private Sub mnuSO5690_Click()
  On Error Resume Next
   frmSO5690A.Show vbModal
End Sub

Private Sub mnuSO56A0_Click()
  On Error Resume Next
   frmSO56A0A.Show vbModal
End Sub

Private Sub mnuSO56B0_Click()
  On Error Resume Next
   frmSO56B0A.Show vbModal
End Sub

Private Sub mnuSO56C0_Click()
  On Error Resume Next
   frmSO56C0A.Show vbModal
End Sub

Private Sub mnuSO56D0_Click()
  On Error Resume Next
   frmSO56D0A.Show vbModal
End Sub

Private Sub mnuSO56E0_Click()
  On Error Resume Next
   frmSO56E0A.Show vbModal
End Sub

Private Sub mnuSO56F0_Click()
  On Error Resume Next
   frmSO56F0A.Show vbModal
End Sub

Private Sub mnuSO56G0_Click()
  On Error Resume Next
   frmSO56G0A.Show vbModal
End Sub

Private Sub mnuSO56H0_Click()
  On Error Resume Next
   frmSO56H0A.Show vbModal
End Sub

Private Sub mnuSO56I0_Click()
  On Error Resume Next
   frmSO56I0A.Show vbModal
End Sub
Private Sub mnuSO56J0_Click()
  On Error Resume Next
   frmSO56J0A.Show vbModal
End Sub

Private Sub mnuSO5710_Click()
  On Error Resume Next
   frmSO5710A.Show vbModal
End Sub

Private Sub mnuSO5720_Click()
  On Error Resume Next
   frmSO5720A.Show vbModal
End Sub

Private Sub mnuSO5730_Click()
  On Error Resume Next
   frmSO5730A.Show vbModal
End Sub

Private Sub mnuSO5740_Click()
  On Error Resume Next
   frmSO5740A.Show vbModal
End Sub

Private Sub mnuSO5750_Click()
  On Error Resume Next
   frmSO5750A.Show vbModal
End Sub

Private Sub mnuSO5760_Click()
  On Error Resume Next
   frmSO5760A.Show vbModal
End Sub

Private Sub mnuSO5770_Click()
  On Error Resume Next
   frmSO5770A.Show vbModal
End Sub

Private Sub mnuSO5780_Click()
  On Error Resume Next
   frmSO5780A.Show vbModal
End Sub

Private Sub mnuSO5790_Click()
  On Error Resume Next
   frmSO5790A.Show vbModal
End Sub

Private Sub mnuSO57A0_Click()
  On Error Resume Next
   frmSO57A0A.Show vbModal
End Sub

Private Sub mnuSO57B0_Click()
  On Error Resume Next
   frmSo57B0A.Show vbModal
End Sub

Private Sub mnuSO5810_Click()
  On Error Resume Next
   frmSO5810A.Show vbModal
End Sub

Private Sub mnuSO5820_Click()
  On Error Resume Next
   frmSO5820A.Show vbModal
End Sub

Private Sub mnuSO5830_Click()
  On Error Resume Next
   frmSO5830A.Show vbModal
End Sub

Private Sub mnuSO5840_Click()
  On Error Resume Next
   frmSO5840A.Show vbModal
End Sub

Private Sub mnuSO5910_Click()
  On Error Resume Next
   frmSO5910A.Show vbModal
End Sub

Private Sub mnuSO5920_Click()
  On Error Resume Next
   frmSO5920A.Show vbModal
End Sub

Private Sub mnuSO5930_Click()
  On Error Resume Next
   frmSO5930A.Show vbModal
End Sub

Private Sub mnuSO5940_Click()
  On Error Resume Next
   frmSO5940A.Show vbModal
End Sub

Private Sub mnuSO5950_Click()
  On Error Resume Next
   frmSO5950A.Show vbModal
End Sub

Private Sub mnuSO5960_Click()
  On Error Resume Next
   frmSO5960A.Show vbModal
End Sub

Private Sub mnuSO5970_Click()
  On Error Resume Next
   frmSO5970A.Show vbModal
End Sub

Private Sub mnuSO5980_Click()
  On Error Resume Next
   frmSO5980A.Show vbModal
End Sub

Private Sub mnuSO5990_Click()
  On Error Resume Next
   frmSO5990A.Show vbModal
End Sub

Private Sub mnuSO5A11_Click()
On Error Resume Next
   frmSO5A10A.Show vbModal
End Sub

Private Sub mnuSO5A12_Click()
  On Error Resume Next
    frmSO5A12A.Show vbModal
End Sub

Private Sub mnuSO5A20_Click()
  On Error Resume Next
   frmSo5A20A.Show vbModal
End Sub

Private Sub mnuSO5A30_Click()
  On Error Resume Next
   frmSo5A30A.Show vbModal
End Sub

Private Sub mnuSO5A40_Click()
  On Error Resume Next
   frmSo5A40A.Show vbModal
End Sub

Private Sub mnuSO5A50_Click()
  On Error Resume Next
   frmSo5A50A.Show vbModal
End Sub

Private Sub mnuSO5A60_Click()
  On Error Resume Next
   frmSo5A60A.Show vbModal
End Sub

'Private Sub mnuSO5A70_Click()
'  On Error Resume Next
'   frmSO5A70A.Show vbModal
'End Sub

Private Sub mnuSO5A80_Click()
  On Error Resume Next
   frmSO5A80A.Show vbModal
End Sub

Private Sub mnuSO5A90_Click()
  On Error Resume Next
   frmSO5A90A.Show vbModal
End Sub

Private Sub mnuSO5AA0_Click()
  On Error Resume Next
   frmSO5AA0A.Show vbModal
End Sub

Private Sub mnuSO5AB0_Click()
  On Error Resume Next
   frmSO5AB0A.Show vbModal
End Sub

Private Sub mnuSO5AC0_Click()
  On Error Resume Next
   frmSO5AC0A.Show vbModal
End Sub

Private Sub mnuSO5AD0_Click()
  On Error Resume Next
   frmSO5AD0A.Show vbModal
End Sub

Private Sub mnuSO5B10_Click()
  On Error Resume Next
   frmSO5B10A.Show vbModal
End Sub

Private Sub mnuSO5B20_Click()
  On Error Resume Next
   frmSO5B20A.Show vbModal
End Sub

Private Sub mnuSO5B30_Click()
  On Error Resume Next
   frmSO5B30A.Show vbModal
End Sub

Private Sub mnuSO5B40_Click()
  On Error Resume Next
   frmSO5B40A.Show vbModal
End Sub

Private Sub mnuSO5B50_Click()
  On Error Resume Next
   frmSO5B50A.Show vbModal
End Sub

Private Sub mnuSO5B60_Click()
  On Error Resume Next
   frmSO5B60A.Show vbModal
End Sub

Private Sub mnuSO5B70_Click()
  On Error Resume Next
   frmSO5B70A.Show vbModal
End Sub

Private Sub mnuSO5B80_Click()
  On Error Resume Next
   frmSO5B80A.Show vbModal
End Sub

Private Sub mnuSO5B90_Click()
  On Error Resume Next
   frmSO5B90A.Show vbModal
End Sub

Private Sub mnuSO5BA0_Click()
  On Error GoTo ChkErr
    On Error Resume Next
      frmSO5BA0A.Show vbModal
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "mnuSO5BA0_Click")
End Sub

Private Sub mnuSO5BB0_Click()
  On Error Resume Next
    frmSO5BB0A.Show vbModal
End Sub

Private Sub mnuSO5BC0_Click()
  On Error Resume Next
   frmSO5BC0A.Show vbModal
End Sub

Private Sub mnuSO5BD0_Click()
  On Error Resume Next
   frmSO5BD0A.Show vbModal
End Sub

Private Sub mnuSO5BE0_Click()
  On Error Resume Next
   frmSO5BE0A.Show vbModal
End Sub

Private Sub mnuSO5BF0_Click()
  On Error Resume Next
   frmSO5BF0A.Show vbModal
End Sub

Private Sub mnuSO5BG0_Click()
  On Error Resume Next
   frmSO5BG0A.Show vbModal
End Sub

Private Sub mnuSO5BH0_Click()
  On Error Resume Next
   frmSO5BH0A.Show vbModal
End Sub

Private Sub mnuSO5BI0_Click()
  On Error Resume Next
   frmSO5BI0A.Show vbModal
    
End Sub

Private Sub mnuSO5C10_Click()
  On Error Resume Next
   frmSO5C10A.Show vbModal
End Sub

Private Sub mnuSO5C20_Click()
  On Error Resume Next
   frmSO5C20A.Show vbModal
End Sub

Private Sub mnuSO5C30_Click()
  On Error Resume Next
   frmSO5C30A.Show vbModal
End Sub

Private Sub mnuSO5C40_Click()
  On Error Resume Next
   frmSO5C40A.Show vbModal
End Sub

Private Sub mnuSO5C50_Click()
  On Error Resume Next
   frmSO5C50A.Show vbModal
End Sub

Private Sub mnuSO5C60_Click()
  On Error Resume Next
   frmSO5C60A.Show vbModal
End Sub

Private Sub mnuSO5C70_Click()
  On Error Resume Next
    frmSO5C70A.Show vbModal
End Sub

Private Sub mnuSO5C80_Click()
  On Error Resume Next
    frmSO5C80A.Show vbModal
End Sub

Private Sub mnuSO5C90_Click()
  On Error Resume Next
    frmSO5C90A.Show vbModal
End Sub

Private Sub mnuSO5CA0_Click()
  On Error Resume Next
    frmSO5CA0A.Show vbModal
End Sub

Private Sub mnuSO5CB0_Click()
  On Error Resume Next
    frmSO5CB0A.Show vbModal
End Sub

Private Sub mnuSO5CC0_Click()
  On Error Resume Next
    frmSO5CC0A.Show vbModal
End Sub

Private Sub mnuSO5CD0_Click()
  On Error Resume Next
    frmSO5CD0A.Show vbModal
End Sub

Private Sub mnuSO5D00_Click()
  On Error Resume Next
  Dim strCompName As String
  Dim strServiceType As String
  strCompName = ""
  strServiceType = getServiceType(gCompCode)
  strCompName = gcnGi.Execute("Select Description From CD039 Where CodeNo=" & gCompCode).GetString & ""
  On Error GoTo ChkErr
  CallExternalFile "SO5D00", "SO5D00", GetDSN & " " & _
                                       GetUID & " " & _
                                       GetPWD & " " & _
                                       gCompCode & " " & _
                                       strCompName & " " & _
                                       strServiceType & " " & _
                                       garyGi(1) & " " & _
                                       "V40"
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "mnuSO5D00_Click")
End Sub

Private Sub mnuSO5E00_Click()
  On Error Resume Next
  Dim strCompName As String
  Dim strServiceType As String
  strCompName = ""
  strServiceType = getServiceType(gCompCode)
  strCompName = gcnGi.Execute("Select Description From CD039 Where CodeNo=" & gCompCode).GetString & ""
  On Error GoTo ChkErr
  CallExternalFile "SO5E00", "SO5E00", GetDSN & " " & _
                                       GetUID & " " & _
                                       GetPWD & " " & _
                                       gCompCode & " " & _
                                       strCompName & " " & _
                                       strServiceType & " " & _
                                       garyGi(1) & " " & _
                                      "V40"
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "mnuSO5E00_Click")
End Sub

Private Sub mnuSO5F10_Click()
  On Error Resume Next
    frmSO5F10A.Show vbModal
End Sub

Private Sub mnuSO5F20_Click()
  On Error Resume Next
    frmSO5F20A.Show vbModal
End Sub

Private Sub mnuSO5F30_Click()
  On Error Resume Next
    frmSO5F30A.Show vbModal
End Sub

Private Sub mnuSO5G00_Click()
  On Error GoTo ChkErr
    Screen.MousePointer = vbHourglass
    garyGi(15) = oldCompStr
    ChangeComp frmCompany
    garyGi(15) = garyGi(9)
    garyGi(16) = GetRsValue("select TableOwner From CD039 Where CodeNo = " & garyGi(9)) & ""
    Screen.MousePointer = vbDefault
  Exit Sub
ChkErr:
    ErrSub Me.Name, "mnuSO5G00_Click"
End Sub
