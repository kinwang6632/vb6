VERSION 5.00
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{A0C51418-555F-11D3-BC10-0050BA01A3A9}#9.6#0"; "GiMulti.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmSO3291A 
   BorderStyle     =   1  '單線固定
   Caption         =   "亞太扣款授權資料作業[SO3291A]"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8100
   Icon            =   "SO3291A.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   8100
   StartUpPosition =   1  '所屬視窗中央
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
      Left            =   2340
      TabIndex        =   20
      Top             =   3735
      Width           =   4725
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
      Left            =   7140
      TabIndex        =   19
      Top             =   3750
      Width           =   450
   End
   Begin VB.CheckBox chkContent 
      Caption         =   "銀行名稱依電子檔內容對應(銀行代碼需與來源資料相符)"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   2340
      TabIndex        =   2
      Top             =   1290
      Width           =   5430
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
      Left            =   6315
      TabIndex        =   10
      Top             =   4425
      Width           =   1410
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "F2.確定"
      Default         =   -1  'True
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
      Left            =   495
      TabIndex        =   9
      Top             =   4425
      Width           =   1410
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
      Left            =   7140
      TabIndex        =   7
      Top             =   3345
      Width           =   450
   End
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
      Left            =   2325
      TabIndex        =   6
      ToolTipText     =   "請輸入字檔之路徑及檔名！"
      Top             =   3330
      Width           =   4725
   End
   Begin prjGiList.GiList gilCMCode 
      Height          =   315
      Left            =   2340
      TabIndex        =   4
      Top             =   1620
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
   Begin prjGiList.GiList gilClctEn 
      Height          =   315
      Left            =   2340
      TabIndex        =   5
      Top             =   1980
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
   Begin prjGiList.GiList gilEmpNO 
      Height          =   315
      Left            =   2340
      TabIndex        =   3
      Top             =   2340
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
   Begin MSComDlg.CommonDialog cnd1 
      Left            =   3810
      Top             =   4425
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin prjGiList.GiList gilCD046 
      Height          =   315
      Left            =   2340
      TabIndex        =   1
      Top             =   900
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
   Begin prjGiList.GiList gilBank 
      Height          =   315
      Left            =   2340
      TabIndex        =   0
      Top             =   525
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
   Begin Gi_Multi.GiMulti gmsServArea 
      Height          =   375
      Left            =   390
      TabIndex        =   8
      Top             =   2805
      Width           =   7260
      _ExtentX        =   12806
      _ExtentY        =   661
      ButtonCaption   =   "週期性收費項目"
   End
   Begin prjGiList.GiList gilCompCode 
      Height          =   315
      Left            =   2340
      TabIndex        =   18
      Top             =   150
      Width           =   2730
      _ExtentX        =   4815
      _ExtentY        =   556
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
      FldWidth2       =   1800
      F5Corresponding =   -1  'True
   End
   Begin VB.Label lblErrpath 
      AutoSize        =   -1  'True
      Caption         =   "問題參考檔"
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
      Left            =   1245
      TabIndex        =   21
      Top             =   3780
      Width           =   975
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "銀行"
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
      Left            =   1875
      TabIndex        =   17
      Top             =   615
      Width           =   390
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "發票資料依據"
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
      Left            =   1095
      TabIndex        =   16
      Top             =   990
      Width           =   1170
   End
   Begin VB.Label lblDatapath 
      AutoSize        =   -1  'True
      Caption         =   "資料檔位置"
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
      Left            =   1245
      TabIndex        =   15
      Top             =   3405
      Width           =   975
   End
   Begin VB.Label lblClctEn 
      AutoSize        =   -1  'True
      Caption         =   "付款種類"
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
      Left            =   1485
      TabIndex        =   14
      Top             =   2055
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "收費方式"
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
      Left            =   1485
      TabIndex        =   13
      Top             =   1695
      Width           =   780
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "受理人員"
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
      Left            =   1485
      TabIndex        =   12
      Top             =   2415
      Width           =   780
   End
   Begin VB.Label lblComp 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "公  司  別"
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
      Left            =   1500
      TabIndex        =   11
      Top             =   210
      Width           =   765
   End
End
Attribute VB_Name = "frmSO3291A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public FSO As New FileSystemObject
Private File(1) As TextStream
Private blnFileOpen As Boolean
Dim lngDouble As Long
Private Sub chkContent_Click()
If chkContent.Value = 1 Then
   gilBank.Enabled = False
   Label4.ForeColor = &H404040
Else
   gilBank.Enabled = True
   Label4.ForeColor = vbRed
End If
End Sub

Private Sub cmdCancel_Click()
On Error GoTo ChkErr

   Unload Me
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "cmdCancel_Click")
End Sub

Private Sub cmdDataPath_Click()
On Error GoTo ChkErr
        With cnd1
                .FileName = ""
              ''  .Filter = "文字檔|*.txt"
                .Filter = "文字檔 (*.txt;*.dat)|*.txt;*.dat|所有檔案 (*.*)|*.*"
                .InitDir = ""
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
                .FileName = ""
             ''   .Filter = "文字檔|*.txt"
                 .Filter = "文字檔 (*.txt;*.dat)|*.txt;*.dat|所有檔案 (*.*)|*.*"
                .InitDir = ""
                .Action = 1
                txtErrPath = .FileName
        End With
Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "cmdErrPath_Click")
End Sub

Private Sub cmdOk_Click()
If IsDataOk = False Then Exit Sub
Dim strData As String
Dim S As String
Dim rsSO001 As New ADODB.Recordset
Dim errCount As Integer
Dim testCount As Integer

  S = txtDataPath.Text
  Set File(0) = FSO.OpenTextFile(S, ForReading)
      S = txtErrPath.Text
  If Not FSO.FileExists(S) Then
      Set File(1) = FSO.CreateTextFile(S)
  Else
      Set File(1) = FSO.OpenTextFile(S, ForWriting)
  End If
  blnFileOpen = True
  errCount = 0
  lngDouble = 0
 Do While Not File(0).AtEndOfLine
 testCount = testCount + 1

         strData = File(0).ReadLine
         Call GetRS(rsSO001, "SELECT * FROM " & GetOwner & "SO001 WHERE CustID=" & Trim(Mid(strData, 38, 20)), gcnGi, adUseClient, adOpenDynamic)
         If rsSO001.EOF Or rsSO001.BOF Then
                File(1).WriteLine "客戶主檔裏找不到客戶編號等於 " & Trim(Mid(strData, 38, 20)) & " 的客戶資料　!!"
                errCount = errCount + 1
         Else
                If Trim(Mid(strData, 29, 1)) = 1 Or Trim(Mid(strData, 29, 1)) = 3 Then
                        Call InsertToSO016(strData, rsSO001)
                        Call InsertToSO002A(strData, rsSO001)
                        Call ChkSO003(Trim(Mid(strData, 38, 20)), Trim(Mid(strData, 58, 14)), strData)
                Else
                        Call UpdateStopState(strData)
                End If
         End If
         rsSO001.Close
  Loop
  MsgBox "客戶資料處理完畢 !!"
  If errCount > 0 Then
       MsgBox " 總共產生 " & errCount & " 筆錯誤 !! "
      ''Shell S, vbHide
  End If
  If lngDouble > 0 Then
         MsgBox _
               "總共有 " & lngDouble & " 筆客戶編號資料於客戶帳戶資料檔裏已有資料，不再作新增 !! ", _
               vbInformation + vbOKOnly, "訊息"
  End If
  If blnFileOpen = True Then
    Set FSO = Nothing
    File(0).Close: Set File(0) = Nothing
    File(1).Close: Set File(1) = Nothing
End If
Exit Sub
ChkErr:
     MsgBox Err.HelpContext
    Call ErrSub(Me.Name, "cmdOK_Click")
End Sub

Private Sub Form_Load()
          Screen.MousePointer = vbDefault
            subGil
            subGim
            lngDouble = 0
End Sub
Private Sub subGil()
  On Error GoTo ChkErr
            
            SetgiList gilCompCode, "CodeNo", "Description", "CD039", , , , , , , "WHERE CODENO IN (" & garyGi(15) & ")"
            gilCompCode.ListIndex = 1
            SetgiList gilCD046, "CodeNo", "Description", "CD046"
            gilCD046.SetCodeNo "C"
            gilCD046.Query_Description
   Exit Sub
ChkErr:
   ErrSub Me.Name, "subGil"
End Sub

Private Sub subGim()
Dim strWhereBank As String
On Error GoTo ChkErr

    
Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "subGim")
End Sub

Private Sub gilBank_Change()
Dim strEmpNo  As String
strEmpNo = RPxx(gcnGi.Execute("SELECT EMPNO FROM   " & GetOwner & "CD018 WHERE CODENO =" & gilBank.GetCodeNo).GetString)
gilEmpNO.SetCodeNo strEmpNo
gilEmpNO.Query_Description
End Sub

Private Sub gilCD046_Change()
            SetgiList gilCMCode, "CodeNo", "Description", "CD031", , , , , , , " WHERE STOPFLAG  = 0 AND (ServiceType = '" & gilCD046.GetCodeNo & "'   OR ServiceType IS NULL) AND (REFNO =2 OR REFNO =4)", False
            SetgiList gilClctEn, "CodeNo", "Description", "CD032", , , , , , , " Where  STOPFLAG  = 0 AND  (ServiceType = '" & gilCD046.GetCodeNo & "'     OR ServiceType IS NULL )", False
            Call SetgiMulti(gmsServArea, "CodeNo", "Description", "CD019", "代碼", "名稱", " WHERE  PERIODFLAG=1 AND STOPFLAG  = 0 ", False)
End Sub

Private Sub gilCompCode_Change()
            SetgiList gilEmpNO, "EmpNo", "EmpName", "CM003", , , 1050, 2200, , , " Where CompCode = " & gilCompCode.GetCodeNo & "  AND StopFlag = 0 ", False
            SetgiList gilBank, "CodeNo", "Description", "CD018", , , , , , , " Where CompCode = " & gilCompCode.GetCodeNo & "  AND StopFlag = 0 ", False
End Sub
Private Sub InsertToSO016(ByVal pstrData As String, ByVal prsSO001 As ADODB.Recordset)

Dim strUpdate As String:   strUpdate = ""
Dim strFields As String: strFields = ""
Dim strFieldsValue As String: strFieldsValue = ""
Dim strGetDate As String:    strGetDate = ""
Dim intBankCode As Integer
Dim strBankName As String

On Error GoTo ChkErr

If chkContent.Value = 1 Then
       intBankCode = Trim(Mid(pstrData, 3, 7))
       strBankName = RPxx(gcnGi.Execute("SELECT Description FROM  " & GetOwner & "CD018 WHERE CODENO =" & intBankCode).GetString)
Else
       intBankCode = gilBank.GetCodeNo
       strBankName = gilBank.GetDescription
End If

strGetDate = "TO_DATE('" & Mid(pstrData, 21, 8) & "','YYYYMMDD')"
strFields = "CustID,AccountID,AcceptEn,AcceptName,AcceptTime,Proposer,ID,BankCode,BankName,AccountName,AccountNameID," & _
                 "PropDate,SendDate,SnactionDate,CMCode,CMName,UpdEn,UpdTime,CitemStr,PTCode,PTName,CompCode,StopFlag"

strFieldsValue = Trim(Mid(pstrData, 38, 20)) & ",'" & Trim(Mid(pstrData, 58, 14)) & "','" & _
                            gilEmpNO.GetCodeNo & "','" & gilEmpNO.GetDescription & "'," & _
                            strGetDate & ",'" & prsSO001("CustName") & "','" & _
                            Trim(Mid(pstrData, 10, 11)) & "'," & _
                            intBankCode & ",'" & strBankName & "','" & _
                            prsSO001("CustName") & "','" & _
                            Trim(Mid(pstrData, 10, 11)) & "'," & _
                            strGetDate & "," & strGetDate & "," & strGetDate & "," & _
                            gilCMCode.GetCodeNo & ",'" & gilCMCode.GetDescription & "','" & _
                            garyGi(0) & "','" & _
                            GetDTString(RightNow) & "','" & _
                            GetSO003SeqNO(Trim(Mid(pstrData, 38, 20))) & "'," & _
                            gilClctEn.GetCodeNo & ",'" & gilClctEn.GetDescription & "'," & gilCompCode.GetCodeNo & ",0 "
                   
strUpdate = "INSERT INTO  " & GetOwner & "SO106 (" & strFields & ") VALUES ( " & strFieldsValue & ")"
gcnGi.Execute strUpdate
Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "InsertToSO106")
End Sub
Private Sub InsertToSO002A(ByVal pstrData As String, ByVal prsSO001 As ADODB.Recordset)
Dim strUpdate As String:   strUpdate = ""
Dim strFields As String: strFields = ""
Dim strValue As String: strValue = ""
Dim strGetDate As String:    strGetDate = ""
Dim rsSO002 As New ADODB.Recordset
Dim strChargeAddress As String
Dim strMailAddress As String
Dim intBankCodeSO002A As Integer
Dim strBankNameSO002A As String

On Error GoTo ChkErr

Dim intSO002ACount  As Integer
intSO002ACount = RPxx(gcnGi.Execute("SELECT COUNT(CUSTID) FROM SO002A WHERE CUSTID = " & Trim(Mid(pstrData, 38, 20)) & " AND COMPCODE = " & gilCompCode.GetCodeNo & " AND AccountNo = '" & Trim(Mid(pstrData, 58, 14)) & "'").GetString)
If intSO002ACount > 0 Then
     lngDouble = lngDouble + 1
    File(1).WriteLine "客戶編號  " & Trim(Mid(pstrData, 38, 20)) & " 於客戶帳戶資料檔裏已有資料，不再作新增 !! "
    Exit Sub
End If

If chkContent.Value = 1 Then
       intBankCodeSO002A = Trim(Mid(pstrData, 3, 7))
       strBankNameSO002A = RPxx(gcnGi.Execute( _
                                              "SELECT Description FROM  " & GetOwner & "CD018 WHERE CODENO =" & intBankCodeSO002A _
                                              ).GetString)
Else
       intBankCodeSO002A = gilBank.GetCodeNo
       strBankNameSO002A = gilBank.GetDescription
End If

strChargeAddress = RPxx(gcnGi.Execute("SELECT ADDRESS FROM  " & GetOwner & "SO014 WHERE AddrNo =" & prsSO001("ChargeAddrNO")).GetString)
strMailAddress = RPxx(gcnGi.Execute("SELECT ADDRESS FROM  " & GetOwner & "SO014 WHERE AddrNo =" & prsSO001("MailAddrNO")).GetString)
''以下的SO002 是否還要服務別
GetRS rsSO002, "SELECT * FROM  " & GetOwner & "SO002 WHERE CUSTID =" & prsSO001("CUSTID") & "  AND ServiceType ='" & gilCD046.GetCodeNo & "'", gcnGi, adUseClient, adOpenKeyset
strGetDate = "TO_DATE('" & Mid(pstrData, 21, 8) & "','YYYYMMDD')"

strFields = "CustId,BankCodE,BankName,ID,AccountNo,ChargeAddrNo,ChargeAddress," & _
                 "MailAddrNo,MailAddress,ChargeTitle,InvNo,InvTitle,InvAddress,InvoiceType,CitemStr,CompCode,StopFlag"
                   
strValue = Trim(Mid(pstrData, 38, 20)) & "," & _
                 intBankCodeSO002A & ",'" & strBankNameSO002A & "',0, '" & Trim(Mid(pstrData, 58, 14)) & "'," & _
                 prsSO001("ChargeAddrNo") & ",'" & strChargeAddress & "'," & _
                 prsSO001("MailAddrNO") & ",'" & strMailAddress & "','" & _
                 prsSO001("CustName") & "','" & rsSO002("InvNo") & "','" & rsSO002("InvTitle") & "','" & _
                 rsSO002("InvAddress") & "','" & rsSO002("InvoiceType") & "','" & _
                 GetSO003SeqNO(Trim(Mid(pstrData, 38, 20))) & "'," & gilCompCode.GetCodeNo & ",0 "
                   
strUpdate = "INSERT INTO  " & GetOwner & "SO002A(" & strFields & ") VALUES (" & strValue & ")"
gcnGi.Execute strUpdate
rsSO002.Close
Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "InsertToSO002A")
End Sub
Private Function ChkSO003(ByVal lngCustId As Long, strAccount As String, ByVal ppstrData As String) As Boolean

Dim rsSO003 As New ADODB.Recordset
Dim rsSO003NotNull As New ADODB.Recordset
Dim strSQL As String
Dim strWhere As String

Dim intBankCodeSO033 As Integer
Dim strBankNameSO033 As String

On Error GoTo ChkErr

If chkContent.Value = 1 Then
       intBankCodeSO033 = Trim(Mid(ppstrData, 3, 7))
       strBankNameSO033 = RPxx(gcnGi.Execute( _
                                              "SELECT Description FROM  " & GetOwner & "CD018 WHERE CODENO =" & intBankCodeSO033 _
                                              ).GetString)
Else
       intBankCodeSO033 = gilBank.GetCodeNo
       strBankNameSO033 = gilBank.GetDescription
End If
ChkSO003 = False

If Len(gmsServArea.GetQryStr) > 0 Then
    strWhere = "CustID=" & lngCustId & " AND " & _
                       "CitemCode " & gmsServArea.GetQryStr
Else
    strWhere = "CustID=" & lngCustId & " "
End If

strSQL = "SELECT * FROM  " & GetOwner & "SO003 WHERE " & strWhere & " AND " & _
                "AccountNo IS NULL "
               
Call GetRS(rsSO003, strSQL, gcnGi, adUseClient, adOpenDynamic, adLockOptimistic)
If Not (rsSO003.EOF Or rsSO003.BOF) Then
            ''作更新SO003 的動作
            Call gcnGi.Execute("UPDATE  " & GetOwner & "SO003 SET " & _
                                          "BankCode = " & intBankCodeSO033 & ", BankName='" & strBankNameSO033 & "'," & _
                                          "PTCode = " & gilClctEn.GetCodeNo & ", PTName='" & gilClctEn.GetDescription & "'," & _
                                          "CMCode = " & gilCMCode.GetCodeNo & ", CMName='" & gilCMCode.GetDescription & "'," & _
                                          "AccountNo='" & strAccount & "' WHERE " & strWhere)
Else
            strSQL = "SELECT * FROM  " & GetOwner & "SO003 WHERE " & strWhere & " AND " & _
                            "AccountNo  IS  NOT NULL "
            Call GetRS(rsSO003NotNull, strSQL, gcnGi, adUseClient, adOpenDynamic, adLockOptimistic)
            If Not (rsSO003NotNull.EOF Or rsSO003NotNull.BOF) Then
                 
                 File(1).WriteLine ("客戶編號 " & lngCustId & " 的客戶在客戶週期性收費項目資料檔裏，已有此筆資料且帳號並非空值　!!")
            End If
End If
ChkSO003 = True
rsSO003.Close
Set rsSO003 = Nothing
Exit Function
ChkErr:
ChkSO003 = False
    Call ErrSub(Me.Name, "ChkSO003")
End Function
Private Function IsDataOk() As Boolean
On Error GoTo ChkErr:
Dim blnOK As Boolean

blnOK = True
If Len(gilCompCode.GetCodeNo) = 0 Then IsDataOk = False: Call ShowBlankMessage("公司別", blnOK): gilCompCode.SetFocus:    Exit Function
If chkContent.Value = 0 Then
    If Len(gilBank.GetCodeNo) = 0 Then IsDataOk = False: Call ShowBlankMessage("銀行別", blnOK): gilBank.SetFocus: Exit Function
End If
If Len(gilCD046.GetCodeNo) = 0 Then IsDataOk = False: Call ShowBlankMessage("發票資料依據", blnOK): gilCD046.SetFocus: Exit Function

If Len(gilCMCode.GetCodeNo) = 0 Then IsDataOk = False: Call ShowBlankMessage("收費方式", blnOK): gilCMCode.SetFocus: Exit Function
If Len(gilClctEn.GetCodeNo) = 0 Then IsDataOk = False: Call ShowBlankMessage("付款種類", blnOK): gilClctEn.SetFocus: Exit Function
If Len(gilEmpNO.GetCodeNo) = 0 Then IsDataOk = False: Call ShowBlankMessage("受理人員", blnOK): gilEmpNO.SetFocus: Exit Function
If Len(txtDataPath.Text) = 0 Then IsDataOk = False: Call ShowBlankMessage("位置資料檔", blnOK): txtDataPath.SetFocus: Exit Function
If Len(txtErrPath.Text) = 0 Then IsDataOk = False: Call ShowBlankMessage("問題參考檔", blnOK): txtErrPath.SetFocus: Exit Function
If Len(gmsServArea.GetQueryCode) = 0 Then IsDataOk = False: Call ShowBlankMessage("週期性收費項目", blnOK): gmsServArea.SetFocus: Exit Function
If Len(gmsServArea.GetQueryCode) > 100 Then
    IsDataOk = False
    MsgBox "週期性收費項目的選項太多，請重新調整 !!", vbOKOnly + vbCritical, "條件值設定訊息"
    gmsServArea.SetFocus
    Exit Function
End If
IsDataOk = blnOK
Exit Function
ChkErr:
    Call ErrSub(Me.Name, "IsDataOk")
End Function

Private Sub ShowBlankMessage(ByVal strField As String, ByRef pblnOK As Boolean)
Dim strMessage  As String
 strMessage = "欄位 '" & strField & "'  不可空白　!! "
 MsgBox strMessage, vbCritical + vbOKOnly, "欄位檢查訊息"
 pblnOK = False
End Sub

Private Sub UpdateStopState(ByVal pstrData As String)
Dim strUpdate As String
Dim strCitemCodeQuery As String

On Error GoTo ChkErr  ''Mid(strData, 38, 20)), Trim(Mid(strData, 58, 14)
    strUpdate = "Update  " & GetOwner & "SO106 SET StopDate =TO_DATE('" & Trim(Mid(pstrData, 21, 8)) & "','YYYYMMDD')," & _
                       "StopFlag=1, UpdEn='" & garyGi(0) & "', UpdTime='" & GetDTString(RightNow) & "' WHERE " & _
                       "CustId=" & Trim(Mid(pstrData, 38, 20)) & " AND  AccountID='" & Trim(Mid(pstrData, 58, 14)) & "'"
                       
    gcnGi.Execute strUpdate
    strUpdate = "UPDATE  " & GetOwner & "SO002A SET  StopDate =TO_DATE('" & Trim(Mid(pstrData, 21, 8)) & "','YYYYMMDD')," & _
                       "StopFlag=1 WHERE CustId=" & Trim(Mid(pstrData, 38, 20)) & " AND  AccountNO='" & Trim(Mid(pstrData, 58, 14)) & "'"
    gcnGi.Execute strUpdate
    
    If Len(gmsServArea.GetQryStr) > 0 Then
       strCitemCodeQuery = "and CitemCode " & gmsServArea.GetQryStr
    Else
       strCitemCodeQuery = " "
    End If
    
    strUpdate = "UPDATE  " & GetOwner & "SO003 SET " & _
                        "BankCode=null, BankName=null, AccountNO=null, " & _
                        "UpdTime='" & GetDTString(RightNow) & "'," & _
                        "Upden='" & garyGi(0) & "' WHERE " & _
                        "CustID=" & Trim(Mid(pstrData, 38, 20)) & " AND AccountNO='" & Trim(Mid(pstrData, 58, 14)) & "'  " & strCitemCodeQuery
    gcnGi.Execute strUpdate
Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "UpdateStopState")
End Sub
Private Function GetSO003SeqNO(ByVal strCustid As String) As String
Dim strSeqNo As String
Dim rsSeqNO As New ADODB.Recordset
Dim blnFirst As Boolean

On Error GoTo ChkErr:
blnFirst = True
strSeqNo = ""
Call GetRS(rsSeqNO, "SELECT SeqNO FROM SO003 WHERE " & _
                                     "CUSTID =" & strCustid & " AND CompCode =" & gilCompCode.GetCodeNo & " AND CitemCode " & _
                                     gmsServArea.GetQryStr, gcnGi, adUseClient, adOpenKeyset)

If Not (rsSeqNO.EOF Or rsSeqNO.BOF) Then
   Do While Not rsSeqNO.EOF
     If blnFirst = False Then strSeqNo = strSeqNo & "'',''" & rsSeqNO(0) Else strSeqNo = rsSeqNO(0)
     rsSeqNO.MoveNext
     blnFirst = False
   Loop
   strSeqNo = "''" & strSeqNo & "''"
   GetSO003SeqNO = strSeqNo
Else
     GetSO003SeqNO = ""
End If
Exit Function

ChkErr:
    Call ErrSub(Me.Name, "GetSO003SeqNO")
End Function
