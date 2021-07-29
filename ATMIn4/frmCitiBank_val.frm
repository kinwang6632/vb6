VERSION 5.00
Object = "{3BB22E57-824A-11D3-BD70-0080C8F80BC4}#3.4#0"; "GiGridR.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Begin VB.Form frmCitiBank_val 
   ClientHeight    =   4230
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9345
   BeginProperty Font 
      Name            =   "新細明體"
      Size            =   9.75
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCitiBank_val.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4230
   ScaleWidth      =   9345
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&X)"
      Height          =   465
      Left            =   5775
      TabIndex        =   6
      Top             =   3585
      Width           =   1410
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "F2.確定"
      Default         =   -1  'True
      Height          =   465
      Left            =   2010
      TabIndex        =   5
      Top             =   3585
      Width           =   1410
   End
   Begin VB.Frame fra2 
      Caption         =   "請選擇收費人員及收費方式的填值方式"
      Height          =   3045
      Left            =   105
      TabIndex        =   7
      Top             =   240
      Width           =   9120
      Begin VB.OptionButton optByParameter 
         Caption         =   "依收費參數輸入"
         Height          =   315
         Left            =   4605
         TabIndex        =   3
         Top             =   285
         Width           =   2340
      End
      Begin VB.OptionButton optByScreen 
         Caption         =   "依畫面輸入"
         Height          =   315
         Left            =   180
         TabIndex        =   0
         Top             =   285
         Width           =   1620
      End
      Begin prjGiGridR.GiGridR ggrData 
         Height          =   2145
         Left            =   4905
         TabIndex        =   4
         Top             =   750
         Width           =   4020
         _ExtentX        =   7091
         _ExtentY        =   3784
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
      Begin prjGiList.GiList gilCMCode 
         Height          =   315
         Left            =   1530
         TabIndex        =   1
         Top             =   750
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
         Left            =   1530
         TabIndex        =   2
         Top             =   1155
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
      Begin VB.Label lblCMCode 
         AutoSize        =   -1  'True
         Caption         =   "收費方式"
         Height          =   195
         Left            =   435
         TabIndex        =   9
         Top             =   795
         Width           =   780
      End
      Begin VB.Label lblClctEn 
         AutoSize        =   -1  'True
         Caption         =   "收費人員"
         Height          =   195
         Left            =   435
         TabIndex        =   8
         Top             =   1200
         Width           =   780
      End
   End
End
Attribute VB_Name = "frmCitiBank_val"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private blnUpdate As Boolean        '更新程式是否成功
Private TmpFilePath As String
Private ErrFilePath As String
Private strFilePath As String       'INI檔案路徑
Private strRealDate As String       '入帳日期
Private strSourcePath As String     '原始檔案路徑
Private strUpdEn As String          '記錄異動人員
Private strPrgName As String        '轉帳程式名稱
Private strCMCode As String         '收費方式代碼
Private strCMName As String         '收費方式名稱
Private strClctEn As String         '收費人員代碼
Private strClctName As String       '收費人員名稱
Private strPTCode As String         '付款種類
Private strPTName As String         '付款種類名稱
Private strServiceType As String    '服務類別
Private strCorpId As String
Private intBankId As Integer
Private strBankName As String       '銀行名稱
Private strCompCode As String       '公司別
Private strGetOwner As String       'OwnerName
Private strOption As String          '判斷之前option的狀態
Private rsTmp As New ADODB.Recordset
Private intpara24 As Integer

Private Sub IntoData()
On Error GoTo ChkErr

   With objStorePara
     strFilePath = .FilePath
     strRealDate = .RealDate
     strSourcePath = .SourcePath
     strUpdEn = .UpdEn
     strPrgName = .PrgName
     strPTCode = .PTCode
     strPTName = .PTName
     strServiceType = .ServiceType
     strCompCode = .CompCode
     strGetOwner = .uGetOwner
     intBankId = .BankID
   End With
   strCorpId = GetRsValue("Select CorpId From " & strGetOwner & "CD018 Where CodeNO = " & intBankId) & ""
   
'   lngCount = 0
'   lngErrCount = 0
'   TmpFilePath = strFilePath & "\Tmp" & strPrgName & "In.log"
'   ErrFilePath = strFilePath & "\" & strPrgName & "Err.log"
'   Call GetFilePath(strFilePath)
'   Fso.CopyFile strSourcePath, TmpFilePath
'   Set File = Fso.OpenTextFile(TmpFilePath, ForReading)
'   Set ErrFile = Fso.CreateTextFile(ErrFilePath, True)
   
   If InitData = False Then
      File.Close
      ErrFile.Close
      Set FSO = Nothing
      Exit Sub
   End If
  Exit Sub
ChkErr:
  ErrSub Me.Name, "IntoData"
End Sub


Private Function IntoAccount() As Boolean
On Error GoTo ChkErr
 Dim strBillNo As String
 Dim strRealAmt As String
 Dim strData As String
 Dim strYYMM As String
 Dim rsTmp As New ADODB.Recordset
 Dim strBankName As String
 Dim strSQL As String
 
        If strCorpId = "" Then
            MsgBox "銀行代號沒有設定收件單位代號!!請查證!!", vbExclamation, gimsgPrompt
            Unload Me
            Exit Function
        End If
        lngCount = 0
        lngErrCount = 0
        TmpFilePath = strFilePath & "\Tmp" & strPrgName & "In.log"
        ErrFilePath = strFilePath & "\" & strPrgName & "Err.log"
        Call GetFilePath(strFilePath)
        FSO.CopyFile strSourcePath, TmpFilePath
        Set File = FSO.OpenTextFile(TmpFilePath, ForReading)
        Set ErrFile = FSO.CreateTextFile(ErrFilePath, True)
        
        strNowTime = RightNow
        lngTime = Timer
        strBankName = "Citibank"
        gcnGi.BeginTrans
        '7179
        blnCrossCustCombine = Val(GetRsValue("Select Nvl(CrossCustCombine,0) From " & strOwnerName & "SO041", gcnGi) & "")
        While Not File.AtEndOfLine   '讀取轉帳資料
           strData = File.ReadLine   '讀取一列資料
            strCMCode = ""
            strCMName = ""
           '若非明細列不做以下動作
           If Mid(strData, 1, 1) <> "2" Then GoTo Nextloop
           '取虛擬帳號'888005'+(B67~B75)
           If Mid(strData, 69, 1) = "B" Or Mid(strData, 69, 1) = "T" Then
              strBillNo = strCorpId & Mid(strData, 67, 2) & "0" & Mid(strData, 70, 6)
           Else
              strBillNo = strCorpId & Mid(strData, 67, 9)
           End If
           '若扣帳不成功，則寫入錯誤log檔，錯誤筆數加１
           If Mid(strData, 63, 2) <> "00" And Mid(strData, 63, 2) <> "  " Then
              ErrFile.WriteLine ("單據編號： " & strBillNo & " ATM扣帳不成功")
              lngErrCount = lngErrCount + 1
              GoTo Nextloop
           End If
           '取該筆入/扣帳金額(B41~B52)
           strRealAmt = Mid(strData, 41, 12)
           
           '如果畫面選擇依參數輸入,則才取得收件單位及收費方式資料
           If strOption = "Parameter" Then
              
              '取該筆收件單位(B55~B62)及收件單位名稱
              strClctEn = Mid(strData, 10, 8)
              strClctName = strClctEn
              '取該筆收費方式
              strCMCode = Trim(Mid(strData, 55, 8))
              If strServiceType = "" Then
                strSQL = "Select A.ModeNo,B.Description From " & _
                          strGetOwner & "CD031A A," & strGetOwner & "CD031 B" & _
                          " Where A.ModeNo=B.CodeNo And A.CodeNo = '" & strCMCode & "'"
              Else
                strSQL = "Select A.ModeNo,B.Description From " & _
                          strGetOwner & "CD031A A," & strGetOwner & "CD031 B" & _
                          " Where A.ModeNo=B.CodeNo And A.CodeNo = '" & strCMCode & "' And (A.ServiceType = '" & strServiceType & "' Or A.ServiceType Is Null)"
              End If
              If Not GetRS(rsTmp, strSQL, gcnGi) Then
                 MsgBox "開啟Recordset失敗!", vbExclamation, "警告..."
                 Exit Function
              End If
              If rsTmp.RecordCount > 0 Then
                 strCMCode = rsTmp("ModeNo")
                 strCMName = rsTmp("Description")
              End If
           Else
              strClctEn = gilClctEn.GetCodeNo
              strClctName = gilClctEn.GetDescription
              strCMCode = gilCMCode.GetCodeNo
              strCMName = gilCMCode.GetDescription
           End If
           
           '若入帳日期為空白，則取(B21~B26)為入帳日期
           If Len(strRealDate) = 0 Then
              strRealDate = Mid(strData, 21, 6)
              strRealDate = Trim(str(1911 + Val(Mid(strRealDate, 1, 2))) & Mid(strRealDate, 3, 4))
           End If
           
           '執行ChkData檢查資料是否有誤
           If ChkData(strBillNo, strRealAmt, strClctEn, strClctName, strRealDate, strUpdEn, strCMCode, strCMName, strPTCode, strPTName, "", , , , intpara24, 15, strBillNo) = False Then
              IntoAccount = False
               GoTo Roolback
           End If
           DoEvents
Nextloop:
        Wend
        
        gcnGi.CommitTrans

        IntoAccount = True
        On Error Resume Next
        rsTmp.Close
        Set rsTmp = Nothing
  Exit Function
Roolback:
        gcnGi.RollbackTrans
        frmCitiBank_val.MousePointer = vbDefault
        MsgBox "更新失敗!", vbExclamation, "警告..."
        File.Close
        ErrFile.Close
        Set FSO = Nothing
  Exit Function
ChkErr:
  ErrSub Me.Name, "IntoAccount"
End Function

Private Sub cmdCancel_Click()
On Error Resume Next
   Unload Me
End Sub

Private Sub cmdOK_Click()
On Error GoTo ChkErr
       
       If Not IsDataOk Then Exit Sub
       frmCitiBank_val.MousePointer = vbHourglass
       blnUpdate = IntoAccount
       If blnUpdate = False Then  '更新失敗，將相關檔案關閉
          objStorePara.UpdState = False
          Exit Sub
       Else
          objStorePara.UpdState = True
       End If
      
       Call DefineRS
       Call ScrToRcd
    
       Call MsgResult       '顯示結果訊息
       '若錯誤筆數大於0則用NotePad來顯示檔案內容
       If lngErrCount > 0 Then Call Shell("NotePad " & ErrFilePath, vbNormalFocus)
       frmCitiBank_val.MousePointer = vbDefault
       File.Close
       ErrFile.Close
       Set FSO = Nothing
  
   Exit Sub
ChkErr:
   ErrSub Me.Name, "cmdOK_Click"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   On Error Resume Next
        If Not ChkGiList(KeyCode) Then Exit Sub
        If KeyCode = vbKeyEscape Then Unload Me
        If KeyCode = vbKeyF2 Then If cmdOK.Enabled = True Then cmdOK.Value = True
End Sub

Private Sub Form_Load()
  On Error Resume Next
      
      Me.Caption = objStorePara.BankName & ""
      Call IntoData
      Call subGil
      Call RcdToScr

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
      File.Close
      ErrFile.Close
      Set FSO = Nothing
End Sub

'依畫面輸入
Private Sub optByScreen_Click()
On Error GoTo ChkErr
      
      strOption = "Screen"
      gilCMCode.Enabled = True
      gilClctEn.Enabled = True
      ggrData.Enabled = False
   Exit Sub
ChkErr:
   ErrSub Me.Name, "optByScreen_Click"
End Sub

'依收費參數輸入
Private Sub optByParameter_Click()
On Error GoTo ChkErr

      strOption = "Parameter"

     ggrData.Enabled = True
     gilCMCode.Enabled = False
     gilClctEn.Enabled = False
     Call subGrd
     Call OpenRec
   Exit Sub
ChkErr:
   ErrSub Me.Name, "optByParameter_Click"
End Sub

Private Sub subGrd()
   '設定grid屬性
On Error GoTo ChkErr
   Dim mFlds As New prjGiGridR.GiGridFlds
   With mFlds
         .Add "CodeNo", , , , False, "收費來源代碼", vbLeftJustify
         .Add "Description", , , , False, "收費來源名稱", vbLeftJustify
         .Add "ModeNo", , , , False, "輸入代碼", vbLeftJustify
         .Add "ModeName", , , , False, "輸入名稱", vbLeftJustify
   End With
   ggrData.AllFields = mFlds
   ggrData.SetHead
   Exit Sub
ChkErr:
   Call ErrSub(Me.Name, "SubGrd")
End Sub

Private Sub OpenRec()
On Error GoTo ChkErr
  Dim strSQL As String
  Dim rsData As New ADODB.Recordset
  
     If rsData.State = adStateOpen Then rsData.Close
     If strServiceType = "" Then
        strSQL = "Select A.CodeNo,A.Description,A.ModeNo,B.Description ModeName From " & _
                 strGetOwner & "CD031A A," & strGetOwner & "CD031 B" & _
                 " Where A.ModeNo=B.CodeNo And A.StopFlag=0 And A.CompCode = " & strCompCode & " Order By A.CodeNo"
     Else
        strSQL = "Select A.CodeNo,A.Description,A.ModeNo,B.Description ModeName From " & _
                 strGetOwner & "CD031A A," & strGetOwner & "CD031 B" & _
                 " Where A.ModeNo=B.CodeNo And A.StopFlag=0 And A.CompCode = " & strCompCode & " And (A.ServiceType = '" & strServiceType & "' Or A.ServiceType Is Null) Order By A.CodeNo"
     End If
     If GetRS(rsData, strSQL) Then
        Set ggrData.Recordset = rsData
        ggrData.Refresh
     Else
        MsgBox "無法開啟Recodeset!!", vbExclamation, "警告..."
     End If
     If rsData.State = adStateOpen Then
        rsData.Close
        Set rsData = Nothing
     End If
   Exit Sub
ChkErr:
   ErrSub Me.Name, "OpenRec"
End Sub

Private Sub subGil()
On Error GoTo ChkErr
    intpara24 = GetSystemParaItem("Para24", strCompCode, strServiceType, "SO043")
    SetLst gilClctEn, "EmpNo", "EmpName", 10, 20, , , "CM003", " Where CompCode = " & strCompCode
    If intpara24 = 1 Then
        SetLst gilCMCode, "CodeNo", "Description", 3, 20, , , "CD031", ""
    Else
        SetLst gilCMCode, "CodeNo", "Description", 3, 20, , , "CD031", " Where ServiceType Is Null Or ServiceType = '" & strServiceType & "'"
    End If
   Exit Sub
ChkErr:
   ErrSub Me.Name, "subGil"
End Sub

Private Sub RcdToScr()
On Error GoTo ChkErr
    
    If Dir(strFilePath & "\" & strPrgName & "In4.log") = "" Then
       optByScreen.Value = True
       gilClctEn.SetCodeNo ""
       gilClctEn.SetDescription ""
       gilCMCode.SetCodeNo ""
       gilCMCode.SetDescription ""
    Else
       With rsTmp
            If .State = adStateOpen Then .Close
            .Open strFilePath & "\" & strPrgName & "In4.log"
            If .Fields("Option") = "Screen" Then
               optByScreen.Value = True
               gilClctEn.SetCodeNo .Fields("ClctEn") & ""
               gilClctEn.Query_Description
               gilCMCode.SetCodeNo .Fields("CMCode") & ""
               gilCMCode.Query_Description
            Else
               optByParameter.Value = True
            End If
       End With
    End If

   Exit Sub
ChkErr:
   ErrSub Me.Name, "RcdToScr"
End Sub

Private Sub ScrToRcd()
On Error GoTo ChkErr
    With rsTmp
       .AddNew
       .Fields("Option") = strOption
       If strOption = "Screen" Then
          .Fields("ClctEn") = gilClctEn.GetCodeNo
          .Fields("ClctName") = gilClctEn.GetDescription
          .Fields("CMCode") = gilCMCode.GetCodeNo
          .Fields("CMName") = gilCMCode.GetDescription
       Else
          .Fields("ClctEn") = ""
          .Fields("ClctName") = ""
          .Fields("CMCode") = ""
          .Fields("CMName") = ""
       End If
       Kill strFilePath & "\" & strPrgName & "In2.log"
       .Save strFilePath & "\" & strPrgName & "In2.log"
    End With
  Exit Sub
ChkErr:
  ErrSub Me.Name, "ScrToRcd"
End Sub

Private Sub DefineRS()
On Error GoTo ChkErr
  If rsTmp.State = adStateOpen Then rsTmp.Close
    With rsTmp
       .LockType = adLockOptimistic
       .Fields.Append "Option", adBSTR, 12
       .Fields.Append "ClctEn", adBSTR, 3
       .Fields.Append "ClctName", adBSTR, 20
       .Fields.Append "CMCode", adBSTR, 3
       .Fields.Append "CMName", adBSTR, 20
       .Open
    End With
  Exit Sub
ChkErr:
  ErrSub Me.Name, "DefineRS"
End Sub

Private Function IsDataOk() As Boolean
On Error GoTo ChkErr
     IsDataOk = False
     If strOption = "" Then
        optByScreen.SetFocus
        MsgBox "請選擇收費人員及收費方式的填值方式!", vbExclamation, "警告..."
        Exit Function
     ElseIf strOption = "Screen" Then
        If intpara24 = 0 And Len(gilCMCode.GetCodeNo & "") = 0 Then
           gilCMCode.SetFocus
           MsgBox "請選擇收費方式!", vbExclamation, "警告..."
           Exit Function
        ElseIf Len(gilClctEn.GetCodeNo & "") = 0 Then
           gilClctEn.SetFocus
           MsgBox "請選擇收費人員!", vbExclamation, "警告..."
           Exit Function
        End If
     End If
     IsDataOk = True
   Exit Function
ChkErr:
   ErrSub Me.Name, "IsDataOk"
End Function
