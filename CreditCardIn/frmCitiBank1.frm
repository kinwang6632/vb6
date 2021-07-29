VERSION 5.00
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.4#0"; "GiList.ocx"
Begin VB.Form frmCitiBank 
   BorderStyle     =   1  '單線固定
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6105
   Icon            =   "frmCitiBank.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   6105
   StartUpPosition =   2  '螢幕中央
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
      Left            =   630
      TabIndex        =   1
      Top             =   1965
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
      Left            =   3900
      TabIndex        =   0
      Top             =   1950
      Width           =   1410
   End
   Begin prjGiList.GiList gilCMCode 
      Height          =   315
      Left            =   1980
      TabIndex        =   2
      Top             =   630
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
      Left            =   1980
      TabIndex        =   3
      Top             =   1080
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
      Height          =   195
      Left            =   885
      TabIndex        =   5
      Top             =   675
      Width           =   780
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
      Height          =   195
      Left            =   885
      TabIndex        =   4
      Top             =   1125
      Width           =   780
   End
End
Attribute VB_Name = "frmCitiBank"
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
Private strBankName As String       '銀行名稱
Private strCompCode As String       '公司別
Private strOption As String          '判斷之前option的狀態
Private rsTmp As New ADODB.Recordset
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
   End With
   
   If InitData = False Then
      File.Close
      ErrFile.Close
      Set Fso = Nothing
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
 Dim rsTmp1 As New ADODB.Recordset
 Dim strsql As String
 Dim strCM As String
 
        lngCount = 0
        lngErrCount = 0
        
        TmpFilePath = strFilePath & "\Tmp" & strPrgName & "In.log"
        ErrFilePath = strFilePath & "\" & strPrgName & "Err.log"
        
        Call GetFilePath(strFilePath)
        
        Fso.CopyFile strSourcePath, TmpFilePath
        Set File = Fso.OpenTextFile(TmpFilePath, ForReading)
        Set ErrFile = Fso.CreateTextFile(ErrFilePath)
        
        strNowTime = Now
        lngTime = Timer
        gcnGi.BeginTrans
        While Not File.AtEndOfLine   '讀取轉帳資料
        strData = File.ReadLine
        
        '' 入帳單號 為回登單號 (B097~B112)
        
          strBillNo = Trim(Mid(strData, 97, 16))
        
           '  檢查授權碼(B46~B53)  ，若扣帳不成功，則寫入錯誤log檔，錯誤筆數加１
           ''  授權碼是數字才可以
           
           If Not IsNumeric(Mid(strData, 46, 8)) Or Len(Trim(Mid(strData, 46, 8))) = 0 Then
              ErrFile.WriteLine ("單據編號： " & strBillNo & " 信用卡扣帳失敗")
              lngErrCount = lngErrCount + 1
              GoTo Nextloop
           End If
           
           ''入帳金額
           
             '''   strRealAmt = CStr(CInt(Mid(strData, 38, 8)))
            If Not IsNumeric(Mid(strData, 38, 8)) Then
              strRealAmt = 0
            Else
             strRealAmt = CStr(CInt(Mid(strData, 38, 8)))
            End If
          
           '入帳日期 (B85 ~ B90)
           
              strRealDate = Mid(strData, 85, 6)
              strRealDate = Trim(CStr(2000 + Val(Mid(strRealDate, 1, 2))) & Mid(strRealDate, 3, 4))
              
          ''   strCM = Trim(Mid(strData, 55, 8))
          
              strClctEn = gilClctEn.GetCodeNo
              strClctName = gilClctEn.GetDescription
              strCMCode = gilCMCode.GetCodeNo
              strCMName = gilCMCode.GetDescription
              
              If ChkData( _
                              strBillNo, strRealAmt, strRealDate, strUpdEn, _
                              strCMCode, strCMName, strClctEn, strClctName) = False Then
                              
                    IntoAccount = False
                    GoTo Roolback
              End If
           
Nextloop:

        Wend
        gcnGi.CommitTrans
        IntoAccount = True
        On Error Resume Next
        rsTmp1.Close
        Set rsTmp1 = Nothing
        
  Exit Function
Roolback:
        gcnGi.RollbackTrans
        frmCitiBank.MousePointer = vbDefault
        MsgBox "更新失敗!", vbExclamation, "警告..."
        File.Close
        ErrFile.Close
        Set Fso = Nothing
  Exit Function
ChkErr:
  ErrSub "agenCitiBank", "IntoAccount"
End Function

Private Sub cmdCancel_Click()
On Error Resume Next
   Unload Me
End Sub

Private Sub cmdOK_Click()
On Error GoTo ChkErr
       
       If Not IsDataOk Then Exit Sub
       frmCitiBank.MousePointer = vbHourglass
       
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
       frmCitiBank.MousePointer = vbDefault
       File.Close
       ErrFile.Close
       Set Fso = Nothing
  
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
      Set Fso = Nothing
End Sub

Private Sub OpenRec()
On Error GoTo ChkErr
  Dim strsql As String
  Dim rsData As New ADODB.Recordset
  
     If rsData.State = adStateOpen Then rsData.Close
     strsql = _
          "Select " & _
          "A.CodeNo,A.Description,A.ModeNo,B.Description ModeName " & _
          "From " & _
          "CD031A A,CD031 B " & _
          "Where " & _
          "A.ModeNo=B.CodeNo And A.StopFlag=0 And A.CompCode = " & strCompCode & " And (A.ServiceType = '" & strServiceType & "' Or A.ServiceType Is Null) Order By A.CodeNo"
  
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
    SetLst gilClctEn, "CodeNo", "Description", 10, 20, , , "CD032", " Where CompCode = " & strCompCode
    SetLst gilCMCode, "CodeNo", "Description", 10, 20, , , "CD031", " Where ServiceType Is Null Or ServiceType = '" & strServiceType & "'"
Exit Sub
ChkErr:
   ErrSub Me.Name, "subGil"
End Sub

Private Sub RcdToScr()
On Error GoTo ChkErr
    
    If Dir(strFilePath & "\" & strPrgName & "In2.log") = "" Then
       
       gilClctEn.SetCodeNo ""
       gilClctEn.SetDescription ""
       gilCMCode.SetCodeNo ""
       gilCMCode.SetDescription ""
    Else
       With rsTmp
            If .State = adStateOpen Then .Close
            .Open strFilePath & "\" & strPrgName & "In2.log"
               gilClctEn.SetCodeNo .Fields("ClctEn") & ""
               gilClctEn.Query_Description
               gilCMCode.SetCodeNo .Fields("CMCode") & ""
               gilCMCode.Query_Description
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
       .Fields.Append "ClctName", adBSTR, 12
       .Fields.Append "CMCode", adBSTR, 3
       .Fields.Append "CMName", adBSTR, 12
       .Open
    End With
  Exit Sub
ChkErr:
  ErrSub Me.Name, "DefineRS"
End Sub

Private Function IsDataOk() As Boolean
On Error GoTo ChkErr

     IsDataOk = False

        If Len(gilCMCode.GetCodeNo & "") = 0 Then
           gilCMCode.SetFocus
           MsgBox "請選擇收費方式!", vbExclamation, "警告..."
           Exit Function
        ElseIf Len(gilClctEn.GetCodeNo & "") = 0 Then
           gilClctEn.SetFocus
           MsgBox "請選擇收費人員!", vbExclamation, "警告..."
           Exit Function
        End If
     IsDataOk = True
   Exit Function
ChkErr:
   ErrSub Me.Name, "IsDataOk"
End Function

