VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mediaPost3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
Private strOption As String         '判斷之前option的狀態
Private strCitemQry As String
Private rsTmp As New ADODB.Recordset

Public Sub Action()
  On Error GoTo ChkErr
      Call IntoData
      If Not IntoAccount Then
         objStorePara.UpdState = False
         Exit Sub
      Else
         objStorePara.UpdState = False
      End If
       
      Call MsgResult       '顯示結果訊息
      '若錯誤筆數大於0則用NotePad來顯示檔案內容
      If lngErrCount > 0 Then Call Shell("NotePad " & ErrFilePath, vbNormalFocus)
      File.Close
      ErrFile.Close
      Set fso = Nothing
      
   Exit Sub
ChkErr:
    ErrSub "mediaCShing", "Action"
End Sub

Private Sub IntoData()
On Error GoTo ChkErr

   With objStorePara
     strFilePath = .FilePath
     strRealDate = .RealDate
     strSourcePath = .SourcePath
     strUpdEn = .UpdEn
     strPrgName = .PrgName
     strClctEn = .uClctEn
     strClctName = .uClctName
     strCMCode = .uCMCode
     strCMName = .uCMName
     strServiceType = .ServiceType
     strCompCode = .CompCode
     strBankName = .BankName
     strCitemQry = .uCitemQry
   End With
   
   If InitData = False Then
      File.Close
      ErrFile.Close
      Set fso = Nothing
      Exit Sub
   End If
  Exit Sub
ChkErr:
  ErrSub "mediaCShing", "IntoData"
End Sub

Private Function IntoAccount() As Boolean
On Error GoTo ChkErr
 Dim strBillNo As String
 Dim strRealAmt As String
 Dim strData As String
 Dim strYYMM As String
 Dim rsTmp As New ADODB.Recordset
 Dim strSQL As String
 Dim rsPara As New ADODB.Recordset
 Dim strMedia As String

        lngCount = 0
        lngErrCount = 0
        TmpFilePath = strFilePath & "\Tmp" & strPrgName & "In.log"
        ErrFilePath = strFilePath & "\" & strPrgName & "Err.log"
        Call GetFilePath(strFilePath)
        fso.CopyFile strSourcePath, TmpFilePath
        Set File = fso.OpenTextFile(TmpFilePath, ForReading)
        Set ErrFile = fso.CreateTextFile(ErrFilePath)
        
        strNowTime = RightNow
        strSQL = "Select Para24 From " & strOwnerName & "SO043 Where Rownum=1"
        If Not GetRS(rsPara, strSQL, gcnGi) Then Exit Function
        strMedia = GetFieldValue(rsPara, "Para24") & ""
        
        lngTime = Timer
        gcnGi.BeginTrans
        While Not File.AtEndOfLine   '讀取轉帳資料
            strData = File.ReadLine   '讀取一列資料
'            strCMCode = ""
'            strCMName = ""
            'para24=1,取媒體單號(B44~B54)來入帳
            If strMedia = 1 Then
               strBillNo = MidMbcs(strData, 44, 11)
            'para24=0,取單據編號(B44~B54),並用GetBillNo_New()轉換成15碼單編
            Else
               strBillNo = GetBillNo_Old(MidMbcs(strData, 44, 11), strServiceType)
            End If
            
            '取該筆入/扣帳金額(B33~B43後兩碼為小數)
            strRealAmt = MidMbcs(strData, 33, 11)
            
            '若入帳日期為空白，則取(B9~B14)為交易日期
            If Len(strRealDate) = 0 Then
                '交易日期為中文年
                strRealDate = Val(MidMbcs(strData, 9, 6)) + 19110000
            End If
            
            '若Para24=1,則用媒體單號入帳
        If ChkData(strBillNo, strRealAmt, strClctEn, strClctName, strRealDate, strUpdEn, strCMCode, strCMName, strPTCode, strPTName, strBankName, IIf(strMedia = 1, False, True), strCitemQry, strMedia) = False Then
           IntoAccount = False
           GoTo Rollback
        End If
Nextloop:
        Wend
        gcnGi.CommitTrans
        IntoAccount = True
        On Error Resume Next
        rsTmp.Close
        Set rsTmp = Nothing
        rsPara.Close
        Set rsPara = Nothing
  Exit Function
Rollback:
        gcnGi.RollbackTrans
        MsgBox "更新失敗!", vbExclamation, "警告..."
        File.Close
        ErrFile.Close
        Set fso = Nothing
  Exit Function
ChkErr:
  ErrSub "mediaCShing", "IntoAccount"
End Function

