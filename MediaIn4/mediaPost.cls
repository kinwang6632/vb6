VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mediaPost"
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
Private strBankId As String         '銀行代碼
Private strCompCode As String       '公司別
Private strOption As String         '判斷之前option的狀態
Private strCitemQry As String
Private rsTmp As New ADODB.Recordset
Private rsCD013 As New ADODB.Recordset
Private strFailUpdSQL As String
Private strUPDTime As String
Private intRefNo As Integer
Public Sub Action()
  On Error GoTo ChkErr
    
    blnFailUpd = False
    '*************************************************************************************************************************************
    '#3417 讀取未收原因為5,如果有資料,更新失敗時要回寫UCCode與UCName By Kin 2007/12/11
    If Not GetRS(rsCD013, "Select CodeNo,Description From " & strOwnerName & "CD013 Where Nvl(RefNo,0)=5 And StopFlag<>1 Order by CodeNo Desc", _
                  gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
    If rsCD013.EOF Then
        blnFailUpd = False
    Else
        strUCCode = rsCD013("CodeNo") & ""
        strUCName = rsCD013("Description") & ""
        blnFailUpd = True
    End If
    '*************************************************************************************************************************************

  
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
    Set FSO = Nothing
    CloseRecordset rsCD013
    Exit Sub
ChkErr:
    CloseRecordset rsCD013
    ErrSub "mediaCShing", "Action"
End Sub

Private Sub IntoData()
On Error GoTo ChkErr
    '#5875 增加判斷參考號,如果是1讀取的位置全部往右移1位 By Kin 2010/12/08
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
        strBankId = .BankId
        strCitemQry = .uCitemQry
        strUPDTime = .uUPDTime
        '#7033 增加付款種類 By Kin 2015/04/28
        strPTCode = .uPTCode
        strPTName = .uPTName
        intRefNo = .uRefNo
    End With
   
    If InitData = False Then
       File.Close
       ErrFile.Close
       Set FSO = Nothing
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
    FSO.CopyFile strSourcePath, TmpFilePath
    Set File = FSO.OpenTextFile(TmpFilePath, ForReading)
    Set ErrFile = FSO.CreateTextFile(ErrFilePath)
    
    strNowTime = RightNow
    strSQL = "Select Para24 From " & strOwnerName & "SO043 Where Rownum=1"
    If Not GetRS(rsPara, strSQL, gcnGi) Then Exit Function
    strMedia = GetFieldValue(rsPara, "Para24") & ""
    
    lngTime = Timer
    gcnGi.BeginTrans
    Dim intRedPos As Integer
    intRedPos = 0
    '#5875 如果參考號為1日期讀取7碼,其它位置全部往後移一位 By Kin 2010/12/08
    If intRefNo = 1 Then
        intRedPos = intRedPos + 1
    End If
    While Not File.AtEndOfLine   '讀取轉帳資料
        strData = File.ReadLine   '讀取一列資料
        If strMedia = 1 Then
            strBillNo = MidMbcs(strData, 44 + intRedPos, 11)
        Else
            strBillNo = GetBillNo_Old(MidMbcs(strData, 44 + intRedPos, 11), strServiceType)
        End If
        
        '取該筆入/扣帳金額(B33~B43後兩碼為小數)
        strRealAmt = MidMbcs(strData, 33 + intRedPos, 11)
        
        '若入帳日期為空白，則取(B9~B14)為交易日期
        If Len(strRealDate) = 0 Then
            '交易日期為中文年
            strRealDate = Val(MidMbcs(strData, 9, 6 + intRedPos)) + 19110000
        End If
        
        '若Para24=1,則用媒體單號入帳
    If ChkData(strBillNo, strRealAmt, strClctEn, strClctName, strRealDate, strUpdEn, strCMCode, strCMName, strPTCode, strPTName, strBankName, IIf(strMedia = 1, False, True), strCitemQry, strMedia, , strBankId) = False Then
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
    Set FSO = Nothing
  Exit Function
ChkErr:
  ErrSub "mediaCShing", "IntoAccount"
End Function


