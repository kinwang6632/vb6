VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mediaTCBank"
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
        strBankId = .BankId
        strCitemQry = .uCitemQry
        strUPDTime = .uUPDTime
         '#7033 增加付款種類 By Kin 2015/04/28
        strPTCode = .uPTCode
        strPTName = .uPTName
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
    Dim strFailUpdWhere As String
    Dim strErrCode As String
    Dim strErrDescript As String
    Dim strFail As String
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
    
    '******************************************************************************************
    '#3417 失敗時要更新的語法 By Kin 2007/12/11
    '#4388 增加更新異動人員與異動時間 By Kin 2009/05/04
    strFailUpdSQL = Empty
    If blnFailUpd Then
        strFailUpdSQL = "Update " & strOwnerName & "SO033 Set UCCode=" & strUCCode & _
                        ",UPDTime='" & strUPDTime & "',UPDEN='" & strUpdEn & "'" & _
                        ",UCName='" & strUCName & "',Note=Note"
                                                                                  
    End If
    '******************************************************************************************
    
    lngTime = Timer
    gcnGi.BeginTrans
    While Not File.AtEndOfLine   '讀取轉帳資料
        strData = File.ReadLine   '讀取一列資料
        If strMedia = 1 Then
            strBillNo = MidMbcs(strData, 36, 11)
        Else
            strBillNo = GetBillNo_Old(MidMbcs(strData, 36, 11), strServiceType)
        End If
        strFailUpdWhere = Empty
        '********************************************************************************************
        '#3417 要更新UCCode與UCName的 Where 條件 By Kin 2007/12/11
        strFailUpdWhere = IIf(strMedia = "1", " MediaBillNo='", " BillNo='") & strBillNo & "'" & _
                        " And UCCode is Not Null" & _
                        " And CancelFlag<>1 " & _
                        IIf(strCitemQry <> "", " And CitemCode In (" & strCitemQry & ")", "")
       '*********************************************************************************************
       
       '若扣帳不成功，則寫入錯誤log檔，錯誤筆數加１
       '#4388 UPD時 失敗名稱要在前面,單據編號在後面 By Kin 2009/05/13
        strErrCode = Empty
        strErrDescript = Empty
        Select Case MidMbcs(strData, 57, 2)
            Case "01"
                ErrFile.WriteLine ("單據編號： " & strBillNo & " 存款不足!")
                strErrCode = "01"
                strErrDescript = " 存款不足! ;單據編號： " & strBillNo
                lngErrCount = lngErrCount + 1
            Case "02"
                ErrFile.WriteLine ("單據編號： " & strBillNo & " 此帳號未代繳!")
                strErrCode = "02"
                strErrDescript = " 此帳號未代繳! ;單據編號： " & strBillNo
                lngErrCount = lngErrCount + 1
            Case "03"
                ErrFile.WriteLine ("單據編號： " & strBillNo & " 已終止委託代繳!")
                strErrCode = "03"
                strErrDescript = " 已終止委託代繳! ;單據編號： " & strBillNo
                lngErrCount = lngErrCount + 1
            Case "04"
                ErrFile.WriteLine ("單據編號： " & strBillNo & " 此編號未建檔!")
                strErrCode = "04"
                strErrDescript = " 此編號未建檔! ;單據編號： " & strBillNo
                lngErrCount = lngErrCount + 1
            Case "05"
                ErrFile.WriteLine ("單據編號： " & strBillNo & " 存戶帳號錯誤!")
                strErrCode = "05"
                strErrDescript = " 存戶帳號錯誤! ;單據編號： " & strBillNo
                lngErrCount = lngErrCount + 1
            Case "06"
                ErrFile.WriteLine ("單據編號： " & strBillNo & " 結清戶、靜止戶、繼承戶、已銷戶!")
                strErrCode = "06"
                strErrDescript = " 結清戶、靜止戶、繼承戶、已銷戶! ;單據編號： " & strBillNo
                lngErrCount = lngErrCount + 1
            Case "07"
                ErrFile.WriteLine ("單據編號： " & strBillNo & " 拒往戶、交換票據凍結、扣押金額不足!")
                strErrCode = "07"
                strErrDescript = " 拒往戶、交換票據凍結、扣押金額不足! ;單據編號： " & strBillNo
                lngErrCount = lngErrCount + 1
            Case "98"
                ErrFile.WriteLine ("單據編號： " & strBillNo & " 其他原因!")
                strErrCode = "98"
                strErrDescript = " 其他原因! ;單據編號： " & strBillNo
                lngErrCount = lngErrCount + 1
            Case "00"   '表入帳成功
                '取該筆入/扣帳金額(B18~B30後兩碼為小數)
                strRealAmt = MidMbcs(strData, 18, 11)
                
                '若入帳日期為空白，則取(B49~B56)為入帳日期
                If Len(strRealDate) = 0 Then
                   strRealDate = MidMbcs(strData, 49, 8)
                   strRealDate = Trim(str(1911 + Val(Mid(strRealDate, 1, 2))) & Mid(strRealDate, 3, 4))
                End If
               
                     '若Para24=1,則用媒體單號入帳
                If ChkData(strBillNo, strRealAmt, strClctEn, strClctName, strRealDate, strUpdEn, strCMCode, strCMName, strPTCode, strPTName, strBankName, IIf(strMedia = 1, False, True), strCitemQry, strMedia, , strBankId) = False Then
                   IntoAccount = False
                   GoTo Rollback
                End If
        End Select
        '******************************************************************************************************************
        '#3417 假如有錯誤，更新UCCode與UCName By Kin 2007/12/11
        If strErrDescript <> Empty Then
            If blnFailUpd Then
                '#3913 因為原本寫法為strFailUpdSQL=strFailUpdSQL & "..."，這樣第二次後會連第一次都串進去，多加一個變數儲存　By Kin 2008/04/30
                strFail = Empty
                strFail = strFailUpdSQL & "||'失敗代碼:" & strErrCode & ";失敗原因:" & strErrDescript & _
                                "; 未收原因:" & strUCCode & _
                                "; 未收原因名稱:" & strUCName & "'" & _
                                " Where " & strFailUpdWhere

                gcnGi.Execute strFail
            End If
        End If
        '******************************************************************************************************************

           
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
