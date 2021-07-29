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
'******************************************************************************************'
'此為#2825新增此一功能mediaPost3 By Kin 2006/11/30 Release
'******************************************************************************************'
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
Private intRefNo As Integer
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
    CloseRecordset rsCD013
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
        '#3527 判斷參考號，如果是2以CityBankAtm做為單據依據 By Kin 2007/10/04
        intRefNo = .uRefNo
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
        '*************************************************************************************************
        '問題集2974 原本單據格式為So033.Mediabillno+So106.AccountNameID，但匯出時，則是相反，會導致資料不一致
        '           此需求,將格式修正為一致, 以匯出的格式為準，原本從第55取11碼，現在改為從第65碼取11碼
        '           For Karen 2007/01/10
        '*************************************************************************************************
        If Mid(strData, 1, 1) = 1 Then
            '*********************************************************************************************
            '#3527 判斷參考號是否為2，如果為2則單據使用CityBankATM By Kin 2007/10/04
            If intRefNo <> 2 Then
                If strMedia = 1 Then
                   strBillNo = MidMbcs(strData, 65, 11) '問題集 2974 Modify
                Else
                   strBillNo = GetBillNo_New(MidMbcs(strData, 65, 11)) '問題集 2974 Modify
                End If
            Else
                '#3527 測試不OK，字元55~70應是抓取16個字，不是15個字(原本抓15個) By Kin 2007/10/17
                strBillNo = Trim(MidMbcs(strData, 55, 16))
            End If
            
            '**********************************************************************************************
            '#3417 失敗時要更新的Where條件 By Kin 2007/12/11
            strFailUpdWhere = Empty
            If intRefNo <> 2 Then
                strFailUpdWhere = IIf(strMedia = "1", " MediaBillNo='", " BillNo='") & strBillNo & "'" & _
                                " And UCCode is Not Null" & _
                                " And CancelFlag<>1 " & _
                                IIf(strCitemQry <> "", " And CitemCode In (" & strCitemQry & ")", "")
            Else
                '#3856 CitiBankATM 打錯成 CityBankATM By Kin 2008/04/10
                strFailUpdWhere = " CitiBankATM='" & strBillNo & "'" & _
                                " And UCCode is Not Null" & _
                                " And CancelFlag<>1 " & _
                                IIf(strCitemQry <> "", " And CitemCode In (" & strCitemQry & ")", "")
            End If
            '***********************************************************************************************


            
            '*********************************************************************************************
            '問題集2869 新增錯誤原因，並且增加判斷如果回授原因不為空白，又沒落在下面的原因，則秀出原因：未知
            
            strErrCode = Empty
            strErrDescript = Empty
            '#4388 測試不OK UPD時,原因要在前面 單據編號在後面 By Kin 2009/05/12
            Select Case MidMbcs(strData, 79, 2)
                Case "01"
                    ErrFile.WriteLine ("單據編號： " & strBillNo & " 存款不足!")
                    strErrCode = "01"
                    strErrDescript = " 存款不足! ;單據編號： " & strBillNo
                    lngErrCount = lngErrCount + 1
                Case "02"
                    ErrFile.WriteLine ("單據編號： " & strBillNo & " 未申請代繳或用戶號碼錯誤!")
                    strErrCode = "02"
                    strErrDescript = " 未申請代繳或用戶號碼錯誤! ;單據編號： " & strBillNo
                    lngErrCount = lngErrCount + 1
                Case "03"
                    ErrFile.WriteLine ("單據編號： " & strBillNo & " 已終止代繳!")
                    strErrCode = "03"
                    strErrDescript = " 已終止代繳! ;單據編號： " & strBillNo
                    lngErrCount = lngErrCount + 1
                Case "04"
                    ErrFile.WriteLine ("單據編號： " & strBillNo & " 或用戶號碼錯誤!")
                    strErrCode = "04"
                    strErrDescript = " 或用戶號碼錯誤! ;單據編號： " & strBillNo
                    lngErrCount = lngErrCount + 1
                Case "05"
                    ErrFile.WriteLine ("單據編號： " & strBillNo & " 存簿帳號非活動戶!")
                    strErrCode = "05"
                    strErrDescript = " 存簿帳號非活動戶! ;單據編號： " & strBillNo
                    lngErrCount = lngErrCount + 1
                Case "07"
                    ErrFile.WriteLine ("單據編號： " & strBillNo & " 存簿帳號為凍結戶!")
                    strErrCode = "07"
                    strErrDescript = " 存簿帳號為凍結戶! ;單據編號： " & strBillNo
                    lngErrCount = lngErrCount + 1
                Case "08"
                    ErrFile.WriteLine ("單據編號： " & strBillNo & " 存簿帳號錯誤!")
                    strErrCode = "08"
                    strErrDescript = " 存簿帳號錯誤! ;單據編號： " & strBillNo
                    lngErrCount = lngErrCount + 1
                Case "09"
                    ErrFile.WriteLine ("單據編號： " & strBillNo & " 存簿帳號已終止!")
                    strErrCode = "09"
                    strErrDescript = " 存簿帳號已終止! ;單據編號： " & strBillNo
                    lngErrCount = lngErrCount + 1
                Case "11"
                    ErrFile.WriteLine ("單據編號： " & strBillNo & " (劃撥) 業務支票用戶!")
                    strErrCode = "11"
                    strErrDescript = " (劃撥) 業務支票用戶! ;單據編號： " & strBillNo
                    lngErrCount = lngErrCount + 1
                Case "12"
                    ErrFile.WriteLine ("單據編號： " & strBillNo & " 帳號錯誤或非連線局帳護戶!")
                    strErrCode = "12"
                    strErrDescript = " 帳號錯誤或非連線局帳護戶! ;單據編號： " & strBillNo
                    lngErrCount = lngErrCount + 1
                Case "13"
                    ErrFile.WriteLine ("單據編號： " & strBillNo & " 已終止!")
                    strErrCode = "13"
                    strErrDescript = " 已終止! ;單據編號： " & strBillNo
                    lngErrCount = lngErrCount + 1
                Case "14"
                    ErrFile.WriteLine ("單據編號： " & strBillNo & " 身分證號不符!")
                    strErrCode = "14"
                    strErrDescript = " 身分證號不符! ;單據編號： " & strBillNo
                    lngErrCount = lngErrCount + 1
                Case "19"
                    ErrFile.WriteLine ("單據編號： " & strBillNo & " 系統等其他原因!")
                    strErrCode = "19"
                    strErrDescript = " 系統等其他原因! ;單據編號： " & strBillNo
                    lngErrCount = lngErrCount + 1
                Case Else   '表入帳成功 問題集2869 多增加判斷是不是還有其它失敗代碼 2008/11/16
                    If MidMbcs(strData, 79, 2) = "  " Then
                        '取該筆入/扣帳金額(B44~B52後兩碼為小數)
                        strRealAmt = MidMbcs(strData, 44, 9)
                        '若入帳日期為空白，則取(B9~B14)為交易日期
                        If Len(strRealDate) = 0 Then
                            '2010/07/15 Jacky 5680
                            If MidMbcs(strData, 17, 3) = "   " Then
                                '交易日期為中文年
                                strRealDate = Val(MidMbcs(strData, 10, 7)) + 19110000
                            Else
                                '交易日期為中文年
                                strRealDate = Val(MidMbcs(strData, 10, 6)) + 19110000
                            End If
                        End If
                    '若Para24=1,則用媒體單號入帳
                        If ChkData(strBillNo, strRealAmt, strClctEn, strClctName, strRealDate, strUpdEn, strCMCode, strCMName, strPTCode, strPTName, strBankName, IIf(strMedia = 1, False, True), strCitemQry, strMedia, intRefNo, strBankId) = False Then
                           IntoAccount = False
                           GoTo Rollback
                        End If
                    Else
                        ErrFile.WriteLine ("單據編號： " & strBillNo & " 未知!   代碼：" & MidMbcs(strData, 79, 2))
                        strErrCode = MidMbcs(strData, 79, 2)
                        strErrDescript = " 未知! ;單據編號： " & strBillNo
                        lngErrCount = lngErrCount + 1
                    End If
            End Select
            
            '******************************************************************************************************************
            '#3417 假如有錯誤，更新UCCode與UCName By Kin 2007/12/11
            If strErrDescript <> Empty Then
                If blnFailUpd Then
                    '#3913 只有第一筆會成功，因為原本寫法為strFailUpdSQL=strFailUpdSQL & "..."，這樣第二次後會連第一次都串進去，多加一個變數儲存
                    strFail = Empty
                    strFail = strFailUpdSQL & "||'失敗代碼:" & strErrCode & ";失敗原因:" & strErrDescript & _
                                    "; 未收原因:" & strUCCode & _
                                    "; 未收原因名稱:" & strUCName & "'" & _
                                    " Where " & strFailUpdWhere
    
                    gcnGi.Execute strFail
                End If
            End If
            '******************************************************************************************************************
            
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


