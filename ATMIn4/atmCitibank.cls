VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "atmCitibank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private strFilePath As String       'INI檔案路徑
Private strSourcePath As String     '原始檔案路徑
Private strRealDate As String       '入帳日期
Private strPrgName As String        '轉帳程式名稱
Private strUpdEn As String          '記錄異動人員
Private blnUpdate As Boolean        '更新程式是否成功
Private strClctEn As String           '收費人員代碼
Private strClctName As String         '收費人員名稱
Private strCMCode As String           '收費方式代碼
Private strCMName As String           '收費方式名稱
Private strPTCode As String           '付款種類
Private strPTName As String           '付款種類名稱
Private strServiceType As String      '服務類別

Public Sub Action()
On Error GoTo ChkErr
Dim TmpFilePath As String
Dim ErrFilePath As String

   TmpFilePath = strFilePath & "\Tmp" & strPrgName & "In.log"
   ErrFilePath = strFilePath & "\" & strPrgName & "Err.log"
   Call GetFilePath(strFilePath)
   Fso.CopyFile strSourcePath, TmpFilePath
   Set File = Fso.OpenTextFile(TmpFilePath, ForReading)
   Set ErrFile = Fso.CreateTextFile(ErrFilePath)
   
   If InitData = False Then
      File.Close
      ErrFile.Close
      Set Fso = Nothing
      Exit Sub
   End If
   blnUpdate = IntoAccount
   If blnUpdate = False Then  '更新失敗，將相關檔案關閉
      File.Close
      ErrFile.Close
      Set Fso = Nothing
      Exit Sub
   End If

   Call MsgResult       '顯示結果訊息
   '若錯誤筆數大於0則用NotePad來顯示檔案內容
   If lngErrCount > 0 Then Call Shell("NotePad " & ErrFilePath, vbNormalFocus)
              
   File.Close
   ErrFile.Close
   Set Fso = Nothing
  Exit Sub
ChkErr:
  ErrSub "clsCThrust", "AgencyIn"
End Sub


Private Function IntoAccount() As Boolean
On Error GoTo ChkErr
 Dim strBillNo As String
 Dim strRealAmt As String
 Dim intClctEn As Integer
 Dim strData As String
 Dim strYYMM As String
 Dim rstmp As New ADODB.Recordset
 Dim strSQL As String
 
 
        strNowTime = Now
        lngTime = Timer
        gcnGi.BeginTrans
        While Not File.AtEndOfLine   '讀取轉帳資料
           strData = File.ReadLine   '讀取一列資料
           '若非明細列不做以下動作
           If Mid(strData, 1, 1) <> "2" Then GoTo Nextloop
           
           '取單據編號--應繳年月(B65~B68)+帳單類別(B69~B69)+服務類別+客戶編號(B70~B75)
           strYYMM = Trim(CStr(Val(Mid(strData, 65, 4)) + 191100))
           If strServiceType = "" Then
              strBillNo = strYYMM & Mid(strData, 69, 1) & "I" & GetString(Mid(strData, 70, 6), 7, giRight, True)
           Else
              strBillNo = strYYMM & Mid(strData, 69, 1) & strServiceType & GetString(Mid(strData, 70, 6), 7, giRight, True)
           End If
           
           '若扣帳不成功，則寫入錯誤log檔，錯誤筆數加１
           If Mid(strData, 63, 2) <> "00" And Mid(strData, 63, 2) <> "  " Then
              ErrFile.WriteLine ("單據編號： " & strBillNo & " ATM扣帳不成功")
              lngErrCount = lngErrCount + 1
              GoTo Nextloop
           End If
           '取該筆入/扣帳金額(B41~B52)
           strRealAmt = Mid(strData, 41, 12)
           '取該筆收件單位(B55~B62)及收件單位名稱
           If strClctEn = "" Then
                strClctEn = Val(Mid(strData, 55, 8))
                
                Select Case strClctEn
                    Case 1
                      strClctName = "郵局自動轉帳"
                    Case 2
                      strClctName = "郵局櫃檯代收"
                    Case 3
                      strClctName = "7-ELEVEn"
                    Case 4
                      strClctName = "全家"
                    Case 5
                      strClctName = "ATM"
                End Select
           End If
           '取該筆收費方式
           If strCMCode = "" Then
              strCMCode = "4"
              strSQL = "Select Description From CD031 Where CodeNo = '4'"
              strCMName = GetRsValue(strSQL, gcnGi) & ""
           End If
           '若入帳日期為空白，則取(B21~B26)為入帳日期
           If Len(strRealDate) = 0 Then
              strRealDate = Mid(strData, 21, 6)
              strRealDate = Trim(Str(1911 + Val(Mid(strRealDate, 1, 2))) & Mid(strRealDate, 3, 4))
           End If
           
           '執行ChkData檢查資料是否有誤
           If ChkData(strBillNo, strRealAmt, strClctEn, strClctName, strRealDate, strUpdEn, strCMCode, strCMName, strPTCode, strPTName, True) = False Then
              IntoAccount = False
              Exit Function
           End If
Nextloop:
        Wend
        gcnGi.CommitTrans
        IntoAccount = True
        Set rstmp = Nothing
  Exit Function
ChkErr:
  
  ErrSub "AgencyIn", "IntoAccount"
End Function

Public Property Let FilePath(ByVal vData As String)
    strFilePath = vData
End Property

Public Property Let RealDate(ByVal vData As Variant)
    strRealDate = vData
End Property

Public Property Let SourcePath(ByVal vData As Variant)
    strSourcePath = vData
End Property

Public Property Let UpdEn(ByVal vData As Variant)
    strUpdEn = vData
End Property

Public Property Get UpdState() As Variant
    UpdState = blnUpdate
End Property

Public Property Let PrgName(ByVal vData As Variant)
    strPrgName = vData
End Property

Public Property Set ugcnGi(ByVal vData As ADODB.Connection)
    Set gcnGi = vData
End Property
'收費方式
Public Property Let uCMCode(ByVal vData As Variant)
    strCMCode = vData
End Property
'收費方式名稱
Public Property Let uCMName(ByVal vData As Variant)
    strCMName = vData
End Property
'收費人員
Public Property Let uClctEn(ByVal vData As Variant)
    strClctEn = vData
End Property
'收費人員名稱
Public Property Let uClctName(ByVal vData As Variant)
    strClctName = vData
End Property
'付款種類
Public Property Let uPTCode(ByVal vData As Variant)
    strPTCode = vData
End Property
'付款種類名稱
Public Property Let uPTName(ByVal vData As Variant)
    strPTName = vData
End Property
'服務種類
Public Property Let uServiceType(ByVal vData As Variant)
    strServiceType = vData
End Property
