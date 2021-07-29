Attribute VB_Name = "EmcStb"
''"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
'' 這個模組用於處理 STB 結清作業的後續動作
'' 其中包含了呼叫 東副資訊的相關 DLL 物件
'' 呼叫的相關 DLL 如下
'' --  --
'' CCB_RentToSale
'' CCB_ISSUE
'' CCB_Pair
'' CCB_Retrurn
''--- --
''  接下來 log  一筆設備異動資料一 SO055'
''  修改該筆設備資料  (Wherer Seq No = 該筆設備流水號)
''"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
Option Explicit

Public Function CallSaleReturn(pcompcode As String, pBillNO As String, _
                                               pStbSno As String, pICCNO As String, _
                                               pReturnDate As String, pReturnAmout As Long, _
                                               pItem As String, pSeqSno As String, pPlaceC, pPlaceD) As Boolean
                                                


Dim rsOldCompCode As New ADODB.Recordset
Dim rsStbIcc As New ADODB.Recordset
Dim strSQL As String

Dim pCustName As String
Dim pOldCompcode As String
Dim pStbMSo As String
Dim pICCMSo As String
Dim OrginalBillNo As String

Dim strBugStep As String



On Error GoTo ChkErr

CallSaleReturn = False

' // 取得客戶名稱
OrginalBillNo = pBillNO
pCustName = gcnGi.Execute("SELECT CUSTNAME FROM " & GetOwner & "SO001 WHERE CUSTID = " & gCustId).GetString


'' // 取得舊公司別代碼

strBugStep = 1  ''  記錄目前程式碼執行的位置，以了解出錯的地方
'' ------------------------------------------------------------------------------------------------------------
rsOldCompCode.CursorLocation = adUseClient
rsOldCompCode.Open _
        "SELECT OLDCOMPCODE FROM " & GetOwner & "CD039 WHERE CODENO  = " & pcompcode, _
        gcnGi, adOpenKeyset, adLockReadOnly
If Not (rsOldCompCode.EOF Or rsOldCompCode.BOF) Then
      pOldCompcode = rsOldCompCode(0) & ""
Else
      pOldCompcode = ""
End If
rsOldCompCode.Close
Set rsOldCompCode = Nothing

'' // 取得轉換的單據編號
Dim strm As String
Select Case Mid(pBillNO, 5, 1)
  Case 0
    strm = Mid(pBillNO, 6, 1)
  Case 1
    Select Case Mid(pBillNO, 5, 2)
        Case 0
           strm = "A"
        Case 1
           strm = "B"
        Case 2
           strm = "C"
    End Select
End Select
pBillNO = Mid(pBillNO, 3, 2) & strm & Right(pBillNO, 9)
strSQL = " SELECT FROM " & GetOwner & "SOAC0201A," & GetOwner & "SOAC0201A WHERE  SOAC0201A. = SOAC0201B"

''//  取得 STB 的料號 從 SOAC0201A
strBugStep = 2  ''  記錄目前程式碼執行的位置，以了解出錯的地方
'' ------------------------------------------------------------------------------------------------------------

rsStbIcc.CursorLocation = adUseClient
rsStbIcc.Open _
    "SELECT MaterialNo FROM " & GetOwner & "SOAC0201A WHERE FACISNO = '" & pStbSno & "'", _
    gcnGi, adOpenKeyset, adLockReadOnly
    
If Not (rsStbIcc.EOF Or rsStbIcc.BOF) Then
     pStbMSo = rsStbIcc(0) & ""
Else
     pStbMSo = ""
End If

rsStbIcc.Close

''//  取得 ICC 的料號 從 SOAC0201B
strBugStep = 3 ''  記錄目前程式碼執行的位置，以了解出錯的地方
'' ------------------------------------------------------------------------------------------------------------

rsStbIcc.Open _
    "SELECT MaterialNo FROM " & GetOwner & "SOAC0201B WHERE FACISNO = '" & pICCNO & "'", _
    gcnGi, adOpenKeyset, adLockReadOnly
If Not (rsStbIcc.EOF Or rsStbIcc.BOF) Then
     pICCMSo = rsStbIcc(0) & ""
Else
     pICCMSo = ""
End If

rsStbIcc.Close
Set rsStbIcc = Nothing

strBugStep = 4 ''  記錄目前程式碼執行的位置，以了解出錯的地方
'' ------------------------------------------------------------------------------------------------------------

Dim arrIssueM(7) As String '' CCB_ISSUE 所需的 MASTER 一維陣列
Dim arrIssueD(1, 8) As String '' CCB_ISSUE 所需的 Detail  二維陣列

'' 設定 arrIssueM(7)  內容

        arrIssueM(0) = pOldCompcode
        arrIssueM(1) = pBillNO
        arrIssueM(2) = 5
        arrIssueM(3) = "0" & Format(Date, "eemmdd")
        arrIssueM(4) = "0" & Format(Date, "eemmdd") & Format(Time, "hhdd")
        arrIssueM(5) = garyGi(0)
        arrIssueM(6) = garyGi(1)
        arrIssueM(7) = gCustId
        

        arrIssueD(0, 0) = pStbMSo
        arrIssueD(0, 1) = pStbSno
        arrIssueD(0, 2) = 1
        arrIssueD(0, 3) = 3
        arrIssueD(0, 4) = 1
        arrIssueD(0, 5) = 0
        arrIssueD(0, 6) = pPlaceC
        arrIssueD(0, 7) = pPlaceD
        arrIssueD(0, 8) = 0

        arrIssueD(1, 0) = pICCMSo
        arrIssueD(1, 1) = pICCNO
        arrIssueD(1, 2) = 1
        arrIssueD(1, 3) = 3
        arrIssueD(1, 4) = 1
        arrIssueD(1, 5) = 0
        arrIssueD(1, 6) = pPlaceC
        arrIssueD(1, 7) = pPlaceD
        arrIssueD(1, 8) = 0

'' 整理 DLL CCB_Pair 所需的陣列參數
strBugStep = 5 ''  記錄目前程式碼執行的位置，以了解出錯的地方
'' ------------------------------------------------------------------------------------------------------------

Dim arrPairM(0, 6) As String '' CCB_Pair  所需的陣列

        arrPairM(0, 0) = pOldCompcode
        arrPairM(0, 1) = 2
        arrPairM(0, 2) = pStbMSo
        arrPairM(0, 3) = pStbSno
        arrPairM(0, 4) = pICCMSo
        arrPairM(0, 5) = pICCNO
        arrPairM(0, 6) = "0" & Format(Date, "eemmdd")

' 整理 DLL CCB_Return  所需的陣列參數
strBugStep = 6 ''  記錄目前程式碼執行的位置，以了解出錯的地方
'' ------------------------------------------------------------------------------------------------------------

Dim arrReturnm(1, 5) As String ' CCB_Return 所需的陣列

        arrReturnm(0, 0) = pOldCompcode
        arrReturnm(0, 1) = pBillNO
        arrReturnm(0, 2) = Format(Left(pReturnDate, 4) & "/" & Mid(pReturnDate, 5, 2) & "/" & Right(pReturnDate, 2), "eemmdd")
        arrReturnm(0, 3) = pStbMSo
        arrReturnm(0, 4) = pStbSno
        arrReturnm(0, 5) = Abs(pReturnAmout)
        
        
        arrReturnm(1, 0) = pOldCompcode
        arrReturnm(1, 1) = pBillNO
        arrReturnm(1, 2) = Format(Left(pReturnDate, 4) & "/" & Mid(pReturnDate, 5, 2) & "/" & Right(pReturnDate, 2), "eemmdd")
        arrReturnm(1, 3) = pICCMSo
        arrReturnm(1, 4) = pICCNO
        arrReturnm(1, 5) = 0
        
''     呼叫jacky所寫的模組，將這些陣列傳入給他的 COMMAND
''     陣列整理完成
strBugStep = 7 ''  記錄目前程式碼執行的位置，以了解出錯的地方
'' ------------------------------------------------------------------------------------------------------------

Dim O As Object
Dim intLinkMss As Integer
Dim strMessage As String
intLinkMss = gcnGi.Execute("SELECT LINKMSS FROM " & GetOwner & "CD039 WHERE CODENO  = " & pcompcode).GetString
If intLinkMss = 1 Then
    Dim cnComm As New ADODB.Connection
    Dim lngSeqNo As Long
    Dim blnMSSSynch As Boolean
        '直接與東富介接
        blnMSSSynch = GetRsValue("Select AcrossMSS From " & GetOwner & "Cd039 Where CodeNo = " & gCompCode) = 1
        If Not blnMSSSynch Then
            cnComm.Open GetCommonConnection
            lngSeqNo = GetInvoiceNo3("Send_Mss", cnComm)
            cnComm.BeginTrans
        Else
             '' 以下這一行  2003/04/07 加入的
             cnComm.Open gcnGi.ConnectionString
             frmSO1147A.pic1.Visible = True
        End If
       If MSSIssue(O, arrIssueM, arrIssueD, strMessage, lngSeqNo, cnComm) = False Then
              Call MsgBox(strMessage, vbOKOnly + vbCritical, "訊息 ")
              If Not blnMSSSynch Then cnComm.RollbackTrans
              Exit Function
       End If
       strMessage = ""
       If MSSPair(O, arrPairM, strMessage, lngSeqNo, cnComm) = False Then
              Call MsgBox(strMessage, vbOKOnly + vbCritical, "訊息 ")
              If Not blnMSSSynch Then cnComm.RollbackTrans
              Exit Function
       End If
       strMessage = ""
       
       '' 以下這一段改成新的模組，

'       If MSSReturn(O, arrReturnm, strMessage, lngSeqNo, cnComm) = False Then
'              Call MsgBox(strMessage, vbOKOnly + vbCritical, "訊息 ")
'              If Not blnMSSSynch Then cnComm.RollbackTrans
'              Exit Function
'       End If
       Dim lngA As Long
       Dim rDate As String
       rDate = "0" & CStr(Left(pReturnDate, 4) - 1911) & Right(pReturnDate, 4)
       If MSSReturn(O, 2, lngSeqNo, rDate, lngA, strMessage) = False Then
              Call MsgBox(strMessage, vbOKOnly + vbCritical, "訊息 ")
              If Not blnMSSSynch Then cnComm.RollbackTrans
              Exit Function
       End If
       
       If Not blnMSSSynch Then cnComm.CommitTrans
       
End If
''""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
'\\ log 一筆設備異動資料至 SO055
strBugStep = 8 ''  記錄目前程式碼執行的位置，以了解出錯的地方
'' ------------------------------------------------------------------------------------------------------------

''"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
    Dim stInsertInto As String
    Dim rsSO004 As New ADODB.Recordset
    
    rsSO004.CursorLocation = adUseClient
    rsSO004.Open _
         "SELECT  * FROM " & GetOwner & "SO004 WHERE   SO004.Seqno ='" & pSeqSno & "'", _
         gcnGi, adOpenKeyset, adLockReadOnly
    
    stInsertInto = "INSERT INTO " & GetOwner & "SO055 " & _
                          "(UPDEN,UPDTIME,CUSTID,CUSTNAme,COMPCODE,FUNCTYPE,PRDATEB,PREN1B,PRNAME1B," & _
                          "FACICODE,FACINAME,FACISNO,CVTID,BUYNAME,DUEDATE,AMOUNT,QUANTITY," & _
                          "UNITPRICE,INSTDATE,INSTEN1,INSTEN2,INSTNAME1,INSTNAME2,PREN2,PRNAME2," & _
                          "MEDIACODE,MEDIANAME,INTROID, INTRONAME,SMARTCARDNO )" & _
                          "VALUES" & _
                          "('" & garyGi(0) & "','" & Format(Date, "EEMMDD") & "'," & gCustId & " ,'" & pCustName & "'," & pcompcode & ",1," & _
                          "TO_DATE('" & pReturnDate & "','YYYYMMDD'),'" & _
                          garyGi(0) & "','" & garyGi(1) & _
                          "'," & GetNullString(rsSO004("FACICODE"), giLongV) & "," & GetNullString(rsSO004("FACINAME"), giStringV) & _
                          "," & GetNullString(rsSO004("FACISNO"), giStringV) & "," & GetNullString(rsSO004("CVTID"), giLongV) & "," & GetNullString(rsSO004("BUYNAME"), giStringV) & _
                          "," & GetNullString(rsSO004("DUEDATE"), giDateV) & "," & GetNullString(rsSO004("AMOUNT"), giLongV) & "," & GetNullString(rsSO004("QUANTITY"), giLongV) & _
                          "," & GetNullString(rsSO004("UNITPRICE"), giLongV) & "," & GetNullString(rsSO004("INSTDATE"), giDateV) & "," & GetNullString(rsSO004("INSTEN1"), giStringV) & _
                          "," & GetNullString(rsSO004("INSTEN2"), giStringV) & "," & GetNullString(rsSO004("INSTNAME1"), giStringV) & "," & GetNullString(rsSO004("INSTNAME2"), giStringV) & _
                          "," & GetNullString(rsSO004("PREN2"), giStringV) & "," & GetNullString(rsSO004("PRNAME2"), giStringV) & "," & GetNullString(rsSO004("MEDIACODE"), giLongV) & _
                          "," & GetNullString(rsSO004("MEDIANAME"), giStringV) & "," & GetNullString(rsSO004("INTROID"), giStringV) & "," & GetNullString(rsSO004("INTRONAME"), giStringV) & _
                          "," & GetNullString(rsSO004("SMARTCARDNO"), giStringV) & ")"
    gcnGi.Execute stInsertInto
''"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
'\\ 修改該筆設備資料
strBugStep = 9 ''  記錄目前程式碼執行的位置，以了解出錯的地方
'' ------------------------------------------------------------------------------------------------------------

''"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
 Dim strUpdate  As String

 strUpdate = "UPDATE " & GetOwner & "SO004 SET " & _
                  "PRDATE  = TO_DATE('" & pReturnDate & "','YYYYMMDD')" & " , " & _
                  "PREN1 = '" & garyGi(0) & "'," & _
                  "PRNAME1 = '" & garyGi(1) & "'," & _
                  "NOTE ='" & rsSO004("NOTE") & " ,{結清日銷售退回} ,BillNO:" & OrginalBillNo & ", Item:" & Space(2) & pItem & "'," & _
                  "UPDEN='" & garyGi(0) & "', " & _
                  "UPDTime='" & GetDTString(Now) & "'  " & _
                  "WHERE SEQNO = '" & pSeqSno & "'"

gcnGi.Execute strUpdate

strBugStep = 10 ''  記錄目前程式碼執行的位置，以了解出錯的地方
'' ------------------------------------------------------------------------------------------------------------
 strUpdate = "UPDATE " & GetOwner & "SO004 SET " & _
                  "PRDATE  = TO_DATE('" & pReturnDate & "','YYYYMMDD')" & " , " & _
                  "PREN1 = '" & garyGi(0) & "'," & _
                  "PRNAME1 = '" & garyGi(1) & "'," & _
                  "NOTE ='" & rsSO004("NOTE") & " ,{結清日銷售退回} ,BillNO:" & OrginalBillNo & ", Item:" & Space(2) & pItem & "'," & _
                  "UPDEN='" & garyGi(0) & "', " & _
                  "UPDTime='" & GetDTString(Now) & "'  " & _
                  "WHERE FACISNO = '" & rsSO004("SmartCardNo") & "'" & _
                  " AND CUSTID =" & rsSO004("CustId") & "  AND PRDATE IS  NULL"

gcnGi.Execute strUpdate
                
    rsSO004.Close
    Set rsSO004 = Nothing
CallSaleReturn = True

Exit Function
ChkErr:
    Call ErrSub("emcstb", "CallSaleReturn-" & "  位置碼[" & strBugStep & "]")
End Function

Public Function CallRentToSale(pcompcode As String, pBillNO As String, _
                                                pStbSno As String, pICCSno As String, pSeqSno As String, _
                                                pShouldAmt As Long)

Dim arrRentToSaleM(3) As String '' CCB_RentToSale 所需的 MASTER 一維陣列
Dim arrRentToSaleD(1, 2)  As String  '' CCB_RentToSale 所需的 Detail  二維陣列
Dim rsOldCompCode As New ADODB.Recordset
Dim strSQL As String
Dim rsSOSnoIcc As New ADODB.Recordset
Dim rsStbIcc As New ADODB.Recordset
Dim pOldCompcode As String
Dim pStbMno As String
Dim pICCMno As String
Dim pCustName As String
Dim strBugStep As String

On Error GoTo ChkErr

CallRentToSale = False
'' // 取得舊公司別代碼

strBugStep = 1 ''  記錄目前程式碼執行的位置，以了解出錯的地方
'' ------------------------------------------------------------------------------------------------------------

rsOldCompCode.CursorLocation = adUseClient
rsOldCompCode.Open _
        "SELECT OLDCOMPCODE FROM " & GetOwner & "CD039 WHERE CODENO  = " & pcompcode, _
        gcnGi, adOpenKeyset, adLockReadOnly
If Not (rsOldCompCode.EOF Or rsOldCompCode.BOF) Then
     pOldCompcode = rsOldCompCode(0) & ""
 Else
    pOldCompcode = ""
End If

rsOldCompCode.Close
Set rsOldCompCode = Nothing

'' // 取得轉換的單據編號
strBugStep = 2 ''  記錄目前程式碼執行的位置，以了解出錯的地方
'' ------------------------------------------------------------------------------------------------------------

Dim strm As String
Select Case Mid(pBillNO, 5, 1)
  Case 0
    strm = Mid(pBillNO, 6, 1)
  Case 1
    Select Case Mid(pBillNO, 5, 2)
        Case 0
           strm = "A"
        Case 1
           strm = "B"
        Case 2
           strm = "C"
    End Select
End Select
pBillNO = Mid(pBillNO, 3, 2) & strm & Right(pBillNO, 9)
strSQL = " SELECT FROM " & GetOwner & "SOAC0201A," & GetOwner & "SOAC0201A WHERE  SOAC0201A. = SOAC0201B"

''//  取得 STB 的料號 從 SOAC0201A
strBugStep = 3 ''  記錄目前程式碼執行的位置，以了解出錯的地方
'' ------------------------------------------------------------------------------------------------------------

rsStbIcc.CursorLocation = adUseClient
rsStbIcc.Open _
    "SELECT MaterialNo FROM " & GetOwner & "SOAC0201A WHERE FACISNO = '" & pStbSno & "'", _
    gcnGi, adOpenKeyset, adLockReadOnly
If Not (rsStbIcc.EOF Or rsStbIcc.BOF) Then
     pStbMno = rsStbIcc(0) & ""
Else
     pStbMno = ""
End If
  '''pStbMno = rsStbIcc(0)
rsStbIcc.Close
''//  取得 ICC 的料號 從 SOAC0201B
strBugStep = 4 ''  記錄目前程式碼執行的位置，以了解出錯的地方
'' ------------------------------------------------------------------------------------------------------------

rsStbIcc.Open _
    "SELECT MaterialNo FROM " & GetOwner & "SOAC0201B WHERE FACISNO = '" & pICCSno & "'", _
    gcnGi, adOpenKeyset, adLockReadOnly
If Not (rsStbIcc.EOF Or rsStbIcc.BOF) Then
    pICCMno = rsStbIcc(0) & ""
Else
   pICCMno = ""
End If
rsStbIcc.Close
Set rsStbIcc = Nothing

''  取得客戶名稱
strBugStep = 5 ''  記錄目前程式碼執行的位置，以了解出錯的地方
'' ------------------------------------------------------------------------------------------------------------

pCustName = gcnGi.Execute("select custname from " & GetOwner & "so001 where custid =" & gCustId).GetString


arrRentToSaleM(0) = pOldCompcode
arrRentToSaleM(1) = pBillNO
arrRentToSaleM(2) = "0" & Format(Date, "eemmdd")
arrRentToSaleM(3) = garyGi(0)

arrRentToSaleD(0, 0) = pStbMno & ""
arrRentToSaleD(0, 1) = pStbSno & ""
arrRentToSaleD(0, 2) = pShouldAmt

arrRentToSaleD(1, 0) = pICCMno & ""
arrRentToSaleD(1, 1) = pICCSno & ""
arrRentToSaleD(1, 2) = 0
        
''     呼叫jacky所寫的模組，將這些陣列傳入給他的 COMMAND
Dim O As Object
Dim intLinkMss As Integer
Dim strMessage As String
intLinkMss = gcnGi.Execute("SELECT LINKMSS FROM " & GetOwner & "CD039 WHERE CODENO  = " & pcompcode).GetString
If intLinkMss = 1 Then
    Dim cnComm As New ADODB.Connection
    Dim lngSeqNo As Long
    Dim blnMSSSynch As Boolean
        '直接與東富介接
      blnMSSSynch = GetRsValue("Select AcrossMSS From " & GetOwner & "Cd039 Where CodeNo = " & gCompCode) = 1
      If Not blnMSSSynch Then
          cnComm.Open GetCommonConnection
          lngSeqNo = GetInvoiceNo3("Send_Mss", cnComm)
      Else
              '' 以下這一行  2003/09/23 加入的
         cnComm.Open gcnGi.ConnectionString
          frmSO1147A.pic1.Visible = True
      End If
      
      If MSSRentToSale(O, arrRentToSaleM, arrRentToSaleD, strMessage, lngSeqNo, cnComm) = False Then
            Call MsgBox(strMessage, vbOKOnly + vbCritical, "訊息 ")
            frmSO1147A.pic1.Visible = False
            Exit Function
      End If
      frmSO1147A.pic1.Visible = False
End If
''"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
'\\ log 一筆設備異動資料至 SO055
strBugStep = 6 ''  記錄目前程式碼執行的位置，以了解出錯的地方
'' ------------------------------------------------------------------------------------------------------------

''"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
    Dim stInsertInto As String
    Dim rsSO004 As New ADODB.Recordset
    
    rsSO004.CursorLocation = adUseClient
    rsSO004.Open _
         "SELECT  * FROM " & GetOwner & "SO004 WHERE   SO004.SeqNo ='" & pSeqSno & "'", _
         gcnGi, adOpenKeyset, adLockReadOnly
    
    stInsertInto = "INSERT INTO " & GetOwner & "SO055 " & _
                          "(UPDEN,UPDTIME,CUSTID,CUSTNAme,COMPCODE," & _
                          "FUNCTYPE,BuyCodeB,PRDATEB,PREN1B,PRNAME1B,FACICODE," & _
                          "FACINAME,FACISNO,CVTID,BUYNAME,DUEDATE," & _
                          "AMOUNT,QUANTITY,UNITPRICE,INSTDATE,INSTEN1," & _
                          "INSTEN2,INSTNAME1,INSTNAME2,PREN2,PRNAME2," & _
                          "MEDIACODE,MEDIANAME,INTROID, INTRONAME,SMARTCARDNO )" & _
                          "VALUES" & _
                          "('" & garyGi(0) & "','" & Format(Date, "EEMMDD") & "'," & gCustId & " ,'" & pCustName & "'," & pcompcode & ",1,1," & _
                          GetNullString(rsSO004("PRDATE"), giDateV) & ",'" & _
                          garyGi(0) & "','" & garyGi(1) & _
                          "'," & GetNullString(rsSO004("FACICODE"), giLongV) & "," & GetNullString(rsSO004("FACINAME"), giStringV) & _
                          "," & GetNullString(rsSO004("FACISNO"), giStringV) & "," & GetNullString(rsSO004("CVTID"), giLongV) & ",'賣'" & _
                          "," & GetNullString(rsSO004("DUEDATE"), giDateV) & "," & GetNullString(rsSO004("AMOUNT"), giLongV) & "," & GetNullString(rsSO004("QUANTITY"), giLongV) & _
                          "," & GetNullString(rsSO004("UNITPRICE"), giLongV) & "," & GetNullString(rsSO004("INSTDATE"), giDateV) & "," & GetNullString(rsSO004("INSTEN1"), giStringV) & _
                          "," & GetNullString(rsSO004("INSTEN2"), giStringV) & "," & GetNullString(rsSO004("INSTNAME1"), giStringV) & "," & GetNullString(rsSO004("INSTNAME2"), giStringV) & _
                          "," & GetNullString(rsSO004("PREN2"), giStringV) & "," & GetNullString(rsSO004("PRNAME2"), giStringV) & "," & GetNullString(rsSO004("MEDIACODE"), giLongV) & _
                          "," & GetNullString(rsSO004("MEDIANAME"), giStringV) & "," & GetNullString(rsSO004("INTROID"), giStringV) & "," & GetNullString(rsSO004("INTRONAME"), giStringV) & _
                          "," & GetNullString(rsSO004("SMARTCARDNO"), giStringV) & ")"
    
    gcnGi.Execute stInsertInto
''"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
'\\ 修改該筆設備資料
strBugStep = 7 ''  記錄目前程式碼執行的位置，以了解出錯的地方
'' ------------------------------------------------------------------------------------------------------------

''"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
 Dim strUpdate  As String
Dim strBuyName As String

strBuyName = gcnGi.Execute("SELECT DESCRIPTION from " & GetOwner & "CD034 WHERE CODENO  =1").GetString
strUpdate = "UPDATE " & GetOwner & "SO004 SET  Deposit = 0 ," & _
                   "BUYCODE =1 ,BUYNAME  = '" & strBuyName & "',DUEDATE =  " & _
                   "TO_DATE('" & CStr(Date) & "','YYYY/MM/DD')," & _
                   "UPDEN='" & garyGi(0) & "', " & _
                   "UPDTime='" & GetDTString(Now) & "' ,Amount = " & pShouldAmt & "," & "UnitPrice = " & pShouldAmt & "," & "Quantity = 1 " & _
                   " WHERE SEQNO = '" & pSeqSno & "'  AND PRDATE IS  NULL"
gcnGi.Execute strUpdate

strBugStep = 8 ''  記錄目前程式碼執行的位置，以了解出錯的地方
'' ------------------------------------------------------------------------------------------------------------

 strUpdate = "UPDATE " & GetOwner & "SO004 SET  Deposit = 0," & _
                   "BUYCODE =1 ,BUYNAME  = '" & strBuyName & "',DUEDATE =  " & _
                   "TO_DATE('" & CStr(Date) & "','YYYY/MM/DD')," & _
                   "UPDEN='" & garyGi(0) & "', " & _
                   "UPDTime='" & GetDTString(Now) & "' ,Amount = " & pShouldAmt & "," & "UnitPrice = " & pShouldAmt & "," & "Quantity = 1 " & _
                   " WHERE FACISNO = '" & rsSO004("SmartCardNo") & "'" & _
                   " AND CUSTID =" & rsSO004("CustId") & "  AND PRDATE IS  NULL"
                  
gcnGi.Execute strUpdate

rsSO004.Close
Set rsSO004 = Nothing

CallRentToSale = True
Exit Function
ChkErr:
    Call ErrSub("EmcStb", "CallRentToSale" & "  位置碼[" & strBugStep & "]")
End Function
Private Function GetInvoiceNo3(StrTableName As String, Optional Conn As ADODB.Connection) As String
  On Error GoTo ChkErr
    Dim strSeq As String
    Dim strInv
    strSeq = "S_" & StrTableName
    If Conn Is Nothing Then Set Conn = gcnGi
    strInv = GetRsValue("SELECT Ltrim(To_Char(" & strSeq & ".NextVal, '09999999')) FROM " & GetOwner & "Dual", Conn)
    GetInvoiceNo3 = strInv
  Exit Function
ChkErr:
  ErrSub "Sys_Lib", "GetInvoiceNo3"
End Function
