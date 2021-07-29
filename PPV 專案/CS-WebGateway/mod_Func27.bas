Attribute VB_Name = "mod_Func27"
'#3018 新增API27 By Kin 2007/09/04 Add
Option Explicit
Private cn As Object
Private strOwner As String
Private strXMLresult As String
Public Function JustDoIt(ByRef objWebPool As Object, _
                                        ByRef strPara As String) As String
  On Error GoTo ChkErr
  
    Dim strQry As String
    Dim varPara As Variant
    Dim rs As Object
    Dim rsSO033 As Object
    Dim objCnPool As Object
    Dim strCustID As String
    Dim strMediabillNo As String
    Dim strBillCloseDate As String
    Dim strWhere As String
    Dim lngTotalAmt As Long
    Dim strPostAccountNo As String
    Dim strCitiBankATM As String
    Dim strShopUnit As String, strBarCodeChar As String
    Dim strBarCode1 As String, strBarCode2 As String, strBarCode3 As String
    Dim strPostBillNo As String
    Dim blnBeingTrans As Boolean
    Set objCnPool = objWebPool
    
    Set cn = objCnPool.GetConnection
    
    If cn.State <= 0 Then
        Set cn = objCnPool.GetConnection
        If cn.State <= 0 Then
            strErr = "無法連線資料庫!"
            JustDoIt = "-99,資料庫連線失敗!"
            GoTo 66
        End If
    End If
    
    varPara = Split(strPara, ",")
    If UBound(varPara) < 3 Then
        JustDoIt = "-99,參數不足!"
        GoTo 66
    End If

    On Error Resume Next
    bytComp = Val(varPara(1))
    strCustID = varPara(2)
    strMediabillNo = varPara(3)
    
    On Error GoTo ChkErr
    strOwner = GetOwner(cn)
    
    If (Not IsNumeric(strCustID)) Or (strMediabillNo & "" = "") Then
            strErr = "傳入參數 [種類] 錯誤!!"
            JustDoIt = "-99,參數 [種類] 錯誤!"
        GoTo 66
    End If
    '#3018 常用的Where條件，先用變數儲存，後續可以直接使用 By Kin 2007/09/04
    strWhere = " SO033.MediaBillNo='" & strMediabillNo & "' And SO033.CustId=" & strCustID & _
               " And SO033.CompCode=" & bytComp & " And SO033.Uccode is Not null"
               
    '#3018 檢查是否有符合該媒体單號的應付帳款 By Kin 2007/09/04
    strQry = "Select Count(*) From " & strOwner & "SO033" & Get_DB_Link(cn) & _
            " Where " & strWhere
            
    If cn.Execute(strQry)(0) <= 0 Then
        JustDoIt = "-99,無任何帳務資料"
        GoTo 66
    Else
        JustDoIt = Empty
        
        cn.BeginTrans
        blnBeingTrans = True
        '*****************************************************************************************************************
        '#3018 找BillCloseDate最大值，如果是空值就找ShouldDate By Kin 2007/09/04
        strQry = "Select Max(ShouldDate) ShouldDate,Max(BillCloseDate) BillCloseDate From " & strOwner & "SO033" & _
                Get_DB_Link(cn) & " Where " & strWhere
        strQry = "Select Decode(BillCloseDate,Null,ShouldDate,BillCloseDate) From (" & strQry & ")"
        strBillCloseDate = cn.Execute(strQry)(0)
        If RightDate(cn) > strBillCloseDate Then
            JustDoIt = "-60,失敗-已大於繳款截止日或應收日"
            GoTo 66
        End If
        '*****************************************************************************************************************
        
        '#3018 計算總金額 By Kin 2007/09/04
        strQry = "Select Sum(ShouldAmt) From " & strOwner & "SO033" & Get_DB_Link(cn) & " Where " & strWhere
        lngTotalAmt = cn.Execute(strQry)(0)
        
        '********************************************************************************************************************************************************
        '#3018 Head與Detail所需的資料，將它存成一個rsSO033 By Kin 2007/09/04
        '#3018 測試不OK,這段找不到資料會出錯,如果無資料Show出無任何資料 By Kin 2007/11/06
        '#3018 測試不OK,多加條碼截止日(取得BarCode用) By Kin 2007/11/09
        strQry = "Select SO033.CustId,Decode(SO033.AccountNo,Null,SO001.CustName,SO002A.ChargeTitle) CustName,SO033.CitemCode,SO033.CitemName," & _
                         "SO033.MediaBillNo,SO001.InstAddress,SO004.EBTContNo,SO033.RealStartDate,SO033.RealStopDate,SO033.ShouldAmt,SO033.BarcodeCloseDate " & _
                 " From " & _
                            strOwner & "SO033" & Get_DB_Link(cn) & "," & strOwner & "SO001" & Get_DB_Link(cn) & "," & strOwner & "SO002A" & Get_DB_Link(cn) & _
                            "," & strOwner & "SO004" & Get_DB_Link(cn) & _
                 " Where " & strWhere & " And SO033.Custid=SO001.Custid And SO033.CompCode=SO001.CompCode And SO033.CustId=SO002A.CustId(+)" & _
                         " And SO033.AccountNo=SO002A.AccountNo(+) And SO033.CompCode=SO002A.CompCode(+) And SO033.FaciSeqNo=SO004.SeqNo(+)"
         Set rsSO033 = cn.Execute(strQry)
        '********************************************************************************************************************************************************
        If rsSO033.EOF Then
            JustDoIt = "-98,無任何帳務資料"
            GoTo 66
        End If
        '********************************************************************************************************************************************************
        '#3018 SO033明細的資料CitibankATM必須全部相同，如果有其中一筆不同，回傳XML時用空字串 By Kin 2007/09/04
        strQry = "Select CitibankATM From " & strOwner & "SO033" & Get_DB_Link(cn) & " Where " & strWhere & " Group By CitibankATM"
        Set rs = cn.Execute(strQry)
        If rs.RecordCount = 1 Then
            strCitiBankATM = rs("CitibankATM")
        End If
        '********************************************************************************************************************************************************
        
        '#3018 取出PostAccountNo,ShopUnit,BarCodeChar的值(用SO033.CitemCode 與CD068.CitemCodeStr關聯) By Kin 2007/09/04
        strQry = "Select PostAccountNo,ShopUnit,BarCodeChar From " & strOwner & "CD068" & Get_DB_Link(cn) & _
                 " Where InStr(',' || CitemCodeStr || ',','," & rsSO033("CitemCode") & ",')>0"
        Set rs = cn.Execute(strQry)
        If rs.EOF Then
            JustDoIt = "-99,收費項目對應不到"
            GoTo 66
        End If
        strPostAccountNo = rs("PostAccountNo") & ""
        strShopUnit = rs("ShopUnit") & ""
        strBarCodeChar = rs("BarCodeChar") & ""
        If strPostAccountNo = "" Or strShopUnit = "" Or strBarCodeChar = "" Then
            JustDoIt = "-99,收費項目取值錯誤"
            GoTo 66
        End If
        '#3018  取得BarCode By Kin 2007/09/04
        '#3018 測試不OK,條碼產生改為條碼截止日,所以先判斷截止日是否為空白 By Kin 2007/11/09
        If rsSO033("BarcodeCloseDate") & "" = "" Then
            JustDoIt = "-99,條碼截止日為空值"
            GoTo 66
        End If
        '#3018 取得條碼原本是用BillCloseDate,現在改為BarCodeCloseDate By Kin 2007/11/09
        If Not GetBarcode(rsSO033("BarcodeCloseDate") & "", strShopUnit, strMediabillNo, strBarCodeChar, _
                          lngTotalAmt, strBarCode1, strBarCode2, strBarCode3) Then Exit Function
        
        strPostBillNo = strShopUnit & Mid(strBarCode2, 1, 11)
        
        
        MakeHeadXML "1", strCustID, rsSO033("CustName"), strMediabillNo, _
                     rsSO033("InstAddress"), CStr(lngTotalAmt), _
                    strBillCloseDate, strCitiBankATM, strPostAccountNo, strPostBillNo, _
                    strBarCode1, strBarCode2, strBarCode3
        
        rsSO033.MoveFirst
        Do While Not rsSO033.EOF
            '#3018測試不OK,RealStopDate與RealStartDate有可能為Null,所以要多加"" By Kin 2007/11/09
            MakeDetailXML "2", rsSO033("CitemName") & "", rsSO033("EBTContNo") & "", _
                               rsSO033("RealStartDate") & "", rsSO033("RealStopDate") & "", rsSO033("ShouldAmt") & ""
            rsSO033.MoveNext
        Loop
        strQry = "Update SO033 Set PrtCount = SO033.PrtCount + 1 Where " & strWhere
        cn.Execute strQry
        JustDoIt = "0,成功" & vbCrLf & _
                        "<?xml version=""1.0"" encoding=""BIG5"" standalone=""yes"" ?>" & vbCrLf & _
                        "<DataSet>" & vbCrLf & _
                        "  <DataTable>" & vbCrLf & _
                        strXMLresult & _
                        "  </DataTable>" & vbCrLf & _
                        "</DataSet>" & vbCrLf '
        Debug.Print JustDoIt
        cn.CommitTrans
        blnBeingTrans = False
        
    End If

66:
    On Error Resume Next
    Rlx varPara
    Rlx strCustID
    Rlx strMediabillNo
    Rlx strBillCloseDate
    Rlx strCitiBankATM
    Rlx lngTotalAmt
    Rlx strShopUnit
    Rlx strBarCodeChar
    Rlx strBarCode1
    Rlx strBarCode2
    Rlx strBarCode3
    Rlx strPostBillNo
    If blnBeingTrans Then cn.RollbackTrans
    Set cn = Nothing
    Set objCnPool = Nothing
  Exit Function

ChkErr:
    ErrHandle "mod_Func27", "JustDoIt"
    If blnBeingTrans Then cn.RollbackTrans
    JustDoIt = "-99," & strErr
End Function

Private Function GetBarcode(ByVal strBillCloseDate, _
                            ByVal strShopUnit As String, ByVal strMediabillNo, _
                            ByVal strBarCodeChar As String, ByVal lngAmt As Long, _
                            strBarCode1 As String, strBarCode2 As String, strBarCode3 As String) As Boolean
  On Error GoTo ChkErr
    Dim strBarCode As String
    Dim i As Long
    Dim strChr As String
    Dim intAsc As Integer
    Dim intChkNo1 As Integer
    Dim intChkNo2 As Integer
    Dim strChkCode1 As String
    Dim strChkCode2 As String
    strBarCode1 = Format(strBillCloseDate, "EEMMDD") & strShopUnit
    strBarCode2 = strMediabillNo & strBarCodeChar
    strBarCode = strBarCode1 & "0" & strBarCode2 & Format(RightDate(cn), "EEMM") & "**" & _
                 Format(lngAmt, String(9, "0"))
    For i = 1 To 41
        strChr = Mid(strBarCode, i, 1)
        intAsc = Asc(strChr)
        Select Case True
            Case intAsc >= 65 And intAsc <= 73
               strChr = Chr(intAsc - 16)
            Case intAsc >= 74 And intAsc <= 82
               strChr = Chr(intAsc - 25)
            Case intAsc >= 83 And intAsc <= 90
               strChr = Chr(intAsc - 33)
        End Select
        If i Mod 2 <> 0 Then
            intChkNo1 = intChkNo1 + Val(strChr)
        Else
            intChkNo2 = intChkNo2 + Val(strChr)
        End If
    Next i
    strChkCode1 = CStr(intChkNo1 Mod 11)
    strChkCode2 = CStr(intChkNo2 Mod 11)
    If strChkCode1 = "0" Then strChkCode1 = "A"
    If strChkCode1 = "10" Then strChkCode1 = "B"
    
    If strChkCode2 = "0" Then strChkCode2 = "X"
    If strChkCode2 = "10" Then strChkCode2 = "Y"

    strBarCode3 = Format(RightDate(cn), "EEMM") & _
                           strChkCode1 & strChkCode2 & _
                           Format(lngAmt, String(9, "0"))
    GetBarcode = True
    Exit Function
ChkErr:
    ErrHandle "mod_Funct27", "JustDoIt"
End Function
Private Sub MakeHeadXML(strType As String, strCustID As String, _
                                    strCustName As String, strMediabillNo As String, _
                                    strInstAddress As String, strTotalAMT As String, _
                                    strBillCloseDate As String, strCitiBankATM As String, _
                                    strPostAccountNo As String, strPostBillNo As String, _
                                    strBarCode1 As String, strBarCode2 As String, _
                                    strBarCode3 As String)
  On Error GoTo ChkErr
    strXMLresult = strXMLresult & _
                            "    <DataRow TYPE=""" & strType & """" & _
                                " CUSTID=""" & strCustID & """" & _
                                " CUSTNAME=""" & strCustName & """" & _
                                " MEDIABILLNO=""" & strMediabillNo & """" & _
                                " INSTADDRESS=""" & strInstAddress & """" & _
                                " TOTALAMT=""" & strTotalAMT & """" & _
                                " BILLCLOSEDATE=""" & strBillCloseDate & """" & _
                                " CITIBANKATM=""" & strCitiBankATM & """" & _
                                " POSTACCOUNTNO=""" & strPostAccountNo & """" & _
                                " POSTBILLNO=""" & strPostBillNo & """" & _
                                " BARCODE1=""" & strBarCode1 & """" & _
                                " BARCODE2=""" & strBarCode2 & """" & _
                                " BARCODE3=""" & strBarCode3 & """" & _
                                "/>" & vbCrLf
                                
  Exit Sub
ChkErr:
    ErrHandle "mod_Func27", "MakeHeadXML"
End Sub

Private Sub MakeDetailXML(strType As String, strCitemName As String, _
                                    strEBTContNo As String, strRealStartDate As String, _
                                    strRealStopDate As String, strShouldAmt As String)

                                    
  On Error GoTo ChkErr
    strXMLresult = strXMLresult & _
                            "    <DataRow TYPE=""" & strType & """" & _
                                " CITEMNAME=""" & strCitemName & """" & _
                                " EBTCONTNO=""" & strEBTContNo & """" & _
                                " REALSTARTDATE=""" & strRealStartDate & """" & _
                                " REALSTOPDATE=""" & strRealStopDate & """" & _
                                " SHOULDAMT=""" & strShouldAmt & """" & _
                                "/>" & vbCrLf
                                
  Exit Sub
ChkErr:
    ErrHandle "mod_Func27", "MakeDetailXML"
End Sub


