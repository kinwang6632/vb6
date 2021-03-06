Attribute VB_Name = "mod_Func24"
Option Explicit

' Date : 2006/12/25 聖誕節
' Designer : Power Hammer
' Doc : Web API說明_20061225_Jacy.doc

Private cn As Object
Private strXMLresult As String
Private strCustID As String
Private strCompCode As String
Private strOwner As String
Private strDBLink As String
Private strDbg As String

Public Function JustDoIt(ByRef objWebPool As Object, _
                                        ByRef strPara As String) As String
  On Error GoTo ChkErr
    Dim varAry As Variant
    Dim varPara As Variant
    Dim varElement As Variant
    Dim objCnPool As Object
    Dim rs As Object
    Dim strQry As String
    Dim strChargeAddr As String
    Dim strMailAddr As String
    
    Set objCnPool = objWebPool
    Set cn = objCnPool.GetConnection
    
    If cn.state <= 0 Then
        Set cn = objCnPool.GetConnection
        If cn.state <= 0 Then
            strErr = "無法連線資料庫!"
            JustDoIt = "-99,資料庫連線失敗!"
            GoTo 66
        End If
    End If
    
'    ETHOME -CM相關資訊查詢
'    傳入參數[24, 申請人姓名, 身分種類 ( 1:身分證;  2:統一編號), 身份證字號 或 統一編號,
'    設備種類 (1: MAC; 2:帳號) ,  CM MAC 或 ACS 帳號]

'    If RecordSet筆數 = 0 Then
'       回傳(-20, 查無符合條件之資料)
'    End If
'    If RecordSet筆數 > 0 Then
'      產生XML檔.(依select出來之資料)    (參考: 註二)
'      回傳(0, 成功)及XML檔
'    Else
'       回傳(-20, 查無符合條件之資料)
'    End If

    varPara = Split(strPara, ",")
    
    If UBound(varPara) <> 5 Then '    If 傳入參數 <> 6 Then
        JustDoIt = "-99,參數不足!" '       回傳(-99, 參數不足)
        GoTo 66
    End If '    End If

'    0,成功
'    <?xml version="1.0" encoding="BIG5" standalone="yes" ?>
'    <DataSet>
'      <DataTable>
'        <DataRow SYSID="9" SYSNAME="陽明山" CUSTID="2" DeclarantName="王大林" DialAccount ="a1234" FACISNO="009096D5CC2A" ID="A123456789" Instaddress="台北市中正路100號" ChargeAddress="台北市中正路104號" MailAddress="台北市中正路104號" ContTel="87910168" Contmobile="0935123456" ContTel2="87642312" AgentTel="78324764" OPENSTATUS="開機" ACCOUNTSTATUS="軟開" BPNAME="CM-員工方案:2M半年繳1800元" DYNIPCOUNT="4"/>
'      </DataTable>
'    </DataSet>?

    If varPara(4) = 1 Then
    '    1.若設備種類=1.MAC:
    '     先到CATVN的COM區,以MAC查SO306的公司別,
    '     再以該公司別對應到COM區的SO507 , 查核如何連到SO的資訊
    '     連到該SO , 以該申請人姓名, 身份證字號 / 統一編號, MAC查核 '正常' CM設備
        strDbg = "SELECT COMPCODE FROM COM.SO306 WHERE HFCMAC='" & varPara(5) & "'"
        strCompCode = GetValue(cn, strDbg) & Empty
                                                
    Else
    '    2. 若設備種類=2.帳號:
    '     先到CATVN的COM區,以MAC查SO305A的公司別(新增TABLE: SO305A)
    '     (注意首次由PM來INSERT,後續由小馬每次匯各系統台的帳號時,同時匯一份至該TABLE),
    '     再以該公司別對應到COM區的SO507 , 查核如何連到SO的資訊
    '     連到該SO , 以該申請人姓名, 身份證字號 / 統一編號, 帳號查核 '正常'CM設備
        strDbg = "SELECT COMPCODE FROM COM.SO305A WHERE DIALACCOUNT='" & varPara(5) & "'"
        strCompCode = GetValue(cn, strDbg) & Empty
    End If

    If strCompCode = Empty Then GoTo 99
    
    strDbg = "SELECT DESCRIPTION,LINK FROM COM.SO507 WHERE CODENO=" & strCompCode
    Set rs = cn.Execute(strDbg)
    
    If rs.EOF Then GoTo 99
    
    strOwner = Trim(rs("Description") & "")
    strDBLink = Trim(rs("Link") & "")
    
    If Len(strOwner) > 0 Then strOwner = strOwner & "."
    If Len(strDBLink) > 0 Then strDBLink = "@" & strDBLink

'    傳入參數[24, 申請人姓名, 身分種類 ( 1:身分證;  2:統一編號), 身份證字號 或 統一編號,
'    設備種類 (1: MAC; 2:帳號) ,  CM MAC 或 ACS 帳號]
            
'    需=傳入的"申請人姓名"
'    若設備種類=2(帳號) , 需=傳入的"帳號"
'    若設備種類=1(MAC) , 需=傳入的"MAC"
'    新增欄位 SO004.ContMobile

    strQry = "SELECT COMPCODE,CUSTID,DECLARANTNAME,DECLARANTNAME" & _
                    ",DIALACCOUNT,FACISNO,ID,ID2,CONTTEL,CONTMOBILE,CONTTEL2" & _
                    ",AGENTTEL,CMOPENDATE,CMCLOSEDATE,ENABLEACCOUNT" & _
                    ",DISABLEACCOUNT, BPCODE, BPNAME,IDKINDCODE1" & _
                    " FROM " & strOwner & "SO004" & strDBLink & _
                    " WHERE DECLARANTNAME='" & varPara(1) & "'" & _
                    " AND " & IIf(varPara(4) = 1, "FACISNO", "DIALACCOUNT") & "='" & varPara(5) & "'" & _
                    " AND ( PRDATE IS NULL OR INSTDATE > PRDATE)"
                    
    strDbg = strQry
    Set rs = cn.Execute(strQry)
    
    If rs.EOF Then GoTo 99
    
    'If ChkID(rs("IDKindCode1") & "", Val(varPara(2))) Then
    '    If varPara(2) & "" <> rs("ID") & "" Then GoTo 99
    'End If
    If varPara(3) & "" <> rs("ID") & "" And varPara(3) & "" <> rs("ID2") & "" Then GoTo 99
    
    While Not rs.EOF
        strCustID = rs("CustID")
        If Not GetAddress(rs("FaciSno") & "", strChargeAddr, strMailAddr) Then GoTo 99
        MakeXML strCompCode, GetCompName, strCustID, _
                        rs("DeclarantName") & "", rs("DialAccount") & "", _
                        rs("FaciSno") & "", varPara(3) & "", _
                        GetInstAddr, strChargeAddr, strMailAddr, _
                        rs("ContTel") & "", rs("ContMobile") & "", rs("ContTel2") & "", rs("AgentTel") & "", _
                        GetOpenStatus(rs("CMOpenDate") & "", rs("CMCloseDate") & ""), _
                        GetAccStatus(rs("EnableAccount") & "", rs("DisableAccount") & ""), _
                        rs("BPName") & "", GetDynIPcnt(rs("BPCode") & "")
        rs.MoveNext
    Wend
    
    JustDoIt = "0,成功" & vbCrLf & _
                        "<?xml version=""1.0"" encoding=""BIG5"" standalone=""yes"" ?>" & vbCrLf & _
                        "<DataSet>" & vbCrLf & _
                        "  <DataTable>" & vbCrLf & _
                        strXMLresult & _
                        "  </DataTable>" & vbCrLf & _
                        "</DataSet>" & vbCrLf ' 回傳(0, 成功) 及XML檔
    GoTo 66

'    註二 :  (API-24: ETHOME-CM相關資訊查詢)

'   0:  成功
'   -99 : 參數不足
'   -20, 查無符合條件之資料
    
'    PS:
'        XML (系統台代號, 系統台名稱, 客編, 用戶名稱, ACS撥接帳號, CM MAC, 身份證字號/統編, 裝機地址, 收費地址, 郵寄地址, 申請人聯絡電話, 申請人行動電話, 聯絡人聯絡電話, 代理人聯絡電話, 開通狀態,帳號狀態,申請方案別, IP數)
    
99:
    JustDoIt = "-20, 查無符合條件之資料!" ' -20, 查無符合條件之資料
    'JustDoIt = "-20, 查無符合條件之資料!" & vbCrLf & strDbg ' -20, 查無符合條件之資料
    
66:
    On Error Resume Next
    Rlx varAry
    Rlx varPara
    Rlx varElement
    Rlx strCustID
    Rlx strXMLresult
    Rlx strOwner
    Rlx strDBLink
    Rlx objCnPool
    Rlx strCompCode
    Rlx strChargeAddr
    Rlx strMailAddr
    Rlx strDbg
    Set cn = Nothing
    Set objCnPool = Nothing
  Exit Function
ChkErr:
    ErrHandle "mod_Func24", "JustDoIt"
    JustDoIt = "-99," & strErr & vbCrLf & strDbg
End Function

'' 身份證字號 / 統編(ID)
'Private Function ChkID(strIDKindCode As String, intRefNo As Integer) As Boolean
'  On Error GoTo ChkErr
'    ChkID = Val(GetValue(cn, "SELECT COUNT(*) FROM " & strOwner & "CD070" & strDBLink & _
'                                        " WHERE CODENO=" & strIDKindCode & _
'                                        " AND REFNO=" & intRefNo) & Empty) > 0
'
''    以SO004.IDKindCode1 對應至CD070.CODENO,
''    a.若傳入的身份種類=1(身份證字號) CD070.REFNO = 1
''    b.若傳入的身份種類=2(統一編號) CD070.REFNO = 2
''    --CD070.REFNO需符合,才可抓取SO004.ID=傳入的(身份證字號/統一編號)
'
'  Exit Function
'ChkErr:
'    ErrHandle "mod_Func24", "ChkID"
'End Function

' 系統台名稱 (SYSNAME)
Private Function GetCompName() As String
  On Error GoTo ChkErr
    ' 以SO004.COMPCODE來LINK
    strDbg = "SELECT DESCRIPTION FROM " & strOwner & "CD039" & strDBLink & _
                    " WHERE CODENO=" & strCompCode
    GetCompName = GetValue(cn, strDbg) & Empty
  Exit Function
ChkErr:
    ErrHandle "mod_Func24", "GetCompName"
End Function

' 裝機地址 (Instaddress)
Private Function GetInstAddr() As String
  On Error GoTo ChkErr
    ' 以SO004.CUSTID來LINK
    strDbg = "SELECT INSTADDRESS FROM " & strOwner & "SO001" & strDBLink & _
                    " WHERE CUSTID=" & strCustID & " AND COMPCODE=" & strCompCode
    GetInstAddr = GetValue(cn, strDbg) & Empty
  Exit Function
ChkErr:
    ErrHandle "mod_Func24", "GetInstAddr"
End Function

' 收費地址 (ChargeAddress) / 郵寄地址(MailAddress)
Private Function GetAddress(strFaciSno As String, ByRef strChargeAddr As String, ByRef strMailAddr As String) As Boolean
  On Error GoTo ChkErr
    
    ' 以SO004.FACISNO對應至SO003.FACISNO
    ' 以SO003.ACCOUNTNO對應至SO002A.ACCOUNTNO
    
    ' 有帳號: SO002A.ChargeAddress
    ' 無帳號: SO001.ChargeAddress
    
    ' 有帳號: SO002A.MailAddress
    ' 無帳號: SO001.MailAddress

    Dim strAccNo As String
    Dim rs2A As Object, rs01 As Object
    
    strDbg = "SELECT ACCOUNTNO FROM " & strOwner & "SO003" & strDBLink & _
                    " WHERE CUSTID=" & strCustID & _
                    " AND COMPCODE=" & strCompCode & _
                    " AND FACISNO='" & strFaciSno & "'"
    
    strAccNo = GetValue(cn, strDbg) & Empty
    
    If strAccNo = Empty Then
        GoTo 66
    Else
        strDbg = "SELECT ACCOUNTNO,CHARGEADDRESS,MAILADDRESS FROM " & strOwner & "SO002A" & strDBLink & _
                        " WHERE CUSTID=" & strCustID & _
                        " AND COMPCODE=" & strCompCode & _
                        " AND ACCOUNTNO='" & strAccNo & "'"
        
        Set rs2A = cn.Execute(strDbg)
        
        If Not rs2A.EOF Then
            strChargeAddr = rs2A("ChargeAddress") & ""
            strMailAddr = rs2A("MailAddress") & ""
            GetAddress = True
            GoTo 99
        End If
    End If
    
66:

    strDbg = "SELECT CHARGEADDRESS,MAILADDRESS FROM " & strOwner & "SO001" & strDBLink & _
                    " WHERE CUSTID=" & strCustID & " AND COMPCODE=" & strCompCode
    
    Set rs01 = cn.Execute(strDbg)
    
    If rs01.EOF Then
        GetAddress = False
    Else
        strChargeAddr = rs01("ChargeAddress") & ""
        strMailAddr = rs01("MailAddress") & ""
        GetAddress = True
    End If

99:

    On Error Resume Next
    rs2A.Close
    rs01.Close
    Set rs2A = Nothing
    Set rs01 = Nothing
    
  Exit Function
ChkErr:
    ErrHandle "mod_Func24", "GetAddress"
End Function

' IP數 (DYNIPCOUNT)
Private Function GetDynIPcnt(strBPcode As String) As Integer
  On Error GoTo ChkErr
    ' 以 SO004.BPCODE 來LINK 至 CD078D.BPCODE
    '#3376 BpCode要多加單引號 By Kin 2007/07/26
    strDbg = "SELECT DYNIPCOUNT FROM " & strOwner & "CD078D" & strDBLink & _
                    " WHERE BPCODE='" & strBPcode & "'"
    GetDynIPcnt = Val(GetValue(cn, strDbg) & Empty)
  Exit Function
ChkErr:
    ErrHandle "mod_Func24", "GetDynIPcnt"
End Function

' 開通狀態 (OPENSTATUS)
Private Function GetOpenStatus(strCMOpenDate As String, strCMCloseDate As String) As String
  On Error GoTo ChkErr
    ' 判斷SO004以下二個欄位:
    ' CM開機日期CMOpenDate (開機)
    ' CM關機日期CMCloseDate (關機)
    ' 以日期最大的狀態顯示: 開機 或 關機
    ' 若兩者皆無 , 顯示空
    strDbg = strCMOpenDate & vbTab & strCMCloseDate
    If strCMOpenDate = Empty And strCMCloseDate = Empty Then
        GetOpenStatus = ""
    Else
        If strCMOpenDate <> Empty And strCMCloseDate <> Empty Then
            GetOpenStatus = IIf(CDate(strCMOpenDate) >= CDate(strCMCloseDate), "開機", "關機")
        Else
            If strCMOpenDate <> Empty And strCMCloseDate = Empty Then
                GetOpenStatus = "開機"
            Else
                GetOpenStatus = "關機"
            End If
        End If
    End If
  Exit Function
ChkErr:
    ErrHandle "mod_Func24", "GetOpenStatus"
End Function

' 帳號狀態 (ACCOUNTSTATUS)
Private Function GetAccStatus(strEnableAcc As String, strDisableAcc As String) As String
  On Error GoTo ChkErr
    ' 判斷SO004以下二個欄位:
    ' 帳號啟用日期EnableAccount (軟開)
    ' 帳號停用日期disableaccount (軟關)
    ' 以日期最大的狀態顯示 : 軟開 或 軟關
    ' 若兩者皆無 , 顯示空白
    strDbg = strEnableAcc & vbTab & strDisableAcc
    If strEnableAcc = Empty And strDisableAcc = Empty Then
        GetAccStatus = ""
    Else
        If strEnableAcc <> Empty And strDisableAcc <> Empty Then
            GetAccStatus = IIf(CDate(strEnableAcc) >= CDate(strDisableAcc), "軟開", "軟關")
        Else
            If strEnableAcc <> Empty And strDisableAcc = Empty Then
                GetAccStatus = "軟開"
            Else
                GetAccStatus = "軟關"
            End If
        End If
    End If
  Exit Function
ChkErr:
    ErrHandle "mod_Func24", "GetAccStatus"
End Function

Private Sub MakeXML(strSysID As String, strSysName As String, _
                                    strCustID As String, strDeclarantName As String, _
                                    strDialAccount As String, strFaciSno As String, _
                                    strID As String, strInstAddress As String, _
                                    strChargeAddress As String, strMailAddress As String, _
                                    strContTel As String, strContMobile As String, _
                                    strContTel2 As String, strAgentTel As String, _
                                    strOpenStatus As String, strAccountStatus As String, _
                                    strBPName As String, strDynIPCount As String)
  On Error GoTo ChkErr
'    系統台代號, 系統台名稱, 客編, 用戶名稱, ACS撥接帳號, CM MAC,
'    身份證字號/統編, 裝機地址, 收費地址, 郵寄地址, 申請人聯絡電話,
'    申請人行動電話 , 聯絡人聯絡電話, 代理人聯絡電話, 開通狀態, 帳號狀態, 申請方案別, IP數
    strXMLresult = strXMLresult & _
                            "    <DataRow SYSID=""" & strSysID & """" & _
                                " SYSNAME=""" & strSysName & """" & _
                                " CUSTID=""" & strCustID & """" & _
                                " DECLARANTNAME=""" & strDeclarantName & """" & _
                                " DIALACCOUNT=""" & strDialAccount & """" & _
                                " FACISNO=""" & strFaciSno & """" & _
                                " ID=""" & strID & """" & _
                                " INSTADDRESS=""" & strInstAddress & """" & _
                                " CHARGEADDRESS=""" & strChargeAddress & """" & _
                                " MAILADDRESS=""" & strMailAddress & """" & _
                                " CONTTEL=""" & strContTel & """" & _
                                " CONTMOBILE=""" & strContMobile & """" & _
                                " CONTTEL2=""" & strContTel2 & """" & _
                                " AGENTTEL=""" & strAgentTel & """" & _
                                " OPENSTATUS=""" & strOpenStatus & """" & _
                                " ACCOUNTSTATUS=""" & strAccountStatus & """" & _
                                " BPNAME=""" & strBPName & """" & _
                                " DYNIPCOUNT=""" & strDynIPCount & """/>" & vbCrLf
  Exit Sub
ChkErr:
    ErrHandle "mod_Func24", "MakeXML"
End Sub

