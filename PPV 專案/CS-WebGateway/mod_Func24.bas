Attribute VB_Name = "mod_Func24"
Option Explicit

' Date : 2006/12/25 ¸t½Ï¸`
' Designer : Power Hammer
' Doc : Web API»¡©ú_20061225_Jacy.doc

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
            strErr = "µLªk³s½u¸ê®Æ®w!"
            JustDoIt = "-99,¸ê®Æ®w³s½u¥¢±Ñ!"
            GoTo 66
        End If
    End If
    
'    ETHOME -CM¬ÛÃö¸ê°T¬d¸ß
'    ¶Ç¤J°Ñ¼Æ[24, ¥Ó½Ð¤H©m¦W, ¨­¤ÀºØÃþ ( 1:¨­¤ÀÃÒ;  2:²Î¤@½s¸¹), ¨­¥÷ÃÒ¦r¸¹ ©Î ²Î¤@½s¸¹,
'    ³]³ÆºØÃþ (1: MAC; 2:±b¸¹) ,  CM MAC ©Î ACS ±b¸¹]

'    If RecordSetµ§¼Æ = 0 Then
'       ¦^¶Ç(-20, ¬dµL²Å¦X±ø¥ó¤§¸ê®Æ)
'    End If
'    If RecordSetµ§¼Æ > 0 Then
'      ²£¥ÍXMLÀÉ.(¨Ìselect¥X¨Ó¤§¸ê®Æ)    (°Ñ¦Ò: µù¤G)
'      ¦^¶Ç(0, ¦¨¥\)¤ÎXMLÀÉ
'    Else
'       ¦^¶Ç(-20, ¬dµL²Å¦X±ø¥ó¤§¸ê®Æ)
'    End If

    varPara = Split(strPara, ",")
    
    If UBound(varPara) <> 5 Then '    If ¶Ç¤J°Ñ¼Æ <> 6 Then
        JustDoIt = "-99,°Ñ¼Æ¤£¨¬!" '       ¦^¶Ç(-99, °Ñ¼Æ¤£¨¬)
        GoTo 66
    End If '    End If

'    0,¦¨¥\
'    <?xml version="1.0" encoding="BIG5" standalone="yes" ?>
'    <DataSet>
'      <DataTable>
'        <DataRow SYSID="9" SYSNAME="¶§©ú¤s" CUSTID="2" DeclarantName="¤ý¤jªL" DialAccount ="a1234" FACISNO="009096D5CC2A" ID="A123456789" Instaddress="¥x¥_¥«¤¤¥¿¸ô100¸¹" ChargeAddress="¥x¥_¥«¤¤¥¿¸ô104¸¹" MailAddress="¥x¥_¥«¤¤¥¿¸ô104¸¹" ContTel="87910168" Contmobile="0935123456" ContTel2="87642312" AgentTel="78324764" OPENSTATUS="¶}¾÷" ACCOUNTSTATUS="³n¶}" BPNAME="CM-­û¤u¤è®×:2M¥b¦~Ãº1800¤¸" DYNIPCOUNT="4"/>
'      </DataTable>
'    </DataSet>ÿ

    If varPara(4) = 1 Then
    '    1.­Y³]³ÆºØÃþ=1.MAC:
    '     ¥ý¨ìCATVNªºCOM°Ï,¥HMAC¬dSO306ªº¤½¥q§O,
    '     ¦A¥H¸Ó¤½¥q§O¹ïÀ³¨ìCOM°ÏªºSO507 , ¬d®Ö¦p¦ó³s¨ìSOªº¸ê°T
    '     ³s¨ì¸ÓSO , ¥H¸Ó¥Ó½Ð¤H©m¦W, ¨­¥÷ÃÒ¦r¸¹ / ²Î¤@½s¸¹, MAC¬d®Ö '¥¿±`' CM³]³Æ
        strDbg = "SELECT COMPCODE FROM COM.SO306 WHERE HFCMAC='" & varPara(5) & "'"
        strCompCode = GetValue(cn, strDbg) & Empty
                                                
    Else
    '    2. ­Y³]³ÆºØÃþ=2.±b¸¹:
    '     ¥ý¨ìCATVNªºCOM°Ï,¥HMAC¬dSO305Aªº¤½¥q§O(·s¼WTABLE: SO305A)
    '     (ª`·N­º¦¸¥ÑPM¨ÓINSERT,«áÄò¥Ñ¤p°¨¨C¦¸¶×¦U¨t²Î¥xªº±b¸¹®É,¦P®É¶×¤@¥÷¦Ü¸ÓTABLE),
    '     ¦A¥H¸Ó¤½¥q§O¹ïÀ³¨ìCOM°ÏªºSO507 , ¬d®Ö¦p¦ó³s¨ìSOªº¸ê°T
    '     ³s¨ì¸ÓSO , ¥H¸Ó¥Ó½Ð¤H©m¦W, ¨­¥÷ÃÒ¦r¸¹ / ²Î¤@½s¸¹, ±b¸¹¬d®Ö '¥¿±`'CM³]³Æ
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

'    ¶Ç¤J°Ñ¼Æ[24, ¥Ó½Ð¤H©m¦W, ¨­¤ÀºØÃþ ( 1:¨­¤ÀÃÒ;  2:²Î¤@½s¸¹), ¨­¥÷ÃÒ¦r¸¹ ©Î ²Î¤@½s¸¹,
'    ³]³ÆºØÃþ (1: MAC; 2:±b¸¹) ,  CM MAC ©Î ACS ±b¸¹]
            
'    »Ý=¶Ç¤Jªº"¥Ó½Ð¤H©m¦W"
'    ­Y³]³ÆºØÃþ=2(±b¸¹) , »Ý=¶Ç¤Jªº"±b¸¹"
'    ­Y³]³ÆºØÃþ=1(MAC) , »Ý=¶Ç¤Jªº"MAC"
'    ·s¼WÄæ¦ì SO004.ContMobile

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
    
    JustDoIt = "0,¦¨¥\" & vbCrLf & _
                        "<?xml version=""1.0"" encoding=""BIG5"" standalone=""yes"" ?>" & vbCrLf & _
                        "<DataSet>" & vbCrLf & _
                        "  <DataTable>" & vbCrLf & _
                        strXMLresult & _
                        "  </DataTable>" & vbCrLf & _
                        "</DataSet>" & vbCrLf ' ¦^¶Ç(0, ¦¨¥\) ¤ÎXMLÀÉ
    GoTo 66

'    µù¤G :  (API-24: ETHOME-CM¬ÛÃö¸ê°T¬d¸ß)

'   0:  ¦¨¥\
'   -99 : °Ñ¼Æ¤£¨¬
'   -20, ¬dµL²Å¦X±ø¥ó¤§¸ê®Æ
    
'    PS:
'        XML (¨t²Î¥x¥N¸¹, ¨t²Î¥x¦WºÙ, «È½s, ¥Î¤á¦WºÙ, ACS¼·±µ±b¸¹, CM MAC, ¨­¥÷ÃÒ¦r¸¹/²Î½s, ¸Ë¾÷¦a§}, ¦¬¶O¦a§}, ¶l±H¦a§}, ¥Ó½Ð¤HÁpµ¸¹q¸Ü, ¥Ó½Ð¤H¦æ°Ê¹q¸Ü, Ápµ¸¤HÁpµ¸¹q¸Ü, ¥N²z¤HÁpµ¸¹q¸Ü, ¶}³qª¬ºA,±b¸¹ª¬ºA,¥Ó½Ð¤è®×§O, IP¼Æ)
    
99:
    JustDoIt = "-20, ¬dµL²Å¦X±ø¥ó¤§¸ê®Æ!" ' -20, ¬dµL²Å¦X±ø¥ó¤§¸ê®Æ
    'JustDoIt = "-20, ¬dµL²Å¦X±ø¥ó¤§¸ê®Æ!" & vbCrLf & strDbg ' -20, ¬dµL²Å¦X±ø¥ó¤§¸ê®Æ
    
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

'' ¨­¥÷ÃÒ¦r¸¹ / ²Î½s(ID)
'Private Function ChkID(strIDKindCode As String, intRefNo As Integer) As Boolean
'  On Error GoTo ChkErr
'    ChkID = Val(GetValue(cn, "SELECT COUNT(*) FROM " & strOwner & "CD070" & strDBLink & _
'                                        " WHERE CODENO=" & strIDKindCode & _
'                                        " AND REFNO=" & intRefNo) & Empty) > 0
'
''    ¥HSO004.IDKindCode1 ¹ïÀ³¦ÜCD070.CODENO,
''    a.­Y¶Ç¤Jªº¨­¥÷ºØÃþ=1(¨­¥÷ÃÒ¦r¸¹) CD070.REFNO = 1
''    b.­Y¶Ç¤Jªº¨­¥÷ºØÃþ=2(²Î¤@½s¸¹) CD070.REFNO = 2
''    --CD070.REFNO»Ý²Å¦X,¤~¥i§ì¨úSO004.ID=¶Ç¤Jªº(¨­¥÷ÃÒ¦r¸¹/²Î¤@½s¸¹)
'
'  Exit Function
'ChkErr:
'    ErrHandle "mod_Func24", "ChkID"
'End Function

' ¨t²Î¥x¦WºÙ (SYSNAME)
Private Function GetCompName() As String
  On Error GoTo ChkErr
    ' ¥HSO004.COMPCODE¨ÓLINK
    strDbg = "SELECT DESCRIPTION FROM " & strOwner & "CD039" & strDBLink & _
                    " WHERE CODENO=" & strCompCode
    GetCompName = GetValue(cn, strDbg) & Empty
  Exit Function
ChkErr:
    ErrHandle "mod_Func24", "GetCompName"
End Function

' ¸Ë¾÷¦a§} (Instaddress)
Private Function GetInstAddr() As String
  On Error GoTo ChkErr
    ' ¥HSO004.CUSTID¨ÓLINK
    strDbg = "SELECT INSTADDRESS FROM " & strOwner & "SO001" & strDBLink & _
                    " WHERE CUSTID=" & strCustID & " AND COMPCODE=" & strCompCode
    GetInstAddr = GetValue(cn, strDbg) & Empty
  Exit Function
ChkErr:
    ErrHandle "mod_Func24", "GetInstAddr"
End Function

' ¦¬¶O¦a§} (ChargeAddress) / ¶l±H¦a§}(MailAddress)
Private Function GetAddress(strFaciSno As String, ByRef strChargeAddr As String, ByRef strMailAddr As String) As Boolean
  On Error GoTo ChkErr
    
    ' ¥HSO004.FACISNO¹ïÀ³¦ÜSO003.FACISNO
    ' ¥HSO003.ACCOUNTNO¹ïÀ³¦ÜSO002A.ACCOUNTNO
    
    ' ¦³±b¸¹: SO002A.ChargeAddress
    ' µL±b¸¹: SO001.ChargeAddress
    
    ' ¦³±b¸¹: SO002A.MailAddress
    ' µL±b¸¹: SO001.MailAddress

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

' IP¼Æ (DYNIPCOUNT)
Private Function GetDynIPcnt(strBPcode As String) As Integer
  On Error GoTo ChkErr
    ' ¥H SO004.BPCODE ¨ÓLINK ¦Ü CD078D.BPCODE
    '#3376 BpCode­n¦h¥[³æ¤Þ¸¹ By Kin 2007/07/26
    strDbg = "SELECT DYNIPCOUNT FROM " & strOwner & "CD078D" & strDBLink & _
                    " WHERE BPCODE='" & strBPcode & "'"
    GetDynIPcnt = Val(GetValue(cn, strDbg) & Empty)
  Exit Function
ChkErr:
    ErrHandle "mod_Func24", "GetDynIPcnt"
End Function

' ¶}³qª¬ºA (OPENSTATUS)
Private Function GetOpenStatus(strCMOpenDate As String, strCMCloseDate As String) As String
  On Error GoTo ChkErr
    ' §PÂ_SO004¥H¤U¤G­ÓÄæ¦ì:
    ' CM¶}¾÷¤é´ÁCMOpenDate (¶}¾÷)
    ' CMÃö¾÷¤é´ÁCMCloseDate (Ãö¾÷)
    ' ¥H¤é´Á³Ì¤jªºª¬ºAÅã¥Ü: ¶}¾÷ ©Î Ãö¾÷
    ' ­Y¨âªÌ¬ÒµL , Åã¥ÜªÅ
    strDbg = strCMOpenDate & vbTab & strCMCloseDate
    If strCMOpenDate = Empty And strCMCloseDate = Empty Then
        GetOpenStatus = ""
    Else
        If strCMOpenDate <> Empty And strCMCloseDate <> Empty Then
            GetOpenStatus = IIf(CDate(strCMOpenDate) >= CDate(strCMCloseDate), "¶}¾÷", "Ãö¾÷")
        Else
            If strCMOpenDate <> Empty And strCMCloseDate = Empty Then
                GetOpenStatus = "¶}¾÷"
            Else
                GetOpenStatus = "Ãö¾÷"
            End If
        End If
    End If
  Exit Function
ChkErr:
    ErrHandle "mod_Func24", "GetOpenStatus"
End Function

' ±b¸¹ª¬ºA (ACCOUNTSTATUS)
Private Function GetAccStatus(strEnableAcc As String, strDisableAcc As String) As String
  On Error GoTo ChkErr
    ' §PÂ_SO004¥H¤U¤G­ÓÄæ¦ì:
    ' ±b¸¹±Ò¥Î¤é´ÁEnableAccount (³n¶})
    ' ±b¸¹°±¥Î¤é´Ádisableaccount (³nÃö)
    ' ¥H¤é´Á³Ì¤jªºª¬ºAÅã¥Ü : ³n¶} ©Î ³nÃö
    ' ­Y¨âªÌ¬ÒµL , Åã¥ÜªÅ¥Õ
    strDbg = strEnableAcc & vbTab & strDisableAcc
    If strEnableAcc = Empty And strDisableAcc = Empty Then
        GetAccStatus = ""
    Else
        If strEnableAcc <> Empty And strDisableAcc <> Empty Then
            GetAccStatus = IIf(CDate(strEnableAcc) >= CDate(strDisableAcc), "³n¶}", "³nÃö")
        Else
            If strEnableAcc <> Empty And strDisableAcc = Empty Then
                GetAccStatus = "³n¶}"
            Else
                GetAccStatus = "³nÃö"
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
'    ¨t²Î¥x¥N¸¹, ¨t²Î¥x¦WºÙ, «È½s, ¥Î¤á¦WºÙ, ACS¼·±µ±b¸¹, CM MAC,
'    ¨­¥÷ÃÒ¦r¸¹/²Î½s, ¸Ë¾÷¦a§}, ¦¬¶O¦a§}, ¶l±H¦a§}, ¥Ó½Ð¤HÁpµ¸¹q¸Ü,
'    ¥Ó½Ð¤H¦æ°Ê¹q¸Ü , Ápµ¸¤HÁpµ¸¹q¸Ü, ¥N²z¤HÁpµ¸¹q¸Ü, ¶}³qª¬ºA, ±b¸¹ª¬ºA, ¥Ó½Ð¤è®×§O, IP¼Æ
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

