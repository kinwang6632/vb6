Attribute VB_Name = "mod_Func9"
Option Explicit

Private cn As Object
Private strXMLresult As String
Private blnOneOfOK As Boolean

' 檢核處理中之 STB 的 iPPV 訂購權是否已開啟
Private Function IPPVok(strOwner As String, strSeqNo As String, Optional blnNoDataFound As Boolean = False) As Boolean
  On Error Resume Next
    IPPVok = (Val(RPxx(cn.Execute("SELECT IPPVRIGHT FROM " & _
                                                        strOwner & "SO004" & Get_DB_Link(cn) & _
                                                        "WHERE SEQNO='" & strSeqNo & "'").GetString(2, 1, "", "", "") & "")) = 1)
    If Err.Number = 3021 Then
        Err.Clear
        blnNoDataFound = True
        IPPVok = False
    ElseIf Err.Number <> 0 Then
        ErrHandle "mod_Func9", "IPPVok"
    End If
End Function

' 判斷 有/無 回傳機制
Private Function ReturnBack(strOwner As String, strSeqNo As String) As Boolean
  On Error GoTo ChkErr
    Dim lngModelCode As Long
    Dim strQry As String
    strQry = "SELECT IPPVCALLBACK FROM " & strOwner & "SO041" & Get_DB_Link(cn) & _
                    "WHERE ROWNUM=1"
    ReturnBack = (Val(GetValue(cn, strQry)) = 1)
    If ReturnBack Then
        strQry = "SELECT MODELCODE FROM " & strOwner & "SO004" & Get_DB_Link(cn) & _
                        "WHERE SEQNO='" & strSeqNo & "'"
        lngModelCode = Val(GetValue(cn, strQry))
        If lngModelCode > 0 Then
            strQry = "SELECT FUNCFLAG01 FROM " & strOwner & "CD043" & Get_DB_Link(cn) & _
                            "WHERE CODENO=" & lngModelCode
            ReturnBack = ReturnBack And (Val(GetValue(cn, strQry)) = 1)
        Else
            ReturnBack = False
        End If
    End If
  Exit Function
ChkErr:
    ErrHandle "mod_Func9", "ReturnBack"
End Function

Private Function Get_IPPV_Credit(strOwner As String, strSeqNo As String) As Integer
  On Error GoTo ChkErr
    Get_IPPV_Credit = Val(cn.Execute("SELECT IPPVCREDIT FROM " & strOwner & "SO004" & Get_DB_Link(cn) & _
                                                            "WHERE SEQNO='" & strSeqNo & "'").GetString(2, 1, "", "", 0))
  Exit Function
ChkErr:
    ErrHandle "mod_Func9", "Get_IPPV_Credit"
End Function

Public Function JustDoIt(ByRef objWebPool As Object, _
                                        ByRef strPara As String) As String
  On Error GoTo ChkErr
    Dim strQry As String
    Dim varPara As Variant
    Dim varData As Variant
    Dim objCnPool As Object
    Dim objIPPV As Object
    
    Dim rs As Object
    Dim lngCustID As Long
    Dim intLoop As Integer
    Dim strOwner As String
    Dim strAuthenticID As String
    Dim blnNoData As Boolean
    
'    Set objCnPool = objWebPool
'    Set cn = objCnPool.GetConnection
'
'    If cn.State <= 0 Then
'        Set cn = objCnPool.GetConnection
'        If cn.State <= 0 Then
'            strErr = "無法連線資料庫!"
'            JustDoIt = "-99,資料庫連線失敗!"
'            GoTo 66
'        End If
'    End If
    '************************************************************************
    '#3018 判斷傳入的個數是否正確 By Kin 2007/08/31
    varPara = Split(strPara, vbCrLf)
    If UBound(varPara) < 2 Then
        JustDoIt = "-99,參數不足!"
        GoTo 66
    End If
    For intLoop = 1 To UBound(varPara)
        If (UBound(Split(varPara(intLoop), ",")) < 3) And (varPara(intLoop) <> Empty) Then
            JustDoIt = "-99,參數不足!"
            GoTo 66
        End If
    Next intLoop
    varPara = Split(varPara(0), ",")
    If UBound(varPara) <> 2 Then
        JustDoIt = "-99,參數不足!"
        GoTo 66
    End If
    '************************************************************************
    strAuthenticID = varPara(2)
    bytComp = Val(Left(strAuthenticID, 3))
    lngCustID = Val(Right(strAuthenticID, 7))
    
    varData = Split(strPara, vbCrLf)
    
    Set cn = Get_General_CN2(bytComp)
    
    If cn.State <= 0 Then
        Set cn = Get_General_CN2(bytComp)
        If cn.State <= 0 Then
            strErr = "無法連線資料庫!"
            JustDoIt = "-99,資料庫連線失敗!"
            GoTo 66
        End If
    End If
    strOwner = GetOwner(cn)
    strQry = "SELECT CUSTID FROM " & strOwner & "SO002B" & Get_DB_Link(cn) & _
                    "WHERE MEMBERID='" & varPara(1) & "' AND AUTHENTICID='" & strAuthenticID & "'"
    
    lngCustID = Val(GetValue(cn, strQry) & "")
    
    If lngCustID = 0 Then
        JustDoIt = "-2,訂購PPV認證錯誤"
        GoTo 66
    End If
    
    Set objIPPV = CreateObject("setIPPV.clsNagraSTB")
    
    With objIPPV
            Set .uConn = Get_General_CN
            garyGi(1) = "WEB"
            .ugaryGi = Join(garyGi, Chr(9))
    End With
    
    ''依傳來參數呼叫CA命令點數下載模組 (STBCreDold)

    '點數線上訂購--下載點數
    '9、會員代號、認證ID(折行)
    '設備流水號、STB序號、ICC序號、點數(折行)
    '設備流水號、STB序號、ICC序號、點數(折行)
    '…
    '0 : 成功
    '-8 : 點數下載有誤
    '-19 未啟用IPPV訂購權
    
    ''Loop處理 (第二行參數開始)
    
    blnOneOfOK = False
    strXMLresult = ""
    
    For intLoop = 1 To UBound(varData)
        If RPxx(CStr(varData(intLoop))) <> Empty Then
            varPara = Split(varData(intLoop), ",")
            '檢核處理中之STB的iPPV訂購權是否已開啟(IPPVright=1 ?).
            ''If IPPVright <> 1 Then
                ''  呼叫CA命令非視覺化開啟iPPV訂購權模組 (EnIPPVright)
            ''End If
            
            JustDoIt = ""
            
            If Not IPPVok(strOwner, CStr(varPara(0)), blnNoData) Then
                ''啟動IPPV訂購權
                    'Public Function EnIPPVright(ByVal strCompCode As String, _
                    '    ByVal lngCustId As Long, ByVal strSmartCard As String, _
                    '    ByVal strResvTime As String, ByVal strSTB As String, _
                    '    ByVal strDoCaption As String, ByVal strFaciSNo As String) As Boolean
                    '   Format(Now, "YYYY/MM/DD HH:MM")
                If blnNoData Then
'                    JustDoIt = "-99,設備流水號不存在"
                    JustDoIt = "設備流水號不存在"
                Else
                    If Not objIPPV.EnIPPVright(CStr(bytComp), _
                                                        lngCustID, _
                                                        CStr(varPara(2)), _
                                                        "", _
                                                        CStr(varPara(1)), _
                                                        "", _
                                                        CStr(varPara(1))) Then
                        strErr = objIPPV.uMsgBack
'                        JustDoIt = "-19,未啟用IPPV訂購權"
                        JustDoIt = "未啟用IPPV訂購權 " & strErr
                    End If
                End If
            End If
            
            If JustDoIt <> Empty Then
                MakeXML CStr(varPara(0)), CStr(varPara(1)), CStr(varPara(2)), CStr(varPara(3)), JustDoIt
            Else
                With objIPPV
                    If Not objIPPV.STBCreDold(CStr(bytComp), _
                                                                lngCustID, _
                                                                CStr(varPara(2)), _
                                                                "", _
                                                                CStr(varPara(1)), _
                                                                "", _
                                                                "03", _
                                                                CStr(Val(varPara(3)) + Get_IPPV_Credit(strOwner, CStr(varPara(0)))), _
                                                                Format(Now, "YYYY/MM/DD HH:MM"), _
                                                                ReturnBack(strOwner, CStr(varPara(0)))) Then
                        strErr = .uMsgBack
                        JustDoIt = "點數下載有誤"
'                        JustDoIt = "-8,點數下載有誤"
                        MakeXML CStr(varPara(0)), CStr(varPara(1)), CStr(varPara(2)), CStr(varPara(3)), JustDoIt & " " & strErr
                    Else
                        MakeXML CStr(varPara(0)), CStr(varPara(1)), CStr(varPara(2)), CStr(varPara(3)), .uMsgBack
                        blnOneOfOK = True
    '                    If .uMsgBack <> Empty Then
    '                        JustDoIt = "0,成功(" & .uMsgBack & ")"
    '                    Else
    '                        JustDoIt = "0,成功"
    '                    End If
                    End If
                End With
            End If
            
            ''下載/設定IPPV點數or (無回傳)
            'Public Function STBCreDold(ByVal strCompCode As String, _
            '    ByVal lngCustId As Long, ByVal strSmartCard As String, _
            '    ByVal strResvTime As String, ByVal strSTB As String, _
            '    ByVal strDoCaption As String, ByVal strCreditMode As String, _
            '    ByVal strCredit As String, ByVal strCollectDate As String, ByVal blnReturnBk As Boolean) As Boolean
            
            '    寧可避開誘餌，也勝過在陷阱中掙扎。 說:
            '    ResvTime=''
            '    寧可避開誘餌，也勝過在陷阱中掙扎。 說:
            '    CreditMode='01'
            '    寧可避開誘餌，也勝過在陷阱中掙扎。 說:
            '   CollectDate = sysdate
            
            ''If 點數下載結果不ok Then
                ''  回傳(-8,點數下載有誤)
            ''Else
                ''  回傳(0,成功)
            ''End If
            
            ''呼叫CA命令非視覺化點數下載模組(STBCreDold[有回傳機制call]/ STBCreDoldNb[無回傳機制call])
            ''PS:
            ''有回傳機制定義 : so041. iPPVCallBack=1 and 該STB ModelCode 對應到 CD043. FuncFlag01=1
            ''無回傳機制定義 : so041. iPPVCallBack=0 or (so041. iPPVCallBack=1 and 該STB ModelCode對應到CD043. FuncFlag01=0)
        End If
    Next
    ''End Loop
    
    JustDoIt = IIf(blnOneOfOK, "0,成功", "-8,點數下載有誤")
    
    JustDoIt = JustDoIt & vbCrLf & _
                        "<?xml version=""1.0"" encoding=""BIG5"" standalone=""yes"" ?>" & vbCrLf & _
                        "<DataSet>" & vbCrLf & _
                        "  <DataTable>" & vbCrLf & _
                        strXMLresult & _
                        "  </DataTable>" & vbCrLf & _
                        "</DataSet>" & vbCrLf ' 回傳(0, 成功) 及XML檔
    
    '01=>ADD 增加 Credit
    '02=>SUBTRACT 減少 Credit
    '03=>SET CREDIT 直接設定 Credit 為新值
    '04=>SET BALANCE 清除 Debit, 直接設定 Credit 為新值
    '05=>SUB OFFSET 同時減少Debit 與Credit

66:
    On Error Resume Next
    Rlx rs
    Rlx strQry
    Rlx intLoop
    Rlx varData
    Rlx varPara
    Rlx objIPPV
    Rlx lngCustID
    Rlx blnNoData
    Rlx strAuthenticID
    Rlx strOwner
    Set cn = Nothing
    Set objCnPool = Nothing
  Exit Function
ChkErr:
    ErrHandle "mod_Func9", "JustDoIt"
    JustDoIt = "-99," & strErr
End Function

' --2006/07/30 Modify
'  產生XML檔.(依記錄傳送指令後回傳回來的訊息資料)
'  回傳(0,成功)及XML檔

Private Sub MakeXML(ByVal strSeqNo As String, _
                                    ByVal strFaciSno As String, _
                                    ByVal strSmartCardNo As String, _
                                    ByVal strEPGprice As String, _
                                    ByVal strStatus As String)
  On Error GoTo ChkErr
      ' 設備流水號、STB序號、ICC序號、點數(折行)
    strXMLresult = strXMLresult & _
                            "    <DataRow SEQNO=""" & strSeqNo & """" & _
                                " FACISNO=""" & strFaciSno & """" & _
                                " SMARTCARDNO=""" & strSmartCardNo & """" & _
                                " EPGPRICE=""" & strEPGprice & """" & _
                                " STATUS=""" & strStatus & """/>" & vbCrLf
  Exit Sub
ChkErr:
    ErrHandle "mod_Func9", "MakeXML"
End Sub


'Public Function JustDoIt(ByRef objWebPool As Object, _
'                                        ByRef strPara As String) As String
'  On Error GoTo ChkErr
'    Dim strQry As String
'    Dim varPara As Variant
'    Dim varData As Variant
'    Dim objCnPool As Object
'    Dim objIPPV As Object
'
'    Dim rs As Object
'    Dim lngCustID As Long
'    Dim intLoop As Integer
'    Dim strOwner As String
'    Dim strAuthenticID As String
'    Dim blnNoData As Boolean
'
'    Set objCnPool = objWebPool
'    Set cn = objCnPool.GetConnection
'
'    If cn.state <= 0 Then
'        Set cn = objCnPool.GetConnection
'        If cn.state <= 0 Then
'            strErr = "無法連線資料庫!"
'            JustDoIt = "-99,資料庫連線失敗!"
'            GoTo 66
'        End If
'    End If
'
'    varPara = Split(strPara, vbCrLf)
'    varPara = Split(varPara(0), ",")
'
'    strAuthenticID = varPara(2)
'    bytComp = Val(Left(strAuthenticID, 3))
'    lngCustID = Val(Right(strAuthenticID, 7))
'    strOwner = GetOwner(cn)
'
'    varData = Split(strPara, vbCrLf)
'
'
'    strQry = "SELECT CUSTID FROM " & strOwner & "SO002B" & Get_DB_Link(cn) & _
'                    "WHERE MEMBERID='" & varPara(1) & "' AND AUTHENTICID='" & strAuthenticID & "'"
'
'    lngCustID = Val(GetValue(cn, strQry) & "")
'
'    If lngCustID = 0 Then
'        JustDoIt = "-2,訂購PPV認證錯誤"
'        GoTo 66
'    End If
'
'    Set objIPPV = CreateObject("setIPPV.clsNagraSTB")
'    With objIPPV
'            Set .uConn = Get_General_CN
'            garyGi(1) = "WEB"
'            .ugaryGi = Join(garyGi, Chr(9))
'    End With
'
'    ''依傳來參數呼叫CA命令點數下載模組 (STBCreDold)
'
'    '點數線上訂購--下載點數
'    '9、會員代號、認證ID(折行)
'    '設備流水號、STB序號、ICC序號、點數(折行)
'    '設備流水號、STB序號、ICC序號、點數(折行)
'    '…
'    '0 : 成功
'    '-8 : 點數下載有誤
'    '-19 未啟用IPPV訂購權
'
'    ''Loop處理 (第二行參數開始)
'    For intLoop = 1 To UBound(varData)
'        If RPxx(CStr(varData(intLoop))) <> Empty Then
'            varPara = Split(varData(intLoop), ",")
'            '檢核處理中之STB的iPPV訂購權是否已開啟(IPPVright=1 ?).
'            ''If IPPVright <> 1 Then
'                ''  呼叫CA命令非視覺化開啟iPPV訂購權模組 (EnIPPVright)
'            ''End If
'            If Not IPPVok(strOwner, CStr(varPara(0)), blnNoData) Then
'                ''啟動IPPV訂購權
'                    'Public Function EnIPPVright(ByVal strCompCode As String, _
'                    '    ByVal lngCustId As Long, ByVal strSmartCard As String, _
'                    '    ByVal strResvTime As String, ByVal strSTB As String, _
'                    '    ByVal strDoCaption As String, ByVal strFaciSNo As String) As Boolean
'                    '   Format(Now, "YYYY/MM/DD HH:MM")
'                If blnNoData Then
'                    JustDoIt = "-99,設備流水號不存在"
'                    GoTo 66
'                End If
'                If Not objIPPV.EnIPPVright(CStr(bytComp), _
'                                                    lngCustID, _
'                                                    CStr(varPara(2)), _
'                                                    "", _
'                                                    CStr(varPara(1)), _
'                                                    "", _
'                                                    CStr(varPara(1))) Then
'                    strErr = objIPPV.uMsgBack
'                    JustDoIt = "-19,未啟用IPPV訂購權"
'                    GoTo 66
'                End If
'            End If
'
'            ''下載/設定IPPV點數or (無回傳)
'            'Public Function STBCreDold(ByVal strCompCode As String, _
'            '    ByVal lngCustId As Long, ByVal strSmartCard As String, _
'            '    ByVal strResvTime As String, ByVal strSTB As String, _
'            '    ByVal strDoCaption As String, ByVal strCreditMode As String, _
'            '    ByVal strCredit As String, ByVal strCollectDate As String, ByVal blnReturnBk As Boolean) As Boolean
'
'            '    寧可避開誘餌，也勝過在陷阱中掙扎。 說:
'            '    ResvTime=''
'            '    寧可避開誘餌，也勝過在陷阱中掙扎。 說:
'            '    CreditMode='01'
'            '    寧可避開誘餌，也勝過在陷阱中掙扎。 說:
'            '   CollectDate = sysdate
'
'            With objIPPV
'                If Not objIPPV.STBCreDold(CStr(bytComp), _
'                                                            lngCustID, _
'                                                            CStr(varPara(2)), _
'                                                            "", _
'                                                            CStr(varPara(1)), _
'                                                            "", _
'                                                            "03", _
'                                                            CStr(Val(varPara(3)) + Get_IPPV_Credit(strOwner, CStr(varPara(0)))), _
'                                                            Format(Now, "YYYY/MM/DD HH:MM"), _
'                                                            ReturnBack(strOwner, CStr(varPara(0)))) Then
'                    strErr = .uMsgBack
'                    JustDoIt = "-8,點數下載有誤"
'                    GoTo 66
'                Else
'                    MakeXML CStr(varPara(0)), CStr(varPara(1)), CStr(varPara(2)), CStr(varPara(3)), .uMsgBack
''                    If .uMsgBack <> Empty Then
''                        JustDoIt = "0,成功(" & .uMsgBack & ")"
''                    Else
''                        JustDoIt = "0,成功"
''                    End If
'                End If
'            End With
'
'            ''If 點數下載結果不ok Then
'                ''  回傳(-8,點數下載有誤)
'            ''Else
'                ''  回傳(0,成功)
'            ''End If
'
'            ''呼叫CA命令非視覺化點數下載模組(STBCreDold[有回傳機制call]/ STBCreDoldNb[無回傳機制call])
'            ''PS:
'            ''有回傳機制定義 : so041. iPPVCallBack=1 and 該STB ModelCode 對應到 CD043. FuncFlag01=1
'            ''無回傳機制定義 : so041. iPPVCallBack=0 or (so041. iPPVCallBack=1 and 該STB ModelCode對應到CD043. FuncFlag01=0)
'        End If
'    Next
'    ''End Loop
'    If strXMLresult <> Empty Then
'    End If
'    JustDoIt = "0,成功" & vbCrLf & _
'                        "<?xml version=""1.0"" encoding=""BIG5"" standalone=""yes"" ?>" & vbCrLf & _
'                        "<DataSet>" & vbCrLf & _
'                        "  <DataTable>" & vbCrLf & _
'                        strXMLresult & _
'                        "  </DataTable>" & vbCrLf & _
'                        "</DataSet>" & vbCrLf ' 回傳(0, 成功) 及XML檔
'
'    '01=>ADD 增加 Credit
'    '02=>SUBTRACT 減少 Credit
'    '03=>SET CREDIT 直接設定 Credit 為新值
'    '04=>SET BALANCE 清除 Debit, 直接設定 Credit 為新值
'    '05=>SUB OFFSET 同時減少Debit 與Credit
'
'66:
'    On Error Resume Next
'    Rlx rs
'    Rlx strQry
'    Rlx intLoop
'    Rlx varData
'    Rlx varPara
'    Rlx objIPPV
'    Rlx lngCustID
'    Rlx blnNoData
'    Rlx strAuthenticID
'    Rlx strOwner
'    Set cn = Nothing
'    Set objCnPool = Nothing
'  Exit Function
'ChkErr:
'    ErrHandle "mod_Func9", "JustDoIt"
'    JustDoIt = "-99," & strErr
'End Function

