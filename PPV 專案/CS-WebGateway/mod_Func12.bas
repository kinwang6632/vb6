Attribute VB_Name = "mod_Func12"
Option Explicit

Private cn As Object
Private rs As Object
Private rs105G As Object
Private rs105E As Object

Private strAuthenticID As String
Private strOwner As String
Private intLoop As Integer
Private lngCustID As Long
Private strOrderNo As String
Private strOrderItem As String
Private strName As String
Private strUpdTime As String
Private blnOneOfOK As Boolean
Private strXMLresult As String

'計次付費節目訂購(需經過SMS訂購密碼認證,並包括CA系統訂購授權作業及處理結果)
Public Function JustDoIt(ByRef objWebPool As Object, _
                                        ByRef strPara As String) As String
  On Error GoTo ChkErr
  
    
    Dim varPara As Variant
    Dim varDetail As Variant
    
    Dim strBillData As String
    
    Dim strMsg As String
    
'    Set cn = Get_General_CN
'
'    If cn.State <= 0 Then
'        Set cn = Get_General_CN
'        If cn.State <= 0 Then
'            strErr = "無法連線資料庫!"
'            JustDoIt = "-99,資料庫連線失敗!"
'            GoTo 66
'        End If
'    End If
    
'   12、會員代號、認證ID、PPV後付訂購密碼、2(表來源網站) (折行)

'   XML(訂單單號、訂單項次、STB序號、ICC卡號、節目名稱、狀態)
'   0:  成功
'   -2:  訂購PPV認證錯誤
'   -4 : 未啟用PPV訂購權
'   -11 : 計次付費節目訂購失敗
'***************************************************************************
'#3018 判斷傳入的個數是否正確 By Kin 2007/08/31
    varPara = Split(strPara, vbCrLf)
    
    If UBound(varPara) < 2 Then
        JustDoIt = "-99,參數資料不正確 ! 請確認 !"
        GoTo 66
    End If

    varDetail = Split(varPara(0), ",")
    If UBound(varDetail) < 4 Then
        JustDoIt = "-99,參數不足!"
        GoTo 66
    End If
    For intLoop = 1 To UBound(varPara)
        If (UBound(Split(varPara(intLoop), ",")) < 6) And (varPara(intLoop) <> Empty) Then
            JustDoIt = "-99,參數不足!"
            GoTo 66
        End If
    Next intLoop
'    If UBound(Split(varPara(1), ",")) < 6 Then
'        JustDoIt = "-99,參數不足!"
'        GoTo 66
'    End If
'    If UBound(Split(varPara(2), ",")) < 6 Then
'        JustDoIt = "-99,參數不足!"
'        GoTo 66
'    End If
'***************************************************************************
    strAuthenticID = varDetail(2)
    bytComp = Val(Left(strAuthenticID, 3)) '   公司別代碼
'    lngCustID = Val(Right(strAuthenticID, 7))
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
    
'   Select count(*) as Cnt from so002b where MemberId=[會員代號] and AuthenticId=[認證ID] and PPVOrderPwd=[ PPV後付訂購密碼]
    Set rs = cn.Execute("SELECT COMPCODE,CUSTID,NAME FROM " & strOwner & "SO002B" & Get_DB_Link(cn) & _
                                    " WHERE MEMBERID='" & varDetail(1) & "'" & _
                                    " AND AUTHENTICID='" & strAuthenticID & "'" & _
                                    " AND PPVORDERPWD='" & varDetail(3) & "'")
    
    If rs.AbsolutePage <= 0 Then '    If Cnt = 0 Then
        JustDoIt = "-2,訂購PPV認證錯誤" '        回傳(-2,訂購PPV認證錯誤)
        GoTo 66
    End If '    End If
    
'    API-12規格調整:　若檢核PPVRight<>1時，調整為不回應訊息，改update　so004　ppvright=1及log SOAC0202即可。
'    If Not Chk_DSTB_OK(varPara) Then
'        JustDoIt = "-4, 未啟用PPV訂購權"
'        GoTo 66
'    End If
    
    bytComp = Val(rs("COMPCODE") & "")
    lngCustID = Val(rs("CUSTID") & "")
    strName = rs("NAME") & ""
    
'   依傳入之參數(ICC序號、產品代碼、節目名稱、點數)呼叫CA命令[訂購 PPV 節目]非視覺化模組. 高階代號P30
'   會員代號、認證ID、PPV後付訂購密碼、2(表來源網站) (折行)
'   設備流水號、STB序號、ICC序號、產品代碼、節目名稱、價格(折行)
'   ……

'   先取訂單單號
    If Not GetOrderNo(cn, strOwner, strOrderNo, JustDoIt) Then GoTo 66
'   1． 訂單單號產生: 西元年月+ Sequence Object S_SO105E_OrderNo
'   Ex: '2005030000001

    If Not GetRS(cn, rs105G, "SELECT * FROM " & strOwner & "SO105G" & Get_DB_Link(cn) & _
                                            " WHERE ROWID=''") Then
                                        ' " WHERE COMPCODE=" & bytComp & " AND ORDERNO=''"
        JustDoIt = "-99,[ PPV節目訂購明細檔 ] 開啟失敗 !"
        GoTo 66
    End If

    cn.BeginTrans
    strUpdTime = Format(RightNow(cn), "YYYY/MM/DD HH:MM:SS")
    blnOneOfOK = False
    
    For intLoop = 1 To UBound(varPara) '   Loop處理(第二行參數開始)；處理各DSTB所訂購的節目
        
        strBillData = varPara(intLoop) & ""
        
        If strBillData <> Empty Then
            
            varDetail = Split(strBillData, ",")
            '   設備流水號、STB序號、ICC序號、產品代碼、節目名稱、價格(折行) + 播出時間(折行) 2006/06/20
            '   ……..
            
            '   呼叫CA命令[訂購 PPV 節目]非視覺化模組. 高階代號P30
            '   strNotes= "ProductId" & "~" & "ProductName" & "~" & "CreditAmt"
            If CallNagraProcess(varDetail(2), varDetail(1), varDetail(3) & "~" & varDetail(4) & "~" & varDetail(5), strMsg) Then    '       If 命令成功 Then
                blnOneOfOK = True
                ' 新增訂單單身於SO105G中
                ' (訂單單號,訂單項次,產品代碼,影片代碼,節目名稱,節目名稱(帳單),節目等級,頻道商ID,節目價格,EPG上之點數,手續費,點數,訂購方式代碼, 訂購方式,小計金額,小計點數,設備流水號,物品序號,智慧卡序號,異動人員代號,異動人員,業者服務類別,EventId)
                If Not Ins_2_105G(varDetail(0), varDetail(1), varDetail(2), varDetail(3), varDetail(4), varDetail(5), True) Then
                    JustDoIt = "-99,[PPV節目訂購明細檔] 新增失敗! " & strErr
                    GoTo RollBack
                End If
                MakeXML "成功", varDetail(1), varDetail(2), varDetail(4), varDetail(6)
            Else '       Else
                ' 新增訂單單身於SO105G中
                ' (訂單單號, 訂單項次,產品代碼,影片代碼,節目名稱,節目名稱(帳單),節目等級,頻道商ID,節目價格,EPG上之點數,手續費,點數,訂購方式代碼,訂購方式,小計金額,小計點數,設備流水號,物品序號,智慧卡序號,異動人員代號,異動人員,業者服務類別, EventId,取消人員代號,取消人員,取消時間,訂單取消原因代碼,訂單取消原因 )
                If Not Ins_2_105G(varDetail(0), varDetail(1), varDetail(2), varDetail(3), varDetail(4), varDetail(5), False) Then
                    JustDoIt = "-99,[PPV節目訂購明細檔] 新增失敗! " & strErr
                    GoTo RollBack
                End If
                MakeXML "失敗", varDetail(1), varDetail(2), varDetail(4), varDetail(6)
            End If '       End If
                
        End If
        
    Next '   End Loop
    
    If blnOneOfOK Then ' 成功
        ' 新增一筆訂單單頭於SO105E中(訂單單號, 公司別, 客戶編號, 客戶名稱, PPV訂購方式代碼, PPV訂購方式, 認證編號, 姓名, 受理時間, 受理人員代號, 受理人員名稱, 業者服務類別, 異動人員, 異動時間 )
        If Ins_2_105E Then
            cn.CommitTrans
            ' XML(訂單單號、訂單項次、STB序號、ICC卡號、節目名稱、狀態)
            JustDoIt = "0,成功" & vbCrLf & _
                                "<?xml version=""1.0"" encoding=""BIG5"" standalone=""yes"" ?>" & vbCrLf & _
                                "<DataSet>" & vbCrLf & _
                                "  <DataTable>" & vbCrLf & _
                                strXMLresult & _
                                "  </DataTable>" & vbCrLf & _
                                "</DataSet>" & vbCrLf ' 回傳(0, 成功) 及XML檔
            GoTo 66
        Else
            JustDoIt = "-99,[PPV訂單資料主檔] 新增失敗!" & strErr
            GoTo RollBack
        End If
    Else '   Else If 處理各DSTB所訂購的節目指令全部失敗 Then
        JustDoIt = "-11, 計次付費節目訂購失敗" & strErr ' 回傳(-11, 計次付費節目訂購失敗)
         GoTo RollBack ' Rollback; (即相關訂購資料皆取消不存)
    End If '   End If
    
RollBack:
    cn.RollbackTrans
    
66:
    On Error Resume Next
    Rlx rs
    Rlx cn
    Rlx varPara
    Rlx varDetail
    Rlx strAuthenticID
    Rlx intLoop
    Rlx strOwner
    Rlx strMsg
    Rlx strBillData
    Rlx lngCustID
    Rlx strOrderNo
    Rlx strOrderItem
    Rlx strName
    Rlx strUpdTime
    Rlx blnOneOfOK
    Rlx rs105G
    Rlx rs105E
    Rlx strXMLresult
'    Rlx lngCustID
    Set cn = Nothing
  Exit Function
ChkErr:
    ErrHandle "mod_Func12", "JustDoIt"
    cn.RollbackTrans
    JustDoIt = "-99," & strErr
End Function

Private Sub MakeXML(ByVal strStatus As String, _
                                    ByVal strFaciSno As String, _
                                    ByVal strSmartCardNo As String, _
                                    ByVal strProductName As String, _
                                    ByVal strEventBeginTime As String)
  On Error GoTo ChkErr
    '   XML(訂單單號、訂單項次、STB序號、ICC卡號、節目名稱、狀態(成功／失敗)) + 播出時間
    strXMLresult = strXMLresult & _
                            "    <DataRow ORDERNO=""" & strOrderNo & """" & _
                                " ORDERITEM=""" & strOrderItem & """" & _
                                " FACISNO=""" & strFaciSno & """" & _
                                " SMARTCARDNO=""" & strSmartCardNo & """" & _
                                " PRODUCTNAME=""" & RPC(strProductName) & """" & _
                                " STATUS=""" & strStatus & """" & _
                                " DEVENTBEGINTIME=""" & strEventBeginTime & """/>" & vbCrLf
  Exit Sub
ChkErr:
    ErrHandle "mod_Func12", "MakeXML"
End Sub

Private Function CallNagraProcess(ByVal strSmartCardNo As String, ByVal strFaciSno As String, _
                                                        ByVal strNotes As String, ByRef strMsg As String) As Boolean
  On Error GoTo ChkErr
    Dim objNS As Object
    Dim strQry As String
    Set objNS = CreateObject("setIPPV.clsNagraSTB")
    With objNS
        strMsg = ""
        .uCompCode = bytComp
        Set .uConn = cn
        garyGi(1) = "WEB"
        .ugaryGi = Join(garyGi, Chr(9))
        .uErrPath = Environ("Temp")
        strQry = "SELECT PPVRIGHT FROM " & strOwner & "SO004" & Get_DB_Link(cn) & _
                        " WHERE FACISNO='" & strFaciSno & "'" & _
                        " AND INSTDATE IS NOT NULL AND PRDATE IS NULL"
        If Val(GetValue(cn, strQry) & "") = 0 Then .EnPPVright bytComp, Val(lngCustID), strSmartCardNo, "", strFaciSno, "啟動PPV訂購權", strFaciSno, ""
        CallNagraProcess = .OrdPPVPro(bytComp, lngCustID, strSmartCardNo & "", "", strFaciSno & "", "PPV節目設定", strNotes)
        strMsg = .uMsgBack

'       舊的程式碼
'        If .EnPPVright(bytComp, Val(lngCustID), strSmartCardNo, "", strFaciSno, "啟動PPV訂購權", strFaciSno, "") Then
'            CallNagraProcess = .OrdPPVPro(bytComp, lngCustID, strSmartCardNo & "", "", strFaciSno & "", "PPV節目設定", strNotes)
'        End If
    
    End With
    Set objNS = Nothing
  Exit Function
ChkErr:
    strMsg = "[STB 設定錯誤] - " & Err.Description
    On Error Resume Next
    Set objNS = Nothing
End Function

Private Function Ins_2_105E() As Boolean
  On Error GoTo ChkErr
'   新增一筆訂單單頭於SO105E中
'   (訂單單號, 公司別, 客戶編號, 客戶名稱, PPV訂購方式代碼, PPV訂購方式, 認證編號, 姓名,
'       受理時間, 受理人員代號, 受理人員名稱, 業者服務類別, 異動人員, 異動時間 )
    
    Dim strCustName As String
    strCustName = RPxx(cn.Execute("SELECT CUSTNAME FROM " & strOwner & "SO001" & Get_DB_Link(cn) & _
                                                        " WHERE CUSTID=" & lngCustID & _
                                                        " AND COMPCODE=" & bytComp).GetString(2, 1, "", "", "") & "")
    
    If GetRS(cn, rs105E, "SELECT ORDERNO,COMPCODE,CUSTID,CUSTNAME," & _
                                        "PPVORDERMODENO,PPVORDERMODE,AUTHENTICID,NAME," & _
                                        "ACCEPTTIME,ACCEPTEN,ACCEPTNAME,SERVICETYPE,UPDEN,UPDTIME" & _
                                        " FROM " & strOwner & "SO105E" & Get_DB_Link(cn) & _
                                        " WHERE ROWID=''") Then
                                        ' " WHERE COMPCODE=" & bytComp & " AND ORDERNO=''"
        With rs105E
                .AddNew
                .Fields("ORDERNO") = strOrderNo ' 訂單單號
                .Fields("COMPCODE") = Val(bytComp) ' 公司別
                .Fields("CUSTID") = lngCustID ' 客戶編號
                .Fields("CUSTNAME") = strCustName ' 客戶名稱
                .Fields("PPVORDERMODENO") = 2 ' PPV訂購方式代碼
                ' Liga , 問題集 : 2232 , Web API說明_20060322_Liga.doc
                ' API-12:訂單之PPV訂購方式內容值由「WEB」改為「網站」
                .Fields("PPVORDERMODE") = "網站" ' PPV訂購方式
                .Fields("AUTHENTICID") = strAuthenticID ' 認證編號
                .Fields("NAME") = strName ' 姓名
                .Fields("ACCEPTTIME") = strUpdTime ' 受理時間
                .Fields("ACCEPTEN") = "WEB" ' 受理人員代號
                .Fields("ACCEPTNAME") = "網站" ' 受理人員名稱
                .Fields("SERVICETYPE") = "C" ' 業者服務類別
                .Fields("UPDEN") = "WEB" ' 異動人員
                .Fields("UPDTIME") = strUpdTime ' 異動時間
                .Update
        End With
    End If
    Ins_2_105E = True
    On Error Resume Next
    Rlx strCustName
  Exit Function
ChkErr:
    ErrHandle "mod_Func12", "Ins_2_105E"
End Function

Private Function Ins_2_105G(ByVal strFaciSeqNo As String, _
                                                ByVal strFaciSno As String, _
                                                ByVal strSmartCardNo As String, _
                                                ByVal strProductID As String, _
                                                ByVal strProductName As String, _
                                                ByVal lngAmount As Long, _
                                                blnSucceed As Boolean) As Boolean
  On Error GoTo ChkErr

'   新增訂單單身於SO105G中

'   If 命令成功 Then
'     (訂單單號,訂單項次,產品代碼,節目名稱,節目名稱(帳單),節目等級,頻道商ID,訂購方式代碼, 訂購方式,
'       小計金額,小計點數,設備流水號,物品序號,智慧卡序號,異動人員代號,異動人員,業者服務類別)
'   Else
'     (訂單單號, 訂單項次,產品代碼,節目名稱,節目名稱(帳單),節目等級,頻道商ID,訂購方式代碼,訂購方式,
'       小計金額,小計點數,設備流水號,物品序號,智慧卡序號,異動人員代號,異動人員,業者服務類別,
'       取消人員代號,取消人員,取消時間,訂單取消原因代碼,訂單取消原因 )
'   End If
    
'   6． SO105G資料來源：
'       (1) 由WEB傳參值回填：產品代碼,節目名稱,小計金額,設備流水號,物品序號,智慧卡序號。
'       (2) 用產品代碼(PRODUCTID)至SO155找尋到資料後回填：節目名稱(帳單),節目等級,頻道商ID。
'       (3) 請依其他欄位說明回填：訂購方式代碼, 訂購方式,,異動人員代號,異動人員,業者服務類別,
'               取消人員代碼,取消人員,取消時間,取消原因代碼,取消原因
            
    With rs105G
            .AddNew
            
            .Fields("OrderNo") = strOrderNo ' 訂單單號
            strOrderItem = .RecordCount
            .Fields("OrderItem") = Val(strOrderItem) ' 訂單項次
            .Fields("ProductId") = strProductID ' 產品代碼
            .Fields("ProductName") = strProductName ' 節目名稱
            
            Program strProductID ' 用產品代碼(PRODUCTID)至SO155找尋到資料後回填：節目名稱(帳單),節目等級,頻道商ID。
            
            .Fields("PPVOrderModeNo") = 2 ' 訂購方式代碼
            
            ' Liga , 問題集 : 2232 , Web API說明_20060322_Liga.doc
            ' API-12:訂單之PPV訂購方式內容值由「WEB」改為「網站」
            
            .Fields("PPVOrderMode") = "網站" ' 訂購方式
            
            .Fields("Amount") = lngAmount ' 直接接受Web參數
            
            ' Liga 說不填小計點數嚕
            '.Fields("CreditAmt") = .Fields("CreditAmt") ' 小計點數 ( EPGPrice + HandlingCredit )
            
            .Fields("FaciSeqNo") = strFaciSeqNo ' 設備流水號
            .Fields("FaciSNO") = strFaciSno ' 物品序號
            .Fields("SmartCardNo") = strSmartCardNo ' 智慧卡序號
            .Fields("ModifyEn") = "WEB" ' 異動人員代號
            .Fields("ModifyName") = "網站" ' 異動人員
            .Fields("ServiceType") = "C" ' 業者服務類別
            
            If Not blnSucceed Then
                .Fields("CancelEn") = "WEB" ' 取消人員代號
                .Fields("CancelName") = "網站" ' 取消人員
                .Fields("CancelTime") = strUpdTime ' 取消時間
                 CancelReason ' 取消原因
            End If
            .Update
    End With
    
    Ins_2_105G = True
  Exit Function
ChkErr:
    ErrHandle "mod_Func12", "Ins_2_105G"
End Function

Private Function Program(ByVal strProductID As String) As Boolean
  On Error GoTo ChkErr
    ' 用產品代碼(PRODUCTID)至SO155找尋到資料後回填：節目名稱(帳單),節目等級,頻道商ID。 ( 頻道商ID Liga 說不存 )
    Program = GetRS(cn, rs, "SELECT BILLNAME,DPARENTALRATING FROM " & strOwner & "SO155" & Get_DB_Link(cn) & _
                                            "WHERE PRODUCTID='" & strProductID & "'")
    If Program Then
        With rs105G
            .Fields("BillProductName") = rs!BillName & "" ' 節目名稱(帳單)
            .Fields("ParentalRating") = rs!dParentalRating & "" ' 節目等級
            '.Fields("ContentProviderID") = .Fields("ContentProviderID") ' 頻道商ID Liga 說不存
        End With
    End If
  Exit Function
ChkErr:
    ErrHandle "mod_Func12", "Program"
End Function

Private Function CancelReason() As Boolean ' 取消原因代碼
  On Error GoTo ChkErr
    ' 取消原因代碼、取消原因: 取CD015.REFNO=4的。(若有多筆則取第1筆回填)
    CancelReason = GetRS(cn, rs, "SELECT CODENO,DESCRIPTION FROM " & strOwner & "CD015" & Get_DB_Link(cn) & _
                                                    "WHERE STOPFLAG=0 AND ASSORT=4 AND SERVICETYPE='C' OR SERVICETYPE IS NULL" & _
                                                    " ORDER BY CODENO")
    If CancelReason Then
        With rs105G
            .Fields("ReturnCode") = rs!CodeNo & "" ' 訂單取消原因代碼
            .Fields("ReturnName") = rs!Description & "" ' 訂單取消原因
        End With
    End If
  Exit Function
ChkErr:
    ErrHandle "mod_Func12", "CancelReason"
End Function

'Private Function Chk_DSTB_OK(ByRef varPara As Variant) As Boolean
'  On Error GoTo ChkErr
'    Dim strQry As String
''   Select PPVright as right from so004 where facisno=[ STB序號] and instdate is not null and prdate is null
''   If Right! = 1 Then
''       回傳(-4, 未啟用PPV訂購權)
''   End If
''   設備流水號、STB序號、ICC序號、產品代碼、節目名稱、點數(折行)
'    Chk_DSTB_OK = True
'    For intLoop = 1 To UBound(varPara) '   Loop處理(第二行參數開始)；即檢核各DSTB是否符合訂購資格
'        strQry = "SELECT PPVRIGHT FROM " & strOwner & "SO004" & Get_DB_Link(cn) & _
'                        " WHERE FACISNO='" & Split(varPara(intLoop), ",")(1) & "'" & _
'                        " AND INSTDATE IS NOT NULL AND PRDATE IS NULL"
'        Chk_DSTB_OK = Val(GetValue(cn, strQry) & "") = 1
'        If Not Chk_DSTB_OK Then Exit For
'    Next '   End Loop
'  Exit Function
'ChkErr:
'    ErrHandle "mod_Func12", "Chk_DSTB_OK"
'End Function



' 底下 XXX Block 的 , 為取消的規格
' XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'3． 產品代碼,影片代碼,節目名稱,節目名稱(帳單),節目等級,頻道商ID,節目價格,EPG上之點數, EventId --- 請關聯SO154, SO155, SO155A取得insert。
' XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'4． 訂購方式代碼,訂購方式, 手續費,點數---
'       select codeno, description, charge, credit from so156 where refno=2;
' XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'5． 小計金額=節目價格+手續費、小計點數=EPG上之點數+點數
' XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
    
'7． 其他欄位說明…
'       異動人員代號: 'WEB'
'       異動人員: "網站'
'       業者服務類別: 'C'
'       取消人員代號: 'WEB'
'       取消人員: "網站'
'       取消時間: 系統產生
'       公司別代碼: so002b.CompCode
'       客戶編號: so002b.CustId
'       客戶名稱: so001.CustName
'       PPV訂購方式代碼: 2
'       PPV訂購方式: 'WEB'
'       姓名: so002b.Name
'       受理時間: 由程式產生
'       受理人員代號: 'WEB'
'       受理人員名稱: '網站'
    

