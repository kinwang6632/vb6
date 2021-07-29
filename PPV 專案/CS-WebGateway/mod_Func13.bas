Attribute VB_Name = "mod_Func13"

Option Explicit

Private cn As Object
Private rs As Object

Private strAuthenticID As String
Private strOwner As String
Private intLoop As Integer
Private lngCustID As Long
Private strOrderNo As String
Private strOrderItem As String
Private blnOneOfOK As Boolean
Private strXMLresult As String

'計次付費節目取消(需經過SMS訂購密碼認證,並包括CA系統訂購授權作業及處理結果)
Public Function JustDoIt(ByRef objWebPool As Object, _
                                        ByRef strPara As String) As String
  On Error GoTo ChkErr
  
    
    Dim varPara As Variant
    Dim varDetail As Variant
    
    Dim strBillData As String
    
    Dim strCancelTime As String
    Dim strReturnCode As String
    Dim strReturnName As String
    
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
'******************************************************************************
    '#3018 判斷傳入的個數是否正確 By Kin 2007/09/03
    varPara = Split(strPara, vbCrLf)
    
    If UBound(varPara) < 2 Then
        JustDoIt = "-99,參數資料不正確 ! 請確認 !"
        GoTo 66
    End If
    varDetail = Split(varPara(0), ",")
    If UBound(varDetail) < 9 Then
        JustDoIt = "-99,參數不足!"
        GoTo 66
    End If
    For intLoop = 1 To UBound(varPara)
        If (UBound(Split(varPara(intLoop), ",")) < 1) And (varPara(intLoop) <> Empty) Then
            JustDoIt = "-99,參數不足!"
            GoTo 66
        End If
    Next intLoop
    
'*****************************************************************************
    strAuthenticID = varDetail(2)
    bytComp = Val(Left(strAuthenticID, 3)) '   公司別代碼
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
    
'   Select count(*) as Cnt from so002b where MemberId=[會員代號] and AuthenticId=[認證ID] and PPVOrderPwd=[ PPV後付訂購密碼];
    Set rs = cn.Execute("SELECT COMPCODE,CUSTID FROM " & strOwner & "SO002B" & Get_DB_Link(cn) & _
                                    " WHERE MEMBERID='" & varDetail(1) & "'" & _
                                    " AND AUTHENTICID='" & strAuthenticID & "'" & _
                                    " AND PPVORDERPWD='" & varDetail(3) & "'")
    
    If rs.AbsolutePage <= 0 Then '    If Cnt = 0 Then
        JustDoIt = "-4,訂購PPV認證錯誤" '       回傳(-4,訂購PPV認證錯誤)
        GoTo 66
    End If '    End If
    
    bytComp = Val(rs("COMPCODE") & "")
    lngCustID = Val(rs("CUSTID") & "")
    
    strCancelTime = varDetail(5) & ""
    strReturnCode = varDetail(8) & ""
    strReturnName = varDetail(9) & ""

    cn.BeginTrans

    '   依傳入之參數(ICC序號、產品代碼)
    '   呼叫CA命令[取消 PPV 訂購節目]非視覺化模組. 高階代號P31
    '  13、會員代號、認證ID、PPV後付訂購密碼、2(表來源網站)、取消時間(YYYYMMDD HH24MI)、WEB(取消人員代號)、網站取消(取消人員)、訂單取消原因代碼、訂單取消原因(折行)
    
    blnOneOfOK = False
    For intLoop = 1 To UBound(varPara) 'Loop處理 (第二行開始)
        
        strBillData = varPara(intLoop) & ""
        '   訂單單號、訂單項次(折行)
        
        If strBillData <> Empty Then
            
            varDetail = Split(strBillData, ",")
            
            strOrderNo = varDetail(0) & ""
            strOrderItem = varDetail(1) & ""
            
            '  Select count(*) as Cnt from so105g
            '   where Orderno=[訂單單號] and Orderitem=[訂單項次]
            '   and Accounting=0 and canceltime is null ;
            If GetValue(cn, "SELECT COUNT(*) FROM " & strOwner & "SO105G" & Get_DB_Link(cn) & _
                                            " WHERE ORDERNO='" & strOrderNo & "' AND ORDERITEM=" & strOrderItem & _
                                            " AND ACCOUNTING=0 AND CANCELTIME IS NULL") = 0 Then '   If Cnt = 0 Then
                ' 記錄不成功的訂單單號、訂單項次、STB序號、ICC卡號、節目名稱、"已結算或已取消"
                If Not UnknowStatus("已結算或已取消") Then
                    JustDoIt = "-99,[PPV節目訂購明細檔]查無訂單資料!"
                    GoTo RollBack
                End If
            Else '   Else
                '       依傳入之參數(ICC序號、產品代碼)呼叫CA命令[取消 PPV 訂購節目]非視覺化模組. 高階代號P31
                If CallNagraProcess(strMsg) Then  '       If 命令成功 Then
                    ' Update so105g set CancelEn=[取消人員代號], CancelName=[取消人員], CancelTime=[取消時間],
                    '   ReturnCode=[訂單取消原因代碼], ReturnName=[訂單取消原因]
                    '   Where orderno=[訂單單號] and orderitem=[訂單項次];
                    blnOneOfOK = True
                    cn.Execute "UPDATE " & strOwner & "SO105G" & Get_DB_Link(cn) & _
                                        " SET CANCELEN='WEB',CANCELNAME='網站取消'" & _
                                        ",CANCELTIME=TO_DATE('" & strCancelTime & "','YYYYMMDD HH24MI')" & _
                                        ",RETURNCODE=" & strReturnCode & _
                                        ",RETURNNAME='" & strReturnName & "'" & _
                                        " WHERE ORDERNO='" & strOrderNo & "' AND ORDERITEM=" & strOrderItem
                    
                    '   會員代號、認證ID、PPV後付訂購密碼、2(表來源網站)、取消時間(YYYYMMDD HH24MI)、WEB(取消人員代號)、網站取消(取消人員)、
                    ' 訂單取消原因代碼、訂單取消原因(折行)
                    
                    ' select count(*) as Cnt from so105g where ordreno=[訂單單號] and canceltime is null;
                    If GetValue(cn, "SELECT COUNT(*) FROM " & strOwner & "SO105G" & Get_DB_Link(cn) & _
                                            " WHERE ORDERNO='" & strOrderNo & "' AND CANCELTIME IS NULL") = 0 Then ' If Cnt = 0 Then
                        
                        '  Update so105e set CancelEn=[取消人員代號], CancelName=[取消人員], CancelTime=[取消時間],
                        '   ReturnCode=[訂單取消原因代碼], ReturnName=[訂單取消原因] Where orderno =[訂單單號] ;
                        cn.Execute "UPDATE " & strOwner & "SO105E" & Get_DB_Link(cn) & _
                                            " SET CANCELEN='WEB',CANCELNAME='網站取消'" & _
                                            ",CANCELTIME=TO_DATE('" & strCancelTime & "','YYYYMMDD HH24MI')" & _
                                            ",RETURNCODE=" & strReturnCode & _
                                            ",RETURNNAME='" & strReturnName & "'" & _
                                            " WHERE ORDERNO='" & strOrderNo & "'"
                    
                    End If ' End If
                'Else
                    ' 記錄不成功的訂單單號、訂單項次、STB序號、ICC卡號、節目名稱、"取消失敗"
                    ' UnknowStatus varDetail(0) & "", varDetail(1) & "", "取消失敗"
                End If
            
            End If '  End If
            
        End If
        
    Next '   End Loop
    
'    If strMsg <> Empty Then
'        JustDoIt = "-99, 例外錯誤 !" & strMsg
'        GoTo RollBack '  RollBack;
'    Else
        If Not blnOneOfOK Then 'If 所有欲取消之節目呼叫CA命令(P31)全部失敗 THEN
            JustDoIt = "-12, 計次付費節目取消訂購失敗" & vbCrLf & GetXML '  回傳(-12, 計次付費節目取消訂購失敗)及XML檔
            GoTo RollBack '  RollBack;
        Else 'Else
            cn.CommitTrans
            JustDoIt = "0,成功" & vbCrLf & GetXML '  回傳(0,成功)及XML檔"
            GoTo 66
        End If 'End If
'    End If
    
    'Ps….
    'XML內容如下:
    '訂單單號、訂單項次、STB序號、ICC卡號、節目名稱、狀態
    
RollBack:
    cn.RollbackTrans
    
66:
    On Error Resume Next
    Rlx strCancelTime
    Rlx strReturnCode
    Rlx strReturnName
    Rlx rs
    Rlx cn
    Rlx varPara
    Rlx varDetail
    Rlx strOwner
    Rlx strAuthenticID
    Rlx intLoop
    Rlx lngCustID
    Rlx blnOneOfOK
    Rlx strXMLresult
    Rlx strOrderNo
    Rlx strOrderItem
    Rlx strBillData
    Rlx strMsg
    Set cn = Nothing
    
  Exit Function
ChkErr:
    ErrHandle "mod_Func13", "JustDoIt"
    cn.RollbackTrans
    JustDoIt = "-99," & strErr
End Function

Private Function GetXML() As String
  On Error GoTo ChkErr
    GetXML = "<?xml version=""1.0"" encoding=""BIG5"" standalone=""yes"" ?>" & vbCrLf & _
                    "<DataSet>" & vbCrLf & _
                    "  <DataTable>" & vbCrLf & _
                    strXMLresult & _
                    "  </DataTable>" & vbCrLf & _
                    "</DataSet>" & vbCrLf
  Exit Function
ChkErr:
    ErrHandle "mod_Func13", "JustDoIt"
End Function

Private Function CallNagraProcess(ByRef strMsg As String) As Boolean
  On Error GoTo ChkErr
    Dim objNS As Object
    Dim intErrPos As Integer
    '   strNotes= "ProductId" & "~" & "ProductName" & "~" & "CreditAmt"
    intErrPos = 0
    Set rs = cn.Execute("SELECT FACISNO,SMARTCARDNO,PRODUCTID,PRODUCTNAME" & _
                                    " FROM " & strOwner & "SO105G" & Get_DB_Link(cn) & _
                                    " WHERE ORDERNO='" & strOrderNo & "' AND ORDERITEM=" & strOrderItem)
'    Set rs = cn.Execute("SELECT FACISNO,SMARTCARDNO,PRODUCTID,PRODUCTNAME,CREDITAMT" & _
                                    " FROM " & strOwner & "SO105G" & Get_DB_Link(cn) & _
                                    " WHERE ORDERNO='" & strOrderNo & "' AND ORDERITEM=" & strOrderItem)
    intErrPos = 1
    Set objNS = CreateObject("setIPPV.clsNagraSTB")
    With objNS
        intErrPos = 2
        strMsg = ""
        intErrPos = 3
        .uCompCode = bytComp
        intErrPos = 4
        Set .uConn = cn
        intErrPos = 5
        garyGi(1) = "WEB"
        intErrPos = 6
        .ugaryGi = Join(garyGi, Chr(9))
        intErrPos = 7
        .uErrPath = Environ("Temp")
        intErrPos = 8
        CallNagraProcess = .CancelPPVPro(bytComp, lngCustID, rs!SMARTCARDNO & "", "", rs!FaciSno & "", "取消PPV節目設定", _
                                                                rs!ProductID & "")
'        CallNagraProcess = .CancelPPVPro(bytComp, lngCustID, rs!SmartCardNo & "", "", rs!FaciSno & "", "取消PPV節目設定", _
                                                                rs!ProductID & "~" & rs!ProductName & "~" & rs!CreditAmt)
        intErrPos = 9
        If CallNagraProcess Then
            intErrPos = 10
            MakeXML "取消成功", rs!FaciSno & "", rs!SMARTCARDNO & "", rs!ProductName & ""
        Else
            ' 記錄不成功的訂單單號、訂單項次、STB序號、ICC卡號、節目名稱、"取消失敗"
            intErrPos = 11
            MakeXML "取消失敗", rs!FaciSno & "", rs!SMARTCARDNO & "", rs!ProductName & ""
        End If
        intErrPos = 12
        strMsg = .uMsgBack
        intErrPos = 13
    End With
    Set objNS = Nothing
  Exit Function
ChkErr:
    strMsg = "[STB 設定錯誤] - " & intErrPos & Err.Description
    On Error Resume Next
    Set objNS = Nothing
End Function

Private Function UnknowStatus(strStatus As String) As Boolean
  On Error GoTo ChkErr
    Set rs = cn.Execute("SELECT FACISNO,SMARTCARDNO,PRODUCTNAME FROM " & strOwner & "SO105G" & Get_DB_Link(cn) & _
                                    " WHERE ORDERNO='" & strOrderNo & "' AND ORDERITEM=" & strOrderItem)
    ' 記錄不成功的訂單單號、訂單項次、STB序號、ICC卡號、節目名稱、"已結算或已取消"
    UnknowStatus = rs.RecordCount > 0
    If UnknowStatus Then MakeXML strStatus, rs!FaciSno & "", rs!SMARTCARDNO & "", rs!ProductName & ""
  Exit Function
ChkErr:
    ErrHandle "mod_Func13", "UnknowStatus"
End Function

Private Sub MakeXML(strStatus As String, _
                                    strFaciSno As String, _
                                    strSmartCardNo As String, _
                                    strProductName As String)
  On Error GoTo ChkErr
    strXMLresult = strXMLresult & _
                            "    <DataRow ORDERNO=""" & strOrderNo & """" & _
                                " ORDERITEM=""" & strOrderItem & """" & _
                                " FACISNO=""" & strFaciSno & """" & _
                                " SMARTCARDNO=""" & strSmartCardNo & """" & _
                                " PRODUCTNAME=""" & RPC(strProductName) & """" & _
                                " STATUS=""" & strStatus & """/>" & vbCrLf
  Exit Sub
ChkErr:
    ErrHandle "mod_Func13", "MakeXML"
End Sub


