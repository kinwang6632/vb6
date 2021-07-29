Attribute VB_Name = "mod_Func21"
Option Explicit

Private cn As Object

Public Function JustDoIt(ByRef objWebPool As Object, _
                                        ByRef strPara As String) As String
  On Error GoTo ChkErr
    Dim varAry As Variant
    Dim varPara As Variant
    Dim varElement As Variant
    Dim strCustID As String
    Dim objCnPool As Object
    Dim strCompCode As String
    
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
    
    ' 21
    ' 經營區尋找
    
    ' 傳入參數[21、(系統台別複選)、縣市、鄉鎮市、村里庄埔路街其他、鄰、段、巷、弄、衖、號、其他]
    
    ' 1.  拆解系統台別複選內容
    ' 2.  將傳入地址各參數有值的變數串起來
    ' 3.  依拆解出之系統台別來串Owner並尋找該Owner之SO001.InstAddress
    
    ' Loop 系統台別(多個)
    '   依LIKE方式為搜尋地址方法(SO001.InstAddress)
    '   且以該地址各客戶之最後裝機日期做為取客戶狀態依據
    '   (傳入參數中[縣市]有值時, 先以完整地址查詢;
    '   若找不到時, 請去掉縣市部分在做查詢一次.
    '   若入參數中[縣市]無值時,則直接做查詢)
    
    '   IF DATA FOUND THEN
    '     RETURN CompName(SO001), CustStatusName(SO002)
    '     EXIT
    '   End If
    ' End Loop
    
    ' If LOOP完而找不到時 Then
    '    回傳 -18與"查無經營區資料! 可能不在東森經營區或目前未配線至該地址"
    ' End If
    
    varPara = Split(strPara, ",")
    '*****************************************************
    '#3018 增加判斷傳入的個數是否正確 By Kin 2007/09/03
    If UBound(varPara) < 2 Then
        JustDoIt = "-99,參數不足!"
        GoTo 66
    End If
    '*****************************************************

    
    strCompCode = varPara(1)
    
    ' 系統台別複選格式說明:
    ' 傳入之系統台別代碼分隔符號為 [Tab]
    ' Ex:
    ' 21,1    2,屏東縣,九如鄉,…..
    
    ' 傳進參數 : 21、(系統台別複選)、縣市、鄉鎮市、村里庄埔路街其他、鄰、段、巷、弄、衖、號、其他
    If InStr(1, strCompCode, vbTab) Then
        varAry = Split(strCompCode, vbTab)
        For Each varElement In varAry
            If varElement <> Empty Then
                If IsNumeric(varElement) Then
                    If QueryGo(CStr(varElement), GetAddress(varPara, 2), JustDoIt) Then Exit For
                    If QueryGo(CStr(varElement), GetAddress(varPara, 3), JustDoIt) Then Exit For
                Else
                    JustDoIt = "-99,[系統台] 參數錯誤!"
                    Exit For
                End If
            End If
        Next
    Else
        If IsNumeric(strCompCode) Then
            If Not QueryGo(strCompCode, GetAddress(varPara, 2), JustDoIt) Then
                QueryGo strCompCode, GetAddress(varPara, 3), JustDoIt
            End If
        Else
            JustDoIt = "-99,[系統台] 參數錯誤!"
        End If
    End If
    
    ' 0:  成功
    ' -18 : 查無經營區資料! 可能不在東森經營區或目前未配線至該地址
    
    ' 結果碼、結果訊息字串(折行)
    ' 系統台名稱、客戶狀態
    
    ' PS:
    ' a.客戶狀態回傳值為最後裝機日期客戶之狀態
    ' b.以HomePass找經營區
    
66:
    On Error Resume Next
    Rlx varAry
    Rlx varPara
    Rlx varElement
    Rlx strCustID
    Rlx objCnPool
    Rlx strCompCode
    Set cn = Nothing
    Set objCnPool = Nothing
  Exit Function
ChkErr:
    ErrHandle "mod_Func21", "JustDoIt"
    JustDoIt = "-99," & strErr
End Function

Private Function GetAddress(varAry As Variant, intStart As Integer) As String
  On Error GoTo ChkErr
    Dim intLoop As Integer
    For intLoop = intStart To UBound(varAry)
        GetAddress = GetAddress & varAry(intLoop)
    Next
  Exit Function
ChkErr:
    ErrHandle "mod_Func21", "GetAddress"
End Function

Private Function QueryGo(strCompCode As String, _
                                            strAddress As String, _
                                            strResult As String) As Boolean
  On Error GoTo ChkErr
    Dim strOwner As String
    Dim strCustID As String
    bytComp = CByte(RPxx(strCompCode))
    strOwner = GetOwner(cn)
    strCustID = Search_Cust(strAddress, strOwner)
    QueryGo = False
    If strCustID <> Empty Then
        strResult = Get_Cust_Status(strCustID, strOwner)
        If strResult = Empty Then
            strResult = "-18,查無經營區資料! 可能不在東森經營區或目前未配線至該地址"
        Else
            strResult = "0,成功" & vbCrLf & Get_Sys_Platform(CInt(strCompCode)) & "," & strResult
            QueryGo = True
        End If
    Else
        strResult = "-18,查無經營區資料! 可能不在東森經營區或目前未配線至該地址"
    End If
    On Error Resume Next
    Rlx strOwner
    Rlx strCustID
  Exit Function
ChkErr:
    ErrHandle "mod_Func21", "QueryGo"
End Function

Private Function Get_Cust_Status(strCustID As String, strOwner As String) As String
  On Error GoTo ChkErr
    Get_Cust_Status = GetValue(cn, "SELECT CUSTSTATUSNAME FROM " & strOwner & "SO002" & Get_DB_Link(cn) & _
                                                        " WHERE CUSTID =" & strCustID & " AND INSTTIME=" & _
                                                        "(SELECT MAX(INSTTIME) FROM " & strOwner & "SO002" & Get_DB_Link(cn) & _
                                                        " WHERE CUSTID=" & strCustID & ")") & Empty
    '   a.客戶狀態回傳值為最後裝機日期客戶之狀態
  
  Exit Function
ChkErr:
    ErrHandle "mod_Func21", "Get_Cust_Status"
End Function

Private Function Search_Cust(strAddress As String, strOwner As String) As String
  On Error GoTo ChkErr
 '   Search_Cust = GetValue(cn, "SELECT CUSTID FROM " & strOwner & _
                                                    "SO001 WHERE INSTADDRESS = '" & strAddress & "'") & Empty
    Search_Cust = GetValue(cn, "SELECT CUSTID FROM " & strOwner & "SO001" & Get_DB_Link(cn) & _
                                                    "WHERE INSTADDRESS LIKE '" & strAddress & "%'") & Empty
    '   b.以HomePass找經營區
  Exit Function
ChkErr:
    ErrHandle "mod_Func21", "Search_Cust"
End Function

' 系統台別說明:
' 觀昇=1、屏南=2、南天=3、新頻道=5、豐盟=6、振道=7、全聯=8、陽明山=9、新台北=10、
' 金頻道=11、大安文山=12、新唐城=13、新店=14、北桃園=16

Private Function Get_Sys_Platform(intSysID As Integer) As String
  On Error GoTo ChkErr
    Get_Sys_Platform = Choose(intSysID, "觀昇", "屏南", "南天", "", "新頻道", "豐盟", "振道", _
                                                    "全聯", "陽明山", "新台北", "金頻道", "大安文山", "新唐城", "新店", "", "北桃園")

  Exit Function
ChkErr:
    ErrHandle "mod_Func21", "Get_Sys_Platform"
End Function

