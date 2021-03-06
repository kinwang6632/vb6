VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Gateway"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private objWebPool As Object ' Connection Pool 物件

Private Sub Class_Initialize() ' 初始化
  On Error GoTo ChkErr
'    Dim varAry As Variant
'    Set objGcn = CreateObject("Gcn.conn")
'    Set objCn = CreateObject("ADODB.Connection")
'    strCn = CreateObject("WScript.Shell").RegRead("HKCU\xInfo")
'    varAry = Split(CreateObject("Encryption.Password").Decrypt(strCn), "/")
'    strCn = "Provider=MSDAORA.1;Password=" & varAry(2) & ";User ID=" & varAry(1) & ";Data Source=" & varAry(0) & ";Persist Security Info=True"
    Set objWebPool = CreateObject("WebPool.PoolManager") ' 建立 Connection Pool 物件
    Set mod_Common.objCrypt = CreateObject("Crypt32DLL.Crypt32")
  Exit Sub
ChkErr:
    ErrHandle "", "Class_Initialize"
End Sub

Private Sub Class_Terminate() ' 釋放物件或變數所使用的資源
  On Error Resume Next
    Rlx bytComp ' 公司別 , 宣告於 mod_Common ( 一般模組 )
    Rlx strErr  ' 錯誤訊息字串 , 宣告於 mod_Common ( 一般模組 )
    Rlx strCrLf  ' vbCrLf & vbCrLf , 宣告常式於 mod_Common ( 一般模組 )
    Rlx objWebPool  ' Connection Pool 物件
    Rlx objCrypt  ' Encrypt / Decrypt 元件
'    Rlx strCn
'    Rlx cn
'    Rlx objCn
'    Rlx objGcn
End Sub

'Public Function Executing(ByVal strPara As String) As String
'    Executing = strPara & Chr(255)
'End Function

' strPara 為傳入之參數
Public Function Executing(ByVal strPara As String) As String
  On Error GoTo ChkErr
'    Dim objAPI As Object
'    If InStr(1, strPara, ",") Then
'        Set objAPI = New clsWebAPI
'        Executing = CallByName(objAPI, "Func" & Split(strPara, ",")(0), VbMethod, strPara)

' -------------------------------------------------------------------------------------------------------------------------------------------

'        2007.05.15 調整程式
'        QQ : 3230
'        Doc: Web API說明_20070515_Jacy.Doc
        
'        V2.8新增：
'            1.  新增格式化的query log：
'                (1) 依據意藍傳入QUERY 的參數, 每個參數以TAB 格開,若遇(折行)就取消(折行), 一個QUERY一行,結果列於最後.
'                詳見如下範例: 依序為日期/時間,API-NO,參數1,參數2,參數3,….結果代號,結果名稱
'                EX:
'                (2) 一個月存一個檔,每次新增就 APPEND在檔案後, 檔名: yyyymm.log (ex: 200701.LOG)
'                (3) 該LOG檔存放於API元件下的\QUERYLOG>路徑下
        
        Dim intCmd As Integer
        Dim intPos As Integer
        Dim strTmp As String
        Dim strLogPath As String
        Dim strLogData As String
        
        
        strLogPath = App.Path & "\QueryLog"
        
        If Dir(strLogPath, vbDirectory) = Empty Then MkDir strLogPath
        
        intCmd = Split(strPara, ",")(0)
        
' -------------------------------------------------------------------------------------------------------------------------------------------
        
        Select Case intCmd
                Case 1 ' 有線電視客戶身份SMS帳號密碼認證申請
                        Executing = mod_Func1.JustDoIt(objWebPool, strPara)
                Case 2 ' 訂購PPV認證
                        Executing = mod_Func2.JustDoIt(objWebPool, strPara)
                Case 3
                Case 4 ' 客戶訂購計次付費節目明細查詢(ppv,ippv) [提供個人與管理端]
                        Executing = mod_Func4.JustDoIt(objWebPool, strPara)
                Case 5 ' 客戶帳務資料查詢(日期區間內得歷次帳單及繳款記錄)(個人)
                        Executing = mod_Func5.JustDoIt(objWebPool, strPara)
                Case 6 ' 客戶有線電視收視費轉帳作業/計次付費預付金額轉帳作業(信用卡小額付款機制)--線上付款授權--收查詢
                        Executing = mod_Func6.JustDoIt(objWebPool, strPara)
                Case 7 ' 客戶有線電視收視費轉帳作業/計次付費預付金額轉帳作業(信用卡小額付款機制)--線上付款授權
                        Executing = mod_Func7.JustDoIt(objWebPool, strPara)
                Case 8 ' 點數線上訂購--線上付款授權
                        Executing = mod_Func8.JustDoIt(objWebPool, strPara)
                Case 9 ' 點數線上訂購--下載點數
                        Executing = mod_Func9.JustDoIt(objWebPool, strPara)
                Case 10 ' 客戶變更STB四碼數字Parental PIN Code
                        Executing = mod_Func10.JustDoIt(objWebPool, strPara)
                Case 11 ' 客戶 iPPV訂購PIN Code
                        Executing = mod_Func11.JustDoIt(objWebPool, strPara)
                Case 12 ' 計次付費節目訂購 (需經過SMS訂購密碼認證,並包括CA系統訂購授權作業及處理結果)
                        Executing = mod_Func12.JustDoIt(objWebPool, strPara)
                Case 13 ' 計次付費節目取消 (需經過SMS訂購密碼認證,並包括CA系統訂購授權作業及處理結果)
                        Executing = mod_Func13.JustDoIt(objWebPool, strPara)
                Case 14 ' 計次付費節目取消原因資料
                        Executing = mod_Func14.JustDoIt(objWebPool, strPara)
                Case 15
                Case 16
                Case 17 ' 客戶安裝STB設備資料查詢(已安裝STB型號,編號及位置)
                        Executing = mod_Func17.JustDoIt(objWebPool, strPara)
                Case 18 ' 讀取點數收費資料
                        Executing = mod_Func18.JustDoIt(objWebPool, strPara)
                Case 19 ' 傳回SMS會員資料 For web update
                        Executing = mod_Func19.JustDoIt(objWebPool, strPara)
                Case 20 ' 欠費客戶名單查詢
                        Executing = mod_Func20.JustDoIt(objWebPool, strPara)
                Case 21 ' 經營區尋找
                        Executing = mod_Func21.JustDoIt(objWebPool, strPara)
                Case 22 ' ???
                        Executing = mod_Func22.JustDoIt(objWebPool, strPara)
                Case 23 ' CMCP 相關資訊查詢
                        Executing = mod_Func23.JustDoIt(objWebPool, strPara)
                Case 24 ' ETHOME-CM相關資訊查詢
                        Executing = mod_Func24.JustDoIt(objWebPool, strPara)
                Case 25 'API-25：ESERVICE回裝機單
                        Executing = mod_Func25.JustDoIt(objWebPool, strPara)
                Case 26 'API-26：ESERVICE回停拆移機單
                        Executing = mod_Func26.JustDoIt(objWebPool, strPara)
                Case 27
                        Executing = mod_Func27.JustDoIt(objWebPool, strPara)
                Case 28 'API-28：ESERVICE回維修工單
                        Executing = mod_Func28.JustDoIt(objWebPool, strPara)
        End Select
        
        If Executing = Empty Then Executing = strErr
        
'        一.每個參數間以TAB 格開, 不需列欄位名稱
'
'        日期/時間            API-NO  參數1  參數2 參數3 ....結果(代碼)  結果(名稱)
'
'        2007/01/01 12:00:05  24      A      B     C     -20    查無資料
'        2007/01/01 12:00:09  24      A      B     C     0      成功
'        2007/01/01 12:00:10  22      A      B     C     D      E      F      G      H      0      成功
'        2007/01/01 12:00:11  22      A      B     C     D      E      F      G      H      -10    失拜

'        strLogData = Format(Now, "YYYY/MM/DD HH:MM:SS") & vbTab & _
                                intCmd & vbTab & strPara & vbTab & Executing

        strLogData = Format(Now, "YYYY/MM/DD HH:MM:SS") & vbTab & _
                                intCmd & vbTab

        strLogData = strLogData & Replace(strPara, ",", vbTab, 1)
        strTmp = Split(Executing, ",")(0)
        strLogData = strLogData & vbTab & strTmp & vbTab
        strLogData = strLogData & Mid(Executing, Len(strTmp) + 2)
        strLogData = Replace(strLogData, vbCrLf, vbTab, 1)
        
        intPos = InStr(1, strLogData, "<?xml version=", 1)
        If intPos Then strLogData = Left(strLogData, intPos - 2)
        
        WriteLog strLogPath & "\" & Format(Date, "YYYYMM") & ".Log", strLogData
        
        Executing = Executing & Chr(255)
'    Else
'        Executing = "So nice ! So good !"
'        Select Case True
'                 Case LCase(strPara) = "showcpu"
'                        Executing = Get_CPU_Info
''                 Case Left(LCase(strPara), 4) = "gcn:"
''                        Executing = GetData(Replace(strPara, "gcn:", "", 1, , vbTextCompare)) & "*" & Chr(255) & "*"
''                 Case Else
''                        Executing = FetchingData(strPara) & "*" & Chr(255) & "*"
'        End Select
'    End If
  Exit Function
ChkErr:
    ErrHandle , "Executing"
End Function

Private Sub WriteLog(strFile As String, strData As String)
  On Error GoTo ChkErr
    Dim intFileNo As Integer
    intFileNo = FreeFile
    Open strFile For Append As intFileNo
    Print #intFileNo, strData
    Close intFileNo
  Exit Sub
ChkErr:
    ErrHandle , "WriteLog"
End Sub


'Private Function GetData(ByRef strPara As String) As String
'  On Error GoTo ChkErr
''    If cn Is Nothing Then
''        Set cn = objGcn.OpenCn(False)
''    Else
''        If cn.state <= 0 Then Set cn = objGcn.OpenCn(False)
''    End If
''    If cn.state > 0 Then
''        GetData = cn.Execute(strPara).GetString(2, , ",", vbCrLf, "")
''        On Error Resume Next
''        cn.Close
''        Set cn = Nothing
''    Else
''        GetData = "Can't connecting database !"
''    End If
'  Exit Function
'ChkErr:
'    ErrHandle , "GetData"
'    GetData = strErr
'End Function
'
'Private Function FetchingData(ByRef strPara As String) As String
'  On Error GoTo ChkErr
''    If objCn Is Nothing Then
''        objCn.CursorLocation = 3
''        objCn.Open strCn
''    Else
''        If objCn.state <= 0 Then
''            objCn.CursorLocation = 3
''            objCn.Open strCn
''        Else
''            objCn.Close
''            objCn.CursorLocation = 3
''            objCn.Open strCn
''        End If
''    End If
''    If objCn.state > 0 Then
''        FetchingData = objCn.Execute(strPara).GetString(2, , ",", vbCrLf, "")
''        On Error Resume Next
''        objCn.Close
''        Set objCn = Nothing
''    Else
''        FetchingData = "Can't connecting database !"
''    End If
'  Exit Function
'ChkErr:
'    ErrHandle , "FetchingData"
'    FetchingData = strErr
'End Function
'
'Private Function Get_CPU_Info() As String
'  On Error GoTo ChkErr
'    Dim objLoc As Object
'    Dim objServ As Object
'    Dim objProcSet As Object
'    Dim objProc As Object
'    Dim strCPU0 As String
'    Dim strInfo(35) As String
'
'    Set objLoc = CreateObject("WbemScripting.SWbemLocator")
'    Set objServ = GetObject("winmgmts:")
'    Set objProcSet = objServ.InstancesOf("Win32_Processor")
'    objServ.Security_.ImpersonationLevel = 3
'
'    If objProcSet Is Nothing Then
'        Get_CPU_Info = "Can't detect CPU information !"
'        Exit Function
'    End If
'
'    If objProcSet.Count = 0 Then
'        Get_CPU_Info = "No processor information is available."
'        Exit Function
'    End If
'
'    strCPU0 = "\\" & Environ("COMPUTERNAME") & "\root\cimv2:Win32_Processor.DeviceID=""CPU0"""
'    Set objProc = objProcSet(strCPU0)
'
'    With objProc
'            strInfo(0) = .Name
'            strInfo(1) = "Processor ID: " & .ProcessorID
'            strInfo(2) = "Description: " & .Description
'            strInfo(3) = "Status: " & .Status
'            strInfo(4) = "Manufacturer: " & .Manufacturer
'            strInfo(5) = "Availability: " & Availability(Val(.Availability & ""))
'            strInfo(6) = "Load Percentage: " & .LoadPercentage
'            strInfo(7) = "Current Clock Speed: " & .CurrentClockSpeed & " MHz"
'            strInfo(8) = "Maximum Clock Speed: " & .MaxClockSpeed & " MHz"
'            strInfo(9) = "Level 2 Cache Size: " & .L2CacheSize & " KB"
'            strInfo(10) = "Level 2 Cache Speed: " & .L2CacheSpeed & " KB"
'            strInfo(11) = "Power Management Supported: " & .PowerManagementSupported
'            strInfo(12) = "Current Voltage: " & .CurrentVoltage & " V"
'            strInfo(13) = "Device ID: " & .DeviceID
'            strInfo(14) = "Status: " & GetCPUstatus(.CpuStatus)
'            strInfo(15) = "Data Width: " & .DataWidth
'            strInfo(16) = "Level: " & .Level
'            strInfo(17) = "External Clock: " & .ExtClock & " MHz"
'            strInfo(18) = "System Type: " & .Architecture
'            strInfo(19) = "AddressWidth : " & .AddressWidth
'            strInfo(20) = "Architecture : " & GetArchitecture(.Architecture)
'            strInfo(21) = "CreationClassName  : " & .CreationClassName
'            strInfo(22) = "Family : " & GetFamily(Val(.Family & ""))
'            strInfo(23) = "Level  : " & .Level
'            strInfo(24) = "ProcessorType : " & GetProcessorType(Val(.ProcessorType & ""))
'            strInfo(25) = "Revision : " & .Revision
'            strInfo(26) = "Role : " & .Role
'            strInfo(27) = "SocketDesignation : " & .SocketDesignation
'            strInfo(28) = "StatusInfo : " & GetStatusInfo(Val(.StatusInfo & ""))
'            strInfo(29) = "Stepping : " & .Stepping
'            strInfo(30) = "SystemCreationClassName : " & .SystemCreationClassName
'            strInfo(31) = "SystemName : " & .SystemName
'            strInfo(32) = "UniqueId : " & .UniqueId
'            strInfo(33) = "UpgradeMethod : " & GetUpgradeMethod(Val(.UpgradeMethod & ""))
'            strInfo(34) = "Version : " & .Version
'            strInfo(35) = "VoltageCaps : " & GetVoltageCaps(Val(.VoltageCaps & ""))
'    End With
'    Get_CPU_Info = Trim(vbCrLf & Join(strInfo, vbCrLf))
'    Erase strInfo
'  Exit Function
'ChkErr:
'    Err.Clear
'    Get_CPU_Info = "Now : " & Now
'End Function
'
'Private Function GetVoltageCaps(intCode As Integer) As String
'    Select Case intCode
'             Case 0
'                    GetVoltageCaps = ""
'             Case 1
'                    GetVoltageCaps = "5 V"
'             Case 2
'                    GetVoltageCaps = "3.3 V"
'             Case 4
'                    GetVoltageCaps = "2.9 V"
'    End Select
'End Function
'
'Private Function GetUpgradeMethod(intCode As Integer) As String
'    Select Case intCode
'             Case 0
'                    GetUpgradeMethod = ""
'             Case 1
'                    GetUpgradeMethod = "Other"
'             Case 2
'                    GetUpgradeMethod = "Unknown"
'             Case 3
'                    GetUpgradeMethod = "Daughter Board"
'             Case 4
'                    GetUpgradeMethod = "ZIF Socket"
'             Case 5
'                    GetUpgradeMethod = "Replacement / Piggy Back"
'             Case 6
'                    GetUpgradeMethod = "None"
'             Case 7
'                    GetUpgradeMethod = "LIF Socket"
'             Case 8
'                    GetUpgradeMethod = "Slot 1"
'             Case 9
'                    GetUpgradeMethod = "Slot 2"
'             Case 10
'                    GetUpgradeMethod = "370 Pin Socket"
'             Case 11
'                    GetUpgradeMethod = "Slot A"
'             Case 12
'                    GetUpgradeMethod = "Slot M"
'    End Select
'End Function
'
'Private Function GetStatusInfo(intCode As Integer) As String
'    Select Case intCode
'             Case 0
'                    GetStatusInfo = ""
'             Case 1
'                    GetStatusInfo = "Other"
'             Case 2
'                    GetStatusInfo = "Unknown"
'             Case 3
'                    GetStatusInfo = "Enabled"
'             Case 4
'                    GetStatusInfo = "Disabled"
'             Case 5
'                    GetStatusInfo = "Not Applicable"
'    End Select
'End Function
'
'Private Function GetFamily(intCode As Integer) As String
'    Select Case intCode
'            Case 0
'                    GetFamily = ""
'            Case 1
'                    GetFamily = "Other"
'            Case 2
'                    GetFamily = "Unknown"
'            Case 3
'                    GetFamily = "8086"
'            Case 4
'                    GetFamily = "80286"
'            Case 5
'                    GetFamily = "80386"
'            Case 6
'                    GetFamily = "80486"
'            Case 7
'                    GetFamily = "8087"
'            Case 8
'                    GetFamily = "80287"
'            Case 9
'                    GetFamily = "80387"
'            Case 10
'                    GetFamily = "80487"
'            Case 11
'                    GetFamily = "Pentium�� brand"
'            Case 12
'                    GetFamily = "Pentium�� Pro"
'            Case 13
'                    GetFamily = "Pentium�� II"
'            Case 14
'                    GetFamily = "Pentium�� processor with MMX technology"
'            Case 15
'                    GetFamily = "Celeron��"
'            Case 16
'                    GetFamily = "Pentium�� II Xeon"
'            Case 17
'                    GetFamily = "Pentium�� III"
'            Case 18
'                    GetFamily = "M1 Family"
'            Case 19
'                    GetFamily = "M2 Family"
'            Case 24
'                    GetFamily = "K5 Family"
'            Case 25
'                    GetFamily = "K6 Family"
'            Case 26
'                    GetFamily = "K6 -2"
'            Case 27
'                    GetFamily = "K6 -3"
'            Case 28
'                    GetFamily = "AMD Athlon�� Processor Family"
'            Case 29
'                    GetFamily = "AMD�� Duron�� Processor"
'            Case 30
'                    GetFamily = "AMD2900 Family"
'            Case 31
'                    GetFamily = "K6-2+"
'            Case 32
'                    GetFamily = "Power PC Family"
'            Case 33
'                    GetFamily = "Power PC 601"
'            Case 34
'                    GetFamily = "Power PC 603"
'            Case 35
'                    GetFamily = "Power PC 603+"
'            Case 36
'                    GetFamily = "Power PC 604"
'            Case 37
'                    GetFamily = "Power PC 620"
'            Case 38
'                    GetFamily = "Power PC X704"
'            Case 39
'                    GetFamily = "Power PC 750"
'            Case 48
'                    GetFamily = "Alpha Family"
'            Case 49
'                    GetFamily = "Alpha 21064"
'            Case 50
'                    GetFamily = "Alpha 21066"
'            Case 51
'                    GetFamily = "Alpha 21164"
'            Case 52
'                    GetFamily = "Alpha 21164PC"
'            Case 53
'                    GetFamily = "Alpha 21164a"
'            Case 54
'                    GetFamily = "Alpha 21264"
'            Case 55
'                    GetFamily = "Alpha 21364"
'            Case 64
'                    GetFamily = "MIPS Family"
'            Case 65
'                    GetFamily = "MIPS R4000"
'            Case 66
'                    GetFamily = "MIPS R4200"
'            Case 67
'                    GetFamily = "MIPS R4400"
'            Case 68
'                    GetFamily = "MIPS R4600"
'            Case 69
'                    GetFamily = "MIPS R10000"
'            Case 80
'                    GetFamily = "SPARC Family"
'            Case 81
'                    GetFamily = "SuperSPARC"
'            Case 82
'                    GetFamily = "microSPARC II"
'            Case 83
'                    GetFamily = "microSPARC IIep"
'            Case 84
'                    GetFamily = "UltraSPARC"
'            Case 85
'                    GetFamily = "UltraSPARC II"
'            Case 86
'                    GetFamily = "UltraSPARC III"
'            Case 87
'                    GetFamily = "UltraSPARC III"
'            Case 88
'                    GetFamily = "UltraSPARC IIIi"
'            Case 96
'                    GetFamily = "68040"
'            Case 97
'                    GetFamily = "68xxx Family"
'            Case 98
'                    GetFamily = "68000"
'            Case 99
'                    GetFamily = "68010"
'            Case 100
'                    GetFamily = "68020"
'            Case 101
'                    GetFamily = "68030"
'            Case 112
'                    GetFamily = "Hobbit Family"
'            Case 120
'                    GetFamily = "Crusoe�� TM5000 Family"
'            Case 121
'                    GetFamily = "Crusoe�� TM3000 Family"
'            Case 128
'                    GetFamily = "Weitek"
'            Case 130
'                    GetFamily = "Itanium�� Processor"
'            Case 144
'                    GetFamily = "PA-RISC Family"
'            Case 145
'                    GetFamily = "PA-RISC 8500"
'            Case 146
'                    GetFamily = "PA-RISC 8000"
'            Case 147
'                    GetFamily = "PA-RISC 7300LC"
'            Case 148
'                    GetFamily = "PA-RISC 7200"
'            Case 149
'                    GetFamily = "PA-RISC 7100LC"
'            Case 150
'                    GetFamily = "PA-RISC 7100"
'            Case 160
'                    GetFamily = "V30 Family"
'            Case 176
'                    GetFamily = "Pentium�� III Xeon��"
'            Case 177
'                    GetFamily = "Pentium�� III Processor with Intel�� SpeedStep�� Technology"
'            Case 178
'                    GetFamily = "Pentium�� 4"
'            Case 179
'                    GetFamily = "Intel�� Xeon��"
'            Case 180
'                    GetFamily = "AS400 Family"
'            Case 181
'                    GetFamily = "Intel�� Xeon�� processor MP"
'            Case 182
'                    GetFamily = "AMD AthlonXP�� Family"
'            Case 183
'                    GetFamily = "AMD AthlonMP�� Family"
'            Case 184
'                    GetFamily = "Intel�� Itanium�� 2"
'            Case 185
'                    GetFamily = "AMD Opteron�� Family"
'            Case 190
'                    GetFamily = "K7"
'            Case 200
'                    GetFamily = "IBM390 Family"
'            Case 201
'                    GetFamily = "G4"
'            Case 202
'                    GetFamily = "G5"
'            Case 250
'                    GetFamily = "i860"
'            Case 251
'                    GetFamily = "i960"
'            Case 260
'                    GetFamily = "SH -3"
'            Case 261
'                    GetFamily = "SH -4"
'            Case 280
'                    GetFamily = "ARM"
'            Case 281
'                    GetFamily = "StrongARM"
'            Case 300
'                    GetFamily = "6x86"
'            Case 301
'                    GetFamily = "MediaGX"
'            Case 302
'                    GetFamily = "MII"
'            Case 320
'                    GetFamily = "WinChip"
'            Case 350
'                    GetFamily = "DSP"
'            Case 500
'                    GetFamily = "Video Processor"
'    End Select
'End Function
'
'Private Function GetArchitecture(intCode As Integer) As String
'    Select Case intCode
'             Case 0
'                    GetArchitecture = "x86"
'             Case 1
'                    GetArchitecture = "MIPS"
'             Case 2
'                    GetArchitecture = "Alpha"
'             Case 3
'                    GetArchitecture = "PowerPC"
'             Case 6
'                    GetArchitecture = "Video ia64"
'    End Select
'End Function
'
'Private Function GetProcessorType(intCode As Integer) As String
'    Select Case intCode
'             Case 0
'                    GetProcessorType = ""
'             Case 1
'                    GetProcessorType = "Other"
'             Case 2
'                    GetProcessorType = "Unknown"
'             Case 3
'                    GetProcessorType = "Central Processor"
'             Case 4
'                    GetProcessorType = "Math Processor"
'             Case 5
'                    GetProcessorType = "DSP Processor"
'             Case 6
'                    GetProcessorType = "Video Processor"
'    End Select
'End Function
'
'Private Function GetCPUstatus(intCode As Integer) As String
'    Select Case intCode
'             Case 0
'                    GetCPUstatus = "Unknown"
'             Case 1
'                    GetCPUstatus = "CPU Enabled"
'             Case 2
'                    GetCPUstatus = "CPU Disabled by User via BIOS Setup"
'             Case 3
'                    GetCPUstatus = "CPU Disabled By BIOS (POST Error)"
'             Case 4
'                    GetCPUstatus = "CPU is Idle"
'             Case 5
'                    GetCPUstatus = "Reserved"
'             Case 6
'                    GetCPUstatus = "Reserved"
'             Case 7
'                    GetCPUstatus = "Other"
'    End Select
'End Function
'
'Private Function Availability(intCode As Integer) As String
'    Select Case intCode
'             Case 0
'                    Availability = ""
'             Case 3
'                    Availability = "Running / Full Power"
'             Case 4
'                    Availability = "Warning"
'             Case 5
'                    Availability = "In Test"
'             Case 6
'                    Availability = "Not Applicable"
'             Case 7
'                    Availability = "Power Off"
'             Case 8
'                    Availability = "Off Line"
'             Case 9
'                    Availability = "Off Duty"
'             Case 10
'                    Availability = "Degraded"
'             Case 11
'                    Availability = "Not Installed"
'             Case 12
'                    Availability = "Install Error"
'             Case 13
'                    Availability = "Power Save - Unknown"
'             Case 14
'                    Availability = "Power Save - Low Power Mode"
'             Case 15
'                    Availability = "Power Save - Standby"
'             Case 16
'                    Availability = "Power Cycle"
'             Case 17
'                    Availability = "Power Save - Warning"
'             Case Else
'                    Availability = "Unknown"
'    End Select
'End Function
