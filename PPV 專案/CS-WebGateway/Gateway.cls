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

Private objWebPool As Object ' Connection Pool ª«¥ó

Private Sub Class_Initialize() ' ªì©l¤Æ
  On Error GoTo ChkErr
'    Dim varAry As Variant
'    Set objGcn = CreateObject("Gcn.conn")
'    Set objCn = CreateObject("ADODB.Connection")
'    strCn = CreateObject("WScript.Shell").RegRead("HKCU\xInfo")
'    varAry = Split(CreateObject("Encryption.Password").Decrypt(strCn), "/")
'    strCn = "Provider=MSDAORA.1;Password=" & varAry(2) & ";User ID=" & varAry(1) & ";Data Source=" & varAry(0) & ";Persist Security Info=True"
    Set objWebPool = CreateObject("WebPool.PoolManager") ' «Ø¥ß Connection Pool ª«¥ó
    Set mod_Common.objCrypt = CreateObject("Crypt32DLL.Crypt32")
  Exit Sub
ChkErr:
    ErrHandle "", "Class_Initialize"
End Sub

Private Sub Class_Terminate() ' ÄÀ©ñª«¥ó©ÎÅÜ¼Æ©Ò¨Ï¥Îªº¸ê·½
  On Error Resume Next
    Rlx bytComp ' ¤½¥q§O , «Å§i©ó mod_Common ( ¤@¯ë¼Ò²Õ )
    Rlx strErr  ' ¿ù»~°T®§¦r¦ê , «Å§i©ó mod_Common ( ¤@¯ë¼Ò²Õ )
    Rlx strCrLf  ' vbCrLf & vbCrLf , «Å§i±`¦¡©ó mod_Common ( ¤@¯ë¼Ò²Õ )
    Rlx objWebPool  ' Connection Pool ª«¥ó
    Rlx objCrypt  ' Encrypt / Decrypt ¤¸¥ó
'    Rlx strCn
'    Rlx cn
'    Rlx objCn
'    Rlx objGcn
End Sub

'Public Function Executing(ByVal strPara As String) As String
'    Executing = strPara & Chr(255)
'End Function

' strPara ¬°¶Ç¤J¤§°Ñ¼Æ
Public Function Executing(ByVal strPara As String) As String
  On Error GoTo ChkErr
'    Dim objAPI As Object
'    If InStr(1, strPara, ",") Then
'        Set objAPI = New clsWebAPI
'        Executing = CallByName(objAPI, "Func" & Split(strPara, ",")(0), VbMethod, strPara)

' -------------------------------------------------------------------------------------------------------------------------------------------

'        2007.05.15 ½Õ¾ãµ{¦¡
'        QQ : 3230
'        Doc: Web API»¡©ú_20070515_Jacy.Doc
        
'        V2.8·s¼W¡G
'            1.  ·s¼W®æ¦¡¤Æªºquery log¡G
'                (1) ¨Ì¾Ú·NÂÅ¶Ç¤JQUERY ªº°Ñ¼Æ, ¨C­Ó°Ñ¼Æ¥HTAB ®æ¶},­Y¹J(§é¦æ)´N¨ú®ø(§é¦æ), ¤@­ÓQUERY¤@¦æ,µ²ªG¦C©ó³Ì«á.
'                ¸Ô¨£¦p¤U½d¨Ò: ¨Ì§Ç¬°¤é´Á/®É¶¡,API-NO,°Ñ¼Æ1,°Ñ¼Æ2,°Ñ¼Æ3,¡K.µ²ªG¥N¸¹,µ²ªG¦WºÙ
'                EX:
'                (2) ¤@­Ó¤ë¦s¤@­ÓÀÉ,¨C¦¸·s¼W´N APPEND¦bÀÉ®×«á, ÀÉ¦W: yyyymm.log (ex: 200701.LOG)
'                (3) ¸ÓLOGÀÉ¦s©ñ©óAPI¤¸¥ó¤Uªº\QUERYLOG>¸ô®|¤U
        
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
                Case 1 ' ¦³½u¹qµø«È¤á¨­¥÷SMS±b¸¹±K½X»{ÃÒ¥Ó½Ð
                        Executing = mod_Func1.JustDoIt(objWebPool, strPara)
                Case 2 ' ­qÁÊPPV»{ÃÒ
                        Executing = mod_Func2.JustDoIt(objWebPool, strPara)
                Case 3
                Case 4 ' «È¤á­qÁÊ­p¦¸¥I¶O¸`¥Ø©ú²Ó¬d¸ß(ppv,ippv) [´£¨Ñ­Ó¤H»PºÞ²zºÝ]
                        Executing = mod_Func4.JustDoIt(objWebPool, strPara)
                Case 5 ' «È¤á±b°È¸ê®Æ¬d¸ß(¤é´Á°Ï¶¡¤º±o¾ú¦¸±b³æ¤ÎÃº´Ú°O¿ý)(­Ó¤H)
                        Executing = mod_Func5.JustDoIt(objWebPool, strPara)
                Case 6 ' «È¤á¦³½u¹qµø¦¬µø¶OÂà±b§@·~/­p¦¸¥I¶O¹w¥Iª÷ÃBÂà±b§@·~(«H¥Î¥d¤pÃB¥I´Ú¾÷¨î)--½u¤W¥I´Ú±ÂÅv--¦¬¬d¸ß
                        Executing = mod_Func6.JustDoIt(objWebPool, strPara)
                Case 7 ' «È¤á¦³½u¹qµø¦¬µø¶OÂà±b§@·~/­p¦¸¥I¶O¹w¥Iª÷ÃBÂà±b§@·~(«H¥Î¥d¤pÃB¥I´Ú¾÷¨î)--½u¤W¥I´Ú±ÂÅv
                        Executing = mod_Func7.JustDoIt(objWebPool, strPara)
                Case 8 ' ÂI¼Æ½u¤W­qÁÊ--½u¤W¥I´Ú±ÂÅv
                        Executing = mod_Func8.JustDoIt(objWebPool, strPara)
                Case 9 ' ÂI¼Æ½u¤W­qÁÊ--¤U¸üÂI¼Æ
                        Executing = mod_Func9.JustDoIt(objWebPool, strPara)
                Case 10 ' «È¤áÅÜ§óSTB¥|½X¼Æ¦rParental PIN Code
                        Executing = mod_Func10.JustDoIt(objWebPool, strPara)
                Case 11 ' «È¤á iPPV­qÁÊPIN Code
                        Executing = mod_Func11.JustDoIt(objWebPool, strPara)
                Case 12 ' ­p¦¸¥I¶O¸`¥Ø­qÁÊ (»Ý¸g¹LSMS­qÁÊ±K½X»{ÃÒ,¨Ã¥]¬ACA¨t²Î­qÁÊ±ÂÅv§@·~¤Î³B²zµ²ªG)
                        Executing = mod_Func12.JustDoIt(objWebPool, strPara)
                Case 13 ' ­p¦¸¥I¶O¸`¥Ø¨ú®ø (»Ý¸g¹LSMS­qÁÊ±K½X»{ÃÒ,¨Ã¥]¬ACA¨t²Î­qÁÊ±ÂÅv§@·~¤Î³B²zµ²ªG)
                        Executing = mod_Func13.JustDoIt(objWebPool, strPara)
                Case 14 ' ­p¦¸¥I¶O¸`¥Ø¨ú®ø­ì¦]¸ê®Æ
                        Executing = mod_Func14.JustDoIt(objWebPool, strPara)
                Case 15
                Case 16
                Case 17 ' «È¤á¦w¸ËSTB³]³Æ¸ê®Æ¬d¸ß(¤w¦w¸ËSTB«¬¸¹,½s¸¹¤Î¦ì¸m)
                        Executing = mod_Func17.JustDoIt(objWebPool, strPara)
                Case 18 ' Åª¨úÂI¼Æ¦¬¶O¸ê®Æ
                        Executing = mod_Func18.JustDoIt(objWebPool, strPara)
                Case 19 ' ¶Ç¦^SMS·|­û¸ê®Æ For web update
                        Executing = mod_Func19.JustDoIt(objWebPool, strPara)
                Case 20 ' ¤í¶O«È¤á¦W³æ¬d¸ß
                        Executing = mod_Func20.JustDoIt(objWebPool, strPara)
                Case 21 ' ¸gÀç°Ï´M§ä
                        Executing = mod_Func21.JustDoIt(objWebPool, strPara)
                Case 22 ' ???
                        Executing = mod_Func22.JustDoIt(objWebPool, strPara)
                Case 23 ' CMCP ¬ÛÃö¸ê°T¬d¸ß
                        Executing = mod_Func23.JustDoIt(objWebPool, strPara)
                Case 24 ' ETHOME-CM¬ÛÃö¸ê°T¬d¸ß
                        Executing = mod_Func24.JustDoIt(objWebPool, strPara)
                Case 25 'API-25¡GESERVICE¦^¸Ë¾÷³æ
                        Executing = mod_Func25.JustDoIt(objWebPool, strPara)
                Case 26 'API-26¡GESERVICE¦^°±©î²¾¾÷³æ
                        Executing = mod_Func26.JustDoIt(objWebPool, strPara)
                Case 27
                        Executing = mod_Func27.JustDoIt(objWebPool, strPara)
                Case 28 'API-28¡GESERVICE¦^ºû­×¤u³æ
                        Executing = mod_Func28.JustDoIt(objWebPool, strPara)
        End Select
        
        If Executing = Empty Then Executing = strErr
        
'        ¤@.¨C­Ó°Ñ¼Æ¶¡¥HTAB ®æ¶}, ¤£»Ý¦CÄæ¦ì¦WºÙ
'
'        ¤é´Á/®É¶¡            API-NO  °Ñ¼Æ1  °Ñ¼Æ2 °Ñ¼Æ3 ....µ²ªG(¥N½X)  µ²ªG(¦WºÙ)
'
'        2007/01/01 12:00:05  24      A      B     C     -20    ¬dµL¸ê®Æ
'        2007/01/01 12:00:09  24      A      B     C     0      ¦¨¥\
'        2007/01/01 12:00:10  22      A      B     C     D      E      F      G      H      0      ¦¨¥\
'        2007/01/01 12:00:11  22      A      B     C     D      E      F      G      H      -10    ¥¢«ô

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
'                    GetFamily = "Pentium—û brand"
'            Case 12
'                    GetFamily = "Pentium—û Pro"
'            Case 13
'                    GetFamily = "Pentium—û II"
'            Case 14
'                    GetFamily = "Pentium—û processor with MMX technology"
'            Case 15
'                    GetFamily = "Celeron—ü"
'            Case 16
'                    GetFamily = "Pentium—û II Xeon"
'            Case 17
'                    GetFamily = "Pentium—û III"
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
'                    GetFamily = "AMD Athlon—ü Processor Family"
'            Case 29
'                    GetFamily = "AMD—û Duron—ü Processor"
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
'                    GetFamily = "Crusoe—ü TM5000 Family"
'            Case 121
'                    GetFamily = "Crusoe—ü TM3000 Family"
'            Case 128
'                    GetFamily = "Weitek"
'            Case 130
'                    GetFamily = "Itanium—ü Processor"
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
'                    GetFamily = "Pentium—û III Xeon—ü"
'            Case 177
'                    GetFamily = "Pentium—û III Processor with Intel—û SpeedStep—ü Technology"
'            Case 178
'                    GetFamily = "Pentium—û 4"
'            Case 179
'                    GetFamily = "Intel—û Xeon—ü"
'            Case 180
'                    GetFamily = "AS400 Family"
'            Case 181
'                    GetFamily = "Intel—û Xeon—ü processor MP"
'            Case 182
'                    GetFamily = "AMD AthlonXP—ü Family"
'            Case 183
'                    GetFamily = "AMD AthlonMP—ü Family"
'            Case 184
'                    GetFamily = "Intel—û Itanium—û 2"
'            Case 185
'                    GetFamily = "AMD Opteron—ü Family"
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
