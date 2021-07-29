Attribute VB_Name = "mod_XML"
' PR : Hammer
' Start date : 2005/07/30
' Last Modify : 2005/08/02

Option Explicit

Public rsSO158 As New ADODB.Recordset ' For CAPPV~.. XML 使用 , SO158 (DexTmp)
Public rsSO163 As New ADODB.Recordset ' For CASUB~.. XML 使用 , SO163 (SubChannel)

Public rsSO159 As New ADODB.Recordset ' For Nagra~.. XML 使用 , SO159 (EpgTmp)
Public rsSO162 As New ADODB.Recordset ' For Nagra~.. XML 使用 , SO162 (ChannelInfo)
Public rsSO161 As New ADODB.Recordset ' For Nagra~.. XML 使用 , SO161 (Subscription)

Public str_XML_Class As String ' XML 分類 , 依 XML 檔名前 5 碼來區別
Public objDOM As Object ' M$ 文件物件模型 (DOM)

Public strCreationDate As String
Private strLanguage As String ' 158 , 159 共用
Private strChannelId As String
Private blnAdding As Boolean
Private strDirector_and_Actor As String ' 159
Private strActor As String
Public Const strUpdEn = "系統處理"

' SO162 使用變數
Private blnAddingSO162 As Boolean
Private strCH_ID As String
Private strCH_NO As String
Private strLang162 As String
Private StrName As String
Private strShortName As String
Private strProviderName As String
Private strDescription As String
Private strEngName As String
Private strEngShortName As String
Private strEngDescription As String
Private strServiceId As String
Private strTransportId As String
Private strContentNibble1 As String
Private strContentNibble2 As String
Private strUserNibble1 As String
Private strUserNibble2 As String
'Private strKind As String

' SO161 使用變數
Private blnAddingSO161 As Boolean
Private strLang161 As String
Private strChannel_Id As String
Private strEvent_BeginTime As String
Private lngDuration As Long
Private strEvent_Id As String
Private strEvent_Type As String
Private strFilm_Name As String
Private strEpg_Desc As String
Private strEng_Name As String
Private strEng_Epg_Desc As String
Private strParental_Rating As String
Private strParental_Rating_Name As String
Private strContent_Nibble1 As String
Private strContent_Nibble2 As String
Private strUser_Nibble1 As String
Private strUser_Nibble2 As String
Private strDirector_Actor As String ' 161

'3.  節目表XML檔之檔案命名原則
'(1) 計次付費(PPV)： CAPPV西元年月日時分秒. XML    例：CAPPV20050520083000.XML
'(2) 付費頻道(SUB)： CASUB西元年月日時分秒.XML    例：CASUB20050520173015.XML
'(3) EPG Wrapper ： Nagra_西元年月日時分秒. XML   例：Nagra_20050530031019.XML

Private Function Open158() As Boolean ' 開啟 SO158 ( DexTmp ) 資料錄集 CAPPV~.xml
  On Error GoTo ChkErr
    With rsSO158
         If .State > 0 Then .Close
        .CursorLocation = adUseClient
        .Open "SELECT * FROM SO158 WHERE 0=1", cnCOMM, adOpenStatic, adLockBatchOptimistic
    End With
    blnAdding = False
  Exit Function
ChkErr:
    ErrHandle "mod_XML", "Open158"
End Function

Private Function Open163() As Boolean ' 開啟 SO163 ( SubChannel ) 資料錄集 CASUB~.xml
  On Error GoTo ChkErr
    With rsSO163
         If .State > 0 Then .Close
        .CursorLocation = adUseClient
        .Open "SELECT * FROM SO163 WHERE 0=1", cnCOMM, adOpenStatic, adLockBatchOptimistic
    End With
    blnAdding = False
  Exit Function
ChkErr:
    ErrHandle "mod_XML", "Open163"
End Function

Private Function Open159() As Boolean ' 開啟 SO159 ( EpgTmp ) 資料錄集 NAGRA~.xml
  On Error GoTo ChkErr
    With rsSO159
         If .State > 0 Then .Close
        .CursorLocation = adUseClient
        .Open "SELECT * FROM SO159 WHERE 0=1", cnCOMM, adOpenStatic, adLockBatchOptimistic
    End With
    blnAdding = False
  Exit Function
ChkErr:
    ErrHandle "mod_XML", "Open159"
End Function

Private Function Open162(Optional blnFlag As Boolean = False) As Boolean ' 開啟 SO162 ( ChannelInfo ) 資料錄集 NAGRA~.xml
  On Error GoTo ChkErr
    With rsSO162
         If .State > 0 Then .Close
        .CursorLocation = adUseClient
        If blnFlag Then
            .Open "SELECT * FROM SO162 WHERE CHANNELID='" & strCH_ID & "'", cnCOMM, adOpenStatic, adLockBatchOptimistic
        Else
            .Open "SELECT * FROM SO162 WHERE 0=1", cnCOMM, adOpenStatic, adLockBatchOptimistic
        End If
    End With
'    blnAddingSO162 = False
  Exit Function
ChkErr:
    ErrHandle "mod_XML", "Open162"
End Function

Private Function Open161(Optional blnFlag As Boolean = False) As Boolean ' 開啟 SO161 ( Subscription ) 資料錄集 NAGRA~.xml
  On Error GoTo ChkErr
    With rsSO161
         If .State > 0 Then .Close
        .CursorLocation = adUseClient
        If blnFlag Then
            .Open "SELECT * FROM SO161 WHERE EVENTID='" & strEvent_Id & "'", cnCOMM, adOpenStatic, adLockBatchOptimistic
        Else
            .Open "SELECT * FROM SO161 WHERE 0=1", cnCOMM, adOpenStatic, adLockBatchOptimistic
        End If
    End With
'    blnAddingSO161= False
  Exit Function
ChkErr:
    ErrHandle "mod_XML", "Open161"
End Function

' 透過 DOM 載入 XML 檔案
Public Function LoadXML(strFile As String, _
                                        Optional blnValidate As Boolean = False, _
                                        Optional ByRef ctlList As ListBox) As Boolean
  On Error GoTo ChkErr
    Dim str2crlf As String
    str2crlf = vbCrLf & vbCrLf
    With objDOM
            .async = True ' 非同步執行
            .validateOnParse = blnValidate ' 關掉 Vadidation Parser
            LoadXML = .Load(strFile)
            If Not LoadXML Then ' 當 XML 檔案發生錯誤 , 則顯示訊息
                MsgBox "錯誤碼 : " & .parseError.errorCode & str2crlf & _
                                "錯誤原因 : " & .parseError.reason & vbCrLf & _
                                "錯誤行 : " & .parseError.Line & _
                                "  ( 位置 : " & .parseError.linepos & " )" & str2crlf & _
                                "錯誤來源 : " & .parseError.srcText & str2crlf & _
                                "錯誤檔案 : " & .parseError.url & _
                                "  ( 位置 : " & .parseError.filepos & " )", vbInformation, "訊息"
            End If
    End With
    str2crlf = Empty
    ctlList.Clear
    ' XML 檔案非同步開啟後 , 開啟 XML 檔案對應之資料錄集 , 以供後續資料處理
    Select Case str_XML_Class
                Case "CAPPV"
                            Open158
                Case "CASUB"
                            Open163
                Case "NAGRA"
                            Open159
                            Open162
                            Open161
    End Select
  Exit Function
ChkErr:
    ErrHandle "mod_XML", "LoadXML"
End Function

Public Function Update_XML_Table() As Boolean ' 將 XML 資料存檔至對應的 Table 裡
  On Error GoTo ChkErr
    Update_XML_Table = False
    Select Case str_XML_Class
                Case "CAPPV"
                        Update_Nagra_SO158
                        rsSO158.UpdateBatch
                Case "CASUB"
                        Update_Nagra_SO163
                Case "NAGRA"
                        Update_Nagra_SO159
                        Update_Nagra_SO162
                        Update_Nagra_SO161
    End Select
    Update_XML_Table = True
  Exit Function
ChkErr:
    ErrHandle "mod_XML", "Update_XML_Table"
End Function

Private Sub Update_Nagra_SO163()
  On Error GoTo ChkErr
    With rsSO163
        If Val(cnCOMM.Execute("SELECT COUNT(*) FROM SO163" & _
                                    " WHERE PRODUCTID='" & .Fields("ProductId") & "'" & _
                                    " AND CHANNELID='" & .Fields("ChannelId") & "'").GetString & "") = 0 Then
            .UpdateBatch adAffectCurrent
        Else
            .CancelBatch adAffectCurrent
        End If
    End With
    blnAdding = False
  Exit Sub
ChkErr:
    ErrHandle "mod_XML", "Update_Nagra_SO163"
End Sub

Private Sub Update_Nagra_SO158()
  On Error GoTo ChkErr
    With rsSO158
        .Fields("dFileName") = rsVirtual(1) & ""
        .Fields("dCreationDate") = strCreationDate
        .Fields("UpdTime") = Format(Now, "ee/mm/dd hh:mm:ss")
        .Fields("UpdEn") = strUpdEn
    End With
    blnAdding = False
  Exit Sub
ChkErr:
    ErrHandle "mod_XML", "Update_Nagra_SO158"
End Sub

Private Sub Update_Nagra_SO159()
  On Error GoTo ChkErr
    With rsSO159
        If CP_Str(.Fields("EventType"), "P") Then
            'Debug.Print .Fields("EventType")
            .Fields("eFileName") = rsVirtual(1) & ""
            .Fields("eCreationDate") = strCreationDate
            .Fields("ChannelId") = strChannelId
            .Fields("UpdTime") = Format(Now, "ee/mm/dd hh:mm:ss")
            .Fields("UpdEn") = strUpdEn
            If CP_Str(strLanguage, "eng") Then
                .Fields("EngEpgDesc") = strDirector_and_Actor & vbCrLf & .Fields("EngEpgDesc")
            ElseIf CP_Str(strLanguage, "chi") Or CP_Str(strLanguage, "cht") Then
                .Fields("EpgDesc") = strDirector_and_Actor & vbCrLf & .Fields("EpgDesc")
            End If
            .UpdateBatch adAffectCurrent
        Else
            .CancelBatch adAffectCurrent
        End If
    End With
    strDirector_and_Actor = ""
    blnAdding = False
  Exit Sub
ChkErr:
    ErrHandle "mod_XML", "Update_Nagra_SO159"
End Sub

'B. 需依據Channel資訊(SO162)的PK(ChannelID)判斷是否有重覆
'    若比對不存在 , 則insert一筆資料至SO162
'    若比對已存在 , 則將該筆資料完整update至SO162
Private Sub Update_Nagra_SO162() ' 更新 SO162
  On Error GoTo ChkErr
    Dim blnAddNew As Boolean
    With rsSO162
        If strCH_NO <> Empty Then
            blnAddNew = Val(cnCOMM.Execute("SELECT COUNT(*) FROM SO162 WHERE CHANNELID='" & strCH_ID & "'").GetString & "") = 0
            If blnAddNew Then
                Open162
                .AddNew
            Else
                Open162 True
            End If
            .Fields("ChannelID") = strCH_ID
            .Fields("ChannelNo") = strCH_NO
            .Fields("Language") = strLang162
            .Fields("Name") = StrName
            .Fields("ShortName") = strShortName
            .Fields("ProviderName") = strProviderName
            .Fields("Description") = strDescription
            .Fields("EngName") = strEngName
            .Fields("EngShortName") = Left(strEngShortName, 20)
            .Fields("EngDescription") = strEngDescription
            .Fields("ServiceId") = strServiceId
            .Fields("TransportId") = strTransportId
            .Fields("ContentNibble1") = strContentNibble1
            .Fields("ContentNibble2") = strContentNibble2
            .Fields("UserNibble1") = strUserNibble1
            .Fields("UserNibble2") = strUserNibble2
    '        .Fields("Kind") = strKind
            .UpdateBatch adAffectCurrent
        End If
    End With
    Clear_SO162_Variant
    blnAddingSO162 = False
  Exit Sub
ChkErr:
    ErrHandle "mod_XML", "Update_Nagra_SO162"
End Sub

Private Sub Clear_SO162_Variant() ' 清空 SO162 所使用之變數
  On Error Resume Next
    strCH_ID = ""
    strCH_NO = ""
    strLang162 = ""
    StrName = ""
    strShortName = ""
    strProviderName = ""
    strDescription = ""
    strEngName = ""
    strEngShortName = ""
    strEngDescription = ""
    strServiceId = ""
    strTransportId = ""
    strContentNibble1 = ""
    strContentNibble2 = ""
    strUserNibble1 = ""
    strUserNibble2 = ""
'    strKind = ""
End Sub

'C. 取e.<EventType>≠ ’P’的才insert Sub時段資料檔(SO161)，
'    需判斷SO161的PK(EevntID)是否有重覆
'    若比對不存在 , 則insert一筆資料至SO161
'    若比對已存在 , 則將該筆資料完整update至SO161
Private Sub Update_Nagra_SO161() ' 更新 SO161
  On Error GoTo ChkErr
    Dim blnAddNew As Boolean
    With rsSO161
        If strEvent_Id <> Empty And (Not CP_Str(strEvent_Type, "P")) Then
            blnAddNew = Val(cnCOMM.Execute("SELECT COUNT(*) FROM SO161 WHERE EVENTID='" & strEvent_Id & "'").GetString & "") = 0
            If blnAddNew Then
                Open161
                .AddNew
            Else
                Open161 True
            End If
            .Fields("eFileName") = rsVirtual(1) & ""
            .Fields("eCreationDate") = strCreationDate
            .Fields("ChannelID") = strChannel_Id
            .Fields("EventBeginTime") = strEvent_BeginTime
            .Fields("Duration") = lngDuration
            .Fields("EventId") = strEvent_Id
            .Fields("EventType") = strEvent_Type
            .Fields("Name") = strFilm_Name
            .Fields("EngName") = strEng_Name
            If CP_Str(strLang161, "eng") Then
                .Fields("EngEpgDesc") = strDirector_Actor & vbCrLf & strEng_Epg_Desc
                .Fields("EpgDesc") = ""
            ElseIf CP_Str(strLang161, "chi") Or CP_Str(strLang161, "cht") Then
                .Fields("EpgDesc") = strDirector_Actor & vbCrLf & strEng_Epg_Desc
                .Fields("EngEpgDesc") = ""
            End If
            .Fields("ParentalRating") = strParental_Rating
            .Fields("ParentalRatingName") = strParental_Rating_Name
            .Fields("ContentNibble1") = strContent_Nibble1
            .Fields("ContentNibble2") = strContent_Nibble2
            .Fields("UserNibble1") = strUser_Nibble1
            .Fields("UserNibble2") = strUser_Nibble2
            .Fields("UpdTime") = Format(Now, "ee/mm/dd hh:mm:ss")
            .Fields("UpdEn") = strUpdEn
            .UpdateBatch adAffectCurrent
        End If
    End With
    Clear_SO161_Variant
    strDirector_Actor = ""
    blnAddingSO161 = False
  Exit Sub
ChkErr:
    ErrHandle "mod_XML", "Update_Nagra_SO162"
End Sub

Sub Clear_SO161_Variant() ' 清空 SO161 所使用之變數
  On Error Resume Next
'    strChannel_Id = ""
    strEvent_BeginTime = ""
    strEvent_Id = ""
    strEvent_Type = ""
    strFilm_Name = ""
    strEpg_Desc = ""
    strEng_Name = ""
    strEng_Epg_Desc = ""
    strParental_Rating = ""
    strParental_Rating_Name = ""
    strContent_Nibble1 = ""
    strContent_Nibble2 = ""
    strUser_Nibble1 = ""
    strUser_Nibble2 = ""
    strDirector_Actor = ""
End Sub

'6. XML檔拆解至SO158、SO159之其他注意事項
'   (1) 接收下來的XML檔案，其內容有些會為中文字(encoding="UTF-8" 格式)，故需注意於新增至資料庫時，需做格式上的轉換後再存。
'   (2) 當有提供多語言時，則需增加判斷如下：
'       A.  當Text.Language='chi' 或 'cht' 時，則insert至中文欄位。
'       B.  當Text.Language='eng'時，則insert至英文欄位。
'   (3) XML檔裡所有 Tag 裏面，只要有關時間日期的資料一律是 UTC 時間(格林威治標準時間)，
'       必須轉成 Taiwan Local 時間 ( 即原UTC + 8 小時 )後回填資料庫。
Public Function PrcNode(ByRef objElement As Object, ByRef ctlList As ListBox) ' 處理 XML Node 節點
  On Error GoTo ChkErr
    Dim lngLoop As Long
    Dim objNodeLst As Object
'   (4) 請依照【P7 Schema】SO158、SO163、SO159、SO162、SO161資料表處理規則。
    With objElement
            Select Case .NodeType ' Node Tyype
                                Case 1 ' Node elemnts type
                                        GetAttributes objElement, ctlList
                                Case 3 ' Node text type
                                        Select Case str_XML_Class
'                                                    (1) 將檔名為"CAPPV"開頭的XML檔，取其部分所需的資料insert至DEX XML資料檔(SO158)
'                                                       A. CAPPV所給的Dex XML檔，即使有重覆的資料，則一樣完全拆解至SO158。
                                                    Case "CAPPV"
                                                                Select Case .ParentNode.NodeName
                                                                            Case "ProductId"
                                                                                    If blnAdding Then Update_Nagra_SO158
                                                                                    blnAdding = True
                                                                                    rsSO158.AddNew
                                                                                    ctlList.Clear
                                                                                    rsSO158("ProductId") = .NodeValue
                                                                                    
                                                                            Case "ProductName"
                                                                                    If CP_Str(strLanguage, "eng") Then
                                                                                        rsSO158("ProductName") = Null
                                                                                        rsSO158("EngProductName") = .NodeValue
                                                                                    ElseIf CP_Str(strLanguage, "chi") Or CP_Str(strLanguage, "cht") Then
                                                                                        rsSO158("ProductName") = .NodeValue
                                                                                        rsSO158("EngProductName") = Null
                                                                                    End If
                                                                                    
                                                                            Case "ProductDescription"
                                                                                    If CP_Str(strLanguage, "eng") Then
                                                                                        rsSO158("ProductDescription") = Null
                                                                                        rsSO158("EngProductDescription") = .NodeValue
                                                                                    ElseIf CP_Str(strLanguage, "chi") Or CP_Str(strLanguage, "cht") Then
                                                                                        rsSO158("ProductDescription") = .NodeValue
                                                                                        rsSO158("EngProductDescription") = Null
                                                                                    End If
                                                                                    
                                                                            Case "ProductCategory"
                                                                                    rsSO158("ProductCategory") = .NodeValue
                                                                                    
                                                                            Case "EpgPrice"
                                                                                    rsSO158("EpgPrice") = .NodeValue
                                                                                    
                                                                            Case "EventId"
                                                                                    rsSO158("EventId") = .NodeValue
                                                                                    
                                                                            Case "PreviewTime"
                                                                                    rsSO158("EventPreviewTime") = .NodeValue
                                                                                    
                                                                            Case "ParentalRating"
                                                                                    rsSO158("dParentalRating") = .NodeValue
                                                                End Select
                                                                
'                                                    (2) 將檔名為"CASUB"開頭的XML檔，取其部分所需的資料insert至Sub Channel對應檔(SO163)
'                                                       A. CASUB所給的Dex XML檔，需依據SO163資料表的PK(ProductID、ChannelID)判斷是否有重覆。
'                                                       B. 若比對不存在，則insert一筆資料至SO163。
'                                                       C. 若比對已存在，則將該筆資料完整update至SO163。
                                                    Case "CASUB"
                                                                Select Case .ParentNode.NodeName
                                                                            Case "ProductId"
                                                                                    blnAdding = True
                                                                                    rsSO163.AddNew
                                                                                    ctlList.Clear
                                                                                    rsSO163("ProductId") = .NodeValue
                                                                                    
                                                                            Case "ChannelId"
                                                                                    If blnAdding Then
                                                                                        rsSO163.Fields("ChannelId") = .NodeValue
                                                                                        Update_Nagra_SO163
                                                                                    End If
                                                                End Select
                                                                
'                                                    (3) 將檔名為"Nagra_"開頭的XML檔，取其部分所需的資料insert至SO159、SO162、SO161。
'                                                       A. 取e.<EventType>= 'P'的才insert EPG XML資料檔(SO159)，
'                                                           不需考慮是否有無重覆資料，一樣完全拆解至SO159。
'                                                       B. 需依據Channel資訊(SO162)的PK(ChannelID)判斷是否有重覆
'                                                           若比對不存在, 則insert一筆資料至SO162
'                                                           若比對已存在, 則將該筆資料完整update至SO162
'                                                       C. 取e.<EventType>≠ 'P'的才insert Sub時段資料檔(SO161)，
'                                                           需判斷SO161的PK(EevntID)是否有重覆
'                                                           若比對不存在, 則insert一筆資料至SO161
'                                                           若比對已存在, 則將該筆資料完整update至SO161
                                                    Case "NAGRA"
                                                                Select Case .ParentNode.NodeName
                                                                            Case "ChannelId" ' 159 & 162 & 161
                                                                                    strChannelId = .NodeValue ' 159
                                                                                    If blnAddingSO162 Then ' 162
                                                                                        If strCH_NO = Empty Then
                                                                                            Clear_SO162_Variant
                                                                                        Else
                                                                                            Update_Nagra_SO162
                                                                                        End If
                                                                                    End If
                                                                                    blnAddingSO162 = True
                                                                                    ctlList.Clear
                                                                                    strCH_ID = .NodeValue ' 162
                                                                                    strChannel_Id = .NodeValue ' 161
                                                                                    
                                                                            Case "EventId" ' 159 & 161
                                                                                    rsSO159("EventId") = .NodeValue
                                                                                    strEvent_Id = .NodeValue ' 161
                                                                                    If blnAddingSO161 Then ' 161
                                                                                        Update_Nagra_SO161
                                                                                        Clear_SO161_Variant
                                                                                    End If
                                                                                    blnAddingSO161 = True
                                                                                    ctlList.Clear
                                                                                    
                                                                            Case "EventType" ' 159 & 161
                                                                                    rsSO159("EventType") = .NodeValue
                                                                                    strEvent_Type = .NodeValue ' 161
                                                                                    
                                                                            Case "Name" ' 159 & 161
                                                                                    If CP_Str(strLanguage, "eng") Then
                                                                                        rsSO159("Name") = Null
                                                                                        rsSO159("EngName") = .NodeValue
                                                                                        strFilm_Name = "" ' 161
                                                                                        strEng_Name = .NodeValue ' 161
                                                                                    ElseIf CP_Str(strLanguage, "chi") Or CP_Str(strLanguage, "cht") Then
                                                                                        rsSO159("Name") = .NodeValue
                                                                                        rsSO159("EngName") = Null
                                                                                        strFilm_Name = .NodeValue ' 161
                                                                                        strEng_Name = "" ' 161
                                                                                    End If
'                                                                                    If CP_Str(strLang161, "eng") Then
'                                                                                    ElseIf CP_Str(strLang161, "chi") Or CP_Str(strLang161, "cht") Then
'                                                                                    End If
                                                                                    
                                                                            Case "Description" ' 159
                                                                                    If CP_Str(strLanguage, "eng") Then
                                                                                        rsSO159("EpgDesc") = Null
                                                                                        rsSO159("EngEpgDesc") = .NodeValue
                                                                                        strEpg_Desc = ""
                                                                                        strEng_Epg_Desc = .NodeValue
                                                                                    ElseIf CP_Str(strLanguage, "chi") Or CP_Str(strLanguage, "cht") Then
                                                                                        rsSO159("EpgDesc") = .NodeValue
                                                                                        rsSO159("EngEpgDesc") = Null
                                                                                        strEpg_Desc = .NodeValue
                                                                                        strEng_Epg_Desc = ""
                                                                                    End If

'                                                                           (4) EPG Wrapper XML裡的<ExtendedInfoName>、<ExtendedInfo>、<Description>的內容：
'                                                                               A.拆解後內容 , 一起存放於SO159.EpgDesc欄位
'                                                                               B.  需要能斷行儲存，以利節目表呈現時可一併注重到美觀。

'                                                                           C.舉例:
'                                                                               <EpgText language="chi">
'                                                                                   <Name>變人(普) </Name>
'                                                                                   <Description>機器人管家安德魯【羅賓威廉斯 飾】，原本遭受眾人排擠，而在朝夕相處之下，大家也漸漸地接納他，當他愛上人類亞曼達後決定與她結婚，他必須放棄永不毀損的機器身體而變成人類，他是否能如願以償地與心愛的人廝守一生呢？</Description>
'                                                                                   <ExtendedInfo name="導演:">克里斯哥倫布 </ExtendedInfo>
'                                                                                   <ExtendedInfo name="演員:">羅賓威廉斯、燕伯斯大衛茲 </ExtendedInfo>
'                                                                               </EpgText>

'                                                                           拆解出來的SO159.EpgDesc欄位值為：4
'                                                                               導演:克里斯哥倫布
'                                                                               演員:羅賓威廉斯、燕伯斯大衛茲
'                                                                               機器人管家安德魯【羅賓威廉斯 飾】，原本遭受眾人排擠，而在朝夕相處之下，大家也漸漸地接納他，當他愛上人類亞曼達後決定與她結婚，他必須放棄永不毀損的機器身體而變成人類，他是否能如願以償地與心愛的人廝守一生呢？

                                                                            Case "ExtendedInfo" ' 159
                                                                                    strDirector_and_Actor = strDirector_and_Actor & .NodeValue
                                                                                    strDirector_Actor = strDirector_Actor & .NodeValue
                                                                                    
                                                                            Case "ParentalRating" ' 159 & 161
                                                                                    rsSO159("ParentalRating") = .NodeValue

'                                                                                   Parentalrating於節目表xml檔裡出現的範圍值為0~15之間，
'                                                                                   需對照CD029節目等級代碼檔的年齡分類，取其符合的節目等級代碼、
'                                                                                   名稱回填SO154.ParentalRating、ParentalRatingName。
'                                                                                   Select codeno,description from cd029
'                                                                                   where SO160.Parentalrating between AGE1 and AGE2
                                                                                    strParental_Rating = .NodeValue ' 161
                                                                                    Dim varTmp As Variant
                                                                                    varTmp = "SELECT CODENO,DESCRIPTION FROM CD029 " & _
                                                                                                    " WHERE " & strParental_Rating & _
                                                                                                    " BETWEEN AGE1 AND AGE2 " & _
                                                                                                    " ORDER BY AGE1"
                                                                                    varTmp = cnCOMM.Execute(varTmp).GetString(adClipString, 1, ",", "", "")
                                                                                    varTmp = Split(varTmp, ",")
                                                                                    strParental_Rating = varTmp(0)
                                                                                    strParental_Rating_Name = varTmp(1)
                                                                                    varTmp = Empty
                                                                            
                                                                            Case "ChannelNumber" ' 162
                                                                                    strCH_NO = .NodeValue
                                                                                    
                                                                            Case "ChannelShortName" ' 162
                                                                                    If CP_Str(strLang162, "eng") Then
                                                                                        strEngShortName = .NodeValue
                                                                                    ElseIf CP_Str(strLang162, "chi") Or CP_Str(strLang162, "cht") Then
                                                                                        strShortName = .NodeValue
                                                                                    End If
                                                                                    
                                                                            Case "ChannelName" ' 162
                                                                                    If CP_Str(strLang162, "eng") Then
                                                                                        strEngName = .NodeValue
                                                                                    ElseIf CP_Str(strLang162, "chi") Or CP_Str(strLang162, "cht") Then
                                                                                        StrName = .NodeValue
                                                                                    End If
                                                                                    
                                                                            Case "DvbServiceId" ' 162
                                                                                    strServiceId = .NodeValue
                                                                                    
                                                                            Case "TransportId" ' 162
                                                                                    strTransportId = .NodeValue
                                                                            
                                                                            Case "ChannelDescription" ' 162
                                                                                    If CP_Str(strLang162, "eng") Then
                                                                                        strDescription = .NodeValue
                                                                                    ElseIf CP_Str(strLang162, "chi") Or CP_Str(strLang162, "cht") Then
                                                                                        strEngDescription = .NodeValue
                                                                                    End If
                                                                            
                                                                            Case "ChannelProviderName" ' 162
                                                                                    strProviderName = .NodeValue
                                                                End Select
                                        End Select
                                        If frmMain.chkShowXML.Value Then
                                            ctlList.AddItem "Name: " & .ParentNode.NodeName
                                            ctlList.AddItem "Text: " & .NodeValue
                                        End If
                                Case 4 ' Node cdata type
                                        If frmMain.chkShowXML.Value Then ctlList.AddItem "CDATA: " & .NodeValue
                                Case Else ' Node other case
                                        If frmMain.chkShowXML.Value Then ctlList.AddItem objElement.NodeType & ": " & .NodeName
            End Select
            Set objNodeLst = .childNodes
'            DoEvents
    End With
    With objNodeLst
            For lngLoop = 0 To .length - 1
                PrcNode .Item(lngLoop), ctlList
                If lngLoop Mod 20 = 0 Then DoEvents
            Next
            ctlList.ListIndex = ctlList.ListCount - 1
    End With
  Exit Function
ChkErr:
    ErrHandle "mod_XML", "PrcNode"
End Function

Private Sub GetAttributes(ByRef objElement As Object, ByRef ctlList As ListBox)  ' 處理 XML Attributes 屬性
  On Error GoTo ChkErr
    Dim intLoop As Integer
    With objElement
            Select Case str_XML_Class
                        Case "CAPPV"
                                For intLoop = 0 To .Attributes.length - 1
                                    If intLoop Mod 10 = 0 Then DoEvents
                                    Select Case .NodeName
                                                Case "ProductManagementData"
                                                        If CP_Str(.Attributes(intLoop).NodeName, "creationDate") Then
                                                            strCreationDate = .Attributes(intLoop).NodeValue
                                                        End If
                                                Case "ProductType"
                                                        If blnAdding Then
                                                            If CP_Str(.Attributes(intLoop).NodeName, "productKind") Then
                                                                rsSO158("ProductKind") = .Attributes(intLoop).NodeValue
                                                            End If
                                                            If CP_Str(.Attributes(intLoop).NodeName, "impulsiveFlag") Then
                                                                rsSO158("impulsiveFlag") = .Attributes(intLoop).NodeValue
                                                            End If
                                                            If CP_Str(.Attributes(intLoop).NodeName, "specialPpv") Then
                                                                rsSO158("specialPpv") = .Attributes(intLoop).NodeValue
                                                            End If
                                                        End If
                                                Case "ProductStatus"
                                                        If CP_Str(.Attributes(intLoop).NodeName, "suspendedFlag") Then
                                                            rsSO158("suspendedFlag") = .Attributes(intLoop).NodeValue
                                                        End If
                                                        If CP_Str(.Attributes(intLoop).NodeName, "soldFlag") Then
                                                            rsSO158("soldFlag") = .Attributes(intLoop).NodeValue
                                                        End If
                                                Case "SalePeriod"
                                                        If blnAdding Then
                                                            If CP_Str(.Attributes(intLoop).NodeName, "beginTime") Then
                                                                rsSO158("SaleBeginTime") = .Attributes(intLoop).NodeValue
                                                            End If
                                                            If CP_Str(.Attributes(intLoop).NodeName, "endTime") Then
                                                                rsSO158("SaleEndTime") = .Attributes(intLoop).NodeValue
                                                            End If
                                                        End If
                                                Case "EpgPrice"
                                                        If CP_Str(.Attributes(intLoop).NodeName, "moneyUnit") Then
                                                            rsSO158("moneyUnit") = .Attributes(intLoop).NodeValue
                                                        End If
                                                Case "ValidityPeriod"
                                                        If CP_Str(.Attributes(intLoop).NodeName, "beginTime") Then
                                                            rsSO158("ValidityPeriodBeginTime") = .Attributes(intLoop).NodeValue
                                                        End If
                                                        If CP_Str(.Attributes(intLoop).NodeName, "endTime") Then
                                                            rsSO158("ValidityPeriodEndTime") = .Attributes(intLoop).NodeValue
                                                        End If
                                                Case "ProductText"
                                                        If CP_Str(.Attributes(intLoop).NodeName, "language") Then
                                                            strLanguage = .Attributes(intLoop).NodeValue
                                                        End If
                                    End Select
                                    If frmMain.chkShowXML.Value Then
                                        ctlList.AddItem .NodeName & vbTab & _
                                                                .Attributes(intLoop).NodeName & " = " & _
                                                                .Attributes(intLoop).NodeValue
                                    End If
                                Next
                                
                        Case "CASUB"
                                For intLoop = 0 To .Attributes.length - 1
                                    If frmMain.chkShowXML.Value Then
                                        ctlList.AddItem .NodeName & vbTab & _
                                                            .Attributes(intLoop).NodeName & " = " & _
                                                            .Attributes(intLoop).NodeValue
                                    End If
                                    If intLoop Mod 10 = 0 Then DoEvents
                                Next
                                
                        Case "NAGRA"
                                For intLoop = 0 To .Attributes.length - 1
                                    If intLoop Mod 10 = 0 Then DoEvents
                                    Select Case .NodeName
                                                Case "BroadcastData"
                                                        If CP_Str(.Attributes(intLoop).NodeName, "creationDate") Then
                                                            strCreationDate = .Attributes(intLoop).NodeValue
                                                        End If
                                                        
                                                Case "Event"
                                                        If CP_Str(.Attributes(intLoop).NodeName, "beginTime") Then
                                                            If blnAdding Then Update_Nagra_SO159
'                                                               A. 取e.<EventType>= 'P'的才insert EPG XML資料檔(SO159)，
'                                                                   不需考慮是否有無重覆資料，一樣完全拆解至SO159。
'                                                               B. 需依據Channel資訊(SO162)的PK(ChannelID)判斷是否有重覆
'                                                                   若比對不存在 , 則insert一筆資料至SO162
'                                                                   若比對已存在 , 則將該筆資料完整update至SO162
'                                                               C. 取e.<EventType>≠ ’P’的才insert Sub時段資料檔(SO161)，
'                                                                   需判斷SO161的PK(EevntID)是否有重覆
'                                                                   若比對不存在，則insert一筆資料至SO161
'                                                                   若比對已存在，則將該筆資料完整update至SO161
                                                            blnAdding = True
                                                            rsSO159.AddNew
                                                            ctlList.Clear
                                                            rsSO159("EventBeginTime") = .Attributes(intLoop).NodeValue ' SO159
                                                            strEvent_BeginTime = .Attributes(intLoop).NodeValue ' SO161
                                                            
                                                        End If
                                                        If CP_Str(.Attributes(intLoop).NodeName, "duration") Then
                                                            rsSO159("Duration") = .Attributes(intLoop).NodeValue
                                                            lngDuration = .Attributes(intLoop).NodeValue ' 161
                                                        End If
                                                        
                                                Case "EpgText"
                                                        If CP_Str(.Attributes(intLoop).NodeName, "language") Then
                                                            strLanguage = .Attributes(intLoop).NodeValue
                                                            
                                                            strLang161 = .Attributes(intLoop).NodeValue ' 161 的 Langauge
                                                        End If
                                                        
                                                Case "ExtendedInfo"
                                                        If CP_Str(.Attributes(intLoop).NodeName, "name") Then
                                                            If strDirector_and_Actor = Empty Then
                                                                strDirector_and_Actor = .Attributes(intLoop).NodeValue
                                                            Else
                                                                strDirector_and_Actor = strDirector_and_Actor & vbCrLf & _
                                                                                                        .Attributes(intLoop).NodeValue
                                                            End If
                                                            If strDirector_Actor = Empty Then
                                                                strDirector_Actor = .Attributes(intLoop).NodeValue
                                                            Else
                                                                strDirector_Actor = strDirector_Actor & vbCrLf & _
                                                                                                        .Attributes(intLoop).NodeValue
                                                            End If
                                                        End If
                                                        
                                                Case "Content"
                                                        If CP_Str(.Attributes(intLoop).NodeName, "nibble1") Then
                                                            rsSO159("ContentNibble1") = .Attributes(intLoop).NodeValue
                                                            strContentNibble1 = .Attributes(intLoop).NodeValue
                                                        End If
                                                        If CP_Str(.Attributes(intLoop).NodeName, "nibble2") Then
                                                            rsSO159("ContentNibble2") = .Attributes(intLoop).NodeValue
                                                            strContentNibble2 = .Attributes(intLoop).NodeValue
                                                        End If
                                                        
                                                Case "User"
                                                        If CP_Str(.Attributes(intLoop).NodeName, "nibble1") Then
                                                            rsSO159("UserNibble1") = .Attributes(intLoop).NodeValue
                                                            strUserNibble1 = .Attributes(intLoop).NodeValue
                                                        End If
                                                        If CP_Str(.Attributes(intLoop).NodeName, "nibble2") Then
                                                            rsSO159("UserNibble2") = .Attributes(intLoop).NodeValue
                                                            strUserNibble2 = .Attributes(intLoop).NodeValue
                                                        End If
                                                        
                                                Case "ChannelText"
                                                        If CP_Str(.Attributes(intLoop).NodeName, "language") Then
                                                            strLang162 = .Attributes(intLoop).NodeValue
                                                        End If
                                                        
'                                                Case "TransportId"
'                                                        If CP_Str(.Attributes(intLoop).NodeName, "originalNetworkId") Then
'                                                            strTransportId= .Attributes(intLoop).nodeValue
'                                                        End If
                                    End Select
                                    If frmMain.chkShowXML.Value Then
                                        ctlList.AddItem .NodeName & vbTab & _
                                                            .Attributes(intLoop).NodeName & " = " & _
                                                            .Attributes(intLoop).NodeValue
                                    End If
                                Next
            End Select
    End With
  Exit Sub
ChkErr:
    ErrHandle "mod_XML", "GetAttributes"
End Sub

'6.  XML檔拆解至SO158、SO159之其他注意事項
'   (3) XML檔裡所有 Tag 裏面，只要有關時間日期的資料一律是 UTC 時間(格林威治標準時間)，
'       必須轉成 Taiwan Local 時間 ( 即原UTC + 8 小時 )後回填資料庫。
Private Function Add8H(strOriDate As String) As String
  On Error GoTo ChkErr
    Add8H = Format(strOriDate, "####/##/## ##:##:##")
    Add8H = DateAdd("h", 8, Add8H)
    Add8H = Format(Add8H, "YYYYMMDDHHMMSS")
  Exit Function
ChkErr:
    ErrHandle "mod_XML", "Add8H"
End Function

Public Function GetXMLobj(ByRef objDOMdoc As Object) As Boolean ' 建立 M$ DOM 物件
  On Error Resume Next
    Set objDOMdoc = CreateObject("Msxml2.DOMDocument.4.0")
    If Err.Number <> 0 Then
        Err.Clear
        Set objDOMdoc = CreateObject("Microsoft.XMLDOM")
    End If
    GetXMLobj = (Err.Number = 0)
End Function
