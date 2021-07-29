Attribute VB_Name = "mod_XML"
' PR : Hammer
' Start date : 2005/07/30
' Last Modify : 2005/08/02

Option Explicit

Public rsSO158 As New ADODB.Recordset ' For CAPPV~.. XML �ϥ� , SO158 (DexTmp)
Public rsSO163 As New ADODB.Recordset ' For CASUB~.. XML �ϥ� , SO163 (SubChannel)

Public rsSO159 As New ADODB.Recordset ' For Nagra~.. XML �ϥ� , SO159 (EpgTmp)
Public rsSO162 As New ADODB.Recordset ' For Nagra~.. XML �ϥ� , SO162 (ChannelInfo)
Public rsSO161 As New ADODB.Recordset ' For Nagra~.. XML �ϥ� , SO161 (Subscription)

Public str_XML_Class As String ' XML ���� , �� XML �ɦW�e 5 �X�ӰϧO
Public objDOM As Object ' M$ ��󪫥�ҫ� (DOM)

Public strCreationDate As String
Private strLanguage As String ' 158 , 159 �@��
Private strChannelId As String
Private blnAdding As Boolean
Private strDirector_and_Actor As String ' 159
Private strActor As String
Public Const strUpdEn = "�t�γB�z"

' SO162 �ϥ��ܼ�
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

' SO161 �ϥ��ܼ�
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

'3.  �`�ت�XML�ɤ��ɮשR�W��h
'(1) �p���I�O(PPV)�G CAPPV�褸�~���ɤ���. XML    �ҡGCAPPV20050520083000.XML
'(2) �I�O�W�D(SUB)�G CASUB�褸�~���ɤ���.XML    �ҡGCASUB20050520173015.XML
'(3) EPG Wrapper �G Nagra_�褸�~���ɤ���. XML   �ҡGNagra_20050530031019.XML

Private Function Open158() As Boolean ' �}�� SO158 ( DexTmp ) ��ƿ��� CAPPV~.xml
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

Private Function Open163() As Boolean ' �}�� SO163 ( SubChannel ) ��ƿ��� CASUB~.xml
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

Private Function Open159() As Boolean ' �}�� SO159 ( EpgTmp ) ��ƿ��� NAGRA~.xml
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

Private Function Open162(Optional blnFlag As Boolean = False) As Boolean ' �}�� SO162 ( ChannelInfo ) ��ƿ��� NAGRA~.xml
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

Private Function Open161(Optional blnFlag As Boolean = False) As Boolean ' �}�� SO161 ( Subscription ) ��ƿ��� NAGRA~.xml
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

' �z�L DOM ���J XML �ɮ�
Public Function LoadXML(strFile As String, _
                                        Optional blnValidate As Boolean = False, _
                                        Optional ByRef ctlList As ListBox) As Boolean
  On Error GoTo ChkErr
    Dim str2crlf As String
    str2crlf = vbCrLf & vbCrLf
    With objDOM
            .async = True ' �D�P�B����
            .validateOnParse = blnValidate ' ���� Vadidation Parser
            LoadXML = .Load(strFile)
            If Not LoadXML Then ' �� XML �ɮ׵o�Ϳ��~ , �h��ܰT��
                MsgBox "���~�X : " & .parseError.errorCode & str2crlf & _
                                "���~��] : " & .parseError.reason & vbCrLf & _
                                "���~�� : " & .parseError.Line & _
                                "  ( ��m : " & .parseError.linepos & " )" & str2crlf & _
                                "���~�ӷ� : " & .parseError.srcText & str2crlf & _
                                "���~�ɮ� : " & .parseError.url & _
                                "  ( ��m : " & .parseError.filepos & " )", vbInformation, "�T��"
            End If
    End With
    str2crlf = Empty
    ctlList.Clear
    ' XML �ɮ׫D�P�B�}�ҫ� , �}�� XML �ɮ׹�������ƿ��� , �H�ѫ����ƳB�z
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

Public Function Update_XML_Table() As Boolean ' �N XML ��Ʀs�ɦܹ����� Table ��
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

'B. �ݨ̾�Channel��T(SO162)��PK(ChannelID)�P�_�O�_������
'    �Y��藍�s�b , �hinsert�@����Ʀ�SO162
'    �Y���w�s�b , �h�N�ӵ���Ƨ���update��SO162
Private Sub Update_Nagra_SO162() ' ��s SO162
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

Private Sub Clear_SO162_Variant() ' �M�� SO162 �ҨϥΤ��ܼ�
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

'C. ��e.<EventType>�� ��P�����~insert Sub�ɬq�����(SO161)�A
'    �ݧP�_SO161��PK(EevntID)�O�_������
'    �Y��藍�s�b , �hinsert�@����Ʀ�SO161
'    �Y���w�s�b , �h�N�ӵ���Ƨ���update��SO161
Private Sub Update_Nagra_SO161() ' ��s SO161
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

Sub Clear_SO161_Variant() ' �M�� SO161 �ҨϥΤ��ܼ�
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

'6. XML�ɩ�Ѧ�SO158�BSO159����L�`�N�ƶ�
'   (1) �����U�Ӫ�XML�ɮסA�䤺�e���Ƿ|������r(encoding="UTF-8" �榡)�A�G�ݪ`�N��s�W�ܸ�Ʈw�ɡA�ݰ��榡�W���ഫ��A�s�C
'   (2) �����Ѧh�y���ɡA�h�ݼW�[�P�_�p�U�G
'       A.  ��Text.Language='chi' �� 'cht' �ɡA�hinsert�ܤ������C
'       B.  ��Text.Language='eng'�ɡA�hinsert�ܭ^�����C
'   (3) XML�ɸ̩Ҧ� Tag �ح��A�u�n�����ɶ��������Ƥ@�߬O UTC �ɶ�(��L�ªv�зǮɶ�)�A
'       �����ন Taiwan Local �ɶ� ( �Y��UTC + 8 �p�� )��^���Ʈw�C
Public Function PrcNode(ByRef objElement As Object, ByRef ctlList As ListBox) ' �B�z XML Node �`�I
  On Error GoTo ChkErr
    Dim lngLoop As Long
    Dim objNodeLst As Object
'   (4) �Ш̷ӡiP7 Schema�jSO158�BSO163�BSO159�BSO162�BSO161��ƪ�B�z�W�h�C
    With objElement
            Select Case .NodeType ' Node Tyype
                                Case 1 ' Node elemnts type
                                        GetAttributes objElement, ctlList
                                Case 3 ' Node text type
                                        Select Case str_XML_Class
'                                                    (1) �N�ɦW��"CAPPV"�}�Y��XML�ɡA���䳡���һݪ����insert��DEX XML�����(SO158)
'                                                       A. CAPPV�ҵ���Dex XML�ɡA�Y�Ϧ����Ъ���ơA�h�@�˧�����Ѧ�SO158�C
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
                                                                
'                                                    (2) �N�ɦW��"CASUB"�}�Y��XML�ɡA���䳡���һݪ����insert��Sub Channel������(SO163)
'                                                       A. CASUB�ҵ���Dex XML�ɡA�ݨ̾�SO163��ƪ�PK(ProductID�BChannelID)�P�_�O�_�����СC
'                                                       B. �Y��藍�s�b�A�hinsert�@����Ʀ�SO163�C
'                                                       C. �Y���w�s�b�A�h�N�ӵ���Ƨ���update��SO163�C
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
                                                                
'                                                    (3) �N�ɦW��"Nagra_"�}�Y��XML�ɡA���䳡���һݪ����insert��SO159�BSO162�BSO161�C
'                                                       A. ��e.<EventType>= 'P'���~insert EPG XML�����(SO159)�A
'                                                           ���ݦҼ{�O�_���L���и�ơA�@�˧�����Ѧ�SO159�C
'                                                       B. �ݨ̾�Channel��T(SO162)��PK(ChannelID)�P�_�O�_������
'                                                           �Y��藍�s�b, �hinsert�@����Ʀ�SO162
'                                                           �Y���w�s�b, �h�N�ӵ���Ƨ���update��SO162
'                                                       C. ��e.<EventType>�� 'P'���~insert Sub�ɬq�����(SO161)�A
'                                                           �ݧP�_SO161��PK(EevntID)�O�_������
'                                                           �Y��藍�s�b, �hinsert�@����Ʀ�SO161
'                                                           �Y���w�s�b, �h�N�ӵ���Ƨ���update��SO161
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

'                                                                           (4) EPG Wrapper XML�̪�<ExtendedInfoName>�B<ExtendedInfo>�B<Description>�����e�G
'                                                                               A.��ѫ᤺�e , �@�_�s���SO159.EpgDesc���
'                                                                               B.  �ݭn���_���x�s�A�H�Q�`�ت�e�{�ɥi�@�֪`������[�C

'                                                                           C.�|��:
'                                                                               <EpgText language="chi">
'                                                                                   <Name>�ܤH(��) </Name>
'                                                                                   <Description>�����H�ޮa�w�w�|�iù���·G�� ���j�A�쥻�D�����H�����A�Ӧb�¤i�۳B���U�A�j�a�]�����a���ǥL�A��L�R�W�H���ȰҹF��M�w�P�o���B�A�L�������ä����l������������ܦ��H���A�L�O�_��p�@�H�v�a�P�߷R���H�r�u�@�ͩO�H</Description>
'                                                                                   <ExtendedInfo name="�ɺt:">�J�������ۥ� </ExtendedInfo>
'                                                                                   <ExtendedInfo name="�t��:">ù���·G���B�P�B���j�ï� </ExtendedInfo>
'                                                                               </EpgText>

'                                                                           ��ѥX�Ӫ�SO159.EpgDesc���Ȭ��G4
'                                                                               �ɺt:�J�������ۥ�
'                                                                               �t��:ù���·G���B�P�B���j�ï�
'                                                                               �����H�ޮa�w�w�|�iù���·G�� ���j�A�쥻�D�����H�����A�Ӧb�¤i�۳B���U�A�j�a�]�����a���ǥL�A��L�R�W�H���ȰҹF��M�w�P�o���B�A�L�������ä����l������������ܦ��H���A�L�O�_��p�@�H�v�a�P�߷R���H�r�u�@�ͩO�H

                                                                            Case "ExtendedInfo" ' 159
                                                                                    strDirector_and_Actor = strDirector_and_Actor & .NodeValue
                                                                                    strDirector_Actor = strDirector_Actor & .NodeValue
                                                                                    
                                                                            Case "ParentalRating" ' 159 & 161
                                                                                    rsSO159("ParentalRating") = .NodeValue

'                                                                                   Parentalrating��`�ت�xml�ɸ̥X�{���d��Ȭ�0~15�����A
'                                                                                   �ݹ��CD029�`�ص��ťN�X�ɪ��~�֤����A����ŦX���`�ص��ťN�X�B
'                                                                                   �W�٦^��SO154.ParentalRating�BParentalRatingName�C
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

Private Sub GetAttributes(ByRef objElement As Object, ByRef ctlList As ListBox)  ' �B�z XML Attributes �ݩ�
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
'                                                               A. ��e.<EventType>= 'P'���~insert EPG XML�����(SO159)�A
'                                                                   ���ݦҼ{�O�_���L���и�ơA�@�˧�����Ѧ�SO159�C
'                                                               B. �ݨ̾�Channel��T(SO162)��PK(ChannelID)�P�_�O�_������
'                                                                   �Y��藍�s�b , �hinsert�@����Ʀ�SO162
'                                                                   �Y���w�s�b , �h�N�ӵ���Ƨ���update��SO162
'                                                               C. ��e.<EventType>�� ��P�����~insert Sub�ɬq�����(SO161)�A
'                                                                   �ݧP�_SO161��PK(EevntID)�O�_������
'                                                                   �Y��藍�s�b�A�hinsert�@����Ʀ�SO161
'                                                                   �Y���w�s�b�A�h�N�ӵ���Ƨ���update��SO161
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
                                                            
                                                            strLang161 = .Attributes(intLoop).NodeValue ' 161 �� Langauge
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

'6.  XML�ɩ�Ѧ�SO158�BSO159����L�`�N�ƶ�
'   (3) XML�ɸ̩Ҧ� Tag �ح��A�u�n�����ɶ��������Ƥ@�߬O UTC �ɶ�(��L�ªv�зǮɶ�)�A
'       �����ন Taiwan Local �ɶ� ( �Y��UTC + 8 �p�� )��^���Ʈw�C
Private Function Add8H(strOriDate As String) As String
  On Error GoTo ChkErr
    Add8H = Format(strOriDate, "####/##/## ##:##:##")
    Add8H = DateAdd("h", 8, Add8H)
    Add8H = Format(Add8H, "YYYYMMDDHHMMSS")
  Exit Function
ChkErr:
    ErrHandle "mod_XML", "Add8H"
End Function

Public Function GetXMLobj(ByRef objDOMdoc As Object) As Boolean ' �إ� M$ DOM ����
  On Error Resume Next
    Set objDOMdoc = CreateObject("Msxml2.DOMDocument.4.0")
    If Err.Number <> 0 Then
        Err.Clear
        Set objDOMdoc = CreateObject("Microsoft.XMLDOM")
    End If
    GetXMLobj = (Err.Number = 0)
End Function
