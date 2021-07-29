Attribute VB_Name = "mod_XML"
Option Explicit

Public objDOM As Object ' M$ 文件物件模型 (DOM)
Public rsVirtual As Object

' 透過 DOM 載入 XML 檔案
Private Function LoadXML(strXML As String, _
                                        Optional blnValidate As Boolean = False) As Boolean
  On Error GoTo ChkErr
    Dim str2crlf As String
    str2crlf = vbCrLf & vbCrLf
    With objDOM
            .async = True ' 非同步執行
            .validateOnParse = blnValidate ' 關掉 Vadidation Parser
            
            '判斷傳進的是 XML 資料串流 , 亦或是 XML 路徑+檔案名稱
            If InStr(1, strXML, "?xml", vbTextCompare) Then
                LoadXML = .LoadXML(strXML)
            Else
                LoadXML = .Load(strXML)
            End If
            
            If Not LoadXML Then ' 當 XML 檔案發生錯誤 , 則顯示訊息
                strErrorMessage = "錯誤碼 : " & .parseError.errorCode & str2crlf & _
                                                "錯誤原因 : " & .parseError.reason & vbCrLf & _
                                                "錯誤行 : " & .parseError.Line & _
                                                "  ( 位置 : " & .parseError.linepos & " )" & str2crlf & _
                                                "錯誤來源 : " & .parseError.srcText & str2crlf & _
                                                "錯誤檔案 : " & .parseError.url & _
                                                "  ( 位置 : " & .parseError.filepos & " )"
'                MsgBox strErrorMessage, vbInformation, "訊息"
            End If
    End With
    str2crlf = Empty
  Exit Function
ChkErr:
    ErrHandle "mod_XML", "LoadXML"
End Function
        
' 處理 XML 檔案 , 在存至虛擬記憶體資料錄
Public Function ParseXML(strXMLfile As String) As Boolean
  On Error GoTo ChkErr
    Dim strRet As String
    ParseXML = False
    If GetDOMdocObj(objDOM, strRet) Then
        LoadXML strXMLfile
        Set rsVirtual = CreateObject("ADODB.Recordset")
        With rsVirtual
                .CursorLocation = 3
                .CursorType = 3
                .LockType = 3
                PrcNode objDOM.documentElement.childNodes(0).childNodes(0)
                .Open
        '        PrcNode objDOM.documentElement.childNodes(1)
                PrcNode objDOM.documentElement.childNodes(1).childNodes(0)
                .Update
                .MoveFirst
        End With
        ParseXML = (Err.Number = 0)
    Else
'        MsgBox strRet, vbInformation, "訊息"
    End If
  Exit Function
ChkErr:
    ParseXML = False
    ErrHandle "mod_XML", "ParseXML"
End Function

' 處理 XML Node 節點
Private Function PrcNode(ByRef objElement As Object)
  On Error GoTo ChkErr
    Dim lngLoop As Long
    Dim objNodeLst As Object
    With objElement
            Select Case .NodeType ' Node Tyype
                        Case 1 ' Node elemnts type
                                GetAttributes objElement
                        Case 3 ' Node text type
'                                Debug.Print .ParentNode.NodeName & vbTab & .NodeValue
                                Select Case LCase(.ParentNode.NodeName)
                                            Case "queryresult"
                                                        rsVirtual.AddNew
                                                        rsVirtual("QueryResult") = .NodeValue
                                            Case "operresult"
                                                        rsVirtual.AddNew
                                                        rsVirtual("OperResult") = .NodeValue
                                            Case "faultreason"
                                                        rsVirtual("FaultReason") = .NodeValue
                                            Case "groupid"
                                                        rsVirtual("GROUPID") = .NodeValue
                                            Case "cmstatus"
                                                        rsVirtual("CMStatus") = .NodeValue
                                            Case "cmip"
                                                        rsVirtual("CMIP") = .NodeValue
                                            Case "cmupfreq"
                                                        rsVirtual("CMUPFREQ") = .NodeValue
                                            Case "cmdownfreq"
                                                        rsVirtual("CMDOWNFREQ") = .NodeValue
                                            Case "cmrecpw"
                                                        rsVirtual("CMRECPW") = .NodeValue
                                            Case "cmtranspw"
                                                        rsVirtual("CMTRANSPW") = .NodeValue
                                            Case "cmsnr"
                                                        rsVirtual("CMSNR") = .NodeValue
                                            Case "cmonlinetimes"
                                                        rsVirtual("CMOnlineTimes") = .NodeValue
                                            Case "cminoctets"
                                                        rsVirtual("CMInOctets") = .NodeValue
                                            Case "cmoutoctets"
                                                        rsVirtual("CMOutOctets") = .NodeValue
                                            Case "cminerrors"
                                                        rsVirtual("CMInErrors") = .NodeValue
                                            Case "cmouterrors"
                                                        rsVirtual("CMOutErrors") = .NodeValue
                                            Case "pcip"
                                                        rsVirtual("PCIP") = .NodeValue
                                            Case "cmtsrecpw"
                                                        rsVirtual("CMTSRECPW") = .NodeValue
                                            Case "cmtsrxsnr"
                                                        rsVirtual("CMTSRXSNR") = .NodeValue
                                            Case "cmtsupmod"
                                                        rsVirtual("CMTSUPMOD") = .NodeValue
'                                            Case "cmtsdownmod" ' ***
'                                                        rsVirtual("CMTSDOWNMOD") = .NodeValue
'                                            Case "cmdesc" ' ***
'                                                        rsVirtual("CMDESC") = .NodeValue
'                                            Case "pcips" ' ***
'                                                        rsVirtual("PCIPS") = .NodeValue
                                            Case "hfctype"
                                                        rsVirtual("HFCType") = .NodeValue
                                            Case "offlinetime"
                                                        rsVirtual("OffLineTime") = .NodeValue
                                            Case "onlinetime"
                                                        rsVirtual("OnLineTime") = .NodeValue
                                            Case "detcttime"
                                                        rsVirtual("DetctTime") = .NodeValue
'                                            Case "cusacco"
'                                                        rsVirtual("CUSACCO") = .NodeValue
'                                            Case "cuspassword"
'                                                        rsVirtual("CUSPASSWORD") = .NodeValue
                                            Case "cmmac"
                                                        rsVirtual("CMMAC") = .NodeValue
                                            Case "ctip"
                                                        rsVirtual("CTIP") = .NodeValue
                                            Case "nodeid"
                                                        rsVirtual("NODEID") = .NodeValue
                                            Case "pconlinetime" ' ?? CMPCLog.xml
                                                        rsVirtual("PCOnlineTime") = .NodeValue
                                            Case "pcofflinetime"
                                                        rsVirtual("PCOfflineTime") = .NodeValue
                                End Select
'                                Debug.Print "Name: " & .ParentNode.NodeName & vbTab & "Text: " & .NodeValue
                        Case 4 ' Node cdata type
                                Debug.Print "CDATA: " & .NodeValue
                        Case Else ' Node other case
                                Debug.Print objElement.NodeType & ": " & .NodeName
            End Select
            Set objNodeLst = .childNodes
'            DoEvents
    End With
    With objNodeLst
            For lngLoop = 0 To .length - 1
                PrcNode .Item(lngLoop)
'                If lngLoop Mod 20 = 0 Then DoEvents
            Next
    End With
  Exit Function
ChkErr:
    ErrHandle "mod_XML", "PrcNode"
End Function

' 處理 XML Attributes 屬性
Private Sub GetAttributes(ByRef objElement As Object)
  On Error GoTo ChkErr
    Dim intLoop As Integer
    With objElement
            For intLoop = 0 To .Attributes.length - 1
                If CP_Str(.ParentNode.NodeName, "xs:sequence") Then
                    If CP_Str(.Attributes(intLoop).NodeName, "name") Then
                        rsVirtual.Fields.Append .Attributes(intLoop).NodeValue, 202, 2000
                        Debug.Print .Attributes(intLoop).NodeValue
                    End If
'                    Debug.Print .NodeName & vbTab & _
                                        .Attributes(intLoop).NodeName & " = " & _
                                        .Attributes(intLoop).NodeValue
                End If
'                If intLoop Mod 10 = 0 Then DoEvents
            Next
    End With
  Exit Sub
ChkErr:
    ErrHandle "mod_XML", "GetAttributes"
End Sub
