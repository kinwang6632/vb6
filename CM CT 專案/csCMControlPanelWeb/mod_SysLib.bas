Attribute VB_Name = "mod_SysLib"
Option Explicit

'Public Const FormName = "mduNagraSTB"
Public Const FormName = "CMconsole"
Public Const strCrlf = vbCrLf & vbCrLf
Public strCPno As String
Public blnAccChk As Boolean

Public Function ChkDTok() As Boolean
'  On Error GoTo ChkErr
  On Error Resume Next
   'If Not ChkDTok Then Exit Sub
    If TypeOf Screen.ActiveControl Is GiDate Or _
        TypeOf Screen.ActiveControl Is GiTime Or _
        TypeOf Screen.ActiveControl Is GiYM Then
        If Len(Trim(Screen.ActiveControl.GetValue)) = 0 Then
            ChkDTok = True
        Else
            ChkDTok = Screen.ActiveControl.RaiseValidateEvent
        End If
    Else
        ChkDTok = True
    End If
'  Exit Function
'ChkErr:
'    Call ErrSub("Sys_Lib", "ChkDTok")
End Function

Private Function InsertToOracle(strTable As String, _
            rsSox As ADODB.Recordset, _
            Optional strWhere As String, _
            Optional lngAffected As Long = 0) As Boolean
    On Error GoTo ChkErr
    Dim intLoop As Integer
    Dim rsTmp As New ADODB.Recordset
    Dim strField As String
    Dim strValue As String
    Dim strSQL As String
    Dim strFullField As String
        If rsSox.State = adStateClosed Then Exit Function
        If strWhere = "" Then strWhere = " Where 1 = 0 "
        If Not GetRS(rsTmp, "Select * From " & strTable & " " & strWhere, , adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
        Call ChkHaveField(strFullField, "", rsSox)
        If rsTmp.RecordCount = 0 Then
            For intLoop = 0 To rsTmp.Fields.Count - 1
                '檢查是不有該欄位
                If ChkHaveField(strFullField, rsTmp(intLoop).Name, rsSox) Then
                    If Not IsNull(rsSox(rsTmp(intLoop).Name)) Then
                        strField = strField & "," & rsTmp(intLoop).Name
                        If rsTmp(intLoop).Type = adDBTimeStamp Then
                            strValue = strValue & "," & GetNullString(rsSox(rsTmp(intLoop).Name).Value, giDateV, giOracle, True)
                        Else
                            strValue = strValue & "," & GetNullString(rsSox(rsTmp(intLoop).Name).Value, giStringV, giOracle)
                        End If
                    End If
                End If
            Next
            strValue = "Insert Into " & strTable & " (" & Mid(strField, 2) & ") Values (" & Mid(strValue, 2) & ")"
        Else
            For intLoop = 0 To rsTmp.Fields.Count - 1
                If ChkHaveField(strFullField, rsTmp(intLoop).Name, rsSox) Then
                    If rsTmp(intLoop).Type = adDBTimeStamp Then
                        strField = strField & "," & rsTmp(intLoop).Name & "=" & GetNullString(rsSox(rsTmp(intLoop).Name).Value, giDateV, giOracle, True)
                    Else
                        strField = strField & "," & rsTmp(intLoop).Name & "=" & Replace(GetNullString(rsSox(rsTmp(intLoop).Name).Value, giStringV, giOracle), ",", "'||chr(44)||'")
                    End If
                End If
            Next
            strValue = "Update " & strTable & " Set " & Mid(strField, 2) & " " & strWhere
        End If
        If Not ExecuteCommand(strValue, gcnGi, lngAffected) Then Exit Function
'        If rsTmp.RecordCount = 0 Then rsTmp.AddNew
'        For intLoop = 0 To rsTmp.Fields.Count - 1
'             rsTmp(intLoop) = GetFieldValue(rsSox, rsTmp(intLoop).Name)
'        Next
'        rsTmp.Update
        lngAffected = 1
        Call CloseRecordset(rsTmp)
        InsertToOracle = True
    Exit Function
ChkErr:
    ErrSub FormName, "InsertToOracle"
End Function

Public Sub subShowMessage(strSetType As String, strcmdStatus As String, _
    strErrorCode As String, strErrorMsg As String)
    On Error GoTo ChkErr
        If strcmdStatus = "C" Then
            MsgBox "[" & strSetType & "] 控制訊號已送出完成 !! ", vbInformation, gimsgPrompt
        ElseIf strcmdStatus = "E" Then
            MsgBox strSetType & "失敗!! " & vbCrLf & "錯誤代碼:" & strErrorCode & " ,錯誤原因:" & strErrorMsg, vbInformation, gimsgPrompt
        End If
    Exit Sub
ChkErr:
    ErrSub FormName, "subShowMessage"
End Sub

Public Sub ToolBarButtonClick(ButtonKey As String, _
        objToolBar As Toolbar)
    On Error Resume Next
    Dim strActFormName As String
        If objToolBar.Buttons(ButtonKey).Enabled = False Then Exit Sub
        If objActForm Is Nothing Then Exit Sub
        strActFormName = objActForm.Name
        Select Case ButtonKey
            Case "AddNew"
                Call objActForm.AddNewGo
            Case "Edit"
                Call objActForm.EditGo
            Case "Delete"
                Call objActForm.DeleteGo
            Case "Find"
                Call objActForm.FindGo
            Case "Print"
                Call objActForm.PrintGO
            Case "Save"
                Call objActForm.UpdateGo
            Case "First"
                Call objActForm.FirstGo
            Case "Previous"
                Call objActForm.PreviousGo
            Case "Next"
                Call objActForm.NextGo
            Case "Last"
                Call objActForm.LastGo
            Case "Cancel"
                Call objActForm.CancelGo
        End Select
        If ChkFormLoad(strActFormName) Then If objActForm.Visible And objActForm.Enabled Then objActForm.SetFocus
End Sub

Public Function ChkFormLoad(strForm As String) As Boolean
  On Error Resume Next
    Dim lngLoop As Long
        ChkFormLoad = True
        For lngLoop = 0 To Forms.Count - 1
            If UCase(Forms(lngLoop).Name) = UCase(strForm) Then Exit Function
        Next
        ChkFormLoad = False
End Function

Public Sub FunctionKeyGo(KeyCode As Integer, Shift As Integer, _
        objToolBar As Toolbar)
    On Error GoTo ChkErr
    Dim ButtonKey As String
        If Shift = 0 Then
            Select Case KeyCode
                Case vbKeyF2
                    ButtonKey = "Update"
                Case vbKeyF3
                    ButtonKey = "Find"
                Case vbKeyF6
                    ButtonKey = "AddNew"
                Case vbKeyF11
                    ButtonKey = "Edit"
                Case vbKeyF10
                    ButtonKey = "Delete"
                Case vbKeyF5
                    ButtonKey = "Print"
                Case vbKeyPageUp
                    'ButtonKey = "Previous"
                Case vbKeyPageDown
                    'ButtonKey = "Next"
                Case vbKeyEscape
                    ButtonKey = "Cancel"
            End Select
        ElseIf Shift = 1 Then
            If KeyCode = vbKeyX Then
                ButtonKey = "Cancel"
            End If
        ElseIf Shift = 2 Then
            Select Case KeyCode
                Case vbKeyPageUp
                    ButtonKey = "First"
                Case vbKeyPageDown
                    ButtonKey = "Last"
                Case vbKeyF
                    ButtonKey = "Find"
            End Select
        End If
        If ButtonKey <> "" Then Call ToolBarButtonClick(ButtonKey, objToolBar)
    Exit Sub
ChkErr:
    ErrSub FormName, "FunctionKeyGo"
End Sub
  
Public Sub SetgiList(objGilist As Object, strFldName1 As String, strFldName2 As String, _
            StrTableName As String, Optional lngTop As Long, Optional lngLeft As Long, Optional lngWidth1 As Long, _
            Optional lngWidth2 As Long, Optional lngFldLen1 As Long, Optional lngFldLen2 As Long, _
            Optional strWhere As String, Optional blnStopFlag As Boolean = False)
 '欄位1,欄位2,表格名稱,Top,Left,Width1,Width2,FldLen1,FldLen2
  On Error GoTo ChkErr
    objGilist.SetFldName1 strFldName1
    objGilist.SetFldName2 strFldName2
    objGilist.SetTableName GetOwner & StrTableName
    objGilist.Filter = strWhere
    If lngTop > 0 Then objGilist.Top = lngTop
    If lngLeft > 0 Then objGilist.Left = lngLeft
    If lngWidth1 > 0 Then objGilist.FldWidth1 = lngWidth1
    If lngWidth2 > 0 Then objGilist.FldWidth2 = lngWidth2
    If lngFldLen1 > 0 Then objGilist.FldLen1 = lngFldLen1
    If lngFldLen2 > 0 Then objGilist.FldLen2 = lngFldLen2
    objGilist.FilterStop = blnStopFlag
    Call objGilist.SendConn(gcnGi)
    
  Exit Sub
ChkErr:
    Call ErrSub(FormName, "SetgiList")
End Sub

Public Sub SetgiMulti(objgiMulti As Object, strFldName1 As String, _
        strFldName2 As String, StrTableName As String, strFldCaption1 As String, _
        strFldCaption2 As String, Optional strFilter As String, _
        Optional blnStopFlag As Boolean = False)
 '欄位1,欄位2,Table,中文1,中文2
  On Error GoTo ChkErr
    Call objgiMulti.SendConn(gcnGi)
    objgiMulti.FldName1 = strFldName1
    objgiMulti.FldName2 = strFldName2
    objgiMulti.TableName = GetOwner & StrTableName
    objgiMulti.FldCaption1 = strFldCaption1
    objgiMulti.FldCaption2 = strFldCaption2
    objgiMulti.FilterStop = blnStopFlag
    If strFilter <> "" Then objgiMulti.Filter = strFilter
  Exit Sub
ChkErr:
    Call ErrSub(FormName, "SetgiMulti")
End Sub

Public Sub SetMsQry(objMs As Object, qryStr As String, _
                    Optional FldCaption1 As String = "", _
                    Optional FldCaption2 As String = "", _
                    Optional ColumnCaption As String = "", _
                    Optional ColumnWidth As String = "", _
                    Optional blnStopFlag As Boolean)
  On Error GoTo ChkErr
    With objMs
        .SendConn gcnGi
        .QueryString = qryStr
        .FilterStop = blnStopFlag
         If Len(FldCaption1) > 0 Then .FldCaption1 = FldCaption1
         If Len(FldCaption2) > 0 Then .FldCaption2 = FldCaption2
         If Len(ColumnCaption) > 0 Then .ColumnCaption = ColumnCaption
         If Len(ColumnWidth) > 0 Then .ColumnWidth = ColumnWidth
    End With
  Exit Sub
ChkErr:
    ErrSub FormName, "SetMsQry"
End Sub

Public Function GetCompCodeFilter(objControl As Object, Optional strFilterStr As String) As String
  On Error Resume Next
    If strFilterStr = "" Then Exit Function
    If objControl.Filter = "" Then
        GetCompCodeFilter = "Where CodeNo in (" & strFilterStr & ")"
    Else
        GetCompCodeFilter = objControl.Filter & " And CodeNo in (" & strFilterStr & ")"
    End If
End Function

Public Sub SetgiMultiAddItem(objgiMulti As Object, strCodeNo As String, strDescription As String, Optional strFldCaption1 As String, Optional strFldCaption2 As String)
  On Error GoTo ChkErr
   Dim varCodeNo As Variant
   Dim varDescription As Variant
   Dim varValue As Variant
   Dim intValue As Integer
   objgiMulti.FldCaption1 = strFldCaption1
   objgiMulti.FldCaption2 = strFldCaption2
   objgiMulti.DIY = True
   intValue = 0
   varCodeNo = Split(strCodeNo, ",")
   varDescription = Split(strDescription, ",")
   For Each varValue In varCodeNo
       objgiMulti.AddItem varValue, varDescription(intValue)
       intValue = intValue + 1
   Next
  Exit Sub
ChkErr:
    Call ErrSub(FormName, "SetgiMultiAddItem")
End Sub

Public Sub GiMultiFilter(objMulti As Object, Optional strServiceType As String = "", Optional strCompCode As String = "")
    On Error GoTo ChkErr
    Dim strFilter As String
      objMulti.SetDispStr ""
        If TypeOf objMulti Is GiMulti Then
            If strServiceType <> "" Then
                strFilter = strFilter & " WHERE Nvl(SERVICETYPE,'" & Trim(strServiceType) & "') = '" & Trim(strServiceType) & "'"
            End If
            
            If strCompCode <> "" Then
                strFilter = strFilter & IIf(strFilter = "", " WHERE ", " And ") & " Nvl(COMPCODE," & Trim(strCompCode) & ") = " & Trim(strCompCode) & ""
            End If
            objMulti.Filter = strFilter
        End If
        
    Exit Sub
ChkErr:
    ErrSub "Sys_Lib", "GiMultiFilter"
End Sub

Public Function MustExist(ctl, Optional Flag, Optional strMsg As String, Optional ByRef TabObj As Object, Optional ByVal TabIdx As Integer) As Boolean
  On Error GoTo ChkErr
    MustExist = True
    Flag = IIf(IsMissing(Flag), 0, Flag)
    If strMsg = "" Then
        strMsg = gMsgIsDataOK
    Else
        strMsg = "[" & strMsg & "] 為必要欄位,須有值 !!"
    End If
    ' GiDate
    If Flag = 1 Then
        If ctl.GetValue = "" Then
            If ctl.Enabled Then ctl.SetFocus
            MsgBox strMsg, vbInformation, "訊息"
            MustExist = False
        End If
    ' GiList
    ElseIf Flag = 2 Then
        If ctl.GetCodeNo = "" Or ctl.GetDescription = "" Then
            If ctl.Enabled Then ctl.SetFocus
            MsgBox strMsg, vbInformation, "訊息"
            MustExist = False
        End If
    ' GiAddress2
    ElseIf Flag = 3 Then
        If ctl.IsDataOk = False Or Len(Trim(ctl.GetCodeNo)) = 0 Then
            If ctl.Enabled Then ctl.SetFocus
            MsgBox strMsg, vbInformation, "訊息"
            MustExist = False
        End If
    ' GiAddress1
    ElseIf Flag = 4 Then
        If ctl.AddrNo = 0 Then
            If ctl.Enabled Then
                ctl.SetFocus
            End If
            MsgBox strMsg, vbInformation, "訊息"
            MustExist = False
        End If
    ' Gimulti
    ElseIf Flag = 5 Then
        If ctl.GetDispStr = "" Then
            If ctl.Enabled Then
                ctl.SetFocus
            End If
            MsgBox strMsg, vbInformation, "訊息"
            MustExist = False
        End If
    Else
        If ctl.Text = "" Then
            If Not TabObj Is Nothing Then
                TabObj.Tab = TabIdx
                If ctl.Enabled Then ctl.SetFocus
                MsgBox strMsg, vbInformation, "訊息"
                On Error GoTo 0
                MustExist = False
            Else
                If ctl.Enabled Then ctl.SetFocus
                MsgBox strMsg, vbInformation, "訊息"
                On Error GoTo 0
                MustExist = False
            End If
        End If
    End If
  Exit Function
ChkErr:
    If Err.Number = 5 Then Resume Next: Exit Function
    ErrSub FormName, "MustExist"
End Function

Public Sub objSelectAll(obj As Object)
    On Error GoTo ChkErr
        With obj
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
    Exit Sub
ChkErr:
    ErrSub FormName, "objSelectAll"
End Sub

Public Function GetPremission(strMID As String) As Boolean
    On Error GoTo ChkErr
    Dim rs As New ADODB.Recordset
        If Val(garyGi(4)) = 0 Then GetPremission = True: Exit Function
        If gstrUserPremission = "" Then
            If Not GetRS(rs, "Select Mid From " & GetOwner & "So029 Where Link like 'SO1171%' And Group" & garyGi(4) & " = 1") Then Exit Function
            If Not rs.EOF Then gstrUserPremission = UCase(rs.GetString(, , ",", ",", ""))
        End If
        GetPremission = InStr(1, gstrUserPremission, UCase(strMID)) > 0
    Exit Function
ChkErr:
    ErrSub FormName, "GetPremission"
End Function


