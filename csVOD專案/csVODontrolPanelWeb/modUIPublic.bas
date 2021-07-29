Attribute VB_Name = "modUIPublic"
Option Explicit

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
