VERSION 5.00
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Begin VB.Form frmSO1100D 
   BorderStyle     =   1  '單線固定
   Caption         =   "客戶資料/客戶服務--客戶註銷 [SO1100D]"
   ClientHeight    =   2205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4800
   Icon            =   "SO1100D.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   1290
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   4440
      Begin Gi_Date.GiDate gdaDelDate 
         Height          =   285
         Left            =   1470
         TabIndex        =   2
         Top             =   315
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin prjGiList.GiList gilDelete 
         Height          =   285
         Left            =   1470
         TabIndex        =   4
         Top             =   765
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FldWidth1       =   405
         FldWidth2       =   1620
         F2Corresponding =   -1  'True
         FilterStop      =   -1  'True
      End
      Begin VB.Label lblDelete 
         Caption         =   "註銷原因"
         Height          =   195
         Left            =   660
         TabIndex        =   3
         Top             =   810
         Width           =   855
      End
      Begin VB.Label lblDelDate 
         Caption         =   "註銷日期"
         Height          =   225
         Left            =   660
         TabIndex        =   1
         Top             =   345
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&X)"
      Height          =   375
      Left            =   3000
      TabIndex        =   6
      Top             =   1620
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "F2.確定"
      Height          =   375
      Left            =   450
      TabIndex        =   5
      Top             =   1620
      Width           =   1215
   End
End
Attribute VB_Name = "frmSO1100D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Programmer: Power Hammer
' Last Modify: 2000/12/30

Private Sub cmdCancel_Click()
  On Error GoTo ChkErr
    Unload Me
  Exit Sub
ChkErr:
    ErrSub Me.Name, "cmdCancel_Click"
End Sub

'Private Sub cmdOK_Click()
'  On Error GoTo CnErr
'    Dim cn2 As New ADODB.Connection
'    If Not MustExist(gdaDelDate, 1, "註銷日期") Then Exit Sub
'    If Not MustExist(gilDelete, 2, "註銷原因") Then Exit Sub
'    With cn2
'        .CursorLocation = adUseServer
'        .Open gcnGi.ConnectionString
'        .BeginTrans
'    End With
'  On Error GoTo ProcErr
'    '(1) 更新各種戶數統計欄位:
'    '.若該客戶狀態為正常 (1), 則:
'    With frmSO1100BMDI
'         If .rsSo002!CustStatusCode = 1 Then
'             '1.  將街道代碼檔(CD017)中, 該客戶裝機地址所對應的街道, 將街道安裝戶數減1
'             If ExecuteSQL("UPDATE " & GetOwner & "CD017 SET INSTCNT = INSTCNT - 1 WHERE CODENO = " & .rsSo001!InstAddrNO, cn2) < giOK Then
'                 cn2.RollbackTrans
'                 cn2.Close
'                 Set cn2 = Nothing
'                 Exit Sub
'             End If
''             If ExecuteSQL("UPDATE " & GetOwner & "CD017 SET INSTCNT = INSTCNT - 1 WHERE CODENO = " & .rsSo001!InstAddrNO, cn2) < giOK Then
''                 cn2.RollbackTrans
''                 cn2.Close
''                 Set cn2 = Nothing
''                 Exit Sub
''             End If
'             '2.  若該客戶為大樓戶(該筆客戶資料的MduId欄位有值)
'             If Not IsNull(.rsSo001!MduId) Then
'                 If .rsSo001!MduId <> 0 Then ' 則將大樓資料檔的該大樓的有效戶數減1
'                     If ExecuteSQL("UPDATE " & GetOwner & "SO017 SET INSTCNT = INSTCNT - 1 WHERE MDUID = '" & .rsSo001!MduId & "'", cn2) < giOK Then
'                         cn2.RollbackTrans
'                         cn2.Close
'                         Set cn2 = Nothing
'                         Exit Sub
'                     End If
'                 End If
'             End If
'         ' 若該客戶狀態為已拆機(3)或已停機(2)
'         ElseIf .rsSo002!CustStatusCode = 3 Or .rsSo002!CustStatusCode = 2 Then              ' 且該客戶為大樓戶
'             If Not IsNull(.rsSo001!MduId) Then
'                 If .rsSo001!MduId <> 0 Then ' 則將大樓資料檔的該大樓的大樓停拆／待拆戶數減1
'                     If ExecuteSQL("UPDATE " & GetOwner & "SO017 SET UNINSTCNT = UNINSTCNT - 1 WHERE MDUID = '" & .rsSo001!MduId & "'", cn2) < giOK Then
'                         cn2.RollbackTrans
'                         cn2.Close
'                         Set cn2 = Nothing
'                         Exit Sub
'                     End If
'                 End If
'             End If
'         End If
'
'         cn2.Execute "INSERT INTO " & GetOwner & "SO101 (CUSTID,CUSTNAME,CUSTSTATUSCODE,CUSTSTATUSNAME,DELDATE,DELCODE,DELNAME,CUSTSTATUSCODEB,CUSTSTATUSNAMEB,UPDTIME,UPDEN,MODEID,COMPCODE,SERVICETYPE) VALUES " & _
'                     "(" & gCustId & ",'" & .rsSo001!Custname & "'," & _
'                     .rsSo002!CustStatusCode & ",'" & _
'                     .rsSo002!CustStatusName & "'," & _
'                      GetNullString(gdaDelDate.Text, giDateV) & "," & _
'                      gilDelete.GetCodeNo2 & ",'" & _
'                      gilDelete.GetDescription2 & "',4,'" & _
'                      GetRsValue("select Description FROM " & GetOwner & "CD035 WHERE CodeNo=4", gcnGi) & "','" & _
'                      GetDTString(RightNow) & "','" & _
'                      garyGi(1) & "',4," & _
'                      gCompCode & ",'" & _
'                      gServiceType & "')"
'
'         .rsSo002!DelDate = SolveDate(gdaDelDate.GetValue) & ""
'         .rsSo002!DelCode = gilDelete.GetCodeNo2 & ""
'         .rsSo002!DelName = gilDelete.GetDescription2 & ""
'         .rsSo002!OldStatusCode = Val(.rsSo002!CustStatusCode)
'         .rsSo002!CustStatusCode = 4
'         .rsSo002!CustStatusName = GetRsValue("select Description FROM " & GetOwner & "CD035 WHERE CodeNo=4", gcnGi) & ""
'         .rsSo001!UpdTime = GetDTString(RightNow)
'         .rsSo001!UpdEn = garyGi(1)
'         .rsSo001.Update
'         .rsSo002.Update
'    End With
'    cn2.CommitTrans
'    cn2.Close
'    Set cn2 = Nothing
'
'    frmSO1100BMDI.lblCustStatusString.Caption = frmSO1100BMDI.rsSo002!CustStatusName & ""
'   'frmSO1100BMDI.RcdToScr
'    ServiceCommon.CommonRcdToScr frmSO1100BMDI.ChildForm
'   'frmSO1100BMDI.RcdToScr
'
'    Unload Me '(3) 最後結束本畫面,回到客戶基本資料畫面,並更新畫面內容
'  Exit Sub
'CnErr:
'    MsgBox "連線錯誤而無法註銷！錯誤原因：" & Err.Description, vbCritical, "錯誤"
'  Exit Sub
'ProcErr:
'    MsgBox "資料發生錯誤而無法註銷！錯誤原因：" & Err.Description, vbCritical, "錯誤"
'    On Error Resume Next
'    cn2.RollbackTrans
'    cn2.Close
'    Set cn2 = Nothing
'End Sub
Private Sub cmdOK_Click()
  On Error GoTo CnErr
    
    Dim cn2 As New ADODB.Connection
    
    If Not MustExist(gdaDelDate, 1, "註銷日期") Then Exit Sub
    If Not MustExist(gilDelete, 2, "註銷原因") Then Exit Sub
    
    With cn2
        .CursorLocation = adUseServer
        .Open gcnGi.ConnectionString
        .BeginTrans
    End With
  
  On Error GoTo ProcErr
    
    ' QQ : 1405 , RA Doc : SO1100B_20070316_Debby_調整註銷還原規格.doc
    ' 整段 cmdOK_Click Subroutine 重寫
    
    ' 6.  註銷還原時，請判斷若該客編其它服務別的狀態也為註銷，則一併將其還原。
    ' CustStatusCode : 1,正常    2,停機  3,拆機  4,註銷  5,促銷  6,拆機中
  
    Dim strQry As String
    Dim strInst As String
    Dim strUninst As String
    Dim SvcTp As String
    Dim blnRd As Boolean
    Dim strMdu As String
    Dim intCustStatus As Integer
    Dim strCustStatus As String
    Dim strCust As String
    Dim rs2 As New ADODB.Recordset

    strQry = "SELECT CUSTSTATUSCODE,CUSTSTATUSNAME,OLDSTATUSCODE" & _
                    ",DELCODE,DELNAME,DELDATE,SERVICETYPE" & _
                    " FROM " & GetOwner & "SO002 " & _
                    " WHERE COMPCODE=" & gCompCode & _
                    " AND CUSTID=" & gCustId & _
                    " AND CUSTSTATUSCODE <> 4"

    blnRd = False
    strMdu = RPxx(frmSO1100BMDI.rsSo001!MduId & "")
    strCust = RPxx(frmSO1100BMDI.txtCustName & "")
    strCustStatus = GetRsValue("SELECT DESCRIPTION FROM " & GetOwner & "CD035 WHERE CODENO = 4") & ""
    
    If GetRS(rs2, strQry, cn2, adUseClient, adOpenKeyset, adLockOptimistic) Then
        
        With rs2
            
            While Not .EOF
            
                SvcTp = .Fields("ServiceType") & ""
                intCustStatus = Val(.Fields("CustStatusCode") & "")
                
                '(1) 更新各種戶數統計欄位:
                If Not blnRd Then
                    If intCustStatus = 1 Then '.若該客戶狀態為正常 (1), 則:
                        '1.  將街道代碼檔(CD017)中, 該客戶裝機地址所對應的街道, 將街道安裝戶數減1
                        If ExecuteSQL("UPDATE " & GetOwner & "CD017 SET INSTCNT = INSTCNT - 1" & _
                                                " WHERE CODENO = " & frmSO1100BMDI.rsSo001!InstAddrNO, cn2) = giOK Then
                            blnRd = True
                        Else
                            cn2.RollbackTrans
                            cn2.Close
                            Set cn2 = Nothing
                            Exit Sub
                        End If
                    End If
                End If
                
                If Len(strMdu) > 0 Then '2.  若該客戶為大樓戶(該筆客戶資料的MduId欄位有值)
                    strInst = Switch(SvcTp = "C", "InstCnt", SvcTp = "I", "InstCnt1", SvcTp = "P", "InstCnt3", SvcTp = "D", "InstCnt2")
                    strUninst = Switch(SvcTp = "C", "UnInstCnt", SvcTp = "I", "UnInstCnt1", SvcTp = "P", "UnInstCnt3", SvcTp = "D", "UnInstCnt2")
                    '.若該客戶狀態為正常 (1), 則:
                    If intCustStatus = 1 Then
                        ' 則將大樓資料檔的該大樓的有效戶數減1
                        If ExecuteSQL("UPDATE " & GetOwner & "SO017" & _
                                                " SET " & strInst & " = " & strInst & " - 1" & _
                                                " WHERE MDUID = '" & strMdu & "'", cn2) < giOK Then
                             cn2.RollbackTrans
                             cn2.Close
                             Set cn2 = Nothing
                             Exit Sub
                         End If
                    ElseIf intCustStatus = 2 Or intCustStatus = 3 Then ' 若該客戶狀態為已拆機(3)或已停機(2)
                        ' 則將大樓資料檔的該大樓的大樓停拆／待拆戶數減1
                        If ExecuteSQL("UPDATE " & GetOwner & "SO017" & _
                                                " SET " & strUninst & " = " & strUninst & " - 1" & _
                                                " WHERE MDUID = '" & strMdu & "'", cn2) < giOK Then
                            cn2.RollbackTrans
                            cn2.Close
                            Set cn2 = Nothing
                            Exit Sub
                        End If
                    End If
                End If
                
                ' 客戶資料異動Log檔
                
                cn2.Execute "INSERT INTO " & GetOwner & "SO101" & _
                                    " (CUSTID,CUSTNAME,CUSTSTATUSCODE,CUSTSTATUSNAME" & _
                                    ",DELDATE,DELCODE,DELNAME,CUSTSTATUSCODEB,CUSTSTATUSNAMEB" & _
                                    ",UPDTIME,UPDEN,MODEID,COMPCODE,SERVICETYPE) VALUES " & _
                                    "(" & gCustId & ",'" & strCust & "'," & _
                                    intCustStatus & ",'" & .Fields("CustStatusName") & "'," & _
                                     GetNullString(gdaDelDate.Text, giDateV) & "," & _
                                     gilDelete.GetCodeNo2 & ",'" & gilDelete.GetDescription2 & "'," & _
                                     "4,'" & strCustStatus & "','" & _
                                     GetDTString(RightNow) & "','" & garyGi(1) & "',4," & _
                                     gCompCode & ",'" & SvcTp & "')"
                
                .Fields("DelDate") = SolveDate(gdaDelDate.GetValue) & ""
                .Fields("DelCode") = gilDelete.GetCodeNo2 & ""
                .Fields("DelName") = gilDelete.GetDescription2 & ""
                .Fields("OldStatusCode") = intCustStatus
                .Fields("CustStatusCode") = 4
                .Fields("CustStatusName") = strCustStatus
                .Update
                
                .MoveNext
                
            Wend
        End With
    End If
    
    With frmSO1100BMDI.rsSo001
         .Fields("UpdTime") = GetDTString(RightNow)
         .Fields("UpdEn") = garyGi(1)
         .Update
    End With
    
    cn2.CommitTrans
    cn2.Close
    
    Set cn2 = Nothing
    
    With frmSO1100BMDI
        .ProcessData
        ServiceCommon.CommonRcdToScr frmSO1100BMDI.ChildForm
        .lblCustStatusString = .rsSo002!CustStatusName & ""
    End With
    
    Unload Me '(3) 最後結束本畫面,回到客戶基本資料畫面,並更新畫面內容
  Exit Sub
CnErr:
    MsgBox "連線錯誤而無法註銷！錯誤原因：" & Err.Description, vbCritical, "錯誤"
  Exit Sub
ProcErr:
    MsgBox "資料發生錯誤而無法註銷！錯誤原因：" & Err.Description, vbCritical, "錯誤"
    On Error Resume Next
    cn2.RollbackTrans
    cn2.Close
    Set cn2 = Nothing
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error GoTo ChkErr
    If KeyCode = vbKeyF2 And cmdOK.Enabled Then
        If Not ChkGiList(KeyCode) Then Exit Sub
        cmdOK_Click
        KeyCode = 0
    End If
    If KeyCode = vbKeyEscape And cmdCancel.Enabled Then
        KeyCode = 0
        cmdCancel_Click
        Exit Sub
    End If
   Exit Sub
ChkErr:
    ErrSub Me.Name, "Form_KeyDown"
End Sub

Private Sub Form_Load()
  On Error GoTo ChkErr
'   FrmOnTop Me
    If Crm Then
        Me.Move frmSO1100BMDI.Left + ((frmSO1100BMDI.Width - Me.Width) / 2), frmSO1100BMDI.Top + ((frmSO1100BMDI.Height - Me.Height) / 2)
    Else
        Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    End If
    gdaDelDate.SetValue (Format(RightNow, "YYYYMMDD"))     ' 可輸入註銷日期(預設為今日)
    SetLst gilDelete, "CodeNo", "Description", 3, 12, 405, 1620, "CD012", , True
    gilDelete.Filter = " Where (ServiceType ='" & gServiceType & "' OR ServiceType IS Null)"
    With gilDelete
'       .SetParent Me
        .SetCodeNo (GetRsValue("SELECT CODENO FROM " & GetOwner & "CD012", gcnGi))
        .Query_Description
    End With
    strActFrmName = Me.Name
   Exit Sub
ChkErr:
    ErrSub Me.Name, "Form_Load"
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error GoTo ChkErr
    frmSO1100BMDI.ProcessData
    On Error Resume Next
    ReleaseCOM Me
  Exit Sub
ChkErr:
    ErrSub Me.Name, "Form_Unload"
End Sub
