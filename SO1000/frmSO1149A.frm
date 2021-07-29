VERSION 5.00
Object = "{5C8CED40-8909-11D0-9483-00A0C91110ED}#1.0#0"; "MSDATREP.OCX"
Begin VB.Form frmSO1149A 
   BorderStyle     =   3  '雙線固定對話方塊
   Caption         =   "客戶資料異動記錄 [SO1149A]"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11880
   Icon            =   "frmSO1149A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   Begin MSDataRepeaterLib.DataRepeater dtR101 
      Bindings        =   "frmSO1149A.frx":0442
      Height          =   7140
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   12594
      _StreamID       =   -1412567295
      _Version        =   393216
      ForeColor       =   8388608
      ScrollBars      =   4
      RowDividerStyle =   3
      CaptionStyle    =   0
      Caption         =   $"frmSO1149A.frx":0458
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty RepeatedControlName {21FC0FC0-1E5C-11D1-A327-00AA00688B10} 
         _StreamID       =   -1412567295
         _Version        =   65536
         Name            =   "cs101.ctl101"
      EndProperty
   End
End
Attribute VB_Name = "frmSO1149A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsSO101 As New ADOR.Recordset

Private Sub Form_Load()
  On Error GoTo ChkErr
    If Crm Then
        PolyFrmFunction Me
        Me.Move frmSO1100BMDI.Left + ((frmSO1100BMDI.Width - Me.Width) / 2), frmSO1100BMDI.Top + ((frmSO1100BMDI.Height - Me.Height) / 2)
    Else
        If Not800600 Then Me.Move (Screen.Width - Me.Width) / 2, ((Screen.Height - Me.Height) / 2) + 450
    End If
    Call GetRS(rsSO101, "SELECT CUSTID,CUSTNAME,TEL1,TEL2,TEL3," & _
                " INSTADDRNO,INSTADDRESS,UPDTIME,UPDEN,MODEID,CUSTNAMEB,TEL1B," & _
                " TEL2B,TEL3B,INSTADDRNOB,INSTADDRESSB,MAILADDRNO,MAILADDRESS,MAILADDRNOB," & _
                " MAILADDRESSB,CLASSNAME1,CLASSNAME1B,CHARGEADDRNO,CHARGEADDRNOB, " & _
                " CHARGEADDRESS,CHARGEADDRESSB " & _
                " FROM " & GetOwner & "SO101 WHERE CUSTID=" & gCustId & " AND COMPCODE=" & gCompCode & " ORDER BY UPDTIME DESC", gcnGi, adUseClient, adOpenForwardOnly, adLockReadOnly)
'    Call GetRS(rsSO101, "SELECT CUSTID,CUSTNAME,TEL1,TEL2,TEL3,INSTADDRNO,INSTADDRESS,UPDTIME,UPDEN,MODEID,CUSTNAMEB,TEL1B,TEL2B,TEL3B,INSTADDRNOB,INSTADDRESSB,CLASSNAME1,CLASSNAME1B FROM " & GetOwner & "SO101 WHERE CUSTID=" & gCustId & " AND COMPCODE=" & gCompCode & " AND SERVICETYPE='" & gServiceType & "' ORDER BY UPDTIME DESC", gcnGi, adUseClient, adOpenForwardOnly, adLockReadOnly)
    With dtR101
'       .RepeatedControlName = "cs101.ctl101"
        .RepeaterBindings.Add "FldName", "CUSTNAME"
        .RepeaterBindings.Add "FldNameB", "CUSTNAMEB"
        .RepeaterBindings.Add "FldTEL1", "TEL1"
        .RepeaterBindings.Add "FldTEL1B", "TEL1B"
        .RepeaterBindings.Add "FldTEL2", "TEL2"
        .RepeaterBindings.Add "FldTEL2B", "TEL2B"
        .RepeaterBindings.Add "FldTEL3", "TEL3"
        .RepeaterBindings.Add "FldTEL3B", "TEL3B"
        .RepeaterBindings.Add "FldAddress", "INSTADDRESS"
        .RepeaterBindings.Add "FldAddressB", "INSTADDRESSB"
        .RepeaterBindings.Add "FldMode", "MODEID"
        .RepeaterBindings.Add "FldAddrNo", "INSTADDRNO"
        .RepeaterBindings.Add "FldAddrNoB", "INSTADDRNOB"
        .RepeaterBindings.Add "FldMailAddrNo", "MAILADDRNO"
        .RepeaterBindings.Add "FldMailAddrNoB", "MAILADDRNOB"
        .RepeaterBindings.Add "FldMailAddress", "MAILADDRESS"
        .RepeaterBindings.Add "FldMailAddressB", "MAILADDRESSB"
        '#5937 增加收費地址 By Kin 2011/05/06
        .RepeaterBindings.Add "FldChargeAddrNo", "CHARGEADDRNO"
        .RepeaterBindings.Add "FldChargeAddrNoB", "CHARGEADDRNOB"
        .RepeaterBindings.Add "FldChargeAddress", "CHARGEADDRESS"
        .RepeaterBindings.Add "FldChargeAddressB", "CHARGEADDRESSB"
        .RepeaterBindings.Add "FldClass", "CLASSNAME1"
        .RepeaterBindings.Add "FldClassB", "CLASSNAME1B"
        .RepeaterBindings.Add "FldEn", "UPDEN"
        .RepeaterBindings.Add "FldTime", "UPDTIME"
         Set .DataSource = rsSO101
    End With
    With rsSO101
         If .EOF Or .BOF Or .RecordCount <= 0 Then
            MsgBox "此客戶並無異動過 !!", vbInformation, "訊息"
         End If
    End With
  Exit Sub
ChkErr:
    ErrSub Me.Name, "Form_Load"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error Resume Next
    If KeyCode = 27 Then Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
    frmSO1100BMDI.Pic1.Enabled = True
    rsSO101.Close
    Set rsSO101 = Nothing
    ReleaseCOM Me
End Sub
