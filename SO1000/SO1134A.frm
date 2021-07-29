VERSION 5.00
Object = "{3BB22E57-824A-11D3-BD70-0080C8F80BC4}#3.4#0"; "GiGridR.ocx"
Begin VB.Form frmSO1134A 
   BorderStyle     =   1  '單線固定
   Caption         =   "退費調改資料瀏覽 [SO1134A]"
   ClientHeight    =   5085
   ClientLeft      =   45
   ClientTop       =   1530
   ClientWidth     =   11880
   Icon            =   "SO1134A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   Begin prjGiGridR.GiGridR ggrData 
      Height          =   5100
      Left            =   15
      TabIndex        =   0
      Top             =   0
      Width           =   11850
      _ExtentX        =   20902
      _ExtentY        =   8996
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmSO1134A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Programmer: Power Hammer
' Last Modify: 2006/03/06

Option Explicit

Private rs As New ADODB.Recordset

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error Resume Next
    If KeyCode = 27 Then Unload Me
End Sub

Private Sub Form_Load()
  On Error GoTo ChkErr
    If Crm Then
        Move frmSO1100BMDI.Left + ((frmSO1100BMDI.Width - Width) / 2), frmSO1100BMDI.Top + ((frmSO1100BMDI.Height - Height) / 2)
    Else
        Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    End If
    OpenData
  Exit Sub
ChkErr:
    ErrSub Name, "Form_Load"
End Sub

Private Sub OpenData()
  On Error GoTo ChkErr
  
    'select狀態、受理日期、調改原因、退費調改收費項目名稱、實退金額、實退日期、期數、有效起始日期、有效截止日期
    '   、受理人員、受理人額外說明、欲調改之單號、欲調改之項次、欲調改之收費項目名稱、審核人員姓名1
    '   、審核人額外說明1、核准日期1、審核人員姓名2、審核人額外說明2、核准日期2、退費調改收費單號
    '   from so171 where CustId=[該客戶編號] and ServiceType=[目前服務類別]
    '   order by AcceptTime desc
    
    GetRS rs, "SELECT * FROM " & GetOwner & "SO171" & _
                    " WHERE CUSTID=" & gCustId & _
                    " AND SERVICETYPE='" & gServiceType & "'" & _
                    " ORDER BY ACCEPTTIME DESC", _
                    gcnGi, adUseClient, adOpenStatic, adLockReadOnly
    GrdFmt
  Exit Sub
ChkErr:
    ErrSub Name, "OpenData"
End Sub

Private Sub GrdFmt()
  On Error GoTo ChkErr
  
    ' 狀態、受理日期、調改原因、退費調改收費項目名稱、實退金額、實退日期、期數、有效起始日期、有效截止日期
    '、受理人員、受理人額外說明、欲調改之單號、欲調改之項次、欲調改之收費項目名稱、審核人員姓名1、審核人額外說明1
    '、核准日期1、審核人員姓名2、審核人額外說明2、核准日期2、退費調改收費單號
    Dim mFlds As New GiGridFlds
    With mFlds
            .Add "PSTATUS", , , , , "狀態", vbRightJustify
            .Add "ACCEPTTIME", giControlTypeDate, , , , "  受理時間  ", vbLeftJustify
            .Add "STNAME", , , , , "調改原因", vbLeftJustify
            .Add "CITEMNAME", , , , , "退費調改收費項目名稱", vbLeftJustify
            .Add "REALAMT", , , , , "實退金額", vbRightJustify
            .Add "REALDATE", giControlTypeDate, , , , "  實退日期  ", vbLeftJustify
            .Add "REALPERIOD", , , , , "期數", vbRightJustify
            .Add "REALSTARTDATE", giControlTypeDate, , , , "有效起始日期", vbLeftJustify
            .Add "REALSTOPDATE", giControlTypeDate, , , , "有效截止日期", vbLeftJustify
            .Add "ACCEPTNAME", , , , , "受理人員姓名", vbLeftJustify
            .Add "NOTES1", , , , , "受理人額外說明", vbLeftJustify
            .Add "SBILLNO", , , , , "欲調改之單號", vbLeftJustify
            .Add "SITEM", , , , , "欲調改之項次", vbRightJustify
            .Add "SCITEMNAME", , , , , "欲調改之收費項目名稱", vbLeftJustify
            .Add "PROCESSNAME1", , , , , "審核人員姓名1", vbLeftJustify
            .Add "NOTES2", , , , , "審核人額外說明1", vbLeftJustify
            .Add "PROCESSDATE1", giControlTypeDate, , , , " 核准日期1 ", vbLeftJustify
            .Add "PROCESSNAME2", , , , , "審核人員姓名2", vbLeftJustify
            .Add "NOTES3", , , , , "審核人額外說明2", vbLeftJustify
            .Add "PROCESSDATE2", giControlTypeDate, , , , " 核准日期2 ", vbLeftJustify
            .Add "BILLNO", , , , , "退費調改收費單號", vbLeftJustify
    End With
    With ggrData
            .AllFields = mFlds
            Set .Recordset = rs
            .Refresh
    End With
  Exit Sub
ChkErr:
    ErrSub Name, "GrdFmt"
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
    CloseRS rs
    Rlx rs
End Sub
