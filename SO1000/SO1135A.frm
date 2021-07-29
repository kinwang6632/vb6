VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3BB22E57-824A-11D3-BD70-0080C8F80BC4}#3.4#0"; "GiGridR.ocx"
Begin VB.Form frmSO1135A 
   BorderStyle     =   1  '單線固定
   Caption         =   "CP通信費明細瀏覽 [SO1135A]"
   ClientHeight    =   5085
   ClientLeft      =   45
   ClientTop       =   1530
   ClientWidth     =   11880
   Icon            =   "SO1135A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.StatusBar stbMessage 
      Align           =   2  '對齊表單下方
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   4710
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   21167
            MinWidth        =   21167
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin prjGiGridR.GiGridR ggrData 
      Height          =   4710
      Left            =   -15
      TabIndex        =   0
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   8308
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmSO1135A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Programmer: Power Hammer
' Last Modify: 2006/03/23

' SO1135A CP通信費明細瀏覽
' Doc : (PM Spec)_SO3273臨櫃代收資料產生調整規格文件_20060308_Lawrence.doc

Option Explicit

Private rs As New ADODB.Recordset
Private strCaption As String


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error Resume Next
    If KeyCode = 27 Then Unload Me
End Sub

Private Sub Form_Load()
  On Error GoTo ChkErr
    If Crm Then
        Move frmSO1100BMDI.Left + ((frmSO1100BMDI.Width - Width) / 2), frmSO1100BMDI.Top + ((frmSO1100BMDI.Height - Height) / 2)
    Else
        Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2 + 1800
    End If

'   SO1135A之Caption中須加串入[so033/so034.CitemName]值。
'   Ex: FrmSO1135A.Caption:=' CP通信費明細瀏覽 [SO1135A] '+[so033/so034.CitemName]值
    Select Case strCaption
        Case "CP"
            Caption = "CP通信費明細瀏覽 [SO1135A] " & frmSO1100BMDI.rsSO033SO034!CitemName & ""
        Case "VOD"
            Caption = "VOD 訂購明細瀏覽 [SO1135A] " & frmSO1100BMDI.rsSO033SO034!CitemName & ""
    End Select
    OpenData
    
  Exit Sub
ChkErr:
    ErrSub Name, "Form_Load"
End Sub

Private Sub OpenData()
  On Error GoTo ChkErr
  Dim strSQL As String
    With frmSO1100BMDI.rsSO033SO034
        Select Case strCaption
            Case "CP"
                strSQL = "SELECT TFNBILLNO, CPSHOULDYM, CITEMNAME, SHOULDAMT FROM " & GetOwner & "SO033CP" & _
                         " WHERE BILLNO='" & !BillNo & "' AND ITEM=" & !Item & " AND SERVICETYPE='" & !ServiceType & "'" & _
                         " ORDER BY CPITEM"
            Case "VOD"
'#5482 2010.01.26 by Corey 和SQ(Karen)討論後 維換工單完工後會將SO004.VODAccountID清除，會造成資料對應不到，所以將 SO033VOD.VODAccountID=SO004.VodAccountID 移除
'                strSQL = "Select SO004.FaciSno,SO033Vod.* From " & GetOwner & "SO033VOD," & GetOwner & "SO004 " & _
'                         " Where SO033VOD.CloseBillno ='" & .Fields("BillNo") & "' and SO033VOD.CloseItem=" & .Fields("Item") & _
'                         " and SO033VOD.Custid=" & gCustId & " and SO033VOD.VODAccountID=SO004.VodAccountID and SO033VOD.Custid=SO004.Custid " & _
'                         " and SO004.SEQNO='" & .Fields("FaciSEQNO") & "'"
                strSQL = "Select SO004.FaciSno,SO033Vod.* From " & GetOwner & "SO033VOD," & GetOwner & "SO004 " & _
                         " Where SO033VOD.CloseBillno ='" & .Fields("BillNo") & "' and SO033VOD.CloseItem=" & .Fields("Item") & _
                         " and SO033VOD.Custid=" & gCustId & " and SO033VOD.Custid=SO004.Custid " & _
                         " and SO004.SEQNO='" & .Fields("FaciSEQNO") & "'"
                         
        End Select
    End With
    
    GetRS rs, strSQL, gcnGi, adUseClient, adOpenStatic, adLockReadOnly
    GrdFmt
  
    stbMessage.Panels.Item(1) = "小計 : " & Val(frmSO1100BMDI.rsSO033SO034!ShouldAmt & "")
    
  Exit Sub
ChkErr:
    ErrSub Name, "OpenData"
End Sub

Private Sub GrdFmt()
  On Error GoTo ChkErr
  
    ' ggrData內容與順序：TFN帳單編號(TFNBillNo)、CP通話費年月(CPShouldYM)、費用名稱(CitemName)、金額(ShouldAmt)。
    Dim mFlds As New GiGridFlds
    
    With mFlds
        Select Case strCaption
            Case "CP"
                .Add "TFNBILLNO", , , , , "TFN帳單編號  ", vbLeftJustify
                .Add "CPSHOULDYM", , , , , "CP通話費年月  ", vbLeftJustify
                .Add "CITEMNAME", , , , , "收費項目名稱                                    ", vbLeftJustify
                .Add "SHOULDAMT", , , , , "出單金額  ", vbRightJustify
            Case "VOD"
                .Add "FaciSno", , , , , "設備序號         ", vbLeftJustify
                .Add "InCurredDate", , , , , "訂購日                          ", vbLeftJustify
                .Add "ItemName", , , , , "影片名稱                    ", vbLeftJustify
                .Add "Value", , , , , "消費點數       ", vbLeftJustify
        End Select
    End With
    
    With ggrData
        .AllFields = mFlds
        .SetHead
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

Public Property Let uCaption(ByVal vData As String)
    On Error Resume Next
        strCaption = vData
End Property
