VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3BB22E57-824A-11D3-BD70-0080C8F80BC4}#3.4#0"; "GiGridR.ocx"
Begin VB.Form frmSO1135A 
   BorderStyle     =   1  '��u�T�w
   Caption         =   "CP�q�H�O�����s�� [SO1135A]"
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
      Align           =   2  '������U��
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
         Name            =   "�s�ө���"
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

' SO1135A CP�q�H�O�����s��
' Doc : (PM Spec)_SO3273�{�d�N����Ʋ��ͽվ�W����_20060308_Lawrence.doc

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

'   SO1135A��Caption�����[��J[so033/so034.CitemName]�ȡC
'   Ex: FrmSO1135A.Caption:=' CP�q�H�O�����s�� [SO1135A] '+[so033/so034.CitemName]��
    Caption = "CP�q�H�O�����s�� [SO1135A] " & frmSO1100BMDI.rsSO033SO034!CitemName & ""
    
    OpenData
    
  Exit Sub
ChkErr:
    ErrSub Name, "Form_Load"
End Sub

Private Sub OpenData()
  On Error GoTo ChkErr
  
    'select TFNBillNo, CPShouldYM, CitemName, ShouldAmt from SO033CP
    '   where BillNo=[so033/so034.BillNo]
    '   and Item=[so033/so034.BillNo]
    '   and ServiceType=[so033/so034. ServiceType] order by CPItem
    
    With frmSO1100BMDI.rsSO033SO034
        GetRS rs, "SELECT TFNBILLNO, CPSHOULDYM, CITEMNAME, SHOULDAMT FROM " & GetOwner & "SO033CP" & _
                        " WHERE BILLNO='" & !BillNo & "' AND ITEM=" & !Item & " AND SERVICETYPE='" & !ServiceType & "'" & _
                        " ORDER BY CPITEM", gcnGi, adUseClient, adOpenStatic, adLockReadOnly
    End With
    GrdFmt
  
    'stbMessage�a�J'�p�p:'+[so033/so034.ShouldAmt]�ȡC
    stbMessage.Panels.Item(1) = "�p�p : " & Val(frmSO1100BMDI.rsSO033SO034!ShouldAmt & "")
    
  Exit Sub
ChkErr:
    ErrSub Name, "OpenData"
End Sub

Private Sub GrdFmt()
  On Error GoTo ChkErr
  
    ' ggrData���e�P���ǡGTFN�b��s��(TFNBillNo)�BCP�q�ܶO�~��(CPShouldYM)�B�O�ΦW��(CitemName)�B���B(ShouldAmt)�C
    Dim mFlds As New GiGridFlds
    With mFlds
            .Add "TFNBILLNO", , , , , "TFN�b��s��  ", vbLeftJustify
            .Add "CPSHOULDYM", , , , , "CP�q�ܶO�~��  ", vbLeftJustify
            .Add "CITEMNAME", , , , , "���O���ئW��                                    ", vbLeftJustify
            .Add "SHOULDAMT", , , , , "�X����B  ", vbRightJustify
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
