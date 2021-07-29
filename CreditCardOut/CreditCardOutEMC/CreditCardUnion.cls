VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CreditCardUnion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'' �I�s���e�ݶǤJ�Ѽ�
'' frmCitiBank.PrgName   ����ҥε{���W��
'' Inipath  ��l�Ƹ��|
'' ChooseStr
''""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
Public Sub Action(blnIsTCB As Boolean, pChooseMultiAcc As String, _
                             blnIsTCBCard As Boolean, _
                             Optional lngCCType As Long = 0)
   
  On Error GoTo ChkErr
    With frmUnion
        'Set objAction = frmCitiBank
        .PrgName = "CreditCardUnion"
        .IsTCB = blnIsTCB
        .IsTCBCard = blnIsTCBCard
        .strChooseMultiAcc = pChooseMultiAcc
        .blnfrmChina = False
        .frmCreditCardType = lngCCType
        
        Select Case lngCCType
            Case 0
                .txtAuthorizationBatch.Visible = False
            Case 1
                
                .Label7.Caption = "����N��"
                .Label8.Caption = "�S���N��"
                .txtMerchantName.MaxLength = 13
                .cboMore.Enabled = False
                .chkCancel.Enabled = False
                .optmmyy.Enabled = False
                .opyymm.Enabled = False
                .cboSet.Enabled = False
                .fraOtherBank.Caption = ""
                .txtAuthorizationBatch.Visible = False
                
            Case 2
                .txtSpcNO.MaxLength = 15
                .Label1.Caption = "���v�妸"
                .txtMerchantName.MaxLength = 13
                .cboMore.Enabled = False
                .chkCancel.Enabled = False
                .optmmyy.Enabled = False
                .opyymm.Enabled = False
                .cboSet.Visible = False
                .fraOtherBank.Caption = ""
                .txtAuthorizationBatch.Left = .cboSet.Left
                .txtAuthorizationBatch.Top = .cboSet.Top
                
            Case 3
                .txtSpcNO.MaxLength = 15
                .Label7.Caption = "�S���N��"
                .Label1.Caption = "���v�妸"
                .txtMerchantName.Enabled = False
                .cboMore.Enabled = False
                .chkCancel.Enabled = False
                .optmmyy.Enabled = False
                .opyymm.Enabled = False
                .cboMore.Enabled = True
                .fraOtherBank.Caption = ""
                .cboSet.Visible = False
                .txtSpcNO.MaxLength = 15
                .txtAuthorizationBatch.Visible = True
                .txtAuthorizationBatch.Left = .cboSet.Left
                .txtAuthorizationBatch.Top = .cboSet.Top
                .txtAuthorizationBatch.MaxLength = 5
                .cboMore.Clear
                .cboMore.AddItem "A.Auth(0100)"
                .cboMore.AddItem "S.Sale(0200)"
                .cboMore.AddItem "O.Off-Line Sale(0220)"
                .cboMore.AddItem "R.Refund(0200)"
                .cboMore.ListIndex = 0      '#5310 �w�]�a�Ĥ@�� 2009/09/29 By Kin
                .txtBillMemo.Visible = True
                .lblBillMessage.Visible = True
                .Height = .Height + .txtBillMemo.Height + 200
                .Frame1.Height = .Frame1.Height + .txtBillMemo.Height + 100
                .cmdOK.Top = .Frame1.Height + 300
                .cmdCancel.Top = .Frame1.Height + 300
            Case 5  '#5494 �s�W�ŷs��� By Kin 2009/09/15
                
                .fraOtherBank.Caption = ""
                '.fraOtherBank.Enabled = False
                .txtAuthorizationBatch.Visible = True
                .txtAuthorizationBatch.text = "01"
                .txtAuthorizationBatch.MaxLength = 2
                .Label1.Caption = "�帹"
                
                .cboSet.Visible = False
                .cboMore.Enabled = False
                .txtMerchantName.Enabled = False
                .chkCancel.Enabled = False
                .optmmyy.Visible = False
                .opyymm.Visible = False
                
            '#5609 �W�[�p�X�H�Υd�妸�榡 By Kin 2010/06/03
            Case 6
                .txtAuthorizationBatch.Visible = False
                .txtBillMemo.Visible = True
                .lblBillMessage.Visible = True
                .Height = .Height + .txtBillMemo.Height + 200
                .Frame1.Height = .Frame1.Height + .txtBillMemo.Height + 100
                .cmdOK.Top = .Frame1.Height + 300
                .cmdCancel.Top = .Frame1.Height + 300
             '#7765
            Case 7
               .txtAuthorizationBatch.Visible = False
               .txtBillMemo.Visible = True
               .lblBillMessage.Visible = True
               .txtAuthorizationBatch.Visible = True
               .txtAuthorizationBatch.MaxLength = 4
               .cboSet.Visible = False
               .Label1.Caption = "���v�妸"
               .Height = .Height + .txtBillMemo.Height + 200
               .Frame1.Height = .Frame1.Height + .txtBillMemo.Height + 100
               .cmdOK.Top = .Frame1.Height + 300
               .cmdCancel.Top = .Frame1.Height + 300
        End Select
        .Show vbModal
    End With
   Exit Sub
ChkErr:
   ErrSub "clsCitiBank", "Action"
End Sub