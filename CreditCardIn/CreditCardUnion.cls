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
'' 2003/10/09  �[�J blnfrmChina�P�_�Ѽ� �A�w�]�� false �A�p�G�Otrue ���ܦ����s���i�n�Ϊ�����ӻ�
Public Sub Action(blnIsTCB As Boolean, blnIsTCBCard As Boolean, _
        Optional lngCCType As Long = 0, Optional pstrCitemCodeOut As String = "", _
        Optional pstrRealDate As String = "")
  On Error GoTo ChkErr
        With frmUnion
                'Set objAction = frmCitiBank
               '' .PrgName = "CreditCardCitiBank"
               .IsTcb = blnIsTCB
               .isTcbCard = blnIsTCBCard
               .frmCreditCardType = lngCCType
               .pstrCitemCode = pstrCitemCodeOut
               .pstrRealDatefrm = pstrRealDate
              
               
'               If .frmCreditCardType = lngHSBC Then
'             ''       .gdtRealDate.Visible = True
'          ''          .lblRealDate.Visible = True
'               End If
               
                .Show vbModal
               
        End With
   Exit Sub
ChkErr:
   ErrSub "clsCitiBank", "Action"
End Sub
