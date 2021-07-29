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
'' 2003/10/09  加入 blnfrmChina判斷參數 ，預設為 false ，如果是true 表示此為新視波要用的中國商銀
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

