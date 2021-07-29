VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CreditCardCitiBank"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'' 呼叫之前需傳入參數
'' frmCitiBank.PrgName   表單所用程式名稱
'' Inipath  初始化路徑
'' ChooseStr
''"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
Public Sub Action(pChooseMultiAcc As String, pintPara24 As Integer)
  On Error GoTo ChkErr
        With frmCitiBank
                'Set objAction = frmCitiBank
                .PrgName = "CreditCardCitiBank"
                .strChooseMultiAcc = pChooseMultiAcc
                 .intPara24 = pintPara24
                .Show vbModal
        End With
   Exit Sub
ChkErr:
   ErrSub "clsCitiBank", "Action"
End Sub


