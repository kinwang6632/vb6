VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsACHauthShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub FormOutShow(strArray As String, cn As ADODB.Connection)
    On Error GoTo ChkErr
        Set gcnGi = cn
        Call GetGlobal(strArray)
        frmSO3293A.Show vbModal
    Exit Sub
ChkErr:
    ErrSub "", "FormShow"
End Sub
