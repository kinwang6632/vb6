VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MediaPost4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Sub Action()
  On Error GoTo ChkErr
        With frmPost4
                .PrgName = "mediaPost4"
                .Show vbModal
        End With
   Exit Sub
ChkErr:
    ErrSub "mediaPost4", "Action"
End Sub


