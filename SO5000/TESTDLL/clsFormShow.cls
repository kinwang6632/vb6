VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFormShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Sub FormShow(ByVal strArray As String, ByRef cn As ADODB.Connection)
    On Error GoTo ChkErr
        Set gcnGi = cn
        Call GetGlobal2(strArray)
        Set cnn = GetTmpMdbCn
        gCompCode = garyGi(9)
        garyGi(15) = garyGi(9)
        frmSO5210A.Show vbModal
    Exit Sub
ChkErr:

    ErrSub "clsFormShow", "FormShow"
End Sub
Private Function GetGlobal2(strArray As String)
  On Error GoTo ChkErr
    Dim strParams() As String
    Dim varElement  As Variant
    Dim intLoop As Integer
    strParams = Split(strArray, Chr(9))
    ReDim Preserve strParams(20) As String
    For Each varElement In strParams
        garyGi(intLoop) = varElement
        If intLoop = 10 Then
            If garyGi(10) <> "" Then
                Call GetIniPath1(garyGi(10))
            End If
        ElseIf intLoop = 11 Then
            If garyGi(11) <> "" Then
                Call GetIniPath2(garyGi(11))
            End If
        End If
       intLoop = intLoop + 1
    Next
    Exit Function
ChkErr:
    Call ErrSub("clsFormShow", "GetGlobal2")
End Function
Public Sub objClose()
  On Error Resume Next
    Dim fm As Form
    For Each fm In Forms
        Unload fm
    End If
End Sub
Private Sub Class_Initialize()
    Set objNothing = Me
End Sub
