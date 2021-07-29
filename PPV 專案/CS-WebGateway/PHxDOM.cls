VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PHxDOM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vFldNameS() As String
Private vRowValue() As String
Private vDataXML() As String
Private vElement As Variant
Private lngLoop As Long
Private lngLoop2 As Long
Private lngRcdCnt As Long
Private strNode1 As String
Private strNode2 As String
Private bytFldCnt As Byte
Private bytLoopCnt As Byte

Public Function RScvtXML(ByVal rsObj As Object, _
                            strRootName As String, _
                            Optional strRootProperty As String = "", _
                            Optional strParameter As String = "", _
                            Optional strNodeName As String, _
                            Optional strNodeProperty As String = "", _
                            Optional strElement As String = "", _
                            Optional strOnlyRow As Boolean = False, _
                            Optional FilterSymbol As Boolean = True) As String
  
  On Error GoTo ErrHnd
  
    Dim strTag As String
    Dim fld As Object
    Dim fn As Long
    
    Dim hFile As Long
    Dim Ret As Long
    
    bytFldCnt = rsObj.Fields.Count
    bytLoopCnt = bytFldCnt - 1
    
    ReDim Preserve vFldNameS(bytFldCnt) As String
    ReDim Preserve vRowValue(bytFldCnt * 10) As String
    ReDim Preserve vDataXML(364) As String
    
    strNode1 = "    <" & strElement
    strNode2 = "/>" & vbCrLf
    
    With rsObj
         lngLoop = 0
         
         For Each fld In .Fields
            vFldNameS(lngLoop) = " " & fld.Name & "="""
            lngLoop = lngLoop + 1
         Next
         
        .MoveFirst
        
         If Not strOnlyRow Then
            
            strTag = "<?xml version=""1.0"" encoding=""BIG5"" standalone=""yes"" ?>"
            strTag = strTag & vbCrLf & "<" & strRootName & IIf(strRootProperty = Empty, "", " " & strRootProperty) & ">"
            If strParameter <> Empty Then strTag = strTag & vbCrLf & "  " & strParameter
            strTag = strTag & vbCrLf & "  <" & strNodeName & IIf(strNodeProperty = Empty, "", " " & strNodeProperty) & ">" & vbCrLf
            
            vDataXML(0) = strTag
            lngRcdCnt = lngRcdCnt + 1
         
         End If
         
         If FilterSymbol Then
            Do Until .EOF
               vDataXML(lngRcdCnt) = PHparser2(.GetRows(10, 0))
               lngRcdCnt = lngRcdCnt + 1
            Loop
         Else
            Do Until .EOF
               vDataXML(lngRcdCnt) = PHparser(.GetRows(10, 0))
               lngRcdCnt = lngRcdCnt + 1
            Loop
         End If
         
         
         If Not strOnlyRow Then vDataXML(lngRcdCnt) = "  </" & strNodeName & ">" & vbCrLf & "</" & strRootName & ">"

    End With
    
    RScvtXML = Join(vDataXML, "")
    GbgCollection
    
  Exit Function
ErrHnd:
    ErrHandle "PHxDOM", "RScvtXML"
End Function
    
Private Function PHparser(ByVal vRow As Variant, Optional Filter As Boolean = False) As String
  On Error GoTo ErrHnd
    lngLoop2 = 0
    For Each vElement In vRow
        vRowValue(lngLoop2) = MergeX(vElement)
        lngLoop2 = lngLoop2 + 1
    Next
    ReDim Preserve vRowValue(lngLoop2 - 1) As String
    PHparser = Join(vRowValue(), "")
  Exit Function
ErrHnd:
    ErrHandle "PHxDOM", "PHparser"
End Function

Private Function PHparser2(ByVal vRow As Variant, Optional Filter As Boolean = False) As String
  On Error GoTo ErrHnd
    lngLoop2 = 0
    For Each vElement In vRow
        vRowValue(lngLoop2) = MergeX2(vElement)
        lngLoop2 = lngLoop2 + 1
    Next
    ReDim Preserve vRowValue(lngLoop2 - 1) As String
    PHparser2 = Join(vRowValue(), "")
  Exit Function
ErrHnd:
    ErrHandle "PHxDOM", "PHparser2"
End Function

Private Function RPC(strValue) As String
  On Error Resume Next
    If IsNull(strValue) Then RPC = "": Exit Function
    strValue = Replace(strValue, "&", "&amp;")
    strValue = Replace(strValue, "<", "&lt;")
    strValue = Replace(strValue, ">", "&gt;")
    strValue = Replace(strValue, """", "&quot;")
    strValue = Replace(strValue, "'", "&apos;")
    RPC = strValue
End Function

Private Function MergeX(Value As Variant) As String
  On Error GoTo ErrHnd
    Static bytCnt As Byte
    If bytCnt = bytLoopCnt Then
        MergeX = vFldNameS(bytCnt) & Value & """" & strNode2
        If bytLoopCnt = 0 Then MergeX = strNode1 & MergeX
        bytCnt = 0
    ElseIf bytCnt = 0 Then
        MergeX = strNode1 & vFldNameS(bytCnt) & Value & """"
        bytCnt = bytCnt + 1
    Else
        MergeX = vFldNameS(bytCnt) & Value & """"
        bytCnt = bytCnt + 1
    End If
  Exit Function
ErrHnd:
    ErrHandle "PHxDOM", "MergeX"
End Function

Private Function MergeX2(Value As Variant) As String
  On Error GoTo ErrHnd
    Static bytCnt As Byte
    If bytCnt = bytLoopCnt Then
        MergeX2 = vFldNameS(bytCnt) & RPC(Value) & """" & strNode2
        If bytLoopCnt = 0 Then MergeX2 = strNode1 & MergeX2
        bytCnt = 0
    ElseIf bytCnt = 0 Then
        MergeX2 = strNode1 & vFldNameS(bytCnt) & RPC(Value) & """"
        bytCnt = bytCnt + 1
    Else
        MergeX2 = vFldNameS(bytCnt) & RPC(Value) & """"
        bytCnt = bytCnt + 1
    End If
  Exit Function
ErrHnd:
    ErrHandle "PHxDOM", "MergeX2"
End Function

Private Sub GbgCollection()
  On Error Resume Next
    ReDim Preserve vFldNameS(0) As String
    ReDim Preserve vRowValue(0) As String
    ReDim Preserve vDataXML(0) As String
End Sub

Private Sub Class_Terminate()
  On Error Resume Next
    Erase vFldNameS
    Erase vRowValue
    Erase vDataXML
    Rlx vElement
    Rlx lngLoop
    Rlx lngLoop2
    Rlx lngRcdCnt
    Rlx strNode1
    Rlx strNode2
    Rlx bytFldCnt
    Rlx bytLoopCnt
End Sub
