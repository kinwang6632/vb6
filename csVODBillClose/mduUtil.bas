Attribute VB_Name = "mduUtil"
Option Explicit

'Public Function GetFaciStatus(ByVal rs004 As ADODB.Recordset, Optional ByVal Value) As String
'    On Error Resume Next
'    Dim FaciWipStatus As String
'    Dim rs As New ADODB.Recordset
'    Dim strPRField As String
'        With rs004
'            If .State = adStateClosed Then Exit Function
'            If .EOF Or .BOF Then Exit Function
''            1.正常=>InstDate is not null and PRDate is null
''            2.停機=>InstDate is not null and PRDate is null And FaciStatusCode = 3
''            3.拆機/未取回=>PRDate is not null and GetDate is null
''            4.拆機/取回=>PRDate is not null and GetDate is not null
''            5.無:InstDate is null and PrDate is null
'            If GetSystemParaItem("FaciAuthority", gCompCode, gServiceType, "SO042") = 2 Then
'                strPRField = "GetDate"
'            Else
'                strPRField = "PRDate"
'            End If
'            If .Fields("InstDate") & "" <> "" And .Fields(strPRField) & "" = "" Then
'                If StrToDate(.Fields("StopTime") & "", True) > StrToDate(.Fields("ReInstTime") & "", True) Then
'                    FaciWipStatus = "停機"
'                Else
'                    '2009/07/23 Jacky 5141 MinChen
''                    If .Fields("FaciStatusCode") & "" = 3 Then
''                        FaciWipStatus = "停機"
''                    Else
'                        FaciWipStatus = "正常"
''                    End If
'                End If
'            ElseIf .Fields(strPRField) & "" <> "" Then
'                FaciWipStatus = "拆機"
'                If .Fields("GetDate") & "" = "" Then
'                    FaciWipStatus = FaciWipStatus & "/未取回"
'                End If
'            Else
'                FaciWipStatus = "無"
'            End If
'            If .Fields("GetDate") & "" <> "" Then
'                FaciWipStatus = FaciWipStatus & "/取回"
'            End If
'        End With
'        GetFaciStatus = FaciWipStatus
'End Function
