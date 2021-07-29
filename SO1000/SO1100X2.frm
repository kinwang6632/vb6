VERSION 5.00
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Begin VB.Form frmSO1100X2 
   BorderStyle     =   1  '單線固定
   Caption         =   "bb點數轉換功能 [SO1100X2]"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5250
   Icon            =   "SO1100X2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   5250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所屬視窗中央
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&X)"
      Height          =   345
      Left            =   3780
      TabIndex        =   3
      Top             =   4050
      Width           =   810
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "F2.確定"
      Height          =   345
      Left            =   390
      Picture         =   "SO1100X2.frx":0442
      TabIndex        =   2
      Top             =   4050
      Width           =   780
   End
   Begin VB.Frame fraEdit 
      Height          =   3855
      Left            =   60
      TabIndex        =   4
      Top             =   90
      Width           =   5130
      Begin VB.TextBox txtTotalSavePoint 
         Alignment       =   1  '靠右對齊
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   4050
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   25
         Text            =   "0"
         Top             =   2400
         Width           =   870
      End
      Begin VB.TextBox txtTotalBonus 
         Alignment       =   1  '靠右對齊
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   4050
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   24
         Text            =   "0"
         Top             =   2670
         Width           =   870
      End
      Begin VB.TextBox txtMallbonus 
         Alignment       =   1  '靠右對齊
         Height          =   270
         Left            =   1830
         MaxLength       =   8
         TabIndex        =   16
         Text            =   "0"
         Top             =   2700
         Width           =   870
      End
      Begin VB.TextBox txtMallpoint 
         Alignment       =   1  '靠右對齊
         Height          =   270
         Left            =   1830
         MaxLength       =   8
         TabIndex        =   14
         Text            =   "0"
         Top             =   2400
         Width           =   870
      End
      Begin VB.TextBox txtRate 
         Alignment       =   1  '靠右對齊
         Height          =   270
         Left            =   1830
         MaxLength       =   8
         TabIndex        =   12
         Text            =   "0"
         Top             =   2100
         Width           =   870
      End
      Begin VB.TextBox txtUsedSavePoint 
         Alignment       =   1  '靠右對齊
         Height          =   270
         Left            =   1830
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   10
         Text            =   "0"
         Top             =   990
         Width           =   870
      End
      Begin VB.TextBox txtUsedBonus 
         Alignment       =   1  '靠右對齊
         Height          =   270
         Left            =   1830
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   8
         Text            =   "0"
         Top             =   1290
         Width           =   870
      End
      Begin VB.TextBox txtChangePoint 
         Alignment       =   1  '靠右對齊
         Height          =   270
         Left            =   1830
         MaxLength       =   8
         TabIndex        =   1
         Text            =   "0"
         Top             =   690
         Width           =   870
      End
      Begin prjGiList.GiList gilChangeStore 
         Height          =   315
         Left            =   1830
         TabIndex        =   0
         Top             =   330
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FldWidth1       =   500
         F2Corresponding =   -1  'True
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "剩餘紅利點數:"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   2880
         TabIndex        =   23
         Top             =   2730
         Width           =   1125
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "剩餘儲值點數:"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   2880
         TabIndex        =   22
         Top             =   2430
         Width           =   1125
      End
      Begin VB.Label lblbbTranspriority 
         AutoSize        =   -1  'True
         Caption         =   "轉換優先順序:"
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   1680
         TabIndex        =   21
         Top             =   3210
         Width           =   1125
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "轉換優先順序:"
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   450
         TabIndex        =   20
         Top             =   3210
         Width           =   1125
      End
      Begin VB.Label lblCondition 
         AutoSize        =   -1  'True
         Caption         =   "111"
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   1680
         TabIndex        =   19
         Top             =   3510
         Width           =   270
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "轉換條件:"
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   810
         TabIndex        =   18
         Top             =   3510
         Width           =   765
      End
      Begin VB.Line Line2 
         X1              =   30
         X2              =   5070
         Y1              =   3030
         Y2              =   3030
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "商城紅利:"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   930
         TabIndex        =   17
         Top             =   2730
         Width           =   765
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "商城點數:"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   930
         TabIndex        =   15
         Top             =   2430
         Width           =   765
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "轉換結果"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   1920
         TabIndex        =   13
         Top             =   1830
         Width           =   720
      End
      Begin VB.Line Line1 
         X1              =   30
         X2              =   5070
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "轉換匯率:"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   930
         TabIndex        =   11
         Top             =   2130
         Width           =   765
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "扣抵bb紅利點數:"
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   420
         TabIndex        =   9
         Top             =   1350
         Width           =   1305
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "扣抵bb儲值點數:"
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   420
         TabIndex        =   7
         Top             =   1050
         Width           =   1305
      End
      Begin VB.Label lblbonus 
         AutoSize        =   -1  'True
         Caption         =   "轉換點數:"
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   960
         TabIndex        =   6
         Top             =   720
         Width           =   765
      End
      Begin VB.Label lblBonusKind 
         AutoSize        =   -1  'True
         Caption         =   "轉換對象:"
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   960
         TabIndex        =   5
         Top             =   390
         Width           =   765
      End
   End
End
Attribute VB_Name = "frmSO1100X2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private fBBVodAccountId As String
Private rsSO193 As Recordset
Private rsSO004 As Recordset
Private totalSavepoint As Double
Private totalBonus As Double
'1.有到期日紅利2.無到期日紅利3.儲值點數
Private arybbTranspriority(2) As Byte
Private dicMinusPoint As New Dictionary
Private rsCalculate As ADODB.Recordset
Private ChangeRate As Double
Private ChangeCondition As Byte
Private ChangeRefNo As Integer
Private ChangeCitemCode As String
Private Transseqno As String
Private DefaultUCCode As String
Private DefaultUCName As String
Private IsSave As Boolean
Enum MinusType
    HaveStopDate = 1
    NoStopDate = 2
    SavePoint = 3
End Enum
Private Sub DefineRs()
  On Error GoTo ChkErr
    If Not rsCalculate Is Nothing Then
        CloseRecordset rsCalculate
        Set rsCalculate = Nothing
    End If
    Set rsCalculate = New ADODB.Recordset
    With rsCalculate
        .Fields.Append "RowId", adBSTR, 50
        .Fields.Append "MinusType", adBSTR, 50
        .Fields.Append "CitemCode", adBSTR, 20
        .Fields.Append "MinusPoint", adBSTR, 20
        .Fields.Append "SeqNo", adBSTR, 30
        .CursorLocation = adUseClient
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open
    End With
'    InsertChargeField
    Exit Sub
ChkErr:
  ErrSub Me.Name, "DefineRs"
End Sub
Private Sub subGil()
    On Error GoTo ChkErr
        SetgiList gilChangeStore, "CodeNo", "Description", "CD129", , , 800, , , , " Where Property = 1 And Nvl(StopFlag,0) = 0 ", True
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "subGil")
End Sub
Private Sub DefaultValue()
    On Error GoTo ChkErr
        Dim i As Integer
        Dim rsCD013 As New Recordset
        For i = LBound(arybbTranspriority) To UBound(arybbTranspriority)
            arybbTranspriority(i) = 0
        Next
        Dim strbbTranspriority As String
        strbbTranspriority = GetRsValue("Select bbTranspriority From " & GetOwner & "SO041 ", gcnGi) & ""
        strbbTranspriority = Replace(strbbTranspriority, ",", "")
        For i = 0 To Len(strbbTranspriority) - 1
            If i > 2 Then Exit For
            arybbTranspriority(i) = Mid(strbbTranspriority, i + 1, 1)
        Next
        If Not GetRS(rsCD013, "Select CodeNo,Description From " & GetOwner & "CD013 Where RefNo = 3", _
                        gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
        If Not rsCD013.EOF Then
            DefaultUCCode = rsCD013("CodeNo") & ""
            DefaultUCName = rsCD013("Description") & ""
        End If
        For i = LBound(arybbTranspriority) To UBound(arybbTranspriority)
            Select Case arybbTranspriority(i)
                Case 1
                    lblbbTranspriority.Caption = IIf(i = 0, "1.有到期日紅利", lblbbTranspriority.Caption & " " & i + 1 & "有到期日紅利")
                Case 2
                    lblbbTranspriority.Caption = IIf(i = 0, "1.無到期日紅利", lblbbTranspriority.Caption & " " & i + 1 & "無到期日紅利")
                Case 3
                    lblbbTranspriority.Caption = IIf(i = 0, "1.儲值點數", lblbbTranspriority.Caption & " " & i + 1 & "儲值點數")
            End Select
        Next
    Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "DefaultValue")
End Sub
Private Sub AddCalculateRecord(ByVal RowId As String, ByVal MinusType As String, _
                                            ByVal CitemCode As String, ByVal MinusValue As String, ByVal SeqNo As String)
  On Error GoTo ChkErr
    With rsCalculate
        .AddNew
        .Fields("RowId") = RowId
        .Fields("MinusType") = MinusType
        .Fields("CitemCode") = CitemCode
        .Fields("MinusPoint") = MinusValue
        .Fields("SeqNo") = SeqNo
        .Update
    End With
       
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "AddCalculateRecord")
End Sub
                                            
Private Function MinusPoint(ByVal totalChangePoint As Double, ByVal MinusType As MinusType) As Double
  On Error GoTo ChkErr
        Select Case MinusType
            Case HaveStopDate
                rsSO193.MoveFirst
                Do While Not rsSO193.EOF
                    If (Len(rsSO193("bonusStopDate") & "") > 0) And (Len(rsSO193("closeBillno") & "") = 0) Then
                        If CDate(rsSO193("bonusStopDate")) >= RightNow Then
                            If Val(rsSO193("bonus") & "") - Val(rsSO193("Usedbonus") & "") > 0 Then
                                If Val(rsSO193("bonus") & "") - Val(rsSO193("Usedbonus") & "") >= totalChangePoint Then
                                    Call AddCalculateRecord(rsSO193("RowId"), "2", rsSO193("CitemCode") & "", totalChangePoint, rsSO193("SeqNo"))
                                    totalChangePoint = 0
                                    Exit Do
                                Else
                                    Call AddCalculateRecord(rsSO193("RowId"), "2", rsSO193("CitemCode") & "", _
                                                       str(Val(rsSO193("bonus") & "") - Val(rsSO193("Usedbonus") & "")), rsSO193("SeqNo"))
                                    totalChangePoint = totalChangePoint - (Val(rsSO193("bonus") & "") - Val(rsSO193("Usedbonus") & ""))
                                End If
                            End If
                        End If
                    End If
                    rsSO193.MoveNext
                Loop
                
            Case NoStopDate
                rsSO193.MoveFirst
                Do While Not rsSO193.EOF
                    If (Len(rsSO193("bonusStopDate") & "") = 0) And (Len(rsSO193("closeBillno") & "") = 0) Then
                        If Val(rsSO193("bonus") & "") - Val(rsSO193("Usedbonus") & "") > 0 Then
                            If Val(rsSO193("bonus") & "") - Val(rsSO193("Usedbonus") & "") >= totalChangePoint Then
                                Call AddCalculateRecord(rsSO193("RowId"), _
                                                            "2", rsSO193("CitemCode") & "", totalChangePoint, rsSO193("SeqNo"))
                                totalChangePoint = 0
                                Exit Do
                            Else
                               Call AddCalculateRecord(rsSO193("RowId"), _
                                                            "2", rsSO193("CitemCode") & "", _
                                                             str(Val(rsSO193("bonus") & "") - Val(rsSO193("Usedbonus") & "")), rsSO193("SeqNo"))
                                totalChangePoint = totalChangePoint - (Val(rsSO193("bonus") & "") - Val(rsSO193("Usedbonus") & ""))
                            End If
                        End If
                    End If
                    rsSO193.MoveNext
                Loop
            Case Else
                 rsSO193.MoveFirst
                Do While Not rsSO193.EOF
                    If Len(rsSO193("closeBillno") & "") = 0 Then
                        If Val(rsSO193("Savepoint") & "") - Val(rsSO193("UsedSavepoint") & "") > 0 Then
                            If Val(rsSO193("Savepoint") & "") - Val(rsSO193("UsedSavepoint") & "") >= totalChangePoint Then
                                Call AddCalculateRecord(rsSO193("RowId"), _
                                                "1", rsSO193("CitemCode") & "", totalChangePoint, rsSO193("SeqNo"))
                                totalChangePoint = 0
                                Exit Do
                            Else
                                Call AddCalculateRecord(rsSO193("RowId"), _
                                                            "1", rsSO193("CitemCode") & "", _
                                                            str(Val(rsSO193("Savepoint") & "") - Val(rsSO193("UsedSavepoint") & "")), _
                                                            rsSO193("SeqNo"))
                                totalChangePoint = totalChangePoint - (Val(rsSO193("Savepoint") & "") - Val(rsSO193("UsedSavepoint") & ""))
                            End If
                        End If
                    End If
                    rsSO193.MoveNext
                Loop
        End Select
        MinusPoint = totalChangePoint
        
  Exit Function
ChkErr:
    Call ErrSub(Me.Name, "MinusPoint")
End Function
Private Function IsCanChange() As Boolean
  On Error GoTo ChkErr
    Dim totalChangePoint As Double
    Dim i As Integer
    totalChangePoint = Val(txtChangePoint.Text & "")
    Call DefineRs
    If rsSO193.RecordCount = 0 Then IsCanChange = False: Exit Function
    '判斷扣點順序 1.有到期日紅利2.無到期日紅利3.儲值點數
    For i = LBound(arybbTranspriority) To UBound(arybbTranspriority)
        If totalChangePoint <= 0 Then Exit For
        Select Case arybbTranspriority(i)
            Case 0
            Case 1
               totalChangePoint = MinusPoint(totalChangePoint, HaveStopDate)
            Case 2
               totalChangePoint = MinusPoint(totalChangePoint, NoStopDate)
            Case 3
                totalChangePoint = MinusPoint(totalChangePoint, SavePoint)
        End Select
    Next i
    IsCanChange = (totalChangePoint <= 0)
    Exit Function
ChkErr:
  Call ErrSub(Me.Name, "IsCanChange")
End Function
Private Function InsSO033bbTrans() As Boolean
  On Error GoTo ChkErr
    Dim rsSO033bbTrans As New ADODB.Recordset
    Dim transSavepoint As Double
    Dim transbonus As Double
    transbonus = 0
    transSavepoint = 0
    rsSO193.MoveFirst
    rsCalculate.MoveFirst
    Do While Not rsCalculate.EOF
        Select Case Val(rsCalculate("MinusType"))
            Case 2
                transbonus = transbonus + Val(rsCalculate("MinusPoint"))
            Case 1
                transSavepoint = transSavepoint + Val(rsCalculate("MinusPoint"))
        End Select
        rsCalculate.MoveNext
    Loop
    
    If Not GetRS(rsSO033bbTrans, "Select * From " & GetOwner & "SO033bbTrans Where 1 = 0 ", _
                                    gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
    Transseqno = GetTransseqno
    With rsSO033bbTrans
        .AddNew
        .Fields("bbAccountID") = rsSO193("bbAccountID")
        .Fields("SeqNo") = GetTransseqno
        .Fields("transSavepoint") = transSavepoint
        .Fields("transbonus") = transbonus
        .Fields("transdate") = RightNow
        .Fields("transStoreCode") = gilChangeStore.GetCodeNo
        .Fields("Rate") = ChangeRate
        .Fields("Mallpoint") = RoundCS(transSavepoint * ChangeRate)
        .Fields("Mallbonus") = RoundCS(transbonus * ChangeRate)
        .Fields("UpdTime") = RightNow
        .Fields("UpdEn") = garyGi(1)
        .Update
    End With
    InsSO033bbTrans = True
    On Error Resume Next
     Call CloseRecordset(rsSO033bbTrans)
  Exit Function
ChkErr:
  Call ErrSub(Me.Name, "InsSO033bbTrans")
End Function
Private Function InsSO139B() As Boolean
  On Error GoTo ChkErr
    Dim rs193B As New ADODB.Recordset
    If Not GetRS(rs193B, "Select * From " & GetOwner & "SO193B Where 1 = 0 ", gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
    rsCalculate.MoveFirst
    Do While Not rsCalculate.EOF
        With rs193B
            .AddNew
            .Fields("Saveseqno") = rsCalculate("SeqNo")
            .Fields("Transseqno") = Transseqno
            .Update
        End With
        rsCalculate.MoveNext
    Loop
    InsSO139B = True
    On Error Resume Next
    Call CloseRecordset(rs193B)
  Exit Function
ChkErr:
  Call ErrSub(Me.Name, "InsSO139B")
End Function
'產生負項收費資料
Private Function InsNegativeSO033() As Boolean
  On Error GoTo ChkErr
    Dim dicCitem As New Dictionary
    dicCitem.CompareMode = TextCompare
    Dim CitemCode As String
    Dim IsData As Boolean
    Dim aBillNo As String
    Dim aServiceType As String
    Dim aMinusPoint As Double
    Dim updSql As String
    Dim i As Integer
    Dim rs033 As New Recordset
    IsData = False
    dicCitem.RemoveAll
    rsCalculate.MoveFirst
    aMinusPoint = 0
    Do While Not rsCalculate.EOF
        CitemCode = Empty
        If Val(rsCalculate("MinusType") & "") = 1 Then
            If Len(rsCalculate("CitemCode") & "") > 0 Then
                CitemCode = GetRsValue("Select ReturnCode  From " & GetOwner & "CD019 " & _
                                            " Where CodeNo = " & rsCalculate("CitemCode") & _
                                            " And ReturnCode Is Not Null ", gcnGi) & ""
                If Len(CitemCode) > 0 Then
                    IsData = True
                    aMinusPoint = rsCalculate("MinusPoint")
                    If dicCitem.Exists(CitemCode) Then
                        dicCitem(CitemCode) = aMinusPoint + dicCitem(CitemCode)
                    Else
                        dicCitem.Add CitemCode, aMinusPoint
                    End If
                End If
            End If
        End If
        rsCalculate.MoveNext
    Loop
    
    If Not IsData Then InsNegativeSO033 = True: GoTo 88
    For i = LBound(dicCitem.Keys) To UBound(dicCitem.Keys)
        aServiceType = GetRsValue("Select ServiceType From " & GetOwner & "CD019 " & _
                                " Where CodeNo = " & dicCitem.Keys(i), gcnGi) & ""
        aBillNo = GetInvoiceNo("T", aServiceType)
        Call InsertChargeField(rs033, gCustId, gCompCode, aServiceType, aBillNo, 1, dicCitem.Keys(i), RightDate _
                       , , , rsSO004("SeqNo") & "", rsSO004("FaciSNo") & "")
        rs033.Update
        rs033("UCCode") = DefaultUCCode
        rs033("UCName") = DefaultUCName
        rs033("OldAmt") = 0 - dicCitem.Item(dicCitem.Keys(i))
        rs033("ShouldAmt") = 0 - dicCitem.Item(dicCitem.Keys(i))
        rs033("RealAmt") = 0 - dicCitem.Item(dicCitem.Keys(i))
        rs033("RealDate") = rs033("ShouldDate")
        rs033("UpdTime").Value = GetDTString(RightNow)
        rs033("CreateEn").Value = garyGi(0)
        rs033("UpdEn").Value = garyGi(1)
        rs033.Update
        If Not InsertToOracle("SO033", rs033) Then GoTo 88
        Call CalcVODPoint(rsSO004("VodAccountId"), "-", rs033)
        updSql = "Update " & GetOwner & "SO033BBTRANS " & _
                    " Set Reversalbillno ='" & aBillNo & "' " & _
                    " Where SeqNo = " & Transseqno
        gcnGi.Execute updSql
    Next i
    InsNegativeSO033 = True
88:
    On Error Resume Next
    Call CloseRecordset(rs033)
  Exit Function
ChkErr:
  Call ErrSub(Me.Name, "InsNegativeSO033")
End Function

'產生正項收費資料
Private Function InsPositiveSO033() As Boolean
     On Error GoTo ChkErr
     Dim dicCitem As New Dictionary
     Dim aBillNo As String
     Dim aServiceType As String
     Dim aShouldAmt As String
     Dim IsData As Boolean
     Dim rs033 As ADODB.Recordset
     Dim updSql As String
     Set rs033 = Nothing
     IsData = False
     rsCalculate.MoveFirst
     Do While Not rsCalculate.EOF
       If Val(rsCalculate("MinusType") & "") = 1 Then
           aShouldAmt = Val(aShouldAmt & "") + Val(rsCalculate("MinusPoint") & "")
           IsData = True
       End If
       rsCalculate.MoveNext
     Loop
     If Not IsData Then InsPositiveSO033 = True: GoTo 88
     If Len(ChangeCitemCode & "") = 0 Then InsPositiveSO033 = True: GoTo 88
     aServiceType = GetRsValue("Select ServiceType From " & _
                                           GetOwner & "CD019 Where CodeNo = " & ChangeCitemCode, _
                                           gcnGi)
     aBillNo = GetInvoiceNo("T", aServiceType)
    Call InsertChargeField(rs033, gCustId, gCompCode, aServiceType, aBillNo, 1, ChangeCitemCode, RightDate _
                       , , , rsSO004("SeqNo") & "", rsSO004("FaciSNo") & "")
    rs033.Update
    rs033("UCCode") = DefaultUCCode
    rs033("UCName") = DefaultUCName
    rs033("OldAmt") = aShouldAmt
    rs033("ShouldAmt") = aShouldAmt
    rs033("RealAmt") = aShouldAmt
    rs033("RealDate") = rs033("ShouldDate")
    rs033("UpdTime").Value = GetDTString(RightNow)
    rs033("CreateEn").Value = garyGi(0)
    rs033("UpdEn").Value = garyGi(1)
    rs033.Update
    If Not InsertToOracle("SO033", rs033) Then GoTo 88
    Call CalcVODPoint(rsSO004("VodAccountId"), "+", rs033)
    updSql = "Update " & GetOwner & "SO033BBTRANS " & _
                " Set Mallbillno = '" & aBillNo & "' " & _
                " Where SeqNo = " & Transseqno
    gcnGi.Execute updSql
    
    InsPositiveSO033 = True
88:
    On Error Resume Next
    Call CloseRecordset(rs033)
  Exit Function
ChkErr:
  Call ErrSub(Me.Name, "InsSO033")
End Function
'重算SO193與SO004J金額
Private Function UpdData() As Boolean
  On Error GoTo ChkErr
    Dim updSql As String
    Dim aTotalBonus As Double
    Dim aTotalSavePoint As Double
    aTotalBonus = 0
    aTotalSavePoint = 0
    rsCalculate.MoveFirst
    Do While Not rsCalculate.EOF
        Select Case Val(rsCalculate("MinusType"))
            Case 1
                updSql = "Update " & GetOwner & "SO193 " & _
                        " Set UsedSavepoint = Nvl(UsedSavepoint,0) + " & rsCalculate("MinusPoint") & _
                        " Where RowId = '" & rsCalculate("RowId") & "'"
                aTotalSavePoint = aTotalSavePoint + Val(rsCalculate("MinusPoint") & "")
            Case 2
                updSql = "Update " & GetOwner & "SO193 " & _
                        " Set Usedbonus = Nvl(Usedbonus,0) + " & rsCalculate("MinusPoint") & _
                        " Where RowId = '" & rsCalculate("RowId") & "'"
                aTotalBonus = aTotalBonus + Val(rsCalculate("MinusPoint") & "")
        End Select
        gcnGi.Execute updSql
        rsCalculate.MoveNext
    Loop
    updSql = "Update " & GetOwner & "SO004J " & _
                " Set TotalSavePoint = Nvl(TotalSavePoint,0) - " & aTotalSavePoint & _
                " ,TotalBonus = Nvl(TotalBonus,0)  - " & aTotalBonus & _
                " Where bbAccountID = '" & rsSO004("FaciId") & "'"
    gcnGi.Execute updSql
    
    UpdData = True
  Exit Function
ChkErr:
    Call ErrSub(Me.Name, "UpdData")
End Function
Private Function InsData() As Boolean
  On Error GoTo ChkErr
    If Not InsSO033bbTrans Then
        Exit Function
    End If
    If Not InsSO139B Then
        Exit Function
    End If
    If Not InsPositiveSO033 Then
        Exit Function
    End If
    If Not InsNegativeSO033 Then
        Exit Function
    End If
    InsData = True
  Exit Function
ChkErr:
  
  Call ErrSub(Me.Name, "InsData")
End Function
Private Function IsDataOk() As Boolean
  On Error GoTo ChkErr
    Dim i As Integer
    If Len(gilChangeStore.GetCodeNo & "") = 0 Then
        MsgMustBe ("轉換對象")
        gilChangeStore.SetFocus
        Exit Function
    End If
    If Len(txtChangePoint.Text & "") = 0 Or Val(txtChangePoint.Text & "") = 0 Then
        MsgBox "轉換點數不正確！", vbOKOnly, "訊息"
        txtChangePoint.SetFocus
        Exit Function
    End If
    If Val(txtChangePoint.Text & "") > totalBonus + totalSavepoint Then
        MsgBox "點數不足！", vbOKOnly + vbInformation, "訊息"
        txtChangePoint.SetFocus
        Exit Function
    End If
    Call FilterCondition
    If Not IsCanChange Then
        MsgBox "點數不足！", vbOKOnly + vbInformation, "訊息"
        Exit Function
    End If
    
'    If Len(gdaBonusStopDate.GetValue & "") = 0 Then
'        MsgMustBe ("贈點到期日")
'        gdaBonusStopDate.SetFocus
'        Exit Function
'    End If
'    If Val(gdaBonusStopDate.GetValue) < Val(Format(Now, "yyyymmdd")) Then
'        MsgBox "贈點到期日不可小於今天日期！", vbOKOnly, "警告"
'        gdaBonusStopDate.SetFocus
'        Exit Function
'    End If
    IsDataOk = True
  Exit Function
ChkErr:
  Call ErrSub(Me.Name, "IsDataOK")
End Function

Private Sub cmdCancel_Click()
  On Error Resume Next
    CloseRecordset rsSO004
    CloseRecordset rsSO193
    Unload Me
End Sub
'0=皆可,1=儲值點數,2=紅利點數
Private Sub FilterCondition()
  On Error GoTo ChkErr:
    'Call GetChgDefValue
    rsSO193.Filter = ""
    Select Case ChangeCondition
    Case 0
        rsSO193.Filter = ""
    Case 1
        rsSO193.Filter = "Savepoint > 0 "
    Case 2
         rsSO193.Filter = "bonus > 0  "
    End Select
    Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "FilterCondition")
End Sub
Private Sub cmdOK_Click()
  On Error GoTo ChkErr
    Me.Enabled = False
    Me.MousePointer = vbHourglass
    If Not IsDataOk Then GoTo 88
    gcnGi.BeginTrans
    If Not InsData Then gcnGi.RollbackTrans: GoTo 88
    If Not UpdData Then gcnGi.RollbackTrans: GoTo 88
    'If Not CalcVODPoint(rsSO004("VodAccountId") & "") Then GoTo 88
    gcnGi.CommitTrans
    If Not GetRS(rsSO193, rsSO193.Source, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
    IsSave = True
    MsgBox "轉換成功！", vbInformation, "訊息"
88:
    
    If rsSO193.RecordCount > 0 Then
        rsSO193.MoveFirst
        totalBonus = Val(rsSO193("TotalBonus") & "")
        totalSavepoint = Val(rsSO193("TotalSavePoint") & "")
    End If
    Me.Enabled = True
    Me.MousePointer = vbDefault
    Exit Sub
ChkErr:
  gcnGi.RollbackTrans
  Me.Enabled = True
  Me.MousePointer = vbDefault
  Call ErrSub(Me.Name, "cmdOk_Click")
End Sub
Private Function GetTransseqno() As String
  On Error GoTo ChkErr
    GetTransseqno = Format(Now, "yyyymmdd") & Right("000000000000" & _
                    RPxx(gcnGi.Execute("SELECT " & GetOwner & "S_SO033BBTRANS.NEXTVAL FROM DUAL").GetString & ""), 12)
                        
  Exit Function
ChkErr:
    ErrSub Me.Name, "Get_SeqNo_Seq"
End Function
Private Sub GetChgDefValue()
  On Error GoTo ChkErr
    Dim rsCD129 As New Recordset
    If Not GetRS(rsCD129, "Select Nvl(Condition,0) Condition,RefNo,Nvl(Rate,1) Rate,CitemCode " & _
                        " From " & GetOwner & "CD129  Where CodeNo = " & gilChangeStore.GetCodeNo, _
                        gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
    ChangeCondition = Val(rsCD129("Condition") & "")
    ChangeRate = Val(rsCD129("Rate") & "")
    ChangeRefNo = Val(rsCD129("RefNo") & "")
    ChangeCitemCode = rsCD129("CitemCode") & ""
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "GetChgDefValue")
End Sub
Private Sub FunctionKey(KeyCode As Integer, Shift As Integer)
    On Error GoTo ChkErr
    
        If KeyCode = vbKeyF2 Then Call cmdOK_Click
        If KeyCode = vbKeyEscape Then Call cmdCancel_Click
'        If KeyCode = vbKeyF3 And cmdFind.Enabled Then Call cmdFind_Click
'        If KeyCode = vbKeyF4 And cmdEdit.Enabled Then Call cmdEdit_Click
'        If KeyCode = vbKeyEscape And cmdCancel.Enabled Then Call cmdCancel_Click
'        If KeyCode = vbKeyX And Shift = 2 Then cmdCancel.Value = True
'
'        If KeyCode = vbKeyUp Then
'            KeyCode = 0
'            SendKeys "+{Tab}"
'        End If
'        If KeyCode = vbKeyDown Then
'            KeyCode = 0
'            SendKeys "{Tab}"
'        End If
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "FunctionKey")
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    FunctionKey KeyCode, Shift
End Sub

Private Sub Form_Load()
    IsSave = False
    Me.MousePointer = vbHourglass
    totalBonus = 0
    totalSavepoint = 0
    If rsSO193.RecordCount > 0 Then
        rsSO193.MoveFirst
        totalBonus = Val(rsSO193("TotalBonus") & "")
        totalSavepoint = Val(rsSO193("TotalSavePoint") & "")
        txtTotalBonus.Text = totalBonus
        txtTotalSavePoint.Text = totalSavepoint
    End If
    lblCondition.Caption = "無"
    lblbbTranspriority.Caption = ""
    Call subGil
    Call DefaultValue
    Me.MousePointer = vbDefault
End Sub
Public Property Let BBVodAccountId(ByVal vData As String)
  fBBVodAccountId = vData
End Property
Public Property Set uRS193(ByRef vData As Recordset)
    Set rsSO193 = vData.Clone
End Property
'Public Property Get uRS193() As Recordset
'    Set uRS = rsSO193.Clone
'End Property
Public Property Set uRS004(ByRef vData As Recordset)
    Set rsSO004 = vData
    If rsSO004.RecordCount > 0 Then
        rsSO004.MoveFirst
        fBBVodAccountId = rsSO004("FaciId") & ""
    End If
End Property
'Public Property Get uRS004() As Recordset
'    Set uRS = rsSO004.Clone
'End Property
Public Property Get uSave() As Boolean
  uSave = IsSave
End Property
'
Private Function CalcVODPoint(ByVal VODAccountId As String, ByVal Sign As String, ByRef rs033 As Recordset) As Boolean
    On Error GoTo ChkErr
    If ChangeRefNo <> 1 Or Len(VODAccountId) = 0 Then
        CalcVODPoint = True
        Exit Function
    End If
    
    Dim clsChargeInterface As Object
    Dim clsAlterSO003 As Object
    Set clsChargeInterface = CreateObject("csAlterSO003.clsInterface")
    With clsChargeInterface
        Set .uConn = gcnGi
        .uOwnerName = GetOwner
        Set clsAlterSO003 = .objAction("csAlterSO003.clsAlterSO003")
        If .uErr > 0 Then Exit Function
    End With
    If Sign = "+" Then
        CalcVODPoint = clsAlterSO003.CalcVODPoint(rs033)
    Else
        CalcVODPoint = clsAlterSO003.CalcBBCashPoint(rs033)
    End If
    Set clsChargeInterface = Nothing
    Set clsAlterSO003 = Nothing
    Exit Function
ChkErr:
    Call ErrSub(Me.Name, "CalcVODPoint")
End Function

Private Sub gilChangeStore_Change()
  On Error GoTo ChkErr
    If Len(gilChangeStore.GetCodeNo & "") = 0 Then
        lblCondition.Caption = "無"
        Exit Sub
    End If
  Call GetChgDefValue
  
  Select Case ChangeCondition
    Case 0
        lblCondition.Caption = "皆可"
    Case 1
        lblCondition.Caption = "儲值點數"
    Case 2
        lblCondition.Caption = "紅利點數"
  End Select
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "gilChangeStore_Change")
End Sub

Private Sub txtChangePoint_Validate(Cancel As Boolean)
  On Error GoTo ChkErr:
    Me.Enabled = False
    Me.MousePointer = vbHourglass
    If Not IsDataOk Then
        If Len(gilChangeStore.GetCodeNo & "") = 0 Then
            Cancel = False
            gilChangeStore.SetFocus
        Else
            Cancel = True
        End If
        Me.Enabled = True
        Me.MousePointer = vbDefault
        Exit Sub
    End If
    Dim transSavepoint As Double
    Dim transbonus As Double
    transbonus = 0
    transSavepoint = 0
    If rsSO193.RecordCount > 0 Then rsSO193.MoveFirst
    If rsCalculate.RecordCount > 0 Then
        rsCalculate.MoveFirst
        Do While Not rsCalculate.EOF
        Select Case Val(rsCalculate("MinusType"))
            Case 2
                transbonus = transbonus + Val(rsCalculate("MinusPoint"))
            Case 1
                transSavepoint = transSavepoint + Val(rsCalculate("MinusPoint"))
        End Select
        rsCalculate.MoveNext
    Loop
    End If
    
    txtUsedBonus.Text = transbonus
    txtUsedSavePoint.Text = transSavepoint
    txtRate.Text = ChangeRate
    txtMallbonus.Text = RoundCS(transbonus * ChangeRate, 0)
    txtMallpoint.Text = RoundCS(transSavepoint * ChangeRate, 0)
    txtTotalSavePoint.Text = totalSavepoint - transSavepoint
    txtTotalBonus.Text = totalBonus - transbonus
    Me.Enabled = True
    Me.MousePointer = vbDefault
    Exit Sub
ChkErr:
  Me.Enabled = True
  Me.MousePointer = vbDefault
  Call ErrSub(Me.Name, "txtChangePoint_Validate")
End Sub

