VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmPrint 
   BorderStyle     =   1  '單線固定
   Caption         =   "確定列印"
   ClientHeight    =   8280
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14115
   Icon            =   "frmPrint.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8280
   ScaleWidth      =   14115
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton cmdOK 
      Caption         =   "列印"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   26.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   60
      TabIndex        =   4
      Top             =   7560
      Width           =   1995
   End
   Begin VB.CheckBox chkSellAll 
      Caption         =   "全部選取"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   24
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   525
      Left            =   120
      TabIndex        =   1
      Top             =   180
      Width           =   2595
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdData 
      Height          =   6675
      Left            =   60
      TabIndex        =   0
      ToolTipText     =   "滑鼠左鍵點2下即可選擇"
      Top             =   810
      Width           =   13965
      _ExtentX        =   24633
      _ExtentY        =   11774
      _Version        =   393216
      HighLight       =   2
      SelectionMode   =   1
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "標楷體"
         Size            =   15.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin Crystal.CrystalReport crpWork 
      Left            =   2130
      Top             =   7650
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "裝機派工單"
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label Label1 
      Caption         =   "金額小計"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   24
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   555
      Left            =   6990
      TabIndex        =   3
      Top             =   180
      Width           =   2025
   End
   Begin VB.Label lblAmt 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   24
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   555
      Left            =   9090
      TabIndex        =   2
      Top             =   180
      Width           =   3525
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private strCstRowId As String
Private rsSO033 As New ADODB.Recordset
Private rsTmp As New ADODB.Recordset
Private strCustId As String
Private strTel As String
Private blnPrintOK As Boolean
Private strSQL33 As String
Private lngBatchNo As Long

Private Sub chkSellAll_Click()
  On Error GoTo ChkErr
    Dim i As Long
    Dim lngAmt As Long
    If chkSellAll.Value = 0 Then Exit Sub
    LockWindowUpdate grdData.hwnd
    'grdData.Row = 1
    For i = 1 To grdData.Rows - 1
        grdData.Row = i
        rsTmp.AbsolutePosition = i
        lngAmt = lngAmt + Val(rsTmp("ShouldAmt") & "")
        rsTmp("Choice") = "1"
        rsTmp.Update
        grdData.TextMatrix(grdData.Row, 1) = "V"
        grdData.Col = 1
        grdData.CellForeColor = vbRed
        grdData.Refresh
        'grdData_RowColChange
        'rsTmp.MoveNext
    Next i
    lblAmt.Caption = lngAmt
    LockWindowUpdate 0
    Exit Sub
ChkErr:
  LockWindowUpdate 0
  Call ErrSub(Me.Name, "chkSellAll")
End Sub

Private Sub cmdOK_Click()
  On Error GoTo ChkErr
    Dim strRows As String
    Dim blnCombineBill As Boolean
    Dim lngAffecteds As Long
    lngAffecteds = 0
    strSQL33 = Empty
    Screen.MousePointer = vbHourglass
    rsTmp.MoveFirst
    Do While Not rsTmp.EOF
        If rsTmp("Choice") = "1" Then
            'strRows = IIf(strRows = Empty, "'" & rsTmp("RowId") & "'", strRows & ",'" & rsTmp("RowId") & "'")
            If strSQL33 = "" Then
                strSQL33 = " (A.BillNo='" & rsTmp("BillNo") & "' And A.Item=" & rsTmp("Item") & ")"
            Else
                strSQL33 = strSQL33 & " Or (A.BillNo='" & rsTmp("BillNo") & "' And A.Item=" & rsTmp("Item") & ")"
            End If
            lngAffecteds = lngAffecteds + 1
        End If
        rsTmp.MoveNext
    Loop
    If strSQL33 = Empty Then
        MsgBox "請挑選列印的資料！", vbInformation, "警告"
        blnPrintOK = False
        Screen.MousePointer = vbDefault
        Exit Sub
    Else
        strSQL33 = " (" & strSQL33 & ")"
        lngBatchNo = GetRsValue("SELECT " & GetOwner & "S_SO075_BatchNo.NextVal FROM DUAL", gcnGi)
        If Not InsSO075 Then Exit Sub
        PrintChargeBill crpWork, Me.Name, Me, 1, gCompCode, , strBillPrinter2, strBillTab6, , , , lngAffecteds, , , blnCombineBill
    End If
    blnPrintOK = True
    Unload Me
    Screen.MousePointer = vbDefault
    Exit Sub
ChkErr:
    Screen.MousePointer = vbDefault
  Call ErrSub(Me.Name, "cmdOK_Click")
End Sub

Private Sub Form_Activate()
  On Error Resume Next
    Screen.MousePointer = vbDefault
    
    'SetGrdData
End Sub
Private Function OpenData() As Boolean
  On Error GoTo ChkErr
    Dim strQry As String
    Dim i As Long
    strQry = Empty
    strQry = "Select Distinct A.CustId,B.CustName,C.DeclarantName,D.Description,C.SeqNo,A.CitemCode," & _
            "A.ShouldAmt,A.RealStartDate,A.RealStopDate,C.ContTel,B.Tel1,D.CPDescription, " & _
            "A.MediaBillNo,A.BillNo,A.Item,A.RowId " & _
            " From " & GetOwner & "SO033 A," & GetOwner & "SO001 B," & _
            GetOwner & "SO004 C," & GetOwner & "CD019 D"
           
    If Not GetRS(rsSO033, strQry & " Where 1=0", gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
    With rsTmp
        If .State = adStateOpen Then .Close
        .CursorLocation = adUseClient
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Fields.Append "Choice", adBSTR, 1, adFldIsNullable
        For i = 0 To rsSO033.Fields.Count - 1
            .Fields.Append rsSO033.Fields(i).Name, adBSTR, rsSO033.Fields(i).DefinedSize, adFldIsNullable
            If i = 2 Then
                .Fields.Append "CustNameA", adBSTR, 50, adFldIsNullable
                .Fields.Append "TelA", adBSTR, 20, adFldIsNullable
                .Fields.Append ("CitemNameA"), adBSTR, 40, adFldIsNullable
            End If
        Next i
    End With
    rsTmp.Open
    
    'strQry = subChoose
    strQry = strQry & " Where " & _
            "A.CustId=B.CustId And A.CitemCode=D.CodeNo" & _
            " And A.CustId=C.Custid(+) And A.FaciSeqNo=C.SeqNo(+) " & _
            " And A.CustId=" & strCustId & " And A.CompCode=" & gCompCode & _
            " And A.UCCode is Not Null And A.CancelFlag=0"
    
    If Not GetRS(rsSO033, strQry, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
    If rsSO033.RecordCount > 0 Then
        rsSO033.MoveFirst
    End If
    Do While Not rsSO033.EOF
        rsTmp.AddNew
        For i = 0 To rsSO033.Fields.Count - 1
            If Not IsNull(rsSO033.Fields(i).Value) Then
                rsTmp.Fields(rsSO033.Fields(i).Name).Value = rsSO033.Fields(i).Value & ""
            End If
        Next i
        If Not IsNull(rsSO033("DeclarantName")) Then
            rsTmp("CustNameA") = rsSO033("DeclarantName") & ""
        Else
            rsTmp("CustNameA") = rsSO033("CustName") & ""
        End If
        If strTel <> "" Then
            rsTmp("TelA") = strTel
        Else
            rsTmp("TelA") = rsSO033("Tel1") & ""
        End If
        If Not IsNull(rsSO033("CPDescription")) Then
            rsTmp("CitemNameA") = rsSO033("CPDescription") & ""
        Else
            rsTmp("CitemNameA") = rsSO033("Description") & ""
        End If
        rsTmp.Update
        rsSO033.MoveNext
    Loop
    OpenData = True
    On Error Resume Next
    
    Exit Function
ChkErr:
  ErrSub Me.Name, "OpenData"
End Function
Property Let uRowId(ByVal vData As String)
  On Error Resume Next
    strCstRowId = vData
End Property
Property Let uTel(ByVal vData As String)
  On Error Resume Next
    strTel = vData
End Property
Property Let uCustId(ByVal vData As String)
    On Error Resume Next
    strCustId = vData
End Property
Private Sub subGrid()
'    On Error Resume Next
'    Dim mFlds As New GiGridFlds
'    With mFlds
'        .Add "CustId", , , , , "客戶編號" & Space(10)
'        .Add "CustNameA", , , , , "客戶名稱" & Space(30)
'        .Add "TelA", , , , , "電話" & Space(30)
'        .Add "CitemNameA", , , , , "收費項目" & Space(30)
'        .Add "ShouldAmt", , , , , "應收金額" & Space(10)
'        .Add "RealStartDate", , , , , "起始日" & Space(10)
'        .Add "RealStopDate", , , , , "截止日" & Space(10)
'    End With
'    ggrData.AllFields = mFlds
'    ggrData.SetHead
'    Set ggrData.Recordset = rsTmp
'    ggrData.Refresh
End Sub
Private Sub SetGrdData()
  On Error GoTo ChkErr
    Dim i As Long
    grdData.Clear
    Set grdData.Recordset = rsTmp
    'grdData.FixedCols = 2
    grdData.ColWidth(0, 0) = 400
    For i = 0 To rsTmp.Fields.Count - 1
        Select Case UCase(rsTmp.Fields(i).Name)
            Case "CUSTID"
                grdData.ColWidth(i + 1) = 1500
                grdData.TextMatrix(0, i + 1) = "客戶編號"
                grdData.CellAlignment = vbRightJustify
            Case "CUSTNAMEA"
                grdData.ColWidth(i + 1) = 2000
                grdData.TextMatrix(0, i + 1) = "客戶名稱"
            Case "TELA"
                grdData.ColWidth(i + 1) = 2000
                grdData.TextMatrix(0, i + 1) = "電話"
            Case "CITEMNAMEA"
                grdData.ColWidth(i + 1) = 4000
                grdData.TextMatrix(0, i + 1) = "收費項目"
            Case "SHOULDAMT"
                grdData.ColWidth(i + 1) = 1800
                grdData.TextMatrix(0, i + 1) = "應收金額"
                grdData.CellAlignment = vbRightJustify
            Case "REALSTARTDATE"
                grdData.ColWidth(i + 1) = 1800
                grdData.TextMatrix(0, i + 1) = "起始日"
            Case "REALSTOPDATE"
                grdData.ColWidth(i + 1) = 1800
                grdData.TextMatrix(0, i + 1) = "截止日"
            Case "CHOICE"
                grdData.ColWidth(i + 1) = 800
                grdData.TextMatrix(0, i + 1) = "選擇"
                grdData.Col = i + 1
            Case Else
                grdData.ColWidth(i + 1) = 0
        End Select
    Next i
    If rsTmp.RecordCount = 0 Then grdData.Rows = 2: grdData.Row = 1
    Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "SetGrdData")
End Sub
Private Sub Form_Load()
  On Error GoTo ChkErr
    blnPrintOK = False
    
    OpenData
    SetGrdData
    grdData.Refresh
    Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "Form_Load")
End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  On Error Resume Next
  frmMain.uOK = blnPrintOK
  Close3Recordset rsSO033
  Close3Recordset rsTmp
End Sub

Private Sub grdData_DblClick()
    LockWindowUpdate grdData.hwnd
    Dim lngAbs As Long
    Dim strMediaBillNo As String
    Dim blnChoice As Boolean
    Dim lngAmt As Long
    rsTmp.AbsolutePosition = grdData.Row
    lngAbs = rsTmp.AbsolutePosition
    If rsTmp("Choice") = "1" Then
        grdData.TextMatrix(grdData.Row, 1) = ""
        rsTmp("Choice") = ""
        rsTmp.Update
        strMediaBillNo = rsTmp("MediabillNo") & ""
        blnChoice = False
        If chkSellAll.Value = 1 Then chkSellAll.Value = 0
    Else
        grdData.Col = 1
        grdData.TextMatrix(grdData.Row, 1) = "V"
        grdData.CellForeColor = vbRed
        
        rsTmp("Choice") = "1"
        rsTmp.Update
        strMediaBillNo = rsTmp("MediabillNo") & ""
        blnChoice = True
    End If
    rsTmp.MoveFirst
    Do While Not rsTmp.EOF
        If strMediaBillNo <> "" Then
            grdData.Row = rsTmp.AbsolutePosition
            grdData.CellForeColor = vbRed
            If blnChoice Then
                If rsTmp("MediaBillNo") = strMediaBillNo Then
                    rsTmp("Choice") = "1"
                    rsTmp.Update
                    grdData.TextMatrix(grdData.Row, 1) = "V"
                End If
            Else
                If rsTmp("MediaBillNo") = strMediaBillNo Then
                    rsTmp("Choice") = ""
                    rsTmp.Update
                    grdData.TextMatrix(grdData.Row, 1) = ""
                End If
            End If
            grdData.Refresh
        End If
        If rsTmp("Choice") = "1" Then
            lngAmt = lngAmt + Val(rsTmp("ShouldAmt"))
        End If
        rsTmp.MoveNext
    Loop
    lblAmt.Caption = lngAmt
    rsTmp.AbsolutePosition = lngAbs
    grdData.Refresh
    grdData.Row = lngAbs
'    rsTmp.Update
    grdData.Refresh
    LockWindowUpdate 0
End Sub

Private Sub grdData_RowColChange()
    If grdData.TextMatrix(grdData.Row, 1) = "V" Then
        grdData.Col = 1
        grdData.CellForeColor = vbRed
        grdData.Refresh
    End If
End Sub
'取得所有印表機並得預設印表機名稱
Private Function GetDefPrinter() As String
  On Error GoTo ChkErr
    Dim o As Printer
    Dim strComputer As String
    Dim objWMIService As Object
    Dim colItems As Object
    Dim objItem As Object
    Dim strDefault As String
    strComputer = "."
    Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
    Set colItems = objWMIService.ExecQuery( _
        "SELECT * FROM Win32_Printer", , 48)
    For Each objItem In colItems
        If objItem.Default Then '判斷是否為預設
            strDefault = objItem.Caption    '印表機名稱
        End If
        
    Next
    GetDefPrinter = strDefault
    Exit Function
ChkErr:
  Call ErrSub(Me.Name, "GetDefPrinter")
End Function
Private Function InsSO075() As Boolean
  On Error GoTo ChkErr
    Dim rsSO075 As New ADODB.Recordset
    If Not GetRS(rsSO075, "Select * From SO075 Where 1=0", gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
    gcnGi.BeginTrans
        rsSO075.AddNew
        With rsSO075
            .Fields("CUSTID") = strCustId
            .Fields("BATCHNO") = lngBatchNo
            .Fields("COMPCODE") = gCompCode
            .Update
        End With
    gcnGi.CommitTrans
    On Error Resume Next
    Close3Recordset rsSO075
    InsSO075 = True
    Exit Function
ChkErr:
  Call ErrSub(Me.Name, "InsSO075")
End Function
Property Get uSQL33() As String
    uSQL33 = strSQL33
End Property
Property Get uBatchNo() As Long
    uBatchNo = lngBatchNo
End Property
