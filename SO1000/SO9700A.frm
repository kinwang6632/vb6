VERSION 5.00
Object = "{6DDA0569-6454-11D3-9612-0048544B09DF}#4.1#0"; "GiTime.ocx"
Object = "{3BB22E57-824A-11D3-BD70-0080C8F80BC4}#3.4#0"; "GiGridR.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Begin VB.Form frmSO9700A 
   Caption         =   "Form1"
   ClientHeight    =   6615
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   9420
   BeginProperty Font 
      Name            =   "新細明體"
      Size            =   9.75
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SO9700A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6615
   ScaleWidth      =   9420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton cmdCheck 
      Caption         =   "審單"
      Height          =   405
      Left            =   1200
      TabIndex        =   23
      Top             =   6030
      Width           =   1065
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "查詢"
      Height          =   405
      Left            =   90
      TabIndex        =   21
      Top             =   6030
      Width           =   1065
   End
   Begin VB.Frame Frame2 
      Height          =   3765
      Left            =   60
      TabIndex        =   19
      Top             =   2100
      Width           =   9255
      Begin prjGiGridR.GiGridR ggrData 
         Height          =   3555
         Left            =   60
         TabIndex        =   20
         Top             =   120
         Width           =   9105
         _ExtentX        =   16060
         _ExtentY        =   6271
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1965
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   9195
      Begin VB.CheckBox chkCheckFlag 
         Caption         =   "審單未完成"
         Height          =   315
         Left            =   5910
         TabIndex        =   22
         Top             =   1410
         Width           =   1935
      End
      Begin VB.TextBox txtWip1 
         Height          =   315
         Left            =   1080
         TabIndex        =   10
         Top             =   990
         Width           =   2205
      End
      Begin VB.TextBox txtOrder1 
         Height          =   315
         Left            =   5940
         TabIndex        =   8
         Top             =   570
         Width           =   1995
      End
      Begin VB.TextBox txtCustId2 
         Height          =   315
         Left            =   3240
         TabIndex        =   5
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtCustId1 
         Height          =   315
         Left            =   1080
         TabIndex        =   4
         Top             =   600
         Width           =   1575
      End
      Begin prjGiList.GiList gilCompCode 
         Height          =   315
         Left            =   1080
         TabIndex        =   2
         Top             =   210
         Width           =   2940
         _ExtentX        =   5186
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
         F2Corresponding =   -1  'True
      End
      Begin Gi_Time.GiTime gdtAccTime1 
         Height          =   315
         Left            =   4290
         TabIndex        =   12
         Top             =   990
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Gi_Time.GiTime gdtAccTime2 
         Height          =   315
         Left            =   6780
         TabIndex        =   13
         Top             =   990
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Gi_Time.GiTime gdtRevTime1 
         Height          =   315
         Left            =   1080
         TabIndex        =   16
         Top             =   1410
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Gi_Time.GiTime gdtRevTime2 
         Height          =   315
         Left            =   3540
         TabIndex        =   17
         Top             =   1410
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label9 
         Caption         =   "~"
         Height          =   255
         Left            =   3360
         TabIndex        =   18
         Top             =   1470
         Width           =   165
      End
      Begin VB.Label Label8 
         Caption         =   "預約時間"
         Height          =   255
         Left            =   150
         TabIndex        =   15
         Top             =   1470
         Width           =   795
      End
      Begin VB.Label Label7 
         Caption         =   "~"
         Height          =   255
         Left            =   6570
         TabIndex        =   14
         Top             =   1020
         Width           =   165
      End
      Begin VB.Label Label6 
         Caption         =   "受理時間"
         Height          =   255
         Left            =   3420
         TabIndex        =   11
         Top             =   1080
         Width           =   795
      End
      Begin VB.Label Label5 
         Caption         =   "工單單號"
         Height          =   255
         Left            =   150
         TabIndex        =   9
         Top             =   1080
         Width           =   795
      End
      Begin VB.Label Label4 
         Caption         =   "訂單單號"
         Height          =   255
         Left            =   5040
         TabIndex        =   7
         Top             =   630
         Width           =   795
      End
      Begin VB.Label Label3 
         Caption         =   "~"
         Height          =   255
         Left            =   2880
         TabIndex        =   6
         Top             =   630
         Width           =   165
      End
      Begin VB.Label Label2 
         Caption         =   "客戶編號"
         Height          =   255
         Left            =   150
         TabIndex        =   3
         Top             =   660
         Width           =   795
      End
      Begin VB.Label Label1 
         Caption         =   "公司別"
         Height          =   255
         Left            =   150
         TabIndex        =   1
         Top             =   270
         Width           =   795
      End
   End
End
Attribute VB_Name = "frmSO9700A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private strWhere As String
Private rsSO007 As New ADODB.Recordset

Private Sub cmdCheck_Click()
  On Error GoTo ChkErr
    
    With frmSO9700B
        'Set .uRsSource = rsSO007
        .uCustId = rsSO007("CustId")
        .uOrderNo = rsSO007("OrderNo")
        .Show 1
    End With
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdCheck_Click")
End Sub

Private Sub cmdSearch_Click()
  On Error GoTo ChkErr
    OpenWipData
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdSearch_Click")
End Sub

Private Sub Form_Load()
  On Error GoTo ChkErr
    subGil
    subGrd
    DefaultValue
    
    Exit Sub
ChkErr:
 Call ErrSub(Me.Name, "Form_Load")
End Sub
Private Sub DefaultValue()
  On Error GoTo ChkErr
    gdtAccTime1.ShowSecond
    gdtAccTime2.ShowSecond
    gdtRevTime1.ShowSecond
    gdtRevTime2.ShowSecond
    gdtAccTime1.SetValue Format(Now, "yyyymmdd") & "000000"
    gdtAccTime2.SetValue Format(Now, "yyyymmdd") & "235959"
    
  Exit Sub
ChkErr:
 Call ErrSub(Me.Name, "DefaultValue")
End Sub



Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
  Call CloseRecordset(rsSO007)
End Sub

Private Sub gdtAccTime1_Validate(Cancel As Boolean)
  On Error Resume Next
    If gdtAccTime1.GetValue & "" <> "" Then
        If gdtAccTime2.GetValue & "" = "" Then
            gdtAccTime2.SetValue gdtAccTime1.GetDate & "235959"
        End If
    End If
End Sub

Private Sub gdtAccTime2_Validate(Cancel As Boolean)
  On Error Resume Next
    Cancel = ChkDate2(gdtAccTime1, gdtAccTime2)
End Sub


Private Sub gdtRevTime1_Validate(Cancel As Boolean)
  On Error Resume Next
    If gdtRevTime1.GetValue & "" <> "" Then
        If gdtRevTime2.GetValue & "" = "" Then
            gdtRevTime2.SetValue gdtRevTime1.GetDate & "235959"
        End If
    End If
End Sub

Private Sub gdtRevTime2_Change()
  On Error Resume Next
    Cancel = ChkDate2(gdtRevTime1, gdtRevTime2)
End Sub
Private Sub subGil()
  On Error GoTo ChkErr
    SetgiList gilCompCode, "CodeNo", "Description", "CD039", , , , , 3, 12, GetCompCodeFilter(gilCompCode)
    gilCompCode.SetCodeNo garyGi(9)
    gilCompCode.Query_Description
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "subGil")
End Sub

Private Sub ggrData_ShowCellData(ByVal Col As Long, ByVal Row As Long, ByVal giFld As prjGiGridR.GiGridFld, ByVal Fld As ADODB.Field, Value As Variant)
  On Error GoTo ChkErr
    If UCase(giFld.FieldName) = UCase("APICheckOk") Then
        If Value = "1" Then
            Value = "V"
        End If
    End If
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "ggrData_ShowCellData")
End Sub

Private Sub gilCompCode_Change()
  On Error Resume Next
    If Not ChgComp(gilCompCode, "SO3100", "SO3110") Then Exit Sub
    Call subGil
End Sub
Private Sub OpenWipData()
  On Error GoTo ChkErr
    Call subChoose
    Dim strQry As String
    strQry = "SELECT * FROM " & GetOwner & "SO007 A WHERE " & strWhere
    If Not GetRS(rsSO007, strQry, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
    Set rsSO007.ActiveConnection = Nothing
    Set ggrData.Recordset = rsSO007
    ggrData.Refresh
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "OpenWipData")
End Sub
Private Sub subChoose()
  On Error GoTo ChkErr
    strWhere = Empty
    Call subAnd2(strWhere, "A.CompCode=" & gilCompCode.GetCodeNo)
    If txtCustId1.Text <> Empty Then Call subAnd2(strWhere, "A.CUSTID >= " & txtCustId1.Text)
    If txtCustId2.Text <> Empty Then Call subAnd2(strWhere, "A.CUSTID <= " & txtCustId2.Text)
    If txtOrder1.Text <> Empty Then Call subAnd2(strWhere, "A.ORDERNO = '" & txtOrder1.Text & "'")
    If txtWip1.Text <> Empty Then Call subAnd2(strWhere, "A.SNO = '" & txtWip1.Text & "'")
    If gdtAccTime1.GetValue <> Empty Then
        Call subAnd2(strWhere, "A.AcceptTime >= To_DATE('" & gdtAccTime1.GetValue & "','YYYYMMDDHH24Miss')")
    End If
    If gdtAccTime2.GetValue <> Empty Then
        Call subAnd2(strWhere, "A.AcceptTime <= To_DATE('" & gdtAccTime2.GetValue & "','YYYYMMDDHH24Miss')")
    End If
    If gdtRevTime1.GetValue <> Empty Then
        Call subAnd2(strWhere, "A.ResvTime >= To_DATE('" & gdtRevTime1.GetValue & "','YYYYMMDDHH24Miss')")
    End If
    If gdtRevTime2.GetValue <> Empty Then
        Call subAnd2(strWhere, "A.ResvTime <= To_DATE('" & gdtRevTime2.GetValue & "','YYYYMMDDHH24Miss')")
    End If
    If chkCheckFlag.Value Then
        Call subAnd2(strWhere, "A.CreateByAPI = 1 AND NVL(A.APICheckOk,0) = 0")
    Else
        Call subAnd2(strWhere, "A.CreateByAPI = 1")
    End If
    Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "subChoose")
End Sub
Private Sub subGrd()
  On Error GoTo ChkErr
    Dim mFlds As New GiGridFlds
        With mFlds
            .Add "APICheckOk", , , , , "已審核"
            .Add "CUSTID", , , , , "客戶編號"
            .Add "ORDERNO", , , , , "訂單單號" & Space(10)
            .Add "SNo", , , , , "工單單號" & Space(10)
            .Add "IVRShortSNo", , , , , "IVR 簡碼    "
            .Add "AcceptTime", , , , , "受理日期" & Space(20)
            .Add "ResvTime", , , , , "預約日期" & Space(20)
        End With
        With ggrData
            .UseCellForeColor = True
            .AllFields = mFlds
            .SetHead
        End With
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "subGrd")
End Sub
