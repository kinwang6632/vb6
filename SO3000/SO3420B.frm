VERSION 5.00
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Begin VB.Form frmSO3420B 
   Caption         =   "派工單之服務申告 [SO3420B] "
   ClientHeight    =   2040
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5400
   Icon            =   "SO3420B.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   5400
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "取消(&X)"
      Height          =   435
      Left            =   4170
      TabIndex        =   6
      Top             =   1560
      Width           =   1035
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "F2. 確定"
      Height          =   435
      Left            =   150
      TabIndex        =   5
      Top             =   1560
      Width           =   1035
   End
   Begin VB.Frame fraData 
      Height          =   1395
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   5115
      Begin prjGiList.GiList gilServiceCode 
         Height          =   315
         Left            =   1350
         TabIndex        =   3
         Top             =   300
         Width           =   2940
         _ExtentX        =   5186
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin prjGiList.GiList gilServiceContent 
         Height          =   315
         Left            =   1350
         TabIndex        =   4
         Top             =   690
         Width           =   2940
         _ExtentX        =   5186
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblServiceContent 
         Caption         =   "申告內容"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   390
         TabIndex        =   2
         Top             =   750
         Width           =   915
      End
      Begin VB.Label lblServiceCode 
         Caption         =   "來電分類"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   390
         TabIndex        =   1
         Top             =   330
         Width           =   915
      End
   End
End
Attribute VB_Name = "frmSO3420B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private strServiceType As String
Private blnIsOK As Boolean

Private Sub cmdExit_Click()
  On Error Resume Next
    Unload Me
End Sub

Private Sub cmdOK_Click()
  On Error Resume Next
    If gilServiceCode.GetCodeNo <> "" And gilServiceContent.GetCodeNo <> "" Then
        blnIsOK = True
    Else
        blnIsOK = False
    End If
    Unload Me
End Sub

Private Sub Form_Activate()
On Error Resume Next
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
        If Not ChkGiList(KeyCode) Then Exit Sub
        If KeyCode = vbKeyF2 Then
            Call cmdOK_Click
        End If

End Sub

Private Sub Form_Load()
    Call DefaultValue
    blnIsOK = False
End Sub

Private Sub DefaultValue()
  On Error GoTo chkErr
    '#3579 服務別為Null的也要顯示出來 By Kin 2007/10/22
    '來電分類代碼要過濾ServiceType，並預設為第一筆
'    SetgiList gilServiceCode, "CodeNo", "Description", "CD008", , , , , 3, 12, _
'                "Where (ServiceType='" & strServiceType & "' Or ServiceType is Null) And RefNo=5", True
    '#7026 來電分類 語法調整 By Kin 2015/04/15
     SetgiList gilServiceCode, "CodeNo", "Description", "CD008", , , , , 3, 12, _
                " Where CodeNo IN ( SELECT  A.CODENO FROM  " & GetOwner & "CD008 A," & GetOwner & "CD008C C " & _
                                            " Where A.CODENO = C.SERVICECODE " & _
                                            " AND (EXISTS (SELECT 1 FROM " & GetOwner & "CD008A B WHERE  " & _
                                                " (B.ServiceType = '" & strServiceType & "' Or B.ServiceType is Null) And B.RefNo=5 " & _
                                                " AND Nvl(B.STOPFLAG,0) = 0 AND C.CODENO=B.CODENO) " & _
                                                  " Or Refno = 5) " & _
                                                    " AND Nvl(A.STOPFLAG,0) =0 )"
    gilServiceCode.ListIndex = 1
    If gilServiceCode.GetCodeNo = "" Then
        SetgiList gilServiceContent, "CodeNo", "Description", "CD008A", , , , , , , " Where 1=0"
    End If
    Exit Sub
chkErr:
    ErrSub Me.Name, "DefaultValue"
End Sub
'SO3420B傳進來的ServiceType
Public Property Let uServiceType(ByVal vData As String)
    strServiceType = vData
End Property

Private Sub Form_Unload(Cancel As Integer)
 On Error Resume Next
    frmSO3420A.uIsOk = blnIsOK
    frmSO3420A.uServiceCode = gilServiceCode.GetCodeNo
    frmSO3420A.uServiceName = gilServiceCode.GetDescription
    frmSO3420A.uServiceContentCode = gilServiceContent.GetCodeNo
    frmSO3420A.uServiceContentName = gilServiceContent.GetDescription
End Sub

'來電分類一改變，申告內容要跟著連動
Private Sub gilServiceCode_Change()
  On Error GoTo chkErr
'    Dim strCodeNo As String
'    strCodeNo = gcnGi.Execute("Select CodeNo From " & GetOwner & "CD008C Where ServiceCode=" & gilServiceCode.GetCodeNo)(0).Value
'    SetgiList gilServiceContent, "CodeNo", "Description", "CD008C", , , , , , , " Where ServiceType='" & strServiceType & "' And ServiceCode=" & gilServiceCode.GetCodeNo
    '#3579測試不OK 申告內容的資料CD008A.Code 要In CD008C.CodeNo By Kin 2007/10/26
    SetgiList gilServiceContent, "CodeNo", "Description", "CD008A", , , , , , , " Where (ServiceType='" & strServiceType & "' Or ServiceType is Null) And CodeNo in (" & _
                                " Select CodeNo From " & GetOwner & "CD008C Where ServiceCode=" & gilServiceCode.GetCodeNo & "And (ServiceType='" & _
                                strServiceType & "' Or ServiceType is Null))", True

    gilServiceContent.SetCodeNo ""
    gilServiceContent.Query_Description
    gilServiceContent.ListIndex = 1
    Exit Sub
chkErr:
  ErrSub Me.Name, "gilServiceCode_Change"
End Sub

