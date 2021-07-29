VERSION 5.00
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Begin VB.Form frmSO8K00A 
   BorderStyle     =   1  '單線固定
   Caption         =   "Form1"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "SO8K00A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '螢幕中央
   Begin VB.Frame fraData 
      Height          =   3165
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   4515
      Begin prjGiList.GiList gilServiceCode 
         Height          =   315
         Left            =   1500
         TabIndex        =   1
         Top             =   270
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
         Left            =   1500
         TabIndex        =   4
         Top             =   660
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
      Begin prjGiList.GiList gilContentCode 
         Height          =   315
         Left            =   1500
         TabIndex        =   5
         Top             =   1050
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
         F3Corresponding =   -1  'True
         FilterStop      =   -1  'True
      End
      Begin VB.Label lblContentCode 
         Caption         =   "申告內容細項"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   150
         TabIndex        =   6
         Top             =   1110
         Width           =   1215
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
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   540
         TabIndex        =   3
         Top             =   720
         Width           =   795
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
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   540
         TabIndex        =   2
         Top             =   330
         Width           =   825
      End
   End
End
Attribute VB_Name = "frmSO8K00A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub subGil()
  On Error GoTo ChkErr
    SetLst gilServiceCode, "CodeNo", "Description", 3, 12, , , "CD008", , True      '來電分類
    SetLst gilServiceContent, "CodeNo", "Description", 3, 12, , , "CD008C", , True  '申告內容
    SetLst gilContentCode, "CodeNo", "Description", 3, 12, , , "CD008E", , True    '申告內容細項
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "subGil")
End Sub

Private Sub Form_Load()
  On Error GoTo ChkErr
    Call subGil
ChkErr:
  Call ErrSub(Me.Name, "Form_Load")
End Sub



Private Sub gilServiceCode_Change()
  On Error GoTo ChkErr
    
    If Len(Trim(gilServiceCode.GetCodeNo)) > 0 Then

        SetgiList gilServiceContent, "CodeNo", "Description", "CD008A", , , , , , , " Where CodeNo in (Select CodeNo From " & GetOwner & "CD008C " & _
                                                                " Where ServiceCode=" & gilServiceCode.GetCodeNo & _
                                                                " )", True
    End If
                                                                    

    If gilServiceContent.GetCodeNo <> "" Then
        gilContentCode.Filter = ""

        SetLst gilContentCode, "CodeNo", "Description", 3, 12, , , "CD008E", " WHERE " & _
                                                                   "Exists (Select * From CD008D Where CD008E.CodeNo=CD008D.CodeNo And " & _
                                                                               "CD008E.ServiceType=CD008D.ServiceType And CD008D.StopFlag<>1 or CD008E.ServiceType is Null)"
                                                                             
    Else
        gilContentCode.Filter = ""
        SetLst gilContentCode, "CodeNo", "Description", 3, 12, , , "CD008E", "Where 1=0"
    End If
  Exit Sub
ChkErr:
    ErrSub Me.Name, "gilServiceCode_Change"
End Sub

Private Sub gilServiceContent_Change()
    If gilServiceContent.GetCodeNo <> "" Then
        gilContentCode.Filter = ""

        SetLst gilContentCode, "CodeNo", "Description", 3, 12, , , "CD008E", " WHERE " & _
                                                                   "Exists (Select * From CD008D Where CD008E.CodeNo=CD008D.CodeNo And " & _
                                                                               "CD008E.ServiceType=CD008D.ServiceType And CD008D.StopFlag<>1 or CD008E.ServiceType is Null)"
                                                                             
    Else
        gilContentCode.Filter = ""
        SetLst gilContentCode, "CodeNo", "Description", 3, 12, , , "CD008E", "Where 1=0"
    End If
    'gilServiceCode_Change
End Sub
