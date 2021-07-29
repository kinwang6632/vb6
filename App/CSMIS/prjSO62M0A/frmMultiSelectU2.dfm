object frmMultiSelect1: TfrmMultiSelect1
  Left = 346
  Top = 149
  ActiveControl = ListView
  BorderStyle = bsDialog
  Caption = #35079#36984
  ClientHeight = 513
  ClientWidth = 515
  Color = clBtnFace
  Font.Charset = ANSI_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 14
  object Panel1: TPanel
    Left = 0
    Top = 475
    Width = 515
    Height = 38
    Align = alBottom
    BevelOuter = bvNone
    TabOrder = 0
    object Button1: TButton
      Left = 25
      Top = 2
      Width = 105
      Height = 29
      Action = actConfirm
      TabOrder = 0
    end
    object Button2: TButton
      Left = 147
      Top = 2
      Width = 105
      Height = 29
      Action = actSelectAll
      TabOrder = 1
    end
    object Button3: TButton
      Left = 269
      Top = 2
      Width = 105
      Height = 29
      Action = actClear
      TabOrder = 2
    end
    object Button4: TButton
      Left = 390
      Top = 2
      Width = 105
      Height = 29
      Action = actCancel
      TabOrder = 3
    end
  end
  object Panel2: TPanel
    Left = 0
    Top = 0
    Width = 515
    Height = 475
    Align = alClient
    BevelOuter = bvNone
    BorderWidth = 10
    TabOrder = 1
    object ScrollBox1: TScrollBox
      Left = 10
      Top = 10
      Width = 495
      Height = 455
      Align = alClient
      BevelInner = bvNone
      BevelOuter = bvRaised
      BevelKind = bkFlat
      BorderStyle = bsNone
      TabOrder = 0
      object ListView: TListView
        Left = 0
        Top = 0
        Width = 493
        Height = 453
        Cursor = crHandPoint
        Align = alClient
        BorderStyle = bsNone
        Checkboxes = True
        Columns = <>
        ColumnClick = False
        Font.Charset = ANSI_CHARSET
        Font.Color = clBlue
        Font.Height = -12
        Font.Name = 'Tahoma'
        Font.Style = []
        GridLines = True
        HideSelection = False
        HotTrack = True
        MultiSelect = True
        ReadOnly = True
        RowSelect = True
        ParentFont = False
        ShowWorkAreas = True
        TabOrder = 0
        ViewStyle = vsReport
        OnMouseDown = ListViewMouseDown
        OnSelectItem = ListViewSelectItem
      end
    end
  end
  object Timer1: TTimer
    Enabled = False
    Interval = 250
    OnTimer = Timer1Timer
    Left = 86
    Top = 366
  end
  object ActionList1: TActionList
    Left = 50
    Top = 366
    object actConfirm: TAction
      Caption = 'F2.'#30906#23450
      ShortCut = 113
      OnExecute = actConfirmExecute
    end
    object actSelectAll: TAction
      Caption = #20840#36984'(&A)'
      ShortCut = 32833
      OnExecute = actSelectAllExecute
    end
    object actClear: TAction
      Caption = #28165#38500'(&C)'
      ShortCut = 32835
      OnExecute = actClearExecute
    end
    object actCancel: TAction
      Caption = #21462#28040'(&X)'
      ShortCut = 32856
      OnExecute = actCancelExecute
    end
  end
end
