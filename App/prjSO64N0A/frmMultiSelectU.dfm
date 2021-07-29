object frmMultiSelect: TfrmMultiSelect
  Left = 347
  Top = 97
  ActiveControl = ListView
  BorderStyle = bsDialog
  Caption = #35079#36984
  ClientHeight = 578
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
    Top = 540
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
    Height = 540
    Align = alClient
    BevelOuter = bvNone
    BorderWidth = 10
    TabOrder = 1
    object ScrollBox1: TScrollBox
      Left = 10
      Top = 76
      Width = 495
      Height = 454
      Align = alClient
      BevelInner = bvNone
      BevelOuter = bvRaised
      BevelKind = bkFlat
      BorderStyle = bsNone
      TabOrder = 1
      object ListView: TListView
        Left = 0
        Top = 0
        Width = 493
        Height = 452
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
        ReadOnly = True
        RowSelect = True
        ParentFont = False
        ShowWorkAreas = True
        TabOrder = 0
        ViewStyle = vsReport
        OnChange = ListViewChange
        OnDblClick = ListViewDblClick
      end
    end
    object GroupBox1: TGroupBox
      Left = 10
      Top = 10
      Width = 495
      Height = 66
      Align = alTop
      Caption = #24555#36895#26597#35426
      Ctl3D = True
      ParentCtl3D = False
      TabOrder = 0
      object rdbCodeNo: TcxRadioButton
        Left = 16
        Top = 28
        Width = 49
        Height = 17
        Caption = #20195#30908
        TabOrder = 0
      end
      object rdbDescription: TcxRadioButton
        Left = 72
        Top = 28
        Width = 49
        Height = 17
        Caption = #21517#31281
        Checked = True
        TabOrder = 1
        TabStop = True
      end
      object edtText: TcxTextEdit
        Left = 120
        Top = 26
        Properties.OnChange = edtTextPropertiesChange
        TabOrder = 2
        Width = 161
      end
      object chklShowSelected: TcxCheckBox
        Left = 284
        Top = 26
        Caption = #21482#39023#31034#36984#21462#37096#20998
        Properties.OnChange = chklShowSelectedPropertiesChange
        TabOrder = 3
        Width = 113
      end
      object btnNext: TcxButton
        Left = 396
        Top = 16
        Width = 41
        Height = 41
        Enabled = False
        TabOrder = 4
        OnClick = btnNextClick
        Glyph.Data = {
          520C0000424D520C0000000000003600000028000000210000001F0000000100
          1800000000001C0C0000C40E0000C40E00000000000000000000E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E000E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E000E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E000E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0000000E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0000000000000000000E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E000E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0000000E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0000000000000E0E0E0E0E0E0000000E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E000E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0000000E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E00000
          00000000000000000000000000000000000000000000000000000000000000E0
          E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E000E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E0000000E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0000000E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E000E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0000000E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0E0000000000000000000000000E0E0E00000000000000000000000
          00000000E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E000E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0000000E0E0E0E0E0E00000
          00E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0000000E0E0E0E0E0E0000000E0E0E0E0
          E0E0E0E0E0000000E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E000E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E000000000000000
          0000E0E0E0E0E0E0E0E0E0E0E0E0000000000000000000000000E0E0E0000000
          000000000000000000000000E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E000E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0000000
          E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0000000E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E000E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0000000E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E000000000000000000000
          0000000000000000000000000000000000000000000000E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E000E0E0E0E0E0E0E0E0E0E0E0E000000000000000000000
          0000000000000000000000000000000000000000000000E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0000000E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E000E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0000000E0E0E0E0E0E0E0E0
          E0E0E0E0000000000000000000000000000000000000000000E0E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E000E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E000E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E000E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E00000000000000000000000
          00000000E0E0E0E0E0E0E0E0E0E0E0E0E0E0E000000000000000000000000000
          0000E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E000E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0000000FFFFFF00
          0000000000000000E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0000000FFFFFF000000
          000000000000E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E000E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0000000
          FFFFFF000000000000000000E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0000000FFFF
          FF000000000000000000E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E000E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0000000000000000000000000000000000000000000E0E0E000000000000000
          0000000000000000000000000000E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E000E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0000000000000FFFFFF000000000000000000000000000000000000
          FFFFFF000000000000000000000000000000E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E000E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0000000000000FFFFFF000000000000000000C0C0C00000
          00000000FFFFFF000000000000000000000000000000E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E000E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E0000000000000FFFFFF000000000000000000C0
          C0C0000000000000FFFFFF000000000000000000000000000000E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E000E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0000000000000000000000000
          000000000000000000000000000000000000000000000000000000E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E000E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0000000FFFF
          FF000000000000000000E0E0E0000000FFFFFF000000000000000000E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E000E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E000
          0000000000000000000000000000E0E0E0000000000000000000000000000000
          E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E000E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0E0000000000000000000E0E0E0E0E0E0E0E0E00000000000000000
          00E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E000E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0000000FFFFFF000000E0E0E0E0E0E0E0E0E0000000FF
          FFFF000000E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E000E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E0E0000000000000000000E0E0E0E0E0E0E0E0E0
          000000000000000000E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E000E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E000E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E000}
        LookAndFeel.Kind = lfStandard
      end
      object btnPrior: TcxButton
        Left = 440
        Top = 16
        Width = 41
        Height = 41
        Enabled = False
        TabOrder = 5
        OnClick = btnPriorClick
        Glyph.Data = {
          520C0000424D520C0000000000003600000028000000210000001F0000000100
          1800000000001C0C0000C40E0000C40E00000000000000000000E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E000E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E000E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E000E0E0E0E0E0E0E0E0E0E0E0E00000000000000000000000000000
          00000000000000000000000000000000000000E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0000000000000000000E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E000E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0000000E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0000000000000E0E0E0E0E0E0000000E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E000E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0000000E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E00000
          00000000000000000000000000000000000000000000000000000000000000E0
          E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E000E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E0000000E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0000000E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E000E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0000000E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0E0000000000000000000000000E0E0E00000000000000000000000
          00000000E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E000E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0000000E0E0E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0000000E0E0E0E0E0E0000000E0E0E0E0
          E0E0E0E0E0000000E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E000E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E000000000000000
          0000000000000000E0E0E0E0E0E0000000000000000000000000E0E0E0000000
          000000000000000000000000E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E000E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0000000
          E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0000000E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E000E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0000000E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E000000000000000000000
          0000000000000000000000000000000000000000000000E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E000E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0000000E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0000000E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E000E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0000000E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0000000000000000000000000000000000000000000E0E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E000E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E000E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E000E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E00000000000000000000000
          00000000E0E0E0E0E0E0E0E0E0E0E0E0E0E0E000000000000000000000000000
          0000E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E000E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0000000FFFFFF00
          0000000000000000E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0000000FFFFFF000000
          000000000000E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E000E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0000000
          FFFFFF000000000000000000E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0000000FFFF
          FF000000000000000000E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E000E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0000000000000000000000000000000000000000000E0E0E000000000000000
          0000000000000000000000000000E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E000E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0000000000000FFFFFF000000000000000000000000000000000000
          FFFFFF000000000000000000000000000000E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E000E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0000000000000FFFFFF000000000000000000C0C0C00000
          00000000FFFFFF000000000000000000000000000000E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E000E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E0000000000000FFFFFF000000000000000000C0
          C0C0000000000000FFFFFF000000000000000000000000000000E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E000E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0000000000000000000000000
          000000000000000000000000000000000000000000000000000000E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E000E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0000000FFFF
          FF000000000000000000E0E0E0000000FFFFFF000000000000000000E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E000E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E000
          0000000000000000000000000000E0E0E0000000000000000000000000000000
          E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E000E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0E0000000000000000000E0E0E0E0E0E0E0E0E00000000000000000
          00E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E000E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0000000FFFFFF000000E0E0E0E0E0E0E0E0E0000000FF
          FFFF000000E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E000E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E0E0000000000000000000E0E0E0E0E0E0E0E0E0
          000000000000000000E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E000E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E000E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E000}
        LookAndFeel.Kind = lfStandard
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
