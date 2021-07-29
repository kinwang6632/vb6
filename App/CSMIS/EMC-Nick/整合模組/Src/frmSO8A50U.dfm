object frmSO8A50: TfrmSO8A50
  Left = 381
  Top = 184
  Width = 334
  Height = 349
  Caption = 'frmSO8A50'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 32
    Top = 32
    Width = 73
    Height = 13
    Alignment = taRightJustify
    AutoSize = False
    Caption = #20844#21496'(SO)'#21029
    Font.Charset = ANSI_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = #32048#26126#39636
    Font.Style = []
    ParentFont = False
  end
  object lblRealDate: TLabel
    Left = 45
    Top = 234
    Width = 57
    Height = 13
    Alignment = taRightJustify
    AutoSize = False
    Caption = #23526#25910#24180#26376
    Font.Charset = ANSI_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = #32048#26126#39636
    Font.Style = []
    ParentFont = False
  end
  object lblShouldDate: TLabel
    Left = 45
    Top = 233
    Width = 57
    Height = 13
    Alignment = taRightJustify
    AutoSize = False
    Caption = #25033#25910#24180#26376
    Font.Charset = ANSI_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = #32048#26126#39636
    Font.Style = []
    ParentFont = False
  end
  object Label2: TLabel
    Left = 47
    Top = 65
    Width = 57
    Height = 13
    Alignment = taRightJustify
    AutoSize = False
    Caption = #25910#36027#38917#30446
    Font.Charset = ANSI_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = #32048#26126#39636
    Font.Style = []
    ParentFont = False
  end
  object cobComp: TComboBox
    Left = 110
    Top = 24
    Width = 131
    Height = 24
    Style = csDropDownList
    Enabled = False
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ImeName = #20013#25991' ('#32321#39636') - '#26032#27880#38899
    ItemHeight = 16
    ParentFont = False
    TabOrder = 0
  end
  object btnStartCalculate: TBitBtn
    Left = 24
    Top = 274
    Width = 75
    Height = 31
    Caption = '&S'#38283#22987#35336#31639
    TabOrder = 6
    OnClick = btnStartCalculateClick
  end
  object btnCancel: TBitBtn
    Left = 232
    Top = 274
    Width = 75
    Height = 31
    Caption = '&C'#21462#28040
    TabOrder = 8
    OnClick = btnCancelClick
  end
  object btnReset: TBitBtn
    Left = 128
    Top = 273
    Width = 75
    Height = 32
    Caption = '&R'#37325#35373
    TabOrder = 7
    OnClick = btnResetClick
  end
  inline fraChineseYMD1: TfraChineseYMD
    Left = 2
    Top = 90
    Width = 73
    Height = 35
    TabOrder = 9
    Visible = False
  end
  inline fraChineseYMD2: TfraChineseYMD
    Left = 78
    Top = 90
    Width = 73
    Height = 35
    TabOrder = 10
    Visible = False
  end
  inline fraChineseYMD3: TfraChineseYMD
    Left = 0
    Top = 126
    Width = 73
    Height = 35
    TabOrder = 11
    Visible = False
    inherited mseYMD: TMaskEdit
      Left = 2
      Top = 4
    end
  end
  inline fraChineseYMD4: TfraChineseYMD
    Left = 76
    Top = 126
    Width = 73
    Height = 35
    TabOrder = 12
    Visible = False
    inherited mseYMD: TMaskEdit
      Left = 2
      Top = 4
    end
  end
  object rgpQueryType: TRadioGroup
    Left = 40
    Top = 160
    Width = 225
    Height = 49
    Caption = '   '#26597#35426#26085#26399#26781#20214'   '
    Columns = 2
    Ctl3D = True
    ItemIndex = 0
    Items.Strings = (
      #23526#25910#26085#26399
      #25033#25910#26085#26399)
    ParentCtl3D = False
    TabOrder = 3
    OnClick = rgpQueryTypeClick
  end
  object lbxChargeItem: TListBox
    Left = 111
    Top = 64
    Width = 130
    Height = 89
    ImeName = #20013#25991' ('#32321#39636') - '#26032#27880#38899
    ItemHeight = 13
    TabOrder = 1
  end
  object btnChargeItem: TBitBtn
    Left = 244
    Top = 126
    Width = 24
    Height = 24
    Caption = '...'
    TabOrder = 2
    OnClick = btnChargeItemClick
  end
  inline fraChineseYM1: TfraChineseYM
    Left = 118
    Top = 221
    Width = 100
    Height = 35
    TabOrder = 4
  end
  inline fraChineseYM2: TfraChineseYM
    Left = 118
    Top = 221
    Width = 100
    Height = 35
    TabOrder = 5
  end
  object dsrCodeCD039A: TDataSource
    DataSet = dtmMain2.cdsCodeCD039A
    Left = 4
    Top = 26
  end
end
