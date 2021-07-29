object frmSO8A40: TfrmSO8A40
  Left = 382
  Top = 193
  Width = 403
  Height = 316
  Caption = 'frmSO8A40'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -13
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 16
  object Label1: TLabel
    Left = 48
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
  object Label2: TLabel
    Left = 62
    Top = 114
    Width = 56
    Height = 13
    Alignment = taRightJustify
    AutoSize = False
    Caption = #27512#23660#26085#26399
    Font.Charset = ANSI_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = #32048#26126#39636
    Font.Style = []
    ParentFont = False
  end
  object Label3: TLabel
    Left = 61
    Top = 72
    Width = 57
    Height = 13
    Alignment = taRightJustify
    AutoSize = False
    Caption = #23526#25910#26085#26399
    Font.Charset = ANSI_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = #32048#26126#39636
    Font.Style = []
    ParentFont = False
  end
  object Label4: TLabel
    Left = 201
    Top = 69
    Width = 9
    Height = 16
    Caption = '~'
  end
  object chbShowDetail: TCheckBox
    Left = 46
    Top = 4
    Width = 339
    Height = 17
    Caption = #39023#31034#23458#25142#25910#35222#26126#32048#36039#26009'( '#24517#38920#40670#36984#27492#38917#25165#21487#36681#28858#27491#24335#36039#26009' )'
    TabOrder = 9
    Visible = False
  end
  object btnStartCalculate: TBitBtn
    Left = 32
    Top = 242
    Width = 75
    Height = 31
    Caption = '&S'#38283#22987#35336#31639
    TabOrder = 6
    OnClick = btnStartCalculateClick
  end
  object BitBtn1: TBitBtn
    Left = 256
    Top = 242
    Width = 75
    Height = 31
    Caption = '&C'#21462#28040
    TabOrder = 8
    OnClick = BitBtn1Click
  end
  object cobComp: TComboBox
    Left = 126
    Top = 24
    Width = 115
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
  inline fraChineseYMD1: TfraChineseYMD
    Left = 126
    Top = 56
    Width = 81
    Height = 41
    TabOrder = 1
    inherited mseYMD: TMaskEdit
      Height = 24
    end
  end
  inline fraChineseYMD2: TfraChineseYMD
    Left = 214
    Top = 56
    Width = 100
    Height = 35
    TabOrder = 2
    inherited mseYMD: TMaskEdit
      Height = 24
    end
  end
  inline fraChineseYM1: TfraChineseYM
    Left = 126
    Top = 101
    Width = 100
    Height = 35
    TabOrder = 3
    inherited mseYM: TMaskEdit
      Height = 24
    end
  end
  object chbProviderGroup: TCheckBox
    Left = 47
    Top = 209
    Width = 99
    Height = 17
    Caption = #20381#38971#36947#21830#20998#38913
    TabOrder = 5
  end
  object btnPrint: TBitBtn
    Left = 144
    Top = 241
    Width = 75
    Height = 32
    Caption = '&P'#21015#21360
    TabOrder = 7
    OnClick = btnPrintClick
  end
  object rgpType: TRadioGroup
    Left = 36
    Top = 152
    Width = 321
    Height = 49
    Caption = '  '#23637#29694#22577#34920'  '
    Columns = 2
    Font.Charset = ANSI_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = #26032#32048#26126#39636
    Font.Style = []
    ItemIndex = 0
    Items.Strings = (
      #38971#36947#25286#24115#26126#32048#34920
      #23458#25142#25910#35222#26126#32048#34920)
    ParentFont = False
    TabOrder = 4
  end
  object dsrCodeCD039A: TDataSource
    DataSet = dtmMain2.cdsCodeCD039A
    Left = 16
    Top = 26
  end
  object DataSource1: TDataSource
    DataSet = dtmMain2.cdsSO114SubTotal
    Left = 280
    Top = 24
  end
  object DataSource2: TDataSource
    DataSet = dtmMain2.cdsSO114Rpt
    Left = 280
    Top = 96
  end
end
