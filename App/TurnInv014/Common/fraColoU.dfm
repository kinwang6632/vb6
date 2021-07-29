object fraColo: TfraColo
  Left = 0
  Top = 0
  Width = 693
  Height = 81
  TabOrder = 0
  object lblPercent: TLabel
    Left = 420
    Top = 17
    Width = 42
    Height = 16
    Caption = '百分比'
    FocusControl = dbePercent
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
  end
  object lblAmount: TLabel
    Left = 543
    Top = 49
    Width = 28
    Height = 16
    Caption = '金額'
    FocusControl = dbeAmount
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
  end
  object lblUnitPrice: TLabel
    Left = 216
    Top = 49
    Width = 28
    Height = 16
    Caption = '單價'
    FocusControl = dbeUnitPrice
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
  end
  object lblUnit: TLabel
    Left = 424
    Top = 49
    Width = 28
    Height = 16
    Caption = '單位'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
  end
  object lblCoinID: TLabel
    Left = 15
    Top = 49
    Width = 28
    Height = 16
    Caption = '幣別'
    FocusControl = dblFTCoinName
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
  end
  object lblQuantity: TLabel
    Left = 318
    Top = 49
    Width = 28
    Height = 16
    Caption = '數量'
    FocusControl = dbeQuantity
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
  end
  object dblFromTo: TDBLookupComboBox
    Left = 254
    Top = 12
    Width = 150
    Height = 24
    DataField = 'sFromTo'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ImeName = '新注音'
    KeyField = 'sCompID'
    ListField = 'sCompID;sCompEName'
    ListSource = dsrComp
    ParentFont = False
    TabOrder = 2
    TabStop = False
  end
  object dbcCoload: TDBCheckBox
    Left = 34
    Top = 16
    Width = 57
    Height = 17
    Caption = 'Colo'
    DataField = 'bCoload'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    TabOrder = 0
    ValueChecked = 'True'
    ValueUnchecked = 'False'
  end
  object edbrInOut: TEnumDBRadioGroup
    Left = 98
    Top = 1
    Width = 150
    Height = 40
    Columns = 2
    DataField = 'nInOut'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    Items.Strings = (
      'In'
      'Out')
    ParentFont = False
    TabOrder = 1
    Values.Strings = (
      '0'
      '1')
  end
  object dbcIsSplit: TDBCheckBox
    Left = 537
    Top = 18
    Width = 58
    Height = 17
    Caption = '拆帳'
    DataField = 'bIsSplit'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    TabOrder = 4
    ValueChecked = 'True'
    ValueUnchecked = 'False'
  end
  object dbePercent: TDBEdit
    Left = 468
    Top = 12
    Width = 60
    Height = 24
    DataField = 'fPercent'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ImeName = '新注音'
    ParentFont = False
    TabOrder = 3
  end
  object dbeAmount: TDBEdit
    Left = 578
    Top = 44
    Width = 94
    Height = 24
    DataField = 'fFTAmount'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ImeName = '新注音'
    ParentFont = False
    TabOrder = 8
    OnEnter = dbeAmountEnter
  end
  object dbeUnitPrice: TDBEdit
    Left = 249
    Top = 44
    Width = 57
    Height = 24
    DataField = 'fFTUPrice'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ImeName = '新注音'
    ParentFont = False
    TabOrder = 6
  end
  object dblFTCoinName: TDBLookupComboBox
    Left = 49
    Top = 44
    Width = 149
    Height = 24
    DataField = 'sFTCoinID'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ImeName = '新注音'
    KeyField = 'sCoinID'
    ListField = 'sCoinID;sCoinEName'
    ParentFont = False
    TabOrder = 5
  end
  object dbeQuantity: TDBEdit
    Left = 352
    Top = 44
    Width = 57
    Height = 24
    DataField = 'fFTQuantity'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ImeName = '中文 (繁體) - 新注音'
    ParentFont = False
    TabOrder = 7
  end
  object BitBtn1: TBitBtn
    Left = 595
    Top = 16
    Width = 75
    Height = 25
    Caption = '儲存'
    TabOrder = 9
    Glyph.Data = {
      76010000424D7601000000000000760000002800000020000000100000000100
      04000000000000010000120B0000120B00001000000000000000000000000000
      800000800000008080008000000080008000808000007F7F7F00BFBFBF000000
      FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00555555555555
      555555555555555555555555555555555555555555FF55555555555559055555
      55555555577FF5555555555599905555555555557777F5555555555599905555
      555555557777FF5555555559999905555555555777777F555555559999990555
      5555557777777FF5555557990599905555555777757777F55555790555599055
      55557775555777FF5555555555599905555555555557777F5555555555559905
      555555555555777FF5555555555559905555555555555777FF55555555555579
      05555555555555777FF5555555555557905555555555555777FF555555555555
      5990555555555555577755555555555555555555555555555555}
    NumGlyphs = 2
  end
  object cobFTUnit: TComboBox
    Left = 453
    Top = 44
    Width = 87
    Height = 24
    Style = csDropDownList
    DropDownCount = 4
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ImeName = '新注音'
    ItemHeight = 16
    ParentFont = False
    TabOrder = 10
    Items.Strings = (
      'CBM'
      'KGS'
      '整櫃'
      '')
  end
  object dsrComp: TDataSource
    Left = 533
    Top = 5
  end
  object dsrCoin: TDataSource
    Left = 557
    Top = 5
  end
end
