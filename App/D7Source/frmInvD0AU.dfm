object frmInvD0A: TfrmInvD0A
  Left = 280
  Top = 187
  BorderStyle = bsDialog
  Caption = 'frmInvD0A'
  ClientHeight = 389
  ClientWidth = 600
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poMainFormCenter
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 600
    Height = 32
    Align = alTop
    Color = clSkyBlue
    TabOrder = 0
    object Label5: TLabel
      Left = 10
      Top = 8
      Width = 170
      Height = 16
      Caption = #21463#36104#21934#20301#20195#30908#36039#26009#32173#35703
      Font.Charset = ANSI_CHARSET
      Font.Color = clBlue
      Font.Height = -16
      Font.Name = #26032#32048#26126#39636
      Font.Style = [fsBold]
      ParentFont = False
    end
  end
  object Panel4: TPanel
    Left = 0
    Top = 32
    Width = 600
    Height = 35
    Align = alTop
    TabOrder = 1
    object btnAppend: TBitBtn
      Left = 9
      Top = 4
      Width = 71
      Height = 26
      Caption = #26032#22686
      TabOrder = 0
      OnClick = btnAppendClick
      Glyph.Data = {
        76010000424D7601000000000000760000002800000020000000100000000100
        04000000000000010000130B0000130B00001000000000000000000000000000
        800000800000008080008000000080008000808000007F7F7F00BFBFBF000000
        FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00333333333333
        33333333FF33333333FF333993333333300033377F3333333777333993333333
        300033F77FFF3333377739999993333333333777777F3333333F399999933333
        33003777777333333377333993333333330033377F3333333377333993333333
        3333333773333333333F333333333333330033333333F33333773333333C3333
        330033333337FF3333773333333CC333333333FFFFF77FFF3FF33CCCCCCCCCC3
        993337777777777F77F33CCCCCCCCCC3993337777777777377333333333CC333
        333333333337733333FF3333333C333330003333333733333777333333333333
        3000333333333333377733333333333333333333333333333333}
      NumGlyphs = 2
    end
    object btnEdit: TBitBtn
      Left = 82
      Top = 4
      Width = 71
      Height = 26
      Caption = #20462#25913
      TabOrder = 1
      OnClick = btnEditClick
      Glyph.Data = {
        76010000424D7601000000000000760000002800000020000000100000000100
        04000000000000010000120B0000120B00001000000000000000000000000000
        800000800000008080008000000080008000808000007F7F7F00BFBFBF000000
        FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00333333000000
        000033333377777777773333330FFFFFFFF03FF3FF7FF33F3FF700300000FF0F
        00F077F777773F737737E00BFBFB0FFFFFF07773333F7F3333F7E0BFBF000FFF
        F0F077F3337773F3F737E0FBFBFBF0F00FF077F3333FF7F77F37E0BFBF00000B
        0FF077F3337777737337E0FBFBFBFBF0FFF077F33FFFFFF73337E0BF0000000F
        FFF077FF777777733FF7000BFB00B0FF00F07773FF77373377373330000B0FFF
        FFF03337777373333FF7333330B0FFFF00003333373733FF777733330B0FF00F
        0FF03333737F37737F373330B00FFFFF0F033337F77F33337F733309030FFFFF
        00333377737FFFFF773333303300000003333337337777777333}
      NumGlyphs = 2
    end
    object btnDelete: TBitBtn
      Left = 155
      Top = 4
      Width = 71
      Height = 26
      Caption = #21034#38500
      TabOrder = 2
      OnClick = btnDeleteClick
      Glyph.Data = {
        76010000424D7601000000000000760000002800000020000000100000000100
        04000000000000010000130B0000130B00001000000000000000000000000000
        800000800000008080008000000080008000808000007F7F7F00BFBFBF000000
        FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00333333333333
        333333333333333333FF33333333333330003333333333333777333333333333
        300033FFFFFF3333377739999993333333333777777F3333333F399999933333
        3300377777733333337733333333333333003333333333333377333333333333
        3333333333333333333F333333333333330033333F33333333773333C3333333
        330033337F3333333377333CC3333333333333F77FFFFFFF3FF33CCCCCCCCCC3
        993337777777777F77F33CCCCCCCCCC399333777777777737733333CC3333333
        333333377F33333333FF3333C333333330003333733333333777333333333333
        3000333333333333377733333333333333333333333333333333}
      NumGlyphs = 2
    end
    object btnCancel: TBitBtn
      Left = 228
      Top = 4
      Width = 71
      Height = 26
      Caption = #21462#28040
      TabOrder = 3
      OnClick = btnCancelClick
      Glyph.Data = {
        76010000424D7601000000000000760000002800000020000000100000000100
        04000000000000010000120B0000120B00001000000000000000000000000000
        800000800000008080008000000080008000808000007F7F7F00BFBFBF000000
        FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00500005000555
        555557777F777555F55500000000555055557777777755F75555005500055055
        555577F5777F57555555005550055555555577FF577F5FF55555500550050055
        5555577FF77577FF555555005050110555555577F757777FF555555505099910
        555555FF75777777FF555005550999910555577F5F77777775F5500505509990
        3055577F75F77777575F55005055090B030555775755777575755555555550B0
        B03055555F555757575755550555550B0B335555755555757555555555555550
        BBB35555F55555575F555550555555550BBB55575555555575F5555555555555
        50BB555555555555575F555555555555550B5555555555555575}
      NumGlyphs = 2
    end
    object btnSave: TBitBtn
      Left = 301
      Top = 4
      Width = 71
      Height = 26
      Caption = #20786#23384
      TabOrder = 4
      OnClick = btnSaveClick
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
    object btnExit: TBitBtn
      Left = 485
      Top = 3
      Width = 71
      Height = 26
      Caption = #32080#26463
      TabOrder = 5
      OnClick = btnExitClick
      Glyph.Data = {
        76010000424D7601000000000000760000002800000020000000100000000100
        04000000000000010000130B0000130B00001000000000000000000000000000
        800000800000008080008000000080008000808000007F7F7F00BFBFBF000000
        FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00333333333333
        333333333333333333333333333333333333333FFF33FF333FFF339993370733
        999333777FF37FF377733339993000399933333777F777F77733333399970799
        93333333777F7377733333333999399933333333377737773333333333990993
        3333333333737F73333333333331013333333333333777FF3333333333910193
        333333333337773FF3333333399000993333333337377737FF33333399900099
        93333333773777377FF333399930003999333337773777F777FF339993370733
        9993337773337333777333333333333333333333333333333333333333333333
        3333333333333333333333333333333333333333333333333333}
      NumGlyphs = 2
    end
  end
  object pnlEdit: TPanel
    Left = 0
    Top = 67
    Width = 600
    Height = 100
    Align = alTop
    TabOrder = 2
    object Label8: TLabel
      Left = 24
      Top = 15
      Width = 91
      Height = 13
      Caption = #21463#36104#21934#20301#20195#30908#65306
      Font.Charset = CHINESEBIG5_CHARSET
      Font.Color = clRed
      Font.Height = -13
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
      Transparent = True
    end
    object Label9: TLabel
      Left = 186
      Top = 15
      Width = 39
      Height = 13
      Caption = #21517#31281#65306
      Font.Charset = CHINESEBIG5_CHARSET
      Font.Color = clRed
      Font.Height = -13
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
      Transparent = True
    end
    object Label23: TLabel
      Left = 373
      Top = 41
      Width = 118
      Height = 13
      Caption = #26159#21542#30001#26412#31995#32113'create'#65306
      Font.Charset = CHINESEBIG5_CHARSET
      Font.Color = clTeal
      Font.Height = -13
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
      Transparent = True
    end
    object Label1: TLabel
      Left = 65
      Top = 43
      Width = 52
      Height = 13
      Caption = #21443#32771#34399#65306
      Font.Charset = CHINESEBIG5_CHARSET
      Font.Color = clTeal
      Font.Height = -13
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
      Transparent = True
    end
    object lbl1: TLabel
      Left = 187
      Top = 42
      Width = 65
      Height = 13
      Caption = #32113#19968#32232#34399#65306
      Font.Charset = CHINESEBIG5_CHARSET
      Font.Color = clTeal
      Font.Height = -13
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
      Transparent = True
    end
    object lbl2: TLabel
      Left = 65
      Top = 69
      Width = 52
      Height = 13
      Caption = #24859#24515#30908#65306
      Font.Charset = CHINESEBIG5_CHARSET
      Font.Color = clTeal
      Font.Height = -13
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
      Transparent = True
    end
    object txtItemId: TMaskEdit
      Left = 116
      Top = 10
      Width = 61
      Height = 21
      EditMask = '999;1;_'
      MaxLength = 3
      TabOrder = 0
      Text = '   '
    end
    object txtDescription: TMaskEdit
      Left = 229
      Top = 10
      Width = 320
      Height = 21
      MaxLength = 20
      TabOrder = 1
    end
    object txtIsSelfCreated: TComboBox
      Left = 496
      Top = 37
      Width = 57
      Height = 21
      ItemHeight = 13
      TabOrder = 4
      Text = #26159
      Items.Strings = (
        #26159
        #21542)
    end
    object txtRefNo: TMaskEdit
      Left = 116
      Top = 38
      Width = 61
      Height = 21
      EditMask = '999;1;_'
      MaxLength = 3
      TabOrder = 2
      Text = '   '
    end
    object txtBUSINESSID: TcxTextEdit
      Left = 252
      Top = 38
      Properties.MaxLength = 8
      TabOrder = 3
      OnExit = txtBUSINESSIDExit
      Width = 95
    end
    object txtLoveNum: TcxTextEdit
      Left = 116
      Top = 66
      Properties.MaxLength = 10
      TabOrder = 5
      OnExit = txtBUSINESSIDExit
      Width = 95
    end
  end
  object pnlGrid: TPanel
    Left = 0
    Top = 167
    Width = 600
    Height = 222
    Align = alClient
    Caption = 'pnlGrid'
    TabOrder = 3
    object DBGrid1: TDBGrid
      Left = 1
      Top = 1
      Width = 598
      Height = 203
      Align = alClient
      DataSource = DataSource1
      Options = [dgTitles, dgIndicator, dgColumnResize, dgColLines, dgRowLines, dgTabs, dgRowSelect, dgConfirmDelete, dgCancelOnExit]
      TabOrder = 0
      TitleFont.Charset = DEFAULT_CHARSET
      TitleFont.Color = clWindowText
      TitleFont.Height = -11
      TitleFont.Name = 'MS Sans Serif'
      TitleFont.Style = []
    end
    object Stt_Show: TStaticText
      Left = 1
      Top = 204
      Width = 598
      Height = 17
      Align = alBottom
      Alignment = taCenter
      BevelInner = bvLowered
      BevelKind = bkSoft
      Caption = 'Stt_Show'
      Color = clInfoBk
      ParentColor = False
      TabOrder = 1
    end
  end
  object DataSource1: TDataSource
    DataSet = adoInv041
    Left = 152
    Top = 201
  end
  object adoInv041: TADOQuery
    CacheSize = 1000
    Connection = dtmMain.InvConnection
    AfterScroll = adoInv041AfterScroll
    Parameters = <>
    Left = 190
    Top = 201
  end
  object LoadTimer: TTimer
    Enabled = False
    Interval = 300
    OnTimer = LoadTimerTimer
    Left = 224
    Top = 201
  end
end
