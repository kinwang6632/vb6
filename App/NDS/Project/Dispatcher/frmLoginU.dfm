object frmLogin: TfrmLogin
  Left = 203
  Top = 169
  Width = 379
  Height = 266
  BorderIcons = []
  Caption = #31995#32113#30331#20837
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 371
    Height = 239
    Align = alClient
    BevelInner = bvLowered
    BevelOuter = bvLowered
    TabOrder = 0
    object Label3: TLabel
      Left = 81
      Top = 114
      Width = 65
      Height = 13
      AutoSize = False
      Caption = #20351#29992#32773#20195#34399
      Font.Charset = CHINESEBIG5_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = 'Times New Roman'
      Font.Style = []
      ParentFont = False
    end
    object Label4: TLabel
      Left = 120
      Top = 152
      Width = 26
      Height = 13
      AutoSize = False
      Caption = #23494#30908
      Font.Charset = CHINESEBIG5_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = 'Times New Roman'
      Font.Style = []
      ParentFont = False
    end
    object edtUserID: TEdit
      Left = 155
      Top = 109
      Width = 121
      Height = 21
      Font.Charset = CHINESEBIG5_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = 'Times New Roman'
      Font.Style = []
      ParentFont = False
      TabOrder = 0
    end
    object edtPasswd: TEdit
      Left = 155
      Top = 149
      Width = 121
      Height = 21
      Font.Charset = CHINESEBIG5_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = 'Times New Roman'
      Font.Style = []
      ParentFont = False
      PasswordChar = '*'
      TabOrder = 1
    end
    object bbnConfirm: TBitBtn
      Left = 56
      Top = 188
      Width = 81
      Height = 25
      Caption = #30906#23450
      TabOrder = 2
      OnClick = bbnConfirmClick
      Glyph.Data = {
        76010000424D7601000000000000760000002800000020000000100000000100
        0400000000000001000000000000000000001000000010000000000000000000
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
    object bbnCancel: TBitBtn
      Left = 248
      Top = 188
      Width = 81
      Height = 25
      Caption = #21462#28040
      ModalResult = 2
      TabOrder = 3
      OnClick = bbnCancelClick
      Glyph.Data = {
        76010000424D7601000000000000760000002800000020000000100000000100
        0400000000000001000000000000000000001000000010000000000000000000
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
    object Panel2: TPanel
      Left = 2
      Top = 2
      Width = 367
      Height = 41
      Align = alTop
      BevelOuter = bvNone
      Caption = ' '#27489#36814#20351#29992' NDS Dispatcher'
      Color = clCream
      Font.Charset = ANSI_CHARSET
      Font.Color = clNavy
      Font.Height = -19
      Font.Name = 'Arial Black'
      Font.Style = [fsBold]
      ParentFont = False
      TabOrder = 4
    end
    object Panel3: TPanel
      Left = 2
      Top = 43
      Width = 367
      Height = 41
      Align = alTop
      BevelOuter = bvNone
      Caption = #38283#21338#31185#25216'('#32929')'#20844#21496
      Color = clMoneyGreen
      Font.Charset = CHINESEBIG5_CHARSET
      Font.Color = clBlue
      Font.Height = -16
      Font.Name = 'Times New Roman'
      Font.Style = [fsBold]
      ParentFont = False
      TabOrder = 5
    end
    object pnlPrgName: TPanel
      Left = 2
      Top = 84
      Width = 367
      Height = 5
      Align = alTop
      BevelOuter = bvNone
      Color = clTeal
      Font.Charset = CHINESEBIG5_CHARSET
      Font.Color = clTeal
      Font.Height = -19
      Font.Name = 'Times New Roman'
      Font.Style = []
      ParentFont = False
      TabOrder = 6
    end
    object btnReSet: TBitBtn
      Left = 152
      Top = 188
      Width = 81
      Height = 25
      Caption = #37325#35373
      ModalResult = 2
      TabOrder = 7
      OnClick = btnReSetClick
      Glyph.Data = {
        76010000424D7601000000000000760000002800000020000000100000000100
        0400000000000001000000000000000000001000000010000000000000000000
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
end
