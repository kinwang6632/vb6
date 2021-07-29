object frmLoadData: TfrmLoadData
  Left = 109
  Top = 55
  Width = 734
  Height = 536
  Caption = 'frmLoadData'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnClose = FormClose
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 726
    Height = 33
    Align = alTop
    Alignment = taLeftJustify
    Caption = 'Panel1'
    Color = clSkyBlue
    Font.Charset = ANSI_CHARSET
    Font.Color = clWhite
    Font.Height = -24
    Font.Name = #26032#32048#26126#39636
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 0
  end
  object Panel2: TPanel
    Left = 0
    Top = 33
    Width = 726
    Height = 456
    Align = alClient
    TabOrder = 1
    object lblComp: TLabel
      Left = 145
      Top = 15
      Width = 39
      Height = 13
      Alignment = taRightJustify
      Caption = #20844#21496#21029
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
    end
    object lblICCardPath: TLabel
      Left = 109
      Top = 50
      Width = 78
      Height = 13
      Alignment = taRightJustify
      Caption = #27284#26696#23384#25918#36335#24465
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
    end
    object lblChannelPath: TLabel
      Left = 109
      Top = 115
      Width = 78
      Height = 13
      Alignment = taRightJustify
      Caption = #27284#26696#23384#25918#36335#24465
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
    end
    object lblCAProductPath: TLabel
      Left = 109
      Top = 91
      Width = 78
      Height = 13
      Alignment = taRightJustify
      Caption = #27284#26696#23384#25918#36335#24465
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
    end
    object lblProdPurchasePath: TLabel
      Left = 109
      Top = 154
      Width = 78
      Height = 13
      Alignment = taRightJustify
      Caption = #27284#26696#23384#25918#36335#24465
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
    end
    object lblEntitlementPath: TLabel
      Left = 109
      Top = 195
      Width = 78
      Height = 13
      Alignment = taRightJustify
      Caption = #27284#26696#23384#25918#36335#24465
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
    end
    object lblTransferPath: TLabel
      Left = 109
      Top = 223
      Width = 78
      Height = 13
      Alignment = taRightJustify
      Caption = #27284#26696#23384#25918#36335#24465
      Font.Charset = ANSI_CHARSET
      Font.Color = clRed
      Font.Height = -13
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
    end
    object lblLogAndXmlPath: TLabel
      Left = 159
      Top = 249
      Width = 26
      Height = 13
      Alignment = taRightJustify
      Caption = #26085#26399
      Font.Charset = ANSI_CHARSET
      Font.Color = clBlue
      Font.Height = -13
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
    end
    object Label1: TLabel
      Left = 277
      Top = 249
      Width = 8
      Height = 16
      Alignment = taRightJustify
      Caption = '~'
      Font.Charset = ANSI_CHARSET
      Font.Color = clBlack
      Font.Height = -16
      Font.Name = #32048#26126#39636
      Font.Style = []
      ParentFont = False
    end
    object medICCardPath: TEdit
      Left = 200
      Top = 45
      Width = 321
      Height = 21
      ReadOnly = True
      TabOrder = 2
    end
    object cobCompCode: TComboBox
      Left = 200
      Top = 10
      Width = 129
      Height = 21
      ItemHeight = 13
      TabOrder = 0
      Text = 'cobCompCode'
    end
    object btnLoad: TBitBtn
      Left = 18
      Top = 204
      Width = 80
      Height = 24
      Caption = #20786#23384
      Enabled = False
      TabOrder = 17
      Visible = False
      OnClick = btnLoadClick
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
      Left = 628
      Top = 277
      Width = 80
      Height = 24
      Caption = #32080#26463
      Font.Charset = ANSI_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
      TabOrder = 21
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
    object btnICCardPath: TButton
      Left = 526
      Top = 44
      Width = 22
      Height = 22
      Caption = '...'
      TabOrder = 3
      OnClick = btnICCardPathClick
    end
    object medChannelPath: TEdit
      Left = 200
      Top = 110
      Width = 321
      Height = 21
      ReadOnly = True
      TabOrder = 7
    end
    object btnChannelPath: TButton
      Left = 526
      Top = 109
      Width = 22
      Height = 22
      Caption = '...'
      TabOrder = 8
      OnClick = btnChannelPathClick
    end
    object btnXml: TBitBtn
      Left = 507
      Top = 277
      Width = 106
      Height = 24
      Caption = #26032#22686
      TabOrder = 20
      OnClick = btnXmlClick
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
    object medCAProductPath: TEdit
      Left = 200
      Top = 86
      Width = 321
      Height = 21
      ReadOnly = True
      TabOrder = 5
    end
    object btnCAProductPath: TButton
      Left = 526
      Top = 85
      Width = 22
      Height = 22
      Caption = '...'
      TabOrder = 6
      OnClick = btnCAProductPathClick
    end
    object medProdPurchasePath: TEdit
      Left = 200
      Top = 149
      Width = 321
      Height = 21
      ReadOnly = True
      TabOrder = 10
    end
    object btnProdPurchasePath: TButton
      Left = 526
      Top = 148
      Width = 22
      Height = 22
      Caption = '...'
      TabOrder = 11
      OnClick = btnProdPurchasePathClick
    end
    object medEntitlementPath: TEdit
      Left = 200
      Top = 190
      Width = 321
      Height = 21
      ReadOnly = True
      TabOrder = 13
    end
    object btnEntitlementPath: TButton
      Left = 526
      Top = 189
      Width = 22
      Height = 22
      Caption = '...'
      TabOrder = 14
      OnClick = btnEntitlementPathClick
    end
    object chkICCard: TCheckBox
      Left = 64
      Top = 29
      Width = 97
      Height = 17
      Caption = 'chkICCard'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlue
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      TabOrder = 1
    end
    object chkCAProduct: TCheckBox
      Left = 64
      Top = 68
      Width = 97
      Height = 17
      Caption = 'chkCAProduct'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlue
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      TabOrder = 4
    end
    object chkProdPurchase: TCheckBox
      Left = 64
      Top = 131
      Width = 97
      Height = 17
      Caption = 'chkProdPurchase'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlue
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      TabOrder = 9
    end
    object chkEntitlement: TCheckBox
      Left = 64
      Top = 172
      Width = 97
      Height = 17
      Caption = 'chkEntitlement'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlue
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      TabOrder = 12
    end
    object memErrorLog: TMemo
      Left = 1
      Top = 325
      Width = 724
      Height = 130
      Align = alBottom
      Color = clInfoBk
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clRed
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      ReadOnly = True
      ScrollBars = ssVertical
      TabOrder = 22
    end
    object btnTransfer: TBitBtn
      Left = 9
      Top = 277
      Width = 140
      Height = 24
      Caption = #20786#23384
      TabOrder = 18
      OnClick = btnTransferClick
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
    object medTransferPath: TEdit
      Left = 200
      Top = 218
      Width = 321
      Height = 21
      TabOrder = 15
    end
    object btnTransferPath: TButton
      Left = 526
      Top = 217
      Width = 22
      Height = 22
      Caption = '...'
      TabOrder = 16
      OnClick = btnTransferPathClick
    end
    object btnTransferSMS: TBitBtn
      Left = 159
      Top = 277
      Width = 160
      Height = 24
      Caption = #20462#25913
      TabOrder = 19
      OnClick = btnTransferSMSClick
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
    object btnTransferSMSSO005: TBitBtn
      Left = 335
      Top = 277
      Width = 160
      Height = 24
      Caption = #20462#25913
      TabOrder = 23
      OnClick = btnTransferSMSSO005Click
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
    inline fraYMD3: TfraYMD
      Left = 200
      Top = 239
      Width = 73
      Height = 36
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlue
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 24
      OnExit = fraYMD3Exit
    end
    inline fraYMD4: TfraYMD
      Left = 290
      Top = 240
      Width = 71
      Height = 32
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlue
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 25
    end
  end
  object stbStatus: TStatusBar
    Left = 0
    Top = 489
    Width = 726
    Height = 20
    Panels = <
      item
        Alignment = taRightJustify
        Width = 100
      end
      item
        Alignment = taRightJustify
        Width = 200
      end
      item
        Alignment = taRightJustify
        Width = 50
      end>
    SimplePanel = False
  end
  object OpenDialog1: TOpenDialog
    Left = 272
    Top = 8
  end
end
