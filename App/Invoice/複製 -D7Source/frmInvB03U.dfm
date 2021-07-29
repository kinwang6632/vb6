object frmInvB03: TfrmInvB03
  Left = 416
  Top = 277
  BorderStyle = bsDialog
  Caption = 'frmInvB03'
  ClientHeight = 272
  ClientWidth = 349
  Color = clBtnFace
  Font.Charset = ANSI_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poMainFormCenter
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 14
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 349
    Height = 32
    Align = alTop
    Color = clSkyBlue
    TabOrder = 0
    object Label1: TLabel
      Left = 16
      Top = 7
      Width = 170
      Height = 16
      Caption = #30332#31080#36039#26009#20462#25913#25480#27402#20316#26989
      Font.Charset = ANSI_CHARSET
      Font.Color = clBlue
      Font.Height = -16
      Font.Name = #26032#32048#26126#39636
      Font.Style = [fsBold]
      ParentFont = False
    end
  end
  object Panel2: TPanel
    Left = 0
    Top = 32
    Width = 349
    Height = 240
    Align = alClient
    TabOrder = 1
    object Label3: TLabel
      Left = 12
      Top = 17
      Width = 65
      Height = 13
      Caption = #30332#31080#34399#30908#65306
      Font.Charset = CHINESEBIG5_CHARSET
      Font.Color = clTeal
      Font.Height = -13
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
      Transparent = True
    end
    object Label5: TLabel
      Left = 12
      Top = 64
      Width = 65
      Height = 13
      Caption = #30332#31080#29992#36884#65306
      Font.Charset = CHINESEBIG5_CHARSET
      Font.Color = clTeal
      Font.Height = -13
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
      Transparent = True
    end
    object Label2: TLabel
      Left = 10
      Top = 88
      Width = 91
      Height = 13
      Caption = #30332#31080#38283#31435#31278#39006#65306
      Font.Charset = CHINESEBIG5_CHARSET
      Font.Color = clTeal
      Font.Height = -13
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
      Transparent = True
    end
    object btnOK: TBitBtn
      Left = 80
      Top = 171
      Width = 81
      Height = 26
      Caption = #22519#34892
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
      TabOrder = 5
      OnClick = btnOKClick
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
      Left = 192
      Top = 171
      Width = 81
      Height = 26
      Caption = #32080#26463
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
      TabOrder = 6
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
    object medSInvID: TMaskEdit
      Left = 84
      Top = 12
      Width = 85
      Height = 22
      EditMask = 'aaaaaaaaaa;1;_'
      MaxLength = 10
      TabOrder = 0
      Text = '          '
      OnChange = medSInvIDChange
      OnExit = medSInvIDExit
    end
    object medEInvID: TMaskEdit
      Left = 184
      Top = 12
      Width = 89
      Height = 22
      EditMask = 'aaaaaaaaaa;1;_'
      MaxLength = 10
      TabOrder = 1
      Text = '          '
      OnChange = medSInvIDChange
      OnExit = medEInvIDExit
    end
    object lblMsg: TcxLabel
      Left = 13
      Top = 206
      AutoSize = False
      Caption = 'lblMsg'
      ParentColor = False
      ParentFont = False
      Properties.Alignment.Horz = taCenter
      Style.Font.Charset = ANSI_CHARSET
      Style.Font.Color = clBlue
      Style.Font.Height = -12
      Style.Font.Name = 'Tahoma'
      Style.Font.Style = [fsBold]
      Style.IsFontAssigned = True
      Height = 25
      Width = 324
    end
    object chkInvUse: TCheckBox
      Left = 84
      Top = 40
      Width = 65
      Height = 17
      Caption = #25424#36104
      Color = clBtnFace
      Font.Charset = ANSI_CHARSET
      Font.Color = clTeal
      Font.Height = -12
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentColor = False
      ParentFont = False
      TabOrder = 2
      OnClick = chkInvUseClick
    end
    object txtInvUseItem: TcxButtonEdit
      Left = 83
      Top = 60
      Properties.Buttons = <
        item
          Glyph.Data = {
            72020000424D720200000000000036000000280000000E0000000D0000000100
            1800000000003C020000C40E0000C40E00000000000000000000C8D0D4C8D0D4
            C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D40000
            80FFFFFF0000C8D0D4C8D0D4000080FFFFFFC8D0D4C8D0D4C8D0D4C8D0D4C8D0
            D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D40000C8D0D4000080000080000080FFFF
            FFC8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4000080FFFFFFC8D0D40000C8D0
            D4000080000080000080FFFFFFC8D0D4C8D0D4C8D0D4C8D0D4C8D0D4000080FF
            FFFFC8D0D4C8D0D40000C8D0D4C8D0D4000080000080000080FFFFFFC8D0D4C8
            D0D4C8D0D4000080000080FFFFFFC8D0D4C8D0D40000C8D0D4C8D0D4C8D0D400
            0080000080000080FFFFFFC8D0D4000080000080FFFFFFC8D0D4C8D0D4C8D0D4
            0000C8D0D4C8D0D4C8D0D4C8D0D4000080000080000080000080000080FFFFFF
            C8D0D4C8D0D4C8D0D4C8D0D40000C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4000080
            000080000080FFFFFFC8D0D4C8D0D4C8D0D4C8D0D4C8D0D40000C8D0D4C8D0D4
            C8D0D4C8D0D4000080000080000080000080000080FFFFFFC8D0D4C8D0D4C8D0
            D4C8D0D40000C8D0D4C8D0D4C8D0D4000080000080000080FFFFFFC8D0D40000
            80FFFFFFC8D0D4C8D0D4C8D0D4C8D0D40000C8D0D40000800000800000800000
            80FFFFFFC8D0D4C8D0D4C8D0D4000080000080FFFFFFC8D0D4C8D0D400000000
            80000080000080000080FFFFFFC8D0D4C8D0D4C8D0D4C8D0D4C8D0D400008000
            0080FFFFFFC8D0D40000000080000080FFFFFFC8D0D4C8D0D4C8D0D4C8D0D4C8
            D0D4C8D0D4C8D0D4C8D0D4000080000080FFFFFF0000}
          Kind = bkGlyph
        end
        item
          Default = True
          Kind = bkEllipsis
        end>
      Properties.ReadOnly = True
      Properties.OnButtonClick = txtInvUseItemPropertiesButtonClick
      Style.LookAndFeel.Kind = lfStandard
      Style.LookAndFeel.NativeStyle = False
      StyleDisabled.LookAndFeel.Kind = lfStandard
      StyleDisabled.LookAndFeel.NativeStyle = False
      StyleFocused.LookAndFeel.Kind = lfStandard
      StyleFocused.LookAndFeel.NativeStyle = False
      StyleHot.LookAndFeel.Kind = lfStandard
      StyleHot.LookAndFeel.NativeStyle = False
      TabOrder = 4
      Width = 227
    end
    object chkEInvoice: TCheckBox
      Left = 103
      Top = 85
      Width = 88
      Height = 17
      Caption = #38651#23376#30332#31080
      Color = clBtnFace
      Font.Charset = ANSI_CHARSET
      Font.Color = clTeal
      Font.Height = -12
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentColor = False
      ParentFont = False
      TabOrder = 3
    end
    object grpImport: TGroupBox
      Left = 7
      Top = 111
      Width = 323
      Height = 55
      Caption = #21295#20837#30332#31080#36039#26009
      Font.Charset = ANSI_CHARSET
      Font.Color = clTeal
      Font.Height = -13
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
      TabOrder = 8
      object edtImport: TEdit
        Left = 10
        Top = 21
        Width = 263
        Height = 21
        ReadOnly = True
        TabOrder = 0
      end
      object btnImport: TButton
        Left = 278
        Top = 22
        Width = 33
        Height = 21
        Caption = '....'
        TabOrder = 1
        OnClick = btnImportClick
      end
    end
  end
  object OpenDialog1: TOpenDialog
    Filter = 'Text files (*.txt)|*.txt'
    Left = 298
    Top = 118
  end
end
