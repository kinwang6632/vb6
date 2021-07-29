object frmInvA05: TfrmInvA05
  Left = 278
  Top = 117
  BorderStyle = bsDialog
  Caption = 'frmInvA05'
  ClientHeight = 435
  ClientWidth = 368
  Color = clBtnFace
  Font.Charset = ANSI_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 14
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 368
    Height = 32
    Align = alTop
    Color = clSkyBlue
    TabOrder = 0
    object Label1: TLabel
      Left = 16
      Top = 7
      Width = 153
      Height = 16
      Caption = #38651#23376#35336#31639#27231#30332#31080#22871#21360
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
    Width = 368
    Height = 403
    Align = alClient
    TabOrder = 1
    object Label2: TLabel
      Left = 28
      Top = 46
      Width = 65
      Height = 13
      Caption = #30332#31080#26085#26399#65306
      Font.Charset = CHINESEBIG5_CHARSET
      Font.Color = clTeal
      Font.Height = -13
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
      Transparent = True
    end
    object Label3: TLabel
      Left = 28
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
    object Label4: TLabel
      Left = 42
      Top = 74
      Width = 52
      Height = 13
      Caption = #29151#26989#21029#65306
      Font.Charset = CHINESEBIG5_CHARSET
      Font.Color = clTeal
      Font.Height = -13
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
      Transparent = True
    end
    object Label5: TLabel
      Left = 29
      Top = 99
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
    object Label6: TLabel
      Left = 29
      Top = 125
      Width = 65
      Height = 13
      Caption = #21463#36104#21934#20301#65306
      Font.Charset = CHINESEBIG5_CHARSET
      Font.Color = clTeal
      Font.Height = -13
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
      Transparent = True
    end
    inline fraEInvDate: TfraYMD
      Left = 240
      Top = 25
      Width = 113
      Height = 49
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -16
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
      TabOrder = 3
      inherited mseYMD: TMaskEdit
        Top = 16
        Width = 103
        Font.Charset = CHINESEBIG5_CHARSET
        Font.Color = clBlack
        Font.Height = -13
        Font.Name = #26032#32048#26126#39636
        ParentFont = False
      end
    end
    inline fraSInvDate: TfraYMD
      Left = 116
      Top = 25
      Width = 113
      Height = 49
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -16
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
      TabOrder = 2
      OnExit = fraSInvDateExit
      inherited mseYMD: TMaskEdit
        Left = 1
        Top = 16
        Width = 101
        Font.Charset = CHINESEBIG5_CHARSET
        Font.Color = clBlack
        Font.Height = -13
        Font.Name = #26032#32048#26126#39636
        ParentFont = False
      end
    end
    object rgpOrder: TRadioGroup
      Left = 20
      Top = 151
      Width = 329
      Height = 49
      Caption = '  '#25490#24207#26041#24335'  '
      Columns = 3
      Font.Charset = CHINESEBIG5_CHARSET
      Font.Color = clTeal
      Font.Height = -13
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ItemIndex = 0
      Items.Strings = (
        #30332#31080#34399#30908
        #30332#31080#26085#26399
        #23458#25142#32232#34399)
      ParentFont = False
      TabOrder = 5
    end
    object medSInvID: TMaskEdit
      Left = 116
      Top = 12
      Width = 101
      Height = 21
      EditMask = 'aaaaaaaaaa;1;_'
      Font.Charset = CHINESEBIG5_CHARSET
      Font.Color = clBlack
      Font.Height = -13
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      MaxLength = 10
      ParentFont = False
      TabOrder = 0
      Text = '          '
      OnExit = medSInvIDExit
    end
    object rgpTransferType: TRadioGroup
      Left = 20
      Top = 203
      Width = 329
      Height = 49
      Caption = '  '#21015#21360#26041#24335'  '
      Columns = 3
      Font.Charset = CHINESEBIG5_CHARSET
      Font.Color = clTeal
      Font.Height = -13
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ItemIndex = 0
      Items.Strings = (
        #30332#31080#21015#21360
        #36681#20837#25991#23383#27284)
      ParentFont = False
      TabOrder = 6
    end
    object btnOK: TBitBtn
      Left = 66
      Top = 365
      Width = 81
      Height = 27
      Caption = #30906#23450
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
      TabOrder = 7
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
      Left = 234
      Top = 365
      Width = 81
      Height = 27
      Caption = #32080#26463
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
      TabOrder = 8
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
    object medEInvID: TMaskEdit
      Left = 240
      Top = 12
      Width = 103
      Height = 21
      EditMask = 'aaaaaaaaaa;1;_'
      Font.Charset = CHINESEBIG5_CHARSET
      Font.Color = clBlack
      Font.Height = -13
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      MaxLength = 10
      ParentFont = False
      TabOrder = 1
      Text = '          '
      OnExit = medEInvIDExit
    end
    object cobBusiness: TComboBox
      Left = 116
      Top = 68
      Width = 103
      Height = 21
      Font.Charset = CHINESEBIG5_CHARSET
      Font.Color = clBlack
      Font.Height = -13
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ItemHeight = 13
      ItemIndex = 0
      ParentFont = False
      TabOrder = 4
      Text = #20840#37096
      Items.Strings = (
        #20840#37096
        #29151#26989#20154
        #38750#29151#26989#20154)
    end
    object txtInvUseItem: TcxButtonEdit
      Left = 117
      Top = 95
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
      TabOrder = 9
      Width = 227
    end
    object chkInvUse: TCheckBox
      Left = 239
      Top = 69
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
      TabOrder = 10
      OnClick = chkInvUseClick
    end
    object grpInvKind: TGroupBox
      Left = 20
      Top = 256
      Width = 327
      Height = 47
      Caption = #30332#31080#38283#31435#31278#39006
      Font.Charset = ANSI_CHARSET
      Font.Color = clTeal
      Font.Height = -13
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
      TabOrder = 11
      object chkNormalInv: TCheckBox
        Left = 10
        Top = 22
        Width = 115
        Height = 17
        Caption = #38651#23376#35336#31639#27231#30332#31080
        Checked = True
        State = cbChecked
        TabOrder = 0
        OnClick = chkNormalInvClick
      end
      object chkEInv: TCheckBox
        Left = 182
        Top = 22
        Width = 115
        Height = 17
        Caption = #38651#23376#30332#31080
        TabOrder = 1
        OnClick = chkNormalInvClick
      end
    end
    object grpImport: TGroupBox
      Left = 20
      Top = 306
      Width = 323
      Height = 55
      Caption = #21295#20837#30332#31080#36039#26009
      Font.Charset = ANSI_CHARSET
      Font.Color = clTeal
      Font.Height = -13
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
      TabOrder = 12
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
    object txtGiftUnit: TcxButtonEdit
      Left = 117
      Top = 123
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
      Properties.OnButtonClick = txtGiftUnitPropertiesButtonClick
      Style.LookAndFeel.Kind = lfStandard
      Style.LookAndFeel.NativeStyle = False
      StyleDisabled.LookAndFeel.Kind = lfStandard
      StyleDisabled.LookAndFeel.NativeStyle = False
      StyleFocused.LookAndFeel.Kind = lfStandard
      StyleFocused.LookAndFeel.NativeStyle = False
      StyleHot.LookAndFeel.Kind = lfStandard
      StyleHot.LookAndFeel.NativeStyle = False
      TabOrder = 13
      Width = 227
    end
  end
  object OpenDialog1: TOpenDialog
    Filter = 'Text files (*.txt)|*.txt'
    Left = 338
    Top = 328
  end
end
