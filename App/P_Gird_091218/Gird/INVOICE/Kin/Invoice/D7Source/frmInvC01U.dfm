object frmInvC01: TfrmInvC01
  Left = 331
  Top = 184
  BorderStyle = bsDialog
  Caption = 'frmInvC01'
  ClientHeight = 404
  ClientWidth = 574
  Color = clBtnFace
  Font.Charset = ANSI_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poMainFormCenter
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 14
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 574
    Height = 32
    Align = alTop
    Color = clSkyBlue
    TabOrder = 0
    object Label1: TLabel
      Left = 16
      Top = 8
      Width = 119
      Height = 16
      Caption = #25910#36027#26126#32048#32113#35336#34920
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
    Width = 574
    Height = 372
    Align = alClient
    TabOrder = 1
    object Panel3: TPanel
      Left = 1
      Top = 1
      Width = 572
      Height = 370
      Align = alClient
      TabOrder = 0
      object Label2: TLabel
        Left = 72
        Top = 14
        Width = 60
        Height = 14
        Caption = #30332#31080#26085#26399#65306
        Transparent = True
      end
      object Label3: TLabel
        Left = 72
        Top = 74
        Width = 60
        Height = 14
        Caption = #30332#31080#34399#30908#65306
        Transparent = True
      end
      object Label4: TLabel
        Left = 84
        Top = 192
        Width = 48
        Height = 14
        Alignment = taRightJustify
        Caption = #29151#26989#21029#65306
        Transparent = True
      end
      object Label5: TLabel
        Left = 72
        Top = 102
        Width = 60
        Height = 14
        Caption = #25910#36027#38917#30446#65306
        Transparent = True
      end
      object Label6: TLabel
        Left = 60
        Top = 163
        Width = 72
        Height = 14
        Alignment = taRightJustify
        Caption = #26159#21542#21487#20462#25913#65306
        Transparent = True
      end
      object Label7: TLabel
        Left = 261
        Top = 164
        Width = 60
        Height = 14
        Alignment = taRightJustify
        Caption = #30332#31080#26684#24335#65306
        Transparent = True
      end
      object Label8: TLabel
        Left = 273
        Top = 190
        Width = 48
        Height = 14
        Alignment = taRightJustify
        Caption = #35506#31237#21029#65306
        Transparent = True
      end
      object Label9: TLabel
        Left = 84
        Top = 219
        Width = 48
        Height = 14
        Alignment = taRightJustify
        Caption = #38283#31435#21029#65306
        Transparent = True
      end
      object Label10: TLabel
        Left = 72
        Top = 45
        Width = 60
        Height = 14
        Caption = #25910#36027#26085#26399#65306
        Transparent = True
      end
      object Label11: TLabel
        Left = 40
        Top = 247
        Width = 92
        Height = 14
        Alignment = taRightJustify
        Caption = 'Excel '#23384#25918#36335#24465#65306
        Transparent = True
      end
      object Label12: TLabel
        Left = 84
        Top = 127
        Width = 48
        Height = 14
        Caption = #26381#21209#21029#65306
        Transparent = True
      end
      inline fraEInvDate: TfraYMD
        Left = 269
        Top = 1
        Width = 111
        Height = 34
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -16
        Font.Name = #26032#32048#26126#39636
        Font.Style = []
        ParentFont = False
        TabOrder = 1
        inherited mseYMD: TMaskEdit
          Left = 2
          Width = 105
          Height = 20
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = #26032#32048#26126#39636
          ParentFont = False
        end
      end
      inline fraSInvDate: TfraYMD
        Left = 134
        Top = 2
        Width = 111
        Height = 35
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -16
        Font.Name = #26032#32048#26126#39636
        Font.Style = []
        ParentFont = False
        TabOrder = 0
        OnExit = fraSInvDateExit
        inherited mseYMD: TMaskEdit
          Width = 101
          Height = 20
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = #26032#32048#26126#39636
          ParentFont = False
        end
      end
      object rgpRptType: TRadioGroup
        Left = 69
        Top = 286
        Width = 191
        Height = 68
        Caption = '  '#22577#34920#26684#24335'  '
        Columns = 2
        ItemIndex = 0
        Items.Strings = (
          #32113#35336#34920
          #26126#32048#34920
          #32113#35336'Excel'
          #26126#32048#34920'Excel')
        TabOrder = 14
        OnClick = rgpRptTypeClick
        OnEnter = rgpRptTypeClick
      end
      object medSInvID: TMaskEdit
        Left = 135
        Top = 69
        Width = 101
        Height = 22
        EditMask = 'aaaaaaaaaa;1;_'
        Font.Charset = ANSI_CHARSET
        Font.Color = clBlack
        Font.Height = -12
        Font.Name = 'Tahoma'
        Font.Style = []
        MaxLength = 10
        ParentFont = False
        TabOrder = 4
        Text = '          '
        OnExit = medSInvIDExit
      end
      object btnOK: TBitBtn
        Left = 405
        Top = 324
        Width = 73
        Height = 26
        Caption = #30906#23450
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = #26032#32048#26126#39636
        Font.Style = []
        ParentFont = False
        TabOrder = 17
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
        Left = 485
        Top = 324
        Width = 73
        Height = 26
        Caption = #32080#26463
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = #26032#32048#26126#39636
        Font.Style = []
        ParentFont = False
        TabOrder = 18
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
        Left = 271
        Top = 69
        Width = 105
        Height = 22
        EditMask = 'aaaaaaaaaa;1;_'
        Font.Charset = ANSI_CHARSET
        Font.Color = clBlack
        Font.Height = -12
        Font.Name = 'Tahoma'
        Font.Style = []
        MaxLength = 10
        ParentFont = False
        TabOrder = 5
        Text = '          '
        OnExit = medEInvIDExit
      end
      object cobBusiness: TComboBox
        Left = 135
        Top = 187
        Width = 105
        Height = 22
        Font.Charset = ANSI_CHARSET
        Font.Color = clBlack
        Font.Height = -12
        Font.Name = 'Tahoma'
        Font.Style = []
        ItemHeight = 14
        ItemIndex = 0
        ParentFont = False
        TabOrder = 8
        Text = #20840#37096
        OnExit = cobBusinessExit
        Items.Strings = (
          #20840#37096
          #29151#26989#20154
          #38750#29151#26989#20154)
      end
      inline fraEChargeDate: TfraYMD
        Left = 271
        Top = 33
        Width = 113
        Height = 32
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -16
        Font.Name = #26032#32048#26126#39636
        Font.Style = []
        ParentFont = False
        TabOrder = 3
        inherited mseYMD: TMaskEdit
          Width = 105
          Height = 20
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = #26032#32048#26126#39636
          ParentFont = False
        end
      end
      inline fraSChargeDate: TfraYMD
        Left = 135
        Top = 33
        Width = 113
        Height = 32
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -16
        Font.Name = #26032#32048#26126#39636
        Font.Style = []
        ParentFont = False
        TabOrder = 2
        OnExit = fraSChargeDateExit
        inherited mseYMD: TMaskEdit
          Width = 101
          Height = 20
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = #26032#32048#26126#39636
          ParentFont = False
        end
      end
      object cobIsPrinted: TComboBox
        Left = 135
        Top = 159
        Width = 105
        Height = 22
        Font.Charset = ANSI_CHARSET
        Font.Color = clBlack
        Font.Height = -12
        Font.Name = 'Tahoma'
        Font.Style = []
        ItemHeight = 14
        ParentFont = False
        TabOrder = 6
        OnExit = cobIsPrintedExit
        Items.Strings = (
          #19981#29702#26371
          'Yes'
          'No')
      end
      object cobInvFormat: TComboBox
        Left = 324
        Top = 158
        Width = 111
        Height = 22
        Font.Charset = ANSI_CHARSET
        Font.Color = clBlack
        Font.Height = -12
        Font.Name = 'Tahoma'
        Font.Style = []
        ItemHeight = 14
        ParentFont = False
        TabOrder = 7
        OnExit = cobInvFormatExit
        Items.Strings = (
          '0-'#20840#37096
          '1-'#38651#23376#24335
          '2-'#25163#24037#20108#32879
          '3-'#25163#24037#19977#32879)
      end
      object cobTaxType: TComboBox
        Left = 324
        Top = 187
        Width = 111
        Height = 22
        Font.Charset = ANSI_CHARSET
        Font.Color = clBlack
        Font.Height = -12
        Font.Name = 'Tahoma'
        Font.Style = []
        ItemHeight = 14
        ParentFont = False
        TabOrder = 9
        OnExit = cobTaxTypeExit
        Items.Strings = (
          '999-'#20840#37096
          '1-'#25033#31237
          '2-'#38646#31237#29575
          '3-'#20813#31237)
      end
      object cobToCreate: TComboBox
        Left = 135
        Top = 216
        Width = 161
        Height = 22
        Style = csDropDownList
        Font.Charset = ANSI_CHARSET
        Font.Color = clBlack
        Font.Height = -12
        Font.Name = 'Tahoma'
        Font.Style = []
        ItemHeight = 14
        ParentFont = False
        TabOrder = 10
        OnExit = cobToCreateExit
        Items.Strings = (
          '0-'#20840#37096
          '1-mis'#25291#27284', '#38928#38283
          '2-mis'#25291#27284', '#24460#38283
          '3-'#29694#22580#38283#31435
          '4-'#19968#33324#38283#31435)
      end
      object edtExcelPath: TEdit
        Left = 135
        Top = 244
        Width = 281
        Height = 22
        Font.Charset = ANSI_CHARSET
        Font.Color = clBlack
        Font.Height = -12
        Font.Name = 'Tahoma'
        Font.Style = []
        ParentFont = False
        ReadOnly = True
        TabOrder = 12
        Text = 'C:\'
      end
      object chkPrintMemo1: TCheckBox
        Left = 408
        Top = 275
        Width = 120
        Height = 17
        Caption = #26159#21542#21152#21360#20633#35387'1'
        TabOrder = 15
      end
      object chkPrintMemo2: TCheckBox
        Left = 408
        Top = 299
        Width = 120
        Height = 17
        Caption = #26159#21542#21152#21360#20633#35387'2'
        TabOrder = 16
      end
      object btnTransferPath: TButton
        Left = 415
        Top = 243
        Width = 22
        Height = 22
        Caption = '...'
        TabOrder = 13
        OnClick = btnTransferPathClick
      end
      object RadioGroup2: TRadioGroup
        Left = 275
        Top = 286
        Width = 125
        Height = 67
        Caption = #20844#21496#21029
        ItemIndex = 0
        Items.Strings = (
          #20491#21029#20844#21496
          #20840#37096#20844#21496)
        TabOrder = 11
      end
      object txtChargeItem: TcxButtonEdit
        Left = 135
        Top = 99
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
        Properties.OnButtonClick = txtChargeItemPropertiesButtonClick
        Style.LookAndFeel.Kind = lfStandard
        Style.LookAndFeel.NativeStyle = False
        StyleDisabled.LookAndFeel.Kind = lfStandard
        StyleDisabled.LookAndFeel.NativeStyle = False
        StyleFocused.LookAndFeel.Kind = lfStandard
        StyleFocused.LookAndFeel.NativeStyle = False
        StyleHot.LookAndFeel.Kind = lfStandard
        StyleHot.LookAndFeel.NativeStyle = False
        TabOrder = 19
        Width = 298
      end
      object cmbServiceType: TcxComboBox
        Left = 135
        Top = 125
        Properties.DropDownListStyle = lsEditFixedList
        Style.LookAndFeel.Kind = lfStandard
        Style.LookAndFeel.NativeStyle = False
        StyleDisabled.LookAndFeel.Kind = lfStandard
        StyleDisabled.LookAndFeel.NativeStyle = False
        StyleFocused.LookAndFeel.Kind = lfStandard
        StyleFocused.LookAndFeel.NativeStyle = False
        StyleHot.LookAndFeel.Kind = lfStandard
        StyleHot.LookAndFeel.NativeStyle = False
        TabOrder = 20
        OnExit = cmbServiceTypeExit
        Width = 131
      end
    end
  end
  object OpenDialog1: TOpenDialog
    Left = 12
    Top = 60
  end
end
