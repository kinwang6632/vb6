object frmInvB01: TfrmInvB01
  Left = 330
  Top = 136
  Width = 580
  Height = 561
  Caption = 'frmInvB01'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
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
    Width = 572
    Height = 32
    Align = alTop
    Color = clSkyBlue
    TabOrder = 0
    object Label5: TLabel
      Left = 18
      Top = 8
      Width = 102
      Height = 16
      Caption = #30332#31080#21443#25976#35373#23450
      Font.Charset = ANSI_CHARSET
      Font.Color = clBlue
      Font.Height = -16
      Font.Name = #26032#32048#26126#39636
      Font.Style = [fsBold]
      ParentFont = False
    end
  end
  object pnlEdit: TPanel
    Left = 0
    Top = 74
    Width = 572
    Height = 196
    Align = alTop
    BevelOuter = bvNone
    TabOrder = 1
    object MainPage: TcxPageControl
      Left = 0
      Top = 0
      Width = 572
      Height = 196
      ActivePage = cxTabSheet1
      Align = alClient
      TabOrder = 0
      ClientRectBottom = 192
      ClientRectLeft = 2
      ClientRectRight = 568
      ClientRectTop = 23
      object cxTabSheet1: TcxTabSheet
        Caption = '  '#19968#33324#21443#25976'  '
        ImageIndex = 0
        object Label8: TLabel
          Left = 143
          Top = 13
          Width = 60
          Height = 14
          Caption = #31995#32113#20195#30908#65306
          Transparent = True
        end
        object Label1: TLabel
          Left = 253
          Top = 13
          Width = 60
          Height = 14
          Caption = #20844#21496#21517#31281#65306
          Transparent = True
        end
        object Label2: TLabel
          Left = 435
          Top = 13
          Width = 60
          Height = 14
          Caption = #20844#21496#22411#21029#65306
          Transparent = True
        end
        object Label3: TLabel
          Left = 32
          Top = 42
          Width = 125
          Height = 14
          Caption = #37559#36008#36864#22238'/'#25240#35731#36039#26009#27284#65306
          Transparent = True
        end
        object Label4: TLabel
          Left = 37
          Top = 64
          Width = 120
          Height = 14
          Caption = #38928#38283#30332#31080#25291#27284#36039#26009#27284#65306
          Transparent = True
        end
        object Label6: TLabel
          Left = 37
          Top = 86
          Width = 120
          Height = 14
          Caption = #24460#38283#30332#31080#25291#27284#36039#26009#27284#65306
          Transparent = True
        end
        object Label7: TLabel
          Left = 37
          Top = 109
          Width = 120
          Height = 14
          Caption = #23186#39636#30003#22577#27284#36664#20986#36335#24465#65306
          Transparent = True
        end
        object Label9: TLabel
          Left = 73
          Top = 129
          Width = 84
          Height = 14
          Caption = #27284#26696#20633#20221#36335#24465#65306
          Transparent = True
        end
        object Label10: TLabel
          Left = 23
          Top = 13
          Width = 60
          Height = 14
          Caption = #20844#21496#20195#34399#65306
          Transparent = True
        end
        object medSysID: TMaskEdit
          Left = 206
          Top = 10
          Width = 38
          Height = 22
          EditMask = 'aaaa;1;_'
          MaxLength = 4
          TabOrder = 0
          Text = '    '
        end
        object medCompName: TMaskEdit
          Left = 320
          Top = 10
          Width = 105
          Height = 22
          EditMask = 'aaaaaaaaaaaaaaaaaaaaaaaaaaaaaa;1; '
          MaxLength = 30
          TabOrder = 1
          Text = '                              '
        end
        object cobCompType: TComboBox
          Left = 505
          Top = 10
          Width = 43
          Height = 22
          ItemHeight = 14
          ItemIndex = 0
          TabOrder = 2
          Text = '1'
          Items.Strings = (
            '1'
            '2'
            '3')
        end
        object btnRejectionInfoFile: TButton
          Left = 524
          Top = 36
          Width = 19
          Height = 22
          Caption = '...'
          TabOrder = 3
          OnClick = btnRejectionInfoFileClick
        end
        object medRejectionInfoFile: TMaskEdit
          Left = 169
          Top = 38
          Width = 355
          Height = 22
          ReadOnly = True
          TabOrder = 4
        end
        object btnPInvoiceFilePath: TButton
          Left = 524
          Top = 59
          Width = 19
          Height = 21
          Caption = '...'
          TabOrder = 5
          OnClick = btnPInvoiceFilePathClick
        end
        object medPInvoiceFilePath: TMaskEdit
          Left = 169
          Top = 60
          Width = 355
          Height = 22
          ReadOnly = True
          TabOrder = 6
        end
        object btnInvoiceFilePath: TButton
          Left = 524
          Top = 81
          Width = 19
          Height = 21
          Caption = '...'
          TabOrder = 7
          OnClick = btnInvoiceFilePathClick
        end
        object medInvoiceFilePath: TMaskEdit
          Left = 169
          Top = 82
          Width = 355
          Height = 22
          ReadOnly = True
          TabOrder = 8
        end
        object btnMediaFilePath: TButton
          Left = 524
          Top = 103
          Width = 19
          Height = 21
          Caption = '...'
          TabOrder = 9
          OnClick = btnMediaFilePathClick
        end
        object medMediaFilePath: TMaskEdit
          Left = 169
          Top = 104
          Width = 355
          Height = 22
          ReadOnly = True
          TabOrder = 10
        end
        object btnBackupFilePath: TButton
          Left = 524
          Top = 125
          Width = 19
          Height = 21
          Caption = '...'
          TabOrder = 11
          OnClick = btnBackupFilePathClick
        end
        object medBackupFilePath: TMaskEdit
          Left = 169
          Top = 126
          Width = 355
          Height = 22
          ReadOnly = True
          TabOrder = 12
        end
        object medCompID: TMaskEdit
          Left = 91
          Top = 10
          Width = 47
          Height = 22
          EditMask = 'aaaaaa;1;_'
          MaxLength = 6
          TabOrder = 13
          Text = '      '
        end
      end
      object cxTabSheet2: TcxTabSheet
        Caption = ' '#30332#31080#21295#20986'/'#21015#21360#21443#25976' '
        ImageIndex = 1
        object rgpUseBaseCompName: TcxRadioGroup
          Left = 22
          Top = 16
          Caption = '  '#22577#34920#20844#21496#25260#38957#20197#30332#31080#21443#25976#35373#23450#28858#28310'  '
          ItemIndex = 0
          Properties.Columns = 2
          Properties.Items = <
            item
              Caption = #26159
              Value = 0
            end
            item
              Caption = #21542
              Value = 0
            end>
          TabOrder = 0
          Height = 52
          Width = 223
        end
        object rgpExpAddrType: TcxRadioGroup
          Left = 256
          Top = 16
          Caption = #30332#31080#21295#20986#25991#23383#27284#23492#20214#22320#22336#20381#25818
          ItemIndex = 0
          Properties.Columns = 2
          Properties.Items = <
            item
              Caption = #37109#23492#22320#22336
            end
            item
              Caption = #30332#31080#22320#22336
            end>
          TabOrder = 1
          Height = 52
          Width = 233
        end
        object rgpPrintAddr: TcxRadioGroup
          Left = 22
          Top = 86
          Caption = '  '#26159#21542#21015#21360#22320#22336' '
          ItemIndex = 0
          Properties.Columns = 2
          Properties.Items = <
            item
              Caption = #26159
            end
            item
              Caption = #21542
            end>
          TabOrder = 2
          Height = 52
          Width = 111
        end
        object rgpPrintTitle: TcxRadioGroup
          Left = 141
          Top = 86
          Caption = ' '#24375#21046#21015#21360#30332#31080#25260#38957'  '
          ItemIndex = 0
          ParentShowHint = False
          Properties.Columns = 2
          Properties.Items = <
            item
              Caption = #26159
            end
            item
              Caption = #21542
            end>
          ShowHint = False
          TabOrder = 3
          Height = 52
          Width = 156
        end
        object rgpIfPrintCheck: TcxRadioGroup
          Left = 308
          Top = 85
          Caption = '  '#21015#21360'/'#21295#20986#30332#31080#26159#21542#38656#35201#25480#27402'  '
          ItemIndex = 0
          Properties.Columns = 2
          Properties.Items = <
            item
              Caption = #26159
            end
            item
              Caption = #21542
            end>
          TabOrder = 4
          Height = 52
          Width = 179
        end
      end
      object cxTabSheet3: TcxTabSheet
        Caption = ' '#38646#31237#29575#21443#25976' '
        ImageIndex = 2
        object Label11: TLabel
          Left = 25
          Top = 19
          Width = 76
          Height = 14
          Caption = #37559#21806#20154#32291#24066#21029':'
          Transparent = True
        end
        object Label12: TLabel
          Left = 49
          Top = 48
          Width = 52
          Height = 14
          Caption = #22806#37559#26041#24335':'
          Transparent = True
        end
        object Label13: TLabel
          Left = 25
          Top = 77
          Width = 76
          Height = 14
          Caption = #36890#38364#26041#24335#35387#35352':'
          Transparent = True
        end
        object Label14: TLabel
          Left = 24
          Top = 108
          Width = 76
          Height = 14
          Caption = #20986#21475#22577#21934#39006#21029':'
          Transparent = True
        end
        object cmbCity: TcxComboBox
          Left = 108
          Top = 16
          Properties.DropDownListStyle = lsEditFixedList
          Properties.Items.Strings = (
            ''
            'A.'#21488#21271#24066
            'B.'#21488#20013#24066
            'C.'#22522#38534#24066
            'D.'#21488#21335#24066
            'E.'#39640#38596#24066
            'F.'#21488#21271#32291
            'G.'#23452#34349#32291
            'H.'#26691#22290#24066
            'I.'#22025#32681#24066
            'J.'#26032#31481#32291
            'K.'#33495#26647#32291
            'L.'#21488#20013#32291
            'M.'#21335#25237#32291
            'N.'#24432#21270#32291
            'O.'#26032#31481#24066
            'P.'#38642#26519#32291
            'Q.'#22025#32681#32291
            'R.'#21488#21335#32291
            'S.'#39640#38596#32291
            'T.'#23631#26481#32291
            'U.'#33457#34030#32291
            'V.'#21488#26481#32291
            'X.'#37329#38272#32291
            'X.'#28558#28246#32291
            'Z.'#36899#27743#32291)
          Style.StyleController = dtmMain.cxEditStyle
          TabOrder = 0
          Width = 137
        end
        object cmbExportMode: TcxComboBox
          Left = 108
          Top = 45
          Properties.DropDownListStyle = lsEditFixedList
          Properties.Items.Strings = (
            ''
            '1.'#22806#37559#36008#29289
            '2.'#33287#22806#37559#26377#38364#20043#21214#21209#65292#25110#22312#22283#20839#25552#20379#32780#22312#22283#22806#20351#29992#20043#21214#21209
            '3.'#20381#27861#35373#31435#20043#21453#31237#21830#24215#37559#21806#33287#36942#22659#25110#20986#22659#26053#23458#20043#36008#29289
            '4.'#20986#21475#21312#12289#31185#23416#24037#26989#22290#21312#12289#20445#31237#24037#24288#25110#20445#31237#20489#24235#20043#27231#22120#35373#20633#21407#26009#12289#29289#26009#12289#29123#26009#12289#21322#35069#21697
            '5.'#22283#38555#36939#36664
            '6.'#22283#38555#36939#36664#29992#20043#33337#33334#12289#33322#31354#22120#65292#21450#36960#27915#28417#33337
            '7.'#37559#21806#33287#22283#38555#36939#36664#29992#20043#33337#33334#12289#33322#31354#22120#21450#36960#27915#28417#33337#25152#20351#29992#20043#36008#29289#25110#20462#32341#21214#21209)
          Style.StyleController = dtmMain.cxEditStyle
          TabOrder = 1
          Width = 400
        end
        object cmbImportNote: TcxComboBox
          Left = 108
          Top = 74
          Properties.DropDownListStyle = lsEditFixedList
          Properties.Items.Strings = (
            #20854#23427' ('#31354#30333#20195#34920#36914#38917#24977#35657#12289#37559#36008#36864#22238#25110#25240#35731#35657#26126#21934#21450#38750#36969#29992#38646#31237#29575#20043#37559#38917#24977#35657')'
            '1.'#38750#32147#28023#38364' ('#20195#34920#36969#29992#38646#31237#29575#38750#32147#28023#38364#20986#21475#25033#38468#35657#26126#25991#20214#20043#37559#38917')'
            '2.'#32147#28023#38364' ('#20195#34920#36969#29992#38646#31237#29575#32147#28023#38364#20986#21475#20813#38468#35657#26126#25991#20214#20043#37559#38917')')
          Style.StyleController = dtmMain.cxEditStyle
          TabOrder = 2
          Width = 400
        end
        object txtEntryClass: TcxTextEdit
          Left = 108
          Top = 104
          Properties.MaxLength = 2
          Style.StyleController = dtmMain.cxEditStyle
          TabOrder = 3
          Width = 74
        end
      end
    end
  end
  object pnlGrid: TPanel
    Left = 0
    Top = 270
    Width = 572
    Height = 263
    Align = alClient
    BevelOuter = bvNone
    Caption = 'pnlGrid'
    TabOrder = 2
    object Bevel1: TBevel
      Left = 0
      Top = 0
      Width = 572
      Height = 3
      Align = alTop
      Shape = bsSpacer
    end
    object DBGrid1: TDBGrid
      Left = 0
      Top = 3
      Width = 572
      Height = 242
      Align = alClient
      DataSource = dsrInv003
      ReadOnly = True
      TabOrder = 0
      TitleFont.Charset = DEFAULT_CHARSET
      TitleFont.Color = clWindowText
      TitleFont.Height = -12
      TitleFont.Name = 'Tahoma'
      TitleFont.Style = []
    end
    object Stt_Show: TStaticText
      Left = 0
      Top = 245
      Width = 572
      Height = 18
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
  object Panel4: TPanel
    Left = 0
    Top = 32
    Width = 572
    Height = 42
    Align = alTop
    BevelOuter = bvNone
    TabOrder = 3
    object btnAppend: TcxButton
      Left = 7
      Top = 5
      Width = 80
      Height = 27
      Caption = #26032#22686
      TabOrder = 0
      OnClick = btnAppendClick
      Glyph.Data = {
        36060000424D3606000000000000360000002800000020000000100000000100
        1800000000000006000000000000000000000000000000000000FF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
        9B7C6B9D7E6D9C7E6D9C7E6D9C7E6D9C7D6D9C7D6CBFABA1007500007000006D
        00BFABA1FF00FFFF00FFFF00FFFF00FFA47878A47878A47878A47878A47878A4
        7878A47878BEA3A3000000000000000000BEA3A3FF00FFFF00FFFF00FFFF00FF
        9B7766FFFFFFFAF4E9FAF4E9FAF4E9FAF3E8FAF3E7FCF8F1007D0044DD770072
        00BDA99DFF00FFFF00FFFF00FFFF00FF857769FF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FF000000FF00FF000000BEA3A3FF00FFFF00FFFF00FFFF00FF
        A27F6FFFFFFFDDC2B5DDC2B5DCC2B5E9D6CD00870000850000810048E17B007A
        00007500007000FF00FFFF00FFFF00FFA47878FF00FFDCC1C1DCC1C1DCC1C1E9
        D4D4000000000000000000FF00FF000000000000000000FF00FFFF00FFFF00FF
        A38070FFFFFFDBC3BBDBC3BADBC2B8EBDCD5008D005EF7915AF38D53EC8648E1
        7B45DE78007800FF00FFFF00FFFF00FFA47878FF00FFDCC1C1DCC1C1DCC1C1EB
        D7D7000000FF00FFFF00FFFF00FFFF00FFFF00FF000000FF00FFFF00FFFF00FF
        A98778FFFFFFDBC7C2DBC6C1DBC4BCEBDDD7009100008D00008B005AF38D0083
        00008100007D00FF00FFFF00FFFF00FFA98888FF00FFDCC1C1DCC1C1DCC1C1EB
        D7D7000000000000000000FF00FF000000000000000000FF00FFFF00FFFF00FF
        AB897AFFFFFFDBC7C3DBC7C2DBC5BEEBDED9EBDCD6EBDBD3008F005EF7910089
        00C8B7AEFF00FFFF00FFFF00FFFF00FFA98888FF00FFDCC1C1DCC1C1DCC1C1EC
        D9D9EBD7D7EBD7D7000000FF00FF000000C2AEAEFF00FFFF00FFFF00FFFF00FF
        AF8E7FFFFFFFDCC5C0DCC5BFDBC4BDDBC2B9DBBFB4EBDBD3009300009100008D
        00BFABA1FF00FFFF00FFFF00FFFF00FFB18B8BFF00FFDCC1C1DCC1C1DCC1C1DC
        C1C1DABBBBEBD7D7000000000000000000BEA3A3FF00FFFF00FFFF00FFFF00FF
        AF8F80FFFFFFFEFEFEFEFEFEFEFDFDFEFCFBFDFBF8FEFCFAFDFCF8FDFBF7FCF6
        ED9B7C6CFF00FFFF00FFFF00FFFF00FFB18B8BFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFA47878FF00FFFF00FFFF00FFFF00FF
        AF8F80FFFFFFDFCECCDFCDCBDECAC6DEC6C0DEC4BADEC1B4DEBEADDEBEABFCF6
        EE9C7C6DFF00FFFF00FFFF00FFFF00FFB18B8BFF00FFDDCBCBDDCBCBE0C7C7DC
        C1C1DEBDBDDEBDBDDEBDBDDEBDBDFF00FFA47878FF00FFFF00FFFF00FFFF00FF
        B19080FFFFFFFFFFFFFFFFFFFFFFFFFEFEFEFEFDFBFDFBF8FCF9F4F9F4EEF0E8
        E09D7F6FFF00FFFF00FFFF00FFFF00FFB18B8BFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFA47878FF00FFFF00FFFF00FFFF00FF
        B79787FFFFFFFFFFFFFFFFFFFFFFFFFEFEFEFEFDFBFDFBF9A78270A78270A782
        70A78270FF00FFFF00FFFF00FFFF00FFB39292FF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFA47878A47878A47878A47878FF00FFFF00FFFF00FFFF00FF
        B89888FFFFFFFFFFFFFFFFFFFEFEFEFEFEFDFEFCFAFDFBF9A78270F5E2D9B18E
        7EAB9E98FF00FFFF00FFFF00FFFF00FFBD9494FF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFA47878FF00FFB18B8BAA9F9EFF00FFFF00FFFF00FFFF00FF
        B89888FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFA78270B18E7EAB9E
        98FF00FFFF00FFFF00FFFF00FFFF00FFBD9494FF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFA47878B18B8BAA9F9EFF00FFFF00FFFF00FFFF00FFFF00FF
        B89888B89888B49383B49383B08E7DB08E7DAC8877AC8877A78270AB9E98FF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFBD9494BD9494B39292B39292B18B8BB1
        8B8BA98888A98888A47878AA9F9EFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF}
      NumGlyphs = 2
    end
    object btnExit: TcxButton
      Left = 483
      Top = 5
      Width = 80
      Height = 27
      Cancel = True
      Caption = #32080#26463
      TabOrder = 1
      OnClick = btnExitClick
      Glyph.Data = {
        36060000424D3606000000000000360000002800000020000000100000000100
        1800000000000006000000000000000000000000000000000000FF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFCB6611FF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FF8C6060FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFCB6611CB6611FF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FF8C60608C6060FF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFCB6611FFE0B9CB66
        11FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FF8C6060FF00FF8C6060FF00FFFF00FFFF00FFFF00FFFF00FF
        FF00FFFF00FFCB6611CB6611CB6611CB6611CB6611CB6611CB6611FFE0B9FFE0
        B9CB6611FF00FFFF00FFFF00FFFF00FFFF00FFFF00FF8C60608C60608C60608C
        60608C60608C60608C6060FF00FFFF00FF8C6060FF00FFFF00FFFF00FFFF00FF
        FF00FFFF00FFCB6611FFE0B9FFE0B9FFE0B9FFE0B9FFE0B9FFE0B9FFE0B9FFE0
        B9FFE0B9CB6611FF00FFFF00FFFF00FFFF00FFFF00FF8C6060FF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF8C6060FF00FFAC5D27AC5D27
        AC5D27AC5D27CB6611FFE0B9FFE0B9FFE0B9FFE0B9FFE0B9FFE0B9FFE0B9FFE0
        B9FFE0B9FFE0B9CB66118857578857578857578857578C6060FF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF8C6060AC5D27FFFFFF
        FFFFFFFFFFFFAC5D27FFE0B9FFE0B9FFE0B9FFE0B9FFE0B9FFE0B9FFE0B9FFE0
        B9FFE0B9CB6611FF00FF885757FF00FFFF00FFFF00FF885757FF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF8C6060FF00FFAC5D27FFFFFF
        B58868B58868CB6611CB6611CB6611CB6611CB6611CB6611CB6611FFE0B9FFE0
        B9CB6611FF00FFFF00FF885757FF00FFB18B8BB18B8B8C60608C60608C60608C
        60608C60608C60608C6060FF00FFFF00FF8C6060FF00FFFF00FFAC5D27FFFFFF
        FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFAC5D27FF00FFCB6611FFE0B9CB66
        11FF00FFFF00FFFF00FF885757FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FF885757FF00FF8C6060FF00FF8C6060FF00FFFF00FFFF00FFAC5D27FFFFFF
        B58868B58868B58868B58868B58868FFFFFFAC5D27FF00FFCB6611CB6611FF00
        FFFF00FFFF00FFFF00FF885757FF00FFB18B8BB18B8BB18B8BB18B8BB18B8BFF
        00FF885757FF00FF8C60608C6060FF00FFFF00FFFF00FFFF00FFAC5D27FFFFFF
        FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFAC5D27FF00FFCB6611FF00FFFF00
        FFFF00FFFF00FFFF00FF885757FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FF885757FF00FF8C6060FF00FFFF00FFFF00FFFF00FFFF00FFAC5D27AC5D27
        AC5D27AC5D27AC5D27AC5D27AC5D27AC5D27AC5D27FF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FF88575788575788575788575788575788575788575788
        5757885757FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFAC5D27AC5D27
        AC5D27AC5D27AC5D27AC5D27AC5D27AC5D27AC5D27FF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FF88575788575788575788575788575788575788575788
        5757885757FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF}
      NumGlyphs = 2
    end
    object btnSave: TcxButton
      Left = 346
      Top = 5
      Width = 80
      Height = 27
      Caption = #23384#27284
      TabOrder = 2
      OnClick = btnSaveClick
      Glyph.Data = {
        36060000424D3606000000000000360000002800000020000000100000000100
        1800000000000006000000000000000000000000000000000000FF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
        9B7C6B9D7E6D9C7E6D9C7E6D9F8271C7BAB0709F642B811E1E710831781D6C96
        61C6BBAFFF00FFFF00FFFF00FFFF00FFA47878A47878A47878A47878A47878A4
        78787072743A3A3A3A3A3A3A3A3A6D6D6C857769FF00FFFF00FFFF00FFFF00FF
        9B7766FFFFFFFAF4E9FAF4E9F6F5EE41A43E009A0000A2070097000E89002E71
        003C7C2DFF00FFFF00FFFF00FFFF00FF857769FF00FFFF00FFFF00FFFF00FF42
        41410000000000000000000000003A3A3A3A3A3AFF00FFFF00FFFF00FFFF00FF
        A27F6FFFFFFFDDC2B5E6D2C882B7790DA71A2CB53FFDFBFF6FD4880098000095
        003070005A8A5AFF00FFFF00FFFF00FFA47878FF00FFDCC1C1E6CECE8181813A
        3A3A424141FF00FF8786870000000000003A3A3A59585CFF00FFFF00FFFF00FF
        A38070FFFFFFDBC3BBECDFDA43AA472EB33CF5ECF4FFF6FFFFFFFF6FD588009A
        001288002A751AFF00FFFF00FFFF00FFA47878FF00FFDCC1C1ECD9D94A4B4B3A
        3A3AFF00FFFF00FFFF00FF8786870000000000003A3A3AFF00FFFF00FFFF00FF
        A98778FFFFFFDBC7C2EEE4E23AAE445DB964C0D6BC03AD219DDBA7FFFFFF6DD4
        870095001B7106FF00FFFF00FFFF00FFA98888FF00FFDCC1C1EEE3E34241415E
        6062C5C2C33A3A3AA9A9A9FF00FF8786870000003A3A3AFF00FFFF00FFFF00FF
        AB897AFFFFFFDBC7C3ECE1DE4EB55743CD6D19BB4225C04F0BB22E9DD9A5FFFF
        FF67CF7C237A1BFF00FFFF00FFFF00FFA98888FF00FFDCC1C1EFDEDE59585C6D
        6D6C4241414A4B4B3A3A3AA9A9A9FF00FF8181813A3A3AFF00FFFF00FFFF00FF
        AF8E7FFFFFFFDCC5C0E5D4CF85C07E72DD9635CE6A2EC75E25BE4B09AF278BD0
        944EBB595C925CFF00FFFF00FFFF00FFB18B8BFF00FFDCC1C1E7D2D287868797
        97976A6A6A5C5F624A4B4B3A3A3A97979759585C5C5F62FF00FFFF00FFFF00FF
        AF8F80FFFFFFFEFEFEFEFEFEF9FBF95AC06071DD964AD37A29C2521CB73B0DAA
        22369530FF00FFFF00FFFF00FFFF00FFB18B8BFF00FFFF00FFFF00FFFF00FF5C
        5F629797977777774A4B4B3A3A3A3A3A3A3A3A3AFF00FFFF00FFFF00FFFF00FF
        AF8F80FFFFFFDFCECCDFCDCBDFCBC7E9DFDA86C07E52B85B3CB14844AC4889C4
        86C7BAAFFF00FFFF00FFFF00FFFF00FFB18B8BFF00FFDDCBCBDDCBCBE0C7C7EC
        D9D987868759585C4A4B4B4A4B4B878687A47878FF00FFFF00FFFF00FFFF00FF
        B19080FFFFFFFFFFFFFFFFFFFFFFFFFEFEFEFEFEFCFEFDFBFEFCFAFCF9F6F4EE
        E89E7F71FF00FFFF00FFFF00FFFF00FFB18B8BFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFA47878FF00FFFF00FFFF00FFFF00FF
        B79787FFFFFFFFFFFFFFFFFFFFFFFFFEFEFEFEFDFBFDFBF9A78270A78270A782
        70A78270FF00FFFF00FFFF00FFFF00FFB39292FF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFA47878A47878A47878A47878FF00FFFF00FFFF00FFFF00FF
        B89888FFFFFFFFFFFFFFFFFFFEFEFEFEFEFDFEFCFAFDFBF9A78270F5E2D9B18E
        7EAB9E98FF00FFFF00FFFF00FFFF00FFBD9494FF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFA47878FF00FFB18B8BA47878FF00FFFF00FFFF00FFFF00FF
        B89888FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFA78270B18E7EAB9E
        98FF00FFFF00FFFF00FFFF00FFFF00FFBD9494FF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFA47878B18B8BA47878FF00FFFF00FFFF00FFFF00FFFF00FF
        B89888B89888B49383B49383B08E7DB08E7DAC8877AC8877A78270AB9E98FF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFBD9494BD9494B39292B39292B18B8BB1
        8B8BA98888A98888A47878A47878FF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF}
      NumGlyphs = 2
    end
    object btnCancel: TcxButton
      Left = 261
      Top = 5
      Width = 80
      Height = 27
      Cancel = True
      Caption = #21462#28040
      TabOrder = 3
      OnClick = btnCancelClick
      Glyph.Data = {
        36060000424D3606000000000000360000002800000020000000100000000100
        1800000000000006000000000000000000000000000000000000FF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFC58634D48A
        2CC99350FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFB18B8B808080808080FF00FFFF00FFFF00FFFF00FF
        C09363FF00FFFF00FFFF00FFFF00FFFF00FFC28A4DD48729DF9239FFD0A0FFDB
        B5D68C2CFF00FFFF00FFFF00FFFF00FFBD9494FF00FFFF00FFFF00FFFF00FFFF
        00FFB18B8B808080808080FF00FFFF00FF808080FF00FFFF00FFFF00FFB78A5C
        BE854FFF00FFFF00FFFF00FFC78C4FD6842BEAA04CFFC98EFFC88EFFC78FFFCF
        A0DE8E2AFF00FFFF00FFFF00FFB18B8BB18B8BFF00FFFF00FFFF00FF80808080
        8080808080FF00FFFF00FFFF00FFFF00FF808080FF00FFFF00FFFF00FFA96D44
        BD8450FF00FFFF00FFFF00FFE2933EFFE2B6FFD19DFFCD99FFCC9AFFC48BFFCB
        95E79E4FBC8F5AFF00FFFF00FFA47878A47878FF00FFFF00FFFF00FF808080FF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FF808080B18B8BFF00FFFF00FFB96727
        B58360FF00FFFF00FFFF00FFC29162C28241B6753AD78628FFC587FFC286FFC7
        8EF7BA79BF833BFF00FFFF00FF8C6060A47878FF00FFFF00FFFF00FF808080A4
        7878A47878C7A386FF00FFFF00FFFF00FF808080A47878FF00FFFF00FFCC6A18
        B67A4EFF00FFFF00FFFF00FFFF00FFFF00FFC0813DFFBB77FFC282FEBE7DFFCE
        9BFFCF98C27C2AFF00FFFF00FF8C6060A47878FF00FFFF00FFFF00FFFF00FFFF
        00FFA47878FF00FFFF00FFFF00FFFF00FFFF00FFA47878FF00FFFF00FFCC6813
        CA732BFF00FFFF00FFFF00FFFF00FFFF00FFD9872DFFBF79FFC483D58930D890
        3EFFD8AAC9832FFF00FFFF00FF8C6060A47878FF00FFFF00FFFF00FFFF00FFFF
        00FF808080FF00FFFF00FF808080808080FF00FF808080FF00FFFF00FFBF651D
        EB7F16FF00FFFF00FFFF00FFFF00FFC67621FFBB6EFFBB71FCB86FC08340B587
        5BFFD7AAD48730FF00FFFF00FF8C6060A47878FF00FFFF00FFFF00FFFF00FFA4
        7878FF00FFFF00FFFF00FFA47878B18B8BFF00FF808080FF00FFFF00FFA8673D
        FFA223DB7314BC7231BC7539CD7720FFB35CFFB262FFBD73C67C2EFF00FFFF00
        FFD08635C48D4FFF00FFFF00FF808080D9AF8BA47878A47878A47878A4787880
        8080FF00FFFF00FFA47878FF00FFFF00FF808080808080FF00FFFF00FFFF00FF
        E27A1BFF9F2AFF9F2DFFA740FFAA4BFFAB50FFC986D27D2CB89979FF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFA47878FF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFA47878BD9494FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
        B6744DEC8C2FFFC679FFC37CFFCB8BFFC886CA7326B99471FF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFA47878FF00FFFF00FFFF00FFFF00FFFF
        00FFA47878BD9494FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
        FF00FFB28361B05F24C66618BF671DB07240FF00FFFF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFA478788C60608C60608C6060A4
        7878FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF}
      NumGlyphs = 2
    end
    object btnDelete: TcxButton
      Left = 176
      Top = 5
      Width = 80
      Height = 27
      Cancel = True
      Caption = #21034#38500
      TabOrder = 4
      OnClick = btnDeleteClick
      Glyph.Data = {
        36060000424D3606000000000000360000002800000020000000100000000100
        18000000000000060000F00A0000F00A00000000000000000000FF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF0205C3
        0A14AF867EB8D4C8C32E2DB32526B8B7ACC3A588799B7C6B9B7C6B9A7C6A997B
        689B7C6BFF00FFFF00FFFF00FFB79787B79787B79787D4C8C3B79787B79787B7
        9787A588799B7C6B9B7C6B9A7C6A997B689B7C6BFF00FFFF00FFFF00FF0203AD
        2C72FF1534D42135C4174EFF155CFF363AC5FDF9F3FAF2E6FAF1E4F9EFE0F8ED
        DA977967FF00FFFF00FFFF00FFB79787FF00FFB79787B79787FF00FFFF00FFB7
        9787FF00FFFF00FFFF00FFFF00FFFF00FF977967FF00FFFF00FFFF00FF0000C4
        1225C52C67FF255DFF205BFF141FBAD3C9D8DFC3B5DCBBA9DCBAA5DCBAA3FAF1
        E2987968FF00FFFF00FFFF00FFB79787B79787FF00FFFF00FFFF00FFB79787D3
        C9D8DFC3B5DCBBA9DCBAA5DCBAA3FF00FF987968FF00FFFF00FFFF00FFFF00FF
        A79EC11022BF2D66FF1C49F47471C4ECDDD6DBBDAFDABBAADAB8A5DCB9A5FAF3
        E6997A6AFF00FFFF00FFFF00FFFF00FFA79EC1B79787FF00FFB79787B79787FF
        00FFDBBDAFDABBAADAB8A5DCB9A5FF00FF997A6AFF00FFFF00FFFF00FFFF00FF
        2725A63F7DFF1C3FDD2961FF1223C4D2CADADEC4B8DBBCADDABAA8DCBBA7FBF4
        E8997B6BFF00FFFF00FFFF00FFFF00FFB79787FF00FFFF00FFFF00FFB79787FF
        00FFDEC4B8DBBCADDABAA8DCBBA7FF00FF997B6BFF00FFFF00FFFF00FFFF00FF
        151CC0396DFF8683CC3738B42B6DFF292ABDEEE2DCDBBDAFDBBAA9DCBCA9FBF5
        EA9A7C6BFF00FFFF00FFFF00FFFF00FFB79787FF00FFB79787B79787FF00FFB7
        9787FF00FFDBBDAFDBBAA9DCBCA9FF00FF9A7C6BFF00FFFF00FFFF00FFFF00FF
        A59BCA4242C4F0E6E4CAC3DA3231BBBDB8E5E7D4CCDBBDAFDBBAAADDBBA9FBF5
        EC9B7C6BFF00FFFF00FFFF00FFFF00FFA59BCAB79787FF00FFFF00FFB79787FF
        00FFFF00FFDBBDAFDBBAAADDBBA9FF00FF9B7C6BFF00FFFF00FFFF00FFFF00FF
        BCA194FFFFFFFEFEFEFEFEFEFFFEFEFEFDFCFDFBF8FDFAF6FCF9F3FCF7F0FCF6
        ED9B7C6CFF00FFFF00FFFF00FFFF00FFBCA194FF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FF9B7C6CFF00FFFF00FFFF00FFFF00FF
        AF8F80FFFFFFDFCECCDFCDCBDFCBC7DEC6C0DEC4BADEC1B4DEBEADDEBEABFCF6
        EE9C7C6DFF00FFFF00FFFF00FFFF00FFAF8F80FF00FFDFCECCDFCDCBDFCBC7DE
        C6C0DEC4BADEC1B4DEBEADDEBEABFF00FF9C7C6DFF00FFFF00FFFF00FFFF00FF
        B19080FFFFFFFFFFFFFFFFFFFFFFFFFEFEFEFEFDFBFDFBF8FCF9F4F9F4EEF0E8
        E09F8070FF00FFFF00FFFF00FFFF00FFB19080FF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FF9F8070FF00FFFF00FFFF00FFFF00FF
        B79787FFFFFFFFFFFFFFFFFFFFFFFFFEFEFEFEFDFBFDFBF9A78270A78270A782
        70A78270FF00FFFF00FFFF00FFFF00FFB79787FF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFA78270A78270A78270A78270FF00FFFF00FFFF00FFFF00FF
        B89888FFFFFFFFFFFFFFFFFFFEFEFEFEFEFDFEFCFAFDFBF9A78270F5E2D9B18E
        7EA78270FF00FFFF00FFFF00FFFF00FFB89888FF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFA78270FF00FFB18E7EA78270FF00FFFF00FFFF00FFFF00FF
        B89888FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFA78270B18E7EA782
        70FF00FFFF00FFFF00FFFF00FFFF00FFB89888FF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFA78270B18E7EA78270FF00FFFF00FFFF00FFFF00FFFF00FF
        B89888B89888B49383B49383B08E7DB08E7DAC8877AC8877A78270A78270FF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFB89888B89888B49383B49383B08E7DB0
        8E7DAC8877AC8877A78270A78270FF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF}
      NumGlyphs = 2
    end
    object btnEdit: TcxButton
      Left = 91
      Top = 5
      Width = 80
      Height = 27
      Caption = #20462#25913
      TabOrder = 5
      OnClick = btnEditClick
      Glyph.Data = {
        36060000424D3606000000000000360000002800000020000000100000000100
        1800000000000006000000000000000000000000000000000000FF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
        9B7C6B9D7E6D9C7E6D9C7E6D9C7E6D9C7D6D9C7D6C9B7C6B9B7C6B9A7C6A997B
        689B7C6BFF00FFFF00FFFF00FFFF00FFA47878A47878A47878A47878A47878A4
        7878A47878A47878A47878857769857769A47878FF00FFFF00FFFF00FFFF00FF
        9B7766FFFFFFFAF4E9FAF4E9EEE9DEE8E2D8F7F0E4FAF2E6FAF1E4F9EFE0F8ED
        DA977967FF00FFFF00FFFF00FFFF00FF857769FF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FF857769FF00FFFF00FFFF00FFFF00FF
        A27F6FFFFFFFDDC2B5DDC2B5B5A9A4B1A19ADBC2B4DCBBA9DCBAA5DCBAA3FAF1
        E2987968FF00FFFF00FFFF00FFFF00FFA47878FF00FFDCC1C1DCC1C1B1A6A6B3
        9F9FDCC1C1DABBBBDABBBBDABBBBFF00FF857769FF00FFFF00FFFF00FFFF00FF
        A38070FFFFFFDBC3BBEADDD76F6D71928B96C3CDB9E8D6CCDAB8A5DCB9A5FAF3
        E6997A6AFF00FFFF00FFFF00FFFF00FFA47878FF00FFDCC1C1EBD7D76D6D6C94
        9292C5C2C3E8D0D0DABBBBDABBBBFF00FF857769FF00FFFF00FFFF00FFFF00FF
        A98778FFFFFFDBC7C2E9DDDAA8BBCE34B3562CB44471B46FECDBD2DCBBA7FBF4
        E8997B6BFF00FFFF00FFFF00FFFF00FFA98888FF00FFDCC1C1ECD9D9B6B6B659
        585C424141737373ECD9D9DABBBBFF00FF857769FF00FFFF00FFFF00FFFF00FF
        AB897AFFFFFFDBC7C3E6D7D499D0A766FF985AEC862EAD4687BE81EAD8CCFBF5
        EA9A7C6BFF00FFFF00FFFF00FFFF00FFA98888FF00FFDCC1C1E9D4D4A9A9A999
        99998786874A4B4B878687EAD6D6FF00FFA47878FF00FFFF00FFFF00FFFF00FF
        AF8E7FFFFFFFDCC5C0DEC9C3DBE4D657E27F6AFF9D55E17C2AA43C9CC494FCF8
        F29B7C6BFF00FFFF00FFFF00FFFF00FFB18B8BFF00FFDCC1C1E0C7C7DCD7D781
        81819999998181813A3A3A9D9592FF00FFA47878FF00FFFF00FFFF00FFFF00FF
        AF8F80FFFFFFFEFEFEFEFEFEFEFDFDD1F8DC54EE8368FF9B50DC77249E38B7DC
        B6B69F94FF00FFFF00FFFF00FFFF00FFB18B8BFF00FFFF00FFFF00FFFF00FFD7
        D8D88282839999997777773A3A3AB6B6B6B79F9FFF00FFFF00FFFF00FFFF00FF
        AF8F80FFFFFFDFCECCDFCDCBDECAC6E9D9D5ADEAB95AF68864FF9742DA693487
        3EC6BDB6FF00FFFF00FFFF00FFFF00FFB18B8BFF00FFDDCBCBDDCBCBE0C7C7EA
        D6D6B6B6B6878687979797686868424141C6BFC0FF00FFFF00FFFF00FFFF00FF
        B19080FFFFFFFFFFFFFFFFFFFFFFFFFEFEFEFEFEFD91EFAC55FC889AC1A4CDBB
        B66E6D8CFF00FFFF00FFFF00FFFF00FFB18B8BFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFABACAC878687A9A9A9D2BCBC707274FF00FFFF00FFFF00FFFF00FF
        B79787FFFFFFFFFFFFFFFFFFFFFFFFFEFEFEFEFDFBFBFCFAA5C5A6FFFFFF7892
        F5203DDC292AA1FF00FFFF00FFFF00FFB39292FF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFA9A9A9FF00FF9492923A3A3A3A3A3AFF00FFFF00FFFF00FF
        B89888FFFFFFFFFFFFFFFFFFFEFEFEFEFEFDFEFCFAFDFBF9DED0C98C99DE4277
        FF2D4AD81F209FFF00FFFF00FFFF00FFBD9494FF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFDDCBCB9999997777774A4B4B3A3A3AFF00FFFF00FFFF00FF
        B89888FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFB39384C1B9D03545
        C52C30A7FF00FFFF00FFFF00FFFF00FFBD9494FF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFB39292C5C2C34241413A3A3AFF00FFFF00FFFF00FFFF00FF
        B89888B89888B49383B49383B08E7DB08E7DAC8877AC8877A78270B3A8A2FF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFBD9494BD9494B39292B39292B18B8BB1
        8B8BA98888A98888A47878B1A6A6FF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF}
      NumGlyphs = 2
    end
  end
  object dsrInv003: TDataSource
    DataSet = dtmMainJ.adoInv003Code
    OnDataChange = dsrInv003DataChange
    Left = 16
    Top = 368
  end
  object getXMLDocument: TXMLDocument
    Left = 16
    Top = 404
    DOMVendorDesc = 'MSXML'
  end
  object OpenDialog1: TOpenDialog
    Filter = 'Text files (*.txt)|*.txt'
    Left = 16
    Top = 332
  end
end
