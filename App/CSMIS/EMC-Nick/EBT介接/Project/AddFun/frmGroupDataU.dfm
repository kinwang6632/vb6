object frmGroupData: TfrmGroupData
  Left = 66
  Top = 59
  Width = 696
  Height = 480
  Caption = 'frmGroupData'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -13
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 16
  object StatusBar1: TStatusBar
    Left = 0
    Top = 434
    Width = 688
    Height = 19
    Panels = <>
    SimplePanel = False
  end
  object pnlFun: TPanel
    Left = 0
    Top = 0
    Width = 688
    Height = 41
    Align = alTop
    TabOrder = 1
    object btnAppend: TBitBtn
      Left = 31
      Top = 11
      Width = 61
      Height = 20
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
      Left = 96
      Top = 11
      Width = 61
      Height = 20
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
      Left = 161
      Top = 11
      Width = 61
      Height = 20
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
      Left = 226
      Top = 11
      Width = 61
      Height = 20
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
      Left = 291
      Top = 11
      Width = 61
      Height = 20
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
      Left = 357
      Top = 11
      Width = 61
      Height = 20
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
  object pnlSingleData: TPanel
    Left = 0
    Top = 41
    Width = 688
    Height = 144
    Align = alTop
    TabOrder = 2
    object dbtCodeNo: TDBEdit
      Left = 56
      Top = 18
      Width = 121
      Height = 24
      DataField = 'CODENO'
      DataSource = dsrGroupData
      TabOrder = 0
    end
    object StaticText1: TStaticText
      Left = 24
      Top = 18
      Width = 28
      Height = 20
      Caption = #20195#30908
      TabOrder = 1
    end
    object StaticText3: TStaticText
      Left = 24
      Top = 50
      Width = 28
      Height = 20
      Caption = #21517#31281
      TabOrder = 2
    end
    object DBRadioGroup1: TDBRadioGroup
      Left = 188
      Top = 10
      Width = 97
      Height = 121
      Caption = ' '#26159#21542#20572#29992' '
      DataField = 'STOPFLAG'
      DataSource = dsrGroupData
      Items.Strings = (
        #26159
        #21542)
      TabOrder = 3
      Values.Strings = (
        '1'
        '0')
    end
    object dbtDesc: TDBEdit
      Left = 56
      Top = 50
      Width = 121
      Height = 24
      DataField = 'DESCRIPTION'
      DataSource = dsrGroupData
      TabOrder = 4
    end
    object rgbSo: TGroupBox
      Left = 294
      Top = 10
      Width = 371
      Height = 119
      Caption = ' '#21487#29992'SO '
      TabOrder = 5
      object chb9: TCheckBox
        Tag = 9
        Left = 24
        Top = 24
        Width = 97
        Height = 17
        Caption = #38525#26126#23665
        TabOrder = 0
      end
      object chb10: TCheckBox
        Tag = 10
        Left = 24
        Top = 48
        Width = 97
        Height = 17
        Caption = #26032#21488#21271
        TabOrder = 1
      end
      object chb12: TCheckBox
        Tag = 12
        Left = 24
        Top = 72
        Width = 97
        Height = 17
        Caption = #22823#23433#25991#23665
        TabOrder = 2
      end
      object chb11: TCheckBox
        Tag = 11
        Left = 24
        Top = 96
        Width = 97
        Height = 17
        Caption = #37329#38971#36947
        TabOrder = 3
      end
      object chb8: TCheckBox
        Tag = 8
        Left = 112
        Top = 24
        Width = 97
        Height = 17
        Caption = #20840#32879
        TabOrder = 4
      end
      object chb13: TCheckBox
        Tag = 13
        Left = 112
        Top = 48
        Width = 97
        Height = 17
        Caption = #26032#21776#22478
        TabOrder = 5
      end
      object chb14: TCheckBox
        Tag = 14
        Left = 112
        Top = 72
        Width = 97
        Height = 17
        Caption = #22823#26032#24215
        TabOrder = 6
      end
      object chb16: TCheckBox
        Tag = 16
        Left = 112
        Top = 96
        Width = 97
        Height = 17
        Caption = #21271#26691#22290
        TabOrder = 7
      end
      object chb7: TCheckBox
        Tag = 7
        Left = 192
        Top = 24
        Width = 97
        Height = 17
        Caption = #25391#36947
        TabOrder = 8
      end
      object chb6: TCheckBox
        Tag = 6
        Left = 192
        Top = 48
        Width = 97
        Height = 17
        Caption = #35920#30431
        TabOrder = 9
      end
      object chb5: TCheckBox
        Tag = 5
        Left = 192
        Top = 72
        Width = 97
        Height = 17
        Caption = #26032#38971#36947
        TabOrder = 10
      end
      object chb3: TCheckBox
        Tag = 3
        Left = 192
        Top = 96
        Width = 97
        Height = 17
        Caption = #21335#22825
        TabOrder = 11
      end
      object chb1: TCheckBox
        Tag = 1
        Left = 272
        Top = 24
        Width = 97
        Height = 17
        Caption = #35264#26119
        TabOrder = 12
      end
      object chb2: TCheckBox
        Tag = 2
        Left = 272
        Top = 48
        Width = 97
        Height = 17
        Caption = #23631#21335
        TabOrder = 13
      end
    end
  end
  object pnlMultiData: TPanel
    Left = 0
    Top = 185
    Width = 688
    Height = 249
    Align = alClient
    TabOrder = 3
    object dbgGroupData: TDBGrid
      Left = 1
      Top = 1
      Width = 686
      Height = 247
      Align = alClient
      Color = clInfoBk
      DataSource = dsrGroupData
      ReadOnly = True
      TabOrder = 0
      TitleFont.Charset = DEFAULT_CHARSET
      TitleFont.Color = clWindowText
      TitleFont.Height = -13
      TitleFont.Name = 'MS Sans Serif'
      TitleFont.Style = []
      Columns = <
        item
          Expanded = False
          FieldName = 'CODENO'
          Title.Caption = #20195#30908
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'DESCRIPTION'
          Title.Caption = #21517#31281
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'SONAME'
          Title.Caption = #21487#29992'SO'
          Width = 120
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'STOPFLAGDESC'
          Title.Caption = #26159#21542#20572#29992
          Width = 50
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'UPDTIME'
          Title.Caption = #30064#21205#26178#38291
          Width = 100
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'UPDEN'
          Title.Caption = #30064#21205#20154#21729
          Visible = True
        end>
    end
  end
  object dsrGroupData: TDataSource
    DataSet = dtmMain.qryGroupData
    Left = 568
    Top = 16
  end
end
