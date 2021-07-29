object frmParam: TfrmParam
  Left = 476
  Top = 303
  Width = 403
  Height = 319
  Caption = #31995#32113#21443#25976#35373#23450
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
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 395
    Height = 292
    Align = alClient
    BevelOuter = bvLowered
    TabOrder = 0
    object Panel2: TPanel
      Left = 1
      Top = 1
      Width = 393
      Height = 41
      Align = alTop
      TabOrder = 0
      object btnExit: TBitBtn
        Left = 322
        Top = 9
        Width = 65
        Height = 24
        Caption = #32080#26463
        TabOrder = 0
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
      object btnEdit: TBitBtn
        Left = 106
        Top = 9
        Width = 65
        Height = 24
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
      object btnCancel: TBitBtn
        Left = 178
        Top = 9
        Width = 65
        Height = 24
        Caption = #21462#28040
        TabOrder = 2
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
        Left = 251
        Top = 9
        Width = 65
        Height = 24
        Caption = #20786#23384
        TabOrder = 3
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
    end
    object pnlEdit: TPanel
      Left = 1
      Top = 42
      Width = 393
      Height = 249
      Align = alClient
      TabOrder = 1
      object RadioGroup1: TRadioGroup
        Left = 24
        Top = 11
        Width = 345
        Height = 57
        Caption = ' '#31243#24335#36039#35338#35373#23450' '
        TabOrder = 0
      end
      object DBEdit1: TDBEdit
        Left = 94
        Top = 33
        Width = 251
        Height = 24
        DataField = 'sSysName'
        DataSource = dsrParam
        TabOrder = 1
      end
      object StaticText3: TStaticText
        Left = 40
        Top = 36
        Width = 52
        Height = 20
        Caption = #31243#24335#21517#31281
        TabOrder = 2
      end
      object RadioGroup2: TRadioGroup
        Left = 24
        Top = 81
        Width = 345
        Height = 144
        Caption = ' '#22519#34892#21443#25976#35373#23450' '
        TabOrder = 3
      end
      object DBEdit2: TDBEdit
        Left = 66
        Top = 103
        Width = 39
        Height = 24
        DataField = 'nProcessPeriod'
        DataSource = dsrParam
        TabOrder = 4
      end
      object StaticText2: TStaticText
        Left = 109
        Top = 108
        Width = 180
        Height = 20
        Caption = #20998#37912#34389#29702#19968#27425#36039#26009'('#36039#26009#34389#29702#36913#26399')'
        TabOrder = 5
      end
      object StaticText1: TStaticText
        Left = 49
        Top = 107
        Width = 16
        Height = 20
        Caption = #27599
        TabOrder = 6
      end
      object DBCheckBox1: TDBCheckBox
        Left = 49
        Top = 134
        Width = 150
        Height = 17
        Caption = #21855#21205#36039#26009#34389#29702'log'
        DataField = 'bLogData'
        DataSource = dsrParam
        TabOrder = 7
        ValueChecked = 'True'
        ValueUnchecked = 'False'
      end
      object DBCheckBox2: TDBCheckBox
        Left = 49
        Top = 159
        Width = 150
        Height = 17
        Caption = #39023#31034#36039#26009#34389#29702'UI'
        DataField = 'bShowUI'
        DataSource = dsrParam
        TabOrder = 8
        ValueChecked = 'True'
        ValueUnchecked = 'False'
      end
      object DBCheckBox3: TDBCheckBox
        Left = 49
        Top = 184
        Width = 255
        Height = 17
        Caption = 'Update '#26178#23559#31354#30333#27396#20301#23531#20837#31995#32113#21488
        DataField = 'bAcceptNullData'
        DataSource = dsrParam
        TabOrder = 9
        ValueChecked = 'True'
        ValueUnchecked = 'False'
      end
    end
  end
  object dsrParam: TDataSource
    DataSet = dtmMain.cdsParam
    Left = 304
    Top = 148
  end
end
