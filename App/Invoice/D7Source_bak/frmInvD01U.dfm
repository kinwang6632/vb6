object frmInvD01: TfrmInvD01
  Left = 179
  Top = 112
  Width = 718
  Height = 503
  Caption = 'frmInvD01'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnCloseQuery = FormCloseQuery
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Label11: TLabel
    Left = 8
    Top = 83
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
  object Label12: TLabel
    Left = 128
    Top = 83
    Width = 65
    Height = 13
    Caption = #31237#31821#32232#34399#65306
    Font.Charset = CHINESEBIG5_CHARSET
    Font.Color = clTeal
    Font.Height = -13
    Font.Name = #26032#32048#26126#39636
    Font.Style = []
    ParentFont = False
    Transparent = True
  end
  object Label13: TLabel
    Left = 384
    Top = 83
    Width = 52
    Height = 13
    Caption = #36000#36012#20154#65306
    Font.Charset = CHINESEBIG5_CHARSET
    Font.Color = clTeal
    Font.Height = -13
    Font.Name = #26032#32048#26126#39636
    Font.Style = []
    ParentFont = False
    Transparent = True
  end
  object Label20: TLabel
    Left = 8
    Top = 113
    Width = 104
    Height = 13
    Caption = #31237#25424#31293#26680#34389#21517#31281#65306
    Font.Charset = CHINESEBIG5_CHARSET
    Font.Color = clTeal
    Font.Height = -13
    Font.Name = #26032#32048#26126#39636
    Font.Style = []
    ParentFont = False
    Transparent = True
  end
  object Label21: TLabel
    Left = 128
    Top = 113
    Width = 71
    Height = 13
    Caption = #31237#20989#32232#34399'1'#65306
    Font.Charset = CHINESEBIG5_CHARSET
    Font.Color = clTeal
    Font.Height = -13
    Font.Name = #26032#32048#26126#39636
    Font.Style = []
    ParentFont = False
    Transparent = True
  end
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 710
    Height = 32
    Align = alTop
    Color = clSkyBlue
    TabOrder = 0
    object Label1: TLabel
      Left = 16
      Top = 8
      Width = 102
      Height = 16
      Caption = #20844#21496#36039#26009#32173#35703
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
    Width = 710
    Height = 438
    Align = alClient
    TabOrder = 1
    object Panel3: TPanel
      Left = 1
      Top = 1
      Width = 708
      Height = 35
      Align = alTop
      TabOrder = 0
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
        Left = 581
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
    object pnlMaster: TPanel
      Left = 1
      Top = 36
      Width = 708
      Height = 401
      Align = alClient
      Caption = 'pnlMaster'
      TabOrder = 1
      object pnlGrid: TPanel
        Left = 1
        Top = 307
        Width = 706
        Height = 93
        Align = alClient
        Caption = 'pnlGrid'
        TabOrder = 0
        object DBGrid1: TDBGrid
          Left = 1
          Top = 1
          Width = 704
          Height = 74
          Align = alClient
          DataSource = dsrInv001
          Options = [dgTitles, dgIndicator, dgColumnResize, dgColLines, dgRowLines, dgTabs, dgConfirmDelete, dgCancelOnExit]
          TabOrder = 0
          TitleFont.Charset = DEFAULT_CHARSET
          TitleFont.Color = clWindowText
          TitleFont.Height = -11
          TitleFont.Name = 'MS Sans Serif'
          TitleFont.Style = []
        end
        object Stt_Show: TStaticText
          Left = 1
          Top = 75
          Width = 704
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
      object pnlEdit: TPanel
        Left = 1
        Top = 1
        Width = 706
        Height = 306
        Align = alTop
        TabOrder = 1
        object Label3: TLabel
          Left = 8
          Top = 15
          Width = 65
          Height = 13
          Caption = #20844#21496#20195#34399#65306
          Font.Charset = CHINESEBIG5_CHARSET
          Font.Color = clTeal
          Font.Height = -13
          Font.Name = #26032#32048#26126#39636
          Font.Style = []
          ParentFont = False
          Transparent = True
        end
        object Label2: TLabel
          Left = 144
          Top = 15
          Width = 65
          Height = 13
          Caption = #20844#21496#31777#31281#65306
          Font.Charset = CHINESEBIG5_CHARSET
          Font.Color = clTeal
          Font.Height = -13
          Font.Name = #26032#32048#26126#39636
          Font.Style = []
          ParentFont = False
          Transparent = True
        end
        object Label4: TLabel
          Left = 400
          Top = 15
          Width = 65
          Height = 13
          Caption = #20844#21496#20840#21517#65306
          Font.Charset = CHINESEBIG5_CHARSET
          Font.Color = clTeal
          Font.Height = -13
          Font.Name = #26032#32048#26126#39636
          Font.Style = []
          ParentFont = False
          Transparent = True
        end
        object Label5: TLabel
          Left = 8
          Top = 39
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
        object Label6: TLabel
          Left = 144
          Top = 39
          Width = 65
          Height = 13
          Caption = #31237#31821#32232#34399#65306
          Font.Charset = CHINESEBIG5_CHARSET
          Font.Color = clTeal
          Font.Height = -13
          Font.Name = #26032#32048#26126#39636
          Font.Style = []
          ParentFont = False
          Transparent = True
        end
        object Label7: TLabel
          Left = 300
          Top = 39
          Width = 52
          Height = 13
          Caption = #36000#36012#20154#65306
          Font.Charset = CHINESEBIG5_CHARSET
          Font.Color = clTeal
          Font.Height = -13
          Font.Name = #26032#32048#26126#39636
          Font.Style = []
          ParentFont = False
          Transparent = True
        end
        object Label8: TLabel
          Left = 8
          Top = 88
          Width = 65
          Height = 13
          Caption = #29151#26989#22320#22336#65306
          Font.Charset = CHINESEBIG5_CHARSET
          Font.Color = clTeal
          Font.Height = -13
          Font.Name = #26032#32048#26126#39636
          Font.Style = []
          ParentFont = False
          Transparent = True
        end
        object Label9: TLabel
          Left = 21
          Top = 63
          Width = 52
          Height = 13
          Caption = #30003#22577#21029#65306
          Font.Charset = CHINESEBIG5_CHARSET
          Font.Color = clTeal
          Font.Height = -13
          Font.Name = #26032#32048#26126#39636
          Font.Style = []
          ParentFont = False
          Transparent = True
        end
        object Label10: TLabel
          Left = 496
          Top = 39
          Width = 65
          Height = 13
          Caption = #20844#21496#38651#35441#65306
          Font.Charset = CHINESEBIG5_CHARSET
          Font.Color = clTeal
          Font.Height = -13
          Font.Name = #26032#32048#26126#39636
          Font.Style = []
          ParentFont = False
          Transparent = True
        end
        object Label14: TLabel
          Left = 21
          Top = 112
          Width = 91
          Height = 13
          Caption = #20844#24235#21123#25765#24115#34399#65306
          Font.Charset = CHINESEBIG5_CHARSET
          Font.Color = clTeal
          Font.Height = -13
          Font.Name = #26032#32048#26126#39636
          Font.Style = []
          ParentFont = False
          Transparent = True
        end
        object Label15: TLabel
          Left = 232
          Top = 112
          Width = 104
          Height = 13
          Caption = #32013#31237#20154#21123#25765#24115#34399#65306
          Font.Charset = CHINESEBIG5_CHARSET
          Font.Color = clTeal
          Font.Height = -13
          Font.Name = #26032#32048#26126#39636
          Font.Style = []
          ParentFont = False
          Transparent = True
        end
        object Label16: TLabel
          Left = 448
          Top = 112
          Width = 65
          Height = 13
          Caption = #29151#26989#31237#29575#65306
          Font.Charset = CHINESEBIG5_CHARSET
          Font.Color = clTeal
          Font.Height = -13
          Font.Name = #26032#32048#26126#39636
          Font.Style = []
          ParentFont = False
          Transparent = True
        end
        object Label17: TLabel
          Left = 8
          Top = 158
          Width = 104
          Height = 13
          Caption = #31237#25424#31293#26680#34389#21517#31281#65306
          Font.Charset = CHINESEBIG5_CHARSET
          Font.Color = clTeal
          Font.Height = -13
          Font.Name = #26032#32048#26126#39636
          Font.Style = []
          ParentFont = False
          Transparent = True
        end
        object Label18: TLabel
          Left = 40
          Top = 182
          Width = 71
          Height = 13
          Caption = #31237#20989#32232#34399'1'#65306
          Font.Charset = CHINESEBIG5_CHARSET
          Font.Color = clTeal
          Font.Height = -13
          Font.Name = #26032#32048#26126#39636
          Font.Style = []
          ParentFont = False
          Transparent = True
        end
        object Label19: TLabel
          Left = 380
          Top = 182
          Width = 71
          Height = 13
          Caption = #31237#20989#32232#34399'2'#65306
          Font.Charset = CHINESEBIG5_CHARSET
          Font.Color = clTeal
          Font.Height = -13
          Font.Name = #26032#32048#26126#39636
          Font.Style = []
          ParentFont = False
          Transparent = True
        end
        object Label22: TLabel
          Left = 470
          Top = 63
          Width = 91
          Height = 13
          Caption = #31237#20989#21462#24471#26085#26399#65306
          Font.Charset = CHINESEBIG5_CHARSET
          Font.Color = clTeal
          Font.Height = -13
          Font.Name = #26032#32048#26126#39636
          Font.Style = []
          ParentFont = False
          Transparent = True
        end
        object Label23: TLabel
          Left = 232
          Top = 63
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
        object Label24: TLabel
          Left = 422
          Top = 88
          Width = 91
          Height = 13
          Caption = #26989#32773#26381#21209#31278#39006#65306
          Font.Charset = CHINESEBIG5_CHARSET
          Font.Color = clTeal
          Font.Height = -13
          Font.Name = #26032#32048#26126#39636
          Font.Style = []
          ParentFont = False
          Transparent = True
        end
        object Label25: TLabel
          Left = 8
          Top = 134
          Width = 104
          Height = 13
          Caption = #33287#23458#26381#31995#32113#20171#25509#65306
          Font.Charset = CHINESEBIG5_CHARSET
          Font.Color = clTeal
          Font.Height = -13
          Font.Name = #26032#32048#26126#39636
          Font.Style = []
          ParentFont = False
          Transparent = True
        end
        object Label26: TLabel
          Left = 233
          Top = 134
          Width = 117
          Height = 13
          Caption = #23458#26381#31995#32113#23565#25033#32232#34399#65306
          Font.Charset = CHINESEBIG5_CHARSET
          Font.Color = clTeal
          Font.Height = -13
          Font.Name = #26032#32048#26126#39636
          Font.Style = []
          ParentFont = False
          Transparent = True
        end
        object Label27: TLabel
          Left = 409
          Top = 135
          Width = 133
          Height = 13
          Caption = #33258#21205#25563#38283#30332#31080#26126#32048#31558#25976':'
          Font.Charset = CHINESEBIG5_CHARSET
          Font.Color = clTeal
          Font.Height = -13
          Font.Name = #26032#32048#26126#39636
          Font.Style = []
          ParentFont = False
          Transparent = True
        end
        object Label28: TLabel
          Left = 615
          Top = 118
          Width = 32
          Height = 61
          AutoSize = False
          Caption = #35373#20633#24207#34399#26159#21542#21512#20341
          Font.Charset = CHINESEBIG5_CHARSET
          Font.Color = clTeal
          Font.Height = -13
          Font.Name = #26032#32048#26126#39636
          Font.Style = []
          ParentFont = False
          Transparent = True
          WordWrap = True
        end
        object Label29: TLabel
          Left = 8
          Top = 208
          Width = 103
          Height = 13
          Caption = 'EMAIL'#22846#27284#36335#24465#65306
          Font.Charset = CHINESEBIG5_CHARSET
          Font.Color = clTeal
          Font.Height = -13
          Font.Name = #26032#32048#26126#39636
          Font.Style = []
          ParentFont = False
          Transparent = True
        end
        object Label30: TLabel
          Left = 380
          Top = 208
          Width = 96
          Height = 13
          Caption = 'Email '#22846#27284#26684#24335#65306
          Font.Charset = CHINESEBIG5_CHARSET
          Font.Color = clTeal
          Font.Height = -13
          Font.Name = #26032#32048#26126#39636
          Font.Style = []
          ParentFont = False
          Transparent = True
        end
        object Label31: TLabel
          Left = 369
          Top = 158
          Width = 98
          Height = 13
          Caption = #25976#20301#26381#21209'Owner'#65306
          Font.Charset = CHINESEBIG5_CHARSET
          Font.Color = clTeal
          Font.Height = -13
          Font.Name = #26032#32048#26126#39636
          Font.Style = []
          ParentFont = False
          Transparent = True
        end
        object Label32: TLabel
          Left = 20
          Top = 231
          Width = 84
          Height = 13
          Caption = 'CM'#31995#32113#20195#30908#65306
          Font.Charset = CHINESEBIG5_CHARSET
          Font.Color = clTeal
          Font.Height = -13
          Font.Name = #26032#32048#26126#39636
          Font.Style = []
          ParentFont = False
          Transparent = True
        end
        object medCompID: TMaskEdit
          Left = 74
          Top = 10
          Width = 47
          Height = 21
          EditMask = 'aaaaaa;1;_'
          MaxLength = 6
          TabOrder = 0
          Text = '      '
        end
        object medCompSName: TMaskEdit
          Left = 216
          Top = 8
          Width = 177
          Height = 21
          TabOrder = 1
          Text = 'medCompSName'
          OnEnter = medCompSNameEnter
        end
        object medCompName: TMaskEdit
          Left = 472
          Top = 8
          Width = 177
          Height = 21
          TabOrder = 2
          Text = 'medCompName'
        end
        object medBusinessId: TMaskEdit
          Left = 74
          Top = 34
          Width = 63
          Height = 21
          EditMask = 'aaaaaaaa;1;_'
          MaxLength = 8
          TabOrder = 3
          Text = '        '
          OnExit = medBusinessIdExit
        end
        object medTaxId: TMaskEdit
          Left = 216
          Top = 34
          Width = 73
          Height = 21
          EditMask = 'aaaaaaaaa;1;_'
          MaxLength = 9
          TabOrder = 4
          Text = '         '
        end
        object medManager: TMaskEdit
          Left = 360
          Top = 34
          Width = 129
          Height = 21
          TabOrder = 5
        end
        object medTel: TMaskEdit
          Left = 560
          Top = 34
          Width = 89
          Height = 21
          TabOrder = 6
        end
        inline fraTaxGetDate: TfraYMD
          Left = 559
          Top = 52
          Width = 94
          Height = 30
          Font.Charset = ANSI_CHARSET
          Font.Color = clWindowText
          Font.Height = -16
          Font.Name = #26032#32048#26126#39636
          Font.Style = []
          ParentFont = False
          TabOrder = 9
          inherited mseYMD: TMaskEdit
            Left = 1
            Top = 6
            Width = 88
            Font.Charset = CHINESEBIG5_CHARSET
            Font.Color = clBlack
            Font.Height = -13
            Font.Name = #26032#32048#26126#39636
            ParentFont = False
          end
        end
        object medCompAddr: TMaskEdit
          Left = 74
          Top = 84
          Width = 328
          Height = 21
          TabOrder = 10
        end
        object medDecaCount: TMaskEdit
          Left = 114
          Top = 108
          Width = 79
          Height = 21
          EditMask = 'aaaaaaaaa;1;_'
          MaxLength = 9
          TabOrder = 12
          Text = '         '
        end
        object medPayCount: TMaskEdit
          Left = 338
          Top = 108
          Width = 63
          Height = 21
          EditMask = 'aaaaaaa;1;_'
          MaxLength = 7
          TabOrder = 13
          Text = '       '
        end
        object medTaxRate: TMaskEdit
          Left = 514
          Top = 108
          Width = 83
          Height = 21
          EditMask = 'aaaaa;1;_'
          MaxLength = 5
          TabOrder = 14
          Text = '     '
        end
        object medTaxDivName: TMaskEdit
          Left = 114
          Top = 155
          Width = 252
          Height = 21
          Font.Charset = ANSI_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
          TabOrder = 18
        end
        object medTaxNum1: TMaskEdit
          Left = 115
          Top = 179
          Width = 250
          Height = 21
          TabOrder = 20
        end
        object medTaxNum2: TMaskEdit
          Left = 457
          Top = 179
          Width = 140
          Height = 21
          TabOrder = 21
        end
        object cobNotiType: TComboBox
          Left = 74
          Top = 59
          Width = 89
          Height = 21
          ItemHeight = 13
          ItemIndex = 0
          TabOrder = 7
          Text = #26399
          Items.Strings = (
            #26399
            #26376)
        end
        object cobIsSelfCreated: TComboBox
          Left = 360
          Top = 59
          Width = 81
          Height = 21
          ItemHeight = 13
          ItemIndex = 0
          TabOrder = 8
          Text = #26159
          Items.Strings = (
            #26159
            #21542)
        end
        object medServiceTypeStr: TMaskEdit
          Left = 514
          Top = 84
          Width = 83
          Height = 21
          TabOrder = 11
          OnExit = medServiceTypeStrExit
        end
        object cmbLinkToMis: TComboBox
          Left = 114
          Top = 131
          Width = 75
          Height = 21
          Style = csDropDownList
          ItemHeight = 13
          TabOrder = 15
          Items.Strings = (
            #26159
            #21542)
        end
        object edtConnSeq: TEdit
          Left = 352
          Top = 131
          Width = 49
          Height = 21
          MaxLength = 3
          TabOrder = 16
          OnKeyPress = edtConnSeqKeyPress
        end
        object edtAutoCreateNum: TcxMaskEdit
          Left = 547
          Top = 131
          Properties.MaskKind = emkRegExpr
          Properties.EditMask = '(\d )|( \d\d)'
          Properties.MaxLength = 0
          Style.StyleController = dtmMain.cxEditStyle
          TabOrder = 17
          Width = 49
        end
        object chkFaciCombine: TcxCheckBox
          Left = 619
          Top = 94
          Properties.Alignment = taLeftJustify
          Properties.NullStyle = nssUnchecked
          Properties.ValueChecked = 1
          Properties.ValueGrayed = 0
          Properties.ValueUnchecked = '0'
          Style.StyleController = dtmMain.cxEditStyle
          TabOrder = 29
          Width = 21
        end
        object chkShowFaci: TcxCheckBox
          Left = 192
          Top = 229
          Caption = #26159#21542#39023#31034#35373#20633#24207#34399
          ParentFont = False
          Properties.Alignment = taLeftJustify
          Properties.NullStyle = nssUnchecked
          Properties.ValueChecked = 1
          Properties.ValueGrayed = 0
          Properties.ValueUnchecked = '0'
          Style.Font.Charset = ANSI_CHARSET
          Style.Font.Color = clWindowText
          Style.Font.Height = -13
          Style.Font.Name = #26032#32048#26126#39636
          Style.Font.Style = []
          Style.StyleController = dtmMain.cxEditStyle
          Style.TextColor = clTeal
          Style.IsFontAssigned = True
          TabOrder = 25
          Width = 137
        end
        object chkStarEInvoice: TcxCheckBox
          Left = 0
          Top = 255
          Caption = #21855#29992#38651#23376#30332#31080
          ParentFont = False
          Properties.Alignment = taLeftJustify
          Properties.DisplayChecked = '1'
          Properties.DisplayUnchecked = '0'
          Properties.NullStyle = nssUnchecked
          Properties.ValueChecked = 1
          Properties.ValueUnchecked = 0
          Properties.OnEditValueChanged = chkStarEInvoicePropertiesEditValueChanged
          Style.Font.Charset = ANSI_CHARSET
          Style.Font.Color = clWindowText
          Style.Font.Height = -13
          Style.Font.Name = #26032#32048#26126#39636
          Style.Font.Style = []
          Style.StyleController = dtmMain.cxEditStyle
          Style.TextColor = clTeal
          Style.IsFontAssigned = True
          TabOrder = 26
          Width = 120
        end
        object chkStarEmail: TcxCheckBox
          Left = 99
          Top = 255
          Caption = '1.'#21855#29992'Email'#30332#36865
          ParentFont = False
          Properties.Alignment = taLeftJustify
          Properties.DisplayChecked = '1'
          Properties.DisplayUnchecked = '0'
          Properties.NullStyle = nssUnchecked
          Properties.ValueChecked = 1
          Properties.ValueUnchecked = 0
          Style.Font.Charset = ANSI_CHARSET
          Style.Font.Color = clWindowText
          Style.Font.Height = -13
          Style.Font.Name = #26032#32048#26126#39636
          Style.Font.Style = []
          Style.StyleController = dtmMain.cxEditStyle
          Style.TextColor = clTeal
          Style.IsFontAssigned = True
          TabOrder = 27
          Width = 120
        end
        object chkStarMessage: TcxCheckBox
          Left = 208
          Top = 255
          Caption = '2.'#21855#29992#31777#35338#30332#36865
          ParentFont = False
          Properties.Alignment = taLeftJustify
          Properties.DisplayChecked = '1'
          Properties.DisplayUnchecked = '0'
          Properties.NullStyle = nssUnchecked
          Properties.ValueChecked = 1
          Properties.ValueUnchecked = 0
          Style.Font.Charset = ANSI_CHARSET
          Style.Font.Color = clWindowText
          Style.Font.Height = -13
          Style.Font.Name = #26032#32048#26126#39636
          Style.Font.Style = []
          Style.StyleController = dtmMain.cxEditStyle
          Style.TextColor = clTeal
          Style.IsFontAssigned = True
          TabOrder = 28
          Width = 115
        end
        object chkStarTVMail: TcxCheckBox
          Left = 314
          Top = 255
          Caption = '3.'#21855#29992'TVMail'#30332#36865
          ParentFont = False
          Properties.Alignment = taLeftJustify
          Properties.DisplayChecked = '1'
          Properties.DisplayUnchecked = '0'
          Properties.NullStyle = nssUnchecked
          Properties.ValueChecked = 1
          Properties.ValueUnchecked = 0
          Style.Font.Charset = ANSI_CHARSET
          Style.Font.Color = clWindowText
          Style.Font.Height = -13
          Style.Font.Name = #26032#32048#26126#39636
          Style.Font.Style = []
          Style.StyleController = dtmMain.cxEditStyle
          Style.TextColor = clTeal
          Style.IsFontAssigned = True
          TabOrder = 30
          Width = 126
        end
        object chkStarCMSend: TcxCheckBox
          Left = 435
          Top = 255
          Caption = '4.'#21855#29992'CM'#23566#27969#30332#36865
          ParentFont = False
          Properties.Alignment = taLeftJustify
          Properties.DisplayChecked = '1'
          Properties.DisplayUnchecked = '0'
          Properties.NullStyle = nssUnchecked
          Properties.ValueChecked = 1
          Properties.ValueUnchecked = 0
          Style.Font.Charset = ANSI_CHARSET
          Style.Font.Color = clWindowText
          Style.Font.Height = -13
          Style.Font.Name = #26032#32048#26126#39636
          Style.Font.Style = []
          Style.StyleController = dtmMain.cxEditStyle
          Style.TextColor = clTeal
          Style.IsFontAssigned = True
          TabOrder = 31
          Width = 130
        end
        object medGraphPath: TMaskEdit
          Left = 114
          Top = 204
          Width = 251
          Height = 21
          MaxLength = 60
          TabOrder = 22
        end
        object cboEmailFileType: TComboBox
          Left = 474
          Top = 202
          Width = 83
          Height = 21
          ItemHeight = 13
          TabOrder = 23
          Text = '0. JPG'#22294#27284
          Items.Strings = (
            '0. JPG'#22294#27284
            '1. PDF'#27284
            '2. CSV'#27284)
        end
        object chkAutoNotify: TcxCheckBox
          Left = 328
          Top = 229
          Caption = #21855#29992#33258#21205#38651#23376#30332#31080#36890#30693
          ParentFont = False
          Properties.Alignment = taLeftJustify
          Properties.DisplayChecked = '1'
          Properties.DisplayUnchecked = '0'
          Properties.NullStyle = nssUnchecked
          Properties.ValueChecked = 1
          Properties.ValueUnchecked = 0
          Properties.OnEditValueChanged = chkStarEInvoicePropertiesEditValueChanged
          Style.Font.Charset = ANSI_CHARSET
          Style.Font.Color = clWindowText
          Style.Font.Height = -13
          Style.Font.Name = #26032#32048#26126#39636
          Style.Font.Style = []
          Style.StyleController = dtmMain.cxEditStyle
          Style.TextColor = clTeal
          Style.IsFontAssigned = True
          TabOrder = 33
          Width = 160
        end
        object edtMisOwner: TEdit
          Left = 468
          Top = 155
          Width = 129
          Height = 21
          MaxLength = 15
          TabOrder = 19
        end
        object edtSysID: TEdit
          Left = 114
          Top = 227
          Width = 75
          Height = 21
          MaxLength = 3
          TabOrder = 24
        end
        object chkMaskInvNo: TcxCheckBox
          Left = 0
          Top = 277
          Caption = #36974#32617#25424#36104#30340#30332#31080#34399#30908
          ParentFont = False
          Properties.Alignment = taLeftJustify
          Properties.DisplayChecked = '1'
          Properties.DisplayUnchecked = '0'
          Properties.NullStyle = nssUnchecked
          Properties.ValueChecked = 1
          Properties.ValueUnchecked = 0
          Properties.OnEditValueChanged = chkStarEInvoicePropertiesEditValueChanged
          Style.Font.Charset = ANSI_CHARSET
          Style.Font.Color = clWindowText
          Style.Font.Height = -13
          Style.Font.Name = #26032#32048#26126#39636
          Style.Font.Style = []
          Style.StyleController = dtmMain.cxEditStyle
          Style.TextColor = clTeal
          Style.IsFontAssigned = True
          TabOrder = 34
          Width = 153
        end
        object chkNotifyPrize: TcxCheckBox
          Left = 146
          Top = 276
          Caption = #21855#29992#20013#29518#36890#30693
          ParentFont = False
          Properties.Alignment = taLeftJustify
          Properties.DisplayChecked = '1'
          Properties.DisplayUnchecked = '0'
          Properties.NullStyle = nssUnchecked
          Properties.ValueChecked = 1
          Properties.ValueUnchecked = 0
          Properties.OnEditValueChanged = chkStarEInvoicePropertiesEditValueChanged
          Style.Font.Charset = ANSI_CHARSET
          Style.Font.Color = clWindowText
          Style.Font.Height = -13
          Style.Font.Name = #26032#32048#26126#39636
          Style.Font.Style = []
          Style.StyleController = dtmMain.cxEditStyle
          Style.TextColor = clTeal
          Style.IsFontAssigned = True
          TabOrder = 35
          Width = 111
        end
        object chkStarCMTVMail: TcxCheckBox
          Left = 564
          Top = 255
          Caption = '5.'#21855#29992'CM'#20197'TVMAIL'
          ParentFont = False
          Properties.Alignment = taLeftJustify
          Properties.DisplayChecked = '1'
          Properties.DisplayUnchecked = '0'
          Properties.NullStyle = nssUnchecked
          Properties.ValueChecked = 1
          Properties.ValueUnchecked = 0
          Style.Font.Charset = ANSI_CHARSET
          Style.Font.Color = clWindowText
          Style.Font.Height = -13
          Style.Font.Name = #26032#32048#26126#39636
          Style.Font.Style = []
          Style.StyleController = dtmMain.cxEditStyle
          Style.TextColor = clTeal
          Style.IsFontAssigned = True
          TabOrder = 32
          Width = 141
        end
      end
    end
  end
  object dsrInv001: TDataSource
    DataSet = dtmMainJ.adoInv001Code
    OnDataChange = dsrInv001DataChange
    Left = 574
    Top = 284
  end
end
