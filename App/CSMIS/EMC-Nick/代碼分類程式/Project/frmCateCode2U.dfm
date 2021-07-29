object frmCateCode2: TfrmCateCode2
  Left = 1
  Top = 0
  Width = 798
  Height = 548
  Caption = 'SO8F10_2'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -13
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  WindowState = wsMaximized
  OnClose = FormClose
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 16
  object pnlMFunction: TPanel
    Left = 0
    Top = 0
    Width = 790
    Height = 73
    Align = alTop
    BevelInner = bvLowered
    TabOrder = 0
    object spbMInsert: TSpeedButton
      Left = 43
      Top = 34
      Width = 67
      Height = 31
      Caption = '&A'#26032#22686
      NumGlyphs = 2
      OnClick = spbMInsertClick
    end
    object spbMEdit: TSpeedButton
      Left = 110
      Top = 34
      Width = 67
      Height = 31
      Caption = '&M'#20462#25913
      NumGlyphs = 2
      OnClick = spbMEditClick
    end
    object spbMDelete: TSpeedButton
      Left = 177
      Top = 34
      Width = 67
      Height = 31
      Caption = '&D'#21034#38500
      NumGlyphs = 2
      OnClick = spbMDeleteClick
    end
    object spbMCancel: TSpeedButton
      Left = 244
      Top = 34
      Width = 67
      Height = 31
      Caption = '&C'#21462#28040
      NumGlyphs = 2
      OnClick = spbMCancelClick
    end
    object spbMPost: TSpeedButton
      Left = 311
      Top = 34
      Width = 67
      Height = 31
      Caption = '&S'#23384#27284
      NumGlyphs = 2
      OnClick = spbMPostClick
    end
    object spbExit: TSpeedButton
      Left = 378
      Top = 34
      Width = 67
      Height = 31
      Caption = '&O'#32080#26463
      NumGlyphs = 2
      OnClick = spbExitClick
    end
    object spbNext: TSpeedButton
      Left = 633
      Top = 34
      Width = 78
      Height = 31
      Caption = '&N'#19979#19968#31558
      NumGlyphs = 2
      OnClick = spbNextClick
    end
    object spbLast: TSpeedButton
      Left = 710
      Top = 34
      Width = 78
      Height = 31
      Caption = '&L'#26368#24460#19968#31558
      NumGlyphs = 2
      OnClick = spbLastClick
    end
    object spbFirst: TSpeedButton
      Left = 480
      Top = 34
      Width = 78
      Height = 31
      Caption = '&F'#31532#19968#31558
      NumGlyphs = 2
      OnClick = spbFirstClick
    end
    object spbPrior: TSpeedButton
      Left = 557
      Top = 34
      Width = 78
      Height = 31
      Caption = '&P'#19978#19968#31558
      NumGlyphs = 2
      OnClick = spbPriorClick
    end
    object stxTableName: TStaticText
      Left = 136
      Top = 8
      Width = 117
      Height = 24
      Caption = 'stxTableName'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlue
      Font.Height = -16
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      TabOrder = 0
    end
    object StaticText7: TStaticText
      Left = 48
      Top = 8
      Width = 72
      Height = 20
      Caption = #20998#39006#38917#30446
      Font.Charset = ANSI_CHARSET
      Font.Color = clRed
      Font.Height = -16
      Font.Name = #32048#26126#39636
      Font.Style = [fsBold]
      ParentFont = False
      TabOrder = 1
    end
  end
  object pnlMainMaster: TPanel
    Left = 0
    Top = 73
    Width = 790
    Height = 150
    Align = alTop
    BevelOuter = bvNone
    TabOrder = 1
    object pnlRMaster: TPanel
      Left = 441
      Top = 0
      Width = 349
      Height = 150
      Align = alClient
      TabOrder = 0
      object DBGrid1: TDBGrid
        Left = 1
        Top = 1
        Width = 347
        Height = 148
        Align = alClient
        Color = clInfoBk
        DataSource = dsrMaster
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
            Width = 50
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'DESCRIPTION'
            Title.Caption = #35498#26126
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'REFNO'
            Title.Caption = #21443#32771#34399
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'ServiceTypeName'
            Title.Caption = #26381#21209#21029
            Width = 60
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'StopFlagName'
            Title.Caption = #26159#21542#20572#29992
            Width = 50
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'OPERATOR'
            Title.Caption = #30064#21205#20154#21729
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'UPDTIME'
            Title.Caption = #30064#21205#26178#38291
            Visible = True
          end>
      end
    end
    object pnlMaster: TPanel
      Left = 0
      Top = 0
      Width = 441
      Height = 150
      Align = alLeft
      TabOrder = 1
      object DBEdit1: TDBEdit
        Left = 88
        Top = 16
        Width = 121
        Height = 24
        DataField = 'CODENO'
        DataSource = dsrMaster
        TabOrder = 0
      end
      object DBEdit2: TDBEdit
        Left = 88
        Top = 60
        Width = 121
        Height = 24
        DataField = 'DESCRIPTION'
        DataSource = dsrMaster
        TabOrder = 1
      end
      object DBEdit3: TDBEdit
        Left = 88
        Top = 105
        Width = 121
        Height = 24
        DataField = 'REFNO'
        DataSource = dsrMaster
        TabOrder = 2
      end
      object StaticText1: TStaticText
        Left = 40
        Top = 19
        Width = 28
        Height = 20
        Caption = #20195#30908
        TabOrder = 3
      end
      object StaticText2: TStaticText
        Left = 40
        Top = 63
        Width = 28
        Height = 20
        Caption = #35498#26126
        TabOrder = 4
      end
      object StaticText3: TStaticText
        Left = 32
        Top = 108
        Width = 40
        Height = 20
        Caption = #21443#32771#34399
        TabOrder = 5
      end
      object rgpServiceType: TRadioGroup
        Left = 240
        Top = 16
        Width = 185
        Height = 57
        Caption = ' '#26381#21209#21029' '
        Columns = 2
        ItemIndex = 0
        Items.Strings = (
          #26377#32218#38651#35222
          'CM')
        TabOrder = 6
      end
      object DBRadioGroup1: TDBRadioGroup
        Left = 240
        Top = 80
        Width = 185
        Height = 49
        Caption = ' '#26159#21542#20572#29992' '
        Columns = 2
        DataField = 'STOPFLAG'
        DataSource = dsrMaster
        Items.Strings = (
          #26159
          #21542)
        TabOrder = 7
        Values.Strings = (
          '1'
          '0')
      end
    end
  end
  object pnlMainDetail: TPanel
    Left = 0
    Top = 273
    Width = 790
    Height = 248
    Align = alClient
    TabOrder = 2
    object pnlSingleDetail: TPanel
      Left = 1
      Top = 1
      Width = 788
      Height = 97
      Align = alTop
      BevelOuter = bvLowered
      Enabled = False
      TabOrder = 0
      object DBEdit5: TDBEdit
        Left = 77
        Top = 56
        Width = 121
        Height = 24
        DataField = 'DETAILCODENO'
        DataSource = dsrDetail
        TabOrder = 0
      end
      object StaticText4: TStaticText
        Left = 33
        Top = 19
        Width = 40
        Height = 20
        Caption = #20844#21496#21029
        TabOrder = 1
      end
      object StaticText5: TStaticText
        Left = 43
        Top = 56
        Width = 28
        Height = 20
        Caption = #20195#30908
        TabOrder = 2
      end
      object DBRadioGroup3: TDBRadioGroup
        Left = 261
        Top = 8
        Width = 185
        Height = 73
        Caption = ' '#26159#21542#20572#29992' '
        Columns = 2
        DataField = 'STOPFLAG'
        DataSource = dsrDetail
        Items.Strings = (
          #26159
          #21542)
        TabOrder = 3
        Values.Strings = (
          '1'
          '0')
      end
      object dblCompInfo: TDBLookupComboBox
        Left = 77
        Top = 16
        Width = 145
        Height = 24
        DataField = 'COMPCODE'
        DataSource = dsrDetail
        KeyField = 'CompCode'
        ListField = 'CompName'
        ListSource = dsrCompInfo
        TabOrder = 4
      end
    end
    object pnlMultiDetail: TPanel
      Left = 1
      Top = 98
      Width = 788
      Height = 149
      Align = alClient
      BevelOuter = bvLowered
      Enabled = False
      TabOrder = 1
      object dbgDetail: TDBGrid
        Left = 1
        Top = 1
        Width = 786
        Height = 147
        Align = alClient
        Color = clInfoBk
        DataSource = dsrDetail
        ImeName = #26032#27880#38899
        Options = [dgTitles, dgIndicator, dgColumnResize, dgColLines, dgRowLines, dgTabs, dgConfirmDelete, dgCancelOnExit]
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
            FieldName = 'refCompName'
            Title.Caption = #20844#21496#21029
            Width = 120
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'MASTERCODENO'
            ImeName = #20013#25991' ('#32321#39636') - '#27880#38899
            Title.Caption = #27512#23660#20195#30908
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'DETAILCODENO'
            Title.Caption = #20195#30908
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'StopFlagName'
            Title.Caption = #26159#21542#20572#29992
            Width = 60
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'OPERATOR'
            Title.Caption = #30064#21205#20154#21729
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'UPDTIME'
            Title.Caption = #30064#21205#26178#38291
            Visible = True
          end>
      end
    end
  end
  object pnlDFunction: TPanel
    Left = 0
    Top = 223
    Width = 790
    Height = 50
    Align = alTop
    BevelInner = bvLowered
    Enabled = False
    TabOrder = 3
    object spbDInsert: TSpeedButton
      Left = 104
      Top = 10
      Width = 72
      Height = 31
      Caption = #26032#22686
      NumGlyphs = 2
      OnClick = spbDInsertClick
    end
    object spbDEdit: TSpeedButton
      Left = 177
      Top = 10
      Width = 72
      Height = 31
      Caption = #20462#25913
      NumGlyphs = 2
      OnClick = spbDEditClick
    end
    object spbDDelete: TSpeedButton
      Left = 249
      Top = 10
      Width = 72
      Height = 31
      Caption = #21034#38500
      NumGlyphs = 2
      OnClick = spbDDeleteClick
    end
    object spbDCancel: TSpeedButton
      Left = 322
      Top = 10
      Width = 72
      Height = 31
      Caption = #21462#28040
      NumGlyphs = 2
      OnClick = spbDCancelClick
    end
    object spbDPost: TSpeedButton
      Left = 394
      Top = 10
      Width = 72
      Height = 31
      Caption = #23384#27284
      NumGlyphs = 2
      OnClick = spbDPostClick
    end
    object chbDContinueAppend: TCheckBox
      Left = 476
      Top = 16
      Width = 101
      Height = 26
      Caption = #36899#32396#26032#22686
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -16
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      TabOrder = 0
    end
    object cobCompInfo: TComboBox
      Left = 636
      Top = 16
      Width = 145
      Height = 24
      Style = csDropDownList
      ItemHeight = 16
      TabOrder = 1
      OnCloseUp = cobCompInfoCloseUp
    end
    object StaticText6: TStaticText
      Left = 576
      Top = 16
      Width = 52
      Height = 24
      Caption = #20844#21496#21029
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -16
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 2
    end
    object StaticText8: TStaticText
      Left = 24
      Top = 16
      Width = 72
      Height = 20
      Caption = #25152#23660#20195#30908
      Font.Charset = ANSI_CHARSET
      Font.Color = clRed
      Font.Height = -16
      Font.Name = #32048#26126#39636
      Font.Style = [fsBold]
      ParentFont = False
      TabOrder = 3
    end
  end
  object dsrDetail: TDataSource
    DataSet = dtmMain.adoCD067C
    Left = 720
    Top = 298
  end
  object dsrMaster: TDataSource
    DataSet = dtmMain.adoCD067B
    OnDataChange = dsrMasterDataChange
    Left = 704
    Top = 9
  end
  object dsrCompInfo: TDataSource
    DataSet = dtmMain.cdsCompInfo
    Left = 552
    Top = 344
  end
end
