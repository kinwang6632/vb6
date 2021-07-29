object frmNightRunLog: TfrmNightRunLog
  Left = 15
  Top = 0
  Width = 788
  Height = 612
  AutoSize = True
  Caption = 'frmNightRunLog'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 780
    Height = 41
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
    object btnExit: TBitBtn
      Left = 700
      Top = 9
      Width = 59
      Height = 24
      Caption = #32080#26463
      Font.Charset = ANSI_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
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
  end
  object Panel2: TPanel
    Left = 0
    Top = 41
    Width = 780
    Height = 115
    Align = alTop
    TabOrder = 1
    object Panel4: TPanel
      Left = 1
      Top = 1
      Width = 400
      Height = 113
      Align = alLeft
      BevelWidth = 3
      BorderStyle = bsSingle
      TabOrder = 0
      object lblComp: TLabel
        Left = 7
        Top = 23
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
      object lblQueryDate: TLabel
        Left = 15
        Top = 67
        Width = 26
        Height = 13
        Alignment = taRightJustify
        Caption = #26085#26399
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = #26032#32048#26126#39636
        Font.Style = []
        ParentFont = False
      end
      object Label5: TLabel
        Left = 177
        Top = 69
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
      object lblQueryType: TLabel
        Left = 173
        Top = 24
        Width = 52
        Height = 13
        Alignment = taRightJustify
        Caption = #26597#35426#31278#39006
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = #26032#32048#26126#39636
        Font.Style = []
        ParentFont = False
      end
      object cobCompCode: TComboBox
        Left = 56
        Top = 18
        Width = 113
        Height = 21
        ItemHeight = 13
        TabOrder = 0
        Text = 'cobCompCode'
      end
      inline fraYMD1: TfraYMD
        Left = 56
        Top = 57
        Width = 79
        Height = 35
        TabOrder = 2
        OnExit = fraYMD1Exit
      end
      inline fraYMD2: TfraYMD
        Left = 200
        Top = 57
        Width = 98
        Height = 35
        TabOrder = 4
      end
      object btnOK: TBitBtn
        Left = 324
        Top = 67
        Width = 59
        Height = 24
        Caption = #26032#22686
        TabOrder = 6
        OnClick = btnOKClick
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
      object medETime: TMaskEdit
        Left = 273
        Top = 65
        Width = 39
        Height = 24
        EditMask = '!90:00;0; '
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        MaxLength = 5
        ParentFont = False
        TabOrder = 5
      end
      object medSTime: TMaskEdit
        Left = 128
        Top = 65
        Width = 37
        Height = 24
        EditMask = '!90:00;0; '
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        MaxLength = 5
        ParentFont = False
        TabOrder = 3
      end
      object cobQueryType: TComboBox
        Left = 232
        Top = 18
        Width = 129
        Height = 21
        ItemHeight = 13
        TabOrder = 1
        Text = 'cobQueryType'
      end
    end
    object Panel5: TPanel
      Left = 401
      Top = 1
      Width = 376
      Height = 113
      BevelWidth = 3
      BorderStyle = bsSingle
      TabOrder = 1
      object lblComp2: TLabel
        Left = 7
        Top = 37
        Width = 39
        Height = 13
        Alignment = taRightJustify
        Caption = #20844#21496#21029
        Font.Charset = ANSI_CHARSET
        Font.Color = clBlue
        Font.Height = -13
        Font.Name = #26032#32048#26126#39636
        Font.Style = []
        ParentFont = False
      end
      object lblQueryType2: TLabel
        Left = 173
        Top = 38
        Width = 52
        Height = 13
        Alignment = taRightJustify
        Caption = #26597#35426#31278#39006
        Font.Charset = ANSI_CHARSET
        Font.Color = clBlue
        Font.Height = -13
        Font.Name = #26032#32048#26126#39636
        Font.Style = []
        ParentFont = False
      end
      object lblQueryDate2: TLabel
        Left = 15
        Top = 69
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
      object lblActionAgain: TLabel
        Left = 9
        Top = 9
        Width = 39
        Height = 13
        Caption = #20844#21496#21029
        Font.Charset = ANSI_CHARSET
        Font.Color = clBlue
        Font.Height = -13
        Font.Name = #26032#32048#26126#39636
        Font.Style = []
        ParentFont = False
      end
      object Label1: TLabel
        Left = 133
        Top = 69
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
      object cobCompCode2: TComboBox
        Left = 56
        Top = 32
        Width = 113
        Height = 21
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlue
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ItemHeight = 13
        ParentFont = False
        TabOrder = 0
        Text = 'cobCompCode'
      end
      object cobQueryType2: TComboBox
        Left = 233
        Top = 32
        Width = 128
        Height = 21
        Font.Charset = ANSI_CHARSET
        Font.Color = clBlue
        Font.Height = -13
        Font.Name = #26032#32048#26126#39636
        Font.Style = []
        ItemHeight = 13
        ParentFont = False
        TabOrder = 1
        Text = 'cobQueryType'
      end
      inline fraYMD3: TfraYMD
        Left = 56
        Top = 59
        Width = 73
        Height = 36
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlue
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 2
        OnExit = fraYMD3Exit
      end
      object btnAction: TBitBtn
        Left = 232
        Top = 61
        Width = 129
        Height = 24
        Caption = #26032#22686
        Font.Charset = ANSI_CHARSET
        Font.Color = clBlue
        Font.Height = -13
        Font.Name = #26032#32048#26126#39636
        Font.Style = []
        ParentFont = False
        TabOrder = 3
        OnClick = btnActionClick
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
      inline fraYMD4: TfraYMD
        Left = 146
        Top = 57
        Width = 71
        Height = 35
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlue
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 4
      end
    end
  end
  object Panel3: TPanel
    Left = 0
    Top = 156
    Width = 780
    Height = 429
    Align = alClient
    Caption = 'Panel3'
    TabOrder = 2
    object dbgNightRunLog: TDBGrid
      Left = 1
      Top = 1
      Width = 778
      Height = 338
      Align = alClient
      DataSource = dsrNightRunLod
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clRed
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 0
      TitleFont.Charset = DEFAULT_CHARSET
      TitleFont.Color = clWindowText
      TitleFont.Height = -11
      TitleFont.Name = 'MS Sans Serif'
      TitleFont.Style = []
      OnDblClick = dbgNightRunLogDblClick
    end
    object memErrorLog: TMemo
      Left = 1
      Top = 339
      Width = 778
      Height = 89
      Align = alBottom
      TabOrder = 1
    end
  end
  object dsrNightRunLod: TDataSource
    DataSet = dtmMain.adoQryNightRunLog
    Left = 104
    Top = 200
  end
end
