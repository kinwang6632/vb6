object frmInvA09: TfrmInvA09
  Left = 286
  Top = 154
  Width = 401
  Height = 497
  Caption = 'frmInvA09'
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
  object Label3: TLabel
    Left = 25
    Top = 34
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
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 393
    Height = 32
    Align = alTop
    Color = clSkyBlue
    TabOrder = 0
    object Label6: TLabel
      Left = 10
      Top = 8
      Width = 102
      Height = 16
      Caption = #30332#31080#36039#26009#27298#26680
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
    Width = 393
    Height = 438
    Align = alClient
    TabOrder = 1
    object btnExit: TBitBtn
      Left = 240
      Top = 402
      Width = 71
      Height = 26
      Caption = #32080#26463
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
      TabOrder = 4
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
    object btnQuery: TBitBtn
      Left = 85
      Top = 402
      Width = 71
      Height = 26
      Caption = #26597#35426
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
      TabOrder = 3
      OnClick = btnQueryClick
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
    object Panel3: TPanel
      Left = 9
      Top = 6
      Width = 377
      Height = 91
      TabOrder = 0
      object Label2: TLabel
        Left = 41
        Top = 66
        Width = 65
        Height = 13
        Caption = #30332#31080#24180#26376#65306
        Font.Charset = CHINESEBIG5_CHARSET
        Font.Color = clTeal
        Font.Height = -13
        Font.Name = #26032#32048#26126#39636
        Font.Style = []
        ParentFont = False
        Transparent = True
      end
      object Label10: TLabel
        Left = 206
        Top = 64
        Width = 65
        Height = 13
        Caption = #30332#31080#23383#36556#65306
        Font.Charset = CHINESEBIG5_CHARSET
        Font.Color = clTeal
        Font.Height = -13
        Font.Name = #26032#32048#26126#39636
        Font.Style = []
        ParentFont = False
        Transparent = True
      end
      inline fraInvYM: TfraYM
        Left = 108
        Top = 56
        Width = 98
        Height = 30
        TabOrder = 3
        inherited mseYM: TMaskEdit
          Top = 3
          Width = 65
        end
      end
      object medPrefix: TMaskEdit
        Left = 270
        Top = 60
        Width = 57
        Height = 21
        EditMask = 'aa;1;_'
        Font.Charset = CHINESEBIG5_CHARSET
        Font.Color = clBlack
        Font.Height = -13
        Font.Name = #26032#32048#26126#39636
        Font.Style = []
        MaxLength = 2
        ParentFont = False
        TabOrder = 4
        Text = '  '
        OnExit = medPrefixExit
      end
      object chkBusinessID: TCheckBox
        Left = 40
        Top = 34
        Width = 121
        Height = 17
        Caption = #32113#19968#32232#34399#27298#26680
        Font.Charset = CHINESEBIG5_CHARSET
        Font.Color = clNavy
        Font.Height = -13
        Font.Name = #26032#32048#26126#39636
        Font.Style = []
        ParentFont = False
        TabOrder = 1
      end
      object chkJump: TCheckBox
        Left = 168
        Top = 34
        Width = 121
        Height = 17
        Caption = #36339#34399#27298#26680
        Font.Charset = CHINESEBIG5_CHARSET
        Font.Color = clNavy
        Font.Height = -13
        Font.Name = #26032#32048#26126#39636
        Font.Style = []
        ParentFont = False
        TabOrder = 2
        OnClick = chkJumpClick
      end
      object chkNumber: TCheckBox
        Left = 16
        Top = 12
        Width = 137
        Height = 17
        Caption = #32113#32232#27298#26680
        Font.Charset = CHINESEBIG5_CHARSET
        Font.Color = clMaroon
        Font.Height = -13
        Font.Name = #26032#32048#26126#39636
        Font.Style = [fsBold]
        ParentFont = False
        TabOrder = 0
      end
    end
    object Panel4: TPanel
      Left = 9
      Top = 100
      Width = 377
      Height = 87
      Caption = 'Panel4'
      TabOrder = 1
      object Label1: TLabel
        Left = 25
        Top = 34
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
      object chkAmount: TCheckBox
        Left = 8
        Top = 8
        Width = 137
        Height = 17
        Caption = #20027#21103#27284#37329#38989#27298#26597
        Font.Charset = CHINESEBIG5_CHARSET
        Font.Color = clMaroon
        Font.Height = -13
        Font.Name = #26032#32048#26126#39636
        Font.Style = [fsBold]
        ParentFont = False
        TabOrder = 0
      end
      inline fraEInvDate1: TfraYMD
        Left = 237
        Top = 20
        Width = 111
        Height = 31
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -16
        Font.Name = #26032#32048#26126#39636
        Font.Style = []
        ParentFont = False
        TabOrder = 2
        inherited mseYMD: TMaskEdit
          Top = 6
          Width = 105
          Font.Charset = CHINESEBIG5_CHARSET
          Font.Color = clBlack
          Font.Height = -13
          Font.Name = #26032#32048#26126#39636
          ParentFont = False
        end
      end
      inline fraSInvDate1: TfraYMD
        Left = 101
        Top = 24
        Width = 111
        Height = 33
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -16
        Font.Name = #26032#32048#26126#39636
        Font.Style = []
        ParentFont = False
        TabOrder = 1
        OnExit = fraSInvDate1Exit
        inherited mseYMD: TMaskEdit
          Top = 3
          Width = 101
          Font.Charset = CHINESEBIG5_CHARSET
          Font.Color = clBlack
          Font.Height = -13
          Font.Name = #26032#32048#26126#39636
          ParentFont = False
        end
      end
      object chkCoverObsolete: TCheckBox
        Left = 24
        Top = 60
        Width = 137
        Height = 17
        Caption = #21253#21547#20316#24290#30332#31080
        Font.Charset = CHINESEBIG5_CHARSET
        Font.Color = clTeal
        Font.Height = -13
        Font.Name = #26032#32048#26126#39636
        Font.Style = []
        ParentFont = False
        TabOrder = 3
      end
    end
    object Panel5: TPanel
      Left = 9
      Top = 190
      Width = 377
      Height = 203
      BevelInner = bvRaised
      TabOrder = 2
      object Label4: TLabel
        Left = 25
        Top = 170
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
      object chkMisData: TCheckBox
        Left = 8
        Top = 12
        Width = 177
        Height = 17
        Caption = #30332#31080#38283#31435#33287#23458#26381#36039#26009#27604#23565
        Font.Charset = CHINESEBIG5_CHARSET
        Font.Color = clMaroon
        Font.Height = -13
        Font.Name = #26032#32048#26126#39636
        Font.Style = [fsBold]
        ParentFont = False
        TabOrder = 0
      end
      object chkMisData1: TCheckBox
        Left = 40
        Top = 40
        Width = 321
        Height = 17
        Caption = #30332#31080#25910#20214#20154#33287#23458#26381#20027#27284#22995#21517#19981#21516#25110#33287#24115#34399#25910#20214#20154#19981#21516
        Font.Charset = CHINESEBIG5_CHARSET
        Font.Color = clNavy
        Font.Height = -13
        Font.Name = #26032#32048#26126#39636
        Font.Style = []
        ParentFont = False
        TabOrder = 1
      end
      object chkMisData2: TCheckBox
        Left = 40
        Top = 64
        Width = 321
        Height = 17
        Caption = #25033#28858#19977#32879#38283#25104#20108#32879'('#26377#24115#34399')'#25110#25033#28858#19977#32879#38283#25104#20108#32879
        Font.Charset = CHINESEBIG5_CHARSET
        Font.Color = clNavy
        Font.Height = -13
        Font.Name = #26032#32048#26126#39636
        Font.Style = []
        ParentFont = False
        TabOrder = 2
      end
      object chkMisData4: TCheckBox
        Left = 40
        Top = 111
        Width = 321
        Height = 17
        Caption = #30332#31080#25260#38957#33287#23458#26381#25152#24314#20043#30332#31080#25260#38957#19981#21516
        Font.Charset = CHINESEBIG5_CHARSET
        Font.Color = clNavy
        Font.Height = -13
        Font.Name = #26032#32048#26126#39636
        Font.Style = []
        ParentFont = False
        TabOrder = 4
      end
      object chkMisData5: TCheckBox
        Left = 40
        Top = 136
        Width = 321
        Height = 17
        Caption = #30332#31080#38283#31435#20043#32113#32232#33287#23458#26381#31995#32113#25152#24314#20043#32113#32232#19981#21516
        Font.Charset = CHINESEBIG5_CHARSET
        Font.Color = clNavy
        Font.Height = -13
        Font.Name = #26032#32048#26126#39636
        Font.Style = []
        ParentFont = False
        TabOrder = 5
      end
      inline fraSInvDate2: TfraYMD
        Left = 101
        Top = 160
        Width = 111
        Height = 33
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -16
        Font.Name = #26032#32048#26126#39636
        Font.Style = []
        ParentFont = False
        TabOrder = 6
        OnExit = fraSInvDate2Exit
        inherited mseYMD: TMaskEdit
          Top = 3
          Width = 101
          Font.Charset = CHINESEBIG5_CHARSET
          Font.Color = clBlack
          Font.Height = -13
          Font.Name = #26032#32048#26126#39636
          ParentFont = False
        end
      end
      inline fraEInvDate2: TfraYMD
        Left = 237
        Top = 157
        Width = 111
        Height = 31
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -16
        Font.Name = #26032#32048#26126#39636
        Font.Style = []
        ParentFont = False
        TabOrder = 7
        inherited mseYMD: TMaskEdit
          Top = 6
          Width = 105
          Font.Charset = CHINESEBIG5_CHARSET
          Font.Color = clBlack
          Font.Height = -13
          Font.Name = #26032#32048#26126#39636
          ParentFont = False
        end
      end
      object chkMisData3: TCheckBox
        Left = 40
        Top = 88
        Width = 321
        Height = 17
        Caption = #22810#20491#23458#25142#38283#31435#22312#21516#19968#24373#30332#31080#19978
        Font.Charset = CHINESEBIG5_CHARSET
        Font.Color = clNavy
        Font.Height = -13
        Font.Name = #26032#32048#26126#39636
        Font.Style = []
        ParentFont = False
        TabOrder = 3
      end
    end
  end
end
