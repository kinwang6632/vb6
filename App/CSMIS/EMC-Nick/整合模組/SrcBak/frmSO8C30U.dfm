object frmSO8C30: TfrmSO8C30
  Left = 341
  Top = 224
  Width = 550
  Height = 420
  Caption = 'frmSO8C30'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -13
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poMainFormCenter
  OnClose = FormClose
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 16
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 542
    Height = 392
    Align = alClient
    BevelInner = bvRaised
    BevelOuter = bvLowered
    Font.Charset = ANSI_CHARSET
    Font.Color = clWindowText
    Font.Height = -12
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
    TabOrder = 0
    object Label1: TLabel
      Left = 163
      Top = 24
      Width = 9
      Height = 14
      Caption = '~'
    end
    object Label2: TLabel
      Left = 189
      Top = 58
      Width = 9
      Height = 14
      Caption = '~'
    end
    object Label3: TLabel
      Left = 19
      Top = 51
      Width = 52
      Height = 28
      Caption = #21934#34399#31684#22285' '#13#10'('#21547#23383#38957')'
    end
    object Label15: TLabel
      Left = 81
      Top = 304
      Width = 87
      Height = 13
      AutoSize = False
      Caption = 'Excel'#20786#23384#36335#24465':'
    end
    object BitBtn1: TBitBtn
      Left = 175
      Top = 348
      Width = 75
      Height = 25
      Caption = #26597#35426
      TabOrder = 8
      OnClick = BitBtn1Click
    end
    object BitBtn2: TBitBtn
      Left = 287
      Top = 348
      Width = 75
      Height = 25
      Caption = #32080#26463
      TabOrder = 9
      OnClick = BitBtn2Click
    end
    object StaticText2: TStaticText
      Left = 17
      Top = 131
      Width = 56
      Height = 18
      Caption = #38936#21934#20154#21729' '
      TabOrder = 10
    end
    object StaticText3: TStaticText
      Left = 17
      Top = 24
      Width = 56
      Height = 18
      Caption = #38936#21934#26085#26399' '
      TabOrder = 11
    end
    inline fraChineseYMD1: TfraChineseYMD
      Left = 76
      Top = 16
      Width = 83
      Height = 36
      TabOrder = 0
      inherited mseYMD: TMaskEdit
        Top = 6
        Height = 22
      end
    end
    inline fraChineseYMD2: TfraChineseYMD
      Left = 191
      Top = 14
      Width = 84
      Height = 37
      TabOrder = 1
      inherited mseYMD: TMaskEdit
        Height = 22
      end
    end
    object lbxEmp: TListBox
      Left = 76
      Top = 131
      Width = 405
      Height = 97
      Columns = 4
      ItemHeight = 14
      TabOrder = 4
    end
    object edtBeginNum: TEdit
      Left = 77
      Top = 55
      Width = 109
      Height = 22
      TabOrder = 2
    end
    object edtEndNum: TEdit
      Left = 201
      Top = 54
      Width = 109
      Height = 22
      TabOrder = 3
    end
    object rgpReportType: TRadioGroup
      Left = 340
      Top = 10
      Width = 169
      Height = 110
      Caption = ' '#22577#34920' '
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -12
      Font.Name = 'Tahoma'
      Font.Style = []
      ItemIndex = 0
      Items.Strings = (
        #25163#38283#21934#25818#38936#29992#26126#32048#34920
        #26410#22238#22577#25163#38283#21934#25818#28165#20874
        #25163#38283#21934#20351#29992#24773#24418
        #38936#21934#32396#29992#32000#37636#27284)
      ParentFont = False
      TabOrder = 12
      OnClick = rgpReportTypeClick
    end
    object BitBtn3: TBitBtn
      Left = 485
      Top = 132
      Width = 25
      Height = 25
      Caption = '...'
      TabOrder = 5
      OnClick = BitBtn3Click
    end
    object rgpType2: TRadioGroup
      Left = 77
      Top = 238
      Width = 188
      Height = 51
      Caption = '  '#22577#34920#31278#39006'  '
      Columns = 3
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = #32048#26126#39636
      Font.Style = []
      ItemIndex = 0
      Items.Strings = (
        #22577#34920
        'Excel')
      ParentFont = False
      TabOrder = 6
    end
    object edtPath1: TEdit
      Left = 174
      Top = 300
      Width = 303
      Height = 21
      TabOrder = 7
    end
  end
end
