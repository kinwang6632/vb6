object frmSo8C60: TfrmSo8C60
  Left = 420
  Top = 214
  Width = 390
  Height = 215
  Caption = 'frmSo8C60'
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
    Width = 382
    Height = 187
    Align = alClient
    BevelInner = bvLowered
    BevelOuter = bvNone
    TabOrder = 0
    object lblCounts: TLabel
      Left = 269
      Top = 61
      Width = 65
      Height = 16
      Caption = 'lblCounts'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clNavy
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object lblEndNumber: TLabel
      Left = 125
      Top = 117
      Width = 99
      Height = 16
      Caption = 'lblEndNumber'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clNavy
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object lblBeginNumber: TLabel
      Left = 125
      Top = 93
      Width = 112
      Height = 16
      Caption = 'lblBeginNumber'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clNavy
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
    inline fraChineseYMD3: TfraChineseYMD
      Left = 80
      Top = 22
      Width = 78
      Height = 31
      TabOrder = 0
      inherited mseYMD: TMaskEdit
        Top = 3
        Height = 24
      end
    end
    object StaticText5: TStaticText
      Left = 25
      Top = 28
      Width = 52
      Height = 20
      Caption = #38936#21934#26085#26399
      TabOrder = 5
    end
    object StaticText6: TStaticText
      Left = 169
      Top = 28
      Width = 52
      Height = 20
      Caption = #38936#21934#20154#21729
      TabOrder = 6
    end
    object btnGetPaperUser: TButton
      Left = 349
      Top = 23
      Width = 23
      Height = 25
      Caption = '...'
      TabOrder = 1
      OnClick = btnGetPaperUserClick
    end
    object StaticText7: TStaticText
      Left = 24
      Top = 62
      Width = 64
      Height = 20
      Caption = #36215#22987#27969#27700#34399
      TabOrder = 7
    end
    object edtName: TEdit
      Left = 224
      Top = 24
      Width = 121
      Height = 24
      ReadOnly = True
      TabOrder = 8
    end
    object StaticText9: TStaticText
      Left = 234
      Top = 60
      Width = 28
      Height = 20
      Caption = #24373#25976
      TabOrder = 9
    end
    object edtBeginNumber: TEdit
      Left = 100
      Top = 59
      Width = 121
      Height = 24
      TabOrder = 2
      OnExit = edtBeginNumberExit
    end
    object StaticText1: TStaticText
      Left = 24
      Top = 116
      Width = 88
      Height = 20
      Caption = #21407#22987#25130#27490#27969#27700#34399
      TabOrder = 10
    end
    object BitBtn1: TBitBtn
      Left = 102
      Top = 152
      Width = 75
      Height = 25
      Caption = #30906#23450
      TabOrder = 3
      OnClick = BitBtn1Click
    end
    object BitBtn2: TBitBtn
      Left = 206
      Top = 152
      Width = 75
      Height = 25
      Caption = #32080#26463
      TabOrder = 4
      OnClick = BitBtn2Click
    end
    object StaticText2: TStaticText
      Left = 24
      Top = 93
      Width = 88
      Height = 20
      Caption = #21407#22987#36215#22987#27969#27700#34399
      TabOrder = 11
    end
  end
end
