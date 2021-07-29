object frmSO8C20: TfrmSO8C20
  Left = 324
  Top = 199
  Width = 502
  Height = 276
  Caption = 'frmSO8C20'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -13
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnClose = FormClose
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 16
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 494
    Height = 248
    Align = alClient
    BevelInner = bvRaised
    BevelOuter = bvLowered
    TabOrder = 0
    object StaticText1: TStaticText
      Left = 8
      Top = 25
      Width = 96
      Height = 20
      Caption = #20316#24290#21934#34399'('#21547#23383#38957')'
      TabOrder = 5
    end
    object edtPaperSNum: TEdit
      Left = 104
      Top = 22
      Width = 121
      Height = 24
      TabOrder = 0
    end
    object BitBtn1: TBitBtn
      Left = 288
      Top = 200
      Width = 75
      Height = 25
      Caption = #20316#24290
      TabOrder = 3
      OnClick = BitBtn1Click
    end
    object BitBtn2: TBitBtn
      Left = 400
      Top = 200
      Width = 75
      Height = 25
      Caption = #32080#26463
      TabOrder = 4
      OnClick = BitBtn2Click
    end
    object btnQuery: TBitBtn
      Left = 371
      Top = 22
      Width = 62
      Height = 25
      Caption = #26597#35426
      TabOrder = 2
      OnClick = btnQueryClick
    end
    object StaticText7: TStaticText
      Left = 232
      Top = 25
      Width = 13
      Height = 20
      Caption = '~'
      TabOrder = 6
    end
    object edtPaperENum: TEdit
      Left = 248
      Top = 22
      Width = 121
      Height = 24
      TabOrder = 1
    end
    object dbgPaperData: TDBGrid
      Left = 8
      Top = 54
      Width = 481
      Height = 139
      DataSource = dsrPaperData
      ReadOnly = True
      TabOrder = 7
      TitleFont.Charset = DEFAULT_CHARSET
      TitleFont.Color = clWindowText
      TitleFont.Height = -13
      TitleFont.Name = 'MS Sans Serif'
      TitleFont.Style = []
    end
  end
  object dsrPaperData: TDataSource
    Left = 432
    Top = 16
  end
end
