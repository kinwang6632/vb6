object frmSo18C5: TfrmSo18C5
  Left = 245
  Top = 211
  Width = 313
  Height = 130
  Caption = 'frmSo18C5'
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
    Width = 305
    Height = 103
    Align = alClient
    BevelInner = bvRaised
    BevelOuter = bvLowered
    TabOrder = 0
    object Label1: TLabel
      Left = 168
      Top = 24
      Width = 9
      Height = 16
      Caption = '~'
    end
    object btnQuery: TBitBtn
      Left = 58
      Top = 64
      Width = 75
      Height = 25
      Caption = #26597#35426
      TabOrder = 2
      OnClick = btnQueryClick
    end
    object btnExit: TBitBtn
      Left = 162
      Top = 64
      Width = 75
      Height = 25
      Caption = #32080#26463
      TabOrder = 3
      OnClick = btnExitClick
    end
    object StaticText3: TStaticText
      Left = 17
      Top = 24
      Width = 55
      Height = 20
      Caption = #30064#21205#26085#26399' '
      TabOrder = 4
    end
    object dtpStartDate: TDateTimePicker
      Left = 72
      Top = 21
      Width = 89
      Height = 24
      CalAlignment = dtaLeft
      Date = 37588.6523481018
      Time = 37588.6523481018
      DateFormat = dfShort
      DateMode = dmComboBox
      Kind = dtkDate
      ParseInput = False
      TabOrder = 0
    end
    object dtpEndDate: TDateTimePicker
      Left = 185
      Top = 21
      Width = 89
      Height = 24
      CalAlignment = dtaLeft
      Date = 37588.6523481018
      Time = 37588.6523481018
      DateFormat = dfShort
      DateMode = dmComboBox
      Kind = dtkDate
      ParseInput = False
      TabOrder = 1
    end
  end
  object DataSource1: TDataSource
    DataSet = dtmMain.cdsReport4
    Left = 248
    Top = 56
  end
end
