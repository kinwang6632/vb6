object frmSo18C3: TfrmSo18C3
  Left = 221
  Top = 124
  Width = 542
  Height = 317
  Caption = 'frmSo18C3'
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
    Width = 534
    Height = 290
    Align = alClient
    BevelInner = bvRaised
    BevelOuter = bvLowered
    TabOrder = 0
    object Label1: TLabel
      Left = 163
      Top = 24
      Width = 9
      Height = 16
      Caption = '~'
    end
    object Label2: TLabel
      Left = 230
      Top = 74
      Width = 9
      Height = 16
      Caption = '~'
    end
    object BitBtn1: TBitBtn
      Left = 164
      Top = 246
      Width = 75
      Height = 25
      Caption = #26597#35426
      TabOrder = 0
      OnClick = BitBtn1Click
    end
    object BitBtn2: TBitBtn
      Left = 276
      Top = 246
      Width = 75
      Height = 25
      Caption = #32080#26463
      TabOrder = 1
      OnClick = BitBtn2Click
    end
    object StaticText2: TStaticText
      Left = 17
      Top = 124
      Width = 55
      Height = 20
      Caption = #38936#21934#20154#21729' '
      TabOrder = 2
    end
    object StaticText3: TStaticText
      Left = 17
      Top = 24
      Width = 55
      Height = 20
      Caption = #38936#21934#26085#26399' '
      TabOrder = 3
    end
    object StaticText4: TStaticText
      Left = 17
      Top = 73
      Width = 99
      Height = 20
      Caption = #21934#34399#31684#22285' ('#21547#23383#38957')'
      TabOrder = 4
    end
    inline fraChineseYMD1: TfraChineseYMD
      Left = 76
      Top = 16
      Width = 73
      Height = 35
      TabOrder = 5
      inherited mseYMD: TMaskEdit
        Top = 6
        Height = 24
      end
    end
    inline fraChineseYMD2: TfraChineseYMD
      Left = 191
      Top = 14
      Width = 73
      Height = 35
      TabOrder = 6
      inherited mseYMD: TMaskEdit
        Height = 24
      end
    end
    object lbxEmp: TListBox
      Left = 76
      Top = 124
      Width = 405
      Height = 97
      Columns = 4
      ItemHeight = 16
      TabOrder = 7
    end
    object edtBeginNum: TEdit
      Left = 116
      Top = 72
      Width = 109
      Height = 24
      TabOrder = 8
    end
    object edtEndNum: TEdit
      Left = 244
      Top = 72
      Width = 109
      Height = 24
      TabOrder = 9
    end
    object rgpReportType: TRadioGroup
      Left = 364
      Top = 14
      Width = 153
      Height = 105
      Caption = ' '#22577#34920' '
      ItemIndex = 0
      Items.Strings = (
        #25163#38283#21934#25818#38936#29992#26126#32048#34920
        #26410#22238#22577#25163#38283#21934#25818#28165#20874
        #25163#38283#21934#20351#29992#24773#24418)
      TabOrder = 10
    end
    object BitBtn3: TBitBtn
      Left = 489
      Top = 125
      Width = 25
      Height = 25
      Caption = '...'
      TabOrder = 11
      OnClick = BitBtn3Click
    end
  end
end
