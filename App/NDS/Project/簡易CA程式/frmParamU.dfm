object frmParam: TfrmParam
  Left = 348
  Top = 168
  Width = 436
  Height = 327
  Caption = #21443#25976#35373#23450
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clTeal
  Font.Height = -13
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  FormStyle = fsStayOnTop
  OldCreateOrder = False
  Position = poScreenCenter
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 16
  object BitBtn1: TBitBtn
    Left = 328
    Top = 270
    Width = 75
    Height = 25
    Caption = #38364#38281#35222#31383
    TabOrder = 3
    OnClick = BitBtn1Click
  end
  object StaticText2: TStaticText
    Left = 16
    Top = 144
    Width = 166
    Height = 20
    Caption = 'Reportback Availability'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clTeal
    Font.Height = -13
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 4
  end
  object ComboBox1: TComboBox
    Left = 200
    Top = 142
    Width = 201
    Height = 24
    Style = csDropDownList
    ItemHeight = 16
    ItemIndex = 2
    TabOrder = 0
    Text = 'A (00:00 ~ 23:59. local time)'
    Items.Strings = (
      'D (05:00 ~ 23:00. local time)'
      'E (23:00 ~ 05:00. local time)'
      'A (00:00 ~ 23:59. local time)'
      'N (None)')
  end
  object StaticText3: TStaticText
    Left = 16
    Top = 184
    Width = 122
    Height = 20
    Caption = 'Reportback Date'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clTeal
    Font.Height = -13
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 5
  end
  object ComboBox2: TComboBox
    Left = 160
    Top = 182
    Width = 145
    Height = 24
    Style = csDropDownList
    ItemHeight = 16
    ItemIndex = 0
    TabOrder = 1
    Text = '1'
    Items.Strings = (
      '1'
      '2'
      '3'
      '4'
      '5'
      '6'
      '7'
      '8'
      '9'
      '10'
      '11'
      '12'
      '13'
      '14'
      '15'
      '16'
      '17'
      '18'
      '19'
      '20'
      '21'
      '22'
      '23'
      '24'
      '25'
      '26'
      '27'
      '28')
  end
  object StaticText4: TStaticText
    Left = 16
    Top = 230
    Width = 261
    Height = 20
    Caption = 'Population ID ('#26368#22810'2'#20491'bytes. '#35531#22635#20837#25976#23383')'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clTeal
    Font.Height = -13
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 6
  end
  object edtPopulationID: TEdit
    Left = 283
    Top = 226
    Width = 121
    Height = 24
    MaxLength = 2
    TabOrder = 2
    Text = '1'
    OnExit = edtPopulationIDExit
  end
  object GroupBox1: TGroupBox
    Left = 16
    Top = 8
    Width = 297
    Height = 117
    Caption = ' CA Command Gateway'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlue
    Font.Height = -13
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    TabOrder = 7
    object StaticText1: TStaticText
      Left = 54
      Top = 28
      Width = 18
      Height = 20
      Caption = 'IP'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlue
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      TabOrder = 0
    end
    object StaticText5: TStaticText
      Left = 38
      Top = 55
      Width = 32
      Height = 20
      Caption = 'Port'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlue
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      TabOrder = 1
    end
    object edtIP: TEdit
      Left = 72
      Top = 24
      Width = 121
      Height = 24
      TabOrder = 2
    end
    object edtPort: TEdit
      Left = 72
      Top = 53
      Width = 121
      Height = 24
      TabOrder = 3
      OnExit = edtPortExit
    end
    object StaticText6: TStaticText
      Left = 12
      Top = 83
      Width = 60
      Height = 20
      Caption = 'Timeout'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlue
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      TabOrder = 4
    end
    object edtTimeout: TEdit
      Left = 72
      Top = 83
      Width = 121
      Height = 24
      TabOrder = 5
      OnExit = edtTimeoutExit
    end
    object StaticText7: TStaticText
      Left = 204
      Top = 83
      Width = 16
      Height = 20
      Caption = #31186
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlue
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      TabOrder = 6
    end
  end
end
