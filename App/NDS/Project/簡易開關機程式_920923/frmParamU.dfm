object frmParam: TfrmParam
  Left = 348
  Top = 210
  Width = 436
  Height = 272
  Caption = #21443#25976#35373#23450
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clTeal
  Font.Height = -13
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  FormStyle = fsStayOnTop
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 16
  object BitBtn1: TBitBtn
    Left = 328
    Top = 192
    Width = 75
    Height = 25
    Caption = #38364#38281#35222#31383
    TabOrder = 4
    OnClick = BitBtn1Click
  end
  object StaticText1: TStaticText
    Left = 16
    Top = 32
    Width = 264
    Height = 20
    Caption = 'Region Key ('#26368#22810'8'#20491'bytes. '#25976#23383','#23383#20803#22343#21487')'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clTeal
    Font.Height = -13
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 5
  end
  object Edit1: TEdit
    Left = 280
    Top = 29
    Width = 121
    Height = 24
    MaxLength = 8
    TabOrder = 0
    Text = '12345678'
  end
  object StaticText2: TStaticText
    Left = 16
    Top = 66
    Width = 166
    Height = 20
    Caption = 'Reportback Availability'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clTeal
    Font.Height = -13
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 6
  end
  object ComboBox1: TComboBox
    Left = 200
    Top = 64
    Width = 201
    Height = 24
    Style = csDropDownList
    ItemHeight = 16
    ItemIndex = 2
    TabOrder = 1
    Text = 'A (00:00 ~ 23:59. local time)'
    Items.Strings = (
      'D (05:00 ~ 23:00. local time)'
      'E (23:00 ~ 05:00. local time)'
      'A (00:00 ~ 23:59. local time)'
      'N (None)')
  end
  object StaticText3: TStaticText
    Left = 16
    Top = 106
    Width = 122
    Height = 20
    Caption = 'Reportback Date'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clTeal
    Font.Height = -13
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 7
  end
  object ComboBox2: TComboBox
    Left = 160
    Top = 104
    Width = 145
    Height = 24
    Style = csDropDownList
    ItemHeight = 16
    ItemIndex = 0
    TabOrder = 2
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
    Top = 152
    Width = 261
    Height = 20
    Caption = 'Population ID ('#26368#22810'2'#20491'bytes. '#35531#22635#20837#25976#23383')'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clTeal
    Font.Height = -13
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 8
  end
  object Edit2: TEdit
    Left = 283
    Top = 148
    Width = 121
    Height = 24
    MaxLength = 2
    TabOrder = 3
    Text = '99'
    OnExit = Edit2Exit
  end
end
