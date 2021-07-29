object frmMain: TfrmMain
  Left = 195
  Top = 117
  Width = 395
  Height = 309
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 387
    Height = 282
    Align = alClient
    BevelOuter = bvLowered
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    TabOrder = 0
    object SpeedButton1: TSpeedButton
      Left = 324
      Top = 22
      Width = 23
      Height = 22
      Caption = '...'
      OnClick = SpeedButton1Click
    end
    object BitBtn1: TBitBtn
      Left = 94
      Top = 240
      Width = 75
      Height = 25
      Caption = #35373#23450
      TabOrder = 0
      OnClick = BitBtn1Click
    end
    object CheckBox1: TCheckBox
      Left = 48
      Top = 128
      Width = 65
      Height = 17
      Caption = #38525#26126#23665
      Checked = True
      State = cbChecked
      TabOrder = 1
    end
    object CheckBox2: TCheckBox
      Left = 48
      Top = 152
      Width = 65
      Height = 17
      Caption = #26032#21488#21271
      Checked = True
      State = cbChecked
      TabOrder = 2
    end
    object CheckBox3: TCheckBox
      Left = 48
      Top = 176
      Width = 65
      Height = 17
      Caption = #22823#23433#25991#23665
      Checked = True
      State = cbChecked
      TabOrder = 3
    end
    object CheckBox4: TCheckBox
      Left = 48
      Top = 200
      Width = 65
      Height = 17
      Caption = #37329#38971#36947
      Checked = True
      State = cbChecked
      TabOrder = 4
    end
    object CheckBox5: TCheckBox
      Left = 136
      Top = 128
      Width = 65
      Height = 17
      Caption = #20840#32879
      Checked = True
      State = cbChecked
      TabOrder = 5
    end
    object CheckBox6: TCheckBox
      Left = 136
      Top = 152
      Width = 65
      Height = 17
      Caption = #22823#26032#24215
      Checked = True
      State = cbChecked
      TabOrder = 6
    end
    object CheckBox7: TCheckBox
      Left = 136
      Top = 176
      Width = 65
      Height = 17
      Caption = #26032#21776#22478
      Checked = True
      State = cbChecked
      TabOrder = 7
    end
    object CheckBox8: TCheckBox
      Left = 224
      Top = 128
      Width = 65
      Height = 17
      Caption = #25391#36947
      Checked = True
      State = cbChecked
      TabOrder = 8
    end
    object CheckBox9: TCheckBox
      Left = 224
      Top = 152
      Width = 65
      Height = 17
      Caption = #35920#30431
      Checked = True
      State = cbChecked
      TabOrder = 9
    end
    object CheckBox10: TCheckBox
      Left = 224
      Top = 176
      Width = 65
      Height = 17
      Caption = #26032#38971#36947
      Checked = True
      State = cbChecked
      TabOrder = 10
    end
    object CheckBox11: TCheckBox
      Left = 304
      Top = 128
      Width = 65
      Height = 17
      Caption = #21335#22825
      Checked = True
      State = cbChecked
      TabOrder = 11
    end
    object CheckBox12: TCheckBox
      Left = 304
      Top = 152
      Width = 65
      Height = 17
      Caption = #35264#26119
      Checked = True
      State = cbChecked
      TabOrder = 12
    end
    object CheckBox13: TCheckBox
      Left = 304
      Top = 176
      Width = 65
      Height = 17
      Caption = #23631#21335
      Checked = True
      State = cbChecked
      TabOrder = 13
    end
    object CheckBox14: TCheckBox
      Left = 136
      Top = 200
      Width = 65
      Height = 17
      Caption = #20489#31649
      Checked = True
      State = cbChecked
      TabOrder = 14
    end
    object Button1: TButton
      Left = 208
      Top = 240
      Width = 75
      Height = 25
      Caption = #32080#26463
      TabOrder = 15
      OnClick = Button1Click
    end
    object StaticText1: TStaticText
      Left = 32
      Top = 23
      Width = 40
      Height = 20
      Caption = #35373#23450#27284
      TabOrder = 16
    end
    object edtIniFileName: TEdit
      Left = 77
      Top = 21
      Width = 244
      Height = 24
      ReadOnly = True
      TabOrder = 17
    end
    object CheckBox15: TCheckBox
      Left = 224
      Top = 200
      Width = 97
      Height = 17
      Caption = #21271#26691#22290
      Checked = True
      State = cbChecked
      TabOrder = 18
    end
    object StaticText2: TStaticText
      Left = 32
      Top = 57
      Width = 112
      Height = 20
      Caption = 'callback data '#23531#20837
      TabOrder = 19
    end
    object cobCallbackDataAlias: TComboBox
      Left = 144
      Top = 53
      Width = 177
      Height = 24
      Style = csDropDownList
      ItemHeight = 16
      TabOrder = 20
      Items.Strings = (
        #21271#24179#21488
        #20013#24179#21488)
    end
    object StaticText3: TStaticText
      Left = 31
      Top = 88
      Width = 112
      Height = 20
      Caption = #26234#24935#21345#36039#26009#25152#23660#24179#21488
      TabOrder = 21
    end
    object cobIccDataAlias: TComboBox
      Left = 143
      Top = 84
      Width = 177
      Height = 24
      Style = csDropDownList
      ItemHeight = 16
      TabOrder = 22
      Items.Strings = (
        #21271#24179#21488
        #20013#24179#21488)
    end
  end
  object OpenDialog1: TOpenDialog
    Filter = '*.ini|*.ini'
    Left = 336
    Top = 56
  end
end
