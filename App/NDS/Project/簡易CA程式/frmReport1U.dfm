object frmReport1: TfrmReport1
  Left = 200
  Top = 172
  Width = 308
  Height = 207
  BorderIcons = [biMinimize]
  Caption = #20316#26989#27511#31243#34920
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
    Width = 300
    Height = 180
    Align = alClient
    BevelOuter = bvLowered
    TabOrder = 0
    object btnPrint: TBitBtn
      Left = 56
      Top = 144
      Width = 75
      Height = 25
      Caption = #21015#21360
      TabOrder = 4
      OnClick = btnPrintClick
    end
    object btnCancel: TBitBtn
      Left = 160
      Top = 144
      Width = 75
      Height = 25
      Caption = #21462#28040
      TabOrder = 5
      OnClick = btnCancelClick
    end
    object RadioGroup1: TRadioGroup
      Left = 19
      Top = 11
      Width = 262
      Height = 114
      Caption = ' '#21015#21360#26781#20214' '
      ItemIndex = 0
      Items.Strings = (
        'Subscriber ID'
        'ICC NO'
        #20316#26989#26085#26399)
      TabOrder = 0
      OnClick = RadioGroup1Click
    end
    object edtSubscruberID: TEdit
      Left = 139
      Top = 27
      Width = 121
      Height = 24
      TabOrder = 1
    end
    object edtIccNo: TEdit
      Left = 139
      Top = 59
      Width = 121
      Height = 24
      TabOrder = 2
    end
    object DateTimePicker1: TDateTimePicker
      Left = 139
      Top = 91
      Width = 121
      Height = 24
      CalAlignment = dtaLeft
      Date = 37687.4899206134
      Time = 37687.4899206134
      DateFormat = dfShort
      DateMode = dmComboBox
      Kind = dtkDate
      ParseInput = False
      TabOrder = 3
    end
  end
end
