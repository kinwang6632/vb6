object fraDBLookupCob: TfraDBLookupCob
  Left = 0
  Top = 0
  Width = 400
  Height = 25
  TabOrder = 0
  object DBLookupComboBox1: TDBLookupComboBox
    Left = 2
    Top = 2
    Width = 326
    Height = 21
    ImeName = '中文 (繁體) - 新注音'
    KeyField = 'sVesselID'
    ListField = 'sVesselID;sVesselName'
    TabOrder = 0
  end
  object CheckBox1: TCheckBox
    Left = 331
    Top = 4
    Width = 66
    Height = 17
    Caption = 'By Desc'
    TabOrder = 1
    OnClick = CheckBox1Click
  end
end
