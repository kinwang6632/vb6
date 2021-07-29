object fraChineseYM: TfraChineseYM
  Left = 0
  Top = 0
  Width = 109
  Height = 35
  TabOrder = 0
  object mseYM: TMaskEdit
    Left = 0
    Top = 8
    Width = 64
    Height = 21
    EditMask = '!9999/99;1;_'
    MaxLength = 7
    TabOrder = 0
    Text = '    /  '
    OnExit = mseYMExit
  end
end
