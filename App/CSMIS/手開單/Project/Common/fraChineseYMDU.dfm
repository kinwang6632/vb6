object fraChineseYMD: TfraChineseYMD
  Left = 0
  Top = 0
  Width = 100
  Height = 35
  TabOrder = 0
  object mseYMD: TMaskEdit
    Left = 0
    Top = 8
    Width = 68
    Height = 21
    EditMask = '!999/99/99;1;_'
    MaxLength = 9
    TabOrder = 0
    Text = '   /  /  '
    OnExit = mseYMDExit
  end
end
