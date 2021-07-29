object fraYM: TfraYM
  Left = 0
  Top = 0
  Width = 100
  Height = 36
  TabOrder = 0
  object mseYM: TMaskEdit
    Left = 8
    Top = 8
    Width = 57
    Height = 21
    EditMask = '!9999/99;1'
    MaxLength = 7
    TabOrder = 0
    Text = '    /  '
    OnExit = mseYMExit
  end
end
