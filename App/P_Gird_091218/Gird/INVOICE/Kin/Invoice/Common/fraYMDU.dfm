object fraYMD: TfraYMD
  Left = 0
  Top = 0
  Width = 104
  Height = 35
  TabOrder = 0
  object mseYMD: TMaskEdit
    Left = 0
    Top = 8
    Width = 70
    Height = 21
    EditMask = '!9999/99/99;1;_'
    MaxLength = 10
    TabOrder = 0
    Text = '    /  /  '
    OnExit = mseYMDExit
  end
end
