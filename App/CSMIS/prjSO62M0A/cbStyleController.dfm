object StyleController: TStyleController
  OldCreateOrder = False
  Left = 438
  Top = 309
  Height = 262
  Width = 294
  object EditorStyle: TcxEditStyleController
    Style.Font.Charset = ANSI_CHARSET
    Style.Font.Color = clWindowText
    Style.Font.Height = -12
    Style.Font.Name = 'Tahoma'
    Style.Font.Style = []
    Style.HotTrack = False
    Style.LookAndFeel.Kind = lfFlat
    Style.TextColor = clBlack
    Style.TextStyle = []
    Style.IsFontAssigned = True
    StyleDisabled.Color = 15659506
    StyleDisabled.LookAndFeel.Kind = lfFlat
    StyleFocused.BorderColor = clBlue
    StyleFocused.BorderStyle = ebsThick
    StyleFocused.Color = 15400959
    StyleFocused.LookAndFeel.Kind = lfFlat
    StyleHot.BorderStyle = ebsFlat
    StyleHot.LookAndFeel.Kind = lfFlat
    Left = 60
    Top = 24
  end
  object CheckBoxStyle: TcxEditStyleController
    Style.BorderStyle = ebsFlat
    Style.HotTrack = False
    StyleFocused.BorderColor = clBlue
    StyleFocused.BorderStyle = ebsThick
    StyleHot.BorderStyle = ebsFlat
    Left = 172
    Top = 20
  end
  object GridStyleRepository: TcxStyleRepository
    Left = 60
    Top = 92
    PixelsPerInch = 96
    object HeaderStyle: TcxStyle
      AssignedValues = [svColor]
      Color = 16777156
    end
    object RowActiveStyle: TcxStyle
      AssignedValues = [svColor, svTextColor]
      Color = clMenuHighlight
      TextColor = clWindowText
    end
    object GridBackgroundStyle: TcxStyle
      AssignedValues = [svColor]
      Color = 15400959
    end
    object GroupStyle: TcxStyle
      AssignedValues = [svFont]
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = [fsBold]
    end
  end
  object HintStyleController: TcxHintStyleController
    HintStyle.CallOutPosition = cxbpAuto
    HintStyle.CaptionFont.Charset = DEFAULT_CHARSET
    HintStyle.CaptionFont.Color = clWindowText
    HintStyle.CaptionFont.Height = -11
    HintStyle.CaptionFont.Name = 'Tahoma'
    HintStyle.CaptionFont.Style = []
    HintStyle.Font.Charset = DEFAULT_CHARSET
    HintStyle.Font.Color = clWindowText
    HintStyle.Font.Height = -11
    HintStyle.Font.Name = 'Tahoma'
    HintStyle.Font.Style = []
    HintStyle.Rounded = True
    Left = 172
    Top = 92
  end
  object cxStyleRepository1: TcxStyleRepository
    Left = 171
    Top = 153
    PixelsPerInch = 96
    object cxStyle1: TcxStyle
      AssignedValues = [svColor, svFont]
      Color = clBtnFace
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = []
    end
  end
end
