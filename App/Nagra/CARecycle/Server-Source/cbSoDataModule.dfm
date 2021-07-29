object SoDataModule: TSoDataModule
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  OnDestroy = DataModuleDestroy
  Left = 468
  Top = 247
  Height = 255
  Width = 294
  object SoConnection: TADOConnection
    LoginPrompt = False
    Provider = 'MSDAORA.1'
    Left = 56
    Top = 20
  end
  object SoDataReader: TADOQuery
    CacheSize = 1000
    Connection = SoConnection
    CursorType = ctOpenForwardOnly
    Parameters = <>
    Left = 56
    Top = 76
  end
  object SoDataWriter: TADOQuery
    Connection = SoConnection
    Parameters = <>
    Left = 56
    Top = 133
  end
  object CAConnection: TADOConnection
    LoginPrompt = False
    Provider = 'MSDAORA.1'
    Left = 136
    Top = 20
  end
  object CADataReader: TADOQuery
    CacheSize = 1000
    Connection = CAConnection
    CursorType = ctOpenForwardOnly
    Parameters = <>
    Left = 136
    Top = 76
  end
  object CADataWriter: TADOQuery
    Connection = CAConnection
    Parameters = <>
    Left = 137
    Top = 133
  end
end
