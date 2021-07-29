object PoolManager: TPoolManager
  OldCreateOrder = False
  OnCreate = MtsDataModuleCreate
  OnDestroy = MtsDataModuleDestroy
  Pooled = True
  Height = 209
  Width = 268
  object ADOConnection: TADOConnection
    LoginPrompt = False
    Provider = 'MSDAORA.1'
    Left = 64
    Top = 48
  end
  object ADODataSet: TADODataSet
    CacheSize = 1000
    Connection = ADOConnection
    Parameters = <>
    Prepared = True
    Left = 156
    Top = 48
  end
  object ADOCommand: TADOCommand
    Connection = ADOConnection
    Prepared = True
    Parameters = <>
    Left = 64
    Top = 104
  end
end
