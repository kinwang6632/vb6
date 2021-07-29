object DBController: TDBController
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  OnDestroy = DataModuleDestroy
  Left = 628
  Top = 124
  Height = 251
  Width = 293
  object DataConnection: TADOConnection
    ConnectionString = 
      'Provider=MSDAORA.1;Password=emcdw;User ID=emcdw;Data Source=MIS;' +
      'Persist Security Info=True'
    LoginPrompt = False
    Provider = 'MSDAORA.1'
    Left = 44
    Top = 24
  end
  object DataReader: TADOQuery
    CacheSize = 1000
    Connection = DataConnection
    Parameters = <>
    Prepared = True
    Left = 45
    Top = 88
  end
  object DataWriter: TADOQuery
    Connection = DataConnection
    Parameters = <>
    Left = 45
    Top = 148
  end
  object CodeReader: TADOQuery
    CacheSize = 1000
    Connection = DataConnection
    Parameters = <>
    Prepared = True
    Left = 116
    Top = 88
  end
  object ComConnection: TADOConnection
    ConnectionString = 
      'Provider=MSDAORA.1;Password=emcdw;User ID=emcdw;Data Source=MIS;' +
      'Persist Security Info=True'
    LoginPrompt = False
    Provider = 'MSDAORA.1'
    Left = 158
    Top = 24
  end
  object ComReader: TADOQuery
    CacheSize = 1000
    Connection = ComConnection
    Parameters = <>
    Prepared = True
    Left = 197
    Top = 96
  end
end
