object DBController: TDBController
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  OnDestroy = DataModuleDestroy
  Left = 412
  Top = 219
  Height = 251
  Width = 268
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
    CursorType = ctStatic
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
end
