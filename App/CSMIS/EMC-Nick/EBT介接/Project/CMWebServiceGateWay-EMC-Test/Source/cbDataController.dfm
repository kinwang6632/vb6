object DataController: TDataController
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  OnDestroy = DataModuleDestroy
  Left = 338
  Top = 331
  Height = 245
  Width = 304
  object LogConnection: TADOConnection
    LoginPrompt = False
    Left = 127
    Top = 24
  end
  object LogWriter: TADOQuery
    CacheSize = 1000
    Connection = LogConnection
    Parameters = <>
    Left = 128
    Top = 75
  end
  object LogReader: TADOQuery
    CacheSize = 1000
    Connection = LogConnection
    Parameters = <>
    Left = 129
    Top = 132
  end
  object CMConnection: TADOConnection
    ConnectionString = 
      'Provider=MSDAORA.1;Password=emc2916nty;User ID=emcnty;Data Sourc' +
      'e=CATVN;Persist Security Info=True'
    LoginPrompt = False
    Provider = 'MSDAORA.1'
    Left = 44
    Top = 24
  end
  object CMDataReader: TADOQuery
    CacheSize = 1000
    Connection = CMConnection
    Parameters = <>
    Left = 44
    Top = 76
  end
  object CMDataWriter: TADOQuery
    Connection = CMConnection
    Parameters = <>
    Left = 44
    Top = 132
  end
  object XMLDoc: TXMLDocument
    Left = 204
    Top = 75
    DOMVendorDesc = 'MSXML'
  end
end
