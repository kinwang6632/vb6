object ConfigModule: TConfigModule
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  OnDestroy = DataModuleDestroy
  Left = 489
  Top = 316
  Height = 239
  Width = 319
  object AccessConnection: TADOConnection
    ConnectionString = 
      'Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Data Source=D:\Ap' +
      'p\Nagra2005\Exe\Config.ini;Mode=Share Deny None;Extended Propert' +
      'ies="";Persist Security Info=False;Jet OLEDB:System database="";' +
      'Jet OLEDB:Registry Path="";Jet OLEDB:Database Password=cyc841772' +
      '82;Jet OLEDB:Engine Type=5;Jet OLEDB:Database Locking Mode=1;Jet' +
      ' OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transacti' +
      'ons=1;Jet OLEDB:New Database Password="";Jet OLEDB:Create System' +
      ' Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don'#39't' +
      ' Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica ' +
      'Repair=False;Jet OLEDB:SFP=False'
    LoginPrompt = False
    Provider = 'Microsoft.Jet.OLEDB.4.0'
    Left = 72
    Top = 31
  end
  object DataReader: TADOQuery
    CacheSize = 1000
    Connection = AccessConnection
    CursorType = ctStatic
    Parameters = <>
    Prepared = True
    Left = 72
    Top = 95
  end
end
