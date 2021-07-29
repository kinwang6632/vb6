object ConfigModule: TConfigModule
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  OnDestroy = DataModuleDestroy
  Left = 502
  Top = 313
  Height = 239
  Width = 306
  object AccessConnection: TADOConnection
    ConnectionString = 
      'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\App\Nagra\Progra' +
      'mSchedule\Bin\Config.cfg;Mode=Share Deny Read|Share Deny Write;P' +
      'ersist Security Info=False;Jet OLEDB:Database Password=cyc841772' +
      '82'
    LoginPrompt = False
    Mode = cmShareExclusive
    Provider = 'Microsoft.Jet.OLEDB.4.0'
    Left = 88
    Top = 48
  end
  object DataReader: TADOQuery
    CacheSize = 100
    Connection = AccessConnection
    CursorType = ctStatic
    Parameters = <>
    SQL.Strings = (
      'select * from period')
    Left = 88
    Top = 112
  end
end
