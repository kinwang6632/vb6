object ConfigModule: TConfigModule
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  OnDestroy = DataModuleDestroy
  Height = 186
  Width = 282
  object ConfigConnection: TADOConnection
    ConnectionString = 
      'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\App\DVN system\P' +
      'roject\Bin\'#25152#26377#24179#21488'.cfg;Persist Security Info=False;Jet OLEDB:Databa' +
      'se Password=cyc84177282'
    LoginPrompt = False
    Provider = 'Microsoft.Jet.OLEDB.4.0'
    Left = 40
    Top = 24
  end
  object DataReader: TADOQuery
    CacheSize = 1000
    Connection = ConfigConnection
    Parameters = <>
    Prepared = True
    Left = 128
    Top = 24
  end
  object DataWriter: TADOQuery
    CacheSize = 1000
    Connection = ConfigConnection
    Parameters = <>
    Prepared = True
    Left = 204
    Top = 24
  end
end
