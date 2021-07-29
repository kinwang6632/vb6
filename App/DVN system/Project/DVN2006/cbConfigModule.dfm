object ConfigModule: TConfigModule
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  OnDestroy = DataModuleDestroy
  Left = 512
  Top = 247
  Height = 281
  Width = 366
  object ConfigConnection: TADOConnection
    ConnectionString = 
      'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\App\DVN system\P' +
      'roject\Bin\'#25152#26377#24179#21488'.cfg;Persist Security Info=False;Jet OLEDB:Databa' +
      'se Password=cyc84177282'
    LoginPrompt = False
    Provider = 'Microsoft.Jet.OLEDB.4.0'
    Left = 52
    Top = 24
  end
  object DataReader: TADOQuery
    CacheSize = 1000
    Connection = ConfigConnection
    Parameters = <>
    Prepared = True
    Left = 52
    Top = 84
  end
  object DataWriter: TADOQuery
    CacheSize = 1000
    Connection = ConfigConnection
    Parameters = <>
    Prepared = True
    Left = 52
    Top = 144
  end
  object CmdErrorEnv: TClientDataSet
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'ErrorFlag'
        DataType = ftInteger
      end
      item
        Name = 'ErrorCode'
        DataType = ftString
        Size = 50
      end
      item
        Name = 'ErrorDesc'
        DataType = ftString
        Size = 255
      end>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 148
    Top = 25
  end
  object HighCmdEnv: TClientDataSet
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'HighLevelCmd'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'Description'
        DataType = ftString
        Size = 100
      end
      item
        Name = 'LowLevelCmd'
        DataType = ftString
        Size = 255
      end
      item
        Name = 'CmdType'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'CmdTypePriority'
        DataType = ftInteger
      end
      item
        Name = 'CmdPriority'
        DataType = ftInteger
      end
      item
        Name = 'CmdByNote'
        DataType = ftString
        Size = 255
      end>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 148
    Top = 84
  end
  object LowCmdEnv: TClientDataSet
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'LowLevelCmd'
        DataType = ftString
        Size = 50
      end
      item
        Name = 'Description'
        DataType = ftString
        Size = 255
      end>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 148
    Top = 143
  end
end
