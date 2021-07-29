object LogModule: TLogModule
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  OnDestroy = DataModuleDestroy
  Left = 478
  Top = 261
  Height = 250
  Width = 312
  object AccessConnection: TADOConnection
    ConnectionString = 
      'Provider=Microsoft.Jet.OLEDB.4.0;Password="";Data Source=D:\App\' +
      'Nagra\ProgramSchedule\Bin\Logs.cfg;Mode=Share Deny Read|Share De' +
      'ny Write;Persist Security Info=True;Jet OLEDB:Database Password=' +
      'cyc84177282'
    LoginPrompt = False
    Mode = cmShareExclusive
    Provider = 'Microsoft.Jet.OLEDB.4.0'
    Left = 63
    Top = 29
  end
  object DataReader: TADOQuery
    CacheSize = 1000
    Connection = AccessConnection
    CursorType = ctStatic
    Parameters = <>
    Left = 63
    Top = 89
  end
  object ActionDataSet: TClientDataSet
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'KeyId'
        DataType = ftWideString
        Size = 15
      end
      item
        Name = 'ActDate'
        DataType = ftWideString
        Size = 10
      end
      item
        Name = 'ActTime'
        DataType = ftWideString
        Size = 10
      end
      item
        Name = 'ActFileName'
        DataType = ftWideString
        Size = 100
      end
      item
        Name = 'ActFileSize'
        DataType = ftWideString
        Size = 20
      end
      item
        Name = 'ActFilePath'
        DataType = ftWideString
        Size = 100
      end
      item
        Name = 'ActSource'
        DataType = ftWideString
        Size = 10
      end
      item
        Name = 'ActStatus'
        DataType = ftWideString
        Size = 1
      end
      item
        Name = 'ActCost'
        DataType = ftInteger
      end
      item
        Name = 'ActProgress'
        DataType = ftInteger
      end
      item
        Name = 'ActErrMsg'
        DataType = ftMemo
      end
      item
        Name = 'ActFlag'
        DataType = ftWideString
        Size = 1
      end>
    IndexDefs = <
      item
        Name = 'IDX1'
        Fields = 'KeyId'
        Options = [ixPrimary]
      end>
    Params = <>
    StoreDefs = True
    OnNewRecord = ActionDataSetNewRecord
    Left = 151
    Top = 29
  end
  object DataWriter: TADOQuery
    Connection = AccessConnection
    Parameters = <>
    Left = 64
    Top = 148
  end
end
