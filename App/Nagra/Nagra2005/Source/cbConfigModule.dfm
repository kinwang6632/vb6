object ConfigModule: TConfigModule
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  OnDestroy = DataModuleDestroy
  Left = 646
  Top = 304
  Height = 239
  Width = 319
  object AccessConnection: TADOConnection
    LoginPrompt = False
    Mode = cmShareDenyNone
    Provider = 'Microsoft.Jet.OLEDB.4.0'
    Left = 56
    Top = 39
  end
  object DataReader: TADOQuery
    CacheSize = 1000
    Connection = AccessConnection
    Parameters = <>
    Prepared = True
    Left = 56
    Top = 107
  end
  object HighCmdEnv: TClientDataSet
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'CMD_TYPE'
        DataType = ftWideString
        Size = 255
      end
      item
        Name = 'HIGH_LEVEL_CMD'
        DataType = ftWideString
        Size = 255
      end
      item
        Name = 'DESCRIPTION'
        DataType = ftWideString
        Size = 255
      end
      item
        Name = 'LOW_LEVEL_CMD'
        DataType = ftWideString
        Size = 255
      end
      item
        Name = 'IRD_CMD'
        DataType = ftString
        Size = 255
      end
      item
        Name = 'NOTES_USE'
        DataType = ftWideString
        Size = 255
      end
      item
        Name = 'CMD_TYPE_PRIORITY'
        DataType = ftWideString
        Size = 255
      end
      item
        Name = 'CMD_PRIORITY'
        DataType = ftWideString
        Size = 255
      end
      item
        Name = 'NEW_CMD_COUNT'
        DataType = ftWideString
        Size = 255
      end
      item
        Name = 'OLD_CMD_COUNT'
        DataType = ftWideString
        Size = 255
      end>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 144
    Top = 39
  end
  object LowCmdEnv: TClientDataSet
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'LOW_LEVEL_CMD'
        DataType = ftWideString
        Size = 255
      end
      item
        Name = 'DESCRIPTION'
        DataType = ftWideString
        Size = 255
      end>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 144
    Top = 107
  end
  object CmdError: TClientDataSet
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'ERROR_FLAG'
        DataType = ftWideString
        Size = 255
      end
      item
        Name = 'ERROR_CODE'
        DataType = ftWideString
        Size = 255
      end
      item
        Name = 'ERROR_DESC'
        DataType = ftWideString
        Size = 255
      end>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 220
    Top = 39
  end
end
