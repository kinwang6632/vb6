object ParseController: TParseController
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  OnDestroy = DataModuleDestroy
  Left = 482
  Top = 314
  Height = 332
  Width = 378
  object DataConnection: TADOConnection
    LoginPrompt = False
    Provider = 'MSDAORA.1'
    Left = 65
    Top = 26
  end
  object DataReader: TADOQuery
    CacheSize = 1000
    Connection = DataConnection
    CursorType = ctOpenForwardOnly
    Parameters = <>
    Left = 153
    Top = 26
  end
  object DataWriter: TADOQuery
    CacheSize = 1000
    Connection = DataConnection
    CursorType = ctOpenForwardOnly
    Parameters = <>
    Left = 153
    Top = 90
  end
  object ParseOrderDataSet: TClientDataSet
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'XmlFileType'
        DataType = ftInteger
      end
      item
        Name = 'XmlFileName'
        DataType = ftString
        Size = 200
      end
      item
        Name = 'XmlFilePrefix'
        DataType = ftString
        Size = 8
      end
      item
        Name = 'XmlFileDateTime'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'Status'
        DataType = ftString
        Size = 1
      end>
    IndexDefs = <
      item
        Name = 'IDX1'
        Fields = 'XmlFileType;XmlFileDateTime'
      end>
    Params = <>
    StoreDefs = True
    Left = 67
    Top = 86
  end
  object DataInsert: TADOQuery
    Connection = DataConnection
    Parameters = <>
    SQL.Strings = (
      '')
    Left = 233
    Top = 26
  end
  object DataUpdate: TADOQuery
    Connection = DataConnection
    Parameters = <>
    Left = 233
    Top = 90
  end
  object DataDelete: TADOQuery
    Connection = DataConnection
    Parameters = <>
    Left = 233
    Top = 150
  end
  object DataSetProvider: TDataSetProvider
    DataSet = DataReader
    Options = [poAllowCommandText]
    Left = 67
    Top = 150
  end
end
