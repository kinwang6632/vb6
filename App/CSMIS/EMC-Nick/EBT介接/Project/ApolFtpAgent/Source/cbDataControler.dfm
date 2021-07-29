object DataControler: TDataControler
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  OnDestroy = DataModuleDestroy
  Left = 552
  Top = 259
  Height = 270
  Width = 287
  object DataConnection: TADOConnection
    CommandTimeout = 15
    ConnectionString = 
      'Provider=MSDAORA.1;Password=emcdw;User ID=emcdw;Data Source=MIS;' +
      'Persist Security Info=True'
    LoginPrompt = False
    Provider = 'MSDAORA.1'
    Left = 52
    Top = 28
  end
  object DataReader: TADOQuery
    CacheSize = 1000
    Connection = DataConnection
    CursorType = ctStatic
    CommandTimeout = 15
    Parameters = <>
    Left = 52
    Top = 84
  end
  object ExportDataSet: TClientDataSet
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'COMPCODE'
        DataType = ftInteger
      end
      item
        Name = 'COMPNAME'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'COMMANDTYPE'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'CUSTID'
        DataType = ftInteger
      end
      item
        Name = 'CMMAC'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'DIALACCOUNT'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'DIALPASSWORD'
        DataType = ftString
        Size = 24
      end
      item
        Name = 'CUSTNAME'
        DataType = ftString
        Size = 100
      end
      item
        Name = 'TELDAY'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'TELNIGHT'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'FINTIME'
        DataType = ftDateTime
      end
      item
        Name = 'SENDDATE'
        DataType = ftDateTime
      end
      item
        Name = 'FTPDATE'
        DataType = ftDateTime
      end
      item
        Name = 'CMDSEQNO'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'ID'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'FILENAME'
        DataType = ftString
        Size = 50
      end>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 140
    Top = 84
  end
  object DataWriter: TADOQuery
    Connection = DataConnection
    CommandTimeout = 15
    Parameters = <>
    Prepared = True
    Left = 52
    Top = 140
  end
end
