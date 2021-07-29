object SoDataModule: TSoDataModule
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  OnDestroy = DataModuleDestroy
  Left = 661
  Top = 293
  Height = 338
  Width = 383
  object CAConnection: TADOConnection
    ConnectionString = 
      'Provider=MSDAORA.1;Password=com;User ID=com;Data Source=MIS;Pers' +
      'ist Security Info=True'
    LoginPrompt = False
    Provider = 'MSDAORA.1'
    Left = 60
    Top = 41
  end
  object CADataReader: TADOQuery
    CacheSize = 1000
    Connection = CAConnection
    Parameters = <>
    Left = 60
    Top = 97
  end
  object SoConnection: TADOConnection
    ConnectionString = 
      'Provider=MSDAORA.1;Password=com;User ID=com;Data Source=MIS;Pers' +
      'ist Security Info=True'
    LoginPrompt = False
    Provider = 'MSDAORA.1'
    Left = 144
    Top = 42
  end
  object SoDataReader: TADOQuery
    CacheSize = 1000
    Connection = SoConnection
    Parameters = <>
    Left = 144
    Top = 98
  end
  object cdsInput: TClientDataSet
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'SEQNO'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'ICCNO'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'STBNO'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'STBFLAG'
        DataType = ftInteger
      end
      item
        Name = 'STBAUTOCB'
        DataType = ftInteger
      end
      item
        Name = 'COMPCODE'
        DataType = ftInteger
      end
      item
        Name = 'RCSTARTDATE'
        DataType = ftDateTime
      end
      item
        Name = 'RCENDDATE'
        DataType = ftDateTime
      end
      item
        Name = 'OPERATOR'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'UPDTIME'
        DataType = ftDateTime
      end
      item
        Name = 'LASTDOWNLOADTIME'
        DataType = ftDateTime
      end
      item
        Name = 'CMDFLAG'
        DataType = ftInteger
      end
      item
        Name = 'RECYCLETEXT'
        DataType = ftString
        Size = 50
      end
      item
        Name = 'TRANSNUM'
        DataType = ftString
        Size = 512
      end
      item
        Name = 'CONFIRM'
        DataType = ftString
        Size = 20
      end>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 60
    Top = 212
  end
  object CADataWriter: TADOQuery
    Connection = CAConnection
    Parameters = <>
    Left = 60
    Top = 156
  end
  object AccessConnection: TADOConnection
    ConnectionString = 
      'Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Data Source=D:\Ap' +
      'p\Nagra\CARecycle\Bin\Config.cfg;Mode=Share Deny None;Extended P' +
      'roperties="";Persist Security Info=False;Jet OLEDB:System databa' +
      'se="";Jet OLEDB:Registry Path="";Jet OLEDB:Database Password=cyc' +
      '84177282;Jet OLEDB:Engine Type=5;Jet OLEDB:Database Locking Mode' +
      '=1;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Tra' +
      'nsactions=1;Jet OLEDB:New Database Password="";Jet OLEDB:Create ' +
      'System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB' +
      ':Don'#39't Copy Locale on Compact=False;Jet OLEDB:Compact Without Re' +
      'plica Repair=False;Jet OLEDB:SFP=False'
    LoginPrompt = False
    Provider = 'Microsoft.Jet.OLEDB.4.0'
    Left = 234
    Top = 41
  end
  object AccessDataReader: TADOQuery
    CacheSize = 100
    Connection = AccessConnection
    Parameters = <>
    Left = 236
    Top = 97
  end
  object cdsNoCall: TClientDataSet
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'SEQNO'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'ICCNO'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'STBNO'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'STB_FLAG'
        DataType = ftInteger
      end
      item
        Name = 'STBAUTOCB'
        DataType = ftInteger
      end
      item
        Name = 'COMPCODE'
        DataType = ftInteger
      end
      item
        Name = 'RCSTARTDATE'
        DataType = ftDateTime
      end
      item
        Name = 'RCENDDATE'
        DataType = ftDateTime
      end
      item
        Name = 'OPERATOR'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'UPDTIME'
        DataType = ftDateTime
      end
      item
        Name = 'LASTDOWNLOADTIME'
        DataType = ftDateTime
      end
      item
        Name = 'CMDFLAG'
        DataType = ftInteger
      end
      item
        Name = 'RECYCLETEXT'
        DataType = ftString
        Size = 50
      end
      item
        Name = 'TRANSNUM'
        DataType = ftString
        Size = 512
      end
      item
        Name = 'CONFIRM'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'ERRMSG'
        DataType = ftString
        Size = 512
      end
      item
        Name = 'STATE'
        DataType = ftString
        Size = 1
      end
      item
        Name = 'CMD52_1'
        DataType = ftString
        Size = 1
      end
      item
        Name = 'CMD8'
        DataType = ftString
        Size = 1
      end
      item
        Name = 'CMD97_96'
        DataType = ftString
        Size = 1
      end
      item
        Name = 'CMD7'
        DataType = ftString
        Size = 1
      end
      item
        Name = 'CMD48'
        DataType = ftString
        Size = 2
      end
      item
        Name = 'CMD52_2'
        DataType = ftString
        Size = 1
      end>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 144
    Top = 212
  end
  object cdsCall: TClientDataSet
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'SEQNO'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'ICCNO'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'STBNO'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'STB_FLAG'
        DataType = ftInteger
      end
      item
        Name = 'STBAUTOCB'
        DataType = ftInteger
      end
      item
        Name = 'COMPCODE'
        DataType = ftInteger
      end
      item
        Name = 'RCSTARTDATE'
        DataType = ftDateTime
      end
      item
        Name = 'RCENDDATE'
        DataType = ftDateTime
      end
      item
        Name = 'OPERATOR'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'UPDTIME'
        DataType = ftDateTime
      end
      item
        Name = 'LASTDOWNLOADTIME'
        DataType = ftDateTime
      end
      item
        Name = 'CMDFLAG'
        DataType = ftInteger
      end
      item
        Name = 'RECYCLETEXT'
        DataType = ftString
        Size = 50
      end
      item
        Name = 'TRANSNUM'
        DataType = ftString
        Size = 512
      end
      item
        Name = 'CONFIRM'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'ERRMSG'
        DataType = ftString
        Size = 512
      end
      item
        Name = 'STATE'
        DataType = ftString
        Size = 1
      end
      item
        Name = 'CMD52_1'
        DataType = ftString
        Size = 1
      end
      item
        Name = 'CMD62'
        DataType = ftString
        Size = 1
      end
      item
        Name = 'CMD8'
        DataType = ftString
        Size = 1
      end
      item
        Name = 'CMD60'
        DataType = ftString
        Size = 1
      end
      item
        Name = 'CMD7'
        DataType = ftString
        Size = 1
      end
      item
        Name = 'CMD48'
        DataType = ftString
        Size = 2
      end
      item
        Name = 'CMD99'
        DataType = ftString
        Size = 1
      end
      item
        Name = 'CMD52_2'
        DataType = ftString
        Size = 1
      end>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 240
    Top = 212
  end
  object AccessDataWriter: TADOQuery
    Connection = AccessConnection
    Parameters = <>
    Left = 239
    Top = 156
  end
end
