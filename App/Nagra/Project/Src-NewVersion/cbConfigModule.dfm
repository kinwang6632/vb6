object ConfigModule: TConfigModule
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  OnDestroy = DataModuleDestroy
  Left = 523
  Top = 351
  Height = 214
  Width = 280
  object AccessConnection: TADOConnection
    ConnectionString = 
      'Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Data Source=D:\Ap' +
      'p\Nagra\Project\Bin\'#25152#26377#24179#21488'.ini;Mode=Share Deny None;Extended Prope' +
      'rties="";Persist Security Info=False;Jet OLEDB:System database="' +
      '";Jet OLEDB:Registry Path="";Jet OLEDB:Database Password=cyc8417' +
      '7282;Jet OLEDB:Engine Type=5;Jet OLEDB:Database Locking Mode=1;J' +
      'et OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transac' +
      'tions=1;Jet OLEDB:New Database Password="";Jet OLEDB:Create Syst' +
      'em Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don' +
      #39't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replic' +
      'a Repair=False;Jet OLEDB:SFP=False'
    LoginPrompt = False
    Provider = 'Microsoft.Jet.OLEDB.4.0'
    Left = 60
    Top = 32
  end
  object DataReader: TADOQuery
    CacheSize = 1000
    Connection = AccessConnection
    CursorType = ctStatic
    Parameters = <>
    SQL.Strings = (
      'SELECT A.*, '
      '       0 AS CMD_TYPE, '
      '       1 AS CMD_PRIORTY '
      'FROM SEND_NAGRA A')
    Left = 60
    Top = 92
  end
  object CmdEnv: TClientDataSet
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'HIGH_LEVEL_CMD'
        DataType = ftWideString
        Size = 4
      end
      item
        Name = 'DESCRIPTION'
        DataType = ftWideString
        Size = 100
      end
      item
        Name = 'LOW_LEVE_CMD'
        DataType = ftWideString
        Size = 128
      end
      item
        Name = 'CMD_TYPE'
        DataType = ftMemo
      end
      item
        Name = 'CMD_TYPE_PRIORITY'
        DataType = ftInteger
      end
      item
        Name = 'CMD_PRIORITY'
        DataType = ftInteger
      end>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 152
    Top = 68
  end
end
