inherited dtmMain: TdtmMain
  OldCreateOrder = True
  Height = 298
  Width = 400
  inherited ADOConnection1: TADOConnection
    ConnectionString = 
      'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\App\TransLang\DB' +
      '\TransLang.mdb;Persist Security Info=False'
    Top = 52
  end
  object ADOPrjInfoQry: TADOQuery
    Active = True
    CacheSize = 1000
    Connection = ADOConnection1
    CursorType = ctStatic
    Parameters = <>
    SQL.Strings = (
      'select * from ProjInfo')
    Left = 56
    Top = 120
    object ADOPrjInfoQrynSeq: TIntegerField
      FieldName = 'nSeq'
      ReadOnly = True
    end
    object ADOPrjInfoQrysProjName: TWideStringField
      FieldName = 'sProjName'
      Size = 100
    end
    object ADOPrjInfoQrysUnitName: TWideStringField
      FieldName = 'sUnitName'
      Size = 100
    end
  end
  object CDSTmpQry: TClientDataSet
    Active = True
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'sRepeated'
        DataType = ftString
        Size = 1
      end
      item
        Name = 'sProjName'
        DataType = ftString
        Size = 100
      end
      item
        Name = 'sUnitName'
        DataType = ftString
        Size = 100
      end>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 156
    Top = 52
    Data = {
      740000009619E0BD010000001800000003000000000003000000740009735265
      7065617465640100490000000100055749445448020002000100097350726F6A
      4E616D6501004900000001000557494454480200020064000973556E69744E61
      6D6501004900000001000557494454480200020064000000}
  end
  object ADOExecSQL: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    Left = 56
    Top = 192
  end
end
