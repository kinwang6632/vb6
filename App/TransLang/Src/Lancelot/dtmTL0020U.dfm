inherited dtmTL0020: TdtmTL0020
  OldCreateOrder = True
  Height = 479
  Width = 741
  inherited ADOConnection1: TADOConnection
    Connected = True
    ConnectionString = 
      'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\App\TransLang\DB' +
      '\TransLang.mdb;Persist Security Info=False'
    Mode = cmShareDenyNone
    Provider = 'Microsoft.Jet.OLEDB.4.0'
  end
  object adoWordsInfo: TADOQuery
    Connection = ADOConnection1
    CursorType = ctStatic
    Parameters = <>
    SQL.Strings = (
      'select * from WordsInfo'
      'where nWordsSeq<0')
    Left = 160
    Top = 112
  end
  object adoCheckExist: TADOQuery
    Connection = ADOConnection1
    CursorType = ctStatic
    Parameters = <
      item
        Name = 'sSrcWords'
        DataType = ftString
        Size = -1
        Value = Null
      end
      item
        Name = 'sHead'
        DataType = ftString
        Size = -1
        Value = Null
      end>
    SQL.Strings = (
      'select * from SrcWordsInfo'
      'where sSrcWords=:sSrcWords')
    Left = 48
    Top = 112
  end
end
