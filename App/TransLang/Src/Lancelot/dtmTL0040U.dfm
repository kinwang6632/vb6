inherited dtmTL0040: TdtmTL0040
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
  object adoRWordsInfo: TADOQuery
    Connection = ADOConnection1
    CursorType = ctStatic
    Parameters = <>
    SQL.Strings = (
      'select * from RWordsInfo')
    Left = 56
    Top = 120
  end
  object adoUpdateWordsInfo: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    Left = 56
    Top = 208
  end
end
