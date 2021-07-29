inherited dtmTL0030: TdtmTL0030
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
  object adoTransLanguage: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    Left = 152
    Top = 48
  end
end
