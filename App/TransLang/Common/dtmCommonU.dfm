inherited dtmCommon: TdtmCommon
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
  object adoProjName: TADOQuery
    Connection = ADOConnection1
    CursorType = ctStatic
    Parameters = <>
    SQL.Strings = (
      'select distinct sProjName from ProjInfo')
    Left = 160
    Top = 120
  end
  object adoProjInfo: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    Left = 160
    Top = 48
  end
  object adoCommon: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    Left = 56
    Top = 120
  end
end
