object dtmMain: TdtmMain
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  Left = 254
  Top = 178
  Height = 212
  Width = 236
  object ADOConnection1: TADOConnection
    ConnectionString = 
      'Provider=OraOLEDB.Oracle.1;Persist Security Info=False;User ID=g' +
      'icmis;Data Source=gicmis'
    LoginPrompt = False
    Mode = cmReadWrite
    Provider = 'OraOLEDB.Oracle.1'
    OnLogin = ADOConnection1Login
    Left = 40
    Top = 16
  end
  object qryCommon: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    Left = 144
    Top = 16
  end
end
