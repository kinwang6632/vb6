object dtmMain: TdtmMain
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  Left = 192
  Top = 107
  Height = 402
  Width = 484
  object ADOConnection1: TADOConnection
    LoginPrompt = False
    Left = 32
    Top = 16
  end
  object ADOQuery1: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    Left = 128
    Top = 16
  end
  object cdsCD024: TClientDataSet
    Aggregates = <>
    Params = <>
    Left = 32
    Top = 72
  end
  object DataSetProvider1: TDataSetProvider
    Constraints = True
    Options = [poAllowCommandText]
    Left = 208
    Top = 16
  end
  object adoLogConnection: TADOConnection
    LoginPrompt = False
    Mode = cmShareDenyNone
    Provider = 'Microsoft.Jet.OLEDB.4.0'
    Left = 40
    Top = 160
  end
  object qryReport1: TADOQuery
    Connection = adoLogConnection
    Parameters = <>
    Left = 128
    Top = 160
  end
end
