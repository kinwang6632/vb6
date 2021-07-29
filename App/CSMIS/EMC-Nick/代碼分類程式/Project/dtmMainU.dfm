object dtmMain: TdtmMain
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  Left = 130
  Top = 182
  Height = 329
  Width = 507
  object adoConnection: TADOConnection
    LoginPrompt = False
    Provider = 'MSDAORA.1'
    Left = 48
    Top = 8
  end
  object adoCD067A: TADOQuery
    Connection = adoConnection
    Parameters = <>
    SQL.Strings = (
      'select * from CD067A')
    Left = 152
    Top = 8
    object adoCD067ATABLENAME: TStringField
      DisplayLabel = 'Table Name'
      FieldName = 'TABLENAME'
      Size = 30
    end
    object adoCD067ATABLEDESCRIPTION: TStringField
      DisplayLabel = 'Description'
      FieldName = 'TABLEDESCRIPTION'
      Size = 50
    end
    object adoCD067AOPERATOR: TStringField
      DisplayLabel = #30064#21205#20154#21729
      FieldName = 'OPERATOR'
    end
    object adoCD067AUPDTIME: TStringField
      DisplayLabel = #30064#21205#26178#38291
      FieldName = 'UPDTIME'
    end
  end
  object adoCD067B: TADOQuery
    Connection = adoConnection
    AfterScroll = adoCD067BAfterScroll
    OnCalcFields = adoCD067BCalcFields
    Parameters = <>
    SQL.Strings = (
      'select * from CD067B')
    Left = 48
    Top = 64
    object adoCD067BCODENO: TIntegerField
      FieldName = 'CODENO'
    end
    object adoCD067BTABLENAME: TStringField
      FieldName = 'TABLENAME'
      Size = 30
    end
    object adoCD067BDESCRIPTION: TStringField
      FieldName = 'DESCRIPTION'
    end
    object adoCD067BREFNO: TIntegerField
      FieldName = 'REFNO'
    end
    object adoCD067BSERVICETYPE: TStringField
      FieldName = 'SERVICETYPE'
      FixedChar = True
      Size = 1
    end
    object adoCD067BSTOPFLAG: TIntegerField
      FieldName = 'STOPFLAG'
    end
    object adoCD067BOPERATOR: TStringField
      FieldName = 'OPERATOR'
    end
    object adoCD067BUPDTIME: TStringField
      FieldName = 'UPDTIME'
    end
    object adoCD067BServiceTypeName: TStringField
      FieldKind = fkCalculated
      FieldName = 'ServiceTypeName'
      Calculated = True
    end
    object adoCD067BStopFlagName: TStringField
      FieldKind = fkCalculated
      FieldName = 'StopFlagName'
      Calculated = True
    end
  end
  object adoCD067C: TADOQuery
    Connection = adoConnection
    OnCalcFields = adoCD067CCalcFields
    Parameters = <>
    SQL.Strings = (
      'select * from CD067C')
    Left = 152
    Top = 64
    object adoCD067CCOMPCODE: TIntegerField
      FieldName = 'COMPCODE'
    end
    object adoCD067CTABLENAME: TStringField
      FieldName = 'TABLENAME'
      Size = 30
    end
    object adoCD067CMASTERCODENO: TIntegerField
      FieldName = 'MASTERCODENO'
    end
    object adoCD067CDETAILCODENO: TIntegerField
      FieldName = 'DETAILCODENO'
    end
    object adoCD067CSTOPFLAG: TIntegerField
      FieldName = 'STOPFLAG'
    end
    object adoCD067COPERATOR: TStringField
      FieldName = 'OPERATOR'
    end
    object adoCD067CUPDTIME: TStringField
      FieldName = 'UPDTIME'
    end
    object adoCD067CStopFlagName: TStringField
      FieldKind = fkCalculated
      FieldName = 'StopFlagName'
      Calculated = True
    end
    object adoCD067CrefCompName: TStringField
      FieldKind = fkCalculated
      FieldName = 'refCompName'
      Calculated = True
    end
  end
  object cdsCompInfo: TClientDataSet
    Aggregates = <>
    Params = <>
    Left = 264
    Top = 16
    object cdsCompInfoCompCode: TIntegerField
      FieldName = 'CompCode'
    end
    object cdsCompInfoCompName: TStringField
      FieldName = 'CompName'
    end
    object cdsCompInfoIndex: TIntegerField
      FieldName = 'Index'
    end
  end
  object adoCodeCD067A: TADOQuery
    Connection = adoConnection
    Parameters = <>
    SQL.Strings = (
      'SELECT * FROM CD067A')
    Left = 48
    Top = 128
  end
  object adoRptData: TADOQuery
    Connection = adoConnection
    Parameters = <>
    Left = 160
    Top = 128
  end
  object qryCommon: TADOQuery
    Connection = adoConnection
    Parameters = <>
    Left = 272
    Top = 80
  end
  object qryPriv: TADOQuery
    Connection = adoConnection
    Parameters = <>
    Left = 272
    Top = 136
  end
end
