object dtmMain4: TdtmMain4
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  Left = 376
  Top = 176
  Height = 329
  Width = 546
  object adoCD067A: TADOQuery
    Connection = ADOConnection2
    Parameters = <>
    SQL.Strings = (
      'select * from CD067A')
    Left = 96
    Top = 16
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
    Connection = ADOConnection2
    AfterScroll = adoCD067BAfterScroll
    OnCalcFields = adoCD067BCalcFields
    Parameters = <>
    SQL.Strings = (
      'select * from CD067B')
    Left = 24
    Top = 16
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
    Connection = ADOConnection2
    OnCalcFields = adoCD067CCalcFields
    Parameters = <>
    SQL.Strings = (
      'select * from CD067C')
    Left = 96
    Top = 72
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
    Left = 152
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
    Connection = ADOConnection2
    Parameters = <>
    SQL.Strings = (
      'SELECT * FROM CD067A')
    Left = 24
    Top = 80
  end
  object adoRptData: TADOQuery
    Connection = ADOConnection2
    Parameters = <>
    Left = 96
    Top = 128
  end
  object qryCommon: TADOQuery
    Connection = ADOConnection2
    Parameters = <>
    Left = 152
    Top = 72
  end
  object qryPriv: TADOQuery
    Connection = ADOConnection2
    Parameters = <>
    Left = 152
    Top = 128
  end
  object cdsIniInfo: TClientDataSet
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'DataAreaNo'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'DataArea'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'Alias'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'UserID'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'Password'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'CompCode'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'CompName'
        DataType = ftString
        Size = 30
      end>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 216
    Top = 16
    object cdsIniInfoDataAreaNo: TStringField
      FieldName = 'DataAreaNo'
    end
    object cdsIniInfoDataArea: TStringField
      FieldName = 'DataArea'
    end
    object cdsIniInfoAlias: TStringField
      FieldName = 'Alias'
    end
    object cdsIniInfoUserID: TStringField
      FieldName = 'UserID'
    end
    object cdsIniInfoPassword: TStringField
      FieldName = 'Password'
    end
    object cdsIniInfoCompCode: TStringField
      FieldName = 'CompCode'
    end
    object cdsIniInfoCompName: TStringField
      FieldName = 'CompName'
      Size = 30
    end
  end
  object adoQryCD067A: TADOQuery
    Parameters = <>
    Left = 400
    Top = 32
  end
  object adoQryCD067B: TADOQuery
    Parameters = <>
    Left = 400
    Top = 88
  end
  object adoQryCD067C: TADOQuery
    Parameters = <>
    Left = 400
    Top = 144
  end
  object cdsDBError: TClientDataSet
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'ErrorMsg'
        DataType = ftString
        Size = 100
      end
      item
        Name = 'CompCode'
        DataType = ftString
        Size = 20
      end>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 216
    Top = 72
    object cdsDBErrorErrorMsg: TStringField
      FieldName = 'ErrorMsg'
      Size = 100
    end
    object cdsDBErrorCompCode: TStringField
      FieldName = 'CompCode'
    end
  end
  object ADOConnection2: TADOConnection
    LoginPrompt = False
    Left = 48
    Top = 208
  end
end
