object dtmMain: TdtmMain
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  Left = 625
  Top = 335
  Height = 400
  Width = 215
  object adoConn: TADOConnection
    LoginPrompt = False
    Provider = 'MSDAORA.1'
    Left = 40
    Top = 24
  end
  object qryComm: TADOQuery
    Connection = adoConn
    Parameters = <>
    Left = 112
    Top = 24
  end
  object qryGroupData: TADOQuery
    Connection = adoConn
    AfterInsert = qryGroupDataAfterInsert
    AfterScroll = qryGroupDataAfterScroll
    OnCalcFields = qryGroupDataCalcFields
    Parameters = <>
    SQL.Strings = (
      'select * from EMC_EBT003 where CODENO is null')
    Left = 40
    Top = 80
    object qryGroupDataCODENO: TIntegerField
      FieldName = 'CODENO'
    end
    object qryGroupDataDESCRIPTION: TStringField
      FieldName = 'DESCRIPTION'
      Size = 60
    end
    object qryGroupDataSONAME: TStringField
      FieldName = 'SONAME'
      Size = 120
    end
    object qryGroupDataSOCODE: TStringField
      FieldName = 'SOCODE'
      Size = 50
    end
    object qryGroupDataSTOPFLAG: TIntegerField
      FieldName = 'STOPFLAG'
    end
    object qryGroupDataUPDTIME: TDateTimeField
      FieldName = 'UPDTIME'
    end
    object qryGroupDataUPDEN: TStringField
      FieldName = 'UPDEN'
    end
    object qryGroupDataSTOPFLAGDESC: TStringField
      FieldKind = fkCalculated
      FieldName = 'STOPFLAGDESC'
      Calculated = True
    end
  end
  object qryUserData: TADOQuery
    Connection = adoConn
    CursorType = ctStatic
    AfterInsert = qryUserDataAfterInsert
    OnCalcFields = qryUserDataCalcFields
    Parameters = <>
    SQL.Strings = (
      'select * from EMC_EBT002 where UserID is null')
    Left = 112
    Top = 80
    object qryUserDataUSERID: TStringField
      FieldName = 'USERID'
      Size = 8
    end
    object qryUserDataUSERNAME: TStringField
      FieldName = 'USERNAME'
      Size = 30
    end
    object qryUserDataUSERPASSWD: TStringField
      FieldName = 'USERPASSWD'
      Size = 10
    end
    object qryUserDataUSERGROUP: TIntegerField
      FieldName = 'USERGROUP'
    end
    object qryUserDataISSUPERVISOR: TIntegerField
      FieldName = 'ISSUPERVISOR'
    end
    object qryUserDataISSYSOP: TIntegerField
      FieldName = 'ISSYSOP'
    end
    object qryUserDataSTOPFLAG: TIntegerField
      FieldName = 'STOPFLAG'
    end
    object qryUserDataUPDTIME: TDateTimeField
      FieldName = 'UPDTIME'
    end
    object qryUserDataUPDEN: TStringField
      FieldName = 'UPDEN'
    end
    object qryUserDataCOMPCODE: TIntegerField
      FieldName = 'COMPCODE'
    end
    object qryUserDataDEPTCODE: TIntegerField
      FieldName = 'DEPTCODE'
    end
    object qryUserDataTEL: TStringField
      FieldName = 'TEL'
    end
    object qryUserDataTELEXT: TIntegerField
      FieldName = 'TELEXT'
    end
    object qryUserDataEMAIL: TStringField
      FieldName = 'EMAIL'
      Size = 50
    end
    object qryUserDataCREATER: TStringField
      FieldName = 'CREATER'
    end
    object qryUserDataCREATEDATE: TDateTimeField
      FieldName = 'CREATEDATE'
    end
    object qryUserDataSTOPFLAGDESC: TStringField
      FieldKind = fkCalculated
      FieldName = 'STOPFLAGDESC'
      Calculated = True
    end
    object qryUserDataISSUPERVISORDESC: TStringField
      FieldKind = fkCalculated
      FieldName = 'ISSUPERVISORDESC'
      Calculated = True
    end
    object qryUserDataISSYSOPDESC: TStringField
      FieldKind = fkCalculated
      FieldName = 'ISSYSOPDESC'
      Calculated = True
    end
    object qryUserDataCOMPCODEDESC: TStringField
      FieldKind = fkCalculated
      FieldName = 'COMPCODEDESC'
      Size = 30
      Calculated = True
    end
    object qryUserDataDEPTCODEDESC: TStringField
      FieldKind = fkCalculated
      FieldName = 'DEPTCODEDESC'
      Size = 30
      Calculated = True
    end
    object qryUserDataGROUPCODEDESC: TStringField
      FieldKind = fkCalculated
      FieldName = 'GROUPCODEDESC'
      Size = 30
      Calculated = True
    end
  end
  object qryCompCode: TADOQuery
    Connection = adoConn
    Parameters = <>
    SQL.Strings = (
      'select * from EMC_EBT010 where STOPFLAG=0')
    Left = 40
    Top = 248
  end
  object qryDeptCode: TADOQuery
    Connection = adoConn
    Parameters = <>
    SQL.Strings = (
      'select * from EMC_EBT011 where STOPFLAG=0')
    Left = 40
    Top = 200
  end
  object qryGroupCode: TADOQuery
    Connection = adoConn
    Parameters = <>
    SQL.Strings = (
      'select * from EMC_EBT003 where STOPFLAG=0 ')
    Left = 40
    Top = 304
  end
  object qryAllGroupCode: TADOQuery
    Connection = adoConn
    Parameters = <>
    SQL.Strings = (
      'select * from EMC_EBT003 order by CODENO')
    Left = 120
    Top = 304
  end
  object qryAllCompCode: TADOQuery
    Connection = adoConn
    Parameters = <>
    SQL.Strings = (
      'select * from EMC_EBT010  order by CODENO')
    Left = 120
    Top = 248
  end
  object qryAllDeptCode: TADOQuery
    Connection = adoConn
    Parameters = <>
    SQL.Strings = (
      'select * from EMC_EBT011  order by CODENO')
    Left = 120
    Top = 200
  end
end
