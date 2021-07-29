inherited dtmCompData: TdtmCompData
  OldCreateOrder = True
  inherited ADOQuery1: TADOQuery
    SQL.Strings = (
      'select * from EMC_EBT010')
  end
  inherited ClientDataSet1: TClientDataSet
    AfterInsert = ClientDataSet1AfterInsert
    BeforePost = ClientDataSet1BeforePost
    OnCalcFields = ClientDataSet1CalcFields
    object ClientDataSet1CODENO: TIntegerField
      FieldName = 'CODENO'
    end
    object ClientDataSet1DESCRIPTION: TStringField
      FieldName = 'DESCRIPTION'
      Size = 30
    end
    object ClientDataSet1STOPFLAG: TIntegerField
      FieldName = 'STOPFLAG'
    end
    object ClientDataSet1UPDTIME: TDateTimeField
      FieldName = 'UPDTIME'
    end
    object ClientDataSet1UPDEN: TStringField
      FieldName = 'UPDEN'
    end
    object ClientDataSet1STOPFLAGDESC: TStringField
      FieldKind = fkCalculated
      FieldName = 'STOPFLAGDESC'
      Size = 10
      Calculated = True
    end
  end
end
