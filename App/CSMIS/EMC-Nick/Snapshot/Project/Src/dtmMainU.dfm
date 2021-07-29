object dtmMain: TdtmMain
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  OnDestroy = DataModuleDestroy
  Left = 102
  Top = 143
  Height = 321
  Width = 542
  object ADOConnection1: TADOConnection
    LoginPrompt = False
    Provider = 'MSDAORA.1'
    Left = 48
    Top = 16
  end
  object qrySO030A: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    SQL.Strings = (
      'SELECT * FROM SO030A WHERE JOBID IS NULL')
    Left = 48
    Top = 80
  end
  object dspSO030A: TDataSetProvider
    DataSet = qrySO030A
    Constraints = True
    Options = [poAllowCommandText]
    Left = 48
    Top = 128
  end
  object cdsSO030A: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'dspSO030A'
    Left = 48
    Top = 176
  end
  object qrySO001: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    SQL.Strings = (
      'SELECT * FROM SO001 WHERE CustID IS NULL')
    Left = 112
    Top = 80
  end
  object dspSO001: TDataSetProvider
    DataSet = qrySO001
    Constraints = True
    Options = [poAllowCommandText]
    Left = 112
    Top = 128
  end
  object cdsSO001: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'dspSO001'
    Left = 112
    Top = 176
  end
  object cdsCustIDCounts: TClientDataSet
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'CompCode'
        DataType = ftInteger
      end
      item
        Name = 'cdsCustIDCountsField2'
      end>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 400
    Top = 80
  end
  object qrySO509: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    SQL.Strings = (
      'SELECT * FROM SO509 WHERE TABLECODE IS NULL')
    Left = 176
    Top = 80
  end
  object dspSO509: TDataSetProvider
    DataSet = qrySO509
    Constraints = True
    Options = [poAllowCommandText]
    Left = 176
    Top = 128
  end
  object cdsSO509: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'dspSO509'
    Left = 176
    Top = 176
  end
  object qrySO509B: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    SQL.Strings = (
      
        'SELECT CompCode,CompName,TableCode,TableName,MasterComputeDate,M' +
        'asterRecordCounts,SnapComputeDate,SnapRecordCounts,DiffCounts,to' +
        '_char(DiffCounts) StrDiffCounts,ResultFlag  FROM SO509B'
      ' WHERE TABLECODE IS NULL')
    Left = 315
    Top = 80
  end
  object dspSO509B: TDataSetProvider
    DataSet = qrySO509B
    Constraints = True
    Options = [poAllowCommandText]
    Left = 315
    Top = 128
  end
  object cdsSO509B: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'dspSO509B'
    Left = 315
    Top = 176
    object cdsSO509BCOMPCODE: TIntegerField
      FieldName = 'COMPCODE'
      Visible = False
    end
    object cdsSO509BCOMPNAME: TStringField
      DisplayLabel = #20844#21496#21029
      FieldName = 'COMPNAME'
      Visible = False
    end
    object cdsSO509BTABLECODE: TStringField
      DisplayWidth = 20
      FieldName = 'TABLECODE'
      Visible = False
      Size = 30
    end
    object cdsSO509BTABLENAME: TStringField
      DisplayLabel = 'Table'#21517#31281
      DisplayWidth = 30
      FieldName = 'TABLENAME'
      Size = 60
    end
    object cdsSO509BMASTERCOMPUTEDATE: TDateTimeField
      DisplayLabel = 'SO'#32113#35336#26085#26399
      FieldName = 'MASTERCOMPUTEDATE'
    end
    object cdsSO509BMASTERRECORDCOUNTS: TIntegerField
      DisplayLabel = 'SO Table '#31558#25976
      FieldName = 'MASTERRECORDCOUNTS'
    end
    object cdsSO509BSNAPCOMPUTEDATE: TDateTimeField
      DisplayLabel = 'Snapshot'#32113#35336#26085#26399
      FieldName = 'SNAPCOMPUTEDATE'
    end
    object cdsSO509BSNAPRECORDCOUNTS: TIntegerField
      DisplayLabel = 'Snapshot Table '#31558#25976
      FieldName = 'SNAPRECORDCOUNTS'
    end
    object cdsSO509BDIFFCOUNTS: TIntegerField
      DisplayLabel = #24046#30064
      FieldName = 'DIFFCOUNTS'
    end
    object cdsSO509BRESULTFLAG: TIntegerField
      DisplayLabel = #29376#24907
      FieldName = 'RESULTFLAG'
    end
    object cdsSO509BSTRDIFFCOUNTS: TStringField
      FieldName = 'STRDIFFCOUNTS'
      ReadOnly = True
      Visible = False
      Size = 40
    end
  end
  object qrySO509A: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    SQL.Strings = (
      'SELECT * FROM SO509A WHERE TABLECODE IS NULL')
    Left = 240
    Top = 80
  end
  object dspSO509A: TDataSetProvider
    DataSet = qrySO509A
    Constraints = True
    Options = [poAllowCommandText]
    Left = 240
    Top = 128
  end
  object cdsSO509A: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'dspSO509A'
    Left = 240
    Top = 176
  end
  object qryCom: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    Left = 160
    Top = 8
  end
  object cdsTempSO509B: TClientDataSet
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'CompName'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'TableName'
        DataType = ftString
        Size = 30
      end
      item
        Name = 'MasterComputeDate'
        DataType = ftDateTime
      end
      item
        Name = 'MasterRecordCounts'
        DataType = ftInteger
      end
      item
        Name = 'SnapComputeDate'
        DataType = ftDateTime
      end
      item
        Name = 'SnapRecordCounts'
        DataType = ftInteger
      end
      item
        Name = 'DiffCounts'
        DataType = ftInteger
      end
      item
        Name = 'ResultFlag'
        DataType = ftString
        Size = 20
      end>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 320
    Top = 232
    object cdsTempSO509BCompName: TStringField
      DisplayLabel = #20844#21496#21029
      DisplayWidth = 15
      FieldName = 'CompName'
    end
    object cdsTempSO509BTableName: TStringField
      DisplayLabel = 'Table'#21517#31281
      DisplayWidth = 20
      FieldName = 'TableName'
      Size = 30
    end
    object cdsTempSO509BMasterComputeDate: TDateTimeField
      DisplayLabel = 'SO'#32113#35336#26085#26399
      FieldName = 'MasterComputeDate'
    end
    object cdsTempSO509BMasterRecordCounts: TIntegerField
      DisplayLabel = 'SO Table '#31558#25976
      FieldName = 'MasterRecordCounts'
    end
    object cdsTempSO509BSnapComputeDate: TDateTimeField
      DisplayLabel = 'Snapshot'#32113#35336#26085#26399
      FieldName = 'SnapComputeDate'
    end
    object cdsTempSO509BSnapRecordCounts: TIntegerField
      DisplayLabel = 'Snapshot Table '#31558#25976
      FieldName = 'SnapRecordCounts'
    end
    object cdsTempSO509BDiffCounts: TIntegerField
      DisplayLabel = #24046#30064
      FieldName = 'DiffCounts'
    end
    object cdsTempSO509BResultFlag: TStringField
      DisplayLabel = #29376#24907
      FieldName = 'ResultFlag'
    end
  end
end
