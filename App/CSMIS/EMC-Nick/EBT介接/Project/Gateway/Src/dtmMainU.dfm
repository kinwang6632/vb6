object dtmMain: TdtmMain
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  OnDestroy = DataModuleDestroy
  Left = 438
  Top = 240
  Height = 280
  Width = 297
  object cdsUser: TClientDataSet
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'sID'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'sName'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'sPassword'
        DataType = ftString
        Size = 20
      end>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 16
    Top = 24
    object cdsUsersID: TStringField
      DisplayLabel = #20195#30908
      FieldName = 'sID'
    end
    object cdsUsersName: TStringField
      DisplayLabel = #22995#21517
      FieldName = 'sName'
    end
    object cdsUsersPassword: TStringField
      DisplayLabel = #23494#30908
      FieldName = 'sPassword'
    end
  end
  object cdsParam: TClientDataSet
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'sSysName'
        DataType = ftString
        Size = 30
      end
      item
        Name = 'nProcessPeriod'
        DataType = ftSmallint
      end
      item
        Name = 'bLogData'
        DataType = ftBoolean
      end
      item
        Name = 'bShowUI'
        DataType = ftBoolean
      end
      item
        Name = 'bAcceptNullData'
        DataType = ftBoolean
      end>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 72
    Top = 24
    object cdsParamSysName: TStringField
      FieldName = 'sSysName'
      Size = 30
    end
    object cdsParamnProcessPeriod: TSmallintField
      FieldName = 'nProcessPeriod'
    end
    object cdsParambLogData: TBooleanField
      FieldName = 'bLogData'
    end
    object cdsParambShowUI: TBooleanField
      FieldName = 'bShowUI'
    end
    object cdsParambAcceptNullData: TBooleanField
      FieldName = 'bAcceptNullData'
    end
  end
  object connCSIS: TADOConnection
    Left = 128
    Top = 24
  end
  object ADOQuery1: TADOQuery
    Parameters = <>
    Left = 32
    Top = 88
  end
  object EBT007_UIDataSet: TClientDataSet
    Aggregates = <>
    FieldDefs = <>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 40
    Top = 164
  end
end
