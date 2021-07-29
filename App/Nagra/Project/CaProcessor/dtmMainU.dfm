object dtmMain: TdtmMain
  OldCreateOrder = False
  Left = 192
  Top = 107
  Height = 99
  Width = 142
  object cdsParam: TClientDataSet
    Aggregates = <>
    FieldDefs = <>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 16
    Top = 16
    object cdsParamsServerIP2: TStringField
      FieldName = 'sServerIP'
      Size = 15
    end
    object cdsParamnSPortNo2: TIntegerField
      FieldName = 'nSPortNo'
    end
    object cdsParamnRPortNo2: TIntegerField
      FieldName = 'nRPortNo'
    end
    object cdsParamsSysName2: TStringField
      FieldName = 'sSysName'
      Size = 40
    end
    object cdsParamsSysVersion2: TStringField
      FieldName = 'sSysVersion'
      Size = 5
    end
    object cdsParamsLogPath: TStringField
      FieldName = 'sLogPath'
      Size = 60
    end
    object cdsParamnTimeOut2: TIntegerField
      FieldName = 'nTimeOut'
    end
    object cdsParambCommandLog2: TBooleanField
      FieldName = 'bCommandLog'
    end
    object cdsParambResponseLog: TBooleanField
      FieldName = 'bResponseLog'
    end
    object cdsParamdUptTime2: TDateTimeField
      FieldName = 'dUptTime'
    end
    object cdsParamsUptName2: TStringField
      FieldName = 'sUptName'
    end
    object cdsParambShowUI2: TBooleanField
      FieldName = 'bShowUI'
    end
    object cdsParamnVersion: TIntegerField
      FieldName = 'nVersion'
    end
    object cdsParamsSecuritytype: TStringField
      FieldName = 'sSecuritytype'
      Size = 5
    end
    object cdsParamnToID: TIntegerField
      FieldName = 'nToID'
    end
    object cdsParamnConnectionID: TIntegerField
      FieldName = 'nConnectionID'
    end
    object cdsParamnForwardTolerance: TIntegerField
      FieldName = 'nForwardTolerance'
    end
    object cdsParamnBackwardTolerance: TIntegerField
      FieldName = 'nBackwardTolerance'
    end
    object cdsParamnCurrency: TIntegerField
      FieldName = 'nCurrency'
    end
    object cdsParamnFromID: TIntegerField
      FieldName = 'nFromID'
    end
    object cdsParamnTimeZoneOffset: TIntegerField
      FieldName = 'nTimeZoneOffset'
    end
    object cdsParamsCountryNumber: TStringField
      FieldName = 'sCountryNumber'
      Size = 6
    end
    object cdsParamnTimeZoneOffset2: TIntegerField
      FieldName = 'nTimeZoneOffset2'
    end
    object cdsParamsKey: TStringField
      FieldName = 'sKey'
      Size = 100
    end
  end
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
    Left = 48
    Top = 16
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
end
