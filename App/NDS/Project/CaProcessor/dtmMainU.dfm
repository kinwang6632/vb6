object dtmMain: TdtmMain
  OldCreateOrder = False
  Left = 258
  Top = 107
  Height = 165
  Width = 169
  object cdsParam: TClientDataSet
    Active = True
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'sServerIP'
        DataType = ftString
        Size = 15
      end
      item
        Name = 'nSPortNo'
        DataType = ftInteger
      end
      item
        Name = 'nRPortNo'
        DataType = ftInteger
      end
      item
        Name = 'sSysName'
        DataType = ftString
        Size = 40
      end
      item
        Name = 'sSysVersion'
        DataType = ftString
        Size = 5
      end
      item
        Name = 'sLogPath'
        DataType = ftString
        Size = 60
      end
      item
        Name = 'nTimeOut'
        DataType = ftInteger
      end
      item
        Name = 'bCommandLog'
        DataType = ftBoolean
      end
      item
        Name = 'bResponseLog'
        DataType = ftBoolean
      end
      item
        Name = 'dUptTime'
        DataType = ftDateTime
      end
      item
        Name = 'sUptName'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'bShowUI'
        DataType = ftBoolean
      end
      item
        Name = 'nVersion'
        DataType = ftInteger
      end
      item
        Name = 'sSecuritytype'
        DataType = ftString
        Size = 5
      end
      item
        Name = 'nToID'
        DataType = ftInteger
      end
      item
        Name = 'nConnectionID'
        DataType = ftInteger
      end
      item
        Name = 'nForwardTolerance'
        DataType = ftInteger
      end
      item
        Name = 'nBackwardTolerance'
        DataType = ftInteger
      end
      item
        Name = 'nCurrency'
        DataType = ftInteger
      end
      item
        Name = 'nFromID'
        DataType = ftInteger
      end
      item
        Name = 'nTimeZoneOffset'
        DataType = ftInteger
      end
      item
        Name = 'sCountryNumber'
        DataType = ftString
        Size = 6
      end
      item
        Name = 'nTimeZoneOffset2'
        DataType = ftInteger
      end
      item
        Name = 'sKey'
        DataType = ftString
        Size = 100
      end
      item
        Name = 'nPopulationID'
        DataType = ftInteger
      end>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 32
    Top = 16
    Data = {
      5D0200009619E0BD0100000018000000190000000000030000005D0209735365
      7276657249500100490000000100055749445448020002000F00086E53506F72
      744E6F0400010000000000086E52506F72744E6F040001000000000008735379
      734E616D6501004900000001000557494454480200020028000B735379735665
      7273696F6E010049000000010005574944544802000200050008734C6F675061
      74680100490000000100055749445448020002003C00086E54696D654F757404
      000100000000000B62436F6D6D616E644C6F6702000300000000000C62526573
      706F6E73654C6F670200030000000000086455707454696D6508000800000000
      0008735570744E616D6501004900000001000557494454480200020014000762
      53686F7755490200030000000000086E56657273696F6E04000100000000000D
      7353656375726974797479706501004900000001000557494454480200020005
      00056E546F494404000100000000000D6E436F6E6E656374696F6E4944040001
      0000000000116E466F7277617264546F6C6572616E6365040001000000000012
      6E4261636B77617264546F6C6572616E63650400010000000000096E43757272
      656E63790400010000000000076E46726F6D494404000100000000000F6E5469
      6D655A6F6E654F666673657404000100000000000E73436F756E7472794E756D
      6265720100490000000100055749445448020002000600106E54696D655A6F6E
      654F666673657432040001000000000004734B65790100490000000100055749
      4454480200020064000D6E506F70756C6174696F6E4944040001000000000000
      00}
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
    object cdsParamnPopulationID: TIntegerField
      FieldName = 'nPopulationID'
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
    Left = 80
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
