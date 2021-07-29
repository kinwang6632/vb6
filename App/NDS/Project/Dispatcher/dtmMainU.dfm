object dtmMain: TdtmMain
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  OnDestroy = DataModuleDestroy
  Left = 366
  Top = 191
  Height = 243
  Width = 287
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
    Top = 8
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
        Name = 'nMaxCommandCount'
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
        Name = 'nCmdRefreshRate1'
        DataType = ftInteger
      end
      item
        Name = 'nCmdRefreshRate2'
        DataType = ftInteger
      end
      item
        Name = 'nCmdResentTimes'
        DataType = ftInteger
      end
      item
        Name = 'bShowUI'
        DataType = ftBoolean
      end
      item
        Name = 'sSHotTime'
        DataType = ftString
        Size = 4
      end
      item
        Name = 'sEHotTime'
        DataType = ftString
        Size = 4
      end
      item
        Name = 'nVersion'
        DataType = ftInteger
      end
      item
        Name = 'sSecuritytype'
        DataType = ftString
        Size = 1
      end
      item
        Name = 'nFromID'
        DataType = ftInteger
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
        Name = 'sKey'
        DataType = ftString
        Size = 20
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
      end>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 16
    Top = 8
    Data = {
      E60200009619E0BD01000000180000001E000000000003000000E60209735365
      7276657249500100490000000100055749445448020002000F00086E53506F72
      744E6F0400010000000000086E52506F72744E6F040001000000000008735379
      734E616D6501004900000001000557494454480200020028000B735379735665
      7273696F6E010049000000010005574944544802000200050008734C6F675061
      74680100490000000100055749445448020002003C00086E54696D654F757404
      00010000000000106E4D6178436F6D6D616E64436F756E740400010000000000
      0B62436F6D6D616E644C6F6702000300000000000C62526573706F6E73654C6F
      670200030000000000086455707454696D65080008000000000008735570744E
      616D650100490000000100055749445448020002001400106E436D6452656672
      65736852617465310400010000000000106E436D645265667265736852617465
      3204000100000000000F6E436D64526573656E7454696D657304000100000000
      00076253686F7755490200030000000000097353486F7454696D650100490000
      000100055749445448020002000400097345486F7454696D6501004900000001
      00055749445448020002000400086E56657273696F6E04000100000000000D73
      5365637572697479747970650100490000000100055749445448020002000100
      076E46726F6D49440400010000000000056E546F494404000100000000000D6E
      436F6E6E656374696F6E4944040001000000000004734B657901004900000001
      00055749445448020002001400116E466F7277617264546F6C6572616E636504
      00010000000000126E4261636B77617264546F6C6572616E6365040001000000
      0000096E43757272656E637904000100000000000F6E54696D655A6F6E654F66
      6673657404000100000000000E73436F756E7472794E756D6265720100490000
      000100055749445448020002000600106E54696D655A6F6E654F666673657432
      04000100000000000000}
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
    object cdsParamnMaxCommandCount2: TIntegerField
      FieldName = 'nMaxCommandCount'
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
    object cdsParamnCmdRefreshRate1: TIntegerField
      FieldName = 'nCmdRefreshRate1'
    end
    object cdsParamnCmdRefreshRate22: TIntegerField
      FieldName = 'nCmdRefreshRate2'
    end
    object cdsParamnCmdResentTimes2: TIntegerField
      FieldName = 'nCmdResentTimes'
    end
    object cdsParambShowUI2: TBooleanField
      FieldName = 'bShowUI'
    end
    object cdsParamsSHotTime2: TStringField
      FieldName = 'sSHotTime'
      Size = 4
    end
    object cdsParamsEHotTime2: TStringField
      FieldName = 'sEHotTime'
      Size = 4
    end
    object cdsParamnVersion: TIntegerField
      FieldName = 'nVersion'
    end
  end
  object adoQrySendNdsData: TADOQuery
    Parameters = <>
    Left = 48
    Top = 64
  end
  object cdsDbLinkSet: TClientDataSet
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'Group'
        DataType = ftString
        Size = 2
      end
      item
        Name = 'ALIAS'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'USERID'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'PASSWORD'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'COMPCODE'
        DataType = ftString
        Size = 5
      end
      item
        Name = 'COMPNAME'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'PROCESSORIP'
        DataType = ftString
        Size = 20
      end>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 208
    Top = 8
    object cdsDbLinkSetGroup: TStringField
      DisplayLabel = #32068#21029
      FieldName = 'Group'
      Size = 2
    end
    object cdsDbLinkSetCOMPNAME: TStringField
      DisplayLabel = #20844#21496#21517#31281
      FieldName = 'COMPNAME'
    end
    object cdsDbLinkSetALIAS: TStringField
      DisplayWidth = 15
      FieldName = 'ALIAS'
    end
    object cdsDbLinkSetUSERID: TStringField
      DisplayLabel = #24115#34399
      DisplayWidth = 15
      FieldName = 'USERID'
    end
    object cdsDbLinkSetPASSWORD: TStringField
      DisplayLabel = #23494#30908
      DisplayWidth = 15
      FieldName = 'PASSWORD'
    end
    object cdsDbLinkSetPROCESSORIP: TStringField
      DisplayLabel = 'Processor_IP'
      DisplayWidth = 15
      FieldName = 'PROCESSORIP'
    end
    object cdsDbLinkSetCOMPCODE: TStringField
      DisplayLabel = #20844#21496#21029
      FieldName = 'COMPCODE'
      Size = 5
    end
  end
end
