object dtmMain: TdtmMain
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  OnDestroy = DataModuleDestroy
  Left = 446
  Top = 100
  Height = 213
  Width = 313
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
      end
      item
        Name = 'nMaxCommandCount'
        DataType = ftInteger
      end
      item
        Name = 'nCmdResentTimes'
        DataType = ftInteger
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
        Name = 'nCmdRefreshRate1'
        DataType = ftInteger
      end
      item
        Name = 'nCmdRefreshRate2'
        DataType = ftInteger
      end>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 32
    Top = 16
    Data = {
      FC0200009619E0BD01000000180000001F000000000003000000FC0209735365
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
      4454480200020064000D6E506F70756C6174696F6E4944040001000000000010
      6E4D6178436F6D6D616E64436F756E7404000100000000000F6E436D64526573
      656E7454696D65730400010000000000097353486F7454696D65010049000000
      0100055749445448020002000400097345486F7454696D650100490000000100
      055749445448020002000400106E436D64526566726573685261746531040001
      0000000000106E436D6452656672657368526174653204000100000000000000}
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
    object cdsParamnMaxCommandCount: TIntegerField
      FieldName = 'nMaxCommandCount'
    end
    object cdsParamnCmdResentTimes: TIntegerField
      FieldName = 'nCmdResentTimes'
    end
    object cdsParamsSHotTime: TStringField
      FieldName = 'sSHotTime'
      Size = 4
    end
    object cdsParamsEHotTime: TStringField
      FieldName = 'sEHotTime'
      Size = 4
    end
    object cdsParamnCmdRefreshRate1: TIntegerField
      FieldName = 'nCmdRefreshRate1'
    end
    object cdsParamnCmdRefreshRate2: TIntegerField
      FieldName = 'nCmdRefreshRate2'
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
  object adoQrySendNdsData: TADOQuery
    Parameters = <>
    Left = 48
    Top = 64
  end
  object cdsXmlData: TClientDataSet
    Active = True
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'SeqNO'
        DataType = ftString
        Size = 8
      end
      item
        Name = 'CompCode'
        DataType = ftInteger
      end
      item
        Name = 'CompName'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'CommandID'
        DataType = ftString
        Size = 4
      end
      item
        Name = 'SubscriberID'
        DataType = ftString
        Size = 8
      end
      item
        Name = 'IccNo'
        DataType = ftString
        Size = 12
      end
      item
        Name = 'SubBeginDate'
        DataType = ftDateTime
      end
      item
        Name = 'SubEndDate'
        DataType = ftDateTime
      end
      item
        Name = 'PinCode'
        DataType = ftInteger
      end
      item
        Name = 'PopulationID'
        DataType = ftInteger
      end
      item
        Name = 'RegionKey'
        DataType = ftString
        Size = 8
      end
      item
        Name = 'ReportbackAvailability'
        DataType = ftString
        Size = 1
      end
      item
        Name = 'ReportbackDate'
        DataType = ftInteger
      end
      item
        Name = 'Notes'
        DataType = ftString
        Size = 4000
      end
      item
        Name = 'CmdStatus'
        DataType = ftString
        Size = 1
      end
      item
        Name = 'ErrorCode'
        DataType = ftString
        Size = 5
      end
      item
        Name = 'ErrorDesc'
        DataType = ftString
        Size = 30
      end
      item
        Name = 'Operator'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'ResentTimes'
        DataType = ftInteger
      end
      item
        Name = 'ProcessingDate'
        DataType = ftDateTime
      end
      item
        Name = 'UpdTime'
        DataType = ftDateTime
      end
      item
        Name = 'PinControl'
        DataType = ftString
        Size = 1
      end
      item
        Name = 'ServiceID'
        DataType = ftString
        Size = 5
      end
      item
        Name = 'ResponseData'
        DataType = ftString
        Size = 1024
      end
      item
        Name = 'ClientIP'
        DataType = ftString
        Size = 20
      end>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 48
    Top = 120
    Data = {
      AF0200009619E0BD010000001800000019000000000003000000AF0205536571
      4E4F010049000000010005574944544802000200080008436F6D70436F646504
      0001000000000008436F6D704E616D6501004900000001000557494454480200
      0200140009436F6D6D616E644944010049000000010005574944544802000200
      04000C5375627363726962657249440100490000000100055749445448020002
      000800054963634E6F0100490000000100055749445448020002000C000C5375
      62426567696E4461746508000800000000000A537562456E6444617465080008
      00000000000750696E436F646504000100000000000C506F70756C6174696F6E
      4944040001000000000009526567696F6E4B6579010049000000010005574944
      5448020002000800165265706F72746261636B417661696C6162696C69747901
      004900000001000557494454480200020001000E5265706F72746261636B4461
      74650400010000000000054E6F74657302004900000001000557494454480200
      0200A00F09436D64537461747573010049000000010005574944544802000200
      0100094572726F72436F64650100490000000100055749445448020002000500
      094572726F72446573630100490000000100055749445448020002001E00084F
      70657261746F7201004900000001000557494454480200020014000B52657365
      6E7454696D657304000100000000000E50726F63657373696E67446174650800
      0800000000000755706454696D6508000800000000000A50696E436F6E74726F
      6C01004900000001000557494454480200020001000953657276696365494401
      004900000001000557494454480200020005000C526573706F6E736544617461
      020049000000010005574944544802000200000408436C69656E744950010049
      00000001000557494454480200020014000000}
    object cdsXmlDataSeqNO: TStringField
      FieldName = 'SeqNO'
      Size = 8
    end
    object cdsXmlDataCompCode: TIntegerField
      FieldName = 'CompCode'
    end
    object cdsXmlDataCompName: TStringField
      FieldName = 'CompName'
    end
    object cdsXmlDataCommandID: TStringField
      FieldName = 'CommandID'
      Size = 4
    end
    object cdsXmlDataSubscriberID: TStringField
      FieldName = 'SubscriberID'
      Size = 8
    end
    object cdsXmlDataIccNo: TStringField
      FieldName = 'IccNo'
      Size = 12
    end
    object cdsXmlDataSubBeginDate: TDateTimeField
      FieldName = 'SubBeginDate'
    end
    object cdsXmlDataSubEndDate: TDateTimeField
      FieldName = 'SubEndDate'
    end
    object cdsXmlDataPinCode: TIntegerField
      FieldName = 'PinCode'
    end
    object cdsXmlDataPopulationID: TIntegerField
      FieldName = 'PopulationID'
    end
    object cdsXmlDataRegionKey: TStringField
      FieldName = 'RegionKey'
      Size = 8
    end
    object cdsXmlDataReportbackAvailability: TStringField
      FieldName = 'ReportbackAvailability'
      Size = 1
    end
    object cdsXmlDataReportbackDate: TIntegerField
      FieldName = 'ReportbackDate'
    end
    object cdsXmlDataNotes: TStringField
      FieldName = 'Notes'
      Size = 4000
    end
    object cdsXmlDataCmdStatus: TStringField
      FieldName = 'CmdStatus'
      Size = 1
    end
    object cdsXmlDataErrorCode: TStringField
      FieldName = 'ErrorCode'
      Size = 5
    end
    object cdsXmlDataErrorDesc: TStringField
      FieldName = 'ErrorDesc'
      Size = 30
    end
    object cdsXmlDataOperator: TStringField
      FieldName = 'Operator'
    end
    object cdsXmlDataResentTimes: TIntegerField
      FieldName = 'ResentTimes'
    end
    object cdsXmlDataProcessingDate: TDateTimeField
      FieldName = 'ProcessingDate'
    end
    object cdsXmlDataUpdTime: TDateTimeField
      FieldName = 'UpdTime'
    end
    object cdsXmlDataPinControl: TStringField
      FieldName = 'PinControl'
      Size = 1
    end
    object cdsXmlDataServiceID: TStringField
      FieldName = 'ServiceID'
      Size = 5
    end
    object cdsXmlDataResponseData: TStringField
      FieldName = 'ResponseData'
      Size = 1024
    end
    object cdsXmlDataClientIP: TStringField
      FieldName = 'ClientIP'
    end
  end
end
