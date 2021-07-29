object dtmMain: TdtmMain
  OldCreateOrder = False
  Left = 519
  Top = 211
  Height = 310
  Width = 364
  object cdsNetworkID: TClientDataSet
    Aggregates = <>
    FieldDefs = <>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 40
    Top = 32
    object cdsNetworkIDCompCode: TIntegerField
      DisplayLabel = #20844#21496#20195#30908
      FieldName = 'CompCode'
    end
    object cdsNetworkIDCompName: TStringField
      DisplayLabel = #20844#21496#21517#31281
      DisplayWidth = 20
      FieldName = 'CompName'
      Size = 50
    end
    object cdsNetworkIDNetworkID: TStringField
      FieldName = 'NetworkID'
    end
    object cdsNetworkIDOperation: TStringField
      DisplayLabel = #24190#31186#20839#35201#29983#25928
      FieldName = 'Operation'
    end
  end
  object cdsOperation: TClientDataSet
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'sCommandType'
        DataType = ftString
        Size = 2
      end>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 285
    Top = 43
    object cdsOperationsCommandType: TStringField
      FieldName = 'sCommandType'
      Size = 2
    end
  end
  object cdsFeedback: TClientDataSet
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'sCommandType'
        DataType = ftString
        Size = 2
      end>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 237
    Top = 43
    object cdsFeedbacksCommandType: TStringField
      FieldName = 'sCommandType'
      Size = 2
    end
  end
  object cdsProduct: TClientDataSet
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'sCommandType'
        DataType = ftString
        Size = 2
      end>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 169
    Top = 36
    object cdsProductsCommandType: TStringField
      FieldName = 'sCommandType'
      Size = 2
    end
  end
  object cdsEmm: TClientDataSet
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'sCommandType'
        DataType = ftString
        Size = 2
      end
      item
        Name = 'sBroadcastMode'
        DataType = ftString
        Size = 1
      end
      item
        Name = 'sAddressType'
        DataType = ftString
        Size = 1
      end>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 104
    Top = 96
    object cdsEmmsCommandType: TStringField
      FieldName = 'sCommandType'
      Size = 2
    end
    object cdsEmmsBroadcastMode: TStringField
      FieldName = 'sBroadcastMode'
      Size = 1
    end
    object cdsEmmsAddressType: TStringField
      FieldName = 'sAddressType'
      Size = 1
    end
  end
  object cdsControl: TClientDataSet
    Active = True
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'sCommandType'
        DataType = ftString
        Size = 2
      end
      item
        Name = 'sBroadcastMode'
        DataType = ftString
        Size = 2
      end
      item
        Name = 'sAddressType'
        DataType = ftString
        Size = 2
      end>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 32
    Top = 96
    Data = {
      7F0000009619E0BD0100000018000000030000000000030000007F000C73436F
      6D6D616E645479706501004900000001000557494454480200020002000E7342
      726F6164636173744D6F64650100490000000100055749445448020002000200
      0C73416464726573735479706501004900000001000557494454480200020002
      000000}
    object cdsControlsCommandType: TStringField
      FieldName = 'sCommandType'
      Size = 2
    end
    object cdsControlsBroadcastMode: TStringField
      FieldName = 'sBroadcastMode'
      Size = 2
    end
    object cdsControlsAddressType: TStringField
      FieldName = 'sAddressType'
      Size = 2
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
    Left = 152
    Top = 96
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
    Params = <>
    Left = 184
    Top = 176
    object cdsParamsServerIP: TStringField
      FieldName = 'sServerIP'
      Size = 15
    end
    object cdsParamnSPortNo: TIntegerField
      FieldName = 'nSPortNo'
    end
    object cdsParamnRPortNo: TIntegerField
      FieldName = 'nRPortNo'
    end
    object cdsParamsSysName: TStringField
      FieldName = 'sSysName'
      Size = 40
    end
    object cdsParamsSysVersion: TStringField
      FieldName = 'sSysVersion'
      Size = 5
    end
    object cdsParamsLogPath: TStringField
      FieldName = 'sLogPath'
      Size = 60
    end
    object cdsParamnTimeOut: TIntegerField
      FieldName = 'nTimeOut'
    end
    object cdsParamnMaxCommandCount: TIntegerField
      FieldName = 'nMaxCommandCount'
    end
    object cdsParambCommandLog: TBooleanField
      FieldName = 'bCommandLog'
    end
    object cdsParambErrorLog: TBooleanField
      FieldName = 'bErrorLog'
    end
    object cdsParamnResponseLog: TBooleanField
      FieldName = 'nResponseLog'
    end
    object cdsParamdUptTime: TDateTimeField
      FieldName = 'dUptTime'
    end
    object cdsParamsUptName: TStringField
      FieldName = 'sUptName'
    end
    object cdsParambSetZipCode: TBooleanField
      FieldName = 'bSetZipCode'
    end
    object cdsParamnCreditLimit: TIntegerField
      FieldName = 'nCreditLimit'
    end
    object cdsParamnSourceID: TIntegerField
      FieldName = 'nSourceID'
    end
    object cdsParamnDestID: TIntegerField
      FieldName = 'nDestID'
    end
    object cdsParamnMopPPID: TIntegerField
      FieldName = 'nMopPPID'
    end
    object cdsParamsBroadcastSDate: TStringField
      FieldName = 'sBroadcastSDate'
      EditMask = '!9999/99/99;1;_'
      Size = 10
    end
    object cdsParamsBroadcastEDate: TStringField
      FieldName = 'sBroadcastEDate'
      EditMask = '!9999/99/99;1;_'
      Size = 10
    end
    object cdsParamnCmdRefreshRate1: TIntegerField
      FieldName = 'nCmdRefreshRate1'
    end
    object cdsParamnCmdRefreshRate2: TIntegerField
      FieldName = 'nCmdRefreshRate2'
    end
    object cdsParamnCmdResentTimes: TIntegerField
      FieldName = 'nCmdResentTimes'
    end
    object cdsParambShowUI: TBooleanField
      FieldName = 'bShowUI'
    end
    object cdsParambAssignBroadcastDate: TBooleanField
      FieldName = 'bAssignBroadcastDate'
    end
    object cdsParamsSHotTime: TStringField
      FieldName = 'sSHotTime'
      Size = 4
    end
    object cdsParamsEHotTime: TStringField
      FieldName = 'sEHotTime'
      Size = 4
    end
  end
end
