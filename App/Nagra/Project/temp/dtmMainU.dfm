object dtmMain: TdtmMain
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  OnDestroy = DataModuleDestroy
  Left = 537
  Top = 185
  Height = 150
  Width = 332
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
    Left = 184
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
        Name = 'nFormID'
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
      end>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 16
    Top = 8
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
    object cdsParamsSecuritytype: TStringField
      FieldName = 'sSecuritytype'
      Size = 1
    end
    object cdsParamnFormID: TIntegerField
      FieldName = 'nFormID'
    end
    object cdsParamnToID: TIntegerField
      FieldName = 'nToID'
    end
    object cdsParamnConnectionID: TIntegerField
      FieldName = 'nConnectionID'
    end
    object cdsParamsKey: TStringField
      FieldName = 'sKey'
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
  end
  object ClientDataSet1: TClientDataSet
    Aggregates = <>
    FieldDefs = <>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 104
    Top = 8
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
    object cdsParamsMisCommandPath: TStringField
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
    object cdsParamnResponseLog: TBooleanField
      FieldName = 'nResponseLog'
    end
    object cdsParamdUptTime: TDateTimeField
      FieldName = 'dUptTime'
    end
    object cdsParamsUptName: TStringField
      FieldName = 'sUptName'
    end
    object cdsParamnCmdRefreshRate: TIntegerField
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
    object cdsParamsSHotTime: TStringField
      FieldName = 'sSHotTime'
      Size = 4
    end
    object cdsParamsEHotTime: TStringField
      FieldName = 'sEHotTime'
      Size = 4
    end
  end
  object ClientDataSet2: TClientDataSet
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
        Name = 'nResponseLog'
        DataType = ftInteger
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
        Name = 'nFormID'
        DataType = ftInteger
      end
      item
        Name = 'nToID'
        DataType = ftInteger
      end>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 112
    Top = 56
    object StringField1: TStringField
      FieldName = 'sServerIP'
      Size = 15
    end
    object IntegerField1: TIntegerField
      FieldName = 'nSPortNo'
    end
    object IntegerField2: TIntegerField
      FieldName = 'nRPortNo'
    end
    object StringField2: TStringField
      FieldName = 'sSysName'
      Size = 40
    end
    object StringField3: TStringField
      FieldName = 'sSysVersion'
      Size = 5
    end
    object StringField4: TStringField
      FieldName = 'sLogPath'
      Size = 60
    end
    object IntegerField3: TIntegerField
      FieldName = 'nTimeOut'
    end
    object IntegerField4: TIntegerField
      FieldName = 'nMaxCommandCount'
    end
    object BooleanField1: TBooleanField
      FieldName = 'bCommandLog'
    end
    object IntegerField5: TIntegerField
      FieldName = 'nResponseLog'
    end
    object DateTimeField1: TDateTimeField
      FieldName = 'dUptTime'
    end
    object StringField5: TStringField
      FieldName = 'sUptName'
    end
    object IntegerField6: TIntegerField
      FieldName = 'nCmdRefreshRate1'
    end
    object IntegerField7: TIntegerField
      FieldName = 'nCmdRefreshRate2'
    end
    object IntegerField8: TIntegerField
      FieldName = 'nCmdResentTimes'
    end
    object BooleanField2: TBooleanField
      FieldName = 'bShowUI'
    end
    object StringField6: TStringField
      FieldName = 'sSHotTime'
      Size = 4
    end
    object StringField7: TStringField
      FieldName = 'sEHotTime'
      Size = 4
    end
    object IntegerField9: TIntegerField
      FieldName = 'nVersion'
    end
    object StringField8: TStringField
      FieldName = 'sSecuritytype'
      Size = 1
    end
    object IntegerField10: TIntegerField
      FieldName = 'nFormID'
    end
    object IntegerField11: TIntegerField
      FieldName = 'nToID'
    end
  end
end
