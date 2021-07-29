object dtmMain: TdtmMain
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  OnDestroy = DataModuleDestroy
  Left = 260
  Top = 133
  Height = 297
  Width = 401
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
    Left = 128
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
      Visible = False
    end
  end
  object cdsSGSParam: TClientDataSet
    Active = True
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'Group'
        DataType = ftString
        Size = 2
      end
      item
        Name = 'DataAreaNo'
        DataType = ftInteger
      end
      item
        Name = 'DataArea'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'CompCode'
        DataType = ftString
        Size = 5
      end
      item
        Name = 'CompName'
        DataType = ftString
        Size = 50
      end
      item
        Name = 'Alias'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'UserID'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'Password'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'FirstDate'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'DataPath1'
        DataType = ftString
        Size = 100
      end
      item
        Name = 'DataPath2'
        DataType = ftString
        Size = 100
      end
      item
        Name = 'AdministratorMail1'
        DataType = ftString
        Size = 100
      end
      item
        Name = 'AdministratorMail2'
        DataType = ftString
        Size = 100
      end
      item
        Name = 'Version'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'CasID'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'SrcCode'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'SrcIP'
        DataType = ftString
        Size = 50
      end
      item
        Name = 'SrcType'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'DstCode'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'DstIP'
        DataType = ftString
        Size = 50
      end
      item
        Name = 'DstType'
        DataType = ftString
        Size = 20
      end>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 208
    Top = 16
    Data = {
      720200009619E0BD01000000180000001500000000000300000072020547726F
      757001004900000001000557494454480200020002000A44617461417265614E
      6F04000100000000000844617461417265610100490000000100055749445448
      02000200140008436F6D70436F64650100490000000100055749445448020002
      00050008436F6D704E616D650100490000000100055749445448020002003200
      05416C6961730100490000000100055749445448020002001400065573657249
      4401004900000001000557494454480200020014000850617373776F72640100
      4900000001000557494454480200020014000946697273744461746501004900
      0000010005574944544802000200140009446174615061746831010049000000
      0100055749445448020002006400094461746150617468320100490000000100
      0557494454480200020064001241646D696E6973747261746F724D61696C3101
      004900000001000557494454480200020064001241646D696E6973747261746F
      724D61696C320100490000000100055749445448020002006400075665727369
      6F6E010049000000010005574944544802000200140005436173494401004900
      0000010005574944544802000200140007537263436F64650100490000000100
      0557494454480200020014000553726349500100490000000100055749445448
      0200020032000753726354797065010049000000010005574944544802000200
      140007447374436F646501004900000001000557494454480200020014000544
      7374495001004900000001000557494454480200020032000744737454797065
      01004900000001000557494454480200020014000000}
    object cdsSGSParamGroup: TStringField
      FieldName = 'Group'
      Size = 2
    end
    object cdsSGSParamDataAreaNo: TIntegerField
      FieldName = 'DataAreaNo'
      Visible = False
    end
    object cdsSGSParamDataArea: TStringField
      FieldName = 'DataArea'
      Visible = False
    end
    object cdsSGSParamCompCode: TStringField
      FieldName = 'CompCode'
      Visible = False
      Size = 5
    end
    object cdsSGSParamCompName: TStringField
      DisplayWidth = 20
      FieldName = 'CompName'
      Size = 50
    end
    object cdsSGSParamAlias: TStringField
      FieldName = 'Alias'
    end
    object cdsSGSParamUserID: TStringField
      FieldName = 'UserID'
    end
    object cdsSGSParamPassword: TStringField
      FieldName = 'Password'
    end
    object cdsSGSParamFirstDate: TStringField
      FieldName = 'FirstDate'
    end
    object cdsSGSParamDataPath1: TStringField
      FieldName = 'DataPath1'
      Size = 100
    end
    object cdsSGSParamDataPath2: TStringField
      FieldName = 'DataPath2'
      Size = 100
    end
    object cdsSGSParamAdministratorMail1: TStringField
      FieldName = 'AdministratorMail1'
      Size = 100
    end
    object cdsSGSParamAdministratorMail2: TStringField
      FieldName = 'AdministratorMail2'
      Size = 100
    end
    object cdsSGSParamVersion: TStringField
      FieldName = 'Version'
    end
    object cdsSGSParamCasID: TStringField
      FieldName = 'CasID'
    end
    object cdsSGSParamSrcCode: TStringField
      FieldName = 'SrcCode'
    end
    object cdsSGSParamSrcIP: TStringField
      FieldName = 'SrcIP'
      Size = 50
    end
    object cdsSGSParamSrcType: TStringField
      FieldName = 'SrcType'
    end
    object cdsSGSParamDstCode: TStringField
      FieldName = 'DstCode'
    end
    object cdsSGSParamDstIP: TStringField
      FieldName = 'DstIP'
      Size = 50
    end
    object cdsSGSParamDstType: TStringField
      FieldName = 'DstType'
    end
  end
  object adoQryQueryData: TADOQuery
    Parameters = <>
    Left = 48
    Top = 16
  end
  object adoQryNightRunLog: TADOQuery
    Parameters = <>
    Left = 48
    Top = 64
  end
  object adoQryQueryLogData: TADOQuery
    Parameters = <>
    Left = 280
    Top = 128
  end
  object cdsQueryLogData: TClientDataSet
    Active = True
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'DstMsgID'
        DataType = ftString
        Size = 10
      end
      item
        Name = 'ModeType'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'InstQuery'
        DataType = ftString
        Size = 10
      end
      item
        Name = 'DstCode'
        DataType = ftString
        Size = 10
      end
      item
        Name = 'SrcCode'
        DataType = ftString
        Size = 10
      end
      item
        Name = 'HandleEDateTime'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'ExecDateTime'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'SrcMsgID'
        DataType = ftString
        Size = 10
      end
      item
        Name = 'ErrorCode'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'ErrorMsg'
        DataType = ftString
        Size = 1000
      end>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 280
    Top = 192
    Data = {
      470100009619E0BD01000000180000000A000000000003000000470108447374
      4D736749440100490000000100055749445448020002000A00084D6F64655479
      7065010049000000010005574944544802000200140009496E73745175657279
      0100490000000100055749445448020002000A0007447374436F646501004900
      00000100055749445448020002000A0007537263436F64650100490000000100
      055749445448020002000A000F48616E646C65454461746554696D6501004900
      000001000557494454480200020014000C457865634461746554696D65010049
      0000000100055749445448020002001400085372634D73674944010049000000
      0100055749445448020002000A00094572726F72436F64650100490000000100
      055749445448020002001400084572726F724D73670200490000000100055749
      44544802000200E8030000}
    object cdsQueryLogDataDstMsgID: TStringField
      FieldName = 'DstMsgID'
      Size = 10
    end
    object cdsQueryLogDataModeType: TStringField
      FieldName = 'ModeType'
    end
    object cdsQueryLogDataInstQuery: TStringField
      FieldName = 'InstQuery'
      Size = 10
    end
    object cdsQueryLogDataDstCode: TStringField
      FieldName = 'DstCode'
      Size = 10
    end
    object cdsQueryLogDataSrcCode: TStringField
      FieldName = 'SrcCode'
      Size = 10
    end
    object cdsQueryLogDataHandleEDateTime: TStringField
      FieldName = 'HandleEDateTime'
    end
    object cdsQueryLogDataExecDateTime: TStringField
      FieldName = 'ExecDateTime'
    end
    object cdsQueryLogDataSrcMsgID: TStringField
      FieldName = 'SrcMsgID'
      Size = 10
    end
    object cdsQueryLogDataErrorCode: TStringField
      FieldName = 'ErrorCode'
      Visible = False
    end
    object cdsQueryLogDataErrorMsg: TStringField
      FieldName = 'ErrorMsg'
      Size = 1000
    end
  end
  object cdsChannelID: TClientDataSet
    Active = True
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'ProductCode'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'Provider'
        DataType = ftString
        Size = 40
      end
      item
        Name = 'ChannelName'
        DataType = ftString
        Size = 40
      end>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 280
    Top = 72
    Data = {
      770000009619E0BD01000000180000000300000000000300000077000B50726F
      64756374436F646501004900000001000557494454480200020014000850726F
      766964657201004900000001000557494454480200020028000B4368616E6E65
      6C4E616D6501004900000001000557494454480200020028000000}
    object cdsChannelIDProductCode: TStringField
      FieldName = 'ProductCode'
    end
    object cdsChannelIDProvider: TStringField
      FieldName = 'Provider'
      Size = 40
    end
    object cdsChannelIDChannelName: TStringField
      FieldName = 'ChannelName'
      Size = 40
    end
  end
  object adoQryLoadExecCounts: TADOQuery
    Parameters = <>
    Left = 56
    Top = 128
  end
end
