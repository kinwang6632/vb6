object dtmMain: TdtmMain
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  OnDestroy = DataModuleDestroy
  Left = 372
  Top = 226
  Height = 375
  Width = 544
  object qryActivityCodeStat: TQuery
    DatabaseName = 'symposium'
    SQL.Strings = (
      
        'select distinct ActivityCode, ActivityCodeName, sum(Occurrences)' +
        '  Occurrence'
      'from iActivityCodeStat'
      'where ActivityCode=""'
      'group by ActivityCode'
      'order by ActivityCode')
    Left = 68
    Top = 80
    object qryActivityCodeStatActivityCode: TStringField
      FieldName = 'ActivityCode'
      Origin = 'SYMPOSIUM.iActivityCodeStat.ActivityCode'
      Size = 32
    end
    object qryActivityCodeStatActivityCodeName: TStringField
      FieldName = 'ActivityCodeName'
      Origin = 'SYMPOSIUM.iActivityCodeStat.ActivityCodeName'
      Size = 200
    end
    object qryActivityCodeStatOccurrence: TIntegerField
      FieldName = 'Occurrence'
      Origin = 'SYMPOSIUM.iActivityCodeStat.Occurrences'
    end
  end
  object cdsActivityCodeStat: TClientDataSet
    Active = True
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'nGroup'
        DataType = ftInteger
      end
      item
        Name = 'sActivityCode'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'sActivityCodeName'
        DataType = ftString
        Size = 200
      end
      item
        Name = 'nOccurrence1'
        DataType = ftInteger
      end
      item
        Name = 'nOccurrence2'
        DataType = ftInteger
      end
      item
        Name = 'nOccurrence3'
        DataType = ftInteger
      end
      item
        Name = 'nOccurrence'
        DataType = ftInteger
      end>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 208
    Top = 80
    Data = {
      C40000009619E0BD010000001800000007000000000003000000C400066E4772
      6F757004000100000000000D734163746976697479436F646501004900000001
      0005574944544802000200140011734163746976697479436F64654E616D6501
      0049000000010005574944544802000200C8000C6E4F6363757272656E636531
      04000100000000000C6E4F6363757272656E63653204000100000000000C6E4F
      6363757272656E63653304000100000000000B6E4F6363757272656E63650400
      0100000000000000}
    object cdsActivityCodeStatnGroup: TIntegerField
      FieldName = 'nGroup'
    end
    object cdsActivityCodeStatsActivityCode: TStringField
      FieldName = 'sActivityCode'
    end
    object cdsActivityCodeStatsActivityCodeName: TStringField
      FieldName = 'sActivityCodeName'
      Size = 200
    end
    object cdsActivityCodeStatnOccurrence1: TIntegerField
      FieldName = 'nOccurrence1'
    end
    object cdsActivityCodeStatnOccurrence2: TIntegerField
      FieldName = 'nOccurrence2'
    end
    object cdsActivityCodeStatnOccurrence3: TIntegerField
      FieldName = 'nOccurrence3'
    end
    object cdsActivityCodeStatnOccurrence: TIntegerField
      FieldName = 'nOccurrence'
    end
  end
  object qryAgentPerformance: TQuery
    DatabaseName = 'symposium'
    SQL.Strings = (
      'select AgentLogin, AgentGivenName, '
      'SUM(TalkTime) TalkTime, '
      'Sum(NotReadyTime) NotReadyTime, '
      'SUM(CallsOffered) CallsPresented, '
      'SUM(LoggedInTime) LoggedInTime, '
      'SUM(CallsAnswered) CallsAnswered'
      'from iAgentPerformanceStat '
      'where AgentLogin =""'
      'group by AgentLogin'
      'order by AgentLogin '
      ''
      ''
      '')
    Left = 68
    Top = 136
  end
  object qryAgentBySkillsetPerformance: TQuery
    DatabaseName = 'symposium'
    SQL.Strings = (
      'select AgentSurName,'
      'SUM(TalkTime) TalkTime, '
      'Sum(NotReadyTime) NotReadyTime, '
      'SUM(CallsAnswered) CallsAnswered,'
      'SUM(TalkTime+NotReadyTime) WorkTime'
      'from iAgentBySkillsetStat '
      
        'where Timestamp between "2001/10/23 00:00:01" and "2001/10/23 23' +
        ':59:59"'
      'and AgentLogin=""'
      'group by AgentLogin '
      'order by AgentLogin ')
    Left = 68
    Top = 196
    object qryAgentBySkillsetPerformanceAgentSurName: TStringField
      FieldName = 'AgentSurName'
      Size = 64
    end
    object qryAgentBySkillsetPerformanceTalkTime: TIntegerField
      FieldName = 'TalkTime'
    end
    object qryAgentBySkillsetPerformanceNotReadyTime: TIntegerField
      FieldName = 'NotReadyTime'
    end
    object qryAgentBySkillsetPerformanceCallsAnswered: TIntegerField
      FieldName = 'CallsAnswered'
    end
    object qryAgentBySkillsetPerformanceWorkTime: TIntegerField
      FieldName = 'WorkTime'
    end
  end
  object cdsAgentPerformance: TClientDataSet
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'Group'
        DataType = ftInteger
      end
      item
        Name = 'AgentLogin'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'AgentGivenName'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'CallsAnswered'
        DataType = ftInteger
      end
      item
        Name = 'TalkTime'
        DataType = ftInteger
      end
      item
        Name = 'NotReadyTime'
        DataType = ftInteger
      end
      item
        Name = 'CallsPresented'
        DataType = ftInteger
      end
      item
        Name = 'LoggedInTime'
        DataType = ftInteger
      end
      item
        Name = 'ShortCallsAns'
        DataType = ftInteger
      end>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 208
    Top = 136
    object cdsAgentPerformanceGroup: TIntegerField
      FieldName = 'Group'
    end
    object cdsAgentPerformanceAgentLogin: TStringField
      FieldName = 'AgentLogin'
    end
    object cdsAgentPerformanceAgentGivenName: TStringField
      FieldName = 'AgentGivenName'
    end
    object cdsAgentPerformanceCallsAnswered: TIntegerField
      FieldName = 'CallsAnswered'
    end
    object cdsAgentPerformanceTalkTime: TIntegerField
      FieldName = 'TalkTime'
    end
    object cdsAgentPerformanceNotReadyTime: TIntegerField
      FieldName = 'NotReadyTime'
    end
    object cdsAgentPerformanceCallsPresented: TIntegerField
      FieldName = 'CallsPresented'
    end
    object cdsAgentPerformanceLoggedInTime: TIntegerField
      FieldName = 'LoggedInTime'
    end
    object cdsAgentPerformanceShortCallsAns: TIntegerField
      FieldName = 'ShortCallsAns'
    end
  end
  object Table1: TTable
    DatabaseName = 'symposium'
    Left = 208
    Top = 24
  end
  object Database1: TDatabase
    AliasName = 'ICCM_PREVIEW_DSN'
    DatabaseName = 'symposium'
    SessionName = 'Default'
    Left = 68
    Top = 24
  end
end
