object dtmComm: TdtmComm
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  OnDestroy = DataModuleDestroy
  Left = 192
  Top = 114
  Height = 380
  Width = 492
  object InvConnection: TADOConnection
    LoginPrompt = False
    Provider = 'MSDAORA.1'
    Left = 22
    Top = 10
  end
  object adoINV008: TADOQuery
    Connection = InvConnection
    Parameters = <>
    Left = 18
    Top = 54
  end
  object LogConnection: TADOConnection
    LoginPrompt = False
    Provider = 'MSDAORA.1'
    Left = 46
    Top = 116
  end
  object adoLog: TADOQuery
    Connection = LogConnection
    CursorType = ctStatic
    Parameters = <>
    SQL.Strings = (
      'SELECT 1 TYPE,SENDTIME,                     '
      '       NVL(SENDSTATE,0) SENDSTATE,SENDCNT,'
      '       RETURNCODE,RETURNMSG              '
      '       FROM INV036LOG WHERE INVID = '#39'ABCDE'#39)
    Left = 36
    Top = 170
  end
end
