object OraModule: TOraModule
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  OnDestroy = DataModuleDestroy
  Left = 513
  Top = 255
  Height = 203
  Width = 277
  object OraSession: TOraSession
    LoginPrompt = False
    Left = 44
    Top = 40
  end
  object OraTable: TOraTable
    Session = OraSession
    Left = 116
    Top = 40
  end
  object OraSQL: TOraSQL
    Session = OraSession
    Left = 44
    Top = 100
  end
  object OraQuery: TOraQuery
    Session = OraSession
    Left = 184
    Top = 40
  end
  object OraScript: TOraScript
    Session = OraSession
    Left = 116
    Top = 100
  end
end
