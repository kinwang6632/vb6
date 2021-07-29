object ROServerModule: TROServerModule
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  OnDestroy = DataModuleDestroy
  Height = 273
  Width = 403
  object ROSuperTcpServer: TROSuperTcpServer
    Dispatchers = <
      item
        Name = 'ROServerBinMessage'
        Message = ROServerBinMessage
        Enabled = True
      end>
    DefaultResponse = 'ROSC:Invalid connection string'
    Left = 88
    Top = 28
  end
  object ROSessionManager: TROInMemorySessionManager
    Left = 88
    Top = 104
  end
  object ROServerBinMessage: TROBinMessage
    Left = 88
    Top = 180
  end
  object ROSrvCallbackEventRepository: TROInMemoryEventRepository
    Message = ROCallbackEventBinMessage
    SessionManager = ROSessionManager
    Left = 272
    Top = 28
  end
  object ROCallbackEventBinMessage: TROBinMessage
    Left = 272
    Top = 104
  end
end
