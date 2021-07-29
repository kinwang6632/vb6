object ROClientModule: TROClientModule
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  OnDestroy = DataModuleDestroy
  Left = 589
  Top = 265
  Height = 337
  Width = 368
  object ROSuperTcpChannel: TROSuperTcpChannel
    Host = '127.0.0.1'
    ConnectionWaitTimeout = 5000
    ServerLocators = <>
    DispatchOptions = []
    Left = 60
    Top = 29
  end
  object ROLoginSrvBinMessage: TROBinMessage
    Left = 192
    Top = 29
  end
  object ROAnnSrvBinMessage: TROBinMessage
    Left = 192
    Top = 90
  end
  object ROCallbackSrvBinMessage: TROBinMessage
    Left = 192
    Top = 151
  end
  object ROCallbackSrv2BinMessage: TROBinMessage
    Left = 192
    Top = 212
  end
end
