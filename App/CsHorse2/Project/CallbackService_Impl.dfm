object CallbackService: TCallbackService
  OldCreateOrder = True
  OnCreate = RORemoteDataModuleCreate
  OnDestroy = RORemoteDataModuleDestroy
  SessionManager = ROServerModule.ROSessionManager
  EventRepository = ROServerModule.ROSrvCallbackEventRepository
  Left = 514
  Top = 270
  Height = 300
  Width = 300
end
