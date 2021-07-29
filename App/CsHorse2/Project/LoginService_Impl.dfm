object LoginService: TLoginService
  OldCreateOrder = True
  OnCreate = RORemoteDataModuleCreate
  OnDestroy = RORemoteDataModuleDestroy
  SessionManager = ROServerModule.ROSessionManager
  EventRepository = ROServerModule.ROSrvCallbackEventRepository
  Height = 245
  Width = 300
end
