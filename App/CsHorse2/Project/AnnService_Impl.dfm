object AnnService: TAnnService
  OldCreateOrder = True
  OnCreate = RORemoteDataModuleCreate
  OnDestroy = RORemoteDataModuleDestroy
  SessionManager = ROServerModule.ROSessionManager
  Height = 300
  Width = 300
end
