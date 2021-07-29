object LicenceManager: TLicenceManager
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  OnDestroy = DataModuleDestroy
  Left = 459
  Top = 278
  Height = 207
  Width = 260
  object LbDES: TLbDES
    CipherMode = cmECB
    Left = 100
    Top = 60
  end
end
