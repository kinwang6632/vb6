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
    Left = 96
    Top = 64
  end
end
