object MainUIModule: TMainUIModule
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  OnDestroy = DataModuleDestroy
  Left = 518
  Top = 301
  Height = 219
  Width = 290
  object dsUserBuffer: TDataSource
    DataSet = UserBuffer
    Left = 76
    Top = 96
  end
  object UserBuffer: TVirtualTable
    Left = 76
    Top = 36
    Data = {03000000000000000000}
  end
end
