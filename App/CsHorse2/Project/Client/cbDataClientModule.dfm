object DataClientModule: TDataClientModule
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  OnDestroy = DataModuleDestroy
  Left = 429
  Top = 245
  Height = 196
  Width = 366
  object dsSo021: TDataSource
    AutoEdit = False
    DataSet = BfSo021
    Left = 180
    Top = 22
  end
  object BfSo021: TVirtualTable
    Left = 183
    Top = 76
    Data = {03000000000000000000}
  end
  object dsCd042: TDataSource
    AutoEdit = False
    DataSet = BfCd042
    Left = 24
    Top = 22
  end
  object BfCd042: TVirtualTable
    Left = 24
    Top = 76
    Data = {03000000000000000000}
  end
  object dsSo022: TDataSource
    AutoEdit = False
    DataSet = BfSo022
    Left = 76
    Top = 22
  end
  object BfSo022: TVirtualTable
    Left = 77
    Top = 76
    Data = {03000000000000000000}
  end
  object dsSo023: TDataSource
    AutoEdit = False
    DataSet = BfSo023
    Left = 128
    Top = 22
  end
  object BfSo023: TVirtualTable
    Left = 130
    Top = 76
    Data = {03000000000000000000}
  end
  object BfUserList: TVirtualTable
    Left = 240
    Top = 76
    Data = {03000000000000000000}
  end
  object dsUserList: TDataSource
    AutoEdit = False
    DataSet = BfUserList
    Left = 240
    Top = 21
  end
end
