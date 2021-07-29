object FTPControler: TFTPControler
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  Left = 660
  Top = 325
  Height = 181
  Width = 226
  object IdFTP: TIdFTP
    MaxLineAction = maException
    ProxySettings.ProxyType = fpcmNone
    ProxySettings.Port = 0
    Left = 44
    Top = 36
  end
end
