object frmMain: TfrmMain
  Left = 192
  Top = 107
  Width = 285
  Height = 195
  Caption = 'GiSchema之EngCaption欄位轉入程式'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -13
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  PixelsPerInch = 96
  TextHeight = 16
  object BitBtn1: TBitBtn
    Left = 80
    Top = 32
    Width = 105
    Height = 25
    Caption = '&S開始轉入'
    TabOrder = 0
    OnClick = BitBtn1Click
    Glyph.Data = {
      76010000424D7601000000000000760000002800000020000000100000000100
      04000000000000010000120B0000120B00001000000000000000000000000000
      800000800000008080008000000080008000808000007F7F7F00BFBFBF000000
      FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00555555555555
      555555555555555555555555555555555555555555FF55555555555559055555
      55555555577FF5555555555599905555555555557777F5555555555599905555
      555555557777FF5555555559999905555555555777777F555555559999990555
      5555557777777FF5555557990599905555555777757777F55555790555599055
      55557775555777FF5555555555599905555555555557777F5555555555559905
      555555555555777FF5555555555559905555555555555777FF55555555555579
      05555555555555777FF5555555555557905555555555555777FF555555555555
      5990555555555555577755555555555555555555555555555555}
    NumGlyphs = 2
  end
  object BitBtn2: TBitBtn
    Left = 80
    Top = 72
    Width = 105
    Height = 25
    Caption = '&X結束'
    TabOrder = 1
    OnClick = BitBtn2Click
    Glyph.Data = {
      76010000424D7601000000000000760000002800000020000000100000000100
      04000000000000010000130B0000130B00001000000000000000000000000000
      800000800000008080008000000080008000808000007F7F7F00BFBFBF000000
      FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00333333333333
      333333333333333333333333333333333333333FFF33FF333FFF339993370733
      999333777FF37FF377733339993000399933333777F777F77733333399970799
      93333333777F7377733333333999399933333333377737773333333333990993
      3333333333737F73333333333331013333333333333777FF3333333333910193
      333333333337773FF3333333399000993333333337377737FF33333399900099
      93333333773777377FF333399930003999333337773777F777FF339993370733
      9993337773337333777333333333333333333333333333333333333333333333
      3333333333333333333333333333333333333333333333333333}
    NumGlyphs = 2
  end
  object pgbStatus: TProgressBar
    Left = 0
    Top = 144
    Width = 277
    Height = 24
    Align = alBottom
    Min = 0
    Max = 100
    TabOrder = 2
  end
  object adoSrc: TADOConnection
    Connected = True
    ConnectionString = 
      'Provider=MSDASQL.1;Persist Security Info=False;Data Source=Trans' +
      'LangDB;Extended Properties="DSN=TransLangDB;DBQ=D:\App\TransLang' +
      '\DB\TransLang.mdb;DriverId=25;FIL=MS Access;MaxBufferSize=2048;P' +
      'ageTimeout=5;"'
    LoginPrompt = False
    Left = 24
    Top = 16
  end
  object qrySrc: TADOQuery
    Connection = adoSrc
    Parameters = <>
    SQL.Strings = (
      'select distinct sBody,sEngWords from WordsInfo')
    Left = 72
    Top = 16
  end
  object dbDest: TDatabase
    AliasName = 'GS'
    Connected = True
    DatabaseName = 'dbDest'
    LoginPrompt = False
    SessionName = 'Default'
    Left = 24
    Top = 104
  end
  object qryDest: TQuery
    CachedUpdates = True
    DatabaseName = 'dbDest'
    SQL.Strings = (
      'select distinct Caption, EngCaption from GiColumn')
    UpdateObject = usqDest
    Left = 64
    Top = 104
  end
  object usqDest: TUpdateSQL
    ModifySQL.Strings = (
      'update GiColumn'
      'set'
      '  EngCaption = :EngCaption'
      'where'
      '  TableName = :OLD_TableName and'
      '  OrderNo = :OLD_OrderNo')
    InsertSQL.Strings = (
      'insert into GiColumn'
      '  (EngCaption)'
      'values'
      '  (:EngCaption)')
    DeleteSQL.Strings = (
      'delete from GiColumn'
      'where'
      '  TableName = :OLD_TableName and'
      '  OrderNo = :OLD_OrderNo')
    Left = 128
    Top = 104
  end
  object qryCommon: TQuery
    DatabaseName = 'dbDest'
    SQL.Strings = (
      'update GiColumn'
      'set EngCaption=:EngCaption'
      'where Caption=:Caption')
    Left = 184
    Top = 104
    ParamData = <
      item
        DataType = ftString
        Name = 'EngCaption'
        ParamType = ptInput
      end
      item
        DataType = ftString
        Name = 'Caption'
        ParamType = ptInput
      end>
  end
end
