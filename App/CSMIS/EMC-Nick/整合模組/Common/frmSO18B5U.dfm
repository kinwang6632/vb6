object frmSO18B5: TfrmSO18B5
  Left = 208
  Top = 103
  Width = 696
  Height = 651
  Caption = 'frmSO18B5'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -13
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 16
  object StatusBar1: TStatusBar
    Left = 0
    Top = 595
    Width = 688
    Height = 29
    Panels = <>
    SimplePanel = False
  end
  object pnlQuery: TPanel
    Left = 0
    Top = 0
    Width = 688
    Height = 121
    Align = alTop
    BevelInner = bvRaised
    BevelOuter = bvLowered
    TabOrder = 1
    object rgpQryCond: TRadioGroup
      Left = 24
      Top = 8
      Width = 305
      Height = 107
      Caption = ' '#26597#35426#26781#20214' '
      ItemIndex = 0
      Items.Strings = (
        #23458#32232
        #25910#36027#38917#30446
        #20323#37329#27512#23660#32773)
      TabOrder = 0
    end
    object edtCustID: TEdit
      Left = 115
      Top = 25
      Width = 105
      Height = 24
      TabOrder = 1
    end
    object edtCItem: TEdit
      Left = 115
      Top = 54
      Width = 105
      Height = 24
      ReadOnly = True
      TabOrder = 5
    end
    object edtCm003: TEdit
      Left = 115
      Top = 84
      Width = 105
      Height = 24
      ReadOnly = True
      TabOrder = 6
    end
    object BitBtn1: TBitBtn
      Left = 225
      Top = 55
      Width = 33
      Height = 25
      Caption = '...'
      TabOrder = 3
      OnClick = BitBtn1Click
    end
    object BitBtn2: TBitBtn
      Left = 225
      Top = 84
      Width = 33
      Height = 25
      Caption = '...'
      TabOrder = 4
      OnClick = BitBtn2Click
    end
    object BitBtn3: TBitBtn
      Left = 225
      Top = 25
      Width = 75
      Height = 25
      Caption = #26597#35426
      TabOrder = 2
      OnClick = BitBtn3Click
    end
    object Memo1: TMemo
      Left = 384
      Top = 24
      Width = 273
      Height = 89
      Lines.Strings = (
        'Memo1')
      TabOrder = 7
    end
  end
  object Panel2: TPanel
    Left = 0
    Top = 121
    Width = 688
    Height = 248
    Align = alTop
    BevelOuter = bvLowered
    TabOrder = 2
    object pnlSingleData: TPanel
      Left = 1
      Top = 48
      Width = 686
      Height = 199
      Align = alBottom
      TabOrder = 0
      object StaticText1: TStaticText
        Left = 31
        Top = 16
        Width = 28
        Height = 20
        Caption = #23458#32232
        TabOrder = 4
      end
      object dbtCustID: TDBEdit
        Left = 64
        Top = 16
        Width = 121
        Height = 24
        DataField = 'CUSTID'
        DataSource = dsrSO132
        TabOrder = 0
      end
      object StaticText2: TStaticText
        Left = 205
        Top = 16
        Width = 64
        Height = 20
        Caption = #20323#37329#27512#23660#32773
        TabOrder = 5
      end
      object dbtCm0032: TDBEdit
        Left = 272
        Top = 16
        Width = 161
        Height = 24
        DataField = 'EMPNAME'
        DataSource = dsrSO132
        ReadOnly = True
        TabOrder = 3
      end
      object memCItem: TMemo
        Left = 64
        Top = 56
        Width = 561
        Height = 129
        Color = clMoneyGreen
        ReadOnly = True
        ScrollBars = ssBoth
        TabOrder = 6
      end
      object StaticText3: TStaticText
        Left = 7
        Top = 55
        Width = 52
        Height = 20
        Caption = #25910#36027#38917#30446
        TabOrder = 7
      end
      object BitBtn9: TBitBtn
        Left = 630
        Top = 56
        Width = 33
        Height = 25
        Caption = '...'
        TabOrder = 2
        OnClick = BitBtn9Click
      end
      object BitBtn4: TBitBtn
        Left = 438
        Top = 15
        Width = 33
        Height = 25
        Caption = '...'
        TabOrder = 1
        OnClick = BitBtn4Click
      end
    end
    object pnlCmd: TPanel
      Left = 1
      Top = 1
      Width = 686
      Height = 47
      Align = alClient
      TabOrder = 1
      object btnAppend: TBitBtn
        Left = 23
        Top = 10
        Width = 75
        Height = 25
        Caption = #26032#22686
        TabOrder = 0
        OnClick = btnAppendClick
      end
      object btnEdit: TBitBtn
        Left = 103
        Top = 10
        Width = 75
        Height = 25
        Caption = #20462#25913
        TabOrder = 1
        OnClick = btnEditClick
      end
      object btnDelete: TBitBtn
        Left = 183
        Top = 10
        Width = 75
        Height = 25
        Caption = #21034#38500
        TabOrder = 2
        OnClick = btnDeleteClick
      end
      object btnCancel: TBitBtn
        Left = 263
        Top = 10
        Width = 75
        Height = 25
        Caption = #21462#28040
        TabOrder = 3
        OnClick = btnCancelClick
      end
      object btnSave: TBitBtn
        Left = 343
        Top = 10
        Width = 75
        Height = 25
        Caption = #20786#23384
        TabOrder = 4
        OnClick = btnSaveClick
      end
      object btnExit: TBitBtn
        Left = 423
        Top = 10
        Width = 75
        Height = 25
        Caption = #32080#26463
        TabOrder = 5
        OnClick = btnExitClick
      end
    end
  end
  object pnlMultiData: TPanel
    Left = 0
    Top = 369
    Width = 688
    Height = 226
    Align = alClient
    BevelInner = bvLowered
    BevelOuter = bvNone
    TabOrder = 3
    object DBGrid1: TDBGrid
      Left = 1
      Top = 1
      Width = 686
      Height = 224
      Align = alClient
      Color = clInfoBk
      DataSource = dsrSO132
      ReadOnly = True
      TabOrder = 0
      TitleFont.Charset = DEFAULT_CHARSET
      TitleFont.Color = clWindowText
      TitleFont.Height = -13
      TitleFont.Name = 'MS Sans Serif'
      TitleFont.Style = []
    end
  end
  object ADOConnection1: TADOConnection
    ConnectionString = 
      'Provider=MSDAORA.1;Password=v30;User ID=v30;Data Source=mis;Pers' +
      'ist Security Info=True'
    Provider = 'MSDAORA.1'
    Left = 344
    Top = 24
  end
  object qrySO132: TADOQuery
    Connection = ADOConnection1
    BeforePost = qrySO132BeforePost
    Parameters = <>
    SQL.Strings = (
      'select * from SO132 where compcode=-1')
    Left = 392
    Top = 24
    object qrySO132COMPCODE: TIntegerField
      FieldName = 'COMPCODE'
      Visible = False
    end
    object qrySO132CUSTID: TIntegerField
      DisplayLabel = #23458#32232
      FieldName = 'CUSTID'
    end
    object qrySO132CITEMCODE: TIntegerField
      DisplayLabel = #25910#36027#38917#30446#20195#30908
      FieldName = 'CITEMCODE'
    end
    object qrySO132CITEMNAME: TStringField
      DisplayLabel = #25910#36027#38917#30446
      FieldName = 'CITEMNAME'
    end
    object qrySO132EMPNO: TStringField
      DisplayLabel = #27512#23660#32773#20195#30908
      FieldName = 'EMPNO'
      Size = 10
    end
    object qrySO132EMPNAME: TStringField
      DisplayLabel = #27512#23660#32773
      FieldName = 'EMPNAME'
    end
    object qrySO132DATACREATOR: TStringField
      DisplayLabel = #36039#26009#24314#31435#32773
      FieldName = 'DATACREATOR'
    end
    object qrySO132UPDEN: TStringField
      DisplayLabel = #30064#21205#20154#21729
      FieldName = 'UPDEN'
    end
    object qrySO132UPDTIME: TDateTimeField
      DisplayLabel = #30064#21205#26178#38291
      FieldName = 'UPDTIME'
    end
  end
  object dsrSO132: TDataSource
    DataSet = qrySO132
    OnDataChange = dsrSO132DataChange
    Left = 432
    Top = 24
  end
  object qryCm003: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    SQL.Strings = (
      'select  EMPNO,EMPNAME from CM003'
      'where STOPFLAG=0 and WORKCLASS='#39'C'#39
      '')
    Left = 392
    Top = 56
    object qryCm003EMPNO: TStringField
      DisplayLabel = #27512#23660#32773#20195#30908
      FieldName = 'EMPNO'
      Size = 10
    end
    object qryCm003EMPNAME: TStringField
      DisplayLabel = #27512#23660#32773
      FieldName = 'EMPNAME'
    end
  end
  object qryCD019: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    SQL.Strings = (
      'select CODENO,DESCRIPTION from CD019'
      'where StopFlag=0 and RefNo=2')
    Left = 392
    Top = 88
    object qryCD019CODENO: TIntegerField
      DisplayLabel = #25910#36027#38917#30446#20195#30908
      FieldName = 'CODENO'
    end
    object qryCD019DESCRIPTION: TStringField
      DisplayLabel = #25910#36027#38917#30446
      FieldName = 'DESCRIPTION'
    end
  end
  object dsrCm003: TDataSource
    DataSet = qryCm003
    Left = 432
    Top = 56
  end
  object dsrCD019: TDataSource
    DataSet = qryCD019
    Left = 432
    Top = 88
  end
  object ADOCommand1: TADOCommand
    Connection = ADOConnection1
    Parameters = <>
    Left = 344
    Top = 64
  end
  object qrySo133: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    SQL.Strings = (
      'select * from SO133 where seq <0')
    Left = 472
    Top = 24
    object qrySo133SEQ: TBCDField
      FieldName = 'SEQ'
      Precision = 10
      Size = 0
    end
    object qrySo133COMPCODE: TIntegerField
      FieldName = 'COMPCODE'
    end
    object qrySo133OLDCITEMCODE: TIntegerField
      FieldName = 'OLDCITEMCODE'
    end
    object qrySo133OLDCITEMNAME: TStringField
      FieldName = 'OLDCITEMNAME'
    end
    object qrySo133NEWCITEMCODE: TIntegerField
      FieldName = 'NEWCITEMCODE'
    end
    object qrySo133NEWCITEMNAME: TStringField
      FieldName = 'NEWCITEMNAME'
    end
    object qrySo133OLDEMPNO: TStringField
      FieldName = 'OLDEMPNO'
      Size = 10
    end
    object qrySo133OLDEMPNAME: TStringField
      FieldName = 'OLDEMPNAME'
    end
    object qrySo133NEWEMPNO: TStringField
      FieldName = 'NEWEMPNO'
      Size = 10
    end
    object qrySo133NEWEMPNAME: TStringField
      FieldName = 'NEWEMPNAME'
    end
    object qrySo133OPERATIONMODE: TStringField
      FieldName = 'OPERATIONMODE'
      FixedChar = True
      Size = 1
    end
    object qrySo133UPDEN: TStringField
      FieldName = 'UPDEN'
    end
    object qrySo133UPDTIME: TDateTimeField
      FieldName = 'UPDTIME'
    end
  end
  object qryCommon: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    Left = 472
    Top = 64
  end
end
