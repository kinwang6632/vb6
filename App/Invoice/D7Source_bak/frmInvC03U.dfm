object frmInvC03: TfrmInvC03
  Left = 383
  Top = 201
  Width = 440
  Height = 389
  Caption = #30332#31080#38283#31435#30064#24120#22577#34920
  Color = clBtnFace
  Font.Charset = ANSI_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poMainFormCenter
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  PixelsPerInch = 96
  TextHeight = 14
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 432
    Height = 32
    Align = alTop
    Color = clSkyBlue
    TabOrder = 0
    object Label6: TLabel
      Left = 10
      Top = 8
      Width = 136
      Height = 16
      Caption = #30332#31080#38283#31435#30064#24120#22577#34920
      Font.Charset = ANSI_CHARSET
      Font.Color = clBlue
      Font.Height = -16
      Font.Name = #26032#32048#26126#39636
      Font.Style = [fsBold]
      ParentFont = False
    end
  end
  object Panel2: TPanel
    Left = 0
    Top = 32
    Width = 432
    Height = 330
    Align = alClient
    TabOrder = 1
    object rgpCondition: TcxGroupBox
      Left = 19
      Top = 16
      Caption = '  '#26597#35426#26781#20214'  '
      TabOrder = 0
      Height = 151
      Width = 390
      object Label5: TLabel
        Left = 237
        Top = 85
        Width = 9
        Height = 14
        Caption = '~'
      end
      object Label7: TLabel
        Left = 238
        Top = 115
        Width = 9
        Height = 14
        Caption = '~'
      end
      object rdoBillNo: TcxRadioButton
        Left = 28
        Top = 24
        Width = 77
        Height = 17
        Caption = #25910#36027#21934#34399
        TabOrder = 0
      end
      object rdoCustId: TcxRadioButton
        Left = 28
        Top = 54
        Width = 77
        Height = 17
        Caption = #23458#32232
        TabOrder = 1
      end
      object rdoChargeDate: TcxRadioButton
        Left = 28
        Top = 85
        Width = 77
        Height = 17
        Caption = #25910#36027#26085#26399
        TabOrder = 2
      end
      object rdoLogDate: TcxRadioButton
        Left = 28
        Top = 115
        Width = 77
        Height = 17
        Caption = 'Log'#26085#26399
        TabOrder = 3
      end
      object txtBillNo: TcxTextEdit
        Left = 115
        Top = 22
        Properties.MaxLength = 15
        Style.StyleController = dtmMain.cxEditStyle
        TabOrder = 4
        OnEnter = txtBillNoEnter
        Width = 138
      end
      object txtCustId: TcxTextEdit
        Left = 115
        Top = 52
        Properties.MaxLength = 8
        Style.StyleController = dtmMain.cxEditStyle
        TabOrder = 5
        OnEnter = txtCustIdEnter
        Width = 99
      end
      object txtChageDateSt: TcxDateEdit
        Left = 115
        Top = 82
        ParentFont = False
        Style.Font.Charset = DEFAULT_CHARSET
        Style.Font.Color = clWindowText
        Style.Font.Height = -12
        Style.Font.Name = 'Tahoma'
        Style.Font.Style = []
        Style.StyleController = dtmMain.cxEditStyle
        Style.IsFontAssigned = True
        TabOrder = 6
        OnEnter = txtChageDateStEnter
        OnExit = txtChageDateStExit
        Width = 115
      end
      object txtChageDateEd: TcxDateEdit
        Left = 252
        Top = 82
        ParentFont = False
        Style.Font.Charset = DEFAULT_CHARSET
        Style.Font.Color = clWindowText
        Style.Font.Height = -12
        Style.Font.Name = 'Tahoma'
        Style.Font.Style = []
        Style.StyleController = dtmMain.cxEditStyle
        Style.IsFontAssigned = True
        TabOrder = 7
        OnEnter = txtChageDateStEnter
        Width = 115
      end
      object txtLogDateSt: TcxDateEdit
        Left = 115
        Top = 112
        ParentFont = False
        Style.Font.Charset = DEFAULT_CHARSET
        Style.Font.Color = clWindowText
        Style.Font.Height = -12
        Style.Font.Name = 'Tahoma'
        Style.Font.Style = []
        Style.StyleController = dtmMain.cxEditStyle
        Style.IsFontAssigned = True
        TabOrder = 8
        OnEnter = txtLogDateStEnter
        OnExit = txtLogDateStExit
        Width = 115
      end
      object txtLogDateEd: TcxDateEdit
        Left = 252
        Top = 112
        ParentFont = False
        Style.Font.Charset = DEFAULT_CHARSET
        Style.Font.Color = clWindowText
        Style.Font.Height = -12
        Style.Font.Name = 'Tahoma'
        Style.Font.Style = []
        Style.StyleController = dtmMain.cxEditStyle
        Style.IsFontAssigned = True
        TabOrder = 9
        OnEnter = txtLogDateStEnter
        Width = 115
      end
    end
    object btnOk: TcxButton
      Left = 121
      Top = 284
      Width = 85
      Height = 29
      Caption = #30906#23450
      Default = True
      TabOrder = 1
      OnClick = btnQueryClick
      Glyph.Data = {
        36060000424D3606000000000000360000002800000020000000100000000100
        18000000000000060000F00A0000F00A00000000000000000000FF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
        FF00FFFF00FF008B00008200038805008A00FF00FFFF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF97796797796797796797
        7967FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
        FF00FF00A300008E0013AC2715B32B06940E009300FF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF977967977967FF00FFFF00FF97
        7967977967FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
        00B10000940016AE2E17B23114AD2A13B12906920B009500FF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FF977967977967FF00FFFF00FFFF00FFFF
        00FF977967977967FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF00A900
        009A0019B2331CB63A18B33434BC4D17B03013B02A06910B009500FF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FF977967977967FF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FF977967977967FF00FFFF00FFFF00FFFF00FFFF00FFFF00FF0CA316
        1FB73E20BA421FB84011A824018E0162C77119B23213B12B06910B009500FF00
        FFFF00FFFF00FFFF00FFFF00FF977967FF00FFFF00FFFF00FFB89888977967FF
        00FFFF00FFFF00FF977967977967FF00FFFF00FFFF00FFFF00FFFF00FF28B338
        42C8651FBB4615B32C00A40000920000930061C87119B13113B22C06900B0095
        00FF00FFFF00FFFF00FFFF00FF977967FF00FFFF00FFB8988897796797796797
        7967FF00FFFF00FFFF00FF977967977967FF00FFFF00FFFF00FFFF00FF0EB924
        43C45636BC4E00AF08FF00FFFF00FF00AA0000920063CA7217B13114B12C0691
        09009B00FF00FFFF00FFFF00FF977967FF00FFFF00FF977967977967FF00FFB8
        9888977967FF00FFFF00FFFF00FF977967977967FF00FFFF00FFFF00FFFF00FF
        10D2240DBE20009700FF00FFFF00FFFF00FF00A80000930163CA7217B13114B3
        2D069009008F00FF00FFFF00FFFF00FF977967977967977967FF00FFFF00FFFF
        00FFB89888977967FF00FFFF00FFFF00FF977967977967FF00FFFF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF00AD0000940064CB7515B1
        3116B330028904FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFB89888977967FF00FFFF00FFFF00FF977967FF00FFFF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF00AB0001970366CD
        782ABC46008A02FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFB89888977967FF00FFFF00FF977967FF00FFFF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF009A000094
        01049306008400FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FF977967977967977967977967FF00FFFF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF}
      NumGlyphs = 2
    end
    object btnClose: TcxButton
      Left = 232
      Top = 284
      Width = 85
      Height = 29
      Caption = #32080#26463
      ModalResult = 2
      TabOrder = 2
      OnClick = btnExitClick
      Glyph.Data = {
        36030000424D3603000000000000360000002800000010000000100000000100
        18000000000000030000120B0000120B00000000000000000000FF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FF0000A40206B0030AAE0000A6000098FF00FFFF
        00FFFF00FFFF00FF0000A700009A0000A7FF00FFFF00FFFF00FFFF00FF0000A9
        1844F6194DF81031D20102AB0000B6FF00FFFF00FF0000B10928D7092ED70313
        B30000ACFF00FFFF00FFFF00FF0103B32451F91F52FF1D4FFF1744E8040BB000
        00B00000AC0D2EDD1142F90D3DF50B3BF0041ABC0000A5FF00FFFF00FF0000AE
        1832DB285BFF2456FF2253FF1B4BF1050DB10F30DD164AFE1344F91041F60E3E
        F60A3CF000009FFF00FFFF00FF0000BE1F37DD3A6FFF2C5EFF295AFF2657FF20
        52FC1C4FFF194AFD1646FA1445FA0F3DF2020AB10000A8FF00FFFF00FFFF00FF
        0000C8121DC83D6AFB3567FF2C5DFF2859FF2253FF1D4EFF1A4DFF123DED0002
        AC0000BAFF00FFFF00FFFF00FFFF00FFFF00FF0000CC0000B62E4EE73668FF2E
        5EFF2859FF2254FF163DEA0000A80000ABFF00FFFF00FFFF00FFFF00FFFF00FF
        FF00FFFF00FF0000BF253FDF3B6DFF3464FF2E5EFF2759FF1B46EA0001AC0000
        A9FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF0203C84B7CFF4170FF3B
        6BFF396CFF2D5EFF2558FF1336D70000B6FF00FFFF00FFFF00FFFF00FFFF00FF
        FF00FF0000D9263CDB5080FF4575FF3662FA0D14C33C6DFF2A5BFF2053FD0B1D
        C40000C0FF00FFFF00FFFF00FFFF00FFFF00FF0000CB527CFA5081FF4B7DFF0B
        13C90000BB0E15C7386AFF2456FF1A4AF20207B30000B5FF00FFFF00FFFF00FF
        FF00FF131CDD6A9CFF5788FF2B46E70000CDFF00FF0000CD0F1BCB3065FF1F51
        FF1439DD0000B1FF00FFFF00FFFF00FFFF00FF0000DE3A52E45782FB0101D0FF
        00FFFF00FFFF00FF0000CC1426D6265AFF0F2EE30103B8FF00FFFF00FFFF00FF
        FF00FFFF00FF0000CF0000C00000CEFF00FFFF00FFFF00FFFF00FF0000C40001
        B80000B5000077FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF}
    end
    object cxGroupBox3: TcxGroupBox
      Left = 21
      Top = 180
      Caption = #21015#21360#24335
      TabOrder = 3
      Height = 94
      Width = 390
      object Panel3: TPanel
        Left = 14
        Top = 18
        Width = 130
        Height = 32
        BevelOuter = bvNone
        TabOrder = 0
        object rdoReport: TcxRadioButton
          Left = 7
          Top = 8
          Width = 49
          Height = 17
          Caption = #22577#34920
          Checked = True
          TabOrder = 0
          TabStop = True
        end
        object rdoExcel: TcxRadioButton
          Left = 60
          Top = 8
          Width = 63
          Height = 17
          Caption = 'Excel '#27284
          TabOrder = 1
        end
      end
      object txtFile: TcxButtonEdit
        Left = 144
        Top = 24
        Properties.Buttons = <
          item
            Default = True
            Kind = bkEllipsis
          end>
        Properties.OnButtonClick = txtFilePropertiesButtonClick
        Style.StyleController = dtmMain.cxEditStyle
        TabOrder = 1
        Width = 223
      end
      object PBar: TcxProgressBar
        Left = 23
        Top = 60
        Style.StyleController = dtmMain.cxEditStyle
        TabOrder = 2
        Width = 344
      end
    end
  end
  object OpenDialog1: TOpenDialog
    Filter = 'Excel'#27284#26696'(*.xls)|*.xls|'#25152#26377#27284#26696'(*.*)|*.*'
    Left = 56
    Top = 256
  end
end
