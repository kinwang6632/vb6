object frmSysParam: TfrmSysParam
  Left = 117
  Top = -17
  Width = 599
  Height = 572
  Caption = 'frmSysParam'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 591
    Height = 33
    Align = alTop
    Color = clMoneyGreen
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 0
    object Label7: TLabel
      Left = 14
      Top = -1
      Width = 260
      Height = 33
      Caption = 'System Parameter'
      Font.Charset = ANSI_CHARSET
      Font.Color = clWhite
      Font.Height = -24
      Font.Name = 'Arial Black'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object btnExit: TBitBtn
      Left = 511
      Top = 8
      Width = 61
      Height = 20
      Caption = #32080#26463
      TabOrder = 0
      OnClick = btnExitClick
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
  end
  object Panel2: TPanel
    Left = 0
    Top = 33
    Width = 591
    Height = 512
    Align = alClient
    Caption = 'Panel2'
    TabOrder = 1
    object PageControl1: TPageControl
      Left = 1
      Top = 1
      Width = 589
      Height = 510
      ActivePage = TabSheet1
      Align = alClient
      Images = ImageList1
      MultiLine = True
      Style = tsButtons
      TabIndex = 0
      TabOrder = 0
      object TabSheet1: TTabSheet
        Caption = #31995#32113#21443#25976
        object Panel4: TPanel
          Left = 0
          Top = 0
          Width = 581
          Height = 33
          Align = alTop
          TabOrder = 0
          object btnParamEdit: TBitBtn
            Left = 360
            Top = 7
            Width = 61
            Height = 20
            Caption = '&E'#20462#25913
            TabOrder = 0
            OnClick = btnParamEditClick
            Glyph.Data = {
              76010000424D7601000000000000760000002800000020000000100000000100
              04000000000000010000120B0000120B00001000000000000000000000000000
              800000800000008080008000000080008000808000007F7F7F00BFBFBF000000
              FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00333333000000
              000033333377777777773333330FFFFFFFF03FF3FF7FF33F3FF700300000FF0F
              00F077F777773F737737E00BFBFB0FFFFFF07773333F7F3333F7E0BFBF000FFF
              F0F077F3337773F3F737E0FBFBFBF0F00FF077F3333FF7F77F37E0BFBF00000B
              0FF077F3337777737337E0FBFBFBFBF0FFF077F33FFFFFF73337E0BF0000000F
              FFF077FF777777733FF7000BFB00B0FF00F07773FF77373377373330000B0FFF
              FFF03337777373333FF7333330B0FFFF00003333373733FF777733330B0FF00F
              0FF03333737F37737F373330B00FFFFF0F033337F77F33337F733309030FFFFF
              00333377737FFFFF773333303300000003333337337777777333}
            NumGlyphs = 2
          end
          object btnParamCancel: TBitBtn
            Left = 425
            Top = 7
            Width = 61
            Height = 20
            Caption = '&C'#21462#28040
            TabOrder = 1
            OnClick = btnParamCancelClick
            Glyph.Data = {
              76010000424D7601000000000000760000002800000020000000100000000100
              04000000000000010000120B0000120B00001000000000000000000000000000
              800000800000008080008000000080008000808000007F7F7F00BFBFBF000000
              FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00500005000555
              555557777F777555F55500000000555055557777777755F75555005500055055
              555577F5777F57555555005550055555555577FF577F5FF55555500550050055
              5555577FF77577FF555555005050110555555577F757777FF555555505099910
              555555FF75777777FF555005550999910555577F5F77777775F5500505509990
              3055577F75F77777575F55005055090B030555775755777575755555555550B0
              B03055555F555757575755550555550B0B335555755555757555555555555550
              BBB35555F55555575F555550555555550BBB55575555555575F5555555555555
              50BB555555555555575F555555555555550B5555555555555575}
            NumGlyphs = 2
          end
          object btnParamSave: TBitBtn
            Left = 490
            Top = 7
            Width = 61
            Height = 20
            Caption = '&S'#20786#23384
            TabOrder = 2
            OnClick = btnParamSaveClick
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
        end
        object pnlSysParam: TPanel
          Left = 0
          Top = 33
          Width = 581
          Height = 445
          Align = alClient
          BevelInner = bvLowered
          TabOrder = 1
          object GroupBox5: TGroupBox
            Left = 7
            Top = 232
            Width = 562
            Height = 265
            Caption = ' '#20854#20182' '
            TabOrder = 4
            object spbChangeDir: TSpeedButton
              Left = 351
              Top = 18
              Width = 19
              Height = 17
              Caption = '...'
              OnClick = spbChangeDirClick
            end
            object Label4: TLabel
              Left = 20
              Top = 45
              Width = 53
              Height = 13
              Alignment = taRightJustify
              AutoSize = False
              Caption = 'Version'
              FocusControl = dbtID
            end
            object Label5: TLabel
              Left = 344
              Top = 243
              Width = 113
              Height = 13
              Alignment = taRightJustify
              AutoSize = False
              Caption = 'FormID('#30446#21069#29992#19981#21040')'
              FocusControl = dbtID
              Visible = False
            end
            object Label6: TLabel
              Left = 402
              Top = 77
              Width = 53
              Height = 13
              Alignment = taRightJustify
              AutoSize = False
              Caption = 'ToID'
              FocusControl = dbtID
            end
            object Label8: TLabel
              Left = 186
              Top = 77
              Width = 92
              Height = 13
              Alignment = taRightJustify
              Caption = #36215#22987'Connection ID'
              FocusControl = dbtID
            end
            object Label9: TLabel
              Left = 180
              Top = 45
              Width = 37
              Height = 13
              Alignment = taRightJustify
              AutoSize = False
              Caption = 'Key'
              FocusControl = dbtID
            end
            object Label10: TLabel
              Left = 23
              Top = 109
              Width = 226
              Height = 13
              Alignment = taRightJustify
              Caption = 'Reportback '#30340#21069#32622#37327#22825#25976'(Forward Tolerance)'
              FocusControl = dbtID
            end
            object Label11: TLabel
              Left = 17
              Top = 133
              Width = 232
              Height = 13
              Alignment = taRightJustify
              Caption = 'Reportback '#30340#24460#32622#37327#22825#25976'(Backward tolerance)'
              FocusControl = dbtID
            end
            object Label12: TLabel
              Left = 4
              Top = 77
              Width = 75
              Height = 13
              Alignment = taRightJustify
              Caption = 'Currency ('#36008#24163')'
              FocusControl = dbtID
            end
            object Label13: TLabel
              Left = 16
              Top = 162
              Width = 81
              Height = 13
              Caption = 'Time zone offset '
            end
            object Label14: TLabel
              Left = 232
              Top = 162
              Width = 118
              Height = 13
              Caption = 'Country Number(16'#36914#20301')'
              FocusControl = dbtCountryNumber
            end
            object Label15: TLabel
              Left = 10
              Top = 190
              Width = 87
              Height = 13
              Caption = 'Time zone offset2 '
              FocusControl = DBEdit3
            end
            object Label16: TLabel
              Left = 160
              Top = 162
              Width = 24
              Height = 13
              Caption = #23567#26178
            end
            object Label17: TLabel
              Left = 286
              Top = 186
              Width = 64
              Height = 13
              Caption = 'Population ID'
              FocusControl = dbtCountryNumber
            end
            object StaticText5: TStaticText
              Left = 20
              Top = 19
              Width = 58
              Height = 17
              Caption = 'Log'#27284#36335#24465
              TabOrder = 13
            end
            object dbtLogPath: TDBEdit
              Left = 84
              Top = 16
              Width = 267
              Height = 21
              DataField = 'sLogPath'
              DataSource = dsrParam
              TabOrder = 0
            end
            object dbtVersion: TDBEdit
              Left = 84
              Top = 43
              Width = 77
              Height = 21
              DataField = 'nVersion'
              DataSource = dsrParam
              MaxLength = 4
              TabOrder = 1
            end
            object dbtFormID: TDBEdit
              Left = 464
              Top = 238
              Width = 81
              Height = 21
              DataField = 'nFromID'
              DataSource = dsrParam
              MaxLength = 4
              TabOrder = 12
              Visible = False
            end
            object dbtToID: TDBEdit
              Left = 468
              Top = 72
              Width = 77
              Height = 21
              DataField = 'nToID'
              DataSource = dsrParam
              MaxLength = 4
              TabOrder = 5
            end
            object rgpSecuritytype: TRadioGroup
              Left = 297
              Top = 101
              Width = 257
              Height = 51
              Caption = '  Security type  '
              Columns = 2
              Items.Strings = (
                'M (NDS Proprietary)'
                'S (Standard MD5)')
              TabOrder = 8
            end
            object dbtConnectionID: TDBEdit
              Left = 288
              Top = 72
              Width = 81
              Height = 21
              DataField = 'nConnectionID'
              DataSource = dsrParam
              TabOrder = 4
            end
            object dbtCurrency: TDBEdit
              Left = 86
              Top = 72
              Width = 73
              Height = 21
              DataField = 'nCurrency'
              DataSource = dsrParam
              TabOrder = 3
            end
            object dbtKey: TDBEdit
              Left = 238
              Top = 43
              Width = 307
              Height = 21
              Hint = #21508#32068#38291#20197#39'~'#39#28858#21312#38548','#32780'FromID'#33287'Key'#38291#20197#39','#39#28858#21312#38548
              DataField = 'sKey'
              DataSource = dsrParam
              ParentShowHint = False
              ShowHint = True
              TabOrder = 2
            end
            object dbtForwardTolerance: TDBEdit
              Left = 256
              Top = 106
              Width = 33
              Height = 21
              DataField = 'nForwardTolerance'
              DataSource = dsrParam
              TabOrder = 6
            end
            object dbtBackwardTolerance: TDBEdit
              Left = 256
              Top = 130
              Width = 33
              Height = 21
              DataField = 'nBackwardTolerance'
              DataSource = dsrParam
              TabOrder = 7
            end
            object DBEdit1: TDBEdit
              Left = 100
              Top = 157
              Width = 53
              Height = 21
              DataField = 'nTimeZoneOffset'
              DataSource = dsrParam
              TabOrder = 9
            end
            object dbtCountryNumber: TDBEdit
              Left = 352
              Top = 157
              Width = 82
              Height = 21
              DataField = 'sCountryNumber'
              DataSource = dsrParam
              TabOrder = 10
            end
            object DBEdit3: TDBEdit
              Left = 99
              Top = 187
              Width = 54
              Height = 21
              DataField = 'nTimeZoneOffset2'
              DataSource = dsrParam
              TabOrder = 11
            end
            object dbtPopulationID: TDBEdit
              Left = 352
              Top = 184
              Width = 81
              Height = 21
              DataField = 'nPopulationID'
              DataSource = dsrParam
              TabOrder = 14
            end
          end
          object GroupBox1: TGroupBox
            Left = 7
            Top = 4
            Width = 215
            Height = 126
            Caption = ' '#31995#32113#36039#35338' '
            TabOrder = 0
            object DBText1: TDBText
              Left = 74
              Top = 100
              Width = 122
              Height = 13
              DataField = 'dUptTime'
              DataSource = dsrParam
              Font.Charset = DEFAULT_CHARSET
              Font.Color = clTeal
              Font.Height = -11
              Font.Name = 'MS Sans Serif'
              Font.Style = []
              ParentFont = False
            end
            object DBText2: TDBText
              Left = 75
              Top = 77
              Width = 68
              Height = 14
              DataField = 'sUptName'
              DataSource = dsrParam
              Font.Charset = DEFAULT_CHARSET
              Font.Color = clTeal
              Font.Height = -11
              Font.Name = 'MS Sans Serif'
              Font.Style = []
              ParentFont = False
            end
            object StaticText3: TStaticText
              Left = 13
              Top = 20
              Width = 28
              Height = 17
              Caption = #21517#31281
              TabOrder = 0
            end
            object StaticText4: TStaticText
              Left = 13
              Top = 51
              Width = 28
              Height = 17
              Caption = #29256#26412
              TabOrder = 1
            end
            object dbtSysName: TDBEdit
              Left = 42
              Top = 17
              Width = 164
              Height = 21
              DataField = 'sSysName'
              DataSource = dsrParam
              TabOrder = 2
            end
            object DBEdit6: TDBEdit
              Left = 41
              Top = 49
              Width = 105
              Height = 21
              DataField = 'sSysVersion'
              DataSource = dsrParam
              TabOrder = 3
            end
            object StaticText10: TStaticText
              Left = 13
              Top = 100
              Width = 52
              Height = 17
              Caption = #30064#21205#26178#38291
              TabOrder = 4
            end
            object StaticText11: TStaticText
              Left = 13
              Top = 77
              Width = 52
              Height = 17
              Caption = #30064#21205#20154#21729
              TabOrder = 5
            end
          end
          object GroupBox2: TGroupBox
            Left = 228
            Top = 4
            Width = 150
            Height = 126
            Caption = ' SMS Gateway '
            TabOrder = 1
            object StaticText1: TStaticText
              Left = 20
              Top = 20
              Width = 17
              Height = 17
              Caption = 'IP '
              TabOrder = 4
            end
            object StaticText2: TStaticText
              Left = 10
              Top = 47
              Width = 36
              Height = 17
              Caption = 'S-Port '
              TabOrder = 5
            end
            object dbtIP: TDBEdit
              Left = 46
              Top = 18
              Width = 96
              Height = 21
              DataField = 'sServerIP'
              DataSource = dsrParam
              TabOrder = 0
            end
            object dbtSPort: TDBEdit
              Left = 45
              Top = 44
              Width = 97
              Height = 21
              DataField = 'nSPortNo'
              DataSource = dsrParam
              TabOrder = 1
            end
            object StaticText6: TStaticText
              Left = 15
              Top = 103
              Width = 63
              Height = 17
              Caption = 'Timeout('#31186') '
              TabOrder = 6
            end
            object dbtTimeout: TDBEdit
              Left = 78
              Top = 101
              Width = 65
              Height = 21
              DataField = 'nTimeOut'
              DataSource = dsrParam
              TabOrder = 3
            end
            object StaticText12: TStaticText
              Left = 10
              Top = 78
              Width = 37
              Height = 17
              Caption = 'R-Port '
              TabOrder = 7
            end
            object dbtRPort: TDBEdit
              Left = 45
              Top = 73
              Width = 97
              Height = 21
              DataField = 'nRPortNo'
              DataSource = dsrParam
              TabOrder = 2
            end
          end
          object GroupBox3: TGroupBox
            Left = 392
            Top = 7
            Width = 97
            Height = 98
            Caption = ' Log '
            TabOrder = 2
            object DBCheckBox1: TDBCheckBox
              Left = 13
              Top = 20
              Width = 75
              Height = 13
              Caption = #25351#20196#20659#36865
              DataField = 'bCommandLog'
              DataSource = dsrParam
              TabOrder = 0
              ValueChecked = 'True'
              ValueUnchecked = 'False'
            end
            object DBCheckBox4: TDBCheckBox
              Left = 13
              Top = 59
              Width = 75
              Height = 17
              Caption = #25351#20196#22238#25033
              DataField = 'bResponseLog'
              DataSource = dsrParam
              TabOrder = 1
              ValueChecked = 'True'
              ValueUnchecked = 'False'
            end
          end
          object GroupBox4: TGroupBox
            Left = 8
            Top = 129
            Width = 561
            Height = 96
            Caption = ' '#25351#20196#20659#36865' '
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            ParentFont = False
            TabOrder = 3
            object DBCheckBox5: TDBCheckBox
              Left = 20
              Top = 22
              Width = 289
              Height = 17
              Caption = #26159#21542#39023#31034#20659#36865'UI [ '#26371#20351#24471#25351#20196#20659#36865#36895#24230#35722#24930' ]'
              DataField = 'bShowUI'
              DataSource = dsrParam
              TabOrder = 0
              ValueChecked = 'True'
              ValueUnchecked = 'False'
            end
          end
        end
      end
      object TabSheet2: TTabSheet
        Caption = #20351#29992#32773
        ImageIndex = 1
        object pnlMultiData: TPanel
          Left = 0
          Top = 130
          Width = 581
          Height = 348
          Align = alBottom
          BevelOuter = bvNone
          TabOrder = 0
          object DBGrid1: TDBGrid
            Left = 0
            Top = 0
            Width = 581
            Height = 348
            Align = alClient
            Color = clInfoBk
            DataSource = dsrUser
            ImeName = #20013#25991' ('#32321#39636') - '#26032#27880#38899
            TabOrder = 0
            TitleFont.Charset = DEFAULT_CHARSET
            TitleFont.Color = clWindowText
            TitleFont.Height = -11
            TitleFont.Name = 'MS Sans Serif'
            TitleFont.Style = []
          end
        end
        object Panel3: TPanel
          Left = 0
          Top = 0
          Width = 581
          Height = 33
          Align = alTop
          TabOrder = 1
          object btnAppend: TBitBtn
            Left = 94
            Top = 7
            Width = 61
            Height = 20
            Caption = #26032#22686
            TabOrder = 0
            OnClick = btnAppendClick
            Glyph.Data = {
              76010000424D7601000000000000760000002800000020000000100000000100
              04000000000000010000130B0000130B00001000000000000000000000000000
              800000800000008080008000000080008000808000007F7F7F00BFBFBF000000
              FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00333333333333
              33333333FF33333333FF333993333333300033377F3333333777333993333333
              300033F77FFF3333377739999993333333333777777F3333333F399999933333
              33003777777333333377333993333333330033377F3333333377333993333333
              3333333773333333333F333333333333330033333333F33333773333333C3333
              330033333337FF3333773333333CC333333333FFFFF77FFF3FF33CCCCCCCCCC3
              993337777777777F77F33CCCCCCCCCC3993337777777777377333333333CC333
              333333333337733333FF3333333C333330003333333733333777333333333333
              3000333333333333377733333333333333333333333333333333}
            NumGlyphs = 2
          end
          object btnEdit: TBitBtn
            Left = 159
            Top = 7
            Width = 61
            Height = 20
            Caption = #20462#25913
            TabOrder = 1
            OnClick = btnEditClick
            Glyph.Data = {
              76010000424D7601000000000000760000002800000020000000100000000100
              04000000000000010000120B0000120B00001000000000000000000000000000
              800000800000008080008000000080008000808000007F7F7F00BFBFBF000000
              FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00333333000000
              000033333377777777773333330FFFFFFFF03FF3FF7FF33F3FF700300000FF0F
              00F077F777773F737737E00BFBFB0FFFFFF07773333F7F3333F7E0BFBF000FFF
              F0F077F3337773F3F737E0FBFBFBF0F00FF077F3333FF7F77F37E0BFBF00000B
              0FF077F3337777737337E0FBFBFBFBF0FFF077F33FFFFFF73337E0BF0000000F
              FFF077FF777777733FF7000BFB00B0FF00F07773FF77373377373330000B0FFF
              FFF03337777373333FF7333330B0FFFF00003333373733FF777733330B0FF00F
              0FF03333737F37737F373330B00FFFFF0F033337F77F33337F733309030FFFFF
              00333377737FFFFF773333303300000003333337337777777333}
            NumGlyphs = 2
          end
          object btnDelete: TBitBtn
            Left = 224
            Top = 7
            Width = 61
            Height = 20
            Caption = #21034#38500
            TabOrder = 2
            OnClick = btnDeleteClick
            Glyph.Data = {
              76010000424D7601000000000000760000002800000020000000100000000100
              04000000000000010000130B0000130B00001000000000000000000000000000
              800000800000008080008000000080008000808000007F7F7F00BFBFBF000000
              FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00333333333333
              333333333333333333FF33333333333330003333333333333777333333333333
              300033FFFFFF3333377739999993333333333777777F3333333F399999933333
              3300377777733333337733333333333333003333333333333377333333333333
              3333333333333333333F333333333333330033333F33333333773333C3333333
              330033337F3333333377333CC3333333333333F77FFFFFFF3FF33CCCCCCCCCC3
              993337777777777F77F33CCCCCCCCCC399333777777777737733333CC3333333
              333333377F33333333FF3333C333333330003333733333333777333333333333
              3000333333333333377733333333333333333333333333333333}
            NumGlyphs = 2
          end
          object btnCancel: TBitBtn
            Left = 289
            Top = 7
            Width = 61
            Height = 20
            Caption = #21462#28040
            TabOrder = 3
            OnClick = btnCancelClick
            Glyph.Data = {
              76010000424D7601000000000000760000002800000020000000100000000100
              04000000000000010000120B0000120B00001000000000000000000000000000
              800000800000008080008000000080008000808000007F7F7F00BFBFBF000000
              FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00500005000555
              555557777F777555F55500000000555055557777777755F75555005500055055
              555577F5777F57555555005550055555555577FF577F5FF55555500550050055
              5555577FF77577FF555555005050110555555577F757777FF555555505099910
              555555FF75777777FF555005550999910555577F5F77777775F5500505509990
              3055577F75F77777575F55005055090B030555775755777575755555555550B0
              B03055555F555757575755550555550B0B335555755555757555555555555550
              BBB35555F55555575F555550555555550BBB55575555555575F5555555555555
              50BB555555555555575F555555555555550B5555555555555575}
            NumGlyphs = 2
          end
          object btnSave: TBitBtn
            Left = 354
            Top = 7
            Width = 61
            Height = 20
            Caption = #20786#23384
            TabOrder = 4
            OnClick = btnSaveClick
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
        end
        object pnlSingleData: TPanel
          Left = 0
          Top = 33
          Width = 581
          Height = 97
          Align = alClient
          BevelOuter = bvLowered
          TabOrder = 2
          object Label1: TLabel
            Left = 20
            Top = 13
            Width = 33
            Height = 13
            AutoSize = False
            Caption = #20195#30908
            FocusControl = dbtID
          end
          object Label2: TLabel
            Left = 20
            Top = 39
            Width = 33
            Height = 13
            AutoSize = False
            Caption = #22995#21517
            FocusControl = DBEdit4
          end
          object Label3: TLabel
            Left = 189
            Top = 39
            Width = 33
            Height = 13
            AutoSize = False
            Caption = #23494#30908
            FocusControl = DBEdit5
          end
          object dbtID: TDBEdit
            Left = 52
            Top = 13
            Width = 59
            Height = 21
            DataField = 'sID'
            DataSource = dsrUser
            ImeName = #20013#25991' ('#32321#39636') - '#26032#27880#38899
            TabOrder = 0
          end
          object DBEdit4: TDBEdit
            Left = 52
            Top = 39
            Width = 117
            Height = 21
            DataField = 'sName'
            DataSource = dsrUser
            ImeName = #20013#25991' ('#32321#39636') - '#26032#27880#38899
            TabOrder = 1
          end
          object DBEdit5: TDBEdit
            Left = 221
            Top = 39
            Width = 117
            Height = 21
            DataField = 'sPassword'
            DataSource = dsrUser
            ImeName = #20013#25991' ('#32321#39636') - '#26032#27880#38899
            TabOrder = 2
          end
        end
      end
    end
  end
  object dsrParam: TDataSource
    DataSet = dtmMain.cdsParam
    Left = 272
    Top = 8
  end
  object dsrUser: TDataSource
    DataSet = dtmMain.cdsUser
    Left = 309
    Top = 1
  end
  object ImageList1: TImageList
    Left = 458
    Top = 1
    Bitmap = {
      494C010108000900040010001000FFFFFFFFFF10FFFFFFFFFFFFFFFF424D3600
      0000000000003600000028000000400000003000000001001800000000000024
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000FF0000FF0000FF0000FF0000
      FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF00
      00FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF
      0000FF0000FF0000FF0000FF0000FF0000FF0000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      000000000000000000000000000000000000848400848400C6C6C6848400C6C6
      C684840084840084840084840084840084840000840084840084840084840000
      00FF848484FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFC6C6C6C6C6C6C6C6C6C6C6C6
      C6C6C6C6C6C6C6C6C6FFFFFFFFFFFF0000FF0000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000848484
      0000000000000000000000000000000000008484008484008484008484008484
      00848400848400848400000000848400848400848400848400848400FFFFFF00
      00FF848484848484FFFFFFFFFFFFFFFFFF840000840000840000840000840000
      C6C6C6C6C6C6C6C6C6C6C6C6FFFFFF0000FF0000000000000000000000000000
      0000000000000000000000848400000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000848484000000848484
      0000000000000000000000000000000000008484008484008484008484000000
      00FFFFFF000000000000000000FFFF0084840084840084840084840084840000
      00FF848484848484848484840000840000840000840000840000840000840000
      840000840000C6C6C6C6C6C6C6C6C60000FF0000000000000000000084840000
      0000848400000000848400000000848400000000848400000000000000000000
      0000000000000000000000848484000000848484000000C6C6C6000000000000
      0000000000008484840000000000000000008484008484000000000000000000
      0000000000000000000000000084840084840084840084840084840084840000
      00FF848484848484840000840000840000840000008484840000840000840000
      840000840000840000C6C6C6C6C6C60000FF0000000000000000000000000000
      0000000000848400000000000000848400848400000000000000000000000000
      0000000000000000848484C6C6C6C6C6C6848484848484848484848484848484
      848484848484000000848484000000000000848400C6C6C60000000000000000
      00000000C6C6C600000000000084000084840084840084840084840084840000
      00FF848484848484840000840000840000840000008484840000FF0000840000
      840000840000840000C6C6C6C6C6C60000FFC6C6C60084840084840084840084
      8400848400848400848400848400000000840000848400848400000000000000
      0000000000000000000000C6C6C6C6C6C6C6C6C6C6C6C6C6C6C6C6C6C6FFFFFF
      C6C6C6848484000000000000000000000000FFFF000000000000000000000000
      00000000000000000000000000000000FFFF00FFFF00FFFF00FFFF00FFFF0000
      00FF848484840000840000840000840000008484008484008484FF0000FF0000
      840000FF0000840000008484C6C6C60000FFFFFFFF0084840084840084840084
      8400848400848400848400848400848400FFFF00000000848400FFFF00000000
      0000000000848484848484C6C6C6C6C6C6C6C6C6C6C6C6000000848484848484
      FFFFFF848484000000848484848484000000FFFF00C6C6C60000000000000000
      00000000000000000000000000000000000000FFFF00FFFF00FFFF00FFFF0000
      00FF848484840000840000840000840000008484008484008484FF0000FF0000
      FF0000FF0000840000008484C6C6C60000FFFFFFFF0084840084840084840084
      8400848400848400FFFF00FFFF00000000848400FFFF00FFFF00000000000000
      0000848484000000000000FFFFFFC6C6C6C6C6C6C6C6C6000000848484848484
      C6C6C6848484000000000000000000000000FFFF00FFFF00000000C6C6C6C6C6
      C6C6C6C6FFFF00FFFF00FFFF00FFFF00000000FFFF00FFFF00FFFF00FFFF0000
      00FF848484840000840000840000840000008484008484FF0000FF0000FF0000
      FF0000008484008484008484C6C6C60000FFFFFFFFFFFFFF0084840084840084
      8400848400FFFF00FFFF00FFFF00FFFF00FFFF00FFFF000000008484FFFFFF00
      0000848484C6C6C6C6C6C6FFFFFF848484000000000000000000000000000000
      C6C6C6848484C6C6C6848484000000000000FFFF00FFFF00FFFF000000FF0000
      FFFFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF0000
      00FF848484840000840000840000840000008484FF0000FF0000FF0000FF0000
      FF0000008484008484008484C6C6C60000FF840000FFFFFFC6C6C600FFFF0000
      0000000000848400FFFF00FFFF008484000000008484008484000000FFFFFF84
      0000848484000000000000FFFFFF848484C6C6C6FFFFFF000000C6C6C6C6C6C6
      C6C6C6848484000000000000000000000000FFFF00FFFF00FFFF00000000C6C6
      C6FFFF00FFFFFFFFFFFFFFFF00FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF00
      00FF848484848484840000840000840000008484008484008484FF0000FF0000
      FF0000FF0000840000C6C6C6FFFFFF0000FF840000FFFFFFC6C6C600FFFF00FF
      FF00FFFF000000000000000000008484008484008484008484FFFFFFFFFFFF84
      0000000000000000000000C6C6C6000000848484C6C6C6000000C6C6C6C6C6C6
      C6C6C6848484000000000000000000000000FFFF00FFFF00FFFF00FFFFFFFFFF
      00FFFF00FFFF00FFFF00FFFF00FFFF00FFFFFFFFFFFFFFFFFFFFFF00FFFFFF00
      00FF848484848484840000840000008484008484008484008484FF0000FF0000
      FF0000008484840000FFFFFFFFFFFF0000FF840000840000FFFFFF00000000FF
      FF00FFFF00848400848400848400FFFF00FFFF00FFFF008484FFFFFF84000084
      0000000000000000000000C6C6C6FFFFFF000000848484848484C6C6C6C6C6C6
      C6C6C6848484000000848484000000000000FFFFFFC6C6C6C6C6C6FFFFFFFFFF
      00FFFF00FFFF00FFFF00FFFF00FFFF00FFFFFFFFFFFFFFFFFFFFFF00FFFF0000
      00FF848484848484848484840000008484008484008484008484008484FF0000
      840000008484FFFFFFFFFFFFFFFFFF0000FF8400008400000000000000000000
      0000848400848400FFFF008484000000000000000000FFFFFF84000084000084
      0000000000000000848484C6C6C6C6C6C6C6C6C6FFFFFFFFFFFFFFFFFFC6C6C6
      848484848484000000000000000000000000FFFFFFFFFFFFFFFFFFFFFF00FFFF
      00FFFF00FFFF00FFFF00FFFFFFFFFFFFFFFFFFFFFFFFFFFF00FFFF00FFFF0000
      00FF848484848484848484848484848484840000008484840000840000840000
      848484848484848484FFFFFFFFFFFF0000FF0000000000000000000000000000
      0000000000000000000000000000000000000000000000000084000084000084
      0000000000000000000000848484000000000000000000C6C6C6000000848484
      000000000000000000000000000000000000FFFF00FFFF00FFFF00FFFF00FFFF
      00FFFF00FFFF00FFFF00FFFFFFFFFFFFFFFFFFFFFF00FFFF00FFFF00FFFF0000
      00FF848484848484848484848484848484848484848484848484848484848484
      848484848484848484848484FFFFFF0000FF0000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000084000084
      0000000000000000000000000000000000000000000000C6C6C6000000848484
      000000000000000000000000000000000000FFFF00FFFFFFFFFF00FFFF00FFFF
      00FFFF00FFFF00FFFF00FFFF00FFFF00FFFFFFFFFF00FFFF00FFFF00FFFF0000
      00FF848484848484848484848484848484848484848484848484848484848484
      8484848484848484848484848484840000FF0000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000848484848484848484848484
      000000000000000000000000000000000000000000008400008400000000FFFF
      FFFFFFFFFFFFFF000000000000FFFFFF000000000000FFFFFF000000000000FF
      FFFF000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000FF0000FF0000FF0000FF0000
      FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF00
      00FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF
      0000FF0000FF0000FF0000FF0000FF0000FF00000084848400FF000084000000
      0000008400000000000000848400000000848400848400000000848400FFFF00
      0000848484848484848484848484848484848484848484008484008484000000
      848484848484000000000000000000000000FF0000840000FFFFFF0000008400
      00FFFFFF84000084000084000084000084000084000084000084000084000000
      00FFFF0000000000000000000000000000000000000000000000000000000000
      00FFFF000000000000000000FF00000000FF00000084848400FF000084000000
      840000FF000084000000008484008484FFFFFF00FFFF008484C6C6C600848400
      0000000000C6C6C600FF00C6C6C6C6C6C684000000FFFF000000008484000000
      840000840000840000840000000000000000FF0000FF0000FFFFFFFFFFFFFFFF
      FFFFFFFF840000840000000000000000C6C6C6C6C6C6C6C6C600000084000000
      00FFFF0000000000000000000000000000000000000000000000000000000000
      008484000000000000000000FF00000000FF00000084848400FF000084000000
      840000FF00848400FFFF008484C6C6C600000000000000848400FFFF00848400
      FFFF000000C6C6C6C6C6C6C6C6C684000000FFFF000000008484000000840000
      C6C6C6840000FF0000840000840000000000FF0000FF0000FFFFFF0000000000
      00FFFFFF000000C6C6C6C6C6C6C6C6C6C6C6C6C6C6C6C6C6C6C6C6C684000000
      00FFFF0000000000000000000000000000000000000000000000C6C6C6C6C6C6
      FF0000000000000000000000FF00000000FF00000084848400FF000084000000
      840000FF000084008484C6C6C6C6C6C6008484C6C6C600000000848400FFFF00
      000000000000000000000000000000FFFF000000008484000000C6C6C6FFFFFF
      C6C6C6848484840000840000000000000000FF0000FF0000FFFFFFFFFFFF0000
      00000000C6C6C6C6C6C6C6C6C6C6C6C6C6C6C6000000848484C6C6C684000000
      00FFFF0000000000000000000000000000000000000000C6C6C6FF0000000000
      FF0000000000000000000000FF00000000FF000000848484FFFFFF0084000000
      840000FF00008484000000000000FFFFC6C6C6C6C6C6000000C6C6C600000000
      000000000000000000000000FFFF000000008484000000848484FFFFFFFFFFFF
      848484000000840000000000000000000000FF0000FF0000FF0000FFFFFF0000
      00C6C6C6C6C6C6C6C6C6C6C6C600000084848484848484848400000084000000
      00FFFF0000FF0000000000FF0000000000000000000000000000FF0000000000
      FF0000000000000000FF0000FF00000000FF0000000084008484840084000000
      840000FF00008484000000848400848400000000FFFF00000000848400000000
      000084848484848400FFFF000000008484000000000000000000FFFFFFC6C6C6
      C6C6C6C6C6C6C6C6C6000000000000000000FF0000FF0000FF0000000000C6C6
      C6C6C6C6C6C6C6C6C6C6848484848484848484848484C6C6C684000084000000
      00FFFF0000FF0000FF0000FF0000000000000000000000008484FF0000FF0000
      FF0000000000000000FF0000FF00000000FF000000848484C6C6C60000000000
      84FFFFFF000084840000FFFF00840000848484C6C6C600000000000000000000
      0000848484000084848484008484000000000000000000000000000000FFFFFF
      FFFFFFFFFFFFFFFFFF000000000000000000FF0000FF0000FF0000C6C6C6C6C6
      C6C6C6C6C6C6C6FFFFFF84848484848484848484848400000084000084000000
      00FFFF0000FF0000FF0000FF0000000000000000FF0000000000FF0000FF0000
      FF0000000000000000FF0000FF00000000FF000000848484C6C6C60000000000
      00848484000084840000FFFF0084000000000000000000000000000000000000
      0000000084FF00FF000084000000000000000000000000000000848484FFFFFF
      FFFFFFFFFFFFC6C6C6000000000000000000FF0000FF0000000000C6C6C6C6C6
      C6C6C6C6848484FFFFFFFFFFFF84848484848484848484000084000084000000
      00FFFF0000FF0000FF0000FF0000000000000000FF0000FFFFFFFF0000FF0000
      FF0000000000C6C6C6FF0000FF00000000FF000000848484C6C6C6000000FFFF
      FFC6C6C6000000840000FFFF0084000000000000000000000000000000000000
      0000FF00FF000084000000000000000000000000000000000000FFFFFFFFFFFF
      FFFFFFFFFFFFFFFFFFC6C6C6000000000000FF0000FF0000000000C6C6C6C6C6
      C6000000848484848484FFFFFFFFFFFFC6C6C684000084000084000084000000
      00FFFF0000FF0000FF0000FF0000000000000000FF0000FF0000FF0000FF0000
      FF0000FF0000C6C6C6FF0000FF00000000FF848484FFFFFF848484FFFFFF0000
      00C6C6C6000000840000FFFFFF84000000000000000000000000000000000000
      0000000084000000C6C6C6000000000000000000000000000000848484FFFFFF
      FFFFFFC6C6C6FF0000C6C6C6000000000000FF0000FF0000FFFFFFC6C6C60000
      00000000000000000000000000000000FFFFFFFFFFFF84000084000084000000
      00FFFF0000FF0000FF0000FF0000000000000000FF0000FF0000FF0000FF0000
      FF0000FF0000000000FF0000FF00000000FFFFFFFF848484FFFFFF8484848484
      84848484000000FFFFFF84000000000000000000000000000000000000000000
      0000848484000000C6C6C6000000000000000000000000000000000000000000
      C6C6C6FFFFFF848484000000000000000000FF0000FF0000FFFFFFC6C6C68484
      84848484848484848484848484FF0000FFFFFF00000084000084000084000000
      00FFFF0000FF0000FF0000FF0000000000000000FF0000FF0000FF0000FF0000
      FF0000FF0000000000FF0000FF00000000FF848484FFFFFF000000C6C6C68484
      84FFFFFF848484FFFFFF84848400000000000000000000000000000000000000
      0000848484000000000000000000000000000000000000000000000000000000
      000000FFFFFFFFFFFFC6C6C6000000000000FF0000FF0000FFFFFFC6C6C68484
      84848484848484000000FF0000FF0000FFFFFFFFFFFFFFFFFF84000084000000
      00FFFF0000FF0000FF0000FF0000000000FF0000FF0000FF0000FF0000FF0000
      FF0000FF0000FF0000FF0000FF00000000FF0000008484840000008484848484
      84C6C6C684848400000084848484848400000000000000000000000000000000
      0000000000848484C6C6C6C6C6C6C6C6C6000000000000000000000000000000
      000000000000000000000000000000000000FF0000FF00000000000000008484
      84848484000000FF0000FF0000FF0000FF0000FF0000FF000084000084000000
      00FFFF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FFFFFFFF0000
      FF0000FF0000FF0000FF0000FF00000000FF0000000000000000000000000000
      00848484C6C6C6FFFFFFFFFFFFFFFFFFFFFFFF84848484848400000000000000
      0000000000000000848484848484848484848484000000000000000000000000
      000000000000000000000000000000000000FF0000FF0000FF00000000000000
      00FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF000084000000
      00FFFF0000FF0000FF0000FF0000FFFFFFFF0000FF0000FF0000FF0000FF0000
      FF0000FF0000FF0000FF0000FF00000000FF0000000000000000000000000000
      0000000084848484848484848484848484848484848400000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      000000000000000000000000000000000000FF0000FF0000FF0000FF0000FF00
      00FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF000000
      00FFFF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000
      FF0000FF0000FF0000FF0000FF00000000FF424D3E000000000000003E000000
      2800000040000000300000000100010000000000800100000000000000000000
      000000000000000000000000FFFFFF0000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000FFFFFFFF00000000FF7FFE3F
      00000000E00FFC3F00000000C00FE027000000000007C003000000000003E007
      0000000000038001000000000000000100000000000000010000000000000001
      000000000000E007000000000000E003000000000800C007000000007FF0E42F
      00000000FFFCFC3F00000000FFFEFC3F80000007000000008000000300000000
      800081030000000080008201000000008000F403000000008001080300000000
      80011003000000008007000300000000803F000300000000803F000100000000
      003F000100000000003F400300000000003F7001000000008007800100000000
      F807C00100000000FC0FFE030000000000000000000000000000000000000000
      000000000000}
  end
  object dlgOpen: TOpenDialog
    Filter = 'all files|*.*'
    Left = 424
  end
end
