object frmSysParam: TfrmSysParam
  Left = 419
  Top = 162
  Width = 595
  Height = 591
  Caption = #31995#32113#21443#25976#35373#23450
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -10
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnActivate = FormActivate
  OnClose = FormClose
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 587
    Height = 33
    Align = alTop
    TabOrder = 0
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
    Width = 587
    Height = 531
    Align = alClient
    Caption = 'Panel2'
    TabOrder = 1
    object PageControl1: TPageControl
      Left = 1
      Top = 1
      Width = 585
      Height = 529
      ActivePage = frmSysParam_Tab_1
      Align = alClient
      Images = ImageList1
      MultiLine = True
      Style = tsButtons
      TabOrder = 0
      object frmSysParam_Tab_1: TTabSheet
        Caption = #31995#32113#21443#25976
        object Panel4: TPanel
          Left = 0
          Top = 0
          Width = 577
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
          Width = 577
          Height = 439
          Align = alClient
          BevelInner = bvLowered
          TabOrder = 1
          object frmSysParam_GB_1: TGroupBox
            Left = 7
            Top = 7
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
            object frmSysParam_Lbl_1: TStaticText
              Left = 13
              Top = 20
              Width = 28
              Height = 17
              Caption = #21517#31281
              TabOrder = 0
            end
            object frmSysParam_Lbl_2: TStaticText
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
            object frmSysParam_Lbl_4: TStaticText
              Left = 8
              Top = 100
              Width = 57
              Height = 17
              AutoSize = False
              Caption = #30064#21205#26178#38291
              TabOrder = 4
            end
            object frmSysParam_Lbl_3: TStaticText
              Left = 9
              Top = 77
              Width = 65
              Height = 17
              AutoSize = False
              Caption = #30064#21205#20154#21729
              TabOrder = 5
            end
          end
          object GroupBox1: TGroupBox
            Left = 228
            Top = 7
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
          object GroupBox4: TGroupBox
            Left = 392
            Top = 7
            Width = 97
            Height = 98
            Caption = ' Log '
            TabOrder = 2
            object frmSysParam_Chb_1: TDBCheckBox
              Left = 13
              Top = 28
              Width = 75
              Height = 13
              Caption = #25351#20196#20659#36865
              DataField = 'bCommandLog'
              DataSource = dsrParam
              TabOrder = 0
              ValueChecked = 'True'
              ValueUnchecked = 'False'
            end
            object frmSysParam_Chb_2: TDBCheckBox
              Left = 13
              Top = 59
              Width = 75
              Height = 17
              Caption = #25351#20196#22238#25033
              DataField = 'nResponseLog'
              DataSource = dsrParam
              TabOrder = 1
              ValueChecked = 'True'
              ValueUnchecked = 'False'
            end
          end
          object frmSysParam_GB_3: TGroupBox
            Left = 7
            Top = 335
            Width = 562
            Height = 96
            Caption = ' '#20854#20182' '
            TabOrder = 4
            object spbChangeDir: TSpeedButton
              Left = 367
              Top = 18
              Width = 19
              Height = 17
              Caption = '...'
              OnClick = spbChangeDirClick
            end
            object Label4: TLabel
              Left = 33
              Top = 47
              Width = 48
              Height = 13
              Caption = 'Source ID'
              FocusControl = DBEdit1
            end
            object Label5: TLabel
              Left = 154
              Top = 47
              Width = 36
              Height = 13
              Caption = 'Dest ID'
              FocusControl = DBEdit2
            end
            object Label6: TLabel
              Left = 248
              Top = 47
              Width = 49
              Height = 13
              Caption = 'Mop PPID'
              FocusControl = DBEdit3
            end
            object frmSysParam_Lbl_13: TStaticText
              Left = 16
              Top = 19
              Width = 72
              Height = 17
              AutoSize = False
              Caption = 'Log'#27284#36335#24465
              TabOrder = 0
            end
            object dbtLogPath: TDBEdit
              Left = 94
              Top = 16
              Width = 267
              Height = 21
              DataField = 'sLogPath'
              DataSource = dsrParam
              TabOrder = 1
              OnExit = dbtLogPathExit
            end
            object DBEdit1: TDBEdit
              Left = 92
              Top = 45
              Width = 33
              Height = 21
              DataField = 'nSourceID'
              DataSource = dsrParam
              TabOrder = 2
            end
            object DBEdit2: TDBEdit
              Left = 196
              Top = 45
              Width = 33
              Height = 21
              DataField = 'nDestID'
              DataSource = dsrParam
              TabOrder = 3
            end
            object DBEdit3: TDBEdit
              Left = 307
              Top = 45
              Width = 46
              Height = 21
              DataField = 'nMopPPID'
              DataSource = dsrParam
              TabOrder = 4
            end
          end
          object frmSysParam_GB_2: TGroupBox
            Left = 8
            Top = 144
            Width = 561
            Height = 177
            Caption = ' '#25351#20196#20659#36865' '
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            ParentFont = False
            TabOrder = 3
            object frmSysParam_Chb_3: TDBCheckBox
              Left = 20
              Top = 18
              Width = 185
              Height = 17
              Caption = #38283#21345#26178#26159#21542#19968#20341#35373#23450' Zip Code'
              DataField = 'bSetZipCode'
              DataSource = dsrParam
              TabOrder = 0
              ValueChecked = 'True'
              ValueUnchecked = 'False'
            end
            object StaticText9: TStaticText
              Left = 244
              Top = 18
              Width = 55
              Height = 17
              Caption = 'Credit Limit'
              TabOrder = 1
            end
            object DBEdit7: TDBEdit
              Left = 300
              Top = 15
              Width = 81
              Height = 21
              DataField = 'nCreditLimit'
              DataSource = dsrParam
              TabOrder = 2
            end
            object frmSysParam_Chb_4: TDBCheckBox
              Left = 20
              Top = 38
              Width = 289
              Height = 17
              Caption = #26159#21542#39023#31034#20659#36865'UI [ '#26371#20351#24471#25351#20196#20659#36865#36895#24230#35722#24930' ]'
              DataField = 'bShowUI'
              DataSource = dsrParam
              TabOrder = 3
              ValueChecked = 'True'
              ValueUnchecked = 'False'
            end
            object frmSysParam_Chb_5: TDBCheckBox
              Left = 20
              Top = 57
              Width = 521
              Height = 17
              Caption = #25351#23450'broadcast '#26085#26399' [ '#33509#19981#25351#23450', StartDate'#28858#26152#22825', EndDate'#28858#26126#22825'(Nagra'#24314#35696#19981#25351#23450')]'
              DataField = 'bAssignBroadcastDate'
              DataSource = dsrParam
              TabOrder = 4
              ValueChecked = 'True'
              ValueUnchecked = 'False'
            end
            object StaticText13: TStaticText
              Left = 24
              Top = 78
              Width = 181
              Height = 17
              Caption = 'Broadcast Start Date(YYYY/MM/DD)'
              TabOrder = 5
            end
            object dbtBroadcastSDate: TDBEdit
              Left = 208
              Top = 76
              Width = 73
              Height = 21
              DataField = 'sBroadcastSDate'
              DataSource = dsrParam
              TabOrder = 6
            end
            object StaticText14: TStaticText
              Left = 296
              Top = 78
              Width = 178
              Height = 17
              Caption = 'Broadcast End Date(YYYY/MM/DD)'
              TabOrder = 7
            end
            object dbtBroadcastEDate: TDBEdit
              Left = 476
              Top = 76
              Width = 73
              Height = 21
              DataField = 'sBroadcastEDate'
              DataSource = dsrParam
              TabOrder = 8
            end
            object frmSysParam_Lbl_8: TStaticText
              Left = 8
              Top = 128
              Width = 56
              Height = 17
              AutoSize = False
              Caption = #23574#23792'-'#27599
              TabOrder = 12
            end
            object dbtCmdRefreshRate1: TDBEdit
              Left = 68
              Top = 124
              Width = 35
              Height = 21
              DataField = 'nCmdRefreshRate1'
              DataSource = dsrParam
              TabOrder = 13
            end
            object frmSysParam_Lbl_9: TStaticText
              Left = 108
              Top = 128
              Width = 157
              Height = 17
              AutoSize = False
              Caption = #31186#35712#21462#19968#27425#25351#20196#36039#26009#24235
              TabOrder = 16
            end
            object frmSysParam_Lbl_6: TStaticText
              Left = 290
              Top = 102
              Width = 64
              Height = 17
              Caption = #21487#21516#26178#34389#29702
              TabOrder = 17
            end
            object dbtMaxCommandCount: TDBEdit
              Left = 372
              Top = 98
              Width = 33
              Height = 21
              DataField = 'nMaxCommandCount'
              DataSource = dsrParam
              TabOrder = 11
            end
            object frmSysParam_Lbl_7: TStaticText
              Left = 411
              Top = 102
              Width = 62
              Height = 17
              AutoSize = False
              Caption = #20491#25351#20196
              TabOrder = 18
            end
            object frmSysParam_Lbl_11: TStaticText
              Left = 8
              Top = 154
              Width = 56
              Height = 17
              AutoSize = False
              Caption = #38626#23792'-'#27599
              TabOrder = 19
            end
            object dbtCmdRefreshRate2: TDBEdit
              Left = 68
              Top = 150
              Width = 35
              Height = 21
              DataField = 'nCmdRefreshRate2'
              DataSource = dsrParam
              TabOrder = 15
            end
            object frmSysParam_Lbl_12: TStaticText
              Left = 108
              Top = 154
              Width = 161
              Height = 17
              AutoSize = False
              Caption = #31186#35712#21462#19968#27425#25351#20196#36039#26009#24235
              TabOrder = 20
            end
            object frmSysParam_Lbl_10: TStaticText
              Left = 276
              Top = 128
              Width = 88
              Height = 17
              AutoSize = False
              Caption = #25351#20196#37325#20659#27425#25976
              TabOrder = 21
            end
            object dbtResentTimes: TDBEdit
              Left = 372
              Top = 124
              Width = 33
              Height = 21
              DataField = 'nCmdResentTimes'
              DataSource = dsrParam
              TabOrder = 14
            end
            object frmSysParam_Lbl_5: TStaticText
              Left = 8
              Top = 102
              Width = 89
              Height = 17
              AutoSize = False
              Caption = #23574#23792#26178#27573#23450#32681
              TabOrder = 22
            end
            object StaticText21: TStaticText
              Left = 147
              Top = 102
              Width = 16
              Height = 17
              Caption = #21040
              TabOrder = 23
            end
            object dbtSHotTime: TDBEdit
              Left = 106
              Top = 98
              Width = 35
              Height = 21
              DataField = 'sSHotTime'
              DataSource = dsrParam
              TabOrder = 9
            end
            object dbtEHotTime: TDBEdit
              Left = 162
              Top = 98
              Width = 35
              Height = 21
              DataField = 'sEHotTime'
              DataSource = dsrParam
              TabOrder = 10
            end
          end
        end
      end
      object frmSysParam_Tab_2: TTabSheet
        Caption = #20351#29992#32773
        ImageIndex = 1
        object pnlMultiData: TPanel
          Left = 0
          Top = 104
          Width = 577
          Height = 368
          Align = alBottom
          BevelOuter = bvNone
          TabOrder = 0
          object dbgUser: TDBGrid
            Left = 0
            Top = 0
            Width = 577
            Height = 368
            Align = alClient
            Color = clInfoBk
            DataSource = dsrUser
            TabOrder = 0
            TitleFont.Charset = DEFAULT_CHARSET
            TitleFont.Color = clWindowText
            TitleFont.Height = -10
            TitleFont.Name = 'MS Sans Serif'
            TitleFont.Style = []
            Columns = <
              item
                Expanded = False
                FieldName = 'sID'
                Visible = True
              end
              item
                Expanded = False
                FieldName = 'sName'
                Width = 64
                Visible = True
              end
              item
                Expanded = False
                FieldName = 'sPassword'
                Width = 64
                Visible = True
              end>
          end
        end
        object Panel3: TPanel
          Left = 0
          Top = 0
          Width = 577
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
          Width = 577
          Height = 71
          Align = alClient
          BevelOuter = bvLowered
          TabOrder = 2
          object frmSysParam_Lbl_14: TLabel
            Left = 20
            Top = 15
            Width = 33
            Height = 13
            AutoSize = False
            Caption = #20195#30908
            FocusControl = dbtID
          end
          object frmSysParam_Lbl_15: TLabel
            Left = 20
            Top = 41
            Width = 33
            Height = 13
            AutoSize = False
            Caption = #22995#21517
            FocusControl = DBEdit4
          end
          object frmSysParam_Lbl_16: TLabel
            Left = 189
            Top = 41
            Width = 33
            Height = 13
            AutoSize = False
            Caption = #23494#30908
            FocusControl = DBEdit5
          end
          object dbtID: TDBEdit
            Left = 52
            Top = 12
            Width = 59
            Height = 21
            DataField = 'sID'
            DataSource = dsrUser
            TabOrder = 0
          end
          object DBEdit4: TDBEdit
            Left = 52
            Top = 39
            Width = 117
            Height = 21
            DataField = 'sName'
            DataSource = dsrUser
            TabOrder = 1
          end
          object DBEdit5: TDBEdit
            Left = 221
            Top = 39
            Width = 117
            Height = 21
            DataField = 'sPassword'
            DataSource = dsrUser
            TabOrder = 2
          end
        end
      end
      object TabSheet3: TTabSheet
        Caption = 'Emm Command'
        ImageIndex = 2
        object Panel5: TPanel
          Left = 0
          Top = 0
          Width = 577
          Height = 33
          Align = alTop
          TabOrder = 0
          object btnEmmEdit: TBitBtn
            Left = 224
            Top = 7
            Width = 61
            Height = 20
            Caption = #20462#25913
            TabOrder = 0
            OnClick = btnEmmEditClick
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
          object btnEmmCancel: TBitBtn
            Left = 289
            Top = 7
            Width = 61
            Height = 20
            Caption = #21462#28040
            TabOrder = 1
            OnClick = btnEmmCancelClick
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
          object btnEmmSave: TBitBtn
            Left = 354
            Top = 7
            Width = 61
            Height = 20
            Caption = #20786#23384
            TabOrder = 2
            OnClick = btnEmmSaveClick
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
        object pnlEmmParam: TPanel
          Left = 0
          Top = 33
          Width = 577
          Height = 439
          Align = alClient
          BevelInner = bvLowered
          TabOrder = 1
          object Label7: TLabel
            Left = 26
            Top = 59
            Width = 78
            Height = 13
            Caption = 'Broadcast Mode'
          end
          object Label8: TLabel
            Left = 26
            Top = 20
            Width = 74
            Height = 13
            Caption = 'Command Type'
            FocusControl = dbtEmmCommandType
          end
          object Label9: TLabel
            Left = 27
            Top = 91
            Width = 65
            Height = 13
            Caption = 'Address Type'
          end
          object dbtEmmCommandType: TDBEdit
            Left = 111
            Top = 18
            Width = 26
            Height = 21
            DataField = 'sCommandType'
            DataSource = dsrEmm
            TabOrder = 0
          end
          object DBEdit8: TDBEdit
            Left = 112
            Top = 56
            Width = 26
            Height = 21
            DataField = 'sBroadcastMode'
            DataSource = dsrEmm
            TabOrder = 1
          end
          object DBEdit13: TDBEdit
            Left = 112
            Top = 88
            Width = 26
            Height = 21
            DataField = 'sAddressType'
            DataSource = dsrEmm
            TabOrder = 2
          end
        end
      end
      object TabSheet4: TTabSheet
        Caption = 'Control Command'
        ImageIndex = 3
        object Panel6: TPanel
          Left = 0
          Top = 0
          Width = 577
          Height = 33
          Align = alTop
          TabOrder = 0
          object btnControlEdit: TBitBtn
            Left = 224
            Top = 7
            Width = 61
            Height = 20
            Caption = #20462#25913
            TabOrder = 0
            OnClick = btnControlEditClick
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
          object btnControlCancel: TBitBtn
            Left = 289
            Top = 7
            Width = 61
            Height = 20
            Caption = #21462#28040
            TabOrder = 1
            OnClick = btnControlCancelClick
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
          object btnControlSave: TBitBtn
            Left = 354
            Top = 7
            Width = 61
            Height = 20
            Caption = #20786#23384
            TabOrder = 2
            OnClick = btnControlSaveClick
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
        object pnlControlParam: TPanel
          Left = 0
          Top = 33
          Width = 577
          Height = 439
          Align = alClient
          BevelInner = bvLowered
          TabOrder = 1
          object Label10: TLabel
            Left = 26
            Top = 20
            Width = 74
            Height = 13
            Caption = 'Command Type'
            FocusControl = dbtControlCommandType
          end
          object Label14: TLabel
            Left = 26
            Top = 51
            Width = 81
            Height = 13
            Caption = 'Broadcasd Mode'
            FocusControl = dbtBroadcastMode
          end
          object Label15: TLabel
            Left = 26
            Top = 84
            Width = 65
            Height = 13
            Caption = 'Address Type'
            FocusControl = dbtAddressType
          end
          object Label16: TLabel
            Left = 156
            Top = 52
            Width = 402
            Height = 13
            Caption = 
              '(N or B. Normal -Standard broadcasting mode. Batch - Command wit' +
              'h a lower priority.)'
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clNavy
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            ParentFont = False
          end
          object Label17: TLabel
            Left = 29
            Top = 106
            Width = 251
            Height = 13
            Caption = '(EMM Addredding mode for EMM commands. U or G)'
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clNavy
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            ParentFont = False
          end
          object Label18: TLabel
            Left = 29
            Top = 122
            Width = 270
            Height = 13
            Caption = '(Unique - An ICC is addressed by its unique address (UA).'
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clNavy
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            ParentFont = False
          end
          object Label19: TLabel
            Left = 29
            Top = 138
            Width = 211
            Height = 13
            Caption = '(Global - All ICCs of the MOP are addressed.)'
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clNavy
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            ParentFont = False
          end
          object dbtControlCommandType: TDBEdit
            Left = 119
            Top = 18
            Width = 30
            Height = 21
            DataField = 'sCommandType'
            DataSource = dsrControl
            TabOrder = 0
          end
          object dbtBroadcastMode: TDBEdit
            Left = 119
            Top = 48
            Width = 30
            Height = 21
            DataField = 'sBroadcastMode'
            DataSource = dsrControl
            TabOrder = 1
          end
          object dbtAddressType: TDBEdit
            Left = 119
            Top = 79
            Width = 30
            Height = 21
            DataField = 'sAddressType'
            DataSource = dsrControl
            TabOrder = 2
          end
        end
      end
      object TabSheet5: TTabSheet
        Caption = 'PRODUCT_DEF Command'
        ImageIndex = 4
        object Panel7: TPanel
          Left = 0
          Top = 0
          Width = 577
          Height = 33
          Align = alTop
          TabOrder = 0
          object btnProductEdit: TBitBtn
            Left = 224
            Top = 7
            Width = 61
            Height = 20
            Caption = #20462#25913
            TabOrder = 0
            OnClick = btnProductEditClick
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
          object btnProductCancel: TBitBtn
            Left = 289
            Top = 7
            Width = 61
            Height = 20
            Caption = #21462#28040
            TabOrder = 1
            OnClick = btnProductCancelClick
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
          object btnProductSave: TBitBtn
            Left = 354
            Top = 7
            Width = 61
            Height = 20
            Caption = #20786#23384
            TabOrder = 2
            OnClick = btnProductSaveClick
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
        object pnlProductParam: TPanel
          Left = 0
          Top = 33
          Width = 577
          Height = 439
          Align = alClient
          BevelInner = bvLowered
          TabOrder = 1
          object Label11: TLabel
            Left = 26
            Top = 20
            Width = 74
            Height = 13
            Caption = 'Command Type'
            FocusControl = dbtProductCommandType
          end
          object dbtProductCommandType: TDBEdit
            Left = 111
            Top = 18
            Width = 26
            Height = 21
            DataField = 'sCommandType'
            DataSource = dsrProduct
            TabOrder = 0
          end
        end
      end
      object TabSheet6: TTabSheet
        Caption = 'FEEDBACK Command'
        ImageIndex = 5
        object Panel9: TPanel
          Left = 0
          Top = 0
          Width = 577
          Height = 33
          Align = alTop
          TabOrder = 0
          object btnFeedbackEdit: TBitBtn
            Left = 224
            Top = 7
            Width = 61
            Height = 20
            Caption = #20462#25913
            TabOrder = 0
            OnClick = btnFeedbackEditClick
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
          object btnFeedbackCancel: TBitBtn
            Left = 289
            Top = 7
            Width = 61
            Height = 20
            Caption = #21462#28040
            TabOrder = 1
            OnClick = btnFeedbackCancelClick
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
          object btnFeedbackSave: TBitBtn
            Left = 354
            Top = 7
            Width = 61
            Height = 20
            Caption = #20786#23384
            TabOrder = 2
            OnClick = btnFeedbackSaveClick
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
        object pnlFeedbackParam: TPanel
          Left = 0
          Top = 33
          Width = 577
          Height = 439
          Align = alClient
          BevelInner = bvLowered
          TabOrder = 1
          object Label12: TLabel
            Left = 26
            Top = 20
            Width = 74
            Height = 13
            Caption = 'Command Type'
            FocusControl = dbtFeedbackCommandType
          end
          object dbtFeedbackCommandType: TDBEdit
            Left = 111
            Top = 18
            Width = 26
            Height = 21
            DataField = 'sCommandType'
            DataSource = dsrFeedback
            TabOrder = 0
          end
        end
      end
      object TabSheet7: TTabSheet
        Caption = 'OPERATION Command'
        ImageIndex = 6
        object Panel11: TPanel
          Left = 0
          Top = 0
          Width = 577
          Height = 33
          Align = alTop
          TabOrder = 0
          object btnOperationEdit: TBitBtn
            Left = 224
            Top = 7
            Width = 61
            Height = 20
            Caption = #20462#25913
            TabOrder = 0
            OnClick = btnOperationEditClick
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
          object btnOperationCancel: TBitBtn
            Left = 289
            Top = 7
            Width = 61
            Height = 20
            Caption = #21462#28040
            TabOrder = 1
            OnClick = btnOperationCancelClick
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
          object btnOperationSave: TBitBtn
            Left = 354
            Top = 7
            Width = 61
            Height = 20
            Caption = #20786#23384
            TabOrder = 2
            OnClick = btnOperationSaveClick
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
        object pnlOperationParam: TPanel
          Left = 0
          Top = 33
          Width = 577
          Height = 439
          Align = alClient
          BevelInner = bvLowered
          TabOrder = 1
          object Label13: TLabel
            Left = 26
            Top = 20
            Width = 74
            Height = 13
            Caption = 'Command Type'
            FocusControl = dbtOperationCommandType
          end
          object dbtOperationCommandType: TDBEdit
            Left = 111
            Top = 18
            Width = 26
            Height = 21
            DataField = 'sCommandType'
            DataSource = dsrOperation
            TabOrder = 0
          end
        end
      end
      object frmSysParam_Tab_3: TTabSheet
        Caption = #32178#36335#32232#34399#35373#23450
        ImageIndex = 7
        object Panel8: TPanel
          Left = 0
          Top = 0
          Width = 577
          Height = 33
          Align = alTop
          TabOrder = 0
          object btnNetworkIDAppend: TBitBtn
            Left = 94
            Top = 7
            Width = 61
            Height = 20
            Caption = #26032#22686
            TabOrder = 0
            OnClick = btnNetworkIDAppendClick
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
          object btnNetworkIDEdit: TBitBtn
            Left = 159
            Top = 7
            Width = 61
            Height = 20
            Caption = #20462#25913
            TabOrder = 1
            OnClick = btnNetworkIDEditClick
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
          object btnNetworkIDDelete: TBitBtn
            Left = 224
            Top = 7
            Width = 61
            Height = 20
            Caption = #21034#38500
            TabOrder = 2
            OnClick = btnNetworkIDDeleteClick
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
          object btnNetworkIDCancel: TBitBtn
            Left = 289
            Top = 7
            Width = 61
            Height = 20
            Caption = #21462#28040
            TabOrder = 3
            OnClick = btnNetworkIDCancelClick
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
          object btnNetworkIDSave: TBitBtn
            Left = 354
            Top = 7
            Width = 61
            Height = 20
            Caption = #20786#23384
            TabOrder = 4
            OnClick = btnNetworkIDSaveClick
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
        object pnlSingleNetworkIdData: TPanel
          Left = 0
          Top = 33
          Width = 577
          Height = 173
          Align = alClient
          BevelOuter = bvLowered
          TabOrder = 1
          object frmSysParam_Lbl_17: TLabel
            Left = 41
            Top = 16
            Width = 57
            Height = 13
            AutoSize = False
            Caption = #20844#21496#20195#30908
            FocusControl = DBEdit9
          end
          object frmSysParam_Lbl_18: TLabel
            Left = 41
            Top = 45
            Width = 53
            Height = 13
            AutoSize = False
            Caption = #20844#21496#21517#31281
            FocusControl = DBEdit10
          end
          object Label24: TLabel
            Left = 35
            Top = 74
            Width = 58
            Height = 13
            AutoSize = False
            Caption = 'Network ID'
            FocusControl = DBEdit11
          end
          object frmSysParam_Lbl_19: TLabel
            Left = 29
            Top = 99
            Width = 67
            Height = 13
            AutoSize = False
            Caption = #24190#31186#20839#29983#25928
            FocusControl = DBEdit11
          end
          object DBEdit9: TDBEdit
            Left = 100
            Top = 13
            Width = 59
            Height = 21
            DataField = 'CompCode'
            DataSource = dsrNetworkId
            TabOrder = 0
          end
          object DBEdit10: TDBEdit
            Left = 100
            Top = 42
            Width = 349
            Height = 21
            DataField = 'CompName'
            DataSource = dsrNetworkId
            TabOrder = 1
          end
          object DBEdit11: TDBEdit
            Left = 100
            Top = 71
            Width = 117
            Height = 21
            DataField = 'NetworkID'
            DataSource = dsrNetworkId
            TabOrder = 2
          end
          object DBEdit12: TDBEdit
            Left = 100
            Top = 96
            Width = 117
            Height = 21
            DataField = 'Operation'
            DataSource = dsrNetworkId
            TabOrder = 3
          end
        end
        object pnlMultiNetworkIdData: TPanel
          Left = 0
          Top = 206
          Width = 577
          Height = 266
          Align = alBottom
          BevelOuter = bvNone
          TabOrder = 2
          object dbgNetwordID: TDBGrid
            Left = 0
            Top = 0
            Width = 577
            Height = 266
            Align = alClient
            Color = clInfoBk
            DataSource = dsrNetworkId
            TabOrder = 0
            TitleFont.Charset = DEFAULT_CHARSET
            TitleFont.Color = clWindowText
            TitleFont.Height = -10
            TitleFont.Name = 'MS Sans Serif'
            TitleFont.Style = []
            Columns = <
              item
                Expanded = False
                FieldName = 'CompCode'
                Visible = True
              end
              item
                Expanded = False
                FieldName = 'CompName'
                Visible = True
              end
              item
                Expanded = False
                FieldName = 'NetworkID'
                Width = 64
                Visible = True
              end
              item
                Expanded = False
                FieldName = 'Operation'
                Width = 64
                Visible = True
              end>
          end
        end
      end
    end
  end
  object dsrUser: TDataSource
    DataSet = dtmMain.cdsUser
    Left = 501
    Top = 473
  end
  object ImageList1: TImageList
    Left = 250
    Top = 9
    Bitmap = {
      494C010108000900040010001000FFFFFFFFFF10FFFFFFFFFFFFFFFF424D3600
      0000000000003600000028000000400000003000000001002000000000000030
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
      0000000000000000000000000000000000000000FF000000FF000000FF000000
      FF000000FF000000FF000000FF000000FF000000FF000000FF000000FF000000
      FF000000FF000000FF000000FF000000FF000000FF000000FF000000FF000000
      FF000000FF000000FF000000FF000000FF000000FF000000FF000000FF000000
      FF000000FF000000FF000000FF000000FF000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000008484000084840000C6C6C6008484
      0000C6C6C6008484000084840000848400008484000084840000848400000084
      00008484000084840000848400000000FF0084848400FFFFFF00FFFFFF00FFFF
      FF00FFFFFF00FFFFFF00C6C6C600C6C6C600C6C6C600C6C6C600C6C6C600C6C6
      C600C6C6C600FFFFFF00FFFFFF000000FF000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000084848400000000000000
      0000000000000000000000000000000000008484000084840000848400008484
      0000848400008484000084840000848400000000000084840000848400008484
      00008484000084840000FFFFFF000000FF008484840084848400FFFFFF00FFFF
      FF00FFFFFF008400000084000000840000008400000084000000C6C6C600C6C6
      C600C6C6C600C6C6C600FFFFFF000000FF000000000000000000000000000000
      0000000000000000000000000000000000000084840000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000848484000000000084848400000000000000
      0000000000000000000000000000000000008484000084840000848400008484
      000000000000FFFFFF00000000000000000000000000FFFF0000848400008484
      00008484000084840000848400000000FF008484840084848400848484008400
      0000840000008400000084000000840000008400000084000000840000008400
      0000C6C6C600C6C6C600C6C6C6000000FF000000000000000000000000000084
      8400000000000084840000000000008484000000000000848400000000000084
      8400000000000000000000000000000000000000000000000000000000008484
      8400000000008484840000000000C6C6C6000000000000000000000000000000
      0000848484000000000000000000000000008484000084840000000000000000
      0000000000000000000000000000000000000000000084840000848400008484
      00008484000084840000848400000000FF008484840084848400840000008400
      0000840000008400000000848400840000008400000084000000840000008400
      000084000000C6C6C600C6C6C6000000FF000000000000000000000000000000
      0000000000000000000000848400000000000000000000848400008484000000
      000000000000000000000000000000000000000000000000000084848400C6C6
      C600C6C6C6008484840084848400848484008484840084848400848484008484
      84000000000084848400000000000000000084840000C6C6C600000000000000
      00000000000000000000C6C6C600000000000000000084000000848400008484
      00008484000084840000848400000000FF008484840084848400840000008400
      000084000000840000000084840084000000FF00000084000000840000008400
      000084000000C6C6C600C6C6C6000000FF00C6C6C60000848400008484000084
      8400008484000084840000848400008484000084840000000000008400000084
      840000848400000000000000000000000000000000000000000000000000C6C6
      C600C6C6C600C6C6C600C6C6C600C6C6C600C6C6C600FFFFFF00C6C6C6008484
      840000000000000000000000000000000000FFFF000000000000000000000000
      0000000000000000000000000000000000000000000000000000FFFF0000FFFF
      0000FFFF0000FFFF0000FFFF00000000FF008484840084000000840000008400
      000084000000008484000084840000848400FF000000FF00000084000000FF00
      00008400000000848400C6C6C6000000FF00FFFFFF0000848400008484000084
      840000848400008484000084840000848400008484000084840000FFFF000000
      00000084840000FFFF000000000000000000000000008484840084848400C6C6
      C600C6C6C600C6C6C600C6C6C600000000008484840084848400FFFFFF008484
      840000000000848484008484840000000000FFFF0000C6C6C600000000000000
      000000000000000000000000000000000000000000000000000000000000FFFF
      0000FFFF0000FFFF0000FFFF00000000FF008484840084000000840000008400
      000084000000008484000084840000848400FF000000FF000000FF000000FF00
      00008400000000848400C6C6C6000000FF00FFFFFF0000848400008484000084
      840000848400008484000084840000FFFF0000FFFF00000000000084840000FF
      FF0000FFFF00000000000000000000000000848484000000000000000000FFFF
      FF00C6C6C600C6C6C600C6C6C600000000008484840084848400C6C6C6008484
      840000000000000000000000000000000000FFFF0000FFFF000000000000C6C6
      C600C6C6C600C6C6C600FFFF0000FFFF0000FFFF0000FFFF000000000000FFFF
      0000FFFF0000FFFF0000FFFF00000000FF008484840084000000840000008400
      0000840000000084840000848400FF000000FF000000FF000000FF0000000084
      84000084840000848400C6C6C6000000FF00FFFFFF00FFFFFF00008484000084
      8400008484000084840000FFFF0000FFFF0000FFFF0000FFFF0000FFFF0000FF
      FF000000000000848400FFFFFF000000000084848400C6C6C600C6C6C600FFFF
      FF00848484000000000000000000000000000000000000000000C6C6C6008484
      8400C6C6C600848484000000000000000000FFFF0000FFFF0000FFFF00000000
      FF000000FF00FFFF0000FFFF0000FFFF0000FFFF0000FFFF0000FFFF0000FFFF
      0000FFFF0000FFFF0000FFFF00000000FF008484840084000000840000008400
      00008400000000848400FF000000FF000000FF000000FF000000FF0000000084
      84000084840000848400C6C6C6000000FF0084000000FFFFFF00C6C6C60000FF
      FF0000000000000000000084840000FFFF0000FFFF0000848400000000000084
      84000084840000000000FFFFFF0084000000848484000000000000000000FFFF
      FF0084848400C6C6C600FFFFFF0000000000C6C6C600C6C6C600C6C6C6008484
      840000000000000000000000000000000000FFFF0000FFFF0000FFFF00000000
      0000C6C6C600FFFF0000FFFFFF00FFFFFF00FFFF0000FFFFFF00FFFFFF00FFFF
      FF00FFFFFF00FFFFFF00FFFFFF000000FF008484840084848400840000008400
      000084000000008484000084840000848400FF000000FF000000FF000000FF00
      000084000000C6C6C600FFFFFF000000FF0084000000FFFFFF00C6C6C60000FF
      FF0000FFFF0000FFFF0000000000000000000000000000848400008484000084
      840000848400FFFFFF00FFFFFF0084000000000000000000000000000000C6C6
      C6000000000084848400C6C6C60000000000C6C6C600C6C6C600C6C6C6008484
      840000000000000000000000000000000000FFFF0000FFFF0000FFFF0000FFFF
      FF00FFFF0000FFFF0000FFFF0000FFFF0000FFFF0000FFFF0000FFFFFF00FFFF
      FF00FFFFFF00FFFF0000FFFFFF000000FF008484840084848400840000008400
      000000848400008484000084840000848400FF000000FF000000FF0000000084
      840084000000FFFFFF00FFFFFF000000FF008400000084000000FFFFFF000000
      000000FFFF0000FFFF0000848400008484000084840000FFFF0000FFFF0000FF
      FF0000848400FFFFFF008400000084000000000000000000000000000000C6C6
      C600FFFFFF00000000008484840084848400C6C6C600C6C6C600C6C6C6008484
      840000000000848484000000000000000000FFFFFF00C6C6C600C6C6C600FFFF
      FF00FFFF0000FFFF0000FFFF0000FFFF0000FFFF0000FFFF0000FFFFFF00FFFF
      FF00FFFFFF00FFFF0000FFFF00000000FF008484840084848400848484008400
      00000084840000848400008484000084840000848400FF000000840000000084
      8400FFFFFF00FFFFFF00FFFFFF000000FF008400000084000000000000000000
      000000000000008484000084840000FFFF000084840000000000000000000000
      0000FFFFFF00840000008400000084000000000000000000000084848400C6C6
      C600C6C6C600C6C6C600FFFFFF00FFFFFF00FFFFFF00C6C6C600848484008484
      840000000000000000000000000000000000FFFFFF00FFFFFF00FFFFFF00FFFF
      0000FFFF0000FFFF0000FFFF0000FFFF0000FFFFFF00FFFFFF00FFFFFF00FFFF
      FF00FFFF0000FFFF0000FFFF00000000FF008484840084848400848484008484
      8400848484008400000000848400840000008400000084000000848484008484
      840084848400FFFFFF00FFFFFF000000FF000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000008400000084000000840000000000000000000000000000008484
      8400000000000000000000000000C6C6C6000000000084848400000000000000
      000000000000000000000000000000000000FFFF0000FFFF0000FFFF0000FFFF
      0000FFFF0000FFFF0000FFFF0000FFFF0000FFFFFF00FFFFFF00FFFFFF00FFFF
      0000FFFF0000FFFF0000FFFF00000000FF008484840084848400848484008484
      8400848484008484840084848400848484008484840084848400848484008484
      84008484840084848400FFFFFF000000FF000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000084000000840000000000000000000000000000000000
      0000000000000000000000000000C6C6C6000000000084848400000000000000
      000000000000000000000000000000000000FFFF0000FFFFFF00FFFF0000FFFF
      0000FFFF0000FFFF0000FFFF0000FFFF0000FFFF0000FFFF0000FFFFFF00FFFF
      0000FFFF0000FFFF0000FFFF00000000FF008484840084848400848484008484
      8400848484008484840084848400848484008484840084848400848484008484
      84008484840084848400848484000000FF000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000084848400848484008484840084848400000000000000
      0000000000000000000000000000000000000000000000840000008400000000
      0000FFFFFF00FFFFFF00FFFFFF000000000000000000FFFFFF00000000000000
      0000FFFFFF000000000000000000FFFFFF000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000FF000000FF000000FF000000
      FF000000FF000000FF000000FF000000FF000000FF000000FF000000FF000000
      FF000000FF000000FF000000FF000000FF000000FF000000FF000000FF000000
      FF000000FF000000FF000000FF000000FF000000FF000000FF000000FF000000
      FF000000FF000000FF000000FF000000FF00000000008484840000FF00000084
      0000000000000000840000000000000000000084840000000000008484000084
      8400000000000084840000FFFF00000000008484840084848400848484008484
      8400848484008484840084848400008484000084840000000000848484008484
      840000000000000000000000000000000000FF00000084000000FFFFFF000000
      000084000000FFFFFF0084000000840000008400000084000000840000008400
      00008400000084000000840000000000FF00FF00000000000000000000000000
      000000000000000000000000000000000000000000000000000000FFFF000000
      00000000000000000000FF0000000000FF00000000008484840000FF00000084
      0000000084000000FF0000008400000000000084840000848400FFFFFF0000FF
      FF0000848400C6C6C600008484000000000000000000C6C6C60000FF0000C6C6
      C600C6C6C6008400000000FFFF00000000000084840000000000840000008400
      000084000000840000000000000000000000FF000000FF000000FFFFFF00FFFF
      FF00FFFFFF00FFFFFF0084000000840000000000000000000000C6C6C600C6C6
      C600C6C6C60000000000840000000000FF00FF00000000000000000000000000
      0000000000000000000000000000000000000000000000000000008484000000
      00000000000000000000FF0000000000FF00000000008484840000FF00000084
      0000000084000000FF000084840000FFFF0000848400C6C6C600000000000000
      00000084840000FFFF000084840000FFFF0000000000C6C6C600C6C6C600C6C6
      C6008400000000FFFF0000000000008484000000000084000000C6C6C6008400
      0000FF000000840000008400000000000000FF000000FF000000FFFFFF000000
      000000000000FFFFFF0000000000C6C6C600C6C6C600C6C6C600C6C6C600C6C6
      C600C6C6C600C6C6C600840000000000FF00FF00000000000000000000000000
      000000000000000000000000000000000000C6C6C600C6C6C600FF0000000000
      00000000000000000000FF0000000000FF00000000008484840000FF00000084
      0000000084000000FF000000840000848400C6C6C600C6C6C60000848400C6C6
      C600000000000084840000FFFF00000000000000000000000000000000000000
      000000FFFF00000000000084840000000000C6C6C600FFFFFF00C6C6C6008484
      840084000000840000000000000000000000FF000000FF000000FFFFFF00FFFF
      FF000000000000000000C6C6C600C6C6C600C6C6C600C6C6C600C6C6C6000000
      000084848400C6C6C600840000000000FF00FF00000000000000000000000000
      0000000000000000000000000000C6C6C600FF00000000000000FF0000000000
      00000000000000000000FF0000000000FF000000000084848400FFFFFF000084
      0000000084000000FF0000008400840000000000000000FFFF00C6C6C600C6C6
      C60000000000C6C6C600000000000000000000000000000000000000000000FF
      FF0000000000008484000000000084848400FFFFFF00FFFFFF00848484000000
      000084000000000000000000000000000000FF000000FF000000FF000000FFFF
      FF0000000000C6C6C600C6C6C600C6C6C600C6C6C60000000000848484008484
      84008484840000000000840000000000FF00FF000000FF00000000000000FF00
      000000000000000000000000000000000000FF00000000000000FF0000000000
      000000000000FF000000FF0000000000FF000000000000840000848484000084
      0000000084000000FF00000084008400000000848400008484000000000000FF
      FF0000000000008484000000000000000000848484008484840000FFFF000000
      000000848400000000000000000000000000FFFFFF00C6C6C600C6C6C600C6C6
      C600C6C6C600000000000000000000000000FF000000FF000000FF0000000000
      0000C6C6C600C6C6C600C6C6C600C6C6C6008484840084848400848484008484
      8400C6C6C60084000000840000000000FF00FF000000FF000000FF000000FF00
      000000000000000000000000000000848400FF000000FF000000FF0000000000
      000000000000FF000000FF0000000000FF000000000084848400C6C6C6000000
      000000008400FFFFFF000000840084000000FFFF00008400000084848400C6C6
      C600000000000000000000000000000000008484840000008400848484000084
      84000000000000000000000000000000000000000000FFFFFF00FFFFFF00FFFF
      FF00FFFFFF00000000000000000000000000FF000000FF000000FF000000C6C6
      C600C6C6C600C6C6C600C6C6C600FFFFFF008484840084848400848484008484
      84000000000084000000840000000000FF00FF000000FF000000FF000000FF00
      00000000000000000000FF00000000000000FF000000FF000000FF0000000000
      000000000000FF000000FF0000000000FF000000000084848400C6C6C6000000
      000000000000848484000000840084000000FFFF000084000000000000000000
      00000000000000000000000000000000000000008400FF00FF00000084000000
      00000000000000000000000000000000000084848400FFFFFF00FFFFFF00FFFF
      FF00C6C6C600000000000000000000000000FF000000FF00000000000000C6C6
      C600C6C6C600C6C6C60084848400FFFFFF00FFFFFF0084848400848484008484
      84008400000084000000840000000000FF00FF000000FF000000FF000000FF00
      00000000000000000000FF000000FFFFFF00FF000000FF000000FF0000000000
      0000C6C6C600FF000000FF0000000000FF000000000084848400C6C6C6000000
      0000FFFFFF00C6C6C6000000000084000000FFFF000084000000000000000000
      000000000000000000000000000000000000FF00FF0000008400000000000000
      000000000000000000000000000000000000FFFFFF00FFFFFF00FFFFFF00FFFF
      FF00FFFFFF00C6C6C6000000000000000000FF000000FF00000000000000C6C6
      C600C6C6C600000000008484840084848400FFFFFF00FFFFFF00C6C6C6008400
      00008400000084000000840000000000FF00FF000000FF000000FF000000FF00
      00000000000000000000FF000000FF000000FF000000FF000000FF000000FF00
      0000C6C6C600FF000000FF0000000000FF0084848400FFFFFF0084848400FFFF
      FF0000000000C6C6C6000000000084000000FFFFFF0084000000000000000000
      0000000000000000000000000000000000000000840000000000C6C6C6000000
      00000000000000000000000000000000000084848400FFFFFF00FFFFFF00C6C6
      C600FF000000C6C6C6000000000000000000FF000000FF000000FFFFFF00C6C6
      C600000000000000000000000000000000000000000000000000FFFFFF00FFFF
      FF008400000084000000840000000000FF00FF000000FF000000FF000000FF00
      00000000000000000000FF000000FF000000FF000000FF000000FF000000FF00
      000000000000FF000000FF0000000000FF00FFFFFF0084848400FFFFFF008484
      8400848484008484840000000000FFFFFF008400000000000000000000000000
      0000000000000000000000000000000000008484840000000000C6C6C6000000
      0000000000000000000000000000000000000000000000000000C6C6C600FFFF
      FF0084848400000000000000000000000000FF000000FF000000FFFFFF00C6C6
      C6008484840084848400848484008484840084848400FF000000FFFFFF000000
      00008400000084000000840000000000FF00FF000000FF000000FF000000FF00
      00000000000000000000FF000000FF000000FF000000FF000000FF000000FF00
      000000000000FF000000FF0000000000FF0084848400FFFFFF0000000000C6C6
      C60084848400FFFFFF0084848400FFFFFF008484840000000000000000000000
      0000000000000000000000000000000000008484840000000000000000000000
      000000000000000000000000000000000000000000000000000000000000FFFF
      FF00FFFFFF00C6C6C6000000000000000000FF000000FF000000FFFFFF00C6C6
      C60084848400848484008484840000000000FF000000FF000000FFFFFF00FFFF
      FF00FFFFFF0084000000840000000000FF00FF000000FF000000FF000000FF00
      000000000000FF000000FF000000FF000000FF000000FF000000FF000000FF00
      0000FF000000FF000000FF0000000000FF000000000084848400000000008484
      840084848400C6C6C60084848400000000008484840084848400000000000000
      0000000000000000000000000000000000000000000084848400C6C6C600C6C6
      C600C6C6C6000000000000000000000000000000000000000000000000000000
      000000000000000000000000000000000000FF000000FF000000000000000000
      0000848484008484840000000000FF000000FF000000FF000000FF000000FF00
      0000FF00000084000000840000000000FF00FF000000FF000000FF000000FF00
      0000FF000000FF000000FF000000FF000000FFFFFF00FF000000FF000000FF00
      0000FF000000FF000000FF0000000000FF000000000000000000000000000000
      00000000000084848400C6C6C600FFFFFF00FFFFFF00FFFFFF00FFFFFF008484
      8400848484000000000000000000000000000000000000000000848484008484
      8400848484008484840000000000000000000000000000000000000000000000
      000000000000000000000000000000000000FF000000FF000000FF0000000000
      000000000000FF000000FF000000FF000000FF000000FF000000FF000000FF00
      0000FF000000FF000000840000000000FF00FF000000FF000000FF000000FF00
      0000FFFFFF00FF000000FF000000FF000000FF000000FF000000FF000000FF00
      0000FF000000FF000000FF0000000000FF000000000000000000000000000000
      0000000000000000000084848400848484008484840084848400848484008484
      8400000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      000000000000000000000000000000000000FF000000FF000000FF000000FF00
      0000FF000000FF000000FF000000FF000000FF000000FF000000FF000000FF00
      0000FF000000FF000000FF0000000000FF00FF000000FF000000FF000000FF00
      0000FF000000FF000000FF000000FF000000FF000000FF000000FF000000FF00
      0000FF000000FF000000FF0000000000FF00424D3E000000000000003E000000
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
    Left = 176
    Top = 8
  end
  object dsrEmm: TDataSource
    DataSet = dtmMain.cdsEmm
    Left = 469
    Top = 473
  end
  object dsrControl: TDataSource
    DataSet = dtmMain.cdsControl
    Left = 437
    Top = 508
  end
  object dsrProduct: TDataSource
    DataSet = dtmMain.cdsProduct
    Left = 405
    Top = 508
  end
  object dsrFeedback: TDataSource
    DataSet = dtmMain.cdsFeedback
    Left = 469
    Top = 508
  end
  object dsrOperation: TDataSource
    DataSet = dtmMain.cdsOperation
    Left = 437
    Top = 472
  end
  object dsrNetworkId: TDataSource
    DataSet = dtmMain.cdsNetworkID
    Left = 504
    Top = 504
  end
  object dsrParam: TDataSource
    DataSet = dtmMain.cdsParam
    Left = 408
    Top = 472
  end
end
