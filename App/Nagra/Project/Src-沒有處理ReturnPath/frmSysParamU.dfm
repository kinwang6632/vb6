object frmSysParam: TfrmSysParam
  Left = 346
  Top = 24
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
      ActivePage = frmSysParam_Tab_2
      Align = alClient
      Images = ImageList1
      MultiLine = True
      Style = tsButtons
      TabIndex = 1
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
                Visible = True
              end
              item
                Expanded = False
                FieldName = 'sPassword'
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
                Visible = True
              end
              item
                Expanded = False
                FieldName = 'Operation'
                Visible = True
              end>
          end
        end
      end
    end
  end
  object dsrUser: TDataSource
    Left = 501
    Top = 473
  end
  object ImageList1: TImageList
    Left = 250
    Top = 9
    Bitmap = {
      494C010108000900040010001000FFFFFFFFFF10FFFFFFFFFFFFFFFF424D3600
      0000000000003600000028000000400000003000000001001000000000000018
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
      000000000000000000000000000000000000007C007C007C007C007C007C007C
      007C007C007C007C007C007C007C007C007C007C007C007C007C007C007C007C
      007C007C007C007C007C007C007C007C007C0000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000001002100218631002186310021002
      10021002100210020002100210021002007C1042FF7FFF7FFF7FFF7FFF7F1863
      186318631863186318631863FF7FFF7F007C0000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000010420000000000000000000000001002100210021002100210021002
      1002000010021002100210021002FF7F007C10421042FF7FFF7FFF7F10001000
      1000100010001863186318631863FF7F007C0000000000000000000000000000
      0000004200000000000000000000000000000000000000000000000000000000
      10420000104200000000000000000000000010021002100210020000FF7F0000
      00000000FF0310021002100210021002007C1042104210421000100010001000
      10001000100010001000186318631863007C0000000000000042000000420000
      0042000000420000004200000000000000000000000000001042000010420000
      1863000000000000000010420000000000001002100200000000000000000000
      00000000100210021002100210021002007C1042104210001000100010000042
      10001000100010001000100018631863007C0000000000000000000000000042
      0000000000420042000000000000000000000000000010421863186310421042
      1042104210421042104200001042000000001002186300000000000000001863
      00000000100010021002100210021002007C1042104210001000100010000042
      10001F00100010001000100018631863007C1863004200420042004200420042
      0042004200000002004200420000000000000000000000001863186318631863
      18631863FF7F186310420000000000000000FF03000000000000000000000000
      000000000000FF03FF03FF03FF03FF03007C1042100010001000100000420042
      00421F001F0010001F00100000421863007CFF7F004200420042004200420042
      004200420042E07F00000042E07F000000000000104210421863186318631863
      000010421042FF7F10420000104210420000FF03186300000000000000000000
      0000000000000000FF03FF03FF03FF03007C1042100010001000100000420042
      00421F001F001F001F00100000421863007CFF7F004200420042004200420042
      E07FE07F00000042E07FE07F000000000000104200000000FF7F186318631863
      000010421042186310420000000000000000FF03FF030000186318631863FF03
      FF03FF03FF030000FF03FF03FF03FF03007C1042100010001000100000420042
      1F001F001F001F000042004200421863007CFF7FFF7F0042004200420042E07F
      E07FE07FE07FE07FE07F00000042FF7F0000104218631863FF7F104200000000
      000000000000186310421863104200000000FF03FF03FF03007C007CFF03FF03
      FF03FF03FF03FF03FF03FF03FF03FF03007C1042100010001000100000421F00
      1F001F001F001F000042004200421863007C1000FF7F1863E07F000000000042
      E07FE07F00420000004200420000FF7F1000104200000000FF7F10421863FF7F
      000018631863186310420000000000000000FF03FF03FF0300001863FF03FF7F
      FF7FFF03FF7FFF7FFF7FFF7FFF7FFF7F007C1042104210001000100000420042
      00421F001F001F001F0010001863FF7F007C1000FF7F1863E07FE07FE07F0000
      000000000042004200420042FF7FFF7F10000000000000001863000010421863
      000018631863186310420000000000000000FF03FF03FF03FF7FFF03FF03FF03
      FF03FF03FF03FF7FFF7FFF7FFF03FF7F007C1042104210001000004200420042
      00421F001F001F0000421000FF7FFF7F007C10001000FF7F0000E07FE07F0042
      00420042E07FE07FE07F0042FF7F100010000000000000001863FF7F00001042
      104218631863186310420000104200000000FF7F18631863FF7FFF03FF03FF03
      FF03FF03FF03FF7FFF7FFF7FFF03FF03007C1042104210421000004200420042
      004200421F0010000042FF7FFF7FFF7F007C1000100000000000000000420042
      E07F0042000000000000FF7F100010001000000000001042186318631863FF7F
      FF7FFF7F1863104210420000000000000000FF7FFF7FFF7FFF03FF03FF03FF03
      FF03FF7FFF7FFF7FFF7FFF03FF03FF03007C1042104210421042104210000042
      100010001000104210421042FF7FFF7F007C0000000000000000000000000000
      0000000000000000000000001000100010000000000000001042000000000000
      186300001042000000000000000000000000FF03FF03FF03FF03FF03FF03FF03
      FF03FF7FFF7FFF7FFF03FF03FF03FF03007C1042104210421042104210421042
      1042104210421042104210421042FF7F007C0000000000000000000000000000
      0000000000000000000000000000100010000000000000000000000000000000
      186300001042000000000000000000000000FF03FF7FFF03FF03FF03FF03FF03
      FF03FF03FF03FF7FFF03FF03FF03FF03007C1042104210421042104210421042
      10421042104210421042104210421042007C0000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000001042
      1042104210420000000000000000000000000000000200020000FF7FFF7FFF7F
      00000000FF7F00000000FF7F00000000FF7F0000000000000000000000000000
      000000000000000000000000000000000000007C007C007C007C007C007C007C
      007C007C007C007C007C007C007C007C007C007C007C007C007C007C007C007C
      007C007C007C007C007C007C007C007C007C00001042E0030002000000400000
      0000004200000042004200000042E07F00001042104210421042104210421042
      0042004200001042104200000000000000001F001000FF7F00001000FF7F1000
      10001000100010001000100010001000007C1F00000000000000000000000000
      000000000000E07F0000000000001F00007C00001042E00300020040007C0040
      000000420042FF7FE07F004218630042000000001863E003186318631000E07F
      0000004200001000100010001000000000001F001F00FF7FFF7FFF7FFF7F1000
      10000000000018631863186300001000007C1F00000000000000000000000000
      00000000000000420000000000001F00007C00001042E00300020040007C0042
      E07F00421863000000000042E07F0042E07F00001863186318631000E07F0000
      004200001000186310001F001000100000001F001F00FF7F00000000FF7F0000
      18631863186318631863186318631000007C1F00000000000000000000000000
      0000186318631F000000000000001F00007C00001042E00300020040007C0040
      0042186318630042186300000042E07F00000000000000000000E07F00000042
      00001863FF7F1863104210001000000000001F001F00FF7FFF7F000000001863
      18631863186318630000104218631000007C1F00000000000000000000000000
      18631F0000001F000000000000001F00007C00001042FF7F00020040007C0040
      10000000E07F186318630000186300000000000000000000E07F000000420000
      1042FF7FFF7F1042000010000000000000001F001F001F00FF7F000018631863
      18631863000010421042104200001000007C1F001F0000001F00000000000000
      00001F0000001F00000000001F001F00007C00000002104200020040007C0040
      1000004200420000E07F000000420000000010421042E07F0000004200000000
      0000FF7F18631863186318630000000000001F001F001F000000186318631863
      18631042104210421042186310001000007C1F001F001F001F00000000000000
      00421F001F001F00000000001F001F00007C00001042186300000040FF7F0040
      1000FF0310001042186300000000000000001042004010420042000000000000
      00000000FF7FFF7FFF7FFF7F0000000000001F001F001F001863186318631863
      FF7F1042104210421042000010001000007C1F001F001F001F00000000001F00
      00001F001F001F00000000001F001F00007C0000104218630000000010420040
      1000FF03100000000000000000000000000000401F7C00400000000000000000
      00001042FF7FFF7FFF7F18630000000000001F001F0000001863186318631042
      FF7FFF7F104210421042100010001000007C1F001F001F001F00000000001F00
      FF7F1F001F001F00000018631F001F00007C0000104218630000FF7F18630000
      1000FF0310000000000000000000000000001F7C004000000000000000000000
      0000FF7FFF7FFF7FFF7FFF7F1863000000001F001F0000001863186300001042
      1042FF7FFF7F18631000100010001000007C1F001F001F001F00000000001F00
      1F001F001F001F001F0018631F001F00007C1042FF7F1042FF7F000018630000
      1000FF7F10000000000000000000000000000040000018630000000000000000
      00001042FF7FFF7F18631F001863000000001F001F00FF7F1863000000000000
      000000000000FF7FFF7F100010001000007C1F001F001F001F00000000001F00
      1F001F001F001F001F0000001F001F00007CFF7F1042FF7F1042104210420000
      FF7F100000000000000000000000000000001042000018630000000000000000
      0000000000001863FF7F10420000000000001F001F00FF7F1863104210421042
      104210421F00FF7F0000100010001000007C1F001F001F001F00000000001F00
      1F001F001F001F001F0000001F001F00007C1042FF7F000018631042FF7F1042
      FF7F104200000000000000000000000000001042000000000000000000000000
      0000000000000000FF7FFF7F1863000000001F001F00FF7F1863104210421042
      00001F001F00FF7FFF7FFF7F10001000007C1F001F001F001F0000001F001F00
      1F001F001F001F001F001F001F001F00007C0000104200001042104218631042
      0000104210420000000000000000000000000000104218631863186300000000
      0000000000000000000000000000000000001F001F0000000000104210420000
      1F001F001F001F001F001F0010001000007C1F001F001F001F001F001F001F00
      1F00FF7F1F001F001F001F001F001F00007C0000000000000000000010421863
      FF7FFF7FFF7FFF7F104210420000000000000000000010421042104210420000
      0000000000000000000000000000000000001F001F001F00000000001F001F00
      1F001F001F001F001F001F001F001000007C1F001F001F001F00FF7F1F001F00
      1F001F001F001F001F001F001F001F00007C0000000000000000000000001042
      1042104210421042104200000000000000000000000000000000000000000000
      0000000000000000000000000000000000001F001F001F001F001F001F001F00
      1F001F001F001F001F001F001F001F00007C1F001F001F001F001F001F001F00
      1F001F001F001F001F001F001F001F00007C424D3E000000000000003E000000
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
    Left = 469
    Top = 473
  end
  object dsrControl: TDataSource
    Left = 437
    Top = 508
  end
  object dsrProduct: TDataSource
    Left = 405
    Top = 508
  end
  object dsrFeedback: TDataSource
    Left = 469
    Top = 508
  end
  object dsrOperation: TDataSource
    Left = 437
    Top = 472
  end
  object dsrNetworkId: TDataSource
    Left = 504
    Top = 504
  end
  object dsrParam: TDataSource
    Left = 408
    Top = 472
  end
end
