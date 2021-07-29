object fmConfig: TfmConfig
  Left = 302
  Top = 215
  BorderStyle = bsDialog
  Caption = #36984#38917
  ClientHeight = 395
  ClientWidth = 644
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poMainFormCenter
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object Bevel1: TBevel
    Left = 0
    Top = 351
    Width = 644
    Height = 3
    Align = alBottom
    Shape = bsBottomLine
  end
  object Panel1: TPanel
    Left = 0
    Top = 354
    Width = 644
    Height = 41
    Align = alBottom
    BevelOuter = bvNone
    TabOrder = 0
    object btnConfirm: TcxButton
      Left = 467
      Top = 6
      Width = 81
      Height = 25
      Caption = #30906#23450
      Default = True
      ModalResult = 1
      TabOrder = 0
      LookAndFeel.Kind = lfFlat
      LookAndFeel.NativeStyle = True
    end
    object btnCancel: TcxButton
      Left = 554
      Top = 6
      Width = 81
      Height = 25
      Cancel = True
      Caption = #21462#28040
      ModalResult = 2
      TabOrder = 1
      LookAndFeel.Kind = lfFlat
      LookAndFeel.NativeStyle = True
    end
  end
  object ConfigPage: TcxPageControl
    Left = 0
    Top = 0
    Width = 644
    Height = 351
    ActivePage = cxTabSheet1
    Align = alClient
    LookAndFeel.Kind = lfOffice11
    LookAndFeel.NativeStyle = True
    Style = 9
    TabOrder = 1
    OnChange = ConfigPageChange
    ClientRectBottom = 351
    ClientRectRight = 644
    ClientRectTop = 20
    object cxTabSheet1: TcxTabSheet
      Caption = #25511#21046#21488#21443#25976'   '
      ImageIndex = 0
      object Label1: TLabel
        Left = 50
        Top = 76
        Width = 253
        Height = 13
        Caption = #25511#21046#21488#27599'                      '#31186#22519#34892#27298#26680#36039#26009#24235#38598#21312#36899#32218#12290
      end
      object Label2: TLabel
        Left = 50
        Top = 111
        Width = 413
        Height = 13
        Caption = #25511#21046#21488#20351#29992#36039#26009#24235#38598#21312#29289#20214#26178',  '#37325#35430#37782#23450#21487#29992#36039#26009#24235#38598#21312#29289#20214', '#26368#22810'                      '#27425#12290
      end
      object Label3: TLabel
        Left = 50
        Top = 148
        Width = 318
        Height = 13
        Caption = #23567#24171#25163#30340#20351#29992#32773#36039#26009#27599'                       '#31186#24478' SO026 '#21516#27493#21040' CH005'#12290
      end
      object Label4: TLabel
        Left = 50
        Top = 186
        Width = 314
        Height = 13
        Caption = #25511#21046#21488#20171#38754', '#27599'                        '#31186#26356#26032#19968#27425#20351#29992#32773#32218#19978'/'#38626#32218#29376#24907#12290
      end
      object HeaderPanel: TPanel
        Left = 10
        Top = 17
        Width = 625
        Height = 42
        BevelOuter = bvNone
        Color = 14211288
        ParentBackground = False
        TabOrder = 0
        object HeaderImage: TImage
          Left = 11
          Top = 6
          Width = 32
          Height = 32
          AutoSize = True
          Transparent = True
        end
      end
      object txtDbValidateSessionFreq: TcxSpinEdit
        Left = 105
        Top = 73
        ParentFont = False
        Properties.MaxValue = 999.000000000000000000
        Properties.MinValue = 3.000000000000000000
        Style.StyleController = StyleModule.EditorStyle
        TabOrder = 1
        Value = 30
        Width = 52
      end
      object txtDbGetPoolObjectRetryCount: TcxSpinEdit
        Left = 409
        Top = 108
        ParentFont = False
        Properties.MaxValue = 999.000000000000000000
        Properties.MinValue = 3.000000000000000000
        Style.StyleController = StyleModule.EditorStyle
        TabOrder = 2
        Value = 30
        Width = 52
      end
      object txtDbSyncUser: TcxSpinEdit
        Left = 178
        Top = 145
        ParentFont = False
        Properties.MaxValue = 999.000000000000000000
        Properties.MinValue = 3.000000000000000000
        Style.StyleController = StyleModule.EditorStyle
        TabOrder = 3
        Value = 30
        Width = 52
      end
      object txtUIGetUserListFreq: TcxSpinEdit
        Left = 138
        Top = 183
        ParentFont = False
        Properties.MaxValue = 999.000000000000000000
        Properties.MinValue = 3.000000000000000000
        Style.StyleController = StyleModule.EditorStyle
        TabOrder = 4
        Value = 10
        Width = 52
      end
    end
    object cxTabSheet2: TcxTabSheet
      Caption = #23458#25142#31471#21443#25976'   '
      ImageIndex = 1
      object lblBusyTimeReadFrquence: TLabel
        Left = 661
        Top = 79
        Width = 153
        Height = 13
        Caption = #27599'                  '#31186#35712#21462#31995#32113#21488#36039#26009
      end
      object Label5: TLabel
        Left = 50
        Top = 76
        Width = 322
        Height = 13
        Caption = #23567#24171#25163#23458#25142#31471#33287#25511#21046#21488#20043#38291#36039#26009#20132#25563#27169#24335#28858'                                  '#12290
      end
      object Label6: TLabel
        Left = 50
        Top = 111
        Width = 176
        Height = 13
        Caption = #30070#36039#26009#20132#25563#27169#24335#28858#12304#33258#21205#26356#26032#12305#26178#21063
      end
      object Label7: TLabel
        Left = 74
        Top = 141
        Width = 186
        Height = 13
        Caption = #27599'                       '#31186', '#26356#26032#30331#20837#32773#27402#38480#12290
      end
      object Label8: TLabel
        Left = 74
        Top = 175
        Width = 186
        Height = 13
        Caption = #27599'                       '#31186', '#26356#26032#20844#20296#27396#36039#26009#12290
      end
      object Label9: TLabel
        Left = 74
        Top = 210
        Width = 208
        Height = 13
        Caption = #27599'                       '#31186', '#26356#26032#20351#29992#32773#32218#19978#28165#21934#12290
      end
      object Label10: TLabel
        Left = 50
        Top = 246
        Width = 252
        Height = 13
        Caption = #30070#32178#36335#26039#32218#26178', '#23458#25142#31471#27599'                       '#31186#33258#21205#37325#36899#12290
      end
      object txtAutoRefresh: TcxComboBox
        Left = 288
        Top = 72
        Properties.DropDownListStyle = lsFixedList
        Properties.Items.Strings = (
          #33258#21205#26356#26032
          #25163#21205#26356#26032)
        Properties.OnChange = txtAutoRefreshPropertiesChange
        Style.StyleController = StyleModule.EditorStyle
        TabOrder = 0
        Text = #33258#21205#26356#26032
        Width = 81
      end
      object txtAuthorizeRefreshRate: TcxSpinEdit
        Left = 96
        Top = 138
        ParentFont = False
        Properties.MaxValue = 999.000000000000000000
        Properties.MinValue = 3.000000000000000000
        Style.StyleController = StyleModule.EditorStyle
        TabOrder = 1
        Value = 30
        Width = 52
      end
      object txtAnnRefreshRate: TcxSpinEdit
        Left = 96
        Top = 172
        ParentFont = False
        Properties.MaxValue = 999.000000000000000000
        Properties.MinValue = 3.000000000000000000
        Style.StyleController = StyleModule.EditorStyle
        TabOrder = 2
        Value = 30
        Width = 52
      end
      object txtUserRefreshRate: TcxSpinEdit
        Left = 96
        Top = 207
        ParentFont = False
        Properties.MaxValue = 999.000000000000000000
        Properties.MinValue = 3.000000000000000000
        Style.StyleController = StyleModule.EditorStyle
        TabOrder = 3
        Value = 30
        Width = 52
      end
      object txtTryReconnectRate: TcxSpinEdit
        Left = 185
        Top = 243
        ParentFont = False
        Properties.MaxValue = 999.000000000000000000
        Properties.MinValue = 3.000000000000000000
        Style.StyleController = StyleModule.EditorStyle
        TabOrder = 4
        Value = 30
        Width = 52
      end
    end
    object cxTabSheet3: TcxTabSheet
      Caption = #31995#32113#21488#36039#26009#24235'   '
      ImageIndex = 2
      object lblTest: TLabel
        Left = 247
        Top = 63
        Width = 385
        Height = 33
        AutoSize = False
        Caption = 'lblTest'
        Transparent = True
        WordWrap = True
      end
      object SoTree: TcxTreeList
        Left = 12
        Top = 102
        Width = 620
        Height = 218
        Bands = <
          item
            FixedKind = tlbfLeft
          end
          item
          end>
        BufferedPaint = True
        LookAndFeel.Kind = lfOffice11
        LookAndFeel.NativeStyle = True
        OptionsBehavior.GoToNextCellOnEnter = True
        OptionsBehavior.Sorting = False
        OptionsView.CellEndEllipsis = True
        OptionsView.Indicator = True
        OptionsView.TreeLineStyle = tllsNone
        TabOrder = 0
        object SoTreeCol1: TcxTreeListColumn
          PropertiesClassName = 'TcxCheckBoxProperties'
          Properties.ImmediatePost = True
          Properties.NullStyle = nssUnchecked
          Caption.AlignHorz = taCenter
          Caption.Text = #36984#21462
          DataBinding.ValueType = 'String'
          Width = 80
          Position.ColIndex = 0
          Position.RowIndex = 0
          Position.BandIndex = 0
        end
        object SoTreeCol2: TcxTreeListColumn
          PropertiesClassName = 'TcxTextEditProperties'
          Caption.Text = #20844#21496#20195#30908
          DataBinding.ValueType = 'String'
          Width = 70
          Position.ColIndex = 3
          Position.RowIndex = 0
          Position.BandIndex = 0
        end
        object SoTreeCol3: TcxTreeListColumn
          PropertiesClassName = 'TcxTextEditProperties'
          Caption.Text = #31995#32113#21488#21517#31281
          DataBinding.ValueType = 'String'
          Width = 100
          Position.ColIndex = 4
          Position.RowIndex = 0
          Position.BandIndex = 0
        end
        object SoTreeCol4: TcxTreeListColumn
          PropertiesClassName = 'TcxTextEditProperties'
          Caption.Text = #24115#34399
          DataBinding.ValueType = 'String'
          Width = 90
          Position.ColIndex = 3
          Position.RowIndex = 0
          Position.BandIndex = 1
        end
        object SoTreeCol5: TcxTreeListColumn
          PropertiesClassName = 'TcxTextEditProperties'
          Properties.EchoMode = eemPassword
          Caption.Text = #23494#30908
          DataBinding.ValueType = 'String'
          Width = 90
          Position.ColIndex = 5
          Position.RowIndex = 0
          Position.BandIndex = 1
        end
        object SoTreeCol6: TcxTreeListColumn
          PropertiesClassName = 'TcxTextEditProperties'
          Caption.Text = #36039#26009#24235#21029#21517
          DataBinding.ValueType = 'String'
          Width = 121
          Position.ColIndex = 2
          Position.RowIndex = 0
          Position.BandIndex = 1
        end
        object SoTreeCol7: TcxTreeListColumn
          PropertiesClassName = 'TcxCheckBoxProperties'
          Properties.ImmediatePost = True
          Properties.NullStyle = nssUnchecked
          Caption.AlignHorz = taCenter
          Caption.Text = #21516#27493' CH005'
          DataBinding.ValueType = 'String'
          Width = 80
          Position.ColIndex = 6
          Position.RowIndex = 0
          Position.BandIndex = 1
        end
        object SoTreeCol8: TcxTreeListColumn
          PropertiesClassName = 'TcxSpinEditProperties'
          Properties.ImmediatePost = True
          Properties.MaxValue = 999.000000000000000000
          Properties.MinValue = 1.000000000000000000
          Caption.AlignHorz = taCenter
          Caption.Text = #26368#22823' Pool '#25976
          DataBinding.ValueType = 'String'
          Width = 82
          Position.ColIndex = 7
          Position.RowIndex = 0
          Position.BandIndex = 1
        end
      end
      object btnDeleteSo: TcxButton
        Left = 86
        Top = 64
        Width = 67
        Height = 26
        Caption = #21034#38500
        ParentShowHint = False
        ShowHint = False
        TabOrder = 1
        OnClick = btnDeleteSoClick
        Glyph.Data = {
          36040000424D3604000000000000360000002800000010000000100000000100
          2000000000000004000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          A4140206B0CB030AAEE10000A67B000098070000000000000000000000000000
          00000000A78500009A8B0000A72A000000000000000000000000000000000000
          A98D1844F6FF194DF8FF1031D2FF0102ABD00000B62400000000000000010000
          B1930928D7FF092ED7FF0313B3FF0000AC8F0000000000000000000000000103
          B3D02451F9FF1F52FFFF1D4FFFFF1744E8FF040BB0EC0000B03A0000AC970D2E
          DDFF1142F9FF0D3DF5FF0B3BF0FF041ABCFF0000A55D00000000000000000000
          AE421832DBFF285BFFFF2456FFFF2253FFFF1B4BF1FF050DB1FB0F30DDFF164A
          FEFF1344F9FF1041F6FF0E3EF6FF0A3CF0FF00009FBF00000000000000000000
          BE511F37DDFF3A6FFFFF2C5EFFFF295AFFFF2657FFFF2052FCFF1C4FFFFF194A
          FDFF1646FAFF1445FAFF0F3DF2FF020AB1E70000A84300000000000000000000
          00010000C866121DC8F03D6AFBFF3567FFFF2C5DFFFF2859FFFF2253FFFF1D4E
          FFFF1A4DFFFF123DEDFF0002ACCB0000BA1A0000000000000000000000000000
          0000000000000000CC1C0000B6BA2E4EE7FF3668FFFF2E5EFFFF2859FFFF2254
          FFFF163DEAFF0000A8BE0000AB0B000000000000000000000000000000000000
          000000000000000000000000BF4D253FDFFF3B6DFFFF3464FFFF2E5EFFFF2759
          FFFF1B46EAFF0001ACCF0000A908000000000000000000000000000000000000
          000000000000000000000203C8C14B7CFFFF4170FFFF3B6BFFFF396CFFFF2D5E
          FFFF2558FFFF1336D7FF0000B693000000000000000000000000000000000000
          0000000000000000D92C263CDBFE5080FFFF4575FFFF3662FAFF0D14C3FE3C6D
          FFFF2A5BFFFF2053FDFF0B1DC4FD0000C0470000000000000000000000000000
          0000000000000000CB8B527CFAFF5081FFFF4B7DFFFF0B13C9EA0000BB420E15
          C7E9386AFFFF2456FFFF1A4AF2FF0207B3E30000B50F00000000000000000000
          000000000000131CDDE06A9CFFFF5788FFFF2B46E7FF0000CD5C000000000000
          CD320F1BCBF23065FFFF1F51FFFF1439DDFF0000B19C00000000000000000000
          0000000000000000DE913A52E4FE5782FBFF0101D0C200000000000000000000
          00000000CC471426D6FA265AFFFF0F2EE3FF0103B8CF00000000000000000000
          000000000000000000000000CF3C0000C0A40000CE3000000000000000000000
          0000000000000000C4630001B8BC0000B5490000770200000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000}
        LookAndFeel.Kind = lfFlat
        LookAndFeel.NativeStyle = True
      end
      object btnAddSo: TcxButton
        Left = 12
        Top = 64
        Width = 67
        Height = 26
        Caption = #26032#22686
        ParentShowHint = False
        ShowHint = False
        TabOrder = 2
        OnClick = btnAddSoClick
        Glyph.Data = {
          36040000424D3604000000000000360000002800000010000000100000000100
          2000000000000004000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000038307CE0E8E18FF0B8A15FF0A8814FF0985
          12FF007703C70000000000000000000000000000000000000000000000000000
          0000000000000000000000000000048609CE51DA7BFF3ACF69FF39CD67FF32C2
          5BFF017A04C70000000000000000000000000000000000000000000000000000
          000000000000000000000000000005870ACE5CE084FF3ED46EFF3DD36DFF34C5
          5FFF027C05C70000000000000000000000000000000000000000000000000000
          000000000000000000000000000006890CCE64E48AFF41D771FF3FD56FFF37C8
          61FF027D06C70000000000000000000000000000000000000000000000000893
          10D4089211D4079010D4078E0FD4088E10F056E282FF44DA74FF41D872FF39CB
          64FF038209EC037E07CE037C06CE027A05CE017804CE000000000000000021A7
          2DFF60E487FF4BDD79FF4ADC78FF49DC77FF4AE07AFF46DD77FF44DB75FF40D6
          70FF3BCD67FF39CB65FF36C862FF33C45EFF0A8413FF000000000000000022A9
          2FFF76F099FF5EEA8AFF5AE888FF56E684FF53E481FF4EE17DFF47DE78FF44DB
          75FF41D872FF3FD56FFF3DD36DFF39CD67FF0B8715FF000000000000000024AB
          30FF7EF39FFF68EE91FF64ED8EFF60EA8BFF59E886FF54E482FF4EE17DFF46DD
          77FF44DA74FF41D771FF3ED46EFF3ACF69FF0C8916FF000000000000000025AD
          32FF91F7ABFF8DF6A8FF8BF5A6FF89F4A5FF7AF09BFF59E886FF53E481FF4ADF
          7AFF5BE486FF67E58CFF5EE186FF53DB7CFF0F8D1AFF0000000000000000099D
          14CE0A9B14CE099913CE099713CE0B9615EF89F4A5FF60EA8BFF56E684FF45D8
          72FF078D0FEC06890DCE06860CCE05840BCE048209CE00000000000000000000
          0000000000000000000000000000099713CE8BF5A6FF64ED8EFF5AE888FF45D7
          71FF068B0DC70000000000000000000000000000000000000000000000000000
          00000000000000000000000000000A9914CE8DF6A8FF68EE91FF5EEA8AFF46D8
          72FF068E0EC70000000000000000000000000000000000000000000000000000
          00000000000000000000000000000B9A15CE91F7ABFF7FF39FFF76F099FF58DF
          7FFF078F0FC70000000000000000000000000000000000000000000000000000
          00000000000000000000000000000A9C16CE26AD33FF25AB32FF23A930FF21A6
          2EFF079110C70000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000}
        LookAndFeel.Kind = lfFlat
        LookAndFeel.NativeStyle = True
      end
      object btnTestSo: TcxButton
        Left = 159
        Top = 64
        Width = 67
        Height = 26
        Caption = #28204#35430
        ParentShowHint = False
        ShowHint = False
        TabOrder = 3
        OnClick = btnTestSoClick
        Glyph.Data = {
          36050000424D3605000000000000360400002800000010000000100000000100
          0800000000000001000000000000000000000001000000010000000000000000
          80000080000000808000800000008000800080800000C0C0C000808080000000
          FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00080808001717
          1700272727003737370047474700575757006767670077777700878787009797
          9700A7A7A700B7B7B700C7C7C700D7D7D700E7E7E700F7F7F70000004C000000
          7F000000B2000000E5001919FF004C4CFE007F7FFF00B2B2FF0000104C00001B
          7F000026B2000031E500194AFF004C72FE007F9AFF00B2C2FF0000204C000036
          7F00004CB2000062E500197BFF004C99FE007FB6FF00B2D3FF0000314C000051
          7F000072B2000093E50019ADFF004CBFFE007FD1FF00B2E3FF0000414C00006D
          7F000099B20000C4E50019DEFF004CE5FE007FECFF00B2F4FF00004C4700007F
          760000B2A50000E5D50019FFEE004CFEF2007FFFF500B2FFF900004C3600007F
          5B0000B27F0000E5A30019FFBD004CFECC007FFFDA00B2FFE900004C2600007F
          3F0000B2590000E5720019FF8C004CFEA5007FFFBF00B2FFD800004C1500007F
          240000B2330000E5410019FF5B004CFE7F007FFFA300B2FFC800004C0500007F
          090000B20C0000E5100019FF29004CFE59007FFF8800B2FFB7000A4C0000127F
          000019B2000020E500003AFF190066FE4C0091FF7F00BDFFB2001B4C00002D7F
          00003FB2000051E500006BFF19008CFE4C00ADFF7F00CDFFB2002B4C0000487F
          000065B2000083E500009CFF1900B2FE4C00C8FF7F00DEFFB2003C4C0000647F
          00008CB20000B4E50000CDFF1900D8FE4C00E3FF7F00EEFFB2004C4C00007F7F
          0000B2B20000E5E50000FFFF1900FEFE4C00FFFF7F00FFFFB2004C3C00007F64
          0000B28C0000E5B40000FFCD1900FED84C00FFE37F00FFEEB2004C2B00007F48
          0000B2660000E5830000FF9C1900FEB24C00FFC87F00FFDEB2004C1B00007F2D
          0000B23F0000E5510000FF6B1900FE8C4C00FFAD7F00FFCDB2004C0A00007F12
          0000B2190000E5200000FF3A1900FE654C00FF917F00FFBDB2004C0005007F00
          0900B2000C00E5001000FF192900FE4C5900FF7F8800FFB2B7004C0015007F00
          2400B2003200E5004100FF195B00FE4C7F00FF7FA300FFB2C8004C0026007F00
          3F00B2005900E5007200FF198C00FE4CA500FF7FBF00FFB2D8004C0036007F00
          5B00B2007F00E500A300FF19BD00FE4CCC00FF7FDA00FFB2E9004C0047007F00
          7600B200A500E500D500FF19EE00FE4CF200FF7FF500FFB2F90041004C006D00
          7F009900B200C400E500DE19FF00E54CFE00EC7FFF00F4B2FF0031004C005100
          7F007200B2009300E500AD19FF00BF4CFE00D17FFF00E3B2FF0020004C003600
          7F004C00B2006200E5007B19FF00994CFE00B67FFF00D3B2FF0010004C001B00
          7F002600B2003100E5004A19FF00724CFE009A7FFF00C2B2FF000F0F0F0F0F0F
          0F0F0F0F0F0F0F0F0F0F0F0F0F0F0F0F0F0F0F0F0F0F0F0F0F0F0F0000001419
          0F0F00000F0F0F0F0F0F0F0F0F0F0F0F0F0F4600000F0F0F0F0F0F000000140F
          0F0F074600000F0F0F0F0F0F0F0F0F0F0F0F19074600000F0F0F0F0000001419
          0F0F0F19074600000F0F0F0F0F0F0F0F0F16000000074600000F0F000000140F
          0F1907464646464600000F0F0F0F0F0F0F0F19074600000F0F0F0F0000001419
          0F0F0F19074600000F0F0F0F0F0F0F0F0F0F0F19070F4600000F0F0000001419
          0F0F0F0F19070F4600000F0F0F0F0F0F0F0F0F0F0F19070F46000F0F0F0F0F0F
          0F0F0F0F0F0F0F0F0F0F0F0F0F0F0F0F0F0F0F0F0F0F0F0F0F0F}
        LookAndFeel.Kind = lfFlat
        LookAndFeel.NativeStyle = True
      end
    end
  end
end
