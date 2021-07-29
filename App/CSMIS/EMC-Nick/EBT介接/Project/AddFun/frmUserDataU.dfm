object frmUserData: TfrmUserData
  Left = 223
  Top = 148
  Width = 696
  Height = 480
  Caption = 'frmUserData'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  WindowState = wsMaximized
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object StatusBar1: TStatusBar
    Left = 0
    Top = 434
    Width = 688
    Height = 19
    Panels = <>
    SimplePanel = False
  end
  object pnlFun: TPanel
    Left = 0
    Top = 129
    Width = 688
    Height = 41
    Align = alTop
    TabOrder = 1
    object btnAppend: TBitBtn
      Left = 31
      Top = 11
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
      Left = 96
      Top = 11
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
      Left = 161
      Top = 11
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
      Left = 226
      Top = 11
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
      Left = 291
      Top = 11
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
    Top = 170
    Width = 688
    Height = 151
    Align = alTop
    TabOrder = 2
    object dbtUserID: TDBEdit
      Left = 56
      Top = 18
      Width = 121
      Height = 21
      DataField = 'USERID'
      DataSource = dsrUserData
      TabOrder = 0
    end
    object StaticText1: TStaticText
      Left = 24
      Top = 18
      Width = 28
      Height = 17
      Caption = #24115#34399
      TabOrder = 1
    end
    object StaticText3: TStaticText
      Left = 189
      Top = 18
      Width = 28
      Height = 17
      Caption = #22995#21517
      TabOrder = 2
    end
    object DBRadioGroup1: TDBRadioGroup
      Left = 292
      Top = 104
      Width = 97
      Height = 35
      Caption = ' '#26159#21542#20572#29992' '
      Columns = 2
      DataField = 'STOPFLAG'
      DataSource = dsrUserData
      Items.Strings = (
        #26159
        #21542)
      TabOrder = 3
      Values.Strings = (
        '1'
        '0')
    end
    object dbtDesc: TDBEdit
      Left = 226
      Top = 18
      Width = 121
      Height = 21
      DataField = 'USERNAME'
      DataSource = dsrUserData
      TabOrder = 4
    end
    object StaticText2: TStaticText
      Left = 357
      Top = 16
      Width = 28
      Height = 17
      Caption = #23494#30908
      TabOrder = 5
    end
    object StaticText4: TStaticText
      Left = 381
      Top = 74
      Width = 28
      Height = 17
      Caption = #32676#32068
      TabOrder = 6
    end
    object StaticText6: TStaticText
      Left = 8
      Top = 74
      Width = 40
      Height = 17
      Caption = #20844#21496#21029
      TabOrder = 7
    end
    object StaticText7: TStaticText
      Left = 184
      Top = 74
      Width = 40
      Height = 17
      Caption = #37096#38272#21029
      TabOrder = 8
    end
    object StaticText9: TStaticText
      Left = 24
      Top = 48
      Width = 28
      Height = 17
      Caption = #38651#35441
      TabOrder = 9
    end
    object StaticText10: TStaticText
      Left = 189
      Top = 47
      Width = 28
      Height = 17
      Caption = #20998#27231
      TabOrder = 10
    end
    object StaticText11: TStaticText
      Left = 301
      Top = 48
      Width = 36
      Height = 17
      Caption = 'EMAIL'
      TabOrder = 11
    end
    object DBEdit1: TDBEdit
      Left = 389
      Top = 16
      Width = 121
      Height = 21
      DataField = 'USERPASSWD'
      DataSource = dsrUserData
      PasswordChar = '*'
      TabOrder = 12
    end
    object DBEdit2: TDBEdit
      Left = 56
      Top = 46
      Width = 121
      Height = 21
      DataField = 'TEL'
      DataSource = dsrUserData
      TabOrder = 13
    end
    object DBEdit3: TDBEdit
      Left = 227
      Top = 46
      Width = 59
      Height = 21
      DataField = 'TELEXT'
      DataSource = dsrUserData
      TabOrder = 14
    end
    object DBEdit4: TDBEdit
      Left = 349
      Top = 48
      Width = 225
      Height = 21
      DataField = 'EMAIL'
      DataSource = dsrUserData
      TabOrder = 15
    end
    object DBLookupComboBox1: TDBLookupComboBox
      Left = 56
      Top = 74
      Width = 123
      Height = 21
      DataField = 'COMPCODE'
      DataSource = dsrUserData
      KeyField = 'CODENO'
      ListField = 'DESCRIPTION'
      ListSource = dsrCompCode
      TabOrder = 16
    end
    object DBLookupComboBox2: TDBLookupComboBox
      Left = 227
      Top = 74
      Width = 145
      Height = 21
      DataField = 'DEPTCODE'
      DataSource = dsrUserData
      KeyField = 'CODENO'
      ListField = 'DESCRIPTION'
      ListSource = dsrDeptCode
      TabOrder = 17
    end
    object DBLookupComboBox3: TDBLookupComboBox
      Left = 411
      Top = 74
      Width = 163
      Height = 21
      DataField = 'USERGROUP'
      DataSource = dsrUserData
      KeyField = 'CODENO'
      ListField = 'DESCRIPTION'
      ListSource = dsrGroupCode
      TabOrder = 18
    end
    object DBRadioGroup2: TDBRadioGroup
      Left = 176
      Top = 104
      Width = 105
      Height = 33
      Caption = ' '#26159#21542#32676#32068#31649#29702#32773' '
      Columns = 2
      DataField = 'ISSUPERVISOR'
      DataSource = dsrUserData
      Items.Strings = (
        #26159
        #21542)
      TabOrder = 19
      Values.Strings = (
        '1'
        '0')
    end
    object DBRadioGroup3: TDBRadioGroup
      Left = 56
      Top = 104
      Width = 113
      Height = 33
      Caption = ' '#26159#21542#31995#32113#31649#29702#32773' '
      Columns = 2
      DataField = 'ISSYSOP'
      DataSource = dsrUserData
      Items.Strings = (
        #26159
        #21542)
      TabOrder = 20
      Values.Strings = (
        '1'
        '0')
    end
  end
  object pnlMultiData: TPanel
    Left = 0
    Top = 321
    Width = 688
    Height = 113
    Align = alClient
    TabOrder = 3
    object dbgGroupData: TDBGrid
      Left = 1
      Top = 1
      Width = 686
      Height = 111
      Align = alClient
      Color = clInfoBk
      DataSource = dsrUserData
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      ReadOnly = True
      TabOrder = 0
      TitleFont.Charset = DEFAULT_CHARSET
      TitleFont.Color = clTeal
      TitleFont.Height = -11
      TitleFont.Name = 'MS Sans Serif'
      TitleFont.Style = []
      OnTitleClick = dbgGroupDataTitleClick
      Columns = <
        item
          Expanded = False
          FieldName = 'USERID'
          Title.Caption = #24115#34399
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'USERNAME'
          Title.Caption = #22995#21517
          Width = 70
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'USERPASSWD'
          Title.Caption = #23494#30908
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'GROUPCODEDESC'
          Title.Caption = #32676#32068
          Width = 64
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'ISSUPERVISORDESC'
          Title.Caption = #26159#21542#32676#32068#31649#29702#32773
          Width = 64
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'ISSYSOPDESC'
          Title.Caption = #26159#21542#31995#32113#31649#29702#32773
          Width = 64
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'COMPCODEDESC'
          Title.Caption = #20844#21496#21029
          Width = 64
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'DEPTCODEDESC'
          Title.Caption = #37096#38272#21029
          Width = 64
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'TEL'
          Title.Caption = #38651#35441
          Width = 64
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'TELEXT'
          Title.Caption = #20998#27231
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'EMAIL'
          Width = 64
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'STOPFLAGDESC'
          Title.Caption = #26159#21542#20572#29992
          Width = 64
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'CREATER'
          Title.Caption = #36039#26009#24314#31435#32773
          Width = 64
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'CREATEDATE'
          Title.Caption = #36039#26009#24314#31435#26178#38291
          Width = 64
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'UPDEN'
          Title.Caption = #30064#21205#20154#21729
          Width = 64
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'UPDTIME'
          Title.Caption = #30064#21205#26178#38291
          Width = 64
          Visible = True
        end>
    end
  end
  object pnlQuery: TPanel
    Left = 0
    Top = 0
    Width = 688
    Height = 129
    Align = alTop
    TabOrder = 4
    object BitBtn1: TBitBtn
      Left = 600
      Top = 16
      Width = 75
      Height = 25
      Caption = #26597#35426
      TabOrder = 0
      OnClick = BitBtn1Click
      Glyph.Data = {
        76010000424D7601000000000000760000002800000020000000100000000100
        04000000000000010000130B0000130B00001000000000000000000000000000
        800000800000008080008000000080008000808000007F7F7F00BFBFBF000000
        FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00333333333333
        333333333333333FF3FF3333333333CC30003333333333773777333333333C33
        3000333FF33337F33777339933333C3333333377F33337F3333F339933333C33
        33003377333337F33377333333333C333300333F333337F33377339333333C33
        3333337FF3333733333F33993333C33333003377FF33733333773339933C3333
        330033377FF73F33337733339933C33333333FF377F373F3333F993399333C33
        330077F377F337F33377993399333C33330077FF773337F33377399993333C33
        33333777733337F333FF333333333C33300033333333373FF7773333333333CC
        3000333333333377377733333333333333333333333333333333}
      NumGlyphs = 2
    end
    object btnExit: TBitBtn
      Left = 600
      Top = 45
      Width = 75
      Height = 25
      Caption = #32080#26463
      TabOrder = 1
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
    object rgpQryCond: TGroupBox
      Left = 8
      Top = 8
      Width = 585
      Height = 105
      Caption = ' '#26597#35426#26781#20214' '
      TabOrder = 2
      object RadioButton0: TRadioButton
        Left = 8
        Top = 16
        Width = 65
        Height = 17
        Caption = #20844#21496#21029
        TabOrder = 0
      end
      object RadioButton1: TRadioButton
        Left = 8
        Top = 48
        Width = 57
        Height = 17
        Caption = #37096#38272
        TabOrder = 1
      end
      object RadioButton2: TRadioButton
        Left = 8
        Top = 80
        Width = 57
        Height = 17
        Caption = #32676#32068
        TabOrder = 2
      end
      object RadioButton3: TRadioButton
        Left = 256
        Top = 16
        Width = 113
        Height = 17
        Caption = #24115#34399
        TabOrder = 3
      end
      object RadioButton4: TRadioButton
        Left = 256
        Top = 48
        Width = 113
        Height = 17
        Caption = #22995#21517
        TabOrder = 4
      end
      object RadioButton5: TRadioButton
        Left = 256
        Top = 80
        Width = 113
        Height = 17
        Caption = #20840#37096#36039#26009
        Checked = True
        TabOrder = 5
        TabStop = True
      end
      object RadioButton6: TRadioButton
        Left = 420
        Top = 16
        Width = 105
        Height = 17
        Caption = #26159#21542#31995#32113#31649#29702#32773
        TabOrder = 6
      end
      object RadioButton7: TRadioButton
        Left = 420
        Top = 48
        Width = 113
        Height = 17
        Caption = #26159#21542#32676#32068#31649#29702#32773
        TabOrder = 7
      end
      object RadioButton8: TRadioButton
        Left = 420
        Top = 80
        Width = 113
        Height = 17
        Caption = #26159#21542#20572#29992
        TabOrder = 8
      end
      object dblCompCode: TDBLookupComboBox
        Left = 64
        Top = 16
        Width = 188
        Height = 21
        KeyField = 'CODENO'
        ListField = 'DESCRIPTION'
        ListSource = dsrAllCompCode
        TabOrder = 9
      end
      object dblDeptCode: TDBLookupComboBox
        Left = 64
        Top = 47
        Width = 188
        Height = 21
        KeyField = 'CODENO'
        ListField = 'DESCRIPTION'
        ListSource = dsrAllDeptCode
        TabOrder = 10
      end
      object dblGroupCode: TDBLookupComboBox
        Left = 65
        Top = 78
        Width = 188
        Height = 21
        KeyField = 'CODENO'
        ListField = 'DESCRIPTION'
        ListSource = dsrAllGroupCode
        TabOrder = 11
      end
      object chbIsSysOp: TComboBox
        Left = 526
        Top = 13
        Width = 49
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        ItemIndex = 0
        TabOrder = 12
        Text = #26159
        Items.Strings = (
          #26159
          #21542)
      end
      object chbIsSupervisor: TComboBox
        Left = 526
        Top = 44
        Width = 49
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        ItemIndex = 0
        TabOrder = 13
        Text = #26159
        Items.Strings = (
          #26159
          #21542)
      end
      object chbStopFlag: TComboBox
        Left = 526
        Top = 76
        Width = 49
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        ItemIndex = 0
        TabOrder = 14
        Text = #26159
        Items.Strings = (
          #26159
          #21542)
      end
      object edtUserID: TEdit
        Left = 302
        Top = 15
        Width = 110
        Height = 21
        TabOrder = 15
      end
      object edtUserName: TEdit
        Left = 300
        Top = 45
        Width = 111
        Height = 21
        TabOrder = 16
      end
    end
  end
  object dsrUserData: TDataSource
    DataSet = dtmMain.qryUserData
    Left = 616
    Top = 184
  end
  object dsrCompCode: TDataSource
    DataSet = dtmMain.qryCompCode
    Left = 648
    Top = 216
  end
  object dsrDeptCode: TDataSource
    DataSet = dtmMain.qryDeptCode
    Left = 648
    Top = 248
  end
  object dsrGroupCode: TDataSource
    DataSet = dtmMain.qryGroupCode
    Left = 648
    Top = 184
  end
  object dsrAllGroupCode: TDataSource
    DataSet = dtmMain.qryAllGroupCode
    Left = 584
    Top = 184
  end
  object dsrAllCompCode: TDataSource
    DataSet = dtmMain.qryAllCompCode
    Left = 584
    Top = 216
  end
  object dsrAllDeptCode: TDataSource
    DataSet = dtmMain.qryAllDeptCode
    Left = 584
    Top = 248
  end
end
