object fmMain: TfmMain
  Left = 154
  Top = 145
  Width = 790
  Height = 590
  Caption = 'DVN'#37197#23565#29138#27231#20316#26989'[SO1621L]'
  Color = clBtnFace
  Font.Charset = ANSI_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnCloseQuery = FormCloseQuery
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  PixelsPerInch = 96
  TextHeight = 14
  object Panel3: TPanel
    Left = 0
    Top = 513
    Width = 782
    Height = 44
    Align = alBottom
    BevelOuter = bvNone
    TabOrder = 2
    OnResize = Panel3Resize
    object btnImport: TcxButton
      Left = 159
      Top = 6
      Width = 95
      Height = 30
      Caption = #25972#25209#21295#20837
      TabOrder = 0
      OnClick = btnImportClick
    end
    object btnSetup: TcxButton
      Left = 281
      Top = 6
      Width = 95
      Height = 30
      Action = actDvnSetup
      TabOrder = 1
    end
    object btnClear: TcxButton
      Left = 404
      Top = 6
      Width = 95
      Height = 30
      Action = actDvnClear
      TabOrder = 2
    end
    object btnCancel: TcxButton
      Left = 527
      Top = 6
      Width = 95
      Height = 30
      Action = actCancel
      TabOrder = 3
    end
  end
  object MainPage: TcxPageControl
    Left = 0
    Top = 39
    Width = 782
    Height = 474
    ActivePage = cxTabNAGRA
    Align = alClient
    TabOrder = 1
    ClientRectBottom = 470
    ClientRectLeft = 2
    ClientRectRight = 778
    ClientRectTop = 23
    object cxTabDVN: TcxTabSheet
      Caption = 'DVN'#37197#23565#29138#27231
      ImageIndex = 0
      object Panel1: TPanel
        Left = 0
        Top = 0
        Width = 776
        Height = 197
        Align = alTop
        BevelOuter = bvNone
        TabOrder = 0
        object GroupBox1: TGroupBox
          Left = 0
          Top = 0
          Width = 776
          Height = 197
          Align = alClient
          TabOrder = 0
          object Label2: TLabel
            Left = 23
            Top = 55
            Width = 64
            Height = 14
            Caption = #27231#19978#30418#24207#34399':'
            Font.Charset = ANSI_CHARSET
            Font.Color = clRed
            Font.Height = -12
            Font.Name = 'Tahoma'
            Font.Style = []
            ParentFont = False
          end
          object Label3: TLabel
            Left = 23
            Top = 85
            Width = 64
            Height = 14
            Caption = #26234#24935#21345#21345#34399':'
            Font.Charset = ANSI_CHARSET
            Font.Color = clRed
            Font.Height = -12
            Font.Name = 'Tahoma'
            Font.Style = []
            ParentFont = False
          end
          object Label1: TLabel
            Left = 26
            Top = 156
            Width = 144
            Height = 14
            Caption = #38971#36947#26377#36984#34920#31034#26371#36865#35430#30475#38971#36947
            Font.Charset = ANSI_CHARSET
            Font.Color = clBlue
            Font.Height = -12
            Font.Name = 'Tahoma'
            Font.Style = []
            ParentFont = False
          end
          object Label4: TLabel
            Left = 195
            Top = 156
            Width = 204
            Height = 14
            Caption = #38928#35373#22825#25976' + '#24037#20316#32233#34909#26085', '#28858#35430#30475#25130#27490#26085
            Font.Charset = ANSI_CHARSET
            Font.Color = clBlue
            Font.Height = -12
            Font.Name = 'Tahoma'
            Font.Style = []
            ParentFont = False
          end
          object Label5: TLabel
            Left = 580
            Top = 123
            Width = 64
            Height = 14
            Caption = #24037#20316#32233#34909#26085':'
            Font.Charset = ANSI_CHARSET
            Font.Color = clBlue
            Font.Height = -12
            Font.Name = 'Tahoma'
            Font.Style = []
            ParentFont = False
          end
          object Panel5: TPanel
            Left = 8
            Top = 16
            Width = 170
            Height = 27
            BevelOuter = bvNone
            TabOrder = 0
            object rdoA1: TcxRadioButton
              Left = 13
              Top = 7
              Width = 55
              Height = 17
              Caption = #38283#27231
              Checked = True
              TabOrder = 0
              TabStop = True
            end
            object rdoA7: TcxRadioButton
              Left = 74
              Top = 7
              Width = 55
              Height = 17
              Caption = #37197#23565
              TabOrder = 1
            end
          end
          object txtDvnStb: TcxTextEdit
            Left = 95
            Top = 51
            Properties.MaxLength = 20
            TabOrder = 1
            OnKeyDown = txtDvnStbKeyDown
            Width = 165
          end
          object txtDvnIcc: TcxTextEdit
            Left = 95
            Top = 79
            Properties.MaxLength = 20
            TabOrder = 2
            OnKeyDown = txtDvnIccKeyDown
            Width = 165
          end
          object btnDvnCannel: TcxButton
            Left = 23
            Top = 118
            Width = 100
            Height = 25
            Caption = #38928#35373#38971#36947
            TabOrder = 3
            OnClick = btnDvnCannelClick
          end
          object btnChannelClear: TcxButton
            Left = 122
            Top = 118
            Width = 25
            Height = 25
            TabOrder = 4
            OnClick = btnChannelClearClick
            Glyph.Data = {
              72020000424D720200000000000036000000280000000E0000000D0000000100
              1800000000003C020000C40E0000C40E00000000000000000000C8D0D4C8D0D4
              C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D40000
              80FFFFFF0000C8D0D4C8D0D4000080FFFFFFC8D0D4C8D0D4C8D0D4C8D0D4C8D0
              D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D40000C8D0D4000080000080000080FFFF
              FFC8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4000080FFFFFFC8D0D40000C8D0
              D4000080000080000080FFFFFFC8D0D4C8D0D4C8D0D4C8D0D4C8D0D4000080FF
              FFFFC8D0D4C8D0D40000C8D0D4C8D0D4000080000080000080FFFFFFC8D0D4C8
              D0D4C8D0D4000080000080FFFFFFC8D0D4C8D0D40000C8D0D4C8D0D4C8D0D400
              0080000080000080FFFFFFC8D0D4000080000080FFFFFFC8D0D4C8D0D4C8D0D4
              0000C8D0D4C8D0D4C8D0D4C8D0D4000080000080000080000080000080FFFFFF
              C8D0D4C8D0D4C8D0D4C8D0D40000C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4000080
              000080000080FFFFFFC8D0D4C8D0D4C8D0D4C8D0D4C8D0D40000C8D0D4C8D0D4
              C8D0D4C8D0D4000080000080000080000080000080FFFFFFC8D0D4C8D0D4C8D0
              D4C8D0D40000C8D0D4C8D0D4C8D0D4000080000080000080FFFFFFC8D0D40000
              80FFFFFFC8D0D4C8D0D4C8D0D4C8D0D40000C8D0D40000800000800000800000
              80FFFFFFC8D0D4C8D0D4C8D0D4000080000080FFFFFFC8D0D4C8D0D400000000
              80000080000080000080FFFFFFC8D0D4C8D0D4C8D0D4C8D0D4C8D0D400008000
              0080FFFFFFC8D0D40000000080000080FFFFFFC8D0D4C8D0D4C8D0D4C8D0D4C8
              D0D4C8D0D4C8D0D4C8D0D4000080000080FFFFFF0000}
          end
          object txtDvnChannel: TcxTextEdit
            Left = 147
            Top = 119
            Properties.ReadOnly = True
            Style.Color = 14737632
            TabOrder = 5
            Width = 417
          end
          object txtDays: TcxTextEdit
            Left = 652
            Top = 119
            Properties.MaxLength = 2
            TabOrder = 6
            Text = '0'
            OnKeyPress = txtDaysKeyPress
            Width = 49
          end
          object chkDvnReSendErr: TcxCheckBox
            Left = 432
            Top = 152
            Caption = #35373#23450#26178', '#37325#26032#34389#29702#37679#35492#25110#36926#26178#30340#25351#20196
            TabOrder = 7
            Width = 229
          end
        end
      end
      object Panel2: TPanel
        Left = 0
        Top = 197
        Width = 776
        Height = 250
        Align = alClient
        BevelOuter = bvNone
        TabOrder = 1
        object Bevel1: TBevel
          Left = 0
          Top = 222
          Width = 776
          Height = 3
          Align = alBottom
          Shape = bsSpacer
        end
        object gdDVN: TcxGrid
          Left = 0
          Top = 0
          Width = 776
          Height = 222
          Align = alClient
          TabOrder = 0
          LookAndFeel.Kind = lfStandard
          object gvDVN: TcxGridDBBandedTableView
            NavigatorButtons.ConfirmDelete = False
            OnCustomDrawCell = gvDVNCustomDrawCell
            DataController.DataSource = dsDVN
            DataController.Summary.DefaultGroupSummaryItems = <>
            DataController.Summary.FooterSummaryItems = <>
            DataController.Summary.SummaryGroups = <>
            OptionsCustomize.ColumnFiltering = False
            OptionsCustomize.ColumnMoving = False
            OptionsCustomize.ColumnSorting = False
            OptionsCustomize.BandMoving = False
            OptionsData.Deleting = False
            OptionsData.Editing = False
            OptionsSelection.CellSelect = False
            OptionsView.GroupByBox = False
            OptionsView.Indicator = True
            Bands = <
              item
                Caption = #35373#20633#36039#35338
                FixedKind = fkLeft
                Width = 290
              end
              item
                Caption = #25351#20196#29376#24907
                Width = 156
              end
              item
                Caption = #34389#29702' / '#21295#20837
                Width = 350
              end>
            object gvDVNSeqNo1: TcxGridDBBandedColumn
              Caption = #25351#20196#24207#34399'1'
              DataBinding.FieldName = 'SEQNO1'
              Visible = False
              Width = 79
              Position.BandIndex = 0
              Position.ColIndex = 0
              Position.RowIndex = 0
            end
            object gvDVNSeqNo2: TcxGridDBBandedColumn
              Caption = #25351#20196#24207#34399'2'
              Visible = False
              Position.BandIndex = 0
              Position.ColIndex = 1
              Position.RowIndex = 0
            end
            object gvDVNStbNo: TcxGridDBBandedColumn
              Caption = #27231#19978#30418#24207#34399
              DataBinding.FieldName = 'STBNO'
              Width = 100
              Position.BandIndex = 0
              Position.ColIndex = 2
              Position.RowIndex = 0
            end
            object gvDVNIccNo: TcxGridDBBandedColumn
              Caption = #26234#24935#21345#21345#34399
              DataBinding.FieldName = 'ICCNO'
              Width = 100
              Position.BandIndex = 0
              Position.ColIndex = 3
              Position.RowIndex = 0
            end
            object gvDVNCmdA1: TcxGridDBBandedColumn
              Caption = #38283#27231'/'#37197#23565
              DataBinding.FieldName = 'CMDA1'
              HeaderAlignmentHorz = taCenter
              Position.BandIndex = 1
              Position.ColIndex = 0
              Position.RowIndex = 0
            end
            object gvDVNCmdB1: TcxGridDBBandedColumn
              Caption = #38283#35430#30475#38971#36947
              DataBinding.FieldName = 'CMDB1'
              HeaderAlignmentHorz = taCenter
              Position.BandIndex = 1
              Position.ColIndex = 1
              Position.RowIndex = 0
            end
            object gvDVNErrMsg: TcxGridDBBandedColumn
              Caption = #35338#24687
              DataBinding.FieldName = 'ERRMSG'
              Position.BandIndex = 2
              Position.ColIndex = 0
              Position.RowIndex = 0
            end
          end
          object glDVN: TcxGridLevel
            GridView = gvDVN
          end
        end
        object Panel4: TPanel
          Left = 0
          Top = 225
          Width = 776
          Height = 25
          Align = alBottom
          BevelOuter = bvNone
          Color = clWhite
          TabOrder = 1
          object lbDvnRecords: TLabel
            Left = 0
            Top = 0
            Width = 776
            Height = 25
            Align = alClient
            Alignment = taCenter
            Caption = ' 0 / 0 '
            Font.Charset = ANSI_CHARSET
            Font.Color = clBlue
            Font.Height = -12
            Font.Name = 'Tahoma'
            Font.Style = []
            ParentFont = False
            Transparent = True
            Layout = tlCenter
          end
          object DvnProgressBar: TcxProgressBar
            Left = 0
            Top = 0
            Align = alClient
            Position = 50.000000000000000000
            Properties.PeakValue = 50.000000000000000000
            TabOrder = 0
            Width = 776
          end
        end
      end
    end
    object cxTabNAGRA: TcxTabSheet
      Caption = 'NAGRA'#37197#23565#29138#27231
      ImageIndex = 1
      object GroupBox2: TGroupBox
        Left = 0
        Top = 0
        Width = 776
        Height = 191
        Align = alTop
        TabOrder = 0
        object Label8: TLabel
          Left = 25
          Top = 54
          Width = 64
          Height = 14
          Caption = #27231#19978#30418#24207#34399':'
          Font.Charset = ANSI_CHARSET
          Font.Color = clRed
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object Label9: TLabel
          Left = 25
          Top = 82
          Width = 64
          Height = 14
          Caption = #26234#24935#21345#21345#34399':'
          Font.Charset = ANSI_CHARSET
          Font.Color = clRed
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object Label10: TLabel
          Left = 26
          Top = 156
          Width = 144
          Height = 14
          Caption = #38971#36947#26377#36984#34920#31034#26371#36865#35430#30475#38971#36947
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlue
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object Label11: TLabel
          Left = 181
          Top = 156
          Width = 204
          Height = 14
          Caption = #38928#35373#22825#25976' + '#24037#20316#32233#34909#26085', '#28858#35430#30475#25130#27490#26085
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlue
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object Label12: TLabel
          Left = 582
          Top = 122
          Width = 64
          Height = 14
          Caption = #24037#20316#32233#34909#26085':'
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlue
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object Label7: TLabel
          Left = 37
          Top = 27
          Width = 52
          Height = 14
          Caption = #25805#20316#27169#24335':'
        end
        object lbOpeartion: TLabel
          Left = 250
          Top = 24
          Width = 279
          Height = 14
          Caption = #38283#27231' -> '#38283#35430#30475#38971#36947' -> '#21462#28040#25152#26377#38971#36947' -> '#38283#35430#30475#38971#36947
        end
        object Label13: TLabel
          Left = 594
          Top = 62
          Width = 52
          Height = 14
          Caption = 'Zip Code:'
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlue
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object Label14: TLabel
          Left = 591
          Top = 92
          Width = 55
          Height = 14
          Caption = 'BAT '#21312#30908':'
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlue
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object txtNagraStb: TcxTextEdit
          Left = 95
          Top = 51
          Properties.MaxLength = 20
          TabOrder = 1
          OnKeyDown = txtDvnStbKeyDown
          Width = 165
        end
        object txtNagraIcc: TcxTextEdit
          Left = 95
          Top = 81
          Properties.MaxLength = 20
          TabOrder = 2
          OnKeyDown = txtDvnIccKeyDown
          Width = 165
        end
        object btnNagraChannel: TcxButton
          Left = 23
          Top = 127
          Width = 100
          Height = 25
          Caption = #38928#35373#38971#36947
          TabOrder = 3
          OnClick = btnDvnCannelClick
        end
        object btnChannelClear2: TcxButton
          Left = 122
          Top = 127
          Width = 25
          Height = 25
          TabOrder = 4
          OnClick = btnChannelClearClick
          Glyph.Data = {
            72020000424D720200000000000036000000280000000E0000000D0000000100
            1800000000003C020000C40E0000C40E00000000000000000000C8D0D4C8D0D4
            C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D40000
            80FFFFFF0000C8D0D4C8D0D4000080FFFFFFC8D0D4C8D0D4C8D0D4C8D0D4C8D0
            D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D40000C8D0D4000080000080000080FFFF
            FFC8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4000080FFFFFFC8D0D40000C8D0
            D4000080000080000080FFFFFFC8D0D4C8D0D4C8D0D4C8D0D4C8D0D4000080FF
            FFFFC8D0D4C8D0D40000C8D0D4C8D0D4000080000080000080FFFFFFC8D0D4C8
            D0D4C8D0D4000080000080FFFFFFC8D0D4C8D0D40000C8D0D4C8D0D4C8D0D400
            0080000080000080FFFFFFC8D0D4000080000080FFFFFFC8D0D4C8D0D4C8D0D4
            0000C8D0D4C8D0D4C8D0D4C8D0D4000080000080000080000080000080FFFFFF
            C8D0D4C8D0D4C8D0D4C8D0D40000C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4000080
            000080000080FFFFFFC8D0D4C8D0D4C8D0D4C8D0D4C8D0D40000C8D0D4C8D0D4
            C8D0D4C8D0D4000080000080000080000080000080FFFFFFC8D0D4C8D0D4C8D0
            D4C8D0D40000C8D0D4C8D0D4C8D0D4000080000080000080FFFFFFC8D0D40000
            80FFFFFFC8D0D4C8D0D4C8D0D4C8D0D40000C8D0D40000800000800000800000
            80FFFFFFC8D0D4C8D0D4C8D0D4000080000080FFFFFFC8D0D4C8D0D400000000
            80000080000080000080FFFFFFC8D0D4C8D0D4C8D0D4C8D0D4C8D0D400008000
            0080FFFFFFC8D0D40000000080000080FFFFFFC8D0D4C8D0D4C8D0D4C8D0D4C8
            D0D4C8D0D4C8D0D4C8D0D4000080000080FFFFFF0000}
        end
        object txtNagraChannel: TcxTextEdit
          Left = 147
          Top = 128
          Properties.ReadOnly = True
          Style.Color = 14737632
          TabOrder = 5
          Width = 417
        end
        object txtNagraDays: TcxTextEdit
          Left = 652
          Top = 119
          Properties.MaxLength = 2
          TabOrder = 8
          Text = '0'
          OnExit = txtNagraDaysExit
          OnKeyPress = txtDaysKeyPress
          Width = 52
        end
        object chkNagraReSendErr: TcxCheckBox
          Left = 398
          Top = 154
          Caption = #35373#23450#26178', '#37325#26032#34389#29702#37679#35492#25110#36926#26178#30340#25351#20196
          TabOrder = 9
          Width = 217
        end
        object cmbNagraOpeartion: TcxComboBox
          Left = 95
          Top = 23
          Properties.DropDownListStyle = lsFixedList
          Properties.Items.Strings = (
            #26032#21697#38283#27231#37197#23565
            #33290#21697#38283#27231#37197#23565
            #33290#21697#22238#25910#28204#35430)
          Properties.OnChange = cmbOpeartionPropertiesChange
          TabOrder = 0
          Width = 137
        end
        object txtZipCode: TcxTextEdit
          Left = 652
          Top = 59
          Properties.MaxLength = 5
          Properties.ReadOnly = True
          TabOrder = 6
          Text = '0'
          OnExit = txtZipCodeExit
          OnKeyPress = txtDaysKeyPress
          Width = 52
        end
        object txtBatId: TcxTextEdit
          Left = 652
          Top = 89
          Properties.MaxLength = 5
          Properties.ReadOnly = True
          TabOrder = 7
          Text = '0'
          OnExit = txtBatIdExit
          OnKeyPress = txtDaysKeyPress
          Width = 52
        end
        object chkCHANCEDAYS: TCheckBox
          Left = 24
          Top = 108
          Width = 203
          Height = 17
          Caption = #38971#36947#38928#35373#35430#30475#22825#25976#28858'0'#32773#21487#20379#21246#36984
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlue
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
          TabOrder = 10
          OnClick = chkCHANCEDAYSClick
        end
      end
      object gdNagra: TcxGrid
        Left = 0
        Top = 191
        Width = 776
        Height = 231
        Align = alClient
        TabOrder = 1
        LookAndFeel.Kind = lfStandard
        object gvNagra: TcxGridDBBandedTableView
          NavigatorButtons.ConfirmDelete = False
          OnCustomDrawCell = gvDVNCustomDrawCell
          DataController.DataSource = dsNagra
          DataController.Summary.DefaultGroupSummaryItems = <>
          DataController.Summary.FooterSummaryItems = <>
          DataController.Summary.SummaryGroups = <>
          OptionsCustomize.ColumnFiltering = False
          OptionsCustomize.ColumnMoving = False
          OptionsCustomize.ColumnSorting = False
          OptionsCustomize.BandMoving = False
          OptionsData.Deleting = False
          OptionsData.Editing = False
          OptionsSelection.CellSelect = False
          OptionsView.GroupByBox = False
          OptionsView.Indicator = True
          Bands = <
            item
              Caption = #35373#20633#36039#35338
              FixedKind = fkLeft
              Width = 290
            end
            item
              Caption = #25351#20196#29376#24907
              Width = 380
            end
            item
              Caption = #34389#29702' / '#21295#20837
              Width = 350
            end>
          object gdNagraCol1: TcxGridDBBandedColumn
            Caption = #25351#20196#24207#34399'1'
            DataBinding.FieldName = 'SEQNO1'
            Visible = False
            Width = 79
            Position.BandIndex = 0
            Position.ColIndex = 0
            Position.RowIndex = 0
          end
          object gdNagraCol2: TcxGridDBBandedColumn
            Caption = #25351#20196#24207#34399'2'
            Visible = False
            Position.BandIndex = 0
            Position.ColIndex = 1
            Position.RowIndex = 0
          end
          object gdNagraCol3: TcxGridDBBandedColumn
            Caption = #27231#19978#30418#24207#34399
            DataBinding.FieldName = 'STBNO'
            Width = 100
            Position.BandIndex = 0
            Position.ColIndex = 2
            Position.RowIndex = 0
          end
          object gdNagraCol4: TcxGridDBBandedColumn
            Caption = #26234#24935#21345#21345#34399
            DataBinding.FieldName = 'ICCNO'
            Width = 100
            Position.BandIndex = 0
            Position.ColIndex = 3
            Position.RowIndex = 0
          end
          object gdNagraCol5: TcxGridDBBandedColumn
            Caption = #38283#27231
            DataBinding.FieldName = 'CMDA1'
            HeaderAlignmentHorz = taCenter
            Width = 81
            Position.BandIndex = 1
            Position.ColIndex = 0
            Position.RowIndex = 0
          end
          object gdNagraCol6: TcxGridDBBandedColumn
            Caption = #38283#35430#30475#38971#36947
            DataBinding.FieldName = 'CMDB1'
            HeaderAlignmentHorz = taCenter
            Width = 98
            Position.BandIndex = 1
            Position.ColIndex = 1
            Position.RowIndex = 0
          end
          object gdNagraCol7: TcxGridDBBandedColumn
            Caption = #21462#28040#25152#26377#38971#36947
            DataBinding.FieldName = 'CMDB3'
            HeaderAlignmentHorz = taCenter
            Width = 111
            Position.BandIndex = 1
            Position.ColIndex = 2
            Position.RowIndex = 0
          end
          object gdNagraCol8: TcxGridDBBandedColumn
            Caption = #38283#35430#30475#38971#36947
            DataBinding.FieldName = 'CMDB1_1'
            HeaderAlignmentHorz = taCenter
            Width = 99
            Position.BandIndex = 1
            Position.ColIndex = 3
            Position.RowIndex = 0
          end
          object gdNagraCol10: TcxGridDBBandedColumn
            Caption = #35338#24687
            DataBinding.FieldName = 'ERRMSG'
            Position.BandIndex = 2
            Position.ColIndex = 0
            Position.RowIndex = 0
          end
        end
        object glNagra: TcxGridLevel
          GridView = gvNagra
        end
      end
      object Panel6: TPanel
        Left = 0
        Top = 422
        Width = 776
        Height = 25
        Align = alBottom
        BevelOuter = bvNone
        Color = clWhite
        TabOrder = 2
        object lbNagraRecords: TLabel
          Left = 0
          Top = 0
          Width = 776
          Height = 25
          Align = alClient
          Alignment = taCenter
          Caption = ' 0 / 0 '
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlue
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
          Transparent = True
          Layout = tlCenter
        end
        object NagraProgressBar: TcxProgressBar
          Left = 0
          Top = 0
          Align = alClient
          Position = 50.000000000000000000
          Properties.EndColor = clNavy
          Properties.PeakValue = 50.000000000000000000
          TabOrder = 0
          Width = 776
        end
      end
    end
  end
  object cxGroupBox1: TcxGroupBox
    Left = 0
    Top = 0
    Align = alTop
    PanelStyle.Active = True
    TabOrder = 0
    Height = 39
    Width = 782
    object Label6: TLabel
      Left = 29
      Top = 12
      Width = 86
      Height = 14
      Caption = #35531#36984#25799'CAS'#31995#32113':'
    end
    object cmbCAS: TcxComboBox
      Left = 124
      Top = 8
      Properties.DropDownListStyle = lsFixedList
      Properties.OnChange = cmbCASPropertiesChange
      TabOrder = 0
      Width = 189
    end
  end
  object dsDVN: TDataSource
    DataSet = cdsDVN
    Left = 83
    Top = 416
  end
  object cdsDVN: TClientDataSet
    Aggregates = <>
    Params = <>
    AfterScroll = cdsDVNAfterScroll
    Left = 115
    Top = 416
  end
  object cxLookAndFeelController1: TcxLookAndFeelController
    Kind = lfStandard
    NativeStyle = True
    Left = 79
    Top = 372
  end
  object ActionList1: TActionList
    Left = 115
    Top = 372
    object actDvnSetup: TAction
      Caption = 'F6.'#35373#23450
      ShortCut = 117
      OnExecute = actDvnSetupExecute
    end
    object actCancel: TAction
      Caption = #21462#28040'(C)'
      ShortCut = 32835
      OnExecute = actCancelExecute
    end
    object actDvnClear: TAction
      Caption = 'F4.'#28165#38500
      ShortCut = 115
      OnExecute = actDvnClearExecute
    end
    object actNagraSetup: TAction
      Caption = 'F6.'#35373#23450
      ShortCut = 117
      OnExecute = actNagraSetupExecute
    end
    object actNagraClear: TAction
      Caption = 'F4.'#28165#38500
      ShortCut = 115
      OnExecute = actNagraClearExecute
    end
  end
  object cdsDvnChannel: TClientDataSet
    Aggregates = <>
    Params = <>
    Left = 147
    Top = 416
  end
  object CmdImageList: TImageList
    Left = 151
    Top = 372
    Bitmap = {
      494C010110001300040010001000FFFFFFFFFF10FFFFFFFFFFFFFFFF424D3600
      0000000000003600000028000000400000005000000001002000000000000050
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      000000000000000000000000000000000000FFFFFF00FFFFFF00FFFFFF00FFFF
      FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
      FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      000000000000000000000000000000000000FFFFFF00FFFFFF00FFFFFF00FFFF
      FF00FFFFFF00F5E5DA00CB885B00C16B2700BE682700C07D5700F1DED600FFFF
      FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      00000000000000000000888A8B00787B7E00B6B7B9002E8D3200318F3500A1CE
      A300000000000000000000000000000000000000000000000000000000000000
      00000000000000000000888A8B00B4B6B7008F909000797979009A5A34007A73
      70009B9B9B00000000000000000000000000FFFFFF00FFFFFF00FFFFFF00FFFF
      FF00F6E1D100CB6F1F00FCB44D00F9B55400F1AD5400E7A04E00B75D1E00EFD6
      CA00FFFFFF00FFFFFF00FFFFFF00FFFFFF000000000000000000000000000000
      0000000000004F73590000000000000000000000000000000000CF722F00CF72
      2F00CF722F00CF722F0000000000000000000000000000000000000000000000
      00008982850090949A00B5ACA400C8B6A300E2D8CE002E8D320060E994003AC3
      5B0053A857000000000000000000000000000000000000000000000000000000
      00008982850090949A00B5ACA400EAE4DD0085858400F1DCCF00A8430500C672
      3A0085858400000000000000000000000000FFFFFF00FFFFFF00FFFFFF00FFFF
      FF00DA843E00FFC26100FFB94E00F7B04D00F1AA4D00EFA84F00F2AD5200C46C
      3200FFFFFF00FFFFFF00FFFFFF00FFFFFF000000000000000000000000000000
      00000000000052AE690051856100000000000000000000000000D29A7600F9E8
      AD00FCD89900CF722F000000000000000000000000000000000000000000928B
      8E00BCB7B100FED8AC00FFE8B600FFE8BD00FFF4DE002E8D320046CA720033D4
      6E004DE9840037AD45000000000000000000000000000000000000000000928B
      8E00BCB7B100FED8AC00FFE8B600FFF5E50093908E00B6662E00A8430500B24C
      0100B06E3B00000000000000000000000000FFFFFF00FFFFFF00FFFFFF00FCF1
      E800DA832A00FFC66900FFBF5600FFB74D00F8B14D00F5AE4D00FAB35200CD75
      2300F6E5DD00FFFFFF00FFFFFF00FFFFFF000000000000000000000000000000
      0000000000009CFEBC0068D28400518561000000000000000000D39D7600F0DF
      A500FCEBB100CF722F00000000000000000000000000000000009C8F9000C7C1
      BA00FFE4B200FFE4B900FFE4B200FFE0B400FFEFDA002C8930004AC16B0027BF
      560033D46E004FED8B002FA340000000000000000000000000009C8F9000C7C1
      BA00FFE4B200FFE4B900FFE4B200FFF5E500B0754A00A8430500B96E3B00BB65
      2100C05C0400000000000000000000000000FFFFFF00FFFFFF00B0B4F3003539
      B600D17F4200C18D8200C6895500FCB75400FFBB5200FFB74E00FFBC5200DC82
      2B00F3DBD000FFFFFF00FFFFFF00FFFFFF000000000000000000000000000000
      00000000000089E9AA0082C19700000000000000000000000000D29A7600D29A
      7600D29A7600CF722F0000000000000000000000000000000000B6B7B900FFE0
      B400FFE4B900FFE4B900FFE5BD00FFE8BD00FFF4DE002E86320083CD950061C7
      7D0028BC55003BA248009EC29E00000000000000000000000000B2B2B800FFE4
      B200FFE4B900FFE4B900FFE4B900FFF5E500B6662E00B3723800C7C5C400BE96
      7400C3610A00CC680300F8BC6B0000000000FFFFFF00B7BCF200203BDD004B6E
      FF00D67C3900EBBBC300DDABB200C1846200DAAE3A00C6B13900C3B93E00BA61
      0C007EBA8600F9FCF900FFFFFF00FFFFFF00CF722F00CF722F00CF722F00CF72
      2F000000000082C1970000000000000000000000000000000000000000000000
      00000000000000000000000000000000000000000000AE9A9B00DCD2C600FFE8
      BD00FFEAC200FFEAC200FFF2C900EAD4AF00F0E6D200388C3B0091D29F006DB5
      73006FAF6B00DAD3CD00989B9D000000000000000000AE9A9B00DCD2C600FFED
      BD00FFEAC200FFEAC200FFF2C900F6EDDD00EFD6BF00FDF1DC00FFF5E500FFF5
      E500E3A56A00D16A0000E3810B0000000000FFFFFF00384FDB00567EFF00496C
      FF00835C7200F0AA8A00E8B6BE00C6B27300D9D76D00ECF48E00E0C664008783
      2E003FB55100499D5100FAFCFA00FFFFFF00D29A7600F9E8AD00FCD89900CF72
      2F00000000000000000000000000000000000000000000000000000000000000
      00000000000000000000000000000000000000000000AC969A00EEDEC500FFF4
      DE00FFEECB00FFF2C900FFF9D10046485000989B9D004391430043914300B9D1
      AD00FFF5E100D3BCA30076797B000000000000000000AC969A00EEDEC500FDF1
      DC00FFEECB00FFF2C900FFF9D1006C6E7400757A7E00FFE4B200FFEAC200FFEA
      C200FFF7E800E1934300DA75020000000000FFFFFF001B41E6005986FF004E73
      FF00446CFF008D5D6600D7782500ECCFBB00DCB25800DF9C3D0093802E0057D0
      76005ECB700047B4550090C19400FFFFFF00D39D7600F0DFA500FCEBB100CF72
      2F00000000000000000000000000000000000000000000000000000000000000
      00000000000000000000000000000000000000000000B19C9F00F5E5C900FFF9
      DE00FFF9DE00FFF9DE00B3AB9200383C43002E303800C1BCB100FFEFDA00FFEF
      D000FFEAC200D8BFA70076797B000000000000000000B19C9F00F5E5C900FFF8
      E000FFF9DE00FFF9DE00B3AB9200383C43002E3038008B816D00F8E0B800FFE4
      B900FFF2D500EDE2D700CE76130000000000FFFFFF00264BED006193FF00537E
      FF004F76FF00496BFF002BBB88003EBED200718B4B006FCA7E0062D7870063CF
      7C0063CC78006BD6820043A14F00FFFFFF00D29A7600D29A7600D29A7600CF72
      2F00000000004F73590000000000000000000000000000000000CF722F00CF72
      2F00CF722F00CF722F00000000000000000000000000BBA0A200F5E9D900FFF9
      DE00FFFFF100CFCBB20001030B00DFD7B600F7EAC6000E111800726A5C00FFEE
      C100FFEAC200C5B7A700878688000000000000000000BBA0A200F6EDDD00FFF9
      DE00FFFFF100CFCBB20001030B00DFD7B600F7EAC6000E111800726A5C00FFED
      C100FFEAC200E0D8D000BDBCBD0000000000FFFFFF00788BF8004C82FF005F91
      FF005986FF00557EFF0046DEA0003AC4D30066D6940070DD96006CD78C0069D4
      870068D385006ED98D0038A34C00FFFFFF000000000000000000000000000000
      00000000000052AE690051856100000000000000000000000000D29A7600F9E8
      AD00FCD89900CF722F0000000000000000000000000000000000E4E0E500FFF2
      C900E6E5D70000000000C3C0AB00FFFFE500FFF9DE00F2E3BF00A0948000FFEE
      C100FFE8B60095979B0000000000000000000000000000000000E4E0E500FFF2
      C900E6E5D70000000000C3C0AB00FFF8E000FFF9DE00F8E0B800A0948000FFED
      C100FFE8B60095979B000000000000000000FFFFFF00F7F7FE004F72FF003566
      FB005487FC00496FFF0030B56F0063CD9A007DE69C0075E1980072DC950070DA
      930070DB920078E19D004DAC5F00FFFFFF000000000000000000000000000000
      0000000000009CFEBC0068D28400518561000000000000000000D39D7600F0DF
      A500FCEBB100CF722F0000000000000000000000000000000000C7ABAC00FFFF
      F100D7CDAF00A2A49D00FFFFF800FFFFE500FFF9D100FFF2C900FFF2C900FFEE
      C100CBC1B5008982850000000000000000000000000000000000C7ABAC00FFFF
      F100D7CDAF00A2A49D00FFFFF800FFF8E000FFF9D100FFF9D100FFF2C900FFED
      C100CBC1B500898285000000000000000000FFFFFF00FFFFFF00FFFFFF00BEC8
      FF008592F8009296F70076CC970077E9980083EBA6007AE59F0078E29B0077E2
      99007BE59F0058C97900A1D5AC00FFFFFF000000000000000000000000000000
      00000000000089E9AA0082C19700000000000000000000000000D29A7600D29A
      7600D29A7600CF722F000000000000000000000000000000000000000000D5BE
      C200FFFFF800FFF9D100FFFFE500FFFFE500FFF9DE00FFF9D100FFE8BD00D5CA
      BF00938D9000000000000000000000000000000000000000000000000000D5BE
      C200FFFFF800FFF9D100FFFFE500FFF8E000FFF8E000FFF9D100FFEDBD00D5CA
      BF0093908E00000000000000000000000000FFFFFF00FFFFFF00FFFFFF00FFFF
      FF00FFFFFF00FFFFFF00F4FDF7005CCC7D0070E294008AF1AC0085EDA80085ED
      AA005ED2820066C37E00FDFEFD00FFFFFF000000000000000000000000000000
      00000000000082C1970000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000C9ABAD00EAE7E800FBF0DF00FFEECB00F5E5C900E2D8CE00BCB8BC009C8F
      9000000000000000000000000000000000000000000000000000000000000000
      0000C9ABAD00EAE7E800FFF5E500FAEBD000F5E5C900E7DACA00BCB8BC009C8F
      900000000000000000000000000000000000FFFFFF00FFFFFF00FFFFFF00FFFF
      FF00FFFFFF00FFFFFF00FFFFFF00F9FDFA0096DCAB0057C977004CC76F0057C3
      7400A3DCB200FDFEFD00FFFFFF00FFFFFF000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      00000000000000000000BBA0A200B49EA100AE9A9B00B19C9F00000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      00000000000000000000BBA0A200B49EA100AE9A9B00AE9A9B00000000000000
      000000000000000000000000000000000000FFFFFF00FFFFFF00FFFFFF00FFFF
      FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
      FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      000000000000000000000000000000000000CB7A3B00CB7A3B00CB7A3B00D175
      3300C26C3400000000000000000000000000000000000000000000000000CB7A
      3B00CB7A3B00CB7A3B00D1753300C26C34000000000000000000000000000000
      000000000000000000000000000000000000000000000000000000000000005F
      7F00000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      000000000000000000000000000000000000D39D7600FCD38F00FBBA7100FCA3
      5100CF722F00000000000000000000000000000000000000000000000000D39D
      7600FCD38F00FBBA7100FCA35100CF722F000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000007FAA0000BF
      FF00005F7F000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000005561
      5800000000000000000000000000000000000000000000000000000000000000
      00000000000000000000000000000000FF000049DC000049DC002557FF002557
      FF002525FF00000000000000000000000000D39D7600F9E8AD00FCD89900FBBA
      7100D0793A00000000000000000000000000000000000000000000000000D39D
      7600F9E8AD00FCD89900FBBA7100D0793A000000000000000000000000000000
      00000000000000000000000000000000000000000000009FD400AADFFF00AADF
      D40055BFFF00005F7F0000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000001454
      240051534F000000000000000000000000000000000000000000000000000000
      000000000000000000000000FF004873FF00D4F0FF008ED4FF006BC6FF006B8F
      FF002525FF00000000000000000000000000D39D7600F0DFA500FCEBB100F9CE
      8C00CB7A3B00000000000000000000000000000000000000000000000000D39D
      7600F0DFA500FCEBB100F9CE8C00CB7A3B000000000000000000000000000000
      0000000000000000000000000000553F7F000000000000000000009FD400AADF
      FF0055DFFF0055BFFF00005F7F00000000000000000000000000000000000000
      0000000000000000000000000000000000004C6B5600344637003E483D004694
      580021602E0066676700000000000000000000000000000000000000000000B9
      00000096000000000000000000004873FF008ED4FF006BC6FF006BC6FF000049
      DC000000DC00000000000000000000000000D39D7600D39D7600D39D7600D39D
      7600D39D76000000000000000000000000000000000000000000D8D3D500D39D
      7600D39D7600D39D7600D39D7600D39D760000000000559FAA00559FAA00557F
      AA00AA9FAA00AABFD400AA3FAA00FF7FFF00AA3F7F000000000000000000009F
      D400AADFFF00009FAA0000000000000000000000000000000000000000000000
      0000000000000000000000000000000000008BD2A2007CDF99005CBB790054AB
      6D0054B66C001C5227009095910000000000000000000000000000B900000096
      000000B9000000000000000000002557FF006BC6FF006BC6FF006B8FFF004873
      FF000049DC000000DC0000000000000000000000000000000000A3918F000000
      00000000000000000000000000000000000000000000D7CFD000A3918F000000
      000000000000000000000000000000000000559FAA00559FD400AADFD400AADF
      D400AAFFFF00AA5FD400FF7FD400FF9FFF00FF7FFF00AA3F7F00000000000000
      0000009FD4000000000000000000000000000000000000000000000000000000
      00000000000000000000000000000000000090D9AB00A0FFC20090F7B00086E2
      A1005FC87B004CA5630096AD9C00000000000000000000960000009631000096
      00000000000000000000000000002557FF006B8FFF002557FF004873FF002557
      FF000055FF000000DC000000B900000000000000000000000000A3918F000000
      000000000000000000000000000000000000D8D3D500A3918F00000000000000
      000000000000000000000000000000000000559FAA0055DFFF0055DFFF00AADF
      FF00AA5FD400FF9FFF00FF9FFF00FF9FFF00AA5FD40000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      00000000000000000000000000000000000083C197008AC59F007EA88B00A0E6
      B80066C8820092AE9B000000000000000000000000000096310000B93D000096
      00000000000000000000000000000055FF004873FF000000DC000000DC000049
      DC000055FF000055FF000000B900000000000000000000000000A3918F000000
      0000000000000000000000000000D7CFCF00A3918F0000000000000000000000
      000000000000000000000000000000000000557FAA0055DFFF00AADFFF00AADF
      D400AADFFF00FF5FD400FFBFFF00AA5FD400AABFD40000000000800000000000
      0000000000000000000000000000000000000000000000000000000000000000
      00000000000000000000000000000000000000000000000000000000000080D0
      9A0086AD9300000000000000000000000000000000000096310000B93D000096
      00000000000000000000000000000000FF000000FF000000FF00000000000000
      DC000000DC000049DC000000B900000000000000000000000000A3918F000000
      00000000000000000000D8D3D500A3918F000000000000000000000000000000
      000000000000000000000000000000000000559FAA0055BFFF00AADFFF00AAFF
      FF0055DFFF0055BFFF00FF5FAA00AABFFF00AABFD400AA5F2A00FF9F5500AA3F
      2A0000000000000000000000000000000000E1C8BC00AD9C9500B2A09900B2A0
      9900B2A09900B2A09900B2A09900B2A09900B2A09900B2A09900B2A099006E99
      7A00AE9C9500000000000000000000000000000000000096310025FF570000B9
      3D00008000000000000000800000008000000080000000000000000000000000
      00000000DC000000DC000000B900000000000000000000000000A3918F000000
      000000000000D7CFCF00A3918F00000000000000000000000000000000000000
      000000000000000000000000000000000000559FAA0055DFFF0055DFD400AADF
      FF00AABFD40055BFFF0055BFFF0055BFD400AA7F5500F0CAA600AABF7F00FF9F
      550080000000000000000000000000000000E1C8BC00C2BFBE004B4A4900B9B4
      B100FBF2EE006E696700FAEDE600C3B8B2004A454200B5A8A100F7E3DA006C62
      5E00AE9C9500000000000000000000000000000000000096000048FF730048FF
      730000B93D0000800000007300008EFFAB000096000000000000000000000000
      00000000B9000000DC000000B900000000000000000000000000A3918F000000
      0000D8D3D500A3918F000000000000000000000000000000000000000000CB7A
      3B00CB7A3B00CB7A3B00D1753300C26C3400559FAA0055DFFF00AADFFF00AADF
      FF0055DFFF0000BFD400009FFF00559FD400AABFD400AA7F5500FFBF7F00FFBF
      7F00AA9F5500AA1F2A000000000000000000E1C8BC0049494900FEF9F7004B4A
      4800FCF4F000837E7C00FAEFE9004A464400F8E8E1004A454200F7E3DA007E74
      6E00AE9D9500000000000000000000000000000000000080000000B93D0048FF
      730048FF73006BFF8F0000B93D008EFFAB000096310000000000000000000000
      00000000B9000000B9000000B90000000000CB7A3B00CB7A3B00CB7A3B00D175
      3300C26C3400000000000000000000000000000000000000000000000000D39D
      7600FCD38F00FBBA7100FCA35100CF722F00559FAA0055BFFF0055DFFF00AAFF
      FF0055DFD40055BFFF00009FD400009FD400557FAA0000000000AA9F5500F0CA
      A600FF5F5500000000000000000000000000E1C8BC00BEBCBB004E4D4C00AEAA
      A800FCF6F200C9C0BB00FAF0EB00BBB2AF004C484500AA9F9A00F7E5DC00B7AB
      A500AF9E960000000000000000000000000000000000000000000080000000B9
      3D006BFF8F006BFF8F008EFFAB008EFFAB0000B93D0000000000000000000000
      B9000000B9000000B9000000000000000000D39D7600FCD38F00FBBA7100FCA3
      5100CF722F00A3918F00A3918F00A3918F00A3918F00A3918F00A3918F00D39D
      7600F9E8AD00FCD89900FBBA7100D0793A00559FD40055DFFF00AADFD400AADF
      FF0055DFFF0000BFD400009FFF00009FD400007F7F000000000000000000AA9F
      550000000000000000000000000000000000E1C8BC00B9B5B400FEFBFA00C6C4
      C3004B4A4900B9B4B100FCF2EE00BBB2AE00F9ECE600C3B7B1004A454200B5A7
      A100B2A099000000000000000000000000000000000000000000000000000073
      0000008000006BFF8F008EFFAB008EFFAB0000B93D0000000000000000000000
      DC000000B900000000000000000000000000D39D7600F9E8AD00FCD89900FBBA
      7100D0793A00000000000000000000000000000000000000000000000000D39D
      7600F0DFA500FCEBB100F9CE8C00CB7A3B00559FAA0055DFFF0055DFFF00AAFF
      FF00AADFFF00AADFFF0055BFFF0055BFD400557FAA0000000000000000000000
      000000000000000000000000000000000000E1C8BC0088888800FEFDFC004C4B
      4A00FDF8F7004B494800FCF3EF00827E7B00FAEDE8004A464400F8E9E1004A44
      4200B2A199000000000000000000000000000000000000000000000000000096
      31006BFF8F008EFFAB008EFFAB008EFFAB003DB9000000000000000000000000
      000000000000000000000000000000000000D39D7600F0DFA500FCEBB100F9CE
      8C00CB7A3B00000000000000000000000000000000000000000000000000D39D
      7600D39D7600D39D7600D39D7600D39D7600559FAA00AADFFF00F0FBFF00F0FB
      FF00F0FBFF00F0FBFF00AAFFFF00AADFD400557FAA0000000000000000000000
      000000000000000000000000000000000000E1C8BC0069696900FEFEFD00B9B7
      B7004C4B4A00AEAAA800F8F1ED0068666500F4EAE400B2A9A40048444200A59B
      9600AE9D95000000000000000000000000000000000000000000000000000096
      3100009631000096310000732500007300000073000000000000000000000000
      000000000000000000000000000000000000D39D7600D39D7600D39D7600D39D
      7600D39D76000000000000000000000000000000000000000000000000000000
      00000000000000000000000000000000000000000000AABFD400559FAA00559F
      AA00559FAA00559FAA00559FAA00AABFD4000000000000000000000000000000
      000000000000000000000000000000000000E1C8BC00DCC4B800DCC4B800DCC3
      B700DCC4B800DCC3B700E0C7BB00E2C9BD00E1C8BC00E2C9BD00E2C9BD00E1C8
      BC00A89791000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000BFBFBF009F9F9F00996666009966660099666600996666007F7F7F00BFBF
      BF00000000000000000000000000000000000000000000000000000000000000
      00000000000083E5A7003B8C39001B7814001B7613003987370083E5A7000000
      0000000000000000000000000000000000000000000000000000000000000000
      00000000000095B2FE00383BBA00141AAD00141AAD003839B10095B2FE000000
      0000000000000000000000000000000000000000000071717100717171007171
      71006D6D6D006D6D6D006D6D6D006D6D6D006D6D6D0000000000000000000000
      000000000000000000000000000000000000000000000000000000000000CC99
      9900CC666600FFC7B100D4F0FF00FFFFFF00FFFFFF00D4F0FF00BFBFBF009F9F
      9F009F9F9F0000000000000000000000000000000000000000000000000083E5
      A70016841600088800000888000013880000217D000032720000266D00001B76
      130083E5A70000000000000000000000000000000000000000000000000095B2
      FE001323C3000020D9000023E1000020D9000019CC000011BD00000BB0001416
      A70095B2FE000000000000000000000000006D6D6D00B1654D008C5036008C50
      36008C5036008C5036006C554C006A635F00717171006D6D6D006D6D6D006D6D
      6D006D6D6D0071717100717171006D6D6D000000000000000000CC999900FFAB
      8E00FFF0D400FFB1B100CC666600CC660000CC660000CC666600FFB1B100FFE3
      D400CC9999009F9F9F000000000000000000000000000000000083E5A7000787
      090004A00800009D010004A00800009D0100009D0100009500001E7E0000396B
      0000126C040083E5A7000000000000000000000000000000000095B2FE00041F
      D500002EFF00002EFF00002DF800002DF8000029EE00001FE5000012CA00000B
      B0000408A40095B2FE000000000000000000E0664200DC513100F9635000FD72
      5C00F1684E00D7744500F1684E00D24C270035593800355938002F982D004890
      37002F982D0025662500355938006D6D6D0000000000BFBFBF00CC999900FFF0
      D400FFAB8E00B93D0000B93D0000FFAB8E00FFFFFF00B93D0000B93D0000CC99
      9900FFF0D40099666600BFBFBF000000000000000000000000001D911E000BA8
      170004A0080036C15400DCF3E9007ADD9B00009D010000950000009D01001285
      0000396B000017751300000000000000000000000000000000001831D5000034
      FF0097A9E8006888F4000034FF00002EFF000031FA0095B2FE004570FB00041F
      D500000BB0001315A7000000000000000000D5885500F9635000FD7C6400FC8B
      6D00D5885500FFCF9C00F36D5300FD6A5700917930004DB34D006AA9660063CA
      630052C252003CB73C00239023006D6D6D0000000000CC999900FFE2B400FFC7
      B100FF572500CC660000DC490000DE9A3E00FFAB8E00B93D0000B93D0000B93D
      0000FFB1B100BFBFBF007F7F7F00000000000000000083E5A70024A82F0008A9
      1B002BB43E00F6F0F500FCF6FA00FFFEFF007ADD9B00009E0700009D01000095
      0000217D0000266D000083E5A700000000000000000095B2FE00254EF8000034
      FF00CACDDC00F5EFDE00B1C0EF000639FC00B1C4FA00FFFFFA00FFFFFA006187
      FD000012CA00000BB00095B2FE000000000000000000E0664200FD846900EB80
      5900FFCF9C00FFCF9C00EF765600FD6B580080AB670080B38000FFEFDF004DB3
      4D0075D275004FC14F00489037006D6D6D0000000000DE9A3E00FFFFFF00FF8F
      6B00FF572500FF572500FF572500FFAB8E00FFFFFF00DC490000B93D0000B93D
      0000CC666600D4F0FF0099666600000000000000000040A23F0039BF4F0024A8
      2F00ECE6EA00F5EEF400FCF6FA00FFFEFF00FFFEFF007ADD9B00009E0700009D
      010000950000327200003987370000000000000000003950DF003D6AFF00083C
      FE0097A3E200F4F1E100F6F3EA00E4E8F300FFFFFA00FFFFFA00F3F7FF00446D
      FD000022E7000011BD00383BBA00000000000000000000000000B1654D004C21
      4F002E3457007F4C6E00B1654D00917930007CD77C008FE08F00FFF7E800BFD9
      AC005BC55B0044A64400000000000000000000000000DE9A3E00FFFFFF00FF73
      4800FF572500FF572500FF572500FF734800FFFFFF00FFC7B100CC660000B93D
      0000B93D0000FFFFFF00996666000000000000000000279D2B0039BF4F00AAC9
      A700FAEAF900D1E1D1001EAF3300DAEEDD00FFFEFF00FFFEFF007ADD9B00009E
      0700009D0100297C00001B76130000000000000000002241E2004B76FF002255
      FF000D40FE005C7AF000EFEDEB00F6F6F300FFFFFA00CDDAFE00174AFF00002E
      FF00002DF8000019CC00121AB10000000000000000001717170005070E000D29
      6800143AA000102F9500081D6C00545454006AA96600427399001579BA00247D
      B600377D57004D4D4D00000000000000000000000000DE9A3E00FFFFFF00FF73
      4800FF734800FF572500FF572500FF572500FF8F6B00FFFFFF00FFE2B400B93D
      0000CC660000FFFFFF009966660000000000000000002AA02D005DD27C0017A3
      25008EBF880011AB24001BB940000BAC2600DAEEDD00FFFEFF00FFFEFF007ADD
      9B00009D01001B8600001B78140000000000000000002241E2006489FF002D5F
      FF003364FF000639FC00DADCED00FBF9F300FFFFFA00ADC0FE000034FF000034
      FF00002EFF000020D900121AB100000000003F3F3F001A1A1A00102C5B001A4D
      B3001C56BC001B51B700102F950054545400699AAE002C92F1003399FF003399
      FF002C92F1002C586F00848484000000000000000000CC999900FFFFFF00FFAB
      8E00FF734800FF734800FFAB8E00FF734800FF572500FF8F6B00FFFFFF00FF73
      4800CC666600FFD4F00099666600000000000000000041A7410077DC95002CC5
      59001BB940002CC5590026C04E0022BB45000BAC2600DAEEDD00FFFEFF00FFFE
      FF0078DB9600138800003890380000000000000000003C57E80083A2FF003D6A
      FF003D6AFF00526DE900FFFBE900B2BFF300D4DAF900FFFFFA006D8FFF000034
      FF000034FF000023E1003940C1000000000012121200282828000F2D93002774
      DA002671D7002671D7001E5AC0005A6064003F95C30040A6FF0040A6FF0040A6
      FF003DA2FF002385C6006D6D6D000000000000000000CC999900FFE2B400FFCA
      CC00FF734800FF8F6B00FFFFFF00FFC7B100FF8F6B00FFC7B100FFFFFF00FF73
      4800FFC7B100FFC7B100CC999900000000000000000083E5A70053C66C0061D8
      890031CA630031CA63002CC5590026C04E0022BB45000BAC2600DAEEDD00FCF6
      FA00CAECD7000E8E080083E5A700000000000000000095B2FE006D8FFF006F95
      FF002958F700BABEDE00F4F1E1001745FC000D40FE00E7ECFB00FFFFFA002D5F
      FF000034FF000020D90095B2FE00000000002C2C2C00363636002C586F003191
      F9003399FF003694F700246AD0005A606400479FD0004BB1FF004BB1FF004DB3
      FF0049AFFF002D92E600666666000000000000000000DCDCDC00FFAB8E00FFFF
      FF00FFCC9900FF734800FFB1B100FFFFFF00FFFFFF00FFF0D400FF8F6B00FFB1
      B100FFF0D400CC666600DCDCDC0000000000000000000000000027A32C0083E5
      A7004CD3790031CA630031CA630027C3530022BB45001CB53A000BA81F007ACC
      850011AB240016841600000000000000000000000000000000002A4CED0095B2
      FE004B76FF005B6CD800546BE1003364FF002D5FFF001745FC00C7D1FA004B76
      FF000034FF001629CA0000000000000000006D6D6D004A4A4A003F3F3F004646
      4600143FA4002060C600135F88007E7E7E0057A9D7004DB3F2004DB3F20055BB
      FF0051B7FF0043A8ED006D6D6D00000000000000000000000000FFC7B100FFB1
      B100FFFFFF00FFCACC00FFAB8E00FF734800FF734800FFAB8E00FFCACC00FFF0
      D400CC999900BFBFBF000000000000000000000000000000000083E5A70028A8
      300081E3A6005ED8880031CA630027C3530022BB45001CB53A0016B02E000BAC
      26000888000083E5A7000000000000000000000000000000000095B2FE002F53
      F30095B2FE006F95FF003966FA003D6AFF002D5FFF002255FF00083CFE00083C
      FE000625DB0095B2FE0000000000000000000000000038383800666666008F8F
      8F00A4A4A4004D4D4D00464646000000000063ABD200247DB60057A9D70063AB
      D200479FD0001372A200699AAE0000000000000000000000000000000000FFC7
      B100FFAB8E00FFE2B400FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFE2B400CC99
      9900BFBFBF0000000000000000000000000000000000000000000000000083E5
      A70027A32C0056CA730077E09C005DD27C0043C8660036C154001EAA2F001D91
      1E0083E5A70000000000000000000000000000000000000000000000000095B2
      FE00294CF2006489FF0088A9FF006D8FFF004F7AFF003D6AFF00204CF6001D37
      DC0095B2FE0000000000000000000000000000000000000000006D6D6D00605E
      5E00605E5E00666666000000000000000000000000003A8BB70084C0E400A3D0
      EA00479FD000699AAE0000000000000000000000000000000000000000000000
      0000DCDCDC00BFBFBF00FFAB8E00CC999900DE9A3E00CC999900CC999900DCDC
      DC00000000000000000000000000000000000000000000000000000000000000
      00000000000083E5A70042AA42002BA33100279D2B0040A63F0083E5A70083E5
      A700000000000000000000000000000000000000000000000000000000000000
      00000000000095B2FE003C57E8002A4CED002A4CED004058E40095B2FE000000
      0000000000000000000000000000000000000000000000000000000000000000
      00000000000000000000000000000000000000000000000000003A8BB7003A8B
      B7003A8BB7000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000007070
      7000707070007070700070707000707070000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000CC6D6D00CC6D
      6D00CC6D6D00CC6D6D00CC6D6D00707070007070700000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      00009D4D25009D4D25009D4D25000000000000000000D0979700DDD9D900E1E0
      E000DBD2D200DDCDCD00C5ADAD00CC6D6D007070700000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      00000000000000000000000000000000000000000000000000001736A4000000
      0000000000000000000000000000000000000000000000000000000000009D4D
      2500FA9D6200E1865800E18658009D4D250000000000D0979700F3F3F300F2F2
      F200DFDEDE00DBC0C000CCA0A000CC6D6D007070700070707000707070007070
      70007373730000000000000000000000000000000000000000007C3600007C36
      00007C3600007C3600007C3600007C3600007C3600007C360000000000000000
      0000000000000000000000000000000000000000000000000000000096000000
      9600000096000000960000009600000096000000960000009600000000000000
      000000000000000000000000000000000000000000001736A4002252E5001736
      A4000000000000000000276128001A671C001A671C0000000000E79A7300FCBE
      9700FBC0A100FBB18200A2563100000000000000000000000000C0979700E0C0
      C000D9C5C500D0979700CC7A7A00BF8A8A00CC6D6D00CC6D6D00CC6D6D00CC6D
      6D007070700070707000000000000000000000000000000000007D390000E4A0
      6F00F2851D00E0771700F2851D00F4900000F49000007D390000000000000000
      0000000000000000000000000000000000000000000000000000000096006BC6
      FF006BC6FF006BC6FF0025AAFF0025AAFF000092DC0000009600000000000000
      0000000000000000000000000000000000002B4AC3002252E500668BF9001736
      A400000000002C7833005AAF5F005BA95F004D9C4F0019771C0000000000E79A
      7300E79A7300E79A730000000000000000000000000000000000AE969600D489
      8900D4898900D4898900C6606000DBC7C700E1E0E000DDD9D900DDCDCD00CCB9
      B900CC6D6D0070707000000000000000000000000000000000007D390000E4A0
      6F00E0771700E0771700E0771700E0771700F49000007D390000000000000000
      0000000000000000000000000000000000000000000000000000000096006BC6
      FF006BC6FF006BC6FF0025AAFF0025AAFF000092DC0000009600000000000000
      0000000000000000000000000000000000002580E2006D90FA00C5D3FE002853
      DF002B82390088D793007ECD880064B36900227A270000000000000000000000
      000056504A0000000000000000000000000000000000C6B69300F6B35300F6B3
      5300F6B35300F6B35300F6B35300F6B35300CB787800DFDEDE00D6B0B000CC93
      9300CC6D6D0070707000000000000000000000000000000000007D390000E4A0
      6F00FF8B5000C76B1E00E0771700E0771700F2851D007D3900007D3900007C36
      00007C3600007C36000000000000000000000000000000000000000096006BC6
      FF006BC6FF0025AAFF0025AAFF0025AAFF000092DC0000009600000096000000
      9600000096000000960000000000000000004972F400D1DBFE002551EE000000
      0000000000002B823900267E2E00227A27000000000056504A00000000005650
      4A000000000000000000000000000000000000000000C6B69300FFBC4600FFBC
      4600FFC04D00FFC04D00FFBE4900F9B14300CB787800D9C5C500D0979700CC6D
      6D00CC6D6D0070707000000000000000000000000000000000007D390000E4A0
      6F00E4A06F00FF8B5000E0771700E0771700F2851D007D390000FF8B5000D074
      0000E07717007D390000000000000000000000000000000000000000960025AA
      FF0025AAFF0025AAFF0025AAFF0025AAFF000092DC00000096006BC6FF0025AA
      FF0025AAFF00000096000000000000000000000000004972F400000000005650
      4A00000000000000000056504A0000000000000000000000000056504A000000
      00000000000000000000000000000000000000000000C6B69300FFCC6600FFCC
      6600FFCC6600FFCC6600FFCC6600FFCA6300CB787800CD848400CD848400CB78
      7800714D4D0070707000707070000000000000000000000000007D3900007D39
      00007D3900007D3900007D3900007D3900007D3900007D390000FF8B5000C76B
      1E00E07717007D39000000000000000000000000000000000000000096000000
      96000000960000009600000096000000960000009600000096006BC6FF0025AA
      FF0025AAFF000000960000000000000000000000000000000000000000000000
      000056504A00000000000000000056504A000000000056504A00000000005650
      4A00000000001736A400000000000000000000000000C6B69300F7CF8500FFD9
      8000FFD98000FFD98000FFD57900FFCC6600CB787800FFB33300FFB33300F6B3
      5300F6B35300CB78780070707000000000000000000000000000000000000000
      000000000000000000007D390000E4A06F00E4A06F00E4A06F00E4A06F00C76B
      1E00C76B1E007D39000000000000000000000000000000000000000000000000
      00000000000000000000000096006BC6FF006BC6FF006BC6FF006BC6FF0025AA
      FF0025AAFF000000960000000000000000000000000000000000000000000000
      00000000000056504A00000000000000000056504A0000000000000000000000
      00001736A4002252E5001736A4000000000000000000C6B69300F7CF8500FFE0
      8F00FFE59900FFE59900FFE59900F7CF8500CB787800FFCC6600FFCC6600FFC7
      5C00FFC75C00CB78780070707000000000000000000000000000000000000000
      000000000000000000007D390000E4A06F00FF8B5000C76B1E00C76B1E00C76B
      1E00C76B1E007D39000000000000000000000000000000000000000000000000
      00000000000000000000000096006BC6FF006BC6FF0025AAFF0025AAFF0025AA
      FF0025AAFF000000960000000000000000000000000000000000000000000000
      0000000000000000000056514D003E6780005DDDE300364B6100000000002B4A
      C3002252E500668BF9001736A4000000000000000000C6B69300F7CF8500FFE0
      8F00F4E7B100FFFCC600FFE79C00EDCA9100CB787800FFCC6600FFCC6600FFCC
      6600FFCC6600CB78780070707000000000000000000000000000000000000000
      000000000000000000007D3900007D3900007D3900007D3900007D3900007D39
      00007D3900007C36000000000000000000000000000000000000000000000000
      0000000000000000000000009600000096000000960000009600000096000000
      9600000096000000960000000000000000000000000000000000000000000000
      0000000000000000000060A3B9005AD6E3009DE2E40023688C00000000002580
      E200668BF900C5D3FE002853DF000000000000000000CBBABA00CBBABA00CBBA
      BA00CBBABA00CFA68900CFA68900CFA68900CB787800FFE59900FFE59900FFDF
      8C00FFD17000CB78780070707000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000005FA4BB007FBCF10096E1E400C7E8EA006596A800000000004972
      F400D1DBFE002556ED0000000000000000000000000000000000000000000000
      00000000000000000000C6B69300F7CF8500FFE08F00FFE89F00FFEAA300FFE5
      9900FBDA8E00CB78780070707000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      00000000000055BADB00A2E2E600BEE8EA0083B3C10000000000000000000000
      00004972F4000000000000000000000000000000000000000000000000000000
      00000000000000000000C6B69300F7CF8500FFE08F00F4E7B100FFFCC600F5DC
      9D00EDCA9100CB78780070707000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      00000000000053BFE100CFF4F60082B6C3000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      00000000000000000000CBBABA00CBBABA00CBBABA00CBBABA00CFA68900CFA6
      8900CFA68900CB78780000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000053BFE100000000000000000000000000000000000000
      000000000000000000000000000000000000424D3E000000000000003E000000
      2800000040000000500000000100010000000000800200000000000000000000
      000000000000000000000000FFFFFF0000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      000000000000000000000000000000000000FFFFFFFFFFFF0000FFFFFC0FFC07
      0000FBC3F007F0070000F9C3E003E0070000F8C3C001C0070000F9C3C001C001
      00000BFF8001800100000FFF8001800100000FFF8001800100000BC380018001
      0000F9C3C003C0030000F8C3C003C0030000F9C3E007E0070000FBFFF00FF00F
      0000FFFFFC3FFC3F0000FFFFFFFFFFFF07E0FFEFFFFFFFFF07E0FFC7FFEFFE07
      07E0FF83FFE7FC0707E0FEC1FF03E60707C08063FF01C603DF9F0037FF018E01
      DF3F007FFF038E01DE7F005FFFE78E21DCFF000F00078471D9FF000700078071
      D3E000030007807107E000470007C0630000006F0007E06707E0007F0007E07F
      07E0007F0007E07F07FF80FF0007FFFFFFFFFFFFFFFFFFFFF00FF81FF81F807F
      E007E007E0070000C003C003C00300008001C003C00300008001800180018000
      800180018001C003800180018001800380018001800100018001800180010001
      80018001800100018001C003C0030001C003C003C0038101E007E007E007C383
      F00FF80FF81FFFC7FFFFFFFFFFFFFFFFE0FFFFFFFFFFFFFFC07FFFFFFFFFFFF1
      807FFFFFFFFFDFE08007C03FC03F8C41C003C03FC03F0823C003C03FC03F0077
      8003C003C00318AF8003C003C003ADDF8001C003C003F6AB8001FC03FC03FB71
      8001FC03FC03FC218001FC03FC03FC218001FFFFFFFFF823FC01FFFFFFFFF877
      FC01FFFFFFFFF8FFFC03FFFFFFFFFDFF00000000000000000000000000000000
      000000000000}
  end
  object OpenDialog: TOpenDialog
    Filter = #32020#25991#23383#27284#26696'(*.txt)|*.txt|CSV'#27284#26696'(*.csv)|*.csv|'#25152#26377#27284#26696'(*.*)|*.*'
    Title = #38283#21855#25972#25209#21295#20837#25991#23383#27284
    Left = 184
    Top = 372
  end
  object cdsNagraChannel: TClientDataSet
    Aggregates = <>
    Params = <>
    Left = 148
    Top = 452
  end
  object cdsNagra: TClientDataSet
    Aggregates = <>
    Params = <>
    AfterScroll = cdsDVNAfterScroll
    Left = 114
    Top = 452
  end
  object dsNagra: TDataSource
    DataSet = cdsNagra
    Left = 84
    Top = 452
  end
  object qryAutoFill: TADOQuery
    Connection = DBController.DataConnection
    Parameters = <>
    Left = 184
    Top = 408
  end
end
