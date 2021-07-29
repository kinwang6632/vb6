object frmCD078B3: TfrmCD078B3
  Left = 448
  Top = 207
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = 'frmCD078B3'
  ClientHeight = 573
  ClientWidth = 789
  Color = clBtnFace
  Font.Charset = ANSI_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = 'Tahoma'
  Font.Style = []
  Icon.Data = {
    0000010002002020100000000000E80200002600000010101000000000002801
    00000E0300002800000020000000400000000100040000000000800200000000
    0000000000000000000000000000000000000000800000800000008080008000
    0000800080008080000080808000C0C0C0000000FF0000FF000000FFFF00FF00
    0000FF00FF00FFFF0000FFFFFF00000000000000000000000000000000000000
    00000000BBBBBBBB000000000000000000000BBBBBBBBBBBBBB0000000000000
    0000BBBBBBBBBBBBBBBB00000000000000BBBBBBBBBBBBBBBBBBBB0000000000
    0BBBBBBBBBBBBBBBBBBBBBB000000000BBBBBBBBB0000000BBBBBBBB00000000
    BBBBBBB00BBBBBBB00BBBBBB0000000BBBBBBB00BBBBBBBBB00BBBBBB00000BB
    BBBBB00BBBBBBBBBBB00BBBBBB0000BBBBBBB0BBBBBBBBBBBBB0BBBBBB0000BB
    BBBB0BBBBBBBBBBBBBBB0BBBBB000BBBBBBB0BBBBBBBBBBBBBBB0BBBBBB00BBB
    BBBB0BBBBBBBBBBBBBBB0BBBBBB00BBBBBBBBBBBBBBBBBBBBBBBBBBBBBB00BBB
    BBBBBBBBBBBBBBBBBBBBBBBBBBB00BBBBBBBBBBBBBBBBBBBBBBBBBBBBBB00BBB
    BBBBBBB00BBBBBB00BBBBBBBBBB00BBBBBBBBB0000BBBB0000BBBBBBBBB00BBB
    BBBBBB0000BBBB0000BBBBBBBBB000BBBBBBBB0000BBBB0000BBBBBBBB0000BB
    BBBBBB0000BBBB0000BBBBBBBB0000BBBBBBBB0000BBBB0000BBBBBBBB00000B
    BBBBBBB00BBBBBB00BBBBBBBB0000000BBBBBBBBBBBBBBBBBBBBBBBB00000000
    BBBBBBBBBBBBBBBBBBBBBBBB000000000BBBBBBBBBBBBBBBBBBBBBB000000000
    00BBBBBBBBBBBBBBBBBBBB00000000000000BBBBBBBBBBBBBBBB000000000000
    00000BBBBBBBBBBBBBB000000000000000000000BBBBBBBB0000000000000000
    0000000000000000000000000000FFF00FFFFF8001FFFE00007FFC00003FF800
    001FF000000FE0000007C0000003C00000038000000180000001800000010000
    0000000000000000000000000000000000000000000000000000000000008000
    00018000000180000001C0000003C0000003E0000007F000000FF800001FFC00
    003FFE00007FFF8001FFFFF00FFF280000001000000020000000010004000000
    0000C00000000000000000000000000000000000000000000000000080000080
    00000080800080000000800080008080000080808000C0C0C0000000FF0000FF
    000000FFFF00FF000000FF00FF00FFFF0000FFFFFF0000000000000000000000
    0BBBBBB00000000BBBBBBBBBB00000BBB000003BBB0000BB00BBBB03BB000BBB
    0BBBBBB0BBB00BBB0BBBBBB0BBB00BBBBBBBBBBBBBB00BBBB00BB00BBBB00BBB
    B00BB00BBBB00BBBB00BB00BBBB000BBB00BB00BBB0000BBBBBBBBBBBB00000B
    BBBBBBBBB00000000BBBBBB000000000000000000000F81F0000E0070000C003
    0000840100008801000000000000000000000000000000000000000000000000
    00008001000080010000C0030000E0070000F81F0000}
  OldCreateOrder = False
  Position = poScreenCenter
  OnClose = FormClose
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 14
  object Panel1: TPanel
    Left = 0
    Top = 541
    Width = 789
    Height = 32
    Align = alBottom
    BevelOuter = bvNone
    TabOrder = 0
    DesignSize = (
      789
      32)
    object edtDML: TcxTextEdit
      Left = 828
      Top = -4
      Anchors = [akRight, akBottom]
      Enabled = False
      ParentFont = False
      Properties.Alignment.Horz = taCenter
      Properties.ReadOnly = True
      Style.Color = 16777111
      Style.Font.Charset = ANSI_CHARSET
      Style.Font.Color = clWindowText
      Style.Font.Height = -13
      Style.Font.Name = 'Tahoma'
      Style.Font.Style = []
      Style.StyleController = StyleController.EditorStyle
      Style.IsFontAssigned = True
      StyleDisabled.Color = 16777111
      StyleDisabled.TextColor = clWindowText
      TabOrder = 0
      Width = 67
    end
    object btnSave: TButton
      Left = 78
      Top = 3
      Width = 69
      Height = 26
      Action = actSave
      TabOrder = 1
    end
    object Button1: TButton
      Left = 158
      Top = 3
      Width = 69
      Height = 26
      Action = actCancel
      Cancel = True
      TabOrder = 2
    end
    object btnUpdate: TButton
      Left = 6
      Top = 3
      Width = 69
      Height = 26
      Action = actUpdate
      TabOrder = 3
    end
  end
  object Panel3: TPanel
    Left = 0
    Top = 0
    Width = 789
    Height = 209
    Align = alTop
    BevelOuter = bvNone
    TabOrder = 1
    object Label3: TLabel
      Left = 7
      Top = 84
      Width = 48
      Height = 14
      Caption = #25351#23450#35373#20633
    end
    object Label18: TLabel
      Left = 7
      Top = 51
      Width = 48
      Height = 28
      Caption = #23565#25033#27491#38917#13#25910#36027#38917#30446
      Font.Charset = ANSI_CHARSET
      Font.Color = clBlue
      Font.Height = -12
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
      WordWrap = True
    end
    object Label1: TLabel
      Left = 7
      Top = 33
      Width = 48
      Height = 14
      Caption = #25910#36027#38917#30446
      Font.Charset = ANSI_CHARSET
      Font.Color = clRed
      Font.Height = -12
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
    end
    object Label2: TLabel
      Left = 7
      Top = 10
      Width = 48
      Height = 14
      Caption = #26381#21209#39006#21029
      Font.Charset = ANSI_CHARSET
      Font.Color = clRed
      Font.Height = -12
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
    end
    object Label9: TLabel
      Left = 15
      Top = 108
      Width = 40
      Height = 14
      Caption = 'CM'#36895#29575
    end
    object Bevel2: TBevel
      Left = 0
      Top = 206
      Width = 789
      Height = 3
      Align = alBottom
      Shape = bsBottomLine
    end
    object Label8: TLabel
      Left = 6
      Top = 130
      Width = 48
      Height = 14
      Caption = #40670#25976#36774#27861
      Font.Charset = ANSI_CHARSET
      Font.Color = clBlue
      Font.Height = -12
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
      WordWrap = True
    end
    object Label19: TLabel
      Left = 5
      Top = 151
      Width = 48
      Height = 28
      Caption = #29305#23450#20778#24800#25130#27490#26085
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -12
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
      WordWrap = True
    end
    object Page1: TcxPageControl
      Left = 359
      Top = 4
      Width = 425
      Height = 197
      ActivePage = cxTabSheet1
      LookAndFeel.Kind = lfFlat
      TabOrder = 6
      OnDrawTabEx = Page1DrawTabEx
      ClientRectBottom = 196
      ClientRectLeft = 1
      ClientRectRight = 424
      ClientRectTop = 22
      object cxTabSheet1: TcxTabSheet
        Caption = #32129#32004#21443#25976
        ImageIndex = 0
        object cxGroupBox1: TcxGroupBox
          Left = 14
          Top = 3
          ParentFont = False
          Style.StyleController = StyleController.EditorStyle
          Style.TextColor = clBlue
          Style.TextStyle = [fsBold]
          TabOrder = 0
          Height = 168
          Width = 406
          object Label4: TLabel
            Left = 85
            Top = 17
            Width = 48
            Height = 14
            Caption = #32129#32004#26376#25976
          end
          object Label5: TLabel
            Left = 27
            Top = 40
            Width = 84
            Height = 14
            Caption = #36949#32004#26178#35336#20729#20381#25818
          end
          object Label6: TLabel
            Left = 15
            Top = 64
            Width = 96
            Height = 14
            Caption = #20778#24800#21040#26399#35336#20729#20381#25818
          end
          object Label7: TLabel
            Left = 51
            Top = 88
            Width = 60
            Height = 14
            Caption = #20445#35657#37329#23660#24615
          end
          object lblMinCount: TLabel
            Left = 194
            Top = 18
            Width = 48
            Height = 14
            Caption = #26368#23569#24190#21488
          end
          object lblMaxCount: TLabel
            Left = 282
            Top = 18
            Width = 48
            Height = 14
            Caption = #26368#22810#24190#21488
          end
          object lblContractType: TLabel
            Left = 51
            Top = 113
            Width = 48
            Height = 14
            Caption = #32129#32004#21407#21063
          end
          object lbl1: TLabel
            Left = 168
            Top = 138
            Width = 69
            Height = 14
            Caption = #25480#27402'Tuner'#25976
          end
          object chkBundle: TcxCheckBox
            Left = 6
            Top = 14
            Caption = #26159#21542#32129#32004
            Properties.NullStyle = nssUnchecked
            Style.StyleController = StyleController.CheckBoxStyle
            TabOrder = 0
            Width = 81
          end
          object edtBundleMon: TcxMaskEdit
            Left = 139
            Top = 14
            ParentFont = False
            Properties.Alignment.Horz = taRightJustify
            Properties.MaskKind = emkRegExpr
            Properties.EditMask = '[0-9]?[0-9]?[0-9]'
            Properties.MaxLength = 0
            Style.StyleController = StyleController.EditorStyle
            TabOrder = 1
            Width = 51
          end
          object cmbPenalType: TcxComboBox
            Left = 117
            Top = 37
            ParentFont = False
            Properties.DropDownListStyle = lsEditFixedList
            Properties.Items.Strings = (
              ''
              '1.'#22238#28335#29260#20729
              '2.'#20778#24800#20729
              '4.'#25033#25910'/'#23526#25910#37329#38989)
            Style.StyleController = StyleController.EditorStyle
            TabOrder = 4
            Width = 160
          end
          object cmbExpireType: TcxComboBox
            Left = 117
            Top = 61
            ParentFont = False
            Properties.DropDownListStyle = lsEditFixedList
            Properties.Items.Strings = (
              ''
              '1.'#24674#24489#26368#26032#20844#21578#29260#20729
              '2.'#20197#26368#24460#19968#38542#32380#32396#20778#24800
              '3.'#21040#26399#19981#20986)
            Style.StyleController = StyleController.EditorStyle
            TabOrder = 5
            Width = 160
          end
          object cmbDepositAttr: TcxComboBox
            Left = 117
            Top = 83
            ParentFont = False
            Properties.DropDownListStyle = lsEditFixedList
            Properties.Items.Strings = (
              ''
              '1.'#19968#33324
              '2.'#23653#32004)
            Properties.OnChange = cmbDepositAttrPropertiesChange
            Style.StyleController = StyleController.EditorStyle
            TabOrder = 6
            Width = 160
          end
          object btnDepositAmt: TBitBtn
            Left = 15
            Top = 135
            Width = 145
            Height = 22
            Hint = #39023#31034#35373#23450#20184#27454#31278#39006#20445#35657#37329#26126#32048
            Caption = #20184#27454#31278#39006#20445#35657#37329#35373#23450
            ParentShowHint = False
            ShowHint = True
            TabOrder = 7
            OnClick = btnDepositAmtClick
            Glyph.Data = {
              36030000424D3603000000000000360000002800000010000000100000000100
              18000000000000030000120B0000120B00000000000000000000FF00FFFF00FF
              FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
              FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF008ECA0085B7007EB300
              7AAD0076A90075A80076A90076ACFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
              FF00FF00A2DD0090C30090C30090C30098CB0091C40071A4006EA10072A60084
              BBFF00FFFF00FFFF00FFFF00FFFF00FF00AAE6019DD100A1D400A0D300A1D636
              D0F8A1EBFE009ED30082B5006B9E0071A40084BEFF00FFFF00FFFF00FF00A6DE
              10A8D800AADD00A9DC00AFE370D8F4CFF2FBD8F5FD7AE3FC00C0F4008FC2006B
              9E0074A8007BB7FF00FFFF00FF00A3D526BCE600B2E500ABE1D4EDF6FFFFFDFF
              FDFDEDF9FCFFFFFFD7F7FF00B6EB0082B5006FA20079ADFF00FFFF00FF17B3E0
              25C5F000BBEF0CADDDCAE7F2DCEEF451C9ED4ACBF0CAEEF8FFFFFE1FC9F60096
              C90071A40079ABFF00FFFF00FF25BDE82BCEF800C4F800B8EB00A6DE00A9DF92
              D8EEEEF6F9FFFCFBFFFFFC06BCEC009BCE0082B5007BAEFF00FFFF00FF26C1ED
              41D9FE00D0FF00BFF157BDDFFFFBF8FFFAF8FFFBF9F6F4F54CCAEE00A9DE00A1
              D4008BBE0080B3FF00FFFF00FF17BEED68E9FF00DDFF00A9DCEDE9ECFFF7F5C7
              E2EC70CCE800A4DF00ADE400AEE200A4D70090C30085B9FF00FFFF00FF00B5E8
              67EAFD2DEEFF00ACDBD9DDE6FFF6F39CD2E6AFDAE9FFF6F2CFE5EF00AEE200A7
              DA0093C6008CC0FF00FFFF00FF00C4FC24CCF475FCFF0CE2F51E9ECDE9E0E5FF
              F1ECFFF2EEF2E9EB39BAE100AFE200A8DB0095C9009ED0FF00FFFF00FFFF00FF
              00D1FF39DFF973FFFF25E5F800AFDE30A6D175C6E200ACE200B7EB00B2E500A3
              D600ACE7FF00FFFF00FFFF00FFFF00FFFF00FF00D2FF21D1F561F3FE60F4FF34
              D4F51ECAF619CCFA17C0EE05ABDE00B5EEFF00FFFF00FFFF00FFFF00FFFF00FF
              FF00FFFF00FF00C3F900B7EC13C1EF20C5EF20C1ED13B7E700A9DE00AAE1FF00
              FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
              00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF}
          end
          object edtMinCount: TcxMaskEdit
            Left = 243
            Top = 14
            ParentFont = False
            Properties.Alignment.Horz = taRightJustify
            Properties.MaskKind = emkRegExpr
            Properties.EditMask = '[0-9]?[0-9]?[0-9]'
            Properties.MaxLength = 0
            Style.StyleController = StyleController.EditorStyle
            TabOrder = 2
            Text = '0'
            Width = 34
          end
          object edtMaxCount: TcxMaskEdit
            Left = 331
            Top = 14
            ParentFont = False
            Properties.Alignment.Horz = taRightJustify
            Properties.MaskKind = emkRegExpr
            Properties.EditMask = '[0-9]?[0-9]?[0-9]'
            Properties.MaxLength = 0
            Style.StyleController = StyleController.EditorStyle
            TabOrder = 3
            Text = '0'
            Width = 34
          end
          object cmbContractType: TcxComboBox
            Left = 117
            Top = 108
            ParentFont = False
            Properties.DropDownListStyle = lsEditFixedList
            Properties.Items.Strings = (
              ''
              '1.'#25563#32004'('#21512#32004#37325#32129')'
              '2.'#25563#26041#26696'('#21512#32004#19981#37325#32129)
            Properties.OnChange = cmbDepositAttrPropertiesChange
            Style.StyleController = StyleController.EditorStyle
            TabOrder = 8
            Width = 160
          end
          object edtAuthTunerCount: TcxMaskEdit
            Left = 241
            Top = 134
            ParentFont = False
            Properties.Alignment.Horz = taRightJustify
            Properties.MaskKind = emkRegExpr
            Properties.EditMask = '[0-9]'
            Properties.MaxLength = 0
            Style.StyleController = StyleController.EditorStyle
            TabOrder = 9
            Text = '0'
            Width = 34
          end
        end
      end
      object cxTabSheet2: TcxTabSheet
        Caption = #37559#21806#29260#20729
        ImageIndex = 1
        object cxGroupBox2: TcxGroupBox
          Left = 9
          Top = 0
          ParentFont = False
          Style.StyleController = StyleController.EditorStyle
          Style.TextColor = clBlue
          Style.TextStyle = [fsBold]
          TabOrder = 0
          Height = 168
          Width = 399
          object Label10: TLabel
            Left = 72
            Top = 20
            Width = 48
            Height = 14
            Caption = #21934#26376#37329#38989
          end
          object Label11: TLabel
            Left = 232
            Top = 19
            Width = 48
            Height = 14
            Caption = #21934#26085#37329#38989
          end
          object Label12: TLabel
            Left = 26
            Top = 43
            Width = 94
            Height = 14
            Caption = #30772#26376#26399#25976'('#27966#24037#21934')'
          end
          object Label13: TLabel
            Left = 60
            Top = 93
            Width = 60
            Height = 14
            Caption = #30772#26376#22522#28310#26085
          end
          object Label14: TLabel
            Left = 172
            Top = 92
            Width = 48
            Height = 14
            Caption = #37329#38989#27512#25972
          end
          object Label15: TLabel
            Left = 12
            Top = 117
            Width = 108
            Height = 14
            Caption = #39318#26399#19981#36275#26376#25910#36027#38917#30446
          end
          object Label16: TLabel
            Left = 271
            Top = 43
            Width = 117
            Height = 53
            AutoSize = False
            Caption = '( '#22823#26044#30772#26376#22522#28310#26085#21063#13'   1. '#22823#26044#26399#25976', '#21453#20043#13'   2. '#23567#26044#26399#25976' )'
          end
          object Label20: TLabel
            Left = 26
            Top = 67
            Width = 94
            Height = 14
            Caption = #30772#26376#26399#25976'('#25910#36027#21934')'
          end
          object edtMonthAmt: TcxCurrencyEdit
            Left = 126
            Top = 15
            EditValue = 0.000000000000000000
            ParentFont = False
            Properties.Alignment.Horz = taRightJustify
            Properties.DecimalPlaces = 3
            Properties.DisplayFormat = ',0.000;-,0.000'
            Properties.MaxValue = 99999999.999000000000000000
            Properties.NullString = '0'
            Properties.OnEditValueChanged = edtMonthAmtPropertiesEditValueChanged
            Style.StyleController = StyleController.EditorStyle
            TabOrder = 0
            Width = 80
          end
          object edtDayAmt: TcxCurrencyEdit
            Left = 286
            Top = 15
            EditValue = 0.000000000000000000
            ParentFont = False
            Properties.Alignment.Horz = taRightJustify
            Properties.DecimalPlaces = 3
            Properties.DisplayFormat = ',0.000;-,0.000'
            Properties.MaxValue = 9999.999000000000000000
            Properties.NullString = '0'
            Style.StyleController = StyleController.EditorStyle
            TabOrder = 1
            Width = 80
          end
          object cmbPFlag1: TcxComboBox
            Left = 125
            Top = 40
            ParentFont = False
            Properties.DropDownListStyle = lsEditFixedList
            Properties.Items.Strings = (
              ''
              '0.'#31561#26044#26399#25976
              '1.'#23567#26044#26399#25976
              '2.'#22823#26044#26399#25976
              '3.'#30772#26376#22522#28310#26085)
            Properties.OnChange = cmbPFlag1PropertiesChange
            Style.StyleController = StyleController.EditorStyle
            TabOrder = 2
            Width = 140
          end
          object edtPFlagDate: TcxMaskEdit
            Left = 125
            Top = 89
            ParentFont = False
            Properties.Alignment.Horz = taRightJustify
            Properties.MaskKind = emkRegExpr
            Properties.EditMask = '[0-9]?[0-9]'
            Properties.MaxLength = 0
            Style.StyleController = StyleController.EditorStyle
            TabOrder = 4
            Text = '1'
            Width = 40
          end
          inline PFlag: TDataLookup
            Left = 124
            Top = 113
            Width = 262
            Height = 23
            HorzScrollBar.Visible = False
            VertScrollBar.Visible = False
            TabOrder = 6
            inherited CodeNo: TcxTextEdit
              Left = 1
              Properties.MaxLength = 5
            end
            inherited CodeName: TcxLookupComboBox
              Properties.OnInitPopup = PFlagCodeNamePropertiesInitPopup
              Width = 197
            end
          end
          object edtTruncAmt: TcxCurrencyEdit
            Left = 225
            Top = 88
            EditValue = 0.000000000000000000
            ParentFont = False
            Properties.Alignment.Horz = taRightJustify
            Properties.DecimalPlaces = 0
            Properties.DisplayFormat = ',0;-,0'
            Properties.MaxValue = 9999.000000000000000000
            Properties.NullString = '0'
            Style.StyleController = StyleController.EditorStyle
            TabOrder = 5
            Width = 40
          end
          object btnInstCodeStr: TButton
            Left = 8
            Top = 139
            Width = 89
            Height = 25
            Caption = #25351#23450#27966#24037#39006#21029
            TabOrder = 7
            OnClick = btnInstCodeStrClick
          end
          object cxButton1: TcxButton
            Left = 97
            Top = 139
            Width = 25
            Height = 25
            TabOrder = 8
            OnClick = cxButton1Click
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
            LookAndFeel.Kind = lfStandard
          end
          object edtInstCodeStr: TcxTextEdit
            Left = 126
            Top = 140
            Enabled = False
            ParentFont = False
            Style.StyleController = StyleController.EditorStyle
            TabOrder = 9
            Width = 258
          end
          object cmbPFlag2: TcxComboBox
            Left = 125
            Top = 64
            ParentFont = False
            Properties.DropDownListStyle = lsEditFixedList
            Properties.Items.Strings = (
              '0.'#31561#26044#26399#25976
              '2.'#22823#26044#26399#25976)
            Properties.OnChange = cmbPFlag1PropertiesChange
            Style.StyleController = StyleController.EditorStyle
            TabOrder = 3
            Text = '0.'#31561#26044#26399#25976
            Width = 140
          end
        end
      end
    end
    inline FaciItem: TDataLookup
      Left = 58
      Top = 79
      Width = 211
      Height = 23
      HorzScrollBar.Visible = False
      VertScrollBar.Visible = False
      TabOrder = 3
      inherited CodeNo: TcxTextEdit
        Left = 1
        Properties.MaxLength = 2
        Properties.OnChange = FaciItemCodeNoPropertiesChange
      end
      inherited CodeName: TcxLookupComboBox
        Properties.OnChange = FaciItemCodeNamePropertiesChange
        Properties.OnInitPopup = FaciItemCodeNamePropertiesInitPopup
        Width = 148
      end
    end
    inline MCItem: TDataLookup
      Left = 59
      Top = 55
      Width = 300
      Height = 23
      HorzScrollBar.Visible = False
      VertScrollBar.Visible = False
      TabOrder = 2
      inherited CodeNo: TcxTextEdit
        Properties.OnChange = MCItemCodeNoPropertiesChange
      end
      inherited CodeName: TcxLookupComboBox
        Left = 61
        Properties.OnChange = MCItemCodeNamePropertiesChange
        Properties.OnInitPopup = MCItemCodeNamePropertiesInitPopup
        Width = 235
      end
    end
    inline CItem: TDataLookup
      Left = 59
      Top = 30
      Width = 300
      Height = 23
      HorzScrollBar.Visible = False
      VertScrollBar.Visible = False
      TabOrder = 1
      inherited CodeNo: TcxTextEdit
        Properties.MaxLength = 5
        Properties.OnChange = CItemCodeNoPropertiesChange
        OnExit = CItemCodeNoExit
      end
      inherited CodeName: TcxLookupComboBox
        Left = 61
        Properties.OnChange = CItemCodeNamePropertiesChange
        Properties.OnInitPopup = CItemCodeNamePropertiesInitPopup
        Width = 235
      end
    end
    inline ServiceType: TDataLookup
      Left = 59
      Top = 6
      Width = 175
      Height = 23
      HorzScrollBar.Visible = False
      VertScrollBar.Visible = False
      TabOrder = 0
      inherited CodeNo: TcxTextEdit
        Properties.CharCase = ecUpperCase
        Properties.MaxLength = 1
        Properties.OnChange = ServiceTypeCodeNoPropertiesChange
      end
      inherited CodeName: TcxLookupComboBox
        Left = 61
        Properties.OnChange = ServiceTypeCodeNamePropertiesChange
        Properties.OnInitPopup = ServiceTypeCodeNamePropertiesInitPopup
        Width = 110
      end
    end
    object btnChangeCItemCode: TBitBtn
      Left = 238
      Top = 6
      Width = 116
      Height = 22
      Hint = #35722#26356#25910#36027#38917#30446
      Caption = #35722#26356#25910#36027#38917#30446
      ParentShowHint = False
      ShowHint = True
      TabOrder = 7
      OnClick = btnChangeCItemCodeClick
      Glyph.Data = {
        36040000424D3604000000000000360000002800000010000000100000000100
        20000000000000040000120B0000120B00000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000DD7A003BCE7B0EC1D88A28F5DB84179E000000000000000000000000C22A
        0006D37B1C8200000000000000000000000000000000DF6B000BD2720A95DC83
        16DEE1943AFEFFD0A0FFFFDBB5FFDF8817DA000000000000000000000000C06A
        1386CA6B119A000000000000000000000000DA730A93DA8325F1EAA04CFFFFC9
        8EFFFFC88EFFFFC78FFFFFCFA0FFE28F26F5C971094F0000000000000000A845
        009BC96A11980000000000000000ED580007E4943DFCFFE2B6FFFFD19DFFFFCD
        99FFFFCC9AFFFFC48BFFFFCB95FFE9A050FECC7208810000000000000000BD59
        0AD1C04F0072000000000000000000000000DA720C77CF6900A0BB5900AAD786
        28FFFFC587FFFFC286FFFFC78EFFF7BA79FFC8720AB10000000000000000D068
        10F1BD50008D0000000000000000000000000000000000000000CB6E09AEFFBB
        77FFFFC282FFFEBE7DFFFFCE9BFFFFCF98FFC97109CB0000000000000000CF67
        0EF6D36408C900000000000000000000000000000000D561002BDB872BF9FFBF
        79FFFFC483FFD9892BF2DC913BF5FFD8AAFFD07E1CDB0000000000000000C35C
        0AE0ED8017FEBE48004D0000000000000000D05F001ECB7113E5FFBB6EFFFFBB
        71FFFCB86FFFC87214B6BE5E007AFFD7AAFFD7882EF80000000000000000A641
        00A6FFA223FFDE7312FAC35F07BEC35E08B3D0751AF3FFB35CFFFFB262FFFFBD
        73FFCC7416D6D4660008B9590006DD7A0DBED6750A930000000000000000A93D
        003EE47B1CFEFF9F2AFFFF9F2DFFFFA740FFFFAA4BFFFFAB50FFFFC986FFD67C
        26F1CE6B05500000000000000000000000000000000000000000000000000000
        0000BF47008EEE8E30FEFFC679FFFFC37CFFFFCB8BFFFFC886FFCE701CEBCC65
        045C000000000000000000000000000000000000000000000000000000000000
        000000000000B6540C79B24B00CAC86413F6C2600EE6B25408A9C56406400000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000}
    end
    object edtCMBaudNo: TcxTextEdit
      Tag = 1
      Left = 59
      Top = 104
      ParentFont = False
      Properties.ReadOnly = True
      Style.StyleController = StyleController.EditorStyle
      TabOrder = 4
      Width = 60
    end
    object edtCMBaudRate: TcxTextEdit
      Tag = 1
      Left = 120
      Top = 104
      ParentFont = False
      Properties.ReadOnly = True
      Style.StyleController = StyleController.EditorStyle
      TabOrder = 5
      Width = 148
    end
    inline MSalePoint: TDataLookup
      Left = 58
      Top = 128
      Width = 300
      Height = 23
      HorzScrollBar.Visible = False
      VertScrollBar.Visible = False
      TabOrder = 8
      inherited CodeNo: TcxTextEdit
        Properties.MaxLength = 5
        Properties.OnChange = MSalePointCodeNoPropertiesChange
      end
      inherited CodeName: TcxLookupComboBox
        Left = 61
        Properties.OnChange = MSalePointCodeNamePropertiesChange
        Properties.OnInitPopup = MSalePointCodeNamePropertiesInitPopup
        Width = 236
      end
    end
    object edtBpStopDate: TMaskEdit
      Left = 57
      Top = 154
      Width = 83
      Height = 22
      EditMask = '!9999/99/99;0;_'
      MaxLength = 10
      TabOrder = 9
      OnExit = edtBpStopDateExit
    end
    object chkStopFlag: TcxCheckBox
      Left = 144
      Top = 154
      Caption = #26159#21542#20572#29992
      Properties.NullStyle = nssUnchecked
      Style.StyleController = StyleController.CheckBoxStyle
      TabOrder = 10
      Width = 81
    end
    object chkLONGPAYflag: TcxCheckBox
      Left = 224
      Top = 154
      Caption = #38263#32371#21029#35387#35352
      Properties.NullStyle = nssUnchecked
      Style.StyleController = StyleController.CheckBoxStyle
      TabOrder = 11
      Width = 91
    end
    object chkPeriodFlag: TcxCheckBox
      Left = 56
      Top = 178
      Caption = #25351#23450#32371#21029#20986#21934
      Properties.NullStyle = nssUnchecked
      Style.StyleController = StyleController.CheckBoxStyle
      TabOrder = 12
      Width = 103
    end
  end
  object Panel4: TPanel
    Left = 0
    Top = 209
    Width = 789
    Height = 332
    Align = alClient
    BevelOuter = bvNone
    TabOrder = 2
    object Bevel1: TBevel
      Left = 0
      Top = 329
      Width = 789
      Height = 3
      Align = alBottom
      Shape = bsBottomLine
    end
    object cxGroupBox4: TcxGroupBox
      Left = 320
      Top = 2
      Caption = #38542#27573#24335#20778#24800#21443#25976
      ParentFont = False
      Style.StyleController = StyleController.EditorStyle
      Style.TextColor = clBlue
      Style.TextStyle = [fsBold]
      TabOrder = 1
      Height = 325
      Width = 466
      object pnlRank: TPanel
        Left = 2
        Top = 38
        Width = 462
        Height = 285
        Align = alCustom
        BevelOuter = bvNone
        TabOrder = 0
        object Label23: TLabel
          Left = 31
          Top = 29
          Width = 48
          Height = 14
          Caption = #20351#29992#26376#25976
        end
        object Label24: TLabel
          Left = 15
          Top = 47
          Width = 16
          Height = 14
          Caption = #19968'.'
        end
        object Label26: TLabel
          Left = 15
          Top = 68
          Width = 16
          Height = 14
          Caption = #20108'.'
        end
        object lblMon1: TLabel
          Left = 62
          Top = 48
          Width = 68
          Height = 17
          Alignment = taCenter
          AutoSize = False
          Caption = '1~12'
        end
        object lblMon2: TLabel
          Left = 62
          Top = 68
          Width = 68
          Height = 17
          Alignment = taCenter
          AutoSize = False
          Caption = '13~24'
        end
        object Label32: TLabel
          Left = 15
          Top = 88
          Width = 16
          Height = 14
          Caption = #19977'.'
        end
        object lblMon3: TLabel
          Left = 62
          Top = 89
          Width = 68
          Height = 17
          Alignment = taCenter
          AutoSize = False
          Caption = '13~24'
        end
        object Label35: TLabel
          Left = 15
          Top = 109
          Width = 16
          Height = 14
          Caption = #22235'.'
        end
        object Label37: TLabel
          Left = 15
          Top = 129
          Width = 16
          Height = 14
          Caption = #20116'.'
        end
        object lblMon4: TLabel
          Left = 62
          Top = 110
          Width = 68
          Height = 17
          Alignment = taCenter
          AutoSize = False
          Caption = '1~12'
        end
        object lblMon5: TLabel
          Left = 62
          Top = 130
          Width = 68
          Height = 17
          Alignment = taCenter
          AutoSize = False
          Caption = '13~24'
        end
        object Label43: TLabel
          Left = 15
          Top = 149
          Width = 16
          Height = 14
          Caption = #20845'.'
        end
        object lblMon6: TLabel
          Left = 62
          Top = 149
          Width = 68
          Height = 17
          Alignment = taCenter
          AutoSize = False
          Caption = '13~24'
        end
        object Label45: TLabel
          Left = 15
          Top = 170
          Width = 16
          Height = 14
          Caption = #19971'.'
        end
        object Label46: TLabel
          Left = 15
          Top = 189
          Width = 16
          Height = 14
          Caption = #20843'.'
        end
        object lblMon7: TLabel
          Left = 62
          Top = 169
          Width = 68
          Height = 17
          Alignment = taCenter
          AutoSize = False
          Caption = '1~12'
        end
        object lblMon8: TLabel
          Left = 62
          Top = 189
          Width = 68
          Height = 17
          Alignment = taCenter
          AutoSize = False
          Caption = '13~24'
        end
        object Label49: TLabel
          Left = 15
          Top = 209
          Width = 16
          Height = 14
          Caption = #20061'.'
        end
        object lblMon9: TLabel
          Left = 62
          Top = 209
          Width = 68
          Height = 17
          Alignment = taCenter
          AutoSize = False
          Caption = '13~24'
        end
        object Label51: TLabel
          Left = 15
          Top = 229
          Width = 16
          Height = 14
          Caption = #21313'.'
        end
        object Label52: TLabel
          Left = 4
          Top = 248
          Width = 28
          Height = 14
          Caption = #21313#19968'.'
        end
        object lblMon10: TLabel
          Left = 62
          Top = 230
          Width = 68
          Height = 17
          Alignment = taCenter
          AutoSize = False
          Caption = '1~12'
        end
        object lblMon11: TLabel
          Left = 62
          Top = 249
          Width = 68
          Height = 17
          Alignment = taCenter
          AutoSize = False
          Caption = '13~24'
        end
        object Label55: TLabel
          Left = 4
          Top = 270
          Width = 28
          Height = 14
          Caption = #21313#20108'.'
        end
        object lblMon12: TLabel
          Left = 63
          Top = 270
          Width = 68
          Height = 17
          Alignment = taCenter
          AutoSize = False
          Caption = '999~999'
        end
        object Label57: TLabel
          Left = 117
          Top = 29
          Width = 48
          Height = 14
          Caption = #32371#36027#26399#25976
        end
        object Label58: TLabel
          Left = 188
          Top = 29
          Width = 48
          Height = 14
          Caption = #36027#29575#20381#25818
        end
        object Label61: TLabel
          Left = 266
          Top = 29
          Width = 48
          Height = 14
          Caption = #20778#24800#37329#38989
        end
        object Label62: TLabel
          Left = 324
          Top = 29
          Width = 48
          Height = 14
          Caption = #21934#26376#37329#38989
        end
        object Label63: TLabel
          Left = 379
          Top = 29
          Width = 48
          Height = 14
          Caption = #21934#26085#37329#38989
        end
        object Label30: TLabel
          Tag = 1
          Left = 48
          Top = 7
          Width = 48
          Height = 14
          Caption = #22810#38542#26041#26696
        end
        object edtMon1: TcxMaskEdit
          Left = 33
          Top = 46
          AutoSize = False
          ParentFont = False
          Properties.Alignment.Horz = taRightJustify
          Properties.MaskKind = emkRegExpr
          Properties.EditMask = '[0-9]?[0-9]?[0-9]'
          Properties.MaxLength = 0
          Style.StyleController = StyleController.EditorStyle
          TabOrder = 0
          Height = 20
          Width = 35
        end
        object edtMon2: TcxMaskEdit
          Left = 33
          Top = 66
          AutoSize = False
          ParentFont = False
          Properties.Alignment.Horz = taRightJustify
          Properties.MaskKind = emkRegExpr
          Properties.EditMask = '[0-9]?[0-9]?[0-9]'
          Properties.MaxLength = 0
          Style.StyleController = StyleController.EditorStyle
          TabOrder = 1
          Height = 20
          Width = 35
        end
        object edtMon3: TcxMaskEdit
          Left = 33
          Top = 86
          AutoSize = False
          ParentFont = False
          Properties.Alignment.Horz = taRightJustify
          Properties.MaskKind = emkRegExpr
          Properties.EditMask = '[0-9]?[0-9]?[0-9]'
          Properties.MaxLength = 0
          Style.StyleController = StyleController.EditorStyle
          TabOrder = 2
          Height = 20
          Width = 35
        end
        object edtMon4: TcxMaskEdit
          Left = 33
          Top = 106
          AutoSize = False
          ParentFont = False
          Properties.Alignment.Horz = taRightJustify
          Properties.MaskKind = emkRegExpr
          Properties.EditMask = '[0-9]?[0-9]?[0-9]'
          Properties.MaxLength = 0
          Style.StyleController = StyleController.EditorStyle
          TabOrder = 3
          Height = 20
          Width = 35
        end
        object edtMon5: TcxMaskEdit
          Left = 33
          Top = 126
          AutoSize = False
          ParentFont = False
          Properties.Alignment.Horz = taRightJustify
          Properties.MaskKind = emkRegExpr
          Properties.EditMask = '[0-9]?[0-9]?[0-9]'
          Properties.MaxLength = 0
          Style.StyleController = StyleController.EditorStyle
          TabOrder = 4
          Height = 20
          Width = 35
        end
        object edtMon6: TcxMaskEdit
          Left = 33
          Top = 146
          AutoSize = False
          ParentFont = False
          Properties.Alignment.Horz = taRightJustify
          Properties.MaskKind = emkRegExpr
          Properties.EditMask = '[0-9]?[0-9]?[0-9]'
          Properties.MaxLength = 0
          Style.StyleController = StyleController.EditorStyle
          TabOrder = 5
          Height = 20
          Width = 35
        end
        object edtMon7: TcxMaskEdit
          Left = 33
          Top = 166
          AutoSize = False
          ParentFont = False
          Properties.Alignment.Horz = taRightJustify
          Properties.MaskKind = emkRegExpr
          Properties.EditMask = '[0-9]?[0-9]?[0-9]'
          Properties.MaxLength = 0
          Style.StyleController = StyleController.EditorStyle
          TabOrder = 6
          Height = 20
          Width = 35
        end
        object edtMon8: TcxMaskEdit
          Left = 33
          Top = 186
          AutoSize = False
          ParentFont = False
          Properties.Alignment.Horz = taRightJustify
          Properties.MaskKind = emkRegExpr
          Properties.EditMask = '[0-9]?[0-9]?[0-9]'
          Properties.MaxLength = 0
          Style.StyleController = StyleController.EditorStyle
          TabOrder = 7
          Height = 20
          Width = 35
        end
        object edtMon9: TcxMaskEdit
          Left = 33
          Top = 206
          AutoSize = False
          ParentFont = False
          Properties.Alignment.Horz = taRightJustify
          Properties.MaskKind = emkRegExpr
          Properties.EditMask = '[0-9]?[0-9]?[0-9]'
          Properties.MaxLength = 0
          Style.StyleController = StyleController.EditorStyle
          TabOrder = 8
          Height = 20
          Width = 35
        end
        object edtMon10: TcxMaskEdit
          Left = 33
          Top = 226
          AutoSize = False
          ParentFont = False
          Properties.Alignment.Horz = taRightJustify
          Properties.MaskKind = emkRegExpr
          Properties.EditMask = '[0-9]?[0-9]?[0-9]'
          Properties.MaxLength = 0
          Style.StyleController = StyleController.EditorStyle
          TabOrder = 9
          Height = 20
          Width = 35
        end
        object edtMon11: TcxMaskEdit
          Left = 33
          Top = 246
          AutoSize = False
          ParentFont = False
          Properties.Alignment.Horz = taRightJustify
          Properties.MaskKind = emkRegExpr
          Properties.EditMask = '[0-9]?[0-9]?[0-9]'
          Properties.MaxLength = 0
          Style.StyleController = StyleController.EditorStyle
          TabOrder = 10
          Height = 20
          Width = 35
        end
        object edtMon12: TcxMaskEdit
          Left = 33
          Top = 266
          AutoSize = False
          ParentFont = False
          Properties.Alignment.Horz = taRightJustify
          Properties.MaskKind = emkRegExpr
          Properties.EditMask = '[0-9]?[0-9]?[0-9]'
          Properties.MaxLength = 0
          Style.StyleController = StyleController.EditorStyle
          TabOrder = 11
          Height = 20
          Width = 35
        end
        object edtPeriod1: TcxMaskEdit
          Left = 123
          Top = 46
          AutoSize = False
          ParentFont = False
          Properties.Alignment.Horz = taRightJustify
          Properties.MaskKind = emkRegExpr
          Properties.EditMask = '[0-9]?[0-9]'
          Properties.MaxLength = 0
          Style.StyleController = StyleController.EditorStyle
          TabOrder = 12
          Height = 20
          Width = 35
        end
        object edtPeriod2: TcxMaskEdit
          Left = 123
          Top = 66
          AutoSize = False
          ParentFont = False
          Properties.Alignment.Horz = taRightJustify
          Properties.MaskKind = emkRegExpr
          Properties.EditMask = '[0-9]?[0-9]'
          Properties.MaxLength = 0
          Style.StyleController = StyleController.EditorStyle
          TabOrder = 13
          Height = 20
          Width = 35
        end
        object edtPeriod3: TcxMaskEdit
          Left = 123
          Top = 86
          AutoSize = False
          ParentFont = False
          Properties.Alignment.Horz = taRightJustify
          Properties.MaskKind = emkRegExpr
          Properties.EditMask = '[0-9]?[0-9]'
          Properties.MaxLength = 0
          Style.StyleController = StyleController.EditorStyle
          TabOrder = 14
          Height = 20
          Width = 35
        end
        object edtPeriod4: TcxMaskEdit
          Left = 123
          Top = 106
          AutoSize = False
          ParentFont = False
          Properties.Alignment.Horz = taRightJustify
          Properties.MaskKind = emkRegExpr
          Properties.EditMask = '[0-9]?[0-9]'
          Properties.MaxLength = 0
          Style.StyleController = StyleController.EditorStyle
          TabOrder = 15
          Height = 20
          Width = 35
        end
        object edtPeriod5: TcxMaskEdit
          Left = 123
          Top = 126
          AutoSize = False
          ParentFont = False
          Properties.Alignment.Horz = taRightJustify
          Properties.MaskKind = emkRegExpr
          Properties.EditMask = '[0-9]?[0-9]'
          Properties.MaxLength = 0
          Style.StyleController = StyleController.EditorStyle
          TabOrder = 16
          Height = 20
          Width = 35
        end
        object edtPeriod6: TcxMaskEdit
          Left = 123
          Top = 146
          AutoSize = False
          ParentFont = False
          Properties.Alignment.Horz = taRightJustify
          Properties.MaskKind = emkRegExpr
          Properties.EditMask = '[0-9]?[0-9]'
          Properties.MaxLength = 0
          Style.StyleController = StyleController.EditorStyle
          TabOrder = 17
          Height = 20
          Width = 35
        end
        object edtPeriod7: TcxMaskEdit
          Left = 123
          Top = 166
          AutoSize = False
          ParentFont = False
          Properties.Alignment.Horz = taRightJustify
          Properties.MaskKind = emkRegExpr
          Properties.EditMask = '[0-9]?[0-9]'
          Properties.MaxLength = 0
          Style.StyleController = StyleController.EditorStyle
          TabOrder = 18
          Height = 20
          Width = 35
        end
        object edtPeriod8: TcxMaskEdit
          Left = 123
          Top = 186
          AutoSize = False
          ParentFont = False
          Properties.Alignment.Horz = taRightJustify
          Properties.MaskKind = emkRegExpr
          Properties.EditMask = '[0-9]?[0-9]'
          Properties.MaxLength = 0
          Style.StyleController = StyleController.EditorStyle
          TabOrder = 19
          Height = 20
          Width = 35
        end
        object edtPeriod9: TcxMaskEdit
          Left = 123
          Top = 206
          AutoSize = False
          ParentFont = False
          Properties.Alignment.Horz = taRightJustify
          Properties.MaskKind = emkRegExpr
          Properties.EditMask = '[0-9]?[0-9]'
          Properties.MaxLength = 0
          Style.StyleController = StyleController.EditorStyle
          TabOrder = 20
          Height = 20
          Width = 35
        end
        object edtPeriod10: TcxMaskEdit
          Left = 123
          Top = 226
          AutoSize = False
          ParentFont = False
          Properties.Alignment.Horz = taRightJustify
          Properties.MaskKind = emkRegExpr
          Properties.EditMask = '[0-9]?[0-9]'
          Properties.MaxLength = 0
          Style.StyleController = StyleController.EditorStyle
          TabOrder = 21
          Height = 20
          Width = 35
        end
        object edtPeriod11: TcxMaskEdit
          Left = 123
          Top = 246
          AutoSize = False
          ParentFont = False
          Properties.Alignment.Horz = taRightJustify
          Properties.MaskKind = emkRegExpr
          Properties.EditMask = '[0-9]?[0-9]'
          Properties.MaxLength = 0
          Style.StyleController = StyleController.EditorStyle
          TabOrder = 22
          Height = 20
          Width = 35
        end
        object edtPeriod12: TcxMaskEdit
          Left = 123
          Top = 266
          AutoSize = False
          ParentFont = False
          Properties.Alignment.Horz = taRightJustify
          Properties.MaskKind = emkRegExpr
          Properties.EditMask = '[0-9]?[0-9]'
          Properties.MaxLength = 0
          Style.StyleController = StyleController.EditorStyle
          TabOrder = 23
          Height = 20
          Width = 35
        end
        object cmbRateType1: TcxComboBox
          Left = 158
          Top = 46
          AutoSize = False
          ParentFont = False
          Properties.DropDownListStyle = lsEditFixedList
          Properties.Items.Strings = (
            ''
            '1.'#20381#36027#29575#34920
            '2.'#20381#20778#24800#37329#38989)
          Style.StyleController = StyleController.EditorStyle
          TabOrder = 24
          Height = 20
          Width = 105
        end
        object cmbRateType2: TcxComboBox
          Left = 158
          Top = 66
          AutoSize = False
          ParentFont = False
          Properties.DropDownListStyle = lsEditFixedList
          Properties.Items.Strings = (
            ''
            '1.'#20381#36027#29575#34920
            '2.'#20381#20778#24800#37329#38989)
          Style.StyleController = StyleController.EditorStyle
          TabOrder = 25
          Height = 20
          Width = 105
        end
        object cmbRateType3: TcxComboBox
          Left = 158
          Top = 86
          AutoSize = False
          ParentFont = False
          Properties.DropDownListStyle = lsEditFixedList
          Properties.Items.Strings = (
            ''
            '1.'#20381#36027#29575#34920
            '2.'#20381#20778#24800#37329#38989)
          Style.StyleController = StyleController.EditorStyle
          TabOrder = 26
          Height = 20
          Width = 105
        end
        object cmbRateType4: TcxComboBox
          Left = 158
          Top = 106
          AutoSize = False
          ParentFont = False
          Properties.DropDownListStyle = lsEditFixedList
          Properties.Items.Strings = (
            ''
            '1.'#20381#36027#29575#34920
            '2.'#20381#20778#24800#37329#38989)
          Style.StyleController = StyleController.EditorStyle
          TabOrder = 27
          Height = 20
          Width = 105
        end
        object cmbRateType5: TcxComboBox
          Left = 158
          Top = 126
          AutoSize = False
          ParentFont = False
          Properties.DropDownListStyle = lsEditFixedList
          Properties.Items.Strings = (
            ''
            '1.'#20381#36027#29575#34920
            '2.'#20381#20778#24800#37329#38989)
          Style.StyleController = StyleController.EditorStyle
          TabOrder = 28
          Height = 20
          Width = 105
        end
        object cmbRateType6: TcxComboBox
          Left = 158
          Top = 146
          AutoSize = False
          ParentFont = False
          Properties.DropDownListStyle = lsEditFixedList
          Properties.Items.Strings = (
            ''
            '1.'#20381#36027#29575#34920
            '2.'#20381#20778#24800#37329#38989)
          Style.StyleController = StyleController.EditorStyle
          TabOrder = 29
          Height = 20
          Width = 105
        end
        object cmbRateType7: TcxComboBox
          Left = 158
          Top = 166
          AutoSize = False
          ParentFont = False
          Properties.DropDownListStyle = lsEditFixedList
          Properties.Items.Strings = (
            ''
            '1.'#20381#36027#29575#34920
            '2.'#20381#20778#24800#37329#38989)
          Style.StyleController = StyleController.EditorStyle
          TabOrder = 30
          Height = 20
          Width = 105
        end
        object cmbRateType8: TcxComboBox
          Left = 158
          Top = 186
          AutoSize = False
          ParentFont = False
          Properties.DropDownListStyle = lsEditFixedList
          Properties.Items.Strings = (
            ''
            '1.'#20381#36027#29575#34920
            '2.'#20381#20778#24800#37329#38989)
          Style.StyleController = StyleController.EditorStyle
          TabOrder = 31
          Height = 20
          Width = 105
        end
        object cmbRateType9: TcxComboBox
          Left = 158
          Top = 206
          AutoSize = False
          ParentFont = False
          Properties.DropDownListStyle = lsEditFixedList
          Properties.Items.Strings = (
            ''
            '1.'#20381#36027#29575#34920
            '2.'#20381#20778#24800#37329#38989)
          Style.StyleController = StyleController.EditorStyle
          TabOrder = 32
          Height = 20
          Width = 105
        end
        object cmbRateType10: TcxComboBox
          Left = 158
          Top = 226
          AutoSize = False
          ParentFont = False
          Properties.DropDownListStyle = lsEditFixedList
          Properties.Items.Strings = (
            ''
            '1.'#20381#36027#29575#34920
            '2.'#20381#20778#24800#37329#38989)
          Style.StyleController = StyleController.EditorStyle
          TabOrder = 33
          Height = 20
          Width = 105
        end
        object cmbRateType11: TcxComboBox
          Left = 158
          Top = 246
          AutoSize = False
          ParentFont = False
          Properties.DropDownListStyle = lsEditFixedList
          Properties.Items.Strings = (
            ''
            '1.'#20381#36027#29575#34920
            '2.'#20381#20778#24800#37329#38989)
          Style.StyleController = StyleController.EditorStyle
          TabOrder = 34
          Height = 20
          Width = 105
        end
        object cmbRateType12: TcxComboBox
          Left = 158
          Top = 266
          AutoSize = False
          ParentFont = False
          Properties.DropDownListStyle = lsEditFixedList
          Properties.Items.Strings = (
            ''
            '1.'#20381#36027#29575#34920
            '2.'#20381#20778#24800#37329#38989)
          Style.StyleController = StyleController.EditorStyle
          TabOrder = 35
          Height = 20
          Width = 105
        end
        object edtDiscountAmt1: TcxCurrencyEdit
          Left = 263
          Top = 46
          AutoSize = False
          EditValue = 0.000000000000000000
          ParentFont = False
          Properties.Alignment.Horz = taRightJustify
          Properties.DecimalPlaces = 0
          Properties.DisplayFormat = ',0;-,0'
          Properties.MaxValue = 99999999.000000000000000000
          Style.StyleController = StyleController.EditorStyle
          TabOrder = 36
          Height = 20
          Width = 56
        end
        object edtMonthAmt1: TcxCurrencyEdit
          Left = 319
          Top = 46
          AutoSize = False
          EditValue = 0.000000000000000000
          ParentFont = False
          Properties.Alignment.Horz = taRightJustify
          Properties.DecimalPlaces = 3
          Properties.DisplayFormat = ',0.000;-,0.000'
          Properties.MaxValue = 99999999.999000000000000000
          Properties.NullString = '0'
          Style.StyleController = StyleController.EditorStyle
          TabOrder = 37
          Height = 20
          Width = 56
        end
        object edtDayAmt1: TcxCurrencyEdit
          Left = 375
          Top = 46
          AutoSize = False
          EditValue = 0.000000000000000000
          ParentFont = False
          Properties.Alignment.Horz = taRightJustify
          Properties.DecimalPlaces = 3
          Properties.DisplayFormat = ',0.000;-,0.000'
          Properties.MaxValue = 9999.999000000000000000
          Style.StyleController = StyleController.EditorStyle
          TabOrder = 38
          Height = 20
          Width = 56
        end
        object edtDiscountAmt2: TcxCurrencyEdit
          Left = 263
          Top = 66
          AutoSize = False
          EditValue = 0.000000000000000000
          ParentFont = False
          Properties.Alignment.Horz = taRightJustify
          Properties.DecimalPlaces = 0
          Properties.DisplayFormat = ',0;-,0'
          Properties.MaxValue = 99999999.000000000000000000
          Style.StyleController = StyleController.EditorStyle
          TabOrder = 39
          Height = 20
          Width = 56
        end
        object edtMonthAmt2: TcxCurrencyEdit
          Left = 319
          Top = 66
          AutoSize = False
          EditValue = 0.000000000000000000
          ParentFont = False
          Properties.Alignment.Horz = taRightJustify
          Properties.DecimalPlaces = 3
          Properties.DisplayFormat = ',0.000;-,0.000'
          Properties.MaxValue = 99999999.999000000000000000
          Properties.NullString = '0'
          Style.StyleController = StyleController.EditorStyle
          TabOrder = 40
          Height = 20
          Width = 56
        end
        object edtDayAmt2: TcxCurrencyEdit
          Left = 375
          Top = 66
          AutoSize = False
          EditValue = 0.000000000000000000
          ParentFont = False
          Properties.Alignment.Horz = taRightJustify
          Properties.DecimalPlaces = 3
          Properties.DisplayFormat = ',0.000;-,0.000'
          Properties.MaxValue = 9999.999000000000000000
          Style.StyleController = StyleController.EditorStyle
          TabOrder = 41
          Height = 20
          Width = 56
        end
        object edtDiscountAmt3: TcxCurrencyEdit
          Left = 263
          Top = 86
          AutoSize = False
          EditValue = 0.000000000000000000
          ParentFont = False
          Properties.Alignment.Horz = taRightJustify
          Properties.DecimalPlaces = 0
          Properties.DisplayFormat = ',0;-,0'
          Properties.MaxValue = 99999999.000000000000000000
          Style.StyleController = StyleController.EditorStyle
          TabOrder = 42
          Height = 20
          Width = 56
        end
        object edtMonthAmt3: TcxCurrencyEdit
          Left = 319
          Top = 86
          AutoSize = False
          EditValue = 0.000000000000000000
          ParentFont = False
          Properties.Alignment.Horz = taRightJustify
          Properties.DecimalPlaces = 3
          Properties.DisplayFormat = ',0.000;-,0.000'
          Properties.MaxValue = 99999999.999000000000000000
          Properties.NullString = '0'
          Style.StyleController = StyleController.EditorStyle
          TabOrder = 43
          Height = 20
          Width = 56
        end
        object edtDayAmt3: TcxCurrencyEdit
          Left = 375
          Top = 86
          AutoSize = False
          EditValue = 0.000000000000000000
          ParentFont = False
          Properties.Alignment.Horz = taRightJustify
          Properties.DecimalPlaces = 3
          Properties.DisplayFormat = ',0.000;-,0.000'
          Properties.MaxValue = 9999.999000000000000000
          Style.StyleController = StyleController.EditorStyle
          TabOrder = 44
          Height = 20
          Width = 56
        end
        object edtDiscountAmt4: TcxCurrencyEdit
          Left = 263
          Top = 106
          AutoSize = False
          EditValue = 0.000000000000000000
          ParentFont = False
          Properties.Alignment.Horz = taRightJustify
          Properties.DecimalPlaces = 0
          Properties.DisplayFormat = ',0;-,0'
          Properties.MaxValue = 99999999.000000000000000000
          Style.StyleController = StyleController.EditorStyle
          TabOrder = 45
          Height = 20
          Width = 56
        end
        object edtMonthAmt4: TcxCurrencyEdit
          Left = 319
          Top = 106
          AutoSize = False
          EditValue = 0.000000000000000000
          ParentFont = False
          Properties.Alignment.Horz = taRightJustify
          Properties.DecimalPlaces = 3
          Properties.DisplayFormat = ',0.000;-,0.000'
          Properties.MaxValue = 99999999.999000000000000000
          Properties.NullString = '0'
          Style.StyleController = StyleController.EditorStyle
          TabOrder = 46
          Height = 20
          Width = 56
        end
        object edtDayAmt4: TcxCurrencyEdit
          Left = 375
          Top = 106
          AutoSize = False
          EditValue = 0.000000000000000000
          ParentFont = False
          Properties.Alignment.Horz = taRightJustify
          Properties.DecimalPlaces = 3
          Properties.DisplayFormat = ',0.000;-,0.000'
          Properties.MaxValue = 9999.999000000000000000
          Style.StyleController = StyleController.EditorStyle
          TabOrder = 47
          Height = 20
          Width = 56
        end
        object edtDiscountAmt5: TcxCurrencyEdit
          Left = 263
          Top = 126
          AutoSize = False
          EditValue = 0.000000000000000000
          ParentFont = False
          Properties.Alignment.Horz = taRightJustify
          Properties.DecimalPlaces = 0
          Properties.DisplayFormat = ',0;-,0'
          Properties.MaxValue = 99999999.000000000000000000
          Style.StyleController = StyleController.EditorStyle
          TabOrder = 48
          Height = 20
          Width = 56
        end
        object edtMonthAmt5: TcxCurrencyEdit
          Left = 319
          Top = 126
          AutoSize = False
          EditValue = 0.000000000000000000
          ParentFont = False
          Properties.Alignment.Horz = taRightJustify
          Properties.DecimalPlaces = 3
          Properties.DisplayFormat = ',0.000;-,0.000'
          Properties.MaxValue = 99999999.999000000000000000
          Properties.NullString = '0'
          Style.StyleController = StyleController.EditorStyle
          TabOrder = 49
          Height = 20
          Width = 56
        end
        object edtDayAmt5: TcxCurrencyEdit
          Left = 375
          Top = 126
          AutoSize = False
          EditValue = 0.000000000000000000
          ParentFont = False
          Properties.Alignment.Horz = taRightJustify
          Properties.DecimalPlaces = 3
          Properties.DisplayFormat = ',0.000;-,0.000'
          Properties.MaxValue = 9999.999000000000000000
          Style.StyleController = StyleController.EditorStyle
          TabOrder = 50
          Height = 20
          Width = 56
        end
        object edtDiscountAmt6: TcxCurrencyEdit
          Left = 263
          Top = 146
          AutoSize = False
          EditValue = 0.000000000000000000
          ParentFont = False
          Properties.Alignment.Horz = taRightJustify
          Properties.DecimalPlaces = 0
          Properties.DisplayFormat = ',0;-,0'
          Properties.MaxValue = 99999999.000000000000000000
          Style.StyleController = StyleController.EditorStyle
          TabOrder = 51
          Height = 20
          Width = 56
        end
        object edtMonthAmt6: TcxCurrencyEdit
          Left = 319
          Top = 146
          AutoSize = False
          EditValue = 0.000000000000000000
          ParentFont = False
          Properties.Alignment.Horz = taRightJustify
          Properties.DecimalPlaces = 3
          Properties.DisplayFormat = ',0.000;-,0.000'
          Properties.MaxValue = 99999999.999000000000000000
          Properties.NullString = '0'
          Style.StyleController = StyleController.EditorStyle
          TabOrder = 52
          Height = 20
          Width = 56
        end
        object edtDayAmt6: TcxCurrencyEdit
          Left = 375
          Top = 146
          AutoSize = False
          EditValue = 0.000000000000000000
          ParentFont = False
          Properties.Alignment.Horz = taRightJustify
          Properties.DecimalPlaces = 3
          Properties.DisplayFormat = ',0.000;-,0.000'
          Properties.MaxValue = 9999.999000000000000000
          Style.StyleController = StyleController.EditorStyle
          TabOrder = 53
          Height = 20
          Width = 56
        end
        object edtDiscountAmt7: TcxCurrencyEdit
          Left = 263
          Top = 166
          AutoSize = False
          EditValue = 0.000000000000000000
          ParentFont = False
          Properties.Alignment.Horz = taRightJustify
          Properties.DecimalPlaces = 0
          Properties.DisplayFormat = ',0;-,0'
          Properties.MaxValue = 99999999.000000000000000000
          Style.StyleController = StyleController.EditorStyle
          TabOrder = 54
          Height = 20
          Width = 56
        end
        object edtMonthAmt7: TcxCurrencyEdit
          Left = 319
          Top = 166
          AutoSize = False
          EditValue = 0.000000000000000000
          ParentFont = False
          Properties.Alignment.Horz = taRightJustify
          Properties.DecimalPlaces = 3
          Properties.DisplayFormat = ',0.000;-,0.000'
          Properties.MaxValue = 99999999.999000000000000000
          Properties.NullString = '0'
          Style.StyleController = StyleController.EditorStyle
          TabOrder = 55
          Height = 20
          Width = 56
        end
        object edtDayAmt7: TcxCurrencyEdit
          Left = 375
          Top = 166
          AutoSize = False
          EditValue = 0.000000000000000000
          ParentFont = False
          Properties.Alignment.Horz = taRightJustify
          Properties.DecimalPlaces = 3
          Properties.DisplayFormat = ',0.000;-,0.000'
          Properties.MaxValue = 9999.999000000000000000
          Style.StyleController = StyleController.EditorStyle
          TabOrder = 56
          Height = 20
          Width = 56
        end
        object edtDiscountAmt8: TcxCurrencyEdit
          Left = 263
          Top = 186
          AutoSize = False
          EditValue = 0.000000000000000000
          ParentFont = False
          Properties.Alignment.Horz = taRightJustify
          Properties.DecimalPlaces = 0
          Properties.DisplayFormat = ',0;-,0'
          Properties.MaxValue = 99999999.000000000000000000
          Style.StyleController = StyleController.EditorStyle
          TabOrder = 57
          Height = 20
          Width = 56
        end
        object edtMonthAmt8: TcxCurrencyEdit
          Left = 319
          Top = 186
          AutoSize = False
          EditValue = 0.000000000000000000
          ParentFont = False
          Properties.Alignment.Horz = taRightJustify
          Properties.DecimalPlaces = 3
          Properties.DisplayFormat = ',0.000;-,0.000'
          Properties.MaxValue = 99999999.999000000000000000
          Properties.NullString = '0'
          Style.StyleController = StyleController.EditorStyle
          TabOrder = 58
          Height = 20
          Width = 56
        end
        object edtDayAmt8: TcxCurrencyEdit
          Left = 375
          Top = 186
          AutoSize = False
          EditValue = 0.000000000000000000
          ParentFont = False
          Properties.Alignment.Horz = taRightJustify
          Properties.DecimalPlaces = 3
          Properties.DisplayFormat = ',0.000;-,0.000'
          Properties.MaxValue = 9999.999000000000000000
          Style.StyleController = StyleController.EditorStyle
          TabOrder = 59
          Height = 20
          Width = 56
        end
        object edtDiscountAmt9: TcxCurrencyEdit
          Left = 263
          Top = 206
          AutoSize = False
          EditValue = 0.000000000000000000
          ParentFont = False
          Properties.Alignment.Horz = taRightJustify
          Properties.DecimalPlaces = 0
          Properties.DisplayFormat = ',0;-,0'
          Properties.MaxValue = 99999999.000000000000000000
          Style.StyleController = StyleController.EditorStyle
          TabOrder = 60
          Height = 20
          Width = 56
        end
        object edtMonthAmt9: TcxCurrencyEdit
          Left = 319
          Top = 206
          AutoSize = False
          EditValue = 0.000000000000000000
          ParentFont = False
          Properties.Alignment.Horz = taRightJustify
          Properties.DecimalPlaces = 3
          Properties.DisplayFormat = ',0.000;-,0.000'
          Properties.MaxValue = 99999999.999000000000000000
          Properties.NullString = '0'
          Style.StyleController = StyleController.EditorStyle
          TabOrder = 61
          Height = 20
          Width = 56
        end
        object edtDayAmt9: TcxCurrencyEdit
          Left = 375
          Top = 206
          AutoSize = False
          EditValue = 0.000000000000000000
          ParentFont = False
          Properties.Alignment.Horz = taRightJustify
          Properties.DecimalPlaces = 3
          Properties.DisplayFormat = ',0.000;-,0.000'
          Properties.MaxValue = 9999.999000000000000000
          Style.StyleController = StyleController.EditorStyle
          TabOrder = 62
          Height = 20
          Width = 56
        end
        object edtDiscountAmt10: TcxCurrencyEdit
          Left = 263
          Top = 226
          AutoSize = False
          EditValue = 0.000000000000000000
          ParentFont = False
          Properties.Alignment.Horz = taRightJustify
          Properties.DecimalPlaces = 0
          Properties.DisplayFormat = ',0;-,0'
          Properties.MaxValue = 99999999.000000000000000000
          Style.StyleController = StyleController.EditorStyle
          TabOrder = 63
          Height = 20
          Width = 56
        end
        object edtMonthAmt10: TcxCurrencyEdit
          Left = 319
          Top = 226
          AutoSize = False
          EditValue = 0.000000000000000000
          ParentFont = False
          Properties.Alignment.Horz = taRightJustify
          Properties.DecimalPlaces = 3
          Properties.DisplayFormat = ',0.000;-,0.000'
          Properties.MaxValue = 99999999.999000000000000000
          Properties.NullString = '0'
          Style.StyleController = StyleController.EditorStyle
          TabOrder = 64
          Height = 20
          Width = 56
        end
        object edtDayAmt10: TcxCurrencyEdit
          Left = 375
          Top = 226
          AutoSize = False
          EditValue = 0.000000000000000000
          ParentFont = False
          Properties.Alignment.Horz = taRightJustify
          Properties.DecimalPlaces = 3
          Properties.DisplayFormat = ',0.000;-,0.000'
          Properties.MaxValue = 9999.999000000000000000
          Style.StyleController = StyleController.EditorStyle
          TabOrder = 65
          Height = 20
          Width = 56
        end
        object edtDiscountAmt11: TcxCurrencyEdit
          Left = 263
          Top = 246
          AutoSize = False
          EditValue = 0.000000000000000000
          ParentFont = False
          Properties.Alignment.Horz = taRightJustify
          Properties.DecimalPlaces = 0
          Properties.DisplayFormat = ',0;-,0'
          Properties.MaxValue = 99999999.000000000000000000
          Style.StyleController = StyleController.EditorStyle
          TabOrder = 66
          Height = 20
          Width = 56
        end
        object edtMonthAmt11: TcxCurrencyEdit
          Left = 319
          Top = 246
          AutoSize = False
          EditValue = 0.000000000000000000
          ParentFont = False
          Properties.Alignment.Horz = taRightJustify
          Properties.DecimalPlaces = 3
          Properties.DisplayFormat = ',0.000;-,0.000'
          Properties.MaxValue = 99999999.999000000000000000
          Properties.NullString = '0'
          Style.StyleController = StyleController.EditorStyle
          TabOrder = 67
          Height = 20
          Width = 56
        end
        object edtDayAmt11: TcxCurrencyEdit
          Left = 375
          Top = 246
          AutoSize = False
          EditValue = 0.000000000000000000
          ParentFont = False
          Properties.Alignment.Horz = taRightJustify
          Properties.DecimalPlaces = 3
          Properties.DisplayFormat = ',0.000;-,0.000'
          Properties.MaxValue = 9999.999000000000000000
          Style.StyleController = StyleController.EditorStyle
          TabOrder = 68
          Height = 20
          Width = 56
        end
        object edtDiscountAmt12: TcxCurrencyEdit
          Left = 263
          Top = 266
          AutoSize = False
          EditValue = 0.000000000000000000
          ParentFont = False
          Properties.Alignment.Horz = taRightJustify
          Properties.DecimalPlaces = 0
          Properties.DisplayFormat = ',0;-,0'
          Properties.MaxValue = 99999999.000000000000000000
          Style.StyleController = StyleController.EditorStyle
          TabOrder = 69
          Height = 20
          Width = 56
        end
        object edtMonthAmt12: TcxCurrencyEdit
          Left = 319
          Top = 266
          AutoSize = False
          EditValue = 0.000000000000000000
          ParentFont = False
          Properties.Alignment.Horz = taRightJustify
          Properties.DecimalPlaces = 3
          Properties.DisplayFormat = ',0.000;-,0.000'
          Properties.MaxValue = 99999999.999000000000000000
          Properties.NullString = '0'
          Style.StyleController = StyleController.EditorStyle
          TabOrder = 70
          Height = 20
          Width = 56
        end
        object edtDayAmt12: TcxCurrencyEdit
          Left = 375
          Top = 266
          AutoSize = False
          EditValue = 0.000000000000000000
          ParentFont = False
          Properties.Alignment.Horz = taRightJustify
          Properties.DecimalPlaces = 3
          Properties.DisplayFormat = ',0.000;-,0.000'
          Properties.MaxValue = 9999.999000000000000000
          Style.StyleController = StyleController.EditorStyle
          TabOrder = 71
          Height = 20
          Width = 56
        end
        object edtStepNo: TcxTextEdit
          Tag = 1
          Left = 101
          Top = 4
          ParentFont = False
          Properties.ReadOnly = True
          Style.StyleController = StyleController.EditorStyle
          TabOrder = 72
          Width = 62
        end
        object chkMasterSale: TcxCheckBox
          Tag = 1
          Left = 378
          Top = 4
          Caption = #20027#25512#22810#38542
          Properties.NullStyle = nssUnchecked
          Properties.ReadOnly = True
          Style.StyleController = StyleController.CheckBoxStyle
          TabOrder = 73
          Width = 74
        end
        object edtDescription: TcxTextEdit
          Tag = 1
          Left = 165
          Top = 4
          ParentFont = False
          Properties.ReadOnly = True
          Style.StyleController = StyleController.EditorStyle
          TabOrder = 74
          Width = 207
        end
        object btnRankList: TBitBtn
          Left = 16
          Top = 2
          Width = 26
          Height = 24
          Hint = #39023#31034#22810#38542#20778#24800#21443#25976#26126#32048
          ParentShowHint = False
          ShowHint = True
          TabOrder = 75
          OnClick = btnRankListClick
          Glyph.Data = {
            36040000424D3604000000000000360000002800000010000000100000000100
            2000000000000004000000000000000000000000000000000000000000000000
            0000000000000000000000000000000000000000000000000000000000000000
            0000000000000000000000000000000000000000000000000000000000000000
            0000A27A66CCC9967EFFC5927CFFC18E7AFFBD8A78FFB98676FFB58274FFB17E
            72FFAD7A70FFA9766EFFA5726CFF845B56CC000000000000000000000000A57C
            66CCEEDACEFFF9F9F9FFF9F9F9FFFAFAFAFFFCFCFCFFFEFEFEFFFFFFFFFFFFFF
            FFFFFFFFFFFFFFFFFFFFFFFFFFFFDBC7C4FF845B56CC0000000000000000D29F
            82FFFFFFFFFF0033FFFF0033FFFF0033FFFF0033FFFF0033FFFFFDFDFDFFF38D
            00FFF48E00FFF69000FFF89200FFFFFFFFFFA5726CFF0000000000000000D7A4
            84FFFFFFFFFFF3F3F3FFF4F4F4FFF5F5F5FFF6F6F6FFF8F8F8FFF9F9F9FFF18B
            00FFF38D00FFF48E00FFF69000FFFFFFFFFFA9766EFF0000000000000000DCA9
            87FFFFFFFFFF00BFF2FF00BFF2FF00BFF2FF00BFF2FF00BFF2FFF7F7F7FFEF89
            00FFF18B00FFF38D00FFF48E00FFFFFFFFFFAE7B70FF0000000000000000E1AE
            8AFFFFFFFFFFECECECFFEDEDEDFFEEEEEEFFF1F1F1FFF3F3F3FFF5F5F5FFF7F7
            F7FFF9F9F9FFFCFCFCFFFFFFFFFFFFFFFFFFB38073FF0000000000000000E5B2
            8CFFFFFFFFFFE37D00FFE47E00FFE57F00FFE78100FFE88200FFF2F2F2FF0033
            FFFF0033FFFF0033FFFF0033FFFFFFFFFFFFB78475FF0000000000000000E9B6
            8EFFFFFFFFFFE17B00FFE37D00FFE37D00FFE57F00FFE68000FFEFEFEFFFF3F3
            F3FFF5F5F5FFF8F8F8FFFAFAFAFFFEFEFEFFBA8777FF0000000000000000EDBA
            90FFFFFFFFFFDF7900FFE07A00FFE27C00FFE37D00FFE57F00FFEEEEEEFF00A9
            DCFF00A9DCFF00A9DCFF00A9DCFFFCFCFCFFBE8B79FF0000000000000000F1BE
            92FFFFFFFFFFE0E0E0FFE3E3E3FFE5E5E5FFE7E7E7FFE9E9E9FFECECECFFEEEE
            EEFFF2F2F2FFF5F5F5FFF8F8F8FFFAFAFAFFC28F7BFF0000000000000000F5C2
            94FFFFFFFFFF0EA71CFF0FA81EFF12AB24FF14AD28FF17B02EFF18B131FF1AB3
            36FF1DB63BFF1FB83FFF21BA43FFF9F9F9FFC6937DFF0000000000000000F9C6
            96FFFFFFFFFF0BA417FF0EA71CFF10A920FF13AC26FF15AE2AFF17B02FFF19B2
            34FF1CB539FF1EB73EFF21BA43FFF9F9F9FFC9967EFF0000000000000000C29A
            76CCFDE8D5FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
            FFFFFFFFFFFFFFFFFFFFFFFFFFFFEEDACEFF885F58CC00000000000000000000
            0000BE9673CCF9C696FFF5C294FFF1BE92FFEDBA90FFE9B68EFFE5B28CFFE1AE
            8AFFDDAA88FFD9A686FFD5A284FF885F58CC0000000000000000000000000000
            0000000000000000000000000000000000000000000000000000000000000000
            0000000000000000000000000000000000000000000000000000}
        end
      end
      object btnRankWhere: TBitBtn
        Left = 17
        Top = 16
        Width = 115
        Height = 22
        Caption = #20778#24800#35215#21063#35373#23450
        ParentShowHint = False
        ShowHint = True
        TabOrder = 1
        OnClick = btnRankWhereClick
      end
    end
    object cxGroupBox3: TcxGroupBox
      Left = 8
      Top = 2
      Caption = #38542#27573#24335#36949#32004#21443#25976
      ParentFont = False
      Style.StyleController = StyleController.EditorStyle
      Style.TextColor = clBlue
      Style.TextStyle = [fsBold]
      TabOrder = 0
      Height = 323
      Width = 306
      object Label17: TLabel
        Left = 23
        Top = 39
        Width = 60
        Height = 14
        Caption = #36949#32004#37329#38917#30446
      end
      object chkPenal: TcxCheckBox
        Left = 3
        Top = 17
        Caption = #36949#32004#37329#25353#26376#25892#25552
        Properties.OnChange = chkPenalPropertiesChange
        Style.StyleController = StyleController.CheckBoxStyle
        TabOrder = 0
        Width = 110
      end
      inline Penal: TDataLookup
        Left = 88
        Top = 34
        Width = 206
        Height = 23
        HorzScrollBar.Visible = False
        VertScrollBar.Visible = False
        ParentBackground = False
        TabOrder = 1
        inherited CodeNo: TcxTextEdit
          Properties.MaxLength = 5
          Width = 50
        end
        inherited CodeName: TcxLookupComboBox
          Left = 53
          Properties.OnInitPopup = PenalCodeNamePropertiesInitPopup
          Width = 152
        end
      end
      object PenalGrid: TcxGrid
        Left = 8
        Top = 88
        Width = 285
        Height = 225
        TabOrder = 2
        LookAndFeel.Kind = lfFlat
        object gvPenal: TcxGridDBTableView
          NavigatorButtons.ConfirmDelete = False
          OnCustomDrawCell = gvPenalCustomDrawCell
          DataController.DataSource = dsCD078A3
          DataController.Summary.DefaultGroupSummaryItems = <>
          DataController.Summary.FooterSummaryItems = <>
          DataController.Summary.SummaryGroups = <>
          OptionsBehavior.FocusFirstCellOnNewRecord = True
          OptionsBehavior.GoToNextCellOnEnter = True
          OptionsBehavior.FocusCellOnCycle = True
          OptionsCustomize.ColumnFiltering = False
          OptionsCustomize.ColumnMoving = False
          OptionsCustomize.ColumnSorting = False
          OptionsData.Deleting = False
          OptionsData.Editing = False
          OptionsData.Inserting = False
          OptionsSelection.CellSelect = False
          OptionsView.GroupByBox = False
          Styles.Background = StyleController.cxStyle1
          object gvPenalCol1: TcxGridDBColumn
            Caption = #38917#27425
            DataBinding.FieldName = 'ITEM'
            GroupSummaryAlignment = taCenter
            HeaderAlignmentHorz = taCenter
            Options.Focusing = False
            Width = 44
          end
          object gvPenalCol2: TcxGridDBColumn
            Caption = #26410#28415#32129#32004#26376#25976
            DataBinding.FieldName = 'PenalMon'
            PropertiesClassName = 'TcxMaskEditProperties'
            Properties.MaskKind = emkRegExpr
            Properties.EditMask = '[\d] | [\d][\d] | [\d][\d][\d] '
            Properties.MaxLength = 0
            Properties.OnValidate = gvPenalCol2PropertiesValidate
            HeaderAlignmentHorz = taCenter
            HeaderAlignmentVert = vaCenter
            Width = 97
          end
          object gvPenalCol3: TcxGridDBColumn
            Caption = #36949#32004#37329#38989
            DataBinding.FieldName = 'PenalAMT'
            PropertiesClassName = 'TcxCurrencyEditProperties'
            Properties.DecimalPlaces = 0
            Properties.DisplayFormat = ',0;-,0'
            Properties.MaxValue = 99999999.000000000000000000
            Properties.NullString = '0'
            Properties.OnValidate = gvPenalCol3PropertiesValidate
            HeaderAlignmentHorz = taRightJustify
            Width = 102
          end
        end
        object glPenal: TcxGridLevel
          GridView = gvPenal
        end
      end
      object btnPenalCondition: TBitBtn
        Left = 8
        Top = 62
        Width = 134
        Height = 22
        Caption = #26781#20214#35215#21063#35373#23450
        ParentShowHint = False
        ShowHint = True
        TabOrder = 3
        OnClick = btnPenalConditionClick
      end
    end
  end
  object ActionList1: TActionList
    Left = 36
    Top = 366
    object actSave: TAction
      Caption = 'F2.'#23384#27284
      ShortCut = 113
      OnExecute = actSaveExecute
    end
    object actCancel: TAction
      Caption = #21462#28040'(&X)'
      ShortCut = 32856
      OnExecute = actCancelExecute
    end
    object actUpdate: TAction
      Caption = 'F11.'#20462#25913
      ShortCut = 122
      OnExecute = actUpdateExecute
    end
  end
  object dsCD078A3: TDataSource
    Left = 72
    Top = 364
  end
end
