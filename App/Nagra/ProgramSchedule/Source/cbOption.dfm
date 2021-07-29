object fmOption: TfmOption
  Left = 377
  Top = 303
  BorderStyle = bsDialog
  Caption = #36984#38917
  ClientHeight = 414
  ClientWidth = 655
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
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 14
  object RzPanel1: TRzPanel
    Left = 0
    Top = 372
    Width = 655
    Height = 42
    Align = alBottom
    BorderOuter = fsGroove
    BorderSides = [sdTop]
    TabOrder = 0
    object btnConfirm: TcxButton
      Left = 469
      Top = 8
      Width = 81
      Height = 25
      Caption = #20786#23384
      Default = True
      TabOrder = 0
      OnClick = btnConfirmClick
      LookAndFeel.Kind = lfFlat
    end
    object btnCancel: TcxButton
      Left = 559
      Top = 8
      Width = 81
      Height = 25
      Cancel = True
      Caption = #21462#28040
      ModalResult = 2
      TabOrder = 1
      LookAndFeel.Kind = lfFlat
    end
  end
  object RzPanel2: TRzPanel
    Left = 0
    Top = 0
    Width = 655
    Height = 372
    Align = alClient
    BorderOuter = fsFlat
    BorderSides = [sdLeft, sdTop, sdRight]
    TabOrder = 1
    object OptionPage: TcxPageControl
      Left = 1
      Top = 1
      Width = 653
      Height = 371
      ActivePage = cxTabSheet5
      Align = alClient
      Images = StyleModule.OptionImageList
      Style = 9
      TabOrder = 0
      ClientRectBottom = 371
      ClientRectRight = 653
      ClientRectTop = 22
      object cxTabSheet5: TcxTabSheet
        Caption = '  FTP '#20282#26381#22120'  '
        ImageIndex = 6
        object cxGroupBox7: TcxGroupBox
          Left = 20
          Top = 25
          Caption = 'FTP '#35373#23450
          Style.StyleController = StyleModule.cxGroupBoxStyle
          TabOrder = 0
          Height = 55
          Width = 617
          object Label20: TLabel
            Left = 30
            Top = 23
            Width = 40
            Height = 14
            Caption = #20282#26381#22120':'
          end
          object Label21: TLabel
            Left = 236
            Top = 23
            Width = 28
            Height = 14
            Caption = #24115#34399':'
          end
          object Label22: TLabel
            Left = 421
            Top = 23
            Width = 28
            Height = 14
            Caption = #23494#30908':'
          end
          object edtFtpServer: TcxTextEdit
            Left = 79
            Top = 20
            ParentFont = False
            Style.StyleController = StyleModule.cxEditStyle
            TabOrder = 0
            Width = 134
          end
          object edtFtpAccount: TcxTextEdit
            Left = 273
            Top = 20
            ParentFont = False
            Properties.EchoMode = eemPassword
            Properties.PasswordChar = '*'
            Style.StyleController = StyleModule.cxEditStyle
            TabOrder = 1
            Width = 121
          end
          object edtFtpPassword: TcxTextEdit
            Left = 457
            Top = 20
            ParentFont = False
            Properties.EchoMode = eemPassword
            Properties.PasswordChar = '*'
            Style.StyleController = StyleModule.cxEditStyle
            TabOrder = 2
            Width = 121
          end
        end
        object cxGroupBox8: TcxGroupBox
          Left = 20
          Top = 98
          Caption = 'FTP '#36960#31471#27284#26696#30446#37636
          Style.StyleController = StyleModule.cxGroupBoxStyle
          TabOrder = 1
          Height = 124
          Width = 617
          object Label23: TLabel
            Left = 36
            Top = 25
            Width = 74
            Height = 14
            Caption = 'DEX'#19979#36617#30446#37636':'
          end
          object Label24: TLabel
            Left = 11
            Top = 55
            Width = 99
            Height = 14
            Caption = 'Wrapper'#19979#36617#30446#37636':'
          end
          object Label25: TLabel
            Left = 24
            Top = 85
            Width = 86
            Height = 14
            Caption = 'AsRun'#19979#36617#30446#37636':'
          end
          object Label26: TLabel
            Left = 321
            Top = 25
            Width = 76
            Height = 14
            Caption = #19979#36617#24460#20633#20221#33267':'
          end
          object Label27: TLabel
            Left = 321
            Top = 55
            Width = 76
            Height = 14
            Caption = #19979#36617#24460#20633#20221#33267':'
          end
          object Label28: TLabel
            Left = 321
            Top = 85
            Width = 76
            Height = 14
            Caption = #19979#36617#24460#20633#20221#33267':'
          end
          object edtFtpDexSrc: TcxTextEdit
            Left = 114
            Top = 22
            ParentFont = False
            Style.StyleController = StyleModule.cxEditStyle
            TabOrder = 0
            Width = 190
          end
          object edtFtpWrapperSrc: TcxTextEdit
            Left = 114
            Top = 52
            ParentFont = False
            Style.StyleController = StyleModule.cxEditStyle
            TabOrder = 2
            Width = 190
          end
          object edtFtpAsRunSrc: TcxTextEdit
            Left = 114
            Top = 82
            ParentFont = False
            Style.StyleController = StyleModule.cxEditStyle
            TabOrder = 4
            Width = 190
          end
          object edtFtpDexBackup: TcxTextEdit
            Left = 403
            Top = 22
            ParentFont = False
            Style.StyleController = StyleModule.cxEditStyle
            TabOrder = 1
            Width = 190
          end
          object edtFtpWrapperBackup: TcxTextEdit
            Left = 403
            Top = 52
            ParentFont = False
            Style.StyleController = StyleModule.cxEditStyle
            TabOrder = 3
            Width = 190
          end
          object edtFtpAsRunBackup: TcxTextEdit
            Left = 403
            Top = 82
            ParentFont = False
            Style.StyleController = StyleModule.cxEditStyle
            TabOrder = 5
            Width = 190
          end
        end
      end
      object cxTabSheet1: TcxTabSheet
        Caption = '  '#26412#27231#30446#37636#35373#23450'  '
        ImageIndex = 0
        object DexGroup: TcxGroupBox
          Left = 16
          Top = 25
          Caption = ' Dex '#27284#26696#36335#24465' '
          Style.StyleController = StyleModule.cxGroupBoxStyle
          TabOrder = 0
          Height = 81
          Width = 622
          object Label1: TLabel
            Left = 20
            Top = 24
            Width = 52
            Height = 14
            Caption = #20358#28304#20301#32622':'
          end
          object Label2: TLabel
            Left = 20
            Top = 48
            Width = 52
            Height = 14
            Caption = #30446#30340#20301#32622':'
          end
          object Lable18: TLabel
            Left = 320
            Top = 24
            Width = 52
            Height = 14
            Caption = #37679#35492#20301#32622':'
          end
          object Label3: TLabel
            Left = 320
            Top = 48
            Width = 52
            Height = 14
            Caption = #20633#20221#20301#32622':'
          end
          object edtDexSrc: TcxButtonEdit
            Left = 79
            Top = 20
            ParentFont = False
            Properties.Buttons = <
              item
                Default = True
                Kind = bkEllipsis
              end>
            Properties.ReadOnly = True
            Properties.OnButtonClick = FilePathButtonClick
            Style.StyleController = StyleModule.cxEditStyle
            TabOrder = 0
            Width = 230
          end
          object edtDexDest: TcxButtonEdit
            Left = 79
            Top = 45
            ParentFont = False
            Properties.Buttons = <
              item
                Default = True
                Kind = bkEllipsis
              end>
            Properties.ReadOnly = True
            Properties.OnButtonClick = FilePathButtonClick
            Style.StyleController = StyleModule.cxEditStyle
            TabOrder = 1
            Width = 230
          end
          object edtDexErr: TcxButtonEdit
            Left = 379
            Top = 20
            ParentFont = False
            Properties.Buttons = <
              item
                Default = True
                Kind = bkEllipsis
              end>
            Properties.ReadOnly = True
            Properties.OnButtonClick = FilePathButtonClick
            Style.StyleController = StyleModule.cxEditStyle
            TabOrder = 2
            Width = 230
          end
          object edtDexBackup: TcxButtonEdit
            Left = 379
            Top = 45
            ParentFont = False
            Properties.Buttons = <
              item
                Default = True
                Kind = bkEllipsis
              end>
            Properties.ReadOnly = True
            Properties.OnButtonClick = FilePathButtonClick
            Style.StyleController = StyleModule.cxEditStyle
            TabOrder = 3
            Width = 230
          end
        end
        object WrapperGroup: TcxGroupBox
          Left = 16
          Top = 125
          Caption = ' Wrapper '#27284#26696#36335#24465' '
          Style.StyleController = StyleModule.cxGroupBoxStyle
          TabOrder = 1
          Height = 81
          Width = 622
          object Label4: TLabel
            Left = 20
            Top = 24
            Width = 52
            Height = 14
            Caption = #20358#28304#20301#32622':'
          end
          object Label5: TLabel
            Left = 20
            Top = 48
            Width = 52
            Height = 14
            Caption = #30446#30340#20301#32622':'
          end
          object Label6: TLabel
            Left = 320
            Top = 24
            Width = 52
            Height = 14
            Caption = #37679#35492#20301#32622':'
          end
          object Label7: TLabel
            Left = 320
            Top = 48
            Width = 52
            Height = 14
            Caption = #20633#20221#20301#32622':'
          end
          object edtWrapperSrc: TcxButtonEdit
            Left = 79
            Top = 20
            ParentFont = False
            Properties.Buttons = <
              item
                Default = True
                Kind = bkEllipsis
              end>
            Properties.ReadOnly = True
            Properties.OnButtonClick = FilePathButtonClick
            Style.StyleController = StyleModule.cxEditStyle
            TabOrder = 0
            Width = 230
          end
          object edtWrapperDest: TcxButtonEdit
            Left = 79
            Top = 45
            ParentFont = False
            Properties.Buttons = <
              item
                Default = True
                Kind = bkEllipsis
              end>
            Properties.ReadOnly = True
            Properties.OnButtonClick = FilePathButtonClick
            Style.StyleController = StyleModule.cxEditStyle
            TabOrder = 1
            Width = 230
          end
          object edtWrapperErr: TcxButtonEdit
            Left = 379
            Top = 20
            ParentFont = False
            Properties.Buttons = <
              item
                Default = True
                Kind = bkEllipsis
              end>
            Properties.ReadOnly = True
            Properties.OnButtonClick = FilePathButtonClick
            Style.StyleController = StyleModule.cxEditStyle
            TabOrder = 2
            Width = 230
          end
          object edtWrapperBackup: TcxButtonEdit
            Left = 379
            Top = 45
            ParentFont = False
            Properties.Buttons = <
              item
                Default = True
                Kind = bkEllipsis
              end>
            Properties.ReadOnly = True
            Properties.OnButtonClick = FilePathButtonClick
            Style.StyleController = StyleModule.cxEditStyle
            TabOrder = 3
            Width = 230
          end
        end
        object AsRunGroup: TcxGroupBox
          Left = 16
          Top = 227
          Caption = ' AsRun '#27284#26696#36335#24465'  '
          Style.StyleController = StyleModule.cxGroupBoxStyle
          TabOrder = 2
          Height = 81
          Width = 622
          object Label8: TLabel
            Left = 20
            Top = 24
            Width = 52
            Height = 14
            Caption = #20358#28304#20301#32622':'
          end
          object Label9: TLabel
            Left = 20
            Top = 48
            Width = 52
            Height = 14
            Caption = #30446#30340#20301#32622':'
          end
          object Label10: TLabel
            Left = 320
            Top = 24
            Width = 52
            Height = 14
            Caption = #37679#35492#20301#32622':'
          end
          object Label11: TLabel
            Left = 320
            Top = 48
            Width = 52
            Height = 14
            Caption = #20633#20221#20301#32622':'
          end
          object edtAsRunSrc: TcxButtonEdit
            Left = 79
            Top = 20
            ParentFont = False
            Properties.Buttons = <
              item
                Default = True
                Kind = bkEllipsis
              end>
            Properties.ReadOnly = True
            Properties.OnButtonClick = FilePathButtonClick
            Style.StyleController = StyleModule.cxEditStyle
            TabOrder = 0
            Width = 230
          end
          object edtAsRunDest: TcxButtonEdit
            Left = 79
            Top = 45
            ParentFont = False
            Properties.Buttons = <
              item
                Default = True
                Kind = bkEllipsis
              end>
            Properties.ReadOnly = True
            Properties.OnButtonClick = FilePathButtonClick
            Style.StyleController = StyleModule.cxEditStyle
            TabOrder = 1
            Width = 230
          end
          object edtAsRunErr: TcxButtonEdit
            Left = 379
            Top = 20
            ParentFont = False
            Properties.Buttons = <
              item
                Default = True
                Kind = bkEllipsis
              end>
            Properties.ReadOnly = True
            Properties.OnButtonClick = FilePathButtonClick
            Style.StyleController = StyleModule.cxEditStyle
            TabOrder = 2
            Width = 230
          end
          object edtAsRunBackup: TcxButtonEdit
            Left = 379
            Top = 45
            ParentFont = False
            Properties.Buttons = <
              item
                Default = True
                Kind = bkEllipsis
              end>
            Properties.ReadOnly = True
            Properties.OnButtonClick = FilePathButtonClick
            Style.StyleController = StyleModule.cxEditStyle
            TabOrder = 3
            Width = 230
          end
        end
      end
      object cxTabSheet2: TcxTabSheet
        Caption = '  '#21295#25972#36913#26399#35373#23450'  '
        ImageIndex = 1
        object PeriodGroup: TcxGroupBox
          Left = 20
          Top = 25
          Caption = ' '#27284#26696#20597#28204#36913#26399' '
          Style.StyleController = StyleModule.cxGroupBoxStyle
          TabOrder = 0
          Height = 230
          Width = 617
          object rdoCheckFix1: TcxRadioButton
            Left = 25
            Top = 32
            Width = 269
            Height = 17
            Caption = #27599'                  '#20998#37912#20597#28204#26159#21542#26377#27284#26696#38920#35201#19979#36617
            TabOrder = 0
            OnClick = PeriodTypeClick
            LookAndFeel.Kind = lfFlat
          end
          object rdoCheckFix2: TcxRadioButton
            Left = 25
            Top = 68
            Width = 290
            Height = 17
            Caption = #22266#23450#19979#21015#26178#38291' ( 24'#23567#26178#21046' )'
            TabOrder = 1
            OnClick = PeriodTypeClick
            LookAndFeel.Kind = lfFlat
          end
          object edtMinute: TcxMaskEdit
            Left = 63
            Top = 30
            ParentFont = False
            Properties.MaskKind = emkRegExpr
            Properties.EditMask = '\d+'
            Properties.MaxLength = 0
            Style.StyleController = StyleModule.cxEditStyle
            TabOrder = 2
            Width = 57
          end
          object chkClock1: TcxCheckBox
            Left = 41
            Top = 97
            Caption = '1:00'
            ParentFont = False
            Style.Font.Charset = DEFAULT_CHARSET
            Style.Font.Color = clWindowText
            Style.Font.Height = -11
            Style.Font.Name = 'MS Sans Serif'
            Style.Font.Style = []
            Style.StyleController = StyleModule.cxEditStyle
            Style.IsFontAssigned = True
            TabOrder = 3
            Width = 54
          end
          object chkClock2: TcxCheckBox
            Left = 41
            Top = 120
            Caption = '2:00'
            ParentFont = False
            Style.Font.Charset = DEFAULT_CHARSET
            Style.Font.Color = clWindowText
            Style.Font.Height = -11
            Style.Font.Name = 'MS Sans Serif'
            Style.Font.Style = []
            Style.StyleController = StyleModule.cxEditStyle
            Style.IsFontAssigned = True
            TabOrder = 4
            Width = 54
          end
          object chkClock3: TcxCheckBox
            Left = 41
            Top = 143
            Caption = '3:00'
            ParentFont = False
            Style.Font.Charset = DEFAULT_CHARSET
            Style.Font.Color = clWindowText
            Style.Font.Height = -11
            Style.Font.Name = 'MS Sans Serif'
            Style.Font.Style = []
            Style.StyleController = StyleModule.cxEditStyle
            Style.IsFontAssigned = True
            TabOrder = 5
            Width = 54
          end
          object chkClock4: TcxCheckBox
            Left = 41
            Top = 166
            Caption = '4:00'
            ParentFont = False
            Style.Font.Charset = DEFAULT_CHARSET
            Style.Font.Color = clWindowText
            Style.Font.Height = -11
            Style.Font.Name = 'MS Sans Serif'
            Style.Font.Style = []
            Style.StyleController = StyleModule.cxEditStyle
            Style.IsFontAssigned = True
            TabOrder = 6
            Width = 54
          end
          object chkClock5: TcxCheckBox
            Left = 41
            Top = 190
            Caption = '5:00'
            ParentFont = False
            Style.Font.Charset = DEFAULT_CHARSET
            Style.Font.Color = clWindowText
            Style.Font.Height = -11
            Style.Font.Name = 'MS Sans Serif'
            Style.Font.Style = []
            Style.StyleController = StyleModule.cxEditStyle
            Style.IsFontAssigned = True
            TabOrder = 7
            Width = 54
          end
          object chkClock6: TcxCheckBox
            Left = 101
            Top = 97
            Caption = '6:00'
            ParentFont = False
            Style.Font.Charset = DEFAULT_CHARSET
            Style.Font.Color = clWindowText
            Style.Font.Height = -11
            Style.Font.Name = 'MS Sans Serif'
            Style.Font.Style = []
            Style.StyleController = StyleModule.cxEditStyle
            Style.IsFontAssigned = True
            TabOrder = 8
            Width = 61
          end
          object chkClock7: TcxCheckBox
            Left = 101
            Top = 120
            Caption = '7:00'
            ParentFont = False
            Style.Font.Charset = DEFAULT_CHARSET
            Style.Font.Color = clWindowText
            Style.Font.Height = -11
            Style.Font.Name = 'MS Sans Serif'
            Style.Font.Style = []
            Style.StyleController = StyleModule.cxEditStyle
            Style.IsFontAssigned = True
            TabOrder = 9
            Width = 61
          end
          object chkClock8: TcxCheckBox
            Left = 101
            Top = 143
            Caption = '8:00'
            ParentFont = False
            Style.Font.Charset = DEFAULT_CHARSET
            Style.Font.Color = clWindowText
            Style.Font.Height = -11
            Style.Font.Name = 'MS Sans Serif'
            Style.Font.Style = []
            Style.StyleController = StyleModule.cxEditStyle
            Style.IsFontAssigned = True
            TabOrder = 10
            Width = 61
          end
          object chkClock9: TcxCheckBox
            Left = 101
            Top = 166
            Caption = '9:00'
            ParentFont = False
            Style.Font.Charset = DEFAULT_CHARSET
            Style.Font.Color = clWindowText
            Style.Font.Height = -11
            Style.Font.Name = 'MS Sans Serif'
            Style.Font.Style = []
            Style.StyleController = StyleModule.cxEditStyle
            Style.IsFontAssigned = True
            TabOrder = 11
            Width = 61
          end
          object chkClock10: TcxCheckBox
            Left = 101
            Top = 190
            Caption = '10:00'
            ParentFont = False
            Style.Font.Charset = DEFAULT_CHARSET
            Style.Font.Color = clWindowText
            Style.Font.Height = -11
            Style.Font.Name = 'MS Sans Serif'
            Style.Font.Style = []
            Style.StyleController = StyleModule.cxEditStyle
            Style.IsFontAssigned = True
            TabOrder = 12
            Width = 61
          end
          object chkClock11: TcxCheckBox
            Left = 164
            Top = 97
            Caption = '11:00'
            ParentFont = False
            Style.Font.Charset = DEFAULT_CHARSET
            Style.Font.Color = clWindowText
            Style.Font.Height = -11
            Style.Font.Name = 'MS Sans Serif'
            Style.Font.Style = []
            Style.StyleController = StyleModule.cxEditStyle
            Style.IsFontAssigned = True
            TabOrder = 13
            Width = 62
          end
          object chkClock12: TcxCheckBox
            Left = 164
            Top = 120
            Caption = '12:00'
            ParentFont = False
            Style.Font.Charset = DEFAULT_CHARSET
            Style.Font.Color = clWindowText
            Style.Font.Height = -11
            Style.Font.Name = 'MS Sans Serif'
            Style.Font.Style = []
            Style.StyleController = StyleModule.cxEditStyle
            Style.IsFontAssigned = True
            TabOrder = 14
            Width = 62
          end
          object chkClock13: TcxCheckBox
            Left = 164
            Top = 143
            Caption = '13:00'
            ParentFont = False
            Style.Font.Charset = DEFAULT_CHARSET
            Style.Font.Color = clWindowText
            Style.Font.Height = -11
            Style.Font.Name = 'MS Sans Serif'
            Style.Font.Style = []
            Style.StyleController = StyleModule.cxEditStyle
            Style.IsFontAssigned = True
            TabOrder = 15
            Width = 62
          end
          object chkClock14: TcxCheckBox
            Left = 164
            Top = 166
            Caption = '14:00'
            ParentFont = False
            Style.Font.Charset = DEFAULT_CHARSET
            Style.Font.Color = clWindowText
            Style.Font.Height = -11
            Style.Font.Name = 'MS Sans Serif'
            Style.Font.Style = []
            Style.StyleController = StyleModule.cxEditStyle
            Style.IsFontAssigned = True
            TabOrder = 16
            Width = 62
          end
          object chkClock15: TcxCheckBox
            Left = 164
            Top = 190
            Caption = '15:00'
            ParentFont = False
            Style.Font.Charset = DEFAULT_CHARSET
            Style.Font.Color = clWindowText
            Style.Font.Height = -11
            Style.Font.Name = 'MS Sans Serif'
            Style.Font.Style = []
            Style.StyleController = StyleModule.cxEditStyle
            Style.IsFontAssigned = True
            TabOrder = 17
            Width = 62
          end
          object chkClock16: TcxCheckBox
            Left = 230
            Top = 97
            Caption = '16:00'
            ParentFont = False
            Style.Font.Charset = DEFAULT_CHARSET
            Style.Font.Color = clWindowText
            Style.Font.Height = -11
            Style.Font.Name = 'MS Sans Serif'
            Style.Font.Style = []
            Style.StyleController = StyleModule.cxEditStyle
            Style.IsFontAssigned = True
            TabOrder = 18
            Width = 63
          end
          object chkClock17: TcxCheckBox
            Left = 230
            Top = 120
            Caption = '17:00'
            ParentFont = False
            Style.Font.Charset = DEFAULT_CHARSET
            Style.Font.Color = clWindowText
            Style.Font.Height = -11
            Style.Font.Name = 'MS Sans Serif'
            Style.Font.Style = []
            Style.StyleController = StyleModule.cxEditStyle
            Style.IsFontAssigned = True
            TabOrder = 19
            Width = 63
          end
          object chkClock18: TcxCheckBox
            Left = 230
            Top = 143
            Caption = '18:00'
            ParentFont = False
            Style.Font.Charset = DEFAULT_CHARSET
            Style.Font.Color = clWindowText
            Style.Font.Height = -11
            Style.Font.Name = 'MS Sans Serif'
            Style.Font.Style = []
            Style.StyleController = StyleModule.cxEditStyle
            Style.IsFontAssigned = True
            TabOrder = 20
            Width = 63
          end
          object chkClock19: TcxCheckBox
            Left = 230
            Top = 166
            Caption = '19:00'
            ParentFont = False
            Style.Font.Charset = DEFAULT_CHARSET
            Style.Font.Color = clWindowText
            Style.Font.Height = -11
            Style.Font.Name = 'MS Sans Serif'
            Style.Font.Style = []
            Style.StyleController = StyleModule.cxEditStyle
            Style.IsFontAssigned = True
            TabOrder = 21
            Width = 63
          end
          object chkClock20: TcxCheckBox
            Left = 230
            Top = 190
            Caption = '20:00'
            ParentFont = False
            Style.Font.Charset = DEFAULT_CHARSET
            Style.Font.Color = clWindowText
            Style.Font.Height = -11
            Style.Font.Name = 'MS Sans Serif'
            Style.Font.Style = []
            Style.StyleController = StyleModule.cxEditStyle
            Style.IsFontAssigned = True
            TabOrder = 22
            Width = 63
          end
          object chkClock21: TcxCheckBox
            Left = 300
            Top = 100
            Caption = '21:00'
            ParentFont = False
            Style.Font.Charset = DEFAULT_CHARSET
            Style.Font.Color = clWindowText
            Style.Font.Height = -11
            Style.Font.Name = 'MS Sans Serif'
            Style.Font.Style = []
            Style.StyleController = StyleModule.cxEditStyle
            Style.IsFontAssigned = True
            TabOrder = 23
            Width = 63
          end
          object chkClock22: TcxCheckBox
            Left = 300
            Top = 123
            Caption = '22:00'
            ParentFont = False
            Style.Font.Charset = DEFAULT_CHARSET
            Style.Font.Color = clWindowText
            Style.Font.Height = -11
            Style.Font.Name = 'MS Sans Serif'
            Style.Font.Style = []
            Style.StyleController = StyleModule.cxEditStyle
            Style.IsFontAssigned = True
            TabOrder = 24
            Width = 63
          end
          object chkClock23: TcxCheckBox
            Left = 300
            Top = 146
            Caption = '23:00'
            ParentFont = False
            Style.Font.Charset = DEFAULT_CHARSET
            Style.Font.Color = clWindowText
            Style.Font.Height = -11
            Style.Font.Name = 'MS Sans Serif'
            Style.Font.Style = []
            Style.StyleController = StyleModule.cxEditStyle
            Style.IsFontAssigned = True
            TabOrder = 25
            Width = 63
          end
          object chkClock24: TcxCheckBox
            Left = 300
            Top = 169
            Caption = '00:00'
            ParentFont = False
            Style.Font.Charset = DEFAULT_CHARSET
            Style.Font.Color = clWindowText
            Style.Font.Height = -11
            Style.Font.Name = 'MS Sans Serif'
            Style.Font.Style = []
            Style.StyleController = StyleModule.cxEditStyle
            Style.IsFontAssigned = True
            TabOrder = 26
            Width = 63
          end
        end
      end
      object cxTabSheet3: TcxTabSheet
        Caption = '  '#36039#26009#24235#35373#23450'  '
        ImageIndex = 2
        object cxGroupBox2: TcxGroupBox
          Left = 20
          Top = 25
          Caption = ' '#36039#26009#24235#30331#20837#35373#23450' '
          Style.StyleController = StyleModule.cxGroupBoxStyle
          TabOrder = 0
          Height = 68
          Width = 621
          object Label15: TLabel
            Left = 25
            Top = 28
            Width = 64
            Height = 14
            Caption = #36039#26009#24235#21029#21517':'
          end
          object Label16: TLabel
            Left = 303
            Top = 28
            Width = 40
            Height = 14
            Caption = #20351#29992#32773':'
          end
          object Label17: TLabel
            Left = 462
            Top = 28
            Width = 28
            Height = 14
            Caption = #23494#30908':'
          end
          object edtDbAccount: TcxTextEdit
            Left = 347
            Top = 25
            ParentFont = False
            Properties.EchoMode = eemPassword
            Properties.ReadOnly = False
            Style.StyleController = StyleModule.cxEditStyle
            TabOrder = 1
            Width = 100
          end
          object edtDbPassword: TcxTextEdit
            Left = 496
            Top = 25
            ParentFont = False
            Properties.EchoMode = eemPassword
            Properties.ReadOnly = False
            Style.StyleController = StyleModule.cxEditStyle
            TabOrder = 2
            Width = 100
          end
          object cmbDbSid: TcxComboBox
            Left = 97
            Top = 25
            ParentFont = False
            Properties.DropDownListStyle = lsEditFixedList
            Style.StyleController = StyleModule.cxEditStyle
            TabOrder = 0
            Width = 185
          end
        end
        object cxGroupBox5: TcxGroupBox
          Left = 20
          Top = 112
          Caption = ' '#26039#32218#34389#29702
          Style.StyleController = StyleModule.cxGroupBoxStyle
          TabOrder = 1
          Height = 68
          Width = 621
          object lbDbRetryFrequence: TLabel
            Left = 25
            Top = 30
            Width = 236
            Height = 14
            Caption = #36039#26009#24235#26039#32218#26178', '#31561#20505'                  '#31186#37325#35430#36899#32218
            Transparent = True
          end
          object edtDbRetrySec: TcxSpinEdit
            Left = 136
            Top = 26
            ParentFont = False
            Properties.MaxValue = 600.000000000000000000
            Properties.MinValue = 30.000000000000000000
            Style.StyleController = StyleModule.cxEditStyle
            TabOrder = 0
            Value = 60
            Width = 58
          end
        end
        object cxGroupBox6: TcxGroupBox
          Left = 20
          Top = 201
          Caption = ' '#27511#31243#35352#37636
          Style.StyleController = StyleModule.cxGroupBoxStyle
          TabOrder = 2
          Height = 68
          Width = 621
          object Label19: TLabel
            Left = 25
            Top = 30
            Width = 476
            Height = 14
            Caption = 
              #31243#24335#21855#21205#26178', '#36617#20837'                   '#22825#21069#25286#35299#27511#21490#35352#37636',  '#20197#21450#36617#20837'                    '#22825 +
              #21069#35338#24687#35352#37636
            Transparent = True
          end
          object edtActLoadDays: TcxSpinEdit
            Left = 125
            Top = 26
            ParentFont = False
            Properties.AssignedValues.MinValue = True
            Properties.MaxValue = 30.000000000000000000
            Style.StyleController = StyleModule.cxEditStyle
            TabOrder = 0
            Value = 3
            Width = 58
          end
          object edtMsgLoadDays: TcxSpinEdit
            Left = 361
            Top = 26
            ParentFont = False
            Properties.AssignedValues.MinValue = True
            Properties.MaxValue = 30.000000000000000000
            Style.StyleController = StyleModule.cxEditStyle
            TabOrder = 1
            Value = 3
            Width = 58
          end
        end
      end
      object cxTabSheet4: TcxTabSheet
        Caption = '  '#30435#25511#36890#30693#27231#21046'  '
        ImageIndex = 3
        object cxGroupBox1: TcxGroupBox
          Left = 20
          Top = 30
          Caption = ' '#36939#20316#37679#35492#36890#30693' '
          Style.StyleController = StyleModule.cxGroupBoxStyle
          TabOrder = 0
          Height = 72
          Width = 621
          object chkErrorNotify: TcxCheckBox
            Left = 24
            Top = 28
            Caption = #21855#29992' ( '#36890#30693#26041#24335#28858' E-Mail )'
            ParentFont = False
            Properties.OnChange = NotifyPropertiesChange
            Style.StyleController = StyleModule.cxEditStyle
            TabOrder = 0
            Width = 170
          end
          object cmbErrList: TcxComboBox
            Left = 196
            Top = 28
            ParentFont = False
            Properties.DropDownListStyle = lsFixedList
            Properties.ReadOnly = True
            Style.StyleController = StyleModule.cxEditStyle
            TabOrder = 1
            Width = 345
          end
          object btnAddErr: TcxButton
            Left = 543
            Top = 28
            Width = 23
            Height = 20
            Caption = '...'
            TabOrder = 2
            OnClick = EmailModifyClick
          end
        end
        object cxGroupBox3: TcxGroupBox
          Left = 20
          Top = 122
          Caption = ' '#31680#30446#34920#30064#21205#36890#30693' '
          Style.StyleController = StyleModule.cxGroupBoxStyle
          TabOrder = 1
          Height = 72
          Width = 621
          object chkChangeNotify: TcxCheckBox
            Left = 24
            Top = 28
            Caption = #21855#29992' ( '#36890#30693#26041#24335#28858' E-Mail )'
            ParentFont = False
            Properties.OnChange = NotifyPropertiesChange
            Style.StyleController = StyleModule.cxEditStyle
            TabOrder = 0
            Width = 170
          end
          object cmbChangeList: TcxComboBox
            Left = 196
            Top = 28
            ParentFont = False
            Properties.DropDownListStyle = lsFixedList
            Properties.ReadOnly = True
            Style.StyleController = StyleModule.cxEditStyle
            TabOrder = 1
            Width = 345
          end
          object btnAddChange: TcxButton
            Left = 543
            Top = 28
            Width = 23
            Height = 20
            Caption = '...'
            TabOrder = 2
            OnClick = EmailModifyClick
          end
        end
        object cxGroupBox4: TcxGroupBox
          Left = 20
          Top = 215
          Caption = ' SMTP'#20282#26381#22120#36039#35338' '
          Style.StyleController = StyleModule.cxGroupBoxStyle
          TabOrder = 2
          Height = 111
          Width = 621
          object Label12: TLabel
            Left = 25
            Top = 29
            Width = 64
            Height = 14
            Caption = #37109#20214#20282#26381#22120':'
          end
          object Label13: TLabel
            Left = 283
            Top = 29
            Width = 88
            Height = 14
            Caption = #36890#30693#32773#38651#23376#37109#20214':'
          end
          object Label14: TLabel
            Left = 25
            Top = 61
            Width = 64
            Height = 14
            Caption = #30331#20837#20351#29992#32773':'
          end
          object Label18: TLabel
            Left = 240
            Top = 61
            Width = 28
            Height = 14
            Caption = #23494#30908':'
          end
          object edtSMTPServer: TcxTextEdit
            Left = 97
            Top = 26
            ParentFont = False
            Style.StyleController = StyleModule.cxEditStyle
            TabOrder = 0
            Width = 170
          end
          object edtSMTPEmail: TcxHyperLinkEdit
            Left = 377
            Top = 26
            ParentFont = False
            Style.StyleController = StyleModule.cxEditStyle
            TabOrder = 1
            Width = 182
          end
          object edtSMTPAccount: TcxTextEdit
            Left = 97
            Top = 58
            ParentFont = False
            Properties.EchoMode = eemPassword
            Style.StyleController = StyleModule.cxEditStyle
            TabOrder = 2
            Width = 112
          end
          object edtSMTPPassword: TcxTextEdit
            Left = 276
            Top = 58
            ParentFont = False
            Properties.EchoMode = eemPassword
            Style.StyleController = StyleModule.cxEditStyle
            TabOrder = 3
            Width = 112
          end
        end
      end
    end
  end
end
