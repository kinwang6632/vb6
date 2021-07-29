object frmOption: TfrmOption
  Left = 322
  Top = 158
  Width = 368
  Height = 290
  BorderIcons = [biSystemMenu, biMinimize]
  Caption = 'Environment Options'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnClose = FormClose
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Bevel1: TBevel
    Left = 0
    Top = 0
    Width = 360
    Height = 225
    Align = alClient
  end
  object Panel1: TPanel
    Left = 0
    Top = 225
    Width = 360
    Height = 38
    Align = alBottom
    BevelOuter = bvNone
    TabOrder = 0
    object BitBtn1: TBitBtn
      Tag = 1
      Left = 196
      Top = 4
      Width = 80
      Height = 33
      Caption = 'OK'
      TabOrder = 0
      OnClick = BitBtn1Click
    end
    object BitBtn3: TBitBtn
      Left = 278
      Top = 4
      Width = 80
      Height = 33
      Caption = 'Cancel'
      TabOrder = 1
      OnClick = BitBtn3Click
    end
  end
  object PageControl1: TPageControl
    Left = 8
    Top = 8
    Width = 345
    Height = 209
    ActivePage = TabSheet1
    TabOrder = 1
    object TabSheet1: TTabSheet
      Caption = 'Path'
      object Label4: TLabel
        Left = 103
        Top = 136
        Width = 136
        Height = 13
        Caption = 'C:\APP\TransLang\src\Max'
      end
      object Label3: TLabel
        Left = 16
        Top = 136
        Width = 81
        Height = 13
        Alignment = taRightJustify
        AutoSize = False
        Caption = 'Full Path:'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object Label2: TLabel
        Left = 8
        Top = 40
        Width = 81
        Height = 13
        Alignment = taRightJustify
        AutoSize = False
        Caption = 'Path:'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object Label1: TLabel
        Left = 8
        Top = 16
        Width = 81
        Height = 13
        Alignment = taRightJustify
        AutoSize = False
        Caption = 'Disk:'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object Label5: TLabel
        Left = 16
        Top = 160
        Width = 81
        Height = 13
        Alignment = taRightJustify
        AutoSize = False
        Caption = 'Default Path:'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clGray
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object Label6: TLabel
        Left = 103
        Top = 160
        Width = 15
        Height = 13
        Caption = 'C:\'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clGray
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object DirectoryListBox1: TDirectoryListBox
        Left = 95
        Top = 34
        Width = 226
        Height = 95
        DirLabel = Label4
        ImeName = '中文 (繁體) - 新注音'
        ItemHeight = 16
        TabOrder = 0
      end
      object DriveComboBox1: TDriveComboBox
        Left = 95
        Top = 14
        Width = 226
        Height = 19
        DirList = DirectoryListBox1
        ImeName = '中文 (繁體) - 新注音'
        TabOrder = 1
        TextCase = tcUpperCase
      end
    end
    object TabSheet2: TTabSheet
      Caption = 'Loader'
      ImageIndex = 1
      object RadioGroup1: TRadioGroup
        Left = 8
        Top = 8
        Width = 321
        Height = 73
        Caption = ' Loader Option '
        Columns = 2
        ItemIndex = 0
        Items.Strings = (
          'Clear All  (Default)'
          'Append')
        TabOrder = 0
      end
    end
    object TabSheet3: TTabSheet
      Caption = 'Exporter'
      ImageIndex = 2
      object RadioGroup2: TRadioGroup
        Left = 8
        Top = 8
        Width = 321
        Height = 73
        Caption = ' Exporter Option '
        Columns = 2
        ItemIndex = 1
        Items.Strings = (
          'Clear All'
          'Append (Default)')
        TabOrder = 0
      end
    end
    object TabSheet4: TTabSheet
      Caption = 'Default'
      ImageIndex = 3
      object GroupBox1: TGroupBox
        Left = 8
        Top = 8
        Width = 321
        Height = 73
        Caption = ' Default Option '
        TabOrder = 0
        object BitBtn2: TBitBtn
          Left = 8
          Top = 22
          Width = 305
          Height = 33
          Caption = 'Click me to load default setting'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlue
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 0
          OnClick = BitBtn2Click
        end
      end
    end
  end
end
