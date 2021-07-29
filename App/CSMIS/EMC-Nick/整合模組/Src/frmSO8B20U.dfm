object frmSO8B20: TfrmSO8B20
  Left = 335
  Top = 141
  Width = 401
  Height = 381
  Caption = 'frmSO8B20'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Label2: TLabel
    Left = 27
    Top = 12
    Width = 67
    Height = 16
    AutoSize = False
    Caption = #20844#21496#21029#65306
    Font.Charset = ANSI_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = #32048#26126#39636
    Font.Style = []
    ParentFont = False
  end
  object Label4: TLabel
    Left = 22
    Top = 228
    Width = 133
    Height = 13
    Caption = 'BOX'#38656#25490#38500#30340#26989#21209#27963#21205
    Font.Charset = ANSI_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = #32048#26126#39636
    Font.Style = []
    ParentFont = False
  end
  object btnStartCalc: TButton
    Left = 303
    Top = 224
    Width = 75
    Height = 31
    Caption = #38283#22987#35336#31639
    TabOrder = 6
    OnClick = btnStartCalcClick
  end
  object btnClose: TButton
    Left = 303
    Top = 290
    Width = 75
    Height = 28
    Caption = #38626#38283
    TabOrder = 8
    OnClick = btnCloseClick
  end
  object stxCom: TStaticText
    Left = 93
    Top = 9
    Width = 46
    Height = 17
    Caption = 'stxCom'
    Font.Charset = ANSI_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = #32048#26126#39636
    Font.Style = []
    ParentFont = False
    TabOrder = 0
  end
  object rgpCompute: TRadioGroup
    Left = 8
    Top = 97
    Width = 57
    Height = 21
    Caption = ' '#35336#31639#20381#25818'  '
    Columns = 2
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ItemIndex = 0
    Items.Strings = (
      #20381#20844#21496#35336#31639
      #20381#37096#38272#35336#31639)
    ParentFont = False
    TabOrder = 1
    Visible = False
  end
  object chbPage: TCheckBox
    Left = 6
    Top = 128
    Width = 43
    Height = 17
    Caption = #20381#35336#31639#20381#25818#20998#38913
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    TabOrder = 2
    Visible = False
  end
  object btnQueryComm: TButton
    Left = 303
    Top = 258
    Width = 75
    Height = 28
    Caption = #20323#29518#37329#26597#35426
    TabOrder = 7
    OnClick = btnQueryCommClick
  end
  object lbxBox: TListBox
    Left = 19
    Top = 244
    Width = 238
    Height = 73
    Columns = 2
    ItemHeight = 13
    TabOrder = 5
  end
  object StatusBar1: TStatusBar
    Left = 0
    Top = 334
    Width = 393
    Height = 19
    Panels = <>
    SimplePanel = True
  end
  object GroupBox1: TGroupBox
    Left = 21
    Top = 36
    Width = 177
    Height = 181
    Caption = ' '#38971#36947#20323#37329#20998#26399'/STB'#29518#37329'  '
    Font.Charset = ANSI_CHARSET
    Font.Color = clBlack
    Font.Height = -13
    Font.Name = #32048#26126#39636
    Font.Style = []
    ParentFont = False
    TabOrder = 3
    object Label1: TLabel
      Left = 14
      Top = 25
      Width = 79
      Height = 20
      AutoSize = False
      Caption = #32113#35336#24180#26376#65306
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = #32048#26126#39636
      Font.Style = []
      ParentFont = False
    end
    object Label3: TLabel
      Left = 14
      Top = 55
      Width = 79
      Height = 20
      AutoSize = False
      Caption = #27512#23660#24180#26376#65306
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = #32048#26126#39636
      Font.Style = []
      ParentFont = False
    end
    inline fraChineseYM1: TfraChineseYM
      Left = 83
      Top = 13
      Width = 76
      Height = 35
      TabOrder = 0
    end
    inline fraChineseYM2: TfraChineseYM
      Left = 83
      Top = 42
      Width = 76
      Height = 35
      TabOrder = 1
      OnExit = fraChineseYM2Exit
    end
    object GroupBox3: TGroupBox
      Left = 8
      Top = 80
      Width = 161
      Height = 89
      Caption = ' '#19978#27425#35336#31639#36039#35338' '
      TabOrder = 2
      object lblLastData1: TLabel
        Left = 8
        Top = 20
        Width = 72
        Height = 12
        Caption = 'lblLastData1'
        Font.Charset = ANSI_CHARSET
        Font.Color = clNavy
        Font.Height = -12
        Font.Name = #32048#26126#39636
        Font.Style = []
        ParentFont = False
      end
      object lblLastData2: TLabel
        Left = 9
        Top = 42
        Width = 72
        Height = 12
        Caption = 'lblLastData2'
        Font.Charset = ANSI_CHARSET
        Font.Color = clNavy
        Font.Height = -12
        Font.Name = #32048#26126#39636
        Font.Style = []
        ParentFont = False
      end
      object lblLastData3: TLabel
        Left = 9
        Top = 66
        Width = 72
        Height = 12
        Caption = 'lblLastData3'
        Font.Charset = ANSI_CHARSET
        Font.Color = clNavy
        Font.Height = -12
        Font.Name = #32048#26126#39636
        Font.Style = []
        ParentFont = False
      end
    end
  end
  object GroupBox2: TGroupBox
    Left = 205
    Top = 36
    Width = 177
    Height = 180
    Caption = ' '#38971#36947#20323#37329#19968#27425#32102#20184'  '
    Font.Charset = ANSI_CHARSET
    Font.Color = clBlack
    Font.Height = -13
    Font.Name = #32048#26126#39636
    Font.Style = []
    ParentFont = False
    TabOrder = 4
    object Label5: TLabel
      Left = 14
      Top = 24
      Width = 79
      Height = 20
      AutoSize = False
      Caption = #32113#35336#24180#26376#65306
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = #32048#26126#39636
      Font.Style = []
      ParentFont = False
    end
    object Label6: TLabel
      Left = 14
      Top = 53
      Width = 79
      Height = 20
      AutoSize = False
      Caption = #27512#23660#24180#26376#65306
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = #32048#26126#39636
      Font.Style = []
      ParentFont = False
    end
    inline fraChineseYM3: TfraChineseYM
      Left = 83
      Top = 20
      Width = 76
      Height = 35
      Enabled = False
      TabOrder = 0
      inherited mseYM: TMaskEdit
        Top = 0
      end
    end
    inline fraChineseYM4: TfraChineseYM
      Left = 83
      Top = 40
      Width = 76
      Height = 35
      TabOrder = 1
      OnExit = fraChineseYM4Exit
    end
    object GroupBox4: TGroupBox
      Left = 8
      Top = 80
      Width = 161
      Height = 89
      Caption = ' '#19978#27425#35336#31639#36039#35338' '
      TabOrder = 2
      object lblLastData4: TLabel
        Left = 8
        Top = 20
        Width = 72
        Height = 12
        Caption = 'lblLastData4'
        Font.Charset = ANSI_CHARSET
        Font.Color = clNavy
        Font.Height = -12
        Font.Name = #32048#26126#39636
        Font.Style = []
        ParentFont = False
      end
      object lblLastData5: TLabel
        Left = 9
        Top = 42
        Width = 72
        Height = 12
        Caption = 'lblLastData5'
        Font.Charset = ANSI_CHARSET
        Font.Color = clNavy
        Font.Height = -12
        Font.Name = #32048#26126#39636
        Font.Style = []
        ParentFont = False
      end
      object lblLastData6: TLabel
        Left = 9
        Top = 66
        Width = 72
        Height = 12
        Caption = 'lblLastData6'
        Font.Charset = ANSI_CHARSET
        Font.Color = clNavy
        Font.Height = -12
        Font.Name = #32048#26126#39636
        Font.Style = []
        ParentFont = False
      end
    end
  end
  object Timer1: TTimer
    Enabled = False
    OnTimer = Timer1Timer
    Left = 205
    Top = 12
  end
end
