object frmSnapShot: TfrmSnapShot
  Left = 18
  Top = 0
  Width = 748
  Height = 547
  Caption = 'SnapShot  '#22519#34892#32000#37636#26597#35426
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
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 740
    Height = 49
    Align = alTop
    TabOrder = 0
    object Label1: TLabel
      Left = 34
      Top = 16
      Width = 35
      Height = 13
      Caption = 'Alias'
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = #32048#26126#39636
      Font.Style = []
      ParentFont = False
    end
    object Label2: TLabel
      Left = 328
      Top = 8
      Width = 56
      Height = 13
      Caption = 'UserName'
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = #32048#26126#39636
      Font.Style = []
      ParentFont = False
      Visible = False
    end
    object Label3: TLabel
      Left = 408
      Top = 8
      Width = 56
      Height = 13
      Caption = 'Password'
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = #32048#26126#39636
      Font.Style = []
      ParentFont = False
      Visible = False
    end
    object Label4: TLabel
      Left = 176
      Top = 16
      Width = 56
      Height = 13
      Caption = #26597#35426#26085#26399
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = #32048#26126#39636
      Font.Style = []
      ParentFont = False
    end
    object edtAlias: TEdit
      Left = 80
      Top = 11
      Width = 73
      Height = 21
      TabOrder = 0
      Text = 'catvnr.world'
    end
    object edtUserName: TEdit
      Left = 320
      Top = 19
      Width = 73
      Height = 21
      TabOrder = 1
      Visible = False
    end
    object edtPassword: TEdit
      Left = 400
      Top = 19
      Width = 73
      Height = 21
      TabOrder = 2
      Visible = False
    end
    inline fraChineseYMD1: TfraChineseYMD
      Left = 240
      Top = 8
      Width = 100
      Height = 35
      TabOrder = 3
      inherited mseYMD: TMaskEdit
        Width = 66
      end
    end
    object btnOK: TBitBtn
      Left = 474
      Top = 16
      Width = 75
      Height = 25
      Caption = #30906#23450
      TabOrder = 4
      OnClick = btnOKClick
    end
    object btnClose: TButton
      Tag = 1
      Left = 629
      Top = 16
      Width = 75
      Height = 25
      Caption = #38626#38283
      ModalResult = 2
      TabOrder = 6
      OnClick = btnCloseClick
    end
    object btnAllReCompute: TBitBtn
      Left = 552
      Top = 16
      Width = 75
      Height = 25
      Caption = #20840#37096#35336#31639
      TabOrder = 5
      OnClick = btnAllReComputeClick
    end
  end
  object Panel2: TPanel
    Left = 0
    Top = 49
    Width = 740
    Height = 471
    Align = alClient
    Caption = 'Panel2'
    TabOrder = 1
    object Panel3: TPanel
      Left = 1
      Top = 1
      Width = 738
      Height = 168
      Align = alTop
      Caption = 'Panel3'
      TabOrder = 0
      object DBGrid1: TDBGrid
        Left = 1
        Top = 1
        Width = 736
        Height = 166
        Align = alClient
        DataSource = dsrSO030A
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 0
        TitleFont.Charset = DEFAULT_CHARSET
        TitleFont.Color = clWindowText
        TitleFont.Height = -11
        TitleFont.Name = 'MS Sans Serif'
        TitleFont.Style = []
        OnCellClick = DBGrid1CellClick
        OnTitleClick = DBGrid1TitleClick
      end
    end
    object TPanel
      Left = 1
      Top = 169
      Width = 738
      Height = 40
      Align = alTop
      TabOrder = 1
      object lblCompName: TLabel
        Left = 16
        Top = 12
        Width = 30
        Height = 13
        Caption = #23458#25142
        Font.Charset = ANSI_CHARSET
        Font.Color = clNavy
        Font.Height = -13
        Font.Name = #32048#26126#39636
        Font.Style = [fsBold]
        ParentFont = False
      end
      object btnReCompute: TButton
        Left = 624
        Top = 8
        Width = 75
        Height = 25
        Caption = #37325#26032#35336#31639
        TabOrder = 0
        OnClick = btnReComputeClick
      end
    end
    object Panel5: TPanel
      Left = 1
      Top = 209
      Width = 738
      Height = 261
      Align = alClient
      Caption = 'Panel5'
      TabOrder = 2
      object DBGrid2: TDBGrid
        Left = 1
        Top = 1
        Width = 736
        Height = 259
        Align = alClient
        DataSource = dsrSO509B
        TabOrder = 0
        TitleFont.Charset = DEFAULT_CHARSET
        TitleFont.Color = clWindowText
        TitleFont.Height = -11
        TitleFont.Name = 'MS Sans Serif'
        TitleFont.Style = []
      end
    end
  end
  object dsrSO030A: TDataSource
    DataSet = dtmMain.cdsSO030A
    Left = 16
    Top = 104
  end
  object dsrSO509B: TDataSource
    DataSet = dtmMain.cdsTempSO509B
    Left = 16
    Top = 296
  end
end
